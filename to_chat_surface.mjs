#!/usr/bin/env node
// Recover the headless Teams page onto the Chat surface (nav click or deep link).
import fs from 'node:fs';

const PORT = 9333;
const CHAT_ID = process.argv[2] || '19:17d579a4-0e2b-4f1d-a130-3917ca8ab605_1bcb80e7-9ab2-4ea2-a5c4-33ac312ff0bd@unq.gbl.spaces';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const targets = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
const page = targets.find((t) => t.type === 'page' && /teams\.microsoft\.com/.test(t.url));
const ws = new WebSocket(page.webSocketDebuggerUrl);
await new Promise((res) => ws.addEventListener('open', res));
let msgId = 0;
const pending = new Map();
ws.addEventListener('message', (ev) => {
  const m = JSON.parse(ev.data);
  if (m.id && pending.has(m.id)) {
    const p = pending.get(m.id);
    pending.delete(m.id);
    p.resolve(m.result);
  }
});
const send = (method, params = {}) =>
  new Promise((resolve) => {
    const id = ++msgId;
    pending.set(id, { resolve });
    ws.send(JSON.stringify({ id, method, params }));
  });
const evaluate = async (expression) => {
  const r = await send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true });
  if (r?.exceptionDetails) return { PAGE_JS_ERROR: JSON.stringify(r.exceptionDetails).slice(0, 400) };
  return r?.result?.value;
};
const surfaceExpr = `(() => ({ title: document.title, url: location.href.slice(0,120), chatItems: [...document.querySelectorAll('span[id^="title-chat-list-item"]')].length }))()`;
const waitSurface = async (timeoutMs) => {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    const v = await evaluate(surfaceExpr).catch(() => null);
    if (v && v.chatItems > 0) return v;
    await sleep(2000);
  }
  return null;
};

console.log('before:', JSON.stringify(await evaluate(surfaceExpr)));

// Attempt 1: click the nav "Chat" button
const navInfo = await evaluate(`(() => {
  const items = [...document.querySelectorAll('[role="navigation"] button, [role="navigation"] a, nav button, nav a')];
  return items.map((b) => ({ tag: b.tagName, aria: b.getAttribute('aria-label') || '', text: (b.textContent || '').trim().slice(0, 30), cls: String(b.className).slice(0, 40) })).slice(0, 15);
})()`);
console.log('nav items:', JSON.stringify(navInfo));
const chatBtn = await evaluate(`(() => {
  const btn = [...document.querySelectorAll('[role="navigation"] button, [role="navigation"] a, nav button, nav a')]
    .find((b) => /^chat/i.test((b.getAttribute('aria-label') || '').trim()) || /^chat/i.test((b.textContent || '').trim()));
  if (!btn) return null;
  btn.click();
  return true;
})()`);
console.log('nav chat click:', JSON.stringify(chatBtn));
if (chatBtn) {
  const v = await waitSurface(60000);
  if (v) {
    console.log('SUCCESS_VIA_NAV', JSON.stringify(v));
    process.exit(0);
  }
}

// Attempt 2: official deep link with RAW chat id
console.log('trying deep link...');
await send('Page.navigate', { url: `https://teams.microsoft.com/l/chat/${CHAT_ID}` });
const v2 = await waitSurface(90000);
if (v2) {
  console.log('SUCCESS_VIA_DEEPLINK', JSON.stringify(v2));
  process.exit(0);
}
console.log('FAILED:', JSON.stringify(await evaluate(surfaceExpr).catch(() => null)));
process.exit(1);
