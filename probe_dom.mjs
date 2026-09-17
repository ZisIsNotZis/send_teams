#!/usr/bin/env node
// One-off DOM probe against the headless Teams page via CDP.
import fs from 'node:fs';

const PORT = 9333;
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

const out = await evaluate(`(() => {
  const norm = (s) => (s || '').toLowerCase().replace(/[(),]/g, ' ').replace(/\\s+/g, ' ').trim();
  const all = [...document.querySelectorAll('body *')];
  const hits = [];
  for (const e of all) {
    const t = norm(e.textContent || '');
    if (t.includes('wu fei')) {
      const own = norm([...e.childNodes].filter((n) => n.nodeType === 3).map((n) => n.textContent).join(' '));
      if (own.includes('wu fei') || e.children.length === 0) hits.push(e);
      if (hits.length > 30) break;
    }
  }
  const describe = (e) => {
    const chain = [];
    let cur = e;
    for (let i = 0; i < 7 && cur && cur.tagName; i++) {
      chain.push({ tag: cur.tagName, role: cur.getAttribute('role'), id: cur.id || '', cls: String(cur.className).slice(0, 70), testid: cur.getAttribute('data-testid') || '', aria: (cur.getAttribute('aria-label') || '').slice(0, 40) });
      cur = cur.parentElement;
    }
    return chain;
  };
  const roles = {};
  for (const e of all) { const r = e.getAttribute('role'); if (r) roles[r] = (roles[r] || 0) + 1; }
  const headers = [...document.querySelectorAll('h1,h2,h3')].map((h) => norm(h.textContent).slice(0, 60)).filter(Boolean).slice(0, 10);
  return {
    title: document.title,
    roles,
    headers,
    hitCount: hits.length,
    hits: hits.slice(0, 12).map((e) => ({ chain: describe(e), text: norm(e.textContent).slice(0, 60) })),
  };
})()`);
console.log(JSON.stringify(out, null, 1));
ws.close();
process.exit(0);
