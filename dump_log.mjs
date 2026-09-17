#!/usr/bin/env node
// Dump the tail of the visible chat message log for send verification.
const PORT = 9333;

const targets = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
const page = targets.find((t) => t.type === 'page' && /teams\.microsoft\.com/.test(t.url));
const ws = new WebSocket(page.webSocketDebuggerUrl);
await new Promise((res) => ws.addEventListener('open', res));
let msgId = 0;
const pending = new Map();
ws.addEventListener('message', (ev) => {
  const m = JSON.parse(ev.data);
  if (m.id && pending.has(m.id)) {
    pending.get(m.id)(m.result);
    pending.delete(m.id);
  }
});const send = (method, params = {}) =>
  new Promise((resolve) => {
    const id = ++msgId;
    pending.set(id, resolve);
    ws.send(JSON.stringify({ id, method, params }));
  });
const ev = async (expression) => {
  const r = await send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true });
  if (r?.exceptionDetails) return { PAGE_JS_ERROR: JSON.stringify(r.exceptionDetails).slice(0, 300) };
  return r?.result?.value;
};

const out = await ev(`(() => {
  const lists = [...document.querySelectorAll('[aria-label*="Message list" i]')];
  const tails = lists.map((l) => (l.innerText || '').trim().slice(-700)).filter(Boolean);
  return { title: document.title, lists: lists.length, tails, bodyTail: (document.body.innerText || '').slice(-700) };
})()`);
console.log(JSON.stringify(out, null, 1));
process.exit(0);
