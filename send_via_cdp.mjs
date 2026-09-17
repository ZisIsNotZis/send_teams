#!/usr/bin/env node
// Drive Teams web running in headless Edge via Chrome DevTools Protocol.
// Zero dependencies: Node >= 22 global fetch/WebSocket.
//
// Usage:
//   node send_via_cdp.mjs "<person>"            # dry run: open chat, report state, NO send
//   node send_via_cdp.mjs "<person>" "<text>"   # open chat, type text, Enter, verify, screenshot

import fs from 'node:fs';

const PORT = 9333;
const [nameQuery = 'Wu, Fei', message = '', chatId = ''] = process.argv.slice(2);
const CHAT_URL = chatId
  ? `https://teams.microsoft.com/v2/l/chat/${encodeURIComponent(chatId)}`
  : null;
const SHOT = '/home/z/vibe/send_teams/evidence-headless.png';

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function main() {
  const targets = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
  const page = targets.find((t) => t.type === 'page' && /teams\.microsoft\.com/.test(t.url));
  if (!page) {
    console.error('NO_TEAMS_TARGET ' + JSON.stringify(targets.map((t) => ({ type: t.type, url: t.url }))));
    process.exit(2);
  }

  const ws = new WebSocket(page.webSocketDebuggerUrl);
  await new Promise((res, rej) => {
    ws.addEventListener('open', res);
    ws.addEventListener('error', () => rej(new Error('ws error')));
  });
  let msgId = 0;
  const pending = new Map();
  ws.addEventListener('message', (ev) => {
    const m = JSON.parse(ev.data);
    if (m.id && pending.has(m.id)) {
      const p = pending.get(m.id);
      pending.delete(m.id);
      m.error ? p.reject(new Error(m.error.message)) : p.resolve(m.result);
    }
  });
  const send = (method, params = {}) =>
    new Promise((resolve, reject) => {
      const id = ++msgId;
      pending.set(id, { resolve, reject });
      ws.send(JSON.stringify({ id, method, params }));
    });

  const evaluate = async (expression) => {
    const r = await send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true });
    if (r?.exceptionDetails) throw new Error('PAGE_JS_ERROR ' + JSON.stringify(r.exceptionDetails).slice(0, 500));
    return r?.result?.value;
  };
  const waitFor = async (expr, timeoutMs, want) => {
    const t0 = Date.now();
    while (Date.now() - t0 < timeoutMs) {
      const v = await evaluate(expr).catch((e) => ({ err: String(e).slice(0, 200) }));
      if (want(v)) return v;
      await sleep(1500);
    }
    return null;
  };

  const HELP = String.raw`
    const norm = (s) => (s || '').toLowerCase().replace(/[(),]/g, ' ').replace(/\s+/g, ' ').trim();
    const CHAT_SELS = '[data-testid="list-item"], [role="treeitem"]';
    const pickEditor = () => {
      const eds = [...document.querySelectorAll('div[contenteditable="true"]')];
      return eds.find((e) => /message/i.test(e.getAttribute('aria-label') || '')) || eds[eds.length - 1] || null;
    };
  `;
  const headingExpr = `(() => { ${HELP} const q = norm(${JSON.stringify(nameQuery)});
    const hs = [...document.querySelectorAll('[role="heading"]')].map((h) => norm(h.textContent)).filter(Boolean);
    return { headings: hs.slice(0, 8), matches: hs.some((h) => h.includes(q)) };
  })()`;
  const listChatsExpr = `(() => { ${HELP} const query = norm(${JSON.stringify(nameQuery)});
    return [...document.querySelectorAll(CHAT_SELS)]
      .filter((e) => norm(e.textContent).includes(query))
      .slice(0, 6)
      .map((e) => ({ tag: e.tagName, id: (e.id || '').slice(0, 60), text: norm(e.textContent).slice(0, 80) }));
  })()`;
  const clickChatExpr = `(() => { ${HELP} const query = norm(${JSON.stringify(nameQuery)});
    let target = null;
    const spans = [...document.querySelectorAll('span[id^="title-chat-list-item"]')]
      .filter((e) => norm(e.textContent).includes(query));
    if (spans.length) {
      target = spans[0].closest('[role="treeitem"], [data-testid="list-item"]');
    } else {
      const cands = [...document.querySelectorAll(CHAT_SELS)].filter((e) => norm(e.textContent).includes(query));
      target = cands.sort((a, b) => norm(a.textContent).length - norm(b.textContent).length)[0] || null;
    }
    if (!target) return null;
    target.scrollIntoView({ block: 'center' });
    target.click();
    return true;
  })()`;
  const stateExpr = `(() => { ${HELP}
    const ed = pickEditor();
    return {
      url: location.href.slice(0, 100),
      signedOut: /sign in|log in/i.test((document.querySelector('body')?.innerText || '').slice(0, 3000)),
      editorFound: !!ed,
      editorLabel: ed ? ed.getAttribute('aria-label') || '' : '',
      editorText: ed ? (ed.textContent || '') : '',
    };
  })()`;
  const focusEditorExpr = `(() => { ${HELP}
    const ed = pickEditor();
    if (!ed) return false;
    ed.focus();
    return document.activeElement === ed;
  })()`;

  // 0. make sure we're on the Chat surface (not Calendar/Teams/etc.)
  const surfaceExpr = `(() => ({ title: document.title, chatItems: [...document.querySelectorAll('span[id^="title-chat-list-item"]')].length }))()`;
  let surf = await evaluate(surfaceExpr);
  if (!surf || !surf.chatItems) {
    console.log('navigating to Chat surface (was: ' + (surf && surf.title) + ')...');
    await send('Page.navigate', { url: 'https://teams.microsoft.com/v2/' });
    surf = await waitFor(surfaceExpr, 120000, (v) => v && !v.err && v.chatItems > 0);
    if (!surf) {
      console.error('CHAT_SURFACE_NOT_REACHED ' + JSON.stringify(await evaluate(surfaceExpr).catch(() => null)));
      process.exit(8);
    }
  }
  console.log('surface:', JSON.stringify(surf));

  // 1. wait for Teams SPA to be interactive
  console.log('waiting for Teams UI...');
  let st = await waitFor(stateExpr, 90000, (v) => v && !v.err && v.editorFound && !v.signedOut);
  if (!st) {
    console.error('TEAMS_NOT_READY ' + JSON.stringify(await evaluate(stateExpr).catch(() => null)));
    process.exit(3);
  }

  // 1b. navigate straight to the target chat if we know its id
  if (CHAT_URL) {
    console.log('navigating to chat deep link...');
    await send('Page.navigate', { url: CHAT_URL });
    await sleep(4000);
    st = await waitFor(stateExpr, 60000, (v) => v && !v.err && v.editorFound && !v.signedOut);
    if (!st) {
      console.error('NAV_FAILED ' + JSON.stringify(await evaluate(stateExpr).catch(() => null)));
      process.exit(3);
    }
  }

  // 2. ensure the right chat is open: wait for the chat rail to list it, click, verify heading
  let hd = await evaluate(headingExpr);
  if (!hd.matches) {
    console.log('waiting for chat list entry...');
    const found = await waitFor(listChatsExpr, 90000, (v) => Array.isArray(v) && v.length > 0);
    if (!found) {
      console.error('CHAT_NOT_IN_LIST for "' + nameQuery + '" ' + JSON.stringify(await evaluate(listChatsExpr)));
      process.exit(4);
    }
    console.log('chat list entry:', JSON.stringify(found[0]));
    await evaluate(clickChatExpr);
    hd = (await waitFor(headingExpr, 30000, (v) => v && v.matches)) || hd;
  }
  if (!hd.matches) {
    console.error('CHAT_OPEN_MISMATCH ' + JSON.stringify(hd));
    process.exit(4);
  }
  console.log('chat header ok:', JSON.stringify(hd.headings));

  // 3. dry run ends here
  if (!message) {
    console.log('DRY_RUN_OK: target chat open, editor ready, nothing sent');
    await screenshot(send);
    process.exit(0);
  }

  // 4. type via trusted CDP input
  if (!(await evaluate(focusEditorExpr))) {
    console.error('EDITOR_FOCUS_FAILED');
    process.exit(5);
  }
  await send('Input.insertText', { text: message });
  await sleep(800);
  const afterType = await evaluate(stateExpr);
  if (!afterType.editorText || !afterType.editorText.includes(message)) {
    console.error('TEXT_NOT_LANDED ' + JSON.stringify(afterType));
    process.exit(6);
  }
  console.log('typed ok, editor now:', JSON.stringify(afterType.editorText));

  // 5. Enter to send
  await send('Input.dispatchKeyEvent', {
    type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 13, text: '\r',
  });
  await send('Input.dispatchKeyEvent', {
    type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 13,
  });

  // 6. verify: compose box cleared
  const cleared = await waitFor(stateExpr, 10000, (v) => v && v.editorFound && !v.editorText);
  if (!cleared) {
    console.error('EDITOR_NOT_CLEARED ' + JSON.stringify(await evaluate(stateExpr).catch(() => null)));
    process.exit(7);
  }

  const fin = await evaluate(stateExpr);
  await screenshot(send);
  console.log('SEND_OK ' + JSON.stringify({ to: nameQuery, message, state: fin }));
  process.exit(0);
}

async function screenshot(send) {
  try {
    const shot = await send('Page.captureScreenshot', { format: 'png' });
    fs.writeFileSync(SHOT, Buffer.from(shot.data, 'base64'));
    console.log('screenshot: ' + SHOT);
  } catch (e) {
    console.log('screenshot failed: ' + e.message);
  }
}

main().catch((e) => {
  console.error('FATAL ' + (e?.stack || e));
  process.exit(1);
});
