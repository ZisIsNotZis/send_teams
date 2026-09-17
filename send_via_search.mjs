#!/usr/bin/env node
// Send a Teams 1:1 message by finding the person via global search (surface-agnostic).
// Usage: node send_via_search.mjs "<person>" "<message>"    (empty message = dry run)

const PORT = 9333;
const [nameQuery = 'Wu, Fei', message = '', hint = 'FT D AA CN'] = process.argv.slice(2);
const SHOT = '/home/z/vibe/send_teams/evidence-headless.png';

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function main() {
  const targets = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
  const page = targets.find((t) => t.type === 'page' && /teams\.microsoft\.com/.test(t.url));
  if (!page) {
    console.error('NO_TEAMS_TARGET');
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
      const v = await evaluate(expr).catch((e) => ({ err: String(e).slice(0, 150) }));
      if (want(v)) return v;
      await sleep(1500);
    }
    return null;
  };
  const screenshot = async () => {
    try {
      const shot = await send('Page.captureScreenshot', { format: 'png' });
      fs.writeFileSync(SHOT, Buffer.from(shot.data, 'base64'));
      console.log('screenshot: ' + SHOT);
    } catch (e) {
      console.log('screenshot failed: ' + e.message);
    }
  };

  const HELP = String.raw`
    const norm = (s) => (s || '').toLowerCase().replace(/[(),]/g, ' ').replace(/\s+/g, ' ').trim();
    const JUNK = /created by|\.(xlsx|pdf|docx|pptx|zip|txt|one)|\+\s*\d+\s*$|^from:|^see more|in all results/;
    const isPersonRow = (t) => (/use arrow keys/.test(t) || /@/.test(t)) && !JUNK.test(t);
    const pickEditor = () => {
      const eds = [...document.querySelectorAll('div[contenteditable="true"]')];
      return eds.find((e) => /message/i.test(e.getAttribute('aria-label') || '')) || eds[eds.length - 1] || null;
    };
  `;
  const searchBoxExpr = `(() => { ${HELP}
    let inp = document.querySelector('[role="search"] input, input[type="search"], input[placeholder*="search" i]');
    if (!inp) {
      const expander = [...document.querySelectorAll('button')].find((b) => /^expand search/i.test(b.getAttribute('aria-label') || ''));
      if (expander) { expander.click(); return { expanded: true }; }
    }
    return { hasInput: !!inp, label: inp ? (inp.getAttribute('aria-label') || inp.getAttribute('placeholder') || '') : '' };
  })()`;
  const searchResultsExpr = `(() => { ${HELP} const q = norm(${JSON.stringify(nameQuery)}); const hint = norm(${JSON.stringify(hint)});
    let opts = [...document.querySelectorAll('[role="option"], [role="listbox"] [role="presentation"], [role="listbox"] button, [role="listbox"] li')]
      .filter((e) => norm(e.textContent).includes(q))
      .filter((e) => !JUNK.test(norm(e.textContent)));
    let persons = opts.filter((e) => isPersonRow(norm(e.textContent)));
    if (persons.length) opts = persons;
    if (hint) {
      const hinted = opts.filter((e) => norm(e.textContent).includes(hint));
      if (hinted.length) opts = hinted;
    }
    opts.sort((a, b) => norm(a.textContent).length - norm(b.textContent).length);
    return opts.slice(0, 5).map((e) => ({ tag: e.tagName, role: e.getAttribute('role'), text: norm(e.textContent).slice(0, 120) }));
  })()`;
  const clickResultExpr = `(() => { ${HELP} const q = norm(${JSON.stringify(nameQuery)}); const hint = norm(${JSON.stringify(hint)});
    let opts = [...document.querySelectorAll('[role="option"], [role="listbox"] [role="presentation"], [role="listbox"] button, [role="listbox"] li')]
      .filter((e) => norm(e.textContent).includes(q))
      .filter((e) => !JUNK.test(norm(e.textContent)));
    let persons = opts.filter((e) => isPersonRow(norm(e.textContent)));
    if (persons.length) opts = persons;
    if (hint) {
      const hinted = opts.filter((e) => norm(e.textContent).includes(hint));
      if (hinted.length) opts = hinted; else return { ambiguous: true, reason: 'no row matches hint', rows: opts.map((e) => norm(e.textContent).slice(0, 80)) };
    }
    if (opts.length !== 1) return { ambiguous: true, reason: 'multiple candidates', rows: opts.map((e) => norm(e.textContent).slice(0, 80)) };
    const el = opts[0];
    if (!el) return null;
    el.scrollIntoView({ block: 'center' });
    el.click();
    return { clicked: true };
  })()`;
  const stateExpr = `(() => { ${HELP}
    const ed = pickEditor();
    return {
      title: document.title,
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

  // 1. wait for the app to be interactive
  console.log('waiting for app...');
  const ready = await waitFor(searchBoxExpr, 120000, (v) => v && !v.err && (v.hasInput || v.expanded));
  if (!ready) {
    console.error('APP_NOT_READY ' + JSON.stringify(await evaluate(searchBoxExpr).catch(() => null)));
    process.exit(3);
  }

  // 2. reset any stale search state, then open + focus search via documented shortcut (trusted input)
  await send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27, nativeVirtualKeyCode: 27 });
  await send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27, nativeVirtualKeyCode: 27 });
  await sleep(800);
  await send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'e', code: 'KeyE', windowsVirtualKeyCode: 69, nativeVirtualKeyCode: 69, modifiers: 3 });
  await send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'e', code: 'KeyE', windowsVirtualKeyCode: 69, nativeVirtualKeyCode: 69, modifiers: 3 });
  await sleep(1500);
  const sb = await evaluate(searchBoxExpr);
  console.log('search box:', JSON.stringify(sb));

  // 2b. clear whatever text is in the search input (Ctrl+A, Delete)
  await send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'a', code: 'KeyA', windowsVirtualKeyCode: 65, nativeVirtualKeyCode: 65, modifiers: 2 });
  await send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'a', code: 'KeyA', windowsVirtualKeyCode: 65, nativeVirtualKeyCode: 65, modifiers: 2 });
  await send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Delete', code: 'Delete', windowsVirtualKeyCode: 46, nativeVirtualKeyCode: 46 });
  await send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Delete', code: 'Delete', windowsVirtualKeyCode: 46, nativeVirtualKeyCode: 46 });
  await sleep(800);

  // 3. type the query as trusted input
  await send('Input.insertText', { text: nameQuery });
  await sleep(3500);

  // 4. wait for person row, then click (hint-disambiguated, refuses ambiguity)
  const resultsReady = await waitFor(searchResultsExpr, 20000, (v) => Array.isArray(v) && v.length > 0);
  if (!resultsReady) {
    const raw = await evaluate(`(() => [...document.querySelectorAll('[role="option"]')].map((e) => e.textContent.slice(0, 90)))()`).catch(() => null);
    console.error('NO_MATCHING_ROWS raw options: ' + JSON.stringify(raw));
    await screenshot();
    process.exit(9);
  }
  console.log('results:', JSON.stringify(resultsReady));
  const clicked = await evaluate(clickResultExpr);
  if (!clicked || !clicked.clicked) {
    console.error('SEARCH_TARGET_NOT_UNIQUE ' + JSON.stringify(clicked));
    await screenshot();
    process.exit(4);
  }

  // 5. wait for the 1:1 chat to open (title shows person + hint, NOT a +N group)
  const hintNorm = hint.toLowerCase().replace(/[(),]/g, ' ').replace(/\s+/g, ' ').trim();
  const opened = await waitFor(stateExpr, 45000, (v) => {
    if (!v || !v.editorFound) return false;
    const t = v.title.toLowerCase().replace(/[(),]/g, ' ').replace(/\s+/g, ' ');
    return t.includes('wu fei') && t.includes(hintNorm) && !/,?\s*\+\s*\d+/.test(t);
  });
  if (!opened) {
    console.error('CHAT_NOT_OPENED ' + JSON.stringify(await evaluate(stateExpr).catch(() => null)));
    await screenshot();
    process.exit(5);
  }
  console.log('chat open:', JSON.stringify({ title: opened.title, editorLabel: opened.editorLabel }));

  // 6. dry run ends here
  if (!message) {
    console.log('DRY_RUN_OK: person chat open, editor ready, nothing sent');
    await screenshot();
    process.exit(0);
  }

  // 7. type + send
  if (!(await evaluate(focusEditorExpr))) {
    console.error('EDITOR_FOCUS_FAILED');
    process.exit(6);
  }
  await send('Input.insertText', { text: message });
  await sleep(800);
  const afterType = await evaluate(stateExpr);
  if (!afterType.editorText || !afterType.editorText.includes(message)) {
    console.error('TEXT_NOT_LANDED ' + JSON.stringify(afterType));
    process.exit(7);
  }
  console.log('typed ok, editor now:', JSON.stringify(afterType.editorText));

  await send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 13, text: '\r' });
  await send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 13 });

  const cleared = await waitFor(stateExpr, 10000, (v) => v && v.editorFound && !v.editorText);
  if (!cleared) {
    console.error('EDITOR_NOT_CLEARED ' + JSON.stringify(await evaluate(stateExpr).catch(() => null)));
    process.exit(8);
  }
  const fin = await evaluate(stateExpr);
  await screenshot();
  console.log('SEND_OK ' + JSON.stringify({ to: nameQuery, message, title: fin.title }));
  process.exit(0);
}

import fs from 'node:fs';
main().catch((e) => {
  console.error('FATAL ' + (e?.stack || e));
  process.exit(1);
});
