#!/usr/bin/env node
// Send Microsoft Teams chat messages via Microsoft Graph API (delegated).
//
// Auth: OAuth 2.0 device-code flow, first-party public client
// "Microsoft Graph Command Line Tools" (14d82eec-204b-4c2f-b7e8-296a70dab67e).
// Scopes: offline_access User.Read Chat.ReadWrite (user-consentable).
//
// Tokens are stored at ~/.local/state/send-teams/tokens.json with 0600 perms
// and are never printed. Pending device codes live in pending.json (0600).
//
// Usage:
//   node teams.mjs login-request          # step 1: get user code (prints USER_CODE / VERIFY_URL)
//   node teams.mjs login-poll             # step 2: poll until login completes, stores tokens
//   node teams.mjs login-browser          # alt: interactive browser login (auth-code + PKCE, loopback);
//                                         #      required when tenant CA blocks device-code flow
//   node teams.mjs whoami                 # sanity check: print my Graph identity
//   node teams.mjs find "<name>"          # resolve a 1:1 chat id by member display name
//   node teams.mjs send "<name>" "<msg>"  # resolve chat then send message
//   node teams.mjs logout                 # delete stored tokens

import fs from 'node:fs';
import http from 'node:http';
import crypto from 'node:crypto';
import path from 'node:path';
import os from 'node:os';

const CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const AUTH_BASE = 'https://login.microsoftonline.com/common/oauth2/v2.0';
const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPE = 'offline_access User.Read Chat.ReadWrite';
const STATE_DIR = path.join(os.homedir(), '.local', 'state', 'send-teams');
const TOKEN_FILE = path.join(STATE_DIR, 'tokens.json');
const PENDING_FILE = path.join(STATE_DIR, 'pending.json');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function ensureStateDir() {
  fs.mkdirSync(STATE_DIR, { recursive: true, mode: 0o700 });
}

function writeJson(file, data) {
  ensureStateDir();
  fs.writeFileSync(file, JSON.stringify(data, null, 1) + '\n', { mode: 0o600 });
}

function readJson(file) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return null;
  }
}

async function postForm(url, params) {
  const res = await fetch(url, { method: 'POST', body: new URLSearchParams(params) });
  const data = await res.json().catch(() => ({}));
  return { ok: res.ok, status: res.status, data };
}

async function cmdLoginRequest() {
  const { ok, status, data } = await postForm(`${AUTH_BASE}/devicecode`, {
    client_id: CLIENT_ID,
    scope: SCOPE,
  });
  if (!ok || !data.user_code) {
    console.error(`AUTH_REQUEST_FAILED status=${status} error=${data.error || 'unknown'}`);
    process.exit(1);
  }
  writeJson(PENDING_FILE, {
    device_code: data.device_code,
    expires_at: Date.now() + data.expires_in * 1000,
    interval: data.interval || 5,
  });
  console.log(`USER_CODE=${data.user_code}`);
  console.log(`VERIFY_URL=${data.verification_uri}`);
  if (data.verification_uri_complete) {
    console.log(`VERIFY_URL_COMPLETE=${data.verification_uri_complete}`);
  }
  console.error(`Login window: ${Math.round(data.expires_in / 60)} minutes.`);
}

async function cmdLoginBrowser() {
  // Interactive auth-code + PKCE with a loopback redirect. The browser session on
  // this machine carries device claims, so Conditional Access (e.g. Siemens'
  // "require managed device", error 530033) can pass here, unlike device-code
  // flow where the polling client presents no device identity.
  const verifier = crypto.randomBytes(48).toString('base64url');
  const challenge = crypto.createHash('sha256').update(verifier).digest('base64url');
  const state = crypto.randomBytes(16).toString('base64url');

  const server = http.createServer();
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  const port = server.address().port;
  const redirectUri = `http://localhost:${port}/`;
  const authUrl = `${AUTH_BASE}/authorize?` +
    new URLSearchParams({
      client_id: CLIENT_ID,
      response_type: 'code',
      redirect_uri: redirectUri,
      response_mode: 'query',
      scope: SCOPE,
      code_challenge: challenge,
      code_challenge_method: 'S256',
      state,
      prompt: 'select_account',
    });

  console.log(`AUTH_URL=${authUrl}`);
  console.error(`Waiting for browser redirect on 127.0.0.1:${port} (up to 8 minutes)...`);

  const timeout = setTimeout(() => {
    console.error('AUTH_BROWSER_TIMEOUT');
    process.exit(3);
  }, 8 * 60 * 1000);

  const code = await new Promise((resolve) => {
    server.on('request', (req, res) => {
      const u = new URL(req.url, `http://localhost:${port}`);
      const err = u.searchParams.get('error');
      const c = u.searchParams.get('code');
      if (u.searchParams.get('state') !== state) {
        res.writeHead(400).end('state mismatch');
        return;
      }
      if (err) {
        res.writeHead(200, { 'Content-Type': 'text/html' });
        res.end(`<html><body><h3>Login failed: ${err}</h3>You can close this window.</body></html>`);
        console.error(`AUTH_BROWSER_ERROR=${err}`);
        process.exit(4);
      }
      if (c) {
        res.writeHead(200, { 'Content-Type': 'text/html' });
        res.end('<html><body><h3>Login complete — you can close this window.</h3></body></html>');
        resolve(c);
      }
    });
  });

  clearTimeout(timeout);
  server.close();

  const { ok, data } = await postForm(`${AUTH_BASE}/token`, {
    client_id: CLIENT_ID,
    grant_type: 'authorization_code',
    code,
    redirect_uri: redirectUri,
    code_verifier: verifier,
    scope: SCOPE,
  });
  if (!ok || !data.access_token) {
    console.error(`AUTH_EXCHANGE_FAILED=${data.error || 'unknown'}`);
    process.exit(5);
  }
  writeJson(TOKEN_FILE, {
    access_token: data.access_token,
    refresh_token: data.refresh_token,
    scope: data.scope || SCOPE,
    expires_at: Date.now() + Math.max(0, (data.expires_in || 3600) - 60) * 1000,
    refreshed_at: new Date().toISOString(),
  });
  console.log('LOGIN_OK');
}

async function cmdLoginPoll() {
  const p = readJson(PENDING_FILE);
  if (!p || Date.now() > p.expires_at) {
    console.error('AUTH_PENDING_EXPIRED (run login-request again)');
    process.exit(3);
  }
  let interval = (p.interval || 5) * 1000;
  while (Date.now() < p.expires_at) {
    await sleep(interval);
    const { ok, data } = await postForm(`${AUTH_BASE}/token`, {
      client_id: CLIENT_ID,
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      device_code: p.device_code,
    });
    if (ok && data.access_token) {
      writeJson(TOKEN_FILE, {
        access_token: data.access_token,
        refresh_token: data.refresh_token,
        scope: data.scope || SCOPE,
        expires_at: Date.now() + Math.max(0, (data.expires_in || 3600) - 60) * 1000,
        refreshed_at: new Date().toISOString(),
      });
      fs.rmSync(PENDING_FILE, { force: true });
      console.log('LOGIN_OK');
      return;
    }
    if (data.error === 'authorization_pending') continue;
    if (data.error === 'slow_down') {
      interval += 5000;
      continue;
    }
    console.error(`AUTH_FAILED=${data.error || 'unknown'}`);
    process.exit(4);
  }
  console.error('AUTH_TIMEOUT');
  process.exit(3);
}

async function getAccessToken() {
  const t = readJson(TOKEN_FILE);
  if (!t) {
    console.error('NOT_LOGGED_IN (run: node teams.mjs login-request)');
    process.exit(5);
  }
  if (Date.now() < t.expires_at) return t.access_token;
  const { ok, data } = await postForm(`${AUTH_BASE}/token`, {
    client_id: CLIENT_ID,
    grant_type: 'refresh_token',
    refresh_token: t.refresh_token,
    scope: SCOPE,
  });
  if (!ok || !data.access_token) {
    console.error(`TOKEN_REFRESH_FAILED=${data.error || 'unknown'} (run login-request again)`);
    process.exit(6);
  }
  const nt = {
    access_token: data.access_token,
    refresh_token: data.refresh_token || t.refresh_token,
    scope: data.scope || t.scope,
    expires_at: Date.now() + Math.max(0, (data.expires_in || 3600) - 60) * 1000,
    refreshed_at: new Date().toISOString(),
  };
  writeJson(TOKEN_FILE, nt);
  return nt.access_token;
}

async function graphFetch(token, pathAndQuery, opts = {}) {
  const res = await fetch(GRAPH + pathAndQuery, {
    ...opts,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(opts.headers || {}),
    },
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`GRAPH_ERROR ${res.status}: ${text.slice(0, 400)}`);
  return text ? JSON.parse(text) : {};
}

const norm = (s) =>
  (s || '').toLowerCase().replace(/[(),]/g, ' ').replace(/\s+/g, ' ').trim();

async function findChat(token, query) {
  const q = norm(query);
  const filter = encodeURIComponent("chatType eq 'oneOnOne'");
  let url = `/me/chats?$filter=${filter}&$expand=members&$top=50`;
  while (url) {
    const page = await graphFetch(token, url);
    for (const chat of page.value || []) {
      const member = (chat.members || []).find((m) => norm(m.displayName).includes(q));
      if (member) {
        return {
          chatId: chat.id,
          displayName: member.displayName,
          memberCount: (chat.members || []).length,
        };
      }
    }
    const next = page['@odata.nextLink'];
    url = next ? next.slice(GRAPH.length) : null;
  }
  return null;
}

async function cmdWhoami() {
  const token = await getAccessToken();
  const me = await graphFetch(token, '/me?$select=displayName,userPrincipalName,mail');
  console.log(JSON.stringify(me));
}

async function cmdFind(query) {
  const token = await getAccessToken();
  const chat = await findChat(token, query);
  if (!chat) {
    console.error(`CHAT_NOT_FOUND for "${query}"`);
    process.exit(2);
  }
  console.log(JSON.stringify(chat));
}

async function cmdSend(query, message) {
  if (!query || !message) {
    console.error('usage: node teams.mjs send "<person name>" "<message>"');
    process.exit(1);
  }
  const token = await getAccessToken();
  const chat = await findChat(token, query);
  if (!chat) {
    console.error(`CHAT_NOT_FOUND for "${query}"`);
    process.exit(2);
  }
  const sent = await graphFetch(token, `/chats/${encodeURIComponent(chat.chatId)}/messages`, {
    method: 'POST',
    body: JSON.stringify({ body: { content: message } }),
  });
  console.log(
    JSON.stringify({
      ok: true,
      to: chat.displayName,
      chatId: chat.chatId,
      messageId: sent.id,
      createdDateTime: sent.createdDateTime,
      content: sent.body?.content,
    })
  );
}

const [cmd, ...args] = process.argv.slice(2);
switch (cmd) {
  case 'login-request':
    await cmdLoginRequest();
    break;
  case 'login-browser':
    await cmdLoginBrowser();
    break;
  case 'login-poll':
    await cmdLoginPoll();
    break;
  case 'whoami':
    await cmdWhoami();
    break;
  case 'find':
    await cmdFind(args[0]);
    break;
  case 'send':
    await cmdSend(args[0], args[1]);
    break;
  case 'logout':
    fs.rmSync(TOKEN_FILE, { force: true });
    console.log('LOGGED_OUT');
    break;
  default:
    console.error('unknown command; see header of teams.mjs');
    process.exit(1);
}
