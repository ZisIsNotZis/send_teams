# send_teams

Send Microsoft Teams 1:1 chat messages from this machine, without touching your
mouse/keyboard or stealing window focus.

## Working method — headless Edge + CDP (Plan A)

Reuses your real Edge session (cookies + token caches copied to a scratch
profile) inside an invisible headless Edge, driven over the Chrome DevTools
Protocol. The person is resolved via Teams global search and disambiguated by
department hint, then the message is typed with trusted CDP input events.

```bash
# 1. scratch profile copy (needed once per boot; ~30 MB, deleted after use)
SRC=~/.config/microsoft-edge DST=~/.cache/teams-headless-profile
rm -rf "$DST"; mkdir -p "$DST/Default"
cp "$SRC/Local State" "$DST/"
cp "$SRC/Default/Cookies" "$SRC/Default/Cookies-journal" "$SRC/Default/Preferences" "$DST/Default/"
cp -r "$SRC/Default/Local Storage" "$SRC/Default/Session Storage" "$SRC/Default/IndexedDB" "$DST/Default/"

# 2. boot invisible Edge with CDP
/usr/bin/microsoft-edge --headless=new --user-data-dir="$DST" \
  --remote-debugging-port=9333 --no-first-run --no-default-browser-check \
  "https://teams.microsoft.com/v2/" &

# 3. send (dry run when message is empty; 3rd arg = department hint for disambiguation)
node send_via_search.mjs "Wu, Fei" ""              # dry run
node send_via_search.mjs "Wu, Fei" "Hello"         # send
node send_via_search.mjs "Wu, Fei" "Hello" "FT D AA CN"
```

Notes:
- Recipient safety: search rows that look like files (`created by`, extensions),
  search-action rows (`from:`, `see more`), and group chats (`Name +N` in title)
  are excluded; ambiguous matches abort instead of guessing.
- Evidence: `evidence-headless.png` is captured on each run.
- Cleanup: kill the headless Edge and `rm -rf ~/.cache/teams-headless-profile`
  (the copy contains live session cookies — do not leave it around).
- Debug helpers: `probe_dom.mjs`, `dump_log.mjs`, `to_chat_surface.mjs`,
  `send_via_cdp.mjs` (older chat-rail approach).

## Blocked method — Microsoft Graph API (Plan B)

`teams.mjs` implements OAuth device-code and browser (PKCE) flows plus
chat resolve/send via Graph (`Chat.ReadWrite`). Currently unusable from this
tenant (Siemens AG):

- AADSTS530033: Conditional Access requires a Siemens-managed device; device-code
  flow cannot satisfy it (known protocol limitation).
- AADSTS50105: "Microsoft Graph Command Line Tools" requires admin-assigned
  access; you are not assigned.

Unblock path: ask IT to assign the "Microsoft Graph Command Line Tools"
enterprise app to you (or permit self-registered apps). Then run:

```bash
node teams.mjs login-browser      # one-time consent in this machine's browser
node teams.mjs send "Wu, Fei" "Hello"
```

Tokens are stored 0600 in `~/.local/state/send-teams/` and never logged.
