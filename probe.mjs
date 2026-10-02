#!/usr/bin/env node
// Throwaway spike: find out which Teams operations your account is allowed to do via Microsoft Graph.
// Everything is read-only unless you explicitly answer "y" to a write test.
import readline from "node:readline/promises";
import { stdin, stdout } from "node:process";

const { CLIENT_ID, TENANT_ID } = process.env;
if (!CLIENT_ID || !TENANT_ID) {
  console.error("Missing CLIENT_ID or TENANT_ID. Copy .env.example to .env and fill it in (see README.md).");
  process.exit(1);
}

const GRAPH = "https://graph.microsoft.com/v1.0";
const LOGIN = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0`;

// Normally users can consent to these themselves.
const BASIC_SCOPES = ["User.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All", "Chat.ReadWrite", "ChannelMessage.Send"];
// These normally require a tenant admin to approve.
const ADMIN_SCOPES = ["ChannelMessage.Read.All", "TeamMember.Read.All"];

const AUTH_HINTS = {
  AADSTS65001: "consent was not granted",
  AADSTS90002: "tenant not found — check TENANT_ID",
  AADSTS90094: "admin approval required",
  AADSTS90099: "admin approval required",
  AADSTS50105: "your account is not assigned to this app",
  AADSTS50194: "app is single-tenant — set TENANT_ID to your tenant ID, not 'organizations'",
  AADSTS53003: "blocked by your company's Conditional Access policy",
  AADSTS700016: "app not found — check CLIENT_ID / TENANT_ID",
  AADSTS700038: "CLIENT_ID is not a valid app ID — copy it from the app registration Overview",
  AADSTS7000218: "enable 'Allow public client flows' in the app registration (Authentication tab)",
};

const C = {
  SIGNIN_BASIC: "Sign in with basic permissions",
  ME: "Read your profile",
  TEAMS: "List your teams",
  CHANNELS: "List channels of a team",
  CHATS: "List your chats",
  CHAT_MEMBERS: "List group chat members",
  CHAT_READ: "Read group chat messages + reactions",
  CHAT_SEND: "Send a message to the group chat",
  CHANNEL_SEND: "Post a message to the channel",
  DM_CREATE: "Open a 1:1 chat with a member (for reminders)",
  SIGNIN_ADMIN: "Sign in with channel-reading permissions",
  CHANNEL_READ: "Read channel messages + reactions",
  TEAM_MEMBERS: "List team members",
};

const rl = readline.createInterface({ input: stdin, output: stdout });
const results = new Map();

function record(name, ok, detail = "") {
  results.set(name, { ok, detail });
  const icon = ok === true ? "✅" : ok === false ? "❌" : "⏭️ ";
  console.log(`${icon} ${name}${detail ? ` — ${detail}` : ""}`);
}

async function check(name, fn) {
  try {
    record(name, true, (await fn()) ?? "");
    return true;
  } catch (e) {
    record(name, false, e.message);
    return false;
  }
}

function describeAuthError(text) {
  const code = text.match(/AADSTS\d+/)?.[0];
  const hint = code && AUTH_HINTS[code];
  return hint ? `${code}: ${hint}` : text.trim().split(/\r?\n| Trace ID:/)[0];
}

async function postForm(path, params) {
  const res = await fetch(`${LOGIN}${path}`, { method: "POST", body: new URLSearchParams({ client_id: CLIENT_ID, ...params }) });
  return res.json();
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// OAuth device-code sign-in, done by hand (not via MSAL) so the AADSTS error reasons stay visible.
// While waiting, pressing Enter cancels (for when the browser says "Need admin approval").
async function signIn(label, scopes) {
  console.log(`\n=== Sign-in: ${label} ===\nScopes: ${scopes.join(", ")}`);
  const code = await postForm("/devicecode", { scope: scopes.join(" ") });
  if (code.error) return { error: describeAuthError(code.error_description ?? code.error) };
  console.log(`\n${code.message}`);
  console.log(`If the browser shows "Need admin approval" or you can't continue, come back here and press Enter.`);

  const waiting = new AbortController();
  let cancelled = false;
  rl.question("", { signal: waiting.signal })
    .then(() => {
      cancelled = true;
      console.log("Cancelling…");
    })
    .catch(() => {});
  try {
    let interval = code.interval * 1000;
    while (!cancelled) {
      await sleep(interval);
      if (cancelled) break;
      const token = await postForm("/token", {
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        device_code: code.device_code,
      });
      if (token.access_token) return { token: token.access_token, scopes: token.scope.split(" ") };
      if (token.error === "slow_down") interval += 5000;
      else if (token.error !== "authorization_pending") return { error: describeAuthError(token.error_description ?? token.error) };
    }
    return { error: "cancelled by you — browser most likely required admin approval" };
  } finally {
    waiting.abort();
  }
}

async function graph(token, method, path, body) {
  const res = await fetch(GRAPH + path, {
    method,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: body && JSON.stringify(body),
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(`${res.status} ${json.error?.code ?? ""}: ${json.error?.message ?? res.statusText}`);
  return json;
}

async function pick(label, items, format) {
  if (!items.length) return null;
  console.log(`\n${label}:`);
  items.forEach((item, i) => console.log(`  ${i + 1}. ${format(item)}`));
  const n = Number(await rl.question(`Pick 1-${items.length} (Enter to skip): `));
  return n >= 1 && n <= items.length ? items[n - 1] : null;
}

async function confirm(question) {
  return /^y(es)?$/i.test((await rl.question(`${question} (y/N): `)).trim());
}

function summarizeReactions(messages) {
  const withReactions = messages.filter((m) => m.reactions?.length);
  const sample = withReactions[0]?.reactions
    .slice(0, 3)
    .map((r) => `${r.displayName ?? r.reactionType} by ${r.user?.user?.displayName ?? r.user?.user?.id ?? "?"}`)
    .join(", ");
  return `${messages.length} messages read, ${withReactions.length} with reactions${sample ? ` (e.g. ${sample})` : ""}`;
}

const chatLabel = (c) =>
  c.topic || (c.members ?? []).map((m) => m.displayName).filter(Boolean).slice(0, 4).join(", ") || "(untitled group chat)";

const TEST_MESSAGE = { body: { content: "🔧 Test message from teams-permission-probe — please ignore." } };

// ---------- Phase 1: basic permissions ----------
let me, team, channel, chat, chatMembers = [];
const basic = await signIn("basic (normally no admin needed)", BASIC_SCOPES);
if (basic.error) {
  record(C.SIGNIN_BASIC, false, basic.error);
} else {
  record(C.SIGNIN_BASIC, true);
  const t = basic.token;

  await check(C.ME, async () => {
    me = await graph(t, "GET", "/me");
    return `${me.displayName} <${me.userPrincipalName}>`;
  });

  let teams = [];
  await check(C.TEAMS, async () => {
    teams = (await graph(t, "GET", "/me/joinedTeams")).value;
    return `${teams.length} teams`;
  });
  team = await pick("Your teams", teams, (x) => x.displayName);

  let channels = [];
  if (team) {
    await check(C.CHANNELS, async () => {
      channels = (await graph(t, "GET", `/teams/${team.id}/channels`)).value;
      return `${channels.length} channels in "${team.displayName}"`;
    });
    channel = await pick("Channels", channels, (x) => x.displayName);
  } else record(C.CHANNELS, null, "no team picked");

  let chats = [];
  await check(C.CHATS, async () => {
    chats = (await graph(t, "GET", "/me/chats?$expand=members&$top=50")).value;
    return `${chats.length} recent chats, ${chats.filter((c) => c.chatType === "group").length} group chats`;
  });
  chat = await pick("Your group chats", chats.filter((c) => c.chatType === "group"), chatLabel);

  if (chat) {
    await check(C.CHAT_MEMBERS, async () => {
      chatMembers = (await graph(t, "GET", `/chats/${chat.id}/members`)).value;
      return `${chatMembers.length} members`;
    });
    await check(C.CHAT_READ, async () => summarizeReactions((await graph(t, "GET", `/chats/${chat.id}/messages?$top=20`)).value));

    if (await confirm(`\nSend a visible test message to "${chatLabel(chat)}"?`)) {
      await check(C.CHAT_SEND, () => graph(t, "POST", `/chats/${chat.id}/messages`, TEST_MESSAGE).then(() => "sent"));
    } else record(C.CHAT_SEND, null, "skipped");

    const other = chatMembers.find((m) => m.userId && m.userId !== me?.id);
    if (other && me && (await confirm(`Open a 1:1 chat with ${other.displayName}? (no message is sent; they won't see anything)`))) {
      const member = (id) => ({
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        roles: ["owner"],
        "user@odata.bind": `${GRAPH}/users('${id}')`,
      });
      await check(C.DM_CREATE, () =>
        graph(t, "POST", "/chats", { chatType: "oneOnOne", members: [member(me.id), member(other.userId)] }).then(() => "ok"),
      );
    } else record(C.DM_CREATE, null, "skipped");
  } else {
    for (const name of [C.CHAT_MEMBERS, C.CHAT_READ, C.CHAT_SEND, C.DM_CREATE]) record(name, null, "no group chat picked");
  }

  if (team && channel && (await confirm(`\nPost a visible test message to channel "${channel.displayName}"?`))) {
    await check(C.CHANNEL_SEND, () =>
      graph(t, "POST", `/teams/${team.id}/channels/${channel.id}/messages`, TEST_MESSAGE).then(() => "posted"),
    );
  } else record(C.CHANNEL_SEND, null, "skipped");
}

// ---------- Phase 2: permissions that usually need admin approval ----------
const admin = await signIn("channel reading (usually needs admin approval)", ADMIN_SCOPES);
if (admin.error) {
  record(C.SIGNIN_ADMIN, false, admin.error);
} else {
  record(C.SIGNIN_ADMIN, true);
  if (team && channel) {
    await check(C.CHANNEL_READ, async () =>
      summarizeReactions((await graph(admin.token, "GET", `/teams/${team.id}/channels/${channel.id}/messages?$top=20`)).value),
    );
  } else record(C.CHANNEL_READ, null, "no channel picked");
  if (team) {
    await check(C.TEAM_MEMBERS, async () => `${(await graph(admin.token, "GET", `/teams/${team.id}/members`)).value.length} members`);
  } else record(C.TEAM_MEMBERS, null, "no team picked");
}
rl.close();

// ---------- Report ----------
console.log("\n================ REPORT ================");
for (const [name, { ok, detail }] of results) {
  console.log(`${ok === true ? "✅" : ok === false ? "❌" : "⏭️ "} ${name}${ok === false ? ` — ${detail}` : ""}`);
}

const ok = (name) => results.get(name)?.ok;
console.log("\n---------------- VERDICT ----------------");
const SETUP_ERRORS = /AADSTS(700016|700038|90002|7000218|50194)\b/;
if (SETUP_ERRORS.test(results.get(C.SIGNIN_BASIC)?.detail ?? "")) {
  console.log("⚠️  Sign-in failed because of a setup problem (see the reason above) — fix it and run again. Nothing was learned yet.");
} else if (ok(C.SIGNIN_BASIC) === false) {
  console.log("❌ You can't sign in to your own app even with basic permissions → Graph-based tools (approach A/B) need IT.");
  console.log("   Next: check whether Power Automate is allowed (approach C).");
} else {
  console.log(ok(C.CHAT_READ)
    ? "✅ GROUP CHATS: you can read messages + reactions without admin → automation works in group chats."
    : "❌ GROUP CHATS: could not read chat reactions (or not tested).");
  console.log(ok(C.CHANNEL_READ)
    ? "✅ CHANNELS: you can read channel messages + reactions."
    : "❌ CHANNELS: reading channel reactions needs admin approval (ChannelMessage.Read.All) → use group chats, or ask IT.");
  console.log(ok(C.DM_CREATE)
    ? "✅ REMINDERS: you can open 1:1 chats → personal reminders are possible."
    : `${ok(C.DM_CREATE) === false ? "❌" : "⏭️ "} REMINDERS: 1:1 chat ${ok(C.DM_CREATE) === false ? "failed" : "not tested"}.`);
}
console.log("ℹ️  Bot approach (A): this script can't test it — check Teams → Apps → Manage your apps → 'Upload a custom app'.");
