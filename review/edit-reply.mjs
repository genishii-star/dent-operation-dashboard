/**
 * One-off corrector — overwrite an already-posted public reply on a single
 * review. Used to fix a reply that was posted to the wrong review / addressed
 * to the wrong guest (the old scraper mis-paired guest names to review_ids).
 *
 * Pulls the encrypted Airbnb session from Worker KV (same as post-approved.mjs),
 * restores Playwright, then runs editReply(review_id, REPLY_TEXT).
 *
 * Run:
 *   CF_ACCESS_CLIENT_ID=... CF_ACCESS_CLIENT_SECRET=... ACCOUNTS_JSON=... \
 *   REPLY_TEXT="Dear ..." \
 *   node edit-reply.mjs --account=NAGAI --review_id=1691126864175092551 [--dry-run] [--gui]
 */

import { chromium } from "playwright";
import { mkdirSync, readFileSync, writeFileSync, existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { ensureLoggedIn, editReply } from "./airbnb-helpers.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA_API = "https://api.dent-inc.com";

const argv = process.argv.slice(2);
const flag = (name, def = null) => {
  const hit = argv.find((a) => a.startsWith(`--${name}=`));
  if (hit) return hit.slice(name.length + 3);
  if (argv.includes(`--${name}`)) return true;
  return def;
};
const ACCOUNT = flag("account");
const REVIEW_ID = flag("review_id");
const DRY_RUN = flag("dry-run", false);
const GUI = flag("gui", false);
const REPLY_TEXT = process.env.REPLY_TEXT ?? "";

if (!ACCOUNT || !/^[A-Za-z0-9_-]{1,32}$/.test(ACCOUNT)) {
  console.error("Usage: node edit-reply.mjs --account=NAME --review_id=DIGITS [--dry-run]");
  process.exit(2);
}
if (!REVIEW_ID || !/^\d{6,25}$/.test(REVIEW_ID)) {
  console.error("missing/invalid --review_id (digits)");
  process.exit(2);
}
if (!REPLY_TEXT.trim()) {
  console.error("missing env REPLY_TEXT");
  process.exit(2);
}
for (const k of ["CF_ACCESS_CLIENT_ID", "CF_ACCESS_CLIENT_SECRET"]) {
  if (!process.env[k]) { console.error(`missing env: ${k}`); process.exit(2); }
}
const CF_HEADERS = {
  "CF-Access-Client-Id": process.env.CF_ACCESS_CLIENT_ID,
  "CF-Access-Client-Secret": process.env.CF_ACCESS_CLIENT_SECRET,
};

async function workerGet(path) {
  const r = await fetch(`${DATA_API}${path}`, { headers: CF_HEADERS });
  if (!r.ok) throw new Error(`GET ${path} ${r.status}: ${(await r.text()).slice(0, 200)}`);
  return r.json();
}
async function workerPost(path, body) {
  const r = await fetch(`${DATA_API}${path}`, {
    method: "POST",
    headers: { ...CF_HEADERS, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!r.ok) throw new Error(`POST ${path} ${r.status}: ${(await r.text()).slice(0, 200)}`);
  return r.json();
}

async function loadSession() {
  const { session } = await workerGet(`/internal/session/${ACCOUNT}`);
  const dir = join(__dirname, "sessions");
  if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
  const path = join(dir, `${ACCOUNT}.json`);
  writeFileSync(path, JSON.stringify(session, null, 2));
  return path;
}
async function saveSession(path) {
  const session = JSON.parse(readFileSync(path, "utf8"));
  await workerPost(`/internal/session/${ACCOUNT}`, { session });
}

function loadAccountCredentials() {
  const blob = process.env.ACCOUNTS_JSON;
  if (!blob) throw new Error("ACCOUNTS_JSON env var not set");
  const creds = JSON.parse(blob)[ACCOUNT];
  if (!creds?.email || !creds?.password) throw new Error(`ACCOUNTS_JSON missing email/password for ${ACCOUNT}`);
  return { email: creds.email, password: creds.password };
}
async function getTotpCode() {
  const r = await workerGet(`/internal/totp/${ACCOUNT}`);
  return r.code;
}
async function probeTotpAvailable() {
  try {
    const r = await fetch(`${DATA_API}/internal/totp/${ACCOUNT}`, { headers: CF_HEADERS });
    if (r.status === 404) return null;
    if (!r.ok) throw new Error(`probe ${r.status}`);
    return getTotpCode;
  } catch (e) {
    console.warn(`[totp] probe failed (${e.message}); assuming no TOTP`);
    return null;
  }
}

async function main() {
  console.log(`[edit] account=${ACCOUNT} review_id=${REVIEW_ID} dry_run=${!!DRY_RUN}`);
  console.log(`[edit] new reply (head): ${REPLY_TEXT.slice(0, 80)}...`);
  if (DRY_RUN) {
    console.log(JSON.stringify({ account: ACCOUNT, review_id: REVIEW_ID, edited: 0, dry_run: true }));
    return;
  }

  const sessionPath = await loadSession();
  const browser = await chromium.launch({ headless: !GUI });
  const context = await browser.newContext({ storageState: sessionPath });
  const page = await context.newPage();

  const totpFetcher = await probeTotpAvailable();
  await ensureLoggedIn(page, context, sessionPath, loadAccountCredentials, totpFetcher);

  const result = await editReply(page, { review_id: REVIEW_ID, draft_text: REPLY_TEXT });

  await context.storageState({ path: sessionPath });
  await saveSession(sessionPath);
  await browser.close();

  if (!result.ok) {
    console.error(`[edit] FAILED: ${result.error}`);
    console.log(JSON.stringify({ account: ACCOUNT, review_id: REVIEW_ID, edited: 0, error: result.error }));
    process.exit(1);
  }
  console.log(JSON.stringify({ account: ACCOUNT, review_id: REVIEW_ID, edited: 1 }));
}

main().catch((e) => { console.error(e); process.exit(1); });
