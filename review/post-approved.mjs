/**
 * Weekly poster — GitHub Actions entry point.
 *
 * Drains status='approved' rows from review_drafts:
 *   1. Pull encrypted session from Worker KV → restore Playwright context
 *   2. GET /internal/review-drafts?account=X&status=approved
 *   3. For each draft, post via Airbnb UI (reply or guest review)
 *      - Page reload + target validation before submit (DOM mutation safety —
 *        see memory feedback_dom_mutation)
 *   4. PATCH status → 'posted' or 'failed'
 *   5. Save (possibly refreshed) session back to KV
 *
 * Run:
 *   CF_ACCESS_CLIENT_ID=... CF_ACCESS_CLIENT_SECRET=... \
 *   node post-approved.mjs --account=NPA [--dry-run] [--gui]
 *
 * Posting helpers are placeholders. Phase 1.5 plugs in the existing
 * post-reply.mjs / post-review.mjs logic.
 */

import { chromium } from "playwright";
import { mkdirSync, readFileSync, writeFileSync, existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import {
  ensureLoggedIn,
  postReply as postReplyImpl,
  postGuestReview as postGuestReviewImpl,
} from "./airbnb-helpers.mjs";

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
const DRY_RUN = flag("dry-run", false);
const GUI = flag("gui", false);

if (!ACCOUNT || !/^[A-Za-z0-9_-]{1,32}$/.test(ACCOUNT)) {
  console.error("Usage: node post-approved.mjs --account=NAME [--dry-run] [--gui]");
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
async function workerPatch(path, body) {
  const r = await fetch(`${DATA_API}${path}`, {
    method: "PATCH",
    headers: { ...CF_HEADERS, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!r.ok) throw new Error(`PATCH ${path} ${r.status}: ${(await r.text()).slice(0, 200)}`);
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

// ----- credentials (for 2FA recovery) -----
function loadAccountCredentials() {
  const blob = process.env.ACCOUNTS_JSON;
  if (!blob) throw new Error("ACCOUNTS_JSON env var not set");
  const parsed = JSON.parse(blob);
  const creds = parsed[ACCOUNT];
  if (!creds?.email || !creds?.password) {
    throw new Error(`ACCOUNTS_JSON missing email/password for ${ACCOUNT}`);
  }
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
    if (!r.ok) throw new Error(`probe ${r.status}: ${(await r.text()).slice(0, 200)}`);
    return getTotpCode;
  } catch (e) {
    console.warn(`[totp] probe failed (${e.message}); assuming no TOTP available`);
    return null;
  }
}

// ----- posters -----
async function postReply(page, draft) {
  return postReplyImpl(page, { review_id: draft.target_id, draft_text: draft.draft_text });
}

async function postGuestReview(page, draft) {
  const ctx = draft.context_json ? JSON.parse(draft.context_json) : {};
  return postGuestReviewImpl(page, {
    reservation_id: draft.target_id,
    editHref: ctx.edit_href ?? null,
    draft_text: draft.draft_text,
  });
}

// ----- main -----
async function main() {
  const { drafts } = await workerGet(`/internal/review-drafts?account=${encodeURIComponent(ACCOUNT)}&status=approved&limit=200`);
  console.log(`[approved] ${drafts.length} drafts to post for ${ACCOUNT}`);
  if (drafts.length === 0) {
    console.log(JSON.stringify({ account: ACCOUNT, posted: 0, failed: 0 }));
    return;
  }

  const sessionPath = await loadSession();
  const browser = await chromium.launch({ headless: !GUI });
  const context = await browser.newContext({ storageState: sessionPath });
  const page = await context.newPage();

  const totpFetcher = await probeTotpAvailable();
  await ensureLoggedIn(page, context, sessionPath, loadAccountCredentials, totpFetcher);

  let posted = 0, failed = 0;

  for (const draft of drafts) {
    if (DRY_RUN) {
      console.log(`[dry-run ${draft.id}] would post ${draft.draft_type} ${draft.target_id}`);
      continue;
    }

    let result;
    try {
      // Force a fresh page state before every submission.
      await page.goto("about:blank");
      if (draft.draft_type === "reply") {
        result = await postReply(page, draft);
      } else if (draft.draft_type === "review_of_guest") {
        result = await postGuestReview(page, draft);
      } else {
        result = { ok: false, error: `unknown draft_type: ${draft.draft_type}` };
      }
    } catch (e) {
      result = { ok: false, error: e.message };
    }

    try {
      if (result.ok) {
        await workerPatch(`/internal/review-drafts/${draft.id}`, { status: "posted" });
        posted++;
      } else {
        await workerPatch(`/internal/review-drafts/${draft.id}`, {
          status: "failed",
          error_message: result.error ?? "unknown error",
        });
        failed++;
        console.error(`[post ${draft.id}] FAILED: ${result.error}`);
      }
    } catch (e) {
      // Status update failed — likely a race (another runner finished it).
      // Don't double-count; log and continue.
      console.error(`[post ${draft.id}] status update failed: ${e.message}`);
    }
  }

  await context.storageState({ path: sessionPath });
  await saveSession(sessionPath);
  await browser.close();

  console.log(JSON.stringify({ account: ACCOUNT, posted, failed }));
}

main().catch((e) => { console.error(e); process.exit(1); });
