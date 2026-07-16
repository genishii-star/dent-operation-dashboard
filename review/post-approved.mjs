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
import Anthropic from "@anthropic-ai/sdk";
import {
  ensureLoggedIn,
  postReply as postReplyImpl,
  postGuestReview as postGuestReviewImpl,
} from "./airbnb-helpers.mjs";

// Host→guest review drafts are authored/edited in the OWNER's language (so the
// owner can actually read what they're approving — not every owner reads
// Japanese; MIM's owner reads Traditional Chinese) and translated to English
// here, at post time — so the posted English always reflects the owner's latest
// edit. Replies are posted in the guest's own language and are never translated.
const MODEL = "claude-haiku-4-5-20251001";
const TRANSLATE_SYSTEM = `You translate an Airbnb host-to-guest review into natural, warm English suitable for posting publicly on Airbnb. The source may be in any language. Keep it positive and concise. Output only the English translation — no notes, no quotes, no preamble.`;

// Any CJK content means the draft isn't English yet and must be translated
// before posting. Testing for "is it Japanese" would miss Chinese drafts and
// post 繁體中文 straight to Airbnb.
function needsEnglishTranslation(s) {
  const cjk = (String(s).match(/[぀-ヿ㐀-鿿가-힣]/g) || []).length;
  return cjk / Math.max(1, String(s).length) > 0.15;
}

async function translateToEn(text) {
  if (!process.env.ANTHROPIC_API_KEY) {
    throw new Error("ANTHROPIC_API_KEY required to translate the review to English");
  }
  const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
  const resp = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 1024,
    system: [{ type: "text", text: TRANSLATE_SYSTEM }],
    messages: [{ role: "user", content: text }],
  });
  const out = resp.content.find((b) => b.type === "text")?.text?.trim() ?? "";
  if (!out) throw new Error("translation returned empty");
  return out;
}

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

// ---------- 抑制ゲート (Gate 2) ----------
//
// 設計: operation/review/DESIGN_review_exclusions.md
//
// 承認済みでも投稿直前に必ず再判定する。承認から投稿までの間にクレームが入り、
// CSが抑制を立てるケースがあるため — 生成時のGate 1では捕まえられない。
// 承認を撤廃すると、ここが投稿を止められる最後の場所になる。
async function loadExclusions() {
  const { exclusions } = await workerGet(`/internal/review-exclusions?limit=1000`);
  const byCode = new Map();
  for (const e of exclusions ?? []) byCode.set(String(e.confirmation_code).toUpperCase(), e);
  console.log(`[exclusions] ${byCode.size}件 の有効な抑制を読み込み`);
  return byCode;
}

/**
 * 投稿してよいかを判定する。
 * @returns {{ post: true } | { post: false, reason: string }}
 */
function checkExclusion(draft, exclusions) {
  const ctx = draft.context_json ? JSON.parse(draft.context_json) : {};
  const code = ctx.confirmation_code ? String(ctx.confirmation_code).toUpperCase() : null;

  if (draft.draft_type !== "review_of_guest") {
    // 返信は既に公開されたレビューへの応答でゲストを新たに突く効果がない。
    // 止めるのは全接触を断つ scope='all' の時だけ。
    const hit = code ? exclusions.get(code) : null;
    return hit?.scope === "all"
      ? { post: false, reason: `抑制リストにより投稿せず (scope=all${hit.reason ? `: ${hit.reason}` : ""})` }
      : { post: true };
  }

  // fail-closed: 確認コードが無い = 身元を特定できない = 抑制リストに載っていない
  // ことを証明できない。損失が非対称 (レビュー1件の取り逃しは軽微だが、悪いレビューは
  // 公開されると恒久的に残る) なので、identify できないものは投稿しない。
  if (!code) {
    return { post: false, reason: "確認コード不明のため投稿を保留 (抑制リストと照合できない)" };
  }

  const hit = exclusions.get(code);
  if (hit) {
    return { post: false, reason: `抑制リストにより投稿せず (${hit.scope}${hit.reason ? `: ${hit.reason}` : ""})` };
  }
  return { post: true };
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
  // Owner-facing draft_text is Japanese → translate to English before posting.
  // (Older drafts already in English pass through unchanged.)
  let text = draft.draft_text;
  if (needsEnglishTranslation(text)) {
    text = await translateToEn(text);
    console.log(`[translate ${draft.id}] →EN: ${text.slice(0, 80)}...`);
  }
  return postGuestReviewImpl(page, {
    reservation_id: draft.target_id,
    editHref: ctx.edit_href ?? null,
    draft_text: text,
  });
}

// ----- main -----
async function main() {
  const { drafts: approved } = await workerGet(`/internal/review-drafts?account=${encodeURIComponent(ACCOUNT)}&status=approved&limit=200`);
  console.log(`[approved] ${approved.length} drafts to post for ${ACCOUNT}`);
  if (approved.length === 0) {
    console.log(JSON.stringify({ account: ACCOUNT, posted: 0, failed: 0, excluded: 0 }));
    return;
  }

  // 抑制ゲート。ブラウザを起動する前に落とす — 投稿してはいけないものを
  // 投稿経路に入れない。
  const exclusions = await loadExclusions();
  const drafts = [];
  let excluded = 0;
  for (const d of approved) {
    const verdict = checkExclusion(d, exclusions);
    if (verdict.post) { drafts.push(d); continue; }
    excluded++;
    console.log(`[excluded ${d.id}] ${d.draft_type} ${d.guest_name ?? ""} — ${verdict.reason}`);
    if (!DRY_RUN) {
      // 'failed' ではなく 'rejected'。抑制は意図した決定であってエラーではないので、
      // 障害シグナルを濁らせない。
      await workerPatch(`/internal/review-drafts/${d.id}`, {
        status: "rejected",
        decided_by: "system:exclusion",
        error_message: verdict.reason,
      });
    }
  }
  // 黙って抑制しない。全件抑制されているのに「順調」と誤認する事故を防ぐ。
  if (excluded) console.log(`[exclusions] ${excluded}件を抑制、残り ${drafts.length}件を投稿対象とする`);
  if (drafts.length === 0) {
    console.log(JSON.stringify({ account: ACCOUNT, posted: 0, failed: 0, excluded }));
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

  console.log(JSON.stringify({ account: ACCOUNT, posted, failed, excluded }));
}

main().catch((e) => { console.error(e); process.exit(1); });
