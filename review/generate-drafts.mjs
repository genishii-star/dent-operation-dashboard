/**
 * Weekly draft generator — GitHub Actions entry point.
 *
 * Flow (per account):
 *   1. Pull encrypted session from Worker KV → write to sessions/{account}.json
 *   2. Launch Playwright, restore session
 *   3. Detect 2FA wall → fetch TOTP code from Worker, re-login, save new session back
 *   4. Scrape unreplied reviews + completed stays needing a guest review
 *   5. For each one not already in D1 with active status, generate a draft via
 *      Claude API and POST to /internal/review-drafts (status='pending')
 *   6. Save (possibly refreshed) session back to KV
 *
 * Approval happens in the dashboard (op.dent-inc.com → レビュータブ → 承認待ち).
 * The Slack #review-approval flow was retired 2026-05-25.
 *
 * Run:
 *   ANTHROPIC_API_KEY=... CF_ACCESS_CLIENT_ID=... CF_ACCESS_CLIENT_SECRET=... \
 *   node generate-drafts.mjs --account=NPA [--dry-run] [--gui]
 */

import { chromium } from "playwright";
import { mkdirSync, readFileSync, writeFileSync, existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import Anthropic from "@anthropic-ai/sdk";
import {
  scrapeUnrepliedReviews,
  scrapePendingGuestReviews,
  ensureLoggedIn,
} from "./airbnb-helpers.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA_API = "https://api.dent-inc.com";
const MODEL = "claude-opus-4-7";

// ---------- CLI ----------
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
  console.error("Usage: node generate-drafts.mjs --account=NAME [--dry-run] [--gui]");
  process.exit(2);
}

// ---------- env ----------
const ENV = ["ANTHROPIC_API_KEY", "CF_ACCESS_CLIENT_ID", "CF_ACCESS_CLIENT_SECRET", "ACCOUNTS_JSON"];
for (const k of ENV) if (!process.env[k]) { console.error(`missing env: ${k}`); process.exit(2); }
const {
  ANTHROPIC_API_KEY, CF_ACCESS_CLIENT_ID, CF_ACCESS_CLIENT_SECRET,
} = process.env;

const anthropic = new Anthropic({ apiKey: ANTHROPIC_API_KEY });

// ---------- Worker helpers ----------
const CF_HEADERS = {
  "CF-Access-Client-Id": CF_ACCESS_CLIENT_ID,
  "CF-Access-Client-Secret": CF_ACCESS_CLIENT_SECRET,
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

// ---------- session round-trip ----------
async function loadSession() {
  console.log(`[session] fetching from Worker for ${ACCOUNT}`);
  const { session } = await workerGet(`/internal/session/${ACCOUNT}`);
  const dir = join(__dirname, "sessions");
  if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
  const path = join(dir, `${ACCOUNT}.json`);
  writeFileSync(path, JSON.stringify(session, null, 2));
  return path;
}

async function saveSession(storageStatePath) {
  const session = JSON.parse(readFileSync(storageStatePath, "utf8"));
  console.log(`[session] saving back to Worker for ${ACCOUNT}`);
  await workerPost(`/internal/session/${ACCOUNT}`, { session });
}

async function getTotpCode() {
  const r = await workerGet(`/internal/totp/${ACCOUNT}`);
  return r.code;
}

// Returns getTotpCode if a seed is registered, else null. Used so accounts
// without a 2-step verification toggle (Airbnb has been quietly removing
// TOTP from the host UI in favor of passkeys) just fall back to "session
// expired → Slack alert → human re-runs login.mjs".
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

// ---------- scraping ----------
//
// Replies use the real Airbnb review_id from /performance/quality/overall/reviews/review/{id}.
// Guest reviews use the reservation_id extracted from the /hosting/reviews/{id} href.
async function scrapeWorkItems(page) {
  console.log("[scrape] unreplied reviews + pending guest reviews");
  // Sequential, not parallel — both use the same page object and would conflict.
  const replies = await scrapeUnrepliedReviews(page);
  const guest_reviews = await scrapePendingGuestReviews(page);
  // Normalize shape for downstream code.
  return {
    replies: replies.map((r) => ({
      review_id: r.review_id,
      guest_name: r.guest_name,
      property_name: r.property_name || null,
      room_no: null,
      original_text: r.original_text,
      language: null, // Claude will infer from original_text
      stay_date: r.date,
    })),
    guest_reviews: guest_reviews.map((r) => ({
      reservation_id: r.reservation_id,
      guest_name: r.guest_name,
      property_name: null,
      room_no: null,
      check_in: null,
      check_out: null,
      nights: null,
      edit_href: r.href, // stored in context_json so post-approved can re-use
    })),
  };
}

// ---------- credentials (for 2FA recovery) ----------
function loadAccountCredentials() {
  const blob = process.env.ACCOUNTS_JSON;
  if (!blob) throw new Error("ACCOUNTS_JSON env var not set (need {account: {email, password}})");
  let parsed;
  try { parsed = JSON.parse(blob); } catch { throw new Error("ACCOUNTS_JSON is not valid JSON"); }
  const creds = parsed[ACCOUNT];
  if (!creds?.email || !creds?.password) {
    throw new Error(`ACCOUNTS_JSON missing email/password for ${ACCOUNT}`);
  }
  return { email: creds.email, password: creds.password };
}

// ---------- facility lookup ----------
let _facilities = null;
async function getFacility(propName, roomNo) {
  if (!_facilities) {
    const { facilities } = await workerGet("/internal/facilities.json");
    _facilities = facilities ?? [];
  }
  const key = (s) => String(s ?? "").trim().toLowerCase();
  // Try exact match by code (single-room) then by name
  return (
    _facilities.find((f) => key(f.code) === key(propName)) ??
    _facilities.find((f) => key(f.name) === key(propName)) ??
    _facilities.find((f) => Array.isArray(f.rooms) && f.rooms.some((r) => key(r) === key(propName + roomNo))) ??
    null
  );
}

// ---------- Claude API: draft generation ----------
//
// We split the system prompt so the long, stable portion gets prompt-cached
// across all drafts in the same run (and across weeks until we change the
// template). The per-draft details go in the user turn, which is uncached.

const REPLY_SYSTEM = `あなたは民泊運営会社 Dent Inc. のホストとして、宿泊ゲストから届いたAirbnbレビューに公開返信を書きます。

# 文体ルール
- ゲストのレビュー本文と同じ言語で返信する (英語のレビュー→英語、日本語→日本語、等)
- 丁寧だが堅すぎない、親しみのあるトーン
- ゲストの名前をオープニングに含める (英語なら "Dear {name}," / 日本語なら "{name}様、")
- 物件の良かった点をゲストが触れていれば、それを引いて感謝する (オウム返しに見えないよう、別の言葉で言い換える)
- ネガティブな指摘があった場合は、まず謝罪し、具体的に何を改善したか/するかに触れる (虚偽の約束はしない)
- 末尾は次回再訪を歓迎する一文で締める
- 150〜300字程度 / 英語なら100〜200 words 程度
- 絵文字は使わない
- Dent Inc. を会社名として明示しない (一個人のホストとして書く)

# 出力形式
返信本文のみを出力。挨拶や説明文は不要。`;

const REVIEW_SYSTEM = `あなたは民泊運営会社 Dent Inc. のホストとして、宿泊が完了したゲストに対するレビューを Airbnb に投稿します。

# 文体ルール
- 英語で書く (Airbnb の host→guest レビューは多言語ゲスト向けの汎用性のため英語が望ましい)
- ポジティブで具体的な内容
- ゲストの良かった点に触れる (清潔さ・コミュニケーション・チェックアウト時刻遵守 など)
- 物件の固有名詞には深入りしない (将来のゲストにも読まれる前提)
- 80〜150 words 程度
- 絵文字は使わない
- 推薦 (recommend) は基本Yes 前提

# 出力形式
レビュー本文のみを出力。挨拶や説明文は不要。`;

function facilityContext(f) {
  if (!f) return "(物件マスタにマッチなし)";
  return [
    `物件名: ${f.name ?? f.code}`,
    `所在地: ${f.address ?? "(不明)"}`,
    `タイプ: ${f.property_type ?? "(不明)"} / ${f.floor_plan ?? ""} / ${f.sqm ?? "?"}㎡`,
    f.amenities?.nearest_station ? `最寄駅: ${f.amenities.nearest_station} 徒歩${f.amenities.walk_time_to_station ?? "?"}` : "",
  ].filter(Boolean).join("\n");
}

async function generateReplyDraft({ item, facility }) {
  const ctx = facilityContext(facility);
  const user = [
    `# 物件情報`, ctx, ``,
    `# ゲスト情報`,
    `名前: ${item.guest_name}`,
    `宿泊月: ${item.stay_date ?? "(不明)"}`,
    `言語: ${item.language ?? "(原文から推測)"}`,
    ``,
    `# 受信レビュー本文`,
    item.original_text,
    ``,
    `上記レビューへの公開返信を書いてください。`,
  ].join("\n");

  const resp = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 1024,
    system: [
      { type: "text", text: REPLY_SYSTEM, cache_control: { type: "ephemeral" } },
    ],
    messages: [{ role: "user", content: user }],
  });

  const text = resp.content.find((b) => b.type === "text")?.text ?? "";
  return { text, usage: resp.usage };
}

async function generateGuestReviewDraft({ item, facility }) {
  const ctx = facilityContext(facility);
  const user = [
    `# 物件情報`, ctx, ``,
    `# 宿泊情報`,
    `ゲスト名: ${item.guest_name}`,
    `チェックイン: ${item.check_in}`,
    `チェックアウト: ${item.check_out}`,
    `泊数: ${item.nights}`,
    ``,
    `上記ゲストへのホスト→ゲストレビューを書いてください (英語)。`,
  ].join("\n");

  const resp = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 1024,
    system: [
      { type: "text", text: REVIEW_SYSTEM, cache_control: { type: "ephemeral" } },
    ],
    messages: [{ role: "user", content: user }],
  });

  const text = resp.content.find((b) => b.type === "text")?.text ?? "";
  return { text, usage: resp.usage };
}

// ---------- D1 dedupe + write ----------
async function alreadyHasActiveDraft({ account, draft_type, target_id }) {
  // We don't have a "by target" endpoint; just list pending+approved and check locally.
  // For low volumes this is fine. Optimize when a single account exceeds ~200 active drafts.
  const [pend, appr] = await Promise.all([
    workerGet(`/internal/review-drafts?account=${encodeURIComponent(account)}&status=pending&limit=500`),
    workerGet(`/internal/review-drafts?account=${encodeURIComponent(account)}&status=approved&limit=500`),
  ]);
  const all = [...(pend.drafts ?? []), ...(appr.drafts ?? [])];
  return all.some((d) => d.draft_type === draft_type && d.target_id === target_id);
}

async function createDraft({ account, draft_type, target_id, guest_name, property_code, context, draft_text }) {
  return workerPost("/internal/review-drafts", {
    account, draft_type, target_id, guest_name, property_code,
    context, draft_text,
  });
}

// ---------- main ----------
async function main() {
  const sessionPath = await loadSession();

  const browser = await chromium.launch({ headless: !GUI });
  const context = await browser.newContext({ storageState: sessionPath });
  const page = await context.newPage();

  // If session expired: auto-recover via TOTP when a seed is registered,
  // otherwise throw SESSION_EXPIRED_NO_AUTO_RECOVERY for the workflow to surface to Slack.
  const totpFetcher = await probeTotpAvailable();
  await ensureLoggedIn(page, context, sessionPath, loadAccountCredentials, totpFetcher);

  const work = await scrapeWorkItems(page);
  console.log(`[scrape] replies: ${work.replies.length}, guest_reviews: ${work.guest_reviews.length}`);

  let created = 0, skipped = 0, failed = 0;

  for (const item of work.replies) {
    try {
      const exists = await alreadyHasActiveDraft({
        account: ACCOUNT, draft_type: "reply", target_id: item.review_id,
      });
      if (exists) { skipped++; continue; }
      const facility = await getFacility(item.property_name, item.room_no);
      const { text } = await generateReplyDraft({ item, facility });
      if (DRY_RUN) { console.log(`[dry-run reply→${item.review_id}]`, text); created++; continue; }
      await createDraft({
        account: ACCOUNT, draft_type: "reply", target_id: item.review_id,
        guest_name: item.guest_name, property_code: facility?.code ?? null,
        context: { property_name: item.property_name, room_no: item.room_no, original_text: item.original_text, language: item.language },
        draft_text: text,
      });
      created++;
    } catch (e) {
      failed++;
      console.error(`[reply ${item.review_id}] ${e.message}`);
    }
  }

  for (const item of work.guest_reviews) {
    try {
      const exists = await alreadyHasActiveDraft({
        account: ACCOUNT, draft_type: "review_of_guest", target_id: item.reservation_id,
      });
      if (exists) { skipped++; continue; }
      const facility = await getFacility(item.property_name, item.room_no);
      const { text } = await generateGuestReviewDraft({ item, facility });
      if (DRY_RUN) { console.log(`[dry-run review→${item.reservation_id}]`, text); created++; continue; }
      await createDraft({
        account: ACCOUNT, draft_type: "review_of_guest", target_id: item.reservation_id,
        guest_name: item.guest_name, property_code: facility?.code ?? null,
        context: { property_name: item.property_name, room_no: item.room_no, check_in: item.check_in, check_out: item.check_out, nights: item.nights, edit_href: item.edit_href },
        draft_text: text,
      });
      created++;
    } catch (e) {
      failed++;
      console.error(`[review ${item.reservation_id}] ${e.message}`);
    }
  }

  // Save (possibly refreshed) cookies back to KV — but only in non-dry-run mode.
  // In dry-run we don't want to clobber a known-good session with whatever state
  // happened to be in the context when scraping finished.
  if (!DRY_RUN) {
    await context.storageState({ path: sessionPath });
    await saveSession(sessionPath);
  }
  await browser.close();

  console.log(JSON.stringify({ account: ACCOUNT, created, skipped, failed }));
}

main().catch((e) => { console.error(e); process.exit(1); });
