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
  scrapeReservationsIndex,
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
  // Completed-reservations table: per-stay details (dates, party size,
  // confirmation code, payout) that the review pages don't expose. Joined to
  // each guest review by reservation_id below.
  const reservations = await scrapeReservationsIndex(page);
  console.log(`[scrape] reservations index: ${reservations.length} rows`);

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
    guest_reviews: guest_reviews.map((r) => {
      const resv = matchReservation(reservations, r.reservation_id);
      if (!resv) {
        // Unidentified stay: no confirmation_code, so nothing downstream can
        // check it against the exclusion list. Surface it rather than letting
        // it look like an ordinary draft.
        console.warn(`[scrape:warn] reservation_id=${r.reservation_id} guest="${r.guest_name}" は予約一覧に無い → 確認コード無しで生成`);
      }
      return {
        reservation_id: r.reservation_id,
        guest_name: r.guest_name,
        property_name: null,
        room_no: null,
        check_in: resv?.check_in_iso ?? null,
        check_out: resv?.check_out_iso ?? null,
        nights: resv?.nights ?? null,
        guests: resv?.guests ?? null,
        guests_label: resv?.guests_label ?? null,
        total_payout: resv?.total_payout ?? null,
        confirmation_code: resv?.confirmation_code ?? null, // join key to Airhost / exclusion list
        airbnb_status: resv?.airbnb_status ?? null,
        expires_soon: resv?.expires_soon ?? false,
        edit_href: r.href, // stored in context_json so post-approved can re-use
      };
    }),
  };
}

// Join a pending guest review to its row in the reservations index by
// reservation_id — the row links to its own /hosting/reviews/{id}/edit, which
// is the same id scrapePendingGuestReviews() returns.
//
// This used to match on guest name (containment, case-insensitive) because the
// review pages show only a first name. That was unsafe: two guests sharing a
// first name resolved to whichever row .find() hit first, silently attaching
// another guest's confirmation_code. The code is the join key to Airhost and
// (per DESIGN_review_exclusions.md) the key an exclusion list would hang off,
// so a wrong code means acting on the wrong reservation. Never match by name.
//
// Returns the enriched row (parsed ISO dates + nights) or null when the stay
// isn't in the table — callers must treat null as "unidentified", not "fine".
function matchReservation(reservations, reservationId) {
  if (!reservationId) return null;
  const hit = reservations.find((r) => r.reservation_id === reservationId);
  if (!hit) return null;
  const ci = parseAirbnbDate(hit.check_in);
  const co = parseAirbnbDate(hit.check_out);
  const nights = ci && co ? Math.round((co - ci) / 86400000) : null;
  return {
    check_in_iso: ci ? ci.toISOString().slice(0, 10) : hit.check_in || null,
    check_out_iso: co ? co.toISOString().slice(0, 10) : hit.check_out || null,
    nights: nights && nights > 0 ? nights : null,
    guests: hit.guests,
    guests_label: hit.guests_label || null,
    total_payout: hit.total_payout || null,
    confirmation_code: hit.confirmation_code || null,
    // Airbnb's own deadline signal ("Review guest - Expires soon"), kept verbatim.
    airbnb_status: hit.status || null,
    expires_soon: /expires soon/i.test(hit.status || ""),
  };
}

// "May 20, 2026" → Date (UTC). Returns null if unparseable.
function parseAirbnbDate(s) {
  if (!s) return null;
  const d = new Date(`${s} UTC`);
  return Number.isNaN(d.getTime()) ? null : d;
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

// ---------- オーナー言語 ----------
//
// ドラフトは「オーナーが読んで確認・編集する言語」で書く。オーナーは日本語話者
// とは限らない (例: MIM = 楊 世祥さん)。ポータルのUI文言だけ訳しても、承認する
// 本文が読めなければ意味がないので、生成側もこの設定を見る。
// 投稿直前に post-approved.mjs が英訳するので、掲載は言語によらず英語。
const OWNER_LANG_LABEL = {
  "ja": { name: "日本語", volume: "日本語で 200〜350字程度" },
  "zh-Hant": { name: "繁體中文 (台灣・香港で使われる正體字)", volume: "繁體中文で 150〜250字程度" },
};

function reviewSystemPrompt(ownerLang) {
  const L = OWNER_LANG_LABEL[ownerLang] ?? OWNER_LANG_LABEL["ja"];
  return `あなたは民泊運営会社 Dent Inc. のhost として、宿泊が完了したゲストへのレビュー(評価コメント)を書きます。この本文はオーナーが${L.name}で確認・編集し、投稿時に英訳して Airbnb に掲載されます。

# 文体ルール
- **${L.name}で書く** (オーナーが読んで確認・編集するため)
- ポジティブで具体的な内容
- ゲストの良かった点に触れる (清潔さ・コミュニケーション・チェックアウト時刻遵守 など)
- 物件の固有名詞には深入りしない (将来のゲストにも読まれる前提)
- 英語にして 80〜150 words 相当 (${L.volume})
- 絵文字は使わない
- 推薦 (recommend) は基本Yes 前提

# 出力形式
レビュー本文(${L.name})のみを出力。挨拶や説明文は不要。`;
}

// アカウント単位の表示言語。取得できなければ既定の ja に倒す
// (言語不明で生成を止めるより、従来どおり日本語で出す方が安全)。
async function loadOwnerLang() {
  try {
    const { lang } = await workerGet(`/internal/owner-lang/${ACCOUNT}`);
    return OWNER_LANG_LABEL[lang] ? lang : "ja";
  } catch (e) {
    console.warn(`[owner-lang] 取得失敗 (${e.message}) → ja で続行`);
    return "ja";
  }
}

function facilityContext(f) {
  if (!f) return "(物件マスタにマッチなし)";
  return [
    `物件名: ${f.name ?? f.code}`,
    `所在地: ${f.address ?? "(不明)"}`,
    `タイプ: ${f.property_type ?? "(不明)"} / ${f.floor_plan ?? ""} / ${f.sqm ?? "?"}㎡`,
    f.amenities?.nearest_station ? `最寄駅: ${f.amenities.nearest_station} 徒歩${f.amenities.walk_time_to_station ?? "?"}` : "",
  ].filter(Boolean).join("\n");
}

// The reply must match the guest's review language. A soft rule in the (all
// Japanese) system prompt isn't enough — the model drifts into Japanese for
// English reviews. So detect the language deterministically (isMostlyJapanese,
// defined below and hoisted) and force it with a prominent directive at the
// very end of the user turn (most salient spot).
async function generateReplyDraft({ item, facility }) {
  const ctx = facilityContext(facility);
  const reviewIsJa = isMostlyJapanese(item.original_text || "");
  const reviewIsZh = isMostlyChinese(item.original_text || "");
  const langDirective = reviewIsJa
    ? `【出力言語】このレビューは日本語です。返信も必ず日本語で書いてください。`
    : reviewIsZh
      ? `【CRITICAL — output language】This guest review is written in Chinese. You MUST write the public reply in Chinese, matching the review's script (Traditional review → Traditional reply, Simplified → Simplified). Do NOT reply in Japanese or English.`
      : `【CRITICAL — output language】This guest review is NOT written in Japanese. You MUST write the public reply in the SAME language as the review above (an English review → an English reply). Do NOT reply in Japanese.`;
  const user = [
    `# 物件情報`, ctx, ``,
    `# ゲスト情報`,
    `名前: ${item.guest_name}`,
    `宿泊月: ${item.stay_date ?? "(不明)"}`,
    `言語: ${item.language ?? (reviewIsJa ? "日本語" : reviewIsZh ? "中国語" : "ゲストのレビュー言語(非日本語)")}`,
    ``,
    `# 受信レビュー本文`,
    item.original_text,
    ``,
    `上記レビューへの公開返信を書いてください。`,
    ``,
    langDirective,
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

async function generateGuestReviewDraft({ item, facility, ownerLang }) {
  const ctx = facilityContext(facility);
  const L = OWNER_LANG_LABEL[ownerLang] ?? OWNER_LANG_LABEL["ja"];
  const user = [
    `# 物件情報`, ctx, ``,
    `# 宿泊情報`,
    `ゲスト名: ${item.guest_name}`,
    `チェックイン: ${item.check_in}`,
    `チェックアウト: ${item.check_out}`,
    `泊数: ${item.nights}`,
    ``,
    `上記ゲストへのホスト→ゲストレビューを書いてください (${L.name})。`,
  ].join("\n");

  const resp = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 1024,
    system: [
      { type: "text", text: reviewSystemPrompt(ownerLang), cache_control: { type: "ephemeral" } },
    ],
    messages: [{ role: "user", content: user }],
  });

  const text = resp.content.find((b) => b.type === "text")?.text ?? "";
  return { text, usage: resp.usage };
}

// ---------- language detection ----------
//
// Kana is the discriminator: Japanese prose effectively always contains it,
// Chinese never does. Counting CJK ideographs alone scores a Chinese review as
// Japanese, which replies to a Chinese guest in Japanese and suppresses the
// 参考訳 they need.
function hasKana(s) {
  return /[぀-ゟ゠-ヿ]/.test(String(s));
}

function isMostlyJapanese(s) {
  if (!hasKana(s)) return false;
  const jp = (String(s).match(/[぀-ヿ㐀-鿿]/g) || []).length;
  return jp / Math.max(1, String(s).length) > 0.15;
}

// Han characters with no kana anywhere → Chinese.
function isMostlyChinese(s) {
  if (hasKana(s)) return false;
  const han = (String(s).match(/[㐀-鿿]/g) || []).length;
  return han / Math.max(1, String(s).length) > 0.15;
}

// ---------- Japanese gloss for owner approval ----------
//
// The owner portal shows this as a 参考訳 so owners who don't read English can
// understand what they're approving. It is NOT posted to Airbnb — the English
// draft_text is what gets posted. Skip when the draft is already Japanese
// (e.g. a reply to a Japanese review).

// 参考訳はオーナーが読む言語で出す。日本語固定だと非日本語話者のオーナー
// (例: MIM = 楊 世祥さん) には返信内容が読めず、承認の判断ができない。
const TRANSLATE_SYSTEM = {
  "ja": `あなたは翻訳者です。与えられた Airbnb のレビュー/返信文を自然で読みやすい日本語に訳してください。訳文のみを出力し、説明・注釈・原文の再掲は不要です。`,
  "zh-Hant": `你是翻譯者。請將提供的 Airbnb 評價／回覆翻譯成自然流暢的繁體中文（正體字，台灣用語）。只輸出譯文，不要加說明、註解或重複原文。`,
};

// 既にオーナーの言語で書かれていれば訳す必要がない。
function alreadyInOwnerLang(text, ownerLang) {
  return ownerLang === "zh-Hant" ? isMostlyChinese(text) : isMostlyJapanese(text);
}

async function translateForOwner(text, ownerLang) {
  if (!text || alreadyInOwnerLang(text, ownerLang)) return null;
  try {
    const resp = await anthropic.messages.create({
      model: MODEL,
      max_tokens: 1024,
      system: [{
        type: "text",
        text: TRANSLATE_SYSTEM[ownerLang] ?? TRANSLATE_SYSTEM["ja"],
        cache_control: { type: "ephemeral" },
      }],
      messages: [{ role: "user", content: text }],
    });
    return resp.content.find((b) => b.type === "text")?.text?.trim() ?? null;
  } catch (e) {
    console.warn(`[translate] failed (${e.message}); skipping gloss`);
    return null;
  }
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
  // ドラフトを書く言語 = オーナーがポータルで読む言語。生成前に確定させる。
  const ownerLang = await loadOwnerLang();
  console.log(`[owner-lang] ${ACCOUNT} → ${ownerLang}`);

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
      if (DRY_RUN) {
        console.log(`[dry-run reply→${item.review_id}] guest="${item.guest_name}" review="${(item.original_text || "").slice(0, 60)}"\n  reply: ${text.slice(0, 120)}`);
        created++; continue;
      }
      const translation = await translateForOwner(text, ownerLang);
      await createDraft({
        account: ACCOUNT, draft_type: "reply", target_id: item.review_id,
        guest_name: item.guest_name, property_code: facility?.code ?? null,
        context: { property_name: item.property_name, room_no: item.room_no, original_text: item.original_text, language: item.language, translation, owner_lang: ownerLang },
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
      // draft_text is Japanese now (owner-facing). It gets translated to
      // English at post time (post-approved.mjs), so no separate translation_ja.
      const { text } = await generateGuestReviewDraft({ item, facility, ownerLang });
      if (DRY_RUN) { console.log(`[dry-run review→${item.reservation_id}]`, text); created++; continue; }
      await createDraft({
        account: ACCOUNT, draft_type: "review_of_guest", target_id: item.reservation_id,
        guest_name: item.guest_name, property_code: facility?.code ?? null,
        context: { owner_lang: ownerLang, property_name: item.property_name, room_no: item.room_no, check_in: item.check_in, check_out: item.check_out, nights: item.nights, guests: item.guests, guests_label: item.guests_label, total_payout: item.total_payout, confirmation_code: item.confirmation_code, airbnb_status: item.airbnb_status, expires_soon: item.expires_soon, edit_href: item.edit_href },
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
