/**
 * Airbnb scrape + login + post helpers, ported from the existing review/ scripts:
 *   - scrapeUnrepliedReviews → from fetch-reviews.mjs STEP 4c + post-reply.mjs
 *   - scrapePendingGuestReviews → from fetch-reviews.mjs STEP 4
 *   - loginWithTotp + ensureLoggedIn → new (extends login.mjs auto-pilot with TOTP)
 *   - postReply → from post-reply.mjs (reload + aria-label verify + submit)
 *   - postGuestReview → from post-review.mjs (7-step wizard)
 *
 * Shared by generate-drafts.mjs and post-approved.mjs.
 */

/**
 * Open the host's reviews list. The new Airbnb UI is /performance/quality/overall/reviews/review/<id>
 * which shows the full list on the left and one review's detail on the right.
 * Pass reviewId=0 (or any sentinel) to just load the list — the right panel will
 * show "Something went wrong" which we don't care about for list scraping.
 */
export async function loadReviewsPage(page, reviewId = "0") {
  const url = `https://www.airbnb.com/performance/quality/overall/reviews/review/${reviewId}`;
  await page.goto(url, { waitUntil: "domcontentloaded", timeout: 30000 });
  await page.waitForTimeout(4000);
  // Lazy-load the list by scrolling
  let prevHeight = 0;
  for (let i = 0; i < 30; i++) {
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await page.waitForTimeout(1200);
    const h = await page.evaluate(() => document.body.scrollHeight);
    if (h === prevHeight) break;
    prevHeight = h;
  }
}

/** Wait until the right "Review details" panel shows content for the currently selected review. */
async function waitForReviewDetailsPanel(page, timeoutMs = 8000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const ready = await page.evaluate(() => {
      const heading = [...document.querySelectorAll("h1, h2, h3, [role='heading']")]
        .find((h) => (h.innerText || "").trim() === "Review details");
      if (!heading) return false;
      let p = heading;
      for (let j = 0; j < 10; j++) {
        p = p.parentElement;
        if (!p) return false;
        const text = (p.innerText || "");
        if (text.length > 200 && !text.includes("Something went wrong")) return true;
        if (p.querySelector("button") && text.length > 200) return true;
      }
      return false;
    });
    if (ready) return true;
    await page.waitForTimeout(500);
  }
  return false;
}

/** Classify the currently selected review (right panel) as 'unreplied', 'replied', or 'unknown'. */
async function classifyRightPanel(page) {
  return page.evaluate(() => {
    const heading = [...document.querySelectorAll("h1, h2, h3, [role='heading']")]
      .find((h) => (h.innerText || "").trim() === "Review details");
    if (!heading) return { status: "unknown", reason: "no Review details heading" };
    let panel = heading;
    for (let j = 0; j < 10; j++) {
      panel = panel.parentElement;
      if (!panel) break;
      if (panel.querySelector("button") && (panel.innerText || "").length > 200) break;
    }
    if (!panel) return { status: "unknown", reason: "no panel container" };
    const buttons = [...panel.querySelectorAll("button")].map((b) => (b.innerText || "").trim());
    if (buttons.includes("Write a public reply")) return { status: "unreplied", buttons };
    if (buttons.includes("Edit") || buttons.includes("Delete")) return { status: "replied", buttons };
    return { status: "unknown", buttons };
  });
}

/**
 * Extract unreplied reviews from "Reviews about you" tab.
 * Returns [{ guest_name, date, original_text }].
 *
 * Note: Airbnb does NOT expose a stable review_id in the public DOM for these
 * cards. The de-facto natural key is (guest_name + date). Callers should use
 * `${guest_name}::${date}` as the D1 `target_id`.
 */
/**
 * Walk the reviews list, opening each entry's right-panel detail to classify it.
 * Returns only the unreplied ones, each with the real Airbnb review_id (from URL).
 *
 * Returns [{ review_id, guest_name, date, original_text, property_name }].
 */
export async function scrapeUnrepliedReviews(page) {
  await loadReviewsPage(page);

  // Snapshot all card-level info from the left list (guest, date, property, rating)
  // before we start clicking, so we don't lose data when the panel updates.
  const cards = await page.evaluate(() => {
    const results = [];
    const seen = new Set();
    const viewBtns = [...document.querySelectorAll("button, a, span")]
      .filter((el) => (el.innerText || "").trim() === "View details");
    for (const btn of viewBtns) {
      // Walk up to find the card container
      let card = btn;
      for (let i = 0; i < 10; i++) {
        card = card.parentElement;
        if (!card) break;
        const text = (card.innerText || "");
        if (text.length > 80 && text.length < 4000 && text.includes("View details")) {
          // Stop at the smallest card-shaped container
          if (text.includes("Overall quality") || text.match(/Rating\s+\d/)) {
            const key = text.substring(0, 100);
            if (seen.has(key)) break;
            seen.add(key);
            const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
            // Line 0: guest name. Line 1: "<dates> • <property name>". Body follows.
            const guest = lines[0] || "";
            const dateLine = lines[1] || "";
            const dateMatch = dateLine.match(/^([A-Za-z]+\s+\d+\s*[-–]\s*\d+,?\s*\d{4})/) ||
                              dateLine.match(/^(\d{1,2}月\d+日\s*[-–]\s*\d+日)/);
            const date = dateMatch ? dateMatch[1] : "";
            const propMatch = dateLine.match(/[•·]\s*(.+)$/);
            const property_name = propMatch ? propMatch[1].trim() : "";
            // Body: lines after the rating/star line, before "View details"
            const bodyStart = lines.findIndex((l) => /^\d+$/.test(l)) + 1;
            const bodyEnd = lines.indexOf("View details");
            const body = bodyStart > 0 && bodyEnd > bodyStart
              ? lines.slice(bodyStart, bodyEnd).join(" ").substring(0, 2000)
              : "";
            results.push({ guest, date, property_name, body });
          }
          break;
        }
      }
    }
    return results;
  });

  // Click each View details one-by-one, harvest review_id from URL + status from panel
  const out = [];
  const N = cards.length;
  for (let i = 0; i < N; i++) {
    const clicked = await page.evaluate((idx) => {
      const els = [...document.querySelectorAll("button, a, span")]
        .filter((el) => (el.innerText || "").trim() === "View details");
      if (!els[idx]) return false;
      els[idx].scrollIntoView({ block: "center" });
      els[idx].click();
      return true;
    }, i);
    if (!clicked) continue;
    await page.waitForTimeout(1500);
    await waitForReviewDetailsPanel(page);
    const url = page.url();
    const m = url.match(/\/review\/(\d+)/);
    if (!m) continue;
    const review_id = m[1];
    const status = await classifyRightPanel(page);
    if (status.status === "unreplied") {
      const c = cards[i] || {};
      out.push({
        review_id,
        guest_name: c.guest || "",
        date: c.date || "",
        property_name: c.property_name || "",
        original_text: c.body || "",
      });
    }
  }
  return out;
}

/**
 * Extract pending guest reviews from /hosting (host→guest reviews not yet posted).
 * Returns [{ reservation_id, guest_name, href, raw_text }].
 *
 * reservation_id is extracted from href: /hosting/reviews/{id}/...
 * guest_name is parsed from text like "Leave Yours a review".
 */
export async function scrapePendingGuestReviews(page) {
  await page.goto("https://www.airbnb.com/hosting", { waitUntil: "domcontentloaded", timeout: 30000 });
  await page.waitForTimeout(3000);

  const raw = await page.evaluate(() =>
    [...document.querySelectorAll('a[href*="/hosting/reviews/"]')].map((a) => ({
      text: a.innerText.trim().substring(0, 400),
      href: a.href,
    })),
  );

  const results = [];
  for (const r of raw) {
    const idMatch = r.href.match(/\/hosting\/reviews\/([^/?#]+)/);
    if (!idMatch) continue;
    const reservation_id = idMatch[1];
    const nameMatch = r.text.match(/Leave\s+(.+?)\s+a review/);
    const guest_name = nameMatch ? nameMatch[1].trim() : null;
    if (!guest_name) continue;
    results.push({ reservation_id, guest_name, href: r.href, raw_text: r.text });
  }
  return results;
}

/**
 * Scrape the host "Completed" reservations table — the only place Airbnb
 * surfaces per-stay details (dates, party size, confirmation code, payout)
 * to the host. Used to enrich guest-review drafts so the owner can see the
 * stay context when approving.
 *
 * Returns [{ name, guests_label, guests, check_in, check_out,
 *            confirmation_code, total_payout }].
 *   - name           : full guest name as shown ("太晴 渡邊", "Daigo Ogasawara")
 *   - guests_label   : raw party string ("5 adults", "2 adults, 1 infant")
 *   - guests         : numeric total persons parsed from guests_label
 *   - check_in/out   : "May 20, 2026" style strings (left as-is; caller parses)
 *   - confirmation_code : "HM84JQA59J" — join key to Airhost reservation data
 *   - total_payout   : "¥23,744" string as shown
 */
export async function scrapeReservationsIndex(page) {
  await page.goto("https://www.airbnb.com/hosting/reservations/completed", {
    waitUntil: "domcontentloaded",
    timeout: 30000,
  });
  await page.waitForTimeout(5000);

  const rows = await page.evaluate(() => {
    const t = document.querySelector("table");
    if (!t) return [];
    const headers = [...t.querySelectorAll("thead th, thead td")].map((th) =>
      th.innerText.trim().toLowerCase(),
    );
    const col = (name) => headers.findIndex((h) => h === name.toLowerCase());
    const ix = {
      guests: col("Guests"),
      checkin: col("Check-in"),
      checkout: col("Checkout"),
      code: col("Confirmation Code"),
      payout: col("Total Payout"),
    };
    return [...t.querySelectorAll("tbody tr")].map((tr) => {
      const cells = [...tr.querySelectorAll("td,th")].map((td) => td.innerText.trim());
      return {
        guests_cell: cells[ix.guests] ?? "",
        check_in: cells[ix.checkin] ?? "",
        check_out: cells[ix.checkout] ?? "",
        confirmation_code: cells[ix.code] ?? "",
        total_payout: cells[ix.payout] ?? "",
      };
    });
  });

  return rows.map((r) => {
    const lines = r.guests_cell.split("\n").map((s) => s.trim()).filter(Boolean);
    const name = lines[0] ?? "";
    const guests_label = lines.slice(1).join(", "); // "5 adults" / "2 adults, 1 infant"
    const guests = (guests_label.match(/\d+/g) ?? []).reduce((a, n) => a + Number(n), 0) || null;
    return {
      name,
      guests_label,
      guests,
      check_in: r.check_in,
      check_out: r.check_out,
      confirmation_code: r.confirmation_code,
      total_payout: r.total_payout,
    };
  }).filter((r) => r.name);
}

// ============================================================
// Login + 2FA recovery
// ============================================================

/**
 * @param {object} opts
 * @param {string} opts.email
 * @param {string} opts.password
 * @param {() => Promise<string>} opts.fetchTotpCode  Returns the current 6-digit code.
 */
export async function loginWithTotp(page, { email, password, fetchTotpCode }) {
  await page.goto("https://www.airbnb.com/login", { waitUntil: "networkidle", timeout: 30000 });
  await page.waitForTimeout(3000);

  // "Continue with email" path
  const emailLink = page.locator(
    'div[data-testid="social-auth-button-email"], button:has-text("Continue with email"), a:has-text("Continue with email"), button:has-text("メールで続行")',
  );
  if ((await emailLink.count()) > 0) {
    await emailLink.first().click();
    await page.waitForTimeout(2000);
  }

  // Email
  const emailInput = page.locator("input:visible").first();
  await emailInput.fill(email);
  await page.waitForTimeout(500);
  await page.locator('button[type="submit"]:visible').first().click();
  await page.waitForTimeout(3000);

  // Password
  const pwInput = page.locator('input[type="password"]:visible');
  if ((await pwInput.count()) > 0) {
    await pwInput.first().fill(password);
    await page.waitForTimeout(500);
    await page.locator('button[type="submit"]:visible').first().click();
    await page.waitForTimeout(5000);
  }

  // 2FA detection: look for a code input or a "verify your identity" page.
  for (let attempt = 0; attempt < 2; attempt++) {
    const bodyText = await page.evaluate(() => document.body.innerText || "");
    const needsCode =
      /authentication code|verification code|認証コード|verify your identity|enter the code/i.test(bodyText);

    if (!needsCode) break;

    // Find the code input — usually a single small input or 6 separate digit inputs.
    const code = await fetchTotpCode();
    const filled = await page.evaluate((c) => {
      const setNative = (el, val) => {
        const proto = el.tagName === "TEXTAREA" ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
        const setter = Object.getOwnPropertyDescriptor(proto, "value").set;
        setter.call(el, val);
        el.dispatchEvent(new Event("input", { bubbles: true }));
      };
      const inputs = [...document.querySelectorAll('input:not([type="hidden"]):not([type="submit"]):not([type="button"])')]
        .filter((i) => i.offsetParent !== null);
      if (inputs.length === 1) { setNative(inputs[0], c); return "single"; }
      if (inputs.length >= 6) {
        const digits = c.split("");
        for (let i = 0; i < 6 && i < inputs.length; i++) setNative(inputs[i], digits[i]);
        return "multi";
      }
      return null;
    }, code);
    if (!filled) throw new Error("2FA input not found on page");
    await page.waitForTimeout(500);

    // Submit / verify button
    await page.evaluate(() => {
      const btn =
        document.querySelector('button[type="submit"]:not([disabled])') ??
        [...document.querySelectorAll("button")].find((b) =>
          /verify|continue|submit|送信|認証/i.test(b.innerText || ""),
        );
      if (btn) btn.click();
    });
    await page.waitForTimeout(5000);

    // If still on a 2FA page (e.g. wrong code due to drift), wait 30s for new code and retry once.
    const stillNeedsCode = /authentication code|verification code|認証コード/i.test(
      await page.evaluate(() => document.body.innerText || ""),
    );
    if (!stillNeedsCode) break;
    if (attempt === 0) {
      console.log("[login] 2FA challenge persists; waiting 30s for next TOTP window...");
      await page.waitForTimeout(30000);
    }
  }

  // "Remember this device" checkbox — increases session lifetime.
  await page.evaluate(() => {
    const candidates = [...document.querySelectorAll("input[type=checkbox]")].filter(
      (cb) => cb.offsetParent !== null && !cb.checked,
    );
    candidates.forEach((cb) => cb.click());
  }).catch(() => {});
  await page.waitForTimeout(1000);

  // Confirm login by checking we're not on /login anymore
  const url = page.url();
  if (url.includes("/login")) {
    throw new Error(`login failed — still at ${url}`);
  }
}

/**
 * Sentinel error message — workflows grep for this to decide whether to send
 * a "manual re-login required" Slack alert vs. a generic failure alert.
 */
export const SESSION_EXPIRED_NO_RECOVERY = "SESSION_EXPIRED_NO_AUTO_RECOVERY";

/**
 * Verify session is alive. If at a login wall and TOTP is available, run
 * loginWithTotp and persist the refreshed storageState. If TOTP is NOT
 * available (no seed registered — common when the account doesn't expose
 * a 2-step verification toggle in Airbnb's UI), throw a labeled error so
 * the caller can notify Slack and a human can re-run `login.mjs` locally.
 *
 * @param {() => Promise<{email: string, password: string}>} loadCredentials
 * @param {(() => Promise<string>) | null} fetchTotpCode  pass null when no TOTP seed exists
 * @param {string} sessionPath
 */
export async function ensureLoggedIn(page, context, sessionPath, loadCredentials, fetchTotpCode) {
  await page.goto("https://www.airbnb.com/hosting", { waitUntil: "domcontentloaded", timeout: 30000 });
  await page.waitForTimeout(3000);
  if (!page.url().includes("/login")) return false; // already logged in

  if (!fetchTotpCode) {
    throw new Error(
      `${SESSION_EXPIRED_NO_RECOVERY}: session expired but no TOTP seed is registered for this account. ` +
        `Run \`node login.mjs <account>\` locally and re-upload the session blob to the Worker.`,
    );
  }

  console.log("[auth] session expired — attempting TOTP re-login");
  const { email, password } = await loadCredentials();
  try {
    await loginWithTotp(page, { email, password, fetchTotpCode });
  } catch (e) {
    // TOTP path failed (wrong code, Airbnb pushed a different challenge, etc.)
    throw new Error(
      `${SESSION_EXPIRED_NO_RECOVERY}: TOTP re-login failed (${(e ).message}). ` +
        `Run \`node login.mjs <account>\` locally and re-upload the session blob.`,
    );
  }
  await context.storageState({ path: sessionPath });
  console.log("[auth] re-login complete; storageState refreshed");
  return true;
}

// ============================================================
// Posting
// ============================================================

/**
 * Post a public reply. Mirrors operation/review/post-reply.mjs:
 *   1. Reload reviews page (DOM-mutation safety)
 *   2. Close any open textareas
 *   3. Find the matching "Leave Public Response" button (smallest container w/ profile link)
 *   4. Verify textarea aria-label contains guest_name
 *   5. Fill + Submit
 *
 * @returns {Promise<{ok: boolean, error?: string}>}
 */
/**
 * Post a public reply via the new Airbnb UI.
 *   1. Navigate directly to /performance/quality/overall/reviews/review/<review_id>
 *      (auto-selects this review in the right panel)
 *   2. Wait for the right panel to render and verify status == 'unreplied'
 *      (if status == 'replied', someone else already replied — abort)
 *   3. Click "Write a public reply" → textarea appears (initially Save disabled)
 *   4. Fill textarea via native setter + dispatch input event (Airbnb listens to it
 *      to enable Save)
 *   5. Click Save once it becomes enabled
 *
 * @param {object} args
 * @param {string} args.review_id  Airbnb review_id (numeric string from URL)
 * @param {string} args.draft_text Reply body
 * @returns {Promise<{ok: boolean, error?: string}>}
 */
export async function postReply(page, { review_id, draft_text }) {
  if (!review_id) return { ok: false, error: "missing review_id" };
  if (!draft_text) return { ok: false, error: "missing draft_text" };

  // Direct goto = no list traversal, no risk of clicking the wrong card
  await loadReviewsPage(page, review_id);
  const panelReady = await waitForReviewDetailsPanel(page, 10000);
  if (!panelReady) return { ok: false, error: "right panel never rendered" };

  const status = await classifyRightPanel(page);
  if (status.status === "replied") {
    return { ok: false, error: "already replied (Edit/Delete present, not Write a public reply)" };
  }
  if (status.status !== "unreplied") {
    return { ok: false, error: `unexpected panel status: ${status.status}; buttons=${JSON.stringify(status.buttons)}` };
  }

  // Click "Write a public reply"
  const clicked = await page.evaluate(() => {
    const btn = [...document.querySelectorAll("button")]
      .find((b) => (b.innerText || "").trim() === "Write a public reply");
    if (!btn) return false;
    btn.scrollIntoView({ block: "center" });
    btn.click();
    return true;
  });
  if (!clicked) return { ok: false, error: "could not click 'Write a public reply'" };
  await page.waitForTimeout(1500);

  // Find the textarea that just appeared
  const taFound = await page.evaluate(() => {
    const tas = [...document.querySelectorAll("textarea")].filter((t) => t.offsetParent !== null);
    return tas.length > 0;
  });
  if (!taFound) return { ok: false, error: "textarea not found after clicking Write a public reply" };

  // Fill via native setter so React/internal state picks it up
  await page.evaluate((text) => {
    const ta = [...document.querySelectorAll("textarea")].find((t) => t.offsetParent !== null);
    if (!ta) return;
    const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value").set;
    setter.call(ta, text);
    ta.dispatchEvent(new Event("input", { bubbles: true }));
    ta.dispatchEvent(new Event("change", { bubbles: true }));
  }, draft_text);
  await page.waitForTimeout(1500);

  // Wait for Save to become enabled (max 5s)
  const enabled = await (async () => {
    const deadline = Date.now() + 5000;
    while (Date.now() < deadline) {
      const ok = await page.evaluate(() => {
        const btn = [...document.querySelectorAll("button")]
          .find((b) => (b.innerText || "").trim() === "Save");
        return !!btn && !btn.disabled;
      });
      if (ok) return true;
      await page.waitForTimeout(500);
    }
    return false;
  })();
  if (!enabled) return { ok: false, error: "Save button never became enabled" };

  // Click Save
  await page.evaluate(() => {
    const btn = [...document.querySelectorAll("button")]
      .find((b) => (b.innerText || "").trim() === "Save");
    if (btn) btn.click();
  });
  await page.waitForTimeout(3000);

  // Verify: panel should now show Edit/Delete (replied state)
  const after = await classifyRightPanel(page);
  if (after.status !== "replied") {
    return { ok: false, error: `post submitted but panel did not transition to 'replied' (got '${after.status}')` };
  }
  return { ok: true };
}

/**
 * Post a host→guest review via the 7-step wizard. Mirrors post-review.mjs.
 *
 * `editHref` is the /hosting/reviews/{reservation_id}/... link captured during
 * scrapePendingGuestReviews. Falls back to constructing it from reservation_id.
 *
 * @returns {Promise<{ok: boolean, error?: string}>}
 */
export async function postGuestReview(page, { reservation_id, editHref, draft_text }) {
  if (!draft_text) return { ok: false, error: "missing draft_text" };
  const url = editHref ?? `https://www.airbnb.com/hosting/reviews/${reservation_id}/edit`;
  await page.goto(url, { waitUntil: "domcontentloaded", timeout: 30000 });
  await page.waitForTimeout(5000);

  // Success markers shown once the review is in (immediate confirmation, or
  // the "already submitted" state if we re-enter an editor for a done review).
  const isSubmittedBody = (b) => {
    // Normalize whitespace first: Airbnb's confirmation copy uses &nbsp;
    // ( ) between words, so a plain-space substring check silently misses
    // even though the text looks identical. \s collapses nbsp too.
    const s = (b || "").replace(/\s+/g, " ").toLowerCase();
    return (
      s.includes("your review has been submitted") ||
      s.includes("thanks for your review") ||
      s.includes("already submitted a review") ||
      s.includes("you’ve already submitted") ||
      s.includes("you've already submitted")
    );
  };

  // Final safety net: when the wizard ends up in an unrecognized state, the
  // submission may actually have gone through (the post-submit confirmation
  // page has no "step N of M" and no radios, so it looks like an "unknown
  // step"). Reload the editor and let Airbnb tell us if the review is done.
  const confirmSubmitted = async () => {
    try {
      await page.goto(url, { waitUntil: "domcontentloaded", timeout: 30000 });
      await page.waitForTimeout(3000);
      const b = await page.evaluate(() => document.body.innerText || "");
      return isSubmittedBody(b);
    } catch {
      return false;
    }
  };

  for (let loop = 0; loop < 12; loop++) {
    const body = await page.evaluate(() => document.body.innerText || "");
    if (isSubmittedBody(body)) {
      return { ok: true };
    }

    const stepMatch = body.match(/step (\d+) out of (\d+)/i);
    const stepNum = stepMatch ? parseInt(stepMatch[1], 10) : -1;

    if (stepNum <= 0) {
      const hasContinue = await page.locator('button:has-text("Continue")').count();
      if (hasContinue > 0) {
        await page.click('button:has-text("Continue")');
        await page.waitForTimeout(3000);
        continue;
      }
      const hasRadio = await page.locator('input[type="radio"]').count();
      if (!hasRadio) {
        if (await confirmSubmitted()) return { ok: true };
        return { ok: false, error: `unknown initial step; body: ${body.substring(0, 200)}` };
      }
      // fall through to step 1-3 handler
    }

    if ((stepNum >= 1 && stepNum <= 3) || stepNum <= 0) {
      await page.evaluate(() => {
        const radios = [...document.querySelectorAll('input[type="radio"]')];
        if (radios.length > 0) radios[radios.length - 1].click(); // ★5
      });
      await page.waitForTimeout(500);
      await page.click('button:has-text("Next")');
      await page.waitForTimeout(3000);
      continue;
    }

    if (stepNum === 4) {
      await page.fill("textarea", draft_text);
      await page.waitForTimeout(500);
      await page.click('button:has-text("Next")');
      await page.waitForTimeout(3000);
      continue;
    }

    if (stepNum === 5) {
      await page.click('button:has-text("Yes")');
      await page.waitForTimeout(500);
      await page.click('button:has-text("Next")');
      await page.waitForTimeout(3000);
      continue;
    }

    if (stepNum === 6) {
      await page.click('button:has-text("Submit")');
      await page.waitForTimeout(5000);
      continue;
    }

    if (await confirmSubmitted()) return { ok: true };
    return { ok: false, error: `unknown step ${stepNum}; body: ${body.substring(0, 200)}` };
  }
  if (await confirmSubmitted()) return { ok: true };
  return { ok: false, error: "exceeded max wizard iterations" };
}
