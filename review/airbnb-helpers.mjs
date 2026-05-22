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

/** Goto /users/reviews and scroll until fully loaded. */
export async function loadReviewsPage(page) {
  await page.goto("https://www.airbnb.com/users/reviews", { waitUntil: "domcontentloaded", timeout: 30000 });
  await page.waitForTimeout(3000);
  let prevHeight = 0;
  for (let i = 0; i < 20; i++) {
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await page.waitForTimeout(1500);
    const h = await page.evaluate(() => document.body.scrollHeight);
    if (h === prevHeight) break;
    prevHeight = h;
  }
}

async function clickTab(page, tabName) {
  await page.evaluate((name) => {
    const btns = [...document.querySelectorAll("button")];
    const tab = btns.find((b) => b.innerText.trim() === name);
    if (tab) tab.click();
  }, tabName);
  await page.waitForTimeout(3000);
}

/**
 * Extract unreplied reviews from "Reviews about you" tab.
 * Returns [{ guest_name, date, original_text }].
 *
 * Note: Airbnb does NOT expose a stable review_id in the public DOM for these
 * cards. The de-facto natural key is (guest_name + date). Callers should use
 * `${guest_name}::${date}` as the D1 `target_id`.
 */
export async function scrapeUnrepliedReviews(page) {
  await loadReviewsPage(page);
  await clickTab(page, "Reviews about you");
  let prevHeight = 0;
  for (let i = 0; i < 20; i++) {
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await page.waitForTimeout(1500);
    const h = await page.evaluate(() => document.body.scrollHeight);
    if (h === prevHeight) break;
    prevHeight = h;
  }

  return page.evaluate(() => {
    const results = [];
    const buttons = [...document.querySelectorAll("button")].filter(
      (b) => b.innerText.trim() === "Leave Public Response",
    );

    for (const btn of buttons) {
      let container = btn;
      for (let i = 0; i < 10; i++) {
        container = container.parentElement;
        if (!container) break;
        const text = container.innerText || "";
        if (text.includes("Leave Public Response") && text.length > 50 && text.length < 3000) {
          // Prefer profile link for guest name (language-agnostic)
          const profileLink = [...container.querySelectorAll("a")].find((a) => {
            const href = a.href || "";
            return href.includes("/users/profile/") || href.includes("/users/show/");
          });
          let guestName = profileLink ? profileLink.innerText.trim() : "";

          const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
          const months = "January|February|March|April|May|June|July|August|September|October|November|December";
          const dateRegex = new RegExp(`(${months})\\s+\\d{4}`);

          // Fallback: parse name from text
          if (!guestName) {
            for (let j = 0; j < lines.length; j++) {
              const m = lines[j].match(dateRegex);
              if (m) {
                const beforeDate = lines[j].substring(0, lines[j].indexOf(m[0])).trim();
                guestName = beforeDate || (j > 0 ? lines[j - 1] : "");
                break;
              }
            }
          }

          let reviewDate = "";
          let reviewText = "";
          for (let j = 0; j < lines.length; j++) {
            const m = lines[j].match(dateRegex);
            if (m) {
              reviewDate = m[0];
              const tl = [];
              for (let k = j + 1; k < lines.length; k++) {
                if (["Leave Public Response", "Private feedback", "Read more"].includes(lines[k])) break;
                tl.push(lines[k]);
              }
              reviewText = tl.join(" ").substring(0, 2000);
              break;
            }
          }
          if (guestName) {
            results.push({ guest_name: guestName, date: reviewDate, original_text: reviewText });
          }
          break;
        }
      }
    }
    return results;
  });
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
export async function postReply(page, { guest_name, draft_text }) {
  if (!guest_name) return { ok: false, error: "missing guest_name" };
  if (!draft_text) return { ok: false, error: "missing draft_text" };

  await loadReviewsPage(page);
  await clickTab(page, "Reviews about you");

  // Close any open forms
  await page.evaluate(() => {
    [...document.querySelectorAll("button")]
      .filter((b) => b.innerText.trim() === "Cancel")
      .forEach((b) => b.click());
  });
  await page.waitForTimeout(1000);

  const clicked = await page.evaluate((name) => {
    const buttons = [...document.querySelectorAll("button")].filter(
      (b) => b.innerText.trim() === "Leave Public Response",
    );
    let best = null, bestSize = Infinity;
    for (const btn of buttons) {
      let container = btn;
      for (let i = 0; i < 10; i++) {
        container = container.parentElement;
        if (!container) break;
        const len = (container.innerText || "").length;
        if (len > 1000) break;
        const link = [...container.querySelectorAll("a")].find((a) => {
          const href = a.href || "";
          return (
            (href.includes("/users/profile/") || href.includes("/users/show/")) &&
            a.innerText.trim() === name
          );
        });
        if (link && container.innerText.includes("Leave Public Response")) {
          if (len < bestSize) { best = btn; bestSize = len; }
          break;
        }
      }
    }
    if (best) { best.scrollIntoView({ block: "center" }); best.click(); return true; }
    return false;
  }, guest_name);

  if (!clicked) return { ok: false, error: `no matching "Leave Public Response" for ${guest_name}` };
  await page.waitForTimeout(2000);

  const info = await page.evaluate((name) => {
    const tas = [...document.querySelectorAll("textarea")];
    for (const ta of tas) {
      const label = ta.getAttribute("aria-label") || ta.placeholder || "";
      if (label.toLowerCase().includes(name.toLowerCase())) {
        return { id: ta.id || null, verified: true, label };
      }
    }
    if (tas.length > 0) {
      const ta = tas[0];
      return { id: ta.id || null, verified: false, label: ta.getAttribute("aria-label") || ta.placeholder || "" };
    }
    return null;
  }, guest_name);

  if (!info) return { ok: false, error: "textarea not found" };
  if (!info.verified) {
    return { ok: false, error: `aria-label mismatch — got "${info.label}", expected to include "${guest_name}"` };
  }

  await page.evaluate(({ id, text }) => {
    const ta = id ? document.getElementById(id) : document.querySelector("textarea");
    if (!ta) return;
    const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value").set;
    setter.call(ta, text);
    ta.dispatchEvent(new Event("input", { bubbles: true }));
  }, { id: info.id, text: draft_text });
  await page.waitForTimeout(1000);

  await page.evaluate(() => {
    const btn = [...document.querySelectorAll("button")].find((b) => b.innerText.trim() === "Submit");
    if (btn) btn.click();
  });
  await page.waitForTimeout(3000);

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

  for (let loop = 0; loop < 12; loop++) {
    const body = await page.evaluate(() => document.body.innerText || "");
    if (body.includes("Your review has been submitted") || body.includes("Thanks for your review")) {
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
      if (!hasRadio) return { ok: false, error: `unknown initial step; body: ${body.substring(0, 200)}` };
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

    return { ok: false, error: `unknown step ${stepNum}; body: ${body.substring(0, 200)}` };
  }
  return { ok: false, error: "exceeded max wizard iterations" };
}
