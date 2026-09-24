/**
 * ============================================================
 *  ABBAS TRADES — Backend (Google Apps Script)
 * ============================================================
 *  SECURITY NOTES:
 *  - ADMIN_KEY below gates every admin action (add/edit/remove
 *    students, viewing the full student list). Treat it exactly
 *    like a password. It already matches the key wired into the
 *    updated admin panel — if you ever need to rotate it, change
 *    it here AND in admin.html's ADMIN_KEY constant together.
 *  - Student passwords are now stored salted + hashed (SHA-256),
 *    never in plain text. Existing plaintext passwords already
 *    in your sheet are migrated automatically, silently, the
 *    next time that student logs in successfully.
 *  - The portal no longer downloads the student list to check
 *    passwords in the browser. It sends one login attempt at a
 *    time and the script verifies it here, server-side.
 *  - Notifications: enrollment form fills and confirmed Razorpay payments
 *    both trigger an email + Telegram alert. Fill in TELEGRAM_BOT_TOKEN
 *    and TELEGRAM_CHAT_ID below before this works.
 *  - RAZORPAY_WEBHOOK_SECRET gates the Razorpay webhook endpoint. Apps
 *    Script can't read custom request headers, so this can't do a real
 *    cryptographic signature check like Razorpay's docs describe — it's
 *    a shared secret passed as a URL query param instead. That's good
 *    enough here since the webhook is notification-only (it can't
 *    activate students or move money), so a spoofed call is just a
 *    fake alert at worst, not a security hole.
 *  - Referral system: lives in a separate "Referrers" sheet tab, created
 *    automatically the first time it's needed. Commission is credited
 *    automatically the moment you click "Activate" on a referred
 *    student in the admin panel (Pending -> Active transition) — that's
 *    the same manual verification step you already do, nothing new to
 *    remember. Withdrawals and course-redemption are still confirmed
 *    manually by you (see notify() alerts for both).
 * ============================================================
 */
const ADMIN_KEY = ''; // Replace with your own secret key before deploying
const NOTIFY_EMAIL = 'abbaseducates@gmail.com';
const TELEGRAM_BOT_TOKEN = '';
const TELEGRAM_CHAT_ID = '';
const RAZORPAY_WEBHOOK_SECRET = '';

// ---- Referral system config ----
const REFERRAL_COMMISSION_RATE = 0.10;       // flat 10% commission per successful referral
const REFERRAL_DISCOUNTED_PRICE = 1899;      // what a referred buyer actually pays (5% off ₹1,999)
const REFERRAL_WITHDRAW_MIN = 1500;          // balance needed to request a cash withdrawal
const REFERRAL_REDEEM_PRICE = 1500;          // balance needed for a non-student to redeem the course

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // ---- Razorpay webhook (checked first, before touching the normal body) ----
  if (e.parameter.wh === RAZORPAY_WEBHOOK_SECRET) {
    return handleRazorpayWebhook(e);
  }

  const data = JSON.parse(e.postData.contents);

  // ---- Public actions (no key required) ----

  if (data.type === 'enrollment') {
    // Form submission from website
    const refCode = String(data.referredBy || '').toUpperCase().trim();
    sheet.appendRow([new Date(), data.name, data.email, data.phone, '', 'Pending', refCode]);
    notify(
      '📝 New Enrollment Form Fill',
      'Name: ' + data.name + '\nEmail: ' + data.email + '\nPhone: ' + data.phone +
      (refCode ? '\nReferred by code: ' + refCode : '') +
      '\n\nStatus: Pending — check Razorpay to confirm payment, then activate in the admin panel.'
    );
    return jsonOut({ ok: true });
  }

  if (data.type === 'login') {
    return handleLogin(data.email, data.password);
  }

  // ---- Forgot password (students & non-student referrers) ----

  if (data.type === 'request_password_reset') {
    return handleRequestPasswordReset(data.email, data.role);
  }

  if (data.type === 'verify_reset_code') {
    return handleVerifyResetCode(data.email, data.code, data.role);
  }

  if (data.type === 'reset_password') {
    return handleResetPassword(data.email, data.code, data.newPassword, data.role);
  }

  // ---- Referral system: public actions (no key required) ----

  if (data.type === 'referrer_signup') {
    return handleReferrerSignup(data, false);
  }

  if (data.type === 'student_referrer_signup') {
    return handleReferrerSignup(data, true);
  }

  if (data.type === 'request_withdrawal') {
    return handleWithdrawalRequest(data.email);
  }

  if (data.type === 'redeem_course') {
    return handleRedeemCourse(data.email);
  }

  // ---- Admin-only actions below (key required) ----

  if (!isAuthorized(data.key)) {
    return jsonOut({ error: 'unauthorized' });
  }

  if (data.type === 'add_student') {
    // Admin adding credentials
    const rows = sheet.getDataRange().getValues();
    let found = false;
    let wasPending = false;
    let referredByCode = '';
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][2] === data.email) {
        wasPending = rows[i][5] === 'Pending';
        referredByCode = rows[i][6] || ''; // column G: ReferredBy
        sheet.getRange(i + 1, 5).setValue(newHashedPassword(data.password));
        sheet.getRange(i + 1, 6).setValue('Active');
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([new Date(), data.name, data.email, '', newHashedPassword(data.password), 'Active']);
    }
    // Credit the referrer's commission only on the Pending -> Active
    // transition, so re-editing an already-active student never double-pays.
    if (wasPending && referredByCode) {
      creditReferralCommission(referredByCode, data.email);
    }

  } else if (data.type === 'edit_student') {
    // Admin editing existing student. Password only changes if one was provided.
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][2] === data.origEmail) {
        sheet.getRange(i + 1, 2).setValue(data.name);
        sheet.getRange(i + 1, 3).setValue(data.email);
        if (data.password) {
          sheet.getRange(i + 1, 5).setValue(newHashedPassword(data.password));
        }
        break;
      }
    }

  } else if (data.type === 'remove_student') {
    // Admin removing a student permanently
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][2] === data.email) {
        sheet.deleteRow(i + 1);
        break;
      }
    }

  } else if (data.type === 'approve_referrer') {
    setReferrerStatus(data.email, 'Active');
    sendBrandedEmail(
      data.email,
      "You're approved! Start earning with Abbas Trades",
      "Good news — your affiliate application has been approved.\n\n" +
      "You can now log in at abbastrades.in/#/refer and share your referral link to start earning 10% commission on every friend who enrolls.\n\n" +
      "If you have any questions, just reply to this email."
    );

  } else if (data.type === 'reject_referrer') {
    const rejectionCount = incrementRejectionCount(data.email);
    setReferrerStatus(data.email, 'Rejected');
    const remaining = Math.max(0, 3 - rejectionCount);
    sendBrandedEmail(
      data.email,
      'Update on your Abbas Trades affiliate application',
      "Thanks for your interest in becoming an Abbas Trades affiliate.\n\n" +
      "After review, we're not able to approve your application at this time.\n\n" +
      (remaining > 0
        ? "You're welcome to apply again — you have " + remaining + " more attempt" + (remaining === 1 ? '' : 's') + " remaining.\n\n"
        : "You've now reached the maximum number of applications, so we're unable to accept further attempts.\n\n") +
      "If you have questions about this decision, feel free to reply to this email."
    );

  } else if (data.type === 'mark_withdrawal_paid') {
    const rSheet = getReferrersSheet();
    const rows = rSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][2]).toLowerCase() === String(data.email).toLowerCase()) {
        rSheet.getRange(i + 1, 9).setValue(0);   // balance -> 0
        rSheet.getRange(i + 1, 12).setValue(''); // clear WithdrawRequested flag
        break;
      }
    }

  } else if (data.type === 'mark_course_given') {
    const rSheet = getReferrersSheet();
    const rows = rSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][2]).toLowerCase() === String(data.email).toLowerCase()) {
        // Deduct only the redemption price, now that access was actually
        // delivered — any balance earned beyond it stays in their account.
        const currentBalance = Number(rows[i][8]) || 0;
        const newBalance = Math.max(0, currentBalance - REFERRAL_REDEEM_PRICE);
        rSheet.getRange(i + 1, 9).setValue(newBalance);
        rSheet.getRange(i + 1, 13).setValue(''); // clear the requested-flag so they can redeem again once balance rebuilds
        break;
      }
    }

  } else if (data.type === 'set_maintenance_status') {
    return handleSetMaintenanceStatus(data);

  } else if (data.type === 'add_update') {
    return handleAddUpdate(data);

  } else if (data.type === 'delete_update') {
    return handleDeleteUpdate(data.id);

  } else if (data.type === 'add_blog_post') {
    return handleAddBlogPost(data);

  } else if (data.type === 'update_blog_post_status') {
    return handleUpdateBlogPostStatus(data.id, data.status);

  } else if (data.type === 'delete_blog_post') {
    return handleDeleteBlogPost(data.id);
  }

  return jsonOut({ ok: true });
}

function doGet(e) {
  if (e.parameter.type === 'login') {
    return handleLogin(e.parameter.email, e.parameter.password);
  }
  if (e.parameter.type === 'get_all') {
    if (!isAuthorized(e.parameter.key)) {
      return jsonOut({ error: 'unauthorized' });
    }
    return getAllRows();
  }

  // ---- Referral system ----
  if (e.parameter.type === 'referrer_login') {
    return handleReferrerLogin(e.parameter.email, e.parameter.password);
  }
  if (e.parameter.type === 'referrer_stats') {
    // Used by the student portal's "Refer & Earn" tab — the student is
    // already authenticated via the portal login, so this just looks up
    // their referrer record by email, no separate password needed.
    return handleReferrerStats(e.parameter.email);
  }
  if (e.parameter.type === 'check_ref_code') {
    return checkRefCode(e.parameter.code);
  }
  if (e.parameter.type === 'get_referrers') {
    if (!isAuthorized(e.parameter.key)) {
      return jsonOut({ error: 'unauthorized' });
    }
    return getAllReferrers();
  }
  if (e.parameter.type === 'get_all_blog_posts') {
    if (!isAuthorized(e.parameter.key)) {
      return jsonOut({ error: 'unauthorized' });
    }
    return handleGetAllBlogPosts();
  }

  // ---- Maintenance mode & site updates (public reads) ----
  if (e.parameter.type === 'maintenance_status') {
    return handleMaintenanceStatus();
  }
  if (e.parameter.type === 'get_updates') {
    return handleGetUpdates();
  }
  if (e.parameter.type === 'get_blog_posts') {
    return handleGetBlogPosts();
  }
  if (e.parameter.type === 'get_blog_post_by_slug') {
    return handleGetBlogPostBySlug(e.parameter.slug);
  }

  return jsonOut({ error: 'not found' });
}

function isAuthorized(key) {
  return !!key && key === ADMIN_KEY;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Password hashing (salted SHA-256) ---------------- */

function makeSalt() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 16);
}

function hashPassword(password, salt) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + password);
  return raw.map(function (b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

function newHashedPassword(password) {
  const salt = makeSalt();
  return salt + ':' + hashPassword(password, salt);
}

// Accepts both new "salt:hash" values and legacy plaintext values still
// sitting in the sheet from before this upgrade.
function verifyPassword(stored, attempt) {
  if (!stored) return false;
  const s = String(stored);
  if (s.indexOf(':') !== -1) {
    const parts = s.split(':');
    return hashPassword(attempt, parts[0]) === parts[1];
  }
  return s === attempt; // legacy plaintext row, not yet migrated
}

/* ---------------- Forgot password (email OTP) ---------------- */

const RESET_CODE_TTL_MINUTES = 15;

function getPasswordResetsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('PasswordResets');
  if (!sheet) {
    sheet = ss.insertSheet('PasswordResets');
    sheet.appendRow(['Email', 'Role', 'Code', 'ExpiresAt', 'Used']);
  }
  return sheet;
}

function generateResetCode() {
  return String(Math.floor(100000 + Math.random() * 900000)); // 6-digit code
}

// role: 'student' or 'referrer'
function emailExistsForRole(email, role) {
  const target = String(email || '').toLowerCase().trim();
  if (!target) return false;
  if (role === 'student') {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][5] === 'Active' && rows[i][2] && String(rows[i][2]).toLowerCase() === target) return true;
    }
    return false;
  }
  if (role === 'referrer') {
    const found = findReferrerRow(target);
    // Student-referrers don't have a password stored on the Referrers sheet
    // (they use their student login instead) — they should reset via the
    // 'student' role, not this one.
    if (!found) return false;
    const isStudent = !!found.row[6] || found.row[6] === 'TRUE';
    return !isStudent;
  }
  return false;
}

// Sends an email that appears to come FROM abbaseducates@gmail.com (the
// business address) rather than whichever personal Google account the
// Apps Script happens to be deployed under.
//
// IMPORTANT: for the "from" address below to actually take effect (rather
// than silently falling back to the script owner's own address), the
// Google account that owns/runs this Apps Script must have
// abbaseducates@gmail.com added and VERIFIED as a "Send mail as" alias:
// Gmail → Settings → Accounts → "Send mail as" → Add another email address.
function sendBrandedEmail(to, subject, body) {
  try {
    GmailApp.sendEmail(to, subject, body, {
      name: 'Abbas Trades',
      from: NOTIFY_EMAIL
    });
  } catch (err) {
    console.log('sendBrandedEmail failed, falling back to MailApp: ' + err);
    try { MailApp.sendEmail(to, subject, body); } catch (err2) { console.log('Fallback email also failed: ' + err2); }
  }
}

function handleRequestPasswordReset(email, role) {
  const target = String(email || '').toLowerCase().trim();
  // Always respond success (don't reveal whether an account exists), but
  // only actually send a code if it does.
  if (target && (role === 'student' || role === 'referrer') && emailExistsForRole(target, role)) {
    const code = generateResetCode();
    const expiresAt = new Date(Date.now() + RESET_CODE_TTL_MINUTES * 60 * 1000);
    const sheet = getPasswordResetsSheet();
    const rows = sheet.getDataRange().getValues();
    let updated = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toLowerCase() === target && rows[i][1] === role) {
        sheet.getRange(i + 1, 3).setValue(code);
        sheet.getRange(i + 1, 4).setValue(expiresAt);
        sheet.getRange(i + 1, 5).setValue(false);
        updated = true;
        break;
      }
    }
    if (!updated) sheet.appendRow([target, role, code, expiresAt, false]);

    sendBrandedEmail(
      target,
      'Your Abbas Trades password reset code',
      'Your verification code is: ' + code +
      '\n\nThis code expires in ' + RESET_CODE_TTL_MINUTES + ' minutes.' +
      '\n\nIf you did not request this, you can safely ignore this email.'
    );
  }
  return jsonOut({ success: true });
}

function findValidResetCode(email, code, role) {
  const target = String(email || '').toLowerCase().trim();
  const sheet = getPasswordResetsSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === target && rows[i][1] === role) {
      const matches = String(rows[i][2]) === String(code);
      const notExpired = rows[i][3] && new Date(rows[i][3]).getTime() > Date.now();
      const notUsed = !rows[i][4];
      if (matches && notExpired && notUsed) return { sheet: sheet, rowIndex: i + 1 };
      return null;
    }
  }
  return null;
}

function handleVerifyResetCode(email, code, role) {
  const found = findValidResetCode(email, code, role);
  return jsonOut({ valid: !!found });
}

function handleResetPassword(email, code, newPassword, role) {
  if (!newPassword || String(newPassword).length < 6) {
    return jsonOut({ success: false, error: 'Password must be at least 6 characters.' });
  }
  const found = findValidResetCode(email, code, role);
  if (!found) {
    return jsonOut({ success: false, error: 'That code is invalid or has expired. Please request a new one.' });
  }

  const target = String(email || '').toLowerCase().trim();
  const hashed = newHashedPassword(newPassword);

  if (role === 'student') {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rows = sheet.getDataRange().getValues();
    let done = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][5] === 'Active' && rows[i][2] && String(rows[i][2]).toLowerCase() === target) {
        sheet.getRange(i + 1, 5).setValue(hashed);
        done = true;
        break;
      }
    }
    if (!done) return jsonOut({ success: false, error: 'Account not found.' });
  } else if (role === 'referrer') {
    const rsheet = getReferrersSheet();
    const rows = rsheet.getDataRange().getValues();
    let done = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][2]).toLowerCase() === target) {
        rsheet.getRange(i + 1, 5).setValue(hashed);
        done = true;
        break;
      }
    }
    if (!done) return jsonOut({ success: false, error: 'Account not found.' });
  } else {
    return jsonOut({ success: false, error: 'Invalid request.' });
  }

  found.sheet.getRange(found.rowIndex, 5).setValue(true); // mark code used
  return jsonOut({ success: true });
}

/* ---------------- Login (server-side password check) ---------------- */

function handleLogin(email, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  const target = String(email || '').toLowerCase().trim();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][5] === 'Active' && rows[i][2] && String(rows[i][2]).toLowerCase() === target) {
      const stored = rows[i][4];
      if (verifyPassword(stored, password)) {
        // Silently upgrade legacy plaintext passwords to hashed form.
        if (String(stored).indexOf(':') === -1) {
          sheet.getRange(i + 1, 5).setValue(newHashedPassword(password));
        }
        return jsonOut({ success: true, name: rows[i][1] });
      }
      break;
    }
  }
  return jsonOut({ success: false });
}

/* ---------------- Admin: full student list ---------------- */

function getAllRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  const all = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2]) {
      all.push({
        timestamp: rows[i][0],
        name: rows[i][1],
        email: rows[i][2],
        phone: rows[i][3],
        hasPassword: !!rows[i][4],
        status: rows[i][5] || 'Pending',
        referredBy: rows[i][6] || ''
      });
    }
  }
  return jsonOut(all);
}

/* ---------------- Notifications (email + Telegram) ---------------- */

function notify(subject, body) {
  try {
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
    console.log('Email sent OK to ' + NOTIFY_EMAIL);
  } catch (err) {
    console.log('Email send failed: ' + err);
  }
  sendTelegramMessage('*' + subject + '*\n' + body);
}

// Run this one manually (select "testEmail" in the function dropdown, then
// click Run) to check email specifically. Check the Execution log after —
// it'll show "Email sent OK" or the exact error if something's wrong.
function testEmail() {
  notify('✅ Test email from Apps Script', 'If you see this in your inbox, email alerts are working.');
}

function sendTelegramMessage(text) {
  if (!TELEGRAM_BOT_TOKEN || TELEGRAM_BOT_TOKEN.indexOf('PASTE_') === 0) return;
  try {
    const res = UrlFetchApp.fetch('' + TELEGRAM_BOT_TOKEN + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: TELEGRAM_CHAT_ID,
        text: text,
        parse_mode: 'Markdown'
      }),
      muteHttpExceptions: true
    });
    console.log('Telegram response code: ' + res.getResponseCode());
    console.log('Telegram response body: ' + res.getContentText());
  } catch (err) {
    console.log('Telegram send threw an error: ' + err);
  }
}

// Run this one manually (select "testTelegram" in the function dropdown,
// then click Run) to see exactly what Telegram says back. Check the
// Execution log after running it — the response code/body will tell us
// precisely what's wrong (bad token, wrong chat id, etc.) instead of
// silently failing.
function testTelegram() {
  sendTelegramMessage('✅ Test message from Apps Script — if you see this, Telegram alerts are working.');
}

/* ---------------- Razorpay webhook (notification-only) ---------------- */

function handleRazorpayWebhook(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const event = data.event || 'unknown';
    const payment = (data.payload && data.payload.payment && data.payload.payment.entity) || {};
    const amountRupees = payment.amount ? (payment.amount / 100).toFixed(2) : '?';
    const email = payment.email || '—';
    const contact = payment.contact || '—';

    notify(
      '💰 Razorpay Payment: ' + event,
      'Amount: ₹' + amountRupees + '\nEmail: ' + email + '\nContact: ' + contact +
      '\n\nCheck your Pending enrollments and activate the matching student.'
    );
  } catch (err) {
    // Still return 200 below even if parsing fails — Razorpay disables
    // webhooks that don't get a prompt 200 response.
  }
  return jsonOut({ ok: true });
}

/* ============================================================
 *  REFERRAL SYSTEM
 *  Sheet columns (auto-created, tab name "Referrers"):
 *  A Timestamp | B Name | C Email | D Phone | E Password (hashed,
 *  blank for students) | F Code | G IsStudent | H Status
 *  (Pending/Active/Rejected) | I Balance | J TotalEarned |
 *  K ReferralCount | L WithdrawRequested (timestamp or blank) |
 *  M Redeemed (timestamp or blank)
 * ============================================================ */

function getReferrersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Referrers');
  if (!sheet) {
    sheet = ss.insertSheet('Referrers');
    sheet.appendRow(['Timestamp', 'Name', 'Email', 'Phone', 'Password', 'Code', 'IsStudent', 'Status', 'Balance', 'TotalEarned', 'ReferralCount', 'WithdrawRequested', 'Redeemed', 'RejectionCount']);
  }
  return sheet;
}

function getReferralLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('ReferralLog');
  if (!sheet) {
    sheet = ss.insertSheet('ReferralLog');
    sheet.appendRow(['Timestamp', 'ReferrerCode', 'ReferrerEmail', 'BuyerEmail', 'Commission', 'Tier', 'RunningTotal']);
  }
  return sheet;
}

function logReferralEvent(code, referrerEmail, buyerEmail, commission, tierName, runningTotal) {
  getReferralLogSheet().appendRow([new Date(), code, referrerEmail, buyerEmail, commission, tierName, runningTotal]);
}

// Masks a buyer's email for privacy when shown on the referrer's own
// dashboard, e.g. "john.doe@gmail.com" -> "jo***@gmail.com".
function maskEmail(email) {
  const str = String(email || '');
  const at = str.indexOf('@');
  if (at <= 0) return '***';
  const name = str.slice(0, at);
  const domain = str.slice(at);
  const visible = name.slice(0, Math.min(2, name.length));
  return visible + '***' + domain;
}

function getReferralHistory(referrerEmail) {
  const sheet = getReferralLogSheet();
  const rows = sheet.getDataRange().getValues();
  const target = String(referrerEmail || '').toLowerCase().trim();
  const history = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).toLowerCase() === target) {
      history.push({
        date: rows[i][0] instanceof Date ? rows[i][0].toISOString() : String(rows[i][0]),
        buyer: maskEmail(rows[i][3]),
        commission: Number(rows[i][4]) || 0,
        tier: rows[i][5] || '',
        runningTotal: Number(rows[i][6]) || 0
      });
    }
  }
  return history; // chronological order (oldest first), matches sheet append order
}

function findReferrerRow(email) {
  const sheet = getReferrersSheet();
  const rows = sheet.getDataRange().getValues();
  const target = String(email || '').toLowerCase().trim();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).toLowerCase() === target) {
      return { sheet: sheet, rowIndex: i + 1, row: rows[i] };
    }
  }
  return null;
}

function handleReferrerSignup(data, isStudent) {
  const sheet = getReferrersSheet();
  const rows = sheet.getDataRange().getValues();
  const email = String(data.email || '').toLowerCase().trim();
  const code = String(data.code || '').toUpperCase().trim();

  if (!email || !code || !data.name) {
    return jsonOut({ success: false, error: 'Missing required fields.' });
  }
  if (!/^[A-Z0-9]{3,15}$/.test(code)) {
    return jsonOut({ success: false, error: 'Code must be 3-15 letters/numbers only, no spaces or symbols.' });
  }
  if (!isStudent && (!data.password || data.password.length < 6)) {
    return jsonOut({ success: false, error: 'Password must be at least 6 characters.' });
  }

  let existingRowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).toLowerCase() === email) {
      existingRowIndex = i;
      continue; // keep scanning so a code-collision check below still runs
    }
    if (String(rows[i][5]).toUpperCase() === code) {
      return jsonOut({ success: false, error: 'That code is already taken — try a different one.' });
    }
  }

  if (existingRowIndex !== -1) {
    const existingStatus = rows[existingRowIndex][7];
    const rejectionCount = getRejectionCount(rows[existingRowIndex]);

    if (existingStatus !== 'Rejected') {
      return jsonOut({ success: false, error: 'You already have a referrer account — try logging in instead.' });
    }
    if (rejectionCount >= MAX_REFERRER_REJECTIONS) {
      return jsonOut({ success: false, error: "You've reached the maximum number of applications and can no longer reapply. Contact " + NOTIFY_EMAIL + ' if you think this is a mistake.' });
    }

    // Rejected, but still has attempts left — update their existing row
    // with the new details and put them back in the Pending queue.
    const passwordHash = isStudent ? '' : newHashedPassword(data.password);
    const sheetRow = existingRowIndex + 1;
    sheet.getRange(sheetRow, 2).setValue(data.name);
    sheet.getRange(sheetRow, 4).setValue(data.phone || '');
    sheet.getRange(sheetRow, 5).setValue(passwordHash);
    sheet.getRange(sheetRow, 6).setValue(code);
    sheet.getRange(sheetRow, 8).setValue('Pending');

    if (!isStudent) {
      notify(
        '🤝 Affiliate Reapplication',
        'Name: ' + data.name + '\nEmail: ' + email + '\nPhone: ' + (data.phone || '') + '\nRequested code: ' + code +
        '\nPrevious rejections: ' + rejectionCount + ' (max ' + MAX_REFERRER_REJECTIONS + ')' +
        '\n\nApprove or reject in the admin panel\'s Affiliates section.'
      );
    }
    return jsonOut({ success: true, status: 'Pending', code: code });
  }

  const passwordHash = isStudent ? '' : newHashedPassword(data.password);
  const status = isStudent ? 'Active' : 'Pending';

  sheet.appendRow([new Date(), data.name, email, data.phone || '', passwordHash, code, isStudent, status, 0, 0, 0, '', '', 0]);

  if (!isStudent) {
    notify(
      '🤝 New Affiliate Signup',
      'Name: ' + data.name + '\nEmail: ' + email + '\nPhone: ' + (data.phone || '') + '\nRequested code: ' + code +
      '\n\nApprove or reject in the admin panel\'s Affiliates section.'
    );
  }

  return jsonOut({ success: true, status: status, code: code });
}

function buildReferrerPayload(row, email) {
  const referralCount = Number(row[10]) || 0;
  return {
    name: row[1],
    code: row[5],
    status: row[7],
    isStudent: !!row[6] || row[6] === 'TRUE',
    balance: Number(row[8]) || 0,
    totalEarned: Number(row[9]) || 0,
    referralCount: referralCount,
    withdrawRequested: !!row[11],
    redeemed: !!row[12],
    commissionRate: REFERRAL_COMMISSION_RATE,
    history: getReferralHistory(email)
  };
}

function handleReferrerLogin(email, password) {
  const found = findReferrerRow(email);
  if (!found) return jsonOut({ success: false });

  const row = found.row;
  if (row[7] === 'Pending') return jsonOut({ success: false, pending: true });
  if (row[7] === 'Rejected') {
    const rejectionCount = getRejectionCount(row);
    const remaining = Math.max(0, MAX_REFERRER_REJECTIONS - rejectionCount);
    return jsonOut({ success: false, rejected: true, remainingChances: remaining, canReapply: remaining > 0 });
  }

  const isStudent = !!row[6] || row[6] === 'TRUE';

  if (isStudent) {
    // Students don't have a password stored in the Referrers sheet (it's
    // intentionally blank — they use their normal student portal login
    // instead). Verify against the main student sheet's password here so
    // the same credentials work on the public /#/refer page too.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rows = sheet.getDataRange().getValues();
    const target = String(email || '').toLowerCase().trim();
    let matched = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][5] === 'Active' && rows[i][2] && String(rows[i][2]).toLowerCase() === target) {
        if (verifyPassword(rows[i][4], password)) {
          matched = true;
          if (String(rows[i][4]).indexOf(':') === -1) {
            sheet.getRange(i + 1, 5).setValue(newHashedPassword(password));
          }
        }
        break;
      }
    }
    if (!matched) return jsonOut({ success: false });
  } else {
    const stored = row[4];
    if (!verifyPassword(stored, password)) return jsonOut({ success: false });
    if (String(stored).indexOf(':') === -1 && stored) {
      found.sheet.getRange(found.rowIndex, 5).setValue(newHashedPassword(password));
    }
  }

  return jsonOut(Object.assign({ success: true }, buildReferrerPayload(row, email)));
}

// Used by the student portal — the student is already authenticated via
// their normal portal login, so this looks them up by email only.
function handleReferrerStats(email) {
  const found = findReferrerRow(email);
  if (!found) return jsonOut({ exists: false });
  return jsonOut(Object.assign({ exists: true }, buildReferrerPayload(found.row, email)));
}

function checkRefCode(code) {
  const sheet = getReferrersSheet();
  const rows = sheet.getDataRange().getValues();
  const target = String(code || '').toUpperCase().trim();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][5]).toUpperCase() === target && rows[i][7] === 'Active') {
      return jsonOut({ valid: true });
    }
  }
  return jsonOut({ valid: false });
}

const MAX_REFERRER_REJECTIONS = 3; // after this many rejections, reapplying is permanently blocked

function setReferrerStatus(email, status) {
  const found = findReferrerRow(email);
  if (found) found.sheet.getRange(found.rowIndex, 8).setValue(status);
}

// Column N (index 13) holds how many times this email has been rejected.
// Returns the new count after incrementing.
function incrementRejectionCount(email) {
  const found = findReferrerRow(email);
  if (!found) return 0;
  const current = Number(found.row[13]) || 0;
  const next = current + 1;
  found.sheet.getRange(found.rowIndex, 14).setValue(next);
  return next;
}

function getRejectionCount(row) {
  return Number(row[13]) || 0;
}

function creditReferralCommission(code, buyerEmail) {
  const sheet = getReferrersSheet();
  const rows = sheet.getDataRange().getValues();
  const target = String(code).toUpperCase().trim();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][5]).toUpperCase() === target && rows[i][7] === 'Active') {
      const newCount = (Number(rows[i][10]) || 0) + 1;
      const commission = Math.round(REFERRAL_DISCOUNTED_PRICE * REFERRAL_COMMISSION_RATE);
      const newBalance = (Number(rows[i][8]) || 0) + commission;
      const newTotal = (Number(rows[i][9]) || 0) + commission;
      sheet.getRange(i + 1, 9).setValue(newBalance);
      sheet.getRange(i + 1, 10).setValue(newTotal);
      sheet.getRange(i + 1, 11).setValue(newCount);
      logReferralEvent(target, rows[i][2], buyerEmail, commission, 'Standard', newTotal);
      notify(
        '💸 Referral Commission Credited',
        'Referrer code: ' + code + '\nNew referral: ' + buyerEmail +
        '\nCommission: ₹' + commission + '\nTheir new balance: ₹' + newBalance + '\nLifetime referrals: ' + newCount
      );
      break;
    }
  }
}

function handleWithdrawalRequest(email) {
  const found = findReferrerRow(email);
  if (!found) return jsonOut({ success: false, error: 'Account not found.' });
  const balance = Number(found.row[8]) || 0;
  if (balance < REFERRAL_WITHDRAW_MIN) {
    return jsonOut({ success: false, error: 'Balance must reach ₹' + REFERRAL_WITHDRAW_MIN + ' to withdraw.' });
  }
  found.sheet.getRange(found.rowIndex, 12).setValue(new Date());
  notify(
    '💰 Withdrawal Requested',
    'Name: ' + found.row[1] + '\nEmail: ' + email + '\nCode: ' + found.row[5] + '\nAmount: ₹' + balance +
    '\n\nPay manually via UPI, then mark it as paid in the admin panel\'s Affiliates section.'
  );
  return jsonOut({ success: true });
}

function handleRedeemCourse(email) {
  const found = findReferrerRow(email);
  if (!found) return jsonOut({ success: false, error: 'Account not found.' });
  if (found.row[6] === true || found.row[6] === 'TRUE') {
    return jsonOut({ success: false, error: 'This option is only for referrers who are not already students.' });
  }
  if (found.row[12]) {
    return jsonOut({ success: false, error: 'A redemption request is already on file for this account.' });
  }
  const balance = Number(found.row[8]) || 0;
  if (balance < REFERRAL_REDEEM_PRICE) {
    return jsonOut({ success: false, error: 'Balance must reach ₹' + REFERRAL_REDEEM_PRICE + ' to redeem.' });
  }
  // Mark the redemption request but leave the balance untouched — it only
  // gets zeroed once the admin confirms the course was actually delivered
  // (see 'mark_course_given'), so nothing is "spent" before it's fulfilled.
  found.sheet.getRange(found.rowIndex, 13).setValue(new Date());
  notify(
    '🎓 Course Redemption Requested',
    'Name: ' + found.row[1] + '\nEmail: ' + email + '\nCode: ' + found.row[5] +
    '\n\nThey want to redeem ₹' + REFERRAL_REDEEM_PRICE + ' of earned commission for course access instead of cash. ' +
    'Their balance stays as-is until you activate their student portal access and mark it given in the admin panel.'
  );
  return jsonOut({ success: true });
}

/* ---------------- Admin: full referrer list ---------------- */

function getAllReferrers() {
  const sheet = getReferrersSheet();
  const rows = sheet.getDataRange().getValues();
  const all = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2]) {
      all.push({
        timestamp: rows[i][0],
        name: rows[i][1],
        email: rows[i][2],
        phone: rows[i][3],
        code: rows[i][5],
        isStudent: !!rows[i][6],
        status: rows[i][7] || 'Pending',
        balance: Number(rows[i][8]) || 0,
        totalEarned: Number(rows[i][9]) || 0,
        referralCount: Number(rows[i][10]) || 0,
        withdrawRequested: !!rows[i][11],
        redeemed: !!rows[i][12]
      });
    }
  }
  return jsonOut(all);
}

/* ---------------- Maintenance mode ---------------- */

function getConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.appendRow(['Key', 'Value']);
    sheet.appendRow(['MaintenanceMode', 'FALSE']);
    sheet.appendRow(['MaintenanceMessage', "We're making a few improvements. Back shortly — thanks for your patience!"]);
    sheet.appendRow(['MaintenanceETA', '']);
  }
  return sheet;
}

function getConfigValue(key) {
  const sheet = getConfigSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) return rows[i][1];
  }
  return '';
}

function setConfigValue(key, value) {
  const sheet = getConfigSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function handleMaintenanceStatus() {
  const enabled = String(getConfigValue('MaintenanceMode')).toUpperCase() === 'TRUE';
  return jsonOut({
    enabled: enabled,
    message: getConfigValue('MaintenanceMessage') || "We're making a few improvements. Back shortly!",
    eta: getConfigValue('MaintenanceETA') || ''
  });
}

function handleSetMaintenanceStatus(data) {
  setConfigValue('MaintenanceMode', data.enabled ? 'TRUE' : 'FALSE');
  if (data.message != null) setConfigValue('MaintenanceMessage', data.message);
  if (data.eta != null) setConfigValue('MaintenanceETA', data.eta);
  return jsonOut({ success: true });
}

/* ---------------- Site updates feed ---------------- */

function getUpdatesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Updates');
  if (!sheet) {
    sheet = ss.insertSheet('Updates');
    sheet.appendRow(['Id', 'Timestamp', 'Title', 'Body']);
  }
  return sheet;
}

function handleGetUpdates() {
  const sheet = getUpdatesSheet();
  const rows = sheet.getDataRange().getValues();
  const updates = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) {
      updates.push({
        id: rows[i][0],
        timestamp: rows[i][1] instanceof Date ? rows[i][1].toISOString() : String(rows[i][1]),
        title: rows[i][2],
        body: rows[i][3]
      });
    }
  }
  updates.reverse(); // newest first
  return jsonOut(updates.slice(0, 10));
}

function handleAddUpdate(data) {
  if (!data.title) return jsonOut({ success: false, error: 'Title required.' });
  const sheet = getUpdatesSheet();
  const id = 'u' + Date.now();
  sheet.appendRow([id, new Date(), data.title, data.body || '']);
  return jsonOut({ success: true, id: id });
}

function handleDeleteUpdate(id) {
  const sheet = getUpdatesSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      sheet.deleteRow(i + 1);
      return jsonOut({ success: true });
    }
  }
  return jsonOut({ success: false, error: 'Not found.' });
}

/* ---------------- Blog posts (draft authoring + published listing) ---------------- */

function getBlogPostsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BlogPosts');
  if (!sheet) {
    sheet = ss.insertSheet('BlogPosts');
    sheet.appendRow(['Id', 'Timestamp', 'Title', 'Category', 'Excerpt', 'Slug', 'RawContent', 'Status']);
  }
  return sheet;
}

// Public: used by blog.html to list published posts.
function handleGetBlogPosts() {
  const sheet = getBlogPostsSheet();
  const rows = sheet.getDataRange().getValues();
  const posts = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] && rows[i][7] === 'Published') {
      posts.push({
        id: rows[i][0],
        timestamp: rows[i][1] instanceof Date ? rows[i][1].toISOString() : String(rows[i][1]),
        title: rows[i][2],
        category: rows[i][3],
        excerpt: rows[i][4],
        slug: rows[i][5]
      });
    }
  }
  posts.reverse(); // newest first
  return jsonOut(posts);
}

// Public: used by blog-post.html to fetch one published article by its slug.
function handleGetBlogPostBySlug(slug) {
  const sheet = getBlogPostsSheet();
  const rows = sheet.getDataRange().getValues();
  const target = String(slug || '').toLowerCase().trim();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] && rows[i][7] === 'Published' && String(rows[i][5]).toLowerCase() === target) {
      return jsonOut({
        found: true,
        title: rows[i][2],
        category: rows[i][3],
        excerpt: rows[i][4],
        slug: rows[i][5],
        rawContent: rows[i][6],
        timestamp: rows[i][1] instanceof Date ? rows[i][1].toISOString() : String(rows[i][1])
      });
    }
  }
  return jsonOut({ found: false });
}

// Admin-only: lists everything (drafts + published) for the admin panel.
function handleGetAllBlogPosts() {
  const sheet = getBlogPostsSheet();
  const rows = sheet.getDataRange().getValues();
  const posts = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) {
      posts.push({
        id: rows[i][0],
        timestamp: rows[i][1] instanceof Date ? rows[i][1].toISOString() : String(rows[i][1]),
        title: rows[i][2],
        category: rows[i][3],
        excerpt: rows[i][4],
        slug: rows[i][5],
        rawContent: rows[i][6],
        status: rows[i][7]
      });
    }
  }
  posts.reverse();
  return jsonOut(posts);
}

function slugifyTitle(title) {
  return String(title || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/(^-|-$)/g, '')
    .slice(0, 60);
}

function handleAddBlogPost(data) {
  if (!data.title || !data.rawContent) {
    return jsonOut({ success: false, error: 'Title and content are required.' });
  }
  const sheet = getBlogPostsSheet();
  const slug = 'blog-' + slugifyTitle(data.title);
  const status = data.status === 'Published' ? 'Published' : 'Draft';

  // If an existing post id was passed in, update that row in place instead
  // of creating a duplicate — this is how editing a saved article works.
  if (data.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.getRange(i + 1, 3).setValue(data.title);
        sheet.getRange(i + 1, 4).setValue(data.category || '');
        sheet.getRange(i + 1, 5).setValue(data.excerpt || '');
        sheet.getRange(i + 1, 6).setValue(slug);
        sheet.getRange(i + 1, 7).setValue(data.rawContent);
        sheet.getRange(i + 1, 8).setValue(status);
        return jsonOut({ success: true, id: data.id, slug: slug });
      }
    }
    // id didn't match anything (e.g. was deleted) — fall through and create new
  }

  const id = 'p' + Date.now();
  sheet.appendRow([
    id, new Date(), data.title, data.category || '', data.excerpt || '',
    slug, data.rawContent, status
  ]);
  return jsonOut({ success: true, id: id, slug: slug });
}

function handleUpdateBlogPostStatus(id, status) {
  const sheet = getBlogPostsSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      sheet.getRange(i + 1, 8).setValue(status === 'Published' ? 'Published' : 'Draft');
      return jsonOut({ success: true });
    }
  }
  return jsonOut({ success: false, error: 'Not found.' });
}

function handleDeleteBlogPost(id) {
  const sheet = getBlogPostsSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      sheet.deleteRow(i + 1);
      return jsonOut({ success: true });
    }
  }
  return jsonOut({ success: false, error: 'Not found.' });
}
