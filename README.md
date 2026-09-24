# Abbas Trades — Forex Trading Education Platform

A full-stack trading education platform using Google Apps Script + Google Sheets for the backend and static HTML/CSS/JavaScript for the frontend.

## Project structure

```text
index.html             # Router / shell
Code.gs                # Google Apps Script backend
pages/
  home.html
  portal.html
  admin.html
  refer.html
  thankyou.html
css/
  router.css
  home.css
  portal.css
  admin.css
  refer.css
  thankyou.css
js/
  router.js
  home.js
  portal.js
  admin.js
  refer.js
  thankyou.js
```

## Security

This repository is intended to contain sanitized source only. Before publishing publicly, verify that no real credentials, tokens, private customer data, payment secrets, or private deployment URLs are present.

The following values are placeholders and must be configured privately for your own deployment:

- `ADMIN_KEY`
- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID`
- `RAZORPAY_WEBHOOK_SECRET`
- Google Apps Script deployment URL

Never commit real secrets to GitHub. If a real secret was previously committed, rotate (replace) it rather than merely deleting the file.

## Backend

`Code.gs` is the Google Apps Script backend. It uses Google Sheets as its database and exposes REST-style endpoints for authentication, admin actions, referrals, password resets, and other application functions.

## Frontend

The original monolithic `index-2.html` contained multiple complete HTML pages encoded as Base64 inside JavaScript. This version separates those pages into normal HTML files and separates their CSS and JavaScript into dedicated files for easier maintenance and GitHub review.


## Public archive sanitization

This archive intentionally contains no live passwords, API tokens, deployment URLs, payment URLs, video URLs, or external website links. Some features therefore require configuration before use.
