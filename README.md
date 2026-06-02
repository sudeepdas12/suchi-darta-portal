# Suchi Darta Portal

This repository contains the Suchi Darta (registration) static front-end and a serverless API used to obtain Backblaze B2 upload URLs.

Contents
- `index.html` — single-page registration form (Nepali UI), upload UI, and client-side upload pipeline.
- `api/get-upload-url.js` — serverless function (Vercel) that authorizes with Backblaze and returns a one-time upload URL/token.

Quick start (local)
1. Install dependencies (if you need to run server code):

```bash
npm install
```

2. Deploy to Vercel (recommended):

```bash
vercel --prod --cwd . -e B2_KEY_ID=<id> -e B2_APP_KEY=<app-key> -e B2_BUCKET_ID=<bucket-id>
```

Notes
- Rotate Backblaze keys immediately if they were exposed. Store keys in Vercel environment variables rather than passing them inline in terminals when possible.
- If you'd like, I can push this repo to your GitHub account and create a short release note.

License
MIT
