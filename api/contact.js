// Vercel Serverless Function — Contact Form Handler
// Stores inquiries to Messages/messages.xlsx (via GitHub REST API in production or local disk in dev)
// and sends email notifications via nodemailer.
//
// Environment variables required in Vercel project settings:
//   GITHUB_TOKEN  — GitHub personal access token (with repo scope)
//   GITHUB_OWNER  — GitHub username (e.g. aibinjmca2527-dotcom)
//   GITHUB_REPO   — Repository name (e.g. my-portfolio)
//   EMAIL_USER    — Gmail/SMTP address used to send notifications
//   EMAIL_PASS    — Gmail App Password (not standard account password)
//   NOTIFY_EMAIL  — Recipient email to receive notifications

const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const XLSX = require('xlsx');

const GITHUB_API = 'https://api.github.com';
const FILE_PATH  = 'Messages/messages.xlsx';

// ─── Sanitization helpers ──────────────────────────────────────────────────────

function sanitizeString(str) {
  if (typeof str !== 'string') return '';
  // Strip control characters and basic HTML tags
  let cleaned = str.replace(/<[^>]*>?/gm, '').trim();
  // Prevent Excel formula injection: prefix leading formula triggers (=, +, -, @, \t, \r) with a single quote
  if (/^[=+\-@\t\r]/.test(cleaned)) {
    cleaned = "'" + cleaned;
  }
  return cleaned;
}

// ─── GitHub API helpers ────────────────────────────────────────────────────────

async function ghGet(token, owner, repo, filePath) {
  const res = await fetch(
    `${GITHUB_API}/repos/${owner}/${repo}/contents/${filePath}`,
    { headers: { Authorization: `token ${token}`, Accept: 'application/vnd.github.v3+json' } }
  );
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`GitHub GET ${res.status}: ${await res.text()}`);
  return res.json();
}

async function ghPut(token, owner, repo, filePath, content64, sha, message) {
  const body = { message, content: content64 };
  if (sha) body.sha = sha;
  const res = await fetch(
    `${GITHUB_API}/repos/${owner}/${repo}/contents/${filePath}`,
    {
      method: 'PUT',
      headers: {
        Authorization: `token ${token}`,
        Accept: 'application/vnd.github.v3+json',
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
    }
  );
  if (!res.ok) throw new Error(`GitHub PUT ${res.status}: ${await res.text()}`);
  return res.json();
}

// ─── Excel helpers ────────────────────────────────────────────────────────────

function buildWorkbook(rows) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['Name', 'Email', 'Message', 'Date', 'Time'],
    ...rows,
  ]);
  // Column widths
  ws['!cols'] = [{ wch: 25 }, { wch: 35 }, { wch: 80 }, { wch: 15 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Messages');
  return wb;
}

function readWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = wb.SheetNames[0] || 'Messages';
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];
  const all = XLSX.utils.sheet_to_json(ws, { header: 1 });
  // Return rows excluding the header row
  return all.slice(1);
}

// ─── Local file fallback helper (for local testing without GitHub token) ──────

function saveLocalExcel(row) {
  const localDir = path.join(process.cwd(), 'Messages');
  const localFile = path.join(localDir, 'messages.xlsx');
  if (!fs.existsSync(localDir)) {
    fs.mkdirSync(localDir, { recursive: true });
  }
  let existingRows = [];
  if (fs.existsSync(localFile)) {
    try {
      const buf = fs.readFileSync(localFile);
      existingRows = readWorkbook(buf);
    } catch (e) {
      console.warn('Could not read existing local workbook:', e.message);
    }
  }
  existingRows.push(row);
  const wb = buildWorkbook(existingRows);
  XLSX.writeFile(wb, localFile);
}

// ─── Email helper ─────────────────────────────────────────────────────────────

async function sendNotification({ emailUser, emailPass, notifyEmail, name, email, message, datetime }) {
  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: { user: emailUser, pass: emailPass },
  });
  await transporter.sendMail({
    from: `"Portfolio Contact" <${emailUser}>`,
    to: notifyEmail,
    replyTo: email,
    subject: `New Portfolio Inquiry from ${name}`,
    html: `
      <div style="font-family:Inter,system-ui,sans-serif;max-width:560px;margin:0 auto;background:#0d1525;color:#e2e8f0;border-radius:12px;overflow:hidden;border:1px solid #1e293b;">
        <div style="background:linear-gradient(135deg,#6366f1,#38bdf8);padding:24px 32px;">
          <h2 style="margin:0;color:#fff;font-size:1.35rem;letter-spacing:-0.5px;">New Portfolio Inquiry</h2>
        </div>
        <div style="padding:28px 32px;">
          <table style="width:100%;border-collapse:collapse;margin-bottom:20px;">
            <tr><td style="padding:8px 0;color:#94a3b8;width:90px;font-size:0.85rem;">Name</td><td style="padding:8px 0;color:#f1f5f9;font-weight:600;font-size:0.95rem;">${name}</td></tr>
            <tr><td style="padding:8px 0;color:#94a3b8;font-size:0.85rem;">Email</td><td style="padding:8px 0;"><a href="mailto:${email}" style="color:#38bdf8;text-decoration:none;">${email}</a></td></tr>
            <tr><td style="padding:8px 0;color:#94a3b8;font-size:0.85rem;">Date/Time</td><td style="padding:8px 0;color:#f1f5f9;font-size:0.9rem;">${datetime}</td></tr>
          </table>
          <div style="margin-top:16px;padding:16px;background:rgba(255,255,255,.05);border-left:3px solid #6366f1;border-radius:6px;">
            <p style="margin:0 0 8px;color:#94a3b8;font-size:0.75rem;font-weight:600;letter-spacing:1px;text-transform:uppercase;">Message</p>
            <p style="margin:0;color:#f1f5f9;line-height:1.7;font-size:0.92rem;white-space:pre-wrap;">${message}</p>
          </div>
        </div>
      </div>
    `,
  });
}

// ─── Validation ───────────────────────────────────────────────────────────────

function validate({ name, email, message }) {
  if (!name || typeof name !== 'string' || name.trim().length < 2 || name.trim().length > 100)
    return 'Name must be 2–100 characters.';
  if (!email || typeof email !== 'string' || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim()) || email.length > 200)
    return 'Please enter a valid email address.';
  if (!message || typeof message !== 'string' || message.trim().length < 10 || message.trim().length > 2000)
    return 'Message must be 10–2000 characters.';
  return null;
}

// ─── Main handler ─────────────────────────────────────────────────────────────

module.exports = async function handler(req, res) {
  // CORS preflight
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  // Env variables
  const { GITHUB_TOKEN, GITHUB_OWNER, GITHUB_REPO, EMAIL_USER, EMAIL_PASS, NOTIFY_EMAIL } = process.env;

  // Validate input
  const err = validate(req.body || {});
  if (err) return res.status(400).json({ error: err });

  // Sanitize fields
  const name    = sanitizeString(req.body.name).slice(0, 100);
  const email   = (req.body.email || '').trim().toLowerCase().slice(0, 200);
  const message = sanitizeString(req.body.message).slice(0, 2000);

  const now      = new Date();
  const dateStr  = now.toLocaleDateString('en-IN', { timeZone: 'Asia/Kolkata', day: '2-digit', month: '2-digit', year: 'numeric' });
  const timeStr  = now.toLocaleTimeString('en-IN', { timeZone: 'Asia/Kolkata', hour: '2-digit', minute: '2-digit' });
  const datetime = `${dateStr} ${timeStr}`;

  // ── Store in GitHub Excel (or local Excel fallback) ──
  const rowData = [name, email, message, dateStr, timeStr];

  if (GITHUB_TOKEN && GITHUB_OWNER && GITHUB_REPO) {
    let retries = 3;
    let stored = false;
    while (retries-- > 0) {
      try {
        const existing = await ghGet(GITHUB_TOKEN, GITHUB_OWNER, GITHUB_REPO, FILE_PATH);
        let existingRows = [];
        if (existing && existing.content) {
          existingRows = readWorkbook(Buffer.from(existing.content, 'base64'));
        }
        existingRows.push(rowData);

        const wb  = buildWorkbook(existingRows);
        const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        const b64 = buf.toString('base64');

        await ghPut(
          GITHUB_TOKEN, GITHUB_OWNER, GITHUB_REPO, FILE_PATH,
          b64,
          existing ? existing.sha : undefined,
          `inquiry: new message from ${name}`
        );
        stored = true;
        break;
      } catch (e) {
        if (retries === 0) {
          console.error('GitHub storage failed after retries:', e.message);
          return res.status(500).json({ error: 'Failed to record message to repository. Please email directly.' });
        }
        await new Promise(r => setTimeout(r, 400));
      }
    }
  } else {
    // Local fallback for local development & testing
    try {
      saveLocalExcel(rowData);
    } catch (localErr) {
      console.error('Local Excel fallback error:', localErr.message);
    }
  }

  // ── Send email notification (non-blocking) ──
  try {
    if (EMAIL_USER && EMAIL_PASS && (NOTIFY_EMAIL || 'aibinjoseph9605573691@gmail.com')) {
      await sendNotification({
        emailUser: EMAIL_USER,
        emailPass: EMAIL_PASS,
        notifyEmail: NOTIFY_EMAIL || 'aibinjoseph9605573691@gmail.com',
        name,
        email,
        message,
        datetime,
      });
    }
  } catch (emailErr) {
    console.error('Email notification failed (non-fatal):', emailErr.message);
  }

  return res.status(200).json({
    success: true,
    message: 'Message sent successfully. Thank you for reaching out — I\'ll get back to you soon!',
  });
};
