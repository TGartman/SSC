#!/usr/bin/env node
/* Blocks committing obvious secrets by scanning STAGED content. */

const { execSync } = require("child_process");

if (process.env.SKIP_SECRET_SCAN === "1") {
  process.exit(0);
}

function sh(cmd) {
  return execSync(cmd, { stdio: ["ignore", "pipe", "pipe"], encoding: "utf8" }).trim();
}

function isProbablyText(buf) {
  // Simple heuristic: if it contains NUL, treat as binary.
  return !buf.includes("\u0000");
}

function getStagedFiles() {
  const out = sh("git diff --cached --name-only --diff-filter=ACM");
  return out ? out.split(/\r?\n/).filter(Boolean) : [];
}

function getStagedFileContent(path) {
  // Read from the index, not working tree
  try {
    return execSync(`git show :${JSON.stringify(path)}`, { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] });
  } catch {
    // If file can't be read (binary or weird path), skip
    return null;
  }
}

const patterns = [
  // Private keys / certs
  { name: "Private key block", re: /-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----/ },

  // Common secret-ish env vars / JSON keys
  { name: "Azure client secret key", re: /\bCLIENT_SECRET\b|\bAZURE_CLIENT_SECRET\b|\bAAD_CLIENT_SECRET\b/i },
  { name: "Generic 'secret' assignment", re: /\b(secret|client_secret|app_secret)\b\s*[:=]\s*["'][^"']{8,}["']/i },
  { name: "Password assignment", re: /\b(password|passwd|pwd)\b\s*[:=]\s*["'][^"']{8,}["']/i },

  // Tokens (broad but helpful)
  { name: "JWT (eyJ...)", re: /\beyJ[a-zA-Z0-9_-]{10,}\.[a-zA-Z0-9_-]{10,}\.[a-zA-Z0-9_-]{10,}\b/ },
  { name: "GitHub token", re: /\bgh[pousr]_[A-Za-z0-9_]{20,}\b/ },
  { name: "Slack token", re: /\bxox[baprs]-[A-Za-z0-9-]{10,}\b/ },
];

const allowFilePatterns = [
  // Allow templates/examples (still fine to scan, but less noisy if you prefer)
  // /(?:^|\/)\.env\.example$/i,
  // /(?:^|\/)local\.settings\.json\.example$/i,
];

function isAllowedFile(path) {
  return allowFilePatterns.some((re) => re.test(path));
}

const files = getStagedFiles();

let violations = [];

for (const file of files) {
  if (isAllowedFile(file)) continue;

  const content = getStagedFileContent(file);
  if (!content) continue;
  if (!isProbablyText(content)) continue;

  // Skip huge files to keep commits snappy
  if (content.length > 2_000_000) continue;

  for (const p of patterns) {
    const m = content.match(p.re);
    if (m) {
      violations.push({ file, rule: p.name, match: m[0].slice(0, 120) });
    }
  }
}

if (violations.length) {
  console.error("\n❌ Commit blocked: possible secret(s) detected in STAGED files:\n");
  for (const v of violations) {
    console.error(`- ${v.file}\n  Rule: ${v.rule}\n  Snip: ${v.match}\n`);
  }

  console.error(
    [
      "Fix options:",
      "1) Remove the secret from the file and commit again.",
      "2) Move it to local.settings.json / .env.local (ignored) and reference via env vars.",
      "3) If this is a false positive, rewrite the value to a safe placeholder.",
      "",
      "Emergency bypass (use sparingly):",
      "  $env:SKIP_SECRET_SCAN='1'; git commit -m \"...\"",
      "",
    ].join("\n")
  );

  process.exit(1);
}

process.exit(0);