/**
 * AES-256-GCM encryption for sensitive settings stored in Google Sheets.
 *
 * Master key lives in MASTER_ENCRYPTION_KEY env var (64 hex chars = 32 bytes).
 * Ciphertext format: "enc:" + base64(IV || authTag || ciphertext)
 *
 * Plaintext values (no "enc:" prefix) pass through unchanged on decrypt,
 * for backward compatibility with values written before encryption was added.
 */

const crypto = require('crypto');

const PREFIX = 'enc:';
const IV_LENGTH = 12;     // 96 bits — GCM standard
const AUTH_TAG_LENGTH = 16; // 128 bits

let _keyBuffer = null;
let _keyWarned = false;

function getKey() {
  if (_keyBuffer) return _keyBuffer;
  const raw = process.env.MASTER_ENCRYPTION_KEY;
  if (!raw) {
    if (!_keyWarned) {
      console.warn('MASTER_ENCRYPTION_KEY not set — secret fields will be stored in plaintext. Generate one with: openssl rand -hex 32');
      _keyWarned = true;
    }
    return null;
  }
  const buf = Buffer.from(raw.trim(), 'hex');
  if (buf.length !== 32) {
    throw new Error(`MASTER_ENCRYPTION_KEY must be exactly 32 bytes (64 hex chars). Got ${buf.length} bytes.`);
  }
  _keyBuffer = buf;
  return _keyBuffer;
}

function isAvailable() {
  return !!process.env.MASTER_ENCRYPTION_KEY;
}

function encrypt(plaintext) {
  if (plaintext === null || plaintext === undefined || plaintext === '') return '';
  const key = getKey();
  if (!key) return String(plaintext); // No key → store as plaintext (with warning above)

  const iv = crypto.randomBytes(IV_LENGTH);
  const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
  const encrypted = Buffer.concat([cipher.update(String(plaintext), 'utf8'), cipher.final()]);
  const authTag = cipher.getAuthTag();

  return PREFIX + Buffer.concat([iv, authTag, encrypted]).toString('base64');
}

function decrypt(value) {
  if (value === null || value === undefined || value === '') return '';
  const s = String(value);
  if (!s.startsWith(PREFIX)) {
    // Plaintext (legacy or no-encryption mode) — return as-is
    return s;
  }
  const key = getKey();
  if (!key) {
    throw new Error('Cannot decrypt: MASTER_ENCRYPTION_KEY not set');
  }

  const raw = Buffer.from(s.slice(PREFIX.length), 'base64');
  const iv = raw.subarray(0, IV_LENGTH);
  const authTag = raw.subarray(IV_LENGTH, IV_LENGTH + AUTH_TAG_LENGTH);
  const encrypted = raw.subarray(IV_LENGTH + AUTH_TAG_LENGTH);

  const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv);
  decipher.setAuthTag(authTag);
  try {
    const decrypted = Buffer.concat([decipher.update(encrypted), decipher.final()]);
    return decrypted.toString('utf8');
  } catch (e) {
    throw new Error(`Decryption failed: ${e.message}. Wrong key, or value was encrypted with a different key.`);
  }
}

function isEncrypted(value) {
  return typeof value === 'string' && value.startsWith(PREFIX);
}

module.exports = { encrypt, decrypt, isEncrypted, isAvailable };
