const config = require('../config');

function isAllowed(email) {
  if (!email) return false;
  if (config.allowedEmails.length === 0) {
    // Fail-safe: empty whitelist blocks everyone. Set ALLOWED_EMAILS env var to grant access.
    return false;
  }
  return config.allowedEmails.includes(String(email).toLowerCase());
}

module.exports = { isAllowed };
