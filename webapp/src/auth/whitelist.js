const config = require('../config');

function isAccountant(email) {
  if (!email) return false;
  return config.accountantEmails.includes(String(email).toLowerCase());
}

function isAllowed(email) {
  if (!email) return false;
  const e = String(email).toLowerCase();
  // Accountant emails can always log in (they get the restricted Regnskap view)
  if (isAccountant(e)) return true;
  if (config.allowedEmails.length === 0) {
    // Fail-safe: empty whitelist blocks everyone. Set ALLOWED_EMAILS env var to grant access.
    return false;
  }
  return config.allowedEmails.includes(e);
}

function roleFor(email) {
  return isAccountant(email) ? 'accountant' : 'full';
}

module.exports = { isAllowed, isAccountant, roleFor };
