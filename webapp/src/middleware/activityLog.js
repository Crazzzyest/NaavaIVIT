// In-memory ring buffer of recent system events.
// Lost on restart — for permanent history, write to a sheet tab later.

const MAX_ENTRIES = 200;
let entries = [];
let nextId = 1;

function log(type, message, meta = {}) {
  const entry = {
    id: nextId++,
    ts: new Date().toISOString(),
    type,         // 'patch' | 'status' | 'cron' | 'cron-error' | 'register' | 'faktura' | 'login' | 'error' | 'admin'
    message,
    meta,
  };
  entries.push(entry);
  if (entries.length > MAX_ENTRIES) {
    entries.splice(0, entries.length - MAX_ENTRIES);
  }
  return entry;
}

function list(limit = 50) {
  return entries.slice(-limit).reverse();
}

function clear() {
  entries = [];
}

// Convenience helpers
const logger = {
  log,
  list,
  clear,
  patch:    (msg, meta) => log('patch', msg, meta),
  status:   (msg, meta) => log('status', msg, meta),
  cron:     (msg, meta) => log('cron', msg, meta),
  cronErr:  (msg, meta) => log('cron-error', msg, meta),
  register: (msg, meta) => log('register', msg, meta),
  faktura:  (msg, meta) => log('faktura', msg, meta),
  login:    (msg, meta) => log('login', msg, meta),
  error:    (msg, meta) => log('error', msg, meta),
  admin:    (msg, meta) => log('admin', msg, meta),
};

module.exports = logger;
