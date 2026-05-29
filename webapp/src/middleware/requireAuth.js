function requireAuth(req, res, next) {
  if (req.session && req.session.user) {
    return next();
  }
  // For API/XHR requests, return 401 JSON so the frontend can react
  const wantsJson = req.path.startsWith('/api/') || (req.get('accept') || '').includes('application/json');
  if (wantsJson) {
    return res.status(401).json({ success: false, error: 'unauthenticated' });
  }
  return res.redirect('/login');
}

module.exports = requireAuth;
