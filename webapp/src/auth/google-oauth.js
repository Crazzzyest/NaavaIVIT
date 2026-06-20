const passport = require('passport');
const { Strategy: GoogleStrategy } = require('passport-google-oauth20');
const config = require('../config');
const { isAllowed, roleFor } = require('./whitelist');

function configurePassport() {
  if (!config.loginOAuth.clientId || !config.loginOAuth.clientSecret) {
    console.warn('Login OAuth not configured (LOGIN_OAUTH_CLIENT_ID/SECRET missing). /auth/google will fail.');
    return;
  }

  passport.use(new GoogleStrategy(
    {
      clientID: config.loginOAuth.clientId,
      clientSecret: config.loginOAuth.clientSecret,
      callbackURL: config.loginOAuth.callbackUrl,
    },
    (accessToken, refreshToken, profile, done) => {
      const email = (profile.emails && profile.emails[0] && profile.emails[0].value || '').toLowerCase();
      if (!email) return done(null, false, { message: 'No email in Google profile' });
      if (!isAllowed(email)) return done(null, false, { message: 'denied' });
      return done(null, {
        email,
        name: profile.displayName || email,
        picture: profile.photos && profile.photos[0] && profile.photos[0].value || null,
        role: roleFor(email),
      });
    },
  ));

  passport.serializeUser((user, done) => done(null, user));
  passport.deserializeUser((user, done) => done(null, user));
}

module.exports = { configurePassport };
