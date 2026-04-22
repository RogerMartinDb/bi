const msal = require('@azure/msal-node');

const SCOPES = ['openid', 'profile', 'email', 'User.Read'];

function createMsalClient() {
  return new msal.ConfidentialClientApplication({
    auth: {
      clientId: process.env.AZURE_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
      clientSecret: process.env.AZURE_CLIENT_SECRET,
    },
  });
}

function registerAuthRoutes(app) {
  const msalClient = createMsalClient();
  const cryptoProvider = new msal.CryptoProvider();

  // Kick off login — redirect to Microsoft
  app.get('/auth/login', async (req, res) => {
    const { verifier, challenge } = await cryptoProvider.generatePkceCodes();
    req.session.pkceverifier = verifier;
    req.session.state = cryptoProvider.createNewGuid();

    const url = await msalClient.getAuthCodeUrl({
      scopes: SCOPES,
      redirectUri: process.env.REDIRECT_URI,
      codeChallenge: challenge,
      codeChallengeMethod: 'S256',
      state: req.session.state,
    });
    res.redirect(url);
  });

  // Microsoft redirects back here with an auth code
  app.get('/auth/callback', async (req, res) => {
    if (req.query.error) {
      console.error('Auth error:', req.query.error, req.query.error_description);
      return res.status(401).send('Authentication failed: ' + req.query.error_description);
    }

    if (req.query.state !== req.session.state) {
      return res.status(400).send('State mismatch — possible CSRF.');
    }

    try {
      const result = await msalClient.acquireTokenByCode({
        code: req.query.code,
        scopes: SCOPES,
        redirectUri: process.env.REDIRECT_URI,
        codeVerifier: req.session.pkceverifier,
      });

      req.session.account = {
        name: result.account.name,
        username: result.account.username,
        homeAccountId: result.account.homeAccountId,
      };
      delete req.session.pkceverifier;
      delete req.session.state;

      res.redirect(req.session.returnTo || '/');
      delete req.session.returnTo;
    } catch (err) {
      console.error('Token acquisition failed:', err);
      res.status(500).send('Authentication error. Please try again.');
    }
  });

  // Logout — clear session and redirect to Microsoft logout
  app.get('/auth/logout', (req, res) => {
    const logoutUri = `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/logout`
      + `?post_logout_redirect_uri=${encodeURIComponent(process.env.POST_LOGOUT_REDIRECT_URI)}`;
    req.session.destroy(() => res.redirect(logoutUri));
  });

  // Returns the current user for the frontend
  app.get('/auth/me', requireAuth, (req, res) => {
    res.json(req.session.account);
  });
}

// Middleware: blocks unauthenticated requests
function requireAuth(req, res, next) {
  if (req.session && req.session.account) return next();
  if (req.path.startsWith('/api/')) return res.status(401).json({ error: 'Unauthorized' });
  req.session.returnTo = req.originalUrl;
  res.redirect('/auth/login');
}

module.exports = { registerAuthRoutes, requireAuth };
