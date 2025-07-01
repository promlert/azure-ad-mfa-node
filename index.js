const express = require("express");
const msal = require("@azure/msal-node");
const session = require("express-session");
const dotenv = require("dotenv");

dotenv.config();

const app = express();
const port = 3000;

app.use(
  session({
    secret: process.env.SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
  })
);

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// MSAL configuration with verbose logging
const msalConfig = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (!containsPii) console.log(`[MSAL] ${message}`);
      },
      logLevel: msal.LogLevel.Verbose,
      piiLoggingEnabled: false,
    },
  },
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

// Routes
app.get("/", (req, res) => {
  if (req.session.isAuthenticated) {
    res.send(`
      <h1>Welcome, ${req.session.account?.name}</h1>
      <p><a href="/logout">Logout</a></p>
    `);
  } else {
    res.send(`
      <h1>Node.js Azure AD MFA Demo</h1>
      <p><a href="/login">Login with Azure AD</a></p>
    `);
  }
});

app.get("/login", async (req, res) => {
  const authCodeUrlParameters = {
    scopes: ["User.Read"],
    redirectUri: process.env.REDIRECT_URI,
  };

  try {
    const response = await cca.getAuthCodeUrl(authCodeUrlParameters);
    res.redirect(response);
  } catch (error) {
    console.error("Login error:", error);
    res.status(500).send(`Login error: ${error.message}`);
  }
});

app.get("/auth/redirect", async (req, res) => {
  try {
    if (!req.query.code) {
      throw new Error("Authorization code missing in redirect URL");
    }
    console.log("Redirect URI:", process.env.REDIRECT_URI);
    console.log("Received auth code:", req.query.code);

    const tokenRequest = {
      code: req.query.code,
      scopes: ["User.Read"],
      redirectUri: process.env.REDIRECT_URI,
    };

    const response = await cca.acquireTokenByCode(tokenRequest);
    req.session.isAuthenticated = true;
    req.session.account = response.account;
    req.session.accessToken = response.accessToken;
    res.redirect("/");
  } catch (error) {
    console.error("Token acquisition error:", error);
    res.status(500).send(`Token acquisition error: ${error.message}`);
  }
});

app.get("/logout", (req, res) => {
  req.session.destroy(() => {
    res.redirect(
      `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/logout?post_logout_redirect_uri=http://localhost:${port}/`
    );
  });
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});