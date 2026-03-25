import "dotenv/config";
import crypto from "crypto";
import express from "express";
import cors from "cors";
import { google } from "googleapis";
import sqlite3 from "sqlite3";

const {
  PORT = "3000",
  GOOGLE_CLIENT_ID,
  GOOGLE_CLIENT_SECRET,
  GOOGLE_REDIRECT_URI,
  APP_URL,
  SQLITE_PATH = "./data.sqlite",
} = process.env;

if (!GOOGLE_CLIENT_ID || !GOOGLE_CLIENT_SECRET || !GOOGLE_REDIRECT_URI || !APP_URL) {
  throw new Error("Missing required env vars. See README for setup.");
}

const db = new sqlite3.Database(SQLITE_PATH);
db.serialize(() => {
  db.run(
    "CREATE TABLE IF NOT EXISTS users (email TEXT PRIMARY KEY, refresh_token TEXT NOT NULL)",
  );
  db.run(
    "CREATE TABLE IF NOT EXISTS sessions (token TEXT PRIMARY KEY, email TEXT NOT NULL, created_at INTEGER NOT NULL)",
  );
  db.run(
    "CREATE TABLE IF NOT EXISTS auth_states (state TEXT PRIMARY KEY, return_url TEXT NOT NULL, created_at INTEGER NOT NULL)",
  );
});

const app = express();
app.use(express.json({ limit: "4mb" }));
app.use(
  cors({
    origin: APP_URL,
    credentials: false,
  }),
);

const oauth2Client = new google.auth.OAuth2(
  GOOGLE_CLIENT_ID,
  GOOGLE_CLIENT_SECRET,
  GOOGLE_REDIRECT_URI,
);

const drive = google.drive({ version: "v3", auth: oauth2Client });
const oauth2 = google.oauth2({ version: "v2", auth: oauth2Client });

function randomToken(size = 32) {
  return crypto.randomBytes(size).toString("hex");
}

function storeAuthState(state, returnUrl) {
  return new Promise((resolve, reject) => {
    db.run(
      "INSERT INTO auth_states (state, return_url, created_at) VALUES (?, ?, ?)",
      [state, returnUrl, Date.now()],
      (err) => (err ? reject(err) : resolve()),
    );
  });
}

function consumeAuthState(state) {
  return new Promise((resolve, reject) => {
    db.get("SELECT return_url FROM auth_states WHERE state = ?", [state], (err, row) => {
      if (err) return reject(err);
      db.run("DELETE FROM auth_states WHERE state = ?", [state], (delErr) => {
        if (delErr) return reject(delErr);
        resolve(row?.return_url || null);
      });
    });
  });
}

function upsertUser(email, refreshToken) {
  return new Promise((resolve, reject) => {
    db.run(
      "INSERT INTO users (email, refresh_token) VALUES (?, ?) ON CONFLICT(email) DO UPDATE SET refresh_token=excluded.refresh_token",
      [email, refreshToken],
      (err) => (err ? reject(err) : resolve()),
    );
  });
}

function getUser(email) {
  return new Promise((resolve, reject) => {
    db.get("SELECT email, refresh_token FROM users WHERE email = ?", [email], (err, row) => {
      if (err) return reject(err);
      resolve(row || null);
    });
  });
}

function createSession(email) {
  return new Promise((resolve, reject) => {
    const token = randomToken(24);
    db.run(
      "INSERT INTO sessions (token, email, created_at) VALUES (?, ?, ?)",
      [token, email, Date.now()],
      (err) => (err ? reject(err) : resolve(token)),
    );
  });
}

function getSession(token) {
  return new Promise((resolve, reject) => {
    db.get("SELECT email FROM sessions WHERE token = ?", [token], (err, row) => {
      if (err) return reject(err);
      resolve(row || null);
    });
  });
}

app.get("/auth/start", async (req, res) => {
  try {
    const returnUrl = req.query.returnUrl || APP_URL;
    const state = randomToken(16);
    await storeAuthState(state, returnUrl);
    const url = oauth2Client.generateAuthUrl({
      access_type: "offline",
      prompt: "consent",
      scope: [
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/userinfo.email",
      ],
      state,
    });
    res.redirect(url);
  } catch (err) {
    res.status(500).send("Auth start failed");
  }
});

app.get("/auth/callback", async (req, res) => {
  try {
    const { code, state } = req.query;
    if (!code || !state) {
      return res.status(400).send("Missing code/state");
    }

    const returnUrl = await consumeAuthState(state);
    if (!returnUrl) {
      return res.status(400).send("Invalid state");
    }

    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);
    const userInfo = await oauth2.userinfo.get();
    const email = userInfo?.data?.email;
    if (!email) {
      return res.status(400).send("Missing user email");
    }

    if (tokens.refresh_token) {
      await upsertUser(email, tokens.refresh_token);
    } else {
      const existing = await getUser(email);
      if (!existing) {
        return res.status(400).send("Refresh token not provided. Re-consent required.");
      }
    }

    const sessionToken = await createSession(email);
    const redirectUrl = new URL(returnUrl);
    redirectUrl.searchParams.set("drive_auth", "success");
    redirectUrl.searchParams.set("drive_token", sessionToken);
    res.redirect(redirectUrl.toString());
  } catch (err) {
    res.status(500).send("Auth callback failed");
  }
});

app.post("/api/drive/upload", async (req, res) => {
  try {
    const authHeader = req.headers.authorization || "";
    const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : "";
    if (!token) {
      return res.status(401).json({ error: "Missing session token" });
    }

    const session = await getSession(token);
    if (!session) {
      return res.status(401).json({ error: "Invalid session" });
    }

    const { filename, content, mimeType = "application/msword" } = req.body || {};
    if (!filename || !content) {
      return res.status(400).json({ error: "Missing filename/content" });
    }

    const user = await getUser(session.email);
    if (!user) {
      return res.status(401).json({ error: "User not found" });
    }

    oauth2Client.setCredentials({ refresh_token: user.refresh_token });
    const file = await drive.files.create({
      requestBody: { name: filename },
      media: {
        mimeType,
        body: Buffer.from(content, "utf8"),
      },
      fields: "id, name",
    });

    res.json({ id: file.data.id, name: file.data.name });
  } catch (err) {
    res.status(500).json({ error: "Upload failed" });
  }
});

app.get("/healthz", (_req, res) => {
  res.json({ ok: true });
});

app.listen(Number(PORT), () => {
  console.log(`Backend running on port ${PORT}`);
});
