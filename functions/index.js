const fs = require("fs");
const path = require("path");
const functions = require("firebase-functions");
const express = require("express");
const admin = require("firebase-admin");

function loadEnvFile(filePath) {
  try {
    if (!fs.existsSync(filePath)) return;
    const raw = fs.readFileSync(filePath, "utf8");
    raw.split(/\r?\n/).forEach((line) => {
      const trimmed = String(line || "").trim();
      if (!trimmed || trimmed.startsWith("#")) return;
      const eqIndex = trimmed.indexOf("=");
      if (eqIndex <= 0) return;
      const key = trimmed.slice(0, eqIndex).trim();
      if (!key || process.env[key]) return;
      let value = trimmed.slice(eqIndex + 1).trim();
      if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
        value = value.slice(1, -1);
      }
      process.env[key] = value;
    });
  } catch (error) {
    console.warn("[env] Failed to load", filePath, error?.message || error);
  }
}

loadEnvFile(path.join(__dirname, ".env"));
loadEnvFile(path.join(__dirname, ".env.local"));
loadEnvFile(path.join(__dirname, "..", ".env"));
loadEnvFile(path.join(__dirname, "..", ".env.local"));

admin.initializeApp();
const db = admin.firestore();

const firebaseWebApiKey = String(process.env.FIREBASE_WEB_API_KEY || "").trim();
const firebaseWebAuthDomain = String(process.env.FIREBASE_WEB_AUTH_DOMAIN || "").trim();
const firebaseWebProjectId = String(process.env.FIREBASE_WEB_PROJECT_ID || process.env.GCLOUD_PROJECT || "").trim();
const firebaseWebAppId = String(process.env.FIREBASE_WEB_APP_ID || "").trim();
const firebaseWebMessagingSenderId = String(process.env.FIREBASE_WEB_MESSAGING_SENDER_ID || "").trim();
const firebaseWebStorageBucket = String(process.env.FIREBASE_WEB_STORAGE_BUCKET || "").trim();

const firebaseAuthAdminEmails = new Set(
  String(process.env.FIREBASE_AUTH_ADMIN_EMAILS || "")
    .split(",")
    .map((value) => String(value || "").trim().toLowerCase())
    .filter(Boolean)
);
const firebaseAuthAgentEmails = new Set(
  String(process.env.FIREBASE_AUTH_AGENT_EMAILS || "")
    .split(",")
    .map((value) => String(value || "").trim().toLowerCase())
    .filter(Boolean)
);

function getFirebaseWebConfig() {
  return {
    apiKey: firebaseWebApiKey,
    authDomain: firebaseWebAuthDomain,
    projectId: firebaseWebProjectId,
    appId: firebaseWebAppId,
    messagingSenderId: firebaseWebMessagingSenderId,
    storageBucket: firebaseWebStorageBucket
  };
}

function hasCompleteFirebaseWebConfig() {
  const config = getFirebaseWebConfig();
  return Boolean(config.apiKey && config.authDomain && config.projectId && config.appId);
}

function resolveFirebaseAuthRole(email) {
  const normalizedEmail = String(email || "").trim().toLowerCase();
  if (!normalizedEmail) return "";
  if (firebaseAuthAdminEmails.has(normalizedEmail)) return "admin";
  if (!firebaseAuthAgentEmails.size) return "agent";
  return firebaseAuthAgentEmails.has(normalizedEmail) ? "agent" : "";
}

function deriveNameFromEmail(email, fallbackDisplayName) {
  const displayName = String(fallbackDisplayName || "").trim();
  if (displayName) return displayName;
  const localPart = String(email || "").split("@")[0] || "User";
  return localPart
    .replace(/[._-]+/g, " ")
    .replace(/\b\w/g, (char) => char.toUpperCase())
    .trim() || "User";
}

const app = express();
app.use(express.json());

app.get("/api/auth/firebase-config", (_req, res) => {
  if (!hasCompleteFirebaseWebConfig()) {
    return res.status(503).json({
      ok: false,
      error: "Firebase login config is not ready"
    });
  }
  return res.json({
    ok: true,
    config: getFirebaseWebConfig()
  });
});

app.post("/api/auth/firebase-session", async (req, res) => {
  const idToken = String(req.body?.idToken || "").trim();
  if (!idToken) {
    return res.status(400).json({ ok: false, error: "idToken is required" });
  }
  try {
    const decoded = await admin.auth().verifyIdToken(idToken, true);
    const email = String(decoded?.email || "").trim().toLowerCase();
    const role = resolveFirebaseAuthRole(email);
    if (!email || !role) {
      return res.status(403).json({ ok: false, error: "This email is not authorized for CRM access" });
    }
    return res.json({
      ok: true,
      user: {
        uid: String(decoded.uid || "").trim(),
        email,
        role,
        name: deriveNameFromEmail(email, decoded?.name || decoded?.displayName || "")
      }
    });
  } catch (error) {
    return res.status(401).json({
      ok: false,
      error: error?.message || "Firebase session verification failed"
    });
  }
});

// VERIFY WEBHOOK
app.get("/webhook", (req, res) => {
  const VERIFY_TOKEN = "unisolvex123";
  const mode = req.query["hub.mode"];
  const token = req.query["hub.verify_token"];
  const challenge = req.query["hub.challenge"];

  if (mode && token === VERIFY_TOKEN) {
    return res.status(200).send(challenge);
  }
  return res.sendStatus(403);
});

// RECEIVE WHATSAPP MESSAGE
app.post("/webhook", async (req, res) => {
  try {
    const body = req.body;
    if (body.entry) {
      const message = body.entry[0].changes[0].value.messages?.[0];
      if (message) {
        await db.collection("messages").add({
          from: message.from,
          text: message.text?.body || "",
          timestamp: Date.now()
        });
      }
    }
    res.sendStatus(200);
  } catch (error) {
    console.error(error);
    res.sendStatus(500);
  }
});

exports.api = functions.https.onRequest(app);
