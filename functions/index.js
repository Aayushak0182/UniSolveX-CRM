const functions = require("firebase-functions");
const express = require("express");
const admin = require("firebase-admin");

admin.initializeApp();
const db = admin.firestore();

const app = express();
app.use(express.json());

// 🔐 VERIFY WEBHOOK
app.get("/webhook", (req, res) => {
  const VERIFY_TOKEN = "unisolvex123";

  const mode = req.query["hub.mode"];
  const token = req.query["hub.verify_token"];
  const challenge = req.query["hub.challenge"];

  if (mode && token === VERIFY_TOKEN) {
    return res.status(200).send(challenge);
  } else {
    return res.sendStatus(403);
  }
});

// 📩 RECEIVE WHATSAPP MESSAGE
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