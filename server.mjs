import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import { WebSocketServer } from 'ws';
import crypto from 'node:crypto';
import admin from 'firebase-admin';

const app = express();
const port = Number(process.env.PORT || 3001);
const verifyToken = process.env.WHATSAPP_VERIFY_TOKEN || '';
const appSecret = process.env.META_APP_SECRET || process.env.WHATSAPP_APP_SECRET || '';
const whatsappAccessToken = process.env.WHATSAPP_ACCESS_TOKEN || '';
const whatsappPhoneNumberId = process.env.WHATSAPP_PHONE_NUMBER_ID || '';
const whatsappInitiationTemplateName = process.env.WHATSAPP_INIT_TEMPLATE_NAME || '';
const whatsappInitiationTemplateLanguage = process.env.WHATSAPP_INIT_TEMPLATE_LANG || 'en_US';
const whatsappInitiationTemplateParamOrder = process.env.WHATSAPP_INIT_TEMPLATE_PARAM_ORDER;
const whatsappInitiationTemplatePreviewText = process.env.WHATSAPP_INIT_TEMPLATE_PREVIEW_TEXT || 'Hello 👋,\n\nThis is UniSolvex team. We would like to start a conversation with you.\n\nPlease reply to this message so we can talk.\n\nThank you!';
const firebaseServiceAccountJson = process.env.FIREBASE_SERVICE_ACCOUNT_JSON || process.env.GOOGLE_SERVICE_ACCOUNT_JSON || '';
const firebaseServiceAccountBase64 = process.env.FIREBASE_SERVICE_ACCOUNT_BASE64 || '';
const firebaseStorageBucketName = process.env.FIREBASE_STORAGE_BUCKET || '';
const whatsappMediaStoragePrefix = process.env.WHATSAPP_MEDIA_STORAGE_PREFIX || 'whatsapp-media';

app.use(cors());
app.disable('x-powered-by');
app.set('trust proxy', 1);
app.use(
    express.json({
        limit: '12mb',
        verify: (req, _res, buf, encoding) => {
            req.rawBody = buf.toString(encoding || 'utf8');
        }
    })
);

const contactsByWaId = new Map();
const messagesByWaId = new Map();
const messageStatusById = new Map();
const sockets = new Set();
const webhookEvents = [];
const apiEvents = [];
const crmState = {
    orderListHtml: '',
    manualContacts: [],
    whatsappContactIdMap: {},
    whatsappReadState: {},
    whatsappContactIdSequence: 100100,
    updatedAt: ''
};

let firebaseDb = null;
let firebaseBucket = null;
let firebasePersistenceInitError = '';
let firebasePersistenceEnabled = false;
let persistedStateLoaded = false;
let persistedStateLoadInFlight = null;
let lastPersistedStateLoadAt = '';
let lastPersistedStateError = '';
let lastFirebaseWriteAt = '';
let lastFirebaseWriteError = '';
let crmStateLoaded = false;
let crmStateLoadInFlight = null;

let loggedMissingAppSecret = false;
let loggedMissingFirebasePersistence = false;

function parseFirebaseServiceAccount() {
    try {
        if (firebaseServiceAccountJson) {
            return JSON.parse(firebaseServiceAccountJson);
        }
        if (firebaseServiceAccountBase64) {
            return JSON.parse(Buffer.from(firebaseServiceAccountBase64, 'base64').toString('utf8'));
        }
    } catch (error) {
        firebasePersistenceInitError = error?.message || 'Invalid Firebase service account JSON';
    }
    return null;
}

function initFirebasePersistence() {
    if (firebasePersistenceEnabled || firebaseDb) return true;
    try {
        const serviceAccount = parseFirebaseServiceAccount();
        if (!serviceAccount && !process.env.GOOGLE_APPLICATION_CREDENTIALS) {
            return false;
        }
        if (!admin.apps.length) {
            const options = {
                ...(firebaseStorageBucketName ? { storageBucket: firebaseStorageBucketName } : {})
            };
            if (serviceAccount) {
                options.credential = admin.credential.cert(serviceAccount);
            }
            admin.initializeApp(options);
        }
        firebaseDb = admin.firestore();
        firebaseBucket = firebaseStorageBucketName ? admin.storage().bucket(firebaseStorageBucketName) : null;
        firebasePersistenceEnabled = true;
        return true;
    } catch (error) {
        firebasePersistenceInitError = error?.message || 'Firebase persistence init failed';
        return false;
    }
}

function warnMissingFirebasePersistenceOnce() {
    if (loggedMissingFirebasePersistence) return;
    loggedMissingFirebasePersistence = true;
    console.warn('[storage] Firebase persistence is not configured. Chat history will remain in-memory only.');
}

function getPersistentMessageDocId(waId, message) {
    const normalizedWaId = normalizeWaId(waId);
    const messageId = String(message?.id || '').trim();
    const seed = messageId
        ? `${normalizedWaId}|${messageId}`
        : [
            normalizedWaId,
            String(message?.direction || ''),
            String(message?.timestamp || ''),
            String(message?.messageType || ''),
            String(message?.text || ''),
            String(message?.attachmentName || ''),
            String(message?.mediaId || ''),
            String(message?.replyTo?.id || '')
        ].join('|');
    return crypto.createHash('sha1').update(seed).digest('hex');
}

function inferFileExtension(mimeType, fallbackName) {
    const lowerMime = String(mimeType || '').toLowerCase();
    if (/\.(\w+)$/.test(String(fallbackName || ''))) {
        return String(fallbackName).split('.').pop();
    }
    if (lowerMime.includes('jpeg')) return 'jpg';
    if (lowerMime.includes('png')) return 'png';
    if (lowerMime.includes('gif')) return 'gif';
    if (lowerMime.includes('webp')) return 'webp';
    if (lowerMime.includes('pdf')) return 'pdf';
    if (lowerMime.includes('mpeg')) return 'mp3';
    if (lowerMime.includes('ogg')) return 'ogg';
    if (lowerMime.includes('mp4')) return 'mp4';
    if (lowerMime.includes('quicktime')) return 'mov';
    if (lowerMime.includes('plain')) return 'txt';
    return 'bin';
}

function sanitizeFileSegment(value) {
    return String(value || '')
        .replace(/[^\w.-]+/g, '_')
        .replace(/^_+|_+$/g, '')
        .slice(0, 80) || 'attachment';
}

function parseDataUrl(dataUrl) {
    const match = /^data:([^;]+);base64,(.+)$/i.exec(String(dataUrl || ''));
    if (!match?.[1] || !match?.[2]) {
        throw new Error('Attachment must be sent as a valid base64 data URL');
    }
    return {
        mimeType: match[1],
        buffer: Buffer.from(match[2], 'base64')
    };
}

async function uploadMediaBufferToFirebaseStorage({ waId, messageId, timestamp, fileName, mimeType, buffer }) {
    if (!initFirebasePersistence() || !firebaseBucket || !buffer?.length) return '';
    const safeWaId = normalizeWaId(waId) || 'unknown';
    const safeName = sanitizeFileSegment(fileName || `media_${messageId || Date.now()}.${inferFileExtension(mimeType, fileName)}`);
    const ts = String(timestamp || new Date().toISOString()).replace(/[:.]/g, '-');
    const objectPath = `${whatsappMediaStoragePrefix}/${safeWaId}/${ts}_${safeName}`;
    const file = firebaseBucket.file(objectPath);
    await file.save(buffer, {
        resumable: false,
        metadata: {
            contentType: mimeType || 'application/octet-stream',
            cacheControl: 'private, max-age=31536000'
        }
    });
    const [signedUrl] = await file.getSignedUrl({
        action: 'read',
        expires: '2035-01-01'
    });
    return signedUrl;
}

async function fetchWhatsappMediaBuffer(mediaId) {
    ensureWhatsappConfig();
    const metaResponse = await fetch(`https://graph.facebook.com/v22.0/${mediaId}`, {
        headers: {
            Authorization: `Bearer ${whatsappAccessToken}`
        }
    });
    const meta = await metaResponse.json();
    if (!metaResponse.ok || !meta?.url) {
        throw new Error(meta?.error?.message || 'Failed to resolve WhatsApp media URL');
    }
    const downloadResponse = await fetch(meta.url, {
        headers: {
            Authorization: `Bearer ${whatsappAccessToken}`
        }
    });
    if (!downloadResponse.ok) {
        throw new Error('Failed to download WhatsApp media');
    }
    const arrayBuffer = await downloadResponse.arrayBuffer();
    return {
        buffer: Buffer.from(arrayBuffer),
        mimeType: meta.mime_type || downloadResponse.headers.get('content-type') || 'application/octet-stream'
    };
}

async function backupIncomingMediaToStorage({ waId, messageId, timestamp, attachmentName, mediaId, mimeType }) {
    if (!mediaId) return '';
    try {
        const { buffer, mimeType: resolvedMimeType } = await fetchWhatsappMediaBuffer(mediaId);
        const fileName = attachmentName && attachmentName !== 'Image' && attachmentName !== 'Video' && attachmentName !== 'Audio' && attachmentName !== 'Document'
            ? attachmentName
            : `media_${mediaId}.${inferFileExtension(mimeType || resolvedMimeType, attachmentName)}`;
        return await uploadMediaBufferToFirebaseStorage({
            waId,
            messageId,
            timestamp,
            fileName,
            mimeType: mimeType || resolvedMimeType,
            buffer
        });
    } catch (error) {
        console.warn('[storage] Failed to back up incoming media:', error?.message || error);
        return '';
    }
}

async function persistContactSnapshot(contact) {
    if (!contact?.waId) return;
    if (!initFirebasePersistence()) {
        warnMissingFirebasePersistenceOnce();
        return;
    }
    try {
        await firebaseDb.collection('wa_contacts').doc(String(contact.waId)).set({
            waId: String(contact.waId),
            profileName: String(contact.profileName || contact.waId),
            updatedAt: String(contact.updatedAt || new Date().toISOString())
        }, { merge: true });
        lastFirebaseWriteAt = new Date().toISOString();
        lastFirebaseWriteError = '';
    } catch (error) {
        lastFirebaseWriteError = error?.message || 'Failed to persist contact snapshot';
        throw error;
    }
}

async function persistMessageSnapshot(waId, message) {
    if (!waId || !message) return '';
    if (!initFirebasePersistence()) {
        warnMissingFirebasePersistenceOnce();
        return '';
    }
    const docId = getPersistentMessageDocId(waId, message);
    try {
        await firebaseDb.collection('wa_messages').doc(docId).set({
            ...message,
            waId: normalizeWaId(waId),
            docId,
            updatedAt: new Date().toISOString()
        }, { merge: true });
        lastFirebaseWriteAt = new Date().toISOString();
        lastFirebaseWriteError = '';
    } catch (error) {
        lastFirebaseWriteError = error?.message || 'Failed to persist message snapshot';
        throw error;
    }
    return docId;
}

async function persistMessageStatusSnapshot(waId, messageId, status, statusTimestamp) {
    if (!waId || !messageId || !status) return;
    if (!initFirebasePersistence()) {
        warnMissingFirebasePersistenceOnce();
        return;
    }
    const docId = getPersistentMessageDocId(waId, { id: messageId });
    try {
        await firebaseDb.collection('wa_messages').doc(docId).set({
            waId: normalizeWaId(waId),
            id: messageId,
            status,
            statusTimestamp: statusTimestamp || new Date().toISOString(),
            updatedAt: new Date().toISOString()
        }, { merge: true });
        lastFirebaseWriteAt = new Date().toISOString();
        lastFirebaseWriteError = '';
    } catch (error) {
        lastFirebaseWriteError = error?.message || 'Failed to persist message status snapshot';
        throw error;
    }
}

function normalizeCrmStatePayload(payload = {}) {
    const nextManualContacts = Array.isArray(payload.manualContacts)
        ? payload.manualContacts
            .filter((row) => row && typeof row === 'object')
            .map((row) => ({
                waId: normalizeWaId(row.waId || ''),
                profileName: String(row.profileName || ''),
                tag: String(row.tag || ''),
                universityName: String(row.universityName || ''),
                semester: String(row.semester || ''),
                timezone: String(row.timezone || '')
            }))
            .filter((row) => row.waId)
        : crmState.manualContacts;
    const nextContactIdMap = payload.whatsappContactIdMap && typeof payload.whatsappContactIdMap === 'object'
        ? Object.fromEntries(
            Object.entries(payload.whatsappContactIdMap)
                .map(([key, value]) => [normalizeWaId(key), String(value || '').trim()])
                .filter(([key, value]) => key && value)
        )
        : crmState.whatsappContactIdMap;
    const nextReadState = payload.whatsappReadState && typeof payload.whatsappReadState === 'object'
        ? Object.fromEntries(
            Object.entries(payload.whatsappReadState)
                .map(([key, value]) => [normalizeWaId(key), String(value || '').trim()])
                .filter(([key, value]) => key && value)
        )
        : crmState.whatsappReadState;
    const nextSequence = Number(payload.whatsappContactIdSequence);
    return {
        orderListHtml: typeof payload.orderListHtml === 'string' ? payload.orderListHtml : crmState.orderListHtml,
        manualContacts: nextManualContacts,
        whatsappContactIdMap: nextContactIdMap,
        whatsappReadState: nextReadState,
        whatsappContactIdSequence: Number.isFinite(nextSequence) && nextSequence > 0
            ? nextSequence
            : crmState.whatsappContactIdSequence,
        updatedAt: new Date().toISOString()
    };
}

async function loadPersistedCrmState() {
    if (!initFirebasePersistence()) return false;
    const snapshot = await firebaseDb.collection('crm_meta').doc('ui_state').get();
    if (!snapshot.exists) {
        crmStateLoaded = true;
        crmState.updatedAt = new Date().toISOString();
        return true;
    }
    const data = snapshot.data() || {};
    const nextState = normalizeCrmStatePayload(data);
    crmState.orderListHtml = nextState.orderListHtml;
    crmState.manualContacts = nextState.manualContacts;
    crmState.whatsappContactIdMap = nextState.whatsappContactIdMap;
    crmState.whatsappReadState = nextState.whatsappReadState;
    crmState.whatsappContactIdSequence = nextState.whatsappContactIdSequence;
    crmState.updatedAt = String(data.updatedAt || nextState.updatedAt);
    crmStateLoaded = true;
    return true;
}

async function ensureCrmStateLoaded(forceRefresh = false) {
    if (crmStateLoadInFlight) return crmStateLoadInFlight;
    if (!forceRefresh && crmStateLoaded) return true;
    if (!initFirebasePersistence()) return false;
    crmStateLoadInFlight = loadPersistedCrmState()
        .finally(() => {
            crmStateLoadInFlight = null;
        });
    return crmStateLoadInFlight;
}

async function persistCrmState(partialPayload = {}) {
    const nextState = normalizeCrmStatePayload({
        ...crmState,
        ...partialPayload
    });
    crmState.orderListHtml = nextState.orderListHtml;
    crmState.manualContacts = nextState.manualContacts;
    crmState.whatsappContactIdMap = nextState.whatsappContactIdMap;
    crmState.whatsappReadState = nextState.whatsappReadState;
    crmState.whatsappContactIdSequence = nextState.whatsappContactIdSequence;
    crmState.updatedAt = nextState.updatedAt;
    if (!initFirebasePersistence()) {
        warnMissingFirebasePersistenceOnce();
        return false;
    }
    await firebaseDb.collection('crm_meta').doc('ui_state').set(nextState, { merge: true });
    lastFirebaseWriteAt = new Date().toISOString();
    lastFirebaseWriteError = '';
    return true;
}

function findExistingMessageIndex(waId, candidateMessage) {
    const thread = messagesByWaId.get(waId) || [];
    const candidateId = String(candidateMessage?.id || '').trim();
    if (candidateId) {
        return thread.findIndex((message) => String(message?.id || '').trim() === candidateId);
    }
    const candidateKey = [
        String(candidateMessage?.direction || ''),
        String(candidateMessage?.timestamp || ''),
        String(candidateMessage?.messageType || ''),
        String(candidateMessage?.text || ''),
        String(candidateMessage?.attachmentName || ''),
        String(candidateMessage?.mediaId || ''),
        String(candidateMessage?.replyTo?.id || '')
    ].join('|');
    return thread.findIndex((message) => ([
        String(message?.direction || ''),
        String(message?.timestamp || ''),
        String(message?.messageType || ''),
        String(message?.text || ''),
        String(message?.attachmentName || ''),
        String(message?.mediaId || ''),
        String(message?.replyTo?.id || '')
    ].join('|')) === candidateKey);
}

function mergePersistedMessageIntoMemory(waId, message) {
    const normalizedWaId = normalizeWaId(waId);
    if (!normalizedWaId || !message) return;
    const nextMessage = {
        id: String(message.id || ''),
        text: String(message.text || ''),
        timestamp: String(message.timestamp || new Date().toISOString()),
        direction: String(message.direction || 'incoming'),
        messageType: String(message.messageType || 'text'),
        attachmentName: String(message.attachmentName || ''),
        attachmentUrl: String(message.attachmentUrl || ''),
        mediaId: String(message.mediaId || ''),
        mimeType: String(message.mimeType || ''),
        replyTo: message.replyTo || null,
        status: String(message.status || ''),
        statusTimestamp: String(message.statusTimestamp || message.timestamp || new Date().toISOString())
    };
    const existingIndex = findExistingMessageIndex(normalizedWaId, nextMessage);
    if (existingIndex >= 0) {
        const thread = messagesByWaId.get(normalizedWaId) || [];
        thread[existingIndex] = {
            ...thread[existingIndex],
            ...nextMessage
        };
        messagesByWaId.set(normalizedWaId, thread);
    } else {
        addMessage(normalizedWaId, nextMessage);
    }
}

function ensureContactForThread(waId, fallbackProfileName = '') {
    const normalizedWaId = normalizeWaId(waId);
    if (!normalizedWaId) return null;
    const existing = contactsByWaId.get(normalizedWaId);
    const thread = messagesByWaId.get(normalizedWaId) || [];
    const lastMessage = thread[thread.length - 1] || null;
    const candidateUpdatedAt = String(
        existing?.updatedAt ||
        lastMessage?.timestamp ||
        new Date().toISOString()
    );
    const nextContact = {
        waId: normalizedWaId,
        profileName: String(existing?.profileName || fallbackProfileName || normalizedWaId),
        updatedAt: candidateUpdatedAt
    };
    if (!existing) {
        contactsByWaId.set(normalizedWaId, nextContact);
        return nextContact;
    }
    const existingUpdatedAtMs = new Date(existing.updatedAt || 0).getTime() || 0;
    const candidateUpdatedAtMs = new Date(candidateUpdatedAt).getTime() || 0;
    if (candidateUpdatedAtMs >= existingUpdatedAtMs) {
        existing.updatedAt = candidateUpdatedAt;
    }
    if (!String(existing.profileName || '').trim() && nextContact.profileName) {
        existing.profileName = nextContact.profileName;
    }
    contactsByWaId.set(normalizedWaId, existing);
    return existing;
}

async function loadPersistedStateIntoMemory() {
    if (!initFirebasePersistence()) return false;
    const contactsSnapshot = await firebaseDb.collection('wa_contacts').get();
    contactsSnapshot.forEach((doc) => {
        const data = doc.data() || {};
        if (!data.waId) return;
        const waId = String(data.waId);
        const existing = contactsByWaId.get(waId);
        const nextContact = {
            waId,
            profileName: String(data.profileName || data.waId),
            updatedAt: String(data.updatedAt || new Date().toISOString())
        };
        if (!existing || new Date(nextContact.updatedAt).getTime() >= new Date(existing.updatedAt || 0).getTime()) {
            contactsByWaId.set(waId, nextContact);
        }
    });

    const messagesSnapshot = await firebaseDb.collection('wa_messages').orderBy('timestamp', 'asc').get();
    messagesSnapshot.forEach((doc) => {
        const data = doc.data() || {};
        const waId = normalizeWaId(data.waId || '');
        if (!waId) return;
        mergePersistedMessageIntoMemory(waId, data);
    });
    Array.from(messagesByWaId.keys()).forEach((waId) => {
        ensureContactForThread(waId);
    });
    persistedStateLoaded = true;
    lastPersistedStateLoadAt = new Date().toISOString();
    lastPersistedStateError = '';
    return true;
}

async function ensurePersistedStateLoaded(forceRefresh = false) {
    if (persistedStateLoadInFlight) {
        return persistedStateLoadInFlight;
    }
    if (!forceRefresh && persistedStateLoaded) {
        return true;
    }
    if (!initFirebasePersistence()) {
        return false;
    }
    persistedStateLoadInFlight = loadPersistedStateIntoMemory()
        .catch((error) => {
            persistedStateLoaded = false;
            lastPersistedStateError = error?.message || 'Failed to load persisted chat state';
            throw error;
        })
        .finally(() => {
            persistedStateLoadInFlight = null;
        });
    return persistedStateLoadInFlight;
}

function safeTimingCompareHex(expectedHex, actualHex) {
    const expected = Buffer.from(String(expectedHex || ''), 'utf8');
    const actual = Buffer.from(String(actualHex || ''), 'utf8');
    if (expected.length !== actual.length) return false;
    return crypto.timingSafeEqual(expected, actual);
}

function verifyMetaSignature(req) {
    if (!appSecret) {
        if (process.env.NODE_ENV === 'production' && !loggedMissingAppSecret) {
            loggedMissingAppSecret = true;
            console.warn('[webhook] WARNING: META_APP_SECRET/WHATSAPP_APP_SECRET not set; signature verification is disabled.');
        }
        return true;
    }

    const rawBody = String(req.rawBody || '');
    const sig256 = req.get('x-hub-signature-256');
    if (sig256) {
        const match = /^sha256=(.+)$/i.exec(String(sig256));
        if (!match?.[1]) return false;
        const provided = match[1];
        const expected = crypto.createHmac('sha256', appSecret).update(rawBody, 'utf8').digest('hex');
        return safeTimingCompareHex(expected, provided);
    }

    const sig1 = req.get('x-hub-signature');
    if (sig1) {
        const match = /^sha1=(.+)$/i.exec(String(sig1));
        if (!match?.[1]) return false;
        const provided = match[1];
        const expected = crypto.createHmac('sha1', appSecret).update(rawBody, 'utf8').digest('hex');
        return safeTimingCompareHex(expected, provided);
    }

    // If we have a secret but no signature header, treat as invalid.
    return false;
}

function pushWebhookEvent(event) {
    webhookEvents.unshift(event);
    if (webhookEvents.length > 100) webhookEvents.pop();
}

function pushApiEvent(event) {
    apiEvents.unshift(event);
    if (apiEvents.length > 100) apiEvents.pop();
}

function normalizeWaId(value) {
    return String(value || '').replace(/[^\d]/g, '');
}

function normalizeWebhookTimestamp(rawTimestamp, fallbackIso = new Date().toISOString()) {
    const fallbackMs = new Date(fallbackIso).getTime();
    const numeric = Number(rawTimestamp || '0');
    const candidateMs = numeric > 0 ? numeric * 1000 : Number.NaN;
    if (!Number.isFinite(candidateMs)) return fallbackIso;

    const minAllowedMs = Date.UTC(2020, 0, 1);
    const maxAllowedMs = fallbackMs + (5 * 60 * 1000);
    if (candidateMs < minAllowedMs || candidateMs > maxAllowedMs) {
        return fallbackIso;
    }
    return new Date(candidateMs).toISOString();
}

function upsertContact(waId, profileName) {
    if (!waId) return null;
    const existing = contactsByWaId.get(waId) || { waId, profileName: waId, updatedAt: new Date().toISOString() };
    if (profileName) existing.profileName = profileName;
    existing.updatedAt = new Date().toISOString();
    contactsByWaId.set(waId, existing);
    return existing;
}

function addMessage(waId, message) {
    if (!waId || !message) return;
    const messageId = String(message.id || '');
    const knownStatus = messageId ? messageStatusById.get(messageId) : null;
    if (knownStatus) {
        message.status = knownStatus.status;
        message.statusTimestamp = knownStatus.statusTimestamp;
    }
    if (messageId && message.status) {
        messageStatusById.set(messageId, {
            status: message.status,
            statusTimestamp: message.statusTimestamp || message.timestamp || new Date().toISOString()
        });
    }
    const existing = messagesByWaId.get(waId) || [];
    existing.push(message);
    messagesByWaId.set(waId, existing);
}

function findMessageById(waId, messageId) {
    if (!waId || !messageId) return null;
    const rows = messagesByWaId.get(waId) || [];
    return rows.find((message) => message?.id === messageId) || null;
}

function makeReplyReference(waId, contextMessageId) {
    if (!contextMessageId) return null;
    const message = findMessageById(waId, contextMessageId);
    if (!message) {
        return {
            id: contextMessageId,
            text: '(message)',
            direction: 'incoming'
        };
    }
    return {
        id: contextMessageId,
        text: message.text || message.attachmentName || '(message)',
        direction: message.direction || 'incoming'
    };
}

function updateMessageStatus(waId, messageId, status, statusTimestamp) {
    if (!waId || !messageId || !status) return false;
    messageStatusById.set(messageId, {
        status,
        statusTimestamp: statusTimestamp || new Date().toISOString()
    });
    const thread = messagesByWaId.get(waId) || [];
    if (!thread.length) return false;

    let updated = false;
    for (let i = thread.length - 1; i >= 0; i -= 1) {
        const item = thread[i];
        if (item?.id !== messageId) continue;
        item.status = status;
        if (statusTimestamp) item.statusTimestamp = statusTimestamp;
        updated = true;
        break;
    }
    if (updated) {
        messagesByWaId.set(waId, thread);
    }
    return updated;
}

function broadcast(payload) {
    const packet = JSON.stringify(payload);
    sockets.forEach((ws) => {
        if (ws.readyState === ws.OPEN) {
            ws.send(packet);
        }
    });
}

function normalizeIncomingMessage(valueMessage) {
    if (!valueMessage) return '';
    if (valueMessage.type === 'text') return valueMessage.text?.body || '';
    if (valueMessage.type === 'button') return valueMessage.button?.text || '';
    if (valueMessage.type === 'interactive') {
        return valueMessage.interactive?.button_reply?.title || valueMessage.interactive?.list_reply?.title || '[Interactive message]';
    }
    if (valueMessage.type === 'image') return '[Image]';
    if (valueMessage.type === 'document') return '[Document]';
    if (valueMessage.type === 'audio') return '[Audio]';
    if (valueMessage.type === 'video') return '[Video]';
    return '[Unsupported message type]';
}

function getIncomingMessageType(valueMessage) {
    return String(valueMessage?.type || 'text').toLowerCase();
}

function getIncomingAttachmentMeta(valueMessage) {
    const type = getIncomingMessageType(valueMessage);
    if (type === 'image') {
        return {
            messageType: 'image',
            attachmentName: valueMessage.image?.caption || 'Image',
            attachmentUrl: '',
            mimeType: valueMessage.image?.mime_type || 'image/*',
            mediaId: valueMessage.image?.id || ''
        };
    }
    if (type === 'document') {
        return {
            messageType: 'document',
            attachmentName: valueMessage.document?.filename || 'Document',
            attachmentUrl: '',
            mimeType: valueMessage.document?.mime_type || 'application/octet-stream',
            mediaId: valueMessage.document?.id || ''
        };
    }
    if (type === 'audio') {
        return {
            messageType: 'audio',
            attachmentName: 'Audio',
            attachmentUrl: '',
            mimeType: valueMessage.audio?.mime_type || 'audio/*',
            mediaId: valueMessage.audio?.id || ''
        };
    }
    if (type === 'video') {
        return {
            messageType: 'video',
            attachmentName: valueMessage.video?.caption || 'Video',
            attachmentUrl: '',
            mimeType: valueMessage.video?.mime_type || 'video/*',
            mediaId: valueMessage.video?.id || ''
        };
    }
    return {
        messageType: 'text',
        attachmentName: '',
        attachmentUrl: '',
        mimeType: '',
        mediaId: ''
    };
}

function ensureWhatsappConfig() {
    if (!whatsappAccessToken || !whatsappPhoneNumberId) {
        throw new Error('WHATSAPP_ACCESS_TOKEN or WHATSAPP_PHONE_NUMBER_ID missing');
    }
}

async function sendGraphJson(path, payload) {
    const response = await fetch(`https://graph.facebook.com/v22.0/${path}`, {
        method: 'POST',
        headers: {
            Authorization: `Bearer ${whatsappAccessToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
    });
    const result = await response.json();
    if (!response.ok) {
        throw new Error(result?.error?.message || 'WhatsApp API request failed');
    }
    return result;
}

function getGraphErrorCode(error) {
    const message = String(error?.message || '');
    const match = /\(#(\d+)\)/.exec(message);
    return match?.[1] || '';
}

function normalizeTemplateParamKey(value) {
    const key = String(value || '').trim().toLowerCase();
    if (!key || key === 'none') return '';
    if (key === 'contact' || key === 'contactname' || key === 'name') return 'contactName';
    if (key === 'agent' || key === 'agentname') return 'agentName';
    return '';
}

function buildTemplateComponentsFromValues(paramKeys, valuesByKey) {
    if (!Array.isArray(paramKeys) || !paramKeys.length) return [];
    return [{
        type: 'body',
        parameters: paramKeys.map((key) => ({
            type: 'text',
            text: String(valuesByKey[key] || '').trim()
        }))
    }];
}

function getInitiationTemplateParamAttempts(contactName, agentName) {
    const valuesByKey = {
        contactName: String(contactName || '').trim() || 'there',
        agentName: String(agentName || '').trim() || 'our team'
    };

    const candidates = [
        [],
        ['contactName'],
        ['agentName'],
        ['contactName', 'agentName']
    ];

    if (typeof whatsappInitiationTemplateParamOrder === 'string') {
        const configuredKeys = whatsappInitiationTemplateParamOrder
            .split(',')
            .map((item) => normalizeTemplateParamKey(item))
            .filter(Boolean);
        candidates.unshift(configuredKeys);
    }

    const seen = new Set();
    return candidates
        .map((keys) => ({
            label: keys.length ? keys.join(',') : 'none',
            components: buildTemplateComponentsFromValues(keys, valuesByKey)
        }))
        .filter((attempt) => {
            if (seen.has(attempt.label)) return false;
            seen.add(attempt.label);
            return true;
        });
}

function getInitiationTemplateLanguageAttempts() {
    const configured = String(whatsappInitiationTemplateLanguage || '').trim() || 'en_US';
    const attempts = [];
    const push = (code) => {
        const value = String(code || '').trim();
        if (!value || attempts.includes(value)) return;
        attempts.push(value);
    };

    push(configured);
    if (/^en(?:_|$)/i.test(configured)) {
        push('en');
        push('en_US');
    }
    if (/^[a-z]{2}_[A-Z]{2}$/.test(configured)) {
        push(configured.split('_')[0]);
    } else if (/^[a-z]{2}$/i.test(configured)) {
        push(`${configured.toLowerCase()}_US`);
    }
    return attempts;
}

function renderInitiationTemplatePreview(contactName, agentName) {
    return String(whatsappInitiationTemplatePreviewText || '')
        .replace(/\{\{\s*contactName\s*\}\}/gi, String(contactName || '').trim() || 'there')
        .replace(/\{\{\s*agentName\s*\}\}/gi, String(agentName || '').trim() || 'our team')
        .replace(/\{\{\s*1\s*\}\}/g, String(contactName || '').trim() || 'there')
        .replace(/\{\{\s*2\s*\}\}/g, String(agentName || '').trim() || 'our team')
        .trim();
}

async function uploadWhatsappMedia({ fileName, mimeType, dataUrl }) {
    const match = /^data:([^;]+);base64,(.+)$/i.exec(String(dataUrl || ''));
    if (!match?.[1] || !match?.[2]) {
        throw new Error('Attachment must be sent as a valid base64 data URL');
    }
    const resolvedMimeType = mimeType || match[1];
    const buffer = Buffer.from(match[2], 'base64');
    const form = new FormData();
    form.append('messaging_product', 'whatsapp');
    form.append('file', new Blob([buffer], { type: resolvedMimeType }), fileName || 'attachment');

    const response = await fetch(`https://graph.facebook.com/v22.0/${whatsappPhoneNumberId}/media`, {
        method: 'POST',
        headers: {
            Authorization: `Bearer ${whatsappAccessToken}`
        },
        body: form
    });
    const result = await response.json();
    if (!response.ok) {
        throw new Error(result?.error?.message || 'WhatsApp media upload failed');
    }
    return result?.id || '';
}

app.get('/', (_req, res) => {
    res.status(200).send('OK. Try /health (status) or /webhook (Meta callback).');
});

app.get('/health', async (_req, res) => {
    let persistenceReachable = false;
    let persistenceCheckError = '';
    try {
        if (await ensurePersistedStateLoaded(false)) {
            persistenceReachable = true;
        } else if (firebasePersistenceInitError) {
            persistenceCheckError = firebasePersistenceInitError;
        }
    } catch (error) {
        persistenceCheckError = error?.message || 'Failed to reach Firebase persistence';
    }
    return res.json({
        ok: true,
        service: 'whatsapp-webhook',
        at: new Date().toISOString(),
        firebasePersistenceEnabled,
        persistedStateLoaded,
        persistenceReachable,
        firebasePersistenceInitError,
        lastPersistedStateLoadAt,
        lastPersistedStateError,
        lastFirebaseWriteAt,
        lastFirebaseWriteError,
        persistenceCheckError
    });
});

app.get('/webhook', (req, res) => {
    const mode = req.query['hub.mode'];
    const token = req.query['hub.verify_token'];
    const challenge = req.query['hub.challenge'];
    if (mode === 'subscribe' && token && token === verifyToken) {
        return res.status(200).send(challenge);
    }
    return res.sendStatus(403);
});

app.post('/webhook', (req, res) => {
    if (!verifyMetaSignature(req)) {
        console.warn('[webhook] invalid signature');
        return res.sendStatus(403);
    }

    const body = req.body;
    const receivedAt = new Date().toISOString();
    pushWebhookEvent({
        receivedAt,
        object: body?.object || '',
        hasEntry: Array.isArray(body?.entry),
        entryCount: Array.isArray(body?.entry) ? body.entry.length : 0
    });
    console.log('[webhook] received', receivedAt, 'object=', body?.object || '(none)');

    if (body?.object !== 'whatsapp_business_account') {
        // Always ack webhook deliveries; non-WA payloads are ignored.
        return res.sendStatus(200);
    }

    const entries = body.entry || [];
    entries.forEach((entry) => {
        const changes = entry.changes || [];
        changes.forEach((change) => {
            const value = change.value || {};
            const contacts = value.contacts || [];
            const messages = value.messages || [];
            const statuses = value.statuses || [];

            messages.forEach((incomingMessage) => {
                const waId = normalizeWaId(incomingMessage.from || '');
                const contactInfo = contacts.find((c) => c.wa_id === waId);
                const profileName = contactInfo?.profile?.name || waId;
                const timestamp = normalizeWebhookTimestamp(incomingMessage.timestamp, receivedAt);
                const text = normalizeIncomingMessage(incomingMessage);
                const attachmentMeta = getIncomingAttachmentMeta(incomingMessage);
                const replyTo = makeReplyReference(waId, incomingMessage.context?.id || '');

                const contact = upsertContact(waId, profileName);
                const storedMessage = {
                    id: incomingMessage.id || '',
                    direction: 'incoming',
                    text,
                    timestamp,
                    ...attachmentMeta,
                    replyTo
                };
                addMessage(waId, storedMessage);
                void persistContactSnapshot(contact).catch((error) => {
                    console.warn('[storage] Failed to persist contact:', error?.message || error);
                });
                if (storedMessage.mediaId) {
                    void backupIncomingMediaToStorage({
                        waId,
                        messageId: storedMessage.id,
                        timestamp,
                        attachmentName: storedMessage.attachmentName,
                        mediaId: storedMessage.mediaId,
                        mimeType: storedMessage.mimeType
                    }).then((attachmentUrl) => {
                        if (!attachmentUrl) return;
                        storedMessage.attachmentUrl = attachmentUrl;
                        return persistMessageSnapshot(waId, storedMessage);
                    }).catch((error) => {
                        console.warn('[storage] Failed to persist incoming media backup:', error?.message || error);
                    });
                }
                void persistMessageSnapshot(waId, storedMessage).catch((error) => {
                    console.warn('[storage] Failed to persist message:', error?.message || error);
                });
                pushWebhookEvent({
                    receivedAt,
                    object: body?.object || '',
                    waId,
                    profileName,
                    direction: 'incoming',
                    text,
                    timestamp
                });

                broadcast({
                    type: 'whatsapp_message',
                    payload: {
                        waId,
                        profileName,
                        id: incomingMessage.id || '',
                        direction: 'incoming',
                        text,
                        timestamp,
                        ...attachmentMeta,
                        replyTo
                    }
                });
            });

            statuses.forEach((statusEvent) => {
                const waId = normalizeWaId(statusEvent.recipient_id || '');
                const messageId = statusEvent.id || '';
                const status = String(statusEvent.status || '').toLowerCase();
                const statusTimestamp = normalizeWebhookTimestamp(statusEvent.timestamp, receivedAt);
                if (!waId || !messageId || !status) return;

                updateMessageStatus(waId, messageId, status, statusTimestamp);
                void persistMessageStatusSnapshot(waId, messageId, status, statusTimestamp).catch((error) => {
                    console.warn('[storage] Failed to persist message status:', error?.message || error);
                });
                pushWebhookEvent({
                    receivedAt,
                    object: body?.object || '',
                    waId,
                    direction: 'outgoing',
                    messageId,
                    status,
                    statusTimestamp
                });
                broadcast({
                    type: 'whatsapp_message_status',
                    payload: {
                        waId,
                        messageId,
                        status,
                        statusTimestamp
                    }
                });
            });
        });
    });

    return res.sendStatus(200);
});

app.get('/api/whatsapp/debug', (_req, res) => {
    const persistence = {
        enabled: firebasePersistenceEnabled,
        loaded: persistedStateLoaded,
        initError: firebasePersistenceInitError,
        lastLoadAt: lastPersistedStateLoadAt,
        lastLoadError: lastPersistedStateError,
        lastWriteAt: lastFirebaseWriteAt,
        lastWriteError: lastFirebaseWriteError
    };
    const recentThreads = Array.from(messagesByWaId.entries())
        .map(([waId, rows]) => {
            const messages = Array.isArray(rows) ? rows : [];
            const lastMessage = messages[messages.length - 1] || null;
            return {
                waId,
                count: messages.length,
                lastMessage: lastMessage ? {
                    id: lastMessage.id || '',
                    direction: lastMessage.direction || '',
                    text: lastMessage.text || '',
                    timestamp: lastMessage.timestamp || '',
                    status: lastMessage.status || '',
                    statusTimestamp: lastMessage.statusTimestamp || ''
                } : null
            };
        })
        .sort((a, b) => {
            const aTs = new Date(a.lastMessage?.timestamp || 0).getTime();
            const bTs = new Date(b.lastMessage?.timestamp || 0).getTime();
            return bTs - aTs;
        })
        .slice(0, 20);

    res.json({
        ok: true,
        persistence,
        crmState: {
            loaded: crmStateLoaded,
            updatedAt: crmState.updatedAt,
            manualContacts: crmState.manualContacts.length,
            hasOrderListHtml: Boolean(crmState.orderListHtml),
            contactIdMapSize: Object.keys(crmState.whatsappContactIdMap || {}).length,
            readStateSize: Object.keys(crmState.whatsappReadState || {}).length,
            contactIdSequence: crmState.whatsappContactIdSequence
        },
        contacts: contactsByWaId.size,
        messageThreads: messagesByWaId.size,
        recentWebhookEvents: webhookEvents.slice(0, 20),
        recentApiEvents: apiEvents.slice(0, 20),
        recentThreads
    });
});

app.get('/api/whatsapp/contacts', async (_req, res) => {
    try {
        await ensurePersistedStateLoaded(false);
    } catch (error) {
        console.warn('[storage] Failed to refresh persisted chat state before contacts response:', error?.message || error);
    }
    const waIds = new Set([
        ...Array.from(contactsByWaId.keys()),
        ...Array.from(messagesByWaId.keys())
    ]);
    const rows = Array.from(waIds)
        .map((waId) => ensureContactForThread(waId))
        .filter(Boolean)
        .sort((a, b) => {
            const aLastMessage = (messagesByWaId.get(a.waId) || []).slice(-1)[0];
            const bLastMessage = (messagesByWaId.get(b.waId) || []).slice(-1)[0];
            const aTs = Math.max(
                new Date(aLastMessage?.timestamp || 0).getTime() || 0,
                new Date(a.updatedAt || 0).getTime() || 0
            );
            const bTs = Math.max(
                new Date(bLastMessage?.timestamp || 0).getTime() || 0,
                new Date(b.updatedAt || 0).getTime() || 0
            );
            return bTs - aTs;
        })
        .map((contact) => ({
            waId: contact.waId,
            profileName: contact.profileName,
            updatedAt: contact.updatedAt,
            messages: (messagesByWaId.get(contact.waId) || []).map((message) => {
                const knownStatus = messageStatusById.get(message.id || '');
                if (!knownStatus) return message;
                return {
                    ...message,
                    status: knownStatus.status,
                    statusTimestamp: knownStatus.statusTimestamp
                };
            })
        }));
    res.json(rows);
});

app.get('/api/crm/state', async (_req, res) => {
    try {
        await ensureCrmStateLoaded(false);
    } catch (error) {
        console.warn('[storage] Failed to load CRM state:', error?.message || error);
    }
    res.json({
        ok: true,
        orderListHtml: crmState.orderListHtml,
        manualContacts: crmState.manualContacts,
        whatsappContactIdMap: crmState.whatsappContactIdMap,
        whatsappReadState: crmState.whatsappReadState,
        whatsappContactIdSequence: crmState.whatsappContactIdSequence,
        updatedAt: crmState.updatedAt
    });
});

app.post('/api/crm/state', async (req, res) => {
    try {
        await persistCrmState(req.body || {});
        return res.json({
            ok: true,
            orderListHtml: crmState.orderListHtml,
            manualContacts: crmState.manualContacts,
            whatsappContactIdMap: crmState.whatsappContactIdMap,
            whatsappReadState: crmState.whatsappReadState,
            whatsappContactIdSequence: crmState.whatsappContactIdSequence,
            updatedAt: crmState.updatedAt
        });
    } catch (error) {
        lastFirebaseWriteError = error?.message || 'Failed to persist CRM state';
        return res.status(500).json({ ok: false, error: error?.message || 'Failed to persist CRM state' });
    }
});

app.post('/api/whatsapp/send', async (req, res) => {
    const waId = normalizeWaId(req.body?.waId || '');
    const text = String(req.body?.text || '').trim();
    const contextMessageId = String(req.body?.contextMessageId || '').trim();

    if (!waId || !text) {
        return res.status(400).json({ ok: false, error: 'waId and text are required' });
    }

    try {
        ensureWhatsappConfig();
        pushApiEvent({
            at: new Date().toISOString(),
            route: 'send',
            waId,
            type: 'text',
            text,
            contextMessageId
        });
        const result = await sendGraphJson(`${whatsappPhoneNumberId}/messages`, {
            messaging_product: 'whatsapp',
            to: waId,
            type: 'text',
            text: { body: text },
            ...(contextMessageId ? { context: { message_id: contextMessageId } } : {})
        });

        const profileName = contactsByWaId.get(waId)?.profileName || waId;
        const timestamp = new Date().toISOString();
        const messageId = result?.messages?.[0]?.id || '';
        const replyTo = makeReplyReference(waId, contextMessageId);
        const contact = upsertContact(waId, profileName);
        const storedMessage = {
            id: messageId,
            direction: 'outgoing',
            text,
            timestamp,
            replyTo,
            status: 'sent',
            statusTimestamp: timestamp
        };
        addMessage(waId, storedMessage);
        await persistContactSnapshot(contact);
        await persistMessageSnapshot(waId, storedMessage);
        pushApiEvent({
            at: timestamp,
            route: 'send',
            waId,
            type: 'text',
            ok: true,
            messageId
        });

        broadcast({
            type: 'whatsapp_message',
            payload: {
                waId,
                profileName,
                id: messageId,
                direction: 'outgoing',
                text,
                timestamp,
                replyTo,
                status: 'sent',
                statusTimestamp: timestamp
            }
        });

        return res.json({ ok: true, id: messageId });
    } catch (error) {
        pushApiEvent({
            at: new Date().toISOString(),
            route: 'send',
            waId,
            type: 'text',
            ok: false,
            error: error?.message || 'Unexpected send error'
        });
        return res.status(500).json({ ok: false, error: error?.message || 'Unexpected send error' });
    }
});

app.post('/api/whatsapp/send-media', async (req, res) => {
    const waId = normalizeWaId(req.body?.waId || '');
    const fileName = String(req.body?.fileName || '').trim();
    const mimeType = String(req.body?.mimeType || '').trim();
    const dataUrl = String(req.body?.dataUrl || '');
    const requestedType = String(req.body?.messageType || '').trim().toLowerCase();
    const caption = String(req.body?.caption || '').trim();
    const contextMessageId = String(req.body?.contextMessageId || '').trim();

    if (!waId || !fileName || !dataUrl) {
        return res.status(400).json({ ok: false, error: 'waId, fileName and dataUrl are required' });
    }

    try {
        ensureWhatsappConfig();
        const messageType = ['image', 'audio', 'video', 'document'].includes(requestedType) ? requestedType : 'document';
        pushApiEvent({
            at: new Date().toISOString(),
            route: 'send-media',
            waId,
            type: messageType,
            fileName,
            caption,
            contextMessageId
        });
        const mediaId = await uploadWhatsappMedia({ fileName, mimeType, dataUrl });
        const mediaPayload = {
            id: mediaId
        };
        if (caption && (messageType === 'image' || messageType === 'video' || messageType === 'document')) {
            mediaPayload.caption = caption;
            if (messageType === 'document') {
                mediaPayload.filename = fileName;
            }
        }
        const result = await sendGraphJson(`${whatsappPhoneNumberId}/messages`, {
            messaging_product: 'whatsapp',
            to: waId,
            type: messageType,
            [messageType]: mediaPayload,
            ...(contextMessageId ? { context: { message_id: contextMessageId } } : {})
        });

        const profileName = contactsByWaId.get(waId)?.profileName || waId;
        const timestamp = new Date().toISOString();
        const messageId = result?.messages?.[0]?.id || '';
        const text = caption || `[${messageType.charAt(0).toUpperCase() + messageType.slice(1)}] ${fileName}`;
        const replyTo = makeReplyReference(waId, contextMessageId);
        let attachmentUrl = '';
        try {
            const parsed = parseDataUrl(dataUrl);
            attachmentUrl = await uploadMediaBufferToFirebaseStorage({
                waId,
                messageId,
                timestamp,
                fileName,
                mimeType: mimeType || parsed.mimeType,
                buffer: parsed.buffer
            });
        } catch (storageError) {
            console.warn('[storage] Failed to back up outgoing media:', storageError?.message || storageError);
        }
        const contact = upsertContact(waId, profileName);
        const storedMessage = {
            id: messageId,
            direction: 'outgoing',
            text,
            timestamp,
            messageType,
            attachmentName: fileName,
            attachmentUrl,
            mediaId,
            mimeType: mimeType || '',
            replyTo,
            status: 'sent',
            statusTimestamp: timestamp
        };
        addMessage(waId, storedMessage);
        await persistContactSnapshot(contact);
        await persistMessageSnapshot(waId, storedMessage);
        pushApiEvent({
            at: timestamp,
            route: 'send-media',
            waId,
            type: messageType,
            ok: true,
            messageId,
            mediaId,
            fileName
        });

        broadcast({
            type: 'whatsapp_message',
            payload: {
                waId,
                profileName,
                id: messageId,
                direction: 'outgoing',
                text,
                timestamp,
                messageType,
                attachmentName: fileName,
                attachmentUrl,
                mediaId,
                mimeType: mimeType || '',
                replyTo,
                status: 'sent',
                statusTimestamp: timestamp
            }
        });

        return res.json({
            ok: true,
            id: messageId,
            text,
            messageType,
            attachmentName: fileName,
            attachmentUrl,
            mediaId,
            mimeType: mimeType || ''
        });
    } catch (error) {
        pushApiEvent({
            at: new Date().toISOString(),
            route: 'send-media',
            waId,
            type: requestedType || 'document',
            ok: false,
            fileName,
            error: error?.message || 'Unexpected media send error'
        });
        return res.status(500).json({ ok: false, error: error?.message || 'Unexpected media send error' });
    }
});

app.post('/api/whatsapp/initiate', async (req, res) => {
    const waId = normalizeWaId(req.body?.waId || '');
    const contactName = String(req.body?.contactName || '').trim() || 'there';
    const agentName = String(req.body?.agentName || '').trim() || 'our team';

    if (!waId) {
        return res.status(400).json({ ok: false, error: 'waId is required' });
    }
    if (!whatsappInitiationTemplateName) {
        return res.status(400).json({ ok: false, error: 'WHATSAPP_INIT_TEMPLATE_NAME is not configured on the backend' });
    }

    try {
        ensureWhatsappConfig();
        const previewText = renderInitiationTemplatePreview(contactName, agentName);
        pushApiEvent({
            at: new Date().toISOString(),
            route: 'initiate',
            waId,
            type: 'template',
            templateName: whatsappInitiationTemplateName,
            templateLanguage: whatsappInitiationTemplateLanguage
        });
        const attempts = getInitiationTemplateParamAttempts(contactName, agentName);
        const languageAttempts = getInitiationTemplateLanguageAttempts();
        let result = null;
        let lastError = null;
        let resolvedTemplateLanguage = whatsappInitiationTemplateLanguage;
        for (const languageCode of languageAttempts) {
            for (const attempt of attempts) {
                try {
                    result = await sendGraphJson(`${whatsappPhoneNumberId}/messages`, {
                        messaging_product: 'whatsapp',
                        to: waId,
                        type: 'template',
                        template: {
                            name: whatsappInitiationTemplateName,
                            language: {
                                code: languageCode
                            },
                            ...(attempt.components.length ? { components: attempt.components } : {})
                        }
                    });
                    resolvedTemplateLanguage = languageCode;
                    lastError = null;
                    break;
                } catch (error) {
                    lastError = error;
                    const graphErrorCode = getGraphErrorCode(error);
                    const languageMismatch = graphErrorCode === '132001' || /language/i.test(String(error?.message || ''));
                    if (graphErrorCode !== '132000' && !languageMismatch) {
                        break;
                    }
                }
            }
            if (result) break;
        }
        if (!result) {
            if (lastError && getGraphErrorCode(lastError) === '132000' && typeof whatsappInitiationTemplateParamOrder !== 'string') {
                throw new Error('Template parameter mismatch. Set WHATSAPP_INIT_TEMPLATE_PARAM_ORDER to match your approved template, e.g. "none", "contactName", "agentName", or "contactName,agentName". Original Meta error: ' + lastError.message);
            }
            throw lastError || new Error('Chat initiation failed');
        }

        const profileName = contactsByWaId.get(waId)?.profileName || waId;
        const timestamp = new Date().toISOString();
        const messageId = result?.messages?.[0]?.id || '';
        const text = previewText || `Chat initiation template sent (${whatsappInitiationTemplateName})`;
        const contact = upsertContact(waId, profileName);
        const storedMessage = {
            id: messageId,
            direction: 'outgoing',
            text,
            timestamp,
            messageType: 'template',
            status: 'sent',
            statusTimestamp: timestamp
        };
        addMessage(waId, storedMessage);
        await persistContactSnapshot(contact);
        await persistMessageSnapshot(waId, storedMessage);
        pushApiEvent({
            at: timestamp,
            route: 'initiate',
            waId,
            type: 'template',
            ok: true,
            messageId,
            templateName: whatsappInitiationTemplateName,
            templateLanguage: resolvedTemplateLanguage
        });

        broadcast({
            type: 'whatsapp_message',
            payload: {
                waId,
                profileName,
                id: messageId,
                direction: 'outgoing',
                text,
                timestamp,
                messageType: 'template',
                status: 'sent',
                statusTimestamp: timestamp
            }
        });

        return res.json({ ok: true, id: messageId, text, previewText, templateName: whatsappInitiationTemplateName, templateLanguage: resolvedTemplateLanguage });
    } catch (error) {
        pushApiEvent({
            at: new Date().toISOString(),
            route: 'initiate',
            waId,
            type: 'template',
            ok: false,
            templateName: whatsappInitiationTemplateName,
            templateLanguage: whatsappInitiationTemplateLanguage,
            error: error?.message || 'Unexpected initiation error'
        });
        return res.status(500).json({ ok: false, error: error?.message || 'Unexpected initiation error' });
    }
});

app.post('/api/whatsapp/initiate-preview', (req, res) => {
    const contactName = String(req.body?.contactName || '').trim() || 'there';
    const agentName = String(req.body?.agentName || '').trim() || 'our team';
    return res.json({
        ok: true,
        templateName: whatsappInitiationTemplateName,
        templateLanguage: whatsappInitiationTemplateLanguage,
        previewText: renderInitiationTemplatePreview(contactName, agentName)
    });
});

app.post('/api/whatsapp/forward', async (req, res) => {
    const targetWaId = normalizeWaId(req.body?.targetWaId || '');
    const messages = Array.isArray(req.body?.messages) ? req.body.messages : [];

    if (!targetWaId || !messages.length) {
        return res.status(400).json({ ok: false, error: 'targetWaId and messages are required' });
    }

    try {
        ensureWhatsappConfig();
        const profileName = contactsByWaId.get(targetWaId)?.profileName || targetWaId;
        const forwardedIds = [];

        for (const message of messages) {
            const messageType = String(message?.messageType || 'text');
            const text = String(message?.text || '').trim();
            if (messageType === 'text' || messageType === 'template') {
                const result = await sendGraphJson(`${whatsappPhoneNumberId}/messages`, {
                    messaging_product: 'whatsapp',
                    to: targetWaId,
                    type: 'text',
                    text: { body: text || '(forwarded message)' }
                });
                const messageId = result?.messages?.[0]?.id || '';
                const timestamp = new Date().toISOString();
                const storedMessage = {
                    id: messageId,
                    direction: 'outgoing',
                    text: text || '(forwarded message)',
                    timestamp,
                    messageType: 'text',
                    status: 'sent',
                    statusTimestamp: timestamp
                };
                addMessage(targetWaId, storedMessage);
                await persistMessageSnapshot(targetWaId, storedMessage);
                forwardedIds.push(messageId);
                continue;
            }

            const mediaId = String(message?.mediaId || '').trim();
            if (!mediaId) {
                throw new Error(`Cannot forward "${message?.attachmentName || 'attachment'}" because mediaId is missing`);
            }
            const payload = { id: mediaId };
            if (text && ['image', 'video', 'document'].includes(messageType)) {
                payload.caption = text;
            }
            if (messageType === 'document' && message?.attachmentName) {
                payload.filename = String(message.attachmentName);
            }
            const result = await sendGraphJson(`${whatsappPhoneNumberId}/messages`, {
                messaging_product: 'whatsapp',
                to: targetWaId,
                type: messageType,
                [messageType]: payload
            });
            const messageId = result?.messages?.[0]?.id || '';
            const timestamp = new Date().toISOString();
            const storedMessage = {
                id: messageId,
                direction: 'outgoing',
                text: text || `[${messageType}] ${message?.attachmentName || ''}`.trim(),
                timestamp,
                messageType,
                attachmentName: String(message?.attachmentName || ''),
                attachmentUrl: String(message?.attachmentUrl || ''),
                mediaId,
                mimeType: String(message?.mimeType || ''),
                status: 'sent',
                statusTimestamp: timestamp
            };
            addMessage(targetWaId, storedMessage);
            await persistMessageSnapshot(targetWaId, storedMessage);
            forwardedIds.push(messageId);
        }

        const contact = upsertContact(targetWaId, profileName);
        await persistContactSnapshot(contact);
        return res.json({ ok: true, forwarded: forwardedIds.length, ids: forwardedIds });
    } catch (error) {
        return res.status(500).json({ ok: false, error: error?.message || 'Unexpected forward error' });
    }
});

app.use((error, _req, res, next) => {
    if (error?.type === 'entity.parse.failed') {
        return res.status(400).json({
            ok: false,
            error: 'Invalid JSON body',
            hint: 'Send a valid JSON payload. In PowerShell, prefer Invoke-RestMethod or pass curl data from a file to avoid quote escaping issues.'
        });
    }
    return next(error);
});

try {
    if (initFirebasePersistence()) {
        await ensurePersistedStateLoaded(true);
        await ensureCrmStateLoaded(true);
        console.log('[storage] Firebase persistence enabled.');
        if (firebaseBucket) {
            console.log('[storage] Media backup bucket:', firebaseBucket.name);
        }
    } else if (firebasePersistenceInitError) {
        console.warn('[storage] Firebase persistence disabled:', firebasePersistenceInitError);
    } else {
        console.warn('[storage] Firebase persistence not configured; using in-memory storage only.');
    }
} catch (error) {
    console.warn('[storage] Failed to load persisted chat state:', error?.message || error);
}

const server = app.listen(port, () => {
    console.log(`WhatsApp webhook server running on http://localhost:${port}`);
});

const wss = new WebSocketServer({ server, path: '/ws' });
wss.on('connection', (ws) => {
    sockets.add(ws);
    ws.on('close', () => sockets.delete(ws));
});
