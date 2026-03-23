import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import { WebSocketServer } from 'ws';
import crypto from 'node:crypto';

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

let loggedMissingAppSecret = false;

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
    if (existing.length > 250) {
        existing.splice(0, existing.length - 250);
    }
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

    if (typeof whatsappInitiationTemplateParamOrder === 'string') {
        const configuredKeys = whatsappInitiationTemplateParamOrder
            .split(',')
            .map((item) => normalizeTemplateParamKey(item))
            .filter(Boolean);
        return [{
            label: configuredKeys.length ? configuredKeys.join(',') : 'none',
            components: buildTemplateComponentsFromValues(configuredKeys, valuesByKey)
        }];
    }

    const candidates = [
        [],
        ['contactName'],
        ['agentName'],
        ['contactName', 'agentName']
    ];
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

app.get('/health', (_req, res) => {
    res.json({ ok: true, service: 'whatsapp-webhook', at: new Date().toISOString() });
});

app.get('/', (_req, res) => {
    res.status(200).send('OK. Try /health (status) or /webhook (Meta callback).');
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
                const tsMillis = Number(incomingMessage.timestamp || '0') * 1000;
                const timestamp = tsMillis > 0 ? new Date(tsMillis).toISOString() : new Date().toISOString();
                const text = normalizeIncomingMessage(incomingMessage);
                const attachmentMeta = getIncomingAttachmentMeta(incomingMessage);
                const replyTo = makeReplyReference(waId, incomingMessage.context?.id || '');

                upsertContact(waId, profileName);
                addMessage(waId, {
                    id: incomingMessage.id || '',
                    direction: 'incoming',
                    text,
                    timestamp,
                    ...attachmentMeta,
                    replyTo
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
                const tsMillis = Number(statusEvent.timestamp || '0') * 1000;
                const statusTimestamp = tsMillis > 0 ? new Date(tsMillis).toISOString() : new Date().toISOString();
                if (!waId || !messageId || !status) return;

                updateMessageStatus(waId, messageId, status, statusTimestamp);
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
        contacts: contactsByWaId.size,
        messageThreads: messagesByWaId.size,
        recentWebhookEvents: webhookEvents.slice(0, 20),
        recentApiEvents: apiEvents.slice(0, 20),
        recentThreads
    });
});

app.get('/api/whatsapp/contacts', (_req, res) => {
    const rows = Array.from(contactsByWaId.values())
        .sort((a, b) => {
            const aTs = new Date(a.updatedAt).getTime();
            const bTs = new Date(b.updatedAt).getTime();
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
        upsertContact(waId, profileName);
        addMessage(waId, {
            id: messageId,
            direction: 'outgoing',
            text,
            timestamp,
            replyTo,
            status: 'sent',
            statusTimestamp: timestamp
        });
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
        upsertContact(waId, profileName);
        addMessage(waId, {
            id: messageId,
            direction: 'outgoing',
            text,
            timestamp,
            messageType,
            attachmentName: fileName,
            attachmentUrl: '',
            mediaId,
            mimeType: mimeType || '',
            replyTo,
            status: 'sent',
            statusTimestamp: timestamp
        });
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
                attachmentUrl: '',
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
            attachmentUrl: '',
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
        let result = null;
        let lastError = null;
        for (const attempt of attempts) {
            try {
                result = await sendGraphJson(`${whatsappPhoneNumberId}/messages`, {
                    messaging_product: 'whatsapp',
                    to: waId,
                    type: 'template',
                    template: {
                        name: whatsappInitiationTemplateName,
                        language: {
                            code: whatsappInitiationTemplateLanguage
                        },
                        ...(attempt.components.length ? { components: attempt.components } : {})
                    }
                });
                lastError = null;
                break;
            } catch (error) {
                lastError = error;
                if (typeof whatsappInitiationTemplateParamOrder === 'string') {
                    break;
                }
                if (getGraphErrorCode(error) !== '132000') {
                    break;
                }
            }
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
        upsertContact(waId, profileName);
        addMessage(waId, {
            id: messageId,
            direction: 'outgoing',
            text,
            timestamp,
            messageType: 'template',
            status: 'sent',
            statusTimestamp: timestamp
        });
        pushApiEvent({
            at: timestamp,
            route: 'initiate',
            waId,
            type: 'template',
            ok: true,
            messageId,
            templateName: whatsappInitiationTemplateName,
            templateLanguage: whatsappInitiationTemplateLanguage
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

        return res.json({ ok: true, id: messageId, text, previewText, templateName: whatsappInitiationTemplateName });
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
                addMessage(targetWaId, {
                    id: messageId,
                    direction: 'outgoing',
                    text: text || '(forwarded message)',
                    timestamp,
                    messageType: 'text',
                    status: 'sent',
                    statusTimestamp: timestamp
                });
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
            addMessage(targetWaId, {
                id: messageId,
                direction: 'outgoing',
                text: text || `[${messageType}] ${message?.attachmentName || ''}`.trim(),
                timestamp,
                messageType,
                attachmentName: String(message?.attachmentName || ''),
                attachmentUrl: '',
                mediaId,
                mimeType: String(message?.mimeType || ''),
                status: 'sent',
                statusTimestamp: timestamp
            });
            forwardedIds.push(messageId);
        }

        upsertContact(targetWaId, profileName);
        return res.json({ ok: true, forwarded: forwardedIds.length, ids: forwardedIds });
    } catch (error) {
        return res.status(500).json({ ok: false, error: error?.message || 'Unexpected forward error' });
    }
});

const server = app.listen(port, () => {
    console.log(`WhatsApp webhook server running on http://localhost:${port}`);
});

const wss = new WebSocketServer({ server, path: '/ws' });
wss.on('connection', (ws) => {
    sockets.add(ws);
    ws.on('close', () => sockets.delete(ws));
});
