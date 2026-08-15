warning: in the working copy of 'server.mjs', LF will be replaced by CRLF the next time Git touches it
[1mdiff --git a/server.mjs b/server.mjs[m
[1mindex df73066..87f86da 100644[m
[1m--- a/server.mjs[m
[1m+++ b/server.mjs[m
[36m@@ -343,6 +343,38 @@[m [mfunction mergePersistedMessageIntoMemory(waId, message) {[m
     }[m
 }[m
 [m
[32m+[m[32mfunction ensureContactForThread(waId, fallbackProfileName = '') {[m
[32m+[m[32m    const normalizedWaId = normalizeWaId(waId);[m
[32m+[m[32m    if (!normalizedWaId) return null;[m
[32m+[m[32m    const existing = contactsByWaId.get(normalizedWaId);[m
[32m+[m[32m    const thread = messagesByWaId.get(normalizedWaId) || [];[m
[32m+[m[32m    const lastMessage = thread[thread.length - 1] || null;[m
[32m+[m[32m    const candidateUpdatedAt = String([m
[32m+[m[32m        existing?.updatedAt ||[m
[32m+[m[32m        lastMessage?.timestamp ||[m
[32m+[m[32m        new Date().toISOString()[m
[32m+[m[32m    );[m
[32m+[m[32m    const nextContact = {[m
[32m+[m[32m        waId: normalizedWaId,[m
[32m+[m[32m        profileName: String(existing?.profileName || fallbackProfileName || normalizedWaId),[m
[32m+[m[32m        updatedAt: candidateUpdatedAt[m
[32m+[m[32m    };[m
[32m+[m[32m    if (!existing) {[m
[32m+[m[32m        contactsByWaId.set(normalizedWaId, nextContact);[m
[32m+[m[32m        return nextContact;[m
[32m+[m[32m    }[m
[32m+[m[32m    const existingUpdatedAtMs = new Date(existing.updatedAt || 0).getTime() || 0;[m
[32m+[m[32m    const candidateUpdatedAtMs = new Date(candidateUpdatedAt).getTime() || 0;[m
[32m+[m[32m    if (candidateUpdatedAtMs >= existingUpdatedAtMs) {[m
[32m+[m[32m        existing.updatedAt = candidateUpdatedAt;[m
[32m+[m[32m    }[m
[32m+[m[32m    if (!String(existing.profileName || '').trim() && nextContact.profileName) {[m
[32m+[m[32m        existing.profileName = nextContact.profileName;[m
[32m+[m[32m    }[m
[32m+[m[32m    contactsByWaId.set(normalizedWaId, existing);[m
[32m+[m[32m    return existing;[m
[32m+[m[32m}[m
[32m+[m
 async function loadPersistedStateIntoMemory() {[m
     if (!initFirebasePersistence()) return false;[m
     const contactsSnapshot = await firebaseDb.collection('wa_contacts').get();[m
[36m@@ -368,6 +400,9 @@[m [masync function loadPersistedStateIntoMemory() {[m
         if (!waId) return;[m
         mergePersistedMessageIntoMemory(waId, data);[m
     });[m
[32m+[m[32m    Array.from(messagesByWaId.keys()).forEach((waId) => {[m
[32m+[m[32m        ensureContactForThread(waId);[m
[32m+[m[32m    });[m
     persistedStateLoaded = true;[m
     lastPersistedStateLoadAt = new Date().toISOString();[m
     lastPersistedStateError = '';[m
[36m@@ -665,23 +700,21 @@[m [mfunction getInitiationTemplateParamAttempts(contactName, agentName) {[m
         agentName: String(agentName || '').trim() || 'our team'[m
     };[m
 [m
[32m+[m[32m    const candidates = [[m
[32m+[m[32m        [],[m
[32m+[m[32m        ['contactName'],[m
[32m+[m[32m        ['agentName'],[m
[32m+[m[32m        ['contactName', 'agentName'][m
[32m+[m[32m    ];[m
[32m+[m
     if (typeof whatsappInitiationTemplateParamOrder === 'string') {[m
         const configuredKeys = whatsappInitiationTemplateParamOrder[m
             .split(',')[m
             .map((item) => normalizeTemplateParamKey(item))[m
             .filter(Boolean);[m
[31m-        return [{[m
[31m-            label: configuredKeys.length ? configuredKeys.join(',') : 'none',[m
[31m-            components: buildTemplateComponentsFromValues(configuredKeys, valuesByKey)[m
[31m-        }];[m
[32m+[m[32m        candidates.unshift(configuredKeys);[m
     }[m
 [m
[31m-    const candidates = [[m
[31m-        [],[m
[31m-        ['contactName'],[m
[31m-        ['agentName'],[m
[31m-        ['contactName', 'agentName'][m
[31m-    ];[m
     const seen = new Set();[m
     return candidates[m
         .map((keys) => ({[m
[36m@@ -695,6 +728,28 @@[m [mfunction getInitiationTemplateParamAttempts(contactName, agentName) {[m
         });[m
 }[m
 [m
[32m+[m[32mfunction getInitiationTemplateLanguageAttempts() {[m
[32m+[m[32m    const configured = String(whatsappInitiationTemplateLanguage || '').trim() || 'en_US';[m
[32m+[m[32m    const attempts = [];[m
[32m+[m[32m    const push = (code) => {[m
[32m+[m[32m        const value = String(code || '').trim();[m
[32m+[m[32m        if (!value || attempts.includes(value)) return;[m
[32m+[m[32m        attempts.push(value);[m
[32m+[m[32m    };[m
[32m+[m
[32m+[m[32m    push(configured);[m
[32m+[m[32m    if (/^en(?:_|$)/i.test(configured)) {[m
[32m+[m[32m        push('en');[m
[32m+[m[32m        push('en_US');[m
[32m+[m[32m    }[m
[32m+[m[32m    if (/^[a-z]{2}_[A-Z]{2}$/.test(configured)) {[m
[32m+[m[32m        push(configured.split('_')[0]);[m
[32m+[m[32m    } else if (/^[a-z]{2}$/i.test(configured)) {[m
[32m+[m[32m        push(`${configured.toLowerCase()}_US`);[m
[32m+[m[32m    }[m
[32m+[m[32m    return attempts;[m
[32m+[m[32m}[m
[32m+[m
 function renderInitiationTemplatePreview(contactName, agentName) {[m
     return String(whatsappInitiationTemplatePreviewText || '')[m
         .replace(/\{\{\s*contactName\s*\}\}/gi, String(contactName || '').trim() || 'there')[m
[36m@@ -954,10 +1009,24 @@[m [mapp.get('/api/whatsapp/contacts', async (_req, res) => {[m
     } catch (error) {[m
         console.warn('[storage] Failed to refresh persisted chat state before contacts response:', error?.message || error);[m
     }[m
[31m-    const rows = Array.from(contactsByWaId.values())[m
[32m+[m[32m    const waIds = new Set([[m
[32m+[m[32m        ...Array.from(contactsByWaId.keys()),[m
[32m+[m[32m        ...Array.from(messagesByWaId.keys())[m
[32m+[m[32m    ]);[m
[32m+[m[32m    const rows = Array.from(waIds)[m
[32m+[m[32m        .map((waId) => ensureContactForThread(waId))[m
[32m+[m[32m        .filter(Boolean)[m
         .sort((a, b) => {[m
[31m-            const aTs = new Date(a.updatedAt).getTime();[m
[31m-            const bTs = new Date(b.updatedAt).getTime();[m
[32m+[m[32m            const aLastMessage = (messagesByWaId.get(a.waId) || []).slice(-1)[0];[m
[32m+[m[32m            const bLastMessage = (messagesByWaId.get(b.waId) || []).slice(-1)[0];[m
[32m+[m[32m            const aTs = Math.max([m
[32m+[m[32m                new Date(aLastMessage?.timestamp || 0).getTime() || 0,[m
[32m+[m[32m                new Date(a.updatedAt || 0).getTime() || 0[m
[32m+[m[32m            );[m
[32m+[m[32m            const bTs = Math.max([m
[32m+[m[32m                new Date(bLastMessage?.timestamp || 0).getTime() || 0,[m
[32m+[m[32m                new Date(b.updatedAt || 0).getTime() || 0[m
[32m+[m[32m            );[m
             return bTs - aTs;[m
         })[m
         .map((contact) => ({[m
[36m@@ -1218,33 +1287,38 @@[m [mapp.post('/api/whatsapp/initiate', async (req, res) => {[m
             templateLanguage: whatsappInitiationTemplateLanguage[m
         });[m
         const attempts = getInitiationTemplateParamAttempts(contactName, agentName);[m
[32m+[m[32m        const languageAttempts = getInitiationTemplateLanguageAttempts();[m
         let result = null;[m
         let lastError = null;[m
[31m-        for (const attempt of attempts) {[m
[31m-            try {[m
[31m-                result = await sendGraphJson(`${whatsappPhoneNumberId}/messages`, {[m
[31m-                    messaging_product: 'whatsapp',[m
[31m-                    to: waId,[m
[31m-                    type: 'template',[m
[31m-                    template: {[m
[31m-                        name: whatsappInitiationTemplateName,[m
[31m-                        language: {[m
[31m-                            code: whatsappInitiationTemplateLanguage[m
[31m-                        },[m
[31m-                        ...(attempt.components.length ? { components: attempt.components } : {})[m
[31m-                    }[m
[31m-                });[m
[31m-                lastError = null;[m
[31m-                break;[m
[31m-            } catch (error) {[m
[31m-                lastError = error;[m
[31m-                if (typeof whatsappInitiationTemplateParamOrder === 'string') {[m
[31m-                    break;[m
[31m-                }[m
[31m-                if (getGraphErrorCode(error) !== '132000') {[m
[32m+[m[32m        let resolvedTemplateLanguage = whatsappInitiationTemplateLanguage;[m
[32m+[m[32m        for (const languageCode of languageAttempts) {[m
[32m+[m[32m            for (const attempt of attempts) {[m
[32m+[m[32m                try {[m
[32m+[m[32m                    result = await sendGraphJson(`${whatsappPhoneNumberId}/messages`, {[m
[32m+[m[32m                        messaging_product: 'whatsapp',[m
[32m+[m[32m                        to: waId,[m
[32m+[m[32m                        type: 'template',[m
[32m+[m[32m                        template: {[m
[32m+[m[32m                            name: whatsappInitiationTemplateName,[m
[32m+[m[32m                            language: {[m
[32m+[m[32m                                code: languageCode[m
[32m+[m[32m                            },[m
[32m+[m[32m                            ...(attempt.components.length ? { components: attempt.components } : {})[m
[32m+[m[32m                        }[m
[32m+[m[32m                    });[m
[32m+[m[32m                    resolvedTemplateLanguage = languageCode;[m
[32m+[m[32m                    lastError = null;[m
                     break;[m
[32m+[m[32m                } catch (error) {[m
[32m+[m[32m                    lastError = error;[m
[32m+[m[32m                    const graphErrorCode = getGraphErrorCode(error);[m
[32m+[m[32m                    const languageMismatch = graphErrorCode === '132001' || /language/i.test(String(error?.message || ''));[m
[32m+[m[32m                    if (graphErrorCode !== '132000' && !languageMismatch) {[m
[32m+[m[32m                        break;[m
[32m+[m[32m                    }[m
                 }[m
             }[m
[32m+[m[32m            if (result) break;[m
         }[m
         if (!result) {[m
             if (lastError && getGraphErrorCode(lastError) === '132000' && typeof whatsappInitiationTemplateParamOrder !== 'string') {[m
[36m@@ -1278,7 +1352,7 @@[m [mapp.post('/api/whatsapp/initiate', async (req, res) => {[m
             ok: true,[m
             messageId,[m
             templateName: whatsappInitiationTemplateName,[m
[31m-            templateLanguage: whatsappInitiationTemplateLanguage[m
[32m+[m[32m            templateLanguage: resolvedTemplateLanguage[m
         });[m
 [m
         broadcast({[m
[36m@@ -1296,7 +1370,7 @@[m [mapp.post('/api/whatsapp/initiate', async (req, res) => {[m
             }[m
         });[m
 [m
[31m-        return res.json({ ok: true, id: messageId, text, previewText, templateName: whatsappInitiationTemplateName });[m
[32m+[m[32m        return res.json({ ok: true, id: messageId, text, previewText, templateName: whatsappInitiationTemplateName, templateLanguage: resolvedTemplateLanguage });[m
     } catch (error) {[m
         pushApiEvent({[m
             at: new Date().toISOString(),[m
[36m@@ -1408,6 +1482,17 @@[m [mapp.post('/api/whatsapp/forward', async (req, res) => {[m
     }[m
 });[m
 [m
[32m+[m[32mapp.use((error, _req, res, next) => {[m
[32m+[m[32m    if (error?.type === 'entity.parse.failed') {[m
[32m+[m[32m        return res.status(400).json({[m
[32m+[m[32m            ok: false,[m
[32m+[m[32m            error: 'Invalid JSON body',[m
[32m+[m[32m            hint: 'Send a valid JSON payload. In PowerShell, prefer Invoke-RestMethod or pass curl data from a file to avoid quote escaping issues.'[m
[32m+[m[32m        });[m
[32m+[m[32m    }[m
[32m+[m[32m    return next(error);[m
[32m+[m[32m});[m
[32m+[m
 try {[m
     if (initFirebasePersistence()) {[m
         await ensurePersistedStateLoaded(true);[m
