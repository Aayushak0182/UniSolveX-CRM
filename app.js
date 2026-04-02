lucide.createIcons();
        // Show agent name from localStorage
        const storedAgentName = localStorage.getItem('agentName');
        if (!storedAgentName) {
            window.location.href = 'login.html';
        }
        const agentName = storedAgentName || 'Aayush';
        document.getElementById('agentNameDisplay').textContent = 'Agent: ' + agentName;
        document.getElementById('agentInitialDisplay').textContent = agentName.charAt(0).toUpperCase();
        const themeToggleBtn = document.getElementById('themeToggleBtn');
        const themeIconSun = document.getElementById('themeIconSun');
        const themeIconMoon = document.getElementById('themeIconMoon');
        const whatsappStatusEl = document.getElementById('whatsappStatus');
        const whatsappLastEventEl = document.getElementById('whatsappLastEvent');
        const contactList = document.getElementById('contactList');
        const contactItems = Array.from(document.querySelectorAll('.contact-item'));
        const contactPanel = document.querySelector('.contact-panel');
        const contactFilter = document.getElementById('contactFilter');
        const contactSearch = document.getElementById('contactSearch');
        const contactToolbar = document.querySelector('.contact-toolbar');
        const toggleContactPanelWidthBtn = document.getElementById('toggleContactPanelWidth');
        const toggleArchivedViewBtn = document.getElementById('toggleArchivedViewBtn');
        const archivedUnreadBadge = document.getElementById('archivedUnreadBadge');
        const addContactBtn = document.getElementById('addContactBtn');
        const orderList = document.getElementById('orderList');
        const orderSearch = document.getElementById('orderSearch');
        const orderDeliveryWindow = document.getElementById('orderDeliveryWindow');
        const orderTabMine = document.getElementById('orderTabMine');
        const orderTabExpert = document.getElementById('orderTabExpert');
        const orderTabAll = document.getElementById('orderTabAll');
        const orderDetailsModal = document.getElementById('orderDetailsModal');
        const clientDetailsModal = document.getElementById('clientDetailsModal');
        const taskServiceType = document.getElementById('taskServiceType');
        const taskClientInput = document.getElementById('taskClientInput');
        const taskModalTitle = document.getElementById('taskModalTitle');
        const clientDetailsNameInput = document.getElementById('clientDetailsName');
        const clientDetailsUniversityInput = document.getElementById('clientDetailsUniversity');
        const clientDetailsSemesterInput = document.getElementById('clientDetailsSemester');
        const clientDetailsTimezoneInput = document.getElementById('clientDetailsTimezone');
        const editClientDetailsBtn = document.getElementById('editClientDetailsBtn');
        const saveClientDetailsBtn = document.getElementById('saveClientDetailsBtn');
        const closeClientDetailsBtn = document.getElementById('closeClientDetailsBtn');
        const cancelClientDetailsBtn = document.getElementById('cancelClientDetailsBtn');
        const liveSessionCreateFields = document.getElementById('liveSessionCreateFields');
        const taskSessionStart = document.getElementById('taskSessionStart');
        const taskSessionDuration = document.getElementById('taskSessionDuration');
        const chatMessages = document.getElementById('chatMessages');
        const chatInput = document.getElementById('chatInput');
        const chatFileInput = document.getElementById('chatFileInput');
        const odActualDeadlineInput = document.getElementById('odActualDeadline');
        const odExpertDeadlineInput = document.getElementById('odExpertDeadline');
        const chatHeaderEl = document.getElementById('chatHeader');
        const chatHeaderTitle = document.getElementById('chatHeaderTitle');
        const chatHeaderAvatar = document.getElementById('chatHeaderAvatar');
        const chatHeaderStatus = document.getElementById('chatHeaderStatus');
        const chatWindowTimerEl = document.getElementById('chatWindowTimer');
        const chatComposerAttachmentsEl = document.getElementById('chatComposerAttachments');
        const chatReplyPreviewEl = document.getElementById('chatReplyPreview');
        const chatReplyPreviewTextEl = document.getElementById('chatReplyPreviewText');
        const clearReplyBtn = document.getElementById('clearReplyBtn');
        const chatForwardToolbarEl = document.getElementById('chatForwardToolbar');
        const chatForwardSelectionCountEl = document.getElementById('chatForwardSelectionCount');
        const chatForwardTargetHintEl = document.getElementById('chatForwardTargetHint');
        const cancelForwardBtn = document.getElementById('cancelForwardBtn');
        const confirmForwardBtn = document.getElementById('confirmForwardBtn');
        const notifyBtn = document.getElementById('notifyBtn');
        const sendBtn = document.getElementById('sendBtn');
        const attachFileBtn = document.getElementById('attachFileBtn');
        let activeOrderCard = null;
        let activeWhatsappWaId = '';
        const orderSequenceByClient = {};
        const ORDER_LIST_STORAGE_KEY = 'unisolvex_order_list_html_v2';
        const MANUAL_CONTACTS_STORAGE_KEY = 'unisolvex_manual_contacts_v1';
        const WHATSAPP_CONTACT_ID_MAP_KEY = 'unisolvex_whatsapp_contact_id_map_v1';
        const WHATSAPP_CONTACT_ID_SEQUENCE_KEY = 'unisolvex_whatsapp_contact_id_seq_v1';
        const WHATSAPP_READ_STATE_STORAGE_KEY = 'unisolvex_whatsapp_read_state_v1';
        const WHATSAPP_API_BASE_STORAGE_KEY = 'whatsappApiBase';
        const WHATSAPP_LAST_WORKING_API_BASE_KEY = 'unisolvex_last_working_whatsapp_api_base_v1';
        const CONTACT_PANEL_WIDE_STORAGE_KEY = 'unisolvex_contact_panel_wide_v1';
        const APP_THEME_STORAGE_KEY = 'unisolvex_theme_v1';
        const configuredWhatsappApiBase = localStorage.getItem(WHATSAPP_API_BASE_STORAGE_KEY);
        const hasHttpOrigin = /^https?:\/\//i.test(window.location.origin || '');
        const RENDER_WHATSAPP_API_BASE = 'https://unisolvex-crm-backend-ra02.onrender.com';
        const defaultWhatsappApiBase = (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1' || !hasHttpOrigin)
            ? 'http://localhost:3001'
            : RENDER_WHATSAPP_API_BASE;
        const WHATSAPP_API_FALLBACK_TIMEOUT_MS = 2500;
        const WHATSAPP_POLL_INTERVAL_MS = 5000;
        const WHATSAPP_DEBUG_POLL_INTERVAL_MS = 8000;
        const WHATSAPP_CHAT_WINDOW_MS = 24 * 60 * 60 * 1000;
        const WHATSAPP_WINDOW_TICK_MS = 30000;
        const CRM_STATE_SAVE_DEBOUNCE_MS = 500;
        const whatsappMessagesByContact = {};
        const whatsappMessageStatusById = {};
        let whatsappPollingTimer = null;
        let whatsappDebugTimer = null;
        let whatsappSocketConnected = false;
        let activeWhatsappApiBase = '';
        let reachableWhatsappApiBases = [];
        let contactOrderCounter = contactItems.length;
        let whatsappWindowTimerId = null;
        let pendingAttachments = [];
        let activeReplyContext = null;
        let forwardMode = false;
        let forwardTargetSelectMode = false;
        let openMessageMenuId = '';
        let openContactMenuId = '';
        let whatsappSendInFlight = false;
        let archivedInboxMode = false;
        let whatsappTextSendProcessing = false;
        const whatsappTextSendQueue = [];
        const forwardSelectionIds = new Set();
        let activeOrderTab = 'mine';
        let activeClientDetailsWaId = '';
        let clientDetailsEditMode = false;
        let crmStateSaveTimer = null;
        let crmStateHydrationComplete = false;
        let crmStatePendingSave = false;

        function applyTheme(theme) {
            const nextTheme = theme === 'dark' ? 'dark' : 'light';
            document.documentElement.setAttribute('data-theme', nextTheme);
            document.documentElement.style.colorScheme = nextTheme;
            localStorage.setItem(APP_THEME_STORAGE_KEY, nextTheme);
            if (themeToggleBtn) {
                themeToggleBtn.setAttribute('aria-pressed', nextTheme === 'dark' ? 'true' : 'false');
                themeToggleBtn.setAttribute('title', nextTheme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme');
                themeToggleBtn.setAttribute('aria-label', nextTheme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme');
            }
            if (themeIconSun) {
                themeIconSun.classList.toggle('hidden', nextTheme === 'dark');
            }
            if (themeIconMoon) {
                themeIconMoon.classList.toggle('hidden', nextTheme !== 'dark');
            }
        }

        const savedTheme = localStorage.getItem(APP_THEME_STORAGE_KEY);
        applyTheme(savedTheme || 'light');

        if (themeToggleBtn) {
            themeToggleBtn.addEventListener('click', () => {
                const currentTheme = document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light';
                applyTheme(currentTheme === 'dark' ? 'light' : 'dark');
            });
        }

        function setWhatsappStatus(text, level) {
            if (!whatsappStatusEl) return;
            whatsappStatusEl.textContent = text || '';
            whatsappStatusEl.classList.remove('text-gray-500', 'text-green-700', 'text-amber-700', 'text-red-700');
            if (level === 'ok') whatsappStatusEl.classList.add('text-green-700');
            else if (level === 'warn') whatsappStatusEl.classList.add('text-amber-700');
            else if (level === 'err') whatsappStatusEl.classList.add('text-red-700');
            else whatsappStatusEl.classList.add('text-gray-500');
        }

        function setWhatsappLastEvent(text) {
            if (!whatsappLastEventEl) return;
            whatsappLastEventEl.textContent = text || '';
        }

        function formatShortTime(ts) {
            const d = ts ? new Date(ts) : new Date();
            if (Number.isNaN(d.getTime())) return '';
            return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        }

        if (contactPanel && toggleContactPanelWidthBtn) {
            const applyContactPanelWide = (value) => {
                const wide = Boolean(value);
                contactPanel.classList.toggle('is-wide', wide);
                toggleContactPanelWidthBtn.classList.toggle('is-wide', wide);
                toggleContactPanelWidthBtn.setAttribute('aria-pressed', wide ? 'true' : 'false');
                localStorage.setItem(CONTACT_PANEL_WIDE_STORAGE_KEY, wide ? '1' : '0');
            };

            const storedWide = localStorage.getItem(CONTACT_PANEL_WIDE_STORAGE_KEY) === '1';
            applyContactPanelWide(storedWide);

            toggleContactPanelWidthBtn.addEventListener('click', () => {
                applyContactPanelWide(!contactPanel.classList.contains('is-wide'));
            });
        }

        if (contactSearch && contactToolbar) {
            const setSearching = (value) => {
                contactToolbar.classList.toggle('is-searching', Boolean(value));
            };
            contactSearch.addEventListener('focus', () => setSearching(true));
            contactSearch.addEventListener('blur', () => setSearching(false));
            contactSearch.addEventListener('keydown', (event) => {
                if (event.key === 'Escape') {
                    contactSearch.blur();
                }
            });
        }

        function sanitizeApiBase(value) {
            return String(value || '').trim().replace(/\/+$/, '');
        }

        function shouldUseNgrokBypassHeader(baseUrl) {
            return /\.ngrok-free\.app$/i.test(String(baseUrl || '').replace(/^https?:\/\//i, '').split('/')[0] || '');
        }

        function mergeFetchOptions(baseUrl, options) {
            const nextOptions = { ...(options || {}) };
            const nextHeaders = { ...(options?.headers || {}) };
            if (shouldUseNgrokBypassHeader(baseUrl)) {
                nextHeaders['ngrok-skip-browser-warning'] = 'true';
            }
            if (Object.keys(nextHeaders).length > 0) {
                nextOptions.headers = nextHeaders;
            }
            return nextOptions;
        }

        function getWhatsappApiCandidates() {
            const rows = [];
            const add = (value) => {
                const normalized = sanitizeApiBase(value);
                if (!normalized) return;
                if (!/^https?:\/\//i.test(normalized)) return;
                if (!rows.includes(normalized)) rows.push(normalized);
            };
            add(configuredWhatsappApiBase);
            add(localStorage.getItem(WHATSAPP_LAST_WORKING_API_BASE_KEY));
            add(defaultWhatsappApiBase);

            const host = (window.location.hostname || '').toLowerCase();
            if (host === '127.0.0.1') add('http://localhost:3001');
            if (host === 'localhost') add('http://127.0.0.1:3001');
            return rows;
        }

        async function probeWhatsappApiBase(baseUrl) {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), WHATSAPP_API_FALLBACK_TIMEOUT_MS);
            try {
                const res = await fetch(baseUrl + '/health', mergeFetchOptions(baseUrl, {
                    method: 'GET',
                    signal: controller.signal
                }));
                return res.ok;
            } catch {
                return false;
            } finally {
                clearTimeout(timeoutId);
            }
        }

        async function resolveWhatsappApiBase(forceRefresh) {
            if (!forceRefresh && activeWhatsappApiBase) {
                return activeWhatsappApiBase;
            }
            const candidates = getWhatsappApiCandidates();
            for (const candidate of candidates) {
                const ok = await probeWhatsappApiBase(candidate);
                if (!ok) continue;
                activeWhatsappApiBase = candidate;
                localStorage.setItem(WHATSAPP_LAST_WORKING_API_BASE_KEY, candidate);
                return activeWhatsappApiBase;
            }
            activeWhatsappApiBase = '';
            return '';
        }

        async function getReachableWhatsappApiBases(forceRefresh) {
            if (!forceRefresh && reachableWhatsappApiBases.length) {
                return [...reachableWhatsappApiBases];
            }
            const candidates = getWhatsappApiCandidates();
            const checks = await Promise.all(candidates.map(async (candidate) => ({
                candidate,
                ok: await probeWhatsappApiBase(candidate)
            })));
            reachableWhatsappApiBases = checks.filter((row) => row.ok).map((row) => row.candidate);
            if (reachableWhatsappApiBases.length) {
                activeWhatsappApiBase = reachableWhatsappApiBases[0];
                localStorage.setItem(WHATSAPP_LAST_WORKING_API_BASE_KEY, activeWhatsappApiBase);
            }
            return [...reachableWhatsappApiBases];
        }

        async function whatsappFetch(path, options) {
            const candidates = [];
            const preferred = await resolveWhatsappApiBase(false);
            if (preferred) candidates.push(preferred);
            getWhatsappApiCandidates().forEach((candidate) => {
                if (!candidates.includes(candidate)) candidates.push(candidate);
            });

            let lastError = null;
            const attempted = [];
            for (const base of candidates) {
                try {
                    attempted.push(base + path);
                    const res = await fetch(base + path, mergeFetchOptions(base, options));
                    if (res.status === 404 || res.status === 405) {
                        continue;
                    }
                    activeWhatsappApiBase = base;
                    localStorage.setItem(WHATSAPP_LAST_WORKING_API_BASE_KEY, base);
                    return res;
                } catch (err) {
                    lastError = err;
                }
            }
            const reason = lastError?.message || 'unknown network/CORS error';
            throw new Error('WhatsApp backend unreachable (' + reason + '). Tried: ' + attempted.join(', '));
        }

        function restoreOrderListFromStorage() {
            const storedOrderListHtml = localStorage.getItem(ORDER_LIST_STORAGE_KEY);
            if (!storedOrderListHtml) return;
            orderList.innerHTML = storedOrderListHtml;
            lucide.createIcons();
        }

        function persistOrderListToStorage(skipRemote) {
            localStorage.setItem(ORDER_LIST_STORAGE_KEY, orderList.innerHTML);
            if (!skipRemote) {
                scheduleCrmStateSave();
            }
        }

        function getStoredWhatsappContactIdSequence() {
            return Number(localStorage.getItem(WHATSAPP_CONTACT_ID_SEQUENCE_KEY) || '100100');
        }

        function setWhatsappContactIdSequenceInStorage(value, skipRemote) {
            localStorage.setItem(WHATSAPP_CONTACT_ID_SEQUENCE_KEY, String(value));
            if (!skipRemote) {
                scheduleCrmStateSave();
            }
        }

        function getCurrentCrmStatePayload() {
            const rawSequence = Number(localStorage.getItem(WHATSAPP_CONTACT_ID_SEQUENCE_KEY) || '100100');
            return {
                orderListHtml: localStorage.getItem(ORDER_LIST_STORAGE_KEY) || '',
                manualContacts: readManualContactsFromStorage(),
                whatsappContactIdMap: readWhatsappContactIdMapFromStorage(),
                whatsappReadState: readWhatsappReadStateFromStorage(),
                whatsappContactIdSequence: Number.isFinite(rawSequence) ? rawSequence : 100100
            };
        }

        async function persistCrmStateToBackend() {
            const payload = getCurrentCrmStatePayload();
            try {
                const res = await whatsappFetch('/api/crm/state', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(payload)
                });
                if (!res.ok) {
                    throw new Error('CRM state save failed with status ' + res.status);
                }
            } catch (err) {
                console.warn('Failed to persist CRM state to backend:', err);
            }
        }

        function scheduleCrmStateSave() {
            if (!crmStateHydrationComplete) {
                crmStatePendingSave = true;
                return;
            }
            if (crmStateSaveTimer) {
                clearTimeout(crmStateSaveTimer);
            }
            crmStateSaveTimer = setTimeout(() => {
                crmStateSaveTimer = null;
                persistCrmStateToBackend();
            }, CRM_STATE_SAVE_DEBOUNCE_MS);
        }

        function rehydrateOrderCardsFromDom() {
            Object.keys(orderSequenceByClient).forEach((key) => {
                delete orderSequenceByClient[key];
            });
            Array.from(orderList.querySelectorAll('.order-card')).forEach((card) => {
                syncOrderServiceTag(card);
                syncOrderPaymentIndicator(card);
                applyCardTypeStyle(card, card.dataset.serviceType || card.querySelector('.order-title')?.textContent || '');
                applyDeadlineStyles(card);
                const cardClientId = (card.dataset.clientId || '').replace(/\D/g, '');
                const cardOrderId = (card.querySelector('.order-id')?.textContent || '').replace(/[^\d]/g, '');
                if (!cardClientId || !cardOrderId || !cardOrderId.startsWith(cardClientId)) return;
                const suffix = Number(cardOrderId.slice(cardClientId.length));
                if (!Number.isFinite(suffix)) return;
                orderSequenceByClient[cardClientId] = Math.max(orderSequenceByClient[cardClientId] || 1009, suffix);
            });
            applyOrderSearchFilter();
        }

        async function hydrateCrmStateFromBackend() {
            try {
                const res = await whatsappFetch('/api/crm/state');
                if (!res.ok) {
                    throw new Error('CRM state load failed with status ' + res.status);
                }
                const data = await res.json();
                const state = data?.state && typeof data.state === 'object' ? data.state : null;
                if (!state) return false;

                const hasRemoteState =
                    typeof state.orderListHtml === 'string' ||
                    Array.isArray(state.manualContacts) ||
                    (state.whatsappContactIdMap && typeof state.whatsappContactIdMap === 'object') ||
                    Number.isFinite(Number(state.whatsappContactIdSequence));

                if (!hasRemoteState) return false;

                localStorage.setItem(ORDER_LIST_STORAGE_KEY, typeof state.orderListHtml === 'string' ? state.orderListHtml : '');
                localStorage.setItem(MANUAL_CONTACTS_STORAGE_KEY, JSON.stringify(Array.isArray(state.manualContacts) ? state.manualContacts : []));
                localStorage.setItem(
                    WHATSAPP_CONTACT_ID_MAP_KEY,
                    JSON.stringify(state.whatsappContactIdMap && typeof state.whatsappContactIdMap === 'object' ? state.whatsappContactIdMap : {})
                );
                localStorage.setItem(
                    WHATSAPP_READ_STATE_STORAGE_KEY,
                    JSON.stringify(state.whatsappReadState && typeof state.whatsappReadState === 'object' ? state.whatsappReadState : {})
                );
                setWhatsappContactIdSequenceInStorage(
                    Number.isFinite(Number(state.whatsappContactIdSequence)) ? Number(state.whatsappContactIdSequence) : 100100,
                    true
                );

                orderList.innerHTML = '';
                restoreOrderListFromStorage();
                rehydrateOrderCardsFromDom();
                restoreManualContactsFromStorage();
                scheduleContactOrderRefresh();
                return true;
            } catch (err) {
                console.warn('Failed to hydrate CRM state from backend:', err);
                return false;
            }
        }

        function getNextOrderId(clientId) {
            const normalizedClientId = (clientId || '').replace(/\D/g, '') || '0000';
            if (typeof orderSequenceByClient[normalizedClientId] !== 'number') {
                orderSequenceByClient[normalizedClientId] = 1010;
            } else {
                orderSequenceByClient[normalizedClientId] += 1;
            }
            return normalizedClientId + String(orderSequenceByClient[normalizedClientId]);
        }

        function parseCardDateTime(value) {
            const raw = String(value || '').trim();
            if (!raw || /^completed$/i.test(raw)) return null;

            const slashFormat = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2}))?$/);
            if (slashFormat) {
                const dd = Number(slashFormat[1]);
                const mm = Number(slashFormat[2]);
                const yyyy = Number(slashFormat[3]);
                const hasTime = typeof slashFormat[4] === 'string' && typeof slashFormat[5] === 'string';
                const hh = hasTime ? Number(slashFormat[4]) : 23;
                const min = hasTime ? Number(slashFormat[5]) : 59;
                const d = new Date(yyyy, mm - 1, dd, hh, min, hasTime ? 0 : 59, hasTime ? 0 : 999);
                if (Number.isNaN(d.getTime())) return null;
                return d;
            }

            const d = new Date(raw);
            if (Number.isNaN(d.getTime())) return null;
            return d;
        }

        function getOrderDeliveryWindowMs() {
            const value = String(orderDeliveryWindow?.value || '').trim();
            if (!value) return 0;
            if (value === '24h') return 24 * 60 * 60 * 1000;
            if (value === '2d') return 2 * 24 * 60 * 60 * 1000;
            if (value === '7d') return 7 * 24 * 60 * 60 * 1000;
            if (value === 'past24h') return -24 * 60 * 60 * 1000;
            if (value === 'past7d') return -7 * 24 * 60 * 60 * 1000;
            return 0;
        }

        function applyOrderSearchFilter() {
            const query = (orderSearch.value || '').trim().toLowerCase().replace('#', '');
            const windowMs = getOrderDeliveryWindowMs();
            const now = new Date();
            const cards = Array.from(orderList.querySelectorAll('.order-card'));
            cards.forEach((card) => {
                const cardOrderId = (card.querySelector('.order-id')?.textContent || '').replace('#', '').trim().toLowerCase();
                const cardClientId = (card.dataset.clientId || '').trim().toLowerCase();
                const matchesQuery = !query || cardOrderId === query || cardClientId === query;
                const assignedTo = String(card.dataset.assignedTo || '').trim().toLowerCase();
                const createdBy = String(card.dataset.createdBy || '').trim().toLowerCase();
                const mineIdentity = String(agentName || '').trim().toLowerCase();
                const deadlineRaw = card.querySelector('.order-deadline')?.dataset.baseValue || card.querySelector('.order-deadline')?.textContent?.trim() || '';
                const due = parseCardDateTime(deadlineRaw);
                const matchesTab = activeOrderTab === 'all'
                    ? true
                    : activeOrderTab === 'expert'
                        ? Boolean(assignedTo) && assignedTo !== mineIdentity
                        : (createdBy === mineIdentity || assignedTo === mineIdentity || !assignedTo);
                const matchesFreshDeadline = activeOrderTab === 'all' || !due || due.getTime() >= now.getTime();

                let matchesWindow = true;
                if (windowMs > 0) {
                    if (!due) {
                        matchesWindow = false;
                    } else {
                        const diffMs = due.getTime() - now.getTime();
                        matchesWindow = diffMs >= 0 && diffMs <= windowMs;
                    }
                } else if (windowMs < 0) {
                    if (!due) {
                        matchesWindow = false;
                    } else {
                        const ageMs = now.getTime() - due.getTime();
                        matchesWindow = ageMs >= 0 && ageMs <= Math.abs(windowMs);
                    }
                }

                card.style.display = matchesQuery && matchesWindow && matchesTab && matchesFreshDeadline ? '' : 'none';
            });
        }

        function setOrderTab(tab) {
            activeOrderTab = ['mine', 'expert', 'all'].includes(tab) ? tab : 'mine';
            [orderTabMine, orderTabExpert, orderTabAll].forEach((btn) => {
                if (!btn) return;
                btn.classList.toggle('is-active', btn.dataset.orderTab === activeOrderTab);
            });
            applyOrderSearchFilter();
        }

        function updateContactStateUI(item) {
            const isPinned = item.dataset.pinned === 'true';
            const unreadCount = Number(item.dataset.unreadCount || 0);
            const isUnread = unreadCount > 0 || item.dataset.unread === 'true';
            const isArchived = String(item.dataset.tag || '').trim() === 'archived';
            const pinBtn = item.querySelector('.contact-action-pin');
            const readToggleBtn = item.querySelector('.contact-action-read');
            const archiveToggleBtn = item.querySelector('.contact-action-archive');
            const unreadDot = item.querySelector('.contact-unread-dot');
            const contactName = item.querySelector('.contact-name');
            const unreadCountBadge = item.querySelector('.contact-unread-count');
            const pinIcon = item.querySelector('.contact-pin-icon');
            const labelBadge = item.querySelector('.contact-label-badge');
            const lastTimeEl = item.querySelector('.contact-last-time');
            const labelValue = String(item.dataset.tag || '').trim();
            const latestActivity = Number(item.dataset.lastActivity || 0);

            if (pinBtn) {
                pinBtn.textContent = isPinned ? 'Unpin chat' : 'Pin chat';
            }
            if (readToggleBtn) {
                readToggleBtn.textContent = isUnread ? 'Mark as read' : 'Mark as unread';
            }
            if (archiveToggleBtn) {
                archiveToggleBtn.textContent = isArchived ? 'Unarchive chat' : 'Archive chat';
            }
            if (unreadDot) {
                unreadDot.classList.toggle('hidden', !isUnread);
            }
            if (contactName) {
                contactName.classList.toggle('font-bold', isUnread);
            }
            if (unreadCountBadge) {
                unreadCountBadge.textContent = String(Math.max(1, unreadCount || 0));
                unreadCountBadge.classList.toggle('hidden', !isUnread);
            }
            if (pinIcon) {
                pinIcon.classList.toggle('hidden', !isPinned);
            }
            if (labelBadge) {
                const labelText = getContactLabelText(labelValue);
                labelBadge.textContent = labelText;
                labelBadge.className = `contact-label-badge${labelText ? '' : ' hidden'}${labelValue ? ` is-${labelValue}` : ''}`;
            }
            if (lastTimeEl) {
                lastTimeEl.textContent = latestActivity ? formatContactLastActivity(latestActivity) : '';
            }
            updateArchivedToggleUI();
        }

        function setActiveContactItem(activeItem) {
            contactItems.forEach((item) => {
                const isActive = item === activeItem;
                item.dataset.active = isActive ? 'true' : 'false';
                item.classList.toggle('bg-indigo-50', isActive);
                item.classList.toggle('border-l-4', isActive);
                item.classList.toggle('border-l-indigo-600', isActive);
            });
        }

        function updateChatHeaderVisibility(waId) {
            if (!chatHeaderEl) return;
            chatHeaderEl.classList.toggle('is-hidden', !normalizeWaId(waId || ''));
        }

        function updateTaskClientInputFromContact(item) {
            if (!taskClientInput) return;
            const name = getContactBaseName(item?.dataset?.profileName || item?.querySelector('.contact-name')?.textContent?.trim() || '', item?.dataset?.waId || activeWhatsappWaId || '');
            const contactId = String(item?.dataset?.contactId || '').trim();
            const waId = normalizeWaId(item?.dataset?.waId || activeWhatsappWaId || '');
            if (!item || !waId) {
                taskClientInput.value = '';
                taskClientInput.placeholder = 'Select a contact first';
                return;
            }
            const displayName = name || waId;
            const suffix = contactId ? ` (${contactId})` : '';
            taskClientInput.value = displayName + suffix;
            taskClientInput.placeholder = 'Selected contact';
        }

        function refreshContactOrder() {
            const sorted = [...contactItems].sort((a, b) => {
                const pinnedDiff = Number(b.dataset.pinned === 'true') - Number(a.dataset.pinned === 'true');
                if (pinnedDiff !== 0) return pinnedDiff;

                const lastActivityDiff = Number(b.dataset.lastActivity || 0) - Number(a.dataset.lastActivity || 0);
                if (lastActivityDiff !== 0) return lastActivityDiff;

                return Number(a.dataset.order) - Number(b.dataset.order);
            });
            sorted.forEach((item) => contactList.appendChild(item));
        }

        function scheduleContactOrderRefresh() {
            refreshContactOrder();
            setTimeout(() => {
                refreshContactOrder();
                applyContactFilter();
            }, 0);
        }

        function applyContactFilter() {
            const filterValue = contactFilter.value;
            const searchValue = (contactSearch.value || '').trim().toLowerCase();
            contactItems.forEach((item) => {
                const itemTag = item.dataset.tag || '';
                const isArchived = itemTag === 'archived';
                const nameText = (item.querySelector('.contact-name')?.textContent || '').toLowerCase();
                const contactIdText = (item.dataset.contactId || '').trim().toLowerCase();
                const idText = (item.querySelector('.contact-meta')?.textContent || item.querySelector('p')?.textContent || '').toLowerCase();
                const matchesTag = !filterValue || itemTag === filterValue;
                const matchesSearch = !searchValue || nameText.includes(searchValue) || contactIdText.includes(searchValue) || idText.includes(searchValue);
                const matchesArchiveMode = archivedInboxMode ? isArchived : !isArchived;
                item.style.display = matchesArchiveMode && matchesTag && matchesSearch ? '' : 'none';
            });
            updateArchivedToggleUI();
        }

        function updateArchivedToggleUI() {
            if (!toggleArchivedViewBtn) return;
            const archivedUnreadCount = contactItems.reduce((sum, item) => {
                if (String(item?.dataset?.tag || '') !== 'archived') return sum;
                return sum + (Math.max(0, Number(item.dataset.unreadCount || 0)) > 0 ? 1 : 0);
            }, 0);
            toggleArchivedViewBtn.classList.toggle('is-active', archivedInboxMode);
            toggleArchivedViewBtn.setAttribute('aria-pressed', archivedInboxMode ? 'true' : 'false');
            toggleArchivedViewBtn.title = archivedInboxMode ? 'Show active chats' : 'Show archived chats';
            if (archivedUnreadBadge) {
                archivedUnreadBadge.textContent = String(Math.max(1, archivedUnreadCount));
                archivedUnreadBadge.classList.toggle('hidden', archivedUnreadCount <= 0);
            }
        }

        function normalizeWaId(rawValue) {
            return String(rawValue || '').replace(/[^\d]/g, '');
        }

        function getContactCodeLabel(waId) {
            const normalized = normalizeWaId(waId);
            if (!normalized) return '';
            return normalized.slice(0, Math.min(3, normalized.length));
        }

        function getContactBaseName(profileName, waId) {
            const normalized = normalizeWaId(waId);
            const safeName = String(profileName || '').trim();
            if (!safeName) return normalized;
            return normalizeWaId(safeName) === normalized ? normalized : safeName;
        }

        function getContactDisplayName(profileName, waId) {
            const code = getContactCodeLabel(waId);
            const baseName = getContactBaseName(profileName, waId);
            return code ? `(${code}) ${baseName}` : baseName;
        }

        function getContactInitials(profileName, waId) {
            const baseName = getContactBaseName(profileName, waId);
            const words = baseName.split(/\s+/).filter(Boolean);
            if (!words.length) return (normalizeWaId(waId).charAt(0) || '?').toUpperCase();
            if (words.length === 1) return words[0].charAt(0).toUpperCase();
            return (words[0].charAt(0) + words[1].charAt(0)).toUpperCase();
        }

        function getContactLabelText(tagValue) {
            const value = String(tagValue || '').trim().toLowerCase();
            if (!value) return '';
            if (value === 'new') return 'New Client';
            if (value === 'old') return 'Old Client';
            if (value === 'tutor') return 'Expert';
            if (value === 'friend') return 'Friend';
            if (value === 'useless') return 'Broker';
            if (value === 'archived') return 'Archived';
            return value;
        }

        function formatContactLastActivity(timestampMs) {
            const time = Number(timestampMs || 0);
            if (!time) return '';
            const date = new Date(time);
            if (Number.isNaN(date.getTime())) return '';
            const now = new Date();
            const isSameDay = now.toDateString() === date.toDateString();
            if (isSameDay) {
                return formatChatTime(date.toISOString());
            }
            const yesterday = new Date();
            yesterday.setDate(yesterday.getDate() - 1);
            if (yesterday.toDateString() === date.toDateString()) {
                return 'Yesterday';
            }
            return date.toLocaleDateString(undefined, { day: 'numeric', month: 'short' });
        }

        function setContactUnreadCount(item, nextCount) {
            if (!item) return;
            const count = Math.max(0, Number(nextCount || 0));
            item.dataset.unreadCount = String(count);
            item.dataset.unread = count > 0 ? 'true' : 'false';
            updateContactStateUI(item);
            scheduleContactOrderRefresh();
        }

        function getPreferredProfileName(existingName, incomingName, waId) {
            const normalizedWaId = normalizeWaId(waId);
            const safeExisting = String(existingName || '').trim();
            const safeIncoming = String(incomingName || '').trim();
            const incomingLooksLikeNumber = safeIncoming && normalizeWaId(safeIncoming) === normalizedWaId;
            if (safeIncoming && !incomingLooksLikeNumber) return safeIncoming;
            if (safeExisting) return safeExisting;
            return safeIncoming || normalizedWaId;
        }

        function getAutoTimezone() {
            try {
                return Intl.DateTimeFormat().resolvedOptions().timeZone || '';
            } catch {
                return '';
            }
        }

        function getManualContactByWaId(waId) {
            const normalizedWaId = normalizeWaId(waId);
            return readManualContactsFromStorage().find((row) => normalizeWaId(row.waId) === normalizedWaId) || null;
        }

        function setClientDetailsEditMode(nextMode) {
            clientDetailsEditMode = Boolean(nextMode);
            [clientDetailsNameInput, clientDetailsUniversityInput, clientDetailsSemesterInput, clientDetailsTimezoneInput].forEach((input) => {
                if (!input) return;
                input.readOnly = !clientDetailsEditMode;
                input.classList.toggle('bg-gray-100', !clientDetailsEditMode);
                input.classList.toggle('text-gray-600', !clientDetailsEditMode);
                input.classList.toggle('cursor-not-allowed', !clientDetailsEditMode);
            });
            if (saveClientDetailsBtn) saveClientDetailsBtn.classList.toggle('hidden', !clientDetailsEditMode);
            if (cancelClientDetailsBtn) cancelClientDetailsBtn.classList.toggle('hidden', !clientDetailsEditMode);
            if (editClientDetailsBtn) editClientDetailsBtn.classList.toggle('hidden', clientDetailsEditMode);
        }

        function openClientDetailsModal(item) {
            const waId = normalizeWaId(item?.dataset?.waId || '');
            if (!waId || !clientDetailsModal) return;
            const saved = getManualContactByWaId(waId);
            activeClientDetailsWaId = waId;
            if (clientDetailsNameInput) clientDetailsNameInput.value = saved?.profileName || item?.dataset?.profileName || '';
            if (clientDetailsUniversityInput) clientDetailsUniversityInput.value = String(saved?.universityName || item?.dataset?.universityName || '');
            if (clientDetailsSemesterInput) clientDetailsSemesterInput.value = String(saved?.semester || item?.dataset?.semester || '');
            if (clientDetailsTimezoneInput) clientDetailsTimezoneInput.value = String(saved?.timezone || item?.dataset?.timezone || getAutoTimezone());
            setClientDetailsEditMode(false);
            clientDetailsModal.classList.remove('hidden');
        }

        function readManualContactsFromStorage() {
            try {
                const raw = localStorage.getItem(MANUAL_CONTACTS_STORAGE_KEY);
                if (!raw) return [];
                const parsed = JSON.parse(raw);
                return Array.isArray(parsed) ? parsed : [];
            } catch {
                return [];
            }
        }

        function writeManualContactsToStorage(rows, skipRemote) {
            localStorage.setItem(MANUAL_CONTACTS_STORAGE_KEY, JSON.stringify(rows));
            if (!skipRemote) {
                scheduleCrmStateSave();
            }
        }

        function saveManualContact(contact) {
            if (!contact?.waId) return;
            const rows = readManualContactsFromStorage();
            const existingIndex = rows.findIndex((row) => row.waId === contact.waId);
            const previous = existingIndex >= 0 ? rows[existingIndex] : null;
            const payload = {
                waId: contact.waId,
                profileName: getPreferredProfileName(previous?.profileName, contact.profileName, contact.waId),
                tag: typeof contact.tag === 'string' ? contact.tag : String(previous?.tag || ''),
                universityName: typeof contact.universityName === 'string' ? contact.universityName : String(previous?.universityName || ''),
                semester: typeof contact.semester === 'string' ? contact.semester : String(previous?.semester || ''),
                timezone: typeof contact.timezone === 'string' ? contact.timezone : String(previous?.timezone || getAutoTimezone())
            };
            if (existingIndex >= 0) {
                rows[existingIndex] = payload;
            } else {
                rows.push(payload);
            }
            writeManualContactsToStorage(rows);
        }

        function restoreManualContactsFromStorage() {
            const rows = readManualContactsFromStorage();
            rows.forEach((contact) => {
                upsertWhatsappContact({
                    waId: normalizeWaId(contact.waId),
                    profileName: contact.profileName,
                    tag: String(contact.tag || ''),
                    universityName: String(contact.universityName || ''),
                    semester: String(contact.semester || ''),
                    timezone: String(contact.timezone || getAutoTimezone())
                }, false);
            });
        }

        function readWhatsappContactIdMapFromStorage() {
            try {
                const raw = localStorage.getItem(WHATSAPP_CONTACT_ID_MAP_KEY);
                if (!raw) return {};
                const parsed = JSON.parse(raw);
                return parsed && typeof parsed === 'object' ? parsed : {};
            } catch {
                return {};
            }
        }

        function readWhatsappReadStateFromStorage() {
            try {
                const raw = localStorage.getItem(WHATSAPP_READ_STATE_STORAGE_KEY);
                if (!raw) return {};
                const parsed = JSON.parse(raw);
                return parsed && typeof parsed === 'object' ? parsed : {};
            } catch {
                return {};
            }
        }

        function writeWhatsappReadStateToStorage(map, skipRemote) {
            localStorage.setItem(WHATSAPP_READ_STATE_STORAGE_KEY, JSON.stringify(map || {}));
            if (!skipRemote) {
                scheduleCrmStateSave();
            }
        }

        function getLatestIncomingMessageTimestamp(waId) {
            const normalizedWaId = normalizeWaId(waId);
            const thread = whatsappMessagesByContact[normalizedWaId] || [];
            for (let i = thread.length - 1; i >= 0; i -= 1) {
                if (String(thread[i]?.direction || '') !== 'incoming') continue;
                return String(thread[i]?.timestamp || '');
            }
            return '';
        }

        function markWhatsappThreadRead(waId, skipRemote) {
            const normalizedWaId = normalizeWaId(waId);
            if (!normalizedWaId) return;
            const latestIncomingTimestamp = getLatestIncomingMessageTimestamp(normalizedWaId);
            const readState = readWhatsappReadStateFromStorage();
            if (latestIncomingTimestamp) {
                readState[normalizedWaId] = latestIncomingTimestamp;
            } else {
                delete readState[normalizedWaId];
            }
            writeWhatsappReadStateToStorage(readState, skipRemote);
            const item = contactList.querySelector('.contact-item[data-wa-id="' + normalizedWaId + '"]');
            if (item) {
                setContactUnreadCount(item, 0);
            }
        }

        function getUnreadCountForThread(waId, thread) {
            const normalizedWaId = normalizeWaId(waId);
            const readState = readWhatsappReadStateFromStorage();
            const lastReadTime = new Date(readState[normalizedWaId] || 0).getTime() || 0;
            return (Array.isArray(thread) ? thread : []).reduce((sum, message) => {
                if (String(message?.direction || '') !== 'incoming') return sum;
                const messageTime = new Date(message?.timestamp || 0).getTime() || 0;
                return messageTime > lastReadTime ? sum + 1 : sum;
            }, 0);
        }

        function writeWhatsappContactIdMapToStorage(map, skipRemote) {
            localStorage.setItem(WHATSAPP_CONTACT_ID_MAP_KEY, JSON.stringify(map));
            if (!skipRemote) {
                scheduleCrmStateSave();
            }
        }

        function getNextWhatsappContactIdValue(map) {
            const stored = getStoredWhatsappContactIdSequence();
            const mapIds = Object.values(map).map((value) => Number(value)).filter((value) => Number.isFinite(value));
            const maxMapId = mapIds.length ? Math.max(...mapIds) : 100100;
            const next = Math.max(stored, maxMapId) + 1;
            setWhatsappContactIdSequenceInStorage(next);
            return next;
        }

        function migrateWhatsappContactIdsToSixDigitSeries() {
            const map = readWhatsappContactIdMapFromStorage();
            const entries = Object.entries(map);
            if (!entries.length) {
                setWhatsappContactIdSequenceInStorage(100100);
                return;
            }

            const shouldMigrate = entries.some((entry) => Number(entry[1]) < 100101);
            if (!shouldMigrate) {
                const maxExisting = Math.max(...entries.map((entry) => Number(entry[1])).filter((value) => Number.isFinite(value)));
                setWhatsappContactIdSequenceInStorage(maxExisting);
                return;
            }

            const migrated = {};
            let nextId = 100101;
            entries
                .sort((a, b) => Number(a[1]) - Number(b[1]))
                .forEach((entry) => {
                    const waId = entry[0];
                    migrated[waId] = String(nextId);
                    nextId += 1;
                });

            writeWhatsappContactIdMapToStorage(migrated);
            setWhatsappContactIdSequenceInStorage(nextId - 1);
        }

        function getOrCreateWhatsappContactId(waId) {
            const normalizedWaId = normalizeWaId(waId);
            if (!normalizedWaId) return '';
            const map = readWhatsappContactIdMapFromStorage();
            if (map[normalizedWaId]) return String(map[normalizedWaId]);
            const nextId = String(getNextWhatsappContactIdValue(map));
            map[normalizedWaId] = nextId;
            writeWhatsappContactIdMapToStorage(map);
            return nextId;
        }

        function applyOrderStatusStyle(statusEl, statusValue) {
            const normalized = (statusValue || '').toLowerCase();
            const card = statusEl?.closest('.order-card');
            if (card) {
                card.dataset.orderStatus = normalized || 'new';
            }
            statusEl.className = 'order-status text-[10px] font-bold uppercase px-2 py-0.5 rounded';
            if (normalized === 'completed') {
                statusEl.classList.add('text-green-800', 'bg-green-100');
                return;
            }
            if (normalized === 'new') {
                statusEl.classList.add('text-blue-600', 'bg-blue-50');
                return;
            }
            if (normalized === 'on hold') {
                statusEl.classList.add('text-yellow-800', 'bg-yellow-100');
                return;
            }
            if (normalized === 'cancelled') {
                statusEl.classList.add('text-red-700', 'bg-red-100');
                return;
            }
            if (normalized === 'refund') {
                statusEl.classList.add('text-red-700', 'bg-red-100');
                return;
            }
            if (normalized === 'archive') {
                statusEl.classList.add('text-gray-700', 'bg-gray-200');
                return;
            }
            statusEl.classList.add('text-green-800', 'bg-green-100');
        }

        function applyCardTypeStyle(card, serviceTypeValue) {
            const value = (serviceTypeValue || '').toLowerCase();
            card.dataset.serviceTone = 'default';
            card.classList.remove('bg-white', 'bg-orange-50', 'bg-amber-50', 'bg-blue-50', 'bg-red-50', 'bg-gray-100', 'border-orange-200', 'border-amber-200', 'border-blue-200', 'border-red-200', 'border-gray-300');

            if (value.includes('live session')) {
                card.dataset.serviceTone = 'live-session';
                card.classList.add('bg-amber-50', 'border-amber-200');
                return;
            }
            if (value.includes('assignment')) {
                card.dataset.serviceTone = 'assignment';
                card.classList.add('bg-blue-50', 'border-blue-200');
                return;
            }
            if (value.includes('project')) {
                card.dataset.serviceTone = 'project';
                card.classList.add('bg-red-50', 'border-red-200');
                return;
            }
            if (value.includes('full course')) {
                card.dataset.serviceTone = 'full-course';
                card.classList.add('bg-gray-100', 'border-gray-300');
                return;
            }
            if (value.includes('consultation')) {
                card.dataset.serviceTone = 'consultation';
            }
            card.classList.add('bg-white', 'border-gray-200');
        }

        function parseAmountToNumber(value) {
            if (!value) return 0;
            const normalized = String(value).replace(/[^0-9.]/g, '');
            const parsed = Number(normalized);
            return Number.isFinite(parsed) ? parsed : 0;
        }

        function syncOrderPaymentIndicator(card) {
            const amountEl = card.querySelector('.order-amount');
            if (!amountEl) return;

            const caLine = amountEl.closest('p');
            if (!caLine) return;

            let indicatorEl = caLine.querySelector('.order-payment-indicator');
            if (!indicatorEl) {
                indicatorEl = document.createElement('span');
                indicatorEl.className = 'order-payment-indicator hidden ml-1';
                caLine.appendChild(indicatorEl);
            }

            const totalAmount = parseAmountToNumber(amountEl.textContent || '');
            const paidAmount = parseAmountToNumber(card.dataset.clientPaidAmount || '');
            const manualStatus = (card.dataset.paymentStatusOverride || 'auto').toLowerCase();
            const hasPaymentInfo = paidAmount > 0 && totalAmount > 0;

            if (manualStatus === 'complete') {
                indicatorEl.className = 'order-payment-indicator ml-1 inline-flex items-center justify-center w-4 h-4 rounded-full bg-green-600 text-white text-[10px] font-bold';
                indicatorEl.innerHTML = '&#10003;';
                return;
            }

            if (manualStatus === 'partial') {
                indicatorEl.className = 'order-payment-indicator ml-1 inline-flex items-center gap-1 text-[10px] font-semibold text-gray-700';
                indicatorEl.innerHTML = '<span class="inline-flex items-center justify-center w-4 h-4 rounded-full bg-black text-white text-[9px] leading-none">...</span><span>Partially Paid</span>';
                return;
            }

            if (!hasPaymentInfo) {
                indicatorEl.className = 'order-payment-indicator hidden ml-1';
                indicatorEl.textContent = '';
                return;
            }

            if (paidAmount >= totalAmount) {
                indicatorEl.className = 'order-payment-indicator ml-1 inline-flex items-center justify-center w-4 h-4 rounded-full bg-green-600 text-white text-[10px] font-bold';
                indicatorEl.innerHTML = '&#10003;';
                return;
            }

            indicatorEl.className = 'order-payment-indicator ml-1 inline-flex items-center gap-1 text-[10px] font-semibold text-gray-700';
            indicatorEl.innerHTML = '<span class="inline-flex items-center justify-center w-4 h-4 rounded-full bg-black text-white text-[9px] leading-none">...</span><span>Partially Paid</span>';
        }

        function syncOrderServiceTag(card) {
            let titleRow = card.querySelector('.order-title-row');
            if (!titleRow) return;

            let tagEl = titleRow.querySelector('.order-service-tag');
            if (!tagEl) {
                tagEl = document.createElement('span');
                tagEl.className = 'order-service-tag text-[10px] font-semibold px-2 py-0.5 rounded-full bg-indigo-100 text-indigo-700';
                titleRow.prepend(tagEl);
            }

            const serviceTypeValue = (card.dataset.serviceType || '').trim();
            tagEl.textContent = serviceTypeValue;
            tagEl.classList.toggle('hidden', !serviceTypeValue);
        }

        function formatDateTimeForCard(dateValue) {
            if (!dateValue) return '';
            if (/^\d{2}\/\d{2}\/\d{4}/.test(dateValue)) return dateValue;
            const d = new Date(dateValue);
            if (Number.isNaN(d.getTime())) return dateValue;
            const dd = String(d.getDate()).padStart(2, '0');
            const mm = String(d.getMonth() + 1).padStart(2, '0');
            const yyyy = d.getFullYear();
            const hh = String(d.getHours()).padStart(2, '0');
            const min = String(d.getMinutes()).padStart(2, '0');
            return dd + '/' + mm + '/' + yyyy + ' ' + hh + ':' + min;
        }

        function getDeadlineVisualState(dateValue, options) {
            const due = parseCardDateTime(dateValue);
            const thresholdMs = (options?.thresholdHours || 48) * 60 * 60 * 1000;
            if (!due) {
                return {
                    tone: 'muted',
                    suffix: ''
                };
            }
            const diffMs = due.getTime() - Date.now();
            if (diffMs < 0) {
                return {
                    tone: 'urgent',
                    suffix: options?.showDueSoon ? 'expired' : ''
                };
            }
            if (diffMs > thresholdMs) {
                return {
                    tone: 'safe',
                    suffix: ''
                };
            }
            return {
                tone: 'urgent',
                suffix: options?.showDueSoon ? ' (due soon)' : ''
            };
        }

        function applyDeadlineStyles(card) {
            if (!card) return;
            const actualEl = card.querySelector('.order-deadline');
            const expertEl = card.querySelector('.order-expert-deadline');
            const orderIdEl = card.querySelector('.order-id');
            const actualRaw = actualEl?.dataset.baseValue || actualEl?.textContent || '';
            const expertRaw = expertEl?.dataset.baseValue || expertEl?.textContent || '';
            let dueSoonChip = card.querySelector('.order-due-soon-chip');

            if (!dueSoonChip && orderIdEl) {
                dueSoonChip = document.createElement('span');
                dueSoonChip.className = 'order-due-soon-chip hidden text-[9px] font-semibold uppercase tracking-[0.08em] px-2 py-0.5 rounded-full';
                orderIdEl.insertAdjacentElement('afterend', dueSoonChip);
            }

            if (actualEl) {
                const actualState = getDeadlineVisualState(actualRaw, { thresholdHours: 48, showDueSoon: true });
                actualEl.dataset.baseValue = String(actualRaw || '');
                actualEl.classList.remove('deadline-safe', 'deadline-urgent', 'deadline-muted');
                actualEl.classList.add(actualState.tone === 'safe' ? 'deadline-safe' : actualState.tone === 'urgent' ? 'deadline-urgent' : 'deadline-muted');
                actualEl.textContent = String(actualRaw || '');
                if (dueSoonChip) {
                    dueSoonChip.classList.toggle('hidden', !actualState.suffix);
                    dueSoonChip.textContent = actualState.suffix === 'expired' ? 'Expired' : actualState.suffix ? 'Due Soon' : '';
                }
            }

            if (expertEl) {
                const expertState = getDeadlineVisualState(expertRaw, { thresholdHours: 48, showDueSoon: false });
                expertEl.dataset.baseValue = String(expertRaw || '');
                expertEl.classList.remove('deadline-safe', 'deadline-urgent', 'deadline-muted');
                expertEl.classList.add(expertState.tone === 'safe' ? 'deadline-safe' : expertState.tone === 'urgent' ? 'deadline-urgent' : 'deadline-muted');
                expertEl.textContent = String(expertRaw || '');
            }
        }

        function toDateTimeLocalValue(value) {
            if (!value) return '';
            if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(value)) return value.slice(0, 16);
            const slashFormat = value.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2}))?$/);
            if (slashFormat) {
                const dd = slashFormat[1];
                const mm = slashFormat[2];
                const yyyy = slashFormat[3];
                const hh = slashFormat[4] || '00';
                const min = slashFormat[5] || '00';
                return yyyy + '-' + mm + '-' + dd + 'T' + hh + ':' + min;
            }
            const d = new Date(value);
            if (Number.isNaN(d.getTime())) return '';
            const yyyy = d.getFullYear();
            const mm = String(d.getMonth() + 1).padStart(2, '0');
            const dd = String(d.getDate()).padStart(2, '0');
            const hh = String(d.getHours()).padStart(2, '0');
            const min = String(d.getMinutes()).padStart(2, '0');
            return yyyy + '-' + mm + '-' + dd + 'T' + hh + ':' + min;
        }

        function parseComments(raw) {
            if (!raw) return [];
            try {
                const parsed = JSON.parse(raw);
                return Array.isArray(parsed) ? parsed : [];
            } catch (err) {
                return [];
            }
        }

        function formatFileSize(bytes) {
            if (!Number.isFinite(bytes) || bytes <= 0) return '0 B';
            const units = ['B', 'KB', 'MB', 'GB'];
            let size = bytes;
            let unitIndex = 0;
            while (size >= 1024 && unitIndex < units.length - 1) {
                size /= 1024;
                unitIndex += 1;
            }
            const precision = unitIndex === 0 ? 0 : 1;
            return size.toFixed(precision) + ' ' + units[unitIndex];
        }

        function detectAttachmentType(fileName, mimeType) {
            const type = String(mimeType || '').toLowerCase();
            const name = String(fileName || '').toLowerCase();
            if (type.startsWith('image/')) return 'image';
            if (type.startsWith('audio/')) return 'audio';
            if (type.startsWith('video/')) return 'video';
            if (name.endsWith('.jpg') || name.endsWith('.jpeg') || name.endsWith('.png') || name.endsWith('.gif') || name.endsWith('.webp')) return 'image';
            if (name.endsWith('.mp3') || name.endsWith('.wav') || name.endsWith('.ogg') || name.endsWith('.m4a')) return 'audio';
            if (name.endsWith('.mp4') || name.endsWith('.mov') || name.endsWith('.webm')) return 'video';
            return 'document';
        }

        function ensureMessageClientKey(message) {
            if (!message) return '';
            if (!message.clientKey) {
                const parts = [
                    message.id || '',
                    message.direction || 'msg',
                    message.timestamp || '',
                    message.messageType || 'text',
                    message.text || '',
                    message.attachmentName || '',
                    message.mediaId || '',
                    message.replyTo?.id || ''
                ];
                message.clientKey = parts.join('|');
            }
            return message.clientKey;
        }

        function getMessagePreviewText(message) {
            if (!message) return '';
            if (message.messageType && message.messageType !== 'text') {
                return `${message.messageType}: ${message.attachmentName || message.text || ''}`.trim();
            }
            return String(message.text || '').trim();
        }

        function setReplyContext(message) {
            activeReplyContext = message ? {
                id: message.id || '',
                clientKey: ensureMessageClientKey(message),
                text: getMessagePreviewText(message),
                direction: message.direction || 'incoming'
            } : null;
            if (!chatReplyPreviewEl || !chatReplyPreviewTextEl) return;
            if (!activeReplyContext?.id) {
                chatReplyPreviewEl.classList.add('hidden');
                chatReplyPreviewEl.classList.remove('flex');
                chatReplyPreviewTextEl.textContent = '';
                return;
            }
            chatReplyPreviewTextEl.textContent = activeReplyContext.text || '(message)';
            chatReplyPreviewEl.classList.remove('hidden');
            chatReplyPreviewEl.classList.add('flex');
        }

        function setForwardMode(nextState) {
            forwardMode = Boolean(nextState);
            if (!forwardMode) {
                forwardSelectionIds.clear();
                forwardTargetSelectMode = false;
            }
            if (chatForwardToolbarEl) {
                chatForwardToolbarEl.classList.toggle('hidden', !forwardMode);
                chatForwardToolbarEl.classList.toggle('flex', forwardMode);
            }
            if (chatForwardSelectionCountEl) {
                chatForwardSelectionCountEl.textContent = `${forwardSelectionIds.size} selected`;
            }
            if (chatForwardTargetHintEl) {
                chatForwardTargetHintEl.textContent = forwardTargetSelectMode
                    ? 'Now click a target contact from the left list.'
                    : 'Select messages, then click "Choose Contact".';
            }
            renderWhatsappMessages(activeWhatsappWaId);
        }

        function renderAttachmentQueue() {
            if (!chatComposerAttachmentsEl) return;
            chatComposerAttachmentsEl.innerHTML = '';
            chatComposerAttachmentsEl.classList.toggle('hidden', pendingAttachments.length === 0);
            pendingAttachments.forEach((file, index) => {
                const chip = document.createElement('div');
                chip.className = 'chat-attachment-chip';
                chip.innerHTML = `
                    <span>${escapeHtml(file.name)} (${formatFileSize(file.size)})</span>
                    <button type="button" data-index="${index}" aria-label="Remove attachment">&times;</button>
                `;
                chatComposerAttachmentsEl.appendChild(chip);
            });
            Array.from(chatComposerAttachmentsEl.querySelectorAll('button[data-index]')).forEach((btn) => {
                btn.onclick = function() {
                    const index = Number(this.dataset.index);
                    pendingAttachments = pendingAttachments.filter((_, itemIndex) => itemIndex !== index);
                    renderAttachmentQueue();
                };
            });
        }

        function queueAttachments(files) {
            const selectedFiles = Array.from(files || []);
            if (!selectedFiles.length) return;
            pendingAttachments.push(...selectedFiles);
            renderAttachmentQueue();
        }

        function getWhatsappWindowState(messages) {
            const rows = Array.isArray(messages) ? messages : [];
            const latestIncoming = [...rows]
                .filter((row) => row?.direction === 'incoming' && row?.timestamp)
                .sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime())[0];
            if (!latestIncoming) {
                return {
                    state: 'waiting',
                    label: '24h window: waiting for client message',
                    shortLabel: 'Waiting for first reply'
                };
            }
            const startedAt = new Date(latestIncoming.timestamp).getTime();
            if (!Number.isFinite(startedAt)) {
                return {
                    state: 'waiting',
                    label: '24h window: waiting for client message',
                    shortLabel: 'Waiting for first reply'
                };
            }
            const expiresAt = startedAt + WHATSAPP_CHAT_WINDOW_MS;
            const remainingMs = expiresAt - Date.now();
            if (remainingMs <= 0) {
                return {
                    state: 'expired',
                    label: '24h window: expired - use chat initiation',
                    shortLabel: 'Expired',
                    expiresAt
                };
            }
            const totalMinutes = Math.floor(remainingMs / 60000);
            const hours = Math.floor(totalMinutes / 60);
            const minutes = totalMinutes % 60;
            const pretty = `${String(hours).padStart(2, '0')}h ${String(minutes).padStart(2, '0')}m left`;
            return {
                state: 'open',
                label: `24h window: ${pretty}`,
                shortLabel: pretty,
                expiresAt
            };
        }

        function updateChatWindowState(waId) {
            if (!chatHeaderStatus || !chatWindowTimerEl) return;
            updateChatHeaderVisibility(waId);
            if (!waId) {
                chatHeaderStatus.textContent = '';
                chatWindowTimerEl.textContent = 'EXPIRED';
                chatWindowTimerEl.classList.remove('text-green-700', 'text-red-600', 'text-amber-700');
                chatWindowTimerEl.classList.add('text-red-600');
                return;
            }
            const state = getWhatsappWindowState(whatsappMessagesByContact[waId] || []);
            const isOpen = state.state === 'open';
            chatHeaderStatus.textContent = '';
            chatWindowTimerEl.textContent = isOpen ? state.shortLabel.replace(' left', '') : 'EXPIRED';
            chatWindowTimerEl.classList.remove('text-green-700', 'text-red-600', 'text-amber-700');
            chatWindowTimerEl.classList.add(isOpen ? 'text-green-700' : 'text-red-600');
        }

        function refreshWhatsappWindowStateUi() {
            updateChatWindowState(activeWhatsappWaId);
        }

        function initWhatsappWindowTimer() {
            if (whatsappWindowTimerId) return;
            refreshWhatsappWindowStateUi();
            whatsappWindowTimerId = setInterval(() => {
                refreshWhatsappWindowStateUi();
            }, WHATSAPP_WINDOW_TICK_MS);
        }

        function renderCommentHistory(comments) {
            const target = document.getElementById('odCommentHistory');
            if (!comments.length) {
                target.innerHTML = '<p class="text-gray-400">No comments yet.</p>';
                return;
            }
            target.innerHTML = comments.map((item) => {
                const when = formatDateTimeForCard(item.ts || '');
                const who = item.agent || 'Agent';
                const text = item.text || '';
                return '<div class="border-b pb-1"><p class="font-semibold">' + who + ' : ' + when + '</p><p>comment ' + text + '</p></div>';
            }).join('');
        }

        function appendCommentToActiveCard(commentText) {
            if (!activeOrderCard || !commentText) return;
            const comments = parseComments(activeOrderCard.dataset.comments || '[]');
            comments.push({
                agent: agentName,
                text: commentText,
                ts: new Date().toISOString()
            });
            activeOrderCard.dataset.comments = JSON.stringify(comments);
            activeOrderCard.dataset.activity = commentText;

            let commentEl = activeOrderCard.querySelector('.order-comment');
            if (!commentEl) {
                commentEl = document.createElement('p');
                commentEl.className = 'order-comment hidden mt-1 text-[11px] text-gray-700 border-t pt-1';
                activeOrderCard.appendChild(commentEl);
            }
            commentEl.textContent = 'Note: ' + agentName + ' - ' + commentText;
            commentEl.classList.remove('hidden');

            renderCommentHistory(comments);
            persistOrderListToStorage();
        }

        function wireContactItem(item) {
            if (!item || item.dataset.wired === 'true') return;
            item.dataset.wired = 'true';
            if (!item.dataset.order) {
                item.dataset.order = String(contactOrderCounter);
                contactOrderCounter += 1;
            }
            if (!item.dataset.pinned) item.dataset.pinned = 'false';
            if (!item.dataset.unread) item.dataset.unread = 'false';
            if (!item.dataset.unreadCount) item.dataset.unreadCount = item.dataset.unread === 'true' ? '1' : '0';
            if (!item.dataset.tag) item.dataset.tag = '';
            if (!item.dataset.active) item.dataset.active = 'false';
            if (!item.dataset.lastActivity) item.dataset.lastActivity = '0';
            updateContactStateUI(item);

            const menuToggleBtn = item.querySelector('.contact-menu-toggle');
            const pinBtn = item.querySelector('.contact-action-pin');
            const readToggleBtn = item.querySelector('.contact-action-read');
            const archiveToggleBtn = item.querySelector('.contact-action-archive');
            const tagSelect = item.querySelector('.contact-menu-label');

            if (menuToggleBtn) {
                menuToggleBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const waId = normalizeWaId(item.dataset.waId || '');
                    openContactMenuId = openContactMenuId === waId ? '' : waId;
                    contactItems.forEach((contactItem) => {
                        const menu = contactItem.querySelector('.contact-card-menu');
                        if (!menu) return;
                        const isOpen = normalizeWaId(contactItem.dataset.waId || '') === openContactMenuId;
                        contactItem.classList.toggle('is-menu-open', isOpen);
                        menu.classList.toggle('hidden', !isOpen);
                    });
                });
            }

            if (pinBtn) {
                pinBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    item.dataset.pinned = item.dataset.pinned === 'true' ? 'false' : 'true';
                    updateContactStateUI(item);
                    scheduleContactOrderRefresh();
                    openContactMenuId = '';
                    item.classList.remove('is-menu-open');
                    item.querySelector('.contact-card-menu')?.classList.add('hidden');
                });
            }

            if (readToggleBtn) {
                readToggleBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const nextCount = Number(item.dataset.unreadCount || 0) > 0 ? 0 : 1;
                    setContactUnreadCount(item, nextCount);
                    openContactMenuId = '';
                    item.classList.remove('is-menu-open');
                    item.querySelector('.contact-card-menu')?.classList.add('hidden');
                });
            }

            if (archiveToggleBtn) {
                archiveToggleBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const nextArchived = String(item.dataset.tag || '') === 'archived' ? '' : 'archived';
                    item.dataset.tag = nextArchived;
                    saveManualContact({
                        waId: normalizeWaId(item.dataset.waId || ''),
                        profileName: item.dataset.profileName || item.querySelector('.contact-name')?.textContent?.trim() || '',
                        tag: nextArchived,
                        universityName: item.dataset.universityName || '',
                        semester: item.dataset.semester || '',
                        timezone: item.dataset.timezone || getAutoTimezone()
                    });
                    updateContactStateUI(item);
                    scheduleContactOrderRefresh();
                    openContactMenuId = '';
                    item.classList.remove('is-menu-open');
                    item.querySelector('.contact-card-menu')?.classList.add('hidden');
                });
            }

            if (tagSelect) {
                tagSelect.addEventListener('click', (e) => e.stopPropagation());
                tagSelect.addEventListener('change', function(e) {
                    e.stopPropagation();
                    item.dataset.tag = this.value;
                    saveManualContact({
                        waId: normalizeWaId(item.dataset.waId || ''),
                        profileName: item.dataset.profileName || item.querySelector('.contact-name')?.textContent?.trim() || '',
                        tag: this.value,
                        universityName: item.dataset.universityName || '',
                        semester: item.dataset.semester || '',
                        timezone: item.dataset.timezone || getAutoTimezone()
                    });
                    updateContactStateUI(item);
                    applyContactFilter();
                });
            }

            const avatarWrap = item.querySelector('.contact-avatar-wrap');
            if (avatarWrap) {
                avatarWrap.addEventListener('click', (e) => {
                    e.stopPropagation();
                    openClientDetailsModal(item);
                });
            }

            item.addEventListener('click', () => {
                const waId = normalizeWaId(item.dataset.waId || '');
                if (!waId) return;
                if (forwardTargetSelectMode) {
                    void forwardSelectedMessagesToContact(waId);
                    return;
                }
                const name = item.querySelector('.contact-name')?.textContent?.trim() || waId;
                activeWhatsappWaId = waId;
                openContactMenuId = '';
                item.classList.remove('is-menu-open');
                item.querySelector('.contact-card-menu')?.classList.add('hidden');
                setActiveContactItem(item);
                markWhatsappThreadRead(waId);
                updateTaskClientInputFromContact(item);
                applyContactFilter();
                if (chatHeaderTitle) chatHeaderTitle.textContent = name;
                if (chatHeaderAvatar) chatHeaderAvatar.textContent = (name.charAt(0) || '?').toUpperCase();
                updateChatHeaderVisibility(waId);
                renderWhatsappMessages(waId);
                updateChatWindowState(waId);
            });
        }

        function escapeHtml(raw) {
            return String(raw || '')
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;');
        }

        function formatChatTime(ts) {
            const d = ts ? new Date(ts) : new Date();
            if (Number.isNaN(d.getTime())) return '';
            const hh = d.getHours() % 12 || 12;
            const mm = String(d.getMinutes()).padStart(2, '0');
            const ampm = d.getHours() >= 12 ? 'PM' : 'AM';
            return hh + ':' + mm + ' ' + ampm;
        }

        function isMediaMessageType(messageType) {
            return ['image', 'document', 'audio', 'video'].includes(String(messageType || '').toLowerCase());
        }

        function getActiveChatContactName() {
            return (chatHeaderTitle?.textContent || '').trim() || normalizeWaId(activeWhatsappWaId);
        }

        function applyChatInitiationPreview(previewText) {
            if (!chatInput) return;
            const normalizedPreview = String(previewText || '').replace(/\s+/g, ' ').trim();
            if (!chatInput.dataset.defaultPlaceholder) {
                chatInput.dataset.defaultPlaceholder = chatInput.getAttribute('placeholder') || 'Type message here...';
            }
            chatInput.setAttribute('placeholder', normalizedPreview || chatInput.dataset.defaultPlaceholder);
            chatInput.title = String(previewText || '').trim();
        }

        function clearChatInitiationPreview() {
            if (!chatInput) return;
            chatInput.setAttribute('placeholder', chatInput.dataset.defaultPlaceholder || 'Type message here...');
            chatInput.title = '';
        }

        function applyExpertDeadlineConstraint() {
            if (!odActualDeadlineInput || !odExpertDeadlineInput) return;
            const actualDeadline = (odActualDeadlineInput.value || '').trim();
            odExpertDeadlineInput.disabled = !actualDeadline;
            odExpertDeadlineInput.max = actualDeadline || '';
            if (!actualDeadline) {
                odExpertDeadlineInput.value = '';
                return;
            }
            if (odExpertDeadlineInput.value && odExpertDeadlineInput.value > actualDeadline) {
                odExpertDeadlineInput.value = actualDeadline;
            }
        }

        function normalizeMessageStatus(rawStatus) {
            const status = String(rawStatus || '').toLowerCase();
            if (status === 'read') return 'read';
            if (status === 'delivered') return 'delivered';
            if (status === 'failed') return 'failed';
            if (status === 'sent' || status === 'accepted' || status === 'queued' || status === 'pending') return 'sent';
            return 'sent';
        }

        function makeTempMessageId(prefix) {
            return `${prefix || 'msg'}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
        }

        function getWhatsappMessageDedupKey(message) {
            if (!message) return '';
            const id = String(message.id || '').trim();
            if (id) return `id:${id}`;
            return [
                String(message.direction || 'outgoing'),
                String(message.timestamp || ''),
                String(message.messageType || 'text'),
                String(message.text || ''),
                String(message.attachmentName || ''),
                String(message.mediaId || ''),
                String(message.replyTo?.id || '')
            ].join('|');
        }

        function findWhatsappMessageIndex(waId, message) {
            const rows = Array.isArray(whatsappMessagesByContact[waId]) ? whatsappMessagesByContact[waId] : [];
            const key = getWhatsappMessageDedupKey(message);
            if (!key) return -1;
            return rows.findIndex((row) => getWhatsappMessageDedupKey(row) === key);
        }

        function findWhatsappOutgoingMatchIndex(waId, message) {
            const rows = Array.isArray(whatsappMessagesByContact[waId]) ? whatsappMessagesByContact[waId] : [];
            if (!rows.length || String(message?.direction || '') !== 'outgoing') return -1;
            const nextTimestamp = new Date(message?.timestamp || 0).getTime();
            const nextReplyId = String(message?.replyTo?.id || '');
            const nextText = String(message?.text || '').trim();
            const nextType = String(message?.messageType || 'text');
            return rows.findIndex((row) => {
                if (String(row?.direction || '') !== 'outgoing') return false;
                if (String(row?.id || '') === String(message?.id || '')) return false;
                if (String(row?.messageType || 'text') !== nextType) return false;
                if (String(row?.text || '').trim() !== nextText) return false;
                if (String(row?.replyTo?.id || '') !== nextReplyId) return false;
                const rowTimestamp = new Date(row?.timestamp || 0).getTime();
                const tsCloseEnough = nextTimestamp && rowTimestamp
                    ? Math.abs(nextTimestamp - rowTimestamp) <= 60000
                    : true;
                if (!tsCloseEnough) return false;
                const rowId = String(row?.id || '');
                return /^(queued|msg|template|media)_/i.test(rowId) || ['pending', 'sent'].includes(String(row?.status || '').toLowerCase());
            });
        }

        function appendLocalOutgoingMessage(waId, message) {
            if (!waId || !message) return null;
            const normalizedWaId = normalizeWaId(waId);
            if (!Array.isArray(whatsappMessagesByContact[normalizedWaId])) {
                whatsappMessagesByContact[normalizedWaId] = [];
            }
            const normalizedMessage = {
                id: message.id || makeTempMessageId('msg'),
                timestamp: message.timestamp || new Date().toISOString(),
                direction: 'outgoing',
                messageType: 'text',
                status: 'sent',
                statusTimestamp: message.statusTimestamp || new Date().toISOString(),
                ...message
            };
            const existingIndex = findWhatsappMessageIndex(normalizedWaId, normalizedMessage);
            const reconcileIndex = existingIndex >= 0 ? existingIndex : findWhatsappOutgoingMatchIndex(normalizedWaId, normalizedMessage);
            if (reconcileIndex >= 0) {
                whatsappMessagesByContact[normalizedWaId][reconcileIndex] = {
                    ...whatsappMessagesByContact[normalizedWaId][reconcileIndex],
                    ...normalizedMessage
                };
            } else {
                whatsappMessagesByContact[normalizedWaId].push(normalizedMessage);
            }
            if (normalizedMessage.id) {
                whatsappMessageStatusById[normalizedMessage.id] = {
                    status: normalizeMessageStatus(normalizedMessage.status),
                    statusTimestamp: normalizedMessage.statusTimestamp || normalizedMessage.timestamp || new Date().toISOString()
                };
            }
            const contactItem = contactList.querySelector('.contact-item[data-wa-id="' + normalizedWaId + '"]');
            if (contactItem) {
                contactItem.dataset.lastActivity = String(new Date(normalizedMessage.timestamp || Date.now()).getTime() || Date.now());
                updateContactStateUI(contactItem);
                scheduleContactOrderRefresh();
            }
            return normalizedMessage;
        }

        function updateLocalWhatsappMessage(waId, messageId, updates) {
            const normalizedWaId = normalizeWaId(waId);
            const rows = Array.isArray(whatsappMessagesByContact[normalizedWaId]) ? whatsappMessagesByContact[normalizedWaId] : [];
            if (!normalizedWaId || !messageId || !rows.length) return false;
            const index = rows.findIndex((row) => String(row?.id || '') === String(messageId || ''));
            if (index < 0) return false;
            rows[index] = {
                ...rows[index],
                ...updates
            };
            whatsappMessagesByContact[normalizedWaId] = rows;
            const nextId = String(rows[index]?.id || '');
            if (nextId) {
                whatsappMessageStatusById[nextId] = {
                    status: normalizeMessageStatus(rows[index]?.status),
                    statusTimestamp: rows[index]?.statusTimestamp || rows[index]?.timestamp || new Date().toISOString()
                };
            }
            return true;
        }

        async function processWhatsappTextSendQueue() {
            if (whatsappTextSendProcessing) return;
            whatsappTextSendProcessing = true;
            while (whatsappTextSendQueue.length) {
                const nextItem = whatsappTextSendQueue[0];
                try {
                    const res = await whatsappFetch('/api/whatsapp/send', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify({
                            waId: nextItem.waId,
                            text: nextItem.text,
                            contextMessageId: nextItem.contextMessageId || ''
                        })
                    });
                    const data = await res.json();
                    if (!res.ok || !data?.ok) {
                        throw new Error(data?.error || 'Send failed');
                    }
                    updateLocalWhatsappMessage(nextItem.waId, nextItem.tempId, {
                        id: data?.id || nextItem.tempId,
                        status: 'sent',
                        statusTimestamp: new Date().toISOString()
                    });
                } catch (err) {
                    console.error('Failed to send queued WhatsApp message:', err);
                    updateLocalWhatsappMessage(nextItem.waId, nextItem.tempId, {
                        status: 'failed',
                        statusTimestamp: new Date().toISOString()
                    });
                } finally {
                    whatsappTextSendQueue.shift();
                    if (activeWhatsappWaId === nextItem.waId) {
                        renderWhatsappMessages(nextItem.waId);
                    }
                    void loadWhatsappContacts();
                }
            }
            whatsappTextSendProcessing = false;
        }

        function setChatSendState(isBusy, label) {
            whatsappSendInFlight = Boolean(isBusy);
            if (sendBtn) {
                sendBtn.disabled = whatsappSendInFlight;
                sendBtn.classList.toggle('is-loading', whatsappSendInFlight);
                sendBtn.classList.toggle('is-disabled', whatsappSendInFlight);
                sendBtn.setAttribute('aria-busy', whatsappSendInFlight ? 'true' : 'false');
                sendBtn.title = whatsappSendInFlight ? (label || 'Sending...') : 'Send message';
            }
            if (notifyBtn) notifyBtn.disabled = whatsappSendInFlight;
            if (attachFileBtn) attachFileBtn.disabled = whatsappSendInFlight;
            if (chatInput) chatInput.readOnly = whatsappSendInFlight;
        }

        function getStatusTickHtml(rawStatus) {
            const status = normalizeMessageStatus(rawStatus);
            if (status === 'read') {
                return '<span class="chat-status chat-status-read" title="Read"><svg viewBox="0 0 16 12" aria-hidden="true"><path d="M1.5 6.5 4.5 9.5 8.5 2.5"></path><path d="M7.5 6.5 10.5 9.5 14.5 2.5"></path></svg></span>';
            }
            if (status === 'delivered') {
                return '<span class="chat-status chat-status-delivered" title="Delivered"><svg viewBox="0 0 16 12" aria-hidden="true"><path d="M1.5 6.5 4.5 9.5 8.5 2.5"></path><path d="M7.5 6.5 10.5 9.5 14.5 2.5"></path></svg></span>';
            }
            if (status === 'failed') {
                return '<span class="chat-status chat-status-failed" title="Failed">!</span>';
            }
            return '<span class="chat-status chat-status-sent" title="Sent"><svg viewBox="0 0 10 10" aria-hidden="true"><path d="M1.5 5.25 3.6 7.35 8.25 1.75"></path></svg></span>';
        }

        function renderWhatsappMessages(waId) {
            if (!chatMessages) return;
            if (!waId) {
                chatMessages.classList.add('is-empty-chat');
                chatMessages.innerHTML = '<div class="chat-empty-state" aria-hidden="true"></div>';
                updateChatWindowState('');
                return;
            }
            const messages = whatsappMessagesByContact[waId] || [];
            chatMessages.classList.remove('is-empty-chat');
            chatMessages.innerHTML = '';
            if (!messages.length) {
                chatMessages.innerHTML = '<p class="text-sm text-gray-500">No WhatsApp messages yet.</p>';
                updateChatWindowState(waId);
                return;
            }
            messages.forEach((msg) => {
                const clientKey = ensureMessageClientKey(msg);
                const wrap = document.createElement('div');
                const incoming = msg.direction === 'incoming';
                const statusTick = incoming ? '' : getStatusTickHtml(msg.status);
                const messageType = msg.messageType || 'text';
                const isTemplateMessage = String(messageType).toLowerCase() === 'template';
                const attachmentName = escapeHtml(msg.attachmentName || '');
                const attachmentUrl = String(msg.attachmentUrl || '').trim();
                const replyQuote = msg.replyTo?.id
                    ? `
                        <div class="chat-reply-quote">
                            <strong>${escapeHtml(msg.replyTo.direction === 'incoming' ? 'Client' : 'You')}</strong>
                            <span>${escapeHtml(msg.replyTo.text || '(message)')}</span>
                        </div>
                    `
                    : '';
                const bubbleBody = isMediaMessageType(messageType)
                    ? `
                        <div class="chat-bubble-file">
                            <p class="font-semibold capitalize">${escapeHtml(messageType)}</p>
                            ${attachmentName ? `<p>${attachmentName}</p>` : ''}
                            ${attachmentUrl ? `<a href="${escapeHtml(attachmentUrl)}" target="_blank" rel="noopener noreferrer" class="text-xs font-medium text-indigo-700 underline">Open attachment</a>` : ''}
                            ${msg.text ? `<p class="text-xs text-gray-600">${escapeHtml(msg.text)}</p>` : ''}
                        </div>
                    `
                    : isTemplateMessage
                        ? `
                            <div class="chat-template-body">
                                <p class="chat-template-label">Template Message</p>
                                <p class="chat-template-text">${escapeHtml(msg.text || '[Template message]')}</p>
                            </div>
                        `
                    : escapeHtml(msg.text || '[Unsupported message type]');
                const bubbleClass = isTemplateMessage
                    ? 'chat-bubble-template p-2.5 shadow-sm text-sm text-gray-800'
                    : `${incoming ? 'chat-bubble-client' : 'chat-bubble-admin'} p-3 max-w-md shadow-sm text-sm text-gray-800`;
                wrap.className = 'flex flex-col ' + (incoming ? 'items-start' : 'items-end');
                wrap.innerHTML = `
                        <div class="chat-message-row">
                        ${forwardMode ? `<input class="chat-forward-check" type="checkbox" data-forward-id="${escapeHtml(clientKey)}" ${forwardSelectionIds.has(clientKey) ? 'checked' : ''}>` : ''}
                        <button type="button" class="chat-message-menu-btn" data-menu-id="${escapeHtml(clientKey)}">&#9662;</button>
                        ${openMessageMenuId === clientKey ? `
                            <div class="chat-message-menu">
                                <button type="button" data-action="reply" data-message-key="${escapeHtml(clientKey)}">Reply</button>
                                <button type="button" data-action="forward" data-message-key="${escapeHtml(clientKey)}">Forward</button>
                            </div>
                        ` : ''}
                        <div class="${bubbleClass}">
                            ${replyQuote}
                            ${bubbleBody}
                        </div>
                    </div>
                    <span class="chat-message-meta text-[10px] text-gray-400 mt-1 ${incoming ? 'ml-1' : 'mr-1'}">${formatChatTime(msg.timestamp)}${statusTick}</span>
                `;
                chatMessages.appendChild(wrap);
            });
            chatMessages.scrollTop = chatMessages.scrollHeight;
            updateChatWindowState(waId);
        }

        function applyWhatsappMessageStatusUpdate(payload) {
            if (!payload || !payload.waId || !payload.messageId) return;
            const waId = normalizeWaId(payload.waId);
            const rows = whatsappMessagesByContact[waId];
            const status = normalizeMessageStatus(payload.status);
            whatsappMessageStatusById[payload.messageId] = {
                status,
                statusTimestamp: payload.statusTimestamp || new Date().toISOString()
            };
            if (!Array.isArray(rows) || !rows.length) return;
            for (let i = rows.length - 1; i >= 0; i -= 1) {
                if (rows[i]?.id !== payload.messageId) continue;
                rows[i].status = status;
                rows[i].statusTimestamp = payload.statusTimestamp || new Date().toISOString();
                break;
            }
            if (activeWhatsappWaId === waId) {
                renderWhatsappMessages(waId);
            }
        }

        function upsertWhatsappContact(contact, markUnread) {
            if (!contact || !contact.waId) return null;
            const normalizedWaId = normalizeWaId(contact.waId);
            if (!normalizedWaId) return null;
            let existing = contactList.querySelector('.contact-item[data-wa-id="' + normalizedWaId + '"]');
            const existingName = existing?.dataset.profileName || existing?.querySelector('.contact-name')?.textContent?.trim() || '';
            const profileName = getPreferredProfileName(existingName, contact.profileName, normalizedWaId);
            const displayName = getContactDisplayName(profileName, normalizedWaId);
            const initials = getContactInitials(profileName, normalizedWaId);
            const savedContact = readManualContactsFromStorage().find((row) => normalizeWaId(row.waId) === normalizedWaId);
            const nextTag = String(contact.tag || savedContact?.tag || existing?.dataset.tag || '');
            const nextUniversityName = String(contact.universityName || savedContact?.universityName || existing?.dataset.universityName || '');
            const nextSemester = String(contact.semester || savedContact?.semester || existing?.dataset.semester || '');
            const nextTimezone = String(contact.timezone || savedContact?.timezone || existing?.dataset.timezone || getAutoTimezone());
            const contactId = getOrCreateWhatsappContactId(normalizedWaId);
            saveManualContact({ waId: normalizedWaId, profileName, tag: nextTag, universityName: nextUniversityName, semester: nextSemester, timezone: nextTimezone });
            if (!existing) {
                const item = document.createElement('div');
                item.className = 'contact-item p-3 border-b hover:bg-gray-50 cursor-pointer';
                item.dataset.waId = normalizedWaId;
                item.dataset.contactId = contactId;
                item.dataset.source = 'whatsapp';
                item.dataset.tag = nextTag;
                item.dataset.lastActivity = '0';
                item.dataset.unreadCount = '0';
                item.dataset.profileName = profileName;
                item.dataset.universityName = nextUniversityName;
                item.dataset.semester = nextSemester;
                item.dataset.timezone = nextTimezone;
                item.innerHTML = `
                    <div class="contact-card-top">
                        <div class="contact-identity">
                            <div class="contact-avatar-wrap">
                                <div class="contact-avatar" aria-hidden="true">${escapeHtml(initials)}</div>
                                <span class="contact-avatar-badge hidden">0</span>
                            </div>
                            <div class="contact-copy">
                                <span class="contact-name font-semibold text-gray-800 text-sm">${escapeHtml(displayName)}</span>
                                <div class="contact-secondary-row">
                                    <p class="contact-meta text-xs text-gray-500">ID: ${escapeHtml(contactId)}</p>
                                    <span class="contact-label-badge hidden"></span>
                                    <span class="contact-pin-icon hidden" aria-hidden="true">
                                        <svg viewBox="0 0 16 16"><path d="M10.3 1.2c1.4 0 2.5 1.1 2.5 2.5 0 .6-.2 1.1-.5 1.5L10.8 7v2.2l1.1 1.1c.2.2.2.6 0 .8s-.6.2-.8 0L10 10H8.4l-3.8 3.8c-.2.2-.6.2-.8 0s-.2-.6 0-.8L7.6 9.2V7.6L6.5 6.5c-.2-.2-.2-.6 0-.8s.6-.2.8 0l1.1 1.1h2.2l1.5-1.5c.4-.4.5-.9.5-1.5 0-.8-.7-1.5-1.5-1.5S9.6 3 9.6 3.8c0 .2-.1.4-.3.5l-1.5 1c-.3.2-.6.1-.8-.2s-.1-.6.2-.8l1.2-.8c.2-1.3 1.3-2.3 2.6-2.3Z"></path><circle cx="10.3" cy="3.8" r="0.8"></circle></svg>
                                    </span>
                                </div>
                            </div>
                        </div>
                        <div class="contact-side">
                            <span class="contact-last-time"></span>
                            <div class="contact-side-meta">
                                <span class="contact-unread-dot hidden"></span>
                                <span class="contact-unread-count hidden">0</span>
                                <button type="button" class="contact-menu-toggle" aria-label="Open contact options">&#9662;</button>
                            </div>
                        </div>
                    </div>
                    <div class="contact-card-menu hidden">
                        <button type="button" class="contact-action-pin">Pin chat</button>
                        <label class="contact-menu-label-wrap">
                            <span>Label chat</span>
                            <select class="contact-menu-label">
                                <option value="">No label</option>
                                <option value="new">New Client</option>
                                <option value="old">Old Client</option>
                                <option value="tutor">Expert</option>
                                <option value="friend">Client's Friend</option>
                                <option value="useless">Broker</option>
                            </select>
                        </label>
                        <button type="button" class="contact-action-archive">Archive chat</button>
                        <button type="button" class="contact-action-read">Mark as unread</button>
                    </div>
                `;
                contactItems.push(item);
                wireContactItem(item);
                contactList.prepend(item);
                existing = item;
            } else {
                const nameEl = existing.querySelector('.contact-name');
                if (nameEl) nameEl.textContent = displayName;
                const avatarEl = existing.querySelector('.contact-avatar');
                if (avatarEl) avatarEl.textContent = initials;
                existing.dataset.contactId = contactId;
                existing.dataset.tag = nextTag;
                existing.dataset.profileName = profileName;
                existing.dataset.universityName = nextUniversityName;
                existing.dataset.semester = nextSemester;
                existing.dataset.timezone = nextTimezone;
                const metaEl = existing.querySelector('.contact-meta') || existing.querySelector('p');
                if (metaEl) {
                    metaEl.classList.add('contact-meta');
                    metaEl.textContent = 'ID: ' + contactId;
                }
            }
            const tagSelect = existing.querySelector('.contact-menu-label');
            if (tagSelect) tagSelect.value = nextTag;
            if (activeWhatsappWaId && activeWhatsappWaId === normalizedWaId) {
                setActiveContactItem(existing);
            }
            const latestMessageTimestamp = Array.isArray(whatsappMessagesByContact[normalizedWaId]) && whatsappMessagesByContact[normalizedWaId].length
                ? whatsappMessagesByContact[normalizedWaId][whatsappMessagesByContact[normalizedWaId].length - 1]?.timestamp
                : '';
            const latestActivityMs = Math.max(
                new Date(latestMessageTimestamp || 0).getTime() || 0,
                new Date(contact.updatedAt || 0).getTime() || 0
            );
            if (latestActivityMs > 0) {
                existing.dataset.lastActivity = String(latestActivityMs);
            }
            if (activeWhatsappWaId && activeWhatsappWaId === normalizedWaId) {
                markWhatsappThreadRead(normalizedWaId);
            } else if (markUnread) {
                setContactUnreadCount(existing, getUnreadCountForThread(normalizedWaId, whatsappMessagesByContact[normalizedWaId] || []));
            } else {
                setContactUnreadCount(existing, getUnreadCountForThread(normalizedWaId, whatsappMessagesByContact[normalizedWaId] || []));
                updateContactStateUI(existing);
            }
            return existing;
        }

        function upsertWhatsappMessage(payload) {
            if (!payload || !payload.waId) return;
            const waId = normalizeWaId(payload.waId);
            if (!waId) return;
            if (!Array.isArray(whatsappMessagesByContact[waId])) {
                whatsappMessagesByContact[waId] = [];
            }
            const nextMessage = {
                id: payload.id || '',
                text: payload.text || '',
                timestamp: payload.timestamp || new Date().toISOString(),
                direction: payload.direction || 'incoming',
                messageType: payload.messageType || 'text',
                attachmentName: payload.attachmentName || '',
                attachmentUrl: payload.attachmentUrl || '',
                mediaId: payload.mediaId || '',
                mimeType: payload.mimeType || '',
                replyTo: payload.replyTo || null,
                status: normalizeMessageStatus(payload.status || whatsappMessageStatusById[payload.id || '']?.status),
                statusTimestamp: payload.statusTimestamp || whatsappMessageStatusById[payload.id || '']?.statusTimestamp || payload.timestamp || new Date().toISOString()
            };
            const existingIndex = findWhatsappMessageIndex(waId, nextMessage);
            const reconcileIndex = existingIndex >= 0 ? existingIndex : findWhatsappOutgoingMatchIndex(waId, nextMessage);
            if (reconcileIndex >= 0) {
                whatsappMessagesByContact[waId][reconcileIndex] = {
                    ...whatsappMessagesByContact[waId][reconcileIndex],
                    ...nextMessage
                };
            } else {
                whatsappMessagesByContact[waId].push(nextMessage);
            }
            upsertWhatsappContact({
                waId: waId,
                profileName: payload.profileName || waId
            }, payload.direction === 'incoming');
            const contactItem = contactList.querySelector('.contact-item[data-wa-id="' + waId + '"]');
            if (contactItem) {
                contactItem.dataset.lastActivity = String(new Date(payload.timestamp || Date.now()).getTime() || Date.now());
                if (activeWhatsappWaId === waId) {
                    markWhatsappThreadRead(waId);
                } else {
                    setContactUnreadCount(contactItem, getUnreadCountForThread(waId, whatsappMessagesByContact[waId] || []));
                }
                updateContactStateUI(contactItem);
                scheduleContactOrderRefresh();
            }

            if (activeWhatsappWaId === waId) {
                renderWhatsappMessages(waId);
            }
        }

        async function loadWhatsappContacts() {
            try {
                const bases = await getReachableWhatsappApiBases(false);
                if (!bases.length) {
                    setWhatsappStatus('WA: backend unreachable', 'err');
                    setWhatsappLastEvent('Last webhook: -');
                    return;
                }
                const responses = await Promise.all(bases.map(async (base) => {
                    try {
                        const res = await fetch(base + '/api/whatsapp/contacts', mergeFetchOptions(base));
                        if (!res.ok) return [];
                        const data = await res.json();
                        return Array.isArray(data) ? data : [];
                    } catch {
                        return [];
                    }
                }));

                const mergedByWaId = new Map();
                responses.flat().forEach((entry) => {
                    const waId = normalizeWaId(entry.waId);
                    if (!waId) return;
                    const existing = mergedByWaId.get(waId) || {
                        waId,
                        profileName: entry.profileName || waId,
                        messages: []
                    };
                    const nextMessages = Array.isArray(entry.messages) ? entry.messages : [];
                    existing.messages.push(...nextMessages);
                    if (entry.profileName && existing.profileName === waId) {
                        existing.profileName = entry.profileName;
                    }
                    mergedByWaId.set(waId, existing);
                });

                mergedByWaId.forEach((entry, waId) => {
                    entry.messages.forEach((m) => {
                        const messageId = m?.id || '';
                        if (!messageId) return;
                        whatsappMessageStatusById[messageId] = {
                            status: normalizeMessageStatus(m.status),
                            statusTimestamp: m.statusTimestamp || m.timestamp || new Date().toISOString()
                        };
                    });
                    const seen = new Set();
                    whatsappMessagesByContact[waId] = entry.messages
                        .map((m) => ({
                            id: m.id || '',
                            text: m.text || '',
                            timestamp: m.timestamp || new Date().toISOString(),
                            direction: m.direction || 'incoming',
                            messageType: m.messageType || 'text',
                            attachmentName: m.attachmentName || '',
                            attachmentUrl: m.attachmentUrl || '',
                            mediaId: m.mediaId || '',
                            mimeType: m.mimeType || '',
                            replyTo: m.replyTo || null,
                            status: normalizeMessageStatus(m.status || whatsappMessageStatusById[m.id || '']?.status),
                            statusTimestamp: m.statusTimestamp || whatsappMessageStatusById[m.id || '']?.statusTimestamp || m.timestamp || new Date().toISOString()
                        }))
                        .filter((m) => {
                            const key = [m.id, m.direction, m.timestamp, m.text].join('|');
                            if (seen.has(key)) return false;
                            seen.add(key);
                            return true;
                        })
                        .sort((a, b) => new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime());
                    upsertWhatsappContact({
                        waId,
                        profileName: entry.profileName || waId,
                        updatedAt: entry.updatedAt || ''
                    }, false);
                    const item = contactList.querySelector('.contact-item[data-wa-id="' + waId + '"]');
                    if (item) {
                        const nextUnreadCount = activeWhatsappWaId === waId ? 0 : getUnreadCountForThread(waId, whatsappMessagesByContact[waId] || []);
                        setContactUnreadCount(item, nextUnreadCount);
                    }
                });
                scheduleContactOrderRefresh();
                if (activeWhatsappWaId) {
                    renderWhatsappMessages(activeWhatsappWaId);
                }
            } catch (err) {
                console.error('Failed to load WhatsApp contacts:', err);
            }
        }

        function initWhatsappPolling() {
            if (whatsappPollingTimer) return;
            whatsappPollingTimer = setInterval(() => {
                loadWhatsappContacts();
            }, WHATSAPP_POLL_INTERVAL_MS);
        }

        async function pollWhatsappDebug() {
            try {
                const baseUrl = await resolveWhatsappApiBase(false);
                if (!baseUrl) {
                    setWhatsappStatus('WA: backend unreachable', 'err');
                    setWhatsappLastEvent('Last webhook: -');
                    return;
                }
                const wsLabel = whatsappSocketConnected ? 'WS on' : 'WS off';
                setWhatsappStatus(`WA: ${wsLabel} (${baseUrl})`, whatsappSocketConnected ? 'ok' : 'warn');

                const res = await fetch(baseUrl + '/api/whatsapp/debug', mergeFetchOptions(baseUrl));
                if (!res.ok) {
                    setWhatsappLastEvent('Last webhook: -');
                    return;
                }
                const data = await res.json();
                const latest = Array.isArray(data?.recentWebhookEvents) ? data.recentWebhookEvents[0] : null;
                if (!latest) {
                    setWhatsappLastEvent('Last webhook: -');
                    return;
                }
                const dir = latest.direction ? String(latest.direction) : '';
                const waId = latest.waId ? String(latest.waId) : '';
                const t = formatShortTime(latest.receivedAt || latest.timestamp);
                const summary = [t ? `Last webhook: ${t}` : 'Last webhook', dir ? `(${dir})` : '', waId ? `- ${waId}` : '']
                    .filter(Boolean)
                    .join(' ');
                setWhatsappLastEvent(summary);
            } catch (err) {
                console.error('WhatsApp debug poll failed:', err);
            }
        }

        function initWhatsappDebugPolling() {
            if (whatsappDebugTimer) return;
            pollWhatsappDebug();
            whatsappDebugTimer = setInterval(() => {
                pollWhatsappDebug();
            }, WHATSAPP_DEBUG_POLL_INTERVAL_MS);
        }

        async function initWhatsappRealtime() {
            const baseUrl = await resolveWhatsappApiBase(false);
            if (!baseUrl) {
                setTimeout(() => {
                    initWhatsappRealtime();
                }, 2500);
                return;
            }
            const wsUrl = baseUrl.replace(/^http/i, 'ws') + '/ws';
            let socket;
            try {
                socket = new WebSocket(wsUrl);
            } catch (err) {
                console.error('WhatsApp socket init failed:', err);
                setTimeout(() => {
                    initWhatsappRealtime();
                }, 2500);
                return;
            }
            socket.onopen = () => {
                whatsappSocketConnected = true;
                pollWhatsappDebug();
            };
            socket.onmessage = (event) => {
                try {
                    const data = JSON.parse(event.data);
                    if (data.type === 'whatsapp_message') {
                        upsertWhatsappMessage(data.payload || {});
                    } else if (data.type === 'whatsapp_message_status') {
                        applyWhatsappMessageStatusUpdate(data.payload || {});
                    }
                } catch (err) {
                    console.error('Invalid WhatsApp socket payload:', err);
                }
            };
            socket.onclose = () => {
                whatsappSocketConnected = false;
                pollWhatsappDebug();
                setTimeout(() => {
                    initWhatsappRealtime();
                }, 2500);
            };
        }

        async function initWhatsappSync() {
            await resolveWhatsappApiBase(false);
            await loadWhatsappContacts();
            scheduleContactOrderRefresh();
            initWhatsappRealtime();
            initWhatsappPolling();
            initWhatsappDebugPolling();
            initWhatsappWindowTimer();
        }

        contactItems.forEach((item, index) => {
            item.dataset.order = String(index);
            wireContactItem(item);
        });

        migrateWhatsappContactIdsToSixDigitSeries();
        setActiveContactItem(null);
        updateTaskClientInputFromContact(null);
        restoreManualContactsFromStorage();
        scheduleContactOrderRefresh();
        initWhatsappSync();
        void hydrateCrmStateFromBackend().then((loaded) => {
            crmStateHydrationComplete = true;
            if (!loaded) {
                scheduleCrmStateSave();
            } else if (crmStatePendingSave) {
                crmStatePendingSave = false;
                scheduleCrmStateSave();
            }
        });
        if (chatHeaderAvatar) chatHeaderAvatar.textContent = '?';
        updateChatHeaderVisibility('');
        setClientDetailsEditMode(false);

        if (addContactBtn) {
            addContactBtn.onclick = function() {
                const enteredName = (window.prompt('Contact name (optional):', '') || '').trim();
                const enteredNumber = (window.prompt('WhatsApp number with country code:', '') || '').trim();
                const waId = normalizeWaId(enteredNumber);

                if (!waId || waId.length < 8) {
                    alert('Please enter a valid WhatsApp number with country code.');
                    return;
                }

                const profileName = enteredName || waId;
                if (!Array.isArray(whatsappMessagesByContact[waId])) {
                    whatsappMessagesByContact[waId] = [];
                }
                upsertWhatsappContact({ waId, profileName }, false);
                saveManualContact({ waId, profileName });
                refreshContactOrder();
                applyContactFilter();
            };
        }

        if (editClientDetailsBtn) {
            editClientDetailsBtn.onclick = function() {
                setClientDetailsEditMode(true);
            };
        }

        if (closeClientDetailsBtn) {
            closeClientDetailsBtn.onclick = function() {
                clientDetailsModal?.classList.add('hidden');
                activeClientDetailsWaId = '';
                setClientDetailsEditMode(false);
            };
        }

        if (cancelClientDetailsBtn) {
            cancelClientDetailsBtn.onclick = function() {
                const activeItem = activeClientDetailsWaId
                    ? contactList.querySelector('.contact-item[data-wa-id="' + normalizeWaId(activeClientDetailsWaId) + '"]')
                    : null;
                if (activeItem) openClientDetailsModal(activeItem);
            };
        }

        if (saveClientDetailsBtn) {
            saveClientDetailsBtn.onclick = function() {
                const waId = normalizeWaId(activeClientDetailsWaId);
                if (!waId) return;
                const activeItem = contactList.querySelector('.contact-item[data-wa-id="' + waId + '"]');
                const profileName = (clientDetailsNameInput?.value || '').trim() || (activeItem?.dataset.profileName || waId);
                const universityName = (clientDetailsUniversityInput?.value || '').trim();
                const semester = (clientDetailsSemesterInput?.value || '').trim();
                const timezone = (clientDetailsTimezoneInput?.value || '').trim() || getAutoTimezone();
                saveManualContact({
                    waId,
                    profileName,
                    tag: activeItem?.dataset.tag || '',
                    universityName,
                    semester,
                    timezone
                });
                upsertWhatsappContact({
                    waId,
                    profileName,
                    tag: activeItem?.dataset.tag || '',
                    universityName,
                    semester,
                    timezone
                }, false);
                if (activeItem && activeWhatsappWaId === waId) {
                    updateTaskClientInputFromContact(activeItem);
                    if (chatHeaderTitle) chatHeaderTitle.textContent = getContactDisplayName(profileName, waId);
                    if (chatHeaderAvatar) chatHeaderAvatar.textContent = (profileName.charAt(0) || '?').toUpperCase();
                }
                setClientDetailsEditMode(false);
            };
        }

        contactFilter.addEventListener('change', applyContactFilter);
        contactSearch.addEventListener('input', applyContactFilter);
        if (toggleArchivedViewBtn) {
            toggleArchivedViewBtn.addEventListener('click', () => {
                archivedInboxMode = !archivedInboxMode;
                applyContactFilter();
            });
        }
        applyContactFilter();
        restoreOrderListFromStorage();
        rehydrateOrderCardsFromDom();
        if (!localStorage.getItem(ORDER_LIST_STORAGE_KEY)) {
            persistOrderListToStorage();
        }
        orderSearch.addEventListener('input', applyOrderSearchFilter);
        if (orderDeliveryWindow) {
            orderDeliveryWindow.addEventListener('change', applyOrderSearchFilter);
        }
        [orderTabMine, orderTabExpert, orderTabAll].forEach((btn) => {
            if (!btn) return;
            btn.addEventListener('click', () => setOrderTab(btn.dataset.orderTab || 'mine'));
        });
        applyExpertDeadlineConstraint();
        if (odActualDeadlineInput) {
            odActualDeadlineInput.addEventListener('input', applyExpertDeadlineConstraint);
            odActualDeadlineInput.addEventListener('change', applyExpertDeadlineConstraint);
        }
        if (odExpertDeadlineInput) {
            odExpertDeadlineInput.addEventListener('change', applyExpertDeadlineConstraint);
        }
        taskServiceType.addEventListener('change', function() {
            const isLiveSession = this.value.toLowerCase() === 'live session';
            liveSessionCreateFields.classList.toggle('hidden', !isLiveSession);
        });

        // Create Agent Button (if present)
        const createAgentBtn = document.getElementById('createAgentBtn');
        if (createAgentBtn) {
            createAgentBtn.onclick = function() {
                alert('Admin: Create Agent ID');
            };
        }
        // Create Task Modal Open
        document.getElementById('createTaskBtn').onclick = function() {
            const activeContactItem = activeWhatsappWaId
                ? contactList.querySelector('.contact-item[data-wa-id="' + normalizeWaId(activeWhatsappWaId) + '"]')
                : null;
            updateTaskClientInputFromContact(activeContactItem);
            if (taskModalTitle) {
                taskModalTitle.textContent = `Create Task (${agentName})`;
            }
            const isLiveSession = taskServiceType.value.toLowerCase() === 'live session';
            liveSessionCreateFields.classList.toggle('hidden', !isLiveSession);
            if (!isLiveSession) {
                taskSessionStart.value = '';
                taskSessionDuration.value = '';
            }
            document.getElementById('taskModal').classList.remove('hidden');
        };
        // Create Task Modal Close
        document.getElementById('closeTaskModalBtn').onclick = function() {
            document.getElementById('taskModal').classList.add('hidden');
        };
        // Create Task Modal Create
        document.getElementById('createTaskModalBtn').onclick = function() {
            const clientText = (taskClientInput?.value || '').trim();
            const serviceType = document.getElementById('taskServiceType').value;
            const title = serviceType;
            if (!clientText || /^select a contact first$/i.test(clientText)) {
                alert('Select a WhatsApp contact first, then create the task card.');
                return;
            }
            const idMatch = clientText.match(/\((\d+)\)/);
            const clientId = idMatch ? idMatch[1] : '0000';
            const orderId = getNextOrderId(clientId);
            const isLiveSession = serviceType.toLowerCase() === 'live session';
            const sessionStart = (taskSessionStart.value || '').trim();
            const sessionDuration = (taskSessionDuration.value || '').trim();

            const card = document.createElement('div');
            card.className = 'order-card p-3 rounded-lg border shadow-sm';
            card.dataset.serviceType = serviceType;
            card.dataset.clientId = clientId;
            card.dataset.assignedTo = agentName;
            card.dataset.createdBy = agentName;
            card.dataset.labels = '';
            card.dataset.description = '';
            card.dataset.instructions = '';
            card.dataset.topic = '';
            card.dataset.activity = '';
            card.dataset.comments = '[]';
            card.dataset.expertDeadline = '';
            card.dataset.expertPayout = 'INR 0';
            card.dataset.clientPaidCurrency = 'INR';
            card.dataset.clientPaidAmount = '';
            card.dataset.paymentStatusOverride = 'auto';
            card.dataset.sessionStart = sessionStart;
            card.dataset.sessionDuration = sessionDuration;
            card.innerHTML = `
                <div class="flex justify-between items-start mb-2">
                    <span class="order-detail-label text-[10px] font-bold uppercase px-2 py-0.5 rounded">Work Details</span>
                    <div class="flex items-center gap-2">
                        <span class="order-id text-[10px] text-gray-400 font-mono">#${orderId}</span>
                        <button class="open-order-details-btn p-1 rounded bg-indigo-50 text-indigo-700 hover:bg-indigo-100" title="Open Details">
                            <i data-lucide="square-arrow-out-up-right" class="w-3.5 h-3.5"></i>
                        </button>
                    </div>
                </div>
                <h4 class="order-title text-sm font-bold text-gray-800">${title || serviceType || 'Untitled Task'}</h4>
                <div class="order-title-row flex items-center gap-2">
                    <span class="order-service-tag text-[10px] font-semibold px-2 py-0.5 rounded-full bg-indigo-100 text-indigo-700">${serviceType}</span>
                    <span class="order-status text-[10px] font-bold uppercase px-2 py-0.5 rounded">New</span>
                </div>
                <p class="order-client-note text-[11px] text-gray-500 mt-0.5">${clientText.replace(/\(\d+\)/, '').trim()}</p>
                <div class="mt-1 grid grid-cols-2 gap-x-3 gap-y-0.5 text-[11px] text-gray-600">
                    <p>C.A - <span class="order-amount font-semibold text-indigo-600">TBD</span><span class="order-payment-indicator hidden ml-1"></span></p>
                    <p>E.D - <span class="order-expert-deadline deadline-muted font-medium"></span></p>
                    <p>C.D - <span class="order-deadline deadline-muted font-medium"></span></p>
                    <p>By - <span class="order-created-by text-gray-700 font-medium">${agentName}</span></p>
                </div>
                <p class="order-session mt-1 text-[11px] text-gray-600 ${isLiveSession ? '' : 'hidden'}">S.T - ${sessionStart || 'TBD'} | Dur - ${sessionDuration || 'TBD'}m</p>
                <p class="order-comment hidden mt-1 text-[11px] text-gray-700 border-t pt-1">Note: </p>
            `;
            syncOrderServiceTag(card);
            syncOrderPaymentIndicator(card);
            applyCardTypeStyle(card, serviceType);
            applyDeadlineStyles(card);
            applyOrderStatusStyle(card.querySelector('.order-status'), 'New');
            orderList.prepend(card);
            lucide.createIcons();
            persistOrderListToStorage();
            applyOrderSearchFilter();
            if (isLiveSession) {
                taskSessionStart.value = '';
                taskSessionDuration.value = '';
            }
            document.getElementById('taskModal').classList.add('hidden');
        };
        orderList.addEventListener('click', function(e) {
            const btn = e.target.closest('.open-order-details-btn');
            if (!btn) return;

            const card = btn.closest('.order-card');
            if (!card) return;

            activeOrderCard = card;

            const title = card.querySelector('.order-title')?.textContent?.trim() || '';
            const status = card.querySelector('.order-status')?.textContent?.trim() || 'In Progress';
            const cardId = (card.querySelector('.order-id')?.textContent || '').replace('#', '').trim();
            const detailsText = card.querySelector('.order-client-note')?.textContent?.trim() || '';
            const clientIdText = card.dataset.clientId || cardId;
            const amount = card.querySelector('.order-amount')?.textContent?.trim() || '';
            const deadline = card.querySelector('.order-deadline')?.dataset.baseValue || card.querySelector('.order-deadline')?.textContent?.trim() || '';

            document.getElementById('odTitle').value = title;
            document.getElementById('odServiceType').value = card.dataset.serviceType || title;
            document.getElementById('odStatus').value = status;
            document.getElementById('odOrderId').value = cardId;
            document.getElementById('odClientId').value = clientIdText;
            document.getElementById('odAssignedTo').value = card.dataset.assignedTo || '';
            document.getElementById('odCreatedBy').value = card.dataset.createdBy || card.dataset.assignedTo || '';
            document.getElementById('odLabels').value = card.dataset.labels || '';
            document.getElementById('odActualDeadline').value = toDateTimeLocalValue(deadline);
            document.getElementById('odExpertDeadline').value = toDateTimeLocalValue(card.querySelector('.order-expert-deadline')?.dataset.baseValue || card.dataset.expertDeadline || '');
            applyExpertDeadlineConstraint();
            const amountMatch = amount.match(/^([A-Za-z]{3})\s+(.+)$/);
            if (amountMatch) {
                document.getElementById('odBaseCurrency').value = amountMatch[1].toUpperCase();
                document.getElementById('odBaseAmountValue').value = amountMatch[2];
            } else {
                document.getElementById('odBaseCurrency').value = 'INR';
                document.getElementById('odBaseAmountValue').value = amount === 'TBD' ? '' : amount;
            }
            const expertPayoutRaw = (card.dataset.expertPayout || '').replace(/^INR\s*/i, '').trim();
            document.getElementById('odExpertPayout').value = expertPayoutRaw;
            document.getElementById('odAdditionalCharges').value = card.dataset.additionalCharges || '';
            document.getElementById('odClientPaidCurrency').value = card.dataset.clientPaidCurrency || 'INR';
            document.getElementById('odClientPaidAmount').value = card.dataset.clientPaidAmount || '';
            document.getElementById('odPaymentStatusOverride').value = card.dataset.paymentStatusOverride || 'auto';
            document.getElementById('odDescription').value = card.dataset.description || '';
            document.getElementById('odTopic').value = card.dataset.topic || '';
            document.getElementById('odSessionStart').value = card.dataset.sessionStart || '';
            document.getElementById('odSessionDuration').value = card.dataset.sessionDuration || '';
            document.getElementById('odActivity').value = '';
            renderCommentHistory(parseComments(card.dataset.comments || '[]'));

            orderDetailsModal.classList.remove('hidden');
        });

        document.getElementById('saveCommentBtn').onclick = function() {
            const text = document.getElementById('odActivity').value.trim();
            if (!text) return;
            appendCommentToActiveCard(text);
            document.getElementById('odActivity').value = '';
        };

        document.getElementById('saveOrderDetailsBtn').onclick = function() {
            if (!activeOrderCard) return;

            const title = document.getElementById('odTitle').value.trim();
            const serviceType = document.getElementById('odServiceType').value.trim();
            const status = document.getElementById('odStatus').value.trim();
            if (status.toLowerCase() === 'archive') {
                activeOrderCard.remove();
                persistOrderListToStorage();
                applyOrderSearchFilter();
                orderDetailsModal.classList.add('hidden');
                activeOrderCard = null;
                return;
            }
            const orderId = (activeOrderCard.querySelector('.order-id')?.textContent || '').replace('#', '').trim();
            const clientId = (activeOrderCard.dataset.clientId || '').trim();
            const assignedTo = document.getElementById('odAssignedTo').value.trim();
            const createdBy = (activeOrderCard.dataset.createdBy || agentName).trim();
            const labels = document.getElementById('odLabels').value.trim();
            const actualDeadline = document.getElementById('odActualDeadline').value.trim();
            const expertDeadline = document.getElementById('odExpertDeadline').value.trim();
            if (actualDeadline && expertDeadline && expertDeadline > actualDeadline) {
                alert('Expert deadline actual deadline ke baad nahi ho sakta.');
                applyExpertDeadlineConstraint();
                return;
            }
            const expertPayout = document.getElementById('odExpertPayout').value.trim();
            const baseCurrency = document.getElementById('odBaseCurrency').value.trim() || 'INR';
            const baseAmountValue = document.getElementById('odBaseAmountValue').value.trim();
            const baseAmount = baseAmountValue ? (baseCurrency + ' ' + baseAmountValue) : 'TBD';
            const additionalCharges = document.getElementById('odAdditionalCharges').value.trim();
            const clientPaidCurrency = document.getElementById('odClientPaidCurrency').value.trim() || 'INR';
            const clientPaidAmount = document.getElementById('odClientPaidAmount').value.trim();
            const paymentStatusOverride = document.getElementById('odPaymentStatusOverride').value.trim() || 'auto';
            const description = document.getElementById('odDescription').value.trim();
            const topic = document.getElementById('odTopic').value.trim();
            const sessionStart = document.getElementById('odSessionStart').value.trim();
            const sessionDuration = document.getElementById('odSessionDuration').value.trim();

            activeOrderCard.querySelector('.order-title').textContent = title || serviceType || 'Untitled';
            activeOrderCard.dataset.serviceType = serviceType || title || '';
            syncOrderServiceTag(activeOrderCard);
            applyCardTypeStyle(activeOrderCard, activeOrderCard.dataset.serviceType);
            const noteEl = activeOrderCard.querySelector('.order-client-note');
            noteEl.textContent = description || noteEl.textContent;
            activeOrderCard.querySelector('.order-id').textContent = '#' + (orderId || clientId || 'NEW');
            activeOrderCard.querySelector('.order-amount').textContent = baseAmount;
            activeOrderCard.dataset.clientPaidCurrency = clientPaidCurrency;
            activeOrderCard.dataset.clientPaidAmount = clientPaidAmount;
            activeOrderCard.dataset.paymentStatusOverride = paymentStatusOverride;
            syncOrderPaymentIndicator(activeOrderCard);
            activeOrderCard.querySelector('.order-deadline').textContent = formatDateTimeForCard(actualDeadline || '');
            activeOrderCard.querySelector('.order-expert-deadline').textContent = formatDateTimeForCard(expertDeadline || '');
            activeOrderCard.querySelector('.order-deadline').dataset.baseValue = formatDateTimeForCard(actualDeadline || '');
            activeOrderCard.querySelector('.order-expert-deadline').dataset.baseValue = formatDateTimeForCard(expertDeadline || '');
            applyDeadlineStyles(activeOrderCard);
            activeOrderCard.querySelector('.order-created-by').textContent = createdBy || agentName;
            const sessionLine = activeOrderCard.querySelector('.order-session');
            if (sessionLine) {
                const hasSessionData = !!(sessionStart || sessionDuration);
                sessionLine.textContent = 'S.T - ' + (sessionStart || 'TBD') + ' | Dur - ' + (sessionDuration || 'TBD') + 'm';
                sessionLine.classList.toggle('hidden', !hasSessionData);
            }
            const comments = parseComments(activeOrderCard.dataset.comments || '[]');
            renderCommentHistory(comments);
            let commentEl = activeOrderCard.querySelector('.order-comment');
            if (!commentEl) {
                commentEl = document.createElement('p');
                commentEl.className = 'order-comment hidden mt-1 text-[11px] text-gray-700 border-t pt-1';
                activeOrderCard.appendChild(commentEl);
            }
            const latestComment = comments.length ? comments[comments.length - 1] : null;
            if (latestComment) {
                commentEl.textContent = 'Note: ' + latestComment.agent + ' - ' + latestComment.text;
                commentEl.classList.remove('hidden');
            } else {
                commentEl.classList.add('hidden');
            }

            const statusEl = activeOrderCard.querySelector('.order-status');
            statusEl.textContent = status || 'In Progress';
            applyOrderStatusStyle(statusEl, status);

            activeOrderCard.dataset.assignedTo = assignedTo;
            activeOrderCard.dataset.createdBy = createdBy || agentName;
            activeOrderCard.dataset.clientId = clientId;
            activeOrderCard.dataset.labels = labels;
            activeOrderCard.dataset.expertDeadline = expertDeadline;
            activeOrderCard.dataset.expertPayout = expertPayout ? ('INR ' + expertPayout) : '';
            activeOrderCard.dataset.additionalCharges = additionalCharges;
            activeOrderCard.dataset.description = description;
            activeOrderCard.dataset.topic = topic;
            activeOrderCard.dataset.sessionStart = sessionStart;
            activeOrderCard.dataset.sessionDuration = sessionDuration;
            activeOrderCard.dataset.activity = activeOrderCard.dataset.activity || '';
            activeOrderCard.dataset.comments = JSON.stringify(comments);
            document.getElementById('odActivity').value = '';

            persistOrderListToStorage();
            orderDetailsModal.classList.add('hidden');
        };

        document.getElementById('closeOrderDetailsBtn').onclick = function() {
            orderDetailsModal.classList.add('hidden');
        };
        // Attach File Button
        document.getElementById('attachFileBtn').onclick = function() {
            if (!chatFileInput) return;
            chatFileInput.click();
        };
        if (chatFileInput) {
            chatFileInput.addEventListener('change', function() {
                queueAttachments(this.files || []);
                this.value = '';
            });
        }
        if (clearReplyBtn) {
            clearReplyBtn.onclick = function() {
                setReplyContext(null);
            };
        }
        if (cancelForwardBtn) {
            cancelForwardBtn.onclick = function() {
                setForwardMode(false);
            };
        }
        if (confirmForwardBtn) {
            confirmForwardBtn.onclick = function() {
                if (!activeWhatsappWaId || !forwardSelectionIds.size) {
                    alert('Select at least one message to forward.');
                    return;
                }
                forwardTargetSelectMode = true;
                if (chatForwardTargetHintEl) {
                    chatForwardTargetHintEl.textContent = 'Now click a target contact from the left list.';
                }
            };
        }
        if (chatMessages) {
            chatMessages.addEventListener('click', function(event) {
                const menuBtn = event.target.closest('.chat-message-menu-btn');
                if (menuBtn) {
                    const messageKey = menuBtn.dataset.menuId || '';
                    openMessageMenuId = openMessageMenuId === messageKey ? '' : messageKey;
                    renderWhatsappMessages(activeWhatsappWaId);
                    return;
                }

                const actionBtn = event.target.closest('[data-action]');
                if (actionBtn) {
                    const messageKey = actionBtn.dataset.messageKey || '';
                    const thread = whatsappMessagesByContact[activeWhatsappWaId] || [];
                    const selectedMessage = thread.find((msg) => ensureMessageClientKey(msg) === messageKey);
                    openMessageMenuId = '';
                    if (!selectedMessage) {
                        renderWhatsappMessages(activeWhatsappWaId);
                        return;
                    }
                    if (actionBtn.dataset.action === 'reply') {
                        setReplyContext(selectedMessage);
                    } else if (actionBtn.dataset.action === 'forward') {
                        if (!forwardMode) setForwardMode(true);
                        forwardSelectionIds.add(messageKey);
                        if (chatForwardSelectionCountEl) {
                            chatForwardSelectionCountEl.textContent = `${forwardSelectionIds.size} selected`;
                        }
                    }
                    renderWhatsappMessages(activeWhatsappWaId);
                    return;
                }

                const checkbox = event.target.closest('.chat-forward-check');
                if (checkbox) {
                    const messageKey = checkbox.dataset.forwardId || '';
                    if (checkbox.checked) forwardSelectionIds.add(messageKey);
                    else forwardSelectionIds.delete(messageKey);
                    if (chatForwardSelectionCountEl) {
                        chatForwardSelectionCountEl.textContent = `${forwardSelectionIds.size} selected`;
                    }
                    return;
                }

                if (openMessageMenuId) {
                    openMessageMenuId = '';
                    renderWhatsappMessages(activeWhatsappWaId);
                }
            });
        }
        async function fileToDataUrl(file) {
            return await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => resolve(String(reader.result || ''));
                reader.onerror = () => reject(new Error('Failed to read attachment: ' + (file?.name || 'file')));
                reader.readAsDataURL(file);
            });
        }

        async function forwardSelectedMessagesToContact(targetWaId) {
            const normalizedTargetWaId = normalizeWaId(targetWaId);
            if (!normalizedTargetWaId || !forwardSelectionIds.size || !activeWhatsappWaId) return;
            const thread = whatsappMessagesByContact[activeWhatsappWaId] || [];
            const selectedMessages = thread.filter((msg) => forwardSelectionIds.has(ensureMessageClientKey(msg)));
            if (!selectedMessages.length) {
                alert('No valid messages selected.');
                return;
            }
            try {
                const res = await whatsappFetch('/api/whatsapp/forward', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        targetWaId: normalizedTargetWaId,
                        messages: selectedMessages
                    })
                });
                const data = await res.json();
                if (!res.ok || !data?.ok) {
                    throw new Error(data?.error || 'Forward failed');
                }
                setForwardMode(false);
                await loadWhatsappContacts();
                alert('Messages forwarded successfully.');
            } catch (err) {
                console.error('Failed to forward messages:', err);
                alert('Forward failed: ' + (err?.message || 'Unknown error'));
            }
        }

        async function sendPendingAttachments(waId, captionText) {
            const normalizedCaption = String(captionText || '').trim();
            for (const file of pendingAttachments) {
                const messageType = detectAttachmentType(file.name, file.type);
                try {
                    const dataUrl = await fileToDataUrl(file);
                    const res = await whatsappFetch('/api/whatsapp/send-media', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify({
                            waId,
                            fileName: file.name,
                            mimeType: file.type || 'application/octet-stream',
                            messageType,
                            caption: normalizedCaption,
                            contextMessageId: activeReplyContext?.id || '',
                            dataUrl
                        })
                    });
                    const data = await res.json();
                    if (!res.ok || !data?.ok) {
                        throw new Error(data?.error || ('Attachment send failed: ' + file.name));
                    }
                    appendLocalOutgoingMessage(waId, {
                        id: data?.id || makeTempMessageId('media'),
                        text: data?.text || normalizedCaption || file.name,
                        messageType: data?.messageType || messageType,
                        attachmentName: data?.attachmentName || file.name,
                        attachmentUrl: data?.attachmentUrl || '',
                        mediaId: data?.mediaId || '',
                        mimeType: data?.mimeType || file.type || 'application/octet-stream',
                        replyTo: activeReplyContext ? { ...activeReplyContext } : null,
                        status: 'sent'
                    });
                } catch (error) {
                    appendLocalOutgoingMessage(waId, {
                        id: makeTempMessageId('media_failed'),
                        text: normalizedCaption || file.name,
                        messageType,
                        attachmentName: file.name,
                        attachmentUrl: '',
                        mimeType: file.type || 'application/octet-stream',
                        replyTo: activeReplyContext ? { ...activeReplyContext } : null,
                        status: 'failed'
                    });
                    throw error;
                }
            }
            pendingAttachments = [];
            renderAttachmentQueue();
        }

        notifyBtn.onclick = async function() {
            if (!activeWhatsappWaId) {
                alert('Select a WhatsApp contact first.');
                return;
            }
            if (whatsappSendInFlight) return;
            try {
                setChatSendState(true, 'Sending template...');
                const contactName = getActiveChatContactName();
                const previewRes = await whatsappFetch('/api/whatsapp/initiate-preview', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        waId: normalizeWaId(activeWhatsappWaId),
                        contactName,
                        agentName
                    })
                });
                const previewData = await previewRes.json();
                const previewText = String(previewData?.previewText || '').trim();
                applyChatInitiationPreview(previewText);
                const shouldSend = window.confirm('Ye template send hoga:\n\n' + (previewText || '(preview unavailable)') + '\n\nAbhi bhejna hai?');
                if (!shouldSend) {
                    return;
                }
                const res = await whatsappFetch('/api/whatsapp/initiate', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        waId: normalizeWaId(activeWhatsappWaId),
                        contactName,
                        agentName
                    })
                });
                const data = await res.json();
                if (!res.ok || !data?.ok) {
                    throw new Error(data?.error || 'Chat initiation failed');
                }
                appendLocalOutgoingMessage(activeWhatsappWaId, {
                    id: data?.id || makeTempMessageId('template'),
                    text: data?.previewText || data?.text || previewText || 'Chat initiation template sent',
                    messageType: 'template',
                    status: 'sent'
                });
                renderWhatsappMessages(activeWhatsappWaId);
                void loadWhatsappContacts();
            } catch (err) {
                console.error('Failed to send chat initiation:', err);
                alert('Chat initiation failed: ' + (err?.message || 'Unknown error'));
            } finally {
                setChatSendState(false);
                clearChatInitiationPreview();
            }
        };
        // Send Button
        sendBtn.onclick = async function() {
            const text = (chatInput?.value || '').trim();
            if (!text && pendingAttachments.length === 0) return;
            if (!activeWhatsappWaId) {
                alert('Select a WhatsApp contact first.');
                return;
            }
            const windowState = getWhatsappWindowState(whatsappMessagesByContact[normalizeWaId(activeWhatsappWaId)] || []);
            if (windowState.state !== 'open') {
                alert('This chat window is expired. Please use the chat initiation template first, then wait for the client reply before sending normal messages.');
                return;
            }

            const hadAttachments = pendingAttachments.length > 0;
            try {
                const waId = normalizeWaId(activeWhatsappWaId);
                if (hadAttachments) {
                    if (whatsappSendInFlight) return;
                    setChatSendState(true, 'Sending attachment...');
                    await sendPendingAttachments(waId, text);
                } else if (text) {
                    const replyContextSnapshot = activeReplyContext ? { ...activeReplyContext } : null;
                    const tempId = makeTempMessageId('queued');
                    appendLocalOutgoingMessage(activeWhatsappWaId, {
                        id: tempId,
                        text,
                        messageType: 'text',
                        replyTo: replyContextSnapshot,
                        status: 'pending',
                        statusTimestamp: new Date().toISOString()
                    });
                    renderWhatsappMessages(activeWhatsappWaId);
                    chatInput.value = '';
                    setReplyContext(null);
                    whatsappTextSendQueue.push({
                        waId,
                        tempId,
                        text,
                        contextMessageId: replyContextSnapshot?.id || ''
                    });
                    void processWhatsappTextSendQueue();
                    return;
                }
                renderWhatsappMessages(activeWhatsappWaId);
                chatInput.value = '';
                setReplyContext(null);
                void loadWhatsappContacts();
            } catch (err) {
                console.error('Failed to send WhatsApp message:', err);
                const reason = err?.message || 'Unknown send error';
                if (text && activeWhatsappWaId) {
                    appendLocalOutgoingMessage(activeWhatsappWaId, {
                        id: makeTempMessageId('msg_failed'),
                        text,
                        messageType: 'text',
                        replyTo: activeReplyContext ? { ...activeReplyContext } : null,
                        status: 'failed'
                    });
                    renderWhatsappMessages(activeWhatsappWaId);
                }
                alert('Message send failed: ' + reason);
            } finally {
                if (hadAttachments) setChatSendState(false);
            }
        };
        if (chatInput) {
            chatInput.addEventListener('keydown', function(event) {
                if (event.key !== 'Enter' || event.shiftKey) return;
                event.preventDefault();
                sendBtn.click();
            });
        }
        document.addEventListener('click', function(event) {
            if (event.target.closest('.contact-card-menu') || event.target.closest('.contact-menu-toggle')) {
                return;
            }
            if (!openContactMenuId) return;
            openContactMenuId = '';
            contactItems.forEach((item) => {
                item.classList.remove('is-menu-open');
                item.querySelector('.contact-card-menu')?.classList.add('hidden');
            });
        });
        // Logout
        document.getElementById('logoutBtn').onclick = function() {
            localStorage.removeItem('agentName');
            window.location.href = 'login.html';
        };

