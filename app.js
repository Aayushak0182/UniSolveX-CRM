lucide.createIcons();
        const storedAuthSessionRaw = localStorage.getItem('crmAuthSession');
        let storedAuthSession = null;
        try {
            storedAuthSession = storedAuthSessionRaw ? JSON.parse(storedAuthSessionRaw) : null;
        } catch {
            storedAuthSession = null;
        }
        // Show agent name from localStorage
        const hasValidAuthSession = Boolean(
            storedAuthSession
            && typeof storedAuthSession === 'object'
            && String(storedAuthSession?.email || '').trim()
            && String(storedAuthSession?.name || '').trim()
            && ['admin', 'agent'].includes(String(storedAuthSession?.role || '').trim())
        );
        const storedAgentName = String(hasValidAuthSession ? storedAuthSession?.name : '').trim();
        const storedAgentRole = String(hasValidAuthSession ? storedAuthSession?.role : 'agent').trim() || 'agent';
        const storedAgentEmail = String(hasValidAuthSession ? storedAuthSession?.email : '').trim().toLowerCase();
        const isAdminUser = storedAgentRole === 'admin';
        const adminQuickMenuWrap = document.getElementById('adminQuickMenuWrap');
        const adminQuickMenuBtn = document.getElementById('adminQuickMenuBtn');
        const adminQuickMenu = document.getElementById('adminQuickMenu');
        const navExportClientsBtn = document.getElementById('navExportClientsBtn');
        const navExportExpertsBtn = document.getElementById('navExportExpertsBtn');
        const navImportExpertsBtn = document.getElementById('navImportExpertsBtn');
        const navMonthlyDataBtn = document.getElementById('navMonthlyDataBtn');
        const importExpertsFileInput = document.getElementById('importExpertsFileInput');
        if (!hasValidAuthSession || !storedAgentName) {
            localStorage.removeItem('crmAuthSession');
            localStorage.removeItem('crmAuthToken');
            localStorage.removeItem('crmAuthEmail');
            localStorage.removeItem('agentName');
            localStorage.removeItem('agentRole');
            window.location.href = 'login.html';
        }
        const adminServerStatusPanel = document.getElementById('adminServerStatusPanel');
        const AGENT_ROSTER_STORAGE_KEY = 'unisolvex_agent_roster_v1';
        const ACTIVE_AGENT_NAME_STORAGE_KEY = 'unisolvex_active_agent_name_v1';
        let agentName = String(localStorage.getItem(ACTIVE_AGENT_NAME_STORAGE_KEY) || '').trim() || storedAgentName || 'Aayush';
        const agentEmail = storedAgentEmail;
        if (adminServerStatusPanel && storedAgentRole !== 'admin') {
            adminServerStatusPanel.classList.add('hidden');
        }
        const assignAgentModal = document.getElementById('assignAgentModal');
        const assignAgentModalTitle = document.getElementById('assignAgentModalTitle');
        const assignAgentModalSubtitle = document.getElementById('assignAgentModalSubtitle');
        const closeAssignAgentModalBtn = document.getElementById('closeAssignAgentModalBtn');
        const cancelAssignAgentModalBtn = document.getElementById('cancelAssignAgentModalBtn');
        const confirmAssignAgentBtn = document.getElementById('confirmAssignAgentBtn');
        const assignAgentList = document.getElementById('assignAgentList');
        const assignAgentAdminPanel = document.getElementById('assignAgentAdminPanel');
        const assignAgentNameInput = document.getElementById('assignAgentNameInput');
        const addAssignAgentNameBtn = document.getElementById('addAssignAgentNameBtn');
        const themeToggleBtn = document.getElementById('themeToggleBtn');
        const assignActiveChatAgentBtn = document.getElementById('assignActiveChatAgentBtn');
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
        const orderSearchRow = document.getElementById('orderSearchRow');
        const monthlyDataControls = document.getElementById('monthlyDataControls');
        const monthlyDataMonthInput = document.getElementById('monthlyDataMonth');
        const orderTabMine = document.getElementById('orderTabMine');
        const orderTabExpert = document.getElementById('orderTabExpert');
        const orderTabAll = document.getElementById('orderTabAll');
        const orderTabMonthly = document.getElementById('orderTabMonthly');
        const orderDetailsModal = document.getElementById('orderDetailsModal');
        const clientDetailsModal = document.getElementById('clientDetailsModal');
        const exportClientsModal = document.getElementById('exportClientsModal');
        const exportClientsFromDateInput = document.getElementById('exportClientsFromDate');
        const exportClientsToDateInput = document.getElementById('exportClientsToDate');
        const closeExportClientsModalBtn = document.getElementById('closeExportClientsModalBtn');
        const cancelExportClientsBtn = document.getElementById('cancelExportClientsBtn');
        const confirmExportClientsBtn = document.getElementById('confirmExportClientsBtn');
        const expertNotifyModal = document.getElementById('expertNotifyModal');
        const monthlyDataModal = document.getElementById('monthlyDataModal');
        const closeMonthlyDataModalBtn = document.getElementById('closeMonthlyDataModalBtn');
        const monthlyDataMonthModalInput = document.getElementById('monthlyDataMonthModal');
        const monthlyDataModalContent = document.getElementById('monthlyDataModalContent');
        const closeExpertNotifyModalBtn = document.getElementById('closeExpertNotifyModalBtn');
        const cancelExpertNotifyBtn = document.getElementById('cancelExpertNotifyBtn');
        const confirmExpertNotifyBtn = document.getElementById('confirmExpertNotifyBtn');
        const expertNotifyTaskMeta = document.getElementById('expertNotifyTaskMeta');
        const expertNotifySearch = document.getElementById('expertNotifySearch');
        const expertNotifySelectAll = document.getElementById('expertNotifySelectAll');
        const expertNotifyList = document.getElementById('expertNotifyList');
        if (adminQuickMenuWrap && isAdminUser) {
            adminQuickMenuWrap.classList.remove('hidden');
        }
        if (!isAdminUser) {
            orderTabMonthly?.classList.add('hidden');
            monthlyDataControls?.classList.add('hidden');
        }
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
        const odExpertPaymentStatusInput = document.getElementById('odExpertPaymentStatus');
        const openAttachmentPageBtn = document.getElementById('openAttachmentPageBtn');
        const openManageExpertPageBtn = document.getElementById('openManageExpertPageBtn');
        const openRecordTransactionBtn = document.getElementById('openRecordTransactionBtn');
        const manageExpertsModal = document.getElementById('manageExpertsModal');
        const closeManageExpertsModalBtn = document.getElementById('closeManageExpertsModalBtn');
        const manageExpertsMeta = document.getElementById('manageExpertsMeta');
        const manageExpertsList = document.getElementById('manageExpertsList');
        const manageExpertsTabNotified = document.getElementById('manageExpertsTabNotified');
        const manageExpertsTabInterested = document.getElementById('manageExpertsTabInterested');
        const manageExpertsTabAssigned = document.getElementById('manageExpertsTabAssigned');
        const transactionRecordsModal = document.getElementById('transactionRecordsModal');
        const closeTransactionRecordsModalBtn = document.getElementById('closeTransactionRecordsModalBtn');
        const openAddTransactionRecordBtn = document.getElementById('openAddTransactionRecordBtn');
        const transactionRecordsList = document.getElementById('transactionRecordsList');
        const transactionRecordsMeta = document.getElementById('transactionRecordsMeta');
        const addTransactionRecordModal = document.getElementById('addTransactionRecordModal');
        const closeAddTransactionRecordModalBtn = document.getElementById('closeAddTransactionRecordModalBtn');
        const cancelAddTransactionRecordBtn = document.getElementById('cancelAddTransactionRecordBtn');
        const saveTransactionRecordBtn = document.getElementById('saveTransactionRecordBtn');
        const trAgentNameInput = document.getElementById('trAgentName');
        const trTransactionIdInput = document.getElementById('trTransactionId');
        const trAmountPaidInput = document.getElementById('trAmountPaid');
        const trCurrencyInput = document.getElementById('trCurrency');
        const trPaymentMethodInput = document.getElementById('trPaymentMethod');
        const trNoteInput = document.getElementById('trNote');
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
        const chatAiModeHintEl = document.getElementById('chatAiModeHint');
        const chatAssignmentHintEl = document.getElementById('chatAssignmentHint');
        const notifyBtn = document.getElementById('notifyBtn');
        const sendBtn = document.getElementById('sendBtn');
        const attachFileBtn = document.getElementById('attachFileBtn');
        const aiAssistBtn = document.getElementById('aiAssistBtn');
        document.getElementById('agentNameDisplay').textContent = 'Agent: ' + agentName;
        document.getElementById('agentInitialDisplay').textContent = agentName.charAt(0).toUpperCase();
        let activeOrderCard = null;
        let activeNotifyOrderCard = null;
        let activeWhatsappWaId = '';
        const orderSequenceByClient = {};
        const ORDER_LIST_STORAGE_KEY = 'unisolvex_order_list_html_v2';
        const ORDER_DETAILS_DRAFTS_STORAGE_KEY = 'unisolvex_order_details_drafts_v1';
        const MANUAL_CONTACTS_STORAGE_KEY = 'unisolvex_manual_contacts_v1';
        const EXPERTS_STORAGE_KEY = 'unisolvex_experts_v1';
        const WHATSAPP_CONTACT_ID_MAP_KEY = 'unisolvex_whatsapp_contact_id_map_v1';
        const WHATSAPP_CONTACT_ID_SEQUENCE_KEY = 'unisolvex_whatsapp_contact_id_seq_v1';
        const EXPERT_ID_SEQUENCE_KEY = 'unisolvex_expert_id_seq_v1';
        const WHATSAPP_READ_STATE_STORAGE_KEY = 'unisolvex_whatsapp_read_state_v1';
        const WHATSAPP_API_BASE_STORAGE_KEY = 'whatsappApiBase';
        const WHATSAPP_LAST_WORKING_API_BASE_KEY = 'unisolvex_last_working_whatsapp_api_base_v1';
        const CONTACT_PANEL_WIDE_STORAGE_KEY = 'unisolvex_contact_panel_wide_v1';
        const APP_THEME_STORAGE_KEY = 'unisolvex_theme_v1';
        const configuredWhatsappApiBase = localStorage.getItem(WHATSAPP_API_BASE_STORAGE_KEY);
        const initialViewMode = new URLSearchParams(window.location.search).get('view');
        const hasHttpOrigin = /^https?:\/\//i.test(window.location.origin || '');
        const RENDER_WHATSAPP_API_BASE = 'https://unisolvex-crm-backend-ra02.onrender.com';
        const defaultWhatsappApiBase = (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1' || !hasHttpOrigin)
            ? 'http://localhost:3001'
            : RENDER_WHATSAPP_API_BASE;
        const WHATSAPP_API_FALLBACK_TIMEOUT_MS = 2500;
        const WHATSAPP_POLL_INTERVAL_MS = 5000;
        const WHATSAPP_DEBUG_POLL_INTERVAL_MS = 8000;
        const CRM_STATE_POLL_INTERVAL_MS = 5000;
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
        let activeManageExpertsTab = 'notified';
        let experts = [];
        let selectedNotifyExpertIds = new Set();
        let assignAgentModalMode = 'all';
        let assignAgentModalWaId = '';
        let selectedAssignAgentName = '';
        let lastRenderedChatWaId = '';
        const ORDER_DETAILS_DRAFT_FIELD_IDS = [
            'odTitle',
            'odStatus',
            'odAssignedTo',
            'odActualDeadline',
            'odExpertDeadline',
            'odBaseCurrency',
            'odBaseAmountValue',
            'odExpertPayout',
            'odAdditionalCharges',
            'odClientPaidCurrency',
            'odClientPaidAmount',
            'odPaymentStatusOverride',
            'odExpertPaymentStatus',
            'odSessionStart',
            'odSessionDuration',
            'odActivity'
        ];
        let activeClientDetailsWaId = '';
        let clientDetailsEditMode = false;
        let crmStateSaveTimer = null;
        let crmStateHydrationComplete = false;
        let crmStatePendingSave = false;
        let crmStatePollingTimer = null;
        let lastCrmStateUpdatedAt = '';
        const cloudCrmState = {
            orderListHtml: '',
            manualContacts: [],
            experts: [],
            agentRoster: [],
            whatsappContactIdMap: {},
            whatsappReadState: {},
            whatsappContactIdSequence: 100100,
            expertIdSequence: 1
        };
        let crmBackendErrorShown = false;

        function setCrmBackendStatus(message, level) {
            setWhatsappStatus(message, level);
            if (level === 'err') {
                setWhatsappLastEvent(message);
            }
        }

        function handleRequiredCrmBackendError(context, err, options = {}) {
            const detail = err?.message ? String(err.message) : 'Unknown backend error';
            const message = `CRM backend unreachable: ${context}. ${detail}`;
            console.warn(message);
            setCrmBackendStatus(message, 'err');
            if (options.alertUser && !crmBackendErrorShown) {
                crmBackendErrorShown = true;
                window.alert('CRM backend unreachable. Cloud sync is required, so local fallback is disabled until the backend is available.');
            }
            return false;
        }

        function clearCrmBackendErrorState() {
            crmBackendErrorShown = false;
        }

        function normalizeAgentName(value) {
            return String(value || '').replace(/\s+/g, ' ').trim();
        }

        function getAgentKey(value) {
            return normalizeAgentName(value).toLowerCase();
        }

        function readAgentRosterFromStorage() {
            try {
                const source = Array.isArray(cloudCrmState.agentRoster) ? cloudCrmState.agentRoster : [];
                const next = [];
                const seen = new Set();
                source.forEach((entry) => {
                    const name = normalizeAgentName(entry);
                    const key = getAgentKey(name);
                    if (!name || seen.has(key)) return;
                    seen.add(key);
                    next.push(name);
                });
                const fallbackName = normalizeAgentName(storedAgentName || agentName);
                if (fallbackName && !seen.has(getAgentKey(fallbackName))) {
                    next.unshift(fallbackName);
                }
                return next;
            } catch {
                return [normalizeAgentName(storedAgentName || agentName)].filter(Boolean);
            }
        }

        function writeAgentRosterToStorage(names, skipRemote) {
            const next = [];
            const seen = new Set();
            (Array.isArray(names) ? names : []).forEach((entry) => {
                const name = normalizeAgentName(entry);
                const key = getAgentKey(name);
                if (!name || seen.has(key)) return;
                seen.add(key);
                next.push(name);
            });
            if (!next.length) {
                next.push(normalizeAgentName(storedAgentName || agentName || 'Agent'));
            }
            cloudCrmState.agentRoster = next.slice();
            if (!skipRemote) scheduleCrmStateSave();
            return next;
        }

        function ensureAgentRoster() {
            return writeAgentRosterToStorage(readAgentRosterFromStorage(), true);
        }

        function setCurrentAgentName(nextName, skipRemote) {
            const resolved = normalizeAgentName(nextName) || normalizeAgentName(storedAgentName || agentName || 'Agent');
            agentName = resolved;
            localStorage.setItem(ACTIVE_AGENT_NAME_STORAGE_KEY, resolved);
            const roster = ensureAgentRoster();
            if (!roster.some((entry) => getAgentKey(entry) === getAgentKey(resolved))) {
                writeAgentRosterToStorage([...roster, resolved], skipRemote);
            }
            document.getElementById('agentNameDisplay').textContent = 'Agent: ' + resolved;
            document.getElementById('agentInitialDisplay').textContent = resolved.charAt(0).toUpperCase();
            updateChatAssignmentState(activeWhatsappWaId);
        }

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
            const storedOrderListHtml = typeof cloudCrmState.orderListHtml === 'string' ? cloudCrmState.orderListHtml : '';
            if (!storedOrderListHtml) return;
            orderList.innerHTML = storedOrderListHtml;
            lucide.createIcons();
        }

        function persistOrderListToStorage(skipRemote) {
            cloudCrmState.orderListHtml = orderList.innerHTML;
            if (!skipRemote) {
                scheduleCrmStateSave();
            }
        }

        function getStoredWhatsappContactIdSequence() {
            return Number(cloudCrmState.whatsappContactIdSequence || '100100');
        }

        function setWhatsappContactIdSequenceInStorage(value, skipRemote) {
            cloudCrmState.whatsappContactIdSequence = Number.isFinite(Number(value)) ? Number(value) : 100100;
            if (!skipRemote) {
                scheduleCrmStateSave();
            }
        }

        function getCurrentCrmStatePayload() {
            const rawSequence = Number(cloudCrmState.whatsappContactIdSequence || '100100');
            const rawExpertSequence = Number(cloudCrmState.expertIdSequence || '1');
            return {
                orderListHtml: orderList.innerHTML || '',
                manualContacts: readManualContactsFromStorage(),
                experts: readExpertsFromStorage(),
                agentRoster: readAgentRosterFromStorage(),
                whatsappContactIdMap: readWhatsappContactIdMapFromStorage(),
                whatsappReadState: readWhatsappReadStateFromStorage(),
                whatsappContactIdSequence: Number.isFinite(rawSequence) ? rawSequence : 100100,
                expertIdSequence: Number.isFinite(rawExpertSequence) ? rawExpertSequence : 1
            };
        }

        function hasMeaningfulCrmState(state) {
            if (!state || typeof state !== 'object') return false;
            const orderHtml = typeof state.orderListHtml === 'string' ? state.orderListHtml : '';
            const hasOrders = /order-card/.test(orderHtml);
            const hasManualContacts = Array.isArray(state.manualContacts) && state.manualContacts.length > 0;
            const hasExperts = Array.isArray(state.experts) && state.experts.length > 0;
            const hasContactMap = state.whatsappContactIdMap && Object.keys(state.whatsappContactIdMap).length > 0;
            const hasReadState = state.whatsappReadState && Object.keys(state.whatsappReadState).length > 0;
            return Boolean(hasOrders || hasManualContacts || hasExperts || hasContactMap || hasReadState);
        }

        function hasMeaningfulLocalCrmState() {
            return hasMeaningfulCrmState(getCurrentCrmStatePayload());
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
                const data = await res.json().catch(() => ({}));
                lastCrmStateUpdatedAt = String(data?.state?.updatedAt || data?.updatedAt || lastCrmStateUpdatedAt || '');
                clearCrmBackendErrorState();
            } catch (err) {
                handleRequiredCrmBackendError('save failed', err);
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
                ensureOrderAnalyticsMeta(card);
                syncOrderServiceTag(card);
                syncOrderPaymentIndicator(card);
                syncOrderExpertPaymentIndicator(card);
                syncOrderExpertSummary(card);
                syncOrderNotifyButton(card);
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

        function applyCrmStateSnapshot(state, options = {}) {
            const includeOrders = options.includeOrders !== false;
            if (!state || typeof state !== 'object') return false;
            const hasRemoteState =
                typeof state.orderListHtml === 'string' ||
                Array.isArray(state.manualContacts) ||
                (state.whatsappContactIdMap && typeof state.whatsappContactIdMap === 'object') ||
                Number.isFinite(Number(state.whatsappContactIdSequence));

            if (!hasRemoteState) return false;
            if (!hasMeaningfulCrmState(state) && hasMeaningfulLocalCrmState()) {
                console.warn('Skipping empty CRM snapshot so current cloud-backed data is preserved.');
                return false;
            }

            cloudCrmState.orderListHtml = typeof state.orderListHtml === 'string' ? state.orderListHtml : '';
            cloudCrmState.manualContacts = Array.isArray(state.manualContacts) ? state.manualContacts : [];
            cloudCrmState.experts = Array.isArray(state.experts) ? state.experts : [];
            cloudCrmState.whatsappContactIdMap = state.whatsappContactIdMap && typeof state.whatsappContactIdMap === 'object' ? state.whatsappContactIdMap : {};
            cloudCrmState.whatsappReadState = state.whatsappReadState && typeof state.whatsappReadState === 'object' ? state.whatsappReadState : {};
            cloudCrmState.whatsappContactIdSequence = Number.isFinite(Number(state.whatsappContactIdSequence)) ? Number(state.whatsappContactIdSequence) : 100100;
            cloudCrmState.expertIdSequence = Number.isFinite(Number(state.expertIdSequence)) ? Number(state.expertIdSequence) : 1;
            if (includeOrders) {
                orderList.innerHTML = '';
                restoreOrderListFromStorage();
                rehydrateOrderCardsFromDom();
            }
            writeAgentRosterToStorage(Array.isArray(state.agentRoster) ? state.agentRoster : [], true);
            setWhatsappContactIdSequenceInStorage(
                Number.isFinite(Number(state.whatsappContactIdSequence)) ? Number(state.whatsappContactIdSequence) : 100100,
                true
            );
            cloudCrmState.expertIdSequence = Number.isFinite(Number(state.expertIdSequence)) ? Number(state.expertIdSequence) : 1;
            experts = readExpertsFromStorage();

            restoreManualContactsFromStorage();
            setCurrentAgentName(localStorage.getItem(ACTIVE_AGENT_NAME_STORAGE_KEY) || agentName, true);
            if (activeWhatsappWaId) {
                updateChatAssignmentState(activeWhatsappWaId);
            }
            scheduleContactOrderRefresh();
            lastCrmStateUpdatedAt = String(state.updatedAt || lastCrmStateUpdatedAt || '');
            clearCrmBackendErrorState();
            return true;
        }

        async function hydrateCrmStateFromBackend() {
            try {
                const res = await whatsappFetch('/api/crm/state');
                if (!res.ok) {
                    throw new Error('CRM state load failed with status ' + res.status);
                }
                const data = await res.json();
                const state = data?.state && typeof data.state === 'object'
                    ? data.state
                    : (data && typeof data === 'object' ? data : null);
                return applyCrmStateSnapshot(state, { includeOrders: true });
            } catch (err) {
                return handleRequiredCrmBackendError('initial load failed', err, { alertUser: true });
            }
        }

        async function syncCrmStateFromBackend() {
            try {
                const res = await whatsappFetch('/api/crm/state');
                if (!res.ok) return false;
                const data = await res.json();
                const state = data?.state && typeof data.state === 'object'
                    ? data.state
                    : (data && typeof data === 'object' ? data : null);
                if (!state) return false;
                const nextUpdatedAt = String(state.updatedAt || '').trim();
                const taskModalEl = document.getElementById('taskModal');
                const isEditingTask = Boolean(orderDetailsModal && !orderDetailsModal.classList.contains('hidden'));
                const isCreatingTask = Boolean(taskModalEl && !taskModalEl.classList.contains('hidden'));
                if (nextUpdatedAt && lastCrmStateUpdatedAt && nextUpdatedAt === lastCrmStateUpdatedAt) {
                    return false;
                }
                return applyCrmStateSnapshot(state, { includeOrders: !(isEditingTask || isCreatingTask) });
            } catch (err) {
                return handleRequiredCrmBackendError('background sync failed', err);
            }
        }

        function initCrmStatePolling() {
            if (crmStatePollingTimer) return;
            crmStatePollingTimer = setInterval(() => {
                syncCrmStateFromBackend();
            }, CRM_STATE_POLL_INTERVAL_MS);
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

        function getDefaultMonthlyDataMonth() {
            const now = new Date();
            return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
        }

        function formatDashboardNumber(value) {
            const numeric = Number(value || 0);
            return new Intl.NumberFormat('en-IN', {
                maximumFractionDigits: 0
            }).format(Number.isFinite(numeric) ? numeric : 0);
        }

        function formatDashboardCurrency(value) {
            return `INR ${formatDashboardNumber(value)}`;
        }

        function getNormalizedOrderStatus(card) {
            const raw = String(
                card?.dataset?.orderStatus
                || card?.querySelector('.order-status')?.textContent
                || ''
            ).trim().toLowerCase();
            if (raw === 'inprogress') return 'in progress';
            if (raw === 'cancel') return 'cancelled';
            return raw || 'new';
        }

        function ensureOrderAnalyticsMeta(card) {
            if (!card || !card.classList.contains('order-card') || card.classList.contains('expert-card')) return;
            const nowIso = new Date().toISOString();
            const fallbackDate =
                card.dataset.updatedAt
                || card.dataset.createdAt
                || card.querySelector('.order-deadline')?.dataset.baseValue
                || card.querySelector('.order-expert-deadline')?.dataset.baseValue
                || nowIso;
            if (!card.dataset.createdAt) {
                const parsed = parseCardDateTime(fallbackDate);
                card.dataset.createdAt = parsed ? parsed.toISOString() : nowIso;
            }
            if (!card.dataset.updatedAt) {
                card.dataset.updatedAt = card.dataset.createdAt || nowIso;
            }
        }

        function getOrderAnalyticsDate(card) {
            ensureOrderAnalyticsMeta(card);
            const candidates = [
                card?.dataset?.updatedAt,
                card?.dataset?.createdAt,
                card?.querySelector('.order-deadline')?.dataset.baseValue,
                card?.querySelector('.order-expert-deadline')?.dataset.baseValue
            ];
            for (const candidate of candidates) {
                const parsed = parseCardDateTime(candidate) || new Date(candidate || '');
                if (parsed instanceof Date && !Number.isNaN(parsed.getTime())) {
                    return parsed;
                }
            }
            return null;
        }

        function getStoredOrderCardsSnapshot() {
            const html = typeof cloudCrmState.orderListHtml === 'string' ? cloudCrmState.orderListHtml : '';
            const temp = document.createElement('div');
            temp.innerHTML = html;
            return Array.from(temp.querySelectorAll('.order-card')).filter((card) => !card.classList.contains('expert-card'));
        }

        function renderMonthlyMetricCard(label, value, tone, helper) {
            return `
                <article class="monthly-metric-card tone-${tone}">
                    <p class="monthly-metric-label">${escapeHtml(label)}</p>
                    <p class="monthly-metric-value">${escapeHtml(value)}</p>
                    <p class="monthly-metric-helper">${escapeHtml(helper)}</p>
                </article>
            `;
        }

        function renderMonthlyStatusChart(statusCounts) {
            const statusConfig = [
                { key: 'completed', label: 'Completed', tone: 'success' },
                { key: 'in progress', label: 'In Progress', tone: 'info' },
                { key: 'cancelled', label: 'Cancelled', tone: 'danger' },
                { key: 'refund', label: 'Refund', tone: 'warning' },
                { key: 'on hold', label: 'On Hold', tone: 'muted' },
                { key: 'new', label: 'New', tone: 'accent' }
            ];
            const maxCount = Math.max(1, ...statusConfig.map((item) => Number(statusCounts[item.key] || 0)));
            return `
                <section class="monthly-chart-card">
                    <div class="monthly-chart-head">
                        <div>
                            <h4 class="monthly-chart-title">Task Status Overview</h4>
                            <p class="monthly-chart-subtitle">Separate count for each order status.</p>
                        </div>
                    </div>
                    <div class="monthly-status-chart">
                        ${statusConfig.map((item) => {
                            const count = Number(statusCounts[item.key] || 0);
                            const width = Math.max(6, Math.round((count / maxCount) * 100));
                            return `
                                <div class="monthly-status-row">
                                    <div class="monthly-status-meta">
                                        <span class="monthly-status-label">${escapeHtml(item.label)}</span>
                                        <span class="monthly-status-count">${escapeHtml(String(count))}</span>
                                    </div>
                                    <div class="monthly-status-track">
                                        <span class="monthly-status-fill tone-${item.tone}" style="width:${width}%"></span>
                                    </div>
                                </div>
                            `;
                        }).join('')}
                    </div>
                </section>
            `;
        }

        function renderMonthlyFinanceChart(title, subtitle, weeklyValues, tone, formatter) {
            const maxValue = Math.max(1, ...weeklyValues.map((item) => Math.abs(Number(item.value || 0))));
            return `
                <section class="monthly-chart-card">
                    <div class="monthly-chart-head">
                        <div>
                            <h4 class="monthly-chart-title">${escapeHtml(title)}</h4>
                            <p class="monthly-chart-subtitle">${escapeHtml(subtitle)}</p>
                        </div>
                    </div>
                    <div class="monthly-bars">
                        ${weeklyValues.map((item) => {
                            const numericValue = Number(item.value || 0);
                            const height = Math.max(10, Math.round((Math.abs(numericValue) / maxValue) * 120));
                            return `
                                <div class="monthly-bar-col">
                                    <span class="monthly-bar-value">${escapeHtml(formatter(numericValue))}</span>
                                    <div class="monthly-bar-track">
                                        <span class="monthly-bar-fill tone-${tone}${numericValue < 0 ? ' is-negative' : ''}" style="height:${height}px"></span>
                                    </div>
                                    <span class="monthly-bar-label">${escapeHtml(item.label)}</span>
                                </div>
                            `;
                        }).join('')}
                    </div>
                </section>
            `;
        }

        function buildMonthlyDataDashboardHtml(selectedMonthInput) {
            const selectedMonth = (selectedMonthInput || getDefaultMonthlyDataMonth()).trim();
            const rows = getStoredOrderCardsSnapshot();
            const monthCards = rows.filter((card) => {
                const date = getOrderAnalyticsDate(card);
                if (!date) return false;
                const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
                return monthKey === selectedMonth;
            });

            if (!monthCards.length) {
                return `
                    <section class="monthly-data-view">
                        <div class="monthly-empty-state">
                            <h3>No monthly data found</h3>
                            <p>No task activity is available for ${escapeHtml(selectedMonth)} yet.</p>
                        </div>
                    </section>
                `;
            }

            const statusCounts = {
                completed: 0,
                'in progress': 0,
                cancelled: 0,
                refund: 0,
                'on hold': 0,
                new: 0
            };
            const weeklyRevenue = [0, 0, 0, 0, 0];
            const weeklyExpense = [0, 0, 0, 0, 0];
            const weeklyProfit = [0, 0, 0, 0, 0];

            let totalRevenue = 0;
            let totalTutorExpense = 0;
            let completedTasks = 0;
            let inProgressTasks = 0;
            let cancelledTasks = 0;
            let refundTasks = 0;

            monthCards.forEach((card) => {
                const status = getNormalizedOrderStatus(card);
                const analyticsDate = getOrderAnalyticsDate(card) || new Date();
                const weekIndex = Math.min(4, Math.max(0, Math.floor((analyticsDate.getDate() - 1) / 7)));
                const orderAmount = parseAmountToNumber(card.querySelector('.order-amount')?.textContent || '');
                const tutorExpense = parseAmountToNumber(card.dataset.expertPayout || card.querySelector('.order-expert-pay')?.textContent || '');

                if (statusCounts[status] === undefined) {
                    statusCounts[status] = 0;
                }
                statusCounts[status] += 1;

                if (status === 'completed') completedTasks += 1;
                if (status === 'in progress') inProgressTasks += 1;
                if (status === 'cancelled') cancelledTasks += 1;
                if (status === 'refund') refundTasks += 1;

                const includeFinancials = status !== 'cancelled' && status !== 'refund';
                const revenueValue = includeFinancials ? orderAmount : 0;
                const expenseValue = includeFinancials ? tutorExpense : 0;
                const profitValue = revenueValue - expenseValue;

                totalRevenue += revenueValue;
                totalTutorExpense += expenseValue;
                weeklyRevenue[weekIndex] += revenueValue;
                weeklyExpense[weekIndex] += expenseValue;
                weeklyProfit[weekIndex] += profitValue;
            });

            const totalProfit = totalRevenue - totalTutorExpense;
            const totalTasks = monthCards.length;
            const [year, month] = selectedMonth.split('-');
            const monthLabel = new Date(Number(year), Math.max(0, Number(month) - 1), 1).toLocaleDateString('en-IN', {
                month: 'long',
                year: 'numeric'
            });
            const weeklyLabels = ['Week 1', 'Week 2', 'Week 3', 'Week 4', 'Week 5'];

            return `
                <section class="monthly-data-view">
                    <div class="monthly-data-hero">
                        <div>
                            <p class="monthly-data-kicker">Monthly Data</p>
                            <h3 class="monthly-data-title">${escapeHtml(monthLabel)} Dashboard</h3>
                            <p class="monthly-data-subtitle">Separate financial and status charts for the selected month.</p>
                        </div>
                        <div class="monthly-data-badge">${escapeHtml(String(totalTasks))} tasks</div>
                    </div>
                    <div class="monthly-metrics-grid">
                        ${renderMonthlyMetricCard('Completed Tasks', String(completedTasks), 'success', `${totalTasks} total tasks this month`)}
                        ${renderMonthlyMetricCard('Revenue', formatDashboardCurrency(totalRevenue), 'info', 'Excludes cancelled and refund orders')}
                        ${renderMonthlyMetricCard('Tutor Expense', formatDashboardCurrency(totalTutorExpense), 'warning', 'Based on expert payout values')}
                        ${renderMonthlyMetricCard('Profit', formatDashboardCurrency(totalProfit), totalProfit >= 0 ? 'success' : 'danger', 'Revenue minus tutor expense')}
                        ${renderMonthlyMetricCard('In Progress', String(inProgressTasks), 'accent', 'Tasks currently active')}
                        ${renderMonthlyMetricCard('Cancelled / Refund', `${cancelledTasks} / ${refundTasks}`, 'muted', 'Status dashboard overview')}
                    </div>
                    <div class="monthly-charts-grid">
                        ${renderMonthlyStatusChart(statusCounts)}
                        ${renderMonthlyFinanceChart('Revenue Chart', 'Weekly booked revenue trend.', weeklyLabels.map((label, index) => ({ label, value: weeklyRevenue[index] })), 'info', formatDashboardCurrency)}
                        ${renderMonthlyFinanceChart('Tutor Expense Chart', 'Weekly expert payout trend.', weeklyLabels.map((label, index) => ({ label, value: weeklyExpense[index] })), 'warning', formatDashboardCurrency)}
                        ${renderMonthlyFinanceChart('Profit Chart', 'Weekly profit trend.', weeklyLabels.map((label, index) => ({ label, value: weeklyProfit[index] })), totalProfit >= 0 ? 'success' : 'danger', formatDashboardCurrency)}
                    </div>
                </section>
            `;
        }

        function renderMonthlyDataDashboard() {
            if (!orderList || activeOrderTab !== 'monthly') return;
            const selectedMonth = (monthlyDataMonthInput?.value || getDefaultMonthlyDataMonth()).trim();
            if (monthlyDataMonthInput && !monthlyDataMonthInput.value) {
                monthlyDataMonthInput.value = selectedMonth;
            }
            orderList.innerHTML = buildMonthlyDataDashboardHtml(selectedMonth);
        }

        function renderMonthlyDataModalContent() {
            if (!monthlyDataModalContent) return;
            const selectedMonth = (monthlyDataMonthModalInput?.value || getDefaultMonthlyDataMonth()).trim();
            if (monthlyDataMonthModalInput && !monthlyDataMonthModalInput.value) {
                monthlyDataMonthModalInput.value = selectedMonth;
            }
            monthlyDataModalContent.innerHTML = buildMonthlyDataDashboardHtml(selectedMonth);
        }

        function openMonthlyDataModal() {
            if (!monthlyDataModal || !isAdminUser) return;
            if (monthlyDataMonthModalInput && !monthlyDataMonthModalInput.value) {
                monthlyDataMonthModalInput.value = monthlyDataMonthInput?.value || getDefaultMonthlyDataMonth();
            }
            renderMonthlyDataModalContent();
            monthlyDataModal.classList.remove('hidden');
        }

        function closeMonthlyDataModal() {
            monthlyDataModal?.classList.add('hidden');
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
            if (activeOrderTab === 'expert') {
                renderExpertList();
                return;
            }
            if (activeOrderTab === 'monthly') {
                renderMonthlyDataDashboard();
                return;
            }
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

                card.style.display = matchesQuery && matchesWindow && matchesTab ? '' : 'none';
            });
        }

        function setOrderTab(tab) {
            const allowedTabs = isAdminUser
                ? ['mine', 'expert', 'all', 'monthly']
                : ['mine', 'expert', 'all'];
            activeOrderTab = allowedTabs.includes(tab) ? tab : 'mine';
            [orderTabMine, orderTabExpert, orderTabAll, orderTabMonthly].forEach((btn) => {
                if (!btn) return;
                btn.classList.toggle('is-active', btn.dataset.orderTab === activeOrderTab);
            });
            if (orderSearchRow) {
                orderSearchRow.classList.toggle('hidden', activeOrderTab === 'monthly');
            }
            if (monthlyDataControls) {
                monthlyDataControls.classList.toggle('hidden', activeOrderTab !== 'monthly' || !isAdminUser);
            }
            if (orderSearch) {
                orderSearch.placeholder = activeOrderTab === 'expert' ? 'Search experts...' : 'Search for orders...';
            }
            if (activeOrderTab === 'monthly') {
                renderMonthlyDataDashboard();
                return;
            }
            if (activeOrderTab === 'expert') {
                renderExpertList();
                return;
            }
            restoreOrderListFromStorage();
            rehydrateOrderCardsFromDom();
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

        function updateAiAssistButtonState(item) {
            if (!aiAssistBtn) return;
            const activeItem = item || (activeWhatsappWaId
                ? contactList.querySelector('.contact-item[data-wa-id="' + normalizeWaId(activeWhatsappWaId) + '"]')
                : null);
            const mode = String(activeItem?.dataset?.aiMode || 'human');
            aiAssistBtn.disabled = !activeItem;
            aiAssistBtn.classList.toggle('is-active', Boolean(activeItem) && mode === 'ai');
            if (chatAiModeHintEl) {
                chatAiModeHintEl.classList.toggle('hidden', !(activeItem && mode === 'ai'));
            }
            aiAssistBtn.title = activeItem
                ? (mode === 'ai' ? 'AI is active for this chat' : 'Transfer this chat to AI')
                : 'Select a contact first';
            updateChatAssignmentState(activeItem?.dataset?.waId || activeWhatsappWaId || '');
        }

        async function setContactAiMode(waId, mode) {
            const normalizedWaId = normalizeWaId(waId);
            if (!normalizedWaId || !['ai', 'human'].includes(String(mode || '').trim().toLowerCase())) {
                throw new Error('Valid waId and mode are required');
            }
            const res = await whatsappFetch('/api/ai-agent/contact-mode', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    waId: normalizedWaId,
                    mode
                })
            });
            const data = await res.json();
            if (!res.ok || !data?.ok) {
                throw new Error(data?.error || 'Failed to update contact mode');
            }
            const activeItem = contactList.querySelector('.contact-item[data-wa-id="' + normalizedWaId + '"]');
            if (activeItem) {
                activeItem.dataset.aiMode = mode;
                updateAiAssistButtonState(activeItem);
            }
            if (activeWhatsappWaId === normalizedWaId) {
                updateChatAssignmentState(normalizedWaId);
            }
            return data;
        }

        async function deleteWhatsappContact(waId) {
            const normalizedWaId = normalizeWaId(waId);
            if (!normalizedWaId) {
                throw new Error('Valid waId is required');
            }
            const res = await whatsappFetch('/api/whatsapp/contact/' + encodeURIComponent(normalizedWaId), {
                method: 'DELETE'
            });
            const data = await res.json().catch(() => ({}));
            if (!res.ok || !data?.ok) {
                throw new Error(data?.error || 'Failed to delete contact');
            }
            return data;
        }

        function removeContactFromUi(waId) {
            const normalizedWaId = normalizeWaId(waId);
            if (!normalizedWaId) return;

            const item = contactList.querySelector('.contact-item[data-wa-id="' + normalizedWaId + '"]');
            if (item) {
                const itemIndex = contactItems.indexOf(item);
                if (itemIndex >= 0) contactItems.splice(itemIndex, 1);
                item.remove();
            }

            delete whatsappMessagesByContact[normalizedWaId];

            removeManualContact(normalizedWaId);
            removeWhatsappContactMeta(normalizedWaId);

            if (activeWhatsappWaId === normalizedWaId) {
                activeWhatsappWaId = '';
                setActiveContactItem(null);
                updateChatHeaderVisibility('');
                updateTaskClientInputFromContact(null);
                renderWhatsappMessages('');
                updateAiAssistButtonState(null);
                updateChatAssignmentState('');
                clearChatInitiationPreview();
            }

            openContactMenuId = '';
            applyContactFilter();
            scheduleContactOrderRefresh();
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

        function toDateInputValue(date) {
            if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        }

        function parseDateInputStart(value) {
            const text = String(value || '').trim();
            if (!text) return null;
            const date = new Date(`${text}T00:00:00`);
            return Number.isNaN(date.getTime()) ? null : date.getTime();
        }

        function parseDateInputEnd(value) {
            const text = String(value || '').trim();
            if (!text) return null;
            const date = new Date(`${text}T23:59:59.999`);
            return Number.isNaN(date.getTime()) ? null : date.getTime();
        }

        function formatExportDateTime(timestampMs) {
            const ts = Number(timestampMs || 0);
            if (!ts) return '';
            const date = new Date(ts);
            if (Number.isNaN(date.getTime())) return '';
            return date.toLocaleString();
        }

        function escapeCsvValue(value) {
            const text = String(value ?? '');
            if (/[",\n]/.test(text)) {
                return `"${text.replace(/"/g, '""')}"`;
            }
            return text;
        }

        function buildClientExportRows(fromMs, toMs) {
            return contactItems
                .map((item) => {
                    const waId = normalizeWaId(item?.dataset?.waId || '');
                    if (!waId) return null;
                    const messages = Array.isArray(whatsappMessagesByContact[waId]) ? whatsappMessagesByContact[waId] : [];
                    const activityTimestamps = messages
                        .map((message) => new Date(message?.timestamp || 0).getTime())
                        .filter((value) => Number.isFinite(value) && value > 0);
                    const fallbackActivity = Number(item?.dataset?.lastActivity || 0) || 0;
                    if (!activityTimestamps.length && fallbackActivity > 0) {
                        activityTimestamps.push(fallbackActivity);
                    }
                    const firstActivityAt = activityTimestamps.length ? Math.min(...activityTimestamps) : 0;
                    const lastActivityAt = activityTimestamps.length ? Math.max(...activityTimestamps) : 0;
                    const matchesRange = (fromMs == null && toMs == null)
                        ? true
                        : activityTimestamps.some((ts) => (fromMs == null || ts >= fromMs) && (toMs == null || ts <= toMs));

                    if (!matchesRange) return null;

                    return {
                        clientName: item.dataset.profileName || item.querySelector('.contact-name')?.textContent?.trim() || waId,
                        clientId: String(item.dataset.contactId || '').trim(),
                        whatsappNumber: waId,
                        firstActivityAt,
                        lastActivityAt
                    };
                })
                .filter(Boolean)
                .sort((a, b) => Number(b.lastActivityAt || 0) - Number(a.lastActivityAt || 0));
        }

        function downloadClientExportCsv(rows, fromValue, toValue) {
            const header = ['Client Name', 'Client ID', 'WhatsApp Number', 'First Activity', 'Last Activity'];
            const lines = [header.map(escapeCsvValue).join(',')];
            rows.forEach((row) => {
                lines.push([
                    row.clientName,
                    row.clientId,
                    row.whatsappNumber,
                    formatExportDateTime(row.firstActivityAt),
                    formatExportDateTime(row.lastActivityAt)
                ].map(escapeCsvValue).join(','));
            });
            const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const anchor = document.createElement('a');
            const fromPart = String(fromValue || '').trim() || 'all';
            const toPart = String(toValue || '').trim() || 'latest';
            anchor.href = url;
            anchor.download = `unisolvex-clients-${fromPart}-to-${toPart}.csv`;
            document.body.appendChild(anchor);
            anchor.click();
            anchor.remove();
            URL.revokeObjectURL(url);
        }

        function openExportClientsModal() {
            if (!exportClientsModal) return;
            const today = new Date();
            const thirtyDaysAgo = new Date(today);
            thirtyDaysAgo.setDate(today.getDate() - 30);
            if (exportClientsFromDateInput && !exportClientsFromDateInput.value) {
                exportClientsFromDateInput.value = toDateInputValue(thirtyDaysAgo);
            }
            if (exportClientsToDateInput && !exportClientsToDateInput.value) {
                exportClientsToDateInput.value = toDateInputValue(today);
            }
            exportClientsModal.classList.remove('hidden');
        }

        function closeExportClientsModal() {
            exportClientsModal?.classList.add('hidden');
        }

        function getManualContactByWaId(waId) {
            const normalizedWaId = normalizeWaId(waId);
            return readManualContactsFromStorage().find((row) => normalizeWaId(row.waId) === normalizedWaId) || null;
        }

        function getAssignedAgentMetaByWaId(waId) {
            const saved = getManualContactByWaId(waId);
            return {
                name: String(saved?.assignedAgent || '').trim(),
                key: String(saved?.assignedAgentKey || '').trim().toLowerCase(),
                email: String(saved?.assignedAgentEmail || '').trim().toLowerCase()
            };
        }

        function canCurrentAgentMessageContact(waId) {
            const assigned = getAssignedAgentMetaByWaId(waId);
            const currentKey = getAgentKey(agentName);
            if (assigned.key) return assigned.key === currentKey;
            if (assigned.name) return getAgentKey(assigned.name) === currentKey;
            if (assigned.email) return assigned.email === agentEmail;
            return false;
        }

        function syncContactAssignmentBadge(item) {
            if (!item) return;
            let badge = item.querySelector('.contact-assigned-badge');
            badge?.remove();
        }

        function updateChatAssignmentState(waId) {
            const assigned = getAssignedAgentMetaByWaId(waId);
            const hasThread = Boolean(waId);
            const canMessage = hasThread && canCurrentAgentMessageContact(waId) && !whatsappSendInFlight;
            const isAssigned = Boolean(assigned.name || assigned.key || assigned.email);
            const hasAnyContacts = contactList?.querySelectorAll('.contact-item').length > 0;

            if (assignActiveChatAgentBtn) {
                assignActiveChatAgentBtn.disabled = !hasAnyContacts;
                assignActiveChatAgentBtn.textContent = hasThread && isAssigned
                        ? `Assigned: ${assigned.name || 'Agent'}`
                        : 'Assign Agent';
                assignActiveChatAgentBtn.title = hasAnyContacts ? 'Open agent assignment popup' : 'No chats available yet';
            }

            if (chatAssignmentHintEl) {
                chatAssignmentHintEl.classList.add('hidden');
                chatAssignmentHintEl.textContent = '';
            }

            if (chatInput) {
                if (!chatInput.dataset.defaultPlaceholder) {
                    chatInput.dataset.defaultPlaceholder = chatInput.getAttribute('placeholder') || 'Type message here...';
                }
                chatInput.readOnly = hasThread ? !canMessage : true;
                chatInput.value = hasThread && !canMessage ? '' : chatInput.value;
                chatInput.setAttribute(
                    'placeholder',
                    !hasThread
                        ? 'Type message here...'
                        : canMessage
                            ? (chatInput.dataset.defaultPlaceholder || 'Type message here...')
                            : isAssigned
                                ? `Assigned to ${assigned.name || 'another agent'}`
                                : 'Assign this chat to start messaging'
                );
            }

            if (sendBtn) sendBtn.disabled = !canMessage;
            if (attachFileBtn) attachFileBtn.disabled = !canMessage;
            if (notifyBtn) notifyBtn.disabled = !canMessage;
        }

        async function assignContactToCurrentAgent(waId) {
            await assignContactToAgent(waId, agentName);
        }

        async function assignContactToAgent(waId, nextAgentName) {
            const normalizedWaId = normalizeWaId(waId);
            if (!normalizedWaId) return;
            const resolvedAgentName = normalizeAgentName(nextAgentName || agentName);
            const item = contactList.querySelector('.contact-item[data-wa-id="' + normalizedWaId + '"]');
            const profileName = item?.dataset?.profileName || item?.querySelector('.contact-name')?.textContent?.trim() || normalizedWaId;
            saveManualContact({
                waId: normalizedWaId,
                profileName,
                tag: item?.dataset?.tag || '',
                universityName: item?.dataset?.universityName || '',
                semester: item?.dataset?.semester || '',
                timezone: item?.dataset?.timezone || getAutoTimezone(),
                assignedAgent: resolvedAgentName,
                assignedAgentKey: getAgentKey(resolvedAgentName),
                assignedAgentEmail: agentEmail
            });
            if (item) {
                item.dataset.assignedAgent = resolvedAgentName;
                item.dataset.assignedAgentKey = getAgentKey(resolvedAgentName);
                item.dataset.assignedAgentEmail = agentEmail;
                syncContactAssignmentBadge(item);
                updateContactStateUI(item);
            }
            try {
                await setContactAiMode(normalizedWaId, 'human');
            } catch {}
            updateChatAssignmentState(normalizedWaId);
            scheduleCrmStateSave();
        }

        async function assignAllChatsToAgent(nextAgentName) {
            const resolvedAgentName = normalizeAgentName(nextAgentName || agentName);
            const currentRows = readManualContactsFromStorage();
            const rowsByWaId = new Map();
            currentRows.forEach((row) => {
                const normalizedWaId = normalizeWaId(row?.waId || '');
                if (!normalizedWaId) return;
                rowsByWaId.set(normalizedWaId, {
                    ...row,
                    waId: normalizedWaId,
                    assignedAgent: resolvedAgentName,
                    assignedAgentKey: getAgentKey(resolvedAgentName),
                    assignedAgentEmail: agentEmail
                });
            });
            contactList.querySelectorAll('.contact-item').forEach((item) => {
                const waId = normalizeWaId(item.dataset.waId || '');
                if (!waId) return;
                rowsByWaId.set(waId, {
                    waId,
                    profileName: item.dataset.profileName || item.querySelector('.contact-name')?.textContent?.trim() || waId,
                    tag: item.dataset.tag || '',
                    universityName: item.dataset.universityName || '',
                    semester: item.dataset.semester || '',
                    timezone: item.dataset.timezone || getAutoTimezone(),
                    assignedAgent: resolvedAgentName,
                    assignedAgentKey: getAgentKey(resolvedAgentName),
                    assignedAgentEmail: agentEmail
                });
                item.dataset.assignedAgent = resolvedAgentName;
                item.dataset.assignedAgentKey = getAgentKey(resolvedAgentName);
                item.dataset.assignedAgentEmail = agentEmail;
                syncContactAssignmentBadge(item);
                updateContactStateUI(item);
            });
            writeManualContactsToStorage([...rowsByWaId.values()]);
            if (activeWhatsappWaId) updateChatAssignmentState(activeWhatsappWaId);
        }

        function renderAssignAgentModal() {
            if (!assignAgentList) return;
            const roster = ensureAgentRoster();
            assignAgentAdminPanel?.classList.toggle('hidden', !isAdminUser);
            if (!roster.length) {
                assignAgentList.innerHTML = '<div class="assign-agent-empty">No agents added yet.</div>';
                return;
            }
            assignAgentList.innerHTML = roster.map((name) => {
                const isSelected = getAgentKey(name) === getAgentKey(selectedAssignAgentName);
                return `
                    <button type="button" class="assign-agent-option${isSelected ? ' is-selected' : ''}" data-agent-name="${escapeHtml(name)}">
                        <span class="assign-agent-option-copy">
                            <span class="assign-agent-option-avatar">${escapeHtml((name.charAt(0) || 'A').toUpperCase())}</span>
                            <span class="assign-agent-option-meta">
                                <span class="assign-agent-option-name">${escapeHtml(name)}</span>
                                <span class="assign-agent-option-subtitle">${getAgentKey(name) === getAgentKey(agentName) ? 'This device is using this name' : 'Click to select this agent'}</span>
                            </span>
                        </span>
                        ${isAdminUser && roster.length > 1 ? `<span class="assign-agent-option-delete" data-delete-agent-name="${escapeHtml(name)}">Delete</span>` : ''}
                    </button>
                `;
            }).join('');
        }

        function openAssignAgentModal(mode = 'all', waId = '') {
            if (!assignAgentModal) return;
            assignAgentModalMode = mode === 'single' ? 'single' : 'all';
            assignAgentModalWaId = normalizeWaId(waId);
            selectedAssignAgentName = normalizeAgentName(getAssignedAgentMetaByWaId(assignAgentModalWaId).name || agentName);
            if (assignAgentModalTitle) {
                assignAgentModalTitle.textContent = assignAgentModalMode === 'all' ? 'Assign All Chats' : 'Assign Chat';
            }
            if (assignAgentModalSubtitle) {
                assignAgentModalSubtitle.textContent = assignAgentModalMode === 'all'
                    ? 'Choose one agent and assign every chat in one click.'
                    : 'Choose one agent for this chat.';
            }
            if (confirmAssignAgentBtn) {
                confirmAssignAgentBtn.textContent = assignAgentModalMode === 'all' ? 'Assign All Chats' : 'Assign Chat';
            }
            renderAssignAgentModal();
            assignAgentModal.classList.remove('hidden');
            document.body.classList.add('modal-open');
        }

        function closeAssignAgentModal() {
            assignAgentModal?.classList.add('hidden');
            document.body.classList.remove('modal-open');
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
                return Array.isArray(cloudCrmState.manualContacts) ? cloudCrmState.manualContacts : [];
            } catch {
                return [];
            }
        }

        function readExpertsFromStorage() {
            try {
                const rows = Array.isArray(cloudCrmState.experts) ? cloudCrmState.experts : [];
                const normalized = dedupeExperts(rows);
                cloudCrmState.experts = normalized;
                cloudCrmState.expertIdSequence = getNextExpertSequenceFromRows(normalized);
                return normalized;
            } catch {
                return [];
            }
        }

        function readOrderDetailsDrafts() {
            try {
                const raw = localStorage.getItem(ORDER_DETAILS_DRAFTS_STORAGE_KEY);
                if (!raw) return {};
                const parsed = JSON.parse(raw);
                return parsed && typeof parsed === 'object' ? parsed : {};
            } catch {
                return {};
            }
        }

        function writeOrderDetailsDrafts(drafts) {
            localStorage.setItem(ORDER_DETAILS_DRAFTS_STORAGE_KEY, JSON.stringify(drafts || {}));
        }

        function getActiveOrderDraftKey(card = activeOrderCard) {
            if (!card) return '';
            const orderId = String(card.querySelector('.order-id')?.textContent || '').replace('#', '').trim();
            const clientId = String(card.dataset.clientId || '').trim();
            return orderId || clientId || '';
        }

        function collectOrderDetailsDraft() {
            const payload = {};
            ORDER_DETAILS_DRAFT_FIELD_IDS.forEach((id) => {
                const field = document.getElementById(id);
                if (!field) return;
                payload[id] = String(field.value || '');
            });
            return payload;
        }

        function applyOrderDetailsDraft(draft) {
            if (!draft || typeof draft !== 'object') return;
            ORDER_DETAILS_DRAFT_FIELD_IDS.forEach((id) => {
                if (!(id in draft)) return;
                const field = document.getElementById(id);
                if (!field) return;
                field.value = String(draft[id] || '');
            });
        }

        function saveActiveOrderDraft() {
            const draftKey = getActiveOrderDraftKey();
            if (!draftKey || !orderDetailsModal || orderDetailsModal.classList.contains('hidden')) return;
            const drafts = readOrderDetailsDrafts();
            drafts[draftKey] = collectOrderDetailsDraft();
            writeOrderDetailsDrafts(drafts);
        }

        function clearOrderDraft(card = activeOrderCard) {
            const draftKey = getActiveOrderDraftKey(card);
            if (!draftKey) return;
            const drafts = readOrderDetailsDrafts();
            if (!(draftKey in drafts)) return;
            delete drafts[draftKey];
            writeOrderDetailsDrafts(drafts);
        }

        function openOrderDetailsModal(card) {
            if (!card) return;
            activeOrderCard = card;

            const title = card.querySelector('.order-title')?.textContent?.trim() || '';
            const status = card.dataset.status || card.querySelector('.order-status')?.textContent?.trim() || 'In Progress';
            const cardId = (card.querySelector('.order-id')?.textContent || '').replace('#', '').trim();
            const clientIdText = card.dataset.clientId || cardId;
            const amount = card.querySelector('.order-amount')?.textContent?.trim() || '';
            const deadline = card.dataset.actualDeadline || card.querySelector('.order-deadline')?.dataset.baseValue || card.querySelector('.order-deadline')?.textContent?.trim() || '';

            document.getElementById('odTitle').value = title;
            document.getElementById('odServiceType').value = card.dataset.serviceType || title;
            document.getElementById('odStatus').value = status;
            document.getElementById('odOrderId').value = cardId;
            document.getElementById('odClientId').value = clientIdText;
            document.getElementById('odAssignedTo').value = card.dataset.assignedTo || '';
            document.getElementById('odCreatedBy').value = card.dataset.createdBy || card.dataset.assignedTo || '';
            document.getElementById('odActualDeadline').value = toDateTimeLocalValue(deadline);
            document.getElementById('odExpertDeadline').value = toDateTimeLocalValue(card.dataset.expertDeadline || card.querySelector('.order-expert-deadline')?.dataset.baseValue || '');
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
            if (odExpertPaymentStatusInput) odExpertPaymentStatusInput.value = card.dataset.expertPaymentStatus || 'pending';
            document.getElementById('odSessionStart').value = card.dataset.sessionStart || '';
            document.getElementById('odSessionDuration').value = card.dataset.sessionDuration || '';
            document.getElementById('odActivity').value = '';
            renderCommentHistory(parseComments(card.dataset.comments || '[]'));
            const draft = readOrderDetailsDrafts()[getActiveOrderDraftKey(card)];
            applyOrderDetailsDraft(draft);
            updateOrderDetailsSummary(card);
            document.body.classList.add('modal-open');
            orderDetailsModal.classList.remove('hidden');
            const scrollArea = document.getElementById('orderDetailsScrollArea');
            if (scrollArea) scrollArea.scrollTop = 0;
        }

        function closeOrderDetailsModal() {
            saveActiveOrderDraft();
            orderDetailsModal.classList.add('hidden');
            document.body.classList.remove('modal-open');
        }

        function writeExpertsToStorage(rows, skipRemote) {
            cloudCrmState.experts = Array.isArray(rows) ? rows : [];
            experts = Array.isArray(rows) ? rows : [];
            if (!skipRemote) {
                scheduleCrmStateSave();
            }
        }

        function getStoredExpertIdSequence() {
            return Math.max(1, Number(cloudCrmState.expertIdSequence || '1') || 1);
        }

        function setExpertIdSequenceInStorage(value, skipRemote) {
            cloudCrmState.expertIdSequence = Math.max(1, Number(value || 1) || 1);
            if (!skipRemote) {
                scheduleCrmStateSave();
            }
        }

        function getExpertIdNumber(expertId) {
            const match = String(expertId || '').trim().match(/^USXT(\d+)$/i);
            return match ? Number(match[1]) || 0 : 0;
        }

        function getNextExpertSequenceFromRows(rows) {
            const maxId = (Array.isArray(rows) ? rows : []).reduce((maxValue, row) => {
                return Math.max(maxValue, getExpertIdNumber(row?.expertId || ''));
            }, 0);
            return maxId > 0 ? maxId + 1 : 1;
        }

        function getPreferredExpertId(...candidates) {
            const values = candidates
                .map((value) => String(value || '').trim())
                .filter(Boolean);
            if (!values.length) return '';
            const withNumbers = values
                .map((value) => ({ value, num: getExpertIdNumber(value) }))
                .filter((entry) => entry.num > 0)
                .sort((a, b) => a.num - b.num);
            return withNumbers[0]?.value || values[0];
        }

        function dedupeExpertise(values) {
            const unique = [];
            (Array.isArray(values) ? values : []).forEach((value) => {
                const skill = String(value || '').trim();
                if (!skill) return;
                if (unique.some((item) => item.toLowerCase() === skill.toLowerCase())) return;
                unique.push(skill);
            });
            return unique;
        }

        function pickPreferredExpertValue(currentValue, nextValue) {
            const current = String(currentValue || '').trim();
            const next = String(nextValue || '').trim();
            if (!current) return next;
            if (!next) return current;
            return next.length >= current.length ? next : current;
        }

        function mergeExpertRecords(baseRow, incomingRow) {
            const base = normalizeExpertRecord(baseRow);
            const incoming = normalizeExpertRecord(incomingRow);
            return normalizeExpertRecord({
                expertId: getPreferredExpertId(base.expertId, incoming.expertId),
                name: pickPreferredExpertValue(base.name, incoming.name),
                expertise: dedupeExpertise([...(base.expertise || []), ...(incoming.expertise || [])]),
                rating: pickPreferredExpertValue(base.rating, incoming.rating),
                education: pickPreferredExpertValue(base.education, incoming.education),
                waId: base.waId || incoming.waId,
                createdAt: [base.createdAt, incoming.createdAt].filter(Boolean).sort()[0] || '',
                updatedAt: [base.updatedAt, incoming.updatedAt].filter(Boolean).sort().pop() || ''
            });
        }

        function dedupeExperts(rows) {
            const sourceRows = Array.isArray(rows) ? rows : [];
            const byWaId = new Map();
            const byFallback = new Map();

            sourceRows.forEach((row) => {
                const normalized = normalizeExpertRecord(row);
                if (!normalized.name) return;
                const mapKey = normalized.waId
                    ? `wa:${normalized.waId}`
                    : normalized.expertId
                        ? `id:${normalized.expertId}`
                        : `name:${normalized.name.toLowerCase()}|edu:${normalized.education.toLowerCase()}`;

                const targetMap = normalized.waId ? byWaId : byFallback;
                const existing = targetMap.get(mapKey);
                targetMap.set(mapKey, existing ? mergeExpertRecords(existing, normalized) : normalized);
            });

            return [...byWaId.values(), ...byFallback.values()]
                .map((row) => normalizeExpertRecord(row))
                .filter((row) => row.name)
                .sort((a, b) => {
                    const diff = getExpertIdNumber(a.expertId) - getExpertIdNumber(b.expertId);
                    if (diff !== 0) return diff;
                    return a.name.localeCompare(b.name);
                });
        }

        function writeManualContactsToStorage(rows, skipRemote) {
            cloudCrmState.manualContacts = Array.isArray(rows) ? rows : [];
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
                timezone: typeof contact.timezone === 'string' ? contact.timezone : String(previous?.timezone || getAutoTimezone()),
                assignedAgent: typeof contact.assignedAgent === 'string' ? contact.assignedAgent : String(previous?.assignedAgent || ''),
                assignedAgentKey: typeof contact.assignedAgentKey === 'string' ? String(contact.assignedAgentKey || '').trim().toLowerCase() : String(previous?.assignedAgentKey || '').trim().toLowerCase(),
                assignedAgentEmail: typeof contact.assignedAgentEmail === 'string' ? String(contact.assignedAgentEmail || '').trim().toLowerCase() : String(previous?.assignedAgentEmail || '').trim().toLowerCase()
            };
            if (existingIndex >= 0) {
                rows[existingIndex] = payload;
            } else {
                rows.push(payload);
            }
            writeManualContactsToStorage(rows);
        }

        function removeManualContact(waId) {
            const normalizedWaId = normalizeWaId(waId);
            if (!normalizedWaId) return;
            const rows = readManualContactsFromStorage().filter((row) => normalizeWaId(row?.waId || '') !== normalizedWaId);
            writeManualContactsToStorage(rows);
        }

        function removeWhatsappContactMeta(waId) {
            const normalizedWaId = normalizeWaId(waId);
            if (!normalizedWaId) return;

            try {
                const idMap = readWhatsappContactIdMapFromStorage();
                delete idMap[normalizedWaId];
                writeWhatsappContactIdMapToStorage(idMap, true);
            } catch {}

            try {
                const readState = readWhatsappReadStateFromStorage();
                delete readState[normalizedWaId];
                writeWhatsappReadStateToStorage(readState, true);
            } catch {}

            scheduleCrmStateSave();
        }

        function normalizeExpertRecord(raw = {}) {
            return {
                expertId: String(raw.expertId || '').trim(),
                name: String(raw.name || '').trim(),
                expertise: Array.isArray(raw.expertise)
                    ? raw.expertise.map((item) => String(item || '').trim()).filter(Boolean)
                    : String(raw.expertise || '')
                        .split(/[|,]/)
                        .map((item) => String(item || '').trim())
                        .filter(Boolean),
                rating: String(raw.rating || '').trim(),
                education: String(raw.education || '').trim(),
                waId: normalizeWaId(raw.waId || ''),
                createdAt: String(raw.createdAt || ''),
                updatedAt: String(raw.updatedAt || '')
            };
        }

        function getNextExpertId() {
            const next = getNextExpertSequenceFromRows(readExpertsFromStorage());
            setExpertIdSequenceInStorage(next + 1);
            return `USXT${String(next).padStart(4, '0')}`;
        }

        function saveExperts(rows) {
            let nextSequence = 1;
            const normalized = dedupeExperts(rows)
                .map((row) => normalizeExpertRecord(row))
                .map((row) => {
                    const nextRow = normalizeExpertRecord(row);
                    if (!nextRow.expertId) {
                        nextRow.expertId = `USXT${String(nextSequence).padStart(4, '0')}`;
                    }
                    nextSequence = Math.max(nextSequence, getExpertIdNumber(nextRow.expertId) + 1);
                    return nextRow;
                })
                .filter((row) => row.expertId && row.name);
            setExpertIdSequenceInStorage(getNextExpertSequenceFromRows(normalized), true);
            writeExpertsToStorage(normalized);
            return normalized;
        }

        function upsertExpert(rawExpert) {
            const expert = normalizeExpertRecord(rawExpert);
            if (!expert.name) return null;
            const nowIso = new Date().toISOString();

            const rows = readExpertsFromStorage();
            const index = rows.findIndex((row) => {
                const current = normalizeExpertRecord(row);
                if (expert.expertId && current.expertId === expert.expertId) return true;
                if (expert.waId && current.waId && current.waId === expert.waId) return true;
                if (!expert.waId && current.name.toLowerCase() === expert.name.toLowerCase() && current.education.toLowerCase() === expert.education.toLowerCase()) return true;
                return false;
            });

            if (index >= 0) {
                const existing = normalizeExpertRecord(rows[index]);
                rows[index] = {
                    ...existing,
                    ...expert,
                    expertId: existing.expertId || expert.expertId || getNextExpertId(),
                    createdAt: existing.createdAt || expert.createdAt || nowIso,
                    updatedAt: nowIso
                };
            } else {
                if (!expert.expertId) {
                    expert.expertId = getNextExpertId();
                }
                if (!expert.createdAt) expert.createdAt = nowIso;
                expert.updatedAt = nowIso;
                rows.push(expert);
            }
            const savedRows = saveExperts(rows);
            return savedRows.find((row) => (expert.waId && row.waId === expert.waId) || row.expertId === expert.expertId) || expert;
        }

        function getExpertInContacts(expert) {
            const waId = normalizeWaId(expert?.waId || '');
            if (!waId) return null;
            return contactList.querySelector('.contact-item[data-wa-id="' + waId + '"]');
        }

        function getExpertByWaId(waId) {
            const normalizedWaId = normalizeWaId(waId);
            if (!normalizedWaId) return null;
            return readExpertsFromStorage().find((row) => normalizeWaId(row?.waId || '') === normalizedWaId) || null;
        }

        function getExpertRatingDisplay(rawRating) {
            const numeric = Math.max(0, Math.min(5, Number(rawRating || 0) || 0));
            const filledStars = Math.round(numeric);
            const stars = '★'.repeat(filledStars) + '☆'.repeat(Math.max(0, 5 - filledStars));
            return {
                value: numeric ? numeric.toFixed(numeric % 1 === 0 ? 0 : 1) : '0',
                stars
            };
        }

        function addExpertToContacts(expertId) {
            const rows = readExpertsFromStorage();
            const expert = rows.find((row) => String(row?.expertId || '') === String(expertId || ''));
            if (!expert) {
                alert('Expert not found.');
                return;
            }
            const waId = normalizeWaId(expert.waId || '');
            if (!waId) {
                alert('This expert does not have a valid WhatsApp number for chat.');
                return;
            }
            if (!Array.isArray(whatsappMessagesByContact[waId])) {
                whatsappMessagesByContact[waId] = [];
            }
            upsertWhatsappContact({
                waId,
                profileName: expert.name,
                tag: 'tutor'
            }, false);
            saveManualContact({
                waId,
                profileName: expert.name,
                tag: 'tutor'
            });
            refreshContactOrder();
            applyContactFilter();
            renderExpertList();
        }

        function promptAddSkillToExpert(expertId) {
            const rows = readExpertsFromStorage();
            const index = rows.findIndex((row) => String(row?.expertId || '') === String(expertId || ''));
            if (index < 0) {
                alert('Expert not found.');
                return;
            }
            const skill = (window.prompt('Add skill / expertise:', '') || '').trim();
            if (!skill) return;
            const existingSkills = normalizeExpertRecord(rows[index]).expertise;
            if (!existingSkills.some((item) => item.toLowerCase() === skill.toLowerCase())) {
                existingSkills.push(skill);
            }
            rows[index] = {
                ...normalizeExpertRecord(rows[index]),
                expertise: existingSkills,
                updatedAt: new Date().toISOString()
            };
            saveExperts(rows);
            renderExpertList();
        }

        function deleteExpert(expertId) {
            const normalizedExpertId = String(expertId || '').trim();
            if (!normalizedExpertId) return;
            const rows = readExpertsFromStorage();
            const existing = rows.find((row) => String(row?.expertId || '') === normalizedExpertId);
            if (!existing) {
                alert('Expert not found.');
                return;
            }
            const confirmed = window.confirm(`Delete expert "${existing.name}" (${normalizedExpertId})?`);
            if (!confirmed) return;

            const nextRows = rows.filter((row) => String(row?.expertId || '') !== normalizedExpertId);
            saveExperts(nextRows);
            if (!nextRows.length) {
                setExpertIdSequenceInStorage(1);
            }

            const waId = normalizeWaId(existing.waId || '');
            if (waId) {
                const contactItem = contactList.querySelector('.contact-item[data-wa-id="' + waId + '"]');
                if (contactItem) {
                    contactItem.dataset.expertId = '';
                    contactItem.querySelector('.contact-expert-id-badge')?.remove();
                }
            }

            renderExpertList();
        }

        function parseCsvLine(line) {
            const values = [];
            let current = '';
            let insideQuotes = false;
            for (let i = 0; i < line.length; i += 1) {
                const char = line[i];
                const next = line[i + 1];
                if (char === '"' && insideQuotes && next === '"') {
                    current += '"';
                    i += 1;
                    continue;
                }
                if (char === '"') {
                    insideQuotes = !insideQuotes;
                    continue;
                }
                if (char === ',' && !insideQuotes) {
                    values.push(current);
                    current = '';
                    continue;
                }
                current += char;
            }
            values.push(current);
            return values.map((value) => String(value || '').trim());
        }

        function parseExpertsCsv(text) {
            const lines = String(text || '')
                .split(/\r?\n/)
                .map((line) => line.trim())
                .filter(Boolean);
            if (lines.length < 2) return [];
            const headers = parseCsvLine(lines[0]).map((header) => header.toLowerCase());
            const findIndex = (aliases) => headers.findIndex((header) => aliases.includes(header));
            const nameIndex = findIndex(['name', 'expert name']);
            const waIndex = findIndex(['whatsapp', 'whatsapp number', 'phone', 'mobile', 'number']);
            const skillIndex = findIndex(['skill', 'skills', 'expertise', 'skill/expertise', 'subject']);
            const ratingIndex = findIndex(['rating', 'ratings']);
            const educationIndex = findIndex(['education', 'qualification']);
            if (nameIndex < 0) {
                throw new Error('CSV must include a Name column.');
            }
            return lines.slice(1).map((line) => {
                const cols = parseCsvLine(line);
                return {
                    name: cols[nameIndex] || '',
                    waId: waIndex >= 0 ? cols[waIndex] || '' : '',
                    expertise: skillIndex >= 0 ? cols[skillIndex] || '' : '',
                    rating: ratingIndex >= 0 ? cols[ratingIndex] || '' : '',
                    education: educationIndex >= 0 ? cols[educationIndex] || '' : ''
                };
            }).filter((row) => String(row.name || '').trim());
        }

        function handleExpertsImportText(text) {
            const parsed = parseExpertsCsv(text);
            if (!parsed.length) {
                alert('No expert rows found in CSV.');
                return;
            }
            parsed.forEach((row) => {
                upsertExpert({
                    name: row.name,
                    waId: row.waId,
                    expertise: row.expertise,
                    rating: row.rating,
                    education: row.education
                });
            });
            renderExpertList();
            alert(`${parsed.length} experts imported successfully.`);
        }

        function downloadExpertsExportCsv() {
            const rows = readExpertsFromStorage()
                .map((expert) => normalizeExpertRecord(expert))
                .filter((expert) => expert.expertId && expert.name)
                .sort((a, b) => String(a.expertId || '').localeCompare(String(b.expertId || '')));

            if (!rows.length) {
                alert('No experts available to export.');
                return;
            }

            const header = ['Expert Name', 'Expert ID', 'Skill/Expertise', 'Ratings', 'Education', 'WhatsApp Number', 'Added To Contacts'];
            const lines = [header.map(escapeCsvValue).join(',')];
            rows.forEach((expert) => {
                lines.push([
                    expert.name,
                    expert.expertId,
                    expert.expertise.join(' | '),
                    expert.rating,
                    expert.education,
                    expert.waId,
                    getExpertInContacts(expert) ? 'Yes' : 'No'
                ].map(escapeCsvValue).join(','));
            });

            const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const anchor = document.createElement('a');
            anchor.href = url;
            anchor.download = `unisolvex-experts-${toDateInputValue(new Date())}.csv`;
            document.body.appendChild(anchor);
            anchor.click();
            anchor.remove();
            URL.revokeObjectURL(url);
        }

        function renderExpertList() {
            if (!orderList || activeOrderTab !== 'expert') return;
            const query = (orderSearch.value || '').trim().toLowerCase();
            orderList.innerHTML = '';

            const filteredExperts = readExpertsFromStorage().filter((expert) => {
                const skills = normalizeExpertRecord(expert).expertise.join(' ').toLowerCase();
                return !query
                    || String(expert.name || '').toLowerCase().includes(query)
                    || String(expert.expertId || '').toLowerCase().includes(query)
                    || skills.includes(query);
            });

            if (!filteredExperts.length) {
                orderList.innerHTML = '<div class="order-card p-4 rounded-lg border text-sm text-gray-500">No experts found. Import experts from the admin menu.</div>';
                return;
            }

            filteredExperts
                .sort((a, b) => String(a.name || '').localeCompare(String(b.name || '')))
                .forEach((expert) => {
                    const normalized = normalizeExpertRecord(expert);
                    const inContacts = Boolean(getExpertInContacts(normalized));
                    const skills = normalized.expertise.length ? normalized.expertise.join(', ') : 'Not added';
                    const rating = getExpertRatingDisplay(normalized.rating);
                    const education = normalized.education || 'Not added';
                    const canAddToContacts = Boolean(normalized.waId) && !inContacts;
                    const initials = getContactInitials(normalized.name, normalized.expertId);

                    const card = document.createElement('div');
                    card.className = 'order-card expert-card p-3 rounded-lg border shadow-sm';
                    card.dataset.expertId = normalized.expertId;
                    card.innerHTML = `
                        <div class="flex items-start justify-between gap-3">
                            <div class="expert-card-head min-w-0">
                                <div class="expert-card-avatar" aria-hidden="true">${escapeHtml(initials)}</div>
                                <div class="min-w-0">
                                    <p class="order-detail-label text-[10px] font-bold uppercase px-2 py-0.5 rounded">Expert</p>
                                    <h4 class="order-title mt-2 text-sm font-bold text-gray-800">${escapeHtml(normalized.name)}</h4>
                                    <p class="text-[11px] text-gray-500 mt-1">Expert ID: <span class="font-semibold text-gray-700">${escapeHtml(normalized.expertId)}</span></p>
                                </div>
                            </div>
                            <div class="expert-card-actions">
                                <button type="button" class="expert-add-skill-btn" data-expert-id="${escapeHtml(normalized.expertId)}">Add Skill</button>
                                <button type="button" class="expert-add-contact-btn ${canAddToContacts ? 'is-ready' : 'is-muted'}" data-expert-id="${escapeHtml(normalized.expertId)}" ${canAddToContacts ? '' : 'disabled'}>${inContacts ? 'Added to Contacts' : (normalized.waId ? 'Add to Contacts' : 'No WhatsApp')}</button>
                                ${isAdminUser ? `<button type="button" class="expert-delete-btn" data-expert-id="${escapeHtml(normalized.expertId)}">Delete</button>` : ''}
                            </div>
                        </div>
                        <div class="mt-3 grid grid-cols-1 gap-2 text-[12px] text-gray-600">
                            <p><span class="font-semibold text-gray-700">Skill/Expertise:</span> ${escapeHtml(skills)}</p>
                            <div class="expert-card-meta-row">
                                <p class="expert-rating-pill"><span class="font-semibold">Rating:</span> <span class="expert-rating-stars" aria-label="Rated ${escapeHtml(rating.value)} out of 5">${escapeHtml(rating.stars)}</span> <span class="expert-rating-value">${escapeHtml(rating.value)}/5</span></p>
                                <p class="expert-education-pill"><span class="font-semibold">Education:</span> ${escapeHtml(education)}</p>
                            </div>
                        </div>
                    `;
                    orderList.appendChild(card);
                });
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
                return cloudCrmState.whatsappContactIdMap && typeof cloudCrmState.whatsappContactIdMap === 'object'
                    ? cloudCrmState.whatsappContactIdMap
                    : {};
            } catch {
                return {};
            }
        }

        function readWhatsappReadStateFromStorage() {
            try {
                return cloudCrmState.whatsappReadState && typeof cloudCrmState.whatsappReadState === 'object'
                    ? cloudCrmState.whatsappReadState
                    : {};
            } catch {
                return {};
            }
        }

        function writeWhatsappReadStateToStorage(map, skipRemote) {
            cloudCrmState.whatsappReadState = map && typeof map === 'object' ? map : {};
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
            cloudCrmState.whatsappContactIdMap = map && typeof map === 'object' ? map : {};
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
            if (normalized === 'in progress') {
                statusEl.classList.add('text-indigo-700', 'bg-indigo-100');
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
            statusEl.classList.add('text-slate-700', 'bg-slate-100');
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
            if (value.includes('tutoring class')) {
                card.dataset.serviceTone = 'tutoring-class';
                card.classList.add('bg-emerald-50', 'border-emerald-200');
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

        function syncOrderExpertPaymentIndicator(card) {
            const payoutEl = card.querySelector('.order-expert-pay');
            if (!payoutEl) return;

            const line = payoutEl.closest('p');
            if (!line) return;

            let indicatorEl = line.querySelector('.order-expert-payment-indicator');
            if (!indicatorEl) {
                indicatorEl = document.createElement('span');
                indicatorEl.className = 'order-expert-payment-indicator hidden ml-1';
                line.appendChild(indicatorEl);
            }

            const payoutAmount = parseAmountToNumber(payoutEl.textContent || '');
            const manualStatus = String(card.dataset.expertPaymentStatus || 'pending').toLowerCase();

            if (!payoutAmount) {
                indicatorEl.className = 'order-expert-payment-indicator hidden ml-1';
                indicatorEl.textContent = '';
                return;
            }

            if (manualStatus === 'complete') {
                indicatorEl.className = 'order-expert-payment-indicator ml-1 inline-flex items-center justify-center w-4 h-4 rounded-full bg-green-600 text-white text-[10px] font-bold';
                indicatorEl.innerHTML = '&#10003;';
                return;
            }

            if (manualStatus === 'partial') {
                indicatorEl.className = 'order-expert-payment-indicator ml-1 inline-flex items-center gap-1 text-[10px] font-semibold text-gray-700';
                indicatorEl.innerHTML = '<span class="inline-flex items-center justify-center w-4 h-4 rounded-full bg-black text-white text-[9px] leading-none">...</span><span>Partially Paid</span>';
                return;
            }

            indicatorEl.className = 'order-expert-payment-indicator hidden ml-1';
            indicatorEl.textContent = '';
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
            const titleRow = card.querySelector('.order-title-row');

            if (!dueSoonChip) {
                dueSoonChip = document.createElement('span');
                dueSoonChip.className = 'order-due-soon-chip hidden text-[9px] font-semibold uppercase tracking-[0.08em] px-2 py-0.5 rounded-full';
            }
            if (titleRow && dueSoonChip.parentElement !== titleRow) {
                titleRow.appendChild(dueSoonChip);
            } else if (!titleRow && orderIdEl && dueSoonChip.parentElement !== orderIdEl.parentElement) {
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

        function parseJsonArray(raw) {
            if (!raw) return [];
            try {
                const parsed = JSON.parse(raw);
                return Array.isArray(parsed) ? parsed : [];
            } catch (err) {
                return [];
            }
        }

        function getNotifiedExpertIds(card) {
            return parseJsonArray(card?.dataset?.notifiedExperts || '[]')
                .map((value) => String(value || '').trim())
                .filter(Boolean);
        }

        function setNotifiedExpertIds(card, values) {
            if (!card) return;
            const unique = Array.from(new Set((Array.isArray(values) ? values : []).map((value) => String(value || '').trim()).filter(Boolean)));
            card.dataset.notifiedExperts = JSON.stringify(unique);
        }

        function getInterestedExpertIds(card) {
            return parseJsonArray(card?.dataset?.interestedExperts || '[]')
                .map((value) => String(value || '').trim())
                .filter(Boolean);
        }

        function setInterestedExpertIds(card, values) {
            if (!card) return;
            const unique = Array.from(new Set((Array.isArray(values) ? values : []).map((value) => String(value || '').trim()).filter(Boolean)));
            card.dataset.interestedExperts = JSON.stringify(unique);
        }

        function getAssignedExpertIds(card) {
            return parseJsonArray(card?.dataset?.assignedExperts || '[]')
                .map((value) => String(value || '').trim())
                .filter(Boolean);
        }

        function setAssignedExpertIds(card, values) {
            if (!card) return;
            const unique = Array.from(new Set((Array.isArray(values) ? values : []).map((value) => String(value || '').trim()).filter(Boolean)));
            card.dataset.assignedExperts = JSON.stringify(unique);
            const firstAssigned = unique[0] || '';
            const expert = getExpertById(firstAssigned);
            card.dataset.assignedExpertId = firstAssigned;
            card.dataset.assignedExpertName = expert?.name || '';
        }

        function getExpertById(expertId) {
            const normalizedId = String(expertId || '').trim();
            if (!normalizedId) return null;
            return readExpertsFromStorage()
                .map((row) => normalizeExpertRecord(row))
                .find((row) => String(row.expertId || '').trim() === normalizedId) || null;
        }

        function buildManageExpertRows(card, mode) {
            if (!card) return [];
            const ids = mode === 'interested'
                ? getInterestedExpertIds(card)
                : mode === 'assigned'
                    ? getAssignedExpertIds(card)
                    : getNotifiedExpertIds(card);
            return ids
                .map((id) => getExpertById(id) || { expertId: id, name: id, expertise: [], rating: '', education: '', waId: '' })
                .filter(Boolean);
        }

        function normalizeInterestedExperts(raw) {
            const values = Array.isArray(raw)
                ? raw
                : String(raw || '')
                    .split(/[\n,]/)
                    .map((item) => item.trim())
                    .filter(Boolean);
            const unique = [];
            values.forEach((value) => {
                if (!value) return;
                if (unique.some((item) => item.toLowerCase() === value.toLowerCase())) return;
                unique.push(value);
            });
            return unique;
        }

        function updateOrderDetailsSummary(card) {
            if (!card) return;
            const setText = (id, value) => {
                const el = document.getElementById(id);
                if (el) el.textContent = String(value || '-').trim() || '-';
            };
            const orderId = (card.querySelector('.order-id')?.textContent || '').replace('#', '').trim() || '-';
            const clientId = String(card.dataset.clientId || '').trim() || '-';
            const serviceType = String(card.dataset.serviceType || '').trim() || card.querySelector('.order-title')?.textContent?.trim() || '-';
            const status = String(card.dataset.status || card.querySelector('.order-status')?.textContent || '').trim() || '-';
            const actualDeadline = card.querySelector('.order-deadline')?.dataset.baseValue || card.querySelector('.order-deadline')?.textContent?.trim() || '-';
            const expertDeadline = card.querySelector('.order-expert-deadline')?.dataset.baseValue || card.querySelector('.order-expert-deadline')?.textContent?.trim() || '-';
            const expertPayment = `${String(card.dataset.expertPaymentStatus || 'pending').trim() || 'pending'} | ${String(card.dataset.expertPayout || 'INR 0').trim() || 'INR 0'}`;
            const orderPayment = `${String(card.dataset.paymentStatusOverride || 'auto').trim() || 'auto'} | ${String(card.dataset.clientPaidCurrency || 'INR').trim() || 'INR'} ${String(card.dataset.clientPaidAmount || '0').trim() || '0'}`;
            const comments = parseComments(card.dataset.comments || '[]');
            const latestComment = comments.length ? comments[comments.length - 1] : null;
            setText('odSummaryOrderId', `#${orderId}`);
            setText('odSummaryClientId', clientId);
            setText('odSummaryService', serviceType);
            setText('odSummaryActualDeadline', actualDeadline);
            setText('odSummaryExpertDeadline', expertDeadline);
            setText('odSummaryExpertPayment', expertPayment);
            setText('odQuickStatus', status);
            setText('odQuickExpertPayment', expertPayment);
            setText('odQuickActualDeadline', actualDeadline);
            setText('odQuickExpertDeadline', expertDeadline);
            setText('odQuickOrderPayment', orderPayment);
            setText('odQuickComment', latestComment ? `${latestComment.agent || 'Agent'} - ${getOrderCommentPreview(latestComment.text || '', 56)}` : '-');
        }

        function getOrderRecordIdentity(card) {
            if (!card) return { orderId: '', clientId: '' };
            return {
                orderId: String(card.querySelector('.order-id')?.textContent || '').replace('#', '').trim(),
                clientId: String(card.dataset.clientId || '').trim()
            };
        }

        function parseOrderTransactions(raw) {
            return parseJsonArray(raw).map((row) => ({
                agent: String(row?.agent || '').trim(),
                currency: String(row?.currency || 'INR').trim().toUpperCase() || 'INR',
                amount: String(row?.amount || '').trim(),
                transactionId: String(row?.transactionId || '').trim(),
                paymentMethod: String(row?.paymentMethod || '').trim(),
                note: String(row?.note || '').trim(),
                createdAt: String(row?.createdAt || '').trim()
            })).filter((row) => row.amount || row.transactionId || row.paymentMethod);
        }

        function parseOrderAttachments(raw) {
            return parseJsonArray(raw).map((row) => ({
                path: String(row?.path || '').trim(),
                name: String(row?.name || '').trim(),
                url: String(row?.url || '').trim(),
                addedAt: String(row?.addedAt || '').trim()
            })).filter((row) => row.name || row.url);
        }

        function renderOrderTransactionHistory(items) {
            if (!odTransactionHistory) return;
            if (!items.length) {
                odTransactionHistory.innerHTML = '<p class="task-empty-state">No transactions recorded yet.</p>';
                return;
            }
            odTransactionHistory.innerHTML = items.map((item, index) => {
                const dateText = formatDateTimeForCard(item.createdAt || '');
                const amountText = [item.currency || 'INR', item.amount || '0'].filter(Boolean).join(' ');
                return `
                    <div class="task-meta-item">
                        <div>
                            <p class="task-meta-item-title">${escapeHtml(amountText)}</p>
                            <p class="task-meta-item-subtitle">${escapeHtml(item.transactionId || 'No transaction ID')}</p>
                        </div>
                        <div class="task-meta-item-actions">
                            <span class="task-meta-item-time">${escapeHtml(dateText || 'Just now')}</span>
                            <button type="button" class="task-meta-remove-btn" data-remove-transaction-index="${index}">Remove</button>
                        </div>
                    </div>
                `;
            }).join('');
        }

        function renderTransactionRecordsModal(items) {
            if (!transactionRecordsList) return;
            if (!items.length) {
                transactionRecordsList.innerHTML = '<div class="transaction-record-empty">No transaction records added yet.</div>';
                return;
            }
            transactionRecordsList.innerHTML = items.map((item, index) => {
                const dateText = formatDateTimeForCard(item.createdAt || '');
                return `
                    <div class="transaction-record-card">
                        <div class="flex items-start justify-between gap-3">
                            <div class="transaction-record-grid flex-1">
                                <div>
                                    <span class="transaction-record-label">Transaction ID</span>
                                    <span class="transaction-record-value">${escapeHtml(item.transactionId || '-')}</span>
                                </div>
                                <div>
                                    <span class="transaction-record-label">Amount Paid</span>
                                    <span class="transaction-record-value">${escapeHtml((item.currency || 'INR') + ' ' + (item.amount || '0'))}</span>
                                </div>
                                <div>
                                    <span class="transaction-record-label">Payment Method</span>
                                    <span class="transaction-record-value">${escapeHtml(item.paymentMethod || '-')}</span>
                                </div>
                                <div>
                                    <span class="transaction-record-label">Date & Time</span>
                                    <span class="transaction-record-value">${escapeHtml(dateText || '-')}</span>
                                </div>
                            </div>
                            ${isAdminUser ? `<button type="button" class="task-meta-remove-btn" data-remove-transaction-record-index="${index}">Remove</button>` : ''}
                        </div>
                    </div>
                `;
            }).join('');
        }

        function renderManageExpertsModal() {
            if (!manageExpertsList || !activeOrderCard) return;
            [manageExpertsTabNotified, manageExpertsTabInterested, manageExpertsTabAssigned].forEach((btn) => {
                if (!btn) return;
                btn.classList.toggle('is-active', btn.dataset.manageExpertsTab === activeManageExpertsTab);
            });
            const rows = buildManageExpertRows(activeOrderCard, activeManageExpertsTab);
            if (!rows.length) {
                manageExpertsList.innerHTML = `<div class="manage-experts-empty">No experts found.</div>`;
                return;
            }
            manageExpertsList.innerHTML = rows.map((expert) => {
                const skills = Array.isArray(expert.expertise) && expert.expertise.length ? expert.expertise.join(', ') : 'No skills added';
                const assignedIds = getAssignedExpertIds(activeOrderCard);
                const interestedIds = getInterestedExpertIds(activeOrderCard);
                const isAssigned = assignedIds.includes(expert.expertId);
                const isInterested = interestedIds.includes(expert.expertId);
                return `
                    <div class="manage-expert-card">
                        <div class="manage-expert-card-grid">
                            <div>
                                <span class="manage-expert-card-label">Expert</span>
                                <span class="manage-expert-card-value">${escapeHtml(expert.name || expert.expertId)}</span>
                            </div>
                            <div>
                                <span class="manage-expert-card-label">Expert ID</span>
                                <span class="manage-expert-card-value">${escapeHtml(expert.expertId || '-')}</span>
                            </div>
                            <div>
                                <span class="manage-expert-card-label">Skills</span>
                                <span class="manage-expert-card-value">${escapeHtml(skills)}</span>
                            </div>
                            <div>
                                <span class="manage-expert-card-label">WhatsApp</span>
                                <span class="manage-expert-card-value">${escapeHtml(expert.waId || '-')}</span>
                            </div>
                        </div>
                        <div class="manage-expert-card-actions">
                            ${activeManageExpertsTab === 'notified' ? `<button type="button" class="manage-expert-action-btn is-primary" data-manage-expert-action="mark-interested" data-expert-id="${escapeHtml(expert.expertId)}">Mark Interested</button>` : ''}
                            ${activeManageExpertsTab === 'interested' ? `<button type="button" class="manage-expert-action-btn is-success" data-manage-expert-action="assign" data-expert-id="${escapeHtml(expert.expertId)}">Assign</button>` : ''}
                            ${activeManageExpertsTab === 'interested' ? `<button type="button" class="manage-expert-action-btn is-muted" data-manage-expert-action="remove-interest" data-expert-id="${escapeHtml(expert.expertId)}">Remove</button>` : ''}
                            ${activeManageExpertsTab === 'assigned' ? `<button type="button" class="manage-expert-action-btn is-muted" data-manage-expert-action="unassign" data-expert-id="${escapeHtml(expert.expertId)}">Unassign</button>` : ''}
                            ${isAssigned ? `<span class="manage-expert-action-btn is-success">Assigned</span>` : ''}
                            ${isInterested && activeManageExpertsTab === 'notified' ? `<span class="manage-expert-action-btn is-muted">Interested</span>` : ''}
                        </div>
                    </div>
                `;
            }).join('');
        }

        function openManageExpertsModal() {
            if (!activeOrderCard || !manageExpertsModal) return;
            const identity = getOrderRecordIdentity(activeOrderCard);
            if (manageExpertsMeta) {
                manageExpertsMeta.textContent = `Order #${identity.orderId || '-'} | Client ${identity.clientId || '-'} | Manage expert flow`;
            }
            activeManageExpertsTab = 'notified';
            renderManageExpertsModal();
            manageExpertsModal.classList.remove('hidden');
        }

        function closeManageExpertsModal() {
            manageExpertsModal?.classList.add('hidden');
        }

        function renderOrderAttachmentList(target, items, kind) {
            if (!target) return;
            if (!items.length) {
                target.innerHTML = '<p class="task-empty-state">No items added yet.</p>';
                return;
            }
            target.innerHTML = items.map((item, index) => {
                const dateText = formatDateTimeForCard(item.addedAt || '');
                const href = item.url ? `href="${escapeHtml(item.url)}"` : '';
                return `
                    <div class="task-meta-item">
                        <div>
                            <p class="task-meta-item-title">${escapeHtml(item.name || 'Untitled')}</p>
                            <p class="task-meta-item-subtitle">${item.url ? `<a ${href} target="_blank" rel="noopener noreferrer" class="task-meta-link">${escapeHtml(item.url)}</a>` : 'No link added'}</p>
                        </div>
                        <div class="task-meta-item-actions">
                            <span class="task-meta-item-time">${escapeHtml(dateText || 'Just now')}</span>
                            <button type="button" class="task-meta-remove-btn" data-remove-attachment-kind="${escapeHtml(kind)}" data-remove-attachment-index="${index}">Remove</button>
                        </div>
                    </div>
                `;
            }).join('');
        }

        function ensureOrderCardEnhancements(card) {
            if (!card || card.classList.contains('expert-card')) return;
            let sessionEl = card.querySelector('.order-session');
            let commentEl = card.querySelector('.order-comment');
            if (!sessionEl) {
                sessionEl = document.createElement('p');
                sessionEl.className = 'order-session hidden mt-1 text-[11px] text-gray-600';
                commentEl = card.querySelector('.order-comment');
                if (commentEl) {
                    card.insertBefore(sessionEl, commentEl);
                } else {
                    card.appendChild(sessionEl);
                }
            }
            card.querySelector('.order-task-meta-grid')?.remove();
            card.querySelector('.order-attachment-summary')?.remove();

            const detailsGrid = card.querySelector('.mt-1.grid');
            const orderHeader = card.querySelector('.mb-2 .flex.items-center.gap-2');
            if (orderHeader) {
                let createdByTop = orderHeader.querySelector('.order-created-by-top');
                if (!createdByTop) {
                    createdByTop = document.createElement('span');
                    createdByTop.className = 'order-created-by-top text-[10px] text-gray-500 font-medium';
                    const orderIdEl = orderHeader.querySelector('.order-id');
                    if (orderIdEl) {
                        orderHeader.insertBefore(createdByTop, orderIdEl);
                    } else {
                        orderHeader.prepend(createdByTop);
                    }
                }
                const currentBy = card.querySelector('.order-created-by')?.textContent?.trim() || card.dataset.createdBy || card.dataset.assignedTo || agentName;
                createdByTop.textContent = `By - ${currentBy}`;

                let notifyBtn = orderHeader.querySelector('.open-expert-notify-btn');
                if (!notifyBtn) {
                    notifyBtn = document.createElement('button');
                    notifyBtn.type = 'button';
                    notifyBtn.className = 'open-expert-notify-btn p-1 rounded';
                    notifyBtn.title = 'Notify Experts';
                    const orderIdEl = orderHeader.querySelector('.order-id');
                    if (orderIdEl) {
                        orderHeader.insertBefore(notifyBtn, orderIdEl.nextSibling);
                    } else {
                        orderHeader.appendChild(notifyBtn);
                    }
                }
                syncOrderNotifyButton(card);
            }

            if (detailsGrid) {
                const caText = card.querySelector('.order-amount')?.textContent?.trim() || 'TBD';
                const cdRaw = card.querySelector('.order-deadline')?.dataset.baseValue || card.querySelector('.order-deadline')?.textContent?.trim() || '';
                const edRaw = card.querySelector('.order-expert-deadline')?.dataset.baseValue || card.querySelector('.order-expert-deadline')?.textContent?.trim() || '';
                const epText = card.querySelector('.order-expert-pay')?.textContent?.trim() || card.dataset.expertPayout || 'INR 0';
                detailsGrid.className = 'mt-1 grid grid-cols-2 gap-x-3 gap-y-0.5 text-[11px] text-gray-600';
                detailsGrid.innerHTML = `
                    <p>C.A - <span class="order-amount font-semibold text-indigo-600">${escapeHtml(caText)}</span><span class="order-payment-indicator hidden ml-1"></span></p>
                    <p>E.P - <span class="order-expert-pay font-semibold text-violet-700">${escapeHtml(epText || 'INR 0')}</span><span class="order-expert-payment-indicator hidden ml-1"></span></p>
                    <p>C.D - <span class="order-deadline deadline-muted font-medium">${escapeHtml(cdRaw)}</span></p>
                    <p>E.D - <span class="order-expert-deadline deadline-muted font-medium">${escapeHtml(edRaw)}</span></p>
                `;
                detailsGrid.querySelector('.order-deadline').dataset.baseValue = cdRaw;
                detailsGrid.querySelector('.order-expert-deadline').dataset.baseValue = edRaw;
            }
        }

        function syncOrderNotifyButton(card) {
            if (!card) return;
            const button = card.querySelector('.open-expert-notify-btn');
            if (!button) return;
            const isNotified = getNotifiedExpertIds(card).length > 0;
            button.classList.toggle('is-notified', isNotified);
            button.title = isNotified ? 'Experts notified' : 'Notify Experts';
            button.innerHTML = isNotified
                ? '<i data-lucide="bell-ring" class="w-3.5 h-3.5"></i>'
                : '<i data-lucide="bell" class="w-3.5 h-3.5"></i>';
            lucide.createIcons({ attrs: { 'stroke-width': 2 } });
        }

        function addTransactionToActiveCard() {
            if (!activeOrderCard) return;
            const transactionId = String(trTransactionIdInput?.value || '').trim();
            const amount = String(trAmountPaidInput?.value || '').trim();
            const currency = String(trCurrencyInput?.value || 'INR').trim().toUpperCase() || 'INR';
            const paymentMethod = String(trPaymentMethodInput?.value || '').trim();
            const note = String(trNoteInput?.value || '').trim();
            const agent = String(trAgentNameInput?.value || agentName).trim() || agentName;
            if (!transactionId) {
                alert('Transaction ID is required.');
                return;
            }
            if (!paymentMethod) {
                alert('Please select payment method.');
                return;
            }
            const rows = parseOrderTransactions(activeOrderCard.dataset.transactionRecords || '[]');
            rows.unshift({
                agent,
                currency,
                amount,
                transactionId,
                paymentMethod,
                note,
                createdAt: new Date().toISOString()
            });
            activeOrderCard.dataset.transactionRecords = JSON.stringify(rows);
            renderTransactionRecordsModal(rows);
            if (trTransactionIdInput) trTransactionIdInput.value = '';
            if (trAmountPaidInput) trAmountPaidInput.value = '';
            if (trCurrencyInput) trCurrencyInput.value = 'INR';
            if (trPaymentMethodInput) trPaymentMethodInput.value = '';
            if (trNoteInput) trNoteInput.value = '';
            closeAddTransactionRecordModal();
            persistOrderListToStorage();
        }

        function removeTransactionFromActiveCard(index) {
            if (!activeOrderCard) return;
            const rows = parseOrderTransactions(activeOrderCard.dataset.transactionRecords || '[]')
                .filter((_, rowIndex) => rowIndex !== index);
            activeOrderCard.dataset.transactionRecords = JSON.stringify(rows);
            renderTransactionRecordsModal(rows);
            persistOrderListToStorage();
        }

        function openTransactionRecordsModal() {
            if (!activeOrderCard || !transactionRecordsModal) return;
            const identity = getOrderRecordIdentity(activeOrderCard);
            if (transactionRecordsMeta) {
                transactionRecordsMeta.textContent = `Order #${identity.orderId || '-'} | Client ${identity.clientId || '-'} | Manage payment records`;
            }
            renderTransactionRecordsModal(parseOrderTransactions(activeOrderCard.dataset.transactionRecords || '[]'));
            transactionRecordsModal.classList.remove('hidden');
        }

        function closeTransactionRecordsModal() {
            transactionRecordsModal?.classList.add('hidden');
        }

        function openAddTransactionRecordModal() {
            if (!activeOrderCard || !addTransactionRecordModal) return;
            if (trAgentNameInput) trAgentNameInput.value = String(activeOrderCard.dataset.createdBy || agentName).trim() || agentName;
            if (trTransactionIdInput) trTransactionIdInput.value = '';
            if (trAmountPaidInput) trAmountPaidInput.value = '';
            if (trCurrencyInput) trCurrencyInput.value = 'INR';
            if (trPaymentMethodInput) trPaymentMethodInput.value = '';
            if (trNoteInput) trNoteInput.value = '';
            addTransactionRecordModal.classList.remove('hidden');
        }

        function closeAddTransactionRecordModal() {
            addTransactionRecordModal?.classList.add('hidden');
        }

        function openOrderHelperPage(page) {
            if (!activeOrderCard) {
                alert('Open an order card first.');
                return;
            }
            const identity = getOrderRecordIdentity(activeOrderCard);
            if (!identity.orderId) {
                alert('Order ID not found.');
                return;
            }
            const targetUrl = `${page}?orderId=${encodeURIComponent(identity.orderId)}&clientId=${encodeURIComponent(identity.clientId || '')}`;
            window.open(targetUrl, '_blank', 'noopener,noreferrer');
        }

        function addAttachmentToActiveCard(kind) {
            if (!activeOrderCard) return;
            const isQuestion = kind === 'question';
            const nameInput = isQuestion ? odQuestionItemNameInput : odSolutionItemNameInput;
            const urlInput = isQuestion ? odQuestionItemUrlInput : odSolutionItemUrlInput;
            const target = isQuestion ? odQuestionAttachmentList : odSolutionAttachmentList;
            const datasetKey = isQuestion ? 'questionAttachments' : 'solutionAttachments';
            const name = String(nameInput?.value || '').trim();
            const url = String(urlInput?.value || '').trim();
            if (!name && !url) {
                alert('Add file name or URL first.');
                return;
            }
            const rows = parseOrderAttachments(activeOrderCard.dataset[datasetKey] || '[]');
            rows.unshift({
                name: name || url,
                url,
                addedAt: new Date().toISOString()
            });
            activeOrderCard.dataset[datasetKey] = JSON.stringify(rows);
            renderOrderAttachmentList(target, rows, kind);
            syncOrderExpertSummary(activeOrderCard);
            if (nameInput) nameInput.value = '';
            if (urlInput) urlInput.value = '';
            persistOrderListToStorage();
        }

        function removeAttachmentFromActiveCard(kind, index) {
            if (!activeOrderCard) return;
            const isQuestion = kind === 'question';
            const target = isQuestion ? odQuestionAttachmentList : odSolutionAttachmentList;
            const datasetKey = isQuestion ? 'questionAttachments' : 'solutionAttachments';
            const rows = parseOrderAttachments(activeOrderCard.dataset[datasetKey] || '[]')
                .filter((_, rowIndex) => rowIndex !== index);
            activeOrderCard.dataset[datasetKey] = JSON.stringify(rows);
            renderOrderAttachmentList(target, rows, kind);
            syncOrderExpertSummary(activeOrderCard);
            persistOrderListToStorage();
        }

        function syncOrderExpertSummary(card) {
            if (!card) return;
            ensureOrderCardEnhancements(card);
            syncOrderPaymentIndicator(card);
            syncOrderExpertPaymentIndicator(card);
        }

        function getFilteredExpertsForNotify(query) {
            const normalizedQuery = String(query || '').trim().toLowerCase();
            return readExpertsFromStorage()
                .map((row) => normalizeExpertRecord(row))
                .filter((expert) => expert.expertId && expert.name)
                .filter((expert) => {
                    if (!normalizedQuery) return true;
                    const haystack = [
                        expert.name,
                        expert.expertId,
                        expert.waId,
                        ...(expert.expertise || [])
                    ].join(' ').toLowerCase();
                    return haystack.includes(normalizedQuery);
                })
                .sort((a, b) => String(a.name || '').localeCompare(String(b.name || '')));
        }

        function renderExpertNotifyList() {
            if (!expertNotifyList) return;
            const rows = getFilteredExpertsForNotify(expertNotifySearch?.value || '');
            if (!rows.length) {
                expertNotifyList.innerHTML = '<div class="expert-notify-empty">No experts found for this search.</div>';
                if (expertNotifySelectAll) expertNotifySelectAll.checked = false;
                return;
            }

            expertNotifyList.innerHTML = rows.map((expert) => {
                const skills = expert.expertise.length ? expert.expertise.join(', ') : 'No skills added';
                const checked = selectedNotifyExpertIds.has(expert.expertId) ? 'checked' : '';
                const initials = getContactInitials(expert.name, expert.expertId);
                const rating = getExpertRatingDisplay(expert.rating || 0);
                const alreadyNotified = activeNotifyOrderCard ? getNotifiedExpertIds(activeNotifyOrderCard).includes(expert.expertId) : false;
                return `
                    <label class="expert-notify-row">
                        <div class="expert-notify-card-main">
                            <div class="expert-notify-avatar-wrap">
                                <div class="expert-notify-avatar" aria-hidden="true">${escapeHtml(initials)}</div>
                                ${alreadyNotified ? '<span class="expert-notify-sent-tick" title="Already notified">&#10003;</span>' : ''}
                            </div>
                            <div class="expert-notify-meta">
                                <div class="expert-notify-title-row">
                                    <p class="expert-notify-name">${escapeHtml(expert.name)}</p>
                                    <span class="expert-notify-pill">${escapeHtml(expert.expertId)}</span>
                                </div>
                                <p class="expert-notify-sub">${expert.waId ? escapeHtml(expert.waId) : 'WhatsApp not added'}</p>
                                <p class="expert-notify-skills">${escapeHtml(skills)}</p>
                                <div class="expert-notify-foot">
                                    <span class="expert-notify-rating">${escapeHtml(rating.stars)} <strong>${escapeHtml(rating.value)}/5</strong></span>
                                </div>
                            </div>
                        </div>
                        <span class="expert-notify-check-wrap">
                            <input type="checkbox" class="expert-notify-checkbox mt-1" data-expert-id="${escapeHtml(expert.expertId)}" ${checked}>
                        </span>
                    </label>
                `;
            }).join('');

            const visibleIds = rows.map((expert) => expert.expertId);
            if (expertNotifySelectAll) {
                expertNotifySelectAll.checked = visibleIds.length > 0 && visibleIds.every((id) => selectedNotifyExpertIds.has(id));
            }
        }

        function openExpertNotifyModal(card) {
            activeNotifyOrderCard = card;
            selectedNotifyExpertIds = new Set(getNotifiedExpertIds(card));
            const taskTitle = card.querySelector('.order-title')?.textContent?.trim() || 'Untitled Task';
            const orderId = (card.querySelector('.order-id')?.textContent || '').replace('#', '').trim();
            const serviceType = String(card.dataset.serviceType || '').trim();
            if (expertNotifyTaskMeta) {
                expertNotifyTaskMeta.textContent = `${taskTitle}${serviceType ? ` | ${serviceType}` : ''}${orderId ? ` | #${orderId}` : ''}`;
            }
            if (expertNotifySearch) expertNotifySearch.value = '';
            renderExpertNotifyList();
            expertNotifyModal?.classList.remove('hidden');
        }

        function closeExpertNotifyModal() {
            expertNotifyModal?.classList.add('hidden');
            activeNotifyOrderCard = null;
            selectedNotifyExpertIds = new Set();
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

        function getOrderCommentPreview(text, maxLength = 82) {
            const normalized = String(text || '').replace(/\s+/g, ' ').trim();
            if (!normalized) return '';
            if (normalized.length <= maxLength) return normalized;
            return normalized.slice(0, Math.max(0, maxLength - 3)).trimEnd() + '...';
        }

        function updateOrderCardPreview(card, overrides = {}) {
            if (!card) return;
            const title = String(overrides.title ?? card.querySelector('.order-title')?.textContent ?? '').trim();
            const status = String(overrides.status ?? card.dataset.status ?? card.querySelector('.order-status')?.textContent ?? 'New').trim() || 'New';
            const actualDeadlineRaw = String(overrides.actualDeadline ?? card.dataset.actualDeadline ?? card.querySelector('.order-deadline')?.dataset.baseValue ?? '').trim();
            const expertDeadlineRaw = String(overrides.expertDeadline ?? card.dataset.expertDeadline ?? card.querySelector('.order-expert-deadline')?.dataset.baseValue ?? '').trim();
            const commentPreview = String(overrides.commentPreview ?? card.dataset.commentPreview ?? '').trim();
            const createdBy = String(overrides.createdBy ?? card.dataset.createdBy ?? card.dataset.assignedTo ?? agentName).trim() || agentName;

            card.dataset.status = status;
            card.dataset.actualDeadline = actualDeadlineRaw;
            card.dataset.expertDeadline = expertDeadlineRaw;
            card.dataset.commentPreview = commentPreview;
            card.dataset.createdBy = createdBy;

            const titleEl = card.querySelector('.order-title');
            if (titleEl) titleEl.textContent = title || 'Untitled Task';

            const statusEl = card.querySelector('.order-status');
            if (statusEl) {
                statusEl.textContent = status;
                applyOrderStatusStyle(statusEl, status);
            }

            const createdByTopEl = card.querySelector('.order-created-by-top');
            if (createdByTopEl) {
                createdByTopEl.textContent = 'By - ' + createdBy;
            }

            const actualEl = card.querySelector('.order-deadline');
            if (actualEl) {
                actualEl.dataset.baseValue = formatDateTimeForCard(actualDeadlineRaw || '');
                actualEl.textContent = formatDateTimeForCard(actualDeadlineRaw || '');
            }

            const expertEl = card.querySelector('.order-expert-deadline');
            if (expertEl) {
                expertEl.dataset.baseValue = formatDateTimeForCard(expertDeadlineRaw || '');
                expertEl.textContent = formatDateTimeForCard(expertDeadlineRaw || '');
            }

            let commentEl = card.querySelector('.order-comment');
            if (!commentEl) {
                commentEl = document.createElement('p');
                commentEl.className = 'order-comment hidden mt-1 text-[11px] text-gray-700 border-t pt-1';
                card.appendChild(commentEl);
            }
            if (commentPreview) {
                commentEl.textContent = 'Note: ' + commentPreview;
                commentEl.classList.remove('hidden');
            } else {
                commentEl.textContent = 'Note: ';
                commentEl.classList.add('hidden');
            }

            applyDeadlineStyles(card);
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
            updateOrderCardPreview(activeOrderCard, {
                commentPreview: agentName + ' - ' + getOrderCommentPreview(commentText)
            });

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
            syncContactAssignmentBadge(item);
            updateContactStateUI(item);

            const menuToggleBtn = item.querySelector('.contact-menu-toggle');
            const assignAgentBtn = item.querySelector('.contact-action-assign-agent');
            const deleteContactBtn = item.querySelector('.contact-action-delete');
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

            if (assignAgentBtn) {
                assignAgentBtn.addEventListener('click', async (e) => {
                    e.stopPropagation();
                    try {
                        openAssignAgentModal('single', item.dataset.waId || '');
                        openContactMenuId = '';
                        item.classList.remove('is-menu-open');
                        item.querySelector('.contact-card-menu')?.classList.add('hidden');
                    } catch (err) {
                        alert('Assign to agent failed: ' + (err?.message || 'Unknown error'));
                    }
                });
            }

            if (deleteContactBtn) {
                deleteContactBtn.addEventListener('click', async (e) => {
                    e.stopPropagation();
                    if (!isAdminUser) return;
                    const waId = normalizeWaId(item.dataset.waId || '');
                    if (!waId) return;
                    const displayName = item.querySelector('.contact-name')?.textContent?.trim() || waId;
                    const confirmed = window.confirm(`Delete contact "${displayName}" and its chat history?`);
                    if (!confirmed) return;
                    try {
                        await deleteWhatsappContact(waId);
                        removeContactFromUi(waId);
                    } catch (err) {
                        alert('Delete contact failed: ' + (err?.message || 'Unknown error'));
                    }
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
                updateAiAssistButtonState(item);
                updateChatAssignmentState(waId);
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
            const activeItem = activeWhatsappWaId
                ? contactList.querySelector('.contact-item[data-wa-id="' + normalizeWaId(activeWhatsappWaId) + '"]')
                : null;
            const baseName = getContactBaseName(activeItem?.dataset?.profileName || '', activeWhatsappWaId || '');
            return baseName || normalizeWaId(activeWhatsappWaId);
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
            const canMessage = canCurrentAgentMessageContact(activeWhatsappWaId);
            if (sendBtn) {
                sendBtn.disabled = whatsappSendInFlight || !canMessage;
                sendBtn.classList.toggle('is-loading', whatsappSendInFlight);
                sendBtn.classList.toggle('is-disabled', whatsappSendInFlight);
                sendBtn.setAttribute('aria-busy', whatsappSendInFlight ? 'true' : 'false');
                sendBtn.title = whatsappSendInFlight ? (label || 'Sending...') : 'Send message';
            }
            if (notifyBtn) notifyBtn.disabled = whatsappSendInFlight || !canMessage;
            if (attachFileBtn) attachFileBtn.disabled = whatsappSendInFlight || !canMessage;
            if (chatInput) chatInput.readOnly = whatsappSendInFlight || !canMessage;
            updateChatAssignmentState(activeWhatsappWaId);
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

        function positionOpenMessageMenus() {
            if (!chatMessages) return;
            const chatRect = chatMessages.getBoundingClientRect();
            const viewportLeft = Math.max(0, chatRect.left);
            const viewportRight = Math.min(window.innerWidth, chatRect.right);
            const viewportTop = Math.max(0, chatRect.top);
            const viewportBottom = Math.min(window.innerHeight, chatRect.bottom);

            chatMessages.querySelectorAll('.chat-message-menu').forEach((menu) => {
                const row = menu.closest('.chat-message-row');
                const btn = row?.querySelector('.chat-message-menu-btn');
                if (!row || !btn) return;

                menu.style.left = '0px';
                menu.style.top = '0px';

                const rowRect = row.getBoundingClientRect();
                const btnRect = btn.getBoundingClientRect();
                const menuRect = menu.getBoundingClientRect();
                const gap = 4;
                const spaceRight = viewportRight - btnRect.right;
                const spaceLeft = btnRect.left - viewportLeft;
                const openToRight = spaceRight >= menuRect.width || spaceRight >= spaceLeft;
                const left = openToRight
                    ? (btnRect.right - rowRect.left + gap)
                    : (btnRect.left - rowRect.left - menuRect.width - gap);
                const spaceBelow = viewportBottom - btnRect.bottom;
                const spaceAbove = btnRect.top - viewportTop;
                const openBelow = spaceBelow >= menuRect.height || spaceBelow >= spaceAbove;
                const top = openBelow
                    ? (btnRect.bottom - rowRect.top + gap)
                    : (btnRect.top - rowRect.top - menuRect.height - gap);

                menu.style.left = `${left}px`;
                menu.style.top = `${top}px`;
                menu.dataset.horizontal = openToRight ? 'right' : 'left';
                menu.dataset.vertical = openBelow ? 'down' : 'up';
            });
        }

        function renderWhatsappMessages(waId) {
            if (!chatMessages) return;
            const previousWaId = lastRenderedChatWaId;
            const wasNearBottom = (chatMessages.scrollHeight - (chatMessages.scrollTop + chatMessages.clientHeight)) < 96;
            const previousOffsetFromBottom = Math.max(0, chatMessages.scrollHeight - chatMessages.scrollTop);
            const shouldStickToBottom = !waId || previousWaId !== waId || wasNearBottom;
            if (!waId) {
                chatMessages.classList.add('is-empty-chat');
                chatMessages.innerHTML = '<div class="chat-empty-state" aria-hidden="true"></div>';
                lastRenderedChatWaId = '';
                updateChatWindowState('');
                return;
            }
            const messages = whatsappMessagesByContact[waId] || [];
            chatMessages.classList.remove('is-empty-chat');
            chatMessages.innerHTML = '';
            if (!messages.length) {
                chatMessages.innerHTML = '<p class="text-sm text-gray-500">No WhatsApp messages yet.</p>';
                lastRenderedChatWaId = waId;
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
                const isHandoffMessage = String(messageType).toLowerCase() === 'handoff';
                const replyQuote = msg.replyTo?.id
                    ? `
                        <button type="button" class="chat-reply-quote" data-reply-target-id="${escapeHtml(msg.replyTo.id || '')}" data-reply-target-key="${escapeHtml(msg.replyTo.clientKey || '')}">
                            <strong>${escapeHtml(msg.replyTo.direction === 'incoming' ? 'Client' : 'You')}</strong>
                            <span>${escapeHtml(msg.replyTo.text || '(message)')}</span>
                        </button>
                    `
                    : '';
                const bubbleBody = isMediaMessageType(messageType)
                    ? (messageType === 'image' && attachmentUrl
                        ? `
                            <div class="chat-bubble-file chat-bubble-media">
                                <a href="${escapeHtml(attachmentUrl)}" target="_blank" rel="noopener noreferrer" class="chat-media-link" title="Open image">
                                    <img src="${escapeHtml(attachmentUrl)}" alt="${attachmentName || 'Image'}" class="chat-media-preview chat-media-image">
                                </a>
                                <div class="chat-media-actions">
                                    <a href="${escapeHtml(attachmentUrl)}" target="_blank" rel="noopener noreferrer">Open</a>
                                    <a href="${escapeHtml(attachmentUrl)}" target="_blank" rel="noopener noreferrer" download>Save</a>
                                </div>
                                ${attachmentName ? `<p class="chat-media-name">${attachmentName}</p>` : ''}
                                ${msg.text ? `<p class="text-xs text-gray-600">${escapeHtml(msg.text)}</p>` : ''}
                            </div>
                        `
                        : (messageType === 'video' && attachmentUrl)
                            ? `
                                <div class="chat-bubble-file chat-bubble-media">
                                    <a href="${escapeHtml(attachmentUrl)}" target="_blank" rel="noopener noreferrer" class="chat-media-link" title="Open video">
                                        <video class="chat-media-preview chat-media-video" preload="metadata" muted playsinline>
                                            <source src="${escapeHtml(attachmentUrl)}" type="${escapeHtml(msg.mimeType || 'video/mp4')}">
                                        </video>
                                    </a>
                                    <div class="chat-media-actions">
                                        <a href="${escapeHtml(attachmentUrl)}" target="_blank" rel="noopener noreferrer">Open</a>
                                        <a href="${escapeHtml(attachmentUrl)}" target="_blank" rel="noopener noreferrer" download>Save</a>
                                    </div>
                                    ${attachmentName ? `<p class="chat-media-name">${attachmentName}</p>` : ''}
                                    ${msg.text ? `<p class="text-xs text-gray-600">${escapeHtml(msg.text)}</p>` : ''}
                                </div>
                            `
                            : (messageType === 'audio' && attachmentUrl)
                                ? `
                                    <div class="chat-bubble-file chat-bubble-audio">
                                        <p class="chat-media-name">${attachmentName || 'Audio message'}</p>
                                        <audio controls preload="metadata" class="chat-audio-player">
                                            <source src="${escapeHtml(attachmentUrl)}" type="${escapeHtml(msg.mimeType || 'audio/mpeg')}">
                                        </audio>
                                        <div class="chat-media-actions">
                                            <a href="${escapeHtml(attachmentUrl)}" target="_blank" rel="noopener noreferrer">Open</a>
                                            <a href="${escapeHtml(attachmentUrl)}" target="_blank" rel="noopener noreferrer" download>Save</a>
                                        </div>
                                        ${msg.text ? `<p class="text-xs text-gray-600">${escapeHtml(msg.text)}</p>` : ''}
                                    </div>
                                `
                                : `
                                    <div class="chat-bubble-file chat-bubble-document">
                                        <div class="chat-doc-card">
                                            <div class="chat-doc-icon" aria-hidden="true">
                                                <svg viewBox="0 0 24 24" fill="none">
                                                    <path d="M7 3.75h7.25L19.5 9v11.25A1.75 1.75 0 0 1 17.75 22h-10.5A1.75 1.75 0 0 1 5.5 20.25V5.5A1.75 1.75 0 0 1 7.25 3.75Z" fill="currentColor" opacity="0.14"></path>
                                                    <path d="M14 3.75v4.5c0 .414.336.75.75.75h4.5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"></path>
                                                    <path d="M14.25 3.75H7.25A1.75 1.75 0 0 0 5.5 5.5v14.75A1.75 1.75 0 0 0 7.25 22h10.5a1.75 1.75 0 0 0 1.75-1.75V9l-5.25-5.25Z" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"></path>
                                                    <path d="M8.75 13h7.5M8.75 16.25h5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"></path>
                                                </svg>
                                            </div>
                                            <div class="chat-doc-copy">
                                                <p class="chat-doc-type">${escapeHtml(messageType)}</p>
                                                <p class="chat-doc-name">${attachmentName || 'Attachment'}</p>
                                            </div>
                                        </div>
                                        ${attachmentUrl ? `
                                            <div class="chat-media-actions">
                                                <a href="${escapeHtml(attachmentUrl)}" target="_blank" rel="noopener noreferrer">Open</a>
                                                <a href="${escapeHtml(attachmentUrl)}" target="_blank" rel="noopener noreferrer" download>Save</a>
                                            </div>
                                        ` : ''}
                                    </div>
                                `)
                    : (isTemplateMessage || isHandoffMessage)
                        ? `
                            <div class="chat-template-body">
                                <p class="chat-template-label">${isHandoffMessage ? 'Human Transfer' : 'Template Message'}</p>
                                <p class="chat-template-text">${escapeHtml(msg.text || (isHandoffMessage ? '[Chat transferred to human]' : '[Template message]'))}</p>
                            </div>
                        `
                    : escapeHtml(msg.text || '[Unsupported message type]');
                const bubbleClass = (isTemplateMessage || isHandoffMessage)
                    ? 'chat-bubble-template p-2.5 shadow-sm text-sm text-gray-800'
                    : `${incoming ? 'chat-bubble-client' : 'chat-bubble-admin'} p-3 max-w-md shadow-sm text-sm text-gray-800`;
                wrap.className = 'flex flex-col ' + ((isTemplateMessage || isHandoffMessage) ? 'items-center' : (incoming ? 'items-start' : 'items-end'));
                wrap.innerHTML = `
                        <div class="chat-message-row" data-message-id="${escapeHtml(msg.id || '')}" data-message-key="${escapeHtml(clientKey)}">
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
            lastRenderedChatWaId = waId;
            if (shouldStickToBottom) {
                chatMessages.scrollTop = chatMessages.scrollHeight;
            } else {
                chatMessages.scrollTop = Math.max(0, chatMessages.scrollHeight - previousOffsetFromBottom);
            }
            positionOpenMessageMenus();
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
            maybeCreateAiTaskCard(waId);
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
            const nextAssignedAgent = String(contact.assignedAgent || savedContact?.assignedAgent || existing?.dataset.assignedAgent || '');
            const nextAssignedAgentKey = String(contact.assignedAgentKey || savedContact?.assignedAgentKey || existing?.dataset.assignedAgentKey || '').trim().toLowerCase();
            const nextAssignedAgentEmail = String(contact.assignedAgentEmail || savedContact?.assignedAgentEmail || existing?.dataset.assignedAgentEmail || '').trim().toLowerCase();
            const contactId = getOrCreateWhatsappContactId(normalizedWaId);
            const linkedExpert = getExpertByWaId(normalizedWaId);
            const expertId = String(linkedExpert?.expertId || '');
            saveManualContact({ waId: normalizedWaId, profileName, tag: nextTag, universityName: nextUniversityName, semester: nextSemester, timezone: nextTimezone, assignedAgent: nextAssignedAgent, assignedAgentKey: nextAssignedAgentKey, assignedAgentEmail: nextAssignedAgentEmail });
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
                item.dataset.assignedAgent = nextAssignedAgent;
                item.dataset.assignedAgentKey = nextAssignedAgentKey;
                item.dataset.assignedAgentEmail = nextAssignedAgentEmail;
                item.dataset.expertId = expertId;
                item.dataset.aiMode = String(contact.aiState?.aiEnabled ? 'ai' : 'human');
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
                                    ${expertId ? `<span class="contact-expert-id-badge" title="Expert ID">${escapeHtml(expertId)}</span>` : ''}
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
                        <button type="button" class="contact-action-assign-agent">Assign to Agent</button>
                        ${isAdminUser ? '<button type="button" class="contact-action-delete text-red-600">Delete Contact</button>' : ''}
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
                existing.dataset.assignedAgent = nextAssignedAgent;
                existing.dataset.assignedAgentKey = nextAssignedAgentKey;
                existing.dataset.assignedAgentEmail = nextAssignedAgentEmail;
                existing.dataset.expertId = expertId;
                existing.dataset.aiMode = String(contact.aiState?.aiEnabled ? 'ai' : existing.dataset.aiMode || 'human');
                const metaEl = existing.querySelector('.contact-meta') || existing.querySelector('p');
                if (metaEl) {
                    metaEl.classList.add('contact-meta');
                    metaEl.textContent = 'ID: ' + contactId;
                }
                let expertBadgeEl = existing.querySelector('.contact-expert-id-badge');
                if (expertId) {
                    if (!expertBadgeEl) {
                        expertBadgeEl = document.createElement('span');
                        expertBadgeEl.className = 'contact-expert-id-badge';
                        existing.querySelector('.contact-secondary-row')?.insertBefore(
                            expertBadgeEl,
                            existing.querySelector('.contact-label-badge')
                        );
                    }
                    expertBadgeEl.textContent = expertId;
                    expertBadgeEl.title = 'Expert ID';
                } else if (expertBadgeEl) {
                    expertBadgeEl.remove();
                }
            }
            syncContactAssignmentBadge(existing);
            const tagSelect = existing.querySelector('.contact-menu-label');
            if (tagSelect) tagSelect.value = nextTag;
            if (activeWhatsappWaId && activeWhatsappWaId === normalizedWaId) {
                setActiveContactItem(existing);
                updateAiAssistButtonState(existing);
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
                senderType: payload.senderType || '',
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
                            senderType: m.senderType || '',
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
                        updatedAt: entry.updatedAt || '',
                        aiState: entry.aiState || null
                    }, false);
                    const item = contactList.querySelector('.contact-item[data-wa-id="' + waId + '"]');
                    if (item) {
                        const nextUnreadCount = activeWhatsappWaId === waId ? 0 : getUnreadCountForThread(waId, whatsappMessagesByContact[waId] || []);
                        setContactUnreadCount(item, nextUnreadCount);
                    }
                    maybeCreateAiTaskCard(waId);
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
            initCrmStatePolling();
        }

        contactItems.forEach((item, index) => {
            item.dataset.order = String(index);
            wireContactItem(item);
        });

        migrateWhatsappContactIdsToSixDigitSeries();
        experts = readExpertsFromStorage();
        setActiveContactItem(null);
        updateTaskClientInputFromContact(null);
        updateAiAssistButtonState(null);
        updateChatAssignmentState('');
        restoreManualContactsFromStorage();
        scheduleContactOrderRefresh();
        initWhatsappSync();
        void hydrateCrmStateFromBackend().then((loaded) => {
            crmStateHydrationComplete = true;
            if (loaded && crmStatePendingSave) {
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

        if (adminQuickMenuBtn && isAdminUser) {
            adminQuickMenuBtn.addEventListener('click', (event) => {
                event.stopPropagation();
                const willOpen = adminQuickMenu?.classList.contains('hidden');
                adminQuickMenu?.classList.toggle('hidden', !willOpen);
                adminQuickMenuBtn.setAttribute('aria-expanded', willOpen ? 'true' : 'false');
            });
        }

        if (navExportClientsBtn && isAdminUser) {
            navExportClientsBtn.addEventListener('click', () => {
                adminQuickMenu?.classList.add('hidden');
                adminQuickMenuBtn?.setAttribute('aria-expanded', 'false');
                openExportClientsModal();
            });
        }

        if (navExportExpertsBtn && isAdminUser) {
            navExportExpertsBtn.addEventListener('click', () => {
                adminQuickMenu?.classList.add('hidden');
                adminQuickMenuBtn?.setAttribute('aria-expanded', 'false');
                downloadExpertsExportCsv();
            });
        }

        if (navImportExpertsBtn && isAdminUser) {
            navImportExpertsBtn.addEventListener('click', () => {
                adminQuickMenu?.classList.add('hidden');
                adminQuickMenuBtn?.setAttribute('aria-expanded', 'false');
                importExpertsFileInput?.click();
            });
        }

        if (navMonthlyDataBtn && isAdminUser) {
            navMonthlyDataBtn.addEventListener('click', () => {
                adminQuickMenu?.classList.add('hidden');
                adminQuickMenuBtn?.setAttribute('aria-expanded', 'false');
                openMonthlyDataModal();
            });
        }

        assignActiveChatAgentBtn?.addEventListener('click', async () => {
            if (!contactList?.querySelector('.contact-item')) return;
            openAssignAgentModal('all', activeWhatsappWaId);
        });

        closeAssignAgentModalBtn?.addEventListener('click', closeAssignAgentModal);
        cancelAssignAgentModalBtn?.addEventListener('click', closeAssignAgentModal);
        assignAgentModal?.addEventListener('click', (event) => {
            if (event.target === assignAgentModal) {
                closeAssignAgentModal();
            }
        });
        assignAgentList?.addEventListener('click', (event) => {
            const deleteTarget = event.target.closest('[data-delete-agent-name]');
            if (deleteTarget) {
                event.stopPropagation();
                if (!isAdminUser) return;
                const targetName = normalizeAgentName(deleteTarget.getAttribute('data-delete-agent-name') || '');
                if (!targetName) return;
                const nextRoster = ensureAgentRoster().filter((entry) => getAgentKey(entry) !== getAgentKey(targetName));
                writeAgentRosterToStorage(nextRoster);
                if (getAgentKey(selectedAssignAgentName) === getAgentKey(targetName)) {
                    selectedAssignAgentName = nextRoster[0] || normalizeAgentName(storedAgentName || 'Agent');
                }
                if (getAgentKey(agentName) === getAgentKey(targetName)) {
                    setCurrentAgentName(nextRoster[0] || normalizeAgentName(storedAgentName || 'Agent'));
                }
                renderAssignAgentModal();
                return;
            }
            const option = event.target.closest('[data-agent-name]');
            if (!option) return;
            selectedAssignAgentName = normalizeAgentName(option.getAttribute('data-agent-name') || '');
            renderAssignAgentModal();
        });
        addAssignAgentNameBtn?.addEventListener('click', () => {
            if (!isAdminUser) return;
            const nextName = normalizeAgentName(assignAgentNameInput?.value || '');
            if (!nextName) {
                alert('Enter an agent name first.');
                return;
            }
            const nextRoster = ensureAgentRoster();
            if (!nextRoster.some((entry) => getAgentKey(entry) === getAgentKey(nextName))) {
                writeAgentRosterToStorage([...nextRoster, nextName]);
            }
            selectedAssignAgentName = nextName;
            if (assignAgentNameInput) assignAgentNameInput.value = '';
            renderAssignAgentModal();
        });
        confirmAssignAgentBtn?.addEventListener('click', async () => {
            const nextName = normalizeAgentName(selectedAssignAgentName || agentName);
            if (!nextName) {
                alert('Select an agent first.');
                return;
            }
            setCurrentAgentName(nextName);
            try {
                if (assignAgentModalMode === 'single' && assignAgentModalWaId) {
                    await assignContactToAgent(assignAgentModalWaId, nextName);
                } else {
                    await assignAllChatsToAgent(nextName);
                }
                closeAssignAgentModal();
            } catch (err) {
                alert('Assign agent failed: ' + (err?.message || 'Unknown error'));
            }
        });

        if (importExpertsFileInput) {
            importExpertsFileInput.addEventListener('change', async (event) => {
                const file = event.target?.files?.[0];
                if (!file) return;
                const fileName = String(file.name || '').toLowerCase();
                if (!fileName.endsWith('.csv')) {
                    alert('Please import a CSV file. In Excel, use Save As -> CSV UTF-8.');
                    importExpertsFileInput.value = '';
                    return;
                }
                try {
                    const text = await file.text();
                    handleExpertsImportText(text);
                } catch (err) {
                    alert('Import failed: ' + (err?.message || 'Unknown error'));
                } finally {
                    importExpertsFileInput.value = '';
                }
            });
        }

        if (closeExportClientsModalBtn) {
            closeExportClientsModalBtn.addEventListener('click', () => {
                closeExportClientsModal();
            });
        }

        if (cancelExportClientsBtn) {
            cancelExportClientsBtn.addEventListener('click', () => {
                closeExportClientsModal();
            });
        }

        if (confirmExportClientsBtn) {
            confirmExportClientsBtn.addEventListener('click', () => {
                const fromValue = exportClientsFromDateInput?.value || '';
                const toValue = exportClientsToDateInput?.value || '';
                const fromMs = parseDateInputStart(fromValue);
                const toMs = parseDateInputEnd(toValue);

                if (fromMs != null && toMs != null && fromMs > toMs) {
                    alert('From date cannot be later than To date.');
                    return;
                }

                const rows = buildClientExportRows(fromMs, toMs);
                if (!rows.length) {
                    alert('No client data found for the selected date range.');
                    return;
                }

                downloadClientExportCsv(rows, fromValue, toValue);
                closeExportClientsModal();
            });
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
        orderSearch.addEventListener('input', applyOrderSearchFilter);
        if (orderDeliveryWindow) {
            orderDeliveryWindow.addEventListener('change', applyOrderSearchFilter);
        }
        if (monthlyDataMonthInput && !monthlyDataMonthInput.value) {
            monthlyDataMonthInput.value = getDefaultMonthlyDataMonth();
        }
        monthlyDataMonthInput?.addEventListener('change', renderMonthlyDataDashboard);
        monthlyDataMonthModalInput?.addEventListener('change', renderMonthlyDataModalContent);
        closeMonthlyDataModalBtn?.addEventListener('click', closeMonthlyDataModal);
        [orderTabMine, orderTabExpert, orderTabAll, orderTabMonthly].forEach((btn) => {
            if (!btn) return;
            btn.addEventListener('click', () => setOrderTab(btn.dataset.orderTab || 'mine'));
        });
        if (isAdminUser && initialViewMode === 'monthly-data') {
            setOrderTab('monthly');
        }
        applyExpertDeadlineConstraint();
        if (odActualDeadlineInput) {
            odActualDeadlineInput.addEventListener('input', applyExpertDeadlineConstraint);
            odActualDeadlineInput.addEventListener('change', applyExpertDeadlineConstraint);
        }
        if (odExpertDeadlineInput) {
            odExpertDeadlineInput.addEventListener('change', applyExpertDeadlineConstraint);
        }
        taskServiceType.addEventListener('change', function() {
            const normalizedServiceType = this.value.toLowerCase();
            const isLiveSession = normalizedServiceType === 'live session' || normalizedServiceType === 'tutoring class';
            liveSessionCreateFields.classList.toggle('hidden', !isLiveSession);
        });

        // Create Agent Button (if present)
        const createAgentBtn = document.getElementById('createAgentBtn');
        if (createAgentBtn) {
            createAgentBtn.onclick = function() {
                alert('Admin: Create Agent ID');
            };
        }

        function createOrderCard(options = {}) {
            const clientText = String(options.clientText || '').trim();
            const serviceType = String(options.serviceType || 'Assignment Help').trim() || 'Assignment Help';
            const title = String(options.title || serviceType || 'Untitled Task').trim();
            const clientId = String(options.clientId || '0000').replace(/[^\d]/g, '') || '0000';
            const orderId = String(options.orderId || getNextOrderId(clientId)).trim();
            const status = String(options.status || 'New').trim() || 'New';
            const assignedTo = String(options.assignedTo || agentName).trim();
            const createdBy = String(options.createdBy || agentName).trim();
            const actualDeadline = String(options.actualDeadline || '').trim();
            const expertDeadline = String(options.expertDeadline || '').trim();
            const expertPayout = String(options.expertPayout || 'INR 0').trim() || 'INR 0';
            const baseAmount = String(options.baseAmount || 'TBD').trim() || 'TBD';
            const sessionStart = String(options.sessionStart || '').trim();
            const sessionDuration = String(options.sessionDuration || '').trim();
            const sourceWaId = normalizeWaId(options.sourceWaId || '');
            const aiSummaryMessageId = String(options.aiSummaryMessageId || '').trim();
            const aiConfirmedMessageId = String(options.aiConfirmedMessageId || '').trim();
            const commentPreview = String(options.commentPreview || '').trim();
            const nowIso = new Date().toISOString();
            const normalizedServiceType = serviceType.toLowerCase();
            const isLiveSession = normalizedServiceType === 'live session' || normalizedServiceType === 'tutoring class';

            const card = document.createElement('div');
            card.className = 'order-card p-3 rounded-lg border shadow-sm';
            card.dataset.serviceType = serviceType;
            card.dataset.clientId = clientId;
            card.dataset.assignedTo = assignedTo;
            card.dataset.createdBy = createdBy;
            card.dataset.createdAt = nowIso;
            card.dataset.updatedAt = nowIso;
            card.dataset.instructions = '';
            card.dataset.activity = '';
            card.dataset.comments = '[]';
            card.dataset.notifiedExperts = '[]';
            card.dataset.interestedExperts = '[]';
            card.dataset.questionAttachments = '[]';
            card.dataset.solutionAttachments = '[]';
            card.dataset.customOrderFolders = '[]';
            card.dataset.customOrderFiles = '[]';
            card.dataset.transactionRecords = '[]';
            card.dataset.assignedExperts = '[]';
            card.dataset.assignedExpertName = '';
            card.dataset.assignedExpertId = '';
            card.dataset.expertNotes = '';
            card.dataset.expertDeadline = expertDeadline;
            card.dataset.expertPayout = expertPayout;
            card.dataset.expertPaymentStatus = 'pending';
            card.dataset.clientPaidCurrency = 'INR';
            card.dataset.clientPaidAmount = '';
            card.dataset.paymentStatusOverride = 'auto';
            card.dataset.sessionStart = sessionStart;
            card.dataset.sessionDuration = sessionDuration;
            card.dataset.status = status;
            card.dataset.actualDeadline = actualDeadline;
            card.dataset.commentPreview = commentPreview;
            if (sourceWaId) card.dataset.sourceWaId = sourceWaId;
            if (aiSummaryMessageId) card.dataset.aiSummaryMessageId = aiSummaryMessageId;
            if (aiConfirmedMessageId) card.dataset.aiConfirmedMessageId = aiConfirmedMessageId;
            card.innerHTML = `
                <div class="flex justify-between items-start mb-2">
                    <span class="order-detail-label text-[10px] font-bold uppercase px-2 py-0.5 rounded">Work Details</span>
                    <div class="flex items-center gap-2">
                        <span class="order-created-by-top text-[10px] text-gray-500 font-medium">By - ${escapeHtml(createdBy)}</span>
                        <span class="order-id text-[10px] text-gray-400 font-mono">#${escapeHtml(orderId)}</span>
                        <button type="button" class="open-expert-notify-btn p-1 rounded" title="Notify Experts">
                            <i data-lucide="bell" class="w-3.5 h-3.5"></i>
                        </button>
                        <button class="open-order-details-btn p-1 rounded bg-indigo-50 text-indigo-700 hover:bg-indigo-100" title="Open Details">
                            <i data-lucide="square-arrow-out-up-right" class="w-3.5 h-3.5"></i>
                        </button>
                    </div>
                </div>
                <h4 class="order-title text-sm font-bold text-gray-800">${escapeHtml(title || serviceType || 'Untitled Task')}</h4>
                <div class="order-title-row flex items-center gap-2">
                    <span class="order-service-tag text-[10px] font-semibold px-2 py-0.5 rounded-full bg-indigo-100 text-indigo-700">${escapeHtml(serviceType)}</span>
                    <span class="order-status text-[10px] font-bold uppercase px-2 py-0.5 rounded">${escapeHtml(status)}</span>
                </div>
                <p class="order-client-note text-[11px] text-gray-500 mt-0.5">${escapeHtml(clientText.replace(/\(\d+\)/, '').trim())}</p>
                <div class="mt-1 grid grid-cols-2 gap-x-3 gap-y-0.5 text-[11px] text-gray-600">
                    <p>C.A - <span class="order-amount font-semibold text-indigo-600">${escapeHtml(baseAmount)}</span><span class="order-payment-indicator hidden ml-1"></span></p>
                    <p>E.P - <span class="order-expert-pay font-semibold text-violet-700">${escapeHtml(expertPayout)}</span><span class="order-expert-payment-indicator hidden ml-1"></span></p>
                    <p>C.D - <span class="order-deadline deadline-muted font-medium">${escapeHtml(formatDateTimeForCard(actualDeadline || ''))}</span></p>
                    <p>E.D - <span class="order-expert-deadline deadline-muted font-medium">${escapeHtml(formatDateTimeForCard(expertDeadline || ''))}</span></p>
                </div>
                <p class="order-session mt-1 text-[11px] text-gray-600 ${isLiveSession ? '' : 'hidden'}">S.T - ${escapeHtml(sessionStart || 'TBD')} | Dur - ${escapeHtml(sessionDuration || 'TBD')}m</p>
                <p class="order-comment ${commentPreview ? '' : 'hidden'} mt-1 text-[11px] text-gray-700 border-t pt-1">${commentPreview ? escapeHtml('Note: ' + commentPreview) : 'Note: '}</p>
            `;
            card.querySelector('.order-deadline').dataset.baseValue = formatDateTimeForCard(actualDeadline || '');
            card.querySelector('.order-expert-deadline').dataset.baseValue = formatDateTimeForCard(expertDeadline || '');
            syncOrderServiceTag(card);
            syncOrderPaymentIndicator(card);
            syncOrderExpertPaymentIndicator(card);
            syncOrderExpertSummary(card);
            syncOrderNotifyButton(card);
            applyCardTypeStyle(card, serviceType);
            updateOrderCardPreview(card, {
                title,
                status,
                actualDeadline,
                expertDeadline,
                commentPreview
            });
            return card;
        }

        function parseAiTaskSummaryMessage(message) {
            const text = String(message?.text || '').trim();
            if (!text || !/short summary of your task/i.test(text)) return null;
            const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
            const summary = {
                summaryMessageId: String(message?.id || '').trim(),
                subjectTopic: '',
                workDetails: '',
                requirements: '',
                deadline: '',
                deliverables: '',
                bulletPoints: []
            };
            lines.forEach((line) => {
                const clean = line.replace(/^-+\s*/, '').trim();
                if (/^subject\/topic\s*:/i.test(clean)) summary.subjectTopic = clean.split(/:/).slice(1).join(':').trim();
                else if (/^work details\s*:/i.test(clean)) summary.workDetails = clean.split(/:/).slice(1).join(':').trim();
                else if (/^requirements\s*:/i.test(clean)) summary.requirements = clean.split(/:/).slice(1).join(':').trim();
                else if (/^deadline\s*:/i.test(clean)) summary.deadline = clean.split(/:/).slice(1).join(':').trim();
                else if (/^deliverables\s*:/i.test(clean)) summary.deliverables = clean.split(/:/).slice(1).join(':').trim();
                else if (line.startsWith('-')) summary.bulletPoints.push(clean);
            });
            if (!summary.subjectTopic && !summary.workDetails && !summary.requirements && !summary.deadline && !summary.deliverables) {
                return null;
            }
            return summary;
        }

        function isPositiveTaskConfirmation(text) {
            const normalized = String(text || '').trim().toLowerCase();
            if (!normalized) return false;
            return [
                'yes', 'y', 'yeah', 'yep', 'ok', 'okay', 'confirm', 'confirmed', 'correct',
                'looks correct', 'this is correct', 'go ahead', 'please proceed', 'proceed'
            ].some((term) => normalized === term || normalized.includes(term));
        }

        function maybeCreateAiTaskCard(waId) {
            const normalizedWaId = normalizeWaId(waId);
            if (!normalizedWaId) return false;
            const thread = Array.isArray(whatsappMessagesByContact[normalizedWaId]) ? whatsappMessagesByContact[normalizedWaId] : [];
            if (!thread.length) return false;

            let summaryMessage = null;
            let confirmationMessage = null;
            for (let index = thread.length - 1; index >= 0; index -= 1) {
                const message = thread[index];
                if (!confirmationMessage && message?.direction === 'incoming' && isPositiveTaskConfirmation(message?.text || '')) {
                    confirmationMessage = message;
                    continue;
                }
                if (!summaryMessage && message?.direction === 'outgoing' && String(message?.senderType || '').toLowerCase() === 'ai') {
                    const parsed = parseAiTaskSummaryMessage(message);
                    if (parsed) {
                        summaryMessage = { message, parsed, index };
                        break;
                    }
                }
            }
            if (!summaryMessage || !confirmationMessage) return false;
            if (thread.indexOf(confirmationMessage) < summaryMessage.index) return false;
            if (!summaryMessage.parsed.summaryMessageId) return false;
            if (Array.from(orderList.querySelectorAll('.order-card')).some((card) => String(card.dataset.aiSummaryMessageId || '') === summaryMessage.parsed.summaryMessageId)) {
                return false;
            }

            const contactItem = contactList.querySelector(`.contact-item[data-wa-id="${normalizedWaId}"]`);
            const clientId = String(contactItem?.dataset?.contactId || getOrCreateWhatsappContactId(normalizedWaId));
            const clientName = contactItem
                ? getContactDisplayName(contactItem.dataset.profileName || normalizedWaId, normalizedWaId)
                : normalizedWaId;
            const title = summaryMessage.parsed.workDetails || summaryMessage.parsed.subjectTopic || 'Assignment Help';
            const serviceType = /session|class|tutor/i.test(title) ? 'Tutoring class' : 'Assignment Help';
            const card = createOrderCard({
                clientText: `${clientName} (${clientId})`,
                clientId,
                serviceType,
                title,
                status: 'New',
                createdBy: 'AI Agent',
                assignedTo: agentName,
                actualDeadline: summaryMessage.parsed.deadline,
                expertDeadline: summaryMessage.parsed.deadline,
                commentPreview: 'AI Agent - ' + getOrderCommentPreview([
                    summaryMessage.parsed.subjectTopic ? `Subject/Topic: ${summaryMessage.parsed.subjectTopic}` : '',
                    summaryMessage.parsed.requirements ? `Requirements: ${summaryMessage.parsed.requirements}` : '',
                    summaryMessage.parsed.deliverables ? `Deliverables: ${summaryMessage.parsed.deliverables}` : ''
                ].filter(Boolean).join(' | ')),
                sourceWaId: normalizedWaId,
                aiSummaryMessageId: summaryMessage.parsed.summaryMessageId,
                aiConfirmedMessageId: String(confirmationMessage.id || '').trim()
            });
            const requirementsText = [
                summaryMessage.parsed.subjectTopic ? `Subject/Topic: ${summaryMessage.parsed.subjectTopic}` : '',
                summaryMessage.parsed.requirements ? `Requirements: ${summaryMessage.parsed.requirements}` : '',
                summaryMessage.parsed.deliverables ? `Deliverables: ${summaryMessage.parsed.deliverables}` : ''
            ].filter(Boolean).join(' | ');
            card.dataset.instructions = requirementsText;
            if (requirementsText) {
                card.dataset.comments = JSON.stringify([{
                    agent: 'AI Agent',
                    text: requirementsText,
                    createdAt: new Date().toISOString()
                }]);
                const commentEl = card.querySelector('.order-comment');
                if (commentEl) {
                    commentEl.textContent = 'Note: AI Agent - ' + requirementsText;
                    commentEl.classList.remove('hidden');
                }
            }
            orderList.prepend(card);
            lucide.createIcons();
            persistOrderListToStorage();
            applyOrderSearchFilter();
            return true;
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
            const normalizedServiceType = taskServiceType.value.toLowerCase();
            const isLiveSession = normalizedServiceType === 'live session' || normalizedServiceType === 'tutoring class';
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
            const normalizedServiceType = serviceType.toLowerCase();
            const isLiveSession = normalizedServiceType === 'live session' || normalizedServiceType === 'tutoring class';
            const sessionStart = (taskSessionStart.value || '').trim();
            const sessionDuration = (taskSessionDuration.value || '').trim();
            const card = createOrderCard({
                clientText,
                clientId,
                serviceType,
                title,
                orderId,
                status: 'New',
                createdBy: agentName,
                assignedTo: agentName,
                sessionStart: isLiveSession ? sessionStart : '',
                sessionDuration: isLiveSession ? sessionDuration : ''
            });
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
            const addContactBtn = e.target.closest('.expert-add-contact-btn');
            if (addContactBtn) {
                addExpertToContacts(addContactBtn.dataset.expertId || '');
                return;
            }

            const addSkillBtn = e.target.closest('.expert-add-skill-btn');
            if (addSkillBtn) {
                promptAddSkillToExpert(addSkillBtn.dataset.expertId || '');
                return;
            }

            const deleteExpertBtn = e.target.closest('.expert-delete-btn');
            if (deleteExpertBtn) {
                deleteExpert(deleteExpertBtn.dataset.expertId || '');
                return;
            }

            const notifyBtn = e.target.closest('.open-expert-notify-btn');
            if (notifyBtn) {
                const card = notifyBtn.closest('.order-card');
                if (!card || card.classList.contains('expert-card')) return;
                openExpertNotifyModal(card);
                return;
            }

            const btn = e.target.closest('.open-order-details-btn');
            if (!btn) return;

            const card = btn.closest('.order-card');
            if (!card) return;
            openOrderDetailsModal(card);
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
                clearOrderDraft(activeOrderCard);
                persistOrderListToStorage();
                applyOrderSearchFilter();
                closeOrderDetailsModal();
                activeOrderCard = null;
                return;
            }
            const orderId = (activeOrderCard.querySelector('.order-id')?.textContent || '').replace('#', '').trim();
            const clientId = (activeOrderCard.dataset.clientId || '').trim();
            const assignedTo = document.getElementById('odAssignedTo').value.trim();
            const createdBy = (activeOrderCard.dataset.createdBy || agentName).trim();
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
            const expertPaymentStatus = (odExpertPaymentStatusInput?.value || 'pending').trim() || 'pending';
            const sessionStart = document.getElementById('odSessionStart').value.trim();
            const sessionDuration = document.getElementById('odSessionDuration').value.trim();

            activeOrderCard.dataset.serviceType = serviceType || title || '';
            syncOrderServiceTag(activeOrderCard);
            applyCardTypeStyle(activeOrderCard, activeOrderCard.dataset.serviceType);
            activeOrderCard.querySelector('.order-id').textContent = '#' + (orderId || clientId || 'NEW');
            activeOrderCard.querySelector('.order-amount').textContent = baseAmount;
            const expertPayEl = activeOrderCard.querySelector('.order-expert-pay');
            if (expertPayEl) {
                expertPayEl.textContent = expertPayout ? ('INR ' + expertPayout) : 'INR 0';
            }
            activeOrderCard.dataset.clientPaidCurrency = clientPaidCurrency;
            activeOrderCard.dataset.clientPaidAmount = clientPaidAmount;
            activeOrderCard.dataset.paymentStatusOverride = paymentStatusOverride;
            activeOrderCard.dataset.expertPaymentStatus = expertPaymentStatus;
            activeOrderCard.dataset.updatedAt = new Date().toISOString();
            syncOrderPaymentIndicator(activeOrderCard);
            syncOrderExpertPaymentIndicator(activeOrderCard);
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
            updateOrderCardPreview(activeOrderCard, {
                title: title || serviceType || 'Untitled',
                status: status || 'In Progress',
                actualDeadline,
                expertDeadline,
                createdBy: createdBy || agentName,
                commentPreview: latestComment ? `${latestComment.agent} - ${getOrderCommentPreview(latestComment.text)}` : ''
            });

            activeOrderCard.dataset.assignedTo = assignedTo;
            activeOrderCard.dataset.createdBy = createdBy || agentName;
            activeOrderCard.dataset.clientId = clientId;
            activeOrderCard.dataset.expertDeadline = expertDeadline;
            activeOrderCard.dataset.expertPayout = expertPayout ? ('INR ' + expertPayout) : '';
            activeOrderCard.dataset.additionalCharges = additionalCharges;
            activeOrderCard.dataset.sessionStart = sessionStart;
            activeOrderCard.dataset.sessionDuration = sessionDuration;
            activeOrderCard.dataset.activity = activeOrderCard.dataset.activity || '';
            activeOrderCard.dataset.comments = JSON.stringify(comments);
            syncOrderExpertSummary(activeOrderCard);
            updateOrderDetailsSummary(activeOrderCard);
            document.getElementById('odActivity').value = '';

            clearOrderDraft(activeOrderCard);
            persistOrderListToStorage();
            closeOrderDetailsModal();
        };

        document.getElementById('closeOrderDetailsBtn').onclick = function() {
            closeOrderDetailsModal();
        };
        openAttachmentPageBtn?.addEventListener('click', function() {
            openOrderHelperPage('order-attachments.html');
        });
        openManageExpertPageBtn?.addEventListener('click', function() {
            openManageExpertsModal();
        });
        openRecordTransactionBtn?.addEventListener('click', function() {
            openTransactionRecordsModal();
        });
        closeManageExpertsModalBtn?.addEventListener('click', closeManageExpertsModal);
        manageExpertsModal?.addEventListener('click', function(event) {
            if (event.target !== manageExpertsModal) return;
            closeManageExpertsModal();
        });
        [manageExpertsTabNotified, manageExpertsTabInterested, manageExpertsTabAssigned].forEach((btn) => {
            btn?.addEventListener('click', function() {
                activeManageExpertsTab = btn.dataset.manageExpertsTab || 'notified';
                renderManageExpertsModal();
            });
        });
        manageExpertsList?.addEventListener('click', function(event) {
            const actionBtn = event.target.closest('[data-manage-expert-action]');
            if (!actionBtn || !activeOrderCard) return;
            const action = String(actionBtn.dataset.manageExpertAction || '').trim();
            const expertId = String(actionBtn.dataset.expertId || '').trim();
            if (!expertId) return;
            const notifiedIds = getNotifiedExpertIds(activeOrderCard);
            const interestedIds = getInterestedExpertIds(activeOrderCard);
            const assignedIds = getAssignedExpertIds(activeOrderCard);

            if (action === 'mark-interested') {
                if (!notifiedIds.includes(expertId)) notifiedIds.push(expertId);
                if (!interestedIds.includes(expertId)) interestedIds.push(expertId);
                setNotifiedExpertIds(activeOrderCard, notifiedIds);
                setInterestedExpertIds(activeOrderCard, interestedIds);
            } else if (action === 'remove-interest') {
                setInterestedExpertIds(activeOrderCard, interestedIds.filter((id) => id !== expertId));
            } else if (action === 'assign') {
                const nextInterested = interestedIds.filter((id) => id !== expertId);
                const nextAssigned = assignedIds.includes(expertId) ? assignedIds : [expertId, ...assignedIds];
                setInterestedExpertIds(activeOrderCard, nextInterested);
                setAssignedExpertIds(activeOrderCard, nextAssigned);
            } else if (action === 'unassign') {
                setAssignedExpertIds(activeOrderCard, assignedIds.filter((id) => id !== expertId));
            }

            persistOrderListToStorage();
            renderManageExpertsModal();
        });
        closeTransactionRecordsModalBtn?.addEventListener('click', closeTransactionRecordsModal);
        transactionRecordsModal?.addEventListener('click', function(event) {
            if (event.target !== transactionRecordsModal) return;
            closeTransactionRecordsModal();
        });
        openAddTransactionRecordBtn?.addEventListener('click', function() {
            openAddTransactionRecordModal();
        });
        closeAddTransactionRecordModalBtn?.addEventListener('click', closeAddTransactionRecordModal);
        cancelAddTransactionRecordBtn?.addEventListener('click', closeAddTransactionRecordModal);
        saveTransactionRecordBtn?.addEventListener('click', function() {
            addTransactionToActiveCard();
        });
        addTransactionRecordModal?.addEventListener('click', function(event) {
            if (event.target !== addTransactionRecordModal) return;
            closeAddTransactionRecordModal();
        });
        transactionRecordsList?.addEventListener('click', function(event) {
            const removeBtn = event.target.closest('[data-remove-transaction-record-index]');
            if (!removeBtn) return;
            const index = Number(removeBtn.dataset.removeTransactionRecordIndex);
            if (!Number.isFinite(index)) return;
            removeTransactionFromActiveCard(index);
        });
        orderDetailsModal?.addEventListener('click', function(event) {
            if (event.target !== orderDetailsModal) return;
            closeOrderDetailsModal();
        });
        closeExpertNotifyModalBtn?.addEventListener('click', closeExpertNotifyModal);
        cancelExpertNotifyBtn?.addEventListener('click', closeExpertNotifyModal);
        expertNotifyModal?.addEventListener('click', function(event) {
            if (event.target !== expertNotifyModal) return;
            closeExpertNotifyModal();
        });
        monthlyDataModal?.addEventListener('click', function(event) {
            if (event.target !== monthlyDataModal) return;
            closeMonthlyDataModal();
        });
        expertNotifySearch?.addEventListener('input', renderExpertNotifyList);
        expertNotifySelectAll?.addEventListener('change', function() {
            const rows = getFilteredExpertsForNotify(expertNotifySearch?.value || '');
            rows.forEach((expert) => {
                if (this.checked) selectedNotifyExpertIds.add(expert.expertId);
                else selectedNotifyExpertIds.delete(expert.expertId);
            });
            renderExpertNotifyList();
        });
        expertNotifyList?.addEventListener('change', function(event) {
            const checkbox = event.target.closest('.expert-notify-checkbox');
            if (!checkbox) return;
            const expertId = String(checkbox.dataset.expertId || '').trim();
            if (!expertId) return;
            if (checkbox.checked) selectedNotifyExpertIds.add(expertId);
            else selectedNotifyExpertIds.delete(expertId);
            renderExpertNotifyList();
        });
        confirmExpertNotifyBtn?.addEventListener('click', function() {
            if (!activeNotifyOrderCard) return;
            const selectedIds = Array.from(selectedNotifyExpertIds);
            if (!selectedIds.length) {
                alert('Select at least one expert.');
                return;
            }
            const expertsToNotify = readExpertsFromStorage()
                .map((row) => normalizeExpertRecord(row))
                .filter((expert) => selectedIds.includes(expert.expertId));
            if (!expertsToNotify.length) {
                alert('No valid experts found for notification.');
                return;
            }
            const taskTitle = activeNotifyOrderCard.querySelector('.order-title')?.textContent?.trim() || 'Untitled Task';
            const serviceType = String(activeNotifyOrderCard.dataset.serviceType || '').trim() || taskTitle;
            const expertDeadline = activeNotifyOrderCard.querySelector('.order-expert-deadline')?.dataset.baseValue || activeNotifyOrderCard.querySelector('.order-expert-deadline')?.textContent?.trim() || 'TBD';
            const expertPayout = String(activeNotifyOrderCard.dataset.expertPayout || activeNotifyOrderCard.querySelector('.order-expert-pay')?.textContent || 'TBD').trim() || 'TBD';
            const taskFileLinks = [
                ...parseOrderAttachments(activeNotifyOrderCard.dataset.questionAttachments || '[]'),
                ...parseOrderAttachments(activeNotifyOrderCard.dataset.customOrderFiles || '[]')
                    .filter((item) => String(item.path || '').trim().toLowerCase().startsWith('expert/task'))
            ]
                .map((item) => String(item.url || '').trim())
                .filter(Boolean);

            const originalText = confirmExpertNotifyBtn.textContent;
            confirmExpertNotifyBtn.disabled = true;
            confirmExpertNotifyBtn.textContent = 'Sending...';

            whatsappFetch('/api/whatsapp/expert-notify', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    taskTitle,
                    serviceType,
                    expertDeadline,
                    expertPayout,
                    taskFileLinks,
                    experts: expertsToNotify.map((expert) => ({
                        expertId: expert.expertId,
                        name: expert.name,
                        waId: expert.waId
                    }))
                })
            })
                .then(async (res) => {
                    const data = await res.json();
                    if (!res.ok || !data?.ok) {
                        throw new Error(data?.error || 'Expert notification failed');
                    }
                    const successfulIds = Array.isArray(data.successes)
                        ? data.successes.map((row) => String(row?.expertId || '').trim()).filter(Boolean)
                        : selectedIds;
                    const mergedIds = Array.from(new Set([...getNotifiedExpertIds(activeNotifyOrderCard), ...successfulIds]));
                    setNotifiedExpertIds(activeNotifyOrderCard, mergedIds);
                    syncOrderNotifyButton(activeNotifyOrderCard);
                    persistOrderListToStorage();
                    closeExpertNotifyModal();
                    if (Array.isArray(data.failures) && data.failures.length) {
                        alert(`Notified ${successfulIds.length} expert(s). ${data.failures.length} failed.`);
                    }
                })
                .catch((error) => {
                    alert('Expert notification failed: ' + (error?.message || 'Unknown error'));
                })
                .finally(() => {
                    confirmExpertNotifyBtn.disabled = false;
                    confirmExpertNotifyBtn.textContent = originalText;
                });
        });
        ORDER_DETAILS_DRAFT_FIELD_IDS.forEach((id) => {
            const field = document.getElementById(id);
            if (!field) return;
            field.addEventListener('input', saveActiveOrderDraft);
            field.addEventListener('change', saveActiveOrderDraft);
        });
        // Attach File Button
        document.getElementById('attachFileBtn').onclick = function() {
            if (!chatFileInput) return;
            chatFileInput.click();
        };
        if (aiAssistBtn) {
            aiAssistBtn.onclick = async function() {
                if (!activeWhatsappWaId) {
                    alert('Select a contact first.');
                    return;
                }
                try {
                    await setContactAiMode(activeWhatsappWaId, 'ai');
                    alert('This chat has been transferred to AI.');
                } catch (err) {
                    alert('AI transfer failed: ' + (err?.message || 'Unknown error'));
                }
            };
        }
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

                const quoteBtn = event.target.closest('.chat-reply-quote[data-reply-target-id], .chat-reply-quote[data-reply-target-key]');
                if (quoteBtn) {
                    const targetId = quoteBtn.dataset.replyTargetId || '';
                    const targetKey = quoteBtn.dataset.replyTargetKey || '';
                    const target = targetId
                        ? chatMessages.querySelector('.chat-message-row[data-message-id="' + CSS.escape(targetId) + '"]')
                        : chatMessages.querySelector('.chat-message-row[data-message-key="' + CSS.escape(targetKey) + '"]');
                    if (target) {
                        target.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        target.classList.add('is-reply-target');
                        setTimeout(() => target.classList.remove('is-reply-target'), 1800);
                    }
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
        window.addEventListener('resize', positionOpenMessageMenus);
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
            if (!canCurrentAgentMessageContact(activeWhatsappWaId)) {
                const assigned = getAssignedAgentMetaByWaId(activeWhatsappWaId);
                alert(assigned.name ? ('This chat is assigned to ' + assigned.name + '.') : 'Assign this chat to yourself before sending messages.');
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
        document.addEventListener('click', function(event) {
            if (event.target.closest('#adminQuickMenuWrap')) {
                return;
            }
            adminQuickMenu?.classList.add('hidden');
            adminQuickMenuBtn?.setAttribute('aria-expanded', 'false');
        });
        ensureAgentRoster();
        setCurrentAgentName(agentName, true);
        // Logout
        document.getElementById('logoutBtn').onclick = function() {
            localStorage.removeItem('crmAuthSession');
            localStorage.removeItem('crmAuthToken');
            localStorage.removeItem('crmAuthEmail');
            localStorage.removeItem('agentName');
            localStorage.removeItem('agentRole');
            localStorage.removeItem(ACTIVE_AGENT_NAME_STORAGE_KEY);
            window.location.href = 'login.html';
        };

