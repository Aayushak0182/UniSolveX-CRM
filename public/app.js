lucide.createIcons();
        // Show agent name from localStorage
        const storedAgentName = localStorage.getItem('agentName');
        if (!storedAgentName) {
            window.location.href = 'login.html';
        }
        const agentName = storedAgentName || 'Aayush';
        document.getElementById('agentNameDisplay').textContent = 'Agent: ' + agentName;
        document.getElementById('agentInitialDisplay').textContent = agentName.charAt(0).toUpperCase();
        const whatsappStatusEl = document.getElementById('whatsappStatus');
        const whatsappLastEventEl = document.getElementById('whatsappLastEvent');
        const contactList = document.getElementById('contactList');
        const contactItems = Array.from(document.querySelectorAll('.contact-item'));
        const contactPanel = document.querySelector('.contact-panel');
        const contactFilter = document.getElementById('contactFilter');
        const contactSearch = document.getElementById('contactSearch');
        const contactToolbar = document.querySelector('.contact-toolbar');
        const toggleContactPanelWidthBtn = document.getElementById('toggleContactPanelWidth');
        const addContactBtn = document.getElementById('addContactBtn');
        const orderList = document.getElementById('orderList');
        const orderSearch = document.getElementById('orderSearch');
        const orderDeliveryWindow = document.getElementById('orderDeliveryWindow');
        const orderDetailsModal = document.getElementById('orderDetailsModal');
        const taskServiceType = document.getElementById('taskServiceType');
        const liveSessionCreateFields = document.getElementById('liveSessionCreateFields');
        const taskSessionStart = document.getElementById('taskSessionStart');
        const taskSessionDuration = document.getElementById('taskSessionDuration');
        const chatMessages = document.getElementById('chatMessages');
        const chatInput = document.getElementById('chatInput');
        const chatFileInput = document.getElementById('chatFileInput');
        const chatHeaderTitle = document.getElementById('chatHeaderTitle');
        const chatHeaderStatus = document.getElementById('chatHeaderStatus');
        let activeOrderCard = null;
        let activeWhatsappWaId = '';
        const orderSequenceByClient = {};
        const ORDER_LIST_STORAGE_KEY = 'unisolvex_order_list_html_v1';
        const MANUAL_CONTACTS_STORAGE_KEY = 'unisolvex_manual_contacts_v1';
        const WHATSAPP_CONTACT_ID_MAP_KEY = 'unisolvex_whatsapp_contact_id_map_v1';
        const WHATSAPP_CONTACT_ID_SEQUENCE_KEY = 'unisolvex_whatsapp_contact_id_seq_v1';
        const WHATSAPP_API_BASE_STORAGE_KEY = 'whatsappApiBase';
        const WHATSAPP_LAST_WORKING_API_BASE_KEY = 'unisolvex_last_working_whatsapp_api_base_v1';
        const CONTACT_PANEL_WIDE_STORAGE_KEY = 'unisolvex_contact_panel_wide_v1';
        const configuredWhatsappApiBase = localStorage.getItem(WHATSAPP_API_BASE_STORAGE_KEY);
        const hasHttpOrigin = /^https?:\/\//i.test(window.location.origin || '');
        const defaultWhatsappApiBase = (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1' || !hasHttpOrigin)
            ? 'http://localhost:3001'
            : window.location.origin;
        const WHATSAPP_API_FALLBACK_TIMEOUT_MS = 2500;
        const WHATSAPP_POLL_INTERVAL_MS = 5000;
        const WHATSAPP_DEBUG_POLL_INTERVAL_MS = 8000;
        const whatsappMessagesByContact = {};
        const whatsappMessageStatusById = {};
        let whatsappPollingTimer = null;
        let whatsappDebugTimer = null;
        let whatsappSocketConnected = false;
        let activeWhatsappApiBase = '';
        let reachableWhatsappApiBases = [];
        let contactOrderCounter = contactItems.length;

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

        function persistOrderListToStorage() {
            localStorage.setItem(ORDER_LIST_STORAGE_KEY, orderList.innerHTML);
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

                let matchesWindow = true;
                if (windowMs > 0) {
                    const deadlineRaw = card.querySelector('.order-deadline')?.textContent?.trim() || '';
                    const due = parseCardDateTime(deadlineRaw);
                    if (!due) {
                        matchesWindow = false;
                    } else {
                        const diffMs = due.getTime() - now.getTime();
                        matchesWindow = diffMs >= 0 && diffMs <= windowMs;
                    }
                }

                card.style.display = matchesQuery && matchesWindow ? '' : 'none';
            });
        }

        function updateContactStateUI(item) {
            const isPinned = item.dataset.pinned === 'true';
            const isUnread = item.dataset.unread === 'true';
            const pinBtn = item.querySelector('.pin-btn');
            const readToggleBtn = item.querySelector('.read-toggle-btn');
            const unreadDot = item.querySelector('.unread-dot');
            const contactName = item.querySelector('.contact-name');

            if (pinBtn) {
                pinBtn.textContent = isPinned ? 'Unpin' : 'Pin';
                pinBtn.classList.toggle('bg-yellow-100', isPinned);
                pinBtn.classList.toggle('text-yellow-800', isPinned);
            }
            if (readToggleBtn) {
                readToggleBtn.textContent = isUnread ? 'Mark Read' : 'Mark Unread';
            }
            if (unreadDot) {
                unreadDot.classList.toggle('hidden', !isUnread);
            }
            if (contactName) {
                contactName.classList.toggle('font-bold', isUnread);
            }
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

        function refreshContactOrder() {
            const sorted = [...contactItems].sort((a, b) => {
                const pinnedDiff = Number(b.dataset.pinned === 'true') - Number(a.dataset.pinned === 'true');
                if (pinnedDiff !== 0) return pinnedDiff;

                const activeDiff = Number(b.dataset.active === 'true') - Number(a.dataset.active === 'true');
                if (activeDiff !== 0) return activeDiff;

                const unreadDiff = Number(b.dataset.unread === 'true') - Number(a.dataset.unread === 'true');
                if (unreadDiff !== 0) return unreadDiff;

                return Number(a.dataset.order) - Number(b.dataset.order);
            });
            sorted.forEach((item) => contactList.appendChild(item));
        }

        function applyContactFilter() {
            const filterValue = contactFilter.value;
            const searchValue = (contactSearch.value || '').trim().toLowerCase();
            contactItems.forEach((item) => {
                const itemTag = item.dataset.tag || '';
                const nameText = (item.querySelector('.contact-name')?.textContent || '').toLowerCase();
                const contactIdText = (item.dataset.contactId || '').trim().toLowerCase();
                const idText = (item.querySelector('.contact-meta')?.textContent || item.querySelector('p')?.textContent || '').toLowerCase();
                const matchesTag = !filterValue || itemTag === filterValue;
                const matchesSearch = !searchValue || nameText.includes(searchValue) || contactIdText.includes(searchValue) || idText.includes(searchValue);
                item.style.display = matchesTag && matchesSearch ? '' : 'none';
            });
        }

        function normalizeWaId(rawValue) {
            return String(rawValue || '').replace(/[^\d]/g, '');
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

        function writeManualContactsToStorage(rows) {
            localStorage.setItem(MANUAL_CONTACTS_STORAGE_KEY, JSON.stringify(rows));
        }

        function saveManualContact(contact) {
            if (!contact?.waId) return;
            const rows = readManualContactsFromStorage();
            const existingIndex = rows.findIndex((row) => row.waId === contact.waId);
            const previous = existingIndex >= 0 ? rows[existingIndex] : null;
            const payload = {
                waId: contact.waId,
                profileName: getPreferredProfileName(previous?.profileName, contact.profileName, contact.waId)
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
                    profileName: contact.profileName
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

        function writeWhatsappContactIdMapToStorage(map) {
            localStorage.setItem(WHATSAPP_CONTACT_ID_MAP_KEY, JSON.stringify(map));
        }

        function getNextWhatsappContactIdValue(map) {
            const stored = Number(localStorage.getItem(WHATSAPP_CONTACT_ID_SEQUENCE_KEY) || '100100');
            const mapIds = Object.values(map).map((value) => Number(value)).filter((value) => Number.isFinite(value));
            const maxMapId = mapIds.length ? Math.max(...mapIds) : 100100;
            const next = Math.max(stored, maxMapId) + 1;
            localStorage.setItem(WHATSAPP_CONTACT_ID_SEQUENCE_KEY, String(next));
            return next;
        }

        function migrateWhatsappContactIdsToSixDigitSeries() {
            const map = readWhatsappContactIdMapFromStorage();
            const entries = Object.entries(map);
            if (!entries.length) {
                localStorage.setItem(WHATSAPP_CONTACT_ID_SEQUENCE_KEY, String(100100));
                return;
            }

            const shouldMigrate = entries.some((entry) => Number(entry[1]) < 100101);
            if (!shouldMigrate) {
                const maxExisting = Math.max(...entries.map((entry) => Number(entry[1])).filter((value) => Number.isFinite(value)));
                localStorage.setItem(WHATSAPP_CONTACT_ID_SEQUENCE_KEY, String(maxExisting));
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
            localStorage.setItem(WHATSAPP_CONTACT_ID_SEQUENCE_KEY, String(nextId - 1));
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
            statusEl.className = 'order-status text-[10px] font-bold uppercase px-2 py-0.5 rounded';
            if (normalized === 'completed') {
                statusEl.classList.add('text-green-600', 'bg-green-50');
                return;
            }
            if (normalized === 'new') {
                statusEl.classList.add('text-blue-600', 'bg-blue-50');
                return;
            }
            if (normalized === 'on hold') {
                statusEl.classList.add('text-amber-700', 'bg-amber-50');
                return;
            }
            if (normalized === 'cancelled') {
                statusEl.classList.add('text-rose-700', 'bg-rose-50');
                return;
            }
            if (normalized === 'refund') {
                statusEl.classList.add('text-violet-700', 'bg-violet-50');
                return;
            }
            if (normalized === 'archive') {
                statusEl.classList.add('text-gray-700', 'bg-gray-200');
                return;
            }
            statusEl.classList.add('text-orange-600', 'bg-orange-50');
        }

        function applyCardTypeStyle(card, serviceTypeValue) {
            const value = (serviceTypeValue || '').toLowerCase();
            card.classList.remove('bg-white', 'bg-orange-50', 'bg-amber-50', 'bg-blue-50', 'bg-red-50', 'bg-gray-100', 'border-orange-200', 'border-amber-200', 'border-blue-200', 'border-red-200', 'border-gray-300');

            if (value.includes('live session')) {
                card.classList.add('bg-amber-50', 'border-amber-200');
                return;
            }
            if (value.includes('assignment')) {
                card.classList.add('bg-blue-50', 'border-blue-200');
                return;
            }
            if (value.includes('project')) {
                card.classList.add('bg-red-50', 'border-red-200');
                return;
            }
            if (value.includes('full course')) {
                card.classList.add('bg-gray-100', 'border-gray-300');
                return;
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
            const titleEl = card.querySelector('.order-title');
            if (!titleEl) return;

            let titleRow = card.querySelector('.order-title-row');
            if (!titleRow) {
                titleRow = document.createElement('div');
                titleRow.className = 'order-title-row flex items-center gap-2';
                titleEl.parentNode.insertBefore(titleRow, titleEl);
                titleRow.appendChild(titleEl);
            }

            let tagEl = titleRow.querySelector('.order-service-tag');
            if (!tagEl) {
                tagEl = document.createElement('span');
                tagEl.className = 'order-service-tag text-[10px] font-semibold px-2 py-0.5 rounded-full bg-indigo-100 text-indigo-700';
                titleRow.appendChild(tagEl);
            }

            const serviceTypeValue = (card.dataset.serviceType || titleEl.textContent || '').trim();
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

        function appendFileMessage(file) {
            if (!chatMessages || !file) return;
            const wrapper = document.createElement('div');
            wrapper.className = 'flex flex-col items-end';
            const now = new Date();
            const hh = now.getHours() % 12 || 12;
            const mm = String(now.getMinutes()).padStart(2, '0');
            const ampm = now.getHours() >= 12 ? 'PM' : 'AM';
            const timestamp = hh + ':' + mm + ' ' + ampm;

            wrapper.innerHTML = `
                <div class="chat-bubble-admin p-3 max-w-md shadow-sm text-sm text-gray-800">
                    <p class="font-semibold">Attachment</p>
                    <p class="mt-1">${file.name}</p>
                    <p class="text-[11px] text-gray-500 mt-1">${formatFileSize(file.size)}</p>
                </div>
                <span class="text-[10px] text-gray-400 mt-1 mr-1">${timestamp}</span>
            `;
            chatMessages.appendChild(wrapper);
            chatMessages.scrollTop = chatMessages.scrollHeight;
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
            if (!item.dataset.tag) item.dataset.tag = '';
            if (!item.dataset.active) item.dataset.active = 'false';
            updateContactStateUI(item);

            const tagSelect = item.querySelector('.contact-tag-select');
            if (tagSelect) {
                tagSelect.addEventListener('click', (e) => e.stopPropagation());
                tagSelect.addEventListener('change', function() {
                    item.dataset.tag = this.value;
                    applyContactFilter();
                });
            }

            const pinBtn = item.querySelector('.pin-btn');
            if (pinBtn) {
                pinBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    item.dataset.pinned = item.dataset.pinned === 'true' ? 'false' : 'true';
                    updateContactStateUI(item);
                    refreshContactOrder();
                    applyContactFilter();
                });
            }

            const readToggleBtn = item.querySelector('.read-toggle-btn');
            if (readToggleBtn) {
                readToggleBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    item.dataset.unread = item.dataset.unread === 'true' ? 'false' : 'true';
                    updateContactStateUI(item);
                    refreshContactOrder();
                    applyContactFilter();
                });
            }

            item.addEventListener('click', () => {
                const waId = normalizeWaId(item.dataset.waId || '');
                if (!waId) return;
                const name = item.querySelector('.contact-name')?.textContent?.trim() || waId;
                activeWhatsappWaId = waId;
                setActiveContactItem(item);
                item.dataset.unread = 'false';
                updateContactStateUI(item);
                refreshContactOrder();
                applyContactFilter();
                if (chatHeaderTitle) chatHeaderTitle.textContent = 'Chat with ' + name;
                if (chatHeaderStatus) chatHeaderStatus.innerHTML = '<span class="w-2 h-2 bg-green-500 rounded-full"></span> WhatsApp Connected';
                renderWhatsappMessages(waId);
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

        function normalizeMessageStatus(rawStatus) {
            const status = String(rawStatus || '').toLowerCase();
            if (status === 'read') return 'read';
            if (status === 'delivered') return 'delivered';
            if (status === 'failed') return 'failed';
            return 'sent';
        }

        function getStatusTickHtml(rawStatus) {
            const status = normalizeMessageStatus(rawStatus);
            if (status === 'read') {
                return '<span class="ml-1 text-[11px] text-blue-500 font-semibold">&#10003;&#10003;</span>';
            }
            if (status === 'delivered') {
                return '<span class="ml-1 text-[11px] text-gray-400 font-semibold">&#10003;&#10003;</span>';
            }
            if (status === 'failed') {
                return '<span class="ml-1 text-[11px] text-red-500 font-semibold">!</span>';
            }
            return '<span class="ml-1 text-[11px] text-gray-400 font-semibold">&#10003;</span>';
        }

        function renderWhatsappMessages(waId) {
            if (!chatMessages || !waId) return;
            const messages = whatsappMessagesByContact[waId] || [];
            chatMessages.innerHTML = '';
            if (!messages.length) {
                chatMessages.innerHTML = '<p class="text-sm text-gray-500">No WhatsApp messages yet.</p>';
                return;
            }
            messages.forEach((msg) => {
                const wrap = document.createElement('div');
                const incoming = msg.direction === 'incoming';
                const statusTick = incoming ? '' : getStatusTickHtml(msg.status);
                wrap.className = 'flex flex-col ' + (incoming ? 'items-start' : 'items-end');
                wrap.innerHTML = `
                    <div class="${incoming ? 'chat-bubble-client' : 'chat-bubble-admin'} p-3 max-w-md shadow-sm text-sm text-gray-800">
                        ${escapeHtml(msg.text || '[Unsupported message type]')}
                    </div>
                    <span class="text-[10px] text-gray-400 mt-1 ${incoming ? 'ml-1' : 'mr-1'}">${formatChatTime(msg.timestamp)}${statusTick}</span>
                `;
                chatMessages.appendChild(wrap);
            });
            chatMessages.scrollTop = chatMessages.scrollHeight;
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
            const existingName = existing?.querySelector('.contact-name')?.textContent?.trim() || '';
            const profileName = getPreferredProfileName(existingName, contact.profileName, normalizedWaId);
            const contactId = getOrCreateWhatsappContactId(normalizedWaId);
            saveManualContact({ waId: normalizedWaId, profileName });
            if (!existing) {
                const item = document.createElement('div');
                item.className = 'contact-item p-3 border-b hover:bg-gray-50 cursor-pointer';
                item.dataset.waId = normalizedWaId;
                item.dataset.contactId = contactId;
                item.dataset.source = 'whatsapp';
                item.innerHTML = `
                    <div class="flex justify-between items-start gap-2">
                        <span class="contact-name font-semibold text-gray-800 text-sm">${escapeHtml(profileName || normalizedWaId)}</span>
                        <select class="contact-tag-select text-xs border rounded px-1.5 py-1 bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500">
                            <option value="">Tag</option>
                            <option value="new">New Client</option>
                            <option value="old">Old Client</option>
                            <option value="tutor">Tutor</option>
                            <option value="friend">Client's Friend</option>
                            <option value="useless">Useless/Broker</option>
                        </select>
                    </div>
                    <p class="contact-meta text-xs text-gray-500">ID: ${escapeHtml(contactId)}</p>
                    <div class="mt-2 flex items-center gap-2">
                        <button class="pin-btn text-[11px] px-2 py-1 rounded border bg-white hover:bg-gray-50">Pin</button>
                        <button class="read-toggle-btn text-[11px] px-2 py-1 rounded border bg-white hover:bg-gray-50">Mark Unread</button>
                        <span class="unread-dot hidden text-[10px] text-red-600 font-semibold">Unread</span>
                    </div>
                `;
                contactItems.push(item);
                wireContactItem(item);
                contactList.prepend(item);
                existing = item;
            } else {
                const nameEl = existing.querySelector('.contact-name');
                if (nameEl) nameEl.textContent = profileName;
                existing.dataset.contactId = contactId;
                const metaEl = existing.querySelector('.contact-meta') || existing.querySelector('p');
                if (metaEl) {
                    metaEl.classList.add('contact-meta');
                    metaEl.textContent = 'ID: ' + contactId;
                }
            }
            if (activeWhatsappWaId && activeWhatsappWaId === normalizedWaId) {
                setActiveContactItem(existing);
            }
            if (markUnread) {
                existing.dataset.unread = 'true';
                updateContactStateUI(existing);
                refreshContactOrder();
                applyContactFilter();
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
            whatsappMessagesByContact[waId].push({
                id: payload.id || '',
                text: payload.text || '',
                timestamp: payload.timestamp || new Date().toISOString(),
                direction: payload.direction || 'incoming',
                status: normalizeMessageStatus(payload.status || whatsappMessageStatusById[payload.id || '']?.status),
                statusTimestamp: payload.statusTimestamp || whatsappMessageStatusById[payload.id || '']?.statusTimestamp || payload.timestamp || new Date().toISOString()
            });
            upsertWhatsappContact({
                waId: waId,
                profileName: payload.profileName || waId
            }, payload.direction === 'incoming' && activeWhatsappWaId !== waId);

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
                    upsertWhatsappContact({
                        waId,
                        profileName: entry.profileName || waId
                    }, false);
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
                });
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
            initWhatsappRealtime();
            initWhatsappPolling();
            initWhatsappDebugPolling();
        }

        contactItems.forEach((item, index) => {
            item.dataset.order = String(index);
            wireContactItem(item);
        });

        migrateWhatsappContactIdsToSixDigitSeries();
        setActiveContactItem(null);
        restoreManualContactsFromStorage();
        initWhatsappSync();

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

        contactFilter.addEventListener('change', applyContactFilter);
        contactSearch.addEventListener('input', applyContactFilter);
        applyContactFilter();
        restoreOrderListFromStorage();
        Array.from(orderList.querySelectorAll('.order-card')).forEach((card) => {
            syncOrderServiceTag(card);
            syncOrderPaymentIndicator(card);
            applyCardTypeStyle(card, card.dataset.serviceType || card.querySelector('.order-title')?.textContent || '');
            const cardClientId = (card.dataset.clientId || '').replace(/\D/g, '');
            const cardOrderId = (card.querySelector('.order-id')?.textContent || '').replace(/[^\d]/g, '');
            if (!cardClientId || !cardOrderId || !cardOrderId.startsWith(cardClientId)) return;
            const suffix = Number(cardOrderId.slice(cardClientId.length));
            if (!Number.isFinite(suffix)) return;
            orderSequenceByClient[cardClientId] = Math.max(orderSequenceByClient[cardClientId] || 1009, suffix);
        });
        if (!localStorage.getItem(ORDER_LIST_STORAGE_KEY)) {
            persistOrderListToStorage();
        }
        orderSearch.addEventListener('input', applyOrderSearchFilter);
        if (orderDeliveryWindow) {
            orderDeliveryWindow.addEventListener('change', applyOrderSearchFilter);
        }
        applyOrderSearchFilter();
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
            const clientText = document.getElementById('taskClientInput').value.trim();
            const serviceType = document.getElementById('taskServiceType').value;
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
                    <span class="order-status text-[10px] font-bold uppercase px-2 py-0.5 rounded">New</span>
                    <div class="flex items-center gap-2">
                        <span class="order-id text-[10px] text-gray-400 font-mono">#${orderId}</span>
                        <button class="open-order-details-btn p-1 rounded bg-indigo-50 text-indigo-700 hover:bg-indigo-100" title="Open Details">
                            <i data-lucide="square-arrow-out-up-right" class="w-3.5 h-3.5"></i>
                        </button>
                    </div>
                </div>
                <div class="order-title-row flex items-center gap-2">
                    <h4 class="order-title text-sm font-bold text-gray-800">${serviceType}</h4>
                    <span class="order-service-tag text-[10px] font-semibold px-2 py-0.5 rounded-full bg-indigo-100 text-indigo-700">${serviceType}</span>
                </div>
                <p class="order-client-note text-[11px] text-gray-500 mt-0.5">${clientText.replace(/\(\d+\)/, '').trim()}</p>
                <div class="mt-1 grid grid-cols-2 gap-x-3 gap-y-0.5 text-[11px] text-gray-600">
                    <p>C.A - <span class="order-amount font-semibold text-indigo-600">TBD</span><span class="order-payment-indicator hidden ml-1"></span></p>
                    <p>E.D - <span class="order-expert-deadline text-amber-700 font-medium"></span></p>
                    <p>C.D - <span class="order-deadline text-red-500 font-medium"></span></p>
                    <p>By - <span class="order-created-by text-gray-700 font-medium">${agentName}</span></p>
                </div>
                <p class="order-session mt-1 text-[11px] text-gray-600 ${isLiveSession ? '' : 'hidden'}">S.T - ${sessionStart || 'TBD'} | Dur - ${sessionDuration || 'TBD'}m</p>
                <p class="order-comment hidden mt-1 text-[11px] text-gray-700 border-t pt-1">Note: </p>
            `;
            syncOrderServiceTag(card);
            syncOrderPaymentIndicator(card);
            applyCardTypeStyle(card, serviceType);
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
            const deadline = card.querySelector('.order-deadline')?.textContent?.trim() || '';

            document.getElementById('odTitle').value = title;
            document.getElementById('odServiceType').value = card.dataset.serviceType || title;
            document.getElementById('odStatus').value = status;
            document.getElementById('odOrderId').value = cardId;
            document.getElementById('odClientId').value = clientIdText;
            document.getElementById('odAssignedTo').value = card.dataset.assignedTo || '';
            document.getElementById('odCreatedBy').value = card.dataset.createdBy || card.dataset.assignedTo || '';
            document.getElementById('odLabels').value = card.dataset.labels || '';
            document.getElementById('odActualDeadline').value = toDateTimeLocalValue(deadline);
            document.getElementById('odExpertDeadline').value = toDateTimeLocalValue(card.dataset.expertDeadline || '');
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
                const selectedFiles = Array.from(this.files || []);
                if (!selectedFiles.length) return;
                selectedFiles.forEach((file) => appendFileMessage(file));
                this.value = '';
            });
        }
        // Notification Button
        document.getElementById('notifyBtn').onclick = function() {
            alert('Notification sent!');
        };
        // Send Button
        document.getElementById('sendBtn').onclick = async function() {
            const text = (chatInput?.value || '').trim();
            if (!text) return;
            if (!activeWhatsappWaId) {
                alert('Select a WhatsApp contact first.');
                return;
            }

            try {
                const res = await whatsappFetch('/api/whatsapp/send', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        waId: normalizeWaId(activeWhatsappWaId),
                        text: text
                    })
                });
                const data = await res.json();
                if (!res.ok || !data?.ok) {
                    throw new Error(data?.error || 'Send failed');
                }
                if (!Array.isArray(whatsappMessagesByContact[activeWhatsappWaId])) {
                    whatsappMessagesByContact[activeWhatsappWaId] = [];
                }
                whatsappMessagesByContact[activeWhatsappWaId].push({
                    id: data?.id || '',
                    text,
                    timestamp: new Date().toISOString(),
                    direction: 'outgoing',
                    status: 'sent',
                    statusTimestamp: new Date().toISOString()
                });
                if (data?.id) {
                    whatsappMessageStatusById[data.id] = {
                        status: 'sent',
                        statusTimestamp: new Date().toISOString()
                    };
                }
                renderWhatsappMessages(activeWhatsappWaId);
                chatInput.value = '';
                await loadWhatsappContacts();
            } catch (err) {
                console.error('Failed to send WhatsApp message:', err);
                const reason = err?.message || 'Unknown send error';
                alert('Message send failed: ' + reason);
            }
        };
        if (chatInput) {
            chatInput.addEventListener('keydown', function(event) {
                if (event.key !== 'Enter' || event.shiftKey) return;
                event.preventDefault();
                document.getElementById('sendBtn').click();
            });
        }
        // Logout
        document.getElementById('logoutBtn').onclick = function() {
            localStorage.removeItem('agentName');
            window.location.href = 'login.html';
        };

