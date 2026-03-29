const APP_VERSION = '12.0.9';

let __loginErrorDismissTimer = null;

/** Plaintext passwords for User rows not yet confirmed by a successful workspace sync (never in localStorage). */
const __pendingPasswordsByUserId = new Map();

const API_AUTH_TOKEN_KEY = 'tasktrack_api_auth_token';

function setApiAuthToken(token) {
    if (token && typeof token === 'string') {
        localStorage.setItem(API_AUTH_TOKEN_KEY, token);
    }
}

function clearApiAuthToken() {
    localStorage.removeItem(API_AUTH_TOKEN_KEY);
}

function rememberPendingUserPasswordForSync(userId, plainPassword) {
    if (!isApiMode() || plainPassword == null || String(plainPassword).length === 0) return;
    const id = typeof userId === 'number' ? userId : parseInt(String(userId), 10);
    if (Number.isNaN(id)) return;
    __pendingPasswordsByUserId.set(id, String(plainPassword));
}

function clearPendingPasswordsForSync() {
    __pendingPasswordsByUserId.clear();
}

function buildWorkspacePayloadForPut() {
    const payload = JSON.parse(JSON.stringify(__workspaceCache));
    if (payload.users && Array.isArray(payload.users)) {
        for (const u of payload.users) {
            const id = typeof u.id === 'number' ? u.id : parseInt(String(u.id), 10);
            if (Number.isNaN(id)) continue;
            const p = __pendingPasswordsByUserId.get(id);
            if (p) u.password = p;
        }
    }
    return payload;
}

function applyWorkspacePutSuccess(normalized) {
    if (normalized && Array.isArray(normalized.users)) {
        for (const u of normalized.users) {
            const id = typeof u.id === 'number' ? u.id : parseInt(String(u.id), 10);
            if (!Number.isNaN(id)) __pendingPasswordsByUserId.delete(id);
        }
    }
    __workspaceCache = normalized;
}

// Data Storage
let currentUser = null;
let currentMonth = new Date().getMonth();
let currentYear = new Date().getFullYear();
let drilldownContext = null; // Stores: { type, title, count, filterFunction, monthValue }

// --- API / cloud sync (when window.API_BASE_URL is set in config.js) ---
let __workspaceCache = null;
let __workspacePushTimer = null;
let __lastWorkspacePutError = '';
const WORKSPACE_PUSH_DEBOUNCE_MS = 400;

function isApiMode() {
    return typeof window.API_BASE_URL === 'string' && window.API_BASE_URL.trim().length > 0;
}

function apiBase() {
    return String(window.API_BASE_URL || '').replace(/\/$/, '');
}

async function apiFetch(path, options = {}) {
    const url = apiBase() + path;
    const headers = { ...(options.headers || {}) };
    if (options.body && typeof options.body === 'string' && !headers['Content-Type']) {
        headers['Content-Type'] = 'application/json';
    }
    const bearer = localStorage.getItem(API_AUTH_TOKEN_KEY);
    if (bearer && !headers.Authorization) {
        headers.Authorization = `Bearer ${bearer}`;
    }
    return fetch(url, { ...options, credentials: 'include', headers });
}

async function apiPullWorkspace() {
    const res = await apiFetch('/api/workspace');
    if (!res.ok) throw new Error('Failed to load workspace');
    const body = await res.json();
    __workspaceCache = normalizeData(body);
}

function scheduleWorkspacePush() {
    if (!isApiMode() || !currentUser || currentUser.isMaster) return;
    clearTimeout(__workspacePushTimer);
    __workspacePushTimer = setTimeout(() => flushWorkspaceToApi(), WORKSPACE_PUSH_DEBOUNCE_MS);
}

async function putWorkspaceCacheToServer() {
    if (!isApiMode() || !__workspaceCache) return false;
    __lastWorkspacePutError = '';
    try {
        const payload = buildWorkspacePayloadForPut();
        const res = await apiFetch('/api/workspace', { method: 'PUT', body: JSON.stringify(payload) });
        if (!res.ok) {
            let detail = `Server returned ${res.status}`;
            try {
                const t = await res.text();
                if (t) {
                    try {
                        const j = JSON.parse(t);
                        if (j.error) detail = j.error;
                    } catch {
                        detail = t.slice(0, 240);
                    }
                }
            } catch (_) { /* ignore */ }
            __lastWorkspacePutError = detail;
            console.error('Workspace sync failed', detail);
            return false;
        }
        const body = await res.json();
        applyWorkspacePutSuccess(normalizeData(body));
        return true;
    } catch (e) {
        __lastWorkspacePutError = e && e.message ? String(e.message) : 'Network or client error';
        console.error('Workspace sync error', e);
        return false;
    }
}

async function flushWorkspaceToApi() {
    if (!isApiMode() || !currentUser || currentUser.isMaster || !__workspaceCache) return;
    await putWorkspaceCacheToServer();
}

/** Immediate PUT (used after import) while session cookie is still valid. Cancels debounced push. */
async function flushWorkspaceToApiNow() {
    if (!isApiMode() || !currentUser || currentUser.isMaster || !__workspaceCache) return false;
    clearTimeout(__workspacePushTimer);
    __workspacePushTimer = null;
    return putWorkspaceCacheToServer();
}

document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden' && isApiMode() && currentUser && !currentUser.isMaster) {
        flushWorkspaceToApi();
    }
});

window.addEventListener('pagehide', () => {
    if (isApiMode() && currentUser && !currentUser.isMaster && __workspaceCache) {
        void flushWorkspaceToApiNow();
    }
});

function defaultWorkspaceShell() {
    return {
        users: [],
        tasks: [],
        locations: [
            { id: 1, name: 'Mundra' },
            { id: 2, name: 'JNPT' },
            { id: 3, name: 'Combine' }
        ],
        segregationTypes: [
            { id: 1, name: 'PSA Reports' },
            { id: 2, name: 'Internal Reports' }
        ],
        holidays: [],
        notes: [],
        learningNotes: [],
        milestones: [],
        dailyPlanner: [],
        locationItems: [],
        codeSnippets: [],
        journal: {}
    };
}

async function uploadRemoteAttachment(locationId, attachmentId, blob, filename, contentType) {
    const fd = new FormData();
    fd.append('locationId', String(locationId));
    fd.append('attachmentId', String(attachmentId));
    fd.append('file', blob, filename || 'file');
    const res = await fetch(`${apiBase()}/api/attachments`, {
        method: 'POST',
        credentials: 'include',
        body: fd
    });
    if (!res.ok) {
        const errText = await res.text().catch(() => '');
        throw new Error(errText || `Attachment upload failed (${res.status})`);
    }
}

function dataUrlToBlob(dataUrl) {
    const parts = dataUrl.split(',');
    const meta = parts[0];
    const b64 = parts[1];
    if (!b64) return null;
    const mime = (meta.match(/data:([^;]+)/) || [])[1] || 'application/octet-stream';
    const binary = atob(b64);
    const len = binary.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
    return new Blob([bytes], { type: mime });
}

// Global error reporter to surface silent failures (shown as alert and console)
function reportError(err, context = 'Unexpected error') {
    const msg = `${context}: ${err && err.message ? err.message : err}`;
    console.error(msg, err);
    try {
        alert(msg);
    } catch (_) {
        // ignore if alert unavailable
    }
}

// Attach global listeners once
if (!window.__taskAppErrorHooksInstalled) {
    window.__taskAppErrorHooksInstalled = true;
    window.addEventListener('error', (e) => {
        // Ignore generic "Script error." which comes from cross-origin scripts (browser extensions, etc.)
        // These errors can't be properly captured due to CORS policies
        const errorMsg = e.error?.message || e.message || '';
        if (errorMsg === 'Script error.' || errorMsg === 'Script error') {
            console.warn('Ignored cross-origin script error (likely from browser extension)');
            return;
        }
        // Only report errors that we can actually handle and that are from our code
        if (errorMsg && errorMsg.trim() !== '') {
            reportError(e.error || errorMsg, 'Runtime error');
        }
    });
    window.addEventListener('unhandledrejection', (e) => {
        reportError(e.reason || e.message || 'Unhandled promise rejection', 'Unhandled promise');
    });
}

// Utility function to format dates as YYYY-MM-DD (avoids timezone issues)
function formatDateString(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

// Utility function to format dates as DD-MM-YYYY for display
function formatDateDisplay(date) {
    if (!date) return '';
    let dateObj;
    if (typeof date === 'string') {
        // Check if it's an ISO date string (contains T or has time component)
        if (date.includes('T') || date.includes(' ')) {
            dateObj = new Date(date);
        } else {
            // Parse YYYY-MM-DD format
            const parts = date.split('-');
            if (parts.length === 3) {
                dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
            } else {
                return date; // Return as-is if not in expected format
            }
        }
    } else if (date instanceof Date) {
        dateObj = date;
    } else {
        return '';
    }

    const day = String(dateObj.getDate()).padStart(2, '0');
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const year = dateObj.getFullYear();
    return `${day}-${month}-${year}`;
}

// Utility: parse a date string in either YYYY-MM-DD or DD-MM-YYYY, return Date or null
function parseDateFlexible(dateStr) {
    if (!dateStr || typeof dateStr !== 'string') return null;
    const parts = dateStr.split('-').map(p => parseInt(p, 10));
    if (parts.length !== 3 || parts.some(isNaN)) return null;

    // Detect format: if first part is 4 digits, assume YYYY-MM-DD, otherwise DD-MM-YYYY
    let y, m, d;
    if (String(parts[0]).length === 4) {
        [y, m, d] = [parts[0], parts[1], parts[2]];
    } else {
        [d, m, y] = [parts[0], parts[1], parts[2]];
    }
    const dt = new Date(y, m - 1, d);
    if (isNaN(dt.getTime())) return null;
    dt.setHours(0, 0, 0, 0);
    return dt;
}

// Debounce utility for search/filter inputs (reduces re-renders on rapid typing)
function debounce(fn, delay = 300) {
    let timeoutId;
    return function (...args) {
        clearTimeout(timeoutId);
        timeoutId = setTimeout(() => fn.apply(this, args), delay);
    };
}

// IndexedDB for location attachments (avoids localStorage 5MB limit and allows more/larger files)
const ATTACHMENTS_DB_NAME = 'TodoAppAttachmentsDB';
const ATTACHMENTS_STORE = 'attachments';
const MAX_ATTACHMENT_SIZE = 50 * 1024 * 1024; // 50MB per file when using IndexedDB
const MAX_ATTACHMENT_FILES = 50; // Max files per location

function openAttachmentsDB() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(ATTACHMENTS_DB_NAME, 1);
        req.onerror = () => reject(req.error);
        req.onsuccess = () => resolve(req.result);
        req.onupgradeneeded = (e) => {
            const db = e.target.result;
            if (!db.objectStoreNames.contains(ATTACHMENTS_STORE)) {
                db.createObjectStore(ATTACHMENTS_STORE, { keyPath: 'key' });
            }
        };
    });
}

function getAttachmentBlob(locationId, attachmentId) {
    const key = `${locationId}_${attachmentId}`;
    return openAttachmentsDB().then(db => {
        return new Promise((resolve, reject) => {
            const tx = db.transaction(ATTACHMENTS_STORE, 'readonly');
            const req = tx.objectStore(ATTACHMENTS_STORE).get(key);
            req.onsuccess = () => resolve(req.result ? req.result.data : null);
            req.onerror = () => reject(req.error);
        });
    });
}

function putAttachmentBlob(locationId, attachmentId, dataBase64) {
    const key = `${locationId}_${attachmentId}`;
    return openAttachmentsDB().then(db => {
        return new Promise((resolve, reject) => {
            const tx = db.transaction(ATTACHMENTS_STORE, 'readwrite');
            const req = tx.objectStore(ATTACHMENTS_STORE).put({ key, data: dataBase64 });
            req.onsuccess = () => resolve();
            req.onerror = () => reject(req.error);
        });
    });
}

function removeAttachmentBlob(locationId, attachmentId) {
    const key = `${locationId}_${attachmentId}`;
    return openAttachmentsDB().then(db => {
        return new Promise((resolve, reject) => {
            const tx = db.transaction(ATTACHMENTS_STORE, 'readwrite');
            const req = tx.objectStore(ATTACHMENTS_STORE).delete(key);
            req.onsuccess = () => resolve();
            req.onerror = () => reject(req.error);
        });
    });
}

// Debounced search handlers (called from onkeyup)
const debouncedFilterTasks = debounce(filterTasks, 300);
const debouncedRenderNotes = debounce(renderNotes, 300);
const debouncedRenderLearningNotes = debounce(renderLearningNotes, 300);
const debouncedRenderLocations = debounce(renderLocations, 300);
const debouncedRenderCodeSnippets = debounce(renderCodeSnippets, 300);
const debouncedJournalSearch = debounce(journalDoSearch, 300);
const debouncedJournalSave = debounce(journalSaveCurrent, 1500);

// Escape HTML to prevent XSS (used across app)
function escapeHtml(text) {
    if (text === null || text === undefined) return '';
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return String(text).replace(/[&<>"']/g, m => map[m]);
}

const MASTER_ACCOUNT_EMAIL = (typeof window.MASTER_ACCOUNT_EMAIL === 'string' && window.MASTER_ACCOUNT_EMAIL.trim())
    ? window.MASTER_ACCOUNT_EMAIL.trim().toLowerCase()
    : 'mansuri.amin1@gmail.com';

function isMasterUserRecord(u) {
    if (!u) return false;
    if (u.isMaster === true) return true;
    const em = (u.email || '').toLowerCase();
    return em === MASTER_ACCOUNT_EMAIL;
}

/** Master Settings: filter dropdown for User Management + account status table. */
function getMasterAccountFilterMode() {
    const wrap = document.getElementById('masterUserMgmtFilterWrap');
    const filt = document.getElementById('masterUserStatusFilter');
    if (!filt || (wrap && wrap.classList.contains('hidden'))) return 'all';
    const v = filt.value;
    return v === 'active' || v === 'disabled' ? v : 'all';
}

function userPassesMasterAccountFilter(u, mode) {
    if (!u) return false;
    if (u.isMaster === true) return true;
    if (mode === 'active') return u.is_active !== false;
    if (mode === 'disabled') return u.is_active === false;
    return true;
}

/** Users shown in assignee / filter dropdowns (hides master for non-admin users). */
function usersVisibleInPickers() {
    const data = getData();
    const list = (data.users || []).filter(u => u.is_active !== false);
    if (!currentUser) return list;
    if (currentUser.role === 'admin' || currentUser.isMaster) return list;
    return list.filter(u => !isMasterUserRecord(u));
}

/**
 * Users an admin may assign tasks to. Tenant admins use workspace-scoped users only.
 * Master account sees all users in GET /workspace — restrict assignment to self only.
 * Non-admin users may only pick themselves.
 */
function taskAssigneePickerUsers() {
    if (!currentUser) return [];
    if (currentUser.isMaster) {
        const data = getData();
        return (data.users || []).filter(u => u.is_active !== false && (u.isMaster === true || u.id === currentUser.id));
    }
    if (currentUser.role !== 'admin') {
        const data = getData();
        const self = (data.users || []).find(u => Number(u.id) === Number(currentUser.id));
        return self && self.is_active !== false ? [self] : [];
    }
    return usersVisibleInPickers();
}

// Utility function to escape CSV values
function escapeCSV(value) {
    if (value === null || value === undefined) return '';
    const str = String(value);
    if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
}

// Dashboard period persistence
function getDashboardPeriod() {
    const from = localStorage.getItem('dashboardPeriodFrom');
    const to = localStorage.getItem('dashboardPeriodTo');
    const today = new Date();
    const currentMonth = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
    return {
        from: from || currentMonth,
        to: to || currentMonth
    };
}
function saveDashboardPeriod() {
    const fromEl = document.getElementById('dashboardPeriodFrom');
    const toEl = document.getElementById('dashboardPeriodTo');
    if (fromEl && fromEl.value) localStorage.setItem('dashboardPeriodFrom', fromEl.value);
    if (toEl && toEl.value) localStorage.setItem('dashboardPeriodTo', toEl.value);
}

// Interactive Dashboard period (from/to month)
function getInteractiveDashboardPeriod() {
    const from = localStorage.getItem('interactiveDashboardPeriodFrom');
    const to = localStorage.getItem('interactiveDashboardPeriodTo');
    const period = getDashboardPeriod();
    return {
        from: from || period.from,
        to: to || period.to
    };
}
function saveInteractiveDashboardPeriod() {
    const fromEl = document.getElementById('filterDashboardMonthFrom');
    const toEl = document.getElementById('filterDashboardMonthTo');
    if (fromEl && fromEl.value) localStorage.setItem('interactiveDashboardPeriodFrom', fromEl.value);
    if (toEl && toEl.value) localStorage.setItem('interactiveDashboardPeriodTo', toEl.value);
}

// Initialize
function init() {
    loadData();
    checkAuth();

    // Initialize planner date with current date
    const plannerDateInput = document.getElementById('plannerDate');
    if (plannerDateInput && !plannerDateInput.value) {
        plannerDateInput.value = formatDateString(new Date());
    }

    const footerVer = document.getElementById('appFooterVersion');
    if (footerVer) footerVer.textContent = APP_VERSION;
    const headerVer = document.getElementById('headerAppVersion');
    if (headerVer) headerVer.textContent = `v${APP_VERSION}`;
    const loginVer = document.getElementById('loginFooterVersion');
    if (loginVer) loginVer.textContent = APP_VERSION;

    // Initialize period filters from localStorage
    const period = getDashboardPeriod();
    const fromEl = document.getElementById('dashboardPeriodFrom');
    const toEl = document.getElementById('dashboardPeriodTo');
    if (fromEl) fromEl.value = period.from;
    if (toEl) toEl.value = period.to;

    // Sync Tasks tab period filters
    const filterFrom = document.getElementById('filterTaskMonthFrom');
    const filterTo = document.getElementById('filterTaskMonthTo');
    if (filterFrom) filterFrom.value = period.from;
    if (filterTo) filterTo.value = period.to;

    // Initialize Interactive Dashboard from/to month (sync with dashboard period)
    const interactivePeriod = getInteractiveDashboardPeriod();
    const interactiveFromEl = document.getElementById('filterDashboardMonthFrom');
    const interactiveToEl = document.getElementById('filterDashboardMonthTo');
    if (interactiveFromEl) interactiveFromEl.value = interactivePeriod.from;
    if (interactiveToEl) interactiveToEl.value = interactivePeriod.to;

    processRecurringTasks();
    renderDashboard();
    renderTasks();
    renderCalendar();
    renderUsers();
    renderSettings();
    renderInteractiveDashboard();
    updateLoginScreenCopy();
}

// Data Management
function loadData() {
    if (isApiMode()) {
        if (__workspaceCache !== null) {
            return normalizeData(__workspaceCache);
        }
        return normalizeData(defaultWorkspaceShell());
    }

    const data = localStorage.getItem('todoAppData');
    if (!data) {
        // Try to load from auto-backup if available
        const backup = loadAutoBackup();
        if (backup) {
            if (confirm('No current data found, but an auto-backup was detected. Would you like to restore from backup?')) {
                const normalizedBackup = normalizeData(backup);
                saveData(normalizedBackup);
                return normalizedBackup;
            }
        }

        const defaultData = {
            users: [],
            tasks: [],
            locations: [
                { id: 1, name: 'Mundra' },
                { id: 2, name: 'JNPT' },
                { id: 3, name: 'Combine' }
            ],
            segregationTypes: [
                { id: 1, name: 'PSA Reports' },
                { id: 2, name: 'Internal Reports' }
            ],
            holidays: [],
            notes: [],
            learningNotes: [],
            milestones: [],
            dailyPlanner: [],
            codeSnippets: [],
            journal: {}
        };
        saveData(defaultData);
        return defaultData;
    }
    try {
        const parsedData = JSON.parse(data);
        // Normalize data to ensure all required fields exist (for backward compatibility)
        const normalizedData = normalizeData(parsedData);
        // Save normalized data if it was different
        if (JSON.stringify(parsedData) !== JSON.stringify(normalizedData)) {
            saveData(normalizedData);
        }
        return normalizedData;
    } catch (error) {
        console.error('Error parsing data:', error);
        // Return default data if parsing fails
        const defaultData = {
            users: [],
            tasks: [],
            locations: [
                { id: 1, name: 'Mundra' },
                { id: 2, name: 'JNPT' },
                { id: 3, name: 'Combine' }
            ],
            segregationTypes: [
                { id: 1, name: 'PSA Reports' },
                { id: 2, name: 'Internal Reports' }
            ],
            holidays: [],
            notes: [],
            learningNotes: [],
            milestones: [],
            dailyPlanner: [],
            codeSnippets: [],
            journal: {}
        };
        saveData(defaultData);
        return defaultData;
    }
}

// Normalize data structure to ensure all required fields exist
function normalizeData(data) {
    if (!data) {
        return {
            users: [],
            tasks: [],
            locations: [],
            locationItems: [],
            segregationTypes: [],
            holidays: [],
            notes: [],
            learningNotes: [],
            milestones: [],
            dailyPlanner: [],
            codeSnippets: [],
            journal: {}
        };
    }

    const rawUsers = Array.isArray(data.users) ? data.users : [];
    const users = rawUsers.map(u => {
        if (!u || typeof u !== 'object') return u;
        const idNum = typeof u.id === 'number' && Number.isFinite(u.id)
            ? u.id
            : parseInt(String(u.id), 10);
        const id = Number.isNaN(idNum) ? u.id : idNum;
        return { ...u, id };
    });
    return {
        users,
        tasks: Array.isArray(data.tasks) ? data.tasks : [],
        locations: Array.isArray(data.locations) ? data.locations : (data.locations || []),
        segregationTypes: Array.isArray(data.segregationTypes) ? data.segregationTypes : (data.segregationTypes || []),
        holidays: Array.isArray(data.holidays) ? data.holidays : [],
        notes: Array.isArray(data.notes) ? data.notes : [],
        learningNotes: Array.isArray(data.learningNotes) ? data.learningNotes : [],
        milestones: Array.isArray(data.milestones) ? data.milestones : [],
        dailyPlanner: Array.isArray(data.dailyPlanner) ? data.dailyPlanner : [],
        locationItems: Array.isArray(data.locationItems) ? data.locationItems : [],
        codeSnippets: Array.isArray(data.codeSnippets) ? data.codeSnippets : [],
        journal: data.journal && typeof data.journal === 'object' ? data.journal : {}
    };
}

function saveData(data) {
    const normalized = normalizeData(data);
    if (isApiMode()) {
        __workspaceCache = normalized;
        if (currentUser) {
            scheduleWorkspacePush();
        }
    } else {
        localStorage.setItem('todoAppData', JSON.stringify(normalized));
    }
    // Auto-export if enabled
    if (localStorage.getItem('autoExportEnabled') === 'true') {
        autoExportData(normalized);
    }
}

function getData() {
    return loadData();
}

function updateData(updater) {
    const data = getData();
    // Ensure notes array exists
    if (!data.notes) {
        data.notes = [];
    }
    updater(data);
    saveData(data);
}

function getNextTaskNumberFromData(data) {
    const tasks = data.tasks || [];
    const nums = tasks.map(t => t.task_number).filter(n => n != null && !isNaN(n) && n > 0);
    return nums.length ? Math.max(...nums) + 1 : 1;
}

// Authentication
function checkAuth() {
    const user = localStorage.getItem('currentUser');
    document.body.classList.remove('user-admin');
    document.body.classList.toggle('app-api-mode', isApiMode());
    const masterHint = document.getElementById('masterToolsHint');
    if (user) {
        currentUser = JSON.parse(user);
        document.getElementById('loginModal').classList.remove('active');
        document.getElementById('currentUser').textContent = currentUser.name;
        if (currentUser.role === 'admin') {
            document.body.classList.add('user-admin');
        }
        if (masterHint) {
            masterHint.style.display = currentUser.isMaster ? 'inline-block' : 'none';
            if (masterHint.getAttribute('data-settings-link-wired') !== '1') {
                masterHint.setAttribute('data-settings-link-wired', '1');
                masterHint.addEventListener('click', (ev) => {
                    ev.preventDefault();
                    switchTab('settings');
                });
            }
        }
        const cpBtn = document.getElementById('changePasswordBtn');
        if (cpBtn) cpBtn.style.display = isApiMode() ? 'inline-block' : 'none';
    } else {
        currentUser = null;
        document.body.classList.toggle('app-api-mode', isApiMode());
        document.getElementById('loginModal').classList.add('active');
        document.getElementById('currentUser').textContent = 'Guest';
        if (masterHint) masterHint.style.display = 'none';
        const cpBtn = document.getElementById('changePasswordBtn');
        if (cpBtn) cpBtn.style.display = 'none';
        const pw = document.getElementById('loginPassword');
        const em = document.getElementById('loginEmail');
        const nm = document.getElementById('loginName');
        if (pw) pw.value = '';
        if (em) em.value = '';
        if (nm) nm.value = '';
        const lat = document.getElementById('loginAccountType');
        if (lat) lat.value = 'org_admin';
        setLoginPanelMode('signin');
    }
}

function setLoginPanelMode(mode) {
    const modal = document.getElementById('loginModal');
    if (!modal) return;
    modal.setAttribute('data-login-mode', mode);
    const nameGroup = document.getElementById('loginNameGroup');
    const submitBtn = document.getElementById('loginSubmitBtn');
    const tSign = document.getElementById('loginTabSignin');
    const tReg = document.getElementById('loginTabRegister');
    const tMas = document.getElementById('loginTabMaster');
    [tSign, tReg, tMas].forEach(t => {
        if (t) {
            t.classList.remove('active');
            t.setAttribute('aria-selected', 'false');
        }
    });
    const accountTypeGroup = document.getElementById('loginAccountTypeGroup');
    const orgAdminGroup = document.getElementById('loginOrgAdminGroup');
    if (mode === 'register') {
        if (nameGroup) nameGroup.classList.remove('hidden');
        if (accountTypeGroup) accountTypeGroup.classList.remove('hidden');
        if (submitBtn) submitBtn.textContent = 'Create account';
        if (tReg) {
            tReg.classList.add('active');
            tReg.setAttribute('aria-selected', 'true');
        }
        if (isApiMode()) {
            void loadOrgAdminsForRegisterDropdown();
        } else {
            loadLocalOrgAdminsForRegisterDropdown();
        }
        updateLoginRegistrationFieldsVisibility();
    } else if (mode === 'master') {
        if (nameGroup) nameGroup.classList.add('hidden');
        if (accountTypeGroup) accountTypeGroup.classList.add('hidden');
        if (orgAdminGroup) orgAdminGroup.classList.add('hidden');
        if (submitBtn) submitBtn.textContent = 'Master sign in';
        if (tMas) {
            tMas.classList.add('active');
            tMas.setAttribute('aria-selected', 'true');
        }
    } else {
        if (nameGroup) nameGroup.classList.add('hidden');
        if (accountTypeGroup) accountTypeGroup.classList.add('hidden');
        if (orgAdminGroup) orgAdminGroup.classList.add('hidden');
        if (submitBtn) submitBtn.textContent = 'Sign in';
        if (tSign) {
            tSign.classList.add('active');
            tSign.setAttribute('aria-selected', 'true');
        }
    }
    updateLoginScreenCopy();
}

function updateLoginScreenCopy() {
    const sub = document.getElementById('loginBrandSubtitle');
    if (!sub) return;
    if (isApiMode()) {
        const modal = document.getElementById('loginModal');
        const mode = (modal && modal.getAttribute('data-login-mode')) || 'signin';
        if (mode === 'register') {
            sub.textContent = 'Account admin creates a new workspace; account user joins an existing admin. Cloud signup uses the server.';
        } else if (mode === 'master') {
            sub.textContent = 'Master sign-in for cross-tenant support tools only.';
        } else {
            sub.textContent = 'Sign in — your workspace loads from the server.';
        }
    } else {
        sub.textContent = 'Sign in or create an account. Add an API URL in config.js to use cloud storage; otherwise data stays in this browser only.';
    }
}

function submitLoginPanel() {
    const modal = document.getElementById('loginModal');
    const mode = (modal && modal.getAttribute('data-login-mode')) || 'signin';
    if (mode === 'register') {
        register();
    } else {
        login();
    }
}

/** Wire login UI without inline onclick (avoids ReferenceError on some static hosts / cached HTML+JS mismatches). */
function wireLoginScreenControls() {
    const modal = document.getElementById('loginModal');
    if (!modal || modal.getAttribute('data-wired-ui') === '1') return;
    modal.setAttribute('data-wired-ui', '1');
    const submitBtn = document.getElementById('loginSubmitBtn');
    if (submitBtn) submitBtn.addEventListener('click', submitLoginPanel);
    const tSign = document.getElementById('loginTabSignin');
    const tReg = document.getElementById('loginTabRegister');
    const tMas = document.getElementById('loginTabMaster');
    if (tSign) tSign.addEventListener('click', () => setLoginPanelMode('signin'));
    if (tReg) tReg.addEventListener('click', () => setLoginPanelMode('register'));
    if (tMas) tMas.addEventListener('click', () => setLoginPanelMode('master'));
    const lat = document.getElementById('loginAccountType');
    if (lat) lat.addEventListener('change', updateLoginRegistrationFieldsVisibility);
}

function updateLoginRegistrationFieldsVisibility() {
    const modal = document.getElementById('loginModal');
    if (!modal || modal.getAttribute('data-login-mode') !== 'register') return;
    const typeEl = document.getElementById('loginAccountType');
    const oag = document.getElementById('loginOrgAdminGroup');
    if (!oag || !typeEl) return;
    if (typeEl.value === 'team_user') {
        oag.classList.remove('hidden');
    } else {
        oag.classList.add('hidden');
    }
}

function loadLocalOrgAdminsForRegisterDropdown() {
    const sel = document.getElementById('loginOrgAdminSelect');
    if (!sel) return;
    const data = getData();
    const admins = (data.users || []).filter(u => u.role === 'admin' && !isMasterUserRecord(u));
    sel.innerHTML = '<option value="">Select account admin…</option>'
        + admins.map(u => `<option value="${u.id}">${escapeHtml(u.name)} (${escapeHtml(u.email)})</option>`).join('');
    updateLoginRegistrationFieldsVisibility();
}

async function loadOrgAdminsForRegisterDropdown() {
    const sel = document.getElementById('loginOrgAdminSelect');
    if (!sel) return;
    sel.innerHTML = '<option value="">Loading…</option>';
    try {
        const res = await apiFetch('/api/auth/org-admins');
        if (!res.ok) throw new Error('load failed');
        const list = await res.json();
        if (!Array.isArray(list) || list.length === 0) {
            sel.innerHTML = '<option value="">No account admins yet — register as account admin first</option>';
        } else {
            const opts = list.map(a => `<option value="${a.id}">${escapeHtml(a.name)} (${escapeHtml(a.email)})</option>`).join('');
            sel.innerHTML = `<option value="">Select account admin…</option>${opts}`;
        }
    } catch (e) {
        console.error(e);
        sel.innerHTML = '<option value="">Could not load account admins</option>';
    }
    updateLoginRegistrationFieldsVisibility();
}

async function login() {
    const email = document.getElementById('loginEmail').value.trim();
    const password = document.getElementById('loginPassword').value;

    if (!email || !password) {
        showError('loginError', 'Please enter email and password.');
        return;
    }

    if (isApiMode()) {
        let res;
        try {
            res = await apiFetch('/api/auth/login', {
                method: 'POST',
                body: JSON.stringify({ email, password })
            });
        } catch (e) {
            console.error(e);
            showError('loginError', 'Could not reach server. Check API URL and network.');
            return;
        }
        if (!res.ok) {
            let errMsg = `Sign-in failed (${res.status})`;
            try {
                const t = await res.text();
                if (t) {
                    try {
                        const j = JSON.parse(t);
                        if (j.error) errMsg = j.error;
                    } catch {
                        errMsg = t.slice(0, 200);
                    }
                }
            } catch (_) { /* ignore */ }
            showError('loginError', errMsg);
            return;
        }
        let body;
        try {
            body = await res.json();
        } catch (parseErr) {
            console.error('Login response parse:', parseErr);
            try {
                const me = await apiFetch('/api/auth/me');
                if (me.ok) {
                    try {
                        currentUser = await me.json();
                    } catch (meParse) {
                        console.error(meParse);
                        throw meParse;
                    }
                    localStorage.setItem('currentUser', JSON.stringify(currentUser));
                    try {
                        await apiPullWorkspace();
                    } catch (pullErr) {
                        console.error('Workspace load after login:', pullErr);
                        __workspaceCache = normalizeData(defaultWorkspaceShell());
                    }
                    const lp = document.getElementById('loginPassword');
                    if (lp) lp.value = '';
                    clearLoginFormError();
                    checkAuth();
                    try {
                        init();
                    } catch (initErr) {
                        console.error('init after login:', initErr);
                    }
                    return;
                }
            } catch (recoverErr) {
                console.error(recoverErr);
            }
            showError('loginError', 'Signed in may have worked — try refreshing. Otherwise check API URL and CORS.');
            return;
        }
        currentUser = body.user;
        if (body.token) setApiAuthToken(body.token);
        localStorage.setItem('currentUser', JSON.stringify(currentUser));
        try {
            await apiPullWorkspace();
        } catch (pullErr) {
            console.error('Workspace load after login:', pullErr);
            __workspaceCache = normalizeData(defaultWorkspaceShell());
        }
        const lp = document.getElementById('loginPassword');
        if (lp) lp.value = '';
        clearLoginFormError();
        checkAuth();
        try {
            init();
        } catch (initErr) {
            console.error('init after login:', initErr);
        }
        return;
    }

    const data = getData();
    const user = data.users.find(u => u.email === email);
    if (!user || user.password !== password) {
        showError('loginError', 'Invalid email or password');
        return;
    }

    if (!user.is_active) {
        showError('loginError', 'Account is disabled');
        return;
    }

    currentUser = { ...user };
    delete currentUser.password;
    localStorage.setItem('currentUser', JSON.stringify(currentUser));
    const lp = document.getElementById('loginPassword');
    if (lp) lp.value = '';
    checkAuth();
    init();
}

async function register() {
    const loginModalEl = document.getElementById('loginModal');
    if (!loginModalEl || loginModalEl.getAttribute('data-login-mode') !== 'register') {
        showError('loginError', 'Use the "Create account" tab to register.');
        return;
    }
    const name = document.getElementById('loginName').value.trim();
    const email = document.getElementById('loginEmail').value.trim();
    const password = document.getElementById('loginPassword').value;

    if (!name || !email || !password || password.length < 6) {
        showError('loginError', 'Please fill all fields. Password must be at least 6 characters.');
        return;
    }

    const accountTypeEl = document.getElementById('loginAccountType');
    const accountType = accountTypeEl && accountTypeEl.value === 'team_user' ? 'team_user' : 'org_admin';
    let orgAdminUserId;
    if (accountType === 'team_user') {
        const sel = document.getElementById('loginOrgAdminSelect');
        orgAdminUserId = sel ? parseInt(sel.value, 10) : NaN;
        if (Number.isNaN(orgAdminUserId) || orgAdminUserId <= 0) {
            showError('loginError', 'Select an account admin for account user signup.');
            return;
        }
    }

    if (isApiMode()) {
        let res;
        const regBody = { name, email, password, accountType };
        if (accountType === 'team_user') regBody.orgAdminUserId = orgAdminUserId;
        try {
            res = await apiFetch('/api/auth/register', {
                method: 'POST',
                body: JSON.stringify(regBody)
            });
        } catch (e) {
            console.error(e);
            showError('loginError', 'Could not reach server. Check API URL and network.');
            return;
        }
        if (!res.ok) {
            let errMsg = `Registration failed (${res.status})`;
            try {
                const t = await res.text();
                if (t) {
                    try {
                        const j = JSON.parse(t);
                        if (j.error) errMsg = j.error;
                    } catch {
                        errMsg = t.slice(0, 200);
                    }
                }
            } catch (_) { /* ignore */ }
            showError('loginError', errMsg);
            return;
        }
        let body;
        try {
            body = await res.json();
        } catch (parseErr) {
            console.error('Register response parse:', parseErr);
            try {
                const me = await apiFetch('/api/auth/me');
                if (me.ok) {
                    try {
                        currentUser = await me.json();
                    } catch (meParse) {
                        console.error(meParse);
                        throw meParse;
                    }
                    localStorage.setItem('currentUser', JSON.stringify(currentUser));
                    try {
                        await apiPullWorkspace();
                    } catch (pullErr) {
                        console.error('Workspace load after register:', pullErr);
                        __workspaceCache = normalizeData(defaultWorkspaceShell());
                    }
                    const lp = document.getElementById('loginPassword');
                    if (lp) lp.value = '';
                    clearLoginFormError();
                    checkAuth();
                    try {
                        init();
                    } catch (initErr) {
                        console.error('init after register:', initErr);
                    }
                    return;
                }
            } catch (recoverErr) {
                console.error(recoverErr);
            }
            showError('loginError', 'Registration may have succeeded — try signing in. If the problem continues, check API URL and CORS settings.');
            return;
        }
        currentUser = body.user;
        if (body.token) setApiAuthToken(body.token);
        localStorage.setItem('currentUser', JSON.stringify(currentUser));
        try {
            await apiPullWorkspace();
        } catch (pullErr) {
            console.error('Workspace load after register:', pullErr);
            __workspaceCache = normalizeData(defaultWorkspaceShell());
        }
        const lp = document.getElementById('loginPassword');
        if (lp) lp.value = '';
        clearLoginFormError();
        checkAuth();
        try {
            init();
        } catch (initErr) {
            console.error('init after register:', initErr);
        }
        return;
    }

    const data = getData();
    if (data.users.find(u => u.email === email)) {
        showError('loginError', 'Email already registered');
        return;
    }

    let newUser;
    if (accountType === 'team_user') {
        const adminUser = data.users.find(u =>
            Number(u.id) === orgAdminUserId && u.role === 'admin' && !isMasterUserRecord(u));
        if (!adminUser) {
            showError('loginError', 'Invalid account admin.');
            return;
        }
        newUser = {
            id: Date.now(),
            name,
            email,
            password,
            role: 'user',
            is_active: true,
            tenant_root_id: Number(adminUser.id)
        };
    } else {
        newUser = {
            id: Date.now(),
            name,
            email,
            password,
            role: 'admin',
            is_active: true
        };
    }

    data.users.push(newUser);
    saveData(data);

    currentUser = { ...newUser };
    delete currentUser.password;
    localStorage.setItem('currentUser', JSON.stringify(currentUser));
    const lp = document.getElementById('loginPassword');
    if (lp) lp.value = '';
    checkAuth();
    init();
}

window.setLoginPanelMode = setLoginPanelMode;
window.submitLoginPanel = submitLoginPanel;
window.openChangePasswordModal = openChangePasswordModal;
window.closeChangePasswordModal = closeChangePasswordModal;
window.submitChangePassword = submitChangePassword;

function openChangePasswordModal() {
    if (!isApiMode() || !currentUser) return;
    const err = document.getElementById('changePasswordError');
    if (err) err.textContent = '';
    const c = document.getElementById('changePasswordCurrent');
    const n = document.getElementById('changePasswordNew');
    const q = document.getElementById('changePasswordConfirm');
    if (c) c.value = '';
    if (n) n.value = '';
    if (q) q.value = '';
    const m = document.getElementById('changePasswordModal');
    if (m) m.classList.add('active');
}

function closeChangePasswordModal() {
    const m = document.getElementById('changePasswordModal');
    if (m) m.classList.remove('active');
}

async function submitChangePassword(event) {
    event.preventDefault();
    const errEl = document.getElementById('changePasswordError');
    if (errEl) errEl.textContent = '';
    const currentPassword = document.getElementById('changePasswordCurrent')?.value || '';
    const newPassword = document.getElementById('changePasswordNew')?.value || '';
    const confirm = document.getElementById('changePasswordConfirm')?.value || '';
    if (newPassword.length < 6) {
        if (errEl) errEl.textContent = 'New password must be at least 6 characters.';
        return;
    }
    if (newPassword !== confirm) {
        if (errEl) errEl.textContent = 'New password and confirmation do not match.';
        return;
    }
    try {
        const res = await apiFetch('/api/auth/change-password', {
            method: 'POST',
            body: JSON.stringify({ currentPassword, newPassword })
        });
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            if (errEl) errEl.textContent = err.error || 'Could not update password.';
            return;
        }
        __pendingPasswordsByUserId.delete(Number(currentUser.id));
        closeChangePasswordModal();
        alert('Password updated. Use your new password next time you sign in on another device.');
    } catch (e) {
        console.error(e);
        if (errEl) errEl.textContent = 'Network error. Check connection and API URL.';
    }
}

async function logout() {
    if (isApiMode()) {
        try {
            await apiFetch('/api/auth/logout', { method: 'POST' });
        } catch (_) {
            /* ignore */
        }
        __workspaceCache = null;
        clearPendingPasswordsForSync();
    }
    clearApiAuthToken();
    localStorage.removeItem('currentUser');
    currentUser = null;
    document.body.classList.remove('user-admin');
    checkAuth();
}

function showError(elementId, message) {
    const element = document.getElementById(elementId);
    if (!element) return;
    if (elementId === 'loginError' && __loginErrorDismissTimer) {
        clearTimeout(__loginErrorDismissTimer);
        __loginErrorDismissTimer = null;
    }
    element.innerHTML = `<div class="alert alert-error">${message}</div>`;
    const ms = elementId === 'loginError' ? 8000 : 3000;
    const t = setTimeout(() => {
        element.innerHTML = '';
        if (elementId === 'loginError') __loginErrorDismissTimer = null;
    }, ms);
    if (elementId === 'loginError') __loginErrorDismissTimer = t;
}

function clearLoginFormError() {
    if (__loginErrorDismissTimer) {
        clearTimeout(__loginErrorDismissTimer);
        __loginErrorDismissTimer = null;
    }
    const el = document.getElementById('loginError');
    if (el) el.innerHTML = '';
}

// Tab Navigation
function switchTab(tabName, eventElement, skipAutoRender = false) {
    const canSettings = currentUser && (currentUser.role === 'admin' || currentUser.isMaster);
    if (tabName === 'settings' && !canSettings) {
        return;
    }
    document.querySelectorAll('.nav-tab').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));

    if (eventElement) {
        eventElement.classList.add('active');
    } else {
        // Find the tab button and activate it
        const tabs = document.querySelectorAll('.nav-tab');
        tabs.forEach(tab => {
            const onclick = tab.getAttribute('onclick');
            if (onclick && onclick.includes(`'${tabName}'`)) {
                tab.classList.add('active');
            }
        });
    }

    document.getElementById(tabName).classList.add('active');

    if (!skipAutoRender) {
        if (tabName === 'dashboard') renderDashboard();
        if (tabName === 'interactive') renderInteractiveDashboard();
        if (tabName === 'drilldown') renderDrilldown();
        if (tabName === 'tasks') renderTasks();
        if (tabName === 'calendar') renderCalendar();
        if (tabName === 'milestones') renderMilestones();
        if (tabName === 'planner') renderDailyPlanner();
        if (tabName === 'notes') renderNotes();
        if (tabName === 'learningNotes') renderLearningNotes();
        if (tabName === 'locations') renderLocations();
        if (tabName === 'snippets') renderCodeSnippets();
        if (tabName === 'journal') renderJournal();
        if (tabName === 'settings') renderSettings();
        if (tabName === 'recurringReport') renderRecurringReport();
    }
}

// Dashboard
function renderDashboard() {
    const data = getData();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayStr = formatDateString(today);

    // Get selected period from filters
    const period = getDashboardPeriod();
    const [fromYear, fromMonth] = period.from.split('-').map(Number);
    const [toYear, toMonth] = period.to.split('-').map(Number);
    const fromDate = new Date(fromYear, fromMonth - 1, 1);
    fromDate.setHours(0, 0, 0, 0);
    const toDate = new Date(toYear, toMonth, 0); // Last day of to month
    toDate.setHours(23, 59, 59, 999);

    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

    let todayTasks = [];
    let overdue = 0;
    let overdueTasks = []; // Array to collect overdue tasks
    let completed = 0;
    let pending = 0;
    let noDueDate = 0;
    let noDueDateTasks = []; // Array to collect no due date tasks
    let inProcessTasks = []; // Array to collect in-process tasks (task_action === 'in_process')
    let workPlan = 0;
    let workPlanTasks = []; // Array to collect work plan tasks
    let auditPoint = 0;
    let adminReview = 0;

    // User-wise statistics
    const userStats = {};
    data.users.forEach(user => {
        userStats[user.id] = {
            name: user.name,
            totalAssigned: 0,
            totalCompleted: 0,
            totalPending: 0,
            totalNeedImprovement: 0
        };
    });

    data.tasks.forEach(task => {
        const dueDate = task.due_date || task.next_due_date;

        // Check if task falls in selected period
        let taskInSelectedPeriod = false;
        if (dueDate) {
            const dateParts = dueDate.split('-');
            const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
            taskDate.setHours(0, 0, 0, 0);
            taskInSelectedPeriod = (taskDate >= fromDate && taskDate <= toDate);
        }

        // For user statistics: count tasks in selected period for all users (admin sees all, regular users see their own)
        if (taskInSelectedPeriod || !dueDate) {
            // For admin: show stats for all users; for regular users: only their own tasks
            const shouldCountForUser = currentUser.role === 'admin' || task.assigned_to === currentUser.id;

            if (shouldCountForUser && userStats[task.assigned_to]) {
                userStats[task.assigned_to].totalAssigned++;

                if (task.task_action === 'completed') {
                    userStats[task.assigned_to].totalCompleted++;
                } else if (task.task_action === 'completed_need_improvement') {
                    userStats[task.assigned_to].totalNeedImprovement++;
                } else {
                    userStats[task.assigned_to].totalPending++;
                }
            }
        }

        // For other dashboard stats (today, overdue, etc.), check if task is visible to current user
        const isAssigned = currentUser.role === 'admin' ||
            task.assigned_to === currentUser.id ||
            (task.is_team_task && currentUser.role === 'admin');

        if (!isAssigned) return;

        // Count tasks without due date (exclude Need Improvement and Not Done)
        if (task.task_type === 'without_due_date' &&
            task.task_action !== 'completed' &&
            task.task_action !== 'completed_need_improvement' &&
            task.task_action !== 'not_done') {
            noDueDate++;
            noDueDateTasks.push(task);
        }

        // Count in-process tasks (task_action === 'in_process')
        if (task.task_action === 'in_process' && !task.removed_at) {
            inProcessTasks.push(task);
        }

        // Count work plan tasks (in selected period; exclude Need Improvement and Not Done)
        if (task.task_type === 'work_plan' &&
            task.task_action !== 'completed' &&
            task.task_action !== 'completed_need_improvement' &&
            task.task_action !== 'not_done') {
            if (dueDate) {
                const dateParts = dueDate.split('-');
                const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
                taskDate.setHours(0, 0, 0, 0);
                if (taskDate >= fromDate && taskDate <= toDate) {
                    workPlan++;
                    workPlanTasks.push(task);
                }
            } else {
                // If no due date, count it
                workPlan++;
                workPlanTasks.push(task);
            }
        }

        // Count audit point tasks (in selected period; exclude Need Improvement and Not Done)
        if (task.task_type === 'audit_point' &&
            task.task_action !== 'completed' &&
            task.task_action !== 'completed_need_improvement' &&
            task.task_action !== 'not_done') {
            if (dueDate) {
                const dateParts = dueDate.split('-');
                const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
                taskDate.setHours(0, 0, 0, 0);
                if (taskDate >= fromDate && taskDate <= toDate) {
                    auditPoint++;
                }
            } else {
                // If no due date, count it
                auditPoint++;
            }
        }

        // Count tasks pending admin review (completed by user but not finalized by admin)
        if (currentUser.role === 'admin' && task.task_action === 'completed' && !task.admin_finalized) {
            adminReview++;
        }

        // Exclude tasks without due date from date-based calculations
        if (task.task_type !== 'without_due_date') {
            const isClosedForOverdue = task.task_action === 'completed' ||
                task.task_action === 'completed_need_improvement' ||
                task.task_action === 'not_done';

            if (dueDate === todayStr && !isClosedForOverdue && task.task_action !== 'in_process') {
                todayTasks.push(task);
            }

            if (dueDate && !isClosedForOverdue && task.task_action !== 'in_process') {
                const dateParts = dueDate.split('-');
                const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
                taskDate.setHours(0, 0, 0, 0);

                if (taskDate < today) {
                    overdue++;
                    overdueTasks.push(task);
                }

                // Pending count for selected period only
                if (taskDate >= fromDate && taskDate <= toDate) {
                    pending++;
                }
            }
        }

        // Completed count: only task_action === 'completed' (exclude Need Improvement from dashboard)
        if (task.task_action === 'completed' &&
            task.completed_at && new Date(task.completed_at) >= thirtyDaysAgo) {
            completed++;
        }
    });

    document.getElementById('statToday').textContent = todayTasks.length;
    document.getElementById('statOverdue').textContent = overdue;
    document.getElementById('statCompleted').textContent = completed;
    document.getElementById('statPending').textContent = pending;
    document.getElementById('statNoDueDate').textContent = noDueDate;
    document.getElementById('statWorkPlan').textContent = workPlan;
    document.getElementById('statAuditPoint').textContent = auditPoint;

    // Show/hide admin review tile
    const adminReviewCard = document.getElementById('adminReviewCard');
    if (currentUser.role === 'admin') {
        adminReviewCard.style.display = 'block';
        document.getElementById('statAdminReview').textContent = adminReview;
    } else {
        adminReviewCard.style.display = 'none';
    }

    // Render user-wise statistics table
    renderUserStatsTable(userStats, fromYear, fromMonth - 1, toYear, toMonth - 1);

    // Render detailed statistics table
    renderDetailedStatsTable(userStats);

    // Make tiles clickable with proper filtering
    const dashboardPeriodValue = period;

    // Use event delegation - attach a single handler to the stats grid container
    const statsGrid = document.querySelector('.stats-grid');
    if (statsGrid) {
        // Remove any existing handler
        statsGrid.onclick = null;

        // Add data attributes to identify each card
        const todayCard = document.getElementById('statToday').parentElement;
        todayCard.setAttribute('data-tile-type', 'today');
        todayCard.style.cursor = 'pointer';
        todayCard.style.transition = 'transform 0.2s';
        todayCard.onmouseenter = () => todayCard.style.transform = 'scale(1.02)';
        todayCard.onmouseleave = () => todayCard.style.transform = 'scale(1)';

        const overdueCard = document.getElementById('statOverdue').parentElement;
        overdueCard.setAttribute('data-tile-type', 'overdue');
        overdueCard.style.cursor = 'pointer';
        overdueCard.style.transition = 'transform 0.2s';
        overdueCard.onmouseenter = () => overdueCard.style.transform = 'scale(1.02)';
        overdueCard.onmouseleave = () => overdueCard.style.transform = 'scale(1)';

        const completedCard = document.getElementById('statCompleted').parentElement;
        completedCard.setAttribute('data-tile-type', 'completed');
        completedCard.style.cursor = 'pointer';
        completedCard.style.transition = 'transform 0.2s';
        completedCard.onmouseenter = () => completedCard.style.transform = 'scale(1.02)';
        completedCard.onmouseleave = () => completedCard.style.transform = 'scale(1)';

        const pendingCard = document.getElementById('statPending').parentElement;
        pendingCard.setAttribute('data-tile-type', 'pending');
        pendingCard.style.cursor = 'pointer';
        pendingCard.style.transition = 'transform 0.2s';
        pendingCard.onmouseenter = () => pendingCard.style.transform = 'scale(1.02)';
        pendingCard.onmouseleave = () => pendingCard.style.transform = 'scale(1)';

        const noDueDateCard = document.getElementById('statNoDueDate').parentElement;
        noDueDateCard.setAttribute('data-tile-type', 'no_due_date');
        noDueDateCard.style.cursor = 'pointer';
        noDueDateCard.style.transition = 'transform 0.2s';
        noDueDateCard.onmouseenter = () => noDueDateCard.style.transform = 'scale(1.02)';
        noDueDateCard.onmouseleave = () => noDueDateCard.style.transform = 'scale(1)';

        const workPlanCard = document.getElementById('statWorkPlan').parentElement;
        workPlanCard.setAttribute('data-tile-type', 'work_plan');
        workPlanCard.style.cursor = 'pointer';
        workPlanCard.style.transition = 'transform 0.2s';
        workPlanCard.onmouseenter = () => workPlanCard.style.transform = 'scale(1.02)';
        workPlanCard.onmouseleave = () => workPlanCard.style.transform = 'scale(1)';

        const auditPointCard = document.getElementById('statAuditPoint').parentElement;
        auditPointCard.setAttribute('data-tile-type', 'audit_point');
        auditPointCard.style.cursor = 'pointer';
        auditPointCard.style.transition = 'transform 0.2s';
        auditPointCard.onmouseenter = () => auditPointCard.style.transform = 'scale(1.02)';
        auditPointCard.onmouseleave = () => auditPointCard.style.transform = 'scale(1)';

        if (currentUser.role === 'admin') {
            const adminReviewCard = document.getElementById('statAdminReview').parentElement;
            adminReviewCard.setAttribute('data-tile-type', 'admin_review');
            adminReviewCard.style.cursor = 'pointer';
            adminReviewCard.style.transition = 'transform 0.2s';
            adminReviewCard.onmouseenter = () => adminReviewCard.style.transform = 'scale(1.02)';
            adminReviewCard.onmouseleave = () => adminReviewCard.style.transform = 'scale(1)';
        }

        // Single event delegation handler
        statsGrid.onclick = (e) => {
            e.preventDefault();
            e.stopPropagation();

            // Find the clicked card (stat-card)
            let clickedCard = e.target;
            while (clickedCard && !clickedCard.hasAttribute('data-tile-type')) {
                clickedCard = clickedCard.parentElement;
            }

            if (!clickedCard || !clickedCard.hasAttribute('data-tile-type')) {
                return false;
            }

            const tileType = clickedCard.getAttribute('data-tile-type');
            console.log('Dashboard tile clicked, type:', tileType);

            // Set drilldown context based on tile type
            // Hide date filters by default (only shown for overdue)
            const dateFilters = document.getElementById('drilldownDateFilters');
            if (dateFilters) dateFilters.style.display = 'none';

            const syncTaskPeriod = () => {
                const p = dashboardPeriodValue;
                const fromEl = document.getElementById('filterTaskMonthFrom');
                const toEl = document.getElementById('filterTaskMonthTo');
                if (fromEl && p?.from) fromEl.value = p.from;
                if (toEl && p?.to) toEl.value = p.to;
            };
            const periodFrom = dashboardPeriodValue?.from;
            const periodTo = dashboardPeriodValue?.to;
            switch (tileType) {
                case 'today':
                    drilldownContext = {
                        type: 'today',
                        title: "Today's Tasks",
                        count: todayTasks.length,
                        filterFunction: () => filterTasksByDate(todayStr),
                        monthValue: periodTo,
                        periodFrom, periodTo,
                        dateStr: todayStr
                    };
                    break;
                case 'overdue':
                    drilldownContext = {
                        type: 'overdue',
                        title: 'Overdue Tasks',
                        count: overdue,
                        filterFunction: () => {
                            const dateFilters = document.getElementById('drilldownDateFilters');
                            if (dateFilters) dateFilters.style.display = 'block';
                            const fromMonthInput = document.getElementById('drilldownFromMonth');
                            const toMonthInput = document.getElementById('drilldownToMonth');
                            const p = getDashboardPeriod();
                            if (fromMonthInput && !fromMonthInput.value) fromMonthInput.value = p.from;
                            if (toMonthInput && !toMonthInput.value) toMonthInput.value = p.to;
                            renderOverdueDrilldown();
                        },
                        monthValue: periodTo,
                        periodFrom, periodTo
                    };
                    break;
                case 'completed':
                    drilldownContext = {
                        type: 'completed',
                        title: 'Completed Tasks (Last 30 Days)',
                        count: completed,
                        filterFunction: () => {
                            syncTaskPeriod();
                            document.getElementById('filterStatus').value = 'completed';
                            document.getElementById('filterType').value = '';
                            document.getElementById('filterTeam').value = '';
                            document.getElementById('searchTasks').value = '';
                            filterTasks();
                        },
                        monthValue: periodTo,
                        periodFrom, periodTo
                    };
                    break;
                case 'pending':
                    drilldownContext = {
                        type: 'pending',
                        title: 'Pending Tasks',
                        count: pending,
                        filterFunction: () => { syncTaskPeriod(); filterTasksByPendingMonth(); },
                        monthValue: periodTo,
                        periodFrom, periodTo
                    };
                    break;
                case 'no_due_date':
                    drilldownContext = {
                        type: 'no_due_date',
                        title: 'Tasks Without Due Date',
                        count: noDueDate,
                        filterFunction: () => {
                            syncTaskPeriod();
                            document.getElementById('filterStatus').value = 'not_completed';
                            document.getElementById('filterType').value = 'without_due_date';
                            document.getElementById('filterTeam').value = '';
                            document.getElementById('searchTasks').value = '';
                            filterTasks();
                        },
                        monthValue: periodTo,
                        periodFrom, periodTo
                    };
                    break;
                case 'work_plan':
                    drilldownContext = {
                        type: 'work_plan',
                        title: 'Work Plan Tasks',
                        count: workPlan,
                        filterFunction: () => {
                            syncTaskPeriod();
                            document.getElementById('filterStatus').value = 'not_completed';
                            document.getElementById('filterType').value = 'work_plan';
                            document.getElementById('filterTeam').value = '';
                            document.getElementById('searchTasks').value = '';
                            filterTasks();
                        },
                        monthValue: periodTo,
                        periodFrom, periodTo
                    };
                    break;
                case 'audit_point':
                    drilldownContext = {
                        type: 'audit_point',
                        title: 'Audit Point Tasks',
                        count: auditPoint,
                        filterFunction: () => {
                            syncTaskPeriod();
                            document.getElementById('filterStatus').value = 'not_completed';
                            document.getElementById('filterType').value = 'audit_point';
                            document.getElementById('filterTeam').value = '';
                            document.getElementById('searchTasks').value = '';
                            filterTasks();
                        },
                        monthValue: periodTo,
                        periodFrom, periodTo
                    };
                    break;
                case 'admin_review':
                    if (currentUser.role === 'admin') {
                        drilldownContext = {
                            type: 'admin_review',
                            title: 'Pending Admin Review',
                            count: adminReview,
                            filterFunction: () => {
                                syncTaskPeriod();
                                document.getElementById('filterStatus').value = '';
                                document.getElementById('filterType').value = '';
                                document.getElementById('filterTeam').value = '';
                                document.getElementById('searchTasks').value = '';
                                filterTasksByAdminReview();
                            },
                            monthValue: periodTo,
                            periodFrom, periodTo
                        };
                    } else {
                        return false;
                    }
                    break;
                default:
                    return false;
            }

            console.log('Drilldown context set:', drilldownContext);
            switchTab('drilldown', null, true);
            setTimeout(() => {
                renderDrilldown();
            }, 50);
            return false;
        };
    }


    todayTasks.sort((a, b) => {
        const priorityOrder = { high: 1, medium: 2, low: 3 };
        const aPriority = priorityOrder[a.priority] || 4;
        const bPriority = priorityOrder[b.priority] || 4;
        if (aPriority !== bPriority) return aPriority - bPriority;
        return (a.due_date || a.next_due_date || '').localeCompare(b.due_date || b.next_due_date || '');
    });

    const todayTasksHtml = todayTasks.length > 0
        ? todayTasks.map(task => renderTaskItem(task, true)).join('')
        : '<p style="text-align: center; color: #999; padding: 20px;">No tasks for today</p>';

    document.getElementById('todayTasks').innerHTML = todayTasksHtml;

    // Render task tiles for quick overview
    renderTaskTiles(todayTasks, overdueTasks, noDueDateTasks, inProcessTasks, workPlanTasks);
}

// Render task tiles on dashboard
function renderTaskTiles(todayTasks, overdueTasks, noDueDateTasks, inProcessTasks, workPlanTasks) {
    const data = getData();

    // Helper function to render a single task item for tiles
    const renderTileTaskItem = (task) => {
        const dueDate = task.due_date || task.next_due_date;
        const dueDateStr = dueDate ? formatDateDisplay(dueDate) : 'No due date';
        const assignedUser = data.users.find(u => u.id === task.assigned_to);

        return `
            <div onclick="openInteractiveTaskPopup(${task.id})" 
                 style="padding: 10px; background: #f8f9fa; border-radius: 5px; cursor: pointer; transition: all 0.2s; border-left: 3px solid #667eea;"
                 onmouseenter="this.style.background='#e9ecef'; this.style.transform='translateX(5px)'"
                 onmouseleave="this.style.background='#f8f9fa'; this.style.transform='translateX(0)'">
                <div style="font-size: 13px; font-weight: 500; color: #333; margin-bottom: 4px;">
                    ${task.task_name}
                </div>
                <div style="font-size: 11px; color: #666; display: flex; justify-content: space-between; align-items: center;">
                    <span>${assignedUser ? assignedUser.name : 'Unassigned'}</span>
                    <span style="color: #999;">${dueDateStr}</span>
                </div>
            </div>
        `;
    };

    // Helper for In Process tile: show expected completion date and allow editing
    const renderInProcessTileItem = (task) => {
        const dueDate = task.due_date || task.next_due_date;
        const dueDateStr = dueDate ? formatDateDisplay(dueDate) : 'No due date';
        const expectedStr = task.expected_completion_date ? formatDateDisplay(task.expected_completion_date) : 'Not set';
        const assignedUser = data.users.find(u => u.id === task.assigned_to);

        return `
            <div style="padding: 10px; background: #f8f9fa; border-radius: 5px; border-left: 3px solid #17a2b8; margin-bottom: 8px;">
                <div onclick="openInteractiveTaskPopup(${task.id})" style="cursor: pointer; margin-bottom: 6px;"
                     onmouseenter="this.style.opacity='0.9'" onmouseleave="this.style.opacity='1'">
                    <div style="font-size: 13px; font-weight: 500; color: #333; margin-bottom: 4px;">${task.task_name}</div>
                    <div style="font-size: 11px; color: #666;">
                        <span>${assignedUser ? assignedUser.name : 'Unassigned'}</span>
                        <span style="color: #999;"> | Due: ${dueDateStr}</span>
                    </div>
                </div>
                <div style="font-size: 11px; color: #17a2b8; display: flex; align-items: center; justify-content: space-between; gap: 6px;">
                    <span>Expected: ${expectedStr}</span>
                    <button type="button" class="btn btn-sm" onclick="event.stopPropagation(); openEditExpectedDateModal(${task.id})" 
                            style="padding: 2px 8px; font-size: 11px; background: #17a2b8; color: white; border: none; border-radius: 4px; cursor: pointer;">Edit date</button>
                </div>
            </div>
        `;
    };

    // Today's Work Tile (all tasks; list has scroll for >5)
    const todayWorkList = document.getElementById('todayWorkTileList');
    if (todayWorkList) {
        const sortedToday = [...todayTasks]
            .filter(t => t.task_action !== 'completed')
            .sort((a, b) => {
                const priorityOrder = { high: 1, medium: 2, low: 3 };
                return (priorityOrder[a.priority] || 4) - (priorityOrder[b.priority] || 4);
            });

        todayWorkList.innerHTML = sortedToday.length > 0
            ? sortedToday.map(renderTileTaskItem).join('')
            : '<div style="text-align: center; color: #999; padding: 15px; font-size: 13px;">No tasks for today</div>';
    }

    // Overdue Tile (all tasks; list has scroll for >5)
    const overdueList = document.getElementById('overdueTileList');
    if (overdueList) {
        const sortedOverdue = [...overdueTasks]
            .sort((a, b) => (a.due_date || a.next_due_date || '').localeCompare(b.due_date || b.next_due_date || ''));

        overdueList.innerHTML = sortedOverdue.length > 0
            ? sortedOverdue.map(renderTileTaskItem).join('')
            : '<div style="text-align: center; color: #999; padding: 15px; font-size: 13px;">No overdue tasks</div>';
    }

    // No Due Date Tile (all tasks; list has scroll for >5)
    const noDueDateList = document.getElementById('noDueDateTileList');
    if (noDueDateList) {
        const sortedNoDueDate = [...noDueDateTasks];

        noDueDateList.innerHTML = sortedNoDueDate.length > 0
            ? sortedNoDueDate.map(renderTileTaskItem).join('')
            : '<div style="text-align: center; color: #999; padding: 15px; font-size: 13px;">No tasks without due date</div>';
    }

    // In Process Tile (tasks with task_action === 'in_process'; show expected completion date, editable)
    const inProcessList = document.getElementById('inProcessTileList');
    if (inProcessList) {
        const sortedInProcess = [...(inProcessTasks || [])]
            .sort((a, b) => {
                const aDate = a.expected_completion_date || a.due_date || a.next_due_date || '9999-12-31';
                const bDate = b.expected_completion_date || b.due_date || b.next_due_date || '9999-12-31';
                return aDate.localeCompare(bDate);
            });
        inProcessList.innerHTML = sortedInProcess.length > 0
            ? sortedInProcess.map(renderInProcessTileItem).join('')
            : '<div style="text-align: center; color: #999; padding: 15px; font-size: 13px;">No tasks in process</div>';
    }

    // Work Plan Tile (all tasks; list has scroll for >5)
    const workPlanList = document.getElementById('workPlanTileList');
    if (workPlanList) {
        const sortedWorkPlan = [...workPlanTasks]
            .sort((a, b) => (a.task_name || '').localeCompare(b.task_name || ''));

        workPlanList.innerHTML = sortedWorkPlan.length > 0
            ? sortedWorkPlan.map(renderTileTaskItem).join('')
            : '<div style="text-align: center; color: #999; padding: 15px; font-size: 13px;">No work plan tasks</div>';
    }
}

// Render overdue tasks in drilldown with date range filters (shows ALL overdue when range not set)
function renderOverdueDrilldown() {
    const data = getData();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Get from and to month values (optional - when not set, show ALL overdue tasks)
    const fromMonthValue = document.getElementById('drilldownFromMonth')?.value;
    const toMonthValue = document.getElementById('drilldownToMonth')?.value;

    let fromDate = null;
    let toDate = null;
    if (fromMonthValue && toMonthValue) {
        const [fromYear, fromMonth] = fromMonthValue.split('-').map(Number);
        const [toYear, toMonth] = toMonthValue.split('-').map(Number);
        fromDate = new Date(fromYear, fromMonth - 1, 1);
        fromDate.setHours(0, 0, 0, 0);
        toDate = new Date(toYear, toMonth, 0);
        toDate.setHours(23, 59, 59, 999);
    }

    // Filter overdue tasks (no date range filter when from/to not set = show all overdue)
    const overdueTasks = data.tasks.filter(task => {
        if (isTaskCompleted(task)) return false;
        if (task.task_action === 'in_process') return false;
        if (task.task_action === 'not_done') return false;
        if (task.removed_at) return false;

        const dueDate = task.due_date || task.next_due_date;
        if (!dueDate) return false;

        const dateParts = dueDate.split('-');
        const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
        taskDate.setHours(0, 0, 0, 0);

        // Task is overdue if due date is before today
        const isOverdue = taskDate < today;

        // Check if task date is within the from-to range (when range is set)
        const inRange = !fromDate || !toDate || (taskDate >= fromDate && taskDate <= toDate);

        // Check user permissions
        const isAssigned = currentUser.role === 'admin' ||
            task.assigned_to === currentUser.id ||
            (task.is_team_task && currentUser.role === 'admin');

        return isOverdue && inRange && isAssigned;
    });

    // Render the tasks in drilldown
    renderDrilldownTasks(overdueTasks);
}

// Reset drilldown date filters
function resetDrilldownDateFilters() {
    const p = getDashboardPeriod();
    document.getElementById('drilldownFromMonth').value = p.from;
    document.getElementById('drilldownToMonth').value = p.to;
    renderOverdueDrilldown();
}

function renderDrilldownTasks(tasks) {
    const container = document.getElementById('drilldownTasksList');
    if (!container) return;

    if (tasks.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #999; padding: 40px;">No tasks found.</p>';
        return;
    }

    // Sort tasks by due date
    tasks.sort((a, b) => {
        const dateA = a.due_date || a.next_due_date || '';
        const dateB = b.due_date || b.next_due_date || '';
        return dateA.localeCompare(dateB);
    });

    window.drilldownFilteredTasks = tasks;
    window.currentFilteredTasksForExport = tasks;

    // Use renderTaskItem with showActions=true for action buttons (Complete, Edit, Copy, Delete)
    container.innerHTML = tasks.map(task => renderTaskItem(task, true)).join('');
}

// Main renderDrilldown function - called when switching to drilldown tab
function renderDrilldown() {
    if (!drilldownContext) return;

    // Update header with title and count
    const header = document.getElementById('drilldownHeader');
    if (header) {
        const monthDisplay = drilldownContext.monthValue ?
            ` - ${formatMonthDisplay(drilldownContext.monthValue)}` : '';
        header.innerHTML = `
            <h3 style="margin: 0;">${drilldownContext.title}${monthDisplay}</h3>
            <p style="margin: 5px 0 0; opacity: 0.9;">Total Count: ${drilldownContext.count}</p>
        `;
    }

    // Call the filter function to display tasks
    if (drilldownContext.filterFunction) {
        drilldownContext.filterFunction();
    }
}

// Helper function to format month display
function formatMonthDisplay(monthValue) {
    if (!monthValue) return '';
    const [year, month] = monthValue.split('-');
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'];
    return `${monthNames[parseInt(month) - 1]} ${year}`;
}

// Render User Statistics Table (year,month for single month; or fromYear,fromMonth,toYear,toMonth for period)
function renderUserStatsTable(userStats, year, month, toYear, toMonth) {
    const tbody = document.getElementById('userStatsTableBody');
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'];

    if (!tbody) return;

    const usersWithTasks = Object.values(userStats).filter(u => u.totalAssigned > 0);
    const periodLabel = (toYear != null && toMonth != null && (year !== toYear || month !== toMonth))
        ? `${monthNames[month]} ${year} - ${monthNames[toMonth]} ${toYear}`
        : `${monthNames[month]} ${year}`;

    if (usersWithTasks.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="5" style="text-align: center; padding: 30px; color: #999;">
                    No tasks found for ${periodLabel}
                </td>
            </tr>
        `;
        return;
    }

    tbody.innerHTML = usersWithTasks.map(user => {
        const completedPercent = user.totalAssigned > 0
            ? ((user.totalCompleted / user.totalAssigned) * 100).toFixed(1)
            : '0.0';
        const pendingPercent = user.totalAssigned > 0
            ? ((user.totalPending / user.totalAssigned) * 100).toFixed(1)
            : '0.0';
        const needImprovementPercent = user.totalAssigned > 0
            ? ((user.totalNeedImprovement / user.totalAssigned) * 100).toFixed(1)
            : '0.0';

        return `
            <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="font-weight: 600; color: #333; padding: 15px;">${user.name}</td>
                <td style="text-align: center; padding: 15px;">
                    <span style="font-size: 18px; font-weight: 600; color: #667eea;">${user.totalAssigned}</span>
                </td>
                <td style="text-align: center; padding: 15px;">
                    <span style="font-size: 18px; font-weight: 600; color: #28a745;">${user.totalCompleted}</span>
                    <span style="font-size: 12px; color: #666; display: block; margin-top: 3px;">${completedPercent}%</span>
                </td>
                <td style="text-align: center; padding: 15px;">
                    <span style="font-size: 18px; font-weight: 600; color: #ffc107;">${user.totalPending}</span>
                    <span style="font-size: 12px; color: #666; display: block; margin-top: 3px;">${pendingPercent}%</span>
                </td>
                <td style="text-align: center; padding: 15px;">
                    <span style="font-size: 18px; font-weight: 600; color: #dc3545;">${user.totalNeedImprovement}</span>
                    <span style="font-size: 12px; color: #666; display: block; margin-top: 3px;">${needImprovementPercent}%</span>
                </td>
            </tr>
        `;
    }).join('');
}

// Render Detailed Statistics Table
function renderDetailedStatsTable(userStats) {
    const tbody = document.getElementById('detailedStatsTableBody');
    if (!tbody) return;

    const usersWithTasks = Object.values(userStats).filter(u => u.totalAssigned > 0);

    if (usersWithTasks.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="3" style="text-align: center; padding: 30px; color: #999;">
                    No statistics available
                </td>
            </tr>
        `;
        return;
    }

    // Calculate totals
    const totalAssigned = usersWithTasks.reduce((sum, u) => sum + u.totalAssigned, 0);
    const totalCompleted = usersWithTasks.reduce((sum, u) => sum + u.totalCompleted, 0);
    const totalPending = usersWithTasks.reduce((sum, u) => sum + u.totalPending, 0);
    const totalNeedImprovement = usersWithTasks.reduce((sum, u) => sum + u.totalNeedImprovement, 0);

    const completedPercent = totalAssigned > 0 ? ((totalCompleted / totalAssigned) * 100).toFixed(1) : '0.0';
    const pendingPercent = totalAssigned > 0 ? ((totalPending / totalAssigned) * 100).toFixed(1) : '0.0';
    const needImprovementPercent = totalAssigned > 0 ? ((totalNeedImprovement / totalAssigned) * 100).toFixed(1) : '0.0';

    tbody.innerHTML = `
        <tr style="background: #f9f9f9;">
            <td style="font-weight: 600; padding: 15px; color: #333;">Total Assigned Tasks</td>
            <td style="text-align: center; padding: 15px; font-size: 18px; font-weight: 600; color: #667eea;">${totalAssigned}</td>
            <td style="text-align: center; padding: 15px; font-size: 18px; font-weight: 600; color: #667eea;">100.0%</td>
        </tr>
        <tr>
            <td style="padding: 15px; color: #333;">Total Completed Tasks</td>
            <td style="text-align: center; padding: 15px; font-size: 16px; font-weight: 600; color: #28a745;">${totalCompleted}</td>
            <td style="text-align: center; padding: 15px; font-size: 16px; font-weight: 600; color: #28a745;">${completedPercent}%</td>
        </tr>
        <tr style="background: #f9f9f9;">
            <td style="padding: 15px; color: #333;">Total Pending Tasks</td>
            <td style="text-align: center; padding: 15px; font-size: 16px; font-weight: 600; color: #ffc107;">${totalPending}</td>
            <td style="text-align: center; padding: 15px; font-size: 16px; font-weight: 600; color: #ffc107;">${pendingPercent}%</td>
        </tr>
        <tr>
            <td style="padding: 15px; color: #333;">Total Need Improvement Tasks</td>
            <td style="text-align: center; padding: 15px; font-size: 16px; font-weight: 600; color: #dc3545;">${totalNeedImprovement}</td>
            <td style="text-align: center; padding: 15px; font-size: 16px; font-weight: 600; color: #dc3545;">${needImprovementPercent}%</td>
        </tr>
    `;
}

// Drilldown
function renderDrilldown() {
    console.log('renderDrilldown called, context:', drilldownContext);

    if (!drilldownContext) {
        document.getElementById('drilldownTitle').textContent = 'Drilldown';
        document.getElementById('drilldownHeader').innerHTML = '<p style="margin: 0; font-size: 16px;">Please click on a dashboard tile to view details.</p>';
        document.getElementById('drilldownTasksList').innerHTML = '';
        return;
    }

    // Sync Tasks tab period with dashboard
    const p = getDashboardPeriod();
    const fromEl = document.getElementById('filterTaskMonthFrom');
    const toEl = document.getElementById('filterTaskMonthTo');
    if (fromEl) fromEl.value = p.from;
    if (toEl) toEl.value = p.to;

    // Update title and header
    document.getElementById('drilldownTitle').textContent = drilldownContext.title;

    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'];
    let monthText = '';
    if (drilldownContext.periodFrom && drilldownContext.periodTo) {
        const [fy, fm] = drilldownContext.periodFrom.split('-').map(Number);
        const [ty, tm] = drilldownContext.periodTo.split('-').map(Number);
        monthText = (fy !== ty || fm !== tm)
            ? ` - ${monthNames[fm - 1]} ${fy} to ${monthNames[tm - 1]} ${ty}`
            : ` - ${monthNames[fm - 1]} ${fy}`;
    } else if (drilldownContext.monthValue) {
        const [year, month] = drilldownContext.monthValue.split('-');
        monthText = ` - ${monthNames[parseInt(month) - 1]} ${year}`;
    }

    document.getElementById('drilldownHeader').innerHTML = `
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div>
                <h3 style="margin: 0; font-size: 24px; font-weight: 600;">${drilldownContext.title}${monthText}</h3>
                <p style="margin: 5px 0 0 0; font-size: 18px; opacity: 0.9;">Total Count: <strong>${drilldownContext.count}</strong></p>
            </div>
        </div>
    `;

    // Get filtered tasks based on type
    let filteredTasks = [];
    let headerHtml = '';

    const data = getData();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayStr = formatDateString(today);

    // Set period filter values from dashboard period
    let filterMonth = null, filterYear = null, toMonth = null, toYear = null;
    if (drilldownContext.periodFrom) {
        const [fy, fm] = drilldownContext.periodFrom.split('-').map(Number);
        filterYear = fy;
        filterMonth = fm - 1;
    }
    if (drilldownContext.periodTo) {
        const [ty, tm] = drilldownContext.periodTo.split('-').map(Number);
        toYear = ty;
        toMonth = tm - 1;
    }

    switch (drilldownContext.type) {
        case 'today':
            const todayDateStr = drilldownContext.dateStr || todayStr;
            filteredTasks = getTodayTasks(todayDateStr);
            headerHtml = `<div class="date-filter-header" style="margin-bottom: 20px; padding: 15px; background: #e3f2fd; border-radius: 8px;">
                <h3 style="margin: 0;">Today's Tasks - ${formatDateDisplay(todayDateStr)} (${filteredTasks.length})</h3>
            </div>`;
            break;
        case 'overdue':
            if (drilldownContext.filterFunction) {
                drilldownContext.filterFunction();
                return;
            }
            filteredTasks = [];
            headerHtml = '';
            break;
        case 'completed':
            filteredTasks = getCompletedTasks(filterMonth, filterYear);
            headerHtml = '';
            break;
        case 'pending':
            if (filterMonth !== null && filterYear !== null) {
                filteredTasks = getPendingTasksByPeriod(filterMonth, filterYear, toMonth, toYear);
                const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                    'July', 'August', 'September', 'October', 'November', 'December'];
                const monthName = monthNames[filterMonth];
                headerHtml = `<div class="date-filter-header" style="margin-bottom: 20px; padding: 15px; background: #d1ecf1; border-radius: 8px;">
                    <h3 style="margin: 0;">Pending Tasks - ${monthName} ${filterYear} (${filteredTasks.length})</h3>
                </div>`;
            } else {
                filteredTasks = getPendingTasks(filterMonth, filterYear);
                headerHtml = `<div class="date-filter-header" style="margin-bottom: 20px; padding: 15px; background: #d1ecf1; border-radius: 8px;">
                    <h3 style="margin: 0;">Pending Tasks - All Periods (${filteredTasks.length})</h3>
                </div>`;
            }
            break;
        case 'no_due_date':
            filteredTasks = getNoDueDateTasks();
            headerHtml = '';
            break;
        case 'work_plan':
            filteredTasks = getWorkPlanTasks(filterMonth, filterYear, toMonth, toYear);
            const monthNamesWP = ['January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'];
            const monthNameWP = filterMonth !== null ? monthNamesWP[filterMonth] : '';
            const yearTextWP = filterYear !== null ? filterYear : '';
            headerHtml = `<div class="date-filter-header" style="margin-bottom: 20px; padding: 15px; background: #e1bee7; border-radius: 8px;">
                <h3 style="margin: 0;">Work Plan Tasks - ${monthNameWP} ${yearTextWP} (${filteredTasks.length})</h3>
            </div>`;
            break;
        case 'audit_point':
            filteredTasks = getAuditPointTasks(filterMonth, filterYear, toMonth, toYear);
            const monthNamesAP = ['January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'];
            const monthNameAP = filterMonth !== null ? monthNamesAP[filterMonth] : '';
            const yearTextAP = filterYear !== null ? filterYear : '';
            headerHtml = `<div class="date-filter-header" style="margin-bottom: 20px; padding: 15px; background: #ffd54f; border-radius: 8px;">
                <h3 style="margin: 0;">Audit Point Tasks - ${monthNameAP} ${yearTextAP} (${filteredTasks.length})</h3>
            </div>`;
            break;
        case 'admin_review':
            filteredTasks = getAdminReviewTasks(filterMonth, filterYear);
            headerHtml = `<div class="date-filter-header" style="margin-bottom: 20px; padding: 15px; background: #fff3cd; border-radius: 8px;">
                <h3 style="margin: 0;">Pending Admin Review (${filteredTasks.length})</h3>
            </div>`;
            break;
        default:
            console.warn('Unknown drilldown type:', drilldownContext.type);
            filteredTasks = [];
            headerHtml = '';
            break;
    }

    // Store for export
    window.drilldownFilteredTasks = filteredTasks;
    window.currentFilteredTasksForExport = filteredTasks;

    // Debug: Log filtered tasks count
    console.log('Drilldown - Type:', drilldownContext.type, 'Filtered Tasks:', filteredTasks.length);

    // Render tasks
    const tasksHtml = filteredTasks.length > 0
        ? filteredTasks.map(task => renderTaskItem(task, true)).join('')
        : '<p style="text-align: center; color: #999; padding: 20px;">No tasks found</p>';

    const drilldownTasksListElement = document.getElementById('drilldownTasksList');
    if (drilldownTasksListElement) {
        drilldownTasksListElement.innerHTML = headerHtml + tasksHtml;
    } else {
        console.error('drilldownTasksList element not found!');
    }
}

// Helper function to check if task is completed (either completed or completed_need_improvement)
function isTaskCompleted(task) {
    return task.task_action === 'completed' || task.task_action === 'completed_need_improvement';
}

// Task Mgmt: show only one recurring task instance per series (the "current" one)
function reduceRecurringTasksForTaskMgmt(tasks, statusFilter) {
    const recurring = tasks.filter(t => t.task_type === 'recurring');
    if (recurring.length === 0) return tasks;

    const others = tasks.filter(t => t.task_type !== 'recurring');

    const dueTime = (t) => {
        const dueStr = t.due_date || t.next_due_date;
        if (!dueStr) return null;
        const [y, m, d] = dueStr.split('-').map(Number);
        return new Date(y, m - 1, d).getTime();
    };

    const completedTime = (t) => {
        const v = t.completed_at || t.completion_date;
        if (!v) return null;
        // completed_at is ISO; completion_date is YYYY-MM-DD
        return t.completed_at ? new Date(v).getTime() : dueTime(t);
    };

    let pool = recurring;
    // When user didn't filter by status, Task Mgmt should focus on "current work"
    if (!statusFilter) {
        pool = recurring.filter(t => t.task_action === 'not_completed' || t.task_action === 'in_process');
        if (pool.length === 0) pool = recurring;
    }

    const pickBetter = (a, b) => {
        const completedMode = statusFilter === 'completed' || statusFilter === 'completed_need_improvement';
        if (completedMode) {
            const ta = completedTime(a);
            const tb = completedTime(b);
            if (ta == null) return false;
            if (tb == null) return true;
            // For completed: keep the latest
            return ta > tb;
        }

        const da = dueTime(a);
        const db = dueTime(b);
        if (da == null) return false;
        if (db == null) return true;
        // For active/not-done/in-process: keep the earliest due date
        return da < db;
    };

    const bySeriesKey = new Map();
    pool.forEach(t => {
        const key = [t.task_name, t.assigned_to, t.location_id, t.frequency].join('::');
        const prev = bySeriesKey.get(key);
        if (!prev) {
            bySeriesKey.set(key, t);
        } else {
            if (pickBetter(t, prev)) bySeriesKey.set(key, t);
        }
    });

    return others.concat(Array.from(bySeriesKey.values()));
}

// Helper: calculate recurrence number for a recurring task within its series
function getTaskRecurrenceNumber(task, data) {
    try {
        if (!task || task.task_type !== 'recurring') return null;
        const dueStr = task.due_date || task.next_due_date;
        if (!dueStr) return null;

        const seriesTasks = (data.tasks || []).filter(t =>
            t.task_type === 'recurring' &&
            t.task_name === task.task_name &&
            t.assigned_to === task.assigned_to &&
            t.location_id === task.location_id &&
            t.frequency === task.frequency
        );

        if (seriesTasks.length === 0) return null;

        seriesTasks.sort((a, b) => {
            const ad = a.due_date || a.next_due_date || '9999-12-31';
            const bd = b.due_date || b.next_due_date || '9999-12-31';
            if (ad !== bd) return ad.localeCompare(bd);
            const ac = a.created_at || '';
            const bc = b.created_at || '';
            if (ac !== bc) return ac.localeCompare(bc);
            return (a.id || 0) - (b.id || 0);
        });

        let index = null;
        seriesTasks.forEach((t, i) => {
            if (t.id === task.id && index === null) {
                index = i;
            }
        });

        return index != null ? index + 1 : null;
    } catch (e) {
        console.warn('Failed to compute recurrence number', e);
        return null;
    }
}

// Helper functions to get filtered tasks
function getTodayTasks(dateStr) {
    const data = getData();
    return data.tasks.filter(task => {
        if (task.task_type === 'without_due_date') return false;
        if (task.removed_at) return false; // Exclude removed tasks
        if (task.task_action === 'not_done') return false; // Exclude Not Done
        const dueDate = task.due_date || task.next_due_date;
        return dueDate === dateStr &&
            !isTaskCompleted(task) &&
            (task.assigned_to === currentUser.id ||
                (task.is_team_task && currentUser.role === 'admin') ||
                currentUser.role === 'admin');
    });
}

function getOverdueTasks(filterMonth, filterYear) {
    const data = getData();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    return data.tasks.filter(task => {
        if (task.task_type === 'without_due_date') return false;
        if (task.removed_at) return false; // Exclude removed tasks
        if (task.task_action === 'in_process') return false; // In Process should not appear in Overdue tab
        if (task.task_action === 'not_done') return false; // Not Done is closed
        const dueDate = task.due_date || task.next_due_date;
        if (!dueDate || isTaskCompleted(task)) return false;

        const dateParts = dueDate.split('-');
        const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
        taskDate.setHours(0, 0, 0, 0);

        const isOverdue = taskDate < today;

        if (filterMonth !== null && filterYear !== null) {
            const taskMonth = taskDate.getMonth();
            const taskYear = taskDate.getFullYear();
            if (taskMonth !== filterMonth || taskYear !== filterYear) {
                return false;
            }
        }

        return isOverdue &&
            (task.assigned_to === currentUser.id ||
                (task.is_team_task && currentUser.role === 'admin') ||
                currentUser.role === 'admin');
    });
}

function getCompletedTasks(filterMonth, filterYear) {
    const data = getData();
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

    // Completed tile: only task_action === 'completed' (exclude Need Improvement from dashboard/pendency)
    return data.tasks.filter(task => {
        if (task.task_action !== 'completed') return false;
        if (!task.completed_at) return false;

        const completedDate = new Date(task.completed_at);
        if (completedDate < thirtyDaysAgo) return false;

        return task.assigned_to === currentUser.id ||
            (task.is_team_task && currentUser.role === 'admin') ||
            currentUser.role === 'admin';
    });
}

function getPendingTasks(filterMonth, filterYear) {
    const data = getData();

    // Show ALL pending tasks from the beginning, regardless of selected period
    // Used for dashboard count - shows total pending tasks
    return data.tasks.filter(task => {
        if (task.task_type === 'without_due_date') return false;
        if (task.removed_at) return false; // Exclude removed tasks
        const dueDate = task.due_date || task.next_due_date;
        if (!dueDate || isTaskCompleted(task)) return false;
        if (task.task_action === 'not_done') return false; // Not Done is closed

        // Return all pending tasks regardless of due date period
        return (task.assigned_to === currentUser.id ||
            (task.is_team_task && currentUser.role === 'admin') ||
            currentUser.role === 'admin');
    });
}

function getPendingTasksByPeriod(filterMonth, filterYear, toMonth, toYear) {
    const data = getData();
    const fromDate = new Date(filterYear, filterMonth, 1);
    const toDate = (toMonth != null && toYear != null)
        ? new Date(toYear, toMonth + 1, 0) : new Date(filterYear, filterMonth + 1, 0);
    toDate.setHours(23, 59, 59, 999);

    return data.tasks.filter(task => {
        if (task.task_type === 'without_due_date') return false;
        if (task.removed_at) return false;
        const dueDate = task.due_date || task.next_due_date;
        if (!dueDate || isTaskCompleted(task)) return false;
        if (task.task_action === 'not_done') return false; // Not Done is closed

        const dateParts = dueDate.split('-');
        const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
        taskDate.setHours(0, 0, 0, 0);

        return taskDate >= fromDate && taskDate <= toDate &&
            (task.assigned_to === currentUser.id ||
                (task.is_team_task && currentUser.role === 'admin') ||
                currentUser.role === 'admin');
    });
}

function getNoDueDateTasks() {
    const data = getData();
    return data.tasks.filter(task => {
        return task.task_type === 'without_due_date' &&
            !isTaskCompleted(task) &&
            task.task_action !== 'not_done' &&
            !task.removed_at && // Exclude removed tasks
            (task.assigned_to === currentUser.id ||
                (task.is_team_task && currentUser.role === 'admin') ||
                currentUser.role === 'admin');
    });
}

function getWorkPlanTasks(filterMonth, filterYear, toMonth, toYear) {
    const data = getData();
    const today = new Date();
    const fm = filterMonth ?? today.getMonth();
    const fy = filterYear ?? today.getFullYear();
    const fromDate = new Date(fy, fm, 1);
    const toDate = (toMonth != null && toYear != null)
        ? new Date(toYear, toMonth + 1, 0) : new Date(fy, fm + 1, 0);
    toDate.setHours(23, 59, 59, 999);

    return data.tasks.filter(task => {
        if (task.task_type !== 'work_plan' || isTaskCompleted(task) || task.removed_at) return false;

        const dueDate = task.due_date || task.next_due_date;
        if (dueDate) {
            const dateParts = dueDate.split('-');
            const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
            taskDate.setHours(0, 0, 0, 0);
            return taskDate >= fromDate && taskDate <= toDate &&
                (task.assigned_to === currentUser.id ||
                    (task.is_team_task && currentUser.role === 'admin') ||
                    currentUser.role === 'admin');
        }
        return task.assigned_to === currentUser.id ||
            (task.is_team_task && currentUser.role === 'admin') ||
            currentUser.role === 'admin';
    });
}

function getAuditPointTasks(filterMonth, filterYear, toMonth, toYear) {
    const data = getData();
    const today = new Date();
    const fm = filterMonth ?? today.getMonth();
    const fy = filterYear ?? today.getFullYear();
    const fromDate = new Date(fy, fm, 1);
    const toDate = (toMonth != null && toYear != null)
        ? new Date(toYear, toMonth + 1, 0) : new Date(fy, fm + 1, 0);
    toDate.setHours(23, 59, 59, 999);

    return data.tasks.filter(task => {
        if (task.task_type !== 'audit_point' || isTaskCompleted(task) || task.removed_at) return false;

        const dueDate = task.due_date || task.next_due_date;
        if (dueDate) {
            const dateParts = dueDate.split('-');
            const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
            taskDate.setHours(0, 0, 0, 0);
            return taskDate >= fromDate && taskDate <= toDate &&
                (task.assigned_to === currentUser.id ||
                    (task.is_team_task && currentUser.role === 'admin') ||
                    currentUser.role === 'admin');
        }
        return task.assigned_to === currentUser.id ||
            (task.is_team_task && currentUser.role === 'admin') ||
            currentUser.role === 'admin';
    });
}

function getAdminReviewTasks(filterMonth, filterYear) {
    const data = getData();
    return data.tasks.filter(task => {
        if (!isTaskCompleted(task) || task.admin_finalized) return false;

        if (filterMonth !== null && filterYear !== null) {
            const dueDate = task.due_date || task.next_due_date;
            if (dueDate) {
                const dateParts = dueDate.split('-');
                const taskMonth = parseInt(dateParts[1]) - 1;
                const taskYear = parseInt(dateParts[0]);
                if (taskMonth !== filterMonth || taskYear !== filterYear) {
                    return false;
                }
            } else {
                return false;
            }
        }

        return task.assigned_to === currentUser.id ||
            (task.is_team_task && currentUser.role === 'admin') ||
            currentUser.role === 'admin';
    });
}

function exportDrilldownCSV() {
    if (!drilldownContext || !window.drilldownFilteredTasks || window.drilldownFilteredTasks.length === 0) {
        alert('No tasks to export. Please click on a dashboard tile first.');
        return;
    }

    const data = getData();
    const tasksToExport = window.drilldownFilteredTasks;

    const headers = [
        'Task Number',
        'Task Name',
        'Description',
        'Assigned to (Name)',
        'Location',
        'Task Type',
        'Frequency',
        'Due date calculation type',
        'Due day of month',
        'Recurrence Type',
        'Start date (for due date calculation)',
        'Estimated Minutes',
        'Team task (true/false)',
        'Status',
        'Due Date',
        'Expected Date of Completion',
        'Completion Date',
        'Completion Remark'
    ];
    const rows = tasksToExport.map(task => {
        const user = data.users.find(u => u.id === task.assigned_to);
        const location = data.locations.find(l => l.id === task.location_id);

        // Format start date as DD-MM-YYYY
        let startDateStr = '';
        if (task.start_date) {
            startDateStr = formatDateDisplay(task.start_date);
        }

        // Format due date as DD-MM-YYYY
        let dueDateStr = '';
        const dueDate = task.due_date || task.next_due_date;
        if (dueDate) {
            dueDateStr = formatDateDisplay(dueDate);
        } else if (task.task_type === 'without_due_date') {
            dueDateStr = 'No Due Date';
        }

        // Format expected completion date as DD-MM-YYYY (when set)
        let expectedCompletionStr = '';
        if (task.expected_completion_date) {
            expectedCompletionStr = formatDateDisplay(task.expected_completion_date);
        }

        // Format completion date as DD-MM-YYYY
        let completionDateStr = '';
        if (task.completion_date) {
            completionDateStr = formatDateDisplay(task.completion_date);
        } else if (task.completed_at) {
            completionDateStr = formatDateDisplay(task.completed_at);
        }

        // Format status
        let statusStr = '';
        if (task.task_action === 'completed') {
            statusStr = 'Completed';
        } else if (task.task_action === 'completed_need_improvement') {
            statusStr = 'Needs Improvement';
        } else if (task.task_action === 'in_process') {
            statusStr = 'In Process';
        } else if (task.task_action === 'not_done') {
            statusStr = 'Not Done';
        } else {
            statusStr = 'Pending';
        }

        return [
            task.task_number != null ? String(task.task_number) : '',
            task.task_name || '',
            task.description || '',
            user ? user.name : '',
            location ? location.name : '',
            task.task_type || '',
            task.frequency || '',
            task.due_date_type || '',
            task.due_day || '',
            task.recurrence_type || '',
            startDateStr,
            task.est_minutes || '',
            task.is_team_task ? 'true' : 'false',
            statusStr,
            dueDateStr,
            expectedCompletionStr,
            completionDateStr,
            task.comment || ''
        ];
    });

    const csv = [
        headers.map(escapeCSV).join(','),
        ...rows.map(r => r.map(escapeCSV).join(','))
    ].join('\n');

    const filename = `${drilldownContext.type}-${formatDateString(new Date())}.csv`;
    downloadFile(csv, filename, 'text/csv');
}

function exportDrilldownJSON() {
    // JSON export always exports ALL data (complete backup), not just filtered drilldown data
    // This ensures full backup/restore capability across browsers with all past and future data
    exportAllData();
}

// Tasks
function renderTasks(skipAutoFilter = false) {
    const period = getDashboardPeriod();
    const fromEl = document.getElementById('filterTaskMonthFrom');
    const toEl = document.getElementById('filterTaskMonthTo');
    if (fromEl && !fromEl.value) fromEl.value = period.from;
    if (toEl && !toEl.value) toEl.value = period.to;
    if (!skipAutoFilter) filterTasks();
}

function filterTasks() {
    const data = getData();
    const statusFilter = document.getElementById('filterStatus').value;
    const typeFilter = document.getElementById('filterType').value;
    const teamFilter = document.getElementById('filterTeam').value;
    const searchFilter = document.getElementById('searchTasks').value.toLowerCase();
    const fromVal = document.getElementById('filterTaskMonthFrom')?.value;
    const toVal = document.getElementById('filterTaskMonthTo')?.value;

    let fromDate = null, toDate = null;
    if (fromVal && toVal) {
        const [fy, fm] = fromVal.split('-').map(Number);
        const [ty, tm] = toVal.split('-').map(Number);
        fromDate = new Date(fy, fm - 1, 1);
        toDate = new Date(ty, tm, 0);
        toDate.setHours(23, 59, 59, 999);
    }

    let filteredTasks = data.tasks.filter(task => {
        const isAssigned = task.assigned_to === currentUser.id ||
            (task.is_team_task && currentUser.role === 'admin') ||
            currentUser.role === 'admin';

        if (!isAssigned) return false;
        if (statusFilter && task.task_action !== statusFilter) return false;
        if (typeFilter && task.task_type !== typeFilter) return false;
        if (teamFilter === 'self' && (task.is_team_task || task.assigned_to !== currentUser.id)) return false;
        if (teamFilter === 'team' && !task.is_team_task) return false;
        if (searchFilter && !task.task_name.toLowerCase().includes(searchFilter) &&
            !(task.description || '').toLowerCase().includes(searchFilter)) return false;

        // From/To month filter is not required for recurring tasks (Task Mgmt shows master data).
        if (fromDate && toDate && task.task_type !== 'recurring') {
            const dueDate = task.due_date || task.next_due_date;
            // Done (completed or need improvement): only visible if completion/due date falls in period
            if (isTaskCompleted(task)) {
                const refDateStr = task.completed_at ? task.completed_at.split('T')[0] : dueDate;
                if (!refDateStr) return false;
                const [y, m, d] = refDateStr.split('-').map(Number);
                const taskDate = new Date(y, m - 1, d);
                if (taskDate < fromDate || taskDate > toDate) return false;
            } else {
                // Pending, In Process, Overdue: only visible if due date in period (or without_due_date)
                if (dueDate) {
                    const [y, m, d] = dueDate.split('-').map(Number);
                    const taskDate = new Date(y, m - 1, d);
                    if (taskDate < fromDate || taskDate > toDate) return false;
                } else {
                    if (task.task_type !== 'without_due_date') return false;
                }
            }
        }
        return true;
    });

    // Task Mgmt enhancement: reduce recurring tasks to a single "current" row per series
    filteredTasks = reduceRecurringTasksForTaskMgmt(filteredTasks, statusFilter);

    // Clear any date-specific filters
    const tasksListElement = document.getElementById('tasksList');
    if (tasksListElement.querySelector('.date-filter-header')) {
        tasksListElement.querySelector('.date-filter-header').remove();
    }

    // Sort by priority and due date (completed tasks last)
    filteredTasks.sort((a, b) => {
        // Completed tasks go to the end
        if (a.task_action === 'completed' && b.task_action !== 'completed') return 1;
        if (a.task_action !== 'completed' && b.task_action === 'completed') return -1;

        const priorityOrder = { high: 1, medium: 2, low: 3 };
        const aPriority = priorityOrder[a.priority] || 4;
        const bPriority = priorityOrder[b.priority] || 4;
        if (aPriority !== bPriority) return aPriority - bPriority;

        const aDate = a.due_date || a.next_due_date || '9999-12-31';
        const bDate = b.due_date || b.next_due_date || '9999-12-31';
        return aDate.localeCompare(bDate);
    });

    // Store filtered tasks for export
    window.currentFilteredTasksForExport = filteredTasks;

    const tasksHtml = filteredTasks.length > 0
        ? filteredTasks.map(task => renderTaskItem(task, true)).join('')
        : '<p style="text-align: center; color: #999; padding: 20px;">No tasks found</p>';

    document.getElementById('tasksList').innerHTML = tasksHtml;
}

function filterTasksByDate(dateStr) {
    // Clear any existing filter headers first
    const tasksListElement = document.getElementById('tasksList');
    if (tasksListElement) {
        // Remove any existing date-filter-header
        const existingHeader = tasksListElement.querySelector('.date-filter-header');
        if (existingHeader) {
            existingHeader.remove();
        }
        tasksListElement.innerHTML = '';
    }

    const data = getData();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const dateTasks = data.tasks.filter(task => {
        // Exclude tasks without due date
        if (task.task_type === 'without_due_date') return false;
        if (task.removed_at) return false; // Exclude removed tasks
        const dueDate = task.due_date || task.next_due_date;
        return dueDate === dateStr &&
            !isTaskCompleted(task) &&
            (task.assigned_to === currentUser.id ||
                (task.is_team_task && currentUser.role === 'admin') ||
                currentUser.role === 'admin');
    });

    const dateParts = dateStr.split('-');
    const displayDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));

    // Store filtered tasks for export
    window.currentFilteredTasksForExport = dateTasks;

    const tasksHtml = dateTasks.length > 0
        ? dateTasks.map(task => renderTaskItem(task, true)).join('')
        : `<p style="text-align: center; color: #999; padding: 20px;">No tasks for ${formatDateDisplay(displayDate)}</p>`;

    document.getElementById('tasksList').innerHTML = `
        <div class="date-filter-header" style="margin-bottom: 20px; padding: 15px; background: #e3f2fd; border-radius: 8px; display: flex; justify-content: space-between; align-items: center;">
            <h3 style="margin: 0;">Today's Tasks - ${formatDateDisplay(displayDate)}</h3>
            <button class="btn btn-secondary" onclick="filterTasks()" style="padding: 5px 15px; font-size: 12px;">Clear Filter</button>
        </div>
        ${tasksHtml}
    `;
}

function filterTasksByOverdue() {
    // Clear any existing filter headers first
    const tasksListElement = document.getElementById('tasksList');
    if (tasksListElement) {
        // Remove any existing date-filter-header
        const existingHeader = tasksListElement.querySelector('.date-filter-header');
        if (existingHeader) {
            existingHeader.remove();
        }
        tasksListElement.innerHTML = '';
    }

    const data = getData();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const fromVal = document.getElementById('filterTaskMonthFrom')?.value;
    const toVal = document.getElementById('filterTaskMonthTo')?.value;
    let fromDate = null, toDate = null;
    if (fromVal && toVal) {
        const [fy, fm] = fromVal.split('-').map(Number);
        const [ty, tm] = toVal.split('-').map(Number);
        fromDate = new Date(fy, fm - 1, 1);
        toDate = new Date(ty, tm, 0);
        toDate.setHours(23, 59, 59, 999);
    }

    const overdueTasks = data.tasks.filter(task => {
        if (task.task_type === 'without_due_date') return false;
        if (task.removed_at) return false;
        if (task.task_action === 'in_process') return false;
        if (task.task_action === 'not_done') return false;
        const dueDate = task.due_date || task.next_due_date;
        if (!dueDate || isTaskCompleted(task)) return false;

        const dateParts = dueDate.split('-');
        const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
        taskDate.setHours(0, 0, 0, 0);

        const isOverdue = taskDate < today;
        if (fromDate && toDate && (taskDate < fromDate || taskDate > toDate)) return false;

        return isOverdue &&
            (task.assigned_to === currentUser.id ||
                (task.is_team_task && currentUser.role === 'admin') ||
                currentUser.role === 'admin');
    });

    // Store filtered tasks for export
    window.currentFilteredTasksForExport = overdueTasks;

    const tasksHtml = overdueTasks.length > 0
        ? overdueTasks.map(task => renderTaskItem(task, true)).join('')
        : '<p style="text-align: center; color: #999; padding: 20px;">No overdue tasks</p>';

    document.getElementById('tasksList').innerHTML = `
        <div class="date-filter-header" style="margin-bottom: 20px; padding: 15px; background: #f8d7da; border-radius: 8px; display: flex; justify-content: space-between; align-items: center;">
            <h3 style="margin: 0;">Overdue Tasks (${overdueTasks.length})</h3>
            <button class="btn btn-secondary" onclick="filterTasks()" style="padding: 5px 15px; font-size: 12px;">Clear Filter</button>
        </div>
        ${tasksHtml}
    `;
}

function filterTasksByPendingMonth() {
    // Clear any existing filter headers first
    const tasksListElement = document.getElementById('tasksList');
    if (tasksListElement) {
        // Remove any existing date-filter-header
        const existingHeader = tasksListElement.querySelector('.date-filter-header');
        if (existingHeader) {
            existingHeader.remove();
        }
        tasksListElement.innerHTML = '';
    }

    const data = getData();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const fromVal = document.getElementById('filterTaskMonthFrom')?.value;
    const toVal = document.getElementById('filterTaskMonthTo')?.value;

    let pendingTasks;
    let headerText;
    if (fromVal && toVal) {
        const [fy, fm] = fromVal.split('-').map(Number);
        const [ty, tm] = toVal.split('-').map(Number);
        pendingTasks = getPendingTasksByPeriod(fm - 1, fy, tm - 1, ty);
        const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        headerText = `Pending Tasks - ${monthNames[fm - 1]} ${fy} to ${monthNames[tm - 1]} ${ty} (${pendingTasks.length})`;
    } else {
        pendingTasks = getPendingTasks(null, null);
        headerText = `Pending Tasks - All Periods (${pendingTasks.length})`;
    }

    // Store filtered tasks for export
    window.currentFilteredTasksForExport = pendingTasks;

    const tasksHtml = pendingTasks.length > 0
        ? pendingTasks.map(task => renderTaskItem(task, true)).join('')
        : '<p style="text-align: center; color: #999; padding: 20px;">No pending tasks</p>';

    document.getElementById('tasksList').innerHTML = `
        <div class="date-filter-header" style="margin-bottom: 20px; padding: 15px; background: #d1ecf1; border-radius: 8px; display: flex; justify-content: space-between; align-items: center;">
            <h3 style="margin: 0;">${headerText}</h3>
            <button class="btn btn-secondary" onclick="filterTasks()" style="padding: 5px 15px; font-size: 12px;">Clear Filter</button>
        </div>
        ${tasksHtml}
    `;
}

function renderTaskItem(task, showActions = false) {
    const data = getData();
    const assignedUser = data.users.find(u => u.id === task.assigned_to);
    const location = data.locations.find(l => l.id === task.location_id);
    const segregation = data.segregationTypes.find(s => s.id === task.segregation_type_id);

    const priorityBadge = task.priority
        ? `<span class="badge badge-${task.priority}">${task.priority.toUpperCase()}</span>`
        : '';

    let statusBadge = '';
    if (task.removed_at) {
        statusBadge = '<span class="badge badge-secondary">Removed</span>';
    } else if (task.task_action === 'completed') {
        if (task.admin_finalized) {
            statusBadge = '<span class="badge badge-completed">Finalized</span>';
        } else if (currentUser.role === 'admin') {
            statusBadge = '<span class="badge badge-warning">Pending Review</span>';
        } else {
            statusBadge = '<span class="badge badge-completed">Completed</span>';
        }
    } else if (task.task_action === 'completed_need_improvement') {
        if (task.admin_finalized) {
            statusBadge = '<span class="badge badge-warning">Needs Improvement - Finalized</span>';
        } else if (currentUser.role === 'admin') {
            statusBadge = '<span class="badge badge-warning">Needs Improvement - Pending Review</span>';
        } else {
            statusBadge = '<span class="badge badge-warning">Needs Improvement</span>';
        }
    } else if (task.task_action === 'in_process') {
        statusBadge = '<span class="badge badge-info">In Process</span>';
    } else if (task.task_action === 'not_done') {
        statusBadge = '<span class="badge badge-not-done">Not Done</span>';
    } else {
        // Check if task was previously rejected
        if (task.rejected_at) {
            statusBadge = '<span class="badge badge-pending">Rejected - Pending</span>';
        } else {
            statusBadge = '<span class="badge badge-pending">Pending</span>';
        }
    }

    let typeBadge = '';
    if (task.task_type === 'recurring') {
        typeBadge = `<span class="badge badge-recurring">Recurring${task.frequency ? ' - ' + task.frequency.charAt(0).toUpperCase() + task.frequency.slice(1) : ''}</span>`;
    } else if (task.task_type === 'without_due_date') {
        typeBadge = '<span class="badge badge-onetime">Without Due Date</span>';
    } else if (task.task_type === 'work_plan') {
        typeBadge = '<span class="badge badge-onetime">Work Plan</span>';
    } else if (task.task_type === 'audit_point') {
        typeBadge = '<span class="badge badge-onetime">Audit Point</span>';
    } else {
        typeBadge = '<span class="badge badge-onetime">One Time</span>';
    }

    const dueDate = task.due_date || task.next_due_date;
    let dueDateStr = 'Not set';
    if (task.task_type === 'without_due_date') {
        dueDateStr = 'No Due Date';
    } else if (dueDate) {
        dueDateStr = formatDateDisplay(dueDate);
    }
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayStr = formatDateString(today);

    // Calculate ageing (days overdue) for overdue tasks
    let ageingStr = '';
    if (!isTaskCompleted(task) && dueDate && task.task_type !== 'without_due_date') {
        const dateParts = dueDate.split('-');
        const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
        taskDate.setHours(0, 0, 0, 0);
        if (taskDate < today) {
            const diffTime = today - taskDate;
            const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
            ageingStr = ` (Ageing: ${diffDays} day${diffDays !== 1 ? 's' : ''})`;
        }
    }

    // Determine task card class based on status
    let cardClass = 'task-item';
    if (isTaskCompleted(task)) {
        if (task.task_action === 'completed_need_improvement') {
            cardClass += ' task-card-needs-improvement';
        } else {
            cardClass += ' task-card-completed';
        }
    } else if (dueDate) {
        const dateParts = dueDate.split('-');
        const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
        taskDate.setHours(0, 0, 0, 0);
        if (taskDate < today) {
            cardClass += ' task-card-overdue';
        } else {
            cardClass += ' task-card-not-due';
        }
    } else {
        cardClass += ' task-card-other';
    }

    let actionsHtml = '';
    if (showActions && !isTaskCompleted(task) && task.task_action !== 'not_done' && !task.removed_at) {
        actionsHtml = `
        <div class="task-actions">
            <button class="btn btn-success" onclick="completeTaskWithRemark(${task.id})">Complete</button>
            <button class="btn btn-warning" onclick="removeTaskWithoutCompletion(${task.id})" style="background: #ff9800; color: white;">Remove</button>
            <button class="btn btn-secondary" onclick="markTaskNotDone(${task.id})">Task Not Done</button>
            ${task.task_type === 'recurring' && !task.recurrence_stopped ? `<button class="btn btn-secondary" onclick="stopRecurringTask(${task.id})">Stop Recurring</button>` : ''}
            <button class="btn btn-primary" onclick="editTask(${task.id})">Edit</button>
            ${currentUser.role === 'admin' ? `<button class="btn btn-danger" onclick="deleteTask(${task.id})">Delete</button>` : ''}
        </div>
        `;
    } else if (showActions && isTaskCompleted(task)) {
        if (currentUser.role === 'admin' && !task.admin_finalized) {
            actionsHtml = `
        <div class="task-actions">
            <button class="btn btn-success" onclick="acceptTaskCompletion(${task.id})">✓ Accept</button>
            <button class="btn btn-danger" onclick="rejectTaskCompletion(${task.id})">✗ Reject</button>
            <button class="btn btn-primary" onclick="editTask(${task.id})">Edit</button>
            <button class="btn btn-info" onclick="copyTask(${task.id})" style="background: #17a2b8; color: white; border: none;">Copy</button>
            <button class="btn btn-danger" onclick="deleteTask(${task.id})">Delete</button>
        </div>
        `;
        } else {
            actionsHtml = `
        <div class="task-actions">
            <button class="btn btn-primary" onclick="editTask(${task.id})">Edit</button>
            <button class="btn btn-info" onclick="copyTask(${task.id})" style="background: #17a2b8; color: white; border: none;">Copy</button>
            ${currentUser.role === 'admin' ? `<button class="btn btn-danger" onclick="deleteTask(${task.id})">Delete</button>` : ''}
        </div>
        `;
        }
    }

    return `
        <div class="${cardClass}" style="padding: 12px; margin-bottom: 8px; cursor: pointer;" data-task-id="${task.id}" data-task-name="${task.task_name}" data-due-date="${dueDate || ''}" data-assigned-to="${task.assigned_to}" onclick="openInteractiveTaskPopup(${task.id})">
            <div style="display: flex; align-items: center; gap: 10px; flex-wrap: wrap;">
                <strong style="font-size: 14px; flex: 1; min-width: 200px;">${task.task_name}</strong>
                ${priorityBadge}
                ${statusBadge}
                ${typeBadge}
                <span style="font-size: 12px; color: #666; white-space: nowrap;">${assignedUser ? assignedUser.name : 'Unknown'}</span>
                <span style="font-size: 12px; color: #666; white-space: nowrap;">${location ? location.name : 'Unknown'}</span>
                <span style="font-size: 12px; color: #666; white-space: nowrap;">Due: ${dueDateStr}${ageingStr}</span>
                ${task.expected_completion_date ? `<span style="font-size: 12px; color: #17a2b8; white-space: nowrap;">Expected: ${formatDateDisplay(task.expected_completion_date)}</span>` : ''}
                ${task.comment ? `<span style="font-size: 11px; color: #4caf50; font-style: italic;" title="${task.comment}">💬</span>` : ''}
                ${showActions ? `
                <div style="display: flex; gap: 5px; margin-left: auto;">
                    ${!isTaskCompleted(task) && !task.removed_at && task.task_action !== 'not_done' ? `<button class="btn btn-success" onclick="completeTaskWithRemark(${task.id})" style="padding: 4px 8px; font-size: 11px;">Complete</button>` : ''}
                    ${!isTaskCompleted(task) && !task.removed_at && task.task_action !== 'in_process' && task.task_action !== 'not_done' ? `<button class="btn btn-info" onclick="markTaskInProcess(${task.id})" style="padding: 4px 8px; font-size: 11px; background: #17a2b8; color: white; border: none;">Mark In Process</button>` : ''}
                    ${task.task_action === 'in_process' ? `<button class="btn btn-info" onclick="openEditExpectedDateModal(${task.id})" style="padding: 4px 8px; font-size: 11px; background: #17a2b8; color: white; border: none;">Edit expected date</button>` : ''}
                    ${!isTaskCompleted(task) && !task.removed_at && task.task_action !== 'not_done' ? `<button class="btn btn-warning" onclick="removeTaskWithoutCompletion(${task.id})" style="padding: 4px 8px; font-size: 11px; background: #ff9800; color: white;">Remove</button>` : ''}
                    ${!isTaskCompleted(task) && !task.removed_at && task.task_action !== 'not_done' ? `<button class="btn btn-secondary" onclick="markTaskNotDone(${task.id})" style="padding: 4px 8px; font-size: 11px; background: #6c757d; color: white;">Task Not Done</button>` : ''}
                    ${!isTaskCompleted(task) && !task.removed_at && task.task_type === 'recurring' && !task.recurrence_stopped && task.task_action !== 'not_done' ? `<button class="btn btn-secondary" onclick="stopRecurringTask(${task.id})" style="padding: 4px 8px; font-size: 11px;">Stop Recurring</button>` : ''}
                    ${isTaskCompleted(task) && currentUser.role === 'admin' && !task.admin_finalized ? `
                        <button class="btn btn-success" onclick="acceptTaskCompletion(${task.id})" style="padding: 4px 8px; font-size: 11px;">✓ Accept</button>
                        <button class="btn btn-danger" onclick="rejectTaskCompletion(${task.id})" style="padding: 4px 8px; font-size: 11px;">✗ Reject</button>
                    ` : ''}
                    <button class="btn btn-primary" onclick="editTask(${task.id})" style="padding: 4px 8px; font-size: 11px;">Edit</button>
                    <button class="btn btn-info" onclick="copyTask(${task.id})" style="padding: 4px 8px; font-size: 11px; background: #17a2b8; color: white; border: none;">Copy</button>
                    ${currentUser.role === 'admin' ? `<button class="btn btn-danger" onclick="deleteTask(${task.id})" style="padding: 4px 8px; font-size: 11px;">Delete</button>` : ''}
                </div>
                ` : ''}
            </div>
            ${task.description ? `<p style="color: #666; margin: 5px 0 0 0; font-size: 12px;">${task.description}</p>` : ''}
            ${task.comment ? `
            <div style="margin-top: 8px; padding: 8px; background: #e8f5e9; border-radius: 4px; font-size: 12px;">
                <strong>User Comment:</strong> ${task.comment}
            </div>
            ` : ''}
            ${task.admin_comment ? `
            <div style="margin-top: 8px; padding: 8px; background: ${task.admin_finalized ? '#fff3cd' : '#f8d7da'}; border-radius: 4px; font-size: 12px;">
                <strong>${task.admin_finalized ? 'Admin Comment (Accepted):' : 'Admin Comment (Rejected):'}</strong> ${task.admin_comment}
                ${task.rejected_at ? `<br><small style="color: #666;">Rejected on: ${formatDateDisplay(new Date(task.rejected_at))}</small>` : ''}
                ${task.finalized_at ? `<br><small style="color: #666;">Accepted on: ${formatDateDisplay(new Date(task.finalized_at))}</small>` : ''}
            </div>
            ` : ''}
        </div>
    `;
}

function completeTaskWithRemark(taskId) {
    try {
        const data = getData();
        // Try to find task by ID (handle both integer and float IDs for recurring tasks)
        let task = data.tasks.find(t => t.id == taskId);

        // If not found, try to find by task name and due date (for recurring instances)
        if (!task) {
            // Get task details from the rendered task item if possible
            const taskElement = document.querySelector(`[data-task-id="${taskId}"]`);
            if (taskElement) {
                const taskName = taskElement.getAttribute('data-task-name');
                const dueDate = taskElement.getAttribute('data-due-date');
                const assignedTo = taskElement.getAttribute('data-assigned-to');

                if (taskName && dueDate) {
                    task = data.tasks.find(t =>
                        t.task_name === taskName &&
                        (t.due_date === dueDate || t.next_due_date === dueDate) &&
                        (!assignedTo || t.assigned_to === parseInt(assignedTo)) &&
                        !isTaskCompleted(t) &&
                        !t.removed_at
                    );
                }
            }
        }

        if (!task) {
            console.error('Task not found when opening completion modal:', taskId);
            console.error('Available task IDs:', data.tasks.map(t => ({ id: t.id, name: t.task_name, due: t.due_date || t.next_due_date })));
            reportError(`Task not found when opening completion modal. Task ID: ${taskId}`, 'Open completion');
            return;
        }

        // Store task details for fallback lookup
        const taskDetails = {
            id: task.id,
            name: task.task_name,
            dueDate: task.due_date || task.next_due_date,
            assignedTo: task.assigned_to,
            frequency: task.frequency
        };

        // Set default completion date to today
        const today = new Date();
        const todayStr = formatDateString(today);

        // Open completion date modal - no date restrictions
        const completionDateInput = document.getElementById('completionDate');
        const completionTaskIdInput = document.getElementById('completionTaskId');
        completionTaskIdInput.value = taskId;
        // Store task details as data attributes for fallback
        completionTaskIdInput.setAttribute('data-task-name', taskDetails.name);
        completionTaskIdInput.setAttribute('data-due-date', taskDetails.dueDate || '');
        completionTaskIdInput.setAttribute('data-assigned-to', taskDetails.assignedTo);
        completionTaskIdInput.setAttribute('data-frequency', taskDetails.frequency || '');

        document.getElementById('completionComment').value = 'DONE'; // Default comment
        completionDateInput.value = todayStr;
        // Remove all date restrictions
        completionDateInput.removeAttribute('min');
        completionDateInput.removeAttribute('max');
        document.getElementById('completionDateModal').classList.add('active');
    } catch (err) {
        reportError(err, 'Open completion modal');
    }
}

function closeCompletionDateModal() {
    document.getElementById('completionDateModal').classList.remove('active');
    document.getElementById('completionDateForm').reset();
    // Reset completion type to default
    document.getElementById('completionType').value = 'completed';
    // Reset button styles
    const needsImprovementBtn = document.querySelector('#completionDateModal button[onclick*="setCompletionType"]');
    if (needsImprovementBtn) {
        needsImprovementBtn.style.background = '#ffc107';
        needsImprovementBtn.style.color = '#000';
        needsImprovementBtn.style.fontWeight = 'normal';
    }
    const submitBtn = document.querySelector('#completionDateForm button[type="submit"]');
    if (submitBtn) {
        submitBtn.style.background = '';
        submitBtn.style.fontWeight = 'normal';
        submitBtn.textContent = 'Complete Task';
    }
}

function setCompletionType(type) {
    document.getElementById('completionType').value = type;
    const needsImprovementBtn = document.querySelector('#completionDateModal button[onclick*="setCompletionType"]');
    const submitBtn = document.querySelector('#completionDateForm button[type="submit"]');

    if (type === 'completed_need_improvement') {
        if (needsImprovementBtn) {
            needsImprovementBtn.style.background = '#dc3545';
            needsImprovementBtn.style.color = '#fff';
            needsImprovementBtn.style.fontWeight = 'bold';
        }
        if (submitBtn) {
            submitBtn.style.background = '#ffc107';
            submitBtn.textContent = 'Complete with Needs Improvement';
        }
    } else {
        if (needsImprovementBtn) {
            needsImprovementBtn.style.background = '#ffc107';
            needsImprovementBtn.style.color = '#000';
            needsImprovementBtn.style.fontWeight = 'normal';
        }
        if (submitBtn) {
            submitBtn.style.background = '';
            submitBtn.textContent = 'Complete Task';
        }
    }
}

function saveTaskCompletion(event) {
    event.preventDefault();

    try {
        const completionTaskIdInput = document.getElementById('completionTaskId');
        let taskId = parseFloat(completionTaskIdInput.value); // Use parseFloat for recurring task IDs
        const comment = document.getElementById('completionComment').value.trim();
        const completionDate = document.getElementById('completionDate').value;
        const completionType = document.getElementById('completionType').value || 'completed';

        if (!comment) {
            alert('Completion remark is required.');
            return;
        }

        if (!completionDate) {
            alert('Completion date is required.');
            return;
        }

        const data = getData();
        // Try to find task by ID (handle both integer and float IDs)
        let task = data.tasks.find(t => t.id == taskId);

        // Fallback: If task not found by ID, try to find by stored details (for recurring instances)
        if (!task) {
            const taskName = completionTaskIdInput.getAttribute('data-task-name');
            const dueDate = completionTaskIdInput.getAttribute('data-due-date');
            const assignedTo = completionTaskIdInput.getAttribute('data-assigned-to');
            const frequency = completionTaskIdInput.getAttribute('data-frequency');

            console.log('Task not found by ID, trying fallback lookup:', {
                taskId,
                taskName,
                dueDate,
                assignedTo,
                frequency
            });

            if (taskName && dueDate) {
                task = data.tasks.find(t =>
                    t.task_name === taskName &&
                    (t.due_date === dueDate || t.next_due_date === dueDate) &&
                    (!assignedTo || t.assigned_to === parseInt(assignedTo)) &&
                    (!frequency || t.frequency === frequency) &&
                    !isTaskCompleted(t) &&
                    !t.removed_at
                );

                if (task) {
                    console.log('Found task via fallback lookup:', task.id);
                    taskId = task.id; // Update taskId to the found task's ID
                }
            }
        }

        if (!task) {
            console.error('Task not found:', {
                searchedId: taskId,
                taskName: completionTaskIdInput.getAttribute('data-task-name'),
                dueDate: completionTaskIdInput.getAttribute('data-due-date'),
                availableTasks: data.tasks.filter(t =>
                    t.task_name === completionTaskIdInput.getAttribute('data-task-name')
                ).map(t => ({
                    id: t.id,
                    name: t.task_name,
                    due: t.due_date || t.next_due_date,
                    completed: isTaskCompleted(t),
                    removed: t.removed_at
                }))
            });
            reportError(`Task not found. Task ID: ${taskId}. Please try again or refresh the page.`, 'Save completion');
            return;
        }

        // Parse completion date flexibly (YYYY-MM-DD or DD-MM-YYYY)
        const completionDateObj = parseDateFlexible(completionDate);
        if (!completionDateObj) {
            alert('Invalid completion date format.');
            return;
        }

        updateData(data => {
            // Find task again inside updateData (handle both integer and float IDs)
            let task = data.tasks.find(t => t.id == taskId);

            // Fallback lookup if still not found
            if (!task) {
                const taskName = completionTaskIdInput.getAttribute('data-task-name');
                const dueDate = completionTaskIdInput.getAttribute('data-due-date');
                const assignedTo = completionTaskIdInput.getAttribute('data-assigned-to');
                const frequency = completionTaskIdInput.getAttribute('data-frequency');

                if (taskName && dueDate) {
                    task = data.tasks.find(t =>
                        t.task_name === taskName &&
                        (t.due_date === dueDate || t.next_due_date === dueDate) &&
                        (!assignedTo || t.assigned_to === parseInt(assignedTo)) &&
                        (!frequency || t.frequency === frequency) &&
                        !isTaskCompleted(t) &&
                        !t.removed_at
                    );

                    if (task) {
                        taskId = task.id; // Update taskId
                    }
                }
            }

            if (task) {
                task.rejected_at = null;
                task.rejected_by = null;
                task.previous_submission_comment = null;
                task.admin_comment = null;
                task.admin_finalized = false;
                task.finalized_at = null;
                task.finalized_by = null;
                // Set task action based on completion type (completed or completed_need_improvement)
                task.task_action = completionType;
                task.comment = comment;
                // Store completion date as ISO string - set to noon to avoid timezone issues
                const completionDateTime = new Date(completionDateObj);
                completionDateTime.setHours(12, 0, 0, 0);
                task.completed_at = completionDateTime.toISOString();
                task.completion_date = formatDateString(completionDateObj); // Store as YYYY-MM-DD
                task.completed_by = currentUser.id;
                // If user is admin, auto-finalize (no admin review needed)
                if (currentUser.role === 'admin') {
                    task.admin_finalized = true;
                    task.finalized_at = new Date().toISOString();
                    task.finalized_by = currentUser.id;
                    task.admin_comment = comment; // Use completion comment as admin comment
                } else {
                    // If user is not admin, mark as pending admin review
                    task.admin_finalized = false;
                }

                // Handle recurring tasks - generate next instance when completed
                if (task.task_type === 'recurring' && task.frequency && !task.recurrence_stopped) {
                    const currentDueDate = task.due_date || task.next_due_date;
                    if (currentDueDate) {
                        // Calculate next due date based on the completed instance's due date
                        const nextDueDate = calculateNextRecurrenceDateForInstance(task, currentDueDate);
                        if (nextDueDate) {
                            // Check if next instance already exists
                            const existingNext = data.tasks.find(t =>
                                t.task_name === task.task_name &&
                                t.assigned_to === task.assigned_to &&
                                t.frequency === task.frequency &&
                                (t.next_due_date === nextDueDate || t.due_date === nextDueDate) &&
                                t.id !== task.id &&
                                !isTaskCompleted(t)
                            );

                            // Only create if it doesn't exist and is in the future
                            const today = new Date();
                            today.setHours(0, 0, 0, 0);
                            const nextDateObj = parseDateFlexible(nextDueDate);

                            if (!existingNext && nextDateObj && nextDateObj > today) {
                                const baseTask = data.tasks.find(t =>
                                    t.task_name === task.task_name &&
                                    t.assigned_to === task.assigned_to &&
                                    t.frequency === task.frequency &&
                                    t.start_date // Base task has start_date
                                ) || task;

                                const newTask = {
                                    ...baseTask,
                                    id: Date.now() + Math.random(),
                                    task_action: 'not_completed',
                                    comment: null,
                                    completed_at: null,
                                    due_date: null,
                                    next_due_date: nextDueDate,
                                    start_date: null, // Generated instances don't have start_date
                                    created_at: new Date().toISOString(),
                                    recurrence_stopped: false,
                                    admin_finalized: false
                                };
                                data.tasks.push(newTask);
                            }
                        }
                    }
                }
            }
        });

        closeCompletionDateModal();
        processRecurringTasks();
        renderTasks();
        renderDashboard();
        renderCalendar();
        renderInteractiveDashboard();
        if (drilldownContext) {
            renderDrilldown();
        }
    } catch (err) {
        console.error('Complete task error', err);
        alert(`Unable to complete task: ${err.message || err}`);
    }
}

function updateTaskStatus(taskId, status) {
    const comment = prompt('Enter completion remark (optional):');

    updateData(data => {
        const task = data.tasks.find(t => t.id == taskId);
        if (task) {
            // Preserve past data - don't modify completed_at if already set
            if (status === 'completed' && !task.completed_at) {
                task.completed_at = new Date().toISOString();
            } else if (status !== 'completed') {
                task.completed_at = null;
                task.admin_finalized = false;
            }

            task.task_action = status;
            if (comment) {
                task.comment = comment;
            }
        }
    });
    processRecurringTasks();
    renderTasks();
    renderDashboard();
    renderCalendar();
    renderInteractiveDashboard();
}

// Accept task completion (Admin finalizes the task)
function acceptTaskCompletion(taskId) {
    const data = getData();
    const task = data.tasks.find(t => t.id == taskId);
    const defaultComment = (task && task.comment && String(task.comment).trim()) || 'DONE';
    const raw = prompt('Admin comment (OK for default):', defaultComment);
    if (raw === null) return;
    const adminComment = (raw.trim() || defaultComment).trim();

    updateData(data => {
        const task = data.tasks.find(t => t.id == taskId);
        if (task) {
            task.admin_finalized = true;
            task.admin_comment = adminComment;
            task.finalized_at = new Date().toISOString();
            task.finalized_by = currentUser.id;
            // Task remains completed, just finalized by admin
        }
    });

    renderTasks();
    renderDashboard();
    renderCalendar();
    renderInteractiveDashboard();
}

// Mark task as In Process with expected completion date
function markTaskInProcess(taskId) {
    const data = getData();
    let task = data.tasks.find(t => t.id == taskId);
    if (!task) {
        const taskElement = document.querySelector(`[data-task-id="${taskId}"]`);
        if (taskElement) {
            const taskName = taskElement.getAttribute('data-task-name');
            const dueDate = taskElement.getAttribute('data-due-date');
            const assignedTo = taskElement.getAttribute('data-assigned-to');
            if (taskName && dueDate) {
                task = data.tasks.find(t =>
                    t.task_name === taskName &&
                    (t.due_date === dueDate || t.next_due_date === dueDate) &&
                    (!assignedTo || t.assigned_to === parseInt(assignedTo)) &&
                    !isTaskCompleted(t) && !t.removed_at
                );
            }
        }
    }
    if (!task) {
        reportError('Task not found.', 'Mark In Process');
        return;
    }
    const today = new Date();
    const todayStr = formatDateString(today); // YYYY-MM-DD
    const defaultDate = task.expected_completion_date || todayStr;
    const dateStr = prompt('Enter expected completion date (DD-MM-YYYY):', formatDateDisplay(defaultDate));
    if (dateStr == null) return;
    const trimmed = dateStr.trim();
    if (!trimmed) {
        alert('Expected completion date is required.');
        return;
    }
    const parsed = parseDateFlexible(trimmed);
    if (!parsed) {
        alert('Invalid date. Please use DD-MM-YYYY format.');
        return;
    }
    const normalized = formatDateString(parsed);
    updateData(data => {
        const t = data.tasks.find(x => x.id == taskId);
        if (t) {
            t.task_action = 'in_process';
            t.expected_completion_date = normalized;
        }
    });
    renderTasks();
    renderDashboard();
    renderCalendar();
    renderInteractiveDashboard();
    if (drilldownContext) renderDrilldown();
}

// Update expected completion date (for In Process tile or task card)
function updateExpectedCompletionDate(taskId, dateStr) {
    const trimmed = (dateStr || '').trim();
    if (!trimmed) return;
    const parsed = parseDateFlexible(trimmed);
    if (!parsed) {
        alert('Invalid date. Please use DD-MM-YYYY format.');
        return false;
    }
    const normalized = formatDateString(parsed);
    updateData(data => {
        const task = data.tasks.find(t => t.id == taskId);
        if (task) task.expected_completion_date = normalized;
    });
    renderDashboard();
    renderTasks();
    renderInteractiveDashboard();
    if (drilldownContext) renderDrilldown();
    return true;
}

// Open modal/prompt to edit expected completion date for an In Process task
function openEditExpectedDateModal(taskId) {
    const data = getData();
    const task = data.tasks.find(t => t.id == taskId);
    if (!task) {
        alert('Task not found.');
        return;
    }
    const current = task.expected_completion_date || task.due_date || task.next_due_date || formatDateString(new Date());
    const dateStr = prompt('Edit expected completion date (DD-MM-YYYY):', formatDateDisplay(current));
    if (dateStr != null) updateExpectedCompletionDate(taskId, dateStr);
}

// Remove task without completion
function removeTaskWithoutCompletion(taskId) {
    const comment = prompt('Enter reason for removing this task (required):');
    if (!comment || comment.trim() === '') {
        alert('Removal reason is required.');
        return;
    }

    updateData(data => {
        const task = data.tasks.find(t => t.id == taskId);
        if (task) {
            // Mark task as removed without completion
            task.removed_at = new Date().toISOString();
            task.removed_by = currentUser.id;
            task.comment = comment.trim();
            task.task_action = 'not_completed'; // Keep as not_completed but mark as removed
            // Don't set completed_at - task is removed, not completed
        }
    });

    processRecurringTasks();
    renderTasks();
    renderDashboard();
    renderCalendar();
    renderInteractiveDashboard();
    if (drilldownContext) {
        renderDrilldown();
    }
}

// Mark task as Not Done (closes current occurrence but keeps future recurring instances)
function markTaskNotDone(taskId) {
    if (!confirm('Mark this task as Not Done?\n\nThis will close the current occurrence but will not stop future recurring instances.')) {
        return;
    }

    updateData(data => {
        const task = data.tasks.find(t => t.id == taskId);
        if (task) {
            task.task_action = 'not_done';
            task.comment = 'Not Done';
            // Clear completion/admin review data
            task.completed_at = null;
            task.completion_date = null;
            task.completed_by = null;
            task.admin_finalized = false;
            task.admin_comment = null;
            task.finalized_at = null;
            task.finalized_by = null;
            // Ensure it's not treated as removed
            task.removed_at = null;
            task.removed_by = null;
        }
    });

    processRecurringTasks();
    renderTasks();
    renderDashboard();
    renderCalendar();
    renderInteractiveDashboard();
    if (drilldownContext) {
        renderDrilldown();
    }
}

// Stop generating future recurring instances (after today)
function stopRecurringTask(taskId) {
    if (!confirm('Stop generating all future recurring instances for this task after today?\n\nExisting past and today\'s occurrences will remain.')) {
        return;
    }

    updateData(data => {
        const reference = data.tasks.find(t => t.id == taskId);
        if (!reference || reference.task_type !== 'recurring') return;

        data.tasks.forEach(t => {
            if (t.task_type === 'recurring' &&
                t.task_name === reference.task_name &&
                t.assigned_to === reference.assigned_to &&
                t.location_id === reference.location_id &&
                t.frequency === reference.frequency) {
                t.recurrence_stopped = true;
            }
        });
    });

    processRecurringTasks();
    renderTasks();
    renderDashboard();
    renderCalendar();
    renderInteractiveDashboard();
    if (drilldownContext) {
        renderDrilldown();
    }
}

// Reject task completion (Admin rejects and sets task back to pending)
function rejectTaskCompletion(taskId) {
    const adminComment = prompt('Enter rejection reason for the assignee (required):');
    if (adminComment === null || adminComment.trim() === '') {
        alert('Rejection reason is required to reject the task.');
        return;
    }

    updateData(data => {
        const task = data.tasks.find(t => t.id == taskId);
        if (task) {
            if (task.comment && String(task.comment).trim()) {
                task.previous_submission_comment = String(task.comment).trim();
            }
            task.task_action = 'not_completed';
            task.admin_finalized = false;
            task.admin_comment = adminComment.trim();
            task.rejected_at = new Date().toISOString();
            task.rejected_by = currentUser.id;
            task.completed_at = null;
            task.completed_by = null;
            task.completion_date = null;
            task.comment = null;
        }
    });

    renderTasks();
    renderDashboard();
    renderCalendar();
    renderInteractiveDashboard();
}

function filterTasksByAdminReview() {
    // Clear any existing filter headers first
    const tasksListElement = document.getElementById('tasksList');
    if (tasksListElement) {
        // Remove any existing date-filter-header
        const existingHeader = tasksListElement.querySelector('.date-filter-header');
        if (existingHeader) {
            existingHeader.remove();
        }
        tasksListElement.innerHTML = '';
    }

    const data = getData();

    const fromVal = document.getElementById('filterTaskMonthFrom')?.value;
    const toVal = document.getElementById('filterTaskMonthTo')?.value;
    let fromDate = null, toDate = null;
    if (fromVal && toVal) {
        const [fy, fm] = fromVal.split('-').map(Number);
        const [ty, tm] = toVal.split('-').map(Number);
        fromDate = new Date(fy, fm - 1, 1);
        toDate = new Date(ty, tm, 0);
        toDate.setHours(23, 59, 59, 999);
    }

    const reviewTasks = data.tasks.filter(task => {
        if (!isTaskCompleted(task) || task.admin_finalized) return false;

        if (fromDate && toDate) {
            const dueDate = task.due_date || task.next_due_date;
            if (dueDate) {
                const [y, m, d] = dueDate.split('-').map(Number);
                const taskDate = new Date(y, m - 1, d);
                if (taskDate < fromDate || taskDate > toDate) return false;
            } else return false;
        }

        return (task.assigned_to === currentUser.id ||
            (task.is_team_task && currentUser.role === 'admin') ||
            currentUser.role === 'admin');
    });

    // Store filtered tasks for export
    window.currentFilteredTasksForExport = reviewTasks;

    const tasksHtml = reviewTasks.length > 0
        ? reviewTasks.map(task => renderTaskItem(task, true)).join('')
        : '<p style="text-align: center; color: #999; padding: 20px;">No tasks pending admin review</p>';

    document.getElementById('tasksList').innerHTML = `
        <div class="date-filter-header" style="margin-bottom: 20px; padding: 15px; background: #fff3cd; border-radius: 8px; display: flex; justify-content: space-between; align-items: center;">
            <h3 style="margin: 0;">Pending Admin Review (${reviewTasks.length})</h3>
            <button class="btn btn-secondary" onclick="filterTasks()" style="padding: 5px 15px; font-size: 12px;">Clear Filter</button>
        </div>
        ${tasksHtml}
    `;

    // Clear filter dropdowns
    document.getElementById('filterStatus').value = '';
    document.getElementById('filterType').value = '';
    document.getElementById('filterTeam').value = '';
    document.getElementById('searchTasks').value = '';
}

function addWorkingDays(startDate, days) {
    let current = new Date(startDate);
    current.setHours(0, 0, 0, 0); // Reset time to midnight
    const data = getData();
    const holidays = data.holidays.map(h => h.date);
    let added = 0;

    // Loop until we've added the required number of working days
    while (added < days) {
        current.setDate(current.getDate() + 1); // Move to next day
        const dayOfWeek = current.getDay();
        const dateStr = formatDateString(current);

        // Skip weekends (Saturday = 6, Sunday = 0) and holidays
        // Only count the day if it's a weekday (Monday-Friday) and not a holiday
        if (dayOfWeek !== 0 && dayOfWeek !== 6 && !holidays.includes(dateStr)) {
            added++;
        }
    }

    return formatDateString(current);
}

// Task Modal
function openTaskModal(taskId = null) {
    const data = getData();
    const modal = document.getElementById('taskModal');
    const form = document.getElementById('taskForm');

    // Populate users (include current assignee on edit even if outside picker rules, e.g. legacy)
    let pickUsers = taskAssigneePickerUsers();
    if (taskId) {
        const t0 = data.tasks.find(t => t.id == taskId);
        if (t0) {
            const aid = Number(t0.assigned_to);
            if (aid && !pickUsers.some(u => Number(u.id) === aid)) {
                const u = data.users.find(x => Number(x.id) === aid);
                if (u) pickUsers = [u, ...pickUsers];
            }
        }
    }
    const userSelect = document.getElementById('taskAssignedTo');
    userSelect.innerHTML = '<option value="">Select user</option>' +
        pickUsers.map(u =>
            `<option value="${u.id}">${escapeHtml(u.name)}</option>`
        ).join('');

    // Populate locations
    const locationSelect = document.getElementById('taskLocation');
    // Find "Combine" location for default
    const combineLocation = data.locations.find(l => l.name.toLowerCase() === 'combine');
    locationSelect.innerHTML = '<option value="">Select location</option>' +
        data.locations.map(l =>
            `<option value="${l.id}" ${combineLocation && l.id === combineLocation.id ? 'selected' : ''}>${l.name}</option>`
        ).join('');

    // Populate segregation types
    const segSelect = document.getElementById('taskSegregation');
    segSelect.innerHTML = '<option value="">Select type</option>' +
        data.segregationTypes.map(s =>
            `<option value="${s.id}">${s.name}</option>`
        ).join('');

    if (taskId) {
        const task = data.tasks.find(t => t.id == taskId);
        if (task) {
            document.getElementById('taskModalTitle').textContent = 'Edit Task';
            document.getElementById('taskId').value = task.id;
            document.getElementById('taskName').value = task.task_name;
            document.getElementById('taskDescription').value = task.description || '';
            document.getElementById('taskAssignedTo').value = task.assigned_to;
            document.getElementById('taskLocation').value = task.location_id;
            document.getElementById('taskType').value = task.task_type;
            document.getElementById('taskDueDate').value = task.due_date || '';
            if (task.task_type === 'without_due_date') {
                document.getElementById('taskPriorityNoDue').value = task.priority || 'medium';
            } else {
                document.getElementById('taskPriority').value = task.priority || 'medium';
            }
            document.getElementById('taskFrequency').value = task.frequency || '';
            document.getElementById('taskDueDateType').value = task.due_date_type || 'calendar_day';
            document.getElementById('taskDueDay').value = task.due_day || 1;

            // Auto-fill start_date if missing for recurring tasks
            if (task.task_type === 'recurring') {
                if (task.start_date) {
                    document.getElementById('taskStartDate').value = task.start_date;
                } else {
                    // If no start_date, use the current due date or today
                    const fallbackDate = task.next_due_date || task.due_date || formatDateString(new Date());
                    document.getElementById('taskStartDate').value = fallbackDate;
                    console.log(`Auto-filled missing start_date for task "${task.task_name}" with: ${fallbackDate}`);
                }
            } else {
                document.getElementById('taskStartDate').value = task.start_date || '';
            }

            document.getElementById('taskRecurrenceType').value = task.recurrence_type || 'calendar_day';
            document.getElementById('taskRecurrenceInterval').value = task.recurrence_interval || 1;
            document.getElementById('taskSegregation').value = task.segregation_type_id || '';
            document.getElementById('taskEstMinutes').value = task.est_minutes || '';
            document.getElementById('taskIsTeam').checked = task.is_team_task || false;
            document.getElementById('taskComment').value = task.comment || '';
            if (task.task_type === 'recurring') {
                document.getElementById('taskStopRecurrence').checked = task.recurrence_stopped || false;
                document.getElementById('taskRecurrenceStopped').value = task.recurrence_stopped ? 'true' : 'false';
            }
            document.getElementById('taskCommentGroup').style.display = task.task_action === 'completed' ? 'block' : 'none';
            toggleRecurrenceFields();
            updateRecurrenceFields();
        }
    } else {
        form.reset();
        document.getElementById('taskModalTitle').textContent = 'New Task';
        document.getElementById('taskId').value = '';

        // Set default Task Type to "one_time"
        document.getElementById('taskType').value = 'one_time';

        // Set default Location to "Combine" if it exists
        const combineLocation = data.locations.find(l => l.name.toLowerCase() === 'combine');
        if (combineLocation) {
            document.getElementById('taskLocation').value = combineLocation.id;
        }

        // Set default Due Date to today
        const today = new Date();
        const todayStr = today.getFullYear() + '-' +
            String(today.getMonth() + 1).padStart(2, '0') + '-' +
            String(today.getDate()).padStart(2, '0');
        document.getElementById('taskDueDate').value = todayStr;

        // Set default Priority to "medium"
        document.getElementById('taskPriority').value = 'medium';
        document.getElementById('taskPriorityNoDue').value = 'medium';

        // Set other defaults
        document.getElementById('taskStartDate').value = todayStr;
        document.getElementById('taskDueDay').value = today.getDate();
        document.getElementById('taskCommentGroup').style.display = 'none';
        document.getElementById('taskStopRecurrence').checked = false;
        toggleRecurrenceFields();
        updateRecurrenceFields();
    }

    modal.classList.add('active');
}

function closeTaskModal() {
    const modal = document.getElementById('taskModal');
    modal.classList.remove('active');
    modal.style.display = ''; // Clear any inline style from legacy code
}

function toggleRecurrenceFields() {
    const taskType = document.getElementById('taskType').value;
    const oneTimeFields = document.getElementById('oneTimeFields');
    const recurringFields = document.getElementById('recurringFields');
    const withoutDueDateFields = document.getElementById('withoutDueDateFields');
    const recurringDueDateGroup = document.getElementById('recurringDueDateGroup');
    const stopRecurrenceGroup = document.getElementById('stopRecurrenceGroup');

    if (taskType === 'recurring') {
        oneTimeFields.style.display = 'none';
        recurringFields.style.display = 'block';
        withoutDueDateFields.style.display = 'none';
        recurringDueDateGroup.style.display = 'block';
        stopRecurrenceGroup.style.display = 'block';
        document.getElementById('taskFrequency').required = true;
        document.getElementById('taskDueDateType').required = true;
        document.getElementById('taskDueDay').required = true;
    } else if (taskType === 'without_due_date') {
        oneTimeFields.style.display = 'none';
        recurringFields.style.display = 'none';
        withoutDueDateFields.style.display = 'block';
        recurringDueDateGroup.style.display = 'none';
        stopRecurrenceGroup.style.display = 'none';
        document.getElementById('taskFrequency').required = false;
        document.getElementById('taskDueDateType').required = false;
        document.getElementById('taskDueDay').required = false;
    } else if (taskType === 'work_plan' || taskType === 'audit_point') {
        // Work Plan and Audit Point behave like one_time - has due date and priority
        oneTimeFields.style.display = 'block';
        recurringFields.style.display = 'none';
        withoutDueDateFields.style.display = 'none';
        recurringDueDateGroup.style.display = 'none';
        stopRecurrenceGroup.style.display = 'none';
        document.getElementById('taskFrequency').required = false;
        document.getElementById('taskDueDateType').required = false;
        document.getElementById('taskDueDay').required = false;
    } else {
        // one_time
        oneTimeFields.style.display = 'block';
        recurringFields.style.display = 'none';
        withoutDueDateFields.style.display = 'none';
        recurringDueDateGroup.style.display = 'none';
        stopRecurrenceGroup.style.display = 'none';
        document.getElementById('taskFrequency').required = false;
        document.getElementById('taskDueDateType').required = false;
        document.getElementById('taskDueDay').required = false;
    }
}

function updateRecurrenceFields() {
    const frequency = document.getElementById('taskFrequency').value;
    const recurrenceTypeGroup = document.getElementById('recurrenceTypeGroup');
    const recurrenceIntervalGroup = document.getElementById('recurrenceIntervalGroup');

    // For daily, weekly, monthly, yearly - hide interval and type fields
    if (frequency && ['daily', 'weekly', 'monthly', 'quarterly', 'halfyearly', 'yearly'].includes(frequency)) {
        recurrenceTypeGroup.style.display = 'none';
        recurrenceIntervalGroup.style.display = 'none';
    } else {
        recurrenceTypeGroup.style.display = 'block';
        recurrenceIntervalGroup.style.display = 'block';
    }
}

function saveTask(event) {
    event.preventDefault();
    const data = getData();
    const taskId = document.getElementById('taskId').value;
    const taskType = document.getElementById('taskType').value;
    const frequency = document.getElementById('taskFrequency').value;
    const stopRecurrence = document.getElementById('taskStopRecurrence').checked;

    // For edit: ensure task exists (handles float IDs from recurring instances; form value is string)
    const isEdit = taskId !== '' && taskId != null;
    const existingTask = isEdit ? data.tasks.find(t => t.id == taskId) : null;

    const assignPick = parseInt(document.getElementById('taskAssignedTo').value, 10);
    const allowedIds = new Set(taskAssigneePickerUsers().map(u => Number(u.id)));
    const assignUnchanged = existingTask && Number(existingTask.assigned_to) === assignPick;
    if (!assignPick || Number.isNaN(assignPick) || (!allowedIds.has(assignPick) && !assignUnchanged)) {
        alert('Please select a user from your organization for this task.');
        return;
    }
    if (isEdit && !existingTask) {
        alert('This task could not be found. It may have been deleted. Please refresh and try again.');
        return;
    }

    let dueDate = null;
    let nextDueDate = null;

    if (taskType === 'one_time' || taskType === 'work_plan' || taskType === 'audit_point') {
        dueDate = document.getElementById('taskDueDate').value;
    } else if (taskType === 'without_due_date') {
        dueDate = null;
    } else if (taskType === 'recurring') {
        const startDate = document.getElementById('taskStartDate').value;
        const dueDateType = document.getElementById('taskDueDateType').value;
        const dueDay = parseInt(document.getElementById('taskDueDay').value);

        if (startDate) {
            nextDueDate = calculateRecurringDueDate(startDate, frequency, dueDateType, dueDay);
        }
    }

    const task = {
        id: existingTask ? existingTask.id : Date.now(),
        task_name: document.getElementById('taskName').value,
        description: document.getElementById('taskDescription').value,
        assigned_to: assignPick,
        location_id: parseInt(document.getElementById('taskLocation').value),
        task_type: taskType,
        due_date: dueDate,
        priority: (taskType === 'one_time' || taskType === 'work_plan' || taskType === 'audit_point') ? document.getElementById('taskPriority').value :
            taskType === 'without_due_date' ? document.getElementById('taskPriorityNoDue').value : null,
        frequency: taskType === 'recurring' ? frequency : null,
        due_date_type: taskType === 'recurring' ? document.getElementById('taskDueDateType').value : null,
        due_day: taskType === 'recurring' ? parseInt(document.getElementById('taskDueDay').value) : null,
        start_date: taskType === 'recurring' ? document.getElementById('taskStartDate').value : null,
        recurrence_type: taskType === 'recurring' && !frequency ? document.getElementById('taskRecurrenceType').value : null,
        recurrence_interval: taskType === 'recurring' && !frequency ? parseInt(document.getElementById('taskRecurrenceInterval').value) : null,
        next_due_date: nextDueDate,
        segregation_type_id: document.getElementById('taskSegregation').value ? parseInt(document.getElementById('taskSegregation').value) : null,
        est_minutes: document.getElementById('taskEstMinutes').value ? parseInt(document.getElementById('taskEstMinutes').value) : null,
        is_team_task: document.getElementById('taskIsTeam').checked,
        task_action: 'not_completed',
        comment: null,
        recurrence_stopped: stopRecurrence,
        created_by: currentUser.id,
        created_at: existingTask ? (existingTask.created_at || new Date().toISOString()) : new Date().toISOString(),
        completed_at: null
    };

    updateData(data => {
        if (taskId) {
            const index = data.tasks.findIndex(t => t.id == taskId);
            if (index !== -1) {
                const existingTaskInData = data.tasks[index];
                // Preserve past data - only update future-related fields
                task.created_at = existingTaskInData.created_at;
                if (existingTaskInData.completed_at) {
                    task.completed_at = existingTaskInData.completed_at;
                }
                if (existingTaskInData.task_number != null && existingTaskInData.task_number !== '') {
                    task.task_number = existingTaskInData.task_number;
                } else {
                    task.task_number = getNextTaskNumberFromData(data);
                }
                data.tasks[index] = { ...existingTaskInData, ...task };

                // If stopping recurrence, remove all future related instances
                if (stopRecurrence && task.task_type === 'recurring') {
                    const today = new Date();
                    today.setHours(0, 0, 0, 0);
                    const relatedTasks = data.tasks.filter(t =>
                        t.id != taskId &&
                        t.task_name === task.task_name &&
                        t.assigned_to === task.assigned_to &&
                        t.frequency === task.frequency &&
                        (t.next_due_date || t.due_date) &&
                        new Date(t.next_due_date || t.due_date) > today
                    );
                    // Remove future instances
                    relatedTasks.forEach(relatedTask => {
                        const relatedIndex = data.tasks.findIndex(t => t.id === relatedTask.id);
                        if (relatedIndex !== -1) {
                            data.tasks.splice(relatedIndex, 1);
                        }
                    });
                }
            }
        } else {
            task.task_number = getNextTaskNumberFromData(data);
            data.tasks.push(task);
        }
    });

    processRecurringTasks();
    closeTaskModal();
    renderTasks();
    renderDashboard();
    renderCalendar();
    renderInteractiveDashboard();
}

// Quick Task Functions
function openQuickTaskModal() {
    const modal = document.getElementById('quickTaskModal');
    const form = document.getElementById('quickTaskForm');

    // Reset form
    form.reset();

    // Set default due date to today (use formatDateString for consistency)
    document.getElementById('quickTaskDueDate').value = formatDateString(new Date());

    // Focus on task name field
    setTimeout(() => {
        document.getElementById('quickTaskName').focus();
    }, 100);

    modal.classList.add('active');
}

function closeQuickTaskModal() {
    document.getElementById('quickTaskModal').classList.remove('active');
    document.getElementById('quickTaskForm').reset();
}

// Interactive Task Popup Functions
function openInteractiveTaskPopup(taskId) {
    const data = getData();
    const task = data.tasks.find(t => t.id == taskId);

    if (!task) return;

    const assignedUser = data.users.find(u => u.id === task.assigned_to);
    const location = data.locations.find(l => l.id === task.location_id);
    const segregation = data.segregationTypes.find(s => s.id === task.segregation_type_id);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const dueDate = task.due_date || task.next_due_date;
    const dueDateStr = dueDate ? formatDateDisplay(dueDate) : 'Not set';

    // Determine status
    let statusClass = '';
    let statusText = '';
    if (task.task_action === 'completed') {
        statusClass = 'badge-completed';
        statusText = 'Completed';
    } else if (task.task_action === 'completed_need_improvement') {
        statusClass = 'badge-warning';
        statusText = 'Need Improvement';
    } else if (task.task_action === 'in_process') {
        statusClass = 'badge-info';
        statusText = 'In Process';
    } else if (task.task_action === 'not_done') {
        statusClass = 'badge-not-done';
        statusText = 'Not Done';
    } else if (dueDate) {
        const dateParts = dueDate.split('-');
        const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
        taskDate.setHours(0, 0, 0, 0);
        if (taskDate < today) {
            statusClass = 'badge-pending';
            statusText = 'Overdue';
        } else {
            statusClass = 'badge-low';
            statusText = 'Pending';
        }
    } else {
        statusClass = 'badge-low';
        statusText = 'No Due Date';
    }

    // Build action buttons (Need Improvement = Done, so they get completed-style actions)
    let actionsHtml = '';
    if (!isTaskCompleted(task) && !task.removed_at) {
        actionsHtml = `
            <div style="margin-top: 20px; padding-top: 20px; border-top: 1px solid #e0e0e0;">
                <h3 style="margin-bottom: 15px; color: #333;">Actions</h3>
                <div style="display: flex; gap: 10px; flex-wrap: wrap;">
                    <button class="btn btn-success" onclick="event.stopPropagation(); completeTaskWithRemark(${task.id}); closeInteractiveTaskPopup(); renderDashboard(); renderInteractiveDashboard();" 
                            style="padding: 10px 20px;">Complete with Remark</button>
                    ${task.task_action !== 'in_process' ? `
                        <button class="btn btn-info" onclick="event.stopPropagation(); markTaskInProcess(${task.id}); closeInteractiveTaskPopup(); renderDashboard(); renderInteractiveDashboard();" 
                                style="padding: 10px 20px; background: #17a2b8; color: white; border: none;">Mark In Process</button>
                    ` : `
                        <button class="btn btn-info" onclick="event.stopPropagation(); openEditExpectedDateModal(${task.id}); closeInteractiveTaskPopup(); renderDashboard(); renderInteractiveDashboard();" 
                                style="padding: 10px 20px; background: #17a2b8; color: white; border: none;">Edit Expected Date</button>
                    `}
                    <button class="btn btn-secondary" onclick="event.stopPropagation(); markTaskNotDone(${task.id}); closeInteractiveTaskPopup(); renderDashboard(); renderInteractiveDashboard();" 
                            style="padding: 10px 20px;">Task Not Done</button>
                    ${task.task_type === 'recurring' && !task.recurrence_stopped ? `
                        <button class="btn btn-secondary" onclick="event.stopPropagation(); stopRecurringTask(${task.id}); closeInteractiveTaskPopup(); renderDashboard(); renderInteractiveDashboard();" 
                                style="padding: 10px 20px;">Stop Recurring</button>
                    ` : ''}
                    <button class="btn btn-primary" onclick="event.stopPropagation(); editTask(${task.id}); closeInteractiveTaskPopup();" 
                            style="padding: 10px 20px;">Edit Task</button>
                    ${currentUser.role === 'admin' ? `
                        <button class="btn btn-info" onclick="event.stopPropagation(); copyTask(${task.id}); closeInteractiveTaskPopup();" 
                                style="padding: 10px 20px; background: #17a2b8; color: white; border: none;">Copy Task</button>
                        <button class="btn btn-danger" onclick="event.stopPropagation(); if(confirm('Are you sure you want to delete this task?')) { deleteTask(${task.id}); closeInteractiveTaskPopup(); renderDashboard(); renderInteractiveDashboard(); }" 
                                style="padding: 10px 20px;">Delete Task</button>
                    ` : ''}
                </div>
            </div>
        `;
    } else if (task.task_action === 'completed' && currentUser.role === 'admin' && !task.admin_finalized) {
        actionsHtml = `
            <div style="margin-top: 20px; padding-top: 20px; border-top: 1px solid #e0e0e0;">
                <h3 style="margin-bottom: 15px; color: #333;">Admin Actions</h3>
                <div style="display: flex; gap: 10px; flex-wrap: wrap;">
                    <button class="btn btn-success" onclick="event.stopPropagation(); acceptTaskCompletion(${task.id}); closeInteractiveTaskPopup(); renderInteractiveDashboard();" 
                            style="padding: 10px 20px;">✓ Accept</button>
                    <button class="btn btn-danger" onclick="event.stopPropagation(); rejectTaskCompletion(${task.id}); closeInteractiveTaskPopup(); renderInteractiveDashboard();" 
                            style="padding: 10px 20px;">✗ Reject</button>
                    <button class="btn btn-primary" onclick="event.stopPropagation(); editTask(${task.id}); closeInteractiveTaskPopup();" 
                            style="padding: 10px 20px;">Edit Task</button>
                    <button class="btn btn-info" onclick="event.stopPropagation(); copyTask(${task.id}); closeInteractiveTaskPopup();" 
                            style="padding: 10px 20px; background: #17a2b8; color: white; border: none;">Copy Task</button>
                    <button class="btn btn-danger" onclick="event.stopPropagation(); if(confirm('Are you sure you want to delete this task?')) { deleteTask(${task.id}); closeInteractiveTaskPopup(); renderInteractiveDashboard(); }" 
                            style="padding: 10px 20px;">Delete Task</button>
                </div>
            </div>
        `;
    } else {
        // Done (completed or need improvement, no pending admin review): Edit, Copy, Delete
        actionsHtml = `
            <div style="margin-top: 20px; padding-top: 20px; border-top: 1px solid #e0e0e0;">
                <h3 style="margin-bottom: 15px; color: #333;">Actions</h3>
                <div style="display: flex; gap: 10px; flex-wrap: wrap;">
                    <button class="btn btn-primary" onclick="event.stopPropagation(); editTask(${task.id}); closeInteractiveTaskPopup();" 
                            style="padding: 10px 20px;">Edit Task</button>
                    <button class="btn btn-info" onclick="event.stopPropagation(); copyTask(${task.id}); closeInteractiveTaskPopup();" 
                            style="padding: 10px 20px; background: #17a2b8; color: white; border: none;">Copy Task</button>
                    ${currentUser.role === 'admin' ? `
                        <button class="btn btn-danger" onclick="event.stopPropagation(); if(confirm('Are you sure you want to delete this task?')) { deleteTask(${task.id}); closeInteractiveTaskPopup(); renderInteractiveDashboard(); }" 
                                style="padding: 10px 20px;">Delete Task</button>
                    ` : ''}
                </div>
            </div>
        `;
    }

    // Build content HTML
    const dataForRecurrence = data;
    const recurrenceNumber = getTaskRecurrenceNumber(task, dataForRecurrence);
    const recurrenceLabel = recurrenceNumber != null ? ` (R${recurrenceNumber})` : '';

    const contentHtml = `
        <div style="line-height: 1.8;">
            <div style="margin-bottom: 20px;">
                <h3 style="color: #667eea; margin-bottom: 10px;">${task.task_name}</h3>
                ${task.task_number != null ? `<div style="margin-bottom: 8px; font-size: 12px; color: #666;">Task #${task.task_number}${recurrenceLabel}${task.task_type === 'recurring' && (task.due_date || task.next_due_date) ? ' - ' + formatDateDisplay(task.due_date || task.next_due_date) : ''}</div>` : ''}
                <div style="display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 15px;">
                    <span class="badge ${statusClass}">${statusText}</span>
                    ${task.priority ? `<span class="badge badge-${task.priority}">${task.priority.toUpperCase()}</span>` : ''}
                    ${task.frequency ? `<span class="badge badge-recurring">${task.frequency.charAt(0).toUpperCase() + task.frequency.slice(1)}</span>` : ''}
                    ${task.is_team_task ? `<span class="badge badge-info">Team Task</span>` : ''}
                </div>
            </div>
            
            <div style="margin-bottom: 15px;">
                <strong style="color: #333;">Description:</strong>
                <div style="margin-top: 5px; padding: 10px; background: #f9f9f9; border-radius: 5px; color: #666; white-space: pre-wrap; word-break: break-word;">${escapeHtml((task.description || 'No description provided').toString())}</div>
            </div>
            
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-bottom: 15px;">
                <div>
                    <strong style="color: #333;">Assigned To:</strong>
                    <div style="color: #666;">${assignedUser ? assignedUser.name : 'Unknown'}</div>
                </div>
                <div>
                    <strong style="color: #333;">Location:</strong>
                    <div style="color: #666;">${location ? location.name : 'Unknown'}</div>
                </div>
                <div>
                    <strong style="color: #333;">Due Date:</strong>
                    <div style="color: #666;">${dueDateStr}</div>
                </div>
                ${task.expected_completion_date ? `
                <div>
                    <strong style="color: #333;">Expected Completion:</strong>
                    <div style="color: #17a2b8;">${formatDateDisplay(task.expected_completion_date)}</div>
                </div>
                ` : ''}
                <div>
                    <strong style="color: #333;">Task Type:</strong>
                    <div style="color: #666;">${task.task_type ? task.task_type.replace('_', ' ').replace(/\b\w/g, l => l.toUpperCase()) : 'N/A'}</div>
                </div>
                ${task.est_minutes ? `
                <div>
                    <strong style="color: #333;">Estimated Minutes:</strong>
                    <div style="color: #666;">${task.est_minutes} minutes</div>
                </div>
                ` : ''}
                ${segregation ? `
                <div>
                    <strong style="color: #333;">Segregation Type:</strong>
                    <div style="color: #666;">${segregation.name}</div>
                </div>
                ` : ''}
            </div>
            
            ${!isTaskCompleted(task) && task.rejected_at && task.admin_comment ? `
            <div style="margin-bottom: 15px; padding: 14px; background: linear-gradient(135deg, #fff8e1, #ffecb3); border-radius: 8px; border-left: 4px solid #ff9800;">
                <strong style="color: #e65100;">Returned — please review and resubmit</strong>
                <div style="margin-top: 10px; color: #5d4037; white-space: pre-wrap;">${escapeHtml(task.admin_comment)}</div>
                ${task.previous_submission_comment ? `<div style="margin-top: 10px; font-size: 12px; color: #795548;">Your previous completion note: ${escapeHtml(task.previous_submission_comment)}</div>` : ''}
            </div>
            ` : ''}
            
            ${task.comment ? `
            <div style="margin-bottom: 15px;">
                <strong style="color: #333;">Completion Remark:</strong>
                <div style="margin-top: 5px; padding: 10px; background: #e3f2fd; border-radius: 5px; color: #666;">
                    ${escapeHtml(task.comment)}
                </div>
                ${task.completed_at ? `
                    <div style="margin-top: 5px; font-size: 12px; color: #999;">
                        Completed on: ${formatDateDisplay(task.completed_at)}
                    </div>
                ` : ''}
            </div>
            ` : ''}
            
            ${task.admin_comment && (isTaskCompleted(task) || task.admin_finalized || !task.rejected_at) ? `
            <div style="margin-bottom: 15px;">
                <strong style="color: #333;">Admin Comment:</strong>
                <div style="margin-top: 5px; padding: 10px; background: ${task.admin_finalized ? '#d4edda' : '#f8d7da'}; border-radius: 5px; color: #666;">
                    ${escapeHtml(task.admin_comment)}
                </div>
                ${task.finalized_at ? `
                    <div style="margin-top: 5px; font-size: 12px; color: #999;">
                        Finalized on: ${formatDateDisplay(task.finalized_at)}
                    </div>
                ` : task.rejected_at ? `
                    <div style="margin-top: 5px; font-size: 12px; color: #999;">
                        Rejected on: ${formatDateDisplay(task.rejected_at)}
                    </div>
                ` : ''}
            </div>
            ` : ''}
            
            ${actionsHtml}
        </div>
    `;

    document.getElementById('interactiveTaskPopupTitle').textContent = task.task_name;
    document.getElementById('interactiveTaskPopupContent').innerHTML = contentHtml;
    document.getElementById('interactiveTaskPopup').classList.add('active');
}

function closeInteractiveTaskPopup() {
    document.getElementById('interactiveTaskPopup').classList.remove('active');
}

function saveQuickTask(event) {
    event.preventDefault();
    const data = getData();

    const taskName = document.getElementById('quickTaskName').value.trim();
    const description = document.getElementById('quickTaskDescription').value.trim();
    const dueDate = document.getElementById('quickTaskDueDate').value;

    if (!taskName) {
        alert('Task Name is required.');
        return;
    }

    if (!dueDate) {
        alert('Due Date is required.');
        return;
    }

    // Get default location (first location or null)
    const defaultLocation = data.locations.length > 0 ? data.locations[0].id : null;

    if (!defaultLocation) {
        alert('No location available. Please add a location in Settings first.');
        return;
    }

    updateData(data => {
        const task = {
            id: Date.now(),
            task_number: getNextTaskNumberFromData(data),
            task_name: taskName,
            description: description,
            assigned_to: currentUser.id,
            location_id: defaultLocation,
            task_type: 'one_time',
            due_date: dueDate,
            priority: 'medium',
            frequency: null,
            due_date_type: null,
            due_day: null,
            start_date: null,
            recurrence_type: null,
            recurrence_interval: null,
            next_due_date: null,
            segregation_type_id: null,
            est_minutes: null,
            is_team_task: false,
            task_action: 'not_completed',
            comment: null,
            recurrence_stopped: false,
            created_by: currentUser.id,
            created_at: new Date().toISOString(),
            completed_at: null
        };
        data.tasks.push(task);
    });

    closeQuickTaskModal();
    renderTasks();
    renderDashboard();
    renderCalendar();
    renderInteractiveDashboard();

    // Show success message
    const messageDiv = document.getElementById('taskUploadMessage');
    if (messageDiv) {
        messageDiv.innerHTML = '<div style="color: green; padding: 10px; background: #d4edda; border-radius: 5px; margin-bottom: 10px;">✓ Quick task created successfully!</div>';
        setTimeout(() => {
            messageDiv.innerHTML = '';
        }, 3000);
    }
}

// Calculate the Nth working day of a given month
// For example: getNthWorkingDayOfMonth(2026, 1, 4) returns the 4th working day of February 2026
// Working days are Monday-Friday, excluding holidays
function getNthWorkingDayOfMonth(year, month, nthDay) {
    const data = getData();
    const holidays = data.holidays.map(h => h.date);

    // Start from the 1st day of the month
    let currentDate = new Date(year, month, 1);
    currentDate.setHours(0, 0, 0, 0);

    let workingDayCount = 0;
    const maxDaysInMonth = new Date(year, month + 1, 0).getDate();

    // Loop through the days of the month
    for (let day = 1; day <= maxDaysInMonth; day++) {
        currentDate.setDate(day);
        const dayOfWeek = currentDate.getDay();
        const dateStr = formatDateString(currentDate);

        // Check if it's a working day (Mon-Fri and not a holiday)
        if (dayOfWeek !== 0 && dayOfWeek !== 6 && !holidays.includes(dateStr)) {
            workingDayCount++;

            // If we've reached the Nth working day, return this date
            if (workingDayCount === nthDay) {
                return new Date(currentDate);
            }
        }
    }

    // If the month doesn't have enough working days, return the last working day of the month
    // This handles edge cases like requesting the 25th working day of a short month
    return currentDate;
}

function calculateRecurringDueDate(startDate, frequency, dueDateType, dueDay) {
    // Parse date string to avoid timezone issues
    const startParts = startDate.split('-');
    const start = new Date(parseInt(startParts[0]), parseInt(startParts[1]) - 1, parseInt(startParts[2]));
    const today = new Date();
    today.setHours(0, 0, 0, 0, 0);

    let targetDate = new Date(start);
    targetDate.setHours(0, 0, 0, 0);

    // For monthly, quarterly, halfyearly, and yearly frequencies
    if (frequency === 'monthly' || frequency === 'quarterly' || frequency === 'halfyearly' || frequency === 'yearly') {
        // If using working day mode, calculate the Nth working day of the month
        if (dueDateType === 'working_day') {
            targetDate = getNthWorkingDayOfMonth(targetDate.getFullYear(), targetDate.getMonth(), dueDay);
        } else if (dueDateType === 'last_working_day') {
            targetDate = getLastWorkingDayOfMonth(targetDate.getFullYear(), targetDate.getMonth());
        } else {
            // Calendar day mode: just set to the Nth day of the month
            const lastDayOfMonth = new Date(targetDate.getFullYear(), targetDate.getMonth() + 1, 0).getDate();
            const dayToUse = Math.min(dueDay, lastDayOfMonth);
            targetDate.setDate(dayToUse);
        }
    }

    // If target date is in the past, move to next occurrence
    while (targetDate < today) {
        if (frequency === 'daily') {
            targetDate.setDate(targetDate.getDate() + 1);
        } else if (frequency === 'weekly') {
            targetDate.setDate(targetDate.getDate() + 7);
        } else if (frequency === 'monthly') {
            targetDate.setMonth(targetDate.getMonth() + 1);
            // Recalculate based on working day or calendar day
            if (dueDateType === 'working_day') {
                targetDate = getNthWorkingDayOfMonth(targetDate.getFullYear(), targetDate.getMonth(), dueDay);
            } else if (dueDateType === 'last_working_day') {
                targetDate = getLastWorkingDayOfMonth(targetDate.getFullYear(), targetDate.getMonth());
            } else {
                const lastDayOfMonth = new Date(targetDate.getFullYear(), targetDate.getMonth() + 1, 0).getDate();
                const dayToUse = Math.min(dueDay, lastDayOfMonth);
                targetDate.setDate(dayToUse);
            }
        } else if (frequency === 'quarterly') {
            targetDate.setMonth(targetDate.getMonth() + 3);
            if (dueDateType === 'working_day') {
                targetDate = getNthWorkingDayOfMonth(targetDate.getFullYear(), targetDate.getMonth(), dueDay);
            } else if (dueDateType === 'last_working_day') {
                targetDate = getLastWorkingDayOfMonth(targetDate.getFullYear(), targetDate.getMonth());
            } else {
                const lastDayOfMonth = new Date(targetDate.getFullYear(), targetDate.getMonth() + 1, 0).getDate();
                const dayToUse = Math.min(dueDay, lastDayOfMonth);
                targetDate.setDate(dayToUse);
            }
        } else if (frequency === 'halfyearly') {
            targetDate.setMonth(targetDate.getMonth() + 6);
            if (dueDateType === 'working_day') {
                targetDate = getNthWorkingDayOfMonth(targetDate.getFullYear(), targetDate.getMonth(), dueDay);
            } else if (dueDateType === 'last_working_day') {
                targetDate = getLastWorkingDayOfMonth(targetDate.getFullYear(), targetDate.getMonth());
            } else {
                const lastDayOfMonth = new Date(targetDate.getFullYear(), targetDate.getMonth() + 1, 0).getDate();
                const dayToUse = Math.min(dueDay, lastDayOfMonth);
                targetDate.setDate(dayToUse);
            }
        } else if (frequency === 'yearly') {
            targetDate.setFullYear(targetDate.getFullYear() + 1);
            if (dueDateType === 'working_day') {
                targetDate = getNthWorkingDayOfMonth(targetDate.getFullYear(), targetDate.getMonth(), dueDay);
            } else {
                const lastDayOfMonth = new Date(targetDate.getFullYear(), targetDate.getMonth() + 1, 0).getDate();
                const dayToUse = Math.min(dueDay, lastDayOfMonth);
                targetDate.setDate(dayToUse);
            }
        }
    }

    // For daily and weekly frequencies, adjust to working day if needed
    // (This ensures if a daily/weekly task falls on a weekend, it moves to Monday)
    if ((frequency === 'daily' || frequency === 'weekly') && dueDateType === 'working_day') {
        targetDate = adjustToWorkingDay(targetDate);
    }

    // Format date as YYYY-MM-DD to avoid timezone issues
    const year = targetDate.getFullYear();
    const month = String(targetDate.getMonth() + 1).padStart(2, '0');
    const day = String(targetDate.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function adjustToWorkingDay(date) {
    const data = getData();
    const holidays = data.holidays.map(h => h.date);

    let checkDate = new Date(date);
    checkDate.setHours(0, 0, 0, 0); // Reset time to midnight
    let attempts = 0;

    while (attempts < 30) { // Max 30 days forward
        const dayOfWeek = checkDate.getDay();
        const dateStr = formatDateString(checkDate);

        // Check if it's a weekend (Sunday = 0, Saturday = 6)
        // If it's Saturday or Sunday, move to next day
        if (dayOfWeek === 0 || dayOfWeek === 6) {
            checkDate.setDate(checkDate.getDate() + 1);
            attempts++;
            continue;
        }

        // Check if it's a holiday
        // If it's a company holiday, move to next day
        if (holidays.includes(dateStr)) {
            checkDate.setDate(checkDate.getDate() + 1);
            attempts++;
            continue;
        }

        // It's a working day (Monday-Friday and not a holiday)
        return checkDate;
    }

    return date; // Fallback to original date if no working day found in 30 days
}

function calculateNextRecurrenceDate(task) {
    if (!task.frequency || task.recurrence_stopped) return null;

    const lastDateStr = task.due_date || task.next_due_date || task.completed_at || formatDateString(new Date());
    const lastDateParts = lastDateStr.split('-');
    const lastDate = new Date(parseInt(lastDateParts[0]), parseInt(lastDateParts[1]) - 1, parseInt(lastDateParts[2]));
    lastDate.setHours(0, 0, 0, 0);

    let nextDate = new Date(lastDate);
    const dueDay = task.due_day || 1;
    const dueDateType = task.due_date_type || 'calendar_day';

    switch (task.frequency) {
        case 'daily':
            nextDate.setDate(nextDate.getDate() + 1);
            break;
        case 'weekly':
            nextDate.setDate(nextDate.getDate() + 7);
            break;
        case 'monthly':
            nextDate.setMonth(nextDate.getMonth() + 1);
            // Calculate based on working day or calendar day
            if (dueDateType === 'working_day') {
                nextDate = getNthWorkingDayOfMonth(nextDate.getFullYear(), nextDate.getMonth(), dueDay);
            } else {
                const lastDayOfMonth = new Date(nextDate.getFullYear(), nextDate.getMonth() + 1, 0).getDate();
                const dayToUse = Math.min(dueDay, lastDayOfMonth);
                nextDate.setDate(dayToUse);
            }
            break;
        case 'yearly':
            nextDate.setFullYear(nextDate.getFullYear() + 1);
            if (dueDateType === 'working_day') {
                nextDate = getNthWorkingDayOfMonth(nextDate.getFullYear(), nextDate.getMonth(), dueDay);
            } else {
                const lastDayOfYearMonth = new Date(nextDate.getFullYear(), nextDate.getMonth() + 1, 0).getDate();
                const dayToUseYear = Math.min(dueDay, lastDayOfYearMonth);
                nextDate.setDate(dayToUseYear);
            }
            break;
        default:
            return null;
    }

    // For daily and weekly frequencies, adjust to working day if needed (skip weekends and holidays)
    if ((task.frequency === 'daily' || task.frequency === 'weekly') && dueDateType === 'working_day') {
        nextDate = adjustToWorkingDay(nextDate);
    }

    return formatDateString(nextDate);
}

// Calculate next recurrence date for a completed recurring task instance
function calculateNextRecurrenceDateForInstance(task, currentDueDateStr) {
    if (!task.frequency || task.recurrence_stopped) return null;

    const dateParts = currentDueDateStr.split('-');
    const currentDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
    currentDate.setHours(0, 0, 0, 0);

    let nextDate = new Date(currentDate);
    const dueDay = task.due_day || 1;
    const dueDateType = task.due_date_type || 'calendar_day';

    switch (task.frequency) {
        case 'daily':
            nextDate.setDate(nextDate.getDate() + 1);
            break;
        case 'weekly':
            nextDate.setDate(nextDate.getDate() + 7);
            break;
        case 'monthly':
            nextDate.setMonth(nextDate.getMonth() + 1);
            // Calculate based on working day or calendar day
            if (dueDateType === 'working_day') {
                nextDate = getNthWorkingDayOfMonth(nextDate.getFullYear(), nextDate.getMonth(), dueDay);
            } else if (dueDateType === 'last_working_day') {
                nextDate = getLastWorkingDayOfMonth(nextDate.getFullYear(), nextDate.getMonth());
            } else {
                const lastDayOfMonth = new Date(nextDate.getFullYear(), nextDate.getMonth() + 1, 0).getDate();
                const dayToUse = Math.min(dueDay, lastDayOfMonth);
                nextDate.setDate(dayToUse);
            }
            break;
        case 'yearly':
            nextDate.setFullYear(nextDate.getFullYear() + 1);
            if (dueDateType === 'working_day') {
                nextDate = getNthWorkingDayOfMonth(nextDate.getFullYear(), nextDate.getMonth(), dueDay);
            } else if (dueDateType === 'last_working_day') {
                nextDate = getLastWorkingDayOfMonth(nextDate.getFullYear(), nextDate.getMonth());
            } else {
                const lastDayOfYearMonth = new Date(nextDate.getFullYear(), nextDate.getMonth() + 1, 0).getDate();
                const dayToUseYear = Math.min(dueDay, lastDayOfYearMonth);
                nextDate.setDate(dayToUseYear);
            }
            break;
        default:
            return null;
    }

    // For daily and weekly frequencies, adjust to working day if needed
    if ((task.frequency === 'daily' || task.frequency === 'weekly') && dueDateType === 'working_day') {
        const adjustedDate = adjustToWorkingDay(nextDate);
        return formatDateString(adjustedDate);
    }

    return formatDateString(nextDate);
}

// Get the last working day (Mon–Fri, excluding holidays) of a given month
function getLastWorkingDayOfMonth(year, monthIndex) {
    const data = getData();
    const holidays = (data.holidays || []).map(h => h.date);
    let date = new Date(year, monthIndex + 1, 0); // last calendar day
    date.setHours(0, 0, 0, 0);

    while (true) {
        const dayOfWeek = date.getDay(); // 0 Sun ... 6 Sat
        const dateStr = formatDateString(date);
        if (dayOfWeek !== 0 && dayOfWeek !== 6 && !holidays.includes(dateStr)) {
            return new Date(date);
        }
        date.setDate(date.getDate() - 1);
    }
}

// Recurring Tasks Performance Report
function renderRecurringReport() {
    const data = getData();
    const fromInput = document.getElementById('recurringReportFromMonth');
    const toInput = document.getElementById('recurringReportToMonth');
    const container = document.getElementById('recurringReportContainer');
    if (!fromInput || !toInput || !container) return;

    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth() + 1;

    // Initialise filters if empty
    if (!fromInput.value) {
        fromInput.value = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;
    }
    if (!toInput.value) {
        toInput.value = fromInput.value;
    }

    const [fromYear, fromMonth] = fromInput.value.split('-').map(Number);
    const [toYear, toMonth] = toInput.value.split('-').map(Number);
    if (isNaN(fromYear) || isNaN(fromMonth) || isNaN(toYear) || isNaN(toMonth)) {
        container.innerHTML = '<p style="text-align: center; color: #999; padding: 20px;">Please select valid From and To months.</p>';
        return;
    }

    // Build list of months in range
    const months = [];
    let y = fromYear;
    let m = fromMonth - 1; // zero-based
    const end = new Date(toYear, toMonth - 1, 1);
    while (new Date(y, m, 1) <= end) {
        months.push({ year: y, monthIndex: m });
        m++;
        if (m > 11) {
            m = 0;
            y++;
        }
    }

    if (months.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #999; padding: 20px;">Please select a valid month range.</p>';
        return;
    }

    // Group recurring tasks by base definition (tasks with start_date are treated as base)
    const baseRecurring = data.tasks.filter(t =>
        t.task_type === 'recurring' &&
        t.start_date &&
        (currentUser.role === 'admin' || t.assigned_to === currentUser.id)
    );

    if (baseRecurring.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #999; padding: 20px;">No recurring tasks found.</p>';
        return;
    }

    const monthHeaders = months.map(({ year, monthIndex }) => {
        const d = new Date(year, monthIndex, 1);
        const label = d.toLocaleString(undefined, { month: 'short', year: 'numeric' });
        return `<th style="white-space: nowrap;">${label}</th>`;
    }).join('');

    const rowsHtml = baseRecurring.map(baseTask => {
        const assignedUser = data.users.find(u => u.id === baseTask.assigned_to);
        const seriesTasks = data.tasks.filter(t =>
            t.task_type === 'recurring' &&
            t.task_name === baseTask.task_name &&
            t.assigned_to === baseTask.assigned_to &&
            t.location_id === baseTask.location_id &&
            t.frequency === baseTask.frequency
        );

        const cells = months.map(({ year, monthIndex }) => {
            const instance = findRecurringInstanceForMonth(seriesTasks, year, monthIndex);
            if (!instance) {
                return '<td style="text-align: center; color: #ccc; background: #f5f5f5;">–</td>';
            }

            let statusText = '';
            let bgColor = '#ffffff';
            let textColor = '#000000';
            if (instance.task_action === 'completed') {
                statusText = 'Completed';
                bgColor = '#b7e1cd'; // light green
            } else if (instance.task_action === 'completed_need_improvement') {
                statusText = 'Needs Improvement';
                bgColor = '#ffe599'; // light orange
            } else if (instance.task_action === 'in_process') {
                statusText = 'In Process';
                bgColor = '#d1ecf1'; // light blue
            } else if (instance.task_action === 'not_done') {
                statusText = 'Not Done';
                bgColor = '#ff4d4f'; // red
                textColor = '#ffffff';
            } else {
                // Pending / Overdue
                const dueStr = instance.due_date || instance.next_due_date;
                if (dueStr) {
                    const [yy, mm, dd] = dueStr.split('-').map(Number);
                    const d = new Date(yy, mm - 1, dd);
                    d.setHours(0, 0, 0, 0);
                    const today0 = new Date();
                    today0.setHours(0, 0, 0, 0);
                    if (d < today0) {
                        statusText = 'Overdue';
                        bgColor = '#f8d7da'; // light red
                    } else {
                        statusText = 'Pending';
                        bgColor = '#fff3cd'; // yellow
                    }
                } else {
                    statusText = 'Pending';
                    bgColor = '#fff3cd';
                }
            }

            let completionPart = '';
            if (isTaskCompleted(instance)) {
                const cd = instance.completion_date ||
                    (instance.completed_at ? instance.completed_at.split('T')[0] : '');
                if (cd) {
                    completionPart = ` (${formatDateDisplay(cd)})`;
                }
            }

            return `<td style="font-size: 12px; background: ${bgColor}; color: ${textColor};">${statusText}${completionPart}</td>`;
        }).join('');

        return `
            <tr>
                <td style="font-weight: 500; white-space: nowrap;">${baseTask.task_name}</td>
                <td style="font-size: 12px;">${assignedUser ? assignedUser.name : ''}</td>
                ${cells}
            </tr>
        `;
    }).join('');

    container.innerHTML = `
        <div class="table-container" style="background: white; border-radius: 8px; overflow: auto; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
            <table class="table" id="recurringReportTable" style="margin: 0; min-width: 600px; border-collapse: collapse;">
                <thead>
                    <tr>
                        <th>Task Name</th>
                        <th>Assigned To</th>
                        ${monthHeaders}
                    </tr>
                </thead>
                <tbody>
                    ${rowsHtml}
                </tbody>
            </table>
        </div>
    `;
}

function findRecurringInstanceForMonth(seriesTasks, year, monthIndex) {
    const matches = seriesTasks.filter(t => {
        const dueStr = t.due_date || t.next_due_date;
        if (!dueStr) return false;
        const [yy, mm] = dueStr.split('-').map(Number);
        return yy === year && (mm - 1) === monthIndex;
    });

    if (matches.length === 0) return null;

    // If multiple, take the earliest due date in that month
    matches.sort((a, b) => {
        const ad = (a.due_date || a.next_due_date || '9999-12-31');
        const bd = (b.due_date || b.next_due_date || '9999-12-31');
        return ad.localeCompare(bd);
    });
    return matches[0];
}

// Export recurring dashboard to Excel (HTML table with preserved styles)
function exportRecurringReportToExcel() {
    try {
        // Ensure latest view
        renderRecurringReport();

        const container = document.getElementById('recurringReportContainer');
        if (!container || !container.innerHTML.trim()) {
            alert('No data to export. Please ensure the dashboard is loaded.');
            return;
        }

        const html = `
            <html>
                <head>
                    <meta charset="UTF-8">
                    <style>
                        table, th, td { border: 1px solid #ddd; }
                        th { background: #667eea; color: #ffffff; }
                    </style>
                </head>
                <body>
                    ${container.innerHTML}
                </body>
            </html>
        `;

        const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `recurring-task-dashboard-${formatDateString(new Date())}.xls`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    } catch (err) {
        console.error('Failed to export recurring report', err);
        alert('Unable to export recurring task dashboard. Please try again.');
    }
}

// Utility function to recalculate due dates for all existing recurring tasks
// This is useful after fixing the working day calculation logic
function recalculateAllRecurringDueDates() {
    const data = getData();
    let updatedCount = 0;

    console.log('Starting recalculation of recurring tasks...');
    console.log('This will ONLY update due dates, NOT the due_day settings');

    updateData(data => {
        data.tasks.forEach(task => {
            // Only process recurring tasks with working_day type that are active
            if (task.task_type === 'recurring' && task.frequency && (task.due_date_type === 'working_day' || task.due_date_type === 'last_working_day')) {

                // Skip if completed or stopped
                if (task.task_action === 'completed' || task.recurrence_stopped) {
                    return;
                }

                // Validate required fields
                if (!task.due_day) {
                    console.log(`Skipping task "${task.task_name}" - missing due_day`);
                    return;
                }

                // Store the ORIGINAL due_day - we will NOT change this
                const originalDueDay = task.due_day;
                const oldDueDate = task.next_due_date || task.due_date;

                // Recalculate the next due date using the existing logic
                let newDueDate;
                if (task.start_date) {
                    // Use start_date to recalculate
                    newDueDate = calculateRecurringDueDate(
                        task.start_date,
                        task.frequency,
                        task.due_date_type,
                        task.due_day
                    );
                } else if (oldDueDate) {
                    // No start_date, use old due date to calculate next occurrence
                    newDueDate = calculateRecurringDueDate(
                        oldDueDate,
                        task.frequency,
                        task.due_date_type,
                        task.due_day
                    );
                } else {
                    console.log(`Skipping task "${task.task_name}" - no start_date or due_date`);
                    return;
                }

                console.log(`Task "${task.task_name}": DueDay=${task.due_day}, Old=${oldDueDate}, New=${newDueDate}`);

                // VERIFY that due_day hasn't changed (safety check)
                if (task.due_day !== originalDueDay) {
                    console.error(`ERROR: due_day was modified! This should never happen!`);
                    task.due_day = originalDueDay; // Restore it
                }

                // Only update if the date changed
                if (newDueDate !== oldDueDate) {
                    if (task.next_due_date) {
                        task.next_due_date = newDueDate;
                    } else {
                        task.due_date = newDueDate;
                    }
                    updatedCount++;
                    console.log(`✓ Updated task "${task.task_name}" from ${oldDueDate} to ${newDueDate} (DueDay=${task.due_day} unchanged)`);
                }
            }
        });
    });

    console.log(`Recalculation complete. Updated ${updatedCount} task(s).`);
    console.log('Note: The due_day field was NOT modified for any tasks.');

    // Refresh all views
    renderDashboard();
    renderTasks();
    renderCalendar();
    renderInteractiveDashboard();

    return updatedCount;
}

// UI handler for the recalculate button
function recalculateExistingTasks() {
    if (confirm('This will recalculate the due dates for all existing recurring tasks. Continue?')) {
        const updatedCount = recalculateAllRecurringDueDates();
        const messageEl = document.getElementById('recalculateMessage');
        if (messageEl) {
            if (updatedCount > 0) {
                messageEl.textContent = `✅ Successfully recalculated ${updatedCount} recurring task(s)!`;
                messageEl.style.color = '#28a745';
            } else {
                messageEl.textContent = '✅ All recurring tasks are already up-to-date.';
                messageEl.style.color = '#666';
            }
            // Clear message after 5 seconds
            setTimeout(() => {
                messageEl.textContent = '';
            }, 5000);
        }
    }
}

// Process recurring tasks to generate future instances
function processRecurringTasks() {
    const data = getData();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayStr = formatDateString(today);
    const futureLimit = new Date();
    futureLimit.setMonth(futureLimit.getMonth() + 12); // Generate 12 months ahead

    updateData(data => {
        // First, mark all tasks with recurrence_stopped flag if they should be stopped
        data.tasks.forEach(task => {
            if (task.task_type === 'recurring' && task.recurrence_stopped) {
                // Remove any future instances that shouldn't exist
                const futureInstances = data.tasks.filter(t =>
                    t.id !== task.id &&
                    t.task_name === task.task_name &&
                    t.assigned_to === task.assigned_to &&
                    t.frequency === task.frequency &&
                    (t.next_due_date || t.due_date) &&
                    new Date(t.next_due_date || t.due_date) > new Date()
                );
                futureInstances.forEach(instance => {
                    const index = data.tasks.findIndex(t => t.id === instance.id);
                    if (index !== -1) {
                        data.tasks.splice(index, 1);
                    }
                });
            }
        });

        // Get all base recurring tasks (have start_date, indicating they're the original)
        const baseRecurringTasks = data.tasks.filter(t =>
            t.task_type === 'recurring' &&
            t.frequency &&
            !t.recurrence_stopped &&
            t.start_date // Base tasks have start_date
        );

        baseRecurringTasks.forEach(task => {
            const dueDay = task.due_day || 1;
            const dueDateType = task.due_date_type || 'calendar_day';
            let currentDate = new Date(task.start_date || task.due_date || today);

            // Generate instances for next 12 months
            for (let i = 0; i < 12; i++) {
                let nextDueDate = null;

                // Calculate next due date based on frequency
                if (task.frequency === 'daily') {
                    currentDate.setDate(currentDate.getDate() + 1);
                    nextDueDate = new Date(currentDate);
                } else if (task.frequency === 'weekly') {
                    currentDate.setDate(currentDate.getDate() + 7);
                    nextDueDate = new Date(currentDate);
                } else if (task.frequency === 'monthly') {
                    currentDate.setMonth(currentDate.getMonth() + 1);
                    if (dueDateType === 'last_working_day') {
                        nextDueDate = getLastWorkingDayOfMonth(currentDate.getFullYear(), currentDate.getMonth());
                        currentDate = new Date(nextDueDate);
                    } else {
                        const lastDayOfMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0).getDate();
                        const dayToUse = Math.min(dueDay, lastDayOfMonth);
                        currentDate.setDate(dayToUse);
                        nextDueDate = new Date(currentDate);
                    }
                } else if (task.frequency === 'yearly') {
                    currentDate.setFullYear(currentDate.getFullYear() + 1);
                    if (dueDateType === 'last_working_day') {
                        nextDueDate = getLastWorkingDayOfMonth(currentDate.getFullYear(), currentDate.getMonth());
                        currentDate = new Date(nextDueDate);
                    } else {
                        const lastDayOfMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0).getDate();
                        const dayToUse = Math.min(dueDay, lastDayOfMonth);
                        currentDate.setDate(dayToUse);
                        nextDueDate = new Date(currentDate);
                    }
                }

                if (!nextDueDate || nextDueDate > futureLimit) break;

                // Adjust for working day if needed
                if (dueDateType === 'working_day') {
                    nextDueDate = adjustToWorkingDay(nextDueDate);
                } else if (dueDateType === 'last_working_day') {
                    nextDueDate = getLastWorkingDayOfMonth(nextDueDate.getFullYear(), nextDueDate.getMonth());
                }

                const nextDateStr = formatDateString(nextDueDate);
                const nextDateObj = new Date(nextDueDate);
                nextDateObj.setHours(0, 0, 0, 0);

                // Only create if date is in the future
                if (nextDateObj <= today) continue;

                // Check if instance already exists
                const existingNext = data.tasks.find(t =>
                    t.task_name === task.task_name &&
                    t.assigned_to === task.assigned_to &&
                    t.frequency === task.frequency &&
                    (t.next_due_date === nextDateStr || t.due_date === nextDateStr) &&
                    t.id !== task.id
                );

                if (!existingNext && nextDateObj <= futureLimit) {
                    const newTask = {
                        ...task,
                        id: Math.floor(Date.now() * 1000) + i, // Unique integer ID
                        task_number: task.task_number != null ? task.task_number : getNextTaskNumberFromData(data),
                        task_action: 'not_completed',
                        comment: null,
                        completed_at: null,
                        due_date: null,
                        next_due_date: nextDateStr,
                        start_date: null, // Generated instances don't have start_date
                        created_at: new Date().toISOString(),
                        recurrence_stopped: false
                    };
                    data.tasks.push(newTask);
                }
            }
        });
    });
}

function editTask(taskId) {
    openTaskModal(taskId);
}

function copyTask(taskId) {
    const data = getData();
    const task = data.tasks.find(t => t.id == taskId);
    if (!task) return;

    // Open modal with task data but without taskId (so it creates a new task)
    openTaskModal(null);

    // Pre-fill all fields with the task data
    document.getElementById('taskModalTitle').textContent = 'Copy Task';
    document.getElementById('taskName').value = task.task_name + ' (Copy)';
    document.getElementById('taskDescription').value = task.description || '';
    document.getElementById('taskAssignedTo').value = task.assigned_to;
    document.getElementById('taskLocation').value = task.location_id;
    document.getElementById('taskType').value = task.task_type;
    document.getElementById('taskDueDate').value = task.due_date || '';
    if (task.task_type === 'without_due_date') {
        document.getElementById('taskPriorityNoDue').value = task.priority || 'medium';
    } else {
        document.getElementById('taskPriority').value = task.priority || 'medium';
    }
    document.getElementById('taskFrequency').value = task.frequency || '';
    document.getElementById('taskDueDateType').value = task.due_date_type || 'calendar_day';
    document.getElementById('taskDueDay').value = task.due_day || 1;
    document.getElementById('taskStartDate').value = task.start_date || '';
    document.getElementById('taskRecurrenceType').value = task.recurrence_type || 'calendar_day';
    document.getElementById('taskRecurrenceInterval').value = task.recurrence_interval || 1;
    document.getElementById('taskSegregation').value = task.segregation_type_id || '';
    document.getElementById('taskEstMinutes').value = task.est_minutes || '';
    document.getElementById('taskIsTeam').checked = task.is_team_task || false;
    document.getElementById('taskCommentGroup').style.display = 'none';
    document.getElementById('taskStopRecurrence').checked = false;
    document.getElementById('taskRecurrenceStopped').value = 'false';
    toggleRecurrenceFields();
    updateRecurrenceFields();
}

function deleteTask(taskId) {
    const data = getData();
    const taskToDelete = data.tasks.find(t => t.id == taskId);

    if (!taskToDelete) return;

    // Check if this is a recurring task
    let confirmMessage = 'Are you sure you want to delete this task?';
    let futureInstancesCount = 0;

    if (taskToDelete.task_type === 'recurring') {
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        // Count future instances
        futureInstancesCount = data.tasks.filter(t => {
            if (t.id === taskId) return false; // Don't count the task being deleted
            return t.task_type === 'recurring' &&
                t.task_name === taskToDelete.task_name &&
                t.assigned_to === taskToDelete.assigned_to &&
                t.frequency === taskToDelete.frequency &&
                (t.next_due_date || t.due_date) &&
                new Date(t.next_due_date || t.due_date) >= today &&
                t.task_action !== 'completed';
        }).length;

        if (futureInstancesCount > 0) {
            confirmMessage = `Are you sure you want to delete this recurring task? This will also delete ${futureInstancesCount} future instance(s).`;
        } else {
            confirmMessage = 'Are you sure you want to delete this recurring task?';
        }
    }

    if (!confirm(confirmMessage)) return;

    updateData(data => {
        // If it's a recurring task, delete all related future instances
        if (taskToDelete.task_type === 'recurring') {
            const today = new Date();
            today.setHours(0, 0, 0, 0);

            // Find all related recurring task instances that are in the future
            // They share: task_name, assigned_to, and frequency
            const relatedFutureTasks = data.tasks.filter(t => {
                // Always include the task being deleted
                if (t.id === taskId) return true;

                // Match related recurring tasks that are future instances
                return t.task_type === 'recurring' &&
                    t.task_name === taskToDelete.task_name &&
                    t.assigned_to === taskToDelete.assigned_to &&
                    t.frequency === taskToDelete.frequency &&
                    (t.next_due_date || t.due_date) &&
                    new Date(t.next_due_date || t.due_date) >= today &&
                    t.task_action !== 'completed'; // Don't delete completed past instances
            });

            // Delete all related future tasks (including the one being deleted)
            const taskIdsToDelete = new Set(relatedFutureTasks.map(t => t.id));
            data.tasks = data.tasks.filter(t => !taskIdsToDelete.has(t.id));
        } else {
            // For non-recurring tasks, just delete the single task
            data.tasks = data.tasks.filter(t => t.id !== taskId);
        }
    });

    renderTasks();
    renderDashboard();
    renderCalendar();
    renderInteractiveDashboard();

    // If drilldown is open, refresh it
    if (window.drilldownContext) {
        renderDrilldown();
    }
}

// Calendar - persist task filter (self / team / all)
const CALENDAR_TASK_FILTER_KEY = 'calendarTaskFilter';

function getCalendarTaskFilter() {
    const v = localStorage.getItem(CALENDAR_TASK_FILTER_KEY);
    return (v === 'self' || v === 'team' || v === 'all') ? v : 'self';
}

function saveCalendarTaskFilter() {
    const el = document.getElementById('calendarTaskFilter');
    if (el) localStorage.setItem(CALENDAR_TASK_FILTER_KEY, el.value || 'self');
}

function calendarTaskFilterPredicate(task) {
    const mode = getCalendarTaskFilter();
    if (mode === 'all') return true;
    if (mode === 'self') return task.assigned_to === currentUser.id;
    if (mode === 'team') return task.assigned_to === currentUser.id || task.is_team_task;
    return task.assigned_to === currentUser.id;
}

// Calendar
function renderCalendar() {
    const data = getData();
    const filterEl = document.getElementById('calendarTaskFilter');
    if (filterEl) {
        const saved = getCalendarTaskFilter();
        filterEl.value = saved;
    }

    const firstDay = new Date(currentYear, currentMonth, 1);
    const lastDay = new Date(currentYear, currentMonth + 1, 0);
    const startDate = new Date(firstDay);
    startDate.setDate(startDate.getDate() - startDate.getDay());

    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'];

    document.getElementById('currentMonthYear').textContent =
        `${monthNames[currentMonth]} ${currentYear}`;

    const weekdays = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    let html = '<div class="calendar">';

    weekdays.forEach(day => {
        html += `<div class="weekday-header">${day}</div>`;
    });

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayStr = formatDateString(today);
    const currentDate = new Date(startDate);
    currentDate.setHours(0, 0, 0, 0);

    for (let i = 0; i < 42; i++) {
        const dateStr = formatDateString(currentDate);
        const isCurrentMonth = currentDate.getMonth() === currentMonth;
        const isToday = dateStr === todayStr;

        const dayTasks = data.tasks.filter(task => {
            const dueDate = task.due_date || task.next_due_date;
            if (dueDate !== dateStr) return false;
            return calendarTaskFilterPredicate(task);
        });

        let classes = 'calendar-day';
        if (!isCurrentMonth) classes += ' other-month';
        if (isToday) classes += ' today';
        if (dayTasks.length > 0) classes += ' has-task';

        html += `<div class="${classes}" onclick="showDateTasks('${dateStr}')" style="cursor: pointer;">`;
        html += `<div class="calendar-day-number">${currentDate.getDate()}</div>`;
        dayTasks.slice(0, 3).forEach(task => {
            html += `<div class="calendar-task" title="${task.task_name}">${task.task_name}</div>`;
        });
        if (dayTasks.length > 3) {
            html += `<div class="calendar-task">+${dayTasks.length - 3} more</div>`;
        }
        html += '</div>';

        currentDate.setDate(currentDate.getDate() + 1);
    }

    html += '</div>';
    document.getElementById('calendarView').innerHTML = html;
}

function showDateTasks(dateStr) {
    // Switch to tasks tab and filter by date
    switchTab('tasks', null);

    // Create a temporary filter for this date (use same filter as calendar: self/team/all)
    const data = getData();
    const dateTasks = data.tasks.filter(task => {
        const dueDate = task.due_date || task.next_due_date;
        return dueDate === dateStr && calendarTaskFilterPredicate(task);
    });

    // Parse date for display
    const dateParts = dateStr.split('-');
    const displayDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));

    // Render tasks for this date
    const tasksHtml = dateTasks.length > 0
        ? dateTasks.map(task => renderTaskItem(task, true)).join('')
        : `<p style="text-align: center; color: #999; padding: 20px;">No tasks for ${formatDateDisplay(displayDate)}</p>`;

    document.getElementById('tasksList').innerHTML = `
        <div style="margin-bottom: 20px; padding: 15px; background: #e3f2fd; border-radius: 8px; display: flex; justify-content: space-between; align-items: center;">
            <h3 style="margin: 0;">Tasks for ${formatDateDisplay(displayDate)}</h3>
            <button class="btn btn-secondary" onclick="filterTasks()" style="padding: 5px 15px; font-size: 12px;">Clear Filter</button>
        </div>
        ${tasksHtml}
    `;

    // Clear other filters
    document.getElementById('filterStatus').value = '';
    document.getElementById('filterType').value = '';
    document.getElementById('filterTeam').value = '';
    document.getElementById('searchTasks').value = '';
}

function changeMonth(direction) {
    currentMonth += direction;
    if (currentMonth < 0) {
        currentMonth = 11;
        currentYear--;
    } else if (currentMonth > 11) {
        currentMonth = 0;
        currentYear++;
    }
    renderCalendar();
}

// Users
function renderUsers() {
    if (currentUser.role !== 'admin' && !currentUser.isMaster) return;

    const data = getData();
    const masterReadOnly = isApiMode() && currentUser.isMaster;
    const filterWrap = document.getElementById('masterUserMgmtFilterWrap');
    const filt = document.getElementById('masterUserStatusFilter');
    if (filterWrap) {
        filterWrap.classList.toggle('hidden', !masterReadOnly);
    }
    if (filt && masterReadOnly && filt.getAttribute('data-wired') !== '1') {
        filt.setAttribute('data-wired', '1');
        filt.addEventListener('change', () => renderSettings());
    }

    const headRow = document.getElementById('userMgmtTableHeadRow');
    if (headRow) {
        if (masterReadOnly) {
            headRow.innerHTML = '<th>Name</th><th>Email</th><th>Role</th><th>Assigned org (account admin)</th><th>Status</th><th>Actions</th>';
        } else {
            headRow.innerHTML = '<th>Name</th><th>Email</th><th>Role</th><th>Status</th><th>Actions</th>';
        }
    }

    const mode = masterReadOnly ? getMasterAccountFilterMode() : 'all';
    const listUsers = masterReadOnly
        ? (data.users || []).filter(u => userPassesMasterAccountFilter(u, mode))
        : (data.users || []);

    const colCount = masterReadOnly ? 6 : 5;
    const html = listUsers.map(user => {
        const active = user.is_active !== false;
        const tenantTd = masterReadOnly
            ? `<td>${escapeHtml(user.tenant_admin_label != null && user.tenant_admin_label !== '' ? user.tenant_admin_label : '—')}</td>`
            : '';
        return `
        <tr>
            <td>${escapeHtml(user.name)}</td>
            <td>${escapeHtml(user.email)}</td>
            <td><span class="badge ${user.role === 'admin' ? 'badge-high' : 'badge-low'}">${user.role}</span></td>
            ${tenantTd}
            <td><span class="badge ${active ? 'badge-completed' : 'badge-pending'}">${active ? 'Active' : 'Disabled'}</span></td>
            <td>${masterReadOnly
        ? '<span style="color:#888;font-size:12px;">Use Master password reset below</span>'
        : `<button type="button" class="btn btn-primary" onclick="editUser(${user.id})" style="padding: 5px 10px; font-size: 12px;">Edit</button>
                ${user.id !== currentUser.id ? `<button type="button" class="btn btn-danger" onclick="deleteUser(${user.id})" style="padding: 5px 10px; font-size: 12px;">Delete</button>` : ''}`}
            </td>
        </tr>`;
    }).join('');

    document.getElementById('usersList').innerHTML = html || `<tr><td colspan="${colCount}" style="text-align: center;">No users</td></tr>`;

    const addBtn = document.getElementById('settingsAddUserBtn');
    if (addBtn) {
        addBtn.style.display = masterReadOnly ? 'none' : 'inline-block';
    }
}

function openUserModal(userId = null) {
    const modal = document.getElementById('userModal');
    const form = document.getElementById('userForm');

    if (userId) {
        const data = getData();
        const user = data.users.find(u => u.id === userId);
        if (user) {
            document.getElementById('userModalTitle').textContent = 'Edit User';
            document.getElementById('userId').value = user.id;
            document.getElementById('userName').value = user.name;
            document.getElementById('userEmail').value = user.email;
            document.getElementById('userPassword').required = false;
            document.getElementById('userPasswordLabel').textContent = 'New Password (leave blank to keep current)';
            document.getElementById('userIsActive').checked = user.is_active;
        }
    } else {
        form.reset();
        document.getElementById('userModalTitle').textContent = 'New User';
        document.getElementById('userId').value = '';
        document.getElementById('userPassword').required = true;
        document.getElementById('userPasswordLabel').textContent = 'Password *';
        document.getElementById('userIsActive').checked = true;
    }

    modal.classList.add('active');
}

function closeUserModal() {
    document.getElementById('userModal').classList.remove('active');
}

async function saveUser(event) {
    event.preventDefault();
    const userId = document.getElementById('userId').value;
    const password = document.getElementById('userPassword').value;

    if (!userId && (!password || password.length < 6)) {
        alert('Password is required (min 6 characters) for new users.');
        return;
    }

    updateData(data => {
        if (userId) {
            const user = data.users.find(u => u.id === parseInt(userId));
            if (user) {
                user.name = document.getElementById('userName').value;
                user.email = document.getElementById('userEmail').value;
                if (password) {
                    user.password = password;
                    rememberPendingUserPasswordForSync(parseInt(userId, 10), password);
                }
                user.is_active = document.getElementById('userIsActive').checked;
            }
        } else {
            const newUser = {
                id: Date.now() * 1000 + Math.floor(Math.random() * 1000),
                name: document.getElementById('userName').value,
                email: document.getElementById('userEmail').value,
                password: password,
                role: 'user',
                is_active: document.getElementById('userIsActive').checked
            };
            data.users.push(newUser);
            rememberPendingUserPasswordForSync(newUser.id, password);
        }
    });

    closeUserModal();
    renderUsers();

    if (isApiMode() && currentUser && !currentUser.isMaster) {
        const ok = await flushWorkspaceToApiNow();
        if (!ok) {
            alert(
                __lastWorkspacePutError
                    ? `${__lastWorkspacePutError}\n\nOpen Settings and try again, or check your connection.`
                    : 'Could not save users to the server. Check your connection, then open Settings and save again.'
            );
        }
    }
}

function editUser(userId) {
    openUserModal(userId);
}

async function deleteUser(userId) {
    if (!confirm('Are you sure you want to delete this user?')) return;
    const uid = typeof userId === 'number' ? userId : parseInt(String(userId), 10);
    if (Number.isNaN(uid)) return;

    __pendingPasswordsByUserId.delete(uid);
    updateData(data => {
        data.users = data.users.filter(u => Number(u.id) !== uid);
        data.tasks = data.tasks.filter(t => Number(t.assigned_to) !== uid && Number(t.created_by) !== uid);
    });

    renderUsers();

    if (isApiMode() && currentUser && !currentUser.isMaster) {
        const ok = await flushWorkspaceToApiNow();
        if (!ok) {
            alert(
                __lastWorkspacePutError
                    ? `${__lastWorkspacePutError}\n\nCheck your connection and try again.`
                    : 'Could not sync user deletion to the server. Check your connection.'
            );
        }
    }
}

// Settings
function renderSettings() {
    if (currentUser.role !== 'admin' && !currentUser.isMaster) return;

    const data = getData();

    // Render users
    renderUsers();

    const masterSec = document.getElementById('masterPasswordResetSection');
    if (masterSec) {
        if (isApiMode() && currentUser && currentUser.isMaster) {
            masterSec.style.display = 'block';
            const opts = data.users.map(u =>
                `<option value="${u.id}">${escapeHtml(u.name)} (${escapeHtml(u.email)})</option>`
            ).join('');
            const orgAdminOptsForMaster = (data.users || [])
                .filter(u => u.role === 'admin' && !isMasterUserRecord(u))
                .map(u =>
                    `<option value="${u.id}">${escapeHtml(u.name)} (${escapeHtml(u.email)})</option>`
                )
                .join('');
            const nonMasterUserOpts = (data.users || [])
                .filter(u => !isMasterUserRecord(u))
                .map(u =>
                    `<option value="${u.id}">${escapeHtml(u.name)} (${escapeHtml(u.email)}) — ${escapeHtml(u.role || 'user')}</option>`
                )
                .join('');
            const mf = getMasterAccountFilterMode();
            const accountRows = (data.users || [])
                .filter(u => !isMasterUserRecord(u))
                .filter(u => userPassesMasterAccountFilter(u, mf))
                .map(u => {
                    const active = u.is_active !== false;
                    const safeName = escapeHtml(u.name || '');
                    const safeEmail = escapeHtml(u.email || '');
                    const safeOrg = escapeHtml(u.tenant_admin_label || '—');
                    return `
                    <tr>
                        <td style="padding:8px 10px;">${safeName}</td>
                        <td style="padding:8px 10px;color:#666;font-size:13px;">${safeEmail}</td>
                        <td style="padding:8px 10px;color:#555;font-size:13px;">${safeOrg}</td>
                        <td style="padding:8px 10px;"><span class="badge ${active ? 'badge-completed' : 'badge-pending'}">${active ? 'Active' : 'Disabled'}</span></td>
                        <td style="padding:8px 10px;">
                            ${active
                    ? `<button type="button" class="btn btn-danger" style="padding:4px 10px;font-size:12px;" onclick="masterSetUserActive(${u.id}, false)">Deactivate</button>`
                    : `<button type="button" class="btn btn-success" style="padding:4px 10px;font-size:12px;" onclick="masterSetUserActive(${u.id}, true)">Activate</button>`}
                        </td>
                    </tr>`;
                }).join('');
            masterSec.innerHTML = `
                <h3 style="margin-top:0;">Self-service registration</h3>
                <p style="color:#666;font-size:13px;">Who may use <strong>Create account</strong> on the login screen. If they are not allowed, they see: <em>Creation not allowed</em>.</p>
                <div style="display:flex;flex-direction:column;gap:8px;margin:12px 0;">
                    <label style="display:flex;align-items:center;gap:8px;cursor:pointer;"><input type="radio" name="masterRegMode" id="masterRegModeOpen" value="open"> Anyone may register</label>
                    <label style="display:flex;align-items:center;gap:8px;cursor:pointer;"><input type="radio" name="masterRegMode" id="masterRegModeEmail" value="email_list"> Only specific email addresses (list below)</label>
                    <label style="display:flex;align-items:center;gap:8px;cursor:pointer;"><input type="radio" name="masterRegMode" id="masterRegModeDomain" value="domain_list"> Only addresses from specific domains (list below)</label>
                </div>
                <div class="form-group">
                    <label>Allowed emails (one per line)</label>
                    <textarea id="masterRegEmails" class="form-control" rows="4" placeholder="user1@company.com&#10;user2@company.com"></textarea>
                </div>
                <div class="form-group">
                    <label>Allowed domains (one per line, e.g. ameyalogistics.com). Subdomains are allowed automatically.</label>
                    <textarea id="masterRegDomains" class="form-control" rows="3" placeholder="company.com"></textarea>
                </div>
                <button type="button" class="btn btn-primary" onclick="saveMasterRegistrationPolicy()">Save registration rules</button>
                <h3 style="margin-top:24px;">Cross-tenant user management</h3>
                <p style="color:#666;font-size:13px;">Add account admins or account users, delete accounts, change admin vs user role, or move a user under another account admin.</p>
                <div class="form-group" style="margin-top:12px;">
                    <label>1 — Add account admin (new tenant)</label>
                    <div style="display:flex;flex-wrap:wrap;gap:8px;align-items:flex-end;">
                        <input type="text" id="masterNewAdminName" class="form-control" placeholder="Name" style="min-width:130px;">
                        <input type="email" id="masterNewAdminEmail" class="form-control" placeholder="Email" style="min-width:160px;">
                        <input type="password" id="masterNewAdminPassword" class="form-control" placeholder="Password" style="min-width:120px;" autocomplete="new-password">
                        <button type="button" class="btn btn-primary" onclick="masterApiCreateOrgAdmin()">Create admin</button>
                    </div>
                </div>
                <div class="form-group">
                    <label>2 — Add account user</label>
                    <div style="display:flex;flex-wrap:wrap;gap:8px;align-items:flex-end;">
                        <input type="text" id="masterNewUserName" class="form-control" placeholder="Name" style="min-width:130px;">
                        <input type="email" id="masterNewUserEmail" class="form-control" placeholder="Email" style="min-width:160px;">
                        <input type="password" id="masterNewUserPassword" class="form-control" placeholder="Password" style="min-width:120px;" autocomplete="new-password">
                        <select id="masterNewUserOrgAdmin" class="form-control" style="min-width:200px;">
                            <option value="">Account admin…</option>${orgAdminOptsForMaster}
                        </select>
                        <button type="button" class="btn btn-primary" onclick="masterApiCreateTeamUser()">Create user</button>
                    </div>
                </div>
                <div class="form-group">
                    <label>User to manage (for actions 3–5)</label>
                    <select id="masterManageUserId" class="form-control" style="max-width:100%;">
                        <option value="">Select user…</option>${nonMasterUserOpts}
                    </select>
                </div>
                <div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:16px;">
                    <button type="button" class="btn btn-danger" onclick="masterApiDeleteUser()">3 — Delete user</button>
                    <select id="masterManageNewRole" class="form-control" style="width:auto;">
                        <option value="admin">4 — Set role: admin</option>
                        <option value="user">4 — Set role: user</option>
                    </select>
                    <button type="button" class="btn btn-secondary" onclick="masterApiPatchUserRole()">Apply role</button>
                    <select id="masterManageMoveOrg" class="form-control" style="min-width:200px;">
                        <option value="">5 — Move to account admin…</option>${orgAdminOptsForMaster}
                    </select>
                    <button type="button" class="btn btn-secondary" onclick="masterApiMoveUserOrg()">Apply org</button>
                </div>
                <hr style="margin:28px 0;border:none;border-top:1px solid #e5e5e5;">
                <h3>Master password reset</h3>
                <p style="color:#666;font-size:13px;">All accounts are listed in <strong>User Management</strong> above. Set a new login password for any user (including other admins) below. Uses the hosted API only.</p>
                <div style="display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end;margin-top:10px;">
                    <div class="form-group" style="margin:0;min-width:220px;">
                        <label>User</label>
                        <select id="masterResetUserId" class="form-control">${opts}</select>
                    </div>
                    <div class="form-group" style="margin:0;min-width:180px;">
                        <label>New password</label>
                        <input type="password" id="masterResetNewPassword" class="form-control" placeholder="Min 6 characters" autocomplete="new-password">
                    </div>
                    <button type="button" class="btn btn-primary" onclick="masterResetUserPassword()">Update password</button>
                </div>
                <h3 style="margin-top:24px;">Activate / deactivate accounts</h3>
                <p style="color:#666;font-size:13px;">Deactivate prevents sign-in for that user (master account cannot be changed). Reactivate to restore access. This table uses the same <strong>Show accounts</strong> filter as User Management above.</p>
                <div style="overflow-x:auto;margin-top:10px;">
                    <table style="width:100%;border-collapse:collapse;font-size:14px;">
                        <thead><tr style="text-align:left;border-bottom:1px solid #ddd;"><th style="padding:8px 10px;">Name</th><th style="padding:8px 10px;">Email</th><th style="padding:8px 10px;">Assigned org (account admin)</th><th style="padding:8px 10px;">Status</th><th style="padding:8px 10px;">Action</th></tr></thead>
                        <tbody>${accountRows || '<tr><td colspan="5" style="padding:12px;color:#999;">No users</td></tr>'}</tbody>
                    </table>
                </div>
            `;
            void masterRegistrationPolicyHydrate();
        } else {
            masterSec.style.display = 'none';
            masterSec.innerHTML = '';
        }
    }

    // Locations
    const locationsHtml = data.locations.map(loc => `
        <div style="display: flex; justify-content: space-between; align-items: center; padding: 10px; background: #f9f9f9; border-radius: 5px; margin-bottom: 10px;">
            <span>${loc.name}</span>
            <button class="btn btn-danger" onclick="removeLocation(${loc.id})" style="padding: 5px 10px; font-size: 12px;">Delete</button>
        </div>
    `).join('');
    document.getElementById('locationsList').innerHTML = locationsHtml || '<p style="color: #999;">No locations</p>';

    // Segregation Types
    const segregationHtml = data.segregationTypes.map(seg => `
        <div style="display: flex; justify-content: space-between; align-items: center; padding: 10px; background: #f9f9f9; border-radius: 5px; margin-bottom: 10px;">
            <span>${seg.name}</span>
            <button class="btn btn-danger" onclick="removeSegregation(${seg.id})" style="padding: 5px 10px; font-size: 12px;">Delete</button>
        </div>
    `).join('');
    document.getElementById('segregationList').innerHTML = segregationHtml || '<p style="color: #999;">No types</p>';

    // Holidays
    const holidaysHtml = data.holidays.map(holiday => `
        <div style="display: flex; justify-content: space-between; align-items: center; padding: 10px; background: #f9f9f9; border-radius: 5px; margin-bottom: 10px;">
            <span>${formatDateDisplay(holiday.date)} - ${holiday.description || 'Holiday'}</span>
            <button class="btn btn-danger" onclick="removeHoliday(${holiday.id})" style="padding: 5px 10px; font-size: 12px;">Delete</button>
        </div>
    `).join('');
    document.getElementById('holidaysList').innerHTML = holidaysHtml || '<p style="color: #999;">No holidays</p>';

    // Set auto-export checkbox state
    const autoExportCheckbox = document.getElementById('autoExportEnabled');
    if (autoExportCheckbox) {
        autoExportCheckbox.checked = localStorage.getItem('autoExportEnabled') === 'true';
    }

    const dataMgmt = document.getElementById('settingsDataManagement');
    const exportBtn = document.getElementById('settingsDataExportBtn');
    const importBtn = document.getElementById('settingsDataImportBtn');
    const autoWrap = document.getElementById('settingsAutoExportWrap');
    const note = document.getElementById('settingsDataMgmtNote');
    if (dataMgmt) {
        if (isApiMode() && currentUser.isMaster) {
            dataMgmt.style.display = 'none';
        } else {
            dataMgmt.style.display = 'block';
            if (exportBtn) {
                exportBtn.textContent = isApiMode() && !currentUser.isMaster
                    ? 'Download backup (cloud database)'
                    : 'Export all data (JSON)';
            }
            if (importBtn) {
                importBtn.textContent = isApiMode() && !currentUser.isMaster
                    ? 'Restore from backup (cloud)'
                    : 'Import data (JSON)';
            }
            if (autoWrap) {
                autoWrap.style.display = isApiMode() ? 'none' : 'block';
            }
            if (note) {
                note.textContent = isApiMode() && !currentUser.isMaster
                    ? 'Backup and restore use your workspace on the server (not browser storage).'
                    : '';
                note.style.display = note.textContent ? 'block' : 'none';
            }
        }
    }
}

async function masterRegistrationPolicyHydrate() {
    if (!isApiMode() || !currentUser || !currentUser.isMaster) return;
    const rOpen = document.getElementById('masterRegModeOpen');
    if (!rOpen) return;
    try {
        const res = await apiFetch('/api/master/registration-policy');
        if (!res.ok) return;
        const p = await res.json();
        const mode = p.registrationMode || 'open';
        const rEmail = document.getElementById('masterRegModeEmail');
        const rDom = document.getElementById('masterRegModeDomain');
        rOpen.checked = mode === 'open';
        if (rEmail) rEmail.checked = mode === 'email_list';
        if (rDom) rDom.checked = mode === 'domain_list';
        const te = document.getElementById('masterRegEmails');
        const td = document.getElementById('masterRegDomains');
        if (te && Array.isArray(p.allowedEmails)) te.value = p.allowedEmails.join('\n');
        if (td && Array.isArray(p.allowedDomains)) td.value = p.allowedDomains.join('\n');
    } catch (e) {
        console.error(e);
    }
}

async function saveMasterRegistrationPolicy() {
    if (!isApiMode() || !currentUser || !currentUser.isMaster) return;
    const sel = document.querySelector('input[name="masterRegMode"]:checked');
    const mode = (sel && sel.value) || 'open';
    const emailsRaw = (document.getElementById('masterRegEmails') && document.getElementById('masterRegEmails').value) || '';
    const domainsRaw = (document.getElementById('masterRegDomains') && document.getElementById('masterRegDomains').value) || '';
    const allowedEmails = emailsRaw.split(/[\n,;]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
    const allowedDomains = domainsRaw.split(/[\n,;]+/).map(s => s.trim().toLowerCase().replace(/^@+/, '')).filter(Boolean);
    try {
        const res = await apiFetch('/api/master/registration-policy', {
            method: 'PUT',
            body: JSON.stringify({ registrationMode: mode, allowedEmails, allowedDomains }),
        });
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            alert(err.error || 'Could not save registration rules.');
            return;
        }
        alert('Registration rules saved.');
    } catch (e) {
        console.error(e);
        alert('Request failed. Check your connection and API URL.');
    }
}

async function masterRefreshAfterUserChange() {
    try {
        await apiPullWorkspace();
    } catch (e) {
        console.error(e);
    }
    renderSettings();
    renderUsers();
    renderTasks();
}

async function masterApiCreateOrgAdmin() {
    if (!isApiMode() || !currentUser || !currentUser.isMaster) return;
    const name = document.getElementById('masterNewAdminName') && document.getElementById('masterNewAdminName').value.trim();
    const email = document.getElementById('masterNewAdminEmail') && document.getElementById('masterNewAdminEmail').value.trim();
    const password = document.getElementById('masterNewAdminPassword') && document.getElementById('masterNewAdminPassword').value;
    if (!name || !email || !password || password.length < 6) {
        alert('Name, email, and password (min 6 characters) required.');
        return;
    }
    try {
        const res = await apiFetch('/api/master/users', {
            method: 'POST',
            body: JSON.stringify({ name, email, password, accountType: 'org_admin' }),
        });
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            alert(err.error || 'Create failed');
            return;
        }
        document.getElementById('masterNewAdminName').value = '';
        document.getElementById('masterNewAdminEmail').value = '';
        document.getElementById('masterNewAdminPassword').value = '';
        alert('Account admin created.');
        await masterRefreshAfterUserChange();
    } catch (e) {
        console.error(e);
        alert('Request failed.');
    }
}

async function masterApiCreateTeamUser() {
    if (!isApiMode() || !currentUser || !currentUser.isMaster) return;
    const name = document.getElementById('masterNewUserName') && document.getElementById('masterNewUserName').value.trim();
    const email = document.getElementById('masterNewUserEmail') && document.getElementById('masterNewUserEmail').value.trim();
    const password = document.getElementById('masterNewUserPassword') && document.getElementById('masterNewUserPassword').value;
    const orgSel = document.getElementById('masterNewUserOrgAdmin');
    const orgAdminUserId = orgSel ? parseInt(orgSel.value, 10) : NaN;
    if (!name || !email || !password || password.length < 6) {
        alert('Name, email, and password (min 6 characters) required.');
        return;
    }
    if (Number.isNaN(orgAdminUserId) || orgAdminUserId <= 0) {
        alert('Select an account admin.');
        return;
    }
    try {
        const res = await apiFetch('/api/master/users', {
            method: 'POST',
            body: JSON.stringify({
                name,
                email,
                password,
                accountType: 'team_user',
                orgAdminUserId,
            }),
        });
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            alert(err.error || 'Create failed');
            return;
        }
        document.getElementById('masterNewUserName').value = '';
        document.getElementById('masterNewUserEmail').value = '';
        document.getElementById('masterNewUserPassword').value = '';
        alert('Account user created.');
        await masterRefreshAfterUserChange();
    } catch (e) {
        console.error(e);
        alert('Request failed.');
    }
}

async function masterApiDeleteUser() {
    if (!isApiMode() || !currentUser || !currentUser.isMaster) return;
    const sel = document.getElementById('masterManageUserId');
    const uid = sel ? parseInt(sel.value, 10) : NaN;
    if (Number.isNaN(uid) || uid <= 0) {
        alert('Select a user.');
        return;
    }
    if (!confirm('Permanently delete this user from the server? This cannot be undone.')) return;
    try {
        const res = await apiFetch(`/api/master/users/${uid}`, { method: 'DELETE' });
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            alert(err.error || 'Delete failed');
            return;
        }
        alert('User deleted.');
        await masterRefreshAfterUserChange();
    } catch (e) {
        console.error(e);
        alert('Request failed.');
    }
}

async function masterApiPatchUserRole() {
    if (!isApiMode() || !currentUser || !currentUser.isMaster) return;
    const sel = document.getElementById('masterManageUserId');
    const uid = sel ? parseInt(sel.value, 10) : NaN;
    const roleEl = document.getElementById('masterManageNewRole');
    const role = roleEl && roleEl.value;
    if (Number.isNaN(uid) || uid <= 0) {
        alert('Select a user.');
        return;
    }
    if (role !== 'admin' && role !== 'user') {
        alert('Invalid role.');
        return;
    }
    try {
        const res = await apiFetch(`/api/master/users/${uid}`, {
            method: 'PATCH',
            body: JSON.stringify({ role }),
        });
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            alert(err.error || 'Update failed');
            return;
        }
        alert('Role updated.');
        await masterRefreshAfterUserChange();
    } catch (e) {
        console.error(e);
        alert('Request failed.');
    }
}

async function masterApiMoveUserOrg() {
    if (!isApiMode() || !currentUser || !currentUser.isMaster) return;
    const sel = document.getElementById('masterManageUserId');
    const uid = sel ? parseInt(sel.value, 10) : NaN;
    const orgSel = document.getElementById('masterManageMoveOrg');
    const orgAdminUserId = orgSel ? parseInt(orgSel.value, 10) : NaN;
    if (Number.isNaN(uid) || uid <= 0) {
        alert('Select a user.');
        return;
    }
    if (Number.isNaN(orgAdminUserId) || orgAdminUserId <= 0) {
        alert('Select an account admin to move the user under.');
        return;
    }
    try {
        const res = await apiFetch(`/api/master/users/${uid}`, {
            method: 'PATCH',
            body: JSON.stringify({ orgAdminUserId }),
        });
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            alert(err.error || 'Update failed');
            return;
        }
        alert('User moved under that account admin.');
        await masterRefreshAfterUserChange();
    } catch (e) {
        console.error(e);
        alert('Request failed.');
    }
}

async function masterSetUserActive(userId, active) {
    if (!isApiMode() || !currentUser || !currentUser.isMaster) return;
    const verb = active ? 'activate' : 'deactivate';
    if (!confirm(`${verb.charAt(0).toUpperCase() + verb.slice(1)} this account?`)) return;
    try {
        const res = await apiFetch(`/api/master/users/${userId}/active`, {
            method: 'POST',
            body: JSON.stringify({ active })
        });
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            alert(err.error || 'Update failed');
            return;
        }
        try {
            await apiPullWorkspace();
        } catch (pullErr) {
            console.error('Workspace refresh after account status:', pullErr);
        }
        renderSettings();
        renderUsers();
        renderTasks();
        alert(active ? 'Account activated.' : 'Account deactivated.');
    } catch (e) {
        console.error(e);
        alert('Request failed. Check your connection and API URL.');
    }
}

async function masterResetUserPassword() {
    if (!isApiMode() || !currentUser || !currentUser.isMaster) return;
    const uid = document.getElementById('masterResetUserId') && document.getElementById('masterResetUserId').value;
    const np = document.getElementById('masterResetNewPassword') && document.getElementById('masterResetNewPassword').value;
    if (!uid) {
        alert('Select a user.');
        return;
    }
    if (!np || np.length < 6) {
        alert('Password must be at least 6 characters.');
        return;
    }
    try {
        const res = await apiFetch(`/api/master/users/${uid}/password`, {
            method: 'POST',
            body: JSON.stringify({ newPassword: np })
        });
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            alert(err.error || 'Reset failed');
            return;
        }
        alert('Password updated for that user.');
        document.getElementById('masterResetNewPassword').value = '';
    } catch (e) {
        console.error(e);
        alert('Request failed. Check your connection and API URL.');
    }
}

function addLocation() {
    const name = document.getElementById('newLocation').value.trim();
    if (!name) return;

    updateData(data => {
        const maxId = data.locations.length > 0 ? Math.max(...data.locations.map(l => l.id)) : 0;
        data.locations.push({ id: maxId + 1, name });
    });

    document.getElementById('newLocation').value = '';
    renderSettings();
}

function removeLocation(id) {
    if (!confirm('Are you sure? This will remove the location from all tasks.')) return;

    updateData(data => {
        data.locations = data.locations.filter(l => l.id !== id);
    });

    renderSettings();
}

function addSegregation() {
    const name = document.getElementById('newSegregation').value.trim();
    if (!name) return;

    updateData(data => {
        const maxId = data.segregationTypes.length > 0 ? Math.max(...data.segregationTypes.map(s => s.id)) : 0;
        data.segregationTypes.push({ id: maxId + 1, name });
    });

    document.getElementById('newSegregation').value = '';
    renderSettings();
}

function removeSegregation(id) {
    if (!confirm('Are you sure? This will remove the type from all tasks.')) return;

    updateData(data => {
        data.segregationTypes = data.segregationTypes.filter(s => s.id !== id);
    });

    renderSettings();
}

function addHoliday() {
    const date = document.getElementById('newHolidayDate').value;
    const description = document.getElementById('newHolidayDesc').value.trim();

    if (!date) return;

    updateData(data => {
        const maxId = data.holidays.length > 0 ? Math.max(...data.holidays.map(h => h.id)) : 0;
        data.holidays.push({ id: maxId + 1, date, description: description || 'Holiday' });
    });

    document.getElementById('newHolidayDate').value = '';
    document.getElementById('newHolidayDesc').value = '';
    renderSettings();
}

function removeHoliday(id) {
    if (!confirm('Are you sure?')) return;

    updateData(data => {
        data.holidays = data.holidays.filter(h => h.id !== id);
    });

    renderSettings();
}

// Export
function updateMonthFilter() {
    const fromEl = document.getElementById('filterDashboardMonthFrom');
    const toEl = document.getElementById('filterDashboardMonthTo');
    const today = new Date();
    const currentMonthStr = today.getFullYear() + '-' + String(today.getMonth() + 1).padStart(2, '0');

    if (fromEl && !fromEl.value) fromEl.value = currentMonthStr;
    if (toEl && !toEl.value) toEl.value = currentMonthStr;
    if (fromEl && toEl && fromEl.value > toEl.value) toEl.value = fromEl.value;

    renderInteractiveDashboard();
}

// Interactive Dashboard
function renderInteractiveDashboard() {
    const data = getData();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayStr = formatDateString(today);

    // Set default from/to month from saved period
    const period = getInteractiveDashboardPeriod();
    const fromEl = document.getElementById('filterDashboardMonthFrom');
    const toEl = document.getElementById('filterDashboardMonthTo');
    const currentMonthStr = today.getFullYear() + '-' + String(today.getMonth() + 1).padStart(2, '0');
    if (fromEl && !fromEl.value) fromEl.value = period.from || currentMonthStr;
    if (toEl && !toEl.value) toEl.value = period.to || currentMonthStr;
    if (fromEl && toEl && fromEl.value && toEl.value && fromEl.value > toEl.value) toEl.value = fromEl.value;

    // Populate filter dropdowns
    const userSelect = document.getElementById('filterDashboardUser');
    const locationSelect = document.getElementById('filterDashboardLocation');
    const statusSelect = document.getElementById('filterDashboardStatus');

    // Store current selections before repopulating
    const currentUserValue = userSelect ? userSelect.value : '';
    const currentStatusValue = statusSelect ? statusSelect.value : '';

    // Check if this is the first time loading (no options exist yet or only "All Users" option)
    const isFirstLoad = !userSelect || userSelect.options.length <= 1;

    // Populate user dropdown
    userSelect.innerHTML = '<option value="">All Users</option>' +
        usersVisibleInPickers().map(u =>
            `<option value="${u.id}">${escapeHtml(u.name)}</option>`
        ).join('');

    // Set default user to logged-in user only on first load, otherwise preserve user selection
    if (isFirstLoad) {
        // First load - set default to logged-in user
        userSelect.value = currentUser.id;
    } else {
        // Not first load - preserve whatever the user selected (including "All Users" = '')
        userSelect.value = currentUserValue;
    }

    // Set default status to "Pending" only on first load, otherwise preserve user selection
    if (statusSelect) {
        if (isFirstLoad) {
            // First load - set default to "Pending"
            statusSelect.value = 'pending';
        } else {
            // Not first load - preserve whatever the user selected (including "All Status" = '')
            statusSelect.value = currentStatusValue;
        }
    }

    locationSelect.innerHTML = '<option value="">All Locations</option>' +
        data.locations.map(l =>
            `<option value="${l.id}">${l.name}</option>`
        ).join('');

    // Get filter values
    const monthFromValue = document.getElementById('filterDashboardMonthFrom')?.value;
    const monthToValue = document.getElementById('filterDashboardMonthTo')?.value;
    const userFilter = document.getElementById('filterDashboardUser').value;
    const taskTypeFilter = document.getElementById('filterDashboardTaskType').value;
    const frequencyFilter = document.getElementById('filterDashboardFrequency').value;
    const statusFilter = document.getElementById('filterDashboardStatus').value;
    const locationFilter = document.getElementById('filterDashboardLocation').value;
    const priorityFilter = document.getElementById('filterDashboardPriority').value;

    // Parse from/to month range
    let fromDateRange = null;
    let toDateRange = null;
    if (monthFromValue) {
        const [y1, m1] = monthFromValue.split('-').map(Number);
        fromDateRange = new Date(y1, m1 - 1, 1);
        fromDateRange.setHours(0, 0, 0, 0);
    }
    if (monthToValue) {
        const [y2, m2] = monthToValue.split('-').map(Number);
        toDateRange = new Date(y2, m2, 0); // Last day of month
        toDateRange.setHours(23, 59, 59, 999);
    }

    // Filter tasks
    let filteredTasks = data.tasks.filter(task => {
        // Exclude removed tasks
        if (task.removed_at) return false;

        // User filter
        if (userFilter && task.assigned_to !== parseInt(userFilter)) return false;

        // Task Type filter
        if (taskTypeFilter && task.task_type !== taskTypeFilter) return false;

        // Frequency filter
        if (frequencyFilter) {
            if (frequencyFilter === 'one_time' && task.task_type !== 'one_time') return false;
            if (frequencyFilter === 'work_plan' && task.task_type !== 'work_plan') return false;
            if (frequencyFilter === 'audit_point' && task.task_type !== 'audit_point') return false;
            if (frequencyFilter === 'without_due_date' && task.task_type !== 'without_due_date') return false;
            if (frequencyFilter !== 'one_time' && frequencyFilter !== 'work_plan' && frequencyFilter !== 'audit_point' && frequencyFilter !== 'without_due_date' &&
                (task.task_type !== 'recurring' || task.frequency !== frequencyFilter)) return false;
        }

        // Location filter
        if (locationFilter && task.location_id !== parseInt(locationFilter)) return false;

        // Priority filter
        if (priorityFilter && task.priority !== priorityFilter) return false;

        // From/To month filter: Done (completed or need improvement) by completion/due date in range; others by due date in range
        if (fromDateRange || toDateRange) {
            const dueDate = task.due_date || task.next_due_date;
            if (isTaskCompleted(task)) {
                const refDateStr = task.completed_at ? task.completed_at.split('T')[0] : dueDate;
                if (!refDateStr) return false;
                const dateParts = refDateStr.split('-');
                const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
                taskDate.setHours(0, 0, 0, 0);
                if (fromDateRange && taskDate < fromDateRange) return false;
                if (toDateRange && taskDate > toDateRange) return false;
            } else {
                if (dueDate) {
                    const dateParts = dueDate.split('-');
                    const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
                    taskDate.setHours(0, 0, 0, 0);
                    if (fromDateRange && taskDate < fromDateRange) return false;
                    if (toDateRange && taskDate > toDateRange) return false;
                } else {
                    return false; // No due date, exclude when month range is set
                }
            }
        }

        // Status filter (Completed = completed or need improvement; both treated as Done)
        if (statusFilter) {
            const dueDate = task.due_date || task.next_due_date;
            if (statusFilter === 'completed' && !isTaskCompleted(task)) return false;
            if (statusFilter === 'pending' && (isTaskCompleted(task) || task.task_action === 'in_process' || task.task_action === 'not_done')) return false;
            if (statusFilter === 'in_process' && task.task_action !== 'in_process') return false;
            if (statusFilter === 'overdue') {
                if (!dueDate || isTaskCompleted(task) || task.task_action === 'in_process' || task.task_action === 'not_done') return false;
                const dateParts = dueDate.split('-');
                const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
                taskDate.setHours(0, 0, 0, 0);
                if (taskDate >= today) return false;
            }
            if (statusFilter === 'not_due') {
                if (!dueDate || isTaskCompleted(task) || task.task_action === 'in_process' || task.task_action === 'not_done') return false;
                const dateParts = dueDate.split('-');
                const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
                taskDate.setHours(0, 0, 0, 0);
                if (taskDate < today) return false;
            }
        }

        // Visibility check
        const isAssigned = task.assigned_to === currentUser.id ||
            (task.is_team_task && currentUser.role === 'admin') ||
            currentUser.role === 'admin';
        return isAssigned;
    });

    // Sort tasks
    filteredTasks.sort((a, b) => {
        // Completed tasks last (including needs improvement)
        if (isTaskCompleted(a) && !isTaskCompleted(b)) return 1;
        if (!isTaskCompleted(a) && isTaskCompleted(b)) return -1;

        // Then by priority
        const priorityOrder = { high: 1, medium: 2, low: 3 };
        const aPriority = priorityOrder[a.priority] || 4;
        const bPriority = priorityOrder[b.priority] || 4;
        if (aPriority !== bPriority) return aPriority - bPriority;

        // Then by due date
        const aDate = a.due_date || a.next_due_date || '9999-12-31';
        const bDate = b.due_date || b.next_due_date || '9999-12-31';
        return aDate.localeCompare(bDate);
    });

    // Render tasks in table format
    const tasksHtml = filteredTasks.length > 0
        ? `
            <div class="table-container" style="background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                <table class="table" style="margin: 0;">
                    <thead>
                        <tr>
                            <th style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; font-weight: 600;">Task Name</th>
                            <th style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; font-weight: 600;">Status</th>
                            <th style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; font-weight: 600;">Assigned To</th>
                            <th style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; font-weight: 600;">Location</th>
                            <th style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; font-weight: 600;">Due Date</th>
                            <th style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; font-weight: 600;">Expected Completion</th>
                            <th style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; font-weight: 600;">Action</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${filteredTasks.map(task => renderInteractiveTaskCard(task)).join('')}
                    </tbody>
                </table>
            </div>
        `
        : '<p style="text-align: center; color: #999; padding: 20px;">No tasks found matching the filters</p>';

    // Store filtered tasks for export
    window.currentFilteredTasks = filteredTasks;

    document.getElementById('interactiveDashboardContent').innerHTML = `
        <div style="margin-bottom: 20px;">
            <h3>Filtered Results: ${filteredTasks.length} task(s)</h3>
        </div>
        ${tasksHtml}
    `;
}

function exportInteractiveDashboardCSV() {
    const data = getData();
    const filteredTasks = window.currentFilteredTasks || [];

    if (filteredTasks.length === 0) {
        alert('No tasks to export. Please apply filters first.');
        return;
    }

    const headers = [
        'Task Number',
        'Task Name',
        'Description',
        'Assigned to (Name)',
        'Location',
        'Task Type',
        'Frequency',
        'Due date calculation type',
        'Due day of month',
        'Recurrence Type',
        'Start date (for due date calculation)',
        'Estimated Minutes',
        'Team task (true/false)',
        'Status',
        'Due Date',
        'Expected Date of Completion',
        'Completion Date',
        'Completion Remark'
    ];
    const rows = filteredTasks.map(task => {
        const user = data.users.find(u => u.id === task.assigned_to);
        const location = data.locations.find(l => l.id === task.location_id);

        // Format start date as DD-MM-YYYY
        let startDateStr = '';
        if (task.start_date) {
            startDateStr = formatDateDisplay(task.start_date);
        }

        // Format due date as DD-MM-YYYY
        let dueDateStr = '';
        const dueDate = task.due_date || task.next_due_date;
        if (dueDate) {
            dueDateStr = formatDateDisplay(dueDate);
        } else if (task.task_type === 'without_due_date') {
            dueDateStr = 'No Due Date';
        }

        // Format expected completion date as DD-MM-YYYY (when set)
        let expectedCompletionStr = '';
        if (task.expected_completion_date) {
            expectedCompletionStr = formatDateDisplay(task.expected_completion_date);
        }

        // Format completion date as DD-MM-YYYY
        let completionDateStr = '';
        if (task.completion_date) {
            completionDateStr = formatDateDisplay(task.completion_date);
        } else if (task.completed_at) {
            completionDateStr = formatDateDisplay(task.completed_at);
        }

        // Format status
        let statusStr = '';
        if (task.task_action === 'completed') {
            statusStr = 'Completed';
        } else if (task.task_action === 'completed_need_improvement') {
            statusStr = 'Needs Improvement';
        } else if (task.task_action === 'in_process') {
            statusStr = 'In Process';
        } else {
            statusStr = 'Pending';
        }

        return [
            task.task_number != null ? String(task.task_number) : '',
            task.task_name || '',
            task.description || '',
            user ? user.name : '',
            location ? location.name : '',
            task.task_type || '',
            task.frequency || '',
            task.due_date_type || '',
            task.due_day || '',
            task.recurrence_type || '',
            startDateStr,
            task.est_minutes || '',
            task.is_team_task ? 'true' : 'false',
            statusStr,
            dueDateStr,
            expectedCompletionStr,
            completionDateStr,
            task.comment || ''
        ];
    });

    // Escape CSV values
    const escapeCSV = (value) => {
        if (value === null || value === undefined) return '';
        const str = String(value);
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
            return `"${str.replace(/"/g, '""')}"`;
        }
        return str;
    };

    const csv = [
        headers.map(escapeCSV).join(','),
        ...rows.map(row => row.map(escapeCSV).join(','))
    ].join('\n');

    const fromVal = document.getElementById('filterDashboardMonthFrom')?.value;
    const toVal = document.getElementById('filterDashboardMonthTo')?.value;
    const filename = (fromVal && toVal)
        ? `tasks-${fromVal}-to-${toVal}-${formatDateString(new Date())}.csv`
        : `tasks-${formatDateString(new Date())}.csv`;

    downloadFile(csv, filename, 'text/csv');
}

function exportInteractiveDashboardCSVByDateRange() {
    const fromDate = prompt('Enter From Date (DD-MM-YYYY):');
    if (!fromDate) return;

    const toDate = prompt('Enter To Date (DD-MM-YYYY):');
    if (!toDate) return;

    // Parse dates
    const fromMatch = fromDate.match(/(\d{2})-(\d{2})-(\d{4})/);
    const toMatch = toDate.match(/(\d{2})-(\d{2})-(\d{4})/);

    if (!fromMatch || !toMatch) {
        alert('Invalid date format. Please use DD-MM-YYYY format.');
        return;
    }

    const fromDateObj = new Date(parseInt(fromMatch[3]), parseInt(fromMatch[2]) - 1, parseInt(fromMatch[1]));
    const toDateObj = new Date(parseInt(toMatch[3]), parseInt(toMatch[2]) - 1, parseInt(toMatch[1]));
    toDateObj.setHours(23, 59, 59, 999); // Include the entire end date

    if (fromDateObj > toDateObj) {
        alert('From date must be before or equal to To date.');
        return;
    }

    const data = getData();

    // Filter tasks by date range
    const filteredTasks = data.tasks.filter(task => {
        const dueDate = task.due_date || task.next_due_date;
        if (!dueDate) return false; // Exclude tasks without due date

        const dateParts = dueDate.split('-');
        const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
        taskDate.setHours(0, 0, 0, 0);

        return taskDate >= fromDateObj && taskDate <= toDateObj &&
            (task.assigned_to === currentUser.id ||
                (task.is_team_task && currentUser.role === 'admin') ||
                currentUser.role === 'admin');
    });

    if (filteredTasks.length === 0) {
        alert('No tasks found in the specified date range.');
        return;
    }

    const headers = [
        'Task Number',
        'Task Name',
        'Description',
        'Assigned to (Name)',
        'Location',
        'Task Type',
        'Frequency',
        'Due date calculation type',
        'Due day of month',
        'Recurrence Type',
        'Start date (for due date calculation)',
        'Estimated Minutes',
        'Team task (true/false)',
        'Status',
        'Due Date',
        'Expected Date of Completion',
        'Completion Date',
        'Completion Remark'
    ];
    const rows = filteredTasks.map(task => {
        const user = data.users.find(u => u.id === task.assigned_to);
        const location = data.locations.find(l => l.id === task.location_id);

        let startDateStr = '';
        if (task.start_date) startDateStr = formatDateDisplay(task.start_date);

        let dueDateStr = '';
        const dueDate = task.due_date || task.next_due_date;
        if (dueDate) dueDateStr = formatDateDisplay(dueDate);
        else if (task.task_type === 'without_due_date') dueDateStr = 'No Due Date';

        let expectedCompletionStr = '';
        if (task.expected_completion_date) expectedCompletionStr = formatDateDisplay(task.expected_completion_date);

        let completionDateStr = '';
        if (task.completion_date) completionDateStr = formatDateDisplay(task.completion_date);
        else if (task.completed_at) completionDateStr = formatDateDisplay(task.completed_at);

        const statusStr = task.task_action === 'completed' ? 'Completed' :
            task.task_action === 'completed_need_improvement' ? 'Needs Improvement' :
            task.task_action === 'in_process' ? 'In Process' :
            task.task_action === 'not_done' ? 'Not Done' : 'Pending';

        return [
            task.task_number != null ? String(task.task_number) : '',
            task.task_name || '',
            task.description || '',
            user ? user.name : '',
            location ? location.name : '',
            task.task_type || '',
            task.frequency || '',
            task.due_date_type || '',
            task.due_day || '',
            task.recurrence_type || '',
            startDateStr,
            task.est_minutes || '',
            task.is_team_task ? 'true' : 'false',
            statusStr,
            dueDateStr,
            expectedCompletionStr,
            completionDateStr,
            task.comment || ''
        ];
    });

    const escapeCSV = (value) => {
        if (value === null || value === undefined) return '';
        const str = String(value);
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
            return `"${str.replace(/"/g, '""')}"`;
        }
        return str;
    };

    const csv = [
        headers.map(escapeCSV).join(','),
        ...rows.map(r => r.map(escapeCSV).join(','))
    ].join('\n');

    const filename = `tasks-${fromDate.replace(/\//g, '-')}-to-${toDate.replace(/\//g, '-')}.csv`;
    downloadFile(csv, filename, 'text/csv');
}

function renderInteractiveTaskCard(task) {
    const data = getData();
    const assignedUser = data.users.find(u => u.id === task.assigned_to);
    const location = data.locations.find(l => l.id === task.location_id);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayStr = formatDateString(today);
    const dueDate = task.due_date || task.next_due_date;

    // Determine status
    let statusClass = '';
    let statusText = '';
    if (task.task_action === 'completed') {
        statusClass = 'badge-completed';
        statusText = 'Completed';
    } else if (task.task_action === 'completed_need_improvement') {
        statusClass = 'badge-warning';
        statusText = 'Need Improvement';
    } else if (task.task_action === 'in_process') {
        statusClass = 'badge-info';
        statusText = 'In Process';
    } else if (dueDate) {
        const dateParts = dueDate.split('-');
        const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
        taskDate.setHours(0, 0, 0, 0);
        if (taskDate < today) {
            statusClass = 'badge-pending';
            statusText = 'Overdue';
        } else {
            statusClass = 'badge-low';
            statusText = 'Pending';
        }
    } else {
        statusClass = 'badge-low';
        statusText = 'No Due Date';
    }

    const priorityBadge = task.priority
        ? `<span class="badge badge-${task.priority}">${task.priority.toUpperCase()}</span>`
        : '';

    const frequencyBadge = task.frequency
        ? `<span class="badge badge-recurring">${task.frequency.charAt(0).toUpperCase() + task.frequency.slice(1)}</span>`
        : '';

    const dueDateStr = dueDate ? formatDateDisplay(dueDate) : 'Not set';
    const expectedStr = task.expected_completion_date ? formatDateDisplay(task.expected_completion_date) : '—';

    return `
        <tr style="border-bottom: 1px solid #e0e0e0; transition: background 0.2s; cursor: pointer;" 
            onmouseenter="this.style.background='#f9f9f9'" 
            onmouseleave="this.style.background=''"
            onclick="openInteractiveTaskPopup(${task.id})">
            <td style="padding: 15px;">
                <div style="font-weight: 600; color: #333; margin-bottom: 5px;">${task.task_name}</div>
                <div style="font-size: 12px; color: #666;">${task.description || 'No description'}</div>
            </td>
            <td style="padding: 15px;">
                <div style="display: flex; gap: 5px; flex-wrap: wrap;">
                    ${priorityBadge}
                    ${frequencyBadge}
                    <span class="badge ${statusClass}">${statusText}</span>
                </div>
            </td>
            <td style="padding: 15px; color: #666; font-size: 14px;">
                ${assignedUser ? assignedUser.name : 'Unknown'}
            </td>
            <td style="padding: 15px; color: #666; font-size: 14px;">
                ${location ? location.name : 'Unknown'}
            </td>
            <td style="padding: 15px; color: #666; font-size: 14px;">
                ${dueDateStr}
            </td>
            <td style="padding: 15px; color: #666; font-size: 14px;">
                ${expectedStr}
            </td>
            <td style="padding: 15px;">
                ${!isTaskCompleted(task) ? `
                    <div style="display: flex; gap: 6px; flex-wrap: wrap;">
                        <button class="btn btn-success" onclick="completeTaskWithRemark(${task.id}); renderInteractiveDashboard();" 
                                style="padding: 6px 12px; font-size: 12px;">Complete</button>
                        ${task.task_action !== 'in_process' ? `
                            <button class="btn btn-info" onclick="markTaskInProcess(${task.id}); renderDashboard(); renderInteractiveDashboard();" 
                                    style="padding: 6px 12px; font-size: 12px; background: #17a2b8; color: white; border: none;">Mark In Process</button>
                        ` : `
                            <button class="btn btn-info" onclick="openEditExpectedDateModal(${task.id}); renderDashboard(); renderInteractiveDashboard();" 
                                    style="padding: 6px 12px; font-size: 12px; background: #17a2b8; color: white; border: none;">Edit Expected Date</button>
                        `}
                    </div>
                ` : `
                    <span class="badge badge-completed">Done</span>
                `}
            </td>
        </tr>
    `;
}

function exportData(format) {
    const data = getData();

    if (format === 'csv') {
        // Use filtered tasks if available, otherwise use all tasks
        const tasksToExport = window.currentFilteredTasksForExport || data.tasks;

        if (tasksToExport.length === 0) {
            alert('No tasks to export. Please apply filters to see tasks first.');
            return;
        }

        const headers = [
            'Task Number',
            'Task Name',
            'Description',
            'Assigned to (Name)',
            'Location',
            'Task Type',
            'Frequency',
            'Due date calculation type',
            'Due day of month',
            'Recurrence Type',
            'Start date (for due date calculation)',
            'Estimated Minutes',
            'Team task (true/false)',
            'Due Date',
            'Expected Date of Completion',
            'Completion Date',
            'Status',
            'Completion Remark'
        ];
        const rows = tasksToExport.map(task => {
            const user = data.users.find(u => u.id === task.assigned_to);
            const location = data.locations.find(l => l.id === task.location_id);

            // Format start date as DD-MM-YYYY
            let startDateStr = '';
            if (task.start_date) {
                startDateStr = formatDateDisplay(task.start_date);
            }

            // Format due date as DD-MM-YYYY
            let dueDateStr = '';
            const dueDate = task.due_date || task.next_due_date;
            if (dueDate) {
                dueDateStr = formatDateDisplay(dueDate);
            } else if (task.task_type === 'without_due_date') {
                dueDateStr = 'No Due Date';
            }

            // Format expected completion date as DD-MM-YYYY (when set)
            let expectedCompletionStr = '';
            if (task.expected_completion_date) {
                expectedCompletionStr = formatDateDisplay(task.expected_completion_date);
            }

            // Format completion date as DD-MM-YYYY
            let completionDateStr = '';
            if (task.completion_date) {
                completionDateStr = formatDateDisplay(task.completion_date);
            } else if (task.completed_at) {
                completionDateStr = formatDateDisplay(task.completed_at);
            }

            const statusStr = task.task_action === 'completed' ? 'Completed' :
                task.task_action === 'completed_need_improvement' ? 'Needs Improvement' :
                task.task_action === 'in_process' ? 'In Process' :
                task.task_action === 'not_done' ? 'Not Done' : 'Pending';
            return [
                task.task_number != null ? String(task.task_number) : '',
                task.task_name || '',
                task.description || '',
                user ? user.name : '',
                location ? location.name : '',
                task.task_type || '',
                task.frequency || '',
                task.due_date_type || '',
                task.due_day || '',
                task.recurrence_type || '',
                startDateStr,
                task.est_minutes || '',
                task.is_team_task ? 'true' : 'false',
                dueDateStr,
                expectedCompletionStr,
                completionDateStr,
                statusStr,
                task.comment || ''
            ];
        });

        // Escape CSV values
        const escapeCSV = (value) => {
            if (value === null || value === undefined) return '';
            const str = String(value);
            if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                return `"${str.replace(/"/g, '""')}"`;
            }
            return str;
        };

        const csv = [
            headers.map(escapeCSV).join(','),
            ...rows.map(r => r.map(escapeCSV).join(','))
        ].join('\n');

        const filename = `tasks-filtered-${formatDateString(new Date())}.csv`;
        downloadFile(csv, filename, 'text/csv');
    } else if (format === 'json') {
        // JSON export always exports ALL data (complete backup), not filtered data
        // This ensures full backup/restore capability across browsers
        exportAllData();
    }
}

// Export/Import Functions (export includes IndexedDB attachment blobs for full restore in local mode)
async function exportAllData() {
    if (isApiMode() && currentUser && currentUser.role === 'admin' && !currentUser.isMaster) {
        try {
            const res = await apiFetch('/api/workspace/backup');
            if (!res.ok) {
                const err = await res.json().catch(() => ({}));
                alert(err.error || 'Could not download backup from server.');
                return;
            }
            const payload = await res.json();
            const filename = `todo-app-cloud-backup-${formatDateString(new Date())}.json`;
            downloadFile(JSON.stringify(payload, null, 2), filename, 'application/json');
            alert('✅ Cloud workspace backup downloaded.\n\nFile: ' + filename + '\n\nThis snapshot is stored on the server for your account.');
        } catch (e) {
            reportError(e, 'Cloud backup');
        }
        return;
    }

    const data = getData();

    const completeData = {
        users: data.users || [],
        tasks: data.tasks || [],
        locations: data.locations || [],
        locationItems: data.locationItems || [],
        segregationTypes: data.segregationTypes || [],
        holidays: data.holidays || [],
        notes: data.notes || [],
        learningNotes: data.learningNotes || [],
        milestones: data.milestones || [],
        dailyPlanner: data.dailyPlanner || [],
        codeSnippets: data.codeSnippets || [],
        journal: data.journal && typeof data.journal === 'object' ? data.journal : {}
    };

    const locationItems = completeData.locationItems || [];
    const blobKeys = [];
    locationItems.forEach(item => {
        (item.attachments || []).forEach(att => {
            if (att.storedInIndexedDB) blobKeys.push({ locationId: item.id, attachmentId: att.id });
        });
    });

    const exportAndDownload = (attachmentBlobs) => {
        const exportData = {
            version: APP_VERSION,
            exportDate: new Date().toISOString(),
        applicationName: 'Task Management System',
            data: completeData,
            attachmentBlobs: attachmentBlobs || {}
        };
        const filename = `todo-app-backup-${formatDateString(new Date())}.json`;
        downloadFile(JSON.stringify(exportData, null, 2), filename, 'application/json');
        alert('✅ Data exported successfully!\n\nFile: ' + filename + '\n\nSave this file and you can import it in any browser to restore all your data.');
    };

    if (blobKeys.length === 0) {
        exportAndDownload({});
        return;
    }

    Promise.all(blobKeys.map(({ locationId, attachmentId }) =>
        getAttachmentBlob(locationId, attachmentId).then(data => ({ key: `${locationId}_${attachmentId}`, data }))
    )).then(results => {
        const attachmentBlobs = {};
        results.forEach(r => { if (r && r.data) attachmentBlobs[r.key] = r.data; });
        exportAndDownload(attachmentBlobs);
    }).catch(err => {
        console.warn('Some attachment blobs could not be read; exporting without them.', err);
        exportAndDownload({});
    });
}

function importAllData(event) {
    const file = event.target.files[0];
    if (!file) return;

    // Check file type
    if (!file.name.toLowerCase().endsWith('.json')) {
        alert('Error: Please select a valid JSON backup file.');
        event.target.value = '';
        return;
    }

    if (!confirm('⚠️ WARNING: This will replace ALL current data!\n\n' +
        '• All tasks\n' +
        '• All users\n' +
        '• All locations\n' +
        '• All settings\n' +
        '• Everything will be replaced with the backup data\n\n' +
        'Are you sure you want to continue?')) {
        event.target.value = '';
        return;
    }

    const reader = new FileReader();
    reader.onload = async function (e) {
        try {
            const imported = JSON.parse(e.target.result);
            let data = imported.data || imported; // Support both new format (v2.0) and old format

            // Validate data structure - ensure all required fields exist
            if (!data) {
                throw new Error('Invalid backup file: No data found');
            }

            // Ensure all required arrays exist, initialize empty if missing
            const completeData = {
                users: Array.isArray(data.users) ? data.users : [],
                tasks: Array.isArray(data.tasks) ? data.tasks : [],
                locations: Array.isArray(data.locations) ? data.locations : [],
                locationItems: Array.isArray(data.locationItems) ? data.locationItems : [],
                segregationTypes: Array.isArray(data.segregationTypes) ? data.segregationTypes : [],
                holidays: Array.isArray(data.holidays) ? data.holidays : [],
                notes: Array.isArray(data.notes) ? data.notes : [],
                learningNotes: Array.isArray(data.learningNotes) ? data.learningNotes : [],
                milestones: Array.isArray(data.milestones) ? data.milestones : [],
                dailyPlanner: Array.isArray(data.dailyPlanner) ? data.dailyPlanner : [],
                codeSnippets: Array.isArray(data.codeSnippets) ? data.codeSnippets : [],
                journal: data.journal && typeof data.journal === 'object' ? data.journal : {}
            };

            // Validate that we have at least some structure
            if (typeof completeData.users !== 'object' || typeof completeData.tasks !== 'object') {
                throw new Error('Invalid backup file: Missing required data structure');
            }

            // Show summary before import
            const summary = `Backup Summary:\n` +
                `• Users: ${completeData.users.length}\n` +
                `• Tasks: ${completeData.tasks.length}\n` +
                `• Locations: ${completeData.locations.length}\n` +
                `• Location Items (Paths): ${completeData.locationItems?.length || 0}\n` +
                `• Segregation Types: ${completeData.segregationTypes.length}\n` +
                `• Holidays: ${completeData.holidays.length}\n` +
                `• Notes: ${completeData.notes.length}\n` +
                `• Learning Notes: ${completeData.learningNotes.length}\n` +
                `• Milestones: ${completeData.milestones.length}\n` +
                `• Daily Planner: ${completeData.dailyPlanner.length}\n` +
                `• Code Snippets: ${completeData.codeSnippets?.length || 0}\n` +
                `\nExport Date: ${imported.exportDate || 'Unknown'}`;

            if (confirm(summary + '\n\nProceed with import?')) {
                if (isApiMode() && currentUser && currentUser.role === 'admin' && !currentUser.isMaster) {
                    try {
                        const res = await apiFetch('/api/workspace/restore', {
                            method: 'POST',
                            body: JSON.stringify(imported)
                        });
                        if (!res.ok) {
                            const err = await res.json().catch(() => ({}));
                            alert(err.error || 'Server restore failed.');
                            event.target.value = '';
                            return;
                        }
                        const merged = await res.json();
                        __workspaceCache = merged;
                        alert('✅ Cloud workspace restored from backup file.\n\nYour session stays active.');
                        event.target.value = '';
                        init();
                    } catch (err) {
                        console.error(err);
                        alert('❌ Restore failed: ' + (err.message || err));
                        event.target.value = '';
                    }
                    return;
                }

                const attachmentBlobs = imported.attachmentBlobs || {};
                const normalizedData = normalizeData(completeData);
                const keys = Object.keys(attachmentBlobs);
                let attachmentWarn = false;

                if (keys.length > 0) {
                    try {
                        await Promise.all(keys.map(key => {
                            const idx = key.indexOf('_');
                            const locationId = parseInt(key.substring(0, idx), 10);
                            const attachmentId = key.substring(idx + 1);
                            const attIdNum = /^\d+\.?\d*$/.test(attachmentId) ? parseFloat(attachmentId) : attachmentId;
                            return putAttachmentBlob(locationId, attIdNum, attachmentBlobs[key]);
                        }));
                    } catch (err) {
                        console.warn('Some attachment blobs could not be restored.', err);
                        attachmentWarn = true;
                    }
                }

                saveData(normalizedData);

                if (isApiMode() && currentUser) {
                    let ok = await flushWorkspaceToApiNow();
                    if (!ok) ok = await flushWorkspaceToApiNow();
                    if (!ok) {
                        alert('❌ Import could not be saved to the server.\n\nYou are still signed in; do not refresh this page yet. Check your connection, then try importing the file again.');
                        event.target.value = '';
                        return;
                    }
                }

                localStorage.removeItem('currentUser');
                currentUser = null;
                if (isApiMode()) __workspaceCache = null;

                if (attachmentWarn) {
                    alert('✅ Data imported with warnings. Some file attachments may be missing.\n\nYou will need to login again.');
                } else {
                    alert('✅ Data imported successfully!\n\nAll data has been restored from the backup file.\n\nYou will need to login again.');
                }
                location.reload();
            } else {
                alert('Import cancelled.');
            }
        } catch (error) {
            alert('❌ Error importing data!\n\n' + error.message + '\n\nPlease ensure the file is a valid backup file created by this application.');
            console.error('Import error:', error);
            console.error('File content:', e.target.result.substring(0, 500)); // Log first 500 chars for debugging
        }
        event.target.value = '';
    };

    reader.onerror = function () {
        alert('❌ Error reading file. Please try again.');
        event.target.value = '';
    };

    reader.readAsText(file);
}

function autoExportData(data) {
    try {
        const completeData = {
            users: data.users || [],
            tasks: data.tasks || [],
            locations: data.locations || [],
            locationItems: data.locationItems || [],
            segregationTypes: data.segregationTypes || [],
            holidays: data.holidays || [],
            notes: data.notes || [],
            learningNotes: data.learningNotes || [],
            milestones: data.milestones || [],
            dailyPlanner: data.dailyPlanner || [],
            codeSnippets: data.codeSnippets || [],
            journal: data.journal && typeof data.journal === 'object' ? data.journal : {}
        };

        const exportData = {
            version: APP_VERSION,
            exportDate: new Date().toISOString(),
        applicationName: 'Task Management System',
            data: completeData
        };

        // Store auto-export in localStorage (limited size, but works)
        const exportString = JSON.stringify(exportData);
        if (exportString.length < 5000000) { // 5MB limit for localStorage
            localStorage.setItem('todoAppAutoBackup', exportString);
        }
    } catch (error) {
        console.warn('Auto-export failed:', error);
    }
}

function loadAutoBackup() {
    try {
        const backup = localStorage.getItem('todoAppAutoBackup');
        if (backup) {
            const imported = JSON.parse(backup);
            const data = imported.data || imported;
            if (data && (data.users || data.tasks || data.locations)) {
                // Return normalized data to ensure all fields exist
                return normalizeData(data);
            }
        }
    } catch (error) {
        console.warn('Auto-backup load failed:', error);
    }
    return null;
}

function toggleAutoExport() {
    const enabled = document.getElementById('autoExportEnabled').checked;
    localStorage.setItem('autoExportEnabled', enabled.toString());
    if (enabled) {
        // Export current data immediately
        const data = getData();
        autoExportData(data);
        alert('Auto-export enabled! Your data will be backed up automatically.');
    } else {
        localStorage.removeItem('todoAppAutoBackup');
        alert('Auto-export disabled.');
    }
}

function showClearDataModal() {
    document.getElementById('clearDataModal').classList.add('active');
    document.getElementById('clearDataPassword').value = '';
    document.getElementById('clearDataError').textContent = '';
}

function closeClearDataModal() {
    document.getElementById('clearDataModal').classList.remove('active');
    document.getElementById('clearDataPassword').value = '';
    document.getElementById('clearDataError').textContent = '';
}

function confirmClearAllData() {
    const password = document.getElementById('clearDataPassword').value;
    const errorDiv = document.getElementById('clearDataError');

    if (!password) {
        errorDiv.textContent = 'Please enter your admin password.';
        return;
    }

    // Verify admin password
    const data = getData();
    const adminUser = data.users.find(u => u.role === 'admin');

    if (!adminUser) {
        errorDiv.textContent = 'No admin user found.';
        return;
    }

    // Simple password check (in production, use proper hashing)
    if (adminUser.password !== password) {
        errorDiv.textContent = 'Incorrect password. Please try again.';
        return;
    }

    // Final confirmation
    if (!confirm('⚠️ FINAL WARNING: This will delete ALL data permanently. This cannot be undone!\n\nAre you absolutely sure?')) {
        return;
    }

    // Clear all data except users
    const currentData = getData();
    const defaultData = {
        users: currentData.users || [], // Preserve users
        tasks: [],
        locations: [
            { id: 1, name: 'Mundra' },
            { id: 2, name: 'JNPT' },
            { id: 3, name: 'Combine' }
        ],
        locationItems: [],
        segregationTypes: [
            { id: 1, name: 'PSA Reports' },
            { id: 2, name: 'Internal Reports' }
        ],
        holidays: [],
        notes: [],
        milestones: [],
        dailyPlanner: [],
        codeSnippets: []
    };

    saveData(defaultData);
    localStorage.removeItem('currentUser');
    localStorage.removeItem('todoAppAutoBackup');
    localStorage.removeItem('autoExportEnabled');

    // Clear IndexedDB attachment blobs
    openAttachmentsDB().then(db => {
        const tx = db.transaction(ATTACHMENTS_STORE, 'readwrite');
        tx.objectStore(ATTACHMENTS_STORE).clear();
        db.close();
    }).catch(() => {});

    closeClearDataModal();
    alert('All data has been cleared. You will be redirected to the login page.');
    location.reload();
}

function downloadFile(content, filename, contentType) {
    const blob = new Blob([content], { type: contentType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}

// CSV Upload Functions
function downloadTaskTemplate() {
    const headers = [
        'Task Number',
        'Task Name',
        'Description',
        'Assigned to (Name)',
        'Location',
        'Task Type',
        'Frequency',
        'Due date calculation type',
        'Due day of month',
        'Recurrence Type',
        'Start date (for due date calculation)',
        'Estimated Minutes',
        'Team task (true/false)'
    ];

    const sampleRows = [
        [
            '1',
            'Sample One Time Task',
            'This is a one-time task description',
            'Amin',
            'JNPT',
            'one_time',
            '',
            '',
            '',
            '',
            '',
            '60',
            'false'
        ],
        [
            '2',
            'Sample Recurring Daily Task',
            'This is a recurring daily task',
            'Amin',
            'Mundra',
            'recurring',
            'daily',
            'calendar_day',
            '1',
            '',
            '01-01-2024',
            '30',
            'true'
        ],
        [
            '3',
            'Sample Recurring Monthly Task',
            'This is a recurring monthly task',
            'Amin',
            'Combine',
            'recurring',
            'monthly',
            'calendar_day',
            '15',
            '',
            '01-01-2024',
            '45',
            'false'
        ],
        [
            '4',
            'Sample Work Plan Task',
            'This is a work plan task',
            'Amin',
            'JNPT',
            'work_plan',
            '',
            '',
            '',
            '',
            '',
            '90',
            'true'
        ],
        [
            '5',
            'Sample Without Due Date Task',
            'Task without due date',
            'Amin',
            'Mundra',
            'without_due_date',
            '',
            '',
            '',
            '',
            '',
            '45',
            'false'
        ]
    ];

    // Create CSV content
    let csv = headers.map(h => escapeCSV(h)).join(',') + '\n';
    sampleRows.forEach(row => {
        csv += row.map(cell => escapeCSV(cell || '')).join(',') + '\n';
    });

    downloadFile(csv, 'task-upload-template.csv', 'text/csv');
}

function handleTaskCSVUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const csv = e.target.result;
            const result = parseAndImportTasks(csv);
            showUploadMessage(result);

            // Clear the file input
            event.target.value = '';
        } catch (error) {
            showUploadMessage({
                success: false,
                message: 'Error reading file: ' + error.message,
                imported: 0,
                errors: []
            });
            event.target.value = '';
        }
    };
    reader.readAsText(file);
}

function parseAndImportTasks(csvContent) {
    const lines = csvContent.split('\n').filter(line => line.trim());
    if (lines.length < 2) {
        return {
            success: false,
            message: 'CSV file must have at least a header row and one data row.',
            imported: 0,
            errors: []
        };
    }

    // Parse header
    const headers = parseCSVLine(lines[0]);
    const expectedHeaders = [
        'Task Name',
        'Assigned to (Name)',
        'Location',
        'Task Type'
    ];

    // Check required headers
    const missingHeaders = expectedHeaders.filter(h => !headers.includes(h));
    if (missingHeaders.length > 0) {
        return {
            success: false,
            message: `Missing required columns: ${missingHeaders.join(', ')}`,
            imported: 0,
            errors: []
        };
    }

    const data = getData();
    const errors = [];
    const importedTasks = [];

    // Get column indices
    const colIndex = {
        taskNumber: headers.indexOf('Task Number'),
        taskName: headers.indexOf('Task Name'),
        description: headers.indexOf('Description'),
        assignedToName: headers.indexOf('Assigned to (Name)'),
        locationName: headers.indexOf('Location'),
        taskType: headers.indexOf('Task Type'),
        frequency: headers.indexOf('Frequency'),
        dueDateType: headers.indexOf('Due date calculation type'),
        dueDay: headers.indexOf('Due day of month'),
        recurrenceType: headers.indexOf('Recurrence Type'),
        startDate: headers.indexOf('Start date (for due date calculation)'),
        estMinutes: headers.indexOf('Estimated Minutes'),
        isTeamTask: headers.indexOf('Team task (true/false)')
    };

    // Process each row
    for (let i = 1; i < lines.length; i++) {
        const row = parseCSVLine(lines[i]);
        if (row.every(cell => !cell.trim())) continue; // Skip empty rows

        const rowNum = i + 1;
        const rowErrors = [];

        // Validate required fields
        const taskName = row[colIndex.taskName]?.trim();
        if (!taskName) {
            rowErrors.push('Task Name is required');
        }

        const assignedToName = row[colIndex.assignedToName]?.trim();
        if (!assignedToName) {
            rowErrors.push('Assigned to (Name) is required');
        }

        const locationName = row[colIndex.locationName]?.trim();
        if (!locationName) {
            rowErrors.push('Location is required');
        }

        const taskType = row[colIndex.taskType]?.trim().toLowerCase();
        if (!taskType || !['one_time', 'recurring', 'without_due_date', 'work_plan', 'audit_point'].includes(taskType)) {
            rowErrors.push('Task Type must be: one_time, recurring, without_due_date, work_plan, or audit_point');
        }

        if (rowErrors.length > 0) {
            errors.push({ row: rowNum, errors: rowErrors });
            continue;
        }

        // Find user by name
        const user = data.users.find(u => u.name.toLowerCase() === assignedToName.toLowerCase());
        if (!user) {
            errors.push({ row: rowNum, errors: [`User with name "${assignedToName}" not found`] });
            continue;
        }

        // Find location by name
        const location = data.locations.find(l => l.name.toLowerCase() === locationName.toLowerCase());
        if (!location) {
            errors.push({ row: rowNum, errors: [`Location "${locationName}" not found`] });
            continue;
        }

        // Parse task data
        let dueDate = null;
        let nextDueDate = null;
        let priority = null;
        let frequency = null;
        let dueDateType = null;
        let dueDay = null;
        let startDate = null;
        let recurrenceType = null;
        let estMinutes = null;
        let isTeamTask = false;

        // Parse recurring task fields
        if (taskType === 'recurring') {
            if (row[colIndex.frequency]) {
                const freqVal = row[colIndex.frequency].trim().toLowerCase();
                if (['daily', 'weekly', 'monthly', 'yearly'].includes(freqVal)) {
                    frequency = freqVal;
                } else {
                    rowErrors.push('Frequency must be: daily, weekly, monthly, or yearly');
                }
            }

            if (row[colIndex.dueDateType]) {
                const dueDateTypeVal = row[colIndex.dueDateType].trim().toLowerCase();
                if (['calendar_day', 'working_day'].includes(dueDateTypeVal)) {
                    dueDateType = dueDateTypeVal;
                } else {
                    rowErrors.push('Due date calculation type must be: calendar_day or working_day');
                }
            }

            if (row[colIndex.dueDay]) {
                const dueDayVal = parseInt(row[colIndex.dueDay].trim());
                if (!isNaN(dueDayVal) && dueDayVal >= 1 && dueDayVal <= 31) {
                    dueDay = dueDayVal;
                } else {
                    rowErrors.push('Due day of month must be a number between 1 and 31');
                }
            }

            // Parse recurrence type (if frequency is not set)
            if (row[colIndex.recurrenceType] && !frequency) {
                const recTypeVal = row[colIndex.recurrenceType].trim().toLowerCase();
                if (['calendar_day', 'working_day'].includes(recTypeVal)) {
                    recurrenceType = recTypeVal;
                }
            }

            if (row[colIndex.startDate]) {
                const startDateStr = row[colIndex.startDate].trim();
                const startDateMatch = startDateStr.match(/(\d{2})-(\d{2})-(\d{4})/);
                if (startDateMatch) {
                    startDate = `${startDateMatch[3]}-${startDateMatch[2]}-${startDateMatch[1]}`;
                } else {
                    rowErrors.push('Invalid Start date format. Use DD-MM-YYYY');
                }
            }

            // Calculate next due date for recurring tasks
            if (startDate && frequency && dueDateType && dueDay) {
                nextDueDate = calculateRecurringDueDate(startDate, frequency, dueDateType, dueDay);
            }
        }

        // Parse estimated minutes
        if (row[colIndex.estMinutes]) {
            const estMinVal = parseInt(row[colIndex.estMinutes].trim());
            if (!isNaN(estMinVal) && estMinVal > 0) {
                estMinutes = estMinVal;
            }
        }

        // Parse is team task
        if (row[colIndex.isTeamTask]) {
            const teamTaskVal = row[colIndex.isTeamTask].trim().toLowerCase();
            isTeamTask = teamTaskVal === 'true' || teamTaskVal === '1' || teamTaskVal === 'yes';
        }

        if (rowErrors.length > 0) {
            errors.push({ row: rowNum, errors: rowErrors });
            continue;
        }

        let taskNumber = null;
        if (colIndex.taskNumber >= 0 && row[colIndex.taskNumber] != null && row[colIndex.taskNumber].toString().trim()) {
            const n = parseInt(row[colIndex.taskNumber].toString().trim(), 10);
            if (!isNaN(n) && n > 0) taskNumber = n;
        }

        // Create task object
        const task = {
            id: Date.now() + i, // Unique ID
            task_number: taskNumber,
            task_name: taskName,
            description: row[colIndex.description]?.trim() || '',
            assigned_to: user.id,
            location_id: location.id,
            task_type: taskType,
            due_date: dueDate,
            priority: null, // Priority not in new CSV format
            frequency: frequency,
            due_date_type: dueDateType,
            due_day: dueDay,
            start_date: startDate,
            recurrence_type: recurrenceType,
            recurrence_interval: null,
            next_due_date: nextDueDate,
            segregation_type_id: null, // Segregation type not in new CSV format
            est_minutes: estMinutes,
            is_team_task: isTeamTask,
            task_action: 'not_completed',
            comment: null,
            recurrence_stopped: false,
            created_by: currentUser.id,
            created_at: new Date().toISOString(),
            completed_at: null
        };

        importedTasks.push(task);
    }

    // Import valid tasks (assign task_number to any task that doesn't have one)
    if (importedTasks.length > 0) {
        updateData(data => {
            let nextNum = getNextTaskNumberFromData(data);
            importedTasks.forEach(task => {
                if (task.task_number == null || task.task_number === '') {
                    task.task_number = nextNum++;
                }
                data.tasks.push(task);
            });
        });

        // Process recurring tasks
        processRecurringTasks();

        // Refresh UI
        renderTasks();
        renderDashboard();
        renderCalendar();
        renderInteractiveDashboard();
    }

    return {
        success: errors.length === 0,
        message: `Imported ${importedTasks.length} task(s) successfully${errors.length > 0 ? `, ${errors.length} row(s) had errors` : ''}.`,
        imported: importedTasks.length,
        errors: errors
    };
}

function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
        const char = line[i];

        if (char === '"') {
            if (inQuotes && line[i + 1] === '"') {
                current += '"';
                i++; // Skip next quote
            } else {
                inQuotes = !inQuotes;
            }
        } else if (char === ',' && !inQuotes) {
            result.push(current);
            current = '';
        } else {
            current += char;
        }
    }
    result.push(current);
    return result;
}

function showUploadMessage(result) {
    const messageDiv = document.getElementById('taskUploadMessage');
    if (!messageDiv) return;

    let html = '';
    if (result.success) {
        html = `<div style="padding: 15px; background: #d4edda; border: 1px solid #c3e6cb; border-radius: 5px; color: #155724;">
            <strong>✓ Success!</strong> ${result.message}
        </div>`;
    } else {
        html = `<div style="padding: 15px; background: ${result.imported > 0 ? '#fff3cd' : '#f8d7da'}; border: 1px solid ${result.imported > 0 ? '#ffeaa7' : '#f5c6cb'}; border-radius: 5px; color: ${result.imported > 0 ? '#856404' : '#721c24'};">
            <strong>${result.imported > 0 ? '⚠ Partial Success' : '✗ Error'}</strong> ${result.message}
        </div>`;

        if (result.errors.length > 0) {
            html += '<div style="margin-top: 10px; max-height: 200px; overflow-y: auto;">';
            result.errors.forEach(err => {
                html += `<div style="font-size: 12px; margin: 5px 0; padding: 5px; background: rgba(0,0,0,0.05); border-radius: 3px;">
                    <strong>Row ${err.row}:</strong> ${err.errors.join(', ')}
                </div>`;
            });
            html += '</div>';
        }
    }

    messageDiv.innerHTML = html;

    // Auto-hide after 10 seconds
    setTimeout(() => {
        if (messageDiv.innerHTML === html) {
            messageDiv.innerHTML = '';
        }
    }, 10000);
}

// Notes Functions
let showHiddenNotes = false; // Toggle to show/hide notes that are marked as hidden

function renderNotes() {
    const data = getData();
    const searchFilter = document.getElementById('noteSearch')?.value.toLowerCase() || '';

    let filteredNotes = data.notes || [];

    // Filter by search
    if (searchFilter) {
        filteredNotes = filteredNotes.filter(note =>
            note.content.toLowerCase().includes(searchFilter)
        );
    }

    // Filter by visibility (unless showHiddenNotes is true)
    if (!showHiddenNotes) {
        filteredNotes = filteredNotes.filter(note => !note.is_hidden);
    }

    // Sort by created date (newest first)
    filteredNotes.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));

    const notesHtml = filteredNotes.length > 0
        ? filteredNotes.map(note => renderNoteCard(note)).join('')
        : '<p style="text-align: center; color: #999; padding: 40px; grid-column: 1 / -1;">No notes found. Click "+ Add Note" to create your first note!</p>';

    document.getElementById('notesList').innerHTML = notesHtml;
}

function renderNoteCard(note) {
    const hiddenClass = note.is_hidden && !showHiddenNotes ? 'hidden' : '';
    const hiddenStyle = note.is_hidden ? 'opacity: 0.6; border: 2px dashed #999;' : '';
    const contentHtml = note.is_password_protected
        ? '<div class="note-content" style="color: #666; font-style: italic;">🔒 Content protected – open to view</div>'
        : `<div class="note-content" style="white-space: pre-wrap; word-break: break-word;">${escapeHtml((note.content || '').replace(/<[^>]*>/g, ' ').trim())}</div>`;

    return `
        <div class="note-card ${hiddenClass}" style="background: ${note.color || '#fff3cd'}; ${hiddenStyle}">
            <div class="note-header">
                <div style="flex: 1;">
                    ${note.is_hidden ? '<span style="font-size: 11px; color: #999; font-weight: 600;">(Hidden)</span>' : ''}
                </div>
            </div>
            ${contentHtml}
            <div class="note-footer">
                <div class="note-date">${formatDateDisplay(note.created_at)}</div>
                <div class="note-actions">
                    <button class="btn btn-primary" onclick="editNote(${note.id})" style="background: #667eea; color: white; padding: 4px 8px; font-size: 11px;">Edit</button>
                    <button class="btn btn-secondary" onclick="toggleNoteVisibility(${note.id})" style="background: ${note.is_hidden ? '#28a745' : '#6c757d'}; color: white; padding: 4px 8px; font-size: 11px;">${note.is_hidden ? 'Show' : 'Hide'}</button>
                    <button class="btn btn-danger" onclick="deleteNote(${note.id})" style="background: #dc3545; color: white; padding: 4px 8px; font-size: 11px;">Delete</button>
                </div>
            </div>
        </div>
    `;
}

function openNoteModal(noteId = null) {
    const modal = document.getElementById('noteModal');
    const form = document.getElementById('noteForm');
    const contentEditor = document.getElementById('noteContentEditor');

    if (noteId) {
        const data = getData();
        const note = data.notes.find(n => n.id === noteId);
        if (note) {
            document.getElementById('noteModalTitle').textContent = 'Edit Note';
            document.getElementById('noteId').value = note.id;

            // Load title and category
            document.getElementById('noteTitle').value = note.title || '';
            document.getElementById('noteCategory').value = note.category || '';

            // Load content (handle both plain text and HTML)
            if (contentEditor) {
                contentEditor.innerHTML = note.content || '';
            }

            // Set color radio
            const colorRadios = document.querySelectorAll('input[name="noteColor"]');
            colorRadios.forEach(radio => {
                if (radio.value === note.color) {
                    radio.checked = true;
                }
            });
            const pwdCheck = document.getElementById('notePasswordProtected');
            if (pwdCheck) pwdCheck.checked = !!note.is_password_protected;
        }
    } else {
        form.reset();
        document.getElementById('noteModalTitle').textContent = 'New Note';
        document.getElementById('noteId').value = '';
        document.getElementById('noteTitle').value = '';
        document.getElementById('noteCategory').value = '';
        if (contentEditor) {
            contentEditor.innerHTML = '';
        }
        const pwdCheck = document.getElementById('notePasswordProtected');
        if (pwdCheck) pwdCheck.checked = false;
        // Set default color
        const defaultColorRadio = document.querySelector('input[name="noteColor"][value="#fff3cd"]');
        if (defaultColorRadio) defaultColorRadio.checked = true;
    }

    modal.classList.add('active');
}

function closeNoteModal() {
    document.getElementById('noteModal').classList.remove('active');
    document.getElementById('noteForm').reset();
}

function saveNote(event) {
    if (event) {
        event.preventDefault();
        event.stopPropagation();
    }

    try {
        const noteIdElement = document.getElementById('noteId');
        const noteTitleElement = document.getElementById('noteTitle');
        const noteCategoryElement = document.getElementById('noteCategory');
        const noteContentEditor = document.getElementById('noteContentEditor');
        const colorRadio = document.querySelector('input[name="noteColor"]:checked');

        if (!noteContentEditor) {
            alert('Error: Note content editor not found.');
            console.error('noteContentEditor element not found');
            return;
        }

        const noteId = noteIdElement ? noteIdElement.value : '';
        const title = noteTitleElement ? noteTitleElement.value.trim() : '';
        const category = noteCategoryElement ? noteCategoryElement.value : '';
        const content = noteContentEditor.innerHTML.trim();
        const color = colorRadio ? colorRadio.value : '#fff3cd';
        const isPasswordProtected = document.getElementById('notePasswordProtected') ? document.getElementById('notePasswordProtected').checked : false;

        // Content is required (check if it has meaningful text, not just empty tags)
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = content;
        const textContent = tempDiv.textContent || tempDiv.innerText || '';

        if (!textContent.trim()) {
            alert('Content is required.');
            return;
        }

        updateData(data => {
            if (!data.notes) {
                data.notes = [];
            }

            if (noteId) {
                const note = data.notes.find(n => n.id === parseInt(noteId));
                if (note) {
                    note.title = title;
                    note.category = category;
                    note.content = content;
                    note.color = color;
                    note.is_password_protected = isPasswordProtected;
                    note.updated_at = new Date().toISOString();
                } else {
                    console.error('Note not found for ID:', noteId);
                }
            } else {
                const newNote = {
                    id: Date.now(),
                    title: title,
                    category: category,
                    content: content,
                    color: color,
                    is_hidden: false,
                    is_password_protected: isPasswordProtected,
                    created_at: new Date().toISOString(),
                    updated_at: null
                };
                data.notes.push(newNote);
            }
        });

        closeNoteModal();
        renderNotes();
    } catch (error) {
        console.error('Error saving note:', error);
        alert('An error occurred while saving the note. Please check the console for details.');
    }
}

function editNote(noteId) {
    openNoteModal(noteId);
}

function deleteNote(noteId) {
    if (!confirm('Are you sure you want to delete this note?')) return;

    updateData(data => {
        data.notes = data.notes.filter(n => n.id !== noteId);
    });

    renderNotes();
}

function toggleNotesVisibility() {
    showHiddenNotes = !showHiddenNotes;
    const btn = document.getElementById('toggleNotesBtn');
    btn.textContent = showHiddenNotes ? 'Hide Hidden Notes' : 'Show Hidden Notes';
    renderNotes();
}

function toggleNoteVisibility(noteId) {
    updateData(data => {
        const note = data.notes.find(n => n.id === noteId);
        if (note) {
            note.is_hidden = !note.is_hidden;
        }
    });
    renderNotes();
}

// Rich text formatting function
function formatText(command, value = null) {
    const editor = document.getElementById('noteContentEditor');
    if (!editor) return;

    editor.focus();

    if (value) {
        document.execCommand(command, false, value);
    } else {
        document.execCommand(command, false, null);
    }
}

// Global variable for category filter
let currentNoteCategory = '';

// Filter notes by category
function filterNotesByCategory(category) {
    currentNoteCategory = category;

    // Update active state of filter buttons
    const filterButtons = document.querySelectorAll('.note-category-filter');
    filterButtons.forEach(btn => {
        if (btn.getAttribute('data-category') === category) {
            btn.classList.add('active');
            btn.style.opacity = '1';
            btn.style.transform = 'scale(1.05)';
        } else {
            btn.classList.remove('active');
            btn.style.opacity = '0.7';
            btn.style.transform = 'scale(1)';
        }
    });

    renderNotes();
}

// Category color mapping
function getCategoryColor(category) {
    const colors = {
        'Work': '#007bff',
        'Systems': '#28a745',
        'Improvement': '#ffc107',
        'Planning': '#6f42c1',
        'Pass': '#17a2b8',
        'Other': '#6c757d'
    };
    return colors[category] || '#6c757d';
}

// Render notes with enhanced features
function renderNotes() {
    const data = getData();
    const container = document.getElementById('notesList');
    const searchTerm = document.getElementById('noteSearch').value.toLowerCase();

    if (!container) return;

    // Filter notes
    let notes = data.notes || [];

    // Filter by category
    if (currentNoteCategory) {
        notes = notes.filter(n => n.category === currentNoteCategory);
    }

    // Filter by search term
    if (searchTerm) {
        notes = notes.filter(n => {
            const title = (n.title || '').toLowerCase();
            const content = (n.content || '').toLowerCase();
            const category = (n.category || '').toLowerCase();
            return title.includes(searchTerm) || content.includes(searchTerm) || category.includes(searchTerm);
        });
    }

    // Filter by visibility
    notes = notes.filter(n => showHiddenNotes || !n.is_hidden);

    // Sort by updated date (most recent first)
    notes.sort((a, b) => {
        const dateA = new Date(a.updated_at || a.created_at);
        const dateB = new Date(b.updated_at || b.created_at);
        return dateB - dateA;
    });

    if (notes.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #999; padding: 40px;">No notes found.</p>';
        return;
    }

    container.innerHTML = notes.map(note => {
        const date = new Date(note.updated_at || note.created_at);
        const dateStr = date.toLocaleDateString();
        const title = note.title || 'Untitled Note';
        const category = note.category || '';
        const categoryBadge = category ?
            `<span style="display: inline-block; padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: 500; background: ${getCategoryColor(category)}; color: white; margin-bottom: 8px;">${category}</span>`
            : '';

        const isPasswordProtected = !!note.is_password_protected;
        let contentPreviewHtml;
        if (isPasswordProtected) {
            contentPreviewHtml = '<div class="note-content" style="flex: 1; color: #666; font-size: 14px; font-style: italic;">🔒 Content protected – open to view</div>';
        } else {
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = note.content || '';
            const plainText = (tempDiv.textContent || tempDiv.innerText || '').trim();
            contentPreviewHtml = `<div class="note-content" style="flex: 1; color: #555; font-size: 14px; line-height: 1.6; overflow: hidden; text-overflow: ellipsis; display: -webkit-box; -webkit-line-clamp: 6; -webkit-box-orient: vertical; white-space: pre-wrap; word-break: break-word;">${escapeHtml(plainText)}</div>`;
        }

        return `
            <div class="note-card ${note.is_hidden ? 'hidden' : ''}" 
                 style="background: ${note.color}; border-radius: 10px; padding: 15px; box-shadow: 0 3px 10px rgba(0,0,0,0.1); 
                        transition: transform 0.2s, box-shadow 0.2s; position: relative; 
                        height: 250px; display: flex; flex-direction: column; cursor: pointer;"
                 onclick="expandNote(${note.id})"
                 onmouseenter="this.style.transform='translateY(-3px)'; this.style.boxShadow='0 5px 15px rgba(0,0,0,0.2)';"
                 onmouseleave="this.style.transform='translateY(0)'; this.style.boxShadow='0 3px 10px rgba(0,0,0,0.1)';">
                
                <div class="note-header" style="margin-bottom: 10px;">
                    ${categoryBadge}
                    <div class="note-title" style="font-weight: 600; font-size: 16px; color: #333; margin-bottom: 5px;">
                        ${title}
                    </div>
                </div>

                ${contentPreviewHtml}

                <div class="note-footer" style="display: flex; justify-content: space-between; align-items: center; 
                                                margin-top: 10px; padding-top: 10px; border-top: 1px solid rgba(0,0,0,0.1);">
                    <div class="note-date" style="font-size: 11px; color: #999;">
                        ${dateStr}
                    </div>
                    <div class="note-actions" style="display: flex; gap: 5px;" onclick="event.stopPropagation();">
                        <button onclick="editNote(${note.id})" 
                                style="padding: 4px 8px; font-size: 11px; background: #667eea; color: white; border: none; 
                                       border-radius: 4px; cursor: pointer; transition: opacity 0.2s;"
                                onmouseenter="this.style.opacity='0.8'" 
                                onmouseleave="this.style.opacity='1'">Edit</button>
                        <button onclick="toggleNoteVisibility(${note.id})" 
                                style="padding: 4px 8px; font-size: 11px; background: #ffc107; color: #333; border: none; 
                                       border-radius: 4px; cursor: pointer; transition: opacity 0.2s;"
                                onmouseenter="this.style.opacity='0.8'" 
                                onmouseleave="this.style.opacity='1'">${note.is_hidden ? 'Show' : 'Hide'}</button>
                        <button onclick="deleteNote(${note.id})" 
                                style="padding: 4px 8px; font-size: 11px; background: #dc3545; color: white; border: none; 
                                       border-radius: 4px; cursor: pointer; transition: opacity 0.2s;"
                                onmouseenter="this.style.opacity='0.8'" 
                                onmouseleave="this.style.opacity='1'">Delete</button>
                    </div>
                </div>
            </div>
        `;
    }).join('');
}

// Expand note in modal to view full content
function expandNote(noteId) {
    const data = getData();
    const note = data.notes.find(n => n.id === noteId);

    if (!note) return;

    const title = note.title || 'Untitled Note';
    const category = note.category || 'No Category';
    const date = new Date(note.updated_at || note.created_at);
    const dateStr = date.toLocaleString();

    // Create modal overlay
    const overlay = document.createElement('div');
    overlay.style.cssText = `
        position: fixed; top: 0; left: 0; right: 0; bottom: 0; 
        background: rgba(0,0,0,0.7); z-index: 10000; display: flex; 
        align-items: center; justify-content: center; padding: 20px;
    `;

    overlay.onclick = (e) => {
        if (e.target === overlay) {
            document.body.removeChild(overlay);
        }
    };

    const modal = document.createElement('div');
    modal.style.cssText = `
        background: ${note.color}; max-width: 800px; width: 100%; 
        max-height: 90vh; overflow-y: auto; border-radius: 15px; 
        padding: 30px; box-shadow: 0 10px 40px rgba(0,0,0,0.3);
    `;

    modal.innerHTML = `
        <div style="display: flex; justify-content: space-between; align-items: start; margin-bottom: 20px;">
            <div>
                <div style="display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 500; 
                            background: ${getCategoryColor(note.category)}; color: white; margin-bottom: 10px;">
                    ${category}
                </div>
                <h2 style="margin: 0; color: #333; font-size: 24px;">${title}</h2>
                <div style="font-size: 12px; color: #666; margin-top: 5px;">${dateStr}</div>
            </div>
            <button onclick="this.closest('.overlay-modal-container').remove()" 
                    style="background: #dc3545; color: white; border: none; border-radius: 50%; 
                           width: 35px; height: 35px; font-size: 24px; cursor: pointer; 
                           display: flex; align-items: center; justify-content: center; line-height: 1;">&times;</button>
        </div>
        <div style="color: #333; font-size: 15px; line-height: 1.7; word-wrap: break-word;">
            ${note.content}
        </div>
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid rgba(0,0,0,0.1); 
                    display: flex; gap: 10px; justify-content: flex-end;">
            <button onclick="editNote(${note.id}); this.closest('.overlay-modal-container').remove();" 
                    style="padding: 10px 20px; background: #667eea; color: white; border: none; 
                           border-radius: 5px; cursor: pointer; font-size: 14px;">Edit Note</button>
            <button onclick="this.closest('.overlay-modal-container').remove()" 
                    style="padding: 10px 20px; background: #6c757d; color: white; border: none; 
                           border-radius: 5px; cursor: pointer; font-size: 14px;">Close</button>
        </div>
    `;

    overlay.className = 'overlay-modal-container';
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
}


// Email Functionality - Format task for email (same format as displayed)
function formatTaskForEmail(task, includeDescription = true) {
    const data = getData();
    const assignedUser = data.users.find(u => u.id === task.assigned_to);
    const location = data.locations.find(l => l.id === task.location_id);
    const segregation = data.segregationTypes.find(s => s.id === task.segregation_type_id);

    // Get priority
    const priority = task.priority ? task.priority.toUpperCase() : '';

    // Get status
    let status = '';
    if (task.task_action === 'completed') {
        if (task.admin_finalized) {
            status = 'Finalized';
        } else if (currentUser.role === 'admin') {
            status = 'Pending Review';
        } else {
            status = 'Completed';
        }
    } else {
        if (task.rejected_at) {
            status = 'Rejected - Pending';
        } else {
            status = 'Pending';
        }
    }

    // Get task type
    let taskType = '';
    if (task.task_type === 'recurring') {
        taskType = task.frequency ? `Recurring - ${task.frequency.charAt(0).toUpperCase() + task.frequency.slice(1)}` : 'Recurring';
    } else if (task.task_type === 'without_due_date') {
        taskType = 'Without Due Date';
    } else if (task.task_type === 'work_plan') {
        taskType = 'Work Plan';
    } else if (task.task_type === 'audit_point') {
        taskType = 'Audit Point';
    } else {
        taskType = 'One Time';
    }

    // Get due date
    const dueDate = task.due_date || task.next_due_date;
    let dueDateStr = 'Not set';
    if (task.task_type === 'without_due_date') {
        dueDateStr = 'No Due Date';
    } else if (dueDate) {
        dueDateStr = formatDateDisplay(dueDate);
    }

    // Build task string
    let taskStr = `Task Name: ${task.task_name}\n`;
    if (priority) taskStr += `Priority: ${priority}\n`;
    taskStr += `Status: ${status}\n`;
    taskStr += `Task Type: ${taskType}\n`;
    taskStr += `Assigned To: ${assignedUser ? assignedUser.name : 'Unknown'}\n`;
    taskStr += `Location: ${location ? location.name : 'Unknown'}\n`;
    taskStr += `Due Date: ${dueDateStr}\n`;
    if (segregation) taskStr += `Segregation Type: ${segregation.name}\n`;
    if (includeDescription && task.description) {
        taskStr += `Description: ${task.description}\n`;
    }
    if (task.comment) {
        taskStr += `User Comment: ${task.comment}\n`;
    }
    if (task.admin_comment) {
        taskStr += `Admin Comment${task.admin_finalized ? ' (Accepted)' : ' (Rejected)'}: ${task.admin_comment}\n`;
    }

    return taskStr;
}

// Format tasks as HTML table for email (matching Interactive Dashboard table format)
function formatTasksAsHTMLTable(tasks) {
    if (!tasks || tasks.length === 0) {
        return '<p>No tasks to display.</p>';
    }

    const data = getData();

    // Start HTML table
    let htmlTable = '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 12px;">';

    // Table header
    htmlTable += '<thead>';
    htmlTable += '<tr style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white;">';
    htmlTable += '<th style="padding: 12px; text-align: left; font-weight: 600;">Task Name</th>';
    htmlTable += '<th style="padding: 12px; text-align: left; font-weight: 600;">Status</th>';
    htmlTable += '<th style="padding: 12px; text-align: left; font-weight: 600;">Assigned To</th>';
    htmlTable += '<th style="padding: 12px; text-align: left; font-weight: 600;">Location</th>';
    htmlTable += '<th style="padding: 12px; text-align: left; font-weight: 600;">Due Date</th>';
    htmlTable += '</tr>';
    htmlTable += '</thead>';

    // Table body
    htmlTable += '<tbody>';

    tasks.forEach((task, index) => {
        const assignedUser = data.users.find(u => u.id === task.assigned_to);
        const location = data.locations.find(l => l.id === task.location_id);
        const dueDate = task.due_date || task.next_due_date;
        const dueDateStr = dueDate ? formatDateDisplay(dueDate) : 'Not set';

        // Determine status (matching renderInteractiveTaskCard logic)
        let statusText = '';
        let statusColor = '';
        if (task.task_action === 'completed') {
            statusText = 'Completed';
            statusColor = '#28a745';
        } else if (task.task_action === 'completed_need_improvement') {
            statusText = 'Need Improvement';
            statusColor = '#ffc107';
        } else {
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            if (dueDate) {
                const dateParts = dueDate.split('-');
                const taskDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
                taskDate.setHours(0, 0, 0, 0);
                if (taskDate < today) {
                    statusText = 'Overdue';
                    statusColor = '#dc3545';
                } else {
                    statusText = 'Pending';
                    statusColor = '#17a2b8';
                }
            } else {
                statusText = 'No Due Date';
                statusColor = '#6c757d';
            }
        }

        // Build status badges HTML
        let statusBadges = `<span style="color: ${statusColor}; font-weight: 600;">${statusText}</span>`;
        if (task.priority) {
            const priorityColors = { high: '#dc3545', medium: '#ffc107', low: '#17a2b8' };
            const priorityColor = priorityColors[task.priority] || '#6c757d';
            statusBadges += ` <span style="background: ${priorityColor}; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px; font-weight: 600;">${task.priority.toUpperCase()}</span>`;
        }
        if (task.frequency) {
            statusBadges += ` <span style="background: #e7f3ff; color: #0066cc; padding: 2px 6px; border-radius: 3px; font-size: 10px;">${task.frequency.charAt(0).toUpperCase() + task.frequency.slice(1)}</span>`;
        }

        // Task name with description
        let taskNameHtml = `<strong>${escapeHtml(task.task_name)}</strong>`;
        if (task.description) {
            taskNameHtml += `<br><span style="color: #666; font-size: 11px;">${escapeHtml(task.description)}</span>`;
        }

        // Row background color (alternating)
        const rowBgColor = index % 2 === 0 ? '#ffffff' : '#f9f9f9';

        htmlTable += `<tr style="background: ${rowBgColor};">`;
        htmlTable += `<td style="padding: 10px;">${taskNameHtml}</td>`;
        htmlTable += `<td style="padding: 10px;">${statusBadges}</td>`;
        htmlTable += `<td style="padding: 10px; color: #666;">${escapeHtml(assignedUser ? assignedUser.name : 'Unknown')}</td>`;
        htmlTable += `<td style="padding: 10px; color: #666;">${escapeHtml(location ? location.name : 'Unknown')}</td>`;
        htmlTable += `<td style="padding: 10px; color: #666;">${dueDateStr}</td>`;
        htmlTable += '</tr>';
    });

    htmlTable += '</tbody>';
    htmlTable += '</table>';

    return htmlTable;
}

// Format tasks for email body - Drilldown format uses same HTML table
// (removed detailed text format, using HTML table for consistency)

// Send email via Outlook using mailto protocol (configured for Windows Desktop Outlook with HTML format)
function sendEmailViaOutlook(tasks, title, source = 'dashboard') {
    if (!tasks || tasks.length === 0) {
        return;
    }

    const data = getData();

    // Get all unique user emails from tasks
    const userEmails = [];
    const userIds = new Set();
    tasks.forEach(task => {
        if (task.assigned_to && !userIds.has(task.assigned_to)) {
            userIds.add(task.assigned_to);
            const user = data.users.find(u => u.id === task.assigned_to);
            if (user && user.email) {
                userEmails.push(user.email);
            }
        }
    });

    // Format email body as HTML table (only table, no extra details)
    const htmlTable = formatTasksAsHTMLTable(tasks);

    // Create subject
    const subject = `${title} - ${tasks.length} Task(s)`;

    // Build mailto link
    const toEmails = userEmails.join(';');
    const encodedSubject = encodeURIComponent(subject);

    // Copy HTML table to clipboard in proper HTML format
    function copyHTMLToClipboard() {
        if (navigator.clipboard && navigator.clipboard.write) {
            // Create HTML blob for clipboard
            const htmlBlob = new Blob([htmlTable], { type: 'text/html' });
            const plainText = htmlTable.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
            const textBlob = new Blob([plainText], { type: 'text/plain' });

            const clipboardItem = new ClipboardItem({
                'text/html': htmlBlob,
                'text/plain': textBlob
            });

            return navigator.clipboard.write([clipboardItem]);
        } else {
            // Fallback: use execCommand
            return new Promise((resolve, reject) => {
                const tempDiv = document.createElement('div');
                tempDiv.contentEditable = true;
                tempDiv.innerHTML = htmlTable;
                tempDiv.style.position = 'fixed';
                tempDiv.style.left = '-9999px';
                document.body.appendChild(tempDiv);

                const range = document.createRange();
                range.selectNodeContents(tempDiv);
                const selection = window.getSelection();
                selection.removeAllRanges();
                selection.addRange(range);

                try {
                    const success = document.execCommand('copy');
                    selection.removeAllRanges();
                    document.body.removeChild(tempDiv);
                    if (success) {
                        resolve();
                    } else {
                        reject(new Error('Copy failed'));
                    }
                } catch (err) {
                    document.body.removeChild(tempDiv);
                    reject(err);
                }
            });
        }
    }

    // Copy HTML to clipboard and open Outlook
    copyHTMLToClipboard().then(() => {
        // Open Outlook with mailto
        let mailtoLink = 'mailto:';
        if (toEmails) {
            mailtoLink += encodeURIComponent(toEmails);
        }
        mailtoLink += `?subject=${encodedSubject}`;

        const link = document.createElement('a');
        link.href = mailtoLink;
        link.style.display = 'none';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }).catch(() => {
        // Fallback: open mailto anyway
        let mailtoLink = 'mailto:';
        if (toEmails) {
            mailtoLink += encodeURIComponent(toEmails);
        }
        mailtoLink += `?subject=${encodedSubject}`;
        window.location.href = mailtoLink;
    });
}

// Send email for Drilldown tab
function sendEmailDrilldown() {
    const tasks = window.drilldownFilteredTasks || window.currentFilteredTasksForExport || [];
    if (tasks.length === 0) {
        alert('No tasks to send. Please click on a dashboard tile to view tasks first.');
        return;
    }

    // Get title from header if available, otherwise use title element
    const headerElement = document.getElementById('drilldownHeader');
    let title = 'Drilldown Tasks';

    if (headerElement) {
        const titleElement = headerElement.querySelector('h3');
        if (titleElement) {
            title = titleElement.textContent.trim();
        } else {
            title = document.getElementById('drilldownTitle')?.textContent || 'Drilldown Tasks';
        }
    } else {
        title = document.getElementById('drilldownTitle')?.textContent || 'Drilldown Tasks';
    }

    // Add task count to title if not already included
    if (tasks.length > 0 && !title.includes(`(${tasks.length})`)) {
        title += ` (${tasks.length} tasks)`;
    }

    sendEmailViaOutlook(tasks, title, 'drilldown');
}

// Send email for Interactive Dashboard tab
function sendEmailInteractiveDashboard() {
    const tasks = window.currentFilteredTasks || [];
    if (tasks.length === 0) {
        alert('No tasks to send. Please apply filters to see tasks first.');
        return;
    }

    // Build title with filter information
    let title = 'Interactive Dashboard Tasks';

    // Get filter values to include in title
    const monthFromFilter = document.getElementById('filterDashboardMonthFrom')?.value;
    const monthToFilter = document.getElementById('filterDashboardMonthTo')?.value;
    const userFilter = document.getElementById('filterDashboardUser')?.value;
    const statusFilter = document.getElementById('filterDashboardStatus')?.value;
    const taskTypeFilter = document.getElementById('filterDashboardTaskType')?.value;

    const data = getData();
    const filters = [];

    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'];
    if (monthFromFilter && monthToFilter) {
        const [y1, m1] = monthFromFilter.split('-').map(Number);
        const [y2, m2] = monthToFilter.split('-').map(Number);
        filters.push(`${monthNames[m1 - 1]} ${y1} - ${monthNames[m2 - 1]} ${y2}`);
    } else if (monthFromFilter) {
        const [y, m] = monthFromFilter.split('-').map(Number);
        filters.push(`${monthNames[m - 1]} ${y}`);
    }

    if (userFilter) {
        const user = data.users.find(u => u.id === parseInt(userFilter));
        if (user) filters.push(`User: ${user.name}`);
    }

    if (statusFilter) {
        filters.push(`Status: ${statusFilter.charAt(0).toUpperCase() + statusFilter.slice(1)}`);
    }

    if (taskTypeFilter) {
        const typeLabels = {
            'one_time': 'One Time',
            'recurring': 'Recurring',
            'without_due_date': 'Without Due Date',
            'work_plan': 'Work Plan',
            'audit_point': 'Audit Point'
        };
        filters.push(`Type: ${typeLabels[taskTypeFilter] || taskTypeFilter}`);
    }

    if (filters.length > 0) {
        title += ` - ${filters.join(', ')}`;
    }

    title += ` (${tasks.length} tasks)`;

    sendEmailViaOutlook(tasks, title, 'interactive');
}

// ==================== MILESTONES TAB ====================

let currentPlannerDate = formatDateString(new Date());

const MILESTONE_YEAR_FILTER_KEY = 'milestoneYearFilter';

function saveMilestoneYearFilter() {
    const el = document.getElementById('milestoneYearFilter');
    if (el) {
        const v = el.value || '';
        if (v) localStorage.setItem(MILESTONE_YEAR_FILTER_KEY, v);
        else localStorage.removeItem(MILESTONE_YEAR_FILTER_KEY);
    }
}

function renderMilestones() {
    const data = getData();
    if (!data.milestones) data.milestones = [];

    // Populate year filter dropdown first
    const yearSelect = document.getElementById('milestoneYearFilter');
    if (yearSelect) {
        const years = new Set();
        data.milestones.forEach(m => {
            if (m.date) {
                const year = m.date.split('-')[0];
                years.add(year);
            }
        });
        const currentYear = new Date().getFullYear();
        years.add(currentYear.toString());

        const savedYear = localStorage.getItem(MILESTONE_YEAR_FILTER_KEY) || '';
        yearSelect.innerHTML = '<option value="">All Years</option>';
        Array.from(years).sort((a, b) => parseInt(b) - parseInt(a)).forEach(year => {
            const option = document.createElement('option');
            option.value = year;
            option.textContent = year;
            yearSelect.appendChild(option);
        });
        // Restore last selected value (if still valid)
        if (savedYear && years.has(savedYear)) {
            yearSelect.value = savedYear;
        }
    }

    const yearFilter = document.getElementById('milestoneYearFilter')?.value || '';
    if (yearFilter) {
        localStorage.setItem(MILESTONE_YEAR_FILTER_KEY, yearFilter);
    } else {
        localStorage.removeItem(MILESTONE_YEAR_FILTER_KEY);
    }

    // Filter milestones
    let filteredMilestones = data.milestones || [];
    if (yearFilter) {
        filteredMilestones = filteredMilestones.filter(m => {
            if (!m.date) return false;
            return m.date.startsWith(yearFilter);
        });
    }

    // Sort by date (newest first)
    filteredMilestones.sort((a, b) => {
        if (!a.date) return 1;
        if (!b.date) return -1;
        return b.date.localeCompare(a.date);
    });

    const milestonesHtml = filteredMilestones.length > 0
        ? filteredMilestones.map(milestone => `
            <div class="task-item milestone-card-click" style="padding: 15px; margin-bottom: 10px; border-left: 4px solid #667eea;" onclick="openMilestoneViewModal(${milestone.id})" role="button" tabindex="0" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();openMilestoneViewModal(${milestone.id});}">
                <div style="display: flex; justify-content: space-between; align-items: start;">
                    <div style="flex: 1;">
                        <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 8px;">
                            <strong style="color: #667eea; font-size: 16px;">${formatDateDisplay(milestone.date)}</strong>
                        </div>
                        <div style="margin-bottom: 8px;">
                            <strong>Description:</strong>
                            <p style="margin: 5px 0; color: #333; white-space: pre-wrap;">${escapeHtml(milestone.description || '')}</p>
                        </div>
                        ${milestone.comment ? `
                        <div>
                            <strong>Comment:</strong>
                            <p style="margin: 5px 0; color: #666; font-style: italic; white-space: pre-wrap;">${escapeHtml(milestone.comment)}</p>
                        </div>
                        ` : ''}
                    </div>
                    <div style="display: flex; gap: 5px;" onclick="event.stopPropagation();">
                        <button type="button" class="btn btn-primary" onclick="editMilestone(${milestone.id})" style="padding: 5px 10px; font-size: 12px;">Edit</button>
                        <button type="button" class="btn btn-danger" onclick="deleteMilestone(${milestone.id})" style="padding: 5px 10px; font-size: 12px;">Delete</button>
                    </div>
                </div>
            </div>
        `).join('')
        : '<p style="text-align: center; color: #999; padding: 20px;">No milestones found. Click "Add Milestone" to create one.</p>';

    document.getElementById('milestonesList').innerHTML = milestonesHtml;
}

function openMilestoneViewModal(milestoneId) {
    const data = getData();
    const milestone = (data.milestones || []).find(m => m.id === milestoneId);
    if (!milestone) return;
    document.getElementById('milestoneViewTitle').textContent = 'Milestone';
    document.getElementById('milestoneViewDate').textContent = formatDateDisplay(milestone.date);
    document.getElementById('milestoneViewDescription').textContent = milestone.description || '';
    const cw = document.getElementById('milestoneViewCommentWrap');
    const cc = document.getElementById('milestoneViewComment');
    if (milestone.comment) {
        cw.style.display = 'block';
        cc.textContent = milestone.comment;
    } else {
        cw.style.display = 'none';
        cc.textContent = '';
    }
    const editBtn = document.getElementById('milestoneViewEditBtn');
    editBtn.onclick = () => {
        closeMilestoneViewModal();
        editMilestone(milestoneId);
    };
    document.getElementById('milestoneViewModal').classList.add('active');
}

function closeMilestoneViewModal() {
    document.getElementById('milestoneViewModal').classList.remove('active');
}

function openMilestoneModal(milestoneId = null) {
    const modal = document.getElementById('milestoneModal');
    const form = document.getElementById('milestoneForm');
    const today = new Date();
    const todayStr = formatDateString(today);

    if (milestoneId) {
        const data = getData();
        const milestone = data.milestones.find(m => m.id === milestoneId);
        if (milestone) {
            document.getElementById('milestoneModalTitle').textContent = 'Edit Milestone';
            document.getElementById('milestoneId').value = milestoneId;
            document.getElementById('milestoneDate').value = milestone.date || todayStr;
            document.getElementById('milestoneDescription').value = milestone.description || '';
            document.getElementById('milestoneComment').value = milestone.comment || '';
        }
    } else {
        document.getElementById('milestoneModalTitle').textContent = 'Add Milestone';
        form.reset();
        document.getElementById('milestoneId').value = '';
        document.getElementById('milestoneDate').value = todayStr;
    }

    modal.classList.add('active');
}

function closeMilestoneModal() {
    document.getElementById('milestoneModal').classList.remove('active');
    document.getElementById('milestoneForm').reset();
}

function saveMilestone(event) {
    if (event) event.preventDefault();

    const milestoneId = document.getElementById('milestoneId').value;
    const date = document.getElementById('milestoneDate').value;
    const description = document.getElementById('milestoneDescription').value.trim();
    const comment = document.getElementById('milestoneComment').value.trim();

    if (!date || !description) {
        alert('Date and Description are required.');
        return;
    }

    updateData(data => {
        if (!data.milestones) data.milestones = [];

        if (milestoneId) {
            const index = data.milestones.findIndex(m => m.id === parseInt(milestoneId));
            if (index !== -1) {
                data.milestones[index] = {
                    ...data.milestones[index],
                    date,
                    description,
                    comment: comment || null,
                    updated_at: new Date().toISOString()
                };
            }
        } else {
            const newMilestone = {
                id: Date.now(),
                date,
                description,
                comment: comment || null,
                created_at: new Date().toISOString(),
                created_by: currentUser.id
            };
            data.milestones.push(newMilestone);
        }
    });

    closeMilestoneModal();
    renderMilestones();
}

function editMilestone(milestoneId) {
    openMilestoneModal(milestoneId);
}

function deleteMilestone(milestoneId) {
    if (!confirm('Are you sure you want to delete this milestone?')) return;

    updateData(data => {
        if (!data.milestones) data.milestones = [];
        data.milestones = data.milestones.filter(m => m.id !== milestoneId);
    });

    renderMilestones();
}

function exportMilestonesCSV() {
    const fromDate = prompt('Enter From Date (DD-MM-YYYY) or leave empty for all:');
    const toDate = prompt('Enter To Date (DD-MM-YYYY) or leave empty for all:');

    const data = getData();
    let milestones = data.milestones || [];

    // Filter by date range if provided
    if (fromDate || toDate) {
        const fromDateObj = fromDate ? parseDateFlexible(fromDate) : null;
        const toDateObj = toDate ? parseDateFlexible(toDate) : null;

        if (fromDate && !fromDateObj) {
            alert('Invalid from date format. Please use DD-MM-YYYY.');
            return;
        }
        if (toDate && !toDateObj) {
            alert('Invalid to date format. Please use DD-MM-YYYY.');
            return;
        }

        milestones = milestones.filter(m => {
            if (!m.date) return false;
            const milestoneDate = parseDateFlexible(m.date);
            if (!milestoneDate) return false;

            if (fromDateObj && milestoneDate < fromDateObj) return false;
            if (toDateObj && milestoneDate > toDateObj) return false;
            return true;
        });
    }

    if (milestones.length === 0) {
        alert('No milestones found in the specified date range.');
        return;
    }

    // Sort by date
    milestones.sort((a, b) => {
        if (!a.date) return 1;
        if (!b.date) return -1;
        return a.date.localeCompare(b.date);
    });

    const headers = ['Date', 'Description', 'Comment'];
    const rows = milestones.map(m => [
        formatDateDisplay(m.date),
        m.description || '',
        m.comment || ''
    ]);

    const csv = [
        headers.map(escapeCSV).join(','),
        ...rows.map(r => r.map(escapeCSV).join(','))
    ].join('\n');

    const filename = `milestones-${fromDate ? formatDateString(parseDateFlexible(fromDate)) : 'all'}-${toDate ? formatDateString(parseDateFlexible(toDate)) : 'all'}-${formatDateString(new Date())}.csv`;
    downloadFile(csv, filename, 'text/csv');
}

// ==================== DAILY PLANNER TAB ====================

function renderDailyPlanner() {
    const plannerDateInput = document.getElementById('plannerDate');
    if (!plannerDateInput) return;

    const selectedDate = plannerDateInput.value || formatDateString(new Date());
    currentPlannerDate = selectedDate;

    const data = getData();
    if (!data.dailyPlanner) data.dailyPlanner = [];

    // Find planner entry for selected date
    const plannerEntry = data.dailyPlanner.find(p => p.date === selectedDate);

    const dateDisplay = formatDateDisplay(selectedDate);

    const plannerHtml = `
        <div class="card" style="margin-bottom: 20px;">
            <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px; border-radius: 8px 8px 0 0;">
                <h3 style="margin: 0;">Daily Planner - ${dateDisplay}</h3>
            </div>
            <div style="padding: 20px;">
                ${plannerEntry ? `
                    <div style="margin-bottom: 15px;">
                        <strong style="color: #667eea;">Part 1:</strong>
                        <p style="margin: 5px 0; padding: 10px; background: #f9f9f9; border-radius: 5px; min-height: 20px; white-space: pre-wrap; word-break: break-word;">${plannerEntry.part1 || '<em style="color: #999;">(not specified)</em>'}</p>
                    </div>
                    <div style="margin-bottom: 15px;">
                        <strong style="color: #667eea;">Part 2:</strong>
                        <p style="margin: 5px 0; padding: 10px; background: #f9f9f9; border-radius: 5px; min-height: 20px; white-space: pre-wrap; word-break: break-word;">${plannerEntry.part2 || '<em style="color: #999;">(not specified)</em>'}</p>
                    </div>
                    <div style="margin-bottom: 15px;">
                        <strong style="color: #667eea;">Part 3:</strong>
                        <p style="margin: 5px 0; padding: 10px; background: #f9f9f9; border-radius: 5px; min-height: 20px; white-space: pre-wrap; word-break: break-word;">${plannerEntry.part3 || '<em style="color: #999;">(not specified)</em>'}</p>
                    </div>
                    <div style="margin-bottom: 15px;">
                        <strong style="color: #667eea;">Part 4:</strong>
                        <p style="margin: 5px 0; padding: 10px; background: #f9f9f9; border-radius: 5px; min-height: 20px; white-space: pre-wrap; word-break: break-word;">${plannerEntry.part4 || '<em style="color: #999;">(not specified)</em>'}</p>
                    </div>
                    <div style="margin-top: 20px;">
                        <button class="btn btn-primary" onclick="editDailyPlanner('${selectedDate}')">Edit Planner</button>
                    </div>
                ` : `
                    <p style="text-align: center; color: #999; padding: 20px;">No planner entry for this date.</p>
                    <div style="text-align: center;">
                        <button class="btn btn-primary" onclick="openPlannerModal('${selectedDate}')">Create Planner Entry</button>
                    </div>
                `}
            </div>
        </div>
    `;

    document.getElementById('dailyPlannerContent').innerHTML = plannerHtml;
}

function changePlannerDate(offset) {
    const plannerDateInput = document.getElementById('plannerDate');
    if (!plannerDateInput) return;

    let currentDate;
    if (offset === 0) {
        currentDate = new Date();
    } else {
        const currentValue = plannerDateInput.value || formatDateString(new Date());
        const dateParts = currentValue.split('-');
        currentDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
        currentDate.setDate(currentDate.getDate() + offset);
    }

    plannerDateInput.value = formatDateString(currentDate);
    renderDailyPlanner();
}

function openPlannerModal(date = null) {
    const modal = document.getElementById('plannerModal');
    const form = document.getElementById('plannerForm');
    const selectedDate = date || document.getElementById('plannerDate')?.value || formatDateString(new Date());

    const data = getData();
    if (!data.dailyPlanner) data.dailyPlanner = [];
    const plannerEntry = data.dailyPlanner.find(p => p.date === selectedDate);

    if (plannerEntry) {
        document.getElementById('plannerModalTitle').textContent = 'Edit Daily Planner';
        document.getElementById('plannerEntryId').value = plannerEntry.id;
        document.getElementById('plannerFormDate').value = plannerEntry.date;
        document.getElementById('plannerPart1').value = plannerEntry.part1 || '';
        document.getElementById('plannerPart2').value = plannerEntry.part2 || '';
        document.getElementById('plannerPart3').value = plannerEntry.part3 || '';
        document.getElementById('plannerPart4').value = plannerEntry.part4 || '';
    } else {
        document.getElementById('plannerModalTitle').textContent = 'Create Daily Planner';
        form.reset();
        document.getElementById('plannerEntryId').value = '';
        document.getElementById('plannerFormDate').value = selectedDate;
        document.getElementById('plannerPart1').value = '';
        document.getElementById('plannerPart2').value = '';
        document.getElementById('plannerPart3').value = '';
        document.getElementById('plannerPart4').value = '';
    }

    modal.classList.add('active');
}

function closePlannerModal() {
    document.getElementById('plannerModal').classList.remove('active');
    document.getElementById('plannerForm').reset();
}

function editDailyPlanner(date) {
    openPlannerModal(date);
}

function saveDailyPlanner(event) {
    if (event) event.preventDefault();

    const entryId = document.getElementById('plannerEntryId').value;
    const date = document.getElementById('plannerFormDate').value;
    const part1 = document.getElementById('plannerPart1').value.trim();
    const part2 = document.getElementById('plannerPart2').value.trim();
    const part3 = document.getElementById('plannerPart3').value.trim();
    const part4 = document.getElementById('plannerPart4').value.trim();

    if (!date) {
        alert('Date is required.');
        return;
    }

    updateData(data => {
        if (!data.dailyPlanner) data.dailyPlanner = [];

        if (entryId) {
            const index = data.dailyPlanner.findIndex(p => p.id === parseInt(entryId));
            if (index !== -1) {
                data.dailyPlanner[index] = {
                    ...data.dailyPlanner[index],
                    date,
                    part1,
                    part2,
                    part3,
                    part4,
                    updated_at: new Date().toISOString()
                };
            }
        } else {
            // Check if entry already exists for this date
            const existingIndex = data.dailyPlanner.findIndex(p => p.date === date);
            if (existingIndex !== -1) {
                data.dailyPlanner[existingIndex] = {
                    ...data.dailyPlanner[existingIndex],
                    part1,
                    part2,
                    part3,
                    part4,
                    updated_at: new Date().toISOString()
                };
            } else {
                const newEntry = {
                    id: Date.now(),
                    date,
                    part1,
                    part2,
                    part3,
                    part4,
                    created_at: new Date().toISOString(),
                    created_by: currentUser.id
                };
                data.dailyPlanner.push(newEntry);
            }
        }
    });

    closePlannerModal();
    renderDailyPlanner();
}

function exportPlannerCSV() {
    const fromDate = prompt('Enter From Date (DD-MM-YYYY) or leave empty for all:');
    const toDate = prompt('Enter To Date (DD-MM-YYYY) or leave empty for all:');

    const data = getData();
    let plannerEntries = data.dailyPlanner || [];

    // Filter by date range if provided
    if (fromDate || toDate) {
        const fromDateObj = fromDate ? parseDateFlexible(fromDate) : null;
        const toDateObj = toDate ? parseDateFlexible(toDate) : null;

        if (fromDate && !fromDateObj) {
            alert('Invalid from date format. Please use DD-MM-YYYY.');
            return;
        }
        if (toDate && !toDateObj) {
            alert('Invalid to date format. Please use DD-MM-YYYY.');
            return;
        }

        plannerEntries = plannerEntries.filter(p => {
            if (!p.date) return false;
            const entryDate = parseDateFlexible(p.date);
            if (!entryDate) return false;

            if (fromDateObj && entryDate < fromDateObj) return false;
            if (toDateObj && entryDate > toDateObj) return false;
            return true;
        });
    }

    if (plannerEntries.length === 0) {
        alert('No planner entries found in the specified date range.');
        return;
    }

    // Sort by date
    plannerEntries.sort((a, b) => {
        if (!a.date) return 1;
        if (!b.date) return -1;
        return a.date.localeCompare(b.date);
    });

    const headers = ['Date', 'Part 1', 'Part 2', 'Part 3', 'Part 4'];
    const rows = plannerEntries.map(p => [
        formatDateDisplay(p.date),
        p.part1 || '',
        p.part2 || '',
        p.part3 || '',
        p.part4 || ''
    ]);

    const csv = [
        headers.map(escapeCSV).join(','),
        ...rows.map(r => r.map(escapeCSV).join(','))
    ].join('\n');

    const filename = `daily-planner-${fromDate ? formatDateString(parseDateFlexible(fromDate)) : 'all'}-${toDate ? formatDateString(parseDateFlexible(toDate)) : 'all'}-${formatDateString(new Date())}.csv`;
    downloadFile(csv, filename, 'text/csv');
}

// ============================================
// LOCATIONS TAB FUNCTIONS
// ============================================

let currentLocationCategory = 'folder_path';
let locationAttachmentsData = [];

// Render Locations
function renderLocations() {
    const data = getData();
    const grid = document.getElementById('locationsGrid');
    const searchQuery = document.getElementById('locationSearch')?.value.toLowerCase() || '';

    if (!data.locationItems) {
        data.locationItems = [];
        saveData(data);
    }

    // Filter by category and search
    let filtered = data.locationItems.filter(item => item.category === currentLocationCategory);

    if (searchQuery) {
        filtered = filtered.filter(item =>
            item.heading.toLowerCase().includes(searchQuery) ||
            item.path.toLowerCase().includes(searchQuery)
        );
    }

    if (filtered.length === 0) {
        grid.innerHTML = `
            <div class="location-empty">
                <div class="location-empty-icon">${getCategoryIcon(currentLocationCategory)}</div>
                <h3>No ${getCategoryLabel(currentLocationCategory)} Yet</h3>
                <p>Click "Add Location" to create your first entry</p>
            </div>
        `;
        return;
    }

    grid.innerHTML = filtered.map(item => `
        <div class="location-card" data-category="${item.category}">
            <div class="location-heading">
                <span class="location-heading-icon">${getCategoryIcon(item.category)}</span>
                <span class="location-heading-text">${escapeHtml(item.heading)}</span>
            </div>
            
            <div class="location-path">
                <a href="${formatPathForLink(item.path)}" class="location-path-link" 
                   data-path="${escapeHtml(item.path)}" data-category="${item.category}"
                   onclick="return handlePathClick(event, this.dataset.path, this.dataset.category)">
                    ${escapeHtml(item.path)}
                </a>
                <div class="location-path-actions">
                    <button class="location-path-btn" data-path="${escapeHtml(item.path)}"
                        onclick="copyToClipboard(this.dataset.path)">
                        📋 Copy
                    </button>
                    <button class="location-path-btn" data-path="${escapeHtml(item.path)}" data-category="${item.category}"
                        onclick="openPathInExplorer(this.dataset.path, this.dataset.category)">
                        🔗 Open
                    </button>
                </div>
            </div>
            
            <div class="location-attachments">
                ${item.attachments && item.attachments.length > 0 ?
            item.attachments.map(att => `
                        <div class="location-attachment">
                            <span class="location-attachment-icon">📎</span>
                            <span class="location-attachment-name" title="${escapeHtml(att.name)}">${escapeHtml(att.name)}</span>
                            <span class="location-attachment-size">(${formatFileSize(att.size)})</span>
                            <button class="location-attachment-download" onclick="downloadLocationAttachment(${item.id}, ${att.id})">
                                ⬇
                            </button>
                        </div>
                    `).join('')
            : '<span style="color: #999; font-size: 12px;">No attachments</span>'
        }
            </div>
            
            <div class="location-actions">
                <button class="location-edit-btn" onclick="editLocationItem(${item.id})">
                    ✏️ Edit
                </button>
                <button class="location-delete-btn" onclick="deleteLocationItem(${item.id})">
                    🗑️ Delete
                </button>
            </div>
        </div>
    `).join('');
}

// Filter by category
function filterLocationByCategory(category) {
    currentLocationCategory = category;

    // Update active tab
    document.querySelectorAll('.location-category-tab').forEach(tab => {
        tab.classList.remove('active');
        if (tab.getAttribute('data-category') === category) {
            tab.classList.add('active');
        }
    });

    renderLocations();
}

// Open Location Modal
function openLocationModal() {
    document.getElementById('locationModal').classList.add('active');
    document.getElementById('locationModalTitle').textContent = 'Add Location';
    document.getElementById('locationForm').reset();
    document.getElementById('locationId').value = '';
    document.getElementById('locationCategory').value = currentLocationCategory;
    document.getElementById('attachmentPreview').innerHTML = '';
    locationAttachmentsData = [];
    const fin = document.getElementById('locationAttachments');
    if (fin) fin.value = '';
}

// Close Location Modal
function closeLocationModal() {
    document.getElementById('locationModal').classList.remove('active');
    locationAttachmentsData = [];
    const fin = document.getElementById('locationAttachments');
    if (fin) fin.value = '';
}

function readLocationFilesAsAttachmentEntries(files) {
    const maxSize = MAX_ATTACHMENT_SIZE;
    return Promise.all(Array.from(files).map((file, index) => new Promise((resolve, reject) => {
        if (file.size > maxSize) {
            alert(`File "${file.name}" is too large (max ${MAX_ATTACHMENT_SIZE / 1024 / 1024}MB)`);
            resolve(null);
            return;
        }
        const reader = new FileReader();
        reader.onload = () => resolve({
            id: Date.now() + index + Math.random(),
            name: file.name,
            data: reader.result,
            type: file.type,
            size: file.size
        });
        reader.onerror = () => reject(reader.error || new Error('File read failed'));
        reader.readAsDataURL(file);
    }))).then(items => items.filter(Boolean));
}

function renderLocationAttachmentPreviewFromData() {
    const preview = document.getElementById('attachmentPreview');
    if (!preview) return;
    const storeHint = isApiMode()
        ? 'stored on server (cloud)'
        : (`stored in IndexedDB, max ${MAX_ATTACHMENT_SIZE / 1024 / 1024}MB each`);
    if (!locationAttachmentsData.length) {
        preview.innerHTML = '';
        return;
    }
    preview.innerHTML = `<div style="font-size: 12px; color: #666; margin-bottom: 8px;">Attachments (${storeHint}):</div>` +
        locationAttachmentsData.map(att => `
            <div style="display: flex; align-items: center; gap: 8px; padding: 6px; background: #f8f9fa; border-radius: 4px; margin-bottom: 4px;">
                <span>📎</span>
                <span style="flex: 1; font-size: 12px;">${escapeHtml(att.name)}</span>
                <span style="font-size: 11px; color: #999;">(${formatFileSize(att.size)})</span>
            </div>
        `).join('');
}

/** Merge new files from input with existing saved attachments (remote / IndexedDB). */
async function handleLocationAttachments() {
    const fileInput = document.getElementById('locationAttachments');
    const files = fileInput && fileInput.files;
    const preserved = locationAttachmentsData.filter(a => a.remoteAttachment || a.storedInIndexedDB);

    if (!files || files.length === 0) {
        locationAttachmentsData = preserved;
        renderLocationAttachmentPreviewFromData();
        return;
    }

    if (files.length > MAX_ATTACHMENT_FILES) {
        alert(`Maximum ${MAX_ATTACHMENT_FILES} files per location. You selected ${files.length}.`);
        return;
    }

    try {
        const newEntries = await readLocationFilesAsAttachmentEntries(files);
        locationAttachmentsData = preserved.concat(newEntries);
        fileInput.value = '';
        renderLocationAttachmentPreviewFromData();
    } catch (e) {
        console.error(e);
        alert('Could not read selected files. Please try again.');
    }
}

async function syncLocationAttachmentsBeforeSave() {
    const fileInput = document.getElementById('locationAttachments');
    if (fileInput && fileInput.files && fileInput.files.length > 0) {
        await handleLocationAttachments();
    }
}

// Save Location (IndexedDB locally, or GridFS when API mode)
async function saveLocation(event) {
    event.preventDefault();
    await syncLocationAttachmentsBeforeSave();

    const data = getData();
    const locId = document.getElementById('locationId').value;
    const locationId = locId ? parseInt(locId) : Date.now();
    const category = document.getElementById('locationCategory').value;
    const heading = document.getElementById('locationHeading').value;
    const path = document.getElementById('locationPath').value;

    if (!data.locationItems) {
        data.locationItems = [];
    }

    const syncAtts = [];
    const asyncJobs = [];

    locationAttachmentsData.forEach(att => {
        if (att.remoteAttachment) {
            syncAtts.push({
                id: att.id,
                name: att.name,
                type: att.type,
                size: att.size,
                remoteAttachment: true
            });
        } else if (att.storedInIndexedDB) {
            syncAtts.push({
                id: att.id,
                name: att.name,
                type: att.type,
                size: att.size,
                storedInIndexedDB: true
            });
        } else if (att.data) {
            if (isApiMode()) {
                asyncJobs.push((async () => {
                    const blob = dataUrlToBlob(att.data);
                    if (!blob) throw new Error('Invalid file data');
                    await uploadRemoteAttachment(locationId, att.id, blob, att.name, att.type || blob.type);
                    return {
                        id: att.id,
                        name: att.name,
                        type: att.type,
                        size: att.size,
                        remoteAttachment: true
                    };
                })());
            } else {
                asyncJobs.push(
                    putAttachmentBlob(locationId, att.id, att.data).then(() => ({
                        id: att.id,
                        name: att.name,
                        type: att.type,
                        size: att.size,
                        storedInIndexedDB: true
                    }))
                );
            }
        } else {
            syncAtts.push(att);
        }
    });

    const doSave = (attachmentsToSave) => {
        if (locId) {
            const index = data.locationItems.findIndex(item => item.id === parseInt(locId));
            if (index !== -1) {
                const oldItem = data.locationItems[index];
                const oldIds = (oldItem.attachments || []).map(a => a.id);
                const newIds = attachmentsToSave.map(a => a.id);
                oldIds.forEach(aid => {
                    if (!newIds.includes(aid)) {
                        const oldA = (oldItem.attachments || []).find(a => a.id === aid);
                        if (oldA && oldA.remoteAttachment && isApiMode()) {
                            apiFetch(`/api/attachments/${oldItem.id}/${aid}`, { method: 'DELETE' }).catch(() => {});
                        }
                        if (oldA && oldA.storedInIndexedDB) {
                            removeAttachmentBlob(oldItem.id, aid).catch(() => {});
                        }
                    }
                });
                data.locationItems[index] = {
                    ...oldItem,
                    category,
                    heading,
                    path,
                    attachments: attachmentsToSave.length > 0 ? attachmentsToSave : [],
                    updatedAt: Date.now()
                };
            }
        } else {
            data.locationItems.push({
                id: locationId,
                category,
                heading,
                path,
                attachments: attachmentsToSave,
                createdAt: Date.now(),
                updatedAt: Date.now()
            });
        }
        saveData(data);
        closeLocationModal();
        renderLocations();
    };

    if (asyncJobs.length > 0) {
        try {
            const results = await Promise.all(asyncJobs);
            doSave(syncAtts.concat(results));
        } catch (err) {
            console.error('Failed to store attachments', err);
            alert('One or more attachments could not be saved. Please try again.');
        }
    } else {
        doSave(syncAtts);
    }
}

// Edit Location Item
function editLocationItem(id) {
    const data = getData();
    const item = data.locationItems.find(i => i.id === id);

    if (!item) return;

    document.getElementById('locationModal').classList.add('active');
    document.getElementById('locationModalTitle').textContent = 'Edit Location';
    document.getElementById('locationId').value = item.id;
    document.getElementById('locationCategory').value = item.category;
    document.getElementById('locationHeading').value = item.heading;
    document.getElementById('locationPath').value = item.path;

    // Show existing attachments
    locationAttachmentsData = item.attachments || [];
    const fin = document.getElementById('locationAttachments');
    if (fin) fin.value = '';
    renderLocationAttachmentPreviewFromData();
}

// Delete Location Item
function deleteLocationItem(id) {
    if (!confirm('Are you sure you want to delete this location?')) return;

    const data = getData();
    const item = data.locationItems.find(i => i.id === id);
    if (item && item.attachments && isApiMode()) {
        item.attachments.forEach(a => {
            if (a.remoteAttachment) {
                apiFetch(`/api/attachments/${id}/${a.id}`, { method: 'DELETE' }).catch(() => {});
            }
            if (a.storedInIndexedDB) {
                removeAttachmentBlob(id, a.id).catch(() => {});
            }
        });
    } else if (item && item.attachments) {
        item.attachments.forEach(a => {
            if (a.storedInIndexedDB) {
                removeAttachmentBlob(id, a.id).catch(() => {});
            }
        });
    }
    data.locationItems = data.locationItems.filter(item => item.id !== id);
    saveData(data);
    renderLocations();
}

// Download attachment (from inline data, IndexedDB, or API / GridFS)
function downloadLocationAttachment(locationId, attachmentId) {
    const data = getData();
    const location = data.locationItems.find(l => l.id === locationId);
    if (!location) return;

    const attachment = location.attachments.find(a => a.id === attachmentId);
    if (!attachment) return;

    if (attachment.remoteAttachment && isApiMode()) {
        apiFetch(`/api/attachments/${locationId}/${attachmentId}`)
            .then(res => {
                if (!res.ok) throw new Error('Download failed');
                return res.blob();
            })
            .then(blob => {
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = attachment.name;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
            })
            .catch(err => {
                console.error('Download from server failed', err);
                alert('Failed to download attachment.');
            });
    } else if (attachment.storedInIndexedDB) {
        getAttachmentBlob(locationId, attachmentId).then(dataUrl => {
            if (!dataUrl) {
                alert('Attachment data not found. It may have been cleared.');
                return;
            }
            const link = document.createElement('a');
            link.href = dataUrl;
            link.download = attachment.name;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }).catch(err => {
            console.error('Download from IndexedDB failed', err);
            alert('Failed to download attachment.');
        });
    } else {
        const link = document.createElement('a');
        link.href = attachment.data;
        link.download = attachment.name;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
}

// ========== Code Snippets ==========
const SNIPPET_LANGUAGES = { python: 'Python', sql: 'SQL', java: 'Java', vba: 'Excel VBA', javascript: 'JavaScript', typescript: 'TypeScript', html: 'HTML', css: 'CSS', bash: 'Bash / Shell', powershell: 'PowerShell', other: 'Other' };

function renderCodeSnippets() {
    const data = getData();
    const snippets = Array.isArray(data.codeSnippets) ? data.codeSnippets : [];
    const langFilter = (document.getElementById('snippetLanguageFilter') || {}).value;
    const search = ((document.getElementById('snippetSearch') || {}).value || '').toLowerCase().trim();

    let filtered = snippets.filter(s => {
        if (langFilter && s.language !== langFilter) return false;
        if (search && !(s.title || '').toLowerCase().includes(search) && !(s.code || '').toLowerCase().includes(search)) return false;
        return true;
    });

    const listEl = document.getElementById('snippetsList');
    const viewPanel = document.getElementById('snippetViewPanel');
    if (!listEl) return;
    listEl.style.display = '';

    listEl.innerHTML = filtered.length === 0
        ? '<p style="text-align: center; color: #999; padding: 40px;">No code snippets yet. Click "+ Add Snippet" to save your first snippet.</p>'
        : filtered.map(s => {
            const preview = (s.code || '').split('\n').slice(0, 4).join('\n') || '(no code)';
            const langLabel = SNIPPET_LANGUAGES[s.language] || s.language || 'Other';
            return `
                <div class="snippet-card">
                    <div class="snippet-card-title">${escapeHtml(s.title || 'Untitled')}</div>
                    <span class="snippet-lang-badge">${escapeHtml(langLabel)}</span>
                    <div class="snippet-code-preview">${escapeHtml(preview)}</div>
                    <div class="snippet-card-actions">
                        <button class="btn btn-primary" onclick="openSnippetView(${s.id})">View</button>
                        <button class="btn btn-secondary" onclick="copySnippet(${s.id}, event)" title="Copy full snippet">Copy</button>
                        <button class="btn btn-secondary" onclick="openSnippetModal(${s.id})">Edit</button>
                        <button class="btn btn-danger" onclick="deleteSnippet(${s.id})">Delete</button>
                    </div>
                </div>`;
        }).join('');

    if (viewPanel) viewPanel.style.display = 'none';
}

function openSnippetModal(id) {
    document.getElementById('snippetModal').classList.add('active');
    document.getElementById('snippetModalTitle').textContent = id ? 'Edit Code Snippet' : 'Add Code Snippet';
    document.getElementById('snippetId').value = id || '';
    document.getElementById('snippetTitle').value = '';
    document.getElementById('snippetLanguage').value = 'python';
    document.getElementById('snippetCode').value = '';

    if (id) {
        const data = getData();
        const s = (data.codeSnippets || []).find(x => x.id === id);
        if (s) {
            document.getElementById('snippetTitle').value = s.title || '';
            document.getElementById('snippetLanguage').value = s.language || 'python';
            document.getElementById('snippetCode').value = s.code || '';
        }
    }
}

function closeSnippetModal() {
    document.getElementById('snippetModal').classList.remove('active');
}

function saveSnippet(event) {
    event.preventDefault();
    const data = getData();
    if (!data.codeSnippets) data.codeSnippets = [];
    const id = document.getElementById('snippetId').value;
    const title = document.getElementById('snippetTitle').value.trim();
    const language = document.getElementById('snippetLanguage').value;
    const code = document.getElementById('snippetCode').value;

    if (id) {
        const idx = data.codeSnippets.findIndex(s => s.id === parseInt(id, 10));
        if (idx !== -1) {
            data.codeSnippets[idx] = { ...data.codeSnippets[idx], title, language, code, updatedAt: Date.now() };
        }
    } else {
        data.codeSnippets.push({
            id: Date.now(),
            title,
            language,
            code,
            createdAt: Date.now(),
            updatedAt: Date.now()
        });
    }
    saveData(data);
    closeSnippetModal();
    renderCodeSnippets();
    const viewPanel = document.getElementById('snippetViewPanel');
    if (viewPanel && viewPanel.style.display !== 'none') {
        const viewId = document.getElementById('snippetViewId').value;
        if (viewId) openSnippetView(parseInt(viewId, 10));
    }
}

function deleteSnippet(id) {
    if (!confirm('Delete this code snippet?')) return;
    updateData(data => {
        data.codeSnippets = (data.codeSnippets || []).filter(s => s.id !== id);
    });
    renderCodeSnippets();
    const viewPanel = document.getElementById('snippetViewPanel');
    if (viewPanel) viewPanel.style.display = 'none';
}

function copySnippet(id, ev) {
    const data = getData();
    const s = (data.codeSnippets || []).find(x => x.id === id);
    if (!s) return;
    const code = s.code || '';
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(code).then(() => {
            const btn = ev && ev.target;
            if (btn && btn.classList) {
                const orig = btn.textContent;
                btn.textContent = 'Copied!';
                setTimeout(() => { btn.textContent = orig; }, 1500);
            } else {
                alert('Snippet copied to clipboard.');
            }
        }).catch(() => { fallbackCopyToClipboard(code); });
    } else {
        fallbackCopyToClipboard(code);
    }
}
function fallbackCopyToClipboard(text) {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed'; ta.style.left = '-9999px';
    document.body.appendChild(ta);
    ta.select();
    try {
        document.execCommand('copy');
        alert('Snippet copied to clipboard.');
    } catch (e) {
        alert('Could not copy. Please select and copy manually.');
    }
    document.body.removeChild(ta);
}

function openSnippetView(id) {
    const data = getData();
    const s = (data.codeSnippets || []).find(x => x.id === id);
    if (!s) return;
    const viewPanel = document.getElementById('snippetViewPanel');
    const listEl = document.getElementById('snippetsList');
    if (viewPanel && listEl) {
        listEl.style.display = 'none';
        viewPanel.style.display = 'block';
    }
    document.getElementById('snippetViewTitle').textContent = s.title || 'Untitled';
    document.getElementById('snippetViewLanguage').textContent = SNIPPET_LANGUAGES[s.language] || s.language || 'Other';
    document.getElementById('snippetViewCode').textContent = s.code || '';
    document.getElementById('snippetViewCode').style.display = 'block';
    const editEl = document.getElementById('snippetViewCodeEdit');
    if (editEl) {
        editEl.value = s.code || '';
        editEl.style.display = 'none';
    }
    document.getElementById('snippetViewId').value = id;
    document.getElementById('snippetViewEditBtn').textContent = 'Edit';
}

function closeSnippetView() {
    const viewPanel = document.getElementById('snippetViewPanel');
    const listEl = document.getElementById('snippetsList');
    if (viewPanel) viewPanel.style.display = 'none';
    if (listEl) listEl.style.display = '';
}

function snippetViewToggleEdit() {
    const pre = document.getElementById('snippetViewCode');
    const edit = document.getElementById('snippetViewCodeEdit');
    const btn = document.getElementById('snippetViewEditBtn');
    if (!pre || !edit || !btn) return;
    if (edit.style.display === 'none') {
        edit.value = pre.textContent || '';
        pre.style.display = 'none';
        edit.style.display = 'block';
        edit.focus();
        btn.textContent = 'Save';
    } else {
        const newCode = edit.value;
        const id = parseInt(document.getElementById('snippetViewId').value, 10);
        updateData(data => {
            const s = (data.codeSnippets || []).find(x => x.id === id);
            if (s) { s.code = newCode; s.updatedAt = Date.now(); }
        });
        pre.textContent = newCode;
        edit.style.display = 'none';
        pre.style.display = 'block';
        btn.textContent = 'Edit';
        renderCodeSnippets();
    }
}

// ========== Journal Tab ==========
function getJournalDateStr() {
    const el = document.getElementById('journalDate');
    return el ? el.value : formatDateString(new Date());
}

function setJournalDateStr(dateStr) {
    const el = document.getElementById('journalDate');
    if (el) el.value = dateStr;
}

function renderJournal() {
    if (!window.journalAutoSaveInit) {
        initJournalAutoSave();
        window.journalAutoSaveInit = true;
    }
    const todayStr = formatDateString(new Date());
    const dateEl = document.getElementById('journalDate');
    if (dateEl && !dateEl.value) {
        dateEl.value = todayStr;
    }
    const viewMode = document.getElementById('journalViewMode');
    if (viewMode && viewMode.value === 'full') {
        journalShowFullView();
    } else {
        journalShowDayView();
        journalLoadDate();
    }
}

function journalShowDayView() {
    document.getElementById('journalFullView').style.display = 'none';
    document.getElementById('journalDayView').style.display = 'block';
}

function journalShowFullView() {
    document.getElementById('journalDayView').style.display = 'none';
    const fullEl = document.getElementById('journalFullView');
    fullEl.style.display = 'block';
    const data = getData();
    const journal = data.journal && typeof data.journal === 'object' ? data.journal : {};
    const dates = Object.keys(journal).filter(d => journal[d]).sort();
    if (dates.length === 0) {
        fullEl.innerHTML = '<p style="color: #999; padding: 20px;">No journal entries yet.</p>';
        return;
    }
    fullEl.innerHTML = dates.map(dateStr => {
        const displayDate = formatDateDisplay(dateStr);
        const content = (journal[dateStr] || '').trim();
        return `<div class="journal-day-block" style="margin-bottom: 24px; padding-bottom: 20px; border-bottom: 1px solid #e0e0e0;">
            <div style="font-weight: 600; color: #667eea; margin-bottom: 10px; font-size: 16px;">${displayDate}</div>
            <div style="color: #333; line-height: 1.6;">${content || '<span style="color: #999;">(empty)</span>'}</div>
        </div>`;
    }).join('');
}

function journalSwitchView() {
    const viewMode = document.getElementById('journalViewMode');
    if (viewMode && viewMode.value === 'full') {
        journalShowFullView();
    } else {
        journalShowDayView();
        journalLoadDate();
    }
}

function journalLoadDate() {
    const dateStr = getJournalDateStr();
    const data = getData();
    const journal = data.journal && typeof data.journal === 'object' ? data.journal : {};
    const editor = document.getElementById('journalEditor');
    if (editor) {
        editor.innerHTML = journal[dateStr] || '';
    }
}

function journalPrevDay() {
    const dateStr = getJournalDateStr();
    const d = dateStr ? new Date(dateStr + 'T12:00:00') : new Date();
    d.setDate(d.getDate() - 1);
    setJournalDateStr(formatDateString(d));
    journalLoadDate();
}

function journalNextDay() {
    const dateStr = getJournalDateStr();
    const d = dateStr ? new Date(dateStr + 'T12:00:00') : new Date();
    d.setDate(d.getDate() + 1);
    setJournalDateStr(formatDateString(d));
    journalLoadDate();
}

function journalSaveCurrent() {
    const dateStr = getJournalDateStr();
    const editor = document.getElementById('journalEditor');
    if (!editor) return;
    const content = editor.innerHTML.trim();
    updateData(data => {
        if (!data.journal || typeof data.journal !== 'object') data.journal = {};
        data.journal[dateStr] = content;
    });
}

function journalDoSearch() {
    const q = (document.getElementById('journalSearch') || {}).value;
    if (!q || !q.trim()) {
        journalSwitchView();
        return;
    }
    const data = getData();
    const journal = data.journal && typeof data.journal === 'object' ? data.journal : {};
    const searchLower = q.toLowerCase().trim();
    const dates = Object.keys(journal).filter(d => {
        const html = journal[d] || '';
        const div = document.createElement('div');
        div.innerHTML = html;
        const text = (div.textContent || div.innerText || '').toLowerCase();
        return text.includes(searchLower);
    }).sort();
    const fullEl = document.getElementById('journalFullView');
    const dayView = document.getElementById('journalDayView');
    fullEl.style.display = 'block';
    dayView.style.display = 'none';
    document.getElementById('journalViewMode').value = 'full';
    if (dates.length === 0) {
        fullEl.innerHTML = '<p style="color: #999; padding: 20px;">No entries match your search.</p>';
        return;
    }
    fullEl.innerHTML = dates.map(dateStr => {
        const displayDate = formatDateDisplay(dateStr);
        const content = (journal[dateStr] || '').trim();
        const div = document.createElement('div');
        div.innerHTML = content;
        let text = div.textContent || div.innerText || '';
        const idx = text.toLowerCase().indexOf(searchLower);
        let snippet = content;
        if (idx >= 0 && content) {
            const span = document.createElement('span');
            span.innerHTML = content;
            text = span.textContent || span.innerText || '';
            const start = Math.max(0, idx - 40);
            const end = Math.min(text.length, idx + searchLower.length + 60);
            snippet = '...' + (text.substring(start, end) || content) + '...';
        }
        return `<div class="journal-day-block" style="margin-bottom: 16px; padding: 12px; background: #f9f9f9; border-radius: 8px;">
            <div style="font-weight: 600; color: #667eea; margin-bottom: 8px;">${displayDate}</div>
            <div style="color: #333; line-height: 1.6;">${snippet || '(empty)'}</div>
        </div>`;
    }).join('');
}

function journalExportToWord() {
    const fromStr = prompt('Export from date (YYYY-MM-DD):', formatDateString(new Date()));
    if (fromStr === null) return;
    const toStr = prompt('Export to date (YYYY-MM-DD):', formatDateString(new Date()));
    if (toStr === null) return;
    const fromDate = fromStr ? new Date(fromStr + 'T12:00:00') : null;
    const toDate = toStr ? new Date(toStr + 'T12:00:00') : null;
    if (!fromDate || isNaN(fromDate.getTime()) || !toDate || isNaN(toDate.getTime())) {
        alert('Please enter valid dates in YYYY-MM-DD format.');
        return;
    }
    const data = getData();
    const journal = data.journal && typeof data.journal === 'object' ? data.journal : {};
    const dates = Object.keys(journal).filter(d => {
        const dt = new Date(d + 'T12:00:00');
        return !isNaN(dt.getTime()) && dt >= fromDate && dt <= toDate && (journal[d] || '').trim();
    }).sort();
    if (dates.length === 0) {
        alert('No journal entries in the selected date range.');
        return;
    }
    const body = dates.map(dateStr => {
        const displayDate = formatDateDisplay(dateStr);
        const content = (journal[dateStr] || '').trim();
        return `<div style="margin-bottom: 24px;"><p style="font-weight: bold; color: #333; margin-bottom: 8px;">${displayDate}</p><div>${content}</div></div>`;
    }).join('');
    const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Journal ${fromStr} to ${toStr}</title></head><body style="font-family: Arial, sans-serif; padding: 20px;">${body}</body></html>`;
    const blob = new Blob(['\ufeff' + html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `journal-${fromStr}-to-${toStr}.doc`;
    a.click();
    URL.revokeObjectURL(url);
}

function initJournalAutoSave() {
    const editor = document.getElementById('journalEditor');
    if (!editor) return;
    editor.addEventListener('input', () => debouncedJournalSave());
    editor.addEventListener('paste', () => setTimeout(() => debouncedJournalSave(), 100));
}

// Utility: Format path for link
function formatPathForLink(path) {
    // If already has file:// protocol, return as-is
    if (path.startsWith('file://')) {
        return path;
    }

    // If it's a Windows path (contains : or \)
    if (path.includes(':') || path.includes('\\')) {
        // Convert backslashes to forward slashes
        const normalizedPath = path.replace(/\\/g, '/');
        return `file:///${normalizedPath}`;
    }

    // Otherwise assume it's a Unix path
    return `file://${path}`;
}

// Handle path click (path link in location card) — open folder/file directly (browser may block file:// from HTTPS)
function handlePathClick(event, path, _category) {
    event.preventDefault();
    tryOpenPath(path);
    return false;
}

// Open path in explorer / file manager
function openPathInExplorer(path, _category) {
    tryOpenPath(path);
}

function tryOpenPath(path) {
    const url = formatPathForLink(path);
    const opened = window.open(url, '_blank', 'noopener,noreferrer');
    if (!opened) {
        const link = document.createElement('a');
        link.href = url;
        link.target = '_blank';
        link.rel = 'noopener';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
}

// Copy to clipboard
function copyToClipboard(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(() => {
            // Success - display temporary notification
            showCopyNotification();
        }).catch(err => {
            // Fallback for older browsers
            fallbackCopyToClipboard(text);
        });
    } else {
        fallbackCopyToClipboard(text);
    }
}

// Fallback copy method
function fallbackCopyToClipboard(text) {
    const textArea = document.createElement('textarea');
    textArea.value = text;
    textArea.style.position = 'fixed';
    textArea.style.top = '-9999px';
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();

    try {
        document.execCommand('copy');
        showCopyNotification();
    } catch (err) {
        console.error('Failed to copy', err);
    }

    document.body.removeChild(textArea);
}

// Show copy notification
function showCopyNotification() {
    const notification = document.createElement('div');
    notification.textContent = '✓ Copied to clipboard!';
    notification.style.cssText = `
        position: fixed;
        bottom: 30px;
        right: 30px;
        background: #28a745;
        color: white;
        padding: 12px 20px;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        z-index: 10000;
        font-size: 14px;
        font-weight: 500;
        animation: slideIn 0.3s ease;
    `;

    document.body.appendChild(notification);

    setTimeout(() => {
        notification.style.animation = 'slideOut 0.3s ease';
        setTimeout(() => {
            document.body.removeChild(notification);
        }, 300);
    }, 2000);
}

// Add CSS animation for notification (add to head if not exists)
if (!document.getElementById('copyNotificationStyles')) {
    const style = document.createElement('style');
    style.id = 'copyNotificationStyles';
    style.textContent = `
        @keyframes slideIn {
            from { transform: translateX(400px); opacity: 0; }
            to { transform: translateX(0); opacity: 1; }
        }
        @keyframes slideOut {
            from { transform: translateX(0); opacity: 1; }
            to { transform: translateX(400px); opacity: 0; }
        }
    `;
    document.head.appendChild(style);
}

// Format file size
function formatFileSize(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

// Get category icon
function getCategoryIcon(category) {
    const icons = {
        'folder_path': '📁',
        'file_path': '📄',
        'learning': '📚',
        'other_document': '📋'
    };
    return icons[category] || '📂';
}

// Get category label
function getCategoryLabel(category) {
    const labels = {
        'folder_path': 'Folder Paths',
        'file_path': 'File Paths',
        'learning': 'Learning Resources',
        'other_document': 'Other Documents'
    };
    return labels[category] || 'Items';
}

// Learning Notes state
let learningSelectedCourse = '';
let learningSelectedNoteId = null;
let learningRevisionMode = false;
let learningFlashcards = [];
let learningFlashcardIndex = 0;

function getLearningPriorityWeight(priority) {
    if (priority === 'high') return 3;
    if (priority === 'medium') return 2;
    return 1;
}

function getLearningRevisionLabel(status) {
    if (status === 'revised_once') return 'Revised Once';
    if (status === 'mastered') return 'Mastered';
    return 'Not Revised';
}

function openLearningNoteModal(noteId = null) {
    const data = getData();
    const modal = document.getElementById('learningNoteModal');
    if (!modal) return;
    const form = document.getElementById('learningNoteForm');
    if (!form) return;

    if (noteId) {
        const note = (data.learningNotes || []).find(n => n.id == noteId);
        if (!note) return;
        document.getElementById('learningNoteModalTitle').textContent = 'Edit Learning Note';
        document.getElementById('learningNoteId').value = note.id;
        document.getElementById('learningCourseName').value = note.course_name || '';
        document.getElementById('learningTopicName').value = note.topic_name || '';
        document.getElementById('learningEntryDate').value = note.entry_date || formatDateString(new Date());
        document.getElementById('learningKeyConcepts').value = (note.key_concepts || []).join('\n');
        document.getElementById('learningDetailedNotes').value = note.detailed_notes || '';
        document.getElementById('learningTags').value = (note.tags || []).join(', ');
        document.getElementById('learningRevisionStatus').value = note.revision_status || 'not_revised';
        document.getElementById('learningLastRevisedDate').value = note.last_revised_date || '';
        document.getElementById('learningPriority').value = note.priority || 'medium';
        document.getElementById('learningIsPinned').checked = !!note.pinned;
    } else {
        form.reset();
        document.getElementById('learningNoteModalTitle').textContent = 'New Learning Note';
        document.getElementById('learningNoteId').value = '';
        document.getElementById('learningEntryDate').value = formatDateString(new Date());
        document.getElementById('learningRevisionStatus').value = 'not_revised';
        document.getElementById('learningPriority').value = 'medium';
    }
    modal.classList.add('active');
}

function closeLearningNoteModal() {
    const modal = document.getElementById('learningNoteModal');
    if (modal) modal.classList.remove('active');
}

function saveLearningNote(event) {
    if (event) event.preventDefault();
    const id = document.getElementById('learningNoteId').value;
    const courseName = document.getElementById('learningCourseName').value.trim();
    const topicName = document.getElementById('learningTopicName').value.trim();
    const entryDate = document.getElementById('learningEntryDate').value;
    const keyConcepts = document.getElementById('learningKeyConcepts').value
        .split('\n').map(s => s.trim()).filter(Boolean);
    const detailedNotes = document.getElementById('learningDetailedNotes').value.trim();
    const tags = document.getElementById('learningTags').value
        .split(',').map(s => s.trim()).filter(Boolean);
    const revisionStatus = document.getElementById('learningRevisionStatus').value;
    const lastRevisedDate = document.getElementById('learningLastRevisedDate').value || '';
    const priority = document.getElementById('learningPriority').value;
    const pinned = document.getElementById('learningIsPinned').checked;

    if (!courseName || !topicName || !entryDate) {
        alert('Course Name, Topic/Module and Date are required.');
        return;
    }

    updateData(data => {
        if (!Array.isArray(data.learningNotes)) data.learningNotes = [];
        if (id) {
            const idx = data.learningNotes.findIndex(n => n.id == id);
            if (idx !== -1) {
                data.learningNotes[idx] = {
                    ...data.learningNotes[idx],
                    course_name: courseName,
                    topic_name: topicName,
                    entry_date: entryDate,
                    key_concepts: keyConcepts,
                    detailed_notes: detailedNotes,
                    tags,
                    revision_status: revisionStatus,
                    last_revised_date: lastRevisedDate,
                    priority,
                    pinned,
                    updated_at: new Date().toISOString()
                };
            }
        } else {
            data.learningNotes.push({
                id: Date.now() + Math.random(),
                course_name: courseName,
                topic_name: topicName,
                entry_date: entryDate,
                key_concepts: keyConcepts,
                detailed_notes: detailedNotes,
                tags,
                revision_status: revisionStatus,
                last_revised_date: lastRevisedDate,
                priority,
                pinned,
                collapsed: false,
                created_at: new Date().toISOString(),
                updated_at: new Date().toISOString(),
                created_by: currentUser ? currentUser.id : null
            });
        }
    });

    closeLearningNoteModal();
    renderLearningNotes();
}

function toggleLearningRevisionMode() {
    learningRevisionMode = !learningRevisionMode;
    const btn = document.getElementById('learningRevisionModeBtn');
    if (btn) btn.textContent = learningRevisionMode ? 'Exit Revision Mode' : 'Revision Mode';
    renderLearningNotes();
}

function setLearningSelectedCourse(courseName) {
    learningSelectedCourse = courseName || '';
    const sel = document.getElementById('learningFilterCourse');
    if (sel) sel.value = learningSelectedCourse;
    renderLearningNotes();
}

function markLearningNoteRevised(noteId) {
    updateData(data => {
        const note = (data.learningNotes || []).find(n => n.id == noteId);
        if (!note) return;
        if (note.revision_status === 'not_revised') note.revision_status = 'revised_once';
        else if (note.revision_status === 'revised_once') note.revision_status = 'mastered';
        note.last_revised_date = formatDateString(new Date());
        note.updated_at = new Date().toISOString();
    });
    renderLearningNotes();
}

function toggleLearningPin(noteId) {
    updateData(data => {
        const note = (data.learningNotes || []).find(n => n.id == noteId);
        if (note) note.pinned = !note.pinned;
    });
    renderLearningNotes();
}

function toggleLearningCollapse(noteId) {
    updateData(data => {
        const note = (data.learningNotes || []).find(n => n.id == noteId);
        if (note) note.collapsed = !note.collapsed;
    });
    renderLearningNotes();
}

function deleteLearningNote(noteId) {
    if (!confirm('Delete this learning note?')) return;
    updateData(data => {
        data.learningNotes = (data.learningNotes || []).filter(n => n.id != noteId);
    });
    renderLearningNotes();
}

function getFilteredLearningNotes() {
    const data = getData();
    const notes = Array.isArray(data.learningNotes) ? data.learningNotes : [];
    const q = (document.getElementById('learningSearchInput')?.value || '').toLowerCase().trim();
    const courseFilter = document.getElementById('learningFilterCourse')?.value || '';
    const tagFilter = document.getElementById('learningFilterTag')?.value || '';
    const revisionFilter = document.getElementById('learningFilterRevision')?.value || '';
    const sortBy = document.getElementById('learningSortBy')?.value || 'date_desc';

    let out = notes.filter(n => {
        if (learningSelectedCourse && n.course_name !== learningSelectedCourse) return false;
        if (courseFilter && n.course_name !== courseFilter) return false;
        if (tagFilter && !(n.tags || []).includes(tagFilter)) return false;
        if (revisionFilter && n.revision_status !== revisionFilter) return false;
        if (q) {
            const hay = [
                n.course_name, n.topic_name, n.detailed_notes,
                ...(n.key_concepts || []), ...(n.tags || [])
            ].join(' ').toLowerCase();
            if (!hay.includes(q)) return false;
        }
        return true;
    });

    out.sort((a, b) => {
        if (!!a.pinned !== !!b.pinned) return a.pinned ? -1 : 1;
        if (sortBy === 'date_asc') return (a.entry_date || '').localeCompare(b.entry_date || '');
        if (sortBy === 'priority_desc') return getLearningPriorityWeight(b.priority) - getLearningPriorityWeight(a.priority);
        if (sortBy === 'last_revised_desc') return (b.last_revised_date || '').localeCompare(a.last_revised_date || '');
        return (b.entry_date || '').localeCompare(a.entry_date || '');
    });

    return out;
}

function renderLearningFlashcard(notes) {
    const box = document.getElementById('learningRevisionFlashcard');
    if (!box) return;
    if (!learningRevisionMode) {
        box.style.display = 'none';
        return;
    }

    learningFlashcards = notes.filter(n =>
        n.priority === 'high' || n.revision_status === 'not_revised'
    );
    if (learningFlashcards.length === 0) {
        box.style.display = 'block';
        box.innerHTML = `<div class="card" style="padding: 12px;"><strong>Revision Mode:</strong> No high priority / weak notes found.</div>`;
        return;
    }
    learningFlashcardIndex = Math.min(learningFlashcardIndex, learningFlashcards.length - 1);
    const n = learningFlashcards[learningFlashcardIndex];
    const firstConcept = (n.key_concepts && n.key_concepts.length) ? n.key_concepts[0] : n.topic_name;
    box.style.display = 'block';
    box.innerHTML = `
        <div class="card" style="padding: 12px; border-left: 4px solid #667eea;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 8px;">
                <strong>Revision Flashcard ${learningFlashcardIndex + 1}/${learningFlashcards.length}</strong>
                <div style="display:flex; gap:6px;">
                    <button class="btn btn-secondary" onclick="learningFlashcardIndex=Math.max(0, learningFlashcardIndex-1); renderLearningNotes()">Prev</button>
                    <button class="btn btn-secondary" onclick="learningFlashcardIndex=Math.min(learningFlashcards.length-1, learningFlashcardIndex+1); renderLearningNotes()">Next</button>
                </div>
            </div>
            <div style="display:grid; grid-template-columns: 1fr 1fr; gap: 10px;">
                <div style="padding: 10px; background:#f8f9fa; border-radius:6px;">
                    <div style="font-size:12px; color:#666; margin-bottom:4px;">Front (Key Concept)</div>
                    <div style="font-weight:600;">${escapeHtml(firstConcept || '(no concept)')}</div>
                </div>
                <div style="padding: 10px; background:#eef6ff; border-radius:6px;">
                    <div style="font-size:12px; color:#666; margin-bottom:4px;">Back (Explanation)</div>
                    <div style="white-space: pre-wrap;">${escapeHtml(n.detailed_notes || n.topic_name || '')}</div>
                </div>
            </div>
            <div style="margin-top:10px;">
                <button class="btn btn-success" onclick="markLearningNoteRevised(${n.id})">Mark as Revised</button>
            </div>
        </div>
    `;
}

function renderLearningNotes() {
    const data = getData();
    const allNotes = Array.isArray(data.learningNotes) ? data.learningNotes : [];
    const courseSel = document.getElementById('learningFilterCourse');
    const tagSel = document.getElementById('learningFilterTag');
    const courseList = document.getElementById('learningCourseList');
    const notesList = document.getElementById('learningNotesList');
    const progressBox = document.getElementById('learningCourseProgress');
    if (!courseSel || !tagSel || !courseList || !notesList || !progressBox) return;

    // Populate filter options
    const courses = Array.from(new Set(allNotes.map(n => n.course_name).filter(Boolean))).sort();
    const tags = Array.from(new Set(allNotes.flatMap(n => n.tags || []))).sort();
    const prevCourse = courseSel.value;
    const prevTag = tagSel.value;
    courseSel.innerHTML = `<option value="">All Courses</option>${courses.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('')}`;
    tagSel.innerHTML = `<option value="">All Tags</option>${tags.map(t => `<option value="${escapeHtml(t)}">${escapeHtml(t)}</option>`).join('')}`;
    if (prevCourse) courseSel.value = prevCourse;
    if (prevTag) tagSel.value = prevTag;

    // Left panel (courses + progress)
    const courseRows = courses.map(c => {
        const items = allNotes.filter(n => n.course_name === c);
        const revised = items.filter(n => n.revision_status === 'revised_once' || n.revision_status === 'mastered').length;
        const pct = items.length ? Math.round((revised / items.length) * 100) : 0;
        const active = (learningSelectedCourse === c) ? 'background:#eef6ff; border:1px solid #cfe2ff;' : 'background:#fff; border:1px solid #eee;';
        return `
            <div onclick='setLearningSelectedCourse(${JSON.stringify(c)})' style="cursor:pointer; padding:8px; border-radius:6px; margin-bottom:6px; ${active}">
                <div style="font-weight:600;">${escapeHtml(c)}</div>
                <div style="font-size:12px; color:#666;">${items.length} note(s) | ${pct}% revised</div>
                <div style="height:6px; background:#eee; border-radius:4px; margin-top:4px;">
                    <div style="height:6px; width:${pct}%; background:#28a745; border-radius:4px;"></div>
                </div>
            </div>
        `;
    }).join('');
    courseList.innerHTML = courseRows || '<div style="color:#999;">No courses yet.</div>';

    const totalRevised = allNotes.filter(n => n.revision_status === 'revised_once' || n.revision_status === 'mastered').length;
    const totalPct = allNotes.length ? Math.round((totalRevised / allNotes.length) * 100) : 0;
    progressBox.textContent = `Overall Progress: ${totalPct}% revised (${totalRevised}/${allNotes.length})`;

    const filtered = getFilteredLearningNotes();
    renderLearningFlashcard(filtered);

    // Right panel notes
    notesList.innerHTML = filtered.length ? filtered.map(n => {
        const priorityColor = n.priority === 'high' ? '#dc3545' : (n.priority === 'medium' ? '#d97706' : '#1890ff');
        const concepts = (n.key_concepts || []).map(k => `<li><span style="color:#5568d3; font-weight:600;">${escapeHtml(k)}</span></li>`).join('');
        const tagsHtml = (n.tags || []).map(t => `<span class="badge badge-info" style="font-size:11px;">${escapeHtml(t)}</span>`).join(' ');
        const collapsed = !!n.collapsed;
        const badgeRevisionClass = n.revision_status === 'mastered' ? 'badge-completed' : (n.revision_status === 'revised_once' ? 'badge-info' : 'badge-warning');
        return `
            <div class="task-item" style="padding:12px; border-left:4px solid ${priorityColor};">
                <div style="display:flex; justify-content:space-between; gap:10px; align-items:flex-start;">
                    <div>
                        <div style="font-weight:700;">${escapeHtml(n.topic_name)}</div>
                        <div style="font-size:12px; color:#666;">${escapeHtml(n.course_name)} | ${formatDateDisplay(n.entry_date)}</div>
                    </div>
                    <div style="display:flex; gap:6px; flex-wrap:wrap;">
                        ${n.pinned ? '<span class="badge badge-warning">Pinned</span>' : ''}
                        <span class="badge ${badgeRevisionClass}">${getLearningRevisionLabel(n.revision_status)}</span>
                        <span class="badge badge-secondary">${(n.priority || 'medium').toUpperCase()}</span>
                    </div>
                </div>
                <div style="margin-top:8px;">${tagsHtml}</div>
                ${collapsed ? '' : `
                    <div style="margin-top:10px;">
                        <strong>Key Concepts</strong>
                        <ul style="padding-left:18px; margin-top:4px;">${concepts || '<li>(none)</li>'}</ul>
                    </div>
                    ${n.detailed_notes ? `<div style="margin-top:8px; white-space:pre-wrap;"><strong>Detailed Notes</strong><div style="background:#f8f9fa; padding:8px; border-radius:4px; margin-top:4px;">${escapeHtml(n.detailed_notes)}</div></div>` : ''}
                `}
                <div style="margin-top:10px; display:flex; gap:6px; flex-wrap:wrap;">
                    <button class="btn btn-secondary" onclick="toggleLearningCollapse(${n.id})">${collapsed ? 'Expand' : 'Collapse'}</button>
                    <button class="btn btn-secondary" onclick="toggleLearningPin(${n.id})">${n.pinned ? 'Unpin' : 'Pin'}</button>
                    <button class="btn btn-success" onclick="markLearningNoteRevised(${n.id})">Mark as Revised</button>
                    <button class="btn btn-primary" onclick="openLearningNoteModal(${n.id})">Edit</button>
                    <button class="btn btn-danger" onclick="deleteLearningNote(${n.id})">Delete</button>
                </div>
                <div style="margin-top:6px; font-size:12px; color:#666;">Last Revised: ${n.last_revised_date ? formatDateDisplay(n.last_revised_date) : '—'}</div>
            </div>
        `;
    }).join('') : '<div style="color:#999; padding: 10px;">No learning notes found.</div>';
}

function exportLearningNotesToWord() {
    const notes = getFilteredLearningNotes();
    const title = 'Learning Notes Export';
    const html = `
        <html><head><meta charset="UTF-8"><title>${title}</title>
        <style>
            body { font-family: Arial, sans-serif; padding: 16px; }
            h1 { color: #333; }
            .course { margin-top: 18px; }
            .note { border: 1px solid #ddd; border-radius: 8px; padding: 10px; margin-top: 8px; }
            .meta { color: #666; font-size: 12px; margin-bottom: 6px; }
            ul { margin-top: 4px; }
        </style></head><body>
        <h1>${title}</h1>
        <div>Generated on: ${formatDateDisplay(new Date())}</div>
        ${notes.map(n => `
            <div class="course">
                <div class="note">
                    <h3>${escapeHtml(n.course_name)} - ${escapeHtml(n.topic_name)}</h3>
                    <div class="meta">Entry Date: ${formatDateDisplay(n.entry_date)} | Revision: ${getLearningRevisionLabel(n.revision_status)} | Priority: ${(n.priority || '').toUpperCase()} | Last Revised: ${n.last_revised_date ? formatDateDisplay(n.last_revised_date) : '—'}</div>
                    <div><strong>Tags:</strong> ${escapeHtml((n.tags || []).join(', ')) || '—'}</div>
                    <div><strong>Key Concepts:</strong>
                        <ul>${(n.key_concepts || []).map(k => `<li>${escapeHtml(k)}</li>`).join('') || '<li>—</li>'}</ul>
                    </div>
                    ${n.detailed_notes ? `<div><strong>Detailed Notes:</strong><div style="white-space: pre-wrap;">${escapeHtml(n.detailed_notes)}</div></div>` : ''}
                </div>
            </div>
        `).join('')}
        </body></html>
    `;
    const blob = new Blob(['\ufeff' + html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `learning-notes-${formatDateString(new Date())}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

async function bootstrapApp() {
    document.body.classList.toggle('app-api-mode', isApiMode());
    if (!isApiMode()) {
        init();
        return;
    }

    document.body.classList.add('app-bootstrapping');
    try {
        let res;
        try {
            res = await apiFetch('/api/auth/me');
        } catch (fetchErr) {
            console.warn('API bootstrap fetch failed', fetchErr);
            currentUser = null;
            __workspaceCache = null;
            clearApiAuthToken();
            localStorage.removeItem('currentUser');
            checkAuth();
            init();
            return;
        }

        if (res.ok) {
            try {
                currentUser = await res.json();
            } catch (parseErr) {
                console.warn('Bootstrap /me parse failed', parseErr);
                currentUser = null;
                clearApiAuthToken();
                localStorage.removeItem('currentUser');
                __workspaceCache = null;
                checkAuth();
                init();
                return;
            }
            localStorage.setItem('currentUser', JSON.stringify(currentUser));
            try {
                await apiPullWorkspace();
            } catch (pullErr) {
                console.error('Workspace load on bootstrap:', pullErr);
                __workspaceCache = normalizeData(defaultWorkspaceShell());
            }
        } else {
            currentUser = null;
            clearApiAuthToken();
            localStorage.removeItem('currentUser');
            __workspaceCache = null;
        }
    } finally {
        document.body.classList.remove('app-bootstrapping');
    }
    checkAuth();
    init();
}

wireLoginScreenControls();

// Initialize on load
window.addEventListener('DOMContentLoaded', () => {
    wireLoginScreenControls();
    bootstrapApp().catch(err => console.error(err));
});


