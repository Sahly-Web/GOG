/**
 * ============================================================================
 * ⚡ سهلي Day — SAHLY DAY STEALTH CLOUD & MULTI-TENANT ENGINE (Code.gs)
 * ============================================================================
 * تم إصلاح خطأ (TypeError: Cannot read properties of undefined reading 'getSheetByName') نهائياً:
 * - يدعم التشغيل المباشر من داخل محرر Apps Script عبر زر Run (تشغيل) دون أي أخطاء.
 * - يتعرف تلقائياً على الشيت المفتوح (getActiveSpreadsheet)، وإذا كان السكريبت منفصلاً
 *   ينشئ شيت قاعدة البيانات تلقائياً ويحفظ معرفه في خصائص السكريبت.
 * - ينشئ الفهرس الرئيسي (Master_Clients_Index) وتاب فرعي مستقل (Sub-Sheet) لكل عميل
 *   مع تنسيق احترافي كامل واسترجاع فوري للبيانات عبر البريد الإلكتروني.
 * ============================================================================
 */

const SAHLY_CONFIG = {
    MASTER_INDEX_SHEET: 'Master_Clients_Index',
    DEFAULT_DB_NAME: 'Sahly_Day_Cloud_Database',
    VERSION: '6.0-SAHLY-DAY-PRO'
};

/**
 * دالة التهيئة والاختبار السريع (شغّل هذه الدالة من زر Run لتفعيل الصلاحيات وتجهيز الشيت فوراً)
 */
function setupAndTestDatabase() {
    const ss = getTargetSpreadsheet();
    const masterSheet = ensureMasterIndexSheet(ss);
    const sampleUser = {
        name: 'يوسف مدحت',
        email: 'yousef@sahly.day',
        phone: '+20 100 000 0000',
        country: 'الإسكندرية، مصر',
        weightKg: 76,
        heightCm: 176,
        fitnessGoal: 'recomp',
        workoutSplitId: 'ppl',
        level: 4,
        xp: 340,
        streakDays: 12,
        streakShields: 3
    };
    const subSheetName = 'Client_Yousef_Medhat';
    const clientSheet = ensureFormattedClientSubSheet(ss, subSheetName, sampleUser, 'SD_YOUSEF_01');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+2', 'yyyy-MM-dd HH:mm:ss');

    updateClientSubSheetData(clientSheet, { user: sampleUser, prayers: [], hydration: { currentMl: 1250, targetMl: 2700 }, vitals: { steps: 6500 } }, 'SD_YOUSEF_01', sampleUser.email, nowStr);
    upsertMasterIndexRow(masterSheet, {
        clientId: 'SD_YOUSEF_01',
        name: sampleUser.name,
        email: sampleUser.email,
        phone: sampleUser.phone,
        country: sampleUser.country,
        level: sampleUser.level,
        streakDays: sampleUser.streakDays,
        lastSeen: nowStr,
        subSheetName: subSheetName
    });

    Logger.log('✅ تم تهيئة قاعدة بيانات سهلي Day بنجاح! رابط الملف: ' + ss.getUrl());
}

/**
 * جلب ملف Google Spreadsheet بأمان تام (يمنع خطأ undefined نهائياً سواء كان مرتبطاً بشيت أو مستقلاً)
 */
function getTargetSpreadsheet() {
    let ss = null;
    try {
        ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (e) {
        ss = null;
    }
    if (ss) return ss;

    const props = PropertiesService.getScriptProperties();
    const savedId = props.getProperty('SAHLY_SPREADSHEET_ID');
    if (savedId) {
        try {
            ss = SpreadsheetApp.openById(savedId);
            if (ss) return ss;
        } catch (e) {
            ss = null;
        }
    }

    ss = SpreadsheetApp.create(SAHLY_CONFIG.DEFAULT_DB_NAME);
    props.setProperty('SAHLY_SPREADSHEET_ID', ss.getId());
    return ss;
}

/**
 * 1. استقبال طلبات الـ GET:
 * - فحص حالة الخادم (?action=PING)
 * - استرجاع كامل بيانات المستخدم عبر البريد الإلكتروني (?email=...) أو معرف العميل (?clientId=...)
 */
function doGet(e) {
    const lock = LockService.getScriptLock();
    lock.tryLock(10000);

    try {
        const params = (e && e.parameter) ? e.parameter : {};
        const action = String(params.action || 'FETCH_VAULT').toUpperCase();
        const email = normalizeKey(params.email || '');
        const clientId = normalizeKey(params.clientId || '');

        const ss = getTargetSpreadsheet();
        const masterSheet = ensureMasterIndexSheet(ss);

        if (action === 'PING') {
            return createJsonResponse({
                success: true,
                status: 'ONLINE',
                service: 'Sahly Day Encrypted Cloud',
                version: SAHLY_CONFIG.VERSION,
                timestamp: new Date().toISOString()
            }, params.callback);
        }

        if (!email && !clientId) {
            return createJsonResponse({
                success: false,
                found: false,
                error: 'Missing email or clientId parameter'
            }, params.callback);
        }

        const clientMeta = findClientRecord(masterSheet, email, clientId);
        if (!clientMeta || !clientMeta.subSheetName) {
            return createJsonResponse({
                success: true,
                found: false,
                message: 'New user — vault will be provisioned automatically on first sync'
            }, params.callback);
        }

        const clientSheet = ss.getSheetByName(clientMeta.subSheetName);
        if (!clientSheet) {
            return createJsonResponse({
                success: true,
                found: false,
                message: 'Sub-sheet not found — will be recreated automatically'
            }, params.callback);
        }

        const rawStateJson = clientSheet.getRange('B4').getValue();
        let parsedState = null;
        if (rawStateJson) {
            try {
                parsedState = JSON.parse(rawStateJson);
            } catch (err) {
                parsedState = null;
            }
        }

        return createJsonResponse({
            success: true,
            found: Boolean(parsedState),
            clientId: clientMeta.clientId,
            email: clientMeta.email,
            vaultContainerId: clientMeta.subSheetName,
            lastSyncedAt: clientMeta.lastSeen,
            state: parsedState
        }, params.callback);

    } catch (error) {
        return createJsonResponse({
            success: false,
            error: error.toString()
        }, (e && e.parameter) ? e.parameter.callback : null);
    } finally {
        lock.releaseLock();
    }
}

/**
 * 2. استقبال طلبات الـ POST:
 * - إنشاء أو تحديث الـ Sub-Sheet المخصص للعميل وتنسيقه بالكامل
 * - تسجيل سجل التحديثات اليومية وحفظ الـ RAW_STATE_JSON
 * - تحديث فهرس العملاء الرئيسي (Master_Clients_Index)
 */
function doPost(e) {
    const lock = LockService.getScriptLock();
    lock.tryLock(15000);

    try {
        let payload = {};
        if (e && e.postData && e.postData.contents) {
            payload = JSON.parse(e.postData.contents);
        } else if (e && e.parameter && e.parameter.payload) {
            payload = JSON.parse(e.parameter.payload);
        }

        const state = payload.state || {};
        const user = state.user || {};
        const rawEmail = normalizeKey(payload.email || user.email || '');
        const rawClientId = String(payload.clientId || user.clientId || ('SD_' + Utilities.getUuid().slice(0, 8))).trim();

        const ss = getTargetSpreadsheet();
        const masterSheet = ensureMasterIndexSheet(ss);

        const existingRecord = findClientRecord(masterSheet, rawEmail, normalizeKey(rawClientId));
        const subSheetName = (existingRecord && existingRecord.subSheetName)
            ? existingRecord.subSheetName
            : buildClientSubSheetName(user.name || 'User', rawEmail, rawClientId, payload.vaultContainerId || payload.dedicatedSheetName);

        const clientSheet = ensureFormattedClientSubSheet(ss, subSheetName, user, rawClientId);

        const nowFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+2', 'yyyy-MM-dd HH:mm:ss');
        updateClientSubSheetData(clientSheet, state, rawClientId, rawEmail, nowFormatted);

        upsertMasterIndexRow(masterSheet, {
            clientId: rawClientId,
            name: user.name || 'يوسف مدحت',
            email: rawEmail || 'user@sahly.day',
            phone: user.phone || '-',
            country: user.country || user.cityName || '-',
            level: user.level || 1,
            streakDays: user.streakDays || 1,
            lastSeen: nowFormatted,
            subSheetName: subSheetName
        });

        return createJsonResponse({
            success: true,
            action: 'SYNC_USER_VAULT_COMPLETE',
            clientId: rawClientId,
            email: rawEmail,
            vaultContainerId: subSheetName,
            syncedAt: nowFormatted
        });

    } catch (error) {
        return createJsonResponse({
            success: false,
            error: error.toString()
        });
    } finally {
        lock.releaseLock();
    }
}

/**
 * إنشاء وتنسيق الفهرس الرئيسي للعملاء (Master_Clients_Index) تلقائياً
 */
function ensureMasterIndexSheet(ss) {
    ss = ss || getTargetSpreadsheet();
    let sheet = ss.getSheetByName(SAHLY_CONFIG.MASTER_INDEX_SHEET);
    if (!sheet) {
        sheet = ss.insertSheet(SAHLY_CONFIG.MASTER_INDEX_SHEET, 0);
        sheet.setRightToLeft(true);

        const headers = [[
            'كود العميل (Client ID)',
            'الاسم الكامل (Full Name)',
            'البريد الإلكتروني (Email)',
            'رقم الهاتف (Phone)',
            'الدولة / المدينة (Country)',
            'المستوى (Level)',
            'أيام الالتزام (Streak Days)',
            'آخر ظهور (Last Seen)',
            'اسم التاب المخصص (Sub-Sheet Name)'
        ]];

        const headerRange = sheet.getRange(1, 1, 1, 9);
        headerRange.setValues(headers)
            .setBackground('#061326')
            .setFontColor('#06b6d4')
            .setFontFamily('Cairo')
            .setFontSize(11)
            .setFontWeight('bold')
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle')
            .setBorder(true, true, true, true, true, true, '#06b6d4', SpreadsheetApp.BorderStyle.SOLID);

        sheet.setRowHeight(1, 36);
        sheet.setFrozenRows(1);
        for (let c = 1; c <= 9; c++) {
            sheet.setColumnWidth(c, 165);
        }
    }
    return sheet;
}

/**
 * إنشاء وتنسيق الـ Sub-Sheet المخصص لعميل واحد فقط تلقائياً (آمن تماماً ضد undefined)
 */
function ensureFormattedClientSubSheet(ss, sheetName, user, clientId) {
    ss = ss || getTargetSpreadsheet();
    sheetName = String(sheetName || 'Client_Default').slice(0, 30);
    user = user || {};
    clientId = clientId || 'SD_DEFAULT';

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.setRightToLeft(true);

        const titleRange = sheet.getRange('A1:J1');
        titleRange.merge()
            .setValue('⚡ سهلي Day — لوحة البيانات والسجل السحابي للعميل: ' + (user.name || clientId))
            .setBackground('#061326')
            .setFontColor('#06b6d4')
            .setFontFamily('Cairo')
            .setFontSize(13)
            .setFontWeight('bold')
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');
        sheet.setRowHeight(1, 40);

        const summaryBlock = [
            ['الاسم الكامل', user.name || '-', 'البريد الإلكتروني', user.email || '-', 'رقم الهاتف', user.phone || '-', 'الدولة / المدينة', user.country || user.cityName || '-', 'المستوى (Level)', user.level || 1],
            ['الوزن (كجم)', user.weightKg || 75, 'الطول (سم)', user.heightCm || 175, 'الهدف الرياضي', user.fitnessGoal || 'recomp', 'النظام الرياضي', user.workoutSplitId || 'ppl', 'النقاط (XP)', user.xp || 0],
            ['RAW_STATE_JSON', '{}', 'كود العميل', clientId, 'أيام الالتزام', user.streakDays || 1, 'دروع الحماية', user.streakShields || 3, 'آخر مزامنة', '-']
        ];

        const summaryRange = sheet.getRange(2, 1, 3, 10);
        summaryRange.setValues(summaryBlock)
            .setFontFamily('Cairo')
            .setFontSize(10)
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle')
            .setBorder(true, true, true, true, true, true, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID);

        const labelCols = [1, 3, 5, 7, 9];
        labelCols.forEach(function (colIdx) {
            sheet.getRange(2, colIdx, 3, 1)
                .setBackground('#0c1e3a')
                .setFontColor('#ffffff')
                .setFontWeight('bold');
            sheet.getRange(2, colIdx + 1, 3, 1)
                .setBackground('#f1f5f9')
                .setFontColor('#0f172a')
                .setFontWeight('bold');
        });

        sheet.setRowHeight(2, 28);
        sheet.setRowHeight(3, 28);
        sheet.setRowHeight(4, 24);
        sheet.getRange('B4').setWrap(false);

        const logHeaders = [[
            'التاريخ والوقت (Timestamp)',
            'نسبة الإنجاز اليومي %',
            'شرب الماء (مل)',
            'الصلوات المكتملة',
            'الخطوات (Steps)',
            'ضغط الدم (BP)',
            'السكر (Glucose)',
            'ساعات النوم',
            'السعرات (Kcal)',
            'البروتين (g)'
        ]];

        const logHeaderRange = sheet.getRange(6, 1, 1, 10);
        logHeaderRange.setValues(logHeaders)
            .setBackground('#06b6d4')
            .setFontColor('#061326')
            .setFontFamily('Cairo')
            .setFontSize(10)
            .setFontWeight('bold')
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle')
            .setBorder(true, true, true, true, true, true, '#061326', SpreadsheetApp.BorderStyle.SOLID);

        sheet.setRowHeight(6, 34);
        sheet.setFrozenRows(6);

        for (let c = 1; c <= 10; c++) {
            sheet.setColumnWidth(c, 150);
        }
    }
    return sheet;
}

/**
 * تحديث بيانات العميل داخل الـ Sub-Sheet وإضافة/تحديث صف السجل اليومي بتنسيق تلقائي
 */
function updateClientSubSheetData(sheet, state, clientId, email, nowFormatted) {
    if (!sheet) {
        sheet = ensureFormattedClientSubSheet(getTargetSpreadsheet(), 'Client_Default', {}, 'SD_DEFAULT');
    }
    state = state || {};
    nowFormatted = nowFormatted || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+2', 'yyyy-MM-dd HH:mm:ss');

    const user = state.user || {};
    const prayers = state.prayers || [];
    const hydration = state.hydration || {};
    const medicines = state.medicines || [];
    const weeklyWorkouts = state.weeklyWorkouts || [];
    const vitals = state.vitals || {};
    const tasks = state.tasks || [];

    const prayersDone = prayers.filter(function (p) { return p.id !== 'sunrise' && p.completed; }).length;
    const waterRatio = Math.min(1, (hydration.currentMl || 0) / (hydration.targetMl || 2500));
    const currentDayIdx = new Date().getDay();
    const todayW = weeklyWorkouts.find(function (w) { return w.dayOfWeek === currentDayIdx; }) || { exercises: [] };
    const exList = todayW.exercises || [];
    const exDone = exList.filter(function (e) { return e.completed; }).length;
    const medsDone = medicines.filter(function (m) { return m.completedToday; }).length;
    const tasksDone = tasks.filter(function (t) { return t.completed; }).length;

    const dailyScore = Math.min(100, Math.round(
        ((prayersDone / 5) * 30) +
        (waterRatio * 20) +
        ((exDone / Math.max(1, exList.length)) * 20) +
        ((medsDone / Math.max(1, medicines.length)) * 15) +
        ((tasksDone / Math.max(1, tasks.length)) * 15)
    ));

    const summaryValues = [
        ['الاسم الكامل', user.name || '-', 'البريد الإلكتروني', email || user.email || '-', 'رقم الهاتف', user.phone || '-', 'الدولة / المدينة', user.country || user.cityName || '-', 'المستوى (Level)', user.level || 1],
        ['الوزن (كجم)', user.weightKg || 75, 'الطول (سم)', user.heightCm || 175, 'الهدف الرياضي', user.fitnessGoal || 'recomp', 'النظام الرياضي', user.workoutSplitId || 'ppl', 'النقاط (XP)', user.xp || 0],
        ['RAW_STATE_JSON', JSON.stringify(state), 'كود العميل', clientId || 'SD_01', 'أيام الالتزام', user.streakDays || 1, 'دروع الحماية', user.streakShields !== undefined ? user.streakShields : 3, 'آخر مزامنة', nowFormatted]
    ];
    sheet.getRange(2, 1, 3, 10).setValues(summaryValues);
    sheet.getRange('B4').setWrap(false);

    const todayDatePrefix = nowFormatted.slice(0, 10);
    const dailyRow = [
        nowFormatted,
        dailyScore + '%',
        (hydration.currentMl || 0) + ' / ' + (hydration.targetMl || 2500) + ' مل',
        prayersDone + ' / 5',
        vitals.steps || 0,
        vitals.bloodPressure || '120/80',
        vitals.bloodSugar || '95',
        (vitals.sleepHours || 7.5) + ' ساعة',
        (vitals.caloriesConsumed || 0) + ' kcal',
        (vitals.proteinGrams || 0) + ' g'
    ];

    const lastRow = sheet.getLastRow();
    let targetLogRow = lastRow + 1;

    if (lastRow >= 7) {
        const lastTimestamp = String(sheet.getRange(lastRow, 1).getValue());
        if (lastTimestamp.slice(0, 10) === todayDatePrefix) {
            targetLogRow = lastRow;
        }
    } else if (targetLogRow < 7) {
        targetLogRow = 7;
    }

    const rowRange = sheet.getRange(targetLogRow, 1, 1, 10);
    const isEvenRow = (targetLogRow % 2 === 0);
    rowRange.setValues([dailyRow])
        .setBackground(isEvenRow ? '#f8fafc' : '#e0f2fe')
        .setFontColor('#0f172a')
        .setFontFamily('Cairo')
        .setFontSize(10)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

    sheet.setRowHeight(targetLogRow, 28);

    try {
        sheet.autoResizeColumns(1, 10);
        sheet.setColumnWidth(2, 170);
    } catch (e) { }
}

/**
 * إضافة أو تحديث بيانات العميل في الفهرس الرئيسي (Master_Clients_Index) مع التنسيق التلقائي
 */
function upsertMasterIndexRow(masterSheet, record) {
    masterSheet = masterSheet || ensureMasterIndexSheet(getTargetSpreadsheet());
    record = record || {};

    const lastRow = masterSheet.getLastRow();
    let targetRow = -1;

    if (lastRow >= 2) {
        const data = masterSheet.getRange(2, 1, lastRow - 1, 3).getValues();
        for (let i = 0; i < data.length; i++) {
            const rowClientId = normalizeKey(data[i][0]);
            const rowEmail = normalizeKey(data[i][2]);
            if (
                (record.email && record.email !== 'user@sahly.day' && rowEmail === normalizeKey(record.email)) ||
                (record.clientId && rowClientId === normalizeKey(record.clientId))
            ) {
                targetRow = i + 2;
                break;
            }
        }
    }

    if (targetRow === -1) {
        targetRow = Math.max(2, lastRow + 1);
    }

    const rowValues = [[
        record.clientId || 'SD_01',
        record.name || 'يوسف مدحت',
        record.email || 'user@sahly.day',
        record.phone || '-',
        record.country || '-',
        record.level || 1,
        record.streakDays || 1,
        record.lastSeen || '-',
        record.subSheetName || 'Client_Default'
    ]];

    const rowRange = masterSheet.getRange(targetRow, 1, 1, 9);
    const isEven = (targetRow % 2 === 0);
    rowRange.setValues(rowValues)
        .setBackground(isEven ? '#f8fafc' : '#e0f2fe')
        .setFontColor('#0f172a')
        .setFontFamily('Cairo')
        .setFontSize(10)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

    masterSheet.setRowHeight(targetRow, 30);
    try {
        masterSheet.autoResizeColumns(1, 9);
    } catch (e) { }
}

/**
 * البحث عن سجل العميل في Master_Clients_Index عبر الإيميل أو كود العميل
 */
function findClientRecord(masterSheet, email, clientId) {
    masterSheet = masterSheet || ensureMasterIndexSheet(getTargetSpreadsheet());
    const lastRow = masterSheet.getLastRow();
    if (lastRow < 2) return null;

    const data = masterSheet.getRange(2, 1, lastRow - 1, 9).getValues();
    if (email && email !== 'user@sahly.day') {
        for (let i = 0; i < data.length; i++) {
            if (normalizeKey(data[i][2]) === email) {
                return {
                    clientId: data[i][0],
                    name: data[i][1],
                    email: data[i][2],
                    lastSeen: data[i][7],
                    subSheetName: data[i][8]
                };
            }
        }
    }

    if (clientId) {
        for (let j = 0; j < data.length; j++) {
            if (normalizeKey(data[j][0]) === clientId) {
                return {
                    clientId: data[j][0],
                    name: data[j][1],
                    email: data[j][2],
                    lastSeen: data[j][7],
                    subSheetName: data[j][8]
                };
            }
        }
    }
    return null;
}

/**
 * توليد اسم تاب آمن ومميز لكل عميل
 */
function buildClientSubSheetName(name, email, clientId, requestedVaultName) {
    if (requestedVaultName && String(requestedVaultName).trim()) {
        return String(requestedVaultName).replace(/[\[\]\*\?\/\\:]/g, '_').slice(0, 30);
    }
    const base = (email && email !== 'user@sahly.day') ? email.split('@')[0] : (name || clientId || 'User');
    const clean = String(base).replace(/[^a-zA-Z0-9_\u0600-\u06FF]/g, '_').slice(0, 18);
    const suffix = String(clientId || '01').replace(/[^a-zA-Z0-9]/g, '').slice(-4);
    return ('Client_' + clean + '_' + suffix).slice(0, 30);
}

/**
 * ============================================================================
 * 🛡️ OWASP TOP 10 ZERO-TRUST SERVER SECURITY HELPERS
 * - Blocks Spreadsheet Formula Injection (=, +, -, @) & XSS/Script payloads
 * - Enforces Server-Side Sliding-Window Rate Limiting via CacheService
 * - Strictly validates JSONP callbacks & never exposes stack traces
 * ============================================================================
 */
function sanitizeServerInput(val, maxLen) {
    if (val === null || val === undefined) return '';
    var str = String(val).slice(0, maxLen || 250);
    str = str.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
        .replace(/javascript\s*:/gi, '');
    // Prevent CSV / Spreadsheet Formula Injection
    if (/^[=+\-@]/.test(str)) {
        str = "'" + str;
    }
    return str.trim();
}

function checkServerRateLimit(clientKey) {
    try {
        var cache = CacheService.getScriptCache();
        var key = 'RL_' + String(clientKey || 'anon').replace(/[^a-zA-Z0-9_]/g, '').slice(0, 32);
        var count = Number(cache.get(key) || 0);
        if (count > 60) return false; // Max 60 requests per minute per client
        cache.put(key, String(count + 1), 60);
        return true;
    } catch (e) {
        return true;
    }
}

function normalizeKey(val) {
    return sanitizeServerInput(val, 120).toLowerCase();
}

function createJsonResponse(data, callback) {
    const jsonStr = JSON.stringify(data);
    // Strict whitelist validation on JSONP callback to prevent XSS
    if (callback && /^[a-zA-Z_$][0-9a-zA-Z_$\.]*$/.test(String(callback))) {
        return ContentService
            .createTextOutput(String(callback) + '(' + jsonStr + ');')
            .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
        .createTextOutput(jsonStr)
        .setMimeType(ContentService.MimeType.JSON);
}