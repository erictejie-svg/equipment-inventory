/*************************************************
 * PHDU INVENTORY SYSTEM - ULTRA BOOST 3
 * Strict, speed-first, modal Fast Scan workflow
 * Locked settings version
 * Locked format + locked workflow
 *************************************************/

var MASTER_SHEET = 'Master Database';
var UPDATE_LOG_SHEET = 'Update_Log_New';
var FAST_SCAN_SHEET = 'Fast Scan Sessions';
var SCAN_AUDIT_SHEET = 'Scan Audit';
var FORM_RESPONSES_SHEET = 'Form Responses 2';

const ARCHIVE_SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1dxzgIw2XStoZEF6_5c7pRAGd7LXPLiYXC6qDY5qXbUM/edit?gid=494021773#gid=494021773';
const FAST_SCAN_EDIT_WINDOW_HOURS = 12;
const SERVICE_DUE_DAYS = 30;
const DASHBOARD_CACHE_KEY = 'PHDU_DASHBOARD_CACHE_V3';
const DASHBOARD_CACHE_SECONDS = 600;

function invalidateDashboardCache_() {
  try {
    CacheService.getScriptCache().remove(DASHBOARD_CACHE_KEY);
  } catch (e) {}
}

function getDashboardSummaryResult_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(MASTER_SHEET);
  if (!master) throw new Error('Master Database sheet not found.');

  var lastRow = master.getLastRow();
  var lastCol = master.getLastColumn();
  var data = lastRow > 0 && lastCol > 0
    ? master.getRange(1, 1, lastRow, lastCol).getValues()
    : [];

  var analysis = buildInventoryAnalysis_(data);
  var logMeta = getLogMeta_();
  var recentUpdates = getRecentUpdates_();

  return {
    inventoryMeta: {
      inventoryStatus: analysis.inventoryComplete ? 'INVENTORY DONE' : 'INVENTORY NOT COMPLETE',
      inventoryComplete: analysis.inventoryComplete,
      serviceAlertText: analysis.serviceAlertText,
      lastCheckedBy: logMeta.lastCheckedBy,
      lastUpdate: logMeta.lastUpdate,
      todayLogs: logMeta.todayLogs,
      weekLogs: logMeta.weekLogs,
      totalLogs: logMeta.totalLogs,
      scannedThisRound: analysis.scannedThisRound,
      accountedCount: analysis.accountedCount,
      totalEquipment: analysis.totals.total,
      completionPercent: analysis.completionPercent,
      alertWorkingCount: analysis.alertWorkingCount,
      unscannedTotal: analysis.unscannedTotal,
      overdueServiceCount: analysis.serviceDueSummary.overdue,
      dueSoonServiceCount: analysis.serviceDueSummary.dueSoon
    },
    recentUpdates: (recentUpdates || []).slice(0, 20),
totals: analysis.totals,
equipment: analysis.equipment,
unscannedItems: analysis.unscannedItems,
alertItems: analysis.alertItems,
serviceDueItems: {
  overdue: analysis.serviceDueItems.overdue,
  dueSoon: analysis.serviceDueItems.dueSoon
}

  };
}
var EMAIL_TO = [
  'erictejie@gmail.com',
  'mnjmohan817@gmail.com',
  'maryam.alhakmani@moh.gov.om'
].join(',');

var EMAIL_SENDER_NAME = 'PHDU Inventory System';

/*************************************************
 * WEB APP
 *************************************************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('PHDU Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/*************************************************
 * DASHBOARD DATA
 *************************************************/
function getInitialData(forceRefresh) {
  try {
    var useCache = !forceRefresh;
    var cache = CacheService.getScriptCache();

    if (useCache) {
      var cached = cache.get(DASHBOARD_CACHE_KEY);
      if (cached) {
        return JSON.parse(cached);
      }
    }

    var result = getDashboardSummaryResult_();

    try {
      cache.put(
        DASHBOARD_CACHE_KEY,
        JSON.stringify(result),
        DASHBOARD_CACHE_SECONDS
      );
    } catch (cacheErr) {
      Logger.log('Dashboard cache skipped: ' + cacheErr.message);
    }

    return result;
    } catch (err) {
    logError_(err, 'getInitialData');
    throw err;
  }
}


/*************************************************
 * FORM SUBMIT
 *************************************************/
function onFormSubmit(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var master = ss.getSheetByName(MASTER_SHEET);
    if (!master) throw new Error('Master Database sheet not found.');
    if (!e || !e.namedValues) throw new Error('No form submission data received.');

    var named = e.namedValues;

    var biomed = getNamedValue_(named, [
      'BIOMEDICAL NUMBER',
      'Biomedical Number',
      'QR Code ID',
      'Biomed Number'
    ]);

    var equipmentName = getNamedValue_(named, [
      'EQUIPMENT NAME',
      'Equipment Name'
    ]);

    var status = getNamedValue_(named, [
      'STATUS',
      'Status'
    ]);

    var location = getNamedValue_(named, [
      'LOCATION',
      'Location',
      'LOCATION/WARD',
      'Location/Ward'
    ]);

    var remarks = getNamedValue_(named, [
      'REMARKS',
      'Remarks'
    ]);

    var checkedBy = getNamedValue_(named, [
      'INVENTORY DONE',
      'Inventory Done',
      'INVENTORY DONE BY',
      'Inventory Done By',
      'INVENTORY DONE BY:',
      'Inventory Done By:',
      'Checked By',
      'CHECKED BY',
      'Done By',
      'DONE BY',
      'Name',
      'Staff Name',
      'EMAIL ADDRESS',
      'Email Address'
    ]) || getNamedValueLoose_(named, 'INVENTORY DONE')
       || getNamedValueLoose_(named, 'DONE BY')
       || getNamedValueLoose_(named, 'CHECKED BY')
       || getNamedValueLoose_(named, 'EMAIL ADDRESS');

    var inventoryDoneBy =
      getNamedValue_(named, [
        'INVENTORY DONE',
        'Inventory Done',
        'INVENTORY DONE BY',
        'Inventory Done By',
        'INVENTORY DONE BY:',
        'Inventory Done By:',
        'Done By',
        'DONE BY',
        'Checked By',
        'CHECKED BY'
      ]) || getNamedValueLoose_(named, 'INVENTORY DONE')
         || getNamedValueLoose_(named, 'DONE BY')
         || getNamedValueLoose_(named, 'CHECKED BY');

    var lastServiceDate =
      getNamedValue_(named, [
        'LAST SERVICE DATE',
        'Last Service Date',
        'LAST SERVICE DATE:',
        'Last Service Date:'
      ]) || getNamedValueLoose_(named, 'LAST SERVICE DATE');

    var nextServiceDue =
      getNamedValue_(named, [
        'NEXT SERVICE DUE',
        'Next Service Due',
        'SERVICE DUE',
        'Service Due',
        'LAST SERVICE DUE',
        'Last Service Due'
      ]) || getNamedValueLoose_(named, 'SERVICE DUE');

    if (!biomed) throw new Error('Biomedical Number not found in form submission.');

    var statusUpper = String(status || '').trim().toUpperCase();
    var markSeen = statusUpper === 'WORKING';
    var clearSeen = !!statusUpper && statusUpper !== 'WORKING';
    var finalCheckedBy = String(inventoryDoneBy || checkedBy || '').trim() || 'Unknown';

    updateMasterByBiomed_(master, {
      biomed: biomed,
      equipmentName: equipmentName,
      status: status,
      location: location,
      remarks: remarks,
      checkedBy: finalCheckedBy,
      inventoryDoneBy: finalCheckedBy,
      lastServiceDate: lastServiceDate,
      nextServiceDue: nextServiceDue,
      markSeen: markSeen,
      clearSeen: clearSeen,
      updateLastScanned: true
    });

    appendUpdateLog_({
      timestamp: new Date(),
      source: 'FORM',
      biomed: biomed,
      equipmentName: equipmentName,
      status: status,
      location: location,
      checkedBy: finalCheckedBy,
      note: buildFormNote_(finalCheckedBy, lastServiceDate, nextServiceDue, remarks)
    });

    invalidateDashboardCache_();
    return 'Form update processed.';
  } catch (err) {
    logError_(err, 'onFormSubmit');
    throw err;
  }
}

function buildFormNote_(staffName, lastServiceDate, nextServiceDue, remarks) {
  var parts = ['FORM QR update by ' + (String(staffName || '').trim() || 'Unknown')];
  if (lastServiceDate) parts.push('Last Service Date updated');
  if (nextServiceDue) parts.push('Next Service Due updated');
  if (remarks) parts.push('Remarks: ' + remarks);
  return parts.join(' | ');
}

/*************************************************
 * FAST SCAN SESSION
 *************************************************/
function startFastScanSession(staffName) {
  try {
    var sh = getOrCreateFastScanSheet_();
    var cleanStaff = String(staffName || '').trim() || 'Unknown';
    var existing = findEditableSessionByStaff_(sh, cleanStaff, FAST_SCAN_EDIT_WINDOW_HOURS);

    if (existing && String(existing.status || '').trim().toUpperCase() === 'OPEN') {
      return {
        sessionId: existing.sessionId,
        staffName: cleanStaff,
        status: 'OPEN',
        scannedCodes: parseStoredCodes_(existing.scannedCodes),
        duplicateCodes: parseStoredCodesRaw_(existing.duplicateCodes),
        notFoundCodes: parseStoredCodes_(existing.notFoundCodes),
        uniqueCount: Number(existing.uniqueCount || 0),
        duplicateCount: Number(existing.duplicateCount || 0),
        notFoundCount: Number(existing.notFoundCount || 0),
        inventoryRound: existing.inventoryRound || buildInventoryRound_(),
        resumed: true,
        editableUntil: formatAnyDate_(existing.editableUntil),
        revisionNo: Number(existing.revisionNo || 0)
      };
    }

    var sessionId = 'FS-' + new Date().getTime();
    var now = new Date();
    var inventoryRound = buildInventoryRound_();

    appendRowsToSheet_(sh, [[
      sessionId,
      cleanStaff,
      now,
      '',
      'OPEN',
      '',
      '',
      '',
      0,
      0,
      0,
      inventoryRound,
      now,
      now,
      0
    ]]);

    return {
      sessionId: sessionId,
      staffName: cleanStaff,
      status: 'OPEN',
      scannedCodes: [],
      duplicateCodes: [],
      notFoundCodes: [],
      uniqueCount: 0,
      duplicateCount: 0,
      notFoundCount: 0,
      inventoryRound: inventoryRound,
      resumed: false,
      editableUntil: formatAnyDate_(addHours_(now, FAST_SCAN_EDIT_WINDOW_HOURS)),
      revisionNo: 0
    };
  } catch (err) {
    logError_(err, 'startFastScanSession');
    throw err;
  }
}

function saveFastScanDraft(payload) {
  try {
    payload = payload || {};

    var sessionId = String(payload.sessionId || '').trim();
    var staffName = String(payload.staffName || '').trim() || 'Unknown';
    if (!sessionId) throw new Error('Session ID is required.');

    var sh = getOrCreateFastScanSheet_();
    var sessionInfo = findSessionRowInfo_(sh, sessionId);
    if (!sessionInfo) throw new Error('Fast scan session not found.');

    var now = new Date();
    var scannedCodes = normalizeUniqueCodesArray_(payload.scannedCodes || []);
    var duplicateCodes = normalizeCodesArray_(payload.duplicateCodes || []);
    var notFoundCodes = normalizeUniqueCodesArray_(payload.notFoundCodes || []);

    sh.getRange(sessionInfo.row, 2, 1, 14).setValues([[
      staffName,
      sessionInfo.startTime || now,
      '',
      'OPEN',
      scannedCodes.join('\n'),
      duplicateCodes.join('\n'),
      notFoundCodes.join('\n'),
      scannedCodes.length,
      duplicateCodes.length,
      notFoundCodes.length,
      sessionInfo.inventoryRound || buildInventoryRound_(),
      now,
      now,
      sessionInfo.revisionNo || 0
    ]]);

    return {
      success: true,
      sessionId: sessionId,
      status: 'OPEN',
      unique: scannedCodes.length,
      duplicate: duplicateCodes.length,
      notFound: notFoundCodes.length,
      message: 'Draft saved.'
    };
  } catch (err) {
    logError_(err, 'saveFastScanDraft');
    throw err;
  }
}

function finishFastScanSession(payload) {
  try {
    payload = payload || {};

    var sessionId = String(payload.sessionId || '').trim();
    var staffName = String(payload.staffName || '').trim() || 'Unknown';
    if (!sessionId) throw new Error('Session ID is required.');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var master = ss.getSheetByName(MASTER_SHEET);
    if (!master) throw new Error('Master Database sheet not found.');

    var sessionSheet = getOrCreateFastScanSheet_();
    var auditSheet = getOrCreateScanAuditSheet_();
    var sessionInfo = findSessionRowInfo_(sessionSheet, sessionId);
    if (!sessionInfo) throw new Error('Fast scan session not found.');

    var scannedCodes = normalizeUniqueCodesArray_(payload.scannedCodes || []);
    var duplicateCodes = normalizeCodesArray_(payload.duplicateCodes || []);
    if (!scannedCodes.length) throw new Error('No scanned codes found.');

    var masterData = master.getDataRange().getValues();
    if (masterData.length < 2) throw new Error('Master Database is empty.');

    var headers = masterData[0];
    var rows = masterData.slice(1);

    var idxBiomed = findHeaderIndex_(headers, ['QR Code ID', 'Biomedical Number', 'BIOMEDICAL NUMBER', 'Biomed Number']);
    var idxName = findHeaderIndex_(headers, ['Equipment Name', 'EQUIPMENT NAME']);
    if (idxBiomed === -1) throw new Error('Biomedical column not found');
    var idxStatus = findHeaderIndex_(headers, ['Status', 'STATUS']);
    var idxLocation = findHeaderIndex_(headers, ['Location/Ward', 'LOCATION/WARD', 'Location', 'LOCATION']);
    var idxRemarks = findHeaderIndex_(headers, ['Remarks', 'REMARKS']);
    var idxSeen = findHeaderIndex_(headers, ['SEEN THIS ROUND', 'Seen This Round']);
    var idxLastScannedDate = findHeaderIndex_(headers, ['LAST SCANNED DATE', 'Last Scanned Date']);
    var idxLastScannedBy = findHeaderIndex_(headers, ['LAST SCANNED BY', 'Last Scanned By']);
    var idxInventoryDone = findHeaderIndex_(headers, ['INVENTORY DONE', 'Inventory Done']);

    if (idxBiomed === -1) throw new Error('Biomedical Number column not found in Master Database.');

    var map = {};
    for (var i = 0; i < rows.length; i++) {
      var code = normalizeScanCode_(rows[i][idxBiomed]);
      if (!code) continue;
      map[code] = {
        rowIndex: i,
        biomed: String(rows[i][idxBiomed] || '').trim(),
        name: idxName > -1 ? String(rows[i][idxName] || '').trim() : '',
        location: idxLocation > -1 ? String(rows[i][idxLocation] || '').trim() : '',
        remarks: idxRemarks > -1 ? String(rows[i][idxRemarks] || '').trim() : ''
      };
    }

    var now = new Date();
    var inventoryRound = sessionInfo.inventoryRound || buildInventoryRound_();
    var matchedCodes = [];
    var notFoundCodes = [];
    var auditRows = [];
    var updateRows = [];
    var archiveFastScanRows = [];
    var scannedSet = {};
    var changedIndexes = {};

    if (idxSeen > -1) {
      for (var r = 0; r < rows.length; r++) {
        if (String(rows[r][idxSeen] || '').trim().toUpperCase() === 'YES') {
          rows[r][idxSeen] = '';
          changedIndexes[r] = true;
        }
      }
    }

    scannedCodes.forEach(function(code) {
      if (scannedSet[code]) return;
      scannedSet[code] = true;

      var match = map[code];
      if (!match) {
        notFoundCodes.push(code);

        auditRows.push([
          now, sessionId, staffName, 'FAST_SCAN', code, code,
          'NOT_FOUND', '', '', '', '', 'Code not found in Master Database'
        ]);

        archiveFastScanRows.push([now, sessionId, code, staffName, inventoryRound, 'NOT_FOUND']);
        return;
      }

      matchedCodes.push(code);

      var row = rows[match.rowIndex];
      if (idxStatus > -1 && String(row[idxStatus] || '').trim().toUpperCase() !== 'WORKING') {
        row[idxStatus] = 'WORKING';
      }
      if (idxSeen > -1) row[idxSeen] = 'YES';
      if (idxLastScannedDate > -1) row[idxLastScannedDate] = now;
      if (idxLastScannedBy > -1) row[idxLastScannedBy] = staffName;
      if (idxInventoryDone > -1) row[idxInventoryDone] = staffName;
      changedIndexes[match.rowIndex] = true;

      updateRows.push([
        now, 'FAST_SCAN', match.biomed, match.name, 'WORKING', match.location, staffName,
        buildFastScanNote_('Scanned via Fast Scan -> set to WORKING', match.remarks)
      ]);

      auditRows.push([
        now, sessionId, staffName, 'FAST_SCAN', code, code,
        'MATCHED', match.biomed, match.name, 'WORKING', match.location,
        'Scanned -> set to WORKING'
      ]);

      archiveFastScanRows.push([now, sessionId, match.biomed, staffName, inventoryRound, 'MATCHED']);
    });

    duplicateCodes.forEach(function(code) {
      auditRows.push([
        now, sessionId, staffName, 'FAST_SCAN', code, code,
        'DUPLICATE_IGNORED', '', '', '', '', 'Duplicate scan ignored'
      ]);

      archiveFastScanRows.push([now, sessionId, code, staffName, inventoryRound, 'DUPLICATE_IGNORED']);
    });

    for (var x = 0; x < rows.length; x++) {
      var rowCode = idxBiomed > -1 ? normalizeScanCode_(rows[x][idxBiomed]) : '';
      if (!rowCode) continue;
      if (scannedSet[rowCode]) continue;

      var statusUpper = idxStatus > -1 ? String(rows[x][idxStatus] || '').trim().toUpperCase() : '';

      if (idxInventoryDone > -1 && rows[x][idxInventoryDone] !== staffName) {
        rows[x][idxInventoryDone] = staffName;
        changedIndexes[x] = true;
      }

      if (statusUpper === 'WORKING') {
        if (idxStatus > -1) rows[x][idxStatus] = 'UNACCOUNTED';
        changedIndexes[x] = true;

        var itemName = idxName > -1 ? String(rows[x][idxName] || '').trim() : '';
        var itemLocation = idxLocation > -1 ? String(rows[x][idxLocation] || '').trim() : '';
        var itemRemarks = idxRemarks > -1 ? String(rows[x][idxRemarks] || '').trim() : '';

        updateRows.push([
          now, 'FAST_SCAN_FINALIZE', rowCode, itemName, 'UNACCOUNTED', itemLocation, staffName,
          buildFastScanNote_('Previously working / previously scanned but not scanned now', itemRemarks)
        ]);
      }
    }

    writeBackRowBlocks_(master, headers.length, rows, Object.keys(changedIndexes));

    if (auditRows.length) {
      appendRowsToSheet_(auditSheet, auditRows);
    }

    if (updateRows.length) {
      appendUpdateLogRows_(updateRows);
    }

    var nextRevision = sessionInfo.status === 'CLOSED'
      ? Number(sessionInfo.revisionNo || 0) + 1
      : Number(sessionInfo.revisionNo || 0);

    sessionSheet.getRange(sessionInfo.row, 2, 1, 14).setValues([[
      staffName,
      sessionInfo.startTime || now,
      now,
      'CLOSED',
      matchedCodes.join('\n'),
      duplicateCodes.join('\n'),
      notFoundCodes.join('\n'),
      matchedCodes.length,
      duplicateCodes.length,
      notFoundCodes.length,
      inventoryRound,
      now,
      now,
      nextRevision
    ]]);

    SpreadsheetApp.flush();

    archiveFastScanRows_(archiveFastScanRows);
    archiveUpdateLogRows_(updateRows, inventoryRound);
    archiveSessionRow_(sessionId);

    invalidateDashboardCache_();

    var analysis = buildInventoryAnalysis_([headers].concat(rows));

    return {
      unique: matchedCodes.length,
      duplicate: duplicateCodes.length,
      notFound: notFoundCodes.length,
      notFoundCodes: notFoundCodes,
      ignoredUrls: 0,
      completionPercent: analysis.completionPercent,
      unscannedItems: analysis.unscannedItems,
      unscannedSummary: analysis.unscannedSummary,
      message:
        'Fast scan saved. ' +
        matchedCodes.length + ' unique, ' +
        duplicateCodes.length + ' duplicate ignored, ' +
        notFoundCodes.length + ' not found.'
    };
  } catch (err) {
    logError_(err, 'finishFastScanSession');
    throw err;
  }
}

function buildFastScanNote_(base, remarks) {
  return remarks ? base + ' | Remarks: ' + remarks : base;
}

function saveFastScanSession(sessionId, scannedCodes, staffName) {
  return finishFastScanSession({
    sessionId: sessionId,
    staffName: staffName || 'Unknown',
    scannedCodes: String(scannedCodes || '').split(/\r?\n/),
    duplicateCodes: []
  });
}

/*************************************************
 * MANUAL FINISH INVENTORY
 *************************************************/
function finalizeInventory(staffName) {
  try {
    var cleanStaff = String(staffName || '').trim() || 'Unknown';

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var master = ss.getSheetByName(MASTER_SHEET);
    if (!master) throw new Error('Master Database sheet not found.');

    var data = master.getDataRange().getValues();
    if (data.length < 2) throw new Error('Master Database is empty.');

    var headers = data[0];
    var rows = data.slice(1);

    var idxBiomed = findHeaderIndex_(headers, ['QR Code ID', 'Biomedical Number', 'BIOMEDICAL NUMBER', 'Biomed Number']);
    var idxStatus = findHeaderIndex_(headers, ['Status', 'STATUS']);
    var idxSeen = findHeaderIndex_(headers, ['SEEN THIS ROUND', 'Seen This Round']);
    var idxInventoryDone = findHeaderIndex_(headers, ['INVENTORY DONE', 'Inventory Done']);

    var now = new Date();
    var changedIndexes = {};

    for (var i = 0; i < rows.length; i++) {
      var biomed = idxBiomed > -1 ? String(rows[i][idxBiomed] || '').trim() : '';
      if (!biomed) continue;

      var status = idxStatus > -1 ? String(rows[i][idxStatus] || '').trim().toUpperCase() : '';
      var seen = idxSeen > -1 ? String(rows[i][idxSeen] || '').trim().toUpperCase() : '';

      if (seen === 'YES') {
        if (idxStatus > -1 && status !== 'WORKING') {
          rows[i][idxStatus] = 'WORKING';
          changedIndexes[i] = true;
        }
      } else {
        if (status === 'WORKING') {
          if (idxStatus > -1) rows[i][idxStatus] = 'UNACCOUNTED';
          changedIndexes[i] = true;
        }
      }

      if (idxInventoryDone > -1 && rows[i][idxInventoryDone] !== cleanStaff) {
        rows[i][idxInventoryDone] = cleanStaff;
        changedIndexes[i] = true;
      }
    }

    writeBackRowBlocks_(master, headers.length, rows, Object.keys(changedIndexes));

    appendUpdateLog_({
      timestamp: now,
      source: 'INVENTORY_DONE',
      biomed: '',
      equipmentName: '',
      status: '',
      location: '',
      checkedBy: cleanStaff,
      note: 'Inventory round finalized'
    });

    SpreadsheetApp.flush();
    invalidateDashboardCache_();

    var analysis = buildInventoryAnalysis_([headers].concat(rows));

    return {
      staffName: cleanStaff,
      total: analysis.totals.total,
      scannedThisRound: analysis.scannedThisRound,
      accountedCount: analysis.accountedCount,
      completionPercent: analysis.completionPercent,
      alertWorkingCount: analysis.alertWorkingCount,
      unscannedTotal: analysis.unscannedTotal,
      unscannedItems: analysis.unscannedItems,
      unscannedSummary: analysis.unscannedSummary
    };
  } catch (err) {
    logError_(err, 'finalizeInventory');
    throw err;
  }
}

/*************************************************
 * EMAILS / REPORTS
 *************************************************/
function sendMissingAlert() {
  try {
    var analysis = buildInventoryAnalysis_(getMasterData_());
    var payload = buildAlertPdfSheet_('PHDU Missing Equipment Alert', analysis.grouped.missing, '#f4cccc', 'Currently marked missing');

    sendPdfEmail_({
      subject: 'PHDU Missing Equipment Alert',
      bodyLines: [
        'PHDU MISSING EQUIPMENT ALERT',
        '',
        'Generated: ' + formatAnyDate_(new Date()),
        'Total Missing: ' + analysis.grouped.missing.length,
        '',
        'Attached: Hospital-style PDF alert'
      ],
      blob: payload.blob
    });

    return 'Missing alert emailed successfully to ' + EMAIL_TO + '.';
  } catch (err) {
    logError_(err, 'sendMissingAlert');
    throw err;
  }
}

function sendWeeklyServiceDueAlert() {
  try {
    var analysis = buildInventoryAnalysis_(getMasterData_());
    var overdue = analysis.serviceDueItems.overdue || [];
    var dueSoon = analysis.serviceDueItems.dueSoon || [];
    var total = overdue.length + dueSoon.length;

    if (!total) {
      MailApp.sendEmail({
        to: EMAIL_TO,
        subject: 'PHDU Weekly Service Due Alert',
        body: [
          'PHDU WEEKLY SERVICE DUE ALERT',
          '',
          'Generated: ' + formatAnyDate_(new Date()),
          'No overdue or due-within-30-days service items found.'
        ].join('\n'),
        name: EMAIL_SENDER_NAME
      });
      return 'Weekly service due alert sent. No due items found.';
    }

    var pdf = buildServiceDuePdf_({
      overdue: overdue,
      dueSoon: dueSoon
    });

    sendPdfEmail_({
      subject: 'PHDU Weekly Service Due Alert',
      bodyLines: [
        'PHDU WEEKLY SERVICE DUE ALERT',
        '',
        'Generated: ' + formatAnyDate_(new Date()),
        'Overdue: ' + overdue.length,
        'Due Within 30 Days: ' + dueSoon.length,
        '',
        'Attached: Hospital-style PDF service due alert'
      ],
      blob: pdf.blob
    });

    return 'Weekly service due alert emailed successfully to ' + EMAIL_TO + '.';
  } catch (err) {
    logError_(err, 'sendWeeklyServiceDueAlert');
    throw err;
  }
}

function sendHospitalPdfReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = null;

  try {
    var master = ss.getSheetByName(MASTER_SHEET);
    if (!master) throw new Error('Master Database sheet not found.');

    var masterData = master.getDataRange().getValues();
    if (masterData.length < 2) throw new Error('Master Database is empty.');

    var analysis = buildInventoryAnalysis_(masterData);
    var auditToday = getTodayAuditSummary_();
    var logMeta = getLogMeta_();

    var tempName = 'Hospital_Report_Temp';
    var old = ss.getSheetByName(tempName);
    if (old) ss.deleteSheet(old);

    sh = ss.insertSheet(tempName);
    buildHospitalReportSheet_(sh, {
      analysis: analysis,
      auditToday: auditToday,
      logMeta: logMeta
    });

    SpreadsheetApp.flush();

    var pdfBlob = exportSheetToPdf_(ss.getId(), sh.getSheetId(), 'PHDU_Hospital_Report.pdf');
    var pdfFileInfo = saveReportPdfToDrive_(pdfBlob);

    MailApp.sendEmail({
      to: EMAIL_TO,
      subject: 'PHDU Hospital Equipment Report',
      body: [
        'PHDU HOSPITAL EQUIPMENT REPORT',
        '',
        'Generated: ' + formatAnyDate_(new Date()),
        'Last Checked By: ' + (logMeta.lastCheckedBy || '-'),
        '',
        'Total Equipment: ' + analysis.totals.total,
        'Working: ' + analysis.totals.working,
        'Missing: ' + analysis.totals.missing,
        'Loaned: ' + analysis.totals.loaned,
        'Under Maintenance: ' + analysis.totals.underMaintenance,
        'Condemned: ' + analysis.totals.condemned,
        'Unaccounted: ' + analysis.totals.unaccounted,
        '',
        'Scanned This Round: ' + analysis.scannedThisRound,
        'Accounted This Round: ' + analysis.accountedCount + ' / ' + analysis.totals.total,
        'Completion: ' + analysis.completionPercent + '%',
        '',
        'Today Scan Audit Summary',
        'Matched: ' + auditToday.matched,
        'Duplicates Ignored: ' + auditToday.duplicates,
        'Not Found: ' + auditToday.notFound,
        'Ignored Form QR: ' + auditToday.ignoredUrls,
        '',
        'Service Alerts',
        'Overdue Service: ' + analysis.serviceDueSummary.overdue,
        'Due Within 30 Days: ' + analysis.serviceDueSummary.dueSoon,
        '',
        'Attached: Hospital-style PDF report'
      ].join('\n'),
      attachments: [pdfBlob],
      name: EMAIL_SENDER_NAME
    });

    archiveReportRow_(analysis, pdfFileInfo);
    upsertMonthlySummary_(analysis);

    return 'Hospital report emailed successfully to ' + EMAIL_TO + '.';
  } catch (err) {
    logError_(err, 'sendHospitalPdfReport');
    throw err;
  } finally {
    try {
      if (sh && ss.getSheetByName(sh.getName())) ss.deleteSheet(sh);
    } catch (ignore) {}
  }
}

function monthlyAutoSendReport() {
  return sendHospitalPdfReport();
}

function createMonthlyReportTrigger() {
  deleteMonthlyReportTriggers();

  ScriptApp.newTrigger('monthlyAutoSendReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(7)
    .create();

  return 'Monthly report trigger created.';
}

function deleteMonthlyReportTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'monthlyAutoSendReport') {
      ScriptApp.deleteTrigger(t);
    }
  });
  return 'Monthly report triggers removed.';
}

function createWeeklyServiceDueTrigger() {
  deleteWeeklyServiceDueTriggers();

  ScriptApp.newTrigger('sendWeeklyServiceDueAlert')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(7)
    .create();

  return 'Weekly service due alert trigger created.';
}

function deleteWeeklyServiceDueTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'sendWeeklyServiceDueAlert') {
      ScriptApp.deleteTrigger(t);
    }
  });
  return 'Weekly service due alert triggers removed.';
}

function monthlyArchiveAndClean() {
  try {
    sendHospitalPdfReport();
    archiveAndCleanLiveSheets_();
    return 'Monthly archive and cleanup completed.';
  } catch (err) {
    logError_(err, 'monthlyArchiveAndClean');
    throw err;
  }
}

function createMonthlyMaintenanceTrigger() {
  deleteMonthlyMaintenanceTriggers();

  ScriptApp.newTrigger('monthlyArchiveAndClean')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();

  return 'Monthly maintenance trigger created.';
}

function deleteMonthlyMaintenanceTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'monthlyArchiveAndClean') {
      ScriptApp.deleteTrigger(t);
    }
  });
  return 'Monthly maintenance triggers removed.';
}

/*************************************************
 * REPORT BUILDERS
 *************************************************/
function buildHospitalReportSheet_(sheet, report) {
  sheet.clear();

  var analysis = report.analysis;
  var row = 1;

  sheet.getRange(row, 1, 1, 8).merge();
  sheet.getRange(row, 1).setValue('PHDU EQUIPMENT INVENTORY REPORT');
  sheet.getRange(row, 1).setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  row++;

  sheet.getRange(row, 1, 1, 8).merge();
  sheet.getRange(row, 1).setValue('Royal Hospital - Pediatric High Dependency Unit');
  sheet.getRange(row, 1).setFontSize(11).setHorizontalAlignment('center');
  row++;

  sheet.getRange(row, 1, 1, 8).merge();
  sheet.getRange(row, 1).setValue('Generated: ' + formatAnyDate_(new Date()));
  sheet.getRange(row, 1).setFontSize(10).setHorizontalAlignment('center');
  row++;

  sheet.getRange(row, 1, 1, 8).merge();
  sheet.getRange(row, 1).setValue('Last Checked By: ' + (report.logMeta.lastCheckedBy || '-'));
  sheet.getRange(row, 1).setFontSize(10).setHorizontalAlignment('center');
  row += 2;

  sheet.getRange(row, 1, 1, 8).merge();
  sheet.getRange(row, 1).setValue('SUMMARY');
  sheet.getRange(row, 1).setFontWeight('bold').setBackground('#d9eaf7');
  row++;

  var summaryRows = [
    ['Total Equipment', analysis.totals.total],
    ['Working', analysis.totals.working],
    ['Missing', analysis.totals.missing],
    ['Loaned', analysis.totals.loaned],
    ['Under Maintenance', analysis.totals.underMaintenance],
    ['Condemned', analysis.totals.condemned],
    ['Unaccounted / Alert', analysis.totals.unaccounted],
    ['Scanned This Round', analysis.scannedThisRound],
    ['Accounted This Round', analysis.accountedCount + ' / ' + analysis.totals.total],
    ['Completion %', analysis.completionPercent + '%'],
    ['Overdue Service', analysis.serviceDueSummary.overdue],
    ['Due Within 30 Days', analysis.serviceDueSummary.dueSoon]
  ];

  sheet.getRange(row, 1, summaryRows.length, 2).setValues(summaryRows);
  sheet.getRange(row, 1, summaryRows.length, 2).setBorder(true, true, true, true, true, true);
  sheet.getRange(row, 1, summaryRows.length, 1).setFontWeight('bold');
  row += summaryRows.length + 2;

  sheet.getRange(row, 1, 1, 8).merge();
  sheet.getRange(row, 1).setValue('TODAY SCAN AUDIT');
  sheet.getRange(row, 1).setFontWeight('bold').setBackground('#fce5cd');
  row++;

  var auditRows = [
    ['Matched scans', report.auditToday.matched],
    ['Duplicates Ignored', report.auditToday.duplicates],
    ['Not Found', report.auditToday.notFound],
    ['Ignored Form QR', report.auditToday.ignoredUrls]
  ];

  sheet.getRange(row, 1, auditRows.length, 2).setValues(auditRows);
  sheet.getRange(row, 1, auditRows.length, 2).setBorder(true, true, true, true, true, true);
  sheet.getRange(row, 1, auditRows.length, 1).setFontWeight('bold');
  row += auditRows.length + 2;

  row = writeSection_(sheet, row, 'ALERT ITEMS (PREVIOUSLY WORKING / PREVIOUSLY SCANNED BUT NOT SCANNED NOW)', analysis.grouped.unaccounted, '#fde9e9');
  row = writeSection_(sheet, row, 'NEXT SERVICE DUE - OVERDUE', analysis.serviceDueItems.overdue, '#ffe6e6');
  row = writeSection_(sheet, row, 'NEXT SERVICE DUE - WITHIN 30 DAYS', analysis.serviceDueItems.dueSoon, '#fff2cc');
  row = writeSection_(sheet, row, 'WORKING EQUIPMENT', analysis.grouped.working, '#d9ead3');
  row = writeSection_(sheet, row, 'MISSING EQUIPMENT', analysis.grouped.missing, '#f4cccc');
  row = writeSection_(sheet, row, 'LOANED EQUIPMENT', analysis.grouped.loaned, '#cfe2f3');
  row = writeSection_(sheet, row, 'UNDER MAINTENANCE', analysis.grouped.maintenance, '#fce5cd');
  row = writeSection_(sheet, row, 'CONDEMNED EQUIPMENT', analysis.grouped.condemned, '#eeeeee');

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 230);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 220);
  sheet.setColumnWidth(8, 240);
  sheet.getRange(1, 1, sheet.getLastRow(), 8).setWrap(true).setVerticalAlignment('middle');
  sheet.setFrozenRows(5);
}

function writeSection_(sheet, startRow, title, items, color) {
  var row = startRow;

  sheet.getRange(row, 1, 1, 8).merge();
  sheet.getRange(row, 1).setValue(title);
  sheet.getRange(row, 1).setFontWeight('bold').setBackground(color);
  row++;

  sheet.getRange(row, 1, 1, 8).setValues([[
    'Biomedical No',
    'Equipment Name',
    'Status',
    'Location',
    'Last Service Date',
    'Next Service Due',
    'Remarks',
    'Note'
  ]]);
  sheet.getRange(row, 1, 1, 8).setFontWeight('bold').setBackground('#d9eaf7');
  sheet.getRange(row, 1, 1, 8).setBorder(true, true, true, true, true, true);
  row++;

  if (!items.length) {
    sheet.getRange(row, 1, 1, 8).merge();
    sheet.getRange(row, 1).setValue('None');
    sheet.getRange(row, 1).setBorder(true, true, true, true, true, true);
    row += 2;
    return row;
  }

  var values = items.map(function(item) {
    return [
      item.biomed || '-',
      item.name || '-',
      item.status || '-',
      item.location || '-',
      item.lastServiceDate || '-',
      item.nextServiceDate || '-',
      item.remarks || '-',
      item.unscannedNote || item.serviceNote || '-'
    ];
  });

  sheet.getRange(row, 1, values.length, 8).setValues(values);
  sheet.getRange(row, 1, values.length, 8).setBorder(true, true, true, true, true, true);
  row += values.length + 2;

  return row;
}

function buildAlertPdfSheet_(title, items, color, defaultNote) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = null;

  try {
    var tempName = 'Alert_Report_Temp';
    var old = ss.getSheetByName(tempName);
    if (old) ss.deleteSheet(old);

    sh = ss.insertSheet(tempName);
    var row = 1;

    sh.getRange(row, 1, 1, 8).merge();
    sh.getRange(row, 1).setValue(title);
    sh.getRange(row, 1).setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
    row++;

    sh.getRange(row, 1, 1, 8).merge();
    sh.getRange(row, 1).setValue('Royal Hospital - Pediatric High Dependency Unit');
    sh.getRange(row, 1).setHorizontalAlignment('center');
    row++;

    sh.getRange(row, 1, 1, 8).merge();
    sh.getRange(row, 1).setValue('Generated: ' + formatAnyDate_(new Date()));
    sh.getRange(row, 1).setHorizontalAlignment('center');
    row += 2;

    sh.getRange(row, 1, 1, 8).merge();
    sh.getRange(row, 1).setValue('DETAILS');
    sh.getRange(row, 1).setFontWeight('bold').setBackground(color);
    row++;

    sh.getRange(row, 1, 1, 8).setValues([[
      'Biomedical No',
      'Equipment Name',
      'Status',
      'Location',
      'Last Service Date',
      'Next Service Due',
      'Remarks',
      'Note'
    ]]);
    sh.getRange(row, 1, 1, 8).setFontWeight('bold').setBackground('#d9eaf7');
    row++;

    if (!items.length) {
      sh.getRange(row, 1, 1, 8).merge();
      sh.getRange(row, 1).setValue('None');
    } else {
      var values = items.map(function(item) {
        return [
          item.biomed || '-',
          item.name || '-',
          item.status || '-',
          item.location || '-',
          item.lastServiceDate || '-',
          item.nextServiceDate || '-',
          item.remarks || '-',
          item.unscannedNote || item.serviceNote || defaultNote || '-'
        ];
      });
      sh.getRange(row, 1, values.length, 8).setValues(values);
    }

    SpreadsheetApp.flush();
    var blob = exportSheetToPdf_(ss.getId(), sh.getSheetId(), title.replace(/[^A-Za-z0-9]+/g, '_') + '.pdf');
    return { blob: blob };
  } finally {
    try {
      if (sh && ss.getSheetByName(sh.getName())) ss.deleteSheet(sh);
    } catch (ignore) {}
  }
}

function buildServiceDuePdf_(payload) {
  var overdue = payload.overdue || [];
  var dueSoon = payload.dueSoon || [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = null;

  try {
    var tempName = 'Service_Due_Report_Temp';
    var old = ss.getSheetByName(tempName);
    if (old) ss.deleteSheet(old);

    sh = ss.insertSheet(tempName);
    var row = 1;

    sh.getRange(row, 1, 1, 7).merge();
    sh.getRange(row, 1).setValue('PHDU SERVICE DUE ALERT');
    sh.getRange(row, 1).setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
    row++;

    sh.getRange(row, 1, 1, 7).merge();
    sh.getRange(row, 1).setValue('Royal Hospital - Pediatric High Dependency Unit');
    sh.getRange(row, 1).setHorizontalAlignment('center');
    row++;

    sh.getRange(row, 1, 1, 7).merge();
    sh.getRange(row, 1).setValue('Generated: ' + formatAnyDate_(new Date()));
    sh.getRange(row, 1).setHorizontalAlignment('center');
    row += 2;

    row = writeServiceSection_(sh, row, 'OVERDUE SERVICE', overdue, '#ffe6e6');
    row = writeServiceSection_(sh, row, 'DUE WITHIN 30 DAYS', dueSoon, '#fff2cc');

    SpreadsheetApp.flush();
    var blob = exportSheetToPdf_(ss.getId(), sh.getSheetId(), 'PHDU_Service_Due_Alert.pdf');
    return { blob: blob };
  } finally {
    try {
      if (sh && ss.getSheetByName(sh.getName())) ss.deleteSheet(sh);
    } catch (ignore) {}
  }
}

function writeServiceSection_(sheet, startRow, title, items, color) {
  var row = startRow;

  sheet.getRange(row, 1, 1, 7).merge();
  sheet.getRange(row, 1).setValue(title);
  sheet.getRange(row, 1).setFontWeight('bold').setBackground(color);
  row++;

  sheet.getRange(row, 1, 1, 7).setValues([[
    'Biomedical No',
    'Equipment Name',
    'Status',
    'Location',
    'Last Service',
    'Next Service Due',
    'Remarks'
  ]]);
  sheet.getRange(row, 1, 1, 7).setFontWeight('bold').setBackground('#d9eaf7');
  row++;

  if (!items.length) {
    sheet.getRange(row, 1, 1, 7).merge();
    sheet.getRange(row, 1).setValue('None');
    row += 2;
    return row;
  }

  var values = items.map(function(item) {
    return [
      item.biomed || '-',
      item.name || '-',
      item.status || '-',
      item.location || '-',
      item.lastServiceDate || '-',
      item.nextServiceDate || '-',
      item.remarks || '-'
    ];
  });

  sheet.getRange(row, 1, values.length, 7).setValues(values);
  row += values.length + 2;
  return row;
}

function sendPdfEmail_(opts) {
  MailApp.sendEmail({
    to: EMAIL_TO,
    subject: opts.subject,
    body: (opts.bodyLines || []).join('\n'),
    attachments: opts.blob ? [opts.blob] : [],
    name: EMAIL_SENDER_NAME
  });
}

/*************************************************
 * ANALYSIS
 *************************************************/
function buildInventoryAnalysis_(masterData) {
  if (!masterData || masterData.length < 2) {
    return {
      inventoryComplete: false,
      completionPercent: '0.0',
      serviceAlertText: 'No service alerts',
      scannedThisRound: 0,
      accountedCount: 0,
      alertWorkingCount: 0,
      unscannedTotal: 0,
      totals: {
        total: 0,
        working: 0,
        missing: 0,
        loaned: 0,
        underMaintenance: 0,
        condemned: 0,
        unaccounted: 0
      },
      grouped: {
        working: [],
        missing: [],
        loaned: [],
        maintenance: [],
        condemned: [],
        unaccounted: []
      },
      unscannedGrouped: {
        unaccounted: [],
        missing: [],
        loaned: [],
        maintenance: []
      },
      unscannedSummary: {
        unaccounted: 0,
        missing: 0,
        loaned: 0,
        maintenance: 0
      },
      serviceDueItems: {
        overdue: [],
        dueSoon: []
      },
      serviceDueSummary: {
        overdue: 0,
        dueSoon: 0
      },
      equipment: [],
      unscannedItems: [],
      alertItems: []
    };
  }

  var headers = masterData[0];
  var rows = masterData.slice(1);

  var idxBiomed = findHeaderIndex_(headers, ['QR Code ID', 'Biomedical Number', 'BIOMEDICAL NUMBER', 'Biomed Number']);
  var idxName = findHeaderIndex_(headers, ['Equipment Name', 'EQUIPMENT NAME']);
  var idxStatus = findHeaderIndex_(headers, ['Status', 'STATUS']);
  var idxLocation = findHeaderIndex_(headers, ['Location/Ward', 'LOCATION/WARD', 'Location', 'LOCATION']);
  var idxRemarks = findHeaderIndex_(headers, ['Remarks', 'REMARKS']);
  var idxSeen = findHeaderIndex_(headers, ['SEEN THIS ROUND', 'Seen This Round']);
  var idxInventoryDone = findHeaderIndex_(headers, ['INVENTORY DONE', 'Inventory Done']);
  var idxLastScannedDate = findHeaderIndex_(headers, ['LAST SCANNED DATE', 'Last Scanned Date']);
  var idxLastScannedBy = findHeaderIndex_(headers, ['LAST SCANNED BY', 'Last Scanned By']);
  var idxLastServiceDate = findHeaderIndex_(headers, ['LAST SERVICE DATE', 'Last Service Date']);
  var idxNextServiceDate = findHeaderIndex_(headers, ['NEXT SERVICE DUE', 'Next Service Due', 'NEXT SERVICE DATE', 'Next Service Date', 'SERVICE DUE']);

  var totals = {
    total: 0,
    working: 0,
    missing: 0,
    loaned: 0,
    underMaintenance: 0,
    condemned: 0,
    unaccounted: 0
  };

  var grouped = {
    working: [],
    missing: [],
    loaned: [],
    maintenance: [],
    condemned: [],
    unaccounted: []
  };

  var unscannedGrouped = {
    unaccounted: [],
    missing: [],
    loaned: [],
    maintenance: []
  };

  var serviceDueItems = {
    overdue: [],
    dueSoon: []
  };

  var equipment = [];
  var unscannedItems = [];
  var alertItems = [];
  var scannedThisRound = 0;
  var accountedCount = 0;
  var overdueCount = 0;
  var dueSoonCount = 0;

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var next30 = new Date(today);
  next30.setDate(next30.getDate() + SERVICE_DUE_DAYS);

  for (var i = 0; i < rows.length; i++) {
    var biomed = idxBiomed > -1 ? String(rows[i][idxBiomed] || '').trim() : '';
    var name = idxName > -1 ? String(rows[i][idxName] || '').trim() : '';
    var status = idxStatus > -1 ? String(rows[i][idxStatus] || '').trim() : '';
    var location = idxLocation > -1 ? String(rows[i][idxLocation] || '').trim() : '';
    var remarks = idxRemarks > -1 ? String(rows[i][idxRemarks] || '').trim() : '';
    var seen = idxSeen > -1 ? String(rows[i][idxSeen] || '').trim().toUpperCase() : '';
    var inventoryDoneBy = idxInventoryDone > -1 ? String(rows[i][idxInventoryDone] || '').trim() : '';
    var lastScannedBy = idxLastScannedBy > -1 ? String(rows[i][idxLastScannedBy] || '').trim() : '';
    var lastScannedDate = idxLastScannedDate > -1 ? rows[i][idxLastScannedDate] : '';
    var lastServiceDate = idxLastServiceDate > -1 ? rows[i][idxLastServiceDate] : '';
    var nextServiceDate = idxNextServiceDate > -1 ? rows[i][idxNextServiceDate] : '';

    if (!biomed && !name && !status) continue;

    totals.total++;

    var statusUpper = String(status || '').trim().toUpperCase();

    var item = {
      biomed: biomed || '-',
      name: name || '-',
      status: status || '-',
      location: location || '-',
      remarks: remarks || '-',
      seenThisRound: seen === 'YES' ? 'YES' : '',
      inventoryDoneBy: inventoryDoneBy || '',
      lastScannedBy: lastScannedBy || '',
      displayCheckedBy: lastScannedBy || inventoryDoneBy || '-',
      lastScannedDate: formatAnyDate_(lastScannedDate),
      lastServiceDate: formatDateOnly_(lastServiceDate),
      nextServiceDate: formatDateOnly_(nextServiceDate),
      unscannedNote: '',
      serviceNote: ''
    };

    if (statusUpper === 'WORKING') {
      totals.working++;
      grouped.working.push(item);
    } else if (statusUpper === 'MISSING') {
      totals.missing++;
      grouped.missing.push(item);
    } else if (statusUpper === 'LOANED') {
      totals.loaned++;
      grouped.loaned.push(item);
    } else if (statusUpper === 'UNDER MAINTENANCE' || statusUpper === 'MAINTENANCE') {
      totals.underMaintenance++;
      grouped.maintenance.push(item);
    } else if (statusUpper === 'CONDEMNED') {
      totals.condemned++;
      grouped.condemned.push(item);
    } else if (statusUpper === 'UNACCOUNTED') {
      totals.unaccounted++;
      grouped.unaccounted.push(item);
    }

    if (seen === 'YES') scannedThisRound++;

    var accounted = (
      statusUpper === 'WORKING' ||
      statusUpper === 'MISSING' ||
      statusUpper === 'LOANED' ||
      statusUpper === 'UNDER MAINTENANCE' ||
      statusUpper === 'MAINTENANCE' ||
      statusUpper === 'CONDEMNED'
    );

    if (accounted) accountedCount++;

    if (idxNextServiceDate > -1) {
      var dueDate = parseSheetDate_(nextServiceDate);
      if (dueDate) {
        dueDate.setHours(0, 0, 0, 0);
        if (dueDate < today) {
          overdueCount++;
          item.serviceNote = 'OVERDUE';
          serviceDueItems.overdue.push(cloneItem_(item));
        } else if (dueDate <= next30) {
          dueSoonCount++;
          item.serviceNote = 'DUE WITHIN 30 DAYS';
          serviceDueItems.dueSoon.push(cloneItem_(item));
        }
      }
    }

    if (statusUpper === 'UNACCOUNTED') {
      item.unscannedNote = 'Previously scanned / previously working but not scanned now';
      unscannedGrouped.unaccounted.push(item);
      unscannedItems.push(item);
      alertItems.push(item);
    } else if (statusUpper === 'MISSING') {
      item.unscannedNote = 'Currently marked MISSING';
      unscannedGrouped.missing.push(item);
      unscannedItems.push(item);
    } else if (statusUpper === 'LOANED') {
      item.unscannedNote = 'Currently marked LOANED';
      unscannedGrouped.loaned.push(item);
      unscannedItems.push(item);
    } else if (statusUpper === 'UNDER MAINTENANCE' || statusUpper === 'MAINTENANCE') {
      item.unscannedNote = 'Currently marked UNDER MAINTENANCE';
      unscannedGrouped.maintenance.push(item);
      unscannedItems.push(item);
    }

    equipment.push(item);
  }

  var orderedUnscanned = []
    .concat(unscannedGrouped.unaccounted)
    .concat(unscannedGrouped.missing)
    .concat(unscannedGrouped.loaned)
    .concat(unscannedGrouped.maintenance);

  var completionPercent = totals.total ? ((accountedCount / totals.total) * 100).toFixed(1) : '0.0';
  var inventoryComplete = totals.total > 0 && totals.unaccounted === 0;

  var serviceAlertText = 'No service alerts';
  if (overdueCount > 0) {
    serviceAlertText = overdueCount + ' overdue service item(s)';
  } else if (dueSoonCount > 0) {
    serviceAlertText = dueSoonCount + ' service item(s) due within 30 days';
  }

  return {
    inventoryComplete: inventoryComplete,
    completionPercent: completionPercent,
    serviceAlertText: serviceAlertText,
    scannedThisRound: scannedThisRound,
    accountedCount: accountedCount,
    alertWorkingCount: alertItems.length,
    unscannedTotal: orderedUnscanned.length,
    totals: totals,
    grouped: grouped,
    unscannedGrouped: unscannedGrouped,
    unscannedSummary: {
      unaccounted: unscannedGrouped.unaccounted.length,
      missing: unscannedGrouped.missing.length,
      loaned: unscannedGrouped.loaned.length,
      maintenance: unscannedGrouped.maintenance.length
    },
    serviceDueItems: serviceDueItems,
    serviceDueSummary: {
      overdue: serviceDueItems.overdue.length,
      dueSoon: serviceDueItems.dueSoon.length
    },
    equipment: equipment,
    unscannedItems: orderedUnscanned,
    alertItems: alertItems
  };
}

function cloneItem_(item) {
  return JSON.parse(JSON.stringify(item));
}

/*************************************************
 * MASTER UPDATE HELPER
 *************************************************/
function updateMasterByBiomed_(master, payload) {
  var data = master.getDataRange().getValues();
  if (data.length < 2) throw new Error('Master Database is empty.');

  var headers = data[0];
  var rows = data.slice(1);

  var idxBiomed = findHeaderIndex_(headers, ['QR Code ID', 'Biomedical Number', 'BIOMEDICAL NUMBER', 'Biomed Number']);
  var idxName = findHeaderIndex_(headers, ['Equipment Name', 'EQUIPMENT NAME']);
  var idxStatus = findHeaderIndex_(headers, ['Status', 'STATUS']);
  var idxLocation = findHeaderIndex_(headers, ['Location/Ward', 'LOCATION/WARD', 'Location', 'LOCATION']);
  var idxRemarks = findHeaderIndex_(headers, ['Remarks', 'REMARKS']);
  var idxLastScannedDate = findHeaderIndex_(headers, ['LAST SCANNED DATE', 'Last Scanned Date']);
  var idxLastScannedBy = findHeaderIndex_(headers, ['LAST SCANNED BY', 'Last Scanned By']);
  var idxSeen = findHeaderIndex_(headers, ['SEEN THIS ROUND', 'Seen This Round']);
  var idxInventoryDone = findHeaderIndex_(headers, ['INVENTORY DONE', 'Inventory Done']);
  var idxLastServiceDate = findHeaderIndex_(headers, ['LAST SERVICE DATE', 'Last Service Date']);
  var idxNextServiceDue = findHeaderIndex_(headers, ['NEXT SERVICE DUE', 'Next Service Due', 'NEXT SERVICE DATE', 'Next Service Date', 'SERVICE DUE']);
  var idxMissingAlert = findHeaderIndex_(headers, ['MISSING ALERT']);

  if (idxBiomed === -1) throw new Error('Biomedical Number / QR Code ID column not found in Master Database.');

  var target = normalizeScanCode_(payload.biomed);
  var foundRowIndex = -1;
  var now = new Date();

  for (var i = 0; i < rows.length; i++) {
    var rowCode = normalizeScanCode_(rows[i][idxBiomed]);
    if (rowCode !== target) continue;

    foundRowIndex = i;

    if (idxName > -1 && payload.equipmentName) rows[i][idxName] = payload.equipmentName;
    if (idxStatus > -1 && payload.status) rows[i][idxStatus] = payload.status;
    if (idxLocation > -1 && payload.location) rows[i][idxLocation] = payload.location;
    if (idxRemarks > -1 && payload.remarks !== undefined && payload.remarks !== '') rows[i][idxRemarks] = payload.remarks;

    if (payload.updateLastScanned) {
      if (idxLastScannedDate > -1) rows[i][idxLastScannedDate] = now;
      if (idxLastScannedBy > -1) rows[i][idxLastScannedBy] = payload.inventoryDoneBy || payload.checkedBy || 'Unknown';
    }

    if (idxSeen > -1) {
      if (payload.markSeen) rows[i][idxSeen] = 'YES';
      if (payload.clearSeen) rows[i][idxSeen] = '';
    }

    if (idxInventoryDone > -1 && (payload.inventoryDoneBy || payload.checkedBy)) {
      rows[i][idxInventoryDone] = payload.inventoryDoneBy || payload.checkedBy;
    }

    if (idxLastServiceDate > -1 && payload.lastServiceDate) {
      var parsedLastService = parseSheetDate_(payload.lastServiceDate);
      if (parsedLastService) rows[i][idxLastServiceDate] = parsedLastService;
    }

    if (idxNextServiceDue > -1) {
      if (payload.nextServiceDue) {
        var parsedNextDue = parseSheetDate_(payload.nextServiceDue);
        if (parsedNextDue) rows[i][idxNextServiceDue] = parsedNextDue;
      } else if (payload.lastServiceDate) {
        var fromLast = parseSheetDate_(payload.lastServiceDate);
        if (fromLast) rows[i][idxNextServiceDue] = addMonths_(fromLast, 11);
      }
    }

    if (idxMissingAlert > -1 && payload.status) {
      var s = String(payload.status).trim().toUpperCase();
      if (s === 'MISSING') rows[i][idxMissingAlert] = 'SENT';
      if (s === 'WORKING') rows[i][idxMissingAlert] = '';
    }

    break;
  }

  if (foundRowIndex === -1) {
    throw new Error('Biomedical Number not found in Master Database: ' + payload.biomed);
  }

  writeBackRowBlocks_(master, headers.length, rows, [String(foundRowIndex)]);
}

/*************************************************
 * LOG HELPERS
 *************************************************/
function appendUpdateLog_(data) {
  appendUpdateLogRows_([[
    data.timestamp || new Date(),
    data.source || '',
    data.biomed || '',
    data.equipmentName || '',
    data.status || '',
    data.location || '',
    String(data.checkedBy || '').trim() || 'Unknown',
    data.note || ''
  ]]);
}

function appendUpdateLogRows_(rows) {
  if (!rows || !rows.length) return;

  var cleaned = dedupeRowsByKey_(rows, function(r) {
    return [
      formatKeyDate_(r[0]),
      String(r[1] || ''),
      String(r[2] || ''),
      String(r[3] || ''),
      String(r[4] || ''),
      String(r[5] || ''),
      String(r[6] || ''),
      String(r[7] || '')
    ].join('|');
  });

  if (!cleaned.length) return;

  var sh = getOrCreateUpdateLogSheet_();
  appendRowsToSheet_(sh, cleaned);
}

function getLogMeta_() {
  var log = getOrCreateUpdateLogSheet_();
  var lastRow = log.getLastRow();

  if (lastRow < 2) {
    return {
      lastCheckedBy: '-',
      lastUpdate: '-',
      todayLogs: 0,
      weekLogs: 0,
      totalLogs: 0
    };
  }

var totalLogs = lastRow - 1;
var readRows = Math.min(120, totalLogs);
var startRow = lastRow - readRows + 1;

var data = log.getRange(startRow, 1, readRows, 8).getValues();

var now = new Date();
var todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
var weekStart = new Date(todayStart);
weekStart.setDate(weekStart.getDate() - 6);

var todayLogs = 0;
var weekLogs = 0;
var last = data[data.length - 1];
  for (var i = 0; i < data.length; i++) {
    var ts = parseSheetDate_(data[i][0]);
    if (!ts) continue;
    if (ts >= todayStart) todayLogs++;
    if (ts >= weekStart) weekLogs++;
  }

  return {
    lastCheckedBy: String(last[6] || '-') || '-',
    lastUpdate: formatAnyDate_(last[0]),
    todayLogs: todayLogs,
    weekLogs: weekLogs,
    totalLogs: totalLogs
  };
}

function getRecentUpdates_() {
  var log = getOrCreateUpdateLogSheet_();
  var lastRow = log.getLastRow();

  if (lastRow < 2) return [];

  var readRows = Math.min(30, lastRow - 1);
  var startRow = lastRow - readRows + 1;
  var rows = log.getRange(startRow, 1, readRows, 8).getValues().reverse();
  var out = [];
  var seen = {};

  for (var i = 0; i < rows.length && out.length < 12; i++) {
    var r = rows[i];

    var timestamp = r[0];
    var source = String(r[1] || '').trim();
    var biomed = String(r[2] || '').trim();
    var equipmentName = String(r[3] || '').trim();
    var status = String(r[4] || '').trim();
    var location = String(r[5] || '').trim();
    var checkedBy = String(r[6] || '').trim() || 'Unknown';
    var note = String(r[7] || '').trim();

    var text = (equipmentName || 'Inventory Update');
    if (biomed) text += ' (' + biomed + ')';
    if (status) text += ' - ' + status;
    text += ' - by ' + checkedBy;
    if (source) text += ' - ' + source;
    if (location) text += ' - ' + location;
    if (note) text += ' - ' + note;

    var timeText = formatAnyDate_(timestamp);
    var key = text + '|' + timeText;
    if (seen[key]) continue;
    seen[key] = true;

    out.push({
      text: text,
      time: timeText
    });
  }

  return out;
}

/*************************************************
 * AUDIT / SHEET HELPERS
 *************************************************/
function getTodayAuditSummary_() {
  var sh = getOrCreateScanAuditSheet_();
  var lastRow = sh.getLastRow();

  if (lastRow < 2) {
    return { matched: 0, duplicates: 0, notFound: 0, ignoredUrls: 0 };
  }

  var data = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  var today = new Date();
  var dayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  var matched = 0;
  var duplicates = 0;
  var notFound = 0;
  var ignoredUrls = 0;

  for (var i = 0; i < data.length; i++) {
    var ts = parseSheetDate_(data[i][0]);
    if (!ts || ts < dayStart) continue;

    var result = String(data[i][6] || '');
    if (result === 'MATCHED') matched++;
    else if (result === 'DUPLICATE_IGNORED') duplicates++;
    else if (result === 'NOT_FOUND') notFound++;
    else if (result === 'IGNORED_URL') ignoredUrls++;
  }

  return {
    matched: matched,
    duplicates: duplicates,
    notFound: notFound,
    ignoredUrls: ignoredUrls
  };
}

function getOrCreateUpdateLogSheet_() {
  return getOrCreateSheetWithHeaders_(UPDATE_LOG_SHEET, [[
    'Timestamp',
    'Source',
    'Biomedical Number',
    'Equipment Name',
    'Status',
    'Location',
    'Checked By',
    'Note'
  ]]);
}

function getOrCreateFastScanSheet_() {
  return getOrCreateSheetWithHeaders_(FAST_SCAN_SHEET, [[
    'Session ID',
    'Staff Name',
    'Start Time',
    'End Time',
    'Session Status',
    'Scanned Codes',
    'Duplicate Codes',
    'Not Found Codes',
    'Unique Count',
    'Duplicate Count',
    'Not Found Count',
    'Inventory Round',
    'Last Saved',
    'Last Activity',
    'Revision No'
  ]]);
}

function getOrCreateScanAuditSheet_() {
  return getOrCreateSheetWithHeaders_(SCAN_AUDIT_SHEET, [[
    'Timestamp',
    'Session ID',
    'Staff Name',
    'Scan Type',
    'Raw Code',
    'Normalized Code',
    'Match Result',
    'Biomedical Number',
    'Equipment Name',
    'Status',
    'Location',
    'Note'
  ]]);
}

function getOrCreateSheetWithHeaders_(name, headerValues) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, headerValues.length, headerValues[0].length).setValues(headerValues);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, headerValues.length, headerValues[0].length).setValues(headerValues);
  }

  return sh;
}

function findSessionRowInfo_(sh, sessionId) {
  var data = sh.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '') === String(sessionId)) {
      return {
        row: i + 1,
        sessionId: data[i][0],
        staffName: data[i][1],
        startTime: data[i][2],
        endTime: data[i][3],
        status: data[i][4],
        scannedCodes: data[i][5],
        duplicateCodes: data[i][6],
        notFoundCodes: data[i][7],
        uniqueCount: data[i][8],
        duplicateCount: data[i][9],
        notFoundCount: data[i][10],
        inventoryRound: data[i][11],
        lastSaved: data[i][12],
        lastActivity: data[i][13],
        revisionNo: data[i][14]
      };
    }
  }

  return null;
}

function findEditableSessionByStaff_(sh, staffName, hours) {
  var data = sh.getDataRange().getValues();
  var now = new Date();

  for (var i = data.length - 1; i >= 1; i--) {
    var sessionStaff = String(data[i][1] || '').trim();
    if (sessionStaff !== staffName) continue;

    var lastActivity = parseSheetDate_(data[i][13]) || parseSheetDate_(data[i][12]) || parseSheetDate_(data[i][3]) || parseSheetDate_(data[i][2]);
    if (!lastActivity) continue;

    var editableUntil = addHours_(lastActivity, hours);
    if (editableUntil < now) continue;

    return {
      row: i + 1,
      sessionId: data[i][0],
      staffName: data[i][1],
      startTime: data[i][2],
      endTime: data[i][3],
      status: data[i][4],
      scannedCodes: data[i][5],
      duplicateCodes: data[i][6],
      notFoundCodes: data[i][7],
      uniqueCount: data[i][8],
      duplicateCount: data[i][9],
      notFoundCount: data[i][10],
      inventoryRound: data[i][11],
      lastSaved: data[i][12],
      lastActivity: data[i][13],
      revisionNo: data[i][14],
      editableUntil: editableUntil
    };
  }

  return null;
}

/*************************************************
 * ARCHIVE HELPERS
 *************************************************/
function getArchiveSpreadsheet_() {
  return SpreadsheetApp.openByUrl(ARCHIVE_SPREADSHEET_URL);
}

function testArchiveConnection() {
  var archive = getArchiveSpreadsheet_();
  var sheets = archive.getSheets().map(function(s) { return s.getName(); });
  Logger.log('Archive sheets: ' + sheets.join(', '));
}

function getOrCreateArchiveSheet_(name, headers) {
  var archive = getArchiveSpreadsheet_();
  var sh = archive.getSheetByName(name);

  if (!sh) {
    sh = archive.insertSheet(name);
    if (headers && headers.length) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (sh.getLastRow() === 0 && headers && headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sh;
}

function archiveFastScanRows_(rows) {
  if (!rows || !rows.length) return;

  var sh = getOrCreateArchiveSheet_('Fast Scan Logs', [
    'Timestamp', 'Session ID', 'Biomedical Number', 'Staff Name', 'Inventory Round', 'Result'
  ]);

  appendUniqueRowsToSheet_(sh, rows, function(r) {
    return [
      formatKeyDate_(r[0]),
      String(r[1] || ''),
      String(r[2] || ''),
      String(r[3] || ''),
      String(r[4] || ''),
      String(r[5] || '')
    ].join('|');
  });
}

function archiveUpdateLogRows_(rows, inventoryRound) {
  if (!rows || !rows.length) return;

  var sh = getOrCreateArchiveSheet_('Update Log Archive', [
    'Timestamp', 'Biomedical Number', 'Equipment Name', 'Source', 'Status', 'Updated By', 'Remarks', 'Inventory Round'
  ]);

  var mapped = rows.map(function(r) {
    return [
      r[0] || '',
      r[2] || '',
      r[3] || '',
      r[1] || '',
      r[4] || '',
      r[6] || '',
      r[7] || '',
      inventoryRound || ''
    ];
  });

  appendUniqueRowsToSheet_(sh, mapped, function(r) {
    return [
      formatKeyDate_(r[0]),
      String(r[1] || ''),
      String(r[2] || ''),
      String(r[3] || ''),
      String(r[4] || ''),
      String(r[5] || ''),
      String(r[6] || ''),
      String(r[7] || '')
    ].join('|');
  });
}

function archiveSessionRow_(sessionId) {
  var sh = getOrCreateFastScanSheet_();
  var info = findSessionRowInfo_(sh, sessionId);
  if (!info) return;

  var target = getOrCreateArchiveSheet_('Session Archive', [
    'Session ID', 'Inventory Round', 'Staff Name', 'Start Time', 'End Time',
    'Status', 'Unique Count', 'Duplicate Count', 'Not Found Count', 'Revision No'
  ]);

  appendUniqueRowsToSheet_(target, [[
    info.sessionId || '',
    info.inventoryRound || '',
    info.staffName || '',
    info.startTime || '',
    info.endTime || '',
    info.status || '',
    info.uniqueCount || 0,
    info.duplicateCount || 0,
    info.notFoundCount || 0,
    info.revisionNo || 0
  ]], function(r) {
    return String(r[0] || '') + '|' + String(r[9] || 0);
  });
}

function saveReportPdfToDrive_(pdfBlob) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var file = DriveApp.getFileById(ss.getId());
    var parents = file.getParents();
    var folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    var saved = folder.createFile(pdfBlob);
    return {
      fileName: saved.getName(),
      fileId: saved.getId(),
      fileUrl: saved.getUrl()
    };
  } catch (err) {
    logError_(err, 'saveReportPdfToDrive_');
    return {
      fileName: pdfBlob.getName(),
      fileId: '',
      fileUrl: ''
    };
  }
}

function archiveReportRow_(analysis, pdfInfo) {
  var sh = getOrCreateArchiveSheet_('Reports Archive', [
    'Report ID', 'Inventory Round', 'Generated At', 'Generated By',
    'Total', 'Working', 'Missing', 'Loaned', 'Maintenance', 'Condemned',
    'PDF File Name', 'PDF URL'
  ]);

  var meta = getLogMeta_();

  appendRowsToSheet_(sh, [[
    'RPT-' + new Date().getTime(),
    buildInventoryRound_(),
    new Date(),
    meta.lastCheckedBy || '-',
    analysis.totals.total || 0,
    analysis.totals.working || 0,
    analysis.totals.missing || 0,
    analysis.totals.loaned || 0,
    analysis.totals.underMaintenance || 0,
    analysis.totals.condemned || 0,
    (pdfInfo && pdfInfo.fileName) || '',
    (pdfInfo && pdfInfo.fileUrl) || ''
  ]]);
}

function upsertMonthlySummary_(analysis) {
  var sh = getOrCreateArchiveSheet_('Monthly Summary', [
    'Month', 'Total Equipment', 'Working', 'Missing', 'Loaned', 'Maintenance', 'Condemned', 'Generated At'
  ]);

  var monthKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  var data = sh.getDataRange().getValues();
  var rowIndex = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '') === monthKey) {
      rowIndex = i + 1;
      break;
    }
  }

  var values = [[
    monthKey,
    analysis.totals.total || 0,
    analysis.totals.working || 0,
    analysis.totals.missing || 0,
    analysis.totals.loaned || 0,
    analysis.totals.underMaintenance || 0,
    analysis.totals.condemned || 0,
    new Date()
  ]];

  if (rowIndex > -1) {
    sh.getRange(rowIndex, 1, 1, values[0].length).setValues(values);
  } else {
    appendRowsToSheet_(sh, values);
  }
}

function archiveAndCleanLiveSheets_() {
  archiveAndCleanFormResponses_();
  archiveAndCleanUpdateLog_();
  archiveAndCleanScanAudit_();
  archiveAndCleanFastScanSessions_();
  invalidateDashboardCache_();
}

function archiveAndCleanFormResponses_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FORM_RESPONSES_SHEET);
  if (!sh) return;

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  var keepStart = getCurrentMonthStart_();
  var headers = data[0];
  var rowsToArchive = [];
  var rowIndexesToDelete = [];

  for (var i = 1; i < data.length; i++) {
    var ts = parseSheetDate_(data[i][0]);
    if (!ts) continue;

    if (ts < keepStart) {
      rowsToArchive.push(data[i]);
      rowIndexesToDelete.push(i + 1);
    }
  }

  if (rowsToArchive.length) {
    var target = getOrCreateArchiveSheet_('Form Responses Archive', headers);
    appendUniqueRowsToSheet_(target, rowsToArchive, function(r) {
      return r.map(function(x) { return String(x || ''); }).join('|');
    });
    deleteRowsDescending_(sh, rowIndexesToDelete);
  }
}

function archiveAndCleanUpdateLog_() {
  var sh = getOrCreateUpdateLogSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  var keepStart = getCurrentMonthStart_();
  var rowsToArchive = [];
  var rowIndexesToDelete = [];

  for (var i = 1; i < data.length; i++) {
    var ts = parseSheetDate_(data[i][0]);
    if (!ts) continue;
    if (ts < keepStart) {
      rowsToArchive.push([
        data[i][0],
        data[i][2],
        data[i][3],
        data[i][1],
        data[i][4],
        data[i][6],
        data[i][7],
        ''
      ]);
      rowIndexesToDelete.push(i + 1);
    }
  }

  if (rowsToArchive.length) {
    var target = getOrCreateArchiveSheet_('Update Log Archive', [
      'Timestamp', 'Biomedical Number', 'Equipment Name', 'Source', 'Status', 'Updated By', 'Remarks', 'Inventory Round'
    ]);

    appendUniqueRowsToSheet_(target, rowsToArchive, function(r) {
      return [
        formatKeyDate_(r[0]),
        String(r[1] || ''),
        String(r[2] || ''),
        String(r[3] || ''),
        String(r[4] || ''),
        String(r[5] || ''),
        String(r[6] || ''),
        String(r[7] || '')
      ].join('|');
    });

    deleteRowsDescending_(sh, rowIndexesToDelete);
  }
}

function archiveAndCleanScanAudit_() {
  var sh = getOrCreateScanAuditSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  var keepStart = getCurrentMonthStart_();
  var rowsToArchive = [];
  var rowIndexesToDelete = [];

  for (var i = 1; i < data.length; i++) {
    var ts = parseSheetDate_(data[i][0]);
    if (!ts) continue;
    if (ts < keepStart) {
      rowsToArchive.push([
        data[i][0],
        data[i][1],
        data[i][7] || data[i][5] || '',
        data[i][2],
        '',
        data[i][6]
      ]);
      rowIndexesToDelete.push(i + 1);
    }
  }

  if (rowsToArchive.length) {
    var target = getOrCreateArchiveSheet_('Fast Scan Logs', [
      'Timestamp', 'Session ID', 'Biomedical Number', 'Staff Name', 'Inventory Round', 'Result'
    ]);

    appendUniqueRowsToSheet_(target, rowsToArchive, function(r) {
      return [
        formatKeyDate_(r[0]),
        String(r[1] || ''),
        String(r[2] || ''),
        String(r[3] || ''),
        String(r[4] || ''),
        String(r[5] || '')
      ].join('|');
    });

    deleteRowsDescending_(sh, rowIndexesToDelete);
  }
}

function archiveAndCleanFastScanSessions_() {
  var sh = getOrCreateFastScanSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  var keepStart = getCurrentMonthStart_();
  var rowIndexesToDelete = [];

  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][4] || '').trim().toUpperCase();
    var endTime = parseSheetDate_(data[i][3]);
    var startTime = parseSheetDate_(data[i][2]);
    var refDate = endTime || startTime;

    if (!refDate) continue;
    if (status === 'OPEN') continue;
    if (refDate >= keepStart) continue;

    rowIndexesToDelete.push(i + 1);
  }

  if (rowIndexesToDelete.length) {
    deleteRowsDescending_(sh, rowIndexesToDelete);
  }
}

function deleteRowsDescending_(sheet, rows) {
  rows.sort(function(a, b) { return b - a; });
  rows.forEach(function(rowIndex) {
    sheet.deleteRow(rowIndex);
  });
}

function getCurrentMonthStart_() {
  var now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), 1);
}

function logError_(err, fnName) {
  try {
    var sh = getOrCreateArchiveSheet_('Error Logs', [
      'Timestamp', 'Function', 'Error Message', 'Stack'
    ]);

    appendRowsToSheet_(sh, [[
      new Date(),
      fnName || '',
      err && err.message ? err.message : String(err || ''),
      err && err.stack ? err.stack : ''
    ]]);
  } catch (ignore) {}
}

/*************************************************
 * GENERAL HELPERS
 *************************************************/
function getMasterData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(MASTER_SHEET);
  if (!master) throw new Error('Master Database sheet not found.');
  return master.getDataRange().getValues();
}

function getNamedValue_(namedValues, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    if (namedValues[candidates[i]] && namedValues[candidates[i]][0] !== undefined) {
      return String(namedValues[candidates[i]][0] || '').trim();
    }
  }
  return '';
}

function getNamedValueLoose_(namedValues, targetText) {
  var target = String(targetText || '').trim().toUpperCase();
  var keys = Object.keys(namedValues || {});
  for (var i = 0; i < keys.length; i++) {
    var key = String(keys[i] || '').trim().toUpperCase();
    if (key.indexOf(target) > -1) {
      var val = namedValues[keys[i]];
      if (val && val[0] !== undefined) return String(val[0] || '').trim();
    }
  }
  return '';
}

function findHeaderIndex_(headers, candidates) {
  var normalizedHeaders = headers.map(function(h) {
    return normalizeHeader_(h);
  });

  for (var i = 0; i < candidates.length; i++) {
    var target = normalizeHeader_(candidates[i]);
    var idx = normalizedHeaders.indexOf(target);
    if (idx > -1) return idx;
  }
  return -1;
}

function normalizeHeader_(value) {
  return String(value || '').trim().toUpperCase().replace(/\s+/g, ' ');
}

function normalizeScanCode_(value) {
  return String(value || '').trim().toUpperCase().replace(/[^0-9A-Z]/g, '');
}

function normalizeCodesArray_(arr) {
  return (arr || []).map(function(x) {
    return normalizeScanCode_(x);
  }).filter(function(x) {
    return !!x;
  });
}

function normalizeUniqueCodesArray_(arr) {
  var out = [];
  var seen = {};
  normalizeCodesArray_(arr).forEach(function(code) {
    if (seen[code]) return;
    seen[code] = true;
    out.push(code);
  });
  return out;
}

function parseStoredCodes_(value) {
  return normalizeUniqueCodesArray_(String(value || '').split(/\r?\n|,/));
}

function parseStoredCodesRaw_(value) {
  return normalizeCodesArray_(String(value || '').split(/\r?\n|,/));
}

function buildInventoryRound_() {
  return 'Inventory ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function parseSheetDate_(value) {
  if (!value) return null;

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return new Date(value);
  }

  var s = String(value).trim();
  if (!s) return null;

  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;

  var parts = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (parts) {
    var month = parseInt(parts[1], 10) - 1;
    var day = parseInt(parts[2], 10);
    var year = parseInt(parts[3], 10);
    if (year < 100) year += 2000;
    var parsed = new Date(year, month, day);
    if (!isNaN(parsed.getTime())) return parsed;
  }

  return null;
}

function formatAnyDate_(value) {
  var d = parseSheetDate_(value);
  if (!d) return '-';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
}

function formatDateOnly_(value) {
  var d = parseSheetDate_(value);
  if (!d) return '-';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
}

function addMonths_(date, months) {
  var d = parseSheetDate_(date);
  if (!d) return '';
  var out = new Date(d);
  out.setMonth(out.getMonth() + months);
  return out;
}

function addHours_(date, hours) {
  var d = parseSheetDate_(date);
  if (!d) return null;
  var out = new Date(d);
  out.setHours(out.getHours() + Number(hours || 0));
  return out;
}

function appendRowsToSheet_(sheet, rows) {
  if (!rows || !rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function writeBackRowBlocks_(sheet, headersLength, rows, changedIndexes) {
  if (!changedIndexes || !changedIndexes.length) return;

  var idxs = changedIndexes.map(function(x) { return Number(x); })
    .filter(function(x) { return !isNaN(x) && x >= 0; })
    .sort(function(a, b) { return a - b; });

  if (!idxs.length) return;

  var start = idxs[0];
  var prev = idxs[0];

  for (var i = 1; i <= idxs.length; i++) {
    var current = idxs[i];
    if (i < idxs.length && current === prev + 1) {
      prev = current;
      continue;
    }

    var blockRows = rows.slice(start, prev + 1);
    sheet.getRange(start + 2, 1, blockRows.length, headersLength).setValues(blockRows);

    if (i < idxs.length) {
      start = current;
      prev = current;
    }
  }
}

function dedupeRowsByKey_(rows, keyBuilder) {
  var out = [];
  var seen = {};
  (rows || []).forEach(function(r) {
    var key = keyBuilder(r);
    if (seen[key]) return;
    seen[key] = true;
    out.push(r);
  });
  return out;
}

function appendUniqueRowsToSheet_(sheet, rows, keyBuilder) {
  if (!rows || !rows.length) return;

  var existing = {};
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow > 1 && lastCol > 0) {
    var current = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    current.forEach(function(r) {
      existing[keyBuilder(r)] = true;
    });
  }

  var toAppend = [];
  rows.forEach(function(r) {
    var key = keyBuilder(r);
    if (existing[key]) return;
    existing[key] = true;
    toAppend.push(r);
  });

  if (toAppend.length) {
    appendRowsToSheet_(sheet, toAppend);
  }
}

function formatKeyDate_(value) {
  var d = parseSheetDate_(value);
  return d ? d.getTime() : String(value || '');
}

/*************************************************
 * PDF EXPORT
 *************************************************/
function exportSheetToPdf_(spreadsheetId, sheetId, fileName) {
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?' + [
    'format=pdf',
    'portrait=true',
    'size=A4',
    'fitw=true',
    'sheetnames=false',
    'printtitle=false',
    'pagenumbers=false',
    'gridlines=false',
    'fzr=false',
    'gid=' + sheetId,
    'top_margin=0.50',
    'bottom_margin=0.50',
    'left_margin=0.50',
    'right_margin=0.50'
  ].join('&');

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Failed to export PDF. Response code: ' + response.getResponseCode());
  }

  return response.getBlob().setName(fileName);
}

function debugBoundSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('NAME: ' + ss.getName());
  Logger.log('ID: ' + ss.getId());
  Logger.log('URL: ' + ss.getUrl());

  var master = ss.getSheetByName('Master Database');
  Logger.log('MASTER EXISTS: ' + !!master);
  Logger.log('MASTER LAST ROW: ' + (master ? master.getLastRow() : 0));
}

function testArchiveConnectionSimple() {
  try {
    var archive = SpreadsheetApp.openByUrl(ARCHIVE_SPREADSHEET_URL);
    Logger.log('ARCHIVE NAME: ' + archive.getName());
    Logger.log('ARCHIVE ID: ' + archive.getId());
    return 'OK';
  } catch (err) {
    Logger.log('ARCHIVE ERROR: ' + err.message);
    throw err;
  }
}

/*************************************************
 * PHDU EXTENSION - ADD MACHINE + QR CODE
 * STRICT IMPROVEMENTS ONLY
 * NO DRIFT
 * USES EXISTING LOCKED SYSTEM VARIABLES/HELPERS
 *************************************************/

/**
 * Add new machine to Master Database.
 * Required:
 * - biomed
 * - name
 * - status
 *
 * Optional:
 * - location
 * - remarks
 * - lastServiceDate
 * - category
 */
function addMachine(payload) {
  try {
payload = payload || {};

    var biomed = normalizeScanCode_(payload.biomed || '');
    var name = String(payload.name || '').trim();
    var status = String(payload.status || '').trim().toUpperCase();
    var location = String(payload.location || '').trim();
    var remarks = String(payload.remarks || '').trim();
    var category = String(payload.category || '').trim();
    var lastServiceDateRaw = payload.lastServiceDate || payload.lastService || '';

    if (!biomed) throw new Error('Biomedical Number is required.');
    if (!name) throw new Error('Equipment Name is required.');
    if (!status) throw new Error('Status is required.');

    var allowedStatuses = {
      'WORKING': true,
      'MISSING': true,
      'LOANED': true,
      'UNDER MAINTENANCE': true,
      'MAINTENANCE': true,
      'CONDEMNED': true,
      'UNACCOUNTED': true
    };
    if (!allowedStatuses[status]) {
      throw new Error('Invalid Status. Use WORKING, MISSING, LOANED, UNDER MAINTENANCE, MAINTENANCE, CONDEMNED, or UNACCOUNTED.');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(MASTER_SHEET);
    if (!sh) throw new Error('Master Database sheet not found.');

    var data = sh.getDataRange().getValues();
    if (!data || !data.length) throw new Error('Master Database has no headers.');

    var headers = data[0];

    var idxBiomed = findHeaderIndex_(headers, ['Biomedical Number', 'BIOMEDICAL NUMBER', 'QR Code ID', 'Biomed Number']);
    var idxName = findHeaderIndex_(headers, ['Equipment Name', 'EQUIPMENT NAME']);
    var idxCategory = findHeaderIndex_(headers, ['Category', 'CATEGORY']);
    var idxStatus = findHeaderIndex_(headers, ['Status', 'STATUS']);
    var idxLocation = findHeaderIndex_(headers, ['Location/Ward', 'LOCATION/WARD', 'Location', 'LOCATION']);
    var idxRemarks = findHeaderIndex_(headers, ['Remarks', 'REMARKS']);
    var idxLastServiceDate = findHeaderIndex_(headers, ['Last Service Date', 'LAST SERVICE DATE']);
    var idxNextServiceDue = findHeaderIndex_(headers, ['Next Service Due', 'NEXT SERVICE DUE', 'Next Service Date', 'NEXT SERVICE DATE', 'SERVICE DUE']);
    var idxSeen = findHeaderIndex_(headers, ['SEEN THIS ROUND', 'Seen This Round']);
    var idxLastScannedDate = findHeaderIndex_(headers, ['LAST SCANNED DATE', 'Last Scanned Date']);
    var idxLastScannedBy = findHeaderIndex_(headers, ['LAST SCANNED BY', 'Last Scanned By']);
    var idxInventoryDone = findHeaderIndex_(headers, ['INVENTORY DONE', 'Inventory Done']);
    var idxQrGenerated = findHeaderIndex_(headers, ['QR GENERATED', 'Qr Generated', 'QR Generated']);
    var idxQrLink = findHeaderIndex_(headers, ['QR LINK', 'Qr Link', 'QR Links', 'Qr Links']);
var idxQrCodeId = findHeaderIndex_(headers, ['QR Code ID', 'Qr Code ID', 'QR CODE ID']);
    if (idxBiomed === -1) throw new Error('Biomedical Number / QR Code ID column not found.');
    if (idxName === -1) throw new Error('Equipment Name column not found.');
    if (idxStatus === -1) throw new Error('Status column not found.');

    for (var i = 1; i < data.length; i++) {
      var existing = normalizeScanCode_(data[i][idxBiomed]);
      if (existing && existing === biomed) {
        return {
          success: false,
          duplicate: true,
          message: 'Duplicate Biomedical Number detected.',
          existing: {
            biomed: String(data[i][idxBiomed] || '').trim(),
            name: idxName > -1 ? String(data[i][idxName] || '').trim() : '',
            status: idxStatus > -1 ? String(data[i][idxStatus] || '').trim() : '',
            location: idxLocation > -1 ? String(data[i][idxLocation] || '').trim() : '',
            remarks: idxRemarks > -1 ? String(data[i][idxRemarks] || '').trim() : ''
          }
        };
      }
    }

    var lastServiceDate = '';
    var nextServiceDue = '';
    if (lastServiceDateRaw) {
      var parsedLast = parseSheetDate_(lastServiceDateRaw);
      if (!parsedLast) throw new Error('Invalid Last Service Date.');
      lastServiceDate = parsedLast;
      nextServiceDue = addMonths_(parsedLast, 11);
    }

    var row = new Array(headers.length);
    for (var c = 0; c < headers.length; c++) row[c] = '';

    row[idxBiomed] = biomed;
    row[idxName] = name;
    row[idxStatus] = status;

    if (idxCategory > -1) row[idxCategory] = category || '';
    if (idxLocation > -1) row[idxLocation] = location || '';
    if (idxRemarks > -1) row[idxRemarks] = remarks || '';
    if (idxLastServiceDate > -1 && lastServiceDate) row[idxLastServiceDate] = lastServiceDate;
    if (idxNextServiceDue > -1 && nextServiceDue) row[idxNextServiceDue] = nextServiceDue;
    if (idxSeen > -1) row[idxSeen] = '';
    if (idxLastScannedDate > -1) row[idxLastScannedDate] = '';
    if (idxLastScannedBy > -1) row[idxLastScannedBy] = '';
    if (idxInventoryDone > -1) row[idxInventoryDone] = '';

    var qrUrl = buildDirectBiomedicalQrUrl_(biomed);
    if (idxQrGenerated > -1) row[idxQrGenerated] = biomed;
    if (idxQrLink > -1) row[idxQrLink] = qrUrl;
if (idxQrCodeId > -1) row[idxQrCodeId] = biomed;
Logger.log('idxBiomed=' + idxBiomed);
Logger.log('headers=' + JSON.stringify(headers));
Logger.log('rowBeforeAppend=' + JSON.stringify(row));

    sh.appendRow(row);

var newRow = sh.getLastRow();

if (idxBiomed > -1) {
  sh.getRange(newRow, idxBiomed + 1).setNumberFormat('@');
  sh.getRange(newRow, idxBiomed + 1).setValue("'" + String(biomed));
}

if (idxQrCodeId > -1) {
  sh.getRange(newRow, idxQrCodeId + 1).setNumberFormat('@');
  sh.getRange(newRow, idxQrCodeId + 1).setValue("'" + String(biomed));
}

    appendUpdateLog_({
      timestamp: new Date(),
      source: 'ADD_MACHINE',
      biomed: biomed,
      equipmentName: name,
      status: status,
      location: location,
      checkedBy: 'SYSTEM',
      note: 'Machine added from dashboard'
    });

    invalidateDashboardCache_();
    SpreadsheetApp.flush();

    return {
      success: true,
      duplicate: false,
      message: 'Machine added successfully.',
      biomed: biomed,
      name: name,
      status: status,
      location: location || '',
      remarks: remarks || '',
      lastServiceDate: formatDateOnly_(lastServiceDate),
      nextServiceDue: formatDateOnly_(nextServiceDue),
      qrUrl: qrUrl
    };
  } catch (err) {
    logError_(err, 'addMachine');
    throw err;
  }
}

/**
 * Replace existing machine row when duplicate Biomedical Number is confirmed.
 */
function replaceMachine(payload) {
  try {
    payload = payload || {};

    var biomed = normalizeScanCode_(payload.biomed || '');
    var name = String(payload.name || '').trim();
    var status = String(payload.status || '').trim().toUpperCase();
    var location = String(payload.location || '').trim();
    var remarks = String(payload.remarks || '').trim();
    var category = String(payload.category || '').trim();
    var lastServiceDateRaw = payload.lastServiceDate || payload.lastService || '';

    if (!biomed) throw new Error('Biomedical Number is required.');
    if (!name) throw new Error('Equipment Name is required.');
    if (!status) throw new Error('Status is required.');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(MASTER_SHEET);
    if (!sh) throw new Error('Master Database sheet not found.');

    var data = sh.getDataRange().getValues();
    if (!data || data.length < 2) throw new Error('Master Database is empty.');

    var headers = data[0];

    var idxBiomed = findHeaderIndex_(headers, ['Biomedical Number', 'BIOMEDICAL NUMBER', 'QR Code ID', 'Biomed Number']);
    var idxName = findHeaderIndex_(headers, ['Equipment Name', 'EQUIPMENT NAME']);
    var idxCategory = findHeaderIndex_(headers, ['Category', 'CATEGORY']);
    var idxStatus = findHeaderIndex_(headers, ['Status', 'STATUS']);
    var idxLocation = findHeaderIndex_(headers, ['Location/Ward', 'LOCATION/WARD', 'Location', 'LOCATION']);
    var idxRemarks = findHeaderIndex_(headers, ['Remarks', 'REMARKS']);
    var idxLastServiceDate = findHeaderIndex_(headers, ['Last Service Date', 'LAST SERVICE DATE']);
    var idxNextServiceDue = findHeaderIndex_(headers, ['Next Service Due', 'NEXT SERVICE DUE', 'Next Service Date', 'NEXT SERVICE DATE', 'SERVICE DUE']);
    var idxQrGenerated = findHeaderIndex_(headers, ['QR GENERATED', 'Qr Generated', 'QR Generated']);
    var idxQrLink = findHeaderIndex_(headers, ['QR LINK', 'Qr Link', 'QR Links', 'Qr Links']);

    if (idxBiomed === -1) throw new Error('Biomedical Number / QR Code ID column not found.');

    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      var existing = normalizeScanCode_(data[i][idxBiomed]);
      if (existing && existing === biomed) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      return addMachine(payload);
    }

    var lastServiceDate = '';
    var nextServiceDue = '';
    if (lastServiceDateRaw) {
      var parsedLast = parseSheetDate_(lastServiceDateRaw);
      if (!parsedLast) throw new Error('Invalid Last Service Date.');
      lastServiceDate = parsedLast;
      nextServiceDue = addMonths_(parsedLast, 11);
    }

    if (idxName > -1) sh.getRange(targetRow, idxName + 1).setValue(name);
    if (idxStatus > -1) sh.getRange(targetRow, idxStatus + 1).setValue(status);
    if (idxCategory > -1) sh.getRange(targetRow, idxCategory + 1).setValue(category || '');
    if (idxLocation > -1) sh.getRange(targetRow, idxLocation + 1).setValue(location || '');
    if (idxRemarks > -1) sh.getRange(targetRow, idxRemarks + 1).setValue(remarks || '');
    if (idxLastServiceDate > -1) sh.getRange(targetRow, idxLastServiceDate + 1).setValue(lastServiceDate || '');
    if (idxNextServiceDue > -1) sh.getRange(targetRow, idxNextServiceDue + 1).setValue(nextServiceDue || '');

    var qrUrl = buildDirectBiomedicalQrUrl_(biomed);
    if (idxQrGenerated > -1) sh.getRange(targetRow, idxQrGenerated + 1).setValue(biomed);
    if (idxQrLink > -1) sh.getRange(targetRow, idxQrLink + 1).setValue(qrUrl);

    appendUpdateLog_({
      timestamp: new Date(),
      source: 'REPLACE_MACHINE',
      biomed: biomed,
      equipmentName: name,
      status: status,
      location: location,
      checkedBy: 'SYSTEM',
      note: 'Duplicate biomedical replaced from dashboard'
    });

    invalidateDashboardCache_();
    SpreadsheetApp.flush();

    return {
      success: true,
      duplicate: false,
      replaced: true,
      message: 'Existing machine replaced successfully.',
      biomed: biomed,
      name: name,
      status: status,
      location: location || '',
      remarks: remarks || '',
      lastServiceDate: formatDateOnly_(lastServiceDate),
      nextServiceDue: formatDateOnly_(nextServiceDue),
      qrUrl: qrUrl
    };
  } catch (err) {
    logError_(err, 'replaceMachine');
    throw err;
  }
}

/**
 * Get one machine for duplicate checking / QR generation.
 */
function getMachineByBiomed(biomed) {
  try {
    var target = normalizeScanCode_(biomed || '');
    if (!target) return null;

    var data = getMasterData_();
    if (!data || data.length < 2) return null;

    var headers = data[0];
    var rows = data.slice(1);

    var idxBiomed = findHeaderIndex_(headers, ['Biomedical Number', 'BIOMEDICAL NUMBER', 'QR Code ID', 'Biomed Number']);
    var idxName = findHeaderIndex_(headers, ['Equipment Name', 'EQUIPMENT NAME']);
    var idxCategory = findHeaderIndex_(headers, ['Category', 'CATEGORY']);
    var idxStatus = findHeaderIndex_(headers, ['Status', 'STATUS']);
    var idxLocation = findHeaderIndex_(headers, ['Location/Ward', 'LOCATION/WARD', 'Location', 'LOCATION']);
    var idxRemarks = findHeaderIndex_(headers, ['Remarks', 'REMARKS']);
    var idxLastServiceDate = findHeaderIndex_(headers, ['Last Service Date', 'LAST SERVICE DATE']);
    var idxNextServiceDue = findHeaderIndex_(headers, ['Next Service Due', 'NEXT SERVICE DUE', 'Next Service Date', 'NEXT SERVICE DATE', 'SERVICE DUE']);

    for (var i = 0; i < rows.length; i++) {
      var code = normalizeScanCode_(rows[i][idxBiomed]);
      if (code === target) {
        return {
          biomed: String(rows[i][idxBiomed] || '').trim(),
          name: idxName > -1 ? String(rows[i][idxName] || '').trim() : '',
          category: idxCategory > -1 ? String(rows[i][idxCategory] || '').trim() : '',
          status: idxStatus > -1 ? String(rows[i][idxStatus] || '').trim() : '',
          location: idxLocation > -1 ? String(rows[i][idxLocation] || '').trim() : '',
          remarks: idxRemarks > -1 ? String(rows[i][idxRemarks] || '').trim() : '',
          lastServiceDate: idxLastServiceDate > -1 ? formatDateOnly_(rows[i][idxLastServiceDate]) : '',
          nextServiceDue: idxNextServiceDue > -1 ? formatDateOnly_(rows[i][idxNextServiceDue]) : '',
          qrUrl: buildDirectBiomedicalQrUrl_(code)
        };
      }
    }

    return null;
  } catch (err) {
    logError_(err, 'getMachineByBiomed');
    throw err;
  }
}

/**
 * NEW QR ONLY or FULL QR SHEET
 * mode:
 * - NEW_ONLY
 * - FULL_QR_SHEET
 */
function generateQrCodePdf(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tempSheet = null;

  try {
    payload = payload || {};
    var mode = String(payload.mode || 'NEW_ONLY').trim().toUpperCase();

    var machines = [];
    if (mode === 'NEW_ONLY') {
      var biomed = normalizeScanCode_(payload.biomed || '');
      if (!biomed) throw new Error('Biomedical Number is required for NEW_ONLY QR.');
      var one = getMachineByBiomed(biomed);
      if (!one) throw new Error('Machine not found in Master Database: ' + biomed);
      machines = [one];
    } else {
      machines = getAllMachinesForQR();
      if (!machines.length) throw new Error('No machines found for QR sheet.');
    }

    var tempName = 'PHDU_QR_TEMP_' + new Date().getTime();
    tempSheet = ss.insertSheet(tempName);


buildQrStickerSheet_(tempSheet, machines.slice(0, 10));
    SpreadsheetApp.flush();

    var fileName = mode === 'NEW_ONLY'
      ? ('PHDU_QR_' + machines[0].biomed + '.pdf')
      : 'PHDU_Full_QR_Sheet.pdf';

    var blob = exportSheetToPdf_(ss.getId(), tempSheet.getSheetId(), fileName);

    return {
      success: true,
      mode: mode,
      count: machines.length,
      fileName: fileName,
      blob: Utilities.base64Encode(blob.getBytes()),
      mimeType: blob.getContentType()
    };
  } catch (err) {
    logError_(err, 'generateQrCodePdf');
    throw err;
  } finally {
    try {
      if (tempSheet && ss.getSheetByName(tempSheet.getName())) {
        ss.deleteSheet(tempSheet);
      }
    } catch (ignore) {}
  }
}

function emailQrCodePdf(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tempSheet = null;

  try {
    payload = payload || {};
    var mode = String(payload.mode || 'NEW_ONLY').trim().toUpperCase();

  var machines = [];

if (mode === 'NEW_ONLY') {
  throw new Error('Use QR selection (14 machines only).');
} 
else if (mode === 'SELECTED_QR') {

  var selected = payload.selected || [];

  if (selected.length !== 14) {
    throw new Error('Exactly 14 machines must be selected.');
  }

  var allData = getAllMachinesForQR();

  var map = {};
  allData.forEach(function(x) {
    map[x.biomed] = x;
  });

  var selectedMachines = selected.map(function(biomed) {
    var m = map[String(biomed)];

    if (!m) {
      throw new Error('Machine not found: ' + biomed);
    }

    return {
      biomed: m.biomed,
      name: m.name,
      qrUrl: 'https://chart.googleapis.com/chart?chs=300x300&cht=qr&chl=' + m.biomed
    };
  });

  machines = selectedMachines;
}    

    var tempName = 'PHDU_QR_EMAIL_TEMP_' + new Date().getTime();
    tempSheet = ss.insertSheet(tempName);

    buildQrStickerSheet_(tempSheet, machines);
    SpreadsheetApp.flush();

    var fileName = 'PHDU_14_QR_Stickers.pdf';

    var blob = exportSheetToPdf_(ss.getId(), tempSheet.getSheetId(), fileName);

    MailApp.sendEmail({
      to: EMAIL_TO,
      subject: mode === 'NEW_ONLY' ? ('PHDU QR Sticker - ' + machines[0].biomed) : 'PHDU Full QR Sheet',
      body: [
        'PHDU QR CODE PDF',
        '',
        'Mode: ' + mode,
        'Total Sticker(s): ' + machines.length,
        'Generated: ' + formatAnyDate_(new Date())
      ].join('\n'),
      attachments: [blob],
      name: EMAIL_SENDER_NAME
    });

    return {
      success: true,
      mode: mode,
      count: machines.length,
      message: 'QR PDF emailed successfully to ' + EMAIL_TO + '.'
    };
  } catch (err) {
    logError_(err, 'emailQrCodePdf');
    throw err;
  } finally {
    try {
      if (tempSheet && ss.getSheetByName(tempSheet.getName())) {
        ss.deleteSheet(tempSheet);
      }
    } catch (ignore) {}
  }

function buildQrStickerSheet_(sheet, machines) {
  sheet.clear();
  sheet.setHiddenGridlines(true);

  var row = 1;
  var colLeft = 1;
  var colRight = 6;
  var blockHeight = 17;
  var count = 0;

  for (var i = 0; i < machines.length; i++) {
    var machine = machines[i];
    var colStart = (count % 2 === 0) ? colLeft : colRight;

    writeQrStickerBlock_(sheet, row, colStart, machine);
    count++;
    if (count % 2 === 0) {
      row += blockHeight;
    }
  }

  for (var c = 1; c <= 10; c++) {
    sheet.setColumnWidth(c, 95);
  }

  sheet.setRowHeights(1, Math.max(sheet.getMaxRows(), row + blockHeight), 24);
}

function writeQrStickerBlock_(sheet, startRow, startCol, machine) {
  var endCol = startCol + 3;

  sheet.getRange(startRow, startCol, 1, 4).merge();
  sheet.getRange(startRow, startCol).setValue('ROYAL HOSPITAL - PHDU');
  sheet.getRange(startRow, startCol).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');

  sheet.getRange(startRow + 1, startCol, 8, 4).merge();
  sheet.getRange(startRow + 1, startCol).setFormula('=IMAGE("' + machine.qrUrl + '",4,220,220)');
  sheet.getRange(startRow + 1, startCol).setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange(startRow + 9, startCol, 1, 4).merge();
  sheet.getRange(startRow + 9, startCol).setValue(machine.name || '-');
  sheet.getRange(startRow + 9, startCol).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center').setWrap(true);

  sheet.getRange(startRow + 10, startCol, 1, 4).merge();
  sheet.getRange(startRow + 10, startCol).setValue('Biomedical Number: ' + (machine.biomed || '-'));
  sheet.getRange(startRow + 10, startCol).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');

  sheet.getRange(startRow + 11, startCol, 1, 4).merge();
  sheet.getRange(startRow + 11, startCol).setValue(machine.category ? ('Category: ' + machine.category) : '');
  sheet.getRange(startRow + 11, startCol).setFontSize(10).setHorizontalAlignment('center');

  sheet.getRange(startRow, startCol, 12, 4).setBorder(true, true, true, true, true, true, '#1f1f1f', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function buildDirectBiomedicalQrUrl_(biomed) {
  return 'https://quickchart.io/qr?text=' + encodeURIComponent(String(biomed || '').trim()) + '&size=260';
}

/**
 * Simple helper for UI preload.
 */
function getAddMachineDefaults() {
  return {
    statuses: [
      'WORKING',
      'MISSING',
      'LOANED',
      'UNDER MAINTENANCE',
      'CONDEMNED'
    ],
    emailTo: EMAIL_TO
  };
}

function testFullQrNow() {
  return emailQrCodePdf({ mode: 'FULL_QR_SHEET' });
}

function testQrMachineCount() {
  try {
    var data = getMasterData_();
    if (!data || !data.length) return 'No data found';
    var rows = data.slice(1);
    return 'TOTAL MACHINES: ' + rows.length;
  } catch (e) {
    return 'ERROR: ' + e.message;
  }
}

function testQrPdfOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tempSheet = null;

  try {
    Logger.log('STEP 1: Getting machines');

    var machines = getAllMachinesForQR();
    Logger.log('STEP 2: Machines count = ' + machines.length);

    if (!machines.length) throw new Error('No machines found for QR sheet.');

    tempSheet = ss.insertSheet('QR_TEST_' + new Date().getTime());
    Logger.log('STEP 3: Sheet created');

    buildQrStickerSheet_(tempSheet, machines.slice(0, 10));
    Logger.log('STEP 4: QR sheet built');

    SpreadsheetApp.flush();
    Logger.log('STEP 5: Flush done');

    Logger.log('STEP 6: Exporting PDF...');
    var blob = exportSheetToPdf_(ss.getId(), tempSheet.getSheetId(), 'QR_TEST.pdf');

    Logger.log('STEP 7: DONE ' + blob.getName());
    return 'PDF OK';
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
    Logger.log(e.stack);
    throw e;
  } finally {
    try {
      if (tempSheet && ss.getSheetByName(tempSheet.getName())) {
        ss.deleteSheet(tempSheet);
      }
    } catch (ignore) {}
  }
}

function testSimple() {
  return 'OK';
}
function generate14QrStickers(selectedBiomeds) {
  if (!selectedBiomeds || !selectedBiomeds.length) {
    throw new Error('No machines selected.');
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master Database');
  var data = sheet.getDataRange().getValues();

  var headers = data.shift();

  var biomedIndex = headers.indexOf('Biomedical Number');
  var nameIndex = headers.indexOf('Equipment Name');
  var ipIndex = headers.indexOf('IP Number'); // if exists
  var locationIndex = headers.indexOf('Location');

  var selectedMachines = data.filter(function(row) {
    return selectedBiomeds.includes(String(row[biomedIndex]));
  });

  if (selectedMachines.length === 0) {
    throw new Error('Selected machines not found.');
  }

  // Limit to max 14 but allow fewer
if (selectedMachines.length > 14) {
  selectedMachines = selectedMachines.slice(0, 14);
}


  // 👉 BUILD HTML FOR PDF
  var html = '<html><head><style>' +
    'body{font-family:Arial;padding:10px}' +
    '.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}' +
    '.card{border:2px solid #000;border-radius:12px;padding:10px;height:180px;position:relative}' +
    '.title{font-weight:800;font-size:12px}' +
    '.code{font-size:14px;font-weight:800;margin-bottom:4px}' +
    '.qr{margin-top:6px}' +
    '.label{font-size:10px;font-weight:800;margin-top:4px}' +
    '.hole{position:absolute;top:4px;left:50%;transform:translateX(-50%);width:10px;height:10px;border-radius:50%;border:2px solid #000}' +
    '</style></head><body>';

  html += '<div class="grid">';

  selectedMachines.forEach(function(row) {
    var biomed = row[biomedIndex];
    var name = row[nameIndex];
    var ip = ipIndex > -1 ? row[ipIndex] : '';

    var fastQR = 'https://chart.googleapis.com/chart?chs=120x120&cht=qr&chl=' + biomed;
    var formQR = 'https://chart.googleapis.com/chart?chs=120x120&cht=qr&chl=FORM_' + biomed;

    html += '<div class="card">';
    html += '<div class="hole"></div>';

    html += '<div class="title">' + name + '</div>';
    html += '<div class="code">BIOMED: ' + biomed + '</div>';

    if (ip) {
      html += '<div class="label">IP: ' + ip + '</div>';
    }

    html += '<div class="qr">';
    html += '<div class="label">FAST SCAN</div>';
    html += '<img src="' + fastQR + '">';
    html += '</div>';

    html += '<div class="qr">';
    html += '<div class="label">ON FORM</div>';
    html += '<img src="' + formQR + '">';
    html += '</div>';

    html += '</div>';
  });

  html += '</div></body></html>';

  var blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf')
    .setName('PHDU_QR_14_STICKERS.pdf');

  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: 'PHDU QR Stickers (14)',
    body: 'Attached QR sticker sheet.',
    attachments: [blob]
  });

  return { message: 'QR sticker PDF sent to your email.' };
}
}

function getAllMachinesForQR() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master Database');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var out = [];
  for (var i = 1; i < data.length; i++) {
    var biomed = normalizeScanCode_(data[i][0]);  // Biomedical Number column
    var name = String(data[i][1] || '').trim();    // Equipment Name column

    if (!biomed) continue; // Skip blank rows

    out.push({
      biomed: biomed,
      name: name
    });
  }

  return out;
}

