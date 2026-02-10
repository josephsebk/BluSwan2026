const SHEET_ID = '1ajTafy2t0VC0mUHPpS4ygHnG6uxfZxSauC_rDtPbqEQ';
const MEETING_LIST_TAB = 'Meeting List';

// ─── GET handler ─────────────────────────────────────────────────────────────
function doGet(e) {
    const action = (e && e.parameter && e.parameter.action) || 'getMeetings';
    let result;
    try {
        if (action === 'getMeetings') {
            result = getMeetings();
        } else if (action === 'ensureStatusColumns') {
            ensureStatusColumns();
            result = { success: true, message: 'Status columns ensured' };
        } else {
            result = { error: 'Unknown action: ' + action };
        }
    } catch (err) {
        result = { error: err.message, stack: err.stack };
    }
    return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
}

// ─── POST handler ────────────────────────────────────────────────────────────
function doPost(e) {
    let result;
    try {
        const body = JSON.parse(e.postData.contents);
        const action = body.action || 'startMeeting';
        if (action === 'startMeeting') {
            result = updateMeetingStatus(body.meetingId, 'Started', body.startedBy || '');
        } else if (action === 'resetMeeting') {
            result = updateMeetingStatus(body.meetingId, 'Pending', '');
        } else {
            result = { error: 'Unknown action: ' + action };
        }
    } catch (err) {
        result = { error: err.message };
    }
    return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
}

// ─── Ensure Status/StartedAt/StartedBy columns exist ─────────────────────────
function ensureStatusColumns() {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(MEETING_LIST_TAB);
    if (!sheet) throw new Error('Meeting List sheet not found');

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Find or create Status column
    let statusCol = headers.indexOf('Status');
    if (statusCol === -1) {
        // Add Status, Started At, Started By headers
        const nextCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextCol).setValue('Status').setFontWeight('bold');
        sheet.getRange(1, nextCol + 1).setValue('Started At').setFontWeight('bold');
        sheet.getRange(1, nextCol + 2).setValue('Started By').setFontWeight('bold');

        // Set all existing rows to "Pending"
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
            for (let i = 2; i <= lastRow; i++) {
                const cellA = sheet.getRange(i, 1).getValue();
                // Only set status for rows with a numeric ID
                if (typeof cellA === 'number') {
                    sheet.getRange(i, nextCol).setValue('Pending');
                }
            }
        }
    }
}

// ─── Read meetings ───────────────────────────────────────────────────────────
function getMeetings() {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(MEETING_LIST_TAB);
    if (!sheet) {
        return { meetings: [], error: 'Meeting List sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { meetings: [] };

    // Find column indices from headers (row 1)
    const headers = data[0].map(function (h) { return String(h).trim().toLowerCase(); });

    // Map headers to indices — flexible matching
    function findCol(keywords) {
        for (var k = 0; k < keywords.length; k++) {
            for (var i = 0; i < headers.length; i++) {
                if (headers[i].indexOf(keywords[k]) !== -1) return i;
            }
        }
        return -1;
    }

    var colId = findCol(['id', '#']);
    var colTimeSlot = findCol(['time', 'slot']);
    var colFund = 3; // Force Column D for Investor Name as requested

    // Find Rep column dynamically — look for "rep" header that isn't the fund/investor column
    var colRep = -1;
    for (var i = 0; i < headers.length; i++) {
        if (i === colFund) continue; // Skip the investor/fund column
        if (headers[i] === 'rep' || headers[i] === 'investor rep' || headers[i] === 'representative') {
            colRep = i;
            break;
        }
    }

    var colCompany = findCol(['company', 'startup']);
    var colFounder = findCol(['founder']);
    var colRoom = findCol(['room']);
    var colRunner = findCol(['runner']);
    var colStatus = findCol(['status']);
    var colStarted = findCol(['started at', 'startedat']);
    var colStartedBy = findCol(['started by', 'startedby']);

    var meetings = [];

    for (var i = 1; i < data.length; i++) {
        var row = data[i];

        // Get ID — skip non-numeric rows (summary rows like "Total Meetings:")
        var id = colId >= 0 ? row[colId] : i;
        if (typeof id !== 'number' || id <= 0) continue;

        var timeSlot = colTimeSlot >= 0 ? String(row[colTimeSlot]).trim() : '';
        var fund = colFund >= 0 ? String(row[colFund]).trim() : '';
        var rep = colRep >= 0 ? String(row[colRep]).trim() : '';
        var company = colCompany >= 0 ? String(row[colCompany]).trim() : '';
        var founder = colFounder >= 0 ? String(row[colFounder]).trim() : '';
        var room = colRoom >= 0 ? String(row[colRoom]).trim() : '';
        var runner = colRunner >= 0 ? String(row[colRunner]).trim() : '';
        var status = colStatus >= 0 ? String(row[colStatus]).trim().toLowerCase() : 'pending';
        var startedAt = colStarted >= 0 && row[colStarted] ? new Date(row[colStarted]).toISOString() : null;
        var startedBy = colStartedBy >= 0 ? String(row[colStartedBy]).trim() : '';

        if (!company && !fund) continue; // Skip empty rows

        // Normalize time slot format: "10:00–10:35" → "10:00 AM – 10:35 AM"
        // Handles 24h format (13:00 → 1:00 PM) and determines AM/PM from hour value
        if (timeSlot && timeSlot.indexOf('PM') === -1 && timeSlot.indexOf('AM') === -1) {
            var parts = timeSlot.split(/[–-]/);
            if (parts.length === 2) {
                var formatTime = function(t) {
                    t = t.trim();
                    var match = t.match(/^(\d{1,2}):(\d{2})$/);
                    if (!match) return t;
                    var h = parseInt(match[1], 10);
                    var m = match[2];
                    var suffix = (h >= 12) ? 'PM' : 'AM';
                    if (h > 12) h = h - 12;
                    if (h === 0) h = 12;
                    return h + ':' + m + ' ' + suffix;
                };
                timeSlot = formatTime(parts[0]) + ' – ' + formatTime(parts[1]);
            }
        }

        meetings.push({
            id: id,
            runner: runner,
            timeSlot: timeSlot,
            founder: founder,
            company: company,
            investor: fund,
            room: room,
            rep: rep,
            isPriority: false,
            status: status || 'pending',
            startedAt: startedAt,
            startedBy: startedBy,
            sheetRow: i + 1
        });
    }

    return {
        meetings: meetings,
        columns: {
            id: colId, timeSlot: colTimeSlot, fund: colFund, rep: colRep,
            company: colCompany, founder: colFounder, room: colRoom,
            runner: colRunner, status: colStatus
        },
        timestamp: new Date().toISOString()
    };
}

// ─── Update meeting status ───────────────────────────────────────────────────
function updateMeetingStatus(meetingId, newStatus, startedBy) {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(MEETING_LIST_TAB);
    if (!sheet) return { error: 'Meeting List sheet not found' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function (h) { return String(h).trim().toLowerCase(); });

    // Find status-related columns
    var colStatus = -1, colStartedAt = -1, colStartedBy = -1;
    for (var c = 0; c < headers.length; c++) {
        if (headers[c].indexOf('status') !== -1) colStatus = c;
        if (headers[c].indexOf('started at') !== -1 || headers[c] === 'startedat') colStartedAt = c;
        if (headers[c].indexOf('started by') !== -1 || headers[c] === 'startedby') colStartedBy = c;
    }

    // If no Status column exists, create them
    if (colStatus === -1) {
        colStatus = sheet.getLastColumn(); // 0-indexed from data, but getRange is 1-indexed
        sheet.getRange(1, colStatus + 1).setValue('Status').setFontWeight('bold');
        colStartedAt = colStatus + 1;
        sheet.getRange(1, colStartedAt + 1).setValue('Started At').setFontWeight('bold');
        colStartedBy = colStatus + 2;
        sheet.getRange(1, colStartedBy + 1).setValue('Started By').setFontWeight('bold');
    }

    // Find the row with matching meeting ID (column A)
    var colId = -1;
    for (var c = 0; c < headers.length; c++) {
        if (headers[c].indexOf('id') !== -1 || headers[c] === '#') { colId = c; break; }
    }

    for (var i = 1; i < data.length; i++) {
        var rowId = colId >= 0 ? data[i][colId] : i;
        if (String(rowId) === String(meetingId)) {
            // Update Status (1-indexed for getRange)
            sheet.getRange(i + 1, colStatus + 1).setValue(newStatus);

            if (newStatus === 'Started') {
                if (colStartedAt >= 0) sheet.getRange(i + 1, colStartedAt + 1).setValue(new Date());
                if (colStartedBy >= 0) sheet.getRange(i + 1, colStartedBy + 1).setValue(startedBy);
                // Highlight green
                sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()).setBackground('#d4edda');
            } else {
                // Reset
                if (colStartedAt >= 0) sheet.getRange(i + 1, colStartedAt + 1).setValue('');
                if (colStartedBy >= 0) sheet.getRange(i + 1, colStartedBy + 1).setValue('');
                sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()).setBackground(null);
            }

            return { success: true, meetingId: meetingId, status: newStatus, timestamp: new Date().toISOString() };
        }
    }

    return { error: 'Meeting ' + meetingId + ' not found' };
}

// ─── Run once to add Status columns ─────────────────────────────────────────
function setup() {
    ensureStatusColumns();
    Logger.log('Status columns ensured on Meeting List sheet');
} 
