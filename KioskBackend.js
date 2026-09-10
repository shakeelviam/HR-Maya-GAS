// ============================================================
// KioskBackend.gs  — additions to the existing GAS project
// ============================================================
// Add this as a new file in your Apps Script project.
// It adds one new GET method to the doGet router:
//   markBackdatedDayOff — employee logs a Day Off for a past date
//
// In your doGet / router, add:
//   case 'markBackdatedDayOff': return wrap_(markBackdatedDayOff(p)); break;
// (where wrap_ is whatever your response-wrapping helper is called)
// ============================================================

/**
 * Kiosk: record a backdated Day Off punch for a specific past date.
 * Called from the kiosk when an employee logs a missed Day Off.
 *
 * Params (all strings via GET):
 *   empId  — employee ID (MT-XXXXX)
 *   pin    — employee PIN
 *   date   — dd-mm-yyyy  (must be in the past, not today)
 *   lat    — latitude (optional, empty if unavailable)
 *   lng    — longitude (optional)
 *
 * Returns: { success, message } or { success:false, error }
 */
function markBackdatedDayOff(params) {
  try {
    const empId = String(params.empId  || '').trim();
    const pin   = String(params.pin    || '').trim();
    const date  = String(params.date   || '').trim();  // dd-mm-yyyy
    const lat   = params.lat !== '' ? parseFloat(params.lat) : null;
    const lng   = params.lng !== '' ? parseFloat(params.lng) : null;

    if (!empId || !pin || !date) {
      return { success: false, error: 'empId, pin and date are required.' };
    }

    // Validate date format dd-mm-yyyy
    if (!/^\d{2}-\d{2}-\d{4}$/.test(date)) {
      return { success: false, error: 'Date must be in dd-mm-yyyy format.' };
    }

    // Parse date and ensure it is in the past (not today or future)
    const parts   = date.split('-');
    const dateObj = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
    const today   = new Date();
    today.setHours(0, 0, 0, 0);
    if (dateObj >= today) {
      return { success: false, error: 'Backdated Day Off must be for a past date.' };
    }

    // Validate employee PIN using the existing validation logic
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const hrSheet = ss.getSheetByName('HR Maya');
    if (!hrSheet) return { success: false, error: 'HR Maya sheet not found.' };

    // Find employee by ID and validate PIN
    const hrData = hrSheet.getRange(2, 1, hrSheet.getLastRow() - 1, hrSheet.getLastColumn()).getValues();
    const headers = hrSheet.getRange(1, 1, 1, hrSheet.getLastColumn()).getValues()[0];
    const colIdx  = (name) => headers.indexOf(name);
    const ID_COL  = colIdx('Employee ID');
    const PIN_COL = colIdx('PIN');
    const NM_COL  = colIdx('Name (English)');

    const empRow = hrData.find(r => String(r[ID_COL]).trim() === empId);
    if (!empRow) return { success: false, error: 'Employee not found.' };
    if (String(empRow[PIN_COL]).trim() !== pin) return { success: false, error: 'Invalid PIN.' };

    const empName = String(empRow[NM_COL] || '').trim();

    // Check Attendance_Log — don't allow duplicate Day Off for same date
    const logSheet = ss.getSheetByName('Attendance_Log');
    if (logSheet && logSheet.getLastRow() > 1) {
      const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();
      const lHeaders = logSheet.getRange(1,1,1,logSheet.getLastColumn()).getValues()[0];
      const L_ID    = lHeaders.indexOf('Employee ID');
      const L_DATE  = lHeaders.indexOf('Date');
      const L_STAGE = lHeaders.indexOf('Stage');
      const alreadyExists = logData.some(r =>
        String(r[L_ID]).trim() === empId &&
        String(r[L_DATE]).trim() === date &&
        String(r[L_STAGE]).trim() === 'Day Off'
      );
      if (alreadyExists) {
        return { success: false, error: 'Day Off already recorded for ' + date + '.' };
      }
    }

    // Write to Attendance_Log
    // Schema: Timestamp | Employee ID | Employee Name | Date | Stage | PIN Used |
    //         Latitude | Longitude | Distance (m) | Location Flag
    const now       = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

    const newRow = [
      timestamp,
      empId,
      empName,
      date,
      'Day Off',
      'KIOSK-BACKDATED',    // PIN column — marks as backdated entry
      lat !== null ? lat : '',
      lng !== null ? lng : '',
      '',
      lat !== null ? 'BACKDATED (recorded)' : 'NO LOCATION'
    ];

    logSheet.appendRow(newRow);

    return {
      success : true,
      message : 'Day Off recorded for ' + date + '.'
    };

  } catch (e) {
    Logger.log('markBackdatedDayOff error: ' + e.message);
    return { success: false, error: 'Server error: ' + e.message };
  }
}