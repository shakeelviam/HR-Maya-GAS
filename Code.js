// ============================================================
// COMPLETE CODE.GS - KIOSK + PAYROLL + LEAVE/ENCASHMENT + WHATSAPP
// ============================================================
// Adds to your version:
//  - preparePayrollInputs, generatePayrollDocs cases (fixes "Invalid method")
//  - Leave/encashment cases (functions in LeaveEncashment.gs)
//  - markKioskAttendance now accepts 'Day Off' stage (today or backdated via date param)
//  - getDailyKioskStatus returns 'day_off' if a Day Off punch exists for today
// Everything else is unchanged from your deployed version.
// NOTE: verify the three Sheet IDs below match your live file before deploying.
// ============================================================

// Sheet IDs
const MAIN_SPREADSHEET_ID = '1FDEgm7v4fogW8Nqj7FN-CCIeiizdnbCYG4cpx6Lw0sM';
const SOURCE_SPREADSHEET_ID = '1k5T2DfZqACaJLbwy16-XSfnGaLq4Vg0u1epGro94c_c';
const PAYROLL_SPREADSHEET_ID = '1QMYgO6B8fciT6pObX9PB3Nmc8TvqOAAYsYQF4bve69M';

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

// Parameters for an endpoint that may be called either way.
// A POST body wins when there is one — that is how a bulk action sends a
// list of ids without pushing them through a URL. Falls back to the query
// string so the same endpoint still works from a browser address bar.
function cgBody_(e) {
  try {
    if (e && e.postData && e.postData.contents) {
      const b = JSON.parse(e.postData.contents);
      // Keep method= and anything else that only came on the query string.
      if (e.parameter) { Object.keys(e.parameter).forEach(k => { if (!(k in b)) b[k] = e.parameter[k]; }); }
      return b;
    }
  } catch (err) { /* malformed body — fall through to the query string */ }
  return (e && e.parameter) ? e.parameter : {};
}

function handleRequest(e) {
  try {
    const method = e.parameter.method || 'getIntegratedEmployees';
    switch(method) {
      case 'getIntegratedEmployees':
        return ContentService.createTextOutput(JSON.stringify(getIntegratedEmployees())).setMimeType(ContentService.MimeType.JSON);
      case 'getEmployees':
        return ContentService.createTextOutput(JSON.stringify(getEmployees())).setMimeType(ContentService.MimeType.JSON);
      case 'getPayrollData':
        return ContentService.createTextOutput(JSON.stringify(getPayrollData())).setMimeType(ContentService.MimeType.JSON);
      case 'getAttendance':
        return ContentService.createTextOutput(JSON.stringify(getAttendance(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'markAttendance':
        const data = JSON.parse(e.postData?.contents || '{}');
        return ContentService.createTextOutput(JSON.stringify(markAttendance(data))).setMimeType(ContentService.MimeType.JSON);
      case 'getLeaveRequests':
        return ContentService.createTextOutput(JSON.stringify(getLeaveRequests())).setMimeType(ContentService.MimeType.JSON);
      case 'submitLeave':
        const leaveData = JSON.parse(e.postData?.contents || '{}');
        return ContentService.createTextOutput(JSON.stringify(submitLeave(leaveData))).setMimeType(ContentService.MimeType.JSON);
      case 'updateLeaveStatus':
        const statusData = JSON.parse(e.postData?.contents || '{}');
        return ContentService.createTextOutput(JSON.stringify(updateLeaveStatus(statusData))).setMimeType(ContentService.MimeType.JSON);
      case 'getReviews':
        return ContentService.createTextOutput(JSON.stringify(getReviews())).setMimeType(ContentService.MimeType.JSON);
      case 'addReview':
        const reviewData = JSON.parse(e.postData?.contents || '{}');
        return ContentService.createTextOutput(JSON.stringify(addReview(reviewData))).setMimeType(ContentService.MimeType.JSON);
      case 'getDashboardStats':
        return ContentService.createTextOutput(JSON.stringify(getDashboardStats())).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 🟢 KIOSK CASES
      // ============================================================
      case 'loginStaffKiosk':
        return ContentService.createTextOutput(JSON.stringify(loginStaffKiosk(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'markKioskAttendance':
        return ContentService.createTextOutput(JSON.stringify(markKioskAttendance(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'submitOtRequest':
        return ContentService.createTextOutput(JSON.stringify(submitOtRequest((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getPendingOt':
        return ContentService.createTextOutput(JSON.stringify(getPendingOt(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'approvePendingOt':
        return ContentService.createTextOutput(JSON.stringify(approvePendingOt(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'rejectPendingOt':
        return ContentService.createTextOutput(JSON.stringify(rejectPendingOt(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getDailyKioskStatus':
        return ContentService.createTextOutput(JSON.stringify(getDailyKioskStatus(e.parameter))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 🎙️ VOICE SUPERVISOR CASES (functions in VoiceSupervisor.gs)
      // ============================================================
      case 'voiceLogin':
        return ContentService.createTextOutput(JSON.stringify(voiceLogin(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'voiceTeamStatus':
        return ContentService.createTextOutput(JSON.stringify(voiceTeamStatus(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'voiceFindEmployee':
        return ContentService.createTextOutput(JSON.stringify(voiceFindEmployee(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'voiceAddPunch':
        return ContentService.createTextOutput(JSON.stringify(voiceAddPunch(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'voiceDeletePunch':
        return ContentService.createTextOutput(JSON.stringify(voiceDeletePunch(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'voiceMarkAbsent':
        return ContentService.createTextOutput(JSON.stringify(voiceMarkAbsent(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'voiceMarkDayOff':
        return ContentService.createTextOutput(JSON.stringify(voiceMarkDayOff(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'voiceUndoLastAction':
        return ContentService.createTextOutput(JSON.stringify(voiceUndoLastAction(e.parameter))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 📅 SUPERVISOR ROSTER CASES (functions in RosterBackend.gs)
      // ============================================================
      case 'getRosterRange':
        return ContentService.createTextOutput(JSON.stringify(getRosterRange(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'saveRosterDraft':
        return ContentService.createTextOutput(JSON.stringify(saveRosterDraft((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'publishRoster':
        return ContentService.createTextOutput(JSON.stringify(publishRoster(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'amendRosterEntry':
        return ContentService.createTextOutput(JSON.stringify(amendRosterEntry((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'copyPreviousRange':
        return ContentService.createTextOutput(JSON.stringify(copyPreviousRange(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getDiscrepancies':
        return ContentService.createTextOutput(JSON.stringify(getDiscrepancies(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'markDiscrepancyReviewed':
        return ContentService.createTextOutput(JSON.stringify(markDiscrepancyReviewed(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getMyTeam':
        return ContentService.createTextOutput(JSON.stringify(getMyTeam(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getActiveLocations':
        return ContentService.createTextOutput(JSON.stringify(getActiveLocations(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getShiftsForLocation':
        return ContentService.createTextOutput(JSON.stringify(getShiftsForLocation(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'requestTransfer':
        return ContentService.createTextOutput(JSON.stringify(requestTransfer((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getIncomingTransferRequests':
        return ContentService.createTextOutput(JSON.stringify(getIncomingTransferRequests(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getOutgoingTransferRequests':
        return ContentService.createTextOutput(JSON.stringify(getOutgoingTransferRequests(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'approveTransfer':
        return ContentService.createTextOutput(JSON.stringify(approveTransfer(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'rejectTransfer':
        return ContentService.createTextOutput(JSON.stringify(rejectTransfer(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'cancelTransferRequest':
        return ContentService.createTextOutput(JSON.stringify(cancelTransferRequest(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'setupBranchLocations':
        return ContentService.createTextOutput(JSON.stringify(setupBranchLocations())).setMimeType(ContentService.MimeType.JSON);
      case 'recomputeGeoFlags':
        return ContentService.createTextOutput(JSON.stringify(recomputeGeoFlags(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getAttendanceLog':
        return ContentService.createTextOutput(JSON.stringify(getAttendanceLog(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'addAttendancePunch':
        return ContentService.createTextOutput(JSON.stringify(addAttendancePunch((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'editAttendancePunch':
        return ContentService.createTextOutput(JSON.stringify(editAttendancePunch((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'deleteAttendancePunch':
        return ContentService.createTextOutput(JSON.stringify(deleteAttendancePunch(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      // ============================================================
      // 🧾 OT VERIFICATION CONSOLE (functions in OtCompile.gs)
      // Separate role from supervisor — allowlist in Admin_Settings
      // under the key 'ot_compilers'.
      // ============================================================
      // ============================================================
      // ✅ SUPERVISOR VERIFICATION (functions in SupervisorVerify.gs)
      // Supervisors confirm OT claims and unresolved attendance days for
      // their own location. Verdicts are recorded, not applied — the OT
      // compiler still decides, but with a statement attached.
      // ============================================================
      case 'getSupervisorVerifyList':
        // Proves who is looking (PIN) and hides his own days from him.
        return ContentService.createTextOutput(JSON.stringify(svWebList(cgBody_(e)))).setMimeType(ContentService.MimeType.JSON);
      case 'getRosterOverview':
        // Every location's roster, read-only, for the HR admin page.
        return ContentService.createTextOutput(JSON.stringify(getRosterOverview(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'svRoles':
        return ContentService.createTextOutput(JSON.stringify(svRoles())).setMimeType(ContentService.MimeType.JSON);
      case 'saveSupervisorVerdicts':
        return ContentService.createTextOutput(JSON.stringify(saveSupervisorVerdicts((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);

      case 'getOtCompileList':
        return ContentService.createTextOutput(JSON.stringify(getOtCompileList(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'previewOtDateFix':
        return ContentService.createTextOutput(JSON.stringify(previewOtDateFix(cgBody_(e)))).setMimeType(ContentService.MimeType.JSON);
      case 'commitOtDateFix':
        return ContentService.createTextOutput(JSON.stringify(commitOtDateFix(cgBody_(e)))).setMimeType(ContentService.MimeType.JSON);
      case 'postApprovedOt':
        return ContentService.createTextOutput(JSON.stringify(postApprovedOt((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);

      case 'previewMissingCheckouts':
        return ContentService.createTextOutput(JSON.stringify(previewMissingCheckouts(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'commitMissingCheckouts':
        return ContentService.createTextOutput(JSON.stringify(commitMissingCheckouts((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getLocations':
        return ContentService.createTextOutput(JSON.stringify(getLocations(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'saveLocation':
        return ContentService.createTextOutput(JSON.stringify(saveLocation((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'deleteLocation':
        return ContentService.createTextOutput(JSON.stringify(deleteLocation(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'migrateBranchToLocation':
        return ContentService.createTextOutput(JSON.stringify(migrateBranchToLocation(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getLocationSupervisors':
        return ContentService.createTextOutput(JSON.stringify(getLocationSupervisors(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'saveLocationSupervisors':
        return ContentService.createTextOutput(JSON.stringify(saveLocationSupervisors((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getEmployeesByLocation':
        return ContentService.createTextOutput(JSON.stringify(getEmployeesByLocation(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'saveLocationShifts':
        return ContentService.createTextOutput(JSON.stringify(saveLocationShifts((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'dailyAdminDigest':
        return ContentService.createTextOutput(JSON.stringify(dailyAdminDigest())).setMimeType(ContentService.MimeType.JSON);
      case 'weeklyExpiryRoundup':
        return ContentService.createTextOutput(JSON.stringify(weeklyExpiryRoundup())).setMimeType(ContentService.MimeType.JSON);
      case 'resetAttendanceLog':
        return ContentService.createTextOutput(JSON.stringify(resetAttendanceLog(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'setAdminEmail':
        return ContentService.createTextOutput(JSON.stringify(setAdminEmail(e.parameter))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // ⏰ SHIFTS (Shifts.gs)
      // ============================================================
      case 'getShifts':
        return ContentService.createTextOutput(JSON.stringify(getShifts(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'saveShift':
        return ContentService.createTextOutput(JSON.stringify(saveShift((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'deleteShift':
        return ContentService.createTextOutput(JSON.stringify(deleteShift(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'assignShift':
        return ContentService.createTextOutput(JSON.stringify(assignShift((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 🕐 ATTENDANCE → OT (AttendanceOT.gs)
      // ============================================================
      case 'getAttendanceOtDraft':
        return ContentService.createTextOutput(JSON.stringify(getAttendanceOtDraft(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'postOtToEntries': {
        // What is already in OT_Entries for a man and day stands — never post it twice.
        const g = otGuardPostEntries((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter);
        let res;
        if (!g.body.rows || !g.body.rows.length) {
          res = { success: true, posted: 0, skipped: g.skipped,
                  message: 'Nothing posted — every ticked row is already in OT_Entries.' };
        } else {
          res = postOtToEntries(g.body);
          if (res && g.skipped.length) {
            res.skipped = g.skipped;
            res.message = (res.message || '') + ' ' + g.skipped.length +
                          ' skipped — already in OT_Entries: ' + g.skipped.join('; ');
          }
        }
        return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
      }
      case 'scrubAttendancePins':
        return ContentService.createTextOutput(JSON.stringify(scrubAttendancePins())).setMimeType(ContentService.MimeType.JSON);
      case 'archiveOtEntriesNow':
        return ContentService.createTextOutput(JSON.stringify(archiveOtEntriesNow(e.parameter))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 📋 MANUAL ATTENDANCE (ManualAttendance.gs) — supervisor backfill
      // ============================================================
      case 'getManualAttnEmployees':
        return ContentService.createTextOutput(JSON.stringify(getManualAttnEmployees(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'saveManualAttendance':
        return ContentService.createTextOutput(JSON.stringify(saveManualAttendance((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getManualAttendance':
        return ContentService.createTextOutput(JSON.stringify(getManualAttendance(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'backfillHeadOfficeAug':
        return ContentService.createTextOutput(JSON.stringify(backfillHeadOfficeAug(e.parameter))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 📱 WHATSAPP PIN BROADCAST LIST (WhatsAppPins.gs)
      // ============================================================
      case 'getPinBroadcastList':
        return ContentService.createTextOutput(JSON.stringify(getPinBroadcastList(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getPinContactsCsv':
        return ContentService.createTextOutput(JSON.stringify(getPinContactsCsv(e.parameter))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 💰 PAYROLL + ATTENDANCE AMENDMENT CASES  (functions in Payroll.gs / PayrollDocs.gs)
      // ============================================================
      case 'computeMonthlyPayroll':
        return ContentService.createTextOutput(JSON.stringify(computeMonthlyPayroll(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'amendAttendance':
        return ContentService.createTextOutput(JSON.stringify(amendAttendance(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'cancelPayrollRun':
        return ContentService.createTextOutput(JSON.stringify(cancelPayrollRun(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'preparePayrollInputs':
        return ContentService.createTextOutput(JSON.stringify(preparePayrollInputs(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'generatePayrollDocs':
        return ContentService.createTextOutput(JSON.stringify(generatePayrollDocs(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'populateBankTransfer':
        return ContentService.createTextOutput(JSON.stringify(populateBankTransfer(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'syncBankDetailsToHrMaya':
        return ContentService.createTextOutput(JSON.stringify(syncBankDetailsToHrMaya(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'syncStaffSubmissions':
        return ContentService.createTextOutput(JSON.stringify(syncStaffSubmissions(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'pullSubmissionsFromDocs':
        return ContentService.createTextOutput(JSON.stringify(pullSubmissionsFromDocs(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'refreshAndSyncSubmissions':
        return ContentService.createTextOutput(JSON.stringify(refreshAndSyncSubmissions(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'pullDesignationsFromRoster':
        return ContentService.createTextOutput(JSON.stringify(pullDesignationsFromRoster(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'refreshDesignationsFull':
        return ContentService.createTextOutput(JSON.stringify(refreshDesignationsFull(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getDesignations':
        return ContentService.createTextOutput(JSON.stringify(getDesignations(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getScheduledDeductions':
        return ContentService.createTextOutput(JSON.stringify(getScheduledDeductions(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'saveScheduledDeduction':
        return ContentService.createTextOutput(JSON.stringify(saveScheduledDeduction((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'cancelScheduledDeduction':
        return ContentService.createTextOutput(JSON.stringify(cancelScheduledDeduction(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      // ── Recruitment (Recruitment.gs) ────────────────────────────
      case 'getRecruitment':
        return ContentService.createTextOutput(JSON.stringify(getRecruitment(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'saveCandidate':
        return ContentService.createTextOutput(JSON.stringify(saveCandidate(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'setCandidateStage':
        return ContentService.createTextOutput(JSON.stringify(setCandidateStage(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'convertCandidateToEmployee':
        return ContentService.createTextOutput(JSON.stringify(convertCandidateToEmployee(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'issueAgencyPin':
        return ContentService.createTextOutput(JSON.stringify(issueAgencyPin(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'revokeAgencyPin':
        return ContentService.createTextOutput(JSON.stringify(revokeAgencyPin(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'checkAgencyLogin':
        return ContentService.createTextOutput(JSON.stringify(checkAgencyLogin(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'agentCandidates':
        return ContentService.createTextOutput(JSON.stringify(agentCandidates(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'agentAddCandidate':
        return ContentService.createTextOutput(JSON.stringify(agentAddCandidate(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      // ── QR terminal check-in (QrPunch.gs) ─────────────────────────
      // The tablet shows a rotating code; the phone scans it; the tablet
      // shows the stage page and takes the punch.
      case 'pairTerminal':
        return ContentService.createTextOutput(JSON.stringify(pairTerminal(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'terminalCodes':
        return ContentService.createTextOutput(JSON.stringify(terminalCodes(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'issueBadge':
        return ContentService.createTextOutput(JSON.stringify(issueBadge(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'badgeByLink':
        return ContentService.createTextOutput(JSON.stringify(badgeByLink(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'badgeIdentify':
        return ContentService.createTextOutput(JSON.stringify(badgeIdentify(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'revokeBadge':
        return ContentService.createTextOutput(JSON.stringify(revokeBadge(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'qrScan':
        return ContentService.createTextOutput(JSON.stringify(qrScan(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'terminalPoll':
        return ContentService.createTextOutput(JSON.stringify(terminalPoll(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'terminalPunch':
        return ContentService.createTextOutput(JSON.stringify(terminalPunch(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'listTerminals':
        return ContentService.createTextOutput(JSON.stringify(listTerminals(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'revokeTerminal':
        return ContentService.createTextOutput(JSON.stringify(revokeTerminal(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);

      case 'saveAgency':
        return ContentService.createTextOutput(JSON.stringify(saveAgency(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);

      // Scheduled extras — bonuses etc. (ScheduledExtras.gs)
      case 'getScheduledExtras':
        return ContentService.createTextOutput(JSON.stringify(getScheduledExtras(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'saveScheduledExtra':
        return ContentService.createTextOutput(JSON.stringify(saveScheduledExtra((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'cancelScheduledExtra':
        return ContentService.createTextOutput(JSON.stringify(cancelScheduledExtra(e.parameter))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 🏢 PAYROLL BY BRANCH (PayrollByBranch.gs)
      // ============================================================
      case 'buildPayrollByBranch':
        return ContentService.createTextOutput(JSON.stringify(buildPayrollByBranch(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 📄 DOCUMENT EXPIRY CHASE (DocExpiryAlert.gs)
      // ============================================================
      case 'getDocChaseList':
        return ContentService.createTextOutput(JSON.stringify(getDocChaseList(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'makeMyDocsLink':
        return ContentService.createTextOutput(JSON.stringify(makeMyDocsLink(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getMyDocuments':
        return ContentService.createTextOutput(JSON.stringify(getMyDocuments(e.parameter))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // ✈️ LEAVE SALARY — paid in advance, by voucher (LeaveSalary.gs)
      // ============================================================
      case 'getLeaveSalaryList':
        return ContentService.createTextOutput(JSON.stringify(getLeaveSalaryList(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'payLeaveSalary':
        return ContentService.createTextOutput(JSON.stringify(payLeaveSalary(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'cancelLeaveSalary':
        return ContentService.createTextOutput(JSON.stringify(cancelLeaveSalary(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'getLeaveSalaryHistory':
        return ContentService.createTextOutput(JSON.stringify(getLeaveSalaryHistory(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'generateLeaveSalaryVoucher':
        return ContentService.createTextOutput(JSON.stringify(generateLeaveSalaryVoucher(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'reprintLeaveSalaryVoucher':
        return ContentService.createTextOutput(JSON.stringify(reprintLeaveSalaryVoucher(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'setLeaveTicketAmount':
        return ContentService.createTextOutput(JSON.stringify(setLeaveTicketAmount(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 🌴 LEAVE LIFECYCLE (LeaveLifecycle.gs)
      // ============================================================
      case 'amendLeaveGrant':
        return ContentService.createTextOutput(JSON.stringify(amendLeaveGrant(((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter)))).setMimeType(ContentService.MimeType.JSON);
      case 'grantLeave':
        return ContentService.createTextOutput(JSON.stringify(grantLeave(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getLeaveGrants':
        return ContentService.createTextOutput(JSON.stringify(getLeaveGrants(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'applyLeaveStatusChanges':
        return ContentService.createTextOutput(JSON.stringify(applyLeaveStatusChanges(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getLeaveStatusReminders':
        return ContentService.createTextOutput(JSON.stringify(getLeaveStatusReminders(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getRemainingVacationTally':
        return ContentService.createTextOutput(JSON.stringify(getRemainingVacationTally(e.parameter))).setMimeType(ContentService.MimeType.JSON);

      // ── Approval gate. grantLeave now only REQUESTS; these decide. ──
      // Decisions come by POST so a bulk approve can carry more grant ids
      // than a URL should, but a GET with a single grantId still works.
      case 'approveLeaveGrant':
        return ContentService.createTextOutput(JSON.stringify(approveLeaveGrant(cgBody_(e)))).setMimeType(ContentService.MimeType.JSON);
      case 'rejectLeaveGrant':
        return ContentService.createTextOutput(JSON.stringify(rejectLeaveGrant(cgBody_(e)))).setMimeType(ContentService.MimeType.JSON);
      case 'cancelLeaveGrant':
        return ContentService.createTextOutput(JSON.stringify(cancelLeaveGrant(cgBody_(e)))).setMimeType(ContentService.MimeType.JSON);
      case 'getLeaveRegister':
        return ContentService.createTextOutput(JSON.stringify(getLeaveRegister(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'auditLeaveGrants':
        return ContentService.createTextOutput(JSON.stringify(auditLeaveGrants(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'recompileLeaveGrants':
        return ContentService.createTextOutput(JSON.stringify(recompileLeaveGrants(cgBody_(e)))).setMimeType(ContentService.MimeType.JSON);
      case 'previewLeaveApprovalMigration':
        return ContentService.createTextOutput(JSON.stringify(previewLeaveApprovalMigration(cgBody_(e)))).setMimeType(ContentService.MimeType.JSON);
      case 'commitLeaveApprovalMigration':
        return ContentService.createTextOutput(JSON.stringify(commitLeaveApprovalMigration(cgBody_(e)))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 🗂️ LEAVE TYPES + ALLOCATION (LeaveTypes.gs)
      // ============================================================
      case 'getLeaveTypesFull':
        return ContentService.createTextOutput(JSON.stringify(getLeaveTypesFull(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'saveLeaveType':
        return ContentService.createTextOutput(JSON.stringify(saveLeaveType((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'deleteLeaveType':
        return ContentService.createTextOutput(JSON.stringify(deleteLeaveType(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'allocateLeave':
        return ContentService.createTextOutput(JSON.stringify(allocateLeave((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'migrateRemainingToAnnual':
        return ContentService.createTextOutput(JSON.stringify(migrateRemainingToAnnual(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getGrantMatrix':
        return ContentService.createTextOutput(JSON.stringify(getGrantMatrix(e.parameter))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 💵 PAYROLL REVIEW (PayrollReview.gs)
      // ============================================================
      case 'getPayrollReview':
        return ContentService.createTextOutput(JSON.stringify(getPayrollReview(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'runPayrollFromReview':
        return ContentService.createTextOutput(JSON.stringify(runPayrollFromReview((e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 🌴 LEAVE + ENCASHMENT CASES  (functions in LeaveEncashment.gs)
      // ============================================================
      case 'ensureLeaveTabs':
        return ContentService.createTextOutput(JSON.stringify(ensureLeaveTabs(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getLeaveTypes':
        return ContentService.createTextOutput(JSON.stringify(getLeaveTypes(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getLeaveBalances':
        return ContentService.createTextOutput(JSON.stringify(getLeaveBalances(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'upsertLeaveBalance':
        return ContentService.createTextOutput(JSON.stringify(upsertLeaveBalance(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'bulkUploadLeaveBalances':
        const bulkData = JSON.parse(e.postData?.contents || '{}');
        return ContentService.createTextOutput(JSON.stringify(bulkUploadLeaveBalances(bulkData))).setMimeType(ContentService.MimeType.JSON);
      case 'encashLeave':
        return ContentService.createTextOutput(JSON.stringify(encashLeave(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'generateEncashmentVoucher':
        return ContentService.createTextOutput(JSON.stringify(generateEncashmentVoucher(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getEncashmentHistory':
        return ContentService.createTextOutput(JSON.stringify(getEncashmentHistory(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'reprintEncashmentVoucher':
        return ContentService.createTextOutput(JSON.stringify(reprintEncashmentVoucher(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'buildOtMergeSheet':
        return ContentService.createTextOutput(JSON.stringify(buildOtMergeSheet(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      // RETIRED: buildAttendanceDaily used the old pre-roster compile and
      // would overwrite roster-aware rows with guessed "Default Day Off".
      // The stub in AttendanceDaily.gs returns a clear error. Kept routed
      // so an old call fails loudly rather than 404-ing as 'Invalid method'.
      case 'buildAttendanceDaily':
        return ContentService.createTextOutput(JSON.stringify(buildAttendanceDaily(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getSupervisorSessions':
        return ContentService.createTextOutput(JSON.stringify(getSupervisorSessions(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'correctSession': {
        // Same rules as the verification list: proven identity, never his
        // own day, never a time that has not happened yet.
        const body = (e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter;
        const g = svGuardCorrection(body, 'correctSession');
        const res = g.ok ? correctSession(body) : { success: false, error: g.error };
        return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
      }
      case 'addSession': {
        // Same rules as the verification list: proven identity, never his
        // own day, never a time that has not happened yet.
        const body = (e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter;
        const g = svGuardCorrection(body, 'addSession');
        const res = g.ok ? addSession(body) : { success: false, error: g.error };
        return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
      }
      case 'deleteSession': {
        // Same rules as the verification list: proven identity, never his
        // own day, never a time that has not happened yet.
        const body = (e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter;
        const g = svGuardCorrection(body, 'deleteSession');
        const res = g.ok ? deleteSession(body) : { success: false, error: g.error };
        return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
      }
      case 'migrateLeaveBalances':
        return ContentService.createTextOutput(JSON.stringify(migrateLeaveBalances(e.parameter))).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // ⏱️ OT CASES  (functions in OT.gs)
      // ============================================================
      case 'addOtEntry':
        return ContentService.createTextOutput(JSON.stringify(addOtEntry(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getOtSummary':
        return ContentService.createTextOutput(JSON.stringify(getOtSummary(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'generateOtReport':
        return ContentService.createTextOutput(JSON.stringify(generateOtReport(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'setupMayaHR':
        return ContentService.createTextOutput(JSON.stringify(setupMayaHR())).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 👤 EMPLOYEE ADMIN CASES  (functions in EmployeeAdmin.gs)
      // ============================================================
      case 'saveEmployee':
        const empData = (e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter;
        return ContentService.createTextOutput(JSON.stringify(saveEmployee(empData))).setMimeType(ContentService.MimeType.JSON);
      case 'getRoster':
        return ContentService.createTextOutput(JSON.stringify(getRoster(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getEmployeeById':
        return ContentService.createTextOutput(JSON.stringify(getEmployeeById(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'getEmployeesFull':
        return ContentService.createTextOutput(JSON.stringify(getEmployeesFull(e.parameter))).setMimeType(ContentService.MimeType.JSON);
      case 'convertDerivedColumnsToFormulas':
        return ContentService.createTextOutput(JSON.stringify(convertDerivedColumnsToFormulas())).setMimeType(ContentService.MimeType.JSON);

      // ============================================================
      // 💬 WHATSAPP ATTENDANCE BOT CASE  (function in WhatsAppAttendance.gs)
      // ============================================================
      case 'waMessage':
        const waData = JSON.parse(e.postData?.contents || '{}');
        return ContentService.createTextOutput(JSON.stringify(handleWaMessage(waData))).setMimeType(ContentService.MimeType.JSON);

      default:
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Invalid method' })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// 1. GET EMPLOYEES FROM SOURCE SHEET
// ============================================================
function getEmployees() {
  try {
    const sheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID).getSheetByName('Confirmed Names');
    if (!sheet) { return { success: true, data: [] }; }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) { return { success: true, data: [] }; }
    const headers = data[0];
    const employees = [];
    const findCol = (name) => {
      for (let i = 0; i < headers.length; i++) {
        if (headers[i] && headers[i].toString().trim().toLowerCase().includes(name.toLowerCase())) { return i; }
      }
      return -1;
    };
    const nameCol = findCol('full name') !== -1 ? findCol('full name') : findCol('Name');
    const civilIdCol = findCol('civil') !== -1 ? findCol('civil') : findCol('ID');
    const basicCol = findCol('basic') !== -1 ? findCol('basic') : findCol('salary');
    const foodCol = findCol('food');
    const accomCol = findCol('accomodation');
    const conveyCol = findCol('conveyance');
    const bonusCol = findCol('bonus');
    const grossCol = findCol('gross');
    const loanCol = findCol('loan');
    const deductionCol = findCol('deduction');
    const totalPayableCol = findCol('total payable') !== -1 ? findCol('total payable') : findCol('payable');
    const branchCol = findCol('الفرع');
    const nameArCol = findCol('arabic') !== -1 ? findCol('arabic') : findCol('الاسم');
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const employee = {
        'Name (English)': nameCol !== -1 ? String(data[i][nameCol] || '').trim() : '',
        'Name (Arabic)': nameArCol !== -1 ? String(data[i][nameArCol] || '').trim() : '',
        'Civil ID': civilIdCol !== -1 ? String(data[i][civilIdCol] || '').trim() : '',
        'Basic Salary': basicCol !== -1 ? data[i][basicCol] || 0 : 0,
        'Food Allowance': foodCol !== -1 ? data[i][foodCol] || 0 : 0,
        'Accommodation': accomCol !== -1 ? data[i][accomCol] || 0 : 0,
        'Conveyance': conveyCol !== -1 ? data[i][conveyCol] || 0 : 0,
        'Bonus': bonusCol !== -1 ? data[i][bonusCol] || 0 : 0,
        'Gross Salary': grossCol !== -1 ? data[i][grossCol] || 0 : 0,
        'Loan': loanCol !== -1 ? data[i][loanCol] || 0 : 0,
        'Other Deductions': deductionCol !== -1 ? data[i][deductionCol] || 0 : 0,
        'Total Payable': totalPayableCol !== -1 ? data[i][totalPayableCol] || 0 : 0,
        'Branch': branchCol !== -1 ? data[i][branchCol] || '' : '',
        'Status': 'Active'
      };
      for (let j = 0; j < headers.length; j++) { employee[headers[j]] = data[i][j]; }
      employees.push(employee);
    }
    return { success: true, data: employees };
  } catch (error) { return { success: false, error: error.toString(), data: [] }; }
}

// ============================================================
// 2. GET PAYROLL DATA FROM SOURCE SHEET
// ============================================================
function getPayrollData() {
  try {
    const sheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID).getSheetByName('Confirmed Names');
    if (!sheet) { return { success: true, data: [] }; }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) { return { success: true, data: [] }; }
    const headers = data[0];
    const payroll = [];
    const findCol = (name) => {
      for (let i = 0; i < headers.length; i++) {
        if (headers[i] && headers[i].toString().trim().toLowerCase().includes(name.toLowerCase())) { return i; }
      }
      return -1;
    };
    const nameCol = findCol('full name') !== -1 ? findCol('full name') : findCol('Name');
    const civilIdCol = findCol('civil') !== -1 ? findCol('civil') : findCol('ID');
    const basicCol = findCol('basic') !== -1 ? findCol('basic') : findCol('salary');
    const foodCol = findCol('food');
    const accomCol = findCol('accomodation');
    const conveyCol = findCol('conveyance');
    const bonusCol = findCol('bonus');
    const grossCol = findCol('gross');
    const loanCol = findCol('loan');
    const deductionCol = findCol('deduction');
    const totalPayableCol = findCol('total payable') !== -1 ? findCol('total payable') : findCol('payable');
    const branchCol = findCol('الفرع');
    const priceHourCol = findCol('price per hour') !== -1 ? findCol('price per hour') : findCol('hour');
    const priceDayCol = findCol('price per day') !== -1 ? findCol('price per day') : findCol('day');
    const otHoursCol = findCol('ot hours') !== -1 ? findCol('ot hours') : findCol('overtime hours');
    const otDaysCol = findCol('ot days') !== -1 ? findCol('ot days') : findCol('overtime days');
    const totalOTCol = findCol('total ot payable') !== -1 ? findCol('total ot payable') : findCol('ot payable');
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const record = {
        'Full Name': nameCol !== -1 ? String(data[i][nameCol] || '').trim() : '',
        'Civil ID Number': civilIdCol !== -1 ? String(data[i][civilIdCol] || '').trim() : '',
        'Basic salary': basicCol !== -1 ? data[i][basicCol] || 0 : 0,
        'Food': foodCol !== -1 ? data[i][foodCol] || 0 : 0,
        'Accomodation Allowance': accomCol !== -1 ? data[i][accomCol] || 0 : 0,
        'Conveyance Allowance': conveyCol !== -1 ? data[i][conveyCol] || 0 : 0,
        'Bonus': bonusCol !== -1 ? data[i][bonusCol] || 0 : 0,
        'Gross Salary': grossCol !== -1 ? data[i][grossCol] || 0 : 0,
        'Loan': loanCol !== -1 ? data[i][loanCol] || 0 : 0,
        'Other Deductions': deductionCol !== -1 ? data[i][deductionCol] || 0 : 0,
        'Total Payable Salary': totalPayableCol !== -1 ? data[i][totalPayableCol] || 0 : 0,
        'الفرع': branchCol !== -1 ? data[i][branchCol] || '' : '',
        'Price Per Hour': priceHourCol !== -1 ? data[i][priceHourCol] || 0 : 0,
        'Price Per Day': priceDayCol !== -1 ? data[i][priceDayCol] || 0 : 0,
        'OT Hours': otHoursCol !== -1 ? data[i][otHoursCol] || 0 : 0,
        'OT Days': otDaysCol !== -1 ? data[i][otDaysCol] || 0 : 0,
        'Total OT Payable': totalOTCol !== -1 ? data[i][totalOTCol] || 0 : 0
      };
      for (let j = 0; j < headers.length; j++) { record[headers[j]] = data[i][j]; }
      payroll.push(record);
    }
    return { success: true, data: payroll };
  } catch (error) { return { success: false, error: error.toString(), data: [] }; }
}

// ============================================================
// 3. GET VACATION DATA FROM PAYROLL SHEET
// ============================================================
function getVacationData() {
  try {
    const sheet = SpreadsheetApp.openById(PAYROLL_SPREADSHEET_ID).getSheetByName('Vacation');
    if (!sheet) { return { success: true, data: [] }; }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) { return { success: true, data: [] }; }
    const headers = data[0];
    const vacation = [];
    const findCol = (name) => {
      for (let i = 0; i < headers.length; i++) {
        if (headers[i] && headers[i].toString().trim().toLowerCase().includes(name.toLowerCase())) { return i; }
      }
      return -1;
    };
    const nameCol = findCol('english') !== -1 ? findCol('english') : findCol('Name');
    const takenCol = findCol('taken') !== -1 ? findCol('taken') : findCol('used');
    const remainingCol = findCol('remaining') !== -1 ? findCol('remaining') : findCol('left');
    const remarksCol = findCol('remarks');
    for (let i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      const record = {
        'Name in English': nameCol !== -1 ? String(data[i][nameCol] || '').trim() : '',
        'Vacation taken days': takenCol !== -1 ? data[i][takenCol] || 0 : 0,
        'Remaining Vacation days': remainingCol !== -1 ? data[i][remainingCol] || 0 : 0,
        'Remarks': remarksCol !== -1 ? data[i][remarksCol] || '' : ''
      };
      for (let j = 0; j < headers.length; j++) { record[headers[j]] = data[i][j]; }
      vacation.push(record);
    }
    return { success: true, data: vacation };
  } catch (error) { return { success: false, error: error.toString(), data: [] }; }
}

// ============================================================
// 4. GET SICK LEAVE DATA FROM PAYROLL SHEET
// ============================================================
function getSickLeaveData() {
  try {
    const sheet = SpreadsheetApp.openById(PAYROLL_SPREADSHEET_ID).getSheetByName('Sick Leave');
    if (!sheet) { return { success: true, data: [] }; }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) { return { success: true, data: [] }; }
    const headers = data[0];
    const sick = [];
    const findCol = (name) => {
      for (let i = 0; i < headers.length; i++) {
        if (headers[i] && headers[i].toString().trim().toLowerCase().includes(name.toLowerCase())) { return i; }
      }
      return -1;
    };
    const nameCol = findCol('english') !== -1 ? findCol('english') : findCol('Name');
    const daysCol = findCol('sickness') !== -1 ? findCol('sickness') : findCol('days');
    const remarksCol = findCol('remarks');
    for (let i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      const record = {
        'Name (English)': nameCol !== -1 ? String(data[i][nameCol] || '').trim() : '',
        'Sickness days': daysCol !== -1 ? data[i][daysCol] || 0 : 0,
        'Remarks': remarksCol !== -1 ? data[i][remarksCol] || '' : ''
      };
      for (let j = 0; j < headers.length; j++) { record[headers[j]] = data[i][j]; }
      sick.push(record);
    }
    return { success: true, data: sick };
  } catch (error) { return { success: false, error: error.toString(), data: [] }; }
}

// ============================================================
// 5. INTEGRATED DATA - MERGES ALL SOURCES
// ============================================================
function getIntegratedEmployees() {
  try {
    const empResult = getEmployees();
    if (!empResult.success) { return { success: false, error: empResult.error, data: [] }; }
    const employees = empResult.data || [];
    const vacationResult = getVacationData();
    const vacation = vacationResult.data || [];
    const sickResult = getSickLeaveData();
    const sick = sickResult.data || [];
    const integrated = employees.map(emp => {
      const empName = emp['Name (English)'] || '';
      const vacationRecord = vacation.find(v => String(v['Name in English'] || '').trim() === empName.trim());
      const sickRecord = sick.find(s => String(s['Name (English)'] || '').trim() === empName.trim());
      if (vacationRecord) {
        emp['Vacation Taken'] = vacationRecord['Vacation taken days'] || 0;
        emp['Remaining Vacation'] = vacationRecord['Remaining Vacation days'] || 0;
        emp['Vacation Remarks'] = vacationRecord['Remarks'] || '';
      }
      if (sickRecord) {
        emp['Sick Days Taken'] = sickRecord['Sickness days'] || 0;
        emp['Sick Remarks'] = sickRecord['Remarks'] || '';
      }
      return emp;
    });
    const assigned = assignEmployeeIds(integrated);
    return { success: true, data: assigned };
  } catch (error) { return { success: false, error: error.toString(), data: [] }; }
}

// ============================================================
// 6. ASSIGN EMPLOYEE IDs
// ============================================================
function assignEmployeeIds(employees) {
  const managementNames = [
    'Abdulrahman Al-Najjar', 'Mishari Al-Mutawaq', 'Abdulwahab Al-Najjar',
    'Abdullah Al-Samhan', 'Mohamed Awda', 'Ayman Mandoub'
  ];
  const management = [];
  const regular = [];
  for (const emp of employees) {
    const name = emp['Name (English)'] || '';
    const isManagement = managementNames.some(m => name.includes(m) || m.includes(name));
    if (isManagement) { management.push(emp); } else { regular.push(emp); }
  }
  management.sort((a, b) => {
    const aName = a['Name (English)'] || '';
    const bName = b['Name (English)'] || '';
    const aIndex = managementNames.findIndex(m => aName.includes(m) || m.includes(aName));
    const bIndex = managementNames.findIndex(m => bName.includes(m) || m.includes(bName));
    return aIndex - bIndex;
  });
  regular.sort((a, b) => {
    const dateA = new Date(a['Date of Join'] || 0);
    const dateB = new Date(b['Date of Join'] || 0);
    return dateA - dateB;
  });
  let idCounter = 1;
  for (const emp of management) { emp['Employee ID'] = `MT-${String(idCounter).padStart(5, '0')}`; idCounter++; }
  for (const emp of regular) { emp['Employee ID'] = `MT-${String(idCounter).padStart(5, '0')}`; idCounter++; }
  return employees;
}

// ============================================================
// 7. ATTENDANCE FUNCTIONS
// ============================================================
function getAttendance(params) {
  try {
    const sheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).getSheetByName('Attendance');
    if (!sheet) { return { success: true, data: [] }; }
    const allData = sheet.getDataRange().getValues();
    if (allData.length <= 1) { return { success: true, data: [] }; }
    const headers = allData[0];
    const attendance = [];
    const startDate = params.startDate;
    const endDate = params.endDate;
    for (let i = 1; i < allData.length; i++) {
      const date = allData[i][0];
      if (date && date >= startDate && date <= endDate) {
        const record = {};
        for (let j = 0; j < headers.length; j++) { record[headers[j]] = allData[i][j]; }
        attendance.push(record);
      }
    }
    return { success: true, data: attendance };
  } catch (error) { return { success: false, error: error.toString(), data: [] }; }
}

function markAttendance(data) {
  try {
    let sheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).getSheetByName('Attendance');
    if (!sheet) {
      sheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).insertSheet('Attendance');
      sheet.appendRow(['Date', 'Employee ID', 'Employee Name', 'Civil ID', 'Check In', 'Check Out', 'Hours', 'Status']);
    }
    const row = [
      data.date || new Date().toISOString().split('T')[0],
      data.employeeId || '', data.employeeName || '', data.civilId || '',
      data.checkIn || new Date().toLocaleTimeString(), data.checkOut || '',
      data.hours || '', data.status || 'Present'
    ];
    sheet.appendRow(row);
    return { success: true, message: 'Attendance marked' };
  } catch (error) { return { success: false, error: error.toString() }; }
}

// ============================================================
// 8. LEAVE REQUEST FUNCTIONS
// ============================================================
function getLeaveRequests() {
  try {
    const sheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).getSheetByName('Leave Requests');
    if (!sheet) { return { success: true, data: [] }; }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) { return { success: true, data: [] }; }
    const headers = data[0];
    const leaves = [];
    for (let i = 1; i < data.length; i++) {
      const leave = {};
      for (let j = 0; j < headers.length; j++) { leave[headers[j]] = data[i][j]; }
      leaves.push(leave);
    }
    return { success: true, data: leaves };
  } catch (error) { return { success: false, error: error.toString(), data: [] }; }
}

function submitLeave(data) {
  try {
    let sheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).getSheetByName('Leave Requests');
    if (!sheet) {
      sheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).insertSheet('Leave Requests');
      sheet.appendRow(['Request ID', 'Employee ID', 'Employee Name', 'Civil ID', 'Start Date', 'End Date', 'Type', 'Status', 'Reason']);
    }
    const nextId = sheet.getLastRow();
    const row = [
      nextId, data.employeeId || '', data.employeeName || '', data.civilId || '',
      data.startDate || '', data.endDate || '', data.type || 'Annual',
      data.status || 'Pending', data.reason || ''
    ];
    sheet.appendRow(row);
    return { success: true, message: 'Leave request submitted', id: nextId };
  } catch (error) { return { success: false, error: error.toString() }; }
}

function updateLeaveStatus(data) {
  try {
    const sheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).getSheetByName('Leave Requests');
    if (!sheet) { return { success: false, error: 'Leave sheet not found' }; }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const requestId = data.requestId;
    const idColumn = headers.indexOf('Request ID') + 1;
    const dataRange = sheet.getDataRange().getValues();
    let rowNum = -1;
    for (let i = 1; i < dataRange.length; i++) {
      if (String(dataRange[i][idColumn-1]) === String(requestId)) { rowNum = i + 1; break; }
    }
    if (rowNum === -1) { return { success: false, error: 'Leave request not found' }; }
    const statusCol = headers.indexOf('Status') + 1;
    sheet.getRange(rowNum, statusCol).setValue(data.status);
    return { success: true, message: 'Leave status updated' };
  } catch (error) { return { success: false, error: error.toString() }; }
}

// ============================================================
// 9. PERFORMANCE REVIEW FUNCTIONS
// ============================================================
function getReviews() {
  try {
    const sheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).getSheetByName('Performance Reviews');
    if (!sheet) { return { success: true, data: [] }; }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) { return { success: true, data: [] }; }
    const headers = data[0];
    const reviews = [];
    for (let i = 1; i < data.length; i++) {
      const review = {};
      for (let j = 0; j < headers.length; j++) { review[headers[j]] = data[i][j]; }
      reviews.push(review);
    }
    return { success: true, data: reviews };
  } catch (error) { return { success: false, error: error.toString(), data: [] }; }
}

function addReview(data) {
  try {
    let sheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).getSheetByName('Performance Reviews');
    if (!sheet) {
      sheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).insertSheet('Performance Reviews');
      sheet.appendRow(['Review ID', 'Employee ID', 'Employee Name', 'Civil ID', 'Review Date', 'Reviewer', 'Rating', 'Comments', 'Goals']);
    }
    const nextId = sheet.getLastRow();
    const row = [
      nextId, data.employeeId || '', data.employeeName || '', data.civilId || '',
      data.reviewDate || new Date().toISOString().split('T')[0], data.reviewer || '',
      data.rating || 0, data.comments || '', data.goals || ''
    ];
    sheet.appendRow(row);
    return { success: true, message: 'Review added', id: nextId };
  } catch (error) { return { success: false, error: error.toString() }; }
}

// ============================================================
// 10. DASHBOARD STATS
// ============================================================
function getDashboardStats() {
  try {
    const employees = getIntegratedEmployees().data || [];
    const today = new Date().toISOString().split('T')[0];
    const attendance = getAttendance({ startDate: today, endDate: today }).data || [];
    const presentToday = attendance.filter(a => a.Status === 'Present').length;
    const leaves = getLeaveRequests().data || [];
    const pendingLeaves = leaves.filter(l => l.Status === 'Pending').length;
    const reviews = getReviews().data || [];
    let avgRating = 0;
    if (reviews.length > 0) {
      const sum = reviews.reduce((acc, r) => acc + parseFloat(r.Rating || 0), 0);
      avgRating = sum / reviews.length;
    }
    const activeEmployees = employees.filter(e => e.Status === 'Active').length;
    return {
      success: true,
      data: {
        totalEmployees: employees.length, activeEmployees: activeEmployees,
        presentToday: presentToday, pendingLeaves: pendingLeaves,
        avgRating: avgRating.toFixed(1),
        recentAttendance: attendance.slice(-5).reverse(),
        pendingLeaveRequests: leaves.filter(l => l.Status === 'Pending').slice(0, 5)
      }
    };
  } catch (error) { return { success: false, error: error.toString() }; }
}

// ============================================================
// 11. HELPER - WRITE DATA TO SHEET TAB
// ============================================================
function writeIntegratedDataToSheet() {
  try {
    const result = getIntegratedEmployees();
    if (!result.success || !result.data || result.data.length === 0) { Logger.log("No data found to write."); return; }
    const employees = result.data;
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Master Data');
    if (!sheet) { sheet = ss.insertSheet('Master Data'); }
    const headers = Object.keys(employees[0]);
    const rows = [headers];
    employees.forEach(emp => {
      const row = headers.map(header => { let val = emp[header]; return (val !== undefined && val !== null) ? val : ''; });
      rows.push(row);
    });
    sheet.clear();
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.autoResizeColumns(1, headers.length);
    Logger.log(`Successfully wrote ${employees.length} employees to the sheet.`);
    return { success: true, message: `Written ${employees.length} records.` };
  } catch (error) { Logger.log("Error writing to sheet: " + error.toString()); return { success: false, error: error.toString() }; }
}

// ============================================================
// 12. WRITE CLEAN INTEGRATED DATA TO 'HR Maya' SHEET
// ============================================================
function writeToHrMaya() {
  try {
    const result = getIntegratedEmployees();
    if (!result.success || !result.data || result.data.length === 0) { Logger.log("No data found to write."); return; }
    const employees = result.data;
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    let sheet = ss.getSheetByName('HR Maya');
    if (!sheet) { sheet = ss.insertSheet('HR Maya'); }
    const cleanHeaders = [
      'Employee ID', 'Name (English)', 'Name (Arabic)', 'Civil ID', 'Basic Salary',
      'Food Allowance', 'Accommodation', 'Conveyance', 'Bonus', 'Gross Salary',
      'Loan', 'Other Deductions', 'Total Payable', 'Branch', 'Status',
      'Price Per Hour', 'Price Per Day', 'OT Hours', 'OT Days', 'Total OT Payable',
      'Vacation Taken', 'Remaining Vacation', 'Vacation Remarks', 'Sick Days Taken', 'Sick Remarks'
    ];
    const rows = [cleanHeaders];
    employees.forEach(emp => {
      const row = cleanHeaders.map(header => { let val = emp[header]; return (val !== undefined && val !== null) ? val : ''; });
      rows.push(row);
    });
    sheet.clear();
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sheet.getRange(1, 1, 1, cleanHeaders.length).setFontWeight("bold");
    sheet.autoResizeColumns(1, cleanHeaders.length);
    Logger.log(`Successfully wrote ${employees.length} clean records to HR Maya.`);
    return { success: true, message: `Written ${employees.length} clean records.` };
  } catch (error) { Logger.log("Error writing to sheet: " + error.toString()); return { success: false, error: error.toString() }; }
}

// ============================================================
// 13. KIOSK LOGIN - READS COLUMN B DIRECTLY
// ============================================================
function loginStaffKiosk(params) {
  try {
    const empId = String(params.empId).trim();
    const pin = String(params.pin).trim();
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HR Maya');
    if (!sheet) return { success: false, error: 'System error' };
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, error: 'No data found' };
    // Resolve columns by HEADER (immune to column re-ordering / added columns).
    const H = data[0];
    const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cPin = H.findIndex(h => String(h).trim().toLowerCase() === 'pin');
    const cName = H.findIndex(h => String(h).trim().toLowerCase() === 'name (english)');
    if (cId === -1 || cPin === -1) return { success: false, error: 'HR Maya is missing Employee ID or PIN column.' };
    let foundEmployee = null;
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][cId]).trim();
      // PIN may be stored as a number; normalize both sides (strip trailing .0).
      let rowPin = String(data[i][cPin]).trim(); if (rowPin.endsWith('.0')) rowPin = rowPin.slice(0, -2);
      if (rowId === empId && rowPin === pin) {
        foundEmployee = { id: rowId, name: cName !== -1 ? String(data[i][cName]).trim() : rowId };
        break;
      }
    }
    if (!foundEmployee) { return { success: false, error: 'Invalid Employee ID or PIN' }; }
    const statusResult = getDailyKioskStatus({ empId: empId });
    return { success: true, name: foundEmployee.name, id: foundEmployee.id, currentStatus: statusResult.currentStage || 'none' };
  } catch (error) { return { success: false, error: error.toString() }; }
}

// ============================================================
// 14. KIOSK MARK ATTENDANCE
// ============================================================
// SESSION HELPERS (overnight-safe) — a session is Check In -> Check Out,
// spanning any dates. All kiosk stage logic derives from these.
// ============================================================
// How many hours an un-closed session keeps the kiosk locked to its next
// stage. Deliberately separate from AC_MAX_SESSION_HOURS (20) — see the
// note in the session builder for why 17.
const AC_STALE_OPEN_HOURS = 17;

function kioskSessionState_(attData, empId) {
  const H = attData[0];
  const cId = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
  const cStage = H.findIndex(h => String(h).trim().toLowerCase() === 'stage');
  const cTs = H.findIndex(h => String(h).trim().toLowerCase() === 'timestamp');
  const punches = [];
  for (let i = 1; i < attData.length; i++) {
    if (String(attData[i][cId]).trim() !== empId) continue;
    const ts = (typeof parseLogTimestamp === 'function') ? parseLogTimestamp(attData[i][cTs]) : null;
    if (!ts) continue;
    punches.push({ ts: ts, stage: String(attData[i][cStage]).trim() });
  }
  punches.sort((a, b) => a.ts - b.ts);
  // A punch in the future has not happened. On 20-09 supervisors wrote
  // check-outs at shift ends that had not come yet; the kiosk read them as
  // "already checked out", refused the man's real Check Out, and his next
  // tap went in as a Check In. Ignore them.
  const now = new Date();
  // How long an un-closed session keeps the kiosk locked to its next
  // stage. This is NOT AC_MAX_SESSION_HOURS, which asks a different
  // question ("are these punches plausible?") and must stay at 20.
  //
  // Two things have to hold at once:
  //   • a man at the end of a real double shift — 15-16h, 12:26 → 03:30 —
  //     must still be able to Check OUT, so the session cannot expire
  //     under about 16 hours;
  //   • a man who forgot to check out last night must be able to Check IN
  //     for today's shift. The tightest real gap is Jabriya/Qurain second
  //     shift: in at 18:30, back at the kiosk 12:28 next day — 17.9h.
  //
  // 17h sits in that gap. At 20h a second-shift man who forgets his
  // check-out is locked out of his next shift until mid-afternoon, which
  // is what happened on 01-10: the kiosk kept showing him Break In.
  const staleH = (typeof AC_STALE_OPEN_HOURS === 'number') ? AC_STALE_OPEN_HOURS : 17;
  let open = false, stage = 'none', openTs = null;
  for (const p of punches) {
    if (p.ts > now) continue;
    if (p.stage === 'Check In') {
      // A Check In while a shift is open is a repeat tap — UNLESS that shift
      // has been open 20h or more: then it is a forgotten check-out from an
      // earlier day and this Check In starts the new shift. (Before 22-09 the
      // repeat rule swallowed these too, so one missed check-out locked the
      // man out for good: every later Check In was ignored and the kiosk kept
      // showing the first screen — Kiran Bhujel 20-09 → 22-09.)
      if (open && openTs && (p.ts - openTs) / 3600000 < staleH) continue;
      open = true; stage = 'checked_in'; openTs = p.ts;
    }
    else if (!open) { continue; }
    else if (p.stage === 'Break Out') { stage = 'on_break'; }
    else if (p.stage === 'Break In') { stage = 'break_done'; }
    else if (p.stage === 'Check Out') { open = false; stage = 'none'; openTs = null; }
  }
  // A check-in left open for 20 hours or more is a forgotten check-out from
  // an earlier shift, not the shift he is standing at the kiosk for. Let him
  // check in; the old day goes on his supervisor's list to be answered.
  if (open && openTs && (now - openTs) / 3600000 >= staleH) {
    return { stage: 'none', open: false, openCheckInTs: null, staleOpenTs: openTs };
  }
  return { stage: stage, open: open, openCheckInTs: openTs };
}

function validateStageAgainstSession_(sess, stage) {
  if (stage === 'Check In') {
    if (sess.open) return 'Your previous shift was not checked out. Please contact HR to close it before checking in again.';
    return '';
  }
  if (!sess.open) return 'You are not checked in. Please Check In first.';
  if (stage === 'Break Out') {
    if (sess.stage === 'on_break') return 'You are already on break.';
    return '';
  }
  if (stage === 'Break In') {
    if (sess.stage !== 'on_break') return 'You are not on break.';
    return '';
  }
  if (stage === 'Check Out') {
    if (sess.stage === 'on_break') return 'Please Break In before checking out.';
    return '';
  }
  return '';
}

// ── Helper: normalise any date cell to dd-mm-yyyy ────────────────────
function normDayStr_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return formatDdMmYyyy_(v);
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return m[1].padStart(2,'0') + '-' + m[2].padStart(2,'0') + '-' + m[3];
  m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return m[3].padStart(2,'0') + '-' + m[2].padStart(2,'0') + '-' + m[1];
  var d = new Date(s);
  return isNaN(d.getTime()) ? s : formatDdMmYyyy_(d);
}

function markKioskAttendance(params) {
  try {
    const empId = String(params.empId).trim();
    const stage = String(params.stage).trim();

    // ── FIX: 'Day Off' is now a valid stage ───────────────────────────
    // Today's Day Off  → no date param; uses today.
    // Backdated Day Off → params.date = dd-mm-yyyy (must be in the past).
    if (!['Check In','Break Out','Break In','Check Out','Day Off'].includes(stage)) {
      return { success: false, error: 'Invalid stage' };
    }

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HR Maya');
    const data = sheet.getDataRange().getValues();
    const H = data[0];
    let cBranch = H.findIndex(h => String(h).trim().toLowerCase() === 'location');
    if (cBranch === -1) cBranch = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
    let empName = 'Unknown', empBranch = '';
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === empId) {
        empName   = String(data[i][2]).trim();
        empBranch = cBranch !== -1 ? String(data[i][cBranch] || '').trim() : '';
        break;
      }
    }

    // --- Geofence (allow but flag) ---
    const geo = checkGeofence_(ss, empBranch, params.lat, params.lng);

    const now = new Date();
    let punchDate = formatDdMmYyyy_(now); // default: today

    // Backdated Day Off: honour params.date if it is a valid past date
    if (stage === 'Day Off' && params.date) {
      const d = String(params.date).trim();
      if (/^\d{2}-\d{2}-\d{4}$/.test(d)) {
        const parts = d.split('-');
        const dObj  = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
        const today0 = new Date(); today0.setHours(0, 0, 0, 0);
        if (dObj >= today0) return { success: false, error: 'Backdated Day Off must be for a past date.' };
        punchDate = d;
      }
    }

    const timestamp = now.toLocaleString('en-GB', { hour12: false }).replace(',', '');
    let attSheet = ss.getSheetByName('Attendance_Log');
    if (!attSheet) {
      attSheet = ss.insertSheet('Attendance_Log');
      attSheet.appendRow(['Timestamp', 'Employee ID', 'Employee Name', 'Date', 'Stage', 'Auth', 'Branch', 'Latitude', 'Longitude', 'Distance (m)', 'Geo Flag']);
    }
    ensureAttendanceGeoCols_(attSheet);

    const existingData = attSheet.getDataRange().getValues();

    if (stage === 'Day Off') {
      // Day Off: no session validation needed. Block duplicate Day Off for same date.
      const H2      = existingData[0];
      const cId2    = H2.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
      const cDate2  = H2.findIndex(h => String(h).trim().toLowerCase() === 'date');
      const cStage2 = H2.findIndex(h => String(h).trim().toLowerCase() === 'stage');
      for (let i = 1; i < existingData.length; i++) {
        if (String(existingData[i][cId2]).trim() === empId &&
            normDayStr_(existingData[i][cDate2]) === punchDate &&
            String(existingData[i][cStage2]).trim() === 'Day Off') {
          return { success: false, error: 'Day Off already recorded for ' + punchDate + '.' };
        }
      }
    } else {
      // Punch stages: session-based validation (overnight-safe).
      const sess = kioskSessionState_(existingData, empId);
      const stageErr = validateStageAgainstSession_(sess, stage);
      if (stageErr) return { success: false, error: stageErr };
    }

    // Auth and Branch can be supplied by another front door (the QR terminal
    // in QrPunch.gs passes 'QR' and the tablet's own location). Left out, this
    // behaves exactly as it always has for the phone kiosk.
    const authLabel = String(params.auth || '').trim() || 'Verified';
    const branchLabel = String(params.branchOverride || '').trim() || (geo.nearest || empBranch);
    attSheet.appendRow([timestamp, empId, empName, punchDate, stage, authLabel, branchLabel,
      geo.lat === null ? '' : geo.lat, geo.lng === null ? '' : geo.lng,
      geo.distance === null ? '' : geo.distance, geo.flag]);

    return { success: true, message: `${stage} recorded successfully`, stage: stage,
      geoFlag: geo.flag, distance: geo.distance, insideFence: geo.inside };
  } catch (error) { return { success: false, error: error.toString() }; }
}

// ============================================================
// GEOFENCE HELPERS — Branch_Locations tab + Haversine distance.
// ============================================================
function ensureAttendanceGeoCols_(attSheet) {
  const H = attSheet.getRange(1, 1, 1, attSheet.getLastColumn()).getValues()[0];
  const need = ['Branch', 'Latitude', 'Longitude', 'Distance (m)', 'Geo Flag'];
  need.forEach(name => {
    if (H.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase()) === -1) {
      attSheet.getRange(1, attSheet.getLastColumn() + 1).setValue(name).setFontWeight('bold');
    }
  });
}

function haversineMetres_(lat1, lon1, lat2, lon2) {
  const R = 6371000, toRad = (d) => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1), dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
  return Math.round(2 * R * Math.asin(Math.sqrt(a)));
}

function checkGeofence_(ss, branch, latRaw, lngRaw) {
  const lat = (latRaw === undefined || latRaw === '' || latRaw === null) ? null : Number(latRaw);
  const lng = (lngRaw === undefined || lngRaw === '' || lngRaw === null) ? null : Number(lngRaw);
  if (lat === null || lng === null || isNaN(lat) || isNaN(lng)) {
    return { lat: null, lng: null, distance: null, inside: false, flag: 'NO LOCATION', nearest: '' };
  }
  const loc = ss.getSheetByName('Branch_Locations');
  if (!loc) return { lat: lat, lng: lng, distance: null, inside: false, flag: 'NO BRANCH TABLE', nearest: '' };
  const d = loc.getDataRange().getValues();
  const H = d[0];
  const cB = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
  const cLa = H.findIndex(h => String(h).trim().toLowerCase() === 'latitude');
  const cLo = H.findIndex(h => String(h).trim().toLowerCase() === 'longitude');
  const cR = H.findIndex(h => String(h).trim().toLowerCase() === 'radius');
  const cG = H.findIndex(h => String(h).trim().toLowerCase() === 'geofenced');

  let empLocGeofenced = true;
  if (branch) {
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][cB] || '').trim().toLowerCase() === String(branch).trim().toLowerCase()) {
        empLocGeofenced = (cG === -1) ? true : String(d[i][cG]).trim().toLowerCase() === 'yes';
        break;
      }
    }
  }

  let nearestAny = null;
  for (let i = 1; i < d.length; i++) {
    const bLat = Number(d[i][cLa]), bLng = Number(d[i][cLo]);
    if (isNaN(bLat) || isNaN(bLng)) continue;
    const dist = haversineMetres_(lat, lng, bLat, bLng);
    if (!nearestAny || dist < nearestAny.dist) nearestAny = { dist: dist, name: String(d[i][cB] || '').trim() };
  }

  if (!empLocGeofenced) {
    return { lat: lat, lng: lng, distance: nearestAny ? nearestAny.dist : null, inside: true, nearest: nearestAny ? nearestAny.name : '',
      flag: nearestAny ? ('NEAREST: ' + nearestAny.name + ' (' + nearestAny.dist + 'm) — not fenced') : 'RECORDED — not fenced' };
  }

  let best = null;
  for (let i = 1; i < d.length; i++) {
    if (cG !== -1 && String(d[i][cG]).trim().toLowerCase() !== 'yes') continue;
    const bLat = Number(d[i][cLa]), bLng = Number(d[i][cLo]);
    if (isNaN(bLat) || isNaN(bLng)) continue;
    const radius = (cR !== -1 && Number(d[i][cR])) ? Number(d[i][cR]) : 50;
    const dist = haversineMetres_(lat, lng, bLat, bLng);
    if (!best || dist < best.dist) best = { dist: dist, radius: radius, name: String(d[i][cB] || '').trim() };
  }
  if (!best) return { lat: lat, lng: lng, distance: null, inside: false, flag: 'NO BRANCHES CONFIGURED', nearest: '' };

  const inside = best.dist <= best.radius;
  return {
    lat: lat, lng: lng, distance: best.dist, inside: inside, nearest: best.name,
    flag: inside ? ('OK (' + best.name + ')') : ('OUTSIDE ALL (nearest ' + best.name + ' ' + best.dist + 'm)')
  };
}

function recomputeGeoFlags(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const log = ss.getSheetByName('Attendance_Log');
    if (!log) return { success: false, error: 'No Attendance_Log.' };
    const data = log.getDataRange().getValues();
    const H = data[0];
    const cLat = H.findIndex(h => String(h).trim().toLowerCase() === 'latitude');
    const cLng = H.findIndex(h => String(h).trim().toLowerCase() === 'longitude');
    const cDist = H.findIndex(h => String(h).trim().toLowerCase() === 'distance (m)');
    const cFlag = H.findIndex(h => String(h).trim().toLowerCase() === 'geo flag');
    if (cLat === -1 || cLng === -1 || cFlag === -1) return { success: false, error: 'Attendance_Log missing Latitude/Longitude/Geo Flag columns.' };
    let updated = 0, stillNoLoc = 0;
    for (let i = 1; i < data.length; i++) {
      if (!String(data[i][1] || '').trim()) continue;
      const lat = data[i][cLat], lng = data[i][cLng];
      if (lat === '' || lng === '' || lat == null || lng == null) { stillNoLoc++; continue; }
      const geo = checkGeofence_(ss, '', lat, lng);
      log.getRange(i + 1, cFlag + 1).setValue(geo.flag);
      if (cDist !== -1) log.getRange(i + 1, cDist + 1).setValue(geo.distance === null ? '' : geo.distance);
      updated++;
    }
    return { success: true, updated: updated, stillNoLocation: stillNoLoc,
      message: 'Recomputed Geo Flag on ' + updated + ' row(s) with coordinates. ' + stillNoLoc + ' still without coordinates (left as NO LOCATION).' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function setupBranchLocations() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Branch_Locations');
  const seed = [
    ['Kaifan', 29.341500, 47.951750, 50, 'Yes', 'Yes'],
    ['Jabriya', 29.317139, 48.028611, 50, 'Yes', 'Yes'],
    ['Ardiya', 29.297167, 47.910611, 50, 'Yes', 'Yes'],
    ['Qurain', 29.198861, 48.052833, 50, 'Yes', 'Yes'],
    ['Khairan', 28.678537, 48.332356, 50, 'Yes', 'Yes'],
    ['Central Kitchen', '', '', '', 'No', 'Yes'],
    ['Head Office', '', '', '', 'No', 'Yes']
  ];
  if (!sheet) { sheet = ss.insertSheet('Branch_Locations'); }
  sheet.clear();
  sheet.appendRow(['Branch', 'Latitude', 'Longitude', 'Radius', 'Geofenced', 'Active']);
  sheet.getRange(2, 1, seed.length, 6).setValues(seed);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return { success: true, message: 'Branch_Locations seeded: 5 geofenced branches + Central Kitchen + Head Office (not geofenced). Manage locations on the Locations page.' };
}

function ensureLocationCols_(ss) {
  let sheet = ss.getSheetByName('Branch_Locations');
  if (!sheet) return setupBranchLocations(), ss.getSheetByName('Branch_Locations');
  const H = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const has = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase()) !== -1;
  if (!has('Geofenced')) {
    const col = sheet.getLastColumn() + 1;
    sheet.getRange(1, col).setValue('Geofenced').setFontWeight('bold');
    const n = sheet.getLastRow() - 1;
    if (n > 0) sheet.getRange(2, col, n, 1).setValue('Yes');
  }
  if (!has('Active')) {
    const col = sheet.getLastColumn() + 1;
    sheet.getRange(1, col).setValue('Active').setFontWeight('bold');
    const n = sheet.getLastRow() - 1;
    if (n > 0) sheet.getRange(2, col, n, 1).setValue('Yes');
  }
  return sheet;
}

function getLocations(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureLocationCols_(ss);
    const d = sheet.getDataRange().getValues();
    const H = d[0];
    const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const cB = c('Branch'), cLa = c('Latitude'), cLo = c('Longitude'), cR = c('Radius'), cG = c('Geofenced'), cA = c('Active'), cSh = c('Shifts');
    const out = [];
    for (let i = 1; i < d.length; i++) {
      const name = String(d[i][cB] || '').trim();
      if (!name) continue;
      out.push({
        name: name,
        latitude: cLa !== -1 && d[i][cLa] !== '' ? Number(d[i][cLa]) : '',
        longitude: cLo !== -1 && d[i][cLo] !== '' ? Number(d[i][cLo]) : '',
        radius: cR !== -1 && d[i][cR] !== '' ? Number(d[i][cR]) : '',
        geofenced: cG !== -1 ? String(d[i][cG]).trim().toLowerCase() === 'yes' : false,
        shifts: cSh !== -1 ? String(d[i][cSh] || '').split(',').map(s => s.trim()).filter(Boolean) : [],
        active: cA === -1 || String(d[i][cA]).trim().toLowerCase() === 'yes'
      });
    }
    return { success: true, data: out };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function saveLocation(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureLocationCols_(ss);
    const name = String(params.name || '').trim();
    if (!name) return { success: false, error: 'Location name required.' };
    const geofenced = String(params.geofenced).trim().toLowerCase() === 'yes' || params.geofenced === true;
    if (geofenced && (params.latitude === '' || params.longitude === '' || params.latitude == null || params.longitude == null)) {
      return { success: false, error: 'Geofenced locations need latitude and longitude.' };
    }
    const d = sheet.getDataRange().getValues();
    const H = d[0];
    const c = (n) => H.findIndex(h => String(h).trim().toLowerCase() === n.toLowerCase());
    const cB = c('Branch'), cLa = c('Latitude'), cLo = c('Longitude'), cR = c('Radius'), cG = c('Geofenced'), cA = c('Active');
    let rowNum = -1;
    for (let i = 1; i < d.length; i++) { if (String(d[i][cB] || '').trim().toLowerCase() === name.toLowerCase()) { rowNum = i + 1; break; } }
    const width = H.length;
    const row = rowNum === -1 ? new Array(width).fill('') : d[rowNum - 1].slice();
    row[cB] = name;
    row[cLa] = geofenced ? Number(params.latitude) : '';
    row[cLo] = geofenced ? Number(params.longitude) : '';
    row[cR] = geofenced ? (Number(params.radius) || 50) : '';
    if (cG !== -1) row[cG] = geofenced ? 'Yes' : 'No';
    if (cA !== -1) row[cA] = (params.active === undefined ? 'Yes' : (String(params.active).trim().toLowerCase() === 'yes' || params.active === true ? 'Yes' : 'No'));
    if (rowNum === -1) sheet.appendRow(row); else sheet.getRange(rowNum, 1, 1, width).setValues([row]);
    return { success: true, message: (rowNum === -1 ? 'Added' : 'Updated') + ' location "' + name + '"' + (geofenced ? ' (geofenced)' : ' (not geofenced)') + '.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function deleteLocation(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ensureLocationCols_(ss);
    const name = String(params.name || '').trim();
    const d = sheet.getDataRange().getValues();
    const H = d[0];
    const cB = H.findIndex(h => String(h).trim().toLowerCase() === 'branch');
    const cA = H.findIndex(h => String(h).trim().toLowerCase() === 'active');
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][cB] || '').trim().toLowerCase() === name.toLowerCase()) {
        if (String(params.hard).toLowerCase() === 'yes') { sheet.deleteRow(i + 1); return { success: true, message: 'Deleted "' + name + '".' }; }
        sheet.getRange(i + 1, cA + 1).setValue('No');
        return { success: true, message: 'Retired location "' + name + '".' };
      }
    }
    return { success: false, error: 'Location "' + name + '" not found.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ============================================================
// 15. KIOSK GET STATUS
// ============================================================
function getDailyKioskStatus(params) {
  try {
    const empId = String(params.empId).trim();
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Attendance_Log');
    if (!sheet) return { success: true, currentStage: 'none' };
    const data = sheet.getDataRange().getValues();

    // ── FIX: if a Day Off punch exists for today, return 'day_off' ──
    // so the kiosk shows the done screen instead of asking again.
    const todayStr = formatDdMmYyyy_(new Date());
    const H      = data[0];
    const cId    = H.findIndex(h => String(h).trim().toLowerCase() === 'employee id');
    const cDate  = H.findIndex(h => String(h).trim().toLowerCase() === 'date');
    const cStage = H.findIndex(h => String(h).trim().toLowerCase() === 'stage');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][cId]).trim() === empId &&
          String(data[i][cStage]).trim() === 'Day Off' &&
          normDayStr_(data[i][cDate]) === todayStr) {
        return { success: true, currentStage: 'day_off', sessionOpen: false };
      }
    }

    // Session-based (overnight-safe): current stage = state of the open session.
    const sess = kioskSessionState_(data, empId);
    let currentStage;
    if (sess.open) {
      currentStage = sess.stage; // 'checked_in' | 'on_break' | 'break_done'
    } else {
      currentStage = 'none';     // no open session -> next action is Check In
    }
    return { success: true, currentStage: currentStage, sessionOpen: sess.open };
  } catch (error) { return { success: false, error: error.toString() }; }
}