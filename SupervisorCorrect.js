// ============================================================
// SupervisorCorrect.gs — standalone supervisor tool: correct + add punches
// ============================================================
// Scoped to 01–03 Sept 2026 cleanup. A supervisor picks their name, sees their
// LOCATION's sessions in the window, corrects broken ones (Check In/Out/Shift),
// and adds missing ones (from the WA list). All writes go to Attendance_Log
// (source of truth), audited to Attendance_Amendments. Rebuild Attendance_Daily
// afterwards to reflect the fixes.
//
// ENDPOINTS: getSupervisorSessions, correctSession, addSession
// ============================================================

const SC_FROM = '01-09-2026';
const SC_TO   = '03-09-2026';
// Upper cutoff: sessions whose CHECK-IN is before this are included. Excludes
// the 03/09 noon shift that is still in progress (would show as "no check-out").
const SC_CUTOFF = new Date(2026, 8, 3, 11, 0, 0); // 03-09-2026 11:00 (month is 0-based)

function sc_inWindow_(dStr) {
  const d = parseFlexibleDate_(dStr);
  return d && d >= startOfDay_(parseFlexibleDate_(SC_FROM)) && d <= endOfDay_(parseFlexibleDate_(SC_TO));
}
// Include a session by its Check-In timestamp: on/after 01/09 start, and
// strictly before the 03/09 11:00 cutoff.
function sc_checkinInWindow_(ciTs) {
  return ciTs && ciTs >= startOfDay_(parseFlexibleDate_(SC_FROM)) && ciTs < SC_CUTOFF;
}
function sc_normDay_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return formatDdMmYyyy_(v);
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/); if (m) return m[1].padStart(2,'0')+'-'+m[2].padStart(2,'0')+'-'+m[3];
  m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/); if (m) return m[3].padStart(2,'0')+'-'+m[2].padStart(2,'0')+'-'+m[1];
  var d = new Date(s); return isNaN(d.getTime()) ? s : formatDdMmYyyy_(d);
}

// Sessions for a supervisor's location within the cleanup window, plus the
// location's employees + shifts (for the add form). Sessions flagged if broken.
function getSupervisorSessions(params) {
  try {
    params = params || {};
    const location = String(params.location || '').trim();
    if (!location) return { success: false, error: 'location required' };
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);

    // Employees at this location + the location's shifts.
    const hr = ss.getSheetByName('HR Maya'); const hd = hr.getDataRange().getValues(); const H = hd[0];
    const hcId=H.findIndex(h=>String(h).trim().toLowerCase()==='employee id');
    const hcName=H.findIndex(h=>String(h).trim().toLowerCase()==='name (english)');
    let hcLoc=H.findIndex(h=>String(h).trim().toLowerCase()==='location'); if(hcLoc===-1) hcLoc=H.findIndex(h=>String(h).trim().toLowerCase()==='branch');
    const nameById={}, locEmployees=[];
    for(let i=1;i<hd.length;i++){ const id=String(hd[i][hcId]||'').trim(); if(!id) continue;
      nameById[id]=hcName!==-1?String(hd[i][hcName]||'').trim():id;
      if(hcLoc!==-1 && String(hd[i][hcLoc]||'').trim().toLowerCase()===location.toLowerCase()) locEmployees.push({id:id,name:nameById[id]}); }
    locEmployees.sort((a,b)=>a.name.localeCompare(b.name));
    let shifts=[]; try{ const lm=getLocationShiftsMap_(ss); shifts=lm[location.toLowerCase()]||[]; }catch(e){}

    // Sessionize Attendance_Log for this location's employees, in window.
    const log = ss.getSheetByName('Attendance_Log');
    const sessions=[];
    if(log){
      const L=log.getDataRange().getValues(); const LH=L[0];
      const cId=LH.findIndex(h=>String(h).trim().toLowerCase()==='employee id');
      const cStage=LH.findIndex(h=>String(h).trim().toLowerCase()==='stage');
      const cTs=LH.findIndex(h=>String(h).trim().toLowerCase()==='timestamp');
      const cLoc=LH.findIndex(h=>String(h).trim().toLowerCase()==='branch');
      const locEmpIds={}; locEmployees.forEach(e=>locEmpIds[e.id]=true);
      const byEmp={};
      for(let i=1;i<L.length;i++){ const id=String(L[i][cId]||'').trim(); if(!id||!locEmpIds[id]) continue;
        const ts=parseLogTimestamp(L[i][cTs]); if(!ts) continue;
        (byEmp[id]=byEmp[id]||[]).push({row:i+1, ts:ts, stage:String(L[i][cStage]||'').trim()}); }
      Object.keys(byEmp).forEach(id=>{
        const p=byEmp[id].sort((a,b)=>a.ts-b.ts);
        let cur=null;
        p.forEach(x=>{
          if(x.stage==='Check In'){ if(cur)sessions.push(cur); cur={empId:id,name:nameById[id],ciRow:x.row,ciTs:x.ts,coRow:null,coTs:null}; }
          else if(!cur){ return; }
          else if(x.stage==='Check Out'){ cur.coRow=x.row; cur.coTs=x.ts; sessions.push(cur); cur=null; }
        });
        if(cur) sessions.push(cur);
      });
    }
    // Filter to window + shape for UI, flag broken (<1h or no checkout).
    const out=[];
    sessions.forEach(s=>{
      const dateStr=formatDdMmYyyy_(s.ciTs);
      if(!sc_checkinInWindow_(s.ciTs)) return; // check-in must be before 03/09 11:00
      let presence=null, broken=false, reason='';
      if(!s.coTs){ broken=true; reason='No check-out'; }
      else { let mins=(s.coTs-s.ciTs)/60000; if(mins<0) mins+=1440; presence=Math.round((mins/60)*100)/100; if(presence<1){ broken=true; reason='Too short ('+presence+'h)'; } }
      out.push({ empId:s.empId, name:s.name, date:dateStr, ciRow:s.ciRow, coRow:s.coRow,
        checkIn:Utilities.formatDate(s.ciTs,tz_(),'HH:mm'), checkOut:s.coTs?Utilities.formatDate(s.coTs,tz_(),'HH:mm'):'',
        presence:presence, broken:broken, reason:reason });
    });
    out.sort((a,b)=>(a.date+a.name).localeCompare(b.date+b.name));
    // Flag duplicates: same employee + same date appearing more than once.
    const seenCount = {};
    out.forEach(r => { const k = r.empId + '|' + r.date; seenCount[k] = (seenCount[k] || 0) + 1; });
    out.forEach(r => { r.duplicate = (seenCount[r.empId + '|' + r.date] > 1); });
    return { success:true, location:location, window:{from:SC_FROM,to:SC_TO}, sessions:out, employees:locEmployees, shifts:shifts };
  } catch (err) { return { success:false, error: err.toString() }; }
}

// Correct a session's Check In / Check Out times (and optional shift note) by row.
// params: { ciRow, coRow, date, checkIn(HH:mm), checkOut(HH:mm), coNextDay(bool), by }
function correctSession(params) {
  try {
    params = (typeof params==='string')?JSON.parse(params):(params||{});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const log = ss.getSheetByName('Attendance_Log');
    const H = log.getRange(1,1,1,log.getLastColumn()).getValues()[0];
    const cTs=H.findIndex(h=>String(h).trim().toLowerCase()==='timestamp');
    const cAuth=H.findIndex(h=>String(h).trim().toLowerCase()==='auth');
    const d=parseFlexibleDate_(params.date); if(!d) return {success:false,error:'bad date'};
    const mkTs=(hhmm,nextDay)=>{ const base=new Date(d.getTime()+(nextDay?86400000:0)); return Utilities.formatDate(base,tz_(),'dd/MM/yyyy')+' '+hhmm+':00'; };
    const changes=[];
    if(params.ciRow && params.checkIn){ log.getRange(Number(params.ciRow),cTs+1).setValue(mkTs(params.checkIn,false)); if(cAuth!==-1) log.getRange(Number(params.ciRow),cAuth+1).setValue('Supervisor corrected'); changes.push('CheckIn->'+params.checkIn); }
    if(params.coRow && params.checkOut){ log.getRange(Number(params.coRow),cTs+1).setValue(mkTs(params.checkOut, !!params.coNextDay)); if(cAuth!==-1) log.getRange(Number(params.coRow),cAuth+1).setValue('Supervisor corrected'); changes.push('CheckOut->'+params.checkOut+(params.coNextDay?'(+1)':'')); }
    if(typeof logAmend_==='function') logAmend_(ss, params.by, 'CORRECT', params.empId||'', params.date, 'session', changes.join('; '));
    return { success:true, message:'Corrected: '+changes.join('; ') };
  } catch (err) { return { success:false, error: err.toString() }; }
}

// Add a missing session (Check In + Check Out) from the WA list.
// params: { empId, date, shift, checkIn, checkOut, coNextDay, location, by }
function addSession(params) {
  try {
    params = (typeof params==='string')?JSON.parse(params):(params||{});
    const empId=String(params.empId||'').trim();
    const date=String(params.date||'').trim();
    if(!empId||!date||!params.checkIn||!params.checkOut) return {success:false,error:'Employee, date, check-in and check-out required.'};
    if(!sc_inWindow_(date)) return {success:false,error:'Date must be 01–03 Sep 2026.'};
    const ss=SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const log=ss.getSheetByName('Attendance_Log');
    const H=log.getRange(1,1,1,log.getLastColumn()).getValues()[0];
    const cId=H.findIndex(h=>String(h).trim().toLowerCase()==='employee id');
    const cDate=H.findIndex(h=>String(h).trim().toLowerCase()==='date');
    const cStage=H.findIndex(h=>String(h).trim().toLowerCase()==='stage');
    // Duplicate warn: does this emp already have a session that date?
    const data=log.getDataRange().getValues();
    const dd=formatDdMmYyyy_(parseFlexibleDate_(date));
    for(let i=1;i<data.length;i++){ if(String(data[i][cId]).trim()===empId && sc_normDay_(data[i][cDate])===dd && String(data[i][cStage]).trim()==='Check In'){ return {success:false, error:'⚠ '+empId+' already has a session on '+dd+'. Correct it instead of adding.'}; } }

    const name=empNameById_(ss,empId)||empId;
    const loc=String(params.location||'').trim();
    const d=parseFlexibleDate_(date);
    const mkTs=(hhmm,nextDay)=>{ const base=new Date(d.getTime()+(nextDay?86400000:0)); return Utilities.formatDate(base,tz_(),'dd/MM/yyyy')+' '+hhmm+':00'; };
    const mkRow=(stage,tsStr)=>{ const r=new Array(H.length).fill('');
      r[cId]=empId; r[H.findIndex(h=>String(h).trim().toLowerCase()==='employee name')]=name;
      r[cDate]=dd; r[cStage]=stage; r[H.findIndex(h=>String(h).trim().toLowerCase()==='timestamp')]=tsStr;
      const cAuth=H.findIndex(h=>String(h).trim().toLowerCase()==='auth'); if(cAuth!==-1) r[cAuth]='Supervisor added';
      const cBranch=H.findIndex(h=>String(h).trim().toLowerCase()==='branch'); if(cBranch!==-1) r[cBranch]=loc;
      const cFlag=H.findIndex(h=>String(h).trim().toLowerCase()==='geo flag'); if(cFlag!==-1) r[cFlag]='MANUAL (supervisor)';
      return r; };
    log.appendRow(mkRow('Check In', mkTs(params.checkIn,false)));
    log.getRange(log.getLastRow(),cDate+1).setNumberFormat('@');
    log.appendRow(mkRow('Check Out', mkTs(params.checkOut, !!params.coNextDay)));
    log.getRange(log.getLastRow(),cDate+1).setNumberFormat('@');
    if(typeof logAmend_==='function') logAmend_(ss, params.by, 'ADD', empId, dd, 'session', 'in '+params.checkIn+' / out '+params.checkOut+(params.coNextDay?'(+1)':''));
    return { success:true, message:'Added session for '+name+' on '+dd+' ('+params.checkIn+'–'+params.checkOut+').' };
  } catch (err) { return { success:false, error: err.toString() }; }
}

// Delete a session's punch rows (Check In + its Check Out) from Attendance_Log.
// Used to remove a duplicate/erroneous session. Audited.
// params: { ciRow, coRow, empId, date, by }
function deleteSession(params) {
  try {
    params = (typeof params === 'string') ? JSON.parse(params) : (params || {});
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const log = ss.getSheetByName('Attendance_Log');
    if (!log) return { success: false, error: 'No Attendance_Log.' };
    // Collect the rows to delete (Check In row, and Check Out row if present),
    // delete highest row number first so indexes don't shift.
    const rowsToDelete = [];
    if (params.ciRow) rowsToDelete.push(Number(params.ciRow));
    if (params.coRow) rowsToDelete.push(Number(params.coRow));
    rowsToDelete.sort((a, b) => b - a);
    rowsToDelete.forEach(r => { if (r >= 2) log.deleteRow(r); });
    if (typeof logAmend_ === 'function') logAmend_(ss, params.by, 'DELETE', params.empId || '', params.date || '', 'session', 'removed duplicate/erroneous session');
    return { success: true, message: 'Session removed.' };
  } catch (err) { return { success: false, error: err.toString() }; }
}