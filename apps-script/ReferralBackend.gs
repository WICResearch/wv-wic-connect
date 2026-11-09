/**
 * WV WIC Connect referral intake — bound Google Apps Script.
 * Deploy only under agency-approved access restrictions.
 * Do not publish this source with actual sheet IDs, credentials or deployment secrets.
 *
 * Setup: open destination Sheet > Extensions > Apps Script, paste this file,
 * run setupReferralSheet once, then deploy as Web App with approved access.
 * The public GitHub Pages form MUST NOT be enabled until authentication,
 * transport and an end-to-end test are confirmed.
 */
const REFERRAL_TAB = 'WIC Referrals';
const HEADERS = [
  'Received At','Referral ID','Organization','Staff Contact','Caregiver Name',
  'Phone','Email','County','Preferred Contact','Referral Reasons','Notes',
  'Consent Confirmed','Status','Assigned To','Follow-up Date','Outcome'
];
function setupReferralSheet() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!ss) throw new Error('Bind this Apps Script to the destination Google Sheet.');
  const sheet=ss.getSheetByName(REFERRAL_TAB)||ss.insertSheet(REFERRAL_TAB);
  if(sheet.getLastRow()===0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#00657f').setFontColor('#ffffff');
  }
}
function validateReferral_(p) {
  const fields=['organization','family','phone','county'];
  for(const field of fields) if(!String(p[field]||'').trim()) throw new Error('Missing required field: '+field);
  if(p.consent!==true) throw new Error('Consent is required');
  if(String(p.family).length>150||String(p.phone).length>40||String(p.notes||'').length>1000) throw new Error('Input exceeds limit');
  if(p.email&&!/^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(String(p.email))) throw new Error('Invalid email');
}
function safeCell_(value) {
  const s=String(value==null?'':value).trim();
  return /^[=+@\\-\\t\\r]/.test(s)?"'"+s:s;
}
/** Call only from an authenticated, agency-approved entry point. */
function recordReferral_(payload) {
  validateReferral_(payload);
  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName(REFERRAL_TAB);
    if(!sheet) throw new Error('Run setupReferralSheet first');
    const id=Utilities.getUuid();
    sheet.appendRow([
      new Date(),id,safeCell_(payload.organization),safeCell_(payload.staff),
      safeCell_(payload.family),safeCell_(payload.phone),safeCell_(payload.email),
      safeCell_(payload.county),safeCell_(payload.contact),
      safeCell_(Array.isArray(payload.reasons)?payload.reasons.join('; '):''),
      safeCell_(payload.notes),'Yes','New','','',''
    ]);
    return {ok:true,referralId:id};
  } finally {lock.releaseLock();}
}
/*
 * Intentionally no anonymous doPost() handler. A public GitHub Pages site
 * cannot keep an authentication secret. Configure an authenticated intake
 * before enabling submission of personal contact details.
 */
