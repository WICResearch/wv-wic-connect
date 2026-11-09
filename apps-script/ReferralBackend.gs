/** WV WIC Connect referral intake, spreadsheet-bound Apps Script. */
const REFERRAL_TAB='WIC Referrals';
const HEADERS=['Received At','Referral ID','Organization','Staff Contact','Caregiver Name','Phone','Email','County','Preferred Contact','Referral Reasons','Notes','Consent Confirmed','Status','Assigned To','Follow-up Date','Outcome'];
function setupReferralSheet(){
 const ss=SpreadsheetApp.getActiveSpreadsheet();
 if(!ss)throw new Error('Bind script to destination spreadsheet');
 const sh=ss.getSheetByName(REFERRAL_TAB)||ss.insertSheet(REFERRAL_TAB);
 if(sh.getLastRow()===0){sh.appendRow(HEADERS);sh.setFrozenRows(1);sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#00657f').setFontColor('#ffffff');}
}
function validateReferral_(p){
 for(const f of ['organization','family','phone','county'])if(!String(p[f]||'').trim())throw new Error('Missing '+f);
 if(p.consent!==true&&p.consent!=='true')throw new Error('Consent required');
 if(String(p.family).length>150||String(p.phone).length>40||String(p.notes||'').length>1000)throw new Error('Field too long');
 if(p.email&&!/^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(String(p.email)))throw new Error('Invalid email');
}
function safeCell_(v){
 const s=String(v==null?'':v).trim();
 return /^[=+@\\-\\t\\r]/.test(s)?"'"+s:s;
}
function recordReferral_(p){
 validateReferral_(p);
 const lock=LockService.getScriptLock();lock.waitLock(10000);
 try{
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REFERRAL_TAB);
  if(!sh)throw new Error('Run setupReferralSheet');
  const id=Utilities.getUuid();
  sh.appendRow([new Date(),id,safeCell_(p.organization),safeCell_(p.staff),safeCell_(p.family),safeCell_(p.phone),safeCell_(p.email),safeCell_(p.county),safeCell_(p.contact),safeCell_(Array.isArray(p.reasons)?p.reasons.join('; '):''),safeCell_(p.notes),'Yes','New','','','']);
  return {ok:true,referralId:id};
 }finally{lock.releaseLock();}
}
/**
 * No-sign-in submission endpoint.
 * Only accepts POST; the browser receives a confirmation page.
 * Public endpoints cannot reliably identify a sender; review spam/abuse.
 */
function doPost(e) {
  try {
    if(!e || !e.parameter) throw new Error('Invalid request');
    const p=e.parameter;
    // Honeypot: ordinary users never fill this hidden field.
    if(p.website) return responsePage_(false,'Submission could not be accepted.');
    const payload={
      organization:p.organization,staff:p.staff,family:p.family,
      phone:p.phone,email:p.email,county:p.county,contact:p.contact,
      reasons:String(p.reasons||'').split('|').filter(Boolean),
      notes:p.notes,consent:p.consent
    };
    if(String(p.startedAt||'') && Date.now()-Number(p.startedAt)<2500)
      throw new Error('Please wait a moment and try again.');
    const cache=CacheService.getScriptCache();
    const fingerprint=Utilities.base64EncodeWebSafe(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
        [payload.organization,payload.family,payload.phone,payload.county].join('|').toLowerCase())
    ).slice(0,40);
    if(cache.get(fingerprint)) throw new Error('This referral was recently submitted.');
    const result=recordReferral_(payload);
    cache.put(fingerprint,'1',300);
    return responsePage_(true,'Referral received. Confirmation: '+result.referralId);
  } catch(err) {
    console.error(err);
    return responsePage_(false,'Referral was not confirmed. Please contact WIC staff or try again.');
  }
}
function responsePage_(success,message) {
  const safe=String(message).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  return HtmlService.createHtmlOutput('<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"></head><body style="font:16px system-ui;padding:28px;max-width:580px;margin:auto"><h1>'+ (success?'Referral received':'Submission not confirmed') +'</h1><p>'+safe+'</p><p>You can close this tab and return to the referral portal.</p></body></html>');
}
function doGet() {
  return HtmlService.createHtmlOutput('WV WIC referral intake endpoint. Submit referrals using the official partner portal.');
}
