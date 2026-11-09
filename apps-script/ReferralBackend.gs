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
  const escapeHtml = value => String(value == null ? '' : value)
    .replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  const match = success ? String(message).match(/Confirmation:\\s*(.+)$/) : null;
  const reference = match ? escapeHtml(match[1]) : '';
  const heading = success ? 'Referral submitted successfully!' : 'We couldn’t confirm your referral';
  const description = success
    ? 'Thank you for helping connect a West Virginia family with WIC services. Your referral has been received and recorded for WIC staff to review and coordinate follow-up.'
    : 'Your referral has not been confirmed. Please try again, or contact WIC staff for assistance. Do not assume the referral was received.';
  const symbol = success ? '&#10003;' : '!';
  const accent = success ? '#006b4f' : '#a65c00';
  const soft = success ? '#e8f7ef' : '#fff3df';
  const referenceBlock = reference ? '<div class="reference"><div class="reference-label">Your confirmation number</div><div class="reference-id">'+reference+'</div><p class="reference-hint">Please save this number for your records.</p></div>' : '';
  const html = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(heading)} | WV WIC Connect</title>
<style>
*{box-sizing:border-box}body{margin:0;background:linear-gradient(150deg,#edf7f4,#f7f9fa 58%,#fceef5);color:#17323d;font:16px/1.65 system-ui,-apple-system,"Segoe UI",sans-serif;min-height:100vh}
.topbar{background:#fff;border-top:5px solid #b51d6b;border-bottom:1px solid #e1e9ea;padding:17px 24px}.top-inner{max-width:800px;margin:auto;display:flex;align-items:center;gap:12px}.brand-mark{display:grid;place-items:center;background:#b51d6b;color:white;border-radius:12px;width:43px;height:43px;font-size:26px;font-weight:bold}.brand-name{font-size:20px;font-weight:800;color:#005b77;letter-spacing:-.04em}.brand-sub{font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:#60747c;font-weight:700}
main{max-width:690px;margin:72px auto;padding:0 20px}.card{background:white;border:1px solid #dfebeb;border-radius:22px;box-shadow:0 20px 55px #004d6114;text-align:center;padding:clamp(28px,6vw,58px)}.status{display:grid;place-items:center;width:76px;height:76px;border-radius:50%;background:${soft};color:${accent};font-size:42px;font-weight:800;margin:0 auto 24px}.eyebrow{font-size:12px;text-transform:uppercase;letter-spacing:.14em;color:#00657f;font-weight:800}h1{font-size:clamp(28px,5vw,39px);line-height:1.18;letter-spacing:-.045em;margin:12px 0 18px}p{color:#526b74;margin:0 auto;max-width:520px}.reference{background:#f1f7f7;border:1px solid #dce9e9;border-radius:14px;padding:19px 14px;margin:30px 0}.reference-label{font-size:12px;text-transform:uppercase;letter-spacing:.1em;color:#526b74;font-weight:800}.reference-id{font-size:clamp(13px,2.5vw,17px);font-weight:800;color:#004d61;overflow-wrap:anywhere;margin:9px 0}.reference-hint{font-size:13px}.button{display:inline-block;background:#00657f;color:white;text-decoration:none;padding:13px 24px;border-radius:11px;font-weight:750;margin-top:28px}.button:hover{background:#004d61}.button:focus-visible{outline:3px solid #b51d6b;outline-offset:3px}.footnote{text-align:center;font-size:13px;color:#60747c;margin:22px auto}@media(max-width:600px){main{margin:35px auto}.card{padding:30px 20px}}
</style></head><body>
<header class="topbar"><div class="top-inner"><span class="brand-mark" aria-hidden="true">&#9829;</span><div><div class="brand-sub">Community resource &amp; referral hub</div><div class="brand-name">FRN &amp; WIC Connect</div></div></div></header>
<main><section class="card" aria-labelledby="status-heading"><div class="status" aria-hidden="true">${symbol}</div><div class="eyebrow">West Virginia WIC</div><h1 id="status-heading">${escapeHtml(heading)}</h1><p>${escapeHtml(description)}</p>${referenceBlock}<a class="button" href="https://wicresearch.github.io/wv-wic-connect/">Return to WV WIC Connect</a></section><p class="footnote">Supporting West Virginia families through trusted community connections.</p></main></body></html>`;
  return HtmlService.createHtmlOutput(html).setTitle(heading);
}
function doGet() {
  return HtmlService.createHtmlOutput('WV WIC referral intake endpoint. Submit referrals using the official partner portal.');
}
