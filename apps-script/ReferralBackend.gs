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
