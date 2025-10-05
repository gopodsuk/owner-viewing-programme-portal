/** Auto-credits + points snapshot + IMMEDIATE owner notifications on Status edit (installable onEdit) */

const CF_CONFIG = {
  NOTIFY_TEST_ONLY: true,
  TEST_OWNER_NUMBER: '321',
  MAIL_FROM_ALIAS: 'owner-viewings@go-pods.co.uk',
  MAIL_SENDER_NAME: 'Go-Pods Owner Viewings',
  MAIL_REPLY_TO: 'sales@go-pods.co.uk',
  CC_LIST: 'rachel@go-pods.co.uk, sales@go-pods.co.uk',
  PORTAL_URL: 'https://www.go-pods.co.uk/owner-viewings-portal',
  COL_OWNER: 'OWNER',
  COL_VIEWER_NAME: 'Viewer Name',
  COL_STATUS: 'Status',
  COL_VIEWING_DATE: 'VIEWING DATE',
  COL_REQUESTED_DATE: 'REQUESTED DATE',
  COL_CREDITS: 'Credits',
  COL_POINTS_SNAPSHOT: 'Points total'
};

const CF_POINTS = { 'SALE': 1, 'VIEWED': 0.25, 'NO SALE': 0.25 };

function onEdit(e){
  // only the INSTALLABLE on-edit should run this (simple trigger has no triggerUid)
  if (!e || !e.triggerUid) return;

  try {
    const sh = e.range.getSheet();
    if (!sh || sh.getName() !== 'Viewings') return;

    const row = e.range.getRow();
    if (row <= 1) return;

    const header = CF_getHeader(sh);
    const iOwner  = header.indexOf(CF_CONFIG.COL_OWNER) + 1;
    const iName   = header.indexOf(CF_CONFIG.COL_VIEWER_NAME) + 1;
    const iStatus = header.indexOf(CF_CONFIG.COL_STATUS) + 1;
    const iVDate  = header.indexOf(CF_CONFIG.COL_VIEWING_DATE) + 1;
    const iReq    = header.indexOf(CF_CONFIG.COL_REQUESTED_DATE) + 1;
    const iCred   = header.indexOf(CF_CONFIG.COL_CREDITS) + 1;
    const iSnap   = header.indexOf(CF_CONFIG.COL_POINTS_SNAPSHOT) + 1;

    if (!(iOwner && iName && iStatus && iVDate && iReq && iCred)) return;

    const col = e.range.getColumn();
    const get = (c) => sh.getRange(row, c).getValue();
    const set = (c, v) => sh.getRange(row, c).setValue(v);

    // STATUS edited by user
    if (col === iStatus) {
      const oldVal = (typeof e.oldValue !== 'undefined') ? String(e.oldValue).toUpperCase().trim() : null;
      const newVal = (typeof e.value     !== 'undefined') ? String(e.value).toUpperCase().trim()     : null;
      if (!newVal || newVal === oldVal) return;

      // Auto-credits (respect manual overrides)
      const targetPts = CF_POINTS[newVal] || 0;
      const currPts = Number(get(iCred) || 0);
      if (targetPts && currPts !== targetPts) set(iCred, targetPts);

      // Refresh snapshot (if present)
      if (iSnap > 0) {
        Utilities.sleep(200);
        const ownerDisplay = String(get(iOwner) || '');
        const ownerNo = CF_ownerDisplayToNumber(ownerDisplay);
        const totals = CF_getOwnerTotals(ownerNo);
        if (totals) set(iSnap, totals.totalPoints);
      }

      // Notify owner immediately (VIEWED/SALE only)
      if (newVal === 'VIEWED' || newVal === 'SALE') {
        const ownerDisplay = String(get(iOwner) || '');
        const ownerNo = CF_ownerDisplayToNumber(ownerDisplay);
        if (CF_shouldNotify(ownerNo)) {
          const viewerName = String(get(iName) || '');
          const when = get(iVDate) || get(iReq) || new Date();
          const pointsAwarded = Number(get(iCred) || (CF_POINTS[newVal] || 0));
          const totals = CF_getOwnerTotals(ownerNo) || { firstName:'there', totalPoints:0 };

          const whenTxt = CF_fmtUK(when);
          const subj = `Go-Pods viewing update - ${viewerName || 'viewer'} on ${whenTxt}`;
          const { textBody, htmlBody } = CF_buildOwnerBodies({
            firstName: totals.firstName,
            viewerName,
            status: newVal,
            pts: pointsAwarded,
            totalNow: totals.totalPoints,
            portalUrl: CF_CONFIG.PORTAL_URL
          });

          CF_sendEmailOwner(`${ownerNo}@go-pod.com`, subj, textBody, htmlBody, CF_CONFIG.CC_LIST);
        }
      }
      return;
    }

    // If Viewing Date edited: format as UK date (no timestamp)
    if (col === iVDate) {
      const v = get(iVDate);
      if (v) {
        const d = (v instanceof Date) ? v : new Date(v);
        if (!isNaN(d)) {
          const cell = sh.getRange(row, iVDate);
          cell.setValue(d);
          cell.setNumberFormat('dd/MM/yyyy');
        }
      }
    }
  } catch (err) {
    console.error('onEdit error:', err && err.message ? err.message : err);
  }
}

function CF_getHeader(sh){ return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String); }
function CF_ownerDisplayToNumber(s){ const m = String(s||'').match(/#\s*(\d{1,3})/); return m ? String(m[1]).padStart(3,'0') : null; }
function CF_shouldNotify(ownerNo){ if (!ownerNo) return false; return CF_CONFIG.NOTIFY_TEST_ONLY ? (String(ownerNo).padStart(3,'0')===String(CF_CONFIG.TEST_OWNER_NUMBER).padStart(3,'0')) : true; }
function CF_getOwnerTotals(ownerNo){
  if (!ownerNo) return null;
  const ss = SpreadsheetApp.getActive(); const sh = ss.getSheetByName('Owners'); if (!sh) return null;
  const vals = sh.getDataRange().getValues(); const header = vals.shift().map(String);
  const idxNum = header.indexOf('#'), idxFN = header.indexOf('First Name'), idxTotal = header.indexOf('TOTAL');
  if (idxNum<0||idxFN<0||idxTotal<0) return null;
  for (let r=0;r<vals.length;r++){
    const row = vals[r]; const num = String(row[idxNum]).replace(/\D/g,'').padStart(3,'0');
    if (num === String(ownerNo).padStart(3,'0')) return { firstName:String(row[idxFN]||'').trim()||'there', totalPoints:Number(row[idxTotal]||0) };
  }
  return null;
}
function CF_buildOwnerBodies({ firstName, viewerName, status, pts, totalNow, portalUrl }){
  const isSale = String(status).toUpperCase()==='SALE';
  const ptsTxt = isSale ? '1' : (pts!=null?pts:0.25);
  const textBody = isSale ? [
    `Hi ${firstName},`,``,`Just a quick email to let you know that we’ve awarded 1 point for your recent viewing with ${viewerName}, as they’ve now placed an order with us for their very own Go-Pod! Your reward points total is now ${totalNow}.`,``,`Thanks so much for your help with this - it genuinely is appreciated. With owners like you “spreading the good word”, we’re consistently growing the Go-Pods community across the UK and beyond.`,``,`You can log into our new “owner viewings portal” here: ${portalUrl}`,``,`If you need anything, please email sales@go-pods.co.uk / rachel@go-pods.co.uk or call 01234 816 832.`,``,`Kind regards,`,``,`The Go-Pods team`
  ].join('\n') : [
    `Hi ${firstName},`,``,`Just a quick email to let you know that we’ve awarded ${ptsTxt} points for your recent viewing with ${viewerName}. Your reward points total is now ${totalNow}.`,``,`You can log into our new “owner viewings portal” here: ${portalUrl}`,``,`Thanks so much for accommodating this viewing and we’ll be in touch to let you know if they place an order!`,``,`Kind regards,`,``,`The Go-Pods team`
  ].join('\n');
  const htmlBody = isSale ? [
    `<p>Hi ${CF_escapeHtml(firstName)},</p>`,
    `<p>Just a quick email to let you know that we’ve awarded <strong>1 point</strong> for your recent viewing with ${CF_escapeHtml(viewerName||'a viewer')}, as they’ve now placed an order with us for their very own Go-Pod! Your reward points total is now <strong>${CF_escapeHtml(String(totalNow))}</strong>.</p>`,
    `<p>Thanks so much for your help with this — it genuinely is appreciated. With owners like you “spreading the good word”, we’re consistently growing the Go-Pods community across the UK and beyond.</p>`,
    `<p>You can log into our new “owner viewings portal” here: <a href="${portalUrl}" target="_blank" rel="noopener">Owner Viewings Portal</a></p>`,
    `<p>If you need anything, please email <a href="mailto:sales@go-pods.co.uk">sales@go-pods.co.uk</a> / <a href="mailto:rachel@go-pods.co.uk">rachel@go-pods.co.uk</a> or call <a href="tel:+441234816832">01234 816 832</a>.</p>`,
    `<p>Kind regards,<br/>The Go-Pods team</p>`
  ].join('') : [
    `<p>Hi ${CF_escapeHtml(firstName)},</p>`,
    `<p>Just a quick email to let you know that we’ve awarded <strong>${CF_escapeHtml(String(ptsTxt))} points</strong> for your recent viewing with ${CF_escapeHtml(viewerName||'a viewer')}. Your reward points total is now <strong>${CF_escapeHtml(String(totalNow))}</strong>.</p>`,
    `<p>You can log into our new “owner viewings portal” here: <a href="${portalUrl}" target="_blank" rel="noopener">Owner Viewings Portal</a></p>`,
    `<p>Thanks so much for accommodating this viewing and we’ll be in touch to let you know if they place an order!</p>`,
    `<p>Kind regards,<br/>The Go-Pods team</p>`
  ].join('');
  return { textBody, htmlBody };
}
function CF_sendEmailOwner(to, subject, textBody, htmlBody, cc){
  const opts = { name: CF_CONFIG.MAIL_SENDER_NAME, from: CF_CONFIG.MAIL_FROM_ALIAS, replyTo: CF_CONFIG.MAIL_REPLY_TO, cc, htmlBody };
  try { GmailApp.sendEmail(to, subject, textBody, opts); }
  catch(e){ try { const {from,...rest}=opts; GmailApp.sendEmail(to, subject, textBody, rest); }
  catch(e2){ MailApp.sendEmail({ to, subject, body: textBody, name: CF_CONFIG.MAIL_SENDER_NAME, replyTo: CF_CONFIG.MAIL_REPLY_TO, cc }); } }
}
function CF_escapeHtml(s){ return String(s || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function CF_fmtUK(d){ const dd=(d instanceof Date)?d:new Date(d); if (isNaN(dd)) return '(no date)'; const tz=SpreadsheetApp.getActive().getSpreadsheetTimeZone()||'Europe/London'; return Utilities.formatDate(dd,tz,'dd/MM/yyyy'); }
