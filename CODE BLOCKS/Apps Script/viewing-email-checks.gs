/** STAFF reminders only (no owner emails) — run daily via time-driven trigger */

const SHEET_NAME = 'Viewings';
const HEADERS_ROW = 1;

const COL = {
  OWNER: 1, VIEWER_NAME: 2, VIEWER_EMAIL: 3, STATUS: 4,
  CONTACT_DATE: 6, DAYS_SINCE_CONTACT: 7, REQUESTED_DATE: 8,
  VIEWING_DATE: 9, FUPPED: 10, CREDITS: 11, FOLLOW_UP_AGAIN: 12, NOTES: 13
};

const TO_SALES_AND_RACHEL = 'sales@go-pods.co.uk, rachel@go-pods.co.uk';
const TO_SALES_ONLY = 'sales@go-pods.co.uk';
const STAFF_SENDER = { from: 'owner-viewings@go-pods.co.uk', name: 'Go-Pods Owner Viewings', replyTo: 'sales@go-pods.co.uk' };
const CUTOFF_MONTHS = 3;

function checkViewings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Europe/London';
  const today = toMidnight(new Date(), tz);
  const cutoffDate = toMidnight(addMonths(today, -CUTOFF_MONTHS), tz);

  const lastRow = sheet.getLastRow(); const lastCol = sheet.getLastColumn();
  if (lastRow <= HEADERS_ROW) return;

  const values = sheet.getRange(HEADERS_ROW + 1, 1, lastRow - HEADERS_ROW, lastCol).getValues();
  const sheetId = sheet.getSheetId().toString();

  values.forEach((row, i) => {
    const rowIndex = HEADERS_ROW + 1 + i;
    const owner = safeStr(row[COL.OWNER - 1]);
    const viewerName = safeStr(row[COL.VIEWER_NAME - 1]);
    const status = safeStr(row[COL.STATUS - 1]).toUpperCase().trim();
    const contactDate   = toDateOrNull(row[COL.CONTACT_DATE - 1]);
    const daysSinceCnt  = toNumberOrNull(row[COL.DAYS_SINCE_CONTACT - 1]);
    const requestedDate = toDateOrNull(row[COL.REQUESTED_DATE - 1]);
    const viewingDate   = toDateOrNull(row[COL.VIEWING_DATE - 1]);
    const fuppedDate    = toDateOrNull(row[COL.FUPPED - 1]);
    const followUpAgain = safeStr(row[COL.FOLLOW_UP_AGAIN - 1]).toUpperCase();

    // Rule 1
    if (status === 'TBC' && daysSinceCnt != null && daysSinceCnt >= 7 && contactDate && toMidnight(contactDate, tz) >= cutoffDate) {
      const key = makeKey(sheetId, rowIndex, 1);
      runOnce(key, () => {
        const subject = `Check up on viewing request - ${viewerName} with ${owner}`;
        const body = [`Follow up with viewer ${viewerName} who requested a viewing with ${owner} on ${fmtDate(requestedDate, tz)}.`,``,`First contacted us on ${fmtDate(contactDate, tz)}, ${daysSinceCnt} days ago.`].join('\n');
        sendTeamEmail(TO_SALES_AND_RACHEL, subject, body);
      });
    }

    // Rule 2
    if (status === 'ARRANGED' && requestedDate && !viewingDate) {
      const daysUntilReq = dateDiffInDays(toMidnight(requestedDate, tz), today);
      if (daysUntilReq >= 0 && daysUntilReq <= 14) {
        const key = makeKey(sheetId, rowIndex, 2);
        runOnce(key, () => {
          const subject = `Check up on arranged viewing - ${viewerName} with ${owner} on ${fmtDate(requestedDate, tz)}`;
          const body = `Double check that ${viewerName} with ${owner} on ${fmtDate(requestedDate, tz)} is definitely confirmed. The “VIEWING DATE” field is currently empty.`;
          sendTeamEmail(TO_SALES_AND_RACHEL, subject, body);
        });
      }
    }

    // Rule 3
    if (status === 'ARRANGED' && requestedDate && viewingDate) {
      const vMid = toMidnight(viewingDate, tz);
      if (vMid < today && vMid >= cutoffDate) {
        const key = makeKey(sheetId, rowIndex, 3);
        runOnce(key, () => {
          const subject = `Check if viewing went ahead - ${viewerName} with ${owner} on ${fmtDate(viewingDate, tz)}`;
          const body = `Suzanne: Follow up with ${owner} to check if their viewing with ${viewerName} went ahead as planned. Update the “Status” field as “VIEWED” once confirmed to allocate viewing points.`;
          sendTeamEmail(TO_SALES_AND_RACHEL, subject, body);
        });
      }
    }

    // Rule 4
    if (status === 'VIEWED' && viewingDate && toMidnight(viewingDate, tz) >= cutoffDate) {
      const key = makeKey(sheetId, rowIndex, 4);
      runOnce(key, () => {
        const subject = `Follow up owner viewing - ${viewerName} with ${owner} on ${fmtDate(viewingDate, tz)}`;
        const body = [`Follow up with ${viewerName}, who viewed with ${owner} on ${fmtDate(viewingDate, tz)}.`,``,`Update the “FUPPED” field once done and set a reminder to follow up again in a week if no reply.`].join('\n');
        sendTeamEmail(TO_SALES_ONLY, subject, body);
      });
    }

    // Rule 5
    if (status === 'VIEWED' && followUpAgain === 'YES') {
      const baseline = fuppedDate ? fuppedDate : viewingDate;
      if (baseline) {
        const bMid = toMidnight(baseline, tz);
        const daysSinceBaseline = dateDiffInDays(today, bMid);
        if (daysSinceBaseline > 7) {
          const key = makeKey(sheetId, rowIndex, 5);
          const lastSentISO = getLastSent(key);
          let shouldSend = false;
          if (!lastSentISO) shouldSend = true;
          else {
            const lastSent = new Date(lastSentISO);
            if (dateDiffInDays(today, lastSent) >= 7) shouldSend = true;
          }
          if (shouldSend) {
            setLastSent(key);
            const subject = `Follow up again with ${viewerName}`;
            const body = `Reminder to follow up again with ${viewerName}, who viewed with ${owner} on ${fmtDate(viewingDate, tz)}.` +
                         (fuppedDate ? ` Already followed up on ${fmtDate(fuppedDate, tz)}.` : ` No previous follow-up date recorded yet.`);
            sendTeamEmail(TO_SALES_ONLY, subject, body);
          }
        }
      }
    }
  });
}

// helpers
function sendTeamEmail(to, subject, body) {
  try { GmailApp.sendEmail(to, subject, body, STAFF_SENDER); }
  catch (e) { MailApp.sendEmail({ to, subject, body, name: STAFF_SENDER.name }); }
}
function runOnce(key, fn){ const p=PropertiesService.getScriptProperties(); if (p.getProperty(key)) return; fn(); p.setProperty(key, new Date().toISOString()); }
function getLastSent(key){ return PropertiesService.getScriptProperties().getProperty(key); }
function setLastSent(key){ PropertiesService.getScriptProperties().setProperty(key, new Date().toISOString()); }
function makeKey(sheetId,rowIndex,rule){ return `viewings_${sheetId}_r${rowIndex}_rule${rule}`; }
function toDateOrNull(v){ if(!v) return null; if (v instanceof Date) return v; const d=new Date(v); return isNaN(d)?null:d; }
function toNumberOrNull(v){ if (v===''||v==null) return null; const n=Number(v); return isNaN(n)?null:n; }
function safeStr(v){ return (v==null)?'':String(v).trim(); }
function toMidnight(date, tz){ const fmt=Utilities.formatDate(date, tz, 'yyyy-MM-dd'); const [y,m,d]=fmt.split('-').map(Number); return new Date(y,m-1,d); }
function dateDiffInDays(a,b){ const MS=86400000; return Math.floor((toMidnight(a,'UTC') - toMidnight(b,'UTC'))/MS); }
function fmtDate(d,tz){ if(!d) return '(no date)'; return Utilities.formatDate(d, tz, 'EEE d MMM yyyy'); }
function addMonths(date, months){ const d=new Date(date); const od=d.getDate(); d.setMonth(d.getMonth()+months); if (d.getDate()<od) d.setDate(0); return d; }
