/** ================= Go-Pods Owner Portal API (namespaced) ==================
 * Rewards catalog + grouped redemption (delivery vs workshop)
 * - Reads REWARD ITEMS
 * - Caches catalog for 10 minutes
 * - Redeem supports: collectAtFitting, workshop booking (chassis + preferred date)
 * - Auto-CCs Catherine if any fitting items are ordered
 *
 * Tabs:
 *   Owners, Viewings, Redeemed Points, REWARD ITEMS, (optional) Config
 *
 * Endpoints (POST JSON):
 *   action: 'ping'
 *   action: 'login'                 { ownerNumber, password }
 *   action: 'me'
 *   action: 'viewings'
 *   action: 'confirmviewing'        { viewingId }
 *   action: 'updateviewingdate'     { viewingId, dateISO }
 *   action: 'ownerfeedback'         { viewingId, feedback }
 *   action: 'rewards'               {}
 *   action: 'redeem'                {
 *       items:[{sku,qty}],
 *       shipping:{line1,line2,town,postcode,phone} | null,
 *       collectAtFitting:boolean,
 *       workshop:{ chassisNumber, preferredDateISO } | null
 *   }
 *   action: 'setactive'             { active:true|false }
 *   action: 'changepassword'        { oldPassword, newPassword }
 *   action: 'logout'
 *   action: 'createviewingrequest'  { ownerDisplay? | ownerNumber?, ... }
 *   action: 'ownersformap'
 */

const GP_CONFIG = {
  APP_NAME: 'Go-Pods Owner Portal',
  MAIL_TO: ['sales@go-pods.co.uk','rachel@go-pods.co.uk','shop@redlioncaravancentre.co.uk'],
  MAIL_FIT_CONTACT: 'catherine@rlcaravans.com',

  TOKEN_TTL_SECS: 60 * 60 * 12,

  // Send-as alias (must be verified on the sales@ account)
  MAIL_FROM_ALIAS: 'owner-viewings@go-pods.co.uk',
  MAIL_SENDER_NAME: 'Go-Pods Owner Viewings',
  MAIL_REPLY_TO: 'sales@go-pods.co.uk',

  // Test gate for owner notifications used elsewhere
  NOTIFY_TEST_ONLY: true,
  TEST_OWNER_NUMBER: '321'
};

/* ---------------- Util ---------------- */
function GP_tz(){ return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Europe/London'; }
function GP_fmtUK(d){
  if (!d) return '';
  const dd = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dd)) return '';
  return Utilities.formatDate(dd, GP_tz(), 'dd/MM/yyyy');
}

/* -------------- Config + Sheet helpers -------------- */
function GP_cfg(key, fallback){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Config');
  if (!sh) return fallback;
  const vals = sh.getDataRange().getValues();
  for (let i=1;i<vals.length;i++){
    if (String(vals[i][0]) === key) return String(vals[i][1]);
  }
  return fallback;
}
function GP_getTab(nameKey, defaultName){
  const name = GP_cfg(nameKey, defaultName);
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error(`Tab not found: ${name} (key ${nameKey})`);
  return sh;
}
function GP_readTable(tabKey, defaultName){
  const sh = GP_getTab(tabKey, defaultName);
  const vals = sh.getDataRange().getValues();
  const header = vals.shift().map(String);
  return { header, rows: vals, sheet: sh };
}
function GP_colIndex(header, key, defaultCol){
  const label = GP_cfg(key, defaultCol);
  const i = header.indexOf(label);
  if (i < 0) throw new Error(`Column not found: ${label} (key ${key})`);
  return i;
}
function GP_json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj || {}))
    .setMimeType(ContentService.MimeType.JSON);
}
function GP_jsonError(msg, code){
  return GP_json({ ok:false, error: msg || 'Error', code: code || 400 });
}

/* ---------------- Mail helper (alias + HTML) ---------------- */
function GP_send(to, subject, body, extra){
  const opts = {
    name: GP_CONFIG.MAIL_SENDER_NAME || GP_CONFIG.APP_NAME,
    replyTo: GP_CONFIG.MAIL_REPLY_TO || undefined,
    from: GP_CONFIG.MAIL_FROM_ALIAS || undefined
  };
  if (extra && extra.cc) opts.cc = extra.cc;
  if (extra && extra.htmlBody) opts.htmlBody = extra.htmlBody;
  try {
    GmailApp.sendEmail(to, subject, body, opts);
  } catch (e) {
    try {
      const {from, ...rest} = opts;
      GmailApp.sendEmail(to, subject, body, rest);
    } catch (e2) {
      MailApp.sendEmail({ to, subject, body, name: opts.name, replyTo: opts.replyTo, cc: extra && extra.cc ? extra.cc : undefined });
    }
  }
}

/* ---------------- Auth/session ---------------- */
function GP_getAuthTokenFromRequest(e){
  const h = e?.headers || {};
  const auth = h['Authorization'] || h['authorization'] || '';
  const m = auth.match(/Bearer\s+(.+)/i);
  if (m) return m[1].trim();
  try { const body = JSON.parse(e.postData.contents || '{}'); if (body.token) return body.token; } catch(_){}
  return null;
}
function GP_makeToken(ownerNumber){
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(`gp_sess_${token}`, String(ownerNumber), GP_CONFIG.TOKEN_TTL_SECS);
  return token;
}
function GP_validateToken(token){
  if (!token) return null;
  return CacheService.getScriptCache().get(`gp_sess_${token}`);
}
function GP_requireAuth(token, fn){
  const ownerNumber = GP_validateToken(token);
  if (!ownerNumber) throw new Error('Unauthorized');
  return fn(ownerNumber);
}

/* ---------------- Owners lookup ---------------- */
function GP_findOwner(ownerNumber){
  const { header, rows, sheet } = GP_readTable('OWNERS_TAB','Owners');
  const iNum   = GP_colIndex(header,'Owners.OwnerNumber','#');
  const iDisp  = GP_colIndex(header,'Owners.OwnerDisplay','Owner');
  const iFN    = GP_colIndex(header,'Owners.FirstName','First Name');
  const iSN    = GP_colIndex(header,'Owners.Surname','Surname');
  const iJoin  = GP_colIndex(header,'Owners.Joined','Joined');
  const iAct   = GP_colIndex(header,'Owners.ActiveFlag','ACTIVE?');
  const iTotal = GP_colIndex(header,'Owners.Total','TOTAL');
  const iViews = GP_colIndex(header,'Owners.ViewingsCount','# of viewings');
  const iHash  = GP_colIndex(header,'Owners.PasswordHash','PasswordHash');
  const iSalt  = GP_colIndex(header,'Owners.Salt','Salt');

  const want = String(ownerNumber).replace(/\D/g,'');
  for (let r=0; r<rows.length; r++){
    const row = rows[r];
    const num = String(row[iNum]).replace(/\D/g,'');
    if (num === want){
      return { rowIndex: r+2, sheet, header, row,
        idx: { iNum,iDisp,iFN,iSN,iJoin,iAct,iTotal,iViews,iHash,iSalt } };
    }
  }
  return null;
}
function GP_hmac(plain, salt){
  const bytes = Utilities.computeHmacSha256Signature(plain, salt);
  return Utilities.base64Encode(bytes);
}
function GP_setOwnerPassword(ownerNumber, plain){
  const rec = GP_findOwner(ownerNumber);
  if (!rec) throw new Error('Owner not found');
  const salt = Utilities.getUuid().slice(0,8);
  const hash = GP_hmac(plain, salt);
  rec.sheet.getRange(rec.rowIndex, rec.idx.iHash+1).setValue(hash);
  rec.sheet.getRange(rec.rowIndex, rec.idx.iSalt+1).setValue(salt);
}

/* ---------------- Rewards catalog ---------------- */
function GP_rewardsRead(){
  const cache = CacheService.getScriptCache();
  const hit = cache.get('gp_rewards_v1');
  if (hit) return JSON.parse(hit);

  const { header, rows } = GP_readTable('REWARDS_TAB','REWARD ITEMS');
  const idx = {
    sku: header.indexOf('SKU'),
    name: header.indexOf('ITEM NAME'),
    desc: header.indexOf('ITEM DESCRIPTION'),
    pts: header.indexOf('POINTS COST'),
    img: header.indexOf('IMAGE URL'),
    active: header.indexOf('ACTIVE?'),
    fit: header.indexOf('REQUIRES FITTING?'),
    max: header.indexOf('MAX QTY / ORDER'),
    cat: header.indexOf('CATEGORY'),
    sort: header.indexOf('SORT ORDER')
  };
  ['sku','name','pts','active','fit'].forEach(k => { if (idx[k] < 0) throw new Error(`REWARD ITEMS missing column: ${k}`); });

  const items = rows
    .filter(r => String(r[idx.active] || 'Y').toUpperCase().startsWith('Y'))
    .map(r => ({
      sku: String(r[idx.sku]).trim(),
      name: String(r[idx.name]).trim(),
      description: String(r[idx.desc] || '').trim(),
      points: Number(r[idx.pts] || 0),
      imageUrl: String(r[idx.img] || '').trim(),
      requiresFitting: String(r[idx.fit] || 'N').toUpperCase().startsWith('Y'),
      maxPerOrder: idx.max >= 0 ? Number(r[idx.max] || 0) : 0,
      category: idx.cat >= 0 ? String(r[idx.cat] || '').trim() : '',
      sort: idx.sort >= 0 ? Number(r[idx.sort] || 0) : 0
    }))
    .sort((a,b)=> (a.sort||0) - (b.sort||0) || a.name.localeCompare(b.name));

  cache.put('gp_rewards_v1', JSON.stringify(items), 600);
  return items;
}
function GP_getRewardMap(){
  const arr = GP_rewardsRead();
  const map = {};
  arr.forEach(it => { map[it.sku] = it; });
  return map;
}

/* ---------------- API: rewards ---------------- */
function GP_apiRewards(){
  const items = GP_rewardsRead();
  return { ok:true, items };
}

/* ---------------- Redeem (grouping + workshop) ---------------- */
function GP_redeem(ownerNumber, items, shipping, collectAtFitting, workshop){
  if (!Array.isArray(items) || !items.length) return { ok:false, error:'No items' };

  // Validate items against catalog
  const rewardMap = GP_getRewardMap();
  const normalized = [];
  for (const it of items){
    const sku = String(it.sku || '').trim();
    const qty = Math.max(1, Number(it.qty || 0));
    const spec = rewardMap[sku];
    if (!sku || !spec) return { ok:false, error:`Unknown item: ${sku}` };
    if (spec.maxPerOrder && qty > spec.maxPerOrder) return { ok:false, error:`${spec.name}: max ${spec.maxPerOrder} per order` };
    normalized.push({ ...spec, qty });
  }

  // Split by requirement
  const fitItems = normalized.filter(i => i.requiresFitting);
  const delivItems = normalized.filter(i => !i.requiresFitting);

  // Validate addresses/booking
  const needsShipping = delivItems.length > 0 && !collectAtFitting;
  if (needsShipping){
    if (!shipping || !shipping.line1 || !shipping.town || !shipping.postcode) {
      return { ok:false, error:'Incomplete shipping address' };
    }
  }
  const needsWorkshop = fitItems.length > 0;
  if (needsWorkshop){
    if (!workshop || !workshop.chassisNumber || !workshop.preferredDateISO) {
      return { ok:false, error:'Workshop booking requires chassis number and preferred date' };
    }
  }

  // Points & balance
  const o = GP_findOwner(ownerNumber);
  if (!o) throw new Error('Owner not found');
  const totalBefore = Number(o.row[o.idx.iTotal] || 0);
  const pointsTotal = normalized.reduce((s,it)=> s + (Number(it.points)||0) * (Number(it.qty)||1), 0);
  if (pointsTotal > totalBefore) return { ok:false, error:`Not enough points. You have ${totalBefore}, need ${pointsTotal}.` };

  // Append to Redeemed Points (A Owner, B Items/desc, C Date, D Points)
  const sh = GP_getTab('REDEEMED_TAB','Redeemed Points');
  const when = new Date();
  const ownerDisplay = String(o.row[o.idx.iDisp]);
  const listToStr = (arr)=> arr.map(it => `${it.sku} ${it.name} x${it.qty} @ ${Number(it.points).toFixed(2)}pts`).join(' | ');
  const descParts = [];
  if (delivItems.length){
    if (collectAtFitting) descParts.push(`DELIVERY ITEMS (to collect at fitting): ${listToStr(delivItems)}`);
    else descParts.push(`DELIVERY ITEMS: ${listToStr(delivItems)}`);
  }
  if (fitItems.length){
    descParts.push(`WORKSHOP FIT ITEMS: ${listToStr(fitItems)}`);
  }
  if (needsShipping){
    const shipMini = [shipping.line1, shipping.town, shipping.postcode].filter(Boolean).join(', ');
    descParts.push(`SHIP TO (summary): ${shipMini}`);
  }
  if (needsWorkshop){
    descParts.push(`WORKSHOP: Chassis ${workshop.chassisNumber}, Pref ${GP_fmtUK(workshop.preferredDateISO)}`);
  }
  const desc = descParts.join(' || ');
  sh.appendRow([ownerDisplay, desc, when, pointsTotal]);
  const newRow = sh.getLastRow();
  sh.getRange(newRow, 3).setNumberFormat('dd/MM/yyyy');

  // Re-read total AFTER
  const o2 = GP_findOwner(ownerNumber);
  const totalAfter = Number(o2.row[o2.idx.iTotal] || (totalBefore - pointsTotal));

  // Emails
  const deliveryLines = delivItems.map(it => ` - ${it.sku} ${it.name} x${it.qty} @ ${Number(it.points).toFixed(2)}pts = ${(it.qty*Number(it.points)).toFixed(2)}`);
  const fittingLines  = fitItems.map(it => ` - ${it.sku} ${it.name} x${it.qty} @ ${Number(it.points).toFixed(2)}pts = ${(it.qty*Number(it.points)).toFixed(2)}`);
  const sections = [];
  if (delivItems.length){
    sections.push(`Delivery items:\n${deliveryLines.join('\n')}`);
    if (collectAtFitting) sections.push(`Note: Delivery items to be collected at fitting (no courier dispatch).`);
  }
  if (fitItems.length){
    sections.push(`Workshop fitting items:\n${fittingLines.join('\n')}`);
  }
  let shipBlock = '';
  if (needsShipping){
    shipBlock = [
      'Ship To:',
      shipping.line1,
      (shipping.line2 || ''),
      shipping.town,
      shipping.postcode,
      (shipping.phone ? `Phone: ${shipping.phone}` : '')
    ].filter(Boolean).join('\n');
  }
  let workshopBlock = '';
  if (needsWorkshop){
    workshopBlock = [
      'Workshop booking request:',
      'Red Lion Caravan Centre',
      '300 Southport Road, Scarisbrick, Lancashire, PR8 5LF',
      `Chassis: ${workshop.chassisNumber}`,
      `Preferred date: ${GP_fmtUK(workshop.preferredDateISO)}`,
      '',
      'Note: Catherine Warden (Aftersales) will contact the owner to confirm the fitting date, as their preferred date may not be available.'
    ].join('\n');
  }

  const baseRecipients = GP_CONFIG.MAIL_TO.slice();
  if (needsWorkshop) baseRecipients.push(GP_CONFIG.MAIL_FIT_CONTACT);
  const toAll = baseRecipients.join(', ');
  const subject = `[${GP_CONFIG.APP_NAME}] Redemption · ${ownerDisplay}`;
  const body = [
    `Owner: ${ownerDisplay}`,
    '',
    sections.join('\n\n'),
    '',
    `Points: ${pointsTotal.toFixed(2)}`,
    `Before: ${totalBefore.toFixed(2)}`,
    `After: ${totalAfter.toFixed(2)}`,
    '',
    shipBlock ? shipBlock + '\n' : '',
    workshopBlock ? workshopBlock + '\n' : ''
  ].join('\n');

  try { GP_send(toAll, subject, body); } catch(_) {}

  return {
    ok:true,
    pointsBefore: totalBefore,
    pointsAfter: totalAfter,
    grouped: {
      delivery: delivItems.map(({sku,name,points,qty}) => ({sku,name,points,qty})),
      fitting:  fitItems.map(({sku,name,points,qty}) => ({sku,name,points,qty})),
      collectAtFitting: !!collectAtFitting
    }
  };
}

/* ---------------- Existing endpoints kept ---------------- */
function GP_login(ownerNumber, password){
  const rec = GP_findOwner(ownerNumber);
  if (!rec) return { ok:false, error:'Invalid credentials' };

  const supplied = String(password || '').trim();
  let salt = rec.row[rec.idx.iSalt];
  let stored = rec.row[rec.idx.iHash];

  if (!salt || !stored){
    const expected = (String(ownerNumber).padStart(3,'0') + '-START').toUpperCase();
    if (supplied.toUpperCase() !== expected){
      return { ok:false, error:'First-time sign-in not completed. Use ###-START (e.g. 321-START) and then change it.' };
    }
    GP_setOwnerPassword(ownerNumber, expected);
    return { ok:true, token: GP_makeToken(String(ownerNumber).padStart(3,'0')) };
  }

  const candidate = GP_hmac(supplied, salt);
  if (candidate !== stored) return { ok:false, error:'Invalid credentials' };

  return { ok:true, token: GP_makeToken(String(ownerNumber).padStart(3,'0')) };
}

function GP_getMe(ownerNumber){
  const o = GP_findOwner(ownerNumber);
  if (!o) throw new Error('Not found');

  const first = String(o.row[o.idx.iFN] || '').trim();
  const sur   = String(o.row[o.idx.iSN] || '').trim();
  const name  = (first || sur) ? `${first} ${sur}`.trim() : String(o.row[o.idx.iDisp] || '');
  const joinedYear   = o.row[o.idx.iJoin];
  const totalPoints  = Number(o.row[o.idx.iTotal] || 0);
  const viewCount    = Number(o.row[o.idx.iViews] || 0);
  const isActive     = String(o.row[o.idx.iAct] || 'Y').toUpperCase().startsWith('Y');

  return {
    ok: true,
    profile: {
      ownerNumber: String(ownerNumber).padStart(3,'0'),
      name, joinedYear, isActive,
      totals: { viewings: viewCount, successful: null, rewardPoints: totalPoints }
    },
    rewards: []
  };
}

function GP_getViewings(ownerNumber){
  const o = GP_findOwner(ownerNumber);
  if (!o) throw new Error('Owner not found');

  const ownerDisplay = String(o.row[o.idx.iDisp]);
  const { header, rows } = GP_readTable('VIEWINGS_TAB','Viewings');
  const iOwner = GP_colIndex(header,'Viewings.OwnerDisplay','OWNER');
  const iStatus= GP_colIndex(header,'Viewings.Status','Status');
  const iName  = GP_colIndex(header,'Viewings.ViewerName','Viewer Name');
  const iEmail = GP_colIndex(header,'Viewings.ViewerEmail','Viewer Email');
  const iVDate = GP_colIndex(header,'Viewings.ViewingDate','VIEWING DATE');
  const iReq   = GP_colIndex(header,'Viewings.RequestedDate','REQUESTED DATE');
  const iCred  = GP_colIndex(header,'Viewings.Credits','Credits');
  const iNotes = GP_colIndex(header,'Viewings.Notes','NOTES');

  const items = [];
  for (let r=0; r<rows.length; r++){
    const row = rows[r];
    if (String(row[iOwner]) !== ownerDisplay) continue;
    items.push({
      viewingId: `ROW-${r+2}`,
      viewerName: row[iName],
      viewerEmail: row[iEmail],
      viewingDate: row[iVDate] || row[iReq] || '',
      status: row[iStatus],
      pointsAllocated: Number(row[iCred] || 0),
      notes: row[iNotes] || ''
    });
  }
  items.sort((a,b)=> String(b.viewingDate).localeCompare(String(a.viewingDate)));
  return { ok:true, items };
}

function GP_statusToCredits(status){
  const s = String(status||'').toUpperCase();
  if (s === 'SALE') return 1;
  if (s === 'VIEWED' || s === 'NO SALE') return 0.25;
  return 0;
}

function GP_confirmViewing(ownerNumber, viewingRowToken){
  const o = GP_findOwner(ownerNumber);
  if (!o) throw new Error('Owner not found');

  const { header, sheet } = GP_readTable('VIEWINGS_TAB','Viewings');
  const iOwner = GP_colIndex(header,'Viewings.OwnerDisplay','OWNER');
  const iStatus= GP_colIndex(header,'Viewings.Status','Status');
  const iName  = GP_colIndex(header,'Viewings.ViewerName','Viewer Name');
  const iVDate = GP_colIndex(header,'Viewings.ViewingDate','VIEWING DATE');
  const iReq   = GP_colIndex(header,'Viewings.RequestedDate','REQUESTED DATE');
  const iCred  = GP_colIndex(header,'Viewings.Credits','Credits');

  const m = String(viewingRowToken||'').match(/^ROW-(\d+)$/);
  if (!m) return { ok:false, error:'Bad viewing id' };
  const rowIndex = Number(m[1]);

  const row = sheet.getRange(rowIndex,1,1,header.length).getValues()[0];
  const ownerDisplay = String(o.row[o.idx.iDisp]);
  if (String(row[iOwner]) !== ownerDisplay) return { ok:false, error:'Not your viewing' };

  const curr = String(row[iStatus]).toUpperCase();
  if (curr === 'ARRANGED') {
    sheet.getRange(rowIndex, iStatus+1).setValue('VIEWED');
    const currentCredits = Number(row[iCred] || 0);
    let pts = currentCredits;
    if (!currentCredits) {
      pts = GP_statusToCredits('VIEWED');
      sheet.getRange(rowIndex, iCred+1).setValue(pts);
    }
    const after = sheet.getRange(rowIndex,1,1,header.length).getValues()[0];
    const vName = after[iName];
    const vDate = after[iVDate] || after[iReq] || new Date();
    GP_notifyOwnerStatusChange(ownerDisplay, vName, vDate, 'VIEWED', pts);
    return { ok:true, updated:{ viewingId:viewingRowToken, status:'VIEWED' } };
  }
  if (curr === 'TBC') return { ok:false, error:'This viewing is TBC and needs admin review first.' };
  return { ok:false, error:`Viewing is already ${curr}` };
}

function GP_updateViewingDate(ownerNumber, viewingRowToken, dateISO){
  const o = GP_findOwner(ownerNumber);
  if (!o) throw new Error('Owner not found');

  const { header, sheet } = GP_readTable('VIEWINGS_TAB','Viewings');
  const iOwner = GP_colIndex(header,'Viewings.OwnerDisplay','OWNER');
  const iVDate = GP_colIndex(header,'Viewings.ViewingDate','VIEWING DATE');

  const m = String(viewingRowToken||'').match(/^ROW-(\d+)$/);
  if (!m) return { ok:false, error:'Bad viewing id' };
  const rowIndex = Number(m[1]);

  const row = sheet.getRange(rowIndex,1,1,header.length).getValues()[0];
  const ownerDisplay = String(o.row[o.idx.iDisp]);
  if (String(row[iOwner]) !== ownerDisplay) return { ok:false, error:'Not your viewing' };

  const d = new Date(dateISO);
  if (isNaN(d)) return { ok:false, error:'Invalid date' };
  const cell = sheet.getRange(rowIndex, iVDate+1);
  cell.setValue(d);
  cell.setNumberFormat('dd/MM/yyyy');
  return { ok:true };
}

function GP_ownerFeedback(ownerNumber, viewingRowToken, feedback){
  if (!feedback) return { ok:false, error:'No feedback' };
  const o = GP_findOwner(ownerNumber);
  if (!o) throw new Error('Owner not found');

  const { header, sheet } = GP_readTable('VIEWINGS_TAB','Viewings');
  const iOwner = GP_colIndex(header,'Viewings.OwnerDisplay','OWNER');
  const iName  = GP_colIndex(header,'Viewings.ViewerName','Viewer Name');
  const iVDate = GP_colIndex(header,'Viewings.ViewingDate','VIEWING DATE');
  const idxFb  = header.indexOf('OWNER FEEDBACK');

  const m = String(viewingRowToken||'').match(/^ROW-(\d+)$/);
  if (!m) return { ok:false, error:'Bad viewing id' };
  const rowIndex = Number(m[1]);

  const row = sheet.getRange(rowIndex,1,1,header.length).getValues()[0];
  const ownerDisplay = String(o.row[o.idx.iDisp]);
  if (String(row[iOwner]) !== ownerDisplay) return { ok:false, error:'Not your viewing' };

  if (idxFb >= 0) sheet.getRange(rowIndex, idxFb+1).setValue(String(feedback));

  try {
    const subject = `Owner viewing feedback - ${row[iName] || '(viewer)'} with ${ownerDisplay}`;
    const vtxt = GP_fmtUK(row[iVDate]);
    const body = `Owner: ${ownerDisplay}\nViewer: ${row[iName] || ''}\nViewing date: ${vtxt || '(no date)'}\n\nFeedback:\n${feedback}`;
    GP_send('sales@go-pods.co.uk, rachel@go-pods.co.uk', subject, body);
  } catch(_) {}
  return { ok:true };
}

function GP_setActive(ownerNumber, active){
  const rec = GP_findOwner(ownerNumber);
  if (!rec) throw new Error('Owner not found');
  const val = active ? 'Y' : 'N';
  rec.sheet.getRange(rec.rowIndex, rec.idx.iAct+1).setValue(val);
  return { ok:true, active: !!active };
}

/* -------------- Prospect ingest -------------- */
function GP_createViewingRequest(payload){
  const { header, sheet } = GP_readTable('VIEWINGS_TAB','Viewings');
  const iOwner = GP_colIndex(header,'Viewings.OwnerDisplay','OWNER');
  const iStatus= GP_colIndex(header,'Viewings.Status','Status');
  const iName  = GP_colIndex(header,'Viewings.ViewerName','Viewer Name');
  const iEmail = GP_colIndex(header,'Viewings.ViewerEmail','Viewer Email');
  const iVDate = GP_colIndex(header,'Viewings.ViewingDate','VIEWING DATE');
  const iReq   = GP_colIndex(header,'Viewings.RequestedDate','REQUESTED DATE');
  const iNotes = GP_colIndex(header,'Viewings.Notes','NOTES');

  let ownerDisplay = (payload.ownerDisplay || '').trim();
  if (!ownerDisplay && payload.ownerNumber){
    const o = GP_findOwner(payload.ownerNumber);
    if (!o) return { ok:false, error:'Owner not found' };
    ownerDisplay = String(o.row[o.idx.iDisp] || '').trim();
  }
  if (!ownerDisplay) return { ok:false, error:'ownerDisplay or ownerNumber required' };

  const row = new Array(header.length).fill('');
  row[iOwner] = ownerDisplay;
  row[iStatus]= 'TBC';
  row[iName]  = (payload.viewerName || '').trim();
  row[iEmail] = (payload.viewerEmail || '').trim();

  const source = String(payload.source || '').toLowerCase();
  if (source === 'impromptu') {
    const reqDate = payload.requestedDate ? new Date(payload.requestedDate) : new Date();
    row[iVDate] = reqDate; row[iReq] = '';
    const parts = [];
    parts.push(`Viewing took place at: ${(payload.location || '').trim() || '(unspecified)'}`);
    parts.push(`Notes: ${(payload.notes || '').trim() || '(none provided)'}`);
    if (payload.viewerPhone) parts.push(`Viewer telephone number: ${payload.viewerPhone}`);
    row[iNotes] = parts.join(' | ');
  } else {
    row[iVDate] = '';
    row[iReq]   = payload.requestedDate ? new Date(payload.requestedDate) : new Date();
    const extras = [];
    if (payload.viewerPhone) extras.push(`Phone: ${payload.viewerPhone}`);
    if (payload.viewerPostcode) extras.push(`Postcode: ${payload.viewerPostcode}`);
    if (payload.notes) extras.push(payload.notes);
    row[iNotes] = extras.join(' | ');
  }

  sheet.appendRow(row);

  // Force date formats
  const newRow = sheet.getLastRow();
  if (row[iVDate]) sheet.getRange(newRow, iVDate+1).setNumberFormat('dd/MM/yyyy');
  if (row[iReq])   sheet.getRange(newRow, iReq+1).setNumberFormat('dd/MM/yyyy');

  try {
    if (source === 'impromptu') {
      const vtxt = GP_fmtUK(row[iVDate]);
      const subject = `Impromptu viewing logged by ${ownerDisplay} - ${row[iName] || '(viewer)'} on ${vtxt}`;
      const preface = `${ownerDisplay} has just logged an impromptu viewing record via the portal. ` +
        `Check the sheet and contact the viewer to confirm it is a genuine viewing that has taken place. ` +
        `Change status to award points once done as applicable.\n\n`;
      const body = preface + `Prospect: ${row[iName]} · ${row[iEmail]}\n${row[iNotes]}\nViewing date: ${vtxt}`;
      GP_send('sales@go-pods.co.uk, rachel@go-pods.co.uk', subject, body);
    } else {
      const rtxt = GP_fmtUK(row[iReq]);
      const subject = `[${GP_CONFIG.APP_NAME}] New viewing request · ${ownerDisplay}`;
      const body = `Prospect: ${row[iName]} · ${row[iEmail]}\n${row[iNotes]}\nRequested: ${rtxt}`;
      GP_send(GP_CONFIG.MAIL_TO[0], subject, body);
    }
  } catch(_) {}

  return { ok:true };
}

function GP_ownersForMap(){
  const { header, rows } = GP_readTable('OWNERS_TAB','Owners');
  const iDisp = GP_colIndex(header,'Owners.OwnerDisplay','Owner');
  const iNum  = GP_colIndex(header,'Owners.OwnerNumber','#');
  const iAct  = GP_colIndex(header,'Owners.ActiveFlag','ACTIVE?');

  const idxLat = header.indexOf('Lat'); // optional
  const idxLng = header.indexOf('Lng'); // optional

  const owners = rows
    .filter(r => String(r[iAct] || 'Y').toUpperCase().startsWith('Y'))
    .map(r => ({
      ownerNumber: String(r[iNum]).replace(/\D/g,'').padStart(3,'0'),
      name: String(r[iDisp]),
      lat: idxLat >=0 ? Number(r[idxLat] || '') : null,
      lng: idxLng >=0 ? Number(r[idxLng] || '') : null
    }))
    .filter(o => o.lat && o.lng);
  return { ok:true, owners };
}

/* ---------------- Owner notify (used by GP_confirmViewing) ---------------- */
function GP_ownerDisplayToNumber(ownerDisplay){
  const m = String(ownerDisplay||'').match(/#\s*(\d{1,3})/);
  return m ? String(m[1]).padStart(3,'0') : null;
}
function GP_shouldNotify(ownerNumber){
  if (!ownerNumber) return false;
  if (GP_CONFIG.NOTIFY_TEST_ONLY) {
    return String(ownerNumber).padStart(3,'0') === String(GP_CONFIG.TEST_OWNER_NUMBER).padStart(3,'0');
    }
  return true;
}
function escapeHtml(s){
  return String(s || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
function GP_notifyOwnerStatusChange(ownerDisplay, viewerName, viewingDate, status, pointsAwarded){
  const ownerNo = GP_ownerDisplayToNumber(ownerDisplay);
  if (!GP_shouldNotify(ownerNo)) return;

  const ownerRec = GP_findOwner(ownerNo);
  if (!ownerRec) return;
  const firstName = String(ownerRec.row[ownerRec.idx.iFN] || '').trim() || 'there';
  const totalNow  = Number(ownerRec.row[ownerRec.idx.iTotal] || 0);

  const whenTxt = GP_fmtUK(viewingDate) || '(date tbc)';
  const subj = `Go-Pods viewing update - ${viewerName || 'viewer'} on ${whenTxt}`;

  const portalUrl = 'https://www.go-pods.co.uk/owner-viewings-portal';
  const isSale = String(status).toUpperCase() === 'SALE';
  const ptsTxt = isSale ? '1' : (pointsAwarded != null && pointsAwarded !== '' ? String(pointsAwarded) : '0.25');

  const textBody = isSale ? [
    `Hi ${firstName},`,
    ``,
    `Just a quick email to let you know that we’ve awarded 1 point for your recent viewing with ${viewerName}, as they’ve now placed an order with us for their very own Go-Pod! Your reward points total is now ${totalNow}.`,
    ``,
    `Thanks so much for your help with this - it genuinely is appreciated. With owners like you “spreading the good word”, we’re consistently growing the Go-Pods community across the UK and beyond.`,
    ``,
    `You can log into our new “owner viewings portal” to check your viewing record, points total, provide feedback on viewings and redeem your reward points at this link; ${portalUrl}`,
    ``,
    `If you need anything, please feel free to send us an email at sales@go-pods.co.uk / rachel@go-pods.co.uk, or give us a call on 01234 816 832.`,
    ``,
    `Kind regards,`,
    ``,
    `The Go-Pods team`
  ].join('\n') : [
    `Hi ${firstName},`,
    ``,
    `Just a quick email to let you know that we’ve awarded ${ptsTxt} points for your recent viewing with ${viewerName}. Your reward points total is now ${totalNow}.`,
    ``,
    `You can log into our new “owner viewings portal” to check your viewing record, points total, provide feedback on viewings and redeem your reward points at this link; ${portalUrl}`,
    ``,
    `Thanks so much for accommodating this viewing and we’ll be in touch to let you know if they place an order!`,
    ``,
    `Kind regards,`,
    ``,
    `The Go-Pods team`
  ].join('\n');

  const htmlBody = isSale ? [
    `<p>Hi ${escapeHtml(firstName)},</p>`,
    `<p>Just a quick email to let you know that we’ve awarded <strong>1 point</strong> for your recent viewing with ${escapeHtml(viewerName || 'a viewer')}, as they’ve now placed an order with us for their very own Go-Pod! Your reward points total is now <strong>${escapeHtml(String(totalNow))}</strong>.</p>`,
    `<p>Thanks so much for your help with this - it genuinely is appreciated. With owners like you “spreading the good word”, we’re consistently growing the Go-Pods community across the UK and beyond.</p>`,
    `<p>You can log into our new “owner viewings portal” to check your viewing record, points total, provide feedback on viewings and redeem your reward points at this link: <a href="${portalUrl}" target="_blank" rel="noopener">Owner Viewings Portal</a></p>`,
    `<p>If you need anything, please feel free to email <a href="mailto:sales@go-pods.co.uk">sales@go-pods.co.uk</a> / <a href="mailto:rachel@go-pods.co.uk">rachel@go-pods.co.uk</a>, or give us a call on <a href="tel:+441234816832">01234 816 832</a>.</p>`,
    `<p>Kind regards,<br/>The Go-Pods team</p>`
  ].join('') : [
    `<p>Hi ${escapeHtml(firstName)},</p>`,
    `<p>Just a quick email to let you know that we’ve awarded <strong>${escapeHtml(ptsTxt)} points</strong> for your recent viewing with ${escapeHtml(viewerName || 'a viewer')}. Your reward points total is now <strong>${escapeHtml(String(totalNow))}</strong>.</p>`,
    `<p>You can log into our new “owner viewings portal” to check your viewing record, points total, provide feedback on viewings and redeem your reward points at this link: <a href="${portalUrl}" target="_blank" rel="noopener">Owner Viewings Portal</a></p>`,
    `<p>Thanks so much for accommodating this viewing and we’ll be in touch to let you know if they place an order!</p>`,
    `<p>Kind regards,<br/>The Go-Pods team</p>`
  ].join('');

  const toAddr = `${ownerNo}@go-pod.com`;
  const ccList = 'rachel@go-pods.co.uk, sales@go-pods.co.uk';
  GP_send(toAddr, subj, textBody, { cc: ccList, htmlBody });
}

/* ---------------- Router ---------------- */
function doPost(e){
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = String(body.action || '').toLowerCase();
    const token = GP_getAuthTokenFromRequest(e);

    switch (action) {
      case 'ping':                 return GP_json({ ok:true, ping:'pong', time:new Date().toISOString() });

      case 'login':                return GP_json(GP_login(body.ownerNumber, body.password));
      case 'me':                   return GP_json(GP_requireAuth(token, GP_getMe));
      case 'viewings':             return GP_json(GP_requireAuth(token, GP_getViewings));
      case 'confirmviewing':       return GP_json(GP_requireAuth(token, (n)=>GP_confirmViewing(n, body.viewingId)));
      case 'updateviewingdate':    return GP_json(GP_requireAuth(token, (n)=>GP_updateViewingDate(n, body.viewingId, body.dateISO)));
      case 'ownerfeedback':        return GP_json(GP_requireAuth(token, (n)=>GP_ownerFeedback(n, body.viewingId, body.feedback)));

      case 'rewards':              return GP_json(GP_requireAuth(token, GP_apiRewards));
      case 'redeem':               return GP_json(GP_requireAuth(token, (n)=>GP_redeem(n, body.items, body.shipping, !!body.collectAtFitting, body.workshop || null)));

      case 'setactive':            return GP_json(GP_requireAuth(token, (n)=>GP_setActive(n, !!body.active)));
      case 'changepassword':       return GP_json(GP_requireAuth(token, (n)=>changePassword_GP(n, body.oldPassword, body.newPassword)));
      case 'logout':               return GP_json({ ok:true });

      case 'createviewingrequest': return GP_json(GP_createViewingRequest(body));
      case 'ownersformap':         return GP_json(GP_ownersForMap());
      default: return GP_jsonError('Unknown action');
    }
  } catch (err) {
    return GP_jsonError(err?.message || 'Server error', 500);
  }
}

/* ---------------- Health check ---------------- */
function doGet(e){
  return ContentService
    .createTextOutput(JSON.stringify({ ok:true, service:'Go-Pods Owner Portal', time:new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}
