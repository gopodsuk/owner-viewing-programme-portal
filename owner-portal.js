(function(){
  // --- Configure your Apps Script Web App URL ---
  const API = 'https://script.google.com/macros/s/AKfycbxB99bU28_gTFi_PXs18HLYcQpDhtGqksFR9UEQRG9Hod2e2YjQoOJYsOx6_CNBgmjS/exec';

  // Token + root
  let TOKEN = localStorage.getItem('gp_token') || '';
  const root = document.getElementById('gp-owner-portal');

  // Polyfill CSS.escape for older browsers
  if (typeof CSS === 'undefined' || !CSS.escape){
    window.CSS = { escape: s => String(s).replace(/[^a-zA-Z0-9_\-]/g, ch => '\\' + ch) };
  }

  // Utilities
  function mount(html){ root.innerHTML = `<div class="gp-wrap">${html}</div>`; }
  const esc = s => String(s ?? '').replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));

  // Date helpers (UK + ISO)
  const toISO = d => { if(!d) return ''; const dt = (d instanceof Date)?d:new Date(d); if(isNaN(dt)) return ''; const y=dt.getFullYear(), m=String(dt.getMonth()+1).padStart(2,'0'), da=String(dt.getDate()).padStart(2,'0'); return `${y}-${m}-${da}`; };
  const toUK  = d => { if(!d) return '-'; const dt=(d instanceof Date)?d:new Date(d); if(isNaN(dt)) return '-'; const da=String(dt.getDate()).padStart(2,'0'), m=String(dt.getMonth()+1).padStart(2,'0'), y=dt.getFullYear(); return `${da}/${m}/${y}`; };

  // Status pills
  function statusClassFor(key){
    switch (String(key || '').toUpperCase()) {
      case 'TBC':       return 'status--tbc';
      case 'ARRANGED':  return 'status--arranged';
      case 'VIEWED':    return 'status--viewed';
      case 'SALE':      return 'status--sale';
      case 'NO SALE':   return 'status--no-sale';
      default:          return '';
    }
  }
  function statusPillHTML(key, id){
    const cls = statusClassFor(key);
    const label = String(key || '').toUpperCase();
    let iconLeft = '', iconRight = '';
    if (label === 'VIEWED') iconLeft = '<i>✓</i>';
    if (label === 'SALE')   { iconLeft = '<i>★</i>'; iconRight = '<i>★</i>'; }
    return `<span id="status-${id}" class="status ${cls}">${iconLeft}<span>${label}</span>${iconRight}</span>`;
  }

  // API helper
  async function call(action, payload={}, useAuth=true){
    const body = { action, ...payload };
    if (useAuth && TOKEN) body.token = TOKEN;
    try {
      const res = await fetch(API, { method:'POST', headers:{'Content-Type':'text/plain;charset=utf-8'}, body: JSON.stringify(body) });
      const txt = await res.text();
      try { return JSON.parse(txt); } catch { return { ok:false, error:`Server returned non-JSON: ${txt.slice(0,160)}…` }; }
    } catch (e) { return { ok:false, error: `Network error: ${e?.message || e}` }; }
  }

  /* ------------------------- Login ------------------------- */
  function loginView(msg=''){
    mount(`
      <div class="gp-hero">
        <h2>Go-Pods Owner Portal</h2>
        <button class="gp-btn secondary" style="opacity:.85;pointer-events:none">Secure</button>
      </div>
      <div class="gp-grid" style="grid-template-columns:1fr;">
        <div class="gp-card">
          <h3>Owner Login</h3>
          ${msg ? `<div class="gp-error">${esc(msg)}</div>` : ''}
          <div class="formgrid-2">
            <div>
              <label>Owner Number</label>
              <input id="ownerNumber" class="gp-input" placeholder="e.g. 123" inputmode="numeric"/>
            </div>
            <div>
              <label>Password</label>
              <input id="password" class="gp-input" type="password"/>
            </div>
          </div>
          <div class="gp-row"><button class="gp-btn" id="loginBtn">Login</button></div>
          <p class="gp-muted">First time? Use <code>###-START</code> (e.g. <code>123-START</code>), then change it.</p>
        </div>
      </div>
    `);
    document.getElementById('loginBtn').onclick = async ()=>{
      const ownerNumber = document.getElementById('ownerNumber').value.trim();
      const password    = document.getElementById('password').value.trim();
      const r = await call('login', { ownerNumber, password }, false);
      if (!r.ok) return loginView(r.error || 'Login failed');
      TOKEN = r.token; localStorage.setItem('gp_token', TOKEN);
      await loadCatalog(); // load rewards
      dashboard();
    };
  }

  /* ------------------------- Catalog (Rewards) ------------------------- */
  let CATALOG = [];     // from backend
  let CATALOG_MAP = {}; // sku -> spec
  async function loadCatalog(){
    const r = await call('rewards');
    if (!r.ok) throw new Error(r.error || 'Failed to load rewards');
    CATALOG = r.items || [];
    CATALOG_MAP = Object.fromEntries(CATALOG.map(i => [i.sku, i]));
  }

  /* ------------------------- Dashboard ------------------------- */
  let ME = null;
  async function dashboard(){
    const me = await call('me');
    if (!me.ok) { localStorage.removeItem('gp_token'); return loginView(me.error || 'Session expired—please log in.'); }
    ME = me.profile;

    mount(`
      <div class="gp-hero">
        <h2>Hi ${esc(ME.name || 'Owner')}</h2>
        <button class="gp-btn secondary" id="logoutBtn">Logout</button>
      </div>

      <div class="gp-grid" style="grid-template-columns:1fr;">
        <div class="gp-card">
          <h3 style="margin-top:0">My details</h3>
          <div class="metrics" style="margin-top:8px">
            <div class="metric"><b>Owner #</b><div>${esc(ME.ownerNumber)}</div></div>
            <div class="metric"><b>Joined</b><div>${esc(ME.joinedYear || '-')}</div></div>
            <div class="metric"><b>Total viewings</b><div id="viewCount">${ME.totals.viewings ?? '-'}</div></div>
            <div class="metric">
              <b>Reward points</b>
              <div class="points-callout">
                <span id="pointsNow">${ME.totals.rewardPoints}</span>
                <button class="gp-btn" id="startRedeem">Redeem your points</button>
              </div>
            </div>
          </div>
          <div class="gp-row" style="margin-top:10px">
            <div style="flex:0 0 auto;display:flex;align-items:center;gap:10px">
              <span class="gp-muted">Receive viewing requests:</span>
              <div class="switch" id="activeSwitch"><span></span></div>
              <span id="activeText" class="gp-muted"></span>
            </div>
          </div>
          <div id="detailsMsg"></div>
        </div>
      </div>

      <div class="gp-card" id="viewingsCard">
        <h3>Your viewings</h3>
        <div style="height:1px;background:var(--gp-border);margin:8px 0"></div>
        <table class="gp-table">
          <thead>
            <tr>
              <th>Viewing Date</th>
              <th>Viewer</th>
              <th>Status</th>
              <th>Points awarded</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody id="viewingsTable"><tr><td colspan="5" class="gp-muted" style="padding:12px">Loading…</td></tr></tbody>
        </table>
      </div>

      <div class="gp-card">
        <h3>Log an "impromptu" viewing</h3>
        <p class="gp-muted" style="margin-top:-6px">Log a viewing that took place outside of the system <i>(I.e. a viewing that didn't go through the form on Go-Pods.co.uk)</i> this could be viewings with interested friends/family, while you're away in your Go-Pod etc. We'll review the requests and award points once we've made contact with the viewers.</p>
        <div class="formgrid-2">
          <div><label>Viewing date</label><input id="iv-date" class="gp-input" type="date" value="${toISO(new Date())}"></div>
          <div><label>Location</label><input id="iv-location" class="gp-input" placeholder="e.g. CAMC Chatsworth Park"></div>
        </div>
        <div class="formgrid-3">
          <div><label>Viewer name</label><input id="iv-name" class="gp-input"></div>
          <div><label>Viewer email</label><input id="iv-email" class="gp-input" type="email"></div>
          <div><label>Viewer phone</label><input id="iv-phone" class="gp-input" type="tel"></div>
        </div>
        <div class="gp-row">
          <div><label>Notes (optional)</label><input id="iv-notes" class="gp-input" placeholder="Anything useful to know…"></div>
        </div>
        <div class="gp-row"><button class="gp-btn" id="iv-submit">Submit viewing</button></div>
        <div id="iv-msg"></div>
      </div>

      <div class="gp-card">
        <h3>Change password</h3>
        <div class="formgrid-3">
          <div><label>Current password</label><input id="cp-old" class="gp-input" type="password"></div>
          <div><label>New password</label><input id="cp-new" class="gp-input" type="password"></div>
          <div><label>Confirm new password</label><input id="cp-new2" class="gp-input" type="password"></div>
        </div>
        <div class="gp-row"><button class="gp-btn" id="cp-save">Save new password</button></div>
        <div id="cp-msg"></div>
      </div>
    `);

    document.getElementById('logoutBtn').onclick = async ()=>{ await call('logout',{}); localStorage.removeItem('gp_token'); loginView(); };

    // Active toggle
    const activeSwitch = document.getElementById('activeSwitch');
    const activeText = document.getElementById('activeText');
    setSwitch(!!ME.isActive);
    activeSwitch.onclick = async ()=>{
      const to = !activeSwitch.classList.contains('on');
      setSwitch(to);
      const resp = await call('setactive', { active: to });
      if (!resp.ok) {
        setSwitch(!to);
        document.getElementById('detailsMsg').innerHTML = `<div class="gp-error">${esc(resp.error || 'Couldn’t update status')}</div>`;
      } else {
        document.getElementById('detailsMsg').innerHTML = `<div class="gp-success">Status updated.</div>`;
      }
    };
    function setSwitch(on){ activeSwitch.classList.toggle('on', !!on); activeText.textContent = on ? 'Active' : 'Inactive'; }

    // Redeem flow
    document.getElementById('startRedeem').onclick = ()=> redeemFlow();

    // Impromptu submission
    document.getElementById('iv-submit').onclick = async ()=>{
      const btn   = document.getElementById('iv-submit');
      const dateISO = document.getElementById('iv-date').value;
      const loc   = document.getElementById('iv-location').value.trim();
      const name  = document.getElementById('iv-name').value.trim();
      const email = document.getElementById('iv-email').value.trim();
      const phone = document.getElementById('iv-phone').value.trim();
      const notes = document.getElementById('iv-notes').value.trim();
      const box   = document.getElementById('iv-msg');
      box.innerHTML = '';

      if (!dateISO || !name || (!email && !phone)) {
        box.innerHTML = `<div class="gp-error">Please provide a date, viewer name and at least one contact method.</div>`;
        return;
      }

      btn.disabled = true;

      const payload = {
        ownerNumber: ME.ownerNumber,
        viewerName: name,
        viewerEmail: email,
        viewerPhone: phone,
        requestedDate: dateISO,
        notes: [`Impromptu at: ${loc || '(unspecified)'}`, notes].filter(Boolean).join(' | '),
        source: 'impromptu'
      };

      const r = await call('createviewingrequest', payload, false);

      if (!r.ok) {
        box.innerHTML = `<div class="gp-error">${esc(r.error || 'Failed to submit')}</div>`;
        btn.disabled = false;
        return;
      }

      // Clear fields & reset
      document.getElementById('iv-location').value = '';
      document.getElementById('iv-name').value = '';
      document.getElementById('iv-email').value = '';
      document.getElementById('iv-phone').value = '';
      document.getElementById('iv-notes').value = '';
      document.getElementById('iv-date').value = toISO(new Date());

      box.innerHTML = `<div class="gp-success"><span class="check">✓</span>&nbsp;Thanks! Logged as TBC for admin review.</div>`;
      btn.disabled = false;

      // Refresh table so the new row appears
      loadViewings();
    };

    // Change password
    document.getElementById('cp-save').onclick = async ()=>{
      const oldP = document.getElementById('cp-old').value;
      const np1  = document.getElementById('cp-new').value;
      const np2  = document.getElementById('cp-new2').value;
      const box  = document.getElementById('cp-msg');
      box.innerHTML = '';
      if (!oldP || !np1 || np1 !== np2) { box.innerHTML = `<div class="gp-error">Please fill all fields and ensure passwords match.</div>`; return; }
      const r = await call('changepassword', { oldPassword: oldP, newPassword: np1 });
      if (!r.ok) { box.innerHTML = `<div class="gp-error">${esc(r.error || 'Couldn’t change password')}</div>`; return; }
      box.innerHTML = `<div class="gp-success">Password updated.</div>`;
      ['cp-old','cp-new','cp-new2'].forEach(id=>document.getElementById(id).value='');
    };

    loadViewings();
  }

  /* ------------------------- Viewings ------------------------- */
  async function loadViewings(){
    const body = document.getElementById('viewingsTable');
    const r = await call('viewings');
    if (!r.ok) { body.innerHTML = `<tr><td colspan="5"><div class="gp-error">Couldn’t load viewings: ${esc(r.error||'')}</div></td></tr>`; return; }
    if (!r.items.length) { body.innerHTML = `<tr><td colspan="5" class="gp-muted" style="padding:12px">No viewings yet.</td></tr>`; return; }

    const rows = r.items.map(v => {
      const id  = esc(v.viewingId);
      const iso = toISO(v.viewingDate);
      const uk  = toUK(v.viewingDate);
      const statusKey = String(v.status || '').toUpperCase();
      const canComplete = (statusKey === 'ARRANGED');
      return `
        <tr data-row="${id}">
          <td>
            <div class="date-view" data-date-view="${id}">
              <strong>${uk}</strong>
              <button class="link" data-date-edit="${id}">Change viewing date</button>
            </div>
            <div class="date-editor" data-date-editor="${id}" style="display:none;gap:8px;align-items:center">
              <input type="date" class="gp-input" value="${iso}" data-date-input="${id}" style="max-width:220px">
              <button class="gp-btn secondary small" data-date-save="${id}">Save</button>
              <button class="gp-btn secondary small" data-date-cancel="${id}">Cancel</button>
            </div>
            <div class="date-done gp-muted" data-date-done="${id}" style="display:none">
              <span class="check">✓</span>&nbsp;Viewing date changed. <button class="link" data-date-edit-again="${id}">Change again</button>
            </div>
          </td>
          <td>${esc(v.viewerName || '')}</td>
          <td>${statusPillHTML(statusKey, id)}</td>
          <td><span id="pts-${id}">${(v.pointsAllocated ?? '').toString()}</span></td>
          <td>
            ${canComplete
              ? `<button class="gp-btn secondary small" data-complete="${id}">Mark viewing as complete</button>
                 <span class="gp-pill done" data-complete-done="${id}" style="display:none"><span class="check">✓</span> Completed</span>`
              : `<span class="gp-pill done" ${statusKey==='TBC'?'style="display:none"':''}><span class="check">✓</span> Completed</span>`}
          </td>
        </tr>
        <tr class="gp-subrow">
          <td colspan="5">
            <div class="subattach">
              <div class="gp-subflex">
                <button class="gp-pill" data-fbopen="${id}">Provide feedback for this viewing</button>
                <div class="gp-muted">Opens a quick text box below</div>
              </div>
              <div class="feedback-box" id="fb-${id}" style="display:none">
                <textarea class="gp-input" rows="3" placeholder="Enter your feedback…"></textarea>
                <div class="gp-row" style="margin-top:8px">
                  <button class="gp-btn small" data-fbsend="${id}">Send feedback</button>
                  <button class="gp-btn secondary small" data-fbcancel="${id}">Cancel</button>
                </div>
                <div class="gp-muted" style="margin-top:4px">This will be emailed to sales@go-pods.co.uk and rachel@go-pods.co.uk and added to our sheet.</div>
              </div>
            </div>
          </td>
        </tr>
      `;
    }).join('');
    body.innerHTML = rows;

    // Date edit UI
    const showEditor = id => { sel(`[data-date-view="${CSS.escape(id)}"]`).style.display='none'; sel(`[data-date-done="${CSS.escape(id)}"]`).style.display='none'; sel(`[data-date-editor="${CSS.escape(id)}"]`).style.display='flex'; };
    const showView   = id => { sel(`[data-date-editor="${CSS.escape(id)}"]`).style.display='none'; sel(`[data-date-done="${CSS.escape(id)}"]`).style.display='none'; sel(`[data-date-view="${CSS.escape(id)}"]`).style.display='flex'; };
    const showDone   = id => { sel(`[data-date-editor="${CSS.escape(id)}"]`).style.display='none'; sel(`[data-date-view="${CSS.escape(id)}"]`).style.display='none'; sel(`[data-date-done="${CSS.escape(id)}"]`).style.display='inline-block'; };
    function sel(q){ return body.querySelector(q); }

    body.querySelectorAll('[data-date-edit],[data-date-edit-again]').forEach(b=> b.onclick = ()=> showEditor(b.dataset.dateEdit || b.dataset.dateEditAgain));
    body.querySelectorAll('[data-date-cancel]').forEach(b=> b.onclick = ()=> showView(b.dataset.dateCancel));
    body.querySelectorAll('[data-date-save]').forEach(b=>{
      b.onclick = async ()=>{
        const id = b.dataset.dateSave;
        const inp = sel(`[data-date-input="${CSS.escape(id)}"]`);
        if (!inp.value) return;
        b.disabled = true;
        const resp = await call('updateviewingdate', { viewingId:id, dateISO: inp.value });
        if (!resp.ok) { alert(resp.error || 'Failed to update date'); b.disabled = false; return; }
        const uk = toUK(inp.value);
        const view = sel(`[data-date-view="${CSS.escape(id)}"]`); if (view) view.querySelector('strong').textContent = uk;
        b.disabled = false; showDone(id);
      };
    });

    // Mark complete
    body.querySelectorAll('[data-complete]').forEach(btn=>{
      btn.onclick = async ()=>{
        const id = btn.dataset.complete;
        btn.disabled = true;
        const r2 = await call('confirmviewing', { viewingId:id });
        if (!r2.ok) { alert(r2.error || 'Failed'); btn.disabled = false; return; }
        btn.style.display='none';
        const done = body.querySelector(`[data-complete-done="${CSS.escape(id)}"]`); if (done) done.style.display='inline-flex';
        const statusEl = document.getElementById(`status-${id}`);
        const ptsEl = document.getElementById(`pts-${id}`);
        if (statusEl) {
          statusEl.className = 'status ' + statusClassFor('VIEWED');
          statusEl.innerHTML = '<i>✓</i><span>VIEWED</span>';
        }
        if (ptsEl) ptsEl.textContent = '0.25';
        const me2 = await call('me'); if (me2.ok){ document.getElementById('pointsNow').textContent = me2.profile.totals.rewardPoints; document.getElementById('viewCount').textContent = me2.profile.totals.viewings ?? '-'; }
      };
    });

    // Feedback UI
    body.querySelectorAll('[data-fbopen]').forEach(btn=>{
      btn.onclick = ()=>{
        const id = btn.dataset.fbopen;
        const box = document.getElementById(`fb-${id}`);
        box.style.display = (box.style.display==='none'||!box.style.display)?'block':'none';
      };
    });
    body.querySelectorAll('[data-fbcancel]').forEach(btn=>{
      btn.onclick = ()=>{
        const id = btn.dataset.fbcancel;
        const box = document.getElementById(`fb-${id}`);
        box.querySelector('textarea').value='';
        box.style.display='none';
      };
    });
    body.querySelectorAll('[data-fbsend]').forEach(btn=>{
      btn.onclick = async ()=>{
        const id = btn.dataset.fbsend;
        const box = document.getElementById(`fb-${id}`);
        const txt = box.querySelector('textarea').value.trim();
        if (!txt) return alert('Please enter some feedback.');
        btn.disabled = true;
        const r = await call('ownerfeedback', { viewingId:id, feedback: txt });
        if (!r.ok) { alert(r.error || 'Failed to send feedback'); btn.disabled = false; return; }
        box.querySelector('textarea').value=''; box.style.display='none';
      };
    });
  }

  /* ------------------------- Rewards Flow (Live Catalog) ------------------------- */
  function redeemFlow(){
    // state.items will be [{sku, qty}]
    const state = {
      step: 1,
      items: [],
      shipping: { line1:'', line2:'', town:'', postcode:'', phone:'' },
      collectAtFitting: false,
      chassisNumber: '',
      preferredDateISO: ''
    };

    const pointsAvailable = ()=> Number(document.getElementById('pointsNow')?.textContent || ME.totals.rewardPoints || 0);
    const itemCost = ({sku, qty}) => (CATALOG_MAP[sku]?.points || 0) * qty;
    const basketTotal = () => state.items.reduce((s,it)=> s + itemCost(it), 0);
    const clampQty = (qty, max) => {
      const q = Math.max(0, Math.floor(qty || 0));
      return max > 0 ? Math.min(q, max) : q;
    };
    const hasFitting = () => state.items.some(it => CATALOG_MAP[it.sku]?.requiresFitting);
    const hasDelivery = () => state.items.some(it => !CATALOG_MAP[it.sku]?.requiresFitting);

    const render = ()=>{
      const steps = `
        <div class="stepper">
          <div class="step ${state.step===1?'active':''}">1. Choose rewards</div>
          <div class="step ${state.step===2?'active':''}">2. Delivery & fitting</div>
          <div class="step ${state.step===3?'active':''}">3. Review & submit</div>
        </div>`;

      // STEP 1
      if (state.step === 1){
        mount(`
          <div class="gp-hero">
            <h2>Redeem points</h2>
            <button class="gp-btn secondary" id="backDash">Back to dashboard</button>
          </div>
          <div class="gp-grid" style="grid-template-columns:1fr;">
            <div class="gp-card">
              ${steps}
              <div class="points-callout" style="margin:8px 0 12px 0">
                <span>Available points: <b id="availPts">${esc(pointsAvailable().toFixed(2))}</b></span>
                <span>Basket total: <b id="basketPts">0.00</b></span>
              </div>
              <div class="grid" id="rewardsGrid"></div>
              <div id="catMsg" class="gp-error" style="display:none;margin-top:8px"></div>
              <div class="gp-row"><button class="gp-btn" id="toShip" disabled>Next: Delivery & fitting</button></div>
            </div>
          </div>
        `);
        document.getElementById('backDash').onclick = dashboard;

        const grid = document.getElementById('rewardsGrid');
        grid.innerHTML = CATALOG.map((p,i)=>`
          <div class="prod">
            <img src="${esc(p.imageUrl || '')}" alt="">
            <div class="pbody">
              <strong>${esc(p.name)}</strong>
              <div class="gp-muted">${esc(p.description || '')}</div>
              <div class="gp-row" style="margin:0;align-items:center">
                <div>
                  <span class="gp-pill">Pts: ${Number(p.points||0).toFixed(2)}</span>
                  ${p.maxPerOrder>0 ? `<span class="gp-muted" style="margin-left:8px">Max ${p.maxPerOrder}/order</span>` : ''}
                  ${p.requiresFitting ? `<span class="gp-pill warn" style="margin-left:8px">Requires workshop fitting</span>` : ''}
                </div>
                <div style="display:flex;gap:6px;align-items:center">
                  <input type="number" min="1" step="1" value="1" class="gp-input" style="max-width:90px" id="qty-${i}">
                  <button class="gp-btn small" data-add="${p.sku}">Add</button>
                </div>
              </div>
            </div>
          </div>
        `).join('');

        const basketPts = document.getElementById('basketPts');
        const toShip = document.getElementById('toShip');
        const catMsg = document.getElementById('catMsg');

        function refreshBasketUI(){
          const total = basketTotal();
          basketPts.textContent = total.toFixed(2);
          toShip.disabled = total <= 0;
          const avail = pointsAvailable();
          catMsg.style.display = (total > avail) ? 'block' : 'none';
          catMsg.textContent = (total > avail) ? `Basket exceeds your points (${total.toFixed(2)} > ${avail.toFixed(2)}). Remove some items.` : '';
        }

        grid.querySelectorAll('[data-add]').forEach(btn=>{
          btn.onclick = ()=>{
            const sku = btn.dataset.add;
            const spec = CATALOG_MAP[sku];
            if (!spec) return;
            const idx = CATALOG.findIndex(x=>x.sku===sku);
            const qtyInput = grid.querySelector(`#qty-${idx}`);
            const want = Math.max(1, Number(qtyInput?.value || 1));
            const current = state.items.find(it => it.sku === sku)?.qty || 0;
            const nextQty = clampQty(current + want, Number(spec.maxPerOrder||0));
            if (nextQty === current) return; // at max already
            const existing = state.items.find(it => it.sku === sku);
            if (existing) existing.qty = nextQty; else state.items.push({ sku, qty: nextQty });
            refreshBasketUI();
          };
        });

        toShip.onclick = ()=>{
          if (basketTotal() <= 0) return;
          if (basketTotal() > pointsAvailable()) {
            catMsg.style.display = 'block';
            catMsg.textContent = 'Basket exceeds your available points.';
            return;
          }
          state.step = 2; render();
        };
      }

      // STEP 2
      else if (state.step === 2){
        const delivery = hasDelivery();
        const fitting  = hasFitting();

        mount(`
          <div class="gp-hero">
            <h2>Redeem points</h2>
            <button class="gp-btn secondary" id="backTo1">Back</button>
          </div>
          <div class="gp-grid" style="grid-template-columns:1fr;">
            <div class="gp-card">
              ${steps}

              ${delivery ? `
                <div id="collect-at-fitting-wrap" style="margin:8px 0; ${fitting?'':'display:none'}">
                  <label><input type="checkbox" id="collectAtFitting" ${state.collectAtFitting?'checked':''}> Collect delivery items at your fitting appointment</label>
                </div>
                <div id="shipping-fields" style="margin:12px 0; ${state.collectAtFitting?'display:none':''}">
                  <div class="formgrid-2">
                    <input id="addr1" class="gp-input" placeholder="Address line 1 *" value="${esc(state.shipping.line1)}">
                    <input id="addr2" class="gp-input" placeholder="Address line 2 (optional)" value="${esc(state.shipping.line2)}">
                  </div>
                  <div class="formgrid-2">
                    <input id="town" class="gp-input" placeholder="Town/City *" value="${esc(state.shipping.town)}">
                    <input id="postcode" class="gp-input" placeholder="Postcode *" value="${esc(state.shipping.postcode)}">
                  </div>
                  <div class="gp-row">
                    <input id="phone" class="gp-input" placeholder="Phone (optional)" value="${esc(state.shipping.phone)}">
                  </div>
                </div>
              ` : `
                <div class="gp-muted" style="margin:8px 0">No delivery items in your basket.</div>
              `}

              ${fitting ? `
                <div id="fitting-fields" style="margin:12px 0">
                  <div class="formgrid-2">
                    <input id="chassis" class="gp-input" placeholder="Chassis number *" value="${esc(state.chassisNumber)}">
                    <input id="prefdate" class="gp-input" type="date" value="${esc(state.preferredDateISO || '')}">
                  </div>
                  <div class="gp-muted" style="margin-top:4px">Catherine will confirm your fitting appointment by email.</div>
                </div>
              ` : `
                <div class="gp-muted" style="margin:8px 0">No workshop-fitting items in your basket.</div>
              `}

              <div id="shipMsg" class="gp-error" style="display:none;margin-top:8px"></div>
              <div class="gp-row"><button class="gp-btn" id="toReview">Next: Review</button></div>
            </div>
          </div>
        `);

        document.getElementById('backTo1').onclick = ()=>{ state.step = 1; render(); };

        const collectBox = document.getElementById('collectAtFitting');
        if (collectBox) collectBox.onchange = ()=>{
          state.collectAtFitting = !!collectBox.checked;
          render();
        };

        document.getElementById('toReview').onclick = ()=>{
          const err = (msg)=>{ const box = document.getElementById('shipMsg'); box.style.display='block'; box.textContent = msg; };

          if (hasDelivery() && !state.collectAtFitting){
            state.shipping.line1 = document.getElementById('addr1').value.trim();
            state.shipping.line2 = document.getElementById('addr2').value.trim();
            state.shipping.town  = document.getElementById('town').value.trim();
            state.shipping.postcode = document.getElementById('postcode').value.trim();
            state.shipping.phone = document.getElementById('phone').value.trim();
            if (!state.shipping.line1 || !state.shipping.town || !state.shipping.postcode){
              return err('Please complete address line 1, town/city and postcode.');
            }
          }

          if (hasFitting()){
            state.chassisNumber = document.getElementById('chassis').value.trim();
            state.preferredDateISO = document.getElementById('prefdate').value || '';
            if (!state.chassisNumber) return err('Please provide your chassis number (needed for workshop booking).');
          }

          state.step = 3; render();
        };
      }

      // STEP 3
      else if (state.step === 3){
        const before = pointsAvailable();
        const total  = basketTotal();
        const after  = (before - total).toFixed(2);

        mount(`
          <div class="gp-hero">
            <h2>Redeem points</h2>
            <button class="gp-btn secondary" id="backTo2">Back</button>
          </div>
          <div class="gp-grid" style="grid-template-columns:1fr;">
            <div class="gp-card">
              ${steps}
              <h3 style="margin:6px 0 10px">Review your order</h3>
              ${state.items.map(it => {
                const spec = CATALOG_MAP[it.sku] || {};
                const subtotal = (Number(spec.points||0) * it.qty).toFixed(2);
                return `
                  <div class="gp-row" style="align-items:center">
                    <div style="flex:2"><strong>${esc(it.sku)}</strong> — ${esc(spec.name || '')}</div>
                    <div style="flex:1">Pts each: ${Number(spec.points||0).toFixed(2)}</div>
                    <div style="flex:1">Qty: ${it.qty}</div>
                    <div style="flex:1">Subtotal: ${subtotal}</div>
                  </div>
                `;
              }).join('')}
              <div class="gp-divider"></div>
              <div class="points-callout"><span>Total points to spend: <b>${total.toFixed(2)}</b></span><span>After order: <b>${after}</b></span></div>

              ${hasDelivery() ? `
                <h4 style="margin:14px 0 6px">Delivery</h4>
                <div class="gp-muted">
                  ${state.collectAtFitting ? 'Collect delivery items at fitting appointment' : `
                    ${esc(state.shipping.line1)} ${esc(state.shipping.line2)}<br>
                    ${esc(state.shipping.town)}<br>
                    ${esc(state.shipping.postcode)}<br>
                    ${state.shipping.phone ? `Phone: ${esc(state.shipping.phone)}` : ''}
                  `}
                </div>
              `:''}

              ${hasFitting() ? `
                <h4 style="margin:14px 0 6px">Workshop fitting</h4>
                <div class="gp-muted">
                  Chassis: ${esc(state.chassisNumber)}<br>
                  ${state.preferredDateISO ? `Preferred date: ${esc(state.preferredDateISO)}` : 'Preferred date: (not specified)'}
                </div>
              `:''}

              <div class="gp-row"><button class="gp-btn" id="submitOrder">Place order</button></div>
              <div id="orderMsg"></div>
            </div>
          </div>
        `);

        document.getElementById('backTo2').onclick = ()=>{ state.step = 2; render(); };

        document.getElementById('submitOrder').onclick = async ()=>{
          const msg = document.getElementById('orderMsg'); msg.className=''; msg.textContent='';

          if (basketTotal() > pointsAvailable()){
            msg.className='gp-error'; msg.textContent='Basket exceeds your available points.'; return;
          }

          const items = state.items.map(it => ({ sku: it.sku, qty: it.qty }));
          const shipping = (hasDelivery() && !state.collectAtFitting)
            ? { line1: state.shipping.line1, line2: state.shipping.line2, town: state.shipping.town, postcode: state.shipping.postcode, phone: state.shipping.phone }
            : null;
          const workshop = hasFitting()
            ? { chassisNumber: state.chassisNumber, preferredDateISO: state.preferredDateISO || null }
            : null;

          const r = await call('redeem', {
            items,
            shipping,
            collectAtFitting: !!state.collectAtFitting,
            workshop
          });

          if (!r.ok){
            msg.className='gp-error';
            msg.textContent = r.error || 'Order failed';
            return;
          }

          msg.className='gp-success';
          msg.innerHTML = `Order submitted! Points: <b>${r.pointsBefore}</b> → <b>${r.pointsAfter}</b>.`;

          setTimeout(async ()=>{
            const me2 = await call('me');
            if (me2.ok){ document.getElementById('pointsNow').textContent = me2.profile.totals.rewardPoints; }
            dashboard();
          }, 1200);
        };
      }
    };

    render();
  }

  // Init
  (async ()=>{
    try{
      if (TOKEN) {
        const me = await call('me');
        if (me.ok) {
          await loadCatalog();
          return dashboard();
        }
        localStorage.removeItem('gp_token');
      }
      loginView();
    } catch (e){
      mount(`<div class="gp-card"><div class="gp-error">Init error: ${esc(e.message||e)}</div></div>`);
    }
  })();
})();
