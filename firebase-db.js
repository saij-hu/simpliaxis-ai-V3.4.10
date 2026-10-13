/* ============================================================
   FIREBASE / FIRESTORE — overall dashboard shared data layer
   ------------------------------------------------------------
   This layer is intentionally additive and non-destructive:
   - Google Sheets remains the source for live analytical rows.
   - IndexedDB remains the browser cache for large live datasets/snapshots.
   - Firestore becomes the shared cloud store for operational dashboard state:
       users (metadata only), tickets, updates, notifications, settings,
       sales-target mirrors, and live dashboard metadata.
   - When Firebase is not configured, all existing local behavior remains intact.
   ============================================================ */
(function(){
  'use strict';

  const STORAGE_KEY = 'pulse_firebase_config';
  const COLLECTION_LIMIT = 2500;
  let db = null;
  let app = null;
  let initialized = false;
  let syncing = false;
  const listeners = new Map();

  function safeConfig(raw){
    if(!raw || typeof raw!=='object') return null;
    const cfg = {
      apiKey: String(raw.apiKey||'').trim(),
      authDomain: String(raw.authDomain||'').trim(),
      projectId: String(raw.projectId||'').trim(),
      storageBucket: String(raw.storageBucket||'').trim(),
      messagingSenderId: String(raw.messagingSenderId||'').trim(),
      appId: String(raw.appId||'').trim()
    };
    return cfg.apiKey && cfg.authDomain && cfg.projectId && cfg.appId ? cfg : null;
  }

  function currentConfig(){
    let cfg = safeConfig(window.SIMPLIAXIS_FIREBASE_CONFIG);
    if(cfg) return cfg;
    try{ cfg = safeConfig(JSON.parse(localStorage.getItem(STORAGE_KEY)||'null')); }catch(e){}
    return cfg;
  }

  function configured(){ return !!currentConfig() && typeof window.firebase!=='undefined' && !!window.firebase.firestore; }

  function status(){
    return {
      configured: configured(),
      initialized,
      projectId: currentConfig()?.projectId || '',
      mode: configured() ? 'FIRESTORE' : 'LOCAL FALLBACK'
    };
  }

  function logWarn(...args){ try{ console.warn('[FirebaseDB]', ...args); }catch(e){} }
  function logError(...args){ try{ console.error('[FirebaseDB]', ...args); }catch(e){} }

  async function init(){
    if(initialized) return status();
    const cfg = currentConfig();
    if(!cfg || typeof window.firebase==='undefined' || !window.firebase.firestore){
      initialized = true;
      return status();
    }
    try{
      app = window.firebase.apps && window.firebase.apps.length ? window.firebase.app() : window.firebase.initializeApp(cfg);
      db = window.firebase.firestore(app);
      // Offline persistence is best-effort. Firestore will continue online if the browser
      // or browser context does not permit persistence.
      try{
        if(db.enablePersistence) db.enablePersistence({synchronizeTabs:true}).catch(()=>{});
      }catch(e){}
      initialized = true;
      await bootstrapOperationalCollections();
      startRealtimeListeners();
      return status();
    }catch(e){
      logError('Initialization failed; existing local dashboard behavior remains active.', e);
      initialized = true;
      return status();
    }
  }

  function collection(name){
    if(!db) throw new Error('Firebase is not connected');
    return db.collection(name);
  }

  function cleanUser(u){
    const x = {...(u||{})};
    // Never upload password material, OTPs, lock internals, or other credential secrets.
    ['password','passwordEnc','passwordHash','pendingOtp','pendingOtpExpires','failedAttempts','lockUntil'].forEach(k=>delete x[k]);
    return x;
  }

  function stripUndefined(obj){
    const out = {};
    Object.keys(obj||{}).forEach(k=>{ if(obj[k]!==undefined) out[k]=obj[k]; });
    return out;
  }

  async function upsert(name, id, value){
    if(!db || !id) return false;
    try{
      await collection(name).doc(String(id)).set(stripUndefined({...value, _firebaseUpdatedAt: window.firebase.firestore.FieldValue.serverTimestamp()}), {merge:true});
      return true;
    }catch(e){ logWarn(`upsert ${name}/${id} failed`, e); return false; }
  }

  async function remove(name, id){
    if(!db || !id) return false;
    try{ await collection(name).doc(String(id)).delete(); return true; }
    catch(e){ logWarn(`remove ${name}/${id} failed`, e); return false; }
  }

  async function batchUpsert(name, rows, idFn, sanitizer){
    if(!db || !Array.isArray(rows) || !rows.length) return 0;
    let written = 0;
    const size = 450; // safely below Firestore's 500-operation batch limit
    for(let start=0; start<rows.length; start+=size){
      const batch = db.batch();
      const chunk = rows.slice(start, start+size);
      chunk.forEach((row, idx)=>{
        const id = idFn(row, start+idx);
        if(!id) return;
        const value = sanitizer ? sanitizer(row) : row;
        batch.set(collection(name).doc(String(id)), stripUndefined({...value, _firebaseUpdatedAt: window.firebase.firestore.FieldValue.serverTimestamp()}), {merge:true});
      });
      try{ await batch.commit(); written += chunk.length; }
      catch(e){ logWarn(`batchUpsert ${name} failed at ${start}`, e); }
    }
    return written;
  }

  async function read(name, limit=COLLECTION_LIMIT){
    if(!db) return [];
    try{
      const snap = await collection(name).orderBy('_firebaseUpdatedAt','desc').limit(limit).get();
      return snap.docs.map(d=>({id:d.id,...d.data()}));
    }catch(orderErr){
      try{
        const snap = await collection(name).limit(limit).get();
        return snap.docs.map(d=>({id:d.id,...d.data()}));
      }catch(e){ logWarn(`read ${name} failed`, e); return []; }
    }
  }

  async function replaceMirror(name, rows, idFn, sanitizer){
    return batchUpsert(name, rows, idFn, sanitizer);
  }

  async function bootstrapOperationalCollections(){
    if(!db || syncing) return;
    syncing = true;
    try{
      // Updates / tickets / notifications are operational data and should become shared.
      // If Firestore is empty, seed from the current browser state; otherwise remote data
      // is merged into local state without deleting local-only items.
      const [remoteUpdates, remoteTickets, remoteNotifications] = await Promise.all([
        read('updates', 250), read('tickets', 500), read('notifications', 500)
      ]);

      if(remoteUpdates.length){
        const byId = new Map((STATE.updates||[]).map(x=>[String(x.id),x]));
        remoteUpdates.forEach(x=>byId.set(String(x.id), {...byId.get(String(x.id)), ...x}));
        STATE.updates = [...byId.values()].sort((a,b)=>new Date(b.dateTime||b.createdAt||0)-new Date(a.dateTime||a.createdAt||0));
        try{ saveUpdates(); }catch(e){}
      }else if(STATE.updates?.length){ await batchUpsert('updates', STATE.updates, x=>x.id); }

      if(remoteTickets.length){
        const byId = new Map((STATE.tickets||[]).map(x=>[String(x.id),x]));
        remoteTickets.forEach(x=>byId.set(String(x.id), {...byId.get(String(x.id)), ...x}));
        STATE.tickets = [...byId.values()].sort((a,b)=>new Date(b.createdAt||0)-new Date(a.createdAt||0));
        try{ saveTickets(); }catch(e){}
      }else if(STATE.tickets?.length){ await batchUpsert('tickets', STATE.tickets, x=>x.id); }

      if(remoteNotifications.length){
        const byId = new Map((STATE.notifications||[]).map(x=>[String(x.id),x]));
        remoteNotifications.forEach(x=>byId.set(String(x.id), {...byId.get(String(x.id)), ...x}));
        STATE.notifications = [...byId.values()].sort((a,b)=>new Date(b.ts||0)-new Date(a.ts||0)).slice(0,300);
        try{ saveNotifications(); }catch(e){}
      }else if(STATE.notifications?.length){ await batchUpsert('notifications', STATE.notifications, x=>x.id); }

      // Mirror targets, but keep the existing Target Management shared endpoint as the
      // authoritative write/read path until Firebase Auth/rules are configured for that module.
      if(Array.isArray(STATE.tmSalesVsTarget) && STATE.tmSalesVsTarget.length){
        await batchUpsert('targets', STATE.tmSalesVsTarget, x=>`${x.month||'unknown'}__${x.cre||'unknown'}`);
      }
    }finally{
      syncing = false;
    }
  }

  function startRealtimeListeners(){
    if(!db) return;
    subscribe('tickets', snap=>{
      const remote = snap.docs.map(d=>({id:d.id,...d.data()}));
      if(!remote.length) return;
      const byId = new Map((STATE.tickets||[]).map(x=>[String(x.id),x]));
      remote.forEach(x=>byId.set(String(x.id), {...byId.get(String(x.id)), ...x}));
      STATE.tickets = [...byId.values()].sort((a,b)=>new Date(b.createdAt||0)-new Date(a.createdAt||0));
      try{ if(typeof renderTicketsPage==='function') renderTicketsPage(); }catch(e){}
    });
    subscribe('updates', snap=>{
      const remote = snap.docs.map(d=>({id:d.id,...d.data()}));
      if(!remote.length) return;
      const byId = new Map((STATE.updates||[]).map(x=>[String(x.id),x]));
      remote.forEach(x=>byId.set(String(x.id), {...byId.get(String(x.id)), ...x}));
      STATE.updates = [...byId.values()].sort((a,b)=>new Date(b.dateTime||b.createdAt||0)-new Date(a.dateTime||a.createdAt||0));
      try{ if(typeof renderUpdatesPage==='function') renderUpdatesPage(); }catch(e){}
    });
    subscribe('notifications', snap=>{
      const remote = snap.docs.map(d=>({id:d.id,...d.data()}));
      if(!remote.length) return;
      const byId = new Map((STATE.notifications||[]).map(x=>[String(x.id),x]));
      remote.forEach(x=>byId.set(String(x.id), {...byId.get(String(x.id)), ...x}));
      STATE.notifications = [...byId.values()].sort((a,b)=>new Date(b.ts||0)-new Date(a.ts||0)).slice(0,300);
      try{ saveNotifications(); updateNotifBadge(); renderNotifPanel(); }catch(e){}
    });
  }

  function subscribe(name, onNext){
    if(!db || listeners.has(name)) return;
    try{
      const unsub = collection(name).limit(500).onSnapshot(onNext, e=>logWarn(`listener ${name} failed`, e));
      listeners.set(name, unsub);
    }catch(e){ logWarn(`subscribe ${name} failed`, e); }
  }

  async function saveUsers(users){
    if(!db) return 0;
    return batchUpsert('users', users||[], x=>(x.username||x.email||x.id||'').toLowerCase(), cleanUser);
  }

  async function saveTickets(tickets){
    if(!db) return 0;
    return batchUpsert('tickets', tickets||[], x=>x.id);
  }

  async function saveUpdates(updates){
    if(!db) return 0;
    return batchUpsert('updates', updates||[], x=>x.id);
  }

  async function saveNotifications(items){
    if(!db) return 0;
    return batchUpsert('notifications', (items||[]).slice(0,300), x=>x.id);
  }

  async function saveTargets(rows){
    if(!db) return 0;
    return batchUpsert('targets', rows||[], x=>`${x.month||'unknown'}__${x.cre||'unknown'}`);
  }

  async function saveSettings(username, settings){
    if(!db || !username) return false;
    return upsert('settings', String(username).toLowerCase(), settings||{});
  }

  async function saveLiveMeta(meta){
    if(!db) return false;
    return upsert('dashboardMeta', 'current', meta||{});
  }

  async function testConnection(){
    if(!configured()) return {ok:false,error:'Firebase is not configured. Add the Firebase Web App config first.'};
    try{
      if(!initialized) await init();
      if(!db) return {ok:false,error:'Firebase SDK initialized but Firestore is unavailable.'};
      await collection('dashboardMeta').doc('_connection_test').set({ts:window.firebase.firestore.FieldValue.serverTimestamp()},{merge:true});
      return {ok:true,projectId:currentConfig()?.projectId||''};
    }catch(e){ return {ok:false,error:e.message||String(e)}; }
  }

  async function configureAndPersist(config){
    const cfg = safeConfig(config);
    if(!cfg) throw new Error('Invalid Firebase Web App config. Paste the complete config object from Firebase Console.');
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
    window.SIMPLIAXIS_FIREBASE_CONFIG = cfg;
    // Reloading the SDK is unnecessary; initialize against the newly supplied config.
    initialized = false; db = null; app = null;
    const result = await init();
    return result;
  }

  window.FirebaseDB = {
    init,
    status,
    configured,
    testConnection,
    configureAndPersist,
    saveUsers,
    saveTickets,
    saveUpdates,
    saveNotifications,
    saveTargets,
    saveSettings,
    saveLiveMeta,
    upsert,
    remove,
    read
  };
})();
