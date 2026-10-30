// ════════════════════════════════════════════════════════════════
// Cocinería Sra. Jake · Pedidos de almuerzo — Codigo.gs
// ════════════════════════════════════════════════════════════════
//
// ════════ HOJAS (se crean solas la primera vez) ════════
//
// CONFIG      A:CLAVE  B:VALOR
//   precio_colacion | 5000
//   recargo_hipo    | 1000
//   cierre_empresas | 10:30
//   pin_cocina      | 1234        ← cámbialo después de instalar
//
// MENU        A:ID  B:NOMBRE  C:CUPO  D:LLEVA_AGREGADO  E:ACTIVO  F:ORDEN
//   (el menú de hoy; el cupo es de porciones por día)
//
// AGREGADOS   A:NOMBRE  B:ACTIVO
//
// EMPRESAS    A:ID  B:NOMBRE  C:CODIGO  D:ACTIVA
//   (el código se lo das a la empresa para que pueda reservar a su cuenta)
//
// PEDIDOS     A:ID  B:N  C:FECHA  D:HORA  E:TIPO  F:EMPRESA_ID  G:EMPRESA
//             H:CLIENTE  I:ESTADO  J:TOTAL  K:PAGO  L:CREADO
//
// ITEMS       A:PEDIDO_ID  B:FECHA  C:FONDO_ID  D:FONDO  E:AGREGADO
//             F:HIPOCALORICO  G:CANTIDAD  H:PARA  I:PRECIO_UNIT  J:SUBTOTAL
//
// ════════════════════════════════════════════════════════════════

const TZ = 'America/Santiago';
const SH = { CONFIG:'CONFIG', MENU:'MENU', AGREGADOS:'AGREGADOS', EMPRESAS:'EMPRESAS', PEDIDOS:'PEDIDOS', ITEMS:'ITEMS' };
const HEADERS = {
  CONFIG:    ['CLAVE','VALOR'],
  MENU:      ['ID','NOMBRE','CUPO','LLEVA_AGREGADO','ACTIVO','ORDEN'],
  AGREGADOS: ['NOMBRE','ACTIVO'],
  EMPRESAS:  ['ID','NOMBRE','CODIGO','ACTIVA'],
  PEDIDOS:   ['ID','N','FECHA','HORA','TIPO','EMPRESA_ID','EMPRESA','CLIENTE','ESTADO','TOTAL','PAGO','CREADO'],
  ITEMS:     ['PEDIDO_ID','FECHA','FONDO_ID','FONDO','AGREGADO','HIPOCALORICO','CANTIDAD','PARA','PRECIO_UNIT','SUBTOTAL']
};
const ESTADOS = ['pendiente','listo','entregado','anulado'];

// ─── Entrada web ────────────────────────────────────────────────

function jsonOut(d){
  return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  const p = e && e.parameter || {};
  if (p.action) {
    let payload = {};
    try { payload = JSON.parse(p.payload || '{}'); } catch (err) {}
    return jsonOut(api(p.action, payload));
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Cocinería Sra. Jake · Pedidos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e){
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return jsonOut({ ok:false, error:'Solicitud inválida' }); }
  return jsonOut(api(body.action, body.payload || {}));
}

// Llamado por fetch (GitHub Pages) y por google.script.run (cuando se abre desde Apps Script).
function api(action, payload){
  const publicas = { menu: getMenu, pedir: crearPedido };
  const cocina   = { panel: getPanel, estado: cambiarEstado, guardarMenu: guardarMenu, cuentas: getCuentas, login: function(){ return { ok:true }; } };
  try {
    if (publicas[action]) return publicas[action](payload || {});
    if (cocina[action]) {
      if (String((payload || {}).pin || '') !== String(cfg_().pin_cocina)) return { ok:false, error:'PIN incorrecto', auth:false };
      return cocina[action](payload);
    }
    return { ok:false, error:'Acción desconocida: ' + action };
  } catch (err) {
    console.error(err);
    return { ok:false, error: String(err && err.message || err) };
  }
}

// ─── Utilidades ─────────────────────────────────────────────────

function getSS(){
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('Cocinería Sra. Jake · Datos');
  props.setProperty('SS_ID', ss.getId());
  return ss;
}

function sheet_(name){
  const ss = getSS();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    seed_(name, sh);
  }
  return sh;
}

function seed_(name, sh){
  const rows = {
    CONFIG: [['precio_colacion',5000],['recargo_hipo',1000],['cierre_empresas','10:30'],['pin_cocina','1234']],
    MENU: [
      ['pollo','Pollo asado',40,true,true,1],
      ['cazuela','Cazuela de cerdo',25,false,true,2],
      ['panitas','Panitas de vacuno escabechada',15,true,true,3],
      ['olla','Carne de cerdo a la olla',30,true,true,4]
    ],
    AGREGADOS: [['Arroz',true],['Tallarín',true],['Puré',true]]
  }[name];
  if (rows) sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  if (name === 'CONFIG') sh.getRange('B:B').setNumberFormat('@');
}

function rows_(name){
  const v = sheet_(name).getDataRange().getValues();
  v.shift();
  return v.filter(function(r){ return String(r[0]) !== ''; });
}

function cfg_(){
  const c = {};
  rows_(SH.CONFIG).forEach(function(r){ c[String(r[0]).trim()] = r[1]; });
  return {
    precio: Number(c.precio_colacion) || 5000,
    hipo: Number(c.recargo_hipo) || 0,
    cierre: hhmm_(c.cierre_empresas) || '10:30',
    pin_cocina: String(c.pin_cocina || '1234')
  };
}

function setCfg_(key, value){
  const sh = sheet_(SH.CONFIG);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  sh.appendRow([key, value]);
}

function hhmm_(v){
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  const m = String(v || '').match(/^(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}
function hoy_(){ return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function ahora_(){ return Utilities.formatDate(new Date(), TZ, 'HH:mm'); }
function fecha_(v){ return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '').slice(0, 10); }
function bool_(v){ return v === true || /^(true|si|sí|1|x)$/i.test(String(v).trim()); }
function uid_(){ return Utilities.getUuid().slice(0, 8); }

function fondos_(soloActivos){
  return rows_(SH.MENU).map(function(r){
    return { id:String(r[0]), nombre:String(r[1]), cupo:Number(r[2]) || 0, agregado:bool_(r[3]), activo:bool_(r[4]), orden:Number(r[5]) || 0 };
  }).filter(function(f){ return !soloActivos || f.activo; })
    .sort(function(a, b){ return a.orden - b.orden; });
}
function agregados_(){ return rows_(SH.AGREGADOS).filter(function(r){ return bool_(r[1]); }).map(function(r){ return String(r[0]); }); }
function empresas_(){
  return rows_(SH.EMPRESAS).map(function(r){ return { id:String(r[0]), nombre:String(r[1]), codigo:String(r[2]), activa:bool_(r[3]) }; });
}

// Porciones pedidas hoy por fondo (sin contar pedidos anulados).
function vendidosHoy_(){
  const fecha = hoy_();
  const anulados = {};
  rows_(SH.PEDIDOS).forEach(function(r){ if (fecha_(r[2]) === fecha && r[8] === 'anulado') anulados[r[0]] = true; });
  const v = {};
  rows_(SH.ITEMS).forEach(function(r){
    if (fecha_(r[1]) !== fecha || anulados[r[0]]) return;
    v[r[2]] = (v[r[2]] || 0) + (Number(r[6]) || 0);
  });
  return v;
}

// ─── Público ────────────────────────────────────────────────────

function getMenu(){
  const c = cfg_(), vend = vendidosHoy_();
  return {
    ok: true,
    fecha: hoy_(), hora: ahora_(),
    cfg: { precio:c.precio, hipo:c.hipo, cierre:c.cierre, empresasAbiertas: ahora_() <= c.cierre },
    fondos: fondos_(true).map(function(f){
      return { id:f.id, nombre:f.nombre, agregado:f.agregado, cupo:f.cupo, disponible: Math.max(0, f.cupo - (vend[f.id] || 0)) };
    }),
    agregados: agregados_(),
    empresas: empresas_().filter(function(e){ return e.activa; }).map(function(e){ return { id:e.id, nombre:e.nombre }; })
  };
}

// payload: { tipo:'publico'|'empresa', empresaId, codigo, cliente, items:[{fondoId, agregado, hipo, cant, para}] }
function crearPedido(p){
  const items = Array.isArray(p.items) ? p.items : [];
  if (!items.length) return { ok:false, error:'El pedido no tiene almuerzos.' };
  const cliente = String(p.cliente || '').trim().slice(0, 80);
  if (!cliente) return { ok:false, error:'Falta el nombre de quien pide.' };
  const tipo = p.tipo === 'empresa' ? 'empresa' : 'publico';
  const c = cfg_();

  let empresa = null;
  if (tipo === 'empresa') {
    empresa = empresas_().filter(function(e){ return e.activa && e.id === p.empresaId; })[0];
    if (!empresa) return { ok:false, error:'Empresa no encontrada.' };
    if (String(p.codigo || '').trim().toUpperCase() !== empresa.codigo.trim().toUpperCase()) return { ok:false, error:'El código de la empresa no es correcto.' };
    if (ahora_() > c.cierre) return { ok:false, error:'Las reservas de empresa cerraron a las ' + c.cierre + '. Puede pedir como público general.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const fondos = {}; fondos_(true).forEach(function(f){ fondos[f.id] = f; });
    const ags = agregados_();
    const vend = vendidosHoy_();
    const pide = {};
    const limpios = [];
    for (let i = 0; i < items.length; i++) {
      const it = items[i], f = fondos[it.fondoId];
      const cant = Math.floor(Number(it.cant));
      if (!f) return { ok:false, error:'Un plato del pedido ya no está en el menú.' };
      if (!(cant >= 1 && cant <= 200)) return { ok:false, error:'Cantidad inválida para ' + f.nombre + '.' };
      const hipo = f.agregado && !!it.hipo;   // hipocalórico = sin agregado, solo aplica a platos que lo llevan
      let agregado = '';
      if (f.agregado && !hipo) {
        agregado = String(it.agregado || '');
        if (ags.indexOf(agregado) < 0) return { ok:false, error:'Elige un agregado para ' + f.nombre + '.' };
      }
      pide[f.id] = (pide[f.id] || 0) + cant;
      limpios.push({ f:f, cant:cant, hipo:hipo, agregado:agregado, para:String(it.para || '').trim().slice(0, 60), precio: c.precio + (hipo ? c.hipo : 0) });
    }
    for (const id in pide) {
      const disp = fondos[id].cupo - (vend[id] || 0);
      if (pide[id] > disp) return { ok:false, error: disp > 0 ? 'Solo quedan ' + disp + ' de ' + fondos[id].nombre + '.' : fondos[id].nombre + ' está agotado.', menu:getMenu() };
    }

    const fecha = hoy_(), hora = ahora_(), id = uid_();
    const n = rows_(SH.PEDIDOS).filter(function(r){ return fecha_(r[2]) === fecha; }).length + 1;
    const total = limpios.reduce(function(s, x){ return s + x.precio * x.cant; }, 0);
    sheet_(SH.PEDIDOS).appendRow([id, n, fecha, hora, tipo, empresa ? empresa.id : '', empresa ? empresa.nombre : '', cliente, 'pendiente', total,
      tipo === 'empresa' ? 'cuenta empresa' : 'al retirar', new Date()]);
    const it = limpios.map(function(x){ return [id, fecha, x.f.id, x.f.nombre, x.agregado, x.hipo, x.cant, x.para, x.precio, x.precio * x.cant]; });
    const shI = sheet_(SH.ITEMS);
    shI.getRange(shI.getLastRow() + 1, 1, it.length, it[0].length).setValues(it);
    return { ok:true, pedido:{ id:id, n:n, hora:hora, total:total } };
  } finally {
    lock.releaseLock();
  }
}

// ─── Cocina / administración (requieren PIN) ────────────────────

function pedidosDe_(fecha){
  const items = {};
  rows_(SH.ITEMS).forEach(function(r){
    if (fecha_(r[1]) !== fecha) return;
    (items[r[0]] = items[r[0]] || []).push({ fondoId:String(r[2]), fondo:String(r[3]), agregado:String(r[4] || ''), hipo:bool_(r[5]), cant:Number(r[6]) || 0, para:String(r[7] || ''), precio:Number(r[8]) || 0 });
  });
  return rows_(SH.PEDIDOS).filter(function(r){ return fecha_(r[2]) === fecha; }).map(function(r){
    return { id:String(r[0]), n:Number(r[1]), hora:hhmm_(r[3]), tipo:String(r[4]), empresaId:String(r[5]), empresa:String(r[6]), cliente:String(r[7]), estado:String(r[8]), total:Number(r[9]) || 0, pago:String(r[10]), items: items[r[0]] || [] };
  });
}

function getPanel(p){
  const fecha = /^\d{4}-\d{2}-\d{2}$/.test(p.fecha || '') ? p.fecha : hoy_();
  const c = cfg_();
  return {
    ok: true, fecha: fecha, hora: ahora_(),
    cfg: { precio:c.precio, hipo:c.hipo, cierre:c.cierre },
    fondos: fondos_(false), agregados: agregados_(), empresas: empresas_(),
    pedidos: pedidosDe_(fecha)
  };
}

function cambiarEstado(p){
  if (ESTADOS.indexOf(p.estado) < 0) return { ok:false, error:'Estado inválido' };
  const sh = sheet_(SH.PEDIDOS), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(p.id)) { sh.getRange(i + 1, 9).setValue(p.estado); return { ok:true }; }
  return { ok:false, error:'Pedido no encontrado' };
}

// payload: { cfg:{precio,hipo,cierre,pinNuevo}, fondos:[...], agregados:[...], empresas:[...] }
function guardarMenu(p){
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (p.cfg) {
      if (p.cfg.precio != null) setCfg_('precio_colacion', Math.max(0, Math.round(Number(p.cfg.precio) || 0)));
      if (p.cfg.hipo != null) setCfg_('recargo_hipo', Math.max(0, Math.round(Number(p.cfg.hipo) || 0)));
      if (hhmm_(p.cfg.cierre)) setCfg_('cierre_empresas', hhmm_(p.cfg.cierre));
      if (p.cfg.pinNuevo && String(p.cfg.pinNuevo).trim().length >= 4) setCfg_('pin_cocina', String(p.cfg.pinNuevo).trim());
    }
    if (Array.isArray(p.fondos)) replace_(SH.MENU, p.fondos.map(function(f, i){
      return [f.id || uid_(), String(f.nombre || '').trim(), Math.max(0, Math.floor(Number(f.cupo) || 0)), !!f.agregado, f.activo !== false, i + 1];
    }).filter(function(r){ return r[1]; }));
    if (Array.isArray(p.agregados)) replace_(SH.AGREGADOS, p.agregados.map(function(a){ return [String(a).trim(), true]; }).filter(function(r){ return r[0]; }));
    if (Array.isArray(p.empresas)) replace_(SH.EMPRESAS, p.empresas.map(function(e){
      return [e.id || uid_(), String(e.nombre || '').trim(), String(e.codigo || '').trim().toUpperCase() || codigo_(), e.activa !== false];
    }).filter(function(r){ return r[1]; }));
    return getPanel({});
  } finally {
    lock.releaseLock();
  }
}

function replace_(name, rows){
  const sh = sheet_(name), last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, HEADERS[name].length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, HEADERS[name].length).setValues(rows);
}

function codigo_(){
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = ''; for (let i = 0; i < 5; i++) s += abc.charAt(Math.floor(Math.random() * abc.length));
  return s;
}

// Resumen mensual para cobrar a cada empresa. payload: { mes:'yyyy-MM' }
function getCuentas(p){
  const mes = /^\d{4}-\d{2}$/.test(p.mes || '') ? p.mes : hoy_().slice(0, 7);
  const por = {};
  rows_(SH.PEDIDOS).forEach(function(r){
    if (String(r[4]) !== 'empresa' || String(r[8]) === 'anulado' || fecha_(r[2]).slice(0, 7) !== mes) return;
    const k = String(r[5]);
    por[k] = por[k] || { empresaId:k, empresa:String(r[6]), pedidos:0, total:0, dias:{} };
    por[k].pedidos++; por[k].total += Number(r[9]) || 0; por[k].dias[fecha_(r[2])] = true;
  });
  const colaciones = {};
  const anulados = {};
  rows_(SH.PEDIDOS).forEach(function(r){ if (String(r[8]) === 'anulado') anulados[r[0]] = true; });
  const empPorPedido = {};
  rows_(SH.PEDIDOS).forEach(function(r){ if (String(r[4]) === 'empresa') empPorPedido[r[0]] = String(r[5]); });
  rows_(SH.ITEMS).forEach(function(r){
    const e = empPorPedido[r[0]];
    if (!e || anulados[r[0]] || fecha_(r[1]).slice(0, 7) !== mes) return;
    colaciones[e] = (colaciones[e] || 0) + (Number(r[6]) || 0);
  });
  return { ok:true, mes:mes, cuentas: Object.keys(por).map(function(k){
    return { empresaId:k, empresa:por[k].empresa, dias:Object.keys(por[k].dias).length, colaciones:colaciones[k] || 0, total:por[k].total };
  }).sort(function(a, b){ return b.total - a.total; }) };
}

// Ejecuta esto una vez desde el editor para crear las hojas y autorizar permisos.
function instalar(){
  Object.keys(SH).forEach(function(k){ sheet_(SH[k]); });
  Logger.log('Planilla lista: ' + getSS().getUrl());
}
