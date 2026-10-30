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
//   negocio_nombre, negocio_rut, negocio_telefono, negocio_direccion, negocio_email
//                                 (encabezado de cotizaciones, detalles y balances)
//
// MENU        A:ID  B:NOMBRE  C:CUPO  D:LLEVA_AGREGADO  E:ACTIVO  F:ORDEN
//   (el menú de hoy; el cupo es de porciones por día)
//
// AGREGADOS   A:NOMBRE  B:ACTIVO
//
// EMPRESAS    A:ID  B:NOMBRE  C:CODIGO  D:ACTIVA  E:EMAIL  F:CONTACTO  G:RUT
//   (el código se lo das a la empresa para que entre a su portal y anote a su gente)
//
// PEDIDOS     A:ID  B:N  C:FECHA  D:HORA  E:TIPO  F:EMPRESA_ID  G:EMPRESA
//             H:CLIENTE  I:ESTADO  J:TOTAL  K:PAGO  L:CREADO
//
// ITEMS       A:PEDIDO_ID  B:FECHA  C:FONDO_ID  D:FONDO  E:AGREGADO
//             F:HIPOCALORICO  G:CANTIDAD  H:PARA  I:PRECIO_UNIT  J:SUBTOTAL  K:ITEM_ID
//   (en las listas de empresa, PARA es el nombre del trabajador que se anotó)
//
// COTIZACIONES A:ID  B:N  C:FECHA  D:CLIENTE  E:RUT  F:CONTACTO  G:EMAIL  H:TELEFONO
//             I:LINEAS_JSON  J:DESCUENTO_PCT  K:TOTAL  L:VALIDEZ_DIAS  M:NOTAS
//             N:ESTADO  O:ENVIADA_EL
//
// ════════════════════════════════════════════════════════════════

const TZ = 'America/Santiago';
const SH = { CONFIG:'CONFIG', MENU:'MENU', AGREGADOS:'AGREGADOS', EMPRESAS:'EMPRESAS', PEDIDOS:'PEDIDOS', ITEMS:'ITEMS', COTIZACIONES:'COTIZACIONES' };
const HEADERS = {
  CONFIG:    ['CLAVE','VALOR'],
  MENU:      ['ID','NOMBRE','CUPO','LLEVA_AGREGADO','ACTIVO','ORDEN'],
  AGREGADOS: ['NOMBRE','ACTIVO'],
  EMPRESAS:  ['ID','NOMBRE','CODIGO','ACTIVA','EMAIL','CONTACTO','RUT'],
  PEDIDOS:   ['ID','N','FECHA','HORA','TIPO','EMPRESA_ID','EMPRESA','CLIENTE','ESTADO','TOTAL','PAGO','CREADO'],
  ITEMS:     ['PEDIDO_ID','FECHA','FONDO_ID','FONDO','AGREGADO','HIPOCALORICO','CANTIDAD','PARA','PRECIO_UNIT','SUBTOTAL','ITEM_ID'],
  COTIZACIONES: ['ID','N','FECHA','CLIENTE','RUT','CONTACTO','EMAIL','TELEFONO','LINEAS_JSON','DESCUENTO_PCT','TOTAL','VALIDEZ_DIAS','NOTAS','ESTADO','ENVIADA_EL']
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
  const publicas = { menu: getMenu, pedir: crearPedido, empresaEntrar: empresaEntrar, empresaAnotar: empresaAnotar, empresaQuitar: empresaQuitar, empresaMes: empresaMes };
  const cocina   = { panel: getPanel, estado: cambiarEstado, guardarMenu: guardarMenu, cuentas: getCuentas, login: function(){ return { ok:true }; },
                     balance: getBalance, cotizaciones: listarCotizaciones, guardarCotizacion: guardarCotizacion,
                     documento: getDocumento, pdf: getPdf, enviar: enviarDocumento };
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
  } else if (sh.getLastColumn() < HEADERS[name].length) {
    // planilla creada con una versión anterior: agrega los encabezados nuevos
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
  }
  return sh;
}

function seed_(name, sh){
  const rows = {
    CONFIG: [['precio_colacion',5000],['recargo_hipo',1000],['cierre_empresas','10:30'],['pin_cocina','1234'],
      ['negocio_nombre','Cocinería Sra. Jake'],['negocio_rut',''],['negocio_telefono',''],['negocio_direccion',''],['negocio_email','']],
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
    pin_cocina: String(c.pin_cocina || '1234'),
    negocio: {
      nombre: String(c.negocio_nombre || 'Cocinería Sra. Jake'), rut: String(c.negocio_rut || ''),
      telefono: String(c.negocio_telefono || ''), direccion: String(c.negocio_direccion || ''), email: String(c.negocio_email || '')
    }
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
  return rows_(SH.EMPRESAS).map(function(r){
    return { id:String(r[0]), nombre:String(r[1]), codigo:String(r[2]), activa:bool_(r[3]), email:String(r[4] || ''), contacto:String(r[5] || ''), rut:String(r[6] || '') };
  });
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
    cfg: { precio:c.precio, hipo:c.hipo, cierre:c.cierre },
    fondos: fondos_(true).map(function(f){
      return { id:f.id, nombre:f.nombre, agregado:f.agregado, cupo:f.cupo, disponible: Math.max(0, f.cupo - (vend[f.id] || 0)) };
    }),
    agregados: agregados_()
  };
}

// Valida los platos pedidos contra el menú y el cupo de hoy. Debe llamarse con el lock tomado.
// Devuelve { limpios:[...] } o { error, menu }.
function validarItems_(items, c){
  const fondos = {}; fondos_(true).forEach(function(f){ fondos[f.id] = f; });
  const ags = agregados_();
  const vend = vendidosHoy_();
  const pide = {}, limpios = [];
  for (let i = 0; i < items.length; i++) {
    const it = items[i], f = fondos[it.fondoId];
    const cant = Math.floor(Number(it.cant));
    if (!f) return { error:'Un plato del pedido ya no está en el menú.' };
    if (!(cant >= 1 && cant <= 200)) return { error:'Cantidad inválida para ' + f.nombre + '.' };
    const hipo = f.agregado && !!it.hipo;   // hipocalórico = sin agregado, solo aplica a platos que lo llevan
    let agregado = '';
    if (f.agregado && !hipo) {
      agregado = String(it.agregado || '');
      if (ags.indexOf(agregado) < 0) return { error:'Elige un agregado para ' + f.nombre + '.' };
    }
    pide[f.id] = (pide[f.id] || 0) + cant;
    limpios.push({ f:f, cant:cant, hipo:hipo, agregado:agregado, para:String(it.para || '').trim().slice(0, 60), precio: c.precio + (hipo ? c.hipo : 0) });
  }
  for (const id in pide) {
    const disp = fondos[id].cupo - (vend[id] || 0);
    if (pide[id] > disp) return { error: disp > 0 ? 'Solo quedan ' + disp + ' de ' + fondos[id].nombre + '.' : fondos[id].nombre + ' está agotado.', menu:getMenu() };
  }
  return { limpios:limpios };
}

function agregarItems_(pedidoId, fecha, limpios){
  const it = limpios.map(function(x){ return [pedidoId, fecha, x.f.id, x.f.nombre, x.agregado, x.hipo, x.cant, x.para, x.precio, x.precio * x.cant, uid_()]; });
  const shI = sheet_(SH.ITEMS);
  shI.getRange(shI.getLastRow() + 1, 1, it.length, it[0].length).setValues(it);
  return it.map(function(r){ return r[10]; });
}

function nuevoPedido_(tipo, empresa, cliente, total){
  const fecha = hoy_(), hora = ahora_(), id = uid_();
  const n = rows_(SH.PEDIDOS).filter(function(r){ return fecha_(r[2]) === fecha; }).length + 1;
  sheet_(SH.PEDIDOS).appendRow([id, n, fecha, hora, tipo, empresa ? empresa.id : '', empresa ? empresa.nombre : '', cliente, 'pendiente', total,
    tipo === 'empresa' ? 'cuenta empresa' : 'al retirar', new Date()]);
  return { id:id, n:n, fecha:fecha, hora:hora };
}

// Pedido del público general. payload: { cliente, items:[{fondoId, agregado, hipo, cant, para}] }
function crearPedido(p){
  const items = Array.isArray(p.items) ? p.items : [];
  if (!items.length) return { ok:false, error:'El pedido no tiene almuerzos.' };
  const cliente = String(p.cliente || '').trim().slice(0, 80);
  if (!cliente) return { ok:false, error:'Falta el nombre de quien pide.' };
  const c = cfg_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const v = validarItems_(items, c);
    if (v.error) return { ok:false, error:v.error, menu:v.menu };
    const total = v.limpios.reduce(function(s, x){ return s + x.precio * x.cant; }, 0);
    const ped = nuevoPedido_('publico', null, cliente, total);
    agregarItems_(ped.id, ped.fecha, v.limpios);
    return { ok:true, pedido:{ id:ped.id, n:ped.n, hora:ped.hora, total:total } };
  } finally {
    lock.releaseLock();
  }
}

// ─── Portal de empresas (requieren el código de la empresa) ─────
//
// Cada empresa tiene UNA lista por día (un pedido tipo "empresa").
// Cada trabajador se anota en ella con su nombre y su plato hasta la
// hora de cierre, o hasta que la cocina empiece a preparar la lista.

function empresaPorCodigo_(codigo){
  const c = String(codigo || '').trim().toUpperCase();
  if (!c) return null;
  return empresas_().filter(function(e){ return e.activa && e.codigo.trim().toUpperCase() === c; })[0] || null;
}

function listaHoyRow_(empresaId){
  const sh = sheet_(SH.PEDIDOS), v = sh.getDataRange().getValues(), fecha = hoy_();
  for (let i = v.length - 1; i >= 1; i--) {
    if (String(v[i][4]) === 'empresa' && String(v[i][5]) === empresaId && fecha_(v[i][2]) === fecha && String(v[i][8]) !== 'anulado') return i + 1;
  }
  return 0;
}

function recalcularTotal_(pedidoId){
  let total = 0;
  rows_(SH.ITEMS).forEach(function(r){ if (String(r[0]) === pedidoId) total += Number(r[9]) || 0; });
  const sh = sheet_(SH.PEDIDOS), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === pedidoId) { sh.getRange(i + 1, 10).setValue(total); break; }
  return total;
}

function estadoEmpresa_(empresa){
  const c = cfg_();
  const row = listaHoyRow_(empresa.id);
  const ped = row ? sheet_(SH.PEDIDOS).getRange(row, 1, 1, HEADERS.PEDIDOS.length).getValues()[0] : null;
  const estado = ped ? String(ped[8]) : 'pendiente';
  const abierta = ahora_() <= c.cierre && estado === 'pendiente';
  const lista = ped ? rows_(SH.ITEMS).filter(function(r){ return String(r[0]) === String(ped[0]); }).map(function(r){
    return { itemId:String(r[10] || ''), nombre:String(r[7] || ''), fondoId:String(r[2]), fondo:String(r[3]), agregado:String(r[4] || ''), hipo:bool_(r[5]), cant:Number(r[6]) || 0, precio:Number(r[8]) || 0 };
  }) : [];
  return {
    ok:true,
    empresa:{ id:empresa.id, nombre:empresa.nombre },
    abierta:abierta, estado:estado, cierre:c.cierre,
    motivo: abierta ? '' : (estado !== 'pendiente' ? 'La cocina ya está preparando la lista de hoy.' : 'La lista de hoy cerró a las ' + c.cierre + '.'),
    pedido: ped ? { id:String(ped[0]), n:Number(ped[1]), total:Number(ped[9]) || 0 } : null,
    lista:lista,
    menu:getMenu()
  };
}

// payload: { codigo }
function empresaEntrar(p){
  const e = empresaPorCodigo_(p.codigo);
  if (!e) return { ok:false, error:'No encontramos una empresa con ese código.' };
  return estadoEmpresa_(e);
}

// payload: { codigo, nombre, items:[{fondoId, agregado, hipo, cant}] }
function empresaAnotar(p){
  const e = empresaPorCodigo_(p.codigo);
  if (!e) return { ok:false, error:'No encontramos una empresa con ese código.' };
  const nombre = String(p.nombre || '').trim().slice(0, 60);
  if (!nombre) return { ok:false, error:'Escribe el nombre de quien almuerza.' };
  const items = (Array.isArray(p.items) ? p.items : []).map(function(it){ return { fondoId:it.fondoId, agregado:it.agregado, hipo:it.hipo, cant:it.cant || 1, para:nombre }; });
  if (!items.length) return { ok:false, error:'Elige un plato.' };
  const c = cfg_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const st = estadoEmpresa_(e);
    if (!st.abierta) return Object.assign(st, { ok:false, error:st.motivo });
    const v = validarItems_(items, c);
    if (v.error) return Object.assign(estadoEmpresa_(e), { ok:false, error:v.error });
    const pedidoId = st.pedido ? st.pedido.id : nuevoPedido_('empresa', e, 'Lista de la empresa', 0).id;
    agregarItems_(pedidoId, hoy_(), v.limpios);
    recalcularTotal_(pedidoId);
    return estadoEmpresa_(e);
  } finally {
    lock.releaseLock();
  }
}

// payload: { codigo, itemId }
function empresaQuitar(p){
  const e = empresaPorCodigo_(p.codigo);
  if (!e) return { ok:false, error:'No encontramos una empresa con ese código.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const st = estadoEmpresa_(e);
    if (!st.abierta) return Object.assign(st, { ok:false, error:st.motivo });
    if (!st.pedido) return st;
    const sh = sheet_(SH.ITEMS), v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) {
      if (String(v[i][10]) === String(p.itemId) && String(v[i][0]) === st.pedido.id) { sh.deleteRow(i + 1); break; }
    }
    recalcularTotal_(st.pedido.id);
    return estadoEmpresa_(e);
  } finally {
    lock.releaseLock();
  }
}

// Consumo de la empresa por día en un mes. payload: { codigo, mes:'yyyy-MM' }
function empresaMes(p){
  const e = empresaPorCodigo_(p.codigo);
  if (!e) return { ok:false, error:'No encontramos una empresa con ese código.' };
  const mes = /^\d{4}-\d{2}$/.test(p.mes || '') ? p.mes : hoy_().slice(0, 7);
  const peds = {};
  rows_(SH.PEDIDOS).forEach(function(r){
    if (String(r[4]) === 'empresa' && String(r[5]) === e.id && String(r[8]) !== 'anulado' && fecha_(r[2]).slice(0, 7) === mes) peds[r[0]] = fecha_(r[2]);
  });
  const dias = {};
  rows_(SH.ITEMS).forEach(function(r){
    const f = peds[r[0]]; if (!f) return;
    dias[f] = dias[f] || { fecha:f, colaciones:0, total:0 };
    dias[f].colaciones += Number(r[6]) || 0; dias[f].total += Number(r[9]) || 0;
  });
  return { ok:true, mes:mes, dias: Object.keys(dias).sort().map(function(k){ return dias[k]; }) };
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
    cfg: { precio:c.precio, hipo:c.hipo, cierre:c.cierre, negocio:c.negocio },
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
      if (p.cfg.negocio) ['nombre','rut','telefono','direccion','email'].forEach(function(k){
        if (p.cfg.negocio[k] != null) setCfg_('negocio_' + k, String(p.cfg.negocio[k]).trim().slice(0, 120));
      });
    }
    if (Array.isArray(p.fondos)) replace_(SH.MENU, p.fondos.map(function(f, i){
      return [f.id || uid_(), String(f.nombre || '').trim(), Math.max(0, Math.floor(Number(f.cupo) || 0)), !!f.agregado, f.activo !== false, i + 1];
    }).filter(function(r){ return r[1]; }));
    if (Array.isArray(p.agregados)) replace_(SH.AGREGADOS, p.agregados.map(function(a){ return [String(a).trim(), true]; }).filter(function(r){ return r[0]; }));
    if (Array.isArray(p.empresas)) replace_(SH.EMPRESAS, p.empresas.map(function(e){
      return [e.id || uid_(), String(e.nombre || '').trim(), String(e.codigo || '').trim().toUpperCase() || codigo_(), e.activa !== false,
        String(e.email || '').trim(), String(e.contacto || '').trim(), String(e.rut || '').trim()];
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

// ─── Balances, detalles y cotizaciones (requieren PIN) ──────────

function rango_(p){
  const re = /^\d{4}-\d{2}-\d{2}$/;
  const hasta = re.test(p.hasta || '') ? p.hasta : hoy_();
  const desde = re.test(p.desde || '') ? p.desde : hasta.slice(0, 8) + '01';
  if (desde > hasta) throw new Error('La fecha "desde" es posterior a "hasta".');
  return { desde:desde, hasta:hasta };
}

// Pedidos no anulados entre dos fechas, con sus platos.
function pedidosRango_(desde, hasta, empresaId){
  const peds = {}, lista = [];
  rows_(SH.PEDIDOS).forEach(function(r){
    const f = fecha_(r[2]);
    if (f < desde || f > hasta || String(r[8]) === 'anulado') return;
    if (empresaId && String(r[5]) !== empresaId) return;
    const o = { id:String(r[0]), n:Number(r[1]), fecha:f, tipo:String(r[4]), empresaId:String(r[5]), empresa:String(r[6]), cliente:String(r[7]), estado:String(r[8]), total:Number(r[9]) || 0, items:[] };
    peds[o.id] = o; lista.push(o);
  });
  rows_(SH.ITEMS).forEach(function(r){
    const o = peds[r[0]]; if (!o) return;
    o.items.push({ fondo:String(r[3]), agregado:String(r[4] || ''), hipo:bool_(r[5]), cant:Number(r[6]) || 0, para:String(r[7] || ''), precio:Number(r[8]) || 0, subtotal:Number(r[9]) || 0 });
  });
  return lista.sort(function(a, b){ return a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : a.n - b.n; });
}

function balance_(desde, hasta){
  const peds = pedidosRango_(desde, hasta);
  const acc = function(){ return { pedidos:0, colaciones:0, total:0 }; };
  const R = { total:acc(), publico:acc(), cobrado:acc(), porCobrar:acc(), empresas:acc(), hipo:0 };
  const dias = {}, platos = {}, ags = {}, emps = {};
  peds.forEach(function(o){
    const n = o.items.reduce(function(s, i){ return s + i.cant; }, 0);
    const add = function(a){ a.pedidos++; a.colaciones += n; a.total += o.total; };
    add(R.total);
    if (o.tipo === 'empresa') {
      add(R.empresas);
      const e = emps[o.empresaId] = emps[o.empresaId] || { empresaId:o.empresaId, empresa:o.empresa, dias:{}, colaciones:0, total:0 };
      e.dias[o.fecha] = true; e.colaciones += n; e.total += o.total;
    } else {
      add(R.publico); add(o.estado === 'entregado' ? R.cobrado : R.porCobrar);
    }
    const d = dias[o.fecha] = dias[o.fecha] || { fecha:o.fecha, publico:0, empresas:0, colaciones:0, total:0 };
    d[o.tipo === 'empresa' ? 'empresas' : 'publico'] += n; d.colaciones += n; d.total += o.total;
    o.items.forEach(function(i){
      const pl = platos[i.fondo] = platos[i.fondo] || { fondo:i.fondo, colaciones:0, hipo:0, total:0 };
      pl.colaciones += i.cant; pl.total += i.subtotal; if (i.hipo) { pl.hipo += i.cant; R.hipo += i.cant; }
      const k = i.hipo ? 'Hipocalórico' : (i.agregado || 'Sin agregado');
      ags[k] = (ags[k] || 0) + i.cant;
    });
  });
  const vals = function(o){ return Object.keys(o).map(function(k){ return o[k]; }); };
  return {
    desde:desde, hasta:hasta, resumen:R,
    porDia: vals(dias).sort(function(a, b){ return a.fecha < b.fecha ? -1 : 1; }),
    porPlato: vals(platos).sort(function(a, b){ return b.colaciones - a.colaciones; }),
    porAgregado: Object.keys(ags).map(function(k){ return { agregado:k, colaciones:ags[k] }; }).sort(function(a, b){ return b.colaciones - a.colaciones; }),
    porEmpresa: vals(emps).map(function(e){ return { empresaId:e.empresaId, empresa:e.empresa, dias:Object.keys(e.dias).length, colaciones:e.colaciones, total:e.total }; })
      .sort(function(a, b){ return b.total - a.total; })
  };
}

// payload: { desde, hasta }
function getBalance(p){
  const r = rango_(p);
  return Object.assign({ ok:true }, balance_(r.desde, r.hasta));
}

function cotizaciones_(){
  return rows_(SH.COTIZACIONES).map(function(r){
    let lineas = []; try { lineas = JSON.parse(r[8] || '[]'); } catch (e) {}
    return { id:String(r[0]), n:Number(r[1]), fecha:fecha_(r[2]), cliente:String(r[3]), rut:String(r[4] || ''), contacto:String(r[5] || ''), email:String(r[6] || ''),
      telefono:String(r[7] || ''), lineas:lineas, descuento:Number(r[9]) || 0, total:Number(r[10]) || 0, validez:Number(r[11]) || 0, notas:String(r[12] || ''),
      estado:String(r[13] || 'borrador'), enviadaEl:r[14] ? String(r[14]) : '' };
  });
}

function listarCotizaciones(){
  return { ok:true, cotizaciones: cotizaciones_().sort(function(a, b){ return b.n - a.n; }).slice(0, 100) };
}

function totalCotizacion_(lineas, descuento){
  const sub = lineas.reduce(function(s, l){ return s + l.cantidad * l.precio; }, 0);
  const desc = Math.round(sub * descuento / 100);
  return { subtotal:sub, descuento:desc, total:sub - desc };
}

// payload: { id?, cliente, rut, contacto, email, telefono, lineas:[{descripcion,cantidad,precio}], descuento, validez, notas }
function guardarCotizacion(p){
  const cliente = String(p.cliente || '').trim().slice(0, 100);
  if (!cliente) return { ok:false, error:'Falta el nombre del cliente.' };
  const lineas = (Array.isArray(p.lineas) ? p.lineas : []).map(function(l){
    return { descripcion:String(l.descripcion || '').trim().slice(0, 160), cantidad:Math.max(0, Number(l.cantidad) || 0), precio:Math.max(0, Math.round(Number(l.precio) || 0)) };
  }).filter(function(l){ return l.descripcion && l.cantidad > 0; });
  if (!lineas.length) return { ok:false, error:'Agrega al menos una línea con descripción y cantidad.' };
  const descuento = Math.min(100, Math.max(0, Number(p.descuento) || 0));
  const validez = Math.max(1, Math.floor(Number(p.validez) || 15));
  const total = totalCotizacion_(lineas, descuento).total;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(SH.COTIZACIONES), v = sh.getDataRange().getValues();
    const fila = function(id, n, fecha, estado, enviada){
      return [id, n, fecha, cliente, String(p.rut || '').trim(), String(p.contacto || '').trim(), String(p.email || '').trim(), String(p.telefono || '').trim(),
        JSON.stringify(lineas), descuento, total, validez, String(p.notas || '').trim().slice(0, 1000), estado, enviada];
    };
    for (let i = 1; i < v.length; i++) if (p.id && String(v[i][0]) === String(p.id)) {
      sh.getRange(i + 1, 1, 1, HEADERS.COTIZACIONES.length).setValues([fila(v[i][0], v[i][1], hoy_(), v[i][13] || 'borrador', v[i][14] || '')]);
      return { ok:true, cotizacion: cotizaciones_().filter(function(c){ return c.id === String(p.id); })[0] };
    }
    const n = v.slice(1).reduce(function(m, r){ return Math.max(m, Number(r[1]) || 0); }, 0) + 1;
    const id = uid_();
    sh.appendRow(fila(id, n, hoy_(), 'borrador', ''));
    return { ok:true, cotizacion: cotizaciones_().filter(function(c){ return c.id === id; })[0] };
  } finally {
    lock.releaseLock();
  }
}

// ── Documentos: el mismo HTML sirve para vista previa, PDF y correo ──

function clp_(n){ return '$' + String(Math.round(n || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function ddmm_(iso){ const p = String(iso).split('-'); return p.length === 3 ? p[2] + '-' + p[1] + '-' + p[0] : String(iso); }
function h_(s){ return String(s == null ? '' : s).replace(/[&<>"']/g, function(c){ return { '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]; }); }
function sumarDias_(iso, n){ const d = new Date(iso + 'T12:00:00'); d.setDate(d.getDate() + n); return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }

function docShell_(neg, titulo, meta, cuerpo){
  const datos = [neg.rut && 'RUT ' + h_(neg.rut), h_(neg.direccion), h_(neg.telefono), h_(neg.email)].filter(String).join(' · ');
  return '<div style="font-family:Arial,Helvetica,sans-serif;color:#1f1c19;max-width:720px;margin:0 auto;font-size:13px;line-height:1.45">' +
    '<table style="width:100%;border-collapse:collapse;background:#24211d;color:#f4efe6"><tr>' +
    '<td style="padding:16px 18px"><div style="font-size:11px;letter-spacing:2px;text-transform:uppercase;opacity:.75">Cocinería</div>' +
    '<div style="font-size:24px;font-weight:bold;color:#e9a416">' + h_(neg.nombre) + '</div>' +
    (datos ? '<div style="font-size:11px;opacity:.8;margin-top:2px">' + datos + '</div>' : '') + '</td>' +
    '<td style="padding:16px 18px;text-align:right;vertical-align:top"><div style="font-size:18px;font-weight:bold">' + h_(titulo) + '</div>' +
    meta.map(function(m){ return '<div style="font-size:12px;opacity:.85">' + m + '</div>'; }).join('') + '</td></tr></table>' +
    '<div style="padding:16px 4px">' + cuerpo + '</div></div>';
}
const TD_ = 'padding:6px 8px;border-bottom:1px solid #e5e1d8;';
const TH_ = 'padding:6px 8px;border-bottom:2px solid #1f1c19;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#6a645c;';
function tabla_(cols, filas, pie){
  const al = function(c){ return c.n ? 'text-align:right;' : ''; };
  return '<table style="width:100%;border-collapse:collapse;margin:8px 0 18px">' +
    '<tr>' + cols.map(function(c){ return '<th style="' + TH_ + al(c) + '">' + h_(c.t) + '</th>'; }).join('') + '</tr>' +
    filas.map(function(f){ return '<tr>' + f.map(function(v, i){ return '<td style="' + TD_ + al(cols[i]) + '">' + v + '</td>'; }).join('') + '</tr>'; }).join('') +
    (pie ? '<tr>' + pie.map(function(v, i){ return '<td style="padding:8px;font-weight:bold;' + al(cols[i]) + '">' + v + '</td>'; }).join('') + '</tr>' : '') +
    '</table>';
}
function h2_(t){ return '<h2 style="font-size:15px;margin:18px 0 4px">' + h_(t) + '</h2>'; }

// payload: { tipo:'balance'|'detalle'|'cotizacion', desde, hasta, empresaId, id }
function documento_(p){
  const c = cfg_(), neg = c.negocio, hoy = hoy_();
  if (p.tipo === 'balance') {
    const r = rango_(p), b = balance_(r.desde, r.hasta), R = b.resumen;
    const cuerpo =
      tabla_([{ t:'Concepto' }, { t:'Pedidos', n:1 }, { t:'Colaciones', n:1 }, { t:'Monto', n:1 }], [
        ['Público general · cobrado', R.cobrado.pedidos, R.cobrado.colaciones, clp_(R.cobrado.total)],
        ['Público general · por cobrar', R.porCobrar.pedidos, R.porCobrar.colaciones, clp_(R.porCobrar.total)],
        ['Empresas · a cuenta', R.empresas.pedidos, R.empresas.colaciones, clp_(R.empresas.total)]
      ], ['Total del período', R.total.pedidos, R.total.colaciones, clp_(R.total.total)]) +
      h2_('Por empresa') + (b.porEmpresa.length ? tabla_([{ t:'Empresa' }, { t:'Días', n:1 }, { t:'Colaciones', n:1 }, { t:'Monto', n:1 }],
        b.porEmpresa.map(function(e){ return [h_(e.empresa), e.dias, e.colaciones, clp_(e.total)]; })) : '<p>Sin consumo de empresas.</p>') +
      h2_('Por plato') + tabla_([{ t:'Plato' }, { t:'Colaciones', n:1 }, { t:'Hipocalóricas', n:1 }, { t:'Monto', n:1 }],
        b.porPlato.map(function(x){ return [h_(x.fondo), x.colaciones, x.hipo, clp_(x.total)]; })) +
      h2_('Por día') + tabla_([{ t:'Fecha' }, { t:'Público', n:1 }, { t:'Empresas', n:1 }, { t:'Monto', n:1 }],
        b.porDia.map(function(d){ return [ddmm_(d.fecha), d.publico, d.empresas, clp_(d.total)]; }));
    const per = ddmm_(r.desde) + ' al ' + ddmm_(r.hasta);
    return { titulo:'Balance', asunto:'Balance ' + per + ' · ' + neg.nombre, archivo:'Balance ' + r.desde + ' a ' + r.hasta,
      html: docShell_(neg, 'Balance', ['Período: ' + per, 'Emitido: ' + ddmm_(hoy)], cuerpo),
      texto: neg.nombre + ' — Balance ' + per + ': ' + R.total.colaciones + ' colaciones, total ' + clp_(R.total.total) +
        ' (público ' + clp_(R.publico.total) + ', empresas ' + clp_(R.empresas.total) + ').', para:'', telefono:'' };
  }
  if (p.tipo === 'detalle') {
    const r = rango_(p);
    const e = empresas_().filter(function(x){ return x.id === p.empresaId; })[0];
    if (!e) throw new Error('Empresa no encontrada.');
    const peds = pedidosRango_(r.desde, r.hasta, e.id).filter(function(o){ return o.tipo === 'empresa'; });
    let n = 0, total = 0;
    const filas = [];
    peds.forEach(function(o){
      o.items.forEach(function(i){
        n += i.cant; total += i.subtotal;
        filas.push([ddmm_(o.fecha), h_(i.para || '—'), h_(i.fondo) + (i.hipo ? ' · Hipocalórico' : i.agregado ? ' · ' + h_(i.agregado) : ''), i.cant, clp_(i.subtotal)]);
      });
    });
    const dias = {}; peds.forEach(function(o){ dias[o.fecha] = (dias[o.fecha] || 0) + o.total; });
    const per = ddmm_(r.desde) + ' al ' + ddmm_(r.hasta);
    const cuerpo =
      '<table style="width:100%;margin-bottom:12px"><tr><td style="vertical-align:top"><div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#6a645c">Cliente</div>' +
      '<div style="font-size:16px;font-weight:bold">' + h_(e.nombre) + '</div>' + (e.rut ? '<div>RUT ' + h_(e.rut) + '</div>' : '') + (e.contacto ? '<div>' + h_(e.contacto) + '</div>' : '') + '</td>' +
      '<td style="text-align:right;vertical-align:top"><div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#6a645c">Total del período</div>' +
      '<div style="font-size:24px;font-weight:bold">' + clp_(total) + '</div><div>' + n + ' colaciones · ' + Object.keys(dias).length + ' días</div></td></tr></table>' +
      h2_('Resumen por día') + tabla_([{ t:'Fecha' }, { t:'Monto', n:1 }], Object.keys(dias).sort().map(function(d){ return [ddmm_(d), clp_(dias[d])]; }), ['Total', clp_(total)]) +
      h2_('Detalle de colaciones') + (filas.length ? tabla_([{ t:'Fecha' }, { t:'Persona' }, { t:'Plato' }, { t:'Cant.', n:1 }, { t:'Monto', n:1 }], filas, ['', '', 'Total', n, clp_(total)]) : '<p>Sin consumo en el período.</p>') +
      '<p style="color:#6a645c;font-size:12px">Cada colación incluye sopa, ensalada, pan y vaso de jugo.</p>';
    return { titulo:'Detalle de consumo', asunto:'Detalle de consumo ' + per + ' · ' + e.nombre, archivo:'Detalle ' + e.nombre + ' ' + r.desde + ' a ' + r.hasta,
      html: docShell_(neg, 'Detalle de consumo', ['Período: ' + per, 'Emitido: ' + ddmm_(hoy)], cuerpo),
      texto: 'Hola, le enviamos el detalle de consumo de ' + e.nombre + ' del ' + per + ': ' + n + ' colaciones, total ' + clp_(total) + '. ' + neg.nombre + '.',
      para:e.email, telefono:'' };
  }
  if (p.tipo === 'cotizacion') {
    const q = cotizaciones_().filter(function(x){ return x.id === String(p.id); })[0];
    if (!q) throw new Error('Cotización no encontrada.');
    const t = totalCotizacion_(q.lineas, q.descuento);
    const filas = q.lineas.map(function(l){ return [h_(l.descripcion), l.cantidad, clp_(l.precio), clp_(l.cantidad * l.precio)]; });
    const totales = '<table style="margin-left:auto;border-collapse:collapse;min-width:260px">' +
      '<tr><td style="padding:4px 8px">Subtotal</td><td style="padding:4px 8px;text-align:right">' + clp_(t.subtotal) + '</td></tr>' +
      (q.descuento ? '<tr><td style="padding:4px 8px">Descuento ' + q.descuento + '%</td><td style="padding:4px 8px;text-align:right">−' + clp_(t.descuento) + '</td></tr>' : '') +
      '<tr><td style="padding:8px;font-weight:bold;font-size:16px;border-top:2px solid #1f1c19">Total</td><td style="padding:8px;text-align:right;font-weight:bold;font-size:16px;border-top:2px solid #1f1c19">' + clp_(t.total) + '</td></tr></table>';
    const cuerpo =
      '<div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#6a645c">Para</div>' +
      '<div style="font-size:16px;font-weight:bold">' + h_(q.cliente) + '</div>' +
      [q.rut && 'RUT ' + h_(q.rut), h_(q.contacto), h_(q.email), h_(q.telefono)].filter(String).map(function(x){ return '<div>' + x + '</div>'; }).join('') +
      tabla_([{ t:'Descripción' }, { t:'Cantidad', n:1 }, { t:'Precio unitario', n:1 }, { t:'Total', n:1 }], filas) + totales +
      (q.notas ? h2_('Condiciones') + '<p style="white-space:pre-wrap">' + h_(q.notas) + '</p>' : '') +
      '<p style="color:#6a645c;font-size:12px;margin-top:18px">Cada colación incluye plato de fondo con agregado, sopa, ensalada, pan y vaso de jugo. ' +
      'Cotización válida hasta el ' + ddmm_(sumarDias_(q.fecha, q.validez)) + '.</p>';
    return { titulo:'Cotización N° ' + q.n, asunto:'Cotización N° ' + q.n + ' · ' + neg.nombre, archivo:'Cotizacion ' + q.n + ' ' + q.cliente,
      html: docShell_(neg, 'Cotización', ['N° ' + q.n, 'Fecha: ' + ddmm_(q.fecha), 'Válida hasta: ' + ddmm_(sumarDias_(q.fecha, q.validez))], cuerpo),
      texto: 'Hola ' + q.cliente + ', le enviamos la cotización N° ' + q.n + ' de ' + neg.nombre + ': ' +
        q.lineas.map(function(l){ return l.cantidad + ' × ' + l.descripcion + ' (' + clp_(l.precio) + ')'; }).join('; ') +
        '. Total ' + clp_(t.total) + (q.descuento ? ' con ' + q.descuento + '% de descuento' : '') + '. Válida hasta el ' + ddmm_(sumarDias_(q.fecha, q.validez)) + '.',
      para:q.email, telefono:q.telefono };
  }
  throw new Error('Tipo de documento desconocido.');
}

function getDocumento(p){ return Object.assign({ ok:true }, documento_(p)); }

function pdfBlob_(d){
  const page = '<html><head><meta charset="utf-8"></head><body style="margin:24px">' + d.html + '</body></html>';
  return Utilities.newBlob(page, 'text/html', 'doc.html').getAs('application/pdf').setName(d.archivo.replace(/[\\/:*?"<>|]/g, '') + '.pdf');
}

function getPdf(p){
  const d = documento_(p), b = pdfBlob_(d);
  return { ok:true, nombre:b.getName(), base64:Utilities.base64Encode(b.getBytes()) };
}

// payload: { ...documento, para, mensaje }
function enviarDocumento(p){
  const para = String(p.para || '').trim();
  if (!/^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$/.test(para)) return { ok:false, error:'Escribe un correo válido.' };
  const d = documento_(p), neg = cfg_().negocio;
  const msg = String(p.mensaje || '').trim();
  const intro = msg ? '<p style="font-family:Arial,Helvetica,sans-serif;font-size:14px;white-space:pre-wrap">' + h_(msg) + '</p>' : '';
  const opts = { to:para, subject:d.asunto, htmlBody: intro + d.html, name:neg.nombre, attachments:[pdfBlob_(d)] };
  if (neg.email) opts.replyTo = neg.email;
  MailApp.sendEmail(opts);
  if (p.tipo === 'cotizacion') {
    const sh = sheet_(SH.COTIZACIONES), v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(p.id)) { sh.getRange(i + 1, 14, 1, 2).setValues([['enviada', new Date()]]); break; }
  }
  return { ok:true, enviadoA:para };
}

// Ejecuta esto una vez desde el editor para crear las hojas y autorizar permisos.
function instalar(){
  Object.keys(SH).forEach(function(k){ sheet_(SH[k]); });
  Logger.log('Planilla lista: ' + getSS().getUrl());
}
