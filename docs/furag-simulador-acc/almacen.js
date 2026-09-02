/* Simulador FURAG · capa de almacenamiento
   Un solo contrato para dos modos:
     · local   — todo vive en el navegador (localStorage). Para probar sin backend.
     · remoto  — todo vive en la hoja de cálculo, a través del Web App de Apps Script.
   El modo se decide en config.js: si FURAG_CONFIG.api está vacío, es local. */
(function () {
  'use strict';

  var CFG = window.FURAG_CONFIG || {};
  var API = String(CFG.api || '').trim();
  var CLAVE_STORE = 'FURAG_SIM_ACC_v1';
  var CLAVE_TOKEN = 'FURAG_SIM_ACC_token';
  var CLAVE_SESION = 'FURAG_SIM_ACC_sesion';

  function norm(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
  function ahora() { return new Date().toISOString().slice(0, 16).replace('T', ' '); }
  function hoy() { return new Date().toISOString().slice(0, 10); }

  /* ───────────────── modo remoto ───────────────── */

  function token() { try { return sessionStorage.getItem(CLAVE_TOKEN) || ''; } catch (e) { return ''; } }
  function guardarToken(t) { try { t ? sessionStorage.setItem(CLAVE_TOKEN, t) : sessionStorage.removeItem(CLAVE_TOKEN); } catch (e) {} }

  function llamar(accion, datos) {
    var cuerpo = Object.assign({ accion: accion, token: token() }, datos || {});
    return fetch(API, {
      method: 'POST',
      // text/plain evita la petición previa CORS que Apps Script no responde.
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(cuerpo),
      redirect: 'follow'
    }).then(function (r) {
      if (!r.ok) throw new Error('El servidor respondió ' + r.status + '. Revise la URL del Web App.');
      return r.text();
    }).then(function (t) {
      var d;
      try { d = JSON.parse(t); }
      catch (e) { throw new Error('Respuesta no válida del servidor. ¿La implementación está publicada con acceso «Cualquier persona»?'); }
      if (d.error) {
        var m = d.error;
        if (m === 'SESION_EXPIRADA') { var e2 = new Error('Su sesión expiró. Vuelva a ingresar.'); e2.expirada = true; throw e2; }
        if (m === 'PERMISO') throw new Error('Operación no permitida para este usuario.');
        if (m === 'APROBADA') throw new Error('La pregunta ya fue aprobada: no admite cambios.');
        if (m === 'OBS_OBLIGATORIA') throw new Error('El rechazo exige registrar el motivo.');
        if (m === 'CREDENCIALES') { var e3 = new Error('CREDENCIALES'); e3.credenciales = true; throw e3; }
        throw new Error(m);
      }
      return d;
    });
  }

  var Remoto = {
    remoto: true,
    login: function (u, c) {
      return llamar('login', { usuario: u, clave: c }).then(function (d) {
        guardarToken(d.token);
        return d.sesion;
      }).catch(function (e) {
        if (e.credenciales) return null;
        throw e;
      });
    },
    salir: function () {
      var t = token(); guardarToken('');
      if (!t) return Promise.resolve();
      return llamar('salir', { token: t }).catch(function () {});
    },
    cargar: function () { return llamar('cargar', {}); },
    sondear: function (v) { return llamar('sondear', { version: v }); },
    sembrar: function (catalogo) { return llamar('sembrar', { catalogo: catalogo }); },
    guardar: function (cod, resp, resumen, cat) { return llamar('guardar', { cod: cod, resp: resp, resumen: resumen, cat: cat }); },
    evidencias: function (cod, evis, cat, detalle) { return llamar('evidencias', { cod: cod, evis: evis, cat: cat, detalle: detalle }); },
    enviar: function (cod, cat) { return llamar('enviar', { cod: cod, cat: cat }); },
    revisar: function (cod, estado, obs, cat) { return llamar('revisar', { cod: cod, estado: estado, obs: obs, cat: cat }); },
    limpiar: function () { return llamar('limpiar', {}); }
  };

  /* ───────────────── modo local ───────────────── */

  function leerLocal() {
    try {
      var s = JSON.parse(localStorage.getItem(CLAVE_STORE) || '{}');
      if (!s.reg) s.reg = {};
      if (!s.bitacora) s.bitacora = [];
      return s;
    } catch (e) { return { reg: {}, bitacora: [] }; }
  }
  function escribirLocal(s) {
    try { localStorage.setItem(CLAVE_STORE, JSON.stringify(s)); }
    catch (e) { alert('No se pudo guardar en este navegador: ' + e.message); }
  }
  function regLocal(s, cod) {
    if (!s.reg[cod]) s.reg[cod] = { resp: null, evis: [], estado: 'borrador', obs: '' };
    if (!s.reg[cod].evis) s.reg[cod].evis = [];
    return s.reg[cod];
  }
  var sesionLocal = null;

  var Local = {
    remoto: false,
    login: function (u, c) {
      return new Promise(function (res) {
        var un = norm(u), enc = null;
        window.FURAG.usuarios.forEach(function (x) { if (x.u === un) enc = x; });
        if (!enc || window.sha256FURAG(un + ':' + c + ':' + window.FURAG.salt) !== enc.h) return res(null);
        sesionLocal = { usuario: enc.u, nombre: enc.nombre, rol: enc.rol, dep: enc.dep };
        res(sesionLocal);
      });
    },
    salir: function () { sesionLocal = null; return Promise.resolve(); },
    cargar: function () {
      var s = leerLocal();
      return Promise.resolve({ reg: s.reg, bitacora: s.bitacora, version: 0, servidor: ahora() });
    },
    sondear: function () { return Promise.resolve({ sinCambios: true }); },
    sembrar: function () { return Promise.resolve({ ok: true }); },
    guardar: function (cod, resp) {
      var s = leerLocal(), r = regLocal(s, cod);
      r.resp = resp; r.actualizado = ahora(); r.por = sesionLocal ? sesionLocal.nombre : '';
      escribirLocal(s); return Promise.resolve({ ok: true });
    },
    evidencias: function (cod, evis, cat, detalle) {
      var s = leerLocal(), r = regLocal(s, cod);
      r.evis = evis; r.actualizado = ahora();
      s.bitacora.push({ fecha: ahora(), qid: cod, accion: detalle || 'evidencias_actualizadas',
        por: sesionLocal ? sesionLocal.nombre : '', rol: sesionLocal ? sesionLocal.rol : '', detalle: evis.length + ' evidencia(s)' });
      escribirLocal(s); return Promise.resolve({ ok: true });
    },
    enviar: function (cod) {
      var s = leerLocal(), r = regLocal(s, cod), previo = r.estado;
      r.estado = 'pendiente'; r.envFecha = ahora(); r.envPor = sesionLocal ? sesionLocal.nombre : '';
      s.bitacora.push({ fecha: ahora(), qid: cod, accion: previo === 'rechazado' ? 'subsanacion_enviada' : 'enviada_a_revision',
        por: r.envPor, rol: 'dependencia', detalle: '' });
      escribirLocal(s); return Promise.resolve({ ok: true });
    },
    revisar: function (cod, estado, obs) {
      var s = leerLocal(), r = regLocal(s, cod);
      r.estado = estado; r.obs = estado === 'rechazado' ? String(obs).trim() : '';
      r.revFecha = ahora(); r.revPor = sesionLocal ? sesionLocal.nombre : '';
      s.bitacora.push({ fecha: ahora(), qid: cod, accion: estado === 'aprobado' ? 'aprobada' : 'rechazada',
        por: r.revPor, rol: 'admin', detalle: r.obs });
      escribirLocal(s); return Promise.resolve({ ok: true });
    },
    limpiar: function () {
      escribirLocal({ reg: {}, bitacora: [] });
      return Promise.resolve({ ok: true, mensaje: 'Se borró el avance guardado en este navegador.' });
    },
    /* solo en modo local: respaldo por archivo */
    exportar: function () { return leerLocal(); },
    importar: function (s) { escribirLocal(s); },
    borrar: function () { escribirLocal({ reg: {}, bitacora: [] }); }
  };

  window.Almacen = API ? Remoto : Local;
  window.Almacen.api = API;
  window.Almacen.claveSesion = CLAVE_SESION;
  window.Almacen.sondeoMs = Math.max(10, Number(CFG.sondeoSegundos || 25)) * 1000;
  window.Almacen.hoy = hoy;
  window.Almacen.ahora = ahora;
})();
