/* Simulador FURAG · Agencia Catastral de Cundinamarca
   Interfaz. Todo el acceso a datos pasa por window.Almacen (almacen.js), que decide
   si el avance vive en este navegador o en la hoja de cálculo compartida. */
(function () {
  'use strict';

  var D = window.FURAG;
  if (!D) { document.body.innerHTML = '<p style="padding:40px">No se pudo cargar datos.js</p>'; return; }
  var A = window.Almacen;

  var $ = function (s, r) { return (r || document).querySelector(s); };

  /* ═════════════════════ utilidades ═════════════════════ */
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function norm(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
  function hoy() { return new Date().toISOString().slice(0, 10); }
  function ahora() { return new Date().toISOString().slice(0, 16).replace('T', ' '); }
  function fmt(n) { return Number(n || 0).toLocaleString('es-CO'); }
  function pct(a, b) { return b ? Math.round(a * 1000 / b) / 10 : 0; }

  var POL = {}; D.politicas.forEach(function (p) { POL[p.id] = p; });
  var DEP = {}; D.dependencias.forEach(function (d) { DEP[d.id] = d; });
  /* Nombres visibles de las dependencias. Es solo texto: los identificadores
  DEP-01 … DEP-11 no cambian, así que no afecta la asignación de preguntas,
  ni el login, ni lo que está guardado en la hoja. */
  var NOMBRE_DEP = {
    'DEP-GT': 'Gerencia Técnica (Administrador)',
    'DEP-01': 'Gerencia General',
    'DEP-02': 'Planeación',
    'DEP-03': 'Subgerencia Administrativa y Financiera',
    'DEP-04': 'Talento Humano',
    'DEP-05': 'Subgerencia de Cartografía, Tecnología e Innovación',
    'DEP-06': 'Jurídica',
    'DEP-07': 'Contratación',
    'DEP-08': 'Atención al Ciudadano',
    'DEP-09': 'Gestión Documental',
    'DEP-10': 'Control Interno',
    'DEP-11': 'Comunicaciones'
  };
  D.dependencias.forEach(function (d) { if (NOMBRE_DEP[d.id]) d.nombre = NOMBRE_DEP[d.id]; });
  var Q = {}; D.preguntas.forEach(function (p) { Q[p.id] = p; });

  var ETIQ_TIPO = {
    unica: 'Selección única', multiple: 'Selección múltiple',
    multiple_num: 'Selección múltiple numérica', numerica: 'Abierta numérica',
    texto: 'Abierta / Texto', matricial: 'Matricial'
  };
  var ETIQ_ESTADO = {
    borrador: 'Sin enviar', pendiente: 'Pendiente de revisión',
    aprobado: 'Aprobado', rechazado: 'Rechazado'
  };

  /* Ficha que acompaña cada escritura: le permite al servidor crear la fila
     con su política y dependencia si todavía no existe. */
  function cat(p) {
    return { pol: p.polx, dep: p.dep, tipo: p.tipoOriginal, ev: p.evidencia ? 'Sí' : 'No' };
  }

  /* ═════════════════════ estado en memoria ═════════════════════ */
  var store = { reg: {}, bitacora: [] };
  var versionServidor = 0;
  var sesion = null;
  var temporizador = null;

  function reg(qid) {
    if (!store.reg[qid]) store.reg[qid] = { resp: null, evis: [], estado: 'borrador', obs: '' };
    if (!store.reg[qid].evis) store.reg[qid].evis = [];
    return store.reg[qid];
  }

  try { sesion = JSON.parse(sessionStorage.getItem(A.claveSesion) || 'null'); } catch (e) { sesion = null; }

  /* ═════════════════════ permisos (interfaz) ═══════════════════════
     En modo compartido el servidor vuelve a validar cada operación y solo
     entrega a cada dependencia sus propias preguntas; esto es la primera capa. */
  function esAdmin() { return !!sesion && sesion.rol === 'admin'; }
  function puedeVer(q) { return !!sesion && (esAdmin() || q.dep === sesion.dep); }
  function preguntasVisibles() { return D.preguntas.filter(puedeVer); }
  function puedeEditar(q) {
    if (!sesion || esAdmin()) return false;
    if (q.dep !== sesion.dep) return false;
    return reg(q.id).estado !== 'aprobado';
  }

  /* ═════════════════════ indicador de sincronización ═════════════════════ */
  var elSync = null;
  function marcarSync(estado, detalle) {
    if (!elSync) return;
    var t = { ok: '● Sincronizado', trabajando: '◌ Guardando…', error: '▲ Sin conexión' }[estado] || '';
    var c = { ok: 'var(--st-bien)', trabajando: 'var(--tenue)', error: 'var(--st-critica)' }[estado];
    elSync.textContent = t;
    elSync.style.color = c;
    elSync.title = detalle || (estado === 'ok' ? 'Último guardado ' + ahora() : '');
  }

  var pendientes = 0;
  function sincronizar(promesa) {
    if (!A.remoto) return promesa;
    pendientes++; marcarSync('trabajando');
    return promesa.then(function (d) {
      if (d && d.version) versionServidor = d.version;
      if (--pendientes <= 0) { pendientes = 0; marcarSync('ok'); }
      return d;
    }).catch(function (e) {
      pendientes = Math.max(0, pendientes - 1);
      marcarSync('error', e.message);
      if (e.expirada) { alert(e.message); salir(); return; }
      alert('No se pudo guardar en el servidor: ' + e.message +
        '\n\nSe recargará el avance real para no dejar datos inconsistentes.');
      return recargar();
    });
  }

  function recargar() {
    return A.cargar().then(function (d) {
      store.reg = d.reg || {}; store.bitacora = d.bitacora || [];
      versionServidor = d.version || 0;
      marcarSync('ok');
      pintarContenido();
    }).catch(function (e) { marcarSync('error', e.message); });
  }

  function arrancarSondeo() {
    if (!A.remoto) return;
    if (temporizador) clearInterval(temporizador);
    temporizador = setInterval(function () {
      if (pendientes) return;
      A.sondear(versionServidor).then(function (d) {
        if (d.sinCambios) { marcarSync('ok'); return; }
        store.reg = d.reg || {}; store.bitacora = d.bitacora || [];
        versionServidor = d.version || 0;
        marcarSync('ok');
        pintarContenido();
      }).catch(function (e) {
        if (e.expirada) { clearInterval(temporizador); alert(e.message); salir(); return; }
        marcarSync('error', e.message);
      });
    }, A.sondeoMs);
  }

  /* ═════════════════════ mutaciones ═════════════════════ */
  function guardarRespuesta(qid, valor) {
    var q = Q[qid];
    if (!q || !puedeEditar(q)) throw new Error('Operación no permitida para este usuario.');
    var r = reg(qid);
    r.resp = valor; r.actualizado = ahora(); r.por = sesion.nombre;
    return sincronizar(A.guardar(qid, valor, resumenRespuesta(q, valor), cat(q)));
  }
  /* Espera antes de mandar un guardado al servidor. Agrupa varios clics o
     teclas en una sola escritura: con 20 dependencias trabajando a la vez, es
     lo que evita que la cola de escrituras se desborde. */
  var MS_GUARDADO = 2000;
  var debounce = {};
  function guardarRespuestaDiferido(qid, valor) {
    var q = Q[qid];
    if (!q || !puedeEditar(q)) return;
    var r = reg(qid);
    r.resp = valor; r.actualizado = ahora();
    clearTimeout(debounce[qid]);
    debounce[qid] = setTimeout(function () {
      sincronizar(A.guardar(qid, valor, resumenRespuesta(q, valor), cat(q)));
    }, MS_GUARDADO);
  }
  function agregarEvidencia(qid, nombre, url) {
    var q = Q[qid];
    if (!q || !puedeEditar(q)) throw new Error('Operación no permitida para este usuario.');
    var r = reg(qid);
    r.evis.push({ id: 'EVI-' + Date.now().toString(36), nombre: nombre, url: url,
      fecha: hoy(), por: sesion.nombre });
    return sincronizar(A.evidencias(qid, r.evis, cat(q), 'evidencia_agregada'));
  }
  function quitarEvidencia(qid, eid) {
    var q = Q[qid];
    if (!q || !puedeEditar(q)) throw new Error('Operación no permitida para este usuario.');
    var r = reg(qid);
    r.evis = r.evis.filter(function (e) { return e.id !== eid; });
    return sincronizar(A.evidencias(qid, r.evis, cat(q), 'evidencia_eliminada'));
  }
  function enviarARevision(qid) {
    var q = Q[qid];
    if (!q || !puedeEditar(q)) throw new Error('Operación no permitida para este usuario.');
    var r = reg(qid);
    r.estado = 'pendiente'; r.envFecha = ahora(); r.envPor = sesion.nombre;
    return sincronizar(A.enviar(qid, cat(q)));
  }
  function revisar(qid, estado, obs) {
    if (!esAdmin()) throw new Error('Solo la Gerencia Técnica puede aprobar o rechazar.');
    if (estado === 'rechazado' && !String(obs || '').trim())
      throw new Error('El rechazo exige registrar el motivo.');
    var r = reg(qid);
    r.estado = estado;
    r.obs = estado === 'rechazado' ? String(obs).trim() : '';
    r.revFecha = ahora(); r.revPor = sesion.nombre;
    return sincronizar(A.revisar(qid, estado, r.obs, cat(Q[qid])));
  }

  /* ═════════════════════ login ═════════════════════ */
  var formLogin = $('#formLogin'), errLogin = $('#errLogin'), btnLogin = $('#formLogin button[type=submit]');

  /* La lista de dependencias se arma con las siglas del catálogo de usuarios.
     El valor enviado es la sigla; el servidor la resuelve contra la pestaña USUARIOS. */
  (function llenarDependencias() {
    var sel = $('#usuario');
    if (!sel || sel.tagName !== 'SELECT') return;
    (D.usuarios || []).forEach(function (u) {
      var o = document.createElement('option');
      o.value = u.sigla || u.u;
      o.textContent = NOMBRE_DEP[u.dep] || ETIQUETAS[u.sigla] || ((u.sigla || u.u) + ' — ' + u.nombre);
      sel.appendChild(o);
    });
  })();

  function mostrarError(msg) {
    errLogin.hidden = false; errLogin.textContent = msg;
    $('#clave').value = ''; $('#clave').focus();
  }

  formLogin.addEventListener('submit', function (ev) {
    ev.preventDefault();
    var u = $('#usuario').value, c = $('#clave').value;
    errLogin.hidden = true;
    btnLogin.disabled = true; btnLogin.textContent = 'Verificando…';
    A.login(u, c).then(function (s) {
      if (!s) { mostrarError('Usuario o contraseña incorrectos. Verifique las mayúsculas y las tildes.'); return; }
      sesion = s;
      try { sessionStorage.setItem(A.claveSesion, JSON.stringify(s)); } catch (e) {}
      return abrirApp();
    }).catch(function (e) {
      mostrarError(e.message || 'No se pudo contactar el servidor.');
    }).then(function () {
      btnLogin.disabled = false; btnLogin.textContent = 'Ingresar';
    });
  });

  function salir() {
    if (temporizador) clearInterval(temporizador);
    A.salir();
    sesion = null; store = { reg: {}, bitacora: [] };
    try { sessionStorage.removeItem(A.claveSesion); } catch (e) {}
    $('#vApp').hidden = true; $('#vLogin').hidden = false; $('#sesion').hidden = true;
    $('#usuario').value = ''; $('#clave').value = ''; errLogin.hidden = true;
    window.scrollTo(0, 0);
  }
  $('#btnSalir').addEventListener('click', salir);

  /* ═════════════════════ estado de la interfaz ═════════════════════ */
  var vista = 'preguntas';
  var filtro = { pol: '', estado: '', dep: '', q: '' };

  function abrirApp() {
    $('#vLogin').hidden = true; $('#vApp').hidden = false; $('#sesion').hidden = false;
    $('#quienNombre').textContent = sesion.nombre;
    $('#quienRol').textContent = esAdmin()
      ? 'Administrador — Gerencia Técnica'
      : 'Dependencia responsable · ' + preguntasVisibles().length + ' preguntas asignadas';
    $('#subtitulo').textContent = D.etiquetaVigencia + ' — Agencia Catastral de Cundinamarca';
    if (!elSync) {
      elSync = document.createElement('span');
      elSync.id = 'sync';
      elSync.style.cssText = 'font-size:12px;font-weight:700;white-space:nowrap';
      $('#sesion').insertBefore(elSync, $('#btnSalir'));
    }
    elSync.hidden = !A.remoto;
    $('#pieApp').innerHTML = 'Agencia Catastral de Cundinamarca · ' + esc(D.etiquetaVigencia);
    vista = esAdmin() ? 'dashboard' : 'preguntas';
    filtro = { pol: '', estado: '', dep: '', q: '' };
    $('#contenido').innerHTML = '<div class="vacio">Cargando el avance…</div>';
    pintarTabs();

    return A.cargar().then(function (d) {
      store.reg = d.reg || {}; store.bitacora = d.bitacora || [];
      versionServidor = d.version || 0;
      // Primera vez en modo compartido: el administrador siembra el catálogo de preguntas.
      if (A.remoto && esAdmin() && !Object.keys(store.reg).length) {
        var catalogo = D.preguntas.map(function (p) {
          return { cod: p.id, pol: p.polx, dep: p.dep, tipo: p.tipoOriginal, ev: p.evidencia ? 'Sí' : 'No' };
        });
        return A.sembrar(catalogo).then(function () { return A.cargar(); }).then(function (d2) {
          store.reg = d2.reg || {}; versionServidor = d2.version || 0;
        });
      }
    }).then(function () {
      marcarSync('ok');
      arrancarSondeo();
      render();
      window.scrollTo(0, 0);
    }).catch(function (e) {
      $('#contenido').innerHTML = '<div class="vacio"><p><b>No se pudo cargar el avance.</b></p><p>' +
        esc(e.message) + '</p></div>';
      marcarSync('error', e.message);
    });
  }

  function render() { pintarTabs(); pintarBarra(); pintarContenido(); }

  function pintarTabs() {
    var t = esAdmin()
      ? [['dashboard', 'Tablero de seguimiento'], ['revision', 'Revisión de evidencias']]/*, ['datos', 'Sincronización y respaldo']]*/
      : [['preguntas', 'Mis preguntas'], ['resumen', 'Mi avance']];
    $('#tabs').innerHTML = t.map(function (x) {
      return '<button class="mini' + (vista === x[0] ? ' on' : '') + '" data-tab="' + x[0] + '" type="button">' + x[1] + '</button>';
    }).join('');
    Array.prototype.forEach.call($('#tabs').children, function (b) {
      b.addEventListener('click', function () { vista = b.dataset.tab; filtro.q = ''; render(); });
    });
  }

  function pintarBarra() {
    var b = $('#barra'), caja = b.parentNode;
    if (vista === 'datos' || vista === 'resumen') { caja.hidden = true; return; }
    caja.hidden = false;
    var visibles = preguntasVisibles();
    var pols = {};
    visibles.forEach(function (p) { pols[p.pol] = true; });
    var opPol = D.politicas.filter(function (p) { return pols[p.id]; })
      .map(function (p) { return '<option value="' + p.id + '"' + (filtro.pol === p.id ? ' selected' : '') + '>' + esc(p.nombre) + '</option>'; }).join('');
    var opDep = D.dependencias.filter(function (d) { return d.id !== 'DEP-GT'; })
      .map(function (d) { return '<option value="' + d.id + '"' + (filtro.dep === d.id ? ' selected' : '') + '>' + esc(d.nombre) + '</option>'; }).join('');
    var opEst = ['borrador', 'pendiente', 'aprobado', 'rechazado']
      .map(function (e) { return '<option value="' + e + '"' + (filtro.estado === e ? ' selected' : '') + '>' + ETIQ_ESTADO[e] + '</option>'; }).join('');

    b.innerHTML =
      (esAdmin() ? '<label for="fDep">Dependencia</label><select id="fDep"><option value="">Todas</option>' + opDep + '</select>' : '') +
      '<label for="fPol">Política</label><select id="fPol"><option value="">Todas</option>' + opPol + '</select>' +
      '<label for="fEst">Estado</label><select id="fEst"><option value="">Todos</option>' + opEst + '</select>' +
      '<input type="search" id="fQ" placeholder="Buscar por código o texto…" aria-label="Buscar pregunta" value="' + esc(filtro.q) + '">' +
      '<span class="conteo" id="cont"></span>';

    if ($('#fDep')) $('#fDep').addEventListener('change', function () { filtro.dep = this.value; pintarContenido(); });
    $('#fPol').addEventListener('change', function () { filtro.pol = this.value; pintarContenido(); });
    $('#fEst').addEventListener('change', function () { filtro.estado = this.value; pintarContenido(); });
    var tmr;
    $('#fQ').addEventListener('input', function () {
      var v = this.value; clearTimeout(tmr);
      tmr = setTimeout(function () { filtro.q = v; pintarContenido(); }, 220);
    });
  }

  function filtrar() {
    var t = norm(filtro.q);
    return preguntasVisibles().filter(function (p) {
      if (filtro.pol && p.pol !== filtro.pol) return false;
      if (filtro.dep && p.dep !== filtro.dep) return false;
      if (filtro.estado && reg(p.id).estado !== filtro.estado) return false;
      if (t && norm(p.id + ' ' + p.enunciado).indexOf(t) < 0) return false;
      return true;
    });
  }

  function pintarContenido() {
    if (!sesion) return;
    if (vista === 'preguntas') return listaPreguntas(false);
    if (vista === 'revision') return listaPreguntas(true);
    if (vista === 'dashboard') return tablero();
    if (vista === 'resumen') return resumenDependencia();
    if (vista === 'datos') return panelDatos();
  }

  /* ═════════════════════ conteos ═════════════════════ */
  function conteos(lista) {
    var c = { total: lista.length, borrador: 0, pendiente: 0, aprobado: 0, rechazado: 0,
      exigeEv: 0, conEv: 0, respondidas: 0 };
    lista.forEach(function (p) {
      var r = reg(p.id);
      if (c[r.estado] == null) c[r.estado] = 0;
      c[r.estado]++;
      if (p.evidencia) { c.exigeEv++; if (r.evis.length) c.conEv++; }
      if (tieneRespuesta(p, r.resp)) c.respondidas++;
    });
    return c;
  }
  function tieneRespuesta(p, v) {
    if (!v) return false;
    if (p.tipo === 'unica') return !!v.op;
    if (p.tipo === 'multiple') return !!(v.ops && v.ops.length);
    if (p.tipo === 'multiple_num') return Object.keys(v.num || {}).some(function (k) { return v.num[k] !== ''; });
    if (p.tipo === 'numerica') return v.n !== '' && v.n != null;
    if (p.tipo === 'texto') return !!String(v.t || '').trim();
    if (p.tipo === 'matricial') return Object.keys(v.m || {}).length > 0;
    return false;
  }

  /* ═════════════════════ lista de preguntas ═════════════════════ */
  function listaPreguntas(modoAdmin) {
    var lista = filtrar();
    if ($('#cont')) $('#cont').textContent = fmt(lista.length) + ' de ' + fmt(preguntasVisibles().length) + ' preguntas';
    var c = conteos(preguntasVisibles());
    var html = '<div class="kpis">' + [
      { v: fmt(c.total), e: 'Preguntas asignadas' },
      { v: fmt(c.respondidas), e: 'Con respuesta diligenciada' },
      { v: fmt(c.borrador), e: 'Sin enviar' },
      { v: fmt(c.pendiente), e: 'Pendientes de revisión', k: 'pend' },
      { v: fmt(c.aprobado), e: 'Aprobadas', k: 'bien' },
      { v: fmt(c.rechazado), e: 'Rechazadas · por subsanar', k: 'mal' }
    ].map(function (k) {
      return '<div class="kpi ' + (k.k || '') + '"><div class="v">' + k.v + '</div><div class="e">' + k.e + '</div></div>';
    }).join('') + '</div>';

    if (!lista.length) {
      html += '<div class="vacio"><p><b>No hay preguntas para mostrar.</b></p><p>' +
        (preguntasVisibles().length ? 'Ajuste los filtros de la barra superior.'
          : 'Su dependencia no tiene preguntas asignadas en esta vigencia. Comuníquese con la Gerencia Técnica.') +
        '</p></div>';
      $('#contenido').innerHTML = html;
      return;
    }
    html += '<div class="preguntas">' + lista.map(function (p) { return tarjeta(p, modoAdmin); }).join('') + '</div>';
    $('#contenido').innerHTML = html;
    lista.forEach(function (p) { cablear(p, modoAdmin); });
  }

  function tarjeta(p, modoAdmin) {
    var r = reg(p.id), pol = POL[p.pol] || { nombre: p.seccion };
    var editable = puedeEditar(p);
    var h = '<article class="q ' + r.estado + '" id="q-' + p.id + '">';
    h += '<div class="q-cab">' +
      '<span class="cod">' + esc(p.id) + '</span>' +
      '<span class="meta"><b>' + esc(pol.nombre) + '</b>' +
        (modoAdmin ? esc((DEP[p.dep] || {}).nombre || '') + ' · ' : '') +
        'Pregunta ' + p.num + ' de 452 · página ' + p.pagina + ' del formulario' +
        (p.evidencia ? ' · exige evidencia' : '') + '</span>' +
      '<span class="pill p-tipo">' + ETIQ_TIPO[p.tipo] + '</span>' +
      '<span class="pill p-' + r.estado + '">' + ETIQ_ESTADO[r.estado] + '</span></div>';

    h += '<div class="q-cuerpo' + (editable ? '' : ' solo-lectura') + '">';
    h += '<p class="q-enun">' + esc(p.enunciado) + '</p>';
    if (p.nota) h += '<p class="q-nota">' + esc(p.nota) + '</p>';
    h += control(p, r, editable);

    if (p.evidencia || (r.evis && r.evis.length)) h += bloqueEvidencias(p, r, editable);

    if (r.estado === 'rechazado' && r.obs) {
      h += '<div class="obs"><b>Motivo del rechazo</b>' + esc(r.obs) +
        '<div class="meta">Registrado por ' + esc(r.revPor || '') + ' el ' + esc(r.revFecha || '') + '</div></div>';
    }
    if (r.estado === 'aprobado') {
      h += '<div class="obs" style="background:#E3F5E9;border-color:#BCE3C9;color:#0b6b2e">' +
        '<b>Evidencia aprobada</b>Aprobada por ' + esc(r.revPor || '') + ' el ' + esc(r.revFecha || '') + '</div>';
    }

    h += '<div class="acciones">';
    if (modoAdmin && esAdmin()) {
      h += '<button class="mini ok" data-acc="aprobar" data-q="' + p.id + '" type="button">Aprobar</button>' +
           '<button class="mini no" data-acc="rechazar" data-q="' + p.id + '" type="button">Rechazar…</button>';
      if (r.estado === 'aprobado' || r.estado === 'rechazado')
        h += '<span class="msg" style="color:var(--tenue);font-weight:600">Puede modificar el estado en cualquier momento.</span>';
      if (r.estado === 'borrador')
        h += '<span class="msg" style="color:var(--tenue);font-weight:600">La dependencia aún no ha enviado esta pregunta.</span>';
      if (r.envPor)
        h += '<span class="msg" style="color:var(--tenue);font-weight:600">Enviada por ' + esc(r.envPor) + ' el ' + esc(r.envFecha || '') + '</span>';
    } else if (editable) {
      h += '<button class="btn" data-acc="enviar" data-q="' + p.id + '" type="button">' +
        (r.estado === 'rechazado' ? 'Reenviar subsanada' : 'Enviar a revisión') + '</button>';
      if (r.actualizado) h += '<span class="msg" style="color:var(--tenue);font-weight:600">Guardado ' + esc(r.actualizado) + '</span>';
    } else if (!esAdmin() && r.estado === 'aprobado') {
      h += '<span class="msg" style="color:var(--tenue);font-weight:600">Aprobada por la Gerencia Técnica: ya no admite cambios.</span>';
    }
    h += '</div></div></article>';
    return h;
  }

  /* ── controles por tipo de respuesta ── */
  function control(p, r, ed) {
    var v = r.resp || {}, dis = ed ? '' : ' disabled';
    var h = '';
    if (p.tipo === 'unica' || p.tipo === 'multiple') {
      var multi = p.tipo === 'multiple';
      var sel = multi ? (v.ops || []) : (v.op ? [v.op] : []);
      h += '<div class="opciones">';
      p.opciones.forEach(function (o) {
        var m = sel.indexOf(o.id) >= 0;
        h += '<label class="op' + (m ? ' marcada' : '') + '" data-op="' + o.id + '">' +
          '<input type="' + (multi ? 'checkbox' : 'radio') + '" name="' + p.id + '" value="' + o.id + '"' +
          (m ? ' checked' : '') + dis + '>' +
          '<span class="txt">' + esc(o.texto) +
          (o.txt ? '<span class="extra"><input type="text" data-txt="' + o.id + '" placeholder="Detalle su respuesta" value="' +
            esc((v.txt || {})[o.id] || '') + '"' + dis + '></span>' : '') +
          '</span></label>';
      });
      h += '</div>';
    } else if (p.tipo === 'multiple_num') {
      h += '<div class="opciones">';
      p.opciones.forEach(function (o) {
        h += '<div class="op-num"><span>' + esc(o.texto) + '</span>' +
          '<input type="number" step="any" min="0" data-num="' + o.id + '" placeholder="0" value="' +
          esc((v.num || {})[o.id] != null ? v.num[o.id] : '') + '"' + dis + '></div>';
      });
      h += '</div>';
    } else if (p.tipo === 'numerica') {
      h += '<div style="max-width:280px"><label for="n-' + p.id + '">Valor</label>' +
        '<input type="number" step="any" id="n-' + p.id + '" data-n="1" value="' + esc(v.n != null ? v.n : '') + '"' + dis + '></div>';
    } else if (p.tipo === 'matricial' && p.matriz) {
      var m = v.m || {};
      h += '<div class="tabla-wrap"><table class="matriz"><thead><tr><th></th>' +
        p.matriz.columnas.map(function (c) { return '<th class="c">' + esc(c) + '</th>'; }).join('') +
        '</tr></thead><tbody>';
      p.matriz.filas.forEach(function (f, i) {
        h += '<tr><th scope="row" style="font-weight:600">' + esc(f) + '</th>' +
          p.matriz.columnas.map(function (c, j) {
            return '<td class="c"><input type="radio" name="' + p.id + '-f' + i + '" data-m="' + i + '" value="' + j + '"' +
              (String(m[i]) === String(j) ? ' checked' : '') + dis +
              ' aria-label="' + esc(f + ' — ' + c) + '"></td>';
          }).join('') + '</tr>';
      });
      h += '</tbody></table></div>';
    } else {
      h += '<textarea data-t="1" placeholder="Escriba su respuesta"' + dis + '>' + esc(v.t || '') + '</textarea>';
    }
    return h;
  }

  function bloqueEvidencias(p, r, ed) {
    var h = '<div class="bloque"><h4>Evidencias' + (p.evidencia ? ' (obligatorias para esta pregunta)' : '') + '</h4>';
    h += '<div class="evi-lista">';
    if (!r.evis.length) h += '<div class="evi" style="color:var(--tenue)">Sin evidencias registradas.</div>';
    r.evis.forEach(function (e) {
      h += '<div class="evi"><span class="nom"><a href="' + esc(e.url) + '" target="_blank" rel="noopener">' +
        esc(e.nombre) + '</a></span><span class="f">' + esc(e.fecha) + '</span>' +
        (ed ? '<button class="mini no" data-acc="quitar-evi" data-q="' + p.id + '" data-e="' + e.id + '" type="button">Quitar</button>' : '') +
        '</div>';
    });
    h += '</div>';
    if (ed) {
      h += '<div class="evi-form">' +
        '<div><label>Nombre del soporte</label><input type="text" data-evinom="' + p.id + '" placeholder="Ej.: Acta comité 2025-03"></div>' +
        '<div><label>Enlace en SharePoint</label><input type="url" data-eviurl="' + p.id + '" placeholder="https://acatastral.sharepoint.com/…"></div>' +
        '<button class="mini on" data-acc="add-evi" data-q="' + p.id + '" type="button">Registrar</button></div>' +
        '<p style="font-size:12.5px;color:var(--tenue);margin:8px 0 0">Suba el archivo de la justificación de su respuesta a la ' +
        '<a href="' + esc(D.sharepointEvidencias) + '" target="_blank" rel="noopener">carpeta de evidencias de la Gerencia Técnica</a> ' +
        'y pegue aquí el enlace del archivo.</p>';
    }
    h += '</div>';
    return h;
  }

  /* ── eventos de cada tarjeta ── */
  function cablear(p, modoAdmin) {
    var card = $('#q-' + p.id);
    if (!card) return;
    var editable = puedeEditar(p);

    function leer() {
      var v = {};
      if (p.tipo === 'unica') {
        var s = card.querySelector('input[type=radio]:checked');
        v.op = s ? s.value : ''; v.txt = {};
        Array.prototype.forEach.call(card.querySelectorAll('input[data-txt]'), function (i) {
          if (i.value) v.txt[i.dataset.txt] = i.value;
        });
      } else if (p.tipo === 'multiple') {
        v.ops = Array.prototype.map.call(card.querySelectorAll('input[type=checkbox]:checked'), function (i) { return i.value; });
        v.txt = {};
        Array.prototype.forEach.call(card.querySelectorAll('input[data-txt]'), function (i) {
          if (i.value) v.txt[i.dataset.txt] = i.value;
        });
      } else if (p.tipo === 'multiple_num') {
        v.num = {};
        Array.prototype.forEach.call(card.querySelectorAll('input[data-num]'), function (i) { v.num[i.dataset.num] = i.value; });
      } else if (p.tipo === 'numerica') {
        var n = card.querySelector('input[data-n]');
        v.n = n ? n.value : '';
      } else if (p.tipo === 'matricial') {
        v.m = {};
        Array.prototype.forEach.call(card.querySelectorAll('input[data-m]:checked'), function (i) { v.m[i.dataset.m] = i.value; });
      } else {
        var t = card.querySelector('textarea[data-t]');
        v.t = t ? t.value : '';
      }
      return v;
    }

    if (editable) {
      card.addEventListener('change', function (ev) {
        if (!ev.target.matches('input[type=radio],input[type=checkbox],input[type=number],input[data-txt],textarea[data-t]')) return;
        // «Ninguna de las anteriores» y «No aplica» excluyen al resto
        if (p.tipo === 'multiple' && ev.target.type === 'checkbox' && ev.target.checked) {
          var o = null;
          p.opciones.forEach(function (x) { if (x.id === ev.target.value) o = x; });
          var excl = o && (o.ninguna || o.na);
          Array.prototype.forEach.call(card.querySelectorAll('input[type=checkbox]'), function (i) {
            if (i === ev.target) return;
            var oo = null;
            p.opciones.forEach(function (x) { if (x.id === i.value) oo = x; });
            if (excl) i.checked = false;
            else if (oo && (oo.ninguna || oo.na)) i.checked = false;
          });
        }
        Array.prototype.forEach.call(card.querySelectorAll('.op'), function (l) {
          var i = l.querySelector('input'); l.classList.toggle('marcada', !!(i && i.checked));
        });
        guardarRespuestaDiferido(p.id, leer());
      });
      card.addEventListener('input', function (ev) {
        if (!ev.target.matches('textarea[data-t],input[data-txt],input[data-num],input[data-n]')) return;
        guardarRespuestaDiferido(p.id, leer());
      });
    }

    card.addEventListener('click', function (ev) {
      var b = ev.target.closest ? ev.target.closest('button[data-acc]') : null;
      if (!b) return;
      var acc = b.dataset.acc;
      try {
        if (acc === 'enviar') {
          if (Q[p.id].evidencia && !reg(p.id).evis.length &&
              !confirm('Esta pregunta exige evidencia y no ha registrado ninguna. ¿Enviar de todos modos?')) return;
          clearTimeout(debounce[p.id]);
          b.disabled = true;
          Promise.resolve(guardarRespuesta(p.id, leer()))
            .then(function () { return enviarARevision(p.id); })
            .then(pintarContenido);
        } else if (acc === 'add-evi') {
          var nom = card.querySelector('input[data-evinom]').value.trim();
          var url = card.querySelector('input[data-eviurl]').value.trim();
          if (!nom || !url) { alert('Registre el nombre del soporte y el enlace.'); return; }
          if (!/^https?:\/\//i.test(url)) { alert('El enlace debe empezar por http:// o https://'); return; }
          b.disabled = true;
          Promise.resolve(agregarEvidencia(p.id, nom, url)).then(pintarContenido);
        } else if (acc === 'quitar-evi') {
          if (!confirm('¿Quitar esta evidencia?')) return;
          Promise.resolve(quitarEvidencia(p.id, b.dataset.e)).then(pintarContenido);
        } else if (acc === 'aprobar') {
          b.disabled = true;
          Promise.resolve(revisar(p.id, 'aprobado', '')).then(pintarContenido);
        } else if (acc === 'rechazar') {
          modalRechazo(p.id);
        }
      } catch (e) { alert(e.message); b.disabled = false; }
    });
  }

  /* ═════════════════════ modal de rechazo ═════════════════════ */
  function modalRechazo(qid) {
    var p = Q[qid];
    var wrap = document.createElement('div');
    wrap.className = 'velo';
    wrap.innerHTML =
      '<div class="modal" role="dialog" aria-modal="true" aria-labelledby="mt">' +
      '<h3 id="mt">Rechazar evidencia</h3>' +
      '<p class="sub">' + esc(p.id) + ' · ' + esc((DEP[p.dep] || {}).nombre || '') + '</p>' +
      '<p style="font-size:14px;margin:0 0 14px;color:var(--tinta-2)">' + esc(p.enunciado.slice(0, 180)) + '</p>' +
      '<label for="mObs">Motivo del rechazo / observaciones</label>' +
      '<textarea id="mObs" placeholder="Explique qué debe corregir la dependencia para subsanar."></textarea>' +
      '<div class="err" id="mErr" hidden>Debe registrar el motivo antes de confirmar el rechazo.</div>' +
      '<div class="pie"><button class="mini" id="mCancel" type="button">Cancelar</button>' +
      '<button class="btn" id="mOk" type="button" style="background:var(--st-critica);border-color:var(--st-critica)">Confirmar rechazo</button></div>' +
      '</div>';
    $('#modales').appendChild(wrap);
    var obs = $('#mObs', wrap); obs.focus();
    function cerrar() { wrap.parentNode.removeChild(wrap); document.removeEventListener('keydown', alEscape); }
    function alEscape(e) { if (e.key === 'Escape') cerrar(); }
    document.addEventListener('keydown', alEscape);
    $('#mCancel', wrap).addEventListener('click', cerrar);
    wrap.addEventListener('click', function (e) { if (e.target === wrap) cerrar(); });
    $('#mOk', wrap).addEventListener('click', function () {
      if (!obs.value.trim()) { $('#mErr', wrap).hidden = false; obs.focus(); return; }
      try {
        var t = obs.value;
        cerrar();
        Promise.resolve(revisar(qid, 'rechazado', t)).then(pintarContenido);
      } catch (e) { alert(e.message); }
    });
  }

  /* ═════════════════════ gráficos ═════════════════════ */
  var COLOR = { borrador: 'var(--linea)', pendiente: 'var(--st-media)',
                aprobado: 'var(--st-bien)', rechazado: 'var(--st-critica)' };

  function barrasApiladas(filas) {
    return '<div class="barras">' + filas.map(function (f) {
      var tot = f.total || 1;
      return '<div class="fila-b"><span class="et" title="' + esc(f.et) + '">' + esc(f.et) + '</span>' +
        '<span class="pista">' + ['aprobado', 'pendiente', 'rechazado', 'borrador'].map(function (k) {
          return f[k] ? '<i style="width:' + (f[k] * 100 / tot) + '%;background:' + COLOR[k] + '" title="' +
            ETIQ_ESTADO[k] + ': ' + f[k] + '"></i>' : '';
        }).join('') + '</span>' +
        '<span class="n">' + fmt(f.total) + '</span></div>';
    }).join('') + '</div>';
  }
  function leyendaEstados() {
    return '<div class="leyenda">' + ['aprobado', 'pendiente', 'rechazado', 'borrador'].map(function (k) {
      return '<span><i style="background:' + COLOR[k] + '"></i>' + ETIQ_ESTADO[k] + '</span>';
    }).join('') + '</div>';
  }
  function kpis(arr) {
    return '<div class="kpis">' + arr.map(function (k) {
      return '<div class="kpi ' + (k.k || '') + '"><div class="v">' + k.v + '</div><div class="e">' + k.e + '</div></div>';
    }).join('') + '</div>';
  }
  function agrupar(lista, llave, nombre, base) {
    var origen = base || preguntasVisibles();
    return lista.map(function (g) {
      var l = origen.filter(function (p) { return p[llave] === g.id; });
      var cc = conteos(l);
      return { et: nombre(g), total: cc.total, borrador: cc.borrador, pendiente: cc.pendiente,
        aprobado: cc.aprobado, rechazado: cc.rechazado };
    }).filter(function (f) { return f.total; }).sort(function (a, b) { return b.total - a.total; });
  }

  /* ═════════════════════ tablero del administrador ═════════════════════ */
  function tablero() {
    var todas = filtrar(), c = conteos(todas);
    var evis = 0, evisPend = 0, evisApro = 0, evisRech = 0;
    todas.forEach(function (p) {
      var r = reg(p.id);
      evis += r.evis.length;
      if (r.evis.length) {
        if (r.estado === 'aprobado') evisApro += r.evis.length;
        else if (r.estado === 'rechazado') evisRech += r.evis.length;
        else if (r.estado === 'pendiente') evisPend += r.evis.length;
      }
    });
    var h = kpis([
      { v: fmt(c.total), e: 'Preguntas del formulario' },
      { v: fmt(c.borrador), e: 'Sin enviar por la dependencia' },
      { v: fmt(c.pendiente), e: 'Pendientes de revisión', k: 'pend' },
      { v: fmt(c.aprobado), e: 'Aprobadas', k: 'bien' },
      { v: fmt(c.rechazado), e: 'Rechazadas', k: 'mal' },
      { v: pct(c.aprobado, c.total) + '%', e: 'Avance de aprobación' }
    ]) + kpis([
      { v: fmt(c.exigeEv), e: 'Preguntas que exigen evidencia' },
      { v: fmt(c.conEv), e: 'Con evidencia registrada' },
      { v: fmt(evis), e: 'Evidencias cargadas en total' },
      { v: fmt(evisPend), e: 'Evidencias pendientes de revisión', k: 'pend' },
      { v: fmt(evisApro), e: 'Evidencias aprobadas', k: 'bien' },
      { v: fmt(evisRech), e: 'Evidencias rechazadas', k: 'mal' }
    ]);

    var porDep = agrupar(D.dependencias.filter(function (d) { return d.id !== 'DEP-GT'; }), 'dep', function (d) { return d.nombre; }, todas);
    var porPol = agrupar(D.politicas, 'pol', function (p) { return p.nombre; }, todas);

    h += '<div class="dos">' +
      '<div class="panel"><h3>Estado por dependencia</h3><p class="sub">Preguntas asignadas a cada dependencia y su situación de revisión.</p>' +
        leyendaEstados() + barrasApiladas(porDep) + '</div>' +
      '<div class="panel"><h3>Estado por política MIPG</h3><p class="sub">Las ' + porPol.length + ' políticas y secciones del formulario.</p>' +
        leyendaEstados() + barrasApiladas(porPol) + '</div>' +
      '</div>';

    var rech = todas.filter(function (p) { return reg(p.id).estado === 'rechazado'; });
    h += '<div class="panel"><h3>Preguntas devueltas para subsanación</h3>' +
      '<p class="sub">' + (rech.length ? fmt(rech.length) + ' preguntas esperan corrección de la dependencia.' : 'No hay preguntas rechazadas.') + '</p>';
    if (rech.length) {
      h += '<div class="tabla-wrap"><table><thead><tr><th>Código</th><th>Dependencia</th><th>Motivo</th><th>Revisada</th></tr></thead><tbody>' +
        rech.map(function (p) {
          var r = reg(p.id);
          return '<tr><td><b>' + esc(p.id) + '</b></td><td>' + esc((DEP[p.dep] || {}).nombre || '') + '</td><td>' +
            esc(r.obs) + '</td><td>' + esc(r.revFecha || '') + '</td></tr>';
        }).join('') + '</tbody></table></div>';
    }
    h += '</div>';
    $('#contenido').innerHTML = h;
    if ($('#cont')) $('#cont').textContent = fmt(todas.length) + ' preguntas';
  }

  /* ═════════════════════ resumen de la dependencia ═════════════════════ */
  function resumenDependencia() {
    var todas = preguntasVisibles(), c = conteos(todas);
    var h = kpis([
      { v: fmt(c.total), e: 'Preguntas asignadas' },
      { v: fmt(c.respondidas), e: 'Respondidas' },
      { v: fmt(c.borrador), e: 'Sin enviar' },
      { v: fmt(c.pendiente), e: 'En revisión', k: 'pend' },
      { v: fmt(c.aprobado), e: 'Aprobadas', k: 'bien' },
      { v: fmt(c.rechazado), e: 'Por subsanar', k: 'mal' }
    ]);

    h += '<div class="panel"><h3>Mi avance por política</h3><p class="sub">' +
      esc(sesion.nombre) + ' · ' + esc((DEP[sesion.dep] || {}).ambito || '') + '</p>' +
      leyendaEstados() + barrasApiladas(agrupar(D.politicas, 'pol', function (p) { return p.nombre; })) + '</div>';

    var rech = todas.filter(function (p) { return reg(p.id).estado === 'rechazado'; });
    h += '<div class="panel"><h3>Lo que debo subsanar</h3><p class="sub">' +
      (rech.length ? 'La Gerencia Técnica devolvió estas preguntas con observaciones.' : 'No tiene preguntas devueltas.') + '</p>';
    if (rech.length) {
      h += '<div class="tabla-wrap"><table><thead><tr><th>Código</th><th>Pregunta</th><th>Motivo del rechazo</th></tr></thead><tbody>' +
        rech.map(function (p) {
          return '<tr><td><b>' + esc(p.id) + '</b></td><td>' + esc(p.enunciado.slice(0, 90)) + '…</td><td>' +
            esc(reg(p.id).obs) + '</td></tr>';
        }).join('') + '</tbody></table></div>';
    }
    h += '</div>';

    /*
    h += '<div class="panel"><h3>' + (A.remoto ? 'Copia de mi avance' : 'Respaldo de mi avance') + '</h3>' +
      '<p class="sub">' + (A.remoto
        ? 'Su avance ya está guardado en la hoja institucional; esta descarga es solo una copia de consulta.'
        : 'Descargue el archivo si va a continuar desde otro equipo o navegador.') + '</p>' +
      '<div class="acciones"><button class="mini on" id="expJson" type="button">Descargar copia (JSON)</button>' +
      (A.remoto ? '' :
        '<button class="mini" id="impJson" type="button">Cargar respaldo…</button>' +
        '<input type="file" id="fileJson" accept=".json" hidden>') +
      '</div></div>';*/

    $('#contenido').innerHTML = h;
    cablearRespaldo();
  }

  /* ═════════════════════ sincronización y respaldo (admin) ═════════════════════ */
  
  function panelDatos() {
    var h = '';

    if (A.remoto) {
      h += '<div class="panel"><h3>Estado de la sincronización</h3>' +
        '<p class="sub">Todas las dependencias escriben en la misma hoja de cálculo. ' +
        'Cada aprobación o rechazo actualiza al instante la hoja <code>ESTADOS</code>, ' +
        'con las columnas <code>Politica, Cod_Pregunta, Aprobada, Rechazada</code> del Excel de SharePoint.</p>' +
        '<div class="acciones"><button class="mini on" id="refrescar" type="button">Traer los últimos cambios</button>' +
        '<button class="mini no" id="limpiar" type="button">Dejar la base en limpio</button>' +
        '<a class="mini" href="' + esc(A.api.replace(/\/exec.*$/, '')) + '" target="_blank" rel="noopener" ' +
        'style="text-decoration:none">Abrir el proyecto del backend</a>' +
        '<span class="msg" id="msgRef" style="color:var(--tenue);font-weight:600"></span></div>' +
        '<p style="font-size:13px;color:var(--tenue);margin:12px 0 0">Si edita a mano la hoja ' +
        '<code>ESTADOS</code>, el cambio se refleja en el aplicativo en la siguiente carga. ' +
        'La hoja se descarga como .xlsx desde <i>Archivo → Descargar → Microsoft Excel</i> y ese archivo ' +
        'es el que se sube al Excel de SharePoint.</p></div>';
    } else {
      h += '<div class="panel"><h3>Modo local</h3>' +
        '<p class="sub">Este navegador guarda su propio avance y no lo comparte con nadie. ' +
        'Para trabajar entre varias dependencias, configure el backend siguiendo ' +
        '<code>backend/INSTALACION.md</code> y pegue la URL del Web App en <code>config.js</code>.</p></div>';
    }

    h += '<div class="panel"><h3>Exportar para el Excel de SharePoint</h3>' +
      '<p class="sub">Dos archivos: el de estados conserva exactamente la estructura del libro actual; ' +
      'el consolidado añade respuestas, evidencias y observaciones.</p>' +
      '<div class="acciones">' +
      '<button class="mini on" id="expEstados" type="button">Exportar estados (CSV)</button>' +
      '<button class="mini" id="expDetalle" type="button">Exportar consolidado (CSV)</button>' +
      (A.remoto ? '' :
        '<button class="mini" id="impEstados" type="button">Importar estados desde CSV…</button>' +
        '<input type="file" id="fileCsv" accept=".csv,text/csv" hidden>') +
      '</div></div>';

    h += '<div class="panel"><h3>' + (A.remoto ? 'Copia completa' : 'Respaldo completo') + '</h3>' +
      '<p class="sub">Incluye respuestas, evidencias, estados, observaciones y la bitácora de revisión.</p>' +
      '<div class="acciones"><button class="mini on" id="expJson" type="button">Descargar copia (JSON)</button>' +
      (A.remoto ? '' :
        '<button class="mini" id="impJson" type="button">Cargar respaldo…</button>' +
        '<input type="file" id="fileJson" accept=".json" hidden>' +
        '<button class="mini no" id="borrar" type="button">Borrar todo el avance de este navegador</button>') +
      '</div></div>';

    var b = (store.bitacora || []).slice(-60).reverse();
    h += '<div class="panel"><h3>Bitácora de trazabilidad</h3><p class="sub">Últimos ' + b.length + ' movimientos registrados.</p>';
    h += b.length ? '<div class="tabla-wrap"><table><thead><tr><th>Fecha</th><th>Pregunta</th><th>Acción</th><th>Usuario</th><th>Detalle</th></tr></thead><tbody>' +
      b.map(function (x) {
        return '<tr><td>' + esc(x.fecha) + '</td><td><b>' + esc(x.qid) + '</b></td><td>' + esc(x.accion) +
          '</td><td>' + esc(x.por) + '</td><td>' + esc(String(x.detalle).slice(0, 80)) + '</td></tr>';
      }).join('') + '</tbody></table></div>' : '<p style="color:var(--tenue)">Sin movimientos todavía.</p>';
    h += '</div>';

    $('#contenido').innerHTML = h;
    cablearRespaldo();

    if ($('#limpiar')) $('#limpiar').addEventListener('click', function () {
      if (!confirm('Se borrarán todas las respuestas, evidencias, estados, observaciones y la bitácora ' +
        'de la hoja de cálculo. El catálogo de las 452 preguntas se conserva.\n\n' +
        'Esto afecta a todas las dependencias. ¿Continuar?')) return;
      if (!confirm('Confirme una vez más: el avance de prueba se perderá y no se puede deshacer.')) return;
      var b = this; b.disabled = true; b.textContent = 'Limpiando…';
      Promise.resolve(A.limpiar()).then(function (r) {
        return recargar().then(function () {
          alert((r && r.mensaje) || 'Base en limpio.');
        });
      }).catch(function (e) { alert('No se pudo limpiar: ' + e.message); })
        .then(function () { pintarContenido(); });
    });

    if ($('#refrescar')) $('#refrescar').addEventListener('click', function () {
      $('#msgRef').textContent = 'Consultando…';
      recargar().then(function () { if ($('#msgRef')) $('#msgRef').textContent = 'Actualizado ' + ahora(); });
    });

    $('#expEstados').addEventListener('click', function () {
      var f = [['Politica', 'Cod_Pregunta', 'Aprobada', 'Rechazada']];
      D.preguntas.forEach(function (p) {
        var r = reg(p.id);
        f.push([p.polx, p.id,
          r.estado === 'aprobado' ? 1 : (r.estado === 'rechazado' ? 0 : ''),
          r.estado === 'rechazado' ? 1 : (r.estado === 'aprobado' ? 0 : '')]);
      });
      bajarCSV(f, 'estado_pregunta_simulador_furag.csv');
    });
    $('#expDetalle').addEventListener('click', function () {
      var f = [['Cod_Pregunta', 'Politica', 'Dependencia', 'Tipo_Respuesta', 'Requiere_Evidencia',
        'Estado', 'Respuesta_Resumen', 'Evidencias', 'Observacion_Rechazo',
        'Enviado_Por', 'Fecha_Envio', 'Revisado_Por', 'Fecha_Revision']];
      D.preguntas.forEach(function (p) {
        var r = reg(p.id);
        f.push([p.id, p.polx, (DEP[p.dep] || {}).nombre || '',
          p.tipoOriginal, p.evidencia ? 'Sí' : 'No', ETIQ_ESTADO[r.estado],
          resumenRespuesta(p, r.resp), r.evis.map(function (e) { return e.nombre + ' <' + e.url + '>'; }).join(' | '),
          r.obs || '', r.envPor || '', r.envFecha || '', r.revPor || '', r.revFecha || '']);
      });
      bajarCSV(f, 'consolidado_furag_simulador.csv');
    });
    if ($('#impEstados')) {
      $('#impEstados').addEventListener('click', function () { $('#fileCsv').click(); });
      $('#fileCsv').addEventListener('change', function () {
        var f = this.files[0]; if (!f) return;
        var fr = new FileReader();
        fr.onload = function () {
          try {
            var n = aplicarCSVEstados(String(fr.result));
            alert('Se actualizaron ' + n + ' preguntas desde el archivo.');
            pintarContenido();
          } catch (e) { alert('No se pudo leer el archivo: ' + e.message); }
        };
        fr.readAsText(f, 'utf-8');
        this.value = '';
      });
    }
    if ($('#borrar')) $('#borrar').addEventListener('click', function () {
      if (!confirm('Se borrarán todas las respuestas, evidencias y estados guardados en este navegador. ¿Continuar?')) return;
      A.borrar(); store = { reg: {}, bitacora: [] }; pintarContenido();
    });
  }

  function resumenRespuesta(p, v) {
    if (!v) return '';
    if (p.tipo === 'unica') { var o = null; p.opciones.forEach(function (x) { if (x.id === v.op) o = x; });
      return (o ? o.texto : '') + (v.txt && v.txt[v.op] ? ' — ' + v.txt[v.op] : ''); }
    if (p.tipo === 'multiple') return (v.ops || []).map(function (id) {
      var t = ''; p.opciones.forEach(function (x) { if (x.id === id) t = x.texto; });
      return t + (v.txt && v.txt[id] ? ' — ' + v.txt[id] : '');
    }).join(' | ');
    if (p.tipo === 'multiple_num') return Object.keys(v.num || {}).filter(function (k) { return v.num[k] !== ''; })
      .map(function (k) { var t = ''; p.opciones.forEach(function (x) { if (x.id === k) t = x.texto; }); return t + '=' + v.num[k]; }).join(' | ');
    if (p.tipo === 'numerica') return v.n != null ? String(v.n) : '';
    if (p.tipo === 'matricial') return Object.keys(v.m || {}).map(function (i) {
      return p.matriz.filas[i] + ' = ' + p.matriz.columnas[v.m[i]];
    }).join(' | ');
    return String(v.t || '');
  }

  function cablearRespaldo() {
    var ej = $('#expJson'); if (!ej) return;
    ej.addEventListener('click', function () {
      bajar(JSON.stringify({ app: 'furag-simulador-acc', vigencia: D.vigencia, fecha: ahora(),
        modo: A.remoto ? 'compartido' : 'local', store: store }, null, 1),
        'respaldo_furag_' + hoy() + '.json', 'application/json');
    });
    if (!$('#impJson')) return;
    $('#impJson').addEventListener('click', function () { $('#fileJson').click(); });
    $('#fileJson').addEventListener('change', function () {
      var f = this.files[0]; if (!f) return;
      var fr = new FileReader();
      fr.onload = function () {
        try {
          var d = JSON.parse(String(fr.result));
          if (!d.store || !d.store.reg) throw new Error('El archivo no tiene la estructura esperada.');
          if (!confirm('Se reemplazará el avance guardado en este navegador. ¿Continuar?')) return;
          store = d.store; if (!store.bitacora) store.bitacora = [];
          A.importar(store); render();
          alert('Respaldo cargado.');
        } catch (e) { alert('No se pudo cargar: ' + e.message); }
      };
      fr.readAsText(f, 'utf-8');
      this.value = '';
    });
  }

  /* CSV → aplicación (solo en modo local; en modo compartido se edita la hoja ESTADOS) */
  function aplicarCSVEstados(txt) {
    var filas = parseCSV(txt);
    if (!filas.length) return 0;
    var cab = filas[0].map(function (x) { return norm(x); });
    var iCod = cab.indexOf('cod_pregunta'), iA = cab.indexOf('aprobada'), iR = cab.indexOf('rechazada');
    var iObs = cab.indexOf('observacion_rechazo');
    if (iCod < 0) throw new Error('Falta la columna Cod_Pregunta.');
    var n = 0;
    for (var i = 1; i < filas.length; i++) {
      var cod = String(filas[i][iCod] || '').trim();
      if (!Q[cod]) continue;
      var a = iA >= 0 ? String(filas[i][iA] || '').trim() : '';
      var rj = iR >= 0 ? String(filas[i][iR] || '').trim() : '';
      var r = reg(cod), nuevo = null;
      if (a === '1') nuevo = 'aprobado';
      else if (rj === '1') nuevo = 'rechazado';
      if (!nuevo) continue;
      r.estado = nuevo;
      if (nuevo === 'rechazado' && iObs >= 0 && filas[i][iObs]) r.obs = filas[i][iObs];
      r.revFecha = ahora(); r.revPor = sesion.nombre + ' (importado del Excel)';
      n++;
    }
    A.importar(store);
    return n;
  }

  function parseCSV(t) {
    var filas = [], f = [], c = '', q = false;
    t = t.replace(/^﻿/, '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    for (var i = 0; i < t.length; i++) {
      var ch = t[i];
      if (q) {
        if (ch === '"') { if (t[i + 1] === '"') { c += '"'; i++; } else q = false; }
        else c += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',' || ch === ';') { f.push(c); c = ''; }
      else if (ch === '\n') { f.push(c); filas.push(f); f = []; c = ''; }
      else c += ch;
    }
    if (c !== '' || f.length) { f.push(c); filas.push(f); }
    return filas.filter(function (x) { return x.length > 1 || (x[0] || '').trim(); });
  }
  function bajarCSV(filas, nombre) {
    var t = filas.map(function (f) {
      return f.map(function (c) {
        c = String(c == null ? '' : c);
        return /[",;\n]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c;
      }).join(',');
    }).join('\n');
    bajar('﻿' + t, nombre, 'text/csv;charset=utf-8');
  }
  function bajar(txt, nombre, tipo) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([txt], { type: tipo }));
    a.download = nombre;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.parentNode.removeChild(a); }, 1200);
  }

  /* ═════════════════════ arranque ═════════════════════ */
  var avisoModo = $('#avisoModo');
  if (avisoModo && A.remoto) {
    avisoModo.innerHTML = '<b>Modo compartido.</b> El avance de todas las dependencias se guarda en ' +
      'la hoja de cálculo institucional. Puede cerrar el navegador y continuar después, desde este ' +
      'equipo o desde otro: nada se pierde y la Gerencia Técnica ve el avance en tiempo real.';
    avisoModo.style.background = '#E3F5E9';
    avisoModo.style.borderColor = '#BCE3C9';
    avisoModo.style.color = '#0b6b2e';
  }
  if (sesion) abrirApp();
  else $('#usuario').focus();

})();
