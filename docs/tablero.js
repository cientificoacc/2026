/* Tablero MIPG · FURAG — Agencia Catastral de Cundinamarca
   Sin dependencias externas: los gráficos se dibujan como SVG con las mismas
   variables CSS de la identidad, así el modo de alto contraste los arrastra. */
(function () {
  'use strict';

  var D = window.DATOS;
  var $ = function (s) { return document.querySelector(s); };
  var NS = 'http://www.w3.org/2000/svg';

  var POL = {}, DIM = {}, RES = {};
  D.politicas.forEach(function (p) { POL[p.id] = p; });
  D.dimensiones.forEach(function (d) { DIM[d.id] = d; });
  D.responsables.forEach(function (r) { RES[r.id] = r; });

  /* ── estados de respuesta y prioridades: color + etiqueta, nunca color solo ── */
  var ESTADO = [
    { k: 'bien',    et: 'Acreditada completa', v: 'var(--st-bien)' },
    { k: 'parcial', et: 'Avance parcial',      v: 'var(--st-media)' },
    { k: 'incump',  et: 'Incumplimiento',      v: 'var(--st-alta)' },
    { k: 'sinresp', et: 'Sin responder',       v: 'var(--st-critica)' }
  ];
  var PRIO = [
    { k: 'alta',       et: 'Alta',        v: 'var(--st-critica)', pill: 'p-critica' },
    { k: 'media_alta', et: 'Media-alta',  v: 'var(--st-alta)',    pill: 'p-alta' },
    { k: 'media',      et: 'Media',       v: 'var(--st-media)',   pill: 'p-media' }
  ];
  var ET_PRIO = {}; PRIO.forEach(function (p) { ET_PRIO[p.k] = p; });

  function estadoDe(p) {
    if (!p.brecha) return 'bien';
    if (p.cat === 'sin_responder') return 'sinresp';
    if (p.cat === 'incumplimiento') return 'incump';
    return 'parcial';
  }
  D.preguntas.forEach(function (p) { p.estado = estadoDe(p); });

  var fmt = function (n) { return n.toLocaleString('es-CO'); };
  var pct = function (a, b) { return b ? Math.round(a * 1000 / b) / 10 : 0; };
  var esc = function (s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  };

  /* ───────────────────────── estado del filtro ───────────────────────── */
  var filtro = '';
  var V = {};   // vista calculada

  function calcular() {
    V.preguntas = filtro ? D.preguntas.filter(function (p) { return p.res === filtro; }) : D.preguntas;
    V.acciones  = filtro ? D.acciones.filter(function (a) { return a.res === filtro; }) : D.acciones;
    V.brechas   = V.preguntas.filter(function (p) { return p.brecha; });
    V.cump = V.preguntas.reduce(function (s, p) { return s + (p.cump || 0); }, 0);
    V.tot  = V.preguntas.reduce(function (s, p) { return s + (p.tot || 0); }, 0);
    V.cortes = V.acciones.reduce(function (s, a) { return s + a.cortes.length; }, 0);
    V.alta = V.acciones.filter(function (a) { return a.prio === 'alta'; }).length;
    V.exigeEv = V.preguntas.filter(function (p) { return p.exigeEv; });
    V.conEv = V.exigeEv.filter(function (p) { return p.nEv > 0; });
    V.sinSoporte = V.acciones.filter(function (a) { return a.soporte; }).length;
  }

  /* ───────────────────────── utilidades SVG ───────────────────────── */
  function el(n, a) {
    var e = document.createElementNS(NS, n);
    for (var k in a) if (a[k] !== undefined && a[k] !== null) e.setAttribute(k, a[k]);
    return e;
  }
  function txt(x, y, s, cls, anchor) {
    var t = el('text', { x: x, y: y, class: cls || '', 'text-anchor': anchor || 'start' });
    t.textContent = s;
    return t;
  }
  function limpiar(svg, w, h) {
    while (svg.firstChild) svg.removeChild(svg.firstChild);
    svg.setAttribute('viewBox', '0 0 ' + w + ' ' + h);
    svg.setAttribute('preserveAspectRatio', 'xMinYMin meet');
    return svg;
  }
  var tip = $('#tip');
  function conTip(nodo, html) {
    nodo.addEventListener('mouseenter', function (e) {
      tip.innerHTML = html; tip.style.opacity = '1'; mover(e);
    });
    nodo.addEventListener('mousemove', mover);
    nodo.addEventListener('mouseleave', function () { tip.style.opacity = '0'; });
    function mover(e) {
      var x = e.clientX + 14, y = e.clientY + 16;
      var r = tip.getBoundingClientRect();
      if (x + r.width > innerWidth - 8) x = e.clientX - r.width - 14;
      if (y + r.height > innerHeight - 8) y = e.clientY - r.height - 16;
      tip.style.left = x + 'px'; tip.style.top = y + 'px';
    }
  }

  /* ¿pantalla angosta? define el ancho del lienzo y el canal de etiquetas */
  function angosto() { return window.innerWidth < 700; }
  function recortar(t, n) { return t.length > n ? t.slice(0, n - 1) + '…' : t; }

  /* Barras horizontales de una sola serie (magnitud). Etiqueta directa siempre. */
  function barrasH(svg, datos, opts) {
    opts = opts || {};
    var ang = angosto();
    var W = ang ? 400 : 640, fila = opts.fila || (ang ? 24 : 26), gap = 6;
    var izq = ang ? 128 : (opts.izq || 210), der = ang ? 46 : 58;
    var H = Math.max(40, datos.length * (fila + gap) + 14);
    limpiar(svg, W, H);
    var max = Math.max.apply(null, datos.map(function (d) { return d.v; }).concat([1]));
    var ancho = W - izq - der;
    datos.forEach(function (d, i) {
      var y = i * (fila + gap) + 6;
      var w = Math.max(d.v > 0 ? 3 : 0, Math.round(d.v / max * ancho));
      var g = el('g', { class: 'item' });
      g.appendChild(txt(izq - 8, y + fila * 0.68, recortar(d.et, ang ? 17 : 30), 'rot', 'end'));
      var r = el('rect', { x: izq, y: y, width: w, height: fila, rx: 3, class: 'marca-d',
                           fill: d.color || 'var(--serie)' });
      g.appendChild(r);
      g.appendChild(txt(izq + w + 8, y + fila * 0.68, fmt(d.v) + (opts.suf || ''), 'val'));
      var hit = el('rect', { x: 0, y: y - 3, width: W, height: fila + 6, class: 'hit' });
      g.appendChild(hit);
      conTip(g, '<b>' + esc(d.et) + '</b>' + '<span class="n">' + fmt(d.v) +
                (opts.tipSuf || '') + '</span>' + (d.extra ? '<br>' + esc(d.extra) : ''));
      svg.appendChild(g);
    });
    if (!datos.length) svg.appendChild(txt(10, 24, 'Sin datos para esta dependencia.', 'eje'));
  }

  /* Barras apiladas horizontales por categoría de estado. Separador de 2px. */
  function apiladasH(svg, filas, series, opts) {
    opts = opts || {};
    var ang = angosto();
    var W = ang ? 400 : 640, fila = ang ? 22 : 24, gap = 8;
    var izq = ang ? 128 : 210, der = ang ? 40 : 46;
    var H = Math.max(40, filas.length * (fila + gap) + 12);
    limpiar(svg, W, H);
    var max = Math.max.apply(null, filas.map(function (f) { return f.total; }).concat([1]));
    var ancho = W - izq - der;
    filas.forEach(function (f, i) {
      var y = i * (fila + gap) + 6, x = izq;
      var g = el('g', { class: 'item' });
      g.appendChild(txt(izq - 8, y + fila * 0.7, recortar(f.et, ang ? 17 : 30), 'rot', 'end'));
      series.forEach(function (s) {
        var v = f[s.k] || 0;
        if (!v) return;
        var w = Math.max(2, Math.round(v / max * ancho));
        var seg = el('g', { class: 'item' });
        seg.appendChild(el('rect', { x: x, y: y, width: w, height: fila, rx: 3,
                                     class: 'marca-d', fill: s.v }));
        conTip(seg, '<b>' + esc(f.et) + '</b>' + esc(s.et) + ': <span class="n">' +
                    fmt(v) + '</span> de ' + fmt(f.total));
        g.appendChild(seg);
        x += w + 2;                       /* separador de 2px entre segmentos */
      });
      g.appendChild(txt(x + 8, y + fila * 0.7, fmt(f.total), 'val'));
      svg.appendChild(g);
    });
    if (!filas.length) svg.appendChild(txt(10, 24, 'Sin datos para esta dependencia.', 'eje'));
  }

  /* Barras verticales apiladas (cronograma mensual). */
  function apiladasV(svg, cols, series) {
    var ang = angosto();
    var W = ang ? 380 : 620, H = ang ? 230 : 258, base = H - 40, top = 30, izq = ang ? 32 : 40;
    limpiar(svg, W, H);
    var max = Math.max.apply(null, cols.map(function (c) { return c.total; }).concat([1]));
    var paso = (W - izq - 14) / cols.length, bw = Math.min(ang ? 40 : 70, paso * 0.62);
    [0, 0.5, 1].forEach(function (f) {
      var y = base - f * (base - top);
      svg.appendChild(el('line', { x1: izq, x2: W - 8, y1: y, y2: y, class: 'rejilla' }));
      svg.appendChild(txt(izq - 8, y + 4, fmt(Math.round(max * f)), 'eje', 'end'));
    });
    cols.forEach(function (c, i) {
      var cx = izq + paso * i + paso / 2, y = base;
      var g = el('g', { class: 'item' });
      series.forEach(function (s) {
        var v = c[s.k] || 0;
        if (!v) return;
        var h = Math.max(2, Math.round(v / max * (base - top)));
        y -= h;
        var seg = el('g', { class: 'item' });
        seg.appendChild(el('rect', { x: cx - bw / 2, y: y, width: bw, height: h, rx: 3,
                                     class: 'marca-d', fill: s.v }));
        conTip(seg, '<b>' + esc(c.et) + '</b>' + esc(s.et) + ': <span class="n">' +
                    fmt(v) + '</span> acciones');
        g.appendChild(seg);
        y -= 2;                            /* separador de 2px */
      });
      g.appendChild(txt(cx, y - 7, fmt(c.total), 'val', 'middle'));
      g.appendChild(txt(cx, base + 17, ang ? c.et.slice(0, 3) : c.et, 'eje', 'middle'));
      svg.appendChild(g);
    });
    svg.appendChild(el('line', { x1: izq, x2: W - 8, y1: base, y2: base, class: 'rejilla' }));
  }

  function leyenda(cont, series) {
    cont.innerHTML = series.map(function (s) {
      return '<span><i style="background:' + s.v + '"></i>' + esc(s.et) + '</span>';
    }).join('');
  }

  /* ───────────────────────── pintado ───────────────────────── */
  function pintar() {
    calcular();
    var nombre = filtro ? RES[filtro].nombre : 'toda la Agencia';

    $('#ambito').innerHTML = 'Mostrando <b>' + esc(nombre) + '</b> · ' +
      fmt(V.preguntas.length) + ' preguntas y ' + fmt(V.acciones.length) + ' acciones';

    /* KPIs */
    var kpis = [
      [fmt(V.preguntas.length), 'preguntas del FURAG a cargo', false],
      [fmt(V.brechas.length), 'quedaron con brecha en 2025', V.brechas.length > 0],
      [pct(V.cump, V.tot) + ' %', 'prácticas acreditadas en 2025', false],
      [fmt(V.acciones.length), 'acciones para cerrar en 2026', false],
      [fmt(V.alta), 'de prioridad alta', V.alta > 0],
      [fmt(V.cortes), 'reportes mensuales por hacer', false]
    ];
    $('#kpis').innerHTML = kpis.map(function (k) {
      return '<div class="kpi"><div class="v' + (k[2] ? ' alerta' : '') + '">' + k[0] +
             '</div><div class="e">' + k[1] + '</div></div>';
    }).join('');

    /* ── acto 1 ── */
    var porDim = {};
    V.preguntas.forEach(function (p) {
      var d = (POL[p.pol] || {}).dim || '—';
      porDim[d] = (porDim[d] || 0) + 1;
    });
    var datosDim = D.dimensiones.map(function (d) {
      return { et: d.nombre, v: porDim[d.id] || 0, extra: d.proposito };
    }).filter(function (d) { return d.v > 0; }).sort(function (a, b) { return b.v - a.v; });
    barrasH($('#gDim'), datosDim, { tipSuf: ' preguntas' });
    $('#sub1').textContent = filtro
      ? 'Preguntas a cargo de ' + nombre + ', agrupadas por dimensión del MIPG.'
      : 'Número de preguntas del FURAG que corresponden a cada dimensión.';

    var polFilas = D.politicas.map(function (p) {
      var qs = V.preguntas.filter(function (q) { return q.pol === p.id; });
      return { p: p, n: qs.length, b: qs.filter(function (q) { return q.brecha; }).length };
    }).filter(function (f) { return f.n > 0; }).sort(function (a, b) { return b.n - a.n; });
    $('#tabPol tbody').innerHTML = polFilas.length ? polFilas.map(function (f) {
      return '<tr><td>' + esc(f.p.nombre) + '</td><td>' +
             esc((DIM[f.p.dim] || {}).nombre || '—') + '</td><td>' + esc(f.p.lider) +
             '</td><td class="n">' + fmt(f.n) + '</td><td class="n">' + fmt(f.b) + '</td></tr>';
    }).join('') : '<tr><td colspan="6" class="vacio">Esta dependencia no tiene políticas asignadas.</td></tr>';

    /* ── acto 2 ── */
    barrasH($('#gPol'), polFilas.slice(0, 10).map(function (f) {
      return { et: f.p.nombre, v: f.n, extra: f.b + ' con brecha' };
    }), { tipSuf: ' preguntas' });

    var porTipo = {};
    V.preguntas.forEach(function (p) { porTipo[p.tipo] = (porTipo[p.tipo] || 0) + 1; });
    barrasH($('#gTipo'), Object.keys(porTipo).map(function (t) {
      return { et: t, v: porTipo[t] };
    }).sort(function (a, b) { return b.v - a.v; }), { izq: 178, tipSuf: ' preguntas' });

    /* ── acto 3 ── */
    leyenda($('#legEstado'), ESTADO);
    var filasEstado = polFilas.map(function (f) {
      var o = { et: f.p.nombre, total: f.n };
      ESTADO.forEach(function (s) { o[s.k] = 0; });
      V.preguntas.forEach(function (q) { if (q.pol === f.p.id) o[q.estado]++; });
      return o;
    });
    apiladasH($('#gEstado'), filasEstado, ESTADO);
    $('#sub3').textContent = 'Las ' + fmt(V.preguntas.length) + ' preguntas de ' + nombre +
      ', clasificadas según lo que se acreditó ante el DAFP.';

    var sinEv = V.exigeEv.length - V.conEv.length;
    barrasH($('#gEvid'), [
      { et: 'Con soporte aportado', v: V.conEv.length, color: 'var(--st-bien)' },
      { et: 'Sin soporte', v: sinEv, color: 'var(--st-critica)' },
      { et: 'No exige soporte', v: V.preguntas.length - V.exigeEv.length,
        color: 'var(--serie-suave)' }
    ], { izq: 178, tipSuf: ' preguntas' });

    barrasH($('#gPrio'), PRIO.map(function (p) {
      return { et: 'Prioridad ' + p.et.toLowerCase(),
               v: V.brechas.filter(function (b) { return b.prio === p.k; }).length,
               color: p.v };
    }), { izq: 178, tipSuf: ' brechas' });

    tablaBrechas();

    /* ── acto 4 ── */
    leyenda($('#legPrio'), PRIO.map(function (p) {
      return { k: p.k, et: 'Prioridad ' + p.et.toLowerCase(), v: p.v };
    }));
    var cols = D.meses.map(function (m) {
      var o = { et: m.nombre, total: 0 };
      PRIO.forEach(function (p) { o[p.k] = 0; });
      V.acciones.forEach(function (a) {
        if (a.cortes.some(function (c) { return c.m === m.n; })) { o[a.prio]++; o.total++; }
      });
      return o;
    });
    apiladasV($('#gCrono'), cols, PRIO);

    tablaAcciones();
  }

  /* ───────────────────────── tablas ───────────────────────── */
  var ORD = { alta: 0, media_alta: 1, media: 2 };
  function dep(id) { return (RES[id] || {}).nombre || '—'; }

  function tablaBrechas() {
    var q = ($('#bBrecha').value || '').trim().toLowerCase();
    var filas = V.brechas.slice().sort(function (a, b) {
      return (ORD[a.prio] - ORD[b.prio]) || a.cod.localeCompare(b.cod);
    });
    if (q) filas = filas.filter(function (f) {
      return (f.cod + ' ' + f.enun + ' ' + ((POL[f.pol] || {}).nombre || '') + ' ' +
              dep(f.res) + ' ' + f.falta).toLowerCase().indexOf(q) >= 0;
    });
    $('#cBrecha').textContent = fmt(filas.length) + ' de ' + fmt(V.brechas.length) + ' brechas';
    $('#tabBrecha tbody').innerHTML = filas.length ? filas.slice(0, 300).map(function (f) {
      var p = ET_PRIO[f.prio] || ET_PRIO.media;
      return '<tr><td class="cod">' + esc(f.cod) + '</td>' +
        '<td><span class="pill ' + p.pill + '">' + esc(p.et) + '</span></td>' +
        '<td class="dep">' + esc(dep(f.res)) + '</td>' +
        '<td><span class="clamp" title="' + esc(f.enun) + '">' + esc(f.enun) + '</span>' +
          '<span class="sub-cel">' + esc((POL[f.pol] || {}).nombre || '') + '</span></td>' +
        '<td class="n">' + fmt(f.cump) + ' de ' + fmt(f.tot) + '</td>' +
        '<td><span class="clamp" title="' + esc(f.falta || '') + '">' +
          esc(f.falta || '—') + '</span></td></tr>';
    }).join('') + (filas.length > 300
        ? '<tr><td colspan="6" class="vacio">Se muestran las primeras 300. Afine la búsqueda.</td></tr>' : '')
      : '<tr><td colspan="6" class="vacio">Sin brechas con ese criterio. ' +
        (V.brechas.length ? '' : 'Esta dependencia no tiene brechas registradas.') + '</td></tr>';
  }

  function tablaAcciones() {
    var q = ($('#bAccion').value || '').trim().toLowerCase();
    var filas = V.acciones.slice().sort(function (a, b) {
      return (ORD[a.prio] - ORD[b.prio]) || a.id.localeCompare(b.id);
    });
    if (q) filas = filas.filter(function (f) {
      return (f.id + ' ' + f.desc + ' ' + f.cod + ' ' + ((POL[f.pol] || {}).nombre || '') +
              ' ' + dep(f.res)).toLowerCase().indexOf(q) >= 0;
    });
    $('#cAccion').textContent = fmt(filas.length) + ' de ' + fmt(V.acciones.length) + ' acciones';
    $('#tabAccion tbody').innerHTML = filas.length ? filas.map(function (f) {
      var p = ET_PRIO[f.prio] || ET_PRIO.media;
      var meses = f.cortes.map(function (c) {
        return (D.meses.filter(function (m) { return m.n === c.m; })[0] || {}).nombre || '';
      });
      return '<tr><td class="cod">' + esc(f.id) +
          (f.soporte ? '<span class="sub-cel" style="color:var(--st-alta)">exige soporte</span>' : '') + '</td>' +
        '<td><span class="pill ' + p.pill + '">' + esc(p.et) + '</span></td>' +
        '<td class="dep">' + esc(dep(f.res)) + '</td>' +
        '<td><span class="clamp" title="' + esc(f.desc) + '">' + esc(f.desc) + '</span></td>' +
        '<td class="cod">' + esc(f.cod || '—') + '</td>' +
        '<td class="n" title="' + esc(meses.join(', ')) + '">' + f.cortes.length + '</td></tr>';
    }).join('')
      : '<tr><td colspan="6" class="vacio">Sin acciones con ese criterio.</td></tr>';
  }

  /* ───────────────────────── arranque ───────────────────────── */
  /* Solo se listan las dependencias con algo que mostrar. Gerencia General no tiene
     preguntas ni acciones asignadas: al elegirla el tablero quedaba en blanco. */
  var sel = $('#selRes');
  function conDatos(r) {
    return D.preguntas.some(function (p) { return p.res === r.id; }) ||
           D.acciones.some(function (a) { return a.res === r.id; });
  }
  D.responsables.filter(conDatos)
    .sort(function (a, b) { return a.nombre.localeCompare(b.nombre); })
    .forEach(function (r) {
      var o = document.createElement('option');
      o.value = r.id; o.textContent = r.nombre;
      sel.appendChild(o);
    });

  var hash = new URLSearchParams(location.search).get('area');
  if (hash && RES[hash] && conDatos(RES[hash])) { filtro = hash; sel.value = hash; }

  sel.addEventListener('change', function () {
    filtro = sel.value;
    var u = new URL(location.href);
    if (filtro) u.searchParams.set('area', filtro); else u.searchParams.delete('area');
    history.replaceState(null, '', u);
    pintar();
  });
  $('#bBrecha').addEventListener('input', tablaBrechas);
  $('#bAccion').addEventListener('input', tablaAcciones);

  /* El tablero abre y permanece en modo claro; los botones quedan ocultos por CSS
     pero se dejan cableados por si el comité pide devolverlos. */
  document.documentElement.setAttribute('data-theme', 'light');
  var bT = $('#btnTema'), bI = $('#btnImprimir');
  if (bT) bT.addEventListener('click', function () {
    var oscuro = document.documentElement.getAttribute('data-theme') === 'dark';
    document.documentElement.setAttribute('data-theme', oscuro ? 'light' : 'dark');
  });
  if (bI) bI.addEventListener('click', function () { window.print(); });

  $('#fCorte').textContent = D.corte;
  $('#fNota').innerHTML = 'Diagnóstico de la vigencia ' + D.vigencia.diagnostico +
    ' · cierre de brechas durante ' + D.vigencia.cierre +
    ' · reporte al DAFP en ' + D.vigencia.reporte + '.';

  var anchoPrev = window.innerWidth;
  function alRedimensionar() {
    clearTimeout(window._rz);
    window._rz = setTimeout(function () {
      /* solo se repinta si cruzó el umbral o cambió bastante el ancho */
      if (Math.abs(window.innerWidth - anchoPrev) > 40 ||
          (window.innerWidth < 700) !== (anchoPrev < 700)) {
        anchoPrev = window.innerWidth;
        pintar();
      }
    }, 200);
  }
  addEventListener('resize', alRedimensionar);
  addEventListener('orientationchange', alRedimensionar);

  pintar();
})();
