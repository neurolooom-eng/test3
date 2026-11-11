/*
 * HTML renderer + deck controller (browser only).
 * Draws the primitives from src/slides.js as absolutely positioned elements on
 * a 1280 × 720 px canvas (13.333 × 7.5 in at 96 px/in), then scales to fit.
 */
(function () {
  'use strict';

  var cfg = window.AIDA_DECK_CONFIG, S = window.AIDA_SLIDES, calc = window.AIDA_CALC;
  var PX = 96, PT = 96 / 72;
  var state = S.initialState(cfg);
  var current = 0, slides = [], showNotes = false;

  function palette() { return cfg.theme.palettes[cfg.theme.activePalette] || cfg.theme.palettes.light; }
  function col(token) {
    if (!token) return 'transparent';
    var p = palette();
    return '#' + (p[token] || String(token).replace('#', ''));
  }

  /* ---------- theme → CSS variables ---------- */
  function applyTheme() {
    var p = palette(), r = document.documentElement.style;
    Object.keys(p).forEach(function (k) { r.setProperty('--c-' + k, '#' + p[k]); });
    r.setProperty('--font-heading', cfg.theme.fonts.headingStack);
    r.setProperty('--font-body', cfg.theme.fonts.bodyStack);
    document.documentElement.setAttribute('data-palette', cfg.theme.activePalette);
  }

  /* ---------- primitive renderers ---------- */
  function box(p) {
    var el = document.createElement('div');
    el.className = 'p p-' + p.type;
    el.style.left = p.x * PX + 'px';
    el.style.top = p.y * PX + 'px';
    el.style.width = p.w * PX + 'px';
    el.style.height = p.h * PX + 'px';
    if (p.fill) el.style.background = col(p.fill);
    if (p.line) el.style.border = (p.lineW || 1) * PT + 'px ' + (p.dash ? 'dashed' : 'solid') + ' ' + col(p.line);
    return el;
  }

  function renderText(p) {
    var el = box(p);
    el.style.fontFamily = p.font === 'heading' ? 'var(--font-heading)' : 'var(--font-body)';
    el.style.fontSize = p.size * PT + 'px';
    el.style.color = col(p.color);
    el.style.fontWeight = p.bold ? '700' : '400';
    if (p.italic) el.style.fontStyle = 'italic';
    el.style.textAlign = p.align;
    el.style.justifyContent = { top: 'flex-start', middle: 'center', bottom: 'flex-end' }[p.valign] || 'flex-start';
    if (p.margin) el.style.padding = p.margin * PT + 'px';
    if (p.role === 'title') el.setAttribute('role', 'heading');
    var inner = document.createElement('div');
    inner.className = 'p-text-inner';
    if (p.runs) {
      p.runs.forEach(function (r) {
        var sp = document.createElement('span');
        sp.textContent = r.text;
        if (r.bold !== undefined) sp.style.fontWeight = r.bold ? '700' : '400';
        if (r.italic) sp.style.fontStyle = 'italic';
        if (r.color) sp.style.color = col(r.color);
        if (r.size) sp.style.fontSize = r.size * PT + 'px';
        inner.appendChild(sp);
        if (r.breakLine) inner.appendChild(document.createElement('br'));
      });
    } else inner.textContent = p.text;
    el.appendChild(inner);
    return el;
  }

  function renderLine(p) {
    var pad = 8, w = Math.max(p.w * PX, 0), h = Math.max(p.h * PX, 0);
    var svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('class', 'p p-line');
    svg.setAttribute('width', w + 2 * pad);
    svg.setAttribute('height', h + 2 * pad);
    svg.style.left = p.x * PX - pad + 'px';
    svg.style.top = p.y * PX - pad + 'px';
    var c = col(p.color), sw = (p.width || 1) * PT;
    var l = document.createElementNS(svg.namespaceURI, 'line');
    var x2 = pad + w, y2 = pad + h;
    if (p.arrow) { var a = sw * 3.2; if (w >= h) x2 -= a; else y2 -= a; }
    l.setAttribute('x1', pad); l.setAttribute('y1', pad); l.setAttribute('x2', x2); l.setAttribute('y2', y2);
    l.setAttribute('stroke', c); l.setAttribute('stroke-width', sw);
    if (p.dash) l.setAttribute('stroke-dasharray', sw * 3 + ' ' + sw * 2);
    svg.appendChild(l);
    if (p.arrow) {
      var tri = document.createElementNS(svg.namespaceURI, 'polygon'), s = sw * 3.4, ex = pad + w, ey = pad + h;
      var pts = w >= h
        ? [[ex, ey], [ex - s, ey - s * 0.6], [ex - s, ey + s * 0.6]]
        : [[ex, ey], [ex - s * 0.6, ey - s], [ex + s * 0.6, ey - s]];
      tri.setAttribute('points', pts.map(function (q) { return q.join(','); }).join(' '));
      tri.setAttribute('fill', c);
      svg.appendChild(tri);
    }
    return svg;
  }

  function renderInputs(p) {
    var el = box(p);
    el.classList.add('inputs');
    var rowH = p.h / p.rows.length * PX;
    p.rows.forEach(function (r) {
      var row = document.createElement('label');
      row.className = 'inp-row';
      row.style.height = rowH + 'px';
      var lab = document.createElement('span');
      lab.className = 'inp-label';
      lab.textContent = r.label;
      row.appendChild(lab);
      var ctl;
      if (r.kind === 'select') {
        ctl = document.createElement('select');
        r.options.forEach(function (o) {
          var op = document.createElement('option');
          op.value = o.value; op.textContent = o.label;
          if (o.value === r.value) op.selected = true;
          ctl.appendChild(op);
        });
      } else {
        ctl = document.createElement('input');
        ctl.type = r.kind === 'text' ? 'text' : 'number';
        if (r.kind !== 'text') { ctl.min = '0'; ctl.step = 'any'; ctl.inputMode = 'decimal'; }
        ctl.value = r.kind === 'percent' ? Math.round(r.value * 1000) / 10 : r.value;
      }
      ctl.className = 'inp-ctl';
      ctl.setAttribute('aria-label', r.label);
      ctl.dataset.key = r.key;
      ctl.dataset.kind = r.kind;
      ctl.addEventListener('change', onInput);
      row.appendChild(ctl);
      el.appendChild(row);
    });
    return el;
  }

  function onInput(e) {
    var k = e.target.dataset.key, kind = e.target.dataset.kind, v = e.target.value;
    if (k === 'releaseType') { state.releaseId = v; state.overrides = {}; }
    else if (k === 'userScenario') { state.userId = v; delete state.overrides.aidaHoursPerDocument; }
    else if (kind === 'percent') state.overrides[k] = v === '' ? null : parseFloat(v) / 100;
    else if (kind === 'number') state.overrides[k] = v === '' ? null : parseFloat(v);
    else state.overrides[k] = v;
    renderAll();
    var again = document.querySelector('.slide.active [data-key="' + k + '"]');
    if (again) again.focus();
  }

  function renderPrimitive(p) {
    if (p.type === 'text') return renderText(p);
    if (p.type === 'line') return renderLine(p);
    if (p.type === 'inputs') return renderInputs(p);
    var el = box(p);
    if (p.type === 'ellipse') el.style.borderRadius = '50%';
    return el;
  }

  function renderChrome(slideEl, index) {
    var c = S.chrome(cfg);
    slideEl.appendChild(renderPrimitive(c.footerRule));
    slideEl.appendChild(renderPrimitive(c.footerText));
    var n = renderText(Object.assign({ type: 'text', text: String(index + 1), align: 'right', valign: 'middle', font: 'body' }, c.slideNumber));
    if (cfg.theme.footer.showSlideNumber) slideEl.appendChild(n);
    var lg = c.logo, el;
    if (lg.imagePath) {
      el = document.createElement('img');
      el.src = lg.imagePath; el.alt = 'Logo';
      el.className = 'p logo-img';
      Object.assign(el.style, { left: lg.x * PX + 'px', top: lg.y * PX + 'px', width: lg.w * PX + 'px', height: lg.h * PX + 'px' });
    } else {
      el = renderText({ type: 'text', x: lg.x, y: lg.y, w: lg.w, h: lg.h, text: lg.text, size: 11, bold: true, color: 'muted', align: 'center', valign: 'middle', line: 'neutral', lineW: 1, dash: lg.dashed, font: 'body' });
    }
    slideEl.appendChild(el);
  }

  /* ---------- deck ---------- */
  function renderAll() {
    applyTheme();
    slides = S.buildDeck(cfg, state, calc);
    var stage = document.getElementById('deck');
    stage.innerHTML = '';
    slides.forEach(function (sl, i) {
      var el = document.createElement('section');
      el.className = 'slide' + (i === current ? ' active' : '');
      el.setAttribute('aria-label', 'Slide ' + (i + 1) + ': ' + sl.title);
      el.dataset.slide = sl.id;
      el.style.background = col('background');
      sl.primitives.forEach(function (p) { el.appendChild(renderPrimitive(p)); });
      renderChrome(el, i);
      stage.appendChild(el);
    });
    document.title = cfg.brand.toolName + ' leadership deck';
    updateUi();
  }

  function updateUi() {
    document.querySelectorAll('.slide').forEach(function (el, i) { el.classList.toggle('active', i === current); });
    document.getElementById('counter').textContent = (current + 1) + ' / ' + slides.length;
    var sl = slides[current];
    document.getElementById('notes-meta').textContent = 'Slide ' + (current + 1) + ' · ' + sl.owner + ' · ' + sl.time;
    document.getElementById('notes-body').textContent = sl.notes.split('\n\n').slice(1).join('\n\n');
    document.body.classList.toggle('with-notes', showNotes);
    fit();
  }

  function fit() {
    var wrap = document.getElementById('viewport');
    var s = Math.min(wrap.clientWidth / 1280, wrap.clientHeight / 720);
    document.getElementById('deck').style.transform = 'scale(' + s + ')';
    document.getElementById('deck').style.width = '1280px';
    document.getElementById('scaler').style.width = 1280 * s + 'px';
    document.getElementById('scaler').style.height = 720 * s + 'px';
  }

  function go(d) { current = Math.max(0, Math.min(slides.length - 1, current + d)); updateUi(); }

  function exportPptx() {
    var btn = document.getElementById('btn-pptx');
    if (!window.PptxGenJS) { alert('PptxGenJS failed to load (vendor/pptxgenjs/pptxgen.bundle.js).'); return; }
    btn.disabled = true; btn.textContent = 'Exporting…';
    var pptx = window.AIDA_PPTX.build(window.PptxGenJS, cfg, state, calc);
    pptx.writeFile({ fileName: cfg.brand.toolName + '-leadership-deck.pptx' })
      .catch(function (err) { alert('Export failed: ' + err); })
      .then(function () { btn.disabled = false; btn.textContent = 'Export PPTX'; });
  }

  document.addEventListener('keydown', function (e) {
    var tag = (e.target.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'select') return;
    if (['ArrowRight', 'PageDown', ' '].indexOf(e.key) >= 0) { e.preventDefault(); go(1); }
    else if (['ArrowLeft', 'PageUp'].indexOf(e.key) >= 0) { e.preventDefault(); go(-1); }
    else if (e.key === 'Home') { current = 0; updateUi(); }
    else if (e.key === 'End') { current = slides.length - 1; updateUi(); }
    else if (e.key === 'n' || e.key === 'N') { showNotes = !showNotes; updateUi(); }
  });
  window.addEventListener('resize', fit);
  document.getElementById('btn-prev').onclick = function () { go(-1); };
  document.getElementById('btn-next').onclick = function () { go(1); };
  document.getElementById('btn-notes').onclick = function () { showNotes = !showNotes; updateUi(); };
  document.getElementById('btn-theme').onclick = function () {
    cfg.theme.activePalette = cfg.theme.activePalette === 'dark' ? 'light' : 'dark';
    renderAll();
  };
  document.getElementById('btn-pptx').onclick = exportPptx;
  document.getElementById('btn-pdf').onclick = function () { window.print(); };

  var m = /[?&]slide=(\d+)/.exec(location.search);
  if (m) current = Math.max(0, Math.min(6, parseInt(m[1], 10) - 1));
  var pal = /[?&]palette=(light|dark)/.exec(location.search);
  if (pal) cfg.theme.activePalette = pal[1];
  renderAll();
  window.AIDA_DECK = { state: state, rerender: renderAll, go: function (i) { current = i; updateUi(); } };
})();
