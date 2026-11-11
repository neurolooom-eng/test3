/*
 * Slide model — one source of truth for both renderers.
 *
 * Each slide is a list of simple primitives positioned in INCHES on a
 * 13.333 × 7.5 in (16:9) canvas. The HTML renderer (src/render-html.js) and
 * the PowerPoint exporter (src/pptx-builder.js) both draw from this list, so
 * every diagram exports as native, editable PowerPoint shapes and text.
 *
 * Design rule: slides carry a headline and one visual (about 30 words or
 * fewer). Everything else lives in the speaker notes (notesFor below).
 *
 * Primitive types
 *   rect / ellipse : x, y, w, h, fill, line, lineW (pt), dash, radius (in)
 *   line           : x, y, w, h (w or h = 0), color, width (pt), dash, arrow, arrowStart
 *   text           : x, y, w, h, text | runs[], size (pt), color, bold, italic,
 *                    align, valign, font, fill, line, radius, role, footnote
 *   inputs         : savings assumptions; `overlay: true` = hidden HTML panel, not in PPTX
 *   toggle         : HTML-only button that shows/hides an overlay (skipped in PPTX)
 * Colours are palette token names (see config.theme.palettes).
 */
(function (root) {
  'use strict';

  var W = 13.333, H = 7.5, M = 0.6, CW = W - 2 * M;
  var LAYOUT = { width: W, height: H, margin: M, contentWidth: CW, footerY: 6.97, footerRuleY: 6.9 };
  var R = 0.1; // corner radius (overridden by theme.diagramStyle.cardRadius)

  function makeTpl(cfg) {
    var p = cfg.presenters;
    return function (s) {
      if (typeof s !== 'string') return s;
      return s.replace(/\{tool\}/g, cfg.brand.toolName)
        .replace(/\{presenterOne\}/g, p.presenterOne)
        .replace(/\{presenterTwo\}/g, p.presenterTwo);
    };
  }

  function ownerName(cfg, owner) {
    if (owner === 'both') return cfg.presenters.presenterOne + ' & ' + cfg.presenters.presenterTwo;
    return cfg.presenters[owner] || owner;
  }

  /* ---------- primitive helpers ---------- */
  function rect(x, y, w, h, o) { return Object.assign({ type: 'rect', x: x, y: y, w: w, h: h }, o || {}); }
  function ellipse(x, y, w, h, o) { return Object.assign({ type: 'ellipse', x: x, y: y, w: w, h: h }, o || {}); }
  function line(x, y, w, h, o) { return Object.assign({ type: 'line', x: x, y: y, w: w, h: h, color: 'neutral', width: 1.75 }, o || {}); }
  function text(x, y, w, h, t, o) {
    var p = Object.assign({ type: 'text', x: x, y: y, w: w, h: h, size: 18, color: 'text', font: 'body', align: 'left', valign: 'top' }, o || {});
    if (Array.isArray(t)) p.runs = t; else p.text = t;
    return p;
  }
  function caps(s) { return String(s).toUpperCase(); }
  function pillW(label, size) { return Math.max(0.9, String(label).length * (size || 11) * 0.0095 + 0.4); }
  function footnote(t) { return text(M, 6.56, CW, 0.26, t, { size: 9, italic: true, color: 'muted', valign: 'middle', footnote: true }); }

  /* Status pill: current = filled blue, future = dashed outline, outcome = solid outline.
     Each state differs by fill/border style AND its text — never colour alone. */
  function statusPill(x, y, w, label, kind) {
    var base = { size: 10.5, bold: true, align: 'center', valign: 'middle', radius: 0.16 };
    if (kind === 'current') return text(x, y, w, 0.32, caps(label), Object.assign(base, { fill: 'accent', color: 'onAccent' }));
    if (kind === 'future') return text(x, y, w, 0.32, caps(label), Object.assign(base, { fill: 'background', line: 'secondary', lineW: 1.25, dash: 'dash', color: 'secondary' }));
    return text(x, y, w, 0.32, caps(label), Object.assign(base, { fill: 'background', line: 'accent', lineW: 1.25, color: 'accent' }));
  }
  function numberDot(x, y, d, n, fill, color) {
    return [
      ellipse(x, y, d, d, { fill: fill || 'accent' }),
      text(x, y, d, d, String(n), { size: Math.round(d * 32), bold: true, color: color || 'onAccent', align: 'center', valign: 'middle' })
    ];
  }
  function livePill(x, y) {
    return text(x, y, 1.4, 0.36, caps('Live demo'), { size: 11, bold: true, fill: 'panel', color: 'onPanel', align: 'center', valign: 'middle', radius: 0.18 });
  }

  /* Slide header: section pill, presenter, one-line headline. */
  function header(cfg, tpl, idx, section, headlineRuns, emphasis) {
    var s = [], label = caps(tpl(section));
    s.push(text(M, 0.42, pillW(label), 0.32, label, { size: 10.5, bold: true, color: 'accent', fill: 'accentSoft', align: 'center', valign: 'middle', radius: 0.16 }));
    s.push(text(6.4, 0.42, 4.8, 0.32, ownerName(cfg, cfg.agenda[idx].owner), { size: 11, color: 'muted', align: 'right', valign: 'middle' }));
    s.push(text(M, 0.92, CW, 0.82, headlineRuns.map(function (r) {
      return r.emphasis ? { text: tpl(r.text), bold: true, color: emphasis || 'accent' } : { text: tpl(r.text) };
    }), { size: 34, font: 'heading', color: 'text', valign: 'middle', role: 'title' }));
    return s;
  }

  /* ---------- slides ---------- */

  function slideOpening(cfg, tpl) {
    var o = cfg.opening, s = [], px = 8.35;
    s.push(text(M, 1.55, 7.4, 0.34, caps(tpl(cfg.brand.deckSubtitle)), { size: 12, bold: true, color: 'accent', valign: 'middle' }));
    s.push(text(M, 1.88, 7.4, 1.68, cfg.brand.toolName, { size: 100, bold: true, color: 'accent', font: 'heading', valign: 'middle' }));
    s.push(text(M, 3.55, 7.4, 0.7, tpl(cfg.brand.deckTitle), { size: 32, bold: true, color: 'text', font: 'heading', valign: 'middle', role: 'title' }));
    s.push(rect(M, 4.45, 0.9, 0.08, { fill: 'warning' }));
    s.push(text(M, 4.7, 7.3, 0.55, tpl(o.keyMessage), { size: 22, color: 'secondary', valign: 'middle' }));
    s.push(text(M, 6.2, 7.3, 0.36, cfg.presenters.presenterOne + '  ·  ' + cfg.presenters.presenterTwo + '  ·  ' + cfg.brand.roleLabel,
      { size: 13, bold: true, color: 'muted', valign: 'middle' }));

    // Right — three review modes as a vertical flow on navy
    s.push(rect(px, 0, W - px, H, { fill: 'panel' }));
    var cx = px + 0.75, d = 0.9, ys = [2.2, 3.55, 4.9];
    s.push(text(cx, 1.45, 4, 0.34, caps(tpl(o.modesLabel)), { size: 12, bold: true, color: 'onPanelMuted', valign: 'middle' }));
    s.push(line(cx + d / 2, ys[0] + d, 0, ys[2] - ys[0] - d, { color: 'onPanelMuted', width: 1.5 }));
    o.operatingModes.forEach(function (m, i) {
      s.push.apply(s, numberDot(cx, ys[i], d, i + 1));
      s.push(text(cx + 1.2, ys[i], 2.6, d, tpl(m.title), { size: 30, bold: true, color: 'onPanel', valign: 'middle', font: 'heading' }));
    });
    return s;
  }

  function slideUc1Problem(cfg, tpl) {
    var u = cfg.useCaseOne, s = header(cfg, tpl, 1, u.sectionProblem, u.headlineProblem, 'warning');
    var n = u.artefacts.length, bw = 2.15, gap = (CW - n * bw) / (n - 1), by = 2.35, bh = 1.15, midY = by + bh / 2;
    var gaps = {};
    u.issueLabels.forEach(function (g) { gaps[g.afterArtefact] = g; });
    function bx(i) { return M + i * (bw + gap); }
    for (var i = 0; i < n; i++) {
      s.push(text(bx(i), by, bw, bh, tpl(u.artefacts[i]), {
        size: 18, bold: true, align: 'center', valign: 'middle', fill: 'background',
        line: i === 0 ? 'accent' : 'accentLine', lineW: i === 0 ? 2.25 : 1.5, radius: R, margin: 5
      }));
      if (i < n - 1) {
        var lx = bx(i) + bw, g = gaps[i];
        if (g) {
          s.push(line(lx, midY, gap, 0, { color: 'warning', width: 2, dash: 'dash' }));
          var cxm = lx + gap / 2, dd = 0.36;
          s.push(ellipse(cxm - dd / 2, midY - dd / 2, dd, dd, { fill: 'warning' }));
          s.push(text(cxm - dd / 2, midY - dd / 2, dd, dd, '!', { size: 15, bold: true, color: 'onAccent', align: 'center', valign: 'middle' }));
          s.push(text(cxm - 1.0, by + bh + 0.15, 2.0, 0.36, tpl(g.text), { size: 14, bold: true, color: 'warning', align: 'center', valign: 'middle' }));
        } else {
          s.push(line(lx + 0.04, midY, gap - 0.08, 0, { color: 'accent', width: 2, arrow: true }));
        }
      }
    }
    // Red return loop with the cycle count as the focal number
    var x0 = bx(0) + bw / 2, x4 = bx(n - 1) + bw / 2, ly = 5.15;
    s.push(line(x4, by + bh, 0, ly - by - bh, { color: 'warning', width: 2.25 }));
    s.push(line(x0, ly, x4 - x0, 0, { color: 'warning', width: 2.25 }));
    s.push(line(x0, by + bh + 0.02, 0, ly - by - bh - 0.02, { color: 'warning', width: 2.25, arrowStart: true }));
    s.push(text(W / 2 - 1.9, ly - 0.75, 3.8, 1.5, [
      { text: tpl(u.reviewCycleRange) + '×', size: 54, bold: true, breakLine: true },
      { text: tpl(u.loopLabel), size: 16, bold: true }
    ], { color: 'warning', fill: 'background', align: 'center', valign: 'middle', font: 'heading' }));
    return s;
  }

  function slideUc1Approach(cfg, tpl) {
    var u = cfg.useCaseOne, s = header(cfg, tpl, 2, u.sectionApproach, u.headlineApproach);
    var lh = 1.42, lys = [2.1, 3.75], sx = 3.35, sw = 2.55, sg = (W - M - 0.3 - sx - 3 * sw) / 2;
    u.lanes.forEach(function (lane, li) {
      var y = lys[li], fut = lane.future;
      s.push(rect(M, y, CW, lh, fut ? { fill: 'background', line: 'accentLine', lineW: 1.25, dash: 'dash', radius: R } : { fill: 'surface', radius: R }));
      s.push(text(M + 0.35, y + 0.28, 2.4, 0.5, caps(tpl(lane.label)), { size: 24, bold: true, color: fut ? 'secondary' : 'accent', valign: 'middle', font: 'heading' }));
      s.push(statusPill(M + 0.35, y + 0.86, 1.45, tpl(lane.tag), fut ? 'future' : 'current'));
      lane.steps.forEach(function (st, k) {
        var x = sx + k * (sw + sg), hl = !!st.highlight;
        s.push(text(x, y + 0.26, sw, lh - 0.52, tpl(st.text), hl
          ? { size: 19, bold: true, fill: 'accent', color: 'onAccent', align: 'center', valign: 'middle', radius: R }
          : { size: 19, bold: true, fill: 'background', line: fut ? 'secondary' : 'neutral', lineW: 1, dash: fut, align: 'center', valign: 'middle', radius: R }));
        if (k < lane.steps.length - 1) s.push(line(x + sw + 0.08, y + lh / 2, sg - 0.16, 0, { color: 'accent', width: 2, arrow: true }));
      });
    });
    // Target cycles (explicitly a target) + demo cue
    var ky = 5.5;
    s.push(text(M, ky, 3, 0.3, caps(tpl(u.targetLabel)), { size: 11, bold: true, color: 'muted', valign: 'middle' }));
    s.push(text(M, ky + 0.3, 6, 0.75, [
      { text: tpl(u.reviewCycleRange) + '  →  ', color: 'muted' },
      { text: tpl(u.targetReviewCycles), color: 'accent' }
    ], { size: 36, bold: true, font: 'heading', valign: 'middle' }));
    s.push(livePill(W - M - 1.4, ky + 0.5));
    return s;
  }

  function slideUc2Problem(cfg, tpl) {
    var u = cfg.useCaseTwo, s = header(cfg, tpl, 3, u.sectionProblem, u.headlineProblem);
    s.push(statusPill(M, 2.08, 1.45, tpl(u.currentUseTag), 'current'));
    s.push(text(W - M - 4, 2.08, 4, 0.32, tpl(u.changeArrowLabel), { size: 13, bold: true, italic: true, color: 'warning', align: 'right', valign: 'middle' }));
    s.push(line(M, 2.6, CW, 0, { color: 'warning', width: 1.5, arrow: true }));

    var n = u.lifecycleLabels.length, g = 0.14, pw = (CW - (n - 1) * g) / n, py = 2.85, ph = 0.78;
    var hl = {};
    u.checkpointLabels.forEach(function (c, i) { hl[c.stage] = i + 1; });
    function px(i) { return M + i * (pw + g); }
    u.lifecycleLabels.forEach(function (lab, i) {
      s.push(text(px(i), py, pw, ph, tpl(lab), hl[i]
        ? { size: 18, bold: true, fill: 'accent', color: 'onAccent', align: 'center', valign: 'middle', radius: 0.39 }
        : { size: 17, fill: 'surface', color: 'muted', align: 'center', valign: 'middle', radius: 0.39 }));
    });
    var cy = 4.15, ch = 2.15, cw = 2 * pw + g;
    u.checkpointLabels.forEach(function (c, i) {
      var x = px(c.stage);
      s.push(line(x + pw / 2, py + ph, 0, cy - py - ph, { color: 'accent', width: 2 }));
      s.push(rect(x, cy, cw, ch, { fill: 'background', line: 'accentLine', lineW: 1.5, radius: R }));
      s.push.apply(s, numberDot(x + 0.35, cy + ch / 2 - 0.35, 0.7, i + 1));
      s.push(text(x + 1.3, cy, cw - 1.5, ch, tpl(c.question), { size: 30, bold: true, font: 'heading', valign: 'middle' }));
    });
    if (u.codeReviewStatus && u.codeReviewStatus.showOnSlide) s.push(footnote('Scope note: ' + tpl(u.codeReviewStatus.text)));
    return s;
  }

  function slideUc2Approach(cfg, tpl) {
    var u = cfg.useCaseTwo, s = header(cfg, tpl, 4, u.sectionApproach, u.headlineApproach);
    var top = 2.2, chipW = 2.45, chipH = 0.62, hubX = 3.75, hubW = 4.7, hubY = top, hubH = 3.2;
    u.sources.forEach(function (src, k) {
      var y = top + 0.1 + k * 0.82;
      s.push(text(M, y, chipW, chipH, tpl(src), { size: 16, bold: true, fill: 'surface', color: 'secondary', align: 'center', valign: 'middle', radius: R }));
      s.push(line(M + chipW + 0.06, y + chipH / 2, hubX - M - chipW - 0.12, 0, { color: 'accentLine', width: 1.75, arrow: true }));
    });
    s.push(rect(hubX, hubY, hubW, hubH, { fill: 'panel', radius: 0.16 }));
    s.push(text(hubX, hubY + 0.35, hubW, 1.0, tpl(u.hubTitle), { size: 48, bold: true, color: 'onPanel', align: 'center', valign: 'middle', font: 'heading' }));
    var cg = 0.15, cw2 = (hubW - 0.6 - 2 * cg) / 3;
    u.reusableSetup.forEach(function (m, k) {
      s.push(text(hubX + 0.3 + k * (cw2 + cg), hubY + 1.75, cw2, 0.62, tpl(m), { size: 17, bold: true, fill: 'accent', color: 'onAccent', align: 'center', valign: 'middle', radius: 0.31 }));
    });
    var ox = 9.25, ow = W - M - ox, oh = 1.25, oys = [hubY + 0.2, hubY + hubH - oh - 0.2];
    u.checkpointLabels.forEach(function (c, i) {
      var y = oys[i];
      s.push(line(hubX + hubW + 0.06, y + oh / 2, ox - hubX - hubW - 0.12, 0, { color: 'accent', width: 2, arrow: true }));
      s.push(rect(ox, y, ow, oh, { fill: 'background', line: 'accentLine', lineW: 1.5, radius: R }));
      s.push.apply(s, numberDot(ox + 0.28, y + oh / 2 - 0.27, 0.54, i + 1));
      s.push(text(ox + 1.0, y, ow - 1.1, oh, tpl(c.title), { size: 21, bold: true, font: 'heading', valign: 'middle' }));
    });
    s.push(livePill(W - M - 1.4, 5.95));
    return s;
  }

  function slideValue(cfg, tpl) {
    var v = cfg.value, s = header(cfg, tpl, 5, v.section, v.headline);
    var top = 2.1, fx = M, fw = 3.7, ax = fx + fw, aw = 0.9, tx = ax + aw, tw = 5.1, pxx = tx + tw + 0.3, pw = W - M - pxx;
    v.shifts.forEach(function (r, i) {
      var y = top + i * 1.18, h = 0.95;
      s.push(text(fx, y, fw, h, tpl(r.from), { size: 22, color: 'muted', fill: 'surface', align: 'center', valign: 'middle', radius: R }));
      s.push(line(ax + 0.15, y + h / 2, aw - 0.3, 0, { color: 'accent', width: 2.5, arrow: true }));
      s.push(text(tx, y, tw, h, tpl(r.to), r.kind === 'future'
        ? { size: 26, bold: true, color: 'text', fill: 'background', line: 'secondary', lineW: 1.5, dash: 'dash', align: 'center', valign: 'middle', radius: R }
        : { size: 26, bold: true, color: 'text', fill: 'accentSoft', line: 'accent', lineW: 1.5, align: 'center', valign: 'middle', radius: R }));
      s.push(statusPill(pxx, y + h / 2 - 0.16, pw, tpl(r.tag), r.kind));
    });
    s.push(rect(M, 5.68, CW, 0.78, { fill: 'panel', radius: R }));
    s.push(rect(M + 0.3, 5.88, 0.07, 0.38, { fill: 'warning' }));
    s.push(text(M + 0.6, 5.68, CW - 0.9, 0.78, tpl(v.bottomLine), { size: 22, bold: true, color: 'onPanel', valign: 'middle' }));
    s.push(footnote(tpl(v.caveatShort)));
    return s;
  }

  function assumptionRows(sv, tpl, state, inp, fmt) {
    var L = sv.labels;
    return [
      { key: 'releaseType', label: tpl(L.releaseType), kind: 'select', value: state.releaseId, display: inp.releaseType,
        options: sv.releaseScenarios.map(function (x) { return { value: x.id, label: tpl(x.releaseType) }; }) },
      { key: 'userScenario', label: tpl(L.userScenario), kind: 'select', value: state.userId, display: inp.userScenario,
        options: sv.userScenarios.map(function (x) { return { value: x.id, label: tpl(x.label) }; }) },
      { key: 'eligibleDocumentsPerRelease', label: tpl(L.eligibleDocumentsPerRelease), kind: 'number', value: inp.eligibleDocumentsPerRelease, display: fmt(+inp.eligibleDocumentsPerRelease) },
      { key: 'adoptionRate', label: tpl(L.adoptionRate), kind: 'percent', value: inp.adoptionRate, display: fmt(inp.adoptionRate * 100, 0) + '%' },
      { key: 'manualHoursPerDocument', label: tpl(L.manualHoursPerDocument), kind: 'number', value: inp.manualHoursPerDocument, display: fmt(+inp.manualHoursPerDocument) },
      { key: 'aidaHoursPerDocument', label: tpl(L.aidaHoursPerDocument), kind: 'number', value: inp.aidaHoursPerDocument, display: fmt(+inp.aidaHoursPerDocument) },
      { key: 'releasesPerYear', label: tpl(L.releasesPerYear), kind: 'number', value: inp.releasesPerYear, display: fmt(+inp.releasesPerYear) },
      { key: 'baselineReviewCycles', label: tpl(L.baselineReviewCycles), kind: 'text', value: inp.baselineReviewCycles, display: tpl(String(inp.baselineReviewCycles)) },
      { key: 'aidaTargetReviewCycles', label: tpl(L.aidaTargetReviewCycles), kind: 'text', value: inp.aidaTargetReviewCycles, display: tpl(String(inp.aidaTargetReviewCycles)) }
    ];
  }

  function slideSavings(cfg, tpl, state, calc) {
    var sv = cfg.savings, L = sv.labels, s = header(cfg, tpl, 6, sv.section, sv.headline);
    var inp = calc.resolveInputs(sv, state.releaseId, state.userId, state.overrides);
    var r = calc.calculateCapacity(inp), fmt = calc.fmt;
    var rows = assumptionRows(sv, tpl, state, inp, fmt);

    // Hero number
    var top = 2.1, hw = 5.7, hh = 2.75;
    s.push(rect(M, top, hw, hh, { fill: 'panel', radius: 0.14 }));
    s.push(text(M + 0.45, top + 0.3, hw - 0.9, 1.5, fmt(r.qeCapacityReleasedHours) + ' h', { size: 88, bold: true, color: 'onPanel', font: 'heading', valign: 'middle' }));
    s.push(text(M + 0.45, top + 1.85, hw - 0.9, 0.5, tpl(L.resultReleased), { size: 17, color: 'onPanelMuted', valign: 'middle' }));

    // Supporting numbers
    var rx = M + hw + 0.35, rw = W - M - rx, kw = (rw - 0.25) / 2;
    [[r.qeCapacityReleasedPercent === null ? '—' : fmt(r.qeCapacityReleasedPercent, 0) + '%', tpl(L.resultPercent)],
     [r.annualQeCapacityReleasedHours === null ? '—' : fmt(r.annualQeCapacityReleasedHours) + ' h', tpl(L.resultAnnual)]].forEach(function (k, i) {
      var x = rx + i * (kw + 0.25);
      s.push(rect(x, top, kw, 1.3, { fill: 'accentSoft', radius: R }));
      s.push(text(x + 0.25, top + 0.12, kw - 0.5, 0.75, k[0], { size: 38, bold: true, color: 'accent', font: 'heading', valign: 'middle' }));
      s.push(text(x + 0.25, top + 0.85, kw - 0.5, 0.36, k[1], { size: 13, color: 'secondary', valign: 'middle' }));
    });

    // Before → after bars
    var cur = r.currentQeHoursPerRelease, aid = r.aidaAssistedQeHoursPerRelease;
    var max = Math.max(cur || 0, aid || 0), bx = rx + 1.45, barMax = rw - 1.45 - 1.05;
    function barW(v) { return v && max > 0 ? Math.max(0.05, barMax * Math.max(v, 0) / max) : 0.05; }
    [[tpl(L.currentBar), cur, 'secondary'], [tpl(L.aidaBar), aid, 'accent']].forEach(function (b, i) {
      var y = top + 1.6 + i * 0.62, w = barW(b[1]);
      s.push(text(rx, y, 1.4, 0.48, b[0], { size: 14, bold: true, valign: 'middle' }));
      s.push(rect(bx, y, w, 0.48, { fill: b[2], radius: 0.06 }));
      s.push(text(bx + w + 0.1, y, 1.0, 0.48, fmt(b[1]) + ' h', { size: 15, bold: true, valign: 'middle' }));
    });

    // Reinvestment
    var ry = 5.2;
    s.push(text(M, ry, 1.9, 0.56, caps(tpl(sv.reinvestmentLead)) + '  →', { size: 12, bold: true, color: 'accent', valign: 'middle' }));
    var ix = M + 1.9, iw = (W - M - ix - 0.15 * (sv.reinvestmentItems.length - 1)) / sv.reinvestmentItems.length;
    sv.reinvestmentItems.forEach(function (it, k) {
      s.push(text(ix + k * (iw + 0.15), ry, iw, 0.56, tpl(it), { size: 16, bold: true, fill: 'accentSoft', color: 'text', align: 'center', valign: 'middle', radius: 0.28 }));
    });

    // Illustrative tag + assumptions toggle (HTML) + overlay panel
    if (sv.showIllustrativeTag) s.push(text(M, 6.02, 1.5, 0.32, caps(tpl(sv.illustrativeTag)), { size: 10, bold: true, color: 'secondary', line: 'secondary', lineW: 1, dash: 'dash', align: 'center', valign: 'middle', radius: 0.16 }));
    s.push({ type: 'toggle', x: M + 1.65, y: 6.02, w: 1.6, h: 0.32, label: tpl(L.assumptionsButton), target: 'assumptions' });
    s.push({ type: 'inputs', overlay: true, id: 'assumptions', x: M, y: top, w: hw, h: hh + 0.9, rows: rows });
    s.push(footnote(tpl(sv.disclaimerShort)));
    return { primitives: s, inputs: inp, results: r, rows: rows };
  }

  /* ---------- speaker notes: everything removed from the slide face ---------- */
  function notesFor(cfg, tpl, i, extra) {
    var a = cfg.agenda[i], who = ownerName(cfg, a.owner);
    var head = 'Presenter: ' + who + ' · Time guide: ' + a.time + '\n\n';
    var u1 = cfg.useCaseOne, u2 = cfg.useCaseTwo, v = cfg.value, sv = cfg.savings;
    function list(arr) { return arr.map(function (d, k) { return (k + 1) + '. ' + tpl(d); }).join('\n'); }
    var body = [
      tpl(cfg.opening.notes),
      'Problem: ' + tpl(u1.problemStatement) + '\n\nAn average document goes through ' + tpl(u1.reviewCycleRange) + ' review cycles (configure actual value).\n\nEvidence relationship: ' + tpl(u1.evidenceRelationship) +
        '\n\nRelease Plan, Requirements Document, Test Protocol, Code Review Report / PR evidence, and the RTM are tightly coupled: one release change can affect all of them.\n\nTypical review failure points:\n- ' + u1.failurePoints.map(tpl).join('\n- '),
      'TODAY (current use): ' + tpl(u1.currentStateNote) + '\n\nNEXT (future / adoption model): ' + tpl(u1.futureStateNote) + '\n\n' + tpl(u1.outcomeNote) +
        '\n\nTarget: reduce QE review cycles from ' + tpl(u1.reviewCycleRange) + ' toward ' + tpl(u1.targetReviewCycles) + ' for eligible checks. This is a target / assumption, not a guaranteed reduction.' +
        '\n\nLIVE DEMO (one focused scenario):\n' + list(u1.demoSteps) +
        '\n\nHANDOVER to ' + cfg.presenters.presenterTwo + ': “' + tpl(cfg.speakerNotes.handoverAfterSlide3) + '”',
      'Problem: ' + tpl(u2.problemStatement) + '\n\n' + u2.checkpointNotes.map(tpl).join('\n\n') +
        '\n\nPositioning: source-driven lifecycle quality assurance, not “review without documents”; a test protocol or design artefact can still be an input. Already in QE use at multiple stages.\n\nBoundary: ' + tpl(u2.codeReviewStatus.text) + ' Only mention if asked.',
      'One configured setup, reused at each checkpoint (current QE use). Sources: requirements, detailed design, acceptance criteria, test protocol / test suite.\n- LLM-based: requirement quality, content context, INCOSE-aligned review.\n- Hybrid: contextual assessment plus deterministic validation.\n- Script-based: link coverage, mapping, completeness, and traceability validation.\n\nLIVE DEMO (two rapid checkpoints, not a feature tour):\n' + list(u2.demoSteps),
      v.pillarNotes.map(tpl).join('\n\n') + '\n\nBottom line: ' + tpl(v.bottomLineNote) + '\n\nCaveat: ' + tpl(v.caveat) + ' Say “full-population, evidence-based review” or “100% coverage of applicable, configured checks”. Never say “100% sampling”.',
      'Use “QE capacity released”, not “time saved”: capacity is reinvested, not removed. Figures are configurable assumptions and targets, not guarantees.\n\nReinvest released QE capacity into: detailed technical quality reviews · risk-based assessment · requirements and design quality · early lifecycle quality support.' +
        (extra ? '\n\nASSUMPTIONS USED (synthetic placeholders):\n' + extra.map(function (r) { return '- ' + r.label + ': ' + r.display; }).join('\n') : '') +
        '\n\nDisclaimer: ' + sv.disclaimerLines.map(tpl).join(' ')
    ][i];
    return head + body;
  }

  /* ---------- deck ---------- */
  function buildDeck(cfg, state, calc) {
    var tpl = makeTpl(cfg);
    if (cfg.theme.diagramStyle && cfg.theme.diagramStyle.cardRadius !== undefined) R = cfg.theme.diagramStyle.cardRadius;
    var sav = slideSavings(cfg, tpl, state, calc);
    var slides = [
      { id: 'opening', variant: 'title', primitives: slideOpening(cfg, tpl), title: cfg.brand.toolName + ': ' + tpl(cfg.brand.deckTitle) },
      { id: 'uc1-problem', primitives: slideUc1Problem(cfg, tpl), title: tpl(cfg.useCaseOne.titleProblem) },
      { id: 'uc1-approach', primitives: slideUc1Approach(cfg, tpl), title: tpl(cfg.useCaseOne.titleApproach) },
      { id: 'uc2-problem', primitives: slideUc2Problem(cfg, tpl), title: tpl(cfg.useCaseTwo.titleProblem) },
      { id: 'uc2-approach', primitives: slideUc2Approach(cfg, tpl), title: tpl(cfg.useCaseTwo.titleApproach) },
      { id: 'value', primitives: slideValue(cfg, tpl), title: tpl(cfg.value.title) },
      { id: 'savings', primitives: sav.primitives, title: tpl(cfg.savings.title), savings: { inputs: sav.inputs, results: sav.results } }
    ];
    slides.forEach(function (sl, i) {
      sl.notes = notesFor(cfg, tpl, i, i === 6 ? sav.rows : null);
      sl.owner = ownerName(cfg, cfg.agenda[i].owner);
      sl.time = cfg.agenda[i].time;
    });
    return slides;
  }

  /* Master-slide chrome. The title variant shows only the logo. */
  function chrome(cfg, variant) {
    var tpl = makeTpl(cfg), ls = cfg.theme.logoStyle, title = variant === 'title';
    return {
      footer: !title,
      footerRule: rect(M, LAYOUT.footerRuleY, CW, 0.012, { fill: 'neutral' }),
      footerText: text(M, LAYOUT.footerY, 9, 0.3, tpl(cfg.brand.footerText), { size: 10, color: 'muted', valign: 'middle' }),
      slideNumber: { x: W - M - 1.0, y: LAYOUT.footerY, w: 1.0, h: 0.3, size: 10, color: 'muted' },
      logo: {
        x: title ? M : ls.x, y: title ? 0.5 : ls.y - 0.1, w: ls.w, h: ls.h, dashed: ls.dashed,
        text: cfg.brand.logoPlaceholder.text, imagePath: cfg.brand.logoPlaceholder.imagePath
      }
    };
  }

  function initialState(cfg) {
    return { releaseId: cfg.savings.defaultReleaseScenario, userId: cfg.savings.defaultUserScenario, overrides: {} };
  }

  /* Visible words on a slide face (excludes footnotes and overlays) — used by tests. */
  function visibleWords(slide) {
    var n = 0;
    slide.primitives.forEach(function (p) {
      if (p.type !== 'text' || p.footnote) return;
      var t = p.runs ? p.runs.map(function (r) { return r.text; }).join(' ') : p.text;
      // Count words with 2+ letters (numbers and units like "h" are visuals, not words).
      n += String(t).split(/\s+/).filter(function (w) { return /[A-Za-z]{2,}/.test(w); }).length;
    });
    return n;
  }

  var api = { buildDeck: buildDeck, chrome: chrome, initialState: initialState, LAYOUT: LAYOUT, makeTpl: makeTpl, visibleWords: visibleWords };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.AIDA_SLIDES = api;
})(typeof self !== 'undefined' ? self : this);
