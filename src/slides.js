/*
 * Slide model — one source of truth for both renderers.
 *
 * Each slide is a list of simple primitives positioned in INCHES on a
 * 13.333 × 7.5 in (16:9) canvas. The HTML renderer (src/render-html.js) and
 * the PowerPoint exporter (src/pptx-builder.js) both draw from this list, so
 * every diagram exports as native, editable PowerPoint shapes and text.
 *
 * Primitive types
 *   rect / ellipse : x, y, w, h, fill, line, lineW (pt), dash, radius (in)
 *   line           : x, y, w, h (w or h = 0), color, width (pt), dash, arrow, arrowStart
 *   text           : x, y, w, h, text | runs[], size (pt), color, bold, italic,
 *                    align, valign, font ('heading'|'body'), fill, line, radius, role
 *   inputs         : savings assumptions (interactive in HTML, a static list in PPTX)
 * Colours are palette token names (see config.theme.palettes).
 */
(function (root) {
  'use strict';

  var W = 13.333, H = 7.5, M = 0.6, CW = W - 2 * M;
  var LAYOUT = { width: W, height: H, margin: M, contentWidth: CW, footerY: 6.97, footerRuleY: 6.9 };
  var R = 0.1; // default corner radius (overridden by theme.diagramStyle.cardRadius)

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

  /* Standard slide header: section pill, presenter, action headline, tagline. */
  function header(cfg, tpl, idx, section, headlineRuns, tagline, emphasis) {
    var s = [], label = caps(tpl(section));
    s.push(text(M, 0.42, pillW(label), 0.32, label, { size: 10.5, bold: true, color: 'accent', fill: 'accentSoft', align: 'center', valign: 'middle', radius: 0.16 }));
    s.push(text(6.4, 0.42, 4.8, 0.32, ownerName(cfg, cfg.agenda[idx].owner), { size: 11, color: 'muted', align: 'right', valign: 'middle' }));
    s.push(text(M, 0.86, CW, 0.98, headlineRuns.map(function (r) {
      return r.emphasis ? { text: tpl(r.text), bold: true, color: emphasis || 'accent' } : { text: tpl(r.text) };
    }), { size: 28, font: 'heading', color: 'text', valign: 'bottom', role: 'title' }));
    if (tagline) {
      s.push(rect(M, 1.93, 0.05, 0.34, { fill: emphasis || 'accent' }));
      s.push(text(M + 0.2, 1.93, CW - 0.2, 0.34, tpl(tagline), { size: 15, italic: true, color: 'muted', valign: 'middle' }));
    }
    return s;
  }

  function livePill(x, y, label) {
    var l = caps(label || 'Live demo');
    return text(x, y, pillW(l), 0.32, l, { size: 10.5, bold: true, fill: 'panel', color: 'onPanel', align: 'center', valign: 'middle', radius: 0.16 });
  }

  /* ---------- slides ---------- */

  function slideOpening(cfg, tpl) {
    var o = cfg.opening, s = [], px = 8.35;
    // Left — identity
    s.push(text(M, 1.5, 7.4, 0.34, caps(tpl(cfg.brand.deckSubtitle)), { size: 12, bold: true, color: 'accent', valign: 'middle' }));
    s.push(text(M, 1.82, 7.4, 1.4, cfg.brand.toolName, { size: 80, bold: true, color: 'accent', font: 'heading', valign: 'middle' }));
    s.push(text(M, 3.15, 7.4, 0.7, tpl(cfg.brand.deckTitle), { size: 30, bold: true, color: 'text', font: 'heading', valign: 'middle', role: 'title' }));
    s.push(rect(M, 4.02, 0.8, 0.07, { fill: 'warning' }));
    s.push(text(M, 4.3, 7.1, 1.0, tpl(o.keyMessage), { size: 20, color: 'secondary' }));
    s.push(text(M, 6.05, 7.3, 0.62, [
      { text: caps('Presented by'), size: 10.5, bold: true, color: 'muted', breakLine: true },
      { text: cfg.presenters.presenterOne + '  ·  ' + cfg.presenters.presenterTwo + '  ·  ' + cfg.brand.roleLabel, bold: true }
    ], { size: 15, color: 'text' }));

    // Right — how the tool reviews, as a vertical flow on a navy panel
    s.push(rect(px, 0, W - px, H, { fill: 'panel' }));
    var cx = px + 0.55, tw = W - cx - 0.75;
    s.push(text(cx, 1.0, tw + 0.5, 0.3, caps('How ' + cfg.brand.toolName + ' reviews'), { size: 11, bold: true, color: 'onPanelMuted', valign: 'middle' }));
    s.push(text(cx, 1.32, tw + 0.5, 0.65, tpl(o.supportingLine), { size: 15, color: 'onPanel' }));
    var ys = [2.35, 3.4, 4.45], d = 0.5;
    s.push(line(cx + d / 2, ys[0] + d, 0, ys[2] - ys[0] - d, { color: 'onPanelMuted', width: 1.25 }));
    o.operatingModes.forEach(function (m, i) {
      s.push.apply(s, numberDot(cx, ys[i], d, i + 1));
      s.push(text(cx + 0.72, ys[i] - 0.06, tw, 0.36, tpl(m.title), { size: 17, bold: true, color: 'onPanel', valign: 'middle' }));
      s.push(text(cx + 0.72, ys[i] + 0.3, tw, 0.32, tpl(m.text), { size: 13, color: 'onPanelMuted', valign: 'middle' }));
    });
    s.push(text(cx, 5.45, tw + 0.5, 0.3, caps('Solution areas'), { size: 11, bold: true, color: 'onPanelMuted', valign: 'middle' }));
    o.solutionAreas.forEach(function (a, i) {
      var y = 5.82 + i * 0.5;
      s.push(text(cx, y, 3.75, 0.4, a.focus ? a.label + '  ·  today’s focus' : a.label, a.focus
        ? { size: 13, bold: true, fill: 'accent', color: 'onAccent', align: 'center', valign: 'middle', radius: 0.2 }
        : { size: 13, line: 'onPanelMuted', lineW: 1, color: 'onPanelMuted', align: 'center', valign: 'middle', radius: 0.2 }));
    });
    return s;
  }

  function slideUc1Problem(cfg, tpl) {
    var u = cfg.useCaseOne, s = header(cfg, tpl, 1, u.sectionProblem, u.headlineProblem, u.tagline, 'warning');
    var n = u.artefacts.length, bw = 2.15, gap = (CW - n * bw) / (n - 1), by = 2.6, bh = 1.0, midY = by + bh / 2;
    var gaps = {};
    u.issueLabels.forEach(function (g) { gaps[g.afterArtefact] = g; });
    function bx(i) { return M + i * (bw + gap); }
    for (var i = 0; i < n; i++) {
      var a = u.artefacts[i];
      s.push(text(bx(i), by, bw, bh, [
        { text: caps(tpl(a.eyebrow)), size: 10.5, color: 'accent', breakLine: true },
        { text: tpl(a.name), size: 15 }
      ], { bold: true, align: 'center', valign: 'middle', fill: 'background', line: i === 0 ? 'accent' : 'accentLine', lineW: i === 0 ? 2 : 1.25, radius: R, margin: 5 }));
      if (i < n - 1) {
        var lx = bx(i) + bw, g = gaps[i];
        if (g) {
          s.push(line(lx, midY, gap, 0, { color: 'warning', width: 1.75, dash: 'dash' }));
          var cxm = lx + gap / 2, dd = 0.3;
          s.push(ellipse(cxm - dd / 2, midY - dd / 2, dd, dd, { fill: 'warning' }));
          s.push(text(cxm - dd / 2, midY - dd / 2, dd, dd, '!', { size: 13, bold: true, color: 'onAccent', align: 'center', valign: 'middle' }));
          var tw = 2.3;
          s.push(text(cxm - tw / 2, by + bh + 0.14, tw, 0.4, tpl(g.text), { size: 11, bold: true, color: 'warning', fill: 'warningSoft', align: 'center', valign: 'middle', radius: 0.18 }));
        } else {
          s.push(line(lx + 0.03, midY, gap - 0.06, 0, { color: 'accent', width: 1.75, arrow: true }));
        }
      }
    }
    // Return-to-author loop (red)
    var x0 = bx(0) + bw / 2, x4 = bx(n - 1) + bw / 2, ly = 4.42;
    s.push(line(x4, by + bh, 0, ly - by - bh, { color: 'warning', width: 1.75 }));
    s.push(line(x0, ly, x4 - x0, 0, { color: 'warning', width: 1.75 }));
    s.push(line(x0, by + bh + 0.02, 0, ly - by - bh - 0.02, { color: 'warning', width: 1.75, arrowStart: true }));
    var loop = tpl(u.loopLabel).replace('{cycles}', tpl(u.reviewCycleRange));
    s.push(text(W / 2 - 2.9, ly - 0.2, 5.8, 0.4, loop, { size: 14, bold: true, color: 'warning', fill: 'background', align: 'center', valign: 'middle' }));

    // Three impact cards
    var cg = 0.3, cw = (CW - 2 * cg) / 3, cy = 4.98, ch = 1.62;
    u.problemCards.forEach(function (c, k) {
      var x = M + k * (cw + cg);
      s.push(rect(x, cy, cw, ch, { fill: 'surface', radius: R }));
      s.push(rect(x + 0.18, cy + 0.24, 0.06, ch - 0.48, { fill: 'warning' }));
      s.push(text(x + 0.42, cy + 0.2, cw - 0.6, 0.42, tpl(c.title), { size: 18, bold: true, font: 'heading', valign: 'middle' }));
      s.push(text(x + 0.42, cy + 0.66, cw - 0.6, ch - 0.8, tpl(c.text), { size: 14, color: 'muted' }));
    });
    return s;
  }

  function slideUc1Approach(cfg, tpl) {
    var u = cfg.useCaseOne, s = header(cfg, tpl, 2, u.sectionApproach, u.headlineApproach, u.taglineApproach);
    var lh = 1.28, lys = [2.45, 3.86], sx = 3.0, sw = 2.7, sg = (W - M - 0.2 - sx - 3 * sw) / 2;
    u.lanes.forEach(function (lane, li) {
      var y = lys[li], fut = lane.future;
      s.push(rect(M, y, CW, lh, fut ? { fill: 'background', line: 'accentLine', lineW: 1.25, dash: 'dash', radius: R } : { fill: 'surface', radius: R }));
      var tag = tpl(lane.tag).split(' · ');
      s.push(text(M + 0.3, y, 2.1, lh, [
        { text: caps(tag[0]), size: 18, bold: true, color: fut ? 'secondary' : 'accent', breakLine: tag.length > 1 },
        { text: tag.slice(1).join(' · '), size: 12, color: 'muted' }
      ], { valign: 'middle' }));
      lane.steps.forEach(function (st, k) {
        var x = sx + k * (sw + sg), yb = y + 0.18, hb = lh - 0.36;
        var hl = !!st.highlight;
        s.push(text(x, yb, sw, hb, [
          { text: caps(tpl(st.role)), size: 10.5, bold: true, color: hl ? 'onPanelMuted' : 'accent', breakLine: true },
          { text: tpl(st.text), size: 15, bold: true, color: hl ? 'onAccent' : 'text' }
        ], hl
          ? { fill: 'accent', align: 'center', valign: 'middle', radius: R, margin: 4 }
          : { fill: 'background', line: fut ? 'secondary' : 'neutral', lineW: 1, dash: fut, align: 'center', valign: 'middle', radius: R, margin: 4 }));
        if (k < lane.steps.length - 1) s.push(line(x + sw + 0.06, y + lh / 2, sg - 0.12, 0, { color: 'accent', width: 1.75, arrow: true }));
      });
    });

    // Target KPI
    var ky = 5.36;
    s.push(text(M, ky, 4.4, 0.3, caps(tpl(u.targetLabel)), { size: 10.5, bold: true, color: 'muted', valign: 'middle' }));
    s.push(text(M, ky + 0.3, 4.4, 0.7, [
      { text: tpl(u.reviewCycleRange), color: 'muted' },
      { text: '  →  ', color: 'muted' },
      { text: tpl(u.targetReviewCycles), color: 'accent' }
    ], { size: 32, bold: true, font: 'heading', valign: 'middle' }));
    s.push(text(M, ky + 1.0, 4.4, 0.3, tpl(u.targetCaveat), { size: 11.5, italic: true, color: 'muted', valign: 'middle' }));

    // Demo stepper
    var dx = 5.5, dx2 = W - M - 0.5, dy = 5.95, n = u.demoSteps.length, step = (dx2 - dx) / (n - 1), d = 0.4;
    s.push(livePill(dx - 0.2, ky, 'Live demo · one scenario'));
    s.push(line(dx, dy, dx2 - dx, 0, { color: 'accentLine', width: 2 }));
    u.demoSteps.forEach(function (st, k) {
      var cx = dx + k * step;
      s.push.apply(s, numberDot(cx - d / 2, dy - d / 2, d, k + 1));
      s.push(text(cx - 0.85, dy + 0.3, 1.7, 0.4, tpl(st.short), { size: 12, bold: true, align: 'center' }));
    });
    return s;
  }

  function slideUc2Problem(cfg, tpl) {
    var u = cfg.useCaseTwo, s = header(cfg, tpl, 3, u.sectionProblem, u.headlineProblem, u.tagline);
    // "Room to change narrows" — risk signal in red
    s.push(text(M, 2.42, CW, 0.28, tpl(u.changeArrowLabel), { size: 11.5, italic: true, bold: true, color: 'warning', align: 'right', valign: 'middle' }));
    s.push(line(M, 2.78, CW, 0, { color: 'warning', width: 1.25, arrow: true }));

    var n = u.lifecycleLabels.length, g = 0.12, pw = (CW - (n - 1) * g) / n, py = 2.98, ph = 0.6;
    var hl = {};
    u.checkpointLabels.forEach(function (c, i) { hl[c.stage] = i + 1; });
    function px(i) { return M + i * (pw + g); }
    u.lifecycleLabels.forEach(function (lab, i) {
      var on = hl[i];
      s.push(text(px(i), py, pw, ph, on ? [{ text: on + '   ', color: 'onPanelMuted' }, { text: tpl(lab) }] : tpl(lab),
        on ? { size: 16, bold: true, fill: 'accent', color: 'onAccent', align: 'center', valign: 'middle', radius: 0.3 }
           : { size: 15, fill: 'surface', color: 'muted', align: 'center', valign: 'middle', radius: 0.3 }));
    });
    // Checkpoint cards hang from their stage
    var cy = 4.0, ch = 1.88, cw = 2 * pw + g;
    u.checkpointLabels.forEach(function (c) {
      var x = px(c.stage);
      s.push(line(x + pw / 2, py + ph, 0, cy - py - ph, { color: 'accent', width: 1.75 }));
      s.push(rect(x, cy, cw, ch, { fill: 'background', line: 'accentLine', lineW: 1.25, radius: R }));
      s.push(text(x + 0.3, cy + 0.2, cw - 0.6, 0.3, caps(tpl(c.when)), { size: 10.5, bold: true, color: 'accent', valign: 'middle' }));
      s.push(text(x + 0.3, cy + 0.52, cw - 0.6, 0.8, tpl(c.question), { size: 19, bold: true, font: 'heading' }));
      s.push(text(x + 0.3, cy + 1.36, cw - 0.6, 0.36, [
        { text: caps('Catches') + '   ', bold: true, color: 'warning', size: 10.5 },
        { text: tpl(c.value) }
      ], { size: 14, color: 'muted', valign: 'middle' }));
    });
    // Sources
    var sy = 6.18;
    s.push(text(M, sy, 1.9, 0.42, caps(tpl(u.sourcesLabel)), { size: 10.5, bold: true, color: 'muted', valign: 'middle' }));
    var sxs = M + 1.9, sw = (W - M - sxs - 0.12 * (u.sources.length - 1)) / u.sources.length;
    u.sources.forEach(function (src, k) {
      s.push(text(sxs + k * (sw + 0.12), sy, sw, 0.42, tpl(src), { size: 13, fill: 'surface', color: 'secondary', align: 'center', valign: 'middle', radius: 0.21 }));
    });
    if (u.codeReviewStatus && u.codeReviewStatus.showOnSlide) {
      s.push(text(M, 6.66, CW, 0.22, 'Scope note: ' + tpl(u.codeReviewStatus.text), { size: 10, italic: true, color: 'muted', align: 'right', valign: 'middle' }));
    }
    return s;
  }

  function slideUc2Approach(cfg, tpl) {
    var u = cfg.useCaseTwo, s = header(cfg, tpl, 4, u.sectionApproach, u.headlineApproach, u.taglineApproach);
    var top = 2.55;
    // Inputs
    s.push(text(M, top, 2.5, 0.3, caps('Connected sources'), { size: 10.5, bold: true, color: 'muted', valign: 'middle' }));
    var hubX = 3.65, hubW = 5.1, hubY = top, hubH = 3.0, chipW = 2.4;
    u.sources.forEach(function (src, k) {
      var y = top + 0.45 + k * 0.64;
      s.push(text(M, y, chipW, 0.48, tpl(src), { size: 13.5, fill: 'surface', color: 'secondary', align: 'center', valign: 'middle', radius: R }));
      s.push(line(M + chipW + 0.05, y + 0.24, hubX - M - chipW - 0.1, 0, { color: 'accentLine', width: 1.5, arrow: true }));
    });
    // Hub
    s.push(rect(hubX, hubY, hubW, hubH, { fill: 'panel', radius: 0.14 }));
    s.push(text(hubX + 0.35, hubY + 0.2, hubW - 0.7, 0.62, [
      { text: caps(tpl(u.hubTitle)), size: 10.5, color: 'onPanelMuted', breakLine: true },
      { text: 'Configured once by QEs', size: 19 }
    ], { bold: true, color: 'onPanel' }));
    u.reusableSetup.forEach(function (r, k) {
      var y = hubY + 1.0 + k * 0.64;
      s.push(text(hubX + 0.35, y, 1.55, 0.48, tpl(r.mode), { size: 13, bold: true, fill: 'accent', color: 'onAccent', align: 'center', valign: 'middle', radius: 0.24 }));
      s.push(text(hubX + 2.05, y, hubW - 2.3, 0.48, tpl(r.text), { size: 14, color: 'onPanel', valign: 'middle' }));
    });
    // Outputs
    var ox = 9.5, ow = W - M - ox, oh = 1.3, oys = [hubY + 0.05, hubY + hubH - oh - 0.05];
    u.checkpointLabels.forEach(function (c, i) {
      var y = oys[i];
      s.push(line(hubX + hubW + 0.05, y + oh / 2, ox - hubX - hubW - 0.1, 0, { color: 'accent', width: 1.75, arrow: true }));
      s.push(rect(ox, y, ow, oh, { fill: 'background', line: 'accentLine', lineW: 1.25, radius: R }));
      s.push.apply(s, numberDot(ox + 0.22, y + 0.24, 0.44, i + 1));
      s.push(text(ox + 0.8, y + 0.2, ow - 0.95, 0.5, tpl(c.title), { size: 18, bold: true, font: 'heading', valign: 'middle' }));
      s.push(text(ox + 0.22, y + 0.72, ow - 0.35, 0.52, tpl(u.checkpointApplies[i] || ''), { size: 12.5, color: 'muted', valign: 'middle' }));
    });
    // Demo
    var dy = 5.9;
    s.push(livePill(M, dy + 0.12, 'Live demo'));
    var dx = M + 1.55, dw = (W - M - dx - 0.2) / 2;
    u.demoSteps.forEach(function (st, k) {
      var x = dx + k * (dw + 0.2);
      s.push(text(x, dy, dw, 0.56, [
        { text: (k + 1) + '   ', bold: true, color: 'accent' },
        { text: tpl(st.short), bold: true },
        { text: '  ·  ' + tpl(u.checkpointApplies[k] || ''), color: 'muted' }
      ], { size: 14, fill: 'surface', valign: 'middle', radius: 0.28, margin: [0, 14, 0, 14] }));
    });
    return s;
  }

  function slideValue(cfg, tpl) {
    var v = cfg.value, s = header(cfg, tpl, 5, v.section, v.headline, null);
    var top = 2.2, fx = M, fw = 4.55, ax = fx + fw, aw = 0.75, tx = ax + aw, tw = 4.9, px = tx + tw + 0.25, pw = W - M - px;
    s.push(text(fx, top, fw, 0.3, caps('From'), { size: 10.5, bold: true, color: 'muted', valign: 'middle' }));
    s.push(text(tx, top, tw, 0.3, caps('To'), { size: 10.5, bold: true, color: 'accent', valign: 'middle' }));
    v.shifts.forEach(function (r, i) {
      var y = top + 0.42 + i * 1.0, h = 0.82;
      s.push(text(fx, y, fw, h, tpl(r.from), { size: 16, color: 'muted', fill: 'surface', valign: 'middle', radius: R, margin: [0, 18, 0, 18] }));
      s.push(line(ax + 0.12, y + h / 2, aw - 0.24, 0, { color: 'accent', width: 2, arrow: true }));
      s.push(text(tx, y, tw, h, tpl(r.to), r.kind === 'future'
        ? { size: 17, bold: true, color: 'text', fill: 'background', line: 'secondary', lineW: 1.25, dash: 'dash', valign: 'middle', radius: R, margin: [0, 18, 0, 18] }
        : { size: 17, bold: true, color: 'text', fill: 'accentSoft', line: 'accent', lineW: 1.25, valign: 'middle', radius: R, margin: [0, 18, 0, 18] }));
      s.push(statusPill(px, y + h / 2 - 0.16, pw, tpl(r.tag), r.kind));
    });
    s.push(text(M, 5.64, CW, 0.3, [{ text: 'Quality caveat:  ', bold: true }, { text: tpl(v.caveat), italic: true }], { size: 13, color: 'muted', valign: 'middle' }));
    s.push(rect(M, 6.02, CW, 0.74, { fill: 'panel', radius: R }));
    s.push(rect(M + 0.25, 6.2, 0.06, 0.38, { fill: 'warning' }));
    s.push(text(M + 0.5, 6.02, CW - 0.8, 0.74, tpl(v.bottomLine), { size: 17, bold: true, color: 'onPanel', valign: 'middle' }));
    return s;
  }

  function slideSavings(cfg, tpl, state, calc) {
    var sv = cfg.savings, L = sv.labels;
    var s = header(cfg, tpl, 6, sv.section, sv.headline, sv.showIllustrativeTag ? tpl(sv.illustrativeTag) + ' · adjust the assumptions live' : null);
    var inp = calc.resolveInputs(sv, state.releaseId, state.userId, state.overrides);
    var r = calc.calculateCapacity(inp);
    var fmt = calc.fmt;

    // Left — assumptions
    var px = M, py = 2.45, pw = 3.95, ph = 3.3;
    s.push(rect(px, py, pw, ph, { fill: 'surface', radius: R }));
    s.push(text(px + 0.2, py + 0.08, pw - 0.4, 0.3, caps('Assumptions'), { size: 10.5, bold: true, color: 'muted', valign: 'middle' }));
    s.push({
      type: 'inputs', x: px + 0.2, y: py + 0.42, w: pw - 0.4, h: ph - 0.52,
      rows: [
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
      ]
    });

    // Centre — before → after bars
    var cx = 4.9, cw = 3.9;
    s.push(text(cx, py + 0.08, cw, 0.3, caps(tpl(L.barsTitle)), { size: 10.5, bold: true, color: 'muted', valign: 'middle' }));
    var cur = r.currentQeHoursPerRelease, aid = r.aidaAssistedQeHoursPerRelease;
    var max = Math.max(cur || 0, aid || 0), bx = cx + 1.25, barMax = 1.75;
    function barW(v) { return v && max > 0 ? Math.max(0.05, barMax * Math.max(v, 0) / max) : 0.05; }
    [[tpl(L.currentBar), cur, 'secondary'], [tpl(L.aidaBar), aid, 'accent']].forEach(function (b, i) {
      var y = py + 0.6 + i * 0.95, w = barW(b[1]);
      s.push(text(cx, y, 1.2, 0.62, b[0], { size: 14, bold: true, color: 'text', valign: 'middle' }));
      s.push(rect(bx, y, w, 0.62, { fill: b[2], radius: 0.06 }));
      s.push(text(bx + w + 0.1, y, 1.0, 0.62, fmt(b[1]) + ' h', { size: 17, bold: true, valign: 'middle' }));
    });
    s.push(line(cx, py + 2.4, cw, 0, { color: 'neutral', width: 0.75 }));
    s.push(text(cx, py + 2.5, cw, 0.7, [
      { text: caps('Review cycles'), size: 10.5, bold: true, color: 'muted', breakLine: true },
      { text: tpl(String(inp.baselineReviewCycles)) + '  →  ', bold: true, color: 'muted' },
      { text: tpl(String(inp.aidaTargetReviewCycles)), bold: true, color: 'accent' },
      { text: '   ' + tpl(L.cyclesTargetNote), size: 11, italic: true, color: 'muted' }
    ], { size: 17 }));

    // Right — KPIs (one hero, two supporting)
    var rx = 9.1, rw = W - M - rx;
    s.push(rect(rx, py, rw, 1.75, { fill: 'panel', radius: R }));
    s.push(text(rx + 0.3, py + 0.15, rw - 0.6, 0.95, fmt(r.qeCapacityReleasedHours) + ' h', { size: 50, bold: true, color: 'onPanel', font: 'heading', valign: 'middle' }));
    s.push(text(rx + 0.3, py + 1.1, rw - 0.6, 0.5, tpl(L.resultReleased), { size: 13.5, color: 'onPanelMuted', valign: 'middle' }));
    var sw = (rw - 0.2) / 2;
    [[r.qeCapacityReleasedPercent === null ? '—' : fmt(r.qeCapacityReleasedPercent, 0) + '%', tpl(L.resultPercent)],
     [r.annualQeCapacityReleasedHours === null ? '—' : fmt(r.annualQeCapacityReleasedHours) + ' h', tpl(L.resultAnnual)]].forEach(function (k, i) {
      var x = rx + i * (sw + 0.2), y = py + 1.95;
      s.push(rect(x, y, sw, 1.35, { fill: 'accentSoft', radius: R }));
      s.push(text(x + 0.2, y + 0.1, sw - 0.4, 0.66, k[0], { size: 28, bold: true, color: 'accent', font: 'heading', valign: 'middle' }));
      s.push(text(x + 0.15, y + 0.74, sw - 0.3, 0.56, k[1], { size: 11, color: 'secondary' }));
    });

    // Reinvestment row
    var ry = 5.95;
    s.push(text(M, ry, 1.9, 0.42, caps(tpl(sv.reinvestmentLead)) + '  →', { size: 11, bold: true, color: 'accent', valign: 'middle' }));
    var ix = M + 1.95, iw = (W - M - ix - 0.12 * (sv.reinvestmentItems.length - 1)) / sv.reinvestmentItems.length;
    sv.reinvestmentItems.forEach(function (it, k) {
      s.push(text(ix + k * (iw + 0.12), ry, iw, 0.42, tpl(it), { size: 12.5, bold: true, fill: 'accentSoft', color: 'text', align: 'center', valign: 'middle', radius: 0.21 }));
    });
    s.push(text(M, 6.47, CW, 0.38, sv.disclaimerLines.map(tpl).join(' '), { size: 8.5, color: 'muted', valign: 'middle' }));
    return { primitives: s, inputs: inp, results: r };
  }

  /* ---------- speaker notes ---------- */
  function notesFor(cfg, tpl, i) {
    var a = cfg.agenda[i], who = ownerName(cfg, a.owner);
    var head = 'Presenter: ' + who + ' · Time guide: ' + a.time + '\n\n';
    var u1 = cfg.useCaseOne, u2 = cfg.useCaseTwo, v = cfg.value;
    var body = [
      tpl(cfg.opening.notes),
      'Problem: ' + tpl(u1.problemStatement) + '\n\nAverage document: ' + tpl(u1.reviewCycleRange) + ' review cycles (configure actual value).\n\nEvidence relationship: ' + tpl(u1.evidenceRelationship) + '\n\nTypical review failure points:\n- ' + u1.failurePoints.map(tpl).join('\n- '),
      'Current state: ' + tpl(u1.currentStateNote) + '\nFuture: ' + tpl(u1.futureStateNote) + '\nOutcome: ' + tpl(u1.outcomeNote) +
        '\nDo not claim a guaranteed cycle reduction.\n\nDEMO (one focused scenario):\n' +
        u1.demoSteps.map(function (d, k) { return (k + 1) + '. ' + tpl(d.detail); }).join('\n') +
        '\n\nHANDOVER to ' + cfg.presenters.presenterTwo + ': “' + tpl(cfg.speakerNotes.handoverAfterSlide3) + '”',
      'Problem: ' + tpl(u2.problemStatement) + '\n\n' + u2.checkpointNotes.map(tpl).join('\n\n') +
        '\n\nPositioning: source-driven lifecycle quality assurance, not “review without documents”; a test protocol or design artefact can still be an input. Already in QE use at multiple stages.\n\nBoundary: ' + tpl(u2.codeReviewStatus.text) + ' Only mention if asked.',
      'Reusable setup: LLM-based (requirement quality, content context, INCOSE-aligned review); Hybrid (contextual assessment plus deterministic validation); Script-based (link coverage, mapping, completeness, traceability).\n\nDEMO (two rapid checkpoints, not a feature tour):\n' +
        u2.demoSteps.map(function (d, k) { return (k + 1) + '. ' + tpl(d.short) + ': ' + tpl(d.detail); }).join('\n'),
      v.pillarNotes.map(tpl).join('\n\n') + '\n\nUse “full-population, evidence-based review” or “100% coverage of applicable, configured checks”. Never say “100% sampling” or imply 100% assurance.',
      'Use “QE capacity released”, not “time saved”: capacity is reinvested, not removed. Figures are configurable assumptions and targets, not guarantees.\n\nDisclaimer: ' + cfg.savings.disclaimerLines.map(tpl).join(' ')
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
      sl.notes = notesFor(cfg, tpl, i);
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

  var api = { buildDeck: buildDeck, chrome: chrome, initialState: initialState, LAYOUT: LAYOUT, makeTpl: makeTpl };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.AIDA_SLIDES = api;
})(typeof self !== 'undefined' ? self : this);
