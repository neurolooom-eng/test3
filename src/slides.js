/*
 * Slide model — one source of truth for both renderers.
 *
 * Each slide is a list of simple primitives positioned in INCHES on a
 * 13.333 × 7.5 in (16:9) canvas. The HTML renderer (src/render-html.js) and
 * the PowerPoint exporter (src/pptx-builder.js) both draw from this list, so
 * every diagram exports as native, editable PowerPoint shapes and text.
 *
 * Primitive types
 *   rect / ellipse : x, y, w, h, fill, line, lineW (pt), dash
 *   line           : x, y, w, h (w or h = 0), color, width (pt), dash, arrow
 *   text           : x, y, w, h, text | runs[], size (pt), color, bold, italic,
 *                    align, valign, font ('heading'|'body'), fill, line, role
 *   inputs         : savings assumptions (interactive in HTML, a static list in PPTX)
 * Colours are palette token names (see config.theme.palettes).
 */
(function (root) {
  'use strict';

  var W = 13.333, H = 7.5, M = 0.6, CW = W - 2 * M;
  var LAYOUT = { width: W, height: H, margin: M, contentWidth: CW, titleY: 0.68, titleH: 0.62, footerY: 6.97, footerRuleY: 6.9 };

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
  function line(x, y, w, h, o) { return Object.assign({ type: 'line', x: x, y: y, w: w, h: h, color: 'neutral', width: 1.5 }, o || {}); }
  function text(x, y, w, h, t, o) {
    var p = Object.assign({ type: 'text', x: x, y: y, w: w, h: h, size: 18, color: 'text', font: 'body', align: 'left', valign: 'top' }, o || {});
    if (Array.isArray(t)) p.runs = t; else p.text = t;
    return p;
  }
  function caps(s) { return String(s).toUpperCase(); }

  /* Status tag: current = filled accent, future = dashed outline, outcome = solid outline.
     Each state differs by fill/border style AND its text — never colour alone. */
  function statusTag(x, y, w, label, kind) {
    var base = { size: 11, bold: true, align: 'center', valign: 'middle', font: 'body' };
    if (kind === 'current') return text(x, y, w, 0.32, caps(label), Object.assign(base, { fill: 'accent', color: 'onAccent' }));
    if (kind === 'future') return text(x, y, w, 0.32, caps(label), Object.assign(base, { line: 'secondary', lineW: 1.25, dash: 'dash', color: 'secondary' }));
    return text(x, y, w, 0.32, caps(label), Object.assign(base, { line: 'accent', lineW: 1.25, color: 'accent' }));
  }
  function statusCard(x, y, w, h, kind) {
    if (kind === 'current') return rect(x, y, w, h, { fill: 'background', line: 'accent', lineW: 1.25 });
    if (kind === 'future') return rect(x, y, w, h, { fill: 'background', line: 'secondary', lineW: 1.25, dash: 'dash' });
    return rect(x, y, w, h, { fill: 'surface' });
  }
  function numberDot(x, y, d, n) {
    return [
      ellipse(x, y, d, d, { fill: 'accent' }),
      text(x, y, d, d, String(n), { size: Math.round(d * 34), bold: true, color: 'onAccent', align: 'center', valign: 'middle' })
    ];
  }
  function kicker(t) { return text(M, 0.3, 9.5, 0.32, caps(t), { size: 12, bold: true, color: 'accent', valign: 'middle' }); }
  function title(t) { return text(M, LAYOUT.titleY, CW, LAYOUT.titleH, t, { size: 26, bold: true, font: 'heading', color: 'text', valign: 'middle', role: 'title' }); }

  function demoBand(y, label, steps, tpl, opts) {
    var out = [text(M, y, 4.8, 0.34, caps(label), { size: 12, bold: true, fill: 'primary', color: 'background', align: 'center', valign: 'middle' })];
    var n = steps.length, gap = 0.14, cy = y + 0.46, ch = opts.chipH;
    var cw = (CW - gap * (n - 1)) / n;
    steps.forEach(function (s, i) {
      var cx = M + i * (cw + gap);
      out.push(rect(cx, cy, cw, ch, { fill: 'surface' }));
      var runs = [{ text: (i + 1) + '  ', bold: true, color: 'accent' }, { text: tpl(s.short), bold: true }];
      if (opts.withDetail) runs.push({ text: ' — ' + tpl(s.detail), bold: false, color: 'muted' });
      out.push(text(cx + 0.16, cy, cw - 0.3, ch, runs, { size: opts.size, valign: 'middle' }));
    });
    return out;
  }

  /* ---------- slides ---------- */

  function slideOpening(cfg, tpl) {
    var o = cfg.opening, s = [];
    s.push(text(M, 0.3, 9.5, 0.32, 'Presented by ' + cfg.presenters.presenterOne + ' and ' + cfg.presenters.presenterTwo + ', ' + cfg.brand.roleLabel,
      { size: 13, color: 'muted', valign: 'middle' }));
    s.push(text(M, 0.95, CW, 0.34, caps(tpl(cfg.brand.deckSubtitle)), { size: 13, bold: true, color: 'accent', valign: 'middle' }));
    s.push(text(M, 1.3, CW, 0.85, tpl(cfg.brand.deckTitle), { size: 40, bold: true, font: 'heading', valign: 'middle', role: 'title' }));
    s.push(rect(M, 2.25, 0.9, 0.06, { fill: 'accent' }));
    s.push(text(M, 2.45, 11.6, 0.85, tpl(o.keyMessage), { size: 24, bold: true, color: 'text' }));
    s.push(text(M, 3.3, 11.2, 0.65, tpl(o.supportingLine), { size: 18, color: 'muted' }));

    // Solution areas
    s.push(text(M, 4.08, 2.0, 0.4, caps('Solution areas'), { size: 12, bold: true, color: 'muted', valign: 'middle' }));
    var cx = M + 2.0;
    o.solutionAreas.forEach(function (a) {
      var label = a.focus ? a.label + ' · today’s focus' : a.label;
      var w = a.focus ? 3.6 : 2.0;
      s.push(text(cx, 4.08, w, 0.4, label, a.focus
        ? { size: 15, bold: true, color: 'accent', fill: 'accentSoft', line: 'accent', lineW: 1, align: 'center', valign: 'middle' }
        : { size: 15, color: 'muted', line: 'neutral', lineW: 1, align: 'center', valign: 'middle' }));
      cx += w + 0.2;
    });

    // Operating modes — equal-weight cards
    s.push(text(M, 4.72, 6, 0.3, caps(cfg.brand.toolName + ' operating modes'), { size: 12, bold: true, color: 'muted', valign: 'middle' }));
    var gap = 0.25, cw = (CW - 2 * gap) / 3, y = 5.08, h = 1.6;
    o.operatingModes.forEach(function (m, i) {
      var x = M + i * (cw + gap);
      s.push(rect(x, y, cw, h, { fill: 'surface' }));
      s.push(rect(x, y, cw, 0.06, { fill: 'accent' }));
      s.push(text(x + 0.25, y + 0.2, cw - 0.5, 0.4, tpl(m.title), { size: 20, bold: true, font: 'heading' }));
      s.push(text(x + 0.25, y + 0.66, cw - 0.5, 0.85, tpl(m.text), { size: 17, color: 'secondary' }));
    });
    return s;
  }

  function slideUc1Problem(cfg, tpl) {
    var u = cfg.useCaseOne, s = [];
    s.push(kicker('Use case 1 · ' + ownerName(cfg, cfg.agenda[1].owner)));
    s.push(title(tpl(u.titleProblem)));
    s.push(text(M, 1.5, 8.35, 1.3, tpl(u.problemStatement), { size: 18 }));
    // Review-cycle figure
    s.push(rect(9.35, 1.5, 3.38, 1.3, { fill: 'surface' }));
    s.push(text(9.6, 1.58, 2.9, 0.68, tpl(u.reviewCycleRange), { size: 40, bold: true, color: 'accent', font: 'heading', valign: 'middle' }));
    s.push(text(9.6, 2.24, 2.9, 0.5, tpl(u.reviewCycleCaption), { size: 14, color: 'muted' }));

    // Evidence chain
    var n = u.artefactLabels.length, bw = 2.1, gap = (CW - n * bw) / (n - 1), by = 3.35, bh = 1.05, midY = by + bh / 2;
    s.push(text(M, by - 0.42, 1.9, 0.3, caps(tpl(u.changeTag)), { size: 11, bold: true, fill: 'accent', color: 'onAccent', align: 'center', valign: 'middle' }));
    var gaps = {};
    u.issueLabels.forEach(function (g) { gaps[g.afterArtefact] = g; });
    for (var i = 0; i < n; i++) {
      var x = M + i * (bw + gap);
      s.push(text(x, by, bw, bh, tpl(u.artefactLabels[i]), {
        size: 16, bold: true, align: 'center', valign: 'middle', fill: 'background',
        line: i === 0 ? 'accent' : 'text', lineW: i === 0 ? 2.25 : 1, margin: 6
      }));
      if (i < n - 1) {
        var lx = x + bw, g = gaps[i];
        s.push(line(lx, midY, gap, 0, g
          ? { color: 'warning', width: 1.75, dash: 'dash' }
          : { color: 'neutral', width: 1.75, arrow: true }));
        if (g) {
          var cxm = lx + gap / 2, d = 0.36;
          s.push(ellipse(cxm - d / 2, midY - d / 2, d, d, { fill: 'warningSoft', line: 'warning', lineW: 1.5 }));
          s.push(text(cxm - d / 2, midY - d / 2, d, d, '!', { size: 15, bold: true, color: 'warning', align: 'center', valign: 'middle' }));
          s.push(line(cxm, midY + d / 2, 0, 0.42, { color: 'warning', width: 1 }));
          var tw = 2.7;
          s.push(text(Math.min(Math.max(cxm - tw / 2, M), W - M - tw), by + bh + 0.32, tw, 0.62,
            [{ text: 'Gap: ', bold: true }, { text: tpl(g.text) }], { size: 15, color: 'warning', align: 'center', fill: 'warningSoft' }));
        }
      }
    }

    // Takeaway
    s.push(rect(M, 5.5, 0.08, 0.55, { fill: 'accent' }));
    s.push(text(M + 0.28, 5.5, CW - 0.28, 0.55, tpl(u.keyTakeaway), { size: 24, bold: true, font: 'heading', valign: 'middle' }));

    // Subtle source labels
    var sl = cfg.systemLabels, roles = cfg.systemLabelRoles || {};
    var runs = [{ text: 'Evidence sources:  ', bold: true }];
    Object.keys(sl).forEach(function (k, idx) {
      runs.push({ text: (idx ? '  ·  ' : '') + tpl(sl[k]) + (roles[k] ? ' (' + tpl(roles[k]) + ')' : '') });
    });
    s.push(text(M, 6.22, CW, 0.55, runs, { size: 13, color: 'muted' }));
    return s;
  }

  function slideUc1Approach(cfg, tpl) {
    var u = cfg.useCaseOne, s = [];
    s.push(kicker('Use case 1 · ' + ownerName(cfg, cfg.agenda[2].owner)));
    s.push(title(tpl(u.titleApproach)));
    var cols = [[u.currentState, 'current'], [u.futureState, 'future'], [u.outcome, 'outcome']];
    var gap = 0.34, cw = (CW - 2 * gap) / 3, ty = 1.5, cy = 1.92, ch = 3.1;
    cols.forEach(function (c, i) {
      var d = c[0], kind = c[1], x = M + i * (cw + gap);
      s.push(statusTag(x, ty, 2.7, tpl(d.tag), kind));
      s.push(statusCard(x, cy, cw, ch, kind));
      s.push(text(x + 0.25, cy + 0.18, cw - 0.5, 0.7, tpl(d.title), { size: 19, bold: true, font: 'heading' }));
      s.push(text(x + 0.25, cy + 0.95, cw - 0.5, ch - 1.05, tpl(d.text), { size: 17, color: 'secondary' }));
      if (i < cols.length - 1) s.push(line(x + cw + 0.04, cy + ch / 2, gap - 0.08, 0, { color: 'accent', width: 1.75, arrow: true }));
    });
    // Target statement (configurable, explicitly a target)
    var target = tpl(u.targetStatement).replace('{current}', tpl(u.reviewCycleRange)).replace('{target}', tpl(u.targetReviewCycles));
    s.push(rect(M, 5.12, 0.08, 0.6, { fill: 'accent' }));
    s.push(text(M + 0.28, 5.12, CW - 0.28, 0.6, [
      { text: target, bold: true, breakLine: true },
      { text: tpl(u.targetCaveat), italic: true, color: 'muted', size: 13 }
    ], { size: 17, valign: 'middle' }));
    s.push.apply(s, demoBand(5.82, 'Live demo · one end-to-end scenario', u.demoSteps, tpl, { chipH: 0.56, size: 15 }));
    return s;
  }

  function slideUc2Problem(cfg, tpl) {
    var u = cfg.useCaseTwo, s = [];
    s.push(kicker('Use case 2 · ' + ownerName(cfg, cfg.agenda[3].owner)));
    s.push(title(tpl(u.titleProblem)));
    s.push(text(M, 1.45, CW, 0.92, tpl(u.problemStatement), { size: 18 }));
    s.push(statusTag(M, 2.48, 1.7, tpl(u.currentUseTag), 'current'));
    s.push(text(M + 1.85, 2.46, CW - 1.85, 0.74, [
      { text: tpl(u.currentUseText), color: 'secondary', breakLine: true },
      { text: tpl(u.sourceNote), color: 'muted', size: 14 }
    ], { size: 15 }));

    // Timeline
    var ly = 3.58, xs = [1.95, 4.85, 7.75, 10.65];
    s.push(line(M, ly, CW, 0, { color: 'neutral', width: 2, arrow: true }));
    var hl = {};
    u.checkpointLabels.forEach(function (c, i) { hl[c.stage] = i + 1; });
    u.lifecycleLabels.forEach(function (lab, i) {
      var cx = xs[i] !== undefined ? xs[i] : M + (i + 0.5) * CW / u.lifecycleLabels.length;
      if (hl[i]) s.push.apply(s, numberDot(cx - 0.25, ly - 0.25, 0.5, hl[i]));
      else s.push(ellipse(cx - 0.13, ly - 0.13, 0.26, 0.26, { fill: 'background', line: 'neutral', lineW: 2 }));
      s.push(text(cx - 1.35, ly + 0.36, 2.7, 0.5, tpl(lab), { size: 14, bold: !!hl[i], color: hl[i] ? 'text' : 'muted', align: 'center' }));
    });

    // Checkpoint cards
    var gap = 0.33, cw = (CW - gap) / 2, cy = 4.52, ch = 2.28;
    u.checkpointLabels.forEach(function (c, i) {
      var x = M + i * (cw + gap);
      s.push(rect(x, cy, cw, ch, { fill: 'background', line: 'accent', lineW: 1 }));
      s.push.apply(s, numberDot(x + 0.22, cy + 0.2, 0.42, i + 1));
      s.push(text(x + 0.8, cy + 0.14, cw - 1.0, 0.6, [
        { text: tpl(c.title), bold: true, breakLine: true },
        { text: caps(tpl(c.when)), size: 11, bold: true, color: 'muted' }
      ], { size: 18, font: 'heading' }));
      s.push(text(x + 0.25, cy + 0.82, cw - 0.5, 0.92, tpl(c.question), { size: 18 }));
      s.push(text(x + 0.25, cy + 1.74, cw - 0.5, 0.52, [{ text: 'Value: ', bold: true }, { text: tpl(c.value) }], { size: 14, color: 'accent' }));
    });
    // Code review is out of scope for this use case; shown only when configured.
    if (u.codeReviewStatus && u.codeReviewStatus.showOnSlide) {
      s.push(text(M, 1.22, CW, 0.22, 'Scope note: ' + tpl(u.codeReviewStatus.text), { size: 11, italic: true, color: 'muted', align: 'right', valign: 'middle' }));
    }
    return s;
  }

  function slideUc2Approach(cfg, tpl) {
    var u = cfg.useCaseTwo, s = [];
    s.push(kicker('Use case 2 · ' + ownerName(cfg, cfg.agenda[4].owner)));
    s.push(title(tpl(u.titleApproach)));
    s.push(statusTag(M, 1.47, 1.7, 'Current use', 'current'));
    s.push(text(M + 1.85, 1.43, CW - 1.85, 0.4, tpl(u.approachSubtitle), { size: 18, color: 'secondary', valign: 'middle' }));

    var bx = M, by = 2.05, bw = 6.95, bh = 3.15;
    s.push(rect(bx, by, bw, bh, { fill: 'surface' }));
    s.push(text(bx + 0.25, by + 0.12, bw - 0.5, 0.36, caps('Configured quality setup · defined once by QEs'), { size: 12, bold: true, color: 'muted', valign: 'middle' }));
    u.reusableSetup.forEach(function (r, i) {
      var ry = by + 0.6 + i * 0.83;
      s.push(text(bx + 0.25, ry, 1.75, 0.68, tpl(r.mode), { size: 16, bold: true, fill: 'background', line: 'accent', lineW: 1.25, align: 'center', valign: 'middle', color: 'accent' }));
      s.push(text(bx + 2.2, ry, bw - 2.4, 0.68, tpl(r.text), { size: 17, valign: 'middle' }));
    });

    var rx = 8.4, rw = W - M - rx, rh = 1.4, ys = [by, by + bh - rh];
    u.checkpointLabels.forEach(function (c, i) {
      var ry = ys[i];
      s.push(rect(rx, ry, rw, rh, { fill: 'background', line: 'accent', lineW: 1 }));
      s.push.apply(s, numberDot(rx + 0.2, ry + 0.2, 0.42, i + 1));
      var tparts = tpl(c.title).split(' · ');
      s.push(text(rx + 0.75, ry + 0.12, rw - 0.9, 0.56, tparts.length > 1
        ? [{ text: caps(tparts[0]), size: 11, color: 'muted', breakLine: true }, { text: tparts.slice(1).join(' · ') }]
        : tpl(c.title), { size: 17, bold: true, font: 'heading', valign: 'middle' }));
      s.push(text(rx + 0.25, ry + 0.76, rw - 0.4, 0.58, tpl(u.checkpointApplies[i] || ''), { size: 15, color: 'secondary' }));
      s.push(line(bx + bw, ry + rh / 2, rx - bx - bw, 0, { color: 'accent', width: 1.75, arrow: true }));
    });
    s.push(text(bx + bw + 0.05, by + bh / 2 - 0.3, rx - bx - bw - 0.1, 0.6, 'Same checks, reused', { size: 12, color: 'muted', align: 'center', valign: 'middle', italic: true }));

    s.push.apply(s, demoBand(5.4, 'Live demo · two rapid checkpoints', u.demoSteps, tpl, { chipH: 0.92, size: 15, withDetail: true }));
    return s;
  }

  function slideValue(cfg, tpl) {
    var v = cfg.value, s = [];
    s.push(kicker('Combined value · ' + ownerName(cfg, cfg.agenda[5].owner)));
    s.push(title(tpl(v.title)));
    var gap = 0.3, cw = (CW - 2 * gap) / 3, ty = 1.5, cy = 1.92, ch = 3.1;
    v.pillars.forEach(function (p, i) {
      var kind = p.future ? 'future' : (i === 0 ? 'current' : 'outcome');
      var x = M + i * (cw + gap);
      s.push(statusTag(x, ty, 2.7, tpl(p.tag), kind));
      s.push(statusCard(x, cy, cw, ch, kind));
      s.push(text(x + 0.25, cy + 0.18, 0.6, 0.35, String(i + 1).padStart(2, '0'), { size: 15, bold: true, color: 'accent' }));
      s.push(text(x + 0.25, cy + 0.5, cw - 0.5, 0.7, tpl(p.title), { size: 20, bold: true, font: 'heading' }));
      s.push(text(x + 0.25, cy + 1.22, cw - 0.5, ch - 1.32, tpl(p.text), { size: 17, color: 'secondary' }));
    });
    s.push(text(M, 5.1, CW, 0.38, [{ text: 'Quality caveat: ', bold: true }, { text: tpl(v.caveat), italic: true }], { size: 16, color: 'muted', valign: 'middle' }));
    s.push(rect(M, 5.58, CW, 1.08, { fill: 'accentSoft' }));
    s.push(rect(M, 5.58, 0.08, 1.08, { fill: 'accent' }));
    s.push(text(M + 0.35, 5.58, CW - 0.6, 1.08, tpl(v.bottomLine), { size: 22, bold: true, font: 'heading', valign: 'middle' }));
    return s;
  }

  function slideSavings(cfg, tpl, state, calc) {
    var sv = cfg.savings, L = sv.labels, s = [];
    var inp = calc.resolveInputs(sv, state.releaseId, state.userId, state.overrides);
    var r = calc.calculateCapacity(inp);
    var fmt = calc.fmt;

    s.push(kicker('QE capacity · ' + ownerName(cfg, cfg.agenda[6].owner)));
    s.push(title(tpl(sv.title)));

    // Left — assumptions
    var px = M, py = 1.45, pw = 4.2, ph = 3.98;
    s.push(rect(px, py, pw, ph, { fill: 'surface' }));
    s.push(text(px + 0.2, py + 0.1, pw - 0.4, 0.34, caps('Assumptions'), { size: 12, bold: true, color: 'muted', valign: 'middle' }));
    if (sv.showIllustrativeTag) {
      s.push(text(px + 0.2, py + 0.46, pw - 0.4, 0.28, caps(tpl(sv.illustrativeTag)), { size: 9, bold: true, color: 'secondary', line: 'secondary', lineW: 0.75, dash: 'dash', align: 'center', valign: 'middle' }));
    }
    s.push({
      type: 'inputs', x: px + 0.2, y: py + 0.86, w: pw - 0.4, h: ph - 0.96,
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

    // Centre — before → after
    var cx = 5.1, cwid = 3.95;
    s.push(text(cx, py + 0.1, cwid, 0.34, caps('Before → after · QE hours per release'), { size: 12, bold: true, color: 'muted', valign: 'middle' }));
    var cur = r.currentQeHoursPerRelease, aid = r.aidaAssistedQeHoursPerRelease;
    var max = Math.max(cur || 0, aid || 0), barMax = 2.55;
    function barW(v) { return v && max > 0 ? Math.max(0.05, barMax * Math.max(v, 0) / max) : 0.05; }
    var bars = [[tpl(L.currentBar), cur, 'secondary'], [tpl(L.aidaBar), aid, 'accent']];
    bars.forEach(function (b, i) {
      var y0 = py + 0.62 + i * 1.18;
      s.push(text(cx, y0, cwid, 0.34, b[0], { size: 15, color: 'text', valign: 'middle' }));
      var bw = barW(b[1]);
      s.push(rect(cx, y0 + 0.42, bw, 0.55, { fill: b[2] }));
      s.push(text(cx + bw + 0.12, y0 + 0.42, 1.3, 0.55, fmt(b[1]) + ' h', { size: 20, bold: true, valign: 'middle', color: 'text' }));
    });
    s.push(line(cx, py + 3.05, cwid, 0, { color: 'neutral', width: 0.75 }));
    s.push(text(cx, py + 3.12, cwid, 0.66, [
      { text: 'Review cycles: ', bold: true },
      { text: tpl(String(inp.baselineReviewCycles)) + ' → ' + tpl(String(inp.aidaTargetReviewCycles)), breakLine: true },
      { text: tpl(L.cyclesTargetNote) + ' · eligible documents ' + fmt(r.eligibleDocuments), color: 'muted', size: 12 }
    ], { size: 15 }));

    // Right — three results
    var rx = 9.35, rw = W - M - rx, rh = ph / 3;
    var results = [
      [fmt(r.qeCapacityReleasedHours) + ' h', tpl(L.resultReleased)],
      [r.qeCapacityReleasedPercent === null ? '—' : fmt(r.qeCapacityReleasedPercent, 0) + '%', tpl(L.resultPercent)],
      [r.annualQeCapacityReleasedHours === null ? '—' : fmt(r.annualQeCapacityReleasedHours) + ' h', tpl(L.resultAnnual)]
    ];
    results.forEach(function (res, i) {
      var y0 = py + i * rh;
      if (i) s.push(line(rx, y0, rw, 0, { color: 'neutral', width: 0.75 }));
      s.push(text(rx, y0 + 0.08, rw, 0.7, res[0], { size: 34, bold: true, color: 'accent', font: 'heading', valign: 'middle' }));
      s.push(text(rx, y0 + 0.78, rw, 0.42, res[1], { size: 14, color: 'text' }));
    });

    // Reinvestment band
    s.push(rect(M, 5.54, CW, 0.64, { fill: 'accentSoft' }));
    s.push(rect(M, 5.54, 0.08, 0.64, { fill: 'accent' }));
    s.push(text(M + 0.3, 5.54, CW - 0.5, 0.64, [
      { text: tpl(sv.reinvestmentLead) + '  ', bold: true },
      { text: sv.reinvestmentItems.map(tpl).join('  ·  ') }
    ], { size: 16, valign: 'middle' }));

    // Disclaimer
    s.push(text(M, 6.25, CW, 0.6, sv.disclaimerLines.map(tpl).join(' '), { size: 10, color: 'muted', valign: 'top' }));
    return { primitives: s, inputs: inp, results: r };
  }

  /* ---------- speaker notes ---------- */
  function notesFor(cfg, tpl, i) {
    var a = cfg.agenda[i], who = ownerName(cfg, a.owner);
    var head = 'Presenter: ' + who + ' · Time guide: ' + a.time + '\n\n';
    var u1 = cfg.useCaseOne, u2 = cfg.useCaseTwo;
    var body = [
      'Open with the proposition: ' + tpl(cfg.opening.keyMessage) + '\nFocus today: Document Reviews and connected source evidence. Data Insights is out of scope unless relevant to a demo.',
      'Problem: ' + tpl(u1.problemStatement) + '\n\nEvidence relationship: ' + tpl(u1.evidenceRelationship) + '\n\nTypical review failure points:\n- ' + u1.failurePoints.map(tpl).join('\n- '),
      'Current state is in use today; author-initiated pre-review is a future / adoption model. Do not claim a guaranteed cycle reduction.\n\nDEMO (one focused scenario):\n' +
        u1.demoSteps.map(function (d, k) { return (k + 1) + '. ' + tpl(d.detail); }).join('\n') +
        '\n\nHANDOVER to ' + cfg.presenters.presenterTwo + ': “' + tpl(cfg.speakerNotes.handoverAfterSlide3) + '”',
      'Positioning: source-driven lifecycle quality assurance — not “review without documents”; a test protocol or design artefact can still be an input. This is already in QE use at multiple stages.\n\nBoundary: ' + tpl(u2.codeReviewStatus.text) + ' Only mention if asked.',
      'DEMO (two rapid checkpoints, not a feature tour):\n' + u2.demoSteps.map(function (d, k) { return (k + 1) + '. ' + tpl(d.short) + ': ' + tpl(d.detail); }).join('\n'),
      'Use “full-population, evidence-based review” or “100% coverage of applicable, configured checks”. Never say “100% sampling” or imply 100% assurance.\nPillar 2 is a future / adoption model where not yet deployed.',
      'Use “QE capacity released”, not “time saved”: capacity is reinvested, not removed. Figures are configurable assumptions and targets, not guarantees, and not a headcount-reduction claim. Human QE review of outputs remains necessary.'
    ][i];
    return head + body;
  }

  /* ---------- deck ---------- */
  function buildDeck(cfg, state, calc) {
    var tpl = makeTpl(cfg);
    var sav = slideSavings(cfg, tpl, state, calc);
    var slides = [
      { id: 'opening', primitives: slideOpening(cfg, tpl), title: tpl(cfg.brand.deckTitle) },
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

  /* Master-slide chrome shared by every slide: footer rule, footer text, logo. */
  function chrome(cfg) {
    var tpl = makeTpl(cfg), ls = cfg.theme.logoStyle;
    return {
      footerRule: rect(M, LAYOUT.footerRuleY, CW, 0.012, { fill: 'neutral' }),
      footerText: text(M, LAYOUT.footerY, 9, 0.3, tpl(cfg.brand.footerText), { size: 10, color: 'muted', valign: 'middle' }),
      slideNumber: { x: W - M - 1.0, y: LAYOUT.footerY, w: 1.0, h: 0.3, size: 10, color: 'muted' },
      logo: { x: ls.x, y: ls.y - 0.1, w: ls.w, h: ls.h, dashed: ls.dashed, text: cfg.brand.logoPlaceholder.text, imagePath: cfg.brand.logoPlaceholder.imagePath }
    };
  }

  function initialState(cfg) {
    return { releaseId: cfg.savings.defaultReleaseScenario, userId: cfg.savings.defaultUserScenario, overrides: {} };
  }

  var api = { buildDeck: buildDeck, chrome: chrome, initialState: initialState, LAYOUT: LAYOUT, makeTpl: makeTpl };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.AIDA_SLIDES = api;
})(typeof self !== 'undefined' ? self : this);
