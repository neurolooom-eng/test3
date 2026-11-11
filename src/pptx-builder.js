/*
 * PowerPoint builder — converts the slide primitives from src/slides.js into
 * native, editable PowerPoint shapes and text using PptxGenJS.
 * Works in the browser (window.PptxGenJS) and in Node (require('pptxgenjs')).
 *
 * Theme: colours and fonts come from config.theme, and every slide uses one of
 * two slide masters (AIDA_TITLE for the opening slide, AIDA_CONTENT for the rest).
 * Titles are real title placeholders so a corporate theme can restyle them.
 */
(function (root) {
  'use strict';

  var MASTER_TITLE = 'AIDA_TITLE', MASTER_CONTENT = 'AIDA_CONTENT';

  function build(PptxGenJS, cfg, state, calc, slidesApi) {
    var S = slidesApi || root.AIDA_SLIDES;
    var pal = cfg.theme.palettes[cfg.theme.activePalette] || cfg.theme.palettes.light;
    var fonts = cfg.theme.fonts;
    function col(t) { return (pal[t] || String(t)).replace('#', ''); }
    function face(p) { return p.font === 'heading' ? fonts.pptxHeading : fonts.pptxBody; }

    var L = S.LAYOUT;
    var pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'AIDA_16x9', width: L.width, height: L.height });
    pptx.layout = 'AIDA_16x9';
    pptx.theme = { headFontFace: fonts.pptxHeading, bodyFontFace: fonts.pptxBody };
    pptx.title = S.makeTpl(cfg)(cfg.brand.deckTitle);
    pptx.subject = 'Leadership briefing';

    var deck = S.buildDeck(cfg, state || S.initialState(cfg), calc);
    var ch = S.chrome(cfg);

    // Slide model margins follow CSS order [top, right, bottom, left] in points.
    // PptxGenJS 4.0.1 reads a margin array as [left, right, bottom, top].
    function pptMargin(m) {
      if (!m) return 0;
      if (!Array.isArray(m)) return m;
      return [m[3], m[1], m[2], m[0]];
    }
    function textOpts(p) {
      var o = {
        x: p.x, y: p.y, w: p.w, h: p.h,
        fontFace: face(p), fontSize: p.size, color: col(p.color || 'text'),
        bold: !!p.bold, italic: !!p.italic, align: p.align || 'left', valign: p.valign || 'top',
        margin: pptMargin(p.margin), fit: 'none', objectName: p.role === 'title' ? 'Title' : undefined
      };
      if (p.fill) o.fill = { color: col(p.fill) };
      if (p.line) o.line = { color: col(p.line), width: p.lineW || 1, dashType: p.dash ? 'dash' : 'solid' };
      if (p.radius && (p.fill || p.line)) { o.shape = pptx.ShapeType.roundRect; o.rectRadius = p.radius; }
      return o;
    }
    function runs(p) {
      if (!p.runs) return p.text;
      return p.runs.map(function (r) {
        var o = {};
        if (r.bold !== undefined) o.bold = r.bold;
        if (r.italic) o.italic = true;
        if (r.color) o.color = col(r.color);
        if (r.size) o.fontSize = r.size;
        if (r.breakLine) o.breakLine = true;
        return { text: r.text, options: o };
      });
    }

    function masterObjects(variant) {
      var c = S.chrome(cfg, variant), objs = [];
      if (c.footer) {
        objs.push({ rect: { x: c.footerRule.x, y: c.footerRule.y, w: c.footerRule.w, h: c.footerRule.h, fill: { color: col('neutral') }, line: { type: 'none' } } });
        objs.push({ text: { text: c.footerText.text, options: textOpts(c.footerText) } });
      }
      var lg = c.logo;
      if (lg.imagePath) {
        objs.push({ image: { x: lg.x, y: lg.y, w: lg.w, h: lg.h, path: lg.imagePath } });
      } else {
        objs.push({ text: { text: lg.text, options: {
          x: lg.x, y: lg.y, w: lg.w, h: lg.h, fontFace: fonts.pptxBody, fontSize: 11, bold: true, color: col('muted'),
          align: 'center', valign: 'middle', margin: 0, line: { color: col('neutral'), width: 1, dashType: lg.dashed ? 'dash' : 'solid' }
        } } });
      }
      return objs;
    }
    var slideNumber = cfg.theme.footer.showSlideNumber ? {
      x: ch.slideNumber.x, y: ch.slideNumber.y, w: ch.slideNumber.w, h: ch.slideNumber.h,
      fontFace: fonts.pptxBody, fontSize: ch.slideNumber.size, color: col(ch.slideNumber.color), align: 'right', valign: 'middle', margin: 0
    } : undefined;

    function titlePlaceholderFrom(p) {
      return { placeholder: { options: {
        name: 'title', type: 'title', x: p.x, y: p.y, w: p.w, h: p.h,
        fontFace: fonts.pptxHeading, fontSize: p.size, bold: !!p.bold, color: col(p.color || 'text'), align: 'left', valign: p.valign || 'middle', margin: 0
      }, text: '' } };
    }
    function findTitle(sl) { return sl.primitives.filter(function (p) { return p.role === 'title'; })[0]; }

    pptx.defineSlideMaster({ title: MASTER_TITLE, background: { color: col('background') },
      objects: masterObjects('title').concat([titlePlaceholderFrom(findTitle(deck[0]))]) });
    pptx.defineSlideMaster({ title: MASTER_CONTENT, background: { color: col('background') },
      objects: masterObjects('content').concat([titlePlaceholderFrom(findTitle(deck[1]))]), slideNumber: slideNumber });

    deck.forEach(function (sl, i) {
      var slide = pptx.addSlide({ masterName: i === 0 ? MASTER_TITLE : MASTER_CONTENT });
      sl.primitives.forEach(function (p) {
        if (p.type === 'text') {
          var o = textOpts(p);
          if (p.role === 'title') o.placeholder = 'title';
          slide.addText(runs(p), o);
        } else if (p.type === 'rect' || p.type === 'ellipse') {
          var shp = p.type === 'ellipse' ? pptx.ShapeType.ellipse : (p.radius ? pptx.ShapeType.roundRect : pptx.ShapeType.rect);
          slide.addShape(shp, {
            x: p.x, y: p.y, w: p.w, h: p.h, rectRadius: p.radius || undefined,
            fill: p.fill ? { color: col(p.fill) } : { type: 'none' },
            line: p.line ? { color: col(p.line), width: p.lineW || 1, dashType: p.dash ? 'dash' : 'solid' } : { type: 'none' }
          });
        } else if (p.type === 'line') {
          slide.addShape(pptx.ShapeType.line, {
            x: p.x, y: p.y, w: p.w, h: p.h,
            line: { color: col(p.color), width: p.width || 1, dashType: p.dash ? 'dash' : 'solid',
              endArrowType: p.arrow ? 'triangle' : undefined, beginArrowType: p.arrowStart ? 'triangle' : undefined }
          });
        } else if (p.type === 'inputs') {
          // Static snapshot of the current assumptions.
          var rh = p.h / p.rows.length;
          p.rows.forEach(function (r, k) {
            var y = p.y + k * rh;
            slide.addText(r.label, { x: p.x, y: y, w: p.w * 0.7, h: rh, fontFace: fonts.pptxBody, fontSize: 9.5, color: col('muted'), valign: 'middle', margin: 0 });
            slide.addText(String(r.display), { x: p.x + p.w * 0.7, y: y, w: p.w * 0.3, h: rh, fontFace: fonts.pptxBody, fontSize: 11, bold: true, color: col('text'), align: 'right', valign: 'middle', margin: 0 });
          });
        }
      });
      slide.addNotes(sl.notes);
    });
    return pptx;
  }

  var api = { build: build };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.AIDA_PPTX = api;
})(typeof self !== 'undefined' ? self : this);
