#!/usr/bin/env node
/*
 * Command-line PPTX export (same builder as the in-browser "Export PPTX" button).
 *
 *   npm run export:pptx
 *   node scripts/export-pptx.js --out dist/AIDA.pptx --release minor --user experienced --palette dark
 */
'use strict';
const path = require('path');
const fs = require('fs');
const PptxGenJS = require('pptxgenjs');
const cfg = require('../config/deck.config.js');
const calc = require('../src/calc.js');
const slides = require('../src/slides.js');
const { build } = require('../src/pptx-builder.js');

function arg(name, def) {
  const i = process.argv.indexOf('--' + name);
  return i > -1 && process.argv[i + 1] ? process.argv[i + 1] : def;
}

const state = slides.initialState(cfg);
state.releaseId = arg('release', state.releaseId);
state.userId = arg('user', state.userId);
cfg.theme.activePalette = arg('palette', cfg.theme.activePalette);

const out = path.resolve(arg('out', path.join('dist', cfg.brand.toolName + '-leadership-deck.pptx')));
fs.mkdirSync(path.dirname(out), { recursive: true });

const pptx = build(PptxGenJS, cfg, state, calc, slides);
pptx.writeFile({ fileName: out }).then((f) => console.log('Wrote ' + f));
