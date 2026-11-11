'use strict';
const test = require('node:test');
const assert = require('node:assert');
const calc = require('../src/calc.js');
const cfg = require('../config/deck.config.js');
const slides = require('../src/slides.js');

test('calculates QE capacity per the handoff formulas', () => {
  const r = calc.calculateCapacity({
    eligibleDocumentsPerRelease: 10, adoptionRate: 0.5,
    manualHoursPerDocument: 8, aidaHoursPerDocument: 6, releasesPerYear: 3
  });
  assert.strictEqual(r.eligibleDocuments, 5);
  assert.strictEqual(r.currentQeHoursPerRelease, 40);
  assert.strictEqual(r.aidaAssistedQeHoursPerRelease, 30);
  assert.strictEqual(r.qeCapacityReleasedHours, 10);
  assert.strictEqual(r.qeCapacityReleasedPercent, 25);
  assert.strictEqual(r.annualQeCapacityReleasedHours, 30);
});

test('guards divide-by-zero and missing inputs with null (rendered as —)', () => {
  const r = calc.calculateCapacity({ eligibleDocumentsPerRelease: 0, adoptionRate: 1, manualHoursPerDocument: 5, aidaHoursPerDocument: 2 });
  assert.strictEqual(r.currentQeHoursPerRelease, 0);
  assert.strictEqual(r.qeCapacityReleasedPercent, null);
  assert.strictEqual(r.annualQeCapacityReleasedHours, null);
  assert.strictEqual(calc.fmt(r.qeCapacityReleasedPercent), '—');
  const blank = calc.calculateCapacity({ eligibleDocumentsPerRelease: '', adoptionRate: 1, manualHoursPerDocument: null, aidaHoursPerDocument: 2, releasesPerYear: 2 });
  assert.strictEqual(blank.qeCapacityReleasedHours, null);
  assert.strictEqual(blank.qeCapacityReleasedPercent, null);
});

test('scenario resolution applies user scenario and overrides', () => {
  const sv = cfg.savings;
  const a = calc.resolveInputs(sv, 'minor', 'experienced', {});
  const minor = sv.releaseScenarios.find((r) => r.id === 'minor');
  assert.strictEqual(a.aidaHoursPerDocument, minor.aidaHoursPerDocument.experienced);
  assert.strictEqual(a.adoptionRate, sv.adoptionRate);
  const b = calc.resolveInputs(sv, 'minor', 'experienced', { adoptionRate: 1, eligibleDocumentsPerRelease: 20 });
  assert.strictEqual(b.adoptionRate, 1);
  assert.strictEqual(b.eligibleDocumentsPerRelease, 20);
});

test('deck has seven slides, notes, and tool name substitution', () => {
  const deck = slides.buildDeck(cfg, slides.initialState(cfg), calc);
  assert.strictEqual(deck.length, 7);
  deck.forEach((s) => {
    assert.ok(s.notes.length > 20);
    const json = JSON.stringify(s.primitives);
    assert.ok(!json.includes('{tool}'), 'unresolved {tool} on ' + s.id);
    s.primitives.forEach((p) => {
      assert.ok(p.x >= 0 && p.y >= 0 && p.x + p.w <= 13.34 && p.y + p.h <= 7.5, `primitive out of bounds on ${s.id}: ${JSON.stringify(p).slice(0, 120)}`);
    });
  });
});

test('guardrail wording: no "100% sampling" and caveat present', () => {
  const all = JSON.stringify(cfg);
  assert.ok(!/100% sampling/i.test(all.replace(/Never say “100% sampling”/g, '')));
  assert.ok(/not 100% assurance/.test(cfg.value.caveat));
  assert.ok(!/Software Quality Engineer|Design Quality Engineer/.test(all));
});
