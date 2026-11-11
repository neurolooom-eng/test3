/*
 * QE-capacity calculation (pure functions, no DOM).
 * Used by the HTML deck, the PPTX exporter and the unit tests.
 */
(function (root) {
  'use strict';

  function num(v) {
    var n = typeof v === 'string' ? parseFloat(v) : v;
    return typeof n === 'number' && isFinite(n) ? n : null;
  }

  /**
   * @param {object} i inputs: eligibleDocumentsPerRelease, adoptionRate (0–1),
   *   manualHoursPerDocument, aidaHoursPerDocument, releasesPerYear (optional)
   * @returns {object} results; any value that cannot be computed is null
   */
  function calculateCapacity(i) {
    var docs = num(i.eligibleDocumentsPerRelease);
    var adoption = num(i.adoptionRate);
    var manual = num(i.manualHoursPerDocument);
    var aida = num(i.aidaHoursPerDocument);
    var releases = num(i.releasesPerYear);

    var eligibleDocuments = docs !== null && adoption !== null ? docs * adoption : null;
    var currentQeHoursPerRelease = eligibleDocuments !== null && manual !== null ? eligibleDocuments * manual : null;
    var aidaAssistedQeHoursPerRelease = eligibleDocuments !== null && aida !== null ? eligibleDocuments * aida : null;
    var qeCapacityReleasedHours = currentQeHoursPerRelease !== null && aidaAssistedQeHoursPerRelease !== null
      ? currentQeHoursPerRelease - aidaAssistedQeHoursPerRelease : null;
    // Divide-by-zero guard: no baseline → no percentage.
    var qeCapacityReleasedPercent = qeCapacityReleasedHours !== null && currentQeHoursPerRelease > 0
      ? (qeCapacityReleasedHours / currentQeHoursPerRelease) * 100 : null;
    var annualQeCapacityReleasedHours = qeCapacityReleasedHours !== null && releases !== null && releases > 0
      ? qeCapacityReleasedHours * releases : null;

    return {
      eligibleDocuments: eligibleDocuments,
      currentQeHoursPerRelease: currentQeHoursPerRelease,
      aidaAssistedQeHoursPerRelease: aidaAssistedQeHoursPerRelease,
      qeCapacityReleasedHours: qeCapacityReleasedHours,
      qeCapacityReleasedPercent: qeCapacityReleasedPercent,
      annualQeCapacityReleasedHours: annualQeCapacityReleasedHours
    };
  }

  /** Format a number for display; null → em dash. */
  function fmt(v, decimals) {
    if (v === null || v === undefined || !isFinite(v)) return '—';
    var d = decimals === undefined ? (Math.abs(v) >= 100 || Number.isInteger(v) ? 0 : 1) : decimals;
    return v.toLocaleString('en-GB', { minimumFractionDigits: d, maximumFractionDigits: d });
  }

  /** Resolve the effective inputs for a release + user scenario, with overrides. */
  function resolveInputs(savings, releaseId, userId, overrides) {
    var rel = savings.releaseScenarios.filter(function (r) { return r.id === releaseId; })[0] || savings.releaseScenarios[0];
    var user = savings.userScenarios.filter(function (u) { return u.id === userId; })[0] || savings.userScenarios[0];
    var aidaH = rel.aidaHoursPerDocument;
    var base = {
      releaseType: rel.releaseType,
      userScenario: user.label,
      releasesPerYear: rel.releasesPerYear,
      eligibleDocumentsPerRelease: rel.eligibleDocumentsPerRelease,
      baselineReviewCycles: rel.baselineReviewCycles,
      aidaTargetReviewCycles: rel.aidaTargetReviewCycles,
      manualHoursPerDocument: rel.manualHoursPerDocument,
      aidaHoursPerDocument: typeof aidaH === 'object' && aidaH !== null ? aidaH[user.id] : aidaH,
      adoptionRate: rel.adoptionRate !== undefined ? rel.adoptionRate : savings.adoptionRate
    };
    var o = overrides || {};
    Object.keys(o).forEach(function (k) { if (o[k] !== undefined) base[k] = o[k]; });
    return base;
  }

  var api = { calculateCapacity: calculateCapacity, fmt: fmt, resolveInputs: resolveInputs };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.AIDA_CALC = api;
})(typeof self !== 'undefined' ? self : this);
