/*
 * AIDA leadership deck — central configuration
 * =============================================
 * Every piece of editable content lives in this file: wording, system labels,
 * presenters, savings assumptions, disclaimers and the theme.
 * You should not need to edit HTML/CSS/JS for normal updates.
 *
 * Slides are deliberately light on text. Longer explanations live in
 * `notes` fields and appear in the speaker notes (press N, or in PowerPoint).
 *
 * Tokens: any string may contain {tool} (replaced with brand.toolName) and
 * {presenterOne} / {presenterTwo} (replaced with the presenter names).
 *
 * IMPORTANT: every number below is a SYNTHETIC PLACEHOLDER. Replace with
 * validated internal figures before presenting. Do not commit confidential
 * figures to a shared repository.
 */
(function (root) {
  'use strict';

  var config = {
    /* ------------------------------------------------------------------ */
    brand: {
      toolName: 'AIDA',
      deckTitle: 'Evidence-led Quality Reviews',
      deckSubtitle: 'Quality Engineering · Leadership briefing',
      // Set imagePath to a PNG/SVG path (relative to index.html) to replace the placeholder box.
      logoPlaceholder: { text: 'LOGO', imagePath: null },
      footerText: '{tool} · Quality Engineering · Internal — placeholder content',
      roleLabel: 'Quality Engineers'
    },

    /* ------------------------------------------------------------------ */
    systemLabels: {
      templateRepository: 'ARIS',
      workManagement: 'Azure',
      documentContainer: 'Miro / Azure',
      repository: 'Repository / PRs',
      defectTracker: 'Defects / bugs'
    },
    systemLabelRoles: {
      templateRepository: 'approved templates',
      workManagement: 'release scope, work items, planning',
      documentContainer: 'document container',
      repository: 'implementation evidence',
      defectTracker: 'additional evidence'
    },

    /* ------------------------------------------------------------------ */
    presenters: {
      presenterOne: '[Presenter 1]',
      presenterTwo: '[Presenter 2]'
    },

    /* ------------------------------------------------------------------ */
    opening: {
      keyMessage: 'Quality review where decisions are made.',
      modesLabel: 'Three review modes',
      operatingModes: [
        { title: 'LLM', text: 'Content, context, quality, and consistency checks.' },
        { title: 'Hybrid', text: 'LLM reasoning plus Python/script-based validation.' },
        { title: 'Script', text: 'Deterministic mappings, completeness, coverage, and repeatable checks.' }
      ],
      notes: 'Key message: quality review should happen where decisions are made—not only after documents are assembled.\n\n{tool} combines connected evidence, LLM reasoning, and deterministic validation to strengthen release traceability and lifecycle readiness.\n\nOperating modes:\n- LLM-based review: content, context, quality, and consistency checks.\n- Hybrid review: LLM reasoning plus Python/script-based validation.\n- Script-based validation: deterministic mappings, completeness, coverage, and repeatable checks.\n\nTwo solution areas: Document Reviews and Data Insights. Today we focus on Document Reviews and connected source evidence.'
    },

    /* ------------------------------------------------------------------ */
    useCaseOne: {
      sectionProblem: '1 · The problem',
      headlineProblem: [
        { text: 'Gaps surface late — ' },
        { text: 'and cycle back', emphasis: true }
      ],
      titleProblem: 'Use Case 1 — Evidence-based Release Traceability Review',
      titleApproach: 'Use Case 1 — From Formal Review to First Time Right',
      problemStatement: 'When release documents enter formal review, QEs must manually reconcile traceability and compliance across tightly coupled artefacts. This creates repeated review cycles and leaves a risk of missed impact, missing mappings, and inconsistent evidence.',
      // Placeholder — configure actual value.
      reviewCycleRange: '3–6',
      artefacts: ['Release Plan', 'Requirements', 'Tests', 'Code / PRs', 'RTM'],
      // Gap markers on the chain; `afterArtefact` = index of the artefact before the gap.
      issueLabels: [
        { afterArtefact: 1, text: 'Unmapped test' },
        { afterArtefact: 3, text: 'Unlinked PR' }
      ],
      loopLabel: 'review cycles',
      evidenceRelationship: 'Release scope → requirement → acceptance criteria → test coverage → implementation / pull request → defects → consolidated RTM',
      failurePoints: [
        'Impacted documents are incorrectly identified after a change.',
        'Scope, requirements, acceptance criteria, test, PR, and defect mappings are missing or indirect.',
        'Pull requests are linked to tasks rather than the governing feature or requirement.',
        'Bugs are not linked to requirements or test cases.',
        'DHF, GDP, and document-quality checks are repetitive and manually executed.'
      ],

      sectionApproach: '1 · The shift',
      headlineApproach: [
        { text: 'Check early. ' },
        { text: 'Right first time.', emphasis: true }
      ],
      lanes: [
        { label: 'Today', tag: 'Current use', future: false,
          steps: [{ text: 'Author' }, { text: 'QE + {tool}', highlight: true }, { text: 'Approve' }] },
        { label: 'Next', tag: 'Future', future: true,
          steps: [{ text: 'Author + {tool}', highlight: true }, { text: 'QE' }, { text: 'Approve ✓' }] }
      ],
      currentStateNote: 'Today, QEs use {tool} to make formal review evidence-based—reviewing against connected source evidence, not only the document narrative.',
      futureStateNote: 'The same QE-defined, governed checks can be initiated by authors on partial or complete content before formal QE submission. This is a future / adoption model.',
      outcomeNote: 'First Time Right: earlier feedback and consistent baseline checks aim to reduce avoidable rework and move formal QE review toward strategy, risk, tailoring, and expert judgement.',
      // Placeholder target — present as an assumption, never a guarantee.
      targetReviewCycles: '[target]',
      targetLabel: 'Target cycles',
      demoSteps: [
        'Select an in-scope release / work item set.',
        'Show connected evidence for a requirement or change.',
        'Run a {tool} traceability validation.',
        'Surface representative gaps: missing mapping, unlinked PR/defect, coverage issue, or DHF/GDP observation.',
        'Close with the action a QE takes.'
      ]
    },

    /* ------------------------------------------------------------------ */
    useCaseTwo: {
      sectionProblem: '2 · The checkpoints',
      headlineProblem: [
        { text: 'Review quality ' },
        { text: 'at the source', emphasis: true }
      ],
      titleProblem: 'Use Case 2 — Source-driven Lifecycle Quality Assurance',
      titleApproach: 'Use Case 2 — Review Quality at the Source',
      problemStatement: 'Requirement quality, design completeness, and test coverage must be assessed at lifecycle decision points while the product can still be changed—not only after evidence is assembled for formal review.',
      currentUseTag: 'Current use',
      lifecycleLabels: ['Design Input', 'Development', 'Test readiness', 'Verification'],
      changeArrowLabel: 'Change gets harder',
      checkpointLabels: [
        { stage: 0, title: 'Design Input', question: 'Requirements testable?' },
        { stage: 2, title: 'Test readiness', question: 'Every AC tested?' }
      ],
      sources: ['Requirements', 'Design', 'Acceptance criteria', 'Tests'],
      // Code review is NOT part of Use Case 2 today. Keep showOnSlide false unless asked.
      codeReviewStatus: {
        status: 'Not yet tested',
        text: 'Code-review checks are not yet tested for this use case (possible future extension).',
        showOnSlide: false
      },
      checkpointNotes: [
        'Design Input / before development begins: are requirements clearly written, sufficiently complete and testable, aligned to INCOSE guidance, and supported by detailed design where needed? Finds ambiguous requirements and missing design information before implementation.',
        'Test readiness / before verification begins: do requirements plus detailed design have sufficient coverage in the test protocol or test suite? Is traceability complete at acceptance-criteria level? Finds coverage and traceability gaps before verification execution.'
      ],

      sectionApproach: '2 · One setup',
      headlineApproach: [
        { text: 'Configure once. ' },
        { text: 'Reuse everywhere.', emphasis: true }
      ],
      hubTitle: '{tool}',
      reusableSetup: ['LLM', 'Hybrid', 'Script'],
      demoSteps: [
        'Design Input: inspect requirement quality and whether detailed design needs are identified.',
        'Test readiness: combine requirement plus detailed design and check acceptance-criteria-level test coverage.'
      ]
    },

    /* ------------------------------------------------------------------ */
    value: {
      section: '3 · What changes',
      headline: [
        { text: 'From compliance effort to ' },
        { text: 'quality impact', emphasis: true }
      ],
      title: 'What {tool} Changes: From Compliance Effort to Quality Impact',
      shifts: [
        { from: 'Sample', to: 'Full population', tag: 'Current', kind: 'current' },
        { from: 'At submission', to: 'At authoring', tag: 'Future', kind: 'future' },
        { from: 'Repeat checks', to: 'Expert judgement', tag: 'Target', kind: 'outcome' }
      ],
      pillarNotes: [
        'Evidence-based, full-population review (current use): review the complete applicable release scope against connected source evidence—not a limited sample or document narrative alone.',
        'Author-initiated pre-review (future / adoption model where not yet deployed): QE-defined, governed checks can be run on partial or complete content before formal QE submission.',
        'First Time Right, strategy-led QE review (target outcome): repeatable baseline checks are applied consistently, allowing QEs to focus on quality strategy, risk-based decisions, tailoring, and expert judgement.'
      ],
      caveat: 'This is not 100% assurance. It is 100% coverage of the applicable configured checks.',
      caveatShort: 'Not 100% assurance — 100% coverage of applicable configured checks.',
      bottomLine: '{tool} removes repeatable effort. QE judgement goes deeper.',
      bottomLineNote: '{tool} does not replace QE judgement; it removes repeatable review effort so QE expertise can go deeper where it matters.'
    },

    /* ------------------------------------------------------------------ */
    savings: {
      section: '4 · QE capacity',
      headline: [
        { text: 'Releasing ' },
        { text: 'QE capacity', emphasis: true }
      ],
      title: 'Releasing QE Capacity for Deeper Technical Quality Review',
      // Shows an "Illustrative" tag. Set false once real, validated figures are configured.
      showIllustrativeTag: true,
      illustrativeTag: 'Illustrative',
      defaultReleaseScenario: 'major',
      defaultUserScenario: 'typical',
      // Applicable adoption / scope (0–1).
      adoptionRate: 0.8,
      // All values are SYNTHETIC placeholders.
      releaseScenarios: [
        {
          id: 'major', releaseType: 'Major',
          releasesPerYear: 2,
          eligibleDocumentsPerRelease: 12,
          baselineReviewCycles: '3–6',
          aidaTargetReviewCycles: '[target]',
          manualHoursPerDocument: 10,
          aidaHoursPerDocument: { new: 8, typical: 6.5, experienced: 5.5 }
        },
        {
          id: 'minor', releaseType: 'Minor',
          releasesPerYear: 4,
          eligibleDocumentsPerRelease: 8,
          baselineReviewCycles: '3–6',
          aidaTargetReviewCycles: '[target]',
          manualHoursPerDocument: 6,
          aidaHoursPerDocument: { new: 5, typical: 4, experienced: 3.5 }
        },
        {
          id: 'patch', releaseType: 'Patch',
          releasesPerYear: 6,
          eligibleDocumentsPerRelease: 4,
          baselineReviewCycles: '3–6',
          aidaTargetReviewCycles: '[target]',
          manualHoursPerDocument: 3,
          aidaHoursPerDocument: { new: 2.5, typical: 2, experienced: 1.75 }
        }
      ],
      userScenarios: [
        { id: 'new', label: 'New user' },
        { id: 'typical', label: 'Typical user' },
        { id: 'experienced', label: 'Experienced user' }
      ],
      labels: {
        releaseType: 'Release type',
        userScenario: 'User scenario',
        releasesPerYear: 'Releases per year',
        eligibleDocumentsPerRelease: 'Eligible documents per release',
        baselineReviewCycles: 'Current QE review cycles',
        aidaTargetReviewCycles: '{tool}-assisted target cycles',
        manualHoursPerDocument: 'Current QE effort per document (h)',
        aidaHoursPerDocument: '{tool}-assisted QE effort per document (h)',
        adoptionRate: 'Applicable adoption / scope (%)',
        currentBar: 'Today',
        aidaBar: 'With {tool}',
        resultReleased: 'QE hours released per release',
        resultPercent: 'fewer QE review hours',
        resultAnnual: 'released per year',
        assumptionsButton: 'Assumptions',
        cyclesTargetNote: 'Target / assumption'
      },
      reinvestmentLead: 'Reinvest in',
      reinvestmentItems: [
        'Technical reviews',
        'Risk assessment',
        'Design quality',
        'Early lifecycle'
      ],
      disclaimerShort: 'Indicative estimate, not a headcount claim · varies with complexity, documents, rule coverage, data quality, proficiency and adoption · assumes human review of outputs.',
      disclaimerLines: [
        'This is an indicative QE-capacity estimate, not a headcount-reduction claim. Results vary with release complexity, document count, rule coverage, source-data quality, user proficiency, and adoption.',
        'The calculation reflects only applicable {tool}-assisted checks and assumes appropriate human review of outputs.'
      ]
    },

    /* ------------------------------------------------------------------ */
    speakerNotes: {
      handoverAfterSlide3: 'That shows how {tool} makes the formal traceability review evidence-led. Next, we will show how the same quality logic can be applied at lifecycle checkpoints—before the review package is assembled.'
    },

    /* ------------------------------------------------------------------ */
    // Slide owners and time guides (speaker notes and slide header).
    agenda: [
      { owner: 'presenterOne', time: '0:30–0:45' },
      { owner: 'presenterOne', time: '1:00' },
      { owner: 'presenterOne', time: '2:30–3:00' },
      { owner: 'presenterTwo', time: '1:00' },
      { owner: 'presenterTwo', time: '2:30–3:00' },
      { owner: 'both', time: '1:00–1:15' },
      { owner: 'presenterOne', time: '1:00–1:15' }
    ],

    /* ------------------------------------------------------------------ */
    theme: {
      // 'light' or 'dark' — the toolbar can toggle at runtime.
      activePalette: 'light',
      palettes: {
        light: {
          background: 'FFFFFF',
          surface: 'F2F5FA',      // cool blue-grey panels
          text: '0B1F3A',         // deep navy — primary text
          muted: '5B6B84',        // secondary text
          primary: '0B2A5B',      // navy
          secondary: '2E4366',
          accent: '1D4ED8',       // royal blue — lead colour
          accentSoft: 'E7EEFC',
          accentLine: 'B9CBF5',
          success: '1F6F43',
          warning: 'C8102E',      // red — gaps, risks, one emphasis
          warningSoft: 'FCECEE',
          neutral: 'D5DCE7',      // connectors, hairlines
          onAccent: 'FFFFFF',
          panel: '0B2A5B',        // dark navy panels
          onPanel: 'FFFFFF',
          onPanelMuted: 'AFC0E0'
        },
        dark: {
          background: '07142D',
          surface: '0F2145',
          text: 'EEF3FB',
          muted: 'A3B3CF',
          primary: 'EEF3FB',
          secondary: 'C9D5EA',
          accent: '5B8DEF',
          accentSoft: '132B5C',
          accentLine: '2C4A86',
          success: '6FCF97',
          warning: 'FF5C70',
          warningSoft: '3A1522',
          neutral: '263B66',
          onAccent: 'FFFFFF',
          panel: '13295A',
          onPanel: 'FFFFFF',
          onPanelMuted: 'AFC0E0'
        }
      },
      fonts: {
        // HTML font stacks (first available is used).
        headingStack: '"Segoe UI", "Helvetica Neue", Arial, sans-serif',
        bodyStack: '"Segoe UI", "Helvetica Neue", Arial, sans-serif',
        // Single font names written into the .pptx.
        pptxHeading: 'Segoe UI',
        pptxBody: 'Segoe UI'
      },
      logoStyle: { x: 11.43, y: 0.36, w: 1.3, h: 0.45, dashed: true },
      diagramStyle: {
        connectorWidth: 1.75,  // points
        cardRadius: 0.1        // inches (rounded corners on cards and flow boxes)
      },
      footer: { showSlideNumber: true }
    }
  };

  if (typeof module === 'object' && module.exports) module.exports = config;
  else root.AIDA_DECK_CONFIG = config;
})(typeof self !== 'undefined' ? self : this);
