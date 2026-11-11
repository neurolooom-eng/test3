/*
 * AIDA leadership deck — central configuration
 * =============================================
 * Every piece of editable content lives in this file: wording, system labels,
 * presenters, savings assumptions, disclaimers and the theme.
 * You should not need to edit HTML/CSS/JS for normal updates.
 *
 * Tokens: any string may contain {tool}, which is replaced with brand.toolName,
 * and {presenterOne} / {presenterTwo}, replaced with the presenter names.
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
      deckTitle: '{tool}: Evidence-led Quality Reviews',
      deckSubtitle: 'Quality Engineering · Leadership briefing',
      // Set logoPlaceholder.imagePath to a PNG/SVG path (relative to index.html)
      // to replace the dashed placeholder box. Leave null to show the placeholder.
      logoPlaceholder: { text: 'LOGO', imagePath: null },
      footerText: '{tool} · Quality Engineering · Internal — placeholder content',
      roleLabel: 'Quality Engineers'
    },

    /* ------------------------------------------------------------------ */
    systemLabels: {
      templateRepository: 'ARIS',
      workManagement: 'Azure',
      documentContainer: 'Miro / Azure',
      repository: 'Repository / pull requests',
      defectTracker: 'Defects / bugs'
    },
    // What each system provides (shown as subtle source labels on slide 2).
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
      keyMessage: 'Quality review should happen where decisions are made—not only after documents are assembled.',
      supportingLine: '{tool} combines connected evidence, LLM reasoning, and deterministic validation to strengthen release traceability and lifecycle readiness.',
      operatingModes: [
        { title: 'LLM-based review', text: 'Content, context, quality, and consistency checks.' },
        { title: 'Hybrid review', text: 'LLM reasoning plus Python/script-based validation.' },
        { title: 'Script-based validation', text: 'Deterministic mappings, completeness, coverage, and repeatable checks.' }
      ],
      solutionAreas: [
        { label: 'Document Reviews', focus: true },
        { label: 'Data Insights', focus: false }
      ],
      focusNote: 'Today’s focus: Document Reviews and connected source evidence'
    },

    /* ------------------------------------------------------------------ */
    useCaseOne: {
      titleProblem: 'Use Case 1 — Evidence-based Release Traceability Review',
      titleApproach: 'Use Case 1 — From Formal Review to First Time Right',
      problemStatement: 'When release documents enter formal review, QEs must manually reconcile traceability and compliance across tightly coupled artefacts. This creates repeated review cycles and leaves a risk of missed impact, missing mappings, and inconsistent evidence.',
      // Shown as a display figure on slide 2. Placeholder — configure actual value.
      reviewCycleRange: '3–6',
      reviewCycleCaption: 'review cycles for an average document',
      artefactLabels: [
        'Release Plan',
        'Requirements Document',
        'Test Protocol',
        'Code Review Report / PR evidence',
        'Requirements Traceability Matrix (RTM)'
      ],
      changeTag: 'Release change',
      // Gap markers drawn on the evidence chain. `afterArtefact` is the 0-based
      // index of the artefact the gap sits after (i.e. on the connector to the next).
      issueLabels: [
        { afterArtefact: 1, text: 'Acceptance criterion with no mapped test' },
        { afterArtefact: 3, text: 'PR linked to a task, not the requirement' }
      ],
      keyTakeaway: 'One release change can affect multiple linked artefacts.',
      evidenceRelationship: 'Release scope → requirement → acceptance criteria → test coverage → implementation / pull request → defects → consolidated RTM',
      failurePoints: [
        'Impacted documents are incorrectly identified after a change.',
        'Scope, requirements, acceptance criteria, test, PR, and defect mappings are missing or indirect.',
        'Pull requests are linked to tasks rather than the governing feature or requirement.',
        'Bugs are not linked to requirements or test cases.',
        'DHF, GDP, and document-quality checks are repetitive and manually executed.'
      ],
      currentState: {
        tag: 'Current use',
        title: 'Evidence-based QE review',
        text: 'Today, QEs use {tool} to make formal review evidence-based—reviewing against connected source evidence, not only the document narrative.'
      },
      futureState: {
        tag: 'Future · adoption model',
        title: 'Author-initiated pre-review',
        text: 'The same QE-defined, governed checks can be initiated by authors on partial or complete content before formal QE submission.'
      },
      outcome: {
        tag: 'Target outcome',
        title: 'First Time Right',
        text: 'Earlier feedback and consistent baseline checks aim to reduce avoidable rework and move formal QE review toward strategy, risk, tailoring, and expert judgement.'
      },
      // Placeholder target — present as an assumption, never a guarantee.
      targetReviewCycles: '[target]',
      targetStatement: 'Target: reduce QE review cycles from {current} toward {target} for eligible checks.',
      targetCaveat: 'Target / assumption — not a guaranteed reduction.',
      demoSteps: [
        { short: 'Select release scope', detail: 'Select an in-scope release / work item set.' },
        { short: 'Show connected evidence', detail: 'Show connected evidence for a requirement or change.' },
        { short: 'Run traceability check', detail: 'Run a {tool} traceability validation.' },
        { short: 'Surface gaps', detail: 'Surface representative gaps: missing mapping, unlinked PR/defect, coverage issue, or DHF/GDP observation.' },
        { short: 'QE action', detail: 'Close with the action a QE takes.' }
      ]
    },

    /* ------------------------------------------------------------------ */
    useCaseTwo: {
      titleProblem: 'Use Case 2 — Source-driven Lifecycle Quality Assurance',
      titleApproach: 'Use Case 2 — Review Quality at the Source',
      problemStatement: 'Requirement quality, design completeness, and test coverage must be assessed at lifecycle decision points while the product can still be changed—not only after evidence is assembled for formal review.',
      currentUseTag: 'Current use',
      currentUseText: 'Already applied by QEs at multiple lifecycle stages.',
      lifecycleLabels: ['Release grooming / Design Input', 'Development', 'Test readiness', 'Verification'],
      // Index into lifecycleLabels for each highlighted checkpoint.
      checkpointLabels: [
        {
          stage: 0,
          title: 'Checkpoint 1 · Design Input',
          when: 'Before development begins',
          question: 'Are requirements clear, complete, testable and INCOSE-aligned—with detailed design where needed?',
          value: 'Finds ambiguous requirements and missing design information before implementation.'
        },
        {
          stage: 2,
          title: 'Checkpoint 2 · Test readiness',
          when: 'Before verification begins',
          question: 'Do requirements plus detailed design have sufficient test coverage—traceable to each acceptance criterion?',
          value: 'Finds coverage and traceability gaps before verification execution.'
        }
      ],
      sourceNote: '{tool} pulls from connected sources: requirements, detailed design, acceptance criteria, test protocol / suite.',
      // Code review is NOT part of Use Case 2 today. Keep showOnSlide false unless asked.
      codeReviewStatus: {
        status: 'Not yet tested',
        text: 'Code-review checks are not yet tested for this use case (possible future extension).',
        showOnSlide: false
      },
      approachSubtitle: 'Current QE use: one configured quality setup, reused at each lifecycle checkpoint.',
      reusableSetup: [
        { mode: 'LLM-based', text: 'Requirement quality · content context · INCOSE-aligned review' },
        { mode: 'Hybrid', text: 'Contextual assessment + deterministic validation' },
        { mode: 'Script-based', text: 'Link coverage · mapping · completeness · traceability' }
      ],
      checkpointApplies: [
        'Requirement quality; detailed-design needs identified',
        'Requirement + design vs. tests, at acceptance-criteria level'
      ],
      demoSteps: [
        { short: 'Design Input', detail: 'Inspect requirement quality and whether detailed design needs are identified.' },
        { short: 'Test readiness', detail: 'Combine requirement plus detailed design and check acceptance-criteria-level test coverage.' }
      ]
    },

    /* ------------------------------------------------------------------ */
    value: {
      title: 'What {tool} Changes: From Compliance Effort to Quality Impact',
      pillars: [
        {
          tag: 'Current use',
          future: false,
          title: 'Evidence-based, full-population review',
          text: 'Review the complete applicable release scope against connected source evidence—not a limited sample or document narrative alone.'
        },
        {
          tag: 'Future · adoption model',
          future: true,
          title: 'Author-initiated pre-review',
          text: 'QE-defined, governed checks can be run on partial or complete content before formal QE submission.'
        },
        {
          tag: 'Target outcome',
          future: false,
          title: 'First Time Right, strategy-led QE review',
          text: 'Repeatable baseline checks are applied consistently, allowing QEs to focus on quality strategy, risk-based decisions, tailoring, and expert judgement.'
        }
      ],
      caveat: 'This is not 100% assurance. It is 100% coverage of the applicable configured checks.',
      bottomLine: '{tool} does not replace QE judgement; it removes repeatable review effort so QE expertise can go deeper where it matters.'
    },

    /* ------------------------------------------------------------------ */
    savings: {
      title: 'Releasing QE Capacity for Deeper Technical Quality Review',
      // Shows an "Illustrative — synthetic values" tag. Set false once real,
      // validated figures are configured.
      showIllustrativeTag: true,
      illustrativeTag: 'Illustrative · synthetic placeholder values',
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
        currentBar: 'Current QE hours per release',
        aidaBar: '{tool}-assisted QE hours per release',
        resultReleased: 'QE capacity released per release',
        resultPercent: 'Reduction in QE review hours',
        resultAnnual: 'Annual QE capacity released',
        cyclesTargetNote: 'Target / assumption'
      },
      reinvestmentLead: 'Reinvest released QE capacity into:',
      reinvestmentItems: [
        'detailed technical quality reviews',
        'risk-based assessment',
        'requirements and design quality',
        'early lifecycle quality support'
      ],
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
    // Slide owners and time guides (shown in speaker notes and presenter view).
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
          surface: 'F4F6F9',      // subtle panels
          text: '14213D',         // dark navy — primary text
          muted: '5B6577',        // secondary text
          primary: '14213D',      // navy
          secondary: '33415C',
          accent: '0F766E',       // restrained teal
          accentSoft: 'E6F2F1',
          success: '2F6B3F',
          warning: 'B45309',      // amber — issues / gaps only
          warningSoft: 'FDF3E3',
          neutral: 'C9D0DB',      // connectors, hairlines
          onAccent: 'FFFFFF'
        },
        dark: {
          background: '0F1A2E',
          surface: '18253D',
          text: 'EEF2F7',
          muted: 'A9B4C6',
          primary: 'EEF2F7',
          secondary: 'C9D3E2',
          accent: '4FB8AE',
          accentSoft: '173A40',
          success: '7CC48F',
          warning: 'F2A649',
          warningSoft: '3A2C17',
          neutral: '3C4A63',
          onAccent: '0F1A2E'
        }
      },
      fonts: {
        // HTML font stacks (first available is used).
        headingStack: '"Segoe UI", "Helvetica Neue", Arial, sans-serif',
        bodyStack: '"Segoe UI", "Helvetica Neue", Arial, sans-serif',
        // Single font names written into the .pptx (PowerPoint needs one name).
        pptxHeading: 'Segoe UI',
        pptxBody: 'Segoe UI'
      },
      logoStyle: { x: 11.43, y: 0.36, w: 1.3, h: 0.45, dashed: true },
      diagramStyle: {
        connectorWidth: 1.5,   // points
        cardBorderWidth: 1,    // points
        accentBarHeight: 0.06  // inches
      },
      footer: { showSlideNumber: true }
    }
  };

  if (typeof module === 'object' && module.exports) module.exports = config;
  else root.AIDA_DECK_CONFIG = config;
})(typeof self !== 'undefined' ? self : this);
