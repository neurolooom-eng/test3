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
      keyMessage: 'Quality review where decisions are made — not after documents are assembled.',
      supportingLine: 'Connected evidence + LLM reasoning + deterministic validation.',
      operatingModes: [
        { title: 'LLM-based review', text: 'Content, context, quality, consistency' },
        { title: 'Hybrid review', text: 'LLM reasoning + script validation' },
        { title: 'Script-based validation', text: 'Mappings, completeness, coverage' }
      ],
      solutionAreas: [
        { label: 'Document Reviews', focus: true },
        { label: 'Data Insights', focus: false }
      ],
      notes: '{tool} combines connected evidence, LLM reasoning, and deterministic validation to strengthen release traceability and lifecycle readiness. Two solution areas: Document Reviews and Data Insights. Today we focus on Document Reviews and connected source evidence.'
    },

    /* ------------------------------------------------------------------ */
    useCaseOne: {
      sectionProblem: '1 · The problem',
      headlineProblem: [
        { text: 'One release change ripples across ' },
        { text: 'five linked artefacts', emphasis: true },
        { text: ' — gaps surface only at formal review' }
      ],
      tagline: 'Quality is found late, and found again.',
      titleProblem: 'Use Case 1 — Evidence-based Release Traceability Review',
      titleApproach: 'Use Case 1 — From Formal Review to First Time Right',
      problemStatement: 'When release documents enter formal review, QEs must manually reconcile traceability and compliance across tightly coupled artefacts. This creates repeated review cycles and leaves a risk of missed impact, missing mappings, and inconsistent evidence.',
      // Placeholder — configure actual value.
      reviewCycleRange: '3–6',
      artefacts: [
        { eyebrow: 'Scope', name: 'Release Plan' },
        { eyebrow: 'Requirements', name: 'Requirements Document' },
        { eyebrow: 'Verification', name: 'Test Protocol' },
        { eyebrow: 'Implementation', name: 'Code Review / PR evidence' },
        { eyebrow: 'Traceability', name: 'RTM' }
      ],
      // Gap markers on the chain; `afterArtefact` = index of the artefact before the gap.
      issueLabels: [
        { afterArtefact: 1, text: 'AC with no mapped test' },
        { afterArtefact: 3, text: 'PR not linked to requirement' }
      ],
      loopLabel: 'Returned to author · {cycles} review cycles per document',
      problemCards: [
        { title: 'Impact missed', text: 'Affected documents mis-identified after a change.' },
        { title: 'Mappings missing', text: 'Scope → requirement → test → PR → defect links broken or indirect.' },
        { title: 'Manual, repeated checks', text: 'DHF, GDP and document-quality checks re-run by hand every cycle.' }
      ],
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
        { text: 'From formal review to ' },
        { text: 'First Time Right', emphasis: true }
      ],
      taglineApproach: 'Same QE-defined checks — moved earlier, run by authors.',
      lanes: [
        {
          tag: 'Today · current use', future: false,
          steps: [
            { role: 'Author', text: 'Prepares document' },
            { role: 'QE + {tool}', text: 'Evidence-based review', highlight: true },
            { role: 'Approval', text: 'After review cycles' }
          ]
        },
        {
          tag: 'Next · future adoption', future: true,
          steps: [
            { role: 'Author + {tool}', text: 'Pre-review, before submission', highlight: true },
            { role: 'QE', text: 'Strategy, risk, judgement' },
            { role: 'Approval', text: 'First Time Right' }
          ]
        }
      ],
      currentStateNote: 'Today, QEs use {tool} to make formal review evidence-based—reviewing against connected source evidence, not only the document narrative.',
      futureStateNote: 'The same QE-defined, governed checks can be initiated by authors on partial or complete content before formal QE submission. This is a future / adoption model.',
      outcomeNote: 'Earlier feedback and consistent baseline checks aim to reduce avoidable rework and move formal QE review toward strategy, risk, tailoring, and expert judgement.',
      // Placeholder target — present as an assumption, never a guarantee.
      targetReviewCycles: '[target]',
      targetLabel: 'Target review cycles',
      targetCaveat: 'For eligible checks · assumption, not a guarantee',
      demoSteps: [
        { short: 'Select scope', detail: 'Select an in-scope release / work item set.' },
        { short: 'Connected evidence', detail: 'Show connected evidence for a requirement or change.' },
        { short: 'Run validation', detail: 'Run a {tool} traceability validation.' },
        { short: 'Surface gaps', detail: 'Surface representative gaps: missing mapping, unlinked PR/defect, coverage issue, or DHF/GDP observation.' },
        { short: 'QE action', detail: 'Close with the action a QE takes.' }
      ]
    },

    /* ------------------------------------------------------------------ */
    useCaseTwo: {
      sectionProblem: '2 · The checkpoints',
      headlineProblem: [
        { text: 'Review quality ' },
        { text: 'at the source', emphasis: true },
        { text: ' — while the product can still change' }
      ],
      tagline: 'Source-driven lifecycle quality assurance · already in QE use at multiple stages',
      titleProblem: 'Use Case 2 — Source-driven Lifecycle Quality Assurance',
      titleApproach: 'Use Case 2 — Review Quality at the Source',
      problemStatement: 'Requirement quality, design completeness, and test coverage must be assessed at lifecycle decision points while the product can still be changed—not only after evidence is assembled for formal review.',
      lifecycleLabels: ['Design Input', 'Development', 'Test readiness', 'Verification'],
      changeArrowLabel: 'Room to change the product narrows',
      checkpointLabels: [
        {
          stage: 0,
          title: 'Design Input',
          when: 'Before development',
          question: 'Are requirements clear, testable and INCOSE-aligned?',
          value: 'Ambiguity and missing design'
        },
        {
          stage: 2,
          title: 'Test readiness',
          when: 'Before verification',
          question: 'Is every acceptance criterion covered by a test?',
          value: 'Coverage and traceability gaps'
        }
      ],
      sourcesLabel: '{tool} pulls from',
      sources: ['Requirements', 'Detailed design', 'Acceptance criteria', 'Test protocol / suite'],
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

      sectionApproach: '2 · One setup, reused',
      headlineApproach: [
        { text: 'Configure once. ' },
        { text: 'Reuse at every checkpoint.', emphasis: true }
      ],
      taglineApproach: 'Current QE use — the same governed checks applied across the lifecycle.',
      hubTitle: '{tool} quality setup',
      reusableSetup: [
        { mode: 'LLM-based', text: 'Requirement quality · INCOSE' },
        { mode: 'Hybrid', text: 'Context + deterministic checks' },
        { mode: 'Script-based', text: 'Links · mapping · coverage' }
      ],
      checkpointApplies: [
        'Requirement quality · design needs',
        'AC-level test coverage'
      ],
      demoSteps: [
        { short: 'Design Input', detail: 'Inspect requirement quality and whether detailed design needs are identified.' },
        { short: 'Test readiness', detail: 'Combine requirement plus detailed design and check acceptance-criteria-level test coverage.' }
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
        { from: 'Limited sample, document narrative', to: 'Full-population, evidence-based review', tag: 'Current use', kind: 'current' },
        { from: 'Checks only at formal QE submission', to: 'Author-initiated pre-review', tag: 'Future adoption', kind: 'future' },
        { from: 'QE time on repeatable checks', to: 'QE time on strategy, risk, judgement', tag: 'Target outcome', kind: 'outcome' }
      ],
      pillarNotes: [
        'Evidence-based, full-population review: review the complete applicable release scope against connected source evidence—not a limited sample or document narrative alone.',
        'Author-initiated pre-review: QE-defined, governed checks can be run on partial or complete content before formal QE submission. Future / adoption model where not yet deployed.',
        'First Time Right, strategy-led QE review: repeatable baseline checks are applied consistently, allowing QEs to focus on quality strategy, risk-based decisions, tailoring, and expert judgement.'
      ],
      caveat: 'Not 100% assurance — 100% coverage of the applicable configured checks.',
      bottomLine: '{tool} does not replace QE judgement. It removes repeatable effort so QE expertise goes deeper where it matters.'
    },

    /* ------------------------------------------------------------------ */
    savings: {
      section: '4 · QE capacity',
      headline: [
        { text: 'Releasing QE capacity for ' },
        { text: 'deeper technical review', emphasis: true }
      ],
      title: 'Releasing QE Capacity for Deeper Technical Quality Review',
      // Shows an "Illustrative" tag. Set false once real, validated figures are configured.
      showIllustrativeTag: true,
      illustrativeTag: 'Illustrative · synthetic values',
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
        barsTitle: 'QE hours per release',
        resultReleased: 'QE capacity released per release',
        resultPercent: 'Reduction in QE review hours',
        resultAnnual: 'Annual QE capacity released',
        cyclesTargetNote: 'Target / assumption'
      },
      reinvestmentLead: 'Reinvest into',
      reinvestmentItems: [
        'Detailed technical reviews',
        'Risk-based assessment',
        'Requirements & design',
        'Early lifecycle support'
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
