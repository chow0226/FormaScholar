# FormaScholar — Word on the web Add-in Prototype

**Prototype version: 0.4.5**

FormaScholar is a **Microsoft Word task-pane add-in** that validates the document currently open in **Word on the web** against deterministic publication-formatting rules.

Application release: **v0.4.5**. The Office XML manifests use version **1.0.0.0** for manifest-validation compatibility, and the add-in requires **WordApi 1.3**.

v0.4 focuses on detection accuracy: deterministic formatting is resolved from
DOCX OOXML through one effective-formatting resolver, with source provenance
available for diagnostics. Office.js remains responsible for locating and
editing content in Word.

This prototype is focused on the **AIP Conference Proceedings September 2023 template** supplied with the project. It does not use Python, Streamlit, a database, a REST backend, or an LLM.



## v0.4.5 conditional review and structure checks

v0.4.5 adds document-feature applicability so irrelevant manual and submission
checks are reported as not applicable and hidden from the default findings view.
It also automates author association superscript checks, equation tab detection,
and OOXML table top/bottom/header-rule checks with safe correction paths.

## v0.4.4 text and caption normalization

Author marker runs such as superscript affiliation numbers and email letter markers are excluded from italic-body validation and are preserved by the dedicated italic correction. Caption spacing, caption leading whitespace, and conservative repeated internal-space checks are deterministic and separately fixable. Alignment corrections set the requested final alignment directly; they do not perform an unnecessary left-alignment intermediate step.

## v0.4.3 safe correction verification

`LINE_SPACING_001` is a deterministic safe-fix rule. FormaScholar changes only the affected paragraphs, including table-cell paragraphs, then re-reads the document through OOXML. A correction is counted as verified only when the effective spacing resolves to `w:lineRule="auto"` and `w:line="240"`; remaining `exact` or `atLeast` paragraphs are reported as failed verification.

## v0.4.2 cross-reference and object-structure validation

This revision adds one OOXML-derived cross-reference model shared by the validation rules. It extracts numeric reference citations and first-use order, figure/table citations, ordered document blocks, and high-confidence figure/table caption placement. Figure citation/placement checks remain strict AIP checks; reference coverage, reference sequence, and table citation checks are reported as consistency checks and are excluded from the compliance score. Content citations are never auto-fixed.

## v0.3.1 validation coverage + status model

This revision fixes the misleading wall of **Not evaluated** results. FormaScholar now distinguishes between:

- **Automated** rules — deterministically checked and scored as Passed/Failed.
- **Manual review** — requirements that need human judgement or visual confirmation; excluded from the score.
- **Submission checklist** — requirements concerning the final PDF, copyright evidence, or other material outside the open DOCX; excluded from the score.
- **Not applicable** — the manuscript contains no matching object/section.
- **Not evaluated** — reserved for an actual technical limitation or unavailable document data.

Important engine fix: unimplemented paragraph/object rules are now tested for applicability first. For example, a manuscript with no equations no longer shows a large set of equation rules as Not evaluated; those rules become **Not applicable**.

Additional deterministic checks implemented in this revision include:

- exact run-level bold detection for `Abstract.`, `FIGURE n.`, and `TABLE n.` prefixes;
- abstract-period validation;
- removal of leftover `(First/Second/Third Level Heading)` template guidance;
- removal of leftover `Use the Microsoft Word template style:` / font-instruction text;
- Heading 2/3 lowercase minor-word validation;
- sequential figure and table numbering;
- uppercase TABLE label validation;
- equation-number presence, parentheses, and sequence when Equation-style paragraphs are detected; and
- Conclusion / Acknowledgments body Times New Roman 10 pt checks.
- AIP single-line-spacing validation across every meaningful manuscript paragraph;
- Figure-caption 6 pt before spacing;
- OOXML/style-aware line-spacing resolution and table-cell exclusion;
- effective-format tri-state handling when Word cannot expose a value;
- References Heading 1 capitalization and spacing checks;
- author-email italic validation and the strategy-selected 1.00 in bottom margin;
- multi-section page-size, margin and column validation;
- PAGE/NUMPAGES field detection in the document body as well as headers/footers;
- separate Abstract-prefix and Abstract-period findings;
- Abstract remaining-text bold detection; and
- high-confidence AIP template-boilerplate detection.

The task-pane summary now reports **automated scored**, **review**, **checklist**, **not applicable**, and only shows a **not evaluated** count when a real technical limitation remains. Historical source conflicts are resolved explicitly in the active strategy and remain auditable in the raw metadata.

## v0.1.2 correction + classification fix

This revision makes the five previously non-selectable failed findings actionable and also fixes two false-positive sources:

- **Headers/footers:** empty Word header/footer containers are no longer treated as violations. FormaScholar now inspects the actually referenced header/footer OOXML parts for meaningful text, fields, drawings, or objects. If real content exists, the finding can be selected and the approved fix clears all primary/first-page/even-page headers or footers across every section.
- **Title capitalization:** `TITLE_CASE_001` can now be approved. The fixer uses conservative deterministic title case and preserves acronyms, mixed-case technical terms, and tokens containing digits.
- **Figure-caption prefix:** `FIGURE_CAPTION_PREFIX_001` can now be approved. Invalid caption labels are rewritten to `FIGURE n.` according to caption order, and the label is bolded.
- **Reference detection:** paragraphs are no longer classified as references simply because they appear after the `REFERENCES` heading. FormaScholar now identifies the contiguous reference block using the Reference style, explicit numeric prefixes, or Word list numbering, and stops when that block ends.
- **Reference numbering:** `REFERENCES_NUMBERING_001` now understands Word automatic list numbers through `listString`. When a real numbering failure is approved, the reference block is normalized to sequential `1., 2., 3.` text numbering.

The task pane labels these approved operations as **structural fix**, **text fix**, or **safe fix** instead of treating every correction as the same type.

## v0.1.1 OOXML layout parsing fix

This revision fixes the `OOXML unavailable / Unable to parse OOXML part` issue seen in Word on the web for page-width, page-height and margin checks. The DOCX reader now:

- accepts Word slice data as `Uint8Array`, `ArrayBuffer`, any typed-array view, or number arrays;
- verifies that the returned file is actually a ZIP/DOCX package before parsing;
- retries XML parsing after removing only XML-1.0-forbidden control characters;
- falls back to a small namespace-prefix-agnostic section-property parser when browser `DOMParser` rejects `word/document.xml`;
- scans all sections for header/footer references;
- detects PAGE/NUMPAGES fields more narrowly instead of matching arbitrary footer text; and
- reports an individual layout rule as **Not evaluated** rather than a false failure when that specific OOXML value is genuinely absent.

When validation succeeds, the task-pane status now shows whether OOXML used `DOMParser` or the fallback parser.

## Current prototype

The add-in contains a catalogue of **174 AIP Conference Proceedings rules**. Of these, **152 rules are implemented in the validation engine**, while **111 rules support automatic correction** where a deterministic and safe fix is available. The remaining requirements are primarily manual-review, submission/final-PDF, or other rules that cannot be safely evaluated or modified from the open Word document.

Rules are separated into automated scoreable checks, manual review, submission checklist, historical source-conflict metadata, not applicable, consistency checks, and genuine technical not-evaluated cases. The compliance score only includes applicable compliance rules that were actually evaluated as Passed/Failed. The current strategy contains **20 consistency/advisory rules**, which remain visible in findings but are kept separate from the main compliance score.

Implemented automatic checks include:

- Letter page size (8.5 × 11 in)
- All four section margins from DOCX OOXML, including the strategy-selected 1.00-inch bottom margin
- Multi-section page size, margin and column validation
- Single-column layout
- Header/footer/page-number prohibition
- Title typography, alignment, capitalization and non-all-caps presentation
- Author name typography, alignment and 18 pt before/after spacing
- Affiliation and author-email typography
- Abstract presence, typography, alignment, 0.2 in left/right indents, prefix/period, and run-level bold checks
- Body font, size, justification and first-line indentation
- Heading 1–3 font, size, emphasis, alignment, spacing, capitalization, and Heading 2/3 minor-word checks
- Figure/table caption typography, label syntax, run-level bold prefix checks, label capitalization, and numbering sequence
- Equation paragraph centering, number presence/parentheses/sequence, and center/right tab-stop validation
- References heading detection/formatting
- Reference numbering syntax/sequence and hanging indentation
- Reference citation validity, coverage and first-use order as separate consistency checks
- Figure citation coverage and prior-citation order, plus high-confidence figure-caption placement
- High-confidence table-caption placement and table citation consistency checks
- OOXML table top, bottom and header-border validation with supported correction paths
- Ordered OOXML paragraph/table/drawing blocks with ambiguous object relationships left Not evaluated
- Required Abstract and References checks, plus consistency-only Introduction, Methods/approach, Results/evaluation and Conclusion section-family checks
- First-level section recognition that excludes explicit Heading 2/3 paragraphs from major-section presence checks
- Detection of figure captions after REFERENCES
- Leftover template guidance/style-instruction detection
- Conclusion/Acknowledgments body typography checks

## Post-demo refinements

The active AIP strategy includes the following supervisor-approved refinements:

- US Letter correction sets the predefined Letter paper size for every section while preserving portrait/landscape orientation and existing margins. The correction requires WordApiDesktop 1.3; unsupported hosts receive a manual-review finding. Revalidation warns that paper-size changes can affect pagination, table wrapping and figure placement.
- References are inspected through both the paragraph-style layer and the automatic-list layer. The required `Reference` style is Times New Roman 10 pt, justified, based on `Paragraph`, followed by `Reference`, with a 0 in left indent and 0.3 in hanging indent. List numbering is genuine Word automatic Arabic numbering at level 1, with the number at 0.25 in and text at 0.5 in. Manual prefixes are removed only from their exact leading range; reference bodies, fields, hyperlinks and mixed emphasis are not rewritten. Reference correction requires the custom styles to be available and WordApi 1.6; otherwise it is withheld for manual review.
- Manuscript figure citations always use `Figure` or `Figures`; captions and protected field-backed citations are left unchanged or reported for manual review. A separate `FIGURE_CITATION_SEQUENCE_001` check verifies first mentions in numerical order, ignores repeats, and never renumbers or relocates figures automatically.
- The affiliation/email separator accepts exactly one effective separator: either one safe blank paragraph or one leading ordinary manual line break in the first email paragraph. Duplicate blanks, multiple leading breaks, page breaks, fields, drawings, and uncertain OOXML mappings are withheld for manual review. The existing 0 pt before/after spacing rules remain independent and do not insert another separator.

Reference-style correction is intentionally withheld when reference boundaries, leading-number ranges or citation-manager fields cannot be mapped safely. Spacing before/after references and line spacing remain template-dependent and are not inferred from the supervisor screenshot.

The correction engine applies only explicitly approved deterministic fixes directly to the open Word document. Supported corrections include typography, paragraph formatting, indentation, spacing, line spacing, alignment, page margins, headers/footers, page-number fields, caption formatting, numbering, equation tab stops, selected structural spacing issues, and supported OOXML table-border corrections.

After a correction batch, the add-in automatically revalidates the document. Where appropriate, the corrected property is re-read from Word or OOXML so a correction is counted as successful only when the expected result is verified.

## AIP strategy decisions

The supplied source materials contain two historical differences. FormaScholar’s active AIP strategy resolves them explicitly:

- `MARGIN_BOTTOM_001` validates **1.00 in** from the supplied DOCX section configuration.
- `AUTHOR_EMAIL_ITALIC_001` validates **italic** from the written author-email instruction and can be safely corrected.

The raw extracted evidence in `reference/aip-template-derived-metadata.json` is retained unchanged for audit/history. The strategy file is the source of active validation behaviour.

## Architecture

```text
Word on the web
      │
      ├── Office.js
      │     ├── document paragraphs
      │     ├── styles / text
      │     ├── font properties
      │     ├── paragraph properties
      │     └── live correction / selection
      │
      └── getFileAsync(Compressed)
            │
            └── JSZip + OOXML inspection
                  ├── effective formatting
                  ├── page size and section margins
                  ├── columns
                  ├── headers / footers / page fields
                  ├── paragraph and style properties
                  ├── figures and tables
                  ├── table borders
                  ├── equations and tab stops
                  ├── document object ordering
                  └── cross-reference relationships

                ↓
        ValidationEngine
          ↙              ↘
 Built-in AIP Strategy    Imported Strategy
          \              /
                   ↓
        Findings + Score
                   ↓
       User-approved fixes
                   ↓
        CorrectionService
                   ↓
      Automatic revalidation
```

The Word document itself is the viewer. Clicking **Locate** selects the relevant paragraph in Word rather than rendering an HTML copy of the DOCX.

## Requirements

- Microsoft 365 account with access to Word on the web
- Node.js (current LTS recommended)
- A browser supported by Microsoft 365

No Microsoft Word desktop installation is required for the target workflow.

## Run locally

### Option A — Windows shortcut

Double-click:

```text
START_DEV.bat
```

The script installs npm dependencies when needed and starts the HTTPS development server at:

```text
https://localhost:3000
```

The project uses Microsoft's Office development certificate helper so localhost can be served over trusted HTTPS.

**VS Code is not required to run FormaScholar.** `START_DEV.bat` provides a convenient Windows development launcher. Node.js is required only for this local development workflow.

### Option B — command line

```bash
npm install
npm run dev
```

Keep that terminal open while testing the add-in.

## Sideload into Word on the web

1. Start the local server with `START_DEV.bat` or `npm run dev`.
2. Open a `.docx` file in Word on the web.
3. In Word, go to **Home > Add-ins** and open the advanced / more settings view.
4. Choose **Upload My Add-in**.
5. Select this project's `manifest.xml`.
6. Upload it.
7. Open the **FormaScholar** tab/button in the Word ribbon and choose **Open FormaScholar**.
8. Click **Validate current document**.

Microsoft may label the Add-ins menu slightly differently depending on the tenant/UI version, but the development action is **Upload My Add-in** using the XML manifest.

## Typical workflow

```text
Open manuscript in Word on the web
        ↓
Open FormaScholar
        ↓
AIP Conference Proceedings — September 2023
        ↓
Validate current document
        ↓
Compliance score + findings
        ↓
Click Locate to jump to an issue
        ↓
Approve safe fixes
        ↓
Apply approved fixes
        ↓
Automatic revalidation
```

## Template Rules

FormaScholar includes a **Template Rules** workspace for analysing an updated AIP DOCX template.

The template analyser extracts measurable formatting properties such as fonts, font sizes, paragraph alignment, indentation, spacing and page-layout settings, then compares them with the built-in AIP strategy. Proposed rule changes are presented for review; users can accept selected proposals and create a separate **imported strategy**.

Template analysis never overwrites the built-in AIP strategy. The protected default remains available, users can switch between built-in and imported strategies, and an imported strategy can be removed when it is no longer needed. Imported strategies are stored locally in the browser environment rather than synchronized through a remote database.

## Safe-correction design

FormaScholar follows a conservative correction policy:

1. Validate the manuscript first.
2. Show the detected issue, expected value and rule evidence.
3. Allow the user to select only supported fixable findings.
4. Apply only explicitly approved deterministic corrections.
5. Re-read and revalidate the manuscript after correction.
6. Report any correction that cannot be verified instead of assuming success.

FormaScholar does not blindly reformat an entire manuscript or automatically alter academic content.

### Operation safety

Validation, template analysis and correction operations use a shared busy state to prevent conflicting actions. Relevant controls and finding selections are locked while an operation is running, and correction batches use single-flight protection to prevent duplicate execution.

A correction is also blocked when the validated document or active strategy no longer matches the state used to produce the current findings. The manuscript must be validated again first.

## Rule statuses

- **Passed** — automated rule was evaluated and matched.
- **Failed** — automated rule was evaluated and did not match.
- **Manual review** — requires human judgement/visual confirmation, or represents a source conflict; excluded from the score.
- **Submission checklist** — must be confirmed outside the open DOCX (for example final-PDF checks or copyright permissions); excluded from the score.
- **Not applicable** — no matching object/section exists in this manuscript.
- **Not evaluated** — reserved for a genuine technical limitation or unavailable document data.

This separation prevents unsupported checks from looking like unfinished validator failures and prevents manual/submission requirements from inflating the automated compliance score.

## Project structure

```text
FormaScholar_Word_Web_Addin/
├── manifest.xml
├── package.json
├── webpack.config.js
├── tsconfig.json
├── START_DEV.bat
├── BUILD.bat
├── VALIDATE_MANIFEST.bat
├── assets/
├── reference/
│   ├── AIP Template(7).docx
│   └── aip-template-derived-metadata.json
└── src/
    ├── core/
    │   ├── models.ts
    │   ├── utils.ts
    │   ├── ValidationEngine.ts
    │   ├── CorrectionService.ts
    │   └── StrategyStore.ts
    ├── ooxml/
    │   └── DocxInspector.ts
    ├── strategies/
    │   └── aip-2023.json
    ├── taskpane/
    │   ├── taskpane.html
    │   ├── taskpane.css
    │   └── taskpane.ts
    ├── template/
    │   └── TemplateAnalyzer.ts
    └── word/
        └── WordDocumentAdapter.ts
```

The test suite is organized as:

```text
tests/
├── ValidationEngine.test.ts
├── StrategyStore.test.ts
├── TemplateAnalyzer.test.ts
├── BusyState.test.ts
└── StrategyDeletion.test.ts
```

## Build

Run:

```text
BUILD.bat
```

or:

```bash
npm run typecheck
npm run build
```

The static production files are generated into `dist/`.

## Verification and tests

The current project contains **252 automated tests across 5 test files**, covering validation behaviour, strategy management, template analysis, deletion/protection behaviour, and application busy-state/concurrency handling.

Run:

```bash
npm run typecheck
npm test
npm run build
```

Before release, all three commands should complete successfully.

## Production deployment

The add-in has **no server-side backend**, but an Office web add-in still needs its HTML/JS/CSS to be hosted on an **HTTPS web host**.

For deployment:

1. Run `npm run build`.
2. Host `dist/` on an HTTPS static host (for example an organization web server or Azure Static Web Apps).
3. Copy `manifest.production.template.xml` to `manifest.production.xml`.
4. Replace every `https://YOUR-HOST.example.com` value with the actual HTTPS origin.
5. Deploy/sideload that production manifest through your Microsoft 365 organization or the appropriate Office add-in distribution route.

There is no document upload service: validation still happens from the task pane against the open Word document.

In a production deployment, end users would not need Node.js, VS Code or a command-line environment. The compiled FormaScholar web assets would be hosted over HTTPS and users would access the add-in directly from Microsoft Word.

## Prototype limitations

- FormaScholar currently focuses primarily on the **AIP Conference Proceedings September 2023** formatting strategy. Other publishers such as IEEE, Springer and Elsevier do not yet have built-in strategies.
- Requirements that require semantic or visual judgement—such as whether an author is associated with the correct affiliation, whether an image is sufficiently readable, or whether copyright permission has been obtained—remain **Manual review** or **Submission checklist** items.
- FormaScholar deliberately does not automatically modify manuscript content when doing so could change academic meaning. Only deterministic formatting and structural corrections with sufficiently safe behaviour are offered as automatic fixes.
- Target detection relies on Word styles, effective formatting, document position and deterministic text/structural heuristics. Highly unconventional or severely malformed manuscripts may therefore require manual review.
- Figure/table/object associations are evaluated only when FormaScholar can establish a sufficiently reliable relationship. Ambiguous OOXML mappings are reported conservatively rather than guessed.
- Final-PDF requirements such as PDF font embedding, final visual reflow, file-size constraints and copyright evidence cannot be fully validated from the currently open DOCX.
- Template analysis can infer measurable DOCX formatting properties but cannot automatically understand every semantic publication instruction expressed in natural language.
- Imported strategies are currently stored locally rather than synchronized across devices or users.
- The current project is a prototype/development deployment. A production version still requires hosted HTTPS assets and appropriate Microsoft 365 add-in deployment.
- Additional telemetry and error logging could be added for production deployment, subject to privacy requirements.

## Reference material

The `reference/` folder contains the supplied AIP DOCX and a small JSON snapshot of the actual DOCX style/section properties used to construct the strategy.
