# FormaScholar Word Web Prototype — Smoke Test

Applies to **v0.4.5**.

## v0.4.5 conditional review and structure checks

- Applicability is derived from detected authors, objects, equations, tables, references, and submission state.
- The default findings view hides `not_applicable` items; the explicit filter still exposes them for audit.
- Association-marker superscript, equation tab, and table-rule checks use normalized OOXML data.
- Safe fixes for association markers, equation tabs, and table rules preserve manuscript text and are revalidated through the normal validation pass.

## 0.4.4 text and caption normalization

- Superscript affiliation-number and email-letter marker runs are ignored by author italic checks.
- Author italic corrections preserve the leading marker and italicize only the association text.
- Repeated internal spaces are checked conservatively outside tables, equations, and captions.
- Figure/table leading caption whitespace is detected and removed at range level.
- Figure captions use 6 pt before and 0 pt after; table captions use the same template-derived spacing.
- Alignment fixes assign the final target directly.

## 0.4.3 safe correction verification

- `LINE_SPACING_001` is auto-fixable only for its affected paragraph indices.
- The correction uses semantic 12-point single spacing and includes paragraphs inside table cells.
- Verification re-reads OOXML and requires `lineRule=auto`, `line=240`, and normalized multiple `1.0`.
- A partial result must report verified and failed paragraph counts separately.
- Diagnostics identify each offender's paragraph number, location, rule, and raw OOXML line value where available.
- Regression coverage includes `auto+240`, `exact+240`, `atLeast+240`, `exact+200`, inherited style spacing, and table-cell spacing.

## 0.4.2 cross-reference checks

- Numeric citations such as `[1]`, `[1, 2]`, `[1-3]`, and `[1–3, 6]` are expanded deterministically.
- Reference coverage, invalid reference citations, and first-use sequence are reported as consistency findings and do not change the AIP compliance score.
- Figure captions exclude their own label when checking in-text figure citation coverage.
- Table citation coverage is reported as a consistency finding.
- Ordered OOXML blocks pair a table caption above a table and a high-confidence drawing followed by a figure caption; ambiguous relationships are Not evaluated rather than failed.
- Citation and placement findings are not auto-fixable.

## 1. Installation

- Run `START_DEV.bat`.
- Confirm `https://localhost:3000/taskpane.html` responds in a browser.
- In Word on the web, upload `manifest.xml` as a custom add-in.
- Confirm the FormaScholar command opens the task pane.

## 2. AIP template baseline

Open `reference/AIP Template(7).docx` in Word on the web and validate it.

Expected observations:

- The task-pane status should say `Validated with Office.js + OOXML (DOMParser)` or `Validated with Office.js + OOXML (fallback)`.
- `PAGE_WIDTH_001` should be evaluated with detected value **8.5 in**, not `OOXML unavailable`.
- `PAGE_HEIGHT_001` should be evaluated with detected value **11 in**, not `OOXML unavailable`.
- Top/left/right margins should be evaluated at **1.0 in**.
- Bottom margin should be automatically evaluated as **1.00 in** in every document section.
- Author-email italic should be automatically evaluated as **true** and should be selectable for a safe correction.
- Template styles such as Paper Title, Author Name, Abstract and Heading 1–3 should be detected where present.
- Human-judgement requirements should show **Manual review**; final-PDF/copyright requirements should show **Submission checklist**; absent object families should show **Not applicable**. **Not evaluated** should be reserved for a real technical limitation.

Note: the AIP template itself contains instructional/sample text, so it is not expected to produce a perfect manuscript compliance score.

## 3. Deliberate title failure

In a copy of the document:

- Select the title.
- Change it to Arial 12 pt, left aligned.
- Validate.

Expected:

- `TITLE_FONT_001` fails.
- `TITLE_SIZE_001` fails.
- `TITLE_ALIGN_001` fails.
- Clicking Locate selects the title paragraph.
- Selecting those fixable findings and clicking **Apply approved fixes** restores the configured AIP formatting and revalidates.

## 4. Heading failure

Change a Heading 1 paragraph to mixed case and remove bold.

Expected:

- Heading 1 bold check fails.
- Heading 1 all-caps check fails.
- The all-caps correction is selectable.

## 5. Body paragraph failure

Change a normal body paragraph to Calibri 11 pt and left aligned.

Expected:

- Body font, size and alignment rules fail for the affected body paragraph group.
- Safe fixes can be approved and applied.

## 6. OOXML document-level check

Add a page number or header in Word and save/sync the document, then validate again.

Expected:

- Page-number and/or header/footer rules should fail after Word's current document package exposes those changes through `getFileAsync`.

## 7. No References section

Validate a manuscript without a `REFERENCES` heading.

Expected:

- `REFERENCES_REQUIRED_001` fails.
- Rules dependent on reference structure do not pretend to pass.

## 8. Safety check

Approve a formatting correction and verify:

- Only selected deterministic properties are modified.
- FormaScholar does not replace the original file or export a separate copy automatically.
- Word's normal undo history remains available to the user.

## 9. Header/footer false-positive + approved fix

- Validate a document that has no visible header/footer content. Empty header/footer containers should not fail `HEADER_FORBIDDEN_001` or `FOOTER_FORBIDDEN_001`.
- Add visible header text and validate. `HEADER_FORBIDDEN_001` should fail and show a selectable **structural fix**.
- Approve it and apply fixes. Revalidation should pass the header rule. Repeat for a footer.

## 10. Title-case fix

Change the title to lower case while leaving technical tokens such as `YOLOv8`, `5G`, or an all-caps acronym in it.

Expected:

- `TITLE_CASE_001` fails and is selectable.
- Applying the fix title-cases ordinary words while preserving the technical tokens.
- Revalidation passes the title-case rule.

## 11. Figure-caption prefix fix

Give one or more paragraphs styled as Figure Caption an invalid/missing label.

Expected:

- `FIGURE_CAPTION_PREFIX_001` fails and is selectable.
- Applying the fix produces `FIGURE 1.`, `FIGURE 2.`, etc. according to caption order for the failed captions.
- The generated `FIGURE n.` label is bold.

## 12. Reference-block detection and numbering

Open the supplied AIP template. Its seven sample references use the `Reference` style/Word numbering and are followed by additional guidance text.

Expected:

- Only the actual reference block is classified as references; later “Where to Find Further Information” and summary guidance are not counted.
- Correct Word automatic numbering should satisfy `REFERENCES_NUMBERING_001` through the list-number string.
- If numbering is deliberately broken, the rule should become selectable and the approved fix should normalize the detected reference block to sequential `1., 2., 3.` numbering.
## 13. Coverage/status regression test

Validate a manuscript with no equations and no tables.

Expected:

- Equation- and table-specific human-review rules do **not** appear as Not evaluated merely because those objects are absent; they are **Not applicable**.
- Final-PDF and copyright requirements appear under **Checklist** and do not affect compliance score.
- Human/visual requirements appear under **Review** and do not affect compliance score.
- The coverage line only includes a `not evaluated` count if a genuine technical check could not run.

## 14. Newly automated deterministic checks

Using a copy of the AIP template/manuscript, deliberately change one item at a time:

- make only `Abstract.` non-bold;
- make `FIGURE 1.` non-bold;
- change `TABLE 1.` to `Table 1.`;
- break FIGURE or TABLE numbering sequence;
- leave `(Second Level Heading)` in a real heading;
- leave a `Use the Microsoft Word template style:` instruction in the manuscript;
- capitalize a middle minor word such as `With` in a Heading 2/3;
- create Equation-style paragraphs ending `(1)`, `(3)` instead of `(1)`, `(2)`;
- change Conclusion/Acknowledgments body text away from Times New Roman 10 pt.

Expected: the corresponding rules are now automatically evaluated as Passed/Failed instead of Not evaluated.

## 15. v0.3.1 regression checks

- Figure Caption with 6 pt before spacing passes `FIGURE_CAPTION_SPACE_BEFORE_001`.
- Single line spacing resolves OOXML `auto` values such as 240 → 1.0 and 276 → 1.15; exact/atLeast rules fail strict single-spacing validation.
- Paragraphs inside table cells are not classified as normal body paragraphs.
- An email styled as Author Affiliation is classified as `authorEmail` when its text contains an email address.
- `Abstract. text` passes the prefix rule and period rule; `Abstract text` passes only the prefix rule.
- `Abstract.` bold with a normal body passes `ABSTRACT_BODY_NORMAL_001`; an entirely bold abstract fails it.
- A high-confidence AIP instruction phrase fails `TEMPLATE_BOILERPLATE_REMOVED_001`.
- A body PAGE/NUMPAGES field fails `PAGE_NUMBER_FORBIDDEN_001` even when no header/footer field exists.
- A two-section document passes a layout rule only when both sections match; a wrong value in section 2 fails it.
- `Analysis with Deep Learning` does not fail the Heading 2/3 title-case rule or the minor-word rule; `Analysis With Deep Learning` fails only the minor-word rule.
- `References` fails the References all-caps rule while `REFERENCES` passes.
- `Figure 1.` fails the exact prefix rule but passes figure-number sequence.

