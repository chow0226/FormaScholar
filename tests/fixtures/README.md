# DOCX integration fixtures

The local integration fixture is the supplied AIP template at
`reference/AIP Template(7).docx`. It is intentionally excluded from the
public repository because it is reference material supplied with the project.
When the file is available locally, tests load the real DOCX ZIP and run its
`word/document.xml` and `word/styles.xml` through the OOXML formatting parser.
Those two tests are skipped when the private fixture is unavailable, so public
CI can validate the rest of the suite without publishing the DOCX.

Additional manuscript fixtures should be added here as they become available;
the parser tests should assert deterministic findings rather than Office.js
runtime behavior. The supplied `218_ARI_CITIC2025.docx` is not currently
present in the workspace, so its binary integration fixture and the expected
25-paragraph before/after correction assertion still need to be added when the
file is provided.
