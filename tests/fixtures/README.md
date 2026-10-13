# DOCX integration fixtures

The first checked-in integration fixture is the supplied AIP template at
`reference/AIP Template(7).docx`. Tests load the real DOCX ZIP and run its
`word/document.xml` and `word/styles.xml` through the OOXML formatting parser.

Additional manuscript fixtures should be added here as they become available;
the parser tests should assert deterministic findings rather than Office.js
runtime behavior. The supplied `218_ARI_CITIC2025.docx` is not currently
present in the workspace, so its binary integration fixture and the expected
25-paragraph before/after correction assertion still need to be added when the
file is provided.
