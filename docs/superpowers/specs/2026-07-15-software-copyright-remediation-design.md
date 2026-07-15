# Software Copyright Remediation Design

Date: 2026-07-15

## Objective

Rebuild the software copyright submission materials from the current working tree. The code identification material must demonstrate functional originality instead of presenting HTML layout as the software's main body.

## Constraints

- Do not invent, rewrite, or selectively edit source lines for the filing.
- Freeze the current working tree with file hashes before generating materials.
- Preserve full selected files and remove only blank-only lines during extraction.
- Produce A4 code materials whose rendered physical pages are exactly 30 plus 30.
- Keep the software name, version, completion date, source line count, and page count consistent across all artifacts.
- Stop at every user-confirmation gate required by the software-copyright workflow.

## Selected Approach

Use a dual-ended functional ordering. Functional Python code occupies both submitted ends, while HTML templates are concentrated in the omitted middle of the complete source stream.

### Ordered Source Stream

1. `extract_references.py`
2. `db_lookup.py`
3. `lookup_authors.py`
4. `send_emails.py`
5. `tests/test_multi_journal.py`
6. `templates/extract.html`
7. `templates/lookup.html`
8. `templates/notification.html`
9. `main.py`
10. `app.py`
11. `tools/installer_build.py`
12. `tests/test_extract_references.py`
13. `tests/test_installer_build.py`
14. `启动系统.bat`
15. `installer.iss`
16. `build_installer.ps1`

The current estimate is 3,987 nonblank source lines plus 16 trace markers, producing a 4,003-line material stream and 81 logical pages. The submitted front is lines 1-1,500. Lines 1,501-2,550 are omitted. The submitted back is lines 2,551-4,003.

The front ends inside `tests/test_multi_journal.py`. All three HTML files and `main.py` fall inside the omitted middle. The back starts inside `app.py` and contains application, build, test, installer, and launcher logic. The final logical page has three lines; rendered physical pagination remains a separate acceptance gate.

## Material Changes

- Before rebuilding any Markdown, confirm the frozen code baseline, version number, and actual completion date.
- Replace the existing code selection and regenerate the extraction manifest.
- Generate a source-scope manifest for every candidate path with inclusion decision, reason, SHA-256, physical lines, and nonblank lines. The application source count must use the confirmed total from this manifest, not the 3,987 identification-material lines or the old 14,440 scan.
- Exclude backups, generated materials, archives, build outputs, data, dependencies, and the separate email-search project from the source-count scope.
- Set the code identification page count to 60 after rendered verification.
- Change the main software category to `应用软件` and remove YAML from programming languages.
- Mark applicant identity fields for user confirmation before final generation.

## Manual Remediation

- Reassign screenshots to the verified page mapping.
- Remove duplicate or unsupported screenshot references.
- Verify that the referenced overall design, detailed design, and test-case documents exist; otherwise remove those references.
- Keep figures and captions on the same page and headings with their following paragraph.
- Treat screenshots with residual real business data as requiring user replacement or explicit confirmation.
- Remove process-oriented wording from the main-function field.

## Document Generation

Generate Markdown first. After user approval, rebuild DOCX files using the bundled software-copyright toolkit. Use WPS as the available renderer and iterate code typography, margins, and spacing until the physical page counts are exactly 30 and 30 with unique headers numbered 1 through 60.

No existing formal file is overwritten without a timestamped backup. Generated PDFs and PNGs are verification artifacts unless explicitly included in a clean delivery package.

Old ZIP files, `可提交资料*`, and previous formal files are archive-only and cannot enter the new delivery package. The new generation report must record WPS page counts and verification results; it cannot report no warnings when preview or rendering failed.

The legacy product titles in `app.py` and `main.py` require user confirmation. They fall in the omitted middle under this ordering. If the user wants them changed, update the real project source before the final freeze; never alter only the extracted copy.

## Verification

Run structural and rendered checks for:

- exact 30 plus 30 code pages;
- page-number continuity and no duplicate pages;
- no clipping, overlap, blank pages, or overflow;
- source-line traceability and file-hash matches;
- software name, version, date, source count, and page-count consistency;
- screenshot-to-section consistency and privacy;
- absence of comments, tracked changes, macros, hidden attachments, secrets, and personal metadata.

## Agent Team Review

After regeneration, run one review round with three agents:

1. Code originality reviewer: functional-code proportion, algorithmic specificity, and HTML/template exposure.
2. Source authenticity reviewer: file hashes, ordering, full-file extraction, and traceability.
3. Submission consistency reviewer: fields, headers, page counts, manual screenshots, privacy, and packaging.

The release decision is `PASS` only when all blocking findings are resolved. Otherwise the review reports `FAIL` with exact evidence and remediation steps.

## User Gates

After freezing hashes and recalculating material line numbers, write the 16 files and exact order into `代码文件选择.json`, then persist `confirm_stage.py --stage code-selection`. Any selected-file hash change invalidates that confirmation.

Business wording, application fields, screenshot method, and final Markdown still require explicit confirmation before formal Word/TXT generation.
