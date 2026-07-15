# Software Copyright Remediation Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Rebuild the current-code software copyright materials so both submitted code ends demonstrate functional originality and render as exactly 30 plus 30 physical pages.

**Architecture:** Freeze the dirty working tree as the user-authorized source baseline, generate deterministic scope and selection manifests, then use the bundled software-copyright scripts to regenerate Markdown and DOCX. WPS-rendered PDFs are the physical-page oracle. Existing formal files are backed up before replacement.

**Tech Stack:** PowerShell, Python 3.12, bundled software-copyright scripts, python-docx/OOXML, WPS COM, Poppler/pdfplumber, SHA-256.

**Isolation note:** A clean Git worktree would discard the user's uncommitted current code. Execution therefore stays in the current workspace and first creates immutable hashes and a formal-material backup.

---

## Chunk 1: Freeze, Scope, And Selection

### Task 0: Confirm Source-Identity Decisions

**Files:**
- Inspect: `app.py`
- Inspect: `main.py`

- [ ] Show the user the legacy product titles in both files and confirm whether they remain part of the current actual source.
- [ ] If the user wants them changed, edit the real project source first; never edit only the extracted copy.
- [ ] Confirm the updated business positioning and persist `confirm_stage.py --stage business` before rebuilding application or manual Markdown.
- [ ] Start source freezing only after these decisions are final.

### Task 1: Freeze The Current Source Baseline

**Files:**
- Create: `软件著作权申请资料/重制工作区/源码范围清单.json`
- Create: `软件著作权申请资料/重制工作区/源码范围清单.md`
- Create: `软件著作权申请资料/正式资料-修改前备份-<timestamp>/`

- [ ] Create the timestamped formal-material backup with PowerShell `Copy-Item -Recurse`.
- [ ] Inventory the explicit current-project source scope: six runtime Python files, seven templates, four test files, two tool files, BAT/ISS/PS1 launch and build scripts.
- [ ] Record path, inclusion reason, SHA-256, physical lines, and nonblank lines for every included file.
- [ ] Exclude `软件著作权申请资料*/`, `整理归档/`, `build/`, `dist/`, `邮箱检索/`, data, dependencies, caches, and generated artifacts.
- [ ] Verify the manifest totals and write the confirmed physical-line total into the remediation summary.

Run: `Get-FileHash -Algorithm SHA256 <each-file>` and line-count checks.
Expected: no duplicate paths; every included file exists; old `14440` is not reused.

### Task 2: Persist The Approved Code Ordering

**Files:**
- Modify: `软件著作权申请资料/草稿/代码文件选择.json`
- Test: `软件著作权申请资料/草稿/代码文件选择.json`

- [ ] Replace selection order with the 16 approved files from the design spec.
- [ ] Set all unrelated candidates to `selected: false` with explicit exclusion reasons.
- [ ] Recalculate 3,987 nonblank lines plus 16 trace markers, yielding a 4,003-line material stream.
- [ ] Verify front lines 1-1,500, omitted lines 1,501-2,550, and back lines 2,551-4,003.
- [ ] Verify all HTML and `main.py` are in the omitted middle and the back starts inside `app.py`.
- [ ] Record the code-selection gate with `confirm_stage.py --stage code-selection` using the user's approval of approach A.

Expected: any later selected-file hash change invalidates the selection confirmation.

## Chunk 2: Regenerate Draft Materials

### Task 3: Confirm Version Facts Before Markdown

**Files:**
- Modify: `软件著作权申请资料/草稿/申请表字段确认输入-补正.json`
- Modify: `软件著作权申请资料/草稿/申请表信息.md`

- [ ] Present evidence-based values for software name, current project version, actual completion date, source total, category `应用软件`, languages without YAML, and page count 60.
- [ ] Stop for explicit user confirmation of version, completion date, applicant identity, and environment fields.
- [ ] Persist the confirmation with `confirm_stage.py --stage application-fields` only after the user responds.

Expected: no formal DOCX generation before this gate passes.

### Task 4: Generate The Code Markdown And Trace Manifest

**Files:**
- Modify: `软件著作权申请资料/草稿/代码-前30页.md`
- Modify: `软件著作权申请资料/草稿/代码-后30页.md`
- Modify: `软件著作权申请资料/草稿/代码提取清单.json`
- Modify: `软件著作权申请资料/草稿/代码提取清单.md`

- [ ] Run `extract_code_material.py` with the frozen project, confirmed version, and approved selection JSON.
- [ ] Compare the generated file hashes against the frozen scope manifest.
- [ ] Assert front markers contain only Python/test sources.
- [ ] Assert no HTML marker appears in either submitted Markdown and the back starts in `app.py`.
- [ ] Confirm total logical pages are 81 and submitted headers are numbered 1-30 and 31-60.

Expected: front 1,500 material lines; back 1,453 material lines; no fabricated or reordered lines inside a file.

### Task 5: Remediate The Application And Manual Drafts

**Files:**
- Modify: `软件著作权申请资料/草稿/申请表信息.md`
- Modify: `软件著作权申请资料/草稿/操作手册.md`
- Modify: `软件著作权申请资料/草稿/操作手册自检记录.md`
- Modify: `软件著作权申请资料/草稿/操作手册自检记录.json`

- [ ] Replace source count with the confirmed scope-manifest total and page count with 60.
- [ ] Change category to `应用软件`, remove YAML from languages, and move process wording out of the main-function field.
- [ ] Stop and ask the user to confirm reuse of the existing user-supplied screenshots or provide replacements.
- [ ] Persist `confirm_stage.py --stage screenshot-method --method user-supplied` only after that confirmation.
- [ ] Reassign screenshots: 1 home, 2 config, 3 extract/flow, 4 lookup, 5 supplement, 7/9 preview, 8 send/log, 6 download.
- [ ] Remove unsupported duplicate screenshot use and keep captions with images.
- [ ] Verify referenced design/test documents exist or remove those references.
- [ ] Run and update the three manual self-check rounds.

Expected: all functional claims trace to current routes/templates; residual screenshot privacy risks are explicit.

## Chunk 3: Build, Render, And Review

### Task 6: Obtain Final Markdown Approval

**Files:**
- Verify: `软件著作权申请资料/草稿/*.md`
- Modify: `软件著作权申请资料/草稿/最终生成确认.json`

- [ ] Show the user the changed-field summary and exact code composition.
- [ ] Stop for explicit approval of all Markdown drafts.
- [ ] Persist `confirm_stage.py --stage markdown` after approval.

### Task 7: Build And Physically Paginate Formal Documents

**Files:**
- Modify: `软件著作权申请资料/正式资料/申请表信息.txt`
- Modify: `软件著作权申请资料/正式资料/<软件名>-代码(前30页).docx`
- Modify: `软件著作权申请资料/正式资料/<软件名>-代码(后30页).docx`
- Modify: `软件著作权申请资料/正式资料/<软件名>_操作手册.docx`
- Modify: `软件著作权申请资料/正式资料/生成报告.md`

- [ ] Run the bundled `build_docx_from_md.py` only after all gates pass.
- [ ] Export the three DOCX files to PDF with read-only WPS COM.
- [ ] Audit PDF physical page counts, header numbers, blank pages, and character bounds.
- [ ] If either code file is not exactly 30 physical pages, adjust code font, line spacing, or margins in the required direction with a task-local postprocessor, without changing source text, then rerender.
- [ ] Repeat until front is exactly 30 pages numbered 1-30 and back is exactly 30 pages numbered 31-60.
- [ ] Fix manual caption/heading pagination and rerender every page.
- [ ] Update the generation report with actual WPS results and warnings.

Expected: no clipping, overlap, missing page number, duplicate page number, or unreported preview failure.

### Task 8: Run One Three-Agent Final Review

**Files:**
- Create: `软件著作权申请资料/审核工作区/修改后一轮Agent-Team审核报告.md`

- [ ] Agent 1 reviews originality: functional-code ratio, algorithmic specificity, HTML exposure, and prior-rejection risk.
- [ ] Agent 2 reviews authenticity: hashes, full-file extraction, material boundaries, and source traceability.
- [ ] Agent 3 reviews submission consistency: fields, headers, rendered pages, manual screenshots, privacy, and package boundary.
- [ ] Reconcile disagreements against rendered and hash evidence.
- [ ] If any blocker is found, the round result is `FAIL`.
- [ ] After fixing a blocker, start a new complete three-agent review round; never reuse the previous round's `PASS` decision.

Expected: one complete round reports `PASS` only if all three agents report no blocker on the same artifact set.

### Task 9: Final Verification And Handoff

- [ ] Run Python syntax checks for task-local verification scripts.
- [ ] Verify formal-file sizes, timestamps, hashes, and exact page counts.
- [ ] Verify old ZIP and old `可提交资料*` files are excluded from the new handoff.
- [ ] Append the final result and affected files to the Obsidian conversation record.
- [ ] Report any remaining user-supplied legal facts or screenshot replacements without claiming submission readiness.

No generated soft-copyright material or private applicant data is committed to Git.
