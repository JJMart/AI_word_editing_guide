# AI Agent Instructions — AI Word Editing Workflow (AGENTS.md)

This workspace uses a structured workflow for editing Microsoft Word documents. Edits are delivered as Word VBA macros that implement changes as tracked revisions. The writer accepts or rejects each change individually.

This file is the single source of truth for AI behavior in this workflow. It lives in the project root as `AGENTS.md` so it is agent-agnostic — agent-specific loader files (`.github/copilot-instructions.md`, `.roo/rules.md`, `.cursorrules`, `.windsurfrules`, `CLAUDE.md`) simply point here. See the [Agent Compatibility](#agent-compatibility) section at the bottom of this file for the full wiring table. The writer-facing summary lives in `README.md`; if a rule here affects what the writer does (session start/end steps, aggressiveness levels, file descriptions), mirror it there.

---

## Before Suggesting Any Edits

Before reviewing pasted text or suggesting any edits, ask the writer any clarifying questions needed. Do not skip this step. Ask about any of the following that are not already provided:

- **Section name and number** — required for logging acronyms and check items correctly
- **Document type** (peer-reviewed journal article, technical report, grant proposal, etc.) — affects tone, passive voice tolerance, and citation style
- **Edit aggressiveness level** — if not stated, default to *Standard*. The lists below are illustrative, not exhaustive — the authoritative item list is "Types of Edits to Consider" further down, tagged by level:
  - *Conservative*: clear errors only (e.g., grammar, contradictions, undefined acronyms — all `[C]` items)
  - *Standard*: errors plus style improvements (e.g., parallel structure, hedges, informal punctuation — all `[C]` and `[S]` items)
  - *Comprehensive*: everything including flow, word choice, and sentence restructuring (all `[C]`, `[S]`, and `[X]` items)
- **Intended tone** (formal/objective, persuasive, accessible to non-specialists)
- **Section-specific conventions** (e.g., passive voice is expected in Methods)
- **Intentional terms or abbreviations** that should not be changed
- **First pass or partial edit** — whether the section has already been partially edited
- **Co-author or journal style preferences** the writer is aware of
- If the section contains in-text citations, ask the writer to provide the reference list so citation completeness can be checked on the first pass

Only skip questions whose answers are clearly already provided by the writer, in the workspace guide files, or in the `.md` export.

If a section has no errors and no stylistic issues above the requested aggressiveness threshold, report that explicitly and do not produce a macro. Do not invent edits to fill a perceived quota.

---

## Writing VBA Macros

Every edit is implemented as a Word VBA macro. Most edits are `Find`/`Replace` with tracked changes; structural edits use the documented insertion, paragraph-split, reordering, and anchor-range-delete patterns.

### Use the canonical template

**Always start from `VBA_MACRO_TEMPLATE.bas` in the workspace root.** Copy the `ReviewEdits_SectionName` sub, rename it to reflect the section being edited (e.g. `ReviewEdits_2_1_Methods`), and fill **only** the EDIT BLOCKS region. Do not modify the HEADER or FOOTER regions — they are required for consistent per-edit reporting. The template also contains a `TestSetup` sub (a one-time verification macro for first-run writers) and a "Reference Patterns" section with skeletons for all five edit types; do not include those in the writer-delivered macro.

### Supported edit types

1. **Find/Replace** — text substitution. The default pattern; use for rewording, removing hedges, fixing typos, terminology swaps.
2. **Insertion** — add text at an anchor. Use for missing topic sentences, transitional phrases, missing citation anchors. Pattern: `Find + Collapse wdCollapseStart/End + InsertBefore/InsertAfter`.
3. **Paragraph split** — break one paragraph into two. Pattern: `Find + Collapse wdCollapseStart + InsertBefore vbCr`.
4. **Reordering** — move a sentence or clause. Pattern: find source, copy text, delete source, find destination, insert. Produces a tracked delete + tracked insert pair; tell the writer to accept both halves together or reject both — accepting only one leaves the document in a broken state.
5. **Anchor-range delete** — long deletion beyond the ~255 char `Find.Text` limit. Pattern: find start anchor + find end anchor + build range + `.Delete`.

Full code skeletons for all five patterns are in the "Reference Patterns" section at the bottom of `VBA_MACRO_TEMPLATE.bas`. Copy the relevant pattern into the EDIT BLOCKS region and adapt.

### Structural edits are riskier than Find/Replace

Insertion, paragraph split, reordering, and anchor-range-delete change document structure, not just text. Two rules:

1. **Run structural edits in their own macro, not batched with Find/Replace edits.** If the writer needs to reject the macro and re-try, they should not lose unrelated text-substitution edits in the process.
2. **Flag structural edits explicitly when proposing them.** Label each proposed edit as `[text]`, `[insert]`, `[split]`, `[move]`, or `[delete-range]` in the rationale so the writer knows what kind of change is coming. The writer should be able to decline a structural edit and ask for a Find/Replace rewording instead.

### Per-edit success/failure reporting (required)

Every Find/Replace block must be wrapped in an `If .Execute(...) Then / Else` structure that writes one line to `sMsg`. **Report the replacement count, not just success/failure** — `.Execute(Replace:=wdReplaceAll)` returns `True` whether it replaced one occurrence or many, so a bare `[OK]` hides accidental over-replacement when an anchor is not unique. Count the replacements and state the expected count so the writer can confirm:

```vb
' Count occurrences first so over-replacement is visible.
Dim nHits As Long
nHits = 0
With oDoc.Content.Duplicate.Find
    .ClearFormatting
    .Text = "it is important to note that "
    .Wrap = wdFindStop
    Do While .Execute
        nHits = nHits + 1
        .Parent.Collapse wdCollapseEnd
    Loop
End With
With oDoc.Content.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = "it is important to note that "
    .Replacement.Text = ""
    .Wrap = wdFindContinue
    If .Execute(Replace:=wdReplaceAll) Then
        nOK = nOK + 1
        sMsg = sMsg & "[OK]   Edit 3: removed hedge (replaced " & nHits & ", expected 1)" & vbCrLf
    Else
        nFail = nFail + 1
        sMsg = sMsg & "[FAIL] Edit 3: anchor not found - removed hedge" & vbCrLf
    End If
End With
```

For a single-instance edit, the writer should see `expected 1` and a matching count; if the reported count exceeds the expected number, the anchor was not unique and the writer must inspect every changed location. A lighter alternative when uniqueness is certain is to omit the pre-count and simply state the expected count in the `[OK]` string. This replaces the old "bare total edit count" pattern: the writer should see one line per edit in the final MsgBox, investigate every `[FAIL]` line, and verify any count that exceeds what was expected before accepting tracked changes.

For anchor-range-delete operations (long deletions), still wrap the deletion in a conditional that logs success or failure the same way.

### VBA coding rules

- Begin every macro with `oDoc.TrackRevisions = True` (template does this).
- Call `.ClearFormatting` and `.Replacement.ClearFormatting` before every `Find` block.
- Set `.Wrap = wdFindContinue` on Find/Replace blocks that should scan the whole document. Use `.Wrap = wdFindStop` when locating a single anchor for a structural edit (insertion, paragraph split, reordering, anchor-range delete) so the search cannot wrap around the end of the document and match an unintended occurrence. The reference patterns in `VBA_MACRO_TEMPLATE.bas` follow this split — Find/Replace uses `wdFindContinue`; the structural patterns use `wdFindStop`.
- Use `.MatchCase = True` for strings containing proper nouns or meaningful capitalization; `False` otherwise.
- **Also use `.MatchCase = True` when the found text begins with an uppercase letter but the replacement should be lowercase** — with `False`, Word auto-capitalizes the replacement to match the found text's case pattern.
- Use `Replace:=wdReplaceAll` by default; use `wdReplaceOne` only when a phrase appears multiple times and only one instance should change — anchor with sufficient surrounding unique text.
- Include enough surrounding context in search strings to guarantee uniqueness in the document.
- Use `ChrW()` (not `Chr()`) for any Unicode character above code point 255 — e.g., `ChrW(8211)` = en dash, `ChrW(8212)` = em dash, `ChrW(8217)` = smart right apostrophe.
- For contractions and possessives in `.Text` strings, use `ChrW(8217)` for the apostrophe — Word AutoCorrect replaces straight `'` with a curly right single quotation mark that a bare `'` will not match.
- `Find.Text` has a hard limit of ~255 characters; for deletions longer than ~200 characters, use the anchor-range-delete pattern (documented in `AI_ERRORS_TO_AVOID.md`) instead of Find/Replace.
- Build the `MsgBox` summary string in a `Dim sMsg As String` variable and call `MsgBox sMsg` — VBA allows a maximum of 24 line continuations per logical line; exceeding this causes a compile error. The template already uses this pattern.
- Never include `Attribute VB_Name = "..."` in generated macro code — this causes a compile syntax error when pasted directly into the VBA editor.
- Comment every `Find`/`Replace` block with the rationale for the edit (the "Edit N:" comment).

---

## Session Management

- Work one document section per session to avoid context window degradation. Target roughly 500–3,000 words per session — adjust heading level accordingly.
- **At the start of every session, read the following files from the workspace before doing anything else:**
  - `GRAMMATICAL_RULES_FORWARD.md` — apply all style and terminology rules found here to every edit
  - `ITEMS_TO_CHECK.md` — scan for open (unresolved) items; list any you find and ask the writer to confirm, resolve, or defer each one before starting new work. Do not silently skip open items — surface them explicitly. Deferred items stay open and must be surfaced again at the next session start.
  - `WRITING_LESSONS_LEARNED.md` — apply general writing lessons from prior sessions
  - `AI_ERRORS_TO_AVOID.md` — avoid all VBA errors and pitfalls documented here
  - `VBA_MACRO_TEMPLATE.bas` — use as the skeleton for any macro produced this session
  - `AI_WRITING_INDICATORS.md` — do **not** read in full at session start, but be aware it exists: you MUST run the Section 8 (Quick Reference Checklist) check after proposing edits and after writing the macro (see "AI Writing Indicators Check" below). This step is mandatory every session, not optional.
  - If any of the four document/cross-project `.md` guide files (`GRAMMATICAL_RULES_FORWARD.md`, `ITEMS_TO_CHECK.md`, `AI_ERRORS_TO_AVOID.md`, `WRITING_LESSONS_LEARNED.md`) are missing, create them from the starter content in the "Guide File Starter Content" section below — but for `AI_ERRORS_TO_AVOID.md` and `WRITING_LESSONS_LEARNED.md`, first ask the writer whether they have existing versions from a previous project to bring in. If `VBA_MACRO_TEMPLATE.bas` or `AI_WRITING_INDICATORS.md` is missing, tell the writer — do not regenerate either from memory (they are researched cross-project assets, not files to reconstruct).
- **The writer should provide a pandoc-generated Markdown (`.md`) export of the document** so the AI can read section content, heading structure, and tables directly. The recommended command is: `pandoc "MyDocument.docx" --wrap=none --track-changes=accept -o "MyDocument.md"`.
- **When first reading a `.md` file in a session, verify it was converted correctly** by checking that headings appear as `#` markers, tables render as Markdown tables, and special characters are intact. Note: superscripts/subscripts applied as character formatting (not true Unicode) will appear as plain characters — this is expected.
- **The `.md` export is lossy.** It drops comments, cross-reference fields, equation objects, and character-formatted super/subscripts. When a proposed edit touches formatted content (subscripts, footnote markers, fields, partial-run bold), flag this to the writer and ask them to verify the `Find.Text` will match the real Word text before trusting the tracked change. **VBA character handling still applies:** Unicode characters visible in the `.md` (en dashes, smart quotes) must still use `ChrW()` in VBA `Find.Text` strings — never copy them literally into VBA string literals.
- **If the writer provides a reference document (e.g., a call for proposals, a submission template, a style guide), it will arrive as a plain-text `.txt` file** extracted from the original PDF using `pypdf` — pandoc cannot read PDF files. Read the `.txt` file as context when the writer asks you to cross-reference or align the document against it. Do not attempt to open or reference the original PDF. **When first reading the `.txt` file, perform an automatic quality check** before using it as reference context: scan for (a) large runs of garbled or non-English characters, (b) pages that extracted as blank or near-blank, (c) section headings that are missing or appear as garbage, and (d) tables that have collapsed into undifferentiated text. The bar is whether the AI can reliably extract information from the file — minor formatting irregularities that would be hard for a human to read are acceptable as long as the content is intact. Report the result to the writer in one or two sentences (e.g., "Extraction looks clean — all sections present and readable" or "Pages 4–6 appear blank and table content on page 9 is garbled — the PDF may be partially scanned"). If significant quality problems are found, ask the writer to re-extract using an OCR tool before proceeding.
- **When editing figure or table captions, edit only the caption text** (the descriptive text after the label and number). Never modify the `Figure N:` / `Table N:` prefix, the number itself, or its punctuation — Word may be maintaining these via automatic numbering, and changing them can break cross-references throughout the document. If uncertain whether captions are automatic, tell the writer to check via Insert → Cross-reference.
- **`oDoc.Content.Find` operates on the main document body only.** It does not reach footnote, endnote, header, footer, or comment text. If an edit targets content in one of those story ranges, iterate `oDoc.StoryRanges` or explicitly address `oDoc.Footnotes(i).Range.Find` / `oDoc.Endnotes(i).Range.Find`. Flag this to the writer before proceeding so they can confirm the edit should run against the non-body story.
- **When scoping a macro to a specific section**, read the exact heading text directly from the `.md` file (`#` markers); use these as bounding anchors in the macro. If no `.md` file is provided, ask the writer to check the Navigation panel (View → Navigation Pane) and confirm the exact heading text.
- Do not modify the guide files during the session. Collect pending updates and apply them only at session end after the writer confirms. Unlike the companion code-documentation workflow (which writes `[UNCONFIRMED]` changelog entries during a session so nothing is lost across session boundaries), this workflow has no pending-marker mechanism: if a session ends before the writer confirms, the proposed guide-file updates are not yet written. This is acceptable because guide-file entries are cheap to regenerate, but it means you must restate any unwritten proposed updates at the start of the next session if the prior session ended before confirmation.
- At the end of each session, update guide files as needed:
  - `GRAMMATICAL_RULES_FORWARD.md` — new style or terminology decisions
  - `ITEMS_TO_CHECK.md` — newly flagged acronyms, cross-references, or consistency issues
  - `AI_ERRORS_TO_AVOID.md` — any VBA errors encountered
  - `WRITING_LESSONS_LEARNED.md` — any broadly applicable writing insights
- A **second-pass session** should be run after all sections are edited, using `ITEMS_TO_CHECK.md` as the agenda to resolve outstanding items.

---

## AI Writing Indicators Check

After proposing edits for a section **and** after writing the VBA macro, check against the Quick Reference Checklist in Section 8 of `AI_WRITING_INDICATORS.md`. Check **the replacement text destined for the document** and flag any AI writing indicators that already appear in the original text (so the writer is aware). Do not check rationale or explanation prose — only text that will end up in the document. Do not load the full `AI_WRITING_INDICATORS.md` at session start; read it only when performing this check.

Verify that none of the proposed edits or macro replacement strings introduce new AI writing indicators.

---

## Types of Edits to Consider

Each entry is tagged with the lowest aggressiveness level at which it should be flagged: **[C]** = Conservative and above, **[S]** = Standard and above, **[X]** = Comprehensive only. At a given level, apply every tag at or below it (Standard applies [C] and [S]; Comprehensive applies all three).

- **[C]** Grammar errors, spelling, and typographical errors
- **[C]** Contradictory phrasing (e.g., mixing passive/active framing in the same clause)
- **[C]** Undefined acronyms on first use
- **[C]** Citation punctuation errors (e.g., missing period after "al" in "et al") — do not alter citation style, numbering format, or author–year conventions
- **[S]** Parallel structure violations
- **[S]** Unnecessary hedges ("are known to", "it should be noted that")
- **[S]** Colloquial or imprecise terms that have standard technical equivalents
- **[S]** Informal punctuation (slashes in place of "and"/"or")
- **[S]** Redundant phrasing
- **[S]** Unexplained jargon (non-acronym)
- **[X]** Awkward parenthetical constructions that disrupt sentence flow
- **[X]** Passive voice where active is clearer (passive is acceptable in Methods sections regardless of level)
- **[X]** Sentence-level restructuring for flow and cadence
- **[X]** Word-choice improvements beyond standard technical equivalents
- **(all levels)** **Do not introduce em dashes (—) in any suggested edit.** If an em dash already exists in the document, flag it and ask the writer whether it is intentional — do not silently preserve or add them. Most writers do not type em dashes naturally and AI models use them far more than typical writers do. Scope note: in correspondence (letters, emails, memos) even a single em dash is a strong AI tell and should always be flagged; in scholarly prose (journal articles, technical reports, proposals) the em dash is a more accepted device, so flag clusters and any that read as AI-inserted rather than reflexively flagging every one. `AI_WRITING_INDICATORS.md` §7 is the authority on this scope distinction.

Edits should feel surgical; preserve the writer's existing voice. Do not paraphrase for the sake of paraphrasing.

---

## Guide File Starter Content

If any of the four `.md` guide files are missing from the workspace, create them with the structure below. For `AI_ERRORS_TO_AVOID.md` and `WRITING_LESSONS_LEARNED.md`, ask the writer first whether they have cross-project versions to bring in.

### `GRAMMATICAL_RULES_FORWARD.md` (document-scoped)
Section headers only; all section bodies empty on creation. Use bullet lists (not tables) for every section — one bullet per rule or acronym, with the originating section in brackets where relevant:
- Units and Measurements
- Proper Nouns and Capitalization
- Terminology Preferences — format: `- "avoid term" → "prefer term" (reason)`
- Sentence Structure Preferences
- Citation Style
- Tense
- Voice and Person
- Document-Specific Acronyms Defined — format: `- [§X.Y] ACR = full term`

### `ITEMS_TO_CHECK.md` (document-scoped)
Section headers only; all section bodies empty on creation. Use bullet lists (not tables). Each item should include the originating section and enough context to resolve on the second pass. Suggested format: `- [§X.Y] <brief description> — <what to verify>`.
- Acronyms Used — Needs First-Definition Check
- Acronyms Assigned — Needs Usage and Re-definition Check
- Cross-Reference Checks
- Consistency Checks
- Numeric Consistency
- Reference List
- Resolved Items (move items here with a resolution note once confirmed)

### `AI_ERRORS_TO_AVOID.md` (cross-project, writer-owned)
Section headers:
- VBA Errors — empty on creation
- Known Pitfalls (Preemptive) — pre-populate with: formatting bleed-through, non-unique search strings, special characters in VBA strings, and macro not found on Alt+F8

Do not reset this file when starting a new document project — carry it forward.

### `WRITING_LESSONS_LEARNED.md` (cross-project, writer-owned)
Section headers, all empty on creation:
- Workflow and Process
- VBA Macro Editing
- Working with AI on Technical Writing
- General Writing

Do not reset this file when starting a new document project — carry it forward.

---

## Known Rendering Note

Two `~` characters separated by text in a chat message may render as strikethrough due to Markdown formatting. This is a display artifact only and does not affect VBA string literals.

---

## README Sync

`README.md` is the writer-facing summary of this workflow and is displayed on the GitHub project page. If a rule change here affects writer steps (session start/end checklists, aggressiveness levels, file descriptions, pandoc command), update `README.md` to match. Update `README.md` only after the writer confirms the workflow rule change is working as intended — it describes current, verified workflow behavior.

---

## Agent Compatibility

The project instructions live in one file: **`AGENTS.md`** in the project root. Each AI agent uses a different mechanism to load project instructions automatically. The table below shows how each agent is wired to `AGENTS.md` so only one instructions file ever needs to be maintained.

| Agent | Auto-load file | How it is wired |
|-------|----------------|-----------------|
| **Codex** | `AGENTS.md` | No loader needed — Codex reads `AGENTS.md` from the repository root natively. |
| **Cline** | `AGENTS.md` | No loader needed — Cline reads `AGENTS.md` natively (also supports `.clinerules/`). |
| **Roo Code** | `.roo/rules.md` | Contains one line: `Read AGENTS.md in the workspace root and follow all instructions found there.` |
| **GitHub Copilot** | `.github/copilot-instructions.md` | Pointer works in Agent/Edits mode; paste or mirror full contents for Chat mode (see comment block in that file). |
| **Cursor** | `.cursorrules` | Contains one line pointing to `AGENTS.md`. |
| **Windsurf** | `.windsurfrules` | Contains one line pointing to `AGENTS.md`. |
| **Claude Code** | `CLAUDE.md` | Contains one line pointing to `AGENTS.md`. |

> **Compatibility note:** This strategy is most reliable with agents that can read project files and use tools. For inline-completion tools or chat assistants without repository file access, treat `AGENTS.md` as source material to paste into that tool's supported instruction surface.

**To add a new AI agent:** create its config file in the location the agent expects, containing only:
`Read AGENTS.md in the workspace root and follow all instructions found there.`
Then add a row to the table above.
