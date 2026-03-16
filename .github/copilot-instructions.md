# Copilot Instructions — AI Word Editing Workflow

This workspace uses a structured workflow for editing Microsoft Word documents. Edits are delivered as Word VBA macros that implement changes as tracked revisions. The writer accepts or rejects each change individually.

---

## Before Suggesting Any Edits

Before reviewing pasted text or suggesting any edits, ask the writer any clarifying questions needed. Do not skip this step. Ask about any of the following that are not already provided:

- **Section name and number** — required for logging acronyms and check items correctly
- **Document type** (peer-reviewed journal article, technical report, grant proposal, etc.) — affects tone, passive voice tolerance, and citation style
- **Edit aggressiveness level** — if not stated, default to *Standard*:
  - *Conservative*: clear errors only (grammar, contradictions, undefined acronyms)
  - *Standard*: errors plus style improvements (parallel structure, hedges, informal punctuation)
  - *Comprehensive*: everything including flow, word choice, and sentence restructuring
- **Intended tone** (formal/objective, persuasive, accessible to non-specialists)
- **Section-specific conventions** (e.g., passive voice is expected in Methods)
- **Intentional terms or abbreviations** that should not be changed
- **First pass or partial edit** — whether the section has already been partially edited
- **Co-author or journal style preferences** the writer is aware of
Only skip questions whose answers are clearly already provided by the writer, in the workspace guide files, or in the `.txt` export.

---

## Writing VBA Macros

Every edit is implemented as a Word VBA macro using `Find`/`Replace` with tracked changes. Follow these rules:

- Begin every macro with `oDoc.TrackRevisions = True` (idempotent — safe if already on)
- Call `.ClearFormatting` and `.Replacement.ClearFormatting` before every `Find` block
- Set `.Wrap = wdFindContinue` on every `Find` block
- Use `.MatchCase = True` for strings containing proper nouns or meaningful capitalization; `False` otherwise
- **Also use `.MatchCase = True` when the found text begins with an uppercase letter but the replacement should be lowercase** — with `False`, Word auto-capitalizes the replacement to match the found text's case pattern
- Use `Replace:=wdReplaceAll` by default; use `wdReplaceOne` only when a phrase appears multiple times and only one instance should change — anchor with sufficient surrounding unique text
- Include enough surrounding context in search strings to guarantee uniqueness in the document
- Use `ChrW()` (not `Chr()`) for any Unicode character above code point 255 — e.g., `ChrW(8211)` = en dash, `ChrW(8212)` = em dash, `ChrW(8217)` = smart right apostrophe
- For contractions and possessives in `.Text` strings, use `ChrW(8217)` for the apostrophe — Word AutoCorrect replaces straight `'` with a curly right single quotation mark that a bare `'` will not match
- `Find.Text` has a hard limit of ~255 characters; for deletions longer than ~200 characters, use the anchor-range-delete pattern instead of Find/Replace
- Build the `MsgBox` summary string in a `Dim sMsg As String` variable and call `MsgBox sMsg` — VBA allows a maximum of 24 line continuations per logical line; exceeding this causes a compile error
- Never include `Attribute VB_Name = "..."` in generated macro code — this causes a compile syntax error when pasted directly into the VBA editor
- Comment every `Find`/`Replace` block with the rationale for the edit
- End the macro with `MsgBox sMsg`; tell the writer to verify the reported edit count matches expectations before accepting changes

---

## Session Management

- Work one document section per session to avoid context window degradation
- **At the start of every session, read the following files from the workspace before doing anything else:**
  - `GRAMMATICAL_RULES_FORWARD.md` — apply all style and terminology rules found here to every edit
  - `ITEMS_TO_CHECK.md` — be aware of all open items; do not re-flag already resolved items
  - `WRITING_LESSONS_LEARNED.md` — apply general writing lessons from prior sessions
  - `AI_ERRORS_TO_AVOID.md` — avoid all VBA errors and pitfalls documented here
  - If any of these files are missing, create them from the starter content defined in `AI_WORD_EDITING_GUIDE.md` — but for `AI_ERRORS_TO_AVOID.md` and `WRITING_LESSONS_LEARNED.md`, first ask the writer whether they have existing versions from a previous project to bring in
- **The writer should provide a plain-text (`.txt`) export of the document** so the AI can read section content and heading text directly — this eliminates manual section pasting and the need to ask the writer to check the Navigation panel for section-scoping anchors
- **When first reading a `.txt` file in a session, verify it was exported correctly** by checking that special characters are intact. If they appear as `?` or garbled, ask the writer to re-export: `File → Save As → Plain Text → Other encoding → Unicode (UTF-8)`, with **Insert line breaks** and **Allow character substitution** both unchecked. Note: superscripts/subscripts applied as character formatting (not true Unicode) will always appear as plain characters regardless of encoding — this is expected
- **When the first in-text citation is encountered in a section, ask the writer to provide the reference list** so that citation completeness can be checked on the first pass rather than deferred to a second-pass session. If the writer provides it, flag any in-text citations missing from the list or any reference list entries not cited in the text.
- **Before modifying any figure or table caption, ask the writer whether they use automatic Word captions** — editing automatic captions can break Word's automatic numbering links and cross-references throughout the document
- **When scoping a macro to a specific section**, use heading text found in the `.txt` file as bounding anchors; if no `.txt` file is provided, ask the writer to check the Navigation panel (View → Navigation Pane) and confirm the exact heading text
- If guide files are missing from the workspace, ask the writer before creating new ones — `AI_ERRORS_TO_AVOID.md` and `WRITING_LESSONS_LEARNED.md` are cross-project living documents the writer may already have
- At the end of each session, update guide files as needed:
  - `GRAMMATICAL_RULES_FORWARD.md` — new style or terminology decisions
  - `ITEMS_TO_CHECK.md` — newly flagged acronyms, cross-references, or consistency issues
  - `AI_ERRORS_TO_AVOID.md` — any VBA errors encountered
  - `WRITING_LESSONS_LEARNED.md` — any broadly applicable writing insights

---

## Types of Edits to Consider

- Parallel structure violations
- Unnecessary hedges ("are known to", "it should be noted that")
- Contradictory phrasing (e.g., mixing passive/active framing in the same clause)
- Colloquial or imprecise terms that have standard technical equivalents
- Informal punctuation (slashes in place of "and"/"or")
- Unexplained jargon or acronyms
- Awkward parenthetical constructions that disrupt sentence flow
- Redundant phrasing
- Passive voice where active is clearer (acceptable in Methods sections)
- Citation punctuation errors (e.g., missing period after "al" in "et al") are acceptable to correct; do not alter citation style, numbering format, or author–year conventions
- **Do not introduce em dashes (—) in any suggested edit.** If an em dash already exists in the document, flag it and ask the writer whether it is intentional — do not silently preserve or add them. Most writers do not type em dashes naturally and AI models use them far more than typical writers do.

---

## Known Rendering Note

Two `~` characters separated by text in the chat may render as strikethrough due to Markdown formatting. This is a display artifact only and does not affect VBA string literals.

---

## File Sync Note

`AI_WORD_EDITING_GUIDE.md` is the full workflow reference for this project and mirrors the rules in this file in expanded form. If anything in these instructions is worth updating based on a session's experience, make the corresponding update in `AI_WORD_EDITING_GUIDE.md` as well so the two files stay in sync. Do not direct the AI to read `AI_WORD_EDITING_GUIDE.md` during normal editing sessions — everything required is already in these instructions.
