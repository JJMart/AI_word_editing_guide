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
- If `GRAMMATICAL_RULES_FORWARD.md` is not attached, ask whether rules from prior sessions should be applied

Only skip questions whose answers are clearly already provided by the writer or in attached guide files.

---

## Writing VBA Macros

Every edit is implemented as a Word VBA macro using `Find`/`Replace` with tracked changes. Follow these rules:

- Begin every macro with `oDoc.TrackRevisions = True` (idempotent — safe if already on)
- Call `.ClearFormatting` and `.Replacement.ClearFormatting` before every `Find` block
- Set `.Wrap = wdFindContinue` on every `Find` block
- Use `.MatchCase = True` for strings containing proper nouns or meaningful capitalization; `False` otherwise
- Use `Replace:=wdReplaceAll` by default; use `wdReplaceOne` only when a phrase appears multiple times and only one instance should change — anchor with sufficient surrounding unique text
- Include enough surrounding context in search strings to guarantee uniqueness in the document
- Comment every `Find`/`Replace` block with the rationale for the edit
- End the macro with a `MsgBox` stating how many edits were applied

---

## Session Management

- Work one document section per session to avoid context window degradation
- `GRAMMATICAL_RULES_FORWARD.md` and `ITEMS_TO_CHECK.md` must be attached by the writer each session via `#file` — apply rules from these files to all edits
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

---

## Known Rendering Note

Two `~` characters separated by text in the chat may render as strikethrough due to Markdown formatting. This is a display artifact only and does not affect VBA string literals.

---

## File Sync Note

`AI_WORD_EDITING_GUIDE.md` is the full workflow reference for this project and mirrors the rules in this file in expanded form. If anything in these instructions is worth updating based on a session's experience, make the corresponding update in `AI_WORD_EDITING_GUIDE.md` as well so the two files stay in sync. Do not direct the AI to read `AI_WORD_EDITING_GUIDE.md` during normal editing sessions — everything required is already in these instructions.
