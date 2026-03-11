# AI Word Editing Guide

---

## Section 1: For the Writer — How to Run a VBA Macro in Word

1. **Open the VBA Editor:** Press `Alt+F11`
2. **Insert a new module:** In the menu bar, click `Insert` → `Module`
3. **Paste the macro:** Copy the VBA code from the chat and paste it into the blank module window
4. **Return to Word and save:** Press `Alt+F4` to close the VBA editor (or click the X), then save the document (`Ctrl+S`) — this ensures Word registers the macro
5. **Run the macro:** Press `Alt+F8`, select the macro name from the list, and click `Run`
6. **Review tracked changes:** All edits will appear as tracked revisions. Accept or reject each one individually using the Review ribbon

> **Tip:** Make sure macros are enabled. Go to `File → Options → Trust Center → Trust Center Settings → Macro Settings` and select "Enable all macros" or "Enable with notification."

> **Tip:** Save a dated backup copy of your `.docx` before each editing session so you have a rollback point.

> **Tip:** Track Changes does not need to be turned on manually — the macro enables it automatically. If it is already on, this is harmless.

---

## Section 1b: Session Start and End Checklists

### Session Start
1. Save a dated backup copy of the `.docx` (e.g., `MyDoc_2026-03-10.docx`)
2. Open a new chat session
3. Attach `AI_WORD_EDITING_GUIDE.md`, `GRAMMATICAL_RULES_FORWARD.md`, and `ITEMS_TO_CHECK.md` using `#file` in the chat input
4. Tell the AI the following before pasting text:
   - The **document type** (e.g., peer-reviewed paper, technical report, grant proposal)
   - The **section name and number** you are working on
   - Your preferred **edit aggressiveness level** (see Section 2 — General Workflow)
5. Paste the section text
6. **Wait for the AI to ask any clarifying questions before it begins suggesting edits** — answer these before proceeding

### Session End
1. Ask the AI to update `GRAMMATICAL_RULES_FORWARD.md` with any new style or terminology rules established this session
2. Ask the AI to update `ITEMS_TO_CHECK.md` with any newly flagged acronyms, cross-references, or consistency issues
3. Ask the AI to update `AI_ERRORS_TO_AVOID.md` if any VBA errors were encountered
4. Ask the AI to update `WRITING_LESSONS_LEARNED.md` if any broadly applicable writing insights emerged
5. Run the macro and review all tracked changes in Word before closing the session

---

## Section 2: For the AI — Workflow Guidelines and Best Practices

### General Workflow
- The writer pastes a section of text into the chat
- The AI reviews it and proposes specific edits with rationale
- The AI then writes a VBA macro implementing those edits using `Find`/`Replace` with tracked changes
- The writer runs the macro and accepts or rejects each tracked change individually
- **At the start of each session, the writer should declare:**
  - **Document type** (e.g., peer-reviewed journal article, technical report, grant proposal) — affects passive voice tolerance, citation style, and tone
  - **Edit aggressiveness level:**
    - *Conservative* — flag clear errors only (grammar, contradictions, undefined acronyms)
    - *Standard* — errors plus style improvements (parallel structure, hedges, informal punctuation)
    - *Comprehensive* — everything including flow, word choice, and sentence restructuring
  - If not declared, default to *Standard*
- **Before suggesting any edits, the AI must review the pasted text and ask any clarifying questions it considers important.** Do not proceed to edits until the writer has answered. Questions to consider asking if not already declared or apparent from the attached guide files:
  - Document type and target audience, if not stated
  - Edit aggressiveness level, if not stated
  - Intended tone (e.g., formal/objective, persuasive, accessible to non-specialists)
  - Whether specific sections have different conventions (e.g., passive voice is expected in Methods)
  - Whether any terms, phrasings, or abbreviations are intentional and should not be changed
  - Whether the section has already been partially edited or should be treated as a first pass
  - Any co-author or journal style preferences the writer is aware of
  - If `GRAMMATICAL_RULES_FORWARD.md` is not attached, whether rules from prior sessions should be applied
  - Only skip questions whose answers are clearly already provided in the attached guide files or the writer's message

### Session Management
- Work **one document section per chat session** to keep context tight and avoid context window degradation
- **If the writer does not provide a section number or name when pasting text, ask for it before proceeding** — section identity is required to correctly log items in `ITEMS_TO_CHECK.md` and `GRAMMATICAL_RULES_FORWARD.md` (e.g., for ordering acronym first-use)
- At the start of each session, the writer should attach `GRAMMATICAL_RULES_FORWARD.md` and `ITEMS_TO_CHECK.md` using `#file` in the chat input so prior session decisions carry forward
- At the **end of each session**, explicitly update any of the markdown guide files with new rules, errors, or check items discovered that session
- A **second-pass session** should be run after all sections are edited, using `ITEMS_TO_CHECK.md` as the agenda to resolve outstanding items
- Save a dated copy of the `.docx` before each session as a rollback point
- **If any of the five markdown guide files are missing from the working folder, create them** using the starter content defined in the "Markdown Guide Files" section below — but first ask the writer whether they have existing versions of `AI_ERRORS_TO_AVOID.md` or `WRITING_LESSONS_LEARNED.md` to bring in, as these are cross-project living documents

### VBA Writing Rules
- Always include `oDoc.TrackRevisions = True` at the top — it is idempotent (safe if already on)
- Use `Find`/`Replace` with `.MatchCase` set to `True` when the search string contains proper nouns, sentence-opening capitals, or any capitalization that is meaningful and must be matched exactly; use `False` otherwise
- Always call `.ClearFormatting` and `.Replacement.ClearFormatting` before each Find block to prevent formatting bleed-through between replacements
- Set `.Wrap = wdFindContinue` so the search covers the full document regardless of cursor position
- Use `Replace:=wdReplaceAll` unless a phrase appears multiple times and only one instance should be changed — in that case use `wdReplaceOne` and anchor with surrounding unique text
- For long sentence replacements, match enough unique surrounding text to guarantee uniqueness in the document
- End each macro with a `MsgBox` confirming how many edits were applied
- Comment every Find/Replace block with the rationale for the edit

### Markdown Guide Files — Purpose and Starter Content

These five files live in the same folder as the document being edited. If any are missing at the start of a session, create them with the starter content shown below.

| File | Scope | Purpose |
|---|---|---|
| `AI_WORD_EDITING_GUIDE.md` | Document project | Master workflow reference for both writer and AI. Instructions for running VBA, session management rules, VBA coding rules, and this file-descriptions table. |
| `AI_ERRORS_TO_AVOID.md` | **Cross-project (writer-owned)** | Living log of errors encountered during AI-assisted editing sessions (VBA failures, bad find strings, etc.) so they are not repeated. Carries forward across projects. |
| `GRAMMATICAL_RULES_FORWARD.md` | Document project | Document-specific style and terminology decisions that must carry forward across all sections: unit formats, species name capitalization, terminology preferences, acronym definition tracking. |
| `ITEMS_TO_CHECK.md` | Document project | Running checklist of items to resolve on a second pass: acronyms used but not yet confirmed to be defined earlier, acronyms defined (to confirm they are used and not re-defined), figure/table cross-reference checks, and consistency checks. |
| `WRITING_LESSONS_LEARNED.md` | **Cross-project (writer-owned)** | Living document of general writing lessons and insights that apply beyond any specific document or project. Should travel with the writer to future projects. |

**Starter content for `GRAMMATICAL_RULES_FORWARD.md`:** Section headers for Units and Measurements, Species Names, Terminology Preferences (table with Avoid/Prefer/Reason columns), Sentence Structure Preferences, Citation Style, and Document-Specific Acronyms Defined (table with Acronym/Full term/First defined in section columns). All tables start empty.

**Starter content for `ITEMS_TO_CHECK.md`:** Section headers for Acronyms Used (needs first-definition check), Acronyms Assigned (needs usage and re-definition check), Cross-Reference Checks, Consistency Checks, Template for Logging an Item, and Resolved Items. All tables start empty.

**Starter content for `AI_ERRORS_TO_AVOID.md`:** Section headers for VBA Errors and Known Pitfalls (Preemptive). Pre-populate Known Pitfalls with: formatting bleed-through, non-unique search strings, special characters in VBA strings, and macro not found on Alt+F8 (see current file for text). This file is writer-owned and cross-project; do not reset it when starting a new document project — carry it forward and add to it.

**Starter content for `WRITING_LESSONS_LEARNED.md`:** Section headers for Workflow and Process, VBA Macro Editing, Working with AI on Technical Writing, and General Writing. This file is writer-owned and cross-project; do not reset it when starting a new document project — carry it forward and add to it.

> **Living document check:** For `AI_ERRORS_TO_AVOID.md` and `WRITING_LESSONS_LEARNED.md`, if neither file exists in the working folder, ask the writer whether they have previous versions to copy in before generating new ones from scratch.

---

### Copilot Instructions Sync Note

`.github/copilot-instructions.md` is a condensed version of the AI rules in this file and is automatically loaded by GitHub Copilot Chat in VS Code. If a new rule is added to this guide that the AI should always follow — especially anything in General Workflow, VBA Writing Rules, or Types of Edits to Consider — add the corresponding entry to `copilot-instructions.md` as well so it takes effect automatically without manual file attachment.

### README Sync Note

`README.md` is the writer-facing summary of this guide and is displayed on the GitHub project page. If a new rule or workflow step is added to this guide that the writer needs to know — especially anything in Session Start/End checklists, the aggressiveness level definitions, or the file descriptions table — add the corresponding entry to `README.md` as well.

---

### Markdown Rendering Note
- Two `~` characters separated by text in the chat may render as ~~strikethrough~~ due to Markdown formatting. This is a display artifact only and does not affect the VBA string literals.

### Types of Edits to Consider
- Parallel structure violations
- Unnecessary hedges ("are known to", "it should be noted that")
- Contradictory phrasing (e.g., mixing passive/active framing in the same clause)
- Colloquial or imprecise terms that have standard technical equivalents
- Informal punctuation (slashes in place of "and"/"or")
- Unexplained jargon or acronyms
- Awkward parenthetical constructions that disrupt sentence flow
- Redundant phrasing
- Passive voice where active is clearer (use judgment — passive is acceptable in methods sections)
- **Do not introduce em dashes (—) in any suggested edit.** If an em dash already exists in the document, flag it and ask the writer whether it is intentional — do not silently preserve or add them. Most writers do not type em dashes naturally and AI models use them far more than typical writers do.
