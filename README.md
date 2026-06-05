# AI Word Editing Workflow

A structured workflow for editing Microsoft Word documents using an AI assistant (GitHub Copilot, Roo Code, ChatGPT, Claude, etc.) with changes delivered as tracked revisions via VBA macros. The writer retains full control — every suggested edit appears as a tracked change to accept or reject individually. When the AI is uncertain about a phrasing preference, encounters an ambiguous case, or needs context only the writer can provide, it will ask for clarification or present options rather than guessing.

---

## How It Works

1. You convert your document to Markdown using pandoc and place the `.md` file in the project folder
2. The AI reads the `.md` file directly — no need to paste section text manually
3. The AI asks clarifying questions before beginning; during review it will flag any items it is uncertain about and ask you to clarify or choose between options
4. The AI proposes specific edits with rationale, then writes a Word VBA macro implementing those edits with Track Changes enabled. The macro is built from the canonical [`VBA_MACRO_TEMPLATE.bas`](VBA_MACRO_TEMPLATE.bas) skeleton and reports per-edit success or failure in a MsgBox at the end
5. You paste and run the macro in Word — all changes appear as tracked revisions
6. You accept or reject each change individually using the Word Review ribbon

---

## Files in This Repository

| File | Scope | Description |
|---|---|---|
| [`AGENTS.md`](AGENTS.md) | **Cross-project** | Single source of truth for AI behavior in this workflow. All agent-specific loader files (`.github/copilot-instructions.md`, `.roo/rules.md`, `.cursorrules`, `.windsurfrules`, `CLAUDE.md`) redirect here. Update this file when workflow rules change. |
| [`VBA_MACRO_TEMPLATE.bas`](VBA_MACRO_TEMPLATE.bas) | **Cross-project** | Canonical VBA macro skeleton. The AI copies this template and fills only the EDIT BLOCKS region. Ensures every macro has consistent header, footer, per-edit reporting, and coding conventions. |
| [`GRAMMATICAL_RULES_FORWARD.md`](GRAMMATICAL_RULES_FORWARD.md) | Document project | Style and terminology decisions specific to the document being edited (units, capitalization, terminology preferences, acronym tracking). Read by the AI at session start so rules carry forward. |
| [`ITEMS_TO_CHECK.md`](ITEMS_TO_CHECK.md) | Document project | Running checklist of items to verify on a second pass: undefined acronyms, unused acronyms, cross-reference accuracy, numeric consistency, and reference list completeness. |
| [`AI_ERRORS_TO_AVOID.md`](AI_ERRORS_TO_AVOID.md) | **Cross-project** | Living log of errors encountered during AI-assisted VBA editing sessions. Carry forward to future projects and add to it over time. |
| [`WRITING_LESSONS_LEARNED.md`](WRITING_LESSONS_LEARNED.md) | **Cross-project** | Living document of general writing insights gained through AI-assisted editing. Carry forward across all projects. |
| [`AI_WRITING_INDICATORS.md`](AI_WRITING_INDICATORS.md) | **Cross-project** | Reference on indicators of AI-generated writing. The AI reads Section 8 (Quick Reference Checklist) after proposing edits and after writing a macro to flag AI indicators in the original text and verify none are introduced by the edits. |
| `pypdf` (Python package) | Tool | Extracts plain text from PDF reference documents (pandoc cannot read PDF). Install once with `pip install pypdf`. See [Converting a PDF Reference Document to Text](#converting-a-pdf-reference-document-to-text) below. |

> **Cross-project files** are writer-owned and should not be reset when starting a new document project. Copy them into the new project folder and continue adding to them.

---

## Converting the Document to Markdown

Before each session, convert your `.docx` to Markdown using pandoc so the AI can read it directly:

1. Open a terminal (PowerShell) in the project folder
2. Run:
   ```powershell
   pandoc "MyDocument.docx" --wrap=none --track-changes=accept -o "MyDocument.md"
   ```
   - `--wrap=none` — keeps each paragraph as a single line, making it easy for the AI to locate anchor text
   - `--track-changes=accept` — shows the clean accepted-state text; omit this flag if the document has pending tracked changes you want the AI to see
3. Place the `.md` file in the same project folder as the guide files

> **Why pandoc instead of plain-text export?** Pandoc preserves document structure — section headings appear as `#` Markdown headers (so the AI can scope macros to sections without you checking the Navigation pane), tables retain their layout, and lists stay readable. Plain-text export flattens all of this.

> **Lossy export caveat:** pandoc does not preserve comments, cross-reference fields, equation objects, or character-formatted super/subscripts. The AI will flag any edit that touches formatted content so you can verify the Find string matches the real Word text before accepting the tracked change.

> **VBA character note:** Unicode characters (en dashes, smart quotes, etc.) are visible in the `.md` file, but the AI must still use `ChrW()` codes in VBA `Find.Text` strings — never copy those characters literally into VBA string literals. Superscripts and subscripts applied as Word character *formatting* (not true Unicode) will still appear as plain digits or letters — this is expected.

> **Install pandoc:** If pandoc is not already installed, download it from [pandoc.org](https://pandoc.org/installing.html) or install via `winget install JohnMacFarlane.Pandoc` in PowerShell.

---

## Converting a PDF Reference Document to Text

When working on proposals, reports, or other documents that must align with an external reference (a call for proposals, a submission template, a style guide), you may want to provide that reference to the AI as readable text.

**Pandoc cannot read PDF files** — it converts *to* PDF but not *from* it. Use `pypdf` (a Python package) instead:

1. Install `pypdf` once:
   ```powershell
   pip install pypdf
   ```
2. Extract the text from the PDF:
   ```powershell
   python -c "import pypdf; r=pypdf.PdfReader('reference.pdf'); print('\n'.join(p.extract_text() for p in r.pages))" > reference.txt
   ```
   Replace `reference.pdf` with your file name. The output `reference.txt` will be placed in the current directory.
3. Place `reference.txt` in the project folder alongside the other guide files.
4. Tell the AI the filename and ask it to read the file and verify extraction quality before using it as reference context. The AI will check whether the text is coherent and flag any sections that appear garbled or empty — you do not need to review the raw `.txt` yourself.

> **Quality caveat:** PDF text extraction works well for text-based PDFs (those with selectable text). Scanned PDFs or PDFs with heavy formatting may produce garbled or incomplete output. If the AI flags quality problems, re-extract using an OCR tool such as `pytesseract` or Adobe Acrobat's export feature.

---

## Running a VBA Macro in Word

1. **Open the VBA Editor:** Press `Alt+F11`
2. **Insert a new module:** Click `Insert` → `Module`
3. **Paste the macro:** Copy the VBA code from the AI chat and paste it into the blank module window
4. **Return to Word and save:** Close the VBA editor, then save the document (`Ctrl+S`)
5. **Run the macro:** Press `Alt+F8`, select the macro name, and click `Run`
6. **Review the MsgBox report:** The macro reports each edit as `[OK]` or `[FAIL]` with a short description, followed by success and failure totals. Investigate every `[FAIL]` line before accepting any tracked changes — a `[FAIL]` means the search anchor was not found in the document, usually because of a character-matching issue (smart apostrophe, special Unicode, formatted content) or an incorrect anchor string.
7. **Review tracked changes:** Set Review → Display for Review to **All Markup** so individual tracked changes are visible, then accept or reject each revision using the Review ribbon.

> **Enable macros first:** `File → Options → Trust Center → Trust Center Settings → Macro Settings → Enable all macros` (or "Enable with notification")

> **First-run verification:** If you are running this workflow on a new machine, paste the contents of [`VBA_MACRO_TEMPLATE.bas`](VBA_MACRO_TEMPLATE.bas) into a Word module and run the `TestSetup` macro. It inserts a visible marker as a tracked change, immediately undoes it, and shows a MsgBox confirming that macros are enabled, Track Changes works, and tracked edits can be applied and reverted. Run this only once per machine; you do not need to run it before every editing session.

---

## Session Workflow

### Before Each Session
1. Save a dated backup copy of your `.docx` (e.g., `MyDoc_2026-03-10.docx`)
2. Convert the document to Markdown using the pandoc command above and place the `.md` file in the project folder
3. Open a new AI chat session
4. Load the AI behavior rules:
   - **VS Code agents (Copilot, Roo Code, Cline, Cursor, Windsurf, or similar):** no attachment needed — [`AGENTS.md`](AGENTS.md) and the other guide files are read automatically from the workspace via the agent-specific loader file
   - **Other AI tools (ChatGPT, Claude, etc.):** attach [`AGENTS.md`](AGENTS.md), [`GRAMMATICAL_RULES_FORWARD.md`](GRAMMATICAL_RULES_FORWARD.md), [`ITEMS_TO_CHECK.md`](ITEMS_TO_CHECK.md), and [`VBA_MACRO_TEMPLATE.bas`](VBA_MACRO_TEMPLATE.bas). If your tool supports a persistent system prompt file (e.g., a Claude Project instruction file), copy the contents of [`AGENTS.md`](AGENTS.md) into that file so the core behavior rules load automatically and you only need to attach the three document-scoped files each session
5. Tell the AI:
   - The **filename** of the `.md` export and the **section name and number** to work on
   - The **document type** (peer-reviewed paper, technical report, grant proposal, etc.)
   - Your preferred **edit aggressiveness level** (see below)
6. Wait for the AI's clarifying questions before it begins

### Edit Aggressiveness Levels

| Level | What is flagged |
|---|---|
| *Conservative* | Clear errors only: grammar, contradictions, undefined acronyms |
| *Standard* | Errors plus style improvements: parallel structure, unnecessary hedges, informal punctuation *(default)* |
| *Comprehensive* | Everything: flow, word choice, sentence restructuring |

### After Each Session
1. Ask the AI to update [`GRAMMATICAL_RULES_FORWARD.md`](GRAMMATICAL_RULES_FORWARD.md) with any new style or terminology decisions
2. Ask the AI to update [`ITEMS_TO_CHECK.md`](ITEMS_TO_CHECK.md) with any newly flagged acronyms, cross-references, or consistency issues
3. Ask the AI to update [`AI_ERRORS_TO_AVOID.md`](AI_ERRORS_TO_AVOID.md) if any VBA errors were encountered
4. Ask the AI to update [`WRITING_LESSONS_LEARNED.md`](WRITING_LESSONS_LEARNED.md) if any broadly applicable writing insights emerged
5. Run the macro, review the per-edit MsgBox report, investigate any `[FAIL]` lines, and then accept or reject each tracked revision

### Second-Pass Session
After all sections are edited, run a dedicated session using [`ITEMS_TO_CHECK.md`](ITEMS_TO_CHECK.md) as the agenda to resolve outstanding acronym, cross-reference, and consistency issues.

---

## VS Code AI Agent Users (Copilot, Roo Code, Cline, Cursor, Windsurf, etc.)

This workflow is agent-agnostic. All AI behavior rules live in [`AGENTS.md`](AGENTS.md) at the project root. Agent-specific loader files simply point to it:

| Agent | Auto-load file | Notes |
|---|---|---|
| **Codex** | `AGENTS.md` | No loader needed — reads `AGENTS.md` from the repository root natively |
| **Cline** | `AGENTS.md` | No loader needed — reads `AGENTS.md` natively (also supports `.clinerules/`) |
| **Roo Code** | `.roo/rules.md` | Loaded automatically by Roo Code for every session in this workspace |
| **GitHub Copilot** | `.github/copilot-instructions.md` | Auto-injected in Agent/Edits mode; for Chat mode, attach `AGENTS.md` to the session context |
| **Cursor** | `.cursorrules` | Loaded automatically by Cursor |
| **Windsurf** | `.windsurfrules` | Loaded automatically by Windsurf |
| **Claude Code** | `CLAUDE.md` | Loaded automatically by Claude Code |

The agent will also read [`GRAMMATICAL_RULES_FORWARD.md`](GRAMMATICAL_RULES_FORWARD.md), [`ITEMS_TO_CHECK.md`](ITEMS_TO_CHECK.md), [`WRITING_LESSONS_LEARNED.md`](WRITING_LESSONS_LEARNED.md), [`AI_ERRORS_TO_AVOID.md`](AI_ERRORS_TO_AVOID.md), and [`VBA_MACRO_TEMPLATE.bas`](VBA_MACRO_TEMPLATE.bas) automatically at the start of each session — **no manual file attachment is needed**. Simply tell the AI the `.md` filename and which section to work on.

**To add a new AI agent:** create its config file in the location the agent expects, containing only: `Read AGENTS.md in the workspace root and follow all instructions found there.` Then add a row to the table in the [Agent Compatibility](AGENTS.md#agent-compatibility) section of `AGENTS.md`.

---

## Tips

- Work **one section per chat session** to avoid context window degradation. Target roughly 500–3,000 words per session; adjust heading level accordingly
- Track Changes does not need to be on manually — the macro enables it; if already on, this is harmless
- **Always review the per-edit MsgBox report** before accepting tracked changes. Every `[FAIL]` line means an anchor was not found — investigate the cause (smart apostrophe, special Unicode, formatted content, wrong anchor) before moving on
- **Structural edits come in their own macro.** For non-text edits (insertions, paragraph splits, sentence reordering, long deletions), the AI delivers these in a separate macro from text substitutions. This way, if you reject a structural edit, you do not also lose unrelated text-substitution edits. Proposed edits will be labeled with `[text]`, `[insert]`, `[split]`, `[move]`, or `[delete-range]` so you know what kind of change is coming before running the macro
- If the AI's Find/Replace fails on a character (e.g., en dash, middle dot, smart apostrophe), the document likely uses a Unicode character where a plain ASCII one was assumed — the AI should use `ChrW()` codes rather than bare characters for anything above ASCII 255
- For contractions and possessives in VBA search strings, `ChrW(8217)` must be used for the apostrophe — Word AutoCorrect replaces straight apostrophes with curly ones that bare `'` will not match
- Save the Word document after pasting VBA and *before* running `Alt+F8`, or the macro may not appear in the run list
- When asking the AI to edit a figure or table caption, expect it to edit only the caption *text* (the words after `Figure N:` or `Table N:`). It will leave the label and number alone so automatic numbering and cross-references are not disturbed

> **Note for maintainers:** `README.md` is the writer-facing summary of [`AGENTS.md`](AGENTS.md). If writer-relevant workflow steps or file descriptions are updated in [`AGENTS.md`](AGENTS.md), update this file to match.
