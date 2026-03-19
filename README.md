# AI Word Editing Guide

A structured workflow for editing Microsoft Word documents using an AI assistant (e.g., GitHub Copilot, ChatGPT) with changes delivered as tracked revisions via VBA macros. The writer retains full control — every suggested edit appears as a tracked change to accept or reject individually. When the AI is uncertain about a phrasing preference, encounters an ambiguous case, or needs context only the writer can provide, it will ask for clarification or present options rather than guessing.

---

## How It Works

1. You convert your document to Markdown using pandoc and place the `.md` file in the project folder
2. The AI reads the file directly — no need to paste section text manually
3. The AI asks clarifying questions before beginning; during review it will flag any items it is uncertain about and ask the writer to clarify or choose between options before proceeding
4. The AI proposes specific edits with rationale, then writes a Word VBA macro implementing those edits with Track Changes enabled
5. You paste and run the macro in Word — all changes appear as tracked revisions
6. You accept or reject each change individually using the Word Review ribbon

---

## Files in This Repository

| File | Scope | Description |
|---|---|---|
| `AI_WORD_EDITING_GUIDE.md` | Document project | Full guide for both writer and AI, including workflow rules, VBA best practices, and instructions for maintaining the other files. Attach this at the start of each session. |
| `GRAMMATICAL_RULES_FORWARD.md` | Document project | Style and terminology decisions specific to the document being edited (units, capitalization, terminology preferences, acronym tracking). Attach at the start of each session so rules carry forward. |
| `ITEMS_TO_CHECK.md` | Document project | Running checklist of items to verify on a second pass: undefined acronyms, unused acronyms, cross-reference accuracy, numeric consistency, and reference list completeness. |
| `AI_ERRORS_TO_AVOID.md` | **Cross-project** | Living log of errors encountered during AI-assisted VBA editing sessions. Carry this forward to future projects and add to it over time. |
| `WRITING_LESSONS_LEARNED.md` | **Cross-project** | Living document of general writing insights gained through AI-assisted editing. Not document-specific — carry it forward across all projects. |

> **Cross-project files** (`AI_ERRORS_TO_AVOID.md` and `WRITING_LESSONS_LEARNED.md`) are writer-owned and should not be reset when starting a new document project. Copy them into the new project folder and continue adding to them.

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

> **VBA character note:** Unicode characters (en dashes, smart quotes, etc.) are visible in the `.md` file, but you still must use `ChrW()` codes in VBA `Find.Text` strings — never copy those characters literally into VBA string literals. Superscripts and subscripts applied as Word character *formatting* (not true Unicode) will still appear as plain digits or letters — this is expected.

> **Install pandoc:** If pandoc is not already installed, download it from [pandoc.org](https://pandoc.org/installing.html) or install via `winget install JohnMacFarlane.Pandoc` in PowerShell.

---

## Running a VBA Macro in Word

1. **Open the VBA Editor:** Press `Alt+F11`
2. **Insert a new module:** Click `Insert` → `Module`
3. **Paste the macro:** Copy the VBA code from the AI chat and paste it into the blank module window
4. **Return to Word and save:** Close the VBA editor, then save the document (`Ctrl+S`)
5. **Run the macro:** Press `Alt+F8`, select the macro name, and click `Run`
6. **Review tracked changes:** Accept or reject each revision individually using the Review ribbon

> **Enable macros first:** `File → Options → Trust Center → Trust Center Settings → Macro Settings → Enable all macros` (or "Enable with notification")

> **Note for maintainers:** `README.md` is the writer-facing summary of `AI_WORD_EDITING_GUIDE.md`. If writer-relevant workflow steps or file descriptions are updated in that guide, update this file to match.

---

## Session Workflow

### Before Each Session
1. Save a dated backup copy of your `.docx` (e.g., `MyDoc_2026-03-10.docx`)
2. Convert the document to Markdown using the pandoc command above and place the `.md` file in the project folder
3. Open a new AI chat session
4. Attach the guide files:
   - **VS Code + GitHub Copilot users:** no attachment needed — all guide files are read automatically from the workspace
   - **All other AI tools (ChatGPT, Claude, etc.):** attach `AI_WORD_EDITING_GUIDE.md`, `GRAMMATICAL_RULES_FORWARD.md`, and `ITEMS_TO_CHECK.md` — or, if your tool supports a persistent system prompt file (e.g., `claude.md` for Claude Projects), copy the contents of `.github/copilot-instructions.md` into that file so the AI behavior rules load automatically and you only need to attach `GRAMMATICAL_RULES_FORWARD.md` and `ITEMS_TO_CHECK.md` each session
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
1. Ask the AI to update `GRAMMATICAL_RULES_FORWARD.md` with any new style or terminology decisions
2. Ask the AI to update `ITEMS_TO_CHECK.md` with any newly flagged acronyms, cross-references, or consistency issues
3. Ask the AI to update `AI_ERRORS_TO_AVOID.md` if any VBA errors were encountered
4. Ask the AI to update `WRITING_LESSONS_LEARNED.md` if any broadly applicable writing insights emerged
5. Run the macro and verify the edit count in the MsgBox matches the expected number of changes before accepting tracked revisions

### Second-Pass Session
After all sections are edited, run a dedicated session using `ITEMS_TO_CHECK.md` as the agenda to resolve outstanding acronym, cross-reference, and consistency issues.

---

## VS Code + GitHub Copilot Users

If you are using this workflow in VS Code with GitHub Copilot Chat, the file `.github/copilot-instructions.md` is included in this repository. It is automatically injected into every Copilot Chat session when you open this folder as a workspace.

Copilot will also read `GRAMMATICAL_RULES_FORWARD.md`, `ITEMS_TO_CHECK.md`, `WRITING_LESSONS_LEARNED.md`, and `AI_ERRORS_TO_AVOID.md` automatically at the start of each session using its workspace file tools — **no manual file attachment is needed at all**. Simply tell Copilot the `.md` filename and which section to work on.

If you are using a different AI tool (ChatGPT, Claude, etc.), attach `AI_WORD_EDITING_GUIDE.md`, `GRAMMATICAL_RULES_FORWARD.md`, and `ITEMS_TO_CHECK.md` manually at the start of each session instead. If your tool supports a persistent system prompt file (e.g., `claude.md` for Claude Projects), copy the contents of `.github/copilot-instructions.md` into that file so the core AI behavior rules load automatically every session.

---

## Tips

- Work **one section per chat session** to avoid context window degradation
- Track Changes does not need to be on manually — the macro enables it; if already on, this is harmless
- **Always verify the edit count** reported in the macro's MsgBox before accepting tracked changes — if the number seems off, investigate before proceeding
- If the AI's Find/Replace fails on a character (e.g., en dash, middle dot, smart apostrophe), the document likely uses a Unicode character where a plain ASCII one was assumed — the AI should use `ChrW()` codes rather than bare characters for anything above ASCII 255
- For contractions and possessives in VBA search strings, `ChrW(8217)` must be used for the apostrophe — Word AutoCorrect replaces straight apostrophes with curly ones that bare `'` will not match
- Save the Word document after pasting VBA and *before* running `Alt+F8`, or the macro may not appear in the run list
- Do not ask the AI to modify figure or table captions if you use Word's automatic caption numbering — editing those can break cross-references throughout the document
