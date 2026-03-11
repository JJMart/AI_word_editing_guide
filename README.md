# AI Word Editing Guide

A structured workflow for editing Microsoft Word documents using an AI assistant (e.g., GitHub Copilot, ChatGPT) with changes delivered as tracked revisions via VBA macros. The writer retains full control — every suggested edit appears as a tracked change to accept or reject individually.

---

## How It Works

1. You paste a section of document text into the AI chat
2. The AI asks clarifying questions, then reviews the text and proposes specific edits with rationale
3. The AI writes a Word VBA macro implementing those edits with Track Changes enabled
4. You paste and run the macro in Word — all changes appear as tracked revisions
5. You accept or reject each change individually using the Word Review ribbon

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
2. Open a new AI chat session
3. Attach `AI_WORD_EDITING_GUIDE.md`, `GRAMMATICAL_RULES_FORWARD.md`, and `ITEMS_TO_CHECK.md` using the file attachment feature
4. Tell the AI:
   - The **document type** (peer-reviewed paper, technical report, grant proposal, etc.)
   - The **section name and number** you are working on
   - Your preferred **edit aggressiveness level** (see below)
5. Paste the section text and wait for the AI's clarifying questions before it begins

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
5. Run the macro and review all tracked changes in Word

### Second-Pass Session
After all sections are edited, run a dedicated session using `ITEMS_TO_CHECK.md` as the agenda to resolve outstanding acronym, cross-reference, and consistency issues.

---

## VS Code + GitHub Copilot Users

If you are using this workflow in VS Code with GitHub Copilot Chat, the file `.github/copilot-instructions.md` is included in this repository. It is automatically injected into every Copilot Chat session when you open this folder as a workspace — no manual attachment needed for the core AI behavior rules.

You still need to attach `GRAMMATICAL_RULES_FORWARD.md` and `ITEMS_TO_CHECK.md` manually each session using `#file`, as these change as editing progresses.

If you are using a different AI tool (ChatGPT, Claude, etc.), attach `AI_WORD_EDITING_GUIDE.md` manually at the start of each session instead.

---

## Tips

- Work **one section per chat session** to avoid context window degradation
- Track Changes does not need to be on manually — the macro enables it; if already on, this is harmless
- If the AI's Find/Replace fails on a character (e.g., em dash, middle dot), check whether the document uses a special Unicode character where a plain ASCII one was assumed — copy the character directly from the document
- Save the Word document after pasting VBA and *before* running `Alt+F8`, or the macro may not appear in the run list
