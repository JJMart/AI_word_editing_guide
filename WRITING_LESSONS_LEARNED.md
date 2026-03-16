# Writing Lessons Learned — AI-Assisted Editing

This is a living document of general lessons learned from AI-assisted document editing. It is **not document-specific** and should travel with the writer across projects. Add to it at the end of any session where a useful general insight emerges.

---

## Workflow and Process

- Work one document section per chat session to maintain coherent context and avoid performance degradation near the context window limit
- When using VS Code with GitHub Copilot, guide files (`GRAMMATICAL_RULES_FORWARD.md`, `ITEMS_TO_CHECK.md`, `WRITING_LESSONS_LEARNED.md`, `AI_ERRORS_TO_AVOID.md`) are read automatically at session start — no manual attachment needed. When using other AI tools (ChatGPT, Claude, etc.), attach them manually at the start of each session
- Ask the AI to update guide files at the end of each session — do not rely on the AI to do this automatically
- Keep a dated backup copy of the document before each session as a rollback point; VBA macros making bulk find/replace changes can be difficult to reverse manually
- A dedicated second-pass session focused on `ITEMS_TO_CHECK.md` is more effective than trying to resolve cross-document issues (e.g., acronym ordering) inline during editing sessions

## VBA Macro Editing

- The AI-assisted VBA workflow (find/replace with track changes) is reliable when search strings are sufficiently unique; errors most often stem from non-unique or special-character-mismatched strings
- Always save the Word document after pasting VBA into the editor and before running `Alt+F8` — otherwise the macro may not appear in the run list
- Setting `TrackRevisions = True` in a macro when track changes is already on is harmless (idempotent) — safe to always include
- VBA macros in Word cannot target specific instances of repeated text without using `wdReplaceOne` with careful cursor positioning; prefer editing document text to be unique before running a macro if ambiguity exists

## Working with AI on Technical Writing

- Providing the section name or number when pasting text gives the AI the context needed to log acronym definitions and cross-references accurately
- The AI can flag unexplained acronyms, jargon, and informal phrasing, but the writer retains judgment on domain-specific terminology that may be intentional
- Tracking terminology preferences and stylistic decisions in a running file (`GRAMMATICAL_RULES_FORWARD.md`) prevents inconsistency drift across sections edited in separate sessions

## General Writing

- **Number-unit compound modifiers:** When a number-unit pair precedes and modifies a noun, use one hyphen on the number-unit pair only — e.g., "6-mm wide gap," "7.7-cm tall section," not "6-mm-wide gap." The stricter two-hyphen form is correct but the one-hyphen form is acceptable and consistent with common usage in engineering and scientific literature. Do not flag this as an error. No hyphen is needed at all in predicative position (after a linking verb such as "was," "is," "are") — e.g., "the tank was 1.2 m wide" is correct because the modifier follows the noun it describes rather than directly preceding it.

- **Never introduce em dashes (—) in suggested edits.** AI models use em dashes far more frequently than most writers do, and the average writer does not type them naturally (they require Alt codes or special input methods). If an em dash already exists in the document, flag it and confirm with the writer that it is intentional before leaving it in place — do not silently preserve or add them.
