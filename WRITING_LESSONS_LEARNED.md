# Writing Lessons Learned — AI-Assisted Editing

This is a living document of general lessons learned from AI-assisted document editing. It is **not document-specific** and should travel with the writer across projects. Add to it at the end of any session where a useful general insight emerges.

---

## Workflow and Process

- Work one document section per chat session to maintain coherent context and avoid performance degradation near the context window limit
- Attach relevant guide files (`GRAMMATICAL_RULES_FORWARD.md`, `ITEMS_TO_CHECK.md`) at the start of each session using `#file` so prior decisions carry forward without relying on chat history
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

*(Add general writing insights here as they emerge)*
