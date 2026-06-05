# Workflow Logic Review — Findings

*Reviewer: Opus 4.8 logic pass. Date: 2026-06-05. Each finding notes severity, the affected file(s), the problem, and a suggested resolution.*

> **Status: ALL FINDINGS RESOLVED (2026-06-05).** Every finding below was fixed in the same session. A `✅ Resolved` line follows each one describing what was changed. This file is retained as a record of the review and the rationale behind each change.*

Severity legend: **[High]** = can cause a broken macro, lost work, or contradictory AI behavior; **[Medium]** = inconsistency or gap that will cause confusion or occasional failure; **[Low]** = polish, stale reference, or wording.

---

## High-severity findings

### H1. Direct contradiction on `.Wrap` between AGENTS.md and the template's structural patterns
**Files:** `AGENTS.md` (line 79), `VBA_MACRO_TEMPLATE.bas` (Patterns 2–5), `AI_ERRORS_TO_AVOID.md` (anchor-range skeleton).

`AGENTS.md` states as an absolute VBA coding rule: *"Set `.Wrap = wdFindContinue` on every `Find` block."* But four of the five reference patterns in the template (insertion, paragraph split, reordering, anchor-range delete) deliberately use `.Wrap = wdFindStop`, and the `AI_ERRORS_TO_AVOID.md` anchor-range skeleton also uses `wdFindStop`. Only the Find/Replace pattern uses `wdFindContinue`.

This is a real contradiction, not just wording. The template is actually *correct* — `wdFindStop` is the right choice for structural patterns because `wdFindContinue` can wrap past the end of the document and match an unintended earlier/later occurrence when you are locating a single anchor to collapse against. An AI following the AGENTS.md rule literally would change every structural pattern's `.Wrap` to `wdFindContinue` and reintroduce a wrap-around anchor bug.

**Suggested resolution:** Reword the AGENTS.md rule to: "Set `.Wrap = wdFindContinue` on Find/Replace blocks that should scan the whole document; use `.Wrap = wdFindStop` when locating a single anchor for a structural edit (insertion, split, move, delete-range) so the search cannot wrap around and match an unintended occurrence." This aligns the rule with the (correct) template.

✅ **Resolved:** `AGENTS.md` VBA coding rule rewritten to distinguish Find/Replace (`wdFindContinue`) from structural-edit anchor location (`wdFindStop`), and now explicitly notes that the reference patterns follow this split.

### H2. Reordering pattern deletes the source before confirming the destination exists — partial-failure data risk
**Files:** `VBA_MACRO_TEMPLATE.bas` (Pattern 4, lines 277–311).

The reordering skeleton executes `oMoveSrc.Delete` *before* searching for the destination anchor. If the destination `Find` then fails, the macro logs `[FAIL] ... (source already deleted)` — but the source text is already gone (as a tracked deletion). The writer is left with a tracked deletion and no corresponding insertion. The comment honestly notes "source already deleted," but the design still performs a destructive step before the step that can fail.

Because reordering is already flagged as the riskiest edit type and is supposed to run in its own macro, the blast radius is limited — but it still means a single mistyped destination anchor produces a one-sided tracked change that the writer must notice and manually reject.

**Suggested resolution:** Reorder the pattern to find *both* anchors first (capturing `sMoveText` and confirming the destination match) and only call `oMoveSrc.Delete` after both are confirmed. This mirrors the anchor-range-delete pattern (Pattern 5), which correctly tests `bS And bE` before deleting. Worth documenting in `AI_ERRORS_TO_AVOID.md` as a known pitfall as well.

✅ **Resolved:** Pattern 4 in `VBA_MACRO_TEMPLATE.bas` now confirms both `bSrc` and `bDst` before any mutation, deletes only after both are verified, re-finds the destination after the delete shifts positions, and logs a "REJECT this macro and re-try" `[FAIL]` if the destination is somehow lost. A matching pitfall entry was added to `AI_ERRORS_TO_AVOID.md`.

### H3. `wdReplaceAll` + per-edit `[OK]`/`[FAIL]` reporting can silently over-edit or mis-report
**Files:** `AGENTS.md` (lines 71, 82), `VBA_MACRO_TEMPLATE.bas` (Pattern 1).

`.Execute(Replace:=wdReplaceAll)` returns `True` if it made *at least one* replacement — it does not return the count. The reporting pattern logs `[OK]` on `True`. Two failure modes are invisible to the writer:
1. **Over-replacement:** if the anchor string is not unique, `wdReplaceAll` changes every occurrence and still logs a single `[OK]`. The writer sees "succeeded" and has no signal that three places changed instead of one.
2. **The AGENTS.md rule "Include enough surrounding context in search strings to guarantee uniqueness"** is the only thing standing between the writer and silent over-editing, but nothing in the macro *verifies* uniqueness.

**Suggested resolution:** Either (a) default Find/Replace edits to a count-aware pattern (loop `.Execute` without `wdReplaceAll`, increment a counter, and report `[OK] Edit N: replaced K occurrence(s)` so K>1 is visible), or (b) add a sentence to the reporting section telling the AI to report the expected occurrence count in the `[OK]` string and instruct the writer to confirm K matches. Option (a) is more robust; option (b) is a lighter-touch documentation fix.

✅ **Resolved:** Adopted option (a) as the default. The AGENTS.md reporting section, the template EDIT BLOCKS example, and template Pattern 1 now pre-count occurrences and report `(replaced N, expected 1)`, with a note that exceeding the expected count means a non-unique anchor the writer must inspect. The lighter option (b) is documented as acceptable when uniqueness is certain.

---

## Medium-severity findings

### M1. The PDF extraction one-liner can crash on pages with no extractable text
**File:** `README.md` (line 70).

```
python -c "import pypdf; r=pypdf.PdfReader('reference.pdf'); print('\n'.join(p.extract_text() for p in r.pages))" > reference.txt
```

`page.extract_text()` can return `None` for image-only/scanned pages in some pypdf versions, which makes `'\n'.join(...)` raise `TypeError: sequence item N: expected str instance, NoneType found` — the whole command fails and produces an empty or partial `reference.txt`. Ironically this is exactly the scanned-PDF case the quality-check guidance is meant to catch, but here it manifests as a hard crash rather than a garbled file the AI can flag.

**Suggested resolution:** Make the join defensive: `... (p.extract_text() or '' for p in r.pages) ...`. This yields empty strings for unextractable pages (which the AI's quality check then flags as "blank pages") instead of crashing.

✅ **Resolved:** The README one-liner now uses `(p.extract_text() or '')` with an explanatory sentence about the guard.

### M2. Two confirmation-gating models coexist without cross-reference
**Files:** `AGENTS.md` (Session Management; README Sync), the coding-workflow guide this was generalized from.

The writing workflow updates guide files (`GRAMMATICAL_RULES_FORWARD.md`, `ITEMS_TO_CHECK.md`, etc.) only at session end after the writer confirms — a sensible model. The coding workflow this was harmonized with uses an `[UNCONFIRMED]`/`[CONFIRMED]`/`[REVERTED]` changelog marker system precisely because a session can end before the writer verifies a change. The writing workflow has no equivalent: if a session ends after the AI proposes guide-file updates but before the writer runs the macro and confirms, the pending updates are simply lost (the AI was told "do not modify guide files during the session").

This is probably acceptable for the writing workflow (guide-file entries are cheap to regenerate, unlike code dead-ends), but it is an unstated asymmetry. A writer moving between the two workflows may expect the same durability guarantee and not get it.

**Suggested resolution:** Add one sentence to the AGENTS.md Session Management section acknowledging this is intentional: "Unlike the code-documentation workflow, guide-file updates are not written under a pending marker; if a session ends before the writer confirms, restate the proposed updates at the start of the next session." Or, if durability matters, adopt a lightweight pending-marker for `ITEMS_TO_CHECK.md` only.

✅ **Resolved:** Added the acknowledgement sentence to the AGENTS.md "do not modify guide files during the session" bullet, explicitly contrasting with the code-documentation workflow and instructing the AI to restate unwritten proposed updates at the next session start.

### M3. Session-start "read these files" list omits `AI_WRITING_INDICATORS.md`, but the check is mandatory
**Files:** `AGENTS.md` (Session Management lines 96–102 vs. AI Writing Indicators Check lines 120–124).

The session-start reading list deliberately excludes `AI_WRITING_INDICATORS.md` (correctly — line 122 says don't load the full file at session start, only Section 8 when performing the check). But the "AI Writing Indicators Check" is described as something to do "After proposing edits for a section **and** after writing the VBA macro" — i.e., it is mandatory every session. A new AI reading only the session-start list could plausibly treat the indicators check as optional because the file isn't in the must-read list.

**Suggested resolution:** Add a line to the session-start list: "`AI_WRITING_INDICATORS.md` — do not read at session start, but you MUST run the Section 8 check after proposing edits and after writing the macro (see AI Writing Indicators Check below)." This keeps the lazy-load behavior while making the obligation visible from the session-start checklist.

✅ **Resolved:** Added exactly this line to the AGENTS.md session-start reading list.

### M4. "Four `.md` guide files" count is ambiguous / undercounts
**Files:** `AGENTS.md` (line 102, line 153).

The text says "If any of the four `.md` guide files are missing, create them from the starter content." The starter-content section documents exactly four (`GRAMMATICAL_RULES_FORWARD.md`, `ITEMS_TO_CHECK.md`, `AI_ERRORS_TO_AVOID.md`, `WRITING_LESSONS_LEARNED.md`), so the number is internally consistent there. But `AI_WRITING_INDICATORS.md` is also a `.md` guide file that the workflow depends on, and it has *no* starter content and no "create if missing" instruction. If a writer starts a fresh project and copies only the four, the indicators check silently has nothing to read and the AI may either skip it or hallucinate the checklist.

**Suggested resolution:** Either add `AI_WRITING_INDICATORS.md` to the "tell the writer if missing — do not regenerate from memory" category (like `VBA_MACRO_TEMPLATE.bas`), or explicitly note it is a required cross-project file that must be copied into every new project. The key risk is the AI regenerating the indicators list from memory and producing a thinner, less-researched version.

✅ **Resolved:** The "create if missing" bullet now names the four document/cross-project guide files explicitly and places `AI_WRITING_INDICATORS.md` alongside `VBA_MACRO_TEMPLATE.bas` in the "tell the writer — do not regenerate from memory (researched cross-project assets)" category.

### M5. Aggressiveness levels and the tagged edit list can disagree at the boundary
**Files:** `AGENTS.md` ("Before Suggesting Any Edits" lines 15–18 vs. "Types of Edits to Consider" lines 128–148).

The aggressiveness definitions say *Standard* = "errors plus style improvements (parallel structure, hedges, informal punctuation)." The tagged list marks several additional items as `[S]` that are not named in the short definition: "Colloquial or imprecise terms," "Redundant phrasing," "Unexplained jargon." A writer who reads only the short definition at session start and approves "Standard" may be surprised when the AI also flags redundancy and jargon. Not a contradiction (the tagged list is the authoritative superset), but the short definition reads like an exhaustive list when it is actually a sample.

**Suggested resolution:** Change the parenthetical in the Standard definition to "(e.g., parallel structure, hedges, informal punctuation — see the full tagged list under 'Types of Edits to Consider')" so it reads as illustrative, not exhaustive. Same treatment for Conservative and Comprehensive.

✅ **Resolved:** All three aggressiveness definitions now read as illustrative ("e.g., …") and point to the authoritative tagged list, with each level mapped to its tag set (`[C]`; `[C]`+`[S]`; `[C]`+`[S]`+`[X]`).

### M6. Em-dash guidance lives in three places with slightly different scope
**Files:** `AGENTS.md` (line 146), `AI_WRITING_INDICATORS.md` (§7, §8), `WRITING_LESSONS_LEARNED.md` (line 34).

The em-dash rule appears in: the AGENTS.md edit list ("Do not introduce em dashes… flag existing ones"), `AI_WRITING_INDICATORS.md` (where the strongest framing lives — "even one triggers suspicion" in correspondence), and `WRITING_LESSONS_LEARNED.md`. The three are mutually consistent in spirit but differ in scope: AI_WRITING_INDICATORS frames the strongest version as *correspondence-specific* (letters/emails/memos), while AGENTS.md states it as a blanket all-levels rule for any document. For a technical paper or proposal (the workflow's main use case), a single em dash is far less of a tell than in a cover letter, yet the AGENTS.md blanket rule would have the AI flag every existing em dash in a journal article. That may over-trigger on legitimate scholarly prose where em dashes are an accepted (if AI-overused) device.

**Suggested resolution:** This is a judgment call for the writer, not a defect. Decide whether the blanket "flag all em dashes" rule is desired for scholarly documents, or whether it should be softened to "flag em dashes in correspondence unconditionally; in scholarly prose, flag only clusters" — and make the three locations agree. At minimum, have AGENTS.md point to the AI_WRITING_INDICATORS §7 discussion as the authority rather than restating a subtly different scope.

✅ **Resolved:** AGENTS.md now carries a scope note (correspondence = flag every em dash; scholarly prose = flag clusters and AI-inserted ones) and points to `AI_WRITING_INDICATORS.md` §7 as the authority on the distinction. The "never introduce em dashes in edits" rule is unchanged; only the flagging-of-existing scope was reconciled.

---

## Low-severity findings

### L1. Stale `agent.md` reference in `.gitignore` comment
**File:** `.gitignore` (line 9).

The header comment still lists "`agent.md`" among the guide files that are NOT ignored. The file was renamed to `AGENTS.md`. Purely cosmetic (it's a comment, and `AGENTS.md` is not matched by any ignore pattern so it is committed correctly), but it is a stale reference the rename pass missed.

**Suggested resolution:** Update the comment to `AGENTS.md`.

✅ **Resolved:** `.gitignore` header comment updated to `AGENTS.md`.

### L2. README "Files in This Repository" table lists `pypdf` with an empty-looking scope
**File:** `README.md` (line 29).

`pypdf` is given the scope "Tool" in a table whose other rows are "Cross-project" or "Document project." It reads slightly oddly because `pypdf` is not a file in the repository at all — it's an external dependency. Minor, but the table is titled "Files in This Repository."

**Suggested resolution:** Either move `pypdf` (and arguably pandoc) to a small "External tools" note rather than the files table, or rename the table to "Files and Tools." Low priority.

✅ **Resolved:** `pypdf` moved out of the "Files in This Repository" table into a new "External tools (not files in this repository)" sub-table that lists both pandoc and `pypdf`.

### L3. `TestSetup` undo assumption may not hold if AutoCorrect fires
**File:** `VBA_MACRO_TEMPLATE.bas` (TestSetup, lines 57–63).

`TestSetup` inserts `[[VBA_SETUP_TEST_MARKER]]` at `Range(0,0)` then calls `oDoc.Undo` once, assuming a single undo reverts exactly the insertion. If the document's AutoFormat/AutoCorrect settings transform the inserted text (unlikely with this bracketed string, but possible with aggressive AutoCorrect), the insertion could become two undo units and a single `Undo` would leave a fragment. The macro does warn the writer to remove stray marker text if seen, so this is self-healing, but the "single Undo reverts the insertion" assumption is not guaranteed by the Word object model.

**Suggested resolution:** Acceptable as-is given the explicit warning. If hardening is wanted, loop `Undo` until the marker is gone, or delete the inserted range explicitly instead of relying on `Undo`.

✅ **Resolved:** `TestSetup` now follows the `Undo` with an explicit Find/Delete loop that removes any remaining `[[VBA_SETUP_TEST_MARKER]]` fragment, guaranteeing a clean document even if AutoCorrect split the insertion into multiple undo units.

### L4. "Reference Patterns must not be included in delivered macro" relies solely on AI discipline
**Files:** `AGENTS.md` (line 38), `VBA_MACRO_TEMPLATE.bas` (lines 167+).

The reference patterns are commented-out code at the bottom of the template. The instruction says not to include them in the writer-delivered macro. Since they're already comments, an accidental copy-through would be harmless (comments don't execute), so the risk is only verbosity. No change needed; noted for completeness.

### L5. Open VS Code tab references a non-existent / differently-named file
**Observation (not a workflow file):** The editor has tabs open for `AI_WORD_EDITING_GUIDE.md` and `Files_From_Using_Tool_05112026/VBA_MACRO_TEMPLATE.bas`, neither of which exists in the repository root listing. `AI_WORD_EDITING_GUIDE.md` is not referenced by any workflow file (confirmed by search). This is likely a leftover open tab or a renamed file. Not a workflow defect — flagging only so the writer can confirm nothing was meant to be there.

---

## Things that are logically sound (verified, no action needed)

- **The `[UNCONFIRMED]`-style session-boundary problem is correctly avoided for `ITEMS_TO_CHECK.md`** by having the AI re-surface open items at every session start (the strengthened rule). This is the right mechanism for a workflow without changelog markers.
- **The anchor-range-delete pattern (Pattern 5) correctly confirms both anchors (`bS And bE`) before deleting.** This is the safe template that Pattern 4 (reordering) should mirror (see H2).
- **`ChrW()` vs `Chr()`, the 255-char `Find.Text` limit, the 24-line-continuation `MsgBox` limit, smart-apostrophe matching, and the tracked-change-boundary failure mode** are all documented accurately and consistently across `AGENTS.md`, the template header, and `AI_ERRORS_TO_AVOID.md`. These are the workflow's strongest, most battle-tested rules.
- **The caption-editing rule** (edit only caption text, never the `Figure N:` prefix) is sound and consistently stated in both AGENTS.md and README.
- **The story-range limitation** (`oDoc.Content.Find` is body-only) is correctly documented with the right remediation (`StoryRanges` / `Footnotes(i).Range.Find`).
- **The agent-compatibility wiring** (AGENTS.md native for Codex/Cline, loader pointers for the rest) is accurate per current tool documentation.

---

## Summary table

| ID | Severity | One-line | Affected file(s) |
|----|----------|----------|------------------|
| H1 | High | `.Wrap` rule contradicts structural patterns | AGENTS.md vs template |
| H2 | High | Reordering deletes source before confirming destination | VBA_MACRO_TEMPLATE.bas |
| H3 | High | `wdReplaceAll` + boolean reporting hides over-replacement | AGENTS.md, template |
| M1 | Medium | pypdf one-liner crashes on `None` pages | README.md |
| M2 | Medium | No session-boundary durability for guide-file updates | AGENTS.md |
| M3 | Medium | Indicators check mandatory but absent from session-start list | AGENTS.md |
| M4 | Medium | `AI_WRITING_INDICATORS.md` not in "create/flag if missing" set | AGENTS.md |
| M5 | Medium | Aggressiveness short defs read as exhaustive but aren't | AGENTS.md |
| M6 | Medium | Em-dash rule scope differs across three files | AGENTS.md, indicators, lessons |
| L1 | Low | Stale `agent.md` in .gitignore comment | .gitignore |
| L2 | Low | `pypdf` in a table titled "Files in This Repository" | README.md |
| L3 | Low | TestSetup single-Undo assumption | VBA_MACRO_TEMPLATE.bas |
| L4 | Low | Reference-patterns exclusion relies on AI discipline | template |
| L5 | Low | Stray open tab for non-existent guide file | (not a workflow file) |
