# AI Errors to Avoid — AI-Assisted Editing

This is a living document of errors encountered during AI-assisted VBA editing sessions so they are not repeated in the future. It is **not document-specific** and should travel with the writer across projects.

---

## VBA Errors

### Chr() used for Unicode character above code point 255; Runtime error '5'
**What happened:** `.Replacement.Text` used `Chr(8211)` to insert an en dash (–).  
**Why it failed:** VBA's `Chr()` function only accepts values 0–255 (ANSI range). Code point 8211 is out of range, causing runtime error '5': Invalid procedure call or argument.  
**Fix / how to avoid:** Use `ChrW()` for any character above code point 255. `ChrW(8211)` = en dash (–); `ChrW(8212)` = em dash (—). Use `Chr()` only for standard ANSI characters.

---

### Too many line continuations in MsgBox; Compile error
**What happened:** A `MsgBox` call used more than 24 `& _` line continuations to build a long summary string in a single statement.  
**Why it failed:** VBA allows a maximum of **24 line continuations** (`& _`) per logical line. Exceeding this limit causes the compile error "Too many line continuations" immediately on paste.  
**Fix / how to avoid:** Never build a long string inside a single `MsgBox` call. Instead:
1. `Dim sMsg As String` in the variable declarations
2. Build the string with separate `sMsg = sMsg & "..."` lines (no continuations needed)
3. Call `MsgBox sMsg` as the final statement

This pattern has no continuation limit and is always safe for summary messages of any length.

---

### Find.Text string too long; Runtime error '5854'
**What happened:** `Find.Text` was assigned a multi-sentence string built with `& _` concatenation (~450 chars).  
**Why it failed:** Word VBA's `Find.Text` property has a hard limit of approximately 255 characters. Assigning a longer string raises runtime error '5854': String parameter too long.  
**Fix / how to avoid:** For deletions longer than ~200 characters, do not use Find/Replace. Instead:
1. Find a short unique string near the **start** of the block → `oStart.Find.Execute` → range collapses to that anchor
2. Find a short unique string near the **end** of the block → `oEnd.Find.Execute` → range collapses to that anchor
3. Build a deletion range: `Set oDel = oDoc.Range(oStart.Start, oEnd.End)`
4. Call `oDel.Delete` — this creates a tracked deletion when `TrackRevisions = True`

Pattern skeleton:
```vb
Dim oStart As Range, oEnd As Range, oDel As Range
Dim bS As Boolean, bE As Boolean
Set oStart = oRange.Duplicate
With oStart.Find
    .ClearFormatting : .Wrap = wdFindStop
    .Text = "[short unique start anchor]"
    bS = .Execute
End With
Set oEnd = oRange.Duplicate
With oEnd.Find
    .ClearFormatting : .Wrap = wdFindStop
    .Text = "[short unique end anchor]"
    bE = .Execute
End With
If bS And bE Then
    Set oDel = oDoc.Range(oStart.Start, oEnd.End)
    oDel.Delete
End If
```

---

### Word auto-capitalizes replacement text when found text starts with uppercase and MatchCase = False
**What happened:** A replacement like `"IP-based"` → `"network-connected"` produced `"Network-connected"` in the document.  
**Why it failed:** When `.MatchCase = False`, Word's Find & Replace infers the capitalization pattern of the found text and applies it to the replacement. Found text starting with uppercase → replacement first letter uppercased. Found text ALL CAPS → replacement ALL CAPS.  
**Fix / how to avoid:** Use `.MatchCase = True` any time the found text starts with an uppercase letter but the replacement should be lowercase (e.g., replacing an acronym mid-sentence with a lowercase phrase). Only use `.MatchCase = False` when the search must be case-insensitive AND the replacement capitalization matches the found text.

---

### Smart apostrophes not matched by straight apostrophe in Find.Text
**What happened:** `Find.Text` used a straight apostrophe (`'`) in contractions like "wasn't". The macro ran without error but those Find blocks matched nothing.  
**Why it failed:** Word AutoCorrect replaces straight apostrophes with curly/smart right single quotation marks (Unicode U+2019). The literal `'` in a VBA string literal does not match the curly `'` in the document.  
**Fix / how to avoid:** When searching for contractions or possessives, always use `ChrW(8217)` for the apostrophe in `.Text`. Example:
```vb
.Text = "wasn" & ChrW(8217) & "t"   ' matches Word's smart apostrophe
```
The same applies to left single quotation mark (U+2018): `ChrW(8216)`. Never use a bare `'` inside `.Text` strings for contractions.

---

### `Attribute VB_Name` line causes compile syntax error on paste
**What happened:** The macro file began with `Attribute VB_Name = "MacroName"`.
**Why it failed:** This attribute is a file-level declaration handled internally by VBA. It is only valid when a module is **imported** via File > Import File. Pasting the line directly into the VBA editor causes an immediate compile syntax error.
**Fix / how to avoid:** Never include `Attribute VB_Name = ...` in generated macro code. Start the file with the comment block (`' ===...`) or directly with `Sub MacroName()`.

### Tracked-change revision boundary between search string components; Find fails on anchor spanning boundary
**What happened:** A multi-part search string (e.g., `"some text" & ChrW(8212) & "more text"`) failed to match even though all individual characters were confirmed present by a diagnostic macro.
**Why it failed:** A tracked-change revision boundary from a prior co-author edit existed between two adjacent characters in the document. Word's `Find` engine cannot match a search string that spans an internal revision boundary — the characters are in different revision runs, so the concatenated string never matches as a unit. This is distinct from a character-type mismatch: the characters themselves are correct; the boundary is the problem.
**Why this happens in collaborative documents:** When a document is converted to `.md` using `--track-changes=accept`, the markdown shows clean accepted text. However, the source Word document retains all unaccepted tracked changes from co-author reviews. Any location where a reviewer made a tracked insertion or deletion creates a revision boundary in the document story that `Find` cannot cross. The `.md` export does not reveal these boundaries — the text looks continuous.
**Fix / how to avoid:** When a Find anchor crosses a tracked-change revision boundary, shorten the search string to start or end within a single revision run, avoiding the boundary. To diagnose whether a boundary exists, write a short VBA loop that reads `AscW` character-by-character after a known anchor — boundaries appear as empty ranges (zero-length `oChar.Text`). If the character code matches what you're searching for, the problem is a boundary, not the character. **General rule:** If a macro consistently fails on anchors that appear correct in the `.md`, suspect tracked-change boundaries at those locations and shorten the search anchor.

### Reordering edit deleted the source before confirming the destination; one-sided tracked change left behind
**What happened:** A sentence-move (reordering) macro deleted the source sentence as a tracked deletion, then searched for the destination anchor. The destination anchor was mistyped, so the insertion never happened — the document was left with a tracked deletion and no corresponding insertion.
**Why it failed:** The reordering pattern performed the destructive step (`oMoveSrc.Delete`) before the step that can fail (finding the destination). A move is a delete + insert pair; if the insert half cannot run, the delete half is orphaned.
**Fix / how to avoid:** Always confirm BOTH anchors exist before mutating anything. Find the source and the destination, verify both matched (`If bSrc And bDst Then`), and only then delete the source and insert at the destination. This mirrors the anchor-range-delete pattern, which tests `bS And bE` before deleting. The reordering reference pattern in `VBA_MACRO_TEMPLATE.bas` follows this safe ordering. If the destination is somehow lost after the delete (rare — only if the delete disturbed the destination anchor text), the macro logs a `[FAIL]` telling the writer to reject the whole macro and re-try rather than accept a half-applied move.

### Two-phase macro required "Accept All Changes" between phases and destroyed the tracked-change audit trail
**What happened:** A comment-annotated edit was built as two phases: Phase 1 applied the text changes with Track Changes ON; Phase 2 attached explanatory comments. The MsgBox between phases told the writer to "Accept All Changes" before running Phase 2, because Phase 2 searched for the corrected (new) values to anchor the comments.
**Why it failed:** The corrected values only exist as tracked insertions until accepted, so Phase 2 could not reliably find them while changes were still pending. Forcing "Accept All Changes" to make them findable destroyed exactly the audit trail the writer wanted to send to co-reviewers — the old values and the tracked-change record were gone, leaving only comments on already-accepted text. The design conflated two separable concerns: (a) recording the change, and (b) explaining it.
**Fix / how to avoid:** Use a single-pass **find → comment → change** pattern instead. For each correction: (1) find the OLD (current) text and capture it as a `Range`; (2) `oDoc.Comments.Add Range:=oRng, Text:="..."` to anchor the explanation to the original text; (3) `oRng.Text = "<new value>"` to overwrite, recorded as a tracked deletion + insertion under `TrackRevisions = True`. The comment stays anchored to the strikethrough original, so reviewers see the old value, new value, and rationale together with **no acceptance step required**. Use the `DoEdit` helper sub (Reference Pattern 6 in `VBA_MACRO_TEMPLATE.bas`) so each correction stays one logical line and logs `[OK]`/`[FAIL]`. A comment-only flag (no text change) reuses the same helper by passing identical find and replace text.

### Walk-by-occurrence loops re-matched strikethrough (tracked-deleted) text and skipped live targets
**What happened:** A loop intended to comment/change the Nth occurrence of a repeated value (e.g., the 4th and 5th instances of `"0.2703 (0.0730)"` across table rows) walked forward by hit count and called `oRng.Text = "<new>"` on the counted hits. The first edit landed correctly, but the "next" edit landed on the same location again — the actual later target was skipped, and two near-identical comments ended up on the same spot.
**Why it failed:** Under `TrackRevisions = True`, `oRng.Text = "..."` does not remove the original characters; it marks them as a tracked deletion (strikethrough) and inserts the new text alongside. The strikethrough characters remain in the document story, so Word's `Find` engine re-matches them on the next iteration. The "next" hit is the deletion mark left by the previous edit, not a new live occurrence — so a hit-counting walk counts each location twice and never advances to subsequent ones.
**Fix / how to avoid:** Do not walk by occurrence count when the same edit applies to multiple identical targets. Either (a) anchor each instance with unique surrounding text (a neighboring cell value, row label, or paragraph-intro phrase) so each `Find` is unambiguous, or (b) bracket all targets in one `Range` and run `Find.Execute Replace:=wdReplaceAll` on a fresh copy of that Range — `wdReplaceAll` handles strikethrough re-matching internally. For per-instance comments, unique anchors are required because `wdReplaceAll` cannot attach distinct comments.

---

## Template for Logging an Error

```
### [Date] — [Short description]
**What happened:** ...
**Why it failed:** ...
**Fix / how to avoid:** ...
```

---

## Known Pitfalls (Preemptive)

### Formatting bleed-through between Find/Replace blocks
**Issue:** If `.ClearFormatting` is not called before each `Find` block, formatting from a previous replacement (e.g., bold, italic) can carry over to subsequent replacements.  
**Fix:** Always call `.ClearFormatting` and `.Replacement.ClearFormatting` at the start of every `With .Find` block.

### Non-unique search strings
**Issue:** If the search text appears more than once in the document, `wdReplaceAll` will change every instance, not just the intended one.  
**Fix:** Use a longer, more unique anchor string (include surrounding words). Switch to `wdReplaceOne` if only one instance should change.

### Special characters in VBA strings
**Issue:** Characters like smart quotes (`""`), em dashes (`—`), or middle dots (`·`) must match the exact character in the document. Copy-paste from the document rather than typing manually.  
**Fix:** If a Find fails, check whether the document uses a special Unicode character where a plain ASCII character was assumed.

### Macro not found on Alt+F8
**Issue:** After pasting and closing the VBA editor, the macro may not appear in the Alt+F8 list if the document was not saved after the module was inserted.  
**Fix:** Save the document (`Ctrl+S`) after pasting into the VBA editor and before running Alt+F8.

### Tilde strikethrough rendering in chat
**Issue:** Two `~` characters separated by text in a chat message (e.g., two approximate values like `~0.06` and `~0.12` in the same sentence) can trigger Markdown strikethrough formatting, making the text between them appear struck through.
**Fix:** This is a display artifact in the chat only — it does not affect the VBA string literals. No action needed on the macro.

### Tracked-change boundaries in collaborative documents cause Find to fail on correct anchors
**Issue:** When a document has unaccepted tracked changes from co-author reviews, those revisions create internal revision boundaries in the document story. Word's `Find` engine cannot match a search string that spans one of these boundaries — the string looks correct in the `.md` export (which used `--track-changes=accept`) but fails at runtime because the characters are in different revision runs.
**Symptoms:** A `Find.Execute` returns False even though the search string exactly matches the visible text. The `AscW` of individual characters is correct when checked by diagnostic.
**Fix:** Shorten the search string to start or end within a single revision run, avoiding the boundary. To diagnose, write a short VBA loop reading `AscW` character-by-character after a known anchor — boundaries appear as zero-length ranges (empty `oChar.Text`). This issue only affects documents with unaccepted tracked changes — it does not occur if all changes are accepted before running macros.
**Additional pattern — possessives:** Smart apostrophes in possessives (e.g., `"word" & ChrW(8217) & "s"`) are a common boundary location when a co-author modified text near a possessive. If a Find anchor fails on a possessive, shorten the anchor to start after the apostrophe+s rather than including the possessive noun in the search string.
**Prevention via dual `.md` exports:** For a shared document with pending tracked changes, request two pandoc exports and read both — an accepted view (`--track-changes=accept`) as the drafting target, and an all-changes view (`--track-changes=all`) that shows co-author insertions (`{++...++}`) and deletions (`{--...--}`) inline. Before drafting an anchor, check the all-changes view: do not anchor on or inside text a co-author has already inserted or struck through, and treat text immediately adjacent to a revision marker as a likely boundary. This prevents most boundary failures before the macro is ever run, and prevents editing content the co-author has already deleted or modified.
