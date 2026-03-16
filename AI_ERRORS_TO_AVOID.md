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
