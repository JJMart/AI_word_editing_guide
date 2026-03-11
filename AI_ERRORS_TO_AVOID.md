# AI Errors to Avoid — AI-Assisted Editing

This is a living document of errors encountered during AI-assisted VBA editing sessions so they are not repeated in the future. It is **not document-specific** and should travel with the writer across projects.

---

## VBA Errors

*(None logged yet — add entries as errors are encountered)*

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
