' =============================================================================
' VBA MACRO TEMPLATE - AI Word Editing Workflow
' =============================================================================
' This file contains:
'   1. TestSetup           - one-time verification macro for first-run writers
'   2. ReviewEdits_<name>  - canonical edit-macro skeleton (the main template)
'   3. Reference patterns  - commented examples of Find/Replace, insertion,
'                            reordering, paragraph split, and anchor-range
'                            deletion (for reference only, not executed)
'
' The AI copies the ReviewEdits template verbatim and fills ONLY the
' EDIT BLOCKS region. Do NOT modify the HEADER or FOOTER regions - they
' are required for consistent per-edit reporting.
'
' VBA rules (mirror of AGENTS.md - do not deviate):
'   - Use ChrW() (not Chr()) for any Unicode character above code point 255.
'   - Use ChrW(8217) for the apostrophe in contractions and possessives.
'   - Set .MatchCase = True when the found text starts uppercase but the
'     replacement should be lowercase (otherwise Word auto-capitalizes).
'   - Find.Text is limited to ~255 characters. For longer deletions use the
'     anchor-range-delete pattern (see Reference Patterns below).
'   - oDoc.Content.Find covers the main body only - not footnotes, endnotes,
'     headers, footers, or comments. Use oDoc.Footnotes(i).Range.Find etc.
'     for those story ranges, and flag it to the writer first.
'   - Never include "Attribute VB_Name = ..." when pasting into the editor.
' =============================================================================


' =============================================================================
' TESTSETUP - First-run verification macro
' =============================================================================
' Run this once on a new machine or with a new document to confirm:
'   (a) macros are enabled and the VBA editor accepts pasted code,
'   (b) Track Changes can be turned on programmatically,
'   (c) a Find/Replace insertion produces a visible tracked revision,
'   (d) the change can be cleanly undone.
'
' The macro inserts a visible marker at the top of the document as a tracked
' insertion, then immediately undoes it. If the writer sees the MsgBox "[OK]"
' message and no stray text remains in the document, the setup is good.
' =============================================================================

Sub TestSetup()
    Dim oDoc As Document
    Dim sMsg As String
    Dim bTracked As Boolean
    Dim oMarker As Range

    On Error GoTo ErrHandler

    Set oDoc = ActiveDocument
    bTracked = oDoc.TrackRevisions
    oDoc.TrackRevisions = True

    ' Insert a visible marker at the very start of the document as a tracked
    ' insertion. Using a distinctive string so the undo is unambiguous.
    Set oMarker = oDoc.Range(0, 0)
    oMarker.InsertBefore "[[VBA_SETUP_TEST_MARKER]]"

    ' Immediately undo the insertion. In TrackRevisions mode the insertion
    ' itself is a tracked revision; Undo removes both the text and the
    ' revision record, leaving the document identical to before.
    oDoc.Undo

    ' Belt-and-braces: a single Undo normally reverts the insertion, but if
    ' AutoCorrect/AutoFormat split the insertion into more than one undo unit,
    ' Undo could leave a fragment. Explicitly search for any remaining marker
    ' text and delete it so the document is guaranteed clean.
    Dim oLeftover As Range
    Set oLeftover = oDoc.Content.Duplicate
    With oLeftover.Find
        .ClearFormatting
        .Text = "[[VBA_SETUP_TEST_MARKER]]"
        .Wrap = wdFindStop
        Do While .Execute
            oLeftover.Delete
            Set oLeftover = oDoc.Content.Duplicate
            .Text = "[[VBA_SETUP_TEST_MARKER]]"
            .Wrap = wdFindStop
        Loop
    End With

    ' Restore the writer's original TrackRevisions setting.
    oDoc.TrackRevisions = bTracked

    sMsg = "[OK] VBA setup verified." & vbCrLf & vbCrLf
    sMsg = sMsg & "- Macros are enabled and executing." & vbCrLf
    sMsg = sMsg & "- Track Changes toggled successfully." & vbCrLf
    sMsg = sMsg & "- Tracked insertion and undo both worked." & vbCrLf & vbCrLf
    sMsg = sMsg & "If you see any '[[VBA_SETUP_TEST_MARKER]]' text in the" & vbCrLf
    sMsg = sMsg & "document, remove it manually - Undo did not fully revert."
    MsgBox sMsg
    Exit Sub

ErrHandler:
    ' Best-effort cleanup if anything failed midway.
    On Error Resume Next
    oDoc.TrackRevisions = bTracked
    MsgBox "[FAIL] TestSetup error " & Err.Number & ": " & Err.Description & _
           vbCrLf & vbCrLf & "Check that macros are enabled and that a" & _
           vbCrLf & "document is open. Remove any '[[VBA_SETUP_TEST_MARKER]]'" & _
           vbCrLf & "text from the document if present."
End Sub


' =============================================================================
' REVIEWEDITS - Canonical edit-macro skeleton
' =============================================================================
' Usage:
'   1. Copy this Sub verbatim.
'   2. Rename it to reflect the section being edited
'      (e.g. ReviewEdits_2_1_Methods).
'   3. Paste one Edit Block per proposed edit inside the EDIT BLOCKS region.
'   4. Number edits sequentially (Edit 1, Edit 2, ...) and give each a
'      one-line rationale comment.
'   5. Every Edit Block must call the If/Else logging pattern so the MsgBox
'      report shows per-edit [OK] or [FAIL].
'   6. See the "Reference Patterns" section at the bottom of this file for
'     examples of Find/Replace, insertion, reordering, paragraph split, and
'     anchor-range deletion.
' =============================================================================

Sub ReviewEdits_SectionName()

    ' --- HEADER (do not modify) ---------------------------------------------
    Dim oDoc As Document
    Dim sMsg As String
    Dim nOK As Long
    Dim nFail As Long

    Set oDoc = ActiveDocument
    oDoc.TrackRevisions = True

    nOK = 0
    nFail = 0
    sMsg = "Edit Report" & vbCrLf
    sMsg = sMsg & String(60, "-") & vbCrLf
    ' ------------------------------------------------------------------------


    ' =======================================================================
    ' EDIT BLOCKS - AI fills this section. One block per proposed edit.
    ' =======================================================================
    '
    ' See Reference Patterns at the bottom of this file for the full template
    ' of each edit type. The most common is Find/Replace:
    '
    '     ' Edit 1: <one-line rationale>
    '     ' Pre-count occurrences so over-replacement (non-unique anchor) is visible.
    '     Dim nHits1 As Long
    '     nHits1 = 0
    '     With oDoc.Content.Duplicate.Find
    '         .ClearFormatting
    '         .Text = "<search text with enough context to be unique>"
    '         .Wrap = wdFindStop
    '         Do While .Execute
    '             nHits1 = nHits1 + 1
    '             .Parent.Collapse wdCollapseEnd
    '         Loop
    '     End With
    '     With oDoc.Content.Find
    '         .ClearFormatting
    '         .Replacement.ClearFormatting
    '         .Text = "<search text with enough context to be unique>"
    '         .Replacement.Text = "<replacement text>"
    '         .MatchCase = False          ' True if replacement case differs from found
    '         .MatchWholeWord = False
    '         .MatchWildcards = False
    '         .Wrap = wdFindContinue
    '         .Forward = True
    '         If .Execute(Replace:=wdReplaceAll) Then
    '             nOK = nOK + 1
    '             sMsg = sMsg & "[OK]   Edit 1: <short description> (replaced " & nHits1 & ", expected 1)" & vbCrLf
    '         Else
    '             nFail = nFail + 1
    '             sMsg = sMsg & "[FAIL] Edit 1: anchor not found - <short description>" & vbCrLf
    '         End If
    '     End With
    '
    ' =======================================================================
    ' END EDIT BLOCKS
    ' =======================================================================


    ' --- FOOTER (do not modify) ---------------------------------------------
    sMsg = sMsg & String(60, "-") & vbCrLf
    sMsg = sMsg & "Total succeeded: " & nOK & vbCrLf
    sMsg = sMsg & "Total failed:    " & nFail & vbCrLf
    sMsg = sMsg & vbCrLf
    sMsg = sMsg & "Review all [FAIL] lines before accepting tracked changes."
    MsgBox sMsg
    ' ------------------------------------------------------------------------

End Sub


' =============================================================================
' REFERENCE PATTERNS - examples only, do not execute
' =============================================================================
' The patterns below are reference skeletons for the five supported edit
' types. Copy the relevant pattern into the EDIT BLOCKS region and adapt.
' Every pattern logs [OK] / [FAIL] to sMsg the same way so the MsgBox report
' stays consistent.
'
' Structural edits (insertion, reordering, paragraph split, anchor-range
' delete) are riskier than Find/Replace because they change document
' structure, not just text. Recommend running structural edits in their own
' macro (not batched with Find/Replace edits) so the writer can reject the
' whole macro and re-try without losing unrelated changes. Flag structural
' edits explicitly to the writer when proposing them.
' =============================================================================

' --- PATTERN 1: Find/Replace (text substitution) ----------------------------
' Use for: rewording, removing hedges, fixing typos, swapping terminology.
'
' ' Edit N: <rationale>
' ' Pre-count occurrences so over-replacement (non-unique anchor) is visible.
' ' wdReplaceAll returns True for one OR many replacements, so without a count
' ' a single [OK] could hide accidental edits at unintended locations.
' Dim nHitsN As Long
' nHitsN = 0
' With oDoc.Content.Duplicate.Find
'     .ClearFormatting
'     .Text = "<unique search text>"
'     .Wrap = wdFindStop
'     Do While .Execute
'         nHitsN = nHitsN + 1
'         .Parent.Collapse wdCollapseEnd
'     Loop
' End With
' With oDoc.Content.Find
'     .ClearFormatting
'     .Replacement.ClearFormatting
'     .Text = "<unique search text>"
'     .Replacement.Text = "<replacement>"
'     .MatchCase = False
'     .MatchWholeWord = False
'     .MatchWildcards = False
'     .Wrap = wdFindContinue
'     .Forward = True
'     If .Execute(Replace:=wdReplaceAll) Then
'         nOK = nOK + 1
'         sMsg = sMsg & "[OK]   Edit N: <description> (replaced " & nHitsN & ", expected 1)" & vbCrLf
'     Else
'         nFail = nFail + 1
'         sMsg = sMsg & "[FAIL] Edit N: anchor not found - <description>" & vbCrLf
'     End If
' End With
'
' NOTE: if the reported count exceeds the expected number, the anchor was not
' unique - the writer must inspect every changed location. When uniqueness is
' certain, the pre-count loop may be omitted and the expected count simply
' stated in the [OK] string (e.g. "...expected 1)").


' --- PATTERN 2: Insertion (add text at an anchor) ---------------------------
' Use for: adding a missing topic sentence, inserting a transitional phrase,
'          adding a missing citation anchor after a claim.
'
' Strategy: find a unique anchor string, collapse the range to one end of it
' (wdCollapseEnd inserts after; wdCollapseStart inserts before), then use
' InsertAfter / InsertBefore to add the new text. The insertion appears as
' a tracked insertion in the document.
'
' ' Edit N: insert missing topic sentence before "The experiment consisted of..."
' Dim oAnchor As Range
' Set oAnchor = oDoc.Content.Duplicate
' With oAnchor.Find
'     .ClearFormatting
'     .Text = "The experiment consisted of"
'     .MatchCase = True
'     .Wrap = wdFindStop
'     .Forward = True
'     If .Execute Then
'         oAnchor.Collapse wdCollapseStart
'         oAnchor.InsertBefore "This section presents the experimental protocol. "
'         nOK = nOK + 1
'         sMsg = sMsg & "[OK]   Edit N: inserted topic sentence" & vbCrLf
'     Else
'         nFail = nFail + 1
'         sMsg = sMsg & "[FAIL] Edit N: anchor not found - insert topic sentence" & vbCrLf
'     End If
' End With


' --- PATTERN 3: Paragraph split (break one paragraph into two) --------------
' Use for: breaking a too-long paragraph at a natural topic boundary.
'
' Strategy: find a unique anchor within the paragraph at the split point,
' collapse to the start of it, insert a paragraph mark (vbCr) before it.
' The second half becomes its own paragraph, as a tracked insertion.
'
' ' Edit N: split paragraph at "In contrast, the second experiment..."
' Dim oSplit As Range
' Set oSplit = oDoc.Content.Duplicate
' With oSplit.Find
'     .ClearFormatting
'     .Text = "In contrast, the second experiment"
'     .MatchCase = True
'     .Wrap = wdFindStop
'     .Forward = True
'     If .Execute Then
'         oSplit.Collapse wdCollapseStart
'         oSplit.InsertBefore vbCr
'         nOK = nOK + 1
'         sMsg = sMsg & "[OK]   Edit N: split paragraph" & vbCrLf
'     Else
'         nFail = nFail + 1
'         sMsg = sMsg & "[FAIL] Edit N: split anchor not found" & vbCrLf
'     End If
' End With


' --- PATTERN 4: Reordering (move a sentence or clause) ---------------------
' Use for: swapping two adjacent sentences, moving a clause to a better
'          position within the same paragraph.
'
' Strategy: find the start and end of the segment to move, copy its text,
' delete the original, find the new insertion anchor, insert the copied text.
' Both the deletion and insertion appear as tracked revisions.
'
' WARNING: reordering produces a "delete + insert" pair in the tracked-
' changes view. Tell the writer to accept both halves together (or reject
' both) - accepting only one half will leave the document in a broken state.
'
' SAFETY: confirm BOTH the source and destination anchors exist BEFORE
' deleting anything. Deleting the source first and only then searching for
' the destination risks a one-sided tracked deletion if the destination
' anchor is mistyped or absent. Mirror the anchor-range-delete pattern below:
' locate both, verify both, then mutate.
'
' ' Edit N: move "Results are summarized in Table 3." to end of paragraph
' Dim oMoveSrc As Range, oMoveDst As Range
' Dim sMoveText As String
' Dim bSrc As Boolean, bDst As Boolean
' Set oMoveSrc = oDoc.Content.Duplicate
' With oMoveSrc.Find
'     .ClearFormatting
'     .Text = "Results are summarized in Table 3. "
'     .MatchCase = True
'     .Wrap = wdFindStop
'     .Forward = True
'     bSrc = .Execute
' End With
' Set oMoveDst = oDoc.Content.Duplicate
' With oMoveDst.Find
'     .ClearFormatting
'     .Text = "<unique anchor at new position>"
'     .MatchCase = True
'     .Wrap = wdFindStop
'     .Forward = True
'     bDst = .Execute
' End With
' If bSrc And bDst Then
'     ' Both anchors confirmed - now it is safe to mutate.
'     sMoveText = oMoveSrc.Text
'     oMoveSrc.Delete
'     ' Re-find the destination after the deletion shifts character positions,
'     ' so the insertion point is still valid.
'     Set oMoveDst = oDoc.Content.Duplicate
'     With oMoveDst.Find
'         .ClearFormatting
'         .Text = "<unique anchor at new position>"
'         .MatchCase = True
'         .Wrap = wdFindStop
'         .Forward = True
'         If .Execute Then
'             oMoveDst.Collapse wdCollapseEnd
'             oMoveDst.InsertAfter " " & sMoveText
'             nOK = nOK + 1
'             sMsg = sMsg & "[OK]   Edit N: moved sentence" & vbCrLf
'         Else
'             nFail = nFail + 1
'             sMsg = sMsg & "[FAIL] Edit N: destination lost after delete - REJECT this macro and re-try" & vbCrLf
'         End If
'     End With
' Else
'     nFail = nFail + 1
'     sMsg = sMsg & "[FAIL] Edit N: move aborted before any change (source=" & bSrc & ", dest=" & bDst & ")" & vbCrLf
' End If


' --- PATTERN 5: Anchor-range delete (long deletion > 200 chars) -------------
' Use for: deleting a passage too long to fit in Find.Text (~255 char limit).
'
' Strategy: find a short unique anchor at the START of the passage, find a
' short unique anchor at the END, build a Range between them, and delete it.
' With TrackRevisions = True this is a tracked deletion.
'
' ' Edit N: delete obsolete background passage from "In the early 2000s..."
' '         through "...which is no longer relevant."
' Dim oStart As Range, oEnd As Range, oDel As Range
' Dim bS As Boolean, bE As Boolean
' Set oStart = oDoc.Content.Duplicate
' With oStart.Find
'     .ClearFormatting
'     .Wrap = wdFindStop
'     .Text = "In the early 2000s"
'     .MatchCase = True
'     bS = .Execute
' End With
' Set oEnd = oDoc.Content.Duplicate
' With oEnd.Find
'     .ClearFormatting
'     .Wrap = wdFindStop
'     .Text = "which is no longer relevant."
'     .MatchCase = True
'     bE = .Execute
' End With
' If bS And bE Then
'     Set oDel = oDoc.Range(oStart.Start, oEnd.End)
'     oDel.Delete
'     nOK = nOK + 1
'     sMsg = sMsg & "[OK]   Edit N: deleted obsolete passage" & vbCrLf
' Else
'     nFail = nFail + 1
'     sMsg = sMsg & "[FAIL] Edit N: delete anchors not found (start=" & bS & ", end=" & bE & ")" & vbCrLf
' End If
