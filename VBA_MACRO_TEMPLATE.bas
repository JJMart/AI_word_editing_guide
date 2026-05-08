' =============================================================================
' VBA MACRO TEMPLATE - AI Word Editing Workflow
' =============================================================================
' This is the canonical skeleton for every macro produced by the AI.
' Copy it verbatim. Fill ONLY the EDIT BLOCKS section. Do NOT modify the
' HEADER or FOOTER regions - they are required for consistent edit reporting.
'
' Usage:
'   1. Rename the Sub to reflect the section being edited
'      (e.g. ReviewEdits_2_1_Methods).
'   2. Paste one Edit Block per proposed edit inside the EDIT BLOCKS region.
'   3. Number edits sequentially (Edit 1, Edit 2, ...) and give each a
'      one-line rationale comment.
'   4. Every Edit Block must end with the If .Execute / Else / End If pattern
'      so the MsgBox report shows per-edit [OK] or [FAIL].
'
' VBA rules (mirror of agent.md - do not deviate):
'   - Use ChrW() (not Chr()) for any Unicode character above code point 255.
'   - Use ChrW(8217) for the apostrophe in contractions and possessives.
'   - Set .MatchCase = True when the found text starts uppercase but the
'     replacement should be lowercase (otherwise Word auto-capitalizes).
'   - Find.Text is limited to ~255 characters. For longer deletions use the
'     anchor-range-delete pattern (see AI_ERRORS_TO_AVOID.md).
'   - Never include "Attribute VB_Name = ..." when pasting into the editor.
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
    ' Template for a single Find/Replace edit (copy, number, and adapt):
    '
    '     ' Edit 1: <one-line rationale>
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
    '             sMsg = sMsg & "[OK]   Edit 1: <short description>" & vbCrLf
    '         Else
    '             nFail = nFail + 1
    '             sMsg = sMsg & "[FAIL] Edit 1: anchor not found - <short description>" & vbCrLf
    '         End If
    '     End With
    '
    ' For replace-one-instance cases, use Replace:=wdReplaceOne and anchor with
    ' enough surrounding unique text to guarantee the correct instance.
    '
    ' For deletions longer than ~200 chars, use the anchor-range-delete pattern
    ' documented in AI_ERRORS_TO_AVOID.md - still wrap in If/Else so the result
    ' is logged to sMsg the same way.
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
