Attribute VB_Name = "ProposalEngine"
Option Explicit

' === Overview ===
' - Framework for 3 layout variants and CHC rep handling (Parts A/B/C)
' - Manual prompt flow (Method 2) and an Excel-friendly hook (Method 1)
' - Uses content controls by Tag and/or bookmark pairs for reliable placement
'
' Required template anchors (recommended):
'   ContentControls (by Tag):
'     - ProjectNameHeader           (Header, plain text)
'     - PartA_Container             (Cover page area)
'     - PartB_Container             (Bottom of letter/first page area)
'     - PartC_Names, PartC_Titles   (Last page, bottom-right column)
'
'   Optional bookmark pairs (inclusive delete safety):
'     - cover_start / cover_end
'     - letter_start / letter_end
'
' Data Tags supported (content controls or {{Token}} fallback):
'   Date, QuoteNumber, ProjectName, CompanyName, CustomerName,
'   CustomerJobTitle, CurrentLocation, NewLocation,
'   MainCHCRepName, MainCHCRepTitle, MainCHCRepEmail, MainCHCRepPhone

Public Enum ProposalLayoutType
    Layout1_Full = 1 ' cover + letter + 2 standard
    Layout2_LetterOnly = 2 ' letter + 2 standard
    Layout3_StandardOnly = 3 ' 2 standard only
End Enum

Public Type CHCRep
    Name As String
    Title As String
    Email As String
    Phone As String
    SignaturePath As String ' optional .png
End Type

' --- Manual entrypoint (Method 2) -------------------------------------------
Public Sub MaybePromptOnOpen(ByVal d As Document)
    On Error GoTo CleanExit
    If Not IsProposalDoc(d) Then Exit Sub

    If GetVar(d, "LayoutConfigured") = "1" Then Exit Sub

    Dim layoutChoice As Long, totalReps As Long
    If Not PromptLayoutAndReps(layoutChoice, totalReps) Then Exit Sub

    ApplyLayoutAndReps d, layoutChoice, totalReps
    SetVar d, "LayoutConfigured", "1"

    ' Save once after initial layout
    SafeSave d

CleanExit:
End Sub

Public Sub GenerateWithPrompt()
    Dim d As Document: Set d = ActiveDocument
    MaybePromptOnOpen d
End Sub

Public Sub GenerateProposal()
    GenerateWithPrompt
End Sub

Public Sub FinalizeToDocx()
    On Error Resume Next
    FinalizeExport.FinalizeAndExportToDocx
End Sub

' --- Excel-friendly hooks (Method 1) ----------------------------------------
' Accepts a late-bound dictionary (Scripting.Dictionary or similar) where keys
' match the data tags listed above. Excel can call this directly.
Public Sub ApplyDataFromDictionary(ByVal dict As Object, ByVal layout As Long, Optional ByVal totalReps As Long = 1)
    Dim d As Document: Set d = ActiveDocument
    If layout < 1 Or layout > 3 Then layout = 1
    If totalReps < 1 Then totalReps = 1

    ApplyLayoutAndReps d, layout, totalReps
    FillDataTags d, dict
    SafeSave d
End Sub

' --- Core: apply layout + reps ----------------------------------------------
Public Sub ApplyLayoutAndReps(ByVal d As Document, ByVal layout As Long, ByVal totalReps As Long)
    If totalReps < 1 Then totalReps = 1

    Select Case layout
        Case Layout1_Full
            ' Keep cover and letter
            ' Ensure cover has no header/footer; link the rest to letter
            ConfigureHeaders d, True
        Case Layout2_LetterOnly
            RemoveRegionPairsIfPresent d, "cover_start", "cover_end"
            ConfigureHeaders d, False
        Case Layout3_StandardOnly
            RemoveRegionPairsIfPresent d, "cover_start", "cover_end"
            RemoveRegionPairsIfPresent d, "letter_start", "letter_end"
            ' No dedicated letter section remains; still set header/footer as if letter-style
            ConfigureHeaders d, False
        Case Else
            ' Default to full
            ConfigureHeaders d, True
    End Select

    BuildPartA d, layout, totalReps
    BuildPartB d, layout, totalReps
    BuildPartC d, totalReps
End Sub

Private Sub ConfigureHeaders(ByVal d As Document, ByVal hasCover As Boolean)
    On Error Resume Next
    Dim s As Section, i As Long

    ' If cover exists, clear its header/footer and unlink from next section
    If hasCover And d.Sections.Count >= 1 Then
        With d.Sections(1)
            .Headers(wdHeaderFooterPrimary).Range.Text = ""
            .Footers(wdHeaderFooterPrimary).Range.Text = ""
            If d.Sections.Count >= 2 Then
                d.Sections(2).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
                d.Sections(2).Footers(wdHeaderFooterPrimary).LinkToPrevious = False
            End If
        End With
    End If

    ' Link all subsequent sections to the first non-cover section's header/footer
    For i = IIf(hasCover, 3, 2) To d.Sections.Count
        d.Sections(i).Headers(wdHeaderFooterPrimary).LinkToPrevious = True
        d.Sections(i).Footers(wdHeaderFooterPrimary).LinkToPrevious = True
    Next i
End Sub

Private Sub BuildPartA(ByVal d As Document, ByVal layout As Long, ByVal totalReps As Long)
    If layout <> Layout1_Full Then
        ' Remove any placeholder container if present
        ClearContainer d, "PartA_Container"
        Exit Sub
    End If

    Dim r As Range: Set r = ContainerRange(d, "PartA_Container")
    If r Is Nothing Then Exit Sub

    Dim mainName As String, mainTitle As String, mainPhone As String
    mainName = GetTagOrVar(d, "MainCHCRepName", "")
    mainTitle = GetTagOrVar(d, "MainCHCRepTitle", "")
    mainPhone = GetTagOrVar(d, "MainCHCRepPhone", "")

    r.Text = ""
    AppendLine r, mainName
    AppendLine r, mainTitle
    AppendLine r, mainPhone

    Dim i As Long, extra As Long: extra = totalReps - 1
    For i = 1 To extra
        AppendLine r, ""
        AppendLine r, GetTagOrVar(d, "CHCRep" & (i + 1) & "Name", "[Rep " & (i + 1) & " Name]")
        AppendLine r, GetTagOrVar(d, "CHCRep" & (i + 1) & "JobTitle", "[Rep " & (i + 1) & " Job Title]")
        AppendLine r, GetTagOrVar(d, "CHCRep" & (i + 1) & "Phone", "[Rep " & (i + 1) & " Phone]")
    Next i
End Sub

Private Sub BuildPartB(ByVal d As Document, ByVal layout As Long, ByVal totalReps As Long)
    Dim r As Range: Set r = ContainerRange(d, "PartB_Container")

    ' If no container, attempt a safe fallback at end of target page
    If r Is Nothing Then
        Dim pageIdx As Long
        pageIdx = IIf(layout = Layout1_Full, 2, 1)
        Set r = EndOfPageRange(d, pageIdx)
        If r Is Nothing Then Exit Sub
        r.InsertBreak Type:=wdSectionBreakContinuous
    End If

    r.Text = ""

    Dim tbl As Table
    Set tbl = d.Tables.Add(Range:=r, NumRows:=1, NumColumns:=totalReps)
    tbl.AllowAutoFit = True
    tbl.Rows.Alignment = wdAlignRowCenter

    Dim i As Long
    For i = 1 To totalReps
        With tbl.Cell(1, i).Range
            .Text = ""
            InsertRepSignatureBlock d, tbl.Cell(1, i).Range, i
        End With
    Next i

    ' Ensure a New Page after Part B per spec
    tbl.Range.Collapse wdCollapseEnd
    tbl.Range.InsertBreak Type:=wdPageBreak
End Sub

Private Sub InsertRepSignatureBlock(ByVal d As Document, ByVal atRange As Range, ByVal repIndex As Long)
    Dim nm As String, title As String, sig As String
    If repIndex = 1 Then
        nm = GetTagOrVar(d, "MainCHCRepName", "[Name]")
        title = GetTagOrVar(d, "MainCHCRepTitle", "[Job Title]")
        sig = GetTagOrVar(d, "MainCHCRepSignature", "")
    Else
        nm = GetTagOrVar(d, "CHCRep" & repIndex & "Name", "[Name]")
        title = GetTagOrVar(d, "CHCRep" & repIndex & "JobTitle", "[Job Title]")
        sig = GetTagOrVar(d, "CHCRep" & repIndex & "Signature", "")
    End If

    Dim r As Range: Set r = atRange.Duplicate
    r.Text = ""
    If Len(sig) > 0 And Dir$(sig) <> "" Then
        r.InlineShapes.AddPicture fileName:=sig, LinkToFile:=False, SaveWithDocument:=True
        r.Collapse wdCollapseEnd
        r.InsertParagraphAfter
        r.Collapse wdCollapseEnd
    Else
        AppendLine r, "[Handwritten Signature]"
    End If

    AppendLine r, nm
    AppendLine r, title
End Sub

Private Sub BuildPartC(ByVal d As Document, ByVal totalReps As Long)
    Dim names As String, titles As String

    names = GetTagOrVar(d, "MainCHCRepName", "[Name]")
    titles = GetTagOrVar(d, "MainCHCRepTitle", "[Job Title]")

    Dim i As Long
    For i = 2 To totalReps
        names = names & ", " & GetTagOrVar(d, "CHCRep" & i & "Name", "[Name]")
        titles = titles & ", " & GetTagOrVar(d, "CHCRep" & i & "JobTitle", "[Job Title]")
    Next i

    SetTagIfExists d, "PartC_Names", names
    SetTagIfExists d, "PartC_Titles", titles
End Sub

' --- Data fill ---------------------------------------------------------------
Public Sub FillDataTags(ByVal d As Document, ByVal dict As Object)
    On Error Resume Next
    If dict Is Nothing Then Exit Sub

    ' Ensure Date is present if missing
    If Not dict.Exists("Date") Or Len(Trim$(CStr(dict("Date")))) = 0 Then
        dict("Date") = Format(Date, "mm/dd/yy")
    End If

    Dim key As Variant, v As String
    For Each key In dict.Keys
        v = CStr(dict(key))
        SetTagIfExists d, CStr(key), v
        ReplaceTokenIfPresent d, "{{" & CStr(key) & "}}", v
    Next key
End Sub

' --- Helpers: anchors, tags, tokens -----------------------------------------
Private Function ContainerRange(ByVal d As Document, ByVal tag As String) As Range
    Dim ccs As ContentControls
    On Error Resume Next
    Set ccs = d.SelectContentControlsByTag(tag)
    If Not ccs Is Nothing Then
        If ccs.Count > 0 Then Set ContainerRange = ccs(1).Range
    End If
End Function

Private Sub ClearContainer(ByVal d As Document, ByVal tag As String)
    Dim ccs As ContentControls
    On Error Resume Next
    Set ccs = d.SelectContentControlsByTag(tag)
    If Not ccs Is Nothing Then
        If ccs.Count > 0 Then ccs(1).Range.Text = ""
    End If
End Sub

Private Function EndOfPageRange(ByVal d As Document, ByVal pageIndex As Long) As Range
    On Error Resume Next
    Dim rStart As Range, rNext As Range
    Set rStart = d.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pageIndex)
    Set rNext = d.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pageIndex + 1)
    If Not rStart Is Nothing Then
        If Not rNext Is Nothing Then
            Set EndOfPageRange = rNext
            EndOfPageRange.Collapse wdCollapseStart
        Else
            Set EndOfPageRange = d.Content
            EndOfPageRange.Collapse wdCollapseEnd
        End If
    End If
End Function

Private Sub RemoveRegionPairsIfPresent(ByVal d As Document, ByVal startName As String, ByVal endName As String)
    On Error Resume Next
    If Not d.Bookmarks.Exists(startName) Or Not d.Bookmarks.Exists(endName) Then Exit Sub

    Dim r As Range: Set r = d.Bookmarks(startName).Range.Duplicate
    Dim r2 As Range: Set r2 = d.Bookmarks(endName).Range.Duplicate
    If r.Start > r2.End Then
        r.Start = r2.Start: r.End = d.Bookmarks(startName).Range.End
    Else
        r.End = r2.End
    End If
    If r.Start < r.End Then r.Delete
End Sub

Private Sub SetTagIfExists(ByVal d As Document, ByVal tag As String, ByVal value As String)
    On Error Resume Next
    Dim ccs As ContentControls
    Set ccs = d.SelectContentControlsByTag(tag)
    If Not ccs Is Nothing Then
        If ccs.Count > 0 Then ccs(1).Range.Text = value: Exit Sub
    End If
    ' Also scan by Title
    Set ccs = d.SelectContentControlsByTitle(tag)
    If Not ccs Is Nothing Then
        If ccs.Count > 0 Then ccs(1).Range.Text = value: Exit Sub
    End If
    ' Search headers/footers across sections
    Dim s As Section, h As HeaderFooter, f As HeaderFooter
    For Each s In d.Sections
        For Each h In s.Headers
            If Not h Is Nothing Then
                Set ccs = h.Range.ContentControls
                If Not ccs Is Nothing Then
                    Dim cc As ContentControl
                    For Each cc In ccs
                        If LCase$(cc.Tag) = LCase$(tag) Or LCase$(cc.title) = LCase$(tag) Then
                            cc.Range.Text = value
                            Exit Sub
                        End If
                    Next cc
                End If
            End If
        Next h
        For Each f In s.Footers
            If Not f Is Nothing Then
                Set ccs = f.Range.ContentControls
                If Not ccs Is Nothing Then
                    Dim cc2 As ContentControl
                    For Each cc2 In ccs
                        If LCase$(cc2.Tag) = LCase$(tag) Or LCase$(cc2.title) = LCase$(tag) Then
                            cc2.Range.Text = value
                            Exit Sub
                        End If
                    Next cc2
                End If
            End If
        Next f
    Next s
End Sub

Private Function GetTagOrVar(ByVal d As Document, ByVal tag As String, ByVal fallback As String) As String
    On Error Resume Next
    Dim ccs As ContentControls
    Dim val As String
    Set ccs = d.SelectContentControlsByTag(tag)
    If Not ccs Is Nothing Then
        If ccs.Count > 0 Then
            val = ccs(1).Range.Text
            If ccs(1).ShowingPlaceholderText Then val = ""
        End If
    End If
    If Len(val) = 0 Then val = d.Variables(tag).Value
    If Len(val) = 0 Then val = fallback
    GetTagOrVar = val
End Function

Private Sub ReplaceTokenIfPresent(ByVal d As Document, ByVal token As String, ByVal value As String)
    On Error Resume Next
    Dim r As Range: Set r = d.Content
    With r.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = token: .Replacement.Text = value
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Function IsProposalDoc(ByVal d As Document) As Boolean
    On Error Resume Next
    IsProposalDoc = (d.Variables("IsProposalDoc").Value = "1")
End Function

Private Function GetVar(ByVal d As Document, ByVal nm As String) As String
    On Error Resume Next
    GetVar = d.Variables(nm).Value
End Function

Private Sub SetVar(ByVal d As Document, ByVal nm As String, ByVal val As String)
    On Error Resume Next
    d.Variables(nm).Value = val
End Sub

Private Sub SafeSave(ByVal d As Document)
    On Error Resume Next
    d.Save
End Sub

' --- UI prompt helpers -------------------------------------------------------
Public Function PromptLayoutAndReps(ByRef layoutChoice As Long, ByRef totalReps As Long) As Boolean
    On Error GoTo Fallback
    Dim f As ProposalLayoutForm
    Set f = New ProposalLayoutForm
    PromptLayoutAndReps = f.ShowAndGet(layoutChoice, totalReps)
    Exit Function
Fallback:
    ' Fallback to simple inputs if the form is not available
    Dim s As String
    s = InputBox( _
        "Choose Layout (1-3):" & vbCrLf & _
        "  1 = Cover + Letter + 2 standard" & vbCrLf & _
        "  2 = Letter + 2 standard" & vbCrLf & _
        "  3 = 2 standard only", _
        "Proposal Layout", "1")
    If Len(Trim$(s)) = 0 Then Exit Function
    layoutChoice = Val(s): If layoutChoice < 1 Or layoutChoice > 3 Then layoutChoice = 1

    s = InputBox( _
        "Total number of CH Coakley Representatives that will sign this proposal" & vbCrLf & _
        "Enter whole number (1 = just me).", _
        "Total Signers", "1")
    If Len(Trim$(s)) = 0 Then Exit Function
    totalReps = Val(s): If totalReps < 1 Then totalReps = 1
    PromptLayoutAndReps = True
End Function

Private Sub AppendLine(ByRef r As Range, ByVal s As String)
    r.Collapse wdCollapseEnd
    r.Text = r.Text & s & vbCrLf
End Sub
