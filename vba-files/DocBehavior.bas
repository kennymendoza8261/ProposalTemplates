Attribute VB_Name = "DocBehavior"
Option Explicit

Private Const COPY_RECENCY_MIN As Long = 30  ' minutes

' ---------- Persistent ID ----------
Public Function GetProposalGuid(ByVal d As Document) As String
    On Error Resume Next
    GetProposalGuid = d.CustomDocumentProperties("ProposalGuid").Value
    If Err.Number <> 0 Then GetProposalGuid = ""
    Err.Clear
End Function

' ---------- Save pipeline ----------
Public Sub OnBeforeSave(ByVal d As Document, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)
    If Not IsOurProposal(d) Then Exit Sub
    If SaveAsUI Then
        MintNewProposalGuid d
        SetAuthorToCurrent d
    End If
    UpdateDateContentControls d
End Sub

Public Sub MintNewProposalGuid(ByVal d As Document)
    Dim props As Object, newId As String
    Set props = d.CustomDocumentProperties
    newId = CreateGuidString()
    On Error Resume Next
    props("ProposalGuid").Value = newId
    If Err.Number <> 0 Then
        Err.Clear
        props.Add name:="ProposalGuid", LinkToContent:=False, Type:=4, Value:=newId
    End If
    Err.Clear
End Sub

Public Sub SetAuthorToCurrent(ByVal d As Document)
    On Error Resume Next
    d.BuiltInDocumentProperties("Author").Value = Application.UserName
    Err.Clear
End Sub

' ---------- First-open UX ----------
Public Sub FirstOpenWorkHidden(Optional ByVal CallerGuid As String = "")
    If Not Bootstrap.VerifiedCaller(CallerGuid) Then Exit Sub
    If GetVar(ThisDocument, "CosignerPromptDone") <> "1" Then
        UpdateDateContentControls ThisDocument
        RunCoSignerChoice ThisDocument
        SetVar ThisDocument, "CosignerPromptDone", "1"
    End If
End Sub

Public Sub MaybeRunCosigner(ByVal d As Document)
    If Not IsOurProposal(d) Then Exit Sub
    If GetVar(d, "LayoutConfigured") = "1" Then Exit Sub
    If GetVar(d, "CosignerPromptDone") <> "1" Then
        RunCoSignerChoice d
        SetVar d, "CosignerPromptDone", "1"
    End If
End Sub

Private Sub RunCoSignerChoice(ByVal d As Document)
    Dim resp As VbMsgBoxResult
    Dim bmPairs(1 To 4, 1 To 2) As String
    bmPairs(1, 1) = "secondary_sig":       bmPairs(1, 2) = "secondary_sig_end"
    bmPairs(2, 1) = "secondary_sig_2":     bmPairs(2, 2) = "secondary_sig_2end"
    bmPairs(3, 1) = "secondary_sig_cover": bmPairs(3, 2) = "secondary_sig_cover_end"
    bmPairs(4, 1) = "sig_3":               bmPairs(4, 2) = "sig_3_end"

    Do
        resp = MsgBox( _
            "Does this proposal need a co-signer (second salesperson signature)?" & vbCrLf & vbCrLf & _
            "Yes = Keep co-signer (2 salesperson signatures)" & vbCrLf & _
            "No  = Remove the co-signer sections (1 salesperson signature).", _
            vbYesNo + vbQuestion + vbDefaultButton2, _
            "Keep Co-Signers?" _
        )
        If resp = vbYes Then
            DeleteBookmarksOnly d, bmPairs
            Exit Do
        ElseIf resp = vbNo Then
            RemoveCosignerSectionsSafe d, bmPairs
            Exit Do
        End If
    Loop
End Sub

' ---------- Dates ----------
Public Sub UpdateDateContentControls(ByVal d As Document)
    Dim cc As ContentControl
    For Each cc In d.ContentControls
        If cc.Type = wdContentControlDate Then
            Select Case LCase$(cc.Tag)
                Case "datecontrol":  cc.Range.Text = Format(Date, "dddd, mmmm d, yyyy")
                Case "datecontrol2": cc.Range.Text = Format(Date, "mm/dd/yy")
            End Select
        End If
    Next cc
End Sub

' ---------- Parent?Child CC mirroring ----------
Public Sub SyncParentChild(ByVal d As Document, ByVal cc As ContentControl, ByRef Cancel As Boolean)
    If InStr(1, cc.Tag, "parent", vbTextCompare) = 0 Then Exit Sub
    If Len(cc.title) = 0 Then Exit Sub

    Dim cleaned As String, parentText As String
    If cc.ShowingPlaceholderText Then
        cleaned = ""
    Else
        parentText = cc.Range.Text
        cleaned = Replace(Replace(parentText, vbCr, " "), vbLf, " ")
        cleaned = Trim$(cleaned)
    End If

    Dim col As ContentControls, childCC As ContentControl
    On Error Resume Next
    Set col = d.SelectContentControlsByTitle(cc.title)
    On Error GoTo 0
    If col Is Nothing Then Exit Sub

    For Each childCC In col
        If childCC.ID <> cc.ID Then childCC.Range.Text = cleaned
    Next childCC
End Sub

' ---------- Bookmark helpers ----------
Private Sub DeleteBookmarksOnly(ByVal d As Document, ByRef bmPairs() As String)
    Dim i As Long
    For i = LBound(bmPairs, 1) To UBound(bmPairs, 1)
        DeleteBookmarkIfExists d, bmPairs(i, 1)
        DeleteBookmarkIfExists d, bmPairs(i, 2)
    Next i
End Sub

Private Sub DeleteBookmarkIfExists(ByVal d As Document, ByVal bmName As String)
    If d.Bookmarks.Exists(bmName) Then d.Bookmarks(bmName).Delete
End Sub

Private Sub RemoveCosignerSectionsSafe(ByVal d As Document, ByRef bmPairs() As String)
    Dim i As Long, r As Range, issues As String
    Dim ranges() As Range, count As Long

    For i = LBound(bmPairs, 1) To UBound(bmPairs, 1)
        Set r = TryInclusiveRange(d, bmPairs(i, 1), bmPairs(i, 2), issues)
        If Not r Is Nothing Then
            count = count + 1
            ReDim Preserve ranges(1 To count)
            Set ranges(count) = r.Duplicate
        End If
    Next i

    If count = 0 Then
        If Len(issues) > 0 Then MsgBox "No co-signer regions removed:" & vbCrLf & issues, vbInformation
        Exit Sub
    End If

    Dim a As Long, b As Long
    Dim tmp As Range
    For a = 1 To count - 1
        For b = a + 1 To count
            If ranges(a).Start < ranges(b).Start Then
                Set tmp = ranges(a): Set ranges(a) = ranges(b): Set ranges(b) = tmp
            End If
        Next b
    Next a

    For a = 1 To count
        ranges(a).Delete
    Next a

    If Len(issues) > 0 Then
        MsgBox "Some pairs were skipped:" & vbCrLf & issues, vbInformation
    End If
End Sub

Private Function TryInclusiveRange(ByVal d As Document, _
                                   ByVal startName As String, ByVal endName As String, _
                                   ByRef issues As String) As Range
    On Error GoTo fail
    Dim bmStart As Bookmark, bmEnd As Bookmark
    Dim r As Range

    If Not d.Bookmarks.Exists(startName) Then
        issues = issues & "  Missing bookmark: " & startName & vbCrLf
        Exit Function
    End If
    If Not d.Bookmarks.Exists(endName) Then
        issues = issues & "  Missing bookmark: " & endName & vbCrLf
        Exit Function
    End If

    Set bmStart = d.Bookmarks(startName)
    Set bmEnd = d.Bookmarks(endName)

    If bmStart.Range.StoryType <> bmEnd.Range.StoryType Then
        issues = issues & "  Cross-story pair: " & startName & " ? " & endName & vbCrLf
        Exit Function
    End If

    If bmStart.Range.Start <= bmEnd.Range.End Then
        Set r = bmStart.Range.Duplicate
        r.End = bmEnd.Range.End
    Else
        Set r = bmEnd.Range.Duplicate
        r.End = bmStart.Range.End
    End If

    If r.Start >= r.End Then Exit Function

    Set TryInclusiveRange = r
    Exit Function

fail:
    issues = issues & "  Error for pair " & startName & " ? " & endName & _
                      ": " & Err.Number & " - " & Err.Description & vbCrLf
    On Error GoTo 0
End Function

' ---------- Copy detection on open ----------
Public Sub OnOpenCopyCheck(ByVal d As Document)
    If Not IsOurProposal(d) Then Exit Sub

    Dim fso As Object, f As Object
    Dim curPath As String, curCreated As Date
    Dim prevPath As String, prevCreatedStr As String, prevCreated As Date
    Dim isCopy As Boolean

    On Error GoTo CleanExit
    Set fso = CreateObject("Scripting.FileSystemObject")
    curPath = d.FullName
    Set f = fso.GetFile(curPath)
    curCreated = f.DateCreated

    prevPath = GetVar(d, "LastKnownPath")
    prevCreatedStr = GetVar(d, "LastKnownFsCreated")
    If IsDate(prevCreatedStr) Then prevCreated = CDate(prevCreatedStr)

    If Len(prevPath) > 0 Then
        If LCase$(curPath) <> LCase$(prevPath) Then
            isCopy = LikelyCopy(curCreated, prevCreated)
            If isCopy Then
                MintNewProposalGuid d
                SetAuthorToCurrent d
            End If
        End If
    End If

    SetVar d, "LastKnownPath", curPath
    SetVar d, "LastKnownFsCreated", Format$(curCreated, "yyyy-mm-dd hh:nn:ss")

CleanExit:
End Sub

Private Function LikelyCopy(ByVal curCreated As Date, ByVal prevCreated As Date) As Boolean
    If prevCreated = 0 Then
        LikelyCopy = (DateDiff("n", curCreated, Now) >= 0 And DateDiff("n", curCreated, Now) <= COPY_RECENCY_MIN)
        Exit Function
    End If
    LikelyCopy = (Abs(DateDiff("s", curCreated, prevCreated)) > 60) _
                 And (DateDiff("n", curCreated, Now) >= 0 And DateDiff("n", curCreated, Now) <= COPY_RECENCY_MIN)
End Function

' ---------- helpers ----------
Private Function IsOurProposal(ByVal d As Document) As Boolean
    On Error Resume Next
    IsOurProposal = (d.Variables("IsProposalDoc").Value = "1")
End Function

Private Function GetVar(ByVal d As Document, ByVal nm As String) As String
    On Error Resume Next
    GetVar = d.Variables(nm).Value
    Err.Clear
End Function

Private Sub SetVar(ByVal d As Document, ByVal nm As String, ByVal val As String)
    On Error Resume Next
    d.Variables(nm).Value = val
    Err.Clear
End Sub

Private Function CreateGuidString() As String
    On Error GoTo fallback
    CreateGuidString = Mid$(CreateObject("Scriptlet.TypeLib").guid, 2, 36)
    Exit Function
fallback:
    CreateGuidString = "GUID_" & Format(Now, "yyyymmddHHNNSS") & "_" & CStr(Int(Rnd() * 1000000#))
End Function


