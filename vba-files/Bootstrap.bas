Attribute VB_Name = "Bootstrap"
Option Explicit

Public gDoc As DocEvents  ' per-document event sink

Public Function VerifiedCaller(Optional ByVal CallerGuid As String = "") As Boolean
    ' Document-side call guard: allow only actual proposal docs (not the template)
    If LCase$(Right$(ThisDocument.FullName, 5)) = ".dotm" Then Exit Function
    On Error Resume Next
    If ThisDocument.Variables("IsProposalDoc").Value <> "1" Then Exit Function
    On Error GoTo 0

    Dim myGuid As String
    myGuid = DocBehavior.GetProposalGuid(ThisDocument)
    If Len(myGuid) = 0 Then Exit Function

    If Len(CallerGuid) = 0 Then
        VerifiedCaller = True
    Else
        VerifiedCaller = (StrComp(myGuid, CallerGuid, vbTextCompare) = 0)
    End If
End Function

Public Sub BindEventsAndInitHidden(Optional ByVal CallerGuid As String = "")
    If Not VerifiedCaller(CallerGuid) Then Exit Sub
    Set gDoc = New DocEvents
    Set gDoc.d = ThisDocument
End Sub

Public Sub AutoOpen()
    BindEventsAndInitHidden vbNullString
    DocBehavior.OnOpenCopyCheck ThisDocument
    DocBehavior.MaybeRunCosigner ThisDocument
    DocBehavior.UpdateDateContentControls ThisDocument
    ' New: prompt for layout + reps, then build Parts A/B/C
    ProposalEngine.MaybePromptOnOpen ThisDocument
End Sub

