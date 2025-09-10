VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Minimal app-events helper
Private gAppEv As TemplateAppEvents

Private Sub Document_New()
    On Error GoTo TopErr

    Dim d As Document
    Set d = ActiveDocument   ' The new .docm

    ' Ensure it's writable and not read-only
    d.Final = False
    d.ReadOnlyRecommended = False

    ' 1) Remove protection and update date content controls (content-only)
    SafeCall_UnrestrictAll d
    SafeCall_UpdateDateContentControls d

    ' 2) Force initial SaveAs (.docm in "<Name>'s Active Proposals")
    If Not SafeCall_ForceInitialSaveOutsideTemplateFolder(d) Then
        TB_Log "User cancelled or save failed in ForceInitialSaveOutsideTemplateFolder."
        GoTo CleanExit
    End If

    If Len(d.FullName) = 0 Then
        MsgBox "Auto-save failed — the new document isn't saved to disk. Bootstrap aborted.", vbExclamation, "Bootstrap"
        TB_Log "Abort: new doc not saved to disk."
        GoTo CleanExit
    End If

    If d.ReadOnly Then
        MsgBox "The newly created document is read-only. Check file permissions and retry.", vbExclamation, "Bootstrap"
        TB_Log "Abort: new doc is read-only: " & d.FullName
        GoTo CleanExit
    End If

    ' 3) Prevent style updates and NormalTemplate prompts
    On Error Resume Next
    d.UpdateStylesOnOpen = False
    Application.Options.SaveNormalPrompt = False
    On Error GoTo TopErr

    ' 4) Add identity markers and metadata (before guards that check them)
    SafeCall_EnsureProposalGuid d
    On Error Resume Next
    d.Variables("IsProposalDoc").Value = "1"
    On Error GoTo TopErr
    SafeCall_SetAuthorToCurrent d

    ' 5) Copy required modules from template to .docm (VBProject mutation)
    TB_Log "Copying modules into: " & d.FullName
    SafeCall_CopyProjectItemsToDocm_Safe d, Array("ProposalConfig", "Bootstrap", "DocEvents", "DocBehavior", "FinalizeExport")

    ' 6) Update proposal markers
    SafeCall_UpdateCopyMarkers d

    ' Save with retry (avoids OneDrive/AV locks)
    SaveWithRetry d

    ' 7) Sign LAST (no prompts); save again after signing
    CertSign.SignCurrentDocm_NoPrompts d
    SaveWithRetry d

    ' 8) Small delay, then trigger doc-side bootstrap (runs in child project)
    SleepMs 250
    SafeCall_CallDocMacro d, "Bootstrap.AutoOpen"

CleanExit:
    EnsureAppEvents
    Exit Sub

TopErr:
    TB_Log "Document_New ERROR " & Err.Number & " - " & Err.Description
    MsgBox "Template bootstrap encountered an unexpected error: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Inspect the log at: " & vbCrLf & Environ$("TEMP") & "\TemplateBootstrap.log", vbExclamation, "Template Bootstrap Error"
    Resume CleanExit
End Sub

Private Sub Document_Open()
    On Error Resume Next
    EnsureAppEvents
End Sub

Private Sub EnsureAppEvents()
    On Error Resume Next
    If gAppEv Is Nothing Then
        Set gAppEv = New TemplateAppEvents
        gAppEv.Hook Application
    End If
End Sub

' ---- retry helpers -----------------------------------------------------------

Private Sub SaveWithRetry(ByVal d As Document, Optional ByVal tries As Long = 5, Optional ByVal ms As Long = 250)
    Dim i As Long
    For i = 1 To tries
        On Error Resume Next
        d.Save
        If Err.Number = 0 Then Exit Sub
        Err.Clear
        DoEvents
        SleepMs ms
    Next i
End Sub

Private Sub SleepMs(ByVal ms As Long)
    Dim t As Single: t = Timer + (ms / 1000!)
    Do While Timer < t: DoEvents: Loop
End Sub

' ============================================
' Runtime-safe wrappers to TemplateBootstrap
' ============================================

Private Sub TB_Log(ByVal msg As String)
    On Error Resume Next
    Application.Run "TemplateBootstrap.TB_Log", msg
    If Err.Number <> 0 Then
        ' fallback to write to %TEMP%
        Dim tf As Integer, logPath As String
        logPath = Environ$("TEMP") & "\TemplateBootstrap.log"
        tf = FreeFile
        Open logPath For Append As #tf
        Print #tf, Format$(Now, "yyyy-mm-dd HH:nn:ss") & " - " & msg
        Close #tf
        Err.Clear
    End If
End Sub

Private Sub SafeCall_UnrestrictAll(ByVal d As Document)
    On Error Resume Next
    Application.Run "TemplateBootstrap.UnrestrictAll", d
End Sub

Private Sub SafeCall_UpdateDateContentControls(ByVal d As Document)
    On Error Resume Next
    Application.Run "TemplateBootstrap.UpdateDateContentControls", d
End Sub

Private Sub SafeCall_SetAuthorToCurrent(ByVal d As Document)
    On Error Resume Next
    Application.Run "TemplateBootstrap.SetAuthorToCurrent", d
End Sub

Private Sub SafeCall_EnsureProposalGuid(ByVal d As Document)
    On Error Resume Next
    Application.Run "TemplateBootstrap.EnsureProposalGuid", d
End Sub

Private Sub SafeCall_CopyProjectItemsToDocm_Safe(ByVal TargetDoc As Document, ByVal ModuleNames As Variant)
    On Error Resume Next
    Application.Run "TemplateBootstrap.CopyProjectItemsToDocm_Safe", TargetDoc, ModuleNames
End Sub

Private Sub SafeCall_UpdateCopyMarkers(ByVal d As Document)
    On Error Resume Next
    Application.Run "TemplateBootstrap.UpdateCopyMarkers", d
End Sub

Private Sub SafeCall_CallDocMacro(ByVal d As Document, ByVal macroName As String)
    On Error Resume Next
    Application.Run "TemplateBootstrap.CallDocMacro", d, macroName
End Sub

Private Function SafeCall_ForceInitialSaveOutsideTemplateFolder(ByVal d As Document) As Boolean
    On Error Resume Next
    Dim result As Variant
    result = Application.Run("TemplateBootstrap.ForceInitialSaveOutsideTemplateFolder", d)
    SafeCall_ForceInitialSaveOutsideTemplateFolder = CBool(result)
End Function


