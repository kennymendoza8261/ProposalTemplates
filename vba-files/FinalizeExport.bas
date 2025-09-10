Attribute VB_Name = "FinalizeExport"
Option Explicit

' === FinalizeExport.bas ===
' In-place finalize: .DOCM -> .DOCX (no macros), NO writes to Normal.dotm.
' - Creates DOCX in "<Name>'s Finalized Proposals"
' - Marks DOCX Final + sets read-only file attribute
' - Deletes the original DOCM after successful conversion

Private Const WD_FMT_DOCX As Long = 12

Public Sub FinalizeAndExportToDocx()
    Dim d As Document
    Dim srcPath As String, srcName As String
    Dim outFolder As String, outPath As String
    Dim owner As String
    Dim ok As VbMsgBoxResult

    Set d = ActiveDocument

    ' Guardrails: must be our proposal .docm
    If LCase$(Right$(d.name, 5)) <> ".docm" Then
        MsgBox "Open the proposal .docm and run Finalize from there (not from the .dotm).", vbExclamation, "Finalize"
        Exit Sub
    End If
    If Not IsOurProposalDoc(d) Then
        MsgBox "This document isn’t flagged as a proposal. Aborting Finalize.", vbExclamation, "Finalize"
        Exit Sub
    End If

    ok = MsgBox( _
        "You are about to FINALIZE this document." & vbCrLf & vbCrLf & _
        "• A .DOCX copy (no macros) will be created in your 'Finalized Proposals' folder." & vbCrLf & _
        "• The original .DOCM will then be permanently deleted." & vbCrLf & _
        "• To edit the finalized .DOCX later, manually clear:" & vbCrLf & _
        "    - File > Info > Protect Document > Mark as Final" & vbCrLf & _
        "    - Windows file 'Read-only' setting" & vbCrLf & vbCrLf & _
        "This action cannot be undone. Continue?", _
        vbExclamation + vbYesNo + vbDefaultButton2, "Finalize and Delete .DOCM?")
    If ok <> vbYes Then Exit Sub

    If Not EnsureSavable(d) Then Exit Sub

    srcPath = d.FullName
    srcName = d.name

    ' Ensure we can write next to the source and to the finals folder
    If Not CanCreateFile(d.path) Then
        MsgBox "Word cannot write to the document's folder:" & vbCrLf & d.path & vbCrLf & _
               "Check permissions/Controlled Folder Access or move the file.", vbExclamation, "Finalize"
        Exit Sub
    End If

    ' Resolve "<Name>'s Finalized Proposals"
    outFolder = ProposalConfig.FinalsFolderFromDoc(d)
    If Not EnsureFolder(outFolder) Or Not CanCreateFile(outFolder) Then
        owner = ProposalConfig.OwnerNameFromDocLocation(d): If Len(owner) = 0 Then owner = "User"
        outFolder = ProposalConfig.BuildPath(ProposalConfig.DefaultDocumentsPath(), owner & "'s Finalized Proposals_Fallback")
        If Not EnsureFolder(outFolder) Or Not CanCreateFile(outFolder) Then
            MsgBox "Can't write to the finals folder or fallback under Documents:" & vbCrLf & outFolder, vbExclamation, "Finalize"
            Exit Sub
        End If
    End If

    outPath = UniquePath(ProposalConfig.BuildPath(outFolder, SafeBaseForDocx(d) & "_FINAL.docx"))

    ' --- Convert to DOCX WITHOUT touching Normal --------------------------------
    ' Use SaveCopyAs for safety to keep current doc context, then open the copy and harden it.
    On Error GoTo convertFail
    Dim tmpDocx As String: tmpDocx = outPath
    d.SaveCopyAs fileName:=tmpDocx, FileFormat:=WD_FMT_DOCX   ' creates DOCX copy; original remains open
    ' Open the freshly created DOCX to mark it Final and set read-only attribute
    Dim x As Document
    Set x = Documents.Open(fileName:=tmpDocx, ReadOnly:=False, AddToRecentFiles:=True)

    On Error Resume Next
    x.Final = True            ' Mark as Final (document-level flag)
    x.Save
    x.Close SaveChanges:=True
    On Error GoTo 0

    ' Delete the original DOCM now (close it first if needed)
    If LCase$(ActiveDocument.FullName) = LCase$(srcPath) Then
        ' We are still in the source .docm; close it before deleting
        d.Close SaveChanges:=True
    Else
        ' The user might have switched windows; ensure source is closed
        CloseDocIfOpen srcPath
    End If

    If Not DeleteFileWithRetry(srcPath) Then
        MsgBox "DOCX created, but failed to delete the original DOCM (it may be locked):" & vbCrLf & srcPath, vbExclamation, "Finalize"
    Else
        MsgBox "Finalized successfully." & vbCrLf & _
               "DOCX: " & tmpDocx & vbCrLf & _
               "Deleted: " & srcName, vbInformation, "Finalize Complete"
    End If

    Exit Sub

convertFail:
    MsgBox "Finalize failed while converting to DOCX: " & Err.Number & " - " & Err.Description, vbExclamation, "Finalize"
End Sub

' --- Minimal proposal check ---
Private Function IsOurProposalDoc(ByVal d As Document) As Boolean
    On Error Resume Next
    IsOurProposalDoc = (d.Variables("IsProposalDoc").Value = "1")
End Function

' --- Helpers (local) ----------------------------------------------------------
Private Function EnsureSavable(ByVal d As Document) As Boolean
    If Len(d.path) = 0 Then
        If MsgBox("This document hasn't been saved yet. Save now?", vbQuestion + vbOKCancel, "Save Required") = vbOK Then
            d.Save
        Else
            EnsureSavable = False: Exit Function
        End If
    End If
    EnsureSavable = True
End Function

Private Function EnsureFolder(ByVal folderPath As String) As Boolean
    On Error GoTo fail
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    EnsureFolder = True: Exit Function
fail: EnsureFolder = False
End Function

Private Function CanCreateFile(ByVal folderPath As String) As Boolean
    On Error GoTo fail
    Dim f As Integer, probe As String
    f = FreeFile
    probe = folderPath & "\.__wprobe.tmp"
    Open probe For Output As #f: Close #f
    Kill probe
    CanCreateFile = True: Exit Function
fail: CanCreateFile = False
End Function

Private Function SafeBaseForDocx(ByVal d As Document) As String
    Dim b As String: b = GetBaseNameNoExt(d.name)
    If Len(b) > 120 Then b = Left$(b, 120)
    SafeBaseForDocx = b
End Function

Private Function GetBaseNameNoExt(ByVal fileName As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetBaseNameNoExt = fso.GetBaseName(fileName)
End Function

' Prefix-style uniqueness: "(n) Base.ext"
Private Function UniquePath(ByVal candidate As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(candidate) Then UniquePath = candidate: Exit Function

    Dim p As Long, folder As String, leaf As String, dotp As Long
    Dim base As String, ext As String, n As Long, tryPath As String

    p = InStrRev(candidate, "\")
    If p = 0 Then folder = "" Else folder = Left$(candidate, p)
    leaf = Mid$(candidate, p + 1)
    dotp = InStrRev(leaf, ".")
    If dotp > 0 Then base = Left$(leaf, dotp - 1): ext = Mid$(leaf, dotp) Else base = leaf: ext = ""

    n = 2
    Do
        tryPath = folder & "(" & n & ") " & base & ext
        If Not fso.FileExists(tryPath) Then UniquePath = tryPath: Exit Function
        n = n + 1: If n > 500 Then Exit Do
    Loop
    UniquePath = folder & base & "_" & Format(Now, "yyyymmdd_HHNNSS") & ext
End Function

Private Sub CloseDocIfOpen(ByVal fullPath As String)
    On Error Resume Next
    Dim doc As Document
    For Each doc In Application.Documents
        If LCase$(doc.FullName) = LCase$(fullPath) Then
            doc.Close SaveChanges:=True
            Exit For
        End If
    Next
    On Error GoTo 0
End Sub

Private Function DeleteFileWithRetry(ByVal fullPath As String, Optional ByVal tries As Long = 5, Optional ByVal ms As Long = 300) As Boolean
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim i As Long
    On Error Resume Next
    For i = 1 To tries
        Err.Clear
        If fso.FileExists(fullPath) Then
            fso.DeleteFile fullPath, True
            If Err.Number = 0 Then DeleteFileWithRetry = True: Exit Function
        Else
            DeleteFileWithRetry = True: Exit Function
        End If
        DoEvents
        SleepMs ms
    Next i
End Function

Private Sub SleepMs(ByVal ms As Long)
    Dim t As Single: t = Timer + (ms / 1000!)
    Do While Timer < t: DoEvents: Loop
End Sub


