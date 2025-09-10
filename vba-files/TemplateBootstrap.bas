Attribute VB_Name = "TemplateBootstrap"
Option Explicit

' ========================================
' Public methods used by ThisDocument.SafeCall_* wrappers
' ========================================

Public Sub UnrestrictAll(ByVal d As Document)
    On Error Resume Next
    If d.ProtectionType <> wdNoProtection Then d.Unprotect
    d.ReadOnlyRecommended = False
    d.Final = False
End Sub

Public Sub UpdateDateContentControls(ByVal d As Document)
    Dim cc As ContentControl
    For Each cc In d.ContentControls
        If cc.Type = wdContentControlDate Then
            Select Case LCase(cc.Tag)
                Case "datecontrol":  cc.Range.Text = Format(Date, "dddd, mmmm d, yyyy")
                Case "datecontrol2": cc.Range.Text = Format(Date, "mm/dd/yy")
            End Select
        End If
    Next cc
End Sub

Public Sub SetAuthorToCurrent(ByVal d As Document)
    On Error Resume Next
    d.BuiltInDocumentProperties("Author").Value = Application.UserName
    Err.Clear
End Sub

Public Function ForceInitialSaveOutsideTemplateFolder(ByVal d As Document) As Boolean
    Dim fso As Object, targetPath As String
    Dim outFolder As String, baseName As String

    On Error GoTo ErrHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    baseName = CleanFileName(GetDateTemplateBaseName(d))

    outFolder = ProposalConfig.ActiveFolderFromTemplate(d.AttachedTemplate)
    If Len(Trim$(outFolder)) = 0 Then
        outFolder = ProposalConfig.BuildPath(ProposalConfig.DefaultDocumentsPath(), _
            ProposalConfig.ActiveFolderName(ProposalConfig.OwnerNameFromTemplateOrConst(d.AttachedTemplate)))
    End If

    If Not fso.FolderExists(outFolder) Then fso.CreateFolder outFolder

    targetPath = UniqueDocmPath(outFolder, baseName)
    d.SaveAs2 fileName:=targetPath, _
              FileFormat:=wdFormatXMLDocumentMacroEnabled, _
              AddToRecentFiles:=True

    ForceInitialSaveOutsideTemplateFolder = True
    Exit Function

ErrHandler:
    ForceInitialSaveOutsideTemplateFolder = False
    TB_Log "ForceInitialSaveOutsideTemplateFolder error: " & Err.Number & " - " & Err.Description
End Function

Public Sub EnsureProposalGuid(ByVal d As Document)
    Const PROP_NAME As String = "ProposalGuid"
    Dim props As Object
    Set props = d.CustomDocumentProperties
    On Error Resume Next
    Dim cur: cur = props(PROP_NAME).Value
    If Err.Number = 0 And Len(CStr(cur)) > 0 Then Exit Sub
    Err.Clear
    props.Add name:=PROP_NAME, LinkToContent:=False, Type:=4, Value:=CreateGuidString()
End Sub

Public Sub CopyProjectItemsToDocm_Safe(ByVal TargetDoc As Document, ByVal ModuleNames As Variant)
    On Error GoTo fail
    Dim src As String, dst As String
    Dim i As Long, nm As String

    src = ThisDocument.FullName
    dst = TargetDoc.FullName

    If Len(dst) = 0 Then
        MsgBox "Target document must be saved before copying modules.", vbExclamation, "Copy Modules"
        Exit Sub
    End If

    For i = LBound(ModuleNames) To UBound(ModuleNames)
        nm = CStr(ModuleNames(i))
        On Error Resume Next
        Application.OrganizerDelete Source:=dst, name:=nm, Object:=wdOrganizerObjectProjectItems
        Application.OrganizerCopy Source:=src, Destination:=dst, name:=nm, Object:=wdOrganizerObjectProjectItems
        If Err.Number <> 0 Then
            TB_Log "Failed copying module: " & nm & " - " & Err.Description
            Err.Clear
        End If
    Next i

    Exit Sub

fail:
    TB_Log "CopyProjectItemsToDocm_Safe error: " & Err.Number & " - " & Err.Description
    MsgBox "Failed copying project items: " & Err.Description, vbExclamation, "Copy Modules"
End Sub

Public Sub UpdateCopyMarkers(ByVal d As Document)
    On Error Resume Next
    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(d.FullName)
    d.Variables("LastKnownPath").Value = d.FullName
    d.Variables("LastKnownFsCreated").Value = Format$(f.DateCreated, "yyyy-mm-dd hh:nn:ss")
    Err.Clear

    ' Ensure new core modules are present in the child .docm as well
    CopyProjectItemsToDocm_Safe d, Array("ProposalEngine", "ExcelHook", "RibbonCallbacks", "ProposalLayoutForm")
End Sub

Public Sub CallDocMacro(ByVal d As Document, ByVal macroName As String, ParamArray args() As Variant)
    Dim q As String, success As Boolean, hasArg As Boolean
    q = "'" & d.FullName & "'!" & macroName

    On Error Resume Next
    hasArg = (UBound(args) >= 0)
    If Err.Number <> 0 Then hasArg = False: Err.Clear

    If hasArg Then
        Application.Run q, args(0)
    Else
        Application.Run q
    End If

    success = (Err.Number = 0)
    If Not success Then TB_Log "CallDocMacro failed: " & macroName
    Err.Clear
End Sub

' ========================================
' Internal Helpers (can stay private)
' ========================================

Private Function CreateGuidString() As String
    On Error GoTo fallback
    CreateGuidString = Mid$(CreateObject("Scriptlet.TypeLib").guid, 2, 36)
    Exit Function
fallback:
    CreateGuidString = "GUID_" & Format(Now, "yyyymmddHHNNSS") & "_" & CStr(Int(Rnd() * 1000000#))
End Function

Private Function GetDateTemplateBaseName(ByVal d As Document) As String
    Dim fso As Object, tmplName As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    tmplName = fso.GetBaseName(CStr(d.AttachedTemplate.FullName))
    If Len(tmplName) = 0 Then tmplName = "NewDocument"
    ' Naming per spec: MM.DD.YY.<TemplateName>
    GetDateTemplateBaseName = Format(Date, "mm.dd.yy") & "." & tmplName
End Function

Private Function CleanFileName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace$(s, CStr(ch), "_")
    Next
    CleanFileName = Trim$(s)
End Function

Private Function UniqueDocmPath(ByVal folderPath As String, ByVal baseName As String) As String
    Dim fso As Object, candidate As String, n As Long
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    candidate = folderPath & "\" & baseName & ".docm"

    If Not fso.FileExists(candidate) Then
        UniqueDocmPath = candidate
        Exit Function
    End If

    n = 2
    Do
        candidate = folderPath & "\" & "(" & n & ")_" & baseName & ".docm"
        n = n + 1
        If n > 500 Then Exit Do
    Loop While fso.FileExists(candidate)

    UniqueDocmPath = candidate
End Function

Public Sub TB_Log(ByVal msg As String)
    On Error Resume Next
    Dim tf As Integer, logPath As String
    logPath = Environ$("TEMP") & "\TemplateBootstrap.log"
    tf = FreeFile
    Open logPath For Append As #tf
    Print #tf, Format$(Now, "yyyy-mm-dd HH:nn:ss") & " - " & msg
    Close #tf
    Err.Clear
End Sub


