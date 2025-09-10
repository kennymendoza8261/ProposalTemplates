Attribute VB_Name = "ProposalConfig"
Option Explicit
Public Const OWNER_DISPLAY_NAME As String = ""  ' e.g. "Ken"; "" = auto-detect from folder

' === Public APIs used by other modules ===
Public Function ActiveFolderFromTemplate(ByVal tmpl As Template) As String
    Dim root As String: root = RootFromTemplate(tmpl)
    ActiveFolderFromTemplate = BuildPath(root, ActiveFolderName(OwnerNameFromTemplateOrConst(tmpl)))
End Function

Public Function FinalsFolderFromDoc(ByVal d As Document) As String
    Dim root As String, owner As String
    root = ParentFolderPath(d.path)
    owner = OwnerNameFromDocLocation(d)               ' prefer doc location
    If Len(owner) = 0 Then owner = OwnerNameFallback  ' hard fallback
    FinalsFolderFromDoc = BuildPath(root, FinalsFolderName(owner))
End Function

Public Function TemplatesFolderFromTemplate(ByVal tmpl As Template) As String
    Dim root As String: root = RootFromTemplate(tmpl)
    TemplatesFolderFromTemplate = BuildPath(root, TemplatesFolderName(OwnerNameFromTemplateOrConst(tmpl)))
End Function

' ---------- Name builders ----------
Public Function ActiveFolderName(ByVal owner As String) As String
    ActiveFolderName = owner & "'s Active Proposals"
End Function

Public Function FinalsFolderName(ByVal owner As String) As String
    FinalsFolderName = owner & "'s Finalized Proposals"
End Function

Public Function TemplatesFolderName(ByVal owner As String) As String
    TemplatesFolderName = owner & "'s Proposal Templates"
End Function

' ---------- Owner name resolution ----------
Public Function OwnerNameFromTemplateOrConst(ByVal tmpl As Template) As String
    Dim nm As String
    nm = Trim$(OWNER_DISPLAY_NAME)
    If Len(nm) > 0 Then OwnerNameFromTemplateOrConst = nm: Exit Function

    nm = DetectOwnerFromFolderName(LeafFolderName(ParentFolderPath(tmpl.FullName)), "'s Proposal Templates")
    If Len(nm) > 0 Then OwnerNameFromTemplateOrConst = nm: Exit Function

    OwnerNameFromTemplateOrConst = OwnerNameFallback
End Function

Public Function OwnerNameFromDocLocation(ByVal d As Document) As String
    Dim folder As String: folder = LeafFolderName(d.path)
    Dim nm As String
    nm = DetectOwnerFromFolderName(folder, "'s Active Proposals")
    If Len(nm) = 0 Then
        nm = DetectOwnerFromFolderName(folder, "'s Finalized Proposals")
    End If
    OwnerNameFromDocLocation = nm
End Function

Private Function OwnerNameFallback() As String
    On Error Resume Next
    Dim nm As String
    nm = Trim$(OWNER_DISPLAY_NAME)
    If Len(nm) = 0 Then nm = Trim$(Application.UserName)
    If Len(nm) = 0 Then nm = Trim$(Environ$("USERNAME"))
    If Len(nm) = 0 Then nm = "User"
    OwnerNameFallback = nm
End Function

Private Function DetectOwnerFromFolderName(ByVal folderLeaf As String, ByVal suffix As String) As String
    Dim p As Long
    p = InStr(1, folderLeaf, suffix, vbTextCompare)
    If p > 0 Then
        DetectOwnerFromFolderName = Trim$(Left$(folderLeaf, p - 1))
    End If
End Function

' ---------- Roots & paths ----------
Public Function RootFromTemplate(ByVal tmpl As Template) As String
    Dim pf As String
    pf = ParentFolderPath(ParentFolderPath(tmpl.FullName)) ' parent of the template folder
    If Len(pf) = 0 Then
        pf = DefaultDocumentsPath()
    End If
    RootFromTemplate = pf
End Function

Public Function ParentFolderPath(ByVal path As String) As String
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    ParentFolderPath = fso.GetParentFolderName(path)
End Function

Public Function LeafFolderName(ByVal folderPath As String) As String
    Dim i As Long
    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    i = InStrRev(folderPath, "\")
    If i > 0 Then
        LeafFolderName = Mid$(folderPath, i + 1)
    Else
        LeafFolderName = folderPath
    End If
End Function

Public Function BuildPath(ByVal folderPath As String, ByVal leaf As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    BuildPath = fso.BuildPath(folderPath, leaf)
End Function

Public Function DefaultDocumentsPath() As String
    Dim p As String
    On Error Resume Next
    p = Application.Options.DefaultFilePath(wdDocumentsPath)
    If Len(p) = 0 Then p = Environ$("USERPROFILE") & "\Documents"
    If Len(p) = 0 Then p = "C:\"
    DefaultDocumentsPath = p
End Function

