Attribute VB_Name = "ExcelHook"
Option Explicit

' Late-bound Excel integration for Method 1 (optional convenience).
' Excel can skip this and call ProposalEngine.ApplyDataFromDictionary directly.

Public Sub PopulateFromWorkbookRow(ByVal workbookPath As String, _
                                   ByVal sheetName As String, _
                                   ByVal rowIndex As Long, _
                                   ByVal layout As Long, _
                                   Optional ByVal totalReps As Long = 1)
    On Error GoTo fail

    Dim xl As Object, wb As Object, ws As Object
    Dim lastCol As Long, c As Long
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    If Len(Dir$(workbookPath)) = 0 Then
        MsgBox "Workbook not found: " & workbookPath, vbExclamation, "Excel Hook"
        Exit Sub
    End If

    Set xl = GetObject(, "Excel.Application")
    If xl Is Nothing Then Set xl = CreateObject("Excel.Application")
    Set wb = xl.Workbooks.Open(workbookPath, ReadOnly:=True)
    If Len(sheetName) = 0 Then
        Set ws = wb.Sheets(1)
    Else
        Set ws = wb.Sheets(sheetName)
    End If

    lastCol = ws.Cells(1, ws.Columns.Count).End(-4159).Column ' xlToLeft = -4159
    For c = 1 To lastCol
        Dim rawKey As String, key As String, val As String
        rawKey = Trim$(CStr(ws.Cells(1, c).Value))
        If Len(rawKey) > 0 Then
            key = NormalizeHeaderToTag(rawKey)
            val = Trim$(CStr(ws.Cells(rowIndex, c).Value))
            ' Skip QuoteNumber entirely
            If LCase$(key) <> "quotenumber" Then
                ' Optionally skip empty CHC rep fields
                If Not (IsCHCRepField(key) And Len(val) = 0) Then
                    dict(key) = val
                End If
            End If
        End If
    Next c

    ProposalEngine.ApplyDataFromDictionary dict, layout, totalReps

Clean:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not xl Is Nothing Then xl.Quit
    Set ws = Nothing: Set wb = Nothing: Set xl = Nothing
    Exit Sub

fail:
    MsgBox "PopulateFromWorkbookRow failed: " & Err.Number & " - " & Err.Description, vbExclamation, "Excel Hook"
    Resume Clean
End Sub

Private Function NormalizeHeaderToTag(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, " ", "")
    t = Replace(t, ".", "")
    t = Replace(t, "/", "")
    t = Replace(t, "-", "")
    t = Replace(t, "(", "")
    t = Replace(t, ")", "")
    ' Common words case mapping
    t = Replace(t, "ProjectName", "ProjectName")
    t = Replace(t, "QuoteNumber", "QuoteNumber")
    t = Replace(t, "CompanyName", "CompanyName")
    t = Replace(t, "CustomerName", "CustomerName")
    t = Replace(t, "CustomerJobTitle", "CustomerJobTitle")
    t = Replace(t, "CurrentLocation", "CurrentLocation")
    t = Replace(t, "NewLocation", "NewLocation")
    t = Replace(t, "Date", "Date")
    ' CHC reps: "CHCRep2Name" style
    t = Replace(t, "CHCRep", "CHCRep") ' ensure consistent
    ' Accept headers like "CHC Rep 2 Name"
    t = Replace(t, "CHCRep", "CHCRep")
    NormalizeHeaderToTag = t
End Function

Private Function IsCHCRepField(ByVal key As String) As Boolean
    Dim k As String: k = LCase$(key)
    IsCHCRepField = (Left$(k, 6) = "chcrep")
End Function
