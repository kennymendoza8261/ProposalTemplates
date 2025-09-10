VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents d As Word.Document
Attribute d.VB_VarHelpID = -1

Private Sub Class_Initialize()
    ' consumer sets d = ThisDocument (child) from Bootstrap.BindEventsAndInitHidden
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set d = Nothing
End Sub

' --- Save pipeline hook ---
Private Sub d_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    DocBehavior.OnBeforeSave d, SaveAsUI, Cancel
    On Error GoTo 0
End Sub

' --- Content-control events ---
Private Sub d_ContentControlOnExit(ByVal cc As ContentControl, Cancel As Boolean)
    On Error Resume Next
    DocBehavior.SyncParentChild d, cc, Cancel
    On Error GoTo 0
End Sub

Private Sub d_ContentControlOnEnter(ByVal cc As ContentControl)
    On Error Resume Next
    ' placeholder
    On Error GoTo 0
End Sub

Private Sub d_Close()
    On Error Resume Next
    ' cleanup if needed
    On Error GoTo 0
End Sub

