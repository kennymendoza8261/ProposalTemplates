VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

Public Sub Hook(ByVal a As Word.Application)
    Set App = a
End Sub

Private Sub App_WindowActivate(ByVal doc As Document, ByVal Wn As Window)
    On Error Resume Next
    ' reserved for UI invalidation or light logging
End Sub

Private Sub App_DocumentOpen(ByVal doc As Document)
    On Error Resume Next
    ' reserved
End Sub

