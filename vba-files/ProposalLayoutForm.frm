VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProposalLayoutForm 
   Caption         =   "Proposal Layout"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   OleObjectBlob   =   "ProposalLayoutForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProposalLayoutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedLayout As Long
Public TotalReps As Long

Private Sub UserForm_Initialize()
    SelectedLayout = 1
    TotalReps = 1

    Me.Caption = "Generate Proposal"

    ' Create controls programmatically to keep .frm simple
    Dim lbl As MSForms.Label
    Dim opt1 As MSForms.OptionButton, opt2 As MSForms.OptionButton, opt3 As MSForms.OptionButton
    Dim lblReps As MSForms.Label, txt As MSForms.TextBox
    Dim okBtn As MSForms.CommandButton, cancelBtn As MSForms.CommandButton

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblTitle")
    lbl.Caption = "Choose Layout"
    lbl.Left = 12: lbl.Top = 12: lbl.Width = 200

    Set opt1 = Me.Controls.Add("Forms.OptionButton.1", "optLayout1")
    opt1.Caption = "Layout 1: Cover + Letter + 2 Standard"
    opt1.Left = 24: opt1.Top = 36: opt1.Width = 420: opt1.Value = True

    Set opt2 = Me.Controls.Add("Forms.OptionButton.1", "optLayout2")
    opt2.Caption = "Layout 2: Letter + 2 Standard"
    opt2.Left = 24: opt2.Top = 66: opt2.Width = 420

    Set opt3 = Me.Controls.Add("Forms.OptionButton.1", "optLayout3")
    opt3.Caption = "Layout 3: 2 Standard"
    opt3.Left = 24: opt3.Top = 96: opt3.Width = 420

    Set lblReps = Me.Controls.Add("Forms.Label.1", "lblReps")
    lblReps.Caption = "Total CHC Representatives (1 = just me)"
    lblReps.Left = 12: lblReps.Top = 150: lblReps.Width = 350

    Set txt = Me.Controls.Add("Forms.TextBox.1", "txtReps")
    txt.Left = 24: txt.Top = 174: txt.Width = 60: txt.Text = "1"

    Set okBtn = Me.Controls.Add("Forms.CommandButton.1", "cmdOK")
    okBtn.Caption = "OK"
    okBtn.Left = 360: okBtn.Top = 210: okBtn.Width = 72

    Set cancelBtn = Me.Controls.Add("Forms.CommandButton.1", "cmdCancel")
    cancelBtn.Caption = "Cancel"
    cancelBtn.Left = 444: cancelBtn.Top = 210: cancelBtn.Width = 72
End Sub

Private Sub cmdOK_Click()
    Dim n As Long
    n = Val(Me.txtReps.Text)
    If n < 1 Then n = 1

    If Me.optLayout1.Value Then
        SelectedLayout = 1
    ElseIf Me.optLayout2.Value Then
        SelectedLayout = 2
    Else
        SelectedLayout = 3
    End If

    TotalReps = n
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    SelectedLayout = 0
    Me.Hide
End Sub

Public Function ShowAndGet(ByRef layoutChoice As Long, ByRef totalReps As Long) As Boolean
    Me.Show vbModal
    If SelectedLayout = 0 Then Exit Function
    layoutChoice = SelectedLayout
    totalReps = TotalReps
    ShowAndGet = True
End Function
