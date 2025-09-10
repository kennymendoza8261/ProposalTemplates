Attribute VB_Name = "RibbonCallbacks"
Option Explicit

Private g_ribbon As IRibbonUI

Public Sub OnRibbonLoad(ByVal r As IRibbonUI)
    Set g_ribbon = r
End Sub

Public Sub OnGenerate_Click(ByVal control As IRibbonControl)
    On Error Resume Next
    ProposalEngine.GenerateProposal
End Sub

Public Sub OnFinalize_Click(ByVal control As IRibbonControl)
    On Error Resume Next
    ProposalEngine.FinalizeToDocx
End Sub

