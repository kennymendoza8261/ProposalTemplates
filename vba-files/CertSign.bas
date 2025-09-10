Attribute VB_Name = "CertSign"
Option Explicit

' === CONFIG (edit these 3–4 lines) ============================================
Private Const SIGNER_EXE As String = "C:\Tools\VbaSigner\VbaSign.exe"
' Preferred: use a cert already in CurrentUser\My. If empty, code will fall back to PFX import.
Private Const CERT_THUMBPRINT As String = ""  ' e.g., "A1B2C3..."; keep empty to auto-pick newest code-signing cert
' Optional fallback (used only if cert not found in store):
Private Const PFX_PATH As String = "C:\Secure\SigningCert.pfx"
' DO NOT hard-code the password. Set an environment variable named SIGN_PFX_PWD with the password.
Private Const PFX_PWD_ENV As String = "SIGN_PFX_PWD"

' === PUBLIC ENTRY =============================================================

' Signs the passed .docm silently.
Public Sub SignCurrentDocm_NoPrompts(ByVal d As Document)
    Dim docPath As String, thumb As String, ok As Boolean

    If d Is Nothing Then Exit Sub
    docPath = d.FullName
    If Len(docPath) = 0 Then Exit Sub         ' must be saved .docm

    ' 1) Resolve a code-signing certificate (store ? else PFX import ? store)
    thumb = ResolveOrImportCodeSigningThumbprint()
    If Len(thumb) = 0 Then
        CS_Log "No usable code-signing certificate in CurrentUser\My (and no PFX fallback)."
        Exit Sub
    End If

    ' 2) Shell signer by thumbprint (no secrets in code)
    ok = RunSigner_ByThumbprint(docPath, thumb)
    If Not ok Then
        CS_Log "Signing failed for: " & docPath
    End If
End Sub

' === CERT RESOLUTION / IMPORT =================================================

Private Function ResolveOrImportCodeSigningThumbprint() As String
    Dim thumb As String

    ' Prefer an explicit thumbprint if provided
    If Len(Trim$(CERT_THUMBPRINT)) > 0 Then
        thumb = FindThumbprintInMyStore(Trim$(CERT_THUMBPRINT), True) ' exact match
        If Len(thumb) > 0 Then
            ResolveOrImportCodeSigningThumbprint = thumb
            Exit Function
        End If
    End If

    ' Auto-pick newest code-signing cert with private key
    thumb = GetNewestCodeSigningThumbprint()
    If Len(thumb) > 0 Then
        ResolveOrImportCodeSigningThumbprint = thumb
        Exit Function
    End If

    ' Nothing in store – try PFX import silently (only if both path and env pwd present)
    If Len(Dir$(PFX_PATH)) > 0 Then
        Dim pwd As String
        pwd = Environ$(PFX_PWD_ENV)
        If Len(pwd) > 0 Then
            If PS_ImportPfx_Exportable(PFX_PATH, pwd) Then
                ' Retry find
                thumb = GetNewestCodeSigningThumbprint()
                If Len(thumb) > 0 Then
                    ResolveOrImportCodeSigningThumbprint = thumb
                    Exit Function
                End If
            End If
        End If
    End If
End Function

Private Function FindThumbprintInMyStore(ByVal thumb As String, ByVal exact As Boolean) As String
    Dim t As String
    t = NormalizeHex(thumb)
    Dim got As String: got = PS_QueryThumbprintExact(t, exact)
    If Len(got) > 0 Then FindThumbprintInMyStore = got
End Function

Private Function GetNewestCodeSigningThumbprint() As String
    GetNewestCodeSigningThumbprint = PS_GetNewestCodeSigningThumbprint()
End Function

Private Function NormalizeHex(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9A-Fa-f]" Then out = out & UCase$(ch)
    Next
    NormalizeHex = out
End Function

' === SIGNER CALL ==============================================================
Private Function RunSigner_ByThumbprint(ByVal docPath As String, ByVal sha1Thumb As String) As Boolean
    Dim cmd As String, rc As Long
    If Len(Dir$(SIGNER_EXE)) = 0 Then
        CS_Log "Signer not found: " & SIGNER_EXE
        Exit Function
    End If

    cmd = """" & SIGNER_EXE & """" & _
          " /file """ & docPath & """" & _
          " /sha1 " & sha1Thumb & _
          " /store My /user"

    rc = CS_ShellRunWait(cmd, 0)
    RunSigner_ByThumbprint = (rc = 0)
End Function

' === POWERSHELL HELPERS (silent, no UI) ======================================

Private Function PS_ImportPfx_Exportable(ByVal pfxPath As String, ByVal pfxPwd As String) As Boolean
    Dim ps As String, rc As Long, tmp As String
    tmp = Environ$("TEMP") & "\__ps_import_pfx.ps1"
    ps = ""
    ps = ps & "$ErrorActionPreference='Stop';" & vbCrLf
    ps = ps & "$p = '" & CS_PSQ(pfxPath) & "';" & vbCrLf
    ps = ps & "$sec = ConvertTo-SecureString '" & CS_PSQ(pfxPwd) & "' -AsPlainText -Force;" & vbCrLf
    ps = ps & "Import-PfxCertificate -FilePath $p -Password $sec -Exportable -CertStoreLocation Cert:\CurrentUser\My | Out-Null" & vbCrLf

    If Not CS_WriteText(tmp, ps) Then Exit Function
    rc = CS_RunPSFile(tmp)
    PS_ImportPfx_Exportable = (rc = 0)
End Function

Private Function PS_GetNewestCodeSigningThumbprint() As String
    Dim ps As String, rc As Long, tmp As String, out As String
    tmp = Environ$("TEMP") & "\__ps_get_thumb.txt"

    ps = ""
    ps = ps & "$cs='1.3.6.1.5.5.7.3.3';" & vbCrLf
    ps = ps & "$c=(Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.HasPrivateKey -and $_.EnhancedKeyUsageList.ObjectId -contains $cs }) |" & vbCrLf
    ps = ps & " Sort-Object NotBefore -Descending | Select-Object -First 1;" & vbCrLf
    ps = ps & "if ($c) { $c.Thumbprint } else { '' }" & vbCrLf
    CS_PS_WriteAndRun ps, tmp, rc, out
    If rc = 0 Then PS_GetNewestCodeSigningThumbprint = Trim$(out)
End Function

Private Function PS_QueryThumbprintExact(ByVal thumb As String, ByVal exact As Boolean) As String
    Dim ps As String, rc As Long, tmp As String, out As String
    tmp = Environ$("TEMP") & "\__ps_has_thumb.txt"
    If exact Then
        ps = "$t='" & CS_PSQ(thumb) & "';$x=Get-ChildItem Cert:\CurrentUser\My | ? { ($_.Thumbprint -replace ' ','').ToUpper() -eq $t }; if ($x){$x[0].Thumbprint}else{''}"
    Else
        ps = "$t='" & CS_PSQ(thumb) & "';$x=Get-ChildItem Cert:\CurrentUser\My | ? { ($_.Thumbprint -replace ' ','').ToUpper() -like ('*'+$t+'*') }; if ($x){$x[0].Thumbprint}else{''}"
    End If
    CS_PS_WriteAndRun ps, tmp, rc, out
    If rc = 0 Then PS_QueryThumbprintExact = Trim$(out)
End Function

' ==== small PS runner utilities ==============================================
Private Sub CS_PS_WriteAndRun(ByVal ps As String, ByVal outPath As String, ByRef rc As Long, ByRef stdout As String)
    Dim sfile As String: sfile = Environ$("TEMP") & "\__run_tmp.ps1"
    If Not CS_WriteText(sfile, _
        "$ErrorActionPreference='Stop';" & vbCrLf & _
        ps & " | Out-File -FilePath '" & CS_PSQ(outPath) & "' -Encoding ascii -Force") Then
        rc = 1: Exit Sub
    End If
    rc = CS_RunPSFile(sfile)
    If rc = 0 Then stdout = CS_ReadAll(outPath)
End Sub

Private Function CS_RunPSFile(ByVal ps1 As String) As Long
    Dim sh As Object, cmd As String
    Set sh = CreateObject("WScript.Shell")
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """"
    CS_RunPSFile = sh.Run(cmd, 0, True)
End Function

Private Function CS_PSQ(ByVal s As String) As String
    CS_PSQ = Replace(s, "'", "''")
End Function

' === GENERIC FS/SHELL/LOG =====================================================
Private Function CS_ShellRunWait(ByVal cmd As String, ByVal winStyle As VbAppWinStyle) As Long
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    CS_ShellRunWait = sh.Run(cmd, winStyle, True)
End Function

Private Function CS_WriteText(ByVal path As String, ByVal content As String) As Boolean
    Dim ff As Integer: ff = FreeFile
    On Error GoTo Oops
    Open path For Output As #ff
    Print #ff, content
    Close #ff
    CS_WriteText = True
    Exit Function
Oops:
    On Error Resume Next
    Close #ff
End Function

Private Function CS_ReadAll(ByVal path As String) As String
    Dim ff As Integer: ff = FreeFile
    On Error GoTo Oops
    Open path For Input As #ff
    CS_ReadAll = Input(LOF(ff), ff)
    Close #ff
    Exit Function
Oops:
    On Error Resume Next
    Close #ff
End Function

Private Sub CS_Log(ByVal msg As String)
    On Error Resume Next
    Application.Run "TemplateBootstrap.TB_Log", msg
    If Err.Number <> 0 Then
        Dim tf As Integer, logPath As String
        logPath = Environ$("TEMP") & "\TemplateBootstrap.log"
        tf = FreeFile
        Open logPath For Append As #tf
        Print #tf, Format$(Now, "yyyy-mm-dd HH:nn:ss") & " - " & msg
        Close #tf
        Err.Clear
    End If
End Sub


