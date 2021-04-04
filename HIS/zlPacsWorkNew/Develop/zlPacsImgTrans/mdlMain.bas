Attribute VB_Name = "mdlMain"
Option Explicit

Private Sub Main()
    Dim strCmdPath As String
    
On Error GoTo errHandle
    strCmdPath = Command

    strCmdPath = "C:\APPSOFT\TmpImage\TransCmd\"

    If Trim(strCmdPath) = "" Then
        MsgBox "请指定传输命令所在目录."
        Exit Sub
    End If

    If DirExists(strCmdPath) = False Then
        MsgBox "目录 [" & strCmdPath & "] 无效."
        Exit Sub
    End If
    
    Call StartServer(strCmdPath)
Exit Sub
errHandle:
    MsgBox Err.Description
End Sub


Private Sub StartServer(ByVal strCmdPath As String)
    Dim objMain As New frmMain
    
On Error GoTo errHandle
    Call objMain.Start(strCmdPath)
Exit Sub
errHandle:
    MsgBox Err.Description
End Sub
