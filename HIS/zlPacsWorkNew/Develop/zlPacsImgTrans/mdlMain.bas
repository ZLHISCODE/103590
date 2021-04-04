Attribute VB_Name = "mdlMain"
Option Explicit

Private Sub Main()
    Dim strCmdPath As String
    
On Error GoTo errHandle
    strCmdPath = Command

    strCmdPath = "C:\APPSOFT\TmpImage\TransCmd\"

    If Trim(strCmdPath) = "" Then
        MsgBox "��ָ��������������Ŀ¼."
        Exit Sub
    End If

    If DirExists(strCmdPath) = False Then
        MsgBox "Ŀ¼ [" & strCmdPath & "] ��Ч."
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
