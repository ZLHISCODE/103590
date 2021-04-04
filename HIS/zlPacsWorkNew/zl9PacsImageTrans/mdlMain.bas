Attribute VB_Name = "mdlMain"
Option Explicit

Public Const C_BIGFILE_TAG As String = "[BF]"


Public gstrCmdPath As String
Public gblnIsFailed As Boolean
Public gblnSingle As Boolean    '启动独立传输进程
Public gblnWorking As Boolean

Private Sub Main()
    Dim strCmdPath As String
    
On Error GoTo errHandle

    strCmdPath = Command
    gblnIsFailed = False
    gblnSingle = False
    
    If Len(strCmdPath) >= 4 Then
        If Mid(strCmdPath, 1, 4) = "[BF]" Then
            gblnSingle = True
        End If
    End If
    
    If gblnSingle Then
        '启动单独的进程进行文件传输
        strCmdPath = Split(strCmdPath & "-", "-")(1)
        
        Call OpenTrayIcon
        
        gstrCmdPath = Split(strCmdPath, ",")(0) '[BF]-c:\appsoft\tmpimage\transcmd\,1.2.3.840.11276833.32.3559
        Call StartServer(strCmdPath, True)
    Else
        If DirExists(strCmdPath) = False Then strCmdPath = "C:\APPSOFT\Apply\TmpImage\TransCmd\"
        If DirExists(strCmdPath) = False Then strCmdPath = "D:\APPSOFT\Apply\TmpImage\TransCmd\"
        If DirExists(strCmdPath) = False Then strCmdPath = "E:\APPSOFT\Apply\TmpImage\TransCmd\"
        If DirExists(strCmdPath) = False Then strCmdPath = "F:\APPSOFT\Apply\TmpImage\TransCmd\"
    
        If Trim(strCmdPath) = "" Then
            MsgBox "请指定传输命令所在目录."
            Exit Sub
        End If
    
        If DirExists(strCmdPath) = False Then
            MsgBox "目录 [" & strCmdPath & "] 无效."
            Exit Sub
        End If
        
        gstrCmdPath = strCmdPath
        
        Call OpenTrayIcon
        
        Call StartServer(strCmdPath)
        
    End If
    
    
    
Exit Sub
errHandle:
    MsgBox Err.Description
End Sub


Private Sub StartServer(ByVal strCmdPath As String, Optional ByVal blnIsSingleProcess As Boolean = False)
    Dim objMain As New frmMain
    Dim objForm As Object
    
On Error GoTo errHandle
    Call objMain.Start(strCmdPath, blnIsSingleProcess)
    
    If gblnSingle Then
        '卸载所有窗体
        For Each objForm In Forms
            Unload objForm
        Next
    End If
Exit Sub
errHandle:
    MsgBox Err.Description
End Sub


Private Sub OpenTrayIcon()
'打开托盘图标
    frmTrayIcon.Show
    frmTrayIcon.Hide
End Sub


'Public Sub ResetTrayIcon(ByVal lngTransState As TTrayState)
'    '0-常规，1-上传，2-下载
'    Call frmTrayIcon.ResetTrayIcon(lngTransState)
'End Sub


Public Sub ShowTrayMsg(ByVal strMsg As String, ByVal lngMsgType As Long)
    '显示消息
    Call frmTrayIcon.ShowMessage(strMsg, lngMsgType)
End Sub
