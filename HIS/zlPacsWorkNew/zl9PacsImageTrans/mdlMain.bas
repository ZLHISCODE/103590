Attribute VB_Name = "mdlMain"
Option Explicit

Public Const C_BIGFILE_TAG As String = "[BF]"


Public gstrCmdPath As String
Public gblnIsFailed As Boolean
Public gblnSingle As Boolean    '���������������
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
        '���������Ľ��̽����ļ�����
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
            MsgBox "��ָ��������������Ŀ¼."
            Exit Sub
        End If
    
        If DirExists(strCmdPath) = False Then
            MsgBox "Ŀ¼ [" & strCmdPath & "] ��Ч."
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
        'ж�����д���
        For Each objForm In Forms
            Unload objForm
        Next
    End If
Exit Sub
errHandle:
    MsgBox Err.Description
End Sub


Private Sub OpenTrayIcon()
'������ͼ��
    frmTrayIcon.Show
    frmTrayIcon.Hide
End Sub


'Public Sub ResetTrayIcon(ByVal lngTransState As TTrayState)
'    '0-���棬1-�ϴ���2-����
'    Call frmTrayIcon.ResetTrayIcon(lngTransState)
'End Sub


Public Sub ShowTrayMsg(ByVal strMsg As String, ByVal lngMsgType As Long)
    '��ʾ��Ϣ
    Call frmTrayIcon.ShowMessage(strMsg, lngMsgType)
End Sub
