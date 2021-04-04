VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum LoadCommand
    COMMAND_CAPIMG_UPLOAD           '�ɼ�ͼ���ϴ�
    COMMAND_RPTIMG_DOWNLOAD         'ͼ������
End Enum

Public Enum FileLoadType    '�ļ��ϴ���ʽ
    Normal                  '��ͨ
    Service                 'zlPacsServices
End Enum

Public Type TFtpDeviceInf
    strDeviceId As String
    strFTPIP As String
    strFTPUser As String
    strFTPPwd As String
    strFtpDir As String
    strSDDir As String
    strSDUser As String
    strSDPswd As String
End Type

Public Enum TMediaType
    imgTag = 0   'ͼ����
    MULFRAMETAG = 1 '����ͼ
    VIDEOTAG = 2 '��Ƶ���
    AUDIOTAG = 3 '��Ƶ���
End Enum

Public Sub InitPacsControlWithInPro(cnOracle As ADODB.Connection, _
    ByVal lngOwnerHwnd As Long, _
    ByVal lngSys As Long, _
    ByVal lngMoudle As Long, _
    Optional objOwner As Object = Nothing, _
    Optional blnOpenDebug As Boolean = False)
'��ʼ������������ӣ�����ͬһ���̣�

'    Call InitCommon(cnOracle)
    
    Set gobjOwner = objOwner
    glngSys = lngSys
    glngMoudle = lngMoudle
    glngOwnerHwnd = lngOwnerHwnd
    gblnOpenDebug = blnOpenDebug
End Sub


Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
'    If gblnOpenDebug Or blnIsForce Then
'        OutputDebugString Now & " |---> " & strDebug
'    End If
    Call mdlPublic.BUGEX(strDebug, blnIsForce)
End Sub

Public Function MsgboxEx(objParent As Object, sPrompt As String, Optional dwStyle As Long, Optional sTitle As String) As Long
    Dim lngHwnd As Long

BUGEX "MsgboxEx 1"

    If objParent Is Nothing Then
        lngHwnd = GetActiveWindow
    Else
        lngHwnd = objParent.hWnd
    End If

    If lngHwnd = GetDesktopWindow Or lngHwnd = 0 Then
BUGEX "MsgboxEx 2 GetForegroundWindow" & " DesktopWindowHwnd:" & lngHwnd
        lngHwnd = GetForegroundWindow
    End If

BUGEX "MsgBoxEx 3 Hwnd:" & lngHwnd

    MsgboxEx = mdlMsgBox.MsgboxEx(lngHwnd, sPrompt, dwStyle, sTitle)

    '���򿪵���״̬������д�����Ϣ�����Զ���ʾ
    If err.Number <> 0 And gblnOpenDebug Then
        Call mdlMsgBox.MsgboxEx(lngHwnd, "errSource:" & err.Source & "  errDescription:" & err.Description, vbOKOnly, CON_STR_HINT_TITLE)
    End If

BUGEX "MsgBoxEx End"
End Function


Private Sub Class_Initialize()
'��ʼ�������Դ
    Set gcnOracle = Nothing
    Set gobjOwner = Nothing
    
    glngOwnerHwnd = 0
End Sub


Private Sub Class_Terminate()
'�ͷ������Դ
    
End Sub