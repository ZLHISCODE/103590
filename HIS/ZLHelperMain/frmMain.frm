VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ZLHIS�ͻ��˷���"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6780
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":6852
   ScaleHeight     =   3690
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Visible         =   0   'False
   Begin VB.ComboBox cboLogLevel 
      BackColor       =   &H00FAEBDE&
      Height          =   300
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2010
      Width           =   4935
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5400
      TabIndex        =   3
      Top             =   3000
      Width           =   1100
   End
   Begin VB.PictureBox picNotify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   4080
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Timer tmrThis 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   1920
   End
   Begin MSWinsockLib.Winsock wskListener 
      Left            =   4080
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   3120
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblLogPath 
      AutoSize        =   -1  'True
      BackColor       =   &H00DDECFA&
      BackStyle       =   0  'Transparent
      Caption         =   "C:\APPSOFT\Log\��־����\ZLHelperMain_SessionID_1_V1.log"
      Height          =   360
      Left            =   1440
      TabIndex        =   8
      Tag             =   "4680"
      Top             =   2595
      Width           =   5160
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblServerName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "192.168.33.201:1521/TESTBASE35"
      Height          =   180
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   2700
   End
   Begin VB.Label lblLog 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������־�ļ���"
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   2595
      Width           =   1260
   End
   Begin VB.Label lblLogLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������־����"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   2070
      Width           =   1260
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���������ݿ⣺"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Image imgMain 
      Height          =   720
      Left            =   240
      Picture         =   "frmMain.frx":1A15A
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lblComent 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDB986&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":1B024
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   5295
   End
   Begin VB.Image imgNotify 
      Height          =   240
      Left            =   5160
      Picture         =   "frmMain.frx":1B0B1
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@ģ�� frmMain-2019/7/2
'@��д lshuo
'@����
'
'@����
'
'@��ע
'
Option Explicit
'---------------------------------------------------------------------------
'                0��API�ͳ�������
'---------------------------------------------------------------------------
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'---------------------------------------------------------------------------
'                1���������
'---------------------------------------------------------------------------
Private marrTmp                             As Variant
Private mblnListenOk                        As Boolean

Private mblnShow                            As Boolean
Private mblnCurShow                         As Boolean
Private mobjCurJob                          As clsJob                       '��ǰ�ķ���������
Private mblnHaveHisProcess                  As Boolean
Private mstrServer                          As String
Private mstrTmp                             As String
Private mblnFind                            As Boolean

'---------------------------------------------------------------------------
'                2�����Ա����붨��
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                3����������
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                4��˽�з���
'---------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
'����           AddServerException
'����           ���һ�����������������б�
'����ֵ
'����б�:
'������         ����                    ˵��
'strServer      String                  ��ӵķ���������
'-------------------------------------------------------------------------------------------------
Private Sub AddServerException(ByVal strServer As String)
    Dim objServerExp        As clsServerInfo
    Dim i                   As Long
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.frmMain.AddServerException")
    If InCollection(gcllExecption, "K_" & strServer) Then
        gcllExecption.Remove "K_" & strServer
    End If
    'Ѱ�ҵ����һ��û���쳣�ķ���������
    For i = 1 To gcllExecption.Count
        If gcllExecption(i).TryTimes <> 0 Then
            Exit For
        End If
    Next
    Set objServerExp = New clsServerInfo
    objServerExp.Server = strServer
    If gcllExecption.Count = 0 Or i > gcllExecption.Count Then
        gcllExecption.Add objServerExp, "K_" & strServer
    Else
        gcllExecption.Add objServerExp, "K_" & strServer, i
    End If
    Call Logger.PopMethod("ZLHelperMain.frmMain.AddServerException")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.frmMain.AddServerException") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.frmMain.AddServerException")
End Sub
'--------------------------------------------------------------------------------------------------
'����           MoveServerExceptionLast
'����           ���һ�����������������б�
'����ֵ
'����б�:
'������         ����                    ˵��
'strServer      String                  ��ӵķ���������
'blnExecute     Boolean                 �Ƿ�����ִ�е��µ�
'-------------------------------------------------------------------------------------------------
Private Sub MoveServerExceptionLast(ByVal strServer As String, Optional ByVal blnExecute As Boolean)
    Dim objServerExp        As clsServerInfo
    
    Set objServerExp = gcllExecption("K_" & strServer)
    If blnExecute Then objServerExp.LastTry = GetTickCount
    gcllExecption.Remove "K_" & strServer
    gcllExecption.Add objServerExp, "K_" & strServer
End Sub

'@����    RefreshServer
'   ˢ�·������б�
'@����ֵ
'
'@����:
'strServer String In(Optional)
'   ����ʱ�����������б���ڸ÷��������򣬲��ٻ���
'@��ע
'
Private Sub RefreshServer(Optional ByVal strServer As String)
    
End Sub

'---------------------------------------------------------------------------
'                5�����󷽷����¼�
'---------------------------------------------------------------------------
Private Sub cboLogLevel_Click()
    If cboLogLevel.Tag = "" Then Exit Sub
    If LenB(Environment.SessionUserSID) <> 0 Then
        If Environment.Is64bitOS Then
            Call Registry.SetRegValue("HKEY_USERS\" & Environment.SessionUserSID & "\Software\WOW6432Node\VB and VBA Program Settings\ZLSOFT\����ģ��\��־����\����", gobjFSO.GetFileName(Environment.StartExePath), cboLogLevel.ListIndex & "", True)
        Else
            Call Registry.SetRegValue("HKEY_USERS\" & Environment.SessionUserSID & "\Software\VB and VBA Program Settings\ZLSOFT\����ģ��\��־����\����", gobjFSO.GetFileName(Environment.StartExePath), cboLogLevel.ListIndex & "", True)
        End If
    End If
    Logger.CurrentLogLevel = cboLogLevel.ListIndex
End Sub

Private Sub cmdOK_Click()
    mblnShow = Not mblnShow
End Sub

Private Sub Form_Load()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.frmMain.Form_Load")
    On Error GoTo ErrH
    SetWindowPos Me.hwnd, -1, (Screen.Width - Me.Width) / 2 / 15, (Screen.Height - Me.Height) / 2 / 15, 0, 0, 1
    cboLogLevel.AddItem "����¼��־"
    cboLogLevel.AddItem "���󼶱�(ֻ��¼����)"
    cboLogLevel.AddItem "���漶��(����������Ϣ����󼶱�)"
    cboLogLevel.AddItem "��Ϣ����(������Ҫ��Ϣ�뾯�漶��)"
    cboLogLevel.AddItem "���Լ���(����������Ϣ����Ҫ��Ϣ����)"
    cboLogLevel.AddItem "���ټ���(����������Ϣ����Լ���)"
    cboLogLevel.AddItem "��¼������־"
    cboLogLevel.ListIndex = IIf(Logger.CurrentLogLevel < LogLevel_LogOFF, LogLevel_LogOFF, IIf(Logger.CurrentLogLevel > LogLevel_AllLog, LogLevel_AllLog, Logger.CurrentLogLevel))
    cboLogLevel.Tag = "�Ѿ�����"
    lblLogPath.Caption = Logger.LogFile
    Call AddIcon(picNotify.hwnd, imgNotify.Picture, "ZLHIS�ͻ��˷�")
    '�󶨵�����������IP��,�����Ѿ���
    On Error Resume Next
    wskListener.Bind "7534", "0.0.0.0"
    wskListener.Listen
    mblnListenOk = wskListener.State = sckListening
    lblServerName.Caption = "��"
    tmrThis.Enabled = True
    Call Logger.PopMethod("ZLHelperMain.frmMain.Form_Load")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.frmMain.Form_Load") = 1 Then
        Resume
    End If
    tmrThis.Enabled = True
    Call Logger.PopMethod("ZLHelperMain.frmMain.Form_Load")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i           As Long
    On Error Resume Next
'    Erase marrUserCookie
    For i = 1 To wskServer.UBound
        wskServer(i).Close
        Unload wskServer(i)
    Next
    wskListener.Close
    
    Set mobjCurJob = Nothing
    Call RemoveIcon(picNotify.hwnd)
    Call ProcessExit
End Sub



Private Sub picNotify_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '--------------------------------------------------------------------------------------------------
    '����:  ����picNotify�ĸ��ִ����¼�
    '--------------------------------------------------------------------------------------------------

    Select Case Hex(x) '
        Case "1E3C"     'Right-Button-Down
        Case "1E4B"     'Right-Button-Up
        Case "1830"     'Right-Button-Down LARGE FONTS '
        Case "1E1E"     'Left-Button-up
        Case "1E0F"     'Left-Button-Down '
        Case "1E2D"     'Left-Button-Double-Click '
            On Error Resume Next
            mblnShow = Not mblnShow
        Case "1824"     'Left-Button-Double-Click LARGE FONTS
        Case "1E5A"     'Right-Button-Double-Click '
    End Select '
End Sub
Private Sub tmrThis_Timer()
    On Error GoTo ErrH
    tmrThis.Enabled = False
    Call Logger.PushMethod("ZLHelperMain.frmMain.tmrThis_Timer")
    If mblnCurShow <> mblnShow Then
        If mblnShow Then
            ShowWindow Me.hwnd, 5
        Else
            ShowWindow Me.hwnd, 0
        End If
        mblnCurShow = mblnShow
    End If
    If gobjMetux Is Nothing Then
        Set gobjMetux = New clsMutex
        If Not gobjMetux.CheckMutex(G_SINGLE_INSTANCE) Then
            Set gobjHelperMainRECEIVE = New clsMemoryShare
            Set gobjHelperMainSend = New clsMemoryShare
            If Not gobjHelperMainRECEIVE.CreateMemoryShare(G_HELPER_RECEIVE) Or Not gobjHelperMainSend.CreateMemoryShare(G_HELPER_SEND) Then
                Set gobjHelperMainRECEIVE = Nothing
                Set gobjHelperMainSend = Nothing
                Set gobjMetux = Nothing
            End If
        Else
            Set gobjMetux = Nothing
        End If
    End If
    
    If Not gobjMetux Is Nothing Then
        If gobjHelperMainRECEIVE.ReadMemoryOnce() Then
            If gobjServerQueue.Current = "HELPERUPGRADE SAVEANDEXIT" Then
                glngSendProcess = gobjHelperMainRECEIVE.ProcessID
                gblnExitProcess = True
            ElseIf gobjServerQueue.Current = "EXIT" Then
                glngSendProcess = 0
                gblnExitProcess = True
            Else
                gobjServerQueue.EnQueue gobjHelperMainRECEIVE.Data
            End If
            Logger.DebugEx "��ȡ���������̴�������������Ϣ��", "��Ϣ", gobjHelperMainRECEIVE.Data
        End If
    End If
'    gobjServerQueue.EnQueue "127.0.0.1:1521/TESTBASE35"
    If Not gobjServerQueue.IsEmpty Then
        If Not IsEmptyArray(Process.ProcessesByProcessName("ZLHIS+.EXE")) Then
            mblnHaveHisProcess = True
        ElseIf Not IsEmptyArray(Process.ProcessesByProcessName("ZLLIS+.EXE")) Then
            mblnHaveHisProcess = True
        End If
    End If
    '������ķ��������뵽�������
    Do While Not gobjServerQueue.IsEmpty
        'EXCFUNC DB=192.168.33.201:1521/TESTBASE35
        If gobjServerQueue.Current Like "EXCFUNC DB=*" Then
            mstrTmp = Mid(gobjServerQueue.Current, Len("EXCFUNC DB=*"))
            '�Զ��������صķ����������Ǹ÷������Ѿ����������,����Ҫ�ټ���������б�
            If mstrTmp = mstrServer Then
                mstrTmp = ""
            End If
        Else
            mstrTmp = gobjServerQueue.Current
        End If
        If LenB(mstrTmp) <> 0 Then
            mblnFind = False
            If Not mobjCurJob Is Nothing Then
                '��ǰ����ķ������ӵ�֪ͨ�����Զ�����״̬
                If mobjCurJob.Server = mstrTmp Then
                    mblnFind = True
                    If mobjCurJob.IsNeedRestartJob() Then
                        Set mobjCurJob = Nothing
                        Call gcllExecption("K_" & mstrTmp).Restart
                    End If
                End If
            End If
            If Not mblnFind Then
                Call AddServerException(mstrTmp)
            End If
            If mblnHaveHisProcess Then
                gcllExecption("K_" & mstrTmp).IsDelay = True
            End If
        End If
        gobjServerQueue.DeQueue
    Loop
    mstrServer = ""
    If gcllExecption.Count <> 0 Then
        If gblnExitProcess Then
            Call SaveEnv
            Logger.DebugEx "�˳�������ѵ��"
            Unload Me
            Exit Sub
        End If
        If mobjCurJob Is Nothing Then
            If gcllExecption(1).IsCanTryAgain Then '��������ִ�У���ִ��
                Set mobjCurJob = New clsJob
                Call mobjCurJob.InitJobServer(gcllExecption(1))
                Logger.DebugEx "��������", "������", gcllExecption(1).Server
                lblServerName.Caption = mobjCurJob.Server
            Else
                If gcllExecption(1).IsCanDeleteServer Then
                    gcllExecption.Remove "K_" & mobjCurJob.Server
                Else
                    lblServerName.Caption = "��"
                    '�������ԣ����ƶ������
                    Call MoveServerExceptionLast(gcllExecption(1).Server)
                End If
            End If
        End If
    Else
        lblServerName.Caption = "��"
        Set mobjCurJob = Nothing
    End If

    '�Ƿ������Ϣ��ʾ
    If Not gblnMsgBox Then
        '������֤����24Сʱ�������µĳ�������
        If Not mobjCurJob Is Nothing Then
            If mobjCurJob.IsRestart Then
                Set mobjCurJob = Nothing
            End If
        End If
        If Not mobjCurJob Is Nothing Then
            Select Case mobjCurJob.FinishJob()
                Case SC_Finish
                    gcllExecption.Remove "K_" & mobjCurJob.Server
                    mstrServer = mobjCurJob.Server
                    Set mobjCurJob = Nothing
                Case SC_Delay
                    Call MoveServerExceptionLast(mobjCurJob.Server, True)
                    mstrServer = mobjCurJob.Server
                    Set mobjCurJob = Nothing
                Case SC_Wait
                    gcllExecption("K_" & mstrTmp).IsDelay = True
                    mstrServer = mobjCurJob.Server
                    Set mobjCurJob = Nothing
                Case Else
                    '������ѵ����
            End Select
        End If
    End If
    '�󶨵�����������IP��
    If Not mblnListenOk Then
        On Error Resume Next
        wskListener.Bind 7534, "0.0.0.0"
        wskListener.Listen
        mblnListenOk = wskListener.State = sckListening
        If Err.Number <> 0 Then
             Logger.Error "�󶨶˿�ʧ��", "����", Err.Number & "-" & Err.Description
        Else
             Logger.DebugEx "�����°�", "���", mblnListenOk
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.frmMain.tmrThis_Timer")
    tmrThis.Enabled = True
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.frmMain.tmrThis_Timer") = 1 Then
        Resume
    End If
    tmrThis.Enabled = True
    Call Logger.PopMethod("ZLHelperMain.frmMain.tmrThis_Timer")
End Sub

Private Sub wskListener_ConnectionRequest(ByVal requestID As Long)
    Dim i As Long
      
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.frmMain.wskListener_ConnectionRequest", requestID)
    For i = 1 To wskServer.UBound
        If wskServer(i).State = sckClosed Then
            wskServer(i).Accept requestID
            Call Logger.PopMethod("ZLHelperMain.frmMain.wskListener_ConnectionRequest")
            Exit Sub
        End If
    Next
    Load wskServer(i)
    wskServer(i).Accept requestID
    Call Logger.PopMethod("ZLHelperMain.frmMain.wskListener_ConnectionRequest")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.frmMain.wskListener_ConnectionRequest") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.frmMain.wskListener_ConnectionRequest")
End Sub

Private Sub wskListener_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.frmMain.wskListener_Error", Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
    Logger.Error "ͨѶ�쳣", "Error", Number & "-" & Description
    Call Logger.PopMethod("ZLHelperMain.frmMain.wskListener_Error")
    mblnListenOk = False
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.frmMain.wskListener_Error") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.frmMain.wskListener_Error")
End Sub

Private Sub wskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData     As String
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.frmMain.wskServer_DataArrival", Index, bytesTotal)
    wskServer(Index).GetData strData
    Logger.Info "GetData", "Data", strData
    If strData = "����Զ��" Then
        RunCommand "REG ADD HKLM\SYSTEM\CurrentControlSet\Control\Terminal"" ""Server /v fDenyTSConnections /t REG_DWORD /d 0 /f"
        wskServer(Index).SendData "YES"
        Logger.DebugEx "SendData", "Data", "YES"
    Else
        marrTmp = Split(UCase(strData) & ";;;", ";")
        '�ȼ�����У�����ѯʱ�����д�����ֹ���ڽ���������
        Call gobjServerQueue.EnQueue(Trim(marrTmp(1)))
        strData = "ClientResonse;" & marrTmp(2) & ";" & Environment.IP & ";" & Environment.ComputerName
        '�Ե�ǰ���������Ӧ��û�ж��޸��������Ӧ���޸��ǿͻ�������������ͨ�������������Ĵ���
        wskServer(Index).SendData strData
        Logger.DebugEx "SendData", "Data", strData
    End If
    
    Call Logger.PopMethod("ZLHelperMain.frmMain.wskServer_DataArrival")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.frmMain.wskServer_DataArrival") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.frmMain.wskServer_DataArrival")
End Sub

Private Sub wskServer_SendComplete(Index As Integer)
    wskServer(Index).Close
End Sub

