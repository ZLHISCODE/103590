VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmHL7Main 
   Caption         =   "����HL7����"
   ClientHeight    =   3330
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5400
   Icon            =   "frmHL7Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   5400
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer tmFileInput 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1800
      Top             =   2040
   End
   Begin VB.Timer tmIdle 
      Interval        =   60000
      Left            =   4080
      Top             =   1320
   End
   Begin MSWinsockLib.Winsock wsSend 
      Left            =   120
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmHISAction 
      Interval        =   10000
      Left            =   3360
      Top             =   1320
   End
   Begin VB.Timer tmLinkTimeOut 
      Index           =   0
      Interval        =   1000
      Left            =   2520
      Top             =   1320
   End
   Begin VB.Timer tmMsgProcess 
      Interval        =   10000
      Left            =   1800
      Top             =   1320
   End
   Begin MSWinsockLib.Winsock wsHL7Server 
      Index           =   0
      Left            =   120
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   1296
      ButtonWidth     =   1455
      ButtonHeight    =   1138
      Appearance      =   1
      ImageList       =   "imgGray"
      HotImageList    =   "imgColor"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��������"
            Key             =   "��������"
            ImageKey        =   "��������"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ֹͣ����"
            Key             =   "ֹͣ����"
            ImageKey        =   "ֹͣ����"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "����"
            Object.ToolTipText     =   "����"
            Object.Tag             =   "����"
            ImageKey        =   "����"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�"
            Key             =   "�˳�"
            Object.ToolTipText     =   "�˳�"
            Object.Tag             =   "�˳�"
            ImageKey        =   "�˳�"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   120
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":08CA
            Key             =   "Ԥ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":0AE4
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":0CFE
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":0F18
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":1132
            Key             =   "��¼"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":182C
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":1F26
            Key             =   "���"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":2620
            Key             =   "����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":2D1A
            Key             =   "����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":3414
            Key             =   "�ķ�"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":3B0E
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":4208
            Key             =   "����"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":4902
            Key             =   "�޸�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":4FFC
            Key             =   "ֹͣ����"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":56F6
            Key             =   "����"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":5DF0
            Key             =   "��������"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   720
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":64EA
            Key             =   "Ԥ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":6704
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":691E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":6B38
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":6D52
            Key             =   "��¼"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":744C
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":7B46
            Key             =   "���"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":8240
            Key             =   "����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":893A
            Key             =   "����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":9034
            Key             =   "�ķ�"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":972E
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":9E28
            Key             =   "����"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":A522
            Key             =   "�޸�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":AC1C
            Key             =   "ֹͣ����"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":B316
            Key             =   "����"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHL7Main.frx":BA10
            Key             =   "��������"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsHL7Link 
      Index           =   0
      Left            =   720
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mmuParaSetup 
         Caption         =   "��������"
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuService 
      Caption         =   "����"
      Begin VB.Menu munStartService 
         Caption         =   "��������"
      End
      Begin VB.Menu munStopService 
         Caption         =   "ֹͣ����"
      End
      Begin VB.Menu mnuService_1 
         Caption         =   "-"
      End
      Begin VB.Menu mmuShowService 
         Caption         =   "��ʾ��ǰ����"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "��־"
      Begin VB.Menu mmuLog 
         Caption         =   "��¼��־"
         Begin VB.Menu mmuProcessLog 
            Caption         =   "��¼ͨѶ��־"
            Index           =   1
         End
         Begin VB.Menu mmuProcessLog 
            Caption         =   "��¼������־"
            Index           =   2
         End
         Begin VB.Menu mmuProcessLog 
            Caption         =   "��¼��ϸ��־"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mmuShowLog 
         Caption         =   "��ʾ��־"
         Index           =   1
      End
      Begin VB.Menu mmuShowLog 
         Caption         =   "��ʾ��Ϣ��¼"
         Index           =   2
      End
      Begin VB.Menu mmuShowLog 
         Caption         =   "��ʾ������־"
         Index           =   3
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmHL7Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastState As Integer

Private WithEvents mobjIcon As clsTaskIcon  '������
Attribute mobjIcon.VB_VarHelpID = -1
Private mfrmShowLog As frmShowLog

Private intWSListenerCount As Integer       '����������winsock������
Private intWSLinkCount As Integer           '�������������winsock������

Private mblnActionProcess As Boolean        '���붯������
Private mblnSendConnect As Boolean          '����Socket���ӳɹ�
Private mblnSendACK As Boolean              '����Socket���յ�ACK��Ӧ

Private mlngIdleTime As Long                '�ۼ�Socket���е���ʱ�䣬����Ϊ��λ


Private Sub Form_Load()
        
    On Error Resume Next
    
    If WindowState = vbMinimized Then
        LastState = vbNormal
        Me.Hide
    Else
        LastState = WindowState
    End If
    
    
    '----------��������ͼ��
    Set mobjIcon = New clsTaskIcon
    mobjIcon.frmHwnd = tbrMain.hwnd ' hwnd
    mobjIcon.Icon = Icon.Handle
    mobjIcon.Message = "����HL7��������"
    mobjIcon.AddIcon
    '----------��������ͼ��
    
    '����������Access����־��¼��������
    gstrAccessPath = App.Path & "\ZlHL7Log"
    gstrAccessName = gstrAccessPath & ".mdb"
    
    With gcnAccess
        .ConnectionString = "DBQ=" & gstrAccessName & ";DefaultDir=" & App.Path & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then MsgBox "���ܴ򿪱�����־�ļ���ϵͳ���޷���¼���չ��̣�", vbInformation, gstrSysName
    End With
    
    ''��ʼ�����ݣ�����Ĭ��ֵ
    intWSListenerCount = 0
    intWSLinkCount = 0
    ReDim gstrMsgQueue(0) As String
    ReDim gintTimeOut(0) As Integer
    gintQueueIndex = 1
    gblnQueueBusy = False
    mblnActionProcess = False
    mblnSendConnect = False
    mblnSendACK = False
    gblnServiceStart = True
    tbrMain.Buttons(3).Enabled = True
    tbrMain.Buttons(2).Enabled = False
    
    gstrRegPath = "˽��ģ��\HL7�ӿ�"
    
    '��ȡ���ݿ��ע���Ĳ���
    Call ReadPara
    
    '������Ϣ���շ���
    Call subInputServiceSwitch(1)
    
    '���ô����Զ���С��
    Me.WindowState = vbMinimized
    
    '�����������¼����Ϣ��־
    Call WriteMessageLog("HL7��������", Now & " HL7�����������汾Ϊ��" & App.Major & "." & App.Minor & "." & App.Revision & "����¼�û�Ϊ��" & gstrDbUser)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        '���˳���ֻ����С��
        Cancel = True
        Me.WindowState = vbMinimized
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If WindowState = 1 Then
        Me.Hide
        Exit Sub
    End If
    
    If WindowState <> vbMinimized Then
        LastState = WindowState
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim wsSocket As Winsock
    '�������ͼ��
    mobjIcon.DelIcon
    Set mobjIcon = Nothing
    
    'ֹͣ��Ϣ���շ���
    Call subInputServiceSwitch(0)
    
    '�������winsock����
    For Each wsSocket In wsHL7Server
        If wsSocket.State <> sckClosed Then wsSocket.Close
        If wsSocket.Index <> 0 Then Unload wsSocket
    Next
    For Each wsSocket In wsHL7Link
        If wsSocket.State <> sckClosed Then wsSocket.Close
        If wsSocket.Index <> 0 Then Unload wsSocket
    Next
    
    '���˳������¼����Ϣ��־
    Call WriteMessageLog("HL7�����˳�", Now & " HL7�����˳�")
    
End Sub

Private Sub mmuParaSetup_Click()
    Call frmParaSetup.zlSohwMe(Me)
End Sub


Private Sub mmuProcessLog_Click(Index As Integer)
    Dim i As Integer
    
    mmuProcessLog(Index).Checked = Not mmuProcessLog(Index).Checked
    gblnProcessLog = mmuProcessLog(Index).Checked
    glngProcessLogLevel = Index
    
    For i = 1 To 3
        If i <> Index Then
            mmuProcessLog(i).Checked = False
        End If
    Next i
End Sub

Private Sub mmuShowLog_Click(Index As Integer)
    Call subShowLog(Index)
End Sub

Private Sub mmuShowService_Click()
    Call subShowLog(4)
End Sub

Private Sub mnuFileQuit_Click()
    Dim dblSleepTime As Double
    
    gblnServiceStart = False
    
    dblSleepTime = timeGetTime
    
    '��ʱ2�����ٹرգ��Ա���ֹͣ���ͷ���
    While timeGetTime < dblSleepTime + 2000
        DoEvents
    Wend
    
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    gzlComLib.ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mobjIcon_MouseLeftDBClick()
    '�����ʾ��־��ģʽ�����Ѿ����򿪣����˳���������ִ���
    If mfrmShowLog Is Nothing Then
        If WindowState <> 1 Then
            WindowState = vbMinimized
            If WindowState = vbMinimized Then
                Me.Hide
            Else
                Me.Show
            End If
        Else
            WindowState = vbNormal
            Me.Show
        End If
    End If
End Sub

Private Sub munStartService_Click()
    gblnServiceStart = True
    
    tmHISAction.Enabled = True
    
    Call subInputServiceSwitch(1)
End Sub

Private Sub munStopService_Click()
    gblnServiceStart = False
    
    tmHISAction.Enabled = False
    
    Call subInputServiceSwitch(0)
End Sub

Private Sub subInputServiceSwitch(itype As Integer)
    'itype=0��ֹͣ����;itype=1��������
        
    If itype = 0 Then
        'ֹͣ���з���
        Call funListenPorts(0)
        tmFileInput.Enabled = False
    Else
        '��������ʱ��ȷ��ֻ��������һ�ַ���
        If gintInputDataType = 0 Then
            '����HL7��������
            Call funListenPorts(1)
            tmFileInput.Enabled = False
        Else
            tmFileInput.Enabled = True
            Call funListenPorts(0)
        End If
    End If
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "��������"
            tbrMain.Buttons(3).Enabled = True
            Button.Enabled = False
            Call munStartService_Click
        Case "ֹͣ����"
            tbrMain.Buttons(2).Enabled = True
            Button.Enabled = False
            Call munStopService_Click
        Case "����"
            Call mnuHelpAbout_Click
        Case "�˳�"
            Me.WindowState = vbMinimized ' mnuFileQuit_Click
    End Select
End Sub

Private Sub tbrMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mobjIcon.MouseState x
End Sub

Private Function ReadPara() As Boolean
'------------------------------------------------
'���ܣ� ��ȡ��������
'��������
'���أ�True --�ɹ���False -- ʧ��
'------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    
    On Error GoTo err
    
    '��ȡ����IP��ַ
    gstrLocalIP = funcGetLocalIP & ",127.0.0.1"
    
    strSQL = "Select IP��ַ,�˿ں�,��������,���ͳ�������,�����豸����,���ճ�������,�����豸���� From zlhis.hl7�������� " & _
                " Where instr([1],IP��ַ)>0"
    Set rsTemp = gzlDatabase.OpenSQLRecord(strSQL, "��ȡHL7��������", gstrLocalIP)
    
    ReDim HL7Services(rsTemp.RecordCount) As THL7Service
    i = 1
    
    While Not rsTemp.EOF
        HL7Services(i).strIP = Nvl(rsTemp!IP��ַ)
        HL7Services(i).lngPort = Val(Nvl(rsTemp!�˿ں�, 0))
        HL7Services(i).intServiceType = Nvl(rsTemp!��������)
        HL7Services(i).strSendApp = Nvl(rsTemp!���ͳ�������)
        HL7Services(i).strSendFacility = Nvl(rsTemp!�����豸����)
        HL7Services(i).strReceiveApp = Nvl(rsTemp!���ճ�������)
        HL7Services(i).strReceiveFacility = Nvl(rsTemp!�����豸����)
        i = i + 1
        rsTemp.MoveNext
    Wend
    
    gintTimeOutMax = Val(GetSetting("ZLSOFT", gstrRegPath, "��ʱ", 60))
    gintInputDataType = Val(GetSetting("ZLSOFT", gstrRegPath, "������Ϣ��ʽ", 0))
    If gintInputDataType <> 1 Then gintInputDataType = 0
    gstrFileDir = GetSetting("ZLSOFT", gstrRegPath, "�ļ���ϢĿ¼", "")
    gstrFileSuffix = GetSetting("ZLSOFT", gstrRegPath, "�ļ���Ϣ��׺", "")
    gstrFileBackupDir = GetSetting("ZLSOFT", gstrRegPath, "�ļ���Ϣ����Ŀ¼", "")
    
    Exit Function
    
err:
    If gzlComLib.ErrCenter() = 1 Then Resume
    Call gzlComLib.SaveErrLog
End Function

Private Sub funListenPorts(itype As Integer)
'-----------------------------------------------------------------------------
'����:������ֹͣ�Է���˿ڵ�����
'����: iType = 0 ֹͣ������iType = 1 ����������
'���أ�True -- �ɹ���False -- ʧ��
'-----------------------------------------------------------------------------
    Dim strPort As String
    Dim i As Integer
    Dim lngResult As Long
    
    On Error Resume Next
    
    '�����������շ���˿�
    For i = 1 To UBound(HL7Services)
        If HL7Services(i).intServiceType = 1 Then
            If InStr("," & strPort, HL7Services(i).lngPort) = 0 Then
                strPort = strPort & "," & HL7Services(i).lngPort
                If itype = 0 Then   'ֹͣ����
                    'ֹͣ����
                    If funWinsockUnlisten(HL7Services(i).lngPort) = 0 Then
                        HL7Services(i).Started = False
                    End If
                ElseIf itype = 1 Then   '��������
                    '��������
                    lngResult = funWinSockListen(HL7Services(i).lngPort)
                    If lngResult = 0 Then
                        HL7Services(i).Started = True
                    Else
                        HL7Services(i).Started = False
                        MsgBox "�˿ڣ�" & HL7Services(i).lngPort & "�ѱ�ʹ�ã�" & _
                                " ϵͳ�޷����������������ü����˿ڡ�", vbExclamation, gstrSysName
                    End If
                End If
            Else
                HL7Services(i).Started = IIf(itype = 0, False, True)
            End If
        End If
    Next i
        
End Sub

Private Function funWinSockListen(lngPort As Long) As Long
'-----------------------------------------------------------------------------
'����:����winsock�˿ڵ�����
'����:
'       lngPort ---�����Ķ˿ں�
'���أ�0 -- �ɹ���1 -- ʧ��,�˿ڱ�ռ�ã�2-ʧ�ܣ�winsock��������ȷ�������µ�winsockʧ��
'       3 -- ʧ�ܣ���������
'-----------------------------------------------------------------------------
    'wsHL7Server()�����У�wsHL7Server��1��֮���ʵ�����������
    '������ͨѶ����󣬴���һ���µ�wsHL7Linkʵ����������������
    
    On Error GoTo err
    
    intWSListenerCount = intWSListenerCount + 1
    
    '�ȼ�鵱ǰwinsock�Ƿ���Ҫ����
    If wsHL7Server.Count = intWSListenerCount Then
        Load wsHL7Server(intWSListenerCount)
    Else
        'intWSListenerCount�������ԣ�ֱ�ӷ��ش���
        funWinSockListen = 2
        Exit Function
    End If
    
    
    '�ȹر�һ��winsock
    wsHL7Server(intWSListenerCount).Close
    wsHL7Server(intWSListenerCount).Bind lngPort
    wsHL7Server(intWSListenerCount).Listen
    If wsHL7Server(intWSListenerCount).State = sckListening Then
        funWinSockListen = 0        '�����ɹ�
    Else
        funWinSockListen = 3        '����ʧ��,δ֪����
    End If
    
    Exit Function
err:
    '������ʾ��ֱ�ӷ��������˿�ʧ��
    If err.Number = 10048 Then
        funWinSockListen = 1    '�˿ڱ�ռ��
    Else
        funWinSockListen = 3    '����ʧ��,δ֪����
    End If
End Function

Private Function funWinsockUnlisten(lngPort As Long) As Long
'-----------------------------------------------------------------------------
'����:ֹͣwinsock�˿ڵ�����
'����:
'       lngPort ---�����Ķ˿ں�
'���أ�0 -- �ɹ���1 -- ʧ�ܣ���������
'-----------------------------------------------------------------------------
    Dim wsListener As Winsock
    
    On Error GoTo err
    
    'Ĭ�Ϸ���ֵ��0�����û���ҵ��˿ڣ���Ϊ����˿ڵ�������ֹͣ��
    funWinsockUnlisten = 0
    
    '���Ȳ������ڼ����ö˿ڵ�winsock����
    For Each wsListener In wsHL7Server
        If wsListener.LocalPort = lngPort And wsListener.State = sckListening And wsListener.Index <> 0 Then
            wsListener.Close
            Unload wsListener
            '���Ҳ�ɾ������������´�������������ʵ��
            'wsHL7Link
            '�����
            Exit For
        End If
    Next
    
    '���ֻʣ��һ��wsHL7Server�����������ö��������
    If wsHL7Server.Count = 1 Then
        intWSListenerCount = 0
    End If
    Exit Function
err:
    funWinsockUnlisten = 1
End Function


Private Sub tmFileInput_Timer()
    'ͨ���ļ���ʽ������Ϣ�Ķ�ʱ��
    '��ʱ��ѯָ��Ŀ¼���ҵ��ļ��󣬽����ļ��������ļ�
    
    Dim strFileName As String
    Dim strHL7Msg As String
    
    On Error GoTo err
    
        
    'ѭ����ȡ���Ŀ¼�µ������ļ�
    'ʹ��dir()������ѯָ��Ŀ¼���ǰ����ļ���������ģ�����ж����Ϣ��������ʱ���п��ܳ���һ���ļ���Ϊ������ԭ��ʼ�ձ��ŵ����
    strFileName = Dir(gstrFileDir & "\*." & gstrFileSuffix)
    While strFileName <> ""
    
       
        '��¼��ϸ��־
        Call WriteProcessLog("tmFileInput_Timer", "1�����յ��ļ���Ϣ��׼������", "�ļ���Ϊ��" & gstrFileDir & "\" & strFileName, 2)
    
        '��ȡ�ļ�
        strHL7Msg = funReadFile(gstrFileDir & "\" & strFileName)
        
        '��¼��ϸ��־
        Call WriteProcessLog("tmFileInput_Timer", "2����ȡ�ļ�����", "�ļ�����ǰ��Σ�" & Left(strHL7Msg, 150), 3)
        
        
        '���ļ����ݼ�����Ϣ����
        Call MsgInQueue(strHL7Msg)
        
        '��¼��ϸ��־
        Call WriteProcessLog("tmFileInput_Timer", "3����Ϣ���", "�ļ�����" & strFileName, 3)
                
        '���ļ��Ƶ�����Ŀ¼������Ŀ¼ÿ��һ��
        Call funBackupHL7File(gstrFileDir, strFileName)
        
        '��¼��ϸ��־
        Call WriteProcessLog("tmFileInput_Timer", "4���ļ���Ϣ����", "�ļ�����" & strFileName, 3)
        
        '��ȡ��һ���ļ�
        strFileName = Dir(gstrFileDir & "\*." & gstrFileSuffix)
        
        '��¼��ϸ��־
        Call WriteProcessLog("tmFileInput_Timer", "5����һ���ļ���", "�ļ�����" & strFileName, 3)
    Wend
    
    Exit Sub
err:
    '������
    If err.Number = 52 Then
        Call WriteLog(5003, err.Number, "tmFileInput_Timer ���ִ��󣬴��������ǣ�" & err.Description & "��dir�ļ���=" & gstrFileDir & "\*." & gstrFileSuffix)
    Else
        Call WriteLog(5003, err.Number, "tmFileInput_Timer ���ִ��󣬴��������ǣ�" & err.Description)
    End If
End Sub

Private Function funBackupHL7File(strFileDir As String, strFileName As String)
    Dim strPath As String
    
    On Error GoTo err
    '���յ������ڣ�����Ŀ¼��Ŀ¼Ϊ\��\��+��
    strPath = "\" + Format(Date, "yyyy") + "\" + Format(Date, "mmdd") + "\"
    
    
    '��������Ŀ¼
    Call MkLocalDir(gstrFileBackupDir + strPath + "\")
    
    '�����ļ�
    Call FileCopy(strFileDir + "\" + strFileName, gstrFileBackupDir + strPath + strFileName)
    Call Kill(strFileDir + "\" + strFileName)
    
    Exit Function
err:
    Call WriteLog(4004, err.Number, "funBackupHL7File ���ִ���strFileDir=" & strFileDir & "��strFileName=" & strFileName & "�����������ǣ�" & err.Description)
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'���ܣ���������Ŀ¼
'������ strDir��������Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Private Function funReadFile(strFileName As String) As String
'��ȡ�������ļ�����

    Dim fn As Integer
    Dim strContent As String
    
    fn = FreeFile
    
    On Error GoTo err
    
    Open strFileName For Binary Access Read As #fn
    strContent = Space(FileLen(strFileName))
    Get #fn, , strContent
    Close #fn
    
    funReadFile = strContent
        
    Exit Function
err:
    '�ر��ļ�
    Close #fn
    Call WriteLog(4002, err.Number, "funReadFile ���ִ���strFileName=" & "�����������ǣ�" & err.Description)
End Function


Private Sub tmHISAction_Timer()
    '��ʱ��ѯHIS���ݿ��HL7��Ϣ��ʱ��
    On Error Resume Next
    'ȷ������󣬻��ܹ���mblnActionProcess���ó�true
    
    If mblnActionProcess = False Then
        mblnActionProcess = True
        Call funActionProcess
        mblnActionProcess = False
    End If
End Sub

Private Sub tmIdle_Timer()
    
    On Error GoTo err
    '���������¼������־
    
    '���м�ʱ��1��������һ��
    
    '�ڳ�������������10����һ�εļ�ʱ����ÿ���賿3��10�֣������־�ļ���С
    mlngIdleTime = mlngIdleTime + 1
    If mlngIdleTime > 10 Then
        '������賿3��15��֮���15����֮�ڣ����ط���Ϣ�������־�ļ�
        If DateDiff("n", "03:15:00", Time) > 0 And DateDiff("n", "03:15:00", Time) < 15 Then
        
            '��ֹͣtmIdle���������ط��Ĺ����У�����δ���
            tmIdle.Enabled = False
            mlngIdleTime = 1
            

            '�ж���־�ļ��Ƿ񳬹�600M�������򴴽��µ���־�ļ�
            If FileLen(gstrAccessName) > 600000000 Then
                Call subNewLogFile
            End If
        
            '�ط���Ϣ
            Call funReSendMessage
            
        End If
    End If
    
    '����tmIdle
    tmIdle.Enabled = True
    
    
'-----------------------------���·�ʽ�޷�ʹ��-------------------------------
'����GE��MUSEϵͳ���Զ�����socket���ӣ���������ͨѶ�Ŀ���ʱ�䣬Ҳ��һֱ����socket��
'��˲�����socket����=1���ж�socket�Ƿ���С�

'    '�ۼ�SocketͨѶ�Ŀ���ʱ�䣬����ʱ�乻�����������ط���Ϣ�Ļ���
'    'wsHL7Link����Ϊ1����ʾû�������κ�Socket��������Ϣ���ǿ���
'    If wsHL7Link.Count = 1 Then
'        mlngIdleTime = mlngIdleTime + 1
'        '����ʮ����֮���ط�һ����Ϣ
'        If mlngIdleTime > 600 Then
'            mlngIdleTime = 1
'
'            '�ط���Ϣ
'            Call funReSendMessage
'
'            '�ж���־�ļ��Ƿ񳬹�600M�������򴴽��µ���־�ļ�
'            If FileLen(gstrAccessName) > 600000000 Then
'                Call subNewLogFile
'            End If
'
'        End If
'    End If
'---------------------------------------------------------------------------
    Exit Sub
err:
    Call WriteLog(5001, err.Number, "tmIdle_Timer ��ʱ��������־���ط���Ϣ�������������ǣ�" & err.Description)
    tmIdle.Enabled = True
End Sub

Private Sub tmLinkTimeOut_Timer(Index As Integer)
    
    On Error GoTo err
    
    '�����ʱ���Լ���Ϣ��ʱ,����ACK��Ӧ
    If Index <> 0 Then
        gintTimeOut(Index) = gintTimeOut(Index) + 1
        '��ʱ
        If gintTimeOut(Index) > gintTimeOutMax Then
            Dim strACK As String
            
            Call GetMsgACK(tmLinkTimeOut(Index).Tag, strACK)
                            
            '����ACK��Ӧ
            wsHL7Link(Index).SendData strACK
            
            '�رպ�ж����Ϣ��ʱ��
            tmLinkTimeOut(Index).Enabled = False
            Unload tmLinkTimeOut(Index)
            
            gintTimeOut(Index) = 0
            '�ر���Ϣ����
            wsHL7Link(Index).Close
            Unload wsHL7Link(Index)
        End If
    End If
    
    Exit Sub
err:
    Call WriteLog(5002, err.Number, "tmLinkTimeOut_Timer ��ʱ�����ִ���Index=" & Index & "�����������ǣ�" & err.Description)
End Sub

Private Sub tmMsgProcess_Timer()
    On Error GoTo err
    
    '��ֹͣTimer��ʱ������ʱ�䴦����Ϣ
    tmMsgProcess.Enabled = False
    Call funMsgProcess
    
    '����ϵͳ����
    '���wsHL7Linkֻ��һ��������������intWSLinkCount����
    If wsHL7Link.Count = 1 Then
        intWSLinkCount = 0
        ReDim gintTimeOut(0) As Integer
    End If
    
    '��Ϣ������֮����������ʱ
    tmMsgProcess.Enabled = True
    Exit Sub
err:
    Call WriteLog(5004, err.Number, "tmMsgProcess_Timer ��ʱ�����ִ��󣬴��������ǣ�" & err.Description)
    tmMsgProcess.Enabled = True
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub

Private Sub wsHL7Link_Close(Index As Integer)

    On Error Resume Next
    
    'ͬʱж�س�ʱ��ʱ��
    Unload tmLinkTimeOut(Index)
    gintTimeOut(Index) = 0
    
'    wsHL7Link(Index).Close
    '�ڶԷ��ر����ӵ�ʱ�򣬰��Լ�ж����
    Unload wsHL7Link(Index)
End Sub

Private Sub wsHL7Link_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    Dim strACK As String
    Dim strData As String
    Dim lngMsgFullType As Long
    Dim strMessage As String
    Dim strRemain As String
    
    On Error GoTo err
    
    '��¼��ϸ��־
    Call WriteProcessLog("wsHL7Link_DataArrival", "���յ���Ϣ��׼������", "wsHL7Link(" & Index & ")���յ���Ϣ���ݣ�׼������", 2)
    
    '��������
    wsHL7Link(Index).GetData strData, vbString
    
    '��¼ͨѶ��־
    Call WriteProcessLog("wsHL7Link_DataArrival", "wsHL7Link(" & Index & ")���յ���Ϣ���ݣ�", strData, 1)
    
    '���յ���һ�����ݣ��ſ�ʼ��ʱ��ͬʱ��ʼ��tag
    If tmLinkTimeOut(Index).Enabled = False Then
        tmLinkTimeOut(Index).Enabled = True
        gintTimeOut(Index) = 1
        tmLinkTimeOut(Index).Tag = ""
    End If
    
    '�ۼ���Ϣ��tag��
    tmLinkTimeOut(Index).Tag = tmLinkTimeOut(Index).Tag & strData
    
    '��ȡ�͸���TAG�е���Ϣ�������������Ϣ���򷵻�ACK�����
     
    While funGetMessage(tmLinkTimeOut(Index).Tag, strMessage, strRemain) = True
        '����ACK��'��Ϣ���
        If GetMsgACK(strMessage, strACK) = True Then
            'ֻ����Ϣ��ȷ���ű������ݵ�HL7��Ϣ����������
            '��¼��Ϣ��־
            Call WriteMessageLog("wsHL7Link(" & Index & ")���յ��ĵ�����Ϣ", strMessage)
            Call MsgInQueue(strMessage)
        End If
        '������Ϣ�Ƿ���ȷ����������Ϣ��������Ӧ
        wsHL7Link(Index).SendData strACK
        '��¼��Ϣ��־
        Call WriteMessageLog("wsHL7Link(" & Index & ")������ϢACK", strACK)
        
        '����tag������
        tmLinkTimeOut(Index).Tag = strRemain
    Wend
    
    Exit Sub
err:
    '�ݲ�����
    Call WriteLog(3001, err.Number, "wsHL7Link(" & Index & ") wsHL7Link_DataArrival ���ִ��󣬴��������ǣ�" & err.Description)
End Sub


Private Sub wsHL7Server_Close(Index As Integer)
    '���Ҳ�ɾ������������´�������������ʵ��
    'wsHL7Link
    '�����
End Sub

Private Sub wsHL7Server_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    '���յ�ͨѶ��������һ���µ�ʵ��������������󣬲��Ҵ���һ��Timer����¼����ʱ��
    
    '���ü������winsock����������Ϊÿһ������ʵ�����Ǵ������������ӵĹرն��رյ�,ʵ��������������
    intWSLinkCount = intWSLinkCount + 1
    Load wsHL7Link(intWSLinkCount)
    
    '���س�ʱ�ͷְ��������
    Load tmLinkTimeOut(intWSLinkCount)
    
    tmLinkTimeOut(intWSLinkCount).Interval = 1000   '1������һ��
    ReDim Preserve gintTimeOut(intWSLinkCount) As Integer
    
    wsHL7Link(intWSLinkCount).LocalPort = 0
    wsHL7Link(intWSLinkCount).Accept requestID
    
    'ͨѶ��־
    Call WriteProcessLog("wsHL7Server_ConnectionRequest", "���յ���������", _
        "�Ӷ˿ڣ�" & wsHL7Server(Index).LocalPort & "�����յ��������������IP�ǣ�" & wsHL7Server(Index).RemoteHostIP, 1)
        
        '��¼��ϸ��־
    Call WriteProcessLog("wsHL7Server_ConnectionRequest", "������socket-" & intWSLinkCount & "������Ϣ", "wsHL7Link(" & intWSLinkCount & ")״̬Ϊ��" & wsHL7Link(intWSLinkCount).State, 2)
End Sub


Public Function funActionProcess() As Long
'-----------------------------------------------------------------------------
'����:����HIS�Ķ�������ѯHL7�м��һ��������������Ϣ�������ֹͣ������źţ�����ֹͣ��Ϣ�ķ���
'������
'����ֵ����
'-----------------------------------------------------------------------------
     '��ʱ��ѯHIS���ݿ��HL7��Ϣ��ʱ��
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnֱ�ӷ��� As Boolean
    Dim str�������� As String
    Dim strҵ��ID�� As String
    Dim lng��ϢID As Long
    Dim lngSendResult As Long
    
    On Error GoTo err
    
    '�����е㲻�ԣ�һ��������Ϣ����Ӧ������Է��͵���Ϣ���ã�ÿ����Ϣ���ö����ͳɹ��ˣ������ǳɹ���
    
    Call CheckDBConnect
    
    'ֻ����3��֮�ڵ�ҽ�������һ���շѵ���ҽ��3��֮��û�нɷѣ����ٷ���
    strSQL = "Select Id ,��������,ҵ��ID��,���ʹ���,ֱ�ӷ��� From zlhis.hl7������Ϣ Where ����ʱ�� >= Sysdate -3 order by ID"
    Set rsTemp = gzlDatabase.OpenSQLRecord(strSQL, "��ȡ�����͵�HL7��Ϣ")
    
    '��¼��ϸ��־
    Call WriteProcessLog("funActionProcess", "1����ȡ�����͵�HL7��Ϣ", "strSQL = " + Replace(strSQL, "'", "��"), 3)
        
    While Not rsTemp.EOF And gblnServiceStart = True
        
        
        
        'һ������֯��Ϣ���ҷ���
        blnֱ�ӷ��� = IIf(Nvl(rsTemp!ֱ�ӷ���, 0) = 1, True, False)
        str�������� = Nvl(rsTemp!��������)
        strҵ��ID�� = Nvl(rsTemp!ҵ��ID��, 0)
        lng��ϢID = rsTemp!ID
        
        '��¼��ϸ��־
        Call WriteProcessLog("funActionProcess", "2��׼��������Ϣ", "lng��ϢID=" & lng��ϢID & ",strҵ��ID��=" & strҵ��ID��, 3)
    
        lngSendResult = funSendAMessage(lng��ϢID, strҵ��ID��, str��������, blnֱ�ӷ���)
        
        If lngSendResult <> 2 Then
            Call WriteProcessLog("funActionProcess", "3����Ϣ���ͽ����" & IIf(lngSendResult = 0, "�ɹ�", "ʧ��"), "��ϢID��" & lng��ϢID & "��ҵ��ID��" & strҵ��ID�� & "���������ͣ�" & str��������, 3)
            
            '����з���ʧ�ܵģ��ȼ�¼���������ͳɹ��ģ�ɾ�����ݿ���ʱ���¼
            strSQL = "zlhis.b_Hl7interface.HL7������Ϣ_UPDATE(" & lng��ϢID & "," & IIf(lngSendResult = 0, "0", "1") & " ) "
            gzlDatabase.ExecuteProcedure strSQL, "��Ϣ���ͺ���"
        End If
        
        rsTemp.MoveNext
    Wend
    Exit Function
err:
    Call WriteLog(3002, err.Number, "funActionProcess ���ִ��󣬴��������ǣ�" & err.Description)
End Function

Private Sub wsSend_Connect()
    '����¼�������˵�����ӳɹ������Է�����Ϣ��
    mblnSendConnect = True
End Sub

Private Sub wsSend_DataArrival(ByVal bytesTotal As Long)
    '������Ϣ֮�󣬶Է�������Ӧ��Ϣ
    Dim strData As String
    
    On Error GoTo err
    
    '������Ϣ
    wsSend.GetData strData, vbString
    '����Ϣ�ݴ���TAG��
    wsSend.Tag = strData
    mblnSendACK = True
    Exit Sub
err:
    Call WriteLog(3003, err.Number, "wsSend_DataArrival ���ִ��󣬴��������ǣ�" & err.Description)
End Sub


Private Function funSendHL7Message(strIP As String, lngPort As Long, strMessage As String) As Long
'-----------------------------------------------------------------------------
'����:����һ��HL7��Ϣ
'������ strIP -- IP��ַ
'       lngPort -- �˿ں�
'       strMessage -- ��Ϣ����
'����ֵ��0 -- �ɹ��� 1 - ʧ�ܣ�2-ֹͣ����
'-----------------------------------------------------------------------------
    Dim dblSleepTime As Double
    
    On Error GoTo err
                    
    funSendHL7Message = 1
    
    '�ȶϿ����͵����ӣ�����������
    If wsSend.State <> sckClosed Then wsSend.Close
    mblnSendConnect = False
    
    wsSend.RemoteHost = strIP
    wsSend.RemotePort = lngPort
    wsSend.LocalPort = 0    ''�ͻ��˵ı��ض˿�ָ��Ϊ0�������ӵ�ʱ�򣬻��Զ�ʹ�ÿ��д��ڣ�������Ϊ�Ͽ�֮����Ҫ�ȴ�5���Ӳ����ٴ�����
    wsSend.Connect
    If wsSend.State = sckConnecting Then
        Call WriteProcessLog("funSendHL7Message", "1��׼��������Ϣ����������", "IP��ַ�ǣ�" & strIP & "���˿��ǣ�" & lngPort & "��׼�����͵���Ϣ�ǣ�" & strMessage, 3)
    End If

    ''һ����ʱ,����doevents���ܴ���Socket��connect�¼�
    dblSleepTime = timeGetTime
    'ֻ��ʱ�䵽�ˣ����߷��ͳɹ��ˣ������˳������ˣ����˳�ѭ��
    While timeGetTime < dblSleepTime + 2000 And mblnSendConnect = False And gblnServiceStart = True
        DoEvents
    Wend
    
    If mblnSendConnect = True Then
        '���ӳɹ���������Ϣ
        wsSend.SendData strMessage
        Call WriteProcessLog("funSendHL7Message", "2������HL7��Ϣ", "IP��ַ�ǣ�" & strIP & "���˿��ǣ�" & lngPort & "��׼�����͵���Ϣ�ǣ�" & strMessage, 1)
        
        '������Ϣ֮�󣬵ȴ�ACK��Ӧ
        mblnSendACK = False
        '����һ����ʱ���ȴ���Ӧ��Ϣ
        dblSleepTime = timeGetTime
        'ֻ��ʱ�䵽�ˣ������յ�ACK�ɹ��ˣ������˳������ˣ����˳�ѭ��
        While timeGetTime < dblSleepTime + 2000 And mblnSendACK = False And gblnServiceStart = True
            DoEvents
        Wend
        
        '������Ӧ��Ϣ�������AA���գ����¼��Ϣ���ͳɹ�
        If mblnSendACK = True Then
            '���յ�ACK��Ӧ�����������Ӧ
            Call WriteProcessLog("funSendHL7Message", "3���յ�ACK��Ӧ", "ACK��Ϣ�����ǣ�" & wsSend.Tag, 1)
            If funParseACK(strMessage, wsSend.Tag) = 0 Then
                '�������գ�����0
                funSendHL7Message = 0
            Else
                funSendHL7Message = 1
            End If
        Else
            'û���յ�ACK��Ӧ��funSendHL7Message����Ĭ��ֵ1������ʧ��
            Call WriteProcessLog("funSendHL7Message", "3.1������ACK��Ӧ��ʱ", "����ACK��Ӧ��ʱ����Ϣ����ʧ��", 1)
        End If
    Else
        '���Ӳ��ɹ����Ǿ����ǳ�ʱ�ˣ���¼��־
        Call WriteProcessLog("funSendHL7Message", "2.1�����ӳ�ʱ", "IP��ַ�ǣ�" & strIP & "���˿��ǣ�" & lngPort & "��׼�����͵���Ϣ�ǣ�" & strMessage, 1)
    End If
            
    If gblnServiceStart = False Then
        '�����;�����û���Ԥ��ֹͣ�˷��ͷ����򷵻�2
        funSendHL7Message = 2
        Call WriteProcessLog("funSendHL7Message", "4���ֹ�ֹͣ���ͷ���", "��Ϣ����û�з��ͳɹ�", 1)
    End If
        
    '��Ϣ������ɺ��ǲ���Ҫ�Ͽ�wsSend�������أ�
    wsSend.Close
    
    Exit Function
err:
    Call WriteLog(3004, err.Number, "funSendHL7Message ���ִ��󣬴��������ǣ�" & err.Description)
End Function

Private Sub subShowLog(intType As Integer)
    On Error Resume Next
    
    Set mfrmShowLog = New frmShowLog
    mfrmShowLog.intLogType = intType
    mfrmShowLog.Show 1, Me
    Set mfrmShowLog = Nothing
    
End Sub

Private Function funReSendMessage() As Long
'-----------------------------------------------------------------------------
'����:  ���·���HL7��Ϣ������hl7�ط���Ϣ�����е���Ϣ��һ�δ���10����Ϣ
'       �����ֹͣ������źţ�����ֹͣ��Ϣ�ķ���
'������
'����ֵ����
'-----------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnֱ�ӷ��� As Boolean
    Dim str�������� As String
    Dim strҵ��ID�� As String
    Dim lng��ϢID As Long
    Dim lngSendResult As Long
    
    On Error GoTo err
    
    Call WriteProcessLog("funReSendMessage", "1���ط���Ϣ", "׼���ط���Ϣ", 3)
    
    '��ѯ��hl7�ط���Ϣ����
    strSQL = "Select ID,��������,ҵ��ID��,����ʱ��,���ʹ���,ֱ�ӷ���,�ط�ʱ��,��ϢʱЧ From zlhis.hl7�ط���Ϣ  Where  ��ϢʱЧ = 1  Order By �ط�ʱ�� "
    Set rsTemp = gzlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ҫ�ط���ȫ��HL7��Ϣ")
    
    'ѭ��������Ϣ
    While Not rsTemp.EOF And gblnServiceStart = True
        'һ������֯��Ϣ���ҷ���
        blnֱ�ӷ��� = IIf(Nvl(rsTemp!ֱ�ӷ���, 0) = 1, True, False)
        str�������� = Nvl(rsTemp!��������)
        strҵ��ID�� = Nvl(rsTemp!ҵ��ID��, 0)
        lng��ϢID = rsTemp!ID
        
        lngSendResult = funSendAMessage(lng��ϢID, strҵ��ID��, str��������, blnֱ�ӷ���)
        
        If lngSendResult <> 2 Then
            
            Call WriteProcessLog("funReSendMessage", "2���ط���Ϣ�����ͽ����" & IIf(lngSendResult = 0, "�ɹ�", "ʧ��"), "��ϢID��" & lng��ϢID & "��ҵ��ID��" & strҵ��ID�� & "���������ͣ�" & str��������, 3)
            
            '����з���ʧ�ܵģ��ȼ�¼���������ͳɹ��ģ�ɾ�����ݿ���ʱ���¼
            strSQL = "zlhis.b_Hl7interface.HL7�ط���Ϣ_UPDATE(" & lng��ϢID & "," & IIf(lngSendResult = 0, "0", "1") & " ) "
            gzlDatabase.ExecuteProcedure strSQL, "��Ϣ�ط�����"
            
        Else
            '��Ϣû�з��ͣ�������ҽ��δ�Ʒ�
            Call WriteProcessLog("funReSendMessage", "2.2����Ϣû�з���", "������ҽ��û�мƷ�", 3)
        End If
        
        rsTemp.MoveNext
    Wend
    Exit Function
err:
    Call WriteLog(3005, err.Number, "funReSendMessage ���ִ��󣬴��������ǣ�" & err.Description)
End Function

Private Function funSendAMessage(lng��ϢID As Long, strҵ��ID�� As String, str�������� As String, blnֱ�ӷ��� As Boolean) As Long
'-----------------------------------------------------------------------------
'����:  ����һ��HL7��Ϣ
'������ lng��ϢID   ---��ϢID
'       strҵ��ID�� ---��Ϣ��ҵ��ID��
'       str�������� ---��Ϣ�Ķ�������
'       blnֱ�ӷ��� ---�Ƿ�ֱ�ӷ�����Ϣ
'����ֵ��0--���ͳɹ���1--����ʧ�ܣ�2--û�з��ͣ�����δ����
'-----------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsFee As ADODB.Recordset
    Dim arrMsgs As THl7Messages
    Dim iMsg As Integer
    Dim lngResult As Long
    Dim blnSendOK As Boolean
    
    On Error GoTo err
    
    'ͨ��ֱ�ӷ����ֶ��ж���Ϣ�Ƿ�ֱ�ӷ��͸�MUSE������ͼ���շ�״̬
    If blnֱ�ӷ��� = False And str�������� = HL7_MSG_SEND_NEW_ORDER Then
        '�¿�ҽ�������Ҳ���ֱ�ӷ��ͣ����ѯ���ݿ⣬��ȡҽ�����շ����
        strSQL = "Select ҽ��ID From zlhis.����ҽ������  " & _
                 " Where ��¼����=1 And �Ʒ�״̬ In (1,2) And ҽ��id = [1] And ���ͺ� = [2]"
        Set rsFee = gzlDatabase.OpenSQLRecord(strSQL, "��ȡҽ���ļƷ�״̬", CLng(Split(strҵ��ID��, ";")(0)), CLng(Split(strҵ��ID��, ";")(1)))
        
        If rsFee.RecordCount = 0 Then
            '���շ�,���Է���
            blnֱ�ӷ��� = True
        End If
    End If
    
    If blnֱ�ӷ��� = True Then
    
        '��¼��ϸ��־
        Call WriteProcessLog("funSendAMessage", "1��׼����ȡ��Ϣ����", "str��������=" & str��������, 3)
        
        '�����ݿ������ö�����HL7��Ϣ�ṹ
        '�����ݿ���ȡ��Ϣ����
        arrMsgs = getMsgDefFromDB(str��������)
        
        '��ѯ����дÿһ����Ϣ�ε��ֶ�����
        If UBound(arrMsgs.arrMsgs) > 0 Then
            '��¼������Ϣ��ID
            arrMsgs.lngActionID = lng��ϢID
            '����Ƕ�η��͵���Ϣ�����ƶ����Ϣ��¼
'                If Nvl(rsTemp!���ʹ���, 1) > 1 Then
'                    Call funDuplicateMsg(arrMsgs, CInt(Nvl(rsTemp!���ʹ���, 1)))
'                End If
            
            '��¼��ϸ��־
            Call WriteProcessLog("funSendAMessage", "2��׼�������Ϣ����", "strҵ��ID��=" & strҵ��ID��, 3)
        
            '�����ֶ�������д���ݿ���߹̶�����
            Call funfillMsgValue(arrMsgs, strҵ��ID��)
            
            '����ÿһ����Ϣ
            For iMsg = 1 To UBound(arrMsgs.arrMsgs)
                lngResult = funSendHL7Message(arrMsgs.arrMsgs(iMsg).strIP, arrMsgs.arrMsgs(iMsg).lngPort, arrMsgs.arrMsgs(iMsg).strText)
                If lngResult = 0 Then
                    arrMsgs.arrMsgs(iMsg).blnSendOK = True
                    '��¼��Ϣ��¼
                    Call WriteMessageLog(arrMsgs.arrMsgs(iMsg).strActionType, arrMsgs.arrMsgs(iMsg).strText)
                ElseIf lngResult = 2 Then   'ֹͣ���ͷ���
                    Exit For
                End If
            Next iMsg
            
            '����������Ϣ�Ƿ��ͳɹ���ͨ�����̴��������Ϣ��¼
            blnSendOK = True
            '���ÿһ�������õ���Ϣ�Ƿ��ͳɹ�
            For iMsg = 1 To UBound(arrMsgs.arrMsgs)
                If arrMsgs.arrMsgs(iMsg).blnSendOK = False Then
                    blnSendOK = False
                    Exit For
                End If
            Next iMsg
            
        Else
            blnSendOK = False
        End If
        funSendAMessage = IIf(blnSendOK = True, 0, 1)
    Else
        funSendAMessage = 2 '��Ϣû�з���
    End If
    
    Exit Function
err:
    Call WriteLog(3006, err.Number, "funSendAMessage ���ִ��󣬴��������ǣ�" & err.Description)
End Function
