VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Begin VB.Form frmPetitionCapture 
   Caption         =   "���뵥ͼ��"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13380
   Icon            =   "frmPetitionCapture.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   13380
   StartUpPosition =   3  '����ȱʡ
   Begin ScanLibCtl.ImgScan ImageScanner 
      Left            =   3360
      Top             =   0
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   0
      StopScanBox     =   -1  'True
      FileType        =   3
      CompressionType =   0
      CompressionInfo =   0
      ScanTo          =   4
   End
   Begin VB.Frame fmeDcmViewer 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   5520
      TabIndex        =   6
      Top             =   240
      Width           =   4215
      Begin DicomObjects.DicomViewer dcmViewImg 
         Height          =   1575
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
         _Version        =   262147
         _ExtentX        =   3836
         _ExtentY        =   2778
         _StockProps     =   35
         BackColor       =   -2147483640
         UseScrollBars   =   0   'False
         UseMouseWheel   =   -1  'True
      End
      Begin VB.PictureBox picTemp2 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   240
         ScaleHeight     =   1215
         ScaleWidth      =   1695
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         Caption         =   "δ�ҵ�����ͼ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   20
         Top             =   3120
         Width           =   3840
      End
      Begin VB.Image img 
         Height          =   1785
         Left            =   1080
         Picture         =   "frmPetitionCapture.frx":058A
         Top             =   960
         Width           =   2505
      End
   End
   Begin VB.Frame fmeInfoCtrl 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   3210
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.Frame fmePatientInfo 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2535
         Begin VB.Label labSpePosition 
            AutoSize        =   -1  'True
            Caption         =   "labSpePosition"
            Height          =   180
            Left            =   1080
            TabIndex        =   18
            Top             =   2520
            Width           =   1260
         End
         Begin VB.Label labCheckMethod 
            AutoSize        =   -1  'True
            Caption         =   "labCheckWay"
            Height          =   180
            Left            =   1080
            TabIndex        =   17
            Top             =   2160
            Width           =   990
         End
         Begin VB.Label labAge 
            AutoSize        =   -1  'True
            Caption         =   "labAge"
            Height          =   180
            Left            =   1080
            TabIndex        =   16
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label labSex 
            AutoSize        =   -1  'True
            Caption         =   "labSex"
            Height          =   180
            Left            =   1080
            TabIndex        =   15
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label labRoom 
            AutoSize        =   -1  'True
            Caption         =   "labRoom"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   1080
            TabIndex        =   14
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label labNo 
            AutoSize        =   -1  'True
            Caption         =   "labNo"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   1080
            TabIndex        =   13
            Top             =   720
            Width           =   525
         End
         Begin VB.Label labName 
            AutoSize        =   -1  'True
            Caption         =   "labName"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   1080
            TabIndex        =   12
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblCheckNum 
            AutoSize        =   -1  'True
            Caption         =   "�� �� ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   900
         End
         Begin VB.Label lblPatientAge 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��    ��:"
            Height          =   180
            Left            =   120
            TabIndex        =   10
            Top             =   1800
            Width           =   930
         End
         Begin VB.Label lblPatientDept 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���˿���:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   885
         End
         Begin VB.Label lblPatientName 
            AutoSize        =   -1  'True
            Caption         =   "��    ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   915
         End
         Begin VB.Label lblExamineMethod 
            AutoSize        =   -1  'True
            Caption         =   "��鷽��:"
            Height          =   180
            Left            =   120
            TabIndex        =   4
            Top             =   2160
            Width           =   810
         End
         Begin VB.Label lblSpePosition 
            AutoSize        =   -1  'True
            Caption         =   "��鲿λ:"
            Height          =   180
            Left            =   120
            TabIndex        =   3
            Top             =   2520
            Width           =   810
         End
         Begin VB.Label lblPatientSex 
            AutoSize        =   -1  'True
            Caption         =   "��    ��:"
            Height          =   180
            Left            =   120
            TabIndex        =   2
            Top             =   1440
            Width           =   810
         End
      End
   End
   Begin DicomObjects.DicomViewer dcmMiniature 
      Height          =   4935
      Left            =   240
      TabIndex        =   19
      Top             =   3600
      Width           =   4050
      _Version        =   262147
      _ExtentX        =   7144
      _ExtentY        =   8705
      _StockProps     =   35
      BackColor       =   -2147483642
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   2880
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPetitionCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'����������
Private Type TPoint
  X As Integer
  Y As Integer
End Type

'��Ƶ��������
Private Enum TVideoDriverType
  vdtWDM = 0
  vdtVFW = 1
  vdtTWAIN = 2
  '������Ҫ֧�ֵ���������......
End Enum

Private mstrTempDirOfScan As String          'ɨ�����ʱĿ¼
Private mstrScanDeviceTempDir As String      'ɨ���豸��ʱĿ¼
Private mstrBufferDir As String

Private mintScanImageIndex As Integer        'ɨ��ͼ������
Private mintCurImgIndex As Integer           '��ǰ��ѡ�е�ͼ������


Private Const CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE As String = "Scan"  'Ĭ��ɨ���ļ���ʱ�洢·��
Private Const CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME As String = "\TempScan"  'Ĭ��ɨ���ļ���ʱ�洢·��

Private mlngAdviceID As Long           'ҽ��ID
Private mlngCurDeptId As Long          '��ǰ����ID
Private mstrPrivs As String            '��ǰȨ��

Private mstrSaveDeviceID As String      '�洢�豸���豸��
Private miNet As New clsFtp             'FTP��
Private mFtpUser As String              'FTP�˺�
Private mFtpPass As String              'FTP����
Private mFtpDir As String               'FTPĿ¼
Private mFtpIp As String               'FTP��ַ

Private mlngBaseX As Long               'dcmView�����Downʱ��X����
Private mlngBaseY As Long               'dcmView�����Downʱ��Y����
Private mMouseDownPoint As TPoint       '�����DcmImage�ϰ���ʱ��λ��
Private mblndcmViewImgDown As Boolean    '�����ж�dcmView������Ƿ񱻰���
Private mInitScrollPoint As TPoint      '��ʼ�϶�ʱ�ĳ�ʼλ��
Private mCorpSize As TPoint             '�϶�������ƫ��λ��
Private mblnIsExamine As Boolean        '�Ƿ�鿴���뵥
Public mblnIsLogin As Boolean           '�Ƿ��ǵ�¼���ڵ����뵥��ť

'���˻�����Ϣ
Private mstrCheckNo As String           '����
Private mstrDeptName As String          '��������
Private mstrPatientName As String       '��������
Private mstrPatientAge As String        '��������
Private mstrPatientSex As String        '�����Ա�
Private mstrExamineMethod As String     '��鷽��
Private mstrSpePosition As String       '�걾��λ

'�˵�
Private Enum conMenus
    conMenu_Process_RRotate = 503
    conMenu_Process_LRotate = 504
    conMenu_Process_Magnify = 502
    conMenu_Process_Shrink = 513
    conMenu_Process_Restore = 8124
    conMenu_Process_ScamImg = 8101
    conMenu_Process_DeleteImg = 10001
    conMenu_Process_ScanSet = 815
    conMenu_Process_ChoiceEqu = 181
    conMenu_File_Exit = 191
End Enum
Private mblnImgProcess As Boolean       '�Ƿ��ڶ�ѡ��ͼ����зŴ�ȴ���
Private mblnOperate As Boolean          '�Ƿ����ͼ��ɨ��Ȳ���
Private mdcmTmpView As DicomViewer
Private mintImageType As Integer        'ɨ��ͼ���ʽ

Public Event RefreshState(ByVal blnState As Long)             'ˢ�¼���б����뵥����״̬

Public Sub ShowPetitionCaptureWind(ByVal strPrivs As String, lngCurDeptId As Long, strDeptName As String, _
                                   strPatientName As String, strPatientAge As String, strPatientSex As String, _
                                   strExamineMethod As String, strSpePosition As String, blnIsExamine As Boolean, _
                                   blnIsLogin As Boolean, Optional lngAdviceID As Long = 0, Optional intState As Integer = 0, _
                                   Optional dcmTmpView As DicomViewer)
Dim rsTemp As ADODB.Recordset
Dim strSql As String
Dim FTPconn As New clsFtp
On Error GoTo errH

    '����ģ�����
    mstrPrivs = strPrivs
    mlngAdviceID = lngAdviceID
    mblnIsExamine = IIf(intState = 0, blnIsExamine, True)
    mblnIsLogin = blnIsLogin
    mlngCurDeptId = lngCurDeptId
    Set mdcmTmpView = dcmTmpView
    
    '��ʼ�����Ҽ�����
    InitDeptPara mlngCurDeptId
    
    If FTPconn.FuncFtpConnect(mFtpIp, mFtpUser, mFtpPass) = 0 Then
        MsgBox "FTP�����������ӣ������������á�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ͽ�FTP��������
    FTPconn.FuncFtpDisConnect
    
    strSql = "select ���� from Ӱ�����¼  where ҽ��id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "ȡ�ü���", lngAdviceID)
    
    If rsTemp.RecordCount > 0 Then
        mstrCheckNo = Nvl(rsTemp!����)
    End If
    
    mstrDeptName = strDeptName
    mstrPatientName = strPatientName
    mstrPatientAge = strPatientAge
    mstrPatientSex = strPatientSex
    mstrExamineMethod = strExamineMethod
    mstrSpePosition = strSpePosition
    
    If Not mblnIsExamine Then mblnOperate = True
    
    '��ʼ��������Ϣ
    Call InitLables
     
    Call Me.Show(1)
    
    Exit Sub
errH:
    '�Ͽ�FTP����
    FTPconn.FuncFtpDisConnect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 32, 32
    End With
    
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    'ͼ���������������
    Set cbrToolBar = Me.cbrMain.Add("ͼ�������", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True '�ı���ʾ��ͼ���·�
    cbrToolBar.Closeable = False
    
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RRotate, "˳ʱ"): cbrControl.ToolTipText = "˳ʱ����ת90��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_LRotate, "��ʱ"): cbrControl.ToolTipText = "��ʱ����ת90��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Magnify, "�Ŵ�"): cbrControl.ToolTipText = "�Ŵ�ͼ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Shrink, "��С"): cbrControl.ToolTipText = "��Сͼ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Restore, "�ָ�"): cbrControl.ToolTipText = "�ָ�ͼ�񵽳�ʼ״̬"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_ScamImg, "ɨ��ͼ��") '102
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_DeleteImg, "ɾ��ͼ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_ScanSet, "ɨ������") '181
        'Set cbrControl = .Add(xtpControlButton, conMenu_Process_ChoiceEqu, "ѡ���豸")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
         cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    cbrToolBar.Position = xtpBarTop
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHandle
    
    Select Case control.ID
        Case conMenu_Process_RRotate        '˳ʱ
            Call subSetRotate(True)
            
        Case conMenu_Process_LRotate        '��ʱ
            Call subSetRotate(False)
            
        Case conMenu_Process_Magnify        '�Ŵ�
            Call cmdMagnify_Click
            
        Case conMenu_Process_Shrink         '��С
            Call cmdReduce_Click
        
        Case conMenu_Process_Restore        '�ָ�
            Call cmdReset_Click
        
        Case conMenu_Process_ScamImg        'ɨ��ͼ��
            Call cmdScanImg_Click
        
        Case conMenu_Process_DeleteImg      'ɾ��ͼ��
            Call cmdDeleteImg_Click
        
        Case conMenu_Process_ScanSet        'ɨ������
            Call cmdScanSet_Click
        
'        Case conMenu_Process_ChoiceEqu      'ѡ���豸
'            Call cmdChoiceEqu_Click
        
        Case conMenu_File_Exit              '�˳�
            Call Menu_File_Exit
            
    End Select
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHandle
    
    Select Case control.ID
        Case conMenu_Process_RRotate, conMenu_Process_LRotate, conMenu_Process_Magnify, conMenu_Process_Shrink, _
             conMenu_Process_Restore    '˳ʱ,��ʱ,�Ŵ�,��С,�ָ�
            
            control.Enabled = mblnImgProcess
        
        Case conMenu_Process_ScamImg, conMenu_Process_ScanSet
            'ɨ��ͼ��,ɾ��ͼ��,ɨ������
            control.Visible = mblnOperate
            control.Enabled = mblnOperate
            
        Case conMenu_Process_DeleteImg
            control.Visible = mblnOperate
            control.Enabled = mblnOperate And Not mblnIsLogin
    End Select
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_File_Exit()
    Unload Me
End Sub

Private Sub subSetRotate(blnClockwise As Boolean)
'------------------------------------------------
'���ܣ�dcmViewImg��ͼ�����ת
'������blnClockwise��ת�ķ���,True=˳ʱ����ת��False=��ʱ����ת
'���أ��ޣ�ֱ�Ӵ���dcmView�е�ͼ��
'------------------------------------------------
    If dcmViewImg.Images.Count > 0 Then
        Dim iRotateState As Integer
        
        iRotateState = dcmViewImg.Images(1).RotateState
        If blnClockwise = True Then
            iRotateState = iRotateState - 1
        Else
            iRotateState = iRotateState + 1
        End If
        If iRotateState = -1 Then iRotateState = 3
        iRotateState = iRotateState Mod 4
        dcmViewImg.Images(1).RotateState = iRotateState
    End If
End Sub

Private Sub cmdDeleteImg_Click()
On Error GoTo errHandle

    'ɾ������
    If mblnIsLogin Then
        Call subDeleteDcmImage
    Else
        Call subDeleteFTPImage
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdScanImg_Click()
On Error GoTo errHandle
    
    Call ScanImages
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdScanSet_Click()
On Error GoTo errHandle
    '��ɨ�����ô���
    Call frmScanSetup.ShowParameterConfig(ImageScanner, Me)
    mintImageType = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPetitionCapture", "ɨ��ͼ���ʽ", 0))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ScanImages()
    Dim strScanFile As String
    
    'ɾ����������ʱ�洢��ͼ��Ŀ¼
    On Error GoTo continue
    If Dir(mstrTempDirOfScan, vbDirectory) <> "" Then
        Call mdlDir.DeleteFolder(mstrTempDirOfScan, , False)
    End If
continue:

    If Dir(mstrTempDirOfScan, vbDirectory) = "" Then
        Call MkDir(mstrTempDirOfScan)
    End If

    'ɾ��twain�豸��ʱ�洢��Ŀ¼
    On Error GoTo continue1
    If Dir(mstrScanDeviceTempDir, vbDirectory) <> "" Then
        Call mdlDir.DeleteFolder(mstrScanDeviceTempDir, , False)
    End If
continue1:

    If Dir(mstrScanDeviceTempDir, vbDirectory) = "" Then
        Call MkDir(mstrScanDeviceTempDir)
    End If
    
    mintScanImageIndex = 0
  
    If Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPetitionCapture", "ɨ����������", 0)) = vdtWDM Then
        On Error GoTo errProcess
        
        strScanFile = mstrTempDirOfScan & "\" & CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE & strScanFile & ".bmp"
    
        Call frmScanSetup.ShowScanWind(strScanFile, Me)
        
        Exit Sub
    End If

    '����ɨ�����ļ���������
    ImageScanner.FileType = IIf(mintImageType = 0, BMP_Bitmap, JPG_File)
    ImageScanner.StopScanBox = True
    ImageScanner.ShowSetupBeforeScan = True
    ImageScanner.ScanTo = UseFileTemplateOnly
    '���òɼ���ģ���ļ�
    ImageScanner.Image = mstrTempDirOfScan & "\" & CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE


    If Not ImageScanner.ScannerAvailable Then
        ImageScanner.OpenScanner
    End If

    On Error GoTo errProcess
    Call ImageScanner.StartScan
    Call ImageScanner.StopScan
    Call ImageScanner.CloseScanner

    Exit Sub
errProcess:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dcmMiniature_Click()
On Error GoTo errHandle
    If mintCurImgIndex = 0 Then Exit Sub
    
   '��ѡ�е�ͼ�񵥶����ص�dcmViewImg��ȥ
    Call LoadViewImg
    
    mblnImgProcess = True

    Call cbrMain_Resize
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadViewImg()
On Error GoTo errH
    Dim ImgTmpImage As New DicomImage
    
    dcmViewImg.Images.Clear
    Set ImgTmpImage = dcmMiniature.Images.Item(mintCurImgIndex)
    
    dcmViewImg.Images.Add ImgTmpImage.SubImage(0, 0, ImgTmpImage.SizeX, ImgTmpImage.SizeY, 1, ImgTmpImage.Frame)
    dcmViewImg.Visible = True
    Exit Sub
errH:
    MsgBox "LoadViewImg�����쳣:" & err.Description, vbInformation, gstrSysName
End Sub

Private Sub dcmMiniature_OnDataChanged()
    On Error GoTo errHandle

    If dcmMiniature.Images.Count = 0 Then
        RaiseEvent RefreshState(False)
    ElseIf dcmMiniature.Images.Count > 0 Then
        RaiseEvent RefreshState(True)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dcmViewImg_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next

    If mblndcmViewImgDown = True And Button = 1 And dcmViewImg.Images.Count > 0 Then
        dcmViewImg.Images(1).ScrollX = mInitScrollPoint.X - X
        dcmViewImg.Images(1).ScrollY = mInitScrollPoint.Y - Y

        dcmViewImg.Refresh
    End If
End Sub

Private Sub dcmViewImg_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle
    Dim intLabelType As Integer

    If Button = 1 And dcmViewImg.Images.Count > 0 Then
        mMouseDownPoint.X = dcmViewImg.Images(1).ActualScrollX
        mMouseDownPoint.Y = dcmViewImg.Images(1).ActualScrollY
          
        mInitScrollPoint.X = dcmViewImg.Images(1).ScrollX + X
        mInitScrollPoint.Y = dcmViewImg.Images(1).ScrollY + Y
        
        mblndcmViewImgDown = True
        
        '��¼��ǰ�������
        mlngBaseX = X
        mlngBaseY = Y
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dcmViewImg_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle

    If mblndcmViewImgDown = True And Button = 1 And dcmViewImg.Images.Count > 0 Then
        '����ͼ�����ε�ƫ��λ��
        mCorpSize.X = mCorpSize.X + (dcmViewImg.Images(1).ActualScrollX - mMouseDownPoint.X)
        mCorpSize.Y = mCorpSize.Y + (dcmViewImg.Images(1).ActualScrollY - mMouseDownPoint.Y)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dcmViewImg_MouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
On Error GoTo errHandle
    '�������¼� ʵ���϶�
     Dim dblZoom As Double
    dblZoom = dcmViewImg.Images(1).ActualZoom
    dblZoom = dblZoom * (1 + Delta * 0.001)
    If dblZoom < 64 And dblZoom > 0.01 Then
        subCenterZoom dcmViewImg.Images(1), dcmViewImg, dblZoom, mCorpSize
    End If
    mlngBaseY = Y
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdMagnify_Click()
On Error GoTo errHandle
'��ť�Ŵ�
Dim dblZoom As Double

    dblZoom = dcmViewImg.Images(1).ActualZoom
    dblZoom = dblZoom * (1 + 300 * 0.001)
    If dblZoom < 64 And dblZoom > 0.01 Then
        subCenterZoom dcmViewImg.Images(1), dcmViewImg, dblZoom, mCorpSize
    End If
    'mlngBaseY = y
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cmdReduce_Click()
On Error GoTo errHandle
'��ť��С
    Dim dblZoom As Double
    
    dblZoom = dcmViewImg.Images(1).ActualZoom
    dblZoom = dblZoom * (1 + (-240) * 0.001)
    If dblZoom < 64 And dblZoom > 0.01 Then
        subCenterZoom dcmViewImg.Images(1), dcmViewImg, dblZoom, mCorpSize
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cmdReset_Click()
On Error GoTo errHandle
'���ð�ť�Լ�ͼ��
    
    '��ʼ���϶�������ƫ��λ��
    mCorpSize.X = 0
    mCorpSize.Y = 0
    
    '����ͼ��
    Call LoadViewImg
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub subCenterZoom(img As DicomImage, Viewer As DicomViewer, dblZoom As Double, corpSize As TPoint)
'------------------------------------------------
'���ܣ���ͼ��������š��Ե�ǰviewer���ĵ�Ϊ�������ĵ㡣
'������ img -- �������ŵ�ͼ��
'       viewer ���� ͼ�����ڵ�viewer
'       dblZoom ����ͼ���µ����ű���
'------------------------------------------------
    img.Zoom = dblZoom
    img.StretchToFit = False

    img.ScrollX = (img.SizeX * img.ActualZoom - ScaleX(Viewer.Width, vbTwips, vbPixels) / Viewer.MultiColumns) / 2 + corpSize.X
    img.ScrollY = (img.SizeY * img.ActualZoom - ScaleY(Viewer.Height, vbTwips, vbPixels) / Viewer.MultiRows) / 2 + corpSize.Y
End Sub





Private Sub Form_Load()
'��������¼�

Dim strFolder As String
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitCommandBars
    
    mstrTempDirOfScan = App.Path + CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME
    If Len(mstrTempDirOfScan) > 45 Then
        
        Dim pathlen As Long

        strFolder = String(256, 0)
        pathlen = GetTempPath(256, strFolder)
        If pathlen > 0 Then
            mstrTempDirOfScan = Left(strFolder, pathlen - 1) + CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME
        End If
    End If
    
    '���ݲ����ж� ��ǰ�ǲ鿴���뵥����ɨ�����뵥,���ǲ鿴������ĸ�������ť
    If mblnIsExamine Then mblnOperate = False
    
    '��ʼ������ ͼ��߼�����ť
    mblnImgProcess = False
    
    '�����豸��ʱĿ¼
    mstrScanDeviceTempDir = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPetitionCapture", "ɨ����ʱĿ¼", "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA")
    
    mintImageType = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPetitionCapture", "ɨ��ͼ���ʽ", 0))

    '��������ڴ��̵ĸ�Ŀ¼��app.pathΪ��x:\��
    mstrBufferDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
    
    '������ͼ����ص�DcmViewer�ؼ�����ʾ
    Call GetPetitionImages(dcmMiniature, mlngAdviceID, 100)

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub Form_Unload(Cancel As Integer)
On Error GoTo errHandle
    Call SaveWinState(Me, App.ProductName)

    '�رմ���ʱ �Ͽ���ǰFTP����
    miNet.FuncFtpDisConnect
    
    Exit Sub
errHandle:
    '�Ͽ�FTP����
    miNet.FuncFtpDisConnect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub InitLables()
'���ݴ����ֵ�����˻�����Ϣlbl��ֵ
    labName.Caption = mstrPatientName
    labName.ToolTipText = mstrPatientName
    
    labNo.Caption = mstrCheckNo
    labNo.ToolTipText = mstrCheckNo
    
    labRoom.Caption = mstrDeptName
    labRoom.ToolTipText = mstrDeptName
    
    labSex.Caption = mstrPatientSex
    labSex.ToolTipText = mstrPatientSex
    
    labAge.Caption = mstrPatientAge
    labAge.ToolTipText = mstrPatientAge
    
    labCheckMethod.Caption = mstrExamineMethod
    labCheckMethod.ToolTipText = mstrExamineMethod
    
    labSpePosition.Caption = mstrSpePosition
    labSpePosition.ToolTipText = mstrSpePosition

End Sub

Public Sub InitDeptPara(ByVal lngDeptID As Long)
'��ʼ�����Ҽ�����
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo DBError
    mlngCurDeptId = lngDeptID
    
    '��ȡ�����洢�豸��
    mstrSaveDeviceID = GetDeptPara(mlngCurDeptId, "���뵥�洢�豸��")
    gstrSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�õ��豸��", mstrSaveDeviceID)
    If rsTmp.EOF Then
        MsgBox "���뵥�洢�豸δ�������ͣ�ã����飡", vbInformation, gstrSysName
        mstrSaveDeviceID = ""
        Exit Sub
    End If
    
    Call funGetStorageDevice(Me, mstrSaveDeviceID, mFtpDir, mFtpIp, mFtpUser, mFtpPass)
    
    Exit Sub
DBError:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imageScanner_PageDone(ByVal PageNumber As Long)
On Error GoTo errHandle
      Dim strScanFile As String

      If mintScanImageIndex = -1 Then
        Exit Sub
      End If
    
      '����ɨ���ļ�����
      mintScanImageIndex = mintScanImageIndex + 1
    
      
      strScanFile = mintScanImageIndex
    
      'ȡ����Ч��ɨ���ļ�����
      While Len(strScanFile) < 4
        strScanFile = "0" + strScanFile
      Wend
    
      strScanFile = mstrTempDirOfScan & "\" & CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE & strScanFile & ".bmp"
    
      '����ɨ���ͼ��
      Call subCaptureImg(True, strScanFile)

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub subCaptureImg(ByVal RealTimeCap As Boolean, Optional ByVal strFileName As String = "", _
    Optional ByRef picCapture As StdPicture = Nothing, Optional ByVal blnIsAfterCapture As Boolean = False)
'------------------------------------------------
'����: ɨ�貢�洢ͼ��
'��������
'���أ��ޣ�ֱ�ӱ����²ɼ���ͼ��
'------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If dcmMiniature.Images.Count > 9 Then
        MsgBoxD Me, "�Ѿ�����10�����뵥���������ɨ�裬����ɾ��ǰ�治��Ҫ�����뵥��", vbInformation, gstrSysName
        Exit Sub
    End If

    If mblnIsLogin Then
        If funCaptureSingleImage(RealTimeCap, strFileName, picCapture) = False Then
            MsgBoxD Me, "ͼ�����ʧ�ܡ�", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        If funCaptureSingleImage(RealTimeCap, strFileName, picCapture) = True Then
            '���ñ���ͼ�񵽷����� ����
            Call subSaveImage(, mlngAdviceID)
        End If
    End If
    
    lab.Visible = False
    img.Visible = False
    
End Sub




Private Function funCaptureSingleImage(ByVal RealTimeCap As Boolean, _
    Optional ByVal strFileName As String = "", Optional ByRef picCapture As StdPicture = Nothing) As Boolean
'------------------------------------------------
'���ܣ���ͼ���������ͼdcmMiniature�С�
'��������
'���أ��ޣ�ֱ�ӽ��²ɼ���ͼ�����dcmMiniature��
'------------------------------------------------

    Dim ImgTmpImage As New DicomImage

    On Error GoTo SaveFileError

    'ɨ��ͼ��
    Call Clipboard.Clear

    If Not (picCapture Is Nothing) Then
        Set picTemp2.Picture = Nothing
        Set picTemp2.Picture = picCapture

    ElseIf Trim(strFileName) <> "" And Dir(strFileName) <> "" Then
        'ʹ��dcmView��ʾ����ͼƬ������Ҫ�ٲü�
        Set picTemp2.Picture = Nothing
        Set picTemp2.Picture = LoadPicture(strFileName)

    Else
        Set picTemp2.Picture = Nothing

    End If

    '��ͼ���ٴ��ύ�����а�
    If picTemp2.Picture Is Nothing Then
        funCaptureSingleImage = False
        Exit Function
    End If


    Call Clipboard.SetData(picTemp2.Picture, 2)
'    �Ӽ��а�ȡ��ͼ��
    Call ImgTmpImage.Paste

    Call Clipboard.Clear

    '��ͼ���������ͼ��
    Call subInsert2Mini(ImgTmpImage)

    funCaptureSingleImage = True

    Exit Function
SaveFileError:
    If ErrCenter() = 1 Then
        Resume
    End If

    Call SaveErrLog
End Function

Private Sub subInsert2Mini(img As DicomImage)
'------------------------------------------------
'���ܣ���ͼ����ӵ�����ͼdcmMiniature��
'������img���������ͼ��
'���أ��ޣ�ֱ�ӽ�ͼ����ӵ�����ͼdcmMiniature��
'------------------------------------------------
    Dim iRows As Integer
    Dim iCols As Integer

    '��������ͼ��ͼ�񲼾�
    ResizeRegion dcmMiniature.Images.Count + 1, dcmMiniature.Width, dcmMiniature.Height, iRows, iCols

    dcmMiniature.MultiColumns = iCols
    dcmMiniature.MultiRows = iRows

    dcmMiniature.Images.Add img

    '��������ͼ�б�ѡ�е�״̬
    If mintCurImgIndex > 0 And mintCurImgIndex <= dcmMiniature.Images.Count Then
        dcmMiniature.Images(mintCurImgIndex).BorderColour = vbWhite
    End If


    With dcmMiniature.Images(dcmMiniature.Images.Count)
        .BorderWidth = 1
        .BorderStyle = 6
        .BorderColour = vbRed
    End With

    If Not mdcmTmpView Is Nothing Then
        mdcmTmpView.Images.Add img
    End If
    
    mintCurImgIndex = dcmMiniature.Images.Count
    Call dcmMiniature_Click
End Sub


Public Sub subSaveImage(Optional iEncode As Integer = 0, Optional lngAdviceID As Long, Optional dcmTmpView As DicomViewer = Nothing)
'------------------------------------------------
'���ܣ������һ������ͼ���浽���ݿ���
'������iEncode����ѹ����ʽ��1��Run-Length Encoding�г�ѹ����2������������ԭͼ��ѹ����ʽ��������Lossless JPEG encoding JPEG����ѹ��
'���أ���
'------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim ImgTmp As New DicomImage
    
    Dim strReceived As String
    Dim strFileTitle As String       'ͼ���ļ���ͷ
    Dim lngResult As Long         'FTP�������
    Dim blnResult As Boolean
    Dim nowTime As Date
    Dim blnInTrans As Boolean       '�����ﴦ�������
    Dim strRandom As String
    Dim lngImage As Long
    Dim strSql As String
    Dim arrSQL() As String
    Dim arrImages() As String       '�ϴ�FTP�ɹ���ͼƬ
    Dim i As Long
    
    On Error GoTo errHandle
    
    If Not dcmTmpView Is Nothing Then
        If dcmTmpView.Images.Count <= 0 Then Exit Sub
        lngImage = dcmTmpView.Images.Count
    Else
        If dcmMiniature.Images.Count <= 0 Then Exit Sub
        '��ȡ���һ������ͼ
        Set ImgTmp = dcmMiniature.Images(dcmMiniature.Images.Count)
        lngImage = 1
    End If
    
    
    lngResult = miNet.FuncFtpConnect(mFtpIp, mFtpUser, mFtpPass)
    
    If lngResult = 0 Then
        'FTP����ʧ�ܣ���ʾ���󣬲�ɾ������ͼ�е�ͼ��
        MsgBox "FTP����ʧ�ܣ�ͼ���޷����棬�����������á�", vbInformation, gstrSysName

        If dcmTmpView Is Nothing Then
            dcmMiniature.Images.Remove (i)
        End If
            
        Exit Sub
    End If

    nowTime = zlDatabase.Currentdate
    strReceived = Format(nowTime, "yyyymmdd")
    
    '��������Ŀ¼
    MkLocalDir mstrBufferDir & strReceived & "/" & lngAdviceID & "/"
    
    ReDim arrImages(0)
    ReDim arrSQL(0)
    For i = 1 To lngImage
        
        If Not dcmTmpView Is Nothing Then
            Set ImgTmp = dcmTmpView.Images(i)
        End If
        
        '�õ������
        strRandom = CInt(Rnd * 100 + 1)
    
        nowTime = zlDatabase.Currentdate
        strFileTitle = Format(nowTime, "mmddhhmmss") & Format((Timer() * 1000) Mod 1000, "000")
    
        '����ͼ�񵽻���Ŀ¼  Lossless JPEG encoding JPEG����ѹ��
        ImgTmp.WriteFile mstrBufferDir & strReceived & "/" & lngAdviceID & "/" & strFileTitle & lngAdviceID & strRandom, True
    
        ImgTmp.FileExport mstrBufferDir & strReceived & "/" & lngAdviceID & "/" & strFileTitle & lngAdviceID & strRandom & ".jpg", "JPG", 80
    
        ImgTmp.tag = strFileTitle & lngAdviceID & strRandom & ".jpg"
    
        '����ɨ�赥ͼ��
        blnResult = WriteToURL(mstrBufferDir & strReceived & "/" & lngAdviceID & "/" & strFileTitle & lngAdviceID & strRandom, mFtpDir & _
            strReceived & "/" & lngAdviceID & "/" & strFileTitle & lngAdviceID & strRandom)
        
        If blnResult Then
            '�ϴ�FTP�ɹ������뵥��¼�����飬����ʧ�ܺ�ɾ��
            ReDim Preserve arrImages(UBound(arrImages) + 1)
            arrImages(UBound(arrImages)) = mFtpDir & strReceived & "/" & lngAdviceID & "/" & strFileTitle & lngAdviceID & strRandom
            
            'ͼ��洢�ɹ��󣬴洢���ݿ���Ϣ
            strSql = "ZL_Ӱ�����뵥ͼ��_INSERT ('" & lngAdviceID & "','" & strFileTitle & lngAdviceID & strRandom & ".jpg" & "','" & strReceived & "/" & lngAdviceID & "','" & mstrSaveDeviceID & "','" & UserInfo.���� & "',sysdate)"
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSql
        End If
    Next
    
    miNet.FuncFtpDisConnect
    
    '����ͼ��
    gcnOracle.BeginTrans
    
    blnInTrans = True
    For i = 1 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "����ͼ��")
    Next
    
    gcnOracle.CommitTrans
    
    blnInTrans = False
    
    '���mblnIsLogin=true ��ô˵�������ڵǼǽ���ı���ͼ����Ҫ���ò�������Ϊfalse
    If mblnIsLogin Then
        mblnIsLogin = False
    End If
    
    Exit Sub
errHandle:
    If blnInTrans Then
        gcnOracle.RollbackTrans
        blnInTrans = False
    End If
    
    '�Ͽ�FTP����
    miNet.FuncFtpDisConnect
    
    Call CancelImagesUp(arrImages)
    
    MsgBox "���뵥ͼ�񱣴�ʧ�ܡ�", vbInformation, gstrSysName
    
    If dcmTmpView Is Nothing Then
        dcmMiniature.Images.Remove (dcmMiniature.Images.Count)
    End If
End Sub

Private Sub CancelImagesUp(arrImages() As String)
    Dim i As Long
    Dim objFtp As clsFtp
    
    Call miNet.FuncFtpConnect(mFtpIp, mFtpUser, mFtpPass)
    
    For i = 1 To UBound(arrImages)
        RemoveFromURL arrImages(i)
    Next
    
    miNet.FuncFtpDisConnect
End Sub


Public Sub GetPetitionImages(dcmViewer As DicomViewer, lngAdviceID As Long, _
Optional intGetImgNum As Integer = 0, Optional intShowImgNum As Integer = 0)
'------------------------------------------------
'���ܣ�ɾ��dcmViewer�е�ͼ��󣬽���ȡ��ͼ���ļ�����dcmViewer��
'������ dcmViewer������ͼ���DicomViewer
'       lngAdviceID ���� ҽ��ID
'       intGetImgNum �������ζ�ȡ��ͼ������
'       intShowImgNum ����������ʾ��ͼ������
'���أ��ޣ�ֱ���޸�dcmViewer����ʾ��ͼ��
'------------------------------------------------

    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim curImage As DicomImage
    Dim iCols As Integer, iRows As Integer
    Dim objFile As New Scripting.FileSystemObject, strTmpFile As String
    Dim Inet1 As New clsFtp
    Dim strFTPUser As String              'FTP�˺�
    Dim strFtpPass As String              'FTP����
    Dim strFtpDir As String               'FTPĿ¼
    Dim strFtpIp As String                'FTP��ַ
    Dim strFirstDevNo As String
    Dim strNextDevNo As String
    Dim strTmpFolder As String
    Dim i As Long
    
    On Error GoTo DBError
    
    If mblnIsLogin Then
        If mdcmTmpView.Images.Count > 0 Then
            lab.Visible = False
            img.Visible = False
            ResizeRegion mdcmTmpView.Images.Count, dcmViewer.Width, dcmViewer.Height, iRows, iCols
            dcmViewer.MultiColumns = iCols
            dcmViewer.MultiRows = iRows
            
            For i = 1 To mdcmTmpView.Images.Count
                dcmViewer.Images.Add mdcmTmpView.Images(i)
                dcmViewer.Images(i).BorderWidth = 1
            Next
        
        End If
    Else
       strSql = "select ���뵥ͼ��,ɨ����,ɨ��ʱ��,FTP·��,�豸�� from Ӱ�����뵥ͼ�� where ҽ��ID=[1] order by �豸��"
       Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���뵥ͼ����Ϣ", lngAdviceID)

        'dcmViewer.Images.Clear
        If rsTmp.RecordCount > 0 Then
            lab.Visible = False
            img.Visible = False
            ResizeRegion rsTmp.RecordCount, dcmViewer.Width, dcmViewer.Height, iRows, iCols

            dcmViewer.MultiColumns = iCols
            dcmViewer.MultiRows = iRows
            
            mstrBufferDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")

            strFirstDevNo = Nvl(rsTmp("�豸��"))
   
            Do While Not rsTmp.EOF
                strTmpFolder = mstrBufferDir & objFile.GetParentFolderName(Nvl(rsTmp("FTP·��")) & "/" & Mid(Nvl(rsTmp("���뵥ͼ��")), 1, InStr(Nvl(rsTmp("���뵥ͼ��")), ".") - 1))
                '��������Ŀ¼
                If Not objFile.FolderExists(strTmpFolder) Then MkLocalDir (strTmpFolder)
            
                If strFirstDevNo <> strNextDevNo Then
                    Call funGetStorageDevice(Me, Nvl(rsTmp("�豸��")), strFtpDir, strFtpIp, strFTPUser, strFtpPass)
                    
                    '�ж�FTP�Ƿ����ӳɹ�
                    If Inet1.FuncFtpConnect(strFtpIp, strFTPUser, strFtpPass) = 0 Then
                        MsgBoxD Me, "FTP�����������ӣ������������á�"
                        Exit Sub
                    End If
                End If
                
                strTmpFile = mstrBufferDir & Nvl(rsTmp("FTP·��")) & "/" & Mid(Nvl(rsTmp("���뵥ͼ��")), 1, InStr(Nvl(rsTmp("���뵥ͼ��")), ".") - 1)

                If Dir(strTmpFile) = vbNullString Then
                    '���ػ���ͼ�񲻴��ڣ����ȡFTPͼ��

                    If Inet1.FuncDownloadFile(objFile.GetParentFolderName(strFtpDir & Nvl(rsTmp("FTP·��")) & "/" & Mid(Nvl(rsTmp("���뵥ͼ��")), 1, InStr(Nvl(rsTmp("���뵥ͼ��")), ".") - 1)), strTmpFile, objFile.GetFileName(Nvl(rsTmp("FTP·��")) & "/" & Mid(Nvl(rsTmp("���뵥ͼ��")), 1, InStr(Nvl(rsTmp("���뵥ͼ��")), ".") - 1))) <> 0 Then
                        '����ͼ��ʧ��
                        MsgBoxD Me, "���ع�������δ֪��������ϵϵͳ����Ա��"
                        Exit Sub
                    End If
                End If

                If Dir(strTmpFile) <> vbNullString Then
                        
                        Set curImage = dcmViewer.Images.ReadFile(strTmpFile)
                        curImage.tag = Nvl(rsTmp("���뵥ͼ��"))
                        
                        With curImage
                            .BorderStyle = 6
                            .BorderWidth = 1
                            .BorderColour = vbWhite
                        End With

                    'ȡ���Զ���Ӱ,��ΪDicomObjects�ؼ�����Դ����Ӱ��BUG�����ڣ�0028��6100��ʱ�����Զ���ͼ����м�Ӱ��
                    '���½�ú��DSAͼ����������ʾ
                    '��Ȼ����ͼ���mask=0 ,����ȡ����Ӱ������ÿ��ͼ����ӵ��µ�Dicomimages֮���Զ��ֽ�mask���ó�1�ˣ�
                    '�����ڳ������޷��ܺõĿ��ƣ����ֱ��ȥ����0028��6100��������ԡ�
                    If Not IsNull(curImage.Attributes(&H28, &H6100).value) Then
                        curImage.Attributes.Remove &H28, &H6100
                    End If
                End If

                rsTmp.MoveNext
                If Not rsTmp.EOF Then strNextDevNo = Nvl(rsTmp("�豸��"))
            Loop
        End If

        Inet1.FuncFtpDisConnect
    End If
    
    If dcmViewer.Images.Count > 0 Then
        dcmViewer.CurrentIndex = 1
        dcmViewer.Images(dcmViewer.Images.Count).BorderColour = vbRed
        
        mintCurImgIndex = dcmViewer.Images.Count
        Call dcmMiniature_Click
    Else
        lab.Visible = True
        img.Visible = True
        dcmViewImg.Visible = False
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
    End If
        
    Exit Sub
DBError:
    '�Ͽ�FTP����
    Inet1.FuncFtpDisConnect
    If ErrCenter() = 1 Then
        Resume
    End If

    Call SaveErrLog
End Sub

Private Sub subDeleteFTPImage()
'------------------------------------------------
'���ܣ�ɾ������ͼ�б�ѡ�е�ͼ���ȴ����ݿ���ɾ����Ȼ���FTP��ɾ��.
'��������
'���أ��ޣ�ֱ��ɾ������ͼ�����һ��ͼ��
'------------------------------------------------
On Error GoTo errH
    Dim blnResult As Boolean
    Dim i As Integer, iRows As Integer, iCols As Integer
    
    If dcmMiniature.Images.Count > 0 And mintCurImgIndex > 0 And mintCurImgIndex <= dcmMiniature.Images.Count Then
        
        '�����ݿ��FTP��ɾ������ͼ�б�ѡ�е�ͼ��
        blnResult = DelPetitionImg()
        
        If blnResult = True Then    'ɾ���ɹ������޸�����ͼ״̬��������StateChanged�¼�
            '������ͼ��ɾ��ͼ��
            dcmMiniature.Images.Remove mintCurImgIndex
            
            If dcmMiniature.Images.Count = 0 Then
                'ɾ����ʱ��ֻ��һ��ͼ,ɾ����ɺ���ת�Ȱ�ť����Ϊ�����ã��ұߴ�ͼ����
                lab.Visible = True
                img.Visible = True
            
                mblnImgProcess = False
                mintCurImgIndex = 0
                dcmViewImg.Visible = False
            Else
                'ɾ��ʱ�ж���ͼ��ɾ����ɺ��Զ�ѡ��ǰһ��ͼ
                For i = mintCurImgIndex + 1 To dcmMiniature.Images.Count
                    Call dcmMiniature.Images.Move(i, i - 1)
                Next i
                
                '���²���
                '��������ͼ��ͼ�񲼾�
                ResizeRegion dcmMiniature.Images.Count + 1, dcmMiniature.Width, dcmMiniature.Height, iRows, iCols
            
                dcmMiniature.MultiColumns = iCols
                dcmMiniature.MultiRows = iRows
    
                Call dcmMiniature.Refresh

                If mintCurImgIndex > 1 Then
                    mintCurImgIndex = mintCurImgIndex - 1
                Else
                    mintCurImgIndex = 1
                End If
                dcmMiniature.Images(mintCurImgIndex).BorderColour = vbRed

                Call dcmMiniature_Click
            End If
            
            
        
        End If
    End If
    Exit Sub
errH:
    MsgBoxD Me, "ɾ��ʧ��-" & err.Description, vbInformation, gstrSysName
End Sub

Private Sub subDeleteDcmImage()

'������ͼ��ɾ��ͼ��
        dcmMiniature.Images.Remove mintCurImgIndex
        
        If mintCurImgIndex > dcmMiniature.Images.Count Then
            mintCurImgIndex = dcmMiniature.Images.Count
        End If

        If mintCurImgIndex > 0 Then
            dcmMiniature.Images(mintCurImgIndex).BorderColour = vbRed
        End If
        

End Sub


Private Function DelPetitionImg() As Boolean
'------------------------------------------------
'���ܣ������ݿ��FTP��ɾ��ͼ��ɾ������ͼ�б�ѡ�е�ͼ��
'��������
'���أ�True����ɾ���ɹ���False����ɾ��ʧ��

    Dim ImgTmp As New DicomImage
    Dim rsTmp As New ADODB.Recordset
    Dim strReportImage As String
    Dim varTmp As Variant
    Dim strResult As Long
    Dim strSql As String
    Dim strFTPUser As String              'FTP�˺�
    Dim strFtpPass As String              'FTP����
    Dim strFtpDir As String               'FTPĿ¼
    Dim strFtpIp As String                'FTP��ַ
    
    If dcmMiniature.Images.Count < mintCurImgIndex Then Exit Function
    Set ImgTmp = dcmMiniature.Images(mintCurImgIndex)
                
    On Error GoTo errHand
    
    strSql = "select ɨ��ʱ��,�豸�� from Ӱ�����뵥ͼ�� where ҽ��ID=[1] and ���뵥ͼ�� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���뵥ͼ����Ϣ", mlngAdviceID, ImgTmp.tag)
    
    If rsTmp.EOF = True Then
        MsgBoxD Me, "û���ҵ�����ɾ����ͼ��!", vbInformation, gstrSysName
        DelPetitionImg = False
        Exit Function
    End If
    
    Call funGetStorageDevice(Me, Nvl(rsTmp("�豸��")), strFtpDir, strFtpIp, strFTPUser, strFtpPass)
    
    strResult = miNet.FuncFtpConnect(strFtpIp, strFTPUser, strFtpPass)
    If strResult = 0 Then
        MsgBoxD Me, "����FTPʧ�ܣ�����FTP���ӡ�", vbInformation, gstrSysName
        DelPetitionImg = False
        Exit Function
    End If

    gstrSQL = "ZL_Ӱ�����뵥ͼ��_DELETE(" & mlngAdviceID & ",'" & ImgTmp.tag & "')"

    zlDatabase.ExecuteProcedure gstrSQL, "Ӱ��ͼ��ɾ��"

    'ɾ��ͼ���ļ�
    RemoveFromURL strFtpDir & _
        Format(Nvl(rsTmp("ɨ��ʱ��"), CStr(zlDatabase.Currentdate)), "yyyyMMdd") & "/" & _
        mlngAdviceID & "/" & Mid(ImgTmp.tag, 1, InStr(ImgTmp.tag, ".") - 1)

    miNet.FuncFtpDisConnect
    DelPetitionImg = True
    Exit Function
errHand:
    '�Ͽ�FTP����
    miNet.FuncFtpDisConnect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function WriteToURL(ByVal strFileName As String, ByVal strDestFileName As String) As Boolean
'------------------------------------------------
'���ܣ��������ļ����浽Զ��������
'������strFileName--�����ļ�����strDestFileName����Զ���ļ���
'���أ���
'-----------------------------------------------
'���ܣ�
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strPath As String
    Dim lngReturn As Long
    
    '��FTP�д���Ŀ¼
    strPath = objFileSystem.GetParentFolderName(strDestFileName)
    miNet.FuncFtpMkDir "/", strPath
    
    '��FTP�ϴ��ļ�
    lngReturn = miNet.FuncUploadFile(strPath, strFileName, objFileSystem.GetFileName(strDestFileName))
    
    If lngReturn = 0 Then
        WriteToURL = True
    Else
        WriteToURL = False
    End If
End Function


Private Sub RemoveFromURL(ByVal strDestFileName As String)
'------------------------------------------------
'���ܣ���FTPɾ���ļ�
'������strDestFileName����Զ���ļ���
'���أ���
'-----------------------------------------------
    Dim objFileSystem As New Scripting.FileSystemObject
    
    miNet.FuncDelFile objFileSystem.GetParentFolderName(strDestFileName), objFileSystem.GetFileName(strDestFileName)
End Sub

Private Sub dcmMiniature_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle
    Dim i As Integer
    Dim j As Integer

    If Button = 1 Then
        mCorpSize.X = 0
        mCorpSize.Y = 0
        
        '��ѡ��ͼ����ʾ���
        i = dcmMiniature.ImageIndex(X, Y)
        If i <> 0 Then
        
            For j = 1 To dcmMiniature.Images.Count
                dcmMiniature.Images(j).BorderColour = vbWhite
            Next
            dcmMiniature.Images(i).BorderColour = vbRed
            mintCurImgIndex = i
            
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbrMain_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Dim lngHeightdcmMiniature As Long
    
    On Error Resume Next
    
    cbrMain.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    
    lngHeightdcmMiniature = Me.ScaleHeight - fmeInfoCtrl.Height - lngTop - 120

    Call fmeInfoCtrl.Move(0, lngTop, IIf(lngRight > 2700, 2700, lngRight), 3000)
    Call fmePatientInfo.Move(60, 0, fmeInfoCtrl.Width - 100, 3000)
    Call fmeDcmViewer.Move(fmeInfoCtrl.Width, lngTop, Me.ScaleWidth - fmeInfoCtrl.Left - fmeInfoCtrl.Width - 120, Me.ScaleHeight - lngTop - 60)
    Call dcmMiniature.Move(60, fmeInfoCtrl.Top + fmeInfoCtrl.Height + 60, fmeInfoCtrl.Width - 120, lngHeightdcmMiniature) 'LTWH
    Call dcmViewImg.Move(60, 60, fmeDcmViewer.Width - 120, fmeDcmViewer.Height - 60)
    
    Call lab.Move((fmeDcmViewer.Width - lab.Width) / 2, (fmeDcmViewer.Height - lab.Height) / 2)
    Call img.Move((fmeDcmViewer.Width - img.Width) / 2, lab.Top - img.Height - 120)
End Sub



