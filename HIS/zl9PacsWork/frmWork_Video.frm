VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVideoProcess.ocx"
Begin VB.Form frmWork_Video 
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   10425
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmWork_Video.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   10425
   StartUpPosition =   3  '����ȱʡ
   Begin ScanLibCtl.ImgScan ImageScanner 
      Left            =   120
      Top             =   3120
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
   Begin VB.Timer tmrReg 
      Interval        =   10000
      Left            =   30
      Top             =   6660
   End
   Begin VB.Timer timerHook 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   6090
   End
   Begin VB.PictureBox picDock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   480
      ScaleHeight     =   8295
      ScaleWidth      =   9015
      TabIndex        =   3
      Top             =   120
      Width           =   9015
      Begin zl9PACSWork.ucSplitter ucSplitter1 
         Height          =   135
         Left            =   0
         TabIndex        =   4
         Top             =   6015
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   238
         MousePointer    =   7
         SplitType       =   0
         DBClickType     =   2
         SplitLevel      =   3
         Con1MinSize     =   3000
         Con2MinSize     =   1000
         Control1Name    =   "picCapture"
         Control2Name    =   "ucPreview"
      End
      Begin VB.PictureBox picCapture 
         ForeColor       =   &H00000000&
         Height          =   6015
         Left            =   0
         ScaleHeight     =   5955
         ScaleWidth      =   8955
         TabIndex        =   6
         Top             =   0
         Width           =   9015
         Begin VB.PictureBox picView 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   3495
            Left            =   600
            ScaleHeight     =   3495
            ScaleWidth      =   6855
            TabIndex        =   11
            Top             =   240
            Width           =   6855
            Begin ZLDSVideoProcess.DSCapture wdmCapture 
               Height          =   3135
               Left            =   720
               TabIndex        =   12
               Top             =   240
               Width           =   3495
               Object.Visible         =   -1  'True
               AutoScroll      =   0   'False
               AutoSize        =   0   'False
               AxBorderStyle   =   1
               Caption         =   ""
               Color           =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               KeyPreview      =   -1  'True
               PixelsPerInch   =   96
               PrintScale      =   1
               Scaled          =   -1  'True
               DropTarget      =   0   'False
               HelpFile        =   ""
               ScreenSnap      =   0   'False
               SnapBuffer      =   10
               DoubleBuffered  =   0   'False
               Enabled         =   -1  'True
               IsStretch       =   0   'False
               IsShowState     =   -1  'True
               IsFullScreen    =   0   'False
               IsAdjustWindowSize=   0   'False
               IsFit           =   0   'False
               IsEscKeyQuitFullScreen=   -1  'True
               IsDblClickQuitFullScreen=   0   'False
               IsClickQuitFullScreen=   0   'False
               CurWidth        =   233
               CurHeight       =   209
               CurVideoWidth   =   231
               CurVideoHeight  =   189
               ShowModel       =   0
               CapParameterWindPos=   8
               SnatchWay       =   0
               ParameterCfgFileName=   ""
               HideCfgItem     =   0
               AppHandle       =   0
            End
            Begin VB.TextBox txtInputText 
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   5520
               TabIndex        =   14
               Text            =   "Text1"
               Top             =   840
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.PictureBox picVideo 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   3015
               Left            =   1200
               ScaleHeight     =   201
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   224
               TabIndex        =   13
               Top             =   120
               Width           =   3360
            End
            Begin DicomObjects.DicomViewer dcmView 
               Height          =   1575
               Left            =   4440
               TabIndex        =   15
               Top             =   1440
               Width           =   2175
               _Version        =   262147
               _ExtentX        =   3836
               _ExtentY        =   2778
               _StockProps     =   35
               BackColor       =   0
               UseScrollBars   =   0   'False
            End
         End
         Begin VB.PictureBox pbxSize 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Height          =   75
            Index           =   0
            Left            =   360
            MousePointer    =   7  'Size N S
            ScaleHeight     =   75
            ScaleWidth      =   7335
            TabIndex        =   10
            Top             =   120
            Width           =   7335
         End
         Begin VB.PictureBox pbxSize 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Height          =   3975
            Index           =   3
            Left            =   7560
            MousePointer    =   9  'Size W E
            ScaleHeight     =   3975
            ScaleWidth      =   75
            TabIndex        =   9
            Top             =   15
            Width           =   75
         End
         Begin VB.PictureBox pbxSize 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Height          =   3975
            Index           =   2
            Left            =   480
            MousePointer    =   9  'Size W E
            ScaleHeight     =   3975
            ScaleWidth      =   75
            TabIndex        =   8
            Top             =   0
            Width           =   75
         End
         Begin VB.PictureBox pbxSize 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   75
            Index           =   1
            Left            =   360
            MousePointer    =   7  'Size N S
            ScaleHeight     =   75
            ScaleWidth      =   7335
            TabIndex        =   7
            Top             =   3840
            Width           =   7335
         End
         Begin XtremeCommandBars.CommandBars cbrMain 
            Left            =   0
            Top             =   0
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
         End
      End
      Begin zl9PACSWork.ucImagePreview ucPreview 
         Bindings        =   "frmWork_Video.frx":1CCA
         Height          =   2145
         Left            =   0
         TabIndex        =   5
         Top             =   6150
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3784
         BackColor       =   4210752
      End
   End
   Begin DicomObjects.DicomViewer dcmAfter 
      Height          =   735
      Left            =   8880
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1035
      _Version        =   262147
      _ExtentX        =   1826
      _ExtentY        =   1296
      _StockProps     =   35
      BackColor       =   12632319
   End
   Begin VB.PictureBox picBackImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   7680
      Picture         =   "frmWork_Video.frx":1CDE
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Timer tmrComm 
      Interval        =   2
      Left            =   0
      Top             =   5040
   End
   Begin MSCommLib.MSComm commListener 
      Left            =   0
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   0
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picTemp2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   45
      ScaleHeight     =   1455
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmWork_Video"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'���ܣ��ɼ���¼����Ƶͼ��
'
'
'
'�޸���ʷ��
'
'2010-01-19: ��wdm��Ƶ������뵽�ɼ�ģ���У���֧�ֶ�ָ��SDK��Ƶ�ɼ���ʵ��
'
'
'
'�ü�ԭ��˵����
'
'
'
'
'                ------------------------------------
'               |ԭʼͼ���С                        |
'               |                                    |
'               |                                    |
'               |         ------------------         |
'               |        |                  |        |
'               |<-- L-->|     �ü���С     |<-- R-->|
'               |        |                  |        |
'               |         ------------------         |
'               |                                    |
'               |                                    |
'               |                                    |
'                ------------------------------------
'
'L��ʾ��߲ü��Ĵ�С�ٷֱ�
'R��ʾ�ұ߲ü��Ĵ�С�ٷֱ�
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit

Implements IWorkMenu


Private Const M_STR_HINT_NoSelectData As String = "��Ч�ļ�����ݣ���ѡ����Ҫִ�еļ���¼��"
Private Const M_STR_MODULE_MENU_TAG As String = "�ɼ�"


'�ü���Χ modify by tjh 2010-01-19
Private Type TCutRange
  LeftRate As Double
  TopRate As Double
  WidthRate As Double
  HeightRate As Double
End Type


'��Ƶ���� modify by tjh 2010-01-19
Private Type TVideoArea
  Left As Long
  Top As Long
  Width As Long
  Height As Long
End Type


'�ƶ����� modify by tjh 2010-01-19
Private Enum TMoveOrientation
  moUp = 0    '��
  moDown = 1  '��
  moLeft = 2  '��
  moRight = 3 '��
End Enum

'����������
Private Type TPoint
  X As Integer
  Y As Integer
End Type


Private Type TUnLockStudyInf
    lngAdviceID As Long
    lngSendNO As Long
    blnMoved As Boolean
    lngStudyState As Long
End Type

'''��Ƶ�����¼�����
'Public Enum TVideoEventType
'    vetLockStudy = 1
'    vetAddFirstImg = 2
'    vetDelLastImg = 3
'    vetRecVideo = 4
'    vetUpdateImg = 5
'End Enum

Private mstrActiveType                  '���ʽ


Private WithEvents mclsDxDevice As clsDxHidDevice   'ʵ�������ֱ�֮��Ĳɼ���ʽ
Attribute mclsDxDevice.VB_VarHelpID = -1

Public mhCapWnd As Long                 '�ɼ����ڵľ��

Private mlngModul As Long
Private mstrPrivs As String              'ģ��Ȩ��
Private mlngCurDeptId As Long          '��ǰ����
Private mobjOwner As Object

Public pobjPacsCore As zl9PacsCore.clsViewer
Public mblnObserve As Boolean         '�Ƿ��й�Ƭ����Ȩ��   true��  false��


Private mRestoreContainer As Object
Private mParentContainer As Object
Public mIsShowing As Boolean

Private mblnRealTime As Boolean         '��¼��ǰ��ʾ����ʵʱ��ʾ����ͼ�����ڡ�True = ʵʱ��Ƶ���ڣ�False = ͼ������
Private mblnPlayVideo As Boolean        '��¼��ǰ��ʾ��ͼ����������ʾ����ͼƬ����¼��True = ¼��False = ͼƬ
Private mintMouseState As Integer       '��¼��ǰͼ����ʱ�����״̬:1=���ȶԱȶȣ�2=���ţ�3=�ü����ţ�11=��ͷ��ע��12=Բ�α�ע��13=���ֱ�ע


Private mlngBaseX As Long               'dcmView�����Downʱ��X����
Private mlngBaseY As Long               'dcmView�����Downʱ��Y����
Private mMouseDownPoint As TPoint       '�����DcmImage�ϰ���ʱ��λ��
Private mInitScrollPoint As TPoint      '��ʼ�϶�ʱ�ĳ�ʼλ��
Private mCorpSize As TPoint             '�϶�������ƫ��λ��

Private mstrTempDirOfScan As String          'ɨ�����ʱĿ¼
Private mintScanImageIndex As Integer        'ɨ��ͼ������

Private mstrInfor As String

Private mblnMoveDown  As Boolean         '�����ж��Ƿ���������
Private mblnDcmViewDown As Boolean      '�����ж�dcmView������Ƿ񱻰���
Private mintCurImgIndex As Integer      '��ǰ��ѡ�е�ͼ������
Private mdcmSelectLabel As DicomLabel   '��ǰ��ѡ�еı�ע
Private mstrAviFileName As String       '¼���ļ���
Private mstrEncoderName As String
Private mstrBufferDir As String

Private mintCapType As Integer            '��̤������ʽ��0-ֱ�Ӵ�����1-�任������2-��ƽ����
Private mintComInterval As Integer       '��̤��ͼ��ʱ��������λ��
Private mintComState As Integer          'COM�ڵ�״̬
Private mlngComTime As Long              '��¼com�ڱ���״̬��ʱ��
Private mdtLastCapture As Date           '�����̤���µ�ʱ��
Private mblnCTSHolding As Boolean        '��¼��̬ʱ��CTS�ߵĵ�ƽ
Private mstrComPort As Long              '���������Ķ˿ں�
Private mblnUseClipbord As Boolean          '�Ƿ�ʹ�ü�����

Private mobjFtpConnection As New clsFtp
Private mobjBakFtpConnection As New clsFtp

Private mblnUseInetFtp As Boolean

Private mobjFtp As TFtpDeviceInf        'ftp�豸��Ϣ
Private mobjBakFtp As TFtpDeviceInf     'ftp���ݴ洢�豸��Ϣ

Private dcmglbUID As New DicomGlobal    '����UIDRoot=1
Private mblnReadOnly As Boolean         '�Ƿ�ֻ�ܲ鿴True�鿴ģʽ��False�ɼ�ģʽ

Private mblnShowProcessBar As Boolean   '�Ƿ���ʾ��������
Private mstrScanDeviceTempDir As String 'ɨ���豸��ʱĿ¼
Private mblnShowImage As Boolean        '����ƶ�ʱ���Ƿ��Զ���ʾ��ͼ
Private mdblBigImgZoom As Double        '��ͼ�Ŵ���
Private mblnUnload As Boolean           '�Ƿ�����رմ���
Private mblnLocalizerBackward As Boolean    '��λƬ����
Private mblnChangeUser As Boolean       '�Ƿ��������û�����

'���˻�����Ϣ����
Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long

Private mlngAdviceID As Long            'ҽ��ID
Private mlngSendNo As Long
Private mblnMoved As Boolean            '�Ƿ�ת��
Private mlngStudyState As Long

Private mstrStudyUID As String          '���UID
Private mstrModality As String          'Ӱ�����
Private mstrSex As String               '�Ա�
Private mstrBirthDate As String         '��������
Private mstrAge As String               '����
Private mstrName As String              '����
Private mstrCheckNo As String           '����
Private mstrPatientID As String         '����ID
Private mstrInstitution As String       '��λ����


Private mstrAfterTag As String          '��̨�ɼ����
Private mstrAfterStudyUid As String     '��̨�ɼ����UID
Private mstrAfterSeriesUid As String    '��̨�ɼ�����UID
Private mstrAfterModality As String     '��̨�ɼ���Ӱ�����
Private mpanAfterInf As Pane            '��̨�ɼ���Ϣ��ʾ���
Private mlngAfterCurImageCount As Long  '��ǰ��̨�ɼ�ͼ������
Private mblnAfterIsUse As Boolean       '�Ƿ����ú�̨�ɼ�����
Private mstrAfterParentTitle As String

Private mSelectStudyInf As TUnLockStudyInf


'modify by tjh at 2010-01-20////////////////////////////////////////////

'Private pCurrentfrmCapture As frmVideoCapture    '��¼ӵ����ƵԴ�Ĳɼ�����
Private mVideoCapture As clsVideoCapture '��Ƶ�ɼ�����

Private mdblZoomRate As Double  '���ű��ʣ���cbrMain��cbrMain_ResizeClient�¼�����Ҫ���¼����ֵ��
Private mVideoSize As TVideoSize '��Ƶ��С������ص���Ƶ������棩
Private mCurCutRange As TCutRange '��Ƶ�ü���Χ���ã��ò���ͨ��GetString��SaveString������ע����У�
Private mVideoArea As TVideoArea  '��Ƶ�ͻ��������ã���cbrMain��cbrMain_ResizeClient�¼�����Ҫ���¼����ֵ��
Private mVideoDriverType As TVideoDriverType '��Ƶ�������ͣ��ò���ͨ��GetPara��SetPara���������ݿ��У�
Private mblnSoundHint As Boolean    '������ʾ
Private mblnPoputWindowHint As Boolean  '������ʾ

Private Const M_LNG_REFRESHINTERVAL As Long = 600 'ˢ�¼��

Private mstrVideoRegTime As String '������Ƶ����ע��ʱ��
Private mblnIsExecuteReg As Boolean '�ж��Ƿ�ִ��ע�����
Private mblnIsAllowStartupVideo As Boolean '�Ƿ�����������ƵԴ
Private mblnIsLockStudy As Boolean
Public mblnCurCaptureState As Boolean         '���浱ǰ�ɼ�״̬


Private mObjActiveMenuBar As CommandBars

Private mblnRefreshState As Boolean
Private mblnInitState As Boolean


Private Const CAPTURE_PARAMETER_CONFIG_FILE_NAME As String = "ZLVideoProcess.ini"
Private Const CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME As String = "\TempScan"  'Ĭ��ɨ���ļ���ʱ�洢·��
Private Const CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE As String = "Scan"  'Ĭ��ɨ���ļ���ʱ�洢·��

Private Type DlgFileInfo
    iCount As Long
    sPath As String
    sFIle() As String
End Type

Private Enum Dkp_ID
    Dkp_ID_Video = 1     '����б�
    Dkp_ID_Miniature      '��ǰ���˻�����Ϣ
End Enum


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'----------------------------------------------------------------------------------------------------------

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Property Get CaptionEx() As String
    CaptionEx = Me.Tag
End Property

Property Let CaptionEx(value As String)
    Dim hWndParent As Long
    
    Me.Tag = value
    
    hWndParent = GetParent(Me.hWnd)
    
    Call SetWindowText(hWndParent, Me.Tag)
End Property


'��ȡ��Ƶ�ɼ�����
Property Get videoCapture() As clsVideoCapture
    Set videoCapture = mVideoCapture
End Property


'��ȡ��Ƶ�ɼ����ڵĳ�ʼ��״̬
Property Get InitState() As Boolean
    InitState = mblnInitState
End Property

'��ȡ�˵��ӿڶ���
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property

Public Sub NotificationRefresh()
'֪ͨˢ��
    mblnRefreshState = False
End Sub


Private Sub Form_Initialize()
'��ʼ��ģ�����
    mblnInitState = False
End Sub


'�ӿ�ʵ�ֲ���*********************************************************************************

Public Function IWorkMenu_zlGetModuleMenuId() As Long
'��ȡӰ��˵��Ĳ˵�ID
    IWorkMenu_zlGetModuleMenuId = conMenu_Cap_Group
End Function


Public Function IWorkMenu_zlIsModuleMenu(ByVal objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'�жϲ˵��Ƿ����ڸ�ģ��˵�
    IWorkMenu_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
End Function


Public Sub IWorkMenu_zlCreateMenu(objMenuBar As Object)
'����Ӱ���¼��Ӧ�Ĳ˵�
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrControl As CommandBarControl
    
    Dim str3DFuncs() As String
    Dim i As Long
    Dim lng3DFunc As Long
    
    
    Set mObjActiveMenuBar = objMenuBar

    If Not HasMenu(objMenuBar, conMenu_Cap_Group) Then
        Set cbrMenuBar = mObjActiveMenuBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Cap_Group, "�ɼ�", 3, False)
        cbrMenuBar.ID = conMenu_Cap_Group
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_Dynamic, "��̬", "��ʾʵʱ��Ƶ", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_MarkMap, "�ɼ�", "�ɼ�ͼ��", 0, False)
            
            If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_After_Capture, "��̨�ɼ�", "��̨�ɼ����ͼ��", 10020, False)
            End If
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_Record, "¼��", "¼�Ƽ����Ƶͼ��", 0, True)
            
            If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_After_Record, "��̨¼��", "��̨¼�Ƽ����Ƶͼ��", 10021, False)
            End If
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_Record_Stop, "ֹͣ¼��", "ֹͣ��Ƶ¼��", 0, False): cbrControl.Enabled = False
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_RecordAudio, "¼��", "¼��", 0, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_Play, "����", "����¼�����¼��", 0, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_Import, "����", "�ļ�����", 10002, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_SaveAs, "���", "�ļ����", 3091, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_DelImg, "ɾͼ", "ɾ��ͼ��", 10001, False)
            
            If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_OpenStudyList, "�򿪼��", "�򿪼��", 0, True)
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_StudySyncState, "�������", "�������", 10012, False)
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_After_Tag, "��Ǽ��", "��Ǽ��", 10022, False)
            End If
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_DevSet, "��Ƶ����", "��Ƶ����", 815, True)
            
'            If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
'                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "�û�����", "�û�����", 3012, False)
'            End If
            
            '���ұ���ʾ�����ɼ���ť
            Set cbrControl = CreateModuleMenu(mObjActiveMenuBar.ActiveMenuBar.Controls, xtpControlButton, comMenu_Cap_Process, "�����ɼ�", "���������ɼ�����", 0, False)
            cbrControl.flags = xtpFlagRightAlign
        End With
    End If
End Sub


Public Sub IWorkMenu_zlCreateToolBar(objToolBar As Object)
'����������
'    Dim cbrControl As CommandBarControl
'
'    'ֻ����Ƶ�ɼ�վ������û���������
'    If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
'        If HasMenu(objToolBar, conMenu_Manage_ChangeUser) Then Exit Sub
'
'        Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "����", "�������ҽ���ͱ���ҽ��", 3012, True, 4)
'    End If
End Sub


Public Sub IWorkMenu_zlClearMenu()
'����������Ĳ˵�
'    Dim cbrControl As CommandBarControl
'
'    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Cap_Group)
'    If Not cbrControl Is Nothing Then Call cbrControl.Delete
End Sub


Public Sub IWorkMenu_zlClearToolBar()
'��������Ĺ�����
    Dim cbrControl As CommandBarControl
    
    Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Manage_ChangeUser)
    If Not cbrControl Is Nothing Then Call cbrControl.Delete
End Sub

Public Sub IWorkMenu_zlExecuteMenu(ByVal lngMenuId As Long)
'���ݲ˵�IDִ�ж�Ӧ����
    Dim objCbrControl As XtremeCommandBars.CommandBarControl
    
    Select Case lngMenuId
        Case conMenu_Cap_DevSet     '��Ƶ��������
            Call Menu_Cap_VideoConfig
            
'        Case conMenu_Manage_ChangeUser
'            Call SendMsgToMainWindow(Me, wetChangeUser, mlngAdviceID)
            
        Case comMenu_Cap_Process
            Call Menu_Manage_�����ɼ�(True)
            
        Case Else
            Set objCbrControl = Me.cbrMain.FindControl(, lngMenuId)
            
            If Not objCbrControl Is Nothing Then Call zlExecuteCommandBars(objCbrControl)
    End Select
End Sub


Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'���²˵�
    Select Case control.ID
        Case conMenu_Cap_DevSet
            control.Enabled = Me.Visible
            
'        Case conMenu_Manage_ChangeUser
'            control.Visible = mblnChangeUser
            
        Case Else
            Call zlUpdateCommandBars(control)
    End Select
End Sub


Public Sub IWorkMenu_zlPopupMenu(objPopup As XtremeCommandBars.ICommandBar)
'�����Ҽ��˵�
    Exit Sub
End Sub

Public Sub IWorkMenu_zlRefreshSubMenu(objMenuBar As Object)
'ˢ�µ������Ӳ˵�
    Exit Sub
End Sub

'*********************************************************************************************


Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'������ģ���ڵĲ˵�
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If
    
    CreateModuleMenu.ID = lngID '������ﲻָ��id�����ܽ���Щ�˵���ӵ��Ҽ��˵���
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG
End Function


Private Sub Menu_Cap_VideoConfig()
On Error GoTo errHandle
    frmVideoSetup.mlngModul = mlngModul
    frmVideoSetup.strRegName = "frmVideoCapture"    '����ע���Ľڵ�����
    frmVideoSetup.mstrPrivs = mstrPrivs
    frmVideoSetup.mlngCurDepartId = mlngCurDeptId
         
    Set frmVideoSetup.frmParent = Me
          
    'modify by tjh at 2010-01-20
    'frmVideoSetup.Show 1, Me
    
    Call frmWork_Video.SaveParameterCfg
    Call frmVideoSetup.ShowParameterConfig(frmWork_Video.videoCapture, Me)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Manage_�����ɼ�(Optional blnUnload As Boolean = True)
On Error GoTo errHandle

    If Not GetIsValidOfStorageDevice(mlngCurDeptId) Then
      MsgBoxD Me, "Ӱ��洢�豸δ�������ͣ�ã����飡", vbInformation, gstrSysName
      Exit Sub
    End If
    
    'Call frmVideoCapture.SetRestoreContainer(picVideoContainer)
    Call frmVideoDockWindow.Show
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub InitParameter()
'��ʼ����������
    Dim rsTmp As New ADODB.Recordset
    Dim intVideoCapture As Integer
    Dim strRegPath As String        'ע�������ı���·��

    mblnRealTime = True
    mintCurImgIndex = 0
    mblnPlayVideo = False
    mstrVideoRegTime = ""
    
    mblnAfterIsUse = False
    mstrAfterModality = "OT"
    mstrAfterParentTitle = ""
        
    '��������ڴ��̵ĸ�Ŀ¼��app.pathΪ��x:\��
    mstrBufferDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
    mstrAviFileName = mstrBufferDir & "TmpVideo.avi"
    
    mblnUnload = False
    mblnIsExecuteReg = False
        
    
    mstrInstitution = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    
    gint��Ƶ�豸���� = getLicenseCount(LOGIN_TYPE_��Ƶ�豸)
    '��ȡע�����Ϣ--���沼��
    strRegPath = "����ģ��\" & App.ProductName & "\frmVideoCapture"
    
    mblnUseClipbord = GetSetting("ZLSOFT", strRegPath, "UseClipbord", 0)
    Call SaveSetting("ZLSOFT", strRegPath, "UseClipbord", IIf(mblnUseClipbord, 1, 0))
    
    '��ȡ��������
    mVideoDriverType = zlDatabase.GetPara("��Ƶ��������", glngSys, mlngModul, "0")
    
    '��ȡ��ʾ����
    mblnSoundHint = zlDatabase.GetPara("�ɼ���������ʾ", glngSys, mlngModul, True)
    mblnPoputWindowHint = zlDatabase.GetPara("�ɼ��󵯴���ʾ", glngSys, mlngModul, True)
    
    '��ȡɨ���豸��ʱ�洢��ͼ��Ŀ¼
    mstrScanDeviceTempDir = GetSetting("ZLSOFT", strRegPath, "ɨ���豸��ʱĿ¼", "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA")
     
     
    '��ȡ�ü�����
    mCurCutRange.LeftRate = Val(GetSetting("ZLSOFT", strRegPath, "mdblX1Scale", 0))  'ʹ��mdblX1Scale������Ϊ�˱�֤����ǰ�Ĳ������ü���
    mCurCutRange.WidthRate = Val(GetSetting("ZLSOFT", strRegPath, "mdblX2Scale", 0))
    mCurCutRange.TopRate = Val(GetSetting("ZLSOFT", strRegPath, "mdblY1Scale", 0))
    mCurCutRange.HeightRate = Val(GetSetting("ZLSOFT", strRegPath, "mdblY2Scale", 0))

    If (mCurCutRange.LeftRate >= 1) Or (mCurCutRange.LeftRate < 0) Then mCurCutRange.LeftRate = 0
    If (mCurCutRange.WidthRate >= 1) Or (mCurCutRange.WidthRate < 0) Then mCurCutRange.WidthRate = 0
    If (mCurCutRange.TopRate >= 1) Or (mCurCutRange.TopRate < 0) Then mCurCutRange.TopRate = 0
    If (mCurCutRange.HeightRate >= 1) Or (mCurCutRange.HeightRate < 0) Then mCurCutRange.HeightRate = 0
  
    
    '��ȡ���ڵĲ���
    On Error GoTo continue1:
    mstrActiveType = zlDatabase.GetPara("��̤�˿�", glngSys, mlngModul, "1")
    If IsNumeric(mstrActiveType) Then
        mstrComPort = CLng(mstrActiveType)
        mstrActiveType = "COM"
        
        mintCapType = zlDatabase.GetPara("��̤�ɼ���ʽ", glngSys, mlngModul, "1")
        If mintCapType < 0 Or mintCapType > 2 Then
            mintCapType = 1
        End If
        '��ȡ��̤���ʱ��
        mintComInterval = zlDatabase.GetPara("��̤ʱ����", glngSys, mlngModul, "1")
    End If
continue1:

    
    '����ƶ�ʱ���Ƿ��Զ���ʾ��ͼ
     mblnShowImage = zlDatabase.GetPara("����ƶ�ʱ��ʾ��ͼ", glngSys, mlngModul, "0")
     mdblBigImgZoom = zlDatabase.GetPara("�ɼ���ͼ�Ŵ���", glngSys, mlngModul, "1")
     
     If mblnShowImage Then ucPreview.MouseMoveZoom = mdblBigImgZoom
     
     
    '����UIDRoot=1
    dcmglbUID.RegString("UIDRoot") = "1"
    
    '������Ƶ�ɼ������С�Ƿ������޸�
    intVideoCapture = Val(zlDatabase.GetPara("����ı�ɼ������С", glngSys, mlngModul, "1", , InStr(mstrPrivs, ";��������;") > 0))
    
    If intVideoCapture = 0 Then
    
        pbxSize.Item(0).MousePointer = 0
        pbxSize.Item(1).MousePointer = 0
        pbxSize.Item(2).MousePointer = 0
        pbxSize.Item(3).MousePointer = 0
    Else
    
        pbxSize.Item(0).MousePointer = 7
        pbxSize.Item(1).MousePointer = 7
        pbxSize.Item(2).MousePointer = 9
        pbxSize.Item(3).MousePointer = 9
    
    End If
    
    
    '��ʼ�����Ҽ�����==============================================================================
    mblnAfterIsUse = GetDeptPara(mlngCurDeptId, "���ú�̨�ɼ�", 0)
    mstrAfterModality = GetDeptPara(mlngCurDeptId, "��̨Ӱ�����", "OT")
    
    '��ȡ�����洢�豸��
    mobjFtp.strDeviceId = GetDeptPara(mlngCurDeptId, "�洢�豸��")
    mobjBakFtp.strDeviceId = GetDeptPara(mlngCurDeptId, "�����豸��")
    
    mblnLocalizerBackward = Val(GetDeptPara(mlngCurDeptId, "��λƬ����", 0))
    
'    mblnChangeUser = GetDeptPara(mlngCurDeptId, "�������û�", 0) = "1"              '�������û�
    
    '��ȡ���ߴ洢�豸��Ϣ
    gstrSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Tag, mobjFtp.strDeviceId)
    
    If rsTmp.EOF Then
        MsgBox "Ӱ��洢�豸δ�������ͣ�ã����飡", vbInformation, gstrSysName
        mobjFtp.strDeviceId = ""
        mblnReadOnly = True
        Exit Sub
    End If
    
    Call funGetFtpDeviceInf(Me, mobjFtp)
    
    '��ȡ�����豸��Ϣ
    If Val(mobjBakFtp.strDeviceId) > 0 Then
        gstrSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Tag, mobjBakFtp.strDeviceId)
        
        If rsTmp.EOF Then
            mobjBakFtp.strDeviceId = ""
            MsgBox "δȡ����Ч�ı����豸��Ϣ�����ܶԲɼ�ͼ����б��ݲ��������鱸���豸�����Ƿ���ȷ��", vbInformation, gstrSysName
            
            Exit Sub
        End If
        
        Call funGetFtpDeviceInf(Me, mobjBakFtp)
    End If
    
    
End Sub

Public Sub zlInitModule(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long, Optional owner As Object = Nothing)
'��ʼ��ģ�����
    mlngModul = lngModule
    mstrPrivs = strPrivs
    mlngCurDeptId = lngDepartId
    
    If Not owner Is Nothing Then Set mobjOwner = owner
    
    '���ô�����ʽ
    Call SetWindowStyle
    
    Call InitParameter
    
    Call OpenVideoCaptureDevice
    
    Call UpdateAfterCaptureInfo
    
    mblnInitState = True
End Sub




'��ʾ��Ƶ����
Public Sub ShowVideoWindow(ByRef objContainer As Object)
    Dim strRegPath As String
    
    If Not mIsShowing Then
        Call Me.Show
        
        mIsShowing = True
    End If
    
    If objContainer Is Nothing Then Exit Sub
    
    If Not mParentContainer Is Nothing Then
        Call SaveVideoAreaCfg(mParentContainer.Name)
    End If
    
    Set mParentContainer = objContainer
    Call SetParent(Me.hWnd, mParentContainer.hWnd)

    If Me.Height <> mParentContainer.Height Then
        Call LoadVideoAreaCfg(mParentContainer.Name)
    End If

    Call UpdateSize
    
    If TypeOf mParentContainer Is Form Then
        mParentContainer.Caption = Me.Tag
        mParentContainer.Icon = Me.Icon
        
        Me.Width = Me.Width - 140
        Me.Height = Me.Height - 140
    End If
End Sub


Public Sub HideVideoWindow()
'������Ƶ��ʾ����
    Me.Hide
    
    mIsShowing = False
End Sub


'���µ�ǰ��Ƶ���ڴ�С
Public Sub UpdateSize()
On Error GoTo errHandle
    If mParentContainer Is Nothing Then Exit Sub
    
    Me.Left = 0
    Me.Top = 0
    
    Me.Height = mParentContainer.Height
    Me.Width = mParentContainer.Width
    
    If TypeOf mParentContainer Is Form Then
        Me.Width = Me.Width - 140
        Me.Height = Me.Height - 500
    End If
errHandle:
End Sub


'���ûָ�ʱ����������
Public Sub SetRestoreContainer(ByRef objContainer As Object)
    Set mRestoreContainer = objContainer
    
'    mRestoreContainer.Visible = True
End Sub


'�ָ�ԭ�е���Ƶ��ʾ����
Public Sub RestoreContainer()
    If mRestoreContainer Is Nothing Then Exit Sub
    
    If Not mParentContainer Is Nothing Then
        Call SaveVideoAreaCfg(mParentContainer.Name)
    End If
    
    Set mParentContainer = mRestoreContainer
    Call SetParent(Me.hWnd, mRestoreContainer.hWnd)
    
    Me.Left = 0
    Me.Top = 0

    If Me.Height <> mRestoreContainer.Height Then
        '���Ӹ������ڻ��������ڻָ���Ƶ��ʾλ��ʱ�����´�ע����ȡ��Ƶ��ʾλ�ô�С
        Call LoadVideoAreaCfg(mRestoreContainer.Name)
    End If
    
    Me.Height = mRestoreContainer.Height
    Me.Width = mRestoreContainer.Width
    
    
    If TypeOf mRestoreContainer Is Form Then
        mRestoreContainer.Caption = Me.Tag
        mRestoreContainer.Icon = Me.Icon
    End If
End Sub


Property Get ParentContainerObj() As Object
    Set ParentContainerObj = mParentContainer
End Property

Property Set ParentContainerObj(value As Object)
    Set mParentContainer = value
End Property



Property Get RestoreContainerObj() As Object
    Set RestoreContainerObj = mRestoreContainer
End Property

Property Set RestoreContainerObj(value As Object)
    Set mRestoreContainer = value
End Property



Property Get IsLockStudy() As Boolean
    IsLockStudy = mblnIsLockStudy
End Property



Property Get LockPatientName() As String
    LockPatientName = mstrInfor
End Property



'----------------------------------------------------------------------------------------------------------
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'����Ƿ�����������ƵԴ
'����ƵԴû����������ʱ���򲻽���ע�ᣬҲ�������ж�
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckVideoReg(ByVal blnIsStartupVideo As Boolean) As Boolean
  '������Ƶ�����ɹ�������Ҫ����ע��
  
    mblnIsExecuteReg = True
  
    mstrVideoRegTime = FunLogIn(Me, LOGIN_TYPE_��Ƶ�豸)
  
    CheckVideoReg = mstrVideoRegTime <> ""
End Function


Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal lngStudyState As Long, ByVal blnMoved As Boolean)
'����ҽ����Ϣ
    Dim rsTemp As ADODB.Recordset
    
    '����������ĵ�ǰ�����Ϣ
    mSelectStudyInf.lngAdviceID = lngAdviceID
    mSelectStudyInf.blnMoved = blnMoved
    mSelectStudyInf.lngSendNO = lngSendNO
    mSelectStudyInf.lngStudyState = lngStudyState
    
    If mblnIsLockStudy Then Exit Sub
    
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    mblnMoved = blnMoved
    mlngStudyState = lngStudyState
    mblnReadOnly = False
    mblnRefreshState = True
    
    '���ݱ�ת��ʱ��û��Ȩ��ʱ��״̬Ϊָ��״̬ʱ����ģ��Ϊֻ��
    If mlngAdviceID <= 0 Or blnMoved Or lngStudyState = 6 Or lngStudyState = 0 Or lngStudyState = 1 Or InStr(mstrPrivs, "��Ƶ�ɼ�") <= 0 Then
        mblnReadOnly = True
    End If
    
    '��ȡ���˻�����Ϣ,дDICOM����ʱ��
    gstrSQL = "Select /*+Rule */ A.Ӱ�����,A.����,A.�Ա�,A.����,A.��������,A.����,A.����,A.���UID,B.����ID " & _
                " From Ӱ�����¼ A,����ҽ����¼��B " & _
                " Where A.ҽ��ID=[1] And A.ҽ��ID=B.Id"
                
    If mblnMoved Then
        gstrSQL = Replace(gstrSQL, "Ӱ�����¼", "HӰ�����¼")
        gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˻�����Ϣ", lngAdviceID)
    
    If Not rsTemp.EOF Then
        mstrStudyUID = Nvl(rsTemp("���UID"))
        mstrModality = Nvl(rsTemp("Ӱ�����"))
        mstrInfor = Nvl(rsTemp("����"))
        mstrSex = Nvl(rsTemp("�Ա�"))
        mstrAge = Nvl(rsTemp("����"))
        mstrBirthDate = Nvl(rsTemp("��������"))
        mstrName = Nvl(rsTemp("����"))
        mstrCheckNo = Nvl(rsTemp("����"))
        mstrPatientID = Nvl(rsTemp("����ID"))
        
        If mstrSex = "��" Then
            mstrSex = "M"
        ElseIf mstrSex = "Ů" Then
            mstrSex = "F"
        Else
            mstrSex = "O"
        End If
    Else
        mstrStudyUID = ""
        mstrModality = ""
        mstrInfor = ""
        mstrSex = ""
        mstrAge = ""
        mstrCheckNo = ""
        mstrPatientID = ""
        mstrBirthDate = ""
        mstrName = ""
    End If
    
    Me.Tag = "ͼ��ɼ�" & IIf(mstrInfor <> "", "(" & mstrInfor & ")", "")
    Me.CaptionEx = Me.Tag
End Sub


Private Sub LockStudy()
'�������
    mblnIsLockStudy = True
End Sub


Private Sub UnLockStudy()
'�������
    mblnIsLockStudy = False
End Sub


Public Sub zlRefreshFace(Optional blnForceRefresh As Boolean = False)
'ˢ�½���
    Dim rsTemp As ADODB.Recordset
    Dim iRows As Integer
    Dim iCols As Integer
    Dim strStudyUID As String
    
    On Error GoTo errHandle
    
    
    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh Then Exit Sub
    
    mlngTmpAdviceId = mlngAdviceID
    mlngTmpSendNo = mlngSendNo
    mblnRefreshState = True

    Call ConfigVideoShowState(True)
    
'    '��ȡ���˻�����Ϣ,дDICOM����ʱ��
'    gstrSQL = "Select A.���UID From Ӱ�����¼ A,����ҽ����¼��B  Where A.ҽ��ID=B.Id and A.ҽ��ID=[1] "
'
'    If mblnMoved Then
'        gstrSQL = Replace(gstrSQL, "Ӱ�����¼", "HӰ�����¼")
'        gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
'
'        ucPreview.Enable = False
'    End If
'
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˻�����Ϣ", mlngAdviceID)
'
'    If rsTemp.RecordCount <= 0 Then
'        strStudyUID = ""
'    Else
'        strStudyUID = Nvl(rsTemp!���uid)
'    End If


    Call ucPreview.RefreshImage(slStudy, mstrStudyUID, mblnMoved, blnForceRefresh, False)
    
    If ucPreview.ImgViewer.Images.Count > 0 Then
                    
        '����ѡ��ͼ��װ�ص�dcmView��
        dcmView.Images.Clear
        dcmView.Images.Add ucPreview.ImgViewer.Images(ucPreview.SelectIndex)
        
        Dim dblTempZoom As Double
              
        dblTempZoom = dcmView.CurrentImage.ActualZoom
        dcmView.CurrentImage.StretchToFit = False
        
        
        '�жϵ����븡������ʱ�����ű��ʲ���С��0.1
        If dblTempZoom < 0.1 Then
            dblTempZoom = 0.1
        End If
        
              
        Call subCenterZoom(dcmView.CurrentImage, dcmView, dblTempZoom, mCorpSize)
        
        '�����Twain�ɼ�ģʽ��������mblnRealTimeΪfalse
        If IsTwainCaptureWay = True Then mblnRealTime = False

        '��ʾdcmView������picVideo
        dcmView.CurrentImage.BorderWidth = 0
    Else
        Call dcmView.Images.Clear
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub




Private Sub zlStopCapture()
'-----------------------------------------------------------------------------------------
'���ܣ�ֹͣ��ʾ��Ƶ�ɼ�,�ͷ���Ƶ�ɼ����ڣ�
'      �ͷŴ��������Ķ˿�
'��������
'���أ���
'-----------------------------------------------------------------------------------------
    '�ͷŲɼ��豸������
    If Not mVideoCapture Is Nothing Then Call mVideoCapture.StopPreview
    
    '�ر�COMM��
    If commListener.PortOpen Then
        commListener.PortOpen = False
    End If
    
    '����Midi�ӿ����������¼����
    If Not mclsDxDevice Is Nothing Then
        If mclsDxDevice.Handle <> 0 Then Call mclsDxDevice.CloseDxDevice
    End If
    
'    Call ucCapHook.FreeHook
End Sub





Public Sub zlUpdateCommandBars(control As XtremeCommandBars.CommandBarControl)
'ֻ��Ӱ��ɼ�����վ�ž߱���̨�ɼ�����

'���ݵ�ǰ״̬ȷ���˵��Ŀ��ӺͿɲ���

    '���û�г�ʼ����Ƶ��������Ƶ��صİ�ť��������ʹ��
    If mVideoCapture Is Nothing Then
        control.Enabled = False
        Exit Sub
    End If
    
    Select Case control.ID
        Case conMenu_Cap_Dynamic       '��̬��ʾ
            control.Checked = mblnRealTime
            control.Enabled = (Not mblnReadOnly) And Not IsTwainCaptureWay And mVideoCapture.IsStartup ' And (mhCapWnd <> 0) modify by tjh at 2010-01-20
            control.Visible = Not IsTwainCaptureWay
            
            If mblnRealTime Then
                control.IconId = conMenu_Cap_Dynamic
            Else
                control.IconId = 10023
            End If
            
        Case conMenu_Cap_MarkMap       'Ӱ��ɼ�
            control.Enabled = Not mblnReadOnly And (mVideoCapture.IsStartup Or IsTwainCaptureWay) And Not mblnCurCaptureState
            
        Case conMenu_Cap_After_Capture  '��̨�ɼ�
            control.Enabled = mVideoCapture.IsStartup And Not mblnCurCaptureState
            control.Visible = mblnAfterIsUse And mlngModul = G_LNG_VIDEOSTATION_MODULE
            
        Case conMenu_Cap_Import        'Ӱ����
            control.Enabled = Not mblnReadOnly
            
        Case conMenu_Cap_DelImg  'Ӱ��ɾ��
            control.Enabled = (mblnRealTime = False) And (ucPreview.ImgViewer.Images.Count > 0) And (Not mblnReadOnly) And Me.Visible
            
        Case conMenu_Cap_Record        '¼��
            control.Enabled = Not mblnReadOnly And mVideoDriverType = vdtWDM And mVideoCapture.IsStartup
            control.Visible = Not IsTwainCaptureWay
            
        Case conMenu_Cap_After_Record   '��̨¼��
            control.Enabled = mVideoDriverType = vdtWDM And mVideoCapture.IsStartup
            control.Visible = Not IsTwainCaptureWay And mblnAfterIsUse And False
            
        Case conMenu_Cap_Record_Stop 'ֹͣ¼�� modify by tjh at 2010-01-22
            control.Enabled = mblnRealTime And Not mblnReadOnly And (mVideoDriverType = vdtWDM) And mVideoCapture.IsStartup
            control.Visible = Not IsTwainCaptureWay
            
        Case conMenu_Cap_RecordAudio '¼��
            control.Enabled = Not mblnReadOnly
            
'        Case conMenu_Cap_Full_Screen 'ȫ�� modify by tjh at 2010-01-22 (���ʹ���µ���Ƶ�ط��������������øù���)
'            control.Enabled = mblnRealTime And (Not mblnReadOnly) And Not GetIsTwainCaptureWay And mVideoCapture.IsStartup
'            control.Visible = Not GetIsTwainCaptureWay And mstrVideoRegTime <> ""
'
        Case conMenu_Cap_DevSet        '���ã�������ڸ���״̬ʱ�������θð�ť�� modify by tjh at 2010-01-25
            control.Enabled = mblnIsAllowStartupVideo   'mblnEmbedded ' And (Not mblnReadOnly)
            
            '���Ϊ�������壬�����ظ����ð�ť
            'control.Visible = mstrVideoRegTime <> ""
            If Not (mParentContainer Is Nothing) Then
                If TypeOf mParentContainer Is frmVideoDockWindow Then
                    control.Enabled = False
                Else
                    control.Enabled = True
                End If
            End If
            
        '¼�񲥷�,¼��ֹͣ,¼����,¼�����,����¼��
        Case conMenu_Cap_Play, conMenu_Cap_Stop, conMenu_Cap_Forward, _
             conMenu_Cap_Back
            If (mblnRealTime = False) And (dcmView.Images.Count > 0) Then
                control.Visible = dcmView.Images(1).Tag.Tag = VIDEOTAG Or dcmView.Images(1).Tag.Tag = AUDIOTAG
                control.Enabled = dcmView.Images(1).Tag.Tag = VIDEOTAG Or dcmView.Images(1).Tag.Tag = AUDIOTAG
            Else
                control.Visible = False
                control.Enabled = False
            End If
            
        Case conMenu_Cap_SaveAs
            control.Enabled = Me.Visible
            
         '���ȶԱȶ�,����,�ü�����,˳ʱ����ת,��ʱ����ת,��,ƽ��,�߼�����
        Case conMenu_Process_Window, conMenu_Process_Zoom, conMenu_Process_RectZoom, conMenu_Process_RRotate, _
             conMenu_Process_LRotate, conMenu_Process_Sharpness, conMenu_Process_Filter, conMenu_Process_Corp

            control.Enabled = (mblnRealTime = False)
        '��ͷ��ע,Բ�α�ע,���ֱ�ע,
        Case conMenu_Process_Arrow, conMenu_Process_Ellipse, conMenu_Process_Text
            control.Enabled = (mblnRealTime = False)
            
        Case conMenu_Tool_Analyse
            If mblnObserve Then
                control.Enabled = Not mblnReadOnly
            Else
                control.Visible = False
                control.Enabled = False
            End If
            
            
        Case conMenu_Cap_OpenStudyList
            control.Enabled = True
            control.Visible = IIf(mlngModul = G_LNG_VIDEOSTATION_MODULE, True, False)
            
        Case conMenu_Cap_StudySyncState
            control.Enabled = Not mblnReadOnly Or mblnIsLockStudy
            control.Visible = IIf(mlngModul = G_LNG_VIDEOSTATION_MODULE, True, False)
            
        Case conMenu_Cap_After_Tag
            control.Enabled = mVideoCapture.IsStartup
            control.Visible = mblnAfterIsUse And mlngModul = G_LNG_VIDEOSTATION_MODULE
    End Select
End Sub


''''''''''''''''''''''''''''''''''
'ɨ��ͼ��
''''''''''''''''''''''''''''''''''
Private Sub ScanImages()
  'ע��ʧ����ִ�иù���
  If mstrVideoRegTime = "" Then
    Exit Sub
  End If
                
  'ɾ����������ʱ�洢��ͼ��Ŀ¼
  On Error GoTo continue
    If Dir(mstrTempDirOfScan, vbDirectory) <> "" Then
      Call mdlDir.DeleteFolder(mstrTempDirOfScan)
    End If
continue:
      
  If Dir(mstrTempDirOfScan, vbDirectory) = "" Then
    Call MkDir(mstrTempDirOfScan)
  End If
  
  'ɾ��twain�豸��ʱ�洢��Ŀ¼
  On Error GoTo continue1
    If Dir(mstrScanDeviceTempDir, vbDirectory) <> "" Then
      Call mdlDir.DeleteFolder(mstrScanDeviceTempDir)
    End If
continue1:

  If Dir(mstrScanDeviceTempDir, vbDirectory) = "" Then
    Call MkDir(mstrScanDeviceTempDir)
  End If
  
  mintScanImageIndex = 0

  '����ɨ�����ļ���������
  ImageScanner.FileType = BMP_Bitmap
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
    Call ImageScanner.CloseScanner

    MsgBox err.Description
End Sub


Public Sub CaptureImage()
'************************************************************
'
'����Ƶ����¼���вɼ�ͼ��
'
'************************************************************
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    If mstrVideoRegTime = "" Then   '���û��ע�ᣬ������ɼ�
        MsgboxEx Me, "δ��⵽��Ч��ע����Ϣ�����ܽ���ͼ��ɼ�������", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    If Not (Not mblnReadOnly And (mVideoCapture.IsStartup Or IsTwainCaptureWay)) Then Exit Sub  '���Ϊֻ����������Ƶû��������������ɼ�
    
    '�ɼ�ͼ��ʱ��������Ǻ�̨�ɼ��������жϵ�ǰ���ص�ͼ�������ݿ��е�ͼ���¼���Ƿ�һ�£������һ�£�˵���ü�鵱ǰ�������������豸վ��ɼ�
    strSql = "select count(*) as ͼ���� from Ӱ����ͼ�� where ����uid in(select ����UID from Ӱ�������� where ���UID=(select ���UID from Ӱ�����¼ where ҽ��id=[1])) "
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯͼ������", mlngAdviceID)
    
    If rsData.RecordCount > 0 Then
        If Val(Nvl(rsData!ͼ����)) <> ucPreview.ImageTotal Then
            Call MsgBoxD(Me, "��⵽��ǰ���ص�ͼ�����������ݿ��¼����һ�£�����������û��Ըü����вɼ���������ˢ�º����ԡ�", vbInformation + vbOKOnly, "��ʾ")
            Exit Sub
        End If
    End If
            
    If IsTwainCaptureWay Then
      Call ScanImages  'ͨ��TWAIN�ӿڲɼ�ͼ��
    Else
        If mblnRealTime Then 'Ϊʵʱ��ʾʱ�Զ���ʵʱͼ
            Call subCaptureImg(True)
        Else
            Call subCaptureImg(MsgBoxD(Me, "ȷ��Ҫ�ɼ���ǰ��̬ͼ��ѡ������ɼ��豸ʵʱͼ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo)
        End If
    End If
End Sub



Public Sub CaptureAfterImage()
'��̨ͼ��ɼ�
    If mstrVideoRegTime = "" Then   '���û��ע�ᣬ������ɼ�
        MsgboxEx Me, "δ��⵽��Ч��ע����Ϣ�����ܽ���ͼ��ɼ�������", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    If Not mVideoCapture.IsStartup Then Exit Sub  '���Ϊֻ����������Ƶû��������������ɼ�,twain��ʽ�������̨�ɼ�
    
    Call subCaptureImg(True, "", Nothing, True)
    
End Sub


Public Sub zlExecuteCommandBars(control As XtremeCommandBars.CommandBarControl)
  On Error GoTo errHandle
    Select Case control.ID
        Case conMenu_Cap_Dynamic       '��̬��ʾ
            If IsTwainCaptureWay Then
              Call MsgBoxD(Me, "TWAIN�ɼ�ģʽ�£����ܽ��ж�̬��Ƶ����ʾ��", vbOKOnly, "��ʾ")
            Else
              Call ConfigVideoShowState(True)
            End If
            
        Case conMenu_Cap_MarkMap       'Ӱ��ɼ�
            Call CaptureImage
            
        Case conMenu_Cap_After_Capture
            Call CaptureAfterImage
            
        Case conMenu_Cap_Import        'Ӱ����
            Call InputImageFile
            
        Case conMenu_Cap_DelImg  'Ӱ��ɾ��
            Call subDeleteImage
            
        Case conMenu_Cap_Record        '¼��
            If mstrVideoRegTime = "" Then
                MsgboxEx Me, "δ��⵽��Ч��ע����Ϣ�����ܽ���¼�������", vbOKOnly, "��ʾ"
                Exit Sub
            End If
            
            Call subVideoSave
            
        Case conMenu_Cap_Record_Stop  'ֹͣ¼�� modify by tjh at 2010-01-22
            If mstrVideoRegTime = "" Then
                'MsgboxEx Me, "δ��⵽��Ч��ע����Ϣ�����ܽ���¼�������", vbOKOnly, "��ʾ"
                Exit Sub
            End If
            
            Call subStopVideo
            
        Case conMenu_Cap_RecordAudio    '¼��
            If mstrVideoRegTime = "" Then
                MsgboxEx Me, "δ��⵽��Ч��ע����Ϣ�����ܽ���¼��������", vbOKOnly, "��ʾ"
                Exit Sub
            End If
            
            Call frmRecordAudio.ShowRecordAudio(Me)
            
'        Case conMenu_Cap_Full_Screen 'ȫ�� modify by tjh at 2010-01-22
'            Call subFullCall
            
'        Case conMenu_Cap_DevSet        '����
'            Call SaveParameterCfg
'            Call subVideoSetup
            
        Case conMenu_Cap_Play          '¼�񲥷�
            Call subVideoPlay
'
        Case conMenu_Cap_SaveAs        '�ļ����
            Call subVideoSaveAs
            
'        Case conMenu_Process_Cursor
'            subSetMouseState 0
'            control.Checked = True
                
        Case conMenu_Process_Window         '���ȶԱȶ�
            subSetMouseState 1
            control.Checked = True
            
        Case conMenu_Process_Zoom           '����
            subSetMouseState 2
            control.Checked = True
            
        Case conMenu_Process_RectZoom       '�ü�����
            subSetMouseState 3
            control.Checked = True
        
        Case conMenu_Process_RectCapture         '�ü���ɼ�
            Call CaptureFrameSelectImage
            
        Case conMenu_Process_RRotate        '˳ʱ����ת
            subSetRotate True
            
        Case conMenu_Process_LRotate        '��ʱ����ת
            subSetRotate False
            
        Case conMenu_Process_Sharpness      '��
            subSetSharp True
            
        Case conMenu_Process_Filter         'ƽ��
            subSetSharp False
            
        Case conMenu_Process_Corp          '�϶�
           subSetMouseState 14
           control.Checked = True
            
        Case conMenu_Process_Arrow          '��ͷ��ע
            subSetMouseState 11
            control.Checked = True
            
        Case conMenu_Process_Ellipse        'Բ�α�ע
            subSetMouseState 12
            control.Checked = True
            
        Case conMenu_Process_Text           '���ֱ�ע
            subSetMouseState 13
            control.Checked = True
        Case conMenu_Tool_Analyse           '�߼�����
            Call OpenViewer(1, pobjPacsCore, mlngAdviceID, False, Me, "", mblnMoved, mblnLocalizerBackward)
            
        Case conMenu_Cap_OpenStudyList      '�򿪼��ɼ�ͼ��
            Call OpenStudy
            
        Case conMenu_Cap_StudySyncState     '�������
            If control.IconId = 10012 Then
                control.IconId = 8123
                
                Call LockStudy
                
                Call SendMsgToMainWindow(Me, wetLockStudy, mlngAdviceID, mstrInfor)
            Else
                control.IconId = 10012
                
                Call UnLockStudy
                
                If mlngAdviceID <> mSelectStudyInf.lngAdviceID Then
                    Call zlUpdateAdviceInf(mSelectStudyInf.lngAdviceID, mSelectStudyInf.lngSendNO, mSelectStudyInf.lngStudyState, mSelectStudyInf.blnMoved)
                    Call zlRefreshFace
                End If
                
                Call SendMsgToMainWindow(Me, wetUnLockStudy, mlngAdviceID, mstrInfor)
            End If
        Case conMenu_Cap_After_Tag      '���º�̨�ɼ����
            If mstrVideoRegTime = "" Then
                MsgboxEx Me, "δ��⵽��Ч��ע����Ϣ�����ܽ��б�ǣ�", vbOKOnly, "��ʾ"
                Exit Sub
            End If
            
            Call UpdateAfterCaptureInfo
    End Select
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume Next
End Sub


Private Sub OpenStudy()
    Dim cbrControl As CommandBarControl
    
    Dim lngCurAdviceId As Long
    Dim lngSendNO As Long
    Dim lngStudyState As Long
    Dim blnMoved As Boolean
    Dim blnResult As Boolean
    
    blnResult = mobjOwner.OpenPatiListWind(lngCurAdviceId, lngSendNO, lngStudyState, blnMoved)
        
    If lngCurAdviceId > 0 Then
        '��ʼ���µļ����вɼ�
        Call UnLockStudy
        
        Call zlUpdateAdviceInf(lngCurAdviceId, lngSendNO, lngStudyState, blnMoved)
        Call zlRefreshFace
        
        Call LockStudy
                
        '�޸�����״̬
        Set cbrControl = cbrMain.FindControl(, conMenu_Cap_StudySyncState)
        cbrControl.IconId = 8123
       
        '�������˸ı��¼�
        Call SendMsgToMainWindow(Me, wetLockStudy, mlngAdviceID, mstrInfor)
    End If
End Sub


Public Sub zlUnloadMe()
    mblnUnload = True
    Unload Me
End Sub


Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    zlExecuteCommandBars control
End Sub



Private Sub cbrMain_Resize()
    If cbrMain.Count >= 3 Then
        If cbrMain.Item(3).Visible <> mblnShowProcessBar Then
            mblnShowProcessBar = cbrMain.Item(3).Visible
        End If
    End If
    
    If cbrMain.Count >= 3 Then
        If picCapture.Width < 4000 Then
            cbrMain.Item(2).Position = xtpBarTop
            cbrMain.Item(3).Position = xtpBarBottom
        Else
            cbrMain.Item(2).Position = xtpBarLeft
            cbrMain.Item(3).Position = xtpBarRight
        End If
    End If
End Sub

'modify by tjh at 2010-01-19
'ͨ���÷����������ű��ʺ���Ƶ����ʾ��Χ
Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
  
    mVideoArea.Height = Bottom - Top - 4 * pbxSize(0).Height
    mVideoArea.Width = Right - Left - 4 * pbxSize(2).Width
    mVideoArea.Left = Left
    mVideoArea.Top = Top
    
    '�������ű���
    Call CalcVideoZoomRate

    '������Ƶ��ʾ��Χ
    Call ConfigVideoDisplay(wdmCapture)
    Call ConfigVideoDisplay(picVideo)
    
    'ˢ����Ƶ��ʾ
    If Not (mVideoCapture Is Nothing) Then
        Call mVideoCapture.RefreshVideoWindow
    End If
    
    'ˢ��DcmView�е�ͼ����ʾλ��
    If dcmView.Images.Count > 0 Then
        Call subCenterZoom(dcmView.Images(1), dcmView, dcmView.Images(1).ActualZoom, mCorpSize)
    End If

    'ˢ�²ü�����λ��
    Call RefreshPbxSizePos
        
    
    If IsTwainCaptureWay Then
      
        '����ͼ�����ʾ��Χ
        dcmView.Left = Left
        dcmView.Top = Top
        dcmView.Width = Right - Left
        dcmView.Height = Bottom - Top
  
        'ˢ��DcmView��ͼ�����ʾλ��
        If dcmView.Images.Count > 0 Then
            Call subCenterZoom(dcmView.Images(1), dcmView, dcmView.Images(1).ActualZoom, mCorpSize)
        End If
    
    End If
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    zlUpdateCommandBars control
End Sub


Private Sub commListener_OnComm()
On Error GoTo errHandle
    Dim strInput As String
    
    '�����TWAINɨ�裬��֧�ֽ�̤���زɼ�
    If IsTwainCaptureWay Then Exit Sub
    
    If mstrActiveType <> "COM" Then Exit Sub
    
    strInput = ""
    If commListener.InBufferCount > 0 Then strInput = commListener.Input
    
    If Not (commListener.CommEvent = comEvCTS Or commListener.CommEvent = comEvDSR _
        Or commListener.CommEvent = comEvCD Or commListener.CommEvent = comEvRing Or strInput <> "" _
        Or commListener.CommEvent = comEvSend Or commListener.CommEvent = comEvReceive) Then Exit Sub
    
    If mintCapType = 1 Then 'ת������
        If mintComState <> commListener.CommEvent Then
           '����ۼ�ʱ�䳬���˲�ͼʱ��������ɼ�ͼ��
           If mlngComTime > mintComInterval Then
               'If Me.cbrMain.FindControl(, conMenu_Cap_MarkMap).Enabled Then
               If Not mblnReadOnly Then
                    Call subCaptureImg(True)
               End If
           End If
           
           '��¼�µ�COM״̬����ʱ�����㣬����timer
           mintComState = commListener.CommEvent
           mlngComTime = 0
           tmrComm.Enabled = True
        End If
    ElseIf mintCapType = 0 Then   'ֱ�Ӵ���
        '���β��½�̤��ʱ������������3��
        If DateDiff("S", mdtLastCapture, time) < mintComInterval Then
            mdtLastCapture = time
            Exit Sub
        End If
        
        mdtLastCapture = time
        
        If Not mblnReadOnly Then
            Call subCaptureImg(True)
        End If
    Else    '��ƽ����
        '���ڵ�ƽ����������������½�̤��ʱ�򣬶�Ӧ�ߵĵ�ƽ����֣���-��-�ͣ��򣨸�-��-�ߣ��ı仯
        'ͨ����ƽ�仯������ȷ���Ƿ���˽�̤��
        '�����ֵ�������ʱ����Ȼ�����OnComm�¼������ǵ�ƽ���ᷢ���仯��
        'ͨ���жϵ�ǰ��ƽ����̬��ƽ�Ƿ���ͬ��ȷ����ƽ�Ƿ����˱仯��
        
        '�жϵ�ƽ�Ƿ�ı䣬�ж�CTS��
        If mblnCTSHolding <> commListener.CTSHolding Then
            '�����񵴣�ë�������ж����δ�����ʱ���Ƿ�С���趨�ļ��
            If DateDiff("S", mdtLastCapture, time) < mintComInterval Then
                mdtLastCapture = time
                Exit Sub
            End If
            
            mdtLastCapture = time
            
            If Not mblnReadOnly Then
                Call subCaptureImg(True)
            End If
        End If
    End If
errHandle:
End Sub


Private Sub dcmView_DblClick()
On Error GoTo errHandle
    Call subVideoPlay
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


'modify by tjh at 2010-01-20
'������Ƶ���ű���
Private Sub CalcVideoZoomRate()
  If mVideoSize.Width = 0 Or mVideoSize.Height = 0 Then
    mdblZoomRate = 1
    Exit Sub
  End If
  
    
  If (mVideoArea.Height <= 0) Or (mVideoArea.Width <= 0) Then
    mdblZoomRate = 1
    Exit Sub
  End If
  
  '�������ű���
  If (mVideoArea.Height / mVideoArea.Width) > (mVideoSize.Height / mVideoSize.Width) Then
    mdblZoomRate = mVideoArea.Width / mVideoSize.Width
  Else
    mdblZoomRate = mVideoArea.Height / mVideoSize.Height
  End If
  
End Sub


'modify by tjh at 2010-01-20
'������Ƶ��ʾ
Private Sub ConfigVideoDisplay(videoObj As Object)
  '�߿��С
  Const DICOM_VIEWER_BODER_SIZE As Long = 5
  
  If mVideoSize.Width = 0 Or mVideoSize.Height = 0 Then
    Exit Sub
  End If
  
  If (mVideoArea.Height <= 0) Or (mVideoArea.Width <= 0) Then
    Exit Sub
  End If
  
  '������Ƶ��ʾ��Χ
  videoObj.Top = 0 - mdblZoomRate * mVideoSize.Height * mCurCutRange.TopRate
  videoObj.Height = mdblZoomRate * mVideoSize.Height
  picView.Height = mdblZoomRate * mVideoSize.Height * (1 - mCurCutRange.TopRate - mCurCutRange.HeightRate)
  
  videoObj.Left = 0 - mdblZoomRate * mVideoSize.Width * mCurCutRange.LeftRate
  videoObj.Width = mdblZoomRate * mVideoSize.Width
  picView.Width = mdblZoomRate * mVideoSize.Width * (1 - mCurCutRange.LeftRate - mCurCutRange.WidthRate)
  
  picView.Left = mVideoArea.Left + (mVideoArea.Width - picView.Width - 2 * pbxSize(2).Width) / 2 + 3 * pbxSize(2).Width
  picView.Top = mVideoArea.Top + (mVideoArea.Height - picView.Height - 2 * pbxSize(0).Height) / 2 + 3 * pbxSize(0).Height
  
  '����DICOM��ʾͼ��Ĵ�С
  dcmView.Left = DICOM_VIEWER_BODER_SIZE
  dcmView.Top = DICOM_VIEWER_BODER_SIZE
  dcmView.Width = picView.Width - DICOM_VIEWER_BODER_SIZE
  dcmView.Height = picView.Height - DICOM_VIEWER_BODER_SIZE
End Sub


Private Sub ConfigTwainDisplay()
  '�߿��С
  Const DICOM_VIEWER_BODER_SIZE As Long = 5
  
  dcmView.Left = DICOM_VIEWER_BODER_SIZE
  dcmView.Top = DICOM_VIEWER_BODER_SIZE
  dcmView.Width = picView.Width - DICOM_VIEWER_BODER_SIZE
  dcmView.Height = picView.Height - DICOM_VIEWER_BODER_SIZE
End Sub


Private Sub SetWindowStyle()
    Dim lngWindowStyle As Long
    
    lngWindowStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
    lngWindowStyle = lngWindowStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
    
    Call SetWindowLong(Me.hWnd, GWL_STYLE, lngWindowStyle Or WS_CHILD)
End Sub

Private Sub OpenVideoCaptureDevice()
'����Ƶ�ɼ��豸
    Dim blnIsStartupVideo As Boolean

    If mVideoCapture Is Nothing Then
        '������Ƶ�ɼ�����
        Set mVideoCapture = New clsVideoCapture
        
        '������Ƶ������
        Call mVideoCapture.ConnectedVfwDeviceObj(picVideo)
        Call mVideoCapture.ConnectedWdmDeviceObj(wdmCapture)
        
        '��ȡ�����ļ�
        Call mVideoCapture.ReadCaptureParameterFromFile(App.Path & "\" & CAPTURE_PARAMETER_CONFIG_FILE_NAME)
    
        '������Ƶ����ʾģʽ
        Call mVideoCapture.SetVideoShowWay(swStretch)
    
        '�ڶ�ȡ�ļ����ú��޸ĸ����ԣ�ֻ�����ø����ԣ����ܸ��������߿���е��ں���ʾ��
        wdmCapture.AppHandle = Me.hWnd
        wdmCapture.IsShowState = False
        
        mdblZoomRate = 1
    Else
        Call zlStopCapture
    End If

    
    '������Ƶ��������
    mVideoCapture.VideoDriverType = mVideoDriverType

    '��ȡ��Ƶ��С
    mVideoSize.Width = ScaleX(mVideoCapture.VideoSize.Width, vbPixels, vbTwips)
    mVideoSize.Height = ScaleX(mVideoCapture.VideoSize.Height, vbPixels, vbTwips)
    
    '���ý���
    Call CaptureSwitchFace(IsTwainCaptureWay)
    
    mblnIsAllowStartupVideo = FunCheckRegInfo(Me)
    
    '�ж��Ƿ�����������ƵԴ********************************
    If Not mblnIsAllowStartupVideo Then
      mVideoCapture.IsAllowStartupVideo = False
      
      '������������ʱ������twain�Ĳ�������
      mVideoDriverType = vdtTWAIN
      mVideoCapture.VideoDriverType = vdtTWAIN
      '���ý���
      Call CaptureSwitchFace(IsTwainCaptureWay)
      
      Exit Sub
    End If
    '*******************************************************
    
    
    '��ʼ��ƵԤ��********************************************
    If Not IsTwainCaptureWay Then
        mblnRealTime = True
        
        Call mVideoCapture.StartPreview
                
        blnIsStartupVideo = mVideoCapture.IsStartup
    Else
        mblnRealTime = False
        
        blnIsStartupVideo = ImageScanner.ScannerAvailable
    End If
    
    'ע�Ტ�ж��Ƿ���������������Ƶ����������ֹͣ��Ƶ��ʾ
    If Not CheckVideoReg(blnIsStartupVideo) Then
        Call mVideoCapture.StopPreview
        
        If mblnIsExecuteReg Then
            mVideoCapture.IsAllowStartupVideo = False
        End If
    Else
        Call OpenComm(False) '�򿪲ɼ��˿�
    End If
    
    'ע��ʧ�ܺ�������ʾ���棬����twain�Ĳ�������
    '*****************************************************
    '�÷����ɲɼ��������ô��ڵ���
    '�������ע������Ϊ���ܳ��ֲ������ò��ԣ�����Ӳ����������Ƶ�������������û�ж�ϵͳ����ע�ᣬ������ֹ����޷�ʹ��
    '������Ƶ�����������ú��п�����������Ѿ�����ȷ�޸ģ�������Ҫ���½���ע�ᣬ������ع���
    '*****************************************************
    If mstrVideoRegTime = "" Then
      mVideoDriverType = vdtTWAIN
      mVideoCapture.VideoDriverType = vdtTWAIN
      '���ý���
      Call CaptureSwitchFace(IsTwainCaptureWay)
    End If
    '*********************************************************
    
'    If mVideoCapture.IsStartup Then Call ucCapHook.EnableHook
End Sub


Private Sub UpdateAfterCaptureInfo()
'���º�̨�ɼ���Ϣ
    
    'ֻ��Ӱ��ɼ�ģ�鲢�����ú��̨�ɼ�����ʹ�ú�̨�ɼ�
    If mlngModul = G_LNG_VIDEOSTATION_MODULE And Not IsTwainCaptureWay And mblnAfterIsUse And mVideoCapture.IsAllowStartupVideo Then
        Call CreateNewCaptureTag
        Call ShowAfterCaptureInf
    End If
End Sub


Private Sub Form_Load()
  On Error GoTo errHandle
    Dim strRegPath As String
    '���ô�����ʽ
    Call SetWindowStyle
    
    '�÷�����show֮��Żᴥ��
    mIsShowing = False
        
    
    InitCommandBars
    
    Call ucPreview.InitImgPreview(gcnOracle)
        
    strRegPath = "����ģ��\" & App.ProductName & "\frmVideoCapture"
    ucPreview.PageImgCount = Val(GetSetting("ZLSOFT", strRegPath, "�ɼ�����ͼ����", 5))
    
    mstrTempDirOfScan = App.Path + CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME
    If Len(mstrTempDirOfScan) > 45 Then
        Dim strFolder As String
        Dim pathlen As Long
        
        strFolder = String(256, 0)
        pathlen = GetTempPath(256, strFolder)
        If pathlen > 0 Then
            mstrTempDirOfScan = Left(strFolder, pathlen - 1) + CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME
        End If
    End If
    
    
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


'�����Ƿ�ΪTWAIN�Ĳɼ���ʽ
Private Function IsTwainCaptureWay() As Boolean
  IsTwainCaptureWay = IIf(mVideoDriverType = vdtTWAIN, True, False)
End Function


'����TWAINʱ�Ĳɼ�����
Private Sub CaptureSwitchFace(ByVal blnUseTwain As Boolean)
    'ȥ����TWAINɨ�費��ص�һЩ��������
    
    dcmView.Visible = blnUseTwain
    picView.Visible = Not blnUseTwain
    
    pbxSize(0).Visible = Not blnUseTwain
    pbxSize(1).Visible = Not blnUseTwain
    pbxSize(2).Visible = Not blnUseTwain
    pbxSize(3).Visible = Not blnUseTwain
        
    wdmCapture.Visible = False
    picVideo.Visible = False
      
    If blnUseTwain Then
      Set dcmView.Container = picCapture
      Set txtInputText.Container = picCapture
    Else
      Set dcmView.Container = picView
      Set txtInputText.Container = picView
    End If
    
    Call ConfigVideoShowState(True)
    
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    Call cbrMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    Call cbrMain_ResizeClient(lngLeft, lngTop, lngRight, lngBottom)
End Sub


'���²ɼ�����������
Public Sub UpdateCaptureDirver(ByVal videoDirver As TVideoDriverType)
    '���ע��ʧ�ܣ���������������͸���
   If mstrVideoRegTime = "" And mblnIsExecuteReg Then
       Exit Sub
   End If
 
    '��ֹͣ��Ƶ��Ԥ��
    Call mVideoCapture.StopPreview
    
    mVideoDriverType = videoDirver
    mVideoCapture.VideoDriverType = videoDirver
       
    '��ȡ��Ƶ��С
    mVideoSize.Width = ScaleX(mVideoCapture.VideoSize.Width, vbPixels, vbTwips)
    mVideoSize.Height = ScaleX(mVideoCapture.VideoSize.Height, vbPixels, vbTwips)
       
    Call CaptureSwitchFace(videoDirver = vdtTWAIN)
        
    
    '�������Twain�ɼ���ʽ������������Ԥ��
    If videoDirver <> vdtTWAIN Then
      mblnRealTime = True
      
      '��ʼԤ��
      Call mVideoCapture.StartPreview
      
    Else
      mblnRealTime = False
    End If
End Sub


Public Sub SaveVideoAreaCfg(ByVal strAreaName As String)
'������Ƶ�ɼ���������
  Dim strRegPath As String
  
  '����ע������
  strRegPath = "����ģ��\" & App.ProductName & "\" & strAreaName
  SaveSetting "ZLSOFT", strRegPath, "CY1", picCapture.Height
End Sub


Public Sub LoadVideoAreaCfg(ByVal strAreaName As String)
'������Ƶ�ɼ���������
    Dim strRegPath As String
     
    strRegPath = "����ģ��\" & App.ProductName & "\" & strAreaName
    picCapture.Height = Val(GetSetting("ZLSOFT", strRegPath, "CY1", picCapture.Height))
End Sub


'���浱ǰ��������
Public Sub SaveParameterCfg()
  Dim strRegPath As String
  
  '����ע������
  strRegPath = "����ģ��\" & App.ProductName & "\frmVideoCapture"
    
  '�ü���������
  SaveSetting "ZLSOFT", strRegPath, "mdblX1Scale", mCurCutRange.LeftRate
  SaveSetting "ZLSOFT", strRegPath, "mdblX2Scale", mCurCutRange.WidthRate
  SaveSetting "ZLSOFT", strRegPath, "mdblY1Scale", mCurCutRange.TopRate
  SaveSetting "ZLSOFT", strRegPath, "mdblY2Scale", mCurCutRange.HeightRate
  
  
  '��ʾ��������
  SaveSetting "ZLSOFT", strRegPath, "��ʾ��������", mblnShowProcessBar
    
        
  '����ɼ�����
  If Not mVideoCapture Is Nothing Then Call mVideoCapture.SaveCaptureParameterToFile(App.Path & "\" & CAPTURE_PARAMETER_CONFIG_FILE_NAME)
End Sub


Private Sub OpenComm(blnForce As Boolean)
    
    On Error GoTo err
    
    If mstrActiveType = "��" Then Exit Sub
    
    If mstrActiveType = "COM" Then
        
        If commListener.PortOpen Then Exit Sub

        commListener.CommPort = mstrComPort
        commListener.Settings = "9600,N,8,1"
        commListener.InputMode = comInputModeText
        commListener.RThreshold = 1
        commListener.InBufferCount = 0
        commListener.InputLen = 0
        commListener.RTSEnable = True
                        
        commListener.PortOpen = True
            
        '��¼��̬��ƽ��λ
        mblnCTSHolding = commListener.CTSHolding
        
    Else
        
        If mclsDxDevice Is Nothing Then Set mclsDxDevice = New clsDxHidDevice
        
        '��DX�豸
        Call mclsDxDevice.OpenDxDevice(mstrActiveType)
        
        tmrComm.Enabled = True
        tmrComm.Interval = 2
    End If
    
    Exit Sub
err:
    MsgBox "�˿ڴ򿪴���", vbOKOnly, "��ʾ"
End Sub


Private Sub dcmView_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = 1 And dcmView.Images.Count > 0 Then
        Dim intLabelType As Integer
        
        If mintMouseState = 13 And txtInputText.Text <> "" And txtInputText.Visible Then
            If Not mdcmSelectLabel Is Nothing Then mdcmSelectLabel.Text = txtInputText.Text
        End If
        
        mMouseDownPoint.X = dcmView.Images(1).ActualScrollX
        mMouseDownPoint.Y = dcmView.Images(1).ActualScrollY
          
        mInitScrollPoint.X = dcmView.Images(1).ScrollX + X
        mInitScrollPoint.Y = dcmView.Images(1).ScrollY + Y
        
        mblnDcmViewDown = True
        If mintMouseState <> 0 Then
            '��¼��ǰ���λ��
            mlngBaseX = X
            mlngBaseY = Y
            
            Select Case mintMouseState
                'Case 14  'ͼ���϶�
                
                Case 11, 12, 13, 3    '��ͷ����Բ������,��ѡ
                    If mintMouseState = 11 Then
                        intLabelType = doLabelArrow
                    ElseIf mintMouseState = 12 Then
                        intLabelType = doLabelEllipse
                    ElseIf mintMouseState = 13 Then
                        intLabelType = doLabelText
                    ElseIf mintMouseState = 3 Then
                        intLabelType = doLabelRectangle
                    End If
                    
                    dcmView.Images(1).Labels.Add GetNewLabel(intLabelType, dcmView.ImageXPosition(X, Y), dcmView.ImageYPosition(X, Y), 0, 0)
                    
                    Set mdcmSelectLabel = dcmView.Images(1).Labels(dcmView.Images(1).Labels.Count)
                    
                    mdcmSelectLabel.LineWidth = 2
            End Select
        End If
    End If
End Sub


Private Sub dcmView_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mblnDcmViewDown = True And Button = 1 And dcmView.Images.Count > 0 Then
        Select Case mintMouseState
            Case 1  '���ȶԱȶ�
                dcmView.Images(1).Width = dcmView.Images(1).Width + (X - mlngBaseX)
                dcmView.Images(1).Level = dcmView.Images(1).Level + (Y - mlngBaseY)
                mlngBaseX = X
                mlngBaseY = Y
            Case 2  '����
                Dim dblZoom As Double
                dblZoom = dcmView.Images(1).ActualZoom
                dblZoom = dblZoom * (1 + (Y - mlngBaseY) * 0.001)
                If dblZoom < 64 And dblZoom > 0.01 Then
                    subCenterZoom dcmView.Images(1), dcmView, dblZoom, mCorpSize
                End If
                mlngBaseY = Y
'            Case 3  '�ü�����
'                Dim dcmLabel As DicomLabel
'                dcmView.Labels.Clear
'                Set dcmLabel = dcmView.Labels.AddNew
'                dcmLabel.LabelType = doLabelRectangle
'                dcmLabel.Left = mlngBaseX
'                dcmLabel.Top = mlngBaseY
'                dcmLabel.Width = x - mlngBaseX
'                dcmLabel.Height = y - mlngBaseY
            Case 11, 12, 3 '��ͷ��ע'Բ�α�ע,��ѡ
                mdcmSelectLabel.Width = dcmView.ImageXPosition(X, Y) - mdcmSelectLabel.Left
                mdcmSelectLabel.Height = dcmView.ImageYPosition(X, Y) - mdcmSelectLabel.Top
            Case 14
                '�϶�ͼ��......
                dcmView.Images(1).ScrollX = mInitScrollPoint.X - X
                dcmView.Images(1).ScrollY = mInitScrollPoint.Y - Y
        End Select
        dcmView.Refresh
    End If
End Sub


Private Sub dcmView_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mblnDcmViewDown = True And Button = 1 And dcmView.Images.Count > 0 Then
        mblnDcmViewDown = False
        If mintMouseState = 13 Then     '���ֱ�ע
            txtInputText.Left = X * Screen.TwipsPerPixelX
            txtInputText.Top = Y * Screen.TwipsPerPixelY
            txtInputText.Text = ""
            txtInputText.Visible = True
            txtInputText.SetFocus
        ElseIf mintMouseState = 3 Then  '�ü�����
            
            '��ʾͼ�񱣴�˵�
            Call ShowFrameSelectImagePopup
            'ɾ����ѡ�õ���ʱ��ע
            If dcmView.Images(1).Labels.Count > 0 Then
                dcmView.Images(1).Labels.Remove dcmView.Images(1).Labels.Count
            End If
            
            Set mdcmSelectLabel = Nothing
            
            
'            dcmView.Labels.Clear
            
'            dcmView.Labels.Clear
'            RectangleZoom dcmView, dcmView.Images(1), mlngBaseX, mlngBaseY, x - mlngBaseX, y - mlngBaseY
        ElseIf mintMouseState = 14 Then
            '����ͼ�����ε�ƫ��λ��
            mCorpSize.X = mCorpSize.X + (dcmView.Images(1).ActualScrollX - mMouseDownPoint.X)
            mCorpSize.Y = mCorpSize.Y + (dcmView.Images(1).ActualScrollY - mMouseDownPoint.Y)
        End If
        
        dcmView.Refresh
    End If
End Sub


Private Sub RectangleZoom(Viewer As DicomViewer, img As DicomImage, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long)
    Dim newZoom As Double
    Dim dblRatio As Double
    Dim sX As Long
    Dim sY As Long
    Dim oldZoom As Double
    
    If lngWidth > 0 And lngHeight > 0 Then
        oldZoom = img.ActualZoom
        sX = img.ActualScrollX
        sY = img.ActualScrollY
        
        img.StretchToFit = False
        
        dblRatio = Viewer.Width / Screen.TwipsPerPixelX / lngWidth
        If dblRatio > Viewer.Height / Screen.TwipsPerPixelY / lngHeight Then
            dblRatio = Viewer.Height / Screen.TwipsPerPixelY / lngHeight
        End If
        
        newZoom = oldZoom * dblRatio
        img.Zoom = newZoom
        
        img.ScrollX = (sX + lngLeft) * dblRatio
        img.ScrollY = (sY + lngTop) * dblRatio
    End If
End Sub


Private Sub subCenterZoom(img As DicomImage, Viewer As DicomViewer, dblZoom As Double, corpSize As TPoint)
'------------------------------------------------
'���ܣ���ͼ��������š��Ե�ǰviewer���ĵ�Ϊ�������ĵ㡣
'������
'       img -- �������ŵ�ͼ��
'       viewer ���� ͼ�����ڵ�viewer
'       dblZoom ����ͼ���µ����ű���
'���أ��ޣ�ֱ�ӵ���ͼ������ű���
'�ϼ���������̣�frmViewer.Viewer_MouseMove
'�¼���������̣���
'���õ��ⲿ��������
'�����ˣ� �ƽ� 2006-2-10
'------------------------------------------------
    img.Zoom = dblZoom
    img.StretchToFit = False

            
    img.ScrollX = (img.SizeX * img.ActualZoom - ScaleX(Viewer.Width, vbTwips, vbPixels) / Viewer.MultiColumns) / 2 + corpSize.X
    img.ScrollY = (img.SizeY * img.ActualZoom - ScaleY(Viewer.Height, vbTwips, vbPixels) / Viewer.MultiRows) / 2 + corpSize.Y
End Sub


Public Function GetNewLabel(lType As Integer, lLeft As Integer, lTop As Integer, lWidth As Integer, lHeight As Integer) As DicomLabel
'------------------------------------------------
'���ܣ�����һ��LABEL���󣬲���������ʼ����
'������lType--��ע�����ͣ�lLeft--��ע��Leftֵ��lTop--��ע��Topֵ��lWidth--��ע��Widthֵ��lHeight--��ע��Heightֵ��
'���أ������ɵı�ע��
'�����ˣ��ƽ�
'------------------------------------------------
    Dim l As New DicomLabel
    l.LabelType = lType
    l.XOR = True
    l.ImageTied = True
    l.Left = lLeft
    l.Top = lTop
    l.Width = lWidth
    l.Height = lHeight
    l.Margin = 0
    l.AutoSize = True
    l.FontSize = 12
    l.LineWidth = 1
    If l.LabelType = 0 Then     '����
        l.Transparent = False
        l.Width = 200
        l.Height = 10
    End If
    Set GetNewLabel = l
End Function
   
   
Public Sub subCaptureImg(ByVal RealTimeCap As Boolean, Optional ByVal strFileName As String = "", _
    Optional ByRef picCapture As StdPicture = Nothing, Optional ByVal blnIsAfterCapture As Boolean = False)
'------------------------------------------------
'���ܣ��ɼ����洢ͼ��
'��������
'���أ��ޣ�ֱ�ӱ����²ɼ���ͼ��
'------------------------------------------------
On Error GoTo errHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    
    '
    If mblnUseInetFtp Then
        If mblnCurCaptureState Then Exit Sub
        
        mblnCurCaptureState = True
    End If
    
    If funCaptureSingleImage(RealTimeCap, strFileName, picCapture, blnIsAfterCapture) = True Then
        If blnIsAfterCapture Then
            '����Ǻ�̨�ɼ������̨�ɼ��ɹ���ɾ����̨�ɼ���ͼ��
            If subSaveAfterCaptureImage Then Call dcmAfter.Images.Clear
            
            Call ShowAfterCaptureInf
            
            mblnCurCaptureState = False
            Exit Sub
        End If
        
        Call subSaveImage
        
        '����Ӱ����״̬������ɼ���һ��ͼ����ԭ����״̬���ѱ��������޸ĳ��Ѽ��
        If ucPreview.ImgViewer.Images.Count = 1 Then
            
            If mlngStudyState < 3 Then
                strSql = "Zl_Ӱ����_State(" & mlngAdviceID & "," & mlngSendNo & ",3,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCurDeptId & ")"
                zlDatabase.ExecuteProcedure strSql, "�ɼ���һ��ͼ��"
            End If
        End If
        
        
        If ucPreview.ImgViewer.Images.Count = 1 Then
            Call SendMsgToMainWindow(Me, wetCaptureFirstImg, mlngAdviceID, mstrStudyUID)
        Else
            Call SendMsgToMainWindow(Me, wetUpdateImg, mlngAdviceID, mstrStudyUID)
        End If
    End If
    
    
    mblnCurCaptureState = False
Exit Sub
errHandle:
    mblnCurCaptureState = False
    err.Raise err.Number, err.Description
End Sub

Private Function CopyPictureToDicomImg(ByVal lngHDC As Long, ByVal lngPictureHandle As Long, objDcmImg As Object) As Boolean
    Const bitCount As Long = 3
'congpicture�и���ͼ��dicomimage
    Dim hBitmap As OLE_HANDLE
    Dim stucbmp As TBitMap
    Dim lngSize As Long
    Dim lngResult As Long
    Dim aryPixels() As Byte
    Dim stuDipInf As BITMAPINFO
    
    Dim i As Long, j As Long, bytTemp As Byte
    
    
    CopyPictureToDicomImg = False
    hBitmap = lngPictureHandle
    
    'ʹ��GetObject��������ȡ32λ�ĸ�ʽͷ��Ϣ
    lngResult = GetObject(hBitmap, Len(stucbmp), stucbmp)
    If lngResult = 0 Then Exit Function
    
    
    While stucbmp.bmWidth * 3 Mod 4 <> 0
        '��ʹ��GetDIBits����ʱ��ÿһ��������ֽ���������4�ı���������4�ֽڶ���
        stucbmp.bmWidth = stucbmp.bmWidth - 1
    Wend
    
    '����24λͼ���ʽ����ͼ��Ĵ洢��С�����ֽ�Ϊ��λ
    lngSize = stucbmp.bmWidth * 3 * stucbmp.bmHeight 'stucbmp.bmWidthBytes * stucbmp.bmHeight
    
    stuDipInf.bmiHeader.biSize = 40
    stuDipInf.bmiHeader.biHeight = -stucbmp.bmHeight
    stuDipInf.bmiHeader.biPlanes = stucbmp.bmPlanes
    stuDipInf.bmiHeader.biBitCount = 24 'bmp.bmBitsPixel  'ǿ��ʹ��24λ��ʽ�����ں�������ת��
    stuDipInf.bmiHeader.biWidth = stucbmp.bmWidth
    stuDipInf.bmiHeader.biCompression = BI_RGB
    stuDipInf.bmiHeader.biSizeImage = lngSize
    stuDipInf.bmiHeader.biXPelsPerMeter = 0
    stuDipInf.bmiHeader.biYPelsPerMeter = 0
    stuDipInf.bmiHeader.biClrUsed = 0
    stuDipInf.bmiHeader.biClrImportant = 0
    stuDipInf.bmiColors(0).rgbBlue = 8
    stuDipInf.bmiColors(0).rgbGreen = 8
    stuDipInf.bmiColors(0).rgbRed = 8
    stuDipInf.bmiColors(0).rgbReserved = 0
    
'    ReDim aryPixels(1 To stucbmp.bmWidthBytes, 1 To stucbmp.bmHeight, 1 To 1)
    ReDim aryPixels(1 To stucbmp.bmWidth * 3, 1 To stucbmp.bmHeight, 1 To 1)
    
    

'    lngResult = GetBitmapBits(hBitmap, lngSize, aryPixels(1, 1, 1))

    'ֻ��ʹ�øú�����ȡ24λ�����ظ�ʽ�����ʹ��GetBitmapBits����ȡ�Ľ���32λ��ͼ���ʽ
    lngResult = GetDIBits(lngHDC, hBitmap, 0, stucbmp.bmHeight, aryPixels(1, 1, 1), stuDipInf, DIB_RGB_COLORS)
    If lngResult = 0 Then Exit Function
    

    '��bmp��brg�洢��ʽת��Ϊdicom��rgb�洢��ʽ
    For i = 1 To stucbmp.bmWidth * 3 Step 3
        For j = 1 To stucbmp.bmHeight
            bytTemp = aryPixels(i + 2, j, 1)
            aryPixels(i + 2, j, 1) = aryPixels(i, j, 1)
            aryPixels(i, j, 1) = bytTemp
        Next j
    Next i

   
    '����dicom��ͼ���ʽ
    objDcmImg.Attributes.Add &H28, &H2, 3       'stucbmp.bmBitsPixel        'samples per pixel
    objDcmImg.Attributes.Add &H28, &H4, "RGB"                  'Photometric Interpretation
    objDcmImg.Attributes.Add &H28, &H6, 0                      'planar configuration
    objDcmImg.Attributes.Add &H28, &H100, 8                    'Bits Allocated
    objDcmImg.Attributes.Add &H28, &H101, 8                    'Bits Stored
    objDcmImg.Attributes.Add &H28, &H102, 7                    'High Bit
    objDcmImg.Attributes.Add &H28, &H103, 0                    'Pixel Representation
    objDcmImg.Attributes.Add &H28, &H10, stucbmp.bmHeight          'rows
    objDcmImg.Attributes.Add &H28, &H11, stucbmp.bmWidth           'columns
    
    objDcmImg.Pixels = aryPixels

    CopyPictureToDicomImg = True
End Function


Private Function funCaptureSingleImage(ByVal RealTimeCap As Boolean, _
    Optional ByVal strFileName As String = "", Optional ByRef picCapture As StdPicture = Nothing, _
    Optional ByVal blnIsAfterCapture As Boolean = False) As Boolean
'------------------------------------------------
'���ܣ��ɼ���֡��Ƶͼ�񣬽�ͼ��ת����DICOM��ʽ������дDICOM�ļ�ͷ�����ͼ���������ͼdcmMiniature�С�
'��������
'���أ��ޣ�ֱ�ӽ��²ɼ���ͼ�����dcmMiniature��
'------------------------------------------------
'�ɼ���֡ͼ��
On Error GoTo SaveFileError
    Dim ImgTmpImage As New DicomImage

    
    '�ɼ�ͼ�񣬷�Ϊֱ����Ƶ�ɼ��Ͳ���¼��ɼ�

    If Not (picCapture Is Nothing) Then
        Set picTemp2.Picture = Nothing
        Set picTemp2.Picture = picCapture
        
    ElseIf Trim(strFileName) <> "" And Dir(strFileName) <> "" Then
        'ʹ��dcmView��ʾ����ͼƬ������Ҫ�ٲü�
        Set picTemp2.Picture = Nothing
        Set picTemp2.Picture = LoadPicture(strFileName)
        
    Else
        If RealTimeCap = False And mblnPlayVideo = False Then
            'ʹ��dcmView��ʾ����ͼƬ������Ҫ�ٲü�
            Set picTemp2.Picture = Nothing
            
            If dcmView.Images.Count > 0 Then
                Set picTemp2.Picture = dcmView.CurrentImage.Capture(False).Picture
            End If
        Else
            '������ʵʱ��Ƶ��ʾʱ����Ҫ��ͼ����вü�����
            Set picTemp2.Picture = Nothing
                        
            'modify by tjh at 2009-01-20
            Dim curPic As StdPicture
            Set curPic = mVideoCapture.CaptureImageFromMemory

            If curPic Is Nothing Then
                Call MsgBoxD(Me, "��Ƶͼ��ɼ�ʧ�ܣ�������Ƶ���������Ƿ���ȷ(����Ƶ�豸����ʾģʽ��)��", vbOKOnly, "��ʾ")
                
                funCaptureSingleImage = False
                Exit Function
            End If
            
            picTemp2.Width = mVideoSize.Width * (1 - mCurCutRange.WidthRate - mCurCutRange.LeftRate)
            picTemp2.Height = mVideoSize.Height * (1 - mCurCutRange.HeightRate - mCurCutRange.TopRate)

            Call picTemp2.PaintPicture(curPic, 0, 0, picTemp2.Width, picTemp2.Height, _
                                       mVideoSize.Width * mCurCutRange.LeftRate, mVideoSize.Height * mCurCutRange.TopRate, _
                                       picTemp2.Width, picTemp2.Height, vbSrcCopy)
                                               
            picTemp2.Picture = picTemp2.Image

            Set curPic = Nothing
        End If
    End If
    
    
    '��ͼ���ٴ��ύ�����а�
    If picTemp2.Picture Is Nothing Then
        funCaptureSingleImage = False
        Exit Function
    End If
  

    If mblnUseClipbord Then
        'ʹ�ü����巽ʽ
        Call Clipboard.SetData(picTemp2.Picture, 2)
        '�Ӽ��а�ȡ��ͼ��
        Call ImgTmpImage.Paste
        
        Call Clipboard.Clear
    Else
        '��ʹ�ü����巽ʽ����Picture�и���ͼ��ImgTmpImage��,��ʹ�ü����彻������
        If Not CopyPictureToDicomImg(picTemp2.hdc, picTemp2.Image.Handle, ImgTmpImage) Then
            funCaptureSingleImage = False
            Exit Function
        End If
    End If
    
    
    '��дͼ�������DICOM����
    Call subWriteDicomPara(ImgTmpImage, mlngAdviceID, blnIsAfterCapture)
    
    Dim dcmTag As New clsImageTagInf
    dcmTag.Tag = IMGTAG
    
    Set ImgTmpImage.Tag = dcmTag
    
    If blnIsAfterCapture Then
        Call dcmAfter.Images.Add(ImgTmpImage)
    Else
        '��ͼ���������ͼ��
        Call subInsert2Mini(ImgTmpImage)
    End If
    
    
    funCaptureSingleImage = True
    
    Exit Function
SaveFileError:
    If ErrCenter() = 1 Then Resume
End Function


Private Sub subWriteDicomPara(img As DicomImage, lngAdviceID As Long, _
    Optional blnIsAfterCapture As Boolean = False)
'------------------------------------------------
'���ܣ��������ͼ����дDICOM�ļ�ͷ��Ϣ
'������img���������DICOM�ļ�,lngAdviceID����ҽ��ID
'���أ��ޣ�ֱ���ļ�ͷ��Ϣд��img���ļ�ͷ
'------------------------------------------------
    Dim curDate As Date

    curDate = zlDatabase.Currentdate
    
    If blnIsAfterCapture Then
        img.Attributes.Add &H10, &H10, ""                           'Name ����
        img.Attributes.Add &H10, &H20, ""                           'Patient ID ����ID
        img.Attributes.Add &H10, &H30, ""                           'BirthDate ����
        img.Attributes.Add &H10, &H40, ""                           'Sex �Ա�
        img.Attributes.Add &H10, &H1010, ""                         'Age ����
        img.Attributes.Add &H10, &H4000, ""                         'Patient Comment ����ע��
        img.Attributes.Add &H20, &H10, ""                           'Study ID ���ID
        img.Attributes.Add &H8, &H60, mstrModality                   'Modality Ӱ�����
    Else
        img.Attributes.Add &H10, &H10, mstrName                     'Name ����
        img.Attributes.Add &H10, &H20, mstrPatientID                'Patient ID ����ID
        img.Attributes.Add &H10, &H30, mstrBirthDate                'BirthDate ����
        img.Attributes.Add &H10, &H40, mstrSex                      'Sex �Ա�
        img.Attributes.Add &H10, &H1010, mstrAge                    'Age ����
        img.Attributes.Add &H10, &H4000, ""                         'Patient Comment ����ע��
        img.Attributes.Add &H20, &H10, mstrCheckNo                  'Study ID ���ID
        img.Attributes.Add &H8, &H60, mstrModality                   'Modality Ӱ�����
    End If
    
    img.Attributes.Add &H8, &H8, ""                             'ImageType  ��
    img.Attributes.Add &H8, &H16, "1.2.840.10008.5.1.4.1.1.7"   'SOP Class  UID�����β�׽
    img.Attributes.Add &H8, &H20, Format(curDate, "yyyy-mm-dd")     'Study Date �������
    img.Attributes.Add &H8, &H21, Format(curDate, "yyyy-mm-dd")     'Series Date ��������
    img.Attributes.Add &H8, &H22, Format(curDate, "yyyy-mm-dd")     'Acquisition Date �ɼ�����
    img.Attributes.Add &H8, &H23, Format(curDate, "yyyy-mm-dd")     'Image Date   ͼ������
    img.Attributes.Add &H8, &H30, Format(curDate, "HH24:MI:SS")     'Study Time   ���ʱ��
    img.Attributes.Add &H8, &H31, Format(curDate, "HH24:MI:SS")     'Series Time  ����ʱ��
    img.Attributes.Add &H8, &H32, Format(curDate, "HH24:MI:SS")     'Acquisition Time  �ɼ�ʱ��
    img.Attributes.Add &H8, &H33, Format(curDate, "HH24:MI:SS")     'Image Time  ͼ��ʱ��
    img.Attributes.Add &H8, &H50, ""                            'Accession Number ��
    img.Attributes.Add &H8, &H70, "ZLSOFT"                      'Manufacturer ����
    img.Attributes.Add &H8, &H80, mstrInstitution                'Institution Name ��λ����
    img.Attributes.Add &H8, &H90, ""                            'Referring Physician's Name ��
    img.Attributes.Add &H8, &H1030, ""                          'Study Description ������� ��
    img.Attributes.Add &H20, &H11, "1"                          'Series Number ���к�
    img.Attributes.Add &H20, &H13, "1"                          'ImageNumber ͼ���
    img.Attributes.Add &H20, &H20, ""                           'Orientation ��
End Sub


Private Sub subInsert2Mini(img As DicomImage)
'------------------------------------------------
'���ܣ���ͼ����ӵ�����ͼdcmMiniature��
'������img���������DICOMͼ��
'      blnIsLocalImg���Ϊtrue,���ʾΪ��Ƶ
'���أ��ޣ�ֱ�ӽ�ͼ����ӵ�����ͼdcmMiniature��
'------------------------------------------------
    
    '�������Ƶ,������Ƶ������������UID
    '��������ͼ�ļ��UID������UID���޸�img��ֵ
    Call subUniteUID(img, img.Tag.Tag <> VIDEOTAG And img.Tag.Tag <> AUDIOTAG)
    
    ucPreview.AddImage img
End Sub


Private Sub Form_Resize()
On Error GoTo errHandle
    picDock.Left = 0
    picDock.Top = 0
    picDock.Width = Me.ScaleWidth
    picDock.Height = Me.ScaleHeight

    Call ucSplitter1.RePaint(False)
    
errHandle:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String

    
    'ж����Ƶע��
    Call FunLogOut(Me, LOGIN_TYPE_��Ƶ�豸, mstrVideoRegTime)

    '�ȹرղɼ����ں�COMM��
    Call zlStopCapture
  
    '���ֲü�����
    Call SaveParameterCfg
    
    '������Ƶ�ɼ���������
    If Not mRestoreContainer Is Nothing Then
        Call SaveVideoAreaCfg(mRestoreContainer.Name)
    End If

    strRegPath = "����ģ��\" & App.ProductName & "\frmVideoCapture"
    Call SaveSetting("ZLSOFT", strRegPath, "�ɼ�����ͼ����", ucPreview.PageImgCount)
    
'    Call mobjInetFtp.QuitFtp
    
    Set mclsDxDevice = Nothing
    Set mVideoCapture = Nothing
    Set mParentContainer = Nothing
    Set mRestoreContainer = Nothing
    Set mobjOwner = Nothing
End Sub


Private Sub subDeleteImage()
'------------------------------------------------
'���ܣ�ɾ������ͼ�б�ѡ�е�ͼ���ȴ����ݿ���ɾ����Ȼ���FTP��ɾ����ɾ���󴥷�StateChanged�¼�
'��������
'���أ��ޣ�ֱ��ɾ������ͼ�����һ��ͼ��
'------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If ucPreview.ImgViewer.Images.Count > 0 Then
        
        Dim blnResult As Boolean
                 

        '�����ݿ��FTP��ɾ������ͼ�б�ѡ�е�ͼ��
        blnResult = DeleteImages(Me, 1, ucPreview.SelectImage.InstanceUID, "")
        
        If blnResult = True Then    'ɾ���ɹ������޸�����ͼ״̬��������StateChanged�¼�
            '������ͼ��ɾ��ͼ��
            Call ucPreview.DeleteImage(ucPreview.SelectIndex)
            dcmView.Images.Clear

            If Not ucPreview.SelectImage Is Nothing Then
                dcmView.Images.Add ucPreview.SelectImage
            End If
            
            
            '����Ӱ����״̬�����ɾ�����һ��ͼ����ԭ������Ϊ3�����޸�Ϊ2
            If ucPreview.CurImageCount = 0 Then
                
                If mlngStudyState = 3 Then
                    strSql = "Zl_Ӱ����_State(" & mlngAdviceID & "," & mlngSendNo & ",2,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCurDeptId & ")"
                    zlDatabase.ExecuteProcedure strSql, "ɾ�����һ��ͼ��"
                End If
                
                Call SendMsgToMainWindow(Me, wetDelAllImg, mlngAdviceID, mstrStudyUID)
                
                mstrStudyUID = ""
                
                '������ͼ��ɾ��ʱ������ʾʵʱ��Ƶ����
                Call ConfigVideoShowState(True)
            Else
                Call SendMsgToMainWindow(Me, wetUpdateImg, mlngAdviceID, mstrStudyUID)
            End If
        End If
    End If
End Sub


Private Sub subSetMouseState(intMouseState As Integer)
    '�ı䵱ǰ���״̬
    If mintMouseState = intMouseState Then
        mintMouseState = 0
    Else
        mintMouseState = intMouseState
    End If
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Window).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Zoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_RectZoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Arrow).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Ellipse).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Text).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Corp).Checked = False
End Sub


Private Sub subSetSharp(blnSharp As Boolean)
'------------------------------------------------
'���ܣ�dcmView��ͼ���ƽ������
'������blnSharp��ʾͼ����ķ���True=�񻯣�False=ƽ��
'���أ��ޣ�ֱ�Ӵ���dcmView�е�ͼ��
'------------------------------------------------
    If mblnRealTime = False And dcmView.Images.Count > 0 Then
        If blnSharp = True Then
            '�񻯴���
            If dcmView.Images(1).FilterLength <= 0 Then
                dcmView.Images(1).FilterLength = 0
                '��ǰû��ƽ������ֱ�ӽ����񻯴���
                dcmView.Images(1).UnsharpEnhancement = dcmView.Images(1).UnsharpEnhancement + 0.1
            Else
                '�����ǰ�Ѿ���ƽ���������ȵ���ƽ��Ч��
                dcmView.Images(1).FilterLength = dcmView.Images(1).FilterLength - 1
            End If
        Else
            'ƽ������
            '�ж�Zoom�Ƿ�1������ǣ����޸�Ϊ0.9999
            If dcmView.Images(1).ActualZoom = 1 Then
                dcmView.Images(1).Zoom = 0.9999
            End If
            
            If dcmView.Images(1).UnsharpEnhancement <= 0 Then
                dcmView.Images(1).UnsharpEnhancement = 0
                '��ǰû���񻯴���ֱ�ӿ�ʼƽ��
                '�ж�FilterLength�Ƿ�0����ǣ�����2/ActualZoom��2��FilterLength֮����е���
                If dcmView.Images(1).FilterLength = 0 Then
                    dcmView.Images(1).FilterLength = 2 / dcmView.Images(1).ActualZoom + 1
                Else    '���������FilterLength��1
                    dcmView.Images(1).FilterLength = dcmView.Images(1).FilterLength + 1
                End If
            Else
                '��ǰ�Ѿ������񻯴����ȵ����񻯵�Ч��
                dcmView.Images(1).UnsharpEnhancement = dcmView.Images(1).UnsharpEnhancement - 0.1
            End If
        End If
    End If
End Sub


Private Sub subSetRotate(blnClockwise As Boolean)
'------------------------------------------------
'���ܣ�dcmView��ͼ�����ת
'������blnClockwise��ת�ķ���,True=˳ʱ����ת��False=��ʱ����ת
'���أ��ޣ�ֱ�Ӵ���dcmView�е�ͼ��
'------------------------------------------------
    If mblnRealTime = False And dcmView.Images.Count > 0 Then
        Dim iRotateState As Integer
        
        iRotateState = dcmView.Images(1).RotateState
        If blnClockwise = True Then
            iRotateState = iRotateState - 1
        Else
            iRotateState = iRotateState + 1
        End If
        If iRotateState = -1 Then iRotateState = 3
        iRotateState = iRotateState Mod 4
        dcmView.Images(1).RotateState = iRotateState
    End If
End Sub


'modify by tjh at 2010-01-20
'������Ƶ��ʾ״̬
Public Sub ConfigVideoShowState(ByVal blnShowState As Boolean)
  mblnRealTime = blnShowState
  
  Select Case mVideoDriverType
    Case vdtVFW
      picVideo.Visible = blnShowState
      wdmCapture.Visible = False
      dcmView.Visible = Not blnShowState
    Case vdtWDM
      wdmCapture.Visible = blnShowState
      picVideo.Visible = False
      dcmView.Visible = Not blnShowState
    Case vdtTWAIN
      wdmCapture.Visible = False
      picVideo.Visible = False
      mblnRealTime = False
      
      dcmView.Visible = True
  End Select
End Sub


'modify by tjh at 2010-01-20
'������Ƶ�ü����λ��
Private Sub RefreshPbxSizePos()
  '��
  pbxSize(0).Top = picView.Top - pbxSize(0).Height + 5
  pbxSize(0).Left = picView.Left
  pbxSize(0).Width = picView.Width
  
  '��
  pbxSize(1).Top = picView.Top + picView.Height - 5
  pbxSize(1).Left = picView.Left
  pbxSize(1).Width = picView.Width
  
  '��
  pbxSize(2).Top = picView.Top - pbxSize(0).Height
  pbxSize(2).Left = picView.Left - pbxSize(2).Width + 5
  pbxSize(2).Height = picView.Height + pbxSize(0).Height * 2
  
  '��
  pbxSize(3).Top = picView.Top - pbxSize(0).Height
  pbxSize(3).Left = picView.Left + picView.Width - 5
  pbxSize(3).Height = picView.Height + pbxSize(0).Height * 2
  
  'pbxsizeˢ����ʾ
  Call pbxSize(0).Refresh
  Call pbxSize(1).Refresh
  Call pbxSize(2).Refresh
  Call pbxSize(3).Refresh
End Sub


'modify by tjh at 2010-01-20
'�ı���Ƶ�ü���Χ
Private Sub ChangeCutRanage(videoObj As Object, ByVal lngChangeIndex As Long, ByVal X As Long, ByVal Y As Long)
  Dim lngDistanceX As Long
  Dim lngDistanceY As Long
  
  lngDistanceX = X ' - mStartPoint.X
  lngDistanceY = Y ' - mStartPoint.Y
  
  
  Select Case lngChangeIndex
    Case moUp '��--------------------------------------------------
      If (picView.Height - lngDistanceY) <= 50 * mdblZoomRate Then Exit Sub
      If videoObj.Top - lngDistanceY > 0 Then Exit Sub  'lngDistanceY = 0
     
      videoObj.Top = videoObj.Top - lngDistanceY
      
      picView.Top = picView.Top + lngDistanceY
      picView.Height = (picView.Height - lngDistanceY)
    Case moDown '��--------------------------------------------------
      If (picView.Height + lngDistanceY) <= 50 * mdblZoomRate Then Exit Sub
      'If Abs(0 - DSCapture.Top) + Picturexx.Height >= mVideoSize.Height * mdblVZoomRate Then Exit Sub
            
      picView.Height = (picView.Height + lngDistanceY)
      
      If Abs(0 - videoObj.Top) + picView.Height >= mVideoSize.Height * mdblZoomRate Then
        picView.Height = (picView.Height - lngDistanceY)
      End If
    Case moLeft '��--------------------------------------------------
      If (picView.Width - lngDistanceX) <= 50 * mdblZoomRate Then Exit Sub
      If videoObj.Left - lngDistanceX > 0 Then Exit Sub ' lngDistanceX = 0
      
      videoObj.Left = videoObj.Left - lngDistanceX
      
      picView.Left = picView.Left + lngDistanceX
      picView.Width = picView.Width - lngDistanceX
    
    Case moRight '��--------------------------------------------------
      If (picView.Width + lngDistanceX) <= 50 * mdblZoomRate Then Exit Sub
      'If Abs(0 - DSCapture.Left) + Picturexx.Width >= mVideoSize.Width * mdblHZoomRate Then Exit Sub
            
      picView.Width = picView.Width + lngDistanceX
      
      If Abs(0 - videoObj.Left) + picView.Width >= mVideoSize.Width * mdblZoomRate Then
        picView.Width = picView.Width - lngDistanceX
      End If
  End Select
End Sub


'modify by tjh at 2010-01-20
'Ӧ�òü���Χ����
Private Sub ApplayCutRange(videoObj As Object)

   mCurCutRange.LeftRate = Abs(videoObj.Left) / (mVideoSize.Width * mdblZoomRate)
   mCurCutRange.WidthRate = (mVideoSize.Width * mdblZoomRate - picView.Width + videoObj.Left) / (mVideoSize.Width * mdblZoomRate)

   mCurCutRange.TopRate = Abs(videoObj.Top) / (mVideoSize.Height * mdblZoomRate)
   mCurCutRange.HeightRate = (mVideoSize.Height * mdblZoomRate - picView.Height + videoObj.Top) / (mVideoSize.Height * mdblZoomRate)
End Sub


Private Sub imageScanner_PageDone(ByVal PageNumber As Long)
  If mintScanImageIndex = -1 Then
    Exit Sub
  End If

  '����ɨ���ļ�����
  mintScanImageIndex = mintScanImageIndex + 1
  
  Dim curScanFile As String
  curScanFile = CStr(mintScanImageIndex)
  
  'ȡ����Ч��ɨ���ļ�����
  While Len(curScanFile) < 4
    curScanFile = "0" + curScanFile
  Wend
  
  curScanFile = mstrTempDirOfScan & "\" & CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE & curScanFile & ".bmp"
  
  '����ɨ���ͼ��
  Call subCaptureImg(True, curScanFile)
  
  Call ShowScanImage(ucPreview.CurImageCount)
End Sub


Private Sub ShowScanImage(imgIndex As Integer)

    '����ѡ��ͼ��װ�ص�dcmView��
    dcmView.Images.Clear
    dcmView.Images.Add ucPreview.SelectImage
    
    '��ʾdcmView������picVideo
    dcmView.CurrentImage.BorderWidth = 0
    mblnRealTime = False
'    picVideo.Visible = False
'    dcmView.Visible = True
End Sub


Private Sub mclsDxDevice_OnDxKeyPress(ByVal lngButtonNum As Long)
On Error GoTo errHandle
    
    Select Case lngButtonNum
        Case 0  'ǰ̨�ɼ�
            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_MarkMap).Enabled And _
                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_MarkMap).Visible Then
                Call subCaptureImg(True)
            End If
        Case 1  '��̨�ɼ�
            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Capture).Enabled And _
                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Capture).Visible Then
                Call subCaptureImg(True, "", Nothing, True)
            Else
                Call mclsDxDevice_OnDxKeyPress(0)
            End If
        Case 2  '���±��
            
            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Tag).Enabled And _
                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Tag).Visible Then
                Call UpdateAfterCaptureInfo
            Else
                Call mclsDxDevice_OnDxKeyPress(0)
            End If
        Case Else
            Call mclsDxDevice_OnDxKeyPress(0)
    End Select

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

'modify by tjh at 2010-01-20
Private Sub pbxSize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim intVideoCapture As Integer

    intVideoCapture = Val(zlDatabase.GetPara("����ı�ɼ������С", glngSys, mlngModul, "1", , InStr(mstrPrivs, ";��������;") > 0))
  '��ʼִ�вü���Χ����
    If Button = 1 And intVideoCapture = 1 Then
      mblnMoveDown = True
    End If
  
End Sub


'modify by tjh at 2010-01-20
Private Sub pbxSize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  '������Ƶ�ü���Χ
  If mblnMoveDown = True And Button = 1 Then
    If wdmCapture.Visible Then
      Call ChangeCutRanage(wdmCapture, Index, X, Y)
    ElseIf picVideo.Visible Then
      Call ChangeCutRanage(picVideo, Index, X, Y)
    Else
      Call ChangeCutRanage(dcmView, Index, X, Y)
    End If
      
            
    Call RefreshPbxSizePos

  End If
    
End Sub


'modify by tjh at 2010-01-20
Private Sub pbxSize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  If mblnMoveDown = True And Button = 1 Then
          
    'Ӧ�òü�����
    If wdmCapture.Visible Then
      Call ApplayCutRange(wdmCapture)
    ElseIf picVideo.Visible Then
      Call ApplayCutRange(picVideo)
    End If
    
    If IsTwainCaptureWay Then
      ConfigTwainDisplay
    Else
      '������ʾ��Χ
      Call ConfigVideoDisplay(wdmCapture)
      Call ConfigVideoDisplay(picVideo)

      'ˢ����Ƶ��ʾ
      If Not (mVideoCapture Is Nothing) Then
        Call mVideoCapture.RefreshVideoWindow
      End If
    End If

    '���òü��߿�λ��
    Call RefreshPbxSizePos

  End If
    
  mblnMoveDown = False
    
End Sub


Private Sub picCapture_Resize()
On Error GoTo errHandle
    
    '����ͼ���С
    If picCapture.Height < 7000 Or picCapture.Width < 4000 Then
        cbrMain.Options.SetIconSize True, 16, 16
    Else
        cbrMain.Options.SetIconSize True, 32, 32
    End If
errHandle:
End Sub


Private Sub subVideoPlay()
'------------------------------------------------
'���ܣ�dcmView��¼��ͼ��Ĳ���
'��������
'���أ��ޣ�ֱ�Ӳ���dcmView�е�ͼ��
'------------------------------------------------
    If dcmView.Images.Count > 0 Then
        '����¼��������ش��ڣ��򲻽�������
        If dcmView.Images(1).Tag.Tag <> VIDEOTAG And dcmView.Images(1).Tag.Tag <> AUDIOTAG Then
            '���ǵ�Ӱ��ʽ���ܲ���,������ʾ
            Exit Sub
        End If
        
        On Error GoTo continue1
        
            If dcmView.Images(1).Tag.Tag = VIDEOTAG Then
                Call dcmView.Images(1).FileImport(IIf(Len(App.Path) > 3, App.Path & "\..\�����ļ�\aviDownload.bmp", App.Path & "..\�����ļ�\aviDownLoad.bmp"), "DIB/BMP")
        
                '������Ҫ���ŵ���Ƶ
                Call GetSingleImage(dcmView.Images(1).InstanceUID, dcmView.Images(1).SeriesUID, mblnMoved)
            
                Call dcmView.Images(1).FileImport(IIf(Len(App.Path) > 3, App.Path & "\..\�����ļ�\avi.bmp", App.Path & "..\�����ļ�\avi.bmp"), "DIB/BMP")
            Else
                Call dcmView.Images(1).FileImport(IIf(Len(App.Path) > 3, App.Path & "\..\�����ļ�\wavDownload.bmp", App.Path & "..\�����ļ�\wavDownLoad.bmp"), "DIB/BMP")
        
                '������Ҫ���ŵ���Ƶ
                Call GetSingleImage(dcmView.Images(1).InstanceUID, dcmView.Images(1).SeriesUID, mblnMoved)
            
                Call dcmView.Images(1).FileImport(IIf(Len(App.Path) > 3, App.Path & "\..\�����ļ�\wav.bmp", App.Path & "..\�����ļ�\wav.bmp"), "DIB/BMP")
            End If
            
continue1:
            '�򿪲��š���
            Call frmPlaying.Show
            
            'ˢ�²��Ŵ���
'            Call frmPlaying.Refresh
            While Not frmPlaying.IsActive
                Call Sleep(10)
                DoEvents
            Wend
                
            
            Call frmPlaying.OpenVideoFile(Replace(dcmView.Images(1).Tag.VideoFile, "/", "\"), Me)
    End If
End Sub


Private Sub subVideoSaveAs()
'------------------------------------------------
'���ܣ����dcmView�е�ͼ��,֧�ֵĸ�ʽΪAVI,DCM,BMP,JPE
'��������
'���أ���
'------------------------------------------------
    Dim strFileName As String
    Dim strFileType As String
    
    If mblnRealTime = False And dcmView.Images.Count > 0 Then
    
        If dcmView.Images(1).Tag.Tag = VIDEOTAG Then
            dlgOpen.Filter = "(*.AVI)|*.AVI|(*.MPEG)|*.MPEG|(*.*)|*.*"
            
            dlgOpen.ShowSave
            strFileName = dlgOpen.Filename
            
            If strFileName <> "" Then
                '������Ƶ�ļ���ָ��·��
                Call FileCopy(dcmView.Images(1).Tag.VideoFile, strFileName)
            End If
            
            Exit Sub
        End If
            
        If dcmView.Images(1).FrameCount > 1 Then
            dlgOpen.Filter = "¼���ļ�(*.AVI)|*.AVI|DICOM�ļ�(*.dcm)|*.dcm|ͼ���ļ� (*.BMP)|*.BMP|ͼ���ļ�(*.JPG)|*.JPG"
        Else
            dlgOpen.Filter = "DICOM�ļ�(*.dcm)|*.dcm|ͼ���ļ� (*.BMP)|*.BMP|ͼ���ļ�(*.JPG)|*.JPG"
        End If
        
        
        dlgOpen.ShowSave
        strFileName = dlgOpen.Filename
        
        If strFileName <> "" Then
            strFileType = UCase(Right(Trim(strFileName), 3))
            
            Select Case strFileType
                Case "AVI"
                    If dcmView.Images(1).FrameCount > 1 Then
                        dcmView.Images(1).WriteAVI strFileName, 1, dcmView.Images(1).FrameCount, 1, "", 100, False
                    Else
                        MsgBoxD Me, "��̬ͼ���޷������AVI��ʽ��������ѡ��ͼ���ʽ��", vbInformation, gstrSysName
                    End If
                Case "DCM"
                    dcmView.Images(1).WriteFile strFileName, True
                Case "BMP"
                    dcmView.Images(1).FileExport strFileName, "BMP"
                Case "JPG"
                    dcmView.Images(1).FileExport strFileName, "JPG"
            End Select
        End If
    End If
End Sub


Private Sub InputImageFile()
'------------------------------------------------
'���ܣ����ⲿ�ļ�����������ͼ��
'��������
'���أ���
'------------------------------------------------
    Dim DlgInfo As DlgFileInfo
    Dim i As Integer
    Dim ImgTmpImage As New DicomImage
    Dim ImgTmpImages As New DicomImages
    Dim blDicomFile As Boolean              '�Ƿ�DICO�ļ� =TrueΪDICOM�ļ�
    Dim j As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    'ѡ���ļ�
    With Me.dlgOpen
        .CancelError = False
        .MaxFileSize = 32767 '���򿪵��ļ����ߴ�����Ϊ��󣬼�32K
        .flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
        .DialogTitle = "ѡ���ļ�"
        .Filter = "DICOM�ļ���*.dcm��(*.img)|*.dcm;*.img|ͼ���ļ� (*.BMP)(*.JPG)|*.BMP;*.JPG|�����ļ���*.*��|*.*"
        .ShowOpen
        If .Filename <> "" Then
            DlgInfo = GetDlgSelectFileInfo(.Filename)
        End If
        '�ڴ���*.pif�ļ����뽫Filename�����ÿգ�����ѡȡ���*.pif�ļ��󣬵�ǰ·����ı�
        .Filename = ""
    End With
    
    On Error Resume Next
    
    For i = 1 To DlgInfo.iCount
        err.Clear
        Set ImgTmpImage = Nothing
        ImgTmpImages.Clear
        ImgTmpImage.FileImport DlgInfo.sPath & DlgInfo.sFIle(i), ""
        If err <> 0 Then
            err.Clear
            ImgTmpImages.ReadFile DlgInfo.sPath & DlgInfo.sFIle(i)
            If err = 0 Then
                blDicomFile = True
            End If
        End If
        
        If blDicomFile = True And ImgTmpImages.Count > 0 Then
            Set ImgTmpImage = ImgTmpImages(1)
        End If
        
        '����ͼ���DICOM����
        subWriteDicomPara ImgTmpImage, mlngAdviceID
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.Tag = IMGTAG
    
        Set ImgTmpImage.Tag = dcmTag
    
        '��ͼ����뵽����ͼ��
        subInsert2Mini ImgTmpImage
        '����ͼ�񣬲�����ͼ��洢�¼�
        Call subSaveImage
        
        '����Ӱ����״̬������ɼ���һ��ͼ����ԭ����״̬���ѱ��������޸ĳ��Ѽ��
        If ucPreview.CurImageCount = 1 Then
            If mlngStudyState < 3 Then
                strSql = "Zl_Ӱ����_State(" & mlngAdviceID & "," & mlngSendNo & ",3,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCurDeptId & ")"
                zlDatabase.ExecuteProcedure strSql, "�ɼ���һ��ͼ��"
            End If
        End If
        
        If ucPreview.CurImageCount = 1 Then
            Call SendMsgToMainWindow(Me, wetCaptureFirstImg, mlngAdviceID)
        End If
    Next
End Sub


Private Sub subUniteUID(dcmImg As DicomImage, ByVal blnIsUpdateSeriesUid As Boolean)
'------------------------------------------------
'���ܣ���������ͼ��ļ��UID������UID����֤����ͼ��ļ��UID������UID������ͼdcmMiniature�е�һ�¡�
'       ����ӽ�����ͼ����õ�һ��ͼ��ļ��UID������UID��
'       ����ǵ�һ��ͼ������������ļ��UID������ͼ�����Դ��ļ��UID��ͬʱ�����UID������ֵ
'������dcmImg���������DICOMͼ��
'���أ��ޣ�ֱ���޸�ͼ��ļ��UID������UID
'------------------------------------------------
    Dim i As Integer
    
    '��ͼ����ø���һ��ͼ����ͬ�ļ��UID������UID
    If ucPreview.CurImageCount > 0 Then
                
        dcmImg.StudyUID = ucPreview.ImgViewer.Images(1).StudyUID
        
        '�������Ϊtrue�����������img������UID������ʹ���µ�����
        If blnIsUpdateSeriesUid Then
            '����Ϊͼ�������UID
            For i = 1 To ucPreview.ImgViewer.Images.Count
                If ucPreview.ImgViewer.Images(i).Tag.Tag = IMGTAG Then
                    dcmImg.SeriesUID = ucPreview.ImgViewer.Images(i).SeriesUID
                    
                    Exit For
                End If
            Next i
            
        End If
    ElseIf Len(mstrStudyUID) > 0 Then
        dcmImg.StudyUID = mstrStudyUID
    Else
        mstrStudyUID = dcmImg.StudyUID
        
        '�����uid�ı����Ҫ��������ͼ��ʾ����еĲ�ѯֵ
        ucPreview.QueryValue = mstrStudyUID
    End If
End Sub


Private Function GetDlgSelectFileInfo(strFileName As String) As DlgFileInfo
'------------------------------------------------
'���ܣ����ļ���ת��Ϊȫ·������
'������strFileName--�ļ�����ͨ�����ļ��ؼ�����á�
'���أ�ȫ·������
'------------------------------------------------
    Dim sPath, tmpStr As String
    Dim sFIle() As String
    Dim iCount, i As Integer
    On Error GoTo errHandle
    sPath = CurDir()  '��õ�ǰ��·������Ϊ��CommonDialog�иı�·��ʱ��ı䵱ǰ��Path
    tmpStr = Right$(strFileName, Len(strFileName) - Len(sPath)) '���ļ����������
    
    If Left$(tmpStr, 1) = Chr$(0) Then
        'ѡ���˶���ļ�(����Ϊ��һ���ַ�Ϊ�ո�)
        For i = 1 To Len(tmpStr)
            If Mid$(tmpStr, i, 1) = Chr$(0) Then
                iCount = iCount + 1
                ReDim Preserve sFIle(iCount)
            Else
                sFIle(iCount) = sFIle(iCount) & Mid$(tmpStr, i, 1)
            End If
        Next i
    Else
        'ֻѡ����һ���ļ�(ע�⣺��Ŀ¼�µ��ļ�����ȥ·����û��"\"��
        iCount = 1
        ReDim Preserve sFIle(iCount)
        If Left$(tmpStr, 1) = "\" Then tmpStr = Right$(tmpStr, Len(tmpStr) - 1)
        sFIle(iCount) = tmpStr
    End If
    
    GetDlgSelectFileInfo.iCount = iCount
    ReDim GetDlgSelectFileInfo.sFIle(iCount)
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    GetDlgSelectFileInfo.sPath = sPath
    For i = 1 To iCount
        GetDlgSelectFileInfo.sFIle(i) = sFIle(i)
    Next i
    Exit Function
errHandle:
    MsgBox "GetDlgSelectFileInfo����ִ�д���", vbOKOnly + vbCritical, gstrSysName
End Function



Private Sub TimerHook_Timer()
On Error GoTo errHandle
    '��ʹ��hook�ȼ����òɼ�ʱ��ʹ��timer���вɼ�������������ִ�ж��CaptureImage������hookʧЧ
    '���hookʧЧ�Ŀ���ԭ����hook�Ĵ������������ػ�hook��Ĵ���ʱ������������ʧЧ������dicomobjects��fileexport�������ö�����ʧЧ��Ŀǰ��ȥϸ��
    Call CaptureImage
    timerHook.Enabled = False
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub tmrComm_Timer()
    On Error GoTo errHandle
    If mstrActiveType = "COM" Then
        mlngComTime = mlngComTime + 2
        
        '����0.08�룬���Զ�����
        If mlngComTime > 40 Then
            mlngComTime = 0
            tmrComm.Enabled = False
        End If
        
    Else
         If Not mclsDxDevice Is Nothing Then Call mclsDxDevice.PollDxDevice
    End If
    
    Exit Sub
errHandle:
    tmrComm.Enabled = False
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub tmrReg_Timer()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo errHandle:
    If Not mVideoCapture.IsStartup Then
        Exit Sub
    End If
    
    If gint��Ƶ�豸���� <= -1 Then Exit Sub
    
    strSql = "select count(1) ���������� from zltools.zlclients where ������ƵԴ=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "����������")
    
    If rsTemp.RecordCount > gint��Ƶ�豸���� Then
        mstrVideoRegTime = ""
        
        Exit Sub
    End If
    
    If DateDiff("S", mstrVideoRegTime, Now) >= M_LNG_REFRESHINTERVAL Then
        '�ж����ݿ����Ƿ�����Ѿ�ע���ip�����Ѿ�������ƵԴ���������������Ϊû�гɹ�ע��
        If FunCheckRegInfo(Me) Then
            mstrVideoRegTime = Now
        Else
            mstrVideoRegTime = ""
            
            Exit Sub
        End If
        
        
    End If
    
Exit Sub
errHandle:
End Sub

Private Sub txtInputText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then  '''ESC�ͻس����˳�����
        txtInputText.Visible = False
        If Trim(txtInputText.Text) = "" Or KeyAscii = 27 Then
            'ɾ�����ֱ�ע
            dcmView.Images(1).Labels.Remove dcmView.Images(1).Labels.Count
            txtInputText = "1 "
        Else
            mdcmSelectLabel.Text = txtInputText.Text
            dcmView.Refresh
        End If
    End If
End Sub


Private Sub subVideoSave()
'------------------------------------------------
'���ܣ�¼��
'��������
'���أ���¼���ļ���������ͼ
'------------------------------------------------
    
    Dim i As Integer
    Dim dcmTmpImg As New DicomImage
    Dim strError As String
            
    On Error GoTo continue1
      'ɾ����ʷ����Ƶ�ļ�
      If Dir(mstrAviFileName) <> "" Then
        Kill mstrAviFileName
      End If
continue1:
    
    On Error GoTo CapErr
            
    '����Ŀǰ�ķ�ʽ,ʹ��vfw��ʱ���������¼�����
    If mVideoCapture.VideoDriverType = vdtVFW Then
        '¼�����(vfw����¼���ֱ��������ִ��StartVideo�Ժ�����)
        '������vfw��¼����
        Exit Sub
    End If
    
    'modify by tjh at 2010-01-20
    strError = mVideoCapture.StartVideo(mstrAviFileName)
    If Trim(strError) <> "" Then MsgBoxD Me, strError, vbInformation, gstrSysName
    
    '��ȡ��ǰ¼��ı���������
    mstrEncoderName = mVideoCapture.GetEncoderName
    
    Exit Sub
CapErr:
  Call MsgBox(err.Description, vbOKOnly, "��ʾ")
End Sub


'modify by tjh at 2010-01-20
'ֹͣ��Ƶ¼��
Private Sub subStopVideo()
    Dim dcmTmpImg As New DicomImage
            
    If mVideoCapture.VideoDriverType = vdtVFW Then Exit Sub
    
    On Error GoTo continue1
    If Dir(mstrAviFileName) <> "" Then
        Kill mstrAviFileName
    End If
continue1:
    
    On Error GoTo CapErr
    
    Call mVideoCapture.StopVideo
   
    
    '¼�����
    If Dir(mstrAviFileName) <> "" Then
        On Error GoTo continue2
            dcmTmpImg.FileImport IIf(Len(App.Path) > 3, App.Path & "\..\�����ļ�\avi.bmp", App.Path & "..\�����ļ�\avi.bmp"), "DIB/BMP"
continue2:
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.EncoderName = mstrEncoderName
        dcmTag.VideoFile = mstrAviFileName
        dcmTag.CaptureTime = zlDatabase.Currentdate
        dcmTag.RecordTimeLen = mVideoCapture.GetTimeLen
        dcmTag.Tag = VIDEOTAG
        
        Set dcmTmpImg.Tag = dcmTag
        
        subWriteDicomPara dcmTmpImg, mlngAdviceID
        
        subInsert2Mini dcmTmpImg
        
        '������Ƶ¼��
        Call subSaveImage
    End If
    
    '�����¼��Ҳ��Ҫ��״̬���и���
    If ucPreview.CurImageCount = 1 Then
        Call SendMsgToMainWindow(Me, wetCaptureFirstImg, mlngAdviceID)
    End If
    
    Exit Sub
CapErr:
    If ErrCenter() = 1 Then Resume
End Sub


'ֹͣ��Ƶ�ļ�
Public Sub subSaveAudio(ByVal strAudioFile As String, ByVal lngTimeLen As Long)

    Dim i As Integer
    Dim dcmTmpImg As New DicomImage
    
    On Error GoTo CapErr
   
    
    '¼�����
    If Dir(strAudioFile) <> "" Then
        dcmTmpImg.FileImport IIf(Len(App.Path) > 3, App.Path & "\..\�����ļ�\wav.bmp", App.Path & "..\�����ļ�\wav.bmp"), "DIB/BMP"
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.EncoderName = ""
        dcmTag.VideoFile = strAudioFile
        dcmTag.CaptureTime = zlDatabase.Currentdate
        dcmTag.RecordTimeLen = lngTimeLen
        dcmTag.Tag = AUDIOTAG
        
        Set dcmTmpImg.Tag = dcmTag
        
        subWriteDicomPara dcmTmpImg, mlngAdviceID
        
        subInsert2Mini dcmTmpImg
        
        '����¼�Ƶ���Ƶ
        Call subSaveImage
    End If
    
    '�����¼��Ҳ��Ҫ��״̬���и���
    If ucPreview.CurImageCount = 1 Then
        Call SendMsgToMainWindow(Me, wetCaptureFirstImg, mlngAdviceID)
    End If
    
    Exit Sub
CapErr:
    If ErrCenter() = 1 Then Resume
End Sub

'modify by tjh at 2010-01-22
'ȫ����ʾ
Private Sub subFullCall()
  Call mVideoCapture.FullScreen(Me, Me.hWnd)
End Sub


Private Function GetCaptureTag() As String
'ȡ�ú�̨�ɼ����
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
        
    GetCaptureTag = "001"
        
    strSql = "select ���� from Ӱ����ʱ��¼ where ����='��̨'"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    For i = 1 To 999
        rsData.Filter = " ����='" & Lpad(i, 3, "0") & "'"
        If rsData.RecordCount <= 0 Then
            GetCaptureTag = Lpad(i, 3, "0")
            Exit Function
        End If
    Next i
    
    GetCaptureTag = ""
End Function



Private Sub CreateNewCaptureTag()
'ȡ���µĲɼ����
    mstrAfterStudyUid = CreateStudyUid(dcmglbUID.NewUID)
    mstrAfterSeriesUid = CreateSeriesUid(dcmglbUID.NewUID)
    
    mstrAfterTag = GetCaptureTag
    
    mlngAfterCurImageCount = 0
End Sub


Private Sub ShowAfterCaptureInf()
'���º�̨�ɼ�ͼ����Ϣ
    If Not mblnAfterIsUse Then Exit Sub
    
    If mobjOwner Is Nothing Then Exit Sub
    
    If mstrAfterParentTitle = "" Then
        If InStr(mobjOwner.Caption, "      ��̨�ɼ���ǣ�") > 0 Then
            mstrAfterParentTitle = Mid(mobjOwner.Caption, 1, InStr(mobjOwner.Caption, "      ��̨�ɼ���ǣ�") - 1)
        Else
            mstrAfterParentTitle = mobjOwner.Caption
        End If
    End If
    
    mobjOwner.Caption = mstrAfterParentTitle & "      ��̨�ɼ���ǣ�" & mstrAfterTag & "  ��ǰ��̨�ɼ�����" & mlngAfterCurImageCount & "        "
End Sub


Private Function subSaveAfterCaptureImage(Optional iEncode As Integer = 0) As Boolean
'�����̨�ɼ�ͼ��
    Dim i As Long
    Dim lngResult As Long
    Dim strSql As String
    Dim dtNowTime As Date
    Dim strReceivedTime As String
    Dim ImgTmp As DicomImage

    subSaveAfterCaptureImage = False
    
    If dcmAfter.Images.Count <= 0 Then Exit Function
    
    dtNowTime = zlDatabase.Currentdate
    strReceivedTime = Format(dtNowTime, "yyyyMMdd")
    
    If mstrAfterStudyUid = "" Then
        '���uidΪ�գ��򴴽��µ�UID
        mstrAfterStudyUid = dcmglbUID.NewUID
        mstrAfterSeriesUid = dcmglbUID.NewUID
        
        mstrAfterTag = GetCaptureTag()
    End If
    
    If Trim(mstrAfterTag) = "" Then
        Call MsgBoxD(Me, "���ܻ�ȡ��Ч�ĺ�̨�ɼ���ǣ������̨�ɼ��ļ�������Ƿ���������̨�ɼ���������ܳ���1000��", vbOKOnly, Me.Caption)
        Exit Function
    End If

    '��������Ŀ¼
    MkLocalDir mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/"
 
    '��ʹ��inet��ʽʱ����Ҫ�ȳ�ʼ��ftp����
    lngResult = mobjFtpConnection.FuncFtpConnect(mobjFtp.strFTPIP, mobjFtp.strFTPUser, mobjFtp.strFTPPwd)

    If lngResult = 0 Then
        'FTP����ʧ�ܣ���ʾ���󣬲�ɾ������ͼ�е�ͼ��
        MsgBoxD Me, "FTP����ʧ�ܣ���̨�ɼ�ͼ���޷����棬�����������á�", vbInformation, gstrSysName
        Exit Function
    End If
        
    For i = 1 To dcmAfter.Images.Count
    
        Set ImgTmp = dcmAfter.Images(i)
        
        ImgTmp.StudyUID = mstrAfterStudyUid
        ImgTmp.SeriesUID = mstrAfterSeriesUid
        
        If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
            '����ͼ�񵽻���Ŀ¼
            Select Case iEncode
                Case 1          'Run-Length Encoding�г�ѹ��
                    ImgTmp.WriteFile mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.5"
                Case 2          '����������ԭͼ��ѹ����ʽ
                    ImgTmp.WriteFile mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID, True
                Case Else       'Lossless JPEG encoding JPEG����ѹ��
                    ImgTmp.WriteFile mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.4.70"
            End Select
            
            '�洢Ϊ����ͼ��
            ImgTmp.FileExport mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID & ".jpg", "JPG", 80
        End If
        
        If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
            '����dicomͼ��
            WriteToURL mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID, mobjFtp.strFtpDir & _
                strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID
                
            '�ϴ�����ͼ
            WriteToURL mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID & ".jpg", mobjFtp.strFtpDir & _
                strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID & ".jpg"
        Else
            '����¼��
            WriteToURL ImgTmp.Tag.VideoFile, mobjFtp.strFtpDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID
            
            If ImgTmp.Tag.Tag = VIDEOTAG Then
                '��¼���Ƶ���Ӧ��Ŀ¼�У�������������
                Call MoveFile(ImgTmp.Tag.VideoFile, mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID & ".avi")
                
            ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
                '����Ƶ�ļ����Ƶ���Ӧ��Ŀ¼�У�������������
                Call MoveFile(ImgTmp.Tag.VideoFile, mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID & ".wav")
                
            End If
        End If
        
        'ͼ��洢�ɹ��󣬴洢���ݿ���Ϣ
        strSql = "ZL_Ӱ����_��̨�ɼ�('" & mstrAfterModality & "','" & mstrAfterStudyUid & "','" & mstrAfterSeriesUid & "','" & _
                                        ImgTmp.InstanceUID & "','" & mstrAfterTag & "','" & mobjFtp.strDeviceId & "'," & _
                                        "to_Date('" & Format(dtNowTime, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'))"
        
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        mlngAfterCurImageCount = mlngAfterCurImageCount + 1
    Next i
    
    If mblnUseInetFtp Then
        'ʹ��inet ftp��ʽʱ�����ﲻ��Ҫ�Ͽ�����
    Else
        mobjFtpConnection.FuncFtpDisConnect
    End If
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        Call frmCaptureHint.ShowCaptureHint( _
            IIf(mblnPoputWindowHint, mstrBufferDir & strReceivedTime & "/" & mstrAfterStudyUid & "/" & ImgTmp.InstanceUID, ""), _
            mblnSoundHint, hpRB, Me)
            
    End If
    
    subSaveAfterCaptureImage = True
End Function


Private Sub subSaveImage(Optional iEncode As Integer = 0)
'------------------------------------------------
'���ܣ������һ������ͼ���浽���ݿ���
'������iEncode����ѹ����ʽ��1��Run-Length Encoding�г�ѹ����2������������ԭͼ��ѹ����ʽ��������Lossless JPEG encoding JPEG����ѹ��
'���أ���
'------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim ImgTmp As New DicomImage
    
    Dim dtReceived As String
    Dim blnFirstImage As String     '�Ƿ񱾴μ��ĵ�һ��ͼ��
    Dim lngResult As String         'FTP�������
    Dim nowTime As Date
    Dim strReportImages As String
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean       '�����ﴦ�������
    Dim i As Integer
    Dim lngSendNO As Long
    
    '��ȡ���һ������ͼ
    With ucPreview.ImgViewer
        If .Images.Count <= 0 Then Exit Sub
        Set ImgTmp = .Images(.Images.Count)
    End With
    
    '�ȱ���FTPͼ��
    '��ȡ��������
    gstrSQL = "select ���UID ,��������,����ͼ��,���ͺ� from Ӱ�����¼ where ҽ��ID = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, App.ProductName, mlngAdviceID)
    nowTime = zlDatabase.Currentdate
    
    If IsNull(rsTmp("���UID")) Then
        dtReceived = Format(nowTime, "yyyyMMdd")
        blnFirstImage = True
    Else
        dtReceived = Format(rsTmp("��������"), "yyyyMMdd")
        blnFirstImage = False
    End If
    
    '��������Ŀ¼
    MkLocalDir mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/"
    lngSendNO = rsTmp!���ͺ�
    
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        strReportImages = Nvl(rsTmp("����ͼ��"))
    
    
        '��鱨��ͼ��ĳ��ȣ��������4000���ֽڣ�����ʾ�޷�����ͼ��
        If Len(strReportImages & " ;" & ImgTmp.InstanceUID & ".jpg") >= 4000 Then
            MsgBoxD Me, "ͼ�������������ޣ�����ɾ������ͼ����ټ����ɼ�ͼ��", vbInformation, gstrSysName
            Call ucPreview.DeleteImage(ucPreview.CurImageCount)
            Exit Sub
        End If
    
        '����ͼ�񵽻���Ŀ¼
        Select Case iEncode
            Case 1          'Run-Length Encoding�г�ѹ��
                ImgTmp.WriteFile mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.5"
            Case 2          '����������ԭͼ��ѹ����ʽ
                ImgTmp.WriteFile mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID, True
            Case Else       'Lossless JPEG encoding JPEG����ѹ��
                ImgTmp.WriteFile mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.4.70"
        End Select

        ImgTmp.FileExport mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".jpg", "JPG", 80
    End If
    
    '��ʹ��inet��ʽʱ����Ҫ�ȳ�ʼ��ftp����
    lngResult = mobjFtpConnection.FuncFtpConnect(mobjFtp.strFTPIP, mobjFtp.strFTPUser, mobjFtp.strFTPPwd)
    If lngResult = 0 Then
        'FTP����ʧ�ܣ���ʾ���󣬲�ɾ������ͼ�е�ͼ��
        MsgBoxD Me, "FTP����ʧ�ܣ�ͼ���޷����棬�����������á�", vbInformation, gstrSysName
        Call ucPreview.DeleteImage(ucPreview.CurImageCount)
    
        Exit Sub
    End If

    If Val(mobjBakFtp.strDeviceId) > 0 Then
        lngResult = mobjBakFtpConnection.FuncFtpConnect(mobjBakFtp.strFTPIP, mobjBakFtp.strFTPUser, mobjBakFtp.strFTPPwd)
        If lngResult = 0 Then
            mobjBakFtp.strDeviceId = ""
            MsgBoxD Me, "����ftp�豸����ʧ�ܣ��ɼ���ͼ�񽫲��ܽ��б��ݲ��������豸���������̹����еı����豸���á�", vbInformation, gstrSysName
        End If
    End If


    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        '����dicomͼ��
        WriteToURL mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID, mobjFtp.strFtpDir & _
            dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID
            
        WriteToURL mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".jpg", mobjFtp.strFtpDir & _
            dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".jpg"
            
        '���ݵ�ǰ�ɼ���ͼ��
        If mobjBakFtpConnection.hConnection <> 0 Then
            BakImgToURL mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID, mobjBakFtp.strFtpDir & _
                dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID
        End If
    Else
        '����¼��
        WriteToURL ImgTmp.Tag.VideoFile, mobjFtp.strFtpDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID

        '����¼��
        If mobjBakFtpConnection.hConnection <> 0 Then
            BakImgToURL ImgTmp.Tag.VideoFile, mobjBakFtp.strFtpDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID
        End If
        
        
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            '��¼���Ƶ���Ӧ��Ŀ¼�У�������������
            Call FileCopy(ImgTmp.Tag.VideoFile, mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".avi")
            Call Kill(ImgTmp.Tag.VideoFile)
        
            ImgTmp.Tag.VideoFile = mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".avi"
        ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
            '����Ƶ�ļ����Ƶ���Ӧ��Ŀ¼�У�������������
            Call FileCopy(ImgTmp.Tag.VideoFile, mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".wav")
            Call Kill(ImgTmp.Tag.VideoFile)
        
            ImgTmp.Tag.VideoFile = mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID & ".wav"
        End If
    End If
    
    If mblnUseInetFtp Then
        'ʹ��inetftpʱ������Ҫÿ�ζ��Ͽ�����
    Else
        mobjFtpConnection.FuncFtpDisConnect
        
        If mobjBakFtpConnection.hConnection <> 0 Then mobjBakFtpConnection.FuncFtpDisConnect
    End If
    

    'ͼ��洢�ɹ��󣬴洢���ݿ���Ϣ
    On Error GoTo DBError
    arrSQL = Array()
    
    If blnFirstImage Then
        gstrSQL = "ZL_Ӱ�����¼_SET(" & mlngAdviceID & "," & lngSendNO & ",'" & _
            mstrStudyUID & "',null," & _
            "to_Date('" & Format(nowTime, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'),'" & mobjFtp.strDeviceId & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    
    gstrSQL = "Select ����UID From Ӱ��������  Where ����UID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", CStr(ImgTmp.SeriesUID))
    
    '�����µļ������,���Ϊ¼��������µ�����
    If rsTmp.EOF Or ImgTmp.Tag.Tag = VIDEOTAG Or ImgTmp.Tag.Tag = AUDIOTAG Then
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            gstrSQL = "ZL_Ӱ������_INSERT('" & mstrStudyUID & "','" & ImgTmp.SeriesUID & "','��Ƶ¼��',0)"
        ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
            gstrSQL = "ZL_Ӱ������_INSERT('" & mstrStudyUID & "','" & ImgTmp.SeriesUID & "','��Ƶ����',0)"
        Else
            gstrSQL = "ZL_Ӱ������_INSERT('" & mstrStudyUID & "','" & ImgTmp.SeriesUID & "','" & ImgTmp.SeriesDescription & "',0)"
        End If
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        '�����µ�ͼ���¼
        gstrSQL = "ZL_Ӱ��ͼ��_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "',NULL,0, null, sysdate)"
    Else
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            '�����µ���Ƶ��¼
            gstrSQL = "ZL_Ӱ��ͼ��_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "',Null,0" & _
            ",null,sysdate,null,null,null,null,null,null,null,null,null," & VIDEOTAG & ",'" & mstrEncoderName & "'," & ImgTmp.Tag.RecordTimeLen & ")"
        Else
            '�����µ���Ƶ��¼
            gstrSQL = "ZL_Ӱ��ͼ��_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "',Null,0" & _
            ",null,sysdate,null,null,null,null,null,null,null,null,null," & AUDIOTAG & ",''," & ImgTmp.Tag.RecordTimeLen & ")"
        End If
    End If
        
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        '������Ǽ��ͼ���򲻱��汨��ͼ
        gstrSQL = "ZL_Ӱ���鱨��_ADD('" & mstrStudyUID & "','" & ImgTmp.InstanceUID & ".jpg')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    End If
    
    gcnOracle.BeginTrans        '----------����ͼ��
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "����ͼ��")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        Call frmCaptureHint.ShowCaptureHint( _
            IIf(mblnPoputWindowHint, mstrBufferDir & dtReceived & "/" & mstrStudyUID & "/" & ImgTmp.InstanceUID, ""), _
            mblnSoundHint, hpRB, Me)
    End If
    
    Exit Sub
DBError:
    '������������ݿ����������ɾ�����ɼ���ͼ��
    If blnInTrans = True Then gcnOracle.RollbackTrans
    err.Raise err.Number, "���ͼ�񱣴�"
    Call ucPreview.DeleteImage(ucPreview.CurImageCount)
End Sub



Private Sub WriteToURL(ByVal SrcFileName As String, ByVal DestFileName As String)
'------------------------------------------------
'���ܣ��������ļ����浽Զ��������
'������SrcFileName--�����ļ�����DestFileName����Զ���ļ���
'���أ���
'-----------------------------------------------
'���ܣ�
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strPath As String

    '��FTP�д���Ŀ¼
    strPath = objFileSystem.GetParentFolderName(DestFileName)
    mobjFtpConnection.FuncFtpMkDir "/", strPath
    
    '��FTP�ϴ��ļ�
    mobjFtpConnection.FuncUploadFile strPath, SrcFileName, objFileSystem.GetFileName(DestFileName)
End Sub


Private Sub BakImgToURL(ByVal SrcFileName As String, ByVal DestFileName As String)
'------------------------------------------------
'���ܣ�����ͼ��Զ��������
'������SrcFileName--�����ļ�����DestFileName����Զ���ļ���
'���أ���
'-----------------------------------------------
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strPath As String
    
    If mobjBakFtpConnection.hConnection = 0 Then Exit Sub
    
    '��FTP�д���Ŀ¼
    strPath = objFileSystem.GetParentFolderName(DestFileName)
    mobjBakFtpConnection.FuncFtpMkDir "/", strPath
    
    '��FTP�ϴ��ļ�
    mobjBakFtpConnection.FuncUploadFile strPath, SrcFileName, objFileSystem.GetFileName(DestFileName)
End Sub


Private Sub RemoveFromURL(ByVal DestFileName As String)
'------------------------------------------------
'���ܣ���FTPɾ���ļ�
'������DestFileName����Զ���ļ���
'���أ���
'-----------------------------------------------
    Dim objFileSystem As New Scripting.FileSystemObject
    
    mobjFtpConnection.FuncDelFile objFileSystem.GetParentFolderName(DestFileName), objFileSystem.GetFileName(DestFileName)
End Sub


Private Sub InitCommandBars()
'���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim strRegPath As String
    
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
    
    strRegPath = "����ģ��\" & App.ProductName & "\frmVideoCapture"
    
    '�Ƿ���ʾ��������
    mblnShowProcessBar = GetSetting("ZLSOFT", strRegPath, "��ʾ��������", "True")
    
    '�ɼ�����������
    Set cbrToolBar = Me.cbrMain.Add("�ɼ���", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        '�ڷ�TWAIN�ɼ�ģʽ������£�����ʾ�ð�ť
        'If Not GetIsTwainCaptureWay Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Dynamic, "��̬"): cbrControl.ToolTipText = "��ʾʵʱ��Ƶ"
        'End If
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_MarkMap, "�ɼ�"): cbrControl.ToolTipText = "�ɼ�ͼ��"
        
        If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Capture, "��̨�ɼ�")
                cbrControl.ToolTipText = "��̨�ɼ�"
                cbrControl.IconId = 10020
        End If
        
        '�ڷ�TWAIN�ɼ�ģʽ������£�����ʾ�ð�ť
        'If Not GetIsTwainCaptureWay Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Record, "¼��"): cbrControl.ToolTipText = "��ʼ¼��"
                cbrControl.Enabled = True
                
            If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Record, "��̨¼��")
                    cbrControl.ToolTipText = "��̨¼��"
                    cbrControl.IconId = 10021
            End If
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Record_Stop, "ֹͣ¼��"): cbrControl.ToolTipText = "ֹͣ¼��"
                cbrControl.Enabled = False
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_RecordAudio, "¼��"): cbrControl.ToolTipText = "¼��"
        'End If
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Play, "����"): cbrControl.ToolTipText = "����¼��"
            cbrControl.BeginGroup = True
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Import, "����"): cbrControl.ToolTipText = "�ļ�����"
            cbrControl.IconId = 10002
            cbrControl.BeginGroup = True
            
            
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_SaveAs, "���"): cbrControl.ToolTipText = "�ļ����": cbrControl.IconId = 3091
            cbrControl.IconId = 10004
            
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DelImg, "ɾͼ"): cbrControl.ToolTipText = "ɾ��ͼ��": cbrControl.IconId = 10001
            
            
        If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_OpenStudyList, "�򿪼��"): cbrControl.ToolTipText = "�򿪼��"
            cbrControl.BeginGroup = True
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_StudySyncState, "�������"): cbrControl.ToolTipText = "�������"
            cbrControl.IconId = 10012
            
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Tag, "��Ǽ��")
                cbrControl.ToolTipText = "��Ǽ��"
                cbrControl.IconId = 10022
        End If
        
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIcon '  xtpButtonIconAndCaption
        cbrControl.Category = "�ɼ�"
        cbrControl.Enabled = False
    Next
    
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarRight)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Window, "����"): cbrControl.ToolTipText = "��������/�Աȶ�"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Zoom, "����"): cbrControl.ToolTipText = "����ͼ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Corp, "�϶�"): cbrControl.ToolTipText = "�϶�ͼ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RectZoom, "�ü��ɼ�"): cbrControl.ToolTipText = "�ü��ɼ�ͼ��": cbrControl.IconId = 3201
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RRotate, "˳ʱ"): cbrControl.ToolTipText = "˳ʱ����ת"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_LRotate, "��ʱ"): cbrControl.ToolTipText = "��ʱ����ת"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Sharpness, "��"): cbrControl.ToolTipText = "��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Filter, "ƽ��"): cbrControl.ToolTipText = "ƽ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Arrow, "��ͷ"): cbrControl.ToolTipText = "��ͷ��ע"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Ellipse, "Բ��"): cbrControl.ToolTipText = "Բ�α�ע"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Text, "����"): cbrControl.ToolTipText = "���ֱ�ע"
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "�߼�"): cbrControl.ToolTipText = "�߼�����"
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIcon
        cbrControl.Category = "����"
        cbrControl.Enabled = False
    Next
    cbrToolBar.Visible = mblnShowProcessBar
End Sub


Public Sub ShowFrameSelectImagePopup()
'------------------------------------------------
'���ܣ�������ѡͼ���ʱ�� ������Ҽ��ĵ����˵�
'������
'���أ���
'------------------------------------------------

Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    
    '����Ҽ������˵�
    Set cbrToolBar = Me.cbrMain.Add("����Ҽ�", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RectCapture, "�ü��ɼ�")
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub


'DicomViewer�ü���ɼ�ͼ��
Private Sub CaptureFrameSelectImage()
    Dim imgResult As DicomImage
    Dim imgs As New DicomImages
    Dim iPlane As Integer
    Dim dblZoom As Double
    Dim iLeft As Integer
    Dim iRight As Integer
    Dim iTop As Integer
    Dim iBottom As Integer
    Dim iMax As Integer
    Dim img As DicomImage
    Dim lblFrame As DicomLabel
    
    If Me.dcmView.Images.Count <> 1 Then Exit Sub
    If Me.dcmView.Images(1).Labels.Count < 1 Then Exit Sub
    
    Set img = Me.dcmView.Images(1)
    Set lblFrame = Me.dcmView.Images(1).Labels(Me.dcmView.Images(1).Labels.Count)
    
    If Abs(lblFrame.Width) = 0 Or Abs(lblFrame.Height) = 0 Then
        MsgBoxD Me, "��ѡ��ͼ��������ٱ���", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    'ͼ�������=300
    iMax = 300
    
    '����label����ȡ����ѡ�е�ͼ��
    'ͼ��λ��,�ڰ�ͼ��Ϊ1����ɫͼ��Ϊ3
    iPlane = 1
    If Not IsNull(img.Attributes(&H28, &H4).value) And img.Attributes(&H28, &H4).Exists Then
        If img.Attributes(&H28, &H4).value = "RGB" Then
            iPlane = 3
        End If
    End If
    
    'ͼ����λ��
    If lblFrame.Width >= 0 Then
        iLeft = lblFrame.Left
        iRight = iLeft + lblFrame.Width
    Else
        iLeft = lblFrame.Left + lblFrame.Width
        iRight = lblFrame.Left
    End If
    
    If lblFrame.Height >= 0 Then
        iTop = lblFrame.Top
        iBottom = iTop + lblFrame.Height
    Else
        iTop = lblFrame.Top + lblFrame.Height
        iBottom = lblFrame.Top
    End If
    
    '���ƽ��ͼ��Ĵ�С��300*300֮��
    If (iRight - iLeft) > iMax Or (iBottom - iTop) > iMax Then
        dblZoom = iMax / (iRight - iLeft)
        If dblZoom > iMax / (iBottom - iTop) Then dblZoom = iMax / (iBottom - iTop)
    Else
        dblZoom = 1
    End If
    
    img.Labels(img.Labels.Count).Visible = False
    If (img.RotateState = doRotateLeft And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipHorizontal) Then
        'X����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.SizeX - iRight, img.SizeX - iLeft, iTop, iBottom)
    ElseIf (img.RotateState = doRotateLeft And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipVertical) Then
        'Y����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, img.SizeY - iBottom, img.SizeY - iTop)
    ElseIf (img.RotateState = doRotateRight And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateLeft And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipBoth) Then
        'X��Y����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.SizeX - iRight, img.SizeX - iLeft, img.SizeY - iBottom, img.SizeY - iTop)
    Else
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, iTop, iBottom)
    End If
    
    '��imgResultһ��Ψһ�� InstanceUID
    imgResult.InstanceUID = dcmglbUID.NewUID
    
    '�ѽ��ͼ���뵽viewer�в��ұ���
    '����ͼ���DICOM����
    subWriteDicomPara imgResult, mlngAdviceID
    
    Dim dcmTag As New clsImageTagInf
    dcmTag.Tag = IMGTAG
    
    Set imgResult.Tag = dcmTag
    
    '��ͼ����뵽����ͼ��
    subInsert2Mini imgResult
    
    '����ͼ�񣬲�����ͼ��洢�¼�
    Call subSaveImage
    
    If ucPreview.CurImageCount = 1 Then
        Call SendMsgToMainWindow(Me, wetCaptureFirstImg, mlngAdviceID)
    End If
End Sub


Private Sub ucCapHook_OnKeyBoardLHook(ByVal lngVkCode As Long, ByVal lngScanCode As Long, ByVal lngFlags As Long)
On Error GoTo errHandle
    Select Case lngVkCode
        Case 66
            '�жϼ��̰����Ƿ��ɿ���Ϊ0��ʾ���¼���
            If lngScanCode = 128 Then
                'ִ�п�ݲɼ�
'                Call CaptureImage

                If timerHook.Enabled Or mblnCurCaptureState Then Exit Sub
                timerHook.Enabled = True
            End If
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucPreview_OnClick(ByVal lngSelectedIndex As Long)

    mCorpSize.X = 0
    mCorpSize.Y = 0
    
    '��ѡ��ͼ����ʾ���
    If lngSelectedIndex <> 0 Then
        
        '����ѡ��ͼ��װ�ص�dcmView��
        dcmView.Images.Clear
        dcmView.Images.Add ucPreview.ImgViewer.Images(lngSelectedIndex)

        '��ʾdcmView������picVideo
        dcmView.CurrentImage.BorderWidth = 0
        
        'ʹͼ�������ʾ�������������϶�ͼ��
        Dim dblTempZoom As Double
              
        dblTempZoom = dcmView.CurrentImage.ActualZoom
        dcmView.CurrentImage.StretchToFit = False
              
        Call subCenterZoom(dcmView.CurrentImage, dcmView, dblTempZoom, mCorpSize)
        
        '������Ƶ�ĵ�ǰ��ʾ״̬
        Call ConfigVideoShowState(False)
    End If
    
    '�ָ���ǰ�ؼ����㣬�Ա��ܹ�����ͼ��
    ucPreview.SetFocus
End Sub


Private Sub ucPreview_OnDbClick(ByVal lngSelectedIndex As Long, blnContinue As Boolean)
'˫��ʱ��������Ƶ�ļ�
On Error GoTo errHandle
    If lngSelectedIndex <= 0 Or lngSelectedIndex > ucPreview.CurImageCount Then Exit Sub
    
    If Not ucPreview.SelectImage.Tag Is Nothing Then
        If UCase(TypeName(ucPreview.SelectImage.Tag)) = UCase("clsImageTagInf") Then
            If ucPreview.SelectImage.Tag.Tag = VIDEOTAG Or ucPreview.SelectImage.Tag.Tag = AUDIOTAG Then
                Call subVideoPlay
                blnContinue = False
            End If
        End If
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

