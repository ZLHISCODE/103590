VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVideoProcess.ocx"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "*\A..\zl9PacsControl\zl9PacsControl.vbp"
Begin VB.Form frmWork_Video 
   BorderStyle     =   0  'None
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   10410
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmWork_Video.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Tag             =   "��Ƶ�ɼ�"
   Begin zl9PacsControl.ucSplitter ucSplitter1 
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   4620
      Width           =   10410
      _extentx        =   18362
      _extenty        =   238
      mousepointer    =   7
      splittype       =   0
      splitlevel      =   3
      dbclicktype     =   2
      con1minsize     =   3000
      con2minsize     =   1000
      control1name    =   "picCapture"
      control2name    =   "picReportImage"
   End
   Begin ScanLibCtl.ImgScan ImageScanner 
      Left            =   4080
      Top             =   6360
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
      StopScanBox     =   -1  'True
      FileType        =   3
      CompressionType =   0
      CompressionInfo =   0
      ScanTo          =   4
   End
   Begin VB.PictureBox picReportImage 
      Height          =   4125
      Left            =   0
      ScaleHeight     =   4065
      ScaleWidth      =   10350
      TabIndex        =   13
      Top             =   4755
      Width           =   10410
      Begin VB.PictureBox picCacheImage 
         BackColor       =   &H8000000D&
         Height          =   3375
         Left            =   4800
         ScaleHeight     =   3315
         ScaleWidth      =   3795
         TabIndex        =   16
         Top             =   720
         Width           =   3855
         Begin MSComCtl2.DTPicker DTPImg 
            Height          =   255
            Left            =   360
            TabIndex        =   19
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            Format          =   206241793
            CurrentDate     =   42587
         End
         Begin VB.ComboBox cboCacheImage 
            Height          =   300
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   120
            Width           =   1215
         End
         Begin zl9PacsControl.ucImagePreview ucCacheImage 
            Height          =   2565
            Left            =   960
            TabIndex        =   17
            Top             =   600
            Width           =   1890
            _extentx        =   3334
            _extenty        =   4524
            backcolor       =   4210752
         End
      End
      Begin VB.PictureBox picPreview 
         BackColor       =   &H8000000D&
         Height          =   2535
         Left            =   480
         ScaleHeight     =   2475
         ScaleWidth      =   2955
         TabIndex        =   14
         Top             =   600
         Width           =   3015
         Begin zl9PacsControl.ucImagePreview ucPreview 
            Height          =   2205
            Left            =   480
            TabIndex        =   15
            Top             =   120
            Width           =   1890
            _extentx        =   3334
            _extenty        =   4524
            backcolor       =   4210752
         End
      End
      Begin XtremeDockingPane.DockingPane dkpReportImage 
         Left            =   3600
         Top             =   360
         _Version        =   589884
         _ExtentX        =   450
         _ExtentY        =   423
         _StockProps     =   0
      End
   End
   Begin VB.Timer tmrReg 
      Interval        =   10000
      Left            =   1560
      Top             =   6120
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   720
      Top             =   4950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrComm 
      Interval        =   2
      Left            =   0
      Top             =   5040
   End
   Begin VB.Timer timerHook 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15
      Top             =   6090
   End
   Begin VB.Timer timerRePaint 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   135
      Top             =   6780
   End
   Begin VB.PictureBox picCapture 
      ForeColor       =   &H00000000&
      Height          =   4620
      Left            =   0
      ScaleHeight     =   4560
      ScaleWidth      =   10350
      TabIndex        =   2
      Top             =   0
      Width           =   10410
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
         TabIndex        =   11
         Top             =   3840
         Width           =   7335
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
         TabIndex        =   10
         Top             =   0
         Width           =   75
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
         Height          =   75
         Index           =   0
         Left            =   360
         MousePointer    =   7  'Size N S
         ScaleHeight     =   75
         ScaleWidth      =   7335
         TabIndex        =   8
         Top             =   120
         Width           =   7335
      End
      Begin VB.PictureBox picView 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   600
         ScaleHeight     =   3495
         ScaleWidth      =   6855
         TabIndex        =   3
         Top             =   240
         Width           =   6855
         Begin ZLDSVideoProcess.DSCapture wdmCapture 
            Height          =   3135
            Left            =   720
            TabIndex        =   4
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
         Begin VB.PictureBox picVideo 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   3015
            Left            =   1200
            ScaleHeight     =   201
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   224
            TabIndex        =   6
            Top             =   120
            Width           =   3360
         End
         Begin VB.TextBox txtInputText 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   5520
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   840
            Visible         =   0   'False
            Width           =   975
         End
         Begin DicomObjects.DicomViewer dcmView 
            Height          =   1575
            Left            =   4440
            TabIndex        =   7
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
      Begin DicomObjects.DicomViewer dcmAfter 
         Height          =   735
         Left            =   8820
         TabIndex        =   12
         Top             =   3195
         Visible         =   0   'False
         Width           =   1035
         _Version        =   262147
         _ExtentX        =   1826
         _ExtentY        =   1296
         _StockProps     =   35
         BackColor       =   0
         UseScrollBars   =   0   'False
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
   Begin MSCommLib.MSComm commListener 
      Bindings        =   "frmWork_Video.frx":06EA
      Left            =   0
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picTemp2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   8280
      ScaleHeight     =   1455
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   1200
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






Private Const M_STR_HINT_NoSelectData As String = "��Ч�ļ�����ݣ���ѡ����Ҫִ�еļ���¼��"



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

'���������Ļ�����Ϣ
Private Type TUnLockStudyInf
    lngAdviceId As Long
    lngSendNo As Long
    blnMoved As Boolean
    lngStudyState As Long
End Type

'��ǰ���ɼ�ͼ����Ļ�����Ϣ
Private Type TCurStudyBaseInf
    strStudyUid As String          '���UID
    strModality As String          'Ӱ�����
    strSex As String               '�Ա�
    strBirthDate As String         '��������
    strAge As String               '����
    strName As String              '����
    strCheckNo As String           '����
    strPatientID As String         '����ID
End Type


'��̨�ɼ���Ϣ
Private Type TAfterCaptureInf
    intAfterTag As Integer    '��ʶ
    strAfterStudyUid As String     '��̨�ɼ����UID
    strAfterSeriesUid As String    '��̨�ɼ�����UID

    lngAfterCurImageCount As Long  '��ǰ��̨�ɼ�ͼ������
    strAfterParentTitle As String  '��̨�ɼ���Ϣ
End Type

'COM��̤�˿�״̬
Private Type TComPortState
    intComState As Integer          'COM�ڵ�״̬
    lngComTime As Long              '��¼com�ڱ���״̬��ʱ��
    dtLastCapture As Date           '�����̤���µ�ʱ��
    blnCTSHolding As Boolean        '��¼��̬ʱ��CTS�ߵĵ�ƽ
End Type


Private mdcmTmpImg As DicomImage
Private mintCaptureFlag As Integer

Private mobjCustomDevice As Object  'ר����Ƶ�ɼ�����

Private dcmglbUID As New DicomGlobal    '����UIDRoot=1

Private WithEvents mobjDxDevice As clsDxHidDevice   'ʵ�������ֱ�֮��Ĳɼ���ʽ
Attribute mobjDxDevice.VB_VarHelpID = -1
'Private WithEvents mobjHotHook As clsHookKey

Public mhCapWnd As Long                 '�ɼ����ڵľ��
Private WithEvents mfrmParameter As frmVideoSetup
Attribute mfrmParameter.VB_VarHelpID = -1
Private mfrmOpenStudy As frmOpenStudyList
Private mstrAfterStationName As String

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

Private mstrNameInf As String

Private mblnMoveDown  As Boolean        '�����ж��Ƿ���������
Private mblnDcmViewDown As Boolean      '�����ж�dcmView������Ƿ񱻰���
Private mintCurImgIndex As Integer      '��ǰ��ѡ�е�ͼ������
Private mdcmSelectLabel As DicomLabel   '��ǰ��ѡ�еı�ע

Private mstrAviFileName As String       '¼���ļ���
Private mstrEncoderName As String       '
Private mstrBufferDir As String
Private mstrAfterDir As String

Private mcpsComState As TComPortState       'Com�˿�ʹ��״̬

Private mblnUseClipbord As Boolean          '�Ƿ�ʹ�ü�����


Private mobjFtpConnection As New clsFtp
Private mobjBakFtpConnection As New clsFtp

Private mobjFtp As TFtpDeviceInf        'ftp�豸��Ϣ
Private mobjBakFtp As TFtpDeviceInf     'ftp���ݴ洢�豸��Ϣ


Private mblnReadOnly As Boolean         '�Ƿ�ֻ�ܲ鿴True�鿴ģʽ��False�ɼ�ģʽ

Private mblnShowProcessBar As Boolean   '�Ƿ���ʾ��������


'���˻�����Ϣ����
Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long

Private mlngAdviceId As Long            'ҽ��ID
Private mlngSendNo As Long
Private mblnMoved As Boolean            '�Ƿ�ת��
Private mlngStudyState As Long



Private mAfterCaptureInf As TAfterCaptureInf    '��̨�ɼ���Ϣ
Private mSelectStudyInf As TUnLockStudyInf      '�����ļ����Ϣ
Private mcurStudyInf As TCurStudyBaseInf        '��ǰ�����Ϣ

Private mVideoCapture As clsVideoCapture '��Ƶ�ɼ�����

Private mdblZoomRate As Double  '���ű��ʣ���cbrMain��cbrMain_ResizeClient�¼�����Ҫ���¼����ֵ��
Private mVideoSize As TVideoSize '��Ƶ��С������ص���Ƶ������棩
Private mCurCutRange As TCutRange '��Ƶ�ü���Χ���ã��ò���ͨ��GetString��SaveString������ע����У�
Private mVideoArea As TVideoArea  '��Ƶ�ͻ��������ã���cbrMain��cbrMain_ResizeClient�¼�����Ҫ���¼����ֵ��

Private mblnCaptureLockState As Boolean '��Ƶ����״̬

Private mstrInstitution As String       '��λ����

Private Const M_LNG_REFRESHINTERVAL As Long = 600 'ˢ�¼��
Private mstrVideoRegTime As String '������Ƶ����ע��ʱ��
Private mstrMsg As String
Private mblnRefreshState As Boolean
Private mblnInitState As Boolean

Private mrsImageCache As New ADODB.Recordset
Private mobjFile As New FileSystemObject
Private mlngReleationType As Integer    '1--������2--����
Private mdcmUID As New DicomGlobal

Private mdate As Date 'DTP�ؼ�ѡ���ʱ��,���ڹ��˺�̨ͼ
Private mblIsNeedRefresh '�ж��Ƿ���Ҫ���¼��غ�̨ͼ(���ڱ����뱨��ģ��������ͼ��һ��)
Private mintTagMax As Integer '����ʶ
Private mintTagNow As Integer '��ǰ��ʶ
Private mintFontSize As Integer '�ֺ�
Private mblnIsReported As Boolean   '�Ƿ��Ѿ����ڱ�����
Private mstrInstanceUID As String

Private mblnCompareSize As Boolean

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

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Public Event OnIsUseAfterImgChange(ByVal blUse As Boolean)
Public Event OnStateChange(ByVal lngEventType As TVideoEventType, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String)
Public Event OnControlResize(objControl As Object)
Public Event OnImgLoadState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean)

Property Get CaptionEx() As String
    CaptionEx = Me.Tag
End Property

Property Let CaptionEx(value As String)
    Dim hWndParent As Long
    
    Me.Tag = value
    
    hWndParent = GetParent(Me.hwnd)
    
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

'�����Ĳ�������
Property Get LockPatientName() As String
    LockPatientName = mstrNameInf
End Property

'��ǰ����״̬
Property Get LockState() As Boolean
    LockState = mblnCaptureLockState
End Property

Private Sub UnLockStudy()
'�������
    mblnCaptureLockState = False
End Sub

'�Ƿ��Ѿ����ڱ�����
Property Get IsReported() As Boolean
    IsReported = mblnIsReported
End Property
Property Let IsReported(ByVal value As Boolean)
    mblnIsReported = value
End Property

Public Sub NotificationRefresh()
'֪ͨˢ��
    mblnRefreshState = False
End Sub

Private Function GetTag(ByVal FolderName As String, ByRef strType As String) As Integer
'�����ļ��������еı�ʶ�ţ�FolderName��Ŀ��Ŀ¼����strType�� ���ء���ʶ�� �� ����顱
On Error GoTo errH
    Dim i As Integer
    Dim strTmp As String
    
    strType = Mid(FolderName, 1, 2)
    strTmp = Mid(FolderName, 3, Len(FolderName) - 2)
    i = InStr(strTmp, "-")
    GetTag = Val(Mid(strTmp, 1, i - 1))
    
    Exit Function
errH:
    GetTag = 0
End Function

Private Function GetStudyUIDFromFolderName(ByVal FolderName As String) As String
'�����ļ��������еļ��UID�����أ����������ļ�����
On Error GoTo errH
    Dim i As Integer
    Dim j As Integer
    
    i = InStr(FolderName, "-")
    j = Len(FolderName)
    
    GetStudyUIDFromFolderName = Mid(FolderName, i + 1, j - i)
    Exit Function
errH:
    GetStudyUIDFromFolderName = FolderName
End Function

Private Function LoadMiniCache() As Boolean
    Dim i As Integer
    Dim strQueryPath As String
    
    Dim objFolder2 As Folder, objFolder3 As Folder, objFolder4 As Folder
    Dim strStudyUid As String, strSeriesUID As String
    Dim lngStudyNo As Long, lngSeriesNo As Long

    Dim strAfterTime As String
    Dim dtChose As Date
    Dim intTmp As Integer
    Dim strTag As String  '��λ���ı�ʶ
    Dim strType As String
    Dim curDate As Date
        
    If gobjCapturePar.IsUseAfterCapture = False Then Exit Function
    
    curDate = zlDatabase.Currentdate
    dtChose = mdate

    Set mrsImageCache = New ADODB.Recordset
    mrsImageCache.Fields.Append "����", adVarChar, 100
    mrsImageCache.Fields.Append "����", adVarChar, 64
    mrsImageCache.Fields.Append "���UID", adVarChar, 64
    mrsImageCache.Fields.Append "���к�", adVarChar, 18
    mrsImageCache.Fields.Append "����UID", adVarChar, 64
    mrsImageCache.Fields.Append "�������", adVarChar, 20

    mrsImageCache.Fields.Append "·��", adVarChar, 4000
    mrsImageCache.Open
    
    If mobjFile.FolderExists(mstrAfterDir) = False Then Exit Function
    
    If mobjFile.GetFolder(mstrAfterDir).SubFolders.Count > 0 Then
        For Each objFolder2 In mobjFile.GetFolder(mstrAfterDir).SubFolders

                        If InStr(objFolder2.Name, Format(mdate, "yyyymmdd")) > 0 Then ''�������ѡ���ʱ��������

                                If objFolder2.SubFolders.Count > 0 Then
                                
                                        For Each objFolder3 In objFolder2.SubFolders                            '���UID��
                                        
                                                If objFolder3.SubFolders.Count >= 0 Then

                                                        strAfterTime = Format(objFolder3.DateCreated, "YYYY-MM-DD HH:MM:SS")

                                                        strStudyUid = GetStudyUIDFromFolderName(objFolder3.Name)
                                                                                                                  
                                                        lngStudyNo = lngStudyNo + 1
                                                        lngSeriesNo = 0
                                                        
                                                        For Each objFolder4 In objFolder3.SubFolders                    '����UID��
                                                                
                                                                strSeriesUID = objFolder4.Name
                                                                lngSeriesNo = lngSeriesNo + 1
                                                                
                                                                mrsImageCache.AddNew
                                                                strTag = Lpad(GetTag(objFolder3.Name, strType), 3, "0")
                                                                mrsImageCache!���� = strType & strTag
                                                                mrsImageCache!���� = lngStudyNo
                                                                mrsImageCache!���uid = strStudyUid
                                                                mrsImageCache!���к� = lngSeriesNo
                                                                mrsImageCache!����Uid = strSeriesUID
                                                                mrsImageCache!������� = strAfterTime
                                                                mrsImageCache!·�� = objFolder4.Path
                                                                mrsImageCache.Update
                                                                
                                                        Next
                                                        
                                                End If
                    Next
                                    Exit For '��ʱ�Ѿ���������ѡʱ�䣬����ʱ��ѡ��
                                End If
                        End If 'ʱ��ѡ��
        Next
    End If
    
    If mrsImageCache.RecordCount > 0 Then
        mrsImageCache.Sort = "������� desc"
        mrsImageCache.MoveFirst
    End If

    cboCacheImage.Clear
    ucCacheImage.ImgViewer.Images.Clear
    
    For i = 0 To mrsImageCache.RecordCount - 1
        If i = 0 Then strQueryPath = Nvl(mrsImageCache!·��)

        cboCacheImage.AddItem Nvl(mrsImageCache!����) & "     ʱ�䣺" & Format(Nvl(mrsImageCache!�������), "HH:MM:SS")

        mrsImageCache.MoveNext
    Next
    
    If mrsImageCache.RecordCount > 0 Then
        If cboCacheImage.ListIndex < 0 Then
            cboCacheImage.ListIndex = 0
        End If
    End If
  
End Function

Private Sub cboCacheImage_Click()
    Dim strQueryPath As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    
    ucCacheImage.ClearCurrentPageImage
    
    Set rsTmp = mrsImageCache
    rsTmp.Filter = ""
        
    If mrsImageCache.RecordCount <= 0 Then Exit Sub
    
    mrsImageCache.MoveFirst
    Set rsTmp = mrsImageCache

    rsTmp.Filter = "����='" & Trim(Mid(cboCacheImage.Text, 1, 5)) & "'"

    If rsTmp.RecordCount < 1 Then Exit Sub
    strQueryPath = Nvl(rsTmp!·��)
    
    If strQueryPath = "" Then Exit Sub
    Call ucCacheImage.RefreshImage(slLocal, strQueryPath, mblnMoved)
    Call ucCacheImage.RedrawSelf
    Call dkpReportImage.RedrawPanes
    
    
    Exit Sub
errH:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
    err.Clear
End Sub

Private Sub DTPImg_Change()
    On Error GoTo errH
        
    mdate = DTPImg.value
    Call LoadMiniCache
    ucCacheImage.RedrawSelf
    Call dkpReportImage.RedrawPanes
        
    Exit Sub
errH:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub Form_Initialize()
'��ʼ��ģ�����
    mblnInitState = False
End Sub




Public Sub ShowVideoConfig()
On Error GoTo errHandle
'��Ƶ����

BUGEX "ShowVideoConfig 1"
    '�ȱ����޸ĵĲ�������
    Call SaveParameterCfg
BUGEX "ShowVideoConfig 2"

    '�ж��Ƿ���ʵʱģʽ��ʾ״̬
    If mblnRealTime = False Then
        Call ConfigVideoShowState(True)
    End If
    
    '�򿪲������ô���
    If mfrmParameter.ShowParameterConfig(mVideoCapture, Me) = False Then Exit Sub
    
    '���¶�ȡ���ò���------------------------------------------
BUGEX "ShowVideoConfig 3"
    Call InitParameter
    
BUGEX "ShowVideoConfig 4"
    Call ConfigFtpStorageDevice(gobjCapturePar.CurStorageDeviceNo, gobjCapturePar.BakStorageDeviceNo)

BUGEX "ShowVideoConfig 5"
    If gobjCapturePar.IsUseAfterCapture Then
        Call UpdateAfterCaptureInfo
    Else
        Call ShowAfterCaptureInf(False)
    End If
    
BUGEX "ShowVideoConfig 6"
    Call OpenComm
    
    If gobjCapturePar.VideoDirverType = vdtCustom Then Call InitCustomDevice
    
    gstrHotKeyTest = GetSetting("ZLSOFT", "����ģ��", "�ɼ��ȼ�", "F8")
    
''    ����ע��ȫ���ȼ�
'    If gobjCapturePar.strCaptureHot <> "" Then
'        Call mobjHotHook.EnableHook(WM_KEYDOWN, True)
'    Else
'        Call mobjHotHook.FreeHook
'    End If
    '----------------------------------------------------------
    
BUGEX "ShowVideoConfig End"
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
    err.Clear
End Sub


Private Sub InitParameter()
'��ʼ����������
    Dim rsTmp As New ADODB.Recordset
    Dim intVideoCapture As Integer
    Dim strSQL As String

    mintCaptureFlag = 0
    mblnRealTime = True
    mintCurImgIndex = 0
    mblnPlayVideo = False
    
    mstrInstitution = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")

    mAfterCaptureInf.strAfterParentTitle = ""

    '��������ڴ��̵ĸ�Ŀ¼��app.pathΪ��x:\��
    mstrBufferDir = GetAppPath & "\TmpImage\"
    mstrAfterDir = GetAppPath & "\TmpAfterImage\"
    
    mstrAviFileName = mstrBufferDir & "TmpVideo.avi"
    
    gint��Ƶ�豸���� = getLicenseCount(LOGIN_TYPE_��Ƶ�豸)
    
    mblnUseClipbord = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "UseClipbord", 0)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "UseClipbord", IIf(mblnUseClipbord, 1, 0))
    
    timerRePaint.Interval = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "�����ػ���", 500))
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "�����ػ���", timerRePaint.Interval)
    
    mblnCompareSize = IIf(Val(GetSetting("ZLSOFT", "����ģ��\Ftp", "����FTP�ļ���С�Ա�", 1)) <> 0, True, False)
    Call SaveSetting("ZLSOFT", "����ģ��\Ftp", "����FTP�ļ���С�Ա�", IIf(mblnCompareSize, 1, 0))

    '��ȡ�ü�����
    mCurCutRange.LeftRate = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblX1Scale", 0))  'ʹ��mdblX1Scale������Ϊ�˱�֤����ǰ�Ĳ������ü���
    mCurCutRange.WidthRate = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblX2Scale", 0))
    mCurCutRange.TopRate = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblY1Scale", 0))
    mCurCutRange.HeightRate = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblY2Scale", 0))

    If (mCurCutRange.LeftRate >= 1) Or (mCurCutRange.LeftRate < 0) Then mCurCutRange.LeftRate = 0
    If (mCurCutRange.WidthRate >= 1) Or (mCurCutRange.WidthRate < 0) Then mCurCutRange.WidthRate = 0
    If (mCurCutRange.TopRate >= 1) Or (mCurCutRange.TopRate < 0) Then mCurCutRange.TopRate = 0
    If (mCurCutRange.HeightRate >= 1) Or (mCurCutRange.HeightRate < 0) Then mCurCutRange.HeightRate = 0

    '����UIDRoot=1
    dcmglbUID.RegString("UIDRoot") = "1"
    
    '��ȡ�ɼ����ò���
    If gobjCapturePar Is Nothing Then
        Set gobjCapturePar = New clsCaptureParameter
    End If
    
    Call gobjCapturePar.ReadParameter

    '����ƶ�ʱ����ʾ��ͼ
    ucPreview.BigImageWay = gobjCapturePar.ImagePreview
    ucPreview.PreviewTime = gobjCapturePar.PreviewTime
    ucPreview.ImgLoadType = gtFileLoadType
    ucPreview.DoShield = True
    
    If gobjCapturePar.IsAllowChangeSize = False Then
        Me.pbxSize.Item(0).MousePointer = 0
        Me.pbxSize.Item(1).MousePointer = 0
        Me.pbxSize.Item(2).MousePointer = 0
        Me.pbxSize.Item(3).MousePointer = 0
    Else
        Me.pbxSize.Item(0).MousePointer = 7
        Me.pbxSize.Item(1).MousePointer = 7
        Me.pbxSize.Item(2).MousePointer = 9
        Me.pbxSize.Item(3).MousePointer = 9
    End If

    '�������б���ͼ��
    ucPreview.OnlyLoadReportImage = False
End Sub


Private Sub ConfigFtpStorageDevice(ByVal strCurStorageNo As String, ByVal strBakStorageNo As String)
'����ftp�洢�豸
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    '�������ߴ洢�豸��Ϣ
    strSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�洢�豸", strCurStorageNo)
    
    mobjFtp.strDeviceId = ""
    If rsTmp.EOF Then
        MsgboxCus "Ӱ��洢�豸δ�������ͣ�ã����飡", vbInformation, G_STR_HINT_TITLE
        
        mobjFtp.strDeviceId = ""
        mblnReadOnly = True
        Exit Sub
    End If
    
    mobjFtp.strDeviceId = strCurStorageNo
    Call funGetFtpDeviceInf(Me, mobjFtp)
    

    '���ñ����豸��Ϣ
    mobjBakFtp.strDeviceId = ""
    If Val(strBakStorageNo) > 0 Then
        strSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����豸", strBakStorageNo)
        
        If rsTmp.EOF Then
            MsgboxCus "δȡ����Ч�ı����豸��Ϣ�����ܶԲɼ�ͼ����б��ݲ��������鱸���豸�����Ƿ���ȷ��", vbInformation, G_STR_HINT_TITLE
            
            Exit Sub
        End If
        
        mobjBakFtp.strDeviceId = strBakStorageNo
        Call funGetFtpDeviceInf(Me, mobjBakFtp)
    End If
    
End Sub

Public Sub zlRePreview()
'���½�����ƵԤ��
    If mVideoCapture.IsStartup Then
        Call mVideoCapture.StopPreview
        Call mVideoCapture.StartPreview
    End If
End Sub

Public Sub zlInitModule()
BUGEX "zlPacsCapture zlInitModule 0"
'��ʼ��ģ�����
    
    '��ʼ������
    Call InitParameter
    
BUGEX "gobjCapturePar.CurStorageDeviceNo = " & gobjCapturePar.CurStorageDeviceNo
    '����ftp�洢�豸
    Call ConfigFtpStorageDevice(gobjCapturePar.CurStorageDeviceNo, gobjCapturePar.BakStorageDeviceNo)

BUGEX "zlInitModule 1"
    '����Ƶ�ɼ��豸
    Call OpenVideoCaptureDevice

BUGEX "zlInitModule 2"
    '���º�̨�ɼ���Ϣ
    If gobjCapturePar.IsUseAfterCapture Then Call UpdateAfterCaptureInfo
    
BUGEX "zlInitModule 3"
    '��ʼ��ר����Ƶ�ɼ��ӿ�
    Call InitCustomDevice
BUGEX "zlInitModule End"
    mblnInitState = True
End Sub

Private Sub InitCustomDevice()
    Dim strCustomDeviceDir As String        'ר����Ƶ�ɼ�����·��
    Dim strCustomDeviceDllName As String    'ר����Ƶ�ɼ���������
    Dim objFile As New FileSystemObject
    
    '��ʼ��ר����Ƶ�ɼ��ӿ�
    strCustomDeviceDir = gobjCapturePar.CustomDevicePath
    If strCustomDeviceDir <> "" Then
        strCustomDeviceDllName = Trim(Replace(objFile.GetFileName(strCustomDeviceDir), ".dll", ""))
        
        Set mobjCustomDevice = CreateObject(strCustomDeviceDllName & ".cls" & strCustomDeviceDllName)
        
        If Not mobjCustomDevice Is Nothing Then
            Call mobjCustomDevice.zlInit(gcnVideoOracle, UserInfo.id, glngDepartId, glngRootHandle)
        End If
    End If
End Sub


'----------------------------------------------------------------------------------------------------------
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Public Sub zlUpdateAdviceInf(ByVal lngAdviceId As Long, _
                            ByVal lngSendNo As Long, _
                            ByVal lngStudyState As Long, _
                            ByVal blnMoved As Boolean)
'����ҽ����Ϣ
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    '����������ĵ�ǰ�����Ϣ
    mSelectStudyInf.lngAdviceId = lngAdviceId
    mSelectStudyInf.blnMoved = blnMoved
    mSelectStudyInf.lngSendNo = lngSendNo
    mSelectStudyInf.lngStudyState = lngStudyState
    
    If mblnCaptureLockState Then Exit Sub
    
    mlngAdviceId = lngAdviceId
    mlngSendNo = lngSendNo
    mblnMoved = blnMoved
    mlngStudyState = lngStudyState
    
    mblnReadOnly = False
    mblnRefreshState = True
    
    '���ݱ�ת��ʱ��û��Ȩ��ʱ��״̬Ϊָ��״̬ʱ����ģ��Ϊֻ��
    If mlngAdviceId <= 0 Or blnMoved Or lngStudyState = 6 Or lngStudyState = 0 Or lngStudyState = 1 Or InStr(gstrPrivs, "��Ƶ�ɼ�") <= 0 Then
        mblnReadOnly = True
    End If
    
    '��ȡ���˻�����Ϣ,дDICOM����ʱ��
    strSQL = "Select A.Ӱ�����,A.����,A.�Ա�,A.����,A.��������,A.����,A.����,A.���UID,B.����ID " & _
                " From Ӱ�����¼ A,����ҽ����¼��B " & _
                " Where A.ҽ��ID=B.Id And A.ҽ��ID=[1]"
                
    If mblnMoved Then
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˻�����Ϣ", lngAdviceId)
    
    If Not rsTemp.EOF Then
        mcurStudyInf.strStudyUid = Nvl(rsTemp("���UID"))
        mcurStudyInf.strModality = Nvl(rsTemp("Ӱ�����"))
        mcurStudyInf.strSex = Nvl(rsTemp("�Ա�"))
        mcurStudyInf.strAge = Nvl(rsTemp("����"))
        mcurStudyInf.strBirthDate = Nvl(rsTemp("��������"))
        mcurStudyInf.strName = Nvl(rsTemp("����"))
        mcurStudyInf.strCheckNo = Nvl(rsTemp("����"))
        mcurStudyInf.strPatientID = Nvl(rsTemp("����ID"))
        
        mstrNameInf = Nvl(rsTemp("����"))
        
        mcurStudyInf.strSex = IIf(mcurStudyInf.strSex = "��", "M", IIf(mcurStudyInf.strSex = "Ů", "F", "O"))
    Else
        mcurStudyInf.strStudyUid = ""
        mcurStudyInf.strModality = ""
        mcurStudyInf.strSex = ""
        mcurStudyInf.strAge = ""
        mcurStudyInf.strCheckNo = ""
        mcurStudyInf.strPatientID = ""
        mcurStudyInf.strBirthDate = ""
        mcurStudyInf.strName = ""
        
        mstrNameInf = ""
    End If
    
    Me.Tag = "ͼ��ɼ�" & IIf(mstrNameInf <> "", "(" & mstrNameInf & ")", "")
    Me.CaptionEx = Me.Tag
End Sub


Public Sub zlRefreshFace(Optional blnForceRefresh As Boolean = False)
'ˢ�½���
On Error GoTo errHandle
    Dim rsTemp As ADODB.Recordset
    Dim iRows As Integer
    Dim iCols As Integer
    Dim strStudyUid As String
    
BUGEX "zlRefreshFace 0"
    If (mlngTmpAdviceId = mlngAdviceId And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh And Not mblIsNeedRefresh Then Exit Sub

BUGEX "zlRefreshFace 0.1"
    mlngTmpAdviceId = mlngAdviceId
    mlngTmpSendNo = mlngSendNo
    mblnRefreshState = True

BUGEX "zlRefreshFace 1"
    Call ConfigVideoShowState(True)

BUGEX "zlRefreshFace 2"
'����ˢ�¹������ڱ�֤�뱨��ģ�����ʾ����ͬ��
    Call ucPreview.RefreshImage(slStudy, mcurStudyInf.strStudyUid, mblnMoved, blnForceRefresh, False)
    Call ClearEmptyFolder(False) 'ɾ����Ŀ¼
    Call LoadMiniCache
    mblIsNeedRefresh = False
    
    '�л��� ���ͼtab
    dkpReportImage.Panes(1).Selected = True
    
BUGEX "zlRefreshFace 3"
    If ucPreview.ImgViewer.Images.Count > 0 Then
BUGEX "zlRefreshFace 4"
        '����ѡ��ͼ��װ�ص�dcmView��
        Call PreviewThumbnail(ucPreview.SelectIndex)
BUGEX "zlRefreshFace 5"
        '�����Twain��ר����Ƶ�ɼ�ģʽ��������mblnRealTimeΪfalse
        If IsTwainCaptureWay = True Or IsCustomCaptureWay Then mblnRealTime = False
    Else
        Call dcmView.Images.Clear
    End If
BUGEX "zlRefreshFace 6"
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub StopCapture()
'-----------------------------------------------------------------------------------------
'���ܣ�ֹͣ��ʾ��Ƶ�ɼ�,�ͷ���Ƶ�ɼ����ڣ�
'      �ͷŴ��������Ķ˿�
'��������
'���أ���
'-----------------------------------------------------------------------------------------
'    Call mobjHotHook.FreeHook
    
    '�ر�COMM��
    If commListener.PortOpen Then commListener.PortOpen = False
    
    '�ͷŲɼ��豸������
    If Not mVideoCapture Is Nothing Then
        Call mVideoCapture.StopPreview
    End If
    
    '����Midi�ӿ����������¼����
    If Not mobjDxDevice Is Nothing Then
        If mobjDxDevice.Handle <> 0 Then Call mobjDxDevice.CloseDxDevice
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
    
    Select Case control.id
        Case conMenu_Cap_Dynamic       '��̬��ʾ
            control.Checked = mblnRealTime
            control.Enabled = (Not mblnReadOnly) And (Not IsTwainCaptureWay And Not IsCustomCaptureWay) And mVideoCapture.IsStartup ' And (mhCapWnd <> 0) modify by tjh at 2010-01-20
            control.Visible = Not IsTwainCaptureWay And Not IsCustomCaptureWay
            
            If mblnRealTime Then
                control.IconId = conMenu_Cap_Dynamic
            Else
                control.IconId = 10023
            End If
            
        Case conMenu_Cap_MarkMap       'Ӱ��ɼ�
            control.Enabled = Not mblnReadOnly And (mVideoCapture.IsStartup Or IsTwainCaptureWay Or IsCustomCaptureWay)
            
        Case conMenu_Cap_After_Capture  '��̨�ɼ�
            control.Enabled = mVideoCapture.IsStartup
            control.Visible = gobjCapturePar.IsUseAfterCapture And (Not IsCustomCaptureWay)
            
        Case conMenu_Cap_Import        'Ӱ����
            control.Enabled = Not mblnReadOnly
            
        Case conMenu_Cap_DelImg  'Ӱ��ɾ��
            'dkpReportImage.FindPane(1).Selected = true ��ָ�·�չ�ֵı�ǩ��"���ͼ" �����ܣ�ֻҪ�����̨ͼ��ǩ����ť�Ͳ�����
            control.Enabled = (mblnRealTime = False) And (ucPreview.ImgViewer.Images.Count > 0) And (Not mblnReadOnly) And Me.Visible And dkpReportImage.FindPane(1).Selected = True
            
        Case conMenu_Cap_Record        '¼��
            control.Enabled = Not mblnReadOnly And ((gobjCapturePar.VideoDirverType = vdtWDM And mVideoCapture.IsStartup) Or gobjCapturePar.VideoDirverType = vdtCustom)
            control.Visible = Not IsTwainCaptureWay
            
        Case conMenu_Cap_After_Record   '��̨¼��
            control.Enabled = gobjCapturePar.VideoDirverType = vdtWDM And mVideoCapture.IsStartup
            control.Visible = Not IsTwainCaptureWay And Not IsCustomCaptureWay And gobjCapturePar.IsUseAfterCapture And False
            
        Case conMenu_Cap_Record_Stop 'ֹͣ¼�� modify by tjh at 2010-01-22
            control.Enabled = mblnRealTime And Not mblnReadOnly And (gobjCapturePar.VideoDirverType = vdtWDM) And mVideoCapture.IsStartup
            control.Visible = Not IsTwainCaptureWay And Not IsCustomCaptureWay
            
        Case conMenu_Cap_RecordAudio '¼��
            control.Enabled = Not mblnReadOnly
            control.Visible = Not IsCustomCaptureWay
            
'        Case conMenu_Cap_Full_Screen 'ȫ�� modify by tjh at 2010-01-22 (���ʹ���µ���Ƶ�ط��������������øù���)
'            control.Enabled = mblnRealTime And (Not mblnReadOnly) And Not GetIsTwainCaptureWay And mVideoCapture.IsStartup
'            control.Visible = Not GetIsTwainCaptureWay And mstrVideoRegTime <> ""
'
'        Case conMenu_Cap_DevSet        '���ã�������ڸ���״̬ʱ�������θð�ť�� modify by tjh at 2010-01-25
'            control.Enabled = gobjCapturePar.IsUseStartupVideo And mstrVideoRegTime <> ""  'mblnEmbedded ' And (Not mblnReadOnly)
'
'            '���Ϊ�������壬�����ظ����ð�ť
'            'control.Visible = mstrVideoRegTime <> ""
'            If Not (mParentContainer Is Nothing) Then
'                If TypeOf mParentContainer Is frmVideoDockWindow Then
'                    control.Enabled = False
'                Else
'                    control.Enabled = True
'                End If
'            End If
            
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
            
'        Case conMenu_Tool_Analyse
'            If mblnObserve Then
'                control.Enabled = Not mblnReadOnly
'            Else
'                control.Visible = False
'                control.Enabled = False
'            End If
'
            
        Case conMenu_Cap_OpenStudyList
            control.Enabled = True
            control.Visible = gobjCapturePar.IsUseCaptureLock
            
        Case conMenu_Cap_StudySyncState
            control.Enabled = Not mblnReadOnly Or mblnCaptureLockState
            control.Visible = gobjCapturePar.IsUseCaptureLock
            
        Case conMenu_Cap_After_Tag
            control.Enabled = mVideoCapture.IsStartup
            control.Visible = gobjCapturePar.IsUseAfterCapture
            
        ''''''''''''
        Case conMenu_Cap_ImgImport
            
        Case conMenu_Cap_ImgExport          '���͵���̨
            
        Case conMenu_Cap_ReUpLoad, conMenu_Cap_ReDownLoad
            control.Visible = gtFileLoadType = Service
            control.Enabled = ucPreview.CurImageCount > 0
    End Select
End Sub


''''''''''''''''''''''''''''''''''
'ɨ��ͼ��
''''''''''''''''''''''''''''''''''

Private Sub DelScanTmpDir(ByVal strDir As String)
'ɾ��ɨ����ʱ�ļ�
On Error GoTo errHandle
    If Dir(strDir, vbDirectory) <> "" Then
      Call mdlDir.DeleteFolder(strDir)
    End If
errHandle:
End Sub

Private Sub ScanImages()
'ɨ��ͼ��
On Error GoTo errProcess
                  
    'ɾ����������ʱ�洢��ͼ��Ŀ¼
    Call DelScanTmpDir(mstrTempDirOfScan)
        
    If Dir(mstrTempDirOfScan, vbDirectory) = "" Then
      Call MkDir(mstrTempDirOfScan)
    End If
    
    'ɾ��twain�豸��ʱ�洢��Ŀ¼
    Call DelScanTmpDir(gobjCapturePar.ScanDeviceTmpDir)
    
    If Dir(gobjCapturePar.ScanDeviceTmpDir, vbDirectory) = "" Then
      Call MkDir(gobjCapturePar.ScanDeviceTmpDir)
    End If
    
    mintScanImageIndex = 0
    
    '����ɨ�����ļ���������
    ImageScanner.FileType = BMP_Bitmap
    ImageScanner.StopScanBox = True
    ImageScanner.ShowSetupBeforeScan = True
    ImageScanner.ScanTo = UseFileTemplateOnly
    
    '���òɼ���ģ���ļ�
    ImageScanner.Image = mstrTempDirOfScan & "\" & CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE
    
    
    If Not ImageScanner.ScannerAvailable Then ImageScanner.OpenScanner
  
    Call ImageScanner.StartScan
    Call ImageScanner.StopScan
    Call ImageScanner.CloseScanner
    
    Exit Sub
errProcess:
    Call ImageScanner.CloseScanner

    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
    err.Clear
End Sub


Private Function IsVerityCapture() As Boolean
'�ж��Ƿ�Ϊ�����Ĳɼ���ʽ
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    IsVerityCapture = False
    
    '�ɼ�ͼ��ʱ��������Ǻ�̨�ɼ��������жϵ�ǰ���ص�ͼ�������ݿ��е�ͼ���¼���Ƿ�һ�£������һ�£�˵���ü�鵱ǰ�������������豸վ��ɼ�
    '�ô��޸���Ҫ�Ƿ�ֹ�豸������ʦ��Ƚ�̤������ɵ�ͼ��ɼ�
    strSQL = "select count(*) as ͼ���� from Ӱ����ͼ�� where ����uid in(select ����UID from Ӱ�������� where ���UID=(select ���UID from Ӱ�����¼ where ҽ��id=[1])) "
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯͼ������", mlngAdviceId)
    
    If rsData.RecordCount > 0 Then
        If Val(Nvl(rsData!ͼ����)) <> ucPreview.ImageTotal Then
            Call MsgboxCus("��ǰ���ص�ͼ��������ʵ�ʼ�¼����һ�£������Ƿ������û�������в��������޲�����ˢ�º����ԡ�", vbInformation + vbOKOnly, G_STR_HINT_TITLE)
            Exit Function
        End If
    End If
    
    IsVerityCapture = True

End Function


Private Sub CaptureImage()
'************************************************************
'
'����Ƶ����¼���вɼ�ͼ��
'
'************************************************************
    Dim blnIsRealCap As Boolean
    
    If Not (Not mblnReadOnly And (mVideoCapture.IsStartup Or IsTwainCaptureWay Or IsCustomCaptureWay)) Then Exit Sub  '���Ϊֻ����������Ƶû��������������ɼ�
    
    If Not IsVerityCapture Then Exit Sub
            
    If IsTwainCaptureWay Then
        Call ScanImages  'ͨ��TWAIN�ӿڲɼ�ͼ��
    ElseIf IsCustomCaptureWay Then
        Call CustomCapture
    Else
        blnIsRealCap = mblnRealTime 'Ϊʵʱ��ʾʱ�Զ���ʵʱͼ
        
        If Not mblnRealTime Then
            blnIsRealCap = IIf(MsgboxCus("ȷ��Ҫ�ɼ���ǰ��̬ͼ����ѡ���ǡ���ɼ���ǰ����ͼ��", vbQuestion + vbYesNo + vbDefaultButton1, G_STR_HINT_TITLE) = vbYes, False, True)
        End If
        
        '�ɼ�ͼ��
        Call subCaptureImg(blnIsRealCap)
    End If
End Sub

Private Sub CustomCapture()
    Dim objCapPic As StdPicture
    Dim strCapImgFiles As String
    Dim blnUseCustom As Boolean
    
    If mobjCustomDevice Is Nothing Then Exit Sub
    
    '�ɼ�ͼ��
    If Not mobjCustomDevice.zlCaptureImage(mlngAdviceId, objCapPic, strCapImgFiles, blnUseCustom) Then
        Exit Sub
    End If
    
    '����ɨ���ͼ��
    Call subCaptureImg(True, strCapImgFiles, objCapPic, False, blnUseCustom)
  
    Call ShowScanImage(ucPreview.CurImageCount)
End Sub

Private Sub CaptureAfterImage()
'��̨ͼ��ɼ�
    If Not mVideoCapture.IsStartup Then Exit Sub  '���Ϊֻ����������Ƶû��������������ɼ�,twain��ʽ�������̨�ɼ�
    
    Call subCaptureImg(True, "", Nothing, True)
End Sub


Public Sub zlExecuteCommandBars(control As XtremeCommandBars.CommandBarControl)
    On Error GoTo errHandle
        Call VideoCaptureMenuProcess(control)
        
        Call DicomImageMenuProcess(control)
        
        Call ControlImageMenuProcess(control)
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub DoStateChange(ByVal lngEventType As TVideoEventType, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String)
'����StateChange�¼�
On Error GoTo errHandle

BUGEX "DoStateChange(frmWork_Video) 1"
    RaiseEvent OnStateChange(lngEventType, lngAdviceId, lngSendNo, strOther)
    
BUGEX "DoStateChange(frmWork_Video) 2"
    '�㲥ͼ�������Ϣ
    If lngEventType = vetCaptureFirstImg _
        Or lngEventType = vetDelAllImg _
        Or lngEventType = vetUpdateImg Then
        
BUGEX "DoStateChange(frmWork_Video) 3 PostMessage lngAdviceId:" & lngAdviceId
        '���͹㲥��Ϣ
        BoradcastMsg lngAdviceId
    End If
    
BUGEX "DoStateChange(frmWork_Video) End"
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub VideoCaptureMenuProcess(control As XtremeCommandBars.CommandBarControl)
'��Ƶ�ɼ��˵�����
    Select Case control.id
        Case conMenu_Cap_Dynamic       '��̬��ʾ
            If IsTwainCaptureWay Then
                Call MsgboxCus("TWAIN�ɼ�ģʽ�£����ܽ��ж�̬��Ƶ����ʾ��", vbOKOnly, G_STR_HINT_TITLE)
            ElseIf IsCustomCaptureWay Then
                Call MsgboxCus("ר����Ƶ�ɼ�ģʽ�£����ܽ��ж�̬��Ƶ����ʾ��", vbOKOnly, G_STR_HINT_TITLE)
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
                MsgboxCus mstrMsg, vbOKOnly, "��ʾ"
                Exit Sub
            End If
            
            If IsCustomCaptureWay Then
                Call CustomVideoSave
            Else
                Call subVideoSave
            End If
            
        Case conMenu_Cap_Record_Stop  'ֹͣ¼�� modify by tjh at 2010-01-22
            If mstrVideoRegTime = "" Then
                'MsgboxCus  "δ��⵽��Ч��ע����Ϣ�����ܽ���¼�������", vbOKOnly, "��ʾ"
                Exit Sub
            End If
            
            Call subStopVideo
            
        Case conMenu_Cap_RecordAudio    '¼��
            If mstrVideoRegTime = "" Then
                MsgboxCus mstrMsg, vbOKOnly, "��ʾ"
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
                
        Case conMenu_Cap_OpenStudyList      '�򿪼��ɼ�ͼ��
            Call OpenStudy
            
        Case conMenu_Cap_StudySyncState     '�������
            If control.IconId = 10012 Then
                control.IconId = 8123
                
                Call LockStudy(0, 0, 0, 0, 0)
                
                Call DoStateChange(vetLockStudy, mlngAdviceId, mlngSendNo, mstrNameInf)
            Else
                control.IconId = 10012
                
                Call UnLockStudy
                
                If mlngAdviceId <> mSelectStudyInf.lngAdviceId Then
                    Call zlUpdateAdviceInf(mSelectStudyInf.lngAdviceId, mSelectStudyInf.lngSendNo, mSelectStudyInf.lngStudyState, mSelectStudyInf.blnMoved)
                    Call zlRefreshFace
                End If
                
                Call DoStateChange(vetUnLockStudy, mlngAdviceId, mlngSendNo, mstrNameInf)
            End If
        Case conMenu_Cap_After_Tag      '���º�̨�ɼ���ʶ
            
            If mstrVideoRegTime = "" Then
                MsgboxCus mstrMsg, vbOKOnly, "��ʾ"
                Exit Sub
            End If
            
            Call UpdateAfterCaptureInfo
            Call ShowAfterCaptureInf(True)
            
    End Select
End Sub

Private Sub DicomImageMenuProcess(control As XtremeCommandBars.CommandBarControl)
'dicomͼ��˵�����
    If mblnRealTime = True Or dcmView.Images.Count <= 0 Then Exit Sub
    
    Select Case control.id
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
            Call subSetRotate(dcmView.Images(1), True)
            
        Case conMenu_Process_LRotate        '��ʱ����ת
            Call subSetRotate(dcmView.Images(1), False)
            
        Case conMenu_Process_Sharpness      '��
            Call subSetSharp(dcmView.Images(1), True)
            
        Case conMenu_Process_Filter         'ƽ��
            Call subSetSharp(dcmView.Images(1), False)
            
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
    End Select

End Sub

Private Sub ControlImageMenuProcess(control As XtremeCommandBars.CommandBarControl)
    Select Case control.id
        Case conMenu_Cap_SplitPage
            Call ucPreview.ShowPageConfig
            Call ucCacheImage.ShowPageConfig
            
        Case conMenu_Cap_ImgImport          '���͵����
            mlngReleationType = 2
            If ReleationImage Then
                mAfterCaptureInf.lngAfterCurImageCount = GetCountOfCapPic(mintTagNow)
                Call ShowAfterCaptureInf(True)
                
                Call ClearEmptyFolder(False)
            End If
            
        Case conMenu_Cap_ImgExport          '���͵���̨
            mlngReleationType = 1
            Call ReleationImage
            
        Case conMenu_Cap_AddToReport        '���뱨��ͼ
            Call AddImageToReport
            
        Case conMenu_Cap_ReUpLoad           '�����ϴ�
            Call ucPreview_OnReUpload
            
        Case conMenu_Cap_ReDownLoad         '��������
            Call ucPreview.ReLoadFailedImage
            
        Case conMenu_Cap_DelCache
            mlngReleationType = 2
            Call DelTempImage
            Call ClearEmptyFolder(False)

            mAfterCaptureInf.lngAfterCurImageCount = GetCountOfCapPic(mintTagNow)
            Call ShowAfterCaptureInf(True)
            
            
        Case conMenu_Cap_RefreshCache
            Call LoadMiniCache
            
    End Select
End Sub

Private Function AddImageToReport() As Boolean
'------------------------------------------------
'���ܣ���ͼ����ʱ��ӵ������У����ñ���ͼ���
'������ lngActionType �� 1 - ���뱨��ͼ
'���أ�True - �ɹ��� False - ʧ��
'------------------------------------------------
    On Error GoTo err
    
    Dim i As Long
    Dim strSQL As String
    
    '��ȡͼ��UID
    For i = 1 To ucPreview.CurImageCount
        '�޸ı���ͼ��tag�����޸ĳ����еĻ�������
        '���뱨��ͼʱ�����>1�������޸ģ�=0�����޸ģ�=""�ĳ�=0
        If ucPreview.ImgChecked(i) = True Then
            If ucPreview.ImgViewer.Images(i).Tag.ReportImage = "" Then
                ucPreview.ImgViewer.Images(i).Tag.ReportImage = 0
                Call AddToReport(ucPreview.ImgViewer.Images(i))
            End If
            AddImageToReport = True
        End If
    Next i
    
    If AddImageToReport = False Then
        MsgBox "����ѡ��һ��ͼ��", vbOKOnly, "��ʾ��Ϣ"
        Exit Function
    End If
    
    Exit Function
err:
    err.Raise err.Number, err.Description
End Function

Private Sub AddToReport(dcmImage As DicomImage)
    Dim strSQL As String
    '�������ݿ�
    strSQL = "Zl_Ӱ����_���ñ���ͼ('" & dcmImage.InstanceUID & "',1)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    '���¼����±���ͼ��ʾ
    Call DoStateChange(vetAddReportImg, mlngAdviceId, mlngSendNo, dcmImage.InstanceUID)
    '���ɼ�����ͼ�ı���ͼ���
    Call ucPreview.DrawReportImgTag(dcmImage)
End Sub

Private Function ReleationImage() As Boolean
    Dim strHint As String
    Dim rsImageDatas As ADODB.Recordset
    Dim strTmpFile As String
    Dim i As Integer
    
    Set rsImageDatas = GetReleationImageIds()
    
    If rsImageDatas Is Nothing Then
        Call MsgboxEx(Me.hwnd, "��ѡ����Ҫ���й����ļ��ͼ��", vbInformation, G_STR_HINT_TITLE)
        Exit Function
    End If
        
    '��ǰ���UID�����ݿ��в����ڣ����˳�������
    If rsImageDatas.RecordCount <= 0 Then
        Call MsgboxEx(Me.hwnd, "��ѡ����Ҫ���й����ļ��ͼ��", vbInformation, G_STR_HINT_TITLE)
        Exit Function
    End If
    
    If mlngReleationType = 2 Then
        '����ͼ����ʾ
        strHint = GetReleationHintInfo(mlngAdviceId, rsImageDatas)
        
        If strHint = "" Then
            Call MsgboxEx(Me.hwnd, "���ܲ�ѯ����Ҫ������������Ϣ������������", vbOKOnly, G_STR_HINT_TITLE)
            Exit Function
        End If
        
        If MsgboxEx(Me.hwnd, strHint, vbDefaultButton2 + vbQuestion + vbYesNo, G_STR_HINT_TITLE) = vbNo Then Exit Function
        
    Else
        'ȡ��������ʾ
        If MsgboxEx(Me.hwnd, "�Ƿ�ȷ�Ͻ���ѡͼ���͵���̨��", vbDefaultButton2 + vbQuestion + vbYesNo, G_STR_HINT_TITLE) = vbNo Then Exit Function
    End If

    If mlngReleationType = 2 Then '����2��ʾ����ͼ��
        ReleationImage = StartReleation(mlngAdviceId, rsImageDatas)
    Else
        ReleationImage = CancelReleation(mlngAdviceId, rsImageDatas)
    End If
        
    '�����򣬷�ֹ����2������BUG
    For i = 1 To ucPreview.CurImageCount
        ucPreview.ImgViewer.Images(i).BorderColour = vbWhite
    Next
    
    Call DoStateChange(IIf(mlngReleationType = 2, vetImportImage, vetExportImage), mlngAdviceId, mlngSendNo, mcurStudyInf.strStudyUid)
End Function

'ȡ�ù�����ʾ��Ϣ
Private Function GetReleationHintInfo(lngAdviceId As Long, rsReleationImage As ADODB.Recordset) As String
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    GetReleationHintInfo = ""
    
    strSQL = "select ����,����,�Ա�,���� from Ӱ�����¼ where ҽ��ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceId)

    If rsTemp.RecordCount <= 0 Then Exit Function
    
    GetReleationHintInfo = "�Ƿ�ȷ�Ͻ�ѡ���ͼ���͵�[" & Nvl(rsTemp!����) & "(" & Nvl(rsTemp!����) & ") " & Nvl(rsTemp!�Ա�) & " " & Nvl(rsTemp!����) & "]�ļ���У�"
End Function

Private Function StartReleation(ByVal lngAdviceId As Long, rsImageDatas As ADODB.Recordset) As Boolean
'��ʼ����
On Error GoTo errHandle
    Dim strSQL As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim objFtp As New clsFtp
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    Dim arrSql() As String
    Dim i As Long
    
    blnBeginTrans = False
    StartReleation = False
    
    curDate = zlDatabase.Currentdate
    
    strSQL = "select ���UID,�������� from Ӱ�����¼ where ҽ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceId)
    
    If rsTmp.RecordCount <= 0 Then
        Call MsgboxEx(Me.hwnd, "�Ҳ����������ļ����Ϣ��", vbInformation, G_STR_HINT_TITLE)
        Exit Function
    End If
    
    If Trim(Nvl(rsTmp!���uid)) = "" Or Trim(Nvl(rsTmp!��������)) = "" Then

        '��δ�ɼ�ͼ����Ҫ�����µļ��UID
        strNewStudyUID = CreateStudyUid(rsImageDatas!���uid)
        
        Call GetStorageDevice(mlngAdviceId, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgboxEx(Me.hwnd, "����ȡ����Ч�Ĵ洢�豸������洢�豸���á�", vbInformation, G_STR_HINT_TITLE)
            Exit Function
        End If
        
        '���´洢�豸��Ϣ
        strSQL = "Zl_Ӱ����_�����豸(" & mlngAdviceId & ",'" & strNewStudyUID & "','" & strNewDeviceNo & "'," & _
                                        "to_Date('" & Format(curDate, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Else
        strNewStudyUID = Nvl(rsTmp!���uid)
        
        Call GetStorageDevice(mlngAdviceId, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgboxEx(Me.hwnd, "����ȡ����Ч�Ĵ洢�豸������洢�豸���á�", vbInformation, G_STR_HINT_TITLE)
            Exit Function
        End If
    End If
    
    '����FTP
    If objFtp.FuncFtpConnect(strNewFtpIp, strNewFtpUser, strNewFtpPwd) = 0 Then
        Call MsgboxEx(Me.hwnd, "FTP����ʧ�ܣ������������á�", vbInformation, G_STR_HINT_TITLE)
        Exit Function
    End If

    ReDim arrSql(0)
    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        '�����µ�����UID
        strNewSeriesUid = CreateSeriesUid(rsImageDatas!����Uid, strNewStudyUID)
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = "Zl_Ӱ����_ͼ����(" & mlngAdviceId & ",'" & strNewStudyUID & "','" & strNewSeriesUid & "','" & mobjFile.GetFileName(Nvl(rsImageDatas!·��)) & "')"
        '�޸ı����ļ�
        Call UpdateAfterImage(Nvl(rsImageDatas!·��), strNewStudyUID, strNewSeriesUid)
        rsImageDatas.MoveNext
    Wend
    
    '�ƶ�ͼ���ļ�
    If Not MoveImageToStudy(objFtp, rsImageDatas, strNewFtpVirtualPath, objMoveList) Then Exit Function
    
    gcnVideoOracle.BeginTrans
    
    blnBeginTrans = True
    
    For i = 1 To UBound(arrSql)
        '����ͼ���������
        Call zlDatabase.ExecuteProcedure(arrSql(i), Me.Caption)
    Next
    
    '�ύ����
    Call gcnVideoOracle.CommitTrans
    
    '˵��ȫ���ϴ��ɹ�,ɾ��������ʱͼ��
    Call DelTempImages(rsImageDatas)
    
    StartReleation = True
    
    Exit Function
errHandle:
    If blnBeginTrans Then Call gcnVideoOracle.RollbackTrans
    
    Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
    Call OutputDebug("StartReleation", err)
    
    Call RaiseErr(err)  '�����׳�����
End Function

Private Sub UpdateAfterImage(strTmpFile As String, strStudyUid As String, strSeriesUID As String)
'�޸ı����ļ�
    Dim curImage As DicomImage
    Dim dcmViewer As DicomViewer
    
    If Len(Dir(strTmpFile)) = 0 Then Exit Sub
    
    Set curImage = ReadViewImage(strTmpFile, dcmViewer)
    curImage.StudyUID = strStudyUid
    curImage.SeriesUID = strSeriesUID
    curImage.WriteFile strTmpFile, True
End Sub

Private Function CancelReleation(ByVal lngAdviceId As Long, rsImageDatas As ADODB.Recordset) As Boolean
'��������
On Error GoTo errHandle
    Dim strSQL As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim objFtp As New clsFtp
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim rsReportImage As ADODB.Recordset
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    
    blnBeginTrans = False
    CancelReleation = False
    
    curDate = zlDatabase.Currentdate
    
    '����ͼ�����
    strNewStudyUID = CreateStudyUid(mdcmUID.NewUID)
        
    Call GetStorageDevice(mlngAdviceId, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
    If Trim(strNewFtpIp) = "" Then
        Call MsgboxEx(Me.hwnd, "����ȡ����Ч�Ĵ洢�豸������洢�豸���á�", vbInformation, G_STR_HINT_TITLE)
        Exit Function
    End If
    
    '����FTP
    If objFtp.FuncFtpConnect(strNewFtpIp, strNewFtpUser, strNewFtpPwd) = 0 Then
        Call MsgboxEx(Me.hwnd, "FTP����ʧ�ܣ������������á�", vbInformation, G_STR_HINT_TITLE)
        Exit Function
    End If
    
    If Not MoveImageToAfter(objFtp, rsImageDatas, objMoveList) Then Exit Function

    gcnVideoOracle.BeginTrans
    
    blnBeginTrans = True
    
    '��������
    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        strSQL = "Zl_Ӱ����_ͼ�񵼳�(" & mlngAdviceId & ",'" & Nvl(rsImageDatas!ͼ��UID) & "')"
         
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        rsImageDatas.MoveNext
    Wend
    
    Call gcnVideoOracle.CommitTrans
    
    Call ClearFtpImage(rsImageDatas, strNewStudyUID)
    
    CancelReleation = True
Exit Function
errHandle:
    If blnBeginTrans Then Call gcnVideoOracle.RollbackTrans
    Call OutputDebug("CancelReleation", err)
    Call RaiseErr(err)
End Function

'����ͼ����ƶ�
Private Sub CancelImageMove(ByVal strFTPIP As String, ByVal strFTPUser As String, ByVal strFTPPwd As String, objMoveList As Collection)
    Dim i As Long
    Dim objFtp As New clsFtp
    Dim strDestFile As String
    Dim strMoveFile As String
    
    If objMoveList Is Nothing Then Exit Sub
    If objMoveList.Count <= 0 Then Exit Sub
    
On Error GoTo errHandle

    Call objFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
    
    For i = 1 To objMoveList.Count
        strDestFile = objMoveList.Item(i)
        
        strMoveFile = Mid(strDestFile, InStr(strDestFile, ">") + 1, 255)
        strDestFile = Mid(strDestFile, 1, InStr(strDestFile, ">") - 1)
        
        Call objFtp.FuncReNameFile(strMoveFile, strDestFile)
    Next i
        
errHandle:
    objFtp.FuncFtpDisConnect
End Sub

Private Function CreateSeriesUid(ByVal strSeriesUID As String, ByVal strStudyUid As String) As String
'��������UID
    Dim rsData As New ADODB.Recordset
    Dim strSQL As String
    Dim strNewSeriesUid As String
    
    strNewSeriesUid = strSeriesUID 'M_STR_SERIES_UID & "." & Format(Now, "yymmddhhmmss") & "." & Fix(Rnd(1000) * 1000)
    
    strSQL = "select ����UID from Ӱ�������� where ����UID = [1] And ���UID <> [2]"
              
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "PACSͼ�񱣴�", strNewSeriesUid, strStudyUid)
    
    If rsData.RecordCount > 0 Then
        '����һ���µļ��UID
        strSQL = "Select Ӱ����UID���_ID.Nextval From Dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "PACSͼ�񱣴�")
        
        If Len(strNewSeriesUid) <= 55 Then
            strNewSeriesUid = strNewSeriesUid & ".A" & rsData(0)
        Else
            strNewSeriesUid = Left(strNewSeriesUid, 55) & ".A" & rsData(0)
        End If
    End If
    
    CreateSeriesUid = strNewSeriesUid
End Function

Private Sub ClearFtpImage(rsImageDatas As ADODB.Recordset, ByVal strNewStudyUID As String)
On Error GoTo errHandle
'ת��ͼ��ɹ�����ɾ����ʱͼ���ԭ��FTP��ͼ���Ŀ¼���峡�������ִ�����Բ�����
    Dim objSrcFtp As New clsFtp
    Dim strTmpFile As String
    Dim strVirtualPath As String
    Dim strImageUID As String
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        strTmpFile = GetAppPath & "\TmpImage\" & Nvl(rsImageDatas!ͼ��UID)
        
        strImageUID = Nvl(rsImageDatas!ͼ��UID)
        
        strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
                
        If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
            strCurFtpIp = Nvl(rsImageDatas!host)
            strCurFtpUser = Nvl(rsImageDatas!FtpUser)
            strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
            
            Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
        End If
        
'       Ϊ������������ͼ��������ش���ͼ���ļ������ý���ɾ��
        
        If FileExists(strTmpFile) Then Call Kill(strTmpFile)
        If FileExists(strTmpFile & ".jpg") Then Call Kill(strTmpFile & ".jpg")
        
        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID)
        
        'ɾ���յ�ftpĿ¼
        Call objSrcFtp.FuncFtpDelDir(Replace(strVirtualPath, strImageUID, ""), strImageUID)
                
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend
    
    objSrcFtp.FuncFtpDisConnect
Exit Sub
errHandle:
    Call OutputDebug("ClearFtpImage", err)
End Sub

Private Function MoveImageToStudy(objFtp As clsFtp, rsImageDatas As ADODB.Recordset, strNewFtpVirtualPath As String, ByRef objMoveList As Collection) As Boolean
    Dim i As Long

    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        If objFtp.FuncUploadFile(strNewFtpVirtualPath, rsImageDatas!·��, mobjFile.GetFileName(rsImageDatas!·��)) <> 0 Then
            'ʧ�ܺ�ɾ��֮ǰ�ϴ����ļ�
            For i = 0 To objMoveList.Count - 1
                Call objFtp.FuncDelFile(strNewFtpVirtualPath, objMoveList(i))
            Next
            
            Call MsgboxEx(Me.hwnd, "ͼ���ƶ�ʧ�ܣ�����FTP�����Ƿ�������", vbInformation, G_STR_HINT_TITLE)
            
            Exit Function
        Else
            Call objMoveList.Add(rsImageDatas!·��)
        End If
        
        rsImageDatas.MoveNext
    Wend
    
    MoveImageToStudy = True
End Function

Private Function MoveImageToAfter(objFtp As clsFtp, rsImageDatas As ADODB.Recordset, ByRef objMoveList As Collection) As Boolean
    Dim i As Long
    Dim strDestPath As String
    
    rsImageDatas.MoveFirst

    While Not rsImageDatas.EOF
        strDestPath = GetAfterImagePath(rsImageDatas!ͼ��UID, rsImageDatas!����Uid, rsImageDatas!���uid)
        If mobjFile.FolderExists(strDestPath) = False Then Call MkLocalDir(strDestPath)
        
        If objFtp.FuncDownloadFile(rsImageDatas!Root & rsImageDatas!Url, strDestPath & rsImageDatas!ͼ��UID, rsImageDatas!ͼ��UID) <> 0 Then
            'ʧ�ܺ�ɾ��֮ǰ���ص��ļ�
            For i = 0 To objMoveList.Count - 1
                Call mobjFile.DeleteFile(objMoveList(i))
            Next

            Call MsgboxEx(Me.hwnd, "ͼ���ƶ�ʧ�ܣ�����FTP�����Ƿ�������", vbInformation, G_STR_HINT_TITLE)
            
            Exit Function
        Else
            Call objMoveList.Add(strDestPath & rsImageDatas!ͼ��UID)
        End If
        
        rsImageDatas.MoveNext
    Wend
    Call MsgboxEx(Me.hwnd, "�ѽ�ѡ��ͼ���͵�[���" & mintTagNow & "]��", vbInformation, G_STR_HINT_TITLE)
    
    MoveImageToAfter = True
End Function

Public Function GetAfterImagePath(ByVal strImageName As String, ByVal strSeriesUID As String, ByVal strStudyUid As String) As String
'���͵���̨���Ŀ¼ ���Ѿ����ڸñ�ʶ��Ŀ¼ ʹ��֮�����򴴽�
    Dim strTmpPath As String
    Dim objFolder1 As Folder, objFolder2 As Folder, objFolder3 As Folder
    Dim curDate As Date
    Dim strDate As String
    Dim intTmp As Integer
    
    curDate = zlDatabase.Currentdate
    strDate = Format(curDate, "yyyymmdd")
    
    strTmpPath = ""
    
    If mobjFile.FolderExists(mstrAfterDir & "\") Then
        For Each objFolder1 In mobjFile.GetFolder(mstrAfterDir & "\").SubFolders   'ʱ���
            If objFolder1.Name = strDate Then 'ʱ��ֻ�жϵ���

                For Each objFolder2 In mobjFile.GetFolder(objFolder1.Path).SubFolders   '����
                
                    If InStr(objFolder2.Name, "���" & mintTagNow) > 0 Then '�ж��Ƿ���������+��ǰ��ʶ��Ŀ¼�����У�ֱ��ʹ�ã�
                        
                        For Each objFolder3 In mobjFile.GetFolder(objFolder2.Path).SubFolders   '���в�
                                strTmpPath = objFolder3.Path & "\"
                                GoTo step2
                        Next
                   
                    End If
                Next
                
                Exit For '��ֹʱ����ļ��е�����
            End If
        Next
    End If
    

    If strTmpPath = "" Then
        mAfterCaptureInf.strAfterStudyUid = ""
        mintTagNow = mintTagNow + 1
        mintTagMax = mintTagMax + 1
        strTmpPath = mstrAfterDir & "\" & Format(curDate, "yyyymmdd") & "\" & "���" & mintTagNow & "-" & strStudyUid & "\" & strSeriesUID & "\"
    End If
    
    '�ҵ�Ŀ¼��ֹͣǰ��ı�����ֱ�ӽ���step2
step2:
    
    intTmp = mintTagNow + 1
    
    Call ShowAfterCaptureInf(True)
    
    GetAfterImagePath = strTmpPath
End Function


Private Sub GetStorageDevice(ByVal lngAdviceId As Long, ByVal strNewStudyUID As String, _
    ByRef strDeviceNO As String, ByRef strFTPIP As String, _
    ByRef strFtpUrl As String, ByRef strVirtualPath As String, _
    ByRef strFTPUser As String, ByRef strFTPPwd As String)
'��ȡ�µĴ洢�豸��Ϣ������豸�洢��Ϣ�����ڣ�����Ҫ��������
'�����ȡ����������ʹ��strNewStudyUID�����ܴ����ݿ��в��ҵ���Ӧ������
'strDeviceNum:�豸��
'strFtpIp: ftp��ַ
'strFtpUrl: ftpĿ¼
'strVirtualPath: ftp����洢·��
'strFtpUser: ftp�û���
'strFtpPwd: ftp����

    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnIsGetNewDevice As Boolean
    Dim objDestFtp As New clsFtp
    Dim curDate As Date
    
    strFTPIP = ""
    strFtpUrl = ""
    strFTPUser = ""
    strFTPPwd = ""
    strDeviceNO = ""
    
    strSQL = "Select D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd,C.λ��һ,C.λ�ö�,C.λ����,C.��������," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And C.���UID= [1]"
        
    If mblnMoved Then
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    End If
        
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNewStudyUID)
    
    blnIsGetNewDevice = False
    
    If rsData.RecordCount <= 0 Then
        blnIsGetNewDevice = True
    Else
        '���ִ�е����˵����ִ��ͼ�����,��Ҫ�жϵ�ǰ���Ĵ洢�豸�Ƿ���Ч�������Ч�������µĴ洢�豸
        If Trim(rsData!��������) = "" Then
            blnIsGetNewDevice = True
        Else
            strDeviceNO = Nvl(rsData!λ��һ)
            strFTPIP = Nvl(rsData!host)
            strFtpUrl = Nvl(rsData!Root)
            strFTPUser = Nvl(rsData!FtpUser)
            strFTPPwd = Nvl(rsData!FtpPwd)
            strVirtualPath = strFtpUrl & Nvl(rsData!Url)
        End If
    End If
    
    If blnIsGetNewDevice Then
        '�����µļ��UID�ʹ洢�豸,���ִ�е����˵����ȡ������
        
        '��ѯ��ҽ������վ�е�ͼ��洢�豸
        strDeviceNO = GetDeptPara(glngDepartId, "�洢�豸��")
        
        If Val(strDeviceNO) <= 0 Then
            MsgboxEx Me.hwnd, "δ�ҵ�ͼ��洢�豸,��ȷ����Ӱ�����̹������Ƿ�Ըÿ���������ͼ��ɼ��洢�豸��", vbInformation, G_STR_HINT_TITLE
            Exit Sub
        End If
        
        strSQL = "Select �豸��,�豸��,'/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL,FTP�û���,FTP����,IP��ַ " & _
                    " From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
                    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Tag, strDeviceNO)
        
        '����洢�豸ͣ�ã���ֱ���˳�
        If rsTemp.RecordCount <= 0 Then
            MsgboxEx Me.hwnd, "δ�ҵ��洢�豸,��ȷ���豸��Ϊ [" & strDeviceNO & "] ���豸�Ƿ����á�", vbInformation, G_STR_HINT_TITLE
            Exit Sub
        End If
        
        strFtpUrl = Nvl(rsTemp("URL"))
        strFTPIP = Nvl(rsTemp("IP��ַ"))
        strFTPUser = Nvl(rsTemp("FTP�û���"))
        strFTPPwd = Nvl(rsTemp("FTP����"))
        
        strFtpUrl = IIf(strFtpUrl = "/", "//", strFtpUrl)
        
        objDestFtp.FuncFtpConnect strFTPIP, strFTPUser, strFTPPwd
        On Error GoTo errHandle
            curDate = zlDatabase.Currentdate
            
            strVirtualPath = strFtpUrl & Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            '����FTPĿ¼
            objDestFtp.FuncFtpMkDir strFtpUrl, Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            
        Call objDestFtp.FuncFtpDisConnect
errHandle:
        Call objDestFtp.FuncFtpDisConnect
    End If
End Sub


Private Sub DelTempImage()
    Dim rsImageDatas As ADODB.Recordset
    Dim i As Long
    
    '�����ݿ��в�ѯͼ������
    Set rsImageDatas = GetReleationImageIds()
    
    If rsImageDatas Is Nothing Then
        Call MsgboxEx(Me.hwnd, "��ѡ����Ҫɾ���ļ��ͼ��", vbInformation, G_STR_HINT_TITLE)
        Exit Sub
    End If
    
    '��ǰ���UID�����ݿ��в����ڣ����˳�������
    If rsImageDatas.RecordCount <= 0 Then
        Call MsgboxEx(Me.hwnd, "��ѡ����Ҫɾ���ļ��ͼ��", vbInformation, G_STR_HINT_TITLE)
        Exit Sub
    End If
    
    If MsgboxEx(Me.hwnd, "�Ƿ�ȷ��ɾ����ѡͼ��", vbDefaultButton2 + vbQuestion + vbYesNo, G_STR_HINT_TITLE) = vbNo Then Exit Sub
    
    Call DelTempImages(rsImageDatas)
End Sub

Private Function DelTempImages(rsImageDatas As ADODB.Recordset) As Boolean
'ɾ�����ػ����е��ļ����ڽ�����ɾ��ucpre�ؼ���ѡ��ͼ��
    Dim blfinished As Boolean
    Dim i As Long
    Dim curTime As Date
    Dim intTmp As Integer
    
On Error GoTo errHandle
    If rsImageDatas.RecordCount <= 0 Then Exit Function
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        If mobjFile.FileExists(rsImageDatas!·��) Then Call mobjFile.DeleteFile(rsImageDatas!·��)
        
        rsImageDatas.MoveNext
    Wend
    
   
    'ɾ������ͼ��
    blfinished = False
    
    For i = ucCacheImage.CurImageCount To 1 Step -1
        If ucCacheImage.ImgChecked(i) Then
            Call ucCacheImage.DeleteImage(i)
            blfinished = True
        End If
    Next
    

    
    If blfinished = False Then
        Call ucCacheImage.DeleteImage(ucCacheImage.SelectIndex)
    End If
    
    'ͬʱ��Ҫɾ��cbo��Ŀ
    
    If ucCacheImage.CurImageCount = 0 Then
        curTime = zlDatabase.Currentdate
        '�ǵ��첢��ѡ�е��ǵ�ǰ��ʶ���Ͳ�������ղ���
        If Not ((Format(DTPImg.value, "yyyymmdd") = Format(curTime, "yyyymmdd")) And (InStr(cboCacheImage.Text, "��ʶ" & Lpad((mintTagNow), 3, "0")) > 0)) Then
            intTmp = cboCacheImage.ListIndex
            cboCacheImage.RemoveItem (cboCacheImage.ListIndex)
            If cboCacheImage.ListCount > intTmp - 1 Then
                cboCacheImage.ListIndex = intTmp - 1
            Else
                cboCacheImage.ListIndex = 0
            End If
        End If
    End If

    DelTempImages = True
    
    Exit Function
errHandle:
    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Function

Private Function GetReleationImageIds() As ADODB.Recordset
'��ѯ��������Ҫȡ��������ͼ��ID
    Dim i As Long, j As Long
    Dim strSQL As String
    Dim strValues(0 To 80) As String
    Dim strValue As String
    Dim strUninTable As String
    Dim strFilter As String
    Dim strTmpFile As String
    Dim rsImageDatas As ADODB.Recordset

    j = 0
    strUninTable = ""
    strFilter = ""
    strValue = ""
    
    '�����ѯ���
    If mlngReleationType = 1 Then
        For i = 1 To ucPreview.CurImageCount
            If ucPreview.ImgChecked(i) Then
                If j > 79 Then
                    strFilter = strFilter & " Or ͼ��UID ='" & ucPreview.ImgViewer.Images(i).InstanceUID & "'"
                Else
                    If zlCommFun.ActualLen(strValue) > 3600 Then
                         strValues(j) = Mid(strValue, 2)
                         strUninTable = strUninTable & " Union ALL  Select  Column_Value as ͼ��UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
                         
                         strValue = ""
                         j = j + 1
                    End If
                    
                    strValue = strValue & "," & ucPreview.ImgViewer.Images(i).InstanceUID
                End If
            End If
        Next
                
        '������ͼ��û�б�ѡ�еĺ�㣬���к���ͼ����Ϊѡ��
        If Not ucPreview.SelectImage Is Nothing And strValue = "" Then strValue = strValue & "," & ucPreview.SelectImage.InstanceUID
                
    Else
        Set rsImageDatas = New ADODB.Recordset
        rsImageDatas.Fields.Append "����UID", adVarChar, 4000
        rsImageDatas.Fields.Append "���UID", adVarChar, 4000
        rsImageDatas.Fields.Append "·��", adVarChar, 4000
        rsImageDatas.Open
            
        For i = 1 To ucCacheImage.CurImageCount
            If ucCacheImage.ImgChecked(i) Then
                strTmpFile = ucCacheImage.ImgViewer.Images(i).Tag.FilePath
                rsImageDatas.AddNew
                rsImageDatas!����Uid = mobjFile.GetFolder(mobjFile.GetParentFolderName(strTmpFile)).Name
                rsImageDatas!���uid = GetStudyUIDFromFolderName(mobjFile.GetFolder(mobjFile.GetParentFolderName(mobjFile.GetParentFolderName(strTmpFile))).Name)
                'rsImageDatas!���uid = mobjFile.GetFolder(mobjFile.GetParentFolderName(mobjFile.GetParentFolderName(strTmpFile))).Name
                rsImageDatas!·�� = strTmpFile
                rsImageDatas.Update
            End If
        Next
        
        'û��ͼ���ں��״̬����ѡ���к���
        If rsImageDatas.RecordCount = 0 Then
            If ucCacheImage.CurImageCount > 0 Then
                If Not ucCacheImage.SelectImage Is Nothing Then
                    strTmpFile = ucCacheImage.SelectImage.Tag.FilePath
                    rsImageDatas.AddNew
                    rsImageDatas!����Uid = mobjFile.GetFolder(mobjFile.GetParentFolderName(strTmpFile)).Name
                    rsImageDatas!���uid = GetStudyUIDFromFolderName(mobjFile.GetFolder(mobjFile.GetParentFolderName(mobjFile.GetParentFolderName(strTmpFile))).Name)
                    rsImageDatas!·�� = strTmpFile
                    rsImageDatas.Update
                End If
            End If
        End If
                
        If rsImageDatas.RecordCount > 0 Then rsImageDatas.MoveFirst
        
        Set GetReleationImageIds = rsImageDatas
        Exit Function
    End If
    
    If strValue <> "" Then
        strValues(j) = Mid(strValue, 2)
        strUninTable = strUninTable & " Union ALL  Select  Column_Value as ͼ��UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
    End If
    
    '���û����Ҫ���ҵ�ͼ��UID���򷵻ؿ����ݼ�
    If strUninTable <> "" Then
        strUninTable = Mid(strUninTable, 11)
    Else
        Set GetReleationImageIds = Nothing
        Exit Function
    End If
    
    If strFilter <> "" Then strUninTable = strUninTable & " Union All Select ͼ��UID from [Ӱ��ͼ��] where  ( " & Mid(strFilter, 4) & ")"
    
    strSQL = "Select /*+ RULE*/ D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd, Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ) as �豸��," & _
        "D.IP��ַ As Host,B.����UID,B.���UID,C.Ӱ�����, " & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL,A.ͼ��UID, c.����,c.�Ա�,c.����,c.���� " & _
        "From Ӱ����ͼ�� A, Ӱ�������� B, Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,(" & Replace(strUninTable, "[Ӱ��ͼ��]", "Ӱ����ͼ��") & ") E " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And A.����UID=B.����UID and B.���UID=C.���UID and A.ͼ��UID = E.ͼ��UID "
        
    If mblnMoved Then
        strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    End If
    
    Set GetReleationImageIds = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValues(0), strValues(1), strValues(2), strValues(3), _
        strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10), _
        strValues(11), strValues(12), strValues(13), strValues(14), strValues(15), strValues(16), strValues(17), _
        strValues(18), strValues(19), strValues(20), strValues(21), strValues(22), strValues(23), strValues(24), strValues(25), strValues(26), _
        strValues(27), strValues(28), strValues(29), strValues(30), strValues(31), strValues(32), strValues(33), strValues(34), strValues(35), strValues(36), _
        strValues(37), strValues(38), strValues(39), strValues(40), strValues(41), strValues(42), strValues(43), strValues(44), strValues(45), strValues(46), _
        strValues(47), strValues(48), strValues(49), strValues(50), strValues(51), strValues(52), strValues(53), strValues(54), strValues(55), strValues(56), _
        strValues(57), strValues(58), strValues(59), strValues(60), strValues(61), strValues(62), strValues(63), strValues(64), strValues(65), strValues(66), _
        strValues(67), strValues(68), strValues(69), strValues(70), strValues(71), strValues(72), strValues(73), strValues(74), strValues(75), strValues(76), _
        strValues(77), strValues(78), strValues(79), strValues(80))
End Function

Private Sub OpenStudy()
    Dim cbrControl As CommandBarControl
    
    Dim lngCurAdviceId As Long
    Dim lngSendNo As Long
    Dim lngStudyState As Long
    Dim blnResult As Boolean
    
    
    If mfrmOpenStudy Is Nothing Then Set mfrmOpenStudy = New frmOpenStudyList
    
    blnResult = mfrmOpenStudy.ShowStudyWindow(lngCurAdviceId, lngSendNo, lngStudyState, Me)
    
    If blnResult = False Then Exit Sub
        
    If lngCurAdviceId > 0 Then
        '��ʼ���µļ����вɼ�
        Call UnLockStudy
        
        Call zlUpdateAdviceInf(lngCurAdviceId, lngSendNo, lngStudyState, 0)
        Call zlRefreshFace
        
        Call LockStudy(0, 0, 0, 0, 0)
                
        '�޸�����״̬
        Set cbrControl = cbrMain.FindControl(, conMenu_Cap_StudySyncState)
        cbrControl.IconId = 8123
        
        '�������˸ı��¼�
        Call DoStateChange(vetLockStudy, mlngAdviceId, mlngSendNo, mstrNameInf)
    End If
    
End Sub


Public Sub zlUnloadMe()
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
            cbrMain.Item(2).position = xtpBarTop
            cbrMain.Item(3).position = xtpBarBottom
        Else
            cbrMain.Item(2).position = xtpBarLeft
            cbrMain.Item(3).position = xtpBarRight
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
        Call subCenterZoom(Me, dcmView.Images(1), dcmView, dcmView.Images(1).ActualZoom, mCorpSize)
    End If

    'ˢ�²ü�����λ��
    Call RefreshPbxSizePos
        
    
    If IsTwainCaptureWay Or IsCustomCaptureWay Then
      
        '����ͼ�����ʾ��Χ
        dcmView.Left = Left
        dcmView.Top = Top
        dcmView.Width = Right - Left
        dcmView.Height = Bottom - Top
  
        'ˢ��DcmView��ͼ�����ʾλ��
        If dcmView.Images.Count > 0 Then
            Call subCenterZoom(Me, dcmView.Images(1), dcmView, dcmView.Images(1).ActualZoom, mCorpSize)
        End If
    
    End If
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    zlUpdateCommandBars control
End Sub


Private Sub commListener_OnComm()
On Error GoTo errHandle
    Dim strInput As String
    
    '�����TWAINɨ���ר����Ƶ�ɼ�����֧�ֽ�̤���زɼ�
    If IsTwainCaptureWay Or IsCustomCaptureWay Then Exit Sub
    
    If gobjCapturePar.ComPortType <> "COM" Then Exit Sub
    
    strInput = ""
    If commListener.InBufferCount > 0 Then strInput = commListener.Input
    
    If Not (commListener.CommEvent = comEvCTS Or commListener.CommEvent = comEvDSR _
        Or commListener.CommEvent = comEvCD Or commListener.CommEvent = comEvRing Or strInput <> "" _
        Or commListener.CommEvent = comEvSend Or commListener.CommEvent = comEvReceive) Then Exit Sub
    
    If gobjCapturePar.CaptureWay = 1 Then 'ת������
        If mcpsComState.intComState <> commListener.CommEvent Then
           '����ۼ�ʱ�䳬���˲�ͼʱ��������ɼ�ͼ��
           If mcpsComState.lngComTime > gobjCapturePar.ComInterval Then
               'If Me.cbrMain.FindControl(, conMenu_Cap_MarkMap).Enabled Then
               If Not mblnReadOnly Then
                    Call subCaptureImg(True)
               End If
           End If
           
           '��¼�µ�COM״̬����ʱ�����㣬����timer
           mcpsComState.intComState = commListener.CommEvent
           mcpsComState.lngComTime = 0
           
           tmrComm.Enabled = True
        End If
    ElseIf gobjCapturePar.CaptureWay = 0 Then   'ֱ�Ӵ���
        '���β��½�̤��ʱ������������3��
        If DateDiff("S", mcpsComState.dtLastCapture, time) < gobjCapturePar.ComInterval Then
            mcpsComState.dtLastCapture = time
            
            Exit Sub
        End If
        
        mcpsComState.dtLastCapture = time
        
        If Not mblnReadOnly Then
            Call subCaptureImg(True)
        End If
    Else    '��ƽ����
        '���ڵ�ƽ����������������½�̤��ʱ�򣬶�Ӧ�ߵĵ�ƽ����֣���-��-�ͣ��򣨸�-��-�ߣ��ı仯
        'ͨ����ƽ�仯������ȷ���Ƿ���˽�̤��
        '�����ֵ�������ʱ����Ȼ�����OnComm�¼������ǵ�ƽ���ᷢ���仯��
        'ͨ���жϵ�ǰ��ƽ����̬��ƽ�Ƿ���ͬ��ȷ����ƽ�Ƿ����˱仯��
        
        '�жϵ�ƽ�Ƿ�ı䣬�ж�CTS��
        If mcpsComState.blnCTSHolding <> commListener.CTSHolding Then
            '�����񵴣�ë�������ж����δ�����ʱ���Ƿ�С���趨�ļ��
            If DateDiff("S", mcpsComState.dtLastCapture, time) < gobjCapturePar.ComInterval Then
                mcpsComState.dtLastCapture = time
                
                Exit Sub
            End If
            
            mcpsComState.dtLastCapture = time
            
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
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
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
  
  If mVideoSize.Width = 0 Or mVideoSize.Height = 0 Then Exit Sub
  If (mVideoArea.Height <= 0) Or (mVideoArea.Width <= 0) Then Exit Sub

  
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


Public Sub HideBorder()
    '���ش��ڵı����
    Dim lngWindowStyle As Long
    
    lngWindowStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
    lngWindowStyle = lngWindowStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
    
    Call SetWindowLong(Me.hwnd, GWL_STYLE, lngWindowStyle Or WS_CHILD)
End Sub

Private Sub OpenVideoCaptureDevice()
'����Ƶ�ɼ��豸
    Dim blnIsStartupVideo As Boolean

BUGEX "OpenVideoCaptureDevice 1"

    If mVideoCapture Is Nothing Then
        '������Ƶ�ɼ�����
        Set mVideoCapture = New clsVideoCapture
        
        '������Ƶ������
        Call mVideoCapture.ConnectedVfwDeviceObj(picVideo)
        Call mVideoCapture.ConnectedWdmDeviceObj(wdmCapture)
        Call mVideoCapture.ConnectedCustomDeviceObj(mobjCustomDevice)
        
        '��ȡ�����ļ�
        Call mVideoCapture.ReadCaptureParameterFromFile(GetAppPath & "\" & CAPTURE_PARAMETER_CONFIG_FILE_NAME)

        '������Ƶ����ʾģʽ
        Call mVideoCapture.SetVideoShowWay(swStretch)

        '�ڶ�ȡ�ļ����ú��޸ĸ����ԣ�ֻ�����ø����ԣ����ܸ��������߿���е��ں���ʾ��
        wdmCapture.AppHandle = Me.hwnd
        wdmCapture.IsShowState = False

        mdblZoomRate = 1
    End If
    
    mstrVideoRegTime = funVideoRegTime(Me)
    If mstrVideoRegTime = "" Then mstrMsg = "��ƵԴ����������������ϵ����Ա������������н������ã�"
    
    If UCase(Command()) = "DEBUG" Then
        mstrVideoRegTime = Now
    End If
    
    If Not mVideoCapture.IsStartup Then
        
        '������Ƶ��������
        mVideoCapture.VideoDriverType = gobjCapturePar.VideoDirverType
    
        '��ȡ��Ƶ��С
        mVideoSize.Width = ScaleX(mVideoCapture.VideoSize.Width, vbPixels, vbTwips)
        mVideoSize.Height = ScaleX(mVideoCapture.VideoSize.Height, vbPixels, vbTwips)
        
        '���ý���
        Call CaptureSwitchFace(IsTwainCaptureWay Or IsCustomCaptureWay)
        

        '*******************************************************
BUGEX "OpenVideoCaptureDevice 5"
        '��ʼ��ƵԤ��********************************************
        If Not IsTwainCaptureWay And Not IsCustomCaptureWay Then
            mblnRealTime = True
            
            Call mVideoCapture.StartPreview
                    
            blnIsStartupVideo = mVideoCapture.IsStartup
        Else
            mblnRealTime = False
            
            blnIsStartupVideo = ImageScanner.ScannerAvailable
        End If
 

        '*********************************************************
BUGEX "OpenVideoCaptureDevice 8"
    '    If mVideoCapture.IsStartup Then Call ucCapHook.EnableHook
    Else
        Call ConfigVideoShowState(True)
    End If
    
    Call OpenComm   '�򿪲ɼ��˿�
    
'    If gobjCapturePar.strCaptureHot <> "" Then Call mobjHotHook.EnableHook(WM_KEYDOWN, True)
End Sub


Public Sub UpdateAfterCaptureInfo()
'���º�̨�ɼ���Ϣ
    
    'ֻ��Ӱ��ɼ�ģ�鲢�����ú��̨�ɼ�����ʹ�ú�̨�ɼ�
    If Not IsTwainCaptureWay And Not IsCustomCaptureWay Then
        Call CreateNewCaptureTag
        Call ShowAfterCaptureInf(False)
    Else
        Call ShowAfterCaptureInf(False)
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) = 0 Then Exit Sub
    If (Shift And vbShiftMask) = 0 Then Exit Sub
    If (Shift And vbAltMask) = 0 Then Exit Sub
    
    If KeyCode <> 86 Then Exit Sub
    
    Call ShowVideoConfig
End Sub

Private Sub Form_Load()
  On Error GoTo errHandle
    '���ô�����ʽ
'    Call SetWindowStyle
'    Set mobjHotHook = New clsHookKey

    DTPImg.value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    mdate = DTPImg.value
    Call LoadMiniCache
    
    '���������Ըô��ڶ�������ö�������������ִ�д򿪻��߱������ʱ���������ļ�ѡ���λ�ڸô���֮��
    SetWindowPos Me.hwnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3  '�������ö�
    
    mstrAfterStationName = AnalyseComputer
    
    InitCommandBars
    
    Call InitReportImageFace
            
    ucPreview.PageImgCount = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "��Ƶ����ͼ����", 6))
    ucPreview.ShowCheckBox = True
    
    ucCacheImage.PageImgCount = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "��Ƶ����ͼ����", 6))
    ucCacheImage.ShowCheckBox = True
    
    mstrTempDirOfScan = GetAppPath + CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME
    If Len(mstrTempDirOfScan) > 45 Then
        Dim strFolder As String
        Dim pathlen As Long
        
        strFolder = String(256, 0)
        pathlen = GetTempPath(256, strFolder)
        If pathlen > 0 Then
            mstrTempDirOfScan = Left(strFolder, pathlen - 1) + CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME
        End If
    End If
    
    Set mfrmParameter = New frmVideoSetup
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub InitReportImageFace()
    Dim objImage As Pane, objCache As Pane
    
    With dkpReportImage
        .CloseAll
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        .PanelPaintManager.BoldSelected = True
        .TabPaintManager.position = xtpTabPositionLeft  'TAB�ŵ������ʾ
        .TabPaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .TabPaintManager.BoldSelected = True
    End With
    
    Set objImage = dkpReportImage.CreatePane(1, 0, 400, DockLeftOf)
    objImage.Title = "�� �� ͼ"
    objImage.Handle = picPreview.hwnd
    objImage.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    If gobjCapturePar.IsUseAfterCapture = True Then
        Call UseAfterImageChanged(True)
    Else
        picCacheImage.Visible = False
    End If
    
    objImage.Selected = True
End Sub


'�����Ƿ�ΪTWAIN�Ĳɼ���ʽ
Private Function IsTwainCaptureWay() As Boolean
    IsTwainCaptureWay = IIf(gobjCapturePar.VideoDirverType = vdtTWAIN, True, False)
End Function

Private Function IsCustomCaptureWay() As Boolean
    IsCustomCaptureWay = IIf(gobjCapturePar.VideoDirverType = vdtCustom, True, False)
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
Private Sub UpdateCaptureDirver(ByVal videoDirver As TVideoDriverType)

    '��ֹͣ��Ƶ��Ԥ��
    Call mVideoCapture.StopPreview
    
    gobjCapturePar.VideoDirverType = videoDirver
    mVideoCapture.VideoDriverType = videoDirver
       
    '��ȡ��Ƶ��С
    mVideoSize.Width = ScaleX(mVideoCapture.VideoSize.Width, vbPixels, vbTwips)
    mVideoSize.Height = ScaleX(mVideoCapture.VideoSize.Height, vbPixels, vbTwips)
       
    Call CaptureSwitchFace(videoDirver = vdtTWAIN Or videoDirver = vdtCustom)
        
    
    '�������Twain�ɼ���ʽ������������Ԥ��
    If videoDirver <> vdtTWAIN And videoDirver <> vdtCustom Then
        mblnRealTime = True
      
        '��ʼԤ��
        Call mVideoCapture.StartPreview
        
        'ˢ����ƵԤ������
        Call mVideoCapture.RefreshVideoWindow
    Else
        If videoDirver = vdtCustom Then
            '��ʼ��ר����Ƶ�ɼ��ӿ�
            Call InitCustomDevice
        End If
        
        mblnRealTime = False
    End If
End Sub


'���浱ǰ��������
Private Sub SaveParameterCfg()
BUGEX "SaveParameterCfg 1"
    
  '�ü���������
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblX1Scale", mCurCutRange.LeftRate
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblX2Scale", mCurCutRange.WidthRate
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblY1Scale", mCurCutRange.TopRate
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "mdblY2Scale", mCurCutRange.HeightRate
  
  
  '��ʾ��������
  SaveSetting "ZLSOFT", G_STR_REG_PATH_PUBLIC, "��ʾ��������", mblnShowProcessBar
BUGEX "SaveParameterCfg 2"
        
  '����ɼ�����
  If Not mVideoCapture Is Nothing Then Call mVideoCapture.SaveCaptureParameterToFile(GetAppPath & "\" & CAPTURE_PARAMETER_CONFIG_FILE_NAME)
BUGEX "SaveParameterCfg 3"
End Sub


Private Sub OpenComm()
    On Error GoTo err
    
BUGEX "OpenComm 1"
BUGEX "OpenComm ComPortType:" & gobjCapturePar.ComPortType
    If gobjCapturePar.ComPortType = "��" Then Exit Sub
BUGEX "OpenComm 2"
    If gobjCapturePar.ComPortType = "COM" Then
BUGEX "OpenComm 3"
        If commListener.PortOpen Then Exit Sub
BUGEX "OpenComm 4"
        commListener.CommPort = gobjCapturePar.ComPortName
        commListener.Settings = "9600,N,8,1"
        commListener.InputMode = comInputModeText
        commListener.RThreshold = 1
        commListener.InBufferCount = 0
        commListener.InputLen = 0
        commListener.RTSEnable = True
                        
        commListener.PortOpen = True
BUGEX "OpenComm 5"
        '��¼��̬��ƽ��λ
        mcpsComState.blnCTSHolding = commListener.CTSHolding
BUGEX "OpenComm 6"
    Else
BUGEX "OpenComm 7"
        If mobjDxDevice Is Nothing Then
BUGEX "OpenComm 7.1"
            Set mobjDxDevice = New clsDxHidDevice
        Else
BUGEX "OpenComm 7.2"
        End If
BUGEX "OpenComm 8"
        '��DX�豸
        If mobjDxDevice.Handle = 0 Then Call mobjDxDevice.OpenDxDevice(gobjCapturePar.ComPortName)
BUGEX "OpenComm 9"
        tmrComm.Enabled = True
        tmrComm.Interval = 2
    End If
BUGEX "OpenComm 10"
    Exit Sub
err:
BUGEX "OpenComm 11"
    Call MsgboxCus("�˿ڴ򿪴���", vbOKOnly, G_STR_HINT_TITLE)
BUGEX "OpenComm 12"
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
    Dim dblZoom As Double
    
    If mblnDcmViewDown = True And Button = 1 And dcmView.Images.Count > 0 Then
        Select Case mintMouseState
            Case 1  '���ȶԱȶ�
                dcmView.Images(1).Width = dcmView.Images(1).Width + (X - mlngBaseX)
                dcmView.Images(1).Level = dcmView.Images(1).Level + (Y - mlngBaseY)
                
                mlngBaseX = X
                mlngBaseY = Y
            Case 2  '����
                dblZoom = dcmView.Images(1).ActualZoom
                dblZoom = dblZoom * (1 + (Y - mlngBaseY) * 0.001)
                
                If dblZoom < 64 And dblZoom > 0.01 Then
                    subCenterZoom Me, dcmView.Images(1), dcmView, dblZoom, mCorpSize
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

   
Public Sub subCaptureImg(ByVal RealTimeCap As Boolean, _
                        Optional ByVal strFileName As String = "", _
                        Optional ByRef picCapture As StdPicture = Nothing, _
                        Optional ByVal blnIsAfterCapture As Boolean = False, _
                        Optional ByVal blnUseCustom As Boolean = False)
'------------------------------------------------
'���ܣ��ɼ����洢ͼ��
'��������
'���أ��ޣ�ֱ�ӱ����²ɼ���ͼ��
'------------------------------------------------
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    If mstrVideoRegTime = "" Then
        MsgboxCus mstrMsg, vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    If blnIsAfterCapture Then
        If Not mVideoCapture.IsStartup Then Exit Sub
    Else
        If Not (Not mblnReadOnly And (mVideoCapture.IsStartup Or IsTwainCaptureWay)) Then Exit Sub
    End If
    
BUGEX "subCaptureImg 1"
    If funCaptureSingleImage(RealTimeCap, strFileName, picCapture, blnIsAfterCapture) = True Then
        If blnIsAfterCapture Then
            '����Ǻ�̨�ɼ������̨�ɼ��ɹ���ɾ����̨�ɼ���ͼ��
            If subSaveAfterCaptureImage Then Call dcmAfter.Images.Clear
            
            mAfterCaptureInf.lngAfterCurImageCount = GetCountOfCapPic(mintTagNow)
            
            Call ShowAfterCaptureInf(True)
                        
            Exit Sub
        End If
        
        If IsCustomCaptureWay And blnUseCustom Then Exit Sub
        
BUGEX "subCaptureImg 2"
        mintCaptureFlag = 2
        
        Call subSaveImage(mlngAdviceId, mcurStudyInf.strStudyUid)
    End If
BUGEX "subCaptureImg 5"
Exit Sub
errHandle:
    err.Raise err.Number, err.Description
End Sub

Private Function CopyPictureToDicomImg(ByVal lngHDC As Long, ByVal lngPictureHandle As Long, objDcmImg As Object) As Boolean
'congpicture�и���ͼ��dicomimage
    Const bitCount As Long = 3
        
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
                                    Optional ByVal strFileName As String = "", _
                                    Optional ByRef picCapture As StdPicture = Nothing, _
                                    Optional ByVal blnIsAfterCapture As Boolean = False) As Boolean
'------------------------------------------------
'���ܣ��ɼ���֡��Ƶͼ�񣬽�ͼ��ת����DICOM��ʽ������дDICOM�ļ�ͷ�����ͼ���������ͼdcmMiniature�С�
'��������
'���أ��ޣ�ֱ�ӽ��²ɼ���ͼ�����dcmMiniature��
'------------------------------------------------
'�ɼ���֡ͼ��
On Error GoTo SaveFileError
    Dim ImgTmpImage As DicomImage
    Dim dcmTag As clsImageTagInf
    
    '�ɼ�ͼ�񣬷�Ϊֱ����Ƶ�ɼ��Ͳ���¼��ɼ�

    If Not (picCapture Is Nothing) Then
        Set picTemp2.Picture = Nothing
        picTemp2.Width = mVideoSize.Width * (1 - mCurCutRange.WidthRate - mCurCutRange.LeftRate)
        picTemp2.Height = mVideoSize.Height * (1 - mCurCutRange.HeightRate - mCurCutRange.TopRate)
        picTemp2.Picture = picCapture
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
                        
            Dim curPic As StdPicture
            Set curPic = mVideoCapture.CaptureImageFromMemory

            If curPic Is Nothing Then
                Call MsgboxCus("��Ƶͼ��ɼ�ʧ�ܣ�������Ƶ���������Ƿ���ȷ(����Ƶ�豸����ʾģʽ��)��", vbOKOnly, G_STR_HINT_TITLE)
                
                funCaptureSingleImage = False
                Exit Function
            End If
            
            picTemp2.Width = mVideoSize.Width * (1 - mCurCutRange.WidthRate - mCurCutRange.LeftRate)
            picTemp2.Height = mVideoSize.Height * (1 - mCurCutRange.HeightRate - mCurCutRange.TopRate)

            'Ӧ��ͼ��Χ�ü�
            Call picTemp2.PaintPicture(curPic, 0, 0, picTemp2.Width, picTemp2.Height, _
                                       mVideoSize.Width * mCurCutRange.LeftRate, mVideoSize.Height * mCurCutRange.TopRate, _
                                       picTemp2.Width, picTemp2.Height, vbSrcCopy)
                                               
            picTemp2.Picture = picTemp2.Image

            Set curPic = Nothing
        End If
    End If
    
    If picTemp2.Picture Is Nothing Then
        funCaptureSingleImage = False
        Exit Function
    End If

    '����dicom��ʽͼ��
    Set ImgTmpImage = New DicomImage
    
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
    Call subWriteDicomPara(ImgTmpImage, mlngAdviceId, blnIsAfterCapture)
    
    Set dcmTag = New clsImageTagInf
    dcmTag.Tag = imgTag
    
    Set ImgTmpImage.Tag = dcmTag
    
    If blnIsAfterCapture Then
        Call dcmAfter.Images.Add(ImgTmpImage)
    Else
        '��ͼ���������ͼ��
        Call subInsert2Mini(ImgTmpImage)
    End If
    
BUGEX "dcmTag:" & ImgTmpImage.Tag.Tag
    
    funCaptureSingleImage = True

    Exit Function
SaveFileError:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function


Private Sub subWriteDicomPara(img As DicomImage, lngAdviceId As Long, _
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
        img.Attributes.Add &H8, &H60, mcurStudyInf.strModality                   'Modality Ӱ�����
    Else
        img.Attributes.Add &H10, &H10, mcurStudyInf.strName                     'Name ����
        img.Attributes.Add &H10, &H20, mcurStudyInf.strPatientID                'Patient ID ����ID
        img.Attributes.Add &H10, &H30, mcurStudyInf.strBirthDate                'BirthDate ����
        img.Attributes.Add &H10, &H40, mcurStudyInf.strSex                      'Sex �Ա�
        img.Attributes.Add &H10, &H1010, mcurStudyInf.strAge                    'Age ����
        img.Attributes.Add &H10, &H4000, ""                         'Patient Comment ����ע��
        img.Attributes.Add &H20, &H10, mcurStudyInf.strCheckNo                  'Study ID ���ID
        img.Attributes.Add &H8, &H60, mcurStudyInf.strModality                   'Modality Ӱ�����
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

Private Sub UniteUID(img As DicomImage)
    Set mdcmTmpImg = img
    
    '�������Ƶ,������Ƶ������������UID
    '��������ͼ�ļ��UID������UID���޸�img��ֵ
    Call subUniteUID(mdcmTmpImg, mdcmTmpImg.Tag.Tag <> VIDEOTAG And mdcmTmpImg.Tag.Tag <> AUDIOTAG)
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
    
    ucPreview.AddImage img, img.Tag
End Sub

Private Sub Form_Resize()
On Error GoTo errHandle
BUGEX "Form_Resize(frmWork_Video) 1"

    Call ucSplitter1.RePaint(False)
BUGEX "Form_Resize(frmWork_Video) 2"

BUGEX "Form_Resize(frmWork_Video) picCaptureHeight:" & picCapture.Height

Exit Sub
errHandle:
BUGEX "Form_Resize(frmWork_Video) Err:" & err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    mdate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    'LoadMiniCache (True) '����������µĿ�Ŀ¼
    ClearEmptyFolder (True)

BUGEX "VideoForm_UnLoad 1"
    tmrComm.Enabled = False
    timerHook.Enabled = False
    
BUGEX "VideoForm_UnLoad 3"
    '�ȹرղɼ����ں�COMM��
    Call StopCapture
BUGEX "VideoForm_UnLoad 4"
    '���ֲü�����
    Call SaveParameterCfg
BUGEX "VideoForm_UnLoad 5"
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "��Ƶ����ͼ����", ucPreview.PageImgCount)
    
BUGEX "VideoForm_UnLoad 6"
    If Not mfrmParameter Is Nothing Then
        Unload mfrmParameter
    End If
BUGEX "VideoForm_UnLoad 7"
    '�Ͽ�ftp����
    If Not mobjFtpConnection Is Nothing Then
        Call mobjFtpConnection.FuncFtpDisConnect
        Set mobjFtpConnection = Nothing
    End If
BUGEX "VideoForm_UnLoad 8"
    '�Ͽ�����ftp����
    If Not mobjBakFtpConnection Is Nothing Then
        Call mobjBakFtpConnection.FuncFtpDisConnect
        Set mobjBakFtpConnection = Nothing
    End If
    
BUGEX "VideoForm_UnLoad 9"
    If Not mfrmOpenStudy Is Nothing Then
        Unload mfrmOpenStudy
        Set mfrmOpenStudy = Nothing
    End If
    
BUGEX "VideoForm_UnLoad 10"
    wdmCapture.FreeRes
BUGEX "VideoForm_UnLoad 11"

'    Call mobjHotHook.FreeHook
'    Set mobjHotHook = Nothing
    
    Set dcmglbUID = Nothing
    Set mobjDxDevice = Nothing
    Set mVideoCapture = Nothing
    Set mfrmParameter = Nothing
    
    If Not mobjCustomDevice Is Nothing Then
        mobjCustomDevice.zlFree
        Set mobjCustomDevice = Nothing
    End If
BUGEX "VideoForm_UnLoad End"
End Sub


Private Sub subDeleteImage()
'------------------------------------------------
'���ܣ�ɾ������ͼ�б�ѡ�е�ͼ���ȴ����ݿ���ɾ����Ȼ���FTP��ɾ����ɾ���󴥷�StateChanged�¼�
'��������
'���أ��ޣ�ֱ��ɾ������ͼ�����һ��ͼ��
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnResult As Boolean
    Dim i As Long
    Dim lngCount As Integer '���Ͻ��к��ǵ�ͼ����
    Dim strImgUids As String
    Dim lngIndexTemp As Long
    Dim lngDelCount As Long
    
    If ucPreview.ImgViewer.Images.Count > 0 Then
        lngCount = 0
        '�õ���Ҫɾ����ͼ��uid�м���';'����
        For i = 1 To ucPreview.CurImageCount
            If ucPreview.ImgChecked(i) = True Then
                lngIndexTemp = i
                lngCount = lngCount + 1
                strImgUids = IIf(lngCount = 1, ucPreview.ImgViewer.Images(i).InstanceUID, strImgUids & ";" & ucPreview.ImgViewer.Images(i).InstanceUID)
            End If
        Next
                
        If lngCount > 100 Then
            MsgboxCus "�Բ��𣬳���֧��һ����ɾ����ͼ���ܳ���100�ţ�������ѡ��", vbInformation, G_STR_HINT_TITLE
            Exit Sub
        End If
      
        '�����ݿ��FTP��ɾ������ͼ�б�ѡ�е�ͼ��
        If lngCount > 1 Then
            blnResult = DeleteManyImages(strImgUids)
        ElseIf lngCount = 1 Then
        '�������1 ���ܺ��ͺ����ͬһ��ͼ�񣬴�ʱɾ�����ͼ��
            blnResult = DeleteImages(Me, 1, strImgUids, "")
        Else
            blnResult = DeleteImages(Me, 1, ucPreview.SelectImage.InstanceUID, "")
        End If
        
        If blnResult = True Then    'ɾ���ɹ������޸�����ͼ״̬��������StateChanged�¼�
            '������ͼ��ɾ��ͼ��
            lngDelCount = lngCount
            If lngCount > 1 Then
                For i = ucPreview.CurImageCount To 1 Step -1
                    If ucPreview.ImgChecked(i) = True Then
                        If lngDelCount <> 1 Then
                            Call ucPreview.DeleteImage(i, False)
                            lngDelCount = lngDelCount - 1
                        Else
                            Call ucPreview.DeleteImage(i, True, True)
                        End If
                    End If
                Next
            ElseIf lngCount = 1 Then
                Call ucPreview.DeleteImage(lngIndexTemp)
            ElseIf lngCount = 0 Then
                Call ucPreview.DeleteImage(ucPreview.SelectIndex)
            End If
            dcmView.Images.Clear

            If Not ucPreview.SelectImage Is Nothing Then
                dcmView.Images.Add ucPreview.SelectImage
            End If
            
            
            '����Ӱ����״̬�����ɾ�����һ��ͼ����ԭ������Ϊ3�����޸�Ϊ2
            If ucPreview.CurImageCount = 0 Then
                
                If mlngStudyState = 3 Then
                    strSQL = "Zl_Ӱ����_State(" & mlngAdviceId & "," & mlngSendNo & ",2,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & glngDepartId & ")"
                    zlDatabase.ExecuteProcedure strSQL, "ɾ�����һ��ͼ��"
                    mlngStudyState = 2
                End If
                
                Call DoStateChange(vetDelAllImg, mlngAdviceId, mlngSendNo, mcurStudyInf.strStudyUid)
                
                mcurStudyInf.strStudyUid = ""
                
                '������ͼ��ɾ��ʱ������ʾʵʱ��Ƶ����
                Call ConfigVideoShowState(True)
            Else
                Call DoStateChange(vetImgDeled, mlngAdviceId, mlngSendNo, mcurStudyInf.strStudyUid)
            End If
        
        End If
    End If
End Sub


Private Sub subSetMouseState(intMouseState As Integer)
    '�ı䵱ǰ���״̬
    mintMouseState = IIf(mintMouseState = intMouseState, 0, intMouseState)
    
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Window).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Zoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_RectZoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Arrow).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Ellipse).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Text).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Corp).Checked = False
End Sub


'modify by tjh at 2010-01-20
'������Ƶ��ʾ״̬
Private Sub ConfigVideoShowState(ByVal blnShowState As Boolean)
  mblnRealTime = blnShowState
  
  Select Case gobjCapturePar.VideoDirverType
    Case vdtVFW
      picVideo.Visible = blnShowState
      wdmCapture.Visible = False
      dcmView.Visible = Not blnShowState
    Case vdtWDM
      wdmCapture.Visible = blnShowState
      picVideo.Visible = False
      dcmView.Visible = Not blnShowState
    Case vdtTWAIN, vdtCustom
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

Public Sub UseAfterImageChanged(ByVal blUse As Boolean)
    Dim objImage As Pane, objCache As Pane, objTmp As Pane
    Dim blHavePane As Boolean
    Dim i As Integer
    
    If blUse = True Then
    '�Ƿ���Ҫ�������ж�
        blHavePane = False
        For i = 1 To dkpReportImage.PanesCount
            Set objTmp = dkpReportImage.FindPane(i)
            
            If objTmp.Title = "�� �� ͼ" Then
                Set objImage = dkpReportImage.FindPane(i)
            End If
            
            If objTmp.Title = "�� ̨ ͼ" Then
                blHavePane = False
                Exit Sub
            End If
        Next
        If blHavePane = False Then
            picCacheImage.Visible = True
            Set objCache = dkpReportImage.CreatePane(2, 0, 400, DockLeftOf)
            objCache.Title = "�� ̨ ͼ"
            objCache.Handle = picCacheImage.hwnd
            objCache.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            objCache.AttachTo objImage
            objImage.Selected = True
            
            LoadMiniCache
        End If
    Else
    '�Ƿ���Ҫ���ٵ��ж�
        blHavePane = False
        
        For i = 1 To dkpReportImage.PanesCount
            Set objTmp = dkpReportImage.FindPane(i)
            
            If objTmp.Title = "�� ̨ ͼ" Then
                blHavePane = True
                Exit For
            End If
        Next

        If blHavePane = True Then Call dkpReportImage.DestroyPane(objTmp)

        picCacheImage.Visible = False

    End If
  
    Exit Sub
errH:
    Call err.Raise(0, , "��̨ͼ��ǩ�������" & err.Description)
End Sub

Private Sub mfrmParameter_OnIsUseAfterImgChange(ByVal blUse As Boolean)
On Error GoTo errHandle
'��̨�ɼ��������ý���ȷ�Ϻ󴥷����ж��Ƿ���ʾ��̨ͼpane���Ƿ�ɾ����̨ͼpane
    gobjCapturePar.IsUseAfterCapture = blUse
    
    Call UseAfterImageChanged(blUse)
    RaiseEvent OnIsUseAfterImgChange(blUse)
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub mobjDxDevice_OnDxKeyPress(ByVal lngButtonNum As Long)
On Error GoTo errHandle
BUGEX "mobjDxDevice_OnDxKeyPress 1"
BUGEX "mobjDxDevice_OnDxKeyPress ButtonNum:" & lngButtonNum

    Select Case lngButtonNum
        Case 0  'ǰ̨�ɼ�
BUGEX "mobjDxDevice_OnDxKeyPress 2"
            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_MarkMap).Enabled And _
                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_MarkMap).Visible Then
                Call subCaptureImg(True)
            End If
        Case 1  '��̨�ɼ�
BUGEX "mobjDxDevice_OnDxKeyPress 3"
            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Capture).Enabled And _
                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Capture).Visible Then
                Call subCaptureImg(True, "", Nothing, True)
            Else
                Call mobjDxDevice_OnDxKeyPress(0)
            End If
        Case 2  '���±�ʶ
BUGEX "mobjDxDevice_OnDxKeyPress 4"
            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Tag).Enabled And _
                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_After_Tag).Visible Then
                
                If gobjCapturePar.IsUseAfterCapture Then Call UpdateAfterCaptureInfo
            Else
                Call mobjDxDevice_OnDxKeyPress(0)
            End If
        Case Else
            Call mobjDxDevice_OnDxKeyPress(0)
    End Select
    
BUGEX "mobjDxDevice_OnDxKeyPress 5"
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub mfrmParameter_OnVideoDirverChange(ByVal vdtDirverType As TVideoDriverType)
'�����ı�󣬸��²ɼ�����
On Error GoTo errHandle
    Call mVideoCapture.StopPreview
    
    mVideoCapture.VideoDriverType = vdtDirverType
    
    Call UpdateCaptureDirver(vdtDirverType)
    
'    '���ΪTWAIN�ķ�ʽ���򲻽�����Ƶ��ˢ��
'    If mVideoCapture.VideoDriverType <> vdtTWAIN Then
'        Call mVideoCapture.StartPreview
'
'        Call mVideoCapture.RefreshVideoWindow
'    End If
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub mobjHotHook_OnKeyBoardLHook(ByVal lngMsg As Long, ByVal lngVkCode As Long, ByVal lngScanCode As Long, ByVal lngFlags As Long)
'    Dim lngWindowPID As Long
'    Dim lngParentPID As Long
'    Dim lngCurrentPID As Long
'
'    If lngMsg <> WM_KEYDOWN Then Exit Sub
'
'    '�жϴ�����Ϣ���Ƿ�Ϊ��ǰ����
'    Call GetWindowThreadProcessId(GetActiveWindow(), lngWindowPID)
'    Call GetWindowThreadProcessId(glngRootHandle, lngParentPID)
'
'    lngCurrentPID = GetCurrentProcessId
'
'
'    If lngCurrentPID = lngWindowPID Or lngWindowPID = lngParentPID Then
'
'
'
'        'ʹ���ȼ����вɼ�
'        If GetKeyAliasEx(lngVkCode) = gobjCapturePar.strCaptureHot Then
'            If Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_MarkMap).Enabled And _
'                Me.cbrMain.FindControl(xtpControlButton, conMenu_Cap_MarkMap).Visible Then
'                Call subCaptureImg(True)
'            End If
'        End If
'    End If
End Sub

'modify by tjh at 2010-01-20
Private Sub pbxSize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '��ʼִ�вü���Χ����
    If Button = 1 And gobjCapturePar.IsAllowChangeSize Then
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
    
    If IsTwainCaptureWay Or IsCustomCaptureWay Then
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
    
    picCapture.Refresh
    
errHandle:
End Sub


Private Function LoadPlayVideo() As String
'���ز�����Ƶ
On Error GoTo errHandle
    If dcmView.Images(1).Tag.Tag <> VIDEOTAG And dcmView.Images(1).Tag.Tag <> AUDIOTAG Then Exit Function
    
    If dcmView.Images(1).Tag.Tag = VIDEOTAG Then
        Call dcmView.Images(1).FileImport(GetAppPath & "\..\�����ļ�\aviDownload.bmp", "DIB/BMP")
    
        '������Ҫ���ŵ���Ƶ
        LoadPlayVideo = GetSingleImage(dcmView.Images(1).InstanceUID, dcmView.Images(1).SeriesUID, Me, mblnMoved)
    
        Call dcmView.Images(1).FileImport(GetAppPath & "\..\�����ļ�\avi.bmp", "DIB/BMP")
    Else
        Call dcmView.Images(1).FileImport(GetAppPath & "\..\�����ļ�\wavDownload.bmp", "DIB/BMP")
    
        '������Ҫ���ŵ���Ƶ
        LoadPlayVideo = GetSingleImage(dcmView.Images(1).InstanceUID, dcmView.Images(1).SeriesUID, Me, mblnMoved)
    
        Call dcmView.Images(1).FileImport(GetAppPath & "\..\�����ļ�\wav.bmp", "DIB/BMP")
    End If
errHandle:
End Function

Private Sub subVideoPlay()
'------------------------------------------------
'���ܣ�dcmView��¼��ͼ��Ĳ���
'��������
'���أ��ޣ�ֱ�Ӳ���dcmView�е�ͼ��
'------------------------------------------------
    Dim strFile As String
    
    If dcmView.Images.Count > 0 Then
        '����¼��������ش��ڣ��򲻽�������
        If dcmView.Images(1).Tag.Tag <> VIDEOTAG And dcmView.Images(1).Tag.Tag <> AUDIOTAG Then Exit Sub
        
        strFile = LoadPlayVideo
        
        '�򿪲��š���
        Call frmPlaying.Show
        
        'ˢ�²��Ŵ���
'       Call frmPlaying.Refresh
        While Not frmPlaying.IsActive
            Call Sleep(10)
            DoEvents
        Wend
            
        Call frmPlaying.OpenVideoFile(Replace(strFile, "/", "\"), Me)
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
            strFileName = dlgOpen.FileName
            
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
        strFileName = dlgOpen.FileName
        
        If strFileName <> "" Then
            strFileType = UCase(Right(Trim(strFileName), 3))
            
            Select Case strFileType
                Case "AVI"
                    If dcmView.Images(1).FrameCount > 1 Then
                        dcmView.Images(1).WriteAVI strFileName, 1, dcmView.Images(1).FrameCount, 1, "", 100, False
                    Else
                        MsgboxCus "��̬ͼ���޷������AVI��ʽ��������ѡ��ͼ���ʽ��", vbInformation, G_STR_HINT_TITLE
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
On Error Resume Next
    Dim DlgInfo As DlgFileInfo
    Dim i As Integer
    Dim ImgTmpImage As New DicomImage
    Dim j As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    'ѡ���ļ�
    With Me.dlgOpen
        .CancelError = False
        .MaxFileSize = 32767 '���򿪵��ļ����ߴ�����Ϊ��󣬼�32K
        .flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
        .DialogTitle = "ѡ���ļ�"
        .Filter = "DICOM�ļ���*.dcm��(*.img)|*.dcm;*.img|ͼ���ļ� (*.BMP)(*.JPG)|*.BMP;*.JPG|�����ļ���*.*��|*.*"
        .ShowOpen
        If .FileName <> "" Then
            DlgInfo = GetDlgSelectFileInfo(.FileName)
        End If
        '�ڴ���*.pif�ļ����뽫Filename�����ÿգ�����ѡȡ���*.pif�ļ��󣬵�ǰ·����ı�
        .FileName = ""
    End With

    For i = 1 To DlgInfo.iCount
        Set ImgTmpImage = ReadViewImage(DlgInfo.sPath & DlgInfo.sFIle(i))
    
        
        '����ͼ���DICOM����
        subWriteDicomPara ImgTmpImage, mlngAdviceId
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.Tag = imgTag
    
        Set ImgTmpImage.Tag = dcmTag
        
        mintCaptureFlag = 3
        
        '��ͼ����뵽����ͼ��
        subInsert2Mini ImgTmpImage
            
        '����ͼ�񣬲�����ͼ��洢�¼�
        Call subSaveImage(mlngAdviceId, mcurStudyInf.strStudyUid)
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
                If ucPreview.ImgViewer.Images(i).Tag.Tag = imgTag Then
                    dcmImg.SeriesUID = ucPreview.ImgViewer.Images(i).SeriesUID
                    
                    Exit For
                End If
            Next i
            
        End If
    ElseIf Len(mcurStudyInf.strStudyUid) > 0 Then
        dcmImg.StudyUID = mcurStudyInf.strStudyUid
    Else
        mcurStudyInf.strStudyUid = dcmImg.StudyUID
        
        '�����uid�ı����Ҫ��������ͼ��ʾ����еĲ�ѯֵ
        ucPreview.QueryValue = mcurStudyInf.strStudyUid
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
    MsgboxCus "GetDlgSelectFileInfo����ִ�д���", vbOKOnly + vbCritical, G_STR_HINT_TITLE
End Function


Private Sub picDock_Paint()
BUGEX "picDock_Paint(frmWork_Video)"
End Sub

Private Sub picPreview_Resize()
On Error Resume Next
    ucPreview.Left = 0
    ucPreview.Top = 0
    ucPreview.Width = picPreview.ScaleWidth
    ucPreview.Height = picPreview.ScaleHeight
End Sub

Private Sub picCacheImage_Resize()
On Error Resume Next
    DTPImg.Left = 0
    DTPImg.Top = 0
    
    If mintFontSize = 12 Then
        DTPImg.Width = 1750
        DTPImg.Height = 360
    ElseIf mintFontSize = 15 Then
        DTPImg.Width = 2000
        DTPImg.Height = 420
    Else
        '����������ֺ�=9 ����
        DTPImg.Width = 1400
        DTPImg.Height = 300
    End If
    

    cboCacheImage.Left = DTPImg.Width
    cboCacheImage.Top = 0
    cboCacheImage.Width = picCacheImage.ScaleWidth - DTPImg.Width
    cboCacheImage.Height = picCacheImage.ScaleHeight
    
    ucCacheImage.Left = 0
    ucCacheImage.Top = DTPImg.Height
    ucCacheImage.Width = picCacheImage.ScaleWidth
    ucCacheImage.Height = picCacheImage.ScaleHeight - DTPImg.Height
End Sub

Private Sub TimerHook_Timer()
On Error GoTo errHandle
    '��ʹ��hook�ȼ����òɼ�ʱ��ʹ��timer���вɼ�������������ִ�ж��CaptureImage������hookʧЧ
    '���hookʧЧ�Ŀ���ԭ����hook�Ĵ������������ػ�hook��Ĵ���ʱ������������ʧЧ������dicomobjects��fileexport�������ö�����ʧЧ��Ŀǰ��ȥϸ��
    Call CaptureImage
    
    timerHook.Enabled = False
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub TimerRePaint_Timer()
 
    timerRePaint.Enabled = False

    Call cbrMain.RecalcLayout
    Call ucSplitter1.RedrawSelf
    Call ucPreview.RedrawSelf
    Call dcmView.Refresh
    Call picCapture.Refresh

    BUGEX "timerRePaint_Timer 1"
End Sub

Private Sub tmrComm_Timer()
    On Error GoTo errHandle
    If gobjCapturePar.ComPortType = "COM" Then
        mcpsComState.lngComTime = mcpsComState.lngComTime + 2
        
        '����0.08�룬���Զ�����
        If mcpsComState.lngComTime > 40 Then
            mcpsComState.lngComTime = 0
            
            tmrComm.Enabled = False
        End If
        
    Else
         If Not mobjDxDevice Is Nothing Then Call mobjDxDevice.PollDxDevice
    End If
    
    Exit Sub
errHandle:
    tmrComm.Enabled = False
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub tmrReg_Timer()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo errHandle:
    If Not mVideoCapture.IsStartup Then
        Exit Sub
    End If

    If gint��Ƶ�豸���� <= -1 Then Exit Sub
    
    strSQL = "select count(1) ���������� from zltools.zlclients where ������ƵԴ=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����������")
    
    If rsTemp.RecordCount > gint��Ƶ�豸���� Then
        mstrVideoRegTime = ""
        mstrMsg = "��Ȩ��Ƶվ�������뵱ǰʵ��ʹ��������ƥ�䣬���飡"
        Exit Sub
    End If
    
    If DateDiff("S", mstrVideoRegTime, Now) >= M_LNG_REFRESHINTERVAL Then
        '�ж����ݿ����Ƿ�����Ѿ�ע���ip�����Ѿ�������ƵԴ���������������Ϊû�гɹ�ע��
        If FunCheckRegInfo(Me) Then
            mstrVideoRegTime = Now
        Else
            mstrVideoRegTime = ""
            mstrMsg = "��ƵԴ����������������ϵ����Ա������������н������ã�"
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

Private Sub CustomVideoSave()
    Dim dcmTmpImg As New DicomImage
    Dim strVideoFiles As String
    Dim blnUseCustom As Boolean
    Dim strEncoderName As String '����������
    Dim lngRecordTimeLen As Long '¼����Ƶ����
    
    If mobjCustomDevice Is Nothing Then Exit Sub
    
    Call mobjCustomDevice.zlCaptureVideo(mlngAdviceId, strVideoFiles, blnUseCustom, strEncoderName, lngRecordTimeLen)
    
    '¼�����
    If Trim(strVideoFiles) <> "" And Dir(strVideoFiles) <> "" Then
        dcmTmpImg.FileImport GetAppPath & "\..\�����ļ�\avi.bmp", "DIB/BMP"
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.EncoderName = strEncoderName
        dcmTag.VideoFile = strVideoFiles
        dcmTag.CaptureTime = zlDatabase.Currentdate
        dcmTag.RecordTimeLen = lngRecordTimeLen
        dcmTag.Tag = VIDEOTAG
        
        Set dcmTmpImg.Tag = dcmTag
        
        subWriteDicomPara dcmTmpImg, mlngAdviceId
        
        mintCaptureFlag = 4
        
        subInsert2Mini dcmTmpImg
        
        '������Ƶ¼��
        Call subSaveImage(mlngAdviceId, mcurStudyInf.strStudyUid)
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
    If Trim(strError) <> "" Then MsgboxCus strError, vbInformation, G_STR_HINT_TITLE
    
    '��ȡ��ǰ¼��ı���������
    mstrEncoderName = mVideoCapture.GetEncoderName
    
    Exit Sub
CapErr:
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
    err.Clear
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
            dcmTmpImg.FileImport GetAppPath & "\..\�����ļ�\avi.bmp", "DIB/BMP"
continue2:
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.EncoderName = mstrEncoderName
        dcmTag.VideoFile = mstrAviFileName
        dcmTag.CaptureTime = zlDatabase.Currentdate
        dcmTag.RecordTimeLen = mVideoCapture.GetTimeLen
        dcmTag.Tag = VIDEOTAG
        
        Set dcmTmpImg.Tag = dcmTag
        
        subWriteDicomPara dcmTmpImg, mlngAdviceId
        
        mintCaptureFlag = 4
        
        subInsert2Mini dcmTmpImg
        
        '������Ƶ¼��
        Call subSaveImage(mlngAdviceId, mcurStudyInf.strStudyUid)
    End If
    
    Exit Sub
CapErr:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


'ֹͣ��Ƶ�ļ�
Public Sub subSaveAudio(ByVal strAudioFile As String, ByVal lngTimeLen As Long)

    Dim i As Integer
    Dim dcmTmpImg As New DicomImage
    
    On Error GoTo CapErr
   
    
    '¼�����
    If Dir(strAudioFile) <> "" Then
        dcmTmpImg.FileImport GetAppPath & "\..\�����ļ�\wav.bmp", "DIB/BMP"
        
        Dim dcmTag As New clsImageTagInf
        dcmTag.EncoderName = ""
        dcmTag.VideoFile = strAudioFile
        dcmTag.CaptureTime = zlDatabase.Currentdate
        dcmTag.RecordTimeLen = lngTimeLen
        dcmTag.Tag = AUDIOTAG
        
        Set dcmTmpImg.Tag = dcmTag
        
        subWriteDicomPara dcmTmpImg, mlngAdviceId
        
        mintCaptureFlag = 5
        
        subInsert2Mini dcmTmpImg
        
        '����¼�Ƶ���Ƶ
        Call subSaveImage(mlngAdviceId, mcurStudyInf.strStudyUid)
    End If
    
    Exit Sub
CapErr:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

'modify by tjh at 2010-01-22
'ȫ����ʾ
Private Sub subFullCall()
  Call mVideoCapture.FullScreen(Me, Me.hwnd)
End Sub

Private Function CreatFolder(ByVal intTag As Integer) As Boolean
    Dim strAfterImgPath As String
    Dim strReceivedTime As String
    Dim curTime As Date
    Dim i As Integer
    
    curTime = zlDatabase.Currentdate
    CreatFolder = False

    '���uidΪ�գ��򴴽��µ�UID
    mAfterCaptureInf.strAfterStudyUid = dcmglbUID.NewUID
    mAfterCaptureInf.strAfterSeriesUid = dcmglbUID.NewUID
    
    mAfterCaptureInf.intAfterTag = intTag
    strReceivedTime = Format(curTime, "yyyyMMdd")
    
    strAfterImgPath = mstrAfterDir & "\" & strReceivedTime & "\" & "��ʶ" & mAfterCaptureInf.intAfterTag & "-" & mAfterCaptureInf.strAfterStudyUid & "\" & mAfterCaptureInf.strAfterSeriesUid & "\"

    MkLocalDir strAfterImgPath


    
    cboCacheImage.AddItem "��ʶ" & Lpad(mAfterCaptureInf.intAfterTag, 3, "0") & "     ʱ�䣺" & Format(curTime, "HH:MM:SS"), 0

End Function

Private Function GetCaptureTag() As Integer
'��õ�ǰ��ʶ
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    Call GetTodayTagMax(curDate)  '���µ�������ʶ
           
    If GetCountOfCapPic(mintTagMax) = 0 Then
        '��ͼ�񣬼���ʹ�������־
        mintTagNow = mintTagMax
    Else
        '��ͼ�񣬸��±�ʶ���Ҵ�����Ŀ¼
        mintTagNow = mintTagMax + 1
        CreatFolder (mintTagNow)
    End If
    
    '���mintTagNow=1���ж�Ŀ¼��ʶ1�Ƿ���ڲ������Ƿ񴴽�
    If mintTagNow = 1 Then
        If JudgeExistFolderOne = False Then
            Call CreatFolder(1)
        End If
    End If

    GetCaptureTag = mintTagNow
End Function

Private Sub CreateNewCaptureTag()
'ȡ���µĲɼ���ʶ

    mAfterCaptureInf.intAfterTag = GetCaptureTag

    mAfterCaptureInf.lngAfterCurImageCount = 0
End Sub

Public Sub showAfterCapInfo(ByVal intType As Integer, ByVal isNeedRefreshTitle As Boolean)

    If isNeedRefreshTitle = True Then
        mAfterCaptureInf.lngAfterCurImageCount = GetCountOfCapPic(mintTagNow)
        Call ShowAfterCaptureInf(True)
    End If
    
    If intType = 1 Then mblIsNeedRefresh = True
End Sub


Private Sub ShowAfterCaptureInf(ByVal blnShowTag As Boolean)
'���º�̨�ɼ�ͼ����Ϣ
    Dim objPanel As Pane
    Dim i As Long
    
    If gobjOwner Is Nothing Then
        For i = 1 To dkpReportImage.PanesCount
            Set objPanel = dkpReportImage.Panes(i)
            If InStr(objPanel.Title, "�� ̨ ͼ") > 0 Or InStr(objPanel.Title, "��̨ͼ") > 0 Then
                objPanel.Title = "��̨ͼ - ��ǰ��ʶ:" & mintTagNow
                Exit Sub
            End If
        Next i
        
        Exit Sub
    Else
        For i = 1 To dkpReportImage.PanesCount
            Set objPanel = dkpReportImage.Panes(i)
            If InStr(objPanel.Title, "�� ̨ ͼ") > 0 Or InStr(objPanel.Title, "��̨ͼ") > 0 Then
                objPanel.Title = "�� ̨ ͼ"
                Exit For
            End If
        Next i
    End If
    
    If Not gobjCapturePar.IsUseAfterCapture Or blnShowTag = False Then
        If InStr(gobjOwner.Caption, "      ��ǰ��̨�ɼ���ʶ��") > 0 Then
            gobjOwner.Caption = Mid(gobjOwner.Caption, 1, InStr(gobjOwner.Caption, "      ��ǰ��̨�ɼ���ʶ��") - 1)
        End If
            
        Exit Sub
    End If
    
    If mAfterCaptureInf.strAfterParentTitle = "" Then
        If InStr(gobjOwner.Caption, "      ��ǰ��̨�ɼ���ʶ��") > 0 Then
            mAfterCaptureInf.strAfterParentTitle = Mid(gobjOwner.Caption, 1, InStr(gobjOwner.Caption, "      ��ǰ��̨�ɼ���ʶ��") - 1)
        Else
            mAfterCaptureInf.strAfterParentTitle = gobjOwner.Caption
        End If
    End If
    
    gobjOwner.Caption = mAfterCaptureInf.strAfterParentTitle & "      ��ǰ��̨�ɼ���ʶ��" & mintTagNow & "  ��ǰ��ʶͼ������" & mAfterCaptureInf.lngAfterCurImageCount & "        "
End Sub


Private Function subSaveAfterCaptureImage(Optional iEncode As Integer = 0) As Boolean
'�����̨�ɼ�ͼ��
    Dim i As Long
    Dim dtNowTime As Date
    Dim strReceivedTime As String
    Dim ImgTmp As DicomImage
    Dim strAfterImgPath As String
    Dim strTime As String
    Dim objFolder As Folder
    
    subSaveAfterCaptureImage = False
    
    If dcmAfter.Images.Count <= 0 Then Exit Function
    
    dtNowTime = zlDatabase.Currentdate
    strReceivedTime = Format(dtNowTime, "yyyyMMdd")
    
    If mAfterCaptureInf.strAfterStudyUid = "" Then
        '���uidΪ�գ��򴴽��µ�UID
        mAfterCaptureInf.strAfterStudyUid = dcmglbUID.NewUID
        mAfterCaptureInf.strAfterSeriesUid = dcmglbUID.NewUID
        
        mAfterCaptureInf.intAfterTag = GetCaptureTag
    End If


    strAfterImgPath = mstrAfterDir & "\" & strReceivedTime & "\" & "��ʶ" & mAfterCaptureInf.intAfterTag & "-" & mAfterCaptureInf.strAfterStudyUid & "\" & mAfterCaptureInf.strAfterSeriesUid & "\"
    
    MkLocalDir strAfterImgPath
    
    Set objFolder = mobjFile.GetFolder(strAfterImgPath)
    
    strTime = Format(objFolder.DateCreated, "YYYY-MM-DD HH:MM:SS")
    
    mrsImageCache.AddNew
    mrsImageCache!���� = "��ʶ" & Lpad(mintTagNow, 3, "0")
    mrsImageCache!���uid = mAfterCaptureInf.strAfterStudyUid
    mrsImageCache!����Uid = mAfterCaptureInf.strAfterSeriesUid
    mrsImageCache!������� = strTime
    mrsImageCache!·�� = strAfterImgPath
    mrsImageCache.Update
    

    For i = 1 To dcmAfter.Images.Count
        Set ImgTmp = dcmAfter.Images(i)
        
        ImgTmp.StudyUID = mAfterCaptureInf.strAfterStudyUid
        ImgTmp.SeriesUID = mAfterCaptureInf.strAfterSeriesUid
        
        If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
            '����ͼ�񵽻���Ŀ¼
            Select Case iEncode
                Case 1          'Run-Length Encoding�г�ѹ��
                    ImgTmp.WriteFile strAfterImgPath & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.5"
                Case 2          '����������ԭͼ��ѹ����ʽ
                    ImgTmp.WriteFile strAfterImgPath & ImgTmp.InstanceUID, True
                Case Else       'Lossless JPEG encoding JPEG����ѹ��
                    ImgTmp.WriteFile strAfterImgPath & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.4.70"
            End Select
            
            '�洢Ϊ����ͼ��
            'ImgTmp.FileExport strAfterImgPath & ImgTmp.InstanceUID & ".jpg", "JPG", 80
        End If
        
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            '��¼���Ƶ���Ӧ��Ŀ¼�У�������������
            Call MoveFile(ImgTmp.Tag.VideoFile, strAfterImgPath & ImgTmp.InstanceUID & ".avi")
        ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
            '����Ƶ�ļ����Ƶ���Ӧ��Ŀ¼�У�������������
            Call MoveFile(ImgTmp.Tag.VideoFile, strAfterImgPath & ImgTmp.InstanceUID & ".wav")
        End If
        
        mAfterCaptureInf.lngAfterCurImageCount = mAfterCaptureInf.lngAfterCurImageCount + 1
    Next i
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        Call frmCaptureHint.ShowCaptureHint( _
            IIf(gobjCapturePar.IsWindowHint, strAfterImgPath & ImgTmp.InstanceUID, ""), _
            gobjCapturePar.IsSountHint, hpRB, Me)
            
    End If
    
    Call DoStateChange(vetAfterUpdateImg, 0, 0, mAfterCaptureInf.strAfterStudyUid)

    'Call AddDCMImg(strAfterImgPath)
    Dim dcmTag As New zl9PacsControl.clsImageTagInf
    Dim strPathTmp As String
    Set ImgTmp.Tag = dcmTag
    
    strPathTmp = strAfterImgPath & "\" & ImgTmp.InstanceUID
    strPathTmp = Replace(strPathTmp, "/", "\")
    strPathTmp = Replace(strPathTmp, "\\", "\")
    ImgTmp.Tag.FilePath = strPathTmp
    
    Call AddDCMImg(ImgTmp)
    
    subSaveAfterCaptureImage = True
End Function

Private Function ShowMessage(ByVal lngUpLoadResult As Long) As Boolean
'�ļ��ϴ��ɹ�������ʾ,�ļ��ϴ��ɹ�����true����֮����false
    ShowMessage = False
    
    If lngUpLoadResult = 0 Then '�ϴ��ɹ�����������
        ShowMessage = True
    ElseIf lngUpLoadResult = 1 Then 'FTP����ʧ��
        MsgboxCus "FTP����ʧ�ܣ��ļ��޷����棬�����������á�", vbInformation, G_STR_HINT_TITLE
    Else                      '�ļ��ϴ�ʧ��
        MsgboxCus "�ļ��ϴ�ʧ�ܣ������������粻�ȶ���ɡ�", vbInformation, G_STR_HINT_TITLE
    End If
End Function

Private Sub subSaveImage(ByVal lngAdviceId As Long, ByVal strStudyUid As String, Optional iEncode As Integer = 0)
'------------------------------------------------
'���ܣ������һ������ͼ���浽���ݿ���
'������iEncode����ѹ����ʽ��1��Run-Length Encoding�г�ѹ����2������������ԭͼ��ѹ����ʽ��������Lossless JPEG encoding JPEG����ѹ��
'���أ���
'------------------------------------------------
    Dim ImgTmp As New DicomImage
    
    '��ȡ���һ������ͼ
    With ucPreview.ImgViewer
        If .Images.Count <= 0 Then Exit Sub
        Set ImgTmp = .Images(.Images.Count)
    End With
    
    Call SaveImage(ImgTmp, lngAdviceId, strStudyUid, iEncode)
End Sub

Private Sub SaveImage(ImgTmp As DicomImage, ByVal lngAdviceId As Long, ByVal strStudyUid As String, Optional iEncode As Integer = 0, Optional blnSave As Boolean)
    Dim rsTmp As New ADODB.Recordset
    
    
    Dim dtReceived As String
    Dim blnFirstImage As String     '�Ƿ񱾴μ��ĵ�һ��ͼ��
    Dim nowTime As Date
    Dim strReportImages As String
    Dim arrSql() As Variant
    Dim blnInTrans As Boolean       '�����ﴦ�������
    Dim i As Integer
    Dim lngSendNo As Long
    Dim strSQL As String
    Dim imgTag As clsImageTagInf
    Dim blnContinue As Boolean
    
    '�ȱ���FTPͼ��
    '��ȡ��������
    strSQL = "select ����, �Ա�, ����, ���UID ,��������,����ͼ��,���ͺ� from Ӱ�����¼ where ҽ��ID = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngAdviceId)
    
    nowTime = zlDatabase.Currentdate
    
    If IsNull(rsTmp("���UID")) Then
        dtReceived = Format(nowTime, "yyyyMMdd")
        blnFirstImage = True
    Else
        dtReceived = Format(rsTmp("��������"), "yyyyMMdd")
        blnFirstImage = False
    End If
    
    '��������Ŀ¼
    MkLocalDir mstrBufferDir & dtReceived & "\" & strStudyUid & "\"
    lngSendNo = rsTmp!���ͺ�
    
    Set imgTag = ImgTmp.Tag
    
    If Not blnSave Then
        blnContinue = imgTag.Tag <> VIDEOTAG And imgTag.Tag <> AUDIOTAG
    Else
        blnContinue = True
    End If
    
    If blnContinue Then
    
        strReportImages = Nvl(rsTmp("����ͼ��"))
        
        '��鱨��ͼ��ĳ��ȣ��������4000���ֽڣ�����ʾ�޷�����ͼ��
        If Len(strReportImages & " ;" & ImgTmp.InstanceUID & ".jpg") >= 4000 Then
            MsgboxCus "ͼ�������������ޣ�����ɾ������ͼ����ټ����ɼ�ͼ��", vbInformation, G_STR_HINT_TITLE
            Call ucPreview.DeleteImage(ucPreview.CurImageCount)
            Exit Sub
        End If
        
        '����ͼ�񵽻���Ŀ¼
        Select Case iEncode
            Case 1          'Run-Length Encoding�г�ѹ��
                ImgTmp.WriteFile mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.5"
            Case 2          '����������ԭͼ��ѹ����ʽ
                ImgTmp.WriteFile mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID, True
            Case Else       'Lossless JPEG encoding JPEG����ѹ��
                ImgTmp.WriteFile mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.4.70"
        End Select
        
        If gtFileLoadType <> Service Then ImgTmp.FileExport mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID & ".jpg", "JPG", 80
    End If
    
BUGEX "subSaveImage gtFileLoadType = " & gtFileLoadType

    If gtFileLoadType = Service Then
        If Not SaveImageWithService(lngAdviceId, strStudyUid, dtReceived, rsTmp, ImgTmp) Then Exit Sub
    Else
        If Not SaveImageWithNormal(lngAdviceId, strStudyUid, dtReceived, ImgTmp) Then Exit Sub
    End If
    
    'ͼ��洢�ɹ��󣬴洢���ݿ���Ϣ
    On Error GoTo DBError
    arrSql = Array()
    
    If blnFirstImage Then
        strSQL = "ZL_Ӱ�����¼_SET(" & lngAdviceId & "," & lngSendNo & ",'" & _
            strStudyUid & "',null," & _
            "to_Date('" & Format(nowTime, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'),'" & mobjFtp.strDeviceId & "')"
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = strSQL
    End If
    
    strSQL = "Select ����UID From Ӱ��������  Where ����UID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PACSͼ�񱣴�", CStr(ImgTmp.SeriesUID))
    
    '�����µļ������,���Ϊ¼��������µ�����
    If rsTmp.EOF Or ImgTmp.Tag.Tag = VIDEOTAG Or ImgTmp.Tag.Tag = AUDIOTAG Then
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            strSQL = "ZL_Ӱ������_INSERT('" & strStudyUid & "','" & ImgTmp.SeriesUID & "','��Ƶ¼��',0)"
        ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
            strSQL = "ZL_Ӱ������_INSERT('" & strStudyUid & "','" & ImgTmp.SeriesUID & "','��Ƶ����',0)"
        Else
            strSQL = "ZL_Ӱ������_INSERT('" & strStudyUid & "','" & ImgTmp.SeriesUID & "','" & ImgTmp.SeriesDescription & "',0)"
        End If
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = strSQL
    End If
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        '�����µ�ͼ���¼
        strSQL = "ZL_Ӱ��ͼ��_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "',NULL,0, null, sysdate)"
    Else
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            '�����µ���Ƶ��¼
            strSQL = "ZL_Ӱ��ͼ��_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "',Null,0" & _
            ",null,sysdate,null,null,null,null,null,null,null,null,null," & VIDEOTAG & ",'" & mstrEncoderName & "'," & ImgTmp.Tag.RecordTimeLen & ")"
        Else
            '�����µ���Ƶ��¼
            strSQL = "ZL_Ӱ��ͼ��_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "',Null,0" & _
            ",null,sysdate,null,null,null,null,null,null,null,null,null," & AUDIOTAG & ",''," & ImgTmp.Tag.RecordTimeLen & ")"
        End If
    End If
    
    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = strSQL
    
    gcnVideoOracle.BeginTrans        '----------����ͼ��
    blnInTrans = True
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "����ͼ��")
    Next i
    
    gcnVideoOracle.CommitTrans
    blnInTrans = False
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        Call frmCaptureHint.ShowCaptureHint( _
            IIf(gobjCapturePar.IsWindowHint, mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID, ""), _
            gobjCapturePar.IsSountHint, hpRB, Me)
    End If

    If mintCaptureFlag = 1 Or mintCaptureFlag = 4 Or mintCaptureFlag = 5 Then
        If ucPreview.CurImageCount = 1 Then
            Call DoStateChange(vetCaptureFirstImg, lngAdviceId, lngSendNo, strStudyUid)
        Else
            '����ͼ��
            Call DoStateChange(vetUpdateImg, lngAdviceId, lngSendNo, strStudyUid)
        End If
    ElseIf mintCaptureFlag = 2 Then
        '����Ӱ����״̬������ɼ���һ��ͼ����ԭ����״̬���ѱ��������޸ĳ��Ѽ��
        If ucPreview.ImgViewer.Images.Count = 1 Then
            If mlngStudyState < 3 Then
                strSQL = "Zl_Ӱ����_State(" & lngAdviceId & "," & lngSendNo & ",3,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & glngDepartId & ")"
                zlDatabase.ExecuteProcedure strSQL, "�ɼ���һ��ͼ��"
                mlngStudyState = 3
            End If
        End If
        
        If ucPreview.ImgViewer.Images.Count = 1 Then
            '�ɼ���һ��ͼ��
            Call DoStateChange(vetCaptureFirstImg, lngAdviceId, lngSendNo, strStudyUid)
        Else
            '����ͼ��
            Call DoStateChange(vetUpdateImg, lngAdviceId, lngSendNo, strStudyUid)
        End If
    ElseIf mintCaptureFlag = 3 Then
        '����Ӱ����״̬������ɼ���һ��ͼ����ԭ����״̬���ѱ��������޸ĳ��Ѽ��
        If ucPreview.CurImageCount = 1 Then
            If mlngStudyState < 3 Then
                strSQL = "Zl_Ӱ����_State(" & lngAdviceId & "," & lngSendNo & ",3,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & glngDepartId & ")"
                zlDatabase.ExecuteProcedure strSQL, "�ɼ���һ��ͼ��"
                mlngStudyState = 3
            End If
        End If
        
        If ucPreview.CurImageCount = 1 Then
            Call DoStateChange(vetCaptureFirstImg, lngAdviceId, lngSendNo, strStudyUid)
        Else
            '����ͼ��
            Call DoStateChange(vetUpdateImg, lngAdviceId, lngSendNo, strStudyUid)
        End If
    End If
    Exit Sub
DBError:
    '������������ݿ����������ɾ�����ɼ���ͼ��
    If blnInTrans = True Then gcnVideoOracle.RollbackTrans
    err.Raise err.Number, "���ͼ�񱣴�"
    Call ucPreview.DeleteImage(ucPreview.CurImageCount)
End Sub

Private Function SaveImageWithNormal(ByVal lngAdviceId As Long, ByVal strStudyUid As String, ByVal dtReceived As String, ImgTmp As DicomImage) As Boolean
'ʹ����ԭʼ�ķ�ʽ�ϴ�ͼ��
    Dim lngResult As Long
    Dim lngUpLoadResult As Long '�ɹ�:0��FTP����ʧ��:1���ϴ��ļ�ʧ��:2
    
    lngResult = mobjFtpConnection.FuncFtpConnect(mobjFtp.strFTPIP, mobjFtp.strFTPUser, mobjFtp.strFTPPwd)
    If lngResult = 0 Then
        'FTP����ʧ�ܣ���ʾ���󣬲�ɾ������ͼ�е�ͼ��
        MsgboxCus "FTP����ʧ�ܣ�ͼ���޷����棬�����������á�", vbInformation, G_STR_HINT_TITLE
        Call ucPreview.DeleteImage(ucPreview.CurImageCount)
    
        Exit Function
    End If
    
    If Val(mobjBakFtp.strDeviceId) > 0 Then
        lngResult = mobjBakFtpConnection.FuncFtpConnect(mobjBakFtp.strFTPIP, mobjBakFtp.strFTPUser, mobjBakFtp.strFTPPwd)
        If lngResult = 0 Then
            mobjBakFtp.strDeviceId = ""
            MsgboxCus "����ftp�豸����ʧ�ܣ��ɼ���ͼ�񽫲��ܽ��б��ݲ��������豸���������̹����еı����豸���á�", vbInformation, G_STR_HINT_TITLE
        End If
    End If
    
    If ImgTmp.Tag.Tag <> VIDEOTAG And ImgTmp.Tag.Tag <> AUDIOTAG Then
        '����dicomͼ��
        lngUpLoadResult = WriteToURL(mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID, _
                                    mobjFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID, True)
        
        If Not ShowMessage(lngUpLoadResult) Then Exit Function
            
        lngUpLoadResult = WriteToURL(mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID & ".jpg", _
                                    mobjFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID & ".jpg")
        
        If Not ShowMessage(lngUpLoadResult) Then Exit Function
        
        '���ݵ�ǰ�ɼ���ͼ��
        If mobjBakFtpConnection.hConnection <> 0 Then
            lngUpLoadResult = BakImgToURL(mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID, _
                                        mobjBakFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID)
            
            If lngUpLoadResult <> 0 Then
                MsgboxCus "�ļ�����ʧ�ܣ������������粻�ȶ���ɡ�", vbInformation, G_STR_HINT_TITLE
            End If
        End If
    Else
        '����¼��
        lngUpLoadResult = WriteToURL(ImgTmp.Tag.VideoFile, _
                                    mobjFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID)
        
        If Not ShowMessage(lngUpLoadResult) Then Exit Function
        
        '����¼��
        If mobjBakFtpConnection.hConnection <> 0 Then
            lngUpLoadResult = BakImgToURL(ImgTmp.Tag.VideoFile, _
                                        mobjBakFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID)
            
            If lngUpLoadResult <> 0 Then
                MsgboxCus "�ļ�����ʧ�ܣ������������粻�ȶ���ɡ�", vbInformation, G_STR_HINT_TITLE
            End If
        End If
        
        If ImgTmp.Tag.Tag = VIDEOTAG Then
            '��¼���Ƶ���Ӧ��Ŀ¼�У�������������
            Call FileCopy(ImgTmp.Tag.VideoFile, mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID & ".avi")
            Call Kill(ImgTmp.Tag.VideoFile)
        
            ImgTmp.Tag.VideoFile = mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID & ".avi"
        ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
            '����Ƶ�ļ����Ƶ���Ӧ��Ŀ¼�У�������������
            Call FileCopy(ImgTmp.Tag.VideoFile, mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID & ".wav")
            Call Kill(ImgTmp.Tag.VideoFile)
        
            ImgTmp.Tag.VideoFile = mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID & ".wav"
        End If
    End If
    
    mobjFtpConnection.FuncFtpDisConnect
    If mobjBakFtpConnection.hConnection <> 0 Then mobjBakFtpConnection.FuncFtpDisConnect
    
    SaveImageWithNormal = True
End Function

Private Function SaveImageWithService(ByVal lngAdviceId As Long, ByVal strStudyUid As String, ByVal dtReceived As String, rsTmp As ADODB.Recordset, ImgTmp As DicomImage) As Boolean
'ʹ��Service�����̨�ϴ�
    Dim strSQL As String
    Dim fileMsg As TransferFileMsg
    
    If ImgTmp.Tag.Tag = VIDEOTAG Then
        '��¼���ƶ�����Ӧ��Ŀ¼�У�������������
        Name ImgTmp.Tag.VideoFile As mstrBufferDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID & ".avi"
    
        ImgTmp.Tag.VideoFile = mstrBufferDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID & ".avi"
    ElseIf ImgTmp.Tag.Tag = AUDIOTAG Then
        '����Ƶ�ļ��ƶ�����Ӧ��Ŀ¼�У�������������
        Name ImgTmp.Tag.VideoFile As mstrBufferDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID & ".wav"
    
        ImgTmp.Tag.VideoFile = mstrBufferDir & dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID & ".wav"
    End If
    
'BUGEX "strFTPIP = " & mobjFtp.strFTPIP & " strFTPUser = " & mobjFtp.strFTPUser & " strFTPPwd = " & mobjFtp.strFTPPwd
'BUGEX "strFtpDir = " & mobjFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/"
'BUGEX "lngAdviceId = " & lngAdviceId
'BUGEX "LOCALDIR = " & mstrBufferDir & dtReceived & "/" & strStudyUid & "/" & " FILENAME = " & ImgTmp.InstanceUID
'BUGEX "strBakFTPIP = " & mobjBakFtp.strFTPIP & " strBakFTPUser = " & mobjBakFtp.strFTPUser & " strBakFTPPwd = " & mobjBakFtp.strFTPPwd
'BUGEX "strFtpDir = " & mobjBakFtp.strFtpDir & dtReceived & "/" & strStudyUid & "/"
    
    With fileMsg
        fileMsg.strAdviceId = lngAdviceId
        fileMsg.strName = Nvl(rsTmp("����"))
        fileMsg.strSex = Nvl(rsTmp("�Ա�"))
        fileMsg.strAge = Nvl(rsTmp("����"))
        
        fileMsg.ftpInfo.strDeviceId = mobjFtp.strDeviceId
        fileMsg.ftpInfo.strFtpDir = mobjFtp.strFtpDir
        fileMsg.ftpInfo.strFTPIP = mobjFtp.strFTPIP
        fileMsg.ftpInfo.strFTPPwd = mobjFtp.strFTPPwd
        fileMsg.ftpInfo.strFTPUser = mobjFtp.strFTPUser
        fileMsg.ftpInfo.strSDDir = mobjFtp.strSDDir
        fileMsg.ftpInfo.strSDPswd = mobjFtp.strSDPswd
        fileMsg.ftpInfo.strSDUser = mobjFtp.strSDUser
        
        fileMsg.bakFtpInfo.strDeviceId = mobjBakFtp.strDeviceId
        fileMsg.bakFtpInfo.strFtpDir = mobjBakFtp.strFtpDir
        fileMsg.bakFtpInfo.strFTPIP = mobjBakFtp.strFTPIP
        fileMsg.bakFtpInfo.strFTPPwd = mobjBakFtp.strFTPPwd
        fileMsg.bakFtpInfo.strFTPUser = mobjBakFtp.strFTPUser
        fileMsg.bakFtpInfo.strSDDir = mobjBakFtp.strSDDir
        fileMsg.bakFtpInfo.strSDPswd = mobjBakFtp.strSDPswd
        fileMsg.bakFtpInfo.strSDUser = mobjBakFtp.strSDUser
        
        fileMsg.strLocalDir = mstrBufferDir & dtReceived & "\" & strStudyUid & "\" & ImgTmp.InstanceUID
        fileMsg.strFileName = ImgTmp.InstanceUID
        fileMsg.strSubDir = dtReceived & "/" & strStudyUid & "/" & ImgTmp.InstanceUID
        fileMsg.strMediaType = ImgTmp.Tag.Tag
    End With

    If Not SendDataToservice("����ͼ", COMMAND_CAPIMG_UPLOAD, "ͼ��ɼ�", fileMsg) Then
BUGEX "ͼ����Ϣδ�ɹ�������������"
        MsgboxEx Me.hwnd, "��ͼ�����ݷ��������������ʱ����������ZLPacsServerCenter����δ���ã�" & vbCrLf & _
                          "���ݽ���ʱ���浽���أ����´δ򿪷���ʱ�����Զ��ϴ���", vbOKOnly, G_STR_HINT_TITLE
            
        SaveImageWithService = True
        Exit Function
    Else
BUGEX "ͼ����Ϣ�ɹ�������������"
        SaveImageWithService = True
    End If
End Function

Private Function WriteToURL(ByVal SrcFileName As String, ByVal DestFileName As String, Optional ByVal blnCheck As Boolean = False) As Long
'------------------------------------------------
'���ܣ��������ļ����浽Զ��������
'������SrcFileName--�����ļ�����DestFileName����Զ���ļ���
'      blnCheck -- �ϴ��ļ���Ӧ���м�飨�Ƚ�ftp�Ϻͱ��ص��ļ���С�Ƿ���ͬ��
'���أ��ɹ�����0������ʧ�ܷ���1���ϴ��ļ�ʧ�ܷ���2
'-----------------------------------------------
'���ܣ�
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strPath As String, StrMessage As String
    Dim lngFtpFileSzie As Long, lngDestFileSize As Long
    
    '��FTP�д���Ŀ¼
    strPath = objFileSystem.GetParentFolderName(DestFileName)
    mobjFtpConnection.FuncFtpMkDir "/", strPath
    
ReUplcad:
    
    '��FTP�ϴ��ļ�
    WriteToURL = mobjFtpConnection.FuncUploadFile(strPath, Replace(SrcFileName, "/", "\"), objFileSystem.GetFileName(DestFileName))
    
    '�ϴ��ɹ���Ա�һ�´�С���ж��Ƿ������ϴ�
    If blnCheck And WriteToURL = 0 And mblnCompareSize Then
        lngDestFileSize = objFileSystem.GetFile(SrcFileName).Size
        lngFtpFileSzie = mobjFtpConnection.FuncFtpGetFileSize(strPath, objFileSystem.GetFileName(DestFileName))
        
        If lngFtpFileSzie < lngDestFileSize Then
            StrMessage = "�ϴ�����ļ���С[" & lngFtpFileSzie & "]��ԭ�ļ���С[" & lngDestFileSize & "]��һ��" & vbCrLf & _
                         "ԭ�ļ���" & SrcFileName & vbCrLf & _
                         "FTP�ļ���" & strPath & objFileSystem.GetFileName(DestFileName) & vbCrLf & _
                         "�Ƿ���Ҫ�����ϴ���"
            
            If MsgBox(StrMessage, vbQuestion + vbYesNo, "��ʾ") = vbYes Then
                GoTo ReUplcad
            End If
        End If
    End If
End Function


Private Function BakImgToURL(ByVal SrcFileName As String, ByVal DestFileName As String) As Long
'------------------------------------------------
'���ܣ�����ͼ��Զ��������
'������SrcFileName--�����ļ�����DestFileName����Զ���ļ���
'���أ��ɹ�����0������ʧ�ܷ���1���ϴ��ļ�ʧ�ܷ���2
'-----------------------------------------------
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strPath As String
    
    If mobjBakFtpConnection.hConnection = 0 Then Exit Function
    
    '��FTP�д���Ŀ¼
    strPath = objFileSystem.GetParentFolderName(DestFileName)
    mobjBakFtpConnection.FuncFtpMkDir "/", strPath
    
    '��FTP�ϴ��ļ�
    BakImgToURL = mobjBakFtpConnection.FuncUploadFile(strPath, Replace(SrcFileName, "/", "\"), objFileSystem.GetFileName(DestFileName))
End Function


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
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    
    BUGEX "InitCommandBars:Set CommandBar Icon"
    
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons 'imgPublic.Icons '
    
    BUGEX "InitCommandBars:1"
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 32, 32
    End With
    
    BUGEX "InitCommandBars:2"
    
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    BUGEX "InitCommandBars:3"
    
    '�Ƿ���ʾ��������
    mblnShowProcessBar = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "��ʾ��������", "True")
    
    BUGEX "InitCommandBars:4"
    
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
        
        '���ú�̨�ɼ�
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Capture, "��̨�ɼ�"): cbrControl.ToolTipText = "��̨�ɼ�"
            cbrControl.IconId = 10020
        
        '�ڷ�TWAIN�ɼ�ģʽ������£�����ʾ�ð�ť
        'If Not GetIsTwainCaptureWay Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Record, "¼��"): cbrControl.ToolTipText = "��ʼ¼��"
                cbrControl.Enabled = True
                
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Record, "��̨¼��"): cbrControl.ToolTipText = "��̨¼��"
                cbrControl.IconId = 10021
            
            
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
            
            
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_OpenStudyList, "�򿪼��"): cbrControl.ToolTipText = "�򿪼��"
            cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_StudySyncState, "�������"): cbrControl.ToolTipText = "�������"
            cbrControl.IconId = 10012
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Tag, "���±�ʶ"): cbrControl.ToolTipText = "���±�ʶ"
            cbrControl.IconId = 10022
        
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
'        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "�߼�"): cbrControl.ToolTipText = "�߼�����"
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIcon
        cbrControl.Category = "����"
        cbrControl.Enabled = False
    Next
    cbrToolBar.Visible = mblnShowProcessBar
End Sub


Private Sub ShowFrameSelectImagePopup()
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
    
    '�ɼ��ü�ͼ��
    Set imgResult = CutImage(dcmView.Images(1))
    If imgResult Is Nothing Then Exit Sub
    
    '��imgResultһ��Ψһ�� InstanceUID
    imgResult.InstanceUID = dcmglbUID.NewUID
    
    '�ѽ��ͼ���뵽viewer�в��ұ���
    '����ͼ���DICOM����
    subWriteDicomPara imgResult, mlngAdviceId
    
    Dim dcmTag As New clsImageTagInf
    dcmTag.Tag = imgTag
    
    Set imgResult.Tag = dcmTag
    
    mintCaptureFlag = 1
    
    '��ͼ����뵽����ͼ��
    subInsert2Mini imgResult
    
    '����ͼ�񣬲�����ͼ��洢�¼�
    Call subSaveImage(mlngAdviceId, mcurStudyInf.strStudyUid)
End Sub


Private Sub ucCapHook_OnKeyBoardLHook(ByVal lngVkCode As Long, ByVal lngScanCode As Long, ByVal lngFlags As Long)
On Error GoTo errHandle
    Select Case lngVkCode
        Case 66
            '�жϼ��̰����Ƿ��ɿ���Ϊ0��ʾ���¼���
            If lngScanCode = 128 Then
                'ִ�п�ݲɼ�
'                Call CaptureImage

                If timerHook.Enabled Then Exit Sub
                timerHook.Enabled = True
            End If
    End Select
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub ucPreview_OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
'    If ucPreview.ImgViewer.Images.Count <= 0 Then Exit Sub
'    If Button = 1 Then ucPreview.ImgChecked(ucPreview.SelectIndex) = Not ucPreview.ImgChecked(ucPreview.SelectIndex)
End Sub

Private Sub ucPreview_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucPreview.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 2 Then Call ShowPopupImage(1)
End Sub

Private Sub ucCacheImage_OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
'    If ucCacheImage.ImgViewer.Images.Count <= 0 Then Exit Sub
'    If Button = 1 Then ucCacheImage.ImgChecked(ucCacheImage.SelectIndex) = Not ucCacheImage.ImgChecked(ucCacheImage.SelectIndex)
End Sub

Private Sub ucCacheImage_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucCacheImage.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 2 Then Call ShowPopupImage(2)
End Sub

Private Sub ShowPopupImage(ByVal intType As Integer)
'------------------------------------------------
'���ܣ���������Ҽ������˵�
'intType:0--����ͼ��1--����ͼ��2--����ͼ
'------------------------------------------------
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrToolPopup As CommandBarPopup
    
    If intType <> 2 Then
        If ucPreview.CurImageCount < 1 Then Exit Sub
    End If
    
    '����Ҽ������˵�
    Set cbrToolBar = cbrMain.Add("����Ҽ�", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        If intType = 1 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_SplitPage, "��ҳ����")
            
            If glngModule = 1291 And gobjCapturePar.IsUseAfterCapture Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Cap_ImgExport, "���͵���̨")
                cbrControl.BeginGroup = True
            End If
            
            If mblnIsReported = False Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Cap_AddToReport, "���뱨��ͼ")
                cbrControl.BeginGroup = True
            End If
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_ReUpLoad, "�����ϴ�")
                cbrControl.BeginGroup = True
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_ReDownLoad, "��������")
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_SplitPage, "��ҳ����")
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_ImgImport, "���͵����")
                cbrControl.BeginGroup = True
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DelCache, "ɾ��")
                cbrControl.BeginGroup = True
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_RefreshCache, "ˢ��")
        End If
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Private Sub ucPreview_OnClick(ByVal lngSelectedIndex As Long)

    mCorpSize.X = 0
    mCorpSize.Y = 0
    
    '��ѡ��ͼ����ʾ���
    If lngSelectedIndex <> 0 Then
        
        Call PreviewThumbnail(lngSelectedIndex)

        '������Ƶ�ĵ�ǰ��ʾ״̬
        Call ConfigVideoShowState(False)
    End If
    
    '�ָ���ǰ�ؼ����㣬�Ա��ܹ�����ͼ��
    ucPreview.SetFocus
End Sub


Private Sub PreviewThumbnail(ByVal lngImgIndex As Long)
'Ԥ������ͼ
    Dim dblTempZoom As Double
    Dim img As DicomImage
    Dim i As Integer
    
    '����ѡ��ͼ��װ�ص�dcmView��
    dcmView.Images.Clear
    
    If lngImgIndex <= 0 Then Exit Sub
    dcmView.Images.Add ucPreview.ImgViewer.Images(lngImgIndex)
    
    Set img = dcmView.Images(1)
    
    'ȥ���߿�
    For i = 1 To img.Labels.Count
        If img.Labels(i).Tag = "SELECT" Or img.Labels(i).Tag = "BORDER" Or img.Labels(i).Tag = "FAILD" Then
            img.Labels(i).Visible = False
        End If
    Next
    
    '��ʾdcmView������picVideo
    dcmView.CurrentImage.BorderWidth = 0
    
    dblTempZoom = dcmView.CurrentImage.ActualZoom
    dcmView.CurrentImage.StretchToFit = False
        
    '�жϵ����븡������ʱ�����ű��ʲ���С��0.1
    If dblTempZoom < 0.1 Then dblTempZoom = 0.1
                  
    Call subCenterZoom(Me, dcmView.CurrentImage, dcmView, dblTempZoom, mCorpSize)
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
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub ucPreview_OnReUpload()
On Error GoTo errHandle
    
    Call ReloadSelectedImg
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub ReloadSelectedImg()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim dtReceived As String
    Dim objSelectedImg As DicomImage
    Dim fileMsg As TransferFileMsg
    
    If ucPreview.IsFailedImg(ucPreview.SelectIndex) = True Then
        MsgboxEx Me.hwnd, "���ܶԸ�ͼ������ϴ�.", vbOKOnly, G_STR_HINT_TITLE
        Exit Sub
    End If
    
    '�����ϴ�ѡ����ļ�
    Set objSelectedImg = ucPreview.SelectImage
    
    strSQL = "select ����, �Ա�, ����, ���UID ,��������,����ͼ��,���ͺ� from Ӱ�����¼ where ҽ��ID = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, mlngAdviceId)
    
    If rsTmp.RecordCount <= 0 Or objSelectedImg Is Nothing Then Exit Sub
    
    If IsNull(rsTmp("���UID")) Then
        dtReceived = Format(zlDatabase.Currentdate, "yyyyMMdd")
    Else
        dtReceived = Format(rsTmp("��������"), "yyyyMMdd")
    End If
    
'BUGEX "strFTPIP = " & mobjFtp.strFTPIP & " strFTPUser = " & mobjFtp.strFTPUser & " strFTPPwd = " & mobjFtp.strFTPPwd
'BUGEX "strFtpDir = " & mobjFtp.strFtpDir & dtReceived & "/" & objSelectedImg.StudyUID & "/"
'BUGEX "mlngAdviceId = " & mlngAdviceId
'BUGEX "strBakFTPIP = " & mobjBakFtp.strFTPIP & " strBakFTPUser = " & mobjBakFtp.strFTPUser & " strBakFTPPwd = " & mobjBakFtp.strFTPPwd
'BUGEX "strFtpDir = " & mobjBakFtp.strFtpDir & dtReceived & "/" & objSelectedImg.StudyUID & "/"
'BUGEX "LOCALDIR = " & mstrBufferDir & dtReceived & "/" & objSelectedImg.StudyUID & "/" & " FILENAME = " & objSelectedImg.InstanceUID

    With fileMsg
        fileMsg.strAdviceId = mlngAdviceId
        fileMsg.strName = Nvl(rsTmp("����"))
        fileMsg.strSex = Nvl(rsTmp("�Ա�"))
        fileMsg.strAge = Nvl(rsTmp("����"))
        
        fileMsg.ftpInfo.strDeviceId = mobjFtp.strDeviceId
        fileMsg.ftpInfo.strFtpDir = mobjFtp.strFtpDir
        fileMsg.ftpInfo.strFTPIP = mobjFtp.strFTPIP
        fileMsg.ftpInfo.strFTPPwd = mobjFtp.strFTPPwd
        fileMsg.ftpInfo.strFTPUser = mobjFtp.strFTPUser
        fileMsg.ftpInfo.strSDDir = mobjFtp.strSDDir
        fileMsg.ftpInfo.strSDPswd = mobjFtp.strSDPswd
        fileMsg.ftpInfo.strSDUser = mobjFtp.strSDUser
        
        fileMsg.bakFtpInfo.strDeviceId = mobjBakFtp.strDeviceId
        fileMsg.bakFtpInfo.strFtpDir = mobjBakFtp.strFtpDir
        fileMsg.bakFtpInfo.strFTPIP = mobjBakFtp.strFTPIP
        fileMsg.bakFtpInfo.strFTPPwd = mobjBakFtp.strFTPPwd
        fileMsg.bakFtpInfo.strFTPUser = mobjBakFtp.strFTPUser
        fileMsg.bakFtpInfo.strSDDir = mobjBakFtp.strSDDir
        fileMsg.bakFtpInfo.strSDPswd = mobjBakFtp.strSDPswd
        fileMsg.bakFtpInfo.strSDUser = mobjBakFtp.strSDUser
        
        fileMsg.strLocalDir = mstrBufferDir & dtReceived & "\" & objSelectedImg.StudyUID & "\" & objSelectedImg.InstanceUID
        fileMsg.strFileName = objSelectedImg.InstanceUID
        fileMsg.strSubDir = dtReceived & "/" & objSelectedImg.StudyUID & "/" & objSelectedImg.InstanceUID
        fileMsg.strMediaType = objSelectedImg.Tag.Tag
    End With
    
    If Not SendDataToservice("����ͼ", COMMAND_CAPIMG_UPLOAD, "ͼ��ɼ�", fileMsg) Then
        MsgboxEx Me.hwnd, "��ͼ�����ݷ��������������ʱ����������ZLPacsServerCenter����δ���ã�" & vbCrLf & _
                          "���ݽ���ʱ���浽���أ����´δ򿪷���ʱ�����Զ��ϴ���", vbOKOnly, G_STR_HINT_TITLE
    Else
BUGEX "ͼ����Ϣ�ɹ�������������"
    End If
End Sub

Private Sub ucPreview_OnSaveImage(ByVal dcmImage As DicomObjects.DicomImage, ByVal lngImageType As Long)
    Select Case lngImageType
        Case 1
            If dcmImage.Tag.ReportImage = "" Then
                dcmImage.Tag.ReportImage = 0
                '��Ҫ��Ϊ���ͼ
                Call SaveImageToStady(dcmImage, mlngAdviceId)
                Call AddToReport(dcmImage)
                ucPreview.AfterSaveStudy dcmImage
            End If
        Case 2
            Call SaveImageToStady(dcmImage, mlngAdviceId)
    End Select
End Sub

Private Sub ucSplitter1_OnMoveEnd()
On Error Resume Next
    RaiseEvent OnControlResize(picCapture)
err.Clear
End Sub

Private Sub AddDCMImg(ByVal img As DicomImage)
'����ǵ��쵱ǰ��ʶ���������½�ͼƬ��ʾ�ؼ�������ͼƬ���������ֲɼ��ɹ�

On Error GoTo errH
    Dim curTime As Date


    If Not (InStr(cboCacheImage.Text, "��ʶ" & Lpad((mintTagNow), 3, "0")) > 0) Then Exit Sub '���ǵ�ǰ��ʶ�˳�

    curTime = zlDatabase.Currentdate
    If Format(DTPImg.value, "yyyymmdd") <> Format(curTime, "yyyymmdd") Then Exit Sub '���ǵ����˳�
    
    Call ucCacheImage.AddImage(img, img.Tag)
    
    Exit Sub
errH:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub ClearEmptyFolder(ByVal blNoReason As Boolean)
'blNoReason �Ƿ���������ִ�б����̣����ڹرճ����ʱ��ִ��
'��տ�Ŀ¼����Ӧ������������ǰѡ�е��ǵ������±�ʶ����ִ�д˲���
'�����ж�������ǰѡ���Ƿ��ǵ������±�ʶ
    Dim curTime As Date
    Dim strTime As String
    Dim objFolder1 As Folder, objFolder2 As Folder, objFolder3 As Folder
    Dim strType As String
    Dim strTag As String
    Dim i As Long
    On Error GoTo errH
   
    
    If blNoReason = False Then
    
        '����ǰ��ʾ��ͼ�񣬲����в���
        If Not mblIsNeedRefresh And ucCacheImage.CurImageCount > 0 Then Exit Sub
        
        curTime = zlDatabase.Currentdate
    
        '�ǵ��첢��ѡ�е��ǵ�ǰ��ʶ���Ͳ�������ղ���
        If Not mblIsNeedRefresh And (Format(DTPImg.value, "yyyymmdd") = Format(curTime, "yyyymmdd")) And (InStr(cboCacheImage.Text, "��ʶ" & Lpad((mintTagNow), 3, "0")) > 0) Then Exit Sub

    End If
    
    If mobjFile.FolderExists(mstrAfterDir) = False Then Exit Sub
    
    If mobjFile.GetFolder(mstrAfterDir).SubFolders.Count > 0 Then
        For Each objFolder1 In mobjFile.GetFolder(mstrAfterDir).SubFolders '''ʱ��
            If objFolder1.SubFolders.Count > 0 Then
                For Each objFolder2 In objFolder1.SubFolders '''���uid
                    If objFolder2.SubFolders.Count > 0 Then
                    
                        For Each objFolder3 In objFolder2.SubFolders '''����UID
                            If objFolder3.Files.Count = 0 Then
                                If blNoReason = True Then
                                    Call mobjFile.DeleteFolder(objFolder3.Path)
                                Else
                                    If Not (objFolder1.Name = Format(curTime, "yyyymmdd") And mintTagNow = GetTag(objFolder2.Name, strType)) Then
                                        Call mobjFile.DeleteFolder(objFolder3.Path)
                                    End If
                                End If
                                
                            End If
                        Next
                        
                        If objFolder2.SubFolders.Count = 0 Then Call mobjFile.DeleteFolder(objFolder2.Path)
                    Else
                        Call mobjFile.DeleteFolder(objFolder2.Path)
                    End If
                Next
                
                If objFolder1.SubFolders.Count = 0 Then Call mobjFile.DeleteFolder(objFolder1.Path)
            Else
                Call mobjFile.DeleteFolder(objFolder1.Path)
            End If
        Next
    End If
    Exit Sub
errH:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Function GetTodayTagMax(ByVal curDate As Date) As Integer
    '���㵱������ʶ
    Dim strDate As String
    Dim intTmp As Integer
    Dim strType As String
    Dim strStudyUid As String
    Dim objFolder2 As Folder, objFolder3 As Folder
    
    On Error GoTo errH
    
    mintTagMax = 1
    
    strDate = Format(curDate, "yyyymmdd")
    
    If mobjFile.FolderExists(mstrAfterDir) = False Then Exit Function
    
    If mobjFile.GetFolder(mstrAfterDir).SubFolders.Count > 0 Then
        For Each objFolder2 In mobjFile.GetFolder(mstrAfterDir).SubFolders
            If InStr(objFolder2.Name, strDate) > 0 Then                                 'ʱ��ѡ��
            
                If objFolder2.SubFolders.Count > 0 Then                                  'ʱ����Ƿ�����Ŀ¼
                
                    For Each objFolder3 In objFolder2.SubFolders                            '���UID��
                    
                        If objFolder3.SubFolders.Count > 0 Then

                            strStudyUid = GetStudyUIDFromFolderName(objFolder3.Name)
                            
                            intTmp = GetTag(objFolder3.Name, strType)
                            If intTmp > mintTagMax Then mintTagMax = intTmp
                                             
                        End If

                    Next
                    
                End If 'ʱ����Ƿ�����Ŀ¼
                
            End If 'ʱ��ѡ��
        Next
    End If
    
    GetTodayTagMax = mintTagMax
    
    Exit Function
errH:
    BUGEX "GetTodayTagMax output= -1"
    GetTodayTagMax = -1
End Function

Private Function JudgeExistFolderOne() As Boolean
'�ж��Ƿ���ڱ�ʶ1��Ŀ¼  t- ���� f -������ Ĭ�ϲ�����
    On Error GoTo errH
    Dim objFolder2 As Folder, objFolder3 As Folder, objFolder4 As Folder
    Dim strDate As String

    strDate = Format(zlDatabase.Currentdate, "yyyymmdd")
    JudgeExistFolderOne = False
    ''''''''''''''
    If mobjFile.FolderExists(mstrAfterDir) Then
        For Each objFolder2 In mobjFile.GetFolder(mstrAfterDir).SubFolders
            If objFolder2.Name = strDate Then 'ʱ��ֻ�жϵ���

                For Each objFolder3 In objFolder2.SubFolders   '����
                
                    If InStr(objFolder3.Name, "��ʶ1") > 0 Then
                        JudgeExistFolderOne = True
                        Exit Function
                    End If
                Next
                
            End If
        Next
    End If
    
    JudgeExistFolderOne = False
    
    Exit Function
errH:
    BUGEX "JudgeExistFolderOne output= False"
    JudgeExistFolderOne = False
End Function

Private Function GetCountOfCapPic(ByVal intTag As Integer) As Integer '11111
'���㵱ǰ��ʶ��ͼ������,inttag��ʶ
    On Error GoTo errH
    Dim objFolder2 As Folder, objFolder3 As Folder, objFolder4 As Folder
    Dim strDate As String
    

    strDate = Format(zlDatabase.Currentdate, "yyyymmdd")
    GetCountOfCapPic = 0

    If mobjFile.FolderExists(mstrAfterDir) Then
        For Each objFolder2 In mobjFile.GetFolder(mstrAfterDir).SubFolders
        
            If objFolder2.Name = strDate Then 'ʱ��ֻ�жϵ���

                For Each objFolder3 In objFolder2.SubFolders   '����
                
                    If InStr(objFolder3.Name, "���" & intTag) > 0 Or InStr(objFolder3.Name, "��ʶ" & intTag) > 0 Then '
                    'If InStr(objFolder3.Name, "��ʶ" & inttag) > 0 Then  '
                        For Each objFolder4 In objFolder3.SubFolders   '���в�
                            GetCountOfCapPic = objFolder4.Files.Count
                            Exit Function
                        Next
                        
                    End If
                Next
                      
            End If
        Next
    End If

    Exit Function
errH:
    BUGEX "GetCountOfCapPic output= -1"
    GetCountOfCapPic = -1
End Function

Public Sub setFontSize(ByVal intsize As Integer)
'�����������ֺ�
     
    mintFontSize = intsize
    Call picCacheImage_Resize
    
    DTPImg.Font.Size = intsize
    cboCacheImage.Font.Size = intsize
    dkpReportImage.RedrawPanes

End Sub

Private Function DeleteManyImages(ByVal strArrImgs As String) As Boolean
'ɾ�����ͼ��(ɾ�����ݿ�����),strArrImgs:��ɾ����ͼ��uid
    On Error GoTo errH
    Dim i As Integer
    Dim strSQL As String
    Dim iNet As New clsFtp
    Dim rsTemp As ADODB.Recordset
    Dim strFTPIP As String
    Dim strFtpPass As String
    Dim strFTPUser As String
    Dim strRoot As String
    Dim intDeviceUsed As Integer
    Dim lngResult As Long
    Dim strImageUID As String
    
    DeleteManyImages = False
    strSQL = "Select a.ҽ��ID,a.���ͺ�,c.ͼ��UID, " & _
            " Decode(a.��������,Null,'',to_Char(a.��������,'YYYYMMDD')||'/')||a.���UID As ͼ��Ŀ¼, " & _
            "D.FTP�û��� As User1,D.FTP���� As Pwd1,D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1,d.�豸�� as �豸��1," & _
            "E.FTP�û��� As User2,E.FTP���� As Pwd2,E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2,e.�豸�� as �豸��2 " & _
            "From Ӱ�����¼ a,Ӱ�������� b,Ӱ����ͼ�� c,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E, " & _
            "Table(Cast(f_Str2list([1],';') As zlTools.t_Strlist)) F " & _
            "Where a.���UID=b.���UID And b.����UID=c.����UID And c.ͼ��UID = F.Column_Value " & _
            "And a.λ��һ=D.�豸��(+) And a.λ�ö�=E.�豸��(+)"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "PACSɾ�����ͼ��", strArrImgs)
    If rsTemp.EOF = True Then
        MsgboxCus "û���ҵ�����ɾ����ͼ��!", vbInformation, G_STR_HINT_TITLE
        DeleteManyImages = False
        Exit Function
    End If
    
    '�����ݿ���ɾ��ͼ��
    gcnVideoOracle.BeginTrans
    rsTemp.MoveFirst
    While Not rsTemp.EOF
        strSQL = "ZL_Ӱ��ͼ��_DELETE(" & rsTemp("ҽ��ID") & "," & rsTemp("���ͺ�") & ",'" & rsTemp("ͼ��UID") & "','" & "" & "')"
        zlDatabase.ExecuteProcedure strSQL, "Ӱ��ͼ��ɾ��"
        rsTemp.MoveNext
    Wend
    gcnVideoOracle.CommitTrans
    '׼����FTP��ɾ��ͼ���Ȳ����豸һ���ڲ����豸��
    
    rsTemp.MoveFirst
    If Not IsNull(rsTemp!�豸��1) Then
        strFTPIP = Nvl(rsTemp!Host1)
        strFTPUser = Nvl(rsTemp!User1)
        strFtpPass = Nvl(rsTemp!Pwd1)

        intDeviceUsed = 1
        lngResult = iNet.FuncFtpConnect(strFTPIP, strFTPUser, strFtpPass)

        If lngResult = 0 Then
            If Not IsNull(rsTemp!�豸��2) Then
                strFTPIP = Nvl(rsTemp!Host2)
                strFTPUser = Nvl(rsTemp!User2)
                strFtpPass = Nvl(rsTemp!Pwd2)

                intDeviceUsed = 2
                lngResult = iNet.FuncFtpConnect(strFTPIP, strFTPUser, strFtpPass)

                If lngResult = 0 Then
                    If MsgboxCus("����FTPʧ�ܣ��Ƿ����ɾ��ͼ��" & vbCrLf & "��ʱ����ɾ������ֻ��ɾ�����ݿ����ݣ��޷�ɾ��ͼ���ļ���" & vbCrLf & "���ǡ������ɾ����������ɾ����", vbQuestion + vbYesNo, G_STR_HINT_TITLE) = vbNo Then
                        DeleteManyImages = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    strRoot = IIf(intDeviceUsed = 1, Nvl(rsTemp!Root1), Nvl(rsTemp!Root2))
    
    'FTP��ɾ��ͼ���ļ�
    If UBound(Split(strArrImgs, ";")) > 0 Then
        For i = 0 To UBound(Split(strArrImgs, ";"))
            strImageUID = Split(strArrImgs, ";")(i)
            Call iNet.FuncDelFile(strRoot, strImageUID)
            Call iNet.FuncDelFile(strRoot, strImageUID & ".jpg")
        Next
    End If
    
    '�ر�FTP����
    iNet.FuncFtpDisConnect
    DeleteManyImages = True
    
    Exit Function
errH:
    gcnVideoOracle.RollbackTrans
    BUGEX "DeleteManyImages err"
    DeleteManyImages = False
    Call RaiseErr(err)
End Function

Public Sub LockStudy(ByVal intType As Integer, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, _
ByVal lngStudyState As Long, ByVal blnMoved As Boolean)
'���� ���� �������
'intType=0:��ǰ�Ĺ��ܣ�ֻ�ı����ֵ ; 1: ��������(δ����ͬ����λ������б�);   2���������     3:������(�Ѿ�����ͬ����λ������б�)
    Dim cbrControl As CommandBarControl
    
    Set cbrControl = cbrMain.FindControl(, conMenu_Cap_StudySyncState)
     
    '����intType�ж���Ҫ�������ǽ�����Ӱ��VideoCaptureMenuProcess���ж�����
    If intType = 0 Then
        mblnCaptureLockState = True
        Exit Sub
    ElseIf intType = 1 Then
        
        mblnCaptureLockState = False
        Call zlUpdateAdviceInf(lngAdviceId, lngSendNo, lngStudyState, blnMoved)
        mblnCaptureLockState = True
    
        cbrControl.IconId = 10012
    ElseIf intType = 2 Then
        cbrControl.IconId = 8123
    ElseIf intType = 3 Then
        cbrControl.IconId = 10012
    End If
    
    If Not cbrControl Is Nothing Then Call VideoCaptureMenuProcess(cbrControl)
End Sub

Public Function RefreshImageCaptureFace(ByVal blIsDock As Boolean)
'ˢ�²ɼ���������ͼ�߶ȿؼ�λ�ã����¿ؼ��ڲ�����
    Call ucPreview.RefreshFace(blIsDock)
    Call ucCacheImage.RefreshFace(blIsDock)
End Function

Public Function SaveImageToStady(dcmImage As DicomImage, lngAdviceId As Long)
    Dim dcmTag As New clsImageTagInf
    
    If mstrInstanceUID = dcmImage.InstanceUID Then Exit Function
    '����ͼ���DICOM����
    subWriteDicomPara dcmImage, lngAdviceId
    
    dcmTag.Tag = imgTag
    
    Set dcmImage.Tag = dcmTag
    mintCaptureFlag = 1
    
    '��ͼ����뵽����ͼ��
    subInsert2Mini dcmImage
    
    '����ͼ�񣬲�����ͼ��洢�¼�
    Call SaveImage(dcmImage, lngAdviceId, mcurStudyInf.strStudyUid, 0, True)
    
    mstrInstanceUID = dcmImage.InstanceUID
End Function
