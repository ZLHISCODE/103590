VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVideoProcess.ocx"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Begin VB.Form frmVideoCaptureV2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   10410
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmVideoCaptureV2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Tag             =   "��Ƶ�ɼ�"
   Begin ScanLibCtl.ImgScan ImageScanner 
      Left            =   1440
      Top             =   0
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
   Begin VB.PictureBox picAfter 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3840
      ScaleHeight     =   375
      ScaleWidth      =   2655
      TabIndex        =   12
      Top             =   5280
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label labCloseAfter 
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   90
         Width           =   255
      End
      Begin VB.Label labAfterInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "��ʶ:---"
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
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   90
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.PictureBox picLock 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3720
      ScaleHeight     =   375
      ScaleWidth      =   2655
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label labCloseLock 
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   90
         Width           =   255
      End
      Begin VB.Label labLockInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "����:---"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   90
         Width           =   2415
      End
   End
   Begin VB.PictureBox picView 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   1680
      ScaleHeight     =   3615
      ScaleWidth      =   6855
      TabIndex        =   5
      Top             =   1440
      Width           =   6855
      Begin ZLDSVideoProcess.DSCapture wdmCapture 
         Height          =   3375
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4215
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
         CurWidth        =   281
         CurHeight       =   225
         CurVideoWidth   =   279
         CurVideoHeight  =   205
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
         Left            =   4440
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox picCusVideo 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   4440
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   72
         TabIndex        =   7
         Top             =   240
         Width           =   1080
      End
      Begin DicomObjects.DicomViewer dcmView 
         Height          =   855
         Left            =   4440
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
         _Version        =   262147
         _ExtentX        =   1931
         _ExtentY        =   1508
         _StockProps     =   35
         UseScrollBars   =   0   'False
      End
   End
   Begin VB.PictureBox pbxSize 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   0
      Left            =   1440
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   7335
      TabIndex        =   4
      Top             =   1200
      Width           =   7335
   End
   Begin VB.PictureBox pbxSize 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   3975
      Index           =   3
      Left            =   8760
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3975
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   1215
      Width           =   75
   End
   Begin VB.PictureBox pbxSize 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   3975
      Index           =   2
      Left            =   1440
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3975
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   1200
      Width           =   75
   End
   Begin VB.Timer tmrReg 
      Interval        =   10000
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer tmrComm 
      Interval        =   2
      Left            =   480
      Top             =   0
   End
   Begin MSCommLib.MSComm commListener 
      Left            =   2640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picTemp2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pbxSize 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   75
      Index           =   1
      Left            =   1440
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   7335
      TabIndex        =   1
      Top             =   5160
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
Attribute VB_Name = "frmVideoCaptureV2"
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


Private Const conMenu_ImgPro_Window = 501           '���ȶԱȶ�
Private Const conMenu_ImgPro_Zoom = 502             '����
Private Const conMenu_ImgPro_Corp = 512             '�϶�

Private Const conMenu_ImgPro_Rotate_Pop = 503          '˳ʱ����ת
Private Const conMenu_ImgPro_RRotate = 5030          '˳ʱ����ת
Private Const conMenu_ImgPro_LRotate = 5031          '��ʱ����ת

Private Const conMenu_ImgPro_Smooth_Pop = 504        '��
Private Const conMenu_ImgPro_Sharpness = 5040        '��
Private Const conMenu_ImgPro_Smooth = 5041           'ƽ��

Private Const conMenu_ImgPro_Lab_Pop = 505
Private Const conMenu_ImgPro_Text = 5050             '���ֱ�ע
Private Const conMenu_ImgPro_Arrow = 5051            '��ͷ��ע
Private Const conMenu_ImgPro_Ellipse = 5052          'Բ�α�ע

Private Const conMenu_ImgPro_Save = 506         '����
Private Const conMenu_ImgPro_RectSave = 50601        '�ü�����
Private Const conMenu_ImgPro_DirectSave = 50602        'ֱ�ӱ���
Private Const conMenu_ImgPro_RectCapture = 507         '�ü���ɼ�
Private Const conMenu_ImgPro_Restore = 508       '�ָ�



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


'COM��̤�˿�״̬
Private Type TComPortState
    intComState As Integer          'COM�ڵ�״̬
    lngComTime As Long              '��¼com�ڱ���״̬��ʱ��
    dtLastCapture As Date           '�����̤���µ�ʱ��
    blnCTSHolding As Boolean        '��¼��̬ʱ��CTS�ߵĵ�ƽ
    
    lngSignalCount As Long
    lngStartTick As Long
    lngEndTick As Long
End Type


Private Type DlgFileInfo
    iCount As Long
    sPath As String
    sFIle() As String
End Type

Private Enum Dkp_ID
    Dkp_ID_Video = 1     '����б�
    Dkp_ID_Miniature      '��ǰ���˻�����Ϣ
End Enum



Private mobjCapHelper As ICapHelper
Private mstrAfterCapTag As String
Private mstrBufferDir As String
Private mblnIsLock As Boolean
 
Private mintCaptureFlag As Integer

Private mobjCustomDevice As Object  'ר����Ƶ�ɼ�����

Private dcmglbUID As New DicomGlobal    '����UIDRoot=1

Private WithEvents mobjDxDevice As clsDxHidDevice   'ʵ�������ֱ�֮��Ĳɼ���ʽ
Attribute mobjDxDevice.VB_VarHelpID = -1
 
Private WithEvents mfrmParameter As frmVideoSetupV2
Attribute mfrmParameter.VB_VarHelpID = -1
Private mfrmOpenStudy As frmOpenStudyList

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
 

Private mblnMoveDown  As Boolean        '�����ж��Ƿ���������
Private mblnDcmViewDown As Boolean      '�����ж�dcmView������Ƿ񱻰���
Private mintCurImgIndex As Integer      '��ǰ��ѡ�е�ͼ������
Private mdcmSelectLabel As DicomLabel   '��ǰ��ѡ�еı�ע

Private mstrAviFileName As String       '¼���ļ���

Private mstrAfterDir As String

Private mcpsComState As TComPortState       'Com�˿�ʹ��״̬


Private mlngImageSwapWay As Long          '0-�ڴ潻��,1-�����彻����2-�����ļ�����
Private mblnUseBeforeConvert As Boolean     '��ǰת��


Private mblnReadOnly As Boolean         '�Ƿ�ֻ�ܲ鿴True�鿴ģʽ��False�ɼ�ģʽ
 

Private mVideoCapture As clsVideoCapture '��Ƶ�ɼ�����
Private WithEvents mobjPlayWindow As frmPlaying
Attribute mobjPlayWindow.VB_VarHelpID = -1

Private mdblZoomRate As Double  '���ű��ʣ���cbrMain��cbrMain_ResizeClient�¼�����Ҫ���¼����ֵ��
Private mVideoSize As TVideoSize '��Ƶ��С������ص���Ƶ������棩
Private mCurCutRange As TCutRange '��Ƶ�ü���Χ���ã��ò���ͨ��GetString��SaveString������ע����У�
Private mVideoArea As TVideoArea  '��Ƶ�ͻ��������ã���cbrMain��cbrMain_ResizeClient�¼�����Ҫ���¼����ֵ��

Private Const M_LNG_REFRESHINTERVAL As Long = 600 'ˢ�¼��
Private mstrVideoRegTime As String '������Ƶ����ע��ʱ��
Private mstrMsg As String
Private mblnRefreshState As Boolean
Private mblnInitState As Boolean

Private mintFontSize As Integer '�ֺ�
 
Private mblnImageShield As Boolean   '�Ƿ����δ�ͼ
Private mblnTimerState As Boolean '��ʱ����״̬

Private Const CAPTURE_PARAMETER_CONFIG_FILE_NAME As String = "ZLVideoProcess.ini"
Private Const CONST_STR_DEFAULT_TEMP_SCAN_DIR_NAME As String = "\TempScan"  'Ĭ��ɨ���ļ���ʱ�洢·��
Private Const CONST_STR_DEFAULT_SCAN_FILE_TEMPLATE As String = "Scan"  'Ĭ��ɨ���ļ���ʱ�洢·��



'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'----------------------------------------------------------------------------------------------------------

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Public Event OnControlResize(objControl As Object)


'��ȡ��Ƶ�ɼ�����
Property Get videoCapture() As clsVideoCapture
    Set videoCapture = mVideoCapture
End Property


'��ȡ��Ƶ�ɼ����ڵĳ�ʼ��״̬
Property Get InitState() As Boolean
    InitState = mblnInitState
End Property

'��ǰ����״̬
Property Get IsLock() As Boolean
    IsLock = mblnIsLock
End Property

'�Ƿ��̨�ɼ���
Property Get IsAfter() As Boolean
    IsAfter = IIf(Len(mstrAfterCapTag) > 0, True, False)
End Property
'
'Private Sub UnLockStudy()
''�������
'    mblnIsLock = False
'End Sub


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


Private Sub Form_Initialize()
'��ʼ��ģ�����
    mblnInitState = False
    mblnIsLock = False
    mblnTimerState = False
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

    
BUGEX "ShowVideoConfig 6"
        
    If gobjCapturePar.VideoDirverType = vdtCustom Then
        If mobjCustomDevice Is Nothing Then Call InitCustomDevice

        Call mobjCustomDevice.StartPreview
        Call mobjCustomDevice.UpdateWindow(picCusVideo.ScaleWidth, picCusVideo.ScaleHeight)
    End If
    
    Call OpenComm
    
    gstrHotKeyTest = GetSetting("ZLSOFT", "����ģ��", "�ɼ��ȼ�", "F8")
    
    
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
    
 

    '��������ڴ��̵ĸ�Ŀ¼��app.pathΪ��x:\��
    mstrBufferDir = GetAppPath & "\TmpImage\"
    mstrAfterDir = GetAppPath & "\TmpAfterImage\"
    
'    mstrAviFileName = mstrBufferDir & "TmpVideo.avi"
    
    gint��Ƶ�豸���� = getLicenseCount(LOGIN_TYPE_��Ƶ�豸)
    
    mlngImageSwapWay = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "ͼ�񽻻���ʽ", 0))
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "ͼ�񽻻���ʽ", mlngImageSwapWay)
    
    mblnUseBeforeConvert = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "ͼ����ǰת��", 0)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "ͼ����ǰת��", IIf(mblnUseBeforeConvert, 1, 0))

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
 
End Sub


Public Sub zlRePreview(Optional ByVal blnForceStop As Boolean = False)
'���½�����ƵԤ��
    If mVideoCapture.IsStartup Then
        If blnForceStop Or mVideoCapture.VideoDriverType <> vdtWDM Then
            Call mVideoCapture.StopPreview
            Call mVideoCapture.StartPreview
        End If
         
        Call wdmCapture.RePreview
    End If
End Sub

Public Sub zlInitModule(objCapHelper As Object)
BUGEX "zlPacsCapture zlInitModule 0"
'��ʼ��ģ�����
    Set mobjCapHelper = objCapHelper
    
    If mblnInitState Then Exit Sub
    
    '��ʼ������
    Call InitParameter
    
BUGEX "zlInitModule 1"
    '��ʼ��ר����Ƶ�ɼ��ӿ�
    Call InitCustomDevice
    
BUGEX "zlInitModule 2"
    '����Ƶ�ɼ��豸
    Call OpenVideoCaptureDevice
  
BUGEX "zlInitModule End"
    mblnInitState = True
End Sub

Private Sub InitCustomDevice()
    Dim strCustomDeviceDir As String        'ר����Ƶ�ɼ�����·��
    Dim strCustomDeviceDllName As String    'ר����Ƶ�ɼ���������
    Dim objFile As New FileSystemObject
    
    '��ʼ��ר����Ƶ�ɼ��ӿ�
    strCustomDeviceDir = gobjCapturePar.CustomDevicePath
    If Trim(strCustomDeviceDir) <> "" And gobjCapturePar.VideoDirverType = vdtCustom Then
        strCustomDeviceDllName = Trim(Replace(objFile.GetFileName(strCustomDeviceDir), ".dll", ""))
        
        Set mobjCustomDevice = CreateObject(strCustomDeviceDllName & ".cls" & strCustomDeviceDllName)
        
        If Not mobjCustomDevice Is Nothing Then
            Call mobjCustomDevice.zlInit(gcnVideoOracle, UserInfo.id, glngDepartId, picCusVideo.hwnd)
        End If
    End If
End Sub


Public Sub zlRestoreWindow(ByVal blnReadOnly As Boolean, Optional ByVal blnIsMain As Boolean = False, _
    Optional ByVal blnIsOnlyState As Boolean = False)
'ˢ�½���
On Error GoTo errHandle
    mblnReadOnly = blnReadOnly
    
    If blnIsOnlyState Then Exit Sub
    
    If blnIsMain And cbrMain(2).position <> xtpBarRight Then
        cbrMain(2).position = xtpBarRight
        cbrMain.RecalcLayout
    ElseIf blnIsMain = False And cbrMain(2).position <> xtpBarLeft Then
        cbrMain(2).position = xtpBarLeft
        cbrMain.RecalcLayout
    End If
    
    If IsTwainCaptureWay Then Exit Sub
 
    Call ConfigVideoShowState(True)
    
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Public Sub zlPreviewThumbnail(objImg As Object)
'Ԥ������ͼ
    Dim dblTempZoom As Double
    Dim img As DicomImage
    Dim i As Integer
    
    '����ѡ��ͼ��װ�ص�dcmView��
    dcmView.Images.Clear
    
    If objImg Is Nothing Then Exit Sub
    
    If txtInputText.Visible Then txtInputText.Visible = False
 
    dcmView.Images.Add objImg
    
    '��ʾdcmView������picVideo
    dcmView.CurrentImage.BorderWidth = 0
    
    dblTempZoom = dcmView.CurrentImage.ActualZoom
    dcmView.CurrentImage.StretchToFit = False
        
    '�жϵ����븡������ʱ�����ű��ʲ���С��0.1
    If dblTempZoom < 0.1 Then dblTempZoom = 0.1
                  
    Call subCenterZoom(Me, dcmView.CurrentImage, dcmView, dblTempZoom, mCorpSize)
    
    Call ConfigVideoShowState(False)
End Sub


Private Sub StopCapture()
'-----------------------------------------------------------------------------------------
'���ܣ�ֹͣ��ʾ��Ƶ�ɼ�,�ͷ���Ƶ�ɼ����ڣ�
'      �ͷŴ��������Ķ˿�
'��������
'���أ���
'-----------------------------------------------------------------------------------------
    
    '�ر�COMM��
    If commListener.PortOpen Then commListener.PortOpen = False
    
    '����Midi�ӿ����������¼����
    If Not mobjDxDevice Is Nothing Then
        If mobjDxDevice.Handle <> 0 Then Call mobjDxDevice.CloseDxDevice
    End If
    
    '�ͷŲɼ��豸������
    If Not mVideoCapture Is Nothing Then
        Call mVideoCapture.StopPreview
    End If
End Sub



Public Sub zlUpdateCommandBars(Control As XtremeCommandBars.CommandBarControl)
'ֻ��Ӱ��ɼ�����վ�ž߱���̨�ɼ�����

'���ݵ�ǰ״̬ȷ���˵��Ŀ��ӺͿɲ���

    '���û�г�ʼ����Ƶ��������Ƶ��صİ�ť��������ʹ��
    If mVideoCapture Is Nothing Then
        Control.Enabled = False
        Exit Sub
    End If
    
    Select Case Control.id
        Case conMenu_Cap_Dynamic       '��̬��ʾ
            Control.Checked = mblnRealTime
            Control.Enabled = (Not mblnReadOnly Or Len(mstrAfterCapTag) > 0 Or mblnIsLock) And (Not IsTwainCaptureWay) And (mVideoCapture.IsStartup Or IsCustomCaptureWay)    ' And (mhCapWnd <> 0) modify by tjh at 2010-01-20
            Control.Visible = Not IsTwainCaptureWay 'And Not IsCustomCaptureWay
            
            If mblnRealTime Then
                Control.IconId = conMenu_Cap_Dynamic
            Else
                Control.IconId = 10023
            End If
            
        Case conMenu_Cap_MarkMap       'Ӱ��ɼ�
            Control.Enabled = (Not mblnReadOnly Or Len(mstrAfterCapTag) > 0 Or mblnIsLock) And (mVideoCapture.IsStartup Or IsTwainCaptureWay Or IsCustomCaptureWay)
            
'        Case conMenu_Cap_After_Capture  '��̨�ɼ�
'            Control.Enabled = mVideoCapture.IsStartup
'            Control.Visible = gobjCapturePar.IsUseAfterCapture And (Not IsCustomCaptureWay)
            
        Case conMenu_Cap_Record        '¼��
            Control.Enabled = (Not mblnReadOnly Or mblnIsLock) And ((gobjCapturePar.VideoDirverType = vdtWDM And mVideoCapture.IsStartup) Or gobjCapturePar.VideoDirverType = vdtCustom)
            Control.Visible = Not IsTwainCaptureWay And Len(mstrAfterCapTag) <= 0
            
        Case conMenu_Cap_Timer
            Control.Visible = gobjCapturePar.VideoDirverType = vdtWDM
            Control.Enabled = Not mblnReadOnly
            If mblnTimerState Then
                Control.IconId = 10025
                Control.ToolTipText = "�رռ�ʱ"
            Else
                Control.IconId = 10024
                Control.ToolTipText = "������ʱ"
            End If
'        Case conMenu_Cap_After_Record   '��̨¼��
'            Control.Enabled = gobjCapturePar.VideoDirverType = vdtWDM And mVideoCapture.IsStartup
'            Control.Visible = Not IsTwainCaptureWay And Not IsCustomCaptureWay And gobjCapturePar.IsUseAfterCapture And False
            
'        Case conMenu_Cap_Record_Stop 'ֹͣ¼�� modify by tjh at 2010-01-22
'            Control.Enabled = mblnRealTime And Not mblnReadOnly And (gobjCapturePar.VideoDirverType = vdtWDM) And mVideoCapture.IsStartup
'            Control.Visible = Not IsTwainCaptureWay And Not IsCustomCaptureWay
            
        Case conMenu_Cap_RecordAudio '¼��
            Control.Enabled = Not mblnReadOnly Or mblnIsLock
            Control.Visible = Not IsCustomCaptureWay And Len(mstrAfterCapTag) <= 0
            
        '¼�񲥷�,¼��ֹͣ,¼����,¼�����,����¼��
        Case conMenu_Cap_Play, conMenu_Cap_Stop, conMenu_Cap_Forward, _
             conMenu_Cap_Back
            If (mblnRealTime = False) And (dcmView.Images.count > 0) Then
                Control.Visible = dcmView.Images(1).Tag.Tag = VIDEOTAG Or dcmView.Images(1).Tag.Tag = AUDIOTAG
                Control.Enabled = dcmView.Images(1).Tag.Tag = VIDEOTAG Or dcmView.Images(1).Tag.Tag = AUDIOTAG
            Else
                Control.Visible = False
                Control.Enabled = False
            End If
            
         '���ȶԱȶ�,����,�ü�����,˳ʱ����ת,��ʱ����ת,��,ƽ��,�߼�����
        Case conMenu_ImgPro_Window, conMenu_ImgPro_Zoom, conMenu_ImgPro_Save, conMenu_ImgPro_RectSave, conMenu_ImgPro_DirectSave, _
             conMenu_ImgPro_Rotate_Pop, conMenu_ImgPro_RRotate, conMenu_ImgPro_LRotate, _
             conMenu_ImgPro_Smooth_Pop, conMenu_ImgPro_Sharpness, conMenu_ImgPro_Smooth, conMenu_ImgPro_Corp

            Control.Visible = dcmView.Visible
            Control.Enabled = (mblnRealTime = False)
        '��ͷ��ע,Բ�α�ע,���ֱ�ע,
        Case conMenu_ImgPro_Lab_Pop, conMenu_ImgPro_Arrow, conMenu_ImgPro_Ellipse, conMenu_ImgPro_Text
            Control.Visible = dcmView.Visible
            Control.Enabled = (mblnRealTime = False)
            
'        Case conMenu_Cap_OpenStudyList
'            Control.Enabled = True
'            Control.Visible = gobjCapturePar.IsUseCaptureLock
            
        Case conMenu_Cap_StudySyncState
            Control.Enabled = Not mblnReadOnly Or mblnIsLock
'            Control.Visible = gobjCapturePar.IsUseCaptureLock
            
        Case conMenu_Cap_After_Tag
            Control.Enabled = mVideoCapture.IsStartup
'            Control.Visible = gobjCapturePar.IsUseAfterCapture
            
'        ''''''''''''
'        Case conMenu_Cap_ImgImport
'            Control.Enabled = Not mblnReadOnly
            
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

Private Sub DoScanCapture()
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


Public Sub ForeCapture(ByVal blnIsReal As Boolean)
'ǰ̨ͼ��ɼ�
    Dim blnIsRealCap As Boolean
    
    If Not ((Not mblnReadOnly Or Len(mstrAfterCapTag) > 0 Or mblnIsLock) And (mVideoCapture.IsStartup Or IsTwainCaptureWay Or IsCustomCaptureWay)) Then Exit Sub '���Ϊֻ����������Ƶû��������������ɼ�
            
    If IsTwainCaptureWay Then
        Call DoScanCapture  'ͨ��TWAIN�ӿڲɼ�ͼ��
    Else
        If Not blnIsReal Then
            blnIsRealCap = IIf(MsgboxCus("ȷ��Ҫ�ɼ���ǰ��̬ͼ����ѡ���ǡ���ɼ���ǰ����ͼ��", _
                                vbQuestion + vbYesNo + vbDefaultButton1, G_STR_HINT_TITLE) = vbYes, False, True)
        End If
        
        If blnIsReal = False Then
            '��ʵʱ�ɼ�
            Call DoNormalCapture(False)
            Exit Sub
        End If
        
        If IsCustomCaptureWay Then
            '�Զ���ɼ�
            Call DoCustomCapture
        Else
            '�ɼ�ͼ��
            Call DoNormalCapture(True)
        End If
    End If
End Sub



Public Sub zlExecuteCommandBars(Control As XtremeCommandBars.CommandBarControl)
    On Error GoTo errHandle
        Call VideoCaptureMenuProcess(Control)
        
        Call DicomImageMenuProcess(Control)
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub VideoCaptureMenuProcess(Control As XtremeCommandBars.CommandBarControl)
'��Ƶ�ɼ��˵�����
    Select Case Control.id
        Case conMenu_Cap_Dynamic       '��̬��ʾ
            If IsTwainCaptureWay Then
                Call MsgboxCus("TWAIN�ɼ�ģʽ�£����ܽ��ж�̬��Ƶ����ʾ��", vbOKOnly, G_STR_HINT_TITLE)
            Else
                Call ConfigVideoShowState(True)
            End If
            
        Case conMenu_Cap_MarkMap       'Ӱ��ɼ�
            If Len(mstrAfterCapTag) <= 0 Then
                Call ForeCapture(True)
            Else
                Call AfterCapture
            End If
            
        Case conMenu_Cap_Timer
            Call StartTimer
            
'        Case conMenu_Cap_After_Capture  '��̨�ɼ�
'            Call AfterCapture
            
        Case conMenu_Cap_Record        '¼��
            If mstrVideoRegTime = "" Then
                MsgboxCus mstrMsg, vbOKOnly, "��ʾ"
                Exit Sub
            End If
            
            If Control.IconId = conMenu_Cap_Record Then
                Control.IconId = conMenu_Cap_Record_Stop
                
                Call ConfigVideoShowState(True)
                Call StartVideo '��ʼ¼��
            Else
                Control.IconId = conMenu_Cap_Record
                Call StopVideo  'ֹͣ¼��
            End If
            
'        Case conMenu_Cap_Record_Stop  'ֹͣ¼�� modify by tjh at 2010-01-22
'            If mstrVideoRegTime = "" Then
'                'MsgboxCus  "δ��⵽��Ч��ע����Ϣ�����ܽ���¼�������", vbOKOnly, "��ʾ"
'                Exit Sub
'            End If
'
'            If Len(mstrAviFileName) > 0 Then
'                Call StopVideo
'            End If
            
        Case conMenu_Cap_RecordAudio    '¼��
            If mstrVideoRegTime = "" Then
                MsgboxCus mstrMsg, vbOKOnly, "��ʾ"
                Exit Sub
            End If
            
            Call frmRecordAudio.ShowRecordAudio(Me)
            
        Case conMenu_Cap_Play          '¼�񲥷�
            Call PlayCurVideo
                
'        Case conMenu_Cap_OpenStudyList      '�򿪼��ɼ�ͼ��
'            Call mobjCapHelper.OpenLocker
            
        Case conMenu_Cap_StudySyncState     '�������
            If Control.IconId = 10012 Then
                Call CloseAfterCap
                Call LockCapture(Control)
            Else
                Call UnLockCapture(Control)
            End If
        Case conMenu_Cap_After_Tag      '���º�̨�ɼ���ʶ
            
            If mstrVideoRegTime = "" Then
                MsgboxCus mstrMsg, vbOKOnly, "��ʾ"
                Exit Sub
            End If
            
            If mblnIsLock Then
                MsgboxCus "����״̬���ܽ��к�̨���.", vbOKOnly, "��ʾ"
                Exit Sub
            End If
            
            
            Call ResetAfterCaptureTag
            
    End Select
End Sub

Public Sub ResetLockState(ByVal blnIsLock As Boolean)
    Dim objControl As XtremeCommandBars.CommandBarControl
    
    Set objControl = cbrMain.FindControl(, conMenu_Cap_StudySyncState, False, True)
    If objControl Is Nothing Then Exit Sub
    
    If blnIsLock Then
        CloseAfterCap
        Call LockCapture(objControl)
    Else
        Call UnLockCapture(objControl)
    End If
End Sub

Private Sub LockCapture(Control As XtremeCommandBars.CommandBarControl)
    Dim strLocker As String
    
    Control.IconId = 8123
    
    mblnIsLock = True
    
    Call mobjCapHelper.CapLock(strLocker)
    
    
    labLockInfo = "����:" & strLocker & ""
    picLock.Visible = True
    
    Call DrawBorderColor(True)
End Sub

Private Sub UnLockCapture(Control As XtremeCommandBars.CommandBarControl)
    Control.IconId = 10012
    
    mblnIsLock = False
    
    Call mobjCapHelper.CapUnlock
    
    labLockInfo.Caption = ""
    picLock.Visible = False
    
    Call DrawBorderColor(False)
End Sub

Private Sub DrawBorderColor(ByVal blnIsLock As Boolean)
    Dim lngColor As Long
    lngColor = IIf(blnIsLock, vbRed, vbBlue)
    
    pbxSize(0).BackColor = lngColor
    pbxSize(1).BackColor = lngColor
    pbxSize(2).BackColor = lngColor
    pbxSize(3).BackColor = lngColor
End Sub

Private Sub DicomImageMenuProcess(Control As XtremeCommandBars.CommandBarControl)
'dicomͼ��˵�����
    If mblnRealTime = True Or dcmView.Images.count <= 0 Then Exit Sub
    
    Select Case Control.id
        Case conMenu_ImgPro_Window         '���ȶԱȶ�
            subSetMouseState 1
            If mintMouseState = 1 Then
                Control.Checked = True
            End If
            
        Case conMenu_ImgPro_Zoom           '����
            subSetMouseState 2
            If mintMouseState = 2 Then
                Control.Checked = True
            End If
            
        Case conMenu_ImgPro_RectSave       '�ü�����
            subSetMouseState 3
            If mintMouseState = 3 Then
                Control.Checked = True
            End If
            
        Case conMenu_ImgPro_Save, conMenu_ImgPro_DirectSave ' ֱ�ӱ���
            Call CaptureFrameSelectImage
            
        Case conMenu_ImgPro_RectCapture         '�ü���ɼ�
            Call CaptureFrameSelectImage
            
        Case conMenu_ImgPro_Rotate_Pop, conMenu_ImgPro_RRotate         '˳ʱ����ת
            Call subSetRotate(dcmView.Images(1), True)
            
        Case conMenu_ImgPro_LRotate        '��ʱ����ת
            Call subSetRotate(dcmView.Images(1), False)
            
        Case conMenu_ImgPro_Sharpness      '��
            Call subSetSharp(dcmView.Images(1), True)
            
        Case conMenu_ImgPro_Smooth_Pop, conMenu_ImgPro_Smooth         'ƽ��
            Call subSetSharp(dcmView.Images(1), False)
            
        Case conMenu_ImgPro_Corp          '�϶�
           subSetMouseState 14
            If mintMouseState = 14 Then
                Control.Checked = True
            End If
            
        Case conMenu_ImgPro_Arrow          '��ͷ��ע
            subSetMouseState 11
            If mintMouseState = 11 Then
                Control.Checked = True
            End If
            
        Case conMenu_ImgPro_Ellipse        'Բ�α�ע
            subSetMouseState 12
            If mintMouseState = 12 Then
                Control.Checked = True
            End If
            
        Case conMenu_ImgPro_Lab_Pop, conMenu_ImgPro_Text            '���ֱ�ע
            subSetMouseState 13
'            If mintMouseState = 13 Then
'                Control.Checked = True
'            End If
    End Select
    
    If mintMouseState <> 0 Then picView.Refresh
End Sub


Public Sub zlUnloadMe()
    Unload Me
End Sub


Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    zlExecuteCommandBars Control
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
    Call ConfigVideoDisplay(picCusVideo)
    
    'ˢ����Ƶ��ʾ
    If IsCustomCaptureWay Then
        If Not (mobjCustomDevice Is Nothing) Then
            Call mobjCustomDevice.UpdateWindow(picCusVideo.ScaleWidth, picCusVideo.ScaleHeight)
        End If
    Else
        If Not (mVideoCapture Is Nothing) Then
            Call mVideoCapture.RefreshVideoWindow
        End If
    End If
    
    'ˢ��DcmView�е�ͼ����ʾλ��
    If dcmView.Images.count > 0 Then
        Call subCenterZoom(Me, dcmView.Images(1), dcmView, dcmView.Images(1).ActualZoom, mCorpSize)
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
        If dcmView.Images.count > 0 Then
            Call subCenterZoom(Me, dcmView.Images(1), dcmView, dcmView.Images(1).ActualZoom, mCorpSize)
        End If
    
    End If
 
    If pbxSize(0).Top - picLock.Height < Top Then
        picLock.Top = Top
    Else
        picLock.Top = pbxSize(0).Top - picLock.Height
    End If
    
    picLock.Left = Left + ((Right - Left) - picLock.Width) / 2
    
    
    If pbxSize(1).Top + picAfter.Height > Bottom - Top Then
        picAfter.Top = Bottom - picAfter.Height
    Else
        picAfter.Top = pbxSize(1).Top
    End If
    
    picAfter.Left = Left + ((Right - Left) - picAfter.Width) / 2
    
End Sub


Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    zlUpdateCommandBars Control
Exit Sub
errHandle:
    cbrMain.Options.UpdatePeriod = 2147483647
    MsgboxCus "�˵�״̬�����쳣:" & err.Description, vbOKOnly, "��ʾ"
End Sub

Private Sub DoListenCommSingal()
    Dim blnIsTouch As Boolean
    Dim lngTickCount As Long
    
    blnIsTouch = False
    
    lngTickCount = GetTickCount - mcpsComState.lngStartTick
'    strTouch = "���:" & Lpad(lngTickCount, 7) & "      ���:" & mcpsComState.lngSignalCount & "      "
    
    If lngTickCount <= gobjCapturePar.ComInterval Then
        mcpsComState.lngStartTick = GetTickCount
        mcpsComState.lngSignalCount = 1
    Else
        mcpsComState.lngSignalCount = mcpsComState.lngSignalCount + 1
        
        If GetTickCount - mcpsComState.lngEndTick > gobjCapturePar.ComInterval Then
            '�������Ų������źż���
            mcpsComState.lngSignalCount = 1
            mcpsComState.lngEndTick = GetTickCount
'            BUGEX ">>             ", True
        End If
        
        If mcpsComState.lngSignalCount >= gobjCapturePar.ComSignalCount Then
            '�ж��Ƿ�ָ��ʱ���ڽ��ܵ���Ӧ���ź���
            blnIsTouch = True
            
            mcpsComState.lngSignalCount = 0
            mcpsComState.lngStartTick = GetTickCount
            mcpsComState.lngEndTick = mcpsComState.lngStartTick
        End If
        
    End If
    
    If blnIsTouch = True And Not mblnReadOnly Then
'        BUGEX "**********************��̤����*********************", True
        commListener.PortOpen = False
        
        If mstrAfterCapTag = "" Then
            Call ForeCapture(True)
        Else
            Call AfterCapture
        End If
        
        commListener.PortOpen = True
    End If
End Sub


Private Sub DoListenCommStateData()

    Dim strInput As String
    

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
                    If mstrAfterCapTag = "" Then
                        Call ForeCapture(True)
                    Else
                        Call AfterCapture
                    End If
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
            If mstrAfterCapTag = "" Then
                Call ForeCapture(True)
            Else
                Call AfterCapture
            End If
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
                If mstrAfterCapTag = "" Then
                    Call ForeCapture(True)
                Else
                    Call AfterCapture
                End If
            End If
        End If
    End If
End Sub


Private Sub commListener_OnComm()
On Error GoTo errHandle

    '�����TWAINɨ���ר����Ƶ�ɼ�����֧�ֽ�̤���زɼ�
    If IsTwainCaptureWay Or IsCustomCaptureWay Then Exit Sub
    
    If gobjCapturePar.ComPortType <> "COM" Then Exit Sub
    
    If gobjCapturePar.ComSignalCount > 0 Then
        '����̤�ź�����ķ�ʽ���вɼ�
        Call DoListenCommSingal
    Else
        '����̤״̬��������ķ�ʽ���вɼ�
        Call DoListenCommStateData
    End If
    
    Exit Sub
errHandle:
    If commListener.PortOpen = False Then commListener.PortOpen = True
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub dcmView_DblClick()
On Error GoTo errHandle
    Call PlayCurVideo
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
  
  'picview��width��right���Ҷ˺��¶˵Ĳü���Χ���ã���������ı�
  
  If mVideoArea.Width <= picView.Width + pbxSize(2).Width Then
    picView.Left = mVideoArea.Left + pbxSize(2).Width * 2
  Else
    picView.Left = mVideoArea.Left + (mVideoArea.Width - picView.Width - 2 * pbxSize(2).Width) / 2 + 3 * pbxSize(2).Width
  End If
  
  If mVideoArea.Height <= picView.Height + pbxSize(0).Height Then
    picView.Top = mVideoArea.Top + pbxSize(0).Height * 2
  Else
    picView.Top = mVideoArea.Top + (mVideoArea.Height - picView.Height - 2 * pbxSize(0).Height) / 2 + 3 * pbxSize(2).Width
  End If
  
  '����DICOM��ʾͼ��Ĵ�С
  dcmView.Left = DICOM_VIEWER_BODER_SIZE
  dcmView.Top = DICOM_VIEWER_BODER_SIZE
  dcmView.Width = picView.Width - DICOM_VIEWER_BODER_SIZE * 2
  dcmView.Height = picView.Height - DICOM_VIEWER_BODER_SIZE * 2
End Sub


Private Sub ConfigTwainDisplay()
  '�߿��С
  Const DICOM_VIEWER_BODER_SIZE As Long = 5
  
  dcmView.Left = DICOM_VIEWER_BODER_SIZE
  dcmView.Top = DICOM_VIEWER_BODER_SIZE
  dcmView.Width = picView.Width - DICOM_VIEWER_BODER_SIZE * 2
  dcmView.Height = picView.Height - DICOM_VIEWER_BODER_SIZE * 2
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
    Dim lngCusWidth As Long
    Dim lngCusHeight As Long

BUGEX "OpenVideoCaptureDevice 1"

    If mVideoCapture Is Nothing Then
        '������Ƶ�ɼ�����
        Set mVideoCapture = New clsVideoCapture
        
        '������Ƶ������
        Call mVideoCapture.ConnectedVfwDeviceObj(picCusVideo)
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
    
    '������Ƶ��������
    mVideoCapture.VideoDriverType = gobjCapturePar.VideoDirverType
        
    If (Not mVideoCapture.IsStartup) Then
        
        '��ȡ��Ƶ��С
        mVideoSize.Width = ScaleX(mVideoCapture.VideoSize.Width, vbPixels, vbTwips)
        mVideoSize.Height = ScaleX(mVideoCapture.VideoSize.Height, vbPixels, vbTwips)
        
        '���ý���
        Call CaptureSwitchFace(IsTwainCaptureWay) 'Or IsCustomCaptureWay
       

        '*******************************************************
BUGEX "OpenVideoCaptureDevice 5"
        '��ʼ��ƵԤ��********************************************
        If Not IsTwainCaptureWay And Not IsCustomCaptureWay Then
            mblnRealTime = True
            
            Call mVideoCapture.StartPreview
                    
'            blnIsStartupVideo = mVideoCapture.IsStartup
        ElseIf IsCustomCaptureWay Then
            'ר�òɼ�
            mblnRealTime = True
            
            Call mobjCustomDevice.GetSizeInfo(lngCusWidth, lngCusHeight)
            
            mVideoSize.Width = ScaleX(lngCusWidth, vbPixels, vbTwips)
            mVideoSize.Height = ScaleX(lngCusHeight, vbPixels, vbTwips)
'            blnIsStartupVideo = True
        Else
            'twain�ɼ�
            mblnRealTime = False
            
'            blnIsStartupVideo = ImageScanner.ScannerAvailable
        End If
 

        '*********************************************************
    Else
        Call ConfigVideoShowState(True)
    End If
    
    Call OpenComm   '�򿪲ɼ��˿�
End Sub


Public Sub ResetAfterCaptureTag()
'���º�̨�ɼ���Ϣ
    Call mobjCapHelper.AfterTag(mstrAfterCapTag)
    
    labAfterInfo.Caption = "��ʶ:" & mstrAfterCapTag
    labAfterInfo.Visible = True
    picAfter.Visible = True
    
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
    '���������Ըô��ڶ�������ö�������������ִ�д򿪻��߱������ʱ���������ļ�ѡ���λ�ڸô���֮��
    SetWindowPos Me.hwnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3  '�������ö�
     
    Call InitCommandBars
    Call InitScanDir
    
    Set mfrmParameter = New frmVideoSetupV2
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub InitScanDir()
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
    picCusVideo.Visible = False
      
    If blnUseTwain Then
      Set dcmView.Container = Me
      Set txtInputText.Container = Me
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
    If IsCustomCaptureWay Then
        If Not (mobjCustomDevice Is Nothing) Then
            mobjCustomDevice.StopPreview
        End If
    Else
        Call mVideoCapture.StopPreview
    End If
    
    gobjCapturePar.VideoDirverType = videoDirver
    mVideoCapture.VideoDriverType = videoDirver
       
    '��ȡ��Ƶ��С
    mVideoSize.Width = ScaleX(mVideoCapture.VideoSize.Width, vbPixels, vbTwips)
    mVideoSize.Height = ScaleX(mVideoCapture.VideoSize.Height, vbPixels, vbTwips)
       
    Call CaptureSwitchFace(videoDirver = vdtTWAIN)
        
    
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
            'Call InitCustomDevice
            If Not (mobjCustomDevice Is Nothing) Then
                mobjCustomDevice.StartPreview
                Call mobjCustomDevice.UpdateWindow(picCusVideo.ScaleWidth, picCusVideo.ScaleHeight)
            End If
        End If
        
        mblnRealTime = True
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
        commListener.DTREnable = True
        commListener.EOFEnable = False
                        
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
        tmrComm.Interval = 20
    End If
BUGEX "OpenComm 10"
    Exit Sub
err:
BUGEX "OpenComm 11"
    Call MsgboxCus("�˿ڴ򿪴���", vbOKOnly, G_STR_HINT_TITLE)
BUGEX "OpenComm 12"
End Sub


Private Sub dcmView_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = 1 And dcmView.Images.count > 0 Then
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
                    
                    Set mdcmSelectLabel = dcmView.Images(1).Labels(dcmView.Images(1).Labels.count)
                    
                    If mintMouseState = 3 Then
                        mdcmSelectLabel.Tag = "CUT"
                    End If
                    
                    mdcmSelectLabel.LineWidth = 2
            End Select
        End If
    End If
End Sub


Private Sub dcmView_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim dblZoom As Double
    
    If mblnDcmViewDown = True And Button = 1 And dcmView.Images.count > 0 Then
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
    If mblnDcmViewDown = True And Button = 1 And dcmView.Images.count > 0 Then
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
            If dcmView.Images(1).Labels.count > 0 Then
                If dcmView.Images(1).Labels(dcmView.Images(1).Labels.count).Tag = "CUT" Then
                    dcmView.Images(1).Labels.Remove dcmView.Images(1).Labels.count
                End If
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

   
Private Function DoNormalCapture(ByVal blnIsReal As Boolean, Optional objPic As StdPicture = Nothing) As Boolean
'------------------------------------------------
'���ܣ��ɼ����洢ͼ��
'��������
'���أ��ޣ�ֱ�ӱ����²ɼ���ͼ��
'------------------------------------------------
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objImg As DicomImage
    Dim strError As String
    
    DoNormalCapture = False
    
    If mstrVideoRegTime = "" Then
        MsgboxCus mstrMsg, vbOKOnly, "��ʾ"
        Exit Function
    End If
    
     
    Set objImg = ConvertDcmImage(strError, blnIsReal, "", objPic)
    
    If Not objImg Is Nothing Then
         
        mintCaptureFlag = 2
        
        DoNormalCapture = mobjCapHelper.SaveImg(objImg, "", True)
    Else
        MsgboxCus strError, vbOKOnly, "��ʾ"
    End If
Exit Function
errHandle:
    err.Raise err.Number, err.Description
End Function

Private Function DoScanCaptureDown(ByVal strFile As String) As Boolean
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objImg As DicomImage
    Dim strError As String
    
    DoScanCaptureDown = False
    
    If mstrVideoRegTime = "" Then
        MsgboxCus mstrMsg, vbOKOnly, "��ʾ"
        Exit Function
    End If
    
     
    Set objImg = ConvertDcmImage(strError, , strFile)
    
    If Not objImg Is Nothing Then
         
        mintCaptureFlag = 2
        
        DoScanCaptureDown = mobjCapHelper.SaveImg(objImg, "", True)
        
        If DoScanCaptureDown Then Call ShowScanImage(objImg)
    Else
        MsgboxCus strError, vbOKOnly, "��ʾ"
    End If
Exit Function
errHandle:
    err.Raise err.Number, err.Description
End Function
 


Private Function DoCustomCapture() As Boolean
    Dim objCapPic As StdPicture
    Dim strCapImgFile As String
    Dim blnIsCusSave As Boolean
    Dim objImg As DicomImage
    Dim strError As String
    
    DoCustomCapture = False
    
    If mobjCustomDevice Is Nothing Then Exit Function
    
    '�ɼ�ͼ��
    If Not mobjCustomDevice.zlCaptureImage(mobjCapHelper.GetCustomMainID, _
        objCapPic, strCapImgFile, blnIsCusSave) Then
        Exit Function
    End If
    
    Set objImg = ConvertDcmImage(strError, , strCapImgFile, objCapPic)
    If Not objImg Is Nothing Then
        DoCustomCapture = mobjCapHelper.SaveImg(Nothing, "", Not blnIsCusSave)
    Else
        MsgboxCus strError, vbOKOnly, "��ʾ"
    End If
End Function

Public Sub AfterCapture()
'------------------------------------------------
'���ܣ���̨�ɼ�
'��������
'���أ��ޣ�ֱ�ӱ����²ɼ���ͼ��
'------------------------------------------------
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objImg As DicomImage
    Dim strError As String
    
    If mstrVideoRegTime = "" Then
        MsgboxCus mstrMsg, vbOKOnly, "��ʾ"
        Exit Sub
    End If
     
    If Not mVideoCapture.IsStartup Then
        MsgboxCus "��ƵԴ��δ�������ܲɼ���", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    '���û�к�̨�ɼ���ʶ��������µĺ�̨�ɼ���ʶ
    If Len(mstrAfterCapTag) <= 0 Then Call ResetAfterCaptureTag
     
    Set objImg = ConvertDcmImage(strError, True)
    
    If Not objImg Is Nothing Then
         
        mintCaptureFlag = 2
        
        Call mobjCapHelper.SaveImg(objImg, "", True, mstrAfterCapTag)
    Else
        MsgboxCus strError, vbOKOnly, "��ʾ"
    End If
Exit Sub
errHandle:
    err.Raise err.Number, err.Description
End Sub

Private Function PictureToDicomImg(ByVal lngHDC As Long, ByVal lngPictureHandle As Long, _
    objDcmImg As Object, ByRef strError As String) As Boolean
'congpicture�и���ͼ��dicomimage
    Const bitCount As Long = 3
        
    Dim hBitmap As OLE_HANDLE
    Dim stucbmp As TBitMap
    Dim lngSize As Long
    Dim lngResult As Long
    Dim aryPixels() As Byte
    Dim stuDipInf As BITMAPINFO
    
    Dim i As Long, j As Long, bytTemp As Byte
    
On Error GoTo errHandle
    PictureToDicomImg = False
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
    objDcmImg.InstanceUID = dcmglbUID.NewUID

    PictureToDicomImg = True
Exit Function
errHandle:
    strError = err.Description
End Function


Private Function ConvertDcmImage(ByRef strError As String, _
                        Optional ByVal blnRealState As Boolean = True, _
                        Optional ByVal strFileName As String = "", _
                        Optional objCapture As StdPicture = Nothing) As DicomImage
'------------------------------------------------
'���ܣ��ɼ���֡��Ƶͼ�񣬽�ͼ��ת����DICOM��ʽ������дDICOM�ļ�ͷ�����ͼ���������ͼdcmMiniature�С�
'��������
'���أ��ޣ�ֱ�ӽ��²ɼ���ͼ�����dcmMiniature��
'------------------------------------------------
'�ɼ���֡ͼ��
On Error GoTo SaveFileError
    Dim ImgTmpImage As DicomImage
    Dim dcmTag As clsImageTagInf
    Dim strFile As String
    
    '�ɼ�ͼ�񣬷�Ϊֱ����Ƶ�ɼ��Ͳ���¼��ɼ�
    Set ConvertDcmImage = Nothing

    If Not (objCapture Is Nothing) Then
        '��stdPicture��ȡͼ��
        Set picTemp2.Picture = Nothing
        
        picTemp2.Width = mVideoSize.Width * (1 - mCurCutRange.WidthRate - mCurCutRange.LeftRate)
        picTemp2.Height = mVideoSize.Height * (1 - mCurCutRange.HeightRate - mCurCutRange.TopRate)
        
        picTemp2.Picture = objCapture
        
    ElseIf Trim(strFileName) <> "" And Dir(strFileName) <> "" Then
        '���ļ���ȡͼ��
        Set picTemp2.Picture = Nothing
        Set picTemp2.Picture = LoadPicture(strFileName)
        
    Else
        If blnRealState = False And mblnPlayVideo = False Then
            'ʹ��dcmView��ʾ����ͼƬ������Ҫ�ٲü�
            Set picTemp2.Picture = Nothing
            
            If dcmView.Images.count > 0 Then
                Set picTemp2.Picture = dcmView.CurrentImage.Capture(False).Picture
            End If
        Else
            '������ʵʱ��Ƶ��ʾʱ����Ҫ��ͼ����вü�����
            Set picTemp2.Picture = Nothing
                        
            Dim curPic As StdPicture
            Set curPic = mVideoCapture.CaptureImageFromMemory

            If curPic Is Nothing Then
                strError = "��Ƶͼ��ɼ�ʧ�ܣ�������Ƶ���������Ƿ���ȷ(����Ƶ�豸����ʾģʽ��)��"
                Exit Function
            End If
            
            If mCurCutRange.LeftRate > 0.005 Or mCurCutRange.TopRate > 0.005 Then
                picTemp2.Width = mVideoSize.Width * (1 - mCurCutRange.WidthRate - mCurCutRange.LeftRate)
                picTemp2.Height = mVideoSize.Height * (1 - mCurCutRange.HeightRate - mCurCutRange.TopRate)
    
                'Ӧ��ͼ��Χ�ü�
                Call picTemp2.PaintPicture(curPic, 0, 0, picTemp2.Width, picTemp2.Height, _
                                           mVideoSize.Width * mCurCutRange.LeftRate, mVideoSize.Height * mCurCutRange.TopRate, _
                                           picTemp2.Width, picTemp2.Height, vbSrcCopy)
                                                   
                picTemp2.Picture = picTemp2.Image
            Else
                Set picTemp2.Picture = curPic
            End If

            Set curPic = Nothing
        End If
    End If
    
    '���û�вɼ���ͼ����ֱ���˳�
    If picTemp2.Picture Is Nothing Then
        strError = "δ�ɼ���ͼ��."
        Exit Function
    End If
    
    '����dicom��ʽͼ��
    Set ImgTmpImage = New DicomImage

    Select Case mlngImageSwapWay
        Case 0  '�ڴ�
            '��ʹ�ü����巽ʽ����Picture�и���ͼ��ImgTmpImage��,��ʹ�ü����彻������
            If PictureToDicomImg(picTemp2.hdc, picTemp2.Picture.Handle, ImgTmpImage, strError) = False Then
                Exit Function
            End If
        Case 1  '������
            If ClipboardToDicomImg(picTemp2.Picture, ImgTmpImage, dcmglbUID.NewUID, strError) = False Then
                Exit Function
            End If
        Case 2  '�ļ�
            If FileToDicomImg(picTemp2.Picture, ImgTmpImage, strError) = False Then
                Exit Function
            End If
    End Select
    
    Set ConvertDcmImage = ImgTmpImage

    Exit Function
SaveFileError:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function

Private Function FileToDicomImg(objPic As StdPicture, objDcmImg As Object, _
    ByRef strError As String) As Boolean
'�ļ���dicomimage
    Dim strFile As String
    
On Error GoTo errHandle
    FileToDicomImg = False

    strFile = mstrBufferDir & "ImageFile.SWAP"
    Call SavePicture(objPic, strFile)
    
    objDcmImg.FileImport strFile, "BMP"
    objDcmImg.InstanceUID = dcmglbUID.NewUID
    
    FileToDicomImg = True
Exit Function
errHandle:
    strError = err.Description
End Function


Private Sub Form_Resize()
On Error GoTo errHandle
    
    '����ͼ���С
    If Me.ScaleHeight < 7000 Or Me.ScaleWidth < 4000 Then
        cbrMain.Options.SetIconSize True, 16, 16
    Else
        cbrMain.Options.SetIconSize True, 32, 32
    End If

'    cbrMain.RecalcLayout
    
errHandle:
End Sub

Private Sub Form_Terminate()
    Set mdcmSelectLabel = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Long
    

BUGEX "VideoForm_UnLoad 1"
    tmrComm.Enabled = False
    
BUGEX "VideoForm_UnLoad 3"
    '�ȹرղɼ����ں�COMM��
    Call StopCapture
BUGEX "VideoForm_UnLoad 4"
    '���ֲü�����
    Call SaveParameterCfg
    
    
BUGEX "VideoForm_UnLoad 6"
    If Not mfrmParameter Is Nothing Then
        Unload mfrmParameter
    End If
    
BUGEX "VideoForm_UnLoad 8"
    wdmCapture.FreeRes
    
BUGEX "VideoForm_UnLoad 9"
    Set mobjCapHelper = Nothing
    
BUGEX "VideoForm_UnLoad 10"
    Set dcmglbUID = Nothing
    Set mobjDxDevice = Nothing
    Set mVideoCapture = Nothing
    Set mfrmParameter = Nothing
    
    If Not mobjCustomDevice Is Nothing Then
        mobjCustomDevice.zlFree
        Set mobjCustomDevice = Nothing
    End If
    
    If Not mobjPlayWindow Is Nothing Then
        Unload mobjPlayWindow
        Set mobjPlayWindow = Nothing
    End If
    
BUGEX "VideoForm_UnLoad End"
End Sub

 

Private Sub subSetMouseState(intMouseState As Integer)
    '�ı䵱ǰ���״̬
    mintMouseState = IIf(mintMouseState = intMouseState, 0, intMouseState)
    
    If txtInputText.Visible Then txtInputText.Visible = False
    
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_Window, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_Zoom, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_RectSave, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_DirectSave, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_Arrow, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_Ellipse, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_Text, False, True).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_ImgPro_Corp, False, True).Checked = False
'    cbrMain.FindControl(xtpControlSplitButtonPopup, conMenu_ImgPro_Lab_Pop, False, True).Checked = False
End Sub


'modify by tjh at 2010-01-20
'������Ƶ��ʾ״̬
Private Sub ConfigVideoShowState(ByVal blnShowState As Boolean)
  mblnRealTime = blnShowState
  
  Select Case gobjCapturePar.VideoDirverType
    Case vdtVFW, vdtCustom
      picCusVideo.Visible = blnShowState
      wdmCapture.Visible = False
      dcmView.Visible = Not blnShowState
    Case vdtWDM
      wdmCapture.Visible = blnShowState
      picCusVideo.Visible = False
      dcmView.Visible = Not blnShowState
    Case vdtTWAIN, vdtCustom
      wdmCapture.Visible = False
      picCusVideo.Visible = False
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
  Call DoScanCaptureDown(curScanFile)
End Sub


Private Sub ShowScanImage(objImg As DicomImage)

    '����ѡ��ͼ��װ�ص�dcmView��
    dcmView.Images.Clear
    dcmView.Images.Add objImg
    
    '��ʾdcmView������picVideo
    dcmView.CurrentImage.BorderWidth = 0
    mblnRealTime = False
'    picVideo.Visible = False
'    dcmView.Visible = True
End Sub

 

Private Sub labCloseAfter_Click()
On Error GoTo errHandle
    Call CloseAfterCap
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub CloseAfterCap()
    mstrAfterCapTag = ""
    labAfterInfo.Caption = "��ʶ:---"
    picAfter.Visible = False
    
    Call mobjCapHelper.AfterTag("CLOSE")
End Sub

Private Sub labCloseLock_Click()
On Error GoTo errHandle
    Dim Control As XtremeCommandBars.CommandBarControl
    
    Set Control = cbrMain.FindControl(, conMenu_Cap_StudySyncState, False, True)
    
    If Control Is Nothing Then Exit Sub
    
    Call UnLockCapture(Control)
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
                If mstrAfterCapTag <> "" Then Call CloseAfterCap
                
                Call ForeCapture(True)
                
        Case 1  '��̨�ɼ�
                BUGEX "mobjDxDevice_OnDxKeyPress 3"
'                If gobjCapturePar.IsUseAfterCapture Then
                    Call AfterCapture
'                End If
                
        Case 2  '���±�ʶ
                BUGEX "mobjDxDevice_OnDxKeyPress 4"
'                If gobjCapturePar.IsUseAfterCapture Then
                    Call ResetAfterCaptureTag
'                End If
                
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


Private Sub mobjPlayWindow_OnCapture(pic As stdole.StdPicture)
On Error GoTo errHandle
    Call DoNormalCapture(True, pic)
Exit Sub
errHandle:
    
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
    ElseIf picCusVideo.Visible Then
      Call ChangeCutRanage(picCusVideo, Index, X, Y)
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
    ElseIf picCusVideo.Visible Then
      Call ApplayCutRange(picCusVideo)
    End If
    
    If IsTwainCaptureWay Or IsCustomCaptureWay Then
      ConfigTwainDisplay
    Else
      '������ʾ��Χ
      Call ConfigVideoDisplay(wdmCapture)
      Call ConfigVideoDisplay(picCusVideo)

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

Private Sub PlayCurVideo()
'------------------------------------------------
'���ܣ�dcmView��¼��ͼ��Ĳ���
'��������
'���أ��ޣ�ֱ�Ӳ���dcmView�е�ͼ��
'------------------------------------------------
    Dim strFile As String
    
    If mobjPlayWindow Is Nothing Then
        Set mobjPlayWindow = New frmPlaying
    End If
    
    If dcmView.Images.count > 0 Then
        '����¼��������ش��ڣ��򲻽�������
        If dcmView.Images(1).Tag.Tag <> VIDEOTAG And dcmView.Images(1).Tag.Tag <> AUDIOTAG Then Exit Sub
        
        strFile = dcmView.Images(1).Tag.VideoFile
        
        '�򿪲��š���
        Call mobjPlayWindow.Show
        
        'ˢ�²��Ŵ���
        While Not mobjPlayWindow.IsActive
            Call Sleep(10)
            DoEvents
        Wend
            
        Call mobjPlayWindow.OpenVideoFile(Replace(strFile, "/", "\"), Nothing, True)
    End If
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
            dcmView.Images(1).Labels.Remove dcmView.Images(1).Labels.count
            txtInputText = "1 "
        Else
            mdcmSelectLabel.Text = txtInputText.Text
            dcmView.Refresh
        End If
    End If
End Sub

Private Sub StartVideo()
'------------------------------------------------
'���ܣ�¼��
'��������
'���أ���¼���ļ���������ͼ
'------------------------------------------------
    Dim strError As String
    Dim strVideoFile As String
    Dim strEncoderName As String
    Dim lngRecordTimeLen As String
    Dim blnIsSave As Boolean
    
    On Error GoTo continue1
    
      'ɾ����ʷ����Ƶ�ļ�
    mstrAviFileName = mstrBufferDir & "TmpVideo_" & Format(Now, "HHMMSS") & ".avi"
    If Dir(mstrAviFileName) <> "" Then RemoveFile mstrAviFileName
continue1:
    
    On Error GoTo CapErr
            
    '����Ŀǰ�ķ�ʽ,ʹ��vfw��ʱ���������¼�����
    If mVideoCapture.VideoDriverType = vdtVFW Or mVideoCapture.VideoDriverType = vdtTWAIN Then
        '¼�����(vfw����¼���ֱ��������ִ��StartVideo�Ժ�����)
        '������vfw��¼����
        MsgboxCus "������VFW��TWAIN������ʽ��¼����Ƶ��", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    If IsCustomCaptureWay Then
        If mobjCustomDevice Is Nothing Then
            MsgboxCus "ר����Ƶ�ɼ��ӿڵ���ʧ�ܣ�����¼����Ƶ��", vbOKOnly, "��ʾ"
            Exit Sub
        End If
         
        strVideoFile = mobjCustomDevice.zlStartVideo( _
                        mobjCapHelper.GetCustomMainID, _
                        mstrAviFileName, blnIsSave, _
                        strEncoderName, lngRecordTimeLen)
        
        If FileExists(strVideoFile) = False Then
            Exit Sub
        End If
        
        '������ﷵ�����ļ�������ֱ�ӱ���
        If Len(strVideoFile) > 0 Then
            Call mobjCapHelper.SaveVideo(strVideoFile, "", _
                                strEncoderName, lngRecordTimeLen, blnIsSave)
            
            mstrAviFileName = ""
        End If
    Else
        'modify by tjh at 2010-01-20
        strError = mVideoCapture.StartVideo(mstrAviFileName)
        If Trim(strError) <> "" Then MsgboxCus strError, vbInformation, "��ʾ"
    End If
    
    Exit Sub
CapErr:
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
    err.Clear
End Sub


'modify by tjh at 2010-01-20
'ֹͣ��Ƶ¼��
Private Sub StopVideo()
    Dim strVideoFile As String
    Dim strEncoderName As String
    Dim lngRecordTimeLen As Long
    
    Dim blnIsCusSave As Boolean
            
    If mVideoCapture.VideoDriverType = vdtVFW Or mVideoCapture.VideoDriverType = vdtTWAIN Then Exit Sub
    
On Error GoTo continue1
    If Dir(mstrAviFileName) <> "" Then RemoveFile mstrAviFileName
continue1:
    
    On Error GoTo CapErr
    
    If IsCustomCaptureWay Then
        If mobjCustomDevice Is Nothing Then
            MsgboxCus "ר����Ƶ�ɼ��ӿڵ���ʧ�ܡ�", vbOKOnly, "��ʾ"
            Exit Sub
        End If
         
        strVideoFile = mobjCustomDevice.zlstopVideo( _
                        mobjCapHelper.GetCustomMainID, _
                        mstrAviFileName, blnIsCusSave, _
                        strEncoderName, lngRecordTimeLen)
                        
        If FileExists(strVideoFile) = False Then
            MsgboxCus "ר����Ƶ¼���ļ���ȡʧ�ܡ�", vbOKOnly, "��ʾ"
            Exit Sub
        End If
    Else
        Call mVideoCapture.StopVideo
        
        strVideoFile = mstrAviFileName
        strEncoderName = mVideoCapture.GetEncoderName
        lngRecordTimeLen = mVideoCapture.GetTimeLen
        
        blnIsCusSave = False
    End If
       
    Call mobjCapHelper.SaveVideo(strVideoFile, "", _
                    strEncoderName, lngRecordTimeLen, Not blnIsCusSave)
    Exit Sub
CapErr:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


'ֹͣ��Ƶ�ļ�
Public Sub subSaveAudio(ByVal strAudioFile As String, ByVal lngTimeLen As Long)
On Error GoTo CapErr
   
    Call mobjCapHelper.SaveAudio(strAudioFile, "", "", lngTimeLen, True)
    
    Exit Sub
CapErr:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
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
    
    
    Set cbrToolBar = Me.cbrMain.Add("�ɼ�������", xtpBarLeft)
 
'    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False

    With cbrToolBar.Controls
    
        '�ڷ�TWAIN�ɼ�ģʽ������£�����ʾ�ð�ť
        'If Not GetIsTwainCaptureWay Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Dynamic, "��̬"): cbrControl.ToolTipText = "��ʾʵʱ��Ƶ"
        'End If
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_MarkMap, "�ɼ�"): cbrControl.ToolTipText = "�ɼ�ͼ��"
        
'        '���ú�̨�ɼ�
'        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Capture, "��̨�ɼ�"): cbrControl.ToolTipText = "��̨�ɼ�"
'            cbrControl.IconId = 10020
        
        '�ڷ�TWAIN�ɼ�ģʽ������£�����ʾ�ð�ť
        'If Not GetIsTwainCaptureWay Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Record, "¼��"): cbrControl.ToolTipText = "��ʼ¼��"
                cbrControl.Enabled = True
                
            
'            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Record, "��̨¼��"): cbrControl.ToolTipText = "��̨¼��"
'                cbrControl.IconId = 10021
            
            
'            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Record_Stop, "ֹͣ¼��"): cbrControl.ToolTipText = "ֹͣ¼��"
'                cbrControl.Enabled = False
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_RecordAudio, "¼��"): cbrControl.ToolTipText = "¼��"
        'End If
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Play, "����"): cbrControl.ToolTipText = "����¼��"
            cbrControl.BeginGroup = True
            
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_Timer, "��ʱ"): cbrControl.ToolTipText = "������ʱ"
            cbrControl.IconId = 10024
            
'        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_OpenStudyList, "�򿪼��"): cbrControl.ToolTipText = "�򿪼��"
'            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_StudySyncState, "�������"): cbrControl.ToolTipText = "�������"
            cbrControl.IconId = 10012
            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_After_Tag, "���±�ʶ"): cbrControl.ToolTipText = "���±�ʶ"
            cbrControl.IconId = 10022
            
        
            
            
        Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Window, "����"): cbrControl.ToolTipText = "��������/�Աȶ�": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Zoom, "����"): cbrControl.ToolTipText = "����ͼ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Corp, "�϶�"): cbrControl.ToolTipText = "�϶�ͼ��"
        
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_ImgPro_Save, "����"): cbrControl.ToolTipText = "����ͼ��": cbrControl.IconId = 3201
        With cbrControl.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_RectSave, "�ü�����"): cbrControl.ToolTipText = "�ü��ɼ�ͼ�񲢱���": cbrControl.IconId = 0
            Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_DirectSave, "ֱ�ӱ���"): cbrControl.ToolTipText = "���浱ǰ����ͼ��": cbrControl.IconId = 0
        End With
        
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_ImgPro_Rotate_Pop, "��ת"): cbrControl.IconId = 503
        With cbrControl.CommandBar.Controls
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_RRotate, "˳ʱ"): cbrControl.ToolTipText = "˳ʱ����ת": cbrControl.IconId = 503
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_LRotate, "��ʱ"): cbrControl.ToolTipText = "��ʱ����ת": cbrControl.IconId = 504
        End With
        
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_ImgPro_Smooth_Pop, "ƽ��"): cbrControl.IconId = 506
        With cbrControl.CommandBar.Controls
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Sharpness, "��"): cbrControl.ToolTipText = "��": cbrControl.IconId = 505
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Smooth, "ƽ��"): cbrControl.ToolTipText = "ƽ��": cbrControl.IconId = 506
        End With
        
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_ImgPro_Lab_Pop, "��ע"): cbrControl.ToolTipText = "��ע": cbrControl.IconId = 509
        With cbrControl.CommandBar.Controls
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Arrow, "��ͷ"): cbrControl.ToolTipText = "��ͷ��ע"": cbrControl.IconId = 507"
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Ellipse, "Բ��"): cbrControl.ToolTipText = "Բ�α�ע"": cbrControl.IconId = 508"
          Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_Text, "�ı�"): cbrControl.ToolTipText = "���ֱ�ע": cbrControl.IconId = 509
        End With
        
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIcon
        cbrControl.Category = "�ɼ�����"
        cbrControl.Enabled = False
    Next
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
        Set cbrControl = .Add(xtpControlButton, conMenu_ImgPro_RectCapture, "����")
        cbrControl.IconId = 0
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
     
    mintCaptureFlag = 1
    
    Call mobjCapHelper.SaveImg(imgResult, "")
End Sub

Private Sub StartTimer()
On Error GoTo errH
    mblnTimerState = Not mblnTimerState
    Call mVideoCapture.StartTimer(mblnTimerState)
    
    Exit Sub
errH:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub
 
