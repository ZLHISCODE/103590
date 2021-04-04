VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Begin VB.Form frmPetitionCapture 
   Caption         =   "���뵥ͼ��"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11940
   Icon            =   "frmPetitionCapture.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   11940
   StartUpPosition =   3  '����ȱʡ
   Begin ScanLibCtl.ImgScan ImageScanner 
      Left            =   0
      Top             =   6240
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
      Height          =   6735
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10695
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   480
         Top             =   6240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin DicomObjects.DicomViewer dcmMiniature 
         Height          =   4575
         Left            =   1800
         TabIndex        =   8
         ToolTipText     =   "˫��ѡ��ͼ�񣬶�ͼ����в�����"
         Top             =   120
         Width           =   7530
         _Version        =   262147
         _ExtentX        =   13282
         _ExtentY        =   8070
         _StockProps     =   35
         BackColor       =   -2147483642
      End
      Begin DicomObjects.DicomViewer dcmViewImg 
         Height          =   1575
         Left            =   120
         TabIndex        =   15
         Top             =   3120
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
      Begin DicomObjects.DicomViewer dcmView 
         Height          =   1575
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   2175
         _Version        =   262147
         _ExtentX        =   3836
         _ExtentY        =   2778
         _StockProps     =   35
         BackColor       =   0
         UseScrollBars   =   0   'False
      End
      Begin VB.PictureBox picTemp2 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   1695
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame fmeInfoCtrl 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   0
      TabIndex        =   0
      Top             =   6870
      Width           =   11895
      Begin VB.Frame FmeProcess 
         Height          =   1455
         Left            =   7200
         TabIndex        =   16
         Top             =   0
         Width           =   1215
         Begin VB.CommandButton cmdReduce 
            Caption         =   "��С"
            Height          =   330
            Left            =   180
            TabIndex        =   22
            Top             =   670
            Width           =   855
         End
         Begin VB.CommandButton cmdMagnify 
            Caption         =   "�Ŵ�"
            Height          =   330
            Left            =   180
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdReset 
            Caption         =   "��λ"
            Height          =   330
            Left            =   180
            TabIndex        =   17
            Top             =   1100
            Width           =   855
         End
      End
      Begin VB.Frame FmeScanbtn 
         Height          =   1455
         Left            =   8520
         TabIndex        =   6
         Top             =   0
         Width           =   3255
         Begin VB.CommandButton cmdChoiceEqu 
            Caption         =   "ѡ���豸(&C)"
            Height          =   375
            Left            =   1700
            TabIndex        =   12
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton cmdScanSet 
            Caption         =   "ɨ������(&E)"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton cmdDeleteImg 
            Caption         =   "ɾ��ͼ��(&D)"
            Height          =   375
            Left            =   1700
            TabIndex        =   10
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdScanImg 
            Caption         =   "ɨ��ͼ��(&S)"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fmePatientInfo 
         Height          =   1455
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   6975
         Begin VB.Label lblCheckNum 
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
            Height          =   255
            Left            =   480
            TabIndex        =   20
            Top             =   705
            Width           =   2535
         End
         Begin VB.Label lblPatientAge 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��    ��:"
            Height          =   180
            Left            =   5040
            TabIndex        =   19
            Top             =   240
            Width           =   1530
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
            Left            =   480
            TabIndex        =   18
            Top             =   1170
            Width           =   2565
         End
         Begin VB.Label lblPatientName 
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
            Height          =   255
            Left            =   480
            TabIndex        =   5
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label lblExamineMethod 
            Caption         =   "��鷽��:"
            Height          =   240
            Left            =   3120
            TabIndex        =   4
            Top             =   705
            Width           =   3765
         End
         Begin VB.Label lblSpePosition 
            Caption         =   "��鲿λ:"
            Height          =   240
            Left            =   3120
            TabIndex        =   3
            Top             =   1140
            Width           =   3735
         End
         Begin VB.Label lblPatientSex 
            Caption         =   "��    ��:"
            Height          =   255
            Left            =   3120
            TabIndex        =   2
            Top             =   240
            Width           =   1815
         End
      End
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

Private mstrTempDirOfScan As String          'ɨ�����ʱĿ¼
Private mstrScanDeviceTempDir As String      'ɨ���豸��ʱĿ¼
Private mstrBufferDir As String

Private mintScanImageIndex As Integer        'ɨ��ͼ������
Private mintCurImgIndex As Integer           '��ǰ��ѡ�е�ͼ������
Private mintShowPhotoNumber As Integer       '����ͼ����ʾ����


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

Public Sub ShowPetitionCaptureWind(ByVal strPrivs As String, lngCurDeptId As Long, strDeptName As String, _
                                   strPatientName As String, strPatientAge As String, strPatientSex As String, _
                                   strExamineMethod As String, strSpePosition As String, blnIsExamine As Boolean, _
                                   blnIsLogin As Boolean, Optional lngAdviceID As Long = 0, Optional intState As Integer = 0)
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

Private Sub cmdChoiceEqu_Click()
On Error GoTo errHandle

    '��ѡ��ɨ���豸����
    ImageScanner.ShowSelectScanner

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
    ImageScanner.ShowScanPreferences

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ScanImages()

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
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub dcmMiniature_DblClick()
On Error GoTo errHandle
    If mintCurImgIndex = 0 Then
        MsgBoxD Me, "�ò���û����ɨ������뵥��", vbInformation, gstrSysName
        Exit Sub
    End If
    
   '��ѡ�е�ͼ�񵥶����ص�dcmViewImg��ȥ����������
    Call LoadViewImg
    
    FmeProcess.Visible = True
    dcmMiniature.Visible = False
    dcmViewImg.Visible = True
    
    If dcmViewImg.Visible Then
        cmdChoiceEqu.Enabled = False
        cmdDeleteImg.Enabled = False
        cmdScanImg.Enabled = False
        cmdScanSet.Enabled = False
    End If
    
    Call Form_Resize

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadViewImg()
Dim ImgTmpImage As New DicomImage
    
    dcmViewImg.Images.Clear
    '��ͼ��ת����PicBox
    Set picTemp2.Picture = dcmMiniature.Images.Item(mintCurImgIndex).Picture
    '��ͼ���Ƶ�������
    Call Clipboard.SetData(picTemp2.Picture, 2)
'    �Ӽ��а�ȡ��ͼ��
    Call ImgTmpImage.Paste
    
    Call Clipboard.Clear
    '��ͼ���������ͼ��
    dcmViewImg.Images.Add ImgTmpImage
End Sub

Private Sub dcmViewImg_DblClick()
On Error GoTo errHandle

    dcmMiniature.Visible = True
    dcmViewImg.Visible = False
    FmeProcess.Visible = False
    
     '����״̬�� ���ܽ��в���
    If dcmViewImg.Visible = False And Not mblnIsExamine Then
        cmdChoiceEqu.Enabled = True
        cmdDeleteImg.Enabled = True
        cmdScanImg.Enabled = True
        cmdScanSet.Enabled = True
    End If
    
     Call Form_Resize
     
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
    If mblnIsExamine Then
        cmdChoiceEqu.Enabled = False
        cmdDeleteImg.Enabled = False
        cmdScanImg.Enabled = False
        cmdScanSet.Enabled = False
    End If
    
    '��ʼ������ ͼ��߼�����ť
    FmeProcess.Visible = False
    
    '�����豸��ʱĿ¼
    mstrScanDeviceTempDir = "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA"

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
    
    '�ݲ��رգ�ֻ�����ش���
    If mblnIsLogin Then
        Cancel = 1
        Me.Hide
    End If
    
    Exit Sub
errHandle:
    '�Ͽ�FTP����
    miNet.FuncFtpDisConnect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub InitLables()
'���ݴ����ֵ�����˻�����Ϣlbl��ֵ

    lblCheckNum.Caption = "�� �� �ţ�" & mstrCheckNo
    lblPatientDept.Caption = "���˿��ң�" & mstrDeptName
    lblPatientName.Caption = "��    ����" & mstrPatientName
    lblPatientAge.Caption = "��    �䣺" & mstrPatientAge
    lblPatientSex.Caption = "��    ��" & mstrPatientSex
    lblExamineMethod.Caption = "��鷽����" & mstrExamineMethod
    lblSpePosition.Caption = "��鲿λ��" & mstrSpePosition

End Sub

Public Sub InitDeptPara(ByVal lngDeptID As Long)
'��ʼ�����Ҽ�����
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo DBError
    mlngCurDeptId = lngDeptID
    
    '��ȡ�����洢�豸��
    mstrSaveDeviceID = GetDeptPara(mlngCurDeptId, "�洢�豸��")
    gstrSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�õ��豸��", mstrSaveDeviceID)
    If rsTmp.EOF Then
        MsgBox "Ӱ��洢�豸δ�������ͣ�ã����飡", vbInformation, gstrSysName
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

        If dcmView.Images.Count > 0 Then
            Set picTemp2.Picture = dcmView.CurrentImage.Capture(False).Picture
        End If
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

    mintCurImgIndex = dcmMiniature.Images.Count
End Sub


Public Sub subSaveImage(Optional iEncode As Integer = 0, Optional lngAdviceID As Long)
'------------------------------------------------
'���ܣ������һ������ͼ���浽���ݿ���
'������iEncode����ѹ����ʽ��1��Run-Length Encoding�г�ѹ����2������������ԭͼ��ѹ����ʽ��������Lossless JPEG encoding JPEG����ѹ��
'���أ���
'------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim ImgTmp As New DicomImage

    Dim strReceived As String
    Dim strFileTitle As String       'ͼ���ļ���ͷ
    Dim strResult As String         'FTP�������
    Dim nowTime As Date
    Dim blnInTrans As Boolean       '�����ﴦ�������
    Dim strRandom As String
    Dim i As Integer

    
    If dcmMiniature.Images.Count <= 0 Then
        Exit Sub
    End If
    
    '��ȡ���һ������ͼ
    Set ImgTmp = dcmMiniature.Images(dcmMiniature.Images.Count)

    '�õ������
    strRandom = CInt(Rnd * 100 + 1)

    nowTime = zlDatabase.Currentdate
    strFileTitle = Format(nowTime, "mmddhhmmss")
    strReceived = Format(nowTime, "yyyymmdd")
    
    '��������Ŀ¼
    MkLocalDir mstrBufferDir & strReceived & "/" & lngAdviceID & "/"

    '����ͼ�񵽻���Ŀ¼  Lossless JPEG encoding JPEG����ѹ��
    ImgTmp.WriteFile mstrBufferDir & strReceived & "/" & lngAdviceID & "/" & strFileTitle & lngAdviceID & strRandom, True

    ImgTmp.FileExport mstrBufferDir & strReceived & "/" & lngAdviceID & "/" & strFileTitle & lngAdviceID & strRandom & ".jpg", "JPG", 80

    ImgTmp.Tag = strFileTitle & lngAdviceID & strRandom & ".jpg"
    
    strResult = miNet.FuncFtpConnect(mFtpIp, mFtpUser, mFtpPass)

    If strResult = 0 Then
        'FTP����ʧ�ܣ���ʾ���󣬲�ɾ������ͼ�е�ͼ��
        MsgBoxD Me, "FTP����ʧ�ܣ�ͼ���޷����棬�����������á�", vbInformation, gstrSysName
        dcmMiniature.Images.Remove (i)
        Exit Sub
    End If

    '����ɨ�赥ͼ��
    WriteToURL mstrBufferDir & strReceived & "/" & lngAdviceID & "/" & strFileTitle & lngAdviceID & strRandom, mFtpDir & _
        strReceived & "/" & lngAdviceID & "/" & strFileTitle & lngAdviceID & strRandom

    miNet.FuncFtpDisConnect

    'ͼ��洢�ɹ��󣬴洢���ݿ���Ϣ
    On Error GoTo DBError
    
    '�����µ�ͼ���¼
    gstrSQL = "ZL_Ӱ�����뵥ͼ��_INSERT ('" & lngAdviceID & "','" & strFileTitle & lngAdviceID & strRandom & ".jpg" & "','" & strReceived & "/" & lngAdviceID & "','" & mstrSaveDeviceID & "','" & UserInfo.���� & "',sysdate)"

    '����ͼ��
    Call zlDatabase.ExecuteProcedure(CStr(gstrSQL), "����ͼ��")
    
    '���mblnIsLogin=true ��ô˵�������ڵǼǽ���ı���ͼ����Ҫ���ò�������Ϊfalse
    If mblnIsLogin Then
        mblnIsLogin = False
    End If

    Exit Sub
DBError:
    '�Ͽ�FTP����
    miNet.FuncFtpDisConnect
    '������������ݿ����������ɾ�����ɼ���ͼ��
    err.Raise err.Number, "���ͼ�񱣴�"
    dcmMiniature.Images.Remove (dcmMiniature.Images.Count)
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
    
    On Error GoTo DBError

       strSql = "select ���뵥ͼ��,ɨ����,ɨ��ʱ��,FTP·��,�豸�� from Ӱ�����뵥ͼ�� where ҽ��ID=[1] order by �豸��"
       Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���뵥ͼ����Ϣ", lngAdviceID)

        'DcmViewer.Images.Clear
        If rsTmp.RecordCount > 0 Then
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
                        curImage.Tag = Nvl(rsTmp("���뵥ͼ��"))
                        
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
            If dcmViewer.Images.Count > 0 Then
                dcmViewer.CurrentIndex = 1
                dcmViewer.Images(dcmViewer.Images.Count).BorderColour = vbRed
            End If
        Else
            dcmViewer.MultiColumns = 1
            dcmViewer.MultiRows = 1
        End If
    Inet1.FuncFtpDisConnect
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
Dim blnResult As Boolean
    If mintCurImgIndex > 0 And mintCurImgIndex <= dcmMiniature.Images.Count Then
        
        '�����ݿ��FTP��ɾ������ͼ�б�ѡ�е�ͼ��
        blnResult = DelPetitionImg()
        
        If blnResult = True Then    'ɾ���ɹ������޸�����ͼ״̬��������StateChanged�¼�
            '������ͼ��ɾ��ͼ��
            dcmMiniature.Images.Remove mintCurImgIndex
            dcmView.Images.Clear
            
            If mintCurImgIndex > dcmMiniature.Images.Count Then
                mintCurImgIndex = dcmMiniature.Images.Count
            End If

            If mintCurImgIndex > 0 Then
                dcmMiniature.Images(mintCurImgIndex).BorderColour = vbRed
            End If
            
            Call fmeDcmViewer_Resize
        End If
    End If
End Sub

Private Sub subDeleteDcmImage()

'������ͼ��ɾ��ͼ��
        dcmMiniature.Images.Remove mintCurImgIndex
        dcmView.Images.Clear
        
        If mintCurImgIndex > dcmMiniature.Images.Count Then
            mintCurImgIndex = dcmMiniature.Images.Count
        End If

        If mintCurImgIndex > 0 Then
            dcmMiniature.Images(mintCurImgIndex).BorderColour = vbRed
        End If
        
        Call fmeDcmViewer_Resize

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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���뵥ͼ����Ϣ", mlngAdviceID, ImgTmp.Tag)
    
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

    gstrSQL = "ZL_Ӱ�����뵥ͼ��_DELETE(" & mlngAdviceID & ",'" & ImgTmp.Tag & "')"

    zlDatabase.ExecuteProcedure gstrSQL, "Ӱ��ͼ��ɾ��"

    'ɾ��ͼ���ļ�
    RemoveFromURL strFtpDir & _
        Format(Nvl(rsTmp("ɨ��ʱ��"), CStr(zlDatabase.Currentdate)), "yyyyMMdd") & "/" & _
        mlngAdviceID & "/" & Mid(ImgTmp.Tag, 1, InStr(ImgTmp.Tag, ".") - 1)

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


Private Sub WriteToURL(ByVal strFileName As String, ByVal strDestFileName As String)
'------------------------------------------------
'���ܣ��������ļ����浽Զ��������
'������strFileName--�����ļ�����strDestFileName����Զ���ļ���
'���أ���
'-----------------------------------------------
'���ܣ�
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strPath As String
    
    '��FTP�д���Ŀ¼
    strPath = objFileSystem.GetParentFolderName(strDestFileName)
    miNet.FuncFtpMkDir "/", strPath
    
    '��FTP�ϴ��ļ�
    miNet.FuncUploadFile strPath, strFileName, objFileSystem.GetFileName(strDestFileName)
End Sub


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

    If Button = 1 Then
        mCorpSize.X = 0
        mCorpSize.Y = 0
        
        '��ѡ��ͼ����ʾ���
        i = dcmMiniature.ImageIndex(X, Y)
        If i <> 0 Then
        
            If mintCurImgIndex > 0 And mintCurImgIndex <= dcmMiniature.Images.Count Then
                dcmMiniature.Images(mintCurImgIndex).BorderColour = vbWhite
            End If
            dcmMiniature.Images(i).BorderColour = vbRed
            mintCurImgIndex = i
            
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub




Private Sub Form_Resize()
On Error Resume Next
    fmeDcmViewer.Top = 0
    fmeDcmViewer.Left = 0
    fmeDcmViewer.Width = Me.ScaleWidth
    fmeDcmViewer.Height = Me.ScaleHeight - fmeInfoCtrl.Height

    dcmMiniature.Top = 60
    dcmMiniature.Left = 60
    dcmMiniature.Width = fmeDcmViewer.Width - 120
    dcmMiniature.Height = fmeDcmViewer.Height
    
    dcmViewImg.Top = 60
    dcmViewImg.Left = 60
    dcmViewImg.Width = fmeDcmViewer.Width - 120
    dcmViewImg.Height = fmeDcmViewer.Height


    fmeInfoCtrl.Top = fmeDcmViewer.Height
    fmeInfoCtrl.Left = 0
    fmeInfoCtrl.Width = fmeDcmViewer.Width

    fmePatientInfo.Top = 0
    fmePatientInfo.Left = 60
    fmePatientInfo.Width = fmeInfoCtrl.Width - FmeScanbtn.Width - IIf(FmeProcess.Visible, FmeProcess.Width, -60) - 220
    fmePatientInfo.Height = fmeInfoCtrl.Height

    FmeProcess.Top = 0
    FmeProcess.Left = fmePatientInfo.Width + 120
    FmeProcess.Height = fmeInfoCtrl.Height

    FmeScanbtn.Top = 0
    FmeScanbtn.Left = fmePatientInfo.Width + IIf(FmeProcess.Visible, FmeProcess.Width, -60) + 180
    FmeScanbtn.Height = fmeInfoCtrl.Height

End Sub


Private Sub fmeDcmViewer_Resize()
    Dim iRows As Integer
    Dim iCols As Integer
    
    On Error Resume Next
    
    dcmMiniature.Left = 0
    dcmMiniature.Top = 0
    dcmMiniature.Width = fmeDcmViewer.Width
    dcmMiniature.Height = fmeDcmViewer.Height
    
    dcmViewImg.Top = 60
    dcmViewImg.Left = 60
    dcmViewImg.Width = fmeDcmViewer.Width - 120
    dcmViewImg.Height = fmeDcmViewer.Height
    
    '�Զ���ͼ��������
    '��������ͼ��ͼ�񲼾�
    ResizeRegion dcmMiniature.Images.Count, dcmMiniature.Width, dcmMiniature.Height, iRows, iCols
    
    dcmMiniature.MultiColumns = iCols
    dcmMiniature.MultiRows = iRows

End Sub


