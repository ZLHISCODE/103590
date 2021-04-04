VERSION 5.00
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVIDEOPROCESS.OCX"
Begin VB.Form frmPatiImageGatherer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ͼ��ɼ�"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
   Icon            =   "frmPatiImageGatherer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "����"
      Height          =   350
      Index           =   3
      Left            =   5940
      TabIndex        =   6
      Top             =   930
      Width           =   950
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "ȷ��"
      Height          =   350
      Index           =   2
      Left            =   5940
      TabIndex        =   5
      Top             =   510
      Width           =   950
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "�ɼ�"
      Height          =   350
      Index           =   0
      Left            =   5940
      TabIndex        =   4
      Top             =   90
      Width           =   950
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "��Ƶ����"
      Height          =   350
      Index           =   4
      Left            =   5940
      TabIndex        =   3
      Top             =   1350
      Width           =   950
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   0
      ScaleHeight     =   4485
      ScaleWidth      =   5445
      TabIndex        =   0
      Top             =   0
      Width           =   5475
      Begin ZLDSVideoProcess.DSCapture DSCap 
         Height          =   4000
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4000
         Object.Visible         =   -1  'True
         AutoScroll      =   0   'False
         AutoSize        =   0   'False
         AxBorderStyle   =   2
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
         CurWidth        =   267
         CurHeight       =   267
         CurVideoWidth   =   263
         CurVideoHeight  =   245
         ShowModel       =   0
         CapParameterWindPos=   8
         SnatchWay       =   0
         ParameterCfgFileName=   ""
         HideCfgItem     =   0
         AppHandle       =   0
         Begin VB.PictureBox picRect 
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            Height          =   2055
            Left            =   1170
            ScaleHeight     =   2055
            ScaleWidth      =   2385
            TabIndex        =   2
            Top             =   750
            Visible         =   0   'False
            Width           =   2385
         End
      End
   End
   Begin VB.Image imgPerson 
      BorderStyle     =   1  'Fixed Single
      Height          =   1425
      Left            =   5490
      Stretch         =   -1  'True
      Top             =   3090
      Width           =   1800
   End
   Begin VB.Label lblShowInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "ͼ��Ԥ����"
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   2820
      Width           =   915
   End
End
Attribute VB_Name = "frmPatiImageGatherer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnNotLoad As Boolean
Private mstrPictureFilePath As String
Private mblnOk As Boolean '�Ƿ�ɼ��ɹ�
Private Enum Enum_Button
    EM_FUN_�ɼ� = 0
    EM_FUN_ȷ�� = 2
    EM_FUN_���� = 3
    EM_FUN_��Ƶ���� = 4
End Enum
Private mlngButton As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Const SRCCOPY = &HCC0020

Public Function ShowMe(ByVal frmParent As Form, ByRef strPictureFilePath As String) As Boolean
    '-----------------------------------------------
    '���ܣ�������ڣ���ʾͼ��ɼ�����
    '������
    '   frmParent��������
    '   strPictureFilePath���ɼ�ͼƬ����λ��
    '���أ������Ƿ�ɼ��ɹ�
    '���ƣ�Ƚ����
    '���ڣ�2014-6-26
    '-----------------------------------------------
    mstrPictureFilePath = App.Path & "\person.bmp"
    mblnOk = False
    Me.Show 1, frmParent
    strPictureFilePath = mstrPictureFilePath
    ShowMe = mblnOk
End Function

Private Sub DSCap_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    mlngButton = 2
End Sub

Private Sub DSCap_OnMouseMove(ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    Dim blnMouseIn As Boolean '����Ƿ���picRect��
    
    If Not picRect.Visible Then Exit Sub
    
    X = ScaleX(X, vbPixels, vbTwips)
    Y = ScaleX(Y, vbPixels, vbTwips)
    blnMouseIn = picRect.Left < X And X <= picRect.Left + picRect.Width And picRect.Top < Y And Y <= picRect.Top + picRect.Height
    If blnMouseIn Then
        SetCapture picRect.hWnd
        picRect.MousePointer = vbSizePointer
        If mlngButton = vbLeftButton Then
            X = ReleaseCapture()
            SendMessage picRect.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
        End If
    Else
        picRect.MousePointer = vbDefault
        ReleaseCapture
    End If
    picRect.Refresh
End Sub

Private Sub DSCap_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    mlngButton = 0
End Sub

Private Sub Form_Resize()
    DSCap.Width = picBack.Width
    DSCap.Height = picBack.Height
End Sub

Private Sub picRect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnMouseIn As Boolean '����Ƿ���picRect��
    If Not picRect.Visible Then Exit Sub

    blnMouseIn = (0 < X) And (X <= ScaleX(picRect.Width, vbTwips, vbPixels)) And (0 < Y) And (Y <= ScaleX(picRect.Height, vbTwips, vbPixels))
    If blnMouseIn Then
        SetCapture picRect.hWnd
        picRect.MousePointer = vbSizePointer
        If Button = vbLeftButton Then
            X = ReleaseCapture()
            SendMessage picRect.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
        End If
    Else
        picRect.MousePointer = vbDefault
        ReleaseCapture
    End If
    picRect.Refresh
End Sub

Private Sub Form_Load()
    Dim strErr As String
    picRect.ScaleMode = vbPixels
    Call SetRectangle '����һ��͸�����ο�
    cmdButton(EM_FUN_ȷ��).Enabled = False
    DSCap.ReadParameterFromFile
    '����Ԥ��ģʽ
    strErr = DSCap.StartPreview
    If strErr <> "" Then GoTo Errhand
    Exit Sub
Errhand:
    MsgBox strErr, vbExclamation, gstrSysName
    mblnNotLoad = True
    Exit Sub
End Sub

Private Sub Form_Activate()
    If mblnNotLoad Then Unload Me
End Sub

Private Sub cmdButton_Click(Index As Integer)
    Dim strErr As String, strFile As String

    Select Case Index
    Case EM_FUN_�ɼ�
        '���ɼ�ͼ�󱣴�ΪBMP�ļ�
        strErr = DSCap.CaptureBmpImageToFile(mstrPictureFilePath)
        If strErr <> "" Then GoTo Errhand
        '��ʾ�ɼ�ͼ��
        imgPerson.Picture = DSCap.CaptureBmpImage
        cmdButton(EM_FUN_ȷ��).Enabled = True
        DSCap.ShowModel = smAutoFitCut
    Case EM_FUN_ȷ��
        mblnOk = True: Unload Me
    Case EM_FUN_����
        Unload Me
    Case EM_FUN_��Ƶ����
        '��ʾ�ɼ��������öԻ���
        strErr = DSCap.ShowCaptureParameterCfgDialog(Me.hWnd)
        If strErr <> "" Then GoTo Errhand
        '���½���Ԥ��ģʽ
        Call DSCap.RefreshWindow 'Ӧ������
        Call DSCap.RePreview
    End Select
    Exit Sub
Errhand:
    MsgBox strErr, vbExclamation, gstrSysName
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '�ͷ���Դ
    Call DSCap.FreeRes
End Sub

Private Sub SetRectangle()
    '--------------------------------------------
    '���ܣ�����һ��͸�����ο�
    '--------------------------------------------
    Const RGN_DIFF = 4
    Dim lngOuterRgn As Long, lngInnerRgn As Long, lngCombinedRgn As Long
    Dim sinWidth As Single, sinHeight As Single
    Dim sinBorderWidth As Single, sinTitleHeight As Single
    If Not picRect.Visible Then Exit Sub


    '��ȡ�ؼ���Ⱥ͸߶�
    sinWidth = ScaleX(picRect.Width, vbTwips, vbPixels)
    sinHeight = ScaleY(picRect.Height, vbTwips, vbPixels)
    '����ο�
    lngOuterRgn = CreateRectRgn(0, 0, sinWidth, sinHeight)
    '�����ھ��ο��Ⱥ͸߶�
    sinBorderWidth = (sinWidth - picRect.ScaleWidth) / 2
    sinTitleHeight = sinHeight - sinBorderWidth - picRect.ScaleHeight
    '�ھ��ο�
    lngInnerRgn = CreateRectRgn(sinBorderWidth + 3, sinTitleHeight + 3, picRect.ScaleWidth - 3, picRect.ScaleHeight - 3)
    '�Ӵ�����ȥ������������������
    lngCombinedRgn = CreateRectRgn(0, 0, 0, 0)
    CombineRgn lngCombinedRgn, lngOuterRgn, lngInnerRgn, RGN_DIFF
    '���ƴ��ڵ�����
    SetWindowRgn picRect.hWnd, lngCombinedRgn, False
End Sub

