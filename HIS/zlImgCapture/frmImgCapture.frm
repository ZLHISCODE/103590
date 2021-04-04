VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.2#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImgCapture 
   Caption         =   "图像采集"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   ControlBox      =   0   'False
   Icon            =   "frmImgCapture.frx":0000
   ScaleHeight     =   6465
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrComm 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7560
      Top             =   3840
   End
   Begin VB.PictureBox PicPar 
      Height          =   2865
      Left            =   990
      ScaleHeight     =   2805
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   720
      Width           =   4935
      Begin VB.PictureBox PicCli 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   2145
         Left            =   600
         ScaleHeight     =   2145
         ScaleWidth      =   4185
         TabIndex        =   1
         Top             =   240
         Width           =   4185
      End
      Begin VB.PictureBox PicVideo 
         BackColor       =   &H80000007&
         Height          =   1635
         Left            =   360
         ScaleHeight     =   1575
         ScaleWidth      =   3825
         TabIndex        =   11
         Top             =   990
         Width           =   3885
      End
      Begin VB.PictureBox PicTmp1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000D&
         Height          =   675
         Left            =   1050
         ScaleHeight     =   615
         ScaleWidth      =   825
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.PictureBox PicTmp2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000D&
         Height          =   645
         Left            =   0
         ScaleHeight     =   585
         ScaleWidth      =   855
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin MSComDlg.CommonDialog Comm 
      Left            =   7530
      Top             =   2130
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtState 
      Height          =   270
      Left            =   30
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   750
      Visible         =   0   'False
      Width           =   7845
   End
   Begin MCI.MMControl MMControl 
      Height          =   330
      Left            =   900
      TabIndex        =   13
      Top             =   4290
      Visible         =   0   'False
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   582
      _Version        =   393216
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      AutoEnable      =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6
      Left            =   6300
      Top             =   900
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   285
      Left            =   900
      TabIndex        =   12
      Top             =   3930
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   503
      _Version        =   393216
      LargeChange     =   10
      Max             =   100
      TickStyle       =   3
      TextPosition    =   1
   End
   Begin VB.PictureBox PicY 
      Height          =   50
      Left            =   570
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   5625
      TabIndex        =   8
      Top             =   4440
      Width           =   5625
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   9660
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   7200
      FixedBackground1=   0   'False
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   9540
         _ExtentX        =   16828
         _ExtentY        =   1138
         ButtonWidth     =   1138
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "显示"
               Key             =   "显示"
               Object.ToolTipText     =   "显示"
               Object.Tag             =   "显示"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "报告"
               Key             =   "报告"
               Object.ToolTipText     =   "报告"
               Object.Tag             =   "报告"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "采集"
               Key             =   "采集"
               Object.ToolTipText     =   "采集"
               Object.Tag             =   "采集"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "导入"
               Key             =   "导入"
               Object.ToolTipText     =   "导入"
               Object.Tag             =   "导入"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存"
               Key             =   "保存"
               Object.ToolTipText     =   "保存"
               Object.Tag             =   "保存"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "split1"
               Key             =   "split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "录像"
               Key             =   "录像"
               Object.ToolTipText     =   "录像"
               Object.Tag             =   "录像"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "播放"
               Key             =   "播放"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "暂停"
               Key             =   "暂停"
               Object.ToolTipText     =   "暂停"
               Object.Tag             =   "暂停"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "快进"
               Key             =   "快进"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "split2"
               Key             =   "split2"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "设置"
               Key             =   "设置"
               Object.ToolTipText     =   "设置"
               Object.Tag             =   "设置"
               ImageIndex      =   9
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "格式"
                     Text            =   "格式"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "来源"
                     Text            =   "来源"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "压缩方式"
                     Text            =   "压缩方式"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "端口"
               Key             =   "端口"
               Object.ToolTipText     =   "端口"
               Object.Tag             =   "端口"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "split2"
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox PicDicom 
      Height          =   1335
      Left            =   270
      ScaleHeight     =   1275
      ScaleWidth      =   6585
      TabIndex        =   3
      Top             =   4740
      Width           =   6645
      Begin DicomObjects.DicomViewer DicomViewer 
         Height          =   1125
         Left            =   150
         TabIndex        =   4
         Top             =   120
         Width           =   1455
         _Version        =   262146
         _ExtentX        =   2566
         _ExtentY        =   1984
         _StockProps     =   35
         BackColor       =   -2147483635
         BorderStyle     =   1
         MultiColumns    =   3
      End
   End
   Begin VB.PictureBox PicX2 
      BackColor       =   &H80000007&
      Height          =   2955
      Left            =   6030
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2955
      ScaleWidth      =   45
      TabIndex        =   7
      Top             =   810
      Width           =   50
   End
   Begin VB.PictureBox PicX1 
      BackColor       =   &H80000007&
      Height          =   2865
      Left            =   870
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2865
      ScaleWidth      =   45
      TabIndex        =   6
      Top             =   870
      Width           =   50
   End
   Begin VB.PictureBox PicY1 
      BackColor       =   &H80000007&
      Height          =   50
      Left            =   1080
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   4755
      TabIndex        =   5
      Top             =   780
      Width           =   4755
   End
   Begin VB.PictureBox PicY2 
      BackColor       =   &H80000007&
      Height          =   50
      Left            =   1050
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   4875
      TabIndex        =   2
      Top             =   3840
      Width           =   4875
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7320
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImgCapture.frx":17DA
            Key             =   "显示"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImgCapture.frx":2FC4
            Key             =   "采集"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImgCapture.frx":47AE
            Key             =   "导入"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImgCapture.frx":5088
            Key             =   "保存"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImgCapture.frx":6872
            Key             =   "删除"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImgCapture.frx":805C
            Key             =   "录像"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImgCapture.frx":9846
            Key             =   "播放"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImgCapture.frx":B030
            Key             =   "暂停"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImgCapture.frx":C81A
            Key             =   "设置"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImgCapture.frx":E004
            Key             =   "端口"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImgCapture.frx":F7EE
            Key             =   "退出"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImgCapture.frx":10FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImgCapture.frx":11752
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   6105
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   1764
            Picture         =   "frmImgCapture.frx":1202C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11959
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmImgCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1


Const SWP_NOACTIVATE = &H10
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_SHOWWINDOW = &H40

Dim blMoveDown  As Boolean                                  '用于判断是否按下鼠标左键
Dim lngCliWinTop As Long, lngCliWinLeft As Long             '采集子窗体位置
Dim lngParWinWidth As Long, lngParWinHeight As Long         '采集父窗体大小
Dim lngDicomImageHeight As Long                             '采集图像高度
Dim intSelectImage As Integer                               '选中的图像
Dim strTmpFileName As String                                '保存临时路径和文件名
Dim blVideoState As Boolean                                 '录像状态False播放True
Dim lngSpeedPaly As Integer                                 '快进和快退的速度
Dim blSaveMessage As Boolean                                '退出是如有图像没有保存提示
Dim lngReportWidth As Long                                  '报告时窗体宽
Dim lngReportHeight As Long                                 '报告时窗体高
Dim lngReportTop As Long                                    '报告时窗体X位置
Dim lngReportLeft As Long                                   '报告时窗体Y位置
Dim lngReportWinWidth As Long                               '报告窗体宽
Dim lngReportWinHeight As Long                              '报告窗体高
Dim lngReportWinTop As Long                                 '报告窗体Y
Dim lngReportWinLeft As Long                                '报告窗体X
Dim lngCaptureWidth As Long                                 '图像显示窗体宽度
Dim lngCaptureHeight As Long                                '图像显示窗体高度

Dim intComInterval As Integer                               '脚踏采图的时间间隔，单位秒
Dim intCapType As Integer                                   '脚踏触发方式，0-直接触发，1-变换触发
Dim intComState As Integer                                  'COM口的状态
Dim lngComTime As Long                                      '记录com口保持状态的时间
Dim mstrPrivs As String                                    '记录权限

Dim dcmglbUID As New DicomGlobal                            'DICOM全局对象，用来创建新的UID

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Private lScrnOffset As Long
Private iCurImageIndex As Long
Private strPatientID As String, strStudyUID As String, strImgType As String, strSeriesID As String
Private mstrNO As String, mint记录性质 As Integer, mlng单据ID As Long, mlng病历ID As Long, mstr医嘱内容 As String
Private WithEvents mfrmRepEdit As Form
Attribute mfrmRepEdit.VB_VarHelpID = -1
Private mfrmPacsWork As Form
Private lngDeviceNO As String
Private dtLastCapture As Date '最近脚踏踩下的时间

Private MultiImages As New DicomImages
Private strCachePath As String

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Type DlgFileInfo
    iCount As Long
    sPath As String
    sFile() As String
End Type

Private intImgIndex As Integer                      '当前显示的图像index
Private mstrFormMode As String

Private Sub DicomViewer_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    With DicomViewer
        i = .ImageIndex(x, y)
        If i > 0 And i <= .Images.count And i <> iCurImageIndex Then
            .Images(iCurImageIndex).BorderColour = vbWhite
            .Images(i).BorderColour = vbRed
            iCurImageIndex = i
        End If
    End With
End Sub

Private Sub DicomViewer_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim MouseExit As Boolean
    Dim i As Integer
    Dim lngX As Long, lngY As Long
    MouseExit = (0 <= x * Screen.TwipsPerPixelX) And (x * Screen.TwipsPerPixelX <= _
    Me.DicomViewer.Width) And (0 <= y * Screen.TwipsPerPixelY) And (y * Screen.TwipsPerPixelY <= Me.DicomViewer.Height)
    If MouseExit Then
        SetCapture Me.DicomViewer.hwnd
        i = DicomViewer.ImageIndex(x, y)
        If i <> 0 Then
            If Me.tbrThis.Buttons("报告").Caption = "恢复" Then
                If Me.Width > Me.Height Then
                    '横放
                    lngX = x * Screen.TwipsPerPixelX + Me.Left
                    If Me.Top > Screen.Height / 2 Then
                        '放在下方
                        lngY = Me.Top + Me.tbrThis.Height - 550 - Me.PicPar.Height
                    Else
                        '放在上方
                        lngY = Me.Top + Me.Height
                    End If
                Else
                    '竖放
                    lngY = y * Screen.TwipsPerPixelY + Me.Top
                    If lngY + Me.PicCli.Height > Screen.Height - 500 Then
                        lngY = lngY - Abs(Screen.Height - (lngY + Me.PicCli.Height)) - 500
                    End If
                    
                    If Me.Left > Screen.Width / 2 Then
                        '靠右
                        lngX = Me.Left - Me.PicCli.Width
                    Else
                        '靠左
                        lngX = Me.Left + Me.Width
                    End If
                End If
            Else
                lngX = x * Screen.TwipsPerPixelX + Me.Left
                If lngX + Me.PicPar.Width > Screen.Width - 200 Then
                    lngX = Screen.Width - 200 - Me.PicPar.Width
                End If
                lngY = Me.Top + Me.Height - Me.stbThis.Height - Me.PicDicom.Height - Me.PicPar.Height
                
            End If
            If intImgIndex <> i Then
                intImgIndex = i
                frmImgShow.ShowMe DicomViewer.Images(i), Me, lngX, lngY
            End If
        Else
            intImgIndex = 0
            frmImgShow.Hide
        End If
    Else
        ReleaseCapture
        intImgIndex = 0
        frmImgShow.HideMe
    End If
End Sub

Private Sub Form_Load()
    Dim CaptureWinSize As CAPSTATUS
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim i As Integer
    Dim lngWidth As Long, lngHeight As Long
    Dim iGetSavedUI As Integer
    
    On Error GoTo OpenDriverEror
    
    '初始化
    i = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "Drivers", 0)
    mConnCapDevice Me.PicCli.hwnd, i

    CaptureWinSize = mGetCaptureWindowStatus

    Me.PicCli.Width = CaptureWinSize.uiImageWidth * Screen.TwipsPerPixelX
    Me.PicCli.Height = CaptureWinSize.uiImageHeight * Screen.TwipsPerPixelY
    
    
    
    lngParWinWidth = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ParWinWidth", Me.PicCli.Width)
    lngParWinHeight = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ParWinHeight", Me.PicCli.Height)
    lngCliWinTop = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "CliWinTop", 0)
    lngCliWinLeft = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "CliWinLeft", 0)
    lngDicomImageHeight = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "DicomImageHeight", Me.PicY.Top)
    
    iGetSavedUI = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "保存用户界面", 0)
    If iGetSavedUI = 0 Then
        mstrFormMode = "报告"
    Else
        'strMode = "恢复"
        mstrFormMode = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "Reprot_Mode", "报告")
    End If
    If mstrFormMode = "恢复" Then
        lngReportWidth = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "Reprot_frmWidth", 0)
        lngReportHeight = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "Reprot_frmHeight", 0)
        lngReportTop = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "Reprot_frmTop", 0)
        lngReportLeft = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "Reprot_frmLeft", 0)
        
        lngReportWinWidth = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ReprotWindow_frmWidth", 0)
        lngReportWinHeight = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ReprotWindow_frmHeight", 0)
        lngReportWinTop = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ReprotWindow_frmTop", 0)
        lngReportWinLeft = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ReprotWindow_frmLeft", 0)
    End If
    
    strCachePath = App.Path & "\TmpImage\"
    If Not objFileSystem.FolderExists(strCachePath) Then objFileSystem.CreateFolder strCachePath

    strTmpFileName = IIf(Len(App.Path) > 3, App.Path & "\TmpVideo.avi", App.Path & "TmpVideo.avi")
    
    InitPara
    Dim ret As Long
    If App.LogMode <> 0 Then
        '记录原来的window程序地址
        preWinProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
        '用自定义程序代替原来的window程序
        ret = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf Wndproc)
        capDlgVideoCompression hCapWnd
        If hCapWnd <> 0 Then
            Call capSetCallbackOnStatus(hCapWnd, AddressOf StatusProc)
        End If
    End If
       
    '将程序中的UID设置成1
    dcmglbUID.RegString("UIDRoot") = "1"
    Exit Sub
OpenDriverEror:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Resize()
    Dim iRows As Integer, iCols As Integer

    On Error Resume Next

'    If Me.tbrThis.Buttons("隐藏").Caption = "隐藏" Then
        With Me.PicCli
            .Top = lngCliWinTop
            .Left = lngCliWinLeft
        End With
        
        With Me.PicVideo
            .Top = Me.PicCli.Top
            .Left = Me.PicCli.Left
            .Width = Me.PicCli.Width
            .Height = Me.PicCli.Height
        End With
    
        With Me.PicPar
            If Me.tbrThis.Buttons("报告").Caption = "报告" Then
                If lngParWinWidth <> 0 And lngParWinHeight <> 0 Then
                    .Width = lngParWinWidth
                    .Height = lngParWinHeight
                Else
                    .Width = Me.PicCli.Width
                    .Height = Me.PicCli.Height
                End If
        
                .Top = 760
                
                If .Width >= Me.ScaleWidth Then
                    .Left = Me.PicX1.Width
                Else
                    .Left = (Me.ScaleWidth - .Width) / 2
                End If
                
                If lngDicomImageHeight > 0 Then
                    If lngDicomImageHeight <= CLng(.Top + .Height) Then
                        Me.PicY.Top = .Top + .Height + 100
                    Else
                        If lngDicomImageHeight >= Me.Height - Me.tbrThis.Height Then
                            Me.PicY.Top = Me.ScaleHeight - 660 - 1000
                        Else
                            Me.PicY.Top = lngDicomImageHeight
                        End If
                    End If
                    Me.PicY.Width = Me.ScaleWidth
                Else
                    Me.PicY.Top = Me.Height - Me.tbrThis.Height - 1000
                End If
                
                .Top = (Me.PicY.Top - .Height + 660) / 2
                If .Top < 760 Then .Top = 760
            Else
                .Top = 760
                .Left = 0
                .Width = Me.ScaleWidth
                .Height = Me.ScaleWidth * 576 / 768
                
                Me.PicY.Top = .Top + .Height + 100
                
                Me.PicY.Width = Me.ScaleWidth
                
                .Top = (Me.PicY.Top - .Height + 660) / 2
                If .Top < 760 Then .Top = 760
                
            End If
            
        End With
        
        'If Me.tbrThis.Buttons("报告").Caption = "报告" Then
            With Me.PicDicom
                .Top = PicY.Top + Me.PicY.Height - 10
                .Left = 0
                .Width = Me.ScaleWidth
                .Height = Me.ScaleHeight - PicY.Top - stbThis.Height
            End With
'        Else
'            With Me.PicDicom
'                .Top = cbrThis.Height + 10
'                .Left = 0
'                .Width = Me.ScaleWidth
'                .Height = Me.ScaleHeight - cbrThis.Height
'            End With
'        End If
        With Me.DicomViewer
            .Top = 0
            .Left = 0
            .Width = Me.PicDicom.ScaleWidth
            .Height = Me.PicDicom.ScaleHeight
        End With
        
        With Me.PicX1
            .Top = Me.PicPar.Top - Me.PicY1.Height
            .Left = Me.PicPar.Left - .Width
            .Height = Me.PicPar.Height + Me.PicY1.Height * 2 - 10
        End With
    
        With Me.PicX2
            .Top = Me.PicPar.Top - Me.PicY1.Height
            .Left = Me.PicPar.Left + Me.PicPar.Width
            .Height = Me.PicPar.Height + Me.PicY1.Height * 2 - 10
        End With
    
        With Me.PicY1
            .Top = Me.PicPar.Top - .Height
            .Left = Me.PicPar.Left - Me.PicX1.Width
            .Width = Me.PicPar.Width + Me.PicX1.Width * 2 - 10
        End With
    
        With Me.PicY2
            .Top = Me.PicPar.Top + Me.PicPar.Height
            .Left = Me.PicPar.Left - Me.PicX1.Width
            .Width = Me.PicPar.Width + Me.PicX1.Width * 2 - 10
        End With
        
        With Me.Slider1
            .Top = Me.PicY2.Top + Me.PicY2.Height
            .Left = Me.PicY2.Left
            .Width = Me.PicY2.Width
        End With
'    Else
'        With Me.PicDicom
'            .Top = cbrThis.Height + 10
'            .Left = 0
'            .Width = Me.ScaleWidth
'            .Height = Me.ScaleHeight - cbrThis.Height
'        End With
'
'        With Me.DicomViewer
'            .Top = 0
'            .Left = 0
'            .Width = Me.PicDicom.ScaleWidth
'            .Height = Me.PicDicom.ScaleHeight
'        End With
'    End If
    ResizeRegion Me.DicomViewer.Images.count + 1, Me.DicomViewer.Width, Me.DicomViewer.Height, iRows, iCols
    Me.DicomViewer.MultiColumns = iCols
    Me.DicomViewer.MultiRows = iRows
    
    If Me.tbrThis.Buttons("报告").Caption = "报告" Then
        '设置成不缩放
        mResizeCaptureWindow
        capSetScale hCapWnd, False
        
    Else
        '设置成缩放
        Call SetWindowPos(hCapWnd, _
            0&, _
            0&, _
            0&, _
            Me.PicPar.Width / Screen.TwipsPerPixelX, _
            Me.PicPar.Height / Screen.TwipsPerPixelY, _
            SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSENDCHANGING)
        capSetScale hCapWnd, True
        Me.PicCli.Left = 0
        Me.PicCli.Top = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim BITCapTureInfo As BITMAPINFO
    
'    If blSaveMessage = True Then
'        If MsgBox("还有图像没有保存,退出后图像不可恢复." & vbCrLf & _
'                 "是否退出?", vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
'            Cancel = True
'            Exit Sub
'        End If
'    End If
    
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "Reprot_Mode", Me.tbrThis.Buttons("报告").Caption
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "Reprot_frmWidth", Me.Width
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "Reprot_frmHeight", Me.Height
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "Reprot_frmTop", Me.Top
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "Reprot_frmLeft", Me.Left

    
    If Me.tbrThis.Buttons("报告").Caption = "报告" Then
        SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ParWinWidth", Me.PicPar.Width
        SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ParWinHeight", Me.PicPar.Height
        SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "CliWinTop", Me.PicCli.Top
        SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "CliWinLeft", Me.PicCli.Left
    End If
    
    If Me.tbrThis.Buttons("报告").Caption = "报告" Then
        SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "DicomImageHeight", Me.PicDicom.Top               'Me.PicDicom.Height + Me.stbThis.Height
    End If
    
    SendMessage hCapWnd, WM_CAP_GET_VIDEOFORMAT, Len(BITCapTureInfo), BITCapTureInfo
    
    If BITCapTureInfo.bmiHeader.biBitCount <> 0 Then
        SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "CaptureType", BITCapTureInfo.bmiHeader.biBitCount
        SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "CaptureWidth", BITCapTureInfo.bmiHeader.biWidth
        SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "CaptureHeight", BITCapTureInfo.bmiHeader.biHeight
    End If
    
    Me.MMControl.Command = "close"
    
    If Dir(strTmpFileName) <> "" Then
        Kill strTmpFileName
    End If
    
    If Not mfrmRepEdit Is Nothing Then
        Unload mfrmRepEdit
        Set mfrmRepEdit = Nothing
    End If
    
    If App.LogMode <> 0 Then
        '保存压缩设置
        blCompressionStup = False
        blClosefrm = True
        capDlgVideoCompression hCapWnd
    End If
    
    '释放采集设备及窗体
    Call SendMessageAsLong(hCapWnd, WM_CAP_DRIVER_DISCONNECT, gintDeviceIndex, 0&)
    DestroyWindow hCapWnd
    hCapWnd = 0
End Sub

Private Sub labState_Click()

End Sub

Private Sub MSComm1_OnComm()
    Dim strInput As String
    
    On Error Resume Next
    strInput = ""
    If MSComm1.InBufferCount > 0 Then strInput = MSComm1.Input
    
    If Not (MSComm1.CommEvent = comEvCTS Or MSComm1.CommEvent = comEvDSR _
        Or MSComm1.CommEvent = comEvCD Or MSComm1.CommEvent = comEvRing Or strInput <> "" _
        Or MSComm1.CommEvent = comEvSend Or MSComm1.CommEvent = comEvReceive) Then Exit Sub
'     If Not (MSComm1.CommEvent = comEvCTS Or MSComm1.CommEvent = comEvDSR _
'        Or MSComm1.CommEvent = comEvCD Or MSComm1.CommEvent = comEvRing Or strInput <> "") Then Exit Sub
    
    If intCapType = 1 Then '转换触发
        If intComState <> MSComm1.CommEvent Then
           '如果累计时间超过了采图时间间隔，则采集图像
           If lngComTime > intComInterval Then
               If Me.tbrThis.Buttons("采集").Enabled Then
                   subCaptureImage
                   SaveImage CStr(lngDeviceNO)
                   blSaveMessage = True
               End If
           End If
           
           '记录新的COM状态，计时器清零，启动timer
           intComState = MSComm1.CommEvent
           lngComTime = 0
           tmrComm.Enabled = True
        End If
    Else   '直接触发
        '两次踩下脚踏的时间间隔不能少于3秒
        If DateDiff("S", dtLastCapture, time) < intComInterval Then
            dtLastCapture = time
            Exit Sub
        End If
        dtLastCapture = time
        If Me.tbrThis.Buttons("采集").Enabled Then
            subCaptureImage
            SaveImage CStr(lngDeviceNO)
            blSaveMessage = True
        End If
    End If
End Sub

Private Sub PicX1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    blMoveDown = True
End Sub

Private Sub PicX1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngTmpX As Long
    If blMoveDown = True Then
        If Me.PicX1.Left + x <= Me.PicPar.Width + Me.PicPar.Left - 1000 And Me.PicX1.Left + x > Me.PicX2.Left - Me.PicCli.Width - Me.PicX1.Width Then
            Me.PicPar.Left = Me.PicX1.Left + Me.PicX1.Width
            Me.PicPar.Width = Me.PicPar.Width - x
            Me.PicY1.Width = Me.PicPar.Width + Me.PicX1.Width * 2 - 10
            Me.PicY2.Width = Me.PicPar.Width + Me.PicX1.Width * 2 - 10
            Me.PicX1.Left = Me.PicX1.Left + x
            Me.PicY1.Left = Me.PicX1.Left
            Me.PicY2.Left = Me.PicX1.Left
            If Me.PicCli.Left > 0 Then
                Me.PicCli.Left = 0
            Else
                Me.PicCli.Left = Me.PicCli.Left - x
            End If
            Me.PicVideo.Left = Me.PicCli.Left
            Me.Slider1.Left = Me.PicY2.Left
            Me.Slider1.Width = Me.PicY2.Width
        End If
    End If
End Sub

Private Sub PicX1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blMoveDown = False
    GetCaptureWinSize
End Sub

Private Sub PicX2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    blMoveDown = True
End Sub

Private Sub PicX2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If blMoveDown = True Then
        If Me.PicX2.Left + x >= Me.PicX1.Left + 1000 And Me.PicX2.Left + x <= Me.PicX1.Left + Me.PicCli.Width Then
            Me.PicX2.Left = Me.PicX2.Left + x
            Me.PicPar.Width = Me.PicX2.Left - (Me.PicX1.Left + Me.PicX1.Width)
            Me.PicY1.Width = Me.PicX2.Left - Me.PicX1.Left
            Me.PicY2.Width = Me.PicX2.Left - Me.PicX1.Left
            
            If Me.PicCli.Left < 0 Then
                If (Me.PicX2.Left - Me.PicX1.Left - Me.PicX1.Width) >= Me.PicCli.Width - Abs(Me.PicCli.Left) Then
                    Me.PicCli.Left = Me.PicPar.Width - Me.PicCli.Width
                End If
            End If
            Me.PicVideo.Left = Me.PicCli.Left
            Me.Slider1.Left = Me.PicY2.Left
            Me.Slider1.Width = Me.PicY2.Width
        End If
        
    End If
End Sub

Private Sub PicX2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blMoveDown = False
    GetCaptureWinSize
End Sub

Private Sub PicY_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    blMoveDown = True
End Sub

Private Sub PicY_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If blMoveDown = True Then
        If Me.PicY.Top + y < Me.ScaleHeight - 200 And Me.PicY.Top + y > Me.PicPar.Height + Me.tbrThis.Height + Me.stbThis.Height + 200 Then
            Me.PicY.Top = Me.PicY.Top + y
            Me.PicDicom.Top = Me.PicY.Top + Me.PicY.Height
            Me.PicDicom.Height = Me.ScaleHeight - Me.PicY.Top
            Me.DicomViewer.Height = IIf((Me.PicDicom.ScaleHeight - Me.stbThis.Height) < 0, Me.DicomViewer.Height, (Me.PicDicom.ScaleHeight - Me.stbThis.Height))
            lngDicomImageHeight = Me.PicDicom.Top    'Me.PicDicom.Height
        End If
    End If
End Sub

Private Sub PicY_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blMoveDown = False
    Form_Resize
End Sub

Private Sub PicY1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    blMoveDown = True
End Sub

Private Sub PicY1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If blMoveDown = True Then
        If Me.PicY1.Top + y <= Me.PicPar.Height + Me.PicPar.Top - 1000 And Me.PicY1.Top + y >= Me.PicY2.Top - Me.PicCli.Height - Me.PicY1.Height Then
            Me.PicY1.Top = Me.PicY1.Top + y
            Me.PicPar.Height = Me.PicPar.Height - y
            Me.PicPar.Top = Me.PicY1.Top + Me.PicY1.Height
            Me.PicX1.Height = Me.PicPar.Height + Me.PicY1.Height * 2 - 10
            Me.PicX2.Height = Me.PicPar.Height + Me.PicY1.Height * 2 - 10
            Me.PicX1.Top = Me.PicY1.Top
            Me.PicX2.Top = Me.PicY1.Top
            
            If Me.PicCli.Top > 0 Then
                Me.PicCli.Top = 0
            Else
                Me.PicCli.Top = Me.PicCli.Top - y
            End If
            Me.PicVideo.Top = Me.PicCli.Top
        End If
    End If
End Sub

Private Sub PicY1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blMoveDown = False
    GetCaptureWinSize
End Sub

Private Sub PicY2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    blMoveDown = True
End Sub

Private Sub PicY2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If blMoveDown = True Then
        If Me.PicY2.Top + y >= Me.PicY1.Top + 1000 And Me.PicY2.Top + y <= Me.PicY1.Top + Me.PicCli.Height Then
            Me.PicY2.Top = Me.PicY2.Top + y
            Me.PicPar.Height = Me.PicY2.Top - (Me.PicY1.Top + Me.PicY1.Height)
            Me.PicX1.Height = Me.PicY2.Top - Me.PicY1.Top + Me.PicY2.Height
            Me.PicX2.Height = Me.PicY2.Top - Me.PicY1.Top + Me.PicY2.Height
            
            If Me.PicCli.Top < 0 Then
                If (Me.PicY2.Top - Me.PicY1.Top - Me.PicY1.Height) >= Me.PicCli.Height - Abs(Me.PicCli.Top) Then
                    Me.PicCli.Top = Me.PicPar.Height - Me.PicCli.Height
                End If
            End If
            Me.PicVideo.Top = Me.PicCli.Top
        End If
    End If
End Sub

Private Sub PicY2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blMoveDown = False
    GetCaptureWinSize
End Sub

Sub GetCaptureWinSize()
    '得以采集了窗体和父窗体位置和大小
    lngCliWinTop = Me.PicCli.Top
    lngCliWinLeft = Me.PicCli.Left
    lngParWinWidth = Me.PicPar.Width
    lngParWinHeight = Me.PicPar.Height
End Sub
Sub subCaptureImage()
    Dim strTmpPath As String
    Dim ImgTmpImage As New DicomImage
    Dim iRows As Integer, iCols As Integer
    
    
    strTmpPath = App.Path & IIf(Len(App.Path) > 3, "\tmp.bmp", "tmp.bmp")
    
    On Error GoTo SaveFileError
    
    '采集图像
    If blVideoState = False Then
        mCopyImageToClipBoard
        SavePicture Clipboard.GetData(2), strTmpPath
    Else
        Me.PicTmp1.Width = Me.PicCli.Width
        Me.PicTmp1.Height = Me.PicCli.Height
        BitBlt Me.PicTmp1.hdc, 0, 0, Me.PicCli.Width, Me.PicCli.Height, Me.PicVideo.hdc, 0, 0, &HCC0020
        Me.PicTmp1.Picture = Me.PicTmp1.Image
        SavePicture Me.PicTmp1.Picture, strTmpPath
    End If
    
    Me.PicTmp1.Picture = LoadPicture(strTmpPath)
    
    If Me.tbrThis.Buttons("报告").Caption = "报告" Then
        Me.PicTmp2.Width = Me.PicPar.Width
        Me.PicTmp2.Height = Me.PicPar.Height
        
        Me.PicTmp2.PaintPicture Me.PicTmp1.Picture, 0, 0, lngParWinWidth, lngParWinHeight, _
        Abs(lngCliWinLeft), Abs(lngCliWinTop), lngParWinWidth, lngParWinHeight, vbSrcCopy
        
        Me.PicTmp2.Picture = Me.PicTmp2.Image
        
        SavePicture Me.PicTmp2.Picture, strTmpPath
    Else
        Me.PicTmp2.Width = lngParWinWidth
        Me.PicTmp2.Height = lngParWinHeight
        
        Me.PicTmp2.PaintPicture Me.PicTmp1.Picture, 0, 0, lngParWinWidth, lngParWinHeight, _
        Abs(lngCliWinLeft), Abs(lngCliWinTop), lngParWinWidth, lngParWinHeight, vbSrcCopy
        
        Me.PicTmp2.Picture = Me.PicTmp2.Image
        
        SavePicture Me.PicTmp2.Picture, strTmpPath
        'SavePicture Me.PicTmp1.Picture, strTmpPath
    End If
    ResizeRegion Me.DicomViewer.Images.count + 1, Me.DicomViewer.Width, Me.DicomViewer.Height, iRows, iCols
    
    
    
    ImgTmpImage.FileImport strTmpPath, ""
    With Me.DicomViewer
    
        Me.DicomViewer.MultiColumns = iCols
        Me.DicomViewer.MultiRows = iRows
        ImgTmpImage.PatientID = strPatientID
        
        '统一检查UID和序列UID
        If .Images.count > 0 Then
            ImgTmpImage.StudyUID = .Images(1).StudyUID
            ImgTmpImage.SeriesUID = .Images(1).SeriesUID
        ElseIf Len(strStudyUID) > 0 Then
            ImgTmpImage.StudyUID = strStudyUID
            If Len(strSeriesID) > 0 Then ImgTmpImage.SeriesUID = strSeriesID
        Else
            strStudyUID = ImgTmpImage.StudyUID
        End If
        
        Me.DicomViewer.Images.Add ImgTmpImage
        
        Kill strTmpPath
        
        If iCurImageIndex > 0 Then .Images(iCurImageIndex).BorderColour = vbWhite
        
        With .Images(.Images.count)
            .BorderStyle = 6: .BorderWidth = 1: .BorderColour = vbRed
        End With
        iCurImageIndex = .Images.count
    End With
    
    Exit Sub
    
SaveFileError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub ShowMe(ByVal strPrivs As String, ByVal lngAdviceID As Long, ByVal lngSendNO As Long, frmObject As Object, strNO As String, int记录性质 As Integer, _
                        lng单据ID, lng病历ID, str医嘱内容 As String, Optional ByVal strType As String = "", _
                        Optional ByVal strCheckUID As String = "")
    Dim strSQL As String, rsTmp As ADODB.Recordset
     
    mstrPrivs = strPrivs
    strPatientID = lngAdviceID: strStudyUID = "": strSeriesID = ""
    mlngAdviceID = lngAdviceID: mlngSendNO = lngSendNO
    strImgType = strType: strStudyUID = strCheckUID
    mstrNO = strNO: mint记录性质 = int记录性质: mlng单据ID = lng单据ID: mlng病历ID = lng病历ID: mstr医嘱内容 = str医嘱内容
    Set mfrmPacsWork = frmObject
    
    strSQL = "Select 姓名,性别,年龄,检查号,医嘱ID As 检查标识 From 影像检查记录 Where 医嘱ID=[1] And 发送号=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, mlngAdviceID, mlngSendNO)
    If rsTmp.EOF Then Exit Sub
    Me.Caption = "图像采集 - " & Nvl(rsTmp("姓名")) & " " & Nvl(rsTmp("性别")) & " " & Nvl(rsTmp("年龄")) & _
        " 检查号：" & Nvl(rsTmp("检查号")) & " 检查标识：" & Nvl(rsTmp("检查标识"))
    '初始窗体显示
    iCurImageIndex = 0
    blSaveMessage = False
    
    If mstrFormMode = "恢复" Then
        Me.tbrThis.Buttons("报告").Caption = mstrFormMode
        
        Call OpenReportFrm("报告")
    End If
    Form_Resize
    
    GetAllImages DicomViewer, strStudyUID, strSeriesID, strCachePath, iCurImageIndex
    dtLastCapture = time
    
    '将程序中的UID设置成1
    dcmglbUID.RegString("UIDRoot") = "1"
    '默认时实时显示画面
    Me.PicCli.Visible = True
    Me.PicVideo.Visible = False
    Me.Slider1.Visible = False
    Me.Slider1.Value = 1
    blVideoState = False
    Me.stbThis.Panels(2).Text = "状态:实时显示中"
    
    Me.Show , frmObject
End Sub
Private Sub InitPara()
    Dim rsTmp As New ADODB.Recordset
    Dim aDevices() As Variant
    
    '打开串口
    On Error Resume Next
    MSComm1.CommPort = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "脚踏端口", "1"))
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    MSComm1.Settings = "9600,N,8,1"
    MSComm1.InputMode = comInputModeText
    MSComm1.RThreshold = 1
    MSComm1.InBufferCount = 0
    MSComm1.InputLen = 0
    MSComm1.RTSEnable = True
    MSComm1.PortOpen = True
    
    lngDeviceNO = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "设备号", "0")
    On Error GoTo DBError
    
    gstrSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=1"
    OpenRecordset rsTmp, Me.Caption
    If rsTmp.EOF Then
        MsgBox "未定义影像存储设备，请到影像设备目录中设置！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    aDevices = rsTmp.GetRows
    
    lngDeviceNO = GetDefaultDev(aDevices, lngDeviceNO)
    
    intCapType = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "脚踏擦集方式", 1))
    If intCapType < 0 Or intCapType > 1 Then
        intCapType = 1
    End If
    
    intComInterval = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\影像采集", "脚踏时间间隔", 1)
    
    Exit Sub
DBError:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function GetDefaultDev(aSource() As Variant, ByVal lngDev As String) As String
    Dim i As Long
    
    For i = 0 To UBound(aSource, 2)
        If aSource(0, i) = lngDev Then Exit For
    Next
    If i > UBound(aSource, 2) Then i = 0
    GetDefaultDev = aSource(0, i)
End Function

Sub subDelImage()
    Dim iCols As Integer, iRows As Integer
    If iCurImageIndex < 1 Then Exit Sub
    '从数据库中删除图像
    DeleteImage iCurImageIndex, CStr(lngDeviceNO)
    With DicomViewer
        .Images.Remove iCurImageIndex
        ResizeRegion .Images.count, .Width, .Height, iRows, iCols
        .MultiColumns = iCols: .MultiRows = iRows
        
        If iCurImageIndex > .Images.count Then iCurImageIndex = .Images.count
        If iCurImageIndex > 0 Then .Images(iCurImageIndex).BorderColour = vbRed
    End With
End Sub

Private Sub Slider1_Click()
    Dim v As Double

    v = Slider1.Value - Slider1.Min

    v = v / Slider1.Max * (Me.MMControl.length / PlayFPS)
    
    Me.MMControl.To = v * PlayFPS
    Me.MMControl.Command = "seek"
    Me.MMControl.Command = "stop"
    Me.stbThis.Panels(2).Text = "状态:暂停(" & strLalcTime(Me.MMControl.Position) & "/" & strLalcTime(Me.MMControl.length) & ")"
    Me.Timer1.Enabled = False
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.PicPar.Refresh
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Integer
    Dim blShow As Boolean
    Select Case Button.Key
        Case "显示"
            Me.PicCli.Visible = True
            Me.PicVideo.Visible = False
            Me.Slider1.Visible = False
            Me.Slider1.Value = 1
            blVideoState = False
            Me.stbThis.Panels(2).Text = "状态:实时显示中"
        Case "报告"
            OpenReportFrm Button.Caption
            Form_Resize
            mstrFormMode = Me.tbrThis.Buttons("报告").Caption
        Case "采集"
            subCaptureImage
            SaveImage CStr(lngDeviceNO)
            blSaveMessage = True
        Case "导入"
            InputImageFile
            blSaveMessage = True
        Case "保存"
            SaveImages DicomViewer.Images, CStr(lngDeviceNO), strCachePath, , strImgType
            blSaveMessage = False
        Case "删除"
            subDelImage
            '重新保存，排列序号
            'SaveImages DicomViewer.Images, CStr(lngDeviceNO), strCachePath, , strImgType
            blSaveMessage = False
        Case "录像"
            Me.stbThis.Panels(2).Text = "状态:采集中(按下鼠标左键或右键结束采集)"
            subSaveVideo
            Me.stbThis.Panels(2).Text = "状态:采集完成"
        Case "播放"
            subPlayVideo
            Me.stbThis.Panels(2).Text = "状态:播放中"
        Case "暂停"
            If blVideoState = True Then
                MMControl.Command = "Pause"
                Me.Timer1.Enabled = False
            End If
            Me.stbThis.Panels(2).Text = "状态:暂停(" & strLalcTime(Me.MMControl.Position) & "/" & strLalcTime(Me.MMControl.length) & ")"
        Case "快进"
            lngSpeedPaly = lngSpeedPaly + PlayFPS
        Case "端口"
            If frmParaSet.ShowMe(Me) Then InitPara
        Case "退出"
            If Dir(strTmpFileName) <> "" Then
                Kill strTmpFileName
            End If
            
            If Not mfrmRepEdit Is Nothing Then
                Unload mfrmRepEdit
                Set mfrmRepEdit = Nothing
            End If
    
            Me.Hide
    End Select
End Sub
Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim CaptureWinSize As CAPSTATUS
    Select Case ButtonMenu.Key
        Case "格式"
            mViewerFormat
            CaptureWinSize = mGetCaptureWindowStatus
            Me.PicCli.Width = CaptureWinSize.uiImageWidth * Screen.TwipsPerPixelX
            Me.PicCli.Height = CaptureWinSize.uiImageHeight * Screen.TwipsPerPixelY
            Me.PicPar.Width = Me.PicCli.Width
            Me.PicPar.Height = Me.PicCli.Height
            lngCliWinTop = 0
            lngCliWinLeft = 0
            lngParWinWidth = Me.PicPar.Width
            lngParWinHeight = Me.PicPar.Height
            Call Form_Resize
            Me.Refresh
        Case "来源"
            mViewerSource
        Case "压缩方式"
            blCompressionStup = True
            capDlgVideoCompression hCapWnd
    End Select
End Sub
Private Sub subSaveVideo()
    '采集图像
    Dim CapParams As CAPTUREPARMS
    Dim i As Integer
        
    blVideoState = False
    
    
    Me.MMControl.Command = "close"
    If Dir(strTmpFileName) <> "" Then
        Kill strTmpFileName
    End If
    
    capFileSetCaptureFile hCapWnd, strTmpFileName
    
    Me.PicCli.Visible = True
    Me.PicVideo.Visible = False
    Me.Slider1.Visible = False
    Me.Slider1.Value = 1
    
    With CapParams
        .wPercentDropForError = 10
        .fMakeUserHitOKToCapture = True
        .fUsingDOSMemory = True
        .wNumVideoRequested = 32
        .fAbortLeftMouse = -(True)
        .fAbortRightMouse = -(True)
        .wChunkGranularity = 0
        .dwAudioBufferSize = 0
        .fDisableWriteCache = False
        .fMCIControl = False
        .fStepCaptureAt2x = False
        .fYield = False
        .wNumAudioRequested = 4 '10 is max limit
        .AVStreamMaster = AVSTREAMMASTER_NONE
        .dwIndexSize = Val(GetSetting(App.Title, "preferences", "maxframes", INDEX_3_HOURS))
        .dwRequestMicroSecPerFrame = microsSecFromFPS(Val(PlayFPS))
        .fCaptureAudio = False
        .fLimitEnabled = False
        .wTimeLimit = Val(INDEX_3_HOURS)
'        .vKeyAbort = 13
    End With
   
    mcapCaptureSetSetup hCapWnd, CapParams
    capCaptureSequence hCapWnd
End Sub

Private Function microsSecFromFPS(ByVal fps As Long) As Long
    If fps = 0 Then Exit Function
    microsSecFromFPS = 1000000 / fps
End Function
Private Sub subPlayVideo()
'    Dim pMP As IMediaPosition
    Dim v As Double
    '没有时退出
    If Dir(strTmpFileName) = "" Then Exit Sub
    lngSpeedPaly = 0
    Me.PicCli.Visible = False
    Me.PicVideo.Visible = True
    Me.Slider1.Visible = True
    With Me.MMControl
        If blVideoState = False Then
            .DeviceType = "avivideo"
            .FileName = strTmpFileName
            .hWndDisplay = PicVideo.hwnd
            .Command = "Open"
            If .length <> 0 Then
                Me.Slider1.Max = .length / PlayFPS
            Else
                Me.Slider1.Max = 100
            End If
            If Me.Slider1.Max < 10 Then
                Me.Slider1.LargeChange = 1
            End If
            If Me.Slider1.Max < 100 And Me.Slider1.Max >= 10 Then
                Me.Slider1.LargeChange = 5
            End If
            If Me.Slider1.Max >= 100 Then
                Me.Slider1.LargeChange = Me.Slider1.Max / 10
            End If
            
        End If
        If Me.Slider1.Value = Me.Slider1.Max Then
            Me.Slider1.Value = 1
        End If
        v = Slider1.Value - Slider1.Min
        v = v / Slider1.Max * (.length / PlayFPS)
        .To = v * PlayFPS
        .Command = "seek"
        .Command = "play"
    End With
    Timer1.Enabled = True
    blVideoState = True
End Sub

Private Sub Timer1_Timer()
    
    Dim v As Double
    If blVideoState = False Then Exit Sub
    If Me.MMControl.Position = 0 Then Exit Sub
    If lngSpeedPaly <> 0 Then
        Me.MMControl.To = Me.MMControl.Position + lngSpeedPaly
        Me.MMControl.Command = "Seek"
        Me.MMControl.Command = "Play"
    End If
    v = (Me.MMControl.Position / PlayFPS) / (Me.MMControl.length / PlayFPS)
    Me.Slider1.Value = v * (Me.MMControl.length / PlayFPS)
    If Me.Slider1.Value = Me.Slider1.Max Then
        Me.Timer1.Enabled = False
        Me.stbThis.Panels(2).Text = "状态:播放完成"
    End If
    Me.stbThis.Panels(2).Text = "状态:播放中(" & strLalcTime(Me.MMControl.Position) & "/" & strLalcTime(Me.MMControl.length) & ")" & IIf(lngSpeedPaly <> 0, "播放速度:" & lngSpeedPaly / PlayFPS & "倍", "")
End Sub
Private Sub InputImageFile()
    Dim DlgInfo As DlgFileInfo
    Dim i As Integer
    Dim ImgTmpImage As New DicomImage
    Dim ImgTmpImages As New DicomImages
    Dim iRows As Integer, iCols As Integer
    Dim blDicomFile As Boolean              '是否DICO文件 =True为DICOM文件
    Dim j As Integer
    
    '选择文件
    With Me.Comm
        
        .CancelError = False
        .MaxFileSize = 32767 '被打开的文件名尺寸设置为最大，即32K
        .Flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
        .DialogTitle = "选择文件"
        .Filter = "DICOM文件（*.dcm）(*.img)|*.dcm;*.img|图像文件 (*.BMP)(*.JPG)|*.BMP;*.JPG|所有文件（*.*）|*.*"
        .ShowOpen
        If .FileName <> "" Then
            DlgInfo = GetDlgSelectFileInfo(.FileName)
        End If
        .FileName = ""      '在打开了*.pif文件后须将Filename属性置空，
                            '否则当选取多个*.pif文件后，当前路径会改变
    End With
    
    On Error Resume Next
    
    
    For i = 1 To DlgInfo.iCount
        '排版
        ResizeRegion Me.DicomViewer.Images.count + 1, Me.DicomViewer.Width, Me.DicomViewer.Height, iRows, iCols
        err.Clear
        Set ImgTmpImage = Nothing
        ImgTmpImages.Clear
        ImgTmpImage.FileImport DlgInfo.sPath & DlgInfo.sFile(i), ""
        If err <> 0 Then
            err.Clear
            ImgTmpImages.ReadFile DlgInfo.sPath & DlgInfo.sFile(i)
            If err = 0 Then
                blDicomFile = True
            End If
        End If
        With Me.DicomViewer
        
            Me.DicomViewer.MultiColumns = iCols
            Me.DicomViewer.MultiRows = iRows
            ImgTmpImage.PatientID = strPatientID
            
            '统一检查UID和序列UID
            If .Images.count > 0 Then
                ImgTmpImage.StudyUID = .Images(1).StudyUID
                ImgTmpImage.SeriesUID = .Images(1).SeriesUID
            ElseIf Len(strStudyUID) > 0 Then
                ImgTmpImage.StudyUID = strStudyUID
                If Len(strSeriesID) > 0 Then ImgTmpImage.SeriesUID = strSeriesID
            Else
                strStudyUID = ImgTmpImage.StudyUID
            End If
            
            'DICOM文件和非DICOM文件的处理
            If blDicomFile = False Then
                Me.DicomViewer.Images.Add ImgTmpImage
            Else
                For j = 1 To ImgTmpImages.count
                    ImgTmpImages(j).PatientID = strPatientID
                    '统一检查UID和序列UID
                    If .Images.count > 0 Then
                        ImgTmpImages(j).StudyUID = .Images(1).StudyUID
                        ImgTmpImages(j).SeriesUID = .Images(1).SeriesUID
                    ElseIf Len(strStudyUID) > 0 Then
                        ImgTmpImages(j).StudyUID = strStudyUID
                        If Len(strSeriesID) > 0 Then ImgTmpImages(j).SeriesUID = strSeriesID
                    Else
                        strStudyUID = ImgTmpImage.StudyUID
                    End If
                    Me.DicomViewer.Images.Add ImgTmpImages(j)
                    If iCurImageIndex > 0 Then .Images(iCurImageIndex).BorderColour = vbWhite
                    With .Images(.Images.count)
                        .BorderStyle = 6: .BorderWidth = 1: .BorderColour = vbRed
                    End With
                Next
            End If
            
            If iCurImageIndex > 0 Then .Images(iCurImageIndex).BorderColour = vbWhite
            
            With .Images(.Images.count)
                .BorderStyle = 6: .BorderWidth = 1: .BorderColour = vbRed
            End With
            iCurImageIndex = .Images.count
        End With
    Next
    
End Sub
Private Function GetDlgSelectFileInfo(strFileName As String) As DlgFileInfo
'------------------------------------------------
'功能：将文件名转化为全路径数组
'参数：strFileName--文件名，通过打开文件控件来获得。
'返回：全路径数组
'上级函数或过程：mdlFile.funGetFileList
'下级函数或过程：无
'引用的外部参数：无
'编制人：曾超
'------------------------------------------------
    Dim sPath, tmpStr As String
    Dim sFile() As String
    Dim iCount, i As Integer
    On Error GoTo ErrHandle
    sPath = CurDir()  '获得当前的路径，因为在CommonDialog中改变路径时会改变当前的Path
    tmpStr = Right$(strFileName, Len(strFileName) - Len(sPath)) '将文件名分离出来
    
    If Left$(tmpStr, 1) = Chr$(0) Then
        '选择了多个文件(表现为第一个字符为空格)
        For i = 1 To Len(tmpStr)
            If Mid$(tmpStr, i, 1) = Chr$(0) Then
                iCount = iCount + 1
                ReDim Preserve sFile(iCount)
            Else
                sFile(iCount) = sFile(iCount) & Mid$(tmpStr, i, 1)
            End If
        Next i
    Else
        '只选择了一个文件(注意：根目录下的文件名除去路径后没有"\"）
        iCount = 1
        ReDim Preserve sFile(iCount)
        If Left$(tmpStr, 1) = "\" Then tmpStr = Right$(tmpStr, Len(tmpStr) - 1)
        sFile(iCount) = tmpStr
    End If
    
    GetDlgSelectFileInfo.iCount = iCount
    ReDim GetDlgSelectFileInfo.sFile(iCount)
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    GetDlgSelectFileInfo.sPath = sPath
    For i = 1 To iCount
        GetDlgSelectFileInfo.sFile(i) = sFile(i)
    Next i
    Exit Function
ErrHandle:
    MsgBox "GetDlgSelectFileInfo函数执行错误！", vbOKOnly + vbCritical, "自定义函数错误"
End Function
Private Sub OpenReportFrm(strCaption As String)
    Dim i As Integer
    Dim ObjFrm As Object
    Dim blnAudit As Boolean     '可以审核
    Dim blnRollBack As Boolean  '可以驳回
    
    With Me.tbrThis
        If strCaption = "报告" Then
            Me.tbrThis.Buttons("报告").Caption = "恢复"
'            Me.PicPar.Visible = False
            Me.PicPar.Height = 1
            Me.PicPar.Width = 1
            Me.PicX1.Visible = False
            Me.PicX2.Visible = False
            Me.PicY1.Visible = False
            Me.PicY2.Visible = False
            Me.stbThis.Visible = False
            Me.PicY.Visible = False
            Me.WindowState = 0
            If lngReportWidth = 0 Then
                Me.WindowState = 0
                Me.Width = 2000
                Me.Height = Screen.Height - 500
                Me.Top = 0
                Me.Left = Screen.Width - Me.Width
            Else
                Me.WindowState = 0
                Me.Width = lngReportWidth
                Me.Height = lngReportHeight
                Me.Top = lngReportTop
                Me.Left = lngReportLeft
            End If
            If Not mfrmRepEdit Is Nothing Then
                Call ShowWindow(mfrmRepEdit.hwnd, SW_RESTORE)
                Call BringWindowToTop(mfrmRepEdit.hwnd)
            Else
                '打开报告窗体
                blnAudit = (InStr(mstrPrivs, "报告审核") <> 0)
                blnRollBack = (InStr(mstrPrivs, "报告驳回") <> 0)
                EditReport mfrmPacsWork, mstrNO, mint记录性质, mlng单据ID, mlng病历ID, mstr医嘱内容, False, False, ObjFrm, _
                0, False, True, strPatientID, False, IIf(blnAudit, "1", "0") & IIf(blnRollBack, "1", "0") & "1"
                Set mfrmRepEdit = ObjFrm
                If mfrmRepEdit.hwnd <> 0 Then
                    If lngReportWinWidth = 0 Then
                        mfrmRepEdit.WindowState = 0
                        mfrmRepEdit.Width = Me.Left
                        mfrmRepEdit.Height = Screen.Height - 500
                        mfrmRepEdit.Top = 0
                        mfrmRepEdit.Left = 0
                    Else
                        mfrmRepEdit.WindowState = 0
                        mfrmRepEdit.Width = lngReportWinWidth
                        mfrmRepEdit.Height = lngReportWinHeight
                        mfrmRepEdit.Top = lngReportWinTop
                        mfrmRepEdit.Left = lngReportWinLeft
                    End If
                End If
            End If
            
        Else
            Me.tbrThis.Buttons("报告").Caption = "报告"
            Me.PicPar.Visible = True
            Me.PicX1.Visible = True
            Me.PicX2.Visible = True
            Me.PicY1.Visible = True
            Me.PicY2.Visible = True
            Me.stbThis.Visible = True
            Me.PicY.Visible = True
            lngReportWidth = Me.Width
            lngReportHeight = Me.Height
            lngReportTop = Me.Top
            lngReportLeft = Me.Left
            Me.WindowState = 2
            
        End If
        For i = 1 To .Buttons.count
            If .Buttons(i).Key = "采集" Or .Buttons(i).Key = "保存" Or .Buttons(i).Key = "报告" Then
                
            Else
                .Buttons(i).Visible = IIf(Me.tbrThis.Buttons("报告").Caption = "恢复", False, True)
            End If
        Next
    End With
End Sub
Private Sub mfrmRepEdit_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ReprotWindow_frmWidth", mfrmRepEdit.Width
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ReprotWindow_frmHeight", mfrmRepEdit.Height
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ReprotWindow_frmTop", mfrmRepEdit.Top
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.ProductName & "\frmCapture", "ReprotWindow_frmLeft", mfrmRepEdit.Left
    Set mfrmRepEdit = Nothing
    
    If Dir(strTmpFileName) <> "" Then
        Kill strTmpFileName
    End If
    
    If Not mfrmRepEdit Is Nothing Then
        Unload mfrmRepEdit
        Set mfrmRepEdit = Nothing
    End If
    Me.Hide
End Sub

Private Sub tmrComm_Timer()
    lngComTime = lngComTime + 1
    '大于0.8秒，则自动放弃
    If lngComTime > 80 Then
        lngComTime = 0
        tmrComm.Enabled = False
    End If
End Sub
