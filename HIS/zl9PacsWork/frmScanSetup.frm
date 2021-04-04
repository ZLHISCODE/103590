VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{B1790453-7708-48C1-B5CC-75255FA4B066}#1.0#0"; "ZLDSVideoProcess.ocx"
Begin VB.Form frmScanSetup 
   Caption         =   "扫描设置"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6255
   Icon            =   "frmScanSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6255
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取 消(&Q)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确 定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "扫描后继续扫描(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   10
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "扫描(&S)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   9
      Top             =   4920
      Width           =   1100
   End
   Begin VB.PictureBox picImgScan 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   5160
      ScaleHeight     =   4425
      ScaleWidth      =   5760
      TabIndex        =   7
      Top             =   4200
      Width           =   5785
      Begin ZLDSVideoProcess.DSCapture wdmCapture 
         Height          =   4455
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   5775
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
         CurWidth        =   385
         CurHeight       =   297
         CurVideoWidth   =   383
         CurVideoHeight  =   277
         ShowModel       =   0
         CapParameterWindPos=   8
         SnatchWay       =   0
         ParameterCfgFileName=   ""
         HideCfgItem     =   0
         AppHandle       =   0
      End
   End
   Begin TabDlg.SSTab stbConfig 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "基本设置"
      TabPicture(0)   =   "frmScanSetup.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "扫描参数设置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   5535
         Begin VB.OptionButton optImageType 
            Caption         =   "JPG格式"
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   19
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton optImageType 
            Caption         =   "BMP格式"
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   18
            Top             =   960
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton cmdDirSelect 
            Caption         =   "…"
            Height          =   375
            Left            =   4920
            TabIndex        =   15
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtTempDir 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1440
            TabIndex        =   14
            Text            =   "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA"
            Top             =   360
            Width           =   3510
         End
         Begin VB.Label Label1 
            Caption         =   "扫描图像格式"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label labTempDir 
            Caption         =   "扫描临时目录"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   16
            Top             =   465
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "扫描驱动类型"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   5535
         Begin VB.CommandButton cmdVideoSetup 
            Caption         =   "视频设置(&V)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   6
            Top             =   480
            Width           =   1410
         End
         Begin VB.CommandButton cmdChoiceEqu 
            Caption         =   "选择设备(&S)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3720
            TabIndex        =   5
            Top             =   1200
            Width           =   1410
         End
         Begin VB.CommandButton cmdParameterCfg 
            Caption         =   "扫描设置(&S)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   4
            Top             =   1200
            Width           =   1410
         End
         Begin VB.OptionButton optDriver 
            Caption         =   "WDM 驱动"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   540
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton optDriver 
            Caption         =   "TWAIN 驱动"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   2
            Top             =   1280
            Width           =   1350
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgOpenDir 
      Left            =   360
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmScanSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'视频驱动类型
Private Enum TVideoDriverType
  vdtWDM = 0
  vdtVFW = 1
  vdtTWAIN = 2
  '其他需要支持的驱动类型......
End Enum

Private Const CAPTURE_PARAMETER_CONFIG_FILE_NAME As String = "ZLVideoProcess.ini"
Private mScanDriverType As TVideoDriverType '保存当前要使用的驱动类型
Private mobjImgScan As ImgScan
Private mobjParent As Object
Private mstrScanFile As String

Private Sub cmdDirSelect_Click()
  Dim shl As Object
  Set shl = CreateObject("Shell.application")
  
  On Error GoTo final
  
    Dim fd As Object
    Set fd = shl.BrowseForFolder(0, "扫描设备临时目录选择", 0, "\")
  
    If Not fd Is Nothing Then
      txtTempDir.Text = fd.Self.Path
    End If
final:
  Set shl = Nothing
  Set fd = Nothing
End Sub

Private Sub CmdOK_Click()
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\frmPetitionCapture", "扫描驱动类型", mScanDriverType
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\frmPetitionCapture", "扫描临时目录", txtTempDir.Text
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\frmPetitionCapture", "扫描图像格式", IIf(optImageType(0).value, 0, 1)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChoiceEqu_Click()
    If Not mobjImgScan Is Nothing Then
        mobjImgScan.ShowSelectScanner
    End If
End Sub

Private Sub cmdParameterCfg_Click()
    If Not mobjImgScan Is Nothing Then
        Call mobjImgScan.ShowScanPreferences
    End If
End Sub

Private Sub cmdScan_Click(Index As Integer)
    wdmCapture.CaptureBmpImageToFile mstrScanFile
        
    Call mobjParent.subCaptureImg(True, mstrScanFile)
        
    If Index = 1 Then Unload Me '每次只扫描一张图像
End Sub

Private Sub cmdVideoSetup_Click()
    wdmCapture.HideCfgItem = hciVideoShowWay + hciVideoState + hciImageCapture + hciVideoEncoder
    Call wdmCapture.ShowCaptureParameterCfgDialog(0)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = vbKeyEscape Then Unload Me
End Sub

Public Sub ShowParameterConfig(ByRef objImgScan As ImgScan, ByRef objParent As Object)
    Dim intImageType As Integer
    
    Set mobjImgScan = objImgScan
    Set mobjParent = objParent
    
    picImgScan.Visible = False
    cmdScan(0).Visible = False
    cmdScan(1).Visible = False
    
    mScanDriverType = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPetitionCapture", "扫描驱动类型", 0))
    txtTempDir.Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPetitionCapture", "扫描临时目录", "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA")
    
    intImageType = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPetitionCapture", "扫描图像格式", 0))
    
    If intImageType = 0 Then
        optImageType(0).value = True
    Else
        optImageType(1).value = True
    End If
    
    Call UpdateCmdState(mScanDriverType = vdtTWAIN)
    Call LoadDriverType
    Call InitwdmCapture
    
    Call Me.Show(1, objParent)
End Sub

Public Sub ShowScanWind(ByVal strScanFile As String, ByRef objParent As Object)
    mstrScanFile = strScanFile
    Set mobjParent = objParent
    
    stbConfig.Visible = False
    cmdOK.Visible = False
    
    Me.Caption = "扫描图像"
    Me.cmdOK.Caption = "扫描(&S)"
    
    Call InitwdmCapture
    
    Call Me.Show(1, objParent)
End Sub

Private Sub InitwdmCapture()
    Dim strCfgName As String
    
    strCfgName = App.Path & "\" & CAPTURE_PARAMETER_CONFIG_FILE_NAME
    
    wdmCapture.ParameterCfgFileName = strCfgName
    wdmCapture.ReadParameterFromFile
    
    '设置视频的显示模式
    wdmCapture.ShowModel = smStretch

    '在读取文件配置后修改该属性（只有设置该属性，才能根据四条边框进行调节和显示）
    wdmCapture.AppHandle = Me.hWnd
    wdmCapture.IsShowState = False
    wdmCapture.StartPreview
    wdmCapture.RefreshWindow
End Sub

Private Sub LoadDriverType()
    Select Case mScanDriverType
        Case vdtWDM
            optDriver(0).value = True
            
        Case vdtTWAIN
            optDriver(2).value = True
    End Select
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Caption = "扫描图像" Then
        picImgScan.Left = 120
        picImgScan.Top = 120
        picImgScan.Width = Me.ScaleWidth - 240
        picImgScan.Height = Me.ScaleHeight - cmdCancel.Height - 480
        
        cmdScan(0).Top = picImgScan.Top + picImgScan.Height + 200
        cmdScan(0).Left = Me.ScaleWidth - cmdScan(0).Width - cmdScan(1).Width - cmdCancel.Width - 640
        cmdScan(1).Left = cmdScan(0).Left + cmdScan(0).Width + 240
        cmdScan(1).Top = cmdScan(0).Top
        cmdCancel.Left = cmdScan(1).Left + cmdScan(1).Width + 240
        cmdCancel.Top = cmdScan(1).Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjImgScan = Nothing
    wdmCapture.FreeRes
End Sub

Private Sub optDriver_Click(Index As Integer)
On Error GoTo ErrHandle
    Select Case Index
        Case 0
            mScanDriverType = vdtWDM
            
            Call UpdateCmdState(False)
        Case 2
            mScanDriverType = vdtTWAIN
            
            Call UpdateCmdState(True)
    End Select
    
    Exit Sub
ErrHandle:
    MsgBox err.Description, vbOKOnly, gstrSysName
End Sub

Private Sub UpdateCmdState(ByVal blnEnabled As Boolean)
    cmdParameterCfg.Enabled = blnEnabled
    cmdChoiceEqu.Enabled = blnEnabled
    cmdVideoSetup.Enabled = Not blnEnabled
End Sub

Private Sub picImgScan_Resize()
    On Error Resume Next
    
    wdmCapture.Left = 0
    wdmCapture.Top = 0
    wdmCapture.Width = picImgScan.ScaleWidth
    wdmCapture.Height = picImgScan.ScaleHeight
End Sub
