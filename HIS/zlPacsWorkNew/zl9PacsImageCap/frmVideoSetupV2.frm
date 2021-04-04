VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVideoSetupV2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "采集参数设置"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   Icon            =   "frmVideoSetupV2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin ScanLibCtl.ImgScan imageScannerConfig 
      Left            =   3015
      Top             =   4005
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "关 闭(&C)"
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
      Left            =   4770
      TabIndex        =   29
      Top             =   3945
      Width           =   1100
   End
   Begin TabDlg.SSTab stbConfig 
      Height          =   3720
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   6562
      _Version        =   393216
      Style           =   1
      Tab             =   1
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
      TabPicture(0)   =   "frmVideoSetupV2.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboBakDevice"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkAllowChangeSize"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboSaveDevice"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkUseGameModeCap"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkCaptureSound"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "采集设置"
      TabPicture(1)   =   "frmVideoSetupV2.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "optDriver(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdSelectDriver"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtDriverPath"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdParameterCfg"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "optDriver(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "optDriver(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "optDriver(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "脚踏设置"
      TabPicture(2)   =   "frmVideoSetupV2.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblItem(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "labComInterval"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblItem(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "labCaptureWay"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lblItem(3)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblItem(2)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cbxHotKey"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtComInterval"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cboPort"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cboCommCapType"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cboAfterTagHotKey"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cboAfterHotKey"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      Begin VB.CheckBox chkCaptureSound 
         Caption         =   "采集声音提示"
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
         Left            =   -74880
         TabIndex        =   35
         Top             =   2520
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox chkUseGameModeCap 
         Caption         =   "使用游戏手柄采集模式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   34
         Top             =   2040
         Width           =   2400
      End
      Begin VB.ComboBox cboAfterHotKey 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmVideoSetupV2.frx":0060
         Left            =   -73485
         List            =   "frmVideoSetupV2.frx":008B
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2280
         Width           =   4110
      End
      Begin VB.ComboBox cboAfterTagHotKey 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmVideoSetupV2.frx":00C4
         Left            =   -73485
         List            =   "frmVideoSetupV2.frx":00EF
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2730
         Width           =   4110
      End
      Begin VB.Frame Frame1 
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
         Height          =   1240
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   5415
         Begin VB.CommandButton cmdImageCompressConfig 
            Caption         =   "压缩设置(&P)"
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
            Left            =   4005
            TabIndex        =   19
            Top             =   720
            Width           =   1305
         End
         Begin VB.CommandButton cmdSelectScanDevice 
            Caption         =   "设备选择(&D)"
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
            Left            =   2445
            TabIndex        =   18
            Top             =   720
            Width           =   1305
         End
         Begin VB.TextBox tbxTempDir 
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
            Left            =   1800
            TabIndex        =   17
            Text            =   "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA"
            Top             =   255
            Width           =   3150
         End
         Begin VB.CommandButton cmdDirSelect 
            Caption         =   "…"
            Height          =   375
            Left            =   4920
            TabIndex        =   16
            Top             =   255
            Width           =   375
         End
         Begin VB.Label labTempDir 
            Caption         =   "扫描设备临时目录"
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
            TabIndex        =   20
            Top             =   330
            Width           =   1695
         End
      End
      Begin VB.ComboBox cboCommCapType 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmVideoSetupV2.frx":0128
         Left            =   -73485
         List            =   "frmVideoSetupV2.frx":0135
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   930
         Width           =   4110
      End
      Begin VB.ComboBox cboPort 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmVideoSetupV2.frx":0157
         Left            =   -73485
         List            =   "frmVideoSetupV2.frx":0176
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   465
         Width           =   4110
      End
      Begin VB.TextBox txtComInterval 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73485
         TabIndex        =   12
         Text            =   "1"
         Top             =   1395
         Width           =   3810
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
         Left            =   240
         TabIndex        =   11
         Top             =   810
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optDriver 
         Caption         =   "VFW 驱动"
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
         Index           =   1
         Left            =   2280
         TabIndex        =   10
         Top             =   810
         Width           =   1200
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
         Left            =   4320
         TabIndex        =   9
         Top             =   840
         Width           =   1350
      End
      Begin VB.CommandButton cmdParameterCfg 
         Caption         =   "视频设置"
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
         Left            =   4560
         TabIndex        =   8
         Top             =   1150
         Width           =   1050
      End
      Begin VB.ComboBox cboSaveDevice 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73380
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   4020
      End
      Begin VB.CheckBox chkAllowChangeSize 
         Caption         =   "允许改变采集区域大小"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   6
         Top             =   1560
         Width           =   2400
      End
      Begin VB.ComboBox cboBakDevice 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73380
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   945
         Width           =   4020
      End
      Begin VB.ComboBox cbxHotKey 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmVideoSetupV2.frx":01AE
         Left            =   -73485
         List            =   "frmVideoSetupV2.frx":01D9
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1845
         Width           =   4110
      End
      Begin VB.TextBox txtDriverPath 
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
         TabIndex        =   3
         Top             =   1160
         Width           =   2655
      End
      Begin VB.CommandButton cmdSelectDriver 
         Caption         =   "…"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   1150
         Width           =   375
      End
      Begin VB.OptionButton optDriver 
         Caption         =   "专用视频"
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
         Index           =   3
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "后台采集热键"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   -74865
         TabIndex        =   33
         Top             =   2325
         Width           =   1260
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "标记更新热键"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   -74865
         TabIndex        =   32
         Top             =   2805
         Width           =   1260
      End
      Begin VB.Label labCaptureWay 
         Caption         =   "脚踏采集方式"
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
         Left            =   -74865
         TabIndex        =   27
         Top             =   975
         Width           =   1305
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "脚踏端口"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   -74430
         TabIndex        =   26
         Top             =   480
         Width           =   840
      End
      Begin VB.Label labComInterval 
         Caption         =   "脚踏时间间隔                                      秒"
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
         Left            =   -74865
         TabIndex        =   25
         Top             =   1410
         Width           =   5460
      End
      Begin VB.Label Label2 
         Caption         =   "视频驱动类型设置：                  "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   195
         TabIndex        =   24
         Top             =   510
         Width           =   1920
      End
      Begin VB.Label Label4 
         Caption         =   "采集存储设备"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74805
         TabIndex        =   23
         Top             =   540
         Width           =   1380
      End
      Begin VB.Label Label7 
         Caption         =   "备份存储设备"
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
         Left            =   -74805
         TabIndex        =   22
         Top             =   990
         Width           =   1305
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "采集热键"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   -74430
         TabIndex        =   21
         Top             =   1920
         Width           =   840
      End
   End
   Begin MSComDlg.CommonDialog dlgOpenDir 
      Left            =   2445
      Top             =   3945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   3570
      TabIndex        =   28
      Top             =   3945
      Width           =   1100
   End
End
Attribute VB_Name = "frmVideoSetupV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsOK As Boolean

Private DX7 As New DirectX7
Private DxInput As DirectInput
Private DiDevEnum As DirectInputEnumDevices

Private mVideoCapture As clsVideoCapture

Public Event OnVideoDirverChange(ByVal vdtDirverType As TVideoDriverType)


'modify by tjh at 2010-01-21
Public Function ShowParameterConfig(ByRef videoCapture As clsVideoCapture, ByRef owner As Object) As Boolean
BUGEX "ShowParameterConfig 1"
    ShowParameterConfig = False
    Set mVideoCapture = videoCapture
  
    IsOK = False
BUGEX "ShowParameterConfig 2"
    Call LoadDriverType
    Call Me.Show(1, owner)
  
    ShowParameterConfig = IsOK
  
BUGEX "ShowParameterConfig 3"
End Function

'modify by tjh at 2010-01-21
'读取当前使用的驱动类型
Private Sub LoadDriverType()
    If mVideoCapture Is Nothing Then Exit Sub
  
BUGEX "LoadDriverType 1"
    Select Case mVideoCapture.VideoDriverType
        Case vdtTWAIN
BUGEX "LoadDriverType 2"
            optDriver(2).value = True
            Call ConfigScan(True, False)
      
        Case vdtVFW
BUGEX "LoadDriverType 3"
            optDriver(1).value = True
            Call ConfigScan(False, False)
      
        Case vdtWDM
BUGEX "LoadDriverType 4"
            optDriver(0).value = True
            Call ConfigScan(False, False)
    
        Case vdtCustom
BUGEX "LoadDriverType 5"
            optDriver(3).value = True
            Call ConfigScan(False, True)
    End Select
  
BUGEX "LoadDriverType 5"
End Sub

Private Sub cboCommCapType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub ConfigComFace(ByVal blnIsCom As Boolean)
'配置com端口设置界面
    cboCommCapType.Enabled = blnIsCom
    txtComInterval.Enabled = blnIsCom
    labCaptureWay.Enabled = blnIsCom
    labComInterval.Enabled = blnIsCom
End Sub

Private Sub cboPort_Click()
    Dim blnIsCom As Boolean
     
    blnIsCom = IIf(InStr(UCase(cboPort.Text), "COM") > 0, True, False)
    
    Call ConfigComFace(blnIsCom)
End Sub

Private Sub cboPort_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    IsOK = False
    
    Unload Me
End Sub

''''''''''''''''''''''''''''''''''
'选择扫描设备的临时图像存储目录
''''''''''''''''''''''''''''''''''
Private Sub cmdDirSelect_Click()
  Dim shl As Object
  Set shl = CreateObject("Shell.application")
  
  On Error GoTo final
  
    Dim fd As Object
    Set fd = shl.BrowseForFolder(0, "扫描设备临时目录选择", 0, "\")
  
    If Not fd Is Nothing Then
        tbxTempDir.Text = fd.Self.Path
    End If
final:
  Set shl = Nothing
  Set fd = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
'显示压缩设置
''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdImageCompressConfig_Click()
On Error GoTo errHandle
    Call imageScannerConfig.ShowScanPreferences
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub LoadStorageDevice()
'载入存储设备
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=1 and NVL(状态,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.EOF Then Exit Sub

    cboSaveDevice.AddItem ""
    cboBakDevice.AddItem ""
    
    Do While Not rsTemp.EOF
        cboSaveDevice.AddItem rsTemp!设备号 & "-" & Nvl(rsTemp!设备名)
        cboBakDevice.AddItem rsTemp!设备号 & "-" & Nvl(rsTemp!设备名)
        
        If GetDeptPara(glngDepartId, "存储设备号", "") = rsTemp!设备号 Then
            cboSaveDevice.ListIndex = cboSaveDevice.NewIndex
        End If
        
        If GetDeptPara(glngDepartId, "备份设备号", "") = rsTemp!设备号 Then
            cboBakDevice.ListIndex = cboBakDevice.NewIndex
        End If
        
        rsTemp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub
 

Private Sub LoadComPort()
'载入com端口及手柄设备
    Dim i As Long
    
    With cboPort
        .Clear
        .AddItem "无"
        .AddItem "COM1"
        .AddItem "COM2"
        .AddItem "COM3"
        .AddItem "COM4"
        .AddItem "COM5"
        .AddItem "COM6"
        .AddItem "COM7"
        .AddItem "COM8"
    End With
    
    Set DxInput = DX7.DirectInputCreate()
    Set DiDevEnum = DxInput.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
    For i = 1 To DiDevEnum.GetCount
        cboPort.AddItem DiDevEnum.GetItem(i).GetInstanceName
    Next
End Sub

Private Sub ReadDepartmentParameter()
'读取科室通用参数配置
    '允许改变采集区域大小
    chkAllowChangeSize.value = Val(GetDeptPara(glngDepartId, "允许改变采集区域大小", 1))
     
End Sub

Private Sub ReadLocateParameter()
'读取本地参数配置(和机器相关的参数配置)
On Error GoTo ErrorHand

    Dim strExeRoom As String
    Dim strDeviceNO As String, iPortNumber As Integer
    Dim iCapType As Integer
    Dim strTmp() As String
    Dim strHotKey As String
    Dim strAfterHotKey As String
    Dim strAfterTagHotKey As String
    Dim strImgMaxSize As String
    Dim intShowBigImg As Integer
    
    If IsNumeric(zlDatabase.GetPara("脚踏端口", glngSys, glngModule, "1")) Then
        iPortNumber = Val(zlDatabase.GetPara("脚踏端口", glngSys, glngModule, "1"))
        cboPort.ListIndex = iPortNumber
    Else
        SeekIndex cboPort, zlDatabase.GetPara("脚踏端口", glngSys, glngModule, "")
    End If
    
    iCapType = Val(zlDatabase.GetPara("脚踏采集方式", glngSys, glngModule, "1"))
    
    If iCapType = 0 Then
        cboCommCapType.ListIndex = 0
    ElseIf iCapType = 1 Then
        cboCommCapType.ListIndex = 1
    Else
        cboCommCapType.ListIndex = 2
    End If
    
    strHotKey = GetSetting("ZLSOFT", "公共模块", "采集热键", "F8")
    If Trim(strHotKey) = "" Then
        cbxHotKey.ListIndex = 0
    Else
        cbxHotKey.Text = strHotKey
    End If
    
    strAfterHotKey = GetSetting("ZLSOFT", "公共模块", "后台采集热键", "F7")
    If Trim(strAfterHotKey) = "" Then
        cboAfterHotKey.ListIndex = 0
    Else
        cboAfterHotKey.Text = strAfterHotKey
    End If
    
    strAfterTagHotKey = GetSetting("ZLSOFT", "公共模块", "标记更新热键", "F6")
    If Trim(strAfterTagHotKey) = "" Then
        cboAfterTagHotKey.ListIndex = 0
    Else
        cboAfterTagHotKey.Text = strAfterTagHotKey
    End If
    
    tbxTempDir.Text = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "扫描设备临时目录", "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA")
    txtDriverPath.Text = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "专用视频采集")
    
    txtComInterval.Text = zlDatabase.GetPara("脚踏时间间隔", glngSys, glngModule, "1")
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    chkCaptureSound.value = zlDatabase.GetPara("采集后声音提示", glngSys, glngModule, "0")
    
    cmdOK.Enabled = InStr(gstrPrivs, "采集参数设置") > 0
    cmdSelectScanDevice.Enabled = InStr(gstrPrivs, "采集参数设置") > 0
    cmdImageCompressConfig.Enabled = InStr(gstrPrivs, "采集参数设置") > 0
    chkUseGameModeCap.value = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "使用游戏手柄采集模式"))
    
    Exit Sub
ErrorHand:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub SaveDepartmentParameter()
'保存科室通用参数配置

    '保存存储设备
    If cboSaveDevice.Text <> "" Then
        SetDeptPara glngDepartId, "存储设备号", Split(cboSaveDevice.Text, "-")(0)
    Else
        SetDeptPara glngDepartId, "存储设备号", ""
    End If
    
    '保存备份设备
    If cboBakDevice.Text <> "" Then
        SetDeptPara glngDepartId, "备份设备号", Split(cboBakDevice.Text, "-")(0)
    Else
        SetDeptPara glngDepartId, "备份设备号", ""
    End If

 
    '视频大小区域设置
    Call SetDeptPara(glngDepartId, "允许改变采集区域大小", chkAllowChangeSize.value)
     
End Sub

Private Sub SaveLocateParameter()
'保存本地参数配置(和机器相关的参数配置)
On Error GoTo errhand
    If optDriver(3).value Then
        If Trim(txtDriverPath.Text) = "" Then
            MsgboxEx hwnd, "专用视频采集时，需指定对应采集接口部件路径！", vbOKOnly, G_STR_HINT_TITLE
            Exit Sub
        End If
    End If
    
    '9以下是COM口,0表示不使用外部设备
    If cboPort.ListIndex = 0 Then
        Call zlDatabase.SetPara("脚踏端口", "无", glngSys, glngModule)
    ElseIf cboPort.ListIndex < 9 Then
        Call zlDatabase.SetPara("脚踏端口", cboPort.ListIndex, glngSys, glngModule)
    Else
        Call zlDatabase.SetPara("脚踏端口", cboPort.Text, glngSys, glngModule)
    End If
    
    '设置采集热键
    Call SaveSetting("ZLSOFT", "公共模块", "采集热键", cbxHotKey.Text)
    Call SaveSetting("ZLSOFT", "公共模块", "后台采集热键", cboAfterHotKey.Text)
    Call SaveSetting("ZLSOFT", "公共模块", "标记更新热键", cboAfterTagHotKey.Text)

    '保存视频驱动类型，目前只有两种驱动类型
    If optDriver(0).value Then Call zlDatabase.SetPara("视频驱动类型", 0, glngSys, glngModule)
    If optDriver(1).value Then Call zlDatabase.SetPara("视频驱动类型", 1, glngSys, glngModule)
    If optDriver(2).value Then Call zlDatabase.SetPara("视频驱动类型", 2, glngSys, glngModule)
    If optDriver(3).value Then Call zlDatabase.SetPara("视频驱动类型", 3, glngSys, glngModule)
     
    Call zlDatabase.SetPara("采集后声音提示", chkCaptureSound.value, glngSys, glngModule)
    Call zlDatabase.SetPara("脚踏采集方式", cboCommCapType.ListIndex, glngSys, glngModule)
    Call zlDatabase.SetPara("脚踏时间间隔", IIf(Val(txtComInterval.Text) = 0, 1, txtComInterval.Text), glngSys, glngModule)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "扫描设备临时目录", tbxTempDir.Text)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "专用视频采集", txtDriverPath.Text)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "使用游戏手柄采集模式", chkUseGameModeCap.value)

    Exit Sub
errhand:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub cmdOk_Click()
  On Error GoTo errHandle
    
    '保存部门参数设置
    Call SaveDepartmentParameter
    
    Call SaveLocateParameter
    
    IsOK = True
    
    Unload Me
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub cmdParameterCfg_Click()
On Error GoTo errHandle
    Call mVideoCapture.ShowCaptureParameterCfgDialog(Me)
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub cmdSelectDriver_Click()
    Dim strCustomDeviceDllName As String    '专用视频采集部件名称
    Dim objCustomDevice As Object           '专用视频采集部件对象
    Dim objFile As New FileSystemObject
    
    On Error GoTo errHandle
    
    dlgOpenDir.ShowOpen
    
    If dlgOpenDir.FileName <> "" Then txtDriverPath.Text = dlgOpenDir.FileName
    
    strCustomDeviceDllName = Trim(Replace(objFile.GetFileName(txtDriverPath.Text), ".dll", ""))
    
    Set objCustomDevice = CreateObject(strCustomDeviceDllName & ".cls" & strCustomDeviceDllName)
    
    If Not objCustomDevice Is Nothing Then Set objCustomDevice = Nothing
    
    Exit Sub
errHandle:
    MsgboxEx hwnd, "指定的专用视频采集接口部件无效，请重新设置！", vbOKOnly, G_STR_HINT_TITLE
    txtDriverPath.Text = ""
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
'扫描设备选择
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSelectScanDevice_Click()
On Error GoTo errHandle
    Call imageScannerConfig.ShowSelectScanner
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hwnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶
    
    '载入脚踏端口配置
    Call LoadComPort
    '载入存储设备
    Call LoadStorageDevice
    
    '读取部门公共参数
    Call ReadDepartmentParameter
    '读取本机参数配置
    Call ReadLocateParameter
End Sub

Private Sub Form_Terminate()
    Set DxInput = Nothing
    Set DX7 = Nothing
    Set DiDevEnum = Nothing
    Set mVideoCapture = Nothing
End Sub

Private Sub optDriver_Click(Index As Integer)
On Error GoTo errHandle
BUGEX "optDriver_Click 1"
    Select Case Index
        Case 0
BUGEX "optDriver_Click 2"
            Call ConfigScan(False, False)
      
            RaiseEvent OnVideoDirverChange(vdtWDM)
        Case 1
BUGEX "optDriver_Click 3"
            Call ConfigScan(False, False)
          
            RaiseEvent OnVideoDirverChange(vdtVFW)
        Case 2
BUGEX "optDriver_Click 4"
            Call ConfigScan(True, False)
      
            RaiseEvent OnVideoDirverChange(vdtTWAIN)
BUGEX "optDriver_Click 5"
        Case 3
            Call ConfigScan(False, True)
            RaiseEvent OnVideoDirverChange(vdtCustom)
    End Select
BUGEX "optDriver_Click 6"
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub ConfigScan(ByVal blnIsScan As Boolean, ByVal blnIsCustom As Boolean)
BUGEX "ConfigScan 1"
    labTempDir.Enabled = blnIsScan
    tbxTempDir.Enabled = blnIsScan
    cmdDirSelect.Enabled = blnIsScan
BUGEX "ConfigScan 2"
    cmdSelectScanDevice.Enabled = blnIsScan
    cmdImageCompressConfig.Enabled = blnIsScan
BUGEX "ConfigScan 3"
    Frame1.Enabled = blnIsScan
    cmdParameterCfg.Enabled = Not blnIsScan
BUGEX "ConfigScan 4"
    txtDriverPath.Enabled = blnIsCustom
    cmdSelectDriver.Enabled = blnIsCustom
End Sub

Private Sub txtComInterval_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtComInterval_KeyPress(KeyAscii As Integer)
    If InStr("0123456789-." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtComInterval_LostFocus()
    If InStr(txtComInterval.Text, "-") = 0 Then
        If Val(txtComInterval.Text) < 1 Or Val(txtComInterval.Text) > 60 Then
            MsgBox "脚踏时间间隔应在1-60秒范围内，请重新输入。", vbOKOnly, G_STR_HINT_TITLE
            txtComInterval.SetFocus
        End If
    End If
End Sub

Private Function GetComboxIndex(aSource() As Variant, ByVal SeekString As String) As Long
    Dim i As Long
    
    For i = 0 To UBound(aSource, 2)
        If aSource(0, i) = SeekString Then Exit For
    Next
    If i > UBound(aSource, 2) Then i = 0
    GetComboxIndex = i
End Function

'控制文本框输入值
Public Sub TxtInputControl(ByRef TxtBox As TextBox, ByRef KeyAscii As Integer, ByVal intDecimalPointNum As Integer)
'txtBox：文本框控件
'intDecimalPointNum：小数点位数
'KeyAscii:输入的ASC

    If Chr(KeyAscii) = "." Then
        If InStr(TxtBox.Text, ".") > 0 Then KeyAscii = 0
    End If
    
    If InStr(TxtBox.Text, ".") > 0 And KeyAscii <> 8 Then
        If Len(Mid(TxtBox.Text, InStr(TxtBox.Text, ".") + 1)) >= intDecimalPointNum Then KeyAscii = 0
    End If
    
End Sub

 
