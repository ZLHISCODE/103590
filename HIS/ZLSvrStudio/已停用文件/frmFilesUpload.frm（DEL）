VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFilesUpload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "升级文件上传"
   ClientHeight    =   9888
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   15912
   Icon            =   "frmFilesUpload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10296.87
   ScaleMode       =   0  'User
   ScaleWidth      =   15915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picHelp 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1000
      Left            =   0
      ScaleHeight     =   996
      ScaleWidth      =   15912
      TabIndex        =   18
      Top             =   0
      Width           =   15915
      Begin VB.PictureBox picState 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   780
         Index           =   0
         Left            =   1290
         ScaleHeight     =   780
         ScaleWidth      =   8628
         TabIndex        =   28
         Top             =   135
         Visible         =   0   'False
         Width           =   8625
         Begin VB.Label lblEXP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "安装路径缺失：文件的安装路径为空，请修改为有效文件"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   5
            Left            =   15
            TabIndex        =   31
            Top             =   540
            Width           =   4500
         End
         Begin VB.Label lblEXP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "本地文件缺失：请确认该文件路径下文件存在"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   4
            Left            =   15
            TabIndex        =   30
            Top             =   0
            Width           =   3600
         End
         Begin VB.Label lblEXP 
            BackStyle       =   0  'Transparent
            Caption         =   "标准部件缺失：标准文件在当前在用清单中缺失，请检查在用文件清单或修复当前在用文件清单"
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   3
            Left            =   15
            TabIndex        =   29
            Top             =   270
            Width           =   8505
         End
      End
      Begin VB.PictureBox picState 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   780
         Index           =   1
         Left            =   1290
         ScaleHeight     =   780
         ScaleWidth      =   8628
         TabIndex        =   32
         Top             =   135
         Visible         =   0   'False
         Width           =   8625
         Begin VB.Label lblEXP 
            BackStyle       =   0  'Transparent
            Caption         =   "本地文件与标准文件清单的文件不匹配，可能存在风险，请仔细检查警告文件保证兼容性"
            ForeColor       =   &H00007FFF&
            Height          =   225
            Index           =   8
            Left            =   15
            TabIndex        =   35
            Top             =   270
            Width           =   8505
         End
         Begin VB.Label lblEXP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "警告：本地文件与标准文件不符，请确保文件兼容"
            ForeColor       =   &H00007FFF&
            Height          =   180
            Index           =   7
            Left            =   15
            TabIndex        =   34
            Top             =   0
            Width           =   3960
         End
         Begin VB.Label lblEXP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "请确保本地文件正确性，否则可能会造成一些问题"
            ForeColor       =   &H00007FFF&
            Height          =   180
            Index           =   6
            Left            =   30
            TabIndex        =   33
            Top             =   540
            Width           =   3960
         End
      End
      Begin VB.Frame fraBounds 
         Height          =   1170
         Index           =   1
         Left            =   10020
         TabIndex        =   27
         Top             =   -135
         Width           =   30
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主要流程："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   10230
         TabIndex        =   26
         Top             =   405
         Width           =   900
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上传文件"
         Height          =   180
         Index           =   2
         Left            =   14700
         TabIndex        =   25
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收集文件"
         Height          =   180
         Index           =   1
         Left            =   13005
         TabIndex        =   24
         Top             =   75
         Width           =   720
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查文件"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   11250
         TabIndex        =   23
         Top             =   75
         Width           =   720
      End
      Begin VB.Image imgPro 
         Height          =   540
         Index           =   4
         Left            =   13872
         Picture         =   "frmFilesUpload.frx":6852
         Top             =   288
         Width           =   540
      End
      Begin VB.Image imgPro 
         Height          =   540
         Index           =   3
         Left            =   12156
         Picture         =   "frmFilesUpload.frx":807E
         Top             =   288
         Width           =   540
      End
      Begin VB.Image imgPro 
         Height          =   576
         Index           =   2
         Left            =   14688
         Picture         =   "frmFilesUpload.frx":98AA
         Top             =   276
         Width           =   576
      End
      Begin VB.Image imgPro 
         Height          =   576
         Index           =   1
         Left            =   12996
         Picture         =   "frmFilesUpload.frx":B3EC
         Top             =   276
         Width           =   576
      End
      Begin VB.Image imgPro 
         Height          =   576
         Index           =   0
         Left            =   11232
         Picture         =   "frmFilesUpload.frx":CF2E
         Top             =   252
         Width           =   576
      End
      Begin VB.Image imgCaption 
         Height          =   576
         Left            =   300
         Picture         =   "frmFilesUpload.frx":EA70
         Top             =   120
         Width           =   576
      End
      Begin VB.Label lblEXP 
         BackStyle       =   0  'Transparent
         Caption         =   "收集文件：需要将当前环境可上传的文件压缩致临时文件夹，会智能判断已经压缩过的文件，加快收集过程"
         Height          =   225
         Index           =   1
         Left            =   1305
         TabIndex        =   21
         Top             =   405
         Width           =   8505
      End
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查文件：主要检查文件是否异常(本地缺失，标准部件缺失等)，检查文件是否需要上传"
         Height          =   180
         Index           =   0
         Left            =   1305
         TabIndex        =   20
         Top             =   135
         Width           =   7020
      End
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上传文件：会依次向已经启用的文件夹上传当前环境文件"
         Height          =   180
         Index           =   2
         Left            =   1300
         TabIndex        =   19
         Top             =   675
         Width           =   4500
      End
   End
   Begin VB.Frame fraBounds 
      Height          =   30
      Index           =   0
      Left            =   -1695
      TabIndex        =   22
      Top             =   1005
      Width           =   17730
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   13815
      Top             =   1530
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   1048
      ImageHeight     =   27
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":105B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":2519C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":39D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":4E970
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":6355A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":78144
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":8CD2E
            Key             =   "文件"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":8E880
            Key             =   "检查文件"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":903D2
            Key             =   "收集文件"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":91F24
            Key             =   "上传文件"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":93A76
            Key             =   "文件异常"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":955C8
            Key             =   "警告"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":9711A
            Key             =   "异常"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picInformation 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   324
      Left            =   24
      ScaleHeight     =   405.238
      ScaleMode       =   0  'User
      ScaleWidth      =   15847.55
      TabIndex        =   7
      Top             =   9360
      Width           =   15840
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已上传文件："
         Height          =   180
         Index           =   5
         Left            =   13470
         TabIndex        =   13
         Top             =   90
         Width           =   1080
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查警告文件："
         Height          =   180
         Index           =   4
         Left            =   10890
         TabIndex        =   12
         Top             =   75
         Width           =   1260
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "需上传新文件："
         Height          =   180
         Index           =   3
         Left            =   8565
         TabIndex        =   11
         Top             =   90
         Width           =   1260
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "可上传文件："
         Height          =   180
         Index           =   2
         Left            =   5685
         TabIndex        =   10
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态异常文件："
         Height          =   180
         Index           =   1
         Left            =   3045
         TabIndex        =   9
         Top             =   120
         Width           =   1260
      End
      Begin VB.Label lblInformation 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "所有上传文件："
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   465
         TabIndex        =   8
         Top             =   90
         Width           =   1260
      End
      Begin VB.Image imgInformation 
         Height          =   324
         Left            =   12
         Picture         =   "frmFilesUpload.frx":98C6C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15684
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfMain 
      Height          =   7185
      Left            =   60
      TabIndex        =   3
      Top             =   2055
      Visible         =   0   'False
      Width           =   15690
      _cx             =   27675
      _cy             =   12674
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483638
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "上传新的文件(&Q)"
      Enabled         =   0   'False
      Height          =   288
      Left            =   12480
      TabIndex        =   2
      Top             =   1155
      Width           =   1500
   End
   Begin VB.CommandButton cmdAllUpLoad 
      Caption         =   "上传所有文件(&T)"
      Height          =   288
      Left            =   14205
      TabIndex        =   14
      Top             =   1155
      Width           =   1500
   End
   Begin VB.CommandButton cmdMD5Check 
      Caption         =   "重新检查(&D)"
      Height          =   288
      Left            =   11055
      TabIndex        =   5
      Top             =   1155
      Width           =   1200
   End
   Begin MSComctlLib.ProgressBar pgbThis 
      Height          =   330
      Left            =   495
      TabIndex        =   4
      Top             =   1560
      Width           =   9615
      _ExtentX        =   16955
      _ExtentY        =   572
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&Q)"
      Height          =   288
      Left            =   8085
      TabIndex        =   1
      Top             =   1155
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSever 
      Height          =   1155
      Left            =   14490
      TabIndex        =   6
      Top             =   1515
      Visible         =   0   'False
      Width           =   1305
      _cx             =   2302
      _cy             =   2037
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblstatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   1260
      Width           =   90
   End
   Begin VB.Label lblstatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   540
      TabIndex        =   16
      Top             =   1215
      Width           =   120
   End
   Begin VB.Label lblstatus 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   1245
      Width           =   90
   End
   Begin VB.Label lblstatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   1620
      TabIndex        =   0
      Top             =   1260
      Width           =   90
   End
End
Attribute VB_Name = "frmFilesUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'StateDisplay-状态显示
Private Const SDP_准备就绪 = "准备就绪"
Private Const SDP_准备上传 = "准备上传"
Private Const SDP_无需更新 = "无需更新"
Private Const SDP_状态异常 = "状态异常"
Private Const SDP_已经上传 = "已经上传"
Private Const SDP_上传失败 = "上传失败"
Private Const SDP_收集完成 = "收集完成"
Private Const SDP_无需收集 = "无需收集"

'StateColor-状态颜色
Private Const SC_红色 = vbRed
Private Const SC_绿色 = 2188065  'RGB(33, 99, 33)
Private Const SC_蓝色 = 9109504  'RGB(0, 0, 139)
Private Const SC_黄色 = 32767      'RGB(255, 127, 0)

'文件状态
Private Enum FilesState
    FS_默认正常 = 0 '允许上传
    FS_状态异常 = 1 '异常不允许上传
    FS_无需更新 = 2 '不用上传
    FS_无需更新警告文件 = 3
    FS_准备上传警告文件 = 4
    FS_已经上传 = 5 '上传成功
    FS_上传失败 = -1 '上传失败
End Enum

Private Enum UpLoadCol
    Col_序号 = 0
    Col_文件 = 1
    Col_状态 = 2 '状态值 0-正常 1-升级部件缺失(本地文件不存在) 2-本地文件不存在 3-无需更新 4-警告但可以上传 5-已经上传 6-上传失败 7 -警告但不用上传
    Col_警告 = 3
    Col_当前版本 = 4
    Col_标准版本 = 5
    Col_安装路径 = 6
    Col_系统 = 7
    Col_修改日期 = 8
    Col_业务部件 = 9
    Col_文件说明 = 10
    Col_当前md5 = 11
    Col_标准md5 = 12
    Col_本地md5 = 13
    Col_文件地址 = 14
    Col_收集地址 = 15  '收集后文件地址
    Col_收集文件 = 16 '收集后文件名称
    Col_文件类型 = 17
    Col_列数 = 18
End Enum

Private Enum UpSeverCol
    Col_编号 = 0
    Col_类型 = 1
    Col_地址 = 2
    Col_用户名 = 3
    Col_密码 = 4
    Col_端口 = 5
    Col_上传状态 = 6
    Col_服务器列数 = 7
End Enum

Private Enum UploadResult
    Res_未上传 = 0
    Res_上传成功 = 1
    Res_上传失败 = 2
    Res_未知错误 = 3
End Enum

Private Enum lblItemNum
    LN_所有上传文件 = 0
    LN_状态异常文件 = 1
    LN_可上传文件 = 2
    LN_需上传新文件 = 3
    LN_检查警告文件 = 4
    LN_已上传文件 = 5
'    LN_无需更新文件 = 5
    LN_lbl总数 = 6
End Enum

Private Enum pbgstatus
    Sta_百分比 = 0
    Sta_当前操作 = 1
    Sta_操作对象 = 2
    Sta_状态描述 = 3
End Enum

Private mrsTemp As New ADODB.Recordset
Private strParstWarning As String '缺失部件警告
Private mstrSQL As String
Private mintUpFilesCount As Integer '上传文件总数
Private mblnCheckMD5Tag As Boolean '正在检查MD5标志
Private mblnUploadTag As Boolean '正在上传标志
Private mblnAllUploadTag As Boolean '正在全部上传标志
Private mblnAllUpload As Boolean

Private mobjFile As New FileSystemObject
Private mstrScratchFilePath As String '临时文件目录
Private mblnUpLoadSuccess As Boolean '上传成功，更新MD5标志
Private mstrSuccessUploadSever As String '成功上传服务器
Private mcllPath As Collection '安装路径转换实际路径集合

Private mlngAbnormal  As Long '异常
Private mlngCorrect As Long '正常
Private mlngUnchanged As Long '无需更新
Private mlngWarning As Long  '警告
Private mlngUpload As Long
Private mlngNewFile As Long

Public Function ShowMe() As Boolean
    '临时目录初始化
    If gblnInIDE Then
        mstrScratchFilePath = "C:\APPSOFT\ZLUPTMP"
    Else
        mstrScratchFilePath = App.Path & "\ZLUPTMP"
    End If
    Me.Show 1, frmMDIMain
End Function

Private Sub cmdAllUpLoad_Click()
    Dim strSQL As String
    
    strSQL = "update zlfilesupgrade set MD5 = Null"
    gcnOracle.Execute strSQL
    '环境还原
    mblnAllUpload = True
    Call AllUploadRestore
'    Call cmdMD5Check_Click
    Call cmdUpload_Click
    mblnAllUpload = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMD5Check_Click()
    '检查过MD5文件路径检测
    Set mcllPath = CheckAndAdjustFolder()
    ShowStatus "", "", "开始检查", 0
    lblInformation_Click (LN_所有上传文件)
    Call DataLoad
    Call DataCheck
    Call FilesMD5Check
    '检查完成后自动切换界面
    If Val(Split(lblInformation(LN_状态异常文件), "：")(1)) > 0 Then
        lblInformation_Click (LN_状态异常文件)
        ShowStatus "", "", "检查完成，请检查异常文件", 0, , False
    ElseIf Val(Split(lblInformation(LN_检查警告文件), "：")(1)) > 0 Then
        lblInformation_Click (LN_检查警告文件)
        ShowStatus "", "", "检查完成，请检查警告文件", 0, , False
    ElseIf Val(Split(lblInformation(LN_需上传新文件), "：")(1)) > 0 Then
        lblInformation_Click (LN_需上传新文件)
        ShowStatus "", "", "检查完成可以上传新的文件", 0, , False
    Else
        lblInformation_Click (LN_可上传文件)
        ShowStatus "", "", "检查完成可以上传文件", 0, , False
    End If
End Sub

Private Sub cmdUpload_Click()
    Dim strNumber As String
    Dim strSeverType As String
    Dim strServerAddress As String
    Dim strUser As String
    Dim strPassword As String
    Dim strPort As String
    Dim strBatch As String
    Dim i As Integer
    Dim strErrInfor As String '错误信息
    On Error GoTo errH
    
    If mblnAllUpload Then
        Call lblInformation_Click(LN_可上传文件)
    Else
        Call lblInformation_Click(LN_需上传新文件)
    End If
    Call ControlVisible(False)
    '收集文件，压缩至临时文件夹
    If FilesCollections() = True Then
        '清空7z.exe残余系统进程
        Call fun_KillProcess(PROAPPCTION)
        '拷贝收集文件到剪贴板
'        Call FloderToClipBoard(mstrScratchFilePath)
    End If
    
    '读取服务器列表
    Call SeverDataLoad
    imgCaption.Picture = imgList.ListImages("上传文件").Picture
    lblEXP(2).ForeColor = vbBlue
    With vsfSever
        If .Rows < 1 Then MsgBox "请设置至少一个上传服务器!", vbDefaultButton1 + vbInformation, gstrSysName: Exit Sub
        For i = 1 To .Rows - 1
            strNumber = .TextMatrix(i, Col_编号): strSeverType = .TextMatrix(i, Col_类型)
            strServerAddress = .TextMatrix(i, Col_地址): strUser = .TextMatrix(i, Col_用户名)
            strPassword = Decipher(Trim(.TextMatrix(i, Col_密码))): strPort = .TextMatrix(i, Col_端口)
            
            Select Case strSeverType
                Case "0" '共享
                    If CopyFileToShareServer(strNumber, strServerAddress, strUser, strPassword, strErrInfor) = True Then
                        .TextMatrix(i, Col_上传状态) = str(Res_上传成功)
                    Else
                        If strErrInfor <> "" Then
                            MsgBox strErrInfor, vbDefaultButton1 + vbInformation, gstrSysName
                            .TextMatrix(i, Col_上传状态) = str(Res_未知错误)
                       Else
                            .TextMatrix(i, Col_上传状态) = str(Res_上传失败)
                       End If
                    End If
                Case "1" 'FTP
                    If CopyFileToFTPServer(strNumber, strServerAddress, strUser, strPassword, strPort, strErrInfor) = True Then
                        .TextMatrix(i, Col_上传状态) = str(Res_上传成功)
                    Else
                        If strErrInfor <> "" Then
                            MsgBox strErrInfor, vbDefaultButton1 + vbInformation, gstrSysName
                            .TextMatrix(i, Col_上传状态) = str(Res_未知错误)
                       Else
                            .TextMatrix(i, Col_上传状态) = str(Res_上传失败)
                       End If
                    End If
            End Select
        Next
        If UpdateMD5 = True Then '更新MD5
            strBatch = BatchLoad
            strBatch = Trim(str(Nvl(strBatch) + 1))
            BatchUpdate strBatch
        End If
        imgCaption.Picture = imgList.ListImages("文件").Picture
        lblEXP(2).ForeColor = vbBlack
        Call ControlVisible(True)
    End With
    cmdUpload.Enabled = False
    Call UpLoadFilesCount
    Call lblInformation_Click(LN_已上传文件)
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    Call ControlVisible(True)
    If False Then
        Resume
    End If
End Sub

Private Sub Command1_Click()
    Dim Data1 As String
    Dim Data2 As String
    Dim strFilePath As String
    Dim i As Long
    
    With vsfMain
        i = 2
        strFilePath = .TextMatrix(i, Col_文件地址)
        Data1 = Format(FileDateTime(strFilePath), "yyyy-MM-DD hh:mm:ss")
        Data2 = .TextMatrix(i, Col_修改日期)
        Data2 = "2000-04-22 20:10:56"
        Call CompareDate(Data1, Data2)
    End With
End Sub

Private Sub Form_Activate()
    Dim lngWidth As Long
    Dim i As Integer
    
    '控件视图初始化
    pgbThis.Move 100, 600 + picHelp.Height, Me.ScaleWidth - 200, 330
    vsfMain.Move 100, 1000 + picHelp.Height, Me.ScaleWidth - 200, Me.ScaleHeight - 1490 - picHelp.Height

    picInformation.Move 100, 9800, Me.ScaleWidth - 200, 400
    imgInformation.Move 0, 0, picInformation.ScaleWidth, picInformation.ScaleHeight
    imgInformation.Picture = imgList.ListImages.Item(LN_所有上传文件 + 1).Picture

    lngWidth = picInformation.Width / 6
    For i = 0 To LN_lbl总数 - 1
        lblInformation.Item(i).Move (i * lngWidth) + ((lngWidth) - lblInformation.Item(i).Width) / 2, (picInformation.ScaleHeight - lblInformation.Item(i).Height) / 2
    Next
    lblInformation.Item(LN_所有上传文件).ForeColor = vbBlack
    lblInformation.Item(LN_所有上传文件).FontBold = True
    lblInformation.Item(LN_状态异常文件).ForeColor = vbRed
    lblInformation.Item(LN_可上传文件).ForeColor = SC_绿色
    lblInformation.Item(LN_检查警告文件).ForeColor = SC_黄色
    lblInformation.Item(LN_已上传文件).ForeColor = SC_蓝色
    lblInformation.Item(LN_需上传新文件).ForeColor = SC_绿色
'    lblInformation.Item(LN_无需更新文件).ForeColor = SC_蓝色
    vsfMain.Visible = True
    Call cmdMD5Check_Click
End Sub

Private Sub Form_Load()
'   Call cmdMD5Check_Click
End Sub

Private Sub Form_Resize()
'    vsfMain.Move 50, 1200, Me.ScaleWidth - 100, Me.ScaleHeight - 1000
End Sub

Private Function GetSystemName(ByVal strNum As String) As String
'传入系统编号，获得对应系统名称，若未找到
On err GoTo errH
    Select Case strNum
        Case "1", "100"
            GetSystemName = "医院系统标准版"
        Case "2", "200"
            GetSystemName = "人事工资系统"
        Case "3", "300"
            GetSystemName = "病案管理系统"
        Case "4", "400"
            GetSystemName = "物资供应系统"
        Case "5", "500"
            GetSystemName = "财务核算系统"
        Case "6", "600"
            GetSystemName = "设备管理系统"
        Case "7", "700"
            GetSystemName = "成本效益核算系统"
        Case "21", "2100"
            GetSystemName = "体检管理系统"
        Case "22", "2200"
            GetSystemName = "血库管理系统"
        Case "23", "2300"
            GetSystemName = "院感管理系统"
        Case "24", "2400"
            GetSystemName = "手麻管理系统"
        Case "25", "2500"
            GetSystemName = "临床检验管理系统"
        Case "26", "2600"
            GetSystemName = "病人自助服务系统"
    End Select
    Exit Function

errH:
    If False Then
        Resume
    End If
End Function

'数据读取，表格视图加载
Public Sub DataLoad()
    Dim i, j As Long
    Dim strSQL       As String
    Dim strTemp As String
    Dim arrSys As Variant
    On Error GoTo errH

    With vsfMain
        .Redraw = flexRDNone
        .Tag = ""
'        .Redraw = flexRDNone
        .Rows = 1
        .Clear
        .Cols = Col_列数
'        Exit Sub
        .TextMatrix(0, Col_序号) = ""
        .Cell(flexcpAlignment, 0, Col_序号) = flexAlignCenterCenter
        .ColWidth(Col_序号) = 400
        
        .TextMatrix(0, Col_状态) = "状态"
        .Cell(flexcpAlignment, 0, Col_状态) = flexAlignCenterCenter
        .ColWidth(Col_状态) = 1000
        
        .TextMatrix(0, Col_警告) = "文件检查警告"
        .Cell(flexcpAlignment, 0, Col_警告) = flexAlignCenterCenter
        .ColWidth(Col_警告) = 4200
        
        .TextMatrix(0, Col_文件) = "文件"
        .Cell(flexcpAlignment, 0, Col_文件) = flexAlignCenterCenter
        .ColWidth(Col_文件) = 2400
        
        .TextMatrix(0, Col_当前版本) = "当前版本"
        .Cell(flexcpAlignment, 0, Col_当前版本) = flexAlignCenterCenter
        .ColWidth(Col_当前版本) = 900
        
        .TextMatrix(0, Col_标准版本) = "标准版本"
        .Cell(flexcpAlignment, 0, Col_标准版本) = flexAlignCenterCenter
        .ColWidth(Col_标准版本) = 900
        
        .TextMatrix(0, Col_安装路径) = "安装路径"
        .Cell(flexcpAlignment, 0, Col_安装路径) = flexAlignCenterCenter
        .ColWidth(Col_安装路径) = 1800
        
        .TextMatrix(0, Col_系统) = "系统"
        .Cell(flexcpAlignment, 0, Col_系统) = flexAlignCenterCenter
        .ColWidth(Col_系统) = 1000
        .ColHidden(Col_系统) = True
        
        .TextMatrix(0, Col_修改日期) = "修改日期"
        .Cell(flexcpAlignment, 0, Col_修改日期) = flexAlignCenterCenter
        .ColWidth(Col_修改日期) = 1800

        .TextMatrix(0, Col_业务部件) = "业务部件"
'        .Cell(flexcpAlignment, 0, Col_业务部件) = flexAlignCenterCenter
        .ColWidth(Col_业务部件) = 1000
        .ColHidden(Col_业务部件) = True
        
        .TextMatrix(0, Col_文件说明) = "文件说明"
        .Cell(flexcpAlignment, 0, Col_文件说明) = flexAlignCenterCenter
        .ColWidth(Col_文件说明) = 400
        
        .TextMatrix(0, Col_当前md5) = "当前md5"
        .ColWidth(Col_当前md5) = 10
        .ColHidden(Col_当前md5) = True
        
        .TextMatrix(0, Col_标准md5) = "当前md5"
        .ColWidth(Col_标准md5) = 10
        .ColHidden(Col_标准md5) = True

        .TextMatrix(0, Col_本地md5) = "本地md5"
        .ColWidth(Col_本地md5) = 10
        .ColHidden(Col_本地md5) = True
        
        .TextMatrix(0, Col_文件地址) = "文件地址"
        .ColWidth(Col_文件地址) = 10
        .ColHidden(Col_文件地址) = True
        
        .TextMatrix(0, Col_收集地址) = "收集文件地址"
        .ColWidth(Col_收集地址) = 10
        .ColHidden(Col_收集地址) = True
        
        .TextMatrix(0, Col_收集文件) = "收集文件名称"
        .ColWidth(Col_收集文件) = 10
        .ColHidden(Col_收集文件) = True
        
        .TextMatrix(0, Col_文件类型) = "收集文件名称"
        .ColWidth(Col_文件类型) = 10
        .ColHidden(Col_文件类型) = True

'        strSQL = "Select A.文件名 As 文件, A.版本号 As 版本号,b.版本号 as 标准版本, A.安装路径 As 安装路径, A.所属系统 As 系统,A.修改日期 As 修改日期, A.业务部件 As 业务部件, A.文件说明 As 文件说明 " & _
'                      "From zlFilesUpgrade A Left Join Zlfiles B " & _
'                      "On A.文件名 = B.名称 " & _
'                      "order by 文件"
        strSQL = "Select Nvl(a.文件名, b.名称) As 部件名称, a.文件版本号 As 当前版本, b.版本号 As 标准版本, a.安装路径 As 安装路径, a.所属系统 As 系统, a.修改日期 As 修改日期, " & _
                      "a.业务部件 As 业务部件, a.文件说明 As 文件说明, a.Md5 As 当前md5, b.标准md5, Decode(a.文件名, Null, 1, Decode(b.名称, Null, 2)) As 警告, Decode(a.文件类型,null,b.文件类型,a.文件类型) as 文件类型 " & _
                      "From zlFilesUpgrade A Full Join Zlfiles B " & _
                      "On A.文件名 = B.名称 " & _
                      "order by 部件名称"
        Call OpenRecordset(mrsTemp, strSQL, Me.Caption)

        '数据填入
        .Rows = mrsTemp.RecordCount + 1
        i = 1
        Do Until mrsTemp.EOF
            .TextMatrix(i, Col_序号) = i
            .Cell(flexcpAlignment, i, Col_序号) = flexAlignCenterCenter
            
            .TextMatrix(i, Col_状态) = ""  '状态
            .Cell(flexcpAlignment, i, Col_状态) = flexAlignCenterCenter
  
            .TextMatrix(i, Col_警告) = ""
            .Cell(flexcpData, i, Col_警告) = Trim(Nvl(mrsTemp.Fields("警告"), ""))
            .Cell(flexcpAlignment, i, Col_警告) = flexAlignLeftCenter
            
            .TextMatrix(i, Col_文件) = Nvl(mrsTemp.Fields("部件名称"), "")
            .Cell(flexcpAlignment, i, Col_文件) = flexAlignLeftCenter
            
            
            strTemp = Nvl(mrsTemp.Fields("当前版本"), "")
'            .Cell(flexcpData, i, Col_当前版本) = strTemp '未转换版本号
'            strTemp = GetFileVision(strTemp)
            .TextMatrix(i, Col_当前版本) = strTemp  '转换后版本号
            .Cell(flexcpAlignment, i, Col_当前版本) = flexAlignLeftCenter
            
            strTemp = Nvl(mrsTemp.Fields("标准版本"), "")
'            .Cell(flexcpData, i, Col_标准版本) = strTemp
'            strTemp = GetFileVision(strTemp)
            .TextMatrix(i, Col_标准版本) = strTemp
            .Cell(flexcpAlignment, i, Col_标准版本) = flexAlignLeftCenter
            
            .TextMatrix(i, Col_安装路径) = Nvl(mrsTemp.Fields("安装路径"), "")
            .Cell(flexcpAlignment, i, Col_安装路径) = flexAlignLeftCenter

            strTemp = Nvl(mrsTemp.Fields("系统"), "")

            If Trim(strTemp) <> "" Then
                arrSys = Split(Trim(strTemp), ",")
                strTemp = ""
                For j = 0 To UBound(arrSys)
                    If GetSystemName(arrSys(j)) <> "" Then strTemp = strTemp & "，" & GetSystemName(arrSys(j))
                Next
                strTemp = Mid(strTemp, 2)
            Else
                strTemp = ""
            End If
            .TextMatrix(i, Col_系统) = strTemp
            .Cell(flexcpAlignment, i, Col_系统) = flexAlignLeftCenter

            .TextMatrix(i, Col_修改日期) = Nvl(mrsTemp.Fields("修改日期"), "")
            .Cell(flexcpAlignment, i, Col_修改日期) = flexAlignLeftCenter
            
            .TextMatrix(i, Col_业务部件) = Nvl(mrsTemp.Fields("业务部件"), "")
            .Cell(flexcpAlignment, i, Col_业务部件) = flexAlignLeftCenter
            
            .TextMatrix(i, Col_文件说明) = Nvl(mrsTemp.Fields("文件说明"), "")
            .Cell(flexcpAlignment, i, Col_文件说明) = flexAlignLeftCenter
            
            .TextMatrix(i, Col_当前md5) = Trim(Nvl(mrsTemp.Fields("当前md5"), ""))

            .TextMatrix(i, Col_标准md5) = Trim(Nvl(mrsTemp.Fields("标准md5"), ""))
            
            .TextMatrix(i, Col_文件类型) = Trim(Nvl(mrsTemp.Fields("文件类型"), ""))
            
            mrsTemp.MoveNext
            i = i + 1
        Loop
        
        '选中框风格
        .FocusRect = flexFocusSolid
        '最后一列自动列宽
        .ExtendLastCol = True
        '滚动画面跟随
        .ScrollTrack = True
        '自动换行
        .WordWrap = True
        '行高设置
        .RowHeightMin = 300
        .RowHeightMax = 300
        '最大宽度设置
        .ColWidthMax = 7000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .Redraw = flexRDBuffered
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

'读取需要上传的服务器
Public Sub SeverDataLoad()
    Dim i, j As Long
    Dim strSQL       As String
    Dim strTemp As String
    On Error GoTo errH

    With vsfSever
        .Tag = ""
'        .Redraw = flexRDNone
        .Rows = 1
        .Clear
        .Cols = Col_服务器列数
'        Exit Sub
        .TextMatrix(0, Col_编号) = "服务器编号"
        .Cell(flexcpAlignment, 0, Col_编号) = flexAlignCenterCenter
        .ColWidth(Col_编号) = 800
        
        .TextMatrix(0, Col_类型) = "服务器类型"
        .Cell(flexcpAlignment, 0, Col_类型) = flexAlignCenterCenter
        .ColWidth(Col_类型) = 800
        
        .TextMatrix(0, Col_地址) = "服务器地址"
        .Cell(flexcpAlignment, 0, Col_地址) = flexAlignCenterCenter
        .ColWidth(Col_地址) = 3000
        
        .TextMatrix(0, Col_用户名) = "用户名"
        .Cell(flexcpAlignment, 0, Col_用户名) = flexAlignCenterCenter
        .ColWidth(Col_用户名) = 1500
        
        .TextMatrix(0, Col_密码) = "密码"
        .Cell(flexcpAlignment, 0, Col_密码) = flexAlignCenterCenter
        .ColWidth(Col_密码) = 1500
        
        .TextMatrix(0, Col_端口) = "端口"
        .Cell(flexcpAlignment, 0, Col_端口) = flexAlignCenterCenter
        .ColWidth(Col_端口) = 1600
        
        .TextMatrix(0, Col_上传状态) = "状态"
        .Cell(flexcpAlignment, 0, Col_上传状态) = flexAlignCenterCenter
        .ColWidth(Col_上传状态) = 1600

        strSQL = "Select 编号 As 编号, 类型 As 类型, 位置 As 地址, 用户名 As 用户名, 密码 As 密码, 端口 As 端口 " & _
                      "From Zlupgradeserver " & _
                      "Where 是否升级 = 1 " & _
                      "order by 编号"
        Call OpenRecordset(mrsTemp, strSQL, Me.Caption)

        '数据填入
        .Rows = mrsTemp.RecordCount + 1
        i = 1
        Do Until mrsTemp.EOF
            .TextMatrix(i, Col_编号) = Trim(Nvl(mrsTemp.Fields("编号"), ""))

            .TextMatrix(i, Col_类型) = Trim(Nvl(mrsTemp.Fields("类型"), ""))
    
            .TextMatrix(i, Col_地址) = Trim(Nvl(mrsTemp.Fields("地址"), ""))
            
            .TextMatrix(i, Col_用户名) = Trim(Nvl(mrsTemp.Fields("用户名"), ""))

            .TextMatrix(i, Col_密码) = Trim(Nvl(mrsTemp.Fields("密码"), ""))

            .TextMatrix(i, Col_端口) = Trim(Nvl(mrsTemp.Fields("端口"), ""))
            
            .TextMatrix(i, Col_上传状态) = Trim(str(Res_未上传))
            
            mrsTemp.MoveNext
            i = i + 1
        Loop
        
    End With
    Exit Sub
errH:
    If False Then
        Resume
    End If
End Sub

'MD5检查
Public Function FilesMD5Check()
    Dim strMD5 As String '本地文件MD5
    Dim strMD5Upgrade As String '升级文件MD5(当前)
    Dim strMD5Standard As String '标准文件MD5
    Dim lngPercent As Long '进度条百分比
    Dim intPercent As Integer '文字显示百分比
    Dim i As Long
    Dim strSQL As String
    Dim strTemp As String
    
    On Error Resume Next
    Call ControlVisible(False)
    imgCaption.Picture = imgList.ListImages("检查文件").Picture
    lblEXP(0).ForeColor = vbBlue
    With vsfMain
        If .Rows < .FixedRows Then Exit Function
        Call ShowStatus("正在检查", "", "", 0, .Rows - 1)
        mlngCorrect = 0: mlngUnchanged = 0: mlngWarning = 0: mlngNewFile = 0

        For i = .FixedRows To .Rows - 1
'            .Row = i
'            If .Rows - (i + 14) > 0 Then '保持选中行在中间位置
'                .ShowCell i + 14, Col_文件
'            Else
'                .ShowCell i, Col_文件
'            End If
            ShowCenterRow i
            
            Call ShowStatus("正在检查", .TextMatrix(i, Col_文件), "", i)
            If .Cell(flexcpData, i, Col_状态) = "0" Then
                strMD5Standard = .TextMatrix(i, Col_标准md5)
                strMD5Upgrade = .TextMatrix(i, Col_当前md5)
                DoEvents '防止界面卡顿
                strMD5 = FileMD5(Trim(UCase(.TextMatrix(i, Col_文件地址))))
                .TextMatrix(i, Col_本地md5) = strMD5
                
                If strMD5Upgrade = strMD5 Then '本地与升级MD5相同不需要更新
                    If strMD5 = strMD5Standard Then
                        Call FileStateSet(FS_无需更新, i, "文件不存在差异,不需要更新")
                    Else
                        If strMD5Standard = "" Then '标准部件表不存在该文件
                            strTemp = "警告：标准文件清单(zlFiles)中不存在该文件"
                            If .TextMatrix(i, Col_文件类型) = "4" Then strTemp = strTemp & "(三方部件)"
                        Else
                            strTemp = "警告：本地文件与标准文件不符，请确保文件兼容"
                        End If
                        Call FileStateSet(FS_无需更新警告文件, i, strTemp)
                        mlngWarning = mlngWarning + 1
                    End If
                    mlngUnchanged = mlngUnchanged + 1
                    mlngCorrect = mlngCorrect + 1
                Else '本地与升级MD5不相同需要更新
                    If strMD5 <> strMD5Standard Then '本地MD5与标准MD5不同
                        If strMD5Standard = "" Then '标准部件表不存在该文件
                            strTemp = "警告：标准文件清单(zlFiles)中不存在该文件"
                            If .TextMatrix(i, Col_文件类型) = "4" Then strTemp = strTemp & "(三方部件)"
                        Else
                            strTemp = "警告：本地文件与标准文件不符，请确保文件兼容"
                        End If
                        Call FileStateSet(FS_准备上传警告文件, i, strTemp)
                        mlngWarning = mlngWarning + 1
                        mlngCorrect = mlngCorrect + 1
                        mlngNewFile = mlngNewFile + 1
                    Else
                        Call FileStateSet(FS_默认正常, i)
                        mlngCorrect = mlngCorrect + 1
                        mlngNewFile = mlngNewFile + 1
                    End If
                End If
            End If
            lblInformation.Item(LN_可上传文件).Caption = Split(lblInformation.Item(LN_可上传文件).Caption, "：")(0) & "：" & str(mlngCorrect)
            lblInformation.Item(LN_检查警告文件).Caption = Split(lblInformation.Item(LN_检查警告文件).Caption, "：")(0) & "：" & str(mlngWarning)
            lblInformation.Item(LN_需上传新文件).Caption = Split(lblInformation.Item(LN_需上传新文件).Caption, "：")(0) & "：" & str(mlngNewFile)
'            lblInformation.Item(LN_无需更新文件).Caption = Split(lblInformation.Item(LN_无需更新文件).Caption, "：")(0) & "：" & str(mlngUnchanged)
        Next
        .Row = 1
        .ShowCell 1, Col_文件
        Call ShowStatus("检查完成", "", "", pgbThis.Max)
        imgCaption.Picture = imgList.ListImages("文件").Picture
        lblEXP(0).ForeColor = vbBlack
        Call ControlVisible(True)
        Call UpLoadFilesCount
    End With
End Function

'数据检查，并修正视图，标明部件状态
'状态值 0-正常 1-标准部件缺失(本地文件不存在) 2-本地文件不存在 3-无需更新 4-警告但可以上传 5-已经上传
Private Sub DataCheck()
    Dim objFile As New FileSystemObject
    Dim strFileName As String
    Dim strTemp As String
    Dim strStateContent As String
    Dim i As Long
    Dim lngAbnormal As Long '异常文件
    On Error GoTo errH
    '文件存在检查，标红提示
    With vsfMain
        If .Rows < .FixedRows Then Exit Sub
        For i = .FixedRows To .Rows - 1
            strTemp = .Cell(flexcpData, i, Col_警告)
            strStateContent = ""
            If strTemp <> "1" Then '为1说明标准部件缺失
                If Trim(.TextMatrix(i, Col_安装路径)) <> "" Then
                    '安装路径转换成实际路径
                    strFileName = mcllPath("K_" & UCase(.TextMatrix(i, Col_安装路径))) & "\" & UCase(.TextMatrix(i, Col_文件))
                    .TextMatrix(i, Col_文件地址) = UCase(Trim(strFileName))
                    '本地文件存不存在
                    If objFile.FileExists(strFileName) = False Then
                        strStateContent = "本地文件缺失"
                        If .TextMatrix(i, Col_文件类型) = "4" Then strStateContent = strStateContent & "(三方部件)"
                        Call FileStateSet(FS_状态异常, i, strStateContent)
                    Else
                        Call FileStateSet(FS_默认正常, i, strStateContent, "准备就绪", vbBlack)
                    End If
                Else
                    strStateContent = "安装路径缺失"
                    If .TextMatrix(i, Col_文件类型) = "4" Then strStateContent = strStateContent & "(三方部件)"
                    Call FileStateSet(FS_状态异常, i, strStateContent)
                End If
            Else
                strStateContent = "升级文件清单(zlFilesUpgrade)缺失该文件"
                If .TextMatrix(i, Col_文件类型) = "4" Then strStateContent = strStateContent & "(三方部件)"
                Call FileStateSet(FS_状态异常, i, strStateContent)
            End If
        Next
    End With
    Call UpLoadFilesCount
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Function CheckFTPServer(ByVal strIp As String, ByVal strUser As String, ByVal strPass As String, ByVal strPort As String) As Boolean
    '-----------------------------------------------------------------------------
    '功能:检查当前的FTP服务器是否正确
    '返回:当前的文件服务器的各项正确,返回true,否则返回False
    '编制:陈振原
    '日期:2016/07/05
    'strIp - FTP地址
    'strUser - 用户名
    'strPass - 密码
    'strPort - 端口
    '-----------------------------------------------------------------------------
    On Error GoTo errHand:
    
    If strIp = "" Or strUser = "" Or strPass = "" Or strPort = "" Then
        CheckFTPServer = False
        Exit Function
    End If
    
    If IsFtpServer(Trim(strIp), Trim(strUser), Trim(strPass), Trim(strPort)) Then
        CheckFTPServer = True
    Else
        CheckFTPServer = False
        MsgBox "不能连接升级服务器，请检查FTP服务器配置!", vbInformation + vbDefaultButton1, gstrSysName
    End If
    Exit Function
    
errHand:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function CheckShareServer(ByVal strAddress As String, ByVal strUser As String, ByVal strPass As String) As Boolean
    '-----------------------------------------------------------------------------
    '功能:检查当前的文件服务器是否正确
    '返回:当前的文件服务器的各项正确,返回true,否则返回False
    '编制:陈振原
    '日期:2016/07/05
    'strAddress - 地址
    'strUser - 用户
    'strPass - 密码
    '-----------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT

    On Error GoTo errHand:
    
    If strAddress = "" Or strUser = "" Or strPass = "" Then
        CheckShareServer = False
        Exit Function
    End If
    
    If FindFile(Trim(strAddress)) = False Then
        If IsNetServer(Trim(strAddress), Trim(strUser), Trim(strPass)) = False Then
            MsgBox "不能连接升级服务器，请检查共享服务器配置！", vbInformation + vbDefaultButton1, gstrSysName
            CheckShareServer = False
            Exit Function
        End If
    End If
    Call CancelNetServer(Trim(strAddress))
    CheckShareServer = True
    Exit Function
errHand:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function FindFile(ByVal strFileName As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '--功能:查找指定的文件或文夹是否存在
    '--返回: 如果存在此文件为True,否则为Flase
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT
    
    On Error Resume Next
    FindFile = False
    If Len(strFileName) > 0 Then
        apiOpenFile strFileName, typOfStruct, OF_EXIST
        FindFile = typOfStruct.nErrCode <> 2
    End If
End Function

'获取版本的直观显示值
Private Function GetFileVision(ByVal strVision As String) As String
    Dim lng版本号 As Variant
    Dim str版本号 As String
    If Len(strVision) > 0 Then
        lng版本号 = strVision
        str版本号 = Int(lng版本号 / 10 ^ 8)
        If Len(lng版本号) > 9 Then
            lng版本号 = Right(lng版本号, 9) Mod (10 ^ 8)
        Else
            lng版本号 = lng版本号 Mod (10 ^ 8)
        End If
        
        str版本号 = str版本号 & "." & Int(lng版本号 / 10 ^ 4)
        lng版本号 = lng版本号 Mod 10 ^ 4
        str版本号 = str版本号 & "." & lng版本号
        GetFileVision = str版本号
    End If
End Function

Private Sub UpLoadFilesCount()
    Dim i As Long
    Dim lngAbnormal As Long '异常
    Dim lngCorrect As Long '正常
    Dim lngUnchanged As Long '无需更新
    Dim lngWarning As Long  '警告
    Dim lngUpload As Long '已经上传
    Dim lngNewFile As Long '已经上传
    
    With vsfMain
        If .Rows < 1 Then Exit Sub
'        lblInformation.Item(Num_文件总数).Caption = str(.Rows - 1)
        lngAbnormal = 0: lngCorrect = 0: lngUnchanged = 0: lngWarning = 0: lngUpload = 0: lngNewFile = 0
        
        For i = 1 To .Rows - 1
            Select Case .Cell(flexcpData, i, Col_状态)
                Case FS_默认正常 '可以(需要)上传，正常
                    lngNewFile = lngNewFile + 1
                    lngCorrect = lngCorrect + 1
                Case FS_状态异常 '升级文件缺失，异常
                    lngAbnormal = lngAbnormal + 1
                Case FS_无需更新
                    lngUnchanged = lngUnchanged + 1
                    lngCorrect = lngCorrect + 1
                Case FS_无需更新警告文件 '无需更新，正常
                    lngUnchanged = lngUnchanged + 1
                    lngWarning = lngWarning + 1
                    lngCorrect = lngCorrect + 1
                Case FS_准备上传警告文件 '警告，可上传
                    lngNewFile = lngNewFile + 1
                    lngCorrect = lngCorrect + 1
                    lngWarning = lngWarning + 1
                Case FS_已经上传  '已经上传
                    lngUpload = lngUpload + 1
                    lngCorrect = lngCorrect + 1
            End Select
        Next
'        If mblnCheckMD5Flag = True Then
            lblInformation.Item(LN_所有上传文件).Caption = Split(lblInformation.Item(LN_所有上传文件).Caption, "：")(0) & "：" & str(.Rows - 1)
            lblInformation.Item(LN_状态异常文件).Caption = Split(lblInformation.Item(LN_状态异常文件).Caption, "：")(0) & "：" & str(lngAbnormal)
            lblInformation.Item(LN_已上传文件).Caption = Split(lblInformation.Item(LN_已上传文件).Caption, "：")(0) & "：" & str(lngUpload)
            lblInformation.Item(LN_可上传文件).Caption = Split(lblInformation.Item(LN_可上传文件).Caption, "：")(0) & "：" & str(lngCorrect)
            lblInformation.Item(LN_检查警告文件).Caption = Split(lblInformation.Item(LN_检查警告文件).Caption, "：")(0) & "：" & str(lngWarning)
            lblInformation.Item(LN_需上传新文件).Caption = Split(lblInformation.Item(LN_需上传新文件).Caption, "：")(0) & "：" & str(lngNewFile)
'            lblInformation.Item(LN_无需更新文件).Caption = Split(lblInformation.Item(LN_无需更新文件).Caption, "：")(0) & "：" & str(lngUnchanged)
'        Else
'            lblInformation.Item(Num_文件总数).Caption = str(.Rows - 1)
'            lblInformation.Item(Num_状态异常).Caption = str(lngAbnormal)
'            lblInformation.Item(Num_已经上传).Caption = "0"
'            lblInformation.Item(Num_需要上传).Caption = "未检测"
'            lblInformation.Item(Num_MD5警告).Caption = "未检测"
'            lblInformation.Item(Num_无需上传).Caption = "未检测"
'        End If
    End With

    If lngNewFile = 0 Then cmdUpload.Enabled = False
End Sub

Private Sub ControlVisible(blnVisible As Boolean) '界面按键可用性控制
'   cmdExit.Enabled = blnVisible
    cmdMD5Check.Enabled = blnVisible
    cmdUpload.Enabled = blnVisible
    cmdAllUpLoad.Enabled = blnVisible
    picInformation.Enabled = blnVisible
    mblnCheckMD5Tag = IIf(blnVisible = False, True, False)
    mblnUploadTag = IIf(blnVisible = False, True, False)
    mblnAllUploadTag = IIf(blnVisible = False, True, False)
End Sub
Private Sub Form_Unload(Cancel As Integer)
'    If mblnCheckMD5Tag = True Then Cancel = 1
    If mblnUploadTag = True Then Cancel = 1
    If mblnAllUploadTag = True Then Cancel = 1
End Sub

Private Sub lblInformation_Change(Index As Integer)
    Dim lngWidth As Long
    lngWidth = picInformation.Width / 6
    lblInformation.Item(Index).Move (Index * lngWidth) + ((lngWidth) - lblInformation.Item(Index).Width) / 2, (picInformation.ScaleHeight - lblInformation.Item(Index).Height) / 2
End Sub

Private Sub lblInformation_Click(Index As Integer)
'还原粗体状态
    Dim i As Integer
    Dim vPic As Variant
    
    For i = 0 To LN_lbl总数 - 1
        lblInformation.Item(i).FontBold = False
        lblInformation_Change i
    Next
    
    lblInformation.Item(Index).FontBold = True
    lblInformation_Change Index
'    picInformation.Picture = imgList.ListImages(Index + 1).Picture
    imgInformation.Picture = imgList.ListImages(Index + 1).Picture
    
    '异常文件提示
    For Each vPic In picState
        vPic.Visible = False
    Next
    Select Case Index
    Case LN_状态异常文件
        imgCaption.Picture = imgList.ListImages("异常").Picture
        picState(0).Visible = True
    Case LN_检查警告文件
        imgCaption.Picture = imgList.ListImages("警告").Picture
        picState(1).Visible = True
    Case Else
        imgCaption.Picture = imgList.ListImages("文件").Picture
    End Select

    Call DataFilter(Index)
End Sub

'数据过滤
Private Sub DataFilter(intFiler As Integer)
'状态值 0-正常 1-升级部件缺失(本地文件必不存在) 2-本地文件不存在 3-无需更新 4-警告但可以上传 5-已经上传
'strFiler 过滤值代表的状态
'"-1"-显示所有数据、"0"-可以上传、"1"-状态异常、"2"-状态异常、"3"-无需更新、"4"-警告且可上传 、"5"-已经上传

    Dim i As Long
    Dim strData As String
    
    With vsfMain
        If .Rows < 1 Then Exit Sub
        .Redraw = flexRDNone
        For i = 1 To .Rows - 1
            strData = Trim(.Cell(flexcpData, i, Col_状态))
            Select Case intFiler
            Case LN_所有上传文件
                If .RowHidden(i) = True Then .RowHidden(i) = False
            Case LN_可上传文件
                If strData <> FS_状态异常 Then
                    .RowHidden(i) = False
                Else
                    .RowHidden(i) = True
                End If
            Case LN_状态异常文件
                If strData = FS_状态异常 Then
                    .RowHidden(i) = False
                Else
                    .RowHidden(i) = True
                End If
            Case LN_检查警告文件
                If strData = FS_准备上传警告文件 Or strData = FS_无需更新警告文件 Then
                    .RowHidden(i) = False
                Else
                    .RowHidden(i) = True
                End If
            Case LN_已上传文件
                If strData = FS_已经上传 Then
                    .RowHidden(i) = False
                Else
                    .RowHidden(i) = True
                End If
            Case LN_需上传新文件
                If strData = FS_默认正常 Or strData = FS_准备上传警告文件 Then
                    .RowHidden(i) = False
                Else
                    .RowHidden(i) = True
                End If
'            Case LN_无需更新文件
'                If strData = FS_无需更新 Or strData = FS_无需更新警告文件 Then
'                    .RowHidden(i) = False
'                Else
'                    .RowHidden(i) = True
'                End If
            End Select
        Next
        '定位
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                .ShowCell i, Col_文件
                .Row = i
                Exit For
            End If
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Function CopyFileToShareServer(ByVal strNumber As String, ByVal strServerAddress As String, Optional ByVal strUser As String, Optional ByVal strPassword As String, Optional ByRef strErrInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------
    '功能:拷贝文件给指定的服务器
    '参数:strNumber-文件服务器编号
    '     strSourcePath-源文件目录(收集文件夹目录)
    '     strServerAddress-服务器的共享目录
    '     strUser-访问的用户名
    '     strPassword-密码
    '出参:strErrInfor-返回的错误信息
    '返回:拷贝成功,返回true,否则返回False
    '编制:陈振原
    '日期:2016/07/22
    '---------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim i As Long
    Dim strSQL As String
    Dim strTemp As String
    Dim strFilePath As String   '复制文件地址
    Dim strFileSource As String '被复制原文件地址
    Dim BlnState As Boolean
    '界面变量
    Dim blnUpLoadFail As Boolean '上传失败不会更新MD5
    
    '1.检查服务器是否连通
    If CheckShareServer(strServerAddress, strUser, strPassword) = False Then Exit Function '共享服务器连接校验
'    MsgBox strNumber & " 号服服务器 " & """" & strServerAddress & """" & " 测试连接成功", vbDefaultButton1 + vbInformation, gstrSysName
                    
    If objFile.FolderExists(mstrScratchFilePath) = False Then
        strErrInfor = "源文件目录:" & mstrScratchFilePath & "不存在,请检查!"
        Exit Function
    End If
    
    err = 0: On Error GoTo errHand:
    
    With vsfMain
        If .Rows < 1 Then MsgBox "文件列表为空，请检查！", vbDefaultButton1 + vbInformation, gstrSysName: Exit Function
        '界面初始化
        Call ShowStatus("", "", "", 0, .Rows - 1)
        For i = 1 To .Rows - 1
            ShowCenterRow i
            strTemp = UCase(Trim(.TextMatrix(i, Col_收集文件)))
            strFilePath = strServerAddress & "\" & strTemp
            strFileSource = .TextMatrix(i, Col_收集地址)
            Call ShowStatus("正在上传", strTemp & " 至 " & strNumber & " 号服务器 ", "", i)
            strTemp = .Cell(flexcpData, i, Col_状态)
            If mblnAllUpload = True Then
                BlnState = (strTemp <> FS_状态异常) '全部上传，只要非状态异常，均可重新上传
            Else
                BlnState = (strTemp = FS_默认正常 Or strTemp = FS_准备上传警告文件 Or strTemp = FS_已经上传) '更新上传状态为正常和警告的文件可以上传
            End If
            If BlnState Then  '状态为正常和警告的文件可以上传
                err = 0: On Error Resume Next
                .TextMatrix(i, Col_状态) = "正在上传"
'                Call FileStateSet(FS_默认正常, i, , "正在上传")
                DoEvents '防止界面卡顿
                objFile.CopyFile strFileSource, strFilePath, True
                If err <> 0 Then
                    If MsgBox("源文件：" & strFileSource & vbCrLf & " 不能拷贝到目标文件：" & vbCrLf & strFilePath & vbCrLf & "中,是否继续？" & vbNewLine & err.Description, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    Call FileStateSet(FS_上传失败, i, "请检查文件，稍后重新上传")
                    blnUpLoadFail = True '上传失败不能更新MD5
                Else
                    Call FileStateSet(FS_已经上传, i, strNumber & " 号服务器已上传")
                End If
            End If
        Next
        .Row = 1
        .ShowCell 1, Col_文件
        Call ShowStatus("", "", strNumber & " 号服务器上传完成", pgbThis.Max)
    End With
    
    If blnUpLoadFail = False Then
        mblnUpLoadSuccess = True
        mstrSuccessUploadSever = mstrSuccessUploadSever & strNumber & "号," '更新MD5
        CopyFileToShareServer = True
    Else
        CopyFileToShareServer = False
    End If
    Exit Function
errHand:
    strErrInfor = "共享上传过程出现错误:" & vbCrLf & "错误号:" & err.Number & vbCrLf & "错误描述:" & err.Description
    If False Then
        Resume
    End If
End Function

Private Function CopyFileToFTPServer(ByVal strNumber As String, ByVal strServerAddress As String, Optional ByVal strUser As String, Optional ByVal strPassword As String, Optional ByVal strPort As String, Optional ByRef strErrInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------
    '功能:拷贝文件给指定的服务器
    '参数:strNumber-文件服务器编号
    '     strSourcePath-源文件目录
    '     strServerAddress-服务器的共享目录
    '     strUser-访问的用户名
    '     strPassword-密码
    '     strPort-端口
    '出参:strErrInfor-返回的错误信息
    '返回:拷贝成功,返回true,否则返回False
    '编制:陈振原
    '日期:2016/07/22
    '---------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim strTemp As String
    Dim strFileName As String   '复制文件名称
    Dim strFileSource As String '被复制文件源地址
    Dim BlnState As Boolean
    
    Dim i As Long
    Dim strSQL As String
    Dim blnUpLoadFail As Boolean
    
    If CheckFTPServer(strServerAddress, strUser, strPassword, strPort) = False Then Exit Function 'FTP服务器连接校验
'    MsgBox strNumber & " 号服服务器 " & """" & strServerAddress & """" & " 测试连接成功", vbDefaultButton1 + vbInformation, gstrSysName

    err = 0: On Error GoTo errHand:
    
    With vsfMain
        If .Rows < 1 Then MsgBox "文件列表为空，请检查！", vbDefaultButton1 + vbInformation, gstrSysName: Exit Function
        
        Call ShowStatus("", "", "开始上传", 0, .Rows - 1)
        For i = 1 To .Rows - 1
            ShowCenterRow i
            '数据初始化
            strFileName = .TextMatrix(i, Col_收集文件)
            strFileSource = .TextMatrix(i, Col_收集地址)
            Call ShowStatus("正在上传", strFileName & " 至 " & strNumber & " 号服务器 ", "", i)
            strTemp = .Cell(flexcpData, i, Col_状态)
            
            If mblnAllUpload = True Then
                BlnState = (strTemp <> FS_状态异常) '全部上传，只要非状态异常，均可重新上传
            Else
                BlnState = (strTemp = FS_默认正常 Or strTemp = FS_准备上传警告文件 Or strTemp = FS_已经上传) '更新上传状态为正常和警告的文件可以上传
            End If
            
            If BlnState Then
                err = 0: On Error Resume Next
                '文件拷贝,需先判断

'                If UCase(Nvl(strFileName, "")) <> UCase("zlHisCrust.exe") And UCase(Nvl(strFileName, "")) <> UCase("7z.exe") And UCase(Nvl(strFileName, "")) <> UCase("7z.dll") And UCase(Nvl(strFileName, "")) <> UCase("aamd532.dll") And UCase(Nvl(strFileName, "")) <> UCase("zlRunas.exe") And UCase(Nvl(strFileName, "")) <> UCase("RegCom.dll") Then
'                    strFileName = strFileName & ".7Z"
'                End If
                .TextMatrix(i, Col_状态) = "正在上传"
                DoEvents '防止界面卡顿
                If FtpupFile(strFileSource, strFileName) = False Then
'                    If MsgBox("源文件：" & strFileSource & vbCrLf & " 不能拷贝到目标文件：" & vbCrLf & strFileName & vbCrLf & "中,是否继续？" & vbNewLine & err.Description, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    Call FileStateSet(FS_上传失败, i, "请检查文件，稍后重新上传")
                    blnUpLoadFail = True '上传失败不能更新MD5
                Else
                    Call FileStateSet(FS_已经上传, i, strNumber & " 号服务器已上传")
                End If
            End If
        Next
        .Row = 1
        .ShowCell 1, Col_文件
        Call ShowStatus("", "", strNumber & " 号服务器上传完成", pgbThis.Max)
    End With
    
    If blnUpLoadFail = False Then
        mblnUpLoadSuccess = True
        mstrSuccessUploadSever = mstrSuccessUploadSever & strNumber & "号," '更新MD5
        CopyFileToFTPServer = True
    Else
        CopyFileToFTPServer = False
    End If
    Exit Function
errHand:
    strErrInfor = "FTP上传过程出错:" & vbCrLf & "错误号:" & err.Number & vbCrLf & "错误描述:" & err.Description
    If False Then
        Resume
    End If
End Function

Private Function UpdateMD5() As Boolean
    If mblnUpLoadSuccess = False Then UpdateMD5 = False: Exit Function
    Dim i As Long
    Dim lngPercent As Long
    Dim intPercent As Integer
    Dim strSQL As String
    Dim BlnState As String
    mstrSuccessUploadSever = Mid(mstrSuccessUploadSever, 1, Len(mstrSuccessUploadSever) - 1) & "服务器上传完成"
    With vsfMain
        mlngUpload = 0
        Call ShowStatus("", "", "准备更新", i, .Rows - 1)
        For i = 1 To .Rows - 1
            ShowCenterRow i
            If mblnAllUpload = True Then
                BlnState = (.Cell(flexcpData, i, Col_状态) <> FS_状态异常) '全部上传，只要非状态异常，均可重新上传
            Else
                BlnState = (.Cell(flexcpData, i, Col_状态) = FS_默认正常 Or .Cell(flexcpData, i, Col_状态) = FS_准备上传警告文件 Or .Cell(flexcpData, i, Col_状态) = FS_已经上传) '更新上传状态为正常和警告的文件可以上传
            End If
            If BlnState Then
                DoEvents
                Call ShowStatus("正在更新", .TextMatrix(i, Col_文件) & " 的 MD5：", "", i)
                strSQL = "update zlfilesupgrade set MD5 = '" & .TextMatrix(i, Col_本地md5) & "' where  upper(文件名) = '" & UCase(.TextMatrix(i, Col_文件)) & "'"
                gcnOracle.Execute strSQL
                Call FileStateSet(FS_已经上传, i, mstrSuccessUploadSever, "上传完成")
                mlngUpload = mlngUpload + 1
            Else
                Call ShowStatus("正在跳过", .TextMatrix(i, Col_文件), "", i)
            End If
            lblInformation.Item(LN_已上传文件).Caption = Split(lblInformation.Item(LN_已上传文件).Caption, "：")(0) & "：" & str(mlngUpload)
        Next
        
        .Row = 1
        .ShowCell 1, Col_文件
        Call ShowStatus("", "", mstrSuccessUploadSever, pgbThis.Max, , False)
        mstrSuccessUploadSever = ""
        mblnUpLoadSuccess = False
        UpdateMD5 = True
    End With
End Function

Private Function BatchLoad() As String
'读取上传批次
    Dim strSQL As String
    Dim strTemp As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = "select 内容 as 批次 from ZLReginfo where 项目 = '最新升级文件批次'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    
    If rsTemp.EOF Then
        strSQL = "insert into zltools.ZLReginfo(项目,内容) select '最新升级文件批次','0' from dual where not Exists (select 1 from zltools.ZLReginfo where 项目 ='最新升级文件批次')"
        gcnOracle.Execute strSQL
        BatchLoad = "0"
    Else
        BatchLoad = rsTemp.Fields("批次")
    End If
    Exit Function
errH:
    MsgBox "批次读取错误"
End Function

Private Function BatchUpdate(strBath As String) As Boolean
'更新上传批次
    Dim strSQL As String
    Dim i As Long
    
    On Error GoTo errH
    
    strSQL = "update ZLReginfo set 内容 = '" & Trim(strBath) & "' where 项目 = '最新升级文件批次'"
    gcnOracle.Execute strSQL
    
    With vsfSever
        If .Rows < 1 Then BatchUpdate = False: Exit Function
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, Col_上传状态)) = Trim(str(Res_上传成功)) Then
                strSQL = "update ZLUpgradeServer set 批次 = " & strBath & " where 编号 = " & Trim(.TextMatrix(i, Col_编号))
                gcnOracle.Execute strSQL
            End If
        Next
    End With

    BatchUpdate = True
    
    Exit Function
errH:
    BatchUpdate = False
    MsgBox "批次更新错误"
End Function


Private Function funCanWrite(strWritePath As String) As Boolean
'判断远程目录是否具有写权限
    Dim strDest     As String
    Dim objFile As New FileSystemObject
    On Error GoTo errH
            strDest = strWritePath & "\tmp.txt"
            objFile.CreateTextFile strDest
            objFile.DeleteFile strDest, True
            funCanWrite = True
    Exit Function
errH:
    funCanWrite = False
End Function

'Public Function ISCopyFile(ByVal strSourceFile As String, ByVal strTarGetFile As String) As Boolean
'     '---------------------------------------------------------------------------------------------------------------
'    '
'    '功能:判断是否需要拷贝文件(比较版本号,修改时间)
'    '入参数:
'    '   strSourceFile:源文件
'    '   strTargetFile:目标文件
'    '返回:需要拷贝则返回true,否则返回false
'    '---------------------------------------------------------------------------------------------------------------
'    Dim strSource As String, strTarget As String
'
'    ISCopyFile = False
'    err = 0: On Error Resume Next
'    If FindFile(strTarGetFile) = False Then
'        '没有发现文件，则返回true
'        ISCopyFile = True
'        Exit Function
'    End If
'
'    '比较文件版本号
'    strTarget = GetCommpentVersion(strTarGetFile)
'    strSource = GetCommpentVersion(strSourceFile)
'    If RtnVerNum(strTarget) < RtnVerNum(strSource) Then
'        ISCopyFile = True
'        Exit Function
'    End If
'
'    '比较文件的最后修改时间
'    strTarget = Format(FileDateTime(strTarGetFile), "yyyy-MM-DD hh:mm:ss")
'    strSource = Format(FileDateTime(strSourceFile), "yyyy-MM-DD hh:mm:ss")
'    If strTarget < strSource Then
'        ISCopyFile = True
'        Exit Function
'    End If
'End Function

Private Function FilesCollections() As Boolean
    Dim i As Long
    Dim lngPercent As Long '进度条百分比
    Dim intPercent As Integer '文字显示百分比
    Dim blnCollect As Boolean '收集状态 0-不收集 1-收集
    Dim BlnState As Boolean
    Dim strTemp As String
    
    '压缩相关
    Dim strCurFileDirectory As String '目标文件夹
    Dim strCompTxt  As String '压缩脚本
    Dim strSourcePath   As String '压缩源文件路径
    Dim strDescPath     As String '压缩目标文件路径

    Dim objFile As New FileSystemObject
    Dim str7zFile   As String
    Dim driver As Drive
        
    '数据库文件的值
    Dim strFileName As String
    Dim strFilePath As String
    Dim strFileMD5 As String '本地文件MD5值
        
    '检测部件是否用收集
    Dim strEditDate As String '本地文件修改时间
    Dim strEditDateNow As String '当前文件修改时间
    Dim strVersion As String '本地版本号
    Dim strVersionNow As String '当前版本号

    err = 0: On Error GoTo errHand:
    strCurFileDirectory = Trim(mstrScratchFilePath)
    FilesCollections = False
        
    '检查剩余空间
    For Each driver In objFile.Drives
        If driver.IsReady Then
            If driver.DriveLetter = "C" Then
                If driver.FreeSpace < 204800000 Then '小于200M
                    MsgBox "临时收集目录没有足够的空间!", vbInformation, gstrSysName
                    Exit Function
                End If
                Exit For
            End If
        End If
    Next driver

    If FindFile(strCurFileDirectory) = False Then
        On Error Resume Next
        Call mobjFile.CreateFolder(strCurFileDirectory)
        If mobjFile.FolderExists(strCurFileDirectory) = False Then
            MsgBox "临时收集目录不能创建,请检查!" & vbCrLf & strCurFileDirectory, vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If

    '检查7z路径
    If Init7Z = False Then Exit Function
       
    '先清除临时收集文件目录中的所有内容
    err = 0: On Error Resume Next
    
'    If MsgBox("上传前需要先收集本地文件，是否要清空收集文件夹，全部文件重新收集？", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
'        objFile.DeleteFolder strCurFileDirectory & "\*", True
'        objFile.DeleteFile strCurFileDirectory & "\*.*", True
'    End If
    
'    MsgBox "收集过程中会有卡顿，请耐心等待！", vbDefaultButton1 + vbInformation, gstrSysName
    imgCaption.Picture = imgList.ListImages("收集文件").Picture
    lblEXP(1).ForeColor = vbBlue
    With vsfMain
        If .Rows < 1 Then MsgBox "文件列表为空，请检查", vbDefaultButton1 + vbQuestion, gstrSysName: Exit Function

        Call ShowStatus("", "", "准备收集", 0, .Rows - 1, True)
        For i = 1 To .Rows - 1
            ShowCenterRow i
            DoEvents
            '更新界面状态，进度条文字显示等
            Call ShowStatus("正在收集", strFileName & " 至 " & strCurFileDirectory, "", i)
            '初始化当前行数据、名称、地址、本版、修改日期
            strFileName = UCase(.TextMatrix(i, Col_文件))
            strFilePath = .TextMatrix(i, Col_文件地址)
            strFileMD5 = .TextMatrix(i, Col_本地md5)
            strVersion = GetDealVersion(strFilePath)  '本地文件版本
            strVersionNow = .TextMatrix(i, Col_当前版本)  '当前文件版本
            strEditDate = Format(FileDateTime(strFilePath), "yyyy-MM-DD hh:mm:ss") '本地文件修改时间
            strEditDateNow = Format(.TextMatrix(i, Col_修改日期), "yyyy-MM-DD hh:mm:ss") '当前文件修改时间
            blnCollect = False
            If mblnAllUpload = True Then
                BlnState = (.Cell(flexcpData, i, Col_状态) <> FS_状态异常)
            Else
                BlnState = (.Cell(flexcpData, i, Col_状态) = FS_默认正常 Or .Cell(flexcpData, i, Col_状态) = FS_准备上传警告文件)
            End If
            If BlnState Then
                '状态不为异常的可以收集
                '7z进行压缩，5个文件不需要压缩 特殊处理
                If InStr(";ZLHISCRUST.EXE;7Z.EXE;7Z.DLL;AAMD532.DLL;ZLRUNAS.EXE;REGCOM.DLL;GACUTIL.EXE;GACUTIL.EXE.CONFIG;", ";" & UCase(Nvl(strFileName, "")) & ";") > 0 Then
                    strDescPath = strCurFileDirectory & "\" & UCase(strFileName)
                    '文件存在 且版本号相同 且修改日期相同 不需要重新压缩
                    If objFile.FileExists(strDescPath) = True And strVersion = strVersionNow And CompareDate(strEditDate, strEditDateNow) = 0 Then
                        blnCollect = False
                    Else
                        blnCollect = True
                    End If
                    
                    If blnCollect = True Then
                        DoEvents '防止界面卡顿
                        Call objFile.CopyFile(strFilePath, strDescPath, True)
                        Call SaveCollectFilesInformation(i)

                        Call FileStateSet(FS_默认正常, i, "收集完成")
                    Else
                        Call FileStateSet(FS_默认正常, i, "无需收集")
                    End If
                    '存储压缩文件地址及压缩后文件名称
                    .TextMatrix(i, Col_收集地址) = strDescPath
                    .TextMatrix(i, Col_收集文件) = strFileName
                Else
                    strDescPath = strCurFileDirectory & "\" & GetCompressName(Nvl(strFileName, ""))
                    If objFile.FileExists(strDescPath) = True And strVersion = strVersionNow And CompareDate(strEditDate, strEditDateNow) = 0 Then
                        blnCollect = False
                    Else
                        blnCollect = True
                    End If
                    
                    If blnCollect = True Then
                        strCompTxt = CompressionCmd(strDescPath, strFilePath, COMPRESSIONRATE)
                        If strCompTxt <> "" Then
                            DoEvents '防止界面卡顿
                            Call GetCmdTxt(strCompTxt)
                            Call SaveCollectFilesInformation(i)
                            
                            Call FileStateSet(FS_默认正常, i, "收集完成")
                        Else
                            Call FileStateSet(FS_状态异常, i, "文件收集失败,请重试")
                        End If
                    Else
                        .TextMatrix(i, Col_警告) = "无需收集"
                    End If
                    '存储压缩文件地址及压缩后文件名称
                    .TextMatrix(i, Col_收集地址) = strDescPath
                    .TextMatrix(i, Col_收集文件) = GetCompressName(Nvl(strFileName, ""))
                End If
            End If
        Next
        Call ShowStatus("", "", "文件收集完成", pgbThis.Max)
        .Row = 1
        .ShowCell 1, Col_文件
    End With
    imgCaption.Picture = imgList.ListImages("文件").Picture
    lblEXP(1).ForeColor = vbBlack
    FilesCollections = True
    Exit Function
errHand:
    MsgBox "压缩收集过程出现错误:" & vbCrLf & err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Function

Private Function GetCompressName(ByVal strFileName As String) As String
'功能转换为7z后缀的压缩格式名称
    On Error GoTo errH
    GetCompressName = strFileName & ".7z"
    Exit Function
errH:
    If err Then
         MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function SaveCollectFilesInformation(lngRow As Long) As Boolean
'更新上传部件文件信息
'lngRow - 表格内行数，每一行代表一个部件
    Dim strFilesPath As String '文件路径
    Dim strFileName As String '文件名
    Dim strMD5      As String 'MD5
    Dim strEditDate As String '修改日期
    Dim strVision   As String   '版本号
    Dim strSQL As String
    On Error GoTo errH
    
    With vsfMain
        strFilesPath = .TextMatrix(lngRow, Col_文件地址)
        strFileName = .TextMatrix(lngRow, Col_文件)
'        strMD5 = .TextMatrix(lngRow, Col_本地md5)
        strEditDate = Format(FileDateTime(strFilesPath), "yyyy-MM-DD hh:mm:ss")
        strVision = GetDealVersion(strFilesPath)
'        strVision = GetCommpentVersion(strFilesPath)
'        strVision = GetTransVersion(strVision)
        
        If strFileName <> "" Then
            If InStr(";ZLHISCRUST.EXE;7Z.EXE;7Z.DLL;AAMD532.DLL;ZLRUNAS.EXE;REGCOM.DLL;GACUTIL.EXE;GACUTIL.EXE.CONFIG;", ";" & UCase(Nvl(strFileName, "")) & ";") > 0 Then
                strSQL = "update zlfilesupgrade set 版本号='1000350040' ,文件版本号='" & strVision & "',修改日期='" & strEditDate & "' where upper(文件名)='" & UCase(strFileName) & "'"
            Else
                strSQL = "update zlfilesupgrade set 文件版本号='" & strVision & "',修改日期='" & strEditDate & "' where upper(文件名)='" & UCase(strFileName) & "'"
            End If
            gcnOracle.Execute strSQL
            SaveCollectFilesInformation = True
        End If
    End With
    Exit Function
errH:
    If err Then
         MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function GetTransVersion(ByVal strVersion As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:获得转换后的版本号
    '入参:strVersion
    '出参:strVersion 转换后无"."的版本号
    '返回:成功,返回版本号,否则返回空
    '编制:陈振原
    '日期:2016-08-03 16:59:34
    '-----------------------------------------------------------------------------------------------------------
    Dim lngVision   As Double '版本号
    Dim strTmp    As Variant
    
        If strVersion <> "" Then
            strTmp = Split(strVersion, ".")
            lngVision = strTmp(0) * 10 ^ 8 + strTmp(1) * 10 ^ 4 + strTmp(2)
            strVersion = lngVision
        End If
    GetTransVersion = strVersion
End Function

Private Function GetCommpentVersion(ByVal strFile As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取指定控件的版本号
    '入参:
    '出参:
    '返回:成功,返回版本号,否则返回空
    '编制:刘兴洪
    '日期:2009-01-16 16:59:34
    '-----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim strVer As String, varVersion As Variant
    
    err = 0: On Error Resume Next
    '获取文件版本号
    strVer = objFile.GetFileVersion(strFile)
    If err <> 0 Then
        err.Clear: err = 0
        GetCommpentVersion = ""
        Exit Function
    End If
'    If Trim(strVer) <> "" Then
'        varVersion = Split(strVer, ".")
'        If UBound(varVersion) > 2 Then
'            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(3)
'        ElseIf UBound(varVersion) = 2 Then
'            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(2)
'        End If
'    End If
    GetCommpentVersion = strVer
End Function

Private Function CompareDate(Data1 As String, Data2 As String) As Integer
'日期比较 格式必须为 "yyyy-MM-DD hh:mm:ss"
'Data1>Data2 返回 1
'Data1<Data2 返回 -1
'Data1=Data2 返回 0
'错误 返回 3
    If Data1 = "" Or Data2 = "" Then CompareDate = 3: Exit Function
    
    If DateDiff("s", Data1, Data2) < 0 Then
        CompareDate = 1 'data1新
    ElseIf DateDiff("s", Data1, Data2) > 0 Then
        CompareDate = 2 'data2新
    Else
        If Trim(Data1) = Trim(Data2) Then
            CompareDate = 0 '相等
        Else
            CompareDate = 3 '错误
        End If
    End If
End Function

Private Sub tbcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Me.Refresh
    Select Case Item.Index
        Case 0
        Case 1
        Case 2
        Case 3
        Case 4
        Case 5
        Case 6
    End Select
End Sub

Private Sub ShowStatus(strOperation As String, strContent As String, strCondition As String, lngPgbValue As Long, Optional lngPgbMax As Long = 0, Optional blnShowPgbbar As Boolean = True)
    '首次设置了进度条max后只用传入进度条Value即可
    'strCondition不为空，将不会显示strOperation和strContent内容
    '传入进度条Value为-1时，进度百分比不会显示
    'blnShowPgbbar控制隐藏和显示进度条 true-隐藏 false-不隐藏
    On Error Resume Next
    Dim intPercent As Integer
    If lngPgbMax < 0 Then Exit Sub

    If lngPgbValue = -1 Then
        lblstatus(Sta_百分比).Visible = False
    Else
        lblstatus(Sta_百分比).Visible = True
    End If
    
    If lngPgbMax = 0 Then
        If pgbThis.Max <> 0 Then
            intPercent = lngPgbValue / pgbThis.Max * 100
        Else
            Exit Sub
        End If
    Else
        pgbThis.Max = lngPgbMax
        intPercent = lngPgbValue / pgbThis.Max * 100
    End If
    
    If strCondition <> "" Then
        lblstatus(Sta_状态描述).Caption = strCondition
        lblstatus(Sta_百分比).Caption = ""
        lblstatus(Sta_当前操作).Caption = ""
        lblstatus(Sta_操作对象).Caption = ""
    Else
        lblstatus(Sta_状态描述).Caption = ""
        lblstatus(Sta_百分比).Caption = intPercent & "%"
        lblstatus(Sta_当前操作).Caption = strOperation & ""
        lblstatus(Sta_操作对象).Caption = strContent & ""
        pgbThis.value = lngPgbValue
    End If
    
    If blnShowPgbbar Then
        pgbThis.Visible = True
        vsfMain.Move 100, 1000 + picHelp.Height, Me.ScaleWidth - 185, Me.ScaleHeight - 1490 - picHelp.Height
    Else
        pgbThis.Visible = False
        vsfMain.Move 100, pgbThis.Top, Me.ScaleWidth - 185, Me.ScaleHeight - 1800 - pgbThis.Height
    End If
End Sub


Private Sub FileStateSet(emuFileState As FilesState, lngRow As Long, Optional strStateContent As String = "NULL", Optional strStateDisplay As String = "", Optional lngStateColor As Long = -1)
'    默认状态颜色设置
    Select Case emuFileState
        Case FS_默认正常
            If lngStateColor = -1 Then lngStateColor = SC_绿色
            If strStateDisplay = "" Then strStateDisplay = SDP_准备上传
        Case FS_状态异常
            If lngStateColor = -1 Then lngStateColor = SC_红色
            If strStateDisplay = "" Then strStateDisplay = SDP_状态异常
        Case FS_无需更新
            If lngStateColor = -1 Then lngStateColor = SC_蓝色
            If strStateDisplay = "" Then strStateDisplay = SDP_无需更新
        Case FS_无需更新警告文件
            If lngStateColor = -1 Then lngStateColor = SC_蓝色
            If strStateDisplay = "" Then strStateDisplay = SDP_无需更新
        Case FS_准备上传警告文件
            If lngStateColor = -1 Then lngStateColor = SC_黄色
            If strStateDisplay = "" Then strStateDisplay = SDP_准备上传
        Case FS_已经上传
            If lngStateColor = -1 Then lngStateColor = SC_蓝色
            If strStateDisplay = "" Then strStateDisplay = SDP_已经上传
        Case FS_上传失败
            If lngStateColor = -1 Then lngStateColor = SC_红色
            If strStateDisplay = "" Then strStateDisplay = SDP_上传失败
        Case Else
            If lngStateColor = -1 Then lngStateColor = vbBlack
            If strStateDisplay = "" Then strStateDisplay = "TEST"
    End Select
    
    With vsfMain
        .Cell(flexcpData, lngRow, Col_状态) = emuFileState
        .TextMatrix(lngRow, Col_状态) = strStateDisplay
        If strStateContent <> "NULL" Then
            .TextMatrix(lngRow, Col_警告) = strStateContent
        End If
        If FS_状态异常 = emuFileState Then
            .Cell(flexcpForeColor, lngRow, 0, lngRow, Col_列数 - 1) = lngStateColor
        Else
            .Cell(flexcpForeColor, lngRow, Col_状态, lngRow, Col_警告) = lngStateColor
        End If
    End With
End Sub

Private Sub ShowCenterRow(lngRow As Long)

    Dim i As Integer
    Dim intLocation As Integer
    Dim lngShowRow As Integer
    With vsfMain
        If .RowHidden(lngRow) = True Then Exit Sub
        intLocation = 14
        lngShowRow = lngRow
        i = 0
        Do Until i >= intLocation
            If .Rows - (lngShowRow + intLocation) <= 0 Then
                .Row = lngRow
                .ShowCell lngRow, Col_文件
                Exit Sub
            End If
            lngShowRow = lngShowRow + 1
            If .RowHidden(lngShowRow) = False Then i = i + 1
        Loop
        .Row = lngRow
        .ShowCell lngShowRow, Col_文件
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub AllUploadRestore()
    Dim i As Long
    With vsfMain
        If .Rows < .FixedRows Then Exit Sub
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, Col_状态) <> FS_状态异常 Then
                Select Case .Cell(flexcpData, i, Col_状态)
                Case FS_默认正常
                    Call FileStateSet(FS_默认正常, i, "")
                Case FS_无需更新
                    Call FileStateSet(FS_默认正常, i, "")
                Case FS_无需更新警告文件
                    Call FileStateSet(FS_准备上传警告文件, i)
                Case FS_准备上传警告文件
                    Call FileStateSet(FS_准备上传警告文件, i)
                Case Else
                    Call FileStateSet(FS_默认正常, i, "")
                End Select
            End If
        Next
        .Redraw = flexRDBuffered
    End With
    Call UpLoadFilesCount
End Sub
'Private Sub FloderToClipBoard(ByVal strSourceFloder As String)
'    '拷贝临时收集文件目录的文件到剪贴板中去
'    Dim strFile() As String
'    Dim strSourceFile As String
'    Dim strTemp As String
'    Dim i As Integer
'    strSourceFile = strSourceFloder & "\"
'    Erase strFile
'
'
'    If mobjFile.FolderExists(strSourceFile) Then
'        With FileList
'            .Refresh
'            .Path = strSourceFile
'            .FileName = "*.*"
'
'            For i = 0 To .ListCount - 1
'                ReDim Preserve strFile(i)
'                strTemp = strSourceFile & .List(i)
'                strFile(i) = strTemp
'            Next
'
'            If .ListCount <> 0 Then
'                Call clipCopyFiles(strFile)
'            End If
'        End With
'    End If
'End Sub
