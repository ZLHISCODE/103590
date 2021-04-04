VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPacsInterfaceCfg 
   Caption         =   "插件配置管理"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16020
   Icon            =   "frmPacsInterfaceCfg.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10155
   ScaleWidth      =   16020
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picAppCfg 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8895
      Left            =   5760
      ScaleHeight     =   8865
      ScaleWidth      =   10020
      TabIndex        =   4
      Top             =   360
      Width           =   10050
      Begin VB.Frame fraAppFuns 
         BorderStyle     =   0  'None
         Caption         =   "功能配置"
         Height          =   7380
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   9915
         Begin VB.Frame fraVBS 
            Caption         =   "VBS脚本"
            Height          =   3315
            Left            =   5880
            TabIndex        =   23
            Top             =   3840
            Width           =   3555
            Begin VB.CheckBox chkModify 
               Caption         =   "手动调整"
               Height          =   255
               Left            =   960
               TabIndex        =   28
               Top             =   0
               Width           =   1095
            End
            Begin VB.TextBox txtVBS 
               Height          =   2055
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   26
               Top             =   360
               Width           =   2655
            End
         End
         Begin VB.Frame fraFuncs 
            Caption         =   "功能列表"
            Height          =   3495
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   9135
            Begin VB.CommandButton cmdTestFunc 
               Caption         =   "功能验证"
               Height          =   375
               Left            =   3000
               TabIndex        =   32
               Top             =   3000
               Width           =   1215
            End
            Begin VB.CommandButton cmdDelFun 
               Caption         =   "删除方法"
               Height          =   375
               Left            =   7080
               TabIndex        =   31
               Top             =   2880
               Width           =   1215
            End
            Begin VB.CommandButton cmdAddFunc 
               Caption         =   "添加功能"
               Height          =   375
               Left            =   120
               TabIndex        =   21
               Top             =   3000
               Width           =   1215
            End
            Begin VB.CommandButton cmdDelFunc 
               Caption         =   "删除功能"
               Height          =   375
               Left            =   1560
               TabIndex        =   20
               Top             =   3000
               Width           =   1215
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfAppFuns 
               Height          =   2580
               Left            =   180
               TabIndex        =   22
               Top             =   300
               Width           =   8415
               _cx             =   14843
               _cy             =   4551
               Appearance      =   1
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
               BackColorSel    =   16761024
               ForeColorSel    =   0
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   2
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   360
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
               Editable        =   2
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
               Begin VSFlex8Ctl.VSFlexGrid vsfFuncs 
                  Height          =   2445
                  Left            =   4320
                  TabIndex        =   30
                  Top             =   0
                  Width           =   915
                  _cx             =   1614
                  _cy             =   4313
                  Appearance      =   2
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
                  BackColorSel    =   16761024
                  ForeColorSel    =   0
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483633
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483642
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   0   'False
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   1
                  GridLines       =   2
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   10
                  Cols            =   3
                  FixedRows       =   0
                  FixedCols       =   0
                  RowHeightMin    =   360
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
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
                  Editable        =   2
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
            End
         End
         Begin VB.Frame fraFuncParas 
            Caption         =   "参数列表"
            Height          =   3495
            Left            =   240
            TabIndex        =   15
            Top             =   3840
            Width           =   4215
            Begin VB.CommandButton cmdDelPara 
               Caption         =   "删除参数"
               Height          =   375
               Left            =   1680
               TabIndex        =   17
               Top             =   3000
               Width           =   1215
            End
            Begin VB.CommandButton cmdAddPara 
               Caption         =   "添加参数"
               Height          =   375
               Left            =   120
               TabIndex        =   16
               Top             =   3000
               Width           =   1215
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfAppfunPara 
               Height          =   2460
               Left            =   180
               TabIndex        =   18
               Top             =   480
               Width           =   3060
               _cx             =   5397
               _cy             =   4339
               Appearance      =   1
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
               BackColorSel    =   16761024
               ForeColorSel    =   0
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   2
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   360
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
               Editable        =   2
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
               Begin VB.CommandButton cmdConfigWindow 
                  Caption         =   "配置"
                  Height          =   375
                  Left            =   2280
                  TabIndex        =   27
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   735
               End
            End
         End
      End
      Begin VB.Frame fraAppInfo 
         Caption         =   "基本信息"
         Height          =   1150
         Left            =   180
         TabIndex        =   5
         Top             =   120
         Width           =   8955
         Begin VB.TextBox txtAppName 
            Height          =   315
            Left            =   1140
            TabIndex        =   24
            Top             =   300
            Width           =   2895
         End
         Begin VB.ComboBox cboType 
            Height          =   300
            ItemData        =   "frmPacsInterfaceCfg.frx":6852
            Left            =   1140
            List            =   "frmPacsInterfaceCfg.frx":685F
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   720
            Width           =   2895
         End
         Begin VB.ComboBox cboClasses 
            Height          =   300
            Left            =   5220
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   720
            Width           =   2355
         End
         Begin VB.CheckBox chkUseThisApp 
            Caption         =   "启用"
            Height          =   255
            Left            =   7800
            TabIndex        =   8
            Top             =   740
            Width           =   675
         End
         Begin VB.TextBox txtAppDir 
            Height          =   350
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   7
            Tag             =   "VBS动态脚本"
            Top             =   285
            Width           =   3390
         End
         Begin VB.CommandButton cmdSelectApp 
            Caption         =   "…"
            Height          =   350
            Left            =   8595
            TabIndex        =   6
            Top             =   270
            Width           =   260
         End
         Begin VB.Label lblAppName 
            AutoSize        =   -1  'True
            Caption         =   "插件名称："
            Height          =   180
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            Caption         =   "执行类型："
            Height          =   180
            Left            =   240
            TabIndex        =   13
            Top             =   760
            Width           =   900
         End
         Begin VB.Label lblAppDir 
            AutoSize        =   -1  'True
            Caption         =   "程序路径："
            Height          =   180
            Left            =   4140
            TabIndex        =   12
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblClasses 
            AutoSize        =   -1  'True
            Caption         =   "程序集合："
            Height          =   180
            Left            =   4140
            TabIndex        =   11
            Top             =   780
            Width           =   900
         End
      End
   End
   Begin VB.PictureBox picApp 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7755
      Left            =   360
      ScaleHeight     =   7725
      ScaleWidth      =   4545
      TabIndex        =   0
      Top             =   1680
      Width           =   4575
      Begin VB.ComboBox cboStation 
         Height          =   300
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1260
         Width           =   2115
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfApp 
         Height          =   3840
         Left            =   960
         TabIndex        =   2
         Top             =   2280
         Width           =   3540
         _cx             =   6244
         _cy             =   6773
         Appearance      =   1
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
         BackColorSel    =   16761024
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   360
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
      Begin VB.Label lblStation 
         AutoSize        =   -1  'True
         Caption         =   "所属站点"
         Height          =   180
         Left            =   780
         TabIndex        =   3
         Top             =   1320
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   780
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   9795
      Width           =   16020
      _ExtentX        =   28258
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   4154
            MinWidth        =   4154
            Picture         =   "frmPacsInterfaceCfg.frx":6881
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14288
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   2880
      Top             =   900
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPacsInterfaceCfg.frx":7115
      Left            =   2100
      Top             =   1140
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPacsInterfaceCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const C_STR_CUSTOMPARAS = "[[系统号]]|[[模块号]]|[[科室ID]]|[[病人ID]]|[[医嘱ID]]|[[检查号]]|[[门诊号]]|[[住院号]]|[[身份证号]]|[[影像类别]]|[[用户名]]|[[账号名]]|[[当前窗口句柄]]"

Private Enum mAppCol
    序号 = 0: 程序名称: 程序版本: 程序路径: 程序ID: 程序集: 执行类型: 是否启用: 所属模块
End Enum

Private Enum mAppFuncCol
    序号 = 0: 功能名称:  启用功能: 加入右键菜单: 加入工具栏: 自动执行时机: 对应方法: 方法参数: VBS脚本: 功能ID: 验证通过
End Enum

Private Enum mAppFuncsCol
    方法序号 = 0: 功能方法: 方法参数
End Enum

Private Enum mAppFuncParaCol
    序号 = 0: 参数名称: 参数类型: 参数构造
End Enum

Private Enum mExecuteType
    动态创建 = 1: Shell命令: API声明
End Enum

'菜单类型枚举定义
Private Enum TMenuType
    mtFile = 1
    mtSave
    mtCancel
    mtQuit
    
    mtEdit
    mtAdd
    mtMod
    mtDel
    mtUse
    mtCheck
    mtRefresh
End Enum

Private mintTestSta As Integer  '测试状态 0 未测试 1 通过  2 未通过
Private mstr参数类型 As String
Private mblnConfiging As Boolean
Private mblnIsAddCfg As Boolean

Private mlngModule As Long
Private mstrPrivs As String
Private mlngAdviceID As Long
Private mlngSendNo As Long
Private mlngPatId As Long

Public Function ShowPacsInterfaceCfg(objParent As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
                                ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal lngPatId As Long) As Boolean
    
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    mlngPatId = lngPatId
    
    Call Me.Show(1, objParent)
End Function

Private Sub cboClasses_Click()
On Error GoTo ErrorHand
    
    Call FuncsFaceEnabled(mblnConfiging)
    Call ParasFaceEnabled(mblnConfiging)
    Call LoadAllClassFunc(txtAppDir.Text, cboClasses.Text)
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cboStation_Click()
On Error GoTo ErrorHand
    
    Call ClearAllCfg
    Call LoadAppInfo
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHnad
    
    Select Case Control.ID
        Case TMenuType.mtSave
            Call SaveAppCfg
            
        Case TMenuType.mtCancel
            Call CancelAppCfg
            
        Case TMenuType.mtAdd
            Call AddAppCfg
            
        Case TMenuType.mtMod
            Call ModAppCfg
            
        Case TMenuType.mtDel
            Call DelAppCfg
            
        Case TMenuType.mtUse
            Call UseAppCfg(Control)
            
        Case TMenuType.mtRefresh
            Call ClearAllCfg
            Call LoadAppInfo
            
        Case TMenuType.mtQuit
            Call Unload(Me)
            
'---------------------------查看----------------
        Case conMenu_View_ToolBar_Button '工具栏
            Call Menu_View_ToolBar_Button_click(Control)
        Case conMenu_View_ToolBar_Text '按钮文字
            Call Menu_View_ToolBar_Text_click(Control)
        Case conMenu_View_ToolBar_Size '大图标
            Call Menu_View_ToolBar_Size_click(Control)
        Case conMenu_View_StatusBar '状态栏
            Call Menu_View_StatusBar_click(Control)

'--------------------------帮助-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            Call Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
    
    End Select
    
    Exit Sub
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Function ValidData() As Boolean
'------------------------------------------------
'功能：检查输入数据的合法性
'参数： 无
'返回：True--数据输入合格，可以继续；False --有数据输入不合格，需要修改数据
'------------------------------------------------
On Error GoTo ErrorHnad
    
    ValidData = False
    
    '基本信息
    If Trim(txtAppName.Text) = "" Then
        MsgBox "程序名称不能为空，请输入！", vbExclamation, gstrSysName
        txtAppName.SetFocus
        Exit Function
        
    ElseIf Trim(txtAppDir.Text) = "" Then
        MsgBox "程序路径不能为空，请输入！", vbExclamation, gstrSysName
        txtAppDir.SetFocus
        Exit Function
        
    ElseIf cboType.ListIndex < 0 Then
        MsgBox "程序执行类型不能为空，请输入！", vbExclamation, gstrSysName
        cboType.SetFocus
        Exit Function
        
    ElseIf cboClasses.ListIndex < 0 Then
        MsgBox "程序集合不能为空，请输入！", vbExclamation, gstrSysName
        cboClasses.SetFocus
        Exit Function
    End If
    
    If CheckAppFuns() = False Then Exit Function
    If CheckAppfunPara() = False Then Exit Function
    
    ValidData = True
    
    Exit Function
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Function CheckAppFuns() As Boolean
'功能参数列表
    Dim i As Integer
    
    CheckAppFuns = True
    
    With vsfAppFuns
        If .Rows <= 1 Then
            MsgBox "请配置相关功能！", vbExclamation, gstrSysName
            CheckAppFuns = False
            Exit Function
        End If
        
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, mAppFuncCol.功能名称)) = "" Then
                MsgBox "功能名称不能为空，请输入！", vbExclamation, gstrSysName
                CheckAppFuns = False
                Exit Function
            End If
            
            If Trim(vsfFuncs.TextMatrix(0, 0)) = "" Then
                MsgBox "功能对应方法不能为空，请输入！", vbExclamation, gstrSysName
                CheckAppFuns = False
                Exit Function
            End If
        Next
    End With
End Function

Private Function CheckAppfunPara() As Boolean
'参数配置列表
    Dim i As Integer
    
    CheckAppfunPara = True
    
    With vsfAppfunPara
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, mAppFuncParaCol.参数类型)) = "" Then
                MsgBox "参数类型不能为空，请输入！", vbExclamation, gstrSysName
                CheckAppfunPara = False
                Exit Function
            End If
            
            If Trim(.TextMatrix(i, mAppFuncParaCol.参数构造)) = "" And Trim(.TextMatrix(i, mAppFuncParaCol.参数类型)) <> "字符串" Then
                MsgBox "参数构造不能为空，请输入！", vbExclamation, gstrSysName
                CheckAppfunPara = False
                Exit Function
            End If
        Next
    End With
End Function

Private Sub SaveAppCfg()
'------------------------------------------------
'功能：保存配置信息
'参数：无
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim lngAppId As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strParaInfo As String
    Dim strVBS As String
    Dim intType As Integer '执行时机
    
    
    '配置信息的有效性检查
    If Not ValidData Then Exit Sub

    '判断VBS是否修改过，若修改过要先更新到列表中
    If chkModify.value = 1 Then
        If Not (txtVBS.Text = vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.VBS脚本)) Then
            If CheckAppCfg(txtVBS.Text) Then
                vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.VBS脚本) = txtVBS.Text
            Else
                Exit Sub
            End If
        End If
    End If
    
    If Not DoBeforeSave() Then Exit Sub
    
    mblnConfiging = False
    
    Call InputFaceEnabled(False)
    Call AppFaceEnabled(True)
    chkModify.value = 0
    
    '新增时，获取程序ID
    If mblnIsAddCfg Then
        strSql = "Select Nvl(Max(ID), 0) + 1 as 程序ID From 影像插件挂接"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "")
        If rsTemp.RecordCount > 0 Then lngAppId = Val(rsTemp!程序ID)
    Else
        lngAppId = vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.程序ID)
    End If
    
    '保存影像插件挂接信息
    strSql = "ZL_影像插件挂接_Update(" & lngAppId & ",'" & _
                                         txtAppName.Text & "','" & _
                                         txtAppDir.tag & "','" & _
                                         txtAppDir.Text & "','" & _
                                         cboClasses.Text & "'," & _
                                         cboType.ItemData(cboType.ListIndex) & "," & _
                                         chkUseThisApp.value & "," & _
                                         cboStation.ItemData(cboStation.ListIndex) & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, "")
    
    '保存程序配置信息
    strSql = "ZL_影像插件功能_Delete(" & lngAppId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "")
    
    For i = 1 To vsfAppFuns.Rows - 1
        intType = CInt(convertInterfaceTime(vsfAppFuns.TextMatrix(i, mAppFuncCol.自动执行时机), True))
        strSql = "ZL_影像插件功能_Update(" & lngAppId & ",'" & _
                                             vsfAppFuns.TextMatrix(i, mAppFuncCol.功能名称) & "','" & _
                                             vsfAppFuns.TextMatrix(i, mAppFuncCol.对应方法) & "','" & _
                                             vsfAppFuns.TextMatrix(i, mAppFuncCol.方法参数) & "'," & _
                                             IIf(vsfAppFuns.Cell(flexcpChecked, i, mAppFuncCol.启用功能) = 1, 1, 0) & "," & _
                                             IIf(vsfAppFuns.Cell(flexcpChecked, i, mAppFuncCol.加入右键菜单) = 1, 1, 0) & "," & _
                                             IIf(vsfAppFuns.Cell(flexcpChecked, i, mAppFuncCol.加入工具栏) = 1, 1, 0) & "," & _
                                             intType & ",'" & _
                                             vsfAppFuns.TextMatrix(i, mAppFuncCol.VBS脚本) & "')"
        
        Call zlDatabase.ExecuteProcedure(strSql, "")
    Next
    
    If mblnIsAddCfg Then
        For i = 1 To vsfApp.Rows - 1
            With vsfApp
                If vsfApp.TextMatrix(i, mAppCol.程序名称) = "" Then
                    .TextMatrix(i, mAppCol.程序名称) = txtAppName.Text
                    .TextMatrix(i, mAppCol.程序版本) = txtAppDir.tag
                    .TextMatrix(i, mAppCol.程序路径) = txtAppDir.Text
                    .TextMatrix(i, mAppCol.程序ID) = lngAppId
                    .TextMatrix(i, mAppCol.程序集) = cboClasses.Text
                    .TextMatrix(i, mAppCol.执行类型) = cboType.ItemData(cboType.ListIndex)
                    .TextMatrix(i, mAppCol.是否启用) = chkUseThisApp.value
                    .TextMatrix(i, mAppCol.所属模块) = cboStation.ItemData(cboStation.ListIndex)
                    
                    Exit For
                End If
            End With
        Next
    Else
         With vsfApp
            .TextMatrix(.RowSel, mAppCol.程序名称) = txtAppName.Text
            .TextMatrix(.RowSel, mAppCol.程序版本) = txtAppDir.tag
            .TextMatrix(.RowSel, mAppCol.程序路径) = txtAppDir.Text
            .TextMatrix(.RowSel, mAppCol.程序ID) = lngAppId
            .TextMatrix(.RowSel, mAppCol.程序集) = cboClasses.Text
            .TextMatrix(.RowSel, mAppCol.执行类型) = cboType.ItemData(cboType.ListIndex)
            .TextMatrix(.RowSel, mAppCol.是否启用) = chkUseThisApp.value
            .TextMatrix(.RowSel, mAppCol.所属模块) = cboStation.ItemData(cboStation.ListIndex)
        End With
    End If
End Sub

Private Sub CancelAppCfg()
    mblnConfiging = False
    If mblnIsAddCfg Then txtAppName.Text = ""
    
    Call vsfApp_SelChange
    chkModify = 0
    Call InputFaceEnabled(False)
    Call AppFaceEnabled(True)
    
End Sub

Private Sub AddAppCfg()
    mblnConfiging = True
    mblnIsAddCfg = True
    txtAppName.Text = ""
    
    Call ClearInputCfg
    Call AppFaceEnabled(False)
    Call AppInfoFaceEnabled(True)
End Sub

Private Sub ModAppCfg()
    mblnConfiging = True
    mblnIsAddCfg = False
    
    Call AppFaceEnabled(False)
    Call InputFaceEnabled(True)
    cboType.Enabled = False
End Sub

Private Sub DelAppCfg()
    Dim strSql As String
    Dim lngAppId As Long
    
    lngAppId = vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.程序ID)
    If lngAppId <= 0 Then Exit Sub
    
    If MsgBox("确定要删除此程序配置吗？", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strSql = "ZL_影像插件挂接_Delete(" & lngAppId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "")
    
    Call ReLoadVSFList(vsfApp)
    Call vsfApp_SelChange
End Sub

Private Sub ReLoadVSFList(vsfList As VSFlexGrid)
    Dim i As Integer, j As Integer
    
    On Error GoTo ErrorHand
    
    If vsfList Is Nothing Then Exit Sub
    
    With vsfList
        For i = vsfList.RowSel To vsfList.Rows - 2
            For j = 1 To vsfList.Cols - 1
                vsfList.TextMatrix(i, j) = vsfList.TextMatrix(i + 1, j)
            Next
        Next
    End With
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub UseAppCfg(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngAppId As Long
    Dim strSql As String
    
    lngAppId = vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.程序ID)
    
    If lngAppId <= 0 Then Exit Sub
    
    chkUseThisApp.value = IIf(Control.Caption = "启用", 1, 0)
    vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.是否启用) = IIf(Control.Caption = "启用", 1, 0)
    
    strSql = "ZL_影像插件挂接_Update(" & lngAppId & ",'" & _
                                         txtAppName.Text & "','" & _
                                         txtAppDir.tag & "','" & _
                                         txtAppDir.Text & "','" & _
                                         cboClasses.Text & "'," & _
                                         cboType.ItemData(cboType.ListIndex) & "," & _
                                         chkUseThisApp.value & "," & _
                                         cboStation.ItemData(cboStation.ListIndex) & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, "")
End Sub

Private Function CheckAppCfg(ByVal strVBS As String, Optional ByVal blTest As Boolean = False) As Boolean
'调用vbs脚本实现功能,blTest 是否功能测试，若是，则需要完整的测试功能。
    Dim i As Integer
    Dim lngStart As Long, lngEnd As Long
    Dim ary() As String
    Dim strTmpVBS As String, strParaName As String, strParaVal As String
    Dim objCall As Object
    
On Error GoTo ErrorHnad
    
    ary = Split(strVBS, vbCrLf)
    
    For i = 0 To UBound(ary)
        '对于预定义参数，内部赋值
        strTmpVBS = ary(i)
        
        Do While InStr(strTmpVBS, "[[") > 0
            lngStart = InStr(strTmpVBS, "[[")
            lngEnd = InStr(strTmpVBS, "]]") + 2
            
            strParaName = Mid(strTmpVBS, lngStart, lngEnd - lngStart)
            
            Select Case strParaName
                Case "[[用户名]]"
                    strParaVal = "ZLHIS"
                                        
                Case "[[账号名]]"
                    strParaVal = "ZLHIS"
                                        
                Case "[[系统号]]"
                    strParaVal = "100"
                    
                Case "[[模块号]]"
                    strParaVal = "1291"
                    
                Case "[[科室ID]]"
                    strParaVal = "64"
                
                Case "[[病人ID]]"
                    strParaVal = "1"
                    
                Case "[[医嘱ID]]"
                    strParaVal = "101"
                    
                Case "[[检查号]]"
                    strParaVal = "110"
                    
                Case "[[门诊号]]"
                    strParaVal = "1"
                
                Case "[[住院号]]"
                    strParaVal = "110"
                    
                Case "[[身份证号]]"
                    strParaVal = "500105190001010000"
                    
                Case "[[影像类别]]"
                    strParaVal = "CT"
                                        
                Case "[[当前窗口句柄]]"
                    strParaVal = Me.hWnd
                                        
                Case Else
                    MsgBox "发现不能识别的预定义参数，请检查", vbExclamation, gstrSysName
                    CheckAppCfg = False
                    Exit Function
            End Select
            
            If strParaVal <> "------" Then strVBS = Replace(strVBS, strParaName, strParaVal)
            
            strTmpVBS = Trim(Mid(strTmpVBS, lngEnd))
        Loop
    Next
    
    CheckAppCfg = ExecuteSub(strVBS, Me, blTest)
    
    Exit Function
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
    CheckAppCfg = False
End Function

Public Function ExecuteSub(ByVal strVBS As String, ByVal objParent As Object, Optional ByVal blTest As Boolean = False) As Boolean
'调用vbs脚本实现功能
    Dim objCall As Object
    Dim strTempVBS As String
    
On Error GoTo ErrorHnad
    
    '创建脚本执行对象
    Set objCall = CreateObject("ScriptControl")
    objCall.Timeout = 60000
    objCall.Language = "vbscript"
    
    Call objCall.AddCode(strVBS)
    
    If blTest Then
        Call objCall.Run(Trim("ExcuteSub"))
    End If
    ExecuteSub = True
    
    Exit Function
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHnad
    
    Select Case Control.ID
        Case TMenuType.mtRefresh
            Control.Enabled = Not mblnConfiging
                        
        Case TMenuType.mtSave
            Control.Enabled = mblnConfiging
            
        Case TMenuType.mtCancel
            Control.Enabled = mblnConfiging
            
        Case TMenuType.mtAdd
            Control.Enabled = Not mblnConfiging
            
        Case TMenuType.mtMod
            Control.Enabled = vsfApp.RowSel > 0 And vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.程序名称) <> "" And Not mblnConfiging
            
        Case TMenuType.mtDel
            Control.Enabled = vsfApp.RowSel > 0 And vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.程序名称) <> "" And Not mblnConfiging
            
        Case TMenuType.mtUse
            Control.Enabled = vsfApp.RowSel > 0 And vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.程序名称) <> "" And Not mblnConfiging
            Control.Caption = IIf(Val(vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.是否启用)) = 1, "禁用", "启用")
                        
            Control.IconId = IIf(Val(vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.是否启用)) = 1, 3006, 3009)
            
            '用于及时刷新按钮状态
            Control.Enabled = Not Control.Enabled
            Control.Enabled = Not Control.Enabled
        
        Case TMenuType.mtQuit
            Control.Enabled = Not mblnConfiging
    End Select
    
    Exit Sub
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub chkModify_Click()
    txtVBS.Enabled = (chkModify.value = 1)
    cmdDelFun.Enabled = (chkModify.value = 0)
End Sub

Private Sub cmdAddFunc_Click()
    Dim i As Integer

On Error GoTo ErrorHand

    '检查配置
    If vsfAppFuns.Rows > 1 Then
        If Trim(vsfAppFuns.TextMatrix(vsfAppFuns.Rows - 1, mAppFuncCol.功能名称)) = "" Then
            MsgBox "功能名称不能为空，请输入!", vbExclamation, gstrSysName
            Exit Sub
        ElseIf Trim(vsfAppFuns.TextMatrix(vsfAppFuns.Rows - 1, mAppFuncCol.对应方法)) = "" Then
            MsgBox "功能方法不能为空，请输入!", vbExclamation, gstrSysName
            Exit Sub
        End If
        
        '对功能的参数进行检查
        If vsfAppfunPara.Rows <= 1 Then
            '没有配置参数，不需要参数？
            If cboType.ItemData(cboType.ListIndex) <> mExecuteType.动态创建 Then
                If MsgBox("您没有对此功能方法进行参数配置，确定要继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        '对VBS脚本进行验证
        If Not CheckAppCfg(txtVBS.Text) Then Exit Sub
    End If
    
    vsfAppFuns.Rows = vsfAppFuns.Rows + 1
    vsfAppFuns.TextMatrix(vsfAppFuns.Rows - 1, mAppFuncCol.序号) = vsfAppFuns.Rows - 1
    vsfAppFuns.TextMatrix(vsfAppFuns.Rows - 1, mAppFuncCol.启用功能) = 1
    vsfAppFuns.Select vsfAppFuns.Rows - 1, 1
    vsfAppFuns.EditCell
    
    If cboType.Text = "Shell命令" Then vsfFuncs.ColComboList(mAppFuncsCol.功能方法) = txtAppDir.Text
    
    Call ParasFaceEnabled(True)
    Call FuncsFaceEnabled(True)
    Call VBSFaceEnabled(True)
    Call ClearFuncs
    
    '添加功能后，将不允许更改执行类型
    cboType.Enabled = False
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cmdAddPara_Click()
On Error GoTo ErrorHand
    
    vsfAppfunPara.Rows = vsfAppfunPara.Rows + 1
    vsfAppfunPara.TextMatrix(vsfAppfunPara.Rows - 1, mAppFuncParaCol.序号) = vsfAppfunPara.Rows - 1
    vsfAppfunPara.Select vsfAppfunPara.Rows - 1, 1
    vsfAppfunPara.EditCell
    vsfAppfunPara.Cell(flexcpAlignment, 0, 0, vsfAppfunPara.Rows - 1, vsfAppfunPara.Cols - 1) = flexAlignLeftCenter
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cmdConfigWindow_Click()
    Dim strResult As String
    Dim strTxtValueOld
    
    On Error GoTo ErrorHand
    
    strTxtValueOld = vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.参数构造)
    strResult = frmPacsInterfaceParEdit.EditPara(vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.参数构造), Me, (mblnConfiging And chkModify.value = 0))
    
    If mblnConfiging Then
        vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.参数构造) = strResult
        Call vsfAppfunPara_AfterEdit(0, mAppFuncParaCol.参数构造)
    End If
    
    Call RefreshCfg
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cmdDelFun_Click()
On Error GoTo ErrorHand
    Dim LngSel As Integer

    LngSel = vsfFuncs.Row

    If vsfFuncs.RowSel + 1 > vsfFuncs.Rows Then Exit Sub
    If vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.功能方法) = "" Then Exit Sub
    
    vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.方法参数) = ""
    vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.功能方法) = ""
    
    Call ReLoadVSFList(vsfFuncs)
    
    vsfFuncs.RowSel = 0
    vsfFuncs.Row = LngSel
    Call LoadFuncParaCfg
    Call DoFraFuncParasCaption

    Call RefreshCfg
    Call CreateVBS
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cmdDelFunc_Click()
On Error GoTo ErrorHand
    Dim i As Long, j As Long
    Dim lngRow As Long
    Dim blOneRow As Boolean '是否删除前只有一条数据
'''''删除功能前，先选中下一个有效的功能，在功能列表中，若后面有数据，则选中后一条数据，并且从后一条数据开始整体上移一行。若后面没有数据，则选中前一条数据
''''若已经是唯一数据，则提示不允许删除，请直接修改

    blOneRow = False
    If vsfAppFuns.Rows = 1 Then Exit Sub
    If vsfAppFuns.Rows = 2 Then blOneRow = True

    
    lngRow = vsfAppFuns.Row
    If lngRow = 0 Then lngRow = 1
    
    Call vsfAppFuns.RemoveItem(lngRow)
    
    If Not blOneRow Then
    '删除前不只一条数据
        For i = lngRow To vsfAppFuns.Rows - 1
            vsfAppFuns.TextMatrix(i, mAppFuncsCol.方法序号) = i
        Next
    
        If lngRow = vsfAppFuns.Rows Then
            '已经是最后一个，选择前面一个
            vsfAppFuns.Row = lngRow - 1
            vsfAppFuns.RowSel = lngRow - 1
        Else
            '不是最后一个，选择当前
            vsfAppFuns.Row = lngRow
            vsfAppFuns.RowSel = lngRow
        End If
    
        Call vsfAppFuns_SelChange
        Call ParasFaceEnabled(vsfAppFuns.Rows > 1)
        Call VBSFaceEnabled(vsfAppFuns.Rows > 1)
    Else
    '删除前只有一条数据

        Call ClearFuncs
        Call ClearAppFuncParaCfg
        Call DoFraFuncParasCaption
        txtVBS.Text = ""
        Call ParasFaceEnabled(vsfAppFuns.Rows > 1)
        Call VBSFaceEnabled(vsfAppFuns.Rows > 1)
        Call FuncsFaceEnabled(mblnConfiging)
    End If
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cmdDelPara_Click()
    If vsfAppfunPara.Rows <= 1 Then Exit Sub
    
    Call vsfAppfunPara.RemoveItem(vsfAppfunPara.Rows - 1)
    Call RefreshCfg
    
End Sub

Private Sub cmdSelectApp_Click()
On Error GoTo ErrorHand
    Dim strFilePath As String, strFileName As String
    
    dlgFile.Filter = "(*.exe)|*.exe|(*.dll)|*.dll|(*.ocx)|*.ocx|(*.tlb)|*.tlb|(*.*)|*.*"
    dlgFile.ShowOpen
    
    strFilePath = dlgFile.Filename
    strFileName = dlgFile.FileTitle
    cboType.Enabled = True
    
    Call FuncsFaceEnabled(False)
    Call ParasFaceEnabled(False)
    Call VBSFaceEnabled(False)
    Call ClearInputCfg
    Call ClearFuncs
    
    Call LoadAppConfig(strFilePath, strFileName)
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub LoadAppConfig(ByVal strFilePath As String, ByVal strFileName As String)
    Dim objFSO As New FileSystemObject
    
    If strFilePath = "" Then Exit Sub
    
    txtAppDir.Text = strFilePath
    
    txtAppDir.tag = objFSO.GetFileVersion(strFilePath)
    cboClasses.tag = strFileName
    
    vsfFuncs.ColComboList(mAppFuncsCol.功能方法) = ""
    
    '加载程序的所有功能和参数
    Call LoadAllClass(strFilePath, strFileName)
    Call picAppCfg_Resize
End Sub

Private Sub LoadAllClass(ByVal strFilePath As String, ByVal strFileName As String)
'根据dll或ocx部件加载其包含的程序集
    Dim i As Integer
    Dim objClassInfo As TypeLibInfo
    Dim objInterfaceInfo As InterfaceInfo

On Error GoTo ErrorHand
    cboClasses.Clear
    
    Set objClassInfo = TypeLibInfoFromFile(strFilePath)
    
    cboType.Clear
    cboType.AddItem "动态创建"
    cboType.ItemData(cboType.NewIndex) = mExecuteType.动态创建
    cboType.Text = "动态创建"
    cboType.Enabled = False
    cboClasses.Enabled = True
    
    For Each objInterfaceInfo In objClassInfo.Interfaces
        If Not objInterfaceInfo.VTableInterface Is Nothing Then
            cboClasses.AddItem objInterfaceInfo.Parent & "." & Mid(objInterfaceInfo.Name, 2)
            If objInterfaceInfo.Parent & "." & Mid(objInterfaceInfo.Name, 2) = strFileName Then
                cboClasses.ListIndex = cboClasses.NewIndex
            End If
        End If
    Next
    
    Exit Sub
ErrorHand:
    cboType.Clear
    cboType.AddItem "Shell命令"
    cboType.ItemData(cboType.NewIndex) = mExecuteType.Shell命令
    cboType.AddItem "API声明"
    cboType.ItemData(cboType.NewIndex) = mExecuteType.API声明
    
    If vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.程序名称) <> "" And vsfApp.RowSel > 0 Then
        If vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.执行类型) = 2 Then
            cboType.ListIndex = 0
        Else
            cboType.ListIndex = 1
        End If
    Else
        cboType.ListIndex = 0
        cboType.Enabled = True
    End If
    
    cboClasses.AddItem strFileName
    cboClasses.Text = strFileName
    cboClasses.Enabled = False
End Sub

Private Sub LoadAllClassFunc(ByVal strFileName As String, ByVal strClassName As String)
'根据程序集加载对应的方法
    Dim objClassInfo As TypeLibInfo
    Dim objInterfaceInfo As InterfaceInfo
    Dim objMemberInfo As MemberInfo
    Dim strFuncs As String
    
    If cboType.ItemData(cboType.ListIndex) = mExecuteType.动态创建 Then
        Set objClassInfo = TypeLibInfoFromFile(strFileName)
        
        For Each objInterfaceInfo In objClassInfo.Interfaces
            If Not objInterfaceInfo.VTableInterface Is Nothing Then
                If objInterfaceInfo.Parent & "." & Mid(objInterfaceInfo.Name, 2) = strClassName Then
                    For Each objMemberInfo In objInterfaceInfo.Members
                        If objMemberInfo.InvokeKind = INVOKE_FUNC Then
                            '如果是方法则加载到功能列表
                            If objMemberInfo.Name <> "QueryInterface" And objMemberInfo.Name <> "AddRef" _
                                And objMemberInfo.Name <> "Release" And objMemberInfo.Name <> "GetTypeInfoCount" _
                                And objMemberInfo.Name <> "GetTypeInfo" And objMemberInfo.Name <> "GetIDsOfNames" _
                                And objMemberInfo.Name <> "Invoke" Then
                                strFuncs = strFuncs & "|" & objMemberInfo.Name
                            End If
                        End If
                    Next
                End If
            End If
        Next
        
        vsfFuncs.ColComboList(mAppFuncsCol.功能方法) = IIf(strFuncs <> "", Mid(strFuncs, 2), "")
    Else
        vsfFuncs.ColComboList(mAppFuncsCol.功能方法) = ""
        If cboType.Text = "Shell命令" Then vsfFuncs.ColComboList(mAppFuncsCol.功能方法) = txtAppDir.Text
    End If
End Sub

Private Sub LoadParasWithFunc(ByVal strFileName As String, ByVal strClassName As String, ByVal strFuncName As String)
'根据选择的方法获取对应的参数
    Dim objClassInfo As TypeLibInfo
    Dim objInterfaceInfo As InterfaceInfo
    Dim objMemberInfo As MemberInfo
    Dim objParameterInfo As ParameterInfo
    Dim strParas As String
    
    Set objClassInfo = TypeLibInfoFromFile(strFileName)
    
    For Each objInterfaceInfo In objClassInfo.Interfaces
        If Not objInterfaceInfo.VTableInterface Is Nothing Then
            If objInterfaceInfo.Parent & "." & Mid(objInterfaceInfo.Name, 2) = strClassName Then
                For Each objMemberInfo In objInterfaceInfo.Members
                    If objMemberInfo.InvokeKind = INVOKE_FUNC Then
                        '如果是方法
                        If objMemberInfo.Name = strFuncName Then
                            For Each objParameterInfo In objMemberInfo.Parameters
                                Call AddParaToParasList(objParameterInfo.Name)
                            Next
                        End If
                    End If
                Next
            End If
        End If
    Next
End Sub

Private Sub AddParaToParasList(ByVal strParaName As String)
    With vsfAppfunPara
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mAppFuncParaCol.序号) = .Rows - 1
        .TextMatrix(.Rows - 1, mAppFuncParaCol.参数名称) = strParaName
    End With
End Sub

Private Sub cmdTestFunc_Click()
On Error GoTo ErrorHand
    Dim strVBS As String

    strVBS = txtVBS.Text
    
    If InStr(strVBS, "[[") = 0 Then
        '不含有预定义参数，直接进行验证
        If CheckAppCfg(strVBS, True) Then
            mintTestSta = 通过
        Else
            mintTestSta = 未通过
            Exit Sub
        End If
    Else
        '含有预定义参数，需要输入参数后验证。
        If CheckAppCfg(strVBS, False) Then
            mintTestSta = frmPacsInterfaceVBSTest.zlShowMe(strVBS, Me)
        Else
            mintTestSta = 未通过
            Exit Sub
        End If
    End If
    
    If mintTestSta = 未通过 Then MsgBox "验证失败，请检查。", vbExclamation, gstrSysName
    If mintTestSta = 通过 Then
        If (MsgBox("功能验证结束，" & vbLf & "请根据实际情况判断是否正常，正常请选‘是’。", vbYesNo, "测试结果")) = vbYes Then
            vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.验证通过) = 1
            vsfAppFuns.Cell(flexcpBackColor, vsfAppFuns.RowSel, 0) = vsfAppFuns.BackColorFixed
        Else
            mintTestSta = 未通过
            vsfAppFuns.Cell(flexcpBackColor, vsfAppFuns.RowSel, 0) = &HC0C0FF
        End If
    End If
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHand
    
    Call InitCommandBars
    Call InitFaceScheme
    
    Call InitAppfuncParaList
    Call InitAppFunsList
    Call InitAppList
    Call InitAppFuncs
    
    Call InitEdit
    Call InputFaceEnabled(False)
    Call RestoreWinState(Me)
    
    stbThis.Panels(4).Text = "操作人:" & UserInfo.姓名
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub LoadAppInfo()
    Dim i As Integer
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "Select ID,名称,版本,路径,程序集,执行类型,是否启用,所属模块 " & _
             "From 影像插件挂接 Where 所属模块 = [1] Order By ID"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "", cboStation.ItemData(cboStation.ListIndex))
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    With vsfApp
        For i = 1 To rsData.RecordCount
            .TextMatrix(i, mAppCol.程序名称) = zlCommFun.NVL(rsData!名称)
            .TextMatrix(i, mAppCol.程序版本) = zlCommFun.NVL(rsData!版本)
            .TextMatrix(i, mAppCol.程序路径) = zlCommFun.NVL(rsData!路径)
            .TextMatrix(i, mAppCol.程序ID) = zlCommFun.NVL(rsData!ID)
            .TextMatrix(i, mAppCol.程序集) = zlCommFun.NVL(rsData!程序集)
            .TextMatrix(i, mAppCol.执行类型) = zlCommFun.NVL(rsData!执行类型)
            .TextMatrix(i, mAppCol.是否启用) = zlCommFun.NVL(rsData!是否启用)
            .TextMatrix(i, mAppCol.所属模块) = zlCommFun.NVL(rsData!所属模块)
            
            rsData.MoveNext
        Next
        
        If .Rows > 1 Then
            If .Row = 1 Then
                .RowSel = 1
                Call vsfApp_SelChange
            Else
                .Row = 1
                .RowSel = 1
            End If
        End If
    End With
End Sub

Private Sub InitEdit()
    cboStation.Clear
    cboStation.AddItem "全部"
    cboStation.ItemData(cboStation.NewIndex) = 0
    
    cboStation.AddItem "影像医技工作站"
    cboStation.ItemData(cboStation.NewIndex) = 1290
    
    cboStation.AddItem "影像采集工作站"
    cboStation.ItemData(cboStation.NewIndex) = 1291
    
    cboStation.AddItem "影像病理工作站"
    cboStation.ItemData(cboStation.NewIndex) = 1294
    
    cboStation.ListIndex = 0
End Sub

Private Sub InitAppList()
    Dim i As Integer
    
On Error GoTo ErrorHand

    With vsfApp
        .Cols = 9
        .Rows = 51
        
        .ColWidth(mAppCol.序号) = 300
        .ColWidth(mAppCol.程序名称) = 1300
        .ColWidth(mAppCol.程序版本) = 900
        
        .TextMatrix(0, mAppCol.序号) = "≡"
        .TextMatrix(0, mAppCol.程序名称) = "插件名称"
        .TextMatrix(0, mAppCol.程序版本) = "插件版本"
        .TextMatrix(0, mAppCol.程序路径) = "程序路径"
        
        .TextMatrix(0, mAppCol.程序ID) = "程序ID"
        .TextMatrix(0, mAppCol.程序集) = "程序集"
        .TextMatrix(0, mAppCol.执行类型) = "执行类型"
        .TextMatrix(0, mAppCol.是否启用) = "是否启用"
        .TextMatrix(0, mAppCol.所属模块) = "所属模块"
        
        .ExtendLastCol = True
        
        For i = 4 To .Cols - 1
            .ColHidden(i) = True
        Next
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, mAppCol.序号) = i
            .TextMatrix(i, mAppCol.程序ID) = 0
        Next
        
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        .RowSel = 0
    End With
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub InitAppFunsList()
On Error GoTo ErrorHand
    Dim i As Integer
    

    
    With vsfAppFuns
        .Cols = 11
        .Rows = 1
        .WordWrap = True '自动分行显示
        .RowHeight(0) = 500 '大概设置为一个单元格显示2行
        .ColWidth(mAppFuncCol.功能名称) = 1000
        .ColWidth(mAppFuncCol.序号) = 300
        .ColWidth(mAppFuncCol.功能名称) = 1000
        .ColWidth(mAppFuncCol.启用功能) = 500
        .ColWidth(mAppFuncCol.加入右键菜单) = 500
        .ColWidth(mAppFuncCol.加入工具栏) = 500
        .ColWidth(mAppFuncCol.自动执行时机) = 1400
        .ColWidth(mAppFuncCol.对应方法) = 1000
        
        
        .TextMatrix(0, mAppFuncCol.序号) = "≡"
        .TextMatrix(0, mAppFuncCol.功能名称) = "功能名称"
        .TextMatrix(0, mAppFuncCol.启用功能) = "启用功能"
        .TextMatrix(0, mAppFuncCol.加入右键菜单) = "右键菜单"
        .TextMatrix(0, mAppFuncCol.加入工具栏) = "工具栏"
        .TextMatrix(0, mAppFuncCol.自动执行时机) = "自动执行时机"
        .TextMatrix(0, mAppFuncCol.对应方法) = "对应方法"
        .TextMatrix(0, mAppFuncCol.方法参数) = "方法参数"
        .TextMatrix(0, mAppFuncCol.VBS脚本) = "VBS脚本"
        .TextMatrix(0, mAppFuncCol.功能ID) = "功能ID"
        .TextMatrix(0, mAppFuncCol.验证通过) = "验证通过"
          
        .ColHidden(mAppFuncCol.方法参数) = True
        .ColHidden(mAppFuncCol.VBS脚本) = True
        .ColHidden(mAppFuncCol.功能ID) = True
        .ColHidden(mAppFuncCol.验证通过) = True
        
        .ExtendLastCol = True
        .ColDataType(2) = flexDTBoolean
        .ColDataType(3) = flexDTBoolean
        .ColDataType(4) = flexDTBoolean
        .ColComboList(5) = C_STR_INTERFACE_0 & "|" & C_STR_INTERFACE_1 & "|" & C_STR_INTERFACE_2 & "|" & C_STR_INTERFACE_3 & "|" & _
                                      C_STR_INTERFACE_4 & "|" & C_STR_INTERFACE_5 & "|" & C_STR_INTERFACE_6 & "|" & C_STR_INTERFACE_7 & "|" & _
                                      C_STR_INTERFACE_11 & "|" & C_STR_INTERFACE_12 & "|" & C_STR_INTERFACE_13 & "|" & C_STR_INTERFACE_14 & "|" & _
                                      C_STR_INTERFACE_15 & "|" & C_STR_INTERFACE_16 & "|" & C_STR_INTERFACE_17 & "|" & C_STR_INTERFACE_21 & "|" & _
                                      C_STR_INTERFACE_22
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, mAppFuncCol.序号) = i
            .TextMatrix(i, mAppFuncCol.功能ID) = 0
        Next
        
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
    End With
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub InitAppFuncs()
    Dim i As Integer
    With vsfFuncs
        .Clear
        
        .ExtendLastCol = True
        .Cols = 3
        .Rows = 10
        .FixedCols = 1
        .FixedRows = 0
        
        .ColWidth(mAppFuncsCol.方法序号) = 300
        .ColWidthMax = 300 '加上这句才能使 方法序号 宽度为300
        .ColHidden(mAppFuncsCol.方法参数) = True
        
        For i = 0 To .Rows - 1
            .TextMatrix(i, mAppFuncsCol.方法序号) = i + 1
        Next
        
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        
    End With
End Sub

Private Sub InitAppfuncParaList()
    Dim i As Integer

On Error GoTo ErrorHand
    
    With vsfAppfunPara
        .Cols = 4
        .Rows = 1 '清空数据
        
        .ColWidth(mAppFuncParaCol.序号) = 300
        .ColWidth(mAppFuncParaCol.参数名称) = 1100
        .ColWidth(mAppFuncParaCol.参数类型) = 1100
        
        .TextMatrix(0, mAppFuncParaCol.序号) = "≡"
        .TextMatrix(0, mAppFuncParaCol.参数名称) = "参数名称"
        .TextMatrix(0, mAppFuncParaCol.参数类型) = "参数类型"
        .TextMatrix(0, mAppFuncParaCol.参数构造) = "参数构造"
        
        .ColComboList(mAppFuncParaCol.参数类型) = "预定义|字符串|数字型|布尔型"
        
        .ExtendLastCol = True
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, mAppFuncParaCol.序号) = i
        Next
        
    End With
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub InputFaceEnabled(ByVal blnEnabled As Boolean)
    Call AppInfoFaceEnabled(blnEnabled)
    Call FuncsFaceEnabled(blnEnabled)
    Call ParasFaceEnabled(blnEnabled)
    Call VBSFaceEnabled(blnEnabled)
End Sub

Private Sub AppFaceEnabled(ByVal blnEnabled As Boolean)
    lblStation.Enabled = blnEnabled
    cboStation.Enabled = blnEnabled
    vsfApp.Enabled = blnEnabled
End Sub

Private Sub AppInfoFaceEnabled(ByVal blnEnabled As Boolean)
    lblAppDir.Enabled = blnEnabled
    txtAppDir.Enabled = blnEnabled
    cmdSelectApp.Enabled = blnEnabled
    
    lblType.Enabled = blnEnabled
    cboType.Enabled = blnEnabled
    
    lblClasses.Enabled = blnEnabled
    cboClasses.Enabled = blnEnabled
    
    lblAppName.Enabled = blnEnabled
    txtAppName.Enabled = blnEnabled
    
    chkUseThisApp.Enabled = blnEnabled
End Sub

Private Sub FuncsFaceEnabled(ByVal blnEnabled As Boolean)
    Dim blHaveFunc As Boolean 'vsfAppFuns是否有有效数据
    Dim blHaveFun As Boolean 'vsfFuncs是否有有效数据
    Dim i As Long
    
    blHaveFunc = False
    blHaveFun = False
    
    blHaveFunc = vsfAppFuns.Rows > 1
    
    For i = 1 To vsfFuncs.Rows - 1
        If vsfFuncs.TextMatrix(1, mAppFuncsCol.方法序号) <> "" Then
            blHaveFun = True
            Exit For
        End If
    Next

    
    vsfAppFuns.Editable = IIf(blnEnabled, flexEDKbdMouse, flexEDNone)
    cmdAddFunc.Enabled = blnEnabled
    
    cmdDelFunc.Enabled = blnEnabled And blHaveFunc
    cmdDelFun.Enabled = blnEnabled And blHaveFunc And blHaveFun
    
    cmdTestFunc.Enabled = blnEnabled And blHaveFunc
End Sub

Private Sub ParasFaceEnabled(ByVal blnEnabled As Boolean)
    vsfAppfunPara.Editable = IIf(blnEnabled, flexEDKbdMouse, flexEDNone)
    cmdAddPara.Enabled = blnEnabled
    cmdDelPara.Enabled = blnEnabled
    
    If vsfAppFuns.Rows <= 1 Then
        vsfFuncs.Editable = flexEDNone
    Else
        vsfFuncs.Editable = IIf(blnEnabled, flexEDKbdMouse, flexEDNone)
    End If
End Sub

Private Sub VBSFaceEnabled(ByVal blnEnabled As Boolean)
    fraVBS.Enabled = blnEnabled
    chkModify.Enabled = blnEnabled
    txtVBS.Enabled = blnEnabled And chkModify.value = 1
End Sub

Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    '设置菜单栏和工具栏风格
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True                                '显示按钮提示
        .AlwaysShowFullMenus = False                            '不常用的菜单项先隐藏
        .UseFadedIcons = False                                  '图标显示为褪色效果
        .IconsWithShadow = True                                 '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True                                '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True                                      '工具栏显示为大图标
        .SetIconSize True, 24, 24                               '设置大图标的尺寸
        .SetIconSize False, 16, 16                              '设置小图标的尺寸
    End With
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                      '设置控件显示风格
        .EnableCustomization False                             '是否允许自定义设置
        Set .Icons = zlCommFun.GetPubIcons                     '设置关联的图标控件
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '菜单定义
'Begin------------------------编辑菜单--------------------------------------默认可见
    cbrMain.ActiveMenuBar.Title = "菜单"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtFile, "文件(&F)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "保存(&S)"): cbrControl.IconId = 3091
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "取消(&C)"): cbrControl.IconId = 3565
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "退出(&Q)"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtEdit, "编辑(&E)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtAdd, "新增(&N)"): cbrControl.IconId = 4010
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtMod, "修改(&M)"): cbrControl.IconId = 3003
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtDel, "删除(&D)"): cbrControl.IconId = 4008
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtUse, "禁用(&A)"): cbrControl.IconId = 3006
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtRefresh, "刷新(&R)"): cbrControl.IconId = 3823: cbrControl.BeginGroup = True
    cbrControl.ShortcutText = "F5"
    
    'Begin----------------------查看菜单--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 1, "查看(V)")
    Call CreateViewAndHelpMenu(cbrMenuBar, Nothing)

    'Begin----------------------帮助菜单--------------------------------------默认可见
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 2, "帮助(H)")
    Call CreateViewAndHelpMenu(Nothing, cbrMenuBar)
    
    
    '---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "保存", "保存"): cbrControl.IconId = 3091
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "取消", "取消"): cbrControl.IconId = 3565
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtAdd, "新增", "新增"): cbrControl.IconId = 4010: cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtMod, "修改", "修改"): cbrControl.IconId = 3003
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtDel, "删除", "删除"): cbrControl.IconId = 4008
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtUse, "禁用", "禁用"): cbrControl.IconId = 3006
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtRefresh, "刷新", "刷新"): cbrControl.IconId = 791: cbrControl.BeginGroup = True

    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "退出", "退出"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    cbrControl.BeginGroup = True
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub InitFaceScheme()
    Dim Pane1 As Pane, Pane2 As Pane
    
     With Me.dkpMain
        .VisualTheme = ThemeOffice2003
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        
        .PanelPaintManager.BoldSelected = True
        .TabPaintManager.Position = xtpTabPositionLeft  'TAB放到左边显示
        .TabPaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .TabPaintManager.BoldSelected = True
        dkpMain.Options.DefaultPaneOptions = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
        
        Set Pane1 = .CreatePane(1, 300, 100, DockLeftOf)
        Pane1.Handle = picApp.hWnd
        
        Set Pane2 = .CreatePane(2, 500, 100, DockRightOf, Pane1)
        Pane2.Handle = picAppCfg.hWnd
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 15000 Then Me.Width = 15000
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorHand

    mblnConfiging = False
    Call SaveWinState(Me)
    
    Exit Sub
ErrorHand:
End Sub

Private Sub picApp_Resize()
    On Error Resume Next
    
    lblStation.Left = 120
    lblStation.Top = 240
    
    cboStation.Left = lblStation.Left + lblStation.Width + 240
    cboStation.Top = 200
    cboStation.Width = picApp.Width - cboStation.Left - 120
    
    vsfApp.Left = lblStation.Left
    vsfApp.Top = cboStation.Top + cboStation.Height + 120
    vsfApp.Width = picApp.Width - vsfApp.Left * 2
    vsfApp.Height = picApp.Height - vsfApp.Top - 360
End Sub

Private Sub picAppCfg_Resize()
    On Error Resume Next
    
    '基本信息
    fraAppInfo.Left = 120
    fraAppInfo.Top = 120
    fraAppInfo.Width = picAppCfg.Width - fraAppInfo.Left * 2
    
    lblAppDir.Left = txtAppName.Left + txtAppName.Width + 300
    txtAppDir.Width = fraAppInfo.Width - lblAppDir.Left - lblAppDir.Width - cmdSelectApp.Width - 120
    cmdSelectApp.Left = txtAppDir.Left + txtAppDir.Width
    
    lblClasses.Left = cboType.Left + cboType.Width + 300
    cboClasses.Width = fraAppInfo.Width - cboClasses.Left - chkUseThisApp.Width - 300
    chkUseThisApp.Left = fraAppInfo.Width - chkUseThisApp.Width - 60
    
    '功能配置
    fraAppFuns.Left = 120
    fraAppFuns.Top = fraAppInfo.Top + fraAppInfo.Height + 120
    fraAppFuns.Width = fraAppInfo.Width
    fraAppFuns.Height = picAppCfg.Height - fraAppFuns.Top - 120
    
    Call fraFuncs.Move(0, 0, fraAppFuns.Width)
    Call vsfAppFuns.Move(120, 300, fraAppFuns.Width - 240)
    
    vsfFuncs.Left = vsfAppFuns.Cell(flexcpLeft, 0, mAppFuncCol.对应方法) - 10
    vsfFuncs.Top = vsfAppFuns.Cell(flexcpTop, 0, mAppFuncCol.对应方法) + vsfAppFuns.Cell(flexcpHeight, 0, mAppFuncCol.对应方法) - 10
    vsfFuncs.Width = vsfAppFuns.Cell(flexcpWidth, 0, mAppFuncCol.对应方法)
    vsfFuncs.Height = vsfAppFuns.Height - vsfFuncs.Top
    
    If fraFuncs.Width - cmdDelFun.Width - 120 > (cmdTestFunc.Left + cmdTestFunc.Width + 300) Then
        cmdDelFun.Left = (fraFuncs.Width - cmdDelFun.Width) - 120
    Else
        cmdDelFun.Left = cmdTestFunc.Left + cmdTestFunc.Width + 300
    End If
'
    cmdDelFun.Top = cmdAddFunc.Top
    
    Call fraFuncParas.Move(0, fraFuncs.Height + 240, fraAppFuns.Width * 0.5, fraAppFuns.Height - fraFuncs.Height - 240)
        
    Call vsfAppfunPara.Move(120, 300, fraFuncParas.Width - 240, fraFuncParas.Height - 600)
    If cboType.ItemData(cboType.ListIndex) = mExecuteType.动态创建 Then
        Call vsfAppfunPara.Move(120, 300, fraFuncParas.Width - 240, fraFuncParas.Height - 600)
        cmdDelPara.Visible = False
        cmdAddPara.Visible = False
    Else
        Call cmdAddPara.Move(120, fraFuncParas.Height - cmdAddPara.Height - 360)
        Call cmdDelPara.Move(1560, fraFuncParas.Height - cmdDelPara.Height - 360)
        Call vsfAppfunPara.Move(120, 300, fraFuncParas.Width - 240, fraFuncParas.Height - 1200)
        cmdDelPara.Visible = True
        cmdAddPara.Visible = True
    End If
    
    Call fraVBS.Move(fraFuncParas.Left + fraFuncParas.Width, fraFuncParas.Top, fraAppFuns.Width * 0.5, fraAppFuns.Height - fraFuncs.Height - 120)
        
    Call txtVBS.Move(60, 300, fraFuncParas.Width - 120, fraFuncParas.Height - 540)

    
    Call ShowButton
End Sub

Private Sub ClearAllCfg()
    Call ClearEdit
    Call ClearAppCfg
    Call ClearAppFuncCfg
    Call ClearFuncs
    Call ClearAppFuncParaCfg
End Sub

Private Sub ClearInputCfg()
    Call ClearAppFuncCfg
    Call ClearFuncs
    Call ClearAppFuncParaCfg
End Sub

Private Sub ClearAppCfg()
    Dim i As Integer, j As Integer
    
    For i = 1 To vsfApp.Rows - 1
        For j = 1 To vsfApp.Cols - 1
            vsfApp.TextMatrix(i, j) = ""
        Next
    Next
End Sub

Private Sub ClearFuncs()
    Dim i As Integer, j As Integer
    
    For i = 0 To vsfFuncs.Rows - 1
        For j = 1 To vsfFuncs.Cols - 1
            vsfFuncs.TextMatrix(i, j) = ""
        Next
    Next
End Sub

Private Sub ClearEdit()
    txtAppDir.Text = ""
    chkUseThisApp.value = 1
    cboType.Clear
    cboClasses.Clear
End Sub


Private Sub ClearAppFuncCfg()
    vsfAppFuns.Rows = 1
    vsfAppFuns.RowSel = 0
    If mblnIsAddCfg Then
        Call ClearEdit
    End If
    
    txtVBS.Text = ""
End Sub

Private Sub ClearAppFuncParaCfg()
    vsfAppfunPara.Rows = 1
    vsfAppfunPara.RowSel = 0
End Sub

Private Sub txtVBS_KeyUp(KeyCode As Integer, Shift As Integer)
 On Error GoTo ErrorHand
    vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.VBS脚本) = txtVBS.Text
    If chkModify.value = 1 Then
        vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.验证通过) = 0
        Call VerifyAllFuns
    End If
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub txtVBS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error GoTo ErrorHand
    vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.VBS脚本) = txtVBS.Text
    If chkModify.value = 1 And Button = 2 Then
        vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.验证通过) = 0
        Call VerifyAllFuns
    End If
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub vsfApp_SelChange()
    Dim lngAppId As Long
    
On Error GoTo ErrorHand

    txtAppName.Text = ""
    Call ClearEdit
    Call ClearAppFuncCfg
    Call ClearFuncs
    Call ClearAppFuncParaCfg
    
    If vsfApp.RowSel <= 0 Then Exit Sub

    lngAppId = Val(vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.程序ID))
    If lngAppId <= 0 Then Exit Sub
    
    chkUseThisApp.value = IIf(vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.是否启用) = 1, 1, 0)
    txtAppName.Text = vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.程序名称)
    
    Call LoadAppConfig(vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.程序路径), vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.程序集))
    Call LoadCfgData(lngAppId)
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub LoadCfgData(ByVal lngAppId As Long)
    Dim i As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "Select 功能序号,名称,方法,方法参数,是否启用,是否加入右键菜单,是否加入工具栏,自动执行时机,VBS脚本 From 影像插件功能 Where 插件ID = [1] Order By 功能序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "", lngAppId)
    
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    vsfAppFuns.Rows = rsTemp.RecordCount + 1
    vsfAppFuns.RowSel = 0
    
    For i = 1 To rsTemp.RecordCount
        vsfAppFuns.TextMatrix(i, mAppFuncCol.序号) = i
        vsfAppFuns.TextMatrix(i, mAppFuncCol.功能名称) = zlCommFun.NVL(rsTemp!名称)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.启用功能) = zlCommFun.NVL(rsTemp!是否启用)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.加入右键菜单) = zlCommFun.NVL(rsTemp!是否加入右键菜单)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.加入工具栏) = zlCommFun.NVL(rsTemp!是否加入工具栏)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.自动执行时机) = convertInterfaceTime(zlCommFun.NVL(rsTemp!自动执行时机), False)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.对应方法) = zlCommFun.NVL(rsTemp!方法)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.方法参数) = zlCommFun.NVL(rsTemp!方法参数)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.VBS脚本) = zlCommFun.NVL(rsTemp!VBS脚本)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.功能ID) = zlCommFun.NVL(rsTemp!功能序号, 0)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.验证通过) = 1
        
        rsTemp.MoveNext
    Next
   
    vsfAppFuns.Cell(flexcpAlignment, 1, 0, vsfAppFuns.Rows - 1, vsfAppFuns.Cols - 1) = flexAlignLeftCenter
    If vsfAppFuns.Rows > 1 Then
        vsfAppFuns.RowSel = 1
        txtVBS.Text = vsfAppFuns.TextMatrix(1, mAppFuncCol.VBS脚本)
    End If
End Sub

Private Sub CreateVBS()
'根据配置构造VBS脚本
On Error GoTo ErrorHand
    Dim i As Integer, j As Integer
    Dim strVBS As String
    Dim strFuncs As String, strParas As String
    
    Dim strParaVal As String
    Dim strDefine As String, strReg As String, strDefines As String
    
    Dim strParasType As String
    Dim strReturn As String
    
    Dim strFuncInfo As String, strParaInfo As String
    Dim strFuncName As String, strFuncPara As String
    Dim strParaName As String, strParaType As String, strParaValu As String
    
    If vsfAppFuns.Rows < 2 Then Exit Sub
    
    '解析参数,并生成参数串，预定义|字符串|数字型|布尔型
    strFuncInfo = vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.对应方法)
    strParaInfo = vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.方法参数)
    
    For i = 0 To UBound(Split(strFuncInfo, "★"))
        If strFuncInfo <> "" Then
            strParas = ""
            strParasType = ""
            strDefine = ""
                
            If strParaInfo <> "" Then
                strFuncPara = Split(strParaInfo, "★")(i)
                 
                For j = 0 To UBound(Split(strFuncPara, ""))
                    strParaName = Split(Split(strFuncPara, "")(j), "※")(0)
                    strParaType = Split(Split(strFuncPara, "")(j), "※")(1)
                    strParaValu = Split(Split(strFuncPara, "")(j), "※")(2)
                    
                    Select Case strParaType
                        Case "预定义"
                            If cboType.ItemData(cboType.ListIndex) = mExecuteType.Shell命令 Then
                                strParas = strParas & " " & strParaValu
                            Else
                                strParas = strParas & ", """ & strParaValu & """"
                            End If
                            
                            strParasType = strParasType & "s"
                            
                        Case "字符串"
                            If cboType.ItemData(cboType.ListIndex) = mExecuteType.Shell命令 Then
                                strParas = strParas & " " & strParaValu
                            Else
                                strParas = strParas & ", """ & strParaValu & """"
                            End If
                            
                            strParasType = strParasType & "s"
                            
                        Case "数字型"
                            strParas = strParas & ", " & Val(strParaValu)
                            
                            strParasType = strParasType & "l"
                            
                        Case "布尔型"
                            strParas = strParas & ", " & IIf(strParaValu = "", "False", strParaValu)
                            
                            strParasType = strParasType & "s"
                            
                    End Select
                Next
            End If
            
            strFuncName = Split(strFuncInfo, "★")(i)
            strReturn = "s"
            
            If strDefine <> "" Then strDefines = strDefines & strDefine
        
            Select Case cboType.Text
                Case "动态创建"
                    strFuncs = strFuncs & "    Call objExecute." & strFuncName & IIf(strParas = "", "", "(" & Mid(strParas, 2) & ")") & vbCrLf
    
                Case "Shell命令"
                    strFuncs = strFuncs & "    Call objExecute.exec (""" & strFuncName & IIf(strParas = "", """)", " " & Mid(strParas, 2) & """)") & vbCrLf
                    
                Case "API声明"
                    strReg = strReg & "    objExecute.Register """ & cboClasses.Text & """, """ & strFuncName & """, ""i=" & strParasType & """, ""R=" & strReturn & """" & vbCrLf
                    strFuncs = strFuncs & "    Call objExecute." & strFuncName & IIf(strParas = "", "", "(" & Mid(strParas, 2) & ")") & vbCrLf
                    
            End Select
        End If
    Next
    
    '生成脚本
    Select Case cboType.Text
        Case "动态创建"
            strVBS = "Sub ExcuteSub()" & vbCrLf & strDefines & _
                     "    Dim objExecute" & vbCrLf & _
                     "                " & vbCrLf & _
                     "    Set objExecute = CreateObject(""" & cboClasses.Text & """)" & vbCrLf & strFuncs & _
                     "End Sub"
        
        Case "Shell命令"
            strVBS = "Sub ExcuteSub()" & vbCrLf & strDefines & _
                     "    Dim objExecute" & vbCrLf & _
                     "                " & vbCrLf & _
                     "    Set objExecute = CreateObject(""wscript.shell"")" & vbCrLf & strFuncs & _
                     "End Sub"
        
        Case "API声明"
            strVBS = "Sub ExcuteSub()" & vbCrLf & strDefines & _
                     "    Dim objExecute" & vbCrLf & _
                     "                " & vbCrLf & _
                     "    Set objExecute = CreateObject(""DynamicWrapper"")" & vbCrLf & strReg & _
                     "                " & vbCrLf & strFuncs & _
                     "End Sub"
    End Select
    
    txtVBS.Text = strVBS
    vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.VBS脚本) = strVBS
    vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.验证通过) = 0
    
    Call VerifyAllFuns
    
    Exit Sub
ErrorHand:
    err.Raise -1, "CreateVBS", "[GetSelectRowAdviceID]" & vbCrLf & err.Description
End Sub

Private Sub vsfAppfunPara_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error GoTo ErrorHand
    Call ShowButton
Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub vsfAppfunPara_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If chkModify.value = 1 Then
        If Col = mAppFuncParaCol.参数构造 Or Col = mAppFuncParaCol.参数类型 Or Col = mAppFuncParaCol.参数名称 Then Cancel = True
    End If
    If Col = mAppFuncParaCol.参数类型 Then mstr参数类型 = vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.参数类型)
End Sub


Private Sub vsfAppfunPara_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error Resume Next
    
    If NewRow = 0 Or OldRow = NewRow Or vsfAppfunPara.Rows - 1 < OldRow Then Exit Sub
    
    Cancel = True
    
    If Trim(vsfAppfunPara.TextMatrix(OldRow, mAppFuncParaCol.参数类型)) = "" Then
        MsgBox "参数类型不能为空，请输入！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If Trim(vsfAppfunPara.TextMatrix(OldRow, mAppFuncParaCol.参数构造)) = "" And _
       Trim(vsfAppfunPara.TextMatrix(OldRow, mAppFuncParaCol.参数类型)) <> "字符串" Then
        MsgBox "参数构造不能为空，请输入！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    Cancel = False
End Sub

Private Sub vsfAppfunPara_SelChange()
On Error GoTo ErrorHand
    Call ShowButton
Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub vsfAppFuns_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = mAppFuncCol.对应方法 Then
        Call CreateVBS
    End If
End Sub

Private Sub vsfAppFuns_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'104536 操作vsfAppFuns水平滚动条后，同步调整vsfAppFuns内部vsfFuncs的位置
    On Error Resume Next
    vsfFuncs.Left = vsfAppFuns.Cell(flexcpLeft, 0, mAppFuncCol.对应方法) - 10
    vsfFuncs.Width = vsfAppFuns.Cell(flexcpWidth, 0, mAppFuncCol.对应方法)
End Sub

Private Sub vsfAppFuns_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If chkModify.value = 1 Then
        If Col = mAppFuncCol.对应方法 Then Cancel = True
    End If
End Sub

Private Sub vsfAppFuns_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error Resume Next
    
    If OldRow < 1 Or NewRow = 0 Or OldRow = NewRow Or vsfAppFuns.Rows - 1 < OldRow Then Exit Sub
    
    Cancel = True
    
    If Trim(vsfAppFuns.TextMatrix(OldRow, mAppFuncCol.功能名称)) = "" Then
        MsgBox "功能名称不能为空，请输入！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If Trim(vsfFuncs.TextMatrix(0, mAppFuncsCol.功能方法)) = "" And vsfAppFuns.Rows > 2 Then
        MsgBox "功能对应方法不能为空，请输入！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    Cancel = False
End Sub

Private Sub vsfAppFuns_SelChange()
On Error GoTo ErrorHand

    If vsfAppFuns.Row > 0 Then txtVBS.Text = vsfAppFuns.TextMatrix(vsfAppFuns.Row, mAppFuncCol.VBS脚本)

    Call LoadFuncFunCfg
    Call DoFraFuncParasCaption
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub LoadFuncFunCfg()
'加载对应功能的方法
On Error GoTo ErrorHand
    Dim i As Integer
    Dim strFuncInfo As String
    Dim strParaInfo As String
    
    If vsfAppFuns.RowSel <= 0 Then Exit Sub
    
    Call ClearFuncs
    
    If vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.对应方法) <> "" Then
        strFuncInfo = vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.对应方法)
        strParaInfo = vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.方法参数)
        
        For i = 0 To UBound(Split(strFuncInfo, "★"))
            vsfFuncs.TextMatrix(i, mAppFuncsCol.功能方法) = Split(strFuncInfo, "★")(i)
        Next
        
        For i = 0 To UBound(Split(strParaInfo, "★"))
            vsfFuncs.TextMatrix(i, mAppFuncsCol.方法参数) = Split(strParaInfo, "★")(i)
        Next
    End If
    
    If vsfFuncs.RowSel <> 0 Then
        vsfFuncs.Row = 0
        vsfFuncs.RowSel = 0
    Else
        Call LoadFuncParaCfg
    End If
    
    Exit Sub
ErrorHand:
    err.Raise -1, "LoadFuncFunCfg", "[GetSelectRowAdviceID]" & vbCrLf & err.Description
End Sub

Private Sub LoadFuncParaCfg()
'加载方法对应的参数
On Error GoTo ErrorHand
    Dim strParaInfo As String
    Dim i As Integer
    
    If vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.方法参数) <> "" Then
        strParaInfo = vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.方法参数)
        
        If strParaInfo <> "" Then
            vsfAppfunPara.Rows = UBound(Split(strParaInfo, "")) + 2
            
            For i = 0 To UBound(Split(strParaInfo, ""))
                vsfAppfunPara.TextMatrix(i + 1, mAppFuncParaCol.序号) = i + 1
                vsfAppfunPara.TextMatrix(i + 1, mAppFuncParaCol.参数名称) = Split(Split(strParaInfo, "")(i), "※")(0)
                vsfAppfunPara.TextMatrix(i + 1, mAppFuncParaCol.参数类型) = Split(Split(strParaInfo, "")(i), "※")(1)
                vsfAppfunPara.TextMatrix(i + 1, mAppFuncParaCol.参数构造) = Split(Split(strParaInfo, "")(i), "※")(2)
            Next
            
            vsfAppfunPara.Cell(flexcpAlignment, 1, 0, vsfAppfunPara.Rows - 1, vsfAppfunPara.Cols - 1) = flexAlignLeftCenter
        End If
    Else
        For i = 1 To vsfAppfunPara.Rows - 1
            vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.序号) = i
            vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.参数名称) = ""
            vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.参数类型) = ""
            vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.参数构造) = ""
        Next
    
    End If
    

    Exit Sub
ErrorHand:
    err.Raise -1, "LoadFuncParaCfg", "[GetSelectRowAdviceID]" & vbCrLf & err.Description
End Sub

Private Sub vsfAppfunPara_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHand

    Dim str参数类型 As String
        
    str参数类型 = vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.参数类型)
        
    If Col = mAppFuncParaCol.参数构造 Then
        If str参数类型 = "" Then
            MsgBox "请先选择参数赋值类型!", vbExclamation, gstrSysName
            vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.参数构造) = ""
            vsfAppfunPara.EditCell
        End If
    End If
    
    If Col = mAppFuncParaCol.参数类型 Then
        
        If mstr参数类型 <> str参数类型 Then
                
            vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.参数构造) = ""
            
            If str参数类型 = "数字型" Then
                vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.参数构造) = "0"
            ElseIf str参数类型 = "字符串" Then
                vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.参数构造) = ""
            ElseIf str参数类型 = "布尔型" Then
                vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.参数构造) = "False"
            End If
                        
        End If
                
    End If
    
    vsfAppfunPara.Cell(flexcpAlignment, 0, 0, vsfAppfunPara.Rows - 1, vsfAppfunPara.Cols - 1) = flexAlignLeftCenter
    
    Call RefreshCfg
    
    If Col = mAppFuncParaCol.参数类型 Or Col = mAppFuncParaCol.参数构造 Then
        Call CreateVBS
    End If
    
    Call ShowButton

    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub RefreshCfg()
    Dim i As Integer
    Dim strFuncInfo As String
    Dim strParaInfo As String
    
    For i = 1 To vsfAppfunPara.Rows - 1
        If vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.参数类型) = "" And vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.参数构造) <> "" Then
            MsgBox "请先在参数配置列表中选择对应参数赋值类型！", vbExclamation, gstrSysName
            Exit Sub
        End If
        
        strParaInfo = strParaInfo & "" & vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.参数名称) & "※" & vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.参数类型) & "※" & vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.参数构造)
    Next
    
    vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.方法参数) = IIf(strParaInfo <> "", Mid(strParaInfo, 2), "")
    
    strFuncInfo = ""
    strParaInfo = ""
    
    For i = 0 To vsfFuncs.Rows - 1
        If vsfFuncs.TextMatrix(i, mAppFuncsCol.功能方法) <> "" Then
            strFuncInfo = strFuncInfo & "★" & vsfFuncs.TextMatrix(i, mAppFuncsCol.功能方法)
            strParaInfo = strParaInfo & "★" & vsfFuncs.TextMatrix(i, mAppFuncsCol.方法参数)
        End If
    Next
    
    If vsfAppFuns.Rows > 1 Then
        vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.对应方法) = IIf(strFuncInfo <> "", Mid(strFuncInfo, 2), "")
        vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.方法参数) = IIf(strParaInfo <> "", Mid(strParaInfo, 2), "")
    End If

End Sub

Private Sub vsfAppfunPara_EnterCell()
On Error GoTo ErrorHand
    If vsfAppfunPara.ColSel = mAppFuncParaCol.参数构造 Then
        If vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.参数类型) = "预定义" Then
            '?可以动态添加参数类型？
            vsfAppfunPara.ColComboList(mAppFuncParaCol.参数构造) = C_STR_CUSTOMPARAS
        ElseIf vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.参数类型) = "布尔型" Then
            vsfAppfunPara.ColComboList(mAppFuncParaCol.参数构造) = "True|False"
        Else
            vsfAppfunPara.ColComboList(mAppFuncParaCol.参数构造) = ""
        End If
        
    End If
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub ShowButton()
'在指定单元格显示配置按钮
    cmdConfigWindow.Visible = False
    
    With vsfAppfunPara
        If .RowSel < 1 Then Exit Sub
        
        cmdConfigWindow.Left = .Cell(flexcpLeft, .RowSel, mAppFuncParaCol.参数构造) + .Cell(flexcpWidth, .RowSel, mAppFuncParaCol.参数构造) - cmdConfigWindow.Width
        cmdConfigWindow.Top = .Cell(flexcpTop, .RowSel, mAppFuncParaCol.参数构造)
        cmdConfigWindow.Height = .Cell(flexcpHeight, .RowSel, mAppFuncParaCol.参数构造)
        
        If cmdConfigWindow.Top < .RowHeight(0) Then Exit Sub
    
        cmdConfigWindow.Visible = .TextMatrix(.RowSel, mAppFuncParaCol.参数类型) = "字符串"
    End With
End Sub

Private Sub vsfFuncs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHand
    If Col = mAppFuncsCol.功能方法 Then
        If cboType.ItemData(cboType.ListIndex) = mExecuteType.动态创建 Then
            '加载方法对应的参数
            Call ClearAppFuncParaCfg  '清空数据
            Call LoadParasWithFunc(txtAppDir.Text, cboClasses.Text, vsfFuncs.TextMatrix(Row, Col))
        End If
    End If
    
    vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.方法参数) = ""
    
    If Col = mAppFuncsCol.功能方法 Then
        Call CreateVBS
    End If
    
    Call RefreshCfg
    
    
    If Col = mAppFuncsCol.功能方法 Then
        Call DoFraFuncParasCaption
    End If
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub vsfFuncs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrorHand
    If chkModify.value = 1 Then
        If Col = mAppFuncsCol.方法参数 Or Col = mAppFuncsCol.功能方法 Then Cancel = True
    End If
    
    If vsfAppFuns.Rows <= 0 Then
        Cancel = True
        Exit Sub
    End If
    
    If Row = 0 Then Exit Sub
    
    If vsfFuncs.TextMatrix(Row - 1, mAppFuncsCol.功能方法) = "" Then Cancel = True
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub vsfFuncs_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error Resume Next
    If Col = mAppFuncsCol.功能方法 Then
        SendKeys "{ENTER}"
    End If
    err.Clear
End Sub

Private Sub vsfFuncs_SelChange()
On Error GoTo ErrorHand
    Call LoadFuncParaCfg
    Call DoFraFuncParasCaption
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub Menu_Help_Web_Mail_click()
On Error GoTo errHandle
    zlMailTo hWnd
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_About_click()
On Error GoTo errHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Help_click()
'功能：调用帮助主题
On Error GoTo errHandle
    ShowHelp App.ProductName, Me.hWnd, Me.Name
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Forum_click()
On Error GoTo errHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Home_click()
On Error GoTo errHandle
    zlHomePage hWnd
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_StatusBar_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    
    Control.Checked = Not Control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Button_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim i As Integer
    
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    Control.Checked = Not Control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Size_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    Control.Checked = Not Control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).Style
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If
        
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = intStyle
        Next
    Next
    
    Control.Checked = Not Control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function convertInterfaceTime(ByVal strText As String, ByVal intConvetType As Boolean) As String
'将执行时机由汉字转换成数字，默认为0(不执行)
'intConvetType说明 : true ：str转int  false：int转str
    If intConvetType Then
        convertInterfaceTime = 0
        
        Select Case strText
            Case C_STR_INTERFACE_0
                convertInterfaceTime = EInterfaceExeTime.不自动执行
            Case C_STR_INTERFACE_1
                convertInterfaceTime = EInterfaceExeTime.登记后
            Case C_STR_INTERFACE_2
                convertInterfaceTime = EInterfaceExeTime.报到后
            Case C_STR_INTERFACE_3
                convertInterfaceTime = EInterfaceExeTime.采图后
            Case C_STR_INTERFACE_4
                convertInterfaceTime = EInterfaceExeTime.报告保存后
            Case C_STR_INTERFACE_5
                convertInterfaceTime = EInterfaceExeTime.报告签名后
            Case C_STR_INTERFACE_6
                convertInterfaceTime = EInterfaceExeTime.报告审核后
            Case C_STR_INTERFACE_7
                convertInterfaceTime = EInterfaceExeTime.检查完成后
            Case C_STR_INTERFACE_11
                convertInterfaceTime = EInterfaceExeTime.取消登记时
            Case C_STR_INTERFACE_12
                convertInterfaceTime = EInterfaceExeTime.取消报到时
            Case C_STR_INTERFACE_13
                convertInterfaceTime = EInterfaceExeTime.删除图像时
            Case C_STR_INTERFACE_14
                convertInterfaceTime = EInterfaceExeTime.取消报告时
            Case C_STR_INTERFACE_15
                convertInterfaceTime = EInterfaceExeTime.取消签名时
            Case C_STR_INTERFACE_16
                convertInterfaceTime = EInterfaceExeTime.取消审核时
            Case C_STR_INTERFACE_17
                convertInterfaceTime = EInterfaceExeTime.取消完成时
            Case C_STR_INTERFACE_21
                convertInterfaceTime = EInterfaceExeTime.检查切换后
            Case C_STR_INTERFACE_22
                convertInterfaceTime = EInterfaceExeTime.报告驳回后
        End Select
        
    Else
        convertInterfaceTime = C_STR_INTERFACE_0
        
        Select Case strText
            Case EInterfaceExeTime.不自动执行
                convertInterfaceTime = C_STR_INTERFACE_0
            Case EInterfaceExeTime.登记后
                convertInterfaceTime = C_STR_INTERFACE_1
            Case EInterfaceExeTime.报到后
                convertInterfaceTime = C_STR_INTERFACE_2
            Case EInterfaceExeTime.采图后
                convertInterfaceTime = C_STR_INTERFACE_3
            Case EInterfaceExeTime.报告保存后
                convertInterfaceTime = C_STR_INTERFACE_4
            Case EInterfaceExeTime.报告签名后
                convertInterfaceTime = C_STR_INTERFACE_5
            Case EInterfaceExeTime.报告审核后
                convertInterfaceTime = C_STR_INTERFACE_6
            Case EInterfaceExeTime.检查完成后
                convertInterfaceTime = C_STR_INTERFACE_7
            Case EInterfaceExeTime.取消登记时
                convertInterfaceTime = C_STR_INTERFACE_11
            Case EInterfaceExeTime.取消报到时
                convertInterfaceTime = C_STR_INTERFACE_12
            Case EInterfaceExeTime.删除图像时
                convertInterfaceTime = C_STR_INTERFACE_13
            Case EInterfaceExeTime.取消报告时
                convertInterfaceTime = C_STR_INTERFACE_14
            Case EInterfaceExeTime.取消签名时
                convertInterfaceTime = C_STR_INTERFACE_15
            Case EInterfaceExeTime.取消审核时
                convertInterfaceTime = C_STR_INTERFACE_16
            Case EInterfaceExeTime.取消完成时
                convertInterfaceTime = C_STR_INTERFACE_17
            Case EInterfaceExeTime.检查切换后
                convertInterfaceTime = C_STR_INTERFACE_21
            Case EInterfaceExeTime.报告驳回后
                convertInterfaceTime = C_STR_INTERFACE_22
                
        End Select
    End If
    

End Function

Private Sub DoFraFuncParasCaption()
On Error GoTo errH
    Dim strCaption  As String
    strCaption = "参数列表"
    
    If Len(vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.功能名称)) > 0 And Len(vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.功能方法)) > 0 Then
        strCaption = strCaption & "[" & vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.功能名称) & " - "
        strCaption = strCaption & vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.功能方法) & "]"
    ElseIf Len(vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.功能名称)) > 0 And Len(vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.功能方法)) = 0 Then
        strCaption = strCaption & "[" & vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.功能名称) & "]"
    End If
    
    fraFuncParas.Caption = strCaption
    Exit Sub
errH:
    fraFuncParas.Caption = "参数列表"
End Sub

Private Sub VerifyAllFuns()
'检查是否所有功能已经通过验证
'&HC0C0FF 粉红色
On Error GoTo errH
    Dim i As Long

    With vsfAppFuns
        For i = 1 To vsfAppFuns.Rows - 1
            If vsfAppFuns.TextMatrix(i, mAppFuncCol.功能名称) = "" Then Exit Sub

            If vsfAppFuns.TextMatrix(i, mAppFuncCol.验证通过) <> "1" Then
                .Cell(flexcpBackColor, i, 0) = &HC0C0FF
            Else
                .Cell(flexcpBackColor, i, 0) = .BackColorFixed
            End If
        Next
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Function DoBeforeSave() As Boolean
'点击保存前的处理，若存在未验证的功能，则提示需要先验证
On Error GoTo errH
    Dim i As Long

    DoBeforeSave = True
    With vsfAppFuns
        For i = 1 To vsfAppFuns.Rows - 1

            If .Cell(flexcpBackColor, i, 0) = &HC0C0FF Then
                If (MsgBox("存在未经过验证或验证未通过的功能暂时不允许保存，选择‘是’跳过验证操作继续保存；选择‘否’先进行功能验证。", vbYesNo, gstrSysName)) = vbYes Then
                    DoBeforeSave = True
                Else
                    DoBeforeSave = False
                End If
                Exit Function
            End If

        Next
    End With
    Exit Function
errH:
    DoBeforeSave = False
    MsgBox err.Description, vbExclamation, gstrSysName
End Function


