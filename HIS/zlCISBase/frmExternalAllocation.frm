VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmExternalAllocation 
   Caption         =   "三方调用配置"
   ClientHeight    =   9960
   ClientLeft      =   165
   ClientTop       =   870
   ClientWidth     =   14505
   Icon            =   "frmExternalAllocation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   14505
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   73
      Top             =   9585
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmExternalAllocation.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20505
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin VB.PictureBox picEdit 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   5040
      ScaleHeight     =   8655
      ScaleWidth      =   9255
      TabIndex        =   3
      Top             =   120
      Width           =   9255
      Begin VB.Frame fra配置信息ZLBH 
         Caption         =   " 配置信息 "
         Height          =   735
         Left            =   1080
         TabIndex        =   45
         Top             =   3600
         Width           =   7695
         Begin VB.TextBox txtZLBH地址 
            Height          =   270
            Left            =   1035
            MaxLength       =   250
            TabIndex        =   47
            Top             =   315
            Width           =   5415
         End
         Begin VB.Label lblZLBH地址 
            AutoSize        =   -1  'True
            Caption         =   "ZLBH地址"
            Height          =   180
            Left            =   240
            TabIndex        =   46
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Frame fra配置信息EXE 
         Caption         =   " 配置信息 "
         Height          =   735
         Left            =   1080
         TabIndex        =   35
         Top             =   2880
         Width           =   7695
         Begin VB.CommandButton cmd访问路径 
            Caption         =   "…"
            Height          =   270
            Left            =   6480
            TabIndex        =   38
            TabStop         =   0   'False
            Tag             =   "分类"
            ToolTipText     =   "按*打开选择器"
            Top             =   315
            Width           =   270
         End
         Begin VB.TextBox txt程序访问路径 
            Height          =   270
            Left            =   1400
            MaxLength       =   250
            TabIndex        =   37
            Top             =   315
            Width           =   5055
         End
         Begin VB.Label lbl程序访问路径 
            AutoSize        =   -1  'True
            Caption         =   "程序访问路径"
            Height          =   180
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   1080
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   51
         Top             =   5280
         Width           =   7695
         _cx             =   13573
         _cy             =   1296
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
         BackColorBkg    =   16777215
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmExternalAllocation.frx":0E1C
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
         OwnerDraw       =   1
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
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   50
         Top             =   5280
         Width           =   7695
         _cx             =   13573
         _cy             =   1296
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
         BackColorBkg    =   16777215
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmExternalAllocation.frx":0EAD
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
         OwnerDraw       =   1
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
      Begin VB.Frame fra应用场景 
         Caption         =   " 应用场景 "
         Height          =   1815
         Left            =   120
         TabIndex        =   53
         Top             =   6360
         Width           =   7695
         Begin VB.PictureBox pic小图标 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   256
            Left            =   4627
            ScaleHeight     =   174.545
            ScaleMode       =   0  'User
            ScaleWidth      =   182.857
            TabIndex        =   61
            Top             =   682
            Width           =   256
            Begin VB.Image img小图标 
               Height          =   64
               Left            =   0
               Top             =   0
               Width           =   62
            End
         End
         Begin VB.CommandButton cmd大图标 
            Caption         =   "…"
            Height          =   240
            Left            =   5037
            TabIndex        =   70
            TabStop         =   0   'False
            Tag             =   "分类"
            ToolTipText     =   "按*打开选择器"
            Top             =   1410
            Width           =   255
         End
         Begin VB.CommandButton cmd小图标 
            Caption         =   "…"
            Height          =   240
            Left            =   5037
            TabIndex        =   62
            TabStop         =   0   'False
            Tag             =   "分类"
            ToolTipText     =   "按*打开选择器"
            Top             =   690
            Width           =   255
         End
         Begin VB.CommandButton cmd清空大图标 
            Caption         =   "×"
            Height          =   240
            Left            =   5342
            TabIndex        =   71
            TabStop         =   0   'False
            Tag             =   "分类"
            ToolTipText     =   "按*打开选择器"
            Top             =   1410
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmd清空小图标 
            Caption         =   "×"
            Height          =   240
            Left            =   5342
            TabIndex        =   63
            TabStop         =   0   'False
            Tag             =   "分类"
            ToolTipText     =   "按*打开选择器"
            Top             =   690
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox pic大图标 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   350
            Left            =   4627
            ScaleHeight     =   360
            ScaleMode       =   0  'User
            ScaleWidth      =   360
            TabIndex        =   69
            Top             =   1355
            Width           =   360
            Begin VB.Image img大图标 
               Height          =   44
               Left            =   0
               Top             =   0
               Width           =   46
            End
         End
         Begin VB.ComboBox cbo工具栏 
            Height          =   300
            Left            =   1000
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   1020
            Width           =   2055
         End
         Begin VB.CheckBox chk门诊医生工作站 
            Caption         =   "门诊医生工作站"
            Height          =   180
            Left            =   1000
            TabIndex        =   55
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chk住院医生工作站 
            Caption         =   "住院医生工作站"
            Height          =   180
            Left            =   3012
            TabIndex        =   56
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chk住院护士工作站 
            Caption         =   "住院护士工作站"
            Height          =   180
            Left            =   4905
            TabIndex        =   57
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox cbo菜单 
            Height          =   300
            Left            =   1000
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   660
            Width           =   2055
         End
         Begin VB.ComboBox cbo右键菜单 
            Height          =   300
            Left            =   1000
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   1380
            Width           =   2055
         End
         Begin VB.Label lbl显示小图标 
            AutoSize        =   -1  'True
            Caption         =   "小图标"
            Height          =   180
            Left            =   4050
            TabIndex        =   60
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lbl显示大图标 
            AutoSize        =   -1  'True
            Caption         =   "大图标"
            Height          =   180
            Left            =   4050
            TabIndex        =   68
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label lbl右键菜单 
            AutoSize        =   -1  'True
            Caption         =   "右键菜单"
            Height          =   180
            Left            =   240
            TabIndex        =   66
            Top             =   1440
            Width           =   720
         End
         Begin VB.Label lbl工具栏 
            AutoSize        =   -1  'True
            Caption         =   "工具栏"
            Height          =   180
            Left            =   420
            TabIndex        =   64
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label lbl菜单 
            AutoSize        =   -1  'True
            Caption         =   "菜单"
            Height          =   180
            Left            =   600
            TabIndex        =   58
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lbl应用场合 
            AutoSize        =   -1  'True
            Caption         =   "应用场合"
            Height          =   180
            Left            =   240
            TabIndex        =   54
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Frame fra配置信息FTP 
         Caption         =   " 配置信息 "
         Height          =   2175
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   7695
         Begin VB.TextBox txtFTP访问目录 
            Height          =   270
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   28
            Top             =   1035
            Width           =   5175
         End
         Begin VB.CommandButton cmdFTP连接测试 
            Caption         =   "FTP连接测试"
            Height          =   350
            Left            =   5025
            TabIndex        =   34
            Top             =   1710
            Width           =   1335
         End
         Begin VB.TextBox txtFTP地址 
            Height          =   270
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   20
            Top             =   315
            Width           =   2070
         End
         Begin VB.TextBox txtFTP密码 
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   4305
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   26
            Top             =   675
            Width           =   2055
         End
         Begin VB.TextBox txtFTP用户名 
            Height          =   270
            Left            =   1200
            MaxLength       =   25
            TabIndex        =   24
            Top             =   675
            Width           =   2055
         End
         Begin VB.TextBox txtFTP本地目录 
            Height          =   270
            Left            =   1200
            MaxLength       =   250
            TabIndex        =   30
            Top             =   1395
            Width           =   5175
         End
         Begin VB.TextBox txtFTP端口 
            Height          =   270
            Left            =   4305
            MaxLength       =   10
            TabIndex        =   22
            Top             =   315
            Width           =   2055
         End
         Begin VB.TextBox txt文件下载名 
            Height          =   270
            Left            =   1185
            MaxLength       =   50
            TabIndex        =   33
            Top             =   1755
            Width           =   2055
         End
         Begin VB.CommandButton cmd下载目录 
            Caption         =   "…"
            Height          =   270
            Left            =   6405
            TabIndex        =   31
            TabStop         =   0   'False
            Tag             =   "分类"
            ToolTipText     =   "按*打开选择器"
            Top             =   1395
            Width           =   270
         End
         Begin VB.Label lblFTP访问目录 
            AutoSize        =   -1  'True
            Caption         =   "FTP访问目录"
            Height          =   180
            Left            =   180
            TabIndex        =   27
            Top             =   1080
            Width           =   990
         End
         Begin VB.Label lbl文件下载名 
            AutoSize        =   -1  'True
            Caption         =   "文件下载名"
            Height          =   180
            Left            =   240
            TabIndex        =   32
            Top             =   1800
            Width           =   900
         End
         Begin VB.Label lblFTP端口 
            AutoSize        =   -1  'True
            Caption         =   "FTP端口"
            Height          =   180
            Left            =   3630
            TabIndex        =   21
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lblFTP本地目录 
            AutoSize        =   -1  'True
            Caption         =   "FTP本地目录"
            Height          =   180
            Left            =   180
            TabIndex        =   29
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label lblFTP密码 
            AutoSize        =   -1  'True
            Caption         =   "FTP密码"
            Height          =   180
            Left            =   3630
            TabIndex        =   25
            Top             =   720
            Width           =   630
         End
         Begin VB.Label lblFTP用户名 
            AutoSize        =   -1  'True
            Caption         =   "FTP用户名"
            Height          =   180
            Left            =   360
            TabIndex        =   23
            Top             =   720
            Width           =   810
         End
         Begin VB.Label lblFTP地址 
            AutoSize        =   -1  'True
            Caption         =   "FTP地址"
            Height          =   180
            Left            =   540
            TabIndex        =   19
            Top             =   360
            Width           =   630
         End
      End
      Begin VB.Frame fra接入方式 
         Caption         =   " 接入方式 "
         Height          =   650
         Left            =   120
         TabIndex        =   13
         Top             =   1915
         Width           =   7695
         Begin VB.OptionButton opt接入方式 
            Caption         =   "ZLBH"
            Height          =   180
            Index           =   3
            Left            =   6120
            TabIndex        =   17
            Top             =   300
            Width           =   735
         End
         Begin VB.OptionButton opt接入方式 
            Caption         =   "FTP"
            Height          =   180
            Index           =   2
            Left            =   4160
            TabIndex        =   16
            Top             =   300
            Width           =   615
         End
         Begin VB.OptionButton opt接入方式 
            Caption         =   "EXE"
            Height          =   180
            Index           =   1
            Left            =   2200
            TabIndex        =   15
            Top             =   300
            Width           =   615
         End
         Begin VB.OptionButton opt接入方式 
            Caption         =   "URL"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   300
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.Frame fra接口基本信息 
         Caption         =   " 接口基本信息 "
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   7695
         Begin VB.ComboBox cbo接口类别 
            Height          =   300
            Left            =   4250
            TabIndex        =   8
            Top             =   300
            Width           =   2655
         End
         Begin VB.TextBox txt编号 
            Height          =   270
            Left            =   645
            MaxLength       =   5
            TabIndex        =   6
            Top             =   315
            Width           =   2655
         End
         Begin VB.TextBox txt名称 
            Height          =   270
            Left            =   645
            MaxLength       =   50
            TabIndex        =   10
            Top             =   675
            Width           =   6255
         End
         Begin VB.TextBox txt说明 
            Height          =   630
            Left            =   645
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   1080
            Width           =   6255
         End
         Begin VB.Label lbl名称 
            AutoSize        =   -1  'True
            Caption         =   "名称"
            Height          =   180
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lbl说明 
            AutoSize        =   -1  'True
            Caption         =   "说明"
            Height          =   180
            Left            =   240
            TabIndex        =   11
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl接口类别 
            AutoSize        =   -1  'True
            Caption         =   "接口类别"
            Height          =   180
            Left            =   3480
            TabIndex        =   7
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lbl编号 
            AutoSize        =   -1  'True
            Caption         =   "编号"
            Height          =   180
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   360
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   5280
         Width           =   7695
         _cx             =   13573
         _cy             =   1296
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
         BackColorBkg    =   16777215
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmExternalAllocation.frx":0F3E
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
         OwnerDraw       =   1
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
      Begin VB.Frame fra配置信息URL 
         Caption         =   " 配置信息 "
         Height          =   1095
         Left            =   1080
         TabIndex        =   39
         Top             =   3240
         Width           =   7695
         Begin VB.OptionButton opt浏览器类型 
            Caption         =   "Chrome"
            Height          =   180
            Index           =   1
            Left            =   2535
            TabIndex        =   42
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton opt浏览器类型 
            Caption         =   "IE"
            Height          =   180
            Index           =   0
            Left            =   1320
            TabIndex        =   41
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.TextBox txtURL地址 
            Height          =   270
            Left            =   1320
            MaxLength       =   250
            TabIndex        =   44
            Top             =   675
            Width           =   5535
         End
         Begin VB.Label lblURL地址 
            AutoSize        =   -1  'True
            Caption         =   "URL地址"
            Height          =   180
            Left            =   510
            TabIndex        =   43
            Top             =   720
            Width           =   630
         End
         Begin VB.Label lbl浏览器类型 
            AutoSize        =   -1  'True
            Caption         =   "浏览器类型"
            Height          =   180
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Label lbl列表说明 
         AutoSize        =   -1  'True
         Caption         =   "列表说明..."
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   120
         TabIndex        =   52
         Top             =   6120
         Width           =   990
      End
      Begin VB.Label lbl提示信息 
         AutoSize        =   -1  'True
         Caption         =   "提示信息..."
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   120
         TabIndex        =   48
         Top             =   4920
         Width           =   990
      End
      Begin VB.Label lbl图标提示 
         AutoSize        =   -1  'True
         Caption         =   "说明：大图标要求24*24，ico格式；小图标要求16*16，ico格式。点激活键选择使用"
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   120
         TabIndex        =   72
         Top             =   8280
         Width           =   6660
      End
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4575
      ScaleWidth      =   3105
      TabIndex        =   0
      Top             =   1080
      Width           =   3105
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   3570
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2475
         _Version        =   589884
         _ExtentX        =   4366
         _ExtentY        =   6297
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1800
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExternalAllocation.frx":0FCF
            Key             =   "省略"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   900
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
      WordWrap        =   -1  'True
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
   Begin MSComDlg.CommonDialog cdl照片 
      Left            =   240
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.ImageManager imgFunc 
      Left            =   0
      Top             =   480
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmExternalAllocation.frx":32FD
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmExternalAllocation.frx":40D7
      Left            =   945
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmExternalAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const conPane_List = 201
Private Const conPane_Edit = 202
Private Const INTERNET_OPEN_TYPE_DIRECT     As Long = &H1           'direct to net

Private Const INTERNET_SERVICE_FTP          As Long = &H1
Private Const INTERNET_FLAG_KEEP_CONNECTION  As Long = &H400000    ' use keep-alive semantics
Private Const INTERNET_FLAG_PASSIVE         As Long = &H8000000   ' used for FTP connections

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_EDITBOX = &H10
Private Const BIF_USENEWUI = BIF_NEWDIALOGSTYLE Or BIF_EDITBOX
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'功能：打开连接Internet的会话
'说明：
'    sAgent--要调用Internet对话的应用程序名
'    lAccessType--请求的网络访问的类型
'备注：如果lAccessType设置为INTERNET_OPEN_TYPE_PRECONFIG，连接时就要基于
'    HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings
'    注册表路径下的注册表数值ProxyEnable、ProxyServer和 ProxyOverride
'    sProxyName--指定代理服务器的名字，访问类型设置为INTERNET_OPEN_TYPE_PROXY才有效
'    sProxyBypass--指定代理服务器的名字或地址，有设置此项时lpszProxyName指定的将失效
'函数返回值：如果函数调用失败，lngINet 为0。

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'功能：建立Internet连接，打开FTP会话
'说明：
'    hInternetSession--函数InternetOpen返回的Internet会话句柄
'    sServerName--要连接的服务器的名称或IP
'    nServerPort--要连接的Internet端口
'    sUsername--登录的用户帐号
'    sPassword--登录的口令
'    lService--要连接的服务器类型（这里是连接FTP服务器，连接的类型为常数INTERNET_SERVICE_FTP）
'    lFlags--如果传递x8000000，连接将使用被动FTP语义，传递0使用非被动语义
'    lContext--当使用回调函数时使用该参数，不使用回调服务传递0
'函数返回值：如果函数调用失败，lngINetConn 为0

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'功能：关闭Internet连接


Private mstrPrivs As String     '当前使用者权限串
Private mintEditType As Integer    '当前编辑栏状态。0-查看;1-新增;2-修改
Private mlngDelID As Long          '删除的ID。修改通过先删再插的方式
Private mbln显示停用 As Boolean

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'接入方式
Private Enum mTnterfaceType
    URL
    EXE
    FTP
    ZLBH
End Enum

'浏览类型
Private Enum mBrowserType
    IE
    Chrome
End Enum

'表格列
Private Enum mREPORT_COLUMN
    COL_ID
    col_是否停用
    COL_编号
    col_类别
    col_名称
    col_接入方式
    col_说明
End Enum

Private Sub InitCommandBars()
    '功能：初始化菜单。加载全部菜单，工具栏，弹出菜单等
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar

    'CommandBars属性设置
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    Me.cbsThis.VisualTheme = xtpThemeOffice2003

    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With

    Me.cbsThis.EnableCustomization False
    Set cbsThis.Icons = zlcommfun.GetPubIcons
    '-----------------------------------------------------

    '菜单设置
    '-----------------------------------------------------
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagStretched)

    '***文件
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        '文件-预览
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        '文件-打印
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")

        '文件-退出
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        cbrControlMain.BeginGroup = True
    End With

    '***编辑
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        '编辑-新增
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        '编辑-修改
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        '编辑-删除
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")

        '编辑-启用
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用(&U)")
        '编辑-停用
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Pause, "停用(&P)")
    End With

    '***查看
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        '查看-工具栏
        Set cbrControlMain = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
        cbrControl.Checked = True

        '查看-状态栏
        Set cbrControlMain = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        cbrControlMain.Checked = True
        
        '查看-列表
        Set cbrControlMain = .Add(xtpControlPopup, conMenu_View_Append, "列表(&L)")
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, comMenu_LIS_ShowListHead, "显示已停用(&S)", -1, False)
        cbrControl.Checked = True
        mbln显示停用 = True
        
        Set cbrControlMain = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
        cbrControlMain.BeginGroup = True
    End With

    '***帮助
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        '帮助-帮助主题
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")

        '帮助-WEB
        Set cbrControlMain = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False

        '帮助-关于
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…")
        cbrControlMain.BeginGroup = True
    End With
    '-----------------------------------------------------

    '快键绑定
    '-----------------------------------------------------
    With Me.cbsThis.KeyBindings
        '打印
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        '删除
        .Add FCONTROL, Asc("D"), conMenu_Edit_Delete
        '新增
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        '修改
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        '帮助主题
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With
    '-----------------------------------------------------

    '工具栏定义
    '-----------------------------------------------------
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched

    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Print, "打印")
        
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Save, "保存")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
        
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")

        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Edit_Pause, "停用")

        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        cbrControlMain.BeginGroup = True
    End With

    '显示风格
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
    '-----------------------------------------------------
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlControl.RPTCopyToVSF(Me.rptList, Me.vfgList) Is Nothing Then Exit Sub
     
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = "三方调用配置清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub Load接口类别()
    '功能：将左边列表已有的接口类别载入下拉框
    Dim rptRow As ReportRow
    Dim strTemp As String
    
    cbo接口类别.Clear
    
    If Me.rptList.Rows.Count > 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If InStr(";" & strTemp & ";", rptRow.Record(col_类别).Value) < 1 Then
                    cbo接口类别.AddItem rptRow.Record(col_类别).Value
                    strTemp = strTemp & IIf(strTemp = "", "", ";") & rptRow.Record(col_类别).Value
                End If
            End If
        Next
        
        cbo接口类别.ListIndex = -1
    End If

End Sub

Private Sub cbo接口类别_GotFocus()
    zlControl.TxtSelAll cbo接口类别
End Sub

Private Sub cbo接口类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'-+_!@#$%^&*(){}[];:,.<>?/|\、《》，。｛｝【】；：？、￥%……&（）", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    
    Select Case Control.ID
    Case conMenu_File_Preview
        '预览
        Call zlRptPrint(0)
        
    Case conMenu_File_Print
        '打印
        Call zlRptPrint(1)
        
    Case conMenu_File_Exit
        '退出
        Unload Me
        
    Case conMenu_Edit_NewItem
        '新增
        mintEditType = 1
        cmd清空大图标.Visible = False
        cmd清空小图标.Visible = False
        
        Call EnabledControl(mintEditType)
        Call ResetControl
        
    Case conMenu_Edit_Modify
        '修改
        mintEditType = 2
        mlngDelID = rptList.FocusedRow.Record.Item(mREPORT_COLUMN.COL_ID).Value
        Call EnabledControl(mintEditType)

    Case conMenu_Edit_Delete
        '删除
        Call DeleteItem
        
    Case conMenu_Edit_Reuse
        '启用
        Call StopAndStart(0)
        
    Case conMenu_Edit_Pause
        '停用
        Call StopAndStart(1)
    
    Case conMenu_Edit_Save
        '保存
        Call Save
        
    Case conMenu_Edit_Untread
        '取消
        Call Untread
        
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case comMenu_LIS_ShowListHead
        mbln显示停用 = Not mbln显示停用
        Call RefreshList
    Case conMenu_View_Refresh
        Call RefreshList
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    End Select
End Sub

Private Sub StopAndStart(ByVal intMode As Integer)
    '功能：项目停启
    'intMode：0-启动；1-停止
    Dim lngId As Long
    Dim strSql As String
    
    On Error GoTo ErrHandle
    
    lngId = rptList.FocusedRow.Record.Item(mREPORT_COLUMN.COL_ID).Value
    
    strSql = "Zl_三方调用目录_Stop("
    'ID
    strSql = strSql & lngId
    '是否停用
    strSql = strSql & "," & intMode
    strSql = strSql & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call RefreshList(lngId)
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DeleteItem()
    '功能：删除选中的项目
    Dim lngId As Long
    Dim strSql As String
    Dim strMsg As String
    
    On Error GoTo ErrHandle
       
    strMsg = "真的删除该项目？"
    strMsg = strMsg & vbCrLf & "——" & rptList.FocusedRow.Record(mREPORT_COLUMN.col_名称).Value
    If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
    lngId = rptList.FocusedRow.Record.Item(mREPORT_COLUMN.COL_ID).Value
    
    strSql = "Zl_三方调用目录_Delete("
    'ID
    strSql = strSql & lngId
    strSql = strSql & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call RefreshList
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        '新增
        Control.Enabled = (mintEditType = 0 And (zlStr.IsHavePrivs(mstrPrivs, "增删改")))
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        '修改,删除
        If mintEditType <> 0 Or rptList.FocusedRow Is Nothing Or rptList.FocusedRow.GroupRow Then
            Control.Enabled = False
        Else
            Control.Enabled = (zlStr.IsHavePrivs(mstrPrivs, "增删改"))
        End If
    Case conMenu_Edit_Reuse
        '启用
        If mintEditType <> 0 Or rptList.FocusedRow Is Nothing Or rptList.FocusedRow.GroupRow Then
            Control.Enabled = False
        Else
            Control.Enabled = (rptList.FocusedRow.Record.Item(mREPORT_COLUMN.col_是否停用).Value = 1 And (zlStr.IsHavePrivs(mstrPrivs, "增删改")))
        End If
    Case conMenu_Edit_Pause
        '停用
        If mintEditType <> 0 Or rptList.FocusedRow Is Nothing Or rptList.FocusedRow.GroupRow Then
            Control.Enabled = False
        Else
            Control.Enabled = (rptList.FocusedRow.Record.Item(mREPORT_COLUMN.col_是否停用).Value = 0 And (zlStr.IsHavePrivs(mstrPrivs, "增删改")))
        End If
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        '保存，取消
        Control.Enabled = (mintEditType <> 0)
    Case conMenu_File_Preview, conMenu_File_Print
        '预览，打印
        Control.Enabled = (mintEditType = 0)
    
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case comMenu_LIS_ShowListHead
        Control.Checked = mbln显示停用
    Case conMenu_View_Refresh
        Control.Enabled = (mintEditType = 0)
    End Select
End Sub

Private Sub cmdFTP连接测试_Click()
    Dim lngINet As Long
    Dim lngINetConn As Long
    
    lngINet = InternetOpen("FTP Control", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If lngINet <= 0 Then
        MsgBox "测试连接失败！", vbExclamation, "FTP连接测试"
        Exit Sub
    Else
        lngINetConn = InternetConnect(lngINet, txtFTP地址.Text, Val(txtFTP端口.Text), txtFTP用户名.Text, txtFTP密码.Text, INTERNET_SERVICE_FTP, INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_PASSIVE, 0)
        If lngINetConn = 0 Then
            Call InternetCloseHandle(lngINet)
            MsgBox "测试连接失败！", vbExclamation, "FTP连接测试"
            Exit Sub
        Else
            Call InternetCloseHandle(lngINet)
        End If
    End If

    MsgBox "测试连接成功！", vbInformation, "FTP连接测试"
End Sub

Private Sub Save()
    Dim rsTemp As ADODB.Recordset
    Dim lngId As Long       '新增ID
    Dim strSql As String
    Dim str地址 As String
    Dim int接入方式 As Integer
    Dim date发药时间 As Date
    Dim blnInTrans As Boolean
    Dim arrSql As Variant
    Dim i As Integer
    Dim intSel As Integer
    
    On Error GoTo ErrHandle
    
    date发药时间 = sys.Currentdate
    arrSql = Array()
    
    '检验是否录入完整
    '--------------------------
    If Trim(txt编号.Text) = "" Then
        MsgBox "编号未录入！"
        Exit Sub
    End If
    
    If cbo接口类别.Text = "" Then
        MsgBox "接口类别未录入！"
        Exit Sub
    End If
    
    If Trim(txt名称.Text) = "" Then
        MsgBox "名称未录入！"
        Exit Sub
    End If
    
    If opt接入方式(mTnterfaceType.URL).Value Then
        '***URL
        
        If InStr(txtURL地址.Text, "://") < 2 Then
            MsgBox "URL地址格式不正确！"
            Exit Sub
        End If
        
    ElseIf opt接入方式(mTnterfaceType.EXE).Value Then
        '***EXE
        
        If InStr(txt程序访问路径.Text, ".exe") < 2 Then
            MsgBox "程序访问路径格式不正确！"
            Exit Sub
        End If
        
    ElseIf opt接入方式(mTnterfaceType.FTP).Value Then
        '***FTP
        
        If Trim(txtFTP地址.Text) = "" Then
            MsgBox "FTP地址为空！"
            Exit Sub
        End If
        
        If Trim(txtFTP用户名.Text) = "" Then
            MsgBox "FTP用户名为空！"
            Exit Sub
        End If
        
        If Trim(txtFTP端口.Text) = "" Then
            MsgBox "FTP端口为空！"
            Exit Sub
        End If
    ElseIf opt接入方式(mTnterfaceType.ZLBH).Value Then
        '***ZLBH
        
        If Trim(txtZLBH地址.Text) = "" Then
            MsgBox "ZLBH地址为空！"
            Exit Sub
        End If
    End If
    
    '--------------------------
    
    '修改时先删除
    If mintEditType = 2 Then
        gstrSql = "Zl_三方调用目录_Delete("
        'ID
        gstrSql = gstrSql & mlngDelID
        gstrSql = gstrSql & ")"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSql
    End If
    
    '获取新增路径的ID
    '--------------------------
    strSql = "Select 三方调用目录_Id.Nextval ID From Dual"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    lngId = rsTemp!ID
    '--------------------------
    
    '根据不同的接入方式捕获界面上不同的数据
    '--------------------------
    If opt接入方式(mTnterfaceType.URL).Value Then
        intSel = mTnterfaceType.URL
        int接入方式 = 1
        str地址 = "'" & txtURL地址.Text & "'"
    ElseIf opt接入方式(mTnterfaceType.EXE).Value Then
        intSel = mTnterfaceType.EXE
        int接入方式 = 2
        str地址 = "'" & txt程序访问路径.Text & "'"
    ElseIf opt接入方式(mTnterfaceType.FTP).Value Then
        intSel = mTnterfaceType.FTP
        int接入方式 = 3
        str地址 = "Null"
    ElseIf opt接入方式(mTnterfaceType.ZLBH).Value Then
        int接入方式 = 4
        str地址 = "'" & txtZLBH地址.Text & "'"
    End If
    '--------------------------
    
    
    gstrSql = "Zl_三方调用目录_Insert("
    'ID
    gstrSql = gstrSql & lngId
    '编号
    gstrSql = gstrSql & "," & Val(txt编号.Text)
    '类别
    gstrSql = gstrSql & ",'" & cbo接口类别.Text & "'"
    '名称
    gstrSql = gstrSql & ",'" & txt名称.Text & "'"
    '说明
    gstrSql = gstrSql & ",'" & txt说明.Text & "'"
    '接入方式
    gstrSql = gstrSql & "," & int接入方式
    '浏览器类型
    gstrSql = gstrSql & "," & IIf(opt接入方式(mTnterfaceType.URL).Value, IIf(opt浏览器类型(mBrowserType.IE).Value, 1, 2), "Null")
    '应用场合
    gstrSql = gstrSql & ",'" & chk门诊医生工作站.Value & chk住院医生工作站.Value & chk住院护士工作站.Value & "'"
    '地址
    gstrSql = gstrSql & "," & str地址
    '是否停用
    gstrSql = gstrSql & ",0"
    'Ftp地址
    gstrSql = gstrSql & ",'" & IIf(opt接入方式(mTnterfaceType.FTP).Value, txtFTP地址.Text, "") & "'"
    'Ftp访问目录
    gstrSql = gstrSql & ",'" & IIf(opt接入方式(mTnterfaceType.FTP).Value, txtFTP访问目录.Text, "") & "'"
    'Ftp用户名
    gstrSql = gstrSql & ",'" & IIf(opt接入方式(mTnterfaceType.FTP).Value, txtFTP用户名.Text, "") & "'"
    'Ftp密码
    gstrSql = gstrSql & ",'" & IIf(opt接入方式(mTnterfaceType.FTP).Value, zlStr.Sm4EncryptEcb(txtFTP密码.Text), "") & "'"
    'Ftp本地目录
    gstrSql = gstrSql & ",'" & IIf(opt接入方式(mTnterfaceType.FTP).Value, txtFTP本地目录.Text, "") & "'"
    'Ftp端口
    gstrSql = gstrSql & ",'" & IIf(opt接入方式(mTnterfaceType.FTP).Value, txtFTP端口.Text, "") & "'"
    'Ftp文件名
    gstrSql = gstrSql & ",'" & IIf(opt接入方式(mTnterfaceType.FTP).Value, txt文件下载名.Text, "") & "'"
    '菜单显示
    gstrSql = gstrSql & "," & cbo菜单.ListIndex
    '工具栏显示
    gstrSql = gstrSql & "," & cbo工具栏.ListIndex
    '右键菜单显示
    gstrSql = gstrSql & "," & cbo右键菜单.ListIndex
    '修改人
    gstrSql = gstrSql & ",'" & gstrUserName & "'"
    '修改时间
    gstrSql = gstrSql & ",to_date('" & date发药时间 & "','yyyy-MM-dd hh24:mi:ss')"
    gstrSql = gstrSql & ")"
    
    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = gstrSql
    
    If Not opt接入方式(mTnterfaceType.ZLBH).Value Then
        
        
        With vsfList(intSel)
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("序号")) <> "" And .TextMatrix(i, .ColIndex("参数值")) <> "" Then
                    gstrSql = "Zl_三方调用参数_Insert("
                    '接口id
                    gstrSql = gstrSql & lngId
                    '序号
                    gstrSql = gstrSql & "," & .TextMatrix(i, .ColIndex("序号"))
                    '参数值
                    gstrSql = gstrSql & ",'" & .TextMatrix(i, .ColIndex("参数值")) & "'"
                    '备注
                    gstrSql = gstrSql & ",'" & .TextMatrix(i, .ColIndex("备注")) & "'"
                    'Sql
                    gstrSql = gstrSql & ",'" & .TextMatrix(i, .ColIndex("数据源")) & "'"
                    gstrSql = gstrSql & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSql
                End If
            Next
        End With
    End If
    
    '集中处理发药事务
    '--------------------------
    gcnOracle.BeginTrans
    blnInTrans = True

    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "新增三方目录")
    Next
    
    gcnOracle.CommitTrans
    blnInTrans = False
    '--------------------------
    
    '保存图标
    '--------------------------
    Call sys.SaveLob(100, 31, lngId, img小图标.Tag)
    Call sys.SaveLob(100, 32, lngId, img大图标.Tag)
    '--------------------------

    Call RefreshList(lngId)
    
    mintEditType = 0
    Call EnabledControl(mintEditType)
    
    Exit Sub
ErrHandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd大图标_Click()
    Dim pic As stdole.StdPicture
    Dim lngH As Long
    Dim lngW As Long
                    
    With cdl照片
        .CancelError = True
        .Filter = "图片文件(*.ico)|*.ico"
        
        On Error Resume Next
        .ShowOpen
        
        If err <> 0 Then
            '没选中文件
            err.Clear
        Else
            '判断图标的大小尺寸是否符合要求，大图标为24*24像素
            '----------------------------------------------
            Set pic = LoadPicture(.FileName)
            
            lngH = Int(pic.Height * 0.567 / 15 + 0.5)
            lngW = Int(pic.Width * 0.567 / 15 + 0.5)
            
            If lngH <> 24 Or lngW <> 24 Then
                MsgBox "请选择像素为24*24的图标！", vbInformation, gstrSysName
                Exit Sub
            End If
            '----------------------------------------------
            
            img大图标.Picture = LoadPicture(.FileName)

            If err <> 0 Then
                MsgBox "图片文件无效，或文件不存在。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            img大图标.Tag = .FileName
            
            cmd清空大图标.Visible = True
        End If
    End With
End Sub

Private Sub cmd清空大图标_Click()
    Set img大图标.Picture = Nothing
    img大图标.Tag = ""
    cmd清空大图标.Visible = False
End Sub

Private Sub cmd清空小图标_Click()
    Set img小图标.Picture = Nothing
    img小图标.Tag = ""
    cmd清空小图标.Visible = False
End Sub

Private Sub Untread()
    mintEditType = 0
    
    With rptList
        If Me.rptList.Rows.Count > 0 Then
            Call RefreshList
        End If
    End With
    
    Call EnabledControl(mintEditType)
End Sub

Public Function BrowseForFolder(Optional sTitle As String = "请选择文件夹") As String
    Dim intNull As Integer, lngIDList As Long
    Dim strPath As String, udtBI As BrowseInfo

    With udtBI
        .hWndOwner = 0 ' Me.hWnd
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_USENEWUI
    End With
    
    lngIDList = SHBrowseForFolder(udtBI)
    
    If lngIDList Then
        strPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lngIDList, strPath
        CoTaskMemFree lngIDList
        intNull = InStr(strPath, vbNullChar)
        
        If intNull Then
          strPath = Left$(strPath, intNull - 1)
        End If
    End If

    BrowseForFolder = strPath
End Function

Private Sub cmd下载目录_Click()
    '返回下载目录路径
    
    txtFTP本地目录.Text = BrowseForFolder
End Sub

Private Sub cmd小图标_Click()
    Dim pic As stdole.StdPicture
    Dim lngH As Long
    Dim lngW As Long

    With cdl照片
        .CancelError = True
        .Filter = "图片文件(*.ico)|*.ico"
        
        On Error Resume Next
        .ShowOpen
        
        If err <> 0 Then
            '没选中文件
            err.Clear
        Else
            '判断图标的大小尺寸是否符合要求，小图标为16*16像素
            '----------------------------------------------
            Set pic = LoadPicture(.FileName)
            
            lngH = Int(pic.Height * 0.567 / 15 + 0.5)
            lngW = Int(pic.Width * 0.567 / 15 + 0.5)
            
            If lngH <> 16 Or lngW <> 16 Then
                MsgBox "请选择像素为16*16的图标！", vbInformation, gstrSysName
                Exit Sub
            End If
            '----------------------------------------------

            img小图标.Picture = LoadPicture(.FileName)
            
            Debug.Print img小图标.Picture.Width
            
            If err <> 0 Then
                MsgBox "图片文件无效，或文件不存在。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            img小图标.Tag = .FileName
            
            cmd清空小图标.Visible = True
        End If
    End With
End Sub

Private Sub cmd访问路径_Click()
    '返回程序访问路径
    Dim str参数串 As String
    Dim rsTemp As New ADODB.Recordset
    Dim bln缓存 As Boolean
    Dim i As Integer
    
    str参数串 = ""
    bln缓存 = False
    
    If txt程序访问路径.Text <> "" Then
        If InStr(txt程序访问路径.Text, ".exe[") > 0 Then
            str参数串 = Mid(txt程序访问路径.Text, InStr(txt程序访问路径.Text, ".exe[") + 4)
            
            '先对当前的参数列表进行缓存
            If vsfList(mTnterfaceType.EXE).Rows > 1 Then
                With rsTemp
                    If .State = 1 Then .Close
                    
                    .Fields.Append "序号", adDouble, 18, adFldIsNullable
                    .Fields.Append "参数值", adLongVarChar, 200, adFldIsNullable
                    .Fields.Append "备注", adLongVarChar, 500, adFldIsNullable
                    .Fields.Append "数据源", adLongVarChar, 2000, adFldIsNullable
                    
                    .CursorLocation = adUseClient
                    .CursorType = adOpenStatic
                    .LockType = adLockOptimistic
                    .Open
                    
                    bln缓存 = True
                    
                    For i = 1 To vsfList(mTnterfaceType.EXE).Rows - 1
                        .AddNew
                        
                        !序号 = vsfList(mTnterfaceType.EXE).TextMatrix(i, vsfList(mTnterfaceType.EXE).ColIndex("序号"))
                        !参数值 = vsfList(mTnterfaceType.EXE).TextMatrix(i, vsfList(mTnterfaceType.EXE).ColIndex("参数值"))
                        !备注 = vsfList(mTnterfaceType.EXE).TextMatrix(i, vsfList(mTnterfaceType.EXE).ColIndex("备注"))
                        !数据源 = vsfList(mTnterfaceType.EXE).TextMatrix(i, vsfList(mTnterfaceType.EXE).ColIndex("数据源"))
                        
                        .Update
                    Next
                End With
            End If
        End If
    End If
    
    With cdl照片
        .CancelError = True
        .Filter = "可执行文件(*.exe)|*.exe"
        
        On Error Resume Next
        .ShowOpen
        
        If err <> 0 Then
            '没选中文件
            err.Clear
        Else
            If err <> 0 Then
                MsgBox "程序文件无效，或文件不存在。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            txt程序访问路径.Text = .FileName & str参数串
            
            If bln缓存 Then
                rsTemp.Filter = ""
                With vsfList(mTnterfaceType.EXE)
                    .Redraw = flexRDNone
                    
                    For i = 1 To rsTemp.RecordCount
                        .TextMatrix(i, .ColIndex("序号")) = rsTemp!序号
                        .TextMatrix(i, .ColIndex("参数值")) = rsTemp!参数值
                        .TextMatrix(i, .ColIndex("备注")) = zlcommfun.NVL(rsTemp!备注, "")
                        .TextMatrix(i, .ColIndex("数据源")) = zlcommfun.NVL(rsTemp!Sqltext, "")
                        
                        rsTemp.MoveNext
                    Next
                        
                    .Redraw = flexRDDirect
                End With
            End If
        End If
    End With
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = picList.hwnd
    Case conPane_Edit
        Item.Handle = picEdit.hwnd
    End Select
End Sub

Private Sub InitPanes()
    '功能：初始化DockingPane控件
    Dim panList As Pane
    Dim panEdit As Pane

    Set panList = dkpMan.CreatePane(conPane_List, 500, 1000, DockLeftOf, Nothing)
    panList.Title = "信息列表"
    panList.Options = PaneNoCaption
    
    Set panEdit = dkpMan.CreatePane(conPane_Edit, 500, 1000, DockRightOf, panList)
    panEdit.Title = "信息编辑"
    panEdit.Options = PaneNoCaption

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
End Sub

Private Sub InitReportControl()
    '功能：初始化ReportControl控件
    Dim objCol As ReportColumn

    With rptList
        .Columns.DeleteAll
        .AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 800)

        Set objCol = .Columns.Add(COL_ID, "ID", 0, False)
        objCol.Sortable = False
        objCol.Visible = False
        
        Set objCol = .Columns.Add(col_是否停用, "是否停用", 0, False)
        objCol.Sortable = False
        objCol.Visible = False
        
        Set objCol = .Columns.Add(COL_编号, "编号", 150, True)
        objCol.Sortable = True
        objCol.Visible = True
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentRight
        .SortOrder.Add objCol       '默认升序排序
        
        Set objCol = .Columns.Add(col_类别, "类别", 200, True)
        objCol.Sortable = False
        objCol.Visible = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft
        objCol.TreeColumn = True

        Set objCol = .Columns.Add(col_名称, "名称", 400, True)
        objCol.Sortable = True
        objCol.Visible = True
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft

        Set objCol = .Columns.Add(col_接入方式, "接入方式", 200, True)
        objCol.Sortable = True
        objCol.Visible = True
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft

        Set objCol = .Columns.Add(col_说明, "说明", 500, True)
        objCol.Sortable = False
        objCol.Visible = True
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft

        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = (objCol.Index = col_类别)
        Next
        
        .AllowColumnRemove = False
        .MultipleSelection = False  '不允许多行选择。会引发SelectionChanged事件
        .ShowItemsInGroups = False  '不显示已分组的列

        .GroupsOrder.Add .Columns(col_类别)

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的配置..."
        End With
    End With
End Sub

Private Sub InitControl()
    '功能：初始化控件的属性及数据
    
    With cbo菜单
        .Clear
        .AddItem "0-不显示"
        .AddItem "1-显示在子菜单中"
        .AddItem "2-显示在主菜单中"
        .ListIndex = 0
    End With
    
    With cbo工具栏
        .Clear
        .AddItem "0-不显示"
        .AddItem "1-显示在子工具栏中"
        .AddItem "2-显示在主工具栏中"
        .ListIndex = 0
    End With
    
    With cbo右键菜单
        .Clear
        .AddItem "0-不显示"
        .AddItem "1-显示在子菜单中"
        .AddItem "2-显示在主菜单中"
        .ListIndex = 0
    End With
    
    With img大图标
        .Left = pic大图标.ScaleLeft
        .Top = pic大图标.ScaleTop
        .Width = pic大图标.ScaleWidth
        .Height = pic大图标.ScaleHeight
    End With
    
    With img小图标
        .Left = pic小图标.ScaleLeft
        .Top = pic小图标.ScaleTop
        .Width = pic小图标.ScaleWidth
        .Height = pic小图标.ScaleHeight
    End With
End Sub

Private Sub InitvsfList()
    '功能：初始化参数列表
    
    With vsfList(mTnterfaceType.URL)
        .Editable = flexEDKbdMouse
        .ColComboList(.ColIndex("参数值")) = "病人ID|就诊ID|医嘱ID|部门ID|登录用户名|操作员编号|操作员姓名|数据源提取|"
    End With
    
    With vsfList(mTnterfaceType.EXE)
        .Editable = flexEDKbdMouse
        .ColComboList(.ColIndex("参数值")) = "病人ID|就诊ID|医嘱ID|部门ID|登录用户名|操作员编号|操作员姓名|数据源提取|"
    End With
    
    With vsfList(mTnterfaceType.FTP)
        .Editable = flexEDKbdMouse
        .ColComboList(.ColIndex("参数值")) = "病人ID|就诊ID|医嘱ID|部门ID|登录用户名|操作员编号|操作员姓名|数据源提取|"
    End With
    
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    
    mintEditType = 0
    
    Call zlcommfun.SetWindowsInTaskBar(Me.hwnd, False)
    Call InitCommandBars
    Call InitPanes
    Call InitReportControl
    Call InitControl
    Call InitvsfList

    Call RefreshList
    Call EnabledControl(mintEditType)
    
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub EnabledControl(ByVal intEditType As Integer)
    '功能：控制编辑面板是否可用
    
    picList.Enabled = (intEditType = 0)
    
    '文本框
    txt编号.Enabled = (intEditType <> 0)
    cbo接口类别.Enabled = (intEditType <> 0)
    txt名称.Enabled = (intEditType <> 0)
    txt说明.Enabled = (intEditType <> 0)
    
    txt程序访问路径.Enabled = (intEditType <> 0)
    cmd访问路径.Enabled = (intEditType <> 0)
    
    txtURL地址.Enabled = (intEditType <> 0)
    
    txtFTP地址.Enabled = (intEditType <> 0)
    txtFTP访问目录.Enabled = (intEditType <> 0)
    txtFTP用户名.Enabled = (intEditType <> 0)
    txtFTP密码.Enabled = (intEditType <> 0)
    txtFTP端口.Enabled = (intEditType <> 0)
    txt文件下载名.Enabled = (intEditType <> 0)
    txtFTP本地目录.Enabled = (intEditType <> 0)
    cmd下载目录.Enabled = (intEditType <> 0)
    
    txtZLBH地址.Enabled = (intEditType <> 0)
    
    cbo菜单.Enabled = (intEditType <> 0)
    cbo工具栏.Enabled = (intEditType <> 0)
    cbo右键菜单.Enabled = (intEditType <> 0)
    
    cmd小图标.Enabled = (intEditType <> 0)
    cmd大图标.Enabled = (intEditType <> 0)
    
    cmd清空小图标.Enabled = (intEditType <> 0)
    cmd清空大图标.Enabled = (intEditType <> 0)
    
    cmdFTP连接测试.Enabled = (intEditType <> 0)
    
    chk门诊医生工作站.Enabled = (intEditType <> 0)
    chk住院医生工作站.Enabled = (intEditType <> 0)
    chk住院护士工作站.Enabled = (intEditType <> 0)
    
    opt接入方式(mTnterfaceType.EXE).Enabled = (intEditType <> 0)
    opt接入方式(mTnterfaceType.FTP).Enabled = (intEditType <> 0)
    opt接入方式(mTnterfaceType.URL).Enabled = (intEditType <> 0)
    opt接入方式(mTnterfaceType.ZLBH).Enabled = (intEditType <> 0)
    
    vsfList(mTnterfaceType.EXE).Enabled = (intEditType <> 0)
    vsfList(mTnterfaceType.FTP).Enabled = (intEditType <> 0)
    vsfList(mTnterfaceType.URL).Enabled = (intEditType <> 0)
    
    opt浏览器类型(mBrowserType.IE).Enabled = (intEditType <> 0)
    opt浏览器类型(mBrowserType.Chrome).Enabled = (intEditType <> 0)
End Sub

Private Sub ResetControl()
    '功能：清空编辑面板中所有空间的数据
    Dim objTemp As Control
    
    For Each objTemp In Me.Controls
        '清空文本
        If TypeName(objTemp) = "TextBox" Then
            objTemp.Text = ""
        End If
        
        '重置多选框
        If TypeName(objTemp) = "CheckBox" Then
            objTemp.Value = 1
        End If
    Next
    
    '重置单选框
    opt接入方式(mTnterfaceType.URL).Value = True
    opt浏览器类型(mBrowserType.IE).Value = True
    
    '重置表格
    vsfList(mTnterfaceType.URL).Rows = 1
    vsfList(mTnterfaceType.EXE).Rows = 1
    vsfList(mTnterfaceType.FTP).Rows = 1
    
    '清空图标
    Set img小图标.Picture = Nothing
    img小图标.Tag = ""
    
    Set img大图标.Picture = Nothing
    img大图标.Tag = ""
    
    '重置下拉列表
    cbo菜单.ListIndex = 0
    cbo工具栏.ListIndex = 0
    cbo右键菜单.ListIndex = 0

End Sub

Private Sub RefreshList(Optional ByVal lngPart As Long)
    '功能：刷新列表
    Dim rsData As ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim ObjItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    On Error GoTo ErrHandle
    
    'Select
    gstrSql = "Select a.Id, a.是否停用, a.编号, a.类别, a.名称, a.说明, Decode(a.接入方式, 1, 'URL', 2, 'EXE', 3, 'FTP', 4, 'ZLBH',a.接入方式) As 接入方式 "

    'From
    gstrSql = gstrSql & " From 三方调用目录 A "
    
    'Where
    If Not mbln显示停用 Then
        gstrSql = gstrSql & " Where a.是否停用 = 0 "
    End If
    
    'Order By
    gstrSql = gstrSql & " Order By a.编号 "
    
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    '列表加载数据
    '---------------------------------------
    rptList.Records.DeleteAll
    
    Do While Not rsData.EOF
        Set objRecord = rptList.Records.Add()
        
        Set ObjItem = objRecord.AddItem(Val(rsData!ID))
        Set ObjItem = objRecord.AddItem(Val(rsData!是否停用))
        
        Set ObjItem = objRecord.AddItem(String(5 - Len(CStr(rsData!编号)), " ") & CStr(rsData!编号))    '空格补起位数，按数字大小顺序排列
        ObjItem.ForeColor = IIf(Val(rsData!是否停用) = 0, vbBlack, vbRed)
        
        Set ObjItem = objRecord.AddItem(CStr(rsData!类别))
        ObjItem.ForeColor = IIf(Val(rsData!是否停用) = 0, vbBlack, vbRed)
        
        Set ObjItem = objRecord.AddItem(CStr(rsData!名称))
        ObjItem.ForeColor = IIf(Val(rsData!是否停用) = 0, vbBlack, vbRed)
        
        Set ObjItem = objRecord.AddItem(CStr(rsData!接入方式))
        ObjItem.ForeColor = IIf(Val(rsData!是否停用) = 0, vbBlack, vbRed)
        
        Set ObjItem = objRecord.AddItem(CStr("" & rsData!说明))
        ObjItem.ForeColor = IIf(Val(rsData!是否停用) = 0, vbBlack, vbRed)
        
        rsData.MoveNext
    Loop
    
    rptList.Populate
    '---------------------------------------
    
    If lngPart <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(COL_ID).Value) = lngPart Then
                    Set Me.rptList.FocusedRow = rptRow
                    Exit For
                End If
            End If
        Next
    End If
    
    If Me.rptList.FocusedRow Is Nothing And Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    
    '状态栏显示
    '---------------------------------------
    Me.stbThis.Panels(2).Text = "共有" & rsData.RecordCount & "项接口"
    '---------------------------------------
    
    Call rptList_SelectionChanged
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub opt接入方式_Click(Index As Integer)
    Call DynamicArrange
End Sub

Private Sub picList_Resize()
    err = 0: On Error Resume Next
    With Me.rptList
        .Left = Me.picList.ScaleLeft: .Width = Me.picList.ScaleWidth - .Left
        .Top = Me.picList.ScaleTop: .Height = Me.picList.ScaleHeight - .Top
    End With
End Sub

Private Sub picEdit_Resize()
    err = 0: On Error Resume Next
    
    '编辑固定部分
    '-------------------------------------
    With fra接口基本信息
        .Top = 50
        .Left = 100
        .Width = picEdit.Width - .Left - 100
    End With
    
    With fra接入方式
        .Top = fra接口基本信息.Top + fra接口基本信息.Height + 100
        .Left = fra接口基本信息.Left
        .Width = picEdit.Width - .Left - 100
    End With
    '-------------------------------------
    
    '动态排列
    Call DynamicArrange
End Sub

Private Sub DynamicArrange()
    '功能：根据选择的接入类型动态排列控件
    
    err = 0: On Error Resume Next
    
    '动态排列
    '-------------------------------------
    '***URL
    With fra配置信息URL
        .Visible = (opt接入方式(mTnterfaceType.URL).Value)
        .Top = fra接入方式.Top + fra接入方式.Height + 100
        .Left = fra接入方式.Left
        .Width = picEdit.Width - .Left - 100
    End With
    
    '***EXE
    With fra配置信息EXE
        .Visible = (opt接入方式(mTnterfaceType.EXE).Value)
        .Top = fra接入方式.Top + fra接入方式.Height + 100
        .Left = fra接入方式.Left
        .Width = picEdit.Width - .Left - 100
    End With
    
    '***FTP
    With fra配置信息FTP
        .Visible = (opt接入方式(mTnterfaceType.FTP).Value)
        .Top = fra接入方式.Top + fra接入方式.Height + 100
        .Left = fra接入方式.Left
        .Width = picEdit.Width - .Left - 100
    End With
    
    '***ZLBH
    With fra配置信息ZLBH
        .Visible = (opt接入方式(mTnterfaceType.ZLBH).Value)
        .Top = fra接入方式.Top + fra接入方式.Height + 100
        .Left = fra接入方式.Left
        .Width = picEdit.Width - .Left - 100
    End With
    
    '-------------------------------------

    '结尾动态部分
    '-------------------------------------
    With lbl提示信息
        .Left = fra接入方式.Left
        .Visible = True
        If opt接入方式(mTnterfaceType.URL).Value Then
            '--URL
            .Top = fra配置信息URL.Top + fra配置信息URL.Height + 50
            .Caption = "说明：URL地址中的“参数”用[1]，[2]，[3]等表示。如：http://192.168.1.4:8055/All/ResultDetail.aspx?MOD=UIS&&ID=[1]，检查号（2016073074）作参数进入传参。" & vbCrLf & _
                    "系统固定传入参数有：病人ID，就诊ID，部门ID，医嘱ID，登录用户名，操作员编号，操作员姓名。"
        ElseIf opt接入方式(mTnterfaceType.EXE).Value Then
            '--EXE
            .Top = fra配置信息EXE.Top + fra配置信息EXE.Height + 50
            .Caption = "说明：EXE程序调用的“参数”用[1]，[2]，[3]等表示。如：c\appsoft\zlhis+.exe[USER]/[口令]/[连接串]。" & vbCrLf & _
                    "系统固定传入参数有：病人ID，就诊ID，部门ID，医嘱ID，登录用户名，操作员编号，操作员姓名。"
        ElseIf opt接入方式(mTnterfaceType.FTP).Value Then
            '--FTP
            .Top = fra配置信息FTP.Top + fra配置信息FTP.Height + 50
            .Caption = "说明：文件名以参数[1]的方式设置。系统固定传入参数有：病人ID，就诊ID，部门ID，医嘱ID，登录用户名，操作员编号，操作员姓名等，除此外程序还提供从数据源提取的方式来组合传入参数值。"
        ElseIf opt接入方式(mTnterfaceType.ZLBH).Value Then
            .Visible = False
        End If
        
        .ToolTipText = .Caption
    End With
    
    '在动态计算参数列表高度前需确定剩下空间的高度
    lbl列表说明.Caption = "说明：参数值为数据源获取时，SQL涉及参数固定传入[病人ID]，[就诊ID]，[部门ID]，[医嘱ID]" & vbCrLf & _
                "就诊ID：门诊病人传参值为就诊ID；住院病人传参值为主页ID"
    
    If ((picEdit.Height - stbThis.Height) - (lbl提示信息.Top + lbl提示信息.Height) - lbl图标提示.Height - fra应用场景.Height - lbl列表说明.Height - 350 > vsfList(mTnterfaceType.URL).RowHeightMin * 7) Or (opt接入方式(mTnterfaceType.ZLBH).Value) Then
        '顺序排列
        
        With vsfList(mTnterfaceType.URL)
            .Visible = (opt接入方式(mTnterfaceType.URL).Value)
            .Top = lbl提示信息.Top + lbl提示信息.Height + 100
            .Left = fra接入方式.Left
            .Width = picEdit.Width - .Left - 100
            .Height = vsfList(mTnterfaceType.URL).RowHeightMin * 7
        End With
        
        With vsfList(mTnterfaceType.EXE)
            .Visible = (opt接入方式(mTnterfaceType.EXE).Value)
            .Top = lbl提示信息.Top + lbl提示信息.Height + 100
            .Left = fra接入方式.Left
            .Width = picEdit.Width - .Left - 100
            .Height = vsfList(mTnterfaceType.URL).RowHeightMin * 7
        End With
        
        With vsfList(mTnterfaceType.FTP)
            .Visible = (opt接入方式(mTnterfaceType.FTP).Value)
            .Top = lbl提示信息.Top + lbl提示信息.Height + 100
            .Left = fra接入方式.Left
            .Width = picEdit.Width - .Left - 100
            .Height = vsfList(mTnterfaceType.URL).RowHeightMin * 7
        End With
        
        With lbl列表说明
            .Visible = Not (opt接入方式(mTnterfaceType.ZLBH).Value)
            .Top = vsfList(mTnterfaceType.URL).Top + vsfList(mTnterfaceType.URL).Height + 100
            .Left = fra接入方式.Left
            .ToolTipText = .Caption
        End With
        
        With fra应用场景
            If (opt接入方式(mTnterfaceType.ZLBH).Value) Then
                .Top = fra配置信息ZLBH.Top + fra配置信息ZLBH.Height + 100
            Else
                .Top = lbl列表说明.Top + lbl列表说明.Height + 100
            End If
            
            .Left = fra接入方式.Left
            .Width = picEdit.Width - .Left - 100
        End With
        
        With lbl图标提示
            .Top = fra应用场景.Top + fra应用场景.Height + 50
            .Left = fra接入方式.Left
            .ToolTipText = .Caption
        End With
    Else
        '以下从下到上反推top
        With lbl图标提示
            .Top = picEdit.Height - stbThis.Height - .Height - 100
            .Left = fra接入方式.Left
            .ToolTipText = .Caption
        End With
        
        With fra应用场景
            .Top = lbl图标提示.Top - .Height - 50
            .Left = fra接入方式.Left
            .Width = picEdit.Width - .Left - 100
        End With
    
        With lbl列表说明
            .Visible = True
            .Top = fra应用场景.Top - .Height - 100
            .Left = fra接入方式.Left
            .ToolTipText = .Caption
        End With
        
        '表格随窗体扩大而扩大
        With vsfList(mTnterfaceType.URL)
            .Visible = (opt接入方式(mTnterfaceType.URL).Value)
            .Top = lbl提示信息.Top + lbl提示信息.Height + 100
            .Left = fra接入方式.Left
            .Width = picEdit.Width - .Left - 100
            .Height = lbl列表说明.Top - .Top - 50
        End With
        
        With vsfList(mTnterfaceType.EXE)
            .Visible = (opt接入方式(mTnterfaceType.EXE).Value)
            .Top = lbl提示信息.Top + lbl提示信息.Height + 100
            .Left = fra接入方式.Left
            .Width = picEdit.Width - .Left - 100
            .Height = lbl列表说明.Top - .Top - 50
        End With
        
        With vsfList(mTnterfaceType.FTP)
            .Visible = (opt接入方式(mTnterfaceType.FTP).Value)
            .Top = lbl提示信息.Top + lbl提示信息.Height + 100
            .Left = fra接入方式.Left
            .Width = picEdit.Width - .Left - 100
            .Height = lbl列表说明.Top - .Top - 50
        End With
    End If
    '-------------------------------------
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    If Button <> vbRightButton Then Exit Sub
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub RefreshInfo(ByVal lngId As Long)
    '功能：根据id刷新当前显示内容
    Dim rsData As ADODB.Recordset
    Dim str小图标 As String
    Dim str大图标 As String
    Dim i As Integer

    On Error GoTo ErrHandle
    
    If lngId = 0 Then Exit Sub
    
    Call ResetControl
    
    '读取基本数据
    '---------------------------------
    gstrSql = "Select a.编号, a.类别, a.名称, a.说明, a.接入方式, a.浏览器类型, a.应用场合, a.地址, a.Ftp地址, a.Ftp访问目录,a.Ftp用户名, a.Ftp密码, a.Ftp本地目录," & vbNewLine & _
            "       a.Ftp端口, a.Ftp文件名, a.菜单显示, a.工具栏显示, a.右键菜单显示" & vbNewLine & _
            "From 三方调用目录 A" & vbNewLine & _
            "Where a.Id = [1]"

    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngId)
    
    Me.txt编号.Text = rsData!编号

    Call Load接口类别
    For i = 1 To cbo接口类别.ListCount
        If cbo接口类别.List(i - 1) = rsData!类别 Then
            cbo接口类别.ListIndex = i - 1
            Exit For
        End If
    Next

    Me.txt名称.Text = rsData!名称
    Me.txt说明.Text = zlcommfun.NVL(rsData!说明, "")
    
    Me.txtFTP地址.Text = zlcommfun.NVL(rsData!Ftp地址, "")
    Me.txtFTP访问目录.Text = zlcommfun.NVL(rsData!Ftp访问目录, "")
    Me.txtFTP用户名.Text = zlcommfun.NVL(rsData!Ftp用户名, "")
    
    If IsNull(rsData!Ftp密码) Then
        Me.txtFTP密码.Text = ""
    Else
        '解密
        Me.txtFTP密码.Text = zlStr.Sm4DecryptEcb(rsData!Ftp密码)
    End If
    
    Me.txtFTP本地目录.Text = zlcommfun.NVL(rsData!Ftp本地目录, "")
    Me.txtFTP端口.Text = zlcommfun.NVL(rsData!Ftp端口, "")
    Me.txt文件下载名.Text = zlcommfun.NVL(rsData!Ftp文件名, "")
    
    Me.opt接入方式(Val(rsData!接入方式) - 1).Value = True
    
    If opt接入方式(mTnterfaceType.URL).Value Then
        Me.txtURL地址.Text = zlcommfun.NVL(rsData!地址, "")
    ElseIf opt接入方式(mTnterfaceType.EXE).Value Then
        Me.txt程序访问路径.Text = zlcommfun.NVL(rsData!地址, "")
    ElseIf opt接入方式(mTnterfaceType.ZLBH).Value Then
        Me.txtZLBH地址.Text = zlcommfun.NVL(rsData!地址, "")
    End If
    
    If zlcommfun.NVL(rsData!浏览器类型, "") = "" Then
        Me.opt浏览器类型(mBrowserType.IE).Value = True
    Else
        Me.opt浏览器类型(Val(rsData!浏览器类型) - 1).Value = True
    End If
    Me.chk门诊医生工作站.Value = Mid(rsData!应用场合, 1, 1)
    Me.chk住院医生工作站.Value = Mid(rsData!应用场合, 2, 1)
    Me.chk住院护士工作站.Value = Mid(rsData!应用场合, 3, 1)
    
    cbo菜单.ListIndex = rsData!菜单显示
    cbo工具栏.ListIndex = rsData!工具栏显示
    cbo右键菜单.ListIndex = rsData!右键菜单显示
    '---------------------------------
    
    '读取参数据
    '---------------------------------
    gstrSql = "Select a.序号, a.参数值, a.备注, a.Sqltext From 三方调用参数 A Where 接口id = [1]"
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngId)
    
    If opt接入方式(mTnterfaceType.URL).Value Then
        With vsfList(mTnterfaceType.URL)
            .Rows = 1
            .Redraw = flexRDNone
            
            Do While Not rsData.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, .ColIndex("序号")) = rsData!序号
                .TextMatrix(.Rows - 1, .ColIndex("参数值")) = rsData!参数值
                .TextMatrix(.Rows - 1, .ColIndex("备注")) = zlcommfun.NVL(rsData!备注, "")
                .TextMatrix(.Rows - 1, .ColIndex("数据源")) = zlcommfun.NVL(rsData!Sqltext, "")
    
                rsData.MoveNext
            Loop
            
            .Redraw = flexRDDirect
            
        End With
    End If
    
    If opt接入方式(mTnterfaceType.EXE).Value Then
        With vsfList(mTnterfaceType.EXE)
            .Rows = 1
            .Redraw = flexRDNone
            
            Do While Not rsData.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, .ColIndex("序号")) = rsData!序号
                .TextMatrix(.Rows - 1, .ColIndex("参数值")) = rsData!参数值
                .TextMatrix(.Rows - 1, .ColIndex("备注")) = zlcommfun.NVL(rsData!备注, "")
                .TextMatrix(.Rows - 1, .ColIndex("数据源")) = zlcommfun.NVL(rsData!Sqltext, "")
    
                rsData.MoveNext
            Loop
            
            .Redraw = flexRDDirect
            
        End With
    End If
    
    If opt接入方式(mTnterfaceType.FTP).Value Then
        With vsfList(mTnterfaceType.FTP)
            .Rows = 1
            .Redraw = flexRDNone
            
            Do While Not rsData.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, .ColIndex("序号")) = rsData!序号
                .TextMatrix(.Rows - 1, .ColIndex("参数值")) = rsData!参数值
                .TextMatrix(.Rows - 1, .ColIndex("备注")) = zlcommfun.NVL(rsData!备注, "")
                .TextMatrix(.Rows - 1, .ColIndex("数据源")) = zlcommfun.NVL(rsData!Sqltext, "")
    
                rsData.MoveNext
            Loop
            
            .Redraw = flexRDDirect
            
        End With
    End If
    '---------------------------------
    
    '读取图标数据
    '---------------------------------
    str小图标 = sys.Readlob(100, 31, lngId)
    str大图标 = sys.Readlob(100, 32, lngId)
    
    img小图标.Picture = LoadPicture(str小图标)
    img小图标.Tag = str小图标
    
    img大图标.Picture = LoadPicture(str大图标)
    img大图标.Tag = str大图标
    
    cmd清空小图标.Visible = (str小图标 <> "")
    cmd清空大图标.Visible = (str大图标 <> "")
    '---------------------------------
    
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub rptList_SelectionChanged()
    Dim lngId As Long
    
    With rptList
        If .FocusedRow Is Nothing Then
            lngId = 0
        ElseIf .FocusedRow.GroupRow = True Then
            lngId = 0
        Else
            lngId = .FocusedRow.Record.Item(mREPORT_COLUMN.COL_ID).Value
        End If
        Call RefreshInfo(lngId)
    End With
End Sub

Private Sub AutoLoading(ByVal objText As TextBox, ByVal intIndex As Integer)
    '功能：根据URL地址或程序访问路径的参数，自动加载对应参数序号到表格
    '说明：参数以"[1]、[2]、....[n]"的形式存在
    'objText：文本框对象
    'intIndex：表格索引
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim str序号 As String
    Dim strTemp As String
    Dim strExistPars As String
    Dim intCount As Integer
    Dim i As Integer
    Dim n As Integer

    '统计"["的个数
    intCount = (Len(objText.Text) - Len(Replace(objText.Text, "[", ""))) / Len("[")
    
    lngStart = 1
    lngEnd = 1
    
    For i = 1 To intCount
        lngStart = InStr(lngEnd, objText.Text, "[")
        If lngStart = 0 Then Exit For
        
        lngEnd = InStr(lngStart, objText.Text, "]")
        If lngEnd = 0 Then Exit For
        
        If lngStart + 1 < lngEnd Then       '[lngStart + 1 < lngEnd]表示"[?]"中间至少含有一个字符
            str序号 = Mid(objText.Text, lngStart + 1, lngEnd - lngStart - 1)
            strTemp = str序号
            
            '验证是否为纯数字
            '---------------------------
            strTemp = strTemp & vbCr
            For n = 0 To 9
                strTemp = Replace(strTemp, CStr(n), "")
            Next
            
            If strTemp = vbCr Then
                '收集参数串
                strExistPars = strExistPars & IIf(strExistPars = "", "", ",") & str序号
            
                With vsfList(intIndex)
                    If .FindRow(str序号, , .ColIndex("序号")) < 1 Then
                        '新增参数
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, .ColIndex("序号")) = str序号
                    End If
                End With
            End If
            '---------------------------
        End If
    Next
    
    '删除不存在的参数
    With vsfList(intIndex)
        If strExistPars = "" Then
            '无参数时清空表格
            .Rows = 1
        Else
            '有参数，需要做比较后做删除
            strTemp = ""
            
            For i = 1 To .Rows - 1
                If InStr("," & strExistPars & ",", "," & .TextMatrix(i, .ColIndex("序号")) & ",") < 1 Then
                    '收集待删除的参数串
                    strTemp = strTemp & IIf(strTemp = "", "", ",") & .TextMatrix(i, .ColIndex("序号"))
                End If
            Next
                             
            For i = 0 To UBound(Split(strTemp, ","))
                .RemoveItem .FindRow(Split(strTemp, ",")(i), , .ColIndex("序号"))
            Next
        End If
    End With
    
End Sub

Private Sub txtFTP地址_GotFocus()
    zlControl.TxtSelAll txtFTP地址
End Sub

Private Sub txtFTP地址_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789.", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtFTP端口_GotFocus()
    zlControl.TxtSelAll txtFTP端口
End Sub

Private Sub txtFTP端口_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtFTP访问目录_GotFocus()
    zlControl.TxtSelAll txtFTP访问目录
End Sub

Private Sub txtFTP访问目录_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtFTP用户名_GotFocus()
    zlControl.TxtSelAll txtFTP用户名
End Sub

Private Sub txtFTP用户名_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'-+_!@#$%^&*(){}[];:,.<>?/|\、《》，。｛｝【】；：？、￥%……&（）", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtFTP密码_GotFocus()
    zlControl.TxtSelAll txtFTP密码
End Sub

Private Sub txtFTP密码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtURL地址_Change()
    Call AutoLoading(txtURL地址, 0)
End Sub

Private Sub txtURL地址_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtZLBH地址_GotFocus()
    zlControl.TxtSelAll txtZLBH地址
End Sub

Private Sub txtZLBH地址_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt编号_GotFocus()
    zlControl.TxtSelAll txt编号
End Sub

Private Sub txt编号_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub

    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt程序访问路径_Change()
    Call AutoLoading(txt程序访问路径, 1)
End Sub

Private Sub txt程序访问路径_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt名称_GotFocus()
    zlControl.TxtSelAll txt名称
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_GotFocus()
    zlControl.TxtSelAll txt说明
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If InStr(1, "'-+_!@#$%^&*(){}[];:<>?/|\、《》｛｝【】；：？、￥%……&（）", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt文件下载名_Change()
    Call AutoLoading(txt文件下载名, 2)
End Sub

Private Sub txt文件下载名_GotFocus()
    zlControl.TxtSelAll txt文件下载名
End Sub

Private Sub txt文件下载名_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtFTP本地目录_GotFocus()
    zlControl.TxtSelAll txtFTP本地目录
End Sub

Private Sub txtFTP本地目录_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsfList_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim str数据源 As String

    With vsfList(Index)
        If Col = .ColIndex("参数值") Then
            str数据源 = .TextMatrix(Row, .ColIndex("数据源"))
            
            '当参数值不是“数据源提取”时，需要清空数据源字段中的数据。
            .TextMatrix(Row, .ColIndex("数据源")) = ""

            Select Case .TextMatrix(Row, Col)
            Case "数据源提取"
                .TextMatrix(.Row, .ColIndex("数据源")) = frmExternalAllocationData.ShowMe(Me, str数据源)
            End Select
        End If
    End With
End Sub

Private Sub vsfList_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfList(Index)
        If Col <> .ColIndex("序号") Then
            Cancel = False
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfList_ChangeEdit(Index As Integer)
    With vsfList(Index)
        If .Col = .ColIndex("参数值") Then
            .Col = .ColIndex("备注")
        End If
    End With
End Sub

Private Sub vsfList_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(1, "'-+_!@#$%^&*(){}[];:<>?/|\、《》｛｝【】；：？、￥%……&（）", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
