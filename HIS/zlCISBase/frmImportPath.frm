VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportPath 
   Caption         =   "导入标准路径"
   ClientHeight    =   10005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11730
   Icon            =   "frmImportPath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   11730
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame frmImport 
      BorderStyle     =   0  'None
      Height          =   9375
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.PictureBox picFont 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   3
         Left            =   7320
         Picture         =   "frmImportPath.frx":6852
         ScaleHeight     =   495
         ScaleWidth      =   2775
         TabIndex        =   81
         Top             =   3480
         Width           =   2775
      End
      Begin VB.PictureBox picFont 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   2
         Left            =   7320
         Picture         =   "frmImportPath.frx":7983
         ScaleHeight     =   495
         ScaleWidth      =   2295
         TabIndex        =   80
         Top             =   2760
         Width           =   2295
      End
      Begin VB.PictureBox picFont 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   1
         Left            =   7320
         Picture         =   "frmImportPath.frx":86B7
         ScaleHeight     =   495
         ScaleWidth      =   2775
         TabIndex        =   79
         Top             =   2040
         Width           =   2775
      End
      Begin VB.PictureBox picFont 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   7320
         Picture         =   "frmImportPath.frx":97DA
         ScaleHeight     =   495
         ScaleWidth      =   3135
         TabIndex        =   78
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Frame fraRuleDefine 
         Caption         =   "文章结构定义"
         Height          =   3090
         Left            =   1080
         TabIndex        =   49
         Top             =   1080
         Width           =   9690
         Begin VB.CommandButton cmdSize 
            Caption         =   "选择"
            Height          =   350
            Index           =   3
            Left            =   5400
            TabIndex        =   65
            Top             =   2520
            Width           =   550
         End
         Begin VB.CommandButton cmdSize 
            Caption         =   "选择"
            Height          =   350
            Index           =   2
            Left            =   5400
            TabIndex        =   64
            Top             =   1680
            Width           =   550
         End
         Begin VB.CommandButton cmdSize 
            Caption         =   "选择"
            Height          =   350
            Index           =   1
            Left            =   5400
            TabIndex        =   63
            Top             =   1035
            Width           =   550
         End
         Begin VB.CommandButton cmdSize 
            Caption         =   "选择"
            Height          =   350
            Index           =   0
            Left            =   5400
            TabIndex        =   62
            Top             =   275
            Width           =   550
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "加粗"
            Height          =   255
            Index           =   3
            Left            =   4560
            TabIndex        =   61
            Tag             =   "0"
            Top             =   2550
            Width           =   735
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "加粗"
            Height          =   255
            Index           =   2
            Left            =   4560
            TabIndex        =   60
            Tag             =   "1"
            Top             =   1710
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "加粗"
            Height          =   255
            Index           =   1
            Left            =   4560
            TabIndex        =   59
            Tag             =   "0"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "加粗"
            Height          =   255
            Index           =   0
            Left            =   4560
            TabIndex        =   58
            Tag             =   "1"
            Top             =   323
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   7
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   57
            Tag             =   "12"
            Text            =   "12"
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   6
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   56
            Tag             =   "宋体"
            Text            =   "宋体"
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   5
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   55
            Tag             =   "14"
            Text            =   "14"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   4
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   54
            Tag             =   "宋体"
            Text            =   "宋体"
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   3
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   53
            Tag             =   "16"
            Text            =   "16"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   2
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   52
            Tag             =   "宋体"
            Text            =   "宋体"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   1
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   51
            Tag             =   "18"
            Text            =   "18"
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   0
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   50
            Tag             =   "宋体"
            Text            =   "宋体"
            Top             =   300
            Width           =   1095
         End
         Begin VB.Line Line4 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   9720
            Y1              =   3045
            Y2              =   3045
         End
         Begin VB.Line Line3 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   9720
            Y1              =   2295
            Y2              =   2295
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   9720
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   9720
            Y1              =   1590
            Y2              =   1590
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "大小"
            Height          =   180
            Index           =   10
            Left            =   3240
            TabIndex        =   77
            Top             =   2580
            Width           =   360
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "字体名称"
            Height          =   180
            Index           =   9
            Left            =   1200
            TabIndex        =   76
            Top             =   2580
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "正文"
            Height          =   180
            Index           =   3
            Left            =   480
            TabIndex        =   75
            Top             =   2580
            Width           =   360
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "大小"
            Height          =   180
            Index           =   8
            Left            =   3240
            TabIndex        =   74
            Top             =   1740
            Width           =   360
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "字体名称"
            Height          =   180
            Index           =   5
            Left            =   1200
            TabIndex        =   73
            Top             =   1740
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "三级标题"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   72
            Top             =   1740
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "大小"
            Height          =   180
            Index           =   7
            Left            =   3240
            TabIndex        =   71
            Top             =   1140
            Width           =   360
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "字体名称"
            Height          =   180
            Index           =   6
            Left            =   1200
            TabIndex        =   70
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "二级标题"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   69
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "大小"
            Height          =   180
            Index           =   4
            Left            =   3240
            TabIndex        =   68
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "字体名称"
            Height          =   180
            Index           =   13
            Left            =   1200
            TabIndex        =   67
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "大标题"
            Height          =   180
            Index           =   0
            Left            =   300
            TabIndex        =   66
            Top             =   360
            Width           =   540
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPathTable 
         Height          =   2205
         Left            =   1080
         TabIndex        =   17
         Top             =   6720
         Width           =   10305
         _cx             =   1963869953
         _cy             =   1963855674
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   20
         RowHeightMax    =   5000
         ColWidthMin     =   100
         ColWidthMax     =   12000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmImportPath.frx":B05C
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
      Begin VB.Label lblSet 
         Caption         =   "标准路径流程的导入格式如下："
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   27
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "说明5：大标题是指路径名称，二级标题是指流程名称或表单名称，三级标题是指路径流程的项目名称，正文是指路径流程内容。"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   19
         Left            =   1080
         TabIndex        =   26
         Top             =   5520
         Width           =   10575
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "说明4：版本信息默认位于路径标题的下一段并包含关键字""版""。"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   18
         Left            =   1080
         TabIndex        =   25
         Top             =   5160
         Width           =   5175
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "说明2：以上四组定义中,不能存在两组完全相同的字体参数。"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   17
         Left            =   1080
         TabIndex        =   24
         Top             =   4530
         Width           =   5055
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "说明3：路径标题必须包含""临床路径""关键字，表单必须包含""临床路径表单""关键词。"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   16
         Left            =   1080
         TabIndex        =   23
         Top             =   4860
         Width           =   6735
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "说明1：文章结构的解析以字体做主要标识,然后以关键字匹配。"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   20
         Left            =   1080
         TabIndex        =   22
         Top             =   4200
         Width           =   6255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSet 
         Caption         =   "提示：一个单元格里不能出现多余的回车符、一个单元格里不能出现多列或多行"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   21
         Top             =   9120
         Width           =   7935
      End
      Begin VB.Label lblSet 
         Caption         =   "标准路径表单的导入格式如下："
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   20
         Top             =   6360
         Width           =   4095
      End
      Begin VB.Label lblTitle 
         Caption         =   "第二步 导入数据规则与格式说明"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   19
         Tag             =   "第二步 导入数据规则与格式说明"
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "说明6：正文中如果存在其他表，则要将表的格式设置成表单个格式，表单的格式如下"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Index           =   25
         Left            =   1080
         TabIndex        =   18
         Top             =   5880
         Width           =   9495
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frmImport 
      BorderStyle     =   0  'None
      Height          =   9495
      Index           =   4
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Frame fraProcess 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   480
         TabIndex        =   42
         Top             =   4200
         Width           =   10695
         Begin MSComctlLib.ProgressBar prgImp 
            Height          =   360
            Index           =   1
            Left            =   1080
            TabIndex        =   43
            Top             =   630
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSComctlLib.ProgressBar prgImp 
            Height          =   360
            Index           =   0
            Left            =   1080
            TabIndex        =   44
            Top             =   150
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblPrg 
            AutoSize        =   -1  'True
            Caption         =   "当前进度"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   46
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblPrg 
            AutoSize        =   -1  'True
            Caption         =   "总进度"
            Height          =   180
            Index           =   1
            Left            =   420
            TabIndex        =   45
            Top             =   720
            Width           =   540
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsErrInfo 
         Height          =   6615
         Left            =   600
         TabIndex        =   47
         Top             =   960
         Visible         =   0   'False
         Width           =   10455
         _cx             =   2004371465
         _cy             =   2004364692
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
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmImportPath.frx":B0C8
         ScrollTrack     =   -1  'True
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
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "第五步 导入"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   4
         Left            =   480
         TabIndex        =   48
         Tag             =   "第五步 导入"
         Top             =   360
         Width           =   1665
      End
   End
   Begin VB.Frame frmImport 
      BorderStyle     =   0  'None
      Height          =   9495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.CommandButton cmdBraw 
         Caption         =   "浏览"
         Height          =   350
         Left            =   9480
         TabIndex        =   9
         Top             =   960
         Width           =   1110
      End
      Begin VB.TextBox txtFile 
         Height          =   270
         Left            =   2880
         TabIndex        =   8
         Top             =   1005
         Width           =   6495
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "文件夹"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   7
         Top             =   1005
         Width           =   855
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "文件"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   6
         Top             =   1005
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Frame fraFloder 
         Caption         =   "浏览文件夹"
         Height          =   5025
         Left            =   2880
         TabIndex        =   1
         Top             =   1320
         Visible         =   0   'False
         Width           =   6495
         Begin VB.DriveListBox div 
            Height          =   300
            Left            =   120
            TabIndex        =   5
            Top             =   276
            Width           =   6165
         End
         Begin VB.DirListBox dirFloder 
            Height          =   3870
            Left            =   96
            TabIndex        =   4
            Top             =   600
            Width           =   6165
         End
         Begin VB.CommandButton cmdPathOk 
            Caption         =   "确定(&O)"
            Height          =   350
            Left            =   3840
            TabIndex        =   3
            Top             =   4560
            Width           =   1100
         End
         Begin VB.CommandButton cmdPathCancel 
            Caption         =   "取消(&C)"
            Height          =   350
            Left            =   5160
            TabIndex        =   2
            Top             =   4560
            Width           =   1100
         End
      End
      Begin MSComDlg.CommonDialog dlgCom 
         Left            =   120
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "注2：导入多个文件时，时间可能会比较长，请耐心等待"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   11
         Left            =   960
         TabIndex        =   82
         Top             =   3960
         Width           =   4935
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         Caption         =   "第一步 选择你要导入的文件或者文件夹"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   12
         Tag             =   "第一步 选择你要导入的文件或者文件夹"
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "说明：文件名若包含符号""-""时将自动解析文件名,格式：""科室名-编码前缀-编码起始值-版本"",分隔符默认为"".""。"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Index           =   21
         Left            =   960
         TabIndex        =   11
         Top             =   4320
         Width           =   7575
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "注1：导入的文件必须为word文档"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   22
         Left            =   960
         TabIndex        =   10
         Top             =   3600
         Width           =   5175
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frmImport 
      BorderStyle     =   0  'None
      Height          =   9495
      Index           =   2
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VSFlex8Ctl.VSFlexGrid vsDefineImp 
         Height          =   6375
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   11055
         _cx             =   1969703980
         _cy             =   1969695725
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
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmImportPath.frx":B151
         ScrollTrack     =   -1  'True
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
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "说明：分割符可为"".""、""-""。编码起始值请输入自然数,导入路径以此为基础依次自增。"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Index           =   23
         Left            =   240
         TabIndex        =   31
         Top             =   8160
         Width           =   6975
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "第三步 设置导入规则与选择导入文件"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   30
         Tag             =   "第三步 设置导入规则与选择导入文件"
         Top             =   360
         Width           =   5460
      End
   End
   Begin VB.Frame frmImport 
      BorderStyle     =   0  'None
      Height          =   9495
      Index           =   3
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Frame fraProcess 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   480
         TabIndex        =   33
         Top             =   4320
         Width           =   10695
         Begin MSComctlLib.ProgressBar prgImp 
            Height          =   360
            Index           =   2
            Left            =   960
            TabIndex        =   34
            Top             =   570
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSComctlLib.ProgressBar prgImp 
            Height          =   360
            Index           =   3
            Left            =   960
            TabIndex        =   35
            Top             =   90
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblPrg 
            AutoSize        =   -1  'True
            Caption         =   "总进度"
            Height          =   180
            Index           =   2
            Left            =   300
            TabIndex        =   37
            Top             =   660
            Width           =   540
         End
         Begin VB.Label lblPrg 
            AutoSize        =   -1  'True
            Caption         =   "当前进度"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   36
            Top             =   180
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAnalyse 
         Height          =   6720
         Left            =   480
         TabIndex        =   38
         Top             =   960
         Width           =   10605
         _cx             =   2004371730
         _cy             =   2004364877
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
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmImportPath.frx":B243
         ScrollTrack     =   -1  'True
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
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "说明：若要部分导入,请选择相应路径进行导入。"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   24
         Left            =   480
         TabIndex        =   40
         Top             =   8040
         Width           =   4335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "第四步 选择解析结果"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   315
         Index           =   3
         Left            =   480
         TabIndex        =   39
         Tag             =   "第四步 选择解析结果"
         Top             =   360
         Width           =   3150
      End
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "下一步(&N)"
      Height          =   350
      Index           =   1
      Left            =   9120
      TabIndex        =   15
      Top             =   9600
      Width           =   1110
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "退出(&X)"
      Height          =   350
      Index           =   2
      Left            =   10320
      TabIndex        =   14
      Top             =   9600
      Width           =   1110
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "上一步(&P)"
      Height          =   350
      Index           =   0
      Left            =   7920
      TabIndex        =   13
      Top             =   9600
      Width           =   1110
   End
End
Attribute VB_Name = "frmImportPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintPage As Integer
Private mblnFile As Boolean
Private mlngSelFileCount As Long '选择的文件总数
Private mlngSelPathCount As Long '选择的路径总数
Private mlngImpFileCount As Long '已经导入的文件总数
Private mlngImpPathCount As Long '已经导入的路径总数

'字体参数串，字体参数组成的字符串：格式为"字体名,字体大小,是否加粗"
Private mstrFontStr大标题 As String
Private mstrFontStr二级标题 As String
Private mstrFontStr小标题 As String
Private mstrFontStr正文 As String
Private mlngStPathID As Long

Private Enum PageEnu
    PE_PathInput = 0 '文件路径输入
    PE_DefineImp = 1 '导入数据规则
    PE_AnaRules = 2 '解析规则
    PE_AnaResult = 3 '解析结果
    PE_ErrInfo = 4 '错误消息显示
End Enum
Private Enum DefineImpCols
    DC_选择 = 0
    DC_文件名称 = 1
    DC_科室名称 = 2
    DC_版本 = 3
    DC_编码前缀 = 4
    DC_分隔符 = 5
    DC_编码起始值 = 6
    DC_文件路径 = 7
End Enum

Private Enum AnaCols
    AC_选择 = 0
    AC_路径名称 = 1
    AC_科室 = 2
    AC_版本 = 3
    ac_编码 = 4
    AC_正文开始 = 5
    AC_正文结束 = 6
    AC_标题开始 = 7
    AC_文件路径 = 8
End Enum

Private Enum ErrCols
    EC_文件名
    EC_路径名称
    EC_错误信息
End Enum

Private Sub cmdBraw_Click()
'功能：选择文件或文件夹
    If optSelect(1).Value = True Then
        fraFloder.Visible = True
        '锁定当前页面
        Call SetContolStat(PE_PathInput, Not fraFloder.Visible)
        cmdImport(0).Enabled = False
        cmdImport(0).Visible = False
    Else
        With dlgCom
            .FileName = ""
            .DialogTitle = "选择文件"
            .FileName = ""
            .Filter = ".docx"
            .ShowOpen
            If .FileName <> "" Then
              txtFile.Text = .FileName
            End If
        End With
    End If
End Sub

Private Sub cmdImport_Click(Index As Integer)
    Dim blnCanNext As Boolean
    Dim intNextPage As Integer
    Dim intPrePage As Integer
    Dim str编码前缀 As String
    Dim str分隔符 As String
    Dim str编码起始值 As String
    Dim i As Long
    Dim strMsg As String
    
    Select Case Index
        Case 1
            If mintPage = 0 Then mblnFile = optSelect(0).Value
            
            blnCanNext = CheckStepNext(mintPage, mblnFile, intNextPage)
            
            If blnCanNext Then
                If mintPage = 2 Then
                    For i = 1 To vsDefineImp.Rows - 1
                        str编码前缀 = Trim(vsDefineImp.TextMatrix(i, DC_编码前缀))
                        str编码起始值 = Trim(vsDefineImp.TextMatrix(i, DC_编码起始值))
                        str分隔符 = Trim(vsDefineImp.TextMatrix(i, DC_分隔符))
                        If Not zlCommFun.IsNumOrChar(str编码前缀) Then
                            strMsg = "第【" & i & "】行编码前缀只能是数字、字母或者字母与数字的组合！"
                        ElseIf Not IsNumeric(str编码起始值) Then
                            strMsg = "第【" & i & "】行编码起始值只能是自然数！"
                        End If
                        If InStr(".-", Trim(str分隔符)) = 0 Then
                            strMsg = "第【" & i & "】行分隔符只能是【.】或者【-】！"
                        Else
                            If Len(Trim(str分隔符)) > 1 Then
                                strMsg = "第【" & i & "】行分隔符只能是是1个字符长度！"
                            End If
                        End If
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName
                            intNextPage = 2
                            frmImport(mintPage).Visible = False
                            frmImport(intNextPage).Visible = True
                            '先锁定，加载数据后解锁
                            Call SetContolStat(intNextPage, False)
                            If intNextPage = PE_PathInput Or intNextPage = PE_ErrInfo Then
                                cmdImport(0).Enabled = False
                                cmdImport(0).Visible = False
                            End If
                            lblTitle(intNextPage).Caption = lblTitle(intNextPage).Tag
                            cmdImport(1).Caption = "下一步(&N)"
                            If intNextPage = PE_ErrInfo Then
                                cmdImport(1).Caption = "返回(&B)"
                            End If
                            mintPage = intNextPage
                            Call SetContolStat(intNextPage, True)
                            Exit Sub
                        End If
                    Next
                End If
                frmImport(mintPage).Visible = False
                frmImport(intNextPage).Visible = True
                '先锁定，加载数据后解锁
                Call SetContolStat(intNextPage, False)
                If intNextPage = PE_PathInput Or intNextPage = PE_ErrInfo Then
                    cmdImport(0).Enabled = False
                    cmdImport(0).Visible = False
                End If
                lblTitle(intNextPage).Caption = lblTitle(intNextPage).Tag
                cmdImport(1).Caption = "下一步(&N)"
                If intNextPage = PE_ErrInfo Then
                    cmdImport(1).Caption = "返回(&B)"
                End If
                mintPage = intNextPage
                '加载下一页面数据,数据加载后解除锁定
                Select Case intNextPage
                    Case PE_PathInput '从最后一页返回
                        Call ClearPage(-1) '清空数据
                            Call SetContolStat(intNextPage, True)
                    Case PE_DefineImp
                        Call SetContolStat(intNextPage, True)
                        With vsPathTable
                            .Rows = 4
                            .Cols = 4
                            .TextMatrix(0, 0) = "时间"
                            .TextMatrix(0, 1) = "住院第一天"
                            .TextMatrix(0, 2) = "住院第二天(手术日)"
                            .TextMatrix(0, 3) = "住院第三天(出院日）"
                            .TextMatrix(1, 0) = "主要诊疗工作"
                            .TextMatrix(1, 1) = "询问病史及查体"
                            .TextMatrix(1, 2) = "完成眼科特殊检查"
                            .TextMatrix(1, 3) = "完成眼科特殊检查"
                            .TextMatrix(2, 0) = "重点医嘱"
                            .TextMatrix(2, 1) = "长期医嘱:眼科三级护理常规"
                            .TextMatrix(2, 2) = "长期医嘱（术后）：眼科二级护理常规"
                            .TextMatrix(2, 3) = "长期医嘱（术后）：眼科二级护理常规"
                            .TextMatrix(3, 0) = "主要护理工作"
                            .TextMatrix(3, 1) = "执行医嘱、生命体征监测"
                            .TextMatrix(3, 2) = "执行医嘱、生命体征监测"
                            .TextMatrix(3, 3) = "执行医嘱、生命体征监测"
                        End With
                        Call SetVsStyle
                Case PE_AnaRules
                    Call LoadFileList(txtFile.Text, mblnFile)
                    Call SetContolStat(intNextPage, True)
                Case PE_AnaResult
                    If Not LoadAnalyseResult Then
                        intNextPage = 2
                        frmImport(mintPage).Visible = False
                        frmImport(intNextPage).Visible = True
                        '先锁定，加载数据后解锁
                        Call SetContolStat(intNextPage, False)
                        If intNextPage = PE_PathInput Or intNextPage = PE_ErrInfo Then
                            cmdImport(0).Enabled = False
                            cmdImport(0).Visible = False
                        End If
                        lblTitle(intNextPage).Caption = lblTitle(intNextPage).Tag
                        cmdImport(1).Caption = "下一步(&N)"
                        If intNextPage = PE_ErrInfo Then
                            cmdImport(1).Caption = "返回(&B)"
                        End If
                        mintPage = intNextPage
                    End If
                    Call SetContolStat(intNextPage, True)
                Case PE_ErrInfo
                    If LoadPath Then
                        Call SetContolStat(intNextPage, True)
                        MsgBox "录入完毕", vbInformation + vbOKOnly, Me.Caption
                    Else
                        Call SetContolStat(intNextPage, True)
                        MsgBox "录入失败", vbInformation + vbOKOnly, Me.Caption
                    End If
                End Select
            End If
        Case 0
            intPrePage = GetStepPre(mintPage)
            frmImport(mintPage).Visible = False
            frmImport(intPrePage).Visible = True
            cmdImport(0).Enabled = True
            cmdImport(0).Visible = True
            lblTitle(intPrePage).Caption = lblTitle(intPrePage).Tag
            cmdImport(1).Caption = "下一步(&N)"
            If mintPage = PE_PathInput Then
                cmdImport(0).Enabled = False
                cmdImport(0).Visible = False
            End If
        
            '清除以前界面数据
            Select Case intPrePage
            Case PE_PathInput
                 Call ClearPage(-1)
            Case PE_DefineImp
                Call ClearPage(PE_AnaRules)
                Call ClearPage(PE_AnaResult)
                Call ClearPage(PE_ErrInfo)
            Case PE_AnaRules
                Call ClearPage(PE_AnaResult)
                Call ClearPage(PE_ErrInfo)
            Case PE_AnaResult
                Call ClearPage(PE_ErrInfo)
            End Select
            mintPage = intPrePage
        Case 2
            '功能：退出
            Unload Me
    End Select
End Sub

Private Function CheckStepNext(ByVal intPage As Integer, ByVal blnFile As Boolean, ByRef intNextPage As Integer) As Boolean
'功能：进行当前步骤的检查，看是否能进行下一步操作
'      intPage :当前可见页面的index
'      blnfile  :输入的是文件类型
'      intNextPage :下一页面的Index
    Dim fileTemp As File, flrTemp As Folder, objfso As New FileSystemObject
    Dim strPath As String, strTest As String, strTmp As String
    Dim blnCanNext As Boolean, blnReback As Boolean
    Dim i As Long, lngRowCount As Long

    strPath = Trim(txtFile.Text)
    mlngSelFileCount = 0
    mlngSelPathCount = 0
    Select Case intPage
            Case PE_PathInput
                If blnFile Then
                    If objfso.FileExists(strPath) Then '文件存在
                        Set fileTemp = objfso.GetFile(strPath)
                        If (UCase(Right(fileTemp.Name, 4)) = ".DOC" Or UCase(Right(fileTemp.Name, 5)) = ".DOCX") And Mid(fileTemp.Name, 1, 2) <> "~$" Then
                            blnCanNext = True
                        Else
                            MsgBox "你输入的文件类型不是可以导入的Word文件,请重新输入", vbInformation, "系统消息"
                        End If
                    Else
                        MsgBox "你输入的文件不存在,请重新输入", vbInformation, "系统消息"
                        Call txtFile.SetFocus
                    End If
                Else
                    If objfso.FolderExists(strPath) Then  '文件夹存在
                        Set flrTemp = objfso.GetFolder(strPath)
                        For Each fileTemp In flrTemp.Files
                            If (UCase(Right(fileTemp.Name, 4)) = ".DOC" Or UCase(Right(fileTemp.Name, 5)) = ".DOCX") And Mid(fileTemp.Name, 1, 2) <> "~$" Then
                                blnCanNext = True
                                Exit For
                            End If
                        Next
                        If Not blnCanNext Then MsgBox "文件夹不存在可以导入的Word文件,请重新输入", vbInformation, "系统消息"
                    Else
                        MsgBox "你输入的文件夹不存在,请重新输入", vbInformation, "系统消息"
                        Call txtFile.SetFocus
                    End If
                End If
            Case PE_DefineImp
                For i = chkBold.LBound To chkBold.UBound
                    If Trim(txtInfo(i * 2).Text) = "" Then
                        MsgBox "请输入字体名称", vbInformation, "系统消息"
                        txtInfo(i).SetFocus
                        Exit Function
                    End If
                    
                    strTmp = txtInfo(i * 2).Text & "," & txtInfo(i * 2 + 1).Text & "," & IIf(chkBold(i).Value = 1, 1, 0)
                    strTest = strTest & "|" & strTmp
                    lblInfo(i).Tag = strTmp
                Next
                
                For i = chkBold.LBound To chkBold.UBound
                    If HaveMoreStr(strTest, lblInfo(i).Tag) Then
                        MsgBox "存在两种或多种相同的文档结构定义", vbInformation, "系统消息"
                        Exit Function
                    End If
                Next
                blnCanNext = True
                '字体参数串
                mstrFontStr大标题 = lblInfo(0).Tag
                mstrFontStr二级标题 = lblInfo(1).Tag
                mstrFontStr小标题 = lblInfo(2).Tag
                mstrFontStr正文 = lblInfo(3).Tag
            Case PE_AnaRules
                With vsDefineImp
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(i, DC_科室名称) = "" Then
                            MsgBox "科室名称不能从文件文件内部读取,请手工输入科室名称", vbInformation, "系统消息"
                            vsDefineImp.SetFocus
                            vsDefineImp.Select i, DC_科室名称
                            Exit Function
                        End If

                        If .TextMatrix(i, DC_编码前缀) = "" Then
                            MsgBox "编码不能从文件中读取，请确定编码规则", vbInformation, "系统消息"
                            vsDefineImp.SetFocus
                            vsDefineImp.Select i, DC_编码前缀
                            Exit Function
                        End If

                        If .TextMatrix(i, DC_分隔符) = "" Then
                            MsgBox "编码不能从文件中读取，请确定编码规则", vbInformation, "系统消息"
                            vsDefineImp.SetFocus
                            vsDefineImp.Select i, DC_分隔符
                            Exit Function
                        End If

                        If .TextMatrix(i, DC_编码起始值) = "" Then
                            MsgBox "编码不能从文件中读取，请确定编码规则", vbInformation, "系统消息"
                            vsDefineImp.SetFocus
                            vsDefineImp.Select i, DC_编码起始值
                            Exit Function
                        End If

                        If .TextMatrix(i, DC_选择) = "-1" Then
                            mlngSelFileCount = mlngSelFileCount + 1
                        End If
                    Next i
                    blnCanNext = mlngSelFileCount <> 0
                    If Not blnCanNext Then
                        MsgBox "你尚未选择将要导入的文件", vbInformation, "系统消息"
                    End If
                End With
            Case PE_AnaResult
                With vsAnalyse
                    lngRowCount = 0
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(i, AC_路径名称) <> "" Then
                            If .TextMatrix(i, AC_选择) = "-1" Then
                                mlngSelPathCount = mlngSelPathCount + 1
                            End If
                            lngRowCount = lngRowCount + 1
                        End If
                    Next i
                End With
                blnCanNext = mlngSelPathCount <> 0
                blnReback = lngRowCount = 0
                If Not blnCanNext And Not blnReback Then
                    MsgBox "你尚未选择将要导入的路径，请选择路径进行导入", vbInformation, "系统消息"
                    Exit Function
                End If
            Case PE_ErrInfo
                blnCanNext = True
    End Select
    '确定下一页的页号
    If intPage = PE_ErrInfo Then
        intNextPage = PE_PathInput
    Else
        If intPage = PE_AnaResult And blnReback Then '没有解析出路径,返回首页
            MsgBox "解析不到任何路径，请查看输入文件名是否正确，解析跳跃段数是否正确", vbInformation, "系统消息"
            intNextPage = PE_PathInput
        Else
            intNextPage = intPage + 1
        End If
    End If
    CheckStepNext = blnCanNext
End Function

Private Sub LoadFileList(ByVal strPath As String, ByVal blnFile As Boolean)
'功能：加载文件列表
'   strPath:文件夹路径或文件路径
'   blnFile:以文件方式导入
    Dim fileTemp As File, flrTemp As Folder, objfso As New FileSystemObject
    Dim arrTmp As Variant
    Dim i As Long, lngRow  As Long
    Dim strTem As String
    Dim strFileFullName As String, str文件名 As String
    
    vsDefineImp.Rows = vsDefineImp.FixedRows
    If blnFile Then '单个文件加载
        Set fileTemp = objfso.GetFile(strPath)
        str文件名 = fileTemp.Name
        strFileFullName = fileTemp.Path
        Call AddNewFileRow(strFileFullName, str文件名)
    Else
        Set flrTemp = objfso.GetFolder(strPath)
        For Each fileTemp In flrTemp.Files
            If (UCase(Right(fileTemp.Name, 4)) = ".DOC" Or UCase(Right(fileTemp.Name, 5)) = ".DOCX") And Mid(fileTemp.Name, 1, 2) <> "~$" Then '非备份非隐藏文件
                str文件名 = fileTemp.Name
                strFileFullName = fileTemp.Path
                Call AddNewFileRow(strFileFullName, str文件名)
                str文件名 = ""
                strFileFullName = ""
            End If
        Next
    End If
End Sub

Private Sub AddNewFileRow(ByVal strFileFullName As String, ByVal str文件名 As String)
'功能:添加新的一行文件信息
    Dim lngRow As Long, arrTmp As Variant
    With vsDefineImp
        .Rows = .Rows + 1
        lngRow = .Rows - 1
        .Cell(flexcpChecked, lngRow, DC_选择) = True
        .TextMatrix(lngRow, DC_文件名称) = str文件名
        .TextMatrix(lngRow, DC_文件路径) = strFileFullName
        If InStr(str文件名, "-") > 0 Then
            arrTmp = Split(str文件名, "-")
            If UBound(arrTmp) >= 1 Then
                .TextMatrix(lngRow, DC_科室名称) = Trim(arrTmp(0))
                .TextMatrix(lngRow, DC_编码前缀) = Trim(arrTmp(1))
            End If
            
            If UBound(arrTmp) >= 2 Then
                .TextMatrix(lngRow, DC_编码起始值) = Trim(arrTmp(2))
            End If
            
            If UBound(arrTmp) >= 3 Then
                .TextMatrix(lngRow, DC_版本) = Trim(arrTmp(3))
            End If
        End If
        .TextMatrix(lngRow, DC_分隔符) = IIf(.colData(DC_分隔符) = ".", .colData(DC_分隔符), ".")
        .TextMatrix(lngRow, DC_科室名称) = IIf(.TextMatrix(lngRow, DC_科室名称) = "", .colData(DC_科室名称), .TextMatrix(lngRow, DC_科室名称))
        .TextMatrix(lngRow, DC_编码前缀) = IIf(.TextMatrix(lngRow, DC_编码前缀) = "", .colData(DC_编码前缀), .TextMatrix(lngRow, DC_编码前缀))
        .TextMatrix(lngRow, DC_编码起始值) = IIf(.TextMatrix(lngRow, DC_编码起始值) = "", .colData(DC_编码起始值), .TextMatrix(lngRow, DC_编码起始值))
        .TextMatrix(lngRow, DC_版本) = IIf(.TextMatrix(lngRow, DC_版本) = "", .colData(DC_版本), .TextMatrix(lngRow, DC_版本))
        
    End With
End Sub

Private Function LoadAnalyseResult() As Boolean

'功能：加载解析结果
    Dim i As Long, lngCount As Long
    vsAnalyse.Visible = False
    fraProcess(1).Visible = True
    With vsDefineImp
        prgImp(2).Max = mlngSelFileCount
        prgImp(2).Value = 0
        '清空数据
        vsAnalyse.Rows = .FixedRows
        For i = .FixedRows To .Rows - 1
            prgImp(2).Value = lngCount
            If .TextMatrix(i, DC_选择) = "-1" Then
                lngCount = lngCount + 1
                If AnalyseDoc(.TextMatrix(i, DC_文件路径), .TextMatrix(i, DC_文件名称), .TextMatrix(i, DC_科室名称), .TextMatrix(i, DC_版本), .TextMatrix(i, DC_编码前缀) & .TextMatrix(i, DC_分隔符), Val(.TextMatrix(i, DC_编码起始值))) = False Then
                    LoadAnalyseResult = False
                    Exit Function
                End If
            End If
        Next
        If vsAnalyse.Rows = .FixedRows Then
            vsAnalyse.Rows = vsAnalyse.Rows + 1
            LoadAnalyseResult = False
            Exit Function
        Else
            LoadAnalyseResult = True
        End If
    End With
    vsAnalyse.Visible = True
    fraProcess(1).Visible = False
End Function

Private Function AnalyseDoc(ByVal strFilePath As String, ByVal str文件名称 As String, ByVal str科室 As String, ByVal str版本In As String, ByVal str编码 As String, ByVal lng编码起始值 As Long) As Boolean
'功能：分析文档,将分析结果反应在表格上
'      str科室:导入设置的科室名称
'      str版本In:导入设置的版本
'      str编码:导入设置的编码前缀&分隔符
'      lng编码起始值:导入设置的编码起始值

    Dim objWord As Object, objWordApp As Object
    Dim rngFind As Object, rng版本 As Object
    Dim i As Long, j As Long, BlnFind As Boolean, lngRow As Long
    Dim str标准路径Name As String, str版本 As String, str字体参数串 As String
    Dim lngParCount As Long
    Dim blnFont As Boolean, blnNameKey As Boolean
    
    On Error GoTo errH
    
    Set objWordApp = CreateObject("Word.Application")
    If objWordApp Is Nothing Then
        MsgBox "Word.Application创建失败！"
        Exit Function
    End If
    '是否能代开word文件
    Set objWord = objWordApp.Documents.Open(strFilePath, False, True, , , , , , , , , False)
    If objWord Is Nothing Then
        MsgBox "文件" & strFilePath & "打开不成功,可能路径名称有误", vbInformation, "系统消息"
        Exit Function
    End If

    With vsAnalyse
    
        lngParCount = objWord.Paragraphs.Count
        If lngParCount = 0 Then Exit Function
        
        prgImp(3).Max = lngParCount
        prgImp(3).Value = 0
        i = 1
        
        Do
            str标准路径Name = ""
            str版本 = ""
            '判断是否找到标题
            Set rngFind = objWord.Paragraphs(i).Range
            str字体参数串 = rngFind.Font.Name & "," & rngFind.Font.Size & "," & IIf(rngFind.Font.Bold = -1, 1, 0)
            If str字体参数串 = mstrFontStr大标题 Then
                blnFont = True
                str标准路径Name = rngFind.Text
                If InStr(str标准路径Name, "临床路径") > 0 Then
                    blnNameKey = True
                    str标准路径Name = Trim(Replace(Replace(Replace(str标准路径Name, " ", ""), Chr(13), ""), Chr(12), ""))
                    str版本 = objWord.Paragraphs(i + 1).Range.Text
                    If InStr(str版本, "版") > 0 Then
                        str版本 = Trim(Replace(Replace(Replace(str版本, "（", ""), "）", ""), Chr(13), ""))
                    Else
                        str版本 = ""
                    End If
                    BlnFind = True
                Else
                    str标准路径Name = ""
                End If
            End If
            '找到后插入数据
            If BlnFind Then
                lng编码起始值 = lng编码起始值 + 1
                .Rows = .Rows + 1
                lngRow = .Rows - 1
                .TextMatrix(lngRow, AC_路径名称) = str标准路径Name
                .TextMatrix(lngRow, AC_版本) = IIf(str版本 = "", str版本In, str版本)
                .TextMatrix(lngRow, AC_正文开始) = i + 2
                If .TextMatrix(lngRow - 1, AC_正文结束) = "" And lngRow <> .FixedRows Then '上一行已经有值就不进行赋值
                    .TextMatrix(lngRow - 1, AC_正文结束) = i - 1
                End If
                .TextMatrix(lngRow, AC_标题开始) = i
                .TextMatrix(lngRow, ac_编码) = str编码 & lng编码起始值
                .TextMatrix(lngRow, AC_科室) = str科室
                .TextMatrix(lngRow, AC_文件路径) = strFilePath
                BlnFind = False
                prgImp(3).Value = i
            End If
            i = i + 1
        Loop While i < lngParCount
        
        .TextMatrix(lngRow, AC_正文结束) = lngParCount  '修改最后一行的数据
        If str标准路径Name = "" Then
            If blnFont Then
                If Not blnNameKey Then
                    If MsgBox("导入【" & str文件名称 & "】文档的大标题不含【临床路径】关键字，是否继续？)", vbYesNo, gstrSysName) = vbNo Then
                        AnalyseDoc = False
                        Exit Function
                    End If
                End If
            Else
                If MsgBox("导入【" & str文件名称 & "】文档的大标题字体与设置的【" & mstrFontStr大标题 & "】不一致,是否继续？", vbYesNo, gstrSysName) = vbNo Then
                    AnalyseDoc = False
                    Exit Function
                End If
            End If
       End If
    End With
    
    Set objWord = Nothing
    Call objWordApp.Quit
    Set objWordApp = Nothing
    blnFont = False
    blnNameKey = False
    AnalyseDoc = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, "系统消息"
    If 0 = 1 Then
        Resume
    End If
    err.Clear
End Function

Private Sub ClearPage(ByVal intPage As Integer)
'功能：清除指定页面的数据
'       intPage:当前可见页面的index,-1时，清除除解析规则与路径输入界面外的数据
    With vsDefineImp
        '保存第一行的导入规则
        .colData(DC_科室名称) = .TextMatrix(.FixedRows, DC_科室名称)
        .colData(DC_编码前缀) = .TextMatrix(.FixedRows, DC_编码前缀)
        .colData(DC_分隔符) = .TextMatrix(.FixedRows, DC_分隔符)
        .colData(DC_编码起始值) = .TextMatrix(.FixedRows, DC_编码起始值)
        If intPage = PE_DefineImp Or intPage = -1 Then
            '清除数据
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
        End If
    End With
    
    With vsAnalyse
        If intPage = PE_AnaResult Or intPage = -1 Then
            '清除数据
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
        End If
    End With
    
    With vsErrInfo
        If intPage = PE_ErrInfo Or intPage = -1 Then
            '清除数据
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
        End If
    End With
End Sub

Private Function GetStepPre(ByVal intPage As Integer) As Integer
'功能：根据当前步骤获取上一页面
'      intPage :当前可见页面的index
'返回 :上一页面的Index
    If intPage = PE_PathInput Then
        GetStepPre = intPage
    Else
        GetStepPre = intPage - 1
    End If
End Function

Private Sub cmdPathCancel_Click()
    fraFloder.Visible = False
    '解锁当前页面
    Call SetContolStat(PE_PathInput, Not fraFloder.Visible)
    cmdImport(0).Enabled = False
    cmdImport(0).Visible = False
End Sub

Private Sub cmdPathOk_Click()
    txtFile.Text = dirFloder.List(dirFloder.ListIndex)
    fraFloder.Visible = False
    '解锁当前页面
    Call SetContolStat(PE_PathInput, Not fraFloder.Visible)
    cmdImport(0).Enabled = False
    cmdImport(0).Visible = False
End Sub


Private Sub cmdSize_Click(Index As Integer)
    With dlgCom
        .Flags = &H80000 + &H100000
        .ShowFont
        If .FontSize <> 0 Then
            txtInfo(Index * 2 + 1).Text = .FontSize
        End If
        
        If .FontName <> "" Then
            txtInfo(Index * 2).Text = .FontName
        End If
        
        If .FontBold Then
            chkBold(Index).Value = 1
        End If
    End With
End Sub

Private Sub div_Change()
    dirFloder.Path = div.Drive
End Sub

Private Sub Form_Load()
    If mintPage = 0 Then cmdImport(0).Visible = False
End Sub
Private Sub SetVsStyle()
'功能：根据内容设置表单表格的单元格的高度与宽度,以及内容颜色等，以及单元格的合并等

    Dim i As Long, j As Long
    Dim lngmaxHeight As Long
   On Error GoTo errH
    With vsPathTable
        If .Rows = 0 And .Cols = 0 Then Exit Sub
        '修改分类名称，阶段，分类加粗居中
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = 4 '居中
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, 0) = &HE1FFE1
        
        .AutoResize = False
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1, False, 0) '自动调整大小
        '设置阶段字体，颜色，对齐方式
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) = "时间" Then
                .Cell(flexcpAlignment, i, 0, i, .Cols - 1) = 4
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False '设置加粗前要先清除加粗
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = True
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HE1FFE1
            Else
                If .Cols > 1 Then
                    .Cell(flexcpAlignment, i, 1, i, .Cols - 1) = 0
                End If
            End If
        Next
        
        '获取同一行最高的单元格高度赋值给行高
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
                For j = 0 To .Cols - 1
                    If j = 0 Then
                        lngmaxHeight = ComputerLines(.TextMatrix(i, j))
                    Else
                        lngmaxHeight = IIf(lngmaxHeight > ComputerLines(.TextMatrix(i, j)), lngmaxHeight, ComputerLines(.TextMatrix(i, j)))
                    End If
                Next
                .RowHeight(i) = IIf(lngmaxHeight = 0, 5, lngmaxHeight) * Me.TextHeight("字") * 1.5
            Else
                For j = 0 To .Cols - 1
                    .TextMatrix(i, j) = " " '为了合并单元格
                Next
            End If
        Next
        '分割行单元格合并，以及边框颜色设置
        .MergeCells = flexMergeFree
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) = " " Then
                Call .CellBorderRange(i, 0, i, .Cols - 1, &HFFFFFF, 1, 0, 1, 0, 1, 0)
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HFFFFFF
                .MergeRow(i) = True
            End If
        Next
        '实现自由拖动列宽
        .FixedRows = 1
        Call .CellBorderRange(0, 0, 0, .Cols - 1, &H8000&, 0, 0, 1, 1, 1, 1)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetContolStat(ByVal intPage As Integer, ByVal blnInput As Boolean)
'功能:设置指定页面控件状态
'   intPage:设置界面
    Dim i As Long
    
    cmdImport(1).Enabled = blnInput
    cmdImport(2).Enabled = blnInput
    cmdImport(0).Enabled = blnInput
    cmdImport(0).Visible = True
    Select Case intPage
        Case PE_PathInput
            optSelect(0).Enabled = blnInput
            optSelect(1).Enabled = blnInput
            txtFile.Enabled = blnInput
            cmdBraw.Enabled = blnInput
        Case PE_DefineImp
            vsPathTable.Enabled = blnInput
        Case PE_AnaRules
            vsDefineImp.Enabled = blnInput
        Case PE_AnaResult
            vsAnalyse.Enabled = blnInput
    End Select
End Sub

Private Function LoadPath() As Boolean
'功能：加载路径，将路径导入数据库

    Dim i As Long, lngCount As Long
    Dim strFilePath As String
    
    vsErrInfo.Visible = False
    vsErrInfo.Rows = vsErrInfo.FixedRows
    mlngImpFileCount = 0
    mlngImpPathCount = 0
    fraProcess(0).Visible = True
    With vsAnalyse
        prgImp(1).Max = mlngSelPathCount
        prgImp(1).Value = 0
        For i = .FixedRows To .Rows - 1
            prgImp(1).Value = mlngImpPathCount
            If .TextMatrix(i, DC_选择) = "-1" Then
               If ImpSelPathByFile(.TextMatrix(i, AC_文件路径), .TextMatrix(i, DC_科室名称), .TextMatrix(i, DC_版本), .TextMatrix(i, ac_编码), , True, Val(.TextMatrix(i, AC_正文结束)), Val(.TextMatrix(i, AC_标题开始))) Then
                    If strFilePath <> .TextMatrix(i, AC_文件路径) Then
                        strFilePath = .TextMatrix(i, AC_文件路径)
                        mlngImpFileCount = mlngImpFileCount + 1
                    End If
                End If
            End If
        Next
    End With
    If vsErrInfo.Rows <> vsErrInfo.FixedRows Then
        vsErrInfo.Visible = True
        fraProcess(0).Visible = False
        LoadPath = False
    Else
        LoadPath = True
        prgImp(1).Value = prgImp(1).Max
        prgImp(0).Value = 0
    End If
End Function

Private Function ImpSelPathByFile(ByVal strFilePath As String, ByVal str科室 As String, ByVal strVerSionIn As String, ByVal strCode As String, _
            Optional ByVal lngCodeStart As Long, Optional ByVal blnAna As Boolean, Optional ByVal lng正文结束 As Long, Optional ByVal lng标题开始 As Long) As Boolean
'功能：根据选择的文件导入路径
'      strFilePath:文件的全名（带路径)
'      str科室:导入设置的科室名称
'      strVerSionIn:导入设置的版本
'      strCode:导入设置的编码前缀&分隔符
'      lngCodeStart:导入设置的编码起始值
'      blnAna:是否经过解析,经过分析后则按路径导入

    Dim objWord As Object, objWordApp As Object
    Dim rngTitle As Object, rngText As Object, rngTable As Object, rngTotal As Object, rngTableTitle As Object, rngTmp As Object
    Dim i As Long, j As Long, k As Long, m As Long, n As Long, h As Long, l As Long
    Dim lngRows As Long, lngCols As Long, lngParCont As Long, lngCurRow As Long
    Dim strStPathName As String, strVerSion As String, strCodeCur As String, strFontStr As String
    Dim lngPathMark As Long, lngCoursNo As Long, lngStPathID As Long, lng阶段序号 As Long, lng分类序号 As Long
    Dim strTableTitle As String, str表单名称 As String, str段落内容 As String, strCoursContent As String, strDiseaseCodes As String, strOpeCode As String
    Dim str阶段名称 As String, str分类名称 As String, strTableContent As String
    Dim strSql As String, rsTmp As New ADODB.Recordset
    Dim strTCD As String, strTmp As String   '中医疾病编码
    Dim arrSql As Variant
    
    On Error GoTo errH
    
    Set objWordApp = CreateObject("Word.Application")
    If objWordApp Is Nothing Then
        MsgBox "Word.Application创建失败！"
        Exit Function
    End If
    '是否能代开word文件
    Set objWord = objWordApp.Documents.Open(strFilePath, False, True, , , , , , , , , False)
    If objWord Is Nothing Then
        Call err.Raise(200000, "文件打开不成功", "文件" & strFilePath & "打开不成功,可能路径名称有误")
    End If
    
    If objWord.Paragraphs.Count = 0 Then
        Call err.Raise(200000, "文件没有包含内容", "文件" & objWord.Name & "不包含任何内容")
    End If
    Set rngTotal = objWord.Paragraphs(1).Range
    If blnAna Then
        Call rngTotal.SetRange(objWord.Paragraphs(lng标题开始).Range.Start, objWord.Paragraphs(lng正文结束).Range.End)
    Else
        Call rngTotal.SetRange(0, objWord.Paragraphs(objWord.Paragraphs.Count).Range.End)
    End If
    lngParCont = rngTotal.Paragraphs.Count
    If lngParCont = 0 Then
        Call err.Raise(200001, "路径不包含信息", objWord.Name & "中所选路径不包含信息")
    End If
    
    prgImp(0).Max = lngParCont
    prgImp(0).Value = 0
    i = 1
    arrSql = Array()
    Do
        '判断是否找到标题
        Set rngTitle = rngTotal.Paragraphs(i).Range
        prgImp(0).Value = i
        If Trim(Replace(Replace(Replace(rngTitle.Text, " ", ""), Chr(13), ""), Chr(12), "")) <> "" Then
            strFontStr = rngTitle.Font.Name & "," & rngTitle.Font.Size & "," & IIf(rngTitle.Font.Bold = -1, 1, 0)
            If strFontStr = mstrFontStr大标题 Then
                    If InStr(rngTitle.Text, "临床路径") > 0 Then
                        mlngImpPathCount = mlngImpPathCount + 1
                        '初始变量
                        strVerSion = ""
                        strStPathName = ""
                        strVerSion = ""
                        lngStPathID = 0
                        strCodeCur = ""
                        strSql = ""
                        '获取路径总体信息
                        If Not blnAna Then
                            strCodeCur = strCode & lngCodeStart
                            lngCodeStart = lngCodeStart + 1
                        Else
                            strCodeCur = strCode
                        End If
                        strStPathName = Trim(Replace(Replace(Replace(rngTitle.Text, " ", ""), Chr(13), ""), Chr(12), ""))
                        For j = i + 1 To i + 10
                            If j <= rngTotal.Paragraphs.Count Then
                                Set rngText = rngTotal.Paragraphs(j).Range
                                strFontStr = rngText.Font.Name & "," & rngText.Font.Size & "," & IIf(rngText.Font.Bold = -1, 1, 0)
                                If strFontStr = mstrFontStr二级标题 Then i = j - 1: Exit For
                                If InStr(rngText.Text, "版") > 0 Then
                                    strVerSion = Trim(Replace(Replace(Replace(strVerSion, "（", ""), "）", ""), Chr(13), ""))
                                    i = j
                                    Exit For
                                End If
                            End If
                        Next
                        strVerSion = IIf(strVerSion = "", strVerSionIn, strVerSion)
                        
                        strSql = "select ID,科室名称,编码,路径名称 from 标准路径目录 where 科室名称=[1] and 路径名称=[2] and 版本说明=[3] and 编码=[4]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str科室, strStPathName, strVerSion, strCodeCur)
                        
                        If rsTmp.RecordCount = 0 Then
                            strSql = "Zl_标准路径目录_Insert(NULL,'" & str科室 & "','" & strCodeCur & "','" & strStPathName & "','" & strVerSion & "',Null,Null)"
                            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                        Else
                            Call err.Raise(19999, "不能插入该路径", "路径：" & strStPathName & "不能插入,检查【标准路径目录】表中是否存在【科室名称, 编码, 路径名称, 版本说明】相同的数据")
                        End If
                        
                        strSql = "select ID,科室名称,编码,路径名称 from 标准路径目录 where 科室名称=[1] and 路径名称=[2] and 版本说明=[3] and 编码=[4]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str科室, strStPathName, strVerSion, strCodeCur)

                        If rsTmp.RecordCount <> 0 Then
                            rsTmp.MoveFirst
                            lngStPathID = Val(rsTmp!ID & "")
                            mlngStPathID = lngStPathID
                        End If
                        '插其他数据
                        lngPathMark = 1
                        lngCoursNo = 1
                    End If
            ElseIf strFontStr = mstrFontStr二级标题 Or InStr(rngTitle.Text, "临床路径表单") > 0 And InStr(rngTitle.Text, NumberToChar(lngPathMark) & "、") > 0 Then
                    '如果是表单,则路径标记+1，lng路径标记含义为：1：标准路径流程，2：路径表单1，3：路径表单2,....
                    If InStr(rngTitle.Text, "临床路径表单") > 0 Then
                        If InStr(rngTitle.Text, NumberToChar(lngPathMark) & "、") = 0 Then lngPathMark = lngPathMark + 1
                    End If
                    '导入表单
                    If lngPathMark > 1 Then
                        strSql = ""
                        str段落内容 = Trim(Replace(rngTitle.Text, Chr(13), ""))
                        str表单名称 = Mid(str段落内容, InStr(str段落内容, NumberToStr(lngPathMark) & "、") + Len(NumberToStr(lngPathMark) & "、"))
                        str表单名称 = Mid(str表单名称, 1, InStr(str表单名称, "临床路径表单") + 5)
                        '获取该次表单
                        For j = i + 1 To rngTotal.Paragraphs.Count
                            Set rngText = rngTotal.Paragraphs(j).Range
                            strFontStr = rngText.Font.Name & "," & rngText.Font.Size & "," & IIf(rngText.Font.Bold = -1, 1, 0)
                            
                            '下一路径
                            If strFontStr = mstrFontStr大标题 Then
                                If InStr(rngText.Text, "临床路径") > 0 Then
                                    i = j - 1 '将光标到下一个路径开始处
                                    Exit For
                                End If
                            End If
                            '下一表单结束（与上面IF分开，为了提高效率)
                            If strFontStr = mstrFontStr二级标题 Then
                                If InStr(rngText.Text, "临床路径表单") > 0 Then
                                    i = j - 1 '将光标到下一个路径表单开始处
                                    Exit For
                                End If
                            End If
                        Next
                        Set rngTable = objWord.Range(rngTitle.End, rngText.Start)
                        '如果制定范围内存在表，则进行数据解析
                        If rngTable.Tables.Count <> 0 Then
                            Call rngText.SetRange(rngTable.Start, rngTable.Tables(1).Range.Start)
                            strTableTitle = ""
                            For k = 1 To rngText.Paragraphs.Count
                                Set rngTableTitle = rngText.Paragraphs(k).Range
                                If Trim(Replace(Replace(rngTableTitle.Text, Chr(13), ""), Chr(12), "")) <> "" Then
                                    strTableTitle = strTableTitle & rngTableTitle.Text
                                End If
                            Next
                            lng阶段序号 = 1: lng分类序号 = 1
                            '插入表单的表头数据 Zl_标准路径目录_Insert时,已经插入一条表单表头数据
                            ReDim Preserve arrSql(UBound(arrSql) + 1)
                            arrSql(UBound(arrSql)) = "Zl_标准路径表单_Update(" & lngStPathID & "," & IIf(lngPathMark = 2, 1, 0) & ",'" & Trim(str表单名称) & "','" & strTableTitle & "')"
                            
                            '清除默认数据
                            ReDim Preserve arrSql(UBound(arrSql) + 1)
                            arrSql(UBound(arrSql)) = "Zl_标准路径表单_ContentClear(" & lngStPathID & "," & Val(lngPathMark - 1) & ")"
         
                            
                            '读取表格数据
                            For k = 1 To rngTable.Tables.Count
                                lngRows = rngTable.Tables(k).Rows.Count
                                lngCols = rngTable.Tables(k).Columns.Count
                                For m = 1 To rngTable.Tables(k).Columns.Count
                                    lng阶段序号 = lng阶段序号 + 1 '阶段标识，实际阶段序号为lng阶段序号-j(因为每个表的第一列不算作阶段）
                                    '合并单元格读取
                                    str阶段名称 = rngTable.Tables(k).Cell(1, m).Range.Text
                                    str阶段名称 = Trim(Replace(str阶段名称, Chr(13) & Chr(7), ""))
                                    If m <> 1 Then
                                        For n = 2 To lngRows
                                            lng分类序号 = n '分类序号1时，用来存储表单表头
                                            str分类名称 = Trim(Replace(rngTable.Tables(k).Cell(n, 1).Range.Text, Chr(13) & Chr(7), ""))
                                            str分类名称 = Replace(Replace(str分类名称, " ", ""), Chr(13), "")
                                            If InStr(",病情变异记录,责任护士签名,医师签名,", "," & str分类名称 & ",") > 0 Then Exit For
                                            strTableContent = Trim(Replace(rngTable.Tables(k).Cell(n, m).Range.Text, Chr(13) & Chr(7), ""))
                                            If Len(Trim(strTableContent)) > 2000 Then
                                                Call err.Raise(19999, "路径插入不成功", "路径：" & strStPathName & "没有插入成功,检查【" & str分类名称 & "】-【" & str阶段名称 & "】中的内容是否超过了2000个字符长度！")
                                            End If
                                            '插入表单中的路径项目内容
                                            ReDim Preserve arrSql(UBound(arrSql) + 1)
                                            arrSql(UBound(arrSql)) = "Zl_标准路径表单_ContentInsert(" & lngStPathID & "," & Val(lngPathMark - 1) & "," & _
                                                    lng分类序号 & ",'" & str分类名称 & "','" & lng阶段序号 - k & "','" & str阶段名称 & "','" & strTableContent & "')"
                                        Next
                                    End If
                                Next
                            Next
                        End If
                        '标识下一个路径表单的序号
                        lngPathMark = lngPathMark + 1
                    End If
                ElseIf strFontStr = mstrFontStr小标题 And lngPathMark = 1 Then
                    strSql = ""
                    strCoursContent = ""
    '                If lngCoursNo = 7 Then Stop
                    h = i
                    For j = i + 1 To rngTotal.Paragraphs.Count
                        Set rngText = rngTotal.Paragraphs(j).Range
                        strFontStr = rngText.Font.Name & "," & rngText.Font.Size & "," & IIf(rngText.Font.Bold = -1, 1, 0)
                        If strFontStr = lblInfo(2).Tag Or strFontStr = lblInfo(1).Tag Or strFontStr = lblInfo(0).Tag Or _
                            lngCoursNo > 6 And InStr(rngText.Text, "临床路径表单") > 0 And InStr(rngText.Text, NumberToChar(lngPathMark + 1) & "、") > 0 Then
                            i = j - 1 '将光标到下一个路径流程项目开始处
                            Exit For
                        End If
                    Next
                    If h <> i And rngTotal.Paragraphs(h + 1).Range.Start <> rngTotal.Paragraphs(i).Range.End Then
                        If rngTmp Is Nothing Then Set rngTmp = rngTotal.Paragraphs(h + 1).Range
                        Call rngTmp.SetRange(rngTotal.Paragraphs(h + 1).Range.Start, rngTotal.Paragraphs(i).Range.End)
                        If rngTmp.Tables.Count <> 0 Then
                            If rngTotal.Paragraphs(h + 1).Range.Start <> rngTmp.Tables(1).Range.Start Then
                                Call rngText.SetRange(rngTotal.Paragraphs(h + 1).Range.Start, rngTmp.Tables(1).Range.Start)
                                If Trim(Replace(Replace(rngText.Text, Chr(13), ""), Chr(12), "")) <> "" Then '与回车
                                    strCoursContent = strCoursContent & rngText.Text
                                End If
                            End If
                            For j = 1 To rngTmp.Tables.Count
                                For m = 1 To rngTmp.Tables(j).Rows.Count
                                    For n = 1 To rngTmp.Tables(j).Columns.Count
                                        strCoursContent = strCoursContent & "        " & RPAD(rngTmp.Tables(j).Cell(m, n).Range.Text, " ", 15)
                                    Next
                                    strCoursContent = strCoursContent & vbNewLine
                                Next
                                If j < rngTmp.Tables.Count Then
                                    If rngTmp.Tables(j).Range.End <> rngTmp.Tables(j + 1).Range.Start Then
                                        Call rngText.SetRange(rngTmp.Tables(j).Range.End, rngTmp.Tables(j + 1).Range.Start)
                                        If Trim(Replace(Replace(rngText.Text, Chr(13), ""), Chr(12), "")) <> "" Then '与回车
                                            strCoursContent = strCoursContent & rngText.Text
                                        End If
                                    End If
                                Else
                                    If rngTmp.Tables(rngTmp.Tables.Count).Range.End <> rngTmp.End Then
                                        Call rngText.SetRange(rngTmp.Tables(rngTmp.Tables.Count).Range.End, rngTmp.End)
                                        If Trim(Replace(Replace(rngText.Text, Chr(13), ""), Chr(12), "")) <> "" Then '与回车
                                            strCoursContent = strCoursContent & rngText.Text
                                        End If
                                    End If
                                End If
                            Next
                            
                        Else
                            For j = 1 To rngTmp.Paragraphs.Count
                                Set rngText = rngTmp.Paragraphs(j).Range
                                If Trim(Replace(Replace(rngText.Text, Chr(13), ""), Chr(12), "")) <> "" Then '与回车
                                    strCoursContent = strCoursContent & rngText.Text
                                End If
                            Next
                        End If
                    End If
                    '对开始的适用对象做特殊处理，为了获取疾病编码
                    str段落内容 = Trim(Replace(rngTitle.Text, Chr(13), ""))
                    If Not (str段落内容 = "" And strCoursContent = "") Then
'                        If lngCoursNo = 11 Then Stop
                        If InStr(str段落内容, "适用对象") > 0 And lngCoursNo = 1 Or InStr(str段落内容, "进入路径标准") > 0 And lngCoursNo = 2 And (strDiseaseCodes = "" Or strTCD = "" Or strOpeCode = "") Then
                            'strTCD 中医疾病编码
                            strTCD = "" '为解析导入时,下一临床路径需要清空
                            If InStr(strCoursContent, "TCD") > 0 Then
                                strTmp = Replace(Replace(Mid(strCoursContent, InStr(strCoursContent, "TCD")), "）", ")"), "：", ":")
                                For l = 1 To UBound(Split(strCoursContent, "TCD"))
                                    strTmp = Mid(strTmp, InStr(strTmp, ":") + 1)
                                    If Len(strTmp) > 0 And InStr(strTmp, ")") > 0 Then
                                        strTCD = strTCD & "," & Mid(strTmp, 1, InStr(strTmp, ")") - 1)
                                    End If
                                Next
                                strTCD = Mid(strTCD, 2)
                            Else
                                strTCD = ""
                            End If
                            '获取疾病编码与手术编码
                            strDiseaseCodes = ""
                            If InStr(strCoursContent, "ICD-10") > 0 Then
                                strTmp = Replace(Replace(Mid(strCoursContent, InStr(strCoursContent, "ICD-10")), "）", ")"), "：", ":")
                                For l = 1 To UBound(Split(strCoursContent, "ICD-10"))
                                    strTmp = Mid(strTmp, InStr(strTmp, ":") + 1)
                                     If Len(strTmp) > 0 And InStr(strTmp, ")") > 0 Then
                                        strDiseaseCodes = strDiseaseCodes & "," & Mid(strTmp, 1, InStr(strTmp, ")") - 1)
                                    End If
                                Next
                                strDiseaseCodes = Mid(strDiseaseCodes, 2)
                                If Len(strDiseaseCodes) > 200 Then
                                    Call err.Raise(19999, "路径插入不成功", "路径：" & strStPathName & "没有插入成功,检查【" & str段落内容 & "】中的疾病诊断内容是否超过了200个字符长度！")
                                End If
                            Else
                                strDiseaseCodes = ""
                            End If
                            strOpeCode = ""
                            If InStr(strCoursContent, "ICD-9") > 0 Then
                                strOpeCode = Replace(Replace(Mid(strCoursContent, InStr(strCoursContent, "ICD-9")), "）", ")"), "：", ":")
                                strOpeCode = Mid(strOpeCode, InStr(strOpeCode, ":") + 1)
                                If Len(strOpeCode) > 0 And InStr(strOpeCode, ")") > 0 Then
                                    strOpeCode = Mid(strOpeCode, 1, InStr(strOpeCode, ")") - 1)
                                    If Len(strOpeCode) > 100 Then
                                        Call err.Raise(19999, "路径插入不成功", "路径：" & strStPathName & "没有插入成功,检查【" & str段落内容 & "】中的手术内诊断容是否超过了100个字符长度！")
                                    End If
                                End If
                            Else
                                strOpeCode = ""
                            End If
                            ReDim Preserve arrSql(UBound(arrSql) + 1)
                            arrSql(UBound(arrSql)) = "Zl_标准路径病种_Update(" & lngStPathID & ",'" & IIf(strTCD = "", "", strTCD & IIf(strDiseaseCodes = "", "", ",")) & strDiseaseCodes & "','" & strOpeCode & "')"
                        End If
                        If Len(Trim(strCoursContent)) > 4000 Then
                            Call err.Raise(19999, "路径插入不成功", "路径：" & strStPathName & "没有插入成功,检查【" & str段落内容 & "】的内容是否超过了4000个字符长度！")
                        End If
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = "Zl_标准路径流程_Insert(" & lngStPathID & "," & lngCoursNo & ",'" & Trim(str段落内容) & "','" & Trim(strCoursContent) & "')"
                        '标识下一项目的序号
                        lngCoursNo = lngCoursNo + 1
                    End If
            End If
        End If
        i = i + 1
    Loop While i + 1 < lngParCont
    
    '批量提交数据
    For l = LBound(arrSql) To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(l)), Me.Caption)
    Next
    
    If vsErrInfo.Rows > 1 Then
        ImpSelPathByFile = False
    Else
        ImpSelPathByFile = True
    End If
    Set objWord = Nothing
    Call objWordApp.Quit(False)
    Set objWordApp = Nothing
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    If err.Number = 5941 Then
        If err.Number <> 0 Then
            err.Description = "路径：" & strStPathName & "没有插入成功,检查【" & Trim(str表单名称) & "】中的单元格内是否存在多列或者多行的情况！"
        End If
    End If
    With vsErrInfo
        .Rows = .Rows + 1
        lngCurRow = .Rows - 1
        If Not objWord Is Nothing Then
            .TextMatrix(lngCurRow, EC_文件名) = objWord.Name
            .TextMatrix(lngCurRow, EC_路径名称) = strStPathName
            .TextMatrix(lngCurRow, EC_错误信息) = err.Description
        Else
            .TextMatrix(lngCurRow, EC_文件名) = objWord.Name
            .TextMatrix(lngCurRow, EC_路径名称) = err.Source
            .TextMatrix(lngCurRow, EC_错误信息) = err.Description
        End If
    End With
    err.Clear
    Set objWord = Nothing
    Call objWordApp.Quit
    Set objWordApp = Nothing
End Function

Private Function ComputerLines(ByVal strInput As String) As Long
'功能：计算输入文本中回车符的个数
'参数：  strInput   要计算回车符的字符串
'返回：   回车符的个数

    Dim strTmp As String
    Dim Count  As Long, lngPos As Long, lngLen As Long
    
    lngPos = InStr(strInput, Chr(13))
    lngLen = Len(strInput)
    strTmp = strInput
    
    Do While lngPos <> 0
        If Trim(strTmp) = "" Then Exit Do
        If lngPos + 1 <= lngLen Then
            strTmp = Mid(strTmp, lngPos + 1)
            Count = Count + 1
            lngPos = InStr(strTmp, Chr(13))
            lngLen = Len(strTmp)
        End If
    Loop
    
    ComputerLines = Count + 2
    
End Function


Private Sub Form_Unload(Cancel As Integer)
    mintPage = 0
    mblnFile = False
    mlngSelFileCount = 0
    mlngSelPathCount = 0
    mlngImpFileCount = 0
    mlngImpPathCount = 0
    mstrFontStr大标题 = ""
    mstrFontStr二级标题 = ""
    mstrFontStr小标题 = ""
    mstrFontStr正文 = ""
    mlngStPathID = 0
End Sub
Private Sub vsAnalyse_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsAnalyse
        If Not (Col = AC_选择 And .TextMatrix(Row, AC_路径名称) <> "" And Row <> 0) Then
            Cancel = True
        End If
    End With
End Sub

Private Function HaveMoreStr(ByVal strSouce As String, ByVal strJudge As String)
'功能：判断strSouce是否存在两个以上的strJudge
    If Len(strSouce) = Len(Replace(strSouce, strJudge, "")) + Len(strJudge) Then
        HaveMoreStr = False
    Else
        HaveMoreStr = True
    End If
End Function

Public Function ShowMe(frmParent As Object, ByRef lngId As Long) As Boolean
    Me.Show 1, frmParent
    lngId = mlngStPathID
    ShowMe = True
End Function

Private Sub vsDefineImp_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = DC_分隔符 Then
        If Len(Trim(vsDefineImp.TextMatrix(Row, Col))) > 1 Then
            MsgBox "第【" & Row & "】行分隔符只能是一位！", vbInformation, gstrSysName
            zlControl.ControlSetFocus vsDefineImp
            Exit Sub
        End If
    End If
End Sub

Private Sub vsErrInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strMsg As String
    If Shift = 2 And KeyCode = vbKeyC Then
        Clipboard.Clear
        Debug.Print vsErrInfo.MouseRow & "-" & vsErrInfo.MouseCol
        If Not vsErrInfo.MouseRow < 0 And Not vsErrInfo.MouseCol < 0 Then
            strMsg = vsErrInfo.TextMatrix(vsErrInfo.MouseRow, vsErrInfo.MouseCol)
            Clipboard.SetText strMsg
        End If
    End If
End Sub
