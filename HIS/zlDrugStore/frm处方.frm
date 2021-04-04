VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm处方 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picRecInfo_CM 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      ScaleHeight     =   255
      ScaleWidth      =   7455
      TabIndex        =   45
      Top             =   7560
      Width           =   7455
      Begin VB.Label lbl原始付数 
         AutoSize        =   -1  'True
         Caption         =   "原始付数："
         Height          =   180
         Left            =   0
         TabIndex        =   47
         Tag             =   "原始付数:"
         Top             =   60
         Width           =   900
      End
      Begin VB.Label lbl中药煎法 
         AutoSize        =   -1  'True
         Caption         =   "中药煎法："
         Height          =   180
         Left            =   1830
         TabIndex        =   46
         Tag             =   "中药煎法:"
         Top             =   60
         Width           =   900
      End
   End
   Begin VB.PictureBox picRecipe 
      BackColor       =   &H00FFFFFF&
      Height          =   7455
      Index           =   0
      Left            =   240
      ScaleHeight     =   7395
      ScaleWidth      =   8475
      TabIndex        =   6
      Top             =   0
      Width           =   8535
      Begin VB.PictureBox picRecipe 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   4
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   8175
         TabIndex        =   30
         Top             =   6480
         Width           =   8175
         Begin VB.Label lblRP后记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "日期："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   2
            Left            =   3720
            TabIndex        =   44
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblRP后记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "2009年07月07日"
            Height          =   180
            Index           =   3
            Left            =   4320
            TabIndex        =   39
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label lblRP后记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "医师："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   8
            Left            =   3720
            TabIndex        =   38
            Top             =   600
            Width           =   540
         End
         Begin VB.Label lblRP后记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "王小二"
            Height          =   180
            Index           =   9
            Left            =   4320
            TabIndex        =   37
            Top             =   600
            Width           =   540
         End
         Begin VB.Line lineRP后记 
            Index           =   0
            X1              =   240
            X2              =   7680
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label lblRP后记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "应收/实收合计："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   0
            Left            =   270
            TabIndex        =   36
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label lblRP后记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "1013.45/1010.00元"
            Height          =   180
            Index           =   1
            Left            =   1680
            TabIndex        =   35
            Top             =   240
            Width           =   1530
         End
         Begin VB.Label lblRP后记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "收费员："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   4
            Left            =   6000
            TabIndex        =   34
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblRP后记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "陈三娃"
            Height          =   180
            Index           =   5
            Left            =   6840
            TabIndex        =   33
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblRP后记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "配药人："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   6
            Left            =   900
            TabIndex        =   32
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lblRP后记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "张建国"
            Height          =   180
            Index           =   7
            Left            =   1680
            TabIndex        =   31
            Top             =   600
            Width           =   540
         End
      End
      Begin VB.PictureBox picRecipe 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Index           =   3
         Left            =   0
         ScaleHeight     =   3975
         ScaleWidth      =   8175
         TabIndex        =   27
         Top             =   2280
         Width           =   8175
         Begin VSFlex8Ctl.VSFlexGrid vsfRecipe 
            Height          =   1815
            Left            =   480
            TabIndex        =   28
            Top             =   600
            Width           =   2175
            _cx             =   3836
            _cy             =   3201
            Appearance      =   0
            BorderStyle     =   0
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
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   2
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
            ComboSearch     =   0
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
         Begin VB.Line lineRP正文 
            Index           =   0
            X1              =   240
            X2              =   7680
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label lblRP正文 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "RP"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   120
            Width           =   390
         End
      End
      Begin VB.PictureBox picRecipe 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   2
         Left            =   0
         ScaleHeight     =   1575
         ScaleWidth      =   8175
         TabIndex        =   11
         Top             =   720
         Width           =   8175
         Begin VB.TextBox txt诊断内容 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   420
            Left            =   1680
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   48
            Text            =   "frm处方.frx":0000
            Top             =   1100
            Width           =   6135
         End
         Begin VB.Label lblRP前记 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "体重："
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   21
            Left            =   3240
            TabIndex        =   54
            Tag             =   "床号："
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblRP前记 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "55kg"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   22
            Left            =   3840
            TabIndex        =   53
            Tag             =   "体重："
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "1234567890"
            Height          =   180
            Index           =   20
            Left            =   6600
            TabIndex        =   52
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "就诊卡号："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   19
            Left            =   5640
            TabIndex        =   51
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "5"
            Height          =   180
            Index           =   17
            Left            =   5040
            TabIndex        =   43
            Top             =   720
            Width           =   90
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "处方号："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   4
            Left            =   5820
            TabIndex        =   42
            Top             =   0
            Width           =   720
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "险类："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   2
            Left            =   3240
            TabIndex        =   41
            Top             =   0
            Width           =   540
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "自费"
            Height          =   180
            Index           =   1
            Left            =   1680
            TabIndex        =   40
            Top             =   0
            Width           =   360
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "费别："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   0
            Left            =   1125
            TabIndex        =   26
            Top             =   0
            Width           =   540
         End
         Begin VB.Line lineRP前记 
            Index           =   0
            X1              =   240
            X2              =   7680
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "医保"
            Height          =   180
            Index           =   3
            Left            =   3840
            TabIndex        =   25
            Top             =   0
            Width           =   360
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "H00040015"
            Height          =   180
            Index           =   5
            Left            =   6600
            TabIndex        =   24
            Top             =   0
            Width           =   825
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "张三娃娃"
            Height          =   180
            Index           =   7
            Left            =   1680
            TabIndex        =   23
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "男"
            Height          =   180
            Index           =   9
            Left            =   3840
            TabIndex        =   22
            Top             =   360
            Width           =   180
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "40"
            Height          =   180
            Index           =   11
            Left            =   5040
            TabIndex        =   21
            Top             =   360
            Width           =   180
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "门诊内科"
            Height          =   180
            Index           =   15
            Left            =   6240
            TabIndex        =   20
            Top             =   720
            Width           =   720
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "姓名："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   6
            Left            =   1125
            TabIndex        =   19
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "性别："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   8
            Left            =   3240
            TabIndex        =   18
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "年龄："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   10
            Left            =   4440
            TabIndex        =   17
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "科室："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   14
            Left            =   5640
            TabIndex        =   16
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "床号："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   16
            Left            =   4440
            TabIndex        =   15
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "标识号："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   12
            Left            =   945
            TabIndex        =   14
            Top             =   720
            Width           =   720
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "3434242123213"
            Height          =   180
            Index           =   13
            Left            =   1680
            TabIndex        =   13
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label lblRP前记 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "临床诊断："
            ForeColor       =   &H00C0C000&
            Height          =   180
            Index           =   18
            Left            =   765
            TabIndex        =   12
            Top             =   1080
            Width           =   900
         End
      End
      Begin VB.PictureBox picRecipe 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   8175
         TabIndex        =   7
         Top             =   0
         Width           =   8175
         Begin VB.PictureBox picRecipe 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   5
            Left            =   6840
            ScaleHeight     =   435
            ScaleWidth      =   945
            TabIndex        =   8
            Top             =   68
            Width           =   975
            Begin VB.Label lblRP标识 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "普通"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   9
               Top             =   75
               Width           =   720
            End
         End
         Begin VB.Label lblRP标题 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "重庆市第三人民医院"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1080
            TabIndex        =   10
            Top             =   120
            Width           =   5055
         End
      End
   End
   Begin VB.PictureBox picProcess 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   9855
      TabIndex        =   0
      Top             =   7920
      Width           =   9855
      Begin VB.ComboBox cbo核查人 
         Height          =   300
         Left            =   6600
         TabIndex        =   49
         Text            =   "cbo核查人"
         Top             =   0
         Width           =   1935
      End
      Begin VB.ComboBox cbo配药人 
         Height          =   300
         Left            =   3720
         TabIndex        =   3
         Text            =   "cbo配药人"
         Top             =   25
         Width           =   1935
      End
      Begin VB.ComboBox cbo开单医生 
         Height          =   300
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   25
         Width           =   1935
      End
      Begin VB.CommandButton CmdSend 
         Caption         =   "发药(&S)"
         Height          =   350
         Left            =   8640
         TabIndex        =   1
         ToolTipText     =   "热键：F2"
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lbl核查人 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "核查人"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6000
         TabIndex        =   50
         Top             =   90
         Width           =   540
      End
      Begin VB.Label lbl配药人 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "配药人"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3120
         TabIndex        =   5
         Top             =   85
         Width           =   540
      End
      Begin VB.Label lbl开单医生 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开单医生"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   85
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm处方"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'用户定义的处方颜色，从注册表取的字符串，用;分隔
Private mstrUserRecipeColor As String

Private mstrDosUser As String
Private mstrPrivs As String
Private mbln中药处方 As Boolean
Private mbln标志 As Boolean
Private mstr核查人 As String

Private Type Type_Condition
    intListType As Integer
    lng药房ID As Long
    bln自动配药 As Boolean
    bln是否需要配药过程 As Boolean
    bln校验处方 As Boolean
    str配药人 As String
    str核查人 As String
    int金额显示 As Integer      '金额显示方式：0-显示应收金额,1-显示实收金额,2-显示应收和实收金额
    bln处方审查 As Boolean
    intRowNum As Integer
    
End Type
Private mcondition As Type_Condition

'列表类型
Private Enum mListType
    配药确认 = 0
    待配药 = 1
    已配药 = 2
    待发药 = 3
    超时未发 = 4
    退药 = 5
End Enum

'处方类型：普通、儿科、急诊、精二、精一、麻醉
Private Enum 处方类型
    普通 = 0
    儿科 = 1
    急诊 = 2
    精二 = 3
    精一 = 4
    麻醉 = 5
End Enum

Private Enum 西药列名
    分组 = 0
    药品 = 1
    皮试结果 = 2
    总量 = 3
    单量 = 4
    用法 = 5
    嘱托 = 6
    相关id = 7
    
    列数 = 8
End Enum

Private Enum 中药列名
    分组1 = 0
    分组1脚注 = 1
    间隔1 = 2
    分组2 = 3
    分组2脚注 = 4
    间隔2 = 5
    分组3 = 6
    分组3脚注 = 7
    间隔3 = 8
    分组4 = 9
    分组4脚注 = 10
    
    列数 = 8
End Enum

Private Enum 处方签区域
    整体 = 0
    标题 = 1
    前记 = 2
    正文 = 3
    后记 = 4
    标识 = 5
End Enum

Private Enum RP标题
    医院名称 = 0
End Enum

Private Enum RP标识
    标识 = 0
End Enum

Private Enum RP前记
    费别标签 = 0
    费别 = 1
    
    险类标签 = 2
    险类 = 3
    
    处方号标签 = 4
    处方号 = 5
    
    姓名标签 = 6
    姓名 = 7
    
    性别标签 = 8
    性别 = 9
    
    年龄标签 = 10
    年龄 = 11
    
    标识号标签 = 12
    标识号 = 13
    
    科室标签 = 14
    科室 = 15
    
    床号标签 = 16
    床号 = 17
    
    临床诊断标签 = 18
    
    就诊卡号标签 = 19
    就诊卡号 = 20
    
    体重标签 = 21
    体重 = 22
End Enum

Private Enum RP正文
    标识 = 0
End Enum

Private Enum RP后记
    合计金额标签 = 0
    合计金额 = 1
    
    日期标签 = 2
    日期 = 3
    
    收费员标签 = 4
    收费员 = 5
    
    配药人标签 = 6
    配药人 = 7
    
    开单医生标签 = 8
    开单医生 = 9
End Enum

Public Sub CmdProcess()
    If CmdSend.Enabled Then CmdSend_Click
End Sub
Public Sub FormClear()
    '标题
    lblRP标题(RP标题.医院名称).Caption = GetUnitName
    
    '标识
    lblRP标识(RP标识.标识).Caption = "普通"
    
    '前记
    lblRP前记(RP前记.费别).Caption = ""
    lblRP前记(RP前记.险类).Caption = ""
    lblRP前记(RP前记.处方号).Caption = ""
    lblRP前记(RP前记.姓名).Caption = ""
    lblRP前记(RP前记.性别).Caption = ""
    lblRP前记(RP前记.年龄).Caption = ""
    lblRP前记(RP前记.就诊卡号).Caption = ""
    lblRP前记(RP前记.标识号).Caption = ""
    lblRP前记(RP前记.科室).Caption = ""
    lblRP前记(RP前记.床号).Caption = ""
    lblRP前记(RP前记.体重).Caption = ""
    txt诊断内容.Text = ""
    txt诊断内容.Tag = ""
    
    '正文
    vsfRecipe.rows = 1
    
    '后记
    lblRP后记(RP后记.合计金额).Caption = ""
    lblRP后记(RP后记.日期).Caption = ""
    lblRP后记(RP后记.收费员).Caption = ""
    lblRP后记(RP后记.配药人).Caption = ""
    lblRP后记(RP后记.开单医生).Caption = ""
    
    '设置颜色
    SetRecipeColor 0
    
    CmdSend.Enabled = False
End Sub


Private Sub Load医生()
    Dim rsData As ADODB.Recordset
    
    Set rsData = RecipeSendWork_Get医生
    
    Me.cbo开单医生.Clear
    cbo开单医生.AddItem ""
    Do While Not rsData.EOF
        cbo开单医生.AddItem rsData!医生
        rsData.MoveNext
    Loop
    cbo开单医生.ListIndex = 0
End Sub
Public Sub SetParams()
    Dim bln是否配药确认 As Boolean

    mstrUserRecipeColor = zldatabase.GetPara("处方颜色", glngSys, 1341)
    If mstrUserRecipeColor = "" Then mstrUserRecipeColor = GetDefaultRecipeColor
    
    With mcondition
        If .lng药房ID <> Val(zldatabase.GetPara("发药药房", glngSys, 1341)) Then
            .lng药房ID = Val(zldatabase.GetPara("发药药房", glngSys, 1341))
            .bln是否需要配药过程 = RecipeSendWork_DispensingMedi(.lng药房ID, bln是否配药确认)
            Call Load配药人(.lng药房ID)
        End If
        
        .str配药人 = zldatabase.GetPara("配药人", glngSys, 1341)
        .str核查人 = zldatabase.GetPara("核查人", glngSys, 1341)
        
        If .str配药人 = "|当前操作员|" Then
            mstrDosUser = gstrUserName
        Else
            mstrDosUser = .str配药人
        End If
        
        If .str核查人 = "|当前操作员|" Then
            mstr核查人 = gstrUserName
        Else
            mstr核查人 = .str核查人
        End If
    
        .bln自动配药 = (Val(zldatabase.GetPara("自动配药", glngSys, 1341)) = 1)
        .int金额显示 = Val(zldatabase.GetPara("金额显示方式", glngSys, 1341, 0))
        .bln处方审查 = ((gtype_UserSysParms.P240_药房处方审查 = 1 Or gtype_UserSysParms.P240_药房处方审查 = 3) And gtype_UserSysParms.P241_处方审查时机 = 2)
        .intRowNum = gtype_UserSysParms.P213_中药配方每行中药味数
        
        
        If zlStr.IsHavePrivs(mstrPrivs, "配药") = True Then
            If .bln自动配药 = False Then
                Cbo配药人.Enabled = True
            Else
                Cbo配药人.Enabled = False
            End If
        Else
            Cbo配药人.Enabled = False
        End If
        
        cbo核查人.Enabled = True
        .bln校验处方 = IsInString(gstrprivs, "校验处方", ";")
        
        Call Load配药人(.lng药房ID)
        
        Call Load核查人(.lng药房ID)
    End With
End Sub

Private Sub SetRecipeMedi(ByVal rsData As ADODB.Recordset)
    Dim intRow As Integer
    Dim n As Integer
    Dim lng上行相关ID As Long
    Dim lng本行相关ID As Long
    Dim lng下行相关ID As Long
    Dim bln皮试 As Boolean
    Dim i As Integer
    Dim lng药名id As Long
    Dim dblAmount As Double
    Dim strDiag As String
    Dim int门诊 As Integer
    Dim intCol As Integer
    Dim dateCurrent As Date
    
    dateCurrent = Sys.Currentdate
    
    rsData.Filter = ""
    rsData.Sort = "相关ID,序号"
    
    With vsfRecipe
        .Redraw = flexRDNone
        
        Do While Not rsData.EOF
            If rsData!记录性质 = 1 Or (rsData!记录性质 = 2 And (rsData!门诊标志 = 1 Or rsData!门诊标志 = 4)) Then
                int门诊 = 1
            Else
                int门诊 = 2
            End If
            strDiag = RecipeSendWork_GetDiagnosis(int门诊, IIf(int门诊 = 1, Val(rsData!相关id), Val(rsData!病人ID)), Val(rsData!主页id), IIf(mbln中药处方, 1, 2))
            If int门诊 = 1 And rsData!在院 And strDiag = "" Then
                int门诊 = 2
                strDiag = RecipeSendWork_GetDiagnosis(int门诊, IIf(int门诊 = 1, Val(rsData!相关id), Val(rsData!病人ID)), Val(rsData!主页id), IIf(mbln中药处方, 1, 2))
            End If
            
            If strDiag <> "" Then
                strDiag = strDiag & "|"
                For i = 0 To UBound(Split(strDiag, "|"))
                    If Split(strDiag, "|")(i) <> "" Then
                        If InStr(1, txt诊断内容.Text & " ※", "※" & Split(strDiag, "|")(i) & " ※") < 1 Then
                            txt诊断内容.Text = IIf(txt诊断内容.Text = "", " ※", txt诊断内容.Text & " ※") & Split(strDiag, "|")(i)
                            txt诊断内容.Tag = IIf(txt诊断内容.Tag = "", "※ ", txt诊断内容.Tag & vbCrLf & "※ ") & Split(strDiag, "|")(i)
                        End If
                    End If
                Next
            End If
        
            If mbln中药处方 Then
                .MergeCells = flexMergeRestrictColumns
                .MergeCol(中药列名.分组1) = True
                .MergeCol(中药列名.分组2) = True
                .MergeCol(中药列名.分组3) = True
                If mcondition.intRowNum = 4 Then .MergeCol(中药列名.分组4) = True

                If rsData!药名ID <> lng药名id Then
                    If intCol = mcondition.intRowNum Then
                        intCol = 0
                        intRow = intRow + 2
                    ElseIf intCol = 0 Then
                        intRow = intRow + 2
                    End If
                    .rows = intRow + 1
                    
                    lng药名id = rsData!药名ID
                    
                    If NVL(rsData!单量, 0) = 0 Then
                        dblAmount = rsData!数量 * rsData!付数 * rsData!包装 * rsData!剂量系数 / NVL(rsData!原始付数, 1)
                    Else
                        dblAmount = rsData!单量
                    End If
                Else
                    intCol = intCol - 1
                    
                    If NVL(rsData!单量, 0) = 0 Then
                        dblAmount = dblAmount + rsData!数量 * rsData!付数 * rsData!包装 * rsData!剂量系数 / NVL(rsData!原始付数, 1)
                    Else
                        dblAmount = rsData!单量
                    End If
                End If
                
                If intCol = 0 Then
                    .TextMatrix(intRow, 中药列名.分组1) = rsData!药名
                    .TextMatrix(intRow, 中药列名.分组1脚注) = FormatEx(Abs(dblAmount), 1) & rsData!计算单位
                    .TextMatrix(intRow - 1, 中药列名.分组1) = rsData!药名
                    .TextMatrix(intRow - 1, 中药列名.分组1脚注) = IIf(IsNull(rsData!医生嘱托), "", "(" & rsData!医生嘱托 & ")")
                ElseIf intCol = 1 Then
                    .TextMatrix(intRow, 中药列名.分组2) = rsData!药名
                    .TextMatrix(intRow, 中药列名.分组2脚注) = FormatEx(Abs(dblAmount), 1) & rsData!计算单位
                    .TextMatrix(intRow - 1, 中药列名.分组2) = rsData!药名
                    .TextMatrix(intRow - 1, 中药列名.分组2脚注) = IIf(IsNull(rsData!医生嘱托), "", "(" & rsData!医生嘱托 & ")")
                ElseIf intCol = 2 Then
                    .TextMatrix(intRow, 中药列名.分组3) = rsData!药名
                    .TextMatrix(intRow, 中药列名.分组3脚注) = FormatEx(Abs(dblAmount), 1) & rsData!计算单位
                    .TextMatrix(intRow - 1, 中药列名.分组3) = rsData!药名
                    .TextMatrix(intRow - 1, 中药列名.分组3脚注) = IIf(IsNull(rsData!医生嘱托), "", "(" & rsData!医生嘱托 & ")")
                ElseIf intCol = 3 Then
                    .TextMatrix(intRow, 中药列名.分组4) = rsData!药名
                    .TextMatrix(intRow, 中药列名.分组4脚注) = FormatEx(Abs(dblAmount), 1) & rsData!计算单位
                    .TextMatrix(intRow - 1, 中药列名.分组4) = rsData!药名
                    .TextMatrix(intRow - 1, 中药列名.分组4脚注) = IIf(IsNull(rsData!医生嘱托), "", "(" & rsData!医生嘱托 & ")")

                End If
        
                intCol = intCol + 1
                
                .RowHeight(intRow) = 250
                .RowHeight(intRow - 1) = 250

            Else
                intRow = intRow + 1
                .rows = intRow + 1
                
                dblAmount = rsData!数量
                
                .TextMatrix(intRow, 西药列名.药品) = rsData!药品名称 & vbCrLf & rsData!药品规格
                
                If rsData!是否皮试 = 1 Then
                    .TextMatrix(intRow, 西药列名.皮试结果) = Get皮试结果(rsData!病人ID, rsData!药名ID, dateCurrent, rsData!开嘱时间)
                    If .TextMatrix(intRow, 西药列名.皮试结果) <> "" Then
                        bln皮试 = True
                    End If
                End If
                
                .TextMatrix(intRow, 西药列名.总量) = zlStr.FormatEx(dblAmount, 5) & rsData!单位
                
                .ColWidth(西药列名.单量) = 750
                .TextMatrix(intRow, 西药列名.单量) = IIf(IsNull(rsData!单量), "", zlStr.FormatEx(rsData!单量, 5) & "(" & zlStr.NVL(rsData!计算单位) & ")")
                
                .TextMatrix(intRow, 西药列名.用法) = IIf(IsNull(rsData!用法), "", rsData!用法) & " " & IIf(IsNull(rsData!频次), "", rsData!频次)
                .TextMatrix(intRow, 西药列名.嘱托) = IIf(IsNull(rsData!医生嘱托), "", rsData!医生嘱托)
                .TextMatrix(intRow, 西药列名.相关id) = Val(rsData!相关id)
                
                '默认名称+规格分两行写，规格写在第二行；如果名称超过行宽则增加额外的行数；默认每行字高250，每字宽200
                .RowHeight(intRow) = 250 * ((-1 * Int(-1 * Len(rsData!药品名称) / Int((.ColWidth(西药列名.药品) / 200)))) + 1)
            End If
            
            rsData.MoveNext
            
            '如果下一个药品不是一组药品，增加一空行
            If Not rsData.EOF And Not mbln中药处方 Then
                If Val(rsData!相关id) = 0 Then
                    intRow = intRow + 1
                    .rows = intRow + 1
                    .RowHeight(intRow) = 200
                End If
            End If
        Loop
        
        If Not mbln中药处方 Then
            '设置分组
            For n = 1 To .rows - 1
                If Val(.TextMatrix(n, 西药列名.相关id)) <> 0 Then
                    lng本行相关ID = .TextMatrix(n, 西药列名.相关id)
                    If n + 1 <= .rows - 1 Then
                        If Val(.TextMatrix(n + 1, 西药列名.相关id)) <> 0 Then    '如果下行为记录行时
                            lng下行相关ID = IIf(.TextMatrix(n + 1, 西药列名.相关id) = 0, -1, .TextMatrix(n + 1, 西药列名.相关id))
                        ElseIf n + 2 <= .rows - 1 Then  '如果下行为汇总行行时
                            If Val(.TextMatrix(n + 2, 西药列名.相关id)) <> 0 Then    '如果下下行为记录行时
                                lng下行相关ID = IIf(Val(.TextMatrix(n + 2, 西药列名.相关id)) = 0, -1, Val(.TextMatrix(n + 2, 西药列名.相关id)))
                            Else
                                lng下行相关ID = -1
                            End If
                        Else
                            lng下行相关ID = -1
                        End If
                    Else
                        lng下行相关ID = -1
                    End If
                    
                    If lng本行相关ID = lng上行相关ID Then
                        If lng本行相关ID = lng下行相关ID Then
                            .TextMatrix(n, 西药列名.分组) = "│"
                        Else
                            .TextMatrix(n, 西药列名.分组) = "└"
                        End If
                    ElseIf lng本行相关ID = lng下行相关ID Then
                        .TextMatrix(n, 西药列名.分组) = "┌"
                    End If
                
                    lng上行相关ID = IIf(lng本行相关ID = 0, -1, lng本行相关ID)
                End If
            Next
            
            .MergeCells = flexMergeRestrictColumns
            .MergeCol(西药列名.用法) = False
            
            '处理皮试
            If bln皮试 = True Then
                .ColWidth(西药列名.皮试结果) = 800
                For i = 1 To .rows - 1
                    If .TextMatrix(i, 西药列名.皮试结果) = "(+)" Then
                        .Cell(flexcpForeColor, i, 西药列名.皮试结果, i, 西药列名.皮试结果) = vbRed
                    ElseIf .TextMatrix(i, 西药列名.皮试结果) = "(-)" Then
                        .Cell(flexcpForeColor, i, 西药列名.皮试结果, i, 西药列名.皮试结果) = vbBlue
                    Else
                        .Cell(flexcpForeColor, i, 西药列名.皮试结果, i, 西药列名.皮试结果) = &H80000008
                    End If
                Next
            Else
                .ColWidth(西药列名.皮试结果) = 0
            End If
        End If
        
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub IniRecipe()
    
    With vsfRecipe
        .rows = 1
        
        '药品（含规格），总量，单量，用法，嘱托，相关ID
        
        If Not mbln中药处方 Then
            .Cols = 西药列名.列数
            .ColWidth(0) = 500
            .ColWidth(西药列名.药品) = 2000
            .ColWidth(西药列名.皮试结果) = 400
            .ColWidth(西药列名.总量) = 750
            
            
            .ColWidth(西药列名.单量) = 750
            .ColWidth(西药列名.用法) = 2000
            .ColWidth(西药列名.嘱托) = 1500
            .ColWidth(西药列名.相关id) = 0
            
            .FixedAlignment(西药列名.药品) = flexAlignCenterCenter
            .FixedAlignment(西药列名.皮试结果) = flexAlignCenterCenter
            .FixedAlignment(西药列名.总量) = flexAlignCenterCenter
            .FixedAlignment(西药列名.单量) = flexAlignCenterCenter
            .FixedAlignment(西药列名.用法) = flexAlignCenterCenter
            .FixedAlignment(西药列名.嘱托) = flexAlignCenterCenter
            
            .TextMatrix(0, 西药列名.药品) = "药品"
            .TextMatrix(0, 西药列名.皮试结果) = ""
            .TextMatrix(0, 西药列名.总量) = "总量"
            .TextMatrix(0, 西药列名.单量) = "单量"
            .TextMatrix(0, 西药列名.用法) = "用法"
            .TextMatrix(0, 西药列名.嘱托) = "嘱托"
            
            .ColAlignment(西药列名.分组) = flexAlignRightCenter
            .ColAlignment(西药列名.药品) = flexAlignLeftCenter
            .ColAlignment(西药列名.皮试结果) = flexAlignLeftCenter
            .ColAlignment(西药列名.总量) = flexAlignCenterCenter
            .ColAlignment(西药列名.单量) = flexAlignCenterCenter
            .ColAlignment(西药列名.用法) = flexAlignLeftCenter
            .ColAlignment(西药列名.嘱托) = flexAlignLeftCenter
            
            .RowHeight(0) = 255
        Else
            .Cols = 中药列名.列数 + IIf(mcondition.intRowNum = 4, 3, 0)
            
            If mcondition.intRowNum = 4 Then
                .ColWidth(中药列名.分组1) = 1100
                .ColWidth(中药列名.分组2) = 1100
                .ColWidth(中药列名.分组3) = 1100
                .ColWidth(中药列名.分组4) = 1100
                .ColWidth(中药列名.间隔1) = 50
                .ColWidth(中药列名.间隔2) = 50
                .ColWidth(中药列名.间隔3) = 50
                .ColWidth(中药列名.分组1脚注) = 750
                .ColWidth(中药列名.分组2脚注) = 750
                .ColWidth(中药列名.分组3脚注) = 750
                .ColWidth(中药列名.分组4脚注) = 750
                
                .ColAlignment(中药列名.分组4) = flexAlignRightCenter
                .TextMatrix(0, 中药列名.分组4) = ""
            Else
                .ColWidth(中药列名.分组1) = 1700
                .ColWidth(中药列名.分组2) = 1700
                .ColWidth(中药列名.分组3) = 1700
                .ColWidth(中药列名.间隔1) = 50
                .ColWidth(中药列名.间隔2) = 50
                .ColWidth(中药列名.分组1脚注) = 750
                .ColWidth(中药列名.分组2脚注) = 750
                .ColWidth(中药列名.分组3脚注) = 750
            End If
            .ColAlignment(中药列名.分组1) = flexAlignRightCenter
            .ColAlignment(中药列名.分组2) = flexAlignRightCenter
            .ColAlignment(中药列名.分组3) = flexAlignRightCenter
            .ColAlignment(中药列名.分组1脚注) = flexAlignLeftCenter
            .ColAlignment(中药列名.分组2脚注) = flexAlignLeftCenter
            .ColAlignment(中药列名.分组3脚注) = flexAlignLeftCenter
            
            .TextMatrix(0, 中药列名.分组1) = ""
            .TextMatrix(0, 中药列名.分组2) = ""
            .TextMatrix(0, 中药列名.分组3) = ""
            .TextMatrix(0, 中药列名.分组1脚注) = ""
            .TextMatrix(0, 中药列名.分组2脚注) = ""
            .TextMatrix(0, 中药列名.分组3脚注) = ""
            
            .RowHidden(0) = 0
        End If
        
    End With
End Sub
Public Sub ShowRecipe(ByVal intType As Integer)
    Dim i As Integer
    
    With mcondition
        .intListType = intType
        
        If .intListType = mListType.待发药 Or .intListType = mListType.超时未发 Then
            cbo开单医生.Enabled = True
        Else
            cbo开单医生.Enabled = False
        End If
        
        If .intListType <> mListType.退药 Then
            For i = 0 To Cbo配药人.ListCount - 1
                If mstrDosUser = Cbo配药人.List(i) Then
                    Cbo配药人.ListIndex = i
                    Exit For
                End If
            Next
            
            For i = 0 To cbo核查人.ListCount - 1
                If mstr核查人 = cbo核查人.List(i) Then
                    cbo核查人.ListIndex = i
                    Exit For
                End If
            Next
            Lbl配药人.Caption = "配药人"
        Else
            Lbl配药人.Caption = "发药人"
        End If
        
'        cbo配药人.Enabled = (.intListType <> mListType.退药)
        If zlStr.IsHavePrivs(mstrPrivs, "配药") = True Then
            If .bln自动配药 = False Then
                Cbo配药人.Enabled = True
            Else
                Cbo配药人.Enabled = False
            End If
        Else
            Cbo配药人.Enabled = False
        End If

        cbo核查人.Enabled = True
        If .intListType = mListType.退药 Then
            Cbo配药人.Enabled = False
            cbo核查人.Enabled = False
        End If
        
        Select Case .intListType
            Case mListType.配药确认
                Me.cbo开单医生.Enabled = False
                Me.Cbo配药人.Enabled = False
                Me.cbo核查人.Enabled = False
                CmdSend.Caption = "配药确认(&O)"
            Case mListType.待配药
                CmdSend.Caption = "配药(&V)"
            Case mListType.已配药
                CmdSend.Caption = "取消配药(&C)"
            Case mListType.待发药, mListType.超时未发
                CmdSend.Caption = "发药(&S)"
        End Select
        
        CmdSend.Visible = (.intListType <> mListType.退药)
    End With
    
    SetCmdSendPrivs intType
End Sub

Private Sub SetCmdSendPrivs(ByVal int审查结果 As Integer)
    '权限控制
    Select Case mcondition.intListType
    Case mListType.配药确认
       '配药确认
        CmdSend.Enabled = zlStr.IsHavePrivs(mstrPrivs, "配药确认")
    Case mListType.待配药
        '配药
        CmdSend.Enabled = (zlStr.IsHavePrivs(mstrPrivs, "配药") And mcondition.bln自动配药 = False And (mcondition.bln处方审查 = False Or (mcondition.bln处方审查 = True And int审查结果 = 1)))
    Case mListType.已配药
        '取消
        CmdSend.Enabled = zlStr.IsHavePrivs(mstrPrivs, "配药")
    Case mListType.待发药, mListType.超时未发
        '发药
        CmdSend.Enabled = (zlStr.IsHavePrivs(mstrPrivs, "发药") And (mcondition.bln处方审查 = False Or (mcondition.bln处方审查 = True And int审查结果 = 1)))
    End Select
End Sub


Private Sub cbo核查人_Click()
    Dim i As Integer
    
    If mcondition.intListType = mListType.待发药 Or mcondition.intListType = mListType.超时未发 Then
        mstr核查人 = Me.cbo核查人.Text
    End If
End Sub

Private Sub cbo配药人_Click()
    Dim i As Integer
    
    If mcondition.intListType = mListType.待发药 Or mcondition.intListType = mListType.超时未发 Then
        mstrDosUser = Me.Cbo配药人.Text
    End If
End Sub

Private Sub CmdSend_Click()
    If frm药品处方发药New.RecipeWork(mcondition.intListType, frm处方发药明细.mblnInput, frm处方发药明细.vsfList) = False Then
        FormClear
    End If
End Sub

Public Function Get配药人() As String
    If Cbo配药人.ListIndex = -1 Then
        Get配药人 = ""
    ElseIf InStr(Cbo配药人.Text, "-") > 0 Then
        Get配药人 = Mid(Cbo配药人.Text, InStr(Cbo配药人.Text, "-") + 1)
    Else
        Get配药人 = Cbo配药人.Text
    End If
End Function
Public Function Get核查人() As String
    If cbo核查人.ListIndex = -1 Then
        Get核查人 = ""
    ElseIf InStr(cbo核查人.Text, "-") > 0 Then
        Get核查人 = Mid(cbo核查人.Text, InStr(cbo核查人.Text, "-") + 1)
    Else
        Get核查人 = cbo核查人.Text
    End If
End Function
Public Function Get开单医生() As String
    If cbo开单医生.ListIndex = -1 Then
        Get开单医生 = ""
    ElseIf InStr(cbo开单医生.Text, "-") > 0 Then
        Get开单医生 = Mid(cbo开单医生.Text, InStr(cbo开单医生.Text, "-") + 1)
    Else
        Get开单医生 = cbo开单医生.Text
    End If
End Function
Private Sub Form_Load()
    mstrPrivs = gstrprivs
    
    Call SetParams
    Call Load医生
    
    Call IniRecipe
    Call FormClear
    
    If InStr(1, mstrPrivs, "医生查询") = 0 Then
        lblRP后记(9).Visible = False
    End If
End Sub

Private Sub Load配药人(ByVal lng药房ID As Long)
    '配药人
    Dim rsData As ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    gstrSQL = " Select 简码||'-'||姓名 As 姓名,姓名 As 名称 From 人员表  Where ID in " & _
             " (Select Distinct 人员ID From 人员性质说明 Where 人员性质='药房发药人' " & _
             " And 人员ID IN (Select 人员ID From 部门人员 Where 部门ID=[1]))" & _
             " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "取配药人", lng药房ID)
    
    With rsData
        Me.Cbo配药人.Clear
        If .EOF Then Exit Sub
        Do While Not .EOF
            Cbo配药人.AddItem !姓名
            
            If mstrDosUser = !名称 Then
                intIndex = .AbsolutePosition - 1
            End If
            
            .MoveNext
        Loop
        
        Cbo配药人.Enabled = Not Cbo配药人.ListCount = 0
        
        If intIndex <> -1 Then Cbo配药人.ListIndex = intIndex
        
        mstrDosUser = Me.Cbo配药人.Text
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Load核查人(ByVal lng药房ID As Long)
    '核查人
    Dim rsData As ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select 简码||'-'||姓名 As 姓名,姓名 As 名称 From 人员表 Where Id In (Select 人员id from 部门人员 Where 部门id=[1]) " & _
             " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "取审核处方人", lng药房ID)
    
    With rsData
        Me.cbo核查人.Clear
        If .EOF Then Exit Sub
        Do While Not .EOF
            cbo核查人.AddItem !姓名
            
            If mstr核查人 = !名称 Then
                intIndex = .AbsolutePosition - 1
            End If
            
            .MoveNext
        Loop
        
        cbo核查人.Enabled = Not cbo核查人.ListCount = 0
        
        If intIndex <> -1 Then cbo核查人.ListIndex = intIndex
        
        mstr核查人 = Me.cbo核查人.Text
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    With picRecipe(处方签区域.整体)
        .Left = (Me.Width - .Width) / 2
        .Height = Me.Height - IIf(picRecInfo_CM.Visible, picRecInfo_CM.Height, 0) - picProcess.Height - 200
    End With
    
    With picRecInfo_CM
'        If .Visible Then
            .Top = picRecipe(处方签区域.整体).Top + picRecipe(处方签区域.整体).Height + 100
            .Left = picRecipe(处方签区域.整体).Left
            .Width = picRecipe(处方签区域.整体).Width
'        End If
    End With
    
    With picProcess
        .Left = picRecipe(处方签区域.整体).Left
        .Top = Me.Height - .Height - 50
'        .Width = picRecipe(处方签区域.整体).Width
    End With
    
'    If Me.lbl开单医生.Visible = False And mbln标志 = False Then
'        Me.lbl配药人.Left = Me.lbl配药人.Left - 2200
'        Me.cbo配药人.Left = Me.cbo配药人.Left - 2200
'        Me.lbl核查人.Left = Me.lbl核查人.Left - 2200
'        Me.cbo核查人.Left = Me.cbo核查人.Left - 2200
'    End If
    
    mbln标志 = True
End Sub


Private Sub SetRecipeColor(index As Integer)
    Dim lngBackColor As Long
    Dim objTmp As Object
    Dim strTypeName As String
    
    Select Case index
        Case 0
            lblRP标识(RP标识.标识).Caption = "普通"
        Case 1
            lblRP标识(RP标识.标识).Caption = "儿科"
        Case 2
            lblRP标识(RP标识.标识).Caption = "急诊"
        Case 3
            lblRP标识(RP标识.标识).Caption = "精二"
        Case 4
            lblRP标识(RP标识.标识).Caption = "精一"
        Case 5
            lblRP标识(RP标识.标识).Caption = "麻"
    End Select
    
    lngBackColor = Val(Split(mstrUserRecipeColor, ";")(index))
    
    For Each objTmp In lblRP标题
        objTmp.BackColor = lngBackColor
    Next
    
    For Each objTmp In lblRP前记
        objTmp.BackColor = lngBackColor
    Next
    
    For Each objTmp In lblRP正文
        objTmp.BackColor = lngBackColor
    Next
    
    For Each objTmp In lblRP后记
        objTmp.BackColor = lngBackColor
    Next
    
    For Each objTmp In lblRP标识
        objTmp.BackColor = lngBackColor
    Next
    
    For Each objTmp In picRecipe
        objTmp.BackColor = lngBackColor
    Next
    
    With vsfRecipe
        .BackColorFixed = lngBackColor
        .BackColor = lngBackColor
        .BackColorBkg = lngBackColor
    End With
    
    txt诊断内容.BackColor = lngBackColor
End Sub


Private Sub picProcess_Resize()
    With CmdSend
        .Left = picProcess.Width - .Width - 100
    End With
End Sub


Private Sub picRecipe_Resize(index As Integer)
    On Error Resume Next
    
    If index = 0 Then
        With lblRP标题(RP标题.医院名称)
            .Left = 0
            .Width = picRecipe(处方签区域.标题).Width
        End With
        
        With picRecipe(处方签区域.后记)
            .Top = picRecipe(处方签区域.整体).Height - .Height
        End With
        
        With picRecipe(处方签区域.正文)
            .Height = picRecipe(处方签区域.后记).Top - .Top
        End With
        
        With vsfRecipe
            .Left = lblRP正文(RP正文.标识).Left
            .Top = lblRP正文(RP正文.标识).Top + lblRP正文(RP正文.标识).Height + 150
            .Width = picRecipe(处方签区域.正文).Width - .Left - 100
            .Height = picRecipe(处方签区域.正文).Height - .Top - 100
        End With
    ElseIf index = 处方签区域.前记 Then
        picRecipe(处方签区域.前记).Height = txt诊断内容.Top + txt诊断内容.Height + 50
        
        With picRecipe(处方签区域.正文)
            .Top = picRecipe(处方签区域.前记).Top + picRecipe(处方签区域.前记).Height + 50
            .Height = picRecipe(处方签区域.后记).Top - .Top
        End With
    End If
End Sub
Public Sub RefreshRecipe(ByVal rsData As ADODB.Recordset, ByVal strWeight As String, Optional ByVal int可操作 As Integer = 0, Optional int排队状态 As Integer, Optional int审查结果 As Integer)
    Dim dbl应收金额, dbl实收金额 As Double
    Dim str操作员 As String
    Dim IntLocate As Integer
    Dim strDiag As String
    Dim int门诊 As Integer
    Dim i As Integer
 
    FormClear
    
    CmdSend.Enabled = False
    
    With rsData
        .Filter = ""
        
        If .EOF Then Exit Sub
        
        
        mbln中药处方 = False
        If 判断是否中药处方(!药房ID, !单据, !NO) Then
            mbln中药处方 = True
        End If
            
        Call IniRecipe
        
        If !记录性质 = 1 Or (!记录性质 = 2 And (!门诊标志 = 1 Or !门诊标志 = 4)) Then
            int门诊 = 1
        Else
            int门诊 = 2
        End If
        
        '标题
        lblRP标题(RP标题.医院名称).Caption = GetUnitName
        
        '标识
        lblRP标识(RP标识.标识).Caption = Split(gconstrRecipeType, ";")(Val(!处方类型))
        
        '前记
        lblRP前记(RP前记.费别).Caption = IIf(IsNull(!费别), "", !费别)
'        lblRP前记(RP前记.险类).Caption = IIf(IsNull(!险类), "", !险类)
        lblRP前记(RP前记.处方号).Caption = !NO
        lblRP前记(RP前记.姓名).Caption = IIf(IsNull(!姓名), "", !姓名)
        lblRP前记(RP前记.姓名).ForeColor = zldatabase.GetPatiColor(IIf(IsNull(!病人类型), "", !病人类型))
        
        lblRP前记(RP前记.性别).Caption = IIf(IsNull(!性别), "", !性别)
        lblRP前记(RP前记.年龄).Caption = IIf(IsNull(!年龄), "", !年龄)
        
        lblRP前记(RP前记.就诊卡号).Caption = IIf(IsNull(!就诊卡号), "", !就诊卡号)
        
        If !门诊标志 = 1 Or !门诊标志 = 4 Then
            lblRP前记(RP前记.标识号标签).Caption = "门诊号："
        Else
            lblRP前记(RP前记.标识号标签).Caption = "住院号："
        End If
        
        lblRP前记(RP前记.标识号).Caption = IIf(IsNull(!住院号), "", !住院号)
        
        lblRP前记(RP前记.科室).Caption = IIf(IsNull(!科室), "", !科室)
        lblRP前记(RP前记.床号).Caption = IIf(IsNull(!床号), "", !床号)
        lblRP前记(RP前记.体重).Caption = IIf(IsNumeric(strWeight), strWeight & "kg", strWeight)
        
        '诊断信息
        txt诊断内容.Text = ""
        txt诊断内容.Tag = ""
        txt诊断内容.Height = 180
        
        Call picRecipe_Resize(处方签区域.前记)

        '正文
        SetRecipeMedi rsData
        
        '后记
        .Filter = ""
        Do While Not .EOF
            dbl应收金额 = dbl应收金额 + Val(!零售金额)
            dbl实收金额 = dbl实收金额 + Val(!实收金额)
            .MoveNext
        Loop
        .MoveFirst
        
        If mcondition.int金额显示 = 1 Then
            lblRP后记(RP后记.合计金额标签).Caption = "实收合计："
            lblRP后记(RP后记.合计金额).Caption = zlStr.FormatEx(dbl实收金额, 2, , True) & "元"
        ElseIf mcondition.int金额显示 = 2 Then
            lblRP后记(RP后记.合计金额标签).Caption = "应收/实收合计："
            lblRP后记(RP后记.合计金额).Caption = zlStr.FormatEx(dbl应收金额, 2, , True) & "元/" & zlStr.FormatEx(dbl实收金额, 2, , True) & "元"
        Else
            lblRP后记(RP后记.合计金额标签).Caption = "应收合计："
            lblRP后记(RP后记.合计金额).Caption = zlStr.FormatEx(dbl应收金额, 2) & "元"
        End If
        lblRP后记(RP后记.合计金额标签).Left = lblRP后记(RP后记.配药人标签).Left - (lblRP后记(RP后记.合计金额标签).Width - lblRP后记(RP后记.配药人标签).Width)
        
        lblRP后记(RP后记.日期).Caption = IIf(IsNull(!填制日期), "", Format(!填制日期, "yyyy-mm-dd"))
        lblRP后记(RP后记.收费员).Caption = IIf(IsNull(!操作员姓名), "", !操作员姓名)
        lblRP后记(RP后记.配药人).Caption = IIf(IsNull(!配药人), "", !配药人)
        lblRP后记(RP后记.开单医生).Caption = IIf(IsNull(!开单人), "", !开单人)
                
        '设置处方颜色
        SetRecipeColor Val(!处方类型)
        
        '设置开单医生
        Me.cbo开单医生.ListIndex = 0
        If (mcondition.bln校验处方 = False) And zlStr.IsHavePrivs(gstrprivs, "医生查询") Then
            str操作员 = IIf(IsNull(!开单人), "", !开单人)
        Else
            If mcondition.intListType = mListType.退药 And mcondition.bln校验处方 = True Then
                str操作员 = IIf(IsNull(!填制人), "", !填制人)
            Else
                str操作员 = ""
            End If
        End If
        If str操作员 <> "" Then
            '定位医生
            For IntLocate = 1 To cbo开单医生.ListCount
                If Mid(cbo开单医生.List(IntLocate), InStr(1, cbo开单医生.List(IntLocate), "-") + 1) = str操作员 Then
                    cbo开单医生.ListIndex = IntLocate
                    Exit For
                End If
            Next
        End If
        
        cbo开单医生.Enabled = ((mcondition.intListType = mListType.待发药 Or mcondition.intListType = mListType.超时未发) And mcondition.bln校验处方 = True And cbo开单医生.ListIndex = 0)
        
        Lbl配药人.Caption = IIf(mcondition.intListType = mListType.退药, IIf(int可操作 <> 3, "发药人", "退药人"), "配药人")
        '设置配药人
        If mcondition.intListType = mListType.退药 Then
            If IIf(IsNull(!发药人), "", !发药人) <> "" Then
                Me.Cbo配药人 = IIf(IsNull(!发药人), "", !发药人)
            End If
        Else
        
            If IIf(IsNull(!配药人), "", !配药人) <> "" Then
                Me.Cbo配药人 = IIf(IsNull(!配药人), "", !配药人)
            End If
        End If
        
        If mcondition.intListType = mListType.退药 Then
            If IIf(IsNull(!核查人), "", !核查人) <> "" Then
                Me.cbo核查人.Text = IIf(IsNull(!核查人), "", !核查人)
            End If
        End If
        
        '中药处方
        picRecInfo_CM.Visible = False
        If mbln中药处方 Then
            picRecInfo_CM.Visible = True
            Call 中药处方特别处理(!药房ID, !单据, !NO, !记录性质, !门诊标志)
        End If
        Call Form_Resize
    End With
    
    SetCmdSendPrivs int审查结果
    
    If Me.CmdSend.Caption = "配药确认(&O)" And int排队状态 = 1 Then
        Me.CmdSend.Caption = "取消确认(&C)"
    ElseIf Me.CmdSend.Caption = "取消确认(&C)" And int排队状态 = 0 Then
        Me.CmdSend.Caption = "配药确认(&O)"
    End If
End Sub

Private Function 判断是否中药处方(ByVal lngNO药房id As Long, ByVal BillType As Integer, ByVal BillNo As String) As Boolean
    '通过药品id判断是否是中药
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    Dim lng药房ID As Long
    Dim blnMoved As Boolean
    
    On Error GoTo errHandle
    
    lng药房ID = lngNO药房id
    If lngNO药房id = 0 Then lng药房ID = mcondition.lng药房ID
    
    strsql = "Select a.类别 as 类别 From 收费项目目录 a ,药品收发记录 b Where b.药品id=a.Id And b.单据=[2] and b.No=[1] And (b.记录状态=1 Or Mod(b.记录状态,3)=0) and (b.库房ID+0=[3] OR b.库房ID IS NULL) " _
   
    '如果数据转出，则直接从后备表中提取数据
    blnMoved = Sys.IsMovedByNO("药品收发记录", BillNo, " 单据 = ", BillType)
    If blnMoved Then
        gstrSQL = Replace(gstrSQL, "药品收发记录", "H药品收发记录")
    End If
    
    Set rs = zldatabase.OpenSQLRecord(strsql, Me.Caption & "[判断是否中药处方]", BillNo, BillType, lng药房ID)
    
    判断是否中药处方 = IIf(rs!类别 = 7, True, False)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub 中药处方特别处理(ByVal lngNO药房id As Long, ByVal BillStyle As Integer, ByVal BillNo As String, ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer)
    '中药处方显示原始付数和中药煎法
    Dim rs As New ADODB.Recordset
    Dim lng药房ID As Long
    
    On Error GoTo errHandle
    lng药房ID = lngNO药房id
    If lngNO药房id = 0 Then lng药房ID = mcondition.lng药房ID

    gstrSQL = "Select a.外观,b.付数 From 药品收发记录 a ,门诊费用记录 b Where a.费用id=b.Id " _
        & " And a.单据=[2] And a.No=[1] " _
        & " And (a.记录状态=1 Or Mod(a.记录状态,3)=0) and (a.库房ID+0=[3] OR a.库房ID IS NULL) "
    If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    End If
    
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[中药处方特别处理]", BillNo, BillStyle, lng药房ID)
    
    lbl原始付数.Caption = lbl原始付数.Tag & CStr(IIf(IsNull(rs!付数), 1, rs!付数))
    lbl中药煎法.Caption = lbl中药煎法.Tag & IIf(IsNull(rs!外观), "", rs!外观)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt诊断内容_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetTip(txt诊断内容, txt诊断内容.Tag)
End Sub
