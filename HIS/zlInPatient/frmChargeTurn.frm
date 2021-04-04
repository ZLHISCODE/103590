VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmChargeTurn 
   AutoRedraw      =   -1  'True
   Caption         =   "门(急)诊费用转住院"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11715
   Icon            =   "frmChargeTurn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11715
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBill 
      Height          =   2100
      Left            =   90
      ScaleHeight     =   2040
      ScaleWidth      =   10485
      TabIndex        =   21
      Top             =   645
      Width           =   10545
      Begin VSFlex8Ctl.VSFlexGrid mshList 
         Height          =   1470
         Left            =   75
         TabIndex        =   22
         Top             =   90
         Width           =   5490
         _cx             =   9684
         _cy             =   2593
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
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
   End
   Begin VB.PictureBox picBalance 
      Height          =   1950
      Left            =   6285
      ScaleHeight     =   1890
      ScaleWidth      =   2985
      TabIndex        =   19
      Top             =   4035
      Width           =   3045
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   1335
         Left            =   0
         TabIndex        =   20
         Top             =   135
         Width           =   2565
         _cx             =   4524
         _cy             =   2355
         Appearance      =   3
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
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
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         Caption         =   "转出合计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   105
         TabIndex        =   23
         Top             =   1605
         Width           =   1155
      End
   End
   Begin VB.PictureBox picList 
      Height          =   1935
      Left            =   105
      ScaleHeight     =   1875
      ScaleWidth      =   5415
      TabIndex        =   17
      Top             =   3945
      Width           =   5475
      Begin VSFlex8Ctl.VSFlexGrid mshDetail 
         Height          =   1185
         Left            =   30
         TabIndex        =   18
         Top             =   165
         Width           =   5130
         _cx             =   9049
         _cy             =   2090
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
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
   End
   Begin VB.PictureBox picTop 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11715
      TabIndex        =   7
      Top             =   0
      Width           =   11715
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新(&R)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   8670
         TabIndex        =   15
         Top             =   95
         Width           =   1100
      End
      Begin VB.Frame fraPati 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   150
         TabIndex        =   10
         Top             =   -45
         Width           =   2910
         Begin VB.TextBox txtPatient 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   1200
            MaxLength       =   64
            TabIndex        =   11
            ToolTipText     =   "热键：F11"
            Top             =   135
            Width           =   1650
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   345
            Left            =   570
            TabIndex        =   25
            Top             =   135
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   609
            Appearance      =   2
            IDKindStr       =   $"frmChargeTurn.frx":058A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   12
            FontName        =   "宋体"
            IDKind          =   -1
            ShowPropertySet =   -1  'True
            NotContainFastKey=   "F1;CTRL+F1;F2;F3;CTRL+F4;F5;F6;F7;CTRL+F7;F8;F9;F10;F11;F12;CTRL+F12;CTRL+S;CTRL+A;CTRL+R;CTRL+D;CTRL+Q;ESC;ALT+?"
            MustSelectItems =   "姓名,就诊卡"
            BackColor       =   -2147483633
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病人"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   60
            TabIndex        =   12
            Top             =   180
            Width           =   480
         End
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   348
         Left            =   1152
         TabIndex        =   0
         Top             =   96
         Width           =   2664
         _ExtentX        =   4710
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   93323267
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   348
         Left            =   4116
         TabIndex        =   16
         Top             =   96
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   93323267
         CurrentDate     =   36588
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "仅显示可转入数据"
         Height          =   345
         Left            =   6870
         TabIndex        =   24
         Top             =   120
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3840
         TabIndex        =   9
         Top             =   180
         Width           =   120
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发生时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   192
         TabIndex        =   8
         Top             =   168
         Width           =   960
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11715
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7665
      Width           =   11715
      Begin VB.CommandButton cmdParaSet 
         Caption         =   "参数设置(&R)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3804
         TabIndex        =   14
         Top             =   0
         Width           =   1500
      End
      Begin VB.CommandButton cmdSave 
         Cancel          =   -1  'True
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7545
         TabIndex        =   13
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   60
         TabIndex        =   4
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "全清(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   2640
         TabIndex        =   2
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "全选(&A)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   1455
         TabIndex        =   1
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "取消(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8670
         TabIndex        =   3
         Top             =   0
         Width           =   1100
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   8100
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeTurn.frx":0620
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15584
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
            AutoSize        =   2
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChargeTurn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrNOS As String '要进行费用转入的单据信息,格式：单据,票据,结帐ID,险类(非医保为零),单据类型,补充结算单据号:H0000001,F000023,81235,901,收费单(记帐单),S0000001;...
Private mlngPatient As Long, mlng主页ID As Long
Private msngOldY As Single, msngOldX As Single
Private Enum 交易Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
End Enum
Private Enum 医院业务
    support门诊结算作废 = 33        '医保是否支持门诊结算作废，不支持只有个人帐帐户原样退,其余的医保结算方式退为现金,支持的再判断每一种结算方式是否允许退回
    support多单据收费必须全退 = 39  '多单据收费必须全退
End Enum
Private Enum IDKinds
    C0姓名 = 0
    C1医保号 = 1
    C2身份证号 = 2
    C3IC卡号 = 3
    C4门诊号 = 4
    C5住院号 = 5
    C6就诊卡 = 6
End Enum
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

Private mblnSelPati As Boolean '是否选择病人
Private mintPatientRange As Integer
Private mrsInfo As ADODB.Recordset
Private mlngTXTProc As Long
Private mstrPrivs As String, mlngModule As Long
Private mbln门诊转住院先审核 As Boolean
Private mbln立即销帐 As Boolean
Private Enum mObjPancel
    Pan_Search = 1
    Pan_Bill = 2
    Pan_List = 3
    Pan_Balance = 4
    Pan_Bottom = 5
End Enum
Private mrsOneCard  As ADODB.Recordset

'关于消费卡的处理变量
Private Type Ty_SquareCard
    blnExistsObjects As Boolean     '安装了消费卡的
    rsSquare As ADODB.Recordset
    dbl刷卡总额 As Double
    bln卡结算 As Boolean '当前读取的单据是卡结算
    str刷卡结算 As String   '刷卡结算方式;金额;是否允许修改|..."
End Type
Private mtySquareCard As Ty_SquareCard
Private mstrThreeSwapBalance As String
Private mstrThreeSwapCardType As String
Private mstrThreeSwapMoney As String
Private mintIDKind As Integer
Private mobjSquare As Object
Private mblnPassInputCardNo As Boolean  '是否密文输入卡号
Private mblnDefaultPassInputCardNo As Boolean '缺省刷卡是否密文输入卡号
Private mlng医疗卡长度  As Long
Private mblnNotClick As Boolean
Private mstrTitle As String

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:区域设置
    '编制:刘兴洪
    '日期:2011-03-25 17:37:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, strReg As String, panThis As Pane
    Dim panTop As Pane, panBottom As Pane, panRight As Pane
    
    Set panTop = dkpMan.CreatePane(mObjPancel.Pan_Search, 200, 580, DockTopOf, Nothing)
    panTop.Title = "条件窗体"
    panTop.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panTop.Tag = mObjPancel.Pan_Search
    panTop.Handle = picTop.hWnd
    panTop.MaxTrackSize.Height = 495 / Screen.TwipsPerPixelY
    panTop.MinTrackSize.Height = 495 / Screen.TwipsPerPixelY
    
    Set panThis = dkpMan.CreatePane(mObjPancel.Pan_Bill, 250, 580, DockBottomOf, panTop)
    panThis.Title = "门诊转住院列表"
    panThis.Tag = mObjPancel.Pan_Bill
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picBill.hWnd
    

    Set panRight = dkpMan.CreatePane(mObjPancel.Pan_Balance, 1500 / Screen.TwipsPerPixelX, 580, DockRightOf, panThis)
    panRight.Title = "门诊转住院结算信息"
    panRight.Tag = mObjPancel.Pan_Balance
    panRight.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panRight.Handle = picBalance.hWnd
    
    Set panThis = dkpMan.CreatePane(mObjPancel.Pan_List, 250, 580, DockBottomOf, panThis)
    panThis.Title = "单据明细列表"
    panThis.Tag = mObjPancel.Pan_List
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picList.hWnd
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pan_Search
        Item.Handle = picTop.hWnd
    Case Pan_Bill
        Item.Handle = picBill.hWnd
    Case Pan_List
        Item.Handle = picList.hWnd
    Case Pan_Balance
        Item.Handle = picBalance.hWnd
    End Select
End Sub

Public Sub ShowMe(objParent As Object, ByVal lngPatient As Long, ByRef strNos As String, _
    Optional blnSelPati As Boolean = False, Optional intPatientRange As Integer = 0, _
    Optional strPrivs As String, Optional lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊费用转住院费用
    '入参:lngPatient-病人ID
    '      blnSelPati-是否需要选择病人
    '      intPatientRange:(0-所有病人,1-任何费用未结清病人;2-体检未结清的病人;3-住院未结清的病人;4-门诊未结清的病人)
    '出参:
    '   strNOS:要进行费用转入的单据信息,格式：
    '       单据,票据,结帐ID,险类(非医保为零),单据类型,补充结算单据号:H0000001,F000023,81235,901,收费单(记帐单),S0000001;...
    '返回:
    '编制:刘兴洪
    '日期:2010-11-09 17:09:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnSelPati = blnSelPati: mintPatientRange = intPatientRange
    mlngPatient = lngPatient: mstrPrivs = strPrivs: mlngModule = lngModule
    mstrNOS = strNos
    
    If mblnSelPati = False Then
        '此时会先隐式调用事件Form_Load
        Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
        Call SetBillSelected(strNos)
    Else
            If lngPatient <> 0 Then
                If GetPatient(IDKind.GetCurCard, "-" & lngPatient, 0) Then
                    Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
                End If
            Else
                Call ClearData
            End If
    End If
    If mblnSelPati = False Then
        fraPati.Visible = False: cmdSave.Visible = True
    Else
        fraPati.Visible = True: cmdSave.Visible = True
    End If
    Call picTop_Resize
    Call Me.Show(vbModal, objParent)
    strNos = mstrNOS
End Sub

Private Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中输入的负数数量及退回科室是否正确
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-11-09 17:30:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mshList.Redraw = flexRDNone
    mshList.Clear 1: mshList.Rows = 2
    sta.Panels(2).Text = ""
    Call setHeader: Call SetBillColor
    mshList.Redraw = flexRDBuffered
End Sub

Private Sub SetBillSelected(ByVal strNos As String)
'说明:如果转入几天后失败,再进入选择窗体,以前选择的且已被转入的单据现在是"不可转入",所以不应被选择
    Dim i As Long
    With mshList
        For i = 1 To .Rows - 1
            If InStr(";" & strNos, ";" & .TextMatrix(i, .ColIndex("单据号"))) > 0 And .TextMatrix(i, .ColIndex("类别")) = "可转入" Then
                .TextMatrix(i, .ColIndex("选择")) = "√"
            Else
                .TextMatrix(i, .ColIndex("选择")) = ""
            End If
        Next
    End With
End Sub

Public Function CheckExistTurn(ByVal lngPatient As Long, ByRef dat入院时间 As Date) As Boolean
'功能:检查入院时间之后是否存在转入数据
'返回:转入数据的登记时间
    Dim rsTmp As ADODB.Recordset, strSQL As String
        
    On Error GoTo errH
    strSQL = "" & _
    " Select Max(发生时间) 发生时间 " & _
    " From 住院费用记录" & vbNewLine & _
    " Where 记录性质 = 2 And 记录状态 In(1,3) And 病人id = [1] And 主页id Is Null And 标识号 Is Null And 门诊标志=2" & vbNewLine & _
    "       And 摘要='门诊费用转入'"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查是否存在已转单据", lngPatient, dat入院时间)
    
    If Not IsNull(rsTmp!发生时间) Then
        dat入院时间 = rsTmp!发生时间
        CheckExistTurn = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ExecteUpdate(ByVal lngPatient As Long, ByVal str住院号 As String, ByVal lngPageID As Long, ByVal dat入院时间 As Date)
'功能:更新记帐单的主页ID
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Zl_门诊费用转住院_Update(" & lngPatient & "," & str住院号 & "," & lngPageID & _
            ",To_Date('" & Format(dat入院时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    Call zlDatabase.ExecuteProcedure(strSQL, "更新记帐单")
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsYBSingle(ByVal strno As String, ByVal intInsure As Integer) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset, blnInsureSingle As Boolean
    
    blnInsureSingle = gclsInsure.GetCapability(83, , intInsure)
    If blnInsureSingle = False Then
        IsYBSingle = False
        Exit Function
    Else
        strSQL = "Select 1 From 医保结算明细 Where NO = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strno)
        If rsTmp.EOF Then
            IsYBSingle = False
        Else
            If CheckAllTurn(strno) Then
                IsYBSingle = False
            Else
                IsYBSingle = True
            End If
        End If
    End If
End Function

Public Function ExecuteTurn(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal strNos As String, ByVal str住院号 As String, ByVal lng主页ID As Long, _
    ByVal dat入院时间 As Date, ByVal lng入院科室ID As Long, ByVal lng入院病区ID As Long, _
    Optional ByRef strOutDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的单据号序列,执行门诊费用转住院费用,及医保退费结算操作
    '入参:
    '   strNOS:要进行费用转入的单据信息,格式：
    '       单据,票据,结帐ID,险类(非医保为零),单据类型,补充结算单据号:H0000001,F000023,81235,901,收费单(记帐单),S0000001;...
    '   lng住院号-住院号,lng主页ID-主页ID,这两个参数仅在医保入院补充登记时才传入
    '出参:strDelDate-本次转出日期(目前主要是重新获取预交款数据)
    '返回:
    '编制:刘兴洪
    '日期:2011-02-16 10:26:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim DateDel As Date, arrNO As Variant, arrInfo As Variant
    Dim i As Long, j As Long, lngcnt As Long, bln医保单张退 As Boolean
    Dim strSQL As String, strInvoice As String, strInDate As String, strDelDate As String
    Dim cllJzPro As Collection, rsTemp As ADODB.Recordset, str已转结帐ID As String
    Dim blnTrans As Boolean, blnTransMedicare As Boolean, blnExecuteThreeSwap As Boolean
    Dim intInsure As Integer, strAdvance As String, strJzNOs As String
    Dim rsDeposit As ADODB.Recordset, lng结帐ID As Long, blnTransMC As Boolean
    Dim str交易说明 As String, str交易流水号 As String, blnTurnAll As Boolean
    
    '补充结算的单据处理思路：先将费用单据转为住院费用记录，再单独处理门诊退费
    Dim strReplenishNo As String, strReplenishNos As String '格式：单据号,险类
    Dim cllReplenishPro As Collection
    
    mstrPrivs = strPrivs: mlngModule = lngModule
    If strNos = "" Then Exit Function
    
    strInDate = "To_Date('" & Format(dat入院时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    strOutDelDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    strDelDate = "To_Date('" & strOutDelDate & "','YYYY-MM-DD HH24:MI:SS')"
    arrNO = Split(strNos, ";")
    Set cllJzPro = New Collection
    Set cllReplenishPro = New Collection
    
    On Error GoTo errH
    strJzNOs = ""
    i = LBound(arrNO)
    Do While i <= UBound(arrNO)
        lngcnt = 1
        strInvoice = Trim(Split(arrNO(i), ",")(1))
        If strInvoice <> "" Then
            For j = i + 1 To UBound(arrNO)
                If strInvoice = Split(arrNO(j), ",")(1) Then
                    lngcnt = lngcnt + 1
                Else
                    Exit For
                End If
            Next
        End If
        
        '医保要求从最后一张开始退,读出的数据是按日期倒序排列的，所以此处正序即可
        For j = i To i + lngcnt - 1
            arrInfo = Split(arrNO(j), ",")
            bln医保单张退 = False: blnTurnAll = False
            
            strReplenishNo = arrInfo(5)
            If strReplenishNo = "" Then
                If Val(arrInfo(3)) <> 0 Then
                    bln医保单张退 = IsYBSingle(arrInfo(0), Val(arrInfo(3)))
                Else
                    blnTurnAll = CheckAllTurn(arrInfo(0))
                    If InStr("," & str已转结帐ID & ",", "," & arrInfo(2) & ",") > 0 Then blnTurnAll = True
                End If
            
                If mbln立即销帐 Then lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
            End If
            
            If bln医保单张退 Or (Val(arrInfo(3)) = 0 And Not blnTurnAll) Or strReplenishNo <> "" Then
                'Zl_门诊费用转住院_Insert
                strSQL = "Zl_门诊费用转住院_insert("
                '  No_In         住院费用记录.NO%Type,
                strSQL = strSQL & "'" & arrInfo(0) & "',"
                '  住院号_In     住院费用记录.标识号%Type, --医保入院补充登记时才传入
                strSQL = strSQL & "" & IIf(str住院号 = "", "Null", str住院号) & ","
                '  主页id_In     住院费用记录.主页id%Type, --医保入院补充登记时才传入
                strSQL = strSQL & "" & IIf(lng主页ID = 0, "Null", lng主页ID) & ","
                '  入院时间_In   住院费用记录.发生时间%Type,
                strSQL = strSQL & "" & strInDate & ","
                '  入院科室id_In 病人预交记录.科室id%Type,
                strSQL = strSQL & "" & IIf(lng入院科室ID = 0, "NULL", lng入院科室ID) & ","
                '  退费时间_In   住院费用记录.登记时间%Type, --多张单据退费时,每张单据的退费时间相同,都是系统当前时间
                strSQL = strSQL & "" & strDelDate & ","
                '  操作员编号_In 住院费用记录.操作员编号%Type,
                strSQL = strSQL & "'" & UserInfo.编号 & "',"
                '  操作员姓名_In 住院费用记录.操作员姓名%Type,
                strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                '  入院病区id_In 住院费用记录.病人病区id%Type := Null,
                strSQL = strSQL & "" & IIf(lng入院病区ID = 0, "NULL", lng入院病区ID) & ","
                '  单据_In Number:=1(1-门诊收费单;2-记帐单)
                strSQL = strSQL & "" & IIf(arrInfo(4) = "记帐单", 2, 1) & ","
                '  结帐ID_In     住院费用记录.结帐id%Type,
                strSQL = strSQL & "" & IIf(mbln立即销帐 And strReplenishNo = "", lng结帐ID, "NULL") & ","
                '  原结帐id_In   住院费用记录.结帐id%Type := Null,
                strSQL = strSQL & "" & "NULL" & ","
                '  立即销帐_In Number:=1
                strSQL = strSQL & "" & IIf(mbln立即销帐 And strReplenishNo = "", "1", "0") & ")"
                
                blnExecuteThreeSwap = False
                mstrThreeSwapBalance = ""
                mstrThreeSwapCardType = ""
                mstrThreeSwapMoney = ""
                
                If strReplenishNo <> "" Then
                    If InStr(strReplenishNos & ";", ";" & strReplenishNo & "," & arrInfo(3) & ";") = 0 Then
                        strReplenishNos = strReplenishNos & ";" & strReplenishNo & "," & arrInfo(3)
                    End If
                    cllReplenishPro.Add Array(strReplenishNo, strSQL)
                ElseIf arrInfo(4) = "记帐单" And mbln立即销帐 Then
                    If InStr(strJzNOs & ",", "," & arrInfo(0) & ",") = 0 Then
                        strJzNOs = strJzNOs & "," & arrInfo(0)
                        cllJzPro.Add strSQL, "K" & arrInfo(0)
                    End If
                Else
                    gcnOracle.BeginTrans: blnTrans = True
                
                    Call zlDatabase.ExecuteProcedure(strSQL, "门诊费用转住院")
                    '处理医保
                    blnTransMedicare = False
                    intInsure = IIf(arrInfo(4) = "记帐单", 0, Val(arrInfo(3)))
                    If mbln立即销帐 = False Then intInsure = 0  '只有立即销帐的,才会去调用医保接口
                    If intInsure <> 0 Then
                        strAdvance = lng结帐ID & "|0|" & arrInfo(0)
                        If Not gclsInsure.ClinicDelSwap(Abs(Val(arrInfo(2))), , intInsure, strAdvance) Then
                            gcnOracle.RollbackTrans
                            MsgBox "医保结算失败，无法进行门诊费用转住院操作。", vbInformation, gstrSysName
                            Exit Function
                        Else
                            blnTransMedicare = True
                        End If
                    End If
                    gcnOracle.CommitTrans: blnTrans = False
                    If blnTransMedicare And mbln立即销帐 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
                    
                    If mbln立即销帐 And arrInfo(4) <> "记帐单" Then
                        If ExecuteThreeSwap(Val(arrInfo(2)), lng结帐ID, str交易流水号, str交易说明) = True Then
                            blnExecuteThreeSwap = True
                        End If
                        'Zl_门诊转住院_三方卡结算
                        strSQL = "Zl_门诊转住院_三方卡结算("
                        '  No_In         住院费用记录.NO%Type,
                        strSQL = strSQL & "'" & arrInfo(0) & "',"
                        '  操作员编号_In 住院费用记录.操作员编号%Type,
                        strSQL = strSQL & "'" & UserInfo.编号 & "',"
                        '  操作员姓名_In 住院费用记录.操作员姓名%Type,
                        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                        '  退费时间_In   住院费用记录.登记时间%Type, --多张单据退费时,每张单据的退费时间相同,都是系统当前时间
                        strSQL = strSQL & "" & strDelDate & ","
                        '  门诊退费_In   Number := 0,
                        strSQL = strSQL & "" & 0 & ","
                        '  入院科室id_In 病人预交记录.科室id%Type,
                        strSQL = strSQL & "" & IIf(lng入院科室ID = 0, "NULL", lng入院科室ID) & ","
                        '  主页id_In     住院费用记录.主页id%Type, --医保入院补充登记时才传入
                        strSQL = strSQL & "" & IIf(lng主页ID = 0, "Null", lng主页ID) & ","
                        '  三方退费_In   Number := 0,
                        strSQL = strSQL & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
                        '  结帐ID_In     住院费用记录.结帐id%Type)
                        strSQL = strSQL & "" & lng结帐ID & ")"
                        Call zlDatabase.ExecuteProcedure(strSQL, "三方卡结算")
                    End If
                End If
            Else
                If InStr("," & str已转结帐ID & ",", "," & arrInfo(2) & ",") = 0 Then
                    'Zl_门诊费用转住院_Insert
                    strSQL = "Zl_门诊费用转住院_insert("
                    '  No_In         住院费用记录.NO%Type,
                    strSQL = strSQL & "'" & arrInfo(0) & "',"
                    '  住院号_In     住院费用记录.标识号%Type, --医保入院补充登记时才传入
                    strSQL = strSQL & "" & IIf(str住院号 = "", "Null", str住院号) & ","
                    '  主页id_In     住院费用记录.主页id%Type, --医保入院补充登记时才传入
                    strSQL = strSQL & "" & IIf(lng主页ID = 0, "Null", lng主页ID) & ","
                    '  入院时间_In   住院费用记录.发生时间%Type,
                    strSQL = strSQL & "" & strInDate & ","
                    '  入院科室id_In 病人预交记录.科室id%Type,
                    strSQL = strSQL & "" & IIf(lng入院科室ID = 0, "NULL", lng入院科室ID) & ","
                    '  退费时间_In   住院费用记录.登记时间%Type, --多张单据退费时,每张单据的退费时间相同,都是系统当前时间
                    strSQL = strSQL & "" & strDelDate & ","
                    '  操作员编号_In 住院费用记录.操作员编号%Type,
                    strSQL = strSQL & "'" & UserInfo.编号 & "',"
                    '  操作员姓名_In 住院费用记录.操作员姓名%Type,
                    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                    '  入院病区id_In 住院费用记录.病人病区id%Type := Null,
                    strSQL = strSQL & "" & IIf(lng入院病区ID = 0, "NULL", lng入院病区ID) & ","
                    '  单据_In Number:=1(1-门诊收费单;2-记帐单)
                    strSQL = strSQL & "" & IIf(arrInfo(4) = "记帐单", 2, 1) & ","
                    '  结帐ID_In     住院费用记录.结帐id%Type)
                    strSQL = strSQL & "" & IIf(mbln立即销帐, lng结帐ID, "NULL") & ","
                    '  原结帐ID_In     住院费用记录.结帐id%Type,
                    strSQL = strSQL & "" & arrInfo(2) & ","
                    '  立即销帐_In Number:=1
                    strSQL = strSQL & "" & IIf(mbln立即销帐, "1", "0") & ")"
                    
                    blnExecuteThreeSwap = False
                    mstrThreeSwapBalance = ""
                    mstrThreeSwapCardType = ""
                    mstrThreeSwapMoney = ""
                    
                    If arrInfo(4) = "记帐单" And mbln立即销帐 Then
                        If InStr(strJzNOs & ",", "," & arrInfo(0) & ",") = 0 Then
                            strJzNOs = strJzNOs & "," & arrInfo(0)
                            cllJzPro.Add strSQL, "K" & arrInfo(0)
                        End If
                    Else
                        gcnOracle.BeginTrans: blnTrans = True
                        Call zlDatabase.ExecuteProcedure(strSQL, "门诊费用转住院")
                        '处理医保
                        blnTransMedicare = False
                        intInsure = IIf(arrInfo(4) = "记帐单", 0, Val(arrInfo(3)))
                        If mbln立即销帐 = False Then intInsure = 0  '只有立即销帐的,才会去调用医保接口
                        If intInsure <> 0 Then
                            strAdvance = lng结帐ID & "|0"
                            If Not gclsInsure.ClinicDelSwap(Abs(Val(arrInfo(2))), , intInsure, strAdvance) Then
                                gcnOracle.RollbackTrans
                                MsgBox "医保结算失败，无法进行门诊费用转住院操作。", vbInformation, gstrSysName
                                Exit Function
                            Else
                                blnTransMedicare = True
                            End If
                        End If
                        gcnOracle.CommitTrans: blnTrans = False
                        If blnTransMedicare And mbln立即销帐 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
                        
                        If mbln立即销帐 And arrInfo(4) <> "记帐单" Then
                            If ExecuteThreeSwap(Val(arrInfo(2)), lng结帐ID, str交易流水号, str交易说明) = True Then
                                blnExecuteThreeSwap = True
                            End If
                            'Zl_门诊转住院_三方卡结算
                            strSQL = "Zl_门诊转住院_三方卡结算("
                            '  No_In         住院费用记录.NO%Type,
                            strSQL = strSQL & "'" & arrInfo(0) & "',"
                            '  操作员编号_In 住院费用记录.操作员编号%Type,
                            strSQL = strSQL & "'" & UserInfo.编号 & "',"
                            '  操作员姓名_In 住院费用记录.操作员姓名%Type,
                            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                            '  退费时间_In   住院费用记录.登记时间%Type, --多张单据退费时,每张单据的退费时间相同,都是系统当前时间
                            strSQL = strSQL & "" & strDelDate & ","
                            '  门诊退费_In   Number := 0,
                            strSQL = strSQL & "" & 0 & ","
                            '  入院科室id_In 病人预交记录.科室id%Type,
                            strSQL = strSQL & "" & IIf(lng入院科室ID = 0, "NULL", lng入院科室ID) & ","
                            '  主页id_In     住院费用记录.主页id%Type, --医保入院补充登记时才传入
                            strSQL = strSQL & "" & IIf(lng主页ID = 0, "Null", lng主页ID) & ","
                            '  三方退费_In   Number := 0,
                            strSQL = strSQL & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
                            '  结帐ID_In     住院费用记录.结帐id%Type)
                            strSQL = strSQL & "" & lng结帐ID & ")"
                            Call zlDatabase.ExecuteProcedure(strSQL, "三方卡结算")
                        End If
                    End If
                    str已转结帐ID = str已转结帐ID & "," & arrInfo(2)
                End If
            End If
        Next
        i = i + lngcnt
    Loop
    
    '对补充结算单据进行退费处理
    If strReplenishNos <> "" Then
        strReplenishNos = Mid(strReplenishNos, 2)
        If ExecuteReplenishDel(strReplenishNos, cllReplenishPro, lng主页ID, lng入院科室ID, strOutDelDate) = False Then
            Exit Function
        End If
    End If
    
    '对住院结帐进行销帐处理
    If strJzNOs <> "" Then
        strJzNOs = Mid(strJzNOs, 2)
        If DelBalaceMz(strJzNOs, cllJzPro, strOutDelDate) = False Then
            Exit Function
        End If
    End If
    
     '打印预交款部分
     Call PrintPrePayPrint(frmMain, strOutDelDate)
     If strJzNOs <> "" And mbln立即销帐 = True Then
        '显示结帐窗口
        Call SHowBalanceWindows(strOutDelDate)
     End If
    ExecuteTurn = True
    Exit Function
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMedicare And mbln立即销帐 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, intInsure)
    End If
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function ExecuteReplenishDel(ByVal strNos As String, ByVal cllPro As Collection, _
    ByVal lng主页ID As Long, ByVal lng入院科室ID As Long, ByVal strDelDate As String) As Boolean
    '功能:对补充结算的单据进行转费用及退费处理
    '入参:
    '   strNos 补结算单号,格式：单据号,险类;...
    '   cllPro 传入的退费过程的集合：Array(补结算单据号,转费用SQL)
    '   strDelDate 退费时间
    Dim strSQL As String, strNoTemp As String
    Dim varNos As Variant, i As Long, p As Long, blnTrans As Boolean
    Dim strno As String, intInsure As Integer
    Dim lng结算冲销ID  As Long, lng费用冲销ID As Long, lng结算序号 As Long
    Dim lng原结帐ID As Long, strAdvance As String
    
    Err = 0: On Error GoTo errH
    If strNos = "" Then ExecuteReplenishDel = True: Exit Function
    
    varNos = Split(strNos, ";")
    For i = 0 To UBound(varNos)
        '单据号,险类;...
        strno = Split(varNos(i), ",")(0): intInsure = Split(varNos(i), ",")(1)
        
        lng费用冲销ID = zlDatabase.GetNextId("病人结帐记录")
        lng结算冲销ID = zlDatabase.GetNextId("病人结帐记录")
        lng结算序号 = -1 * lng费用冲销ID
        
        gcnOracle.BeginTrans: blnTrans = True
        For p = 1 To cllPro.Count
            'Array(补结算单据号,转费用SQL)
            strNoTemp = cllPro(p)(0): strSQL = cllPro(p)(1)
            If strNoTemp = strno Then
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        Next
        
        'Zl_门诊转住院_补结算转出(
        strSQL = "Zl_门诊转住院_补结算转出("
        '  No_In         费用补充记录.No%Type,
        strSQL = strSQL & "'" & strno & "',"
        '  费用冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & lng费用冲销ID & ","
        '  结算冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & lng结算冲销ID & ","
        '  结算序号_In     病人预交记录.结算序号%Type,
        strSQL = strSQL & "" & lng结算序号 & ","
        '  退费时间_In   住院费用记录.发生时间%Type,
        strSQL = strSQL & "To_Date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  操作员编号_In 住院费用记录.操作员编号%Type,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In 住院费用记录.操作员姓名%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  主页id_In     病人预交记录.主页id%Type,
        strSQL = strSQL & "" & lng主页ID & ","
        '  入院科室id_In 病人预交记录.科室id%Type,
        strSQL = strSQL & "" & lng入院科室ID & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        'Public Function ClinicDelSwap(lngStlID As Long, Optional ByVal bln退费 As Boolean = True, _
            Optional ByVal intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:将门诊退费的明细和结算数据转发送医保前置服务器确认
            '入参:lngStlID-将要退的费记录的结帐ID；，从预交记录中可以检索医保号和密码
            '     bln退费 -表明是退费交易还是改费交易在调用本接口
            '     strAdvance:格式:冲销ID|补充结算标志|…,每位|分隔
            '           第一位:传入冲销ID,医保可以根据冲销ID来进行取数
            '           第二位:补充结算标志,1-补充结算调和;0非补充结算调用
            '           第三位:NO:当前结算的NO
            '           第四位后: 待以后扩展
            '     注意：
            '           strAdvance在10.34.0以前(不含补允结算)
            '               多单据一次结算时,传入的是原结帐IDs:结帐ID1,结帐ID2,...
            '               其他，传入格式为:退费单据总张数|当前退第几张单据
            '出参:strAdvance:1.原样退回时，返回空
            '                2.退费结算方式与收费结算方式不一致时，返回格式为：结算方式|金额||结算方式|金额||…（其中，金额为负）
            '返回：交易成功返回true；否则，返回false
        strAdvance = lng结算冲销ID & "|1"
        lng原结帐ID = zlGetFromNOToLastBalanceID(strno, , , , True)
        If Not gclsInsure.ClinicDelSwap(lng原结帐ID, True, intInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            MsgBox "医保结算失败，无法继续进行门诊费用转住院操作。", vbInformation, gstrSysName
            Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
    Next
    ExecuteReplenishDel = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Function

Private Function zlGetFromNOToLastBalanceID(ByVal strNos As String, _
    Optional ByVal blnNOMoved As Boolean, _
    Optional ByVal bln历史表同步查 As Boolean = False, _
    Optional lng结算序号 As Long, Optional bln补结算 As Boolean = False) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据一张收费单据的NO，返回最后一次有效的结帐的ID
    '入参:blnNoMoved是否在后备表中，查询单据之前的判断需要用这个参数
    '     bln历史表同步查-是否连接历史表一起查询
    '     bln补结算-是否补充结算
    '出参:lng结算序号-返回最后一次有效的结帐序号
    '返回:结帐ID
    '编制:刘兴洪
    '日期:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String
    Dim i As Long
    
    On Error GoTo errHandle:
    '87975
    strSQL = "With c_单据 As (Select Column_Value As NO From Table(f_Str2list([1])))" & vbNewLine & _
            " Select Max(a.结帐id) As 结帐id" & vbNewLine & _
            " From 门诊费用记录 A, c_单据 M" & vbNewLine & _
            " Where a.No = m.No" & vbNewLine & _
            "       And a.登记时间 + 0 =" & vbNewLine & _
            "           (Select Max(m.登记时间)" & vbNewLine & _
            "            From 门诊费用记录 M, c_单据 J" & vbNewLine & _
            "            Where m.No = j.No And Mod(m.记录性质, 10) = 1 And m.记录状态 In (1, 3) And Nvl(m.费用状态, 0) <> 1)" & vbNewLine & _
            "            And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And Nvl(a.费用状态, 0) <> 1"

    If bln补结算 Then
        strSQL = Replace(strSQL, "门诊费用记录", "费用补充记录")
        strSQL = Replace(strSQL, "Max(a.结帐id)", "Max(a.结算id)")
    End If

    strSQL = "" & _
            "   Select /*+ Rule */ A.结帐ID,B.结算序号 " & _
            "   From (" & strSQL & ") A,病人预交记录 B " & _
            "   Where A.结帐ID=B.结帐ID(+) And Rownum<2"

    If Not blnNOMoved And bln历史表同步查 Then
        strSQL1 = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        strSQL1 = Replace(strSQL, "费用补充记录", "H费用补充记录")
        strSQL1 = Replace(strSQL, "病人预交记录", "H病人预交记录")
        strSQL = strSQL & " Union ALL " & strSQL1
    ElseIf blnNOMoved Then
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        strSQL1 = Replace(strSQL, "费用补充记录", "H费用补充记录")
        strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据单据获取最后一次正常结帐的结帐ID", strNos)

    If rsTemp.EOF Then Exit Function

    lng结算序号 = Val(Nvl(rsTemp!结算序号))
    zlGetFromNOToLastBalanceID = Val(Nvl(rsTemp!结帐ID))
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DelBalaceMz(ByVal strNos As String, ByVal cllPro As Collection, ByVal strDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中输入的负数数量及退回科室是否正确
    '入参:strNos-记帐单号(用逗号分离)
    '        cllPro-传入的记帐销帐过程的集号(过程,"K"+NO)
    '        strDelDate-作废时间
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-03-29 14:01:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strBalance As String, strBalanceNo As String, strBalanceNos As String
    Dim strBalanceIDs As String, i As Long, j As Long, lng结帐ID As Long, intInsure As Integer
    Dim varBalance As Variant, varJz As Variant, varTemp As Variant, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim str未结NOs As String
    
    Err = 0: On Error GoTo errH
    '1.门诊结算
    strSQL = "" & _
    "   Select  /*+ rule */  distinct B.ID as 结帐ID,B.NO as 结帐单,A.NO as 记帐单,C.险类 as 医保" & _
    "   From 门诊费用记录 A, 病人结帐记录 B,保险结算记录 C, Table(f_Str2list([1])) J" & _
    "   Where A.NO = J.Column_Value  " & _
    "           And A.结帐id = B.ID And B.记录状态=1  " & _
    "           And A.结帐ID=C.记录ID(+)  " & _
    "           And C.性质(+)=1 And A.记录性质 In (2, 12) " & _
    "   Order by 结帐单,记帐单"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    str未结NOs = "," & strNos & ","
    With rsTemp
        strBalance = "": strBalanceNos = "": strBalanceIDs = ""
        strBalanceNo = ""
        Do While Not .EOF
               If strBalanceNo <> Nvl(rsTemp!结帐单) Then
                    intInsure = Val(Nvl(!医保)): lng结帐ID = Val(Nvl(rsTemp!结帐ID))
                    strBalanceNo = Nvl(rsTemp!结帐单)
                    strBalanceIDs = strBalanceIDs & "," & lng结帐ID
                    strBalanceNos = strBalanceNos & "," & strBalanceNo
                    strBalance = strBalance & "||" & strBalanceNo & "," & lng结帐ID & "," & intInsure & "|"
               End If
               strBalance = strBalance & "," & Nvl(rsTemp!记帐单)
               str未结NOs = "," & Replace(str未结NOs, "," & Nvl(rsTemp!记帐单) & ",", "") & ","
               .MoveNext
        Loop
        '加入未结或结帐冲销部分的单据
        varTemp = Split(str未结NOs, ",")
        strBalance = strBalance & "||,0,0|"
        For i = 0 To UBound(varTemp)
            If Trim(varTemp(i)) <> "" Then
                strBalance = strBalance & "," & varTemp(i)
            End If
        Next
    End With
    '检查是否存在消费卡结算
    If strBalanceNos <> "" Then strBalanceNos = Mid(strBalanceNos, 2)
    If zlIsExistsSquareCard(strBalanceNos, 2) Then
        '消费卡检查
        MsgBox "在结帐单：" & strBalanceNos & "中存在消费卡，暂不支持对消费卡的门诊转住院费用,请检查!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    '检查是否存在一卡通结算
    If strBalanceIDs <> "" Then strBalanceIDs = Mid(strBalanceIDs, 2)
    Set mrsOneCard = zlGetOneCard(strBalanceIDs)
    If mrsOneCard.RecordCount > 0 Then
        MsgBox "在结帐单：" & strBalanceNos & "中存在一卡通结算，暂不支持门诊转住院费用,请检查!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    Set mrsOneCard = zlGetThreeCard(strBalanceIDs)
    If mrsOneCard.RecordCount > 0 Then
        MsgBox "在结帐单：" & strBalanceNos & "中存在三方卡结算，暂不支持门诊转住院费用,请检查!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    Err = 0: On Error GoTo errH:
    '正式处理结算
    '            Dim varBalance As Variant, varJz As Variant, varTemp As Variant
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    '格式:结帐NO,结帐ID,险类|记帐单1,记帐单2,....||结帐NO1...
    varBalance = Split(strBalance, "||")
    For i = 0 To UBound(varBalance)
        varTemp = Split(Split(varBalance(i), "|")(0), ",")
        varJz = Split(Split(varBalance(i), "|")(1), ",")
        intInsure = Val(varTemp(2)): lng结帐ID = Val(varTemp(1)): strBalanceNo = varTemp(0)
        gcnOracle.BeginTrans: blnTrans = True: blnTransMedicare = False
        '记帐单销帐处理
        For j = 0 To UBound(varJz)
            If varJz(j) <> "" Then
                Call zlDatabase.ExecuteProcedure(cllPro("K" & varJz(j)), "门诊费用转住院-记帐销帐")
            End If
        Next
        '结帐单处理
        If lng结帐ID <> 0 Then
            If DelBalance(strDelDate, strBalanceNo, lng结帐ID, intInsure, blnTransMedicare) = False Then
                    If blnTrans Then gcnOracle.RollbackTrans
                    '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
                    If blnTransMedicare And mbln立即销帐 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, intInsure)
                    Exit Function
            End If
        End If
        gcnOracle.CommitTrans: blnTrans = False
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
    Next
    DelBalaceMz = True
    Exit Function
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMedicare And mbln立即销帐 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, intInsure)
    End If
End Function

Private Function SHowBalanceWindows(ByVal strDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示结帐窗口
    ' 入参:strDelDate-作废日期(主要应用于再次结帐时用预交冲)
    '编制:刘兴洪
    '日期:2011-03-29 17:38:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objInExse As Object
    Dim lng病人ID As Long
   '4.创建结帐部件
    If objInExse Is Nothing Then
        Err = 0: On Error Resume Next
        Set objInExse = CreateObject("zl9InExse.clsFeeQuery")
        If Err <> 0 Then
            MsgBox "注意:" & "在创建住院费用部件时出错,可能该部件未正常注册,结帐失败,请注意重新结帐!", vbInformation, gstrSysName
            SHowBalanceWindows = True
            Exit Function
        End If
    End If
    On Error GoTo errHandle
    'zlPatiBalance(ByVal frmMain As Object, _
    '    ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, strDBUser As String, _
    '    ByVal lng病人ID As Long, ByVal lng主页ID As   long ) as boolean
    lng病人ID = 0
    If mlngPatient <> 0 Then
        lng病人ID = mlngPatient
    ElseIf Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    If objInExse.zlPatiBalance(Me, gcnOracle, glngSys, gstrDBUser, lng病人ID, mlng主页ID, strDelDate) = False Then
        '调用结算
    End If
    SHowBalanceWindows = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ShowBills(ByVal lngPatient As Long, ByVal DatBegin As Date, ByVal DatEnd As Date)
'功能:读取并显示病人指定天数内的门诊费用单据
    Dim i As Long, DatTmp As Date, strSQL As String
    Dim rsList As ADODB.Recordset
    Dim strWhere As String, strInsure As String
    If DatBegin > DatEnd Then
        DatTmp = DatEnd
        DatEnd = DatBegin
        DatBegin = DatTmp
    End If
    If mbln门诊转住院先审核 Then
       strWhere = " And A.病人id = [1] "
       strInsure = " And 病人id = [1] "
    Else
        If DatEnd - DatBegin < 4 Then   '36170
            strWhere = " And A.病人id+0 = [1] And A.发生时间 Between [2] And [3]  "
        Else
            strWhere = " And A.病人id = [1] And A.发生时间+0 Between [2] And [3]  "
        End If
    End If
    strInsure = " And 病人id = [1]  "
    sta.Panels(2).Text = "正在读取收费单据,请稍候 ..."
    Screen.MousePointer = 11
    DoEvents
    Me.Refresh
    On Error GoTo errH
        
   strSQL = strSQL & _
            " Select x.选择, x.类别, x.单据, Decode(Nvl(z.险类, 0), 0, '', '√') As 医保, x.No As 单据号, x.票据号, x.开单人, x.应收金额, x.实收金额, x.发生时间, Max(y.结帐id) As 结帐id," & vbNewLine & _
            "       Nvl(z.险类, 0) As 险类" & vbNewLine & _
            " From (Select '√' As 选择, '可转入' As 类别, '收费单' As 单据, a.No," & vbNewLine & _
            "        a.实际票号 As 票据号, a.开单人, LTrim(To_Char(Sum(a.应收金额), '9999999990.0000')) As 应收金额," & vbNewLine & _
            "        LTrim(To_Char(Sum(a.实收金额), '9999999990.0000')) As 实收金额, To_Char(a.发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间" & vbNewLine & _
            "       From 门诊费用记录 A" & vbNewLine & _
            "       Where Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 " & strWhere & " " & _
                        IIf(mbln门诊转住院先审核, "And Exists (Select 1 From 门诊费用记录 M,费用审核记录 J " & _
            "                                                  Where M.ID=J.费用ID And M.病人ID = [1] and M.NO=A.NO And Mod(M.记录性质,10)=Mod(A.记录性质,10) And " & _
            "                                                       J.审核日期 is Not NULL and  nvl(J.记录状态,0)=0 and J.性质=1) " & vbNewLine, " And Not Exists (Select 1 From 门诊费用记录 M,费用审核记录 J Where M.ID=J.费用ID And M.病人ID = [1] and M.NO=A.NO And Mod(M.记录性质,10)=Mod(A.记录性质,10) And J.审核日期 is Not NULL and  nvl(J.记录状态,0) > 0 and J.性质=1)") & _
            " And Exists" & vbNewLine & _
            "        (Select 1" & vbNewLine & _
            "              From 门诊费用记录 K" & vbNewLine & _
            "              Where k.No = a.No And k.病人id = [1] And Mod(k.记录性质, 10) = Mod(a.记录性质, 10) And Nvl(k.附加标志, 0) <> 9" & vbNewLine & _
            "              Group By k.序号" & vbNewLine & _
            "              Having Sum(k.实收金额) <> 0)" & vbNewLine & _
            "       Group By a.No, a.实际票号, a.开单人, a.发生时间) X, 门诊费用记录 Y," & vbNewLine & _
            "     (Select Distinct 记录id, 险类" & vbNewLine & _
            "       From 保险结算记录" & vbNewLine & _
            "       Where 性质 = 1 " & strInsure & ") Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.记录性质, 10) = 1 And y.记录状态 In (1, 3) And y.病人ID = [1]" & _
            " And y.登记时间 = (Select Max(登记时间) From 门诊费用记录 Where NO = x.No And Mod(记录性质, 10) = 1 And 病人ID = [1] And 记录状态 In (1, 3)) And y.结帐id = z.记录id(+)" & _
            " Group By x.选择, x.类别, x.单据, Decode(Nvl(z.险类, 0), 0, '', '√'), x.No, x.票据号, x.开单人, x.应收金额, x.实收金额, x.发生时间, Nvl(z.险类, 0)  "
 
    If chkShow.Value = 0 Then
        strSQL = strSQL & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select x.选择, x.类别, x.单据, Decode(Nvl(z.险类, 0), 0, '', '√') As 医保, x.No As 单据号, x.票据号, x.开单人, x.应收金额, x.实收金额, x.发生时间, Max(y.结帐id) As 结帐id," & vbNewLine & _
            "       Nvl(z.险类, 0) As 险类" & vbNewLine & _
            "From (Select " & vbNewLine & _
            "        '' As 选择, '不可转入' As 类别, '收费单' As 单据, a.No," & vbNewLine & _
            "        a.实际票号 As 票据号, a.开单人, LTrim(To_Char(Sum(a.应收金额), '9999999990.0000')) As 应收金额," & vbNewLine & _
            "        LTrim(To_Char(Sum(a.实收金额), '9999999990.0000')) As 实收金额, To_Char(a.发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间" & vbNewLine & _
            "       From 门诊费用记录 A" & vbNewLine & _
            "       Where Mod(a.记录性质, 10) = 1 And a.记录状态 = 3 " & strWhere & " And Nvl(a.附加标志, 0) <> 9 And Not Exists" & vbNewLine & _
            "        (Select 1" & vbNewLine & _
            "              From 门诊费用记录 K" & vbNewLine & _
            "              Where k.No = a.No And k.病人id = [1] And Mod(k.记录性质, 10) = Mod(a.记录性质, 10) And Nvl(k.附加标志, 0) <> 9" & vbNewLine & _
            "              Group By k.序号" & vbNewLine & _
            "              Having Sum(k.实收金额) <> 0)" & vbNewLine & _
            "       Group By a.No, a.实际票号, a.开单人, a.发生时间) X, 门诊费用记录 Y," & vbNewLine & _
            "     (Select Distinct 记录id, 险类" & vbNewLine & _
            "       From 保险结算记录" & vbNewLine & _
            "       Where 性质 = 1" & strInsure & ") Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.记录性质, 10) = 1 And y.记录状态 In (1, 3) And y.病人ID = [1]" & _
            " And y.登记时间 = (Select Max(登记时间) From 门诊费用记录 Where NO = x.No And Mod(记录性质, 10) = 1 And 病人ID = [1] And 记录状态 In (1, 3)) And y.结帐id = z.记录id(+)" & _
            " Group By x.选择, x.类别, x.单据, Decode(Nvl(z.险类, 0), 0, '', '√'), x.No, x.票据号, x.开单人, x.应收金额, x.实收金额, x.发生时间, Nvl(z.险类, 0)  "

            
        strSQL = strSQL & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select x.选择, x.类别, x.单据, Decode(Nvl(z.险类, 0), 0, '', '√') As 医保, x.No As 单据号, x.票据号, x.开单人, x.应收金额, x.实收金额, x.发生时间, Max(y.结帐id) As 结帐id," & vbNewLine & _
            "       Nvl(z.险类, 0) As 险类" & vbNewLine & _
            "From (Select " & vbNewLine & _
            "        '' As 选择, '不可转入' As 类别, '收费单' As 单据, a.No," & vbNewLine & _
            "        a.实际票号 As 票据号, a.开单人, LTrim(To_Char(Sum(a.应收金额), '9999999990.0000')) As 应收金额," & vbNewLine & _
            "        LTrim(To_Char(Sum(a.实收金额), '9999999990.0000')) As 实收金额, To_Char(a.发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间" & vbNewLine & _
            "       From 门诊费用记录 A" & vbNewLine & _
            "       Where Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 " & strWhere & " " & _
            " And Exists (Select 1 From 门诊费用记录 M,费用审核记录 J Where M.ID=J.费用ID And M.病人ID = [1] and M.NO=A.NO And Mod(M.记录性质,10)=Mod(A.记录性质,10) And J.审核日期 is Not NULL and  nvl(J.记录状态,0) = 1 and J.性质=1)" & _
            " And Exists" & vbNewLine & _
            "        (Select 1" & vbNewLine & _
            "              From 门诊费用记录 K" & vbNewLine & _
            "              Where k.No = a.No And k.病人id = [1] And Mod(k.记录性质, 10) = Mod(a.记录性质, 10) And Nvl(k.附加标志, 0) <> 9" & vbNewLine & _
            "              Group By k.序号" & vbNewLine & _
            "              Having Sum(k.实收金额) <> 0)" & vbNewLine & _
            "       Group By a.No, a.实际票号, a.开单人, a.发生时间) X, 门诊费用记录 Y," & vbNewLine & _
            "     (Select Distinct 记录id, 险类" & vbNewLine & _
            "       From 保险结算记录" & vbNewLine & _
            "       Where 性质 = 1 " & strInsure & ") Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.记录性质, 10) = 1 And y.记录状态 In (1, 3) And y.病人ID = [1]" & _
            " And y.登记时间 = (Select Max(登记时间) From 门诊费用记录 Where NO = x.No And Mod(记录性质, 10) = 1 And 病人ID = [1] And 记录状态 In (1, 3)) And y.结帐id = z.记录id(+)" & _
            " Group By x.选择, x.类别, x.单据, Decode(Nvl(z.险类, 0), 0, '', '√'), x.No, x.票据号, x.开单人, x.应收金额, x.实收金额, x.发生时间, Nvl(z.险类, 0)  "

    End If
     
    strSQL = strSQL & " UNION ALL " & _
            " Select    '√' as 选择,'可转入' as 类别,'记帐单' as 单据,Decode(NULL,Null,'','√') as 医保, A.NO As 单据号, A.实际票号 As 票据号, A.开单人," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.应收金额), '999999999" & gstrDec & "')) As 应收金额," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.实收金额), '999999999" & gstrDec & "')) As 实收金额," & vbNewLine & _
            "       To_Char(A.发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间, 0 as 结帐ID,0 as 险类" & vbNewLine & _
            " From 门诊费用记录 A" & vbNewLine & _
            " Where A.记录性质 =2 And A.记录状态 <> 0 " & strWhere & vbNewLine & _
            "           And Exists (Select 1 From 门诊费用记录 K Where K.NO=A.NO And K.记录性质=A.记录性质 And K.附加标志 <> 9 Group By K.序号 Having Sum(K.数次) <> 0) " & vbNewLine & _
                        IIf(mbln门诊转住院先审核, "           And Exists(Select 1 From 门诊费用记录 M,费用审核记录 J where M.ID=J.费用ID and M.NO=A.NO And M.记录性质=A.记录性质 And J.审核日期 is Not NULL and  nvl(J.记录状态,0)=0 and J.性质=1) " & vbNewLine, " And Not Exists(Select 1 From 门诊费用记录 M,费用审核记录 J where M.ID=J.费用ID and M.NO=A.NO And M.记录性质=A.记录性质 And J.审核日期 is Not NULL and  nvl(J.记录状态,0) > 0 and J.性质=1) ") & _
            "Group By A.NO, A.实际票号, A.开单人, A.发生时间 "
         
    If chkShow.Value = 0 Then
        strSQL = strSQL & " UNION ALL " & _
            " Select C.选择,C.类别,C.单据,C.医保,C.单据号, C.票据号, C.开单人," & vbNewLine & _
            "       LTrim(To_Char(Sum(D.应收金额), '999999999" & gstrDec & "')) As 应收金额," & vbNewLine & _
            "       LTrim(To_Char(Sum(D.实收金额), '999999999" & gstrDec & "')) As 实收金额," & vbNewLine & _
            "       C.发生时间, C.结帐ID, C.险类" & vbNewLine & _
            " From " & _
            " (Select    '' as 选择,'不可转入' as 类别,'记帐单' as 单据,Decode(NULL,Null,'','√') as 医保, A.NO As 单据号, A.实际票号 As 票据号, A.开单人," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.应收金额), '999999999" & gstrDec & "')) As 应收金额," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.实收金额), '999999999" & gstrDec & "')) As 实收金额," & vbNewLine & _
            "       To_Char(A.发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间,0 as 结帐ID,0 as 险类" & vbNewLine & _
            " From 门诊费用记录  A" & vbNewLine & _
            " Where A.记录性质 = 2 And A.记录状态 In (2,3)  And Not Exists (Select 1 From 门诊费用记录 Where NO=A.NO And 记录状态=1 And 记录性质=2) " & strWhere & vbNewLine & _
            "           And Not Exists (Select 1 From 门诊费用记录 K Where K.NO=A.NO And K.记录性质=A.记录性质 And K.附加标志 <> 9 Group By K.序号 Having Sum(K.实收金额) <> 0) " & vbNewLine & _
            " Group By A.NO, A.实际票号, A.开单人, A.发生时间 Having Sum(A.实收金额)=0) C,门诊费用记录 D Where C.单据号=D.NO And D.记录性质=2 And D.记录状态=3" & vbNewLine & _
            " Group By C.选择,C.类别,C.单据,C.医保,C.单据号, C.票据号, C.开单人,C.发生时间, C.结帐ID, C.险类 "
            
        strSQL = strSQL & " UNION ALL " & _
            " Select    '' as 选择,'不可转入' as 类别,'记帐单' as 单据,Decode(NULL,Null,'','√') as 医保, A.NO As 单据号, A.实际票号 As 票据号, A.开单人," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.应收金额), '999999999" & gstrDec & "')) As 应收金额," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.实收金额), '999999999" & gstrDec & "')) As 实收金额," & vbNewLine & _
            "       To_Char(A.发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间, 0 as 结帐ID,0 as 险类" & vbNewLine & _
            " From 门诊费用记录 A" & vbNewLine & _
            " Where A.记录性质 = 2 And A.记录状态 <> 0 " & strWhere & vbNewLine & _
            "           And Exists (Select 1 From 门诊费用记录 K Where K.NO=A.NO And K.记录性质=A.记录性质 And K.附加标志 <> 9 Group By K.序号 Having Sum(K.数次) <> 0) " & vbNewLine & _
            " And  Exists (Select 1 From 门诊费用记录 M,费用审核记录 J where M.ID=J.费用ID and M.NO=A.NO And M.记录性质=A.记录性质 And J.审核日期 is Not NULL and  nvl(J.记录状态,0) = 1 and J.性质=1) " & _
            "Group By A.NO, A.实际票号, A.开单人, A.发生时间 "
        
    End If
    strSQL = strSQL & "Order By 单据,类别, 票据号 Desc, 单据号 Desc"
   '注意:由于医保要求从最后一张开始退,所以排序很关键
    Set rsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatient, DatBegin, DatEnd)
    mshList.Redraw = flexRDNone: mshList.Clear
    mshList.Rows = 2
    Set mshList.DataSource = rsList
    If rsList.EOF Then
        sta.Panels(2).Text = "没有找到指定时间范围的收费或记帐单据!"
        mshList.Rows = 2
    Else
        sta.Panels(2).Text = "共 " & rsList.RecordCount & " 张收费单据"
    End If
    Call setHeader
    Call SetInsure
    Call SetBillColor
    mshList.Redraw = flexRDBuffered
    Call mshList_AfterRowColChange(0, 0, 1, 0)
    If mshList.Rows >= 2 Then mshList.Select 1, 0
    Screen.MousePointer = 0
    Call SetSumMoney
    Me.Refresh
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetInsure()
    Dim intInsure As Integer, lngRow As Long
    Dim str单据 As String, strno As String
    
    With mshList
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("类别")) = "可转入" And .TextMatrix(lngRow, .ColIndex("选择")) = "√" Then
                intInsure = Val(.TextMatrix(lngRow, .ColIndex("险类")))
                str单据 = Trim(.TextMatrix(lngRow, .ColIndex("单据")))
                If intInsure > 0 And str单据 = "收费单" Then
                    If Not gclsInsure.GetCapability(support门诊结算作废, mlngPatient, intInsure) Then
                        .TextMatrix(lngRow, .ColIndex("选择")) = ""
                    End If
                End If
            End If
        Next lngRow
    End With
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub CheckInsure()
    Dim i As Integer, intInsure As Integer, blnSelect As Boolean
    With mshList
        For i = 1 To .Rows - 1
            intInsure = Val(.TextMatrix(i, .ColIndex("险类")))
            blnSelect = .TextMatrix(i, .ColIndex("选择")) <> ""
            If intInsure > 0 And blnSelect Then
                If gclsInsure.GetCapability(support门诊结算作废, mlngPatient, intInsure) = False Then
                    .TextMatrix(i, .ColIndex("选择")) = ""
                End If
            End If
        Next i
    End With
End Sub

Private Function ExecuteThreeSwap(lngBalance As Long, lng冲销ID As Long, Optional ByRef str交易流水号 As String, Optional ByRef str交易说明 As String) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset, strBalanceIDs As String, rsTotal As ADODB.Recordset
    Dim dblMoney As Double, strAll As String, strDetail() As String, strItem() As String, strCardNO As String
    Dim i As Integer, lngCardID As Long
    
    If mobjSquare Is Nothing Then Set mobjSquare = gobjSquare.objSquareCard
    If mobjSquare Is Nothing Then Exit Function
    strSQL = _
        "Select 摘要" & vbNewLine & _
        "    From 病人预交记录" & vbNewLine & _
        "    Where 结算方式 Is Null And 记录性质 = 3 And 记录状态 = 2 And 结帐id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng冲销ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    strAll = Nvl(rsTemp!摘要)
    If strAll = "" Then Exit Function
    
    strDetail = Split(strAll, "|")
    For i = 0 To UBound(strDetail)
        If strDetail(i) <> "" Then
            strItem = Split(strDetail(i), ",")
            If Val(strItem(0)) = 1 Then
                lngCardID = Val(strItem(1))
                dblMoney = -1 * Val(strItem(2))
                strSQL = "Select Distinct a.结帐id" & vbNewLine & _
                            "From 门诊费用记录 A" & vbNewLine & _
                            "Where a.No In (Select Distinct a.No From 门诊费用记录 A Where Mod(a.记录性质, 10) = 1 And a.结帐id = [1]) And Mod(a.记录性质, 10) = 1 And" & vbNewLine & _
                            "      a.记录状态 <> 0"
                strSQL = "Select Min(结帐ID) As 结帐ID,Min(卡号) As 卡号 From 病人预交记录 Where 结帐ID IN (" & strSQL & ") And 卡类别ID = [2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng冲销ID, lngCardID)
                strBalanceIDs = "3|" & Val(Nvl(rsTemp!结帐ID))
                If mobjSquare.zlReturnCheck(Me, mlngModule, lngCardID, False, Nvl(rsTemp!卡号), _
                    strBalanceIDs, dblMoney, str交易流水号, str交易说明, "3|" & lng冲销ID) = False Then Exit Function
                If mobjSquare.zlReturnMoney(Me, mlngModule, lngCardID, False, Nvl(rsTemp!卡号), _
                    strBalanceIDs, dblMoney, str交易流水号, str交易说明, "3|" & lng冲销ID) = False Then Exit Function
            End If
        End If
    Next i
    
    ExecuteThreeSwap = True
End Function

Private Sub setHeader()
    Dim strHead As String
    Dim i As Long
    With mshList
        If .DataSource Is Nothing Then
            strHead = "选择,4,500|类别,4,850|单据,4,800|医保,4,500|单据号,4,850|票据号,4,1100|开单人,4,800|应收金额,7,850|实收金额,7,850|发生时间,4,1850|结帐ID,4,0|险类,4,0"
            .Cols = UBound(Split(strHead, "|")) + 1
            For i = 0 To UBound(Split(strHead, "|"))
                .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
                .ColKey(i) = Trim(.TextMatrix(0, i))
            Next
            .Rows = 2
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        '选择,4,500|类别,4,850|医保,4,500|单据号,4,850|票据号,4,1100|开单人,4,800|应收金额,7,850|实收金额,7,850|发生时间,4,1850|结帐ID,4,0|险类,4,0
        For i = 0 To .Cols - 1
             .FixedAlignment(i) = flexAlignCenterCenter
             .colAlignment(i) = flexAlignLeftCenter
             .ColKey(i) = Trim(.TextMatrix(0, i))
             Select Case .ColKey(i)
             Case "选择", "类别", "单据", "医保", "单据号", "票据号"
                .colAlignment(i) = flexAlignCenterCenter
             Case "应收金额", "实收金额"
                .colAlignment(i) = flexAlignRightCenter
             End Select
             If .ColKey(i) Like "*ID" Or .ColKey(i) = "险类" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
             End If
        Next
        zl_vsGrid_Para_Restore 1131, mshList, Me.Caption, "门诊转住院列表", True
        .RowHeight(0) = 320
        .Row = 1
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Sub SetBillColor()
    Dim i As Long, j As Long
    With mshList
        For i = 1 To .Rows - 1
            .Row = i
            If .TextMatrix(i, .ColIndex("类别")) = "不可转入" Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H8000000C
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = .ForeColor
            End If
        Next
    End With
End Sub

Private Sub cmdParaSet_Click()
    frmChargeTurnParSet.ShowSet Me, 1131, mstrPrivs
    mbln立即销帐 = Val(zlDatabase.GetPara("费用转出立即退费", glngSys, 1131)) = 1
End Sub

Private Sub cmdSave_Click()
    Dim i As Long, strno As String, strNos As String
    Dim strBalanceID As String, strTemp As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng结帐ID As Long, str单据号 As String, lngInsure As Long
    Dim strReplenishNo As String, strNotSelectNos As String
    Dim varData As Variant, blnErrBill As Boolean
    
    mstrNOS = ""
    If mlngPatient = 0 Then
        MsgBox "未发现病人信息，请检查！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With mshList
        strno = ""
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("类别")) = "可转入" And .TextMatrix(i, .ColIndex("选择")) = "√" Then
                lng结帐ID = Val(.TextMatrix(i, .ColIndex("结帐ID")))
                str单据号 = .TextMatrix(i, .ColIndex("单据号"))
                lngInsure = Val(.TextMatrix(i, .ColIndex("险类")))
                strReplenishNo = "": strNotSelectNos = ""
                
                If InStr(1, "," & strno, "," & str单据号 & ",") = 0 Then
                    strno = strno & "," & str单据号
                End If
                
                If .TextMatrix(i, .ColIndex("单据")) = "收费单" Then
                    If CheckBillExistReplenishData(1, , str单据号, strReplenishNo, blnErrBill) Then
                        If mbln立即销帐 Then
                            If blnErrBill Then
                                MsgBox "单据号为[" & str单据号 & "]的记录已进行医保补充结算，但正处于异常结算状态，不能转出，请先到【保险补充结算】进行处理。", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If CheckReplenishAllNosIsSelected(strReplenishNo, .TextMatrix(i, .ColIndex("单据")), strNotSelectNos) = False Then
                                MsgBox "单据号为[" & str单据号 & "]的记录已进行补充结算，以下单据也必须一起转出：" & vbCrLf & strNotSelectNos, vbInformation, gstrSysName
                                Exit Sub
                            End If
                            '获取医保险类
                            lngInsure = GetReplenishInsure(strReplenishNo)
                            If lngInsure = 0 Then
                                MsgBox "单据号为[" & str单据号 & "]的记录已进行补充结算，但未获取到医保险类,不能转出！", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            '检查医保是否能够原样作废
                            strTemp = CheckInsureCancel(mlngPatient, lngInsure, strReplenishNo, True)
                            If strTemp <> "" Then
                                MsgBox strTemp, vbInformation, gstrSysName
                                Exit Sub
                            End If
                        Else
                            MsgBox "单据号为[" & str单据号 & "]的记录已进行补充结算，不能转出！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                
                If lngInsure <> 0 Then
                    '检查医保单据是否全转出
                    If IsYBSingle(str单据号, lngInsure) = False Then
                        If CheckBalanceAllNosIsSelected(lng结帐ID, .TextMatrix(i, .ColIndex("单据"))) = False Then
                            MsgBox "医保单据号为[" & str单据号 & "]的记录本次未转出全部相关结算单据，不能继续！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                Else
                    If CheckAllTurn(str单据号) Then
                        If CheckBalanceAllNosIsSelected(lng结帐ID, .TextMatrix(i, .ColIndex("单据"))) = False Then
                            MsgBox "单据号为[" & str单据号 & "]的记录本次未转出全部相关结算单据，不能继续！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                
                If .TextMatrix(i, .ColIndex("单据")) = "记帐单" Then
                    If zlIsExistsSquareCard(str单据号, 2) Then
                        '消费卡检查
                        MsgBox "在结帐单：[" & str单据号 & "]中存在消费卡，暂不支持对消费卡的门诊转住院费用,请检查!", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                    strBalanceID = ""
                    strSQL = "Select Distinct A.结帐ID From 门诊费用记录 A,病人结帐记录 B" & _
                            " Where A.结帐ID=B.ID And (b.记录状态=1 or nvl(b.结算状态,0)=1)" & _
                            "       and  Mod(A.记录性质,10)=2 And A.No=[1] "
                    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str单据号)
                    Do While Not rsTemp.EOF
                        strBalanceID = strBalanceID & "," & Nvl(rsTemp!结帐ID)
                        rsTemp.MoveNext
                    Loop
                    '检查是否存在一卡通结算
                    If strBalanceID <> "" Then strBalanceID = Mid(strBalanceID, 2)
                    If strBalanceID <> "" Then
                        Set mrsOneCard = zlGetOneCard(strBalanceID)
                        If mrsOneCard.RecordCount > 0 Then
                            MsgBox "在结帐单：[" & str单据号 & "]中存在一卡通结算，暂不支持门诊转住院费用,请检查!", vbOKOnly + vbInformation, gstrSysName
                            Exit Sub
                        End If
                        Set mrsOneCard = zlGetThreeCard(strBalanceID)
                        If mrsOneCard.RecordCount > 0 Then
                            MsgBox "在结帐单：[" & str单据号 & "]中存在三方卡结算，暂不支持门诊转住院费用,请检查!", vbOKOnly + vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                
                mstrNOS = mstrNOS & ";" & str单据号 & "," & .TextMatrix(i, .ColIndex("票据号")) & "," & _
                    lng结帐ID & "," & lngInsure & "," & .TextMatrix(i, .ColIndex("单据")) & "," & strReplenishNo
            End If
        Next
    End With
    If strno <> "" Then strno = Mid(strno, 2)
    If mstrNOS <> "" Then mstrNOS = Mid(mstrNOS, 2)
        
    If mstrNOS = "" Then
        MsgBox "你还未选择要转成住院费用的单据，不能续继！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '不需要选择病人
    If mblnSelPati = False Then Unload Me: Exit Sub
    
    varData = Split(strno, ","): strno = ""
    For i = 0 To UBound(varData)
        If i > 60 Then strno = strno & ",...": Exit For
        strno = strno & IIf(strno = "", "", ",")
        strno = strno & IIf(i > 0 And i Mod 6 = 0, vbCrLf, "")
        strno = strno & varData(i)
    Next
    If MsgBox("你是否真要将如下门诊费用转成住院费用吗？" & vbCrLf & _
        strno, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        mstrNOS = ""
        Exit Sub
    End If
    
    Err = 0: On Error GoTo ErrHand:
    If Val(Nvl(mrsInfo!主页ID)) = 0 Then
        MsgBox "该病人还未入院，不能门诊费用转住院费用！", vbInformation, gstrSysName
        Exit Sub
    End If
    If ExecuteTurn(Me, mlngModule, mstrPrivs, mstrNOS, Val(Nvl(mrsInfo!住院号)), _
        Val(Nvl(mrsInfo!主页ID)), CDate(Format(mrsInfo!入院日期, "yyyy-mm-dd HH:MM:SS")), _
        Val(Nvl(mrsInfo!入院科室ID)), Val(Nvl(mrsInfo!入院病区ID))) = False Then
        '转换未成功
        Call cmdRefresh_Click
        Exit Sub
    End If
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetReplenishAllNos(ByVal strno As String) As String
    '获取补充结算的所有费用单据
    '返回：
    '   补充结算的所有费用单据:A001,A002,...
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strNos As String
    
    On Error GoTo ErrHandler
    strSQL = _
        " Select Distinct a.No" & vbNewLine & _
        " From 门诊费用记录 A, 门诊费用记录 B, 费用补充记录 C" & vbNewLine & _
        " Where a.No = b.No And a.序号 = b.序号 And a.记录性质 In (1, 11)" & vbNewLine & _
        "       And b.结帐id = c.收费结帐id" & vbNewLine & _
        "       And c.记录性质 = 1 And c.附加标志 = 0 And c.No = [1]" & vbNewLine & _
        " Group By a.No, a.序号" & vbNewLine & _
        " Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strno)
    Do While Not rsTmp.EOF
        strNos = strNos & "," & Nvl(rsTmp!NO)
        rsTmp.MoveNext
    Loop
    If strNos <> "" Then strNos = Mid(strNos, 2)
    
    GetReplenishAllNos = strNos
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckReplenishAllNosIsSelected(ByVal strno As String, ByVal str单据 As String, _
    Optional ByRef strNotSelectNos As String) As Boolean
    '检查补充结算的所有剩余未退费用本次是否都选择了转出
    '入参：
    '   str单据 收费单/记帐单
    '出参：
    '   strNotSelectNos 没有被选择的需要一起转出的单据
    Dim i As Integer, k As Long, blnFind As Boolean
    Dim strNos As String, varNos As Variant
    
    On Error GoTo ErrHandler
    strNotSelectNos = ""
    strNos = GetReplenishAllNos(strno)
    
    varNos = Split(strNos, ",")
    With mshList
        For i = 0 To UBound(varNos)
            blnFind = False
            For k = 1 To .Rows - 1
                If .TextMatrix(k, .ColIndex("单据")) = str单据 And .TextMatrix(k, .ColIndex("单据号")) = varNos(i) Then
                    If .TextMatrix(k, .ColIndex("类别")) = "可转入" And .TextMatrix(k, .ColIndex("选择")) = "√" Then
                        blnFind = True: Exit For
                    End If
                End If
            Next
            
            If blnFind = False Then
                strNotSelectNos = strNotSelectNos & "," & varNos(i)
            End If
        Next
    End With
    
    If strNotSelectNos <> "" Then
        strNotSelectNos = Mid(strNotSelectNos, 2)
        Exit Function
    End If
    CheckReplenishAllNosIsSelected = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetReplenishInsure(ByVal strno As String) As Long
    '获取补充结算的医保险类
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSQL = _
        " Select Max(b.险类) As 险类" & vbNewLine & _
        " From 病人预交记录 A, 保险结算记录 B, 费用补充记录 C" & vbNewLine & _
        " Where a.结帐id = b.记录id And a.记录性质 = 6" & vbNewLine & _
        "       And a.结帐id = c.结算id And c.记录性质 = 1" & vbNewLine & _
        "       And c.记录状态 In(1,3) And c.附加标志 = 0 And c.No = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strno)
    If Not rsTmp.EOF Then GetReplenishInsure = Nvl(rsTmp!险类)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBalanceAllNosIsSelected(ByVal lng结帐ID As Long, ByVal str单据 As String) As Boolean
    '检查一次结算的所有剩余未退费用本次是否都选择了转出
    '入参：
    '   str单据 收费单/记帐单
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHandler
    strSQL = _
        " Select Distinct a.No" & vbNewLine & _
        " From 门诊费用记录 A, 门诊费用记录 B" & vbNewLine & _
        " Where a.No = b.No And Mod(a.记录性质,10) = Mod(b.记录性质,10)" & vbNewLine & _
        "       And a.序号=b.序号 And b.结帐id = [1]" & vbNewLine & _
        " Group By a.No,a.序号" & vbNewLine & _
        " Having Nvl(Sum(Nvl(a.付数,1)*a.数次),0) <> 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
    Do While Not rsTmp.EOF
        With mshList
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("单据")) = str单据 And .TextMatrix(i, .ColIndex("单据号")) = Nvl(rsTmp!NO) Then
                    If Not (.TextMatrix(i, .ColIndex("类别")) = "可转入" And .TextMatrix(i, .ColIndex("选择")) = "√") Then
                        Exit Function
                    End If
                End If
            Next
        End With
        rsTmp.MoveNext
    Loop
    CheckBalanceAllNosIsSelected = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Activate()
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Call picTop_Resize
End Sub

Private Sub Form_Load()
    Dim strTmp As String, Datsys As Date
    
    If Not gobjSquare Is Nothing Then
        Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "")
         '初始化相关的本地数据集
        Set mtySquareCard.rsSquare = New ADODB.Recordset
        mtySquareCard.blnExistsObjects = Not gobjSquare.objSquareCard Is Nothing
        If Not gobjSquare.objSquareCard Is Nothing Then IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        Set mobjSquare = gobjSquare.objSquareCard
    End If
    Call GetRegInFor(g私有模块, Me.Name, "idkind", strTmp)
    mintIDKind = Val(strTmp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
    mstrTitle = Me.Caption
    
    Call RestoreWinState(Me, App.ProductName)
    
    mbln门诊转住院先审核 = IIf(Val(zlDatabase.GetPara("门诊转住院先审核", glngSys, 1143, 0)) = 1, True, False)
    mbln立即销帐 = Val(zlDatabase.GetPara("费用转出立即退费", glngSys, 1131)) = 1
    chkShow.Value = IIf(Val(zlDatabase.GetPara("仅显示可转入数据", glngSys, 1131, 1, Array(chkShow))) = 1, 1, 0)
    picBalance.BorderStyle = 0: picList.BorderStyle = 0:    picBill.BorderStyle = 0
    Call InitPancel
    Datsys = zlDatabase.Currentdate
    strTmp = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "开始时间")
    If IsDate(strTmp) Then
        dtpBegin.Value = CDate(strTmp)
    Else
        dtpBegin.Value = Format(DateAdd("d", -3, Datsys), "yyyy-mm-dd 00:00:00")
    End If
    dtpBegin.MaxDate = Format(Datsys, "yyyy-mm-dd 23:59:59")
    If mstrNOS <> "" Then
        strTmp = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "结束时间")
    Else
        strTmp = ""
    End If
    If IsDate(strTmp) Then
        dtpEnd.Value = CDate(strTmp)
    Else
        dtpEnd.Value = Format(Datsys, "yyyy-mm-dd 23:59:59")
    End If
    Call SetVisibleCtl
    Call setHeader: Call SetDetail: Call SetBalanceHead
    Call zlCreateObject
End Sub

Private Sub SetVisibleCtl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的visible属性
    '编制:刘兴洪
    '日期:2011-03-29 21:49:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    dtpBegin.Visible = Not mbln门诊转住院先审核
    dtpEnd.Visible = Not mbln门诊转住院先审核
    lbl至.Visible = Not mbln门诊转住院先审核
    lblDate.Visible = Not mbln门诊转住院先审核
End Sub

Private Sub cmdExit_Click()
    mstrNOS = ""
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdRefresh_Click()
    Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
    If cmdAll(0).Visible And cmdAll(0).Enabled Then Call cmdAll(0).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    mintIDKind = IDKind.IDKind
    Call SaveRegInFor(g私有模块, Me.Name, "idkind", mintIDKind)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "开始时间", Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss")
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "结束时间", Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss")
    Call SaveWinState(Me, App.ProductName)
    Set mtySquareCard.rsSquare = Nothing
    Call zlDatabase.SetPara("仅显示可转入数据", chkShow.Value, glngSys, 1131)
    zlSaveDockPanceToReg Me, dkpMan, "区域"
    zl_vsGrid_Para_Save 1131, mshDetail, Me.Caption, "门诊转住院明细列表", True
    zl_vsGrid_Para_Save 1131, vsBalance, Me.Caption, "门诊转住院结算列表", True
    zl_vsGrid_Para_Save 1131, mshList, Me.Caption, "门诊转住院列表", True
    Call zlCloseObject
End Sub
 
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.名称 Like "*IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then Call FindPati(objCard, True, Trim(txtPatient.Text))
        Exit Sub
    End If
    
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub

   '    zlReadCard(frmMain As Object, _
   '    ByVal lngModule As Long, _
   '    ByVal lngCardTypeID As Long, _
   '    ByVal blnOlnyCardNO As Boolean, _
   '    ByVal strExpand As String, _
   '    ByRef strOutCardNO As String, _
   '    ByRef strOutPatiInforXML As String) As Boolean
   '    '---------------------------------------------------------------------------------------------------------------------------------------------
   '    '功能:读卡接口
   '    '入参:frmMain-调用的父窗口
   '    '       lngModule-调用的模块号
   '    '       strExpand-扩展参数,暂无用
   '    '       blnOlnyCardNO-仅仅读取卡号
   '    '出参:strOutCardNO-返回的卡号
   '    '       strOutPatiInforXML-(病人信息返回.XML串)
   '    '返回:函数返回    True:调用成功,False:调用失败\
   If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
   txtPatient.Text = strOutCardNO
   If txtPatient.Text <> "" Then Call FindPati(objCard, True, Trim(txtPatient.Text))
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNO As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC卡", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mshDetail_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 1131, mshDetail, Me.Caption, "门诊转住院明细列表", True
End Sub

Private Sub mshDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save 1131, mshDetail, Me.Caption, "门诊转住院明细列表", True
End Sub


Private Sub mshList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 1131, mshList, Me.Caption, "门诊转住院列表", True
End Sub

Private Sub mshList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strno As String, str单据 As String
    
    If NewRow = OldRow Then Exit Sub
    With mshList
        strno = Trim(.TextMatrix(NewRow, .ColIndex("单据号")))
        str单据 = Trim(.TextMatrix(NewRow, .ColIndex("单据")))
        If NewRow = 0 Or strno = "" Then
            mshDetail.Clear 1: mshDetail.Rows = 2
            Call SetDetail
        Else
            Call ShowDetail(str单据, strno)
        End If
        .ForeColorSel = mshList.CellForeColor
    End With
End Sub

Private Sub mshList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save 1131, mshList, Me.Caption, "门诊转住院列表", True
End Sub

Private Sub mshList_DblClick()
    With mshList
        If .MouseRow = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("单据号")) = "" Then Exit Sub
        Call SetRowSelected(.Row, Trim(.TextMatrix(.Row, .ColIndex("选择"))) = "")
    End With
    Call SetSumMoney
    
End Sub
Private Sub mshList_KeyPress(KeyAscii As Integer)
     If KeyAscii <> 32 Then Exit Sub
    With mshList
        If .TextMatrix(.Row, .ColIndex("单据号")) = "" Then Exit Sub
       Call SetRowSelected(.Row, Trim(.TextMatrix(.Row, .ColIndex("选择"))) = "")
    End With
    Call SetSumMoney
End Sub

Private Sub cmdAll_Click(Index As Integer)
    Dim i As Long
    
    With mshList
        .Redraw = False
        For i = 1 To .Rows - 1
            If Not SetRowSelected(i, Index = 0) Then
                .Row = i: .Col = 0: .ColSel = .Cols - 1
                Call mshList_AfterRowColChange(0, 0, .Row, .Col)
                Exit For
            End If
        Next
        .Redraw = True
    End With
    Call SetSumMoney(Index = 1)
End Sub

Private Function CheckInsureCancel(ByVal lng病人ID As Long, ByVal lngInsure As Long, _
    ByVal strno As String, Optional ByVal bln补结算 As Long) As String
    '检查医保是否能够原样作废
    '返回：允许原样作废，则返回空；否则，返回提示信息
    Dim strTmp As String, i As Integer
    Dim arrBalanceType As Variant, strBalanceType As String
    
    On Error GoTo ErrHandler
    If Not gclsInsure.GetCapability(support门诊结算作废, lng病人ID, lngInsure) Then
        CheckInsureCancel = IIf(bln补结算, "医保补充结算", "") & "单据[" & strno & "]的病人险类不支持门诊结算作废，不允许转出！"
        Exit Function
    Else
        '再判断该单据的每种结算方式是否支持,正常退费时,可以退为指定结算方式,此处简化规则为不允许退费
        strTmp = GetBalanceType(strno, bln补结算)
        arrBalanceType = Split(strTmp, ",")
        For i = 0 To UBound(arrBalanceType)
            strBalanceType = arrBalanceType(i)
            If Not gclsInsure.GetCapability(support门诊结算作废, lng病人ID, lngInsure, strBalanceType) Then
                CheckInsureCancel = IIf(bln补结算, "医保补充结算", "") & "单据[" & strno & "]的病人险类不支持" & strBalanceType & "结算作废，不允许转出！"
                Exit Function
            End If
        Next
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SetRowSelected(ByVal lngRow As Long, blnSelect As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置一行的选择状态
    '       如果是多张单据中的一张,则还需同时设置多张中的其它单据
    '编制:刘兴洪
    '日期:2011-02-21 16:10:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strno As String, strTmp As String
    Dim str单据 As String
    
    With mshList
        If .TextMatrix(lngRow, .ColIndex("类别")) = "可转入" And .TextMatrix(lngRow, .ColIndex("选择")) <> IIf(blnSelect, "√", "") Then
            intInsure = Val(.TextMatrix(lngRow, .ColIndex("险类")))
            str单据 = Trim(.TextMatrix(lngRow, .ColIndex("单据")))
            strno = .TextMatrix(lngRow, .ColIndex("单据号"))
            
            If intInsure > 0 And blnSelect And str单据 = "收费单" Then
                strTmp = CheckInsureCancel(mlngPatient, intInsure, strno)
                If strTmp <> "" Then
                    sta.Panels(2).Text = strTmp
                    .TextMatrix(lngRow, .ColIndex("选择")) = ""
                    Exit Function
                End If
            End If
            
            .TextMatrix(lngRow, .ColIndex("选择")) = IIf(blnSelect, "√", "")
            If str单据 = "收费单" Then
                If intInsure > 0 Then      '全部选择或取消
                    If gclsInsure.GetCapability(support多单据收费必须全退, mlngPatient, intInsure) _
                        Or Not IsYBSingle(strno, intInsure) Then
                        If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
                    End If
                Else '现金病人需要处理多单据收费情况
                    If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
                End If
            End If
        End If
        If .TextMatrix(lngRow, .ColIndex("类别")) = "不可转入" Then .TextMatrix(lngRow, .ColIndex("选择")) = ""
    End With
    SetRowSelected = True
End Function

Private Function CheckAllTurn(ByVal strno As String) As Boolean
    Dim strSQL As String, rsData As ADODB.Recordset
    strSQL = "Select 1" & vbNewLine & _
            " From 病人预交记录 A," & vbNewLine & _
            "     (Select Distinct 结帐id" & vbNewLine & _
            "       From 门诊费用记录" & vbNewLine & _
            "       Where NO In (Select Distinct NO" & vbNewLine & _
            "                    From 门诊费用记录" & vbNewLine & _
            "                    Where 结帐id In" & vbNewLine & _
            "                          (Select 结帐id" & vbNewLine & _
            "                           From 病人预交记录" & vbNewLine & _
            "                           Where 结算序号 In (Select b.结算序号" & vbNewLine & _
            "                                          From 门诊费用记录 A, 病人预交记录 B" & vbNewLine & _
            "                                          Where a.No = [1] And a.记录性质 = 1 And a.记录状态 <> 0 And a.结帐id = b.结帐id))) And" & vbNewLine & _
            "             记录性质 = 1 And 记录状态 <> 0) B" & vbNewLine & _
            " Where a.结帐id = b.结帐id And a.记录性质 = 3 And (Exists (Select 1 From 医疗卡类别 Where ID = a.卡类别id And 是否全退 = 1) Or Exists" & vbNewLine & _
            "       (Select 1 From 消费卡类别目录 Where 编号 = a.结算卡序号 And 是否全退 = 1))" & vbNewLine & _
            " Group By 结算方式" & vbNewLine & _
            " Having Sum(冲预交) <> 0"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strno)
    If rsData.EOF Then
        CheckAllTurn = False
    Else
        CheckAllTurn = True
    End If
End Function

Private Function SetMultiOther(ByVal lngRow As Long, blnSelect As Boolean, intInsure As Integer) As Boolean
'功能:多张单据整体选择或取消
'     如果医保多张单据要求整体退费,选择其中一张时,全选多张,取消时全取消
    Dim i As Long, j As Long, k As Long, strno As String, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant, blnAllTurn As Boolean
    Dim str单据 As String, strReplenishNo As String, strNotSelectNos As String
    Dim strNos As String, varNos As Variant
    
    With mshList
        str单据 = .TextMatrix(lngRow, .ColIndex("单据"))
        If str单据 = "记帐单" Then SetMultiOther = True: Exit Function
        If intInsure = 0 Then
            '检查是否为补结算单据
            If CheckBillExistReplenishData(1, , .TextMatrix(lngRow, .ColIndex("单据号")), strReplenishNo) Then
                If mbln立即销帐 Then
                    strNos = GetReplenishAllNos(strReplenishNo)
                    varNos = Split(strNos, ",")
                    For i = 0 To UBound(varNos)
                        For k = 1 To .Rows - 1
                            If .TextMatrix(k, .ColIndex("单据")) = str单据 And .TextMatrix(k, .ColIndex("单据号")) = varNos(i) Then
                                .TextMatrix(k, .ColIndex("选择")) = IIf(blnSelect, "√", "")
                                Exit For
                            End If
                        Next
                    Next
                    SetMultiOther = True
                    Exit Function
                End If
            End If
            
            If CheckAllTurn(.TextMatrix(lngRow, .ColIndex("单据号"))) = True Then
                blnAllTurn = True
            Else
                blnAllTurn = False
            End If
            If gblnMultiBalance Or blnAllTurn Then     '   多单据,多种结算方式
                '33635:原因是多单据且多种结算方式,不能部分退
                strno = ""
                For k = 1 To .Rows - 1
                      If .TextMatrix(k, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) _
                            And .TextMatrix(k, .ColIndex("单据")) = str单据 _
                            And Trim(.TextMatrix(lngRow, .ColIndex("结帐ID"))) <> "" Then
                            strno = strno & "," & .TextMatrix(k, .ColIndex("单据号"))
                      End If
                Next
                If strno <> "" Then strno = Mid(strno, 2)
                If InStr(1, strno, ",") > 0 Then    '证明为多单据
                    '不允许部分退,部分退的话,票据收回存在问题
                    'If CheckSingleBalance(strNO) = False Then    '是多种结算方式,则不允许退费,'全选
                        For k = 1 To .Rows - 1
                              If .TextMatrix(k, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) _
                                And .TextMatrix(k, .ColIndex("单据")) = str单据 _
                                And Trim(.TextMatrix(lngRow, .ColIndex("结帐ID"))) <> "" Then
                                    .TextMatrix(k, .ColIndex("选择")) = IIf(blnSelect, "√", "")
                              End If
                        Next
                    'End If
                End If
            End If
            '检查是否存在消费卡的结算,如果存在,现不支持这部分数据的处理
            If strno = "" Then strno = .TextMatrix(lngRow, .ColIndex("单据号"))
'            If zlIsExistsSquareCard(strNO) Then
'                sta.Panels(2).Text = "暂不支持对消费卡数据的转入!"
'                For k = 1 To .Rows - 1
'                      If .TextMatrix(k, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) _
'                        And .TextMatrix(k, .ColIndex("单据")) = str单据 _
'                        And Trim(.TextMatrix(lngRow, .ColIndex("结帐ID"))) <> "" Then
'                            .TextMatrix(k, .ColIndex("选择")) = ""
'                      End If
'                Next
'            End If
            '检查是否存在消费卡,如果多单据中存在消费卡,也必须全选
            SetMultiOther = True
            Exit Function
        End If
        
        If IsYBSingle(.TextMatrix(lngRow, .ColIndex("单据号")), intInsure) Then SetMultiOther = True: Exit Function
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("类别")) = "可转入" _
                And .TextMatrix(i, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) _
                And i <> lngRow Then
                If .TextMatrix(i, .ColIndex("选择")) <> .TextMatrix(lngRow, .ColIndex("选择")) Then
                   If intInsure <> 0 And blnSelect Then
                        strno = .TextMatrix(i, .ColIndex("单据号"))
                        '判断该单据的每种结算方式是否支持,正常退费时,可以退为指定结算方式,此处简化规则为不允许退费
                         strTmp = GetBalanceType(strno)
                         If strTmp <> "" Then
                             arrBalanceType = Split(strTmp, ",")
                             For j = 0 To UBound(arrBalanceType)
                                 strBalanceType = arrBalanceType(j)
                                 If Not gclsInsure.GetCapability(support门诊结算作废, mlngPatient, intInsure, strBalanceType) Then
                                     sta.Panels(2).Text = "单据[" & strno & "]的病人险类不支持" & strBalanceType & "作废,此行不允许选择转入!"
                                     For k = 1 To .Rows - 1
                                        If .TextMatrix(k, .ColIndex("结帐ID")) = .TextMatrix(i, .ColIndex("结帐ID")) _
                                            And .TextMatrix(k, .ColIndex("单据")) = str单据 Then
                                            .TextMatrix(k, .ColIndex("选择")) = ""
                                        End If
                                     Next
                                     Exit Function
                                 End If
                             Next
                         End If
                    End If
                    .TextMatrix(i, .ColIndex("选择")) = IIf(blnSelect, "√", "")
                End If
            End If
        Next
    End With
    SetMultiOther = True
End Function

Private Function zlIsExistsSquareCard(ByVal strNos As String, Optional int记录性质 As Integer = 3) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查该单据是否为卡结算单据
    '入参:strNos-单据号(可以为多张,用逗号分离)
    '       int记录性质:3-门诊收费;2-结帐
    '出参:
    '返回:存在,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-11 12:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNoIns As String
    
    On Error GoTo errHandle
    
    strNoIns = Replace(strNos, "'", "")
    strSQL = "" & _
    "   Select /*+ rule */ A.ID As 卡结算id " & _
    "   From 病人卡结算记录 A, 病人预交记录 B, 门诊费用记录 C,Table( f_Str2list([1])) J " & _
    "   Where A.结算id = B.ID and B.记录性质=[2] And C.NO = J.Column_Value And C.结帐ID = B.结帐ID And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查收费单是否存在刷卡记录", strNoIns, int记录性质)
    zlIsExistsSquareCard = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlGetOneCard(ByVal strIDs As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查该单据是否一卡通结算单据
    '入参:strIDs-结帐ID(可以为多张,用逗号分离)
    '出参:
    '返回:一卡通结帐数据,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-11 12:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    
     strSQL = "" & _
    "   Select /*+ rule */  A.结帐ID,A.单位帐号, A.结算号码, B.医院编码, A.冲预交 as 金额" & vbNewLine & _
    "   From 病人预交记录 A, 一卡通目录 B,Table( f_Num2list([1])) J " & vbNewLine & _
    "   Where A.结帐id = J.Column_Value  And A.结算方式 = B.结算方式" & _
    "   Order by 结帐ID"
    Set zlGetOneCard = zlDatabase.OpenSQLRecord(strSQL, "获取一卡通结算数据", strIDs)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlGetThreeCard(ByVal strIDs As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查该单据是否三方卡结算单据
    '入参:strIDs-结帐ID(可以为多张,用逗号分离)
    '出参:
    '返回:三方卡结帐数据,则返回true,否则返回False
    '编制:刘尔旋
    '日期:2015-12-29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    
     strSQL = "" & _
    "   Select /*+ rule */  A.结帐ID, A.冲预交 as 金额, B.名称 " & vbNewLine & _
    "   From 病人预交记录 A, 医疗卡类别 B,Table( f_Num2list([1])) J " & vbNewLine & _
    "   Where A.结帐id = J.Column_Value  And A.结算方式 = B.结算方式" & _
    "   Order by 结帐ID"
    Set zlGetThreeCard = zlDatabase.OpenSQLRecord(strSQL, "获取三方卡结算数据", strIDs)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckSingleBalance(ByVal strno As String) As Boolean
'功能：判断指定单据中是否只有一种非医保结算方式(冲预交除外)
'       :strNO(格式为"E01,E02"):问题:34035
'参数：
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strno = Replace(strno, "'", "")
    CheckSingleBalance = True
    
    strSQL = "" & _
    " Select /*+ rule */ Count(Distinct A.结算方式) num" & vbNewLine & _
    " From 病人预交记录 A, 结算方式 B,Table( f_Str2list([1])) J" & vbNewLine & _
    " Where   A.记录性质 = 3 And A.记录状态 In (1, 3) " & _
    "           And A.结算方式 = B.名称 And B.性质 In (1, 2)  And A.NO = J.Column_Value"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strno)
    If rsTmp!num > 1 Then CheckSingleBalance = False
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetBalanceType(ByVal strno As String, _
    Optional ByVal bln补结算 As Boolean) As String
    '功能:获取一张单据中的医保结算方式串
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long
        
    On Error GoTo errH
    If bln补结算 Then
        strSQL = _
            " Select Distinct a.结算方式" & vbNewLine & _
            " From 病人预交记录 A, 结算方式 B, 费用补充记录 C" & vbNewLine & _
            " Where a.结算方式 = b.名称 And a.记录性质 = 6 And b.性质 In(3,4)" & vbNewLine & _
            "       And a.结帐id = c.结算id And c.记录性质 = 1" & vbNewLine & _
            "       And c.附加标志 = 0 And Nvl(c.费用状态, 0) <> 2 And c.No = [1]"
    Else
        strSQL = _
            " Select Distinct a.结算方式" & vbNewLine & _
            " From 病人预交记录 A, 结算方式 B, 门诊费用记录 C" & vbNewLine & _
            " Where a.结算方式 = b.名称 And b.性质 In(3,4)" & vbNewLine & _
            "       And a.结帐id = c.结帐ID And c.记录性质 = 1 And c.No = [1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strno)
    Do While Not rsTmp.EOF
        GetBalanceType = GetBalanceType & "," & rsTmp!结算方式
        rsTmp.MoveNext
    Loop
    GetBalanceType = Mid(GetBalanceType, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowDetail(ByVal str单据 As String, ByVal strno As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示明细单据
    '入参:str单据:收费单(记帐单)
    '        strNO-单据号
    '编制:刘兴洪
    '日期:2011-02-22 11:14:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, i As Long, strSQL As String
    Err = 0: On Error GoTo errH
    If mshList.Row < 0 Then Exit Sub
    
    If mshList.TextMatrix(mshList.Row, mshList.ColIndex("类别")) = "可转入" Then
        strSQL = "Select C.名称 As 类别, Nvl(E.名称, B.名称) As 名称, B.规格, A.计算单位 As 单位, Sum(Nvl(A.付数, 1) * A.数次) As 数量," & vbNewLine & _
                "       LTrim(To_Char(A.标准单价, '999990.00000')) As 单价, LTrim(To_Char(Sum(A.应收金额), '99999" & gstrDec & "')) As 应收金额," & vbNewLine & _
                "       LTrim(To_Char(Sum(A.实收金额), '99999" & gstrDec & "')) As 实收金额, D.名称 As 执行科室, 3 As 记录状态" & vbNewLine & _
                "From 门诊费用记录 A, 收费项目目录 B, 收费项目类别 C, 部门表 D, 收费项目别名 E" & vbNewLine & _
                "Where A.收费细目id = B.ID And A.收费类别 = C.编码 And A.执行部门id = D.ID(+) And A.NO = [1] And Mod(A.记录性质,10) = [2] And" & vbNewLine & _
                "      A.记录状态 In (2,3) And A.收费细目id = E.收费细目id(+) And E.码类(+) = 1 And E.性质(+) = 3 And A.附加标志 <> 9 " & vbNewLine & _
                "Group By A.标准单价,A.序号, C.名称, Nvl(E.名称, B.名称), B.规格, A.计算单位, D.名称 Having Sum(A.数次) <> 0 " & vbNewLine & _
                " Union " & vbNewLine & _
                "Select C.名称 As 类别, Nvl(E.名称, B.名称) As 名称, B.规格, A.计算单位 As 单位, Sum(Nvl(A.付数, 1) * A.数次) As 数量," & vbNewLine & _
                "       LTrim(To_Char(A.标准单价, '999990.00000')) As 单价, LTrim(To_Char(Sum(A.应收金额), '99999" & gstrDec & "')) As 应收金额," & vbNewLine & _
                "       LTrim(To_Char(Sum(A.实收金额), '99999" & gstrDec & "')) As 实收金额, D.名称 As 执行科室, 1 As 记录状态" & vbNewLine & _
                "From 门诊费用记录 A, 收费项目目录 B, 收费项目类别 C, 部门表 D, 收费项目别名 E" & vbNewLine & _
                "Where A.收费细目id = B.ID And A.收费类别 = C.编码 And A.执行部门id = D.ID(+) And A.NO = [1] And Mod(A.记录性质,10) = [2] And" & vbNewLine & _
                "      A.记录状态=1 And A.收费细目id = E.收费细目id(+) And E.码类(+) = 1 And E.性质(+) = 3 And A.附加标志 <> 9 " & vbNewLine & _
                "Group By A.标准单价,A.序号, C.名称, Nvl(E.名称, B.名称), B.规格, A.计算单位, D.名称 Having Sum(A.数次) <> 0 " & vbNewLine
    
    ElseIf mshList.TextMatrix(mshList.Row, mshList.ColIndex("类别")) = "不可转入" Then
        strSQL = "Select C.名称 As 类别, Nvl(E.名称, B.名称) As 名称, B.规格, A.计算单位 As 单位, Sum(Nvl(A.付数, 1) * A.数次) As 数量," & vbNewLine & _
                "       LTrim(To_Char(A.标准单价, '999990.00000')) As 单价, LTrim(To_Char(Sum(A.应收金额), '99999" & gstrDec & "')) As 应收金额," & vbNewLine & _
                "       LTrim(To_Char(Sum(A.实收金额), '99999" & gstrDec & "')) As 实收金额, D.名称 As 执行科室, 2 As 记录状态" & vbNewLine & _
                "From 门诊费用记录 A, 收费项目目录 B, 收费项目类别 C, 部门表 D, 收费项目别名 E" & vbNewLine & _
                "Where A.收费细目id = B.ID And A.收费类别 = C.编码 And A.执行部门id = D.ID(+) And A.NO = [1] And Mod(A.记录性质,10) = [2] And" & vbNewLine & _
                "      A.记录状态 In (1,3) And A.收费细目id = E.收费细目id(+) And E.码类(+) = 1 And E.性质(+) = 3 And A.附加标志 <> 9 " & vbNewLine & _
                "Group By A.标准单价,A.序号, C.名称, Nvl(E.名称, B.名称), B.规格, A.计算单位, D.名称 Having Sum(A.数次) <> 0 " & vbNewLine
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strno, IIf(str单据 = "记帐单", 2, 1))
    
    mshDetail.Redraw = flexRDNone
    mshDetail.Clear
    Set mshDetail.DataSource = rsTmp
    If rsTmp.EOF Then mshDetail.Rows = 2
    Call SetDetail
    mshDetail.Redraw = flexRDBuffered
    Exit Sub
errH:
    mshDetail.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    strHead = "类别,1,650|名称,1,1500|规格,1,1450|单位,4,500|数量,7,500|单价,7,850|应收金额,7,850|实收金额,7,850|执行科室,4,1000|记录状态,4,0"
    With mshDetail
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .colAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        .ColHidden(9) = True
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 9)) = 1 Then .Cell(flexcpForeColor, i, 0, i, 9) = vbBlack
            'If Val(.TextMatrix(i, 9)) = 2 Then .Cell(flexcpForeColor, i, 0, i, 9) = vbRed
            If Val(.TextMatrix(i, 9)) = 3 Then .Cell(flexcpForeColor, i, 0, i, 9) = vbBlue
        Next i
        zl_vsGrid_Para_Restore 1131, mshDetail, Me.Caption, "门诊转住院明细列表", True
        .RowHeight(0) = 320
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Sub SetBalanceHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:结算列表
    '编制:刘兴洪
    '日期:2011-03-28 11:27:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strHead As String
    Dim i As Long
    strHead = "序号,4,650|标志,1,600|结算单号,1,1500|结算金额,7,1000|结算发票,1, 2600"
    With vsBalance
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .colAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        zl_vsGrid_Para_Restore 1131, vsBalance, Me.Caption, "门诊转住院结算列表", True
        .RowHeight(0) = 320
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Sub picBill_Resize()
    Err = 0: On Error Resume Next
    With picBill
        mshList.Left = .ScaleLeft
        mshList.Top = .ScaleTop
        mshList.width = .ScaleWidth
        mshList.Height = .ScaleHeight
    End With
End Sub

Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBalance.Left = .ScaleLeft
        vsBalance.Top = .ScaleTop
        vsBalance.width = .ScaleWidth
        lblSum.Top = .ScaleHeight - lblSum.Height
        vsBalance.Height = lblSum.Top - mshDetail.Top
    End With
End Sub

Private Sub picBottom_Resize()
    Err = 0: On Error Resume Next
    With picBottom
            cmdExit.Left = .ScaleLeft + .ScaleWidth - cmdExit.width - 100
            cmdSave.Left = cmdExit.Left - cmdSave.width - 20
            cmdSave.Top = cmdExit.Top
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        mshDetail.Left = .ScaleLeft
        mshDetail.Top = .ScaleTop
        mshDetail.width = .ScaleWidth
        mshDetail.Height = .ScaleHeight
    End With
End Sub

Private Sub picTop_Resize()
    Err = 0: On Error Resume Next
    If mblnSelPati Then
        fraPati.Left = picTop.ScaleLeft
        lblDate.Left = fraPati.Left + fraPati.width + 20
        dtpBegin.Left = lblDate.Left + lblDate.width + 10
        lbl至.Left = dtpBegin.Left + dtpBegin.width + 20
        dtpEnd.Left = lbl至.Left + lbl至.width + 20
    End If
    chkShow.Left = IIf(dtpEnd.Visible, dtpEnd.Left + dtpEnd.width, (fraPati.Left + fraPati.width) * IIf(fraPati.Visible = False, 0, 1) + 50)
    cmdRefresh.Left = chkShow.Left + chkShow.width + 50
End Sub

Private Sub txtPatient_Change()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If Not mobjIDCard Is Nothing And Not txtPatient.Locked Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing And Not txtPatient.Locked Then Call mobjICCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not txtPatient.Locked Then Call IDKind.SetAutoReadCard(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, strInput As String
    If txtPatient.Locked Then Exit Sub
    '病人选择器
    If Not (Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13) Then
       If IDKind.GetCurCard.名称 Like "姓名*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        ElseIf IDKind.GetCurCard.名称 = "门诊号" Or IDKind.GetCurCard.名称 = "住院号" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
        Else
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
        End If
    End If
    
    Me.Refresh
    '刷卡完毕或输入号码后回车
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtPatient.Text))
    End If
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-10-18 16:35:27
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If GetPatient(objCard, strInput, blnCard) Then
        '69526:刘尔旋,2014-02-13,出院病人无法进行门诊转住院操作
        If Val(zlDatabase.GetPara("出院病人允许门诊转住院", glngSys, 1137, "0")) = 0 Then
            If HaveOut(mlngPatient) = True Then
                MsgBox "病人" & mrsInfo!姓名 & "已经出院或还未办理住院，不允许进行门诊费用转住院操作！", vbInformation, gstrSysName
                txtPatient.Text = "": mlngPatient = 0
                Call ClearData
                Set mrsInfo = Nothing
                If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
                Exit Sub
            End If
        End If
        '此时会先隐式调用事件Form_Load
        Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
        If mshList.TextMatrix(1, mshList.ColIndex("单据号")) <> "" Then
            If mshList.TextMatrix(1, mshList.ColIndex("选择")) <> "" Then
                If cmdSave.Visible And cmdSave.Enabled Then Call cmdSave.SetFocus
            Else
                If cmdAll(0).Visible And cmdAll(0).Enabled Then Call cmdAll(0).SetFocus
            End If
        Else
            If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
        End If
    Else
        txtPatient.Text = "": mlngPatient = 0
        Call ClearData
        If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
    End If
    txtPatient.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    Call IDKind.SetAutoReadCard(False)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If txtPatient.Text <> mrsInfo!姓名 Then txtPatient.Text = mrsInfo!姓名
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card '54894
    Set objCard = IDKind.GetIDKindCard("身份证", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人信息
    '入参:blnCard=是否就诊卡刷卡,lng主页ID=读取指定住院次数的病人信息
    '出参:
    '返回:是否读取成功,成功时mrsInfo中包含病人信息,失败时mrsInfo=Close,strInput返回是用来判断是否已提示过,避免再次提示没有找到病人
    '编制:刘兴洪
    '日期:2010-11-09 17:17:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim strPati As String, strRange As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    strSQL = _
    " Select A.病人ID,Nvl(B.主页ID,0) as 主页ID,A.住院号,A.当前床号,B.入院病区ID,B.入院科室ID,B.出院病床," & _
    "        Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别,Nvl(B.年龄,A.年龄) as 年龄,A.IC卡号,A.就诊卡号,A.卡验证码," & _
    "       Nvl(B.费别,A.费别) as 费别,C.名称 as 当前科室,A.当前科室ID,D.名称 as 出院科室,B.出院科室ID,A.险类 as 险类,E.卡号,E.医保号,E.密码," & _
    "       A.登记时间,Nvl(B.状态,0) as 状态,Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,Nvl(B.审核标志,0) as 审核标志,B.入院日期,B.出院日期,B.病人性质,B.病人类型" & _
    " From 病人信息 A,病案主页 B,部门表 C,部门表 D,医保病人档案 E,医保病人关联表 F" & _
    " Where A.停用时间 is NULL And A.病人ID=B.病人ID(+) And " & IIf(lng主页ID = 0, "A.主页ID=B.主页ID(+)", "B.主页ID=[3]") & _
    "           And A.病人ID=F.病人ID(+) And F.标志(+)=1 And F.医保号=E.医保号(+) And F.险类=E.险类(+) And F.中心 = E.中心(+)" & _
    "           And A.当前科室ID=C.ID(+) And B.出院科室ID=D.ID(+) "
        
    If blnCard = True And objCard.名称 Like "姓名*" Then  '刷卡
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSQL = strSQL & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strSQL = strSQL & " And A.门诊号=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        strSQL = strSQL & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [1])"
    Else '当作姓名
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                If Not mrsInfo Is Nothing Then
                    If mrsInfo.State = 1 Then
                        If mrsInfo!姓名 = Trim(txtPatient.Text) Then
                            mlngPatient = Val(Nvl(mrsInfo!病人ID))
                            GetPatient = True
                            Exit Function
                        End If
                    End If
                End If
                If mintPatientRange > 0 Then
                    Select Case mintPatientRange
                        Case 1  '任何费用未结清病人
                            strRange = ""
                        Case 2  '体检未结清的病人
                            strRange = " And C.来源途径 = 4"
                        Case 3  '住院未结清的病人
                            strRange = " And C.来源途径 = 2"
                        Case 4  '门诊未结清的病人
                            strRange = " And C.来源途径 = 1"
                    End Select
                    strPati = " And Exists(Select 1 From 病人未结费用 C Where C.病人id=A.病人ID And Nvl(C.主页ID,0)=A.主页ID" & strRange & ")"
                End If
                 '通过姓名查找
                strPati = "Select A.病人ID as ID,A.病人ID,A.住院号, A.门诊号, Nvl(b.性别, a.性别) As 性别, Nvl(b.年龄, a.年龄) as 年龄, A.住院次数, A.家庭地址, A.工作单位," & vbNewLine & _
                        "To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,  To_Char(B.入院日期,'YYYY-MM-DD') as 入院日期, To_Char(B.出院日期,'YYYY-MM-DD') as 出院日期" & vbNewLine & _
                        "From 病人信息 A, 病案主页 B,在院病人 C" & vbNewLine & _
                        "Where A.病人id = B.病人id(+) And A.主页ID = B.主页id(+) And A.停用时间 Is Null And A.病人ID=C.病人ID And A.姓名 = [1] " & vbNewLine & strPati & vbNewLine & _
                        "Order By Decode(住院号, Null, 1, 0), 入院日期 Desc"
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput)
                            
                If Not mrsInfo Is Nothing Then
                    strInput = Val(mrsInfo!病人ID)
                    strSQL = strSQL & " And A.病人ID=[2]"
                Else
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
                End If
            Case "医保号"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.医保号=[2]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.门诊号=[2]"
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, lng主页ID)
    If mrsInfo.EOF Then Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
    
    txtPatient.Text = Nvl(mrsInfo!姓名): mlngPatient = Val(Nvl(mrsInfo!病人ID))
    If IsDate(Format(mrsInfo!入院日期, "yyyy-mm-dd HH:MM:SS")) Then
        '最大设置为入院日期,不能转入住院过程中的门诊费用
        dtpEnd.MaxDate = CDate(Format(mrsInfo!入院日期, "yyyy-mm-dd 23:59:59"))
        dtpEnd.Value = dtpEnd.MaxDate
        dtpEnd.MaxDate = dtpEnd.MaxDate + 1
        dtpBegin.MaxDate = dtpEnd.Value
        '   问题: 36609比入院时间要多一天,因为可能存在病人在没有门诊结算时,先入院,再去门诊结算,从而造成门诊费用转不了的情况.
    End If
    GetPatient = True
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function
  
Private Function PrintPrePayPrint(ByVal frmMain As Object, ByVal strDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印预交款
    '入参:strDelDate-本次转出日期
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-16 10:30:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, bytPrepayPrint As Byte
    Dim strNos As String
    
    On Error GoTo errHandle
    If InStr(1, mstrPrivs, ";预交款收据打印;") = 0 Then
       PrintPrePayPrint = True: Exit Function '不打印
    End If
    bytPrepayPrint = Val(zlDatabase.GetPara("门诊转住院预交打印", glngSys, 1131))
    If bytPrepayPrint = 0 Then PrintPrePayPrint = True: Exit Function '不打印
    
    strSQL = "Select distinct NO From 病人预交记录 Where 记录性质=1 and 收款时间= [1] and 摘要='门诊转住院预交'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取转预交单", CDate(strDelDate))
    If rsTemp.EOF Then
        '没有转为预交数据，则也不打印
        PrintPrePayPrint = True: Exit Function
    End If
    If bytPrepayPrint = 2 Then   '提示打印
        If MsgBox("本次门诊费用转住院费用时，存在现金等结算方式转为了预交款,您是否要打印预交款票据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
              PrintPrePayPrint = True: Exit Function
        End If
    End If
    
    If Val(zlDatabase.GetPara(283, glngSys, , "0")) = 1 Then '112862
        Do While Not rsTemp.EOF
            strNos = strNos & "," & Nvl(rsTemp!NO)
            rsTemp.MoveNext
        Loop
        If strNos <> "" Then strNos = Mid(strNos, 2)
        If zlPrintInvoice(strNos, strDelDate) = False Then Exit Function
    Else
        With rsTemp
            Do While Not .EOF
                If zlPrintInvoice(Nvl(rsTemp!NO), strDelDate) = False Then Exit Function
                .MoveNext
            Loop
        End With
    End If
    PrintPrePayPrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetSumMoney(Optional blnCls As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置和显示合计
    '编制:刘兴洪
    '日期:2011-03-04 14:17:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblSumMoney As Double
    Dim strJzNOs As String, strSFNos As String
    With mshList
        If blnCls = False Then
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("选择"))) <> "" Then
                    dblSumMoney = dblSumMoney + Val(.TextMatrix(i, .ColIndex("实收金额")))
                End If
                If .TextMatrix(i, .ColIndex("类别")) = "可转入" And .TextMatrix(i, .ColIndex("选择")) = "√" Then
                    If .TextMatrix(i, .ColIndex("单据")) = "记帐单" Then
                        strJzNOs = strJzNOs & "," & .TextMatrix(i, .ColIndex("单据号"))
                    Else
                        strSFNos = strSFNos & "," & .TextMatrix(i, .ColIndex("单据号"))
                    End If
                End If
            Next
        Else
            dblSumMoney = 0
        End If
    End With
    lblSum.Caption = "本次转出合计:" & Format(dblSumMoney, "###0.00;-###0.00;0.00;0.00")
    '加载选择的数据通信
    Call LoadBalance(strJzNOs, strSFNos)
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Bottom = sta.Height + picBottom.Height + 100
End Sub

Private Sub LoadBalance(ByVal strJzNOs As String, ByVal strSFNos As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载结算信息
    '编制:刘兴洪
    '日期:2011-03-28 11:33:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, i As Long
    If strJzNOs = "" And strSFNos = "" Then
        With mshList
            If .TextMatrix(i, .ColIndex("类别")) = "可转入" And .TextMatrix(i, .ColIndex("选择")) = "√" Then
                If .TextMatrix(i, .ColIndex("单据")) = "记帐单" Then
                    strJzNOs = strJzNOs & "," & .TextMatrix(i, .ColIndex("单据号"))
                Else
                    strSFNos = strSFNos & "," & .TextMatrix(i, .ColIndex("单据号"))
                End If
            End If
        End With
    End If
    If strJzNOs = "" Then strJzNOs = ",lxh"
    strJzNOs = Mid(strJzNOs, 2)
    If strSFNos = "" Then strSFNos = ",lxh"
    strSFNos = Mid(strSFNos, 2)
    
    On Error GoTo errHandle
    '将:Wmsys.Wm_Concat改为了f_List2Str(Cast(collect ()))的方式.原因是oracle10g目前只是测试版
    '问题:38528
    
    strSQL = "" & _
    "     Select /*+ rule */  Rownum As 序号, 标志, NO As 结算单号, 结算金额, 发票号 " & _
    "     From (Select A.标志, A.NO, A.结算金额, f_List2str(Cast(COLLECT(distinct C.号码) as t_Strlist))  As 发票号 " & _
    "            From (Select '收费' As 标志, A.NO, To_Char(Sum(a.结帐金额),'9999990.00') As 结算金额 " & _
    "                   From 门诊费用记录 A, Table(f_Str2list([1])) J " & _
    "                   Where A.NO = J.Column_Value And Mod(A.记录性质,10) = 1 " & _
    "                   Group By A.NO) A, 票据打印内容 B, 票据使用明细 C " & _
    "            Where A.NO = B.NO(+) and B.数据性质(+)=1 And B.ID = C.打印id(+) " & _
    "            And C.性质(+)=1 " & _
    "            Group By A.标志, A.NO, A.结算金额 " & _
    "            Union All " & _
    "            Select A.标志, A.NO, A.结算金额, f_List2str(Cast(COLLECT(distinct C.号码) as t_Strlist)) As 发票号 " & _
    "            From (Select '结帐' As 标志, B.NO, To_Char(Sum(a.结帐金额),'9999990.00') As 结算金额 " & _
    "                   From 门诊费用记录 A, 病人结帐记录 B, Table(f_Str2list([2])) J " & _
    "                   Where A.NO = J.Column_Value  And A.结帐id = B.ID  And B.记录状态=1 And A.记录性质 In (2, 12) " & _
    "                   Group By B.NO) A, 票据打印内容 B, 票据使用明细 C " & _
    "            Where A.NO = B.NO(+) and B.数据性质(+)=3 And B.ID = C.打印id(+) " & _
    "            Group By A.标志, A.NO, A.结算金额)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSFNos, strJzNOs)
    Set vsBalance.DataSource = rsTemp
    If rsTemp.RecordCount = 0 Then
        vsBalance.Rows = 2
    End If
    Call SetBalanceHead
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function DelBalance(ByVal strDelDate As String, ByVal strno As String, ByVal lng结帐ID As Long, _
    ByVal intInsure As Integer, Optional ByRef blnTransMC As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:结帐作废
    '入参:strNO-结帐单据号
    '       strDelDate:作废时间
    '出参:blnTransMC-医保结算
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-03-29 11:58:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strAdvance As String
    
    '结帐作废
     '  Zl_门诊转住院结帐_结帐作废
     strSQL = "Zl_门诊转住院结帐_结帐作废("
     '  No_In         病人结帐记录.NO%Type,
     strSQL = strSQL & "'" & strno & "',"
     '  操作员编号_In 病人结帐记录.操作员编号%Type,
     strSQL = strSQL & "'" & UserInfo.编号 & "',"
     '  操作员姓名_In 病人结帐记录.操作员姓名%Type
     strSQL = strSQL & "'" & UserInfo.姓名 & "',"
     '    作废日期_In   病人结帐记录.收费时间%Type
     strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi;ss'))"
     Call zlDatabase.ExecuteProcedure(strSQL, "门诊费用转住院-结帐作废")
    '保险接口
    blnTransMC = False
    If intInsure <> 0 Then
        If gclsInsure.CheckInsureValid(intInsure) = False Then
             Exit Function
        End If
        If gclsInsure.GetCapability(support门诊结算作废, , intInsure) Then
            strAdvance = "1|1"
            If Not gclsInsure.ClinicDelSwap(lng结帐ID, , intInsure, strAdvance) Then
                Exit Function
            Else
                blnTransMC = True
            End If
        Else
            MsgBox "单据(" & strno & ")包含不支持结算作废的医保结算，无法进行门诊费用转住院操作！", vbInformation, gstrSysName
            Exit Function
        End If
  End If

'一卡通，暂不处理
'    ElseIf Not rsOneCard Is Nothing Then
'        If rsOneCard.RecordCount > 0 Then
'            If Not objICCard.ReturnSwap(rsOneCard!单位帐号, rsOneCard!医院编码, "" & rsOneCard!结算号码, rsOneCard!金额) Then
'                MsgBox "一卡通退费交易调用失败！退费操作失败！", vbExclamation, gstrSysName
'                exit function
'            End If
'        End If
'    '4.卡结算处理,暂不处理
'    If zlCallSquare_DelFree(lng结帐ID) = False Then
'        '如果发生错了,在过程中就回退了
'                exit function
'    End If
    DelBalance = True
End Function

Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 1131, vsBalance, Me.Caption, "门诊转住院结算列表", True
End Sub

Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save 1131, vsBalance, Me.Caption, "门诊转住院结算列表", True
End Sub

Private Function zlPrintInvoice(ByVal strNos As String, ByVal strDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发票处理
    '入参：
    '   strNos 本次打印预交单据号，格式：A001,A002,A003,...
    '编制:刘兴洪
    '日期:2011-04-02 09:48:13
    '问题:36984
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngShareUseID As Long, lng领用ID As Long, strInvoice As String
    Dim blnInput As Boolean, blnValid As Boolean
    Dim strSQL As String
    Dim intInvoiceFormat As Integer
    
    '如果严格控制票据使用
    On Error GoTo errHandle
    If gblnPrepayStrict Then
        lngShareUseID = zlDatabase.GetPara("共用预交票据批次", glngSys, 1131, 0)
        '1.严格控制票据时，根据实际的票据张数,重新检查领用ID和票据号
        lng领用ID = GetInvoiceGroupID(2, 1, lng领用ID, lngShareUseID, strInvoice, "2")
        If lng领用ID <= 0 Then
            Select Case lng领用ID
                Case -1
                    MsgBox "预交单据[" & strNos & "]共需要1张票据!" & vbCrLf & _
                        "你没有足够的自用和共用的票据,请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "单据[" & strNos & "]共需要1张票据!" & vbCrLf & _
                        "你没有足够的的共用票据,请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -3
                    MsgBox "单据[" & strNos & "]共需要1张票据!" & vbCrLf & _
                        "票据号[" & strInvoice & "]不在可用领用批次的有效票据号范围内！" & _
                        "请重新输入有效的票据号后重打该单据！", vbInformation, gstrSysName
                Case -4
                    MsgBox "单据[" & strNos & "]共需要1张票据!" & vbCrLf & _
                        "票据号[" & strInvoice & "]所在的领用批次没有足够的票据！" & _
                        "请先打印其它票据,用完当前领用批次后,重打该单据！", vbInformation, gstrSysName
                Case Else
                    MsgBox "票据领用信息访问失败！将来，你可以重打单据[" & strNos & "]", vbInformation, gstrSysName
            End Select
            Exit Function
        End If
        Do
            '根据票据领用读取
            blnInput = False
            strInvoice = GetNextBill(lng领用ID)
            If strInvoice = "" Then
                '如果中途换用靠后的号码,可能造成未用完,但下一号码已超出范围
                strInvoice = UCase(InputBox("无法根据票据领用情况获取将要使用的开始票据号，" & _
                                vbCrLf & "请你输入将要使用的开始票据号码：", gstrSysName, _
                                "", Me.Left + 1500, Me.Top + 1500))
                blnInput = True
            Else
                strInvoice = UCase(InputBox("请确认重打使用的开始票据号码：", gstrSysName, _
                                strInvoice, Me.Left + 1500, Me.Top + 1500))
                blnInput = True
            End If
            
            '用户取消输入,不打印
            If strInvoice = "" Then Exit Function
            '检查输入有效性
            If blnInput Then
                If GetInvoiceGroupID(2, 1, lng领用ID, lngShareUseID, strInvoice, "2") = -3 Then
                    MsgBox "你输入的票据号码不在当前领用批次的有效领用范围内,请重新输入！", vbInformation, gstrSysName
                Else
                    blnValid = True
                End If
            Else
                blnValid = True
            End If
        Loop While Not blnValid
    Else
        '有可能是第一次使用
         Do
             blnInput = False
             '非严格控制时直接从本地读取
             strInvoice = UCase(zlDatabase.GetPara("当前预交票据号", glngSys, 1131, ""))
             If strInvoice = "" Then
                 strInvoice = UCase(InputBox("没有找到已用的最大票据号码，无法确定将要使用的开始票据号。" & _
                                 vbCrLf & "请输入将要使用的开始票据号码：", gstrSysName, _
                                 "", Me.Left + 1500, Me.Top + 1500))
                 blnInput = True
             Else
                 strInvoice = zlCommFun.IncStr(strInvoice)
                 strInvoice = UCase(InputBox("请确认重打使用的开始票据号码：", gstrSysName, _
                                 strInvoice, Me.Left + 1500, Me.Top + 1500))
                 blnInput = True
             End If
                 
             '用户取消输入,允许打印
             If strInvoice = "" Then
                 If MsgBox("你确定不输入票据号继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                 blnValid = True
             Else
                 '检查输入有效性
                 If blnInput Then
                     If zlCommFun.ActualLen(strInvoice) <> gbytPrepayLen Then
                         MsgBox "输入的票据号码长度应该为 " & gbytPrepayLen & " 位！", vbInformation, gstrSysName
                     Else
                         blnValid = True
                     End If
                 Else
                     blnValid = True
                 End If
             End If
         Loop While Not blnValid
    End If
    
    '执行数据处理
    'Zl_病人预交记录_Reprint
    strSQL = "Zl_病人预交记录_Reprint("
    '  单据号_In Varchar2,
    strSQL = strSQL & "'" & strNos & "',"
    '  票据号_In 票据使用明细.号码%Type,
    strSQL = strSQL & "'" & strInvoice & "',"
    '  领用id_In 票据使用明细.领用id%Type,
    strSQL = strSQL & "" & IIf(lng领用ID = 0, "NULL", lng领用ID) & ","
    '  使用人_In 票据使用明细.使用人%Type
    strSQL = strSQL & "'" & UserInfo.姓名 & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '输出票据
    intInvoiceFormat = Val(zlDatabase.GetPara(284, glngSys, , "0"))
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, _
        "NO=" & strNos, "收款时间=" & Format(strDelDate, "yyyy-mm-dd HH:MM:SS"), _
        "病人ID=" & mlngPatient, IIf(intInvoiceFormat = 0, "", "ReportFormat=" & intInvoiceFormat), 2)
    
    '更新本地票据
    If Not gblnPrepayStrict Then
        zlDatabase.SetPara "当前预交票据号", strInvoice, glngSys, 1131
    End If
    zlPrintInvoice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Private Sub zlCreateObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建公共事件对象
    '返回: 创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-08-28 16:16:00
    '说明:
    '问题:54894
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '创建公共对象
    Err = 0: On Error Resume Next
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
         Set mobjICCard.gcnOracle = gcnOracle
    End If
    
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    
End Sub
Private Sub zlCloseObject()
    '关闭相关对象
    Err = 0: On Error Resume Next
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
    End If
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
    End If
    Set mobjIDCard = Nothing
    Set mobjICCard = Nothing
End Sub
