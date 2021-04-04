VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReplenishTheBalanceDel 
   AutoRedraw      =   -1  'True
   Caption         =   "保险补充结算"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   Icon            =   "frmReplenishTheBalanceDel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic退费摘要 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   11265
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4590
      Width           =   11265
      Begin VB.TextBox txt退费摘要 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1035
         MaxLength       =   100
         TabIndex        =   4
         Top             =   45
         Width           =   5820
      End
      Begin VB.Label lbl摘要 
         AutoSize        =   -1  'True
         Caption         =   "退费摘要"
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
         Left            =   45
         TabIndex        =   3
         Top             =   120
         Width           =   960
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   11265
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7404
      Width           =   11265
      Begin VB.CommandButton cmdBillSel 
         Caption         =   "全选当前单据(&B)"
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
         Left            =   3240
         TabIndex        =   23
         ToolTipText     =   "热键：Ctrl+B"
         Top             =   135
         Width           =   2040
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   9375
         TabIndex        =   11
         Top             =   135
         Width           =   1440
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "全清(&R)"
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
         Left            =   1695
         TabIndex        =   18
         ToolTipText     =   "热键：Ctrl+R"
         Top             =   135
         Width           =   1440
      End
      Begin VB.CommandButton cmdSelAll 
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
         Left            =   165
         TabIndex        =   17
         ToolTipText     =   "热键：Ctrl+A"
         Top             =   135
         Width           =   1440
      End
      Begin VB.CommandButton cmdOK 
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
         Left            =   7845
         TabIndex        =   10
         Top             =   144
         Width           =   1440
      End
      Begin VB.Line LineCmd_1 
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   8064
      Width           =   11268
      _ExtentX        =   19870
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmReplenishTheBalanceDel.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12224
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "误差"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VSFlex8Ctl.VSFlexGrid vsBill 
      Align           =   1  'Align Top
      Height          =   3630
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   11265
      _cx             =   19870
      _cy             =   6403
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReplenishTheBalanceDel.frx":0E1E
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picMoney 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   11265
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5445
      Width           =   11265
      Begin VB.TextBox txt退款合计 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   9984
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   75
         Width           =   1275
      End
      Begin VB.TextBox txtAllTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3465
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   75
         Width           =   1275
      End
      Begin VB.TextBox txtCurTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   75
         Width           =   1275
      End
      Begin VB.Label lbl退款合计 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "退款合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   9000
         TabIndex        =   25
         Top             =   132
         Width           =   960
      End
      Begin VB.Label lblAllTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收费合计"
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
         Left            =   2460
         TabIndex        =   8
         Top             =   132
         Width           =   960
      End
      Begin VB.Label lblCurTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前单据"
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
         Left            =   75
         TabIndex        =   6
         Top             =   135
         Width           =   960
      End
   End
   Begin VB.PictureBox picPati 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   11265
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   11265
      Begin zlIDKind.IDKindNew IDKindNO 
         Height          =   300
         Left            =   7725
         TabIndex        =   29
         Top             =   120
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
         ShowSortName    =   0   'False
         IDKindStr       =   "收|收费单号|0|0|0|0|0|0;发|发票号|0|0|0|0|0|0;挂|挂号单号|0|0|0|0|0|0;结|结算单号|0|0|0|0|0|0"
         CaptionAlignment=   1
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
         DefaultCardType =   "0"
         NotAutoAppendKind=   -1  'True
         AllowAutoCommCard=   0   'False
         BackColor       =   -2147483633
      End
      Begin VB.PictureBox picPatiBack 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   540
         ScaleHeight     =   360
         ScaleWidth      =   2640
         TabIndex        =   26
         Top             =   525
         Width           =   2640
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
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   645
            MaxLength       =   100
            TabIndex        =   1
            ToolTipText     =   "定位:F6,输入:-病人ID,*门诊号,+住院号,.挂号单号,例如:*2536表示按门诊号查找"
            Top             =   -15
            Width           =   1980
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   360
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   635
            Appearance      =   2
            IDKindStr       =   "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;就|就诊卡|0"
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
            MustSelectItems =   "姓名"
            BackColor       =   -2147483633
         End
      End
      Begin VB.PictureBox pic退 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   10470
         ScaleHeight     =   360
         ScaleWidth      =   615
         TabIndex        =   20
         Top             =   45
         Visible         =   0   'False
         Width           =   645
         Begin VB.Label lbl退 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "退"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   90
            TabIndex        =   21
            Top             =   0
            Width           =   405
         End
      End
      Begin VB.Frame fraInfo_1 
         Height          =   120
         Left            =   -120
         TabIndex        =   19
         Top             =   390
         Width           =   12000
      End
      Begin VB.TextBox txtNO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9315
         TabIndex        =   0
         Top             =   120
         Width           =   1800
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   6735
         TabIndex        =   28
         Top             =   165
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         Caption         =   "病人: "
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
         Height          =   240
         Left            =   45
         TabIndex        =   14
         Top             =   585
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBalance 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5070
      Width           =   11265
      _cx             =   19870
      _cy             =   661
      Appearance      =   0
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
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483633
      GridColor       =   12632256
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   360
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReplenishTheBalanceDel.frx":0E98
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
      ExplorerBar     =   3
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
Attribute VB_Name = "frmReplenishTheBalanceDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum gEM_ReplenishBalanceDelType
    EM_RBDTY_查看 = 0
    EM_RBDTY_退费 = 1
    EM_RBDTY_异常重退 = 2
End Enum
'----------------------------------------------------------------
'接口变量
Private mstrPrivs As String
Private mbytMode As gEM_ReplenishBalanceDelType
Private mstr结算序号 As String    '要查看或退费的多张单据中结帐序号
Private mblnNOMoved As Boolean '操作的单据是否在后备数据表中
Private mstrDelTime As String '查看退费单据的登记时间(yyyy-MM-dd HH:mm:ss) '只有查看退费单据时才传入时间,以区别正常单据
Private mstr结算单号 As String
'-----------------------------------------------------------
'医保相关设置
Private mstr个人帐户 As String   '医保个人帐户的名称
Private Type TY_Insure
    dbl个帐透支 As Double
    dbl帐户余额 As Double
End Type
Private mTy_Insure As TY_Insure
Private mlngModule  As Long
Private mlng领用ID As Long
Private mblnOK As Boolean
Private mblnPrintView As Boolean    '打印前查看调用
Private mblnFirst As Boolean
Private mblnNotClick As Boolean
Private mstrTittle As String
Private mstrNo As String '要查看或退费的多张单据中的某张NO,退费时可以没有

Private mrs结算方式 As ADODB.Recordset
Private mrs收费对照 As ADODB.Recordset '收费对照 :问题:33634
Private mrsBalance As ADODB.Recordset '记录每张单据的结算情况
Private mrsInfo As ADODB.Recordset

Private mobjPayCards As Cards

Private Type tyBillType
    str结算单 As String
    bln挂号   As Boolean    '是否当前为挂号结算
    strNos As String '实际读出可以退费的单据号
    strAllNOs As String '所有单据号(一次收费的所有单据)
    strDelNOs As String '当前选中要退的单据
    strNosOverFlow As String '超出金额上限的单据号
    strNosPatiDel As String '记录部分退费的单据
    strNotCanDelNOs As String  '(不能退的单据)已经退完的单据或执行不能退的单据
    
    str结算方式 As String '当前结算方式:多张时,用逗号分隔
    bln存在卡结算 As Boolean
    intInsure  As Integer   '医保单据的险类
    bln单张部分退费 As Boolean
    blnExistOnCard As Boolean '是否存在一卡通结算
    blnExistThreeAllDel As Boolean '是否存在一卡通全退的
    strInvoice As String '当前发票号
    lng原结帐ID As Long
    lng结帐ID As Long '重新结帐ID
    lng冲销ID As Long '冲销ID
    lng结算序号 As Long
    lng费用冲销ID As Long '冲销ID
    lng病人ID As Long
    str姓名 As String
    str性别 As String
    str年龄 As String
    str费别 As String
    str病人类型 As String
End Type
Private mCurBillType As tyBillType  '当前单据类型

Private mobjSquare As Object
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
'-------------------------------------------------------------------------------
'列头定义
Private Type TY_ColHead
     strRegColHead As String
     strFeeColHead As String
End Type
Private mTyColHead As TY_ColHead
'-------------------------------------------------------------------------------
'医保相关定义:参数
Private Type TYPE_MedicarePAR
    医保接口打印票据 As Boolean
    分币处理 As Boolean
    退费后打印回单 As Boolean
    医保不走票号  As Boolean        '预结算时有效
    门诊结算作废 As Boolean             '医保是否支持门诊结算作废
    门诊预结算 As Boolean
    先自付 As Boolean
    全自付 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR
'-------------------------------------------------------------------------------
'Api定义
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


'-------------------------------------------------------------------------------
'发票对象
Private mobjInvoice As zlPublicExpense.clsInvoice, mobjFact As zlPublicExpense.clsFactProperty
Private Type Ty_Module_Para
     int提醒剩余票据张数 As Integer
     bln药房单位 As Boolean
     int清单打印方式 As Integer
End Type
Private mtyMoudlePara As Ty_Module_Para
Private mobjDrugPacker  As Object ' 自动发药机(更新发药窗口)
Private mblnDrugPacker As Boolean
Private mobjDrugMachine As Object
Private mblnDrugMachine As Boolean
Private mcllForceDelToCash As Collection '强制退现信息：Array(操作员,卡类别名称,结算方式)
Private mstr排除结算方式 As String '不能使用的结算方式,多个用逗号分隔

Private Sub InitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化模块参数
    '编制:刘兴洪
    '日期:2014-09-16 16:28:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, varTemp As Variant
    With mtyMoudlePara
        .bln药房单位 = zlDatabase.GetPara("药品单位显示", glngSys, mlngModule) = "1"
        .int清单打印方式 = Val(zlDatabase.GetPara("结算清单打印方式", glngSys, mlngModule))
        strTemp = Trim(zlDatabase.GetPara("票据剩余X张时开始提醒收费员", glngSys, mlngModule, "0|10"))
        varTemp = Split(strTemp & "|", "|")
        If Val(varTemp(0)) = 0 Then
            .int提醒剩余票据张数 = -1
        Else
            .int提醒剩余票据张数 = Val(varTemp(1))
        End If
    End With
End Sub
Private Function CheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据关联性检查
    '返回:数据关联检查合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-07 11:41:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mbytMode = EM_RBDTY_查看 Then CheckDepend = True: Exit Function
    
    Set mrs结算方式 = Get结算方式("收费")
    mrs结算方式.Filter = "性质=3"
    If Not mrs结算方式.EOF Then
       mstr个人帐户 = mrs结算方式!名称
    End If
    mrs结算方式.Filter = 0
    If mrs结算方式.RecordCount = 0 Then
        MsgBox "收费场合没有可用的结算方式，请先到结算方式管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    mrs结算方式.MoveFirst
    
    Set mobjPayCards = GetPayCardsObject
    If mobjPayCards Is Nothing Then Exit Function
    If mobjPayCards.Count = 0 Then Exit Function
    
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPayCardsObject() As Cards
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取补结算支持的结算类别对象
    '返回:返回Cards对象
    '编制:刘兴洪
    '日期:2015-03-18 09:56:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, objCards As Cards, objPayCards As Cards
    Dim rsTemp As ADODB.Recordset
    Dim lngKey As Long, i As Long, blnFind As Boolean
    
    On Error GoTo errHandle
    
    Set objCards = New Cards: Set objPayCards = New Cards
    Set rsTemp = Get结算方式("补结算")
    '83533:李南春,2015/3/25,没有有效的补结算
    If rsTemp.RecordCount = 0 Then
        MsgBox "补结算没有可用的结算方式，请先到『结算方式管理』中设置补结算的应用场合。", vbInformation, gstrSysName
        Exit Function
    End If
    If Not gobjSquare Is Nothing Then
        ' zlGetCards(ByVal BytType As Byte)
        '入参:bytType-0-所有医疗卡;
        '             1-启用的医疗卡,
        '             2-所有存在三方账户的三方卡
        '             3-启用的三方账户的医疗卡
       Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
    End If
    
    With rsTemp
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        lngKey = 1
        Do While Not .EOF
            For i = 1 To objCards.Count
                If objCards(i).结算方式 = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If Not blnFind Then
                '83266:李南春,2015/3/18,医疗卡还需判断是否启用
                If InStr(",1,2,", "," & Val(Nvl(rsTemp!性质)) & ",") > 0 _
                    And Val(Nvl(rsTemp!应付款)) <> 1 Then
                    '不加入医保的结算方式或退支票的
                     Set objCard = New Card
                     objCard.短名 = Mid(Nvl(!名称), 1, 1)
                     objCard.接口编码 = Nvl(!编码)
                     objCard.接口程序名 = ""
                     objCard.接口序号 = -1 * lngKey
                     objCard.结算方式 = Nvl(!名称)
                     objCard.名称 = Nvl(!名称)
                     objCard.启用 = True
                     objCard.缺省标志 = Val(Nvl(rsTemp!缺省)) = 1
                     objCard.支付启用 = True
                     objCard.结算性质 = Val(!性质)
                    objPayCards.Add objCard, "K" & lngKey
                    lngKey = lngKey + 1
                End If
            End If
            .MoveNext
        Loop
    End With
    '加三方卡
    For Each objCard In objCards
        If objCard.消费卡 = False Then 'And objCard.是否转帐及代扣 Then
            rsTemp.Filter = "名称='" & objCard.结算方式 & "'"
            If Not rsTemp.EOF Then
                objPayCards.Add objCard, "K" & lngKey
                lngKey = lngKey + 1
            End If
        End If
    Next
    If objPayCards.Count = 0 Then
        MsgBox "结算卡设置有误,原因可能如下:" & vbCrLf & _
            "未正常启用结算卡,请到『医疗卡类别』和『设备配置』中启用", vbInformation, gstrSysName
    End If
    Set GetPayCardsObject = objPayCards
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlShowMe(frmMain As Object, ByVal lngModule As Long, _
    ByVal strPrivs As String, ByVal bytMode As gEM_ReplenishBalanceDelType, _
    Optional ByVal str结算序号 As String, _
    Optional blnPrintView As Boolean, _
    Optional lng领用ID As Long = 0, _
    Optional blnNOMoved As Boolean = False, Optional strDelTime As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:单据查看,退费
    '入参:bytMode-0-多张单据查看,1-多张单据退费,2-退异常的退费单进行重新退费
    '     strPrivs-权限串
    '     str结算序号-退费选中的结算单号
    '     blnPrintView-打印前查看调用
    '     blnNOMoved-是否转到后备数据表
    '     strDelTime-查看退费单据的登记时间(yyyy-MM-dd HH:mm:ss),只有查看退费单据时才传入时间,以区别正常单据
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-30 17:10:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnNOMoved = blnNOMoved: mstrPrivs = strPrivs
    mlng领用ID = lng领用ID: mstr结算序号 = str结算序号
    mlngModule = lngModule: mblnPrintView = blnPrintView
    mbytMode = bytMode: mstrDelTime = strDelTime              '只有查看退费单据时才传入时间,以区别正常单据
    mblnOK = False
    If CheckDepend = False Then Exit Function
    On Error Resume Next
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    On Error GoTo 0
    zlShowMe = mblnOK
End Function

Private Sub Form_Load()
    mblnFirst = True
    Call InitFace
    Call RestoreWinState(Me, App.ProductName, mstrTittle)

    If mstr结算序号 <> "" Then    '指定了结算数据的
        If mbytMode = EM_RBDTY_退费 Then
        'intFindType -0 - 按结算序号查找
        '             1-按收费单据号查找
        '             2.按结算单号查找
        '             3.按输入的发票号查找
        '             4.按挂号单号查找
            If ReadBills(0, mstr结算序号) = False Then Unload Me: Exit Sub
        Else
            If LoadViewBills(mstr结算序号) = False Then Unload Me: Exit Sub
        End If
    End If
    
    '81190,冉俊明,退费业务向发药机上传退费信息
    Call CreateDrugPacker
End Sub

Private Sub CreateDrugPacker()
    '功能:创建自助发药机(自动化药房)
    Dim objComLib As New zl9ComLib.clsComLib
    Dim strPrivs As String
    Dim strMessage As String
    
    '81190,冉俊明,退费业务向发药机上传退费信息
    mblnDrugPacker = False: mblnDrugMachine = False
    If Not (mbytMode = EM_RBDTY_退费 Or mbytMode = EM_RBDTY_异常重退) Then Exit Sub

    Err = 0: On Error Resume Next
    If Val(zlDatabase.GetPara("启用药品自动化设备接口", glngSys, Val("9010-药品自动化设备接口"))) = 1 Then
        '优先新接口
        Set mobjDrugMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err = 0 Then mblnDrugMachine = True
    End If
    
    If mblnDrugMachine = False Then
        '旧部件
        Err = 0
        Set mobjDrugPacker = CreateObject("zlDrugPacker.clsDrugPacker")
        If Err = 0 Then mblnDrugPacker = True
    End If
    
    Err = 0: On Error GoTo 0
    If mblnDrugMachine Then
        '权限检查
        strPrivs = GetPrivFunc(glngSys, Val("9010-药品自动化设备接口"))
        If InStr(";" & strPrivs & ";", ";基本;") > 0 Then
            mblnDrugMachine = mobjDrugMachine.Init(1, objComLib, strMessage)
        Else
            mblnDrugMachine = False
        End If
    ElseIf mblnDrugPacker Then
        mblnDrugPacker = mobjDrugPacker.DYEY_MZ_IniSoap
    End If
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面
    '编制:刘兴洪
    '日期:2014-06-24 14:36:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim TY_Temp As tyBillType, bytTemp As Byte
    
    mCurBillType = TY_Temp
    
    Set mobjInvoice = New zlPublicExpense.clsInvoice
    Call mobjInvoice.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    Set mobjFact = New zlPublicExpense.clsFactProperty
    
    
    Call InitModulePara '初始化模块参数
    
    Call InitBillHead(True)    '设置挂号单据列头
    mTyColHead.strRegColHead = zl_vsGrid_GetCols_Property(vsBill)
    
    Call InitBillHead(False)       '设置费用列头
    mTyColHead.strFeeColHead = zl_vsGrid_GetCols_Property(vsBill)
    
    bytTemp = Val(zlDatabase.GetPara("退费号码输入模式", glngSys, mlngModule, 0))
    IDKindNO.IDKindStr = "收|收费单号;发|发票号;挂|挂号单号;结|结算单号"
    IDKindNO.IDKind = bytTemp
    
    Call NewCardObject
    Call ClearFace
    Call SetFunCtrlVisible
    
    Select Case mbytMode
    Case EM_RBDTY_查看
        mstrTittle = "保险补充结算-查阅"
        Caption = mstrTittle
        vsBill.ColHidden(0) = True
        cmdOK.Visible = False
        cmdCancel.Caption = "退出(&X)"
        If mblnPrintView Then cmdCancel.Caption = "确定(&X)"
        pic退.Visible = mstrDelTime <> ""
        lbl退款合计.Visible = mstrDelTime <> ""
        txt退款合计.Visible = mstrDelTime <> ""
    Case EM_RBDTY_异常重退
        mstrTittle = "保险补充结算-异常退费单重新退费"
        Caption = mstrTittle
        vsBill.ColHidden(0) = True
        cmdOK.Visible = True
        pic退.Visible = mstrDelTime <> ""
        vsBill.Editable = flexEDNone
        Call initCardSquareData
    Case Else 'EM_RBDTY_退费
        mstrTittle = "保险补充结算-退费"
        Caption = mstrTittle
        Call initCardSquareData
    End Select
    If mstr结算序号 <> "" Then
        picPatiBack.Top = txtNO.Top
        lblPati.Top = picPatiBack.Top + (picPatiBack.Height - lblPati.Height) \ 2
        picPati.Height = 480
    End If
End Sub

Private Sub InitBillHead(ByVal bln挂号 As Boolean, Optional blnInit As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化退费或退号的表头列信息
    '入参:bln挂号-是否退号:true退号,False-退费
    '编制:刘兴洪
    '日期:2014-06-24 14:39:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrHead As Variant, strHead As String, i As Long
    Dim varTemp As Variant, intCol As Integer
    
    If bln挂号 Then
        strHead = "" & _
        "选择,300,4;单据号,1000,1;类别,720,1;项目,2800,1;数量,750,7;单位,550,1;单价,1100,7;" & _
        "应收金额,1100,7;实收金额,1100,7;开单科室,1000,1;执行科室,1000,1;医生,850,1;操作员,850,1;" & _
        "登记时间,1400,1;发生时间,1400,1;预约时间,1400;接收时间,1400,1;分诊时间,1400,1;诊室,1000,1;号码,720,1;号序,720,1;结帐ID;" & _
        "原始数量,0,4;准退数量,0,4;医嘱序号,0,4;执行科室ID,0,1"
    Else
        strHead = "" & _
        "选择,300,4;单据号,1000,1;类别,720,1;项目,2800,1;商品名,2000,1;数量,750,7;单位,550,1;单价,1100,7;" & _
        "应收金额,1100,7;实收金额,1100,7;开单科室,1000,1;执行科室,1000,1;操作员,850,1;时间,1260,1;结帐ID;医嘱,1560,1;" & _
        "原始数量,0,4;准退数量,0,4;医嘱序号,0,4;执行科室ID,0,1"
    End If
    
    arrHead = Split(strHead, ";")
    With vsBill
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .COLS = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            varTemp = Split(arrHead(i) & ",,,", ",")
            intCol = .FixedCols + i
            .ColKey(intCol) = Trim(varTemp(0))
            .TextMatrix(.FixedRows - 1, intCol) = varTemp(0)
            If UBound(varTemp) > 0 Then
                .ColHidden(intCol) = False
                .ColWidth(intCol) = Val(varTemp(1))
                If .ColWidth(intCol) = 0 Then .ColHidden(intCol) = True
                .ColAlignment(intCol) = Val(varTemp(2))
            Else
                .ColHidden(intCol) = True
            End If
        Next
         .TextMatrix(.FixedRows - 1, .ColIndex("选择")) = ""
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .COLS - 1) = 4
        If Not bln挂号 Then .ColHidden(.ColIndex("商品名")) = gTy_System_Para.byt药品名称显示 <> 2
        zl_vsGrid_Para_Restore mlngModule, vsBill, mstrTittle, IIf(bln挂号, "挂号列头信息", "费用列头信息")
        
        If Not blnInit Then
            zl_vsGrid_RestoreCols_Property vsBill, IIf(bln挂号, mTyColHead.strRegColHead, mTyColHead.strFeeColHead)
        End If
        .FrozenCols = 2
        .ColHidden(.ColIndex("选择")) = True
        .Editable = flexEDNone
        If mbytMode = EM_RBDTY_退费 Then
            .ColDataType(.ColIndex("选择")) = flexDTBoolean
            .ColHidden(.ColIndex("选择")) = False
            .Editable = flexEDKbdMouse
        End If
    End With
    
End Sub

Private Sub ClearFace(Optional ByVal blnNO As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除界面的信息
    '入参:blnNo=清除单据号
    '编制:刘兴洪
    '日期:2014-06-24 15:19:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tmpBillType As tyBillType
    
    mCurBillType = tmpBillType
    Set mrsBalance = Nothing
    With vsBill
        .Rows = .FixedRows '对非固定行的第一行被隐藏时恢复可见
        .Rows = .FixedRows + 1
        .Row = .FixedRows: .Col = .ColIndex("项目")
        .Clear 1
    End With
    lblPati.Caption = "病人:"
    If blnNO Then txtNO.Text = ""
    
    Call ClearBalance
    With vsBalance
         .COLS = 1
         .TextMatrix(0, 0) = IIf(mstrDelTime = "", "收款结算", "退款结算")
    End With
    txtCurTotal.Text = ""
    txtAllTotal.Text = ""
    txt退款合计.Text = ""
    stbThis.Panels(2).Text = ""
    Call SetFunCtrlVisible
End Sub

Private Sub NewCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化新的卡对象
    '编制:刘兴洪
    '日期:2014-06-24 14:43:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytMode <> EM_RBDTY_查看 Then Exit Sub
   
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    End If
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    End If
    IDKind.SetAutoReadCard (False)
End Sub
Private Sub CloseIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:关闭自助读卡功能
    '编制:刘兴洪
    '日期:2014-06-24 14:43:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        mobjICCard.SetEnabled (False)
        Set mobjICCard = Nothing
    End If
End Sub

Private Function LoadViewBills(ByVal str结算序号 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算序号来加载数据(只针对查看或异常退费)
    '入参:str结算序号-结算序号
    '返回:加载或读取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-24 16:17:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllInvoiceNoInfor As Collection
    Dim rsTemp As ADODB.Recordset, rsAdvice As ADODB.Recordset
    Dim str结帐IDs As String, strNos As String, strAllNOs As String, strFeeNos As String, strRegNos As String
    Dim strTemp As String, strWhere As String, strSQL As String, str医嘱序号 As String
    Dim lng结帐ID As Long, lng原结算序号 As Long, j As Long, lng医嘱序号 As Long, i As Long, lng冲销ID As Long
    Dim intInsure As Integer, intSign As Integer
    Dim dbl合计 As Double
    Dim varData As Variant
 
    Screen.MousePointer = 11
    intSign = IIf(mstrDelTime <> "", -1, 1) '数量,金额正负符号
    On Error GoTo errHandle
    
    str结帐IDs = zlGet结帐ID(Val(str结算序号), strNos, intInsure, mblnNOMoved, lng冲销ID, True)
    
    mCurBillType.str结算单 = strNos
    mCurBillType.lng冲销ID = lng冲销ID
    mCurBillType.intInsure = intInsure
    
    varData = Split(str结帐IDs & ",,", ",")
    If Val(varData(0)) = lng冲销ID Then
         mCurBillType.lng结帐ID = Val(varData(1))
    ElseIf Val(varData(0)) = lng冲销ID Then
         mCurBillType.lng结帐ID = Val(varData(0))
    End If
    
    
    strSQL = "" & _
    " Select A.病人ID,B.姓名,B.性别,B.年龄,B.门诊号,B.费别,b.医疗付款方式 as 付款方式,B.病人类型,B.险类,nvl(A.附加标志,0) as 附加标志" & _
    " From 费用补充记录 A, 人员表 D,病人信息 B" & _
    " Where  A.结算序号=[1] And A.操作员姓名=D.姓名 And A.病人ID=B.病人ID(+) " & _
    "       And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)" & vbNewLine & _
    "       And mod(A.记录性质,10)=1 And Rownum <2 "
    
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "费用补充记录", "H费用补充记录")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(str结算序号))
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        MsgBox "没有找到与结算相关的记录。", vbInformation, gstrSysName
        mCurBillType.lng病人ID = 0
        Exit Function
    End If
    mCurBillType.bln挂号 = Val(Nvl(rsTemp!附加标志)) = 1
    txtPatient.Text = Nvl(rsTemp!姓名)
    lblPati.Caption = "病人:" & IIf(txtPatient.Visible, "       ", rsTemp!姓名) & _
        "　性别:" & Nvl(rsTemp!性别) & _
        "　年龄:" & Nvl(rsTemp!年龄) & _
        "　门诊号:" & Nvl(rsTemp!门诊号) & _
        "　费别:" & Nvl(rsTemp!费别) & _
        "　付款方式:" & rsTemp!付款方式
    
    With mCurBillType
        .lng病人ID = Val(Nvl(rsTemp!病人ID))
        .str性别 = Nvl(rsTemp!性别)
        .str年龄 = Nvl(rsTemp!年龄)
        .str姓名 = Nvl(rsTemp!姓名)
        .str病人类型 = Nvl(rsTemp!病人类型)
        .lng原结帐ID = zlGetFromNOToLastBalanceID(mCurBillType.str结算单, False, , , True)
    End With
    
    
    If mbytMode <> EM_RBDTY_查看 Then
        Call initInsurePara(mCurBillType.intInsure, mCurBillType.lng病人ID, lng结帐ID)
    
        'bytType-0-根据NO来查找;1-根据结帐ID来查找,2-根据结算序号来查找
        If GetBalanceFeeNos(0, mCurBillType.str结算单, strFeeNos, strRegNos, mblnNOMoved) = False Then Exit Function
        If mCurBillType.bln挂号 Then
            mCurBillType.strAllNOs = strRegNos
            strNos = strRegNos
        Else
            mCurBillType.strAllNOs = strFeeNos
            strNos = strFeeNos
        End If
    End If
    
    If CheckPrivsIsValied = False Then Exit Function    '操作权限检查
    lblPati.ForeColor = vbRed
    txtPatient.ForeColor = vbRed
    Call SetPatiColor(txtPatient, Nvl(rsTemp!病人类型), txtPatient.ForeColor)
    lblPati.ForeColor = txtPatient.ForeColor
    
    '加载结算方式
    Set mrsBalance = GetChargeBalance(1, str结算序号, mblnNOMoved)
    Call LoadBalanceInfor
 
    'InStr(str结帐ID, ",") > 0:表示可能存在重收的情况，所以肯定是查的退费记录，所以摘要应该以退费的摘要为准
    strSQL = "" & _
    "   Select A.NO,Nvl(A.价格父号,A.序号) as 序号,A.从属父号,A.开单部门ID,A.执行部门ID,A.收费类别,A.费别,A.收费细目ID," & _
    "          A.费用类型,A.计算单位,max(A.医嘱序号) as 医嘱序号," & _
    "          Avg(Nvl(A.付数,1)*A.数次) as 数次," & _
    "          Sum(A.标准单价) as 单价, Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
    "          Max(A.操作员姓名) as 操作员姓名,max(A.登记时间) as 登记时间," & _
    "           " & IIf(InStr(str结帐IDs, ",") > 0, "Max(Decode(A.记录状态,2,A.摘要,NULL))", "Max(A.摘要)") & " as 摘要,A.结帐ID" & _
    "   From 门诊费用记录 A," & _
    "       (Select 结帐id" & _
    "         From (Select T1.收费结帐id As 结帐id From 费用补充记录 T1 Where T1.结算序号 = [1]" & _
    "                Union All" & _
    "                Select T1.收费结帐id As 结帐id From 费用补充记录 T1 Where T1.结算序号 = [1]" & _
    "                      And Not Exists (Select 1 From 费用补充记录 Where T1.结算序号 = 结算序号 And 记录状态 In (1, 3))" & _
    "                Union All" & _
    "                Select Distinct 结帐id From 病人预交记录 T1 Where 结算序号 = [1] And 结帐id = Abs(结算序号)" & _
    "                       And Not Exists (Select 1 From 费用补充记录 Where 结算序号 = [1] And 记录状态 In (1, 3)))" & _
    "         Group By 结帐id" & _
    "  Having Count(*) <= 1) B" & _
    "   Where Mod(A.记录性质,10)= [2] and A.结帐ID=B.结帐ID  " & _
    "   Group by A.结帐ID,A.NO,Nvl(A.价格父号,A.序号),A.从属父号,A.开单部门ID,A.执行部门ID,A.收费类别,A.费别,A.收费细目ID,A.费用类型,A.计算单位"
    
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        strSQL = Replace(strSQL, "费用补充记录", "H费用补充记录")
    End If
    If mCurBillType.bln挂号 Then
        strSQL = _
            " Select A.NO,A.序号,A.从属父号,A.费别,A.收费细目ID,C.编码 as 类别码,C.名称 as 类别名,B.编码, " & _
            "       Nvl(M1.名称,B.名称) as 名称,Max(Nvl(A.费用类型,B.费用类型)) 费用类型," & _
            "       A.计算单位  as 计算单位,Max(A.医嘱序号) as  医嘱序号," & _
            "       sum(A.数次) as 数次,Max(A.单价) as 单价,Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
            "       1 as 记录标志,0 as 原始数量,0 as 准退数量," & _
            "       D.名称 as 执行科室,A.执行部门ID,E.名称 as 开单科室,Max(A.操作员姓名) As 操作员姓名,Max(B1.执行人) as 医生, " & _
            "       Max(A.登记时间) As 登记时间,Max(B1.发生时间) as 发生时间,max(B1.预约时间) as 预约时间,max(B1.接收时间) as 接收时间,max( B1.分诊时间) as 分诊时间, " & _
            "       Max(B1.诊室) as 诊室,max(B1.号序) as 号序,max( B1.号别) as 号码,  " & _
            "       Max(A.摘要) as 摘要" & _
            " From (" & strSQL & ") A,病人挂号记录 B1,收费项目目录 B,收费项目类别 C,部门表 D,部门表 E,药品规格 X," & _
            "       收费项目别名 M1,收费项目别名 E1" & _
            " Where A.NO=B1.NO  And B1.记录状态 in (1,3) And " & _
            "       A.收费细目ID=B.ID And C.编码=A.收费类别 And A.收费细目ID=X.药品ID(+)" & _
            "       And A.执行部门ID=D.ID(+) And A.开单部门ID=E.ID(+) " & _
            "       And A.收费细目ID=M1.收费细目ID(+) And M1.码类(+)=1 And M1.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
            " Group by A.NO,A.序号,A.从属父号,A.费别,A.收费细目ID,C.编码,C.名称,B.编码,Nvl(M1.名称,B.名称)," & _
            "       E1.名称,B.规格,A.计算单位,D.名称,A.执行部门ID,E.名称,X.药品ID,X." & gstr药房单位 & _
            " Having Sum(A.数次)<>0 " & _
            " Order by NO,序号"
        If mblnNOMoved Then
            strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
        End If
    Else
        strSQL = _
        " Select A.NO,A.序号,A.从属父号,A.费别,A.收费细目ID,C.编码 as 类别码,C.名称 as 类别名,B.编码, " & _
        "       Nvl(M1.名称,B.名称) as 名称,E1.名称 as 商品名 ,B.规格,Max(Nvl(A.费用类型,B.费用类型)) 费用类型," & _
                IIf(mtyMoudlePara.bln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 计算单位," & _
        "       Max(A.医嘱序号) as 医嘱序号," & _
        "       sum(A.数次" & IIf(mtyMoudlePara.bln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ") as 数次," & _
        "       Max(A.单价" & IIf(mtyMoudlePara.bln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ") as 单价," & _
        "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, 1 as 记录标志,0 as 原始数量,0 as 准退数量," & _
        "       D.名称 as 执行科室,A.执行部门ID,E.名称 as 开单科室,Max(a.操作员姓名) As 操作员姓名, Max(a.登记时间) As 登记时间, " & _
        "       Max(A.摘要) as 摘要" & _
        " From (" & strSQL & ") A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 E,药品规格 X," & _
        "       收费项目别名 M1,收费项目别名 E1" & _
        " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.收费细目ID=X.药品ID(+)" & _
        "       And A.执行部门ID=D.ID(+) And A.开单部门ID=E.ID(+) " & _
        "       And A.收费细目ID=M1.收费细目ID(+) And M1.码类(+)=1 And M1.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
        " Group by A.NO,A.序号,A.从属父号,A.费别,A.收费细目ID,C.编码,C.名称,B.编码,Nvl(M1.名称,B.名称)," & _
        "       E1.名称,B.规格,A.计算单位,D.名称,A.执行部门ID,E.名称,X.药品ID,X." & gstr药房单位 & _
        " Having Sum(A.数次)<>0 " & _
        " Order by NO,序号"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str结算序号, IIf(mCurBillType.bln挂号, 4, 1))
    
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        If mbytMode = EM_RBDTY_查看 Then
            MsgBox "没有找到指定结算信息的费用记录,可能因并发原因被他人操作或操作了错误的结算单据。", vbInformation, gstrSysName
        Else
            MsgBox "没有找到与结算信息相关的可以退费的记录。" & _
                vbCrLf & "这些收费记录可能已经退费或已经完全执行。", vbInformation, gstrSysName
        End If
        Call ClearFace(False)
        Exit Function
    End If
    
    If mbytMode <> EM_RBDTY_退费 Then
        pic退费摘要.Enabled = mbytMode = EM_RBDTY_异常重退
        txt退费摘要.Text = Nvl(rsTemp!摘要)
    End If
    str医嘱序号 = ""
    If Not mCurBillType.bln挂号 Then
        With rsTemp
            Do While Not .EOF
                lng医嘱序号 = Val(Nvl(!医嘱序号))
                If InStr(str医嘱序号 & ",", "," & lng医嘱序号 & ",") = 0 And lng医嘱序号 <> 0 Then
                    str医嘱序号 = str医嘱序号 & "," & Val(Nvl(!医嘱序号))
                End If
                .MoveNext
            Loop
            .MoveFirst
        End With
    End If
    
    Set rsAdvice = Nothing
    If str医嘱序号 <> "" Then
        str医嘱序号 = Mid(str医嘱序号, 2)
        Set rsAdvice = zlGetAdviceFromID(str医嘱序号)
    End If
    
    Call InitBillHead(mCurBillType.bln挂号, False)
    stbThis.Panels(2).Text = "当前结算单号:" & mCurBillType.str结算单
    With vsBill
        .Redraw = flexRDNone
        .Rows = .FixedRows + rsTemp.RecordCount
        mCurBillType.strNos = ""
        For i = 1 To rsTemp.RecordCount
            .RowData(i) = Val(Nvl(rsTemp!序号))
            .TextMatrix(i, .ColIndex("选择")) = 0
            .Cell(flexcpData, i, .ColIndex("项目")) = Val(Nvl(rsTemp!从属父号))
            .Cell(flexcpData, i, .ColIndex("结帐ID")) = Nvl(rsTemp!医嘱序号) & "," & Nvl(rsTemp!收费细目ID)
            strTemp = ""
            If Val(Nvl(rsTemp!从属父号)) <> 0 Then
                rsTemp.MoveNext
                strTemp = "┣"
                If rsTemp.EOF Then
                    strTemp = "┗"
                ElseIf Val(.Cell(flexcpData, i, .ColIndex("项目"))) <> Nvl(rsTemp!从属父号) Then
                    strTemp = "┗"
                End If
                rsTemp.MovePrevious
                strTemp = "  " & strTemp & " "
            End If
            .TextMatrix(i, .ColIndex("单据号")) = Nvl(rsTemp!NO)
            .TextMatrix(i, .ColIndex("类别")) = Nvl(rsTemp!类别名)
            If mCurBillType.bln挂号 Then
                .TextMatrix(i, .ColIndex("项目")) = strTemp & rsTemp!名称
                .TextMatrix(i, .ColIndex("数量")) = FormatEx(intSign * rsTemp!数次, 5)
                .Cell(flexcpData, i, .ColIndex("数量")) = intSign * Val(Nvl(rsTemp!数次))
            Else
                .TextMatrix(i, .ColIndex("项目")) = strTemp & rsTemp!名称 & IIf(IsNull(rsTemp!规格), "", " " & rsTemp!规格)
                .TextMatrix(i, .ColIndex("商品名")) = strTemp & Nvl(rsTemp!商品名)
                .TextMatrix(i, .ColIndex("数量")) = FormatEx(intSign * Val(Nvl(rsTemp!数次)), 5)
                .Cell(flexcpData, i, .ColIndex("数量")) = intSign * Val(Nvl(rsTemp!数次))
            End If
            
            .TextMatrix(i, .ColIndex("单位")) = Nvl(rsTemp!计算单位)
            .TextMatrix(i, .ColIndex("单价")) = Format(rsTemp!单价, gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("应收金额")) = Format(intSign * Val(Nvl(rsTemp!应收金额)), gstrDec)
            .TextMatrix(i, .ColIndex("实收金额")) = Format(intSign * Val(Nvl(rsTemp!实收金额)), gstrDec)
            .TextMatrix(i, .ColIndex("开单科室")) = Nvl(rsTemp!开单科室)
            .TextMatrix(i, .ColIndex("执行科室")) = Nvl(rsTemp!执行科室)
            .TextMatrix(i, .ColIndex("操作员")) = rsTemp!操作员姓名
            If mCurBillType.bln挂号 Then
                .TextMatrix(i, .ColIndex("医生")) = Nvl(rsTemp!医生)
                .TextMatrix(i, .ColIndex("登记时间")) = Format(rsTemp!登记时间, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("发生时间")) = Format(rsTemp!发生时间, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("预约时间")) = Format(rsTemp!预约时间, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("接收时间")) = Format(rsTemp!接收时间, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("分诊时间")) = Format(rsTemp!分诊时间, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("诊室")) = Nvl(rsTemp!诊室)
                .TextMatrix(i, .ColIndex("号序")) = Nvl(rsTemp!号序)
                .TextMatrix(i, .ColIndex("号码")) = Nvl(rsTemp!号码)
            Else
                .TextMatrix(i, .ColIndex("时间")) = Format(rsTemp!登记时间, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("医嘱序号")) = Nvl(rsTemp!医嘱序号)
            End If
            .TextMatrix(i, .ColIndex("结帐ID")) = str结帐IDs
            str结算序号 = Val(Nvl(rsTemp!医嘱序号))
            If Not rsAdvice Is Nothing And str医嘱序号 <> "" And Val(str结算序号) <> 0 Then
                rsAdvice.Filter = "医嘱ID=" & Val(str结算序号)
                If rsAdvice.EOF = False Then
                    .TextMatrix(i, .ColIndex("医嘱")) = Nvl(rsAdvice!医嘱内容)
                End If
            End If
            .TextMatrix(i, .ColIndex("原始数量")) = Nvl(rsTemp!原始数量)
            .TextMatrix(i, .ColIndex("准退数量")) = Nvl(rsTemp!准退数量)
            .TextMatrix(i, .ColIndex("医嘱序号")) = Nvl(rsTemp!医嘱序号)
            .TextMatrix(i, .ColIndex("执行科室ID")) = Nvl(rsTemp!执行部门ID)
            .Cell(flexcpData, i, .ColIndex("选择")) = Val(Nvl(rsTemp!记录标志))    '用于判断是否被销帐过,>1表示已销帐
            If Val(Nvl(rsTemp!记录标志)) > 1 And InStr(1, mCurBillType.strNosPatiDel & ",", "," & rsTemp!NO & "") = 0 Then mCurBillType.strNosPatiDel = mCurBillType.strNosPatiDel & "," & rsTemp!NO
            If InStr(mCurBillType.strNos & ",", "," & rsTemp!NO & ",") = 0 Then
                '画出分隔线
                If mCurBillType.strNos <> "" Then
                    .Select i, .FixedCols, i, .COLS - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
                mCurBillType.strNos = mCurBillType.strNos & "," & rsTemp!NO
            End If
            dbl合计 = dbl合计 + Val(Nvl(rsTemp!实收金额))
            rsTemp.MoveNext
        Next
        If mCurBillType.strNos <> "" Then mCurBillType.strNos = Mid(mCurBillType.strNos, 2)
        
        .Row = .FixedRows: .Col = .ColIndex("项目")
        Call vsBill_AfterRowColChange(-1, -1, .Row, .Col)
        .Redraw = flexRDDirect
    End With
    txtAllTotal.Text = Format(intSign * dbl合计, gstrDec)
    Call ReInitPatiInvoice
    txt退款合计.Text = Format(GetDelMoney, "0.00")
    
    Screen.MousePointer = 0
    LoadViewBills = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function PrivsValied(ByVal strNo As String) As Boolean
    '补结算单据操作权限检查
    '问题:81022
    '编制:冉俊明
    '时间:2014-12-22
    Dim strOper As String, vDate As Date
    
    On Error GoTo errHandle
    If Not ReadBillInfo(1, strNo, -3, strOper, vDate) Then
        MsgBox "单据[" & strNo & "]不存在！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If InStr(mstrPrivs, "所有操作员") <= 0 And UserInfo.姓名 <> strOper Then
        MsgBox "你没有""所有操作员""权限，不能对" & strOper & "的单据进行操作！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If Not BillOperCheck(2, strOper, vDate, , strNo, , 1) Then
        Exit Function
    End If
    PrivsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ReadBills(ByVal intFindType As Integer, ByVal strFindValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前输入的结算单据号或票据号
    '入参:intFindType-0-按结算序号查找
    '             1-按收费单据号查找
    '             2.按结算单号查找
    '             3.按输入的发票号查找
    '             4.按挂号单号查找
    '     strFindValue-查找的值(0-结算序号;1-收费单据号;2-结算单据号)
    '返回:读取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-24 15:41:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, strFeeNos As String, strRegNos As String
    Dim strNos As String, strAllNOs As String
    Dim strSQLIn As String, blnNOMoved As Boolean
    Dim strTmp As String, strCanDelNos As String
    Dim i As Long, j As Integer
    Dim dbl合计 As Currency, arrNo As Variant
    Dim strTemp As String, str医嘱序号 As String
    Dim blnNotFind As Boolean
    Dim lng病人ID As Long, cllInvoiceNoInfor As Collection
    Dim str结算序号 As String
    Dim strInvoiceNO As String
    Dim str结算单号 As String, bln挂号补充 As Boolean
    Dim strTittle As String
    
     
    On Error GoTo errH
    
    If mbytMode <> EM_RBDTY_退费 Then Exit Function
    
    Screen.MousePointer = 11
    
    Call ClearFace(False)
    Set cllInvoiceNoInfor = New Collection
    Select Case intFindType
    Case 0  '按结算序号查找
        If Not GetBalanceNO(0, strFindValue, str结算单号, bln挂号补充) Then Exit Function
        strTittle = "结算号"
    Case 1  '按收费单据号查找
        If Not GetBalanceNO(1, strFindValue, str结算单号, bln挂号补充) Then Exit Function
        strTittle = "收费单号"
    Case 2  '按结算单号查找
        If Not GetBalanceNO(4, strFindValue, str结算单号, bln挂号补充) Then Exit Function
        strTittle = "结算单"
    Case 3  '按输入的发票号查找
        If Not GetBalanceNO(2, strFindValue, str结算单号, bln挂号补充) Then Exit Function
        strTittle = "发票号"
    Case 4 '按挂号单号查找
        If Not GetBalanceNO(3, strFindValue, str结算单号, bln挂号补充) Then Exit Function
        strTittle = "挂号单号"
    End Select
    If str结算单号 = "" Then
        Screen.MousePointer = 0
        MsgBox "没有找到" & strTittle & "为" & strFindValue & "相关的结算记录。", vbInformation, gstrSysName
        Exit Function
    End If
    
    mblnNOMoved = zlDatabase.NOMoved("费用补充记录", str结算单号, , 1)
    mCurBillType.str结算单 = str结算单号
    mCurBillType.bln挂号 = bln挂号补充
    
    'bytType-0-根据NO来查找;1-根据结帐ID来查找,2-根据结算序号来查找
    If GetBalanceFeeNos(0, str结算单号, strFeeNos, strRegNos, mblnNOMoved) = False Then Exit Function
    
    '单据操作权限检查
    If Not PrivsValied(str结算单号) Then Screen.MousePointer = 0:  Exit Function
    If bln挂号补充 Then
        mCurBillType.strAllNOs = strRegNos
        strNos = strRegNos
        If CheckDelRegisChargeFeeValied(mCurBillType.strAllNOs, mCurBillType.strNotCanDelNOs, strCanDelNos) = False Then
            Screen.MousePointer = 0:  Exit Function
        End If
    Else
        mCurBillType.strAllNOs = strFeeNos
        strNos = strFeeNos
        
        '升级医嘱执行计价.执行状态
        If Upgrade医嘱执行计价执行状态(strNos) = False Then
            Screen.MousePointer = 0
            MsgBox "医嘱执行计价数据修正失败，不能继续！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If CheckDelChargeIsValied(mCurBillType.strAllNOs, mCurBillType.strNotCanDelNOs, strCanDelNos) = False Then
            Screen.MousePointer = 0:  Exit Function
        End If
    End If
    
    '退费相关检查
    If strCanDelNos <> "" Then strNos = strCanDelNos

    '读取病人信息
    '----------------------------------------------------------------------------------
    strSQL = "" & _
    " Select A.病人ID,E.姓名,E.性别,E.年龄,E.门诊号 as 标识号,E.费别,E.医疗付款方式 as 付款方式,B.险类,E.病人类型" & _
    " From " & IIf(mblnNOMoved, "H", "") & "费用补充记录 A,病人信息 E,保险结算记录 B, 人员表 D" & _
    " Where A.病人ID=E.病人ID(+) And A.结算ID=B.记录ID(+) And B.性质(+)=1 And A.操作员姓名=D.姓名" & _
    "       And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)" & vbNewLine & _
    "       And mod(A.记录性质,10)=1 And A.记录状态 IN(1,3) And A.NO=[1] and rownum <2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str结算单号)
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        MsgBox "没有找到与号码""" & str结算单号 & """相关的结算记录。", vbInformation, gstrSysName
        mCurBillType.lng病人ID = 0
        Exit Function
    End If
    
    mCurBillType.lng原结帐ID = zlGetFromNOToLastBalanceID(str结算单号, blnNOMoved, , , True)
    mCurBillType.intInsure = Val(Nvl(rsTemp!险类))
    
    Call initInsurePara(mCurBillType.intInsure, lng病人ID, mCurBillType.lng原结帐ID)
    If CheckPrivsIsValied = False Then Exit Function    '操作权限检查
    
    
    txtPatient.Text = Nvl(rsTemp!姓名)

    lblPati.Caption = "病人:" & IIf(txtPatient.Visible, "                       ", rsTemp!姓名) & _
        "　性别:" & Nvl(rsTemp!性别) & _
        "　年龄:" & Nvl(rsTemp!年龄) & _
        "　门诊号:" & Nvl(rsTemp!标识号) & _
        "　费别:" & Nvl(rsTemp!费别) & _
        "　付款方式:" & rsTemp!付款方式

    With mCurBillType
        .lng病人ID = Val(Nvl(rsTemp!病人ID))
        .str性别 = Nvl(rsTemp!性别)
        .str年龄 = Nvl(rsTemp!年龄)
        .str姓名 = Nvl(rsTemp!姓名)
    End With

    If Not IsNull(rsTemp!险类) Then
        lblPati.ForeColor = vbRed
        txtPatient.ForeColor = vbRed
    Else
        lblPati.ForeColor = &HC00000
        txtPatient.ForeColor = &HC00000
    End If
    
    Call SetPatiColor(txtPatient, Nvl(rsTemp!病人类型), txtPatient.ForeColor)
    lblPati.ForeColor = txtPatient.ForeColor

    '----------------------------------------------------------------------------------
    '获取结算方式
    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    '查找类型:0-根据结帐ID查找;1-根据结算序号查找,2-根据单据号来获取结算方式
    Set mrsBalance = GetChargeBalance(2, str结算单号, mblnNOMoved)
    
    
    '初始化结算方式相关变量
    Call InitBalanceVar: Call LoadBalanceInfor
    
    str医嘱序号 = ""
    If mCurBillType.bln挂号 Then
        If GetRegListData(strNos, rsTemp) = False Then Exit Function
        If rsTemp.EOF Then
            Screen.MousePointer = 0
            MsgBox "没有找到与挂号单号为""" & Split(strNos, ",")(0) & """相关的可以退号的记录。" & _
                vbCrLf & "这些收费记录可能已经退费或已经完全执行。", vbInformation, gstrSysName
            Call ClearFace(False)
            Exit Function
        End If
    Else
        If GetFeeListData(strNos, rsTemp) = False Then Exit Function
        If rsTemp.EOF Then
            Screen.MousePointer = 0
            MsgBox "没有找到与号码""" & Split(strNos, ",")(0) & """相关的可以退费的记录。" & _
                vbCrLf & "这些收费记录可能已经退费或已经完全执行。", vbInformation, gstrSysName
            Call ClearFace(False)
            Exit Function
        End If
    End If
    mCurBillType.strNosOverFlow = ""
    strTmp = ""
    For i = 0 To UBound(Split(strNos, ","))
        strTmp = Replace(Split(strNos, ",")(i), "'", "")
        '检查是否金额超过上限
        If Not BillOperCheck(IIf(mCurBillType.bln挂号, 1, 2), rsTemp!操作员姓名, rsTemp!登记时间, IIf(mCurBillType.bln挂号, "退号", "退费"), strTmp, , 1, True) Then
            mCurBillType.strNosOverFlow = mCurBillType.strNosOverFlow & " ," & strTmp
        End If
    Next
    If mCurBillType.strNosOverFlow <> "" Then mCurBillType.strNosOverFlow = Mid(mCurBillType.strNosOverFlow, 2)
    
    Call InitBillHead(mCurBillType.bln挂号, False)
    stbThis.Panels(2).Text = "当前结算单号:" & mCurBillType.str结算单
    With vsBill
        .Redraw = flexRDNone
        .Rows = .FixedRows + rsTemp.RecordCount
        mCurBillType.strNos = ""
        For i = 1 To rsTemp.RecordCount
            .Cell(flexcpData, i, .ColIndex("项目")) = Nvl(rsTemp!从属父号)
            .Cell(flexcpData, i, .ColIndex("结帐ID")) = Nvl(rsTemp!医嘱序号) & "," & Nvl(rsTemp!收费细目ID)
            If Not mCurBillType.bln挂号 Then
                If Val(Nvl(rsTemp!医嘱序号)) <> 0 And InStr(str医嘱序号 & ",", "," & Nvl(rsTemp!医嘱序号) & ",") = 0 Then
                    str医嘱序号 = str医嘱序号 & "," & Nvl(rsTemp!医嘱序号)
                End If
            End If
            strTemp = ""
            If Val(Nvl(rsTemp!从属父号)) <> 0 Then
                rsTemp.MoveNext
                strTemp = "┣"
                If rsTemp.EOF Then
                    strTemp = "┗"
                ElseIf Val(.Cell(flexcpData, i, .ColIndex("项目"))) <> Nvl(rsTemp!从属父号) Then
                    strTemp = "┗"
                End If
                rsTemp.MovePrevious
                strTemp = "  " & strTemp & " "
            End If

            .RowData(i) = CLng(rsTemp!序号)
            .TextMatrix(i, .ColIndex("选择")) = 0
            .TextMatrix(i, .ColIndex("单据号")) = rsTemp!NO
            .TextMatrix(i, .ColIndex("类别")) = rsTemp!类别名
            .Cell(flexcpData, i, .ColIndex("类别")) = Nvl(rsTemp!类别码)
            If mCurBillType.bln挂号 Then
                .TextMatrix(i, .ColIndex("项目")) = strTemp & rsTemp!名称
                .TextMatrix(i, .ColIndex("数量")) = FormatEx(rsTemp!数次, 5)
                .Cell(flexcpData, i, .ColIndex("数量")) = Val(Nvl(rsTemp!数次))
            Else
                .TextMatrix(i, .ColIndex("项目")) = strTemp & rsTemp!名称 & IIf(IsNull(rsTemp!规格), "", " " & rsTemp!规格)
                .TextMatrix(i, .ColIndex("商品名")) = strTemp & Nvl(rsTemp!商品名)
                .TextMatrix(i, .ColIndex("数量")) = FormatEx(Nvl(rsTemp!付数, 1) * rsTemp!数次, 5)
                .Cell(flexcpData, i, .ColIndex("数量")) = Nvl(rsTemp!付数, 1) * Val(Nvl(rsTemp!数次))
            End If
            .TextMatrix(i, .ColIndex("单位")) = Nvl(rsTemp!计算单位)
            .TextMatrix(i, .ColIndex("单价")) = Format(Val(Nvl(rsTemp!单价)), gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("应收金额")) = Format(Val(Nvl(rsTemp!应收金额)), gstrDec)
            .TextMatrix(i, .ColIndex("实收金额")) = Format(Val(Nvl(rsTemp!实收金额)), gstrDec)
            .TextMatrix(i, .ColIndex("开单科室")) = Nvl(rsTemp!开单科室)
            .TextMatrix(i, .ColIndex("执行科室")) = Nvl(rsTemp!执行科室)
            .TextMatrix(i, .ColIndex("操作员")) = rsTemp!操作员姓名
            If mCurBillType.bln挂号 Then
                .TextMatrix(i, .ColIndex("医生")) = Nvl(rsTemp!医生)
                .TextMatrix(i, .ColIndex("登记时间")) = Format(rsTemp!登记时间, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("发生时间")) = Format(rsTemp!发生时间, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("预约时间")) = Format(rsTemp!预约时间, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("接收时间")) = Format(rsTemp!接收时间, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("分诊时间")) = Format(rsTemp!分诊时间, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("诊室")) = Nvl(rsTemp!诊室)
                .TextMatrix(i, .ColIndex("号序")) = Nvl(rsTemp!号序)
                .TextMatrix(i, .ColIndex("号码")) = Nvl(rsTemp!号码)
            Else
                .TextMatrix(i, .ColIndex("时间")) = Format(rsTemp!登记时间, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("医嘱")) = Nvl(rsTemp!医嘱内容)
                .TextMatrix(i, .ColIndex("医嘱序号")) = Nvl(rsTemp!医嘱序号)
            End If
            .TextMatrix(i, .ColIndex("执行科室ID")) = Nvl(rsTemp!执行部门ID)
            
            .TextMatrix(i, .ColIndex("结帐ID")) = rsTemp!结帐ID
            .TextMatrix(i, .ColIndex("原始数量")) = Val(Nvl(rsTemp!原始数量))
            .TextMatrix(i, .ColIndex("准退数量")) = Val(Nvl(rsTemp!准退数量))
            If mCurBillType.bln挂号 Then
                .Cell(flexcpChecked, i, .ColIndex("选择")) = -1  '缺省全选
            ElseIf intFindType = 1 And Nvl(rsTemp!NO) = strFindValue Then
                .Cell(flexcpChecked, i, .ColIndex("选择")) = -1 '缺省全选
            End If
            .Cell(flexcpData, i, .ColIndex("选择")) = Val(Nvl(rsTemp!记录标志))    '用于判断是否被销帐过,>1表示已销帐
            If Val(Nvl(rsTemp!记录标志)) > 1 And InStr(1, mCurBillType.strNosPatiDel & ",", "," & rsTemp!NO & "") = 0 Then mCurBillType.strNosPatiDel = mCurBillType.strNosPatiDel & "," & rsTemp!NO
            
            If InStr(mCurBillType.strNos & ",", "," & rsTemp!NO & ",") = 0 Then
                '画出分隔线
                If mCurBillType.strNos <> "" Then
                    .Select i, .FixedCols, i, .COLS - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
                mCurBillType.strNos = mCurBillType.strNos & "," & rsTemp!NO
            End If
            dbl合计 = dbl合计 + Val(Nvl(rsTemp!实收金额))
            rsTemp.MoveNext
        Next
        .Row = .FixedRows: .Col = .ColIndex("项目")
        Call vsBill_AfterRowColChange(-1, -1, .Row, .Col)
        .Redraw = flexRDDirect
    End With
    
    If str医嘱序号 <> "" Then
        Set mrs收费对照 = zlGet诊疗收费对照(Mid(str医嘱序号, 2))
    Else
        Set mrs收费对照 = Nothing
    End If
    
    If mCurBillType.strNos <> "" Then mCurBillType.strNos = Mid(mCurBillType.strNos, 2)
    
    txtAllTotal.Text = Format(dbl合计, gstrDec)
    Call LoadSelDelTotal
    Call SetFunCtrlVisible
    
    Screen.MousePointer = 0
    Call ReInitPatiInvoice
    ReadBills = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function CheckPrivsIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查操作员是否具备操作退费单
    '返回:具备返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-26 16:31:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not (mbytMode = EM_RBDTY_退费 Or mbytMode = EM_RBDTY_异常重退) Then CheckPrivsIsValied = True: Exit Function
    
    If mCurBillType.intInsure = 0 Then
        Screen.MousePointer = 0
        MsgBox "当前病人非医保病人结算单据,不允许进行退费操作！", vbInformation, gstrSysName
        Exit Function
    End If
    '保险退费权限检查
    If zlStr.IsHavePrivs(mstrPrivs, "结算退费") = False Then
        Screen.MousePointer = 0
        MsgBox "你没有权限对进行结算对费操作！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckPrivsIsValied = True: Exit Function
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub cmdBillSel_Click()
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" And _
               .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(.Row, .ColIndex("单据号")) And InStr(1, mCurBillType.strNosOverFlow, vsBill.TextMatrix(i, .ColIndex("单据号"))) <= 0 Then
                .TextMatrix(i, .ColIndex("选择")) = -1
            End If
        Next
    End With
    Call LoadSelDelTotal
End Sub

Private Sub cmdCancel_Click()
    If mCurBillType.strNos <> "" And txtNO.Visible Then
        Call ClearFace
        txtNO.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Function FromNOSelect(ByVal strNo As String, ByVal blnSel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按单据全选或全清单据
    '入参:strNO-指定的NO
    '     blnSel:true表示全选,否则全清
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-08-05 11:06:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" _
                And .TextMatrix(i, .ColIndex("单据号")) = strNo Then
                .TextMatrix(i, .ColIndex("选择")) = IIf(blnSel, -1, 0)
            End If
        Next
    End With
    FromNOSelect = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub cmdClear_Click()
    Dim i As Long, j As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("选择")) = 0
        Next
    End With
    Call LoadSelDelTotal
End Sub
 


Private Sub cmdOK_Click()
    If mbytMode = EM_RBDTY_查看 Then Unload Me: Exit Sub
    
    If mbytMode = EM_RBDTY_异常重退 Then
        '异常单据重新退费
        If ExecuteReDelFee = False Then
            '重新加载异常数据,以便读取正确的结帐数据
            Call LoadViewBills(mstr结算序号)
            Exit Sub
        End If
        mblnOK = True
        Unload Me: Exit Sub
    End If
    
    '退号
    If mCurBillType.bln挂号 Then
        Call ExecuteDelRegister: Exit Sub
    End If
    '退收费费用
    Call ExecuteDelChargeFee
End Sub

Private Sub cmdSelAll_Click()
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" And InStr(1, mCurBillType.strNosOverFlow, vsBill.TextMatrix(i, .ColIndex("单据号"))) <= 0 Then
                .TextMatrix(i, .ColIndex("选择")) = -1
            End If
        Next
    End With
    Call LoadSelDelTotal
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If txtNO.Visible And txtNO.Text = "" Then
        txtNO.SetFocus
    Else
        If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        '###
    ElseIf KeyCode = vbKeyE And Shift = vbCtrlMask Then
        If cmdOK.Visible Then Call cmdOK_Click
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If cmdSelAll.Visible Then Call cmdSelAll_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If cmdClear.Visible Then Call cmdClear_Click
    ElseIf KeyCode = vbKeyEscape Or KeyCode = vbKeyX And Shift = vbAltMask Then
        If cmdCancel.Visible Then Call cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',|~:：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
 

Private Sub Form_Resize()
    Dim staH As Long

    On Error Resume Next

    staH = IIf(stbThis.Visible, stbThis.Height, 0)

    vsBill.Height = Me.ScaleHeight - picCmd.Height - staH - picPati.Height - _
            picMoney.Height - IIf(pic退费摘要.Visible, pic退费摘要.Height, 0) - vsBalance.Height

   
    txtNO.Left = Me.ScaleWidth - txtNO.Width - 45
    IDKindNO.Left = txtNO.Left - IDKindNO.Width - 30
    pic退.Left = Me.ScaleWidth - pic退.Width - 45
    lblFormat.Left = IIf(IDKindNO.Visible, IDKindNO.Left, Me.ScaleWidth) _
            - IIf(pic退.Visible, pic退.Width + 45, 0) - lblFormat.Width - 30
    If Me.ScaleWidth - cmdSelAll.Left - cmdCancel.Width > 5500 Then
        cmdCancel.Left = Me.ScaleWidth - cmdSelAll.Left - cmdCancel.Width
    Else
        cmdCancel.Left = 5500
    End If
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 90

    fraInfo_1.Width = Me.ScaleWidth + 300
    LineCmd_1.x2 = Me.ScaleWidth + 300
    With txt退款合计
        .Left = Me.ScaleWidth - .Width - 100
        lbl退款合计.Left = .Left - lbl退款合计.Width - 20
    End With
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim tyTempBillType As tyBillType
    
    If mbytMode <> EM_RBDTY_查看 Then zlDatabase.SetPara "退费号码输入模式", IDKindNO.IDKind, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
    mCurBillType = tyTempBillType
    mbytMode = EM_RBDTY_查看
    mstrNo = "": mstrDelTime = "": mblnNOMoved = False  '查看时,可能传入true
    Call initCardSquareData: Call CloseIDCard
    
    zl_vsGrid_Para_Save mlngModule, vsBill, mstrTittle, IIf(mCurBillType.bln挂号, "挂号列头信息", "费用列头信息")
    Call SaveWinState(Me, App.ProductName, mstrTittle)
    
    If Not mobjFact Is Nothing Then Set mobjFact = Nothing
    If Not mobjInvoice Is Nothing Then Set mobjFact = Nothing
    If Not mrs结算方式 Is Nothing Then Set mrs结算方式 = Nothing
    If Not mrs收费对照 Is Nothing Then Set mrs收费对照 = Nothing
    If Not mrsBalance Is Nothing Then Set mrsBalance = Nothing
    If Not mrsInfo Is Nothing Then Set mrsInfo = Nothing
    If Not mobjDrugPacker Is Nothing Then Set mobjDrugPacker = Nothing
    If Not mobjDrugMachine Is Nothing Then Set mobjDrugMachine = Nothing
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand As String
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Visible = False Then Exit Sub   '

    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            On Error Resume Next
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            If Err <> 0 Then
                Err = 0: On Error GoTo 0
                Exit Sub
            End If
        End If
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
        Exit Sub
    End If
    lng卡类别ID = objCard.接口序号

    If lng卡类别ID = 0 Then Exit Sub
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
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)

End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub
Private Sub IDKindNO_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    zlControl.TxtSelAll txtNO
    If txtNO.Enabled And txtNO.Visible Then txtNO.SetFocus
End Sub

Private Sub pic退费摘要_Resize()
    Err = 0: On Error Resume Next
    With pic退费摘要
        txt退费摘要.Width = .ScaleWidth - txt退费摘要.Left - 50
    End With
End Sub

Private Sub txtAllTotal_GotFocus()
    zlControl.TxtSelAll txtAllTotal
End Sub

Private Sub txtCurTotal_GotFocus()
    zlControl.TxtSelAll txtCurTotal
End Sub

Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNO
End Sub
Private Function FromNOFind() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据收费单或发票号或挂号单查找允许退费的单据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-08 16:01:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim blnSucces As Boolean
    
    On Error GoTo errHandle
    If Trim(txtNO.Text) = "" Then Exit Function
    
    Set objCard = IDKindNO.GetCurCard
    If objCard Is Nothing Then Exit Function
     
    Select Case objCard.名称
    Case "收费单号"
        txtNO.Text = GetFullNO(txtNO.Text, 13)
        Call zlControl.TxtSelAll(txtNO)
        '入参:intFindType-0-按结算序号查找
         '             1-按收费单据号查找
         '             2.按结算单号查找
         '             3.按输入的发票号查找
         '             4.按挂号单号查找
        blnSucces = ReadBills(1, txtNO.Text)
    Case "发票号"
        Call zlControl.TxtSelAll(txtNO)
        blnSucces = ReadBills(3, txtNO.Text)
    Case "挂号单号"
        txtNO.Text = GetFullNO(txtNO.Text, 12)
        Call zlControl.TxtSelAll(txtNO)
        blnSucces = ReadBills(4, txtNO.Text)
    Case "结算单号"
        txtNO.Text = GetFullNO(txtNO.Text, 13)
        Call zlControl.TxtSelAll(txtNO)
        '入参:intFindType-0-按结算序号查找
         '             1-按收费单据号查找
         '             2.按结算单号查找
         '             3.按输入的发票号查找
         blnSucces = ReadBills(2, txtNO.Text)
    End Select
    Screen.MousePointer = 0
    If blnSucces Then vsBill.SetFocus
    FromNOFind = blnSucces
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 99
        Resume
    End If
End Function

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    Dim strAbc As String, str1 As String, str2 As String
    Dim objCard As Card
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtNO.Text <> "" Then
            Call FromNOFind
            Exit Sub
        End If
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Sub
    End If
    Set objCard = IDKindNO.GetCurCard
    If objCard Is Nothing Then Exit Sub
    Call SetNOInputLimit(txtNO, KeyAscii, IIf(objCard.名称 = "发票号", 1, 0))
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard (False)
End Sub
 
Private Sub txt退费摘要_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt退费摘要, KeyAscii, m文本式
End Sub
Private Sub vsBalance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Row = 1 Then
        With vsBalance
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                .Cell(flexcpForeColor, Row, Col - 1, Row, Col) = vbRed
                .Cell(flexcpFontBold, Row, .Col - 1, Row, .Col) = True
            Else
                .Cell(flexcpForeColor, Row, Col - 1, Row, Col) = Me.ForeColor
                .Cell(flexcpFontBold, Row, .Col - 1, Row, .Col) = False
            End If
        End With
    End If
    Call LoadSelDelTotal
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mbytMode = 2 Or mbytMode = 0 Then Cancel = True: Exit Sub
    With vsBalance
        If Col Mod 2 <> 0 Then Cancel = True: Exit Sub
        If Row <> 1 Then Cancel = True: Exit Sub
        If Val(.ColData(Col)) = 0 Then Cancel = True: Exit Sub
        .ColComboList(Col) = " ||" & Val(.Cell(flexcpData, Row, Col))
    End With
End Sub

Private Sub vsBalance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsBalance.MouseCol > 0 Then vsBalance.ToolTipText = vsBalance.ColData(vsBalance.MouseCol)  '显示结算摘要
End Sub

Private Sub zlSet诊疗固定关系(ByVal lngRow As Long, ByVal Col As Long, Optional lngNotCheckRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据诊疗收费关系,自动进行勾选
    '编制:刘兴洪
    '日期:2014-10-11 11:26:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, bln固定 As Boolean, i As Long, j As Long
    If vsBill.Cell(flexcpData, lngRow, vsBill.ColIndex("结帐ID")) = "" Then Exit Sub
    If mrs收费对照 Is Nothing Then Exit Sub
     
     '问题:33634:如果是固定的项目(诊疗收费关系):即医嘱产生的才判断
    varData = Split(vsBill.Cell(flexcpData, lngRow, vsBill.ColIndex("结帐ID")) & ",", ",")
    If Val(varData(0)) = 0 Then Exit Sub

    mrs收费对照.Filter = "医嘱序号=" & Val(varData(0)) & " And 收费细目ID=" & Val(varData(1))
    If Not mrs收费对照.EOF Then
        bln固定 = Val(Nvl(mrs收费对照!固有对照)) = 1
    Else
        bln固定 = False
    End If
    mrs收费对照.Filter = 0
    If bln固定 = False Then Exit Sub
    With vsBill
        For i = 1 To .Rows - 1
            If i <> lngRow And lngNotCheckRow <> i Then
                varTemp = Split(vsBill.Cell(flexcpData, i, .ColIndex("结帐ID")) & ",", ",")
                If varData(0) = varTemp(0) Then    '是相同的医嘱序号
                     mrs收费对照.Filter = "医嘱序号=" & Val(varTemp(0)) & " And 收费细目ID=" & Val(varTemp(1))
                    If Not mrs收费对照.EOF Then
                        bln固定 = Val(Nvl(mrs收费对照!固有对照)) = 1
                    Else
                        bln固定 = False
                    End If
                    If bln固定 Then
                        .Cell(flexcpChecked, i, .ColIndex("选择")) = .Cell(flexcpChecked, lngRow, .ColIndex("选择"))
                        .TextMatrix(i, .ColIndex("选择")) = .TextMatrix(lngRow, .ColIndex("选择"))
                        '如果是主项,需要检查重项
                        If Val(.Cell(flexcpData, i, .ColIndex("项目"))) = 0 Then  '肯定为父项,因此,需要找从项内容
                            For j = i + 1 To vsBill.Rows - 1
                                If .RowData(i) <> Val(.Cell(flexcpData, j, .ColIndex("项目"))) Then Exit For
                                .Cell(flexcpChecked, j, .ColIndex("选择")) = .Cell(flexcpChecked, i, .ColIndex("选择"))
                                .TextMatrix(j, .ColIndex("选择")) = .TextMatrix(i, .ColIndex("选择"))
                            Next
                        End If
                    End If
                 End If
            End If
        Next
    End With
End Sub

Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, varData As Variant, bln固定 As Boolean
    Dim varTemp As Variant, j As Long
    Dim strNo As String
    With vsBill
        If Col <> .ColIndex("选择") Then Exit Sub
        stbThis.Panels(2).Text = ""
        If mCurBillType.bln挂号 Then
            '按单据号选择
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("单据号")) <> "" And _
                   .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(.Row, .ColIndex("单据号")) And InStr(1, mCurBillType.strNosOverFlow, vsBill.TextMatrix(i, .ColIndex("单据号"))) <= 0 Then
                    .TextMatrix(i, .ColIndex("选择")) = .TextMatrix(Row, .ColIndex("选择"))
                End If
            Next
        Else
            If Val(.Cell(flexcpData, Row, .ColIndex("项目"))) = 0 Then
                For i = Row + 1 To .Rows - 1
                     If Val(.RowData(Row)) <> Val(.Cell(flexcpData, i, .ColIndex("项目"))) Then Exit For
                    .TextMatrix(i, .ColIndex("选择")) = .TextMatrix(Row, .ColIndex("选择"))
                Next
                Call zlSet诊疗固定关系(Row, Col)
            Else
                Call zlSet诊疗固定关系(Row, Col)
                '需要检查主项是否已经被
                For i = Row - 1 To 1 Step -1
                    If Val(.RowData(i)) = Val(.Cell(flexcpData, Row, .ColIndex("项目"))) Then
                        If .TextMatrix(i, .ColIndex("选择")) <> 0 Then
                             .TextMatrix(i, .ColIndex("选择")) = .TextMatrix(Row, .ColIndex("选择"))
                        End If
                        Call zlSet诊疗固定关系(i, Col, Row)
                         Exit For
                    End If
                Next
            End If
        End If
        Call LoadSelDelTotal
    End With
End Sub

Private Sub vsBill_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim dbl合计 As Currency, i As Long
    If NewRow = OldRow Then Exit Sub
    With vsBill
        If Trim(.TextMatrix(NewRow, .ColIndex("单据号"))) = "" Then
            txtCurTotal.Text = Format(dbl合计, gstrDec)
            Exit Sub
        End If
        For i = NewRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, .ColIndex("单据号")) <> .TextMatrix(NewRow, .ColIndex("单据号")) Then Exit For
            dbl合计 = dbl合计 + Val(.TextMatrix(i, .ColIndex("实收金额")))
        Next
        For i = NewRow To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> .TextMatrix(NewRow, .ColIndex("单据号")) Then Exit For
            dbl合计 = dbl合计 + Val(.TextMatrix(i, .ColIndex("实收金额")))
        Next
        txtCurTotal.Text = Format(dbl合计, gstrDec)
    End With
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If mCurBillType.bln挂号 Then
        mTyColHead.strRegColHead = zl_vsGrid_GetCols_Property(vsBill)
    Else
        mTyColHead.strFeeColHead = zl_vsGrid_GetCols_Property(vsBill)
    End If
End Sub

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBill
        If .Col <> .ColIndex("选择") Then Cancel = True: Exit Sub
        If .ColIndex("单据号") < 0 Then Cancel = True: Exit Sub
        If Trim(.TextMatrix(Row, .ColIndex("单据号"))) = "" Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vsBill_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsBill.ColIndex("选择") Then Cancel = True
End Sub

Private Sub GetBillRow(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据的指定行，获取单据的开始行和结速行
    '入参:lngRow-当前行
    '出参:lngBegin-单据的开始行
    '     lngEnd-单据的结束行
    '编制:刘兴洪
    '日期:2014-07-03 17:39:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    lngBegin = lngRow: lngEnd = lngRow
    With vsBill
        For i = lngRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, .ColIndex("单据号")) <> .TextMatrix(lngRow, .ColIndex("单据号")) Then Exit For
            lngBegin = i
        Next
        For i = lngRow To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> .TextMatrix(lngRow, .ColIndex("单据号")) Then Exit For
            lngEnd = i
        Next
    End With
End Sub

Private Sub vsBill_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsBill
        If .ColIndex("单据号") < 0 Then Exit Sub
        '超出限额的设置
        If .TextMatrix(Row, .ColIndex("单据号")) <> "" _
            And InStr(1, mCurBillType.strNosOverFlow, .TextMatrix(Row, .ColIndex("单据号"))) > 0 Then
             .TextMatrix(Row, .ColIndex("选择")) = 0
        End If
    End With
End Sub

Private Sub vsBill_DrawCell(ByVal hDC As Long, ByVal Row As Long, _
    ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, _
    ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:画线
    '编制:刘兴洪
    '日期:2014-07-03 17:41:52
    '说明：1.OwnerDraw要设置为Over(画出单元所有内容)
    '      2.Cell的GridLine从上下左右向内都是从第1根线开始
    '      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT

    With vsBill
        '擦除一并给药相关行列的边线及内容
        lngLeft = .ColIndex("单据号"): lngRight = .ColIndex("单据号")
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        Call GetBillRow(Row, lngBegin, lngEnd)
        If lngBegin = lngEnd Then Exit Sub

        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If

        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsBill_KeyPress(KeyAscii As Integer)
    With vsBill
        Select Case KeyAscii
        Case 32 '空格
            If .ColHidden(.ColIndex("选择")) Then Exit Sub
            KeyAscii = 0
            If Trim(.TextMatrix(.Row, .ColIndex("单据号"))) = "" Then Exit Sub
            
            If .TextMatrix(.Row, .ColIndex("选择")) = 0 _
                And InStr(1, mCurBillType.strNosOverFlow, .TextMatrix(.Row, .ColIndex("单据号"))) <= 0 Then
                 .TextMatrix(.Row, .ColIndex("选择")) = -1
            Else
                 .TextMatrix(.Row, .ColIndex("选择")) = 0
            End If
            Call LoadSelDelTotal
            
            '87675,需要手动触发AfterEdit事件
            Call vsBill_AfterEdit(.Row, .ColIndex("选择"))
        Case 13 '回车
            KeyAscii = 0
            If .Row + 1 <= .Rows - 1 Then
               .Row = .Row + 1: .ShowCell .Row, .Col
            End If
        End Select
    End With
End Sub
Private Sub vsBill_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBill
        Select Case Col
        Case .ColIndex("选择")
        Case Else
             Cancel = True
        End Select
    End With
End Sub

Private Function CheckDelChargeIsValied(ByVal strNos As String, _
    ByRef strNotCanDelNOs As String, _
    ByRef strCanDelNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费单据是否合法
    '入参:strNOs-需要检查的单据号(多个用逗号分离)
    '出参:strNotCanDelNOs-不能退的单据(已经执行及不能退的单据)
    '     strCanDelNos-能退的单号
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-11 11:28:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrNo As Variant, i As Long, intTmp As Integer
    Dim strInfo As String, strFlagPrintInfor As String
    Dim blnFlagPrint As Boolean, strNo As String, strCurNO As String
    Dim blnHaveExe As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle

    '问题:54728
    If Not mbytMode = EM_RBDTY_退费 Then CheckDelChargeIsValied = True: Exit Function   '退费时判断

    arrNo = Split(strNos, ","): strNotCanDelNOs = ""
    strCanDelNos = ""   '记录可以退的单据号
    strInfo = ""        '检查结果提示信息
    strFlagPrintInfor = ""
    For i = 0 To UBound(arrNo)
        strCurNO = Replace(arrNo(i), "'", "")
        If strNo = "" Then strNo = strCurNO

        blnHaveExe = False: blnFlagPrint = False
        intTmp = BillCanDelete(strCurNO, 1, blnHaveExe, , blnFlagPrint)
        If intTmp <> 0 Then
            strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
            Select Case intTmp
                Case 1 '该单据不存在
                    strInfo = strInfo & "指定的单据不存在！" & vbCrLf
                    Exit For
                Case 2 '已经全部完全执行(收费不考虑退费自动退药)
                    strInfo = strInfo & "[" & strCurNO & "]中的项目已经全部完全执行,不能退费!" & vbCrLf
                Case 3 '未完全执行部分剩余数量为0
                    strInfo = strInfo & "[" & strCurNO & "]中未完全执行的项目剩余数量为零,没有可退费用！" & vbCrLf
            End Select

        ElseIf blnHaveExe Then
            If gbln退费申请模式 Then
                '未申请或未审核的单据不能退费
                Set rsTemp = GetApply(strCurNO, 1)
                rsTemp.Filter = "状态<>2"
                If rsTemp.RecordCount = 0 Then
                    strInfo = strInfo & "[" & strCurNO & "]未进行退费申请及审核，不能进行退费！" & vbCrLf
                    strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
                ElseIf IsNull(rsTemp!审核人) Then
                    strInfo = strInfo & "[" & strCurNO & "]未进行退费审核，不能进行退费！" & vbCrLf
                    strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
                Else
                    strInfo = strInfo & "[" & strCurNO & "]中存在已执行的项目，此单据将执行的是部分退费。" & vbCrLf
                    strCanDelNos = strCanDelNos & "," & strCurNO
                End If
            Else
                strInfo = strInfo & "[" & strCurNO & "]中存在已执行的项目，此单据将执行的是部分退费。" & vbCrLf
                strCanDelNos = strCanDelNos & "," & strCurNO
            End If
        ElseIf gbln退费申请模式 Then
            '未申请或未审核的单据不能退费
            Set rsTemp = GetApply(strCurNO, 1)
            rsTemp.Filter = "状态<>2"
            If rsTemp.RecordCount = 0 Then
                strInfo = strInfo & "[" & strCurNO & "]未进行退费申请及审核，不能进行退费！" & vbCrLf
                strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
            ElseIf IsNull(rsTemp!审核人) Then
                strInfo = strInfo & "[" & strCurNO & "]未进行退费审核，不能进行退费！" & vbCrLf
                strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
            Else
                strCanDelNos = strCanDelNos & "," & strCurNO
            End If
        Else
            strCanDelNos = strCanDelNos & "," & strCurNO
        End If

        If blnFlagPrint Then
            '检查对应的条码是否已打印(检验医嘱中的采集方式已执行)
            strFlagPrintInfor = strFlagPrintInfor & "[" & strCurNO & "]检验医嘱的条码已打印。" & vbCrLf
        End If
    Next

    If strNotCanDelNOs <> "" Then strNotCanDelNOs = Mid(strNotCanDelNOs, 2)
    strCanDelNos = Mid(strCanDelNos, 2)

    If strFlagPrintInfor <> "" Then
        If MsgBox("注意:" & vbCrLf & strFlagPrintInfor & vbCrLf & " 是否继续退费？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Screen.MousePointer = 0
            Exit Function
        End If
    End If

    If strCanDelNos = "" Then
        '多张单据因为登记日期一样,必然是一起转出或都没有转出
        '是否已转入后备数据表中
        If zlDatabase.NOMoved("门诊费用记录", strNo, , "1") Then
            If Not ReturnMovedExes(strNo, 1, Me.Caption) Then
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
        Screen.MousePointer = 0
        MsgBox strInfo, vbInformation, gstrSysName
        Exit Function
    End If

    If strInfo <> "" Then
        MsgBox strInfo, vbInformation, gstrSysName
    End If
    strNos = strCanDelNos

    '多张单据因为登记日期一样,必然是一起转出或都没有转出
    '是否已转入后备数据表中
    If zlDatabase.NOMoved("门诊费用记录", strNo, , "1") Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then
            Screen.MousePointer = 0
            Exit Function
        End If
    End If
    CheckDelChargeIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function CheckDelRegisChargeFeeValied(ByVal strNos As String, _
    ByRef strNotCanDelNOs As String, _
    ByRef strCanDelNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查挂号单据退号是否合法
    '入参:strNOs-需要检查的单据号(多个用逗号分离)
    '出参:strNotCanDelNOs-不能退的单据(已经执行及不能退的单据)
    '     strCanDelNos-能退的单号
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-11 11:28:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrNo As Variant, i As Long, intTmp As Integer
    Dim strInfo As String, strFlagPrintInfor As String
    Dim blnFlagPrint As Boolean, strNo As String, strCurNO As String
    Dim blnHaveExe As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle

    '问题:54728
    If Not mbytMode = EM_RBDTY_退费 Then CheckDelRegisChargeFeeValied = True: Exit Function   '退费时判断

    arrNo = Split(strNos, ","): strNotCanDelNOs = ""
    strCanDelNos = ""   '记录可以退的单据号
    strInfo = ""        '检查结果提示信息
    strFlagPrintInfor = ""
    For i = 0 To UBound(arrNo)
        strCurNO = Replace(arrNo(i), "'", "")
        If strNo = "" Then strNo = strCurNO

        blnHaveExe = False: blnFlagPrint = False
        intTmp = BillCanDelete(strCurNO, 4, blnHaveExe, , blnFlagPrint)
        If intTmp <> 0 Then
            strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
            Select Case intTmp
                Case 1 '该单据不存在
                    strInfo = strInfo & "指定的单据不存在！" & vbCrLf
                    Exit For
                Case 2 '已经全部完全执行(收费不考虑退费自动退药)
                    strInfo = strInfo & "[" & strCurNO & "]中的项目已经全部完全执行,不能退号!" & vbCrLf
                Case 3 '未完全执行部分剩余数量为0
                    strInfo = strInfo & "[" & strCurNO & "]中未完全执行的项目剩余数量为零,不能退号！" & vbCrLf
            End Select

        ElseIf blnHaveExe Then
            '存在已执行项目
            strInfo = strInfo & "[" & strCurNO & "]中存在已执行的项目，此单据将执行的是部分退费。" & vbCrLf
            strCanDelNos = strCanDelNos & "," & strCurNO
        Else
            strCanDelNos = strCanDelNos & "," & strCurNO
        End If
        
        If blnFlagPrint Then
            '检查对应的条码是否已打印(检验医嘱中的采集方式已执行)
            strFlagPrintInfor = strFlagPrintInfor & "[" & strCurNO & "]检验医嘱的条码已打印。" & vbCrLf
        End If
    Next

    If strNotCanDelNOs <> "" Then strNotCanDelNOs = Mid(strNotCanDelNOs, 2)
    strCanDelNos = Mid(strCanDelNos, 2)

    If strFlagPrintInfor <> "" Then
        If MsgBox("注意:" & vbCrLf & strFlagPrintInfor & vbCrLf & " 是否继续退号？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Screen.MousePointer = 0
            Exit Function
        End If
    End If

    If strCanDelNos = "" Then
        '多张单据因为登记日期一样,必然是一起转出或都没有转出
        '是否已转入后备数据表中
        If zlDatabase.NOMoved("门诊费用记录", strNo, , "4") Then
            If Not ReturnMovedExes(strNo, 4, Me.Caption) Then
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
        Screen.MousePointer = 0
        MsgBox strInfo, vbInformation, gstrSysName
        Exit Function
    End If

    If strInfo <> "" Then
        MsgBox strInfo, vbInformation, gstrSysName
    End If
    strNos = strCanDelNos

    '多张单据因为登记日期一样,必然是一起转出或都没有转出
    '是否已转入后备数据表中
    If zlDatabase.NOMoved("门诊费用记录", strNo, , "4") Then
        If Not ReturnMovedExes(strNo, 4, Me.Caption) Then
            Screen.MousePointer = 0
            Exit Function
        End If
    End If
    CheckDelRegisChargeFeeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub InitBalanceVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化单据
    '编制:刘兴洪
    '日期:2014-07-04 10:02:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String
    
    If mrsBalance Is Nothing Then Exit Sub
    If mrsBalance.State <> 1 Then Exit Sub
    
    mrsBalance.Filter = "类型<>2 And 类型<>1"
    '       字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '            是否密文,是否全退,是否退现,冲预交
    '       类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    str结算方式 = ""
    mrsBalance.Sort = "类型,结算性质"
    With mrsBalance
        Do While Not .EOF
            If InStr(str结算方式 & ",", "," & Nvl(!结算方式) & ",") = 0 Then
                str结算方式 = str结算方式 & "," & Nvl(!结算方式)
            End If
            If Val(Nvl(!类型)) = 3 Or Val(Nvl(!类型)) = 4 Then mCurBillType.bln存在卡结算 = True
            .MoveNext
        Loop
    End With
    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 2)
    mCurBillType.str结算方式 = str结算方式
    
    '4-一卡通(老)
    mrsBalance.Filter = "类型=4"
    mCurBillType.blnExistOnCard = mrsBalance.EOF = False
    
    '3.一卡通
    mrsBalance.Filter = "类型=3 And  是否全退=1 and 是否退现=0"
    mCurBillType.blnExistThreeAllDel = mrsBalance.EOF = False
    mrsBalance.Filter = 0
End Sub

Private Function ExecuteClinicDelSwap(ByVal lng病人ID As Long, ByVal intInsure As Integer, _
    ByVal lng冲销ID As Long, ByVal lng原结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行医保退费交易
    '入参:lng病人ID-病人ID
    '     intInsure-险类
    '     lng冲销ID-冲销ID
    '     lng原结帐ID-原始结帐ID
    '出参:
    '返回:医保退费交易成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-31 23:38:11
    '说明:
    '   调用接口前,必须先打开事务,完成后,会自动提交事务;失败时,会回退事务
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, strAllBalance As String, strSQL As String
    Dim varData As Variant, varTemp As Variant, i As Long
    
    On Error GoTo errHandle
    
    If intInsure = 0 Then ExecuteClinicDelSwap = True: gcnOracle.CommitTrans: Exit Function
    strAllBalance = GetYBOldBalance(lng病人ID, intInsure, lng原结帐ID)
    
    strAdvance = ""
    If MCPAR.门诊结算作废 Then
        strAdvance = lng冲销ID
        'ClinicDelSwap (医保退费结算)
        '参数名  参数类型    入/出   原参数说明  现调整说明
        'lngStlID    long    IN  将要退费的费用记录的结帐ID(原结帐ID)
        'bln退费 Boolean IN  表明是退费交易还是改费交易在调用本接口
        'intInsure   Intger  In  险类
        'strAdvance  String  In  NULL    冲销ID:增加传入冲销ID
        '医保可以根据冲销ID来进行取数
        '        Out 退费结算：结算方式1|金额||结算方式2|金额...
        '    Boolean 函数返回    True:调用成功,False:调用失败
        '冲销ID|补充结算标志|
        strAdvance = lng冲销ID & "|1"
        If Not gclsInsure.ClinicDelSwap(lng原结帐ID, , intInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
        If strAdvance = CStr(lng冲销ID) & "|1" Then strAdvance = ""
    Else
        strAdvance = strAllBalance
        varData = Split(strAdvance, "||")
        strAdvance = ""
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & "|||", "|")
            strAdvance = strAdvance & "||" & varTemp(0) & "|" & -1 * Val(varTemp(1))
        Next
        If strAdvance <> "" Then strAdvance = Mid(strAdvance, 3)
    End If
    
    If MCPAR.门诊结算作废 Then
        If Not zlInsureCheck(strAllBalance, strAdvance) Or strAdvance = "" Then
            gcnOracle.CommitTrans
            If MCPAR.门诊结算作废 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
            ExecuteClinicDelSwap = True: Exit Function
        End If
        gcnOracle.CommitTrans: gcnOracle.BeginTrans
    End If
    
    '退费和收费不一致时,需要效对
    'Zl_费用补充结算_Modify
    strSQL = "Zl_费用补充结算_Modify("
    '  操作类型_In   Number,
    '  --   0-普通结算方式:
    '  --     结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --   1.三方卡结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    strSQL = strSQL & "" & 2 & ","
    '  结算id_In     In 费用补充记录.结算id%Type,
    strSQL = strSQL & "" & lng冲销ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "'" & strAdvance & "')"
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    '  卡号_In       病人预交记录.卡号%Type := Null,
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    '  完成结算_In Number:=0
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
    ExecuteClinicDelSwap = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mCurBillType.intInsure)
End Function
Private Function isChargeFeeValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退费是否合法
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-11 11:34:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrNo As Variant, strYPNos As String, bln药品 As Boolean, blnSel As Boolean
    Dim i As Long, strDelNOs As String, strNo As String, strOperatorName As String
    Dim varTemp As Variant, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If Not CheckTextLength("退费摘要", txt退费摘要) Then Exit Function
    '检查输入是否正确
    If mCurBillType.strNos = "" Then
        MsgBox "请先确认需要退费的门诊收费单据。", vbInformation, gstrSysName
        If txtNO.Visible Then txtNO.SetFocus: Exit Function
    End If
    
    '检查本次结算单据中是否存在退费异常单据，若存在，则不允许继续退费
    If CheckIsExistDelErrBill(mCurBillType.str结算单, strOperatorName) Then
        MsgBox "注意：" & vbCrLf & _
            "    本次结算中存在异常的结算记录，请先对其进行重新退费！" & _
            IIf(strOperatorName <> UserInfo.姓名, vbCrLf & "    提示：异常单据是操作员【" & strOperatorName & "】收取的。", ""), vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    arrNo = Split(mCurBillType.strNos, ",")
    strYPNos = "": strDelNOs = ""
    bln药品 = False: blnSel = False
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 Then
                blnSel = True
                strNo = .TextMatrix(i, .ColIndex("单据号"))
                If InStr(strDelNOs & ",", "," & strNo & ",") = 0 Then
                    strDelNOs = strDelNOs & "," & strNo
                End If
                
                If .ColIndex("类别") <> -1 And bln药品 = False Then     '47400
                    If .TextMatrix(i, .ColIndex("类别")) Like "*西*药*" _
                        Or .TextMatrix(i, .ColIndex("类别")) Like "*中*药*" _
                        Or .TextMatrix(i, .ColIndex("类别")) Like "*卫材*" Then
                        If InStr(strYPNos & ",", "," & strNo & ",") = 0 Then
                            strYPNos = strYPNos & "," & strNo
                        End If
                        bln药品 = True
                    End If
                End If
            End If
        Next
    End With
    If strDelNOs <> "" Then strDelNOs = Mid(strDelNOs, 2)
    
    If strDelNOs <> "" And gbln退费申请模式 And Not mCurBillType.bln挂号 Then
        Set rsTemp = GetApply(strDelNOs, 1)
        varTemp = Split(strDelNOs, ",")
        For i = 0 To UBound(varTemp)
            strNo = varTemp(i)
            rsTemp.Filter = "NO='" & strNo & "' And 状态<>2"
            If rsTemp.RecordCount = 0 Then
                Screen.MousePointer = 0
                MsgBox "请先对收费单据:" & strNo & " 进行退费申请！", vbInformation, gstrSysName
                Exit Function
            End If
            If IsNull(rsTemp!审核人) Then
                Screen.MousePointer = 0
                MsgBox "单据:" & strNo & " 未进行退费审核，不能进行退费！", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End If
    If blnSel = False Then
        MsgBox "请在单据中至少选择一个要退费的项目。", vbInformation, gstrSysName
        vsBill.SetFocus: Exit Function
    End If
    
    If bln药品 And Not mCurBillType.bln挂号 Then
        If strYPNos <> "" Then strYPNos = Mid(strYPNos, 2)
        If zlCheckDrugIsPutDrug(strYPNos) = False Then Exit Function
    End If
    
    '医保检查
    If mCurBillType.intInsure = 0 Then
        MsgBox "当前结算不是医保病人结算,不允许进行" & IIf(Not mCurBillType.bln挂号, "退费", "退号") & "操作!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If gclsInsure.CheckInsureValid(mCurBillType.intInsure) = False Then Exit Function
    If Not mCurBillType.bln挂号 Then
        If zlCheckIsMzToZY(strDelNOs, 1) Then
              MsgBox "注意:" & vbCrLf & _
                "    该单据已经被门诊费用转住院费用 " & vbCrLf & _
                "    或已经审核了门诊费用转住院费用,不能再退费", vbInformation + vbOKOnly, gstrSysName
              Exit Function
        End If
        
        If MCPAR.门诊结算作废 = False Then '112843
            MsgBox "当前医保不支持门诊结算作废，不能进行退费！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    '三方卡结算方式有效性检查
    If ThreeBalanceCheck(mobjPayCards, mrsBalance, mcllForceDelToCash, mstr排除结算方式) = False Then Exit Function
    
    isChargeFeeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ThreeBalanceCheck(objCards As Cards, ByVal rsBalance As ADODB.Recordset, _
    ByRef cllForceDelToCash As Collection, ByRef str排除结算方式 As String) As Boolean
    '三方卡结算方式有效性检查
    '入参：
    '   objCards 补结算所有有效的支付方式
    '   rsBalance 结算信息
    '出参：
    '   cllForceDelToCash 强制退现信息：Array(操作员,卡类别名称,结算方式)
    '   str排除结算方式 排除结算方式,多个用逗号分隔
    '返回：检查通过，返回True；否则，返回False
    '105432
    Dim objCard As Card
    Dim cllFeeBalance As New Collection, i As Integer
    Dim blnFind As Boolean, blnQuestion As Boolean
    Dim str操作员 As String, strKey As String
    Dim dblMoney  As Double
    Dim j As Integer, lngCount As Long
    Dim varData As Variant
    
    On Error GoTo errHandler
    Set cllForceDelToCash = New Collection
    str排除结算方式 = ""
    If rsBalance Is Nothing Then ThreeBalanceCheck = True: Exit Function
    
    '类型：0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    rsBalance.Filter = "类型= 3"
    '去重
    With rsBalance
        Do While Not .EOF
            strKey = "_" & Val(Nvl(!卡类别ID))
            If CollectionExitsValue(cllFeeBalance, strKey) Then
                dblMoney = cllFeeBalance(strKey)(4) + Val(Nvl(!冲预交))
                cllFeeBalance.Remove strKey
            Else
                dblMoney = Val(Nvl(!冲预交))
            End If
            If RoundEx(dblMoney, 6) > 0 Then '全部退完的就不再加入
                'Array(结算方式,卡类别ID,是否退现,卡类别名称,冲预交,是否全退,是否转帐及代扣)
                cllFeeBalance.Add Array(Nvl(!结算方式), Val(Nvl(!卡类别ID)), Val(Nvl(!是否退现)), _
                    Nvl(!卡类别名称), dblMoney, Val(Nvl(!是否全退)), Nvl(!是否转帐及代扣)), strKey
            End If
            .MoveNext
        Loop
    End With
    If cllFeeBalance.Count = 0 Then ThreeBalanceCheck = True: Exit Function
    
    For i = 1 To cllFeeBalance.Count
        blnQuestion = False
        '医疗卡检查
        If objCards Is Nothing Then
            If MsgBox("『" & cllFeeBalance(i)(3) & "』未启用，该医疗卡支付的金额将被退为其它结算方式，是否继续？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            blnQuestion = True
        Else
            blnFind = False
            For Each objCard In objCards
                If objCard.接口序号 = cllFeeBalance(i)(1) Then blnFind = True: Exit For
            Next
            If blnFind = False Then
                If MsgBox("『" & cllFeeBalance(i)(3) & "』未启用，该医疗卡支付的金额将被退为其它结算方式，是否继续？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnQuestion = True
            End If
        End If
        
        If blnQuestion Then
            If cllFeeBalance(i)(2) = 0 Then '强制退现
                If str操作员 = "" Then '多种卡类别时只验证一次
                    If zlStr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "三方退款强制退现") Then
                        str操作员 = UserInfo.姓名
                    Else
                        str操作员 = zlDatabase.UserIdentifyByUser(Me, "医疗卡『" & cllFeeBalance(i)(3) & "』强制退现，权限验证：", _
                            glngSys, mlngModule, "三方退款强制退现", , True)
                        If str操作员 = "" Then Exit Function
                    End If
                End If
                'Array(操作员,卡类别名称,结算方式)
                cllForceDelToCash.Add Array(str操作员, cllFeeBalance(i)(3), cllFeeBalance(i)(0))
            End If
        ElseIf cllFeeBalance(i)(5) = 1 Then '必须全退
            If cllFeeBalance(i)(2) = 1 Then '允许退现，必须全退
                If cllFeeBalance(i)(6) = 0 Then '不支持转帐及代扣
                    If MsgBox("『" & cllFeeBalance(i)(3) & "』必须全退，因此不能退回原卡。" & _
                        "如果继续操作，那么该医疗卡支付的金额将被退为其它结算方式，是否继续？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    str排除结算方式 = str排除结算方式 & "," & cllFeeBalance(i)(0)
                End If
            ElseIf cllFeeBalance(i)(6) = 0 Then '不允许退现，必须全退，且不支持转帐及代扣
                If MsgBox("『" & cllFeeBalance(i)(3) & "』必须全退且不能退现，同时也不支持转帐及代扣，因此无法退回原卡。" & _
                    "如果继续操作，那么该医疗卡支付的金额将被退为其它结算方式，是否继续？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                If str操作员 = "" Then '多种卡类别时只验证一次
                    If zlStr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "三方退款强制退现") Then
                        str操作员 = UserInfo.姓名
                    Else
                        str操作员 = zlDatabase.UserIdentifyByUser(Me, "『" & cllFeeBalance(i)(3) & "』强制退现，权限验证：", _
                            glngSys, mlngModule, "三方退款强制退现", , True)
                        If str操作员 = "" Then Exit Function
                    End If
                End If
                'Array(操作员,卡类别名称,结算方式)
                cllForceDelToCash.Add Array(str操作员, cllFeeBalance(i)(3), cllFeeBalance(i)(0))
                str排除结算方式 = str排除结算方式 & "," & cllFeeBalance(i)(0)
            End If
        End If
    Next
    If str排除结算方式 <> "" Then str排除结算方式 = Mid(str排除结算方式, 2)
    

    If str排除结算方式 = "" Then ThreeBalanceCheck = True: Exit Function
    '判断是否还有有效的结算方式
    varData = Split(str排除结算方式, ",")
    lngCount = mobjPayCards.Count
    For i = 1 To mobjPayCards.Count
        If mobjPayCards(i).接口序号 <= 0 Or mobjPayCards(i).接口序号 > 0 And mobjPayCards(i).消费卡 Then
            Exit For
        End If
        
        blnFind = False
        For j = 0 To UBound(varData)
            If mobjPayCards(i).结算方式 = varData(j) Then
                lngCount = lngCount - 1: blnFind = True
            End If
        Next
        If blnFind = False Then Exit For
    Next
    If lngCount <= 0 Then
        MsgBox "排除强制退现的结算方式后，已没有可用的结算方式，不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    ThreeBalanceCheck = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCashMoney(ByVal strNo As String) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：医保不支持退个人帐户时,个人帐户退现金,获取现金退款金额
    '参数：
    '   strNO-挂号单号
    '返回：结算金额
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "" & _
            "   Select -1 * a.冲预交 As 现金" & _
            "   From 病人预交记录 A, 门诊费用记录 B, 结算方式 C" & _
            "   Where a.结帐id = b.结帐id And a.结算方式 Is Null" & _
            "         And b.No = [1] And a.记录性质 = 4 And a.记录状态 = 2 And Rownum = 1"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取医保结算金额", strNo)
    
    If Not rsTmp.BOF Then GetCashMoney = CCur(Nvl(rsTmp!现金))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExecuteDelRegister() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行退号操作
    '返回:退号成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-09 16:57:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, cllPro As Collection, strNo As String
    Dim str冲销ID As String, str结算序号 As String, strDelDate As String
    Dim str结算冲销ID As String, strAdvance As String, str个人帐户 As String
    
    On Error GoTo errHandle
    If isRegisterValied(strNo) = False Then Exit Function

    '检查是否允许医保作废
    str个人帐户 = IIf(mstr个人帐户 <> "", mstr个人帐户, "个人帐户")
    If mCurBillType.intInsure <> 0 Then
        If gclsInsure.GetCapability(support门诊结算作废, , mCurBillType.intInsure, str个人帐户) Then
            str个人帐户 = ""     '向过程传入不允许退的结算方式,空表示全部允许
        End If
    End If
    
    Set cllPro = New Collection
    str冲销ID = zlDatabase.GetNextId("病人结帐记录")
    str结算序号 = "-" & str冲销ID
    strDelDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    '1.保存数据
    'Zl_病人挂号补结算_Delete
    strSQL = "Zl_病人挂号补结算_Delete("
    '  单据号_In     门诊费用记录.No%Type,
    strSQL = strSQL & "'" & strNo & "',"
    '  操作员编号_In 门诊费用记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  冲销id_In     门诊费用记录.结帐id%Type := Null,
    strSQL = strSQL & "" & str冲销ID & ","
    '  结算序号_In   病人预交记录.结算序号%Type := Null,
    strSQL = strSQL & "" & str结算序号 & ","
    '  退号时间_In   门诊费用记录.登记时间%Type := Null,
    strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'))"
    '  删除门诊号_In Number:=0
    zlAddArray cllPro, strSQL
    
    str结算冲销ID = zlDatabase.GetNextId("病人结帐记录")
    'Zl_费用补充记录_Delete
    strSQL = "Zl_费用补充记录_Delete("
    '  No_In         In 费用补充记录.No%Type,
    strSQL = strSQL & "'" & mCurBillType.str结算单 & "',"
    '  冲销id_In     In 费用补充记录.结算id%Type,
    strSQL = strSQL & "" & str结算冲销ID & ","
    '  重结id_In     In 费用补充记录.结算id%Type,
    strSQL = strSQL & "NULL,"
    '  结算序号_In   In 费用补充记录.结算序号%Type,
    strSQL = strSQL & "" & str结算序号 & ","
    '  退费结帐id_In varchar2(多个用逗事情分离),
    strSQL = strSQL & "'" & str冲销ID & "',"
    '  操作员编号_In In 费用补充记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In In 费用补充记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  登记时间_In   In 费用补充记录.登记时间%Type := Null
    strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
    '  非原样退结算_In In Varchar2 := Null
    strSQL = strSQL & "'" & str个人帐户 & "')"
    zlAddArray cllPro, strSQL
    Err = 0: On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    '调用医保接口
    '挂号费收取方式|挂号单号|补结算标志,用|分隔
    strAdvance = "0|" & strNo & "|1"
    If Not gclsInsure.RegistDelSwap(mCurBillType.lng原结帐ID, mCurBillType.intInsure, strAdvance) Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    gcnOracle.CommitTrans
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistDelSwap, True, mCurBillType.intInsure)

    If str个人帐户 <> "" Then
        MsgBox "医保不支持[" & str个人帐户 & "]回退，将退为其它结算方式。" & vbCrLf & vbCrLf & "退款共计:" & Format(GetCashMoney(strNo), "0.00") & " 元。", vbInformation, gstrSysName
    End If

    '2.显示结算界面
    mCurBillType.lng结算序号 = Val(str结算序号) '记录用于打印红票
    On Error GoTo errHandle
    Dim frmBalance As New frmReplenishTheBalanceDelWin, objDelBalance As New clsCliniDelBalance
    Set objDelBalance.rsBalance = mrsBalance
    Set objDelBalance.rs结算方式 = mrs结算方式
    With objDelBalance
        .intInsure = mCurBillType.intInsure
        .CurDelNos = mCurBillType.strNos
        .AllNos = mCurBillType.strAllNOs
        .PrintNOs = "'" & mCurBillType.str结算单 & "'"
        .PatiUseType = mobjFact.使用类别
        .SaveBilled = True
        .ShareUserID = mobjFact.共享批次ID
        .病人ID = mCurBillType.lng病人ID
        .冲销ID = Val(str结算冲销ID)
        .当前发票号 = ""
        .回收发票 = ""
        .结算序号 = Val(str结算序号)
        .结帐ID = 0
        .缺省结算方式 = mCurBillType.str结算方式
        .退费合计 = -1 * GetDelMoney
        .费别 = mCurBillType.str费别
        .年龄 = mCurBillType.str年龄
        .性别 = mCurBillType.str性别
        .姓名 = mCurBillType.str姓名
        .病人类型 = mCurBillType.str病人类型
        .医保不走票号 = MCPAR.医保不走票号
        .原结帐ID = mCurBillType.lng原结帐ID
        .退费时间 = strDelDate
        .部分退费 = False
        .原样退 = False
    End With
    
    Call GetAsyncKeyState(VK_RETURN)
    If frmBalance.zlChargeWin(Me, mlngModule, mstrPrivs, EM_BalanceDel, mobjPayCards, objDelBalance, MCPAR.分币处理, _
        mcllForceDelToCash, mstr排除结算方式, True) = False Then Exit Function
    If Not gfrmMain Is Nothing Then
        Call zlExeBalanceWinRefrshData(True, objDelBalance)
    End If
    ExecuteDelRegister = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function isRegisterValied(ByRef strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查挂号退号是否合法
    '出参:strNO-返回挂号单号
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-09 16:58:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnSel  As Boolean, strNos As String, strTemp As String, blnTemp As Boolean
    Dim strOperatorName As String, i As Long
    
    On Error GoTo errHandle
    If Not CheckTextLength("退费摘要", txt退费摘要) Then Exit Function
    '检查输入是否正确
    If mCurBillType.strNos = "" Then
        MsgBox "请先确认需要退号的门诊挂号单据。", vbInformation, gstrSysName
        If txtNO.Visible Then txtNO.SetFocus: Exit Function
    End If
    If mCurBillType.str结算单 = "" Then
        MsgBox "未找到对应的补充结算记录,不允许退号!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
        
    '检查本次结算单据中是否存在退费异常单据，若存在，则不允许继续退费
    If CheckIsExistDelErrBill(mCurBillType.str结算单, strOperatorName) Then
        MsgBox "注意：" & vbCrLf & _
            "    本次结算中存在异常的结算记录，请先对其进行重新退费！" & _
            IIf(strOperatorName <> UserInfo.姓名, vbCrLf & "    提示：异常单据是操作员【" & strOperatorName & "】收取的。", ""), vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If

    blnSel = False
    With vsBill
        strNos = ""
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 Then
                strTemp = Trim(.TextMatrix(i, .ColIndex("单据号")))
                If strTemp <> "" Then
                    If InStr(1, strNos & ",", "," & strTemp & ",") = 0 Then
                        strNos = strNos & "," & strTemp
                        blnTemp = False
                        If Not zlCheckRegBillIsExecuted(strTemp, True, blnTemp) Then vsBill.SetFocus: Exit Function
                        If blnTemp Then
                            MsgBox "挂号单" & strTemp & "已经被医生接诊或下过医嘱,不能退号！", vbInformation, gstrSysName
                             vsBill.SetFocus: Exit Function
                        End If
                    End If
                    blnSel = True
                End If
            End If
        Next
    End With
    
    If blnSel = False Then
        MsgBox "请在单据中至少选择一个要退号的挂号单。", vbInformation, gstrSysName
        vsBill.SetFocus: Exit Function
    End If
    
    If strNos <> "" Then strNos = Mid(strNos, 2)
    If InStr(1, strNos, ",") > 0 Then
        MsgBox "不能一次退多个挂号单据,请检查。", vbInformation, gstrSysName
        vsBill.SetFocus: Exit Function
    End If
    strNo = strNos
    '医保检查
    If mCurBillType.intInsure = 0 Then
        MsgBox "当前结算不是医保病人结算,不允许进行退号操作!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If gclsInsure.CheckInsureValid(mCurBillType.intInsure) = False Then Exit Function
    
    '三方卡结算方式有效性检查
    If ThreeBalanceCheck(mobjPayCards, mrsBalance, mcllForceDelToCash, mstr排除结算方式) = False Then Exit Function
    
    isRegisterValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ExecuteDelChargeFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行退收费费用操作
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-11 11:07:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmBalance As frmReplenishTheBalanceDelWin, objDelBalance As New clsCliniDelBalance
    Dim arrNo As Variant, k As Long, i As Long, j As Long, lngCount As Long
    Dim lngCheck病人ID As Long, intCheckInsure As Integer
    Dim strBalanceInfor As String, strCurSelNos As String, strNo As String, str序号 As String
    Dim strTemp As String, strReclaimInvoice As String, strInvoice As String, strYBPati As String
    Dim str结算冲销ID As String, str重结ID As String, strSQL As String
    Dim str结帐ID As Long, str冲销ID As Long, str结算序号 As Long, lng领用ID As Long
    Dim blnAll部份退费 As Boolean, blnCur部份退费 As Boolean, bln全退 As Boolean, blnTrans As Boolean
    Dim cllPro As Collection, colOrder As New Collection
    Dim cur个帐透支 As Currency
    Dim varTemp As Variant
    Dim dtDelDate As Date
    Dim strReturn As String, strReturnRecipt As String '退费处方信息，格式：NO,药房ID|NO,药房ID|…
    
    If isChargeFeeValied = False Then Exit Function
    
    On Error GoTo Errhand:
    '先判断所有单据是否部份退费,以决定票据的处理方式
    arrNo = Split(mCurBillType.strNos, ",")
    
    blnAll部份退费 = False
    strCurSelNos = ""
    Set cllPro = New Collection
    For i = 0 To UBound(arrNo)
        strNo = arrNo(i)
        str序号 = "":   lngCount = 0
        '收集当前单据要退费的行号
        With vsBill
            k = vsBill.FindRow(strNo, , .ColIndex("单据号"))
            For j = k To vsBill.Rows - 1
                If vsBill.TextMatrix(j, .ColIndex("单据号")) <> strNo Then Exit For
                If Val(vsBill.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                    str序号 = str序号 & "," & CLng(vsBill.RowData(j))
                    If InStr(1, strCurSelNos & ",", "," & strNo & ",") = 0 Then
                        strCurSelNos = strCurSelNos & "," & strNo
                    End If
                    '81190,冉俊明,退费业务向发药机上传退费信息
                    '格式：NO,药房ID|NO,药房ID|…
                    If vsBill.TextMatrix(j, vsBill.ColIndex("类别")) Like "*西*药*" _
                        Or vsBill.TextMatrix(j, vsBill.ColIndex("类别")) Like "*中*药*" Then
                        If InStr(strReturnRecipt & "|", _
                            "|" & vsBill.TextMatrix(j, vsBill.ColIndex("单据号")) & "," & vsBill.TextMatrix(j, vsBill.ColIndex("执行科室ID")) & "|") = 0 Then
                            strReturnRecipt = strReturnRecipt & "|" & vsBill.TextMatrix(j, vsBill.ColIndex("单据号")) & "," & vsBill.TextMatrix(j, vsBill.ColIndex("执行科室ID"))
                        End If
                    End If
                End If
                lngCount = lngCount + 1
            Next
        End With
        str序号 = Mid(str序号, 2)
        If str序号 <> "" Then
            blnCur部份退费 = Not BillDeleteAllNew(strNo, 1)
            If UBound(Split(str序号, ",")) + 1 = lngCount And blnCur部份退费 = False Then str序号 = ""
            blnCur部份退费 = Not (Not blnCur部份退费 And str序号 = "")
            If blnCur部份退费 Then blnAll部份退费 = True '这张单据为部份退费,则所有单据为部份退费
            colOrder.Add str序号, "_" & strNo
        Else
            blnAll部份退费 = True                       '这张单据不退费,则所有单据为部份退费
            colOrder.Add "未选择", "_" & strNo
        End If
    Next
    
    '根据其它单据是否未退完,则可判断出所有单据是否部份退费
    If Not blnAll部份退费 Then
        varTemp = Split(mCurBillType.strAllNOs, ",")
        strTemp = ""
        For i = 0 To UBound(varTemp)
            If InStr(1, "," & mCurBillType.strNos & ",", "," & varTemp(i) & ",") = 0 Then
                strTemp = strTemp & "," & varTemp(i)
                 blnAll部份退费 = True: Exit For
            End If
        Next
    End If
    
    If CheckSelectItemCanDel(strCurSelNos) = False Then Exit Function
    
    '显示回收票据
    If ShowReclaimInvoice(mCurBillType.str结算单, strReclaimInvoice) = False Then Exit Function
    
    If mCurBillType.intInsure <> 0 And MCPAR.医保接口打印票据 Then
        If zlGetInvoiceGroupUseID(lng领用ID) = False Then Exit Function
        strInvoice = GetNextBill(lng领用ID)
    End If
    
    dtDelDate = zlDatabase.Currentdate
    bln全退 = True
'    If blnAll部份退费 Then bln全退 = False
    If bln全退 Then bln全退 = CheckIsAllDel(mCurBillType.strAllNOs)
     '先退医保
    If Not bln全退 Then
        '可能存在重新收费,因此,需要调用身份验证接口(Identifiy)
        'strAdvace:医保部分退时:传入1,表示医保部分退后再重新收费的身份验证;其他传入: 空
        lngCheck病人ID = mCurBillType.lng病人ID
        intCheckInsure = mCurBillType.intInsure
        strYBPati = gclsInsure.Identify(0, lngCheck病人ID, intCheckInsure, 2)
        
        If strYBPati = "" Then
            MsgBox "医保身份验证失败,不允许继续退费!", vbOKOnly + vbDefaultButton1 + vbExclamation, gstrSysName
            Exit Function
        End If
        
        If Val(CLng(Split(strYBPati, ";")(8))) <> mCurBillType.lng病人ID Then
            MsgBox "医保验证的病人与退费的病人不是同一个病人!", vbInformation, gstrSysName
            Call ExecuteYBIdentifyCancel(mCurBillType.lng病人ID, mCurBillType.intInsure)
            Exit Function
        End If
    End If
    
       
    '保存数据:生成要执行的SQL
    str冲销ID = zlDatabase.GetNextId("病人结帐记录")
    str结算序号 = -1 * str冲销ID
    mCurBillType.strDelNOs = ""
    For i = UBound(arrNo) To 0 Step -1
        strNo = arrNo(i)
        If colOrder("_" & strNo) <> "未选择" Then
            ' Zl_门诊收费记录_销帐
            strSQL = "Zl_门诊收费记录_销帐("
            '  No_In         门诊费用记录.No%Type,
            strSQL = strSQL & "'" & strNo & "',"
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  序号_In       Varchar2 := Null,
            strSQL = strSQL & "'" & colOrder("_" & strNo) & "',"
            '  退费时间_In   门诊费用记录.登记时间%Type := Null,
            strSQL = strSQL & "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  退费摘要_In   门诊费用记录.摘要%Type := Null,
            strSQL = strSQL & "" & IIf(Trim(txt退费摘要.Text) = "", "NULL", "'" & Trim(txt退费摘要.Text) & "'") & ","
            '  结帐id_In     病人预交记录.结帐id%Type := Null,
            strSQL = strSQL & str冲销ID & ","
            '  回收票据_In Number:=0
            strSQL = strSQL & "0)"  '结算记录中进行回收
            zlAddArray cllPro, strSQL
            mCurBillType.strDelNOs = mCurBillType.strDelNOs & "," & strNo
        End If
    Next
     
    ' Zl_门诊退费结算_Modify
    strSQL = "Zl_门诊退费结算_Modify("
    '  操作类型_In   Number,
    '  --   0-原样退
    '  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
    '  --   1-普通退费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --     ②冲预交_In:如果涉及预交款,则传入本次的退预交,非正常收费时,传入零(<0 表示退预交款;>0 表示将剩余款生成预交记录
    '  --   2.三方卡退费结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②退预交_In: 传入零
    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --     ②退预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  --   4-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    '  --     ②退预交_In: 传入零
    '  --     ③退支票额_In:传入零
    strSQL = strSQL & "" & 1 & ","
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & mCurBillType.lng病人ID & ","
    '  冲销id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & str冲销ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "" & "NULL" & ")"
    '  退预交_In     病人预交记录.冲预交%Type := Null,
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    '  卡号_In       病人预交记录.卡号%Type := Null,
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    '  缴款_In       病人预交记录.缴款%Type := Null,
    '  找补_In       病人预交记录.找补%Type := Null,
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    '  完成退费_In   Number := 0,
    '  原结帐id_In   病人预交记录.结帐id%Type := Null
    zlAddArray cllPro, strSQL
    
    str结算冲销ID = zlDatabase.GetNextId("病人结帐记录")
    str重结ID = zlDatabase.GetNextId("病人结帐记录")
    
    '先冲销原始的结算记录
    'Zl_费用补充记录_Delete
    strSQL = "Zl_费用补充记录_Delete("
    '  No_In         In 费用补充记录.No%Type,
    strSQL = strSQL & "'" & mCurBillType.str结算单 & "',"
    '  冲销id_In     In 费用补充记录.结算id%Type,
    strSQL = strSQL & "" & str结算冲销ID & ","
    '  重结id_In     In 费用补充记录.结算id%Type,
    strSQL = strSQL & "" & str重结ID & ","
    '  结算序号_In   In 费用补充记录.结算序号%Type,
    strSQL = strSQL & "" & str结算序号 & ","
    '  退费结帐id_In In 费用补充记录.结算id%Type,(费用退费记录的结帐ID)
    strSQL = strSQL & "" & str冲销ID & ","
    '  操作员编号_In In 费用补充记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In In 费用补充记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  登记时间_In   In 费用补充记录.登记时间%Type := Null
    strSQL = strSQL & "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    zlAddArray cllPro, strSQL
    
    
    '调用医保接口
    '先回收票据，预结算之后再产生票据
    If MCPAR.医保接口打印票据 Then
        If Not bln全退 Then '预结算之后再发出票据
            '56963,77058
            strSQL = "zl_门诊收费记录_RePrint('" & mCurBillType.str结算单 & "',NULL," & ZVal(lng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            zlAddArray cllPro, strSQL
        Else '全退费也要生存票据号，北京医保
            strSQL = "zl_门诊收费记录_RePrint('" & mCurBillType.str结算单 & "','" & strInvoice & "'," & ZVal(lng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            zlAddArray cllPro, strSQL
        End If
    End If
    
    blnTrans = True
    '1.数据保存:冲销数据,重结数据
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If ExecuteClinicDelSwap(mCurBillType.lng病人ID, mCurBillType.intInsure, Val(str结算冲销ID), mCurBillType.lng原结帐ID) = False Then Exit Function
    
     blnTrans = False
    '2.数据保存:冲销数据,重结数据
    Set cllPro = New Collection
    
    '重新进行收费处理
    If Not bln全退 Then
        '读取个帐余额
        cur个帐透支 = mTy_Insure.dbl个帐透支
        mTy_Insure.dbl帐户余额 = gclsInsure.SelfBalance(mCurBillType.lng病人ID, CStr(Split(strYBPati, ";")(1)), 10, cur个帐透支, mCurBillType.intInsure)
        mTy_Insure.dbl个帐透支 = cur个帐透支
        '更新重收记录的保险信息
        If GetExcutInsureInforUpdateSQL(str结算序号, strBalanceInfor, cllPro) = False Then Exit Function
        blnTrans = True: zlExecuteProcedureArrAy cllPro, Me.Caption, True
        '77058
        If ExcuteInsureReCharge(mCurBillType.lng病人ID, mCurBillType.intInsure, str重结ID, str结算序号, strBalanceInfor, _
                mCurBillType.str结算单, lng领用ID, strInvoice, dtDelDate) = False Then Exit Function
    End If
    
    blnTrans = False
    '4.显示结算界面
    
    mCurBillType.lng结算序号 = Val(str结算序号) '记录用于打印红票
    Set objDelBalance.rsBalance = mrsBalance
    Set objDelBalance.rs结算方式 = mrs结算方式
    
    With objDelBalance
        .intInsure = mCurBillType.intInsure
        .CurDelNos = mCurBillType.strDelNOs
        .AllNos = mCurBillType.strAllNOs
        .PrintNOs = "'" & mCurBillType.str结算单 & "'"
        .PatiUseType = mobjFact.使用类别
        .SaveBilled = True
        .ShareUserID = mobjFact.共享批次ID
        .病人ID = mCurBillType.lng病人ID
        .冲销ID = str结算冲销ID
        .当前发票号 = strInvoice
        .回收发票 = strReclaimInvoice
        .结算序号 = str结算序号
        .结帐ID = str结帐ID
        .缺省结算方式 = mCurBillType.str结算方式
        .退费合计 = -1 * GetDelMoney
        .费别 = mCurBillType.str费别
        .年龄 = mCurBillType.str年龄
        .性别 = mCurBillType.str性别
        .姓名 = mCurBillType.str姓名
        .病人类型 = mCurBillType.str病人类型
        .医保不走票号 = MCPAR.医保不走票号
        .原结帐ID = mCurBillType.lng原结帐ID
        .退费时间 = dtDelDate
        .部分退费 = Not bln全退
        .原样退 = False
    End With
    Call GetAsyncKeyState(VK_RETURN)
    Set frmBalance = New frmReplenishTheBalanceDelWin
    If frmBalance.zlChargeWin(Me, mlngModule, mstrPrivs, EM_BalanceDel, mobjPayCards, objDelBalance, MCPAR.分币处理, _
        mcllForceDelToCash, mstr排除结算方式, False) = False Then Exit Function

    '81190,冉俊明,退费业务向发药机上传退费信息
    On Error Resume Next
    If mblnDrugMachine Then
        Dim rsTemp As ADODB.Recordset, strData As String '门诊处方退药格式：费用ID1,退药数量1;费用ID2,退药数量2;...
        '本次退的减去重收的就是实际退的
        strSQL = "Select Max(Decode(a.记录状态, 2, a.Id, 0)) As 费用id, -1 * Nvl(Sum(a.付数 * a.数次), 0) As 退药数量" & vbNewLine & _
                " From 门诊费用记录 A,(Select Distinct 结帐ID From 病人预交记录 Where 结算序号 = [1]) B" & vbNewLine & _
                " Where a.结帐id = b.结帐ID And Mod(a.记录性质, 10) = 1 And a.收费类别 In ('5', '6', '7')" & vbNewLine & _
                " Group By NO, Nvl(价格父号, 序号)" & vbNewLine & _
                " Having Nvl(Sum(a.付数 * a.数次), 0) <> 0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询本次退费项目", objDelBalance.结算序号)
        Do While Not rsTemp.EOF
            strData = strData & ";" & Nvl(rsTemp!费用id) & "," & Nvl(rsTemp!退药数量)
            rsTemp.MoveNext
        Loop
        If strData <> "" Then
            strData = Mid(strData, 2)
            Call mobjDrugMachine.Operation(gstrDBUser, Val("24-处方退药(完整/部分)"), strData, strReturn)
        End If
    ElseIf mblnDrugPacker Then
        If strReturnRecipt <> "" Then
            strReturnRecipt = Mid(strReturnRecipt, 2)
            Call mobjDrugPacker.DYEY_MZ_TransRecipeReturn(1, UserInfo.编号, UserInfo.姓名, strReturnRecipt, strReturn)
        End If
    End If
    Err.Clear: On Error GoTo Errhand
    
    If Not gfrmMain Is Nothing Then
        Call zlExeBalanceWinRefrshData(True, objDelBalance)
    End If
    ExecuteDelChargeFee = True
    Exit Function
Errhand:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter = 1 Then Resume
    End If
    Call SaveErrLog
End Function

Private Sub PrintDelBill(ByVal strNo As String, ByVal lng病人ID As Long, _
    ByVal dtDateDel As Date, ByVal bln部分退 As Boolean, _
    ByVal strInvoices As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印相关票据
    '入参:  strNO-当前结算单号
    '       dtDateDel-退费日期
    '编制:刘兴洪
    '日期:2014-10-11 11:36:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInvoiceFormat As Integer, blnPrint As Integer
    Dim str发票号 As String, int票据张数 As Integer
    Dim strSQL As String, strTempNO As String, i As Integer

    On Error GoTo errHandle
    If Not bln部分退 Then
         '税控部件全退时收回处理(全退时，Zl_费用补充记录_Delete中已收回票据)
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strNo)
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        GoTo PrintList:
        Exit Sub
    End If
    '77058
    If bln部分退 And mCurBillType.intInsure <> 0 And MCPAR.医保接口打印票据 Then GoTo PrintList
    If (strInvoices = "无可退票据" Or strInvoices = "") And bln部分退 Then  'a.收回并重新打印门诊收据
        blnPrint = True
        ''0-不打印;1-自动打印;2-提示打印
        blnPrint = True
        If mobjFact.打印方式 = 0 Then blnPrint = False
        If mobjFact.打印方式 = 2 Then
            If MsgBox("是否打印票据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then blnPrint = False
        End If
        
        If blnPrint Then
            intInvoiceFormat = mobjFact.打印格式
            Call zlRePrintReplenishTheBalanceBill(Me, mlngModule, 1, mCurBillType.str结算单, mCurBillType.intInsure, mobjInvoice, mobjFact, True, dtDateDel)
        End If
        GoTo PrintList:
        Exit Sub
    End If


    'b.收费或上一次退时没有打印票据
    If strInvoices <> "无可退票据" And strInvoices <> "" Then
        'c.只收回票据
        strSQL = "Zl_补充结算票据_Reprint('" & strNo & "',Null,0,'" & UserInfo.姓名 & "'," & _
            "To_Date('" & Format(dtDateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,0,'" & strInvoices & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
PrintList:
    '退费发票(红票)打印，91998
    Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_退费收据, mCurBillType.lng病人ID, 0, mCurBillType.intInsure, mobjFact)
    '0-不打印;1-自动打印;2-提示打印
    If mobjFact.打印方式 = 1 Then
        Call zlPrintReplenishTheDelBalanceBill(Me, mlngModule, mCurBillType.lng结算序号, mCurBillType.intInsure, mobjInvoice, mobjFact, True, dtDateDel)
    ElseIf mobjFact.打印方式 = 2 Then
        If MsgBox("是否打印退费票据(红票)？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call zlPrintReplenishTheDelBalanceBill(Me, mlngModule, mCurBillType.lng结算序号, mCurBillType.intInsure, mobjInvoice, mobjFact, True, dtDateDel)
        End If
    End If
    
    If bln部分退 Then
        '打印费用清单
        If zlStr.IsHavePrivs(mstrPrivs, "门诊结算清单") Then
            If mtyMoudlePara.int清单打印方式 = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me, "NO=" & strNo, "药品单位=" & IIf(mtyMoudlePara.bln药房单位, 1, 0), 2)
            ElseIf mtyMoudlePara.int清单打印方式 = 2 Then
                If MsgBox("要打印收费结算清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me, "NO=" & strNo, "药品单位=" & IIf(mtyMoudlePara.bln药房单位, 1, 0), 2)
                End If
            End If
        End If
    End If
    If mCurBillType.intInsure <> 0 And MCPAR.退费后打印回单 And InStr(1, mstrPrivs, "医保退费回单") > 0 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_2", Me, "NO=" & strNo, 2)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
     
Public Function Get实收金额(ByVal strNo As String) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定单据的实收金额
    '返回:返回实收金额
    '编制:刘兴洪
    '日期:2014-09-30 14:03:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) = strNo Then
                Get实收金额 = Get实收金额 + Val(.TextMatrix(i, .ColIndex("实收金额")))
            End If
        Next
    End With
End Function


Private Sub txt退费摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    '选择退费原因
    If KeyCode <> vbKeyReturn Then Exit Sub

    If Trim(txt退费摘要.Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt退费摘要.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If zl_SelectAndNotAddItem(Me, txt退费摘要, Trim(txt退费摘要.Text), "常用退费原因", "常用退费原因选择", True, True) = False Then
        If zlCommFun.IsCharChinese(Trim(txt退费摘要.Text)) = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt退费摘要_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt退费摘要
End Sub
Private Sub txt退费摘要_LostFocus()
    zlCommFun.OpenIme False
    If zlCommFun.ActualLen(txt退费摘要.Text) > 100 Then
        MsgBox "退费摘要最多允许输入100个字符或50个汉字！", vbInformation, gstrSysName
        If txt退费摘要.Visible And txt退费摘要.Enabled Then txt退费摘要.SetFocus
    End If
End Sub

Private Sub txt退费摘要_Change()
    txt退费摘要.Tag = ""
End Sub

Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建或关闭结算卡对象
    '编制:刘兴洪
    '日期:2014-09-30 14:04:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    Set mobjSquare = gobjSquare.objSquareCard
    If mbytMode = 0 Then Exit Sub
    If gobjSquare.objSquareCard Is Nothing Then
        '创建对象
        Call CreateSquareCardObject(gfrmMain, mlngModule)
    End If
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Dim objCard As Card
    Set objCard = IDKind.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    Set mobjSquare = gobjSquare.objSquareCard
End Sub


Private Function CheckBillIsAllDels(ByVal strNo As String, _
    Optional ByRef strSel序号 As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定的单据是否全部选中退费
    '入参:strNO-单据号
    '出参:strSel序号-返回选中的序号
    '返回:0-全部未选择;1-全部选择;2-选择了一部分
    '编制:刘兴洪
    '日期:2014-09-30 14:06:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim k As Long, j As Long, lngCount As Long, str序号 As String
    With vsBill
        k = vsBill.FindRow(strNo, , vsBill.ColIndex("单据号"))
         For j = k To vsBill.Rows - 1
             If vsBill.TextMatrix(j, .ColIndex("单据号")) <> strNo Then Exit For
             If Val(vsBill.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                 str序号 = str序号 & "," & CLng(vsBill.RowData(j))
             End If
             lngCount = lngCount + 1
         Next
     End With

     If str序号 <> "" Then str序号 = Mid(str序号, 2)
     strSel序号 = str序号
     If str序号 = "" Then CheckBillIsAllDels = 0: Exit Function
     
     If lngCount = UBound(Split(str序号, ",")) + 1 Then
        If InStr(1, mCurBillType.strNosPatiDel & ",", "," & strNo & ",") > 0 Then
            CheckBillIsAllDels = 2: Exit Function
        End If
        CheckBillIsAllDels = 1: Exit Function
     End If
    CheckBillIsAllDels = 2
End Function
Private Sub ReInitPatiInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化病人发票信息
    '编制:刘兴洪
    '日期:2014-10-11 11:39:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytMode = EM_RBDTY_查看 Then Exit Sub
    Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_收费收据, mCurBillType.lng病人ID, 0, mCurBillType.intInsure, mobjFact)
    Call ZlShowBillFormat(mlngModule, lblFormat, mobjFact.打印格式)
End Sub

Private Function zlGetInvoiceGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-30 14:15:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjInvoice.zlGetInvoiceGroupID(mlngModule, UserInfo.姓名, EM_收费收据, mobjFact.使用类别, lng领用ID, mobjFact.共享批次ID, lng领用ID, intNum, strInvoiceNO) = False Then Exit Function
    If lng领用ID <= 0 Then
        Select Case lng领用ID
            Case 0 '操作失败
            Case -1
                If Trim(mobjFact.使用类别) = "" Then
                    MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "你没有自用和共用的『" & mobjFact.使用类别 & "』收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mobjFact.使用类别) = "" Then
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "本地的共用票据的『" & mobjFact.使用类别 & "』收费票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
        End Select
    End If
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub ClearBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除结算数据
    '编制:刘兴洪
    '日期:2014-09-30 14:16:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsBalance
        .Clear 1: .COLS = 1
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Sub LoadBalanceInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收款结算方式
    '编制:刘兴洪
    '日期:2014-09-30 14:17:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String, lngRow As Long
    Dim lngCol As Long, i As Long, intSign As Integer
    Dim lngNullCol As Long
    
    
    If mrsBalance Is Nothing Then Exit Sub
    If mrsBalance.State <> 1 Then Exit Sub
    intSign = IIf(mstrDelTime <> "", -1, 1) '数量,金额正负符号
    '字段:类型 ,结帐ID, 记录性质, 结算方式, 摘要, 卡类别ID, 卡类别名称, 自制卡, 结算卡序号, 结算号码, 卡号, 交易流水号, 交易说明, 结算序号, 校对标志, 医保, 消费卡id
    '            是否密文,是否全退,是否退现,冲预交
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    lngRow = 0
    mrsBalance.Filter = "类型=2"
    mrsBalance.Sort = "类型,结算方式"
    
    '1.加载医保结算
    With vsBalance
        .Redraw = flexRDNone
        Call ClearBalance
        
        .TextMatrix(lngRow, 0) = "保险结算"
        
        If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
        Do While Not mrsBalance.EOF
            '--问题:52530
            str结算方式 = Nvl(mrsBalance!结算方式)
            
            If str结算方式 <> "" Then
                '先查找是否存在相同的结算方式,存在直接汇总
                lngCol = -1
                For i = 1 To .COLS - 1 Step 2
                    If str结算方式 = .Cell(flexcpData, lngRow, i) Then
                        lngCol = i: Exit For
                    End If
                Next
                If lngCol = -1 Then
                    .COLS = .COLS + 2
                    .ColAlignment(.COLS - 2) = 7: .ColAlignment(.COLS - 1) = 1
                    lngCol = .COLS - 2
                End If
                .TextMatrix(lngRow, lngCol) = str结算方式 & ":"
                .Cell(flexcpData, lngRow, lngCol) = str结算方式
                .TextMatrix(lngRow, lngCol + 1) = zlFormatNum(Val(.TextMatrix(lngRow, .COLS - 1)) + intSign * Val(Nvl(mrsBalance!冲预交, 0)))
                
                .Cell(flexcpData, lngRow, lngCol + 1, lngRow, lngCol + 1) = Val(Nvl(mrsBalance!是否退现))
                If mbytMode = EM_RBDTY_退费 Then
                    .ColData(lngCol) = "摘要:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!摘要))
                    .ColData(lngCol + 1) = "结算号码:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!结算号码))
                ElseIf mbytMode = EM_RBDTY_异常重退 Then
                    .ColData(lngCol) = "摘要:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!摘要))
                    .ColData(lngCol + 1) = "结算号码:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!结算号码))
                Else
                    If mstrDelTime <> "" Then
                        .ColData(lngCol) = "摘要:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!摘要))
                        .ColData(lngCol + 1) = "结算号码:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!结算号码))
                    Else
                        .ColData(lngCol) = "摘要:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, Nvl(mrsBalance!摘要), "")
                        .ColData(lngCol + 1) = "结算号码:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, Nvl(mrsBalance!结算号码), "")
                    End If
                End If
            End If
             mrsBalance.MoveNext
            .Redraw = flexRDBuffered
         Loop
         
         '合并为零的记录
         i = 1
         Do While i < .COLS - 1
            If Trim(.TextMatrix(lngRow, i + 1)) = "" Then
                For lngCol = i To .COLS - 3
                    .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol + 2)
                    .Cell(flexcpData, lngRow, lngCol) = .Cell(flexcpData, lngRow, lngCol + 2)
                    .ColData(lngCol) = .ColData(lngCol + 2)
                    .Cell(flexcpForeColor, lngRow, lngCol) = .Cell(flexcpForeColor, lngRow, lngCol + 2)
                    .Cell(flexcpForeColor, 1, lngCol) = .Cell(flexcpForeColor, 1, lngCol + 2)
                    .Cell(flexcpFontBold, 1, lngCol) = .Cell(flexcpFontBold, 1, lngCol + 2)
                Next
                .COLS = .COLS - 2
            Else
                i = i + 2
            End If
         Loop
         
         '加载非医保退费部分(退款部分),不支持预交款退款,所以不处理其他的
        mrsBalance.Filter = "类型<>2"
        mrsBalance.Sort = "类型,结算方式"
        If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
        lngRow = 1
        .TextMatrix(lngRow, 0) = "收退信息"
        Do While Not mrsBalance.EOF
            If Val(mrsBalance!类型) = 1 Then '预交款
                str结算方式 = "冲预存款"
            Else
                str结算方式 = Nvl(mrsBalance!结算方式)
            End If
            If str结算方式 <> "" Then
                '先查找是否存在相同的结算方式,存在直接汇总
                lngCol = -1: lngNullCol = -1
                For i = 1 To .COLS - 1 Step 2
                    If str结算方式 = .Cell(flexcpData, lngRow, i) Then
                        lngCol = i: Exit For
                    End If
                    If .Cell(flexcpData, lngRow, i) = "" And lngNullCol = -1 Then
                        lngNullCol = i
                    End If
                Next
                If lngCol = -1 And lngNullCol <> -1 Then lngCol = lngNullCol
                If lngCol = -1 Then
                    .COLS = .COLS + 2
                    .ColAlignment(.COLS - 2) = 7: .ColAlignment(.COLS - 1) = 1
                    lngCol = .COLS - 2
                End If
                
                .TextMatrix(lngRow, lngCol) = str结算方式 & ":"
                .Cell(flexcpData, lngRow, lngCol) = str结算方式
                .TextMatrix(lngRow, lngCol + 1) = FormatEx(Val(.TextMatrix(lngRow, lngCol + 1)) + intSign * Val(Nvl(mrsBalance!冲预交, 0)), 5)
                
                .Cell(flexcpData, lngRow, lngCol + 1, lngRow, lngCol + 1) = Val(Nvl(mrsBalance!是否退现))
                If mbytMode = EM_RBDTY_退费 Then
                    .ColData(lngCol) = "摘要:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!摘要))
                    .ColData(lngCol + 1) = "结算号码:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!结算号码))
                ElseIf mbytMode = EM_RBDTY_异常重退 Then
                    .ColData(lngCol) = "摘要:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!摘要))
                    .ColData(lngCol + 1) = "结算号码:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!结算号码))
                Else
                    If mstrDelTime <> "" Then
                        .ColData(lngCol) = "摘要:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!摘要))
                        .ColData(lngCol + 1) = "结算号码:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!结算号码))
                    Else
                        .ColData(lngCol) = "摘要:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, Nvl(mrsBalance!摘要), "")
                        .ColData(lngCol + 1) = "结算号码:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, Nvl(mrsBalance!结算号码), "")
                    End If
                End If
            End If
             mrsBalance.MoveNext
            .Redraw = flexRDBuffered
         Loop
         
         '合并为零的记录
         i = 1
         Do While i < .COLS - 1
            If .TextMatrix(lngRow, i + 1) = "" Then
                For lngCol = i To .COLS - 3
                    .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol + 2)
                    .Cell(flexcpData, lngRow, lngCol) = .Cell(flexcpData, lngRow, lngCol + 2)
                    .ColData(lngCol) = .ColData(lngCol + 2)
                    .Cell(flexcpForeColor, lngRow, lngCol) = .Cell(flexcpForeColor, lngRow, lngCol + 2)
                    .Cell(flexcpForeColor, 1, lngCol) = .Cell(flexcpForeColor, 1, lngCol + 2)
                    .Cell(flexcpFontBold, 1, lngCol) = .Cell(flexcpFontBold, 1, lngCol + 2)
                Next
                If Trim(.TextMatrix(0, .COLS - 2)) = "" Then
                    .COLS = .COLS - 2
                Else
                  i = i + 2
                End If
            Else
                i = i + 2
            End If
         Loop
         .RowHidden(1) = False
         vsBalance.AutoSizeMode = flexAutoSizeColWidth
         Call vsBalance.AutoSize(0, .COLS - 1)
          ControlResize
    End With
End Sub
 
Private Sub LoadSelDelTotal()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载选择的退款合计
    '编制:刘兴洪
    '日期:2014-10-11 11:41:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txt退款合计 = Format(GetDelMoney, gstrDec)
End Sub

Private Function GetDelMoney() As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取退款合计
    '返回:获取退款合计
    '编制:刘兴洪
    '日期:2014-07-03 17:24:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl退款合计 As Double, i As Long
    With vsBill
        For i = 1 To .Rows - 1
            If Val(vsBill.TextMatrix(i, .ColIndex("选择"))) <> 0 Or _
                mbytMode = EM_RBDTY_异常重退 Or mbytMode = EM_RBDTY_查看 Then
                dbl退款合计 = dbl退款合计 + Val(vsBill.TextMatrix(i, .ColIndex("实收金额")))
            End If
        Next
    End With
    GetDelMoney = RoundEx(dbl退款合计, 6)
End Function

Private Sub ControlResize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调整控件位置
    '编制:刘兴洪
    '日期:2014-10-11 11:42:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnFind As Boolean
    With vsBalance
        For i = 1 To .COLS - 1 Step 2
            If .Cell(flexcpData, 1, i) <> "" Then
                blnFind = True: Exit For
            End If
        Next
        If blnFind = False Then .RowHidden(1) = True
        .Height = IIf(.RowHidden(1), 375, 735)
    End With
    '85153,挂号补充结算退费时隐藏"退费摘要"
    pic退费摘要.Visible = Not mCurBillType.bln挂号
    
    Form_Resize
End Sub

Private Sub txtPatient_Change()
    Dim blnAutoFind As Boolean
    blnAutoFind = False
    If Me.ActiveControl Is txtPatient And txtPatient.Visible Then
        blnAutoFind = txtPatient.Text = ""
    End If
    
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(blnAutoFind)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(blnAutoFind)
    IDKind.SetAutoReadCard (blnAutoFind)

End Sub

Private Sub txtPatient_GotFocus()
    If txtPatient.Locked Or Not txtPatient.Visible Then Exit Sub
    
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "" And txtPatient.Visible)
    zlControl.TxtSelAll txtPatient
    
End Sub
Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnICCard As Boolean, blnCancel As Boolean
    
    On Error GoTo errH
    
    If txtPatient.Locked Then Exit Sub

    If IDKind.GetCurCard.名称 Like "姓名*" Then
        '103563,只要输入的第一个字符是“-+*”，后面是全数字，都认为不是刷卡
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IDKind.GetCurCard.名称 = "门诊号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        txtPatient.IMEMode = 0
    End If

    If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then Exit Sub

    If KeyAscii <> 13 Then
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
    End If
    KeyAscii = 0
    Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtPatient.Text))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2014-09-30 14:29:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnCancel As Boolean
    If objCard.名称 Like "IC卡*" And objCard.系统 Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    
    'a.根据输入读取病人信息失败
    If Not GetPatient(objCard, Trim(txtPatient.Text), blnCancel, blnCard) Then
        If blnCancel Then '取消输入
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            txtPatient.Text = ""
            Exit Sub
        End If
        stbThis.Panels(2) = "未找到该病人，请检查输入内容!"
        If blnCard = True Then
            txtPatient.PasswordChar = "": txtPatient.Text = ""
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
        Else
            txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
        End If
        Set mrsInfo = New ADODB.Recordset
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Call ClearFace
        Exit Sub
    End If
    mCurBillType.lng病人ID = Val("" & mrsInfo!病人ID)
    txtPatient = Nvl(mrsInfo!姓名)

    lblPati.Caption = "病人:" & "                 " & _
        "　性别:" & Nvl(mrsInfo!性别) & _
        "　年龄:" & Nvl(mrsInfo!年龄) & _
        "　门诊号:" & Nvl(mrsInfo!门诊号) & _
        "　费别:" & Nvl(mrsInfo!费别) & _
        "　付款方式:" & mrsInfo!医疗付款方式
        
    With mCurBillType
        .str性别 = Nvl(mrsInfo!性别)
        .str年龄 = Nvl(mrsInfo!年龄)
        .str姓名 = Nvl(mrsInfo!姓名)
        .str费别 = Nvl(mrsInfo!费别)
    End With
    If SelectNO(mCurBillType.lng病人ID) = False Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Call ClearFace
        Exit Sub
    End If
    If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
    blnCancel As Boolean, Optional blnCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:objCard-指定的卡类别
    '     strInput-输入的值
    '     blnCancel-
    '     blnCard-是否刷卡
    '返回:读取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-30 14:30:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strPati As String, strSQL As String
    Dim vRect As RECT, i As Integer, lng卡类别ID As Long, bln存在帐户 As Boolean, lng病人ID As Long, strPassWord As String, strErrMsg As String
    Dim strWhere As String, blnICCard As Boolean
    Dim blnHavePassWord As Boolean
    
    blnCancel = False
    strWhere = ""
    If blnCard And objCard.名称 Like "姓名*" And InStr("-+*", Left(strInput, 1)) = 0 Then  '103563
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.病人ID=[1]"
        strInput = "-" & lng病人ID
        blnHavePassWord = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then  '病人ID
        strWhere = strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then  '住院号(对住(过)院的病人)
        strWhere = strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号(仅对门诊病人)
        strWhere = strWhere & " And A.门诊号=[1]"
        '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    Else '当作姓名
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                strPati = _
                " Select /*+Rule */A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄," & _
                "           A.住院号,B.名称 as 科室,A.当前床号 as 床号," & _
                "           A.出生日期,A.身份证号,A.家庭地址,A.卡验证码 " & _
                " From 病人信息 A,部门表 B" & _
                " Where A.停用时间 is NULL And A.当前科室ID=B.ID(+) And A.姓名 Like [1]" & _
                "   Order by A.姓名"
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", "bytSize=1")
                If Not rsTmp Is Nothing Then
                    strInput = rsTmp!病人ID
                    strWhere = strWhere & " And A.病人ID=[2]"
                Else
                    Set mrsInfo = New ADODB.Recordset: Exit Function
                End If
            Case "医保号"
                strInput = UCase(strInput)
                strWhere = strWhere & " And A.医保号=[2]"
            Case "身份证号", "二代身份证", "身份证"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
                blnICCard = (InStr(1, "-+*.", Left(strInput, 1)) = 0)
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.门诊号=[2]"
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    bln存在帐户 = objCard.是否存在帐户 = 1
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    strSQL = _
    " Select A.病人ID,Nvl(C.主页ID,0) as 主页ID,A.门诊号,Nvl(C.当前病区ID,0) as 病区ID,Nvl(c.出院科室ID,0) as 科室ID,Nvl(A.当前科室ID,0) as 当前科室ID, Nvl(a.在院,0) as 在院," & _
    "           Decode(Nvl(A.主页ID,0),0,A.医疗付款方式,C.医疗付款方式) 医疗付款方式,Nvl(A.病人类型,C.病人类型) as 病人类型," & _
    "           A.姓名,A.性别,A.年龄,Nvl(A.住院号,0) as 住院号,Nvl(C.出院病床,0) as 床号,A.家庭地址,A.卡验证码," & _
    "           B.险类,B.卡号,Nvl(B.医保号,A.医保号) 医保号,B.密码,Nvl(C.费别,A.费别) 费别,A.担保人,A.担保额,Nvl(A.担保性质,0) as 担保性质, C.备注 " & _
    " From 病人信息 A,医保病人档案 B,病案主页 C,医保病人关联表 E " & _
    " Where A.停用时间 is NULL" & _
    "       And A.病人ID=C.病人ID(+) And Nvl(A.主页ID,0)=C.主页ID(+)" & _
    "       And C.病人ID=E.病人ID(+) And E.标志(+)=1  " & _
    "       And E.医保号=B.医保号(+) And E.险类=B.险类(+) And E.中心 = B.中心(+) " & strWhere

    On Error GoTo errH
    txtPatient.ForeColor = &HC00000: lblPati.ForeColor = txtPatient.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If mrsInfo.EOF Then
        Set mrsInfo = New ADODB.Recordset: Exit Function
    End If
    Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(IsNull(mrsInfo!险类), &HC00000, vbRed))
    lblPati.ForeColor = txtPatient.ForeColor
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    txtPatient.ForeColor = &HC00000
    lblPati.ForeColor = txtPatient.ForeColor
End Function

Private Function SelectNO(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID选择合适的退费单据
    '入参:lng病人ID-获取病人ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-09-30 14:47:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, blnCancel As Boolean
    Dim strNo As String, intFindType As Integer
    
    On Error GoTo errHandle
    '80602,冉俊明,2014-12-8,不提取结算作废的补充结算单据（费用状态=2）
    strSQL = "" & _
        "  With 收费单 as ( " & _
        "           Select b.结算单号,Max(a.ID) as ID,max(b.结算序号) as 结算ID ,max(A.结帐ID) as 结帐ID, " & _
        "                  max(decode(a.记录性质,4,'挂号','收费')) as 单据,  max(mod(a.记录性质,10)) as 记录性质ID,a.No as 单据号,  c.名称 as 开单部门, a.开单人, a.操作员编号, a.操作员姓名, a.实际票号, To_Char(a.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 发生时间, " & vbCrLf & _
        "                  To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间 " & vbCrLf & _
        "           From  门诊费用记录 A,( Select distinct  A.NO as 结算单号, A.结算序号,A.收费结帐ID,b.记录性质,b.NO,A.附加标志" & _
        "                   From 费用补充记录 A,门诊费用记录 B " & _
        "                   Where a.收费结帐ID=b.结帐ID And A.病人ID=[1] And nvl(a.费用状态,0)=0) B, " & _
        "                   部门表 C " & vbCrLf & _
        "           Where  A.结帐ID=b.收费结帐ID And nvl(A.附加标志,0)<>9 and A.开单部门ID=C.ID(+)  And a.记录状态 in (1,3) " & vbCrLf & _
        "                And Nvl(a.执行状态, 0) <> 1 And Nvl(a.费用状态, 0) <> 1 " & vbCrLf & _
        "              " & vbCrLf & _
        "          Group by b.结算单号,mod(a.记录性质,10),a.No,a.开单人,c.名称,a.操作员编号, a.操作员姓名, a.实际票号, To_Char(a.发生时间, 'yyyy-mm-dd hh24:mi:ss'),To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') " & vbCrLf & _
        "           )"

     strSQL = strSQL & vbCrLf & _
     "  Select J.*  " & vbCrLf & _
     "  From 收费单 J," & vbCrLf & _
     "           ( Select mod(A.记录性质,10) as 记录性质,A.NO,sum(nvl(A.付数,1)*nvl(A.数次,1)) 数次" & vbCrLf & _
     "             From 门诊费用记录 A,收费单 B  " & vbCrLf & _
     "             Where A.NO=B.单据号 And mod(A.记录性质,10)= b.记录性质ID  And a.价格父号 is null  " & vbCrLf & _
     "             Group by A.记录性质,A.NO " & vbCrLf & _
     "             Having sum(nvl(A.付数,1)*nvl(A.数次,1))>0 ) M" & vbCrLf & _
     "  Where J.单据号=M.NO and J.记录性质ID=M.记录性质 " & vbCrLf
     strSQL = "Select * From (" & strSQL & ") Order by 记录性质ID,登记时间 desc,单据号"
     
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "退费单据", 1, "", "请选择需要退费的单据", False, False, False, 0, 0, 0, blnCancel, False, False, lng病人ID, "bytSize=1")
    If blnCancel = True Then Exit Function
    If rsTemp Is Nothing Then
        MsgBox "该病人不存在补结算费用,请在病人收费管理中进行退费", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "该病人不存在补结算费用,请在病人收费管理中进行退费", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Dim int记录性质 As Integer
    
    strNo = Nvl(rsTemp!单据号): int记录性质 = Nvl(rsTemp!记录性质ID)
    mCurBillType.str结算单 = Nvl(rsTemp!结算单号)
    intFindType = IIf(int记录性质 = 4, 4, 1)
    
    If Not ReadBills(intFindType, strNo) Then
        Call ClearFace: Exit Function
    End If
    SelectNO = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    Dim objCard As Card
    
    If strNo = "" Then Exit Sub
    If Not Me.ActiveControl Is txtPatient _
        Or txtPatient.Locked Or txtPatient.Text <> "" Then Exit Sub
        
    Set objCard = IDKind.GetIDKindCard("IC卡号", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    Call FindPati(objCard, False, strNo)
    If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
End Sub
Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    Dim objCard As Card
    
    If strID = "" Then Exit Sub
    If Not Me.ActiveControl Is txtPatient _
        Or txtPatient.Locked Or txtPatient.Text <> "" Then Exit Sub
        
    Set objCard = IDKind.GetIDKindCard("身份证号", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    Call FindPati(objCard, False, strID)
    If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
End Sub

  
Private Sub SynchronizationSelect(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前行,同步选择关联行
    '编制:刘兴洪
    '日期:2014-10-11 11:51:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    
    With vsBill
        If Val(.Cell(flexcpData, lngRow, .ColIndex("项目"))) = 0 Then
            For i = lngRow + 1 To vsBill.Rows - 1
                 If Val(vsBill.RowData(lngRow)) = Val(vsBill.Cell(flexcpData, i, .ColIndex("项目"))) Then
                       vsBill.TextMatrix(i, .ColIndex("选择")) = vsBill.TextMatrix(lngRow, .ColIndex("选择"))
                 Else
                    Exit For
                 End If
            Next
            Call zlSet诊疗固定关系(lngRow, .ColIndex("选择"))
            Exit Sub
        End If
        
        Call zlSet诊疗固定关系(lngRow, .ColIndex("选择"))
        '需要检查主项是否已经被
        For i = lngRow - 1 To 1 Step -1
            If Val(.RowData(i)) = Val(.Cell(flexcpData, lngRow, .ColIndex("项目"))) Then
                If .TextMatrix(i, .ColIndex("选择")) <> 0 Then
                     .TextMatrix(i, .ColIndex("选择")) = .TextMatrix(lngRow, .ColIndex("选择"))
                End If
                Call zlSet诊疗固定关系(i, .ColIndex("选择"), lngRow)
                 Exit For
            End If
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
 
Public Function CheckDiff(strNos As String, strDiffNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:比较两个单据号是否一致
    '返回:全部一致,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-21 17:19:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long

    On Error GoTo errHandle
    varTemp = Split(Replace(strDiffNos, "'", ""), ",")
    varData = Split(Replace(strNos, "'", ""), ",")
    If UBound(varTemp) <> UBound(varData) Then Exit Function
    For i = 0 To UBound(varData)
        If InStr(1, "," & strDiffNos & ",", "," & varData(i) & ",") = 0 Then Exit Function
    Next
    For i = 0 To UBound(varTemp)
        If InStr(1, "," & strNos & ",", "," & varTemp(i) & ",") = 0 Then Exit Function
    Next
    CheckDiff = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub initInsurePara(ByVal intInsure As Integer, ByVal lng病人ID As Long, ByVal lng结帐ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医保参数
    '编制:刘兴洪
    '日期:2014-06-26 16:25:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If intInsure = 0 Then Exit Sub
    MCPAR.门诊结算作废 = gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure)
    MCPAR.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure, CStr(lng结帐ID))
    MCPAR.退费后打印回单 = gclsInsure.GetCapability(support退费后打印回单, lng病人ID, intInsure)
    MCPAR.门诊预结算 = gclsInsure.GetCapability(support门诊预算, lng病人ID, intInsure)
    MCPAR.先自付 = gclsInsure.GetCapability(support收费帐户首先自付, lng病人ID, intInsure)
    MCPAR.全自付 = gclsInsure.GetCapability(support收费帐户全自费, lng病人ID, intInsure)
    MCPAR.分币处理 = gclsInsure.GetCapability(support分币处理, lng病人ID, intInsure)
End Sub

Private Sub SetFunCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置功能控件的visible属性
    '编制:刘兴洪
    '日期:2014-10-11 11:54:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cmdSelAll.Visible = mbytMode = EM_RBDTY_退费
    cmdClear.Visible = mbytMode = EM_RBDTY_退费
    cmdBillSel.Visible = mbytMode = EM_RBDTY_退费
    If mstr结算序号 <> "" Then   '外面传入时,不用手工输入
        txtNO.Visible = False
        IDKindNO.Visible = False
        picPatiBack.Visible = False
        fraInfo_1.Visible = False
    End If
End Sub

Private Function GetYBOldBalance(ByVal lng病人ID As Long, ByVal intInsure As Integer, ByVal lng原结帐ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医保原结算方式和结算金额
    '返回:返回结算信息,格式:结算方式|结算金额||...
    '编制:刘兴洪
    '日期:2014-07-07 09:57:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String
    
    On Error GoTo errHandle
    If intInsure = 0 Then Exit Function
    
    '字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '     是否密文,是否全退,是否退现,冲预交
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    mrsBalance.Filter = "类型=2 and 结帐ID=" & lng原结帐ID
    If mrsBalance.RecordCount = 0 Then Exit Function
    With mrsBalance
        Do While Not .EOF
            '如果这种结算方式不支持回退,要退为现金,则不用减去
            If MCPAR.门诊结算作废 Then
                If gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure, !结算方式) Then
                    str结算方式 = str结算方式 & "||" & !结算方式 & "|" & Val(Nvl(!冲预交))
                End If
            Else     '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                If !结算方式 <> mstr个人帐户 Then
                    str结算方式 = str结算方式 & "||" & !结算方式 & "|" & Val(Nvl(!冲预交))
                End If
            End If
            .MoveNext
        Loop
    End With
    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 3)
    GetYBOldBalance = str结算方式
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExcuteInsureReCharge(ByVal lng病人ID As Long, ByVal intInsure As Integer, _
    ByVal lng结帐ID As Long, ByVal lng结算序号 As Long, ByVal strBalnaceInfor As String, _
    ByVal strNo As String, ByVal lng领用ID As Long, ByVal strInvoice As String, ByVal dtDelDate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行医保重新收费
    '入参:strBalnaceInfor:结算信息,格式为:实收合计;进入统筹;全自付;先自
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-11 11:55:07
    '说明:参数strNO,lng领用ID,strInvoice,dtDelDate用于医保接口打印票据
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, arrBalance As Variant, str结算方式 As String
    Dim dbl结算金额 As Double, dbl可分配额 As Double, dbl余额 As Double
    Dim strBalance As String, dbl退款合计 As Double, str退回结算 As String
    Dim strSQL As String, rsTmp As ADODB.Recordset, strYbInvoice As String
    Dim i As Long, k As Long, j As Long, cur误差金额 As Double
    Dim strNone As String, strNos As String, varTemp As Variant, cur个帐 As Currency
    
    On Error GoTo errHandle
    If mCurBillType.intInsure = 0 Then
        ExcuteInsureReCharge = False
        gcnOracle.RollbackTrans: Exit Function
    End If
    
    
    strBalance = ""
    If Not MCPAR.门诊预结算 Then '计算个人帐户支付金额
        varTemp = Split(strBalnaceInfor, ";") 'cur实收合计;cur进入统筹;cur全自付;cur先自付
        If mstr个人帐户 <> "" And mTy_Insure.dbl帐户余额 > -1 * mTy_Insure.dbl个帐透支 Then
            If RoundEx(Val(varTemp(0)), 6) >= 0 Then
                cur个帐 = RoundEx(Val(varTemp(1)), 6) + IIf(MCPAR.先自付, RoundEx(Val(varTemp(3)), 6), 0) + IIf(MCPAR.全自付, RoundEx(Val(varTemp(2)), 6), 0)
                If mTy_Insure.dbl帐户余额 - cur个帐 >= -1 * mTy_Insure.dbl个帐透支 Then
                    strBalance = mstr个人帐户 & "|" & cur个帐   '在允许透支范围内足够(允许透支0为特例)
                Else
                    If mTy_Insure.dbl个帐透支 = 0 And mTy_Insure.dbl帐户余额 > 0 Then
                        strBalance = mstr个人帐户 & "|" & mTy_Insure.dbl帐户余额  '不允许透支且有余额
                    Else
                        '超过允许透支范围或不允许透支时无余额
                        If mTy_Insure.dbl个帐透支 <> 0 Then
                            strBalance = mstr个人帐户 & "|" & mTy_Insure.dbl帐户余额 + mTy_Insure.dbl个帐透支 '在允许透支范围内支付
                        Else
                            strBalance = mstr个人帐户 & "|0"
                        End If
                    End If
                End If
            Else
                strBalance = mstr个人帐户 & "|0"
            End If
        End If
    Else
        If ExecuteClinicPreSwap(intInsure, lng结帐ID, lng病人ID, strBalance, strNone, strYbInvoice, strNos) = False Then
            gcnOracle.RollbackTrans
            If strNone <> "" Then
                MsgBox "当前保险结算使用的结算方式" & vbCrLf & vbCrLf & vbTab & strNone & vbCrLf & vbCrLf & _
                    "在门诊未设置，请先到结算方式管理中设置这些结算方式！", vbInformation, gstrSysName
            End If
            Exit Function
        End If
    End If
    
    'Zl_费用补充结算_Modify
    strSQL = "Zl_费用补充结算_Modify("
    '  操作类型_In   Number,
    '  --   0-普通结算方式:
    '  --     结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --   1.三方卡结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    strSQL = strSQL & "" & 2 & ","
    '  结算id_In     In 费用补充记录.结算id%Type,
    strSQL = strSQL & "" & lng结帐ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "'" & strBalance & "')"
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    '  卡号_In       病人预交记录.卡号%Type := Null,
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    '  完成结算_In Number:=0
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If MCPAR.医保接口打印票据 And MCPAR.医保不走票号 = False Then
        '38821,77058
        '票据数据生成(因为不调HIS的打印，医保接口打印，所以先填票据数据)
        strSQL = "zl_门诊收费记录_RePrint('" & strNo & "','" & strInvoice & "'," & ZVal(lng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                  "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    '调用医保结算接口
    If ExecuteClinicSwap(lng病人ID, intInsure, lng结帐ID, lng结算序号, strBalance, strNos, strBalnaceInfor) = False Then Exit Function
    ExcuteInsureReCharge = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mCurBillType.intInsure)
End Function

Private Function ExecuteClinicPreSwap(ByVal intInsure As Integer, _
    ByVal lng结帐ID As Long, ByVal lng病人ID As Long, ByRef strBalance As String, _
    ByRef strNone As String, ByRef strYbInvoice As String, ByRef strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行门诊预结算
    '入参:intInsure-险类
    '     lng结帐ID-重新收费的结帐ID
    '出参:strNone-不存在的结算方式
    '     strBalance-返回结算方式(结算方式|金额||...)
    '     strYbInvoice-医保返回的发票号
    '     strNOs-返回本次结算的NOs
    '返回:预结算成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-07 11:18:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoice As String, varData As Variant
    Dim rsTemp As ADODB.Recordset, strAdvance As String
    Dim i As Long, str结算方式 As String
    Dim varTemp As Variant
    
    
    On Error GoTo errHandle
    
    strInvoice = mCurBillType.strInvoice
    Set rsTemp = zlMakeClinicPreSwapData(strInvoice, lng结帐ID, strNos, True)
    
RePreSwap:
    strAdvance = "3": strBalance = ""
    If Not gclsInsure.ClinicPreSwap(rsTemp, strBalance, intInsure, strAdvance) Then
        Screen.MousePointer = 0
        If MsgBox("重新进行医保收费时,单据预结算失败,是否重新进行预结算?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then GoTo RePreSwap:
        Exit Function
    End If
    If strAdvance <> "" And InStr(1, strAdvance, "|") = 0 Then '医保票据号
        strYbInvoice = strAdvance
    End If
    
    MCPAR.医保不走票号 = False
    If InStr(1, strAdvance, ";") > 0 Then
        varData = Split(strAdvance & ";", ";")
        strYbInvoice = Trim(varData(0))
        '38821:strAdvance:发票号;是否不走票据号
        MCPAR.医保不走票号 = Val(varData(1)) = 1
    End If
    '结算方式;原始(最大)金额;可否修改;改后金额
    varData = Split(strBalance, "|")
    
    '结算方式|结算金额||..
    strBalance = "": strNone = ""
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i), ";")
        str结算方式 = varTemp(0)
        mrs结算方式.Filter = "名称='" & str结算方式 & "' And  性质>=3 and 性质<= 4"
        If mrs结算方式.EOF Then
            strNone = strNone & "," & str结算方式
        End If
        strBalance = strBalance & "||" & varTemp(0) & "|" & Val(varTemp(1))
    Next
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    If strNone <> "" Then
        strNone = Mid(strNone, 2): Exit Function
    End If
    
    ExecuteClinicPreSwap = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function ExecuteClinicSwap(ByVal lng病人ID As Long, _
    ByVal intInsure As Integer, ByVal lng结帐ID As Long, _
    ByVal lng结算序号 As Long, ByVal str预结算 As String, _
    ByVal strNos As String, Optional ByVal strBalnaceInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保结算接口
    '入参:  lng结帐ID:本次结帐的ID
    '       strBalnaceInfor:结算信息,格式为:实收合计;进入统筹;全自付;先自
    '返回:医保调用成功或非医保,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-11 11:55:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTrans As Boolean, blnTransMedicare As Boolean
    Dim cur个帐支付 As Currency, cur医保基金 As Currency
    Dim strSQL As String, strAdvance As String
    Dim varTemp As Variant
    Dim i As Long
    
    
    On Error GoTo errHandle
     
    blnTrans = True
    If MCPAR.医保接口打印票据 And MCPAR.医保不走票号 = False Then
        '不严格控制票据时保存当前票号
        If Not gblnStrictCtrl Then
            zlDatabase.SetPara "当前收费票据号", mCurBillType.strInvoice, glngSys, mlngModule, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
        End If
    End If
    
    
    cur个帐支付 = 0: cur医保基金 = 0
    If str预结算 <> "" Then
        varTemp = Split(str预结算, "||")
        For i = 0 To UBound(varTemp)
            If Split(varTemp(i), "|")(0) = mstr个人帐户 Then
                cur个帐支付 = cur个帐支付 + CCur(Val(Split(varTemp(i), "|")(1)))
            ElseIf Split(varTemp(i), "|")(0) = "医保基金" Then
                cur医保基金 = cur医保基金 + CCur(Val(Split(varTemp(i), "|")(1)))
            End If
        Next
    End If
    varTemp = Split(strBalnaceInfor, ";")  'cur实收合计;cur进入统筹;cur全自付;cur先自付
    strAdvance = CStr(lng结算序号)
    If Not gclsInsure.ClinicSwap(lng结帐ID, cur个帐支付, cur医保基金, _
                        CCur(Val(varTemp(2))), CCur(Val(varTemp(3))), intInsure, strAdvance) Then
        gcnOracle.RollbackTrans:  Exit Function
    End If
  
    
    blnTransMedicare = True
    
    If strAdvance = CStr(lng结算序号) Then strAdvance = ""
     
    If strAdvance = "" Then
       gcnOracle.CommitTrans
       Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, intInsure)
       ExecuteClinicSwap = True: Exit Function
    End If
    
    If Not zlInsureCheck(str预结算, strAdvance) Then
       gcnOracle.CommitTrans
       Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, intInsure)
       ExecuteClinicSwap = True: Exit Function
    End If
     'Zl_费用补充结算_Modify
        strSQL = "Zl_费用补充结算_Modify("
        '  操作类型_In   Number,
        '  --   0-普通结算方式:
        '  --     结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
        '  --   1.三方卡结算:
        '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        '  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
        '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
        strSQL = strSQL & "" & 2 & ","
        '  结算id_In     In 费用补充记录.结算id%Type,
        strSQL = strSQL & "" & lng结帐ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "'" & strAdvance & "')"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        '  卡号_In       病人预交记录.卡号%Type := Null,
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        '  误差金额_In   门诊费用记录.实收金额%Type := Null,
        '  完成结算_In Number:=0
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
         
    gcnOracle.CommitTrans: blnTrans = False
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, intInsure)
    ExecuteClinicSwap = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, False, intInsure)
    Call SaveErrLog
End Function

Private Function ExecuteYBIdentifyCancel(ByVal lng病人ID As Long, ByVal intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消医保病人身份验证
    '返回:返回假时不退出界面或清除操作
    '编制:刘兴洪
    '日期:2014-06-09 14:37:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    ExecuteYBIdentifyCancel = True
    If mbytMode = EM_RBDTY_查看 Or lng病人ID = 0 Then Exit Function
    
    ExecuteYBIdentifyCancel = gclsInsure.IdentifyCancel(0, lng病人ID, intInsure)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Sub zlExeBalanceWinRefrshData(ByVal blnSaveOK As Boolean, _
    ByRef objDelBalance As clsCliniDelBalance)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行退费结算操作后的刷新操作
    '入参:blnSaveOK-是否保存成功
    '     objChargeInfor-结算信息
    '编制:刘兴洪
    '日期:2014-06-17 10:50:41
    '说明:之所要独立出来,主要原因是解决医保调试的问题(模态窗体不好调试)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrintNos As String, strReclaimInvoice As String
    Dim strNo As String
    
    On Error GoTo errHandle
    
    If blnSaveOK = False Then Exit Sub
    
    strPrintNos = objDelBalance.PrintNOs
    
    If gblnBillPrint Then
        If gobjBillPrint.zlEraseBill(strPrintNos, 0) = False Then Exit Sub
    End If
   '打印退费单据
    Call PrintDelBill(strPrintNos, objDelBalance.病人ID, objDelBalance.退费时间, objDelBalance.部分退费, "")

Completed:
    mblnOK = True: Call ClearFace
    If txtNO.Visible Then txtNO.SetFocus: Exit Sub
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function IsFeeAllDel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断是否全退费
    '返回:合退费返回成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-14 16:36:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDel As Boolean, blnAllDel As Boolean
    Dim j As Long
    On Error GoTo errHandle
    '1.看是否为全选，全选就原样退
    If mCurBillType.bln单张部分退费 Then Exit Function
    With vsBill
        For j = 1 To .Rows - 1
            If .TextMatrix(j, .ColIndex("单据号")) <> "" And Abs(Val(.TextMatrix(j, .ColIndex("选择")))) <> 1 Then
                IsFeeAllDel = False: Exit Function
            End If
        Next
    End With
    
    '2.当前退费与本次收费单据完全一致
    If CheckDiff(Replace(mCurBillType.strAllNOs, "'", ""), Replace(mCurBillType.strNos, "'", "")) = False Then Exit Function
    
    
    IsFeeAllDel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetFeeDelNumRecord(ByVal strAllNOs As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取费用的剩余数量集
    '入参:strAllNos-所有单据
    '出参:
    '返回:记录集(NO,序号,原始数量,剩余数量)
    '编制:刘兴洪
    '日期:2014-07-15 11:35:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "" & _
    "   Select A.NO,nvl(A.价格父号,A.序号) as 序号,a.收费细目ID,A.记录性质,A.结帐ID, " & _
    "         Decode(A.记录性质,1, 1,0)*decode(A.记录状态,1,1,3,1,0)*Avg(nvl(A.付数,1) *数次) as 原始数量," & _
    "         Avg(nvl(A.付数,1) *数次) as 数量" & _
    "   From 门诊费用记录 A" & _
    "   Where A.NO in (select J.Column_value From  Table(f_str2List([1])) J )  " & _
    "       And mod(a.记录性质,10)=1 And nvl(A.费用状态,0)<>1" & _
    "   Group by A.NO,nvl(A.价格父号,A.序号),A.记录性质,A.记录状态,A.结帐ID,a.收费细目ID"
    
    strSQL = "" & _
    "   Select /*+ Rule */ A.NO,A.序号,A.收费细目ID," & _
    "      sum(A.原始数量/" & IIf(mtyMoudlePara.bln药房单位, "nvl(B." & gstr药房包装 & ",1)", "1") & ") as 原始数量, " & _
    "      sum(A.数量/" & IIf(mtyMoudlePara.bln药房单位, "nvl(B." & gstr药房包装 & ",1)", "1") & ")  as 剩余数量 " & _
    "   From (" & strSQL & ") A,药品规格 B" & _
    "   Where A.收费细目ID=B.药品ID(+) " & _
    "   Group by A.NO,A.序号,a.收费细目ID" & _
    "   Order by NO,序号"

    On Error GoTo errHandle
    Set GetFeeDelNumRecord = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strAllNOs)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckIsAllDel(ByVal strAllNOs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查所有费用是否全退
    '入参:strAllNos-所有单据,多个用逗号分隔
    '出参:
    '返回:所有全退时,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-15 11:28:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strNo As String, int序号 As Integer
    Dim blnFind As Boolean, dbl剩余数量 As Double
    Dim j As Long, k As Long
    
    On Error GoTo errHandle
    If mbytMode = EM_RBDTY_退费 Then
        With vsBill
            For j = 1 To vsBill.Rows - 1
                If Abs(Val(.TextMatrix(j, .ColIndex("选择")))) <> 1 And InStr(strAllNOs, .TextMatrix(j, .ColIndex("单据号"))) > 0 Then
                   CheckIsAllDel = False: Exit Function
                End If
            Next
        End With
    End If
    Set rsTemp = GetFeeDelNumRecord(strAllNOs)
    With rsTemp
        Do While Not .EOF
            strNo = Nvl(!NO): int序号 = Val(Nvl(!序号))
            dbl剩余数量 = Val(Nvl(!剩余数量))
            If dbl剩余数量 <> 0 Then
                With vsBill
                    k = vsBill.FindRow(strNo, , .ColIndex("单据号"))
                    If k <= 0 Then Exit Function
                    blnFind = False
                    For j = k To vsBill.Rows - 1
                        If .TextMatrix(j, .ColIndex("单据号")) <> strNo Then Exit For
                        If Abs(Val(.TextMatrix(j, .ColIndex("选择")))) <> 1 _
                            And mbytMode <> EM_RBDTY_异常重退 Then
                            CheckIsAllDel = False: Exit Function
                        End If
                        If Val(.RowData(j)) = int序号 Then
                            If Val(dbl剩余数量) <> Val(.Cell(flexcpData, j, .ColIndex("数量"))) Then
                               CheckIsAllDel = False: Exit Function
                            End If
                            blnFind = True: Exit For
                        End If
                    Next
                End With
                If blnFind = False Then Exit Function
            End If
            .MoveNext
        Loop
    End With
    CheckIsAllDel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExecuteReDelFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对异常单据重新退费
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-17 15:43:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmBalance As frmReplenishTheBalanceDelWin, objDelBalance As clsCliniDelBalance
    
    Dim bln全退 As Boolean, str结帐ID As String, lng结帐ID As Long, lng冲销ID As Long
    Dim strNos As String, varData As Variant, strCmdCaptions As String
    Dim cllPro  As New Collection, strReclaimInvoice As String, strInvoice As String
    Dim lngCheck病人ID As Long, intCheckInsure   As Integer, strYBPati As String
    Dim dtDelDate As Date, blnTrans As Boolean, strNo As String
    Dim str序号 As String, j As Long, strPrintNOInfor As String
    Dim strSQL As String, strBalanceInfor As String, cur个帐透支 As Currency
    Dim strReturn As String, strReturnRecipt As String '退费处方信息，格式：NO,药房ID|NO,药房ID|…
    Dim rs药品记录 As ADODB.Recordset, lng领用ID As Long
    Dim rsBalance As ADODB.Recordset
    
    On Error GoTo errHandle
    '并发检查
    If zlIsCheckExistErrBill(Val(mstr结算序号), True) = False Then
        MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(Val(mstr结算序号)) Then
        MsgBox "当前单据正在其它补结算窗口中进行处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    '三方卡结算方式有效性检查
    Set rsBalance = zlFromIDGetChargeBalance(2, mCurBillType.strAllNOs, , , True, IIf(mCurBillType.bln挂号, 4, 1))
    If ThreeBalanceCheck(mobjPayCards, rsBalance, mcllForceDelToCash, mstr排除结算方式) = False Then Exit Function
    
    If ShowReclaimInvoice(mCurBillType.str结算单, strReclaimInvoice) = False Then Exit Function
    bln全退 = CheckIsAllDel(mCurBillType.strAllNOs)
    If Not bln全退 Then
        If MCPAR.医保接口打印票据 Then
            If zlGetInvoiceGroupUseID(lng领用ID) = False Then Exit Function
            strInvoice = GetNextBill(lng领用ID)
        End If
    End If
    With vsBill
        str序号 = "": strNo = ""
        For j = 1 To vsBill.Rows - 1
            If strNo <> Trim(.TextMatrix(j, .ColIndex("单据号"))) Then
                If str序号 <> "" Then
                    strPrintNOInfor = strPrintNOInfor & ";" & strNo & ":" & Mid(str序号, 2)
                End If
                strNo = .TextMatrix(j, .ColIndex("单据号"))
                str序号 = ""
            End If
            str序号 = str序号 & "," & CLng(vsBill.RowData(j))
        Next
    End With
    
    Set objDelBalance = New clsCliniDelBalance
    'bytType-查找类型:0-根据结帐ID查找;1-根据结算序号查找,2-根据单据号来获取结算方式
    Set objDelBalance.rsBalance = zlFromIDGetChargeBalance(1, Val(mstr结算序号), False)
    Set objDelBalance.rs结算方式 = mrs结算方式
    
    lng结帐ID = mCurBillType.lng结帐ID
    lng冲销ID = mCurBillType.lng冲销ID
    dtDelDate = zlDatabase.Currentdate
    
    '先退医保
    If mCurBillType.intInsure <> 0 And lng结帐ID <> 0 Then
        '如果是医保,出现异常,肯定是只有重收部分才出现异常
        '字段:类型 ,结帐ID, 记录性质, 结算方式, 摘要, 卡类别ID, 卡类别名称, 自制卡, 结算卡序号, 结算号码, 卡号, 交易流水号, 交易说明, 结算序号, 校对标志, 医保, 消费卡id
        '            是否密文,是否全退,是否退现,冲预交
        '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        mrsBalance.Filter = "结帐ID=" & lng结帐ID & " And 类型=2 "
        If mrsBalance.EOF Then
            '未进行医保预结算,因此,需要重新预结,然后结算
            '可能存在重新收费,因此,需要调用身份验证接口(Identifiy)
            'strAdvace:医保部分退时:传入1,表示医保部分退后再重新收费的身份验证;其他传入: 空
            lngCheck病人ID = mCurBillType.lng病人ID
            intCheckInsure = mCurBillType.intInsure
            strYBPati = gclsInsure.Identify(0, lngCheck病人ID, intCheckInsure, 1)
            If strYBPati = "" Then
                 MsgBox "医保身份验证失败,不允许继续退费!", vbOKOnly + vbDefaultButton1 + vbExclamation, gstrSysName
                 Exit Function
            End If
             
            If Val(CLng(Split(strYBPati, ";")(8))) <> mCurBillType.lng病人ID Then
                MsgBox "医保验证的病人与退费的病人不是同一个病人!", vbInformation, gstrSysName
                Call ExecuteYBIdentifyCancel(mCurBillType.lng病人ID, mCurBillType.intInsure)
                Exit Function
            End If
            
            If GetExcutInsureInforUpdateSQL(Val(mstr结算序号), strBalanceInfor, cllPro) = False Then Exit Function
            '读取个帐余额
            cur个帐透支 = mTy_Insure.dbl个帐透支
            mTy_Insure.dbl帐户余额 = gclsInsure.SelfBalance(mCurBillType.lng病人ID, CStr(Split(strYBPati, ";")(1)), 10, cur个帐透支, mCurBillType.intInsure)
            mTy_Insure.dbl个帐透支 = cur个帐透支
            
            
            blnTrans = True
            zlExecuteProcedureArrAy cllPro, Me.Caption, True
            
            '重结调用医保接口
            '77058
            If ExcuteInsureReCharge(mCurBillType.lng病人ID, mCurBillType.intInsure, lng结帐ID, Val(mstr结算序号), strBalanceInfor, _
                        mCurBillType.str结算单, lng领用ID, strInvoice, dtDelDate) = False Then Exit Function
            blnTrans = False
        End If
    End If
    
    '4.显示结算界面
    mCurBillType.lng结算序号 = Val(mstr结算序号) '记录用于打印红票
    With objDelBalance
        .intInsure = mCurBillType.intInsure
        .CurDelNos = mCurBillType.strNos
        .AllNos = mCurBillType.strAllNOs
        .PrintNOs = "'" & mCurBillType.str结算单 & "'"
        .PatiUseType = mobjFact.使用类别
        .SaveBilled = True
        .ShareUserID = mobjFact.共享批次ID
        .病人ID = mCurBillType.lng病人ID
        .冲销ID = lng冲销ID
        .当前发票号 = strInvoice
        .回收发票 = strReclaimInvoice
        .结算序号 = Val(mstr结算序号)
        .结帐ID = lng结帐ID
        .缺省结算方式 = mCurBillType.str结算方式
        .退费合计 = -1 * GetDelMoney
        .费别 = mCurBillType.str费别
        .年龄 = mCurBillType.str年龄
        .性别 = mCurBillType.str性别
        .姓名 = mCurBillType.str姓名
        .病人类型 = mCurBillType.str病人类型
        .医保不走票号 = MCPAR.医保不走票号
        .原结帐ID = mCurBillType.lng原结帐ID
        .退费时间 = dtDelDate
        .部分退费 = Not bln全退
    End With
    
    Set frmBalance = New frmReplenishTheBalanceDelWin
    If frmBalance.zlChargeWin(Me, mlngModule, mstrPrivs, EM_BalanceReDel, mobjPayCards, objDelBalance, MCPAR.分币处理, _
        mcllForceDelToCash, mstr排除结算方式, mCurBillType.bln挂号) = False Then Exit Function

    '81190,冉俊明,退费业务向发药机上传退费信息
    On Error Resume Next
    If Not mCurBillType.bln挂号 Then
        If mblnDrugMachine Then
            Dim rsTemp As ADODB.Recordset, strData As String '门诊处方退药格式：费用ID1,退药数量1;费用ID2,退药数量2;...
            '本次退的减去重收的就是实际退的
            strSQL = "Select Max(Decode(a.记录状态, 2, a.Id, 0)) As 费用id, -1 * Nvl(Sum(a.付数 * a.数次), 0) As 退药数量" & vbNewLine & _
                    " From 门诊费用记录 A,(Select Distinct 结帐ID From 病人预交记录 Where 结算序号 = [1]) B" & vbNewLine & _
                    " Where a.结帐id = b.结帐ID And Mod(a.记录性质, 10) = 1 And a.收费类别 In ('5', '6', '7')" & vbNewLine & _
                    " Group By NO, Nvl(价格父号, 序号)" & vbNewLine & _
                    " Having Nvl(Sum(a.付数 * a.数次), 0) <> 0"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询本次退费项目", objDelBalance.结算序号)
            Do While Not rsTemp.EOF
                strData = strData & ";" & Nvl(rsTemp!费用id) & "," & Nvl(rsTemp!退药数量)
                rsTemp.MoveNext
            Loop
            If strData <> "" Then
                strData = Mid(strData, 2)
                Call mobjDrugMachine.Operation(gstrDBUser, Val("24-处方退药(完整/部分)"), strData, strReturn)
            End If
        ElseIf mblnDrugPacker Then
            strSQL = "Select a.No, a.执行部门id" & _
                "   From 门诊费用记录 A, 病人预交记录 B" & _
                "   Where a.结帐id = b.结帐id And a.记录状态=2 And a.收费类别 In ('5', '6', '7') And b.结算序号 = [1]"
            Set rs药品记录 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstr结算序号))
            
            If rs药品记录.RecordCount <> 0 Then
                Do While Not rs药品记录.EOF
                    If InStr(strReturnRecipt & "|", "|" & Nvl(rs药品记录!NO) & "," & Nvl(rs药品记录!执行部门ID) & "|") = 0 Then
                        strReturnRecipt = strReturnRecipt & "|" & Nvl(rs药品记录!NO) & "," & Nvl(rs药品记录!执行部门ID)
                    End If
                    rs药品记录.MoveNext
                Loop
            End If
    
            If strReturnRecipt <> "" Then
                strReturnRecipt = Mid(strReturnRecipt, 2)
                Call mobjDrugPacker.DYEY_MZ_TransRecipeReturn(1, UserInfo.编号, UserInfo.姓名, strReturnRecipt, strReturn)
            End If
        End If
    End If
    Err.Clear: On Error GoTo errHandle
    

    
    If Not gfrmMain Is Nothing Then
        Call zlExeBalanceWinRefrshData(True, objDelBalance)
    End If
    ExecuteReDelFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckIsExistDelErrBill(ByVal strNos As String, Optional ByRef strOperatorName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据号,检查是否存在退费异常记录
    '入参:strNOs=单据号,格式 NO1,NO2,NO3,...
    '出参:
    '     strOperatorName=产生退费异常单据的操作员姓名
    '返回:存在退费异常单据,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-11 12:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    strOperatorName = ""
    If strNos = "" Then Exit Function
    
    On Error GoTo Errhand
    strSQL = "" & _
            " Select 操作员姓名" & _
            " From 费用补充记录 A" & _
            " Where Nvl(费用状态, 0) = 1 And 记录性质 = 1 And 记录状态 = 2 " & _
            "       And a.No In (Select Column_Value From Table(f_Str2list([1])))" & _
            "       And Not Exists (Select 1 From 病人预交记录 B Where a.结算id = b.结帐id And Nvl(b.校对标志, 0) = 0)" & _
            "       And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否存在退费异常记录", strNos)
    
    If Not rsTemp.EOF Then
        strOperatorName = Nvl(rsTemp!操作员姓名)
        CheckIsExistDelErrBill = True
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function
Private Function GetExcutInsureInforUpdateSQL(ByVal lng结算序号 As Long, _
    ByRef strBalanceInfor As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取需要更新的相关SQL
    '出参:strBalanceInfor:项目结算信息(通过GetItemInsure返回),格式: 实收合计;进入统筹;全自付;先自付
    '     cllPro-返回需要执行的SQL
    '返回:获取成功,返回true,否则返回False
    '编制:冉俊明
    '日期:2014-9-16
    '问题:77951
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, strBXInfo As String
    Dim blnTrans As Boolean, cur实收合计 As Currency, cur进入统筹 As Currency, cur全自付 As Currency, cur先自付 As Currency
    Dim cur实收金额 As Currency, cur统筹金额 As Currency, bln保险项目 As Boolean

    
    On Error GoTo Errhand
    strSQL = "" & _
    "   Select max(decode(a.记录状态,2,0,a.Id)) as ID," & _
    "          A.NO,a.病人id, a.收费细目id,a.序号,a.收入项目ID,sum(nvl(a.付数,1)*a.数次) as 数量," & _
    "          Nvl(sum(a.实收金额), 0) As 实收金额,max(decode(a.记录状态,2,'',a.摘要)) as 摘要 " & _
    "   From 门诊费用记录 A,(Select distinct 收费结帐ID From 费用补充记录 Where 结算序号=[1] ) B" & _
    "   Where a.结帐id =B.收费结帐id " & _
    "   Group by  A.NO,a.病人id,a.收费细目id,a.序号,a.收入项目ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取重收费用记录", lng结算序号)
    
    With rsTemp
        If .RecordCount > 0 Then
            Set cllPro = New Collection
            Do While Not .EOF
                '保险项目否(0/1);保险大类ID;进入统筹金额;保险项目编码;摘要;费用类型
                strBXInfo = gclsInsure.GetItemInsure(Nvl(!病人ID), Nvl(!收费细目ID), Val(Nvl(!实收金额)), True, mCurBillType.intInsure, _
                        Nvl(!摘要) & "||" & Val(Nvl(!数量)))
                If strBXInfo <> "" Then
                    '  Zl_门诊收费记录_Update
                    strSQL = "Zl_门诊收费记录_Update("
                    '  Id_In         In 门诊费用记录.Id%Type,
                    strSQL = strSQL & Nvl(!ID) & ","
                    '  保险大类id_In In 门诊费用记录.保险大类id%Type,
                    strSQL = strSQL & ZVal(Split(strBXInfo, ";")(1)) & ","
                    '  保险项目否_In In 门诊费用记录.保险项目否%Type,
                    strSQL = strSQL & Val(Split(strBXInfo, ";")(0)) & ","
                    '  保险编码_In   In 门诊费用记录.保险编码%Type,
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(3)) & "',"
                    '  费用类型_In   In 门诊费用记录.费用类型%Type,
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(5)) & "',"
                    '  统筹金额_In   In 门诊费用记录.统筹金额%Type,
                    strSQL = strSQL & Format(Val(Split(strBXInfo, ";")(2)), gstrDec) & ","
                    '  摘要_In       In 门诊费用记录.摘要%Type
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(4)) & "')"
                    zlAddArray cllPro, strSQL
                    
                    cur统筹金额 = CCur(Val(Split(strBXInfo, ";")(2)))
                    bln保险项目 = Val(Split(strBXInfo, ";")(0)) = 1
                Else
                    cur统筹金额 = Val(Nvl(!统筹金额))
                    bln保险项目 = Val(Nvl(!保险项目否)) = 1
                End If
                
                '统计保险金额
                cur实收金额 = Val(Nvl(!实收金额))
                If cur统筹金额 = 0 Or Not bln保险项目 Then
                    '以原始金额为准,不管分币处理
                    cur全自付 = cur全自付 + cur实收金额
                Else
                    cur进入统筹 = cur进入统筹 + cur统筹金额
                    '以原始金额为准,不管分币处理
                    cur先自付 = cur先自付 + cur实收金额 - cur统筹金额
                End If
                cur实收合计 = cur实收合计 + CCur(Val(Nvl(!实收金额)))
                rsTemp.MoveNext
            Loop
        End If
    End With
    '保险金额信息
    strBalanceInfor = cur实收合计 & ";" & cur进入统筹 & ";" & cur全自付 & ";" & cur先自付
    GetExcutInsureInforUpdateSQL = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetChargeBalance(ByVal bytType As Byte, _
    ByVal strValue As String, Optional blnHistory As Boolean, _
    Optional ByRef blnDel As Boolean, Optional ByVal bln含异常 As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算ID获取收费结算信息
    '入参:bytType-查找类型:0-根据结帐ID查找;1-根据结算序号查找,2-根据单据号来获取结算方式
    '     strValue-要查找的值(为0时,结帐ID,为1时,结算序号,2时为一次收费所涉及的所有单据)
    '     blnDel-退费结算:true-查退费结算;false-非退费结算
    '     bln含异常-是否包含异常结算，根据单据号来获取结算数据时有效
    '返回:收费结算的相关信息集
    '       字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '            是否密文,是否全退,是否退现,冲预交
    '       类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    '编制:刘兴洪
    '日期:2014-06-24 16:37:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTable As String, strWhere As String
    Dim strTable1 As String
    On Error GoTo errHandle
    
    strTable = IIf(blnHistory, "H", "") & "病人预交记录"
    Select Case bytType
    Case 0  '0-根据结帐ID查找
        strWhere = " And  A.结帐ID= [1]"
    Case 1  ';1-根据结算序号查找
        strWhere = "  And A.结算序号= [1]"
    Case 2 '根据单据号来获取结算数据
        strTable1 = "" & _
        "   Select distinct 收费结帐ID as 结帐ID " & _
        "   From 费用补充记录 M " & _
        "   Where M.NO= [2] And Mod(M.记录性质,10)=1" & IIf(bln含异常, "", " And Nvl(M.费用状态,0)<>1")
        strTable1 = strTable1 & " union ALL" & Replace(strTable1, "收费结帐ID", "结算ID")
        strTable1 = ",(" & strTable1 & ") Q1"
        If blnHistory Then strTable1 = Replace(strTable1, "费用补充记录", "H费用补充记录")
        strWhere = " And A.结帐ID=Q1.结帐ID"
    End Select
    
    If blnDel Then
        '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老))(无);5-消费卡(无)
        strSQL = "" & _
        "   Select  A.ID,decode(A.记录状态,2,A.结帐ID,NULL) as 结帐ID," & _
        "        Case when Mod(A.记录性质,10)=1 then 1  " & _
        "             when B.名称 is not null then  2 " & _
        "             when nvl(A.卡类别ID,0)<>0  then  3 " & _
        "             when J.结算方式 is not null   then  4 " & _
        "             else 0 end as 类型, " & _
        "        Mod(A.记录性质,10) as 记录性质,A.结算方式,A.冲预交," & _
        "        decode(A.记录状态,2,A.摘要,NULL) as 摘要,decode(A.记录状态,2,1,0) as 退费," & _
        "        A.卡类别ID,A.结算卡序号, " & _
        "        decode(A.记录状态,2,A.结算号码,NULL) as 结算号码,decode(A.记录状态,2,A.卡号,NULL) as 卡号, " & _
        "        decode(A.记录状态,2,A.交易流水号,NULL) as 交易流水号,nvl(C.是否自制,0) as 自制卡, " & _
        "        nvl(C.是否退现,0) as 是否退现,nvl(C.是否全退,0) as 是否全退, " & _
        "        Decode(C.卡号密文,NULL,0,1) as  是否密文,Nvl(C.是否转帐及代扣,0) as 是否转帐及代扣," & _
        "        C.名称 as 卡类别名称,decode(A.记录状态,2,A.交易说明,NULL) as 交易说明,A.结算序号,decode(A.记录状态,2,A.校对标志,0) as 校对标志, " & _
        "        decode(B.名称,Null,0,1) as 医保,0 as 消费卡id,nvl(q.性质,1) as 结算性质" & _
        "   From " & strTable & " A ,医疗卡类别 C,一卡通目录 J,结算方式 q," & _
        "        (Select 名称 From 结算方式 where 性质 in (3,4)) B " & strTable1 & _
        "   Where A.结算方式=J.结算方式(+) And A.卡类别ID=C.ID(+) " & _
        "         And A.结算方式=B.名称(+) and A.结算方式=q.名称(+) " & _
        "         And (a.记录性质 In (1, 11) Or Nvl(a.结算卡序号, 0) = 0) " & strWhere
     
        strSQL = "" & _
        "   Select  max(结帐id) as 结帐id,类型,max(退费) as 退费,记录性质,结算方式,Max(摘要) as 摘要,卡类别ID,卡类别名称,max(自制卡) as 自制卡,结算卡序号, " & _
        "         max(结算号码) as 结算号码,max(卡号) as 卡号,max(交易流水号) as 交易流水号, max(交易说明) as 交易说明, " & _
        "         结算序号,max(校对标志) as 校对标志,医保,消费卡id,结算性质,max(是否转帐及代扣) as 是否转帐及代扣," & _
        "         max(是否密文) as 是否密文,max(是否全退) as 是否全退,max(是否退现) as 是否退现 , nvl(sum(冲预交),0) as 冲预交" & _
        "   From (" & strSQL & ") " & _
        "   Group by 类型, 记录性质,结算方式,卡类别ID,卡类别名称,结算卡序号,结算序号,医保,消费卡id,结算性质 having  sum(冲预交) <>0"
        Set GetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "获取收费结算方式", Val(strValue), strValue)
        Exit Function
    End If
    
    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老)(无);5-消费卡(无)
    strSQL = "" & _
    "   Select  A.ID,A.结帐ID," & _
    "        Case when Mod(A.记录性质,10)=1 then 1  " & _
    "             when B.名称 is not null then  2 " & _
    "             when nvl(A.卡类别ID,0)<>0  then  3 " & _
    "             when J.结算方式 is not null   then  4 " & _
    "             else 0 end as 类型, " & _
    "        Mod(A.记录性质,10) as 记录性质,A.结算方式,A.冲预交," & _
    "        A.摘要,decode(A.记录状态,2,1,0) as 退费," & _
    "        A.卡类别ID,A.结算卡序号, " & _
    "        A.结算号码,A.卡号,A.交易流水号,nvl(C.是否自制,0) as 自制卡, " & _
    "        nvl(C.是否退现,0) as 是否退现,nvl(C.是否全退,0) as 是否全退, " & _
    "        Decode(C.卡号密文,NULL,0,1) as  是否密文,Nvl(C.是否转帐及代扣,0) as 是否转帐及代扣," & _
    "        C.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志, " & _
    "        decode(B.名称,Null,0,1) as 医保,0 as 消费卡id,nvl(q.性质,1) as 结算性质" & _
    "   From " & strTable & " A ,医疗卡类别 C,一卡通目录 J,结算方式 q," & _
    "        (Select 名称 From 结算方式 where 性质 in (3,4)) B " & strTable1 & _
    "   Where A.结算方式=J.结算方式(+) And A.卡类别ID=C.ID(+) " & _
    "         And A.结算方式=B.名称(+) and A.结算方式=q.名称(+) " & _
    "         And (a.记录性质 In (1, 11) Or Nvl(a.结算卡序号, 0) = 0) " & strWhere
    
    gstrSQL = "" & _
    "   Select  结帐ID,类型,退费,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id,结算性质," & _
    "         max(是否转帐及代扣) as 是否转帐及代扣,max(是否密文) as 是否密文,max(是否全退) as 是否全退,max(是否退现) as 是否退现 , nvl(sum(冲预交),0) as 冲预交" & _
    "   From (" & gstrSQL & ") " & _
    "   Group by 结帐ID,类型,退费,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id,结算性质"
    Set GetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "获取收费结算方式", Val(strValue), strValue)
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetBalanceNO(ByVal intFindType As Integer, _
    ByVal strFindValue As String, _
    ByRef strNo As String, Optional bln挂号补充 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算序号获取结算单据号
    '入参:intFindType:查找类型(0-根据结算序号查找;1-根据收费单号查找;2-根据发票号来查找;3-根据挂号单号来查找;4-根据结算单号查找)
    '      strFindValue-intFindType=0:结算序号;intFindType=1:收费单号;intFindType=2:发票号;intFindType=3:挂号单
    '出参:strNo-返回结算单号
    '     bln挂号补充-是否挂号补充结算
    '返回:获取成功,返回true,读取失败或未找到结算数据,返回False
    '编制:刘兴洪
    '日期:2014-09-29 10:06:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim intMouse As Integer
    
    On Error GoTo errHandle
    
    If strFindValue = "" Then strFindValue = "0"
    
    Select Case intFindType
    Case 0 '根据结算序号查找
        strSQL = "Select NO,附加标志 From 费用补充记录 Where 结算序号=[1] and rownum <2"
    Case 1 '根据收费单号查找
        strSQL = "" & _
        "   Select    A.NO,A.附加标志 From 费用补充记录 A,门诊费用记录 B " & _
        "   Where A.收费结帐ID=B.结帐ID And B.NO=[1] and mod(B.记录性质,10)=1 and rownum <2"
    Case 2 '根据发票号来查找
        strSQL = "" & _
        "   Select  C.NO,C.附加标志 From 票据使用明细 A, 票据打印内容 B,费用补充记录 C " & _
        "   Where 号码 = [1] and B.NO=C.NO and mod(C.记录性质,10)=1 And b.数据性质 = 1 And A.打印id = b.Id and rownum<2"
    Case 3  '根据挂号单号来查找
        strSQL = "" & _
        "   Select    A.NO,A.附加标志 From 费用补充记录 A,门诊费用记录 B " & _
        "   Where A.收费结帐ID=B.结帐ID And B.NO=[1] and mod(B.记录性质,10)=4 and rownum <2"
    Case 4  '4-根据结算单号查找
        strSQL = "Select A.NO,A.附加标志 From 费用补充记录 A Where A.NO=[1] and rownum <2"
    Case Else
        Exit Function
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFindValue)
    If Not rsTemp.EOF Then strNo = Nvl(rsTemp!NO): bln挂号补充 = Val(Nvl(rsTemp!附加标志)) = 1
    GetBalanceNO = True
    Exit Function
errHandle:
    
    intMouse = Me.MousePointer: Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Me.MousePointer = intMouse
        Resume
    End If
    Me.MousePointer = intMouse
End Function

Private Function GetChargeInsure(ByVal str结算ID As String, ByVal strNo As String, _
    ByRef lng病人ID As Long, Optional ByVal blnNOMoved As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费的医保号
    '入参:lng结帐ID-结帐ID
    '     blnNOMoved-是否数据转移
    '出参:lng病人ID-病人ID
    '返回:险类
    '编制:刘兴洪
    '日期:2014-07-02 14:30:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strWhere  As String
    
    On Error GoTo errHandle
    
    lng病人ID = 0
    strWhere = " And A.结算ID=[1]"
    If str结算ID = "" Or str结算ID = "0" Then strWhere = " And A.NO=[2]"
    If str结算ID = "" Then str结算ID = "0"
    
    strSQL = "" & _
    "    Select B.记录ID,B.险类,B.病人ID " & _
    "    From 费用补充记录 A,保险结算记录 B " & _
    "    Where A.结算ID=[1] And  mod(A.记录性质,10)=1 " & _
    "         And B.性质=1 And A.结算ID=B.记录ID and Rownum<2 "
    If blnNOMoved Then
        strSQL = Replace(strSQL, "费用补充记录", "H费用补充记录")
        strSQL = Replace(strSQL, "保险结算记录", "H保险结算记录")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据算ID获取指定的医保险类", str结算ID, strNo)
    If rsTemp.EOF Then Exit Function
    lng病人ID = Nvl(rsTemp!病人ID, 0)
    GetChargeInsure = Nvl(rsTemp!险类, 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function IsRegisterBalance(ByVal strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断当前结算单据是否挂号单据的补充结算
    '入参:strNO-结算单号
    '出参:
    '返回:挂号单补充结算,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-08 16:19:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select A.附加标志 From 费用补充记录 A " & _
    "   where A.NO=[1] And mod(A.记录性质,10)=1   and rownum <2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    If rsTemp.EOF Then Exit Function
    IsRegisterBalance = Val(Nvl(rsTemp!附加标志)) = 1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBalanceFeeNos(ByVal bytType As Byte, _
    ByVal strFindValue As String, _
    Optional ByRef strFeeNos As String, Optional ByRef strRegNos As String, _
    Optional ByVal blnNOMoved As Boolean) As Boolean
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据一张结算单据的NO或结帐ID或结帐序号，返回同一次结算的收费单据NOs
    '入参:bytType-0-根据NO来查找;1-根据结帐ID来查找,2-根据结算序号来查找
    '    strFindValue-查找的值
    '    blnNOMoved-是否在后备表中，查询单据之前的判断需要用这个参数
    '出参:strFeeNos-返回当前结算的费用单号,格式如"AAA,BBB,CCC',..."
    '     strRegNos-返回当前结算的挂号单号,格式如"AAA,BBB,CCC',..."
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNo As String
    On Error GoTo errHandle:
    Select Case bytType
    Case 0 '0-根据结算NO来查找
        strSQL = "" & _
        "   Select distinct A.NO,mod(A.记录性质,10) as 记录性质" & _
        "   From 门诊费用记录 A," & _
        "        (Select distinct 收费结帐ID as 结帐ID From 费用补充记录 Where NO=[1] and 记录性质=1 ) B" & _
        "   Where A.结帐ID=B.结帐ID" & _
        "   Order by 记录性质,NO"
        
    Case 1  '1-根据结帐ID来查找
        strSQL = "" & _
        "    Select Distinct A.No,mod(A.记录性质,10) as 记录性质 " & _
        "    From 门诊费用记录 A," & _
        "        (Select distinct C1.收费结帐ID as 结帐ID " & _
        "         From 费用补充记录 A1,费用补充记录 B1,费用补充记录 C1  " & _
        "         Where A1.结算ID=[2] and A1.记录性质=1  " & _
        "               And A1.NO=B1.NO and A1.记录性质=B1.记录性质 " & _
        "               And B1.结算序号=C1.结算序号 and C1.记录状态 in (1,3) ) B " & _
        "    Where A.结帐ID=B.结帐ID" & _
        "    Order By 记录性质,NO"
    Case 2  '2-根据结算序号来查找
        strSQL = "" & _
        "    Select Distinct A.No,mod(A.记录性质,10) as 记录性质" & _
        "    From 门诊费用记录 A," & _
        "        (Select distinct C1.收费结帐ID as 结帐ID " & _
        "         From 费用补充记录 A1,费用补充记录 B1,费用补充记录 C1  " & _
        "         Where A1.结算序号=[2] and A1.记录性质=1  " & _
        "               And A1.NO=B1.NO and A1.记录性质=B1.记录性质 " & _
        "               And B1.结算序号=C1.结算序号 and C1.记录状态 in (1,3) ) B " & _
        "    Where A.结帐ID=B.结帐ID" & _
        "    Order By 记录性质,NO"
    End Select
    
    If blnNOMoved Then
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        strSQL = Replace(strSQL, "费用补充记录", "H费用补充记录")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取一次结算所涉及的费用单据", strFindValue, Val(strFindValue))
    
    With rsTemp
        strFeeNos = "": strRegNos = ""
        Do While Not .EOF
            strNo = Nvl(!NO)
            If Val(Nvl(!记录性质)) = 1 Then
                If InStr(1, strFeeNos & ",", "," & strNo & ",") = 0 Then
                    strFeeNos = strFeeNos & "," & strNo
                End If
            Else
                If InStr(1, strRegNos & ",", "," & strNo & ",") = 0 Then
                    strRegNos = strRegNos & "," & strNo
                End If
            End If
            .MoveNext
        Loop
    End With
    If strFeeNos <> "" Then strFeeNos = Mid(strFeeNos, 2)
    If strRegNos <> "" Then strRegNos = Mid(strRegNos, 2)
    If strFeeNos = "" And strRegNos = "" Then Exit Function
    GetBalanceFeeNos = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetFeeListData(ByVal strNos As String, ByRef rsFeeList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前准退的费用项目
    '入参:strNos-准退单号
    '出参:rsFeeList-返回准退费集
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-08 17:49:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTableNo As String, strSQLIn As String
    Dim strSQL As String, strSqlSub As String
    
    strSqlSub = "" & _
        " Select /*+cardinality(j,10)*/ A.ID,A.记录性质,A.NO,A.记录状态,A.序号,A.从属父号,A.价格父号,A.收费细目ID, " & _
        "        nvl(A.付数,1) as 付数, nvl(A.数次,0) as 数次, " & _
        "        nvl(A.应收金额,0) as 应收金额 ,nvl(A.实收金额,0) as 实收金额,nvl(A.结帐金额,0) as 结帐金额," & _
        "        Nvl(A.付数,1)*A.数次 as 数量, nvl(标准单价,0)  as 标准单价," & _
                 IIf(mtyMoudlePara.bln药房单位, "nvl(B." & gstr药房包装 & ",1)", "1") & " as 换算系数, " & _
                 IIf(mtyMoudlePara.bln药房单位, " decode(B.药品ID,NULL,A.计算单位,B." & gstr药房单位 & ")", "A.计算单位 ") & " as 计算单位," & _
        "        A.开单部门ID,A.执行部门ID,A.医嘱序号, " & _
        "        A.执行状态,A.费用类型,A.费用状态 ,A.附加标志,A.费别,A.收费类别,A.操作员姓名,A.登记时间,A.结帐ID," & _
        "        B.药品ID" & _
        " From 门诊费用记录 A,药品规格 B,Table(f_Str2list([1])) J  " & _
        " Where mod(A.记录性质,10)=1 And A.NO=J.Column_Value and A.记录状态<>0" & _
        "       And A.收费细目ID=B.药品ID(+)"
    '求准退费(卫材,药品,其他治疗类)
    strTableNo = _
        " With 门诊费用 as (" & strSqlSub & ")," & vbNewLine & _
        "      准退数 as (Select /*+cardinality(j,10)*/ A.费用ID," & _
        "                        Sum(Nvl(A.付数,1)*A.实际数量" & IIf(mtyMoudlePara.bln药房单位, "/Nvl(B." & gstr药房包装 & ",1)", "") & ") as 准退数量" & _
        "                 From 药品收发记录 A,药品规格 B, Table(f_Str2list([1])) J" & _
        "                 Where A.药品ID=B.药品ID(+) And Mod(A.记录状态,3)=1  " & _
        "                       And (A.单据 =8 or a.单据=24) And A.审核人 is NULL And A.NO =J.Column_Value" & _
        "                 Group by A.费用ID"

    '求诊疗相关的准退数
    '*在医嘱执行计价中存在数据时,则按医嘱执行计价中取数
    '*病人医嘱发送.执行状态=1（完成执行）时，准退数为0，不再根据医嘱执行计价来统计准退数,112447
    strTableNo = strTableNo & vbNewLine & _
        "   Union ALL " & vbNewLine & _
        "   Select Max(ID) As 费用ID, Nvl(Sum(数量), 0) As 准退数" & vbNewLine & _
        "   From(Select a.Id, a.医嘱序号 As 医嘱id, a.收费细目id, Decode(b.执行状态, 1, 0, Decode(c.执行状态, 0, 1, 0)) * c.数量 As 数量" & vbNewLine & _
        "        From (" & strSqlSub & ") A, 病人医嘱发送 B, 医嘱执行计价 C, 病人医嘱记录 M" & vbNewLine & _
        "        Where a.医嘱序号 = b.医嘱id And a.No = b.No And b.医嘱id = c.医嘱id And b.医嘱ID = m.id" & vbNewLine & _
        "              And b.发送号 = c.发送号 And a.收费细目id = c.收费细目id + 0 And a.价格父号 Is Null" & vbNewLine & _
        "              And a.记录性质 = 1 And a.记录状态 in (1, 3) And Instr(',5,6,7,', ',' || a.收费类别 || ',') = 0" & vbNewLine & _
        "              And Not Exists(Select 1 From 材料特性 C Where a.收费细目id = c.材料id And c.跟踪在用 = 1)" & vbNewLine & _
        "              And Instr(',C,D,F,G,K,',','||m.诊疗类别||',')=0 And b.记录性质 = 1" & vbNewLine & _
        "        )" & vbNewLine & _
        "   Group By 医嘱ID, 收费细目ID" & vbNewLine & _
        "   Having Max(ID) <> 0" & vbNewLine & _
        "  )"
    
    '整张单据汇总结果(明细到收费细目)
    '执行状态应该在原始记录上判断(部分退药且部份退费的记录)
    '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
    '   *无医嘱执行计价的部分退费无法判断准退数量，不允许退费
    strSQLIn = "" & _
        " Select NO, Nvl(价格父号, 序号) As 序号" & vbNewLine & _
        " From 门诊费用" & vbNewLine & _
        " Where 记录性质 = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1" & vbNewLine & _
        " Minus" & vbNewLine & _
        " Select NO, Nvl(价格父号, 序号) As 序号" & vbNewLine & _
        " From 门诊费用 A1" & vbNewLine & _
        " Where A1.记录性质 = 1 And A1.记录状态 In (1, 3) And Nvl(A1.执行状态, 0) = 2" & vbNewLine & _
        "       And Not Exists(Select 1" & vbNewLine & _
        "                      From 病人医嘱发送 B, 医嘱执行计价 C" & vbNewLine & _
        "                      Where b.医嘱id = A1.医嘱序号 And b.No = A1.No" & vbNewLine & _
        "                            And b.医嘱id = c.医嘱id And b.发送号 = c.发送号" & vbNewLine & _
        "                            And c.收费细目id + 0 = A1.收费细目id And b.记录性质 = 1)" & vbNewLine & _
        "       And Instr('5,6,7', A1.收费类别) = 0" & vbNewLine & _
        "       And Not Exists(Select 1 From 材料特性 Where 材料id = A1.收费细目id And Nvl(跟踪在用, 0) = 1)"
    
    strSQL = _
    " Select A.NO,A.记录状态,A.记录性质,A.执行状态,Nvl(A.价格父号,A.序号) as 序号,A.从属父号," & _
    "       A.费别,C.编码 as 类别码,C.名称 as 类别名,A.收费细目ID,B.编码,B.名称,B.规格,Max(Nvl(A.费用类型,B.费用类型)) 费用类型," & _
    "       A.计算单位,Max(A.医嘱序号) as 医嘱序号, " & _
    "       Avg(Nvl(A.付数,1)) as 付数,Avg(A.数次/A.换算系数) as 数次," & _
    "       Sum(A.标准单价*A.换算系数) as 单价," & _
    "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, " & _
    "       D.名称 as 执行科室,A.执行部门ID,E.名称 as 开单科室" & _
    " From  门诊费用 A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 E" & _
    " Where A.收费细目ID=B.ID And C.编码=A.收费类别" & _
    "       And A.执行部门ID=D.ID(+) And A.开单部门ID=E.ID(+)" & _
    "       And (A.NO,Nvl(A.价格父号,A.序号)) IN( " & strSQLIn & ")  " & _
    "       And A.NO IN( Select NO From 门诊费用 where  记录性质=1 and 记录状态 in (1,3) )" & _
    " Group by A.NO,A.记录性质,A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号),A.费别,A.从属父号," & _
    "       C.编码,C.名称,A.收费细目ID,B.编码,B.名称,B.规格,A.计算单位," & _
    "       D.名称,A.执行部门ID,E.名称,A.药品ID,a.结帐ID "
     
    '最后计算结果
    '当"准退数量=原始数量"时,付数才保留
    '排开已经全部退费的行(执行状态=0的一种可能)
    '有剩余数量无准退数量的有两种情况：
        '1.无对应的收发记录(如普通费用或不跟踪在用的卫材),这时应用剩余数量
        '2.收发记录中已全部发放,即已全部执行,SQL已排除这种记录
    strSQL = strTableNo & vbCrLf & _
    " Select A.NO,A.序号,A.从属父号,A.费别,A.类别码,A.类别名,A.收费细目ID,A.编码,A.名称,A.规格,Max(A.费用类型) As 费用类型,A.计算单位, Max(A.医嘱序号) as 医嘱序号," & _
    "       Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,avg(A.付数),1) as 准退付数," & _
    "       Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Sum(A.数次),Nvl(C.准退数量,Sum(A.付数*A.数次))) as 准退数次," & _
    "       Nvl(C.准退数量,Sum(A.付数*A.数次)) as 准退数量,Sum(A.付数*A.数次) as 剩余数量," & _
    "       A.单价,Sum(A.应收金额) as 剩余应收,Sum(A.实收金额) as 剩余实收,max(q1.记录标志) as 记录标志," & _
    "       A.执行科室,A.执行部门ID,A.开单科室,B.操作员姓名,B.登记时间,B.结帐ID,Max(M.医嘱内容) as 医嘱内容,b.原始数量" & _
    " From (" & strSQL & ") A, 准退数 C,病人医嘱记录 M," & _
    "          ( Select  ID, NO,序号, 收费细目ID,Nvl( 数量,0)/NVL(换算系数,1) as 原始数量,操作员姓名,登记时间,结帐ID" & _
    "            From 门诊费用   " & _
    "            Where  记录状态 IN(1,3) and 记录性质=1 And Nvl( 附加标志,0)<>9 And  价格父号 is NULL )B, " & _
    "            ( Select NO,Max(记录状态) as 记录标志 From 门诊费用  Where 记录状态 in (1,3) Group by NO) Q1" & _
    " Where A.NO=B.NO And A.序号=B.序号 And A.收费细目ID=B.收费细目ID+0  And B.ID=C.费用ID(+)" & _
    "            and A.医嘱序号=M.ID(+) and A.NO=q1.NO(+) " & _
    " Group by A.NO,A.序号,A.从属父号,A.费别,A.类别码,A.类别名,A.收费细目ID,A.编码,A.名称,A.规格," & _
    "       A.计算单位,A.单价,B.原始数量,C.准退数量,A.执行科室,A.执行部门ID,A.开单科室,B.操作员姓名,B.登记时间,B.结帐ID" & _
    " Having Sum(A.付数*A.数次)<>0"

    strSQL = _
    " Select A.NO,A.序号,A.从属父号,A.费别,A.类别码,A.类别名,A.编码,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名," & _
    "       A.规格,A.费用类型,A.计算单位,A.收费细目ID,A.准退付数 as 付数,A.准退数次 as 数次,A.单价, A.医嘱序号 ," & _
    "       A.剩余应收*(A.准退数量/A.剩余数量) as 应收金额," & _
    "       A.剩余实收*(A.准退数量/A.剩余数量) as 实收金额," & _
    "       A.执行科室,A.执行部门ID,A.开单科室,A.操作员姓名,A.登记时间,A.结帐ID,A.医嘱内容,A.记录标志, " & _
    "       A.原始数量,A.准退数量,A.剩余数量" & _
    " From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1" & _
    " Where A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
    " Order by A.NO,A.序号"
    
    On Error GoTo errHandle
    Set rsFeeList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    GetFeeListData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetRegListData(ByVal strNos As String, ByRef rsFeeList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前准退的挂号项目
    '入参:strNos-准退单号
    '出参:rsFeeList-返回准退费集
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-08 17:49:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTableNo As String, strSQLIn As String
    Dim strSQL As String
    
    strTableNo = "" & _
    "   With  门诊费用  as (" & _
    "           Select A.ID,A.记录性质,A.NO,A.记录状态,A.序号,A.从属父号,A.价格父号,A.收费细目ID, " & _
    "                  nvl(A.付数,1) as 付数, nvl(A.数次,0) as 数次, " & _
    "                  nvl(A.应收金额,0) as 应收金额 ,nvl(A.实收金额,0) as 实收金额,nvl(A.结帐金额,0) as 结帐金额," & _
    "                  Nvl(A.付数,1)*A.数次 as 数量, nvl(标准单价,0)  as 标准单价,1 as 换算系数, A.计算单位 as 计算单位," & _
    "                  A.开单部门ID,A.执行部门ID,A.医嘱序号, " & _
    "                  A.执行状态,A.费用类型,A.费用状态 ,A.附加标志,A.费别,A.收费类别,A.操作员姓名,A.结帐ID, " & _
    "                  A.登记时间,A.发生时间,E.预约时间,E.分诊时间,E.接收时间 as 接收时间,E.诊室,E.执行人 as 医生," & _
    "                  Decode(E.号序, Null, A.发药窗口, To_Char(E.号序)) as  号序,To_Char(E.号别)  as  号码  " & _
    "           From 门诊费用记录 A, 病人挂号记录 E" & _
    "           Where mod(A.记录性质,10)=4 And A.NO IN (Select  Column_Value as No From Table(f_Str2list([1]))) " & _
    "                  And A.记录状态<>0  And A.NO=E.NO and E.记录状态 in (1,3)" & _
    "              )"
    
    strSQL = _
    " Select A.NO,A.记录状态,A.记录性质,A.执行状态,Nvl(A.价格父号,A.序号) as 序号,A.从属父号," & _
    "       A.费别,A.收费类别,A.收费细目ID,A.费用类型," & _
    "       A.计算单位,Max(A.医嘱序号) as 医嘱序号, " & _
    "       Avg(Nvl(A.付数,1)) as 付数,Avg(A.数次/A.换算系数) as 数次," & _
    "       Sum(A.标准单价*A.换算系数) as 单价," & _
    "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, " & _
    "       A.开单部门ID,A.执行部门ID" & _
    " From  门诊费用 A" & _
    " Group by A.NO,A.记录性质,A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号),A.费别,A.从属父号," & _
    "          A.收费类别,A.收费细目ID,A.费用类型,A.计算单位,A.开单部门ID,A.执行部门ID,a.结帐ID "
     
 
    strSQL = strTableNo & vbCrLf & _
    " Select A.NO,A.序号,A.从属父号,A.费别,A.收费类别,A.收费细目ID,Max(A.费用类型) As 费用类型,A.计算单位, Max(A.医嘱序号) as 医嘱序号," & _
    "       sum(a.付数*A.数次) as 准退数次,sum(A.付数*A.数次) as 准退数量,Sum(A.付数*A.数次) as 剩余数量," & _
    "       max(A.单价) as 单价,Sum(A.应收金额) as 剩余应收,Sum(A.实收金额) as 剩余实收," & _
    "       max(decode(A.记录状态,2,0,A.记录状态))  as 记录标志," & _
    "       max(A.开单部门ID) as 开单部门ID,max(A.执行部门ID) as 执行部门ID, " & _
    "       max(B.操作员姓名) as 操作员姓名,max(B.医生) as 医生,max(B.登记时间) as 登记时间,max(B.发生时间) as 发生时间, " & _
    "       max(B.预约时间) as 预约时间,max(B.分诊时间) as 分诊时间,max(B.接收时间) as 接收时间, " & _
    "       max(B.诊室) as 诊室,max(B.号序) as 号序,max(B.号码) as 号码, " & _
    "       max(B.结帐ID) as 结帐ID,max(b.原始数量) as 原始数量" & _
    " From (" & strSQL & ") A," & _
    "      ( Select  ID, NO,序号, 收费细目ID,Nvl( 数量,0)/NVL(换算系数,1) as 原始数量," & _
    "               操作员姓名,医生, 登记时间, 发生时间, 预约时间, 分诊时间, 接收时间 as 接收时间, 诊室," & _
    "               号序,号码,结帐ID" & _
    "        From 门诊费用   " & _
    "        Where  记录状态 IN(1,3) and 记录性质=4 And Nvl( 附加标志,0)<>9 And  价格父号 is NULL ) B " & _
    " Where A.NO=B.NO And A.序号=B.序号 And A.收费细目ID=B.收费细目ID+0 " & _
    " Group by A.NO,A.序号,A.从属父号,A.费别,A.收费类别, A.收费细目ID," & _
    "       A.计算单位" & _
    " Having Sum(A.付数*A.数次)<>0"
 
    strSQL = _
    " Select /*+ Rule */ A.NO,A.序号,A.从属父号,A.费别,Q.编码 as 类别码,Q.名称 as 类别名,B1.编码,Nvl(B.名称,B1.名称) as 名称," & _
    "       Nvl(A.费用类型,B1.费用类型) 费用类型,A.计算单位,A.收费细目ID,A.准退数次 as 数次,A.单价,A.医嘱序号," & _
    "       A.剩余应收 as 应收金额,A.剩余实收 as 实收金额," & _
    "       C1.名称 as 执行科室,A.执行部门ID,M.名称 as 开单科室,A.结帐ID,A.记录标志,  " & _
    "       A.操作员姓名,A.医生,A.登记时间,A.发生时间,A.预约时间,A.分诊时间,A.接收时间,A.诊室,A.号序,A.号码, " & _
    "       A.原始数量,A.准退数量,A.剩余数量" & _
    " From (" & strSQL & ") A,收费项目目录 B1,部门表 C1,部门表 M,收费项目别名 B, 收费项目类别 Q" & _
    " Where A.收费细目ID=B1.ID And A.执行部门ID=C1.ID And A.开单部门ID=M.ID And A.收费类别=Q.编码 And   " & _
    "       A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    " Order by A.NO,A.序号"
    On Error GoTo errHandle
    Set rsFeeList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    GetRegListData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ShowReclaimInvoice(ByVal strNos As String, ByRef strReclaimInvoice As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示和返回需要回收的发票
    '入参:strNos-当前的单据号,多个用逗号分离(如果是被充结算,则为补充结算单号)
    '出参:strReclaimInvoice-返回回收的发票号(多个用逗号分隔),格式:AAAA,BBB,....)
    '返回:显示或获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-10 17:53:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmReInvoiceTemp As frmReInvoice
    
    On Error GoTo errHandle
    
    Set frmReInvoiceTemp = New frmReInvoice
    If frmReInvoiceTemp.ShowMe(Me, strNos, 0, 0, strReclaimInvoice, True) = False Then Exit Function
    If Not frmReInvoiceTemp Is Nothing Then Unload frmReInvoiceTemp
    Set frmReInvoiceTemp = Nothing
    ShowReclaimInvoice = True
    ShowReclaimInvoice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CheckSelectItemCanDel(ByVal strNos As String) As Boolean
    '功能：判断选择的退费项目是否可以正常退费，主要检查并发，可能有的项目在提出单据出来后又被执行了
    '参数：
    '   strNos - 本次选择的退费单据号
    '返回：
    '   检查通过，返回True；否则，返回False
    '问题号：105429
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long, j As Long, k As Long
    Dim arrNo As Variant
    Dim dbl剩余数量 As Double, dbl本次数量 As Double
    
    On Error GoTo errHandler
    If Left(strNos, 1) = "," Then strNos = Mid(strNos, 2)
    strNos = Replace(strNos, "'", "")
    If GetFeeListData(strNos, rsTemp) = False Then Exit Function
    If rsTemp.EOF Then
        MsgBox "单据:" & strNos & " 中没有可退费的项目，不能退费！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    arrNo = Split(strNos, ",")
    For i = 0 To UBound(arrNo)
        With vsBill
            k = .FindRow(arrNo(i), , .ColIndex("单据号"))
            For j = k To vsBill.Rows - 1
                If .TextMatrix(j, .ColIndex("单据号")) <> arrNo(i) Then Exit For
                If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                    rsTemp.Filter = "NO='" & arrNo(i) & "' And 序号=" & .RowData(j)
                    If rsTemp.EOF Then
                        MsgBox "单据:" & arrNo(i) & " 中第 " & (j - k + 1) & " 行项目的剩余未退数量为零，不能退费！" & _
                            "请重新获取费用数据！", vbExclamation, gstrSysName
                        .Row = j: .SetFocus
                        Exit Function
                    ElseIf Val(Nvl(rsTemp!原始数量)) > 0 Then
                        '负数收费的不检查
                        dbl剩余数量 = Val(Nvl(rsTemp!付数, 1)) * Val(Nvl(rsTemp!数次))
                        dbl本次数量 = Val(.TextMatrix(j, .ColIndex("数量")))
                        If RoundEx(dbl本次数量, 6) > RoundEx(dbl剩余数量, 6) Then
                            MsgBox "单据:" & arrNo(i) & " 中第 " & (j - k + 1) & " 行项目的本次退费数量(" & _
                                FormatEx(dbl本次数量, 5) & ")大于了剩余未退数量(" & FormatEx(dbl剩余数量, 5) & ")，" & _
                                "不能退费！请重新获取费用数据！", vbExclamation, gstrSysName
                            .Row = j: .SetFocus
                            Exit Function
                        End If
                    End If
                End If
            Next
        End With
    Next
    CheckSelectItemCanDel = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDelXMLExpend() As String
    '获取传入三方卡退费接口zlRetuenCheck中strXMLExpend参数值
    If mbytMode = EM_RBDTY_退费 Then
        GetDelXMLExpend = ZlGetDelXMLExpendByGrid(Me.vsBill)
    ElseIf mbytMode = EM_MULTI_异常重退 Then
        GetDelXMLExpend = ZlGetDelXMLExpend(mstr结算序号, True)
    End If
End Function

