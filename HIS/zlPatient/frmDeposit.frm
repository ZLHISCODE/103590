VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeposit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "预交款单据"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDeposit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   11910
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2325
      Left            =   75
      ScaleHeight     =   2325
      ScaleWidth      =   11775
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1050
      Width           =   11775
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   18
         X1              =   1260
         X2              =   4845
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label lbl手机号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手 机 号"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   70
         Tag             =   "手 机 号 "
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label lbl未缴费用 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未缴费用 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   2640
         TabIndex        =   69
         Tag             =   "未缴费用 "
         ToolTipText     =   "未缴款的划价单费用合计"
         Top             =   1170
         Width           =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   19
         X1              =   3660
         X2              =   4845
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Label lbl医保预结 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保预结 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   5100
         TabIndex        =   68
         Tag             =   "医保预结 "
         ToolTipText     =   "医保预结金额"
         Top             =   795
         Width           =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   17
         X1              =   6105
         X2              =   7680
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblWorkUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工作单位"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   5100
         TabIndex        =   66
         Tag             =   "工作单位 "
         Top             =   1560
         Width           =   960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   16
         X1              =   6105
         X2              =   11640
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   15
         X1              =   1260
         X2              =   2430
         Y1              =   1815
         Y2              =   1815
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   14
         X1              =   3660
         X2              =   4845
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   13
         X1              =   8895
         X2              =   11640
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   12
         X1              =   6105
         X2              =   7680
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   11
         X1              =   1245
         X2              =   2415
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   10
         X1              =   1260
         X2              =   7725
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   9
         X1              =   6105
         X2              =   11640
         Y1              =   1815
         Y2              =   1815
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   8
         X1              =   3660
         X2              =   4845
         Y1              =   1815
         Y2              =   1815
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   7
         X1              =   1245
         X2              =   2415
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   6
         X1              =   8895
         X2              =   11640
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   5
         X1              =   8895
         X2              =   11640
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   4
         X1              =   8895
         X2              =   11640
         Y1              =   345
         Y2              =   345
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   3
         X1              =   5505
         X2              =   7080
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   2
         X1              =   3660
         X2              =   4380
         Y1              =   345
         Y2              =   345
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   0
         X1              =   2115
         X2              =   2835
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   1
         X1              =   780
         X2              =   1500
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label lblMemo 
         AutoSize        =   -1  'True
         Caption         =   "备    注 "
         Height          =   240
         Left            =   5100
         TabIndex        =   63
         Tag             =   "备    注 "
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院科室 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7890
         TabIndex        =   62
         Tag             =   "住院科室 "
         Top             =   405
         Width           =   1080
      End
      Begin VB.Label lbl未审费用 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未审费用 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   61
         Tag             =   "未审费用 "
         ToolTipText     =   "未审核的划价记账费用合计"
         Top             =   1170
         Width           =   1080
      End
      Begin VB.Label lbl应收款 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应 收 款 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7875
         TabIndex        =   60
         Tag             =   "应 收 款 "
         Top             =   1170
         Width           =   1080
      End
      Begin VB.Label lbl医疗付款方式 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医疗付款方式 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7410
         TabIndex        =   59
         Tag             =   "医疗付款方式 "
         Top             =   75
         Width           =   1560
      End
      Begin VB.Label lbl家庭地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   58
         Tag             =   "家庭地址 "
         Top             =   471
         Width           =   1080
      End
      Begin VB.Label lbl担保金额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保金额 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   2640
         TabIndex        =   56
         Tag             =   "担保金额 "
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label lbl担保人 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担 保 人 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   55
         Tag             =   "担 保 人 "
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label lbl费别等级 
         AutoSize        =   -1  'True
         Caption         =   "费别 "
         Height          =   240
         Left            =   4965
         TabIndex        =   54
         Tag             =   "费别 "
         Top             =   90
         Width           =   600
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   1560
         TabIndex        =   53
         Tag             =   "年龄 "
         Top             =   90
         Width           =   600
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   52
         Tag             =   "性别 "
         Top             =   105
         Width           =   600
      End
      Begin VB.Label lbl预交余额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交余额 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   2640
         TabIndex        =   51
         Tag             =   "预交余额 "
         Top             =   795
         Width           =   1080
      End
      Begin VB.Label lbl床号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   3120
         TabIndex        =   50
         Tag             =   "床号 "
         Top             =   90
         Width           =   600
      End
      Begin VB.Label lbl剩余款额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "剩余款额 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   5100
         TabIndex        =   49
         Tag             =   "剩余款额 "
         Top             =   1170
         Width           =   1080
      End
      Begin VB.Label lbl费用余额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未结费用 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7890
         TabIndex        =   48
         Tag             =   "未结费用 "
         ToolTipText     =   "未审核的划价记账费用合计"
         Top             =   795
         Width           =   1080
      End
      Begin VB.Label lbl帐户余额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "帐户余额 "
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   47
         Tag             =   "帐户余额 "
         Top             =   795
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.PictureBox picList 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   0
      ScaleHeight     =   1650
      ScaleWidth      =   11910
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   6660
      Width           =   11910
      Begin VB.CheckBox chk仅显示本次预交 
         Caption         =   "仅显示本次预交"
         Height          =   240
         Left            =   9360
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Frame Frame3 
         Caption         =   "预交清单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   -30
         TabIndex        =   44
         Top             =   0
         Width           =   12015
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
         Height          =   1335
         Left            =   135
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   2355
         _Version        =   393216
         ForeColor       =   -2147483641
         FixedCols       =   0
         RowHeightMin    =   250
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   11910
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   8310
      Width           =   11910
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   420
         Left            =   150
         TabIndex        =   34
         Top             =   60
         Width           =   1500
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   420
         Left            =   10335
         TabIndex        =   32
         ToolTipText     =   "热键:Esc"
         Top             =   45
         Width           =   1500
      End
      Begin VB.CommandButton cmdSetup 
         Caption         =   "打印设置(&S)"
         Height          =   420
         Left            =   1770
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "热键：F10"
         Top             =   60
         Width           =   1620
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   420
         Left            =   8760
         TabIndex        =   28
         ToolTipText     =   "热键：F2"
         Top             =   45
         Width           =   1500
      End
   End
   Begin VB.PictureBox picNO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   75
      ScaleHeight     =   990
      ScaleWidth      =   11760
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   0
      Width           =   11755
      Begin VB.TextBox txtFact 
         ForeColor       =   &H00C00000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   6300
         MaxLength       =   50
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "热键：F3"
         Top             =   570
         Width           =   2370
      End
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9540
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "热键：F12"
         Top             =   570
         Width           =   1830
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "退"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11355
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "热键：F8"
         Top             =   555
         Width           =   420
      End
      Begin VB.Label lblPatientNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   345
         TabIndex        =   57
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据号"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   5565
         TabIndex        =   35
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lblFlag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "退"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   11355
         TabIndex        =   41
         Top             =   570
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交款单据"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   135
         TabIndex        =   45
         Top             =   45
         Width           =   1875
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8760
         TabIndex        =   40
         Top             =   630
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   36
      Top             =   8910
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmDeposit.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16034
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
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
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picFace 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   75
      ScaleHeight     =   3225
      ScaleWidth      =   11775
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3480
      Width           =   11775
      Begin VB.CheckBox chkAllCash 
         Caption         =   "三方账户强制退现"
         Enabled         =   0   'False
         Height          =   240
         Left            =   9360
         TabIndex        =   3
         Top             =   195
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.ComboBox cboPatiPage 
         Height          =   360
         Left            =   3675
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   570
         Width           =   1335
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   585
         TabIndex        =   64
         Top             =   135
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   635
         Appearance      =   2
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
         BackColor       =   -2147483633
      End
      Begin VB.CheckBox chk担保temp 
         Caption         =   "临时担保"
         Enabled         =   0   'False
         Height          =   240
         Left            =   7995
         TabIndex        =   2
         Top             =   195
         Width           =   1335
      End
      Begin VB.ComboBox cboType 
         Height          =   360
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   570
         Width           =   1380
      End
      Begin VB.TextBox txtMan 
         Enabled         =   0   'False
         Height          =   360
         Left            =   7980
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2760
         Width           =   3705
      End
      Begin VB.TextBox txtCode 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7995
         MaxLength       =   30
         TabIndex        =   17
         Top             =   1440
         Width           =   3690
      End
      Begin VB.TextBox txtUnit 
         Height          =   360
         Left            =   7995
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1005
         Width           =   3690
      End
      Begin VB.TextBox txt帐号 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7980
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1890
         Width           =   3705
      End
      Begin VB.ComboBox cboUnit 
         Height          =   360
         Left            =   7995
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   585
         Width           =   3690
      End
      Begin VB.ComboBox cboNote 
         Height          =   360
         Left            =   1230
         TabIndex        =   23
         Text            =   "cboNote"
         Top             =   2325
         Width           =   10485
      End
      Begin VB.TextBox txt开户行 
         Height          =   360
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   19
         Top             =   1890
         Width           =   3765
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1230
         TabIndex        =   1
         ToolTipText     =   "热键：F11"
         Top             =   135
         Width           =   3765
      End
      Begin VB.ComboBox cboStyle 
         Height          =   360
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1440
         Width           =   3765
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   360
         Left            =   1230
         TabIndex        =   25
         Top             =   2760
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin MSCommLib.MSComm com 
         Left            =   -330
         Top             =   2070
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSMask.MaskEdBox txtMoney 
         Height          =   360
         Left            =   1230
         TabIndex        =   11
         Top             =   1005
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   8388608
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label lblPatiPage 
         AutoSize        =   -1  'True
         Caption         =   "住院次数"
         Height          =   240
         Left            =   2685
         TabIndex        =   6
         Top             =   615
         Width           =   960
      End
      Begin VB.Label lblRepairMoney 
         Caption         =   "补交额:"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   5010
         TabIndex        =   65
         Top             =   1050
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label lbl预交类型 
         AutoSize        =   -1  'True
         Caption         =   "预交类型"
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   630
         Width           =   960
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缴款科室"
         Height          =   240
         Left            =   6960
         TabIndex        =   8
         Top             =   645
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "帐号"
         Height          =   240
         Left            =   7440
         TabIndex        =   20
         Top             =   1950
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开户行"
         Height          =   240
         Left            =   435
         TabIndex        =   18
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label lbl缴款单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缴款单位"
         Height          =   240
         Left            =   6960
         TabIndex        =   12
         Top             =   1065
         Width           =   960
      End
      Begin VB.Line Line1 
         X1              =   -135
         X2              =   7755
         Y1              =   -30
         Y2              =   -30
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   120
         TabIndex        =   0
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblMoney 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "金额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   675
         TabIndex        =   10
         Top             =   1065
         Width           =   510
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算号码"
         Height          =   240
         Left            =   6960
         TabIndex        =   16
         Top             =   1500
         Width           =   960
      End
      Begin VB.Label lblStyle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "支付方式"
         Height          =   240
         Left            =   195
         TabIndex        =   14
         Top             =   1500
         Width           =   960
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要"
         Height          =   240
         Left            =   645
         TabIndex        =   22
         Top             =   2385
         Width           =   480
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收款时间"
         Height          =   240
         Left            =   195
         TabIndex        =   24
         Top             =   2820
         Width           =   960
      End
      Begin VB.Label lblMan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收款员"
         Height          =   240
         Left            =   7200
         TabIndex        =   26
         Top             =   2820
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'说明：
'1.退款有两种方式,缺省的方式是在管理界面对指定的单据执行退款功能，或在正常收款状态下使用退款功能，另一种方式
'是以正常收款状态收款,但收款金额以负数表示退款，此时(退款金额<=病人余额)。这两种方式都不影响病人预交款的统计

'入口参数----------------------------------------------------------------------------------
Private mbytInState As Byte '0-收预交款(缺省,可切换到退),1-浏览单据(1),2-作废状态(1); 3-余额退款(37770), 4-转预交
Private mstrInNO As String '要浏览或退款的单据号(mbytInState=1或3时有效),从病人信息登记中调用退卡时为空
Private mblnNOMoved As Boolean '显明细时记录当前选择的单据是否在在线数据表中,以其它操作时无需再判断
Private mblnViewCancel As Boolean '是否浏览退款单据(mbytInState=1时有效)
Private mstrPrivs As String
Private mlngModul As Long
Private mbytCallObject As Byte '调用的对象(0-预交款管理调用;1-病人费用查询调用;2-医疗卡...
Private mlng病人ID As Long, mlng主页ID As Long, mdblDefPreMoney As Double
Private mbytPrepayType As Byte   ' 1-门诊预交;2-住院预交(4时,1,门诊转住院;2时住院转门诊)
Private mblnNotClick As Boolean
Private mstrbrPassWord As String
'程序变量----------------------------------------------------------------------------------
Private mblnUnLoad  As Boolean '用于控制窗体直接退出
Private mrsInfo As New ADODB.Recordset '病人信息(病人ID,姓名,性别,年龄,住院号,床号,在院标志)
Private mdbl剩余款额 As Double
Private mdbl预交余额 As Double
Private mdbl费用余额 As Double
Private mdbl预交余额_三方 As Double, mdbl预交余额_三方备份 As Double
Private mlng领用ID As Long, mstrCardPrivs As String
Private mstr代收款 As String
Private mstrRedFact As String
Private mstr缺省结算方式 As String
Private mblnOK As Boolean, mstr退款操作员 As String
Private mstrPrintDate As String
Private mbln未入科不交预交 As Boolean '51628
Private mbln住院退预交验证 As Boolean   '63113:刘尔旋,2013-10-29,住院预交退款验证
Private mbln允许在院病人余额退款 As Boolean
Private mblnNurseCall As Boolean

Private Enum BalanceType
    C1现金 = 1
    C2非现金 = 2
    C3个人帐户 = 3
    C4医保统筹 = 4
    C5代收款 = 5
End Enum
'医保变量----------------------
Private mcur帐户余额 As Currency '个人帐户余额
Private mstr个人帐户 As String '个人帐户结算方式
Private mstr病人类型 As String
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
'关于结算卡的的处理变量
Private Type Ty_SquareCard
    blnExistsObjects As Boolean     '安装了结算卡的的
    dbl刷卡总额 As Double
    bln卡结算 As Boolean '当前读取的单据是卡结算
End Type
Private mtySquareCard As Ty_SquareCard

'Private mobjSquareCard As Object
Private mblnClickSquareCtrl As Boolean
'短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
Private mcolPayMode As Collection   '卡支付方式
Private mlngCardTypeID  As Long
Private mbln消费卡     As Boolean
Private mstr结算方式      As String
Private mstrBrushCardNo As String
 
Private mlng医疗卡长度 As Long
Private Type Ty_BillInfor
    lng预交ID As Long
    strNo As String
    lng卡类别ID As Long
    bln消费卡 As Boolean
    str卡号 As String
    str名称 As String
    str交易流水号 As String
    str交易说明 As String
    str合作单位 As String
    dbl金额 As Double
    bln转账 As Boolean
    bln退款验卡 As Boolean
    dt收款时间 As Date
    lng消费卡ID As Long
End Type
Private mcurBill As Ty_BillInfor
Private mFactProperty As Ty_FactProperty
Private mblnStartFactUseType As Boolean '是否启用的相关的门诊类别的
Private mblnPassInputCardNo As Boolean  '是否密文输入卡号
Private mblnDefaultPassInputCardNo As Boolean '缺省刷卡是否密文输入卡号
Private mrsDepositBalance As ADODB.Recordset    '当前病人的预交余额
Private mrsDepositInfor As ADODB.Recordset    '当前病人预交情况(按各类型及相关的流水号分类汇总)
Private mbytBackMoneyType As Byte '退款方式:1-禁止;0-提示
Private mbytOracleBackType As Byte '退款检查_In;0-忽略退款金额是否大于了病人余额；1-检查退款金额
Private mblnClearWinInfor As Boolean  '缴款后,是否清除窗体信息
Private mblnCheckPass As Boolean '刷卡时要求输入密码,'0000000000'依位顺序表示各个场合,分别为:1.门诊挂号,2.门诊划价,3.门诊收费,4.门诊记帐,5.入院登记,6.住院记帐,7.病人结帐,8.病人预交款,9.检验技师站,10.影像医技站.'
'外挂评价器对象
Private mobjPlugIn As Object
Private mstrPatiOld As String
Private mstrPatiSex As String
Private mblnOneCard As Boolean  '是否只有一张就诊卡
Private mlngFactModule As Long '发票相关参数模块号

Public Function zlShowEdit(ByVal frmMain As Object, ByVal bytCallObject As Byte, _
    ByVal bytInState As Byte, _
    ByVal strPrivs As String, ByVal lngModule As Long, Optional ByVal bytPrepayType As Byte = 0, Optional strInNo As String = "", _
    Optional ByVal blnViewCancel As Boolean = False, Optional blnNOMoved As Boolean = False, _
    Optional ByVal lng病人ID As Long = 0, Optional lng主页ID As Long = 0, Optional dblDefPreMoney As Double = 0, _
    Optional ByVal blnNurseCall As Boolean = False, _
    Optional blnOneCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口,用于病人预交款信息编辑或查看
    '入参:frmMain-调用的主窗口
    '        bytCallObject:调用的对象(0-预交款管理调用;1-病人费用查询调用;2-医疗卡调用)...
    '        bytInState:0-收预交款(缺省,可切换到退),1-浏览单据(1),2-作废状态(1);3-余额退款(37770)
    '        bytPrepayType-预交类型(0-门诊和住院;1-门诊;2-住院)
    '        strInNo:要浏览或退款的单据号(mbytInState=1或3时有效),从病人信息登记中调用退卡时为空
    '         blnViewCancel:是否浏览退款单据(mbytInState=1时有效)
    '        blnNOMoved:显明细时记录当前选择的单据是否在在线数据表中,以其它操作时无需再判断
    '        dblDefPreMoney-缺省的缴款金额(目前只有病人费用查询中调用时才有效)
    '        blnNurseCall-护士站调用
    '出参:
    '返回:预交款只有一次成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-17 16:11:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    mbytCallObject = bytCallObject: mbytInState = bytInState: mstrPrivs = strPrivs: mlngModul = lngModule
    mstrInNO = strInNo: mblnViewCancel = blnViewCancel: mblnNOMoved = blnNOMoved
    mlng病人ID = lng病人ID: mlng主页ID = lng主页ID: mdblDefPreMoney = dblDefPreMoney
    mbytPrepayType = bytPrepayType
    mblnNurseCall = blnNurseCall
    mblnOneCard = blnOneCard
    mlngFactModule = IIf(mbytCallObject = 2, 1107, mlngModul)
    
    mblnOK = False
    If frmMain Is Nothing Then
        frmDeposit.Show
    Else
        frmDeposit.Show 1, frmMain
    End If
    zlShowEdit = mblnOK
End Function
 
Private Sub cboPatiPage_Click()
    If txtPatient.Tag <> "" And mbytInState = 0 And Not mrsInfo Is Nothing And mrsInfo.State = 1 Then
        If cboPatiPage.ItemData(cboPatiPage.ListIndex) <> Nvl(mrsInfo!主页ID, 0) Then
            Call ShowPatiPageInfo
        End If
    End If
    Call ShowHistoryPrepay("")
End Sub

Private Sub ShowPatiPageInfo()
    Dim lng主页ID As Long
    lng主页ID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    '根据第几次入院更新信息
    Call GetPatient(IDKind.GetfaultCard, txtPatient.Tag, False, False, lng主页ID)
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(Val(Nvl(mrsInfo!住院号)) = 0, "", "住院号:" & mrsInfo!住院号 & "   ") & _
                       IIf(Val(Nvl(mrsInfo!门诊号)) = 0, "", "门诊号:" & mrsInfo!门诊号)
    lbl费别等级.Caption = lbl费别等级.Tag & mrsInfo!费别
    txtPatient.Text = mrsInfo!姓名
    txtPatient.Tag = mrsInfo!病人ID
    lblSex.Caption = lblSex.Tag & IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别)
    lblOld.Caption = lblOld.Tag & IIf(IsNull(mrsInfo!年龄), "", mrsInfo!年龄)
    lbl医疗付款方式.Caption = lbl医疗付款方式.Tag & Nvl(mrsInfo!医疗付款方式)
    lbl科室.Caption = lbl科室.Tag & GET部门名称(mrsInfo!科室ID)
    lbl床号.Caption = lbl床号.Tag & IIf(mrsInfo!床号 = 0, "家庭", mrsInfo!床号)
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(Val(Nvl(mrsInfo!当前科室id)) = 0, Val(Nvl(mrsInfo!科室ID)), Val(Nvl(mrsInfo!当前科室id))))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
End Sub
Private Sub cboPatiPage_KeyDown(KeyCode As Integer, Shift As Integer)
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboType_Click()
    If cboType.ListIndex < 0 Then Exit Sub
    
    '88657:李南春，2015/9/17,切换预交类型刷新预交余额
    If mbytInState = 0 And chkCancel.Value = 0 Or mbytInState = 3 Then
        mlng领用ID = 0
        '问题号:112784,焦博,2017/10/13,获取正确的票据格式
        mFactProperty = zl_GetInvoicePreperty(mlngFactModule, 2, cboType.ItemData(cboType.ListIndex))
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact
        Call ShowPremayBalance(True, 0)
        Call SetCtrlEnabled
        Call ShowHistoryPrepay("")
    ElseIf mbytInState = 4 Then
        mlng领用ID = 0
        '问题号:112784,焦博,2017/10/13,获取正确的票据格式
        mFactProperty = zl_GetInvoicePreperty(mlngFactModule, 2, IIf(cboType.ItemData(cboType.ListIndex) = 1, 2, 1))
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact
        Call ShowPremayBalance(True, 0)
        Call SetCtrlEnabled
    '问题号:114482,焦博,2017/10/10,用户在缴款界面操作预交时，根据右上角“退”按钮来确认是否打印红票。
    ElseIf mbytInState = 2 Or chkCancel.Value = 1 Then
        mlng领用ID = 0
        '问题号:112784,焦博,2017/10/13,获取正确的票据格式
        mFactProperty = zl_GetInvoicePreperty(mlngFactModule, 12, cboType.ItemData(cboType.ListIndex))
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact(False, True)
    End If
    
     '问题号:45666
    If mbytInState = 0 And cboType.Text = "住院预交" Then '交预交款
        chk仅显示本次预交.Visible = True
        chk仅显示本次预交.Value = IIf(zldatabase.GetPara("仅显示本次预交", glngSys, mlngModul, , Array(chk仅显示本次预交), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
    Else
        chk仅显示本次预交.Visible = False
    End If
    lblPatiPage.Visible = cboType.Text = "住院预交": cboPatiPage.Visible = cboType.Text = "住院预交"
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

 
Private Sub chkAllCash_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk担保temp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk仅显示本次预交_Click()
    Call ShowHistoryPrepay("")
End Sub

Private Sub IDKind_Click(objcard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXml As String
    
    If objcard.名称 Like "IC卡*" And objcard.系统 Then
        If mobjICCard Is Nothing Then
               Set mobjICCard = New clsICCard
               Call mobjICCard.SetParent(Me.hWnd)
               Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text = "" Then Exit Sub
        Call FindPati(objcard, False, txtPatient.Text)
        Exit Sub
    End If
     
    lng卡类别ID = objcard.接口序号
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text = "" Then Exit Sub
    Call FindPati(objcard, False, txtPatient.Text)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objcard As zlIDKind.Card)
    Call txtPatient_GotFocus
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub
Private Sub IDKind_ReadCard(ByVal objcard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text = "" Then Exit Sub
    Call FindPati(objcard, True, txtPatient.Text)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objcard As Card
    Set objcard = IDKind.GetIDKindCard("IC卡", CardTypeName)
    If objcard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objcard, True, txtPatient.Text)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objcard As Card
    Set objcard = IDKind.GetIDKindCard("身份证", CardTypeName)
    If objcard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objcard, True, txtPatient.Text)
End Sub
Private Sub SetcmdOkEnabled()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置cmdOk的neable属性
    '编制：刘兴洪
    '日期：2010-07-09 16:24:53
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    If mrsInfo Is Nothing Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = mrsInfo.State = adStateOpen
    End If
    chk仅显示本次预交.Enabled = cmdOK.Enabled
End Sub
Private Sub SetCtrlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Enabled属性
    '编制:刘兴洪
    '日期:2011-07-24 09:30:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean, objCtl As Control
    blnEdit = mbytInState <> 1
    Select Case mbytInState
    Case 4  '转预交
        cboType.Enabled = True
        blnEdit = False
        cboUnit.Enabled = blnEdit
        txtUnit.Enabled = blnEdit
        cboStyle.Enabled = blnEdit: cboStyle.ListIndex = -1
        txtCode.Enabled = blnEdit: txt开户行.Enabled = blnEdit
        txt帐号.Enabled = blnEdit: cboNote.Enabled = blnEdit
    Case Else
        If cboStyle.ListIndex < 0 Then GoTo goEnd:
        Select Case cboStyle.ItemData(cboStyle.ListIndex)
        Case 3 '三方接口
            txtUnit.Enabled = False: txt开户行.Enabled = False
            txt帐号.Enabled = False
        Case 1 '现金
            '现金
            txtUnit.Enabled = False: txt开户行.Enabled = False
            txt帐号.Enabled = False: txtCode.Enabled = False
        Case 2
            blnEdit = cboStyle.Text Like "*票*" Or cboStyle.Text Like "*卡*"
            txtCode.Enabled = blnEdit
            txtUnit.Enabled = True: txt开户行.Enabled = True: txt帐号.Enabled = True
        Case Else
            txtUnit.Enabled = True: txt开户行.Enabled = True: txt帐号.Enabled = True
        End Select
    End Select
goEnd:
    For Each objCtl In Me.Controls
        Select Case UCase(TypeName(objCtl))
        Case UCase("ComBobox")
            objCtl.BackColor = IIf(objCtl.Enabled, &H80000005, Me.BackColor)
        Case UCase("TextBox")
            objCtl.BackColor = IIf(objCtl.Enabled, &H80000005, Me.BackColor)
        Case Else
        End Select
    Next
End Sub

Private Sub cboStyle_Click()
    '当选择支票时才处理上次缴款信息
    Dim strInfo As String, dbl固定金额 As Double
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long
    Dim blnFind As Boolean '问题号:55666
    If mbytInState = 2 Or chkCancel.Value = 1 Then Exit Sub
    If mbytInState = 4 Then Exit Sub
    
    If cboStyle.ListIndex = -1 Then Exit Sub
        
    '问题号:111657,焦博,2017/07/25,使用现金支付预交款时,任会产生三方卡号
    mstrBrushCardNo = ""     '清空三方交易时缓存的卡号
    mcurBill.bln转账 = False
    mcurBill.lng预交ID = 0
    lngIndex = cboStyle.ListIndex + 1
''    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
   If Not mcolPayMode Is Nothing Then
        '问题:56478
        If zlCommFun.GetNeedName(cboStyle.Text) = zlCommFun.GetNeedName(mstr个人帐户) Then
            mlngCardTypeID = 0
            mbln消费卡 = False
            mstr结算方式 = zlCommFun.GetNeedName(mstr个人帐户)
        Else
            mlngCardTypeID = Val(mcolPayMode(lngIndex)(3))
            mbln消费卡 = Val(mcolPayMode(lngIndex)(5)) = 1
            mstr结算方式 = Trim(mcolPayMode(lngIndex)(6))
        End If
        Call ShowPremayBalance(False, 0)
    End If
    Call SetCtrlEnabled
    txtMoney.Enabled = True
    Select Case cboStyle.ItemData(cboStyle.ListIndex)
    Case 3, 1
        txtUnit.Text = "": txt开户行.Text = "": txt帐号.Text = ""
    Case 2
        If cboStyle.Text Like "*票*" Or cboStyle.Text Like "*卡*" Then
            '无支票这种结算性质,所以用名称
            '问题:36611
            If mrsInfo Is Nothing Then Exit Sub
            If mrsInfo.State = adStateClosed Then Exit Sub
            If mrsInfo.EOF Then Exit Sub
            strInfo = GetLastInfo(mrsInfo!病人ID)
            If strInfo <> "" Then
                txtUnit.Text = IIf(Split(strInfo, "|")(0) = "", txtUnit.Text, Split(strInfo, "|")(0))
                txt开户行.Text = IIf(Split(strInfo, "|")(1) = "", txt开户行.Text, Split(strInfo, "|")(1))
                txt帐号.Text = IIf(Split(strInfo, "|")(2) = "", txt帐号.Text, Split(strInfo, "|")(2))
            End If
        End If
    Case 5 ''缺省金额:34705
        varData = Split(mstr代收款, "|"): dbl固定金额 = 0
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i), ":")
            If varTemp(0) = Split(cboStyle.Text & "-", "-")(0) Then
                dbl固定金额 = Val(varTemp(1)): Exit For
            End If
        Next
        If dbl固定金额 <> 0 Then
            txtMoney.Text = Format(dbl固定金额, "##,##0.00;-##,##0.00; ;"): txtMoney.Enabled = False
             txtMoney.Tag = dbl固定金额:
        End If
    End Select
'    '问题号:55666
'     '新单存盘
'    If mrsInfo.State = adStateClosed Then
'        If txtPatient.Visible And cboStyle.Text Like "*卡*" Then
'            MsgBox "没有确定收取预交款的病人,不能进行刷卡操作！", vbExclamation, gstrSysName
'            '设置默认还原成为现金支付
'            For i = 0 To cboStyle.ListCount
'                If cboStyle.List(i) = "现金" Then
'                    blnFind = True
'                    cboStyle.ListIndex = i
'                End If
'            Next
'            If blnFind And cboStyle.ListCount > 0 Then cboStyle.ListIndex = 0: blnFind = False
'        End If
'        If txtPatient.Visible Then txtPatient.SetFocus: Exit Sub
'    End If
'    If IIf(Trim(txtMoney.Text) = "", "0", Trim(txtMoney.Text)) = "0" And txtMoney.Visible And Not mrsInfo Is Nothing And Not txtMoney Is ActiveControl Then
'        MsgBox "没有输入充值金额,不能进行刷卡操作！", vbExclamation, gstrSysName
'        txtMoney.SetFocus
'        For i = 0 To cboStyle.ListCount
'                If cboStyle.List(i) = "现金" Then
'                    blnFind = True
'                    cboStyle.ListIndex = i
'                End If
'            Next
'            If blnFind And cboStyle.ListCount > 0 Then cboStyle.ListIndex = 0: blnFind = False
'        Exit Sub
'    End If
'    '刷卡
'    CheckBrushCard
End Sub

Private Sub cboStyle_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then
        If cboStyle.ListIndex = -1 Then
            Beep
        Else
            'Call cboStyle_Click
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        If cboStyle.Locked Then Exit Sub
        If KeyAscii >= 32 Then
            lngIdx = cbo.MatchIndex(cboStyle.hWnd, KeyAscii)
            If lngIdx = -1 And cboStyle.ListCount > 0 Then lngIdx = 0
            cboStyle.ListIndex = lngIdx
        End If
    End If
End Sub

Private Sub cboStyle_Validate(Cancel As Boolean)
    If cboStyle.Locked Then Exit Sub
    If Not (cboStyle.ListIndex > -1 And (mbytInState = 0 Or mbytInState = 3)) Then Exit Sub
    If cboStyle.ItemData(cboStyle.ListIndex) = BalanceType.C5代收款 Then
         If mbytInState = 0 Then
             If InStr(mstrPrivs, ";代收款收取;") = 0 Then
                 MsgBox "你没有权限进行代收款收取操作！", vbInformation, gstrSysName
                 If cbo.Locate(cboStyle, BalanceType.C1现金, True) Then Cancel = True
             End If
         Else
             If InStr(1, mstrPrivs, ";代收款退款;") = 0 Then
                 MsgBox "你没有权限进行代收款的退款操作！", vbInformation, gstrSysName
                 If cbo.Locate(cboStyle, BalanceType.C1现金, True) Then Cancel = True
             End If
         End If
     ElseIf mbytInState = 0 Then
         If InStr(1, mstrPrivs, ";预交收款;") = 0 Then
             MsgBox "你没有权限进行预交收款操作！", vbInformation, gstrSysName
             If cbo.Locate(cboStyle, BalanceType.C5代收款, True) Then Cancel = True
         End If
     Else
         If InStr(1, mstrPrivs, ";预交退款;") = 0 Then
             MsgBox "你没有权限进行预交退款操作！", vbInformation, gstrSysName
             If cbo.Locate(cboStyle, BalanceType.C5代收款, True) Then Cancel = True
         End If
     End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If SendMessage(cboUnit.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cboUnit.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cboUnit.ListIndex = lngIdx
    '强制要选中一个(第一个)
    If cboUnit.ListIndex = -1 And cboUnit.ListCount <> 0 Then cboUnit.ListIndex = 0
End Sub

Private Sub chkCancel_Click()
    Dim ctlTmp As Control
    Dim strTmp As String
    
    IDKind.Enabled = (chkCancel.Value <> Checked)
    
    If chkCancel.Value = Checked Then
        '按下
        chkCancel.ForeColor = &HFF&
        '清除相关界面和数据
        Set mrsInfo = New ADODB.Recordset '清除病人信息
        txtPatient.Text = "": txtPatient.Locked = True
        Call SetMoneyInfo(True)
                
        txtMoney.Text = "" '可输入退部分款
        cboStyle.ListIndex = -1: cboStyle.Locked = True
        txtCode.Text = "": txtCode.Locked = True
        txtMan.Text = ""
        txtDate.Text = "____-__-__": txtDate.Enabled = False
        cboNote.ListIndex = cboNote.ListCount - 1
                
        picFace.Enabled = True '！！！不允许部份退款！！！
        For Each ctlTmp In Me.Controls
           If ctlTmp.Name <> "com" Then
                If ctlTmp.Container.Name = "picFace" Then
                     If InStr(1, "cboNote,lblNote,txtMan,txtDate", ctlTmp.Name) <= 0 Then
                         strTmp = UCase(TypeName(ctlTmp))
                         If strTmp <> "LABEL" And strTmp <> "LINE" Then
                             On Error Resume Next     'MASKEDBOX不支持locked属性
                             ctlTmp.Enabled = False
                             If strTmp <> "MASKEDBOX" Then ctlTmp.Locked = True    '必须设locked，因为readbill中cboStyle设置listindex时调用Click将会把enabled设置为true
                             On Error GoTo 0
                         End If
                     End If
                End If
           End If
        Next ctlTmp
                
        '待输入退款的单据号
        cboNO.Text = "": cboNO.Tag = ""
        cboNO.Locked = False
        txtFact.Text = ""
        txtFact.Locked = True
        If cboNO.Visible Then cboNO.SetFocus
    Else
        '弹起
        chkCancel.ForeColor = 0
        
        picFace.Enabled = True
        For Each ctlTmp In Me.Controls
           If ctlTmp.Name <> "com" Then
                If ctlTmp.Container.Name = "picFace" Then
                     If InStr(1, "cboNote,lblNote,txtMan,txtDate", ctlTmp.Name) <= 0 Then
                         strTmp = UCase(TypeName(ctlTmp))
                         If strTmp <> "LABEL" And strTmp <> "LINE" Then
                             On Error Resume Next       'MASKEDBOX不支持locked属性
                             ctlTmp.Enabled = True
                             If strTmp <> "MASKEDBOX" Then ctlTmp.Locked = False
                             On Error GoTo 0
                         End If
                     End If
                End If
           End If
        Next ctlTmp
        
        Call ClearBill
    End If
    Call SetCtrlEnabled
End Sub

Private Sub cmdCancel_Click()
    If Not mblnOK Then Unload Me: Exit Sub
    If mbytInState = 0 Then
        If chkCancel.Value = Checked Then
            If MsgBox("确实要放弃退款退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If mrsInfo.State = adStateOpen Then
                If MsgBox("该病人的预交款尚未收取,确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
    End If
    If mbytInState = 3 Then
        If mrsInfo.State = adStateOpen Then
            If MsgBox("该病人的尚未进行退款操作,确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("未进行退款操作,确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    Unload Me
End Sub
Private Sub zlBackDeposit()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：退预交操作
    '编制：刘兴洪
    '日期：2010-06-18 16:34:59
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, blnExistsSquare As Boolean '是否存在结算卡
    Dim blnCanDel As Boolean, intInsure As Integer
    Dim bln打印 As Boolean
    Dim msgBoxResult As String
    
    mbytOracleBackType = 1
'   退款
    If cboNO.Tag = "" Then
        MsgBox "该单据未正确读取,不能退款！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    '金额检查
    If txtMoney.Text = "" Then
        MsgBox "退款金额不能为空,请输入！", vbExclamation, gstrSysName
        Exit Sub
    ElseIf CCur(StrToNum(txtMoney.Text)) = 0 Then
        MsgBox "退款金额不能为零,请输入！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    
    '问题27363 by lesfeng 2010-01-13
    If MsgBox("确实要将单据 " & cboNO.Text & " 作废吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Select Case mFactProperty.intInvoicePrint
        Case 0 '不打印预交发票
           bln打印 = False
        Case 1 '自动打印
           bln打印 = True
        Case 2 '打印提醒
            msgBoxResult = MsgBox("是否需要打印预交红票？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
            bln打印 = (msgBoxResult = vbYes)
        End Select
        
        If mcurBill.lng卡类别ID = 0 Then
            If Not is代收款(cboNO.Text) And gbyt预存款消费验卡 <> 0 Then
                If mbln住院退预交验证 Or cboType.ItemData(cboType.ListIndex) = 1 Then
                    If Not zldatabase.PatiIdentify(Me, glngSys, Val(txtPatient.Tag), Val(StrToNum(txtMoney.Text)), _
                        , , , , , , , (gbyt预存款消费验卡 = 2)) Then Exit Sub
                End If
            End If
        End If
        '医保相关检查
        blnCanDel = True '缺省为支持,考虑过程的一般化处理
        intInsure = ExistInsure(cboNO.Text)
        If intInsure > 0 Then
            '去掉了医保连接匹配检查
            blnCanDel = gclsInsure.GetCapability(support预交退个人帐户, Val(txtPatient.Tag), intInsure)
        End If
        
        '并发操作判断
        If cboStyle.ItemData(cboStyle.ListIndex) <> BalanceType.C5代收款 Then
            Dim dbl预交余额 As Double
            dbl预交余额 = HaveSpare(cboNO.Text)
            If dbl预交余额 = 0 And InStr(mstrPrivs, ";预交结清退款;") = 0 Then
                MsgBox "该病人已没有预交余额，你没有权限作废这张单据！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If HaveBalance(cboNO.Text) <> 0 Then 'And InStr(mstrPrivs, ";预交结帐退款;") = 0 '删掉 预交结帐退款权限 54779
                MsgBox "该笔预交已经被病人在结帐时使用，你不能作废这张单据！", vbInformation, gstrSysName
                Exit Sub
            End If
            '87858
            If CCur(StrToNum(txtMoney.Text)) > dbl预交余额 Then
                '46067
                If mbytBackMoneyType = 1 Then
                    '负数退款,不能大于他本身的余额:37375
                    Call MsgBox("该笔预交的金额比病人当前的余额多，你不能作废这张单据！", vbInformation + vbOKOnly, gstrSysName)
                    Exit Sub
                Else
                    If MsgBox("该笔预交的金额比病人当前的余额多，忽略吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                    mbytOracleBackType = 0
                End If
            End If
        End If
        
        cmdOK.Enabled = False   '防医保延时
        
        '检查三方接口交易是否合法
        '108666:李南春，2017/5/9，恢复确认按钮可用状态
        If zlCheckDepositDelValied(Val(cboNO.Tag), StrToNum(txtMoney.Text)) = False Then cmdOK.Enabled = True: Exit Sub
        
        '执行作废操作
        If Not CancelBill(CLng(cboNO.Tag), blnCanDel, intInsure, bln打印) Then '退款
            MsgBox "操作失败,请重试该操作。如仍有问题,请与系统管理员联系！", vbExclamation, gstrSysName
            cmdOK.Enabled = True
            Exit Sub
        End If
        
        If bln打印 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_1", Me, "NO=" & mstrInNO, "收款时间=" & Format(Now, "yyyy-mm-dd HH:MM:SS"), _
                            IIf(mFactProperty.intInvoiceFormat = 0, "", "ReportFormat=" & mFactProperty.intInvoiceFormat), 2)
            Call zlCheckFactIsEnough
        End If
        
        
        Call RePrintBill '重新打印预交发票
        
        cmdOK.Enabled = True
        
        '医保改动
        For i = 0 To cboStyle.ListCount - 1
            If cboStyle.ItemData(i) = 3 Then
                cboStyle.RemoveItem i: Exit For
            End If
        Next
    End If
    If mbytInState <> 2 Then
        chkCancel.Value = Unchecked '(并激活事件)
    Else
        mblnOK = True
        Unload Me: Exit Sub '退款模式操作后退出
    End If
    mblnOK = True
    Call ClearBill
End Sub

Private Sub RePrintBill()
    '作废后重新打印预交发票
    Dim blnRePrint As Boolean, strNotDelNos As String, strSQL As String
    Dim objFactProperty As Ty_FactProperty
    Dim intInvoiceFormat As Integer, str收款时间 As String
    
    On Error GoTo errHandle
    strNotDelNos = GetTurnMZToZYMultiNOs(cboNO.Text, mblnNOMoved)
    If strNotDelNos = "" Then Exit Sub
    
    objFactProperty = zl_GetInvoicePreperty(mlngModul, 2, cboType.ItemData(cboType.ListIndex))
    Select Case objFactProperty.intInvoicePrint
    Case 0 '不打印预交发票
       blnRePrint = False
    Case 1 '自动打印
       blnRePrint = True
    Case 2 '打印提醒
        blnRePrint = MsgBox("当前预交单据为门诊费用转住院生成的，且该单据所在的发票同时打印了多张预交单据，" & _
            "是否对剩余单据重新打印预交票据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End Select
    If blnRePrint = False Then Exit Sub
    
    intInvoiceFormat = Val(zldatabase.GetPara(284, glngSys, , "0"))
    
    Call GetFact '重新获取发票号,因为当前发票号可能也被红票打印使用

    '票据号检查
    If gblnBill预交 Then
        If Trim(txtFact.Text) = "" Then
            MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
        Else
            mlng领用ID = CheckUsedBill(2, IIf(mlng领用ID > 0, mlng领用ID, objFactProperty.lngShareUseID), _
                txtFact.Text, cboType.ItemData(cboType.ListIndex))
            If mlng领用ID <= 0 Then
                Select Case mlng领用ID
                Case 0 '操作失败
                Case -1
                    MsgBox "你没有自用和共用的预交票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Case -3
                    MsgBox "票据号码不在当前有效领用范围内！", vbInformation, gstrSysName
                End Select
                txtFact.Text = ""
            End If
        End If
    Else
        If Len(txtFact.Text) <> gbyt预交 And txtFact.Text <> "" Then
            MsgBox "票据号码长度应该为 " & gbyt预交 & " 位！", vbInformation, gstrSysName
            txtFact.Text = ""
        End If
    End If
    
    If Trim(txtFact.Text) <> "" Then '发票号无效不打印
        '执行数据处理
        'Zl_病人预交记录_Reprint
        strSQL = "Zl_病人预交记录_Reprint("
        '  单据号_In Varchar2,
        strSQL = strSQL & "'" & strNotDelNos & "',"
        '  票据号_In 票据使用明细.号码%Type,
        strSQL = strSQL & "'" & Trim(txtFact.Text) & "',"
        '  领用id_In 票据使用明细.领用id%Type,
        strSQL = strSQL & "" & IIf(mlng领用ID = 0, "NULL", mlng领用ID) & ","
        '  使用人_In 票据使用明细.使用人%Type
        strSQL = strSQL & "'" & UserInfo.姓名 & "')"
        zldatabase.ExecuteProcedure strSQL, Me.Caption
        
        If Not gblnBill预交 Then
            '松散：保存当前号码
            zldatabase.SetPara "当前预交票据号", Trim(txtFact.Text), glngSys, mlngFactModule
        End If
        
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & strNotDelNos, _
            "收款时间=" & Format(mcurBill.dt收款时间, "yyyy-mm-dd HH:MM:SS"), _
            IIf(intInvoiceFormat = 0, "", "ReportFormat=" & intInvoiceFormat), 2)
        Call zlCheckFactIsEnough
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetTurnMZToZYMultiNOs(ByVal strNo As String, Optional ByVal blnNOMoved As Boolean) As String
    '功能：获取门诊转住院产生的预交单据，并返回一次打印的多张单据号
    '入参:strNo-需要重打NO
    '     blnNOMoved-是否转入历史表空间
    '出参:
    '返回:一次打印的多张单据号，格式：A001,A002,A003,...
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strNos As String
    
    On Error GoTo errHandle
    '应根据最后一次打印的情况来定
    strSQL = _
        "Select a.NO" & vbNewLine & _
        "From 票据打印内容 A" & vbNewLine & _
        "Where a.数据性质 = 2" & vbNewLine & _
        "      And a.ID In (Select ID" & vbNewLine & _
        "                From (Select b.Id" & vbNewLine & _
        "                      From 票据使用明细 A, 票据打印内容 B" & vbNewLine & _
        "                      Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 2 And b.No = [1]" & vbNewLine & _
        "                      Order By a.使用时间 Desc)" & vbNewLine & _
        "                Where Rownum < 2)" & vbNewLine & _
        "      And Not Exists(Select 1 From 病人预交记录 Where 记录性质 = 1 And 记录状态 = 2 And No = a.No)" & vbNewLine & _
        "Order By No"
    If blnNOMoved Then
        strSQL = Replace(strSQL, "票据打印内容", "H票据打印内容")
        strSQL = Replace(strSQL, "票据使用明细", "H票据使用明细")
    End If
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "", strNo)
    If rsTemp.EOF Then Exit Function
    
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & Nvl(rsTemp!NO)
            .MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    GetTurnMZToZYMultiNOs = strNos
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckDataValied() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查数据是否合法
    '返回：合法返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-18 16:38:39
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
   '新单存盘
  If mrsInfo.State = adStateClosed Then
      If mbytInState = 3 Then
            MsgBox "没有确定退预交款的病人,不能退款！", vbExclamation, gstrSysName
      Else
            MsgBox "没有确定收取预交款的病人,不能存盘！", vbExclamation, gstrSysName
      End If
      txtPatient.SetFocus: Exit Function
  End If
  
    If mbytInState = 3 And chkAllCash.Value = 1 Then
        If Val(cboStyle.ItemData(cboStyle.ListIndex)) <> 1 And Val(cboStyle.ItemData(cboStyle.ListIndex)) <> 2 Then
            MsgBox "三方账户强制退现的情况下，只能选择现金或者支票类的结算方式！", vbInformation, gstrSysName
            If cboStyle.Enabled And cboStyle.Visible Then cboStyle.SetFocus
            Exit Function
        End If
    End If
          
  If LenB(StrConv(txtUnit.Text, vbFromUnicode)) > 50 Then
      MsgBox "缴款单位名称只允许 50 个字符或 25 个汉字,请修改！", vbInformation, App.Title
      txtUnit.SetFocus: Exit Function
  End If
  If LenB(StrConv(txt开户行.Text, vbFromUnicode)) > 50 Then
      MsgBox "开户行名称只允许 50 个字符或 25 个汉字,请修改！", vbInformation, App.Title
      txt开户行.SetFocus: Exit Function
  End If
  If LenB(StrConv(cboNote.Text, vbFromUnicode)) > 50 Then
      MsgBox "缴款摘要只允许 50 个字符或 25 个汉字,请修改！", vbInformation, App.Title
      cboNote.SetFocus: Exit Function
  End If
  If mbytInState = 0 Then
    If cboType.ListIndex < 0 Then Exit Function
    '问题:44963
    If mrsInfo Is Nothing Then Exit Function
    If cboType.ItemData(cboType.ListIndex) = 2 Then
        If Val(Nvl(mrsInfo!在院)) = 0 And gblnAllowOut = False Then
            strSQL = "Select 1 From 病案主页 Where 病人ID=[1] And Nvl(主页ID,0)=0 And Nvl(病人性质,0)=0" '预入院
            Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsInfo!病人ID)))
            If rsTemp.EOF Then
                MsgBox "病人还未住院,不能缴住院预交,请检查!", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Else
        If Val(Nvl(mrsInfo!在院)) = 1 And gblnBanIn = True Then
            MsgBox "病人还未出院,不能缴门诊预交,请检查!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
  End If
  '金额检查
  '问题27363 by lesfeng 2010-01-13
  If txtMoney.Text = "" And mblnClickSquareCtrl = False Then
      MsgBox IIf(mbytInState = 3, "退款金额", "收款金额") & "不能为空,请输入！", vbExclamation, gstrSysName
      txtMoney.SetFocus: Exit Function
  ElseIf CCur(StrToNum(txtMoney.Text)) = 0 And mblnClickSquareCtrl = False Then
      MsgBox IIf(mbytInState = 3, "退款金额", "收款金额") & "不能为零,请输入！", vbExclamation, gstrSysName
      txtMoney.SetFocus: Exit Function
  End If

  If InStr(mstrPrivs, ";负数缴款;") = 0 And StrToNum(txtMoney.Text) < 0 Then
      MsgBox IIf(mbytInState = 3, "退款金额", "收款金额") & "不能为负数,请输入", vbExclamation, gstrSysName
      txtMoney.SetFocus: Exit Function
  End If
  mbytOracleBackType = 1
  
  If mbytInState = 3 Then
        If mbln允许在院病人余额退款 = False And cboType.ItemData(cboType.ListIndex) = 2 Then
            If Val(Nvl(mrsInfo!在院)) = 1 Then
                MsgBox "病人在院,不能进行余额退款,请检查!", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If mdbl剩余款额 - IIf(cboStyle.ItemData(cboStyle.ListIndex) < 0, 0, mdbl预交余额_三方) - CCur(StrToNum(txtMoney.Text)) < 0 Then
            '46067
            If mbytBackMoneyType = 1 Then
                Call MsgBox("退款金额比病人当前的余额多,不能退款!", vbInformation + vbOKOnly, gstrSysName)
                txtMoney.SetFocus: Exit Function
            Else
                If MsgBox("退款金额比病人当前的余额多,忽略吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    txtMoney.SetFocus: Exit Function
                End If
                mbytOracleBackType = 0
            End If
        End If
        '问题号:112995,焦博,2017/10/13,退卡退费时提示病人退费金额
        If mblnOneCard Then
            If mdbl剩余款额 - IIf(cboStyle.ItemData(cboStyle.ListIndex) < 0, 0, mdbl预交余额_三方) - CCur(StrToNum(txtMoney.Text)) > 0 Then
                MsgBox "退款金额比病人当前的余额少,请调整退卡金额!", vbInformation + vbOKOnly, gstrSysName
                txtMoney.SetFocus: Exit Function
            End If
        End If
        If cboStyle.ItemData(cboStyle.ListIndex) < 0 Then
            If mdbl预交余额_三方 - CCur(StrToNum(txtMoney.Text)) < 0 Then
              Call MsgBox("" & cboStyle.Text & "最多只能退" & Format(mdbl预交余额_三方, "###0.00;-###0.00;;") & "!", vbInformation + vbOKOnly, gstrSysName)
              txtMoney.SetFocus: Exit Function
            End If
        End If
        If gbyt预存款消费验卡 <> 0 Then
            If mrsInfo Is Nothing Then
                lng病人ID = Val(txtPatient.Tag)
            ElseIf mrsInfo.State <> 1 Then
                lng病人ID = Val(txtPatient.Tag)
            Else
                lng病人ID = mrsInfo!病人ID
            End If
            If mbln住院退预交验证 Or cboType.ItemData(cboType.ListIndex) = 1 Then
                If Not zldatabase.PatiIdentify(Me, glngSys, lng病人ID, Val(StrToNum(txtMoney.Text)), _
                    , , , , , , , (gbyt预存款消费验卡 = 2)) Then Exit Function
            End If
        End If
        
  Else
        If CCur(StrToNum(txtMoney.Text)) < 0 And Abs(CCur(StrToNum(txtMoney.Text))) > mdbl剩余款额 Then
            '46067
            If mbytBackMoneyType = 1 Then
                    '负数退款,不能大于他本身的余额:37375
                    Call MsgBox("退款金额比病人当前的余额多,不能退款!", vbInformation + vbOKOnly, gstrSysName)
                    txtMoney.SetFocus: Exit Function
            Else
                If MsgBox("退款金额比病人当前的余额多,忽略吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    txtMoney.SetFocus: Exit Function
                End If
                mbytOracleBackType = 0
            End If
        End If
  End If
  
  If cboStyle.ListIndex = -1 Then
      MsgBox "请确定结算方式！", vbExclamation, gstrSysName
      cboStyle.SetFocus: Exit Function
  End If
  
  If cboStyle.ItemData(cboStyle.ListIndex) = BalanceType.C5代收款 Then
      If mbytInState = 3 Then
           If InStr(1, mstrPrivs, ";代收款退款;") = 0 Then
                MsgBox "你没有权限进行代收款退款操作！", vbInformation, gstrSysName
                Exit Function
           End If
      Else
            If InStr(mstrPrivs, ";代收款收取;") = 0 Then
                MsgBox "你没有权限进行代收款收取操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
  ElseIf mbytInState = 3 Then
        '退款操作
        If InStr(1, mstrPrivs, ";预交退款;") = 0 Then
            MsgBox "你没有权限进行预交退款操作！", vbInformation, gstrSysName: Exit Function
        End If
  ElseIf InStr(mstrPrivs, ";预交收款;") = 0 Then
      MsgBox "你没有权限进行预交收款操作！", vbInformation, gstrSysName
      Exit Function
  End If
  '医保改动
  If cboStyle.ItemData(cboStyle.ListIndex) = 3 Then
      If mbytInState = 3 Then
            MsgBox "医保病人个人帐户转帐金额不能进行退款。", vbInformation, gstrSysName
            txtMoney.SetFocus: Exit Function
      Else
            If CCur(StrToNum(txtMoney.Text)) < 0 Then
                MsgBox "医保病人个人帐户转帐金额不能为负。", vbInformation, gstrSysName
                txtMoney.SetFocus: Exit Function
            End If
            If CCur(StrToNum(txtMoney.Text)) > mcur帐户余额 Then
                MsgBox "医保病人个人帐户转帐金额不能超过余额:" & Format(mcur帐户余额, "0.00"), vbInformation, gstrSysName
                txtMoney.SetFocus: Exit Function
            End If
        End If
  End If
  
  If mblnClickSquareCtrl Then
      If CCur(StrToNum(txtMoney.Text)) < 0 Then
          MsgBox "结帐卡转预交金额不能为负。", vbInformation, gstrSysName
          txtMoney.SetFocus: Exit Function
      End If
  End If
  
  '问题号:50656
  If mFactProperty.intInvoicePrint = 0 Then CheckDataValied = True: Exit Function
  '票据号码检查
  If gblnBill预交 Then
      If Trim(txtFact.Text) = "" Then
          MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
          txtFact.SetFocus: Exit Function
      End If
      mlng领用ID = CheckUsedBill(2, IIf(mlng领用ID > 0, mlng领用ID, mFactProperty.lngShareUseID), txtFact.Text, cboType.ItemData(cboType.ListIndex))
      If mlng领用ID <= 0 Then
          Select Case mlng领用ID
              Case 0 '操作失败
              Case -1
                  MsgBox "你没有自用和共用的预交票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
              Case -2
                  MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
              Case -3
                  MsgBox "票据号码不在当前有效领用范围内,请重新输入！", vbInformation, gstrSysName
                  txtFact.SetFocus
          End Select
          txtFact.Text = ""
          Exit Function
      End If
  Else
      If Len(txtFact.Text) <> gbyt预交 And txtFact.Text <> "" Then
          MsgBox "票据号码长度应该为 " & gbyt预交 & " 位！", vbInformation, gstrSysName
          txtFact.SetFocus: Exit Function
      End If
  End If
    CheckDataValied = True
End Function

Private Function Select三方退款() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:余额退款选择
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-21 18:01:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset, strSQL As String
    Dim int预交类型 As Integer, bln三方接口 As Boolean
    Dim strWhere As String, lng病人ID As Long
    Dim vRect  As RECT, blnCancel As Boolean
    
    On Error GoTo errHandle
    If mbytInState = 4 Then Select三方退款 = True: Exit Function
    
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    
    If Me.cboStyle.ItemData(cboStyle.ListIndex) >= 0 Then Select三方退款 = True: Exit Function
    
    mcurBill.lng卡类别ID = -1
    
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    mdbl费用余额 = 0: mdbl预交余额 = 0: mdbl剩余款额 = 0
    int预交类型 = cboType.ItemData(cboType.ListIndex)
    
    '余额退款时,需要检查当前支付的是否充足
    '说明： And nvl(A.交易说明,' ') Like '%%' 这一句的目的是使得记录只有一行时不弹出选择器
    strWhere = IIf(mbln消费卡, " And A.结算卡序号=[3] ", " And nvl(A.卡类别ID,0)=[3]")
    strSQL = _
        "Select a.卡类别id, a.结算卡序号, Min(a.收款时间) As 收款时间, a.卡号, a.交易流水号, a.交易说明," & vbNewLine & _
        "       Max(Decode(Sign(a.金额), -1, 0, Decode(a.记录性质, 11, 0, ID))) As 预交id," & vbNewLine & _
        "       Sum(Nvl(金额, 0)) - Sum(Nvl(冲预交, 0)) As 预交余额" & vbNewLine & _
        "From 病人预交记录 A" & vbNewLine & _
        "Where a.病人id = [1] And a.预交类别 = [2] And a.记录性质 In (1, 11) And nvl(A.交易说明,' ') Like '%%'" & strWhere & vbNewLine & _
        "Group By a.卡类别id, a.结算卡序号, a.卡号, a.交易流水号, a.交易说明"
    '108376,焦博 2017,06/13 列表上加一列"交易日期"(预交款的收款时间)
    strSQL = _
        "Select Distinct a.预交id, a.卡类别id, a.结算卡序号 As 消费接口id, Nvl(b.编码, c.编号) As 编码," & vbNewLine & _
        "       Nvl(b.名称, c.名称) As 名称, a.卡号, a.交易流水号, a.交易说明, Nvl(b.是否转帐及代扣, 0) As 转帐," & vbNewLine & _
        "       a.预交余额, a.收款时间 As 交易日期, Nvl(b.是否退款验卡, 0) As 是否退款验卡, n.消费卡id" & vbNewLine & _
        "From (" & strSQL & ") A, 医疗卡类别 B, 消费卡类别目录 C, 病人卡结算记录 N" & vbNewLine & _
        "Where a.卡类别id = b.Id(+) And a.结算卡序号 = c.编号(+) And a.预交id = n.结算id(+)" & vbNewLine & _
        "      And Nvl(a.预交余额, 0) > 0" & vbNewLine & _
        "      And Not Exists (Select 1 From 消费卡信息 Where 接口编号 = a.结算卡序号 And 卡号 = a.卡号 And Nvl(当前状态, 1) <> 1" & vbNewLine & _
        "           And 序号 = (Select Max(序号) From 消费卡信息 Where 接口编号 = a.结算卡序号 And 卡号 = a.卡号))" & vbNewLine & _
        "Order By 卡号"
    
    strSQL = _
        "Select Rownum As ID, 预交id, 卡类别id, 消费接口id, 编码, 名称, 卡号, 交易流水号, 交易说明, " & vbNewLine & _
        "       预交余额, 交易日期, 转帐 As 转帐_ID, 是否退款验卡 As 是否退款验卡_ID, 消费卡id" & vbNewLine & _
        "From (" & strSQL & ")"
    Set rsMoney = zldatabase.ShowSQLSelect(Me, strSQL, 0, cboStyle.Text & "退款", False, "", "请选择需要退款的交易", _
        False, False, False, vRect.Left, vRect.Top, cboStyle.Height, blnCancel, True, True, lng病人ID, int预交类型, _
        mlngCardTypeID)
    If blnCancel Then Exit Function
    If rsMoney Is Nothing Then
        MsgBox cboStyle.Text & "不存在可退余额,不能退款!", vbOKOnly + vbInformation, gstrSysName
        txtMoney.Text = "0.00"
        Exit Function
    End If
    
    With rsMoney
        mcurBill.lng预交ID = Val(Nvl(!预交ID))
        mcurBill.bln消费卡 = mbln消费卡
        mcurBill.lng卡类别ID = mlngCardTypeID
        mcurBill.str交易流水号 = Nvl(!交易流水号)
        mcurBill.str交易说明 = Nvl(!交易说明)
        mcurBill.str卡号 = Nvl(!卡号)
        mcurBill.bln转账 = Val(Nvl(!转帐_ID)) = 1
        mcurBill.bln退款验卡 = Val(Nvl(!是否退款验卡_ID)) = 1
        mcurBill.dbl金额 = Val(Val(Nvl(rsMoney!预交余额)))
        mcurBill.lng消费卡ID = Val(Val(Nvl(rsMoney!消费卡ID)))
        txtMoney.Text = Format(Val(Nvl(rsMoney!预交余额)), "#,###0.00;-#,###0.00;;")
        mdbl预交余额_三方 = Val(Val(Nvl(rsMoney!预交余额)))
    End With
    Select三方退款 = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function
Private Function CheckBrushCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset
    Dim lng病人ID As String
    Dim dblMoney As Double
    Dim strExpand As String '问题号:55666
    Dim dbl账户余额 As Double '问题号:55666
    Dim bln消费卡刷卡 As Boolean '问题号:55666
    Dim bln三方卡刷卡 As Boolean '问题号:55666
    
    On Error GoTo errHandle
    dblMoney = IIf(mbytInState = 3, -1, 1) * StrToNum(txtMoney.Text)
    If cboStyle.ItemData(cboStyle.ListIndex) >= 0 Then CheckBrushCard = True: Exit Function
    If mbytInState = 3 Then
        If mcurBill.lng卡类别ID < 0 Then
            If Select三方退款 = False Then
                Exit Function
            End If
        End If
        dblMoney = StrToNum(txtMoney.Text)
         If zlCheckDepositDelValied(mcurBill.lng预交ID, dblMoney) = False Then Exit Function
         CheckBrushCard = True: Exit Function
    End If
     '弹出刷卡界面
    'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByVal dbl金额 As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln退费 As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln退现 As Boolean = False, _
        Optional ByVal bln余额不足禁止 As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal bln转预交 As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "", _
        Optional ByVal str费用来源 As String, _
        Optional ByVal lng病人ID As Long) As Boolean
    '       strXmlIn-XML入参,目前格式如下:
    '       <IN>
    '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
    '       </IN>
    '       str费用来源 - 当前支付费用的费用来源，多种用逗号分隔(使用消费卡支付时传入)
    '       lng病人ID - 病人ID(使用消费卡支付时传入)
    '问题号:55666
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = adStateOpen Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, mlngCardTypeID, mbln消费卡, _
        Nvl(mrsInfo!姓名), Nvl(mrsInfo!性别), Nvl(mrsInfo!年龄), dblMoney, mstrBrushCardNo, mstrbrPassWord, _
        False, True, False, False, Nothing, False, False, "<IN><CZLX>0</CZLX></IN>", _
        cboType.ItemData(cboType.ListIndex), lng病人ID) = False Then Exit Function
    '保存前,一些数据检查
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, mlngCardTypeID, _
        mbln消费卡, mstrBrushCardNo, dblMoney, "", "") = False Then Exit Function
    '问题号:55666,55851
    gobjSquare.objSquareCard.zlGetAccountMoney Me, mlngModul, mlngCardTypeID, mstrBrushCardNo, strExpand, dbl账户余额, mbln消费卡
    If dbl账户余额 <> 0 Then sta.Panels(2).Text = "账户余额:" & dbl账户余额
    '判断预交金是否超出刷卡的余额
    lblRepairMoney.Visible = CDbl(txtMoney.Text) > dblMoney
    If lblRepairMoney.Visible Then
        lblRepairMoney.Caption = "补交额:" & Format((CDbl(txtMoney.Text) - dblMoney), "###0.00;-###0.00;;")
        txtMoney.Text = Format(dblMoney, "###0.00;-###0.00;;")
    End If
    CheckBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckChangDepositType() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查转预交的类型
    '返回:数据合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 09:54:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    If mrsInfo.State = adStateClosed Then
        MsgBox "没有确定转预交款的病人,不能转预交！", vbExclamation, gstrSysName
       If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus: Exit Function
    End If
    If LenB(StrConv(cboNote.Text, vbFromUnicode)) > 50 Then
        MsgBox "缴款摘要只允许 50 个字符或 25 个汉字,请修改！", vbInformation, App.Title
        If cboNote.Enabled And cboNote.Visible Then cboNote.SetFocus: Exit Function
    End If
    If txtMoney.Text = "" Then
        MsgBox "转预交额不能为空,请输入！", vbExclamation, gstrSysName
        If txtMoney.Enabled And txtMoney.Visible Then txtMoney.SetFocus: Exit Function
    ElseIf CCur(StrToNum(txtMoney.Text)) = 0 Then
        MsgBox "转预交额不能为零,请输入！", vbExclamation, gstrSysName
        If txtMoney.Enabled And txtMoney.Visible Then txtMoney.SetFocus: Exit Function
    End If
    
    If StrToNum(txtMoney.Text) < 0 Then
        MsgBox "转预交额不能为负数,请输入", vbExclamation, gstrSysName
        If txtMoney.Enabled And txtMoney.Visible Then txtMoney.SetFocus: Exit Function
    End If
  
    If mdbl剩余款额 - CCur(StrToNum(txtMoney.Text)) < 0 Then
        Call MsgBox("转预交余额比病人当前的余额多,不能退款!", vbInformation + vbOKOnly, gstrSysName)
        If txtMoney.Enabled And txtMoney.Visible Then txtMoney.SetFocus: Exit Function
    End If
    
    '112999
    If cboType.ListIndex < 0 Then Exit Function
    If cboType.ItemData(cboType.ListIndex) = 1 Then
        If Val(Nvl(mrsInfo!在院)) = 0 And gblnAllowOut = False Then
            strSQL = "Select 1 From 病案主页 Where 病人ID=[1] And Nvl(主页ID,0)=0 And Nvl(病人性质,0)=0" '预入院
            Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsInfo!病人ID)))
            If rsTemp.EOF Then
                MsgBox "病人还未住院，不能门诊预交转住院！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Else
        If Val(Nvl(mrsInfo!在院)) = 1 And gblnBanIn = True Then
            MsgBox "病人还未出院，不能住院预交转门诊！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '票据号码检查
    If gblnBill预交 Then
        If Trim(txtFact.Text) = "" Then
            MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
            txtFact.SetFocus: Exit Function
        End If
        If mbytInState = 4 Then
            mlng领用ID = CheckUsedBill(2, IIf(mlng领用ID > 0, mlng领用ID, mFactProperty.lngShareUseID), txtFact.Text, IIf(cboType.ItemData(cboType.ListIndex) = 1, 2, 1))
        Else
            mlng领用ID = CheckUsedBill(2, IIf(mlng领用ID > 0, mlng领用ID, mFactProperty.lngShareUseID), txtFact.Text, cboType.ItemData(cboType.ListIndex))
        End If
        If mlng领用ID <= 0 Then
            Select Case mlng领用ID
                Case 0 '操作失败
                Case -1
                    MsgBox "你没有自用和共用的预交票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Case -3
                    MsgBox "票据号码不在当前有效领用范围内,请重新输入！", vbInformation, gstrSysName
                    txtFact.SetFocus
            End Select
            txtFact.Text = ""
            Exit Function
        End If
    Else
        If Len(txtFact.Text) <> gbyt预交 And txtFact.Text <> "" Then
            MsgBox "票据号码长度应该为 " & gbyt预交 & " 位！", vbInformation, gstrSysName
            txtFact.SetFocus: Exit Function
        End If
    End If
    CheckChangDepositType = True
End Function
Private Function SaveChageDepositType() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存预交类型转换
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 10:02:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String

    On Error GoTo errHandle
    mstrPrintDate = Format(zldatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    'Zl_病人预交记录_转预交
    strSQL = "Zl_病人预交记录_转预交("
    '  票据号_In     票据使用明细.号码%Type,
    strSQL = strSQL & "'" & txtFact.Text & "',"
    '  病人id_In     病人预交记录.病人id%Type,
    strSQL = strSQL & "" & Val(Nvl(mrsInfo!病人ID)) & ","
    '  主页id_In     病人预交记录.主页id%Type,
    strSQL = strSQL & "" & IIf(mrsInfo!当前科室id <> 0 And mrsInfo!在院 <> 0, mrsInfo!主页ID, "Null") & ","
    '  科室id_In     病人预交记录.科室id%Type,
    strSQL = strSQL & "" & IIf(cboUnit.ItemData(cboUnit.ListIndex) = 0, "NULL", cboUnit.ItemData(cboUnit.ListIndex)) & ","
    '  金额_In       病人预交记录.金额%Type,
    strSQL = strSQL & "" & StrToNum(txtMoney.Text) & ","
    '  操作员编号_In 病人预交记录.操作员编号%Type,
   strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人预交记录.操作员姓名%Type,
   strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  收款时间_In   病人预交记录.收款时间%Type,
   strSQL = strSQL & "to_Date('" & mstrPrintDate & "','yyyy-mm-dd hh24:mi:ss'),"
    '  领用id_In     票据使用明细.领用id%Type,
    strSQL = strSQL & "" & IIf(mlng领用ID = 0, "NULL", mlng领用ID) & ","
    '  预交类别_In   病人预交记录.预交类别%Type,
    strSQL = strSQL & "" & cboType.ItemData(cboType.ListIndex) & ","
    '  摘要_In       病人预交记录.摘要%Type
   strSQL = strSQL & "'" & cboNote.Text & "')"
    zldatabase.ExecuteProcedure strSQL, Me.Caption
    SaveChageDepositType = True
    
    If Not gblnBill预交 And Trim(txtFact.Text) <> "" Then
        '松散：保存当前号码
        zldatabase.SetPara "当前预交票据号", Trim(txtFact.Text), glngSys, mlngFactModule
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    Dim i As Integer, blnCardEnable As Boolean, lng接口编号 As Long, strBlanceInfor As String
    Dim varData As Variant, blnHave结算方式 As Boolean
    Dim blnCanDel As Boolean, intInsure As Integer
    Dim msgBoxResult As VbMsgBoxResult '问题号:50656
    Dim bln打印 As Boolean  '问题号:57624
    Dim lng预交ID As Long
    
    If chkCancel.Value = Checked Then
         Call zlBackDeposit: Exit Sub
    End If
    
    '问题号:57624
    '问题号:50565
    Select Case mFactProperty.intInvoicePrint
    Case 0 '不打印预交发票
       bln打印 = False
    Case 1 '自动打印
       bln打印 = True
    Case 2 '打印提醒
        msgBoxResult = MsgBox("是否需要打印预交票据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        bln打印 = (msgBoxResult = vbYes)
    End Select
    
    If mbytInState = 4 Then
        '--门诊转住院或住院转门诊
        If CheckChangDepositType = False Then Exit Sub
        If SaveChageDepositType = False Then Exit Sub
        If bln打印 Then    '120271
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "收款时间=" & mstrPrintDate, "NO='无' ", "病人ID=" & mrsInfo!病人ID, "ReportFormat=2", 2)
        End If
        mblnOK = True: Unload Me: Exit Sub
    End If
    
    If Not Check未入科不交预交 Then Exit Sub
    If CheckDataValied = False Then Exit Sub
    If CheckBrushCard = False Then Exit Sub
    '存盘
    cmdOK.Enabled = False
    
    If Check退款 = False Then cmdOK.Enabled = True: Exit Sub
    '中间不能有弹出类，避免长时间挂起造成并发
    If Not SaveBill(bln打印, lng预交ID) Then
        MsgBox "预交款单据保存失败,请重试该操作。如果仍有问题,请与系统管理员联系！", vbExclamation, gstrSysName
        cmdOK.Enabled = True: Exit Sub
    Else
        '问题号:57624
        '问题号:50656
        If bln打印 Then '票据号为空就表示不打印发票
            '78751:李南春,2014/10/20,增加预交票据打印格式
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & cboNO.List(0), "病人ID=" & mrsInfo!病人ID, "收款时间=" & Format(Now, "yyyy-mm-dd HH:MM:SS"), _
                            IIf(mFactProperty.intInvoiceFormat = 0, "", "ReportFormat=" & mFactProperty.intInvoiceFormat), 2)
            Call zlCheckFactIsEnough
        End If
        
        '81693:李南春,2015/4/21,评价器
        If Not mobjPlugIn Is Nothing Then
            On Error Resume Next
            Call mobjPlugIn.PatiPrePayAfter(mrsInfo!病人ID, IIf(mbytPrepayType = 2, 1, 0), lng预交ID)
            Err.Clear
        End If
    End If
    '问题号:55666
    '存在补交金额的情况
    If UBound(Split(lblRepairMoney.Caption, ":")) = 1 And Split(lblRepairMoney.Caption, ":")(1) <> "" Then
        txtPatient.Tag = ""
        txtMoney.Text = Split(lblRepairMoney.Caption, ":")(1)
        IDKind.IDKind = IDKind.GetKindIndex("姓名")
        txtPatient.Text = "-" & mrsInfo!病人ID
        txtPatient_KeyPress 13
        '刷新票据
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact
        '定位支付方式
        '84751:李南春,2015/5/14,下拉框定位越界
        For i = 0 To cboStyle.ListCount - 1
            If cboStyle.List(i) = mstr缺省结算方式 Then
                cboStyle.ListIndex = i
            End If
        Next
    
        lblRepairMoney.Visible = False: lblRepairMoney.Caption = "补交额:"
        cmdOK.Enabled = True
        Exit Sub
    End If
    
    '问题:48249
    If mbytCallObject = 1 Or mbytCallObject = 2 Then
        '费用查询时,直接退出
        mblnOK = True: Unload Me: Exit Sub
    End If
    
    If mblnClearWinInfor Then
        Call ClearBill
        Call InitFace(True)
        Call cboStyle_Click
    Else
        '问题号:44732
        SetMoneyInfo False, , , True
        Set mrsInfo = New ADODB.Recordset
        
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact  '重新获取发票号
    End If
    Call SetcmdOkEnabled
    If txtPatient.Enabled Then txtPatient.SetFocus
    mblnOK = True
End Sub

Private Sub ClearBill()
'功能:清除相关界面和数据
    If (mbytInState = 0 Or mbytInState = 3) And gblnLED Then
        zl9LedVoice.DisplayPatient ""
    End If
    
    Set mrsInfo = New ADODB.Recordset '清除病人信息
    txtPatient.Text = "": txtPatient.Locked = False
    txtPatient.Tag = ""
    cboUnit.ListIndex = 0
    txtUnit.Tag = ""
    txtUnit.Text = ""
    mstr退款操作员 = ""
    
    txt开户行.Text = ""
    txt帐号.Text = ""
    SetMoneyInfo True
    
    txtMoney.Text = ""
    If cboStyle.ListCount <> 0 And cboStyle.Tag <> "" Then cboStyle.ListIndex = Val(cboStyle.Tag) '恢复缺省结算方式
    txtCode.Text = "": txtCode.Locked = False
    
    txtMan.Text = UserInfo.姓名
    txtDate.Text = Format(zldatabase.Currentdate(), "yyyy-MM-dd")
    cboNote.Text = ""
    
    '医保改动
    Call Clear个人帐户
    
    '新的一张预交款单据
    cboNO.Text = "": cboNO.Locked = True
    
    txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "修改票据号") And gblnBill预交 '89302
    If mFactProperty.intInvoicePrint <> 0 Then Call GetFact
    txtPatient.SetFocus
End Sub

Private Sub cmdSetup_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub chkAllCash_Click()
    Dim dblMoney As Double, lngRow As Long
    Dim strDBUser As String
    Dim strPrivs As String
    
    If chkAllCash.Value = 1 Then
        If mstr退款操作员 <> "" Then
            mdbl预交余额_三方 = 0
            GoTo ResetMoney
        End If
        If InStr(";" & mstrCardPrivs & ";", ";三方退款强制退现;") = 0 Then
            mstr退款操作员 = zldatabase.UserIdentifyByUser(Me, "强制退现验证", glngSys, 1151, "三方退款强制退现")
            If mstr退款操作员 = "" Then
                MsgBox "录入的操作员验证失败或者录入的操作员不具备强制退现权限，不能强制退现！", vbInformation, gstrSysName
                chkAllCash.Value = 0
                GoTo ResetMoney
            End If
        
            mdbl预交余额_三方 = 0
        Else
            If MsgBox("存在不支持退现的三方卡,是否允许强制将其退现？", _
                                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then chkAllCash.Value = 0: GoTo ResetMoney
            mstr退款操作员 = UserInfo.姓名
            mdbl预交余额_三方 = 0
        End If
    Else
        mdbl预交余额_三方 = mdbl预交余额_三方备份
    End If
ResetMoney:
    If mdbl剩余款额 - mdbl预交余额_三方 > 0 Then
        txtMoney.Text = Format(mdbl剩余款额 - mdbl预交余额_三方, "#,##0.00;-#,##0.00;;")
    Else
        txtMoney.Text = ""
    End If
End Sub

Private Sub Form_Activate()
    If mblnUnLoad Then Unload Me: Exit Sub
    If mbytInState = 0 Or mbytInState = 3 Or mbytInState = 4 Then
        ' mbytInState=3:表示余额退款,4-表示转预交
        If mlng病人ID <> 0 And Trim(txtPatient.Text) = "" Then
            txtPatient.Text = "-" & mlng病人ID
            Call txtPatient_KeyPress(13)
            If mdblDefPreMoney <> 0 And StrToNum(txtMoney.Text) = 0 Then
                txtMoney.Text = Format(mdblDefPreMoney, "###0.00;-###0.00;;")
            End If
        End If
        If gblnLED And Trim(txtPatient.Text) = "" Then
            zl9LedVoice.DisplayPatient ""    '双屏显示窗体必须在当前窗口显示之后调用显示才能移动窗体
        End If
    ElseIf mbytInState = 1 Then
        cmdCancel.SetFocus
    ElseIf mbytInState = 2 Then
        If mstrInNO = "" Then
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        Else
           If cmdOK.Enabled Then cmdOK.SetFocus
        End If
    End If
    '问题号:45666
    If mbytInState = 0 And cboType.Text = "住院预交" Then '交预交款
        chk仅显示本次预交.Visible = True
        chk仅显示本次预交.Value = IIf(zldatabase.GetPara("仅显示本次预交", glngSys, mlngModul, , Array(chk仅显示本次预交), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If cboStyle.ListIndex >= 0 Then
                If cmdOK.Enabled And cmdOK.Visible Then cmdOK_Click
            Else
                If cmdOK.Enabled And cmdOK.Visible Then cmdOK_Click
            End If
        Case vbKeyF3
            If txtFact.Visible And txtFact.Enabled Then txtFact.SetFocus
        Case vbKeyF4
            If Shift = vbCtrlMask And IDKind.Enabled Then
                Dim intIndex As Integer
                intIndex = IDKind.GetKindIndex("IC卡号")
                If intIndex <= 0 Then Exit Sub
                 IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
            End If
            
        Case vbKeyF11
            If txtPatient.Enabled And picFace.Enabled And Not txtPatient.Locked Then txtPatient.SetFocus
        Case vbKeyF12
            If Not cboNO.Locked And picNO.Enabled Then cboNO.SetFocus
        Case vbKeyF10
            If cmdSetup.Enabled And cmdSetup.Visible Then cmdSetup_Click
        Case vbKeyF8
            If chkCancel.Visible And picNO.Enabled Then chkCancel.Value = IIf(chkCancel.Value = 1, 0, 1)
        Case vbKeyEscape
            Call cmdCancel_Click
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub GetFact(Optional blnFirst As Boolean = False, Optional blnRed As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取不同类别的发票
    '编制:刘兴洪
    '日期:2011-07-19 17:47:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
   '票据领用检查及初始
    If gblnBill预交 Then
        mlng领用ID = CheckUsedBill(2, IIf(mlng领用ID > 0, mlng领用ID, mFactProperty.lngShareUseID), "", mFactProperty.strUseType)
        If mlng领用ID <= 0 Then
            Select Case mlng领用ID
                Case 0 '操作失败
                Case -1
                    MsgBox "你没有自用和共用的预交票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End Select
            If blnFirst Then mblnUnLoad = True: Exit Sub
            
        End If
        '严格：取下一个号码
        If Not blnRed Then
            txtFact.Text = GetNextBill(mlng领用ID)
        Else
            mstrRedFact = GetNextBill(mlng领用ID)
        End If
    Else
        '松散：取下一个号码
        If Not blnRed Then
            txtFact.Text = zlCommFun.IncStr(UCase(zldatabase.GetPara("当前预交票据号", glngSys, mlngFactModule, "")))
        Else
            mstrRedFact = zlCommFun.IncStr(UCase(zldatabase.GetPara("当前预交票据号", glngSys, mlngFactModule, "")))
        End If
    End If
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化模块参数
    '编制:刘兴洪
    '日期:2012-02-27 11:23:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstr缺省结算方式 = zldatabase.GetPara("缺省预交结算方式", glngSys, mlngModul)
    mbytBackMoneyType = Val(zldatabase.GetPara("退款禁止方式", glngSys, mlngModul))
    '结算方式:金额|结算方式:金额....
    mstr代收款 = zldatabase.GetPara("代收款设置", glngSys, mlngModul)
    mblnClearWinInfor = IIf(zldatabase.GetPara("缴预交后不清除信息", glngSys, glngModul) <> "1", True, False)
    mbln未入科不交预交 = zldatabase.GetPara("病人未入科不准收预交", glngSys, mlngModul, , , InStr(mstrPrivs, ";参数设置;") > 0) = "1"
    gblnSeekName = Nvl(zldatabase.GetPara("姓名模糊查找", glngSys, mlngModul, 1)) = 1
    mbln住院退预交验证 = zldatabase.GetPara("住院退预交验证", glngSys, mlngModul, "0") = "1"
    mbln允许在院病人余额退款 = zldatabase.GetPara("允许在院病人余额退款", glngSys, mlngModul, "1") = "1"
    '刷卡要求输入密码
    mblnCheckPass = Mid(zldatabase.GetPara(46, glngSys, , "0000000000"), 8, 1) = "1"
End Sub
Private Sub Form_Load()
    Dim lngH As Long
    
    Call InitPara
    mblnOK = False: mblnUnLoad = False
    
    '票据领用检查及初始
    If mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3 Then
        mblnStartFactUseType = zlStartFactUseType(2)
        If mblnStartFactUseType = False Then
            If mFactProperty.intInvoicePrint <> 0 Then Call GetFact(True, mbytInState = 2)
        End If
    End If
    
    zlControl.PicShowFlat picInfo, -1
    zlControl.PicShowFlat picFace, -1
    
    Set mrsInfo = New ADODB.Recordset
    
    If Not InitUnit Then Unload Me: Exit Sub

    Call InitIDKind
    
    Call InitFace
    If mblnUnLoad Then Exit Sub
    
    lblTitle.Caption = gstrUnitName & "预交款单据"
    mstrCardPrivs = GetPrivFunc(glngSys, 1151)
    
    If (mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3) And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.编号 & "号为您服务", mlngModul, gcnOracle
        
        Call zlCheckFactIsEnough
    End If
    If mbytInState = 0 Or mbytInState = 3 Then
        IDKind.IDKind = Val(zldatabase.GetPara("上次输入方式", glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0))
    End If
    
    '81693:李南春,2015/4/21,评价器
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear: On Error GoTo 0
    End If
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInState = 0: mstrInNO = ""
    mblnViewCancel = False: mblnUnLoad = False
    mlng领用ID = 0: mstr个人帐户 = "": mblnNOMoved = False
    mstr退款操作员 = ""
    
    If (mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3) And gblnLED Then
        zl9LedVoice.DisplayPatient "": zl9LedVoice.Reset com
    End If
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
    Call mobjICCard.SetEnabled(False)
    Set mobjICCard = Nothing
    End If
    Set mobjPlugIn = Nothing
'    Call zlCardSquareObject(True)
    Call SaveWinState(Me, App.ProductName)
    If mbytInState = 0 Or mbytInState = 3 Then
        zldatabase.SetPara "上次输入方式", IDKind.IDKind, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    End If
    '问题号:45666
    If mbytInState = 0 And cboType.Text = "住院预交" Then
        zldatabase.SetPara "仅显示本次预交", chk仅显示本次预交.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    End If
End Sub

Private Sub InitPrepayType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化预交类型
    '编制:刘兴洪
    '日期:2011-07-14 18:50:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytInState = 4 Then
        With cboType
            .Clear
            If InStr(1, mstrPrivs, ";门诊预交转住院;") > 0 Then
                .AddItem "门诊转住院": .ItemData(.NewIndex) = 1
                If mbytPrepayType = 1 Then .ListIndex = .NewIndex
            End If
            If InStr(1, mstrPrivs, ";住院预交转门诊;") > 0 Then
                .AddItem "住院转门诊": .ItemData(.NewIndex) = 2
                If mbytPrepayType = 2 Then .ListIndex = .NewIndex
            End If
            
        End With
        lbl预交类型.Caption = "转预交"
        If cboType.ListCount = 0 Then
            MsgBox "你不具备门诊预交转住院或住院预交转门诊权限，请与系统管理员联系!", vbInformation + vbOKOnly, gstrSysName
            mblnUnLoad = True
        End If
        
        Exit Sub
    End If
    With cboType
        .Clear
        If InStr(1, mstrPrivs, ";门诊预交;") > 0 Then
            .AddItem "门诊预交": .ItemData(.NewIndex) = 1
            If mbytPrepayType = 1 Then .ListIndex = .NewIndex
        End If
        If InStr(1, mstrPrivs, ";住院预交;") > 0 Then
            .AddItem "住院预交": .ItemData(.NewIndex) = 2
            If mbytPrepayType = 2 Then .ListIndex = .NewIndex
        End If
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
        If cboType.ListCount = 0 Then
            MsgBox "你不具备门诊预交或住院预交权限，请与系统管理员联系!", vbInformation + vbOKOnly, gstrSysName
            mblnUnLoad = True
        End If
        
     End With
End Sub

Private Sub InitFace(Optional blnSave As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据入口参数设置窗体界面及控制状态
    '编制:刘兴洪
    '日期:2011-07-17 10:36:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, i As Integer, strSQL As String
    Dim ctlTmp As Control
    
    If Not gobjSquare.objSquareCard Is Nothing And blnSave = False Then
        IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
    End If
    
    On Error GoTo errHandle
    
    strSQL = "Select 编码, 名称, 简码, 缺省标志 From 常用预交摘要 Order by 编码"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboNote.Clear
    If rsTmp.RecordCount > 0 Then
        While Not rsTmp.EOF
            cboNote.AddItem Nvl(rsTmp!名称)
            rsTmp.MoveNext
        Wend
    End If
    
    cboNote.ListIndex = -1: Call InitPrepayType
    If mblnUnLoad Then Exit Sub
    
    IDKind.Enabled = (mbytInState = 0 Or mbytInState = 3 Or mbytInState = 4)
    Select Case mbytInState
        Case 0, 3 '收取预交款,余额退款
            '创建卡部件
            Call CreateMobjCard
            cboNO.Text = ""
            txtDate.Text = Format(zldatabase.Currentdate(), "yyyy-MM-dd")
            txtMan.Text = UserInfo.姓名
            
            Call Load支付方式
            '退款权限
            If InStr(mstrPrivs, ";预交退款;") = 0 And InStr(mstrPrivs, ";代收款退款;") = 0 Or mbytInState = 3 Or mbytInState = 4 Then
                chkCancel.Visible = False
            End If
            '只有代收款收取权限
            If InStr(mstrPrivs, ";预交收款;") = 0 Then Call cbo.Locate(cboStyle, 5, True)
            If mbytInState = 3 Then
                lblMoney.Caption = "退款金额": lblMoney.FontBold = True: lblMoney.ForeColor = vbRed
                txtMoney.ForeColor = vbRed: txtMoney.Font.Bold = True
            End If
            txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "修改票据号") And gblnBill预交 '89302
        Case 1 '指定单据浏览
            picList.Visible = False
            Me.Height = Me.Height - picList.Height
            
            If mblnViewCancel Then lblFlag.Visible = True
            cmdSetup.Visible = False
            chkCancel.Visible = False
            cmdOK.Visible = False
            
            cmdCancel.Caption = "退出(&X)"
            
            picNO.Enabled = False
            picFace.Enabled = False
            cboNote.Locked = True
            txtFact.Locked = True
            txtUnit.Locked = True
            txt开户行.Locked = True
            txt帐号.Locked = True
            
            '显示单据内容
            If Not ReadBill(mstrInNO) Then
                MsgBox "不能正确读取该单据内容，请与系统管理员联系！", vbExclamation, gstrSysName
                mblnUnLoad = True
            End If
        Case 2 '指定单据退款
            
            chkCancel.Value = Checked   '在调用的click事件中处理 picFace.Enabled = True '！！！不允许部份退款！！！
            cmdSetup.Visible = False
            txtFact.Locked = True
            If mstrInNO <> "" Then  '病人信息管理中退预交,没有指定单据号
                picNO.Enabled = False
                '显示单据内容
                Dim intBill As Integer
                intBill = ReadBill(mstrInNO)
                If intBill <> -1 Then
                    If intBill <> 3 Then
                        MsgBox "不能正确读取该单据内容，请与系统管理员联系！", vbExclamation, gstrSysName
                    End If
                    mblnUnLoad = True
                End If
            End If
        Case 4
            '创建卡部件
            Call CreateMobjCard
            chkCancel.Visible = False
            txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "修改票据号") And gblnBill预交 '89302
    End Select
    
    If lbl帐户余额.Visible = False Then lbl预交余额.Left = lbl帐户余额.Left
    If lbl帐户余额.Visible Then
        Line2(14).Visible = True: Line2(11).X2 = 2415
    Else
        Line2(14).Visible = False: Line2(11).X2 = Line2(14).X2
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub CreateMobjCard()
    '创建卡部件
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hWnd)
    Set mobjICCard.gcnOracle = gcnOracle
End Sub
Private Sub txtCode_GotFocus()
    txtCode.SelStart = 0: txtCode.SelLength = Len(txtCode.Text)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    Else
        CheckInputLen txtCode, KeyAscii
    End If
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0: txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If IsDate(txtDate.Text) And KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtFact_LostFocus()
'    If Not txtFact.Locked And txtFact.Text <> "" Then
'        txtFact.Text = Format(Left(txtFact.Text, gbyt预交), String(gbyt预交, "0"))
'    End If
End Sub

Private Sub txtMoney_Change()
    '问题27363
    If IsNumeric(StrToNum(txtMoney.Text)) Then
        If mbytInState = 3 Then
            txtMoney.ForeColor = vbRed
        Else
            txtMoney.ForeColor = IIf(CCur(StrToNum(txtMoney.Text)) >= 0, vbBlue, vbRed)
        End If
    End If
End Sub

Private Sub txtMoney_GotFocus()
    txtMoney.SelStart = 0: txtMoney.SelLength = Len(txtMoney.Text)
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    '问题27363
    If KeyAscii <> 13 Then
        If chkCancel.Value = Checked Or mbytInState = 3 Or mbytInState = 4 Then
            '退款时不允许输入负数
            If KeyAscii = Asc(".") And InStr(txtMoney.Text, ".") > 0 Then KeyAscii = 0: Beep: Exit Sub
            If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        Else
            '收款时可以通过负数退款
            If KeyAscii = Asc(".") And InStr(txtMoney.Text, ".") > 0 Then KeyAscii = 0: Beep: Exit Sub
            '权限设置
            If InStr(mstrPrivs, ";预交退款;") = 0 Then
                If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            Else
                If (txtMoney.Text <> "" And txtMoney.SelLength <> Len(txtMoney.Text)) And KeyAscii = Asc("-") Then KeyAscii = 0: Beep: Exit Sub
                If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            End If
        End If
        If (txtMoney.Text <> "" And txtMoney.SelLength <> Len(Format(StrToNum(txtMoney.Text), "##,##0.00;-##,##0.00; ;"))) And _
            (Len(Format(StrToNum(txtMoney.Text), "##,##0.00;-##,##0.00; ;")) >= txtMoney.MaxLength) And _
            InStr(Chr(8), Chr(KeyAscii)) = 0 Then
            If txtMoney.SelLength > 0 And txtMoney.SelLength <= txtMoney.MaxLength Then
            Else
                KeyAscii = 0: Beep: Exit Sub
            End If
        End If
    Else
        If txtMoney.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtMoney_LostFocus()
    '问题27363
    Dim dblMoney  As Double
    If Not IsNumeric(StrToNum(txtMoney.Text)) Then txtMoney.SetFocus: Exit Sub
    If mrsInfo.State = 1 And IsNumeric(StrToNum(txtMoney.Text)) Then
        txtMoney.Text = Format(StrToNum(txtMoney.Text), "##,##0.00;-##,##0.00; ;")
        If txtMoney.MaxLength > 12 Then txtMoney.MaxLength = 12
        '108813:李南春,2017/5/8,语音播报控制
        If mbytInState = 4 Then Exit Sub
        If gblnLED Then
            '#22 1234.56   --预收一千二百三十四点五六元 Y
            '#23 1234.56   --找零一千二百三十四点五六元 Z
            dblMoney = StrToNum(txtMoney.Text)
            If mbytInState = 3 Then dblMoney = -1 * dblMoney
            zl9LedVoice.Speak "#22 " & dblMoney
        End If
    End If
End Sub

Private Sub cboNO_GotFocus()
    If Not cboNO.Locked Then cboNO.SelStart = 0: cboNO.SelLength = Len(cboNO.Text)
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim strOper As String, vDate As Date
    
    If cboNO.Locked Then Exit Sub
    
    '转换成大写(汉字不可处理)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    ElseIf cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 11)
        
        '是否已转入后备数据表中,记录性质为1表示交或冲预交
        If zldatabase.NOMoved("病人预交记录", cboNO.Text, , "1", Me.Caption) Then
            If Not ReturnMovedExes(cboNO.Text, 6, Me.Caption) Then Exit Sub
            mblnNOMoved = False
        End If
        
        '单据权限
        If Not ReadBillInfo(0, cboNO.Text, -2, strOper, vDate) Then
            cboNO.Text = "": cboNO.SetFocus: Exit Sub
        End If
        If Not BillOperCheck(6, strOper, vDate, "退款") Then
            cboNO.Text = "": cboNO.SetFocus: Exit Sub
        End If
        '问题27363
        '读取要作废的预交款单据
        Select Case ReadBill(cboNO.Text)
            Case -1
                If cboStyle.ItemData(cboStyle.ListIndex) = BalanceType.C5代收款 Then
                    If InStr(mstrPrivs, ";代收款退款;") = 0 Then
                        MsgBox "你没有权限进行代收款退款操作！", vbInformation, gstrSysName
                        chkCancel.Value = 0
                    End If
                ElseIf InStr(mstrPrivs, ";预交退款;") = 0 Then
                    MsgBox "你没有权限进行预交退款操作！", vbInformation, gstrSysName
                    chkCancel.Value = 0
                Else
                    If HaveSpare(cboNO.Text) = 0 And InStr(mstrPrivs, ";预交结清退款;") = 0 Then
                        MsgBox "该病人已没有预交余额,你没有权限作废这张单据！", vbInformation, gstrSysName
                        chkCancel.Value = 0
                    ElseIf HaveBalance(cboNO.Text) <> 0 Then
                        MsgBox "该笔预交已经被病人在结帐时使用,你不能作废这张单据！", vbInformation, gstrSysName
                        chkCancel.Value = 0
                    ElseIf Val(StrToNum(txtMoney.Text)) < 0 Then
                        MsgBox "该笔预交金额为负,表示退款,不能执行该操作！", vbExclamation, gstrSysName
                        chkCancel.Value = 0
                    Else
                        If cmdOK.Enabled Then cmdOK.SetFocus
                    End If
                End If
            Case 0
                MsgBox "读取该预交款单据失败！", vbExclamation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus
            Case 1
                MsgBox "该预交款单据不存在！", vbExclamation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus
            Case 2
                MsgBox "该预交款单据已经退款！", vbExclamation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus
            Case 3
                cboNO.Text = "": cboNO.SetFocus
        End Select
    End If
End Sub

Private Sub cboNote_GotFocus()
    cboNote.SelStart = 0: cboNote.SelLength = Len(cboNote.Text)
End Sub

Private Sub cboNote_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtPatient_Change()
    If Not Me.ActiveControl Is txtPatient Or txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjIDCard.SetEnabled(True)
    If Not mobjICCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjICCard.SetEnabled(True)
    txtPatient.Tag = ""
End Sub
Private Sub ClearWinInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除窗体信息
    '编制:刘兴洪
    '日期:2012-02-27 11:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '问题号:55666
    lblRepairMoney.Caption = "补交额:": lblRepairMoney.Visible = False
End Sub
Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    Dim blnCard As Boolean, blnICCard As Boolean
    If txtPatient.Locked Then Exit Sub
    
    Call ClearWinInfor
        
    '特殊字符过滤在Form_KeyPress中进行
    If IDKind.GetCurCard.名称 = "姓名" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.GetCurCard.名称 = "门诊号" Or IDKind.GetCurCard.名称 = "住院号" Or IDKind.GetCurCard.名称 = "手机号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
    End If
    
    If txtPatient.Tag <> "" Then Exit Sub
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        Set frmPatiSelect.mfrmParent = Me
        frmPatiSelect.mbytSize = 1 '大字体(小四)
        frmPatiSelect.Show 1, Me
    End If
    Me.Refresh
    '问题27379
    mstr病人类型 = ""
    txtPatient.ForeColor = &HFF0000
    
    '刷卡完毕或输入号码后回车
    If blnCard And Len(Me.txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtPatient.Text <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtPatient))
    End If
End Sub
Private Sub FindPati(ByVal objcard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnCancel As Boolean, bytPrepayType As Byte
    
    '读取病人信息
    SetMoneyInfo True
    sta.Panels(2) = ""
    If objcard.名称 Like "IC卡*" And objcard.系统 Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If Not GetPatient(objcard, strInput, blnCancel, blnCard) Then
        If blnCancel Then '取消输入
            Call zlControl.TxtSelAll(txtPatient): txtPatient.SetFocus: Exit Sub
        End If
        sta.Panels(2) = "未找到该病人，请检查输入内容!"
        If blnCard = True Then
            txtPatient.PasswordChar = "": txtPatient.Text = ""
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
        Else
            txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
        End If
        Set mrsInfo = New ADODB.Recordset
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Sub
    End If
    '设置病人费用信息
    Call SetMoneyInfo(False, mrsInfo!病人ID)
    Call LoadPatiPage(Val(Nvl(mrsInfo!病人ID)))
    
    '79361:李南春,2014/11/18,缺省病人的预交类型
    '费用查询界面或护士站调用时不自动切换预交类型，以传入的为准
    bytPrepayType = IIf(Val(Nvl(mrsInfo!在院)) = 1, 2, 1)
    If bytPrepayType <> mbytPrepayType And Not (mbytCallObject = 1 Or mblnNurseCall) Then
        mbytPrepayType = bytPrepayType: Call InitPrepayType
    End If
    
    If mrsInfo!当前科室id <> 0 Then
        lbl床号.Caption = lbl床号.Tag & IIf(mrsInfo!床号 = 0, "家庭", mrsInfo!床号)
    End If
            
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(Val(Nvl(mrsInfo!住院号)) = 0, "", "住院号:" & mrsInfo!住院号 & "   ") & _
                           IIf(Val(Nvl(mrsInfo!门诊号)) = 0, "", "门诊号:" & mrsInfo!门诊号)
    lbl科室.Caption = lbl科室.Tag & GET部门名称(mrsInfo!科室ID)
    '46764
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(Val(Nvl(mrsInfo!当前科室id)) = 0, Val(Nvl(mrsInfo!科室ID)), Val(Nvl(mrsInfo!当前科室id))))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
                
    '医保改动-在院病人转个人帐户
    If Not IsNull(mrsInfo!险类) And InStr(mstrPrivs, ";保险转帐;") > 0 And mstr个人帐户 <> "" Then
        If cbo.FindIndex(cboStyle, mstr个人帐户, True) = -1 Then
            cboStyle.AddItem mstr个人帐户
            cboStyle.ItemData(cboStyle.NewIndex) = 3
        End If
        '医保接口
        mcur帐户余额 = gclsInsure.SelfBalance(mrsInfo!病人ID, mrsInfo!医保号, 30, , mrsInfo!险类)
        lbl帐户余额.Caption = lbl帐户余额.Tag & Format(mcur帐户余额, "0.00")
        lbl帐户余额.Visible = True
        lbl预交余额.Left = 2640
        If lbl帐户余额.Visible Then
            Line2(14).Visible = True: Line2(11).X2 = 2415
        Else
            Line2(14).Visible = False: Line2(11).X2 = Line2(14).X2
        End If
    End If
    
    lbl费别等级.Caption = lbl费别等级.Tag & mrsInfo!费别
    lbl担保人.Caption = lbl担保人.Tag & mrsInfo!担保人
    lbl担保金额.Caption = lbl担保金额.Tag & mrsInfo!担保额
    '问题号:116059,焦博,2017/12/7,预交界面显示病人手机号，提取病人信息中的“手机号”
    lbl手机号.Caption = lbl手机号.Tag & mrsInfo!手机号
    chk担保temp.Value = mrsInfo!担保性质
    lblMemo.Caption = lblMemo.Tag & Nvl(mrsInfo!备注)
    '72828,冉俊明,2014-5-9,增加工作单位信息的显示
    lblWorkUnit.Caption = lblWorkUnit.Tag & Nvl(mrsInfo!工作单位)
    
    txtPatient.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    txtPatient.Text = mrsInfo!姓名
    txtPatient.Tag = mrsInfo!病人ID
    '-----------------------------------------------------------------------------------------
    lblSex.Caption = lblSex.Tag & IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别)
    mstrPatiSex = IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别)
    lblOld.Caption = lblOld.Tag & IIf(IsNull(mrsInfo!年龄), "", mrsInfo!年龄)
    mstrPatiOld = IIf(IsNull(mrsInfo!年龄), "", mrsInfo!年龄)
    lbl家庭地址.Caption = lbl家庭地址.Tag & Nvl(mrsInfo!家庭地址)
    lbl医疗付款方式.Caption = lbl医疗付款方式.Tag & Nvl(mrsInfo!医疗付款方式)
    Call Led欢迎信息
    Call SetcmdOkEnabled
    
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Led欢迎信息()
    Dim strInfo As String, lngPatient As Long
    'LED初始化
    If (mbytInState = 0 Or mbytInState = 3) And gblnLED Then
        If gblnLedWelcome Then
            zl9LedVoice.Reset com
            zl9LedVoice.Speak "#1"
            zl9LedVoice.Init UserInfo.编号 & "号为您服务", mlngModul, gcnOracle
        End If
        strInfo = Trim(txtPatient.Text)
        If mrsInfo.State = 1 Then strInfo = strInfo & " " & mrsInfo!性别 & " " & mrsInfo!年龄: lngPatient = Val("" & mrsInfo!病人ID)
        zl9LedVoice.DisplayPatient strInfo, lngPatient
    End If
End Sub
Private Sub Clear个人帐户()
'功能：清除个人帐户信息
    Dim i As Integer
    For i = 0 To cboStyle.ListCount - 1
        If cboStyle.ItemData(i) = 3 Then
            cboStyle.RemoveItem i: Exit For
        End If
    Next
    mcur帐户余额 = 0
    lbl帐户余额.Caption = lbl帐户余额.Tag
    lbl帐户余额.Visible = False: Line2(14).Visible = False
    Line2(11).X2 = Line2(14).X2
    lbl预交余额.Left = lbl帐户余额.Left
End Sub
 

Private Function GetPatient(ByVal objcard As Card, ByVal strInput As String, blnCancel As Boolean, Optional blnCard As Boolean = False, Optional lng主页ID As Long) As Boolean
    '功能：读取病人信息
    '参数：strInput=[刷卡]|[A病人ID]|[B住院号]
    '说明：
    '     1.适用于病人预交款
    '     2.自动识别病人在院状态,读出(病人ID,主页ID,姓名,性别,年龄,住院号,床号,在院标志)
    '返回:是否读取成功,成功时mrsInfo中包含病人信息,失败时mrsInfo=Close
    Dim rsTmp As ADODB.Recordset, strPati As String, strSQL As String
    Dim vRect As RECT, i As Integer, lng卡类别ID As Long, bln存在帐户 As Boolean, lng病人ID As Long, strPassWord As String, strErrMsg As String
    Dim strWhere As String, blnICCard As Boolean
    Dim blnHavePassWord As Boolean
    Dim rsTemp As ADODB.Recordset, str在院病人 As String
    Dim blnIsMobileNO As Boolean
    Dim strRecent As String     '读取最近一次病人信息条件
    
    blnCancel = False
    strWhere = ""
    strRecent = " And Nvl(A.主页ID,0)=C.主页ID(+) "
    If lng主页ID <> 0 Then
        strWhere = strWhere & " And A.病人ID=[2] And C.主页ID=[3]"
        strRecent = ""
        GoTo PatiPage
    End If
    blnIsMobileNO = IDKind.IsMobileNo(strInput)
    Call Clear个人帐户 '清除个人帐户信息
    mdbl预交余额_三方备份 = 0
    chkAllCash.Value = 0: chkAllCash.Visible = False
    mstr退款操作员 = ""
    
    '112999
    If gblnAllowOut = False Then '不允许出院病人缴住院预交
        If mbytInState = 0 And cboType.ItemData(cboType.ListIndex) = 2 _
            Or mbytInState = 4 And cboType.ItemData(cboType.ListIndex) = 1 Then
            
            '在院或预入院
            str在院病人 = " And (Nvl(a.在院, 0) = 1" & vbNewLine & _
                        "       Or Exists (Select 1 From 病案主页" & vbNewLine & _
                        "                  Where 病人id = a.病人id And Nvl(主页id, 0) = 0 And Nvl(病人性质, 0) = 0)) "
        End If
    End If
    
    If (blnCard And objcard.名称 Like "姓名*") _
        And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then   '刷卡或缺省的卡
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then
            If blnIsMobileNO Then
                If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人ID, strPassWord) = False Then
                    GoTo NotFoundPati:
                End If
            Else
                GoTo NotFoundPati:
            End If
        End If
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.病人ID=[1]"
        strInput = "-" & lng病人ID
        blnHavePassWord = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then  '病人ID
        strWhere = strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then  '住院号(对住(过)院的病人)
        strWhere = strWhere & " And (A.病人ID,C.主页ID) In (Select Max(病人id),Max(主页ID) From 病案主页 Where 住院号 = [1])"
        strRecent = ""
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号(仅对门诊病人)
        strWhere = strWhere & " And A.门诊号=[1]"
    Else '当作姓名
        Select Case objcard.名称
            Case "姓名", "姓名或就诊卡"
                
                '限制模糊查长度,如果按照姓查找会影响性能
                If (Not gblnSeekName) Or (gblnSeekName And Len(strInput) < 2) Then
                    Set mrsInfo = New ADODB.Recordset: Exit Function
                End If
                
                strPati = _
                " Select A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄," & _
                "           A.住院号,B.名称 as 科室,A.当前床号 as 床号," & _
                "           A.出生日期,A.身份证号,A.家庭地址,A.卡验证码,Nvl(A.在院,0) As 在院标志 " & _
                " From 病人信息 A,部门表 B " & _
                " Where A.停用时间 is NULL And A.当前科室ID=B.ID(+) And A.姓名 Like [1] " & str在院病人 & _
                "   Order by A.姓名"
                
                
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTmp = zldatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", "bytSize=1")
                If Not rsTmp Is Nothing Then
                    strInput = rsTmp!病人ID
                    strWhere = strWhere & " And A.病人ID=[2]"
                Else
                    Set mrsInfo = New ADODB.Recordset: Exit Function
                End If
            Case "医保号"
                strInput = UCase(strInput)
                strWhere = strWhere & " And A.医保号=[2]"
            Case "IC卡号"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
                blnICCard = (InStr(1, "-+*.", Left(strInput, 1)) = 0) And objcard.系统
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.门诊号=[2]"
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And (A.病人ID,C.主页ID) In (Select Max(病人id),Max(主页ID) From 病案主页 Where 住院号 = [2])"
                strRecent = ""
            Case Else
                '其他类别的,获取相关的病人ID
                If objcard.接口序号 > 0 Then
                    lng卡类别ID = objcard.接口序号
                    bln存在帐户 = objcard.是否存在帐户
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objcard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
PatiPage:
    '问题27379
    '72828,冉俊明,2014-5-9,增加工作单位信息的显示，处理方式：增加A.工作单位字段
    strSQL = _
    " Select A.病人ID,Nvl(C.主页ID,0) as 主页ID,Nvl(C.当前病区ID,0) as 病区ID,Nvl(c.出院科室ID,0) as 科室ID,Nvl(A.当前科室ID,0) as 当前科室ID, Nvl(a.在院,0) as 在院," & _
    "           Decode(Nvl(A.主页ID,0),0,A.医疗付款方式,C.医疗付款方式) 医疗付款方式,Nvl(A.病人类型,C.病人类型) as 病人类型," & _
    "            Nvl(C.姓名, a.姓名) As 姓名, Nvl(C.性别, a.性别) As 性别,A.年龄,Nvl(A.门诊号,0) as 门诊号,Nvl(C.住院号,0) as 住院号,Nvl(C.出院病床,0) as 床号,A.家庭地址,A.卡验证码," & _
    "           B.险类,B.卡号,Nvl(B.医保号,A.医保号) 医保号,B.密码,Nvl(C.费别,A.费别) 费别,A.担保人,A.担保额,Nvl(A.担保性质,0) as 担保性质, A.工作单位,A.手机号,C.备注,Nvl(A.在院,0) As 在院标志" & _
    " From 病人信息 A,医保病人档案 B,病案主页 C,医保病人关联表 E" & _
    " Where A.停用时间 is NULL" & _
    "       And A.病人ID=C.病人ID(+) " & strRecent & _
    "       And C.病人ID=E.病人ID(+) And E.标志(+)=1  " & str在院病人 & _
    "       And E.医保号=B.医保号(+) And E.险类=B.险类(+) And E.中心 = B.中心(+) " & strWhere
    
    On Error GoTo errH
    Set mrsInfo = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, lng主页ID)
    If mrsInfo.EOF Then
        Set mrsInfo = New ADODB.Recordset: Exit Function
    End If
    '需要处理其他
    If mblnCheckPass And (blnCard Or blnICCard Or IDKind.GetCurCard.接口序号 <> 0) Then
        If Not blnHavePassWord Then
            strPassWord = Nvl(mrsInfo!卡验证码)
        End If
        If strPassWord <> "" Then
            If zlCommFun.VerifyPassWord(Me, strPassWord, mrsInfo!姓名, mrsInfo!性别, mrsInfo!年龄) = False Then
                 Set mrsInfo = New ADODB.Recordset: Exit Function
            End If
        End If
    End If
    GetPatient = True
    Exit Function
errH:
     If ErrCenter() = 1 Then Resume
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
End Function

Private Function SaveBill(Optional blnPrintInvoice As Boolean = False, Optional ByRef lng预交ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对当前输入的预交款单据存盘
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 11:15:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String, strSQL As String, i As Integer
    Dim blnInsure As Boolean, strCurDate As String
    Dim blnTrans As Boolean, dblMoney As Double
    Dim lng主页ID As Long
    
    strNo = zldatabase.GetNextNo(11)
    lng预交ID = zldatabase.GetNextId("病人预交记录")
    strCurDate = Format(zldatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    blnInsure = cboStyle.ItemData(cboStyle.ListIndex) = 3 And Not IsNull(mrsInfo!险类)
    '问题27363
    dblMoney = IIf(mbytInState = 3, -1, 1) * StrToNum(txtMoney.Text)
    
Once:
    'Zl_病人预交记录_Insert
    strSQL = "Zl_病人预交记录_Insert("
    '  Id_In         病人预交记录.ID%Type,
    strSQL = strSQL & "" & lng预交ID & ","
    '  单据号_In     病人预交记录.NO%Type,
    strSQL = strSQL & "'" & strNo & "',"
    '  票据号_In     票据使用明细.号码%Type,
    '60669
    If blnPrintInvoice Then
        strSQL = strSQL & "'" & txtFact.Text & "',"
    Else
        strSQL = strSQL & "NULL,"
    End If
    '  病人id_In     病人预交记录.病人id%Type,
    strSQL = strSQL & "" & mrsInfo!病人ID & ","
    '  主页id_In     病人预交记录.主页id%Type,:42329
    '问题:44963
    'mrsInfo!当前科室id <> 0 And mrsInfo!在院 <> 0 And
    
    lng主页ID = IIf(cboType.ItemData(cboType.ListIndex) = 2, Val(Nvl(mrsInfo!主页ID)), 0)
    If cboPatiPage.Visible And cboPatiPage.ListIndex > 0 And cboType.ItemData(cboType.ListIndex) = 2 Then
        lng主页ID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    strSQL = strSQL & "" & IIf(lng主页ID = 0, "NULL", lng主页ID) & ","
    
    '  科室id_In     病人预交记录.科室id%Type,
    strSQL = strSQL & "" & IIf(cboUnit.ItemData(cboUnit.ListIndex) = 0, "NULL", cboUnit.ItemData(cboUnit.ListIndex)) & ","
    '  金额_In       病人预交记录.金额%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  结算方式_In   病人预交记录.结算方式%Type,
    strSQL = strSQL & "'" & mstr结算方式 & "',"
    '  结算号码_In   病人预交记录.结算号码%Type,
    strSQL = strSQL & "'" & txtCode.Text & "',"
    '  缴款单位_In   病人预交记录.缴款单位%Type,
    If blnInsure Then
        strSQL = strSQL & "'" & Nvl(mrsInfo!险类) & "',"
    Else
        strSQL = strSQL & "'" & Trim(txtUnit.Text) & "',"
    End If
    '  单位开户行_In 病人预交记录.单位开户行%Type,
    If blnInsure Then
        strSQL = strSQL & "'" & Nvl(mrsInfo!密码) & "',"
    Else
        strSQL = strSQL & "'" & Trim(txt开户行.Text) & "',"
    End If
    '  单位帐号_In   病人预交记录.单位帐号%Type,
    If blnInsure Then
        strSQL = strSQL & "'" & Nvl(mrsInfo!医保号) & "',"
    Else
        strSQL = strSQL & "'" & Trim(txt帐号.Text) & "',"
    End If
    '  摘要_In       病人预交记录.摘要%Type,
    strSQL = strSQL & "'" & Trim(cboNote.Text) & "',"
    '  操作员编号_In 病人预交记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人预交记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  领用id_In     票据使用明细.领用id%Type,
    strSQL = strSQL & "" & IIf(mlng领用ID = 0, "NULL", mlng领用ID) & ","
    '  预交类别_In   病人预交记录.预交类别%Type := Null,
    strSQL = strSQL & "" & cboType.ItemData(cboType.ListIndex) & ","
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(mlngCardTypeID = 0 Or mbln消费卡 Or chkAllCash.Value = 1, "NULL", mlngCardTypeID) & ","
   '  结算卡序号_in 病人预交记录.结算卡序号%type:=NULL,
    strSQL = strSQL & "" & IIf(mlngCardTypeID = 0 Or Not mbln消费卡, "NULL", mlngCardTypeID) & ","
    '  卡号_In       病人预交记录.卡号%Type := Null,
    If mbytInState = 3 Then
        strSQL = strSQL & "" & IIf(mbytInState = 3 And chkAllCash.Value = 0, "'" & mcurBill.str卡号 & "'", "NULL") & ","
    Else
        strSQL = strSQL & "" & IIf(mstrBrushCardNo = "", "NULL", "'" & mstrBrushCardNo & "'") & ","
    End If
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "" & IIf(mbytInState = 3 And chkAllCash.Value = 0, "'" & mcurBill.str交易流水号 & "'", "NULL") & ","
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "" & IIf(mbytInState = 3, "'" & IIf(chkAllCash.Value = 1, mstr退款操作员 & "强制退现:" & Format(IIf(dblMoney < mdbl预交余额_三方备份, dblMoney, mdbl预交余额_三方备份), "0.00") & "元", mcurBill.str交易说明) & "'", "NULL") & ","
    '  合作单位_In   病人预交记录.合作单位%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  收款时间_In   病人预交记录.收款时间%Type := Null
    strSQL = strSQL & "to_date('" & strCurDate & "','yyyy-mm-dd hh24:mi:ss'),"
    '   操作类型_In Integer:=0 :0-正常缴预交;1-存为划价单;3-余额退款
    strSQL = strSQL & IIf(mbytInState = 3, 3, 0) & ","
    '   结帐id_In     病人预交记录.结帐id%Type := Null
    strSQL = strSQL & "NULL,"
    '   结算性质_In   病人预交记录.结算性质%Type := Null,
    strSQL = strSQL & "NULL,"
    '   退款检查_In   Number := 0
    strSQL = strSQL & mbytOracleBackType & ","
    '   强制退现_In   Number := 0
    strSQL = strSQL & IIf(mbytInState = 3 And chkAllCash.Value = 1, 1, 0) & ","
    '   更新交款余额_In Number := 1,
    strSQL = strSQL & "1,"
    '   是否转账_In     Number := 0
    strSQL = strSQL & IIf(mbytInState = 3 And mcurBill.bln转账, 1, 0) & ")"
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    zldatabase.ExecuteProcedure strSQL, Me.Caption
    If blnInsure Then
        '医保接口
        If Not gclsInsure.TransferSwap(lng预交ID, CCur(dblMoney), mrsInfo!险类) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    End If
    If mbytInState = 3 Then
        If zlDepositDel(mcurBill.lng预交ID, lng预交ID, StrToNum(txtMoney.Text)) = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    Else
        If zlInterfacePrayMoney(lng预交ID, strNo, StrToNum(txtMoney.Text)) = False Then
            '删除无效的预交数据
            gcnOracle.RollbackTrans: Exit Function
            'Call DeletePrepay(strNO): Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    '加入单据历史记录(所有类型单据)
    For i = 0 To cboNO.ListCount - 1
        strNo = strNo & "," & cboNO.List(i)
    Next
    cboNO.Clear
    For i = 0 To UBound(Split(strNo, ","))
        cboNO.AddItem Split(strNo, ",")(i)
        If i = 9 Then Exit For '只显示10个
    Next
    
    If Not gblnBill预交 And blnPrintInvoice And Trim(txtFact.Text) <> "" Then
        '松散：保存当前号码
        zldatabase.SetPara "当前预交票据号", Trim(txtFact.Text), glngSys, mlngFactModule
    End If
    SaveBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If Err.Description Like "*退款金额大于病人剩余预交余额*" And mbytOracleBackType = 1 Then
        If MsgBox("退款金额比病人当前的余额多,是否忽略？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
        mbytOracleBackType = 0
        GoTo Once
    End If
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ReadBill(strNo As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取预交款单据(浏览的、退款的),并填写界面及设置mrsInfo(病人信息),将金额放在Tag中
    '入参:strNO-预交单据号
    '出参:
    '返回: -1-成功;0-失败;1-该单据不存在;2:该单据已经退款(浏览时无效);3-权限不足(已提醒)
    '编制:刘兴洪
    '日期:2011-07-15 11:45:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngPrepayType As Long, rsTemp As New ADODB.Recordset, strFullNO As String
    Dim strWhere As String, lng预交类别 As Long
    Dim i As Long, blnHave As Boolean, strTmp As String
    Dim rs分站点显示 As New ADODB.Recordset
    
    On Error GoTo errH
    strFullNO = GetFullNO(strNo, 11)
    If cboType.ListIndex >= 0 Then
        lng预交类别 = cboType.ItemData(cboType.ListIndex)
    End If
    
    strWhere = IIf(mbytInState = 1, IIf(mblnViewCancel, "And A.记录状态=2", "And A.记录状态 IN(1,3)"), "")
    If mbytCallObject = 1 Or mbytCallObject = 2 Or mbytInState = 3 Then
        strWhere = strWhere & " And Not Exists(Select 1 From 结算方式   Where A.结算方式= 名称 And 性质=5)"
    End If
    
    gstrSQL = "" & _
    "   Select   A.ID,A.预交类别,A.实际票号,A.病人ID,A.主页ID,A.科室ID,A.记录状态,A.摘要,A.金额, " & _
    "               A.结算方式,A.结算号码,A.收款时间,A.操作员姓名,A.缴款单位,A.单位开户行," & _
    "               A.单位帐号,A.卡类别ID,nvl(A.结算卡序号,C.接口编号) as 结算卡序号, " & _
    "               nvl(A.卡号,C.卡号) as 卡号,nvl(A.交易流水号,C.交易流水号) as 交易流水号,A.交易说明,A.合作单位," & _
    "               M.名称 as 卡类别名称, nvl(J.名称,Q.名称) as 消费卡名称,Nvl(M.是否退款验卡,0) as 是否退款验卡,c.消费卡ID " & _
    "   From " & IIf(mblnNOMoved, "H", "") & "病人预交记录 A, " & _
    "          病人卡结算记录 C,消费卡类别目录 J,消费卡类别目录 Q,医疗卡类别 M " & _
    "   Where  A.记录性质=1 And A.No=[1]   " & strWhere & _
    "          And A.ID=c.结算ID(+) and C.接口编号=Q.编号(+)" & _
    "          And A.卡类别ID=M.ID(+) And A.结算卡序号=J.编号(+)"
    
    '72828,冉俊明,2014-5-9,增加工作单位信息的显示，处理方式：增加B.工作单位字段
    gstrSQL = _
    "Select A.实际票号 as 票据号,A.病人ID,A.主页ID,A.科室ID,B.门诊号,B.住院号,nvl(D.姓名,B.姓名) as 姓名,nvl(D.性别,B.性别) as 性别,nvl(D.年龄,B.年龄) as 年龄," & _
    "           A.科室ID As 当前科室ID,B.当前床号,B.家庭地址,A.ID,A.记录状态,A.摘要,A.金额," & _
    "           A.结算方式,C.性质,A.结算号码,A.收款时间,A.操作员姓名,B.合同单位ID," & _
    "           Decode(Nvl(A.主页ID,0),0,B.医疗付款方式,D.医疗付款方式) 医疗付款方式," & _
    "           Decode(Nvl(C.性质,1),3,NULL,A.缴款单位) as 缴款单位," & _
    "           Decode(Nvl(C.性质,1),3,NULL,A.单位开户行) as 单位开户行," & _
    "           Decode(Nvl(C.性质,1),3,NULL,A.单位帐号) as 单位帐号,Nvl(D.费别,B.费别) 费别," & _
    "           B.担保人,B.担保额,Nvl(B.担保性质,0) as 担保性质, B.工作单位,B.手机号," & _
    "           B.病人类型,B.险类," & _
    "           NVL(A.预交类别,0) as 预交类别, " & _
    "           A.卡类别ID,A.结算卡序号,A.卡号,A.交易流水号,A.交易说明,A.合作单位,A.卡类别名称, A.消费卡名称,A.是否退款验卡,a.消费卡ID " & _
    " From (" & gstrSQL & ") A, 病人信息 B,结算方式 C,病案主页 D" & _
    " Where A.病人ID=B.病人ID  And B.病人ID=D.病人ID(+) And nvl(B.主页ID,0)=D.主页ID(+) And A.结算方式=C.名称(+)"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFullNO)
    If rsTemp.RecordCount = 0 Then ReadBill = 1: Exit Function
    If mbytInState = 2 Or chkCancel.Value = 1 Then
        '退款,需要检查是否存在具体的退款权限
        lngPrepayType = Val(Nvl(rsTemp!预交类别))
        If InStr(1, mstrPrivs, IIf(lngPrepayType = 1, ";门诊预交;", ";住院预交;")) = 0 Then
            MsgBox "你不具备对预交单据进行退款的权限,请与系统管理员联系!", vbOKOnly + vbInformation, gstrSysName
            ReadBill = 3
            Exit Function
        End If
        
        If gbln分站点显示 Then
            strTmp = "Select 1 From 部门表 A, 部门人员 B, 人员表 C" & vbNewLine & _
                    " Where a.Id = b.部门id And b.人员id = c.Id And (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null) And c.姓名 =[1]  And Rownum < 2"
    
            Set rs分站点显示 = zldatabase.OpenSQLRecord(strTmp, Me.Caption, Nvl(rsTemp!操作员姓名))
            
            If rs分站点显示.RecordCount = 0 Then
                MsgBox "该预交单据不属于本站点,不允许退款!", vbOKOnly + vbInformation, gstrSysName
                ReadBill = 3: Exit Function
            End If
        End If
    End If
    
    With mcurBill
        .strNo = strFullNO
        .lng预交ID = Val(Nvl(rsTemp!ID))
        .lng卡类别ID = IIf(Val(Nvl(rsTemp!卡类别ID)) = 0, Val(Nvl(rsTemp!结算卡序号)), Val(Nvl(rsTemp!卡类别ID)))
        .bln消费卡 = Val(Nvl(rsTemp!结算卡序号)) <> 0
        .str名称 = IIf(.bln消费卡, Nvl(rsTemp!消费卡名称), Nvl(rsTemp!卡类别名称))
        .str卡号 = Nvl(rsTemp!卡号)
        .bln退款验卡 = Val(Nvl(rsTemp!是否退款验卡)) = 1
        .str交易流水号 = Nvl(rsTemp!交易流水号)
        .str交易说明 = Nvl(rsTemp!交易说明)
        .str合作单位 = Nvl(rsTemp!合作单位)
        .dt收款时间 = Format(rsTemp!收款时间, "yyyy-MM-dd hh:mm:ss")
        .lng消费卡ID = Val(Nvl(rsTemp!消费卡ID))
    End With
    
    cboNO.Text = strFullNO
    cboNO.Tag = rsTemp!ID '以此ID为准退款
    txtPatient.Text = rsTemp!姓名
    txtPatient.Tag = rsTemp!病人ID
    '74426:李南春,2014-7-9,病人姓名显示颜色处理
    Call SetPatiColor(txtPatient, Nvl(rsTemp!病人类型), IIf(IsNull(rsTemp!险类), &HFF0000, vbRed))
    lbl费别等级.Caption = lbl费别等级.Tag & rsTemp!费别
    
    lbl担保人.Caption = lbl担保人.Tag & rsTemp!担保人
    lbl担保金额.Caption = lbl担保金额.Tag & rsTemp!担保额
    lbl手机号.Caption = lbl手机号.Tag & rsTemp!手机号
    chk担保temp.Value = rsTemp!担保性质
    '72828,冉俊明,2014-5-9,增加工作单位信息的显示
    lblWorkUnit.Caption = lblWorkUnit.Tag & Nvl(rsTemp!工作单位)
    
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(IsNull(rsTemp!当前科室id), 0, rsTemp!当前科室id))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    cboType.ListIndex = -1
    lngPrepayType = Val(Nvl(rsTemp!预交类别))
    For i = 0 To cboType.ListCount - 1
         If cboType.ItemData(i) = lngPrepayType Then
            cboType.ListIndex = i: Exit For
         End If
     Next
     
     With cboType
        If cboType.ListIndex < 0 Then
           .AddItem IIf(lngPrepayType = 1, "门诊预交", "住院预交")
           .ItemData(.NewIndex) = IIf(lngPrepayType = 1, 1, 2)
           .ListIndex = .NewIndex
        End If
     End With
     
     With cboPatiPage
        .Clear
        .Visible = lngPrepayType <> 1
        If Val(Nvl(rsTemp!主页ID)) <> 0 Then
            .AddItem "第" & rsTemp!主页ID & "次"
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!主页ID))
            .ListIndex = .NewIndex
        Else
            .AddItem "预约入院"
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!主页ID))
            .ListIndex = .NewIndex
        End If
     End With
     
    txtFact.Text = IIf(IsNull(rsTemp!票据号), "", rsTemp!票据号)
    txtUnit.Text = IIf(IsNull(rsTemp!缴款单位), "", rsTemp!缴款单位)
    txt开户行.Text = IIf(IsNull(rsTemp!单位开户行), "", rsTemp!单位开户行)
    txt帐号.Text = IIf(IsNull(rsTemp!单位帐号), "", rsTemp!单位帐号)
    
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(Val(Nvl(rsTemp!住院号)) = 0, "", "住院号:" & rsTemp!住院号 & "   ") & _
                           IIf(Val(Nvl(rsTemp!门诊号)) = 0, "", "门诊号:" & rsTemp!门诊号)
    lblSex.Caption = lblSex.Tag & IIf(IsNull(rsTemp!性别), "", rsTemp!性别)
    mstrPatiSex = IIf(IsNull(rsTemp!性别), "", rsTemp!性别)
    lblOld.Caption = lblOld.Tag & IIf(IsNull(rsTemp!年龄), "", rsTemp!年龄)
    mstrPatiOld = IIf(IsNull(rsTemp!年龄), "", rsTemp!年龄)
    lbl床号.Caption = lbl床号.Tag & IIf(IsNull(rsTemp!当前床号), "", rsTemp!当前床号)
    lbl科室.Caption = lbl科室.Tag & GET部门名称(IIf(IsNull(rsTemp!当前科室id), 0, rsTemp!当前科室id))
    lbl家庭地址.Caption = lbl家庭地址.Tag & Nvl(rsTemp!家庭地址)
    lbl医疗付款方式.Caption = lbl医疗付款方式.Tag & Nvl(rsTemp!医疗付款方式)
    txtMoney.Text = Format(rsTemp!金额, "##,##0.00;-##,##0.00;;")
    txtMoney.Tag = rsTemp!金额
    If mcurBill.lng卡类别ID <> 0 Then
        cboStyle.ListIndex = cbo.FindIndex(cboStyle, mcurBill.str名称, True)
    Else
        cboStyle.ListIndex = cbo.FindIndex(cboStyle, IIf(IsNull(rsTemp!结算方式), "", rsTemp!结算方式), True)
     End If
    If cboStyle.ListIndex = -1 Then
        If mcurBill.lng卡类别ID <> 0 Then
            cboStyle.AddItem mcurBill.str名称
            cboStyle.ItemData(cboStyle.NewIndex) = -1
        Else
            cboStyle.AddItem IIf(IsNull(rsTemp!结算方式), "", rsTemp!结算方式)
            cboStyle.ItemData(cboStyle.NewIndex) = Val("" & rsTemp!性质)
        End If
        cboStyle.ListIndex = cboStyle.NewIndex
        
    End If
    
    txtCode.Text = IIf(IsNull(rsTemp!结算号码), "", rsTemp!结算号码)
    txtMan.Text = IIf(IsNull(rsTemp!操作员姓名), "", rsTemp!操作员姓名)
    txtDate.Text = Format(rsTemp!收款时间, "yyyy-MM-dd")
    cboNote.Text = IIf(IsNull(rsTemp!摘要), "", rsTemp!摘要)
    '获取病人费用信息
    Call SetMoneyInfo(False, rsTemp!病人ID, strNo)
    ReadBill = -1
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub GetDepositData(ByVal lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新读取预交数据
    '入参:lng病人ID-病人ID巧
    '编制:刘兴洪
    '日期:2011-07-22 17:02:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, int预交类型 As Integer
    Dim strWhere As String, strTittle As String
    
    On Error GoTo errHandle
    If lng病人ID = 0 Then
        If mrsInfo Is Nothing Then Set mrsDepositBalance = Nothing: Exit Sub
        If mrsInfo.State <> 1 Then Set mrsDepositBalance = Nothing: Exit Sub
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    
    mdbl费用余额 = 0: mdbl预交余额 = 0: mdbl剩余款额 = 0
    '按类别先缓存,以提搞性能
    Set mrsDepositBalance = GetMoneyInfo(lng病人ID, , , , True)
    '按类别,分别汇总统计,只有退款时,才会发生
    If mbytInState <> 3 Then Exit Sub
    strSQL = "" & _
    "   Select A.预交类别,nvl(A.卡类别ID,0) as 卡类别ID,nvl(A.结算卡序号,0) as 结算卡序号, " & _
    "           A.卡号,A.交易流水号,A.交易说明," & _
    "           max(decode(sign(金额),-1, decode(A.记录状态,1,0,2,0,ID),ID)) as 预交ID," & _
    "           nvl(sum(金额),0)-nvl(sum(nvl(冲预交,0)),0) as 预交余额 " & _
    "   From 病人预交记录 A " & _
    "   Where   A.病人ID=[1] and (nvl(A.结算卡序号,0)<>0 or nvl(卡类别ID,0)<>0) " & _
    "   Group by A.预交类别,nvl(A.卡类别ID,0),nvl(A.结算卡序号,0),A.卡号,A.交易流水号,A.交易说明" & _
    "   Having nvl(sum(金额),0)-nvl(sum(nvl(冲预交,0)),0)  <>0"
        
    strSQL = "" & _
    "   Select RowNum as ID,A.预交类别, A.预交ID, " & _
    "           A.卡类别ID,A.结算卡序号 as 消费接口ID, " & _
    "          nvl(B.编码,C.编号) as 编码,nvl(B.名称,C.名称) as 名称, " & _
    "          Decode(B.编码,NULL,C.是否全退,B.是否全退) as 是否全退," & _
    "          Decode(B.编码,NULL,C.是否退现,B.是否退现) as 是否退现," & _
    "          A.卡号,A.交易流水号,A.交易说明," & _
    "          A.预交余额 " & _
    "   From (" & strSQL & ") A,医疗卡类别 B,消费卡类别目录 C" & _
    "   Where   A.卡类别ID=B.ID(+)  and A.结算卡序号=C.编号(+)  and nvl(A.预交余额,0)>0" & _
    "   Order by 编码,A.卡号,A.交易流水号,A.交易说明"
    Set mrsDepositInfor = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ShowPremayBalance(ByVal blnreReadData As Boolean, ByVal lng病人ID As Long, _
    Optional ByVal blnNotSelect三方退款 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据相关的结算方式和门诊类型,显示预交余额
    '入参:blnReRead-重读数据
    '       lng病人ID-读取指定的病人ID(0时,从mrsInfo记录中读取病人ID)
    '编制:刘兴洪
    '日期:2011-07-21 15:44:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset, strSQL As String
    Dim int预交类型 As Integer, bln三方接口 As Boolean
    Dim strWhere As String, strTittle As String, strPrevName As String
    Dim dbl金额 As Double, strTemp As String, strBlance As String, str名称 As String
    Dim strNotBalance As String, strCardBalance As String, blnPrevCash As Boolean
    Dim dbl退现 As Double, dbl不能退现 As Double, dbl为审 As Double, dbl未缴 As Double, dblYB As Double
    Dim lng主页ID As Long
    
    On Error GoTo errHandle
    If lng病人ID = 0 Then
        If mrsInfo Is Nothing Then Exit Sub
        If mrsInfo.State <> 1 Then Exit Sub
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    
    If blnreReadData Then Call GetDepositData(lng病人ID)
    If cboStyle.ListIndex < 0 And mbytInState <> 4 Then Exit Sub
    sta.Panels(2).Text = ""
    mdbl费用余额 = 0: mdbl预交余额 = 0: mdbl剩余款额 = 0: mdbl预交余额_三方 = 0
    int预交类型 = cboType.ItemData(cboType.ListIndex)
    If mbytInState = 4 Then
        bln三方接口 = False
    Else
        bln三方接口 = cboStyle.ItemData(cboStyle.ListIndex) = -1
    End If
    strWhere = "And nvl(卡类别ID,0)<>0 or nvl(结算卡序号,0)<>0 "
    If mbytInState = 3 Then
            mcurBill.bln消费卡 = False
            mcurBill.lng卡类别ID = 0
            mcurBill.str交易流水号 = ""
            mcurBill.str交易说明 = ""
            mcurBill.str卡号 = ""
            mcurBill.bln退款验卡 = False
            If bln三方接口 Then
                If mbln消费卡 Then
                    mrsDepositInfor.Filter = "预交类别=" & int预交类型 & " and  消费接口ID=" & mlngCardTypeID
                Else
                    mrsDepositInfor.Filter = "预交类别=" & int预交类型 & " and 卡类别ID=" & mlngCardTypeID
                End If
                sta.Panels(2).Text = ""
                If mrsDepositInfor.RecordCount <> 0 Then
                    strTemp = "": strTittle = "": strBlance = ""
                    With mrsDepositInfor
                        .Sort = "卡号,交易流水号,交易说明"
                        dbl金额 = 0
                        Do While .EOF = False
                            'A.卡号,A.交易流水号,A.交易说明
                            str名称 = Nvl(!名称): strTemp = Nvl(!卡号)
                            If strTemp <> strBlance Then
                                If strBlance <> "" Then
                                    strTittle = strTittle & strBlance & ":" & Format(dbl金额, "###0.00;-###0.00;;") & Space(2)
                                End If
                                strBlance = strTemp: dbl金额 = 0
                            End If
                            dbl金额 = dbl金额 + Val(Nvl(!预交余额, 0))
                            .MoveNext
                        Loop
                        If strBlance <> "" Then
                            strTittle = strTittle & strBlance & ":" & Format(dbl金额, "###0.00;-###0.00;;") & Space(2)
                        End If
                        
                        sta.Panels(2).Text = str名称 & ":" & strTittle
                    End With
                End If
                If blnNotSelect三方退款 = False Then Call Select三方退款
            Else
                mrsDepositInfor.Filter = "预交类别=" & int预交类型
                If mrsDepositInfor.RecordCount <> 0 Then
                    strTemp = "": strTittle = "": strBlance = ""
                    With mrsDepositInfor
                        .Sort = "卡类别ID,消费接口ID,编码"
                        dbl金额 = 0
                        Do While .EOF = False
                            'A.卡号,A.交易流水号,A.交易说明
                            strTemp = Nvl(!卡类别ID) & "-" & Nvl(!消费接口ID) & "-" & Nvl(!编码)
                            If strTemp <> strBlance Then
                                If strBlance <> "" Then
                                    If blnPrevCash Then
                                        strCardBalance = strCardBalance & strPrevName & ":" & Format(dbl金额, "###0.00;-###0.00;;") & Space(2)
                                    Else
                                        strNotBalance = strNotBalance & strPrevName & ":" & Format(dbl金额, "###0.00;-###0.00;;") & Space(2)
                                    End If
                                End If
                                blnPrevCash = Val(Nvl(!是否退现)) = 1
                                strPrevName = Nvl(!名称)
                                strBlance = strTemp: dbl金额 = 0
                            End If
                            str名称 = Nvl(!名称)
                            dbl金额 = dbl金额 + Val(Nvl(!预交余额, 0))
                            If Nvl(!是否退现) <> 1 Then
                                mdbl预交余额_三方 = mdbl预交余额_三方 + Val(Nvl(!预交余额, 0))
                                dbl不能退现 = dbl不能退现 + Val(Nvl(!预交余额, 0))
                            Else
                                dbl退现 = dbl退现 + Val(Nvl(!预交余额, 0))
                            End If
                            .MoveNext
                        Loop
                        If dbl金额 <> 0 Then
                            If blnPrevCash Then
                                strCardBalance = strCardBalance & strPrevName & ":" & Format(dbl金额, "###0.00;-###0.00;;") & Space(5)
                            Else
                                strNotBalance = strNotBalance & strPrevName & ":" & Format(dbl金额, "###0.00;-###0.00;;") & Space(5)
                            End If
                        End If
                        If strBlance <> "" Then strTittle = strTittle & _
                            IIf(strCardBalance = "", "", "允许退现" & Format(dbl退现, "###0.00;-###0.00;;") & "元,其中:" & strCardBalance) & _
                            IIf(strNotBalance = "", "", "不允许退现" & Format(dbl不能退现, "###0.00;-###0.00;;") & "元,其中:" & strNotBalance)
                        sta.Panels(2).Text = strTittle
                    End With
                End If
                If mdbl预交余额_三方 <> 0 Then
                    mdbl预交余额_三方备份 = mdbl预交余额_三方
                    chkAllCash.Visible = True
                    chkAllCash.Enabled = True
                End If
            End If
    End If
    If Not mrsDepositBalance Is Nothing Then
    With mrsDepositBalance
        .Filter = "类型=" & int预交类型
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            mdbl费用余额 = mdbl费用余额 + Val(Nvl(!费用余额))
            mdbl预交余额 = mdbl预交余额 + Val(Nvl(!预交余额))
            .MoveNext
        Loop
    End With
    End If
    '医保预结余额
    If cboPatiPage.Visible And cboPatiPage.ListIndex >= 0 And cboType.ItemData(cboType.ListIndex) = 2 Then
        lng主页ID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    Set rsMoney = New ADODB.Recordset
    If lng主页ID = 0 Then
        strSQL = "Select Sum(金额) As 医保预结 From 保险模拟结算 Where 病人ID = [1] And 主页ID Is Null"
        Set rsMoney = zldatabase.OpenSQLRecord(strSQL, "提取医保预结", lng病人ID)
    Else
        strSQL = "Select Sum(金额) As 医保预结 From 保险模拟结算 Where 病人ID = [1] And 主页ID = [2]"
        Set rsMoney = zldatabase.OpenSQLRecord(strSQL, "提取医保预结", lng病人ID, lng主页ID)
    End If
    
    If Not rsMoney.EOF Then
        If Val(Nvl(rsMoney!医保预结, 0)) > 0 Then
            dblYB = Val(Nvl(rsMoney!医保预结, 0))
            lbl医保预结.Caption = lbl医保预结.Tag & Format(rsMoney!医保预结, "##,##0.00;-##,##0.00; ;")
        Else
            lbl医保预结.Caption = lbl医保预结.Tag
        End If
    Else
        lbl医保预结.Caption = lbl医保预结.Tag
    End If

    mdbl剩余款额 = mdbl预交余额 - mdbl费用余额
    '问题27363
    lbl费用余额.Caption = lbl费用余额.Tag & Format(mdbl费用余额, "##,##0.00;-##,##0.00; ;")
    lbl预交余额.Caption = lbl预交余额.Tag & Format(mdbl预交余额, "##,##0.00;-##,##0.00; ;")
    dbl为审 = GetUnAuditedFee(lng病人ID, , int预交类型)
    dbl未缴 = GetUnAuditedFee(lng病人ID, False, int预交类型)
    lbl未审费用.Caption = lbl未审费用.Tag & Format(dbl为审, "##,##0.00;-##,##0.00; ;")
    lbl未缴费用.Caption = lbl未缴费用.Tag & Format(dbl未缴, "##,##0.00;-##,##0.00; ;")
    lbl剩余款额.Caption = lbl剩余款额.Tag & Format(mdbl剩余款额 - dbl未缴 - dbl为审 + dblYB, "##,##0.00;-##,##0.00; ;")
    If mbytInState = 3 Then
        If bln三方接口 Then
            If mdbl剩余款额 - mdbl预交余额_三方 >= 0 Then
                txtMoney.Text = IIf(mdbl预交余额_三方 > 0, Format(mdbl预交余额_三方, "#,##0.00;-#,##0.00;;"), "0.00")
            Else
                txtMoney.Text = IIf(mdbl剩余款额 > 0, Format(mdbl剩余款额, "#,##0.00;-#,##0.00;;"), "0.00")
            End If
        Else
            If mdbl剩余款额 - mdbl预交余额_三方 >= 0 Then
                txtMoney.Text = Format(mdbl剩余款额 - mdbl预交余额_三方, "#,##0.00;-#,##0.00;;")
            Else
                txtMoney.Text = "0.00"
            End If
        End If
    Else
        If mbytInState = 4 Then
            txtMoney.Text = Format(mdbl剩余款额, "#,##0.00;-#,##0.00;;")
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetMoneyInfo(blnClear As Boolean, Optional lng病人ID As Long, _
    Optional strBackNo As String = "", Optional ByVal blnNotSelect三方退款 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示金额等信息
    '入参:blnClear-清除
    '     lng病人ID-指定病人ID
    '     strBackNO-指定退预交单号(退款时传入,主要是是定位到清单上面去)
    '编制:刘兴洪
    ' 修改:刘兴洪(退号时,增加定位功能),增加参数;strBackNo
    '日期:2011-07-21 15:40:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset, lng主页ID As Long
    Dim strSQL As String, lngRow As Long
    
    If blnClear Then
        lblSex.Caption = lblSex.Tag: mstrPatiSex = ""
        lblOld.Caption = lblOld.Tag: mstrPatiOld = ""
        lblPatientNO.Caption = lblPatientNO.Tag
        lbl床号.Caption = lbl床号.Tag
        lbl科室.Caption = lbl科室.Tag
        lbl家庭地址.Caption = lbl家庭地址.Tag
        lbl医疗付款方式.Caption = lbl医疗付款方式.Tag
        lbl担保人.Caption = lbl担保人.Tag
        lbl担保金额.Caption = lbl担保金额.Tag
        chk担保temp.Value = 0
        '72828,冉俊明,2014-5-9,增加工作单位信息的显示
        lblWorkUnit.Caption = lblWorkUnit.Tag
        
        lbl未审费用.Caption = lbl未审费用.Tag
        lbl未缴费用.Caption = lbl未缴费用.Tag
        lbl费用余额.Caption = lbl费用余额.Tag
        lbl预交余额.Caption = lbl预交余额.Tag
        lbl剩余款额.Caption = lbl剩余款额.Tag
        lbl医保预结.Caption = lbl医保预结.Tag
        lbl手机号.Caption = lbl手机号.Tag
        lbl应收款.Caption = lbl应收款.Tag
        lbl应收款.ForeColor = &H80000007
        
        mdbl费用余额 = 0
        mdbl预交余额 = 0
        mdbl剩余款额 = 0
        
        mshList.Redraw = False
        mshList.Clear
        mshList.Rows = 2
        mshList.Cols = 2
        mshList.Redraw = True
    Else
        On Error GoTo errHandle
        '显示预交余额
        Call ShowPremayBalance(True, lng病人ID, blnNotSelect三方退款)
        '检查是否有应收款
        strSQL = "Select Zl_Patientdue([1]) 剩余应收 From dual"
        Set rsMoney = New ADODB.Recordset
        Set rsMoney = zldatabase.OpenSQLRecord(strSQL, "提取应收款", lng病人ID)
        If Not rsMoney.EOF Then
            If Nvl(rsMoney!剩余应收, 0) > 0 Then
                MsgBox "请注意，该病人尚有 " & rsMoney!剩余应收 & "元 应收款未缴！", vbInformation, gstrSysName
                lbl应收款.Caption = lbl应收款.Tag & Format(rsMoney!剩余应收, "##,##0.00;-##,##0.00; ;")
                lbl应收款.ForeColor = &HFF&
            End If
        End If
        Call ShowHistoryPrepay(strBackNo)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub ShowHistoryPrepay(ByVal strBackNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示历史的预交数据
    '编制:刘兴洪
    '日期:2011-09-16 10:17:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, int类型 As Integer, lngRow As Long, strWhere As String
    Dim rsMoney As ADODB.Recordset
    Dim lng病人ID As Long
    If mrsInfo Is Nothing Then
        lng病人ID = mlng病人ID
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = mlng病人ID
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    
    If cboType.ListIndex < 0 Then
         int类型 = 1 'cboType.ItemData(cboType.ListIndex)
    Else
        int类型 = cboType.ItemData(cboType.ListIndex)
    End If
    
    On Error GoTo errHandle
    '84217,李南春,2015/4/22,显示指定的住院期间缴纳的预交
    If cboType.Text = "住院预交" And chk仅显示本次预交.Value = 1 And cboPatiPage.ListIndex >= 0 Then
        strWhere = " And A.主页ID= " & cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    
    If gbln分站点显示 Then
        strWhere = strWhere & _
                " And Exists (Select 1 From 人员表 C, 部门人员 D, 部门表 E " & _
                " Where C.姓名 =A.操作员姓名 And C.Id = D.人员id And D.部门id = E.Id And (E.站点 = '" & gstrNodeNo & "' Or E.站点 Is Null))"
    End If
    
    If gblnShowHave Then
        '只显示有剩余的历史缴款
        '该子查阅用于消除第一次结帐时的一正一负
        strSQL = _
        "   Select NO,Sum(Nvl(A.金额,0)) as 金额  " & _
        "    From " & IIf(mblnNOMoved, "H", "") & "病人预交记录 A" & _
        "   Where A.结帐ID Is Null And Nvl(A.金额, 0)<>0 And A.病人ID=[1] And A.预交类别=[2] " & _
        "   Group by NO " & _
        "   Having Sum(Nvl(A.金额,0))<>0"
        
        strSQL = _
        " Select LTrim(To_Char(A.收款时间,'YYYY-MM-DD')) as 日期,A.NO as 单据号," & _
        "           C.名称 as 科室,Ltrim(To_Char(Nvl(A.金额,0),'9,999,999,990.00')) as 剩余金额,A.结算方式 as 结算,A.操作员姓名 as 收款人" & _
        " From " & IIf(mblnNOMoved, "H", "") & "病人预交记录 A,(" & strSQL & ") B,部门表 C" & _
        " Where A.结帐ID Is Null And A.预交类别=[2]  And Nvl(A.金额,0)<>0 And A.科室ID=C.ID(+)" & _
        "       And A.结算方式 Not IN(Select 名称 From 结算方式 Where 性质=5)" & _
        "       And A.NO=B.NO And A.病人ID=[1] " & strWhere & _
        " Union All" & _
        " Select Min(LTrim(To_Char(A.收款时间,'YYYY-MM-DD'))) as 日期,A.NO as 单据号," & _
        "           B.名称 as 科室,Ltrim(To_Char(Sum(Nvl(A.金额,0)-Nvl(A.冲预交,0)),'9,999,999,990.00')) as 剩余金额,A.结算方式 as 结算,A.操作员姓名 as 收款人" & _
        " From " & IIf(mblnNOMoved, "H", "") & "病人预交记录 A,部门表 B" & _
        " Where A.记录性质 IN(1,11) And A.结帐ID is Not NULL And A.科室ID=B.ID(+) And A.预交类别=[2] " & _
        "       And Nvl(A.金额,0)<>Nvl(A.冲预交,0) And A.病人ID=[1] " & strWhere & _
        " Having Sum(Nvl(A.金额,0)-Nvl(A.冲预交,0))<>0" & _
        " Group by A.NO,B.名称,A.结算方式,A.操作员姓名" & _
        " Order by 日期,单据号,结算"
    Else
        '所有历史缴款明细清单
        strSQL = _
        " Select Ltrim(To_Char(A.收款时间,'YYYY-MM-DD')) as 日期,A.NO as 单据号,B.名称 as 科室, " & _
        " Ltrim(To_Char(A.金额,'9,999,999,990.00')) as 缴款金额,A.结算方式 as 结算,A.操作员姓名 as 收款人 " & _
        " From " & IIf(mblnNOMoved, "H", "") & "病人预交记录 A,部门表 B" & _
        " Where A.科室ID=B.ID(+) And A.记录性质=1 And A.病人ID=[1]  And A.预交类别=[2] " & strWhere & _
        " Order by A.收款时间 Desc"
    End If
    
    Set rsMoney = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, int类型)
    mshList.Clear
    If Not rsMoney.EOF Then
        Set mshList.DataSource = rsMoney
        mshList.ColWidth(0) = 1350: mshList.ColAlignment(0) = 4
        mshList.ColWidth(1) = 1110: mshList.ColAlignment(1) = 4
        mshList.ColWidth(2) = 1200: mshList.ColAlignment(2) = 1
        mshList.ColWidth(3) = 1600: mshList.ColAlignment(3) = 7
        mshList.ColWidth(4) = 1000: mshList.ColAlignment(4) = 4
        mshList.ColWidth(5) = 1000: mshList.ColAlignment(5) = 1
    End If
    If mshList.Rows > 1 Then
        mshList.Row = 1: mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    End If
    
    '刘兴洪:24386,增加一个定位的功能
    If strBackNo <> "" Then
        lngRow = zlControl.MshGrdFindRow(mshList, strBackNo, 1)
        If lngRow > 0 Then
            mshList.Row = lngRow: mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
            If (Not mshList.RowIsVisible(mshList.Row)) Or ((mshList.Row + 1) * mshList.RowHeight(0)) + 50 > mshList.Height Then mshList.TopRow = mshList.Row
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub txtFact_GotFocus()
    zlControl.TxtSelAll txtFact
End Sub

Private Sub txtFact_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or InStr("0123456789" & Chr(8), Chr(KeyAscii)) > 0) Then
        KeyAscii = 0
    ElseIf Len(txtFact.Text) = txtFact.MaxLength And KeyAscii <> 8 And txtFact.SelLength <> Len(txtFact) Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    '问题27379 by lesfeng 2010-01-18
    If mrsInfo.State = 1 Then
        mstr病人类型 = IIf(IsNull(mrsInfo!病人类型), "", mrsInfo!病人类型)
    End If
    If mstr病人类型 = "" Then
        If mrsInfo.State = 1 Then
            If GetOutPatient(mrsInfo!病人ID) Then
                txtPatient.ForeColor = vbRed
            Else
                txtPatient.ForeColor = &HFF0000
            End If
        Else
            txtPatient.ForeColor = &HFF0000
        End If
    Else
        txtPatient.ForeColor = zldatabase.GetPatiColor(mstr病人类型, True)
    End If
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

Private Sub txtUnit_GotFocus()
    zlControl.TxtSelAll txtUnit
End Sub

Private Sub txt开户行_GotFocus()
    zlControl.TxtSelAll txt开户行
End Sub

Private Sub txt开户行_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("~!%^""'|`", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    Else
        CheckInputLen txt开户行, KeyAscii
    End If
End Sub

Private Sub txt帐号_GotFocus()
    zlControl.TxtSelAll txt帐号
End Sub

Private Sub txt帐号_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("~!%^""'|`", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    Else
        CheckInputLen txt帐号, KeyAscii
    End If
End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("~!%^""'|`", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    Else
        CheckInputLen txtUnit, KeyAscii
    End If
End Sub

Private Function InitUnit() As Boolean
'大改：
'功能：初始化门诊，住院临床科室信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称,A.简码,B.服务对象 " & _
        "from 部门表 A,部门性质说明 B " & _
        "Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
        "and B.部门ID=A.ID and B.服务对象 IN(1,2,3) AND B.工作性质 IN('临床','手术') " & _
        "Order by B.服务对象,A.编码"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboUnit.Clear
    cboUnit.AddItem "无"
    cboUnit.ItemData(0) = 0
    cboUnit.ListIndex = 0
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & IIf(IsNull(rsTmp!名称), "", rsTmp!名称)
            cboUnit.ItemData(cboUnit.ListCount - 1) = rsTmp!ID
            rsTmp.MoveNext
        Next
    End If
    
    If Not gbln缴款科室 Then
        cboUnit.Locked = True
        cboUnit.TabStop = False
    End If
    
    InitUnit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CancelBill(lngID As Long, blnCanDel As Boolean, intInsure As Integer, bln打印 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对指定ID的预交款单据执行退款处理
    '入参:lngID=单据ID
    '        blnCanDel=是否支持退个人帐户
    '        intInsure=单据中所使用的个人帐户的保险类别,无为0
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-07-19 09:28:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, blnTrans As Boolean
    Dim lng冲预交ID As Long
    If mcurBill.lng卡类别ID <> 0 Then
        lng冲预交ID = zldatabase.GetNextId("病人预交记录")
    End If
    On Error GoTo errH
    '111209:谢荣，2017/07/24，预交款单据作废打印红票时,不能退预交。
    strSQL = "zl_病人预交记录_DELETE(" & lngID & ",'" & cboNote.Text & "','" & _
        UserInfo.编号 & "','" & UserInfo.姓名 & "'," & IIf(blnCanDel, 1, 0) & "," & IIf(lng冲预交ID = 0, "NULL", lng冲预交ID) & "," & _
        "'" & IIf(bln打印, mstrRedFact, "") & "'," & IIf(bln打印, IIf(mlng领用ID > 0, mlng领用ID, "Null"), 0) & ")"
    gcnOracle.BeginTrans: blnTrans = True
    zldatabase.ExecuteProcedure strSQL, Me.Caption
    '处理医保接口
    If intInsure <> 0 And blnCanDel Then
        If Not gclsInsure.TransferDelSwap(lngID, intInsure) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    End If
    If zlDepositDel(lngID, lng冲预交ID, StrToNum(txtMoney.Text)) = False Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    
    If Not gblnBill预交 And bln打印 And mstrRedFact <> "" Then
        '松散：保存当前号码
        zldatabase.SetPara "当前预交票据号", mstrRedFact, glngSys, mlngFactModule
    End If
    CancelBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOutPatient(ByVal lngID As Long) As Boolean
'功能：判断门诊病人是否属于医保
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim int险类 As Integer
    
    GetOutPatient = False
    On Error GoTo errH
    
    strSQL = _
        "Select 险类 " & _
        "from 病人信息 " & _
        "Where 病人id = [1] and rownum <= 1 "

    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    
    If Not rsTmp.EOF Then
        int险类 = IIf(IsNull(rsTmp!险类), -1, rsTmp!险类)
        GetOutPatient = int险类 <> -1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:创建或关闭结算卡对象
'    '入参:blnClosed:关闭对象
'    '编制:刘兴洪
'    '日期:2010-01-05 14:51:23
'    '问题:
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strExpend As String
'
'   If mbytInState = 1 Then Exit Sub
'    '只有:执行或退费时,才可能管结算卡的
'    If blnClosed Then
'FromClose:
'        If Not mobjSquareCard Is Nothing Then
'            Call mobjSquareCard.CloseWindows
'            Set mobjSquareCard = Nothing
'        End If
'        Exit Sub
'    End If
'    '创建对象
'    '刘兴洪:增加结算卡的结算:执行或退费时
'    Err = 0: On Error Resume Next
'    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
'    If Err <> 0 Then
'        mtySquareCard.blnExistsObjects = False
'        Exit Sub
'    End If
'    Dim strKind As String
'
'    '安装了结算卡的部件
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    '功能:zlInitComponents (初始化接口部件)
'    '    ByVal frmMain As Object, _
'    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
'    '        ByVal cnOracle As ADODB.Connection, _
'    '        Optional blnDeviceSet As Boolean = False, _
'    '        Optional strExpand As String
'    '出参:
'    '返回:   True:调用成功,False:调用失败
'    '编制:刘兴洪
'    '日期:2009-12-15 15:16:22
'    'HIS调用说明.
'    '   1.进入门诊收费时调用本接口
'    '   2.进入住院结帐时调用本接口
'    '   3.进入预交款时
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    If mobjSquareCard.zlInitComponents(Me, mlngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend) Then
'        mtySquareCard.blnExistsObjects = True
'        mobjSquareCard.mblnYLMgr = mbytCallObject = 2
'    End If
'    strKind = "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;就|就诊卡|0"
'    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, mobjSquareCard, strKind, txtPatient)
'End Sub

Private Sub InitIDKind()
    Dim strKind As String
    strKind = "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;就|就诊卡|0;手|手机号|0"
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKind, txtPatient)
    mtySquareCard.blnExistsObjects = Not gobjSquare.objSquareCard Is Nothing
    gobjSquare.objSquareCard.mblnYLMgr = mbytCallObject = 2
End Sub

Private Function zlCheckDepositDelValied(ByRef lng预交ID As Long, _
    ByVal dbl退款金额 As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费交易接口
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExend As String
    Dim cllSquareBalance As Collection
    
    If mcurBill.lng卡类别ID = 0 Then zlCheckDepositDelValied = True: Exit Function
    
    If Not mtySquareCard.blnExistsObjects Or gobjSquare.objSquareCard Is Nothing Then
            MsgBox "注意:" & vbCrLf & _
                         "      当前的预交款按" & mcurBill.str名称 & " 结算的,但不存在操作的相关部件,不能退款,请与系统管理员联系!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
            Exit Function
    End If
    
    If mbytInState = 3 And mcurBill.bln转账 Then
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, mcurBill.lng卡类别ID, False, Nvl(mrsInfo!姓名), Nvl(mrsInfo!性别), Nvl(mrsInfo!年龄), dbl退款金额, mstrBrushCardNo, mstrbrPassWord, False, False, False, False) = False Then Exit Function
        mcurBill.str卡号 = mstrBrushCardNo
        zlXML.ClearXmlText
        zlXML.AppendNode "IN"
            zlXML.appendData "CZLX", "4"
        zlXML.AppendNode "IN", True
        strXMLExend = zlXML.XmlText
        zlXML.ClearXmlText
        If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, mcurBill.lng卡类别ID, _
            mcurBill.str卡号, dbl退款金额, "", strXMLExend) = False Then
            zlCheckDepositDelValied = False
            Exit Function
        End If
    Else
        Set cllSquareBalance = New Collection
        'Array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文,剩余未退金额)
        cllSquareBalance.Add Array(mcurBill.lng卡类别ID, mcurBill.lng消费卡ID, 0, mcurBill.str卡号, "", "", False, dbl退款金额)
        'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
        ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, ByVal strSwapNo As String, _
        ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:帐户回退交易前的检查
        '入参:frmMain-调用的主窗体
        '       lngModule-调用的模块号
        '       lngCardTypeID-卡类别ID
        '       strCardNo-卡号
        '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
        '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
        '       dblMoney-退款金额
        '       strSwapNo-交易流水号(退款时检查)
        '       strSwapMemo-交易说明(退款时传入)
        '       strXMLExpend    XML IN  可选参数(扩展用).暂未传入
        '返回:退款合法,返回true,否则返回Flase
        If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, mcurBill.lng卡类别ID, mcurBill.bln消费卡, mcurBill.str卡号, _
            "1|" & lng预交ID, dbl退款金额, mcurBill.str交易流水号, mcurBill.str交易说明, strXMLExend) = False Then
              zlCheckDepositDelValied = False
              Exit Function
         End If
         '100610:李南春,2016/10/13，预交退款和余额退款是否验证刷卡
         If mcurBill.bln消费卡 = False And mcurBill.bln退款验卡 _
            Or mcurBill.bln消费卡 And gbln消费卡退费验卡 Then
            '   zlBrushCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal rsClassMoney As ADODB.Recordset, _
            ByVal lngCardTypeID As Long, _
            ByVal bln消费卡 As Boolean, _
            ByVal strPatiName As String, ByVal strSex As String, _
            ByVal strOld As String, ByRef dbl金额 As Double, _
            Optional ByRef strCardNo As String, _
            Optional ByRef strPassWord As String, _
            Optional ByRef bln退费 As Boolean = False, _
            Optional ByRef blnShowPatiInfor As Boolean = False, _
            Optional ByRef bln退现 As Boolean = False, _
            Optional ByVal bln余额不足禁止 As Boolean = True, _
            Optional ByRef varSquareBalance As Variant, _
            Optional ByVal bln转预交 As Boolean = False, _
            Optional ByVal blnAllPay As Boolean = False, _
            Optional ByVal strXmlIn As String = "") As Boolean
            '       strXmlIn-三方卡调用XML入参,目前格式如下:
            '       <IN>
            '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
            '       </IN>
            
            If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, mcurBill.lng卡类别ID, mcurBill.bln消费卡, _
                Trim(txtPatient.Text), mstrPatiSex, mstrPatiOld, dbl退款金额, mstrBrushCardNo, mstrbrPassWord, _
                True, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
            mcurBill.str卡号 = mstrBrushCardNo
        End If
    End If
     
goEnd:
    zlCheckDepositDelValied = True
    Exit Function
End Function

Private Function zlDepositDel(ByRef lng预交ID As Long, ByRef lng冲预交ID As Long, ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：退预交交易
    '入参： lng预交ID-预交ID
    '返回：成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dblCurMoney As Double, dblMoneySum As Double
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strXMLExpend As String, str预交IDs As String
    
    Err = 0: On Error GoTo Errhand:
    If mcurBill.lng卡类别ID = 0 Then zlDepositDel = True: Exit Function
    
    If mcurBill.bln消费卡 Then
        '冲销消费卡金额
        strSQL = _
            "Select 接口编号, 消费卡id, 卡号, -1 * Sum(应收金额) As 应收金额" & vbNewLine & _
            "From 病人卡结算记录" & vbNewLine & _
            "Where 记录性质 = 4 And 结算id = [1]" & vbNewLine & _
            "Group By 接口编号, 消费卡id, 卡号"
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng预交ID)
        
        '可能使用了多张消费卡
        dblMoneySum = dblMoney
        Do While Not rsTemp.EOF
             If Val(Nvl(rsTemp!应收金额)) < dblMoneySum Then
              dblCurMoney = Val(Nvl(rsTemp!应收金额))
              dblMoneySum = Round(dblMoneySum - Val(Nvl(rsTemp!应收金额)), 6)
            Else
              dblCurMoney = dblMoneySum
              dblMoneySum = 0
            End If
            
            'Zl_病人卡结算记录_退款
            strSQL = "Zl_病人卡结算记录_退款("
            '  接口编号_In   消费卡类别目录.编号%Type,
            strSQL = strSQL & "" & Val(Nvl(rsTemp!接口编号)) & ","
            '  卡号_In       消费卡信息.卡号%Type,
            strSQL = strSQL & "'" & Nvl(rsTemp!卡号) & "',"
            '  消费卡id_In   消费卡信息.Id%Type,
            strSQL = strSQL & "" & Val(Nvl(rsTemp!消费卡ID)) & ","
            '  结算金额_In   病人卡结算记录.应收金额%Type,
            strSQL = strSQL & "" & dblCurMoney & ","
            '  原预交id_In   病人卡结算记录.结算id%Type,
            strSQL = strSQL & "" & lng预交ID & ","
            '  新预交id_In   病人卡结算记录.结算id%Type,
            strSQL = strSQL & "" & lng冲预交ID & ","
            '  操作员编号_In 病人卡结算记录.操作员编号%Type,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 病人卡结算记录.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  退款时间_In   病人预交记录.收款时间%Type
            strSQL = strSQL & "To_Date('" & Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'))"

            zldatabase.ExecuteProcedure strSQL, Me.Caption
            
            If dblMoneySum = 0 Then Exit Do
            rsTemp.MoveNext
        Loop
        If dblMoneySum > 0 Then
            MsgBox "剩余可退金额(" & Format(dblMoney - dblMoneySum, "0.00") & ")不足本次退款金额(" & _
                Format(dblMoney, "0.00") & ")，不能退费！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If mbytInState = 3 And mcurBill.bln转账 Then
        strXMLExpend = "<IN><CZLX>4</CZLX></IN>"
        strSwapNO = mcurBill.str交易流水号: strSwapMemo = mcurBill.str交易说明
        strSwapExtendInfor = "1|" & lng冲预交ID
        If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModul, mcurBill.lng卡类别ID, mcurBill.str卡号, _
            lng预交ID, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then Exit Function
    Else
        'Public Function zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long,bln消费卡 as boolean ByVal strCardNo As String, _
            ByVal strBalanceIDs As String, _
            ByVal dblMoney As Double, _
            ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
            ByRef strSwapExtendInfor As String) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:帐户扣款回退交易
            '入参:frmMain-调用的主窗体
            '       lngModule-调用的模块号
            '       lngCardTypeID-卡类别ID:医疗卡类别.ID
            '       strCardNo-卡号
            '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
            '                           格式:收费类型(|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
            '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
            '       dblMoney-退款金额
            '       strSwapNo-交易流水号(扣款时的交易流水号)
            '       strSwapMemo-交易说明(扣款时的交易说明)
            '       strSwapExtendInfor-本次退费的冲销ID：
            '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
            '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
            '出参: strSwapNo-交易流水号(退款交易流水号)
            '         strSwapMemo-交易说明(退款交易说明)
            '       strSwapExtendInfor-交易的扩展信息
            '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
            '返回:函数返回    True:调用成功,False:调用失败
         strSwapNO = mcurBill.str交易流水号: strSwapMemo = mcurBill.str交易说明
         '81489,冉俊明,2015-4-29,退费传入冲销ID
         strSwapExtendInfor = "1|" & lng冲预交ID
         If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, mcurBill.lng卡类别ID, mcurBill.bln消费卡, mcurBill.str卡号, _
            "1|" & lng预交ID, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    End If
    '127450:李南春,2018/6/20，余额退款时，需要获取对应的冲预交记录
    If mbytInState = 3 Then
        strSQL = "Select ID From 病人预交记录 Where 记录性质 In (1, 11) And NO = (Select NO from 病人预交记录 Where ID = [1])"
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "余额退款记录", lng冲预交ID)
        Do While Not rsTemp.EOF
            str预交IDs = str预交IDs & "," & rsTemp!ID
            rsTemp.MoveNext
        Loop
    End If
    If str预交IDs <> "" Then
        str预交IDs = Mid(str预交IDs, 2)
    Else
        str预交IDs = lng冲预交ID
    End If
    If Save三方交易(str预交IDs, mcurBill.lng卡类别ID, mcurBill.bln消费卡, mcurBill.str卡号, strSwapNO, strSwapMemo, _
        strSwapExtendInfor, True, "1|" & lng冲预交ID) = False Then Exit Function
    zlDepositDel = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 

Private Sub Load支付方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:刘兴洪
    '日期:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim str性质 As String
    
                        
    '结算方式:费用查询和医疗卡调用时，一般只支付预交款,不存在代收的情况
    'mbytCallObject:调用的对象(0-预交款管理调用;1-病人费用查询调用;2-医疗卡...
    If InStr(1, mstrPrivs, ";预交收款;") > 0 Or _
        InStr(1, mstrPrivs, ";预交收款;") > 0 Or _
        InStr(1, mstrPrivs, ";预交结清退款;") > 0 Or _
        InStr(1, mstrPrivs, ";门诊预交转住院;") > 0 _
        Or InStr(1, mstrPrivs, ";住院预交转门诊;") > 0 Or mbytCallObject > 0 Then
        str性质 = ",1,2,7,8,3"
    End If
    '只有代收款权限时,不能处理其他性质的预交款
    '问题:45471
    If InStr(1, mstrPrivs, ";代收款退款;") > 0 Or InStr(1, mstrPrivs, ";代收款收取;") > 0 Then
        If mbytCallObject = 0 Then str性质 = str性质 & ",5"
    End If
    If str性质 = "" Then str性质 = ",1,2,7,8,3"
    str性质 = Mid(str性质, 2)
    
    If mblnNurseCall Then
        str性质 = "7,8"
    End If
    
    On Error GoTo errHandle
    Set rsTemp = Get结算方式("预交款", str性质)
    Set mcolPayMode = New Collection
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType: varData = Split(strPayType, ";")
    With cboStyle
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If rsTemp!性质 = 3 And InStr(mstrPrivs, ";保险转帐;") > 0 Then
                mstr个人帐户 = rsTemp!编码 & "-" & rsTemp!名称 '根据病人动态加入
            End If
            '104083:李南春，2016/12/21，个人账户放在最后动态加入
            '性质为8的根据启用医疗卡来处理
            If Not blnFind And InStr(",3,8,", "," & rsTemp!性质 & ",") = 0 Then
                .AddItem Nvl(rsTemp!名称)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!性质))
                mcolPayMode.Add Array("", Nvl(rsTemp!名称), 0, 0, 0, 0, Nvl(rsTemp!名称), 0, 0), "K" & j
                If rsTemp!缺省 = 1 Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                If mstr缺省结算方式 = Nvl(rsTemp!名称) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        For i = 0 To UBound(varData)
            '问题号:116175，焦博，2017/12/8，将医疗卡的缴款方式控制调整为受结算方式管理和设备启用共同控制
            rsTemp.Filter = "名称 ='" & Split(varData(i), "|")(6) & "'"
            If Not rsTemp.EOF Then
                If InStr(1, varData(i), "|") <> 0 And str性质 <> 5 Then
                    varTemp = Split(varData(i), "|")
                    mcolPayMode.Add varTemp, "K" & j
                    .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                    If mstr缺省结算方式 = varTemp(1) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                    j = j + 1
                End If
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    If cboStyle.ListCount = 0 Then
        MsgBox "预交场合没有可用的结算方式,请先到结算方式管理中设置。", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Save三方交易(ByVal str预交IDs As String, _
    ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal str卡号 As String, str交易流水号 As String, str交易说明 As String, _
    strExpend As String, Optional bln退预交 As Boolean = False, Optional strExpendOld As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存三方结算数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSQL As String, varData As Variant, varTemp As Variant, cllPro As Collection, i As Long
     
    Err = 0: On Error GoTo Errhand:
    If bln退预交 = False Then
        '退费时,不更改交易
        '更新交易信息
        '    Zl_三方接口更新_Update
        strSQL = "Zl_三方接口更新_Update("
        '  卡类别id_In   病人预交记录.卡类别id%Type,
        strSQL = strSQL & "" & lng卡类别ID & ","
        '  消费卡_In     Number,
        strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
        '  卡号_In       病人预交记录.卡号%Type,
        strSQL = strSQL & "'" & str卡号 & "',"
        '  结帐ids_In    Varchar2,
        strSQL = strSQL & "'" & str预交IDs & "',"
        '  交易流水号_In 病人预交记录.交易流水号%Type,
        strSQL = strSQL & "'" & str交易流水号 & "',"
        '  交易说明_In   病人预交记录.交易说明%Type
        strSQL = strSQL & "'" & str交易说明 & "',"
        '预交款缴款_In Number := 0
        strSQL = strSQL & "" & 1 & ","
        '退费标志 :1-退费;0-付费
        strSQL = strSQL & "" & IIf(bln退预交, 1, 0) & ")"
        Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    gcnOracle.CommitTrans
    '先提交,这样避免风险,再更新相关的交易信息
    'strExpend:交易扩展信息,格式:项目名称|项目内容||...
    varData = Split(strExpend, "||")
    Dim str交易信息 As String, strTemp As String
    Set cllPro = New Collection
    If strExpendOld <> strExpend Then
        For i = 0 To UBound(varData)
            If Trim(varData(i)) <> "" Then
                varTemp = Split(varData(i) & "|", "|")
                If varTemp(0) <> "" Then
                    strTemp = varTemp(0) & "|" & varTemp(1)
                    If zlCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
                        str交易信息 = Mid(str交易信息, 3)
                        'Zl_三方结算交易_Insert
                        strSQL = "Zl_三方结算交易_Insert("
                        '卡类别id_In 病人预交记录.卡类别id%Type,
                        strSQL = strSQL & "" & lng卡类别ID & ","
                        '消费卡_In   Number,
                        strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
                        '卡号_In     病人预交记录.卡号%Type,
                        strSQL = strSQL & "'" & str卡号 & "',"
                        '结帐ids_In  Varchar2,
                        strSQL = strSQL & "'" & str预交IDs & "',"
                        '交易信息_In Varchar2:交易项目|交易内容||...
                        strSQL = strSQL & "'" & str交易信息 & "',"
                        '预交款缴款_In Number := 0
                        strSQL = strSQL & "1)"
                        zlAddArray cllPro, strSQL
                        str交易信息 = ""
                    End If
                    str交易信息 = str交易信息 & "||" & strTemp
                End If
            End If
        Next
        
        If str交易信息 <> "" Then
            str交易信息 = Mid(str交易信息, 3)
            'Zl_三方结算交易_Insert
            strSQL = "Zl_三方结算交易_Insert("
            '卡类别id_In 病人预交记录.卡类别id%Type,
            strSQL = strSQL & "" & lng卡类别ID & ","
            '消费卡_In   Number,
            strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
            '卡号_In     病人预交记录.卡号%Type,
            strSQL = strSQL & "'" & str卡号 & "',"
            '结帐ids_In  Varchar2,
            strSQL = strSQL & "'" & str预交IDs & "',"
            '交易信息_In Varchar2:交易项目|交易内容||...
            strSQL = strSQL & "'" & str交易信息 & "',"
            '预交款缴款_In Number := 0
            strSQL = strSQL & "1)"
            zlAddArray cllPro, strSQL
        End If
    End If
    Err = 0: On Error GoTo ErrOthers:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    Save三方交易 = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrOthers:
    '    能保存多少,作多少
     Call ErrCenter
End Function


Private Function zlInterfacePrayMoney(ByVal lng预交ID As Long, ByVal strNo As String, ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接口支付金额
    '返回:支付成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    
    If cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    If mlngCardTypeID = 0 Then zlInterfacePrayMoney = True: Exit Function
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, mlngCardTypeID, mbln消费卡, mstrBrushCardNo, "", strNo, _
        dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    If Save三方交易(lng预交ID, mlngCardTypeID, mbln消费卡, mstrBrushCardNo, strSwapGlideNO, strSwapMemo, _
        strSwapExtendInfor) = False Then Exit Function
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub zlCheckFactIsEnough()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前票据是否允足
    '编制:刘兴洪
    '日期:2012-09-06 15:41:52
    '说明:
    '问题:37372
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng剩余数量 As Long, strType As String
    If mbytInState = 1 Or mbytInState = 2 Then Exit Sub
    '需要检查剩余数量是否充足:
    If cboType.ListIndex < 0 Then
        strType = ""
    Else
        strType = cboType.ItemData(cboType.ListIndex)
    End If
    If zlCheckInvoiceOverplusEnough(2, gint提醒剩余票据张数, lng剩余数量, mlng领用ID, strType) = False Then
        MsgBox "注意:" & vbCrLf & _
               "    当前剩余票据(" & lng剩余数量 & ") 小于了报警的张数(" & gint提醒剩余票据张数 & "),请注意更换发票!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
    End If
End Sub
Private Sub LoadPatiPage(ByVal lng病人ID As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人的住院次数
    '编制:刘兴洪
    '日期:2012-12-11 10:19:58
    '说明:
    '问题:51628
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim bln留观 As Boolean
    On Error GoTo errHandle
        
    cboPatiPage.Clear
    strSQL = "" & _
    "   Select 主页ID,病人性质,入院日期,出院日期  " & _
    "   From 病案主页" & _
    "   Where 病人ID=[1]  " & _
    "   Order By Nvl(主页ID,0) Desc"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    
    With cboPatiPage
        Do While Not rsTemp.EOF
            If bln留观 = False And Val(Nvl(rsTemp!病人性质, 0)) <> 0 Then bln留观 = True
            If Val(Nvl(rsTemp!主页ID)) = 0 And Val(Nvl(rsTemp!病人性质)) = 0 Then
                .AddItem "预约入院"
            Else
                .AddItem "第" & rsTemp!主页ID & "次" & IIf(Val("" & rsTemp!病人性质) = 1, "(门诊留观)", IIf(Val("" & rsTemp!病人性质) = 2, "(住院留观)", ""))
            End If
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!主页ID))
            If .ListIndex < 0 Then .ListIndex = .NewIndex
            If mblnNurseCall Then
                If Val(Nvl(rsTemp!主页ID)) = mlng主页ID Then
                    .ListIndex = .NewIndex
                End If
                cboPatiPage.Enabled = False
            Else
                If Val(Nvl(rsTemp!主页ID)) = Val(Nvl(mrsInfo!主页ID)) Then
                    .ListIndex = .NewIndex
                End If
            End If
            rsTemp.MoveNext
        Loop
        If bln留观 = True Then Call cbo.SetListWidth(cboPatiPage.hWnd, 2000)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function Check未入科不交预交() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人是否入科,未入科,不缴预交
    '编制:刘兴洪
    '日期:2012-12-11 10:19:58
    '说明:
    '问题:51628
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lng病人ID As Long, lng主页ID As Long
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If mbln未入科不交预交 = False Then Check未入科不交预交 = True: Exit Function
    '不诊预交不检查
    If cboType.ItemData(cboType.ListIndex) <> 2 Then Check未入科不交预交 = True: Exit Function
    '当前住院次数不为在院的,也不检查
    If Val(Nvl(mrsInfo!在院)) <> 1 Then Check未入科不交预交 = True: Exit Function
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    '不存在住院次数的,也能缴预交,因此不检查
    If cboPatiPage.ListIndex < 0 Then Check未入科不交预交 = True: Exit Function
    lng主页ID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    strSQL = "Select ID From 病人变动记录 Where 病人id = [1] And 主页id = [2] And 开始原因 = 2 And 床号 Is Not Null "
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "注意" & vbCrLf & "   病人『" & mrsInfo!姓名 & "』未入科,不允许缴预交款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Check未入科不交预交 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function Check退款() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人退款前金额是否存在变化
    '编制:李南春
    '日期:2016/2/25 10:21:39
    '说明:
    '问题:93144
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    Dim dbl预交余额 As Double, dbl费用余额 As Double, dbl剩余余额 As Double
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then lng病人ID = 0
    If mrsInfo.State <> 1 Then lng病人ID = 0
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    If lng病人ID <> 0 Then
        Set mrsDepositBalance = GetMoneyInfo(lng病人ID, , , , True)
        If Not mrsDepositBalance Is Nothing Then
            With mrsDepositBalance
                .Filter = "类型=" & cboType.ItemData(cboType.ListIndex)
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not .EOF
                    dbl费用余额 = dbl费用余额 + Val(Nvl(!费用余额))
                    dbl预交余额 = dbl预交余额 + Val(Nvl(!预交余额))
                    .MoveNext
                Loop
            End With
        End If
        dbl剩余余额 = dbl预交余额 - dbl费用余额
        If mdbl剩余款额 <> dbl剩余余额 Then
            MsgBox "病人的剩余款项已发生变化,请重新确定退款金额!", vbInformation + vbOKOnly, gstrSysName
            Call ShowPremayBalance(False, 0)
            txtMoney.SetFocus: Exit Function
        End If
    End If
    Check退款 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
