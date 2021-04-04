VERSION 5.00
Object = "{CC0839AF-B32F-436B-8884-BE2BB3B4C73F}#4.1#0"; "zlIDKind.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCautionMoney 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "押金单据"
   ClientHeight    =   9570
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
   Icon            =   "frmCautionMoney.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   11910
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   75
      ScaleHeight     =   2655
      ScaleWidth      =   11775
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1050
      Width           =   11775
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   20
         X1              =   6105
         X2              =   11640
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   5100
         TabIndex        =   68
         Tag             =   "身份证号 "
         Top             =   1950
         Width           =   960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   18
         X1              =   1260
         X2              =   4845
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Label lbl手机号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手 机 号"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   67
         Tag             =   "手 机 号 "
         Top             =   1950
         Width           =   960
      End
      Begin VB.Label lbl未缴费用 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未缴费用 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   2640
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   63
         Tag             =   "工作单位 "
         Top             =   1560
         Width           =   960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   16
         X1              =   1260
         X2              =   11640
         Y1              =   2550
         Y2              =   2550
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
         X1              =   1245
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
         Left            =   240
         TabIndex        =   60
         Tag             =   "备    注 "
         Top             =   2310
         Width           =   1080
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院科室 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7890
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   53
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
         TabIndex        =   52
         Tag             =   "担 保 人 "
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label lbl费别等级 
         AutoSize        =   -1  'True
         Caption         =   "费别 "
         Height          =   240
         Left            =   4965
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
         Tag             =   "性别 "
         Top             =   105
         Width           =   600
      End
      Begin VB.Label lbl押金余额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "押金余额 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   48
         Tag             =   "押金余额 "
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
         Tag             =   "未结费用 "
         Top             =   795
         Width           =   1080
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   0
      ScaleHeight     =   1650
      ScaleWidth      =   11910
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6960
      Width           =   11910
      Begin VB.CheckBox chk仅显示本次押金 
         Caption         =   "仅显示本次押金"
         Height          =   240
         Left            =   9360
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Frame Frame3 
         Caption         =   "押金清单"
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
         TabIndex        =   42
         Top             =   0
         Width           =   12015
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
         Height          =   1335
         Left            =   135
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   270
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
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   8610
      Width           =   11910
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   420
         Left            =   150
         TabIndex        =   32
         Top             =   60
         Width           =   1500
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   420
         Left            =   10335
         TabIndex        =   31
         ToolTipText     =   "热键:Esc"
         Top             =   45
         Width           =   1500
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   420
         Left            =   8760
         TabIndex        =   27
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
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   0
      Width           =   11755
      Begin VB.TextBox txtFact 
         ForeColor       =   &H00C00000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   6300
         MaxLength       =   50
         TabIndex        =   28
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
         TabIndex        =   29
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
         TabIndex        =   37
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
         TabIndex        =   54
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
         TabIndex        =   33
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
         TabIndex        =   39
         Top             =   570
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "押金单据"
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
         TabIndex        =   43
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
         TabIndex        =   38
         Top             =   630
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   34
      Top             =   9210
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
            Picture         =   "frmCautionMoney.frx":08CA
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
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3800
      Width           =   11775
      Begin VB.ComboBox cbo押金类别 
         Height          =   360
         Left            =   7995
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   3690
      End
      Begin VB.ComboBox cboPatiPage 
         Height          =   360
         Left            =   3675
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   570
         Width           =   1335
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   585
         TabIndex        =   61
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
      Begin VB.ComboBox cboType 
         Height          =   360
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   570
         Width           =   1380
      End
      Begin VB.TextBox txtMan 
         Enabled         =   0   'False
         Height          =   360
         Left            =   7980
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2760
         Width           =   3705
      End
      Begin VB.TextBox txtCode 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7995
         MaxLength       =   30
         TabIndex        =   16
         Top             =   1440
         Width           =   3690
      End
      Begin VB.TextBox txtUnit 
         Height          =   360
         Left            =   7995
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1005
         Width           =   3690
      End
      Begin VB.TextBox txt帐号 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7980
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1890
         Width           =   3705
      End
      Begin VB.ComboBox cboUnit 
         Height          =   360
         Left            =   7995
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   585
         Width           =   3690
      End
      Begin VB.ComboBox cboNote 
         Height          =   360
         Left            =   1230
         TabIndex        =   22
         Text            =   "cboNote"
         Top             =   2325
         Width           =   10485
      End
      Begin VB.TextBox txt开户行 
         Height          =   360
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1890
         Width           =   3765
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1230
         MaxLength       =   100
         TabIndex        =   1
         ToolTipText     =   "热键：F11"
         Top             =   135
         Width           =   3765
      End
      Begin VB.ComboBox cboStyle 
         Height          =   360
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1440
         Width           =   3765
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   360
         Left            =   1230
         TabIndex        =   24
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
         TabIndex        =   10
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
      Begin zlIDKind.ucQRCodePayButton btQRCodePay 
         Height          =   360
         Left            =   5010
         TabIndex        =   70
         ToolTipText     =   "扫码付允许使用快键【F6】进行快速支付"
         Top             =   1425
         Visible         =   0   'False
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   635
         Appearance      =   2
         ToolTipString   =   "扫码付允许使用快键【F6】进行快速支付"
      End
      Begin VB.Label lbl押金类别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "押金类别"
         Height          =   240
         Left            =   6960
         TabIndex        =   69
         Top             =   180
         Width           =   960
      End
      Begin VB.Label lblPatiPage 
         AutoSize        =   -1  'True
         Caption         =   "住院次数"
         Height          =   240
         Left            =   2685
         TabIndex        =   5
         Top             =   615
         Width           =   960
      End
      Begin VB.Label lblRepairMoney 
         Caption         =   "补交额:"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   5010
         TabIndex        =   62
         Top             =   1050
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label lbl押金类型 
         AutoSize        =   -1  'True
         Caption         =   "押金类型"
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   630
         Width           =   960
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缴款科室"
         Height          =   240
         Left            =   6960
         TabIndex        =   7
         Top             =   645
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "帐号"
         Height          =   240
         Left            =   7440
         TabIndex        =   19
         Top             =   1950
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开户行"
         Height          =   240
         Left            =   435
         TabIndex        =   17
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label lbl缴款单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缴款单位"
         Height          =   240
         Left            =   6960
         TabIndex        =   11
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
         TabIndex        =   9
         Top             =   1065
         Width           =   510
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算号码"
         Height          =   240
         Left            =   6960
         TabIndex        =   15
         Top             =   1500
         Width           =   960
      End
      Begin VB.Label lblStyle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "支付方式"
         Height          =   240
         Left            =   195
         TabIndex        =   13
         Top             =   1500
         Width           =   960
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要"
         Height          =   240
         Left            =   645
         TabIndex        =   21
         Top             =   2385
         Width           =   480
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收款时间"
         Height          =   240
         Left            =   195
         TabIndex        =   23
         Top             =   2820
         Width           =   960
      End
      Begin VB.Label lblMan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收款员"
         Height          =   240
         Left            =   7200
         TabIndex        =   25
         Top             =   2820
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCautionMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'说明：
'1.退款有两种方式,缺省的方式是在管理界面对指定的单据执行退款功能，或在正常收款状态下使用退款功能，另一种方式
'是以正常收款状态收款,但收款金额以负数表示退款，此时(退款金额<=病人余额)。这两种方式都不影响病人预交款的统计
Private Enum InStateType
    EM_收押金 = 0
    EM_浏览单据 = 1
    EM_退押金 = 2
    EM_异常重收 = 5
    EM_异常作废 = 6
    EM_异常重退 = 7
End Enum
'入口参数----------------------------------------------------------------------------------
Private mbytInState As Byte '0-收押金(缺省,可切换到退),1-浏览单据(1),2-退押金(1);5-异常重收,6-作废异常单据,7-异常退款重退
Private mstrInNO As String '要浏览或退款的单据号(mbytInState=1或3时有效),从病人信息登记中调用退卡时为空
Private mblnNOMoved As Boolean '显明细时记录当前选择的单据是否在在线数据表中,以其它操作时无需再判断
Private mblnViewCancel As Boolean '是否浏览退款单据(mbytInState=1时有效)
Private mstrPrivs As String
Private mlngModul As Long
Private mblnNotClick As Boolean
Private mstrbrPassWord As String
'程序变量----------------------------------------------------------------------------------
Private mblnUnLoad  As Boolean '用于控制窗体直接退出
Private mdbl剩余款额 As Double
Private mdbl预交余额 As Double
Private mdbl费用余额 As Double
Private mlng领用ID As Long, mstrCardPrivs As String
Private mstrRedFact As String
Private mstr缺省结算方式 As String
Private mblnOK As Boolean
Private mbln未入科不交预交 As Boolean '51628
Private mbln住院退预交验证 As Boolean   '63113:刘尔旋,2013-10-29,住院预交退款验证

'医保变量----------------------
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
Private mobjPayMode As Collection   '结算方式
Private mlngCardTypeID  As Long
Private mstr结算方式      As String
Private mstrBrushCardNo As String

Private Type Ty_YJInfor
    lng押金ID As Long
    strNO As String
    lng卡类别ID As Long
    str卡号 As String
    str名称 As String
    str交易流水号 As String
    str交易说明 As String
    str合作单位 As String
    dbl金额 As Double
    bln退款验卡 As Boolean
    dt收款时间 As Date
End Type
Private mYJinfo As Ty_YJInfor
Private mFactProperty As Ty_FactProperty
Private mblnStartFactUseType As Boolean '是否启用的相关的使用类别的
Private mrsDepositBalance As ADODB.Recordset    '当前病人的预交余额
Private mbytBackMoneyType As Byte '退款方式:1-禁止;0-提示
Private mbytOracleBackType As Byte '退款检查_In;0-忽略退款金额是否大于了病人余额；1-检查退款金额
Private mblnClearWinInfor As Boolean  '缴款后,是否清除窗体信息
Private mblnCheckPass As Boolean '刷卡时要求输入密码,'0000000000'依位顺序表示各个场合,分别为:1.门诊挂号,2.门诊划价,3.门诊收费,4.门诊记帐,5.入院登记,6.住院记帐,7.病人结帐,8.病人预交款,9.检验技师站,10.影像医技站.'
'外挂评价器对象
Private mobjPlugIn As Object
Private mstrPatiOld As String
Private mstrPatiSex As String
Private mlngFactModule As Long '发票相关参数模块号
Private mblnOptErrBill As Boolean '收费模式下处理异常单据
Private mbln排除未缴及未审 As Boolean '剩余款排除未缴及未审金额
Private mstrQRcode  As String    '扫码支付接口返回的二维码串
Private mstr退款操作员 As String
Private mblnCheckSwapFailed As Boolean '异常重结时,是否检查交易失败了(zlSwapIsSucces)
                                                                 'True-检查交易失败；False-检查交易成功
Private mbnQRPay   As Boolean  '是否是扫码付款
Private mpatiInfo As New clsPatientInfo 'zlOneCardComLib.clsPatientInfo

Public Function zlShowEdit(ByVal frmMain As Object, ByVal bytInState As Byte, _
                                          ByVal strPrivs As String, ByVal lngModule As Long, _
                                          Optional strInNo As String = "", _
                                          Optional ByVal blnViewCancel As Boolean = False, _
                                          Optional blnNOMoved As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口,用于病人押金信息编辑或查看
    '入参:frmMain-调用的主窗口
    '        bytInState:0-收押金(缺省,可切换到退),1-浏览单据(1),2-退押金(1)
    '        strInNo:要浏览或退款的单据号(mbytInState=1或3时有效)
    '         blnViewCancel:是否浏览退款单据(mbytInState=1时有效)
    '        blnNOMoved:显明细时记录当前选择的单据是否在在线数据表中,以其它操作时无需再判断
    '出参:
    '返回:押金只有一次成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-17 16:11:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    mbytInState = bytInState: mstrPrivs = strPrivs: mlngModul = lngModule
    mstrInNO = strInNo: mblnViewCancel = blnViewCancel: mblnNOMoved = blnNOMoved
    mlngFactModule = mlngModul
    mblnOK = False
    If frmMain Is Nothing Then
        frmCautionMoney.Show
    Else
        frmCautionMoney.Show 1, frmMain
    End If
    zlShowEdit = mblnOK
End Function

Private Sub btQRCodePay_zlErrShow(ByVal strErrMsg As String, ByVal lngErrNum As Long)
    Call RestorePayStyle '恢复上次选择项
    If strErrMsg = "" Then Exit Sub
    MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
End Sub

Private Sub btQRCodePay_zlGetPayMoney(dblMoney As Double, strExpend As String, blnCancel As Boolean)
    Err = 0: On Error GoTo errHandle:
    
    If Not (mbytInState = EM_收押金 Or mbytInState = EM_异常重收) Then blnCancel = True: Exit Sub

    lblStyle.Tag = cboStyle.ListIndex     '先记录当前选择的支付方式
    '定位到指定卡类别
    If btQRCodePay.Tag = "" Then
        MsgBox "未找到有效的扫码付类别,请检查!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True
        Exit Sub
    End If
    
    '重新处理数据
    dblMoney = StrToNum(txtMoney.Text)

    If dblMoney <> 0 Then
        txtMoney.Text = Format(dblMoney, "0.00")
    End If
    
    If dblMoney < 0 Then
        MsgBox "当前为退款，扫码付不支持退款操作!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True
        Exit Sub
    End If
    
    If dblMoney = 0 Then
        MsgBox "未输入本次应缴金额，不需要进行扫码付款，请输入金额后再扫码付!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True
        zlControl.ControlSetFocus txtMoney
        Exit Sub
    End If
    If CheckDataValied = False Then blnCancel = True:  Exit Sub
    If Not Check未入科不交预交 Then blnCancel = True:  Exit Sub
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    blnCancel = True
End Sub

Private Sub btQRCodePay_zlQRCodePayment(ByVal lngCardTypeID As Long, ByVal strPayMentQRCode As String, ByVal strExpendXML As String, blnCancel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行扫码付款
    '入参:lngCardTypeID-卡类别ID
    '     strPayMentQRCode-二维码付款内码
    '     strExpendXML-暂无
    '出参:strExpendXML-暂无
    '     blnCancel-true表示取消本次扫码付,False-表示本次扫码付成功
    '编制:刘兴洪
    '日期:2019-03-07 11:34:19
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle

    If lngCardTypeID = 0 Or blnCancel Then
        blnCancel = True
        Call RestorePayStyle '恢复上次选择的支付方式
        Exit Sub
    End If

    blnCancel = False
    If LocatePayStyle(lngCardTypeID) = False Then  '定位到扫码付的指定类别上
        blnCancel = True
        MsgBox "不能有效识别当前扫码付的类别，可能本机不支持该类别的扫码付，请与管理员联系！", vbInformation + vbOKOnly, gstrSysName
        Call RestorePayStyle '恢复上次选择的支付方式
        Exit Sub
    End If
    mstrQRcode = strPayMentQRCode
    mbnQRPay = True
    Call cmdOK_Click
    mbnQRPay = False
    mstrQRcode = ""
    If Not mblnClearWinInfor Then Call RestorePayStyle  '恢复上次选择的支付方式

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    blnCancel = True
    Call RestorePayStyle '恢复上次选择的支付方式
End Sub

Private Sub cboPatiPage_Click()
    If txtPatient.Tag <> "" And mbytInState = 0 And mpatiInfo.病人ID > 0 Then
        If cboPatiPage.ItemData(cboPatiPage.ListIndex) <> mpatiInfo.主页ID Then
            Call ShowPatiPageInfo
        End If
    End If
    Call ShowHistoryPrepay("")
End Sub

Private Sub ShowPatiPageInfo()
    Dim lng主页ID As Long
    lng主页ID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    '根据第几次入院更新信息
    Call GetPatient(IDKind.GetfaultCard, txtPatient.Tag, False, False, txtPatient.Tag, lng主页ID)
    If mpatiInfo.病人ID > 0 Then Exit Sub
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(mpatiInfo.住院号 = "", "", "住院号:" & mpatiInfo.住院号 & "   ") & _
                       IIf(mpatiInfo.门诊号 = "", "", "门诊号:" & mpatiInfo.门诊号)
    lbl费别等级.Caption = lbl费别等级.Tag & mpatiInfo.费别
    txtPatient.Text = mpatiInfo.姓名
    txtPatient.Tag = mpatiInfo.病人ID
    lblSex.Caption = lblSex.Tag & mpatiInfo.性别
    lblOld.Caption = lblOld.Tag & mpatiInfo.年龄
    lbl医疗付款方式.Caption = lbl医疗付款方式.Tag & mpatiInfo.医疗付款方式
    lbl科室.Caption = lbl科室.Tag & GET部门名称(mpatiInfo.出院科室ID)
    lbl床号.Caption = lbl床号.Tag & IIf(mpatiInfo.床号 = "", "家庭", mpatiInfo.床号)
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(mpatiInfo.当前科室ID = 0, mpatiInfo.出院科室ID, mpatiInfo.当前科室ID))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    Call Load医保预结(mpatiInfo.病人ID, lng主页ID)
End Sub

Private Sub cboPatiPage_KeyDown(KeyCode As Integer, Shift As Integer)
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboType_Click()
    If cboType.ListIndex < 0 Then Exit Sub
    
    '88657:李南春，2015/9/17,切换预交类型刷新预交余额
    If mbytInState = EM_收押金 And chkCancel.Value = 0 Or mbytInState = EM_异常重收 Then
        mlng领用ID = 0
        '问题号:112784,焦博,2017/10/13,获取正确的票据格式
        mFactProperty = zl_GetInvoicePreperty(mlngFactModule, 21, cboType.ItemData(cboType.ListIndex))
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact
        Call ShowPremayBalance(True, 0)
        Call SetCtrlEnabled
        Call ShowHistoryPrepay("")
    ElseIf mbytInState = EM_退押金 Or chkCancel.Value = 1 Then
        mlng领用ID = 0
        '问题号:112784,焦博,2017/10/13,获取正确的票据格式
        mFactProperty = zl_GetInvoicePreperty(mlngFactModule, 22, cboType.ItemData(cboType.ListIndex))
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact(False, True)
    End If
    
     '问题号:45666
    If mbytInState = EM_收押金 And cboType.Text = "住院押金" Then
        chk仅显示本次押金.Visible = True
        chk仅显示本次押金.Value = IIf(zlDatabase.GetPara("仅显示本次预交", glngSys, mlngModul, , Array(chk仅显示本次押金), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
    Else
        chk仅显示本次押金.Visible = False
    End If
    lblPatiPage.Visible = cboType.Text = "住院押金": cboPatiPage.Visible = cboType.Text = "住院押金"
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo押金类别_Click()
    If mblnNotClick Then Exit Sub
    If Not (mbytInState = EM_异常重收 Or mbytInState = EM_收押金) Then Exit Sub
    With cbo押金类别
        txtMoney.Text = ""
        If Val(.ItemData(.ListIndex)) > 0 Then
            txtMoney.Text = Format(Val(.ItemData(.ListIndex)), "###0.00;-###0.00;;")
        End If
    End With
End Sub

Private Sub cbo押金类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Function zlThirdReturnCashCheck(Optional ByRef blnChange As Boolean) As Boolean
    '功能:三方卡退现检查
    Dim dblMoney As Double, strTKList As String
    Dim strBalanceIDs As String, strXMLExpend As String
    Dim strValue As String, bln允许退现 As Boolean
    Dim int退现状态 As Integer, str缺省退现方式 As String
    
    On Error GoTo errHandle
    If cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then zlThirdReturnCashCheck = True: Exit Function
    If mYJinfo.lng卡类别ID = 0 Or mYJinfo.str名称 = "" Then zlThirdReturnCashCheck = True: Exit Function
    cboStyle.Enabled = False: cboStyle.Locked = True
    dblMoney = roundEx(mYJinfo.dbl金额, 6)
    strBalanceIDs = "8" & "|" & mYJinfo.lng押金ID
    
    strTKList = strTKList & Space(8) & "<TK>" & vbCrLf
    strTKList = strTKList & Space(8) & "    <TKFS>" & mYJinfo.str名称 & "</TKFS>" & vbCrLf
    strTKList = strTKList & Space(8) & "    <TKJE>" & mYJinfo.dbl金额 & "</TKJE>" & vbCrLf
    strTKList = strTKList & Space(8) & "    <JYLSH>" & mYJinfo.str交易流水号 & "</JYLSH>" & vbCrLf
    strTKList = strTKList & Space(8) & "    <JYSM>" & mYJinfo.str交易说明 & "</JYSM>" & vbCrLf
    strTKList = strTKList & Space(8) & "    <KH>" & mYJinfo.str卡号 & "</KH>" & vbCrLf
    strTKList = strTKList & Space(8) & "</TK>" & vbCrLf
        
    strXMLExpend = "<INPUT>" & vbCrLf
    strXMLExpend = strXMLExpend & "    <TKLIST>" & vbCrLf
    strXMLExpend = strXMLExpend & strTKList
    strXMLExpend = strXMLExpend & "    </TKLIST>" & vbCrLf
    strXMLExpend = strXMLExpend & "</INPUT>"

    bln允许退现 = gobjSquare.objSquareCard.zlReturnCashCheck(Me, mlngModul, mYJinfo.lng卡类别ID, mYJinfo.str卡号, _
                          strBalanceIDs, dblMoney, mYJinfo.str交易流水号, mYJinfo.str交易说明, strXMLExpend)
    If zlXML_Init() Then
        If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) Then
            Call zlXML_GetNodeValue("TXZT", , strValue): int退现状态 = Val(strValue)
            Call zlXML_GetNodeValue("QSTKFS", , strValue): str缺省退现方式 = Nvl(strValue)
        End If
    End If
    '接口返回为True-允许退现.
    If bln允许退现 Then
        blnChange = True
        Call Load支付方式(True) '加载性质为1,2的结算方式
        If int退现状态 = 1 Then  '缺省退现
            Call LoadOriginReturnMoneyStyle(True) '加载原始退款方式
            cboStyle.ListIndex = cbo.FindIndex(cboStyle, str缺省退现方式, True)
        Else                               '允许退现
            Call LoadOriginReturnMoneyStyle '加载原始退款方式
        End If
        zlThirdReturnCashCheck = True: Exit Function
    End If
    
    '接口返回为False-允许通过“强制退现”权限来退现.
    If int退现状态 = 1 Then    '允许强制退现
        '有强制退现权限
        If InStr(";" & mstrCardPrivs & ";", ";三方退款强制退现;") > 0 Then
            blnChange = True
            Call Load支付方式(True)                 '加载性质为1,2的结算方式
            Call LoadOriginReturnMoneyStyle '加载原始退款方式
            zlThirdReturnCashCheck = True: Exit Function
        End If
        
        '没有强制退现权限
        mstr退款操作员 = zlDatabase.UserIdentifyByUser(Me, "强制退现验证", glngSys, 1151, "三方退款强制退现")
        If mstr退款操作员 = "" Then
            MsgBox "录入的操作员验证失败或者录入的操作员不具备强制退现权限，不能强制退现！ " & vbCrLf & _
                         "如果要强制退现，请让其他具备”强制退现“的操作员操作。", vbInformation, gstrSysName
        Else
            blnChange = True
            Call Load支付方式(True)                 '加载性质为1,2的结算方式
            Call LoadOriginReturnMoneyStyle '加载原始退款方式
        End If
    End If
    zlThirdReturnCashCheck = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub chk仅显示本次押金_Click()
    Call ShowHistoryPrepay("")
End Sub

Private Sub IDKind_Click(objCard As Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXml As String
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
               Set mobjICCard = New clsICCard
               Call mobjICCard.SetParent(Me.hwnd)
               Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text = "" Then Exit Sub
        Call FindPati(objCard, False, txtPatient.Text)
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text = "" Then Exit Sub
    Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As Card)
    Call txtPatient_GotFocus
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    zlControl.ControlSetFocus txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As Card, objPatiInfor As clsPatientInfo, blnCancel As Boolean)
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text = "" Then Exit Sub
    Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC卡", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("身份证", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub
Private Sub SetcmdOkEnabled()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置cmdOk的neable属性
    '编制：刘兴洪
    '日期：2010-07-09 16:24:53
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    If mpatiInfo.病人ID = 0 Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
    chk仅显示本次押金.Enabled = cmdOK.Enabled
End Sub

Private Sub SetCtrlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Enabled属性
    '编制:刘兴洪
    '日期:2011-07-24 09:30:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean, objCtl As Control
    Dim int性质 As Integer
    blnEdit = mbytInState <> EM_浏览单据
    If cboStyle.ListIndex >= 0 Then int性质 = cboStyle.ItemData(cboStyle.ListIndex)
    Select Case mbytInState
    Case EM_收押金
        If chkCancel.Value = Checked Then GoTo goEnd:
        blnEdit = True
        cbo押金类别.Enabled = blnEdit
        cboType.Enabled = blnEdit
        cboUnit.Enabled = blnEdit
        txtUnit.Enabled = blnEdit And int性质 = 2
        cboStyle.Enabled = blnEdit
        txtCode.Enabled = blnEdit And int性质 = 2
        txt开户行.Enabled = blnEdit And int性质 = 2
        txt帐号.Enabled = blnEdit And int性质 = 2
        cboNote.Enabled = blnEdit
        picNO.Enabled = blnEdit
        cboPatiPage.Enabled = blnEdit
        txtPatient.Enabled = blnEdit
        txtMoney.Enabled = blnEdit
     Case EM_异常重收  '异常重收
        blnEdit = mblnCheckSwapFailed
        cbo押金类别.Enabled = False
        cboType.Enabled = False
        cboUnit.Enabled = blnEdit
        txtUnit.Enabled = blnEdit And int性质 = 2
        cboStyle.Enabled = blnEdit
        txtCode.Enabled = blnEdit And int性质 = 2
        txt开户行.Enabled = blnEdit And int性质 = 2
        txt帐号.Enabled = blnEdit And int性质 = 2
        cboNote.Enabled = blnEdit
        cboPatiPage.Enabled = False
        txtPatient.Enabled = False
        txtMoney.Enabled = blnEdit
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
    Dim strInfo As String
    Dim strStyle As String
    If mbytInState = EM_退押金 Or chkCancel.Value = 1 Then Exit Sub
    
    If cboStyle.ListIndex = -1 Then Exit Sub
        
    '问题号:111657,焦博,2017/07/25,使用现金支付预交款时,任会产生三方卡号
    mstrBrushCardNo = ""     '清空三方交易时缓存的卡号
    mYJinfo.lng押金ID = 0
    strStyle = "_" & cboStyle.List(cboStyle.ListIndex)
  
   If Not mobjPayMode Is Nothing Then
        If CollectionExitsValue(mobjPayMode, strStyle) Then
            mlngCardTypeID = mobjPayMode(strStyle).接口序号
            mstr结算方式 = mobjPayMode(strStyle).结算方式
        End If
        Call ShowPremayBalance(False, 0)
    End If
    Call SetCtrlEnabled
    Select Case cboStyle.ItemData(cboStyle.ListIndex)
    Case 1
        txtUnit.Text = "": txt开户行.Text = "": txt帐号.Text = "": txtCode.Text = ""
    Case 2
        If cboStyle.Text Like "*票*" Or cboStyle.Text Like "*卡*" Then
            '无支票这种结算性质,所以用名称
            '问题:36611
            If mpatiInfo.病人ID = 0 Then Exit Sub
            strInfo = GetLastInfo(mpatiInfo.病人ID)
            If strInfo <> "" Then
                txtUnit.Text = IIf(Split(strInfo, "|")(0) = "", txtUnit.Text, Split(strInfo, "|")(0))
                txt开户行.Text = IIf(Split(strInfo, "|")(1) = "", txt开户行.Text, Split(strInfo, "|")(1))
                txt帐号.Text = IIf(Split(strInfo, "|")(2) = "", txt帐号.Text, Split(strInfo, "|")(2))
                txtCode.Text = IIf(Split(strInfo, "|")(3) = "", txtCode.Text, Split(strInfo, "|")(3))
            End If
        End If
    Case -1
        If CheckParaConfig(mlngCardTypeID) = False Then
            mblnUnLoad = mbytInState = EM_异常重收: Exit Sub
        End If
        If CCur(StrToNum(txtMoney.Text)) < 0 And mbytInState = EM_收押金 Then
            MsgBox "三方卡不允许输入负数！", vbInformation, gstrSysName
            txtMoney.Text = "": zlControl.ControlSetFocus txtMoney
        End If
    End Select
End Sub

Private Function CheckParaConfig(ByVal lngCardTypeID As Long) As Boolean
    Dim i As Integer
    If mlngCardTypeID = 0 Then
        CheckParaConfig = Not mbnQRPay: Exit Function
    End If
    If ZlGetParaConfig(lngCardTypeID, 6) = False Then
        MsgBox "病人预交管理中的押金部分不支持使用" & mstr结算方式 & "进行缴款，请使用其他结算方式缴款！" & vbCrLf & _
                     "如需用" & mstr结算方式 & "进行缴款，请联系接口管理员调整三方接口。", vbInformation, gstrSysName
        If mbytInState = EM_异常重收 And Not mblnCheckSwapFailed Then Exit Function
        With cboStyle
            For i = 0 To .ListCount - 1
                If .ItemData(i) = 1 Then .ListIndex = i
            Next
            If .ItemData(.ListIndex) = -1 Then txtMoney = ""
        End With
        Exit Function
    End If
    CheckParaConfig = True
End Function
    
Private Sub cboStyle_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then
        If cboStyle.ListIndex = -1 Then
            Beep
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        If cboStyle.Locked Then Exit Sub
        If KeyAscii >= 32 Then
            lngIdx = cbo.MatchIndex(cboStyle.hwnd, KeyAscii)
            If lngIdx = -1 And cboStyle.ListCount > 0 Then lngIdx = 0
            cboStyle.ListIndex = lngIdx
        End If
    End If
End Sub

Private Sub cboStyle_Validate(Cancel As Boolean)
    If cboStyle.Locked Then Exit Sub
    If Not (cboStyle.ListIndex > -1 And mbytInState = EM_收押金) Then Exit Sub
    If mbytInState = EM_收押金 Then
         If InStr(1, mstrPrivs, ";押金收款;") = 0 Then
             MsgBox "你没有权限进行押金收款操作！", vbInformation, gstrSysName
         End If
     Else
         If InStr(1, mstrPrivs, ";押金退款;") = 0 Then
             MsgBox "你没有权限进行押金退款操作！", vbInformation, gstrSysName
         End If
     End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If SendMessage(cboUnit.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cboUnit.hwnd, KeyAscii, 0.5)
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
        cmdOK.Enabled = True
        chkCancel.ForeColor = &HFF&
        btQRCodePay.Visible = False
        '清除相关界面和数据
        Set mpatiInfo = New clsPatientInfo '清除病人信息
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
        zlControl.ControlSetFocus cboNO
    Else
        '弹起
        chkCancel.ForeColor = 0
        btQRCodePay.Visible = btQRCodePay.Tag <> ""
        picFace.Enabled = True
        Call Load支付方式
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
    If mbytInState = EM_收押金 Then
        If chkCancel.Value = Checked Then
            If MsgBox("确实要放弃退款退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If mpatiInfo.病人ID > 0 Then
                If MsgBox("该病人的押金款尚未收取,确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
    End If
    Unload Me
End Sub
Private Sub zlBackDepositYJ()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：退押金操作
    '编制：刘兴洪
    '日期：2010-06-18 16:34:59
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim blnCanDel As Boolean, intInsure As Integer
    Dim bln打印 As Boolean, strSQL As String
    Dim msgBoxResult As String, strErrMsg As String
    Dim dbl押金金额 As Double, rsTmp As New ADODB.Recordset
    Dim blnCancel As Boolean
    
    mbytOracleBackType = 1
    '退款
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
    
    '结算方式检查
    If cboStyle.ListIndex = -1 Then
        MsgBox "请确定结算方式！", vbExclamation, gstrSysName
        zlControl.ControlSetFocus cboStyle: Exit Sub
    End If
    '检查单据是否已退，或为异常单据
    If Not CheckBackErrBill(cboNO.Text, strErrMsg) Then
        MsgBox strErrMsg, vbExclamation, gstrSysName
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
            msgBoxResult = MsgBox("是否需要打印押金红票？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
            bln打印 = (msgBoxResult = vbYes)
        End Select
    
        If mYJinfo.lng卡类别ID = 0 Then
            If gbyt预存款消费验卡 <> 0 Then
                If mbln住院退预交验证 Or cboType.ItemData(cboType.ListIndex) = 1 Then
                    If CreatePublicExpense() Then
                        If Not gobjPublicExpense.zlPatiIdentify(mlngModul, Me, Val(txtPatient.Tag), Val(StrToNum(txtMoney.Text)), False) Then Exit Sub
                    End If
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
        strSQL = "Select Nvl(金额, 0) as 押金金额 From 病人押金记录 Where NO = [1] And 记录状态=1 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取押金余额", cboNO.Text)
        If rsTmp.EOF Then
            MsgBox "没有发现要退款的押金记录,该单据可能已经被退！", vbInformation, gstrSysName
            Exit Sub
        End If
        dbl押金金额 = Val(rsTmp!押金金额)
        If CCur(StrToNum(txtMoney.Text)) > dbl押金金额 Then
            If mbytBackMoneyType = 1 Then
                Call MsgBox("该笔押金的退款金额比押金金额多，你不能作废这张单据！", vbInformation + vbOKOnly, gstrSysName)
                Exit Sub
            Else
                If MsgBox("该笔押金的退款金额比押金金额多，忽略吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                mbytOracleBackType = 0
            End If
        End If
        
        cmdOK.Enabled = False   '防医保延时
        
        '检查三方接口交易是否合法
        '108666:李南春，2017/5/9，恢复确认按钮可用状态
        If zlCheckDepositDelValied(Val(cboNO.Tag), StrToNum(txtMoney.Text)) = False Then cmdOK.Enabled = True: Exit Sub
        
        '执行作废操作
        If Not CancelBill(CLng(cboNO.Tag), cboNO.Text, blnCanDel, intInsure, bln打印, cboNote.Text) Then '退款
'            MsgBox "操作失败,请重试该操作。如仍有问题,请与系统管理员联系！", vbExclamation, gstrSysName
            cmdOK.Enabled = True
            Exit Sub
        End If
        
        If bln打印 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_3", Me, "NO=" & cboNO.Text, 2)
            Call zlCheckFactIsEnough
        End If
        
        cmdOK.Enabled = True
        
        '医保改动
        For i = 0 To cboStyle.ListCount - 1
            If cboStyle.ItemData(i) = 3 Then
                cboStyle.RemoveItem i: Exit For
            End If
        Next
     Else
        blnCancel = True
    End If
    If mbytInState <> EM_退押金 Then
        chkCancel.Value = Unchecked '(并激活事件)
    Else
        mblnOK = Not blnCancel
        Unload Me: Exit Sub '退款模式操作后退出
    End If
    mblnOK = Not blnCancel
    Call ClearBill
End Sub

Private Function CheckBackErrBill(ByVal strNO As String, ByRef strErrMsg As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:退款时检查单据是否能退
    '入参:
    '编制:
    '日期:2018-07-20
    '说明:
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select 校对标志 From 病人押金记录 Where  记录状态=3 And No=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If Not rsTmp.EOF Then
        strErrMsg = "单据[" & strNO & "]已退款，请勿重复操作！"
        If Nvl(rsTmp!校对标志, 0) <> 0 Then
            strErrMsg = "单据[" & strNO & "]已生成为异常退款单据，请退出重新读取！"
        End If
        Exit Function
    End If
    CheckBackErrBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckDataValied(Optional ByVal bln打印 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查数据是否合法
    '返回：合法返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-18 16:38:39
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
   '新单存盘
  If mpatiInfo.病人ID = 0 Then
        MsgBox "没有确定收取押金的病人,不能进行押金充值！", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtPatient: Exit Function
  End If
          
  If LenB(StrConv(txtUnit.Text, vbFromUnicode)) > 50 Then
      MsgBox "缴款单位名称只允许 50 个字符或 25 个汉字,请修改！", vbInformation, App.Title
      zlControl.ControlSetFocus txtUnit: Exit Function
  End If
  If LenB(StrConv(txt开户行.Text, vbFromUnicode)) > 50 Then
      MsgBox "开户行名称只允许 50 个字符或 25 个汉字,请修改！", vbInformation, App.Title
      zlControl.ControlSetFocus txt开户行: Exit Function
  End If
  If LenB(StrConv(cboNote.Text, vbFromUnicode)) > 50 Then
      MsgBox "缴款摘要只允许 50 个字符或 25 个汉字,请修改！", vbInformation, App.Title
      zlControl.ControlSetFocus cboNote: Exit Function
  End If
  If CheckParaConfig(mlngCardTypeID) = False Then Exit Function
  If mbytInState = EM_收押金 Then
    If cboType.ListIndex < 0 Then Exit Function
    '问题:44963
    If mpatiInfo Is Nothing Then Exit Function
    If mpatiInfo.病人ID = 0 Then Exit Function
    If cboType.ItemData(cboType.ListIndex) = 2 Then
        If Not mpatiInfo.在院 And gblnAllowOut = False Then
            If Not (mpatiInfo.病人性质 = 0 And mpatiInfo.主页ID = 0 And mpatiInfo.住院状态 = 0) Then
                MsgBox "病人还未住院,不能缴住院押金,请检查!", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Else
        If mpatiInfo.在院 And gblnBanIn = True Then
            MsgBox "病人还未出院,不能缴门诊押金,请检查!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
  End If
  '金额检查
  '问题27363 by lesfeng 2010-01-13
    If txtMoney.Text = "" Then
      MsgBox "收款金额不能为空,请输入！", vbExclamation, gstrSysName
      zlControl.ControlSetFocus txtMoney: Exit Function
    ElseIf CCur(StrToNum(txtMoney.Text)) = 0 Then
      MsgBox "收款金额不能为零,请输入！", vbExclamation, gstrSysName
      zlControl.ControlSetFocus txtMoney: Exit Function
    End If

    mbytOracleBackType = 1

    If cbo押金类别.ListIndex = -1 Then
        MsgBox "请确定押金类别！", vbExclamation, gstrSysName
        zlControl.ControlSetFocus cbo押金类别: Exit Function
    End If
    
    If cboStyle.ListIndex = -1 Then
        MsgBox "请确定结算方式！", vbExclamation, gstrSysName
        zlControl.ControlSetFocus cboStyle: Exit Function
    End If
    
    If InStr(mstrPrivs, ";押金收款;") = 0 Then
        MsgBox "你没有权限进行押金收款操作！", vbInformation, gstrSysName
        Exit Function
    End If
  
    If bln打印 = False Then CheckDataValied = True: Exit Function
    If CheckInvoicePrint = False Then Exit Function

    CheckDataValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInvoicePrint() As Boolean
    '功能:检查票据相关
    
    On Error GoTo errHandle
    If mFactProperty.intInvoicePrint = 0 Then CheckInvoicePrint = True: Exit Function
    If gblnBill预交 Then
        If Trim(txtFact.Text) = "" Then
            MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
            zlControl.ControlSetFocus txtFact: Exit Function
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
                    zlControl.ControlSetFocus txtFact
            End Select
            txtFact.Text = ""
            Exit Function
        End If
    Else
        If Len(txtFact.Text) <> gbyt预交 And txtFact.Text <> "" Then
            MsgBox "票据号码长度应该为 " & gbyt预交 & " 位！", vbInformation, gstrSysName
            zlControl.ControlSetFocus txtFact: Exit Function
        End If
    End If
    CheckInvoicePrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckSwapIsSucces(ByVal lng押金ID As Long, ByVal dblMoney As Double, ByVal intSwapType As Integer, _
                                    ByRef strErrMsg As String, ByRef intState As Integer) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:检查交易是否成功
    '入参:intState(0-交易失败，1-交易正在进行)
    '编制:
    '日期:2018-06-28 15:06:20
    '说明:
    '-------------------------------------------------------------------------------------------------------------------------
    Dim intSwapStatus_Out As Integer, strSwapExtendInfor As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断交易是否成功（10.35.90）
    '入参:  frmMain-调用的主窗体
    '       lngModule-模块号
    '       intSwapType-0-扣款;1-退款；2-转帐
    '       lngCardTypeID-卡类别ID
    '       strCardNO-卡号
    '       dblSwapMoney-交易金额
    '       strBalanceIDs-本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn 收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡
    '       strExpend-扩展参数:退款或作废时，才传入,格式如下 ：
    '        <INPUT>
    '            <TKLIST>
    '                    <TK>
    '                       <JYLSH>交易流水号</JYLSH>
    '                       <KH>卡号</KH>
    '                       <JE>金额</JE>
    '                    </TK>
    '            </TKLIST>
    '        </INPUT>
    '出参:intSwapStatus_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '     strErrMsg- 返回的错误信息:  为空，将不提示,不为空时，界面提示该信息
    '     strXMLExpend-待以后扩展
    '返回：接口调用成功返回true,否则返回Flase
    '日期:2013-06-15 20:22:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If intSwapType = 1 Then strSwapExtendInfor = GetExpendInfo(lng押金ID, False, dblMoney)
    CheckSwapIsSucces = gobjSquare.objSquareCard.zlSwapIsSucces(Me, mlngModul, intSwapType, mlngCardTypeID, "8|" & lng押金ID, mstrBrushCardNo, _
        dblMoney, intSwapStatus_Out, strErrMsg, strSwapExtendInfor)
        
    intState = intSwapStatus_Out
End Function

Private Function GetExpendInfo(ByVal lng押金ID As Long, Optional ByVal blnReturn As Boolean, Optional ByVal dblMoney As Double) As String
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:组成三方接口zlSwapIsSucces扩展入参
    '入参:  lng押金ID
    '       blnReturn-(true-退款交易的入参,false-交易状态检查入参)
    '       dblMoney-退款交易的退款金额
    '编制:
    '日期:2018-07-20
    '说明:
    '       strExpend-扩展参数:退款或作废时，才传入,格式如下 ：
    '        <INPUT>
    '            <TKLIST>
    '                    <TK>
    '                       <JYLSH>交易流水号</JYLSH>
    '                       <KH>卡号</KH>
    '                       <JE>金额</JE>
    '                    </TK>
    '            </TKLIST>
    '        </INPUT>
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strExpend As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    If lng押金ID = 0 Then Exit Function
    strSQL = "Select No,结算方式,交易说明,Decode(记录状态, 2, -1 * 金额, 金额) As 金额,卡号,交易流水号 From 病人押金记录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng押金ID)
    If rsTmp.EOF Then Exit Function
    If blnReturn Then
        strExpend = "<INPUT>" & vbCrLf & _
                    "   <TKLIST>" & vbCrLf & _
                    "      <TK>" & vbCrLf & _
                    "        <JYLSH>" & Nvl(rsTmp!交易流水号) & "</JYLSH>" & vbCrLf & _
                    "        <TKFS>" & Nvl(rsTmp!结算方式) & "</TKFS>" & vbCrLf & _
                    "        <JYSM>" & Nvl(rsTmp!交易说明) & "</JYSM>" & vbCrLf & _
                    "        <DJH>" & Nvl(rsTmp!NO) & "</DJH>" & vbCrLf & _
                    "        <TKJE>" & dblMoney & "</TKJE>" & vbCrLf & _
                    "      </TK>" & vbCrLf & _
                    "   </TKLIST>" & vbCrLf & _
                    "</INPUT>"
    Else
        strExpend = "<INPUT>" & vbCrLf & _
                    "   <TKLIST>" & vbCrLf & _
                    "     <TK>" & vbCrLf & _
                    "       <JYLSH>" & Nvl(rsTmp!交易流水号) & "</JYLSH>" & vbCrLf & _
                    "       <KH>" & Nvl(rsTmp!卡号) & "</KH>" & vbCrLf & _
                    "       <JE>" & dblMoney & "</JE>" & vbCrLf & _
                    "     </TK>" & vbCrLf & _
                    "  </TKLIST>" & vbCrLf & _
                    "</INPUT>"
    End If

    GetExpendInfo = strExpend
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckBrushCard(Optional ByRef blnUnolad As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡
    '出参：是否关闭窗体
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lng病人id As String, strXmlIn As String
    Dim dblMoney As Double
    Dim strExpand As String '问题号:55666
    Dim dbl账户余额 As Double '问题号:55666
    Dim strBrushNo As String
    
    On Error GoTo errHandle
    dblMoney = 1 * StrToNum(txtMoney.Text)
    If cboStyle.ItemData(cboStyle.ListIndex) >= 0 Then CheckBrushCard = True: Exit Function
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
    If mpatiInfo.病人ID > 0 Then lng病人id = mpatiInfo.病人ID

    strBrushNo = mstrBrushCardNo
    
    strXmlIn = "" & _
    "<IN>" & vbCrLf & _
    "   <CZLX>0</CZLX>" & vbCrLf & _
    "   <QRCODE>" & mstrQRcode & "</QRCODE>" & vbCrLf & _
    "</IN>"
    
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, mlngCardTypeID, False, _
        Nvl(mpatiInfo.姓名), mpatiInfo.性别, mpatiInfo.年龄, dblMoney, mstrBrushCardNo, mstrbrPassWord, _
        False, True, False, False, Nothing, False, False, strXmlIn, _
        cboType.ItemData(cboType.ListIndex), lng病人id) = False Then Exit Function
        
    '保存前,一些数据检查
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    
    strXmlIn = "" & _
    "<IN>" & vbCrLf & _
    "   <QRCODE>" & mstrQRcode & "</QRCODE>" & vbCrLf & _
    "   <SFYJ>" & 1 & "</SFYJ>" & vbCrLf & _
    "</IN>"
    
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, mlngCardTypeID, _
        False, mstrBrushCardNo, dblMoney, "", strXmlIn) = False Then
        If gbln费用结算异步控制 Then
            '删除原始单据
            If mbytInState = EM_异常重收 Then
                strSQL = GetDeleteSQL(mstrInNO)
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                MsgBox "三方卡交易失败，已删除该异常单据。", vbInformation, gstrSysName
                blnUnolad = True
            End If
            Exit Function
        End If
    End If
    '问题号:55666,55851
    gobjSquare.objSquareCard.zlGetAccountMoney Me, mlngModul, mlngCardTypeID, mstrBrushCardNo, strExpand, dbl账户余额, False
    If dbl账户余额 <> 0 Then
        sta.Panels(2).Text = "账户余额:" & Format(dbl账户余额, "0.00")
        If dbl账户余额 < dblMoney Then
            MsgBox "注意:" & vbCrLf & _
                         "账户余额为" & Format(dbl账户余额, "0.00") & "元，小于原缴款金额" & Format(dblMoney, "0.00") & _
                         "，本次缴款" & Format(dbl账户余额, "0.00") & "元！", vbInformation, gstrSysName
            lblMoney.Tag = dblMoney
            dblMoney = Format(dbl账户余额, "0.00")
            lblRepairMoney.Visible = True
        End If
    End If
    
    '判断预交金是否超出刷卡的余额
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

Private Function ReCancelBill(ByVal strNO As String, Optional ByVal cllStatusUpdate As Collection) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:重退异常退款单据
    '入参:
    '编制:
    '日期:2018-07-03
    '说明:
    '-------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lng冲销ID As Long, rsBill As ADODB.Recordset
    Dim strSQL As String, blnTrans As Boolean
    Dim msgBoxResult As VbMsgBoxResult
    Dim bln打印 As Boolean
    
    On Error GoTo errHandle
    strSQL = "Select Id From 病人押金记录 Where 记录状态=2 And Nvl(校对标志,0)<>0  And No=[1]"
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsBill.RecordCount = 0 Then
        MsgBox "单据[" & strNO & "]为非异常单据，请刷新后重试！", vbInformation, gstrSysName
        Exit Function
    End If
    lng冲销ID = Nvl(rsBill!ID, 0)
    strSQL = "Select Id From 病人押金记录 Where 记录状态=3  And No=[1]"
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsBill.EOF Then
        MsgBox "没有查找到原始押金单据，无法进行退款操作！", vbInformation, gstrSysName
        Exit Function
    End If
    lngID = Nvl(rsBill!ID, 0)
    
    Select Case mFactProperty.intInvoicePrint
    Case 0 '不打印预交发票
       bln打印 = False
    Case 1 '自动打印
       bln打印 = True
    Case 2 '打印提醒
        msgBoxResult = MsgBox("是否需要打印押金红票？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        bln打印 = (msgBoxResult = vbYes)
    End Select
    If bln打印 Then Call GetFact  '重新获取发票号
    Set cllStatusUpdate = New Collection
    '更新校对标志，进行退款
    strSQL = "zl_病人押金记录_DELETE(" & lngID & ",'" & cboNote.Text & "','" & _
        UserInfo.编号 & "','" & UserInfo.姓名 & "'," & lng冲销ID & "," & _
        IIf(bln打印, "'" & txtFact.Text & "'", "NULL") & "," & IIf(bln打印, IIf(mlng领用ID > 0, mlng领用ID, "Null"), "Null") & ",2)"
    zlAddArray cllStatusUpdate, strSQL
    
    '调用三方接口
    If zlDepositDel(lngID, lng冲销ID, StrToNum(txtMoney.Text), strNO, cllStatusUpdate, blnTrans, , True) = False Then
        Exit Function
    End If
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    If bln打印 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_3", Me, "NO=" & mstrInNO, 2)
        Call zlCheckFactIsEnough
    End If
    ReCancelBill = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim msgBoxResult As VbMsgBoxResult '问题号:50656
    Dim bln打印 As Boolean  '问题号:57624
    Dim lng押金ID As Long, intState As Integer
    Dim strErrMsg As String, blnBeenErr As Boolean
    Dim cllStatusUpdate As Collection
    Dim blnVocherPrint As Boolean, bytP As Byte '打印凭条
    Dim strSavedDate As String '收款日期，用于打印
    Dim blnUnload As Boolean

    If chkCancel.Value = Checked Then
        If mbytInState = EM_退押金 Or mbytInState = EM_收押金 Then
            '退押金
            Call zlBackDepositYJ: Exit Sub
        ElseIf mbytInState = EM_异常重退 Then
            'EM_异常重退
            '重退异常退款单据

            If CheckSwapIsSucces(Val(cboNO.Tag), StrToNum(txtMoney.Text), 1, strErrMsg, intState) = False Then
                If intState = 0 Then
                    '退款交易失败，恢复为正常单据
                    If Not DelDepositErrBill(mstrInNO, 1) Then
                        MsgBox "三方卡退款异常单据删除失败，请稍后重试！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    MsgBox "单据[" & mstrInNO & "]三方卡退款交易已失败，请在收款列表进行重新退款" & IIf(strErrMsg <> "", "，" & vbCrLf & "错误信息如下：" & vbCrLf, "！") & _
                    strErrMsg, vbInformation, gstrSysName
                    mblnOK = True
                Else
                    MsgBox "三方交易正在进行中，无法进行退款，请稍候重试" & IIf(strErrMsg <> "", "，" & vbCrLf & "错误信息如下：" & vbCrLf, "！") & _
                        strErrMsg, vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                mblnOK = ReCancelBill(mstrInNO, cllStatusUpdate)
            End If
            '正常收费模式处理异常单据后，回到收费状态
            If mblnOptErrBill Then
                Call RestoreStatue: Exit Sub
            Else
                Unload Me: Exit Sub
            End If
        End If
    End If
    
    If mbytInState <> EM_异常作废 Then
        Select Case mFactProperty.intInvoicePrint
        Case 0 '不打印预交发票
           bln打印 = False
        Case 1 '自动打印
           bln打印 = True
        Case 2 '打印提醒
            msgBoxResult = MsgBox("是否需要打印押金票据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
            bln打印 = (msgBoxResult = vbYes)
        End Select
    End If
    
    If (mbytInState = EM_收押金 Or mbytInState = EM_异常重收) Then
        bytP = Val(zlDatabase.GetPara("押金凭条打印方式", glngSys, mlngModul))
        Select Case bytP
        Case 0 '不打印预交发票
           blnVocherPrint = False
        Case 1 '自动打印
           blnVocherPrint = True
        Case 2 '打印提醒
            msgBoxResult = MsgBox("是否需要打印押金凭条？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
            blnVocherPrint = (msgBoxResult = vbYes)
        End Select
    End If
    If mbytInState = EM_异常重收 Then
        '异常收款单据进行重收
        If bln打印 Then Call GetFact  '重新获取发票号
        If CheckDataValied(bln打印) = False Then Exit Sub
        If mblnCheckSwapFailed Then
            If CheckBrushCard(blnUnload) = False Then
                If Not blnUnload Then Exit Sub
                '删除异常单据后关闭窗体
                If mblnOptErrBill Then
                    Call RestoreStatue: Exit Sub
                Else
                    mblnOK = True: Unload Me: Exit Sub '刷新数据
                End If
            End If
            mblnOK = ReDepositErrBill(mstrInNO, bln打印, blnVocherPrint)
        Else
            If CheckSwapIsSucces(Val(cboNO.Tag), StrToNum(txtMoney.Text), 0, strErrMsg, intState) = False Then
                If intState = 0 Then
                    If CheckBrushCard(blnUnload) = False Then
                        If blnUnload Then '删除异常单据后关闭窗体
                            If mblnOptErrBill Then
                                Call RestoreStatue: Exit Sub
                            Else
                                mblnOK = True: Unload Me: Exit Sub '刷新数据
                            End If
                        Else
                            mblnCheckSwapFailed = True
                            btQRCodePay.Visible = btQRCodePay.Tag <> ""
                            Call SetCtrlEnabled: Exit Sub
                        End If
                    End If
                    mblnOK = ReDepositErrBill(mstrInNO, bln打印, blnVocherPrint)
                Else
                    MsgBox "三方交易正在进行中，无法进行重收，请稍候重试" & IIf(strErrMsg <> "", "，" & vbCrLf & "错误信息如下：" & vbCrLf, "！") & _
                            strErrMsg, vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                mblnOK = ReDepositErrBill(mstrInNO, bln打印, blnVocherPrint)
            End If
        End If
        '正常收费模式处理异常单据后，回到收费状态
        If mblnOptErrBill Then
            Call RestoreStatue: Exit Sub
        Else
            Unload Me: Exit Sub
        End If
    ElseIf mbytInState = EM_异常作废 Then
        '异常作废
        If CheckSwapIsSucces(Val(cboNO.Tag), StrToNum(txtMoney.Text), 0, strErrMsg, intState) = False Then
            If intState = 0 Then
                If DelDepositErrBill(mstrInNO) = False Then
                    MsgBox "单据作废失败，请稍候重试！", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                MsgBox "三方交易正在进行中，无法作废单据，请稍候重试" & IIf(strErrMsg <> "", "，" & vbCrLf & "错误信息如下：" & vbCrLf, "！") & _
                        strErrMsg, vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            MsgBox "三方交易已成功，不允许作废单据，请重新收费！", vbInformation, gstrSysName
            Exit Sub
        End If
        mblnOK = True: Unload Me: Exit Sub
    End If
    
    If Not Check未入科不交预交 Then Exit Sub
    If CheckDataValied(bln打印) = False Then Exit Sub
    If CheckBrushCard = False Then Exit Sub
    '存盘
    cmdOK.Enabled = False
    
    '中间不能有弹出类，避免长时间挂起造成并发
    If Not SaveBill(bln打印, lng押金ID, blnBeenErr, strSavedDate) Then
        If blnBeenErr Then
            Call SetcmdOkEnabled
            zlControl.ControlSetFocus txtPatient
        Else
            '三方卡时，根据接口信息提示
            If cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then
                MsgBox "押金单据保存失败,请重试该操作。如果仍有问题,请与系统管理员联系！", vbExclamation, gstrSysName
            End If
            cmdOK.Enabled = True: Exit Sub
        End If
    Else
        '问题号:57624
        '问题号:50656
        If bln打印 Then '票据号为空就表示不打印发票
            '78751:李南春,2014/10/20,增加预交票据打印格式
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_2", Me, "NO=" & cboNO.List(0), "病人ID=" & mpatiInfo.病人ID, "收款时间=" & Format(strSavedDate, "yyyy-mm-dd HH:MM:SS"), 2)
            Call zlCheckFactIsEnough
        End If
        If blnVocherPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_4", Me, "NO=" & cboNO.List(0), 2)
        End If
        '81693:李南春,2015/4/21,评价器
        If Not mobjPlugIn Is Nothing Then
            On Error Resume Next
            Call mobjPlugIn.PatiPrePayAfter(mpatiInfo.病人ID, cboType.ItemData(cboType.ListIndex), lng押金ID)
            Err.Clear
        End If
    End If
    '问题号:55666
    '存在补交金额的情况
    If UBound(Split(lblRepairMoney.Caption, ":")) = 1 And Split(lblRepairMoney.Caption, ":")(1) <> "" Then
        txtPatient.Tag = ""
        lblRepairMoney.Tag = Split(lblRepairMoney.Caption, ":")(1)
        IDKind.IDKind = IDKind.GetKindIndex("姓名")
        txtPatient.Text = "-" & mpatiInfo.病人ID
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
        txtMoney.Text = Format(lblRepairMoney.Tag, "0.00")
        lblRepairMoney.Tag = ""
        lblRepairMoney.Visible = False: lblRepairMoney.Caption = "补交额:"
        cmdOK.Enabled = True
        Exit Sub
    End If
    
    If mblnClearWinInfor Then
        Call ClearBill
        Call InitFace(True)
        Call cboStyle_Click
    Else
        '问题号:44732
        SetMoneyInfo False
        Set mpatiInfo = New clsPatientInfo
        
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact  '重新获取发票号
    End If
    Call SetcmdOkEnabled
    zlControl.ControlSetFocus txtPatient
    mblnOK = True
End Sub

Private Sub RestoreStatue()
    '功能：恢复收押金状态
    mbytInState = EM_收押金
    Call ClearBill: Call InitFace
    mblnOptErrBill = False
    cmdOK.Caption = "确定(&O)"
    chkCancel.Value = Unchecked
    Call SetCtrlEnabled
End Sub

Private Function DelDepositErrBill(ByVal strNO As String, Optional ByVal bytOpt As Byte) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:删除押金异常单据记录
    '入参: strno-单据号，Optype-(0-删除异常充值单据，1-删除异常退款单据)
    '编制:
    '日期:2018-06-29
    '说明:
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle

    strSQL = GetDeleteSQL(strNO, bytOpt)
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    DelDepositErrBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ReDepositErrBill(ByVal strNO As String, ByVal blnPrintInvoice As Boolean, ByVal blnPrintVocher As Boolean) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:异常单据重收
    '入参:  strNO-单据号
    '       blnPrintInvoice-是否打印票据 ，为true打印
    '       blnPrintVocher-是否打印凭条 ，为true打印
    '编制:
    '日期:2018-06-28 16:11:16
    '说明:
    '-------------------------------------------------------------------------------------------------------------------------
    Dim rsBill As ADODB.Recordset, strSQL As String, blnTrans As Boolean
    Dim dbl金额 As Double
    Dim strCurDate As String, cllStatusUpdate As Collection
    
    On Error GoTo errHandle
    
    strSQL = "Select Id,主页ID,金额,收款时间 From 病人押金记录 Where 记录状态=0 And Nvl(校对标志,0)<>0  And No=[1]"
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsBill.RecordCount = 0 Then
        MsgBox "单据[" & strNO & "]为非异常单据，请刷新后重试！", vbInformation, gstrSysName
        Exit Function
    End If
    strSQL = ""
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    Set cllStatusUpdate = New Collection
    dbl金额 = IIf(mblnCheckSwapFailed, StrToNum(txtMoney.Text), Nvl(rsBill!金额))
    Call zlGetDepositYJSQL(cllStatusUpdate, rsBill!ID, Nvl(rsBill!主页ID, 0), strNO, dbl金额, blnPrintInvoice, strCurDate, 2)

    '调用三方支付
    If zlInterfacePrayMoney(rsBill!ID, strNO, StrToNum(txtMoney.Text), cllStatusUpdate, blnTrans) = False Then
        Exit Function
    End If
    If blnTrans Then
        gcnOracle.CommitTrans: blnTrans = False
    Else
        If cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then
            zlDatabase.ExecuteProcedure cllStatusUpdate(1), Me.Caption
        End If
    End If
    If blnPrintInvoice Then '票据号为空就表示不打印发票
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_2", Me, "NO=" & cboNO.Text, "病人ID=" & mpatiInfo.病人ID, _
                                 "收款时间=" & Format(strCurDate, "yyyy-mm-dd HH:MM:SS"), 2)
        Call zlCheckFactIsEnough
    End If
    If blnPrintVocher Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_4", Me, "NO=" & cboNO.Text, 2)
    End If
    ReDepositErrBill = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ClearBill()
'功能:清除相关界面和数据
    If mbytInState = EM_收押金 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
    End If
    
    Set mpatiInfo = New clsPatientInfo '清除病人信息
    
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
    If Val(cbo押金类别.ItemData(cbo押金类别.ListIndex)) > 0 Then
            txtMoney.Text = Format(Val(cbo押金类别.ItemData(cbo押金类别.ListIndex)), "###0.00;-###0.00;;")
    End If
    If cboStyle.ListCount <> 0 And cboStyle.Tag <> "" Then cboStyle.ListIndex = Val(cboStyle.Tag) '恢复缺省结算方式
    txtCode.Text = "": txtCode.Locked = False
    
    txtMan.Text = UserInfo.姓名
    txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    cboNote.Text = ""
    
    '新的一张押金单据
    cboNO.Text = "": cboNO.Locked = True
    
    txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "修改票据号") And gblnBill预交 '89302
    If mFactProperty.intInvoicePrint <> 0 Then Call GetFact
    zlControl.ControlSetFocus txtPatient
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Activate()
    If mblnUnLoad Then Unload Me: Exit Sub
    If mbytInState = EM_收押金 Then
        If gblnLED And Trim(txtPatient.Text) = "" Then
            zl9LedVoice.DisplayPatient ""    '双屏显示窗体必须在当前窗口显示之后调用显示才能移动窗体
        End If
    ElseIf mbytInState = EM_浏览单据 Then
        zlControl.ControlSetFocus cmdCancel
    ElseIf mbytInState = EM_退押金 Or mbytInState = EM_异常重退 Then
        If mstrInNO = "" Then
            zlControl.ControlSetFocus cboNO
        Else
           zlControl.ControlSetFocus cmdOK
        End If
        If mbytInState = EM_异常重退 Then txtMoney.Text = Abs(txtMoney.Text): Call InitPatientInfo(mstrInNO)
    ElseIf mbytInState = EM_异常重收 Or mbytInState = EM_异常作废 Then
        '初始化病人信息
        Call InitPatientInfo(mstrInNO)
        txtMoney.Enabled = False
    End If
    '问题号:45666
    If mbytInState = EM_收押金 And cboType.Text = "住院押金" Then '交押金
        chk仅显示本次押金.Visible = True
        chk仅显示本次押金.Value = IIf(zlDatabase.GetPara("仅显示本次预交", glngSys, mlngModul, , Array(chk仅显示本次押金), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
    End If
    
End Sub

Private Sub InitPatientInfo(ByVal strNO As String)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:根据异常单据号初始化病人信息
    '日期:2018-06-28 17:50:31
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsInfo As ADODB.Recordset
    
    On Error GoTo errHandle

    strSQL = "Select 病人id, 卡类别id, 结算方式 From 病人押金记录 Where NO = [1]"
    Set rsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsInfo.RecordCount = 0 Then Exit Sub
    If mlngCardTypeID = 0 Then mlngCardTypeID = Nvl(rsInfo!卡类别ID, 0)
    mstr结算方式 = Nvl(rsInfo!结算方式)
    If GetPatiInfo(rsInfo!病人ID, -1, mpatiInfo) = False Then
        Set mpatiInfo = New clsPatientInfo
        Exit Sub
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
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
            zlControl.ControlSetFocus txtFact
        Case vbKeyF4
            If Shift = vbCtrlMask And IDKind.Enabled Then
                Dim intIndex As Integer
                intIndex = IDKind.GetKindIndex("IC卡号")
                If intIndex <= 0 Then Exit Sub
                 IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
            End If
        Case vbKeyF6
            If btQRCodePay.Visible = False Or btQRCodePay.Enabled = False Then Exit Sub
            Call btQRCodePay.zlReReadQRCode
        Case vbKeyF11
            zlControl.ControlSetFocus txtPatient
        Case vbKeyF12
            zlControl.ControlSetFocus cboNO
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
            txtFact.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("当前预交票据号", glngSys, mlngFactModule, "")))
        Else
            mstrRedFact = zlCommFun.IncStr(UCase(zlDatabase.GetPara("当前预交票据号", glngSys, mlngFactModule, "")))
        End If
    End If
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化模块参数
    '编制:刘兴洪
    '日期:2012-02-27 11:23:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstr缺省结算方式 = zlDatabase.GetPara("缺省预交结算方式", glngSys, mlngModul)
    mbytBackMoneyType = Val(zlDatabase.GetPara("退款禁止方式", glngSys, mlngModul))
    '结算方式:金额|结算方式:金额....
    mblnClearWinInfor = IIf(zlDatabase.GetPara("缴预交后不清除信息", glngSys, glngModul) <> "1", True, False)
    mbln未入科不交预交 = zlDatabase.GetPara("病人未入科不准收预交", glngSys, mlngModul, , , InStr(mstrPrivs, ";参数设置;") > 0) = "1"
    gblnSeekName = Nvl(zlDatabase.GetPara("姓名模糊查找", glngSys, mlngModul, 1)) = 1
    mbln住院退预交验证 = zlDatabase.GetPara("住院退预交验证", glngSys, mlngModul, "0") = "1"
    '刷卡要求输入密码
    mblnCheckPass = Mid(zlDatabase.GetPara(46, glngSys, , "0000000000"), 8, 1) = "1"
    mbln排除未缴及未审 = zlDatabase.GetPara("剩余款排除未缴及未审金额", glngSys, mlngModul, "0") = "1"
    
End Sub

Private Sub Form_Load()
    
    Call InitPara
    mblnOK = False: mblnUnLoad = False

    '票据领用检查及初始
    If mbytInState = EM_收押金 Or mbytInState = EM_退押金 Then
        mblnStartFactUseType = zlStartFactUseType(2)
        If mblnStartFactUseType = False Then
            If mFactProperty.intInvoicePrint <> 0 Then Call GetFact(True, mbytInState = EM_退押金)
        End If
    End If
    
    zlControl.PicShowFlat picInfo, -1
    zlControl.PicShowFlat picFace, -1

    If Not InitUnit Then Unload Me: Exit Sub
    
    Call InitIDKind
    
    mstrCardPrivs = GetPrivFunc(glngSys, 1151)
    Call InitFace
    If mblnUnLoad Then Exit Sub
    
    lblTitle.Caption = gstrUnitName & "押金单据"
    
    If (mbytInState = EM_收押金 Or mbytInState = EM_退押金) And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.编号 & "号为您服务", mlngModul, gcnOracle
        
        Call zlCheckFactIsEnough
    End If

    If mbytInState = EM_收押金 Then
        IDKind.IDKind = Val(zlDatabase.GetPara("上次输入方式", glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0))
    End If
    
    '81693:李南春,2015/4/21,评价器
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear: On Error GoTo 0
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInState = EM_收押金: mstrInNO = ""
    mblnViewCancel = False: mblnUnLoad = False
    mlng领用ID = 0: mblnNOMoved = False
    mblnOptErrBill = False
    mstr退款操作员 = ""
    
    If (mbytInState = EM_收押金 Or mbytInState = EM_退押金) And gblnLED Then
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
    Set mpatiInfo = Nothing
    Set mobjPayMode = Nothing
    mblnCheckSwapFailed = False

    If mbytInState = EM_收押金 Then
        zlDatabase.SetPara "上次输入方式", IDKind.IDKind, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    End If
    '问题号:45666
    If mbytInState = EM_收押金 And cboType.Text = "住院押金" Then
        zlDatabase.SetPara "仅显示本次预交", chk仅显示本次押金.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    End If
End Sub

Private Sub InitPrepayType(Optional bytPrepayType As Byte = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化押金类型
    '编制:刘兴洪
    '日期:2011-07-14 18:50:56
    '---------------------------------------------------------------------------------------------------------------------------------------------

    With cboType
        .Clear
        .AddItem "门诊押金": .ItemData(.NewIndex) = 1
        If bytPrepayType = 1 Then .ListIndex = .NewIndex
        .AddItem "住院押金": .ItemData(.NewIndex) = 2
        If bytPrepayType = 2 Then .ListIndex = .NewIndex
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
     End With
End Sub

Private Sub InitFace(Optional blnSave As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据入口参数设置窗体界面及控制状态
    '编制:刘兴洪
    '日期:2011-07-17 10:36:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strValue As String, varData As Variant, varTemp As Variant
    Dim i As Integer, j As Integer, blnChange As Boolean
    
    If Not gobjSquare.objSquareCard Is Nothing And blnSave = False Then
        IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
    End If
    
    On Error GoTo errHandle
    
    strSQL = "Select 编码, 名称, 简码, 缺省标志  From 常用预交摘要"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    With cboNote
        .Clear
        If rsTmp.RecordCount > 0 Then
            While Not rsTmp.EOF
                .AddItem Nvl(rsTmp!名称)
                If Nvl(rsTmp!缺省标志) = 1 Then .ListIndex = .NewIndex
                rsTmp.MoveNext
            Wend
        End If
        .ListIndex = -1
    End With
    
    strSQL = "Select 编码, 名称, 简码, 缺省标志  From 押金类别 Order By 编码"
    strSQL = " Select Distinct 名称" & _
                  " From (Select b.编码, b.名称, b.简码, Decode(b.缺省标志, 1, 1, 0) As 缺省标志" & _
                  "           From 结算方式应用 A, 结算方式 B" & _
                  "           Where a.应用场合 = '预交款' And b.名称 = a.结算方式 And Nvl(b.性质, 1) = 5 " & _
                  "           Union All" & _
                  " Select 编码, 名称, 简码, Decode(缺省标志, 1, 2, 0) As 缺省标志 From 押金类别 Order By 缺省标志 Desc, 名称)"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cbo押金类别
        .Clear
        If rsTmp.RecordCount > 0 Then
            While Not rsTmp.EOF
                .AddItem Nvl(rsTmp!名称)
                rsTmp.MoveNext
            Wend
        Else
            MsgBox "未找到有效的押金类别，请在字典管理工具中【经济工作】分类下的【押金类别】中设置！", vbExclamation, gstrSysName
            mblnUnLoad = True
            Exit Sub
        End If
    End With
    
    '加载押金类别缺省金额
    strValue = zlDatabase.GetPara("代收款设置", glngSys, mlngModul)
    varData = Split(strValue, "|")
    
    With cbo押金类别
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i), ":")
            For j = 0 To cbo押金类别.ListCount - 1
                If varTemp(0) = cbo押金类别.List(j) Then
                    cbo押金类别.ItemData(j) = Val(varTemp(1)): Exit For
                End If
            Next
        Next
        mblnNotClick = True
        .ListIndex = 0
        mblnNotClick = False
    End With
    
    Call InitPrepayType
    If mblnUnLoad Then Exit Sub

    IDKind.Enabled = mbytInState = EM_收押金
    Select Case mbytInState
        Case EM_收押金 '收取押金
            '创建卡部件
            Call CreateMobjCard
            cboNO.Text = ""
            txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
            txtMan.Text = UserInfo.姓名
            
            Call Load支付方式
            '退款权限
            If InStr(mstrPrivs, ";押金退款;") = 0 Then
                chkCancel.Visible = False
            End If
            txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "修改票据号") And gblnBill预交 '89302
        Case EM_浏览单据 '指定单据浏览
            picList.Visible = False
            Me.Height = Me.Height - picList.Height
            If mblnViewCancel Then lblFlag.Visible = True
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
        Case EM_退押金, EM_异常重退 '指定单据退款
            chkCancel.Value = Checked   '在调用的click事件中处理 picFace.Enabled = True '！！！不允许部份退款！！！
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
                Else
                    If mbytInState = EM_退押金 Then
                        If zlThirdReturnCashCheck(blnChange) Then
                            If cboStyle.ListCount > 1 And blnChange Then
                                cboStyle.Enabled = True: cboStyle.Locked = False
                                cboStyle.BackColor = &H80000005
                            End If
                        End If
                    End If
                End If
            End If
            If mbytInState = EM_异常重退 Then cmdOK.Caption = "重退(&R)": cmdOK.Enabled = True
        Case EM_异常重收
            cmdOK.Caption = "重收(&R)"
            cmdOK.Enabled = True
            chkCancel.Visible = False
            Call Load支付方式
            
            If mstrInNO <> "" Then  '病人信息管理中退预交,没有指定单据号
                '显示单据内容
                Dim intBillErr As Integer
                intBillErr = ReadBill(mstrInNO)
                If intBillErr <> -1 Then
                    If intBillErr <> 3 Then
                        MsgBox "不能正确读取该单据内容，请与系统管理员联系！", vbExclamation, gstrSysName
                    End If
                    mblnUnLoad = True
                End If
            End If
        Case EM_异常作废
            If mblnViewCancel Then lblFlag.Visible = True
            chkCancel.Visible = False
            
            cmdOK.Caption = "作废(&Z)"
            
            picNO.Enabled = False
            picFace.Enabled = False
            txtPatient.Enabled = False
            cboStyle.Enabled = False
            cboPatiPage.Enabled = False
            cboType.Enabled = False
            cboUnit.Enabled = False
            '显示单据内容
            If Not ReadBill(mstrInNO) Then
                MsgBox "不能正确读取该单据内容，请与系统管理员联系！", vbExclamation, gstrSysName
                mblnUnLoad = True
            End If
    End Select

    If mbln排除未缴及未审 Then
        lbl剩余款额.ToolTipText = "剩余款 = 预交余额 + 医保预结算 - 未结费用 - 未缴费用 - 未审费用"
    Else
        lbl剩余款额.ToolTipText = "剩余款 = 预交余额 + 医保预结算 - 未结费用"
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub CreateMobjCard()
    '创建卡部件
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hwnd)
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hwnd)
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

Private Sub txtMoney_Change()
    '问题27363
    If IsNumeric(StrToNum(txtMoney.Text)) Then
        txtMoney.ForeColor = IIf(CCur(StrToNum(txtMoney.Text)) >= 0, vbBlue, vbRed)
    End If
End Sub

Private Sub txtMoney_GotFocus()
    txtMoney.SelStart = 0: txtMoney.SelLength = Len(txtMoney.Text)
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    '问题27363
    If KeyAscii <> 13 Then
        If KeyAscii = Asc(".") And InStr(txtMoney.Text, ".") > 0 Then KeyAscii = 0: Beep: Exit Sub
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
    Else
        If txtMoney.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtMoney_LostFocus()
    '问题27363
    Dim dblMoney  As Double
    If Not IsNumeric(StrToNum(txtMoney.Text)) Then zlControl.ControlSetFocus txtMoney: Exit Sub
    If mpatiInfo.病人ID > 0 And IsNumeric(StrToNum(txtMoney.Text)) Then
        txtMoney.Text = Format(StrToNum(txtMoney.Text), "##,##0.00;-##,##0.00; ;")
        If txtMoney.MaxLength > 12 Then txtMoney.MaxLength = 12
        '108813:李南春,2017/5/8,语音播报控制
        If gblnLED Then
            '#22 1234.56   --预收一千二百三十四点五六元 Y
            '#23 1234.56   --找零一千二百三十四点五六元 Z
            dblMoney = StrToNum(txtMoney.Text)
            zl9LedVoice.Speak "#22 " & dblMoney
        End If
    End If
End Sub

Private Sub cboNO_GotFocus()
    If Not cboNO.Locked Then cboNO.SelStart = 0: cboNO.SelLength = Len(cboNO.Text)
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim strOper As String, vDate As Date
    Dim blnChange As Boolean
    
    If cboNO.Locked Then Exit Sub
    
    '转换成大写(汉字不可处理)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    ElseIf cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 11)
        
        '是否已转入后备数据表中,记录性质为1表示交或冲预交
        If zlDatabase.NOMoved("病人押金记录", cboNO.Text, "", "", Me.Caption) Then
            If Not ReturnMovedExes(cboNO.Text, 6, Me.Caption) Then Exit Sub
            mblnNOMoved = False
        End If
        
        '单据权限
        If Not ReadBillInfo(0, cboNO.Text, -21, strOper, vDate) Then
            cboNO.Text = "": zlControl.ControlSetFocus cboNO: Exit Sub
        End If
        If Not BillOperCheck(6, strOper, vDate, "退款") Then
            cboNO.Text = "": zlControl.ControlSetFocus cboNO: Exit Sub
        End If
        '问题27363
        '读取要作废的预交款单据
        Select Case ReadBill(cboNO.Text)
            Case -1
                If InStr(mstrPrivs, ";押金退款;") = 0 Then
                    MsgBox "你没有权限进行押金退款操作！", vbInformation, gstrSysName
                    chkCancel.Value = 0
                Else
                    If Val(StrToNum(txtMoney.Text)) < 0 Then
                        MsgBox "该笔预交金额为负,表示退款,不能执行该操作！", vbExclamation, gstrSysName
                        chkCancel.Value = 0
                    Else
                        zlControl.ControlSetFocus cmdOK
                    End If
                End If
                If chkCancel.Value <> 0 Then
                    If zlThirdReturnCashCheck(blnChange) Then
                        If cboStyle.ListCount > 1 And blnChange Then
                            cboStyle.Enabled = True: cboStyle.Locked = False
                            cboStyle.BackColor = &H80000005
                        End If
                    End If
                End If
            Case 0
                MsgBox "读取该押金单据失败！", vbExclamation, gstrSysName
                cboNO.Text = "": zlControl.ControlSetFocus cboNO
            Case 1
                MsgBox "该押金单据不存在！", vbExclamation, gstrSysName
                cboNO.Text = "": zlControl.ControlSetFocus cboNO
            Case 2
                MsgBox "该押金单据已经退款！", vbExclamation, gstrSysName
                cboNO.Text = "": zlControl.ControlSetFocus cboNO
            Case 3
                cboNO.Text = "": zlControl.ControlSetFocus cboNO
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
    Dim blnCard As Boolean
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

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnCancel As Boolean, bytPrepayType As Byte
    Dim str担保人 As String, dbl担保额 As Double
    
    Call ClearBill
    '读取病人信息
    SetMoneyInfo True
    sta.Panels(2) = ""
    If objCard.名称 Like "IC卡*" And objCard.系统 Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If Not GetPatient(objCard, strInput, blnCancel, blnCard) Then
        '处理异常单据跳过
        If mblnOptErrBill = False Then
            If blnCancel Then '取消输入
                zlControl.ControlSetFocus txtPatient: Exit Sub
            End If
            sta.Panels(2) = "未找到该病人，请检查输入内容!"
            If blnCard = True Then
                txtPatient.PasswordChar = "": txtPatient.Text = ""
                '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
                txtPatient.IMEMode = 0
            Else
                txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
            End If
            Set mpatiInfo = New clsPatientInfo
            zlControl.ControlSetFocus txtPatient
        End If
        Exit Sub
    End If
    
    '加载病人的住院次数
    Call LoadPatiPage(mpatiInfo.病人ID)
    '设置病人费用信息
    Call SetMoneyInfo(False, mpatiInfo.病人ID)
    
    '缺省病人的押金类型
    bytPrepayType = IIf(mpatiInfo.在院, 2, 1)
    If bytPrepayType <> cboType.ItemData(cboType.ListIndex) Then
        Call InitPrepayType(bytPrepayType)
    End If
    
    If mpatiInfo.当前科室ID <> 0 Then
        lbl床号.Caption = lbl床号.Tag & IIf(mpatiInfo.床号 = "", "家庭", mpatiInfo.床号)
    End If
            
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(mpatiInfo.住院号 = "", "", "住院号:" & mpatiInfo.住院号 & "   ") & _
                           IIf(mpatiInfo.门诊号 = "", "", "门诊号:" & mpatiInfo.门诊号)
    lbl科室.Caption = lbl科室.Tag & GET部门名称(mpatiInfo.出院科室ID)
    '46764
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(mpatiInfo.当前科室ID = 0, mpatiInfo.出院科室ID, mpatiInfo.当前科室ID))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    
    lbl费别等级.Caption = lbl费别等级.Tag & mpatiInfo.费别
    Call Get担保信息(mpatiInfo.病人ID, mpatiInfo.主页ID, dbl担保额, str担保人)
    lbl担保人.Caption = lbl担保人.Tag & str担保人
    lbl担保金额.Caption = lbl担保金额.Tag & Format(dbl担保额, "##,##0.00;-##,##0.00; ;")
    '问题号:116059,焦博,2017/12/7,预交界面显示病人手机号，提取病人信息中的“手机号”
    lbl手机号.Caption = lbl手机号.Tag & mpatiInfo.手机号
    lbl身份证号.Caption = lbl身份证号.Tag & mpatiInfo.身份证号
    lblMemo.Caption = lblMemo.Tag & mpatiInfo.病人备注
    '72828,冉俊明,2014-5-9,增加工作单位信息的显示
    lblWorkUnit.Caption = lblWorkUnit.Tag & mpatiInfo.工作单位
    
    txtPatient.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    txtPatient.Text = mpatiInfo.姓名
    txtPatient.Tag = mpatiInfo.病人ID
    '-----------------------------------------------------------------------------------------
    lblSex.Caption = lblSex.Tag & mpatiInfo.性别
    mstrPatiSex = mpatiInfo.性别
    lblOld.Caption = lblOld.Tag & mpatiInfo.年龄
    mstrPatiOld = mpatiInfo.年龄
    lbl家庭地址.Caption = lbl家庭地址.Tag & mpatiInfo.家庭地址
    lbl医疗付款方式.Caption = lbl医疗付款方式.Tag & mpatiInfo.医疗付款方式
    Call Led欢迎信息
    Call SetcmdOkEnabled
    
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Led欢迎信息()
    Dim strInfo As String, lngPatient As Long
    'LED初始化
    If mbytInState = EM_收押金 And gblnLED Then
        If gblnLedWelcome Then
            zl9LedVoice.Reset com
            zl9LedVoice.Speak "#1"
            zl9LedVoice.Init UserInfo.编号 & "号为您服务", mlngModul, gcnOracle
        End If
        strInfo = Trim(txtPatient.Text)
        If mpatiInfo.当前病区ID > 0 Then strInfo = strInfo & " " & mpatiInfo.性别 & " " & mpatiInfo.年龄: lngPatient = mpatiInfo.病人ID
        zl9LedVoice.DisplayPatient strInfo, lngPatient
    End If
End Sub
 
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
                                           Optional ByRef blnCancel As Boolean, _
                                           Optional ByVal blnCard As Boolean, _
                                           Optional ByVal lng病人id As Long, _
                                           Optional ByVal lng主页ID As Long = -1) As Boolean
    '功能：读取病人信息
    '参数：strInput=[刷卡]|[A病人ID]|[B住院号]
    '          lng病人ID=病人id,根据住院次数过滤押金记录时传入
    '          lng主页ID=-1表示门诊病人或查找所有住院次数;lng主页ID=0表示预入院病人;lng主页ID>0表示住院病人
    '说明：
    '     1.适用于病人预交款
    '     2.自动识别病人在院状态,读出(病人ID,主页ID,姓名,性别,年龄,住院号,床号,在院标志)
    '返回:是否读取成功,成功时mPatiInfo中包含病人信息
    Dim lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim blnHavePassWord As Boolean, blnIsMobileNO As Boolean
    
    blnCancel = False: mstr退款操作员 = ""
    If lng病人id > 0 Then GoTo ReadPati
    
    blnIsMobileNO = IDKind.IsMobileNo(strInput)
    If (blnCard And objCard.名称 Like "姓名*") And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then   '刷卡或缺省的卡
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人id, strPassWord, strErrMsg) = False Then
            If Not blnIsMobileNO Then GoTo NotFoundPati
            If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人id, strPassWord) = False Then GoTo NotFoundPati
        Else
            blnHavePassWord = True
        End If
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then  '病人ID
        lng病人id = Mid(strInput, 2)
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then  '住院号(对住(过)院的病人)
        If Val(Mid(strInput, 2)) = 0 Then GoTo NotFoundPati
        If zlGetPatiIDByInNo(Mid(strInput, 2), lng病人id, lng主页ID) = False Then GoTo NotFoundPati
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号(仅对门诊病人)
        If Val(Mid(strInput, 2)) = 0 Then GoTo NotFoundPati
        If GetPatiID("门诊号", Mid(strInput, 2), lng病人id) = False Then GoTo NotFoundPati
    Else '当作姓名
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                '限制模糊查长度,如果按照姓查找会影响性能
                If (Not gblnSeekName) Or (gblnSeekName And Len(strInput) < 2) Then GoTo NotFoundPati
                If GetPatiIdFromPatiName(txtPatient, strInput, lng病人id, Me, , , , , blnCancel) = False Then GoTo NotFoundPati
            Case "医保号"
                strInput = UCase(strInput)
                If GetPatiID("医保号", strInput, lng病人id) = False Then GoTo NotFoundPati
            Case "门诊号"
                If Not IsNumeric(strInput) Then GoTo NotFoundPati
                If Val(strInput) = 0 Then GoTo NotFoundPati
                If GetPatiID("门诊号", Mid(strInput, 2), lng病人id) = False Then GoTo NotFoundPati
            Case "住院号"
                If Not IsNumeric(strInput) Then GoTo NotFoundPati
                If Val(strInput) = 0 Then GoTo NotFoundPati
                If zlGetPatiIDByInNo(Val(strInput), lng病人id, lng主页ID) = False Then GoTo NotFoundPati
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人id, strPassWord, strErrMsg, lng卡类别ID) = False Then GoTo NotFoundPati
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人id, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati
                End If
                blnHavePassWord = True
        End Select
    End If
    
ReadPati:
    If lng病人id <= 0 Then GoTo NotFoundPati
    If GetPatiInfo(lng病人id, lng主页ID, mpatiInfo) = False Then GoTo NotFoundPati
    If mpatiInfo.病人ID = 0 Then GoTo NotFoundPati
    
    On Error GoTo Errhand
    '处理异常单据
    If mbytInState = EM_收押金 Then
        If OptOthersErrBill(mpatiInfo.病人ID) Then
            Exit Function
        End If
    End If
    '需要处理其他
    If mblnCheckPass And (blnCard Or IDKind.GetCurCard.接口序号 <> 0) Then
        If Not blnHavePassWord Then
            strPassWord = mpatiInfo.卡验证码
        End If
        If strPassWord <> "" Then
            If CreatePublicExpense() Then
                If gobjPublicExpense.zlVerifyPassWord(Me, strPassWord, mpatiInfo.姓名, mpatiInfo.性别, mpatiInfo.年龄) = False Then GoTo NotFoundPati
            End If
        End If
    End If
    GetPatient = True
    Exit Function
Errhand:
     If ErrCenter() = 1 Then
        Resume
     End If
    Call SaveErrLog
NotFoundPati:
    Set mpatiInfo = New clsPatientInfo
End Function

Private Function SaveBill(Optional blnPrintInvoice As Boolean = False, Optional ByRef lng押金ID As Long, Optional ByRef blnBeenErr As Boolean, Optional ByRef strCurDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对当前输入的预交款单据存盘
    '参数: blnBeenErr-是否有异常产生，true-有异常，false-无
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 11:15:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, i As Integer
    Dim blnTrans As Boolean, dblMoney As Double
    Dim lng主页ID As Long
    Dim blnThirdCard As Boolean
    Dim cllDeposit As Collection, cllStatusUpdate As Collection
    
    '当前交易是否三方卡
    blnThirdCard = cboStyle.ItemData(cboStyle.ListIndex) = -1 And mlngCardTypeID <> 0
    
    
    strNO = zlDatabase.GetNextNo(11)
    lng押金ID = zlDatabase.GetNextId("病人预交记录")
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    '问题27363
    dblMoney = 1 * StrToNum(txtMoney.Text)
    
Once:
    
    lng主页ID = IIf(cboType.ItemData(cboType.ListIndex) = 2, mpatiInfo.主页ID, 0)
    If cboPatiPage.Visible And cboPatiPage.ListIndex > 0 And cboType.ItemData(cboType.ListIndex) = 2 Then
        lng主页ID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If

    '获取预交SQL
    Set cllDeposit = New Collection
    Call zlGetDepositYJSQL(cllDeposit, lng押金ID, lng主页ID, strNO, dblMoney, blnPrintInvoice, strCurDate, _
                        IIf(blnThirdCard And mbytInState = EM_收押金 And gbln费用结算异步控制, 1, 0))
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    zlExecuteProcedureArrAy cllDeposit, Me.Caption, True, True
    
    If gbln费用结算异步控制 And blnThirdCard Then
        Set cllStatusUpdate = New Collection
        Call zlGetDepositYJSQL(cllStatusUpdate, lng押金ID, lng主页ID, strNO, dblMoney, blnPrintInvoice, strCurDate, 2)
    End If

    If zlInterfacePrayMoney(lng押金ID, strNO, StrToNum(txtMoney.Text), cllStatusUpdate, blnTrans, blnBeenErr) = False Then
        If blnBeenErr Then mblnOK = True
        Exit Function
    End If
  
    If blnTrans Then
        gcnOracle.CommitTrans: blnTrans = False
    End If
    
    '加入单据历史记录(所有类型单据)
    For i = 0 To cboNO.ListCount - 1
        strNO = strNO & "," & cboNO.List(i)
    Next
    cboNO.Clear
    For i = 0 To UBound(Split(strNO, ","))
        cboNO.AddItem Split(strNO, ",")(i)
        If i = 9 Then Exit For '只显示10个
    Next
    
    If Not gblnBill预交 And blnPrintInvoice And Trim(txtFact.Text) <> "" Then
        '松散：保存当前号码
        zlDatabase.SetPara "当前预交票据号", Trim(txtFact.Text), glngSys, mlngFactModule
    End If
    SaveBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If Err.Description Like "*退款金额大于病人押金余额*" And mbytOracleBackType = 1 Then
        If MsgBox("退款金额比病人押金余额多,是否忽略？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
        mbytOracleBackType = 0
        GoTo Once
    End If
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetDeleteSQL(ByVal strNO As String, Optional ByVal bytOpt As Byte) As String
    Dim strSQL As String
    
    strSQL = "Zl_病人押金异常记录_Delete("
    '  单据号_In       病人押金记录.No%Type
    strSQL = strSQL & "'" & strNO & "'," & bytOpt & ")"
    GetDeleteSQL = strSQL
End Function

Private Function zlGetDepositYJSQL(cllDeposit As Collection, ByVal lng押金ID As Long, ByVal lng主页ID As Long, _
                            ByVal strNO As String, ByVal dblMoney As Double, ByVal blnPrintInvoice As Boolean, _
                            ByVal strCurDate As String, ByVal byt操作状态 As Byte) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : -获取押金充值SQL
    ' 参数 : 出参-cllDeposit，入参：strNo-单据号，dblMoney-金额，blnPrintInvoice-是否打印发票
    '                               strCurDate-收款日期，btyOptType-操作类型
    '                               byt校对标志-校对标志，byt操作状态-操作状态
    '                           --操作状态:0-正常结算，1-保存为异常单据，2-完成异常结算
    '---------------------------------------------------------------------------------------
    
    Dim strSQL As String
    
    'Zl_病人押金记录_Insert_S
    strSQL = "Zl_病人押金记录_Insert_S("
    '  Id_In         病人押金记录.ID%Type,
    strSQL = strSQL & "" & lng押金ID & ","
    '  单据号_In     病人押金记录.NO%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  票据号_In     票据使用明细.号码%Type,
    If blnPrintInvoice Then
        strSQL = strSQL & "'" & txtFact.Text & "',"
    Else
        strSQL = strSQL & "NULL,"
    End If
    '  押金类别_In     病人押金记录.押金类别%Type,
    strSQL = strSQL & "'" & cbo押金类别.Text & "',"
    '  病人id_In     病人押金记录.病人id%Type,
    strSQL = strSQL & "" & mpatiInfo.病人ID & ","
    '  主页id_In     病人押金记录.主页id%Type,
    strSQL = strSQL & "" & ZVal(lng主页ID) & ","
    '  姓名_In         病人押金记录.姓名%Type,
    strSQL = strSQL & "'" & mpatiInfo.姓名 & "',"
    '  性别_In         病人押金记录.性别%Type,
    strSQL = strSQL & "'" & mpatiInfo.性别 & "',"
    '  年龄_In         病人押金记录.年龄%Type,
    strSQL = strSQL & "'" & mpatiInfo.年龄 & "',"
    '  门诊号_In       病人押金记录.门诊号%Type,
    strSQL = strSQL & ZVal(mpatiInfo.门诊号) & ","
    '  住院号_In       病人押金记录.住院号%Type,
    strSQL = strSQL & ZVal(mpatiInfo.住院号) & ","
    '  付款方式名称_In 病人押金记录.付款方式名称%Type,
    strSQL = strSQL & "'" & mpatiInfo.医疗付款方式 & "',"
    '  科室id_In     病人押金记录.科室id%Type,
    strSQL = strSQL & "" & ZVal(cboUnit.ItemData(cboUnit.ListIndex)) & ","
    '  缴款单位_In   病人押金记录.缴款单位%Type,
    strSQL = strSQL & "'" & Trim(txtUnit.Text) & "',"
    '  单位开户行_In 病人押金记录.单位开户行%Type,
    strSQL = strSQL & "'" & Trim(txt开户行.Text) & "',"
    '  单位帐号_In   病人押金记录.单位帐号%Type,
    strSQL = strSQL & "'" & Trim(txt帐号.Text) & "',"
    '  摘要_In       病人押金记录.摘要%Type,
    strSQL = strSQL & "'" & Trim(cboNote.Text) & "',"
    '  金额_In       病人押金记录.金额%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  结算方式_In   病人押金记录.结算方式%Type,
    strSQL = strSQL & "'" & mstr结算方式 & "',"
    '  结算号码_In   病人押金记录.结算号码%Type,
    strSQL = strSQL & "'" & txtCode.Text & "',"
    '  是否门诊_In   病人押金记录.预交类别%Type := Null,
    strSQL = strSQL & "" & IIf(cboType.ItemData(cboType.ListIndex) = 1, 1, 0) & ","
    '  领用id_In     病人押金记录.领用id%Type,
    strSQL = strSQL & "" & ZVal(mlng领用ID) & ","
    '  操作员编号_In 病人押金记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人押金记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  卡号_In       病人押金记录.卡号%Type := Null,
    strSQL = strSQL & "" & IIf(mstrBrushCardNo = "", "NULL", "'" & mstrBrushCardNo & "'") & ","
    '  收款时间_In   病人押金记录.收款时间%Type := Null
    strSQL = strSQL & "to_date('" & strCurDate & "','yyyy-mm-dd hh24:mi:ss'),"
    '  卡类别id_In   病人押金记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & ZVal(mlngCardTypeID) & ","
    '  交易流水号_In 病人押金记录.交易流水号%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  交易说明_In   病人押金记录.交易说明%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '   操作状态_In     Number :=0
    strSQL = strSQL & byt操作状态 & ")"
   
    zlAddArray cllDeposit, strSQL
    
    zlGetDepositYJSQL = True

End Function

Private Function ReadBill(strNO As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取预交款单据(浏览的、退款的),并填写界面及设置mpatiInfo(病人信息),将金额放在Tag中
    '入参:strNO-押金单据号
    '出参:
    '返回: -1-成功;0-失败;1-该单据不存在;2:该单据已经退款(浏览时无效);3-权限不足(已提醒)
    '编制:刘兴洪
    '日期:2011-07-15 11:45:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngPrepayType As Long, rsTemp As New ADODB.Recordset, strFullNO As String
    Dim strWhere As String, i As Long, strTmp As String
    Dim rs分院区显示 As New ADODB.Recordset
    Dim str担保人 As String, dbl担保额 As Double
    Dim lng卡类别ID, str卡类别名称 As String, bln退款验卡 As Boolean
    Dim str结算方式 As String, objCards As New Cards  'zlOneCardComLib.Cards
    On Error GoTo errH
    
    strFullNO = GetFullNO(strNO, 11)
    
    strWhere = IIf(mbytInState = EM_浏览单据 And mblnViewCancel, "And A.记录状态=2", " And A.记录状态 IN(0,1,3) ")
    gstrSQL = "" & _
    "Select a.Id, a.押金类别, a.实际票号, a.病人id, a.主页id, a.科室id As 当前科室ID, a.记录状态, a.摘要, a. 金额, a.结算方式, a.结算号码, a.收款时间, a.操作员姓名, a.缴款单位," & vbNewLine & _
    "       a.单位开户行, a.单位帐号, a.卡类别id, a.卡号, a.交易流水号, a.交易说明, b.性质 As 结算性质, a.是否门诊" & vbNewLine & _
    "From " & IIf(mblnNOMoved, "H", "") & "病人押金记录 A, 结算方式 B" & vbNewLine & _
    "Where a.No = [1]  And a.记录状态 In (0, 1, 3) And a.结算方式 = b.名称(+)" & strWhere
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFullNO)
    If rsTemp.RecordCount = 0 Then ReadBill = 1: Exit Function
    If Val(Nvl(rsTemp!卡类别ID)) > 0 Then
        If gOneCardData.zlGetYLCardObjs(objCards) = False Then Exit Function
        lng卡类别ID = Val(Nvl(rsTemp!卡类别ID))
        If objCards("K" & lng卡类别ID) Is Nothing Then
            MsgBox "未找到卡类别id为" & lng卡类别ID & "的医疗卡信息,请检查是否已启用!", vbOKOnly + vbInformation, gstrSysName
        Else
            str卡类别名称 = objCards("K" & lng卡类别ID).名称
            bln退款验卡 = objCards("K" & lng卡类别ID).是否退款验卡 = 1
        End If
    End If
    If GetPatiInfo(Val(Nvl(rsTemp!病人ID)), IIf(Val(Nvl(rsTemp!主页ID)) = 0, -1, Val(Nvl(rsTemp!主页ID))), mpatiInfo) = False Then Exit Function
    If mpatiInfo.病人ID = 0 Then
        MsgBox "未找到病人信息,请检查!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If

    If mbytInState = EM_退押金 Or chkCancel.Value = 1 Then
        '退款,需要检查是否存在具体的退款权限
        
        If InStr(1, mstrPrivs, ";押金退款;") = 0 Then
            MsgBox "你不具备对押金单据进行退款的权限,请与系统管理员联系!", vbOKOnly + vbInformation, gstrSysName
            ReadBill = 3
            Exit Function
        End If
        
        If gbln分院区显示 Then
            strTmp = "Select 1 From 部门表 A, 部门人员 B, 人员表 C" & vbNewLine & _
                    " Where a.Id = b.部门id And b.人员id = c.Id And (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null) And c.姓名 =[1]  And Rownum < 2"
    
            Set rs分院区显示 = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, Nvl(rsTemp!操作员姓名))
            
            If rs分院区显示.RecordCount = 0 Then
                MsgBox "该押金单据不属于本站点,不允许退款!", vbOKOnly + vbInformation, gstrSysName
                ReadBill = 3: Exit Function
            End If
        End If
    End If
    
    With mYJinfo
        .lng押金ID = Val(Nvl(rsTemp!ID))
        .strNO = strFullNO
        .lng卡类别ID = Val(Nvl(rsTemp!卡类别ID))
        .str卡号 = Nvl(rsTemp!卡号)
        .str名称 = str卡类别名称
        .str交易流水号 = Nvl(rsTemp!交易流水号)
        .dbl金额 = Val(Nvl(rsTemp!金额))
        .str交易说明 = Nvl(rsTemp!交易说明)
        .bln退款验卡 = bln退款验卡
        .dt收款时间 = Format(rsTemp!收款时间, "yyyy-MM-dd hh:mm:ss")
    End With
    
    cboNO.Text = strFullNO
    cboNO.Tag = rsTemp!ID '以此ID为准退款
    txtPatient.Text = mpatiInfo.姓名
    txtPatient.Tag = rsTemp!病人ID
    '74426:李南春,2014-7-9,病人姓名显示颜色处理
    Call SetPatiColor(txtPatient, Nvl(mpatiInfo.病人类型), IIf(Val(mpatiInfo.险类) = 0, &HFF0000, vbRed))
    lbl费别等级.Caption = lbl费别等级.Tag & mpatiInfo.费别
    
    Call Get担保信息(rsTemp!病人ID, Val(Nvl(rsTemp!主页ID)), dbl担保额, str担保人)
    lbl担保人.Caption = lbl担保人.Tag & str担保人
    lbl担保金额.Caption = lbl担保金额.Tag & dbl担保额
    lbl手机号.Caption = lbl手机号.Tag & mpatiInfo.手机号
    lbl身份证号.Caption = lbl身份证号.Tag & mpatiInfo.身份证号

    '72828,冉俊明,2014-5-9,增加工作单位信息的显示
    lblWorkUnit.Caption = lblWorkUnit.Tag & Nvl(mpatiInfo.工作单位)
    
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(IsNull(rsTemp!当前科室ID), 0, rsTemp!当前科室ID))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    cboType.ListIndex = -1
    lngPrepayType = Val(Nvl(rsTemp!是否门诊))
    For i = 0 To cboType.ListCount - 1
         If cboType.ItemData(i) = IIf(lngPrepayType = 1, 1, 2) Then
            cboType.ListIndex = i: Exit For
         End If
     Next
     
     With cboType
        If cboType.ListIndex < 0 Then
           .AddItem IIf(lngPrepayType = 1, "门诊押金", "住院押金")
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
    
    txtFact.Tag = txtFact.Text
    txtFact.Text = IIf(IsNull(rsTemp!实际票号), "", rsTemp!实际票号)
    If mbytInState = EM_异常重收 Then txtFact.Text = txtFact.Tag
    txtUnit.Text = IIf(IsNull(rsTemp!缴款单位), "", rsTemp!缴款单位)
    txt开户行.Text = IIf(IsNull(rsTemp!单位开户行), "", rsTemp!单位开户行)
    txt帐号.Text = IIf(IsNull(rsTemp!单位帐号), "", rsTemp!单位帐号)
    
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(Val(Nvl(mpatiInfo.住院号)) = 0, "", "住院号:" & mpatiInfo.住院号 & "   ") & _
                           IIf(Val(Nvl(mpatiInfo.门诊号)) = 0, "", "门诊号:" & mpatiInfo.门诊号)
    lblSex.Caption = lblSex.Tag & mpatiInfo.性别
    mstrPatiSex = mpatiInfo.性别
    lblOld.Caption = lblOld.Tag & mpatiInfo.年龄
    mstrPatiOld = mpatiInfo.年龄
    lbl床号.Caption = lbl床号.Tag & mpatiInfo.床号
    lbl科室.Caption = lbl科室.Tag & GET部门名称(Val(Nvl(rsTemp!当前科室ID)))
    lbl家庭地址.Caption = lbl家庭地址.Tag & Nvl(mpatiInfo.家庭地址)
    lbl医疗付款方式.Caption = lbl医疗付款方式.Tag & Nvl(mpatiInfo.医疗付款方式)
    txtMoney.Text = Format(rsTemp!金额, "##,##0.00;-##,##0.00;;")
    txtMoney.Tag = rsTemp!金额
    If mYJinfo.lng卡类别ID <> 0 Then
        cboStyle.ListIndex = cbo.FindIndex(cboStyle, mYJinfo.str名称, True)
    Else
        cboStyle.ListIndex = cbo.FindIndex(cboStyle, IIf(IsNull(rsTemp!结算方式), "", rsTemp!结算方式), True)
     End If
    If cboStyle.ListIndex = -1 Then
        str结算方式 = IIf(IsNull(rsTemp!结算方式), "", rsTemp!结算方式)
        If mYJinfo.lng卡类别ID <> 0 Then
            cboStyle.AddItem mYJinfo.str名称
            cboStyle.ItemData(cboStyle.NewIndex) = -1
            Call MakeCardsFrom结算方式(mYJinfo.str名称, str结算方式, Val(Nvl(rsTemp!卡类别ID)))
        Else
            cboStyle.AddItem str结算方式
            cboStyle.ItemData(cboStyle.NewIndex) = Val("" & rsTemp!结算性质)
            Call MakeCardsFrom结算方式(str结算方式, str结算方式)
        End If
        cboStyle.ListIndex = cboStyle.NewIndex
    End If
    
    txtCode.Text = IIf(IsNull(rsTemp!结算号码), "", rsTemp!结算号码)
    txtMan.Text = IIf(IsNull(rsTemp!操作员姓名), "", rsTemp!操作员姓名)
    txtDate.Text = Format(rsTemp!收款时间, "yyyy-MM-dd")
    cboNote.Text = IIf(IsNull(rsTemp!摘要), "", rsTemp!摘要)
    mblnNotClick = True
    If Nvl(rsTemp!押金类别) <> "" Then cbo押金类别.ListIndex = cbo.FindIndex(cbo押金类别, Nvl(rsTemp!押金类别), True)

    mblnNotClick = False
    '获取病人费用信息
    Call SetMoneyInfo(False, rsTemp!病人ID, strNO)
    ReadBill = -1
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetDepositData(ByVal lng病人id As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新读取预交数据
    '入参:lng病人ID-病人ID巧
    '编制:刘兴洪
    '日期:2011-07-22 17:02:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If lng病人id = 0 Then
        If mpatiInfo.病人ID = 0 Then Set mrsDepositBalance = Nothing: Exit Sub
        lng病人id = mpatiInfo.病人ID
    End If

    mdbl费用余额 = 0: mdbl预交余额 = 0: mdbl剩余款额 = 0
    '按类别先缓存,以提搞性能
    Set mrsDepositBalance = GetMoneyInfo(lng病人id)

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ShowPremayBalance(ByVal blnreReadData As Boolean, ByVal lng病人id As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据相关的结算方式和门诊类型,显示病人余额和费用信息
    '入参:blnReRead-重读数据
    '       lng病人ID-读取指定的病人ID(0时,从mPatiInfo记录中读取病人ID)
    '编制:刘兴洪
    '日期:2011-07-21 15:44:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim int押金类型 As Integer, bln三方接口 As Boolean
    Dim strWhere As String
    Dim dbl未审 As Double, dbl未缴 As Double, dblYB As Double
    Dim lng主页ID As Long, dbl剩余款额 As Double
    Dim rsYJMoney As New ADODB.Recordset, dbl押金余额 As Double
    
    On Error GoTo errHandle
    If lng病人id = 0 Then
        If mpatiInfo.病人ID = 0 Then Exit Sub
        lng病人id = mpatiInfo.病人ID
    End If
    
    If blnreReadData Then Call GetDepositData(lng病人id)
    sta.Panels(2).Text = ""
    mdbl费用余额 = 0: mdbl预交余额 = 0: mdbl剩余款额 = 0
    int押金类型 = cboType.ItemData(cboType.ListIndex)
    bln三方接口 = cboStyle.ItemData(cboStyle.ListIndex) = -1
    strWhere = "And nvl(卡类别ID,0)<>0 or nvl(结算卡序号,0)<>0 "

    If Not mrsDepositBalance Is Nothing Then
    With mrsDepositBalance
        .Filter = "类型=" & int押金类型
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
    
    Call Load医保预结(lng病人id, lng主页ID, dblYB)
    strSQL = "Select Sum(Nvl(金额, 0)) as 押金余额 From 病人押金记录 Where nvl(校对标志,0) =0 And 病人id = [1] " & IIf(lng主页ID = 0, "", "And 主页id = [2]")
    Set rsYJMoney = zlDatabase.OpenSQLRecord(strSQL, "读取押金余额", lng病人id, lng主页ID)
    
    dbl押金余额 = Val(Nvl(rsYJMoney!押金余额, 0))
    mdbl剩余款额 = mdbl预交余额 - mdbl费用余额
    '问题27363
    lbl费用余额.Caption = lbl费用余额.Tag & Format(mdbl费用余额, "##,##0.00;-##,##0.00; ;")
    lbl押金余额.Caption = lbl押金余额.Tag & Format(dbl押金余额, "##,##0.00;-##,##0.00; ;")
    dbl未审 = GetUnAuditedFee(lng病人id, , int押金类型)
    dbl未缴 = GetUnAuditedFee(lng病人id, False, int押金类型)
    lbl未审费用.Caption = lbl未审费用.Tag & Format(dbl未审, "##,##0.00;-##,##0.00; ;")
    lbl未缴费用.Caption = lbl未缴费用.Tag & Format(dbl未缴, "##,##0.00;-##,##0.00; ;")
    dbl剩余款额 = IIf(mbln排除未缴及未审, mdbl剩余款额 - dbl未缴 - dbl未审 + dblYB, mdbl剩余款额 + dblYB)
    lbl剩余款额.Caption = lbl剩余款额.Tag & Format(dbl剩余款额, "##,##0.00;-##,##0.00; ;")
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load医保预结(ByVal lng病人id As Long, ByVal lng主页ID As Long, Optional ByRef dblYB As Double)
    '功能:获取病人的医保预结算
    Dim rsMoney As ADODB.Recordset, strSQL As String

    On Error GoTo errHandle
    
    If lng病人id = 0 Then Exit Sub
    Set rsMoney = New ADODB.Recordset
    If lng主页ID = 0 Then
        strSQL = "Select Sum(金额) As 医保预结 From 保险模拟结算 Where 病人ID = [1] And 主页ID Is Null"
        Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, "提取医保预结", lng病人id)
    Else
        strSQL = "Select Sum(金额) As 医保预结 From 保险模拟结算 Where 病人ID = [1] And 主页ID = [2]"
        Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, "提取医保预结", lng病人id, lng主页ID)
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
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetMoneyInfo(blnClear As Boolean, Optional lng病人id As Long, _
    Optional strBackNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示金额等信息
    '入参:blnClear-清除
    '     lng病人ID-指定病人ID
    '     strBackNO-指定退预交单号(退款时传入,主要是是定位到清单上面去)
    '编制:刘兴洪
    ' 修改:刘兴洪(退号时,增加定位功能),增加参数;strBackNo
    '日期:2011-07-21 15:40:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset
    Dim strSQL As String
    
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
        '72828,冉俊明,2014-5-9,增加工作单位信息的显示
        lblWorkUnit.Caption = lblWorkUnit.Tag
        
        lbl未审费用.Caption = lbl未审费用.Tag
        lbl未缴费用.Caption = lbl未缴费用.Tag
        lbl费用余额.Caption = lbl费用余额.Tag
        lbl押金余额.Caption = lbl押金余额.Tag
        lbl剩余款额.Caption = lbl剩余款额.Tag
        lbl医保预结.Caption = lbl医保预结.Tag
        lbl手机号.Caption = lbl手机号.Tag
        lbl身份证号.Caption = lbl身份证号.Tag
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
        '显示病人余额和费用信息
        Call ShowPremayBalance(True, lng病人id)
        '检查是否有应收款
        strSQL = "Select Zl_Patientdue([1]) 剩余应收 From dual"
        Set rsMoney = New ADODB.Recordset
        Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, "提取应收款", lng病人id)
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
    '功能:显示历史的缴押金数据
    '编制:刘兴洪
    '日期:2011-09-16 10:17:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, int类型 As Integer, lngRow As Long, strWhere As String
    Dim rsMoney As ADODB.Recordset
    Dim lng病人id As Long
    
    If mpatiInfo.病人ID = 0 Then
        lng病人id = 0
    Else
        lng病人id = mpatiInfo.病人ID
    End If
    
    If cboType.ListIndex < 0 Then
        int类型 = 1
    Else
        int类型 = IIf(cboType.ItemData(cboType.ListIndex) = 1, 1, 0)
    End If
    
    On Error GoTo errHandle
    '84217,李南春,2015/4/22,显示指定的住院期间缴纳的预交
    If cboType.Text = "住院押金" And chk仅显示本次押金.Value = 1 And cboPatiPage.ListIndex >= 0 Then
        strWhere = " And A.主页ID= " & cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    
    If gbln分院区显示 Then
        strWhere = strWhere & _
                " And Exists (Select 1 From 人员表 C, 部门人员 D, 部门表 E " & _
                " Where C.姓名 =A.操作员姓名 And C.Id = D.人员id And D.部门id = E.Id And (E.站点 = '" & gstrNodeNo & "' Or E.站点 Is Null))"
    End If
            
    '所有历史缴款明细清单
    strSQL = _
    " Select Ltrim(To_Char(A.收款时间,'YYYY-MM-DD')) as 日期,A.NO as 单据号,B.名称 as 科室, " & _
    " Ltrim(To_Char(A.金额,'9,999,999,990.00')) as 缴款金额,A.结算方式 as 结算,A.操作员姓名 as 收款人 " & _
    " From " & IIf(mblnNOMoved, "H", "") & "病人押金记录 A,部门表 B" & _
    " Where A.科室ID=B.ID(+) And  A.病人ID=[1]  And A.是否门诊=[2] " & _
    " And  Nvl(A.校对标志, 0) = 0   " & strWhere & _
    " Order by A.收款时间 Desc"
    
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人id, int类型)
    mshList.Rows = 2: mshList.Cols = 2: mshList.Clear
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
    zlControl.ControlSetFocus txtFact
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
    If mpatiInfo.病人ID > 0 Then
        mstr病人类型 = mpatiInfo.病人类型
    End If
    If mstr病人类型 = "" Then
        If mpatiInfo.病人ID > 0 Then
            If mpatiInfo.险类 > 0 Then
                txtPatient.ForeColor = vbRed
            Else
                txtPatient.ForeColor = &HFF0000
            End If
        Else
            txtPatient.ForeColor = &HFF0000
        End If
    Else
        If CreatePublicPatient() Then
            txtPatient.ForeColor = gobjPublicPatient.GetPatiColor(mstr病人类型, True)
        End If
    End If
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtUnit_GotFocus()
    zlControl.ControlSetFocus txtUnit
End Sub

Private Sub txt开户行_GotFocus()
    zlControl.ControlSetFocus txt开户行
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
    zlControl.ControlSetFocus txt帐号
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
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

Private Function CancelBill(ByVal lngID As Long, ByVal strNO As String, ByVal blnCanDel As Boolean, _
                                        ByVal intInsure As Integer, ByVal bln打印 As Boolean, ByVal strNote As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对指定ID的押金单据执行退款处理
    '入参:lngID=单据ID
    '        blnCanDel=是否支持退个人帐户
    '        intInsure=单据中所使用的个人帐户的保险类别,无为0
    '        strNo=单据号
    '        strNote=摘要
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-07-19 09:28:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, blnTrans As Boolean
    Dim lng冲销ID As Long
    Dim blnThirdCard As Boolean
    Dim cllStatusUpdate As Collection
    
    If mYJinfo.lng卡类别ID <> 0 Then
        lng冲销ID = zlDatabase.GetNextId("病人预交记录")
        blnThirdCard = cboStyle.ItemData(cboStyle.ListIndex) = -1
        If Not blnThirdCard And mstr退款操作员 <> "" Then
            strNote = mstr退款操作员 & "强制退现:" & Format(txtMoney.Text, "0.00") & "元"
        End If
    Else
    End If
    On Error GoTo errH
    'Id_In         病人押金记录.Id%Type,
    '摘要_In       病人押金记录.摘要%Type,
    '操作员编号_In 病人押金记录.操作员编号%Type,
    '操作员姓名_In 病人押金记录.操作员姓名%Type,
    '冲销id_In     病人押金记录.Id%Type := Null,
    '票据号_In     病人押金记录.实际票号%Type := Null,
    '领用id_In     票据领用记录.Id%Type := Null,
    '操作状态_In   Number := 0,
    '三方退现_In   Number := 0,
    '退现方式_In   病人押金记录.结算方式%Type := Null
    strSQL = "zl_病人押金记录_DELETE(" & lngID & ",'" & strNote & "','" & _
        UserInfo.编号 & "','" & UserInfo.姓名 & "'," & IIf(lng冲销ID = 0, "NULL", lng冲销ID) & "," & _
        "'" & IIf(bln打印, mstrRedFact, "") & "'," & IIf(bln打印, IIf(mlng领用ID > 0, mlng领用ID, "Null"), 0) & _
        IIf(blnThirdCard And gbln费用结算异步控制, ",1)", ",0," & IIf(cboStyle.ItemData(cboStyle.ListIndex) <> -1, 1, 0) & ",'" & cboStyle.Text & "')")
    gcnOracle.BeginTrans: blnTrans = True
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If blnThirdCard And gbln费用结算异步控制 Then
        '更新校对标志，进行退款
        strSQL = "zl_病人押金记录_DELETE(" & lngID & ",'" & strNote & "','" & _
            UserInfo.编号 & "','" & UserInfo.姓名 & "'," & IIf(lng冲销ID = 0, "NULL", lng冲销ID) & "," & _
            "'" & IIf(bln打印, mstrRedFact, "") & "'," & IIf(bln打印, IIf(mlng领用ID > 0, mlng领用ID, "Null"), 0) & ",2)"
        Set cllStatusUpdate = New Collection
        zlAddArray cllStatusUpdate, strSQL
    End If
    
    '处理医保接口
    If intInsure <> 0 And blnCanDel Then
        If Not gclsInsure.TransferDelSwap(lngID, intInsure) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    End If

    If blnThirdCard Then
        If zlDepositDel(lngID, lng冲销ID, StrToNum(txtMoney.Text), strNO, cllStatusUpdate, blnTrans) = False Then
    
            Exit Function
        End If
    End If
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    
    
    If Not gblnBill预交 And bln打印 And mstrRedFact <> "" Then
        '松散：保存当前号码
        zlDatabase.SetPara "当前预交票据号", mstrRedFact, glngSys, mlngFactModule
    End If
    CancelBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitIDKind()
    Dim strKind As String
    strKind = "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;留|留观号|0;就|就诊卡|0;手|手机号|0"
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKind, txtPatient)
    mtySquareCard.blnExistsObjects = Not gobjSquare.objSquareCard Is Nothing
End Sub

Private Function zlCheckDepositDelValied(ByRef lng押金ID As Long, _
    ByVal dbl退款金额 As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费交易接口
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExend As String
    Dim cllSquareBalance As Collection
    
    If mYJinfo.lng卡类别ID = 0 Or cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then zlCheckDepositDelValied = True: Exit Function
    If Not mtySquareCard.blnExistsObjects Or gobjSquare.objSquareCard Is Nothing Then
            MsgBox "注意:" & vbCrLf & _
                         "      当前的押金按" & mYJinfo.str名称 & " 结算的,但不存在操作的相关部件,不能退款,请与系统管理员联系!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
            Exit Function
    End If
    
    Set cllSquareBalance = New Collection
    'Array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文,剩余未退金额)
    cllSquareBalance.Add Array(mYJinfo.lng卡类别ID, 0, 0, mYJinfo.str卡号, "", "", False, dbl退款金额)
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
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, mYJinfo.lng卡类别ID, False, mYJinfo.str卡号, _
        "8|" & lng押金ID, dbl退款金额, mYJinfo.str交易流水号, mYJinfo.str交易说明, strXMLExend) = False Then
          zlCheckDepositDelValied = False
          Exit Function
     End If
     '100610:李南春,2016/10/13，预交退款是否验证刷卡
     If mYJinfo.bln退款验卡 Then
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
        
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, mYJinfo.lng卡类别ID, False, _
            Trim(txtPatient.Text), mstrPatiSex, mstrPatiOld, dbl退款金额, mstrBrushCardNo, mstrbrPassWord, _
            True, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
        If mYJinfo.str卡号 <> mstrBrushCardNo Then
            MsgBox "注意:" & vbCrLf & _
                         "      当前卡号[" & mstrBrushCardNo & "]与原交易卡号[" & mYJinfo.str卡号 & "]不一致，请使用原卡交易!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If
     
goEnd:
    zlCheckDepositDelValied = True
    Exit Function
End Function

Private Function zlDepositDel(ByRef lng押金ID As Long, ByRef lng冲销ID As Long, ByVal dblMoney As Double, ByVal strNO As String, ByVal cllStatusUpdate As Collection, _
                                ByRef blnTrans As Boolean, Optional blnBeenErr As Boolean, Optional ByVal blnReCancel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：退预交交易
    '入参： lng押金ID-病人押金记录.ID，blnReCancel-重退，blnTrans-当前事务状态，blnBeenErr-是否产生异常
    '返回：成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strXMLExpend As String, strErrMsg As String, intState As Integer
    Dim rsBalance_Out As ADODB.Recordset, rsExpend_Out As ADODB.Recordset, cllThird As Collection, cllThirdExpend As Collection
    
    Err = 0: On Error GoTo Errhand:
    If mYJinfo.lng卡类别ID = 0 Or cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then zlDepositDel = True: Exit Function
    
    If gbln费用结算异步控制 Then
        If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
        '执行状态更新过程(开启事务不提交)
        If Not blnTrans Then gcnOracle.BeginTrans: blnTrans = True
        zlExecuteProcedureArrAy cllStatusUpdate, Me.Caption, blnTrans, blnTrans
    End If
    
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
        '    blnResolveXMLToRecord-是否解析XML串给记录集(rsBalance_Out,rsExpend_Out）
        '出参: strSwapNo-交易流水号(退款交易流水号)
        '      strSwapMemo-交易说明(退款交易说明)
        '    intStatus_Out-交易状态:接口返回False时，此参数有效: 0-交易调用失败;1-交易正在处理中
        '    strErrMsg_Out-错误信息:为空时，不提示，非空时，提示
        '       strSwapExtendInfor-交易的扩展信息
        '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
        '返回:函数返回    True:调用成功,False:调用失败
     strSwapNO = mYJinfo.str交易流水号: strSwapMemo = mYJinfo.str交易说明
     '81489,冉俊明,2015-4-29,退费传入冲销ID
     strSwapExtendInfor = "8|" & lng冲销ID
     strXMLExpend = GetExpendInfo(lng押金ID, True, dblMoney)
     If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, mYJinfo.lng卡类别ID, False, mYJinfo.str卡号, _
        "8|" & lng押金ID, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor, strXMLExpend, True, rsBalance_Out, rsExpend_Out, intState, strErrMsg) = False Then
        
        '删除无效的预交数据
        'intState为0时，回退事务，删除原始记录，为1时，回退事务
        If gbln费用结算异步控制 Then
            If intState = 1 Then
                gcnOracle.RollbackTrans: blnTrans = False
                MsgBox "三方卡交易正在进行中，已生成异常退款单据[" & strNO & "]" & IIf(strErrMsg <> "", "，" & vbCrLf & "错误信息如下：" & vbCrLf, "！") & _
                    strErrMsg, vbInformation, gstrSysName
                 blnBeenErr = True: Exit Function
            Else
                '回退病人押金记录等
                gcnOracle.RollbackTrans: blnTrans = False
                '删除原始单据
                strSQL = GetDeleteSQL(strNO, 1)
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                
                If blnReCancel Then
                    MsgBox "单据[" & strNO & "]三方卡退款交易失败，请重新执行退款操作" & IIf(strErrMsg <> "", "，" & vbCrLf & "错误信息如下：" & vbCrLf, "！") & _
                            strErrMsg, vbInformation, gstrSysName
                Else
                    MsgBox "三方卡交易失败，请稍后重试" & IIf(strErrMsg <> "", "，" & vbCrLf & "错误信息如下：" & vbCrLf, "！") & _
                        strErrMsg, vbInformation, gstrSysName
                End If
                Exit Function
            End If
        Else
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
    End If
    
    If Not rsBalance_Out Is Nothing Then
         If rsBalance_Out.RecordCount = 0 Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "交易失败，接口调用失败！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        If rsBalance_Out.RecordCount > 1 Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "交易失败，不支持多种结算方式支付！", vbInformation, gstrSysName
            Exit Function
        End If
        If Val(rsBalance_Out!交易金额) <> dblMoney Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "交易失败，实际退款金额与本次退款金额不符！", vbInformation, gstrSysName
            Exit Function
        End If
        '获取更新Sql
        Set cllThird = New Collection
        If GetYJThirdUpdateSQL(lng冲销ID, "", Nvl(rsBalance_Out!结算方式), 0, "", "", "", "", Nvl(rsBalance_Out!是否普通结算, 0), cllThird, True) Then
    
            zlExecuteProcedureArrAy cllThird, Me.Caption, True, True
        End If
    End If
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    zlDepositDel = True

    If rsExpend_Out Is Nothing Then
        If Save三方交易(lng冲销ID, mYJinfo.lng卡类别ID, mYJinfo.str卡号, strSwapNO, strSwapMemo, _
            strSwapExtendInfor, blnTrans, True, "8|" & lng冲销ID) = False Then Exit Function
    Else
        '获取扩展信息Sql
        If zlGetThreeSwapExpendSQL(mYJinfo.lng卡类别ID, lng冲销ID, mYJinfo.str卡号, rsExpend_Out, cllThirdExpend) Then
            '单独事务，不影响其他数据保存
            On Error GoTo ErrExpend:
            zlExecuteProcedureArrAy cllThirdExpend, Me.Caption
        End If
    End If

    Exit Function
Errhand:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter = 1 Then
        Resume
    End If
    Exit Function
ErrExpend:
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub Load支付方式(Optional ByVal bln加载退现方式 As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:刘兴洪
    '日期:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objCard As Card
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim str性质 As String, strErrMsg As String
    Dim strRQCardTypeIDs As String, objPayMode As Cards
    
    If InStr(1, mstrPrivs, ";押金收款;") > 0 Then
        str性质 = "1,2,7,8"
    End If
    If bln加载退现方式 Then str性质 = "1,2"
    If str性质 = "" Then str性质 = "1,2,7,8"
    
    On Error GoTo errHandle
    Set rsTemp = Get结算方式("预交款", str性质)
    'zlGetCards：获取有效的卡对象
    '入参:bytType
    '                   0-所有医疗卡
    '                   1-启用的医疗卡
    '                   2-所有存在三方账户的三方卡
    '                   3-启用的三方账户的医疗卡
    Set objPayMode = gobjSquare.objSquareCard.zlGetCards(3)
    Set mobjPayMode = New Collection
    
    With cboStyle
        .Clear: strRQCardTypeIDs = ""
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 1 To objPayMode.Count
                If objPayMode(i).结算方式 = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            '104083:李南春，2016/12/21，个人账户放在最后动态加入
            '性质为8的根据启用医疗卡来处理
            If Not blnFind And InStr(",3,8,", "," & rsTemp!性质 & ",") = 0 Then
                .AddItem Nvl(rsTemp!名称)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!性质))
                Call MakeCardsFrom结算方式(Nvl(rsTemp!名称), Nvl(rsTemp!名称))
                If rsTemp!缺省 = 1 Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                If mstr缺省结算方式 = Nvl(rsTemp!名称) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 1 To objPayMode.Count
            rsTemp.Filter = "名称 ='" & objPayMode(i).结算方式 & "'"
            If Not rsTemp.EOF Then
                If str性质 <> 5 And Not objPayMode(i).消费卡 Then
                    .AddItem objPayMode(i).名称: .ItemData(.NewIndex) = -1
                    Call MakeCardsFrom结算方式(objPayMode(i).名称, objPayMode(i).结算方式, objPayMode(i).接口序号)
                    If mstr缺省结算方式 = objPayMode(i).名称 Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                    If objPayMode(i).是否支持扫码付 Then
                        strRQCardTypeIDs = strRQCardTypeIDs & "," & objPayMode(i).接口序号
                    End If
                End If
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    If bln加载退现方式 Then Exit Sub
    If cboStyle.ListCount = 0 Then
        MsgBox "预交场合没有可用的结算方式,请先到结算方式管理中设置。", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    If strRQCardTypeIDs <> "" Then strRQCardTypeIDs = Mid(strRQCardTypeIDs, 2)
    '初始化扫码控件
    If btQRCodePay.zlInit(Me, strRQCardTypeIDs, glngSys, mlngModul, gcnOracle, gstrDBUser, strErrMsg) = False Then strRQCardTypeIDs = ""
    btQRCodePay.Tag = strRQCardTypeIDs
    btQRCodePay.Visible = strRQCardTypeIDs <> "" And mbytInState = EM_收押金
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Save三方交易(ByVal lng押金ID As Long, ByVal lng卡类别ID As Long, _
    ByVal str卡号 As String, str交易流水号 As String, str交易说明 As String, strExpend As String, _
    blnTrans As Boolean, Optional bln退押金 As Boolean = False, Optional strExpendOld As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存三方结算数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, varData As Variant, varTemp As Variant, cllPro As Collection, i As Long
     
    Err = 0: On Error GoTo Errhand:
    If bln退押金 = False Then
        '退费时,不更改交易
        '更新交易信息
         '    Zl_三方接口更新_Update
        strSQL = "Zl_三方接口更新_Update("
        '  卡类别id_In   病人预交记录.卡类别id%Type,
        strSQL = strSQL & "" & lng卡类别ID & ","
        '  消费卡_In     Number,
        strSQL = strSQL & "0,"
        '  卡号_In       病人预交记录.卡号%Type,
        strSQL = strSQL & "'" & str卡号 & "',"
        '  结帐ids_In    Varchar2,
        strSQL = strSQL & "'" & lng押金ID & "',"
        '  交易流水号_In 病人预交记录.交易流水号%Type,
        strSQL = strSQL & "'" & str交易流水号 & "',"
        '  交易说明_In   病人预交记录.交易说明%Type
        strSQL = strSQL & "'" & str交易说明 & "',"
        '预交款缴款_In Number := 0
        strSQL = strSQL & "" & 2 & ","
        '退费标志 :1-退费;0-付费
        strSQL = strSQL & "" & IIf(bln退押金, 1, 0) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
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
                        strSQL = strSQL & "0,"
                        '卡号_In     病人预交记录.卡号%Type,
                        strSQL = strSQL & "'" & str卡号 & "',"
                        '结帐ids_In  Varchar2,
                        strSQL = strSQL & "'" & lng押金ID & "',"
                        '交易信息_In Varchar2:交易项目|交易内容||...
                        strSQL = strSQL & "'" & str交易信息 & "',"
                        '预交款缴款_In Number := 0
                        strSQL = strSQL & "2,"
                        '结算方式_In   病人预交记录.结算方式%Type := Null
                        strSQL = strSQL & "Null,"
                        '预交id_In     病人预交记录.Id%Type := Null
                        strSQL = strSQL & "Null,"
                        '性质_In       三方结算交易.性质%Type := Nul
                        strSQL = strSQL & "2)"
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
            strSQL = strSQL & "0,"
            '卡号_In     病人预交记录.卡号%Type,
            strSQL = strSQL & "'" & str卡号 & "',"
            '结帐ids_In  Varchar2,
            strSQL = strSQL & "'" & lng押金ID & "',"
            '交易信息_In Varchar2:交易项目|交易内容||...
            strSQL = strSQL & "'" & str交易信息 & "',"
            '预交款缴款_In Number := 0
            strSQL = strSQL & "2,"
            '结算方式_In   病人预交记录.结算方式%Type := Null
            strSQL = strSQL & "Null,"
            '预交id_In     病人预交记录.Id%Type := Null
            strSQL = strSQL & "Null,"
            '性质_In       三方结算交易.性质%Type := Nul
            strSQL = strSQL & "2)"
            zlAddArray cllPro, strSQL
        End If
    End If
    Err = 0: On Error GoTo ErrOthers: blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption
    blnTrans = False
    Save三方交易 = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrOthers:
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    '    能保存多少,存多少
     Call ErrCenter
End Function

Private Function zlInterfacePrayMoney(ByVal lng押金ID As Long, ByVal strNO As String, ByVal dblMoney As Double, _
                                       ByVal cllStatusUpdate As Collection, ByRef blnTrans As Boolean, Optional blnBeenErr As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接口支付金额
    '参数: intState-交易接口返回失败后的交易状态，0-失败，1-正在进行
    '返回:支付成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSQL As String
    Dim rsBalance_Out As ADODB.Recordset, rsExpend_Out As ADODB.Recordset, strErrMsg As String, intState As Integer
    Dim dbl交易金额 As Double, strTmp As String
    Dim cllThird As Collection, cllThirdExpend As Collection
    
    'intState 0-失败，1-正在进行
    If cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    If mlngCardTypeID = 0 Then zlInterfacePrayMoney = True: Exit Function
    
    If gbln费用结算异步控制 Then
        If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
        '执行状态更新过程(开启事务不提交)
        If Not blnTrans Then gcnOracle.BeginTrans: blnTrans = True
        zlExecuteProcedureArrAy cllStatusUpdate, Me.Caption, blnTrans, blnTrans
    End If
    
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
    strSwapExtendInfor = "" & _
                                "<IN>" & vbCrLf & _
                                "       <QRCODE>" & mstrQRcode & "</QRCODE>" & vbCrLf & _
                                "       <SFYJ>" & 1 & "</SFYJ>" & vbCrLf & _
                                "</IN>"
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, mlngCardTypeID, False, mstrBrushCardNo, "", strNO, _
        dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor, True, rsBalance_Out, rsExpend_Out, intState, strErrMsg) = False Then
         
        '删除无效的预交数据
        'intState为0时，回退事务，删除原始记录，为1时，回退事务
        If gbln费用结算异步控制 Then
            If intState = 1 Then
                gcnOracle.RollbackTrans: blnTrans = False
                MsgBox "三方卡交易正在进行中，已生成异常单据[" & strNO & "]" & IIf(strErrMsg <> "", "，" & vbCrLf & "错误信息如下：" & vbCrLf, "！") & _
                    strErrMsg, vbInformation, gstrSysName
                blnBeenErr = True
                Exit Function
            Else
                '回退人员余额，预交单据余额等
                gcnOracle.RollbackTrans: blnTrans = False
                '删除原始单据
                If mbytInState = EM_收押金 Then
                    strSQL = GetDeleteSQL(strNO)
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                End If
                If mbytInState = EM_异常重收 Then
                    strTmp = "三方卡交易失败，已删除该异常单据。"
                Else
                    strTmp = "三方卡交易失败，请稍后重试" & IIf(strErrMsg <> "", "，" & vbCrLf & "错误信息如下：" & vbCrLf, "！") & _
                                    strErrMsg
                End If
                MsgBox strTmp, vbInformation, gstrSysName
                Exit Function
            End If
        Else
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
    End If
    
    '根据实际返回交易信息更新
    If rsBalance_Out Is Nothing Then
        '原有接口方式
        If Save三方交易(lng押金ID, mlngCardTypeID, mstrBrushCardNo, strSwapGlideNO, strSwapMemo, _
            strSwapExtendInfor, blnTrans) = False Then Exit Function
        zlInterfacePrayMoney = True
        Exit Function
    Else
        If rsBalance_Out.RecordCount = 0 Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "交易失败，接口调用失败！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        If rsBalance_Out.RecordCount > 1 Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "交易失败，不支持多种结算方式支付！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        dbl交易金额 = Nvl(rsBalance_Out!交易金额, 0)
        If dbl交易金额 = 0 Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "交易失败，当前交易金额为0，请检查 ！", vbInformation, gstrSysName:
            Exit Function
        End If
        If dbl交易金额 <> 0 Then sta.Panels(2).Text = "本次交易金额:" & dbl交易金额
        
        If dbl交易金额 < dblMoney Then
            MsgBox "注意:" & vbCrLf & _
                                 "本次交易金额为" & Format(dbl交易金额, "0.00") & "元，跟原缴款金额" & Format(dblMoney, "0.00") & "不一致！", vbInformation, gstrSysName
            dblMoney = Format(dbl交易金额, "0.00")
            lblRepairMoney.Visible = True
        End If
        
        '判断预交金是否超出刷卡的余额
        If lblRepairMoney.Visible Then
            lblRepairMoney.Caption = "补交额:" & Format((CDbl(txtMoney.Text) - dblMoney), "###0.00;-###0.00;;")
            If lblMoney.Tag <> "" Then lblRepairMoney.Caption = "补交额:" & Format((CDbl(lblMoney.Tag) - dblMoney), "###0.00;-###0.00;;")
            lblMoney.Tag = ""
            txtMoney.Text = Format(dblMoney, "###0.00;-###0.00;;")
        End If
        
        '获取更新Sql
        Set cllThird = New Collection
        If GetYJThirdUpdateSQL(lng押金ID, IIf(Nvl(rsBalance_Out!卡号) = "", mstrBrushCardNo, Nvl(rsBalance_Out!卡号)), Nvl(rsBalance_Out!结算方式), Nvl(rsBalance_Out!交易金额, 0), Nvl(rsBalance_Out!结算号码), _
                            strSwapGlideNO, strSwapMemo, Nvl(rsBalance_Out!结算摘要), Nvl(rsBalance_Out!是否普通结算, 0), cllThird) Then
            
            '提交数据
            zlExecuteProcedureArrAy cllThird, Me.Caption, False, True
            blnTrans = False
            zlInterfacePrayMoney = True
        End If
        '获取扩展信息Sql
        If zlGetThreeSwapExpendSQL(mlngCardTypeID, CStr(lng押金ID), IIf(Nvl(rsBalance_Out!卡号) = "", mstrBrushCardNo, Nvl(rsBalance_Out!卡号)), rsExpend_Out, cllThirdExpend) Then
            '单独事务，不影响其他数据保存
            On Error GoTo ErrExpend:
            zlExecuteProcedureArrAy cllThirdExpend, Me.Caption
        End If
        If dblMoney <> Nvl(rsBalance_Out!交易金额, 0) Then txtMoney.Text = Nvl(rsBalance_Out!交易金额, 0)
    End If

    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrExpend:
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetThreeSwapExpendSQL(ByVal lng卡类别ID As Long, ByVal str押金IDs As String, ByVal str卡号 As String, _
                                        ByVal rsExpend As ADODB.Recordset, ByRef cllTirdExpend As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取扩展信息保存的SQL给集合
    '入参:
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:
    '日期:2018-03-27 17:33:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str交易信息 As String, strTemp As String
    On Error GoTo errHandle
    
    If cllTirdExpend Is Nothing Then Set cllTirdExpend = New Collection
    If rsExpend Is Nothing Then zlGetThreeSwapExpendSQL = True: Exit Function
    If rsExpend.State <> 1 Then zlGetThreeSwapExpendSQL = True: Exit Function
    
    With rsExpend
        rsExpend.Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        str交易信息 = ""
        Do While Not .EOF
            If Nvl(!项目名称) <> "" Then
                strTemp = Nvl(!项目名称) & "|" & Nvl(!项目内容)
                If zlCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
                        str交易信息 = Mid(str交易信息, 3)
                        'Zl_三方结算交易_Insert
                        strSQL = "Zl_三方结算交易_Insert("
                        '卡类别id_In 病人预交记录.卡类别id%Type,
                        strSQL = strSQL & "" & lng卡类别ID & ","
                        '消费卡_In   Number,
                        strSQL = strSQL & "" & 0 & ","
                        '卡号_In     病人预交记录.卡号%Type,
                        strSQL = strSQL & "'" & str卡号 & "',"
                        '结帐ids_In  Varchar2,
                        strSQL = strSQL & "'" & str押金IDs & "',"
                        '交易信息_In Varchar2:交易项目|交易内容||...
                        strSQL = strSQL & "'" & str交易信息 & "',"
                        '预交款缴款_In Number := 0
                        strSQL = strSQL & "2,"
                        '结算方式_In   病人预交记录.结算方式%Type := Null
                        strSQL = strSQL & "Null,"
                        '预交id_In     病人预交记录.Id%Type := Null
                        strSQL = strSQL & "Null,"
                        '性质_In       三方结算交易.性质%Type := Nul
                        strSQL = strSQL & "2)"
                        zlAddArray cllTirdExpend, strSQL
                        str交易信息 = ""
                End If
                str交易信息 = str交易信息 & "||" & strTemp
            End If
            .MoveNext
        Loop
        
    End With
    If str交易信息 <> "" Then
        str交易信息 = Mid(str交易信息, 3)
        'Zl_三方结算交易_Insert
        strSQL = "Zl_三方结算交易_Insert("
        '卡类别id_In 病人预交记录.卡类别id%Type,
        strSQL = strSQL & "" & lng卡类别ID & ","
        '消费卡_In   Number,
        strSQL = strSQL & "" & 0 & ","
        '卡号_In     病人预交记录.卡号%Type,
        strSQL = strSQL & "'" & str卡号 & "',"
        '结帐ids_In  Varchar2,
        strSQL = strSQL & "'" & str押金IDs & "',"
        '交易信息_In Varchar2:交易项目|交易内容||...
        strSQL = strSQL & "'" & str交易信息 & "',"
        '预交款缴款_In Number := 0
        strSQL = strSQL & "2,"
        '结算方式_In   病人预交记录.结算方式%Type := Null
        strSQL = strSQL & "Null,"
        '预交id_In     病人预交记录.Id%Type := Null
        strSQL = strSQL & "Null,"
        '性质_In       三方结算交易.性质%Type := Nul
        strSQL = strSQL & "2)"
        zlAddArray cllTirdExpend, strSQL
    End If
    zlGetThreeSwapExpendSQL = True
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
    If mbytInState = EM_浏览单据 Or mbytInState = EM_退押金 Then Exit Sub
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
Private Sub LoadPatiPage(ByVal lng病人id As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人的住院次数
    '编制:刘兴洪
    '日期:2012-12-11 10:19:58
    '说明:
    '问题:51628
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim bln留观 As Boolean
    On Error GoTo errHandle
        
    cboPatiPage.Clear
    With cboPatiPage
        
        If GetPatiPageNum(lng病人id, rsTemp) = False Then Exit Sub
        If rsTemp.State = 0 Then Exit Sub
        
        Do While Not rsTemp.EOF
            If bln留观 = False And Val(Nvl(rsTemp!病人性质, 0)) <> 0 Then bln留观 = True
            If Val(Nvl(rsTemp!主页ID)) = 0 And Val(Nvl(rsTemp!病人性质)) = 0 Then
                .AddItem "预约入院"
            Else
                .AddItem "第" & rsTemp!主页ID & "次" & IIf(Val("" & rsTemp!病人性质) = 1, "(门诊留观)", IIf(Val("" & rsTemp!病人性质) = 2, "(住院留观)", ""))
            End If
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!主页ID))
            If .ListIndex < 0 Then .ListIndex = .NewIndex
            If Val(Nvl(rsTemp!主页ID)) = mpatiInfo.主页ID Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If bln留观 = True Then Call cbo.SetListWidth(cboPatiPage.hwnd, 2000)
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
    Dim str病人id As String, PatiPageInfo As New clsPatientInfo
    Dim lng病人id As Long, lng主页ID As Long
    
    On Error GoTo errHandle
    If mbln未入科不交预交 = False Then Check未入科不交预交 = True: Exit Function
    '不诊预交不检查
    If cboType.ItemData(cboType.ListIndex) <> 2 Then Check未入科不交预交 = True: Exit Function
    '当前住院次数不为在院的,也不检查
    If Not mpatiInfo.在院 Then Check未入科不交预交 = True: Exit Function
    lng病人id = mpatiInfo.病人ID
    '不存在住院次数的,也能缴预交,因此不检查
    If cboPatiPage.ListIndex < 0 Then Check未入科不交预交 = True: Exit Function
    lng主页ID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    str病人id = lng病人id & ":" & lng主页ID
    Call GetPatiPageInforByID(str病人id, PatiPageInfo, False)
    If PatiPageInfo.已入科 = False Then
        MsgBox "注意" & vbCrLf & "   病人『" & mpatiInfo.姓名 & "』未入科,不允许缴押金!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Check未入科不交预交 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OptOthersErrBill(ByVal lng病人id As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:收款操作，检测病人是否有其他操作员产生的异常单据，并处理
    '入参: lng病人ID
    '编制:
    '日期:2018-08-07
    '说明:
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsErrBills As ADODB.Recordset
    Dim str操作员姓名 As String, strTittle
    On Error GoTo errHandle
    '有权限，且为收费状态
    If mbytInState <> EM_收押金 Then Exit Function
    'type: 1-异常充值，2-异常销帐
    strSQL = "Select Type, No , 卡号 ,操作员姓名" & vbNewLine & _
            "From (Select 1 Type, a.No, a.卡号, a.操作员姓名" & vbNewLine & _
            "       From 病人押金记录 a" & vbNewLine & _
            "       Where Nvl(校对标志, 0) <> 0 And 记录状态 = 0 And 病人id = [1]" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 2 Type, a.No, a.卡号, a.操作员姓名" & vbNewLine & _
            "       From 病人押金记录 a" & vbNewLine & _
            "       Where Nvl(校对标志, 0) <> 0 And 病人id = [1] And 记录状态 = 2)" & vbNewLine & _
            "Order By Decode(操作员姓名, [2], 0, 1), Type"
    Set rsErrBills = zlDatabase.OpenSQLRecord(strSQL, "病人异常单据查询", lng病人id, UserInfo.姓名)
    If rsErrBills.EOF Then Exit Function
    
    str操作员姓名 = Nvl(rsErrBills!操作员姓名)
    If Nvl(rsErrBills!type) = 1 Then
        strTittle = "收款"
    ElseIf Nvl(rsErrBills!type) = 2 Then
        strTittle = "销帐"
    End If
    '其他操作员判断权限
    If str操作员姓名 <> UserInfo.姓名 Then
        If InStr(mstrPrivs, ";允许处理他人异常单据;") = 0 Then Exit Function
        If MsgBox("注意:" & vbCrLf & _
            "       该病人存在由操作员【" & str操作员姓名 & "】处产生的异常" & strTittle & "单据！" & vbCrLf & vbCrLf & _
            "       是否对该单据进行" & strTittle & "？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    Else
        If MsgBox("注意:" & vbCrLf & _
            "       该病人存在异常" & strTittle & "单据，是否现在对该单据进行处理？", _
                    vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    mstrInNO = Nvl(rsErrBills!NO)
    mstrBrushCardNo = Nvl(rsErrBills!卡号)
    If Nvl(rsErrBills!type) = 1 Then
        mbytInState = EM_异常重收
    Else
        mbytInState = EM_异常重退
    End If
    Call InitFace
    '初始化病人信息
    Call InitPatientInfo(mstrInNO)
    If mbytInState = EM_异常重退 Then txtMoney.Text = Abs(txtMoney.Text)
    Call SetCtrlEnabled
    mblnOptErrBill = True
    OptOthersErrBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RestorePayStyle()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:恢复到上次选择的支付方式
    '说明:lblStyle.Tag记录的是上次选择的支付方式
    '       cboStyle.Tag记录的是缺省的支付方式
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    On Error GoTo errHandle
    
    If lblStyle.Tag = "" Then Exit Sub
    '有上次选择的支付方式,恢复
    intIndex = Val(lblStyle.Tag)
    lblStyle.Tag = ""
    If intIndex > cboStyle.ListCount - 1 Then cboStyle.ListIndex = Val(cboStyle.Tag): Exit Sub
    cboStyle.ListIndex = intIndex

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LocatePayStyle(ByVal lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡类别ID,定位到指定的支付类别上
    '入参:lngCardTypeID-卡类别ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnFind As Boolean, i As Integer
    If lngCardTypeID = 0 Then Exit Function
    If mobjPayMode Is Nothing Then Exit Function
    If mobjPayMode.Count = 0 Then Exit Function
    With cboStyle
        For i = 1 To mobjPayMode.Count
            If mobjPayMode(i).接口序号 = lngCardTypeID Then
                cboStyle.ListIndex = cbo.FindIndex(cboStyle, mobjPayMode(i).名称)
                blnFind = True: Exit For
            End If
        Next
    End With
    LocatePayStyle = blnFind
End Function

Private Sub LoadOriginReturnMoneyStyle(Optional ByVal bln缺省退现 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人原始退款方式
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mYJinfo.lng卡类别ID = 0 Or mYJinfo.str名称 = "" Then Exit Sub
    cboStyle.AddItem mYJinfo.str名称
    cboStyle.ItemData(cboStyle.NewIndex) = -1
    If Not bln缺省退现 Then cboStyle.ListIndex = cboStyle.NewIndex
End Sub

Private Function ZlGetParaConfig(ByVal lng卡类别ID As Long, ByVal intPara As Long, _
                                                    Optional strErrMsg As String, Optional strExpend As String) As Boolean
    ZlGetParaConfig = gobjSquare.objSquareCard.ZlGetParaConfig(Me, lng卡类别ID, False, intPara, strErrMsg, strExpend)
End Function

Private Function GetPatiInfo(ByVal lng病人id As Long, ByVal lng主页ID As Long, ByRef patiinfo As clsPatientInfo) As Boolean
    '功能：根据病人id和主页id获取病人信息和病案主页中的信息
    '入参：lng病人id-病人id
    '          lng主页id-主页id=-1时表示查询最后一次住院的信息,否则表示读取指定住院次数的信息（主页id=0表示预入院）
    '出差：PatiInfo-病人信息中的信息
    '       ：PatiPageInfo-病案主页中的信息
    '返回：获取成功返回true,否则返回false
    Dim PatiPageInfo As New clsPatientInfo
    Dim str病人id As String, blnLastTime As Boolean
    On Error GoTo errHandle
    
    If GetPatiInforFromPatiID(lng病人id, patiinfo) = False Then Exit Function
    If patiinfo.病人ID = 0 Then Exit Function
    blnLastTime = lng主页ID = -1
    If blnLastTime Then
        '读取最后一次住院的信息
        str病人id = lng病人id
    Else
        '读取指定住院次数住院的信息
        str病人id = lng病人id & ":" & lng主页ID
    End If
    patiinfo.住院状态 = 9
    If GetPatiPageInforByID(str病人id, PatiPageInfo, blnLastTime) = False Then GetPatiInfo = True: Exit Function
    If PatiPageInfo.病人ID > 0 Then
        patiinfo.住院状态 = 0
        patiinfo.当前病区ID = PatiPageInfo.当前病区ID
        patiinfo.出院科室ID = PatiPageInfo.出院科室ID
        patiinfo.医疗付款方式 = IIf(Val(PatiPageInfo.主页ID) = 0, patiinfo.医疗付款方式, PatiPageInfo.医疗付款方式)
        patiinfo.主页ID = PatiPageInfo.主页ID
        If patiinfo.病人类型 = "" Then patiinfo.病人类型 = PatiPageInfo.病人类型
        patiinfo.姓名 = IIf(PatiPageInfo.姓名 = "", patiinfo.姓名, PatiPageInfo.姓名)
        patiinfo.性别 = IIf(PatiPageInfo.性别 = "", patiinfo.性别, PatiPageInfo.性别)
        patiinfo.床号 = PatiPageInfo.床号
        patiinfo.费别 = IIf(PatiPageInfo.费别 = "", patiinfo.费别, PatiPageInfo.费别)
        patiinfo.病人性质 = IIf(PatiPageInfo.病人性质 = 0, patiinfo.病人性质, PatiPageInfo.病人性质)
        patiinfo.病人备注 = IIf(PatiPageInfo.病人备注 = "", patiinfo.病人备注, PatiPageInfo.病人备注)
        patiinfo.出院科室ID = PatiPageInfo.出院科室ID
        patiinfo.已入科 = PatiPageInfo.已入科
    End If
    GetPatiInfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub MakeCardsFrom结算方式(ByVal str名称 As String, ByVal str结算方式 As String, _
                 Optional ByVal lng卡类别ID As Long)
    '功能：根据结算方式构建Cards对象
    Dim objCard As Card
    Set objCard = New Card
    If mobjPayMode Is Nothing Then Set mobjPayMode = New Collection
    objCard.名称 = str名称
    objCard.结算方式 = str结算方式
    objCard.接口序号 = lng卡类别ID
    mobjPayMode.Add objCard, "_" & str名称
End Sub


