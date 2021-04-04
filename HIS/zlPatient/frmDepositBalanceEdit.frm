VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{CC0839AF-B32F-436B-8884-BE2BB3B4C73F}#4.1#0"; "zlIDKind.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmDepositBalanceEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "预交款单据"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13200
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDepositBalanceEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   13200
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   75
      ScaleHeight     =   2400
      ScaleWidth      =   13080
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   585
      Width           =   13080
      Begin VB.Label lblMemo 
         AutoSize        =   -1  'True
         Caption         =   "备    注 "
         Height          =   240
         Left            =   6345
         TabIndex        =   30
         Tag             =   "备    注 "
         Top             =   1995
         Width           =   1080
      End
      Begin VB.Label lblWorkUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工作单位"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6345
         TabIndex        =   29
         Tag             =   "工作单位 "
         Top             =   1605
         Width           =   960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   20
         X1              =   4080
         X2              =   6240
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   3050
         TabIndex        =   28
         Tag             =   "身份证号 "
         Top             =   1995
         Width           =   960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   18
         X1              =   1260
         X2              =   2830
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
         TabIndex        =   27
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
         Left            =   3060
         TabIndex        =   26
         Tag             =   "未缴费用 "
         ToolTipText     =   "未缴款的划价单费用合计"
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   19
         X1              =   4080
         X2              =   6240
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lbl医保预结 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保预结 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6345
         TabIndex        =   25
         Tag             =   "医保预结 "
         ToolTipText     =   "医保预结金额"
         Top             =   840
         Width           =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   17
         X1              =   7350
         X2              =   9225
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   16
         X1              =   7350
         X2              =   12975
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   15
         X1              =   1260
         X2              =   2830
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   14
         X1              =   4080
         X2              =   6240
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   13
         X1              =   10450
         X2              =   12975
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   12
         X1              =   7350
         X2              =   9225
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   11
         X1              =   1245
         X2              =   2830
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   10
         X1              =   1260
         X2              =   8805
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   9
         X1              =   7350
         X2              =   12975
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   8
         X1              =   4080
         X2              =   6240
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   7
         X1              =   1245
         X2              =   2830
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   6
         X1              =   10450
         X2              =   12975
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   5
         X1              =   10450
         X2              =   12975
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   4
         X1              =   10450
         X2              =   12975
         Y1              =   345
         Y2              =   345
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   3
         X1              =   6465
         X2              =   8805
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   2
         X1              =   4380
         X2              =   5760
         Y1              =   345
         Y2              =   345
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   0
         X1              =   2595
         X2              =   3720
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   1
         X1              =   780
         X2              =   1920
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院科室 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   9390
         TabIndex        =   24
         Tag             =   "住院科室 "
         Top             =   471
         Width           =   1080
      End
      Begin VB.Label lbl未审费用 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未审费用 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   23
         Tag             =   "未审费用 "
         ToolTipText     =   "未审核的划价记账费用合计"
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Label lbl应收款 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应 收 款 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   9390
         TabIndex        =   22
         Tag             =   "应 收 款 "
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Label lbl医疗付款方式 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医疗付款方式 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   8910
         TabIndex        =   21
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
         TabIndex        =   20
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
         Left            =   3060
         TabIndex        =   19
         Tag             =   "担保金额 "
         Top             =   1605
         Width           =   1080
      End
      Begin VB.Label lbl担保人 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担 保 人 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Tag             =   "担 保 人 "
         Top             =   1605
         Width           =   1080
      End
      Begin VB.Label lbl费别等级 
         AutoSize        =   -1  'True
         Caption         =   "费别 "
         Height          =   240
         Left            =   5925
         TabIndex        =   17
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
         Left            =   2040
         TabIndex        =   16
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
         TabIndex        =   15
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
         Left            =   3060
         TabIndex        =   14
         Tag             =   "预交余额 "
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lbl床号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   3840
         TabIndex        =   13
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
         Left            =   6345
         TabIndex        =   12
         Tag             =   "剩余款额 "
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Label lbl费用余额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未结费用 "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   9390
         TabIndex        =   11
         Tag             =   "未结费用 "
         ToolTipText     =   "未审核的划价记账费用合计"
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lbl帐户余额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "帐户余额 "
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Tag             =   "帐户余额 "
         Top             =   840
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   13200
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9435
      Width           =   13200
      Begin VB.CommandButton cmdVoucherSet 
         Caption         =   "凭条打印设置(&V)"
         Height          =   420
         Left            =   3960
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   60
         Width           =   2025
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   420
         Left            =   150
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   60
         Width           =   1500
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   420
         Left            =   11625
         TabIndex        =   60
         ToolTipText     =   "热键:Esc"
         Top             =   45
         Width           =   1500
      End
      Begin VB.CommandButton cmdSetup 
         Caption         =   "收据打印设置(&S)"
         Height          =   420
         Left            =   1770
         TabIndex        =   63
         TabStop         =   0   'False
         ToolTipText     =   "热键：F10"
         Top             =   60
         Width           =   2025
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   420
         Left            =   10050
         TabIndex        =   58
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
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   13080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   13080
      Begin VB.TextBox txtFact 
         ForeColor       =   &H00C00000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7710
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "热键：F3"
         Top             =   90
         Width           =   2100
      End
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   10950
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "热键：F12"
         Top             =   90
         Width           =   2100
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据号"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6975
         TabIndex        =   1
         Top             =   150
         Width           =   720
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
         TabIndex        =   8
         Top             =   45
         Width           =   1875
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   10170
         TabIndex        =   6
         Top             =   150
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   10065
      Width           =   13200
      _ExtentX        =   23283
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmDepositBalanceEdit.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18309
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
      Height          =   6330
      Left            =   75
      ScaleHeight     =   6330
      ScaleWidth      =   13095
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3015
      Width           =   13095
      Begin VB.PictureBox picBalance 
         BorderStyle     =   0  'None
         Height          =   4335
         Left            =   8200
         ScaleHeight     =   4335
         ScaleWidth      =   4860
         TabIndex        =   65
         Top             =   1920
         Width           =   4855
         Begin VB.TextBox txt开户行 
            Height          =   360
            Left            =   1065
            MaxLength       =   50
            TabIndex        =   52
            Top             =   2100
            Width           =   3780
         End
         Begin VB.ComboBox cboStyle 
            Height          =   360
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   855
            Width           =   1380
         End
         Begin VB.ComboBox cboNote 
            Height          =   360
            Left            =   1065
            TabIndex        =   56
            Text            =   "cboNote"
            Top             =   3915
            Width           =   3780
         End
         Begin VB.ComboBox cboUnit 
            Height          =   360
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   3000
            Width           =   3780
         End
         Begin VB.TextBox txt帐号 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   1065
            MaxLength       =   50
            TabIndex        =   53
            Top             =   2550
            Width           =   3780
         End
         Begin VB.TextBox txtUnit 
            Height          =   360
            Left            =   1065
            MaxLength       =   50
            TabIndex        =   55
            Top             =   3480
            Width           =   3780
         End
         Begin VB.TextBox txtCode 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   1065
            MaxLength       =   30
            TabIndex        =   51
            Top             =   1650
            Width           =   3780
         End
         Begin MSMask.MaskEdBox txtMoney 
            Height          =   360
            Left            =   2445
            TabIndex        =   49
            Top             =   855
            Width           =   2385
            _ExtentX        =   4207
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
         Begin MSMask.MaskEdBox txtThirdTotal 
            Height          =   360
            Left            =   1065
            TabIndex        =   45
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   8388608
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox txtTotal 
            Height          =   360
            Left            =   1065
            TabIndex        =   47
            Top             =   435
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   8388608
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox txtCashTotal 
            Height          =   360
            Left            =   3435
            TabIndex        =   46
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
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
         Begin MSMask.MaskEdBox txt收款 
            Height          =   360
            Left            =   1065
            TabIndex        =   50
            Top             =   1275
            Width           =   3780
            _ExtentX        =   6668
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   8388608
            Enabled         =   0   'False
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
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "摘要"
            Height          =   240
            Left            =   540
            TabIndex        =   76
            Top             =   3975
            Width           =   480
         End
         Begin VB.Label lblCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结算号码"
            Height          =   240
            Left            =   60
            TabIndex        =   75
            Top             =   1710
            Width           =   960
         End
         Begin VB.Label lblMoney 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "退款"
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
            Left            =   510
            TabIndex        =   74
            Top             =   915
            Width           =   510
         End
         Begin VB.Label lbl缴款单位 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "缴款单位"
            Height          =   240
            Left            =   60
            TabIndex        =   73
            Top             =   3540
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开户行"
            Height          =   240
            Left            =   300
            TabIndex        =   72
            Top             =   2160
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "帐号"
            Height          =   240
            Left            =   540
            TabIndex        =   71
            Top             =   2595
            Width           =   480
         End
         Begin VB.Label lblUnit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "缴款科室"
            Height          =   240
            Left            =   60
            TabIndex        =   70
            Top             =   3060
            Width           =   960
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "退款合计"
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
            Left            =   180
            TabIndex        =   69
            Top             =   525
            Width           =   840
         End
         Begin VB.Label lblCashTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "退现合计"
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
            Left            =   2580
            TabIndex        =   68
            Top             =   75
            Width           =   840
         End
         Begin VB.Label lblThirdTotal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "三方退款"
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
            Left            =   180
            TabIndex        =   67
            Top             =   75
            Width           =   840
         End
         Begin VB.Label lbl找补 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "找补"
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
            Left            =   510
            TabIndex        =   66
            Top             =   1275
            Width           =   510
         End
      End
      Begin VB.ComboBox cboPatiPage 
         Height          =   360
         Left            =   11730
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   120
         Width           =   1305
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1245
         MaxLength       =   100
         TabIndex        =   39
         ToolTipText     =   "热键：F11"
         Top             =   120
         Width           =   2280
      End
      Begin VB.ComboBox cboType 
         Height          =   360
         Left            =   4650
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   120
         Width           =   1290
      End
      Begin VSFlex8Ctl.VSFlexGrid vsThirdTotal 
         Height          =   960
         Left            =   8400
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   855
         Width           =   4605
         _cx             =   8123
         _cy             =   1693
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDepositBalanceEdit.frx":115E
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
         ExplorerBar     =   7
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
      Begin VB.PictureBox picDepositBack 
         BorderStyle     =   0  'None
         Height          =   5625
         Left            =   120
         ScaleHeight     =   5625
         ScaleWidth      =   7815
         TabIndex        =   32
         Top             =   600
         Width           =   7815
         Begin VB.CommandButton cmdDefault 
            Caption         =   "全退(&A)"
            Height          =   420
            Left            =   6120
            TabIndex        =   44
            Top             =   5040
            Width           =   1500
         End
         Begin VSFlex8Ctl.VSFlexGrid vsBlance 
            Height          =   4935
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   7695
            _cx             =   13573
            _cy             =   8705
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   350
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmDepositBalanceEdit.frx":11BF
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
            Begin VB.Image imgDel 
               Height          =   240
               Left            =   75
               Picture         =   "frmDepositBalanceEdit.frx":12D5
               Top             =   45
               Visible         =   0   'False
               Width           =   240
            End
         End
      End
      Begin VB.PictureBox picDeposit 
         Height          =   5700
         Left            =   75
         ScaleHeight     =   5640
         ScaleWidth      =   7875
         TabIndex        =   31
         Top             =   555
         Width           =   7935
         Begin XtremeSuiteControls.TabControl tbPage 
            Height          =   735
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   2175
            _Version        =   589884
            _ExtentX        =   3836
            _ExtentY        =   1296
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picDepositHistory 
         BorderStyle     =   0  'None
         Height          =   1065
         Left            =   3720
         ScaleHeight     =   1065
         ScaleWidth      =   2535
         TabIndex        =   33
         Top             =   1920
         Width           =   2535
         Begin VSFlex8Ctl.VSFlexGrid vsDepositHistory 
            Height          =   645
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   2175
            _cx             =   3836
            _cy             =   1138
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   350
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
            ExplorerBar     =   7
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
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   600
         TabIndex        =   38
         Top             =   120
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
      Begin VB.Label lblPatientNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6060
         TabIndex        =   61
         Top             =   180
         Width           =   840
      End
      Begin VB.Label lblPatiPage 
         AutoSize        =   -1  'True
         Caption         =   "住院次数"
         Height          =   240
         Left            =   10730
         TabIndex        =   59
         Top             =   180
         Width           =   960
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   120
         TabIndex        =   42
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lbl预交类型 
         AutoSize        =   -1  'True
         Caption         =   "预交类型"
         Height          =   240
         Left            =   3615
         TabIndex        =   41
         Top             =   180
         Width           =   960
      End
      Begin VB.Label lblThirdSummary 
         Caption         =   "本次三方退款汇总"
         Height          =   255
         Left            =   8400
         TabIndex        =   36
         Top             =   550
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   -135
         X2              =   7755
         Y1              =   -30
         Y2              =   -30
      End
   End
   Begin MSCommLib.MSComm com 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   600
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDepositBalanceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'入口参数----------------------------------------------------------------------------------
Private mstrPrivs As String
Private mlngModul As Long
Private mbytCallObject As Byte '调用的对象(0-预交款管理调用;1-病人费用查询调用;2-医疗卡管理调用;3-挂号模块调用
Private mlng病人ID As Long, mlng主页ID As Long, mdblDefPreMoney As Double
Private mbytPrepayType As Byte   ' 1-门诊预交;2-住院预交(4时,1,门诊转住院;2时住院转门诊)
Private mblnNotClick As Boolean
'程序变量----------------------------------------------------------------------------------
Private mblnUnLoad  As Boolean '用于控制窗体直接退出
Private mdbl剩余款额 As Double
Private mdbl预交余额 As Double
Private mdbl费用余额 As Double
Private mlng领用ID As Long, mstrCardPrivs As String
Private mstr缺省结算方式 As String
Private mblnOK As Boolean, mstr退款操作员 As String
Private mbln未入科不交预交 As Boolean '51628
Private mbln住院退预交验证 As Boolean   '63113:刘尔旋,2013-10-29,住院预交退款验证
Private mbln允许在院病人余额退款 As Boolean
Private mblnNurseCall As Boolean
Private mblnFirst As Boolean
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

Private mstrBrushCardNo As String

Private Type Ty_BillInfor
    lng预交ID As Long
    strNO As String
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
Private mrsDepositBalance As ADODB.Recordset    '当前病人的预交余额
Private mbytBackMoneyType As Byte '退款方式:1-禁止;0-提示
Private mbytOracleBackType As Byte '退款检查_In;0-忽略退款金额是否大于了病人余额；1-检查退款金额
Private mblnClearWinInfor As Boolean  '缴款后,是否清除窗体信息
Private mblnCheckPass As Boolean '刷卡时要求输入密码,'0000000000'依位顺序表示各个场合,分别为:1.门诊挂号,2.门诊划价,3.门诊收费,4.门诊记帐,5.入院登记,6.住院记帐,7.病人结帐,8.病人预交款,9.检验技师站,10.影像医技站.'
Private mbln排除未缴及未审 As Boolean '剩余款排除未缴及未审金额
'外挂评价器对象
Private mobjPlugIn As Object
Private mstrPatiOld As String
Private mstrPatiSex As String
Private mblnOneCard As Boolean  '是否只有一张就诊卡
Private mlngFactModule As Long '发票相关参数模块号
Private mblnOptErrBill As Boolean '收费模式下处理异常单据
Private mobjThridSwap As clsThirdSwap
Private mobjPtDelItems As clsBalanceItems
Private Enum pg_Page
    pg_预交余额退款 = 1
    pg_预交历史记录 = 2
End Enum

Private Enum PaneId
    EM_Head = 1
    EM_PatiInfo = 2
    EM_BillList = 3
    EM_Cmd = 4
End Enum
Private mpatiInfo As New clsPatientInfo
Private mobjCards As New Cards  'zlOneCardComLib.Cards
Private mobjEInvoice As clsEinvoice '电子票据部件
Private mintPrintType As Integer

Private Sub zlInitBalanceGrid(Optional bln查看 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算列表
    '编制:刘兴洪
    '日期:2015-01-23 14:14:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    With vsBlance
    
        For i = 1 To .Rows - 1
            .RowData(i) = ""
        Next
        .Clear: .Rows = 2: i = 0: .Cols = 23
        .TextMatrix(0, i) = "卡类别ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "消费卡ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "结算性质": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "编辑状态": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "类型": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "退现": .ColWidth(i) = 600: i = i + 1
        .TextMatrix(0, i) = "结算状态": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否退现": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否全退": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "校对标志": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否密文": .ColWidth(i) = 0: i = i + 1
        
        .TextMatrix(0, i) = "单据号": .ColWidth(i) = 1500: i = i + 1
        .TextMatrix(0, i) = "退款方式": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "预交余额": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "退款金额": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "结算号码": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "卡类别名称": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "卡号": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "交易流水号": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "交易说明": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "备注": .ColWidth(i) = 2500: i = i + 1
        .TextMatrix(0, i) = "关联交易ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "是否转账": .ColWidth(i) = 0: i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = flexAlignLeftCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            
            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            Select Case .ColKey(i)
            Case "是否转账", "关联交易ID", "结算性质", "类型", "是否保存", "是否密文", "校对标志", "编辑状态", "是否退现", "是否全退", "结算状态", "是否验证"
                .ColHidden(i) = True
                .ColData(i) = "-1||1"
            Case "退现"
                .ColAlignment(i) = flexAlignRightCenter
                .ColData(i) = "1||0"
                .ColDataType(i) = flexDTBoolean
            Case "退款金额"
                .ColAlignment(i) = flexAlignRightCenter
                .ColData(i) = "1||0"
            Case "预交余额"
                .ColAlignment(i) = flexAlignRightCenter
                .ColData(i) = "0||0"
            Case .ColIndex("退款方式")
                .ColData(i) = """1||0"
            Case "卡类别名称"
                .ColData(i) = "1||2"
            Case .ColIndex("结算号码")
                .ColData(i) = "1||0"
            Case Else
                .ColData(i) = "1||" & IIf(bln查看, "0", "2")
            End Select
            If bln查看 Then .ColData(i) = ""
        Next
        If Not bln查看 Then .Editable = flexEDKbdMouse
        .ExplorerBar = flexExMove
    End With
    zl_vsGrid_Para_Restore mlngModul, vsBlance, Me.Name, "结算列表"
    vsBlance.ColWidth(vsBlance.ColIndex("退现")) = 600
    vsBlance.ColHidden(vsBlance.ColIndex("退现")) = False
    
    With vsDepositHistory
        .Clear: .Rows = 2: i = 0: .Cols = 7
        .TextMatrix(0, i) = "日期": .ColWidth(i) = 1350: i = i + 1
        .TextMatrix(0, i) = "单据号": .ColWidth(i) = 1110: i = i + 1
        .TextMatrix(0, i) = "票据号": .ColWidth(i) = 1110: i = i + 1
        .TextMatrix(0, i) = "科室": .ColWidth(i) = 1200: i = i + 1
        .TextMatrix(0, i) = "缴款金额": .ColWidth(i) = 1200: i = i + 1
        .TextMatrix(0, i) = "结算": .ColWidth(i) = 1000: i = i + 1
        .TextMatrix(0, i) = "收款人": .ColWidth(i) = 1000: i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            Select Case .ColKey(i)
            Case "日期", "单据号", "票据号"
                .ColAlignment(i) = flexAlignCenterCenter
            Case "缴款金额"
                 .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
    End With
    zl_vsGrid_Para_Restore mlngModul, vsDepositHistory, Me.Name, "预交清单"
    zl_vsGrid_Para_Restore mlngModul, vsThirdTotal, Me.Name, "三方退款汇总"
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlShowEdit(ByVal frmMain As Object, ByVal bytCallObject As Byte, ByVal objEInvoice As clsEinvoice, _
    ByVal strPrivs As String, ByVal lngModule As Long, Optional ByVal bytPrepayType As Byte = 0, _
    Optional ByVal lng病人id As Long = 0, Optional lng主页ID As Long = 0, Optional ByVal blnNurseCall As Boolean = False, _
    Optional blnOneCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口,用于病人预交款信息编辑或查看
    '入参:frmMain-调用的主窗口
    '        bytCallObject:调用的对象(0-预交款管理调用;1-病人费用查询调用;2-医疗卡调用,3-门诊挂号调用)...
    '        bytPrepayType-预交类型(0-门诊和住院;1-门诊;2-住院)
    '        strInNo:要浏览或退款的单据号(mbytInState=1或3时有效),从病人信息登记中调用退卡时为空
    '        blnNurseCall-护士站调用
    '出参:
    '返回:预交款只有一次成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-17 16:11:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    mbytCallObject = bytCallObject:   mstrPrivs = strPrivs: mlngModul = lngModule
    mlng病人ID = lng病人id: mlng主页ID = lng主页ID
    mbytPrepayType = bytPrepayType
    mblnNurseCall = blnNurseCall
    mblnOneCard = blnOneCard
    mlngFactModule = IIf(mbytCallObject = 2, 1107, mlngModul)
    Set mobjEInvoice = objEInvoice
    Set mobjThridSwap = New clsThirdSwap
    Call gOneCardData.InitCommon(gcnOracle)
    mblnOK = False
    If frmMain Is Nothing Then
        Me.Show
    Else
        Me.Show 1, frmMain
    End If
    zlShowEdit = mblnOK
End Function

Private Sub cboPatiPage_Click()
    If txtPatient.Tag <> "" And Not mpatiInfo.病人ID = 0 Then
        If cboPatiPage.ItemData(cboPatiPage.ListIndex) <> Val(cboPatiPage.Tag) Then
            cboPatiPage.Tag = cboPatiPage.ItemData(cboPatiPage.ListIndex)
            Call ShowPatiInfoFromPage
            Call ShowPremayBalance(True, mpatiInfo.病人ID)
            Call LoadThirdDelDeposit(Val(cboPatiPage.ItemData(cboPatiPage.ListIndex)))
        End If
    End If
    Call ShowHistoryPrepay
End Sub

Private Sub ShowPatiInfoFromPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人主页显示病人信息
    '编制:刘兴洪
    '日期:2018-11-29 09:46:39
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim lng主页ID As Long
    
    If cboType.ListIndex < 0 Then Exit Sub
    If cboType.ItemData(cboType.ListIndex) = 1 Then Exit Sub    '门诊预交，不存在主页
    If cboPatiPage.ListIndex < 0 Then Exit Sub  '无主页时，不处理
    lng主页ID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    
    '根据第几次入院更新信息
    Call GetPatient(IDKind.GetfaultCard, txtPatient.Tag, False, False, txtPatient.Tag, lng主页ID)
    '加载病人信息给控件
    Call LoadPatiInforToContronl
    
 End Sub
Private Sub LoadPatiInforToContronl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人信息给控件
    '编制:刘兴洪
    '日期:2018-11-29 09:51:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str担保人 As String, dbl担保额 As Double
     
    On Error GoTo errHandle
    
    If mpatiInfo.病人ID = 0 Then Exit Sub
    
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(mpatiInfo.住院号 = "", "", "住院号:" & mpatiInfo.住院号 & "  ") & _
                       IIf(mpatiInfo.门诊号 = "", "", "门诊号:" & mpatiInfo.门诊号)
                       
    txtPatient.IMEMode = 0: txtPatient.PasswordChar = ""    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.Text = mpatiInfo.姓名: txtPatient.Tag = mpatiInfo.病人ID
    
    lblSex.Caption = lblSex.Tag & mpatiInfo.性别: mstrPatiSex = mpatiInfo.性别
    lblOld.Caption = lblOld.Tag & mpatiInfo.年龄: mstrPatiOld = mpatiInfo.年龄
     
    lbl医疗付款方式.Caption = lbl医疗付款方式.Tag & mpatiInfo.医疗付款方式
    
    
    lbl床号.Caption = lbl床号.Tag
    If mpatiInfo.当前科室ID <> 0 Then
        lbl床号.Caption = lbl床号.Tag & IIf(mpatiInfo.床号 = "", "家庭", mpatiInfo.床号)
    End If
    
    lbl科室.Caption = lbl科室.Tag & GET部门名称(mpatiInfo.出院科室ID)
    
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(mpatiInfo.当前科室ID = 0, mpatiInfo.出院科室ID, mpatiInfo.当前科室ID))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    
    '医保改动-在院病人转个人帐户
    If mpatiInfo.险类 > 0 And InStr(mstrPrivs, ";保险转帐;") > 0 And mstr个人帐户 <> "" Then
        If cbo.FindIndex(cboStyle, mstr个人帐户, True) = -1 Then
            cboStyle.AddItem mstr个人帐户
            cboStyle.ItemData(cboStyle.NewIndex) = 3
        End If
        
        '医保接口
        mcur帐户余额 = gclsInsure.SelfBalance(mpatiInfo.病人ID, mpatiInfo.医保号, 30, , mpatiInfo.险类)
        lbl帐户余额.Caption = lbl帐户余额.Tag & Format(mcur帐户余额, "0.00")
        lbl帐户余额.Visible = True
        lbl预交余额.Left = lbl未缴费用.Left
        If lbl帐户余额.Visible Then
            Line2(14).Visible = True: Line2(11).x2 = Line2(7).x2
        Else
            Line2(14).Visible = False: Line2(11).x2 = Line2(14).x2
        End If
    End If
    
    lbl费别等级.Caption = lbl费别等级.Tag & mpatiInfo.费别
    Call Get担保信息(mpatiInfo.病人ID, mpatiInfo.主页ID, dbl担保额, str担保人)
    lbl担保人.Caption = lbl担保人.Tag & str担保人
    lbl担保金额.Caption = lbl担保金额.Tag & Format(dbl担保额, "##,##0.00;-##,##0.00; ;")
    
    lbl手机号.Caption = lbl手机号.Tag & mpatiInfo.手机号
    lbl身份证号.Caption = lbl身份证号.Tag & mpatiInfo.身份证号
    
    lblMemo.Caption = lblMemo.Tag & mpatiInfo.病人备注
    lblWorkUnit.Caption = lblWorkUnit.Tag & mpatiInfo.工作单位
    lbl家庭地址.Caption = lbl家庭地址.Tag & mpatiInfo.家庭地址
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboPatiPage_KeyDown(KeyCode As Integer, Shift As Integer)
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboType_Click()

    If cboType.ListIndex < 0 Then Exit Sub
    If Val(cboType.Tag) = cboType.ListIndex Then Exit Sub
    cboType.Tag = cboType.ListIndex
    
    mlng领用ID = 0
    mFactProperty = zl_GetInvoicePreperty(mlngFactModule, 2, cboType.ItemData(cboType.ListIndex))
    Call GetFact
    If cboType.Text = "住院预交" Then
        If cboPatiPage.ListCount > 0 Then cboPatiPage.Tag = "0": cboPatiPage.ListIndex = 0
    End If
    Call ShowPremayBalance(True, 0)
    '重新加载当前余额退款信息
    If Not mblnNotClick Then
        Call ShowHistoryPrepay
        Call LoadThirdDelDeposit
    End If
    Call SetCtrlEnabled
    
    lblPatiPage.Visible = cboType.Text = "住院预交": cboPatiPage.Visible = cboType.Text = "住院预交"
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub
   
Private Sub cmdDefault_Click()
    Call ReCalePtBalanceMoney(2)
End Sub

Private Sub cmdVoucherSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1103_2", Me)
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
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
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
    cmdOK.Enabled = mpatiInfo.病人ID > 0
End Sub


Private Sub SetCtrlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Enabled属性
    '编制:刘兴洪
    '日期:2011-07-24 09:30:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean, objCtl As Control
    Dim int性质 As Integer
    
    If cboStyle.ListIndex >= 0 Then int性质 = cboStyle.ItemData(cboStyle.ListIndex)
    blnEdit = True
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
    Dim lngIndex As Long
    
    If cboStyle.ListIndex = -1 Then Exit Sub
        
    '问题号:111657,焦博,2017/07/25,使用现金支付预交款时,任会产生三方卡号
    mstrBrushCardNo = ""     '清空三方交易时缓存的卡号
    mcurBill.bln转账 = False
    mcurBill.lng预交ID = 0
    lngIndex = cboStyle.ListIndex + 1
    Call SetCtrlEnabled
    txtMoney.Enabled = True
    
    Select Case cboStyle.ItemData(cboStyle.ListIndex)
    Case 3, 1
        txtUnit.Text = "": txt开户行.Text = "": txt帐号.Text = ""
    Case 2
        If cboStyle.Text Like "*票*" Or cboStyle.Text Like "*卡*" Then
            If mpatiInfo.病人ID = 0 Then Exit Sub
            strInfo = GetLastInfo(mpatiInfo.病人ID)
            If strInfo <> "" Then
                txtUnit.Text = IIf(Split(strInfo, "|")(0) = "", txtUnit.Text, Split(strInfo, "|")(0))
                txt开户行.Text = IIf(Split(strInfo, "|")(1) = "", txt开户行.Text, Split(strInfo, "|")(1))
                txt帐号.Text = IIf(Split(strInfo, "|")(2) = "", txt帐号.Text, Split(strInfo, "|")(2))
                txtCode.Text = IIf(Split(strInfo, "|")(3) = "", txtCode.Text, Split(strInfo, "|")(3))
            End If
        End If
    End Select
End Sub

Private Sub cboStyle_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then
        If cboStyle.ListIndex = -1 Then Beep: Exit Sub
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    If cboStyle.Locked Then Exit Sub
    If KeyAscii >= 32 Then
        lngIdx = cbo.MatchIndex(cboStyle.hwnd, KeyAscii)
        If lngIdx = -1 And cboStyle.ListCount > 0 Then lngIdx = 0
        cboStyle.ListIndex = lngIdx
    End If
End Sub

Private Sub cboStyle_Validate(Cancel As Boolean)
    If cboStyle.Locked Or cboStyle.ListIndex = -1 Then Exit Sub
    
    If InStr(1, mstrPrivs, ";预交退款;") = 0 Then
        MsgBox "你没有权限进行预交退款操作！", vbInformation, gstrSysName
        If cbo.Locate(cboStyle, BalanceType.C5代收款, True) Then Cancel = True
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
Private Sub cmdCancel_Click()
    If Not mblnOK Then Unload Me: Exit Sub
    If mpatiInfo.病人ID > 0 Then
        If MsgBox("该病人的尚未进行退款操作,确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Function CheckDataValied(ByRef objSetFocus As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查数据是否合法
    '返回：合法返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-18 16:38:39
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng病人id As Long
    Dim int预交类别 As Integer, dblCashTotal As Double, dblCash As Double, dblPt As Double
    Dim objItem As clsBalanceItem, i As Long

    On Error GoTo errHandle
        
    If mpatiInfo.病人ID = 0 Then
        lng病人id = Val(txtPatient.Tag)
    Else
        lng病人id = mpatiInfo.病人ID
    End If
    
    '退款操作
    If InStr(1, mstrPrivs, ";预交退款;") = 0 Then
        MsgBox "你没有权限进行预交退款操作！", vbInformation, gstrSysName: Exit Function
    End If
    
    If mpatiInfo.病人ID = 0 Then
        MsgBox "没有确定退预交款的病人,不能退款！", vbExclamation, gstrSysName
        Set objSetFocus = txtPatient
        Exit Function
    End If
     
    If LenB(StrConv(txtUnit.Text, vbFromUnicode)) > 50 Then
        MsgBox "缴款单位名称只允许 50 个字符或 25 个汉字,请修改！", vbInformation, App.Title
        Set objSetFocus = txtUnit
        Exit Function
    End If
    
    If LenB(StrConv(txt开户行.Text, vbFromUnicode)) > 50 Then
        MsgBox "开户行名称只允许 50 个字符或 25 个汉字,请修改！", vbInformation, App.Title
        Set objSetFocus = txt开户行
        Exit Function
    End If
    
    If LenB(StrConv(cboNote.Text, vbFromUnicode)) > 50 Then
        MsgBox "缴款摘要只允许 50 个字符或 25 个汉字,请修改！", vbInformation, App.Title
        Set objSetFocus = cboNote
        Exit Function
    End If
     
    If CCur(StrToNum(txtCashTotal.Text)) = 0 And CCur(StrToNum(txtThirdTotal.Text)) = 0 Then
        MsgBox "退款金额不能为空或零,请输入！", vbExclamation, gstrSysName
        Set objSetFocus = txtCashTotal
        Exit Function
    End If
    
    If StrToNum(txtCashTotal.Text) < 0 Then
        Call MsgBox("该病人无余额可退,不能进行余额退款操作！", vbInformation + vbOKOnly, gstrSysName)
        Set objSetFocus = txtCashTotal
        Exit Function
    End If
    
    If Val(lblCashTotal.Tag) <> StrToNum(txtCashTotal.Text) Then
         If MsgBox("你当前输入的退款金额与退款列表中的退现金额不一致,是否自动计算退现金额？" & vbCrLf & vbCrLf & _
               "退现合计:" & Format(Val(lblCashTotal.Tag), "###0.00###") & vbCrLf & _
               "输入金额:" & Format(StrToNum(txtCashTotal.Text), "###0.00###") & vbCrLf & _
               "", vbQuestion + vbYes + vbDefaultButton1, gstrSysName) = vbNo Then Set objSetFocus = txtCashTotal: Exit Function
        '自动分摊
        Call AutoShareBalanceMoney(StrToNum(txtCashTotal.Text))
        If Val(lblCashTotal.Tag) <> StrToNum(txtCashTotal.Text) Then
            MsgBox "未分摊完成，请检查!", vbInformation + vbOKOnly, gstrSysName
            Set objSetFocus = txtCashTotal: Exit Function
        End If
    Else
        '可能存在三方退现金额录入过小，造成的，所以需提示并重新计算
        With vsBlance
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("退款方式")) <> "" Then
                    If zlGetBalanceItemFromBalanceGrid(i, objItem) = False Then Exit Function
                    If GetVsGridBoolColVal(vsBlance, i, .ColIndex("退现")) Then
                          dblCashTotal = roundEx(dblCashTotal + objItem.剩余金额, 5)
                          dblCash = roundEx(dblCash + objItem.结算金额, 5)
                    End If
                End If
            Next
        End With
        
        If dblCash < StrToNum(txtCashTotal.Text) Then
           If MsgBox("你输入的退款金额大于了退现合计，是否重新计算退款？" & vbCrLf & "输入金额:" & txtCashTotal.Text & vbCrLf & "退现合计:" & Format(dblPt + dblCash, "###0.00"), vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call ReCalePtBalanceMoney(1) '只计算退现部分
           End If
           Exit Function
        End If
    End If

    If Val(lblCashTotal.Tag) <> StrToNum(txtMoney.Text) - StrToNum(txt收款.Text) Then
        Call MsgBox("当前退款金额(" & Format(CCur(StrToNum(txtMoney.Text) - StrToNum(txt收款.Text)), "0.00") & ")与本次退现合计(" & Format(Val(lblCashTotal.Tag), "0.00") & ")不一致,不能退款!", vbInformation + vbOKOnly, gstrSysName)
        Set objSetFocus = txtMoney
        Exit Function
    End If
    
    If mdbl剩余款额 - CCur(StrToNum(txtCashTotal.Text)) - CCur(StrToNum(txtThirdTotal.Text)) < 0 Then
        If mbytBackMoneyType = 1 Then
            Call MsgBox("退款金额(" & Format(CCur(StrToNum(txtCashTotal.Text)) + CCur(StrToNum(txtThirdTotal.Text)), "0.00") & ")大于了病人当前的剩余款(" & Format(mdbl剩余款额, "0.00") & "),不能退款!", vbInformation + vbOKOnly, gstrSysName)
            Set objSetFocus = txtCashTotal
            Exit Function
        Else
            If MsgBox("退款金额(" & Format(CCur(StrToNum(txtCashTotal.Text)) + CCur(StrToNum(txtThirdTotal.Text)), "0.00") & ")大于了病人当前的剩余款(" & Format(mdbl剩余款额, "0.00") & "),忽略吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Set objSetFocus = txtCashTotal
                Exit Function
            End If
            mbytOracleBackType = 0
        End If
    End If

    If cboStyle.ListIndex = -1 And CCur(StrToNum(txtMoney.Text)) <> 0 Then
        MsgBox "未确定当前退款方式，不能进行余额退款！", vbExclamation, gstrSysName
        Set objSetFocus = cboType
        Exit Function
    End If
    
    If cboStyle.ListIndex >= 0 Then
        If mobjThridSwap.objPayCards(cboStyle.ItemData(cboStyle.ListIndex)).结算性质 = 3 Then
            MsgBox "医保病人个人帐户转帐金额不能进行余额退款操作。", vbInformation, gstrSysName
            Set objSetFocus = txtMoney
            Exit Function
        End If
    End If
    
    If cboType.ListIndex >= 0 Then int预交类别 = cboType.ItemData(cboType.ListIndex)
    Select Case int预交类别
    Case 1 '门诊预交
        If InStr(1, mstrPrivs, ";门诊病人余额退款;") = 0 Then
           MsgBox "你没有权限进行余额退款操作,请与管理员联系授予余额退款权限！", vbInformation, gstrSysName: Exit Function
        End If
                
        If gbyt预存款消费验卡 <> 0 Then
            If CreatePublicExpense() Then
                If Not gobjPublicExpense.zlPatiIdentify(mlngModul, Me, lng病人id, Val(StrToNum(txtCashTotal.Text)), True) Then
                    Set objSetFocus = cboType
                    Exit Function
                End If
            End If
        End If

    Case 2 '住院预交
    
        If mbln允许在院病人余额退款 = False And mpatiInfo.在院 Then
            MsgBox "病人在院,不能进行余额退款,请检查！", vbInformation, gstrSysName
            Set objSetFocus = txtMoney
            Exit Function
        End If
    
         If Not mblnNurseCall And InStr(1, mstrPrivs, ";在院病人余额退款;") = 0 And mpatiInfo.在院 Then
            MsgBox "你没有权限对在院病人进行余额退款操作,请与管理员联系授予余额退款权限！", vbInformation, gstrSysName: Exit Function
         End If
         
         If Not mblnNurseCall And InStr(1, mstrPrivs, ";出院病人余额退款;") = 0 And Not mpatiInfo.在院 Then
            MsgBox "你没有权限对出院病人进行余额退款操作,请与管理员联系授予余额退款权限！", vbInformation, gstrSysName: Exit Function
         End If
         
        If gbyt预存款消费验卡 <> 0 And mbln住院退预交验证 Then
            If CreatePublicExpense() Then
                If Not gobjPublicExpense.zlPatiIdentify(mlngModul, Me, lng病人id, Val(StrToNum(txtCashTotal.Text)), True) Then
                    Set objSetFocus = cboType
                    Exit Function
                End If
            End If
        End If
        
    Case Else
        MsgBox "未选择余额退款的预交类型,请输入", vbExclamation, gstrSysName
        Set objSetFocus = txtMoney
        Exit Function
    End Select
    CheckDataValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function CheckFactIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查发票是否有效(同时生成发票号)
    '返回:发票合返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-31 09:32:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
     
    Dim int预交类型 As Integer
    
    On Error GoTo errHandle
    If cboType.ListIndex <> -1 Then int预交类型 = cboType.ItemData(cboType.ListIndex)
    If mobjEInvoice.zlIsStartEInvoice(0, int预交类型) Then
        If mobjEInvoice.zlGetTranPaperInvoiceModule = 0 Then CheckFactIsValied = True: Exit Function
        If Trim(txtFact.Text) = "" Then
            MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
            zlControl.ControlSetFocus txtFact: Exit Function
        End If
        CheckFactIsValied = True: Exit Function
    End If
  
    If mFactProperty.intInvoicePrint = 0 Then CheckFactIsValied = True: Exit Function
    If Trim(txtFact.Text) = "" Then Call GetFact
    
    '票据号码检查
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
                    txtFact.SetFocus
            End Select
            txtFact.Text = ""
            Exit Function
        End If
        CheckFactIsValied = True: Exit Function
    End If
    
    If Len(txtFact.Text) <> gbyt预交 And txtFact.Text <> "" Then
        MsgBox "票据号码长度应该为 " & gbyt预交 & " 位！", vbInformation, gstrSysName
        txtFact.SetFocus: Exit Function
    End If
    CheckFactIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetItemsFromRecord(ByVal int预交类型 As Integer, ByVal dblMoney As Double, ByVal rsMoney As ADODB.Recordset, ByRef objItems_out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据记录集，获取所有的结算项
    '入参:dblMoney-当前分摊金额
    '     rsMoney-当前记录集
    '出参:objItems_out-返回分摊项信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-09-11 11:38:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, bln消费卡 As Boolean
    Dim dblTemp As Double, objItem As clsBalanceItem
    Dim objCard  As Card
    
    On Error GoTo errHandle
    If dblMoney = 0 Then GetItemsFromRecord = True: Exit Function
    
    If objItems_out Is Nothing Then Set objItems_out = New clsBalanceItems
    If rsMoney.RecordCount = 0 Then GetItemsFromRecord = True: Exit Function
    
    rsMoney.MoveFirst
    rsMoney.Sort = "收款时间"
    With rsMoney
        Do While Not .EOF
            lngCardTypeID = Val(Nvl(!卡类别ID))
            bln消费卡 = Val(Nvl(!消费卡)) = 1
            dblTemp = roundEx(Val(Nvl(!冲预交)), 6)
            If dblTemp <> 0 Then
                
                If dblMoney > dblTemp Then
                    dblMoney = roundEx(dblMoney - dblTemp, 6)
                Else
                    dblTemp = dblMoney: dblMoney = 0
                End If
                rsMoney!冲预交 = roundEx(Nvl(!冲预交, 0) - dblTemp, 6)
                rsMoney.Update
                Set objCard = mobjThridSwap.zlGetCardFromCardType(lngCardTypeID, bln消费卡, Nvl(!结算方式))
                Set objItem = New clsBalanceItem
                Set objItem.objCard = objCard
                objItem.结算方式 = Nvl(!结算方式)
                If objItem.结算方式 = "" Then objItem.结算方式 = objCard.结算方式
                objItem.卡号 = Nvl(!卡号)
                objItem.卡类别ID = lngCardTypeID
                objItem.是否允许删除 = True
                objItem.预交ID = Val(Nvl(!预交ID))
                objItem.消费卡 = bln消费卡
                objItem.校对标志 = 1
                objItem.是否退款分交易 = True
                objItem.是否预交 = True
                objItem.是否密文 = Val(objCard.卡号密文规则) <> 0
                objItem.结算性质 = objCard.结算性质
                objItem.结算金额 = roundEx(dblTemp, 2)
                objItem.剩余金额 = roundEx(Val(Nvl(!预交余额)), 2)
                objItem.原始金额 = roundEx(Val(Nvl(!金额)), 2)
                objItem.关联交易ID = Val(Nvl(!关联交易ID))
                objItem.交易流水号 = Trim(Nvl(!交易流水号))
                objItem.交易说明 = Trim(Nvl(!交易说明))
                objItem.结算号码 = Trim(Nvl(!结算号码))
                objItem.结算摘要 = Trim(Nvl(!摘要))
                objItem.门诊预交 = int预交类型 = 1
                 
                objItems_out.AddItem objItem
                objItems_out.结算金额 = roundEx(objItems_out.结算金额 + objItem.结算金额, 6)
            End If
            If dblMoney = 0 Then GetItemsFromRecord = True: Exit Function
            .MoveNext
        Loop
    End With
    GetItemsFromRecord = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadThirdDelDeposit(Optional int主页ID As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载三方退款信息
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-21 18:01:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItemsTemp As clsBalanceItems, objItem As clsBalanceItem
    Dim objFsItems As clsBalanceItems '负数预交集
    Dim rsMoney As ADODB.Recordset, strSQL As String
    Dim int预交类型 As Integer, objCard As Card
    Dim strWhere As String, lng病人id As Long, strDefaultBalance As String, strTemp As String
    Dim lngCardTypeID As Long, bln消费卡 As Boolean, blnDelCash As Boolean
    Dim lngRow As Long, dblThirdMoney As Double, dblCashMoney As Double
    Dim blnAdd As Boolean, dblMoney As Double
    Dim intKind As Integer
    Dim i As Integer
     
    On Error GoTo errHandle
    
    If mpatiInfo.病人ID = 0 Then Exit Function
    
    lng病人id = mpatiInfo.病人ID
    
    strWhere = " And Not Exists(Select 1 From 结算方式   Where B.结算方式= 名称 And 性质=5)"    '不含代收款
    If int主页ID > 0 Then strWhere = strWhere & " And b.主页ID=[3] "
    
    int预交类型 = cboType.ItemData(cboType.ListIndex)
    strSQL = "" & _
    "    Select b.no,a.预交id, a.病人id, a.预交类别, nvl(a.预交余额,0) as 预交余额,b.金额,a.预交余额 as 冲预交,b.结算方式, Nvl(b.卡类别id, b.结算卡序号) As 卡类别id, " & vbCrLf & _
    "           Decode(Nvl(b.结算卡序号, 0), 0, 0, 1) As 消费卡, b.卡号, " & vbCrLf & _
    "           b.交易流水号, b.交易说明, b.关联交易id, b.收款时间,b.结算号码,b.摘要,nvl(c.是否退现,0) as 消费卡是否退现,M.性质 " & vbCrLf & _
    "    From 预交单据余额 A, 病人预交记录 B,消费卡类别目录 C,结算方式 M " & vbCrLf & _
    "    Where a.病人id = [1] And a.预交类别 = [2] And a.预交id = b.Id and B.结算卡序号=C.编号(+) and nvl(a.预交余额,0)<>0 And B.结算方式=M.名称(+) " & strWhere & vbCrLf & _
    "    Order By b.收款时间"
    
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人id, int预交类型, int主页ID)
    
    Set rsMoney = zlDatabase.CopyNewRec(rsMoney)
    
    Set objFsItems = New clsBalanceItems
    
    '先处理负数预交处理
    rsMoney.Filter = "预交余额<0"
    rsMoney.Sort = "收款时间"
    With rsMoney
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            lngCardTypeID = Val(Nvl(!卡类别ID))
            bln消费卡 = Val(Nvl(!消费卡)) = 1
            dblMoney = roundEx(Val(Nvl(!预交余额)), 6)
            
            Set objCard = mobjThridSwap.zlGetCardFromCardType(lngCardTypeID, bln消费卡, Nvl(!结算方式))
            Set objItem = New clsBalanceItem
            Set objItem.objCard = objCard
            objItem.结算方式 = Nvl(!结算方式)
            If objItem.结算方式 = "" Then objItem.结算方式 = objCard.结算方式
            objItem.单据号 = Nvl(!NO)
            objItem.卡号 = Nvl(!卡号)
            objItem.卡类别ID = lngCardTypeID
            objItem.是否允许删除 = True
            objItem.预交ID = Val(Nvl(!预交ID))
            objItem.消费卡 = bln消费卡
            objItem.校对标志 = 1
            objItem.是否退款分交易 = True
            objItem.是否预交 = True
            objItem.是否密文 = Val(objCard.卡号密文规则) <> 0
            objItem.结算性质 = objCard.结算性质
            objItem.结算金额 = roundEx(Val(Nvl(!预交余额)), 2)
            objItem.剩余金额 = objItem.结算金额
            objItem.原始金额 = roundEx(Val(Nvl(!金额)), 2)
            objItem.关联交易ID = Val(Nvl(!关联交易ID))
            objItem.交易流水号 = Trim(Nvl(!交易流水号))
            objItem.交易说明 = Trim(Nvl(!交易说明))
            objItem.结算号码 = Trim(Nvl(!结算号码))
            objItem.结算摘要 = Trim(Nvl(!摘要))
            objItem.门诊预交 = int预交类型 = 1
            objItem.是否转帐 = objCard.是否转帐及代扣
            objFsItems.AddItem objItem
            objFsItems.结算金额 = roundEx(objFsItems.结算金额 + objItem.结算金额, 6)
            .MoveNext
        Loop
    End With
    
    '其次处理存在关联交易ID负数预交
    For Each objItem In objFsItems
         If objItem.卡类别ID > 0 And objItem.消费卡 = False Then
            '三方卡
            rsMoney.Filter = "卡类别ID=" & objItem.卡类别ID & " And 消费卡=0 And 关联交易ID=" & objItem.关联交易ID & " And 冲预交>0"
            If objItem.objTag Is Nothing Then Set objItem.objTag = New clsBalanceItems
            Set objItemsTemp = objItem.objTag
            dblMoney = roundEx(-1 * objItem.结算金额 - objItemsTemp.结算金额, 6)
            If dblMoney >= 0 Then
                Call GetItemsFromRecord(int预交类型, dblMoney, rsMoney, objItemsTemp)
                Set objItem.objTag = objItemsTemp
            End If
         End If
    Next
    
    '再处理关联交易ID存在，但不存在对应的的记录
    For Each objItem In objFsItems
         If objItem.卡类别ID > 0 And objItem.消费卡 = False Then
            '三方卡
            rsMoney.Filter = "卡类别ID=" & objItem.卡类别ID & "  And 消费卡=0 And 关联交易ID=0 And 冲预交>0"
            If objItem.objTag Is Nothing Then Set objItem.objTag = New clsBalanceItems
            Set objItemsTemp = objItem.objTag
            dblMoney = roundEx(-1 * objItem.结算金额 - objItemsTemp.结算金额, 6)
            
            If dblMoney >= 0 Then
                Call GetItemsFromRecord(int预交类型, dblMoney, rsMoney, objItemsTemp)
                Set objItem.objTag = objItemsTemp
            End If
         End If
    Next
         
    '最后处理普通的分摊数据
    For Each objItem In objFsItems
        '三方卡
        rsMoney.Filter = "冲预交>0"
        If objItem.objTag Is Nothing Then Set objItem.objTag = New clsBalanceItems
        Set objItemsTemp = objItem.objTag
        dblMoney = roundEx(-1 * objItem.结算金额 - objItemsTemp.结算金额, 6)
        If dblMoney >= 0 Then
            Call GetItemsFromRecord(int预交类型, dblMoney, rsMoney, objItemsTemp)
            Set objItem.objTag = objItemsTemp
        End If
    Next
    
    Call SaveAutoRelevanceData(lng病人id, objFsItems)
    
    Call ClearVsBalance
    lngRow = 1
    dblCashMoney = 0: dblThirdMoney = 0: strDefaultBalance = ""
    vsBlance.Redraw = flexRDNone
    
    rsMoney.Filter = "冲预交>0"
    rsMoney.Sort = "收款时间"
    With rsMoney
        Do While Not .EOF
            lngCardTypeID = Val(Nvl(!卡类别ID))
            bln消费卡 = Val(Nvl(!消费卡)) = 1
            dblMoney = roundEx(Val(Nvl(!冲预交)), 6)
    
            Set objCard = mobjThridSwap.zlGetCardFromCardType(lngCardTypeID, bln消费卡, Nvl(!结算方式))
            Set objItem = New clsBalanceItem
            Set objItem.objCard = objCard
            objItem.结算方式 = Nvl(!结算方式)
            If objItem.结算方式 = "" Then objItem.结算方式 = objCard.结算方式
            objItem.单据号 = Nvl(!NO)
            objItem.卡号 = Nvl(!卡号)
            objItem.卡类别ID = lngCardTypeID
            objItem.是否允许删除 = True
            objItem.预交ID = Val(Nvl(!预交ID))
            objItem.消费卡 = bln消费卡
            objItem.校对标志 = 1
            objItem.是否退款分交易 = True
            objItem.是否预交 = True
            objItem.是否密文 = Val(objCard.卡号密文规则) <> 0
            objItem.结算性质 = objCard.结算性质
            objItem.结算金额 = dblMoney
            objItem.未退金额 = dblMoney
            objItem.剩余金额 = roundEx(Val(Nvl(!预交余额)), 2)
            objItem.原始金额 = roundEx(Val(Nvl(!金额)), 2)
            objItem.关联交易ID = Val(Nvl(!关联交易ID))
            objItem.交易流水号 = Trim(Nvl(!交易流水号))
            objItem.交易说明 = Trim(Nvl(!交易说明))
            objItem.结算号码 = Trim(Nvl(!结算号码))
            objItem.结算摘要 = Trim(Nvl(!摘要))
            objItem.门诊预交 = int预交类型 = 1
            objItem.是否转帐 = objCard.是否转帐及代扣
            If objCard Is Nothing Then
                objItem.结算性质 = Val(Nvl(!性质))
            Else
                objItem.结算性质 = objCard.结算性质
            End If
        
            Set objItemsTemp = New clsBalanceItems
            objItemsTemp.AddItem objItem
            objItemsTemp.结算金额 = objItem.结算金额
            objItemsTemp.收费类型 = 1
            blnAdd = False
            
            If bln消费卡 Then
                objItem.是否允许退现 = Val(Nvl(!消费卡是否退现)) = 1
                objItem.是否强制退现 = True
                objItem.是否允许删除 = True
                objItem.Tag = IIf(objItem.是否允许退现, "缺省退现", "")
                blnDelCash = IIf(objItem.是否允许退现, True, False)
            ElseIf lngCardTypeID > 0 Then
                If Not mobjThridSwap.zlThirdReturnCashCheck(objCard, objItemsTemp, blnDelCash, strTemp) Then
                    '1.禁止退现
                    objItem.是否允许退现 = False
                    objItem.是否强制退现 = blnDelCash
                    objItem.是否允许删除 = True
                    blnDelCash = False
                    blnAdd = True
                Else
                    objItem.是否允许编辑 = False
                    objItem.是否允许删除 = True
                    objItem.是否强制退现 = True
                    objItem.是否允许退现 = True
                    
                    If blnDelCash = False Then  '是否缺省退现
                        '允许退现，可以删除
                        objItem.Tag = ""
                    Else
                        objItem.Tag = "缺省退现"
                        If strTemp <> "" And strDefaultBalance = "" Then strDefaultBalance = strTemp
                    End If
                End If
            End If
            
            If lngCardTypeID <> 0 Then
                objItem.结算类型 = IIf(bln消费卡, 5, 3)
            ElseIf objCard.结算性质 = 7 Then
                objItem.结算类型 = 4 '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                blnDelCash = IIf(objItem.是否允许退现, True, False)
            Else
                objItem.结算类型 = 0
                objItem.是否允许删除 = True
                objItem.是否强制退现 = True
                objItem.是否允许退现 = True
                blnDelCash = True
            End If
            
            With vsBlance
                .TextMatrix(lngRow, .ColIndex("退现")) = IIf(objItem.是否允许退现 And blnDelCash, 1, 0)
                .TextMatrix(lngRow, .ColIndex("类型")) = objItem.结算类型
                .TextMatrix(lngRow, .ColIndex("卡类别ID")) = objItem.卡类别ID
                .TextMatrix(lngRow, .ColIndex("消费卡ID")) = objItem.消费卡ID
                .TextMatrix(lngRow, .ColIndex("结算性质")) = objItem.结算性质
                .TextMatrix(lngRow, .ColIndex("编辑状态")) = IIf(objItem.是否允许编辑, "1", "0") & "|" & IIf(objItem.是否允许删除, "1", "0")      '是否允许编辑|是否允许删除
                .TextMatrix(lngRow, .ColIndex("是否退现")) = IIf(objCard.是否退现, 1, 0)
                .TextMatrix(lngRow, .ColIndex("是否全退")) = IIf(objCard.是否全退, 1, 0)
                .TextMatrix(lngRow, .ColIndex("校对标志")) = objItem.校对标志
                .TextMatrix(lngRow, .ColIndex("是否密文")) = IIf(objItem.是否密文, 1, 0)
                .TextMatrix(lngRow, .ColIndex("卡类别名称")) = objCard.名称
                .TextMatrix(lngRow, .ColIndex("单据号")) = objItem.单据号
                .TextMatrix(lngRow, .ColIndex("退款方式")) = objItem.结算方式
                .TextMatrix(lngRow, .ColIndex("预交余额")) = IIf(objItem.结算性质 = 9, Format(objItem.剩余金额, "###0.00#####"), Format(objItem.剩余金额, "0.00"))
                .TextMatrix(lngRow, .ColIndex("退款金额")) = IIf(objItem.结算性质 = 9, Format(objItem.结算金额, "###0.00#####"), Format(objItem.结算金额, "0.00"))
                .TextMatrix(lngRow, .ColIndex("结算号码")) = objItem.结算号码
                .TextMatrix(lngRow, .ColIndex("备注")) = objItem.结算摘要
                .TextMatrix(lngRow, .ColIndex("交易流水号")) = objItem.交易流水号
                .TextMatrix(lngRow, .ColIndex("交易说明")) = objItem.交易说明
                .TextMatrix(lngRow, .ColIndex("卡号")) = IIf(objItem.是否密文, String(Len(objItem.卡号), "*"), objItem.卡号)
                .RowData(lngRow) = objItem
                .Rows = .Rows + 1
                lngRow = .Rows - 1
            End With
            dblThirdMoney = roundEx(dblThirdMoney + IIf(objItem.是否允许退现 And blnDelCash, 0, objItem.结算金额), 2)
            dblCashMoney = roundEx(dblCashMoney + IIf(objItem.是否允许退现 And blnDelCash, objItem.结算金额, 0), 2)
            .MoveNext
        Loop

    End With
    With vsBlance
        If .Rows > 2 Then
            If .TextMatrix(.Rows - 1, .ColIndex("退款方式")) = "" Then
                .Rows = .Rows - 1
            End If
        ElseIf .Rows <= 1 Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
        End If
    End With
    Call LoadThirdTotal
    vsBlance.Redraw = flexRDBuffered
    
    txtThirdTotal.Text = Format(dblThirdMoney, "#,##0.00")
    lblCashTotal.Tag = dblCashMoney
    txtCashTotal.Text = Format(dblCashMoney, "#,##0.00")
    txtMoney.Text = Format(dblCashMoney, "#,##0.00")
    txtTotal.Text = Format(dblCashMoney + dblThirdMoney, "#,##0.00")
    If mdbl预交余额 <> mdbl剩余款额 Then Call AutoShareBalanceMoney(mdbl剩余款额, True)
    '缺省定位到当前缺省的结算方式上（第一个)
    For i = 0 To cboStyle.ListCount - 1
        intKind = cboStyle.ItemData(i)
        If mobjThridSwap.objPayCards(intKind).结算方式 = strDefaultBalance Then
            cboStyle.ListIndex = i: Exit For
        End If
    Next
    
    LoadThirdDelDeposit = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    vsBlance.Redraw = flexRDBuffered
End Function

Private Sub SaveAutoRelevanceData(ByVal lng病人id As Long, ByVal objItems As clsBalanceItems)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存负数预交自动关联消费项目
    '入参:objItems-项目信息
     '编制:刘兴洪
    '日期:2018-09-11 17:40:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strHead As String, strTemp As String, str结算信息 As String
    Dim objItem As clsBalanceItem, objItemsTemp As clsBalanceItems, objItemTemp As clsBalanceItem
    Dim cllPro As Collection, blnTrans As Boolean
    Dim strDate As String, lng结帐ID As Long
    
    On Error GoTo errHandle
    
    If objItems Is Nothing Then Exit Sub
    If objItems.Count = 0 Then Exit Sub
    Set cllPro = New Collection
    strDate = "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')"
    
    For Each objItem In objItems
        Set objItemsTemp = objItem.objTag
        If Not objItemsTemp Is Nothing Then
            '    Zl_病人预交记录_Relevance
            strHead = "Zl_病人预交记录_Relevance("
            '    病人id_In     病人预交记录.病人id%Type,
            strHead = strHead & "" & lng病人id & ","
            '    预交id_In     病人预交记录.Id%Type,
            strHead = strHead & "" & objItem.预交ID & ","
            str结算信息 = ""
            For Each objItemTemp In objItemsTemp
                '原预交ID|金额||....
                strTemp = "||" & objItemTemp.预交ID & "|" & objItemTemp.结算金额
                If zlCommFun.ActualLen(str结算信息 & strTemp) > 4000 Then
                    lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
                    str结算信息 = Mid(str结算信息, 3)
                    strSQL = strHead
                    '    结算信息_In   Varchar2 := Null,
                    strSQL = strSQL & "'" & str结算信息 & "',"
                    '   结帐id_In     病人预交记录.结帐id%Type,
                    strSQL = strSQL & "" & lng结帐ID & ","
                    '    操作员编号_In 病人预交记录.操作员编号%Type,
                    strSQL = strSQL & "'" & UserInfo.编号 & "',"
                    '    操作员姓名_In 病人预交记录.操作员姓名%Type,
                    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                    '    收款时间_In   病人预交记录.收款时间%Type := Null,
                    strSQL = strSQL & "" & strDate & ","
                    '    校对标志_In   病人预交记录.校对标志%Type := 0,
                    strSQL = strSQL & "" & 0 & ")"
                    '    缴款组id_In   病人预交记录.缴款组id%Type := -1
                    zlAddArray cllPro, strSQL
                    str结算信息 = ""
                End If
                str结算信息 = str结算信息 & strTemp
            Next
            If str结算信息 <> "" Then
                lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
                str结算信息 = Mid(str结算信息, 3)
                strSQL = strHead
                '    结算信息_In   Varchar2 := Null,
                strSQL = strSQL & "'" & str结算信息 & "',"
                '   结帐id_In     病人预交记录.结帐id%Type,
                strSQL = strSQL & "" & lng结帐ID & ","
                '    操作员编号_In 病人预交记录.操作员编号%Type,
                strSQL = strSQL & "'" & UserInfo.编号 & "',"
                '    操作员姓名_In 病人预交记录.操作员姓名%Type,
                strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                '    收款时间_In   病人预交记录.收款时间%Type := Null,
                strSQL = strSQL & "" & strDate & ","
                '    校对标志_In   病人预交记录.校对标志%Type := 0,
                strSQL = strSQL & "" & 0 & ")"
                '    缴款组id_In   病人预交记录.缴款组id%Type := -1
                zlAddArray cllPro, strSQL
            End If
        End If
    Next
    blnTrans = True:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ReCalePtBalanceMoney(Optional intCalceType As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算普通退款金额
    '入参:intCalceType ：=0表示只计算退现合计及三方退款合计
    '                               =1表示:只处理退现部分：将剩余款作为本次退款
    '                               =2表示:所有三方结算，剩余款作为本次退款
    '编制:刘兴洪
    '日期:2018-09-07 09:42:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblCashMoney As Double, dblThirdDelMoney As Double
    Dim objItem As clsBalanceItem
    
    On Error GoTo errHandle

    With vsBlance
        dblCashMoney = 0: dblThirdDelMoney = 0
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("退款方式")) <> "" Then
                If zlGetBalanceItemFromBalanceGrid(i, objItem) Then
                   If GetVsGridBoolColVal(vsBlance, i, .ColIndex("退现")) Then
                        If intCalceType = 1 Or intCalceType = 2 Then
                            objItem.结算金额 = objItem.剩余金额
                            .TextMatrix(i, .ColIndex("退款金额")) = Format(objItem.结算金额, "###0.00" & IIf(objItem.结算性质 = 9, "###", ""))
                        End If
                        dblCashMoney = roundEx(dblCashMoney + objItem.结算金额, 6)
                   Else
                        If intCalceType = 2 Then
                            objItem.结算金额 = objItem.剩余金额
                            .TextMatrix(i, .ColIndex("退款金额")) = Format(objItem.结算金额, "###0.00" & IIf(objItem.结算性质 = 9, "###", ""))
                        End If
                        dblThirdDelMoney = roundEx(dblThirdDelMoney + objItem.结算金额, 6)
                   End If
                End If
            End If
        Next
    End With
    
    Call LoadThirdTotal
    
    txtThirdTotal.Text = Format(dblThirdDelMoney, "#,##0.00")
    txtCashTotal.Text = Format(dblCashMoney, "#,##0.00")
    lblCashTotal.Tag = dblCashMoney
    dblCashMoney = roundEx(dblCashMoney, 6)
    txtMoney.Text = Format(dblCashMoney, "#,##0.00")
    txtTotal.Text = Format(dblCashMoney + dblThirdDelMoney, "#,##0.00")
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub ClearVsBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除结算方式列表(三方)
    '编制:刘兴洪
    '日期:2018-08-30 13:36:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    
    With vsBlance
        For i = 1 To .Rows - 1
            .RowData(i) = ""
        Next
        .Rows = 2
        .Clear 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LockScreen(ByVal blnLocked As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:锁屏操作，防止操作员重复操作
    '入参:blnLocked-表示锁定
    '编制:刘兴洪
    '日期:2018-08-31 09:59:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEnabled As Boolean
    On Error GoTo errHandle
    
    blnEnabled = Not blnLocked
    cmdOK.Enabled = blnEnabled
    cmdCancel.Enabled = blnEnabled
    cmdHelp.Enabled = blnEnabled
    cmdSetup.Enabled = blnEnabled
    cmdVoucherSet.Enabled = blnEnabled
    picFace.Enabled = blnEnabled
    picInfo.Enabled = blnEnabled
    txtFact.Enabled = blnEnabled
    cboNO.Enabled = blnEnabled
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdOK_Click()
 
    Dim objPati As clsPatiInfo, lng预交ID As Long
    Dim objSefocus As Object
    Dim objDelCashThird As clsBalanceItems
    Dim bytP As Byte, blnVocherPrint As Boolean
    Dim strNos As String, int预交类型 As Integer

    If cmdOK.Enabled = False Then Exit Sub '防止重复执行
    
     
    Call LockScreen(True)
    
    If Not Check未入科不交预交 Then Call LockScreen(False):  Exit Sub
    
    If CheckDataValied(objSefocus) = False Then
        Call LockScreen(False):
        If Not objSefocus Is Nothing Then Call zlControl.ControlSetFocus(objSefocus)
        Exit Sub
    End If
    
    If Check退款 = False Then
        Call LockScreen(False): Call zlControl.ControlSetFocus(txtMoney): Exit Sub
    End If
    
    If GetPatiObject(objPati) = False Then
        Call LockScreen(False): Call zlControl.ControlSetFocus(txtPatient): Exit Sub
    End If
    
    If cboType.ListIndex <> -1 Then int预交类型 = cboType.ItemData(cboType.ListIndex)
    '作废电子票据
    If zlCancelEInvoiceBat(mpatiInfo, strNos) = False Then
        MsgBox "作废电子票据失败，禁止余额退款", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    bytP = Val(zlDatabase.GetPara("凭条打印方式", glngSys, mlngModul))
    Select Case bytP
    Case 0 '不打印预交发票
       blnVocherPrint = False
    Case 1 '自动打印
       blnVocherPrint = True
    Case 2 '打印提醒
        If MsgBox("是否需要打印预交凭条？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then blnVocherPrint = True
    End Select
    '暂时分单据进行冲预交
    'str收款时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    'str结帐ID = zlDatabase.GetNextId("病人结帐记录")
    'str结算序号 = "-" & str结帐ID
    
    '1.再保存三方卡及预交款的原样退
    Set objDelCashThird = New clsBalanceItems
    If Excute_BalanceList_ReturnMoney(objPati, lng预交ID, objDelCashThird, blnVocherPrint) = False Then
        Call LockScreen(False): Call zlControl.ControlSetFocus(vsBlance): Exit Sub
    End If
    
    '2.先保存普通的余额退款
    If Excute_CashAndOther_ReturnMoney(objPati, objDelCashThird, lng预交ID, blnVocherPrint) = False Then
        Call LockScreen(False): Call zlControl.ControlSetFocus(txtMoney): Exit Sub
    End If
    
    '重新开具电子票据
    Call zlCreateEInvoiceBat(mpatiInfo, strNos)
    
    '3.完成后，按评价器
    Call Excute_Plug_PatiPrePayAfter(objPati, lng预交ID)
    
    '4.解锁
    Call LockScreen(False)
    '问题:48249
    If mbytCallObject <> 0 Then '其他模块调用预交缴款时,直接退出
        mblnOK = True: txtPatient.Tag = "": Unload Me: Exit Sub
    End If
    
    If mblnClearWinInfor Then
        Call ClearBill
        Call InitFace(True)
        Call cboStyle_Click
    Else
        SetMoneyInfo False
        Set mpatiInfo = New clsPatientInfo
        Call GetFact  '重新获取发票号
        txtPatient.Tag = ""
    End If
    
    Call SetcmdOkEnabled
    If txtPatient.Enabled Then txtPatient.SetFocus
    mblnOK = True
End Sub

Private Function DelDepositErrBill(ByVal strNO As String, Optional ByVal bytOpt As Byte) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:删除预交异常单据记录
    '入参: strno-单据号，Optype-(0-删除异常充值单据，1-删除异常退款单据，2-删除异常余额退款单据)
    '编制:
    '日期:2018-06-29
    '说明:
    '-------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, blnTrans As Boolean
    
    On Error GoTo errHandle
    Set cllPro = New Collection
    If mobjThridSwap.zlGetDeleteSQL(strNO, bytOpt, cllPro) = False Then Exit Function
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption
    DelDepositErrBill = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    Call ErrCenter
End Function

Private Sub ClearBill()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除相关界面和数据
    '编制:刘兴洪
    '日期:2018-11-29 10:03:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    If gblnLED Then zl9LedVoice.DisplayPatient ""
    
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
    
    txtMoney.Text = "0.00": txt收款.Text = "0.00"
    lblCashTotal.Tag = "": txtCashTotal.Text = "0.00"
    txtTotal.Text = "0.00": txtThirdTotal.Text = "0.00"
    
    If cboStyle.ListCount <> 0 And cboStyle.Tag <> "" Then cboStyle.ListIndex = Val(cboStyle.Tag) '恢复缺省结算方式
    txtCode.Text = "": txtCode.Locked = False
    
    cboNote.Text = ""
    
    Call ClearVsBalance
    '医保改动
    Call Clear个人帐户
    
    '新的一张预交款单据
    cboNO.Text = "": cboNO.Locked = True
    
    vsBlance.Rows = 1: vsBlance.Rows = 2
    vsDepositHistory.Rows = 1: vsDepositHistory.Rows = 2
    vsThirdTotal.Rows = 1: vsThirdTotal.Rows = 2
    
    txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "修改票据号") And gblnBill预交 '89302
    Call GetFact
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSetup_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub
 
 
Private Sub Form_Activate()
    If mblnUnLoad Then Unload Me: Exit Sub
    
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    Call vsBlance_GotFocus
    
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
    zlControl.ControlSetFocus txtPatient
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
        Case vbKeyEscape
            Call cmdCancel_Click
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub GetFact(Optional blnFirst As Boolean = False, Optional ByVal int险类 As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取不同类别的发票
    '编制:刘兴洪
    '日期:2011-07-19 17:47:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPati As Collection
    Dim strFactNO As String, int预交类型 As Integer
    
    '票据领用检查及初始
    '电子票据处理
    If mobjEInvoice Is Nothing Then Exit Sub
    txtFact.Text = ""
    If cboType.ListIndex <> -1 Then int预交类型 = cboType.ItemData(cboType.ListIndex)
    If mobjEInvoice.zlIsStartEInvoice(int险类, int预交类型) Then
        If blnFirst Then Exit Sub
        If mobjEInvoice.zlGetTranPaperInvoiceModule = 0 Then Exit Sub
        If mobjEInvoice.zlIsHisManagerInvoice = False Then
            Call mobjEInvoice.zlGetPatiCollectFromPatiObject(mpatiInfo, cllPati)
            Call mobjEInvoice.zlGetNextInvoiceNo(Me, strFactNO, cllPati, mlng领用ID)
            If strFactNO <> "" Then txtFact.Text = strFactNO
            Exit Sub
        End If
    End If
    
    If mFactProperty.intInvoicePrint = 0 Then Exit Sub
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
        txtFact.Text = GetNextBill(mlng领用ID)
    Else
        '松散：取下一个号码
        txtFact.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("当前预交票据号", glngSys, mlngFactModule, "")))
    End If
End Sub

Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化模块参数
    '编制:刘兴洪
    '日期:2012-02-27 11:23:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
        
    On Error GoTo errHandle
    
    mstr缺省结算方式 = zlDatabase.GetPara("缺省预交结算方式", glngSys, mlngModul)
    mbytBackMoneyType = Val(zlDatabase.GetPara("退款禁止方式", glngSys, mlngModul))
    '结算方式:金额|结算方式:金额....
    mblnClearWinInfor = IIf(zlDatabase.GetPara("缴预交后不清除信息", glngSys, glngModul) <> "1", True, False)
    mbln未入科不交预交 = zlDatabase.GetPara("病人未入科不准收预交", glngSys, mlngModul, , , InStr(mstrPrivs, ";参数设置;") > 0) = "1"
    gblnSeekName = Nvl(zlDatabase.GetPara("姓名模糊查找", glngSys, mlngModul, 1)) = 1
    mbln住院退预交验证 = zlDatabase.GetPara("住院退预交验证", glngSys, mlngModul, "0") = "1"
    mbln允许在院病人余额退款 = zlDatabase.GetPara("允许在院病人余额退款", glngSys, mlngModul, "1") = "1"
    '刷卡要求输入密码
    mblnCheckPass = Mid(zlDatabase.GetPara(46, glngSys, , "0000000000"), 8, 1) = "1"
    mbln排除未缴及未审 = zlDatabase.GetPara("剩余款排除未缴及未审金额", glngSys, mlngModul, "0") = "1"

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub Form_Load()
    mblnFirst = True
    mintPrintType = -1
    Call InitPara
    mblnOK = False: mblnUnLoad = False
    
    '票据领用检查及初始
    mblnStartFactUseType = zlStartFactUseType(2)
    If mblnStartFactUseType = False Then
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact(True)
    End If
    
    
    zlControl.PicShowFlat picInfo, -1
    zlControl.PicShowFlat picFace, -1
    
    Set mpatiInfo = New clsPatientInfo
    If gOneCardData.zlGetYLCardObjs(mobjCards) = False Then Unload Me: Exit Sub

    If Not InitUnit Then Unload Me: Exit Sub
   
    Call InitIDKind
    Call InitFace
    If mblnUnLoad Then Exit Sub
    
    Call InitTab
    Call InitPanel
    
    lblTitle.Caption = gstrUnitName & "余额退款"
    mstrCardPrivs = ";" & GetPrivFunc(glngSys, 1151) & ";"
    
    If gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.编号 & "号为您服务", mlngModul, gcnOracle
    End If
    
    Call zlCheckFactIsEnough
    
    IDKind.IDKind = Val(zlDatabase.GetPara("上次输入方式", glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0))
    
    '81693:李南春,2015/4/21,评价器
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear: On Error GoTo 0
    End If
   
    Call zlInitBalanceGrid
    Call RestoreWinState(Me, App.ProductName)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    If txtPatient.Tag <> "" Then
        If MsgBox("你当前正在进行余额退款，你是否真的要退出?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    mblnUnLoad = False
    mlng领用ID = 0: mstr个人帐户 = ""
    mstr退款操作员 = "": mblnOptErrBill = False
    
    If gblnLED Then
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
    Call SaveWinState(Me, App.ProductName)
    zlDatabase.SetPara "上次输入方式", IDKind.IDKind, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    Set mobjThridSwap = Nothing
    Set mobjCards = Nothing
    Set mpatiInfo = Nothing
End Sub

Private Sub InitPrepayType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化预交类型
    '编制:刘兴洪
    '日期:2011-07-14 18:50:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    With cboType
        .Clear
        cboType.Tag = "-1"
        mblnNotClick = True
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
        mblnNotClick = False
     End With
End Sub


Private Sub InitFace(Optional blnSave As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据入口参数设置窗体界面及控制状态
    '编制:刘兴洪
    '日期:2011-07-17 10:36:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    If Not gobjSquare.objSquareCard Is Nothing And blnSave = False Then
        IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
    End If
    
    On Error GoTo errHandle
    
    strSQL = "Select 编码, 名称, 简码, 缺省标志 From 常用预交摘要 Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboNote.Clear
    If rsTmp.RecordCount > 0 Then
        While Not rsTmp.EOF
            cboNote.AddItem Nvl(rsTmp!名称)
            rsTmp.MoveNext
        Wend
    End If
    
    cboNote.ListIndex = -1: Call InitPrepayType
    If mblnUnLoad Then Exit Sub
    
    IDKind.Enabled = True
    
    '创建卡部件
    Call CreateIDAndICCardObject
    cboNO.Text = ""
    
    Call Load支付方式
    
    lblMoney.Caption = "退款金额": lblMoney.FontBold = True: lblMoney.ForeColor = vbRed
    txtMoney.ForeColor = vbRed: txtMoney.Font.Bold = True
    
    txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "修改票据号") And gblnBill预交 '89302
     
    If lbl帐户余额.Visible = False Then lbl预交余额.Left = lbl帐户余额.Left
    
    If lbl帐户余额.Visible Then
        Line2(14).Visible = True: Line2(11).x2 = 2415
    Else
        Line2(14).Visible = False: Line2(11).x2 = Line2(14).x2
    End If
    
    Call mobjThridSwap.zlInitCompents(Me, mlngModul, mobjICCard)
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

Public Sub CreateIDAndICCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建IC和ID对象
    '编制:刘兴洪
    '日期:2018-08-30 14:39:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hwnd)
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hwnd)
    Set mobjICCard.gcnOracle = gcnOracle
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub picDeposit_Resize()
    Err = 0: On Error Resume Next
    With picDeposit
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Height = .ScaleHeight
        tbPage.Width = .ScaleWidth
    End With
    zlControl.PicShowFlat picFace, -1
End Sub

Private Sub picDepositBack_Resize()
    Err = 0: On Error Resume Next
    With picDepositBack
        vsBlance.Left = .ScaleLeft
        vsBlance.Top = .ScaleTop
        vsBlance.Height = .ScaleHeight - cmdDefault.Height - 100
        vsBlance.Width = .ScaleWidth
        cmdDefault.Left = .ScaleWidth - cmdDefault.Width - 100
        cmdDefault.Top = .ScaleHeight - cmdDefault.Height - 50
    End With
End Sub

Private Sub picDepositHistory_Resize()
    Err = 0: On Error Resume Next
    With picDepositHistory
        vsDepositHistory.Left = .ScaleLeft
        vsDepositHistory.Top = .ScaleTop
        vsDepositHistory.Height = .ScaleHeight
        vsDepositHistory.Width = .ScaleWidth
    End With
End Sub

Private Sub picFace_Resize()
    Err = 0: On Error Resume Next
    With picFace
        picDeposit.Height = .ScaleHeight - picDeposit.Top - 100
        picBalance.Top = .ScaleHeight - picBalance.Height - 100
    End With
    
    With vsThirdTotal
        .Height = picBalance.Top - .Top - 100
        .ColWidth(.ColIndex("退款金额")) = IIf(.Rows * .RowHeight(0) <= .Height, 1855, 1620)
    End With
End Sub

Private Sub picInfo_Resize()
    zlControl.PicShowFlat picInfo, -1
    zlControl.PicShowFlat picFace, -1
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim objTemp As Object
    With tbPage
        Select Case Val(.Selected.Tag)
            Case pg_Page.pg_预交余额退款
                Set objTemp = picDepositBack
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            Case pg_Page.pg_预交历史记录
                Set objTemp = picDepositHistory
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
        End Select
    End With
End Sub

Private Sub txtCashTotal_Change()
    If IsNumeric(StrToNum(txtCashTotal.Text)) Then
        txtCashTotal.ForeColor = vbRed
    End If
End Sub

Private Sub txtCashTotal_GotFocus()
    txtCashTotal.SelStart = 0: txtCashTotal.SelLength = Len(txtCashTotal.Text)
End Sub

Private Sub txtCashTotal_KeyPress(KeyAscii As Integer)
    '问题27363
    If KeyAscii = 13 Then
        If txtCashTotal.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    '退款时不允许输入负数
    If KeyAscii = Asc(".") And InStr(txtCashTotal.Text, ".") > 0 Then KeyAscii = 0: Beep: Exit Sub
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
   
    If (txtCashTotal.Text <> "" And txtCashTotal.SelLength <> Len(Format(StrToNum(txtCashTotal.Text), "##,##0.00;-##,##0.00; ;"))) And _
        (Len(Format(StrToNum(txtCashTotal.Text), "##,##0.00;-##,##0.00; ;")) >= txtCashTotal.MaxLength) And _
        InStr(Chr(8), Chr(KeyAscii)) = 0 Then
        
        If txtCashTotal.SelLength > 0 And txtCashTotal.SelLength <= txtCashTotal.MaxLength Then
        Else
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
End Sub

Private Sub txtCashTotal_LostFocus()
    Dim dblMoney As Double
    If mpatiInfo.病人ID = 0 Then Exit Sub
    dblMoney = StrToNum(txtCashTotal)
    
    If Val(lblCashTotal.Tag) <> dblMoney Then
        If MsgBox("你当前输入的退款金额与退款列表中的退现金额不一致,是否自动计算退现金额？" & vbCrLf & vbCrLf & _
               "退现合计:" & Format(Val(lblCashTotal.Tag), "###0.00###") & vbCrLf & _
               "输入金额:" & Format(dblMoney, "###0.00###") & vbCrLf & _
               "", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Call zlControl.ControlSetFocus(txtCashTotal): Exit Sub
        End If
        '自动分摊
        Call AutoShareBalanceMoney(dblMoney)
    End If
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

Private Sub txtMoney_Change()
    Dim dbl退款金额 As Double, dbl差额 As Double
    dbl退款金额 = Val(lblCashTotal.Tag)
    dbl差额 = dbl退款金额 - Val(txtMoney.Text)
    If dbl差额 < 0 Then
        txt收款.Text = Format(-1 * dbl差额, "0.00")
    Else
        txt收款.Text = "0.00"
    End If
End Sub

Private Sub txtMoney_GotFocus()
    txtMoney.SelStart = 0: txtMoney.SelLength = Len(txtMoney.Text)
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    '问题27363
    If KeyAscii = 13 Then
        If txtMoney.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    '退款时不允许输入负数
    If KeyAscii = Asc(".") And InStr(txtMoney.Text, ".") > 0 Then KeyAscii = 0: Beep: Exit Sub
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub

    If (txtMoney.Text <> "" And txtMoney.SelLength <> Len(Format(StrToNum(txtMoney.Text), "##,##0.00;-##,##0.00; ;"))) And _
        (Len(Format(StrToNum(txtMoney.Text), "##,##0.00;-##,##0.00; ;")) >= txtMoney.MaxLength) And _
        InStr(Chr(8), Chr(KeyAscii)) = 0 Then

        If txtMoney.SelLength > 0 And txtMoney.SelLength <= txtMoney.MaxLength Then
        Else
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
 
End Sub

Private Sub txtMoney_LostFocus()
    '问题27363
    Dim dblMoney  As Double
    If Not IsNumeric(StrToNum(txtMoney.Text)) Then txtMoney.SetFocus: Exit Sub
    If mpatiInfo.病人ID = 0 Or IsNumeric(StrToNum(txtMoney.Text)) = False Then Exit Sub
    txtMoney.Text = Format(StrToNum(txtMoney.Text), "##,##0.00;-##,##0.00; ;")
    If txtMoney.MaxLength > 12 Then txtMoney.MaxLength = 12
    '108813:李南春,2017/5/8,语音播报控制
    If gblnLED Then
        '#22 1234.56   --预收一千二百三十四点五六元 Y
        '#23 1234.56   --找零一千二百三十四点五六元 Z
        dblMoney = StrToNum(txtMoney.Text)
        dblMoney = -1 * dblMoney
        zl9LedVoice.Speak "#22 " & dblMoney
    End If
End Sub

Private Sub cboNO_GotFocus()
    If Not cboNO.Locked Then cboNO.SelStart = 0: cboNO.SelLength = Len(cboNO.Text)
End Sub
Private Sub cboNote_GotFocus()
    cboNote.SelStart = 0: cboNote.SelLength = Len(cboNote.Text)
End Sub

Private Sub cboNote_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtMoney_Validate(Cancel As Boolean)
    Dim dblMoney As Double
    If mpatiInfo.病人ID = 0 Then Exit Sub
    dblMoney = StrToNum(txtMoney.Text)
    If Val(lblCashTotal.Tag) > dblMoney Then
        MsgBox "输入退款金额（" & Format(dblMoney, "###0.00###") & "）小于本次应退金额（" & Format(Val(lblCashTotal.Tag), "###0.00###") & _
                     "），将调整退款金额为" & Format(Val(lblCashTotal.Tag), "###0.00###") & "元！" & vbCrLf & "", vbInformation, gstrSysName
        txtMoney.Text = Format(lblCashTotal.Tag, "#,##0.00")
        zlControl.TxtSelAll txtMoney
        Cancel = True: Exit Sub
    End If
End Sub

Private Sub txtPatient_Change()
    If Not Me.ActiveControl Is txtPatient Or txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    txtPatient.Tag = ""
End Sub

Private Sub txtPatient_GotFocus()
    txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjIDCard.SetEnabled(True)
    If Not mobjICCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjICCard.SetEnabled(True)
    
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    Dim blnSel As Boolean
    
    If txtPatient.Locked Then Exit Sub
        
        
    If txtPatient.Tag <> "" And KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
        
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
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        Set frmPatiSelect.mfrmParent = Me
        frmPatiSelect.mbytSize = 1 '大字体(小四)
        frmPatiSelect.Show 1, Me
        blnSel = True
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
        If blnSel Then zlCommFun.PressKey vbKeyTab
    End If
    
End Sub
Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnCancel As Boolean, bytPrepayType As Byte
    
    Call ClearBill
    '读取病人信息
    SetMoneyInfo True
    sta.Panels(2) = ""
    If objCard.名称 Like "IC卡*" And objCard.系统 Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If Not GetPatient(objCard, strInput, blnCancel, blnCard) Then
        '处理异常单据跳过
        If mblnOptErrBill = False Then
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
            Set mpatiInfo = New clsPatientInfo
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        End If
        Exit Sub
    End If
    
    '设置病人费用信息
    Call SetMoneyInfo(False, mpatiInfo.病人ID)
    Call LoadPatiPage(mpatiInfo.病人ID)
    
    '79361:李南春,2014/11/18,缺省病人的预交类型
    bytPrepayType = IIf(mpatiInfo.在院, 2, 1)
    If bytPrepayType <> mbytPrepayType Then
        mbytPrepayType = bytPrepayType: Call InitPrepayType
    End If
    Call LoadPatiInforToContronl '加载病人信息
    
    Call Led欢迎信息
    Call SetcmdOkEnabled
    Call zlCommFun.PressKey(vbKeyTab)
    
    '重新加载当前余额退款信息
    Call LoadThirdDelDeposit
    '加载历史预交记录
    Call ShowHistoryPrepay
End Sub

Private Sub Led欢迎信息()
    Dim strInfo As String, lngPatient As Long
    'LED初始化
    If Not gblnLED Then Exit Sub
    If gblnLedWelcome Then
        zl9LedVoice.Reset com
        zl9LedVoice.Speak "#1"
        zl9LedVoice.Init UserInfo.编号 & "号为您服务", mlngModul, gcnOracle
    End If
    strInfo = Trim(txtPatient.Text)
    If mpatiInfo.病人ID > 0 Then strInfo = strInfo & " " & mpatiInfo.性别 & " " & mpatiInfo.年龄: lngPatient = mpatiInfo.病人ID
    zl9LedVoice.DisplayPatient strInfo, lngPatient

End Sub

Private Sub Clear个人帐户()
    '功能：清除个人帐户信息
    Dim i As Integer
    
    On Error GoTo errHandle
    
    For i = 0 To cboStyle.ListCount - 1
        If cboStyle.ItemData(i) = 3 Then
            cboStyle.RemoveItem i: Exit For
        End If
    Next
    mcur帐户余额 = 0
    lbl帐户余额.Caption = lbl帐户余额.Tag
    lbl帐户余额.Visible = False: Line2(14).Visible = False
    Line2(11).x2 = Line2(14).x2
    lbl预交余额.Left = lbl帐户余额.Left
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
                                           Optional ByRef blnCancel As Boolean, _
                                           Optional ByVal blnCard As Boolean, _
                                           Optional ByVal lng病人id As Long, _
                                           Optional ByVal lng主页ID As Long = -1) As Boolean
    '功能：读取病人信息
    '参数：strInput=[刷卡]|[A病人ID]|[B住院号]
    '          lng主页ID=-1表示门诊病人或查找所有住院次数;lng主页ID=0表示预入院病人;lng主页ID>0表示住院病人
    '说明：
    '     1.适用于病人预交款
    '     2.自动识别病人在院状态,读出(病人ID,主页ID,姓名,性别,年龄,住院号,床号,在院标志)
    '返回:是否读取成功,成功时mPatiInfo中包含病人信息,失败时清空mPatiInfo
    Dim lng卡类别ID As Long
    Dim strWhere As String, strPassWord As String, strErrMsg As String
    Dim blnHavePassWord As Boolean, blnIsMobileNO As Boolean
    
    blnCancel = False: mstr退款操作员 = ""

    If lng病人id > 0 Then GoTo ReadPati

    blnIsMobileNO = IDKind.IsMobileNo(strInput)
    Call Clear个人帐户 '清除个人帐户信息
    
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
                If GetPatiID("门诊号", strInput, lng病人id) = False Then GoTo NotFoundPati
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
    If OptOthersErrBill(mpatiInfo.病人ID) Then
        Exit Function
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

Private Function GetPatiObject(ByRef objPati_Out As clsPatiInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息对象
    '出参:objPati_Out-返回病人信息对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-30 15:14:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng主页ID As Long
    On Error GoTo errHandle
    If mpatiInfo.病人ID = 0 Then Exit Function
    
    Set objPati_Out = New clsPatiInfo
    With objPati_Out
        .姓名 = mpatiInfo.姓名
        .性别 = mpatiInfo.性别
        .年龄 = mpatiInfo.年龄
        .主页ID = mpatiInfo.主页ID
        .病人ID = mpatiInfo.病人ID
        .门诊号 = mpatiInfo.门诊号
        .住院号 = mpatiInfo.住院号
        .医疗付款方式 = mpatiInfo.医疗付款方式
    End With
    lng主页ID = IIf(cboType.ItemData(cboType.ListIndex) = 2, mpatiInfo.主页ID, 0)
    If cboPatiPage.Visible And cboPatiPage.ListIndex > 0 And cboType.ItemData(cboType.ListIndex) = 2 Then
        lng主页ID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    objPati_Out.主页ID = lng主页ID
        
    GetPatiObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Excute_Third_ReturnMoney(ByVal objPati As clsPatiInfo, _
    ByRef objCurItem As clsBalanceItem, ByRef objDelItem_Out As clsBalanceItem, _
    ByRef blnSave_out As Boolean, ByVal blnVocherPrint As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行列表中的三方退款
    '入参:objPati-当前病人信息
    '     objCurItem-当前退款项
    '     blnVocherPrint-是否打印预交凭条
    '出参:objdelItem_Out-当前有效的退款项
    '     blnSave_out-是否已经保存了数据
    '     str结帐ID-使用完成后，返回"",否则返回原结帐ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-31 10:30:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItemTemp As clsBalanceItem
    Dim blnChangeMoney As Boolean, cllPro As Collection
    Dim int预交类型 As Integer, bln电子票据 As Boolean, int险类 As Integer
    
    On Error GoTo errHandle
    
    blnSave_out = False
    If cboType.ListIndex <> -1 Then int预交类型 = cboType.ItemData(cboType.ListIndex)
    int险类 = IIf(objCurItem.结算性质 = 3, mpatiInfo.险类, 0)
    bln电子票据 = mobjEInvoice.zlIsStartEInvoice(int险类, int预交类型)
    objPati.险类 = int险类
    Call GetFact(False, int险类)
    objCurItem.发票号 = IIf(bln电子票据, "", txtFact.Text)
    objCurItem.领用ID = mlng领用ID
    
    If mobjThridSwap.zlThird_ReturnMoney_IsValied(objPati, objCurItem, 2, objItemTemp, False) = False Then
        Exit Function
    End If
    
    Set cllPro = New Collection
    If mobjThridSwap.zlThird_ReturnMoney(objPati, objCurItem, cllPro, objDelItem_Out, blnSave_out, False, blnChangeMoney, , int险类, bln电子票据) = False Then
        Exit Function
    End If
       
    Excute_Third_ReturnMoney = True
    '加入NO
    Call AddComboxNoFromNo(objDelItem_Out.单据号)
    '打印余额票据
    Call PrintDepostBill(objDelItem_Out.单据号, blnVocherPrint, int险类)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Excute_Square_ReturnMoney(ByVal objPati As clsPatiInfo, _
    ByRef objCurItem As clsBalanceItem, ByRef objDelItem_Out As clsBalanceItem, _
    ByRef blnSave_out As Boolean, ByVal blnVocherPrint As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行列表中的性质为2的其他结算方式退款
    '入参:objPati-当前病人信息
    '     objCurItem-当前退款项
    '     blnVocherPrint-是否打印预交凭条
    '出参:objdelItem_Out-当前有效的退款项
    '     blnSave_out-是否已经保存了数据
    '     str结帐ID-使用完成后，返回"",否则返回原结帐ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-31 10:30:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllSquare As Collection
    Dim int预交类型 As Integer, bln电子票据 As Boolean, int险类 As Integer
    On Error GoTo errHandle
    blnSave_out = False
    
    If cboType.ListIndex <> -1 Then int预交类型 = cboType.ItemData(cboType.ListIndex)
    int险类 = IIf(objCurItem.结算性质 = 3, mpatiInfo.险类, 0)
    bln电子票据 = mobjEInvoice.zlIsStartEInvoice(int险类, int预交类型)
    objPati.险类 = int险类
    Call GetFact(False, int险类)
    objCurItem.发票号 = IIf(bln电子票据, "", txtFact.Text)
    objCurItem.领用ID = mlng领用ID
    
    If Not mobjThridSwap.zlSquare_ReturnMoneySQL(objPati, cllSquare, objCurItem, , int险类, bln电子票据) Then Exit Function
    
    blnSave_out = True
    Excute_Square_ReturnMoney = True
    objCurItem.是否保存 = True
    objCurItem.是否结算 = True
    objCurItem.是否预交 = True
    objCurItem.是否允许编辑 = False
    objCurItem.是否允许删除 = False
    objCurItem.是否允许退现 = False
    
    Set objDelItem_Out = objCurItem
    Call AddComboxNoFromNo(objCurItem.单据号)
    '打印余额票据
    Call PrintDepostBill(objCurItem.单据号, blnVocherPrint, int险类)
        
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Excute_ListOther_ReturnMoney(ByVal objPati As clsPatiInfo, _
    ByRef objCurItem As clsBalanceItem, ByRef objDelItem_Out As clsBalanceItem, _
    ByRef blnSave_out As Boolean, ByVal blnVocherPrint As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行列表中的性质为2的其他结算方式退款
    '入参:objPati-当前病人信息
    '     objCurItem-当前退款项
    '     blnVocherPrint-是否打印预交凭条
    '出参:objdelItem_Out-当前有效的退款项
    '     blnSave_out-是否已经保存了数据
    '     str结帐ID-使用完成后，返回"",否则返回原结帐ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-31 10:30:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, blnTrans As Boolean
    Dim int预交类型 As Integer, bln电子票据 As Boolean, int险类 As Integer
    
    On Error GoTo errHandle
    blnSave_out = False
    
    '加入三方退现部分
    If cboType.ListIndex <> -1 Then int预交类型 = cboType.ItemData(cboType.ListIndex)
    int险类 = IIf(objCurItem.结算性质 = 3, mpatiInfo.险类, 0)
    bln电子票据 = mobjEInvoice.zlIsStartEInvoice(int险类, int预交类型)
    objPati.险类 = int险类
    Call GetFact(False, int险类)
    objCurItem.发票号 = IIf(bln电子票据, "", txtFact.Text)
    objCurItem.领用ID = mlng领用ID
    
    If mobjThridSwap.zlGetSaveSQLfromItem(objPati, objCurItem, 0, cllPro, True, , int险类, bln电子票据) = False Then Exit Function
    
    blnTrans = True
    Call zlExecuteProcedureArrAy(cllPro, Me.Caption)
    blnTrans = False
    
    
    blnSave_out = True
    objCurItem.是否保存 = True
    objCurItem.是否结算 = True
    objCurItem.是否预交 = True
    objCurItem.是否允许编辑 = False
    objCurItem.是否允许删除 = False
    objCurItem.是否允许退现 = False
    
    Set objDelItem_Out = objCurItem
    Excute_ListOther_ReturnMoney = True '加入NO
    Call AddComboxNoFromNo(objCurItem.单据号)
    '打印余额票据
    Call PrintDepostBill(objCurItem.单据号, blnVocherPrint, int险类)
    
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Excute_Third_TransferAccounts(ByVal objPati As clsPatiInfo, _
    ByRef objCurItem As clsBalanceItem, ByRef objDelItem_Out As clsBalanceItem, _
    ByRef blnSave_out As Boolean, ByVal blnVocherPrint As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行列表中的三方退款
    '入参:objPati-当前病人信息
    '     objCurItem-当前退款项
    '     blnVocherPrint-是否打印预交凭条
    '出参:objdelItem_Out-当前有效的退款项
    '     blnSave_out-是否已经保存了数据
    '     str结帐ID-使用完成后，返回"",否则返回原结帐ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-31 10:30:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String, i As Long
    Dim objDelItems As clsBalanceItems, cllPro As Collection
    Dim int预交类型 As Integer, bln电子票据 As Boolean, int险类 As Integer
    
    On Error GoTo errHandle
    
    blnSave_out = False
    Set cllPro = New Collection
    Set objDelItems = objCurItem.objTag
    
    If cboType.ListIndex <> -1 Then int预交类型 = cboType.ItemData(cboType.ListIndex)
    int险类 = IIf(objCurItem.结算性质 = 3, mpatiInfo.险类, 0)
    bln电子票据 = mobjEInvoice.zlIsStartEInvoice(int险类, int预交类型)
    objPati.险类 = int险类
    Call GetFact(False, int险类)
    objCurItem.发票号 = IIf(bln电子票据, "", txtFact.Text)
    objCurItem.领用ID = mlng领用ID
    
    If mobjThridSwap.zlThird_TransferAccounts(objPati, objCurItem, cllPro, strErrMsg, blnSave_out, False, int险类, bln电子票据) = False Then
        If blnSave_out Then
            For i = 1 To objDelItems.Count
                vsBlance.Cell(flexcpForeColor, objDelItems(i).行号, 0, objDelItems(i).行号, vsBlance.Cols - 1) = vbRed
                vsBlance.RowData(objDelItems(i).行号) = objDelItems(i)
            Next
        End If
        Exit Function
    End If
    
    Set objDelItem_Out = objCurItem
    For i = 1 To objDelItems.Count
        vsBlance.Cell(flexcpForeColor, objDelItems(i).行号, 0, objDelItems(i).行号, vsBlance.Cols - 1) = vbGrayed
        vsBlance.RowData(objDelItems(i).行号) = objDelItems(i)
    Next
    Excute_Third_TransferAccounts = True
    
    '加入NO
    Call AddComboxNoFromNo(objDelItem_Out.单据号)
    
    '打印余额票据
    Call PrintDepostBill(objDelItem_Out.单据号, blnVocherPrint, int险类)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Excute_BalanceList_ReturnMoney(ByVal objPati As clsPatiInfo, ByRef lng预交ID_out As Long, _
    ByRef objDelCashItems_out As clsBalanceItems, ByVal blnVocherPrint As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行列表中的三方退款
    '入参:objPati-当前病人信息
    '   blnVocherPrint-是否打印预交凭条
    '出参:lng预交ID_Out-最后一个预交ID
    '     objDelCashItems_out-当前结算列表中退现的项目
    '     str结帐ID-使用完成后，返回"",否则返回原结帐ID
    '返回:执行成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-30 15:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTranItems As clsBalanceItems, objItems As clsBalanceItems, objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim blnFind As Boolean
    Dim i As Long, blnSaveed As Boolean
    Dim bln退现 As Boolean
    
    On Error GoTo errHandle
    
    Set objTranItems = New clsBalanceItems
    If objDelCashItems_out Is Nothing Then Set objDelCashItems_out = New clsBalanceItems
    With vsBlance
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("退款方式")) <> "" Then
                
                Set objItem = Nothing
                Call zlGetBalanceItemFromBalanceGrid(i, objItem)
                If objItem Is Nothing Then
                    MsgBox "在第" & i & "行中的退款信息有误，请检查!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                objItem.行号 = i
                objItem.缴款单位 = Trim(txtUnit.Text)
                objItem.科室ID = IIf(cboUnit.ItemData(cboUnit.ListIndex) = 0, 0, cboUnit.ItemData(cboUnit.ListIndex))
                bln退现 = GetVsGridBoolColVal(vsBlance, i, .ColIndex("退现"))
                
                If bln退现 Then
                    '退现处理
                    objDelCashItems_out.AddItem objItem
                    objDelCashItems_out.结算金额 = roundEx(objDelCashItems_out.结算金额 + objItem.结算金额, 6)
                    
                ElseIf objItem.是否转帐 And objItem.消费卡 = False Then
                    If objItem.是否结算 = False Then
                        '先计算转帐
                        blnFind = False
                        For Each objItemTemp In objTranItems
                            If objItemTemp.卡类别ID = objItem.卡类别ID Then
                                '同一种类别的，按类别一起转
                                Set objItems = objItemTemp.objTag
                                If objItems Is Nothing Then Set objItems = New clsBalanceItems
                                objItems.AddItem objItem
                                objItems.结算金额 = roundEx(objItems.结算金额 + objItem.结算金额, 2)
                                Set objItemTemp.objTag = objItems
                                
                                objItemTemp.结算金额 = objItems.结算金额
                                blnFind = True
                                objTranItems.结算金额 = roundEx(objTranItems.结算金额 + objItem.结算金额, 2)
                                Exit For
                            End If
                        Next
                        If blnFind = False Then
                            Set objItemTemp = mobjThridSwap.zlCopyNewItemFromBalanceItem(objItem)
                            Set objItems = New clsBalanceItems
                            objItems.AddItem objItemTemp
                            objItems.结算金额 = roundEx(objItems.结算金额 + objItemTemp.结算金额, 2)
                            Set objItem.objTag = objItems
                            objTranItems.AddItem objItem
                            objTranItems.结算金额 = roundEx(objTranItems.结算金额 + objItem.结算金额, 2)
                        End If
                    End If
                ElseIf objItem.消费卡 Then '消费卡退款
                    '消费卡相关处理
                    If objItem.是否结算 = False Then
                        Set objItems = New clsBalanceItems
                        objItems.AddItem objItem
                        objItems.结算金额 = roundEx(objItems.结算金额 + objItem.结算金额, 2)
                        Set objItem.objTag = objItems
                        If Excute_Square_ReturnMoney(objPati, objItem, objItemTemp, blnSaveed, blnVocherPrint) = False Then
                            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                            If blnSaveed Then
                                objItemTemp.行号 = i
                                Call zlSetVsBalanceEditStatus(objItemTemp, True)
                                .RowData(i) = objItemTemp
                            End If
                            Exit Function
                        End If
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbGrayed
                        lng预交ID_out = objItemTemp.ID
                        objItemTemp.行号 = i
                        Call zlSetVsBalanceEditStatus(objItemTemp, True)
                    End If
                Else
                    If objItem.是否结算 = False Then
                        If objItem.卡类别ID <= 0 Then
                            '支票等不原样退
                            If Excute_ListOther_ReturnMoney(objPati, objItem, objItemTemp, blnSaveed, blnVocherPrint) = False Then
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                                If blnSaveed Then
                                    objItemTemp.行号 = i
                                    Call zlSetVsBalanceEditStatus(objItemTemp, True)
                                    .RowData(i) = objItemTemp
                                End If
                                Exit Function
                            Else
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbGrayed
                            End If
                        Else
                            If Not Excute_Third_ReturnMoney(objPati, objItem, objItemTemp, blnSaveed, blnVocherPrint) Then
                                If blnSaveed Then
                                    objItemTemp.行号 = i
                                    Call zlSetVsBalanceEditStatus(objItemTemp, True)
                                    .RowData(i) = objItemTemp
                                End If
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                                Exit Function
                            Else
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbGrayed
                            End If
                        End If
                        lng预交ID_out = objItemTemp.ID
                        objItemTemp.行号 = i
                        Call zlSetVsBalanceEditStatus(objItemTemp, True)
                    End If
                End If
            End If
        Next
    End With
    
    '执行转帐操作
    For Each objItem In objTranItems
        If Excute_Third_TransferAccounts(objPati, objItem, objItemTemp, blnSaveed, blnVocherPrint) = False Then Exit Function
        lng预交ID_out = objItemTemp.ID
        Call zlSetVsBalanceEditStatus(objItemTemp, True)
    Next
    
    Excute_BalanceList_ReturnMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Excute_CashAndOther_ReturnMoney(ByRef objPati As clsPatiInfo, _
    ByRef objDelCashThird As clsBalanceItems, Optional lng预交ID_out As Long, _
    Optional ByVal blnVocherPrint As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:余额退款(普通结算方式)
    '参数:lng预交ID_out-保存的预交ID
    '     当前退现的三方卡
    '   blnVocherPrint-是否打印预交凭条
    '出参:
    '   str结帐ID-使用完成后，返回"",否则返回原结帐ID
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 11:15:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCurItem As clsBalanceItem, objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim cllPro As Collection, objCard As Card
    Dim blnTrans As Boolean, dblMoney As Double, dblTotal As Double, dblTemp As Double
    Dim objDelCashItems As clsBalanceItems, int险类 As Integer
    Dim int预交类型 As Integer, bln电子票据 As Boolean
    
    On Error GoTo errH
    
    '先检查发票号是否合法
    If CheckFactIsValied = False Then Exit Function
    
    Set cllPro = New Collection
    Set objCurItem = New clsBalanceItem
    Set objCard = mobjThridSwap.zlGetCardFromBalanceName(cboStyle.Text)
    
    If objDelCashThird Is Nothing Then Set objDelCashThird = New clsBalanceItems
    Set objDelCashItems = New clsBalanceItems
    
    If StrToNum(txtCashTotal.Text) = 0 Then Excute_CashAndOther_ReturnMoney = True: Exit Function
    
    If cboType.ListIndex <> -1 Then int预交类型 = cboType.ItemData(cboType.ListIndex)
    int险类 = IIf(objCurItem.结算性质 = 3, mpatiInfo.险类, 0)
    Call GetFact(False, int险类)
    bln电子票据 = mobjEInvoice.zlIsStartEInvoice(int险类, int预交类型)
    objPati.险类 = int险类
    
    With objCurItem
        Set .objCard = objCard
        .结算金额 = StrToNum(txtCashTotal.Text)
        .结算类型 = 0
        .结算方式 = objCard.结算方式
        .结算摘要 = Trim(cboNote.Text)
        .结算号码 = Trim(txtCode.Text)
        .开户行 = Trim(txt开户行.Text)
        .科室ID = IIf(cboUnit.ItemData(cboUnit.ListIndex) = 0, 0, cboUnit.ItemData(cboUnit.ListIndex))
        .缴款单位 = Trim(txtUnit.Text)
        .门诊预交 = IIf(cboType.ItemData(cboType.ListIndex) = 1, True, False)
        .发票号 = IIf(bln电子票据, "", txtFact.Text)
        .领用ID = mlng领用ID
    End With

    '加入三方退现部分
    For Each objItemTemp In objDelCashThird
        objDelCashItems.AddItem objItemTemp
        objDelCashItems.结算金额 = roundEx(objDelCashItems.结算金额 + objItemTemp.结算金额, 2)
    Next
    
    If roundEx(objCurItem.结算金额, 2) - roundEx(objDelCashThird.结算金额, 2) > 0 Then
        If MsgBox("当前退款金额比病人可退余额还多(不含三方帐户结算退款），是否继续?" & vbCrLf & _
            "可退余额:" & Format(roundEx(objDelCashThird.结算金额, 2), "0.00") & vbCrLf & _
            "本次退款:" & Format(objCurItem.结算金额, "0.00") & vbCrLf & "注： 可退余额=普通结算余额+三方帐户允许退现合计 ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    Set objCurItem.objTag = objDelCashItems
    If mobjThridSwap.zlGetSaveSQLfromItem(objPati, objCurItem, 0, cllPro, True, , int险类, bln电子票据) = False Then Exit Function
    
    blnTrans = True
    Call zlExecuteProcedureArrAy(cllPro, Me.Caption)
    blnTrans = False
    
    '加入NO
    Call AddComboxNoFromNo(objCurItem.单据号)
    
    '打印余额票据
    Call PrintDepostBill(objCurItem.单据号, blnVocherPrint, int险类)
    
    lng预交ID_out = objCurItem.ID
    Excute_CashAndOther_ReturnMoney = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If Err.Description Like "*退款金额大于病人剩余预交余额*" And mbytOracleBackType = 1 Then
        If MsgBox("退款金额比病人当前的余额多,是否忽略？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
        mbytOracleBackType = 0
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlSetVsBalanceEditStatus(ByVal objItem As clsBalanceItem, Optional blnSetRowData As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置编辑状态
    '入参:blnSetRowData-是否将objItem设置给Rowdata属性
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-09-03 10:06:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    If objItem Is Nothing Then Exit Sub
    
    lngRow = objItem.行号
    With vsBlance
        If lngRow < 0 Or lngRow > .Rows - 1 Then Exit Sub
        .TextMatrix(lngRow, .ColIndex("结算状态")) = IIf(objItem.是否结算, 1, 0)
        .TextMatrix(lngRow, .ColIndex("编辑状态")) = IIf(objItem.是否允许编辑, 1, 0) & "|" & IIf(objItem.是否允许删除, 1, 0)
        If blnSetRowData Then .RowData(lngRow) = objItem
    End With
End Sub


Private Sub AddComboxNoFromNo(ByVal strDepositNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据号，将单据号加入单据号下拉框中
    '编制:刘兴洪
    '日期:2018-08-31 09:50:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, i As Long
    
    
    On Error GoTo errHandle
    '加入单据历史记录(所有类型单据)
    strNO = strDepositNo
    For i = 0 To cboNO.ListCount - 1
        strNO = strNO & "," & cboNO.List(i)
    Next
    cboNO.Clear
    For i = 0 To UBound(Split(strNO, ","))
        cboNO.AddItem Split(strNO, ",")(i)
        If i = 9 Then Exit For '只显示10个
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub PrintDepostBill(ByVal strNO As String, ByVal blnVocherPrint As Boolean, Optional ByVal int险类 As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印预交票据
    '入参:blnPrint-是否打印预交票据
    '       blnVocherPrint-是否打印预交凭条
    '       strNO-单据号
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-31 09:43:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int预交类型 As Integer
    
    On Error GoTo errHandle
    If cboType.ListIndex <> -1 Then int预交类型 = cboType.ItemData(cboType.ListIndex)
    If mobjEInvoice.zlIsStartEInvoice(int险类, int预交类型) Then Exit Sub
    
    If blnVocherPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1103_2", Me, "NO=" & strNO, 2)
    End If
    
    If mintPrintType < 0 Then
        Select Case mFactProperty.intInvoicePrint
            Case 0 '不打印预交发票
               mintPrintType = 0
            Case 1 '自动打印
               mintPrintType = 1
            Case 2 '打印提醒
                If MsgBox("是否需要打印预交票据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mintPrintType = 1
                Else
                    mintPrintType = 0
                End If
        End Select
    End If
    
    If mintPrintType = 0 Then Exit Sub

    If Not gblnBill预交 And Trim(txtFact.Text) <> "" Then
        '松散：保存当前号码
        zlDatabase.SetPara "当前预交票据号", Trim(txtFact.Text), glngSys, mlngFactModule
    End If
    
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & cboNO.List(0), "病人ID=" & mpatiInfo.病人ID, "收款时间=" & Format(Now, "yyyy-mm-dd HH:MM:SS"), _
    IIf(mFactProperty.intInvoiceFormat = 0, "", "ReportFormat=" & mFactProperty.intInvoiceFormat), 2)
    
    Call zlCheckFactIsEnough
'    Call GetFact
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub GetDepositData(ByVal lng病人id As Long, Optional ByVal int主页ID As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新读取预交数据
    '入参:lng病人ID-病人ID巧
    '编制:刘兴洪
    '日期:2011-07-22 17:02:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strWhere As String
    
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
    '功能:根据相关的结算方式和门诊类型,显示预交余额
    '入参:blnReRead-重读数据
    '       lng病人ID-读取指定的病人ID(0时,从mPatiInfo中读取病人ID)
    '编制:刘兴洪
    '日期:2011-07-21 15:44:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset, strSQL As String
    Dim int预交类型 As Integer
    Dim dbl未审 As Double, dbl未缴 As Double, dblYB As Double
    Dim lng主页ID As Long, dbl剩余款额 As Double, int主页ID As Integer
    
    On Error GoTo errHandle
    If lng病人id = 0 Then
        If mpatiInfo.病人ID = 0 Then Exit Sub
        lng病人id = mpatiInfo.病人ID
    End If
    If cboPatiPage.Visible And cboPatiPage.ListIndex >= 0 Then
        int主页ID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    If blnreReadData Then Call GetDepositData(lng病人id, int主页ID)
    If cboType.ListIndex < 0 Then Exit Sub
    
    sta.Panels(2).Text = ""
    mdbl费用余额 = 0: mdbl预交余额 = 0: mdbl剩余款额 = 0
    int预交类型 = cboType.ItemData(cboType.ListIndex)
    
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
    
    mdbl剩余款额 = Format(mdbl预交余额 - mdbl费用余额, "0.00")
    '问题27363
    lbl费用余额.Caption = lbl费用余额.Tag & Format(mdbl费用余额, "##,##0.00;-##,##0.00; ;")
    lbl预交余额.Caption = lbl预交余额.Tag & Format(mdbl预交余额, "##,##0.00;-##,##0.00; ;")
    dbl未审 = GetUnAuditedFee(lng病人id, , int预交类型)
    dbl未缴 = GetUnAuditedFee(lng病人id, False, int预交类型)
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


Private Sub SetMoneyInfo(blnClear As Boolean, Optional lng病人id As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示金额等信息
    '入参:blnClear-清除
    '     lng病人ID-指定病人ID
    '编制:刘兴洪
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
        lblWorkUnit.Caption = lblWorkUnit.Tag
        
        lbl未审费用.Caption = lbl未审费用.Tag
        lbl未缴费用.Caption = lbl未缴费用.Tag
        lbl费用余额.Caption = lbl费用余额.Tag
        lbl预交余额.Caption = lbl预交余额.Tag
        lbl剩余款额.Caption = lbl剩余款额.Tag
        lbl医保预结.Caption = lbl医保预结.Tag
        lbl手机号.Caption = lbl手机号.Tag
        lbl身份证号.Caption = lbl身份证号.Tag
        lbl应收款.Caption = lbl应收款.Tag
        lbl应收款.ForeColor = &H80000007
        
        mdbl费用余额 = 0
        mdbl预交余额 = 0
        mdbl剩余款额 = 0
    Else
        On Error GoTo errHandle
        '显示预交余额
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
    If mpatiInfo.病人ID = 0 Then
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
        Call SetPatiColor(txtPatient, mstr病人类型)
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

Private Sub txt收款_GotFocus()
    Call zlControl.TxtSelAll(txt收款)
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化门诊，住院临床科室信息
    '编制:刘兴洪
    '日期:2018-11-29 10:33:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
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

Private Sub InitIDKind()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化IDKind控件的识别项
    '编制:刘兴洪
    '日期:2018-11-29 10:36:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKind As String
    
    On Error GoTo errHandle
    
    strKind = "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;留|留观号|0;就|就诊卡|0;手|手机号|0"
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKind, txtPatient)
    mtySquareCard.blnExistsObjects = Not gobjSquare.objSquareCard Is Nothing
'    gobjSquare.objSquareCard.mblnYLMgr = mbytCallObject = 2
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load支付方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:刘兴洪
    '日期:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objPayCard As Cards, str性质 As String
    Dim objCard As Card
    
    '结算方式:费用查询和医疗卡调用时，一般只支付预交款,不存在代收的情况
    'mbytCallObject:调用的对象(0-预交款管理调用;1-病人费用查询调用;2-医疗卡管理调用;3-门诊挂号调用...
    If InStr(1, mstrPrivs, ";预交收款;") > 0 Or _
        InStr(1, mstrPrivs, ";预交收款;") > 0 Or _
        InStr(1, mstrPrivs, ";预交结清退款;") > 0 Or _
        InStr(1, mstrPrivs, ";门诊预交转住院;") > 0 _
        Or InStr(1, mstrPrivs, ";住院预交转门诊;") > 0 Or mbytCallObject > 0 Then
        str性质 = ",1,2,7,8,3"
    End If
    
    If str性质 = "" Then str性质 = ",1,2,7,8,3"
    
    str性质 = Mid(str性质, 2)
    
    If mblnNurseCall Then
        str性质 = "7,8"
    End If
    
    If mobjThridSwap.zlReGetPayCards(str性质, "预交款", objPayCard) = False Then
        MsgBox "预交场合没有可用的结算方式,请先到结算方式管理中设置。", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    If objPayCard.Count = 0 Then
        MsgBox "预交场合没有可用的结算方式,请先到结算方式管理中设置。", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    
    '余额退款，只加载普通结算方式的余额退款
    For i = 1 To objPayCard.Count
        Set objCard = objPayCard(i)
        If objCard.接口序号 <= 0 And objCard.结算性质 <> 3 Then
             cboStyle.AddItem objCard.结算方式
             cboStyle.ItemData(cboStyle.NewIndex) = objCard.结算性质
             If objCard.缺省标志 And cboStyle.ListIndex < 0 Then cboStyle.ListIndex = cboStyle.NewIndex
             If objCard.结算方式 = mstr缺省结算方式 Then cboStyle.ListIndex = cboStyle.NewIndex
        End If
        If cboStyle.ListIndex < 0 And cboStyle.ListCount <> 0 Then cboStyle.ListIndex = 0
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function zlCheckFactIsEnough() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前票据是否允足
    '编制:刘兴洪
    '日期:2012-09-06 15:41:52
    '说明:
    '问题:37372
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng剩余数量 As Long, strType As String
    
    '需要检查剩余数量是否充足:
    If cboType.ListIndex < 0 Then
        strType = ""
    Else
        strType = cboType.ItemData(cboType.ListIndex)
    End If
    
    If zlCheckInvoiceOverplusEnough(2, gint提醒剩余票据张数, lng剩余数量, mlng领用ID, strType) = False Then
        MsgBox "注意:" & vbCrLf & _
               "    当前剩余票据(" & lng剩余数量 & ") 小于了报警的张数(" & gint提醒剩余票据张数 & "),请注意更换发票!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        zlCheckFactIsEnough = False: Exit Function
    End If
    zlCheckFactIsEnough = True
End Function

Private Sub LoadPatiPage(ByVal lng病人id As Long)
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
    With cboPatiPage
        .AddItem "所有住院": .Tag = 0
        .ItemData(.NewIndex) = -1
        
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
            If mblnNurseCall Then
                If Val(Nvl(rsTemp!主页ID)) = mlng主页ID Then
                    .ListIndex = .NewIndex
                End If
                cboPatiPage.Enabled = False
            Else
                If Val(Nvl(rsTemp!主页ID)) = mpatiInfo.主页ID Then
                    .ListIndex = .NewIndex
                End If
            End If
            rsTemp.MoveNext
        Loop
        If .ListCount > 0 Then .ListIndex = 0
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
    Dim lng病人id As Long, lng主页ID As Long
    Dim str病人id As String, PatiPageInfo As clsPatientInfo
    
    On Error GoTo errHandle
    
    If mbln未入科不交预交 = False Then Check未入科不交预交 = True: Exit Function
    
    '不诊预交不检查
    If cboType.ItemData(cboType.ListIndex) <> 2 Then Check未入科不交预交 = True: Exit Function
    
    '当前住院次数不为在院的,也不检查
    If mpatiInfo.在院 = False Then Check未入科不交预交 = True: Exit Function
    
    lng病人id = mpatiInfo.病人ID
    '不存在住院次数的,也能缴预交,因此不检查
    If cboPatiPage.ListIndex < 0 Then Check未入科不交预交 = True: Exit Function
    
    lng主页ID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    str病人id = lng病人id & ":" & lng主页ID
    Call GetPatiPageInforByID(str病人id, PatiPageInfo, False)
    If PatiPageInfo.已入科 = 0 Then
        MsgBox "注意" & vbCrLf & "   病人『" & mpatiInfo.姓名 & "』未入科,不允许缴预交款!", vbInformation + vbOKOnly, gstrSysName
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
    Dim lng病人id As Long
    Dim dbl预交余额 As Double, dbl费用余额 As Double, dbl剩余余额 As Double
    Dim intIndex As Integer
    Dim objCard As Card
    On Error GoTo errHandle
    
    If mpatiInfo.病人ID = 0 Then Exit Function
    
    If cboType.ListIndex < 0 Then
        If StrToNum(txtMoney.Text) <> 0 Then
            MsgBox "未选择指定的结算信息!", vbInformation + vbOKOnly, gstrSysName
        End If
        Exit Function
    End If
    Check退款 = True
    
    Exit Function
    
    intIndex = cboType.ItemData(cboType.ListIndex)
    Set objCard = mobjThridSwap.objPayCards(intIndex)
    lng病人id = mpatiInfo.病人ID
    Set mrsDepositBalance = GetMoneyInfo(lng病人id)
    If Not mrsDepositBalance Is Nothing Then
        With mrsDepositBalance
            .Filter = "类型=" & objCard.结算性质
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                dbl费用余额 = dbl费用余额 + Val(Nvl(!费用余额))
                dbl预交余额 = dbl预交余额 + Val(Nvl(!预交余额))
                .MoveNext
            Loop
        End With
    End If
    dbl剩余余额 = Format(dbl预交余额 - dbl费用余额, "0.00")
    If mdbl剩余款额 <> dbl剩余余额 Then
        MsgBox "病人的剩余款项已发生变化,请重新确定退款金额!", vbInformation + vbOKOnly, gstrSysName
        Call ShowPremayBalance(False, 0)
        Exit Function
    End If
    
    Check退款 = True
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
    Dim str操作员姓名 As String, strTittle As String
    Dim strNO As String
    
    On Error GoTo errHandle
    '有权限，且为收费状态
    'type: 1-异常充值，2-异常销帐，3-异常余额退款
    strSQL = "Select Type, No , 卡号 ,操作员姓名" & vbNewLine & _
            "From (Select 2 Type, a.No, a.卡号, a.操作员姓名" & vbNewLine & _
            "       From 病人预交记录 a" & vbNewLine & _
            "       Where Nvl(校对标志, 0) <> 0 And 记录性质 = 1 And 病人id = [1] And 记录状态 = 2 " & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 3 Type, a.No, a.卡号, a.操作员姓名" & vbNewLine & _
            "       From 病人预交记录 a" & vbNewLine & _
            "       Where Nvl(校对标志, 0) <>0 And 记录性质 = 1 And 病人id = [1] And 记录状态 = 0 And A.附加标志=1)" & vbNewLine & _
            "Order By Decode(操作员姓名, [2], 0, 1), Type"
    Set rsErrBills = zlDatabase.OpenSQLRecord(strSQL, "病人异常单据查询", lng病人id, UserInfo.姓名)
    If rsErrBills.EOF Then Exit Function
    
    str操作员姓名 = Nvl(rsErrBills!操作员姓名)
    If Nvl(rsErrBills!type) = 2 Then
        strTittle = "销帐"
    Else
        strTittle = "余额退款"
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
    strNO = Nvl(rsErrBills!NO)
    '需重新处理单据
    If mobjEInvoice Is Nothing Then Exit Function
    If frmDeposit.zlShowEdit(Me, mbytCallObject, 7, mobjEInvoice, mstrPrivs, mlngModul, mbytPrepayType, strNO) = False Then Exit Function
    OptOthersErrBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub Excute_Plug_PatiPrePayAfter(ByRef objPati As clsPatiInfo, ByVal lng预交ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行外挂的评价器接口
    '入参:objpati-病人信息对象
    '编制:刘兴洪
    '日期:2018-08-31 10:25:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPlugIn Is Nothing Then Exit Sub
    '81693:李南春,2015/4/21,评价器
    On Error Resume Next
    Call mobjPlugIn.PatiPrePayAfter(objPati.病人ID, IIf(mbytPrepayType = 2, 1, 0), lng预交ID)
    Err.Clear
End Sub

Private Sub vsBlance_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    With vsBlance
       If .ColIndex("退现") = Position Or Col = .ColIndex("退现") Then
            Position = Col
       End If
    End With
End Sub

Private Sub vsBlance_GotFocus()
    vsBlance.BackColorSel = &HFFEBD7
End Sub
Private Sub vsBlance_LostFocus()
   vsBlance.BackColorSel = &HE0E0E0
End Sub

Private Sub vsBlance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim objItem As clsBalanceItem
    Dim strInput As String
    
    With vsBlance
        Select Case Col
        Case .ColIndex("退款方式")
        Case .ColIndex("退现")
            Call ReCalePtBalanceMoney '重新计算退款金额
        Case .ColIndex("退款金额")
            If Not zlGetBalanceItemFromBalanceGrid(Row, objItem) Then Exit Sub
             objItem.结算金额 = roundEx(Val(.TextMatrix(Row, Col)), 2)
             .RowData(Row) = objItem
             Call ReCalePtBalanceMoney '重新计算退款金额
        Case Else
        End Select
    End With
End Sub

Private Sub vsBlance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBlance, Me.Name, "结算列表"
End Sub

Private Sub vsBlance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vsBlance, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsBlance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBlance, Me.Name, "结算列表"
End Sub

Private Sub vsBlance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim objItem As clsBalanceItem
    
    If mpatiInfo.病人ID = 0 Then Cancel = True: Exit Sub
    
    With vsBlance
        If Val(.TextMatrix(Row, .ColIndex("结算状态"))) = 1 Then Cancel = True: Exit Sub
        
        varTemp = Split(.TextMatrix(Row, .ColIndex("编辑状态")) & "|||", "|")
        .ComboList = ""
        
        Select Case Col
        Case .ColIndex("退款方式")
            If Not zlGetBalanceItemFromBalanceGrid(Row, objItem) Then Cancel = True: Exit Sub
            If Val(varTemp(1)) <> 1 Then Cancel = True: Exit Sub
            '是否允许编辑|是否允许删除
             .ColComboList(.ColIndex("退款方式")) = ""
             .ComboList = "..."
             .CellButtonPicture = imgDel
        Case .ColIndex("退现")
            If .TextMatrix(Row, .ColIndex("结算性质")) = 1 Then Cancel = True: Exit Sub
            If Not CheckDelCashColIsEdit(Row) Then Cancel = True: Exit Sub
        Case .ColIndex("退款金额")
            If ChecklDelMoneyIsEdit(.Row) = False Then Cancel = True: Exit Sub
        Case Else
            Cancel = True: Exit Sub
        End Select
    End With
End Sub

Private Sub vsBlance_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsBlance.ColIndex("退现") Then Cancel = True
End Sub

Private Sub vsBlance_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim varData As Variant
    With vsBlance
        '是否允许编辑|是否允许删除
        varData = Split(.TextMatrix(Row, .ColIndex("编辑状态")) & "||", "|")
        If varData(1) <> 1 Then Exit Sub
    End With
    
    Call DeletePayInfor(Row)
    
    Call ReCalePtBalanceMoney '重新计算退款信息
End Sub

Private Sub vsBlance_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
     
     With vsBlance
        If .Row > .Rows - 1 Or .Row < 1 Then Exit Sub
        
        If KeyCode <> vbKeyReturn And (KeyCode <> Asc("*")) And KeyCode <> vbKeySpace _
            And KeyCode <> vbKeyShift Then
            If Shift = 1 And (KeyCode = 56 Or KeyCode <> Asc("*")) Then
                Call vsBlance_CellButtonClick(.Row, .Col)
                Exit Sub
            End If
        End If
        
        '删除
        If KeyCode = vbKeyDelete Then
            Call vsBlance_CellButtonClick(.Row, .Col)
            Exit Sub
        End If
    End With
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsBlance
        Select Case .Col
        Case .ColIndex("退款方式")
            If Trim(.TextMatrix(.Row, .ColIndex("退款方式"))) = "" And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        Case .ColIndex("退款金额")
            If (Trim(.TextMatrix(.Row, .ColIndex("退款方式"))) = "" Or Val(.TextMatrix(.Row, .ColIndex("退款金额"))) = 0) And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        Case Else
            If (Trim(.TextMatrix(.Row, .ColIndex("退款方式"))) = "" Or Val(.TextMatrix(.Row, .ColIndex("退款金额"))) = 0) And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        End Select
        Call zlVsMoveGridCell(vsBlance, .ColIndex("退款方式"), , False, lngRow)
    End With
End Sub

Private Sub vsBlance_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim strKey As String, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsBlance
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "") '暂不处理输入
        Select Case Col
        Case .ColIndex("退款方式")
        Case .ColIndex("退款金额")
        Case Else
        End Select
        Call zlVsMoveGridCell(vsBlance, .ColIndex("退款方式"), -1, False, lngRow)
    End With

End Sub
Private Sub vsBlance_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsBlance
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "") '暂不处理输入
        Select Case Col
        Case .ColIndex("退款方式")
        Case .ColIndex("退款金额")
        Case Else
        End Select
    End With
End Sub


Private Sub vsBlance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Err = 0: On Error GoTo Errhand:
    With vsBlance
        If .MouseRow < 1 Or .MouseRow > .Rows - 1 Then Exit Sub
        If .MouseCol < 0 Or .MouseCol > .Cols - 1 Then Exit Sub
        If .MouseCol = .ColIndex("退现") Then .ToolTipText = "": Exit Sub
        If .ToolTipText = Trim(.TextMatrix(.MouseRow, .MouseCol)) Then Exit Sub
       .ToolTipText = Trim(.TextMatrix(.MouseRow, .MouseCol))
    End With
Errhand:
    Exit Sub
End Sub

Private Sub vsBlance_LeaveCell()
    OS.OpenIme False
End Sub
 
 
Private Sub vsBlance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim objItem As clsBalanceItem
    Dim strInput As String, str结算方式 As String
    
    With vsBlance
        If Row <= 0 Then Exit Sub
        
        Select Case Col
        Case .ColIndex("退现")
            If CheckIsAllowBackCash(Row) = False Then Cancel = True: Exit Sub
        Case .ColIndex("退款金额")
            If Trim(.EditText) = "" Then .EditText = 0
            strInput = Trim(.EditText): strInput = Replace(strInput, Chr(vbKeyReturn), ""): strInput = Replace(strInput, Chr(10), "")
            If Not zlGetBalanceItemFromBalanceGrid(Row, objItem) Then Exit Sub
            str结算方式 = Trim(.TextMatrix(.Row, .ColIndex("退款方式")))
            If Val(Abs(strInput)) > Abs(objItem.剩余金额) Then
                MsgBox "输入的""" & str结算方式 & """退款金额不能超过 " & Format(objItem.剩余金额, "0.00") & " ！", vbInformation, gstrSysName
                .EditCell: .EditSelStart = 0
                .EditSelLength = zlCommFun.ActualLen(.EditText)
                Cancel = True: Exit Sub
            End If
        Case Else
        End Select
    End With
End Sub

Private Function zlGetBalanceItemFromBalanceGrid(ByVal lngRow As Long, ByRef objBalanceItem_Out As clsBalanceItem) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算网格中的数据，提取指定行的BalanceItem数据
    '入参:lngRow-指定的行
    '出参:objBalanceItem-
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-30 15:22:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String, lng卡类别ID As Long, lng消费卡ID As Long
    Dim varTemp As Variant
    Dim objCard As Card
    
    On Error GoTo errHandle
    
    With vsBlance
    
        If lngRow = 0 Then lngRow = .Row
        If lngRow > .Rows - 1 Or lngRow < 1 Then Exit Function
        If UCase(TypeName(.RowData(lngRow))) = UCase("clsBalanceItem") Then
            Set objBalanceItem_Out = .RowData(lngRow)
            If Not objBalanceItem_Out Is Nothing Then zlGetBalanceItemFromBalanceGrid = True: Exit Function
        End If
        
        str结算方式 = .TextMatrix(lngRow, .ColIndex("退款方式"))
        If str结算方式 = "" Then Exit Function
        lng卡类别ID = Val(.TextMatrix(lngRow, .ColIndex("卡类别ID")))
        lng消费卡ID = Val(.TextMatrix(lngRow, .ColIndex("消费卡ID")))
        
        If lng卡类别ID = 0 Then
            Set objCard = mobjThridSwap.zlGetCardFromBalanceName(str结算方式)
        Else
            Call gobjSquare.objSquareCard.zlGetCard(lng卡类别ID, lng消费卡ID <> 0, objCard)
        End If
        
        varTemp = Split(.TextMatrix(lngRow, .ColIndex("编辑状态")) & "|", "|")
        Set objBalanceItem_Out = New clsBalanceItem
        With objBalanceItem_Out
            Set .objCard = objCard
            .关联交易ID = Val(vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("关联交易ID")))
            .交易流水号 = vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("交易流水号"))
            .交易说明 = vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("交易说明"))
            .结算号码 = vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("结算号码"))
            .结算摘要 = vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("备注"))
            .卡号 = vsBlance.Cell(flexcpData, lngRow, vsBlance.ColIndex("卡号"))
            .是否密文 = Val(vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("是否密文"))) = 1
            .结算金额 = Val(vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("退款金额")))
            .是否允许编辑 = Val(varTemp(0)) = 1
            .是否允许删除 = Val(varTemp(1)) = 1
            .限制类别 = CStr(vsBlance.Cell(flexcpData, lngRow, vsBlance.ColIndex("卡类别ID")))
            .消费卡 = lng消费卡ID <> 0
            .消费卡ID = lng消费卡ID
            .卡类别ID = lng卡类别ID
            .密码 = ""
            .校对标志 = Val(vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("校对标志")))
            .结算性质 = objCard.结算性质
            .是否转帐 = IsTransfer(lng卡类别ID)
        End With
       .RowData(lngRow) = objBalanceItem_Out
    End With
    zlGetBalanceItemFromBalanceGrid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function IsTransfer(ByVal lng卡类别ID As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查医疗卡是否支持转账
    ' 参数 :lng卡类别ID-医疗卡类别.id
    ' 日期 : 2019/01/22
    ' 说明 :
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    If lng卡类别ID = 0 Then Exit Function
    If mobjCards("K" & lng卡类别ID) Is Nothing Then Exit Function
    IsTransfer = mobjCards("K" & lng卡类别ID).是否转帐及代扣
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDelCashColIsEdit(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退现现是否允许编译
    '入参:lngRow-指定的行
    '返回:允许编辑返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-31 11:00:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    
    If GetVsGridBoolColVal(vsBlance, lngRow, vsBlance.ColIndex("退现")) = True Then CheckDelCashColIsEdit = True: Exit Function '本身退现，再改为不退现，则允许编辑
    
    If Not zlGetBalanceItemFromBalanceGrid(lngRow, objItem) Then Exit Function
    If objItem.是否允许退现 = False And objItem.是否强制退现 = False Then Exit Function
    If objItem.是否结算 Then Exit Function
    
    If objItem.是否保存 Then '如果已经保存了的,需要调用判断交易是否成功的交易
        'If mobjThridSwap.zlThird_IsSwapIsSucces(objItem, intSwapStatu, strErrMsg) Then Exit Function '交易成功，不允许退现
        'If intSwapStatu <> 0 Then
        '    strNotes = "注意:" & vbCrLf & _
        '    "    " & objCard.名称 & " 交易正在进行中，不能进行退现操作"
        '    If strErrMsg <> "" Then strNotes = strNotes & "，详细错误信息如下：" & vbCrLf & strErrMsg
        '    If strErrMsg = "" Then strNotes = strNotes & "。"
        '    MsgBox strNotes, vbInformation + vbOKOnly, gstrSysName
        '    Exit Function
        'End If
        ''先删除，然后再看能否退现
        '
        Exit Function
    End If
    CheckDelCashColIsEdit = True: Exit Function
End Function

Private Function CheckIsAllowBackCash(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否允许退现
    '入参:lngRow-指定的行
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-31 11:30:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem, intSwapStatu As Integer, strErrMsg As String, strNotes As String
    Dim str操作员姓名 As String
    On Error GoTo errHandle
        
    If GetVsGridBoolColVal(vsBlance, lngRow, vsBlance.ColIndex("退现")) = True Then CheckIsAllowBackCash = True: Exit Function '本身退现，再改为不退现，则允许编辑
    If zlGetBalanceItemFromBalanceGrid(lngRow, objItem) = False Then Exit Function
    
    If objItem.结算性质 = 2 Then CheckIsAllowBackCash = True: Exit Function
    
    If objItem.是否保存 Then '如果已经保存了的,需要调用判断交易是否成功的交易
        If mobjThridSwap.zlThird_IsSwapIsSucces(objItem, intSwapStatu, strErrMsg) Then Exit Function '交易成功，不允许退现
        If intSwapStatu <> 0 Then
            strNotes = "注意:" & vbCrLf & _
            "    " & objItem.objCard.名称 & " 交易正在进行中，不能进行退现操作"
            If strErrMsg <> "" Then strNotes = strNotes & "，详细错误信息如下：" & vbCrLf & strErrMsg
            If strErrMsg = "" Then strNotes = strNotes & "。"
            MsgBox strNotes, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        '明确失败,先删除，然后再看能否退现
        If DelDepositErrBill(objItem.单据号, 2) Then Exit Function
        objItem.是否保存 = False
        objItem.是否结算 = False
        '检查是否允许退现
        Exit Function
    End If
    
    If objItem.是否允许退现 Then CheckIsAllowBackCash = True: Exit Function

    If InStr(";" & mstrCardPrivs & ";", ";三方退款强制退现;") > 0 Then
        '具备强制退现权限
        If MsgBox(objItem.objCard.名称 & " 不支持退现，你是否强制退现？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Function
        objItem.Tag = UserInfo.姓名
        CheckIsAllowBackCash = True
        Exit Function
    End If
    
    '已经验证过的，不再验证
    str操作员姓名 = zlDatabase.UserIdentifyByUser(Me, "强制退现验证", glngSys, 1151, "三方退款强制退现")
    If str操作员姓名 = "" Then
        MsgBox "录入的操作员验证失败或者录入的操作员不具备强制退现权限，不能强制退现！", vbInformation, gstrSysName
        Exit Function
    End If
    objItem.Tag = str操作员姓名



    CheckIsAllowBackCash = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub DeletePayInfor(ByVal lngDelRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除行
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-31 11:38:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, objItem   As clsBalanceItem
    
    On Error GoTo errHandle
    With vsBlance
        If lngDelRow > .Rows - 1 Or lngDelRow < 1 Then Exit Sub
        If zlGetBalanceItemFromBalanceGrid(lngDelRow, objItem) = False Then Exit Sub
        
        If objItem.是否结算 Then Exit Sub
        If objItem.是否保存 Then
            If mobjThridSwap.zlThird_IsCancelFromItems(objItem) = False Then Exit Sub
            '明确失败,先删除，然后再看能否退现
            If DelDepositErrBill(objItem.单据号, 2) = False Then Exit Sub
        End If
        lngRow = lngDelRow
        If .Rows <= 2 Then
            .Clear 1
            .RowData(1) = "": Set objItem = Nothing
            .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        Else
            .RowData(lngDelRow) = ""
            Set objItem = Nothing
            vsBlance.RemoveItem lngDelRow
        End If
        
        If lngRow <= 1 Then
            lngRow = 1
        ElseIf lngRow >= .Rows - 1 Then
            lngRow = .Rows - 1
        Else
            lngRow = lngDelRow + 1
        End If
        If lngRow > .Rows - 1 Or lngRow <= 1 Then lngRow = 1
        .Row = lngRow
        If .RowIsVisible(.Row) = False Then .ShowCell .Row, .Col
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub AutoShareBalanceMoney(ByVal dblMoney As Double, Optional ByVal blnAllMoney As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动分摊费用
    '入参:dblMoney-退款金额
    '       blnAllMoney-是否分摊所有费用（现金+三方）
    '编制:刘兴洪
    '日期:2018-09-07 09:42:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblCashMoney As Double, dblTotal As Double, dblThirdDelMoney As Double
    Dim objItem As clsBalanceItem
    
    On Error GoTo errHandle
    If dblMoney < 0 And blnAllMoney Then Exit Sub
    If dblMoney < 0 Then dblMoney = 0
    dblTotal = dblMoney
    With vsBlance
        dblCashMoney = 0: dblThirdDelMoney = 0
        '先合并负数部分
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("退款方式")) <> "" Then
                If zlGetBalanceItemFromBalanceGrid(i, objItem) Then
                   If GetVsGridBoolColVal(vsBlance, i, .ColIndex("退现")) Or blnAllMoney Then
                        If objItem.结算金额 < 0 Then dblTotal = roundEx(dblTotal - objItem.结算金额, 5)
                   End If
                End If
            End If
        Next
        
        '再分摊金额
        dblThirdDelMoney = 0: dblCashMoney = 0
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("退款方式")) <> "" Then
                If zlGetBalanceItemFromBalanceGrid(i, objItem) Then
                   If GetVsGridBoolColVal(vsBlance, i, .ColIndex("退现")) Then
                        If objItem.剩余金额 > 0 Then
                            If dblTotal > objItem.剩余金额 Then
                                dblTotal = roundEx(dblTotal - objItem.剩余金额, 5)
                                objItem.结算金额 = objItem.剩余金额
                            Else
                                objItem.结算金额 = dblTotal
                                dblTotal = 0
                            End If
                            .RowData(i) = objItem
                            .TextMatrix(i, .ColIndex("退款金额")) = Format(objItem.结算金额, "####0.00" & IIf(objItem.结算性质 = 9, "####", ""))
                        End If
                        dblCashMoney = roundEx(dblCashMoney + objItem.结算金额, 5)
                   Else
                        If objItem.剩余金额 > 0 And blnAllMoney Then
                            If dblTotal > objItem.剩余金额 Then
                                dblTotal = roundEx(dblTotal - objItem.剩余金额, 5)
                                objItem.结算金额 = objItem.剩余金额
                            Else
                                objItem.结算金额 = dblTotal
                                dblTotal = 0
                            End If
                            .RowData(i) = objItem
                            .TextMatrix(i, .ColIndex("退款金额")) = Format(objItem.结算金额, "####0.00" & IIf(objItem.结算性质 = 9, "####", ""))
                        End If
                        dblThirdDelMoney = roundEx(dblThirdDelMoney + objItem.结算金额, 5)
                   End If
                End If
            End If
        Next
    End With
    txtCashTotal.Text = Format(dblCashMoney, "#,##0.00")
    lblCashTotal.Tag = dblCashMoney
    txtMoney.Text = Format(dblCashMoney, "#,##0.00")
    dblCashMoney = roundEx(dblCashMoney, 6)
    txtThirdTotal.Text = Format(dblThirdDelMoney, "#,##0.00")
    If dblCashMoney <> dblMoney And Not blnAllMoney Then
      If MsgBox("当前输入的金额未分摊完成，是否以分摊的退现金额为本次退款金额?" & vbCrLf & vbCrLf & _
               "当前输入：" & Format(dblMoney, "0.00") & vbCrLf & _
               "分摊退现：" & Format(dblCashMoney, "0.00") & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            txtMoney.Text = Format(dblCashMoney, "#,##0.00")
       End If
    End If
    txtTotal.Text = Format(dblCashMoney + dblThirdDelMoney, "#,##0.00")
    Call LoadThirdTotal
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitTab()
    '功能：初始化分页控件
    Dim objItem As TabControlItem
    
    Err = 0: On Error GoTo Errhand:
    With tbPage
        picDeposit.BorderStyle = 0
        picDepositBack.BorderStyle = 0
        picDepositHistory.BorderStyle = 0
        picDeposit.BackColor = &H8000000F
        picDepositBack.BackColor = &H8000000F
        picDepositHistory.BackColor = &H8000000F
        
        Set objItem = .InsertItem(pg_Page.pg_预交余额退款, "退款列表", picDepositBack.hwnd, 0)
        objItem.Tag = pg_Page.pg_预交余额退款
        Set objItem = .InsertItem(pg_Page.pg_预交历史记录, "历史记录", picDepositHistory.hwnd, 0)
        objItem.Tag = pg_Page.pg_预交历史记录
        
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .Item(0).Selected = True
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ShowHistoryPrepay()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示历史的预交数据
    '编制:刘兴洪
    '日期:2011-09-16 10:17:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, int类型 As Integer, lngRow As Long, strWhere As String
    Dim rsMoney As ADODB.Recordset
    Dim lng病人id As Long, i As Integer
    
    If mpatiInfo.病人ID = 0 Then
        lng病人id = mlng病人ID
    Else
        lng病人id = mpatiInfo.病人ID
    End If
    
    If cboType.ListIndex < 0 Then
         int类型 = 1
    Else
        int类型 = cboType.ItemData(cboType.ListIndex)
    End If
    
    On Error GoTo errHandle
    '84217,李南春,2015/4/22,显示指定的住院期间缴纳的预交
    If cboType.Text = "住院预交" And cboPatiPage.ListIndex > 0 Then
        strWhere = " And A.主页ID= " & cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    
    If gbln分院区显示 Then
        strWhere = strWhere & _
                " And Exists (Select 1 From 人员表 C, 部门人员 D, 部门表 E " & _
                " Where C.姓名 =A.操作员姓名 And C.Id = D.人员id And D.部门id = E.Id And (E.站点 = '" & gstrNodeNo & "' Or E.站点 Is Null))"
    End If
            
    If gblnShowHave Then
        '只显示有剩余的历史缴款
        '该子查阅用于消除第一次结帐时的一正一负
        strSQL = _
        "   Select NO,Sum(Nvl(A.金额,0)) as 金额  " & _
        "    From 病人预交记录 A" & _
        "   Where A.结帐ID Is Null And Nvl(A.金额, 0)<>0 And A.病人ID=[1] And A.预交类别=[2] " & _
        "   Group by NO " & _
        "   Having Sum(Nvl(A.金额,0))<>0"
        
        strSQL = _
        " Select LTrim(To_Char(A.收款时间,'YYYY-MM-DD')) as 日期,A.NO as 单据号,A.实际票号 as 票据号," & _
        "           C.名称 as 科室,Ltrim(To_Char(Nvl(A.金额,0),'9,999,999,990.00')) as 缴款金额,A.结算方式 as 结算,A.操作员姓名 as 收款人" & _
        " From 病人预交记录 A,(" & strSQL & ") B,部门表 C" & _
        " Where A.结帐ID Is Null And A.预交类别=[2]  And Nvl(A.金额,0)<>0 And A.科室ID=C.ID(+)" & _
        "       And A.结算方式 Not IN(Select 名称 From 结算方式 Where 性质=5)" & _
        "       And A.NO=B.NO And A.病人ID=[1] And Not Exists (Select 1 From 病人预交记录 Where No = a.No And Nvl(校对标志, 0) <> 0 And 病人ID=[1]) " & strWhere & _
        " Union All" & _
        " Select Min(LTrim(To_Char(A.收款时间,'YYYY-MM-DD'))) as 日期,A.NO as 单据号,Max(A.实际票号) as 票据号," & _
        "           B.名称 as 科室,Ltrim(To_Char(Sum(Nvl(A.金额,0)-Nvl(A.冲预交,0)),'9,999,999,990.00')) as 缴款金额,A.结算方式 as 结算,A.操作员姓名 as 收款人" & _
        " From 病人预交记录 A,部门表 B" & _
        " Where A.记录性质 IN(1,11) And A.结帐ID is Not NULL And A.科室ID=B.ID(+) And A.预交类别=[2] " & _
        "       And Nvl(A.金额,0)<>Nvl(A.冲预交,0) And A.病人ID=[1] And Not Exists (Select 1 From 病人预交记录 Where No = a.No And Nvl(校对标志, 0) <> 0 And 病人ID=[1]) " & strWhere & _
        " Having Sum(Nvl(A.金额,0)-Nvl(A.冲预交,0))<>0" & _
        " Group by A.NO,B.名称,A.结算方式,A.操作员姓名" & _
        " Order by 日期,单据号,结算"
    Else
        '所有历史缴款明细清单
        strSQL = _
        " Select Ltrim(To_Char(A.收款时间,'YYYY-MM-DD')) as 日期,A.NO as 单据号,A.实际票号 as 票据号,B.名称 as 科室, " & _
        " Ltrim(To_Char(A.金额,'9,999,999,990.00')) as 缴款金额,A.结算方式 as 结算,A.操作员姓名 as 收款人 " & _
        " From 病人预交记录 A,部门表 B" & _
        " Where A.科室ID=B.ID(+) And A.记录性质=1 And A.病人ID=[1]  And A.预交类别=[2] " & _
        " And Not Exists (Select 1 From 病人预交记录 Where No = a.No And Nvl(校对标志, 0) <> 0 And 病人ID=[1]) " & strWhere & _
        " Order by A.收款时间 Desc"
    End If
    
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人id, int类型)
    If Not rsMoney.EOF Then
        With vsDepositHistory
            If gblnShowHave Then
                .TextMatrix(0, .ColIndex("缴款金额")) = "剩余金额"
            Else
                .TextMatrix(0, .ColIndex("缴款金额")) = "缴款金额"
            End If
            .Rows = rsMoney.RecordCount + 1
            For i = 1 To rsMoney.RecordCount
                .TextMatrix(i, .ColIndex("日期")) = Nvl(rsMoney!日期)
                .TextMatrix(i, .ColIndex("单据号")) = Nvl(rsMoney!单据号)
                .TextMatrix(i, .ColIndex("票据号")) = Nvl(rsMoney!票据号)
                .TextMatrix(i, .ColIndex("科室")) = Nvl(rsMoney!科室)
                .TextMatrix(i, .ColIndex("缴款金额")) = Nvl(rsMoney!缴款金额)
                .TextMatrix(i, .ColIndex("结算")) = Nvl(rsMoney!结算)
                .TextMatrix(i, .ColIndex("收款人")) = Nvl(rsMoney!收款人)
                rsMoney.MoveNext
            Next
        End With
    End If
    If vsDepositHistory.Rows > 1 Then
        vsDepositHistory.Row = 1: vsDepositHistory.Col = 0: vsDepositHistory.ColSel = vsDepositHistory.Cols - 1
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ChecklDelMoneyIsEdit(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退款金额是否允许编辑
    '入参:lngRow-指定的行
    '返回:允许编辑返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-31 11:00:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    If Not zlGetBalanceItemFromBalanceGrid(lngRow, objItem) Then Exit Function
    If objItem.objCard.是否全退 Then Exit Function
    If objItem.是否结算 Then Exit Function
    ChecklDelMoneyIsEdit = True: Exit Function
End Function

Private Sub vsDepositHistory_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsDepositHistory, Me.Name, "预交清单"
End Sub

Private Sub vsDepositHistory_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsDepositHistory, Me.Name, "预交清单"
End Sub

Private Sub LoadThirdTotal()
    '功能:加载三方退款汇总列表
    Dim str结算方式 As String, strThirdMoney As String, strTmp As String
    Dim i As Integer, j As Integer, dblThird As Double
    Dim var结算方式 As Variant, varData As Variant, varTmp As Variant
    
    On Error GoTo errHandle
    
    vsThirdTotal.Rows = 2: vsThirdTotal.Cell(flexcpText, 1, 0, 1, 1) = ""
    With vsBlance
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("退款方式")) <> "" And Val(.TextMatrix(i, .ColIndex("退款金额"))) <> 0 Then
                If .TextMatrix(i, .ColIndex("退现")) = 0 Then
                    If InStr("," & str结算方式 & ",", "," & .TextMatrix(i, .ColIndex("退款方式")) & ",") = 0 Then
                        str结算方式 = str结算方式 & "," & .TextMatrix(i, .ColIndex("退款方式"))
                    End If
                    strThirdMoney = strThirdMoney & "|" & .TextMatrix(i, .ColIndex("退款方式")) & "," & Val(.TextMatrix(i, .ColIndex("退款金额")))
                End If
            End If
        Next
        
        str结算方式 = Mid(str结算方式, 2)
        strThirdMoney = Mid(strThirdMoney, 2)
        If str结算方式 = "" Or strThirdMoney = "" Then Exit Sub
        var结算方式 = Split(str结算方式, ",")
        varData = Split(strThirdMoney, "|")
        For i = 0 To UBound(var结算方式)
            dblThird = 0
            For j = 0 To UBound(varData)
                varTmp = Split(varData(j), ",")
                If var结算方式(i) = varTmp(0) Then
                    dblThird = dblThird + Val(varTmp(1))
                End If
            Next
            strTmp = strTmp & "|" & var结算方式(i) & "," & dblThird
        Next
        
        strTmp = Mid(strTmp, 2)
        If strTmp = "" Then Exit Sub
    End With
    
    With vsThirdTotal
        varData = Split(strTmp, "|")
        .Rows = UBound(varData) + 2
        For i = 1 To UBound(varData) + 1
            varTmp = Split(varData(i - 1), ",")
            .TextMatrix(i, .ColIndex("退款方式")) = varTmp(0)
            .TextMatrix(i, .ColIndex("退款金额")) = Format(varTmp(1), "0.00")
        Next
        .ColWidth(.ColIndex("退款金额")) = IIf(.Rows * .RowHeight(0) <= .Height, 1855, 1620)
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsDepositHistory_GotFocus()
    vsDepositHistory.BackColorSel = &HFFEBD7
End Sub

Private Sub vsDepositHistory_LostFocus()
    vsDepositHistory.BackColorSel = &HE0E0E0
End Sub

Private Sub vsThirdTotal_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsThirdTotal, Me.Name, "三方退款汇总"
End Sub

Private Sub vsThirdTotal_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsThirdTotal, Me.Name, "三方退款汇总"
End Sub

Private Sub vsThirdTotal_GotFocus()
     vsThirdTotal.BackColorSel = &HFFEBD7
End Sub

Private Sub vsThirdTotal_LostFocus()
     vsThirdTotal.BackColorSel = &HE0E0E0
End Sub

Private Sub InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区域
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane

    Err = 0: On Error GoTo ErrHandler
    
    Set objPane = dkpMain.CreatePane(PaneId.EM_Head, 150, 30, DockTopOf, Nothing)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable
    objPane.Tag = PaneId.EM_Head: objPane.Handle = picNO.hwnd
    objPane.MaxTrackSize.Height = 30: objPane.MinTrackSize.Height = 30
    
    Set objPane = dkpMain.CreatePane(PaneId.EM_PatiInfo, 150, 150, DockBottomOf, objPane)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Tag = PaneId.EM_PatiInfo: objPane.Handle = picInfo.hwnd
    objPane.MaxTrackSize.Height = 150: objPane.MinTrackSize.Height = 30
    
    Set objPane = dkpMain.CreatePane(PaneId.EM_BillList, 150, 430, DockBottomOf, objPane)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Tag = PaneId.EM_BillList: objPane.Handle = picFace.hwnd
    objPane.MinTrackSize.Height = 430
    
    Set objPane = dkpMain.CreatePane(PaneId.EM_Cmd, 150, 30, DockBottomOf, objPane)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Tag = PaneId.EM_Cmd
    objPane.MaxTrackSize.Height = 30: objPane.MinTrackSize.Height = 30
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True '实时拖动
        .Options.AlphaDockingContext = True
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

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
    If GetPatiPageInforByID(str病人id, PatiPageInfo, blnLastTime) = False Then GetPatiInfo = True: Exit Function
    If PatiPageInfo.病人ID > 0 Then
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

Private Function zlCancelEInvoiceBat(ByVal objPati As clsPatientInfo, Optional ByRef strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:批量作废电子发票（余额退款）
    '入参:objPati-当前病人信息
    '       strNos-预交单据号，多个用逗号分隔
    '返回:执行成功返回true,否则返回False
    '编制:焦博
    '日期:2020-04-07 17:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, int险类 As Integer, int预交类别 As Integer
    
    On Error GoTo errHandle
    If mobjEInvoice Is Nothing Then Exit Function
    int预交类别 = cboType.ItemData(cboType.ListIndex)
    With vsBlance
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("退款方式")) <> "" And Val(.TextMatrix(i, .ColIndex("退款金额"))) <> 0 Then
                int险类 = IIf(Val(.TextMatrix(i, .ColIndex("结算性质"))) = 3, mpatiInfo.险类, 0)
                objPati.险类 = int险类
                strNos = strNos & "," & .TextMatrix(i, .ColIndex("单据号"))
                If mobjEInvoice.zlCancelEInvoiceFromBalanceInfor(Me, objPati, .TextMatrix(i, .ColIndex("单据号"))) = False Then Exit Function
            End If
        Next
    End With
    strNos = Mid(strNos, 2)
    zlCancelEInvoiceBat = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlCreateEInvoiceBat(ByVal objPati As clsPatientInfo, ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:批量开具电子发票（余额退款）
    '入参:objPati-当前病人信息
    '       strNos-预交单号，多个用逗号分隔
    '返回:执行成功返回true,否则返回False
    '编制:焦博
    '日期:2020-04-07 17:37:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllSwapData As Collection, strDate As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim int预交类别 As Integer, int险类 As Integer
    
    On Error GoTo errHandle
    If mobjEInvoice Is Nothing Then Exit Function
    If strNos = "" Then Exit Function
    strDate = "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')"
    int预交类别 = cboType.ItemData(cboType.ListIndex)
    strSQL = "" & _
            "Select a.No, a.Id As 冲销id, c.原预交id, a.病人id, a.预交类别, -1 * a.冲预交 As 冲销金额, a.结算方式, d.性质" & vbNewLine & _
            "From 病人预交记录 A," & vbNewLine & _
            "     (Select 结帐id" & vbNewLine & _
            "              From 病人预交记录" & vbNewLine & _
            "              Where 病人id = [1]  And 记录性质 = 1 And 附加标志 = 1) B," & vbNewLine & _
            "     (Select a.No, a.Id As 原预交id" & vbNewLine & _
            "       From 病人预交记录 A, Table(f_Str2List([2])) B" & vbNewLine & _
            "       Where a.记录性质 = 1 And a.记录状态 = 1 And a.No = b.Column_Value) C, 结算方式 D " & vbNewLine & _
            "Where a.病人id = [1]  And a.记录性质 = 11 And a.结帐id = b.结帐id And Nvl(a.冲预交, 0) > 0 And a.No = c.No And a.结算方式 = d.名称(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取预交余额", objPati.病人ID, strNos, int预交类别)
    If rsTemp.EOF Then zlCreateEInvoiceBat = True: Exit Function
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            int险类 = IIf(Val(Nvl(!性质)) = 3, mpatiInfo.险类, 0)
            If mobjEInvoice.zlIsStartEInvoice(int险类, int预交类别) And Nvl(!结算方式) <> "" Then
                Set cllSwapData = Nothing
                Call GetFact(False, int险类)
                objPati.险类 = int险类
                If mobjEInvoice.zlGetEinvoiceSwapCollect(objPati, Nvl(!原预交ID), Nvl(!NO), Val(Nvl(!冲销金额)), strDate, txtFact.Text, cllSwapData, Nvl(!冲销ID), mlng领用ID) Then
                    '开具电子票据
                    Call mobjEInvoice.zlCreateEInvoice(Me, cllSwapData, , , 2, 1, False)
                End If
            End If
            .MoveNext
        Loop
    End With
       
    zlCreateEInvoiceBat = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


