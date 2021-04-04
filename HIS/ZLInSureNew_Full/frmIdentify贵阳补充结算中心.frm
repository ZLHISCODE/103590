VERSION 5.00
Begin VB.Form frmIdentify贵阳补充结算中心 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "补充录入医保中心结算信息"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   Icon            =   "frmIdentify贵阳补充结算中心.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame frmDetail 
      Height          =   7110
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   9870
      Begin VB.TextBox txt公务员补助起付线 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   44
         Top             =   4215
         Width           =   1860
      End
      Begin VB.TextBox txt普通门诊公务员补助累计 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   53
         Top             =   4680
         Width           =   1860
      End
      Begin VB.TextBox txt公务员补助起付标准 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   47
         Top             =   4215
         Width           =   1860
      End
      Begin VB.TextBox txt超大额限额公务员补助 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   50
         Top             =   4680
         Width           =   1860
      End
      Begin VB.CommandButton cmd公务员补助起付线 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify贵阳补充结算中心.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   4215
         Width           =   300
      End
      Begin VB.CommandButton cmd普通门诊公务员补助累计 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify贵阳补充结算中心.frx":006C
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   4680
         Width           =   300
      End
      Begin VB.CommandButton cmd公务员补助起付标准 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify贵阳补充结算中心.frx":00CC
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   4215
         Width           =   300
      End
      Begin VB.CommandButton cmd超大额限额公务员补助 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify贵阳补充结算中心.frx":012C
         Style           =   1  'Graphical
         TabIndex        =   51
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   4680
         Width           =   300
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1170
         Width           =   2220
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1215
         Width           =   2220
      End
      Begin VB.TextBox txt卡号 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   255
         Width           =   2220
      End
      Begin VB.TextBox txt病人ID 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   2220
      End
      Begin VB.TextBox txt主页ID 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   780
         Width           =   2220
      End
      Begin VB.TextBox txt医保号 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   315
         Width           =   2220
      End
      Begin VB.TextBox txt统筹支付 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   26
         Top             =   2820
         Width           =   1860
      End
      Begin VB.TextBox txt统筹自付 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   29
         Top             =   2820
         Width           =   1860
      End
      Begin VB.TextBox txt全自付 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   14
         Top             =   1905
         Width           =   1860
      End
      Begin VB.TextBox txt起付线 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   20
         Top             =   2370
         Width           =   1860
      End
      Begin VB.TextBox txt基数自付 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   23
         Top             =   2370
         Width           =   1860
      End
      Begin VB.TextBox txt公务员补助 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   41
         Top             =   3750
         Width           =   1860
      End
      Begin VB.TextBox txt结算总费用 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   56
         Top             =   5160
         Width           =   1860
      End
      Begin VB.TextBox txt大病统筹 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   32
         Top             =   3285
         Width           =   1860
      End
      Begin VB.TextBox txt超限自付 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   38
         Top             =   3750
         Width           =   1860
      End
      Begin VB.TextBox txt医保总费用 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   59
         Top             =   5160
         Width           =   1860
      End
      Begin VB.TextBox txt大病自付 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   35
         Top             =   3285
         Width           =   1860
      End
      Begin VB.TextBox txt就诊顺序号 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   68
         Top             =   6135
         Width           =   1860
      End
      Begin VB.TextBox txt结算日期 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   71
         Top             =   6135
         Width           =   1860
      End
      Begin VB.TextBox txtHIS总费用 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   62
         Top             =   5640
         Width           =   1860
      End
      Begin VB.TextBox txt结算编号 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   65
         Top             =   5640
         Width           =   1860
      End
      Begin VB.CommandButton cmd起付线 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify贵阳补充结算中心.frx":018C
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   2370
         Width           =   300
      End
      Begin VB.CommandButton cmd全自付 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify贵阳补充结算中心.frx":01EC
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   1905
         Width           =   300
      End
      Begin VB.CommandButton cmd超限自付 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify贵阳补充结算中心.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   3750
         Width           =   300
      End
      Begin VB.CommandButton cmd大病统筹 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify贵阳补充结算中心.frx":02AC
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   3285
         Width           =   300
      End
      Begin VB.CommandButton cmd公务员补助 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify贵阳补充结算中心.frx":030C
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   3750
         Width           =   300
      End
      Begin VB.CommandButton cmd结算日期 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify贵阳补充结算中心.frx":036C
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   6135
         Width           =   300
      End
      Begin VB.CommandButton cmd结算编号 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify贵阳补充结算中心.frx":03CC
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   5640
         Width           =   300
      End
      Begin VB.CommandButton cmd结算总费用 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify贵阳补充结算中心.frx":042C
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   5160
         Width           =   300
      End
      Begin VB.CommandButton cmd医保总费用 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify贵阳补充结算中心.frx":048C
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   5160
         Width           =   300
      End
      Begin VB.CommandButton cmd大病自付 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify贵阳补充结算中心.frx":04EC
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   3285
         Width           =   300
      End
      Begin VB.CommandButton cmd统筹自付 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify贵阳补充结算中心.frx":054C
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   2820
         Width           =   300
      End
      Begin VB.CommandButton cmd基数自付 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify贵阳补充结算中心.frx":05AC
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   2370
         Width           =   300
      End
      Begin VB.CommandButton cmd挂钩自付 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify贵阳补充结算中心.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   1905
         Width           =   300
      End
      Begin VB.CommandButton cmd就诊顺序号 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify贵阳补充结算中心.frx":066C
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   6135
         Width           =   300
      End
      Begin VB.CommandButton cmdHIS总费用 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify贵阳补充结算中心.frx":06CC
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   5640
         Width           =   300
      End
      Begin VB.CommandButton cmd统筹支付 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify贵阳补充结算中心.frx":072C
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   2820
         Width           =   300
      End
      Begin VB.TextBox txt挂钩自付 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   17
         Top             =   1905
         Width           =   1860
      End
      Begin VB.TextBox txt特殊结算说明 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   77
         Top             =   6600
         Width           =   1860
      End
      Begin VB.TextBox txt特殊结算方式 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   74
         Top             =   6600
         Width           =   1860
      End
      Begin VB.CommandButton cmd特殊结算说明 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify贵阳补充结算中心.frx":078C
         Style           =   1  'Graphical
         TabIndex        =   78
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   6600
         Width           =   300
      End
      Begin VB.CommandButton cmd特殊结算方式 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify贵阳补充结算中心.frx":07EC
         Style           =   1  'Graphical
         TabIndex        =   75
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   6600
         Width           =   300
      End
      Begin VB.Label lab公务员补助起付线 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "公务员补助起付线"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   600
         TabIndex        =   43
         Top             =   4260
         Width           =   1680
      End
      Begin VB.Label lab普通门诊公务员补助累计 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "普通门诊公务员补助累计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4755
         TabIndex        =   52
         Top             =   4725
         Width           =   2310
      End
      Begin VB.Label lab公务员补助起付标准 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "公务员补助起付标准"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5175
         TabIndex        =   46
         Top             =   4260
         Width           =   1890
      End
      Begin VB.Label lab超大额限额公务员补助 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "超大额限额公务员补助"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   180
         TabIndex        =   49
         Top             =   4725
         Width           =   2100
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0080FFFF&
         X1              =   -150
         X2              =   15850
         Y1              =   1710
         Y2              =   1710
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   -150
         X2              =   15850
         Y1              =   1695
         Y2              =   1695
      End
      Begin VB.Label lbl姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1860
         TabIndex        =   9
         Top             =   1215
         Width           =   420
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6435
         TabIndex        =   11
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1860
         TabIndex        =   1
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1650
         TabIndex        =   5
         Top             =   765
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "主页ID"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6435
         TabIndex        =   7
         Top             =   825
         Width           =   630
      End
      Begin VB.Label lbl卡号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6435
         TabIndex        =   3
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lab统筹支付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹支付"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1440
         TabIndex        =   25
         Top             =   2865
         Width           =   840
      End
      Begin VB.Label lab统筹自付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹自付"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6225
         TabIndex        =   28
         Top             =   2865
         Width           =   840
      End
      Begin VB.Label lab全自付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "全自付"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1650
         TabIndex        =   13
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label lab起付线 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "起付线"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1650
         TabIndex        =   19
         Top             =   2415
         Width           =   630
      End
      Begin VB.Label lab基数自付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "基数自付"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6225
         TabIndex        =   22
         Top             =   2415
         Width           =   840
      End
      Begin VB.Label lab挂钩自付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "挂钩自付"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6225
         TabIndex        =   16
         Top             =   1950
         Width           =   840
      End
      Begin VB.Label lab公务员补助 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "公务员补助"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6015
         TabIndex        =   40
         Top             =   3795
         Width           =   1050
      End
      Begin VB.Label lab结算总费用 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结算总费用"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1230
         TabIndex        =   55
         Top             =   5205
         Width           =   1050
      End
      Begin VB.Label lab大病统筹 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "大病统筹"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1440
         TabIndex        =   31
         Top             =   3330
         Width           =   840
      End
      Begin VB.Label lab超限自付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "超限自付"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1440
         TabIndex        =   37
         Top             =   3795
         Width           =   840
      End
      Begin VB.Label lab医保总费用 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保总费用"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6015
         TabIndex        =   58
         Top             =   5175
         Width           =   1050
      End
      Begin VB.Label lab大病自付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "大病自付"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6225
         TabIndex        =   34
         Top             =   3330
         Width           =   840
      End
      Begin VB.Label lab就诊顺序号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "就诊顺序号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1230
         TabIndex        =   67
         Top             =   6180
         Width           =   1050
      End
      Begin VB.Label lab结算日期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结算日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6225
         TabIndex        =   70
         Top             =   6180
         Width           =   840
      End
      Begin VB.Label labHIS总费用 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "HIS总费用"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1335
         TabIndex        =   61
         Top             =   5685
         Width           =   945
      End
      Begin VB.Label lab结算编号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结算编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6225
         TabIndex        =   64
         Top             =   5685
         Width           =   840
      End
      Begin VB.Label lab特殊结算说明 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "特殊结算说明"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5805
         TabIndex        =   76
         Top             =   6645
         Width           =   1260
      End
      Begin VB.Label lab特殊结算方式 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "特殊结算方式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1020
         TabIndex        =   73
         Top             =   6645
         Width           =   1260
      End
   End
   Begin VB.PictureBox P2 
      Height          =   495
      Left            =   1440
      Picture         =   "frmIdentify贵阳补充结算中心.frx":084C
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   82
      Top             =   7335
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox P1 
      Height          =   495
      Left            =   75
      Picture         =   "frmIdentify贵阳补充结算中心.frx":092A
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   81
      Top             =   7335
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Height          =   435
      Left            =   8610
      TabIndex        =   80
      Top             =   7365
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&0)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7110
      TabIndex        =   79
      Top             =   7365
      Width           =   1335
   End
End
Attribute VB_Name = "frmIdentify贵阳补充结算中心"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_OkCancel            As Boolean

Public Property Get OkCancel() As Boolean
    OkCancel = m_OkCancel
End Property

Private Sub cmdCancel_Click()
    With g补充结算
        .blnYn = False
    End With
    Unload Me
End Sub

Private Sub cmdHIS总费用_Click()
    txtHIS总费用.Text = Format(Val(txtHIS总费用.Text), "0.00")
    cmdHIS总费用.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmdHIS总费用.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txtHIS总费用.SetFocus
    txtHIS总费用.SelStart = 0
    txtHIS总费用.SelLength = Len(txtHIS总费用.Text)
End Sub

Private Sub cmdOK_Click()
    If Val(txt全自付.Text) < 0 Then
        txt全自付.SetFocus
        Exit Sub
    End If
    If Val(txt挂钩自付.Text) < 0 Then
        txt挂钩自付.SetFocus
        Exit Sub
    End If
    If Val(txt起付线.Text) < 0 Then
        txt起付线.SetFocus
        Exit Sub
    End If
    
    If Val(txt基数自付.Text) < 0 Then
        txt基数自付.SetFocus
        Exit Sub
    End If
    If Val(txt统筹支付.Text) < 0 Then
        txt统筹支付.SetFocus
        Exit Sub
    End If
    If Val(txt统筹自付.Text) < 0 Then
        txt统筹自付.SetFocus
        Exit Sub
    End If
    If Val(txt大病统筹.Text) < 0 Then
        txt大病统筹.SetFocus
        Exit Sub
    End If
    If Val(txt大病自付.Text) < 0 Then
        txt大病自付.SetFocus
        Exit Sub
    End If
    If Val(txt超限自付.Text) < 0 Then
        txt超限自付.SetFocus
        Exit Sub
    End If
    If Val(txt医保总费用.Text) < 0 Then
        txt医保总费用.SetFocus
        Exit Sub
    End If
    If Val(txt公务员补助.Text) < 0 Then
        txt公务员补助.SetFocus
        Exit Sub
    End If
    
    If Val(txt公务员补助起付线.Text) < 0 Then
        txt公务员补助起付线.SetFocus
        Exit Sub
    End If
    
    If Val(txt公务员补助起付标准.Text) < 0 Then
        txt公务员补助起付标准.SetFocus
        Exit Sub
    End If
    
    If Val(txt超大额限额公务员补助.Text) < 0 Then
        txt超大额限额公务员补助.SetFocus
        Exit Sub
    End If
    
    If Val(txt普通门诊公务员补助累计.Text) < 0 Then
        txt普通门诊公务员补助累计.SetFocus
        Exit Sub
    End If
    
    If Val(txt结算总费用.Text) < 0 Then
        txt结算总费用.SetFocus
        Exit Sub
    End If
    If Val(txtHIS总费用.Text) < 0 Then
        MsgBox "HIS总费用必须大于0", vbCritical, gstrSysName
        txtHIS总费用.SetFocus
        Exit Sub
    End If
    If Len(txt结算编号.Text) <= 0 Then
        MsgBox "结算编号不能为空！", vbCritical, gstrSysName
        txt结算编号.SetFocus
        Exit Sub
    End If
    If Len(txt就诊顺序号.Text) < 0 Then
        MsgBox "就诊顺序号不能为空！", vbCritical, gstrSysName
        txt就诊顺序号.SetFocus
        Exit Sub
    End If
    If Not IsDate(txt结算日期.Text) Then
        MsgBox "结算日期必须为日期类型！", vbCritical, gstrSysName
        txt结算日期.SetFocus
        Exit Sub
    End If
    With g补充结算
        .blnYn = True
        .m_全自付 = Val(txt全自付.Text)
        .m_挂钩自付 = Val(txt挂钩自付.Text)
        .m_起付线 = Val(txt起付线.Text)
        .m_基数自付 = Val(txt基数自付.Text)
        .m_统筹支付 = Val(txt统筹支付.Text)
        .m_统筹自付 = Val(txt统筹自付.Text)
        .m_大病统筹 = Val(txt大病统筹.Text)
        .m_大病自付 = Val(txt大病自付.Text)
        .m_超限自付 = Val(txt超限自付.Text)
        .m_医保总费用 = Val(txt医保总费用.Text)
        .m_公务员补助 = Val(txt公务员补助.Text)
        .m_结算总费用 = Val(txt结算总费用.Text)
        .m_HIS总费用 = Val(txtHIS总费用.Text)
        .m_结算编号 = txt结算编号.Text
        .m_就诊顺序号 = txt就诊顺序号.Text
        .m_结算日期 = Format(txt结算日期.Text, "yyyy-mm-dd hh:mm:ss")
        .m_公务员补助起付标准 = Val(txt公务员补助起付标准.Text)
        .m_公务员补助起付线 = Val(txt公务员补助起付线.Text)
        .m_普通门诊公务员补助累计 = Val(txt普通门诊公务员补助累计.Text)
        .m_超大额限额公务员补助 = Val(txt超大额限额公务员补助.Text)
        .m_特殊结算方式 = txt特殊结算方式.Text
        .m_特殊结算说明 = txt特殊结算说明.Text
    End With
    Unload Me
End Sub

Private Sub cmd超大额限额公务员补助_Click()
    txt超大额限额公务员补助.Text = Format(Val(txt超大额限额公务员补助.Text), "0.00")
    cmd超大额限额公务员补助.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd超大额限额公务员补助.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt超大额限额公务员补助.SetFocus
    txt超大额限额公务员补助.SelStart = 0
    txt超大额限额公务员补助.SelLength = Len(txt超大额限额公务员补助.Text)
End Sub

Private Sub cmd超限自付_Click()
    txt超限自付.Text = Format(Val(txt超限自付.Text), "0.00")
    cmd超限自付.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd超限自付.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt超限自付.SetFocus
    txt超限自付.SelStart = 0
    txt超限自付.SelLength = Len(txt超限自付.Text)
End Sub

Private Sub cmd大病统筹_Click()
    txt大病统筹.Text = Format(Val(txt大病统筹.Text), "0.00")
    cmd大病统筹.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd大病统筹.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt大病统筹.SetFocus
    txt大病统筹.SelStart = 0
    txt大病统筹.SelLength = Len(txt大病统筹.Text)
End Sub

Private Sub cmd大病自付_Click()
    txt大病自付.Text = Format(Val(txt大病自付.Text), "0.00")
    cmd大病自付.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd大病自付.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt大病自付.SetFocus
    txt大病自付.SelStart = 0
    txt大病自付.SelLength = Len(txt大病自付.Text)
End Sub

Private Sub cmd公务员补助_Click()
    txt公务员补助.Text = Format(Val(txt公务员补助.Text), "0.00")
    cmd公务员补助.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd公务员补助.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt公务员补助.SetFocus
    txt公务员补助.SelStart = 0
    txt公务员补助.SelLength = Len(txt公务员补助.Text)
End Sub

Private Sub cmd公务员补助起付标准_Click()
    txt公务员补助起付标准.Text = Format(Val(txt公务员补助起付标准.Text), "0.00")
    cmd公务员补助起付标准.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd公务员补助起付标准.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt公务员补助起付标准.SetFocus
    txt公务员补助起付标准.SelStart = 0
    txt公务员补助起付标准.SelLength = Len(txt公务员补助起付标准.Text)
End Sub

Private Sub cmd公务员补助起付线_Click()
    txt公务员补助起付线.Text = Format(Val(txt公务员补助起付线.Text), "0.00")
    cmd公务员补助起付线.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd公务员补助起付线.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt公务员补助起付线.SetFocus
    txt公务员补助起付线.SelStart = 0
    txt公务员补助起付线.SelLength = Len(txt公务员补助起付线.Text)
End Sub

Private Sub cmd挂钩自付_Click()
    txt挂钩自付.Text = Format(Val(txt挂钩自付.Text), "0.00")
    cmd挂钩自付.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd挂钩自付.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt挂钩自付.SetFocus
    txt挂钩自付.SelStart = 0
    txt挂钩自付.SelLength = Len(txt挂钩自付.Text)
End Sub

Private Sub cmd基数自付_Click()
    txt基数自付.Text = Format(Val(txt基数自付.Text), "0.00")
    cmd基数自付.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd基数自付.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt基数自付.SetFocus
    txt基数自付.SelStart = 0
    txt基数自付.SelLength = Len(txt基数自付.Text)
End Sub

Private Sub cmd结算编号_Click()

    cmd结算编号.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd结算编号.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    
    txt结算编号.SetFocus
    txt结算编号.SelStart = 0
    txt结算编号.SelLength = Len(txt结算编号.Text)
End Sub

Private Sub cmd结算日期_Click()
    txt结算日期.Text = Format(txt结算日期.Text, "yyyy-mm-dd hh:mm:ss")
    cmd结算日期.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd结算日期.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    cmd结算日期.Picture = P1.Picture
    txt结算日期.SetFocus
    txt结算日期.SelStart = 0
    txt结算日期.SelLength = Len(txt结算日期.Text)
End Sub

Private Sub cmd结算总费用_Click()
    txt结算总费用.Text = Format(Val(txt结算总费用.Text), "0.00")
    cmd结算总费用.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd结算总费用.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    cmd结算总费用.Picture = P1.Picture
    txt结算总费用.SetFocus
    txt结算总费用.SelStart = 0
    txt结算总费用.SelLength = Len(txt结算总费用.Text)
End Sub

Private Sub cmd就诊顺序号_Click()
    cmd就诊顺序号.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd就诊顺序号.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt就诊顺序号.SetFocus
    txt就诊顺序号.SelStart = 0
    txt就诊顺序号.SelLength = Len(txt就诊顺序号.Text)
End Sub

Private Sub cmd普通门诊公务员补助累计_Click()
    txt普通门诊公务员补助累计.Text = Format(Val(txt普通门诊公务员补助累计.Text), "0.00")
    cmd普通门诊公务员补助累计.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd普通门诊公务员补助累计.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    cmd普通门诊公务员补助累计.Picture = P1.Picture
    txt普通门诊公务员补助累计.SetFocus
    txt普通门诊公务员补助累计.SelStart = 0
    txt普通门诊公务员补助累计.SelLength = Len(txt普通门诊公务员补助累计.Text)
End Sub

Private Sub cmd起付线_Click()
    txt起付线.Text = Format(Val(txt起付线.Text), "0.00")
    cmd起付线.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd起付线.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt起付线.SetFocus
    txt起付线.SelStart = 0
    txt起付线.SelLength = Len(txt起付线.Text)
End Sub

Private Sub cmd全自付_Click()

    txt全自付.Text = Format(Val(txt全自付.Text), "0.00")
    cmd全自付.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd全自付.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt全自付.SetFocus
    txt全自付.SelStart = 0
    txt全自付.SelLength = Len(txt全自付.Text)

End Sub

Private Sub cmd特殊结算方式_Click()
    cmd特殊结算方式.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd特殊结算方式.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt特殊结算方式.SetFocus
    txt特殊结算方式.SelStart = 0
    txt特殊结算方式.SelLength = Len(txt特殊结算方式.Text)
End Sub

Private Sub cmd特殊结算说明_Click()

    cmd特殊结算说明.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd特殊结算说明.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
 
    cmd特殊结算说明.Picture = P1.Picture
    txt特殊结算说明.SetFocus
    txt特殊结算说明.SelStart = 0
    txt特殊结算说明.SelLength = Len(txt特殊结算说明.Text)
 
End Sub

Private Sub cmd统筹支付_Click()
    txt统筹支付.Text = Format(Val(txt统筹支付.Text), "0.00")
    cmd统筹支付.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd统筹支付.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    cmd统筹支付.Picture = P1.Picture
    txt统筹支付.SetFocus
    txt统筹支付.SelStart = 0
    txt统筹支付.SelLength = Len(txt统筹支付.Text)
End Sub

Private Sub cmd统筹自付_Click()
    txt统筹自付.Text = Format(Val(txt统筹自付.Text), "0.00")
    cmd统筹自付.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd统筹自付.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt统筹自付.SetFocus
    txt统筹自付.SelStart = 0
    txt统筹自付.SelLength = Len(txt统筹自付.Text)
End Sub

Private Sub cmd医保总费用_Click()
    txt医保总费用.Text = Format(Val(txt医保总费用.Text), "0.00")
    cmd医保总费用.Tag = IIf(cmd结算编号.Tag = "1", "2", "1")
    cmd医保总费用.Picture = IIf(cmd结算编号.Tag = "1", P2.Picture, P1.Picture)
    txt医保总费用.SetFocus
    txt医保总费用.SelStart = 0
    txt医保总费用.SelLength = Len(txt医保总费用.Text)
End Sub

Private Sub Form_Load()
    txt卡号.Text = g补充结算.str卡号
    txt医保号.Text = g补充结算.str医保号
    txt病人ID.Text = g补充结算.lng病人ID
    txt主页ID.Text = g补充结算.lng主页ID
    txt姓名.Text = g补充结算.str姓名
    txt住院号.Text = g补充结算.str住院号
    
End Sub

Private Sub txtHIS总费用_LostFocus()
    txtHIS总费用.Text = Format(Val(txtHIS总费用.Text), "0.00")
End Sub

Private Sub txt超大额限额公务员补助_LostFocus()
    txt超大额限额公务员补助.Text = Format(Val(txt超大额限额公务员补助.Text), "0.00")
End Sub

Private Sub txt超限自付_LostFocus()
    txt超限自付.Text = Format(Val(txt超限自付.Text), "0.00")
End Sub

Private Sub txt大病统筹_LostFocus()
    txt大病统筹.Text = Format(Val(txt大病统筹.Text), "0.00")
End Sub

Private Sub txt大病自付_LostFocus()
    txt大病自付.Text = Format(Val(txt大病自付.Text), "0.00")
End Sub

Private Sub txt公务员补助_LostFocus()
    txt公务员补助.Text = Format(Val(txt公务员补助.Text), "0.00")
End Sub

Private Sub txt公务员补助起付标准_LostFocus()
    txt公务员补助起付标准.Text = Format(Val(txt公务员补助起付标准.Text), "0.00")
End Sub

Private Sub txt公务员补助起付线_LostFocus()
    txt公务员补助起付线.Text = Format(Val(txt公务员补助起付线.Text), "0.00")
End Sub

Private Sub txt挂钩自付_LostFocus()
    txt挂钩自付.Text = Format(Val(txt挂钩自付.Text), "0.00")
End Sub

Private Sub txt基数自付_LostFocus()
    txt基数自付.Text = Format(Val(txt基数自付.Text), "0.00")
End Sub

Private Sub txt结算总费用_LostFocus()
    txt结算总费用.Text = Format(Val(txt结算总费用.Text), "0.00")
End Sub

Private Sub txt普通门诊公务员补助累计_LostFocus()
    txt普通门诊公务员补助累计.Text = Format(Val(txt普通门诊公务员补助累计.Text), "0.00")
End Sub

Private Sub txt起付线_LostFocus()
    txt起付线.Text = Format(Val(txt起付线.Text), "0.00")
End Sub

Private Sub txt全自付_LostFocus()
    txt全自付.Text = Format(Val(txt全自付.Text), "0.00")
End Sub

Private Sub txt统筹支付_LostFocus()
    txt统筹支付.Text = Format(Val(txt统筹支付.Text), "0.00")
End Sub

Private Sub txt统筹自付_LostFocus()
    txt统筹自付.Text = Format(Val(txt统筹自付.Text), "0.00")
End Sub

Private Sub txt医保总费用_LostFocus()
    txt医保总费用.Text = Format(Val(txt医保总费用.Text), "0.00")
End Sub

