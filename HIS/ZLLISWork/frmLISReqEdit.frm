VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLISReqEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "检验登记"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9075
   Icon            =   "frmLISReqEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9075
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cbo医生 
      Height          =   300
      Left            =   7125
      TabIndex        =   5
      Text            =   "cbo医生"
      Top             =   435
      Width           =   1380
   End
   Begin VB.ComboBox cbo开单科室 
      Height          =   300
      ItemData        =   "frmLISReqEdit.frx":08CA
      Left            =   1245
      List            =   "frmLISReqEdit.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   450
      Width           =   2145
   End
   Begin VB.PictureBox picAdvice 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   3825
      Left            =   -120
      ScaleHeight     =   3825
      ScaleWidth      =   9195
      TabIndex        =   56
      Top             =   2790
      Width           =   9195
      Begin VB.ComboBox cbo付款方式 
         Height          =   300
         Left            =   6795
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox cbo费别 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1440
         Width           =   2160
      End
      Begin VB.Frame fraSample 
         Caption         =   "标本信息"
         Height          =   1455
         Left            =   180
         TabIndex        =   76
         Top             =   1875
         Width           =   8955
         Begin VB.CheckBox chkEmerge 
            Caption         =   "急"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   3540
            Style           =   1  'Graphical
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   225
            Width           =   420
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   4575
            TabIndex        =   37
            Top             =   630
            Width           =   1320
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   2
            Left            =   4575
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1005
            Width           =   1320
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            Left            =   4575
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   630
            Visible         =   0   'False
            Width           =   1320
         End
         Begin zl9LisWork.VsfGrid vsf2 
            Height          =   1095
            Left            =   120
            TabIndex        =   35
            Top             =   225
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   1931
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   0
            Left            =   7005
            TabIndex        =   39
            Top             =   630
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   88539139
            CurrentDate     =   38222
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   1
            Left            =   7005
            TabIndex        =   43
            Top             =   1005
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   88539139
            CurrentDate     =   38222
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "采样时间(&T)"
            Height          =   180
            Index           =   3
            Left            =   6000
            TabIndex        =   38
            Top             =   690
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "标本形态(&X)"
            Height          =   180
            Index           =   5
            Left            =   3525
            TabIndex        =   36
            Top             =   690
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "检 验 人(&J)"
            Height          =   180
            Index           =   2
            Left            =   3525
            TabIndex        =   40
            Top             =   1065
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "检验时间(&D)"
            Height          =   180
            Index           =   6
            Left            =   6000
            TabIndex        =   42
            Top             =   1065
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "采 样 人(&R)"
            Height          =   180
            Index           =   0
            Left            =   3555
            TabIndex        =   78
            Top             =   690
            Visible         =   0   'False
            Width           =   990
         End
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   7920
         TabIndex        =   75
         Top             =   3420
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   6660
         TabIndex        =   74
         Top             =   3420
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   270
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   3405
         Width           =   1100
      End
      Begin VB.CommandButton cmd采集 
         Height          =   285
         Left            =   6645
         Picture         =   "frmLISReqEdit.frx":08CE
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "选择检验标本"
         Top             =   360
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt采集 
         Height          =   300
         Left            =   4740
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.TextBox txt附加 
         Height          =   300
         Left            =   6735
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chk开始时间 
         BackColor       =   &H80000004&
         Caption         =   "要求时间"
         Height          =   225
         Left            =   315
         TabIndex        =   23
         ToolTipText     =   "是否安排时间"
         Top             =   420
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txt单量 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7230
         MaxLength       =   3
         TabIndex        =   31
         Top             =   1080
         Width           =   1380
      End
      Begin VB.TextBox txt频率 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1350
         TabIndex        =   29
         Top             =   1080
         Width           =   2500
      End
      Begin VB.TextBox txt总量 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4740
         MaxLength       =   3
         TabIndex        =   30
         Top             =   1080
         Width           =   1500
      End
      Begin VB.CheckBox chk紧急 
         BackColor       =   &H80000004&
         Caption         =   "紧急(&J)"
         Height          =   225
         Left            =   7710
         TabIndex        =   27
         Top             =   405
         Width           =   945
      End
      Begin VB.CommandButton cmdExt 
         Height          =   285
         Left            =   8340
         Picture         =   "frmLISReqEdit.frx":09C4
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "选择检验标本"
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   285
         Left            =   5280
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   0
         Width           =   285
      End
      Begin VB.ComboBox cbo执行科室 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "frmLISReqEdit.frx":0ABA
         Left            =   1350
         List            =   "frmLISReqEdit.frx":0ABC
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1440
         Width           =   1995
      End
      Begin VB.TextBox txt医嘱内容 
         Height          =   300
         Left            =   1350
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   0
         Width           =   3945
      End
      Begin VB.TextBox txt医生嘱托 
         Height          =   300
         Left            =   1350
         MaxLength       =   100
         TabIndex        =   28
         Top             =   720
         Width           =   7245
      End
      Begin VB.CommandButton cmd频率 
         Enabled         =   0   'False
         Height          =   240
         Left            =   3575
         Picture         =   "frmLISReqEdit.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(F4)"
         Top             =   1110
         Width           =   270
      End
      Begin MSComCtl2.DTPicker txt开始时间 
         Height          =   300
         Left            =   1350
         TabIndex        =   24
         Top             =   360
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   88539139
         CurrentDate     =   38022
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "付款方式"
         Height          =   420
         Left            =   6345
         TabIndex        =   81
         Top             =   1410
         Width           =   435
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   240
         Left            =   3645
         TabIndex        =   80
         Top             =   1500
         Width           =   480
      End
      Begin VB.Label lbl采集 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "采集方式"
         Height          =   180
         Left            =   3930
         TabIndex        =   68
         Top             =   405
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Line lineTitleSplit 
         BorderColor     =   &H80000000&
         X1              =   400
         X2              =   1440
         Y1              =   320
         Y2              =   320
      End
      Begin VB.Label lbl附加 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "检验标本"
         Height          =   180
         Left            =   5940
         TabIndex        =   67
         Top             =   45
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbl单量 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "每次"
         Height          =   180
         Left            =   6840
         TabIndex        =   66
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lbl单量单位 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   8460
         TabIndex        =   65
         Top             =   1140
         Width           =   15
      End
      Begin VB.Label lbl频率 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "频率"
         Height          =   180
         Left            =   960
         TabIndex        =   64
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lbl总量单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   6150
         TabIndex        =   63
         Top             =   1140
         Width           =   15
      End
      Begin VB.Label lbl总量 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "共"
         Height          =   180
         Left            =   4455
         TabIndex        =   62
         Top             =   1140
         Width           =   180
      End
      Begin VB.Label lbl执行科室 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行科室"
         Height          =   180
         Left            =   600
         TabIndex        =   61
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label lbl医嘱内容 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请项目(&I)"
         Height          =   180
         Left            =   330
         TabIndex        =   60
         Top             =   45
         Width           =   990
      End
      Begin VB.Label lbl开始时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "要求时间"
         Height          =   180
         Left            =   600
         TabIndex        =   59
         Top             =   435
         Width           =   720
      End
      Begin VB.Label lbl医生嘱托 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医生嘱托"
         Height          =   180
         Left            =   585
         TabIndex        =   58
         Top             =   795
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   ">>"
      Height          =   300
      Left            =   8520
      TabIndex        =   72
      TabStop         =   0   'False
      ToolTipText     =   "更多病人信息"
      Top             =   450
      Width           =   315
   End
   Begin VB.TextBox txt姓名 
      Height          =   300
      Left            =   1245
      MaxLength       =   20
      TabIndex        =   1
      ToolTipText     =   "数字为就诊卡号、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号"
      Top             =   60
      Width           =   2160
   End
   Begin VB.TextBox txt年龄 
      Height          =   300
      Left            =   6300
      MaxLength       =   10
      TabIndex        =   3
      Top             =   60
      Width           =   2220
   End
   Begin VB.ComboBox cbo性别 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3990
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   1635
   End
   Begin MSComctlLib.ImageList iLstItem 
      Left            =   8280
      Top             =   3480
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
            Picture         =   "frmLISReqEdit.frx":0BB4
            Key             =   "元素"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstbrMain 
      Left            =   7680
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":0CC6
            Key             =   "预览"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":0EE2
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":10FE
            Key             =   "修改"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":131A
            Key             =   "删除"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":1536
            Key             =   "Sample"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":1752
            Key             =   "History"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":196E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":1B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":1DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":1FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":21E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":23FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":261A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":2834
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":2A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":31C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":33E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":35FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":3816
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":3A30
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":41AA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":4924
            Key             =   "SpecChar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":4B3E
            Key             =   "toText"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":4D58
            Key             =   "Copy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstbrMainHot 
      Left            =   6360
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":53D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":55F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":5812
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":5A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":5C52
            Key             =   "Sample"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":5E72
            Key             =   "History"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":6092
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":62AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":64CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":66EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":690C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":6B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":6D46
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":6F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":717A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":78F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":7B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":7D28
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":7F42
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":815C
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":88D6
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":9050
            Key             =   "SpecChar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":926A
            Key             =   "toText"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":9484
            Key             =   "Copy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iLstTab 
      Left            =   6960
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":9AFE
            Key             =   "申请"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISReqEdit.frx":A098
            Key             =   "报告"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt门诊号 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1245
      MaxLength       =   10
      TabIndex        =   71
      Top             =   450
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   -30
      TabIndex        =   44
      Top             =   840
      Width           =   9135
      Begin VB.CommandButton cmd单位名称 
         Caption         =   "…"
         Height          =   285
         Left            =   8220
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "热键:F3"
         Top             =   780
         Width           =   285
      End
      Begin VB.CommandButton cmd家庭地址 
         Caption         =   "…"
         Height          =   285
         Left            =   8220
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "热键:F3"
         Top             =   1170
         Width           =   285
      End
      Begin VB.TextBox txt家庭邮编 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7275
         MaxLength       =   6
         TabIndex        =   18
         Top             =   1560
         Width           =   1260
      End
      Begin VB.TextBox txt家庭电话 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5460
         MaxLength       =   20
         TabIndex        =   17
         Top             =   1560
         Width           =   1260
      End
      Begin VB.TextBox txt家庭地址 
         Height          =   300
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1170
         Width           =   6945
      End
      Begin VB.TextBox txt单位邮编 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3315
         MaxLength       =   6
         TabIndex        =   16
         Top             =   1560
         Width           =   1260
      End
      Begin VB.TextBox txt单位电话 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   15
         Top             =   1560
         Width           =   1260
      End
      Begin VB.TextBox txt单位名称 
         Height          =   300
         Left            =   1260
         MaxLength       =   100
         TabIndex        =   11
         Top             =   780
         Width           =   6945
      End
      Begin VB.TextBox txt身份证号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1260
         MaxLength       =   18
         TabIndex        =   10
         Top             =   390
         Width           =   7245
      End
      Begin VB.ComboBox cbo职业 
         Height          =   300
         Left            =   7275
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   0
         Width           =   1260
      End
      Begin VB.ComboBox cbo婚姻 
         Height          =   300
         Left            =   5460
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   0
         Width           =   1260
      End
      Begin VB.ComboBox cbo民族 
         Height          =   300
         Left            =   3315
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   0
         Width           =   1260
      End
      Begin VB.ComboBox cbo国籍 
         Height          =   300
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "邮编"
         Height          =   180
         Left            =   6825
         TabIndex        =   55
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭电话"
         Height          =   240
         Left            =   4680
         TabIndex        =   54
         Top             =   1620
         Width           =   840
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址"
         Height          =   240
         Left            =   480
         TabIndex        =   53
         Top             =   1230
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "邮编"
         Height          =   180
         Left            =   2865
         TabIndex        =   52
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位电话"
         Height          =   240
         Left            =   480
         TabIndex        =   51
         Top             =   1620
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称"
         Height          =   240
         Left            =   480
         TabIndex        =   50
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   240
         Left            =   480
         TabIndex        =   49
         Top             =   450
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍"
         Height          =   240
         Left            =   825
         TabIndex        =   48
         Top             =   60
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         Height          =   240
         Left            =   2865
         TabIndex        =   47
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   240
         Left            =   6825
         TabIndex        =   46
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   240
         Left            =   4680
         TabIndex        =   45
         Top             =   60
         Width           =   840
      End
   End
   Begin VB.Label lbl开嘱医生 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "医生"
      Height          =   180
      Left            =   6735
      TabIndex        =   83
      Top             =   495
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "申请科室"
      Height          =   180
      Left            =   435
      TabIndex        =   82
      Top             =   510
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名(&N)"
      Height          =   180
      Left            =   570
      TabIndex        =   0
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "年龄"
      Height          =   240
      Left            =   5850
      TabIndex        =   70
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      Height          =   240
      Left            =   3525
      TabIndex        =   69
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmLISReqEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public strPrivs As String       '用户具有本程序的具体权限
Private blnOK As Boolean

Private FileID As String
Private PatientID As String '病人ID
Private CheckID As String '病案ID或挂号单ID
Private PatientType As Integer '1=门诊病人 2=住院病人
Private FileTypeID As String '病历模板文件ID
Private bSample As Boolean '是否示范
Private bln护士站 As Boolean
Private ParentForm As Object
Private DeptID As Long '开单科室
Private ItemType As Integer  '申请项目类别 1=PACS 2=LIS
Private ItemDeptID As Long '项目执行科室

Private PatientDate As Date '病人就诊或入院时间
Private AdviceID As Long, SendNO As Long '医嘱ID、发送号
Private sCheckNo As String '发送单据号
Private iRecordType As Integer '记录性质
Private alngFileID(1) As Long '申请和报告ID
Private intType As Integer '诊疗类别:-1=其他、0=检查组合、1=手术、2=中药、4=检验
Private iTabIndex As Integer
Private mlng前提ID As Long, bln医技执行 As Boolean

'医嘱编辑
Private strAdviceText As String '医嘱内容
Private str类别 As String, lngClinicID As Long, strClinicName As String, str标本部位 As String
Private strSequence As String, lng频率次数 As Long, lng频率间隔 As Long, str间隔单位 As String '频率
Private int计价特性 As Integer, int执行性质 As Integer, lng病人科室ID As Long
Private mstr性别 As String
Private mstrLike As String
Private gint过敏登记有效天数 As Integer
Private rsRelativeAdvice As ADODB.Recordset '相关医嘱
Private strExtData As String '附加项目

Private ifInitItem As Boolean '是否在进入申请时直接显示申请项目

Private iInputType As Integer
'病人姓名当前输入状态，如果一直以该状态可以不输入前导符
'0：就诊卡
'1：病人ID
'2：住院号
'3：门诊号
'4：挂号单
'5：收费单据号
'6：姓名

Private mlngDefaultDevice As Long '默认的检验仪器ID
Private blnComm As Boolean '是否允许双向通讯
Private mbln微生物项目 As Boolean
Private objLISComm As Object
Private mblnSample As Boolean
Private mlngNoneHomeKey() As Long
Private mblnContiAdd As Boolean '是否连续输入
Private blnEmerge As Boolean '是否区分急诊标本
Private mstrCurrentNO As String '当前手工标本号

Private Declare Function GetParent Lib "user32" (ByVal Hwnd As Long) As Long

Public Function ShowMe_Request(frmParent As Object, ByVal lngDeptID As Long, Optional ByVal iItemType As Integer = 1, Optional ByVal ModalWindow As Boolean = True, Optional ByVal lng前提ID As Long = 0, _
    Optional ByVal blnSample As Boolean = True, Optional ByVal lngDefaultDevice As Long = -1, _
    Optional objComm As Object = Nothing) As Boolean
    
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String '诊疗项目名称
    Dim strDrAdvice As String '医生嘱托
    Dim bAllowEdit As Boolean
    
    On Error Resume Next
    '初始化
    Set rsRelativeAdvice = Nothing
    
    alngFileID(0) = 0
    PatientType = 1: AdviceID = 0: PatientID = 0: CheckID = ""
    mlng前提ID = lng前提ID: ItemType = iItemType: ItemDeptID = lngDeptID
    lngClinicID = 0: strDiagName = "": strDrAdvice = ""
    strExtData = ""
    mblnSample = blnSample
    mblnContiAdd = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "连续登记申请", 0))
    '初始化结束
    
    '获取病人信息
    PatientDate = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
    DeptID = UserInfo.部门ID
    
    '初始输入项
    Me.txt开始时间 = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    '初始医生列表
'    Call Get开嘱医生(0, bln护士站, "", 0, Me.cbo医生, PatientType)
    
    '初始标本核收参数
    If blnSample Then
        mlngDefaultDevice = lngDefaultDevice
    
'        blnComm = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "核收允许双向", 0))
        blnComm = Val(zldatabase.GetPara("核收允许双向", 100, 1208, 0))
'        blnEmerge = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "急诊标本", 0))
        blnEmerge = Val(zldatabase.GetPara("急诊标本", 100, 1208, 0))
        Me.chkEmerge.Visible = blnEmerge
    
        If InitSampleData = False Then
            Exit Function
        End If
    
        '启动仪器数据接收初始化
        Set objLISComm = objComm
    End If
    
    Set ParentForm = frmParent
    
    Call InitForm
    Me.cmdCancel.Caption = IIf(mblnContiAdd, "关闭(&C)", "取消(&C)")
    ifInitItem = True
    
    If ModalWindow Then
        Me.Show vbModal, frmParent
    Else
        Me.Show , frmParent
    End If
    ShowMe_Request = blnOK
End Function

Private Sub ClearData()
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strDiagName As String '诊疗项目名称
    Dim strDrAdvice As String '医生嘱托
    Dim bAllowEdit As Boolean
    
    On Error Resume Next
    '初始化
    alngFileID(0) = 0
    PatientType = 1: AdviceID = 0: PatientID = 0: CheckID = ""
    strDiagName = "": strDrAdvice = ""
    '初始化结束
    
    '获取病人信息
    PatientDate = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    DeptID = UserInfo.部门ID
    
    '获取标本信息
    If mblnSample Then ReadSampleData
    
    '清空屏幕
    Me.txt姓名 = "": Me.txt身份证号 = "": Me.txt单位名称 = "": Me.txt家庭地址 = ""
    Me.txt单位电话 = "": Me.txt单位邮编 = "": Me.txt家庭电话 = "": Me.txt家庭邮编 = ""

    Me.txt开始时间 = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
End Sub

Private Function InitSampleData() As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能：初始核收参数
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset, strSQL As String
    
    On Error GoTo ErrHand
    
    With vsf2
        .Cols = 0
        .NewColumn "", 255, 4
'        .NewColumn "检验仪器", 2000, 1, "...", 1
        .NewColumn "检验仪器", 2000, 1, , 0
        .NewColumn "标本号", 800, 1, , 1, GetMaxLength("病人医嘱发送", "样本条码")
        .NewColumn "", 0, 1
        .NewColumn "", 0, 1
        .NoDouble = True
        .FixedCols = 1
    End With
    
        
    dtp(1).Value = Format(zldatabase.Currentdate, dtp(1).CustomFormat)
    dtp(0).Value = dtp(1).Value
    
    strSQL = "SELECT 名称,0 AS ID FROM 检验标本形态"
    OpenRecord rs, strSQL, Me.Caption
    If rs.BOF = False Then Call AddComboData(cbo(0), rs)
    
    InitSampleData = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ReadSampleData() As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能：读取标本信息
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset, mstrSQL As String
    Dim strSubQry As String, i As Integer, aAdvices() As String

    On Error GoTo ErrHand

    Call ResetVsf(vsf2)
    Call LoadDefaultSample(True)

    '2.读取可选的检验人员
    cbo(1).Clear
    cbo(2).Clear

    mstrSQL = "SELECT A.姓名 AS 名称,A.ID,DECODE(A.ID," & UserInfo.ID & ",1,0) AS 缺省 " & _
                "FROM 人员表 A,部门人员 B " & _
                "WHERE A.ID=B.人员id AND B.部门id=[1] "
    Set rs = zldatabase.OpenSQLRecord(mstrSQL, Me.Caption, ItemDeptID)

    If Not rs.EOF Then
        Call AddComboData(cbo(1), rs, False)
        Call AddComboData(cbo(2), rs, False)
    End If

    If cbo(2).ListIndex = -1 And cbo(2).ListCount > 0 Then cbo(2).ListIndex = 0
    If cbo(1).ListIndex = -1 And cbo(1).ListCount > 0 Then cbo(1).ListIndex = 0

    ReadSampleData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub LoadDefaultSample(Optional ByVal blnAll As Boolean = False)
    '--------------------------------------------------------------------------------------------------------
    '功能:
    'blnAll：是否一次增加所有标本
    '--------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strNO As String
    Dim lngDefaultRec As Long
    Dim strConnectDevIDs As String
    Dim strSubQry As String, i As Integer, aAdvices() As String
    Dim mRs As New ADODB.Recordset, mstrSQL As String, rs As New ADODB.Recordset, aItems() As Variant, lngIndex As Long
    Dim mlngLoop As Long
    Dim strItems As String, lngTmpNO As Long
    
    strSubQry = ""
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        strSubQry = strSubQry & " Union All " & "Select " & rsRelativeAdvice("ID") & " As ID From Dual"
        
        rsRelativeAdvice.MoveNext
    Loop
    If Len(strSubQry) > 0 Then strSubQry = Mid(strSubQry, 12)
    rsRelativeAdvice.MoveFirst
    
    '读取申请的检验项目(检验指标)
    mstrSQL = "SELECT ID,检验项目,缩写,项目类别,结果类型," & _
        "TRIM(仪器1||' '||仪器2||' '||仪器3||' '||仪器4||' '||仪器5) AS 仪器,ROWNUM AS 序号,0 As 选择 " & _
        "FROM " & _
        "(SELECT C.ID,C.中文名 AS 检验项目,D.缩写,D.项目类别,D.结果类型," & _
        " Max(Decode(Mod(Rownum,5),0,E.仪器ID,'')) As 仪器1," & _
        " Max(Decode(Mod(Rownum,5),1,E.仪器ID,'')) As 仪器2," & _
        " Max(Decode(Mod(Rownum,5),2,E.仪器ID,'')) As 仪器3," & _
        " Max(Decode(Mod(Rownum,5),3,E.仪器ID,'')) As 仪器4," & _
        " Max(Decode(Mod(Rownum,5),4,E.仪器ID,'')) As 仪器5 " & _
        " FROM 检验报告项目 B,诊治所见项目 C,检验项目 D,检验仪器项目 E,(" & strSubQry & ") S " & _
        " WHERE B.诊疗项目ID=S.ID " & _
            "AND B.报告项目ID=C.ID " & _
            "AND D.诊治项目ID=C.ID AND B.报告项目ID=E.项目ID(+)" & _
        " GROUP BY C.ID,C.中文名,D.缩写,D.项目类别,D.结果类型)"
    Call OpenRecord(rs, mstrSQL, Me.Caption)
    mbln微生物项目 = False: vsf2.Tag = ""
    If rs.BOF = False Then
        mbln微生物项目 = (zlCommFun.Nvl(rs("项目类别"), 0) = 2)
        vsf2.Tag = rs.RecordCount
        aItems = rs.GetRows
    Else
        aItems = Array()
    End If

    '获取本机连接的检验仪器
    strConnectDevIDs = GetConnectDevs
    On Error GoTo ErrHand

    '读取相应的检验仪器列表
    mstrSQL = "SELECT DISTINCT NVL(E.ID,-1) AS ID,NVL(E.名称,'[手工]') AS 名称,NVL(D.缺省仪器,-1) AS 缺省仪器 " & _
                    "FROM 检验报告项目 B, 检验仪器项目 D, 检验仪器 E,(" & strSubQry & ") S " & _
                    "Where B.诊疗项目ID=S.ID " & _
                    "AND B.报告项目ID = D.项目id(+) AND D.仪器id = E.ID(+)" & _
                    "ORDER BY NVL(D.缺省仪器,-1)  DESC"
                      
    Call OpenRecord(mRs, mstrSQL, Me.Caption)
    If mRs.BOF = False Then
        '如果一次增加所有标本，则增加N条空记录
        If blnAll Then vsf2.Rows = mRs.RecordCount + 1
        
        For lngLoop = 1 To vsf2.Rows - 1
            If Val(vsf2.RowData(lngLoop)) = 0 Then
                
                '检验仪器是否已经使用,如已使用,则取一个仪器,如没有下一个,则取最后一个
                lngDefaultRec = -1: mRs.MoveFirst
                Do While Not mRs.EOF
                    If CheckHave(zlCommFun.Nvl(mRs("ID"), 0)) = False Then
                        If zlCommFun.Nvl(mRs("ID"), 0) = mlngDefaultDevice Then
                            '取过滤条件指定的检验仪器
                            lngDefaultRec = mRs.AbsolutePosition
                            Exit Do
                        Else
                            If InStr(";" & strConnectDevIDs & ";", ";" & zlCommFun.Nvl(mRs("ID"), 0) & ";") > 0 Then
                                '默认取本机连接的检验仪器
                                lngDefaultRec = mRs.AbsolutePosition
'                                Exit Do
                            Else
                                If lngDefaultRec = -1 Then lngDefaultRec = mRs.AbsolutePosition
                            End If
                        End If
                    End If
                    mRs.MoveNext
                Loop
                If lngDefaultRec = -1 Then
                    mRs.MoveLast
                Else
                    mRs.AbsolutePosition = lngDefaultRec
                End If
                
                vsf2.TextMatrix(lngLoop, 1) = zlCommFun.Nvl(mRs("名称"))
                vsf2.RowData(lngLoop) = zlCommFun.Nvl(mRs("ID"), 0)
                
                '产生本仪器在本日的下一标本号
                strNO = ""
                For mlngLoop = 1 To vsf2.Rows - 1
                    If mlngLoop <> lngLoop Then
                        If Val(vsf2.RowData(lngLoop)) = Val(vsf2.RowData(mlngLoop)) Then
                            '已有此仪器
                            If Val(strNO) < Val(vsf2.TextMatrix(mlngLoop, 2)) Then
                                strNO = Val(vsf2.TextMatrix(mlngLoop, 2))
                            End If
                        End If
                    End If
                Next
                
                If strNO = "" Then
                    If vsf2.RowData(lngLoop) <> -1 Then
                        vsf2.TextMatrix(lngLoop, 2) = CalcNextSampleNO(zlCommFun.Nvl(mRs("ID"), 0), lngLoop, _
                            IIf(blnEmerge And Me.chkEmerge.Value = 1, 1, 0))
                    Else
                        '手工，从当前号累加标本号
                        If Len(mstrCurrentNO) = 0 Then
                            '取初始标本号
                            lngTmpNO = Val(CalcNextSampleNO(zlCommFun.Nvl(mRs("ID"), 0), lngLoop, _
                                IIf(blnEmerge And Me.chkEmerge.Value = 1, 1, 0)))
                            If gblnManualPH Then
                                vsf2.TextMatrix(lngLoop, 2) = TransSampleNO_PH(lngTmpNO, vsf2.RowData(lngLoop))
                            Else
                                vsf2.TextMatrix(lngLoop, 2) = lngTmpNO
                            End If
                        Else
                            '从当前号累加
                            If gblnManualPH Then
                                If Val(Split(mstrCurrentNO, "-")(1)) = gintNumberPH Then
                                    vsf2.TextMatrix(lngLoop, 2) = Format(Val(Split(mstrCurrentNO, "-")(0)) + 1, "000") & "-0001"
                                Else
                                    vsf2.TextMatrix(lngLoop, 2) = Format(Val(Split(mstrCurrentNO, "-")(0)), "000") & "-" & _
                                        Format(Val(Split(mstrCurrentNO, "-")(1)) + 1, "0000")
                                End If
                            Else
                                vsf2.TextMatrix(lngLoop, 2) = Val(mstrCurrentNO) + 1
                            End If
                        End If
                    End If
                Else
                    vsf2.TextMatrix(lngLoop, 2) = Val(strNO) + 1
                End If
                
                '将项目添加到标本中
                If Val(vsf2.Tag) > 0 Then
                    strItems = ""
                    
                    If UBound(aItems) > -1 Then
                        For i = 0 To UBound(aItems, 2)
                            If InStr(" " & IIf(aItems(5, i) = "" Or IsNull(aItems(5, i)), "-1", aItems(5, i)) & " ", " " & vsf2.RowData(lngLoop) & " ") > 0 _
                                And aItems(7, i) = 0 Then
                                strItems = strItems & "|" & aItems(0, i) & "^" & aItems(1, i) & "^" & aItems(2, i) & "^" & aItems(5, i) & "^^" & aItems(4, i)
                                aItems(7, i) = 1
                            End If
                        Next
                    End If
                    If strItems <> "" Then vsf2.TextMatrix(lngLoop, 3) = Mid(strItems, 2)
                End If
            End If
        Next
        For lngLoop = 1 To vsf2.Rows - 1
            '如果该标本没有指标，则删除
            If vsf2.TextMatrix(lngLoop, 3) = "" And vsf2.Rows > 2 Then
                vsf2.RemoveItem lngLoop
                lngLoop = lngLoop - 1
                If lngLoop = vsf2.Rows - 1 Then Exit For
            End If
        Next
        vsf2.EditMode(1) = 1
        vsf2.ComboList(1) = "..."
    Else
        vsf2.TextMatrix(vsf2.Row, 1) = "[未指明仪器]"
        vsf2.EditMode(1) = 0
        vsf2.ComboList(1) = ""
        
        '产生本仪器在本日的下一标本号
        strNO = ""
        For mlngLoop = 1 To vsf2.Rows - 1
            If mlngLoop <> vsf2.Row Then
                If Val(vsf2.RowData(vsf2.Row)) = Val(vsf2.RowData(mlngLoop)) Then
                    '已有此仪器
                    If Val(strNO) < Val(vsf2.TextMatrix(mlngLoop, 2)) Then
                        strNO = Val(vsf2.TextMatrix(mlngLoop, 2))
                    End If
                End If
            End If
        Next
        
        If strNO = "" Then
            vsf2.TextMatrix(vsf2.Row, 2) = CalcNextSampleNO(0, vsf2.Row, IIf(blnEmerge And Me.chkEmerge.Value = 1, 1, 0))
        Else
            vsf2.TextMatrix(vsf2.Row, 2) = Val(strNO) + 1
        End If
        
    End If
    
    Exit Sub
    
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:检验是否已经使用过
    '参数:
    '返回:
    '--------------------------------------------------------------------------------------------------------
    Dim mlngLoop As Long
    For mlngLoop = 1 To vsf2.Rows - 1
        If vsf2.RowData(mlngLoop) = lngKey Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function

Private Function CalcNextSampleNO(ByVal lngKey As Long, ByVal intRow As Integer, ByVal iType As Integer) As String
    '--------------------------------------------------------------------------------------------------------
    '功能:计算指定仪器在当天内的下一个缺省标本号
    '参数:lngKey                检验仪器ID
    '     iType                 标本类别：0=普通、1=急诊
    '返回:缺省标本号码
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset, mstrSQL As String
    Dim strToday As String
    Dim strTmp As String
    Dim lng次数 As Long, mlngLoop As Long
    Dim strLabNo As String, strLabQCNo As String '检验标本、质控标本
    
    '时间,仪器,标本号
    On Error GoTo ErrHand
    
    strToday = Format(zldatabase.Currentdate, "YYYY-MM-DD")
    
    On Error GoTo point1
    
    mstrSQL = "SELECT NVL(MAX(TO_NUMBER(标本序号)),0) AS 最大序号 FROM 检验标本记录 " & _
                "WHERE 核收时间 BETWEEN [2] and [3] " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL ", "AND 仪器id= [1] ") & " And 医嘱ID Is Not Null" & _
                    IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1")
    Set rs = zldatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")))
    
    If Not rs.EOF Then strLabNo = zlCommFun.Nvl(rs("最大序号"))
    
    On Error GoTo ErrHand
    GoTo point2
    
point1:
    On Error GoTo ErrHand
    
    mstrSQL = "SELECT NVL(MAX(标本序号),'') AS 最大序号 FROM 检验标本记录 " & _
                "WHERE 核收时间 BETWEEN [2] and [3] " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL ", "AND 仪器id= [1] ") & " And 医嘱ID Is Not Null" & _
                    IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1")
    Set rs = zldatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")))
    
    If Not rs.EOF Then strLabNo = zlCommFun.Nvl(rs("最大序号"))
    
point2:
    On Error GoTo point3
    
    mstrSQL = "SELECT NVL(MAX(TO_NUMBER(标本序号)),0) AS 最大序号 FROM 检验标本记录 " & _
                "WHERE 核收时间 BETWEEN [2] and [3] " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL ", "AND 仪器id= [1] ") & _
                    IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1")
    Set rs = zldatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")))
    
    If Not rs.EOF Then strLabQCNo = zlCommFun.Nvl(rs("最大序号"))
    
    On Error GoTo ErrHand
    GoTo point4
    
point3:
    On Error GoTo ErrHand
    
    mstrSQL = "SELECT NVL(MAX(标本序号),'') AS 最大序号 FROM 检验标本记录 " & _
                "WHERE 核收时间 BETWEEN [2] and [3] " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL ", "AND 仪器id=[1] ") & _
                    IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1")
    Set rs = zldatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")))
    
    If Not rs.EOF Then strLabQCNo = zlCommFun.Nvl(rs("最大序号"))
    
point4:
    CalcNextSampleNO = strLabQCNo
    If Val(strLabQCNo) > Val(strLabNo) + 100 Then CalcNextSampleNO = strLabNo
    
'    CalcNextSampleNO = zlCommFun.NVL(rs("最大序号"))
'
    For mlngLoop = 1 To vsf2.Rows - 1
        If mlngLoop <> intRow Then
            If Val(vsf2.RowData(mlngLoop)) = lngKey Then
                If Val(CalcNextSampleNO) < Val(vsf2.TextMatrix(mlngLoop, 2)) Then
                    CalcNextSampleNO = Val(vsf2.TextMatrix(mlngLoop, 2))
                End If
            End If
        End If
    Next
    
    If Val(CalcNextSampleNO) = 0 Then
        CalcNextSampleNO = "1"
        Exit Function
    End If
    
    '1.为纯数字的情况处理
'        If CheckStrType(CalcNextSampleNO, 1) Then
    
        '是为纯数字,则直接加1
        CalcNextSampleNO = Val(CalcNextSampleNO) + 1
        Exit Function
        
'        End If
    
    '2.有字符的情况处理
'        lng次数 = 0
'        strTmp = ""
'        For mlngLoop = Len(CalcNextSampleNO) To 1 Step -1
'            If Mid(CalcNextSampleNO, mlngLoop, 1) >= "0" And Mid(CalcNextSampleNO, mlngLoop, 1) <= "9" Then
'                strTmp = Mid(CalcNextSampleNO, mlngLoop, 1) & strTmp
'
'                If mlngLoop = Len(CalcNextSampleNO) Then lng次数 = 1
'            Else
'                lng次数 = lng次数 + 1
'            End If
'
'            If lng次数 > 1 And Trim(strTmp) <> "" Then
'                CalcNextSampleNO = Mid(CalcNextSampleNO, 1, mlngLoop) & Str(Val(strTmp) + 1) & Mid(CalcNextSampleNO, mlngLoop + Len(strTmp) + 1)
'                Exit Function
'            End If
'        Next
'
'        If Trim(strTmp) <> "" Then
'            CalcNextSampleNO = strTmp & Mid(CalcNextSampleNO, Len(strTmp) + 1)
'        End If
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub InitForm()
    intType = Switch(ItemType = 1, 0, ItemType = 2, 4)
    Select Case intType
        Case 0
            Me.Caption = "检查登记"
        Case 1
            Me.Caption = "手术登记"
        Case 4
            Me.Caption = "检验登记"
        Case Else
            Me.Caption = "登记"
    End Select

    SetItemFormat
End Sub

Private Sub SetItemFormat()   '根据申请项目决定显示方式
    Select Case intType
        Case 0
            Me.lbl医嘱内容.Caption = "检查项目": Me.lbl附加.Caption = "检查部位": Me.cmdExt.ToolTipText = "选择检查部位"
            Me.lbl附加.Visible = True: Me.txt附加.Visible = True: Me.cmdExt.Visible = True
        Case 1
            Me.lbl医嘱内容.Caption = "手术项目": Me.lbl附加.Caption = "麻醉方式": Me.cmdExt.ToolTipText = "选择麻醉方式"
            Me.lbl附加.Visible = True: Me.txt附加.Visible = True: Me.cmdExt.Visible = True
        Case 4
            Me.lbl医嘱内容.Caption = "检验项目": Me.lbl附加.Caption = "检验标本": Me.cmdExt.ToolTipText = "选择检验标本"
            Me.lbl附加.Visible = True: Me.txt附加.Visible = True: Me.cmdExt.Visible = True
            Me.lbl采集.Visible = True: Me.txt采集.Visible = True: Me.cmd采集.Visible = True
        Case Else
            Me.lbl附加.Visible = False: Me.txt附加.Visible = False: Me.cmdExt.Visible = False
    End Select
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim mlngLoop As Long
    
    If KeyAscii = vbKeyReturn Then
        
        For mlngLoop = 0 To cbo(Index).ListCount - 1
            If Mid(cbo(Index).List(mlngLoop), 1, InStr(cbo(Index).List(mlngLoop), "-") - 1) = cbo(Index).Text Then
                cbo(Index).Text = cbo(Index).List(mlngLoop)
                Exit For
            End If
        Next
        
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub cbo开单科室_Click()
    InitDoctors cbo开单科室.ItemData(cbo开单科室.ListIndex)
End Sub

Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub cbo医生_GotFocus()
    Call zlControl.TxtSelAll(cbo医生)
End Sub

Private Sub chkEmerge_Click()
    Dim lngLoop As Long, mlngLoop As Long
    Dim strNO As String
    
    On Error Resume Next
    
    For lngLoop = 1 To vsf2.Rows - 1
        '产生本仪器在本日的下一标本号
         strNO = ""
         For mlngLoop = 1 To vsf2.Rows - 1
             If mlngLoop <> lngLoop Then
                 If Val(vsf2.RowData(lngLoop)) = Val(vsf2.RowData(mlngLoop)) Then
                     '已有此仪器
                     If Val(strNO) < Val(vsf2.TextMatrix(mlngLoop, 2)) Then
                         strNO = Val(vsf2.TextMatrix(mlngLoop, 2))
                     End If
                 End If
             End If
         Next
         
         If strNO = "" Then
             vsf2.TextMatrix(lngLoop, 2) = CalcNextSampleNO(Val(vsf2.RowData(lngLoop)), lngLoop, _
                 IIf(blnEmerge And Me.chkEmerge.Value = 1, 1, 0))
         Else
             vsf2.TextMatrix(lngLoop, 2) = Val(strNO) + 1
         End If
    Next

    vsf2.Col = 2
    vsf2.ShowCell vsf2.Row, vsf2.Col
    vsf2.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdMore_Click()
    Me.Frame1.Visible = Not Me.Frame1.Visible
    If Me.Frame1.Visible Then
        Me.cbo国籍.SetFocus
    Else
        Me.txt医嘱内容.SetFocus
    End If
    Me.Height = Me.Height + IIf(Me.Frame1.Visible, 1, -1) * Me.Frame1.Height
    
    Form_Resize
End Sub

Private Sub CmdOk_Click()
    If Len(sCheckNo) > 0 Then
        If MsgBox("当前申请项目将与收费单据：" & sCheckNo & " 关联，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    If SaveFile Then blnOK = True
    
    If mblnContiAdd Then
        ClearData
        Me.txt姓名.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Sub cmd采集_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strItemID As String
    
    If Len(strExtData) > 0 Then
        strItemID = Split(strExtData, ";")(0)
        If Len(strItemID) > 0 Then strItemID = Split(strItemID, ",")(0)
    End If
    Set rsTmp = SelectCap(Val(strItemID))
    Me.txt采集.SetFocus
    If Not rsTmp Is Nothing Then
        Me.cmd采集.Tag = rsTmp("ID")
        Me.txt采集 = rsTmp("名称"): Me.txt采集.Tag = Me.txt采集
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    If picAdvice.Enabled And ifInitItem Then
        Me.txt医嘱内容 = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "申请项目", "")
        Me.txt附加 = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "申请附加", "")
        
        '默认为最近一次的项目
        If Len(Trim(Me.txt医嘱内容)) > 0 Then
            On Error Resume Next
            Call txt医嘱内容_KeyPress(vbKeyReturn)
            
            Me.txt姓名.SetFocus
        End If
        
        ifInitItem = False
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnShowDetail As Boolean
    
    On Error GoTo errH
    
    blnShowDetail = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "病人详细信息", "False")
    Me.Height = Me.Height - IIf(blnShowDetail, 0, Me.Frame1.Height)
    Me.Frame1.Visible = blnShowDetail
    
    blnOK = False
    iInputType = -1
    '有关医嘱的参数
    mstrLike = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    
    strSQL = "Select 参数号,参数名,参数值 from 系统参数表"
    Call OpenRecord(rsTmp, strSQL, "mdlCISCore")
    '皮试结果有效时间
    rsTmp.Filter = "参数号=2"
    If Not rsTmp.EOF Then gint过敏登记有效天数 = Val(Nvl(rsTmp!参数值, 0))
    
    '---------权限控制-------------
    'strPrivs = gstrPrivs
    '初始病人信息
    lng病人科室ID = UserInfo.部门ID
    Call InitData
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    Dim lngTxtWidth As Single
    Dim lngDistance As Single
    
    If WindowState = 1 Then Exit Sub
    lngTools = 0
    lngStatus = 0
    lngDistance = 300
    
    On Error Resume Next
    With picAdvice
        .Width = Me.ScaleWidth
    End With
    With Me.chk紧急
        .Left = picAdvice.Width - Me.lbl开始时间.Left - .Width
        If .Left < Me.txt采集.Left + Me.txt采集.Width + lngDistance Then .Left = Me.txt采集.Left + Me.txt采集.Width + lngDistance
    End With
'    With Me.chk紧急
'        .Left = picAdvice.Width - Me.lbl开始时间.Left - .Width
'        If .Left < Me.txt开始时间.Left + Me.txt开始时间.Width + lngDistance Then .Left = Me.txt开始时间.Left + Me.txt开始时间.Width + lngDistance
'    End With
    
    lngTxtWidth = (picAdvice.ScaleWidth - Me.lbl开始时间.Left - Me.cmdSel.Width - Me.txt医嘱内容.Left - lngDistance - _
        Me.lbl附加.Width - Me.cmdExt.Width - 60) / 2
    With Me.txt医嘱内容
        .Width = lngTxtWidth
        Me.cmdSel.Left = .Left + .Width
        Me.lbl附加.Left = Me.cmdSel.Left + Me.cmdSel.Width + lngDistance
    End With
    With Me.txt附加
        .Left = Me.lbl附加.Left + Me.lbl附加.Width + 30
        .Width = lngTxtWidth
        Me.cmdExt.Left = .Left + .Width
    End With
    Me.lineTitleSplit.X2 = Me.cmdExt.Left + Me.cmdExt.Width + 200

    With Me.txt医生嘱托
        .Width = picAdvice.Width - Me.lbl开始时间.Left - .Left
    End With
    
    lngTxtWidth = (picAdvice.Width - Me.lbl开始时间.Left - Me.txt频率.Left - Me.txt频率.Width - _
        (Me.lbl总量单位.Width + Me.lbl总量.Width + lngDistance + 2 * 30) - _
        (Me.lbl单量单位.Width + Me.lbl单量.Width + lngDistance + 2 * 30)) / 2
    If lngTxtWidth < 1000 Then lngTxtWidth = 1000
    Me.lbl总量.Left = Me.txt频率.Left + Me.txt频率.Width + lngDistance
    With Me.txt总量
        .Left = Me.lbl总量.Left + Me.lbl总量.Width + 30
        .Width = lngTxtWidth
    End With
    Me.lbl总量单位.Left = Me.txt总量.Left + Me.txt总量.Width + 30
    Me.lbl单量.Left = Me.lbl总量单位.Left + Me.lbl总量单位.Width + lngDistance
    With Me.txt单量
        .Left = Me.lbl单量.Left + Me.lbl单量.Width + 30
        .Width = lngTxtWidth
    End With
    Me.lbl单量单位.Left = Me.txt单量.Left + Me.txt单量.Width + 30
    
    With Me.cbo医生
        .Left = Me.txt单量.Left
'        .Width = picAdvice.Width - Me.lbl开始时间.Left - .Left
    End With
    Me.lbl开嘱医生.Left = Me.cbo医生.Left - Me.lbl开嘱医生.Width

    Me.picAdvice.Top = Me.Frame1.Top + IIf(Me.Frame1.Visible, Me.Frame1.Height, 0)
    
    With Me.cmdMore
        .Caption = IIf(Me.Frame1.Visible, "<<", ">>")
        .ToolTipText = IIf(Me.Frame1.Visible, "基本病人信息", "详细病人信息")
    End With
    
    If Not mblnSample And Me.fraSample.Visible Then
        Me.fraSample.Visible = False
        
        Me.cmdHelp.Top = Me.cmdHelp.Top - Me.fraSample.Height
        Me.cmdCancel.Top = Me.cmdCancel.Top - Me.fraSample.Height
        Me.cmdOK.Top = Me.cmdOK.Top - Me.fraSample.Height
        Me.Height = Me.Height - Me.fraSample.Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlCommFun.OpenIme False
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "病人详细信息", Me.Frame1.Visible
    '保存最近的申请项目
    If Len(Trim(Me.txt医嘱内容)) > 0 Then
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "申请项目", Me.txt医嘱内容
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "申请附加", Me.txt附加
    End If
End Sub

Private Function SaveFile() As Boolean
    Dim sTmpFileID As String
    
    SaveFile = False
        
    '保存申请
    
    If Not ValidAdvice Then Exit Function
    If Not SaveAdvice Then Exit Function

    SaveFile = True
End Function

'检查医嘱内容的合法性
Private Function ValidAdvice() As Boolean
    ValidAdvice = True
    
    On Error Resume Next
'    If txt门诊号.Text = "" Then
'        ValidAdvice = False
'        MsgBox "请输入病人的门诊号！", vbInformation, gstrSysName
'        txt门诊号.SetFocus: Exit Function
'    End If
    If cbo费别.ListIndex = -1 Then
        ValidAdvice = False
        MsgBox "请选择病人的费别！", vbInformation, gstrSysName
        cbo费别.SetFocus: Exit Function
    End If
    If txt姓名.Text = "" Then
        ValidAdvice = False
        MsgBox "请输入病人的姓名！", vbInformation, gstrSysName
        txt姓名.SetFocus: Exit Function
    End If
    
    If Len(Trim(strAdviceText)) = 0 Then
        ValidAdvice = False
        MsgBox "必须输入申请项目！", vbInformation, gstrSysName
        Me.txt医嘱内容.SetFocus: Exit Function
    End If
    If Len(Trim(strSequence)) = 0 Then
        ValidAdvice = False
        MsgBox "必须指定频率！", vbInformation, gstrSysName
        Me.txt频率.SetFocus: Exit Function
    End If
    If Not Check开始时间(CStr(Me.txt开始时间)) Then
        ValidAdvice = False
        Me.txt开始时间.SetFocus: Exit Function
    End If
    If Len(Trim(Me.txt总量)) = 0 Then
        ValidAdvice = False
        MsgBox "请输入总量！", vbInformation, gstrSysName
        Me.txt总量.SetFocus: Exit Function
    End If
    If Len(Trim(Me.txt单量)) = 0 And Me.txt单量.Enabled Then
        ValidAdvice = False
        MsgBox "请输入单量！", vbInformation, gstrSysName
        Me.txt单量.SetFocus: Exit Function
    End If
    If Val(Me.txt单量) > Val(Me.txt总量) Then
        ValidAdvice = False
        MsgBox "单量不能大于总量！", vbInformation, gstrSysName
        Me.txt总量.SetFocus: Exit Function
    End If
    If Me.cbo开单科室.ListIndex = -1 Then
        ValidAdvice = False
        MsgBox "请指定开单科室！", vbInformation, gstrSysName
        Me.cbo开单科室.SetFocus: Exit Function
    End If
'    If Me.cbo医生.ListIndex = -1 Then
'        ValidAdvice = False
'        MsgBox "请指定开单医生！", vbInformation, gstrSysName
'        Me.cbo医生.SetFocus: Exit Function
'    End If
    
    If mblnSample Then
        If Not ValidSampleData(IIf(blnEmerge And Me.chkEmerge.Value = 1, 1, 0)) Then
            ValidAdvice = False: Exit Function
        End If
    End If
End Function
'保存医嘱
Private Function SaveAdvice() As Boolean
    On Error GoTo DBError
    SaveAdvice = True
    
    SaveAdviceData
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    SaveAdvice = False
    SaveErrLog
End Function

Private Sub SaveAdviceData()
    Dim strSQL As String, strDate As String, strNO As String
    Dim lngAdviceID As Long, lngTmpID As Long, lngSendNO As Long
    Dim iMaxSeq As Integer, iSendSeq As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim lng开嘱科室ID As Long, lng病人id As Long, strDoctor As String, i As Integer
    Dim str执行科室ID As String, str执行科室ID1 As String, lngDept As Long
    Dim rsCard As ADODB.Recordset
    Dim tmpstr类别 As String, tmplngClinicID As Long, tmpint计价特性 As Integer, tmpint执行性质 As Integer
    Dim rsDept As ADODB.Recordset

    gcnOracle.BeginTrans
    On Error GoTo DBError
    
    '保存病人信息
    strDate = "To_Date('" & Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    If PatientType = 1 Then '门诊病人
        If PatientID > 0 Then '已有的病人
            lng病人id = PatientID
            strSQL = _
                "zl_挂号病人病案_INSERT(3," & lng病人id & "," & IIf(Len(Trim(txt门诊号.Text)) = 0, "Null", txt门诊号.Text) & "," & _
                "'',''," & _
                "'" & txt姓名.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & "'," & _
                "'" & NeedName(cbo费别.Text) & "','" & NeedName(cbo付款方式.Text) & "'," & _
                "'" & NeedName(cbo国籍.Text) & "','" & NeedName(cbo民族.Text) & "','" & NeedName(cbo婚姻.Text) & "'," & _
                "'" & NeedName(cbo职业.Text) & "','" & txt身份证号.Text & "','" & txt单位名称.Text & "'," & _
                Val(txt单位名称.Tag) & ",'" & txt单位电话.Text & "','" & txt单位邮编.Text & "','" & txt家庭地址.Text & "'," & _
                "'" & txt家庭电话.Text & "','" & txt家庭邮编.Text & "'," & strDate & ",NULL)"
        Else '新病人
            lng病人id = zldatabase.GetNextNo(1)
            strSQL = _
                "zl_挂号病人病案_INSERT(1," & lng病人id & "," & IIf(Len(Trim(txt门诊号.Text)) = 0, "Null", txt门诊号.Text) & "," & _
                "'',''," & _
                "'" & txt姓名.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & "'," & _
                "'" & NeedName(cbo费别.Text) & "','" & NeedName(cbo付款方式.Text) & "'," & _
                "'" & NeedName(cbo国籍.Text) & "','" & NeedName(cbo民族.Text) & "','" & NeedName(cbo婚姻.Text) & "'," & _
                "'" & NeedName(cbo职业.Text) & "','" & txt身份证号.Text & "','" & txt单位名称.Text & "'," & _
                Val(txt单位名称.Tag) & ",'" & txt单位电话.Text & "','" & txt单位邮编.Text & "','" & txt家庭地址.Text & "'," & _
                "'" & txt家庭电话.Text & "','" & txt家庭邮编.Text & "'," & strDate & ",NULL)"
        End If
        Call SQLTest(App.ProductName, Me.Caption, strSQL)
        gcnOracle.Execute strSQL, , adCmdStoredProc
        Call SQLTest
    Else
        lng病人id = PatientID
    End If
    '保存医嘱并发送
    lngAdviceID = zldatabase.GetNextId("病人医嘱记录")
    iMaxSeq = 0
    
    lng开嘱科室ID = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex) 'Get开嘱科室ID(Me.cbo医生.ItemData(Me.cbo医生.ListIndex), lng病人科室ID, PatientType)
    lng病人科室ID = lng开嘱科室ID
    
    i = InStr(Me.cbo医生.Text, "-")
    If i > 0 Then
        strDoctor = Trim(Mid(Me.cbo医生.Text, i + 1))
    Else
        strDoctor = Trim(Me.cbo医生.Text)
    End If
    If Len(Me.cbo执行科室.Text) = 0 Then
        str执行科室ID = "NULL"
    Else
        str执行科室ID = Me.cbo执行科室.ItemData(Me.cbo执行科室.ListIndex)
    End If
    
    tmpstr类别 = str类别: tmplngClinicID = lngClinicID: tmpint计价特性 = int计价特性
    tmpint执行性质 = int执行性质
    iSendSeq = 1
    If intType = 4 Then
        '检验项目将采集方式作为主医嘱
        strSQL = "Select * From 诊疗项目目录 Where ID=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Me.cmd采集.Tag)
        tmpstr类别 = rsTmp("类别"): tmplngClinicID = rsTmp("ID"): tmpint计价特性 = Nvl(rsTmp("计价性质"), 0)
        tmpint执行性质 = Nvl(rsTmp("执行科室"), 0)
        '取采集方式的执行部门
        Set rsDept = GetExeDepart(rsTmp("ID"), PatientType + 1, DeptID)
        If rsDept Is Nothing Then
            str执行科室ID1 = "NULL"
        Else
            str执行科室ID1 = rsDept("ID")
        End If
        lngSendNO = zldatabase.GetNextNo(10)
        If Len(sCheckNo) = 0 Then
            strNO = zldatabase.GetNextNo(IIf(PatientType = 1, 13, 14))
        Else
            strNO = sCheckNo
        End If
    End If
    
    If intType <> 4 Then
        iMaxSeq = iMaxSeq + 1
        strSQL = "ZL_病人医嘱记录_Insert(" & lngAdviceID & ",NULL," & _
            iMaxSeq & "," & PatientType & "," & lng病人id & "," & IIf(PatientType = 2, CheckID, "NULL") & "," & _
            "0,1," & _
            "1,'" & tmpstr类别 & "'," & _
            tmplngClinicID & ",NULL,NULL," & _
            IIf(Len(Trim(Me.txt单量)) = 0, "NULL", Me.txt单量) & "," & _
            IIf(Len(Trim(Me.txt总量)) = 0, "NULL", Me.txt总量) & "," & _
            "'" & Replace(strAdviceText, "'", "''") & "','" & Replace(Me.txt医生嘱托, "'", "''") & "'," & _
            "'" & str标本部位 & "','" & strSequence & "'," & _
            IIf(lng频率次数 = 0, "NULL", lng频率次数) & "," & _
            IIf(lng频率间隔 = 0, "NULL", lng频率间隔) & "," & _
            "'" & str间隔单位 & "',NULL," & _
            tmpint计价特性 & "," & _
            str执行科室ID & "," & _
            tmpint执行性质 & "," & Me.chk紧急.Value & "," & _
            IIf(Me.chk开始时间.Visible And Me.chk开始时间.Value = 0, "NULL,", "To_Date('" & Format(Me.txt开始时间.Value, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),") & _
            "NULL," & _
            lng病人科室ID & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
            "Sysdate,'" & IIf(PatientType = 2, "", CheckID) & "'," & _
            IIf(mlng前提ID = 0, "Null", mlng前提ID) & ")"
    
        Call SQLTest(App.ProductName, Me.Caption, strSQL)
        gcnOracle.Execute strSQL, , adCmdStoredProc
        Call SQLTest
        '发送医嘱
        lngSendNO = zldatabase.GetNextNo(10)
        If Len(sCheckNo) = 0 Then
            strNO = zldatabase.GetNextNo(IIf(PatientType = 1, 13, 14))
        Else
            strNO = sCheckNo
        End If
    End If
    '保存相关医嘱
    If Not rsRelativeAdvice Is Nothing Then
        i = 2
        rsRelativeAdvice.MoveFirst
        Do While Not rsRelativeAdvice.EOF
            lngTmpID = zldatabase.GetNextId("病人医嘱记录")
            With rsRelativeAdvice
                strSQL = "ZL_病人医嘱记录_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                    (iMaxSeq + i) & "," & PatientType & "," & lng病人id & "," & IIf(PatientType = 2, CheckID, "NULL") & "," & _
                    "0,1," & _
                    "1,'" & .Fields("类别") & "'," & _
                    .Fields("ID") & ",NULL,NULL," & _
                    IIf(Len(Trim(Me.txt单量)) = 0, "NULL", Me.txt单量) & "," & _
                    IIf(Len(Trim(Me.txt总量)) = 0, "NULL", Me.txt总量) & "," & _
                    "'" & Replace(.Fields("名称"), "'", "''") & "','" & Replace(Me.txt医生嘱托, "'", "''") & "'," & _
                    "'" & IIf(intType = 4, str标本部位, .Fields("标本部位")) & "','" & strSequence & "'," & _
                    IIf(lng频率次数 = 0, "NULL", lng频率次数) & "," & _
                    IIf(lng频率间隔 = 0, "NULL", lng频率间隔) & "," & _
                    "'" & str间隔单位 & "',NULL," & _
                    .Fields("计价性质") & "," & _
                    str执行科室ID & "," & _
                    .Fields("执行科室") & "," & Me.chk紧急.Value & "," & _
                    IIf(Me.chk开始时间.Visible And Me.chk开始时间.Value = 0, "NULL,", "To_Date('" & Format(Me.txt开始时间.Value, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),") & _
                    "NULL," & _
                    lng病人科室ID & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
                    "Sysdate,'" & IIf(PatientType = 1, CheckID, "") & "'," & _
                    IIf(mlng前提ID = 0, "Null", mlng前提ID) & ")"
                    Call SQLTest(App.ProductName, Me.Caption, strSQL)
                    gcnOracle.Execute strSQL, , adCmdStoredProc
                    Call SQLTest
                
                iSendSeq = iSendSeq + 1
                strSQL = "ZL_病人医嘱发送_Insert(" & _
                    lngTmpID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
                    iSendSeq & "," & Me.txt总量 & ",NULL,NULL," & _
                    "Sysdate+1/(24*3600)," & _
                    "0," & str执行科室ID & "," & IIf(Len(sCheckNo) = 0, 0, 1) & ",0)"
                Call SQLTest(App.ProductName, Me.Caption, strSQL)
                gcnOracle.Execute strSQL, , adCmdStoredProc
                Call SQLTest
                
                i = i + 1
                .MoveNext
            End With
        Loop
    End If
    If intType = 4 Then
        '检验申请的采集方式放到最后
        iMaxSeq = iMaxSeq + 1
        strSQL = "ZL_病人医嘱记录_Insert(" & lngAdviceID & ",NULL," & _
            iMaxSeq & "," & PatientType & "," & lng病人id & "," & IIf(PatientType = 2, CheckID, "NULL") & "," & _
            "0,1," & _
            "1,'" & tmpstr类别 & "'," & _
            tmplngClinicID & ",NULL,NULL," & _
            IIf(Len(Trim(Me.txt单量)) = 0, "NULL", Me.txt单量) & "," & _
            IIf(Len(Trim(Me.txt总量)) = 0, "NULL", Me.txt总量) & "," & _
            "'" & Replace(strAdviceText, "'", "''") & "','" & Replace(Me.txt医生嘱托, "'", "''") & "'," & _
            "'" & str标本部位 & "','" & strSequence & "'," & _
            IIf(lng频率次数 = 0, "NULL", lng频率次数) & "," & _
            IIf(lng频率间隔 = 0, "NULL", lng频率间隔) & "," & _
            "'" & str间隔单位 & "',NULL," & _
            tmpint计价特性 & "," & _
            str执行科室ID1 & "," & _
            tmpint执行性质 & "," & Me.chk紧急.Value & "," & _
            IIf(Me.chk开始时间.Visible And Me.chk开始时间.Value = 0, "NULL,", "To_Date('" & Format(Me.txt开始时间.Value, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),") & _
            "NULL," & _
            lng病人科室ID & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
            "Sysdate,'" & IIf(PatientType = 2, "", CheckID) & "'," & _
            IIf(mlng前提ID = 0, "Null", mlng前提ID) & ")"
        Call SQLTest(App.ProductName, Me.Caption, strSQL)
        gcnOracle.Execute strSQL, , adCmdStoredProc
        Call SQLTest
        
        iSendSeq = iSendSeq + 1
    End If
    
    '发送主医嘱
    If intType <> 4 Then iSendSeq = 1 '非检验类的主医嘱放在前面
    strSQL = "ZL_病人医嘱发送_Insert(" & _
        lngAdviceID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
        iSendSeq & "," & Me.txt总量 & ",NULL,NULL," & _
        "Sysdate+1/(24*3600)," & _
        "0," & str执行科室ID & "," & IIf(Len(sCheckNo) = 0, 0, 1) & ",1)"
'        "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
'        "0," & str执行科室ID & ",0,1)"
    Call SQLTest(App.ProductName, Me.Caption, strSQL)
    gcnOracle.Execute strSQL, , adCmdStoredProc
    Call SQLTest
    '修改费用记录的医嘱序号
    If Len(sCheckNo) > 0 Then
        strSQL = "zl_病人费用记录_医嘱('" & strNO & "',1," & lngAdviceID & ")"
        Call SQLTest(App.ProductName, Me.Caption, strSQL)
        gcnOracle.Execute strSQL, , adCmdStoredProc
        Call SQLTest
    End If
    
    AdviceID = lngAdviceID
    '核收标本
    If mblnSample Then SaveSample

    gcnOracle.CommitTrans
    
    '发送申请信息
    If mblnSample And blnComm Then
        For i = 1 To vsf2.Rows - 1
            If mlngNoneHomeKey(i) = 0 Then
                If Not objLISComm.SendSample(IIf(Val(vsf2.RowData(i)) = -1, 0, Val(vsf2.RowData(i))), _
                    Format(dtp(1).Value, "yyyy-MM-dd HH:mm:ss"), TransSampleNO(vsf2.TextMatrix(i, 2)), , , IIf(blnEmerge And Me.chkEmerge.Value = 1, 1, 0)) Then
                    MsgBox "第 " & CLng(vsf2.TextMatrix(i, 2)) & " 号标本未能传送到仪器(" & _
                        vsf2.TextMatrix(i, 1) & ")，请稍后手动传送", vbInformation + vbOKOnly, gstrSysName
                End If
            End If
        Next
    End If
    
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "病人医嘱保存"
End Sub

Private Function SaveSample() As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能：
    '--------------------------------------------------------------------------------------------------------
    Dim strNow As String
    Dim varTmp As Variant
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim blnMuliQuest As Boolean
    Dim lngMuliQuestKey As Long
    Dim mlngKey As Long '医嘱ID
    Dim lngKey As Long '标本ID
    Dim i As Integer, varAdviceIDs As Variant '指标对应的若干医嘱ID
    Dim strItemRecords As String
    Dim mlngLoop As Long
    Dim lngTmpNO As Long '标本号
    
    On Error GoTo ErrHand
    
        
    strNow = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    ReDim strSQL(1 To 1)
    For mlngLoop = 1 To vsf2.Rows - 1
        
        '检查是否为多张申请对应一个标本的情况,如果是，那么不填写检验标本记录，只填写检验普通结果和修改医嘱发送记录
        
        blnMuliQuest = False
        lngMuliQuestKey = 0
        
        If CheckMuliQuest(PatientID, IIf(Val(vsf2.RowData(mlngLoop)) = -1, 0, Val(vsf2.RowData(mlngLoop))), TransSampleNO(vsf2.TextMatrix(mlngLoop, 2)), lngMuliQuestKey, _
            IIf(blnEmerge And Me.chkEmerge.Value = 1, 1, 0)) = False Then Exit Function
            
        If lngMuliQuestKey = 0 Then
        
            If mlngNoneHomeKey(mlngLoop) = 0 Then
                lngKey = zldatabase.GetNextId("检验标本记录")
            Else
                lngKey = mlngNoneHomeKey(mlngLoop)
            End If
        
        Else
            lngKey = lngMuliQuestKey
        End If
        
        mlngKey = AdviceID '核收的默认医嘱ID
        lngTmpNO = TransSampleNO(vsf2.TextMatrix(mlngLoop, 2))
        strSQL(ReDimArray(strSQL)) = "ZL_检验标本记录_标本核收(" & lngKey & "," & _
                                                                mlngKey & ",'" & _
                                                                lngTmpNO & "'," & _
                                                                "TO_DATE('" & Format(dtp(0).Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & _
                                                                IIf(InStr(cbo(2).Text, "-") > 0, zlCommFun.GetNeedName(cbo(2).Text), cbo(2).Text) & "'," & _
                                                                IIf(Val(vsf2.RowData(mlngLoop)) = -1, 0, Val(vsf2.RowData(mlngLoop))) & "," & _
                                                                "TO_DATE('" & Format(dtp(1).Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & _
                                                                IIf(InStr(cbo(0).Text, "-") > 0, zlCommFun.GetNeedName(cbo(0).Text), cbo(0).Text) & "'," & _
                                                                "0,'" & _
                                                                IIf(InStr(cbo(2).Text, "-") > 0, zlCommFun.GetNeedName(cbo(2).Text), cbo(2).Text) & "'," & _
                                                                "TO_DATE('" & Format(dtp(1).Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')," & mlngNoneHomeKey(mlngLoop) & "," & IIf(mbln微生物项目, 1, 0) & "," & lngMuliQuestKey & "," & _
                                                                IIf(blnEmerge And Me.chkEmerge.Value = 1, 1, 0) & ")"
        
        
        '注意：如果是微生物检验项目，则核收时不填写检验普通结果记录
'        If mbln微生物项目 = False And mlngNoneHomeKey(mlngLoop) <= 0 Then
        If mbln微生物项目 = False Then
            varTmp = Split(vsf2.TextMatrix(mlngLoop, 3), "|")
            strItemRecords = ""
            For lngLoop = 0 To UBound(varTmp)
                mlngKey = AdviceID '指标对应的医嘱ID
                strItemRecords = strItemRecords & "|" & mlngKey & "^" & Val(Split(varTmp(lngLoop), "^")(0)) & "^" & IIf(Val(Split(varTmp(lngLoop), "^")(5)) = 3, "-", "") & "^^"
            Next lngLoop
            If Len(strItemRecords) > 0 Then
                strItemRecords = Mid(strItemRecords, 2)
                    
                strSQL(ReDimArray(strSQL)) = "ZL_检验普通结果_BATCHINSERT(" & lngKey & "," & _
                    IIf(Val(vsf2.RowData(mlngLoop)) = -1, 0, Val(vsf2.RowData(mlngLoop))) & ",'" & _
                    strItemRecords & "')"
            End If
        End If
        
        If vsf2.RowData(mlngLoop) = -1 Then mstrCurrentNO = vsf2.TextMatrix(mlngLoop, 2)
    Next
    
    strSQL(ReDimArray(strSQL)) = "ZL_检验试剂记录_BatchInsert(" & AdviceID & ")"
    '实际不是核收，只是改变医嘱执行状态
    lngTmpNO = TransSampleNO(vsf2.TextMatrix(vsf2.Rows - 1, 2))
    strSQL(ReDimArray(strSQL)) = "ZL_检验标本记录_标本核收(" & lngKey & "," & _
                                                            AdviceID & ",'" & _
                                                            lngTmpNO & "'," & _
                                                            "TO_DATE('" & Format(dtp(0).Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & _
                                                            IIf(InStr(cbo(2).Text, "-") > 0, zlCommFun.GetNeedName(cbo(2).Text), cbo(2).Text) & "'," & _
                                                            IIf(Val(vsf2.RowData(vsf2.Rows - 1)) = -1, 0, Val(vsf2.RowData(vsf2.Rows - 1))) & "," & _
                                                            "TO_DATE('" & Format(dtp(1).Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & _
                                                            IIf(InStr(cbo(0).Text, "-") > 0, zlCommFun.GetNeedName(cbo(0).Text), cbo(0).Text) & "'," & _
                                                            "1,'" & _
                                                            IIf(InStr(cbo(2).Text, "-") > 0, zlCommFun.GetNeedName(cbo(2).Text), cbo(2).Text) & "'," & _
                                                            "TO_DATE('" & Format(dtp(1).Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')," & mlngNoneHomeKey(vsf2.Rows - 1) & "," & IIf(mbln微生物项目, 1, 0) & ",1," & _
                                                            IIf(blnEmerge And Me.chkEmerge.Value = 1, 1, 0) & ")"
    
    
    For mlngLoop = 1 To UBound(strSQL)
        If strSQL(mlngLoop) <> "" Then Call ExecuteProc(strSQL(mlngLoop), Me.Caption)
    Next
    
    SaveSample = True
    
    Exit Function
ErrHand:
    Err.Raise Err.Number, "标本核收"
End Function

Private Function CheckMuliQuest(ByVal lng病人id As Long, ByVal lng仪器id As Long, ByVal strNO As String, ByRef lngKey As Long, ByVal iType As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If mbln微生物项目 Then
        CheckMuliQuest = True
        Exit Function
    End If
    
    If lng仪器id > 0 Then
        strSQL = "SELECT A.ID FROM 检验标本记录 A,病人医嘱记录 B WHERE A.样本状态=1 AND A.医嘱id=B.id AND B.病人id=[1]" & _
        " AND A.仪器id=[2] AND A.核收时间 Between [3] And [4] AND A.标本序号= [5] " & _
        IIf(iType = 1, " And A.标本类别=1", " And Nvl(A.标本类别,0)<>1")
    Else
        strSQL = "SELECT A.ID FROM 检验标本记录 A,病人医嘱记录 B WHERE A.样本状态=1 AND A.医嘱id=B.id AND B.病人id=[1]" & _
        " AND A.仪器id IS NULL AND A.核收时间 Between [3] And [4] AND A.标本序号= [5] " & _
        IIf(iType = 1, " And A.标本类别=1", " And Nvl(A.标本类别,0)<>1")
    End If
    Set rs = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人id, lng仪器id, _
        CDate(Format(dtp(1).Value, "yyyy-MM-dd 00:00:00")), _
        CDate(Format(dtp(1).Value, "yyyy-MM-dd 23:59:59")), strNO)
    
    If rs.BOF = False Then
        If MsgBox("当前的标本号本日已经使用，请问是否为多张申请合并一个标本？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            lngKey = zlCommFun.Nvl(rs("ID"), 0)
        Else
            Exit Function
        End If
    End If
    
    CheckMuliQuest = True
    
    Exit Function
    
ErrHand:
    
End Function

Private Function ValidSampleData(ByVal iType As Integer) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能：
    '--------------------------------------------------------------------------------------------------------
    Dim varTmp As Variant
    Dim strTmp As String
    Dim strError As String
    Dim lngLoop As Long
    Dim lngCount As Long
    Dim rs As New ADODB.Recordset
    Dim i As Integer, mlngLoop As Long
    Dim mstrSQL As String
    
    '1.检查每一个标本指定的检验仪器是否正确
    For mlngLoop = 1 To vsf2.Rows - 1
        If Trim(vsf2.TextMatrix(mlngLoop, 2)) = "" Then
            strError = "第" & mlngLoop & "个标本没有标本号！"
            
            vsf2.Row = mlngLoop
            vsf2.Col = 2
            vsf2.SetFocus
            vsf2.ShowCell vsf2.Row, vsf2.Col
            GoTo ErrHand
            
        End If
        
'        If Left(Trim(vsf2.TextMatrix(mlngLoop, 2)), 1) = "0" Then
'            strError = "第" & mlngLoop & "个标本无效，必须为数字型！"
'
'            vsf2.Row = mlngLoop
'            vsf2.Col = 2
'            vsf2.SetFocus
'            vsf2.ShowCell vsf2.Row, vsf2.Col
'
'            GoTo errHand
'
'        End If
'
'        If CheckStrType(Trim(vsf2.TextMatrix(mlngLoop, 2)), 99, "0123456789") = False Then
'            strError = "第" & mlngLoop & "个标本无效，必须为数字型！"
'
'            vsf2.Row = mlngLoop
'            vsf2.Col = 2
'            vsf2.SetFocus
'            vsf2.ShowCell vsf2.Row, vsf2.Col
'
'            GoTo errHand
'        End If
        
    Next
    
    If cbo(2).ListIndex = -1 Then
        strError = "核收标本时必须指定检验人员！"
        cbo(2).SetFocus
        GoTo ErrHand
    End If
    
    ReDim mlngNoneHomeKey(vsf2.Rows - 1)
    
    For i = 1 To vsf2.Rows - 1
'    If LngCount = 1 And Val(vsf2.RowData(1)) > 0 Then
        If Val(vsf2.RowData(i)) > 0 Then
    
            '检查是否有效
            mstrSQL = "SELECT ID,标本序号 FROM 检验标本记录 WHERE 医嘱id IS NULL AND Nvl(是否质控品,0)<>1 AND 仪器id= [1] " & _
                " AND 核收时间 Between [2] AND [3] AND 标本序号=[4]" & _
                IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1")
            Set rs = zldatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(vsf2.RowData(i)), _
                CDate(Format(dtp(1).Value, "yyyy-MM-dd 00:00:00")), _
                CDate(Format(dtp(1).Value, "yyyy-MM-dd 23:59:59")), TransSampleNO(Trim(vsf2.TextMatrix(i, 2))))
            
            If rs.BOF = False Then
                If MsgBox("你设置的标本号已经存在一个无主标本，是否对应无主标本！", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    vsf2.TextMatrix(i, 2) = TransSampleNO_PH(rs("标本序号").Value, Val(vsf2.RowData(i)))
                    mlngNoneHomeKey(i) = rs("ID").Value
                Else
                    Exit Function
                End If
            End If
            
        End If
    Next
    
    ValidSampleData = True
    
    Exit Function
ErrHand:
    ValidSampleData = False
    MsgBox strError, vbInformation, gstrSysName
End Function

Private Function GetOneDept(lng收费细目ID As Long) As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.执行部门ID From 收费细目 A,收费执行部门 B Where B.收费细目ID=A.ID And A.ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng收费细目ID)
    If Not rsTmp.EOF Then
        GetOneDept = rsTmp!执行部门ID '默认取第一个(如有多个)
    Else
        GetOneDept = UserInfo.部门ID '如没有指定，则取操作员所在科室
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'========以下是医嘱编辑==========

Private Sub cbo执行科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub chk紧急_Click()
    On Error Resume Next
    Me.txt医生嘱托.SetFocus
End Sub

Private Sub chk紧急_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub chk开始时间_Click()
    On Error Resume Next
    If Me.chk开始时间.Value = 1 Then
        Me.txt开始时间.Enabled = True: Me.txt开始时间.SetFocus
    Else
        Me.txt开始时间.Enabled = False
    End If
    
    If str类别 = "D" Then
        strAdviceText = Get检查手术内容(1, strClinicName)
    ElseIf str类别 = "F" Then
        strAdviceText = Get检查手术内容(2, strClinicName)
    End If
End Sub

Private Sub chk开始时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo医生_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo医生_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo医生.ListIndex <> -1 Then Exit Sub '已选中
    If cbo医生.Text = "" Then '无输入
        Exit Sub
    End If
    
    strInput = UCase(NeedName(cbo医生.Text))
    '全院医生
    strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN(1,2,3)"
    strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
        " And B.部门ID IN(" & strSQL & ")" & _
        " And (Upper(A.编号) Like [1] Or Upper(A.姓名) Like [2] Or Upper(A.简码) Like [2])" & _
        " Order by A.简码"
    
    On Error GoTo errH
    vRect = GetControlRect(cbo医生.Hwnd)
    Set rsTmp = zldatabase.ShowSQLSelect(Me, strSQL, 0, lbl开嘱医生.Caption, False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo医生.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If Not rsTmp Is Nothing Then
        cbo医生.Text = rsTmp!姓名
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的医生。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdExt_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim tmpExtData As String
    
    frmAdviceEditEx.mlngHwnd = Me.cbo医生.Hwnd 'txt附加.Hwnd
    frmAdviceEditEx.mintType = IIf(intType = 4, 3, intType)
    frmAdviceEditEx.mint期效 = 1
    frmAdviceEditEx.mstr性别 = mstr性别
    If intType = 4 Then
        '检验项目
        frmAdviceEditEx.mlng项目id = 0 'Split(strExtData, ";")(0)
        frmAdviceEditEx.mstrExtData = strExtData ' Split(strExtData, ";")(1)
    Else
        frmAdviceEditEx.mlng项目id = lngClinicID
        frmAdviceEditEx.mstrExtData = strExtData
    End If
    frmAdviceEditEx.mint服务对象 = PatientType

    On Error Resume Next
    frmAdviceEditEx.Show 1, Me

    If Not frmAdviceEditEx.mblnOK Then
        zlControl.TxtSelAll Me.txt附加
        Me.txt附加.SetFocus
        Exit Sub
    Else
        tmpExtData = frmAdviceEditEx.mstrExtData
        If intType = 4 Then
            strExtData = Split(strExtData, ";")(0) + ";" + tmpExtData
        Else
            strExtData = tmpExtData
        End If
    End If
    Select Case intType
        Case 0 '检查组合部位
            Call AdviceSet检查手术(1, strExtData)
            strAdviceText = Get检查手术内容(1, strClinicName)
            Me.txt附加 = Get部位名称
        Case 1 '麻醉项目
            Call AdviceSet检查手术(2, strExtData)
            txt医嘱内容.Text = Get检查手术名称(2, strClinicName)
            strAdviceText = Get检查手术内容(2, strClinicName)
            Me.txt附加 = Get麻醉名称
        Case 4 '检验项目
            strAdviceText = strClinicName & "(" & tmpExtData & ")"
            Me.txt附加 = tmpExtData: str标本部位 = tmpExtData
    End Select
    txt附加.Tag = txt附加.Text
    Me.txt附加.SetFocus
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdSel_Click()
    Dim rsTmp As ADODB.Recordset
    
    If intType = 4 Then
        '检验项目
        If LabsInput Then
            txt医嘱内容.Tag = txt医嘱内容.Text
            txt附加.Tag = txt附加.Text
            
            If mblnSample Then ReadSampleData
            
            Me.txt医嘱内容.SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            '恢复原值
            txt医嘱内容.Text = txt医嘱内容.Tag
            txt附加.Text = txt附加.Tag
            zlControl.TxtSelAll txt医嘱内容
            txt医嘱内容.SetFocus
        End If
        Exit Sub
    End If
    
    With txt医嘱内容
        .Text = ""
        Set rsTmp = SelectDiagItem()
    End With
    
    If rsTmp Is Nothing Then '取消或无数据
        '恢复原值
        zlControl.TxtSelAll txt医嘱内容
        txt医嘱内容.SetFocus: Exit Sub
    End If
    '新项目的录入
    
    '根据选择项目设置缺省医嘱信息
    If AdviceInput(rsTmp) Then
        '显示已缺省设置的值
        txt医嘱内容.Tag = txt医嘱内容.Text
        txt附加.Tag = txt附加.Text
        Me.txt医嘱内容.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        '恢复原值
        txt医嘱内容.Text = txt医嘱内容.Tag
        txt附加.Text = txt附加.Tag
        zlControl.TxtSelAll txt医嘱内容
        txt医嘱内容.SetFocus
    End If
End Sub

Private Sub cmd频率_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int范围 As Integer, vRect As RECT
        
    int范围 = 1
    strSQL = "Select Rownum as ID,A.编码,A.名称,A.简码," & _
        " A.英文名称,A.频率次数,A.频率间隔,nvl(A.间隔单位,' ') As 间隔单位" & _
        " From 诊疗频率项目 A Where A.适用范围=" & int范围 & _
        " Order by A.编码"
    vRect = GetControlRect(txt频率.Hwnd)
    Set rsTmp = zldatabase.ShowSelect(Me, strSQL, 0, "诊疗频率", , , , , , True, vRect.Left, vRect.Top, txt频率.Height, blnCancel, , True)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "没有可用的诊疗频率项目，请先到医嘱频率管理中设置。", vbInformation, gstrSysName
        End If
        txt频率.Text = strSequence
        Call zlControl.TxtSelAll(txt频率)
        txt频率.SetFocus: Exit Sub
    End If
    Me.cmd频率.Tag = rsTmp("名称"): Me.txt频率 = Me.cmd频率.Tag: strSequence = Me.cmd频率.Tag
    lng频率次数 = rsTmp("频率次数"): lng频率间隔 = rsTmp("频率间隔"): str间隔单位 = Trim(rsTmp("间隔单位"))

    txt频率.SetFocus
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt采集_GotFocus()
    Call zlControl.TxtSelAll(txt采集)
End Sub

Private Sub txt采集_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strItemID As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt采集.Text = txt采集.Tag Then
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    If Len(strExtData) > 0 Then
        strItemID = Split(strExtData, ";")(0)
        If Len(strItemID) > 0 Then strItemID = Split(strItemID, ",")(0)
    End If
    Set rsTmp = SelectCap(Val(strItemID), Me.txt采集)
    If Not rsTmp Is Nothing Then
        Me.cmd采集.Tag = rsTmp("ID")
        Me.txt采集 = rsTmp("名称"): Me.txt采集.Tag = Me.txt采集
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt采集_Validate(Cancel As Boolean)
    '恢复人为的改变
    If txt采集.Text <> txt采集.Tag Then
        txt采集.Text = txt采集.Tag
    End If
End Sub

Private Sub txt单量_GotFocus()
    zlControl.TxtSelAll txt单量
End Sub

Private Sub txt单量_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or ifEditKey(KeyAscii, False)) Then KeyAscii = 0
End Sub

Private Sub txt单量_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txt单量) Then Me.txt单量 = 1: Exit Sub
    Me.txt单量 = CInt(Me.txt单量)
    If CInt(Me.txt单量) < 1 Then Me.txt单量 = 1
End Sub

Private Sub txt附加_DblClick()
    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
End Sub

Private Sub txt附加_GotFocus()
    Call zlControl.TxtSelAll(txt附加)
End Sub

Private Sub txt附加_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txt附加)
    End If
End Sub

Private Sub txt附加_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt附加.Text = txt附加.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        cmdExt_Click
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt附加_Validate(Cancel As Boolean)
    '恢复人为的改变
    If txt附加.Text <> txt附加.Tag Then
        txt附加.Text = txt附加.Tag
    End If
End Sub

Private Sub txt开始时间_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt开始时间_Validate(Cancel As Boolean)
    On Error Resume Next
    If Not Check开始时间(CStr(txt开始时间)) Then
        Cancel = True
        txt开始时间.SetFocus
    Else
        If str类别 = "D" Then
            strAdviceText = Get检查手术内容(1, strClinicName)
        ElseIf str类别 = "F" Then
            strAdviceText = Get检查手术内容(2, strClinicName)
        End If
    End If
End Sub

Private Sub txt频率_GotFocus()
    Call zlControl.TxtSelAll(txt频率)
End Sub

Private Sub txt频率_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int范围 As Integer, vRect As RECT
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmd频率.Tag <> "" And txt频率.Text = strSequence And txt频率.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt频率.Text = "" Then
            If cmd频率.Enabled And cmd频率.Visible Then cmd频率_Click
        Else
            int范围 = 1 '可选频率
            strSQL = "Select Rownum as ID,A.编码,A.名称,A.简码," & _
                " A.英文名称,A.频率次数,A.频率间隔,A.间隔单位" & _
                " From 诊疗频率项目 A Where A.适用范围=" & int范围 & _
                " And (A.编码 Like '" & UCase(txt频率.Text) & "%'" & _
                " Or Upper(A.名称) Like '" & mstrLike & UCase(txt频率.Text) & "%'" & _
                " Or Upper(A.简码) Like '" & mstrLike & UCase(txt频率.Text) & "%'" & _
                " Or Upper(A.英文名称) Like '" & mstrLike & UCase(txt频率.Text) & "%')" & _
                " Order by A.编码"
            vRect = GetControlRect(txt频率.Hwnd)
            Set rsTmp = zldatabase.ShowSelect(Me, strSQL, 0, "诊疗频率", , , , , , True, vRect.Left, vRect.Top, txt频率.Height, blnCancel, , True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配的诊疗频率项目。", vbInformation, gstrSysName
                End If
                txt频率.Text = strSequence
                Call zlControl.TxtSelAll(txt频率)
                txt频率.SetFocus: Exit Sub
            End If
            Me.cmd频率.Tag = rsTmp("名称"): Me.txt频率 = Me.cmd频率.Tag: strSequence = Me.cmd频率.Tag
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt频率_Validate(Cancel As Boolean)
    If cmd频率.Tag <> "" And txt频率.Text <> strSequence Then
        txt频率.Text = strSequence
    End If
End Sub

Private Sub txt姓名_Validate(Cancel As Boolean)
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strField As String
    
    If Len(Trim(txt姓名)) = 0 Then Exit Sub
    Set rsTmp = GetPatient(txt姓名)
    If rsTmp.EOF Then
        Me.txt单位电话 = ""
        Me.txt单位名称 = ""
        Me.txt单位邮编 = ""
        Me.txt家庭地址 = ""
        Me.txt家庭电话 = ""
        Me.txt家庭邮编 = ""
        Me.txt门诊号 = ""
        Me.txt年龄 = ""
        Me.txt身份证号 = ""
        If InStr("+-*.", Left(Me.txt姓名.Text, 1)) > 0 Then Me.txt姓名.Text = "": Cancel = True
        
        PatientID = 0: PatientType = 1: CheckID = "": sCheckNo = ""
    Else
        On Error Resume Next
        Me.txt姓名.Text = Nvl(rsTmp("姓名"))
        Me.txt单位电话 = Nvl(rsTmp("单位电话"))
        Me.txt单位名称 = Nvl(rsTmp("工作单位"))
        Me.txt单位邮编 = Nvl(rsTmp("单位邮编"))
        Me.txt家庭地址 = Nvl(rsTmp("家庭地址"))
        Me.txt家庭电话 = Nvl(rsTmp("家庭电话"))
        Me.txt家庭邮编 = Nvl(rsTmp("户口邮编"))
        Me.txt门诊号 = Nvl(rsTmp("门诊号"))
        Me.txt年龄 = Nvl(rsTmp("年龄"))
        Me.txt身份证号 = Nvl(rsTmp("身份证号"))
        Me.cbo费别 = Nvl(rsTmp("费别")) 'CombIndex(cbo费别, Nvl(rsTmp("费别")))
        Me.cbo付款方式 = Nvl(rsTmp("医疗付款方式")) ' CombIndex(cbo付款方式, Nvl(rsTmp("医疗付款方式")))
        Me.cbo国籍 = Nvl(rsTmp("国籍")) ' CombIndex(cbo国籍, Nvl(rsTmp("国籍")))
        Me.cbo婚姻 = Nvl(rsTmp("婚姻状况")) 'CombIndex(cbo婚姻, Nvl(rsTmp("婚姻状况")))
        Me.cbo民族 = Nvl(rsTmp("民族")) 'CombIndex(cbo民族, Nvl(rsTmp("民族")))
        Me.cbo性别 = Nvl(rsTmp("性别")) 'CombIndex(cbo性别, Nvl(rsTmp("性别")))
        Me.cbo职业 = Nvl(rsTmp("职业")) 'CombIndex(cbo职业, Nvl(rsTmp("职业")))
        
        PatientID = Nvl(rsTmp("病人ID"), 0): PatientType = Nvl(rsTmp("PatientType"), 1): CheckID = Nvl(rsTmp("主页ID"))
        '设置默认开单科室、医生
        For i = 0 To Me.cbo开单科室.ListCount - 1
            If Me.cbo开单科室.ItemData(i) = Nvl(rsTmp("病人科室"), 0) Then
                Me.cbo开单科室.ListIndex = i
                Exit For
            End If
        Next
        DoEvents
        strField = ""
        strField = rsTmp.Fields("医生").Name
        If strField = "医生" Then
            Me.cbo医生.Text = Nvl(rsTmp("医生"))
            For i = 0 To Me.cbo医生.ListCount - 1
                If Me.cbo医生.List(i) Like Nvl(rsTmp("医生")) Then
                    Me.cbo医生.ListIndex = i
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Private Sub txt医生嘱托_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt医生嘱托_Validate(Cancel As Boolean)
    On Error Resume Next
    If zlCommFun.ActualLen(txt医生嘱托.Text) > txt医生嘱托.MaxLength Then
        MsgBox "输入内容不过超过 " & txt医生嘱托.MaxLength \ 2 & " 个汉字或 " & txt医生嘱托.MaxLength & " 个字符。", vbInformation, gstrSysName
        txt医生嘱托.SetFocus
        Cancel = True
    End If
End Sub

Private Sub txt医嘱内容_DblClick()
    If cmdSel.Visible And cmdSel.Enabled Then cmdSel_Click
End Sub

Private Sub txt医嘱内容_GotFocus()
    Call zlControl.TxtSelAll(txt医嘱内容)
End Sub

Private Sub txt医嘱内容_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txt医嘱内容)
    End If
End Sub

Private Sub txt医嘱内容_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt医嘱内容.Text = "" Then cmdSel_Click: Exit Sub
        If txt医嘱内容.Text = txt医嘱内容.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        
        With txt医嘱内容
            Set rsTmp = SelectDiagItem()
        End With
        
        If rsTmp Is Nothing Then '取消或无数据
            '恢复原值
            txt医嘱内容.Text = txt医嘱内容.Tag
            zlControl.TxtSelAll txt医嘱内容
            txt医嘱内容.SetFocus: Exit Sub
        End If
        '新项目的录入
        
        '根据选择项目设置缺省医嘱信息
        If AdviceInput(rsTmp) Then
            '显示已缺省设置的值
            txt医嘱内容.Tag = txt医嘱内容.Text
            txt附加.Tag = txt附加.Text
            
            If mblnSample Then ReadSampleData
            
            If Not ifInitItem Then Call zlCommFun.PressKey(vbKeyTab)
        Else
            '恢复原值
            txt医嘱内容.Text = txt医嘱内容.Tag
            txt附加.Text = txt附加.Tag
            zlControl.TxtSelAll txt医嘱内容
            If Not ifInitItem Then txt医嘱内容.SetFocus: Exit Sub
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        If cmdSel.Visible And cmdSel.Enabled Then Call cmdSel_Click
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt医嘱内容_Validate(Cancel As Boolean)
    '恢复人为的改变
    If txt医嘱内容.Text <> txt医嘱内容.Tag Then
        txt医嘱内容.Text = txt医嘱内容.Tag
    End If
End Sub

Private Sub txt总量_GotFocus()
    Call zlControl.TxtSelAll(Me.txt总量)
End Sub

Private Sub txt总量_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or ifEditKey(KeyAscii, False)) Then KeyAscii = 0
End Sub

Private Sub txt总量_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txt总量) Then Me.txt总量 = 1: Exit Sub
    Me.txt总量 = CInt(Me.txt总量)
    If CInt(Me.txt总量) < 1 Then Me.txt总量 = 1
End Sub

'判断是否为编辑键
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Private Function Check开始时间(ByVal strStart As String, _
    Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'功能：检查输入的开始时间是否合法
'说明：
'1.开始时间不能小于病人的入院时间
'2.开始时间必须小于终止时间
'3.正常录入时,开始时间不能小于当前时间之前30分钟(从而可能造成开嘱时间大于开始时间30分钟)
'4.补录的医嘱开始时间不能大于当前时间
    Dim strInDate As String
    
    If Not IsDate(strStart) Then
        MsgBox "输入的医嘱开始执行时间无效。", vbInformation, gstrSysName
        Exit Function
    End If
        
    strInDate = Format(PatientDate, "yyyy-MM-dd HH:mm")
    If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
        strMsg = "医嘱的开始执行时间不能小于病人的" & IIf(PatientType = 2, "入院", "就诊") & "时间 " & strInDate & " 。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
'    If IsDate(strEnd) Then
'        If Format(strStart, "yyyy-MM-dd HH:mm") >= Format(strEnd, "yyyy-MM-dd HH:mm") Then
'            strMsg = "医嘱的开始执行时间必须小于执行终止时间。"
'            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    
    If DateDiff("n", CDate(strStart), zldatabase.Currentdate) > 30 Then
        strMsg = "开始执行时间不能太早于当前时间。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    Check开始时间 = True
End Function
Private Function SelectDiagItem() As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select Distinct A.ID,A.编码,A.名称,nvl(A.计算单位,'次') As 计算单位,nvl(A.标本部位,' ') As 标本部位," + _
        "Decode(A.类别,'H',Decode(A.操作类型,'1','护理等级','护理常规')," + _
        "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法',4,'中药用法','其它')," + _
        "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院','其它'),A.操作类型) As 项目特性,A.类别 As 类别ID,A.ID As 诊疗项目ID,nvl(执行频率,0) As 执行频率ID,nvl(计算方式,0) As 计算方式ID,nvl(执行安排,0) As 执行安排ID,nvl(计价性质,0) As 计价性质ID,nvl(执行科室,0) As 执行科室ID "
    Select Case ItemType
        Case 1 'PACS
            strSQL = strSQL + "From 诊疗项目目录 A,影像检查项目 B,诊疗项目别名 C,诊疗执行科室 D Where A.ID=B.诊疗项目ID And A.ID=C.诊疗项目ID And A.ID=D.诊疗项目ID And D.执行科室ID=" & ItemDeptID
        Case 2 'LIS
            strSQL = strSQL + "From 诊疗项目目录 A,诊疗项目别名 C,诊疗执行科室 D Where A.ID=C.诊疗项目ID And A.ID=D.诊疗项目ID And A.类别='C' And D.执行科室ID=" & ItemDeptID
    End Select
    strSQL = strSQL + " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
        "And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) IN (" + _
        IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
        " And Nvl(A.执行频率,0) IN(0,1)" + _
        " And (A.编码 Like '" + txt医嘱内容 + "%' Or Upper(A.名称) Like '" + mstrLike + txt医嘱内容 + "%' Or Upper(C.简码) Like '" + mstrLike + UCase(txt医嘱内容) + "%')"
            
    With txt医嘱内容
        Set SelectDiagItem = zldatabase.ShowSelect(Me, strSQL, 0, "选择申请项目", True, .Text, "", True, True, True, .Left + Me.picAdvice.Left + Me.Left, .Top + Me.picAdvice.Top + Me.Top, .Height, False, True)
    End With
End Function

Private Function SelectCap(Optional ByVal lngItemID As Long = 0, Optional ByVal QryStr As String = "", Optional blnNotSelect As Boolean = False) As ADODB.Recordset
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim tmpRect As RECT
    
    On Error GoTo DBError
    If Len(QryStr) > 0 Then
        strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
            "From 诊疗项目目录 A,诊疗项目别名 C,诊疗用法用量 D Where A.ID=C.诊疗项目ID And A.ID=D.用法ID" + _
            " And A.类别='E' And A.操作类型='6'" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
            " And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.适用性别,0) IN (" + _
            IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
            " And Nvl(A.执行频率,0) IN(0,1)" + _
            " And D.项目ID=" & lngItemID & _
            " And (A.编码 Like '" + QryStr + "%' Or Upper(A.名称) Like '" + mstrLike + QryStr + "%' Or Upper(C.简码) Like '" + mstrLike + UCase(QryStr) + "%')"
        OpenRecord rsTmp, strSQL, Me.Caption
        If rsTmp.EOF Then
            strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
                "From 诊疗项目目录 A,诊疗项目别名 C Where A.ID=C.诊疗项目ID" + _
                " And A.类别='E' And A.操作类型='6'" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
                " And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.适用性别,0) IN (" + _
                IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
                " And Nvl(A.执行频率,0) IN(0,1)" + _
                " And (A.编码 Like '" + QryStr + "%' Or Upper(A.名称) Like '" + mstrLike + QryStr + "%' Or Upper(C.简码) Like '" + mstrLike + UCase(QryStr) + "%')"
        End If
    Else
        strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
            "From 诊疗项目目录 A,诊疗用法用量 D Where A.ID=D.用法ID" + _
            " And A.类别='E' And A.操作类型='6'" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
            " And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.适用性别,0) IN (" + _
            IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
            " And Nvl(A.执行频率,0) IN(0,1)" + _
            " And D.项目ID=" & lngItemID
        OpenRecord rsTmp, strSQL, Me.Caption
        If rsTmp.EOF Then
            strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
                "From 诊疗项目目录 A Where " + _
                " A.类别='E' And A.操作类型='6'" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
                " And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.适用性别,0) IN (" + _
                IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
                " And Nvl(A.执行频率,0) IN(0,1)"
        End If
    End If
    If blnNotSelect Then
        If rsTmp.State = adStateOpen Then rsTmp.Close: Set rsTmp = New ADODB.Recordset
        OpenRecord rsTmp, strSQL, Me.Caption
        If Not rsTmp.EOF Then Set SelectCap = rsTmp
    Else
        tmpRect = GetControlRect(Me.txt采集.Hwnd)
        Set SelectCap = zldatabase.ShowSelect(Me, strSQL, 0, "采集方式", True, , , , , True, _
            tmpRect.Left, tmpRect.Top, Me.txt采集.Height, , , True)
    End If
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceInput(rsInput As ADODB.Recordset) As Boolean
'功能：根据新输的诊疗项目(新增或更换)设置缺省的医嘱数据
'参数：rsInput=输入或选择返回的记录集
'返回：本次录入是否有效
    Dim str过敏 As String, blnGroup As Boolean, i As Long
    Dim lng用法ID As Long, lngGroupRow As Long
    Dim lngPreRow As Long, lngNextRow As Long
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim intTmpType As Integer
    Dim strSQL As String

    On Error GoTo errH

    '项目附加数据输入及输入合法性检查
    '---------------------------------------------------------------------------------------------------------------
    txt医嘱内容.Text = rsInput!名称 '暂时显示

    '需要输入更多数据的一些项目
    '---------------------------------------------------------------------------------------------------------------
    intTmpType = -1
    If rsInput!类别ID = "D" And zlCommFun.Nvl(GetItemField(rsInput!诊疗项目ID, "组合项目"), 0) = 1 Then
        '检查组合项目
        intTmpType = 0
        strHelpText = "检查部位"
    ElseIf rsInput!类别ID = "F" Then
        '手术：需要输入麻醉项目，及可选择附加手术
        intTmpType = 1
        strHelpText = "附加手术及麻醉方式"
    ElseIf InStr(",7,8,", rsInput!类别ID) > 0 Then
        '中药配方(单味草药当配方处理)
        intTmpType = 2
    ElseIf rsInput!类别ID = "C" Then
        '检验项目选择检验标本
        intTmpType = 4
        strHelpText = "检验项目"
    End If

    If intTmpType <> -1 Then
        frmAdviceEditEx.mlngHwnd = Me.cbo执行科室.Hwnd ' txt医嘱内容.Hwnd
        frmAdviceEditEx.mintType = intTmpType
        frmAdviceEditEx.mint期效 = 1
        frmAdviceEditEx.mstr性别 = mstr性别
        frmAdviceEditEx.mlng项目id = IIf(intTmpType = 4, 0, rsInput!诊疗项目ID)
        frmAdviceEditEx.mstrExtData = IIf(intTmpType = 4, rsInput!诊疗项目ID & ";" & IIf(ifInitItem, Me.txt附加, ""), "") '新输入项目
        frmAdviceEditEx.mint服务对象 = PatientType

        On Error Resume Next
        If Not ifInitItem Then frmAdviceEditEx.Show 1, Me
        On Error GoTo errH

        If Not ifInitItem And Not frmAdviceEditEx.mblnOK Then Exit Function
        If frmAdviceEditEx.mstrExtData = "" Or Mid(frmAdviceEditEx.mstrExtData, 1, 1) = ";" Then Exit Function
        
        If rsInput!类别ID = "D" And frmAdviceEditEx.mstrExtData <> "" Then
            strAdviceText = txt医嘱内容.Text
            strExtData = frmAdviceEditEx.mstrExtData
            str标本部位 = Trim(rsInput("标本部位"))
            
            '检查的组合部位行
            Call AdviceSet检查手术(1, strExtData)
            txt医嘱内容.Text = Get检查手术名称(1, rsInput!名称)
            strAdviceText = Get检查手术内容(1, rsInput!名称)
            Me.txt附加 = Get部位名称
        ElseIf rsInput!类别ID = "F" And frmAdviceEditEx.mstrExtData <> "" Then
            strAdviceText = txt医嘱内容.Text
            strExtData = frmAdviceEditEx.mstrExtData
            str标本部位 = Trim(rsInput("标本部位"))
            
            '手术的附加手术及麻醉项目行
            Call AdviceSet检查手术(2, strExtData)
            txt医嘱内容.Text = Get检查手术名称(2, rsInput!名称)
            strAdviceText = Get检查手术内容(2, rsInput!名称)
            Me.txt附加 = Get麻醉名称
        ElseIf rsInput!类别ID = "C" And frmAdviceEditEx.mstrExtData <> "" Then
            '获取采集方式
            Set rsTmp = SelectCap(Split(Split(frmAdviceEditEx.mstrExtData, ";")(0), ",")(0), , True)
            If rsTmp Is Nothing Then
                MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
                Exit Function
            End If
            Me.cmd采集.Tag = rsTmp("ID")
            Me.txt采集 = rsTmp("名称"): Me.txt采集.Tag = Me.txt采集
            
            strAdviceText = txt医嘱内容.Text
            strExtData = frmAdviceEditEx.mstrExtData
            str标本部位 = Trim(rsInput("标本部位"))
            
            '检验项目
            strSQL = "Select Distinct A.ID,A.编码,A.名称,nvl(A.计算单位,'次') As 计算单位,nvl(A.标本部位,' ') As 标本部位," + _
                "Decode(A.类别,'H',Decode(A.操作类型,'1','护理等级','护理常规')," + _
                "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法',4,'中药用法','其它')," + _
                "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院','其它'),A.操作类型) As 项目特性,A.类别 As 类别ID,A.ID As 诊疗项目ID,nvl(执行频率,0) As 执行频率ID,nvl(计算方式,0) As 计算方式ID,nvl(执行安排,0) As 执行安排ID,nvl(计价性质,0) As 计价性质ID,nvl(执行科室,0) As 执行科室ID " + _
                "From 诊疗项目目录 A,诊疗项目别名 C Where A.ID=C.诊疗项目ID " + _
                "And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
                "And A.服务对象 IN([1],3) And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) IN (" + _
                IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
                " And Nvl(A.执行频率,0) IN(0,1)" + _
                " And A.ID=[2]"
            If rsInput.State = adStateOpen Then rsInput.Close: Set rsInput = New ADODB.Recordset
            Set rsInput = zldatabase.OpenSQLRecord(strSQL, Me.Caption, PatientType, Split(Split(strExtData, ";")(0), ",")(0))
            
            Call AdviceSet检查手术(3, strExtData)
            txt医嘱内容.Text = Get检查手术名称(2, "")
            strAdviceText = txt医嘱内容.Text & "(" & Split(strExtData, ";")(1) & ")"
            Me.txt附加 = Split(strExtData, ";")(1)
            str标本部位 = Me.txt附加
        End If
    Else
        str标本部位 = Trim(rsInput("标本部位"))
        txt医嘱内容.Text = txt医嘱内容.Text & "(" & str标本部位 & ")"
        strAdviceText = txt医嘱内容.Text
        
        '检查的组合部位行
        Call AdviceSet检查手术(1, "")
    End If
    
    '开始时间
    Me.txt开始时间 = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If rsInput("执行安排ID") = 1 Then
        Me.lbl开始时间.Visible = False: Me.chk开始时间.Visible = True: Me.chk开始时间.Value = 0
        Me.txt开始时间.Enabled = False
    Else
        Me.lbl开始时间.Visible = True: Me.chk开始时间.Visible = False
        Me.txt开始时间.Enabled = True
    End If
    
    '处理频率
    If rsInput("执行频率ID") = 1 Then
        Me.txt频率.Enabled = False: Me.txt频率 = "一次性": Me.cmd频率.Enabled = False
    Else
        Me.txt频率.Enabled = True: Me.txt频率 = "": Me.cmd频率.Enabled = True
    End If
    strSequence = Me.txt频率
    
    '总量
    Me.txt总量 = "1": Me.lbl总量单位.Caption = rsInput("计算单位")
    
    '单量
    If (rsInput("执行频率ID") = 0 And InStr(",1,2,", rsInput("计算方式ID")) > 0) _
                    Or InStr(",5,6,", rsInput("类别ID")) > 0 Then
        Me.txt单量.Enabled = True: Me.txt单量 = "": Me.txt单量.BackColor = Me.txt医嘱内容.BackColor: Me.lbl单量单位.Caption = rsInput("计算单位")
    Else
        Me.txt单量.Enabled = False: Me.txt单量 = "": Me.txt单量.BackColor = Me.BackColor: Me.lbl单量单位.Caption = "" ' rsInput("计算单位")
    End If
    
    '执行科室
    Set rsTmp = GetExeDepart(rsInput("ID"), PatientType, ItemDeptID)
    If rsTmp Is Nothing Then
        Me.cbo执行科室.Clear: Me.cbo执行科室.Enabled = False: Me.cbo执行科室.BackColor = Me.BackColor
    ElseIf rsTmp.RecordCount = 1 Then
        Me.cbo执行科室.Clear
        Me.cbo执行科室.AddItem rsTmp("名称"): Me.cbo执行科室.ItemData(0) = rsTmp("ID"): Me.cbo执行科室.ListIndex = 0
        Me.cbo执行科室.Enabled = False: Me.cbo执行科室.BackColor = Me.txt医嘱内容.BackColor
    Else
        Me.cbo执行科室.Clear
        Do While Not rsTmp.EOF
            Me.cbo执行科室.AddItem rsTmp("名称"): Me.cbo执行科室.ItemData(Me.cbo执行科室.ListCount - 1) = rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
        Me.cbo执行科室.ListIndex = 0
        Me.cbo执行科室.Enabled = True: Me.cbo执行科室.BackColor = Me.txt医嘱内容.BackColor
    End If
    
    '开嘱医生
    If Me.cbo医生.Text = "" Then Me.cbo医生.ListIndex = 0
    
    intType = intTmpType
    SetItemFormat '根据申请项目决定显示方式
    
    str类别 = rsInput("类别ID"): lngClinicID = rsInput("诊疗项目ID")
    int计价特性 = rsInput("计价性质ID"): int执行性质 = rsInput("执行科室ID"): strClinicName = IIf(intType = 4, Me.txt医嘱内容, rsInput("名称"))
    
    AdviceInput = True: Form_Resize
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LabsInput() As Boolean
'功能：编辑检验项目
'返回：本次录入是否有效
    Dim str过敏 As String, blnGroup As Boolean, i As Long
    Dim lng用法ID As Long, lngGroupRow As Long
    Dim lngPreRow As Long, lngNextRow As Long
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim intTmpType As Integer
    Dim strSQL As String, rsInput As New ADODB.Recordset

    On Error GoTo errH
    
    intTmpType = 4
    strHelpText = "检验项目"

    frmAdviceEditEx.mlngHwnd = Me.cbo执行科室.Hwnd ' txt医嘱内容.Hwnd
    frmAdviceEditEx.mintType = intTmpType
    frmAdviceEditEx.mint期效 = 1
    frmAdviceEditEx.mstr性别 = mstr性别
    frmAdviceEditEx.mlng项目id = 0 ' FileTypeID
    frmAdviceEditEx.mstrExtData = strExtData
    frmAdviceEditEx.mint服务对象 = PatientType

    On Error Resume Next
    frmAdviceEditEx.Show 1, Me
    On Error GoTo errH

    If Not frmAdviceEditEx.mblnOK Then Exit Function
    If frmAdviceEditEx.mstrExtData = "" Or Mid(frmAdviceEditEx.mstrExtData, 1, 1) = ";" Then Exit Function
    '获取采集方式
    Set rsTmp = SelectCap(Split(Split(frmAdviceEditEx.mstrExtData, ";")(0), ",")(0), , True)
    If rsTmp Is Nothing Then
        MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    Me.cmd采集.Tag = rsTmp("ID")
    Me.txt采集 = rsTmp("名称"): Me.txt采集.Tag = Me.txt采集
    
    strAdviceText = txt医嘱内容.Text
    strExtData = frmAdviceEditEx.mstrExtData

    strSQL = "Select Distinct A.ID,A.编码,A.名称,nvl(A.计算单位,'次') As 计算单位,nvl(A.标本部位,' ') As 标本部位," + _
        "Decode(A.类别,'H',Decode(A.操作类型,'1','护理等级','护理常规')," + _
        "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法',4,'中药用法','其它')," + _
        "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院','其它'),A.操作类型) As 项目特性,A.类别 As 类别ID,A.ID As 诊疗项目ID,nvl(执行频率,0) As 执行频率ID,nvl(计算方式,0) As 计算方式ID,nvl(执行安排,0) As 执行安排ID,nvl(计价性质,0) As 计价性质ID,nvl(执行科室,0) As 执行科室ID " + _
        "From 诊疗项目目录 A,诊疗项目别名 C Where A.ID=C.诊疗项目ID " + _
        "And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
        "And A.服务对象 IN([1],3) And Nvl(A.单独应用,0)=1 And Nvl(A.适用性别,0) IN (" + _
        IIf(Len(Trim(mstr性别)) = 0, "0)", IIf(mstr性别 Like "*男*", "1,0)", "2,0)")) + _
        " And Nvl(A.执行频率,0) IN(0,1)" + _
        " And A.ID=[2]"
    If rsInput.State = adStateOpen Then rsInput.Close: Set rsInput = New ADODB.Recordset
    Set rsInput = zldatabase.OpenSQLRecord(strSQL, Me.Caption, PatientType, Split(Split(strExtData, ";")(0), ",")(0))
    
    Call AdviceSet检查手术(3, strExtData)
    txt医嘱内容.Text = Get检查手术名称(2, "")
    strAdviceText = txt医嘱内容.Text & "(" & Split(strExtData, ";")(1) & ")"
    Me.txt附加 = Split(strExtData, ";")(1)
    str标本部位 = Me.txt附加
    
    '开始时间
    Me.txt开始时间 = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If rsInput("执行安排ID") = 1 Then
        Me.lbl开始时间.Visible = False: Me.chk开始时间.Visible = True: Me.chk开始时间.Value = 0
        Me.txt开始时间.Enabled = False
    Else
        Me.lbl开始时间.Visible = True: Me.chk开始时间.Visible = False
        Me.txt开始时间.Enabled = True
    End If
    
    '处理频率
    If rsInput("执行频率ID") = 1 Then
        Me.txt频率.Enabled = False: Me.txt频率 = "一次性": Me.cmd频率.Enabled = False
    Else
        Me.txt频率.Enabled = True: Me.txt频率 = "": Me.cmd频率.Enabled = True
    End If
    strSequence = Me.txt频率
    
    '总量
    Me.txt总量 = "1": Me.lbl总量单位.Caption = rsInput("计算单位")
    
    '单量
    If (rsInput("执行频率ID") = 0 And InStr(",1,2,", rsInput("计算方式ID")) > 0) _
                    Or InStr(",5,6,", rsInput("类别ID")) > 0 Then
        Me.txt单量.Enabled = True: Me.txt单量 = "": Me.txt单量.BackColor = Me.txt医嘱内容.BackColor: Me.lbl单量单位.Caption = rsInput("计算单位")
    Else
        Me.txt单量.Enabled = False: Me.txt单量 = "": Me.txt单量.BackColor = Me.BackColor: Me.lbl单量单位.Caption = "" ' rsInput("计算单位")
    End If
    
    '执行科室
    Set rsTmp = GetExeDepart(rsInput("ID"), PatientType, ItemDeptID)
    If rsTmp Is Nothing Then
        Me.cbo执行科室.Clear: Me.cbo执行科室.Enabled = False: Me.cbo执行科室.BackColor = Me.BackColor
    ElseIf rsTmp.RecordCount = 1 Then
        Me.cbo执行科室.Clear
        Me.cbo执行科室.AddItem rsTmp("名称"): Me.cbo执行科室.ItemData(0) = rsTmp("ID"): Me.cbo执行科室.ListIndex = 0
        Me.cbo执行科室.Enabled = False: Me.cbo执行科室.BackColor = Me.txt医嘱内容.BackColor
    Else
        Me.cbo执行科室.Clear
        Do While Not rsTmp.EOF
            Me.cbo执行科室.AddItem rsTmp("名称"): Me.cbo执行科室.ItemData(Me.cbo执行科室.ListCount - 1) = rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
        Me.cbo执行科室.ListIndex = 0
        Me.cbo执行科室.Enabled = True: Me.cbo执行科室.BackColor = Me.txt医嘱内容.BackColor
    End If
    
    '开嘱医生
    If Me.cbo医生.Text = "" Then Me.cbo医生.ListIndex = 0
    
    intType = intTmpType
    SetItemFormat '根据申请项目决定显示方式
    
    str类别 = rsInput("类别ID"): lngClinicID = rsInput("诊疗项目ID")
    int计价特性 = rsInput("计价性质ID"): int执行性质 = rsInput("执行科室ID"): strClinicName = IIf(intType = 4, Me.txt医嘱内容, rsInput("名称"))
    
    LabsInput = True: Form_Resize
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSet检查手术(ByVal int类型 As Integer, ByVal strDataIDs As String)
'功能：1.重新设置指定检查组合项目的部位行,用于新输入检查组合项目或修改部位
'      2.重新设置指定手术项目的附加手术及麻醉项目行,用于新输入手术项目或手术项目的附加手术及麻醉项目
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
'      strDataIDs=检查:包含检查部位信息,手术:包含附加手术及麻醉项目信息,其中可能没有附加手术和麻醉
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    '重新加入部位行或附加手术行及麻醉项目行
    If int类型 = 2 Then
        strDataIDs = Trim(Replace(strDataIDs, ";", ","))
        If Left(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 2)
        If Right(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 1, Len(strDataIDs) - 1)
    ElseIf int类型 = 3 Then
        '处理检验项目
        strDataIDs = Mid(strDataIDs, 1, InStr(strDataIDs, ";") - 1)
    End If
    
    If strDataIDs <> "" Then
        If Not rsRelativeAdvice Is Nothing Then
            rsRelativeAdvice.Close
        Else
            Set rsRelativeAdvice = New ADODB.Recordset
        End If
        strSQL = "Select ID,编码,名称,nvl(标本部位,' ') As 标本部位," + _
        "类别,nvl(计价性质,0) As 计价性质,nvl(执行科室,0) As 执行科室 From 诊疗项目目录 Where ID IN(" & strDataIDs & ")"
        OpenRecord rsRelativeAdvice, strSQL, Me.Caption
    Else
        If Not rsRelativeAdvice Is Nothing Then rsRelativeAdvice.Close: Set rsRelativeAdvice = Nothing
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get检查手术内容(ByVal int类型 As Integer, ByVal txtMainAdvice As String) As String
'功能：重新生成检查手术内容的医嘱内容
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
    Dim lngBegin As Long, i As Long
    Dim str麻醉 As String, strTmp As String
    Dim strDate As String
    
    strDate = IIf(Me.chk开始时间.Visible And Me.chk开始时间.Value = 0, "", Format(Me.txt开始时间, "yy年MM月dd日"))
    
    If rsRelativeAdvice Is Nothing Then
        If int类型 = 1 Then
            Get检查手术内容 = txtMainAdvice & IIf(Len(str标本部位) = 0, "", "(" & str标本部位 & ")"): Exit Function
        Else
            Get检查手术内容 = IIf(Len(strDate) = 0, "", strDate & " 行 ") & txtMainAdvice & IIf(Len(str标本部位) = 0, "", "(" & str标本部位 & ")"): Exit Function
        End If
    End If
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If int类型 = 1 Then
            If Len(Trim(rsRelativeAdvice("标本部位"))) > 0 Then
                strTmp = strTmp & "," & rsRelativeAdvice("标本部位")
            End If
        ElseIf Len(Trim(rsRelativeAdvice("名称"))) > 0 Then
            If rsRelativeAdvice("类别") = "G" Then
                str麻醉 = rsRelativeAdvice("名称")
            Else
                strTmp = strTmp & "," & rsRelativeAdvice("名称")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If int类型 = 1 Then
        If strTmp <> "" Then
            Get检查手术内容 = txtMainAdvice & "(" & Mid(strTmp, 2) & ")"
        Else
            Get检查手术内容 = txtMainAdvice
        End If
    Else
        If strTmp <> "" Or str麻醉 <> "" Then
            If str麻醉 <> "" Then
                Get检查手术内容 = IIf(Len(strDate) = 0, "", strDate & " ") & "在 " & str麻醉 & " 下行 " & txtMainAdvice
            Else
                Get检查手术内容 = IIf(Len(strDate) = 0, "", strDate & " 行 ") & txtMainAdvice
            End If
            If strTmp <> "" Then
                Get检查手术内容 = Get检查手术内容 & " 及 " & Mid(strTmp, 2)
            End If
        Else
            Get检查手术内容 = IIf(Len(strDate) = 0, "", strDate & " 行 ") & txtMainAdvice
        End If
    End If
End Function

Private Function Get检查手术名称(ByVal int类型 As Integer, ByVal txtMainAdvice As String) As String
'功能：重新生成检查手术内容的医嘱内容
'参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
    Dim lngBegin As Long, i As Long
    Dim str麻醉 As String, strTmp As String
    Dim strDate As String
    
    If rsRelativeAdvice Is Nothing Or int类型 = 1 Then Get检查手术名称 = txtMainAdvice: Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("名称"))) > 0 Then
            If rsRelativeAdvice("类别") <> "G" Then
                strTmp = strTmp & "," & rsRelativeAdvice("名称")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If strTmp <> "" Then
        Get检查手术名称 = IIf(Len(Trim(txtMainAdvice)) = 0, "", txtMainAdvice & " 及 ") & Mid(strTmp, 2)
    Else
        Get检查手术名称 = txtMainAdvice
    End If
End Function

Private Function Get麻醉名称() As String
    If rsRelativeAdvice Is Nothing Then Get麻醉名称 = "": Exit Function
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("名称"))) > 0 Then
            If rsRelativeAdvice("类别") = "G" Then
                Get麻醉名称 = rsRelativeAdvice("名称")
            End If
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
End Function

Private Function Get部位名称() As String
    If rsRelativeAdvice Is Nothing Then Get部位名称 = "": Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("标本部位"))) > 0 Then
            Get部位名称 = Get部位名称 & "," & rsRelativeAdvice("标本部位")
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    If Len(Get部位名称) > 0 Then Get部位名称 = Mid(Get部位名称, 2)
End Function

Private Function GetExeDepart(ByVal lngDiagItem As Long, ByVal iPatientType As Integer, Optional ByVal lngDepartID As Long = 0) As ADODB.Recordset
'功能：获取执行科室
'   iPatientType：病人类型 1=门诊、2=住院
'   lngDepartID：开单科室
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo DBError
    
    If lngDepartID = 0 Then lngDepartID = UserInfo.部门ID
    
    zldatabase.OpenRecordset rsTmp, "Select B.ID,B.编码,B.名称 From 部门表 B Where B.ID=" & lngDepartID & " Order by B.编码", Me.Caption
    
    If Not rsTmp.EOF Then Set GetExeDepart = rsTmp
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetGroupCount(lng组合ID As Long) As Long
'功能：获取组合项目中的项目数
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Count(*) as NUM From 诊疗项目组合 Where 诊疗组合ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng组合ID)
    If Not rsTmp.EOF Then GetGroupCount = zlCommFun.Nvl(rsTmp!NUM, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get缺省用法ID(int类型 As Integer) As Long
'功能：返回缺省的给药途径或中药煎法
'参数：int类型=2-给药途径,3-中药煎法,4-中药用法
'      str性别=病人性别
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ID From 诊疗项目目录" & _
        " Where 类别='E' And 操作类型=[1]" & _
        " And (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
        " Order by 编码"
    
    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "mdlCISWork", int类型)
    If Not rsTmp.EOF Then Get缺省用法ID = rsTmp!ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetItemField(ByVal lng项目ID As Long, ByVal strField As String) As Variant
'功能：获取指定诊疗项目的指定字段信息
'说明：未处理NULL值
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select " & strField & " From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng项目ID)
    If Not rsTmp.EOF Then GetItemField = rsTmp.Fields(strField).Value
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get开嘱医生(ByVal lng病人id As Long, ByVal bln护士站 As Boolean, str缺省医生 As String, lng医生ID As Long, _
    Optional objCbo As Object, Optional ByVal int范围 As Integer = 2) As Boolean
'功能：获取可用的开嘱医生在指定的下拉框中
'参数：lng病人科室ID=病人所在科室ID
'      bln护士站=是否由护士代医生下医嘱
'      objCbo=要加入医生清单的下拉框
'      str缺省医生=缺省定位的医生,如果不传objCbo,则先优先定位,再返回缺省医生和医生ID
'      int范围=1-门诊,2-住院(缺省)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
        
    On Error GoTo errH
    
    If bln护士站 Then
        '病人所在科室的医生
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & IIf(objCbo Is Nothing, ",B.部门ID", "") & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
            " And B.部门ID=" & lng病人科室ID & _
            " Order by A.简码"
        '病人所在病区各科的医生
        strSQL = "Select Distinct 病区ID From 床位状况记录 Where 科室ID=" & lng病人科室ID
        strSQL = "Select Distinct 科室ID From 床位状况记录 Where 病区ID=(" & strSQL & ")"
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & IIf(objCbo Is Nothing, ",B.部门ID", "") & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
            " And B.部门ID IN(" & strSQL & ")" & _
            " Order by A.简码"
        '全院住院科室的医生
        strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN(" & int范围 & ",3)"
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & IIf(objCbo Is Nothing, ",B.部门ID", "") & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
            " And B.部门ID IN(" & strSQL & ")" & _
            " Order by A.简码"
    Else '医生下医嘱时,限制为只能为医生本人
        strSQL = "Select ID,编号,姓名,简码 From 人员表 Where ID=" & UserInfo.ID
    End If

    OpenRecord rsTmp, strSQL, "zlCISCore"
    If objCbo Is Nothing Then
        If Not rsTmp.EOF Then
            If Not bln护士站 Then
                lng医生ID = rsTmp!ID
                str缺省医生 = rsTmp!姓名
            ElseIf bln护士站 Then
                If str缺省医生 <> "" Then
                    '缺省医生(住院医师)优先
                    rsTmp.Filter = "姓名='" & str缺省医生 & "'"
                Else
                    '病人科室的医生优先
                    rsTmp.Filter = "部门ID=" & lng病人科室ID
                End If
                If rsTmp.EOF Then rsTmp.Filter = 0
                lng医生ID = rsTmp!ID
                str缺省医生 = rsTmp!姓名
            End If
        End If
    Else
        objCbo.Clear
        For i = 1 To rsTmp.RecordCount
            objCbo.AddItem rsTmp!姓名
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
            If rsTmp!姓名 = str缺省医生 Then
                Call zlControl.CboSetIndex(objCbo.Hwnd, objCbo.NewIndex)
            End If
            rsTmp.MoveNext
        Next
    End If
    Get开嘱医生 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get开嘱科室ID(ByVal lng医生ID As Long, ByVal lng病人科室ID As Long, Optional ByVal int范围 As Integer = 2) As Long
'功能：由医生确定开嘱科室
'参数：int范围=1-门诊,2-住院(缺省)
'说明：在医生所属科室范围内,优先顺序如下：
'      1、病人科室
'      2、服务于门诊/住院病人的科室且为默认科室
'      3、服务于门诊/住院病人的科室
'      4、默认科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim arr科室ID(1 To 4) As Long
    
    '可能部门没有性质
    strSQL = "Select Distinct C.编码,A.部门ID,Nvl(A.缺省,0) as 缺省,Nvl(B.服务对象,0) as 服务对象" & _
        " From 部门人员 A,部门性质说明 B,部门表 C" & _
        " Where A.部门ID=C.ID And A.部门ID=B.部门ID(+) And A.人员ID=[1]" & _
        " Order by C.编码"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng医生ID)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!部门ID = lng病人科室ID Then
            arr科室ID(1) = rsTmp!部门ID
        ElseIf InStr("," & int范围 & ",3,", rsTmp!服务对象) > 0 And rsTmp!缺省 = 1 Then
            arr科室ID(2) = rsTmp!部门ID
        ElseIf InStr("," & int范围 & ",3,", rsTmp!服务对象) > 0 Then
            If arr科室ID(3) = 0 Then arr科室ID(3) = rsTmp!部门ID
        ElseIf rsTmp!缺省 = 1 Then
            arr科室ID(4) = rsTmp!部门ID
        End If
        rsTmp.MoveNext
    Next
    For i = LBound(arr科室ID) To UBound(arr科室ID)
        If arr科室ID(i) <> 0 Then
            Get开嘱科室ID = arr科室ID(i)
            Exit For
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'===以下为病人信息
Private Sub cbo费别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 And cbo费别.ListIndex <> -1 Then Call zlCommFun.PressKey(vbKeyTab)
    
    If SendMessage(cbo费别.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo费别.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo费别.ListIndex = lngIdx
    If cbo费别.ListIndex = -1 And cbo费别.ListCount > 0 Then cbo费别.ListIndex = 0
End Sub

Private Sub cbo付款方式_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo付款方式.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo付款方式.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo付款方式.ListIndex = lngIdx
End Sub

Private Sub cbo国籍_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo国籍.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo国籍.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo国籍.ListIndex = lngIdx
End Sub

Private Sub cbo婚姻_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo婚姻.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo婚姻.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo婚姻.ListIndex = lngIdx
End Sub

Private Sub cbo民族_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo民族.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo民族.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo民族.ListIndex = lngIdx
End Sub

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo性别.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo性别.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo性别.ListIndex = lngIdx
    If cbo性别.ListIndex = -1 And cbo性别.ListCount > 0 Then cbo性别.ListIndex = 0
End Sub

Private Sub cbo职业_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo职业.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo职业.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo职业.ListIndex = lngIdx
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmd单位名称_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = zldatabase.ShowSelect(Me, _
            " Select ID,上级ID,末级,编码,名称,地址,电话,开户银行,帐号,联系人 From  合约单位" & _
            " Start With 上级ID is NULL Connect by Prior ID=上级ID", _
            2, "单位", , txt单位名称.Text)
    If Not rsTmp Is Nothing Then
        txt单位名称.Tag = rsTmp!ID
        txt单位名称.Text = rsTmp!名称
        txt单位名称.SelStart = Len(txt单位名称.Text)
    End If
    txt单位名称.SetFocus
End Sub

Private Sub cmd家庭地址_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = zldatabase.ShowSelect(Me, _
            " Select Distinct Substr(名称,1,2) as ID,NULL as 上级ID,0 as 末级,NULL as 编码," & _
            " Substr(名称,1,2) as 名称 From 地区" & _
            " Union All" & _
            " Select 编码 as ID,Substr(名称,1,2) as 上级ID,1 as 末级,编码,名称 " & _
            " From 地区 Order by 编码", 2, "地区", , txt家庭地址.Text)
    If Not rsTmp Is Nothing Then
        txt家庭地址.Text = rsTmp!名称
        txt家庭地址.SelStart = Len(txt家庭地址.Text)
    End If
    txt家庭地址.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
        DoEvents
    ElseIf KeyCode = vbKeyPageDown Then
        CmdOk_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Function InitData() As Boolean
'功能：初始化必要数据
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    '性别
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("性别")
    cbo性别.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo性别.AddItem rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo性别.ItemData(cbo性别.NewIndex) = 1
                cbo性别.ListIndex = cbo性别.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    
    '费别
    Init费别 True

    '医疗付款方式
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("医疗付款方式")
    cbo付款方式.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo付款方式.AddItem rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo付款方式.ItemData(cbo付款方式.NewIndex) = 1
                cbo付款方式.ListIndex = cbo付款方式.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    '国籍
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("国籍")
    cbo国籍.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo国籍.AddItem rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo国籍.ItemData(cbo国籍.NewIndex) = 1
                cbo国籍.ListIndex = cbo国籍.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    '民族
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("民族")
    cbo民族.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo民族.AddItem rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo民族.ItemData(cbo民族.NewIndex) = 1
                cbo民族.ListIndex = cbo民族.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    '婚姻状况
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("婚姻状况")
    cbo婚姻.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo婚姻.AddItem rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo婚姻.ItemData(cbo婚姻.NewIndex) = 1
                cbo婚姻.ListIndex = cbo婚姻.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    '职业
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("职业")
    cbo职业.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo职业.AddItem rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo职业.ItemData(cbo职业.NewIndex) = 1
                cbo职业.ListIndex = cbo职业.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    
    '初始开单科室
    InitDepts
    
    InitData = True
End Function

Private Function Init费别(bln初诊 As Boolean, Optional blnKeepIndex As Boolean) As Boolean
'参数：bln初诊=是否允许仅限初诊的项目
'      blnKeepIndex=是否保持原有的费别选择
    Dim strSQL As String, i As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strKeep As String
    
    On Error GoTo errH
    
    strKeep = cbo费别.Text
    
    '费别:身份唯一性项目(包含了缺省费别),可以是初诊,不管有效期间及科室
    strSQL = "Select 编码,名称,简码," & _
        " Nvl(仅限初诊,0) as 初诊,Nvl(缺省标志,0) as 缺省" & _
        " From 费别 Where 属性=1" & IIf(Not bln初诊, " And Nvl(仅限初诊,0)=0", "") & _
        " Order by 编码"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, Me.Caption, strSQL) 'SQLTest
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    
    cbo费别.Clear
    Do While Not rsTmp.EOF
        cbo费别.AddItem rsTmp!名称
        If rsTmp!缺省 = 1 Then
            If cbo费别.ListIndex = -1 Then
                cbo费别.ItemData(cbo费别.NewIndex) = 1
                cbo费别.ListIndex = cbo费别.NewIndex
            End If
        End If
        
        '保持原有费别选择
        If blnKeepIndex Then
            If strKeep = rsTmp!编码 & "-" & rsTmp!名称 Then
                cbo费别.ListIndex = cbo费别.NewIndex
            End If
        End If
        
        '记录初诊项目:不会是本地缺省及系统缺省
        If rsTmp!初诊 = 1 Then
            cbo费别.ItemData(cbo费别.NewIndex) = 2
        End If
        rsTmp.MoveNext
    Loop
    
    Init费别 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt单位电话_GotFocus()
    zlControl.TxtSelAll txt单位电话
End Sub

Private Sub txt单位电话_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt单位电话, KeyAscii
End Sub

Private Sub txt单位名称_GotFocus()
    zlControl.TxtSelAll txt单位名称
    zlCommFun.OpenIme True
End Sub

Private Sub txt单位名称_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And cmd单位名称.Enabled And cmd单位名称.Visible Then cmd单位名称_Click
End Sub

Private Sub txt单位名称_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt单位名称, KeyAscii
End Sub

Private Sub txt单位名称_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt单位邮编_GotFocus()
    zlControl.TxtSelAll txt单位邮编
End Sub

Private Sub txt单位邮编_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt单位邮编, KeyAscii
End Sub

Private Sub txt家庭地址_GotFocus()
    zlControl.TxtSelAll txt家庭地址
    zlCommFun.OpenIme True
End Sub

Private Sub txt家庭地址_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And cmd家庭地址.Enabled And cmd家庭地址.Visible Then cmd家庭地址_Click
End Sub

Private Sub txt家庭地址_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt家庭地址, KeyAscii
End Sub

Private Sub txt家庭地址_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt家庭电话_GotFocus()
    zlControl.TxtSelAll txt家庭电话
End Sub

Private Sub txt家庭电话_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt家庭电话, KeyAscii
End Sub

Private Sub txt家庭邮编_GotFocus()
    zlControl.TxtSelAll txt家庭邮编
End Sub

Private Sub txt家庭邮编_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt家庭邮编, KeyAscii
End Sub

Private Sub txt门诊号_GotFocus()
    zlControl.TxtSelAll txt门诊号
End Sub

Private Sub txt门诊号_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt门诊号, KeyAscii
End Sub

Private Sub txt年龄_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt年龄.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt年龄.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt年龄_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt年龄.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt年龄_GotFocus()
    zlControl.TxtSelAll txt年龄
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt年龄, KeyAscii
End Sub

Private Sub txt身份证号_GotFocus()
    zlControl.TxtSelAll txt身份证号
End Sub

Private Sub txt身份证号_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt身份证号, KeyAscii
End Sub

Private Sub txt姓名_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt姓名.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt姓名.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt姓名_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt姓名.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
    zlCommFun.OpenIme True
End Sub

Private Sub txt姓名_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 13 Then
        KeyCode = Asc(UCase(Chr(KeyCode)))
        CheckLen txt姓名, KeyCode
    End If
End Sub

Private Function CombIndex(objComboBox As Object, ByVal strText As String) As Integer
    Dim i As Integer
    CombIndex = 0
    For i = 0 To objComboBox.ListCount - 1
        With objComboBox
            If .List(i) Like "*-" & strText Then CombIndex = i: Exit For
        End With
    Next
End Function

Private Sub txt姓名_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Function GetPatient(strCode As String) As ADODB.Recordset
'功能：读取病人信息，并显示该病人存在的医嘱时间
    Dim strSQL As String, i As Long
    Dim strNO As String, str姓名 As String, lng病人id As Long
    Dim strSeek As String
    
    On Error GoTo errH
    
    sCheckNo = ""
    strSeek = strCode
    '判断当前输入模式
    If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) And iInputType = -1 Then '刷卡
        iInputType = 0
    ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '病人ID
        iInputType = 1
        strSeek = Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then '住院号
        iInputType = 2
        strSeek = Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '门诊号
        iInputType = 3
        strSeek = Mid(strCode, 2)
    ElseIf Left(strCode, 1) = "G" Or Left(strCode, 1) = "." Then '挂号单
        iInputType = 4
        strSeek = Mid(strCode, 2)
    ElseIf Left(strCode, 1) = "/" Then '收费单据号
        iInputType = 5
        strSeek = Mid(strCode, 2)
    ElseIf iInputType = -1 Then '当作姓名
        iInputType = 6
    End If
    
    If iInputType = 0 Then '刷卡
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,Nvl(A.住院次数,0) As 主页ID,Decode(A.当前科室id,Null,Nvl(B.执行部门ID,0),A.当前科室id) As 病人科室,B.执行人 As 医生,A.*" & _
            " From 病人信息 A,病人挂号记录 B Where A.就诊卡号=[1] And A.病人ID=B.病人ID(+) And A.门诊号=B.门诊号(+)" & _
            " And (A.当前科室id IS NOT NULL Or NVL(B.执行状态,1) IN (0,2))"
    ElseIf iInputType = 1 Then '病人ID
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,Nvl(A.住院次数,0) As 主页ID,Nvl(A.当前科室id,0) As 病人科室,A.*" & _
            " From 病人信息 A Where A.病人ID=[2]"
    ElseIf iInputType = 2 Then '住院号
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,Nvl(A.住院次数,0) As 主页ID,Decode(A.当前科室id,Null,Nvl(B.入院科室ID,0),A.当前科室id) As 病人科室,B.住院医师 As 医生,A.*" & _
            " From 病人信息 A,病案主页 B Where A.住院号=[2] And A.病人ID=B.病人ID And A.当前科室id IS NOT NULL And B.出院日期 Is NULL"
    ElseIf iInputType = 3 Then '门诊号
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,Nvl(A.住院次数,0) As 主页ID,Decode(A.当前科室id,Null,Nvl(B.执行部门ID,0),A.当前科室id) As 病人科室,B.执行人 As 医生,A.*" & _
            " From 病人信息 A,病人挂号记录 B Where A.门诊号=[2] And A.病人ID=B.病人ID(+) And A.门诊号=B.门诊号(+)" & _
            " And (A.当前科室id IS NOT NULL Or NVL(B.执行状态,1) IN (0,2))"
    ElseIf iInputType = 4 Then '挂号单
        strNO = GetFullNO(strSeek, 12)
        strSQL = "Select Decode(B.主页ID,Null,1,2) As PatientType,Nvl(B.主页ID,0) As 主页ID,Nvl(B.执行部门ID,0) As 病人科室,B.执行人 As 医生,A.*" & _
            " From 病人信息 A,病人费用记录 B" & _
            " Where B.记录性质=4 And B.记录状态 IN(1,3) And B.NO=[3] And B.病人ID=A.病人ID"
    ElseIf iInputType = 5 Then '收费单据号
        strNO = GetFullNO(strSeek, 13)
        sCheckNo = strNO
        
        strSQL = "Select Decode(B.主页ID,Null,1,2) As PatientType,Nvl(B.主页ID,0) As 主页ID,B.开单部门ID As 病人科室,B.开单人 As 医生,B.姓名,B.性别,B.年龄," & _
            "A.病人ID,A.单位电话,A.工作单位,A.单位邮编,A.家庭地址,A.家庭电话,A.户口邮编,A.门诊号,A.身份证号,A.费别,A.医疗付款方式," & _
            "A.国籍,A.婚姻状况,A.民族,A.职业 From 病人信息 A,病人费用记录 B" & _
            " Where B.记录性质=1 And B.记录状态 IN(1,3) And B.NO=[3] And B.病人ID=A.病人ID(+) And B.医嘱序号 Is Null"
    Else '当作姓名
        strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,Nvl(A.住院次数,0) As 主页ID,Nvl(A.当前科室id,0) As 病人科室,A.*" & _
            " From 病人信息 A Where A.姓名=[1]"
    End If
    
    Set GetPatient = zldatabase.OpenSQLRecord(strSQL, Me.Caption, strCode, Val(strSeek), strNO)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select 编号规则,Sysdate as 日期 From 号码控制表 Where 项目序号=" & intNum
    Call OpenRecord(rsTmp, strSQL, "mdlPublic")
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!编号规则, 0)
        curDate = rsTmp!日期
    End If

    If intType = 1 Then
        '按日编号
        strSQL = Format(CDate("1992-" & Format(rsTmp!日期, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '按年编号
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zldatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Private Function GetDictData(strDict As String) As ADODB.Recordset
'功能：从指定的字典中读取数据
'参数：strDict=字典对应的表名
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, "mdlRegEvent", strSQL) 'SQLTest
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    
    If Not rsTmp.EOF Then Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitDepts() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where B.部门ID = A.ID " & _
        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
        " And (B.工作性质 IN('临床','体检') Or A.ID=" & ItemDeptID & " Or A.ID=" & UserInfo.部门ID & ")" & _
        " Order by A.编码"
    Me.cbo开单科室.Clear
    
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    For i = 1 To rsTmp.RecordCount
        cbo开单科室.AddItem rsTmp!名称
        cbo开单科室.ItemData(cbo开单科室.NewIndex) = rsTmp!ID
        
        rsTmp.MoveNext
    Next
    If cbo开单科室.ListCount > 0 Then cbo开单科室.ListIndex = 0
    InitDepts = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitDoctors(ByVal lng科室ID As Long)
'功能：读取当前开单科室中包含的所有人员
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    Me.cbo医生.Clear
    
    '科室医生或护士
    strSQL = _
        "Select Distinct A.ID,B.部门ID,A.编号,A.姓名,Upper(A.简码) as 简码," & _
        " C.人员性质,Nvl(A.聘任技术职务,0) as 职务" & _
        " From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
        " And C.人员性质 IN('医生') And B.部门ID=[1]"
    strSQL = strSQL & " Order by 简码,人员性质 Desc"
    
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng科室ID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo医生.AddItem rsTmp!姓名
            cbo医生.ItemData(cbo医生.ListCount - 1) = rsTmp!部门ID
            
            If rsTmp!ID = UserInfo.ID And cbo医生.ListIndex = -1 Then cbo医生.ListIndex = cbo医生.NewIndex
            rsTmp.MoveNext
        Next
        
        If cbo医生.ListCount = 1 And cbo医生.ListIndex = -1 Then cbo医生.ListIndex = 0
    End If
End Sub

Private Sub vsf2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strPh As String, strMsg As String
        
    If vsf2.RowData(Row) = -1 And Col = 2 Then
        '手工标本号
        If gblnManualPH Then
            strPh = ValidPH(vsf2.TextMatrix(Row, Col), strMsg)
            If Len(strMsg) > 0 Then
                MsgBox strMsg, vbOKOnly + vbInformation, gstrSysName
                vsf2.TextMatrix(Row, Col) = ""
            Else
                vsf2.TextMatrix(Row, Col) = strPh
            End If
        End If
    End If
End Sub

Private Sub vsf2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsf2.Col = 2
End Sub

Private Sub vsf2_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf2_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf2_GotFocus()
    vsf2.Col = 2
End Sub

Private Sub vsf2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Col
    Case 2
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'KeyAscii = FilterKeyAscii(KeyAscii, 99, "ZXCVBNMASDFGHJKLQWERTYUIOP01234567890,-")
        If vsf2.RowData(vsf2.Row) <> -1 Then
            KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
        Else
            '手工标本号
            If gblnManualPH Then
                KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789-")
            Else
                KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
            End If
        End If
    End Select
End Sub

