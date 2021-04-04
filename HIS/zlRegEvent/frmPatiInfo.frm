VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "*\A..\ZlPatiAddress\ZlPatiAddress.vbp"
Begin VB.Form frmPatiInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "挂号病人信息"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatiInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picTaskPanelOther 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   8190
      ScaleHeight     =   825
      ScaleWidth      =   1755
      TabIndex        =   103
      Top             =   7440
      Visible         =   0   'False
      Width           =   1755
      Begin XtremeSuiteControls.TaskPanel wndTaskPanelOther 
         Height          =   435
         Left            =   330
         TabIndex        =   104
         Top             =   150
         Width           =   855
         _Version        =   589884
         _ExtentX        =   1508
         _ExtentY        =   767
         _StockProps     =   64
         VisualTheme     =   7
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   2010
      Top             =   7620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   420
      Left            =   90
      TabIndex        =   53
      Top             =   7815
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "返回(&X)"
      Height          =   420
      Left            =   6450
      TabIndex        =   51
      ToolTipText     =   "热键：F2"
      Top             =   7785
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   4875
      TabIndex        =   52
      Top             =   7815
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picCard 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   11610
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   0
      Width           =   11610
      Begin VB.TextBox txt验证 
         Enabled         =   0   'False
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   6375
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   120
         Width           =   1725
      End
      Begin VB.TextBox txt密码 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3795
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   120
         Width           =   1725
      End
      Begin VB.TextBox txt卡号 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1230
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   105
         Width           =   1725
      End
      Begin VB.Label lbl验证 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "验证"
         Height          =   240
         Left            =   5790
         TabIndex        =   75
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lbl密码 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
         Height          =   240
         Left            =   3210
         TabIndex        =   74
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lblICCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "卡号"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   615
         TabIndex        =   73
         Top             =   150
         Width           =   510
      End
   End
   Begin VB.PictureBox picInfo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6885
      Left            =   60
      ScaleHeight     =   6885
      ScaleWidth      =   11490
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   975
      Width           =   11490
      Begin VB.TextBox txtMobile 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   8670
         MaxLength       =   20
         TabIndex        =   32
         Top             =   4110
         Width           =   1890
      End
      Begin ZlPatiAddress.PatiAddress padd家庭地址 
         Height          =   360
         Left            =   1170
         TabIndex        =   18
         Tag             =   "现住址"
         Top             =   2100
         Visible         =   0   'False
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   100
      End
      Begin ZlPatiAddress.PatiAddress padd户口地址 
         Height          =   360
         Left            =   1170
         TabIndex        =   21
         Tag             =   "户口地址"
         Top             =   2505
         Visible         =   0   'False
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   100
      End
      Begin VB.Frame fraUnit 
         Caption         =   "单位信息"
         Height          =   750
         Left            =   30
         TabIndex        =   115
         Top             =   5250
         Width           =   11415
         Begin VB.CommandButton cmd单位名称 
            Caption         =   "…"
            Height          =   360
            Left            =   5520
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   270
            Width           =   360
         End
         Begin VB.TextBox txt单位名称 
            Height          =   360
            Left            =   660
            MaxLength       =   100
            TabIndex        =   39
            Top             =   270
            Width           =   4860
         End
         Begin VB.TextBox txt单位邮编 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   6540
            MaxLength       =   6
            TabIndex        =   40
            Top             =   270
            Width           =   1680
         End
         Begin VB.TextBox txt单位电话 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   9120
            MaxLength       =   20
            TabIndex        =   41
            Top             =   270
            Width           =   2205
         End
         Begin VB.Label lbl单位名称 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "名称"
            Height          =   240
            Left            =   135
            TabIndex        =   119
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl单位邮编 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "邮编"
            Height          =   240
            Left            =   6015
            TabIndex        =   117
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl单位电话 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "电话"
            Height          =   240
            Left            =   8580
            TabIndex        =   116
            Top             =   330
            Width           =   480
         End
      End
      Begin VB.Frame fraContact 
         Caption         =   "联系人信息"
         Height          =   720
         Left            =   30
         TabIndex        =   110
         Top             =   4500
         Width           =   11415
         Begin VB.TextBox txt联系人身份证 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   9120
            MaxLength       =   18
            TabIndex        =   38
            Top             =   270
            Width           =   2205
         End
         Begin VB.TextBox txt联系人电话 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   3450
            MaxLength       =   18
            TabIndex        =   35
            Top             =   270
            Width           =   1590
         End
         Begin VB.TextBox txt联系人姓名 
            Height          =   360
            Left            =   630
            MaxLength       =   64
            TabIndex        =   34
            Top             =   270
            Width           =   2160
         End
         Begin VB.TextBox txt其他关系 
            Height          =   360
            Left            =   6975
            MaxLength       =   30
            TabIndex        =   37
            Top             =   270
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.ComboBox cbo联系人关系 
            Height          =   360
            Left            =   5790
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   270
            Width           =   2445
         End
         Begin VB.Label lbl联系人身份证 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "身份证"
            Height          =   240
            Left            =   8355
            TabIndex        =   114
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lbl联系人姓名 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名"
            Height          =   240
            Left            =   135
            TabIndex        =   113
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl联系人电话 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "电话"
            Height          =   240
            Left            =   2925
            TabIndex        =   112
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl联系人关系 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "关系"
            Height          =   240
            Left            =   5250
            TabIndex        =   111
            Top             =   330
            Width           =   480
         End
      End
      Begin VB.TextBox txt户口地址邮编 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   10140
         MaxLength       =   6
         TabIndex        =   22
         Top             =   2504
         Width           =   1290
      End
      Begin VB.CommandButton cmdPicCollect 
         Caption         =   "采集"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9855
         TabIndex        =   108
         Top             =   1665
         Width           =   600
      End
      Begin VB.CommandButton cmdPicFile 
         Caption         =   "文件"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9210
         TabIndex        =   107
         Top             =   1665
         Width           =   585
      End
      Begin VB.CommandButton cmdPicClear 
         Caption         =   "清除"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   10485
         TabIndex        =   106
         Top             =   1665
         Width           =   600
      End
      Begin VB.PictureBox picPatient 
         Height          =   1620
         Left            =   9090
         ScaleHeight     =   1560
         ScaleWidth      =   2025
         TabIndex        =   105
         Top             =   20
         Width           =   2085
         Begin VB.Image imgPatient 
            Height          =   1545
            Left            =   15
            Stretch         =   -1  'True
            Top             =   15
            Width           =   2010
         End
      End
      Begin VB.CommandButton cmdRegLocation 
         Caption         =   "…"
         Height          =   360
         Left            =   8070
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   2504
         Width           =   360
      End
      Begin VB.CommandButton cmdBirthLocation 
         Caption         =   "…"
         Height          =   360
         Left            =   7080
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox txtBirthLocation 
         Height          =   360
         Left            =   1125
         MaxLength       =   100
         TabIndex        =   29
         Top             =   3720
         Width           =   5955
      End
      Begin VB.TextBox txt监护人 
         Height          =   360
         IMEMode         =   2  'OFF
         Left            =   8670
         MaxLength       =   20
         TabIndex        =   30
         Top             =   3720
         Width           =   2775
      End
      Begin VB.TextBox txt过敏反应 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4230
         MaxLength       =   200
         TabIndex        =   84
         Top             =   6645
         Visible         =   0   'False
         Width           =   990
      End
      Begin XtremeSuiteControls.TaskPanel TaskPanel1 
         Height          =   30
         Left            =   1680
         TabIndex        =   83
         Top             =   375
         Width           =   30
         _Version        =   589884
         _ExtentX        =   53
         _ExtentY        =   53
         _StockProps     =   64
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
      Begin VB.TextBox txt验证密码 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5790
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   1277
         Width           =   2895
      End
      Begin VB.TextBox txt支付密码 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1170
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1277
         Width           =   2895
      End
      Begin VB.CommandButton cmd区域 
         Caption         =   "…"
         Height          =   360
         Left            =   7080
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "热键：F3"
         Top             =   4110
         Width           =   375
      End
      Begin VB.ComboBox cbo付款方式 
         Height          =   360
         ItemData        =   "frmPatiInfo.frx":0E42
         Left            =   8670
         List            =   "frmPatiInfo.frx":0E44
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3322
         Width           =   2775
      End
      Begin VB.ComboBox cbo费别 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4710
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3322
         Width           =   2775
      End
      Begin VB.TextBox txtPatiMCNO 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   5790
         MaxLength       =   30
         TabIndex        =   16
         Top             =   1686
         Width           =   2895
      End
      Begin VB.TextBox txtPatiMCNO 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1170
         MaxLength       =   30
         TabIndex        =   15
         Top             =   1686
         Width           =   2895
      End
      Begin VB.ComboBox cbo年龄单位 
         Height          =   360
         Left            =   7920
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   465
         Width           =   780
      End
      Begin VB.CheckBox chk复诊 
         Caption         =   "复诊"
         Height          =   240
         Left            =   10695
         TabIndex        =   33
         Top             =   4185
         Width           =   795
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   1515
         Left            =   2850
         TabIndex        =   76
         Top             =   6735
         Visible         =   0   'False
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   2672
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.TextBox txt门诊号 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5790
         MaxLength       =   18
         TabIndex        =   5
         Top             =   50
         Width           =   2895
      End
      Begin VB.TextBox txtPatient 
         Height          =   360
         Left            =   1170
         MaxLength       =   100
         TabIndex        =   4
         Top             =   50
         Width           =   2895
      End
      Begin VB.TextBox txt年龄 
         Height          =   360
         IMEMode         =   2  'OFF
         Left            =   7185
         MaxLength       =   5
         TabIndex        =   9
         Top             =   465
         Width           =   690
      End
      Begin VB.ComboBox cbo性别 
         Height          =   360
         IMEMode         =   3  'DISABLE
         ItemData        =   "frmPatiInfo.frx":0E46
         Left            =   1170
         List            =   "frmPatiInfo.frx":0E48
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   459
         Width           =   825
      End
      Begin VB.ComboBox cbo国籍 
         Height          =   360
         Left            =   4710
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2913
         Width           =   2775
      End
      Begin VB.ComboBox cbo民族 
         Height          =   360
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2913
         Width           =   2775
      End
      Begin VB.ComboBox cbo婚姻 
         Height          =   360
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3322
         Width           =   2775
      End
      Begin VB.ComboBox cbo职业 
         Height          =   360
         Left            =   8670
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2913
         Width           =   2775
      End
      Begin VB.TextBox txt身份证号 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1170
         MaxLength       =   18
         TabIndex        =   11
         Top             =   868
         Width           =   2895
      End
      Begin VB.TextBox txt家庭电话 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5790
         MaxLength       =   20
         TabIndex        =   12
         Top             =   855
         Width           =   2895
      End
      Begin VB.TextBox txt家庭邮编 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   10140
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2095
         Width           =   1290
      End
      Begin VB.CommandButton cmd家庭地址 
         Caption         =   "…"
         Height          =   360
         Left            =   8070
         TabIndex        =   0
         ToolTipText     =   "热键F3"
         Top             =   2085
         Width           =   360
      End
      Begin VB.CommandButton cmd过敏 
         Caption         =   "…"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6090
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "热键:F3"
         Top             =   6540
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txt过敏 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1995
         MaxLength       =   50
         TabIndex        =   56
         Top             =   7980
         Visible         =   0   'False
         Width           =   990
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh过敏 
         Height          =   1215
         Left            =   30
         TabIndex        =   42
         ToolTipText     =   "F4:修改,F3:选择"
         Top             =   6135
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2143
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         FormatString    =   "<过敏药物                            "
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   8100
         Top             =   6450
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPatiInfo.frx":0E4A
               Key             =   "Itemps"
               Object.Tag             =   "Itemgm"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPatiInfo.frx":13E4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox txt出生时间 
         Height          =   360
         Left            =   5445
         TabIndex        =   8
         Top             =   465
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt出生日期 
         Height          =   360
         Left            =   3420
         TabIndex        =   7
         Top             =   465
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "YYYY-MM-DD"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt区域 
         Height          =   360
         Left            =   1125
         MaxLength       =   50
         TabIndex        =   31
         Top             =   4110
         Width           =   5955
      End
      Begin VB.TextBox txtRegLocation 
         Height          =   360
         Left            =   1170
         MaxLength       =   100
         TabIndex        =   20
         Top             =   2504
         Width           =   6900
      End
      Begin VB.ComboBox cbo家庭地址 
         Height          =   360
         Left            =   1170
         TabIndex        =   17
         Top             =   2100
         Width           =   6915
      End
      Begin VB.Label lblMobile 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手机号"
         Height          =   240
         Left            =   7920
         TabIndex        =   123
         Top             =   4170
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间"
         Height          =   240
         Left            =   4875
         TabIndex        =   120
         Top             =   510
         Width           =   480
      End
      Begin VB.Label lbl户口地址邮编 
         Alignment       =   1  'Right Justify
         Caption         =   "户口邮编"
         Height          =   240
         Left            =   8595
         TabIndex        =   109
         Top             =   2564
         Width           =   1515
      End
      Begin VB.Label lblRegLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口地址"
         Height          =   240
         Left            =   150
         TabIndex        =   101
         Top             =   2564
         Width           =   960
      End
      Begin VB.Label lblBirthLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点"
         Height          =   240
         Left            =   150
         TabIndex        =   100
         Top             =   3780
         Width           =   960
      End
      Begin VB.Label lbl监护人 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  监护人"
         Height          =   240
         Left            =   7680
         TabIndex        =   98
         Top             =   3780
         Width           =   960
      End
      Begin VB.Label lbl验证密码 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "验证密码"
         Height          =   240
         Left            =   4800
         TabIndex        =   82
         Top             =   1335
         Width           =   960
      End
      Begin VB.Label lbl支付密码 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "支付密码"
         Height          =   240
         Left            =   150
         TabIndex        =   81
         Top             =   1337
         Width           =   960
      End
      Begin VB.Label lbl区域 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "区域"
         Height          =   240
         Left            =   630
         TabIndex        =   54
         Top             =   4170
         Width           =   480
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   240
         Left            =   4170
         TabIndex        =   80
         Top             =   3375
         Width           =   480
      End
      Begin VB.Label lbl付款方式 
         BackStyle       =   0  'Transparent
         Caption         =   "付款方式"
         Height          =   300
         Left            =   7680
         TabIndex        =   60
         Top             =   3352
         Width           =   960
      End
      Begin VB.Label lblPatiMCNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "验证医保号"
         Height          =   240
         Index           =   1
         Left            =   4560
         TabIndex        =   79
         Top             =   1740
         Width           =   1200
      End
      Begin VB.Label lblPatiMCNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         Height          =   240
         Index           =   0
         Left            =   390
         TabIndex        =   78
         Top             =   1746
         Width           =   720
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         Height          =   240
         Left            =   2430
         TabIndex        =   77
         Top             =   525
         Width           =   960
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   -150
         X2              =   7695
         Y1              =   7785
         Y2              =   7785
      End
      Begin VB.Label lbl门诊号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   240
         Left            =   5040
         TabIndex        =   72
         Top             =   105
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   240
         Left            =   630
         TabIndex        =   71
         Top             =   110
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   240
         Left            =   630
         TabIndex        =   70
         Top             =   519
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   240
         Left            =   6660
         TabIndex        =   69
         Top             =   525
         Width           =   480
      End
      Begin VB.Label lbl婚姻 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   240
         Left            =   150
         TabIndex        =   68
         Top             =   3382
         Width           =   960
      End
      Begin VB.Label lbl职业 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   240
         Left            =   8160
         TabIndex        =   67
         Top             =   2970
         Width           =   480
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         Height          =   240
         Left            =   660
         TabIndex        =   66
         Top             =   2973
         Width           =   480
      End
      Begin VB.Label lbl国籍 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍"
         Height          =   240
         Left            =   4170
         TabIndex        =   65
         Top             =   2970
         Width           =   480
      End
      Begin VB.Label lbl身份证 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   240
         Left            =   150
         TabIndex        =   64
         Top             =   930
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "现住址"
         Height          =   240
         Left            =   390
         TabIndex        =   63
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label lbl家庭电话 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "电话"
         Height          =   240
         Left            =   5280
         TabIndex        =   62
         Top             =   915
         Width           =   480
      End
      Begin VB.Label lbl家庭邮编 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "现住址邮编"
         Height          =   240
         Left            =   8910
         TabIndex        =   61
         Top             =   2160
         Width           =   1200
      End
   End
   Begin XtremeSuiteControls.TabControl tbcPage 
      Height          =   6780
      Left            =   30
      TabIndex        =   85
      Top             =   570
      Width           =   10395
      _Version        =   589884
      _ExtentX        =   18336
      _ExtentY        =   11959
      _StockProps     =   64
   End
   Begin VB.PictureBox PicHealth 
      BorderStyle     =   0  'None
      Height          =   7230
      Left            =   120
      ScaleHeight     =   7230
      ScaleWidth      =   11400
      TabIndex        =   86
      Top             =   990
      Width           =   11400
      Begin VB.Frame fraCertificate 
         Height          =   105
         Left            =   1020
         TabIndex        =   122
         Top             =   2535
         Width           =   10335
      End
      Begin VB.CommandButton cmdMedicalWarning 
         Caption         =   "…"
         Height          =   330
         Left            =   10995
         TabIndex        =   97
         Top             =   135
         Width           =   330
      End
      Begin VB.ComboBox cboBloodType 
         Height          =   360
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   120
         Width           =   1410
      End
      Begin VB.ComboBox cboBH 
         Height          =   360
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   120
         Width           =   1410
      End
      Begin VB.TextBox txtMedicalWarning 
         Height          =   360
         Left            =   6135
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   120
         Width           =   4860
      End
      Begin VB.TextBox txtOtherWaring 
         Height          =   360
         Left            =   1725
         MaxLength       =   100
         TabIndex        =   46
         Top             =   525
         Width           =   9630
      End
      Begin VB.Frame frameLinkMan 
         Height          =   105
         Left            =   1320
         TabIndex        =   89
         Top             =   1020
         Width           =   10020
      End
      Begin VB.Frame Frame1 
         Height          =   105
         Left            =   1050
         TabIndex        =   88
         Top             =   5370
         Width           =   10275
      End
      Begin VB.Frame Frame2 
         Height          =   105
         Left            =   1050
         TabIndex        =   87
         Top             =   3930
         Width           =   10290
      End
      Begin VSFlex8Ctl.VSFlexGrid vsLinkMan 
         Height          =   975
         Left            =   30
         TabIndex        =   47
         Top             =   1320
         Width           =   11310
         _cx             =   19950
         _cy             =   1720
         Appearance      =   1
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsOtherInfo 
         Height          =   3195
         Left            =   15
         TabIndex        =   50
         Top             =   5640
         Width           =   11310
         _cx             =   19950
         _cy             =   5636
         Appearance      =   1
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsInoculate 
         Height          =   975
         Left            =   45
         TabIndex        =   49
         Top             =   4185
         Width           =   11310
         _cx             =   19950
         _cy             =   1720
         Appearance      =   1
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCertificate 
         Height          =   975
         Left            =   30
         TabIndex        =   48
         Top             =   2775
         Width           =   11310
         _cx             =   19950
         _cy             =   1720
         Appearance      =   1
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblCertificate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "证件信息"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -405
         TabIndex        =   121
         Top             =   2445
         Width           =   1860
      End
      Begin VB.Label lblBloodType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "血型"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   870
         TabIndex        =   96
         Top             =   150
         Width           =   1020
      End
      Begin VB.Label lblRH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "RH"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2940
         TabIndex        =   95
         Top             =   173
         Width           =   885
      End
      Begin VB.Label lblMedicalWarning 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "医学警示"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4665
         TabIndex        =   94
         Top             =   173
         Width           =   1860
      End
      Begin VB.Label lblOtherWaring 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "其他医学警示"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   15
         TabIndex        =   93
         Top             =   585
         Width           =   1875
      End
      Begin VB.Label lblLinkman 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "联系人信息"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -300
         TabIndex        =   92
         Top             =   945
         Width           =   1860
      End
      Begin VB.Label lblOtherInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "其他信息"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   -420
         TabIndex        =   91
         Top             =   5325
         Width           =   1860
      End
      Begin VB.Label lblInoculate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "接种情况"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -420
         TabIndex        =   90
         Top             =   3870
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frmPatiInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mbytFun As Byte  '0-编辑或查看病人信息,1-就诊卡发放,或绑定就诊卡
Public mstrCard As String '入:卡号,出成功:卡号,绑定就诊卡时不传入
Public mblnChange As Boolean
Public mblnInRange As Boolean '一卡通模式刷卡时,卡号是否在领用批次范围内
Public mrs家庭地址 As ADODB.Recordset  '缓存家庭地址,初始时读取地区表
Public mlngOutModeMC As Long '本地医保设置的外挂式医保险类
Public mrsBaseDict As ADODB.Recordset '国籍,民族,婚姻状况,职业
Public mintNOLength As Integer '门诊号长度
Public mbln发卡 As Boolean '问题号:56599
Public mstrPrivs As String
Public mlngModul As Long
Public mstr年龄 As String '原年龄
Public mstr性别 As String '原性别
Public mstr姓名 As String '原姓名
Public mstr年龄单位 As String
Public mstr出生日期 As String
Public mstr出生时间 As String
Public mstr身份证号 As String
Public mstrFirstCode As String '第一种证件类型的编码
Private mbln基本信息调整 As Boolean '是否允许调整病人基本信息
Private mblnCancel As Boolean
Private mlng磁卡领用ID As Long
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private mblnStructAdress As Boolean  '病人地址结构化录入
Private mblnShowTown As Boolean      '乡镇地址结构化录入
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private mfrmMain As Object
Private mDateSys As Date
Private mblnCheckNOValidity As Boolean
Private mobjKeyboard As Object
Private mbln扫描身份证 As Boolean '判断病人信息是否是通过扫描身份证得到
Private mbln禁止输入年龄 As Boolean
Public mrsPatiInfo As Recordset '当前His病人信息
Public mlng病人ID As Long '当前病人ID   :        '35233
Private mbln家庭地址输入    As Boolean      '家庭地址输入是否联想
Public Event PatiMerged(病人ID As Long)     '病人合并触发该事件
Private mbln扫描身份证签约 As Boolean
Public mbln监护人录入 As Boolean
Public mlng监护人年龄 As Long
Private mintDefaultBlood As Integer '默认血型序号
Private Enum mPageIndex
    基本 = 1
    健康档案 = 2
    附加信息 = 3
End Enum
Private mdic医疗卡属性 As New Dictionary '问题号56599
Private Const C_InoculateHeader = "接种日期,4,2400,1;接种名称,4,2400,1;接种日期,4,2400,1;接种名称,4,2400,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_LinkManColumHeader = "姓名,4,1200,1;关系,4,2400,1;身份证号,4,2400,1;电话,4,1200,1;附加信息,4,2400,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_OtherInfoColumHeader = "信息名,4,2400,1;信息值,4,2400,1;信息名,4,2400,1;信息值,4,2400,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_CertificateHeader = "证件类型,4,2400,1;证件号码,4,2400,1;证件类型,4,2400,1;证件号码,4,2400,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
'Private Const C_血型 = "A型,B型,O型,AB型,不详"
Private Const C_BH = "阴,阳,不详,未查"
Public Event ReturnVisitClick()     '点击复诊复选框改变对应的费别显示
Public mlngPlugInHwnd As Long
Public mblnPlugin As Boolean '插建是否创建成功
Public mrsEMPIOut As ADODB.Recordset 'EMPI返回的数据
Public mstrPlugChange As String

'74430,冉俊明,2014-7-7,挂号中的病人信息编辑功能中提供采集照片功能
Private mstr采集图片 As String '采集图片本地保存路径
Public mlng图像操作 As Long '指明当前对病人图像操作的类型(1-文件 2-采集 3-清除 4-身份证提取)
Private mstrIDImageFile As String
Public mblnSavePati As Boolean '病人照片信息或附加信息是否已保存
Public mobjPubPatient As Object
Public mblnNewPatient As Boolean
Private mblnNameChange As Boolean
Public Event 付款方式Click(index As Long)     '点击付款方式
Public mstrPriceGrade As String

Private mobjProPati As Collection '在挂号前，保存病人信息集
Private mblnGetBirth As Boolean '判断是否允许通过年龄计算生日

Private Sub cbo付款方式_Click()
    RaiseEvent 付款方式Click(cbo付款方式.ListIndex)
End Sub

Private Sub cbo国籍_Change()
    mstrPlugChange = mstrPlugChange & ",国籍"
End Sub

Private Sub cbo婚姻_Change()
    mstrPlugChange = mstrPlugChange & ",婚姻状况"
End Sub

Private Sub cbo家庭地址_Change()
    If Not mblnStructAdress Then mstrPlugChange = mstrPlugChange & ",现住址"
End Sub

Private Sub cbo家庭地址_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub
Private Sub cbo家庭地址_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub cbo家庭地址_KeyDown(KeyCode As Integer, Shift As Integer)
    '此过程处理本机缓存数据的删除,以及按下拉键时弹出下拉列表
    '下拉列表弹出时,如果按下删除键时,则删除缓存记录
    
    Dim str家庭地址 As String
    
    If KeyCode = vbKeyDelete Then
        str家庭地址 = cbo家庭地址.Text
        If Not mrs家庭地址 Is Nothing And mbln家庭地址输入 Then
            If mrs家庭地址.State = 1 And str家庭地址 <> "" Then
                If cbo家庭地址.SelText = str家庭地址 And SendMessage(cbo家庭地址.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = True Then
                    mrs家庭地址.Filter = "名称='" & str家庭地址 & "'"
                    If Not mrs家庭地址.EOF Then
                        mrs家庭地址.Delete adAffectCurrent
                        mrs家庭地址.Update
                    End If
                End If
            End If
        End If
    ElseIf KeyCode = vbKeyDown And cbo家庭地址.Text <> "" Then
        If SendMessage(cbo家庭地址.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, True, 0&)
    ElseIf KeyCode = vbKeyF3 Then
        cmd家庭地址.SetFocus
        Call cmd家庭地址_Click
    End If
End Sub

Private Sub cbo家庭地址_KeyUp(KeyCode As Integer, Shift As Integer)
    '此时text中已接收输入的信息
    '此事件处理删除和退格键,删除部分输入项目后,下拉列表数据中做对应的数据筛选
    '如果全部文字都删除了,则清空下拉列表数据
        
    Dim str家庭地址 As String, i As Long
    Dim lng位置 As Long
    
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If mrs家庭地址 Is Nothing Or mbln家庭地址输入 = False Then Exit Sub
        
        str家庭地址 = cbo家庭地址.Text                      '此时,如果选择了部分文字,则选择的文字已经被删除
        lng位置 = cbo家庭地址.SelStart
        
        If mrs家庭地址.State = 1 And Len(str家庭地址) > 1 Then
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(str家庭地址, 1))) > 0 Then
                mrs家庭地址.Filter = "简码 like '" & gstrLike & UCase(str家庭地址) & "*'"
            Else
                mrs家庭地址.Filter = "名称 Like '" & gstrLike & str家庭地址 & "*'"
            End If
            
            If Not mrs家庭地址.EOF Then
                
                If mrs家庭地址.RecordCount <> cbo家庭地址.ListCount Then
                    Call SendMessage(cbo家庭地址.Hwnd, CB_RESETCONTENT, 0, 0)
                    mrs家庭地址.Sort = "次数 Desc,名称"
                    For i = 1 To mrs家庭地址.RecordCount
                        AddComboItem cbo家庭地址.Hwnd, CB_ADDSTRING, 0, mrs家庭地址!名称
                        mrs家庭地址.MoveNext
                    Next
                    If SendMessage(cbo家庭地址.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, True, 0&)
                                        
                    cbo家庭地址.Text = str家庭地址
                    cbo家庭地址.SelStart = lng位置
                End If
            Else
                Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, False, 0&)
            End If
        ElseIf str家庭地址 = "" Then
            cbo家庭地址.Clear
            Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, False, 0&)
        End If
    End If
End Sub

Private Sub cbo家庭地址_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim str简码 As String
    Dim str家庭地址 As String
    Dim lng中间输入点 As Long
    
    If (mrs家庭地址 Is Nothing Or mbln家庭地址输入 = False) And KeyAscii <> 13 Then Exit Sub
    
    '用本地缓存匹配输入
    If KeyAscii <> 13 And KeyAscii <> vbKeyF4 And KeyAscii <> vbKeyEscape And _
        KeyAscii <> vbKeyBack And KeyAscii <> 26 And KeyAscii <> 3 And KeyAscii <> 22 Then   '26表示ctrl+z,3-ctrl+c,22-ctrl+v
            
        If mrs家庭地址.State = 0 Or cbo家庭地址.Text = "" Then  '输第一个字时不匹配
            Exit Sub
        End If
       
        '选中中间部分文本再输入的情况
        If cbo家庭地址.SelText <> "" And (cbo家庭地址.SelStart + cbo家庭地址.SelLength) <> Len(cbo家庭地址.Text) Then
            lng中间输入点 = cbo家庭地址.SelStart + 1
            cbo家庭地址.Text = Mid(cbo家庭地址.Text, 1, cbo家庭地址.SelStart) & Chr(KeyAscii) & Mid(cbo家庭地址.Text, cbo家庭地址.SelStart + cbo家庭地址.SelLength + 1)
            cbo家庭地址.SelText = ""
            str家庭地址 = cbo家庭地址.Text
        Else
            '输入点在尾部,或在中间时,后面的已选中
            If cbo家庭地址.SelStart = Len(cbo家庭地址.Text) Or (cbo家庭地址.SelStart + cbo家庭地址.SelLength) = Len(cbo家庭地址.Text) Then
                str家庭地址 = Mid(cbo家庭地址.Text, 1, cbo家庭地址.SelStart) & Chr(KeyAscii)
            Else
                str家庭地址 = Mid(cbo家庭地址.Text, 1, cbo家庭地址.SelStart) & Chr(KeyAscii) & Mid(cbo家庭地址.Text, cbo家庭地址.SelStart + 1)
                lng中间输入点 = cbo家庭地址.SelStart + 1
            End If
        End If
        
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(str家庭地址, 1))) > 0 Then
            mrs家庭地址.Filter = "简码 like '" & gstrLike & UCase(str家庭地址) & "*'"
        Else
            mrs家庭地址.Filter = "名称 Like '" & gstrLike & str家庭地址 & "*'"
        End If
        
        If Not mrs家庭地址.EOF Then
            If mrs家庭地址.RecordCount <> cbo家庭地址.ListCount Then
                Call SendMessage(cbo家庭地址.Hwnd, CB_RESETCONTENT, 0, 0)
                mrs家庭地址.Sort = "次数 Desc,名称"
                For i = 1 To mrs家庭地址.RecordCount
                    AddComboItem cbo家庭地址.Hwnd, CB_ADDSTRING, 0, mrs家庭地址!名称
                    mrs家庭地址.MoveNext
                Next
                If SendMessage(cbo家庭地址.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, True, 0&)
            End If
            
            i = KeyAscii    '用来后面判断是否是按退格删除键
            KeyAscii = 0
            cbo家庭地址.Text = str家庭地址
            cbo家庭地址.SelStart = Len(cbo家庭地址.Text)

            mrs家庭地址.MoveFirst   '如果不是输入的简码,相同则取下一个更多的
            If mrs家庭地址!名称 = str家庭地址 And i <> vbKeyBack Then
                mrs家庭地址.MoveNext
            End If
            If Not mrs家庭地址.EOF Then
                If InStr(1, mrs家庭地址!名称, str家庭地址) > 0 Or mrs家庭地址!简码 = UCase(str家庭地址) Then    '输入内容属于已有内容的一部分,则选中缓存多余文字
                    i = Len(cbo家庭地址.Text)
                    cbo家庭地址.Text = mrs家庭地址!名称
                    cbo家庭地址.SelStart = i
                    cbo家庭地址.SelLength = Len(cbo家庭地址.Text) - cbo家庭地址.SelStart
                    
                    If mrs家庭地址.RecordCount = 1 Then Exit Sub
                End If
            End If
            
        '没有找到匹配的缓存数据时,需清除下拉列表数据
        Else
            Call SendMessage(cbo家庭地址.Hwnd, CB_RESETCONTENT, 0, 0)
            If SendMessage(cbo家庭地址.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, False, 0&)
            KeyAscii = 0
            cbo家庭地址.Text = str家庭地址
            cbo家庭地址.SelStart = Len(cbo家庭地址.Text)
        End If
        
        If lng中间输入点 > 0 Then cbo家庭地址.SelStart = lng中间输入点: cbo家庭地址.SelText = ""
        
    ElseIf KeyAscii = 13 Then
        'a.在没有选中任何文字,且输入内容为空,光标为于末端时,确认输入,并保存信息到本地缓存
        Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, False, 0&)
        
        If cbo家庭地址.Text = "" Then
            If gbln家庭地址 And txtPatient.Text <> "" Then
                Exit Sub
            Else
                Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            End If
        End If
        
        '下拉列表弹出时按回车,则定位到末尾
        If cbo家庭地址.SelText = cbo家庭地址.Text Then cbo家庭地址.SelStart = Len(cbo家庭地址.Text): Exit Sub
        
        If mrs家庭地址.State = 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If zlCommFun.ActualLen(cbo家庭地址.Text) > 100 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
       
        'a.非下拉状态下按回车,没有选择文本
        If cbo家庭地址.SelText = "" Then
            str家庭地址 = cbo家庭地址.Text
            mrs家庭地址.Filter = "名称='" & str家庭地址 & "'"
            If mrs家庭地址.EOF Then
                str简码 = Mid(zlCommFun.zlGetSymbol(str家庭地址), 1, 10)
                If str简码 <> UCase(str家庭地址) Then
                    With mrs家庭地址
                        .AddNew
                        !类别 = "用户"
                        !名称 = str家庭地址
                        !简码 = str简码
                        !次数 = 1
                        .Update                 '在窗体Unload中save
                    End With
                End If
            Else
                mrs家庭地址!次数 = mrs家庭地址!次数 + 1
                mrs家庭地址.Update
                
                If zlCommFun.IsCharAlpha(str家庭地址) Then
                    If mrs家庭地址.RecordCount = 1 Then
                        cbo家庭地址.Text = mrs家庭地址!名称
                    Else
                        Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, True, 0&)
                        Exit Sub
                    End If
                End If
            End If
            
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo联系人关系_Change()
    mstrPlugChange = mstrPlugChange & ",联系人关系"
End Sub

Private Sub cbo联系人关系_Click()
    With cbo联系人关系
        If .ListIndex = 8 And txt其他关系.Visible = False Then
            .Width = 1200: txt其他关系.Visible = True
        ElseIf .ListIndex <> 8 And txt其他关系.Visible Then
            .Width = 2445: txt其他关系.Visible = False
        ElseIf .ListIndex = -1 Then
            .Width = 2445
        End If
    End With
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("关系") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("关系")) = zlCommFun.GetNeedName(cbo联系人关系.Text)
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("附加信息")) = zlCommFun.GetNeedName(txt其他关系.Text)
    End If
End Sub

Private Sub cbo民族_Change()
    mstrPlugChange = mstrPlugChange & ",民族"
End Sub

Private Sub cbo性别_Change()
    mstrPlugChange = mstrPlugChange & ",性别"
End Sub

Private Sub cbo职业_Change()
    mstrPlugChange = mstrPlugChange & ",职业"
End Sub

Private Sub chk复诊_Click()
    RaiseEvent ReturnVisitClick
End Sub

Private Sub cmdMedicalWarning_Click()
'问题号:56599
    Dim rsTemp As Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txtMedicalWarning.Hwnd)
    strSQL = "" & _
    "       Select 编码 as ID,名称,简码 From 医学警示 Where 名称 Not Like '其他%'"
    Set rsTemp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "医学警示", False, txtMedicalWarning.Text, "", False, False, False, vRect.Left, vRect.Top - 180, 500, blnCancel, False, True)
    If blnCancel Then Exit Sub
    If Not rsTemp Is Nothing Then
      While rsTemp.EOF = False
        strTemp = strTemp & "," & rsTemp!名称
        rsTemp.MoveNext
      Wend
    End If
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    If strTemp <> "" Then txtMedicalWarning.Text = strTemp
End Sub
''Private Sub SetPatiBaseModiPropertyEanbled()
''   '---------------------------------------------------------------------------------------------------------------------------------------------
''    '功能:设置病人的基本信息的编辑属性
''    '编制:刘兴洪
''    '日期:2013-11-04 11:59:46
''    '---------------------------------------------------------------------------------------------------------------------------------------------
''    Dim blnEnabled As Boolean
''    Dim lngColor As Long
''    On Error GoTo errHandle
''
''    blnEnabled = mbytFun <> 0 Or mlng病人ID = 0
''
''    txtPatient.Enabled = blnEnabled
''    cbo性别.Enabled = blnEnabled
''    txt年龄.Enabled = blnEnabled
''    cbo年龄单位.Enabled = blnEnabled
''    txt出生日期.Enabled = blnEnabled
''
''    lngColor = IIf(blnEnabled, vbWhite, Me.BackColor)
''    txtPatient.BackColor = lngColor
''    cbo性别.BackColor = lngColor
''    txt年龄.BackColor = lngColor
''    cbo年龄单位.BackColor = lngColor
''    txt出生日期.BackColor = lngColor
''
''    Exit Sub
''errHandle:
''    If ErrCenter() = 1 Then
''        Resume
''    End If
'' End Sub
''

Private Sub cmdPicClear_Click()
    '问题号:56599
    imgPatient.Picture = Nothing
    mlng图像操作 = 3
End Sub

Private Sub cmdPicCollect_Click()
    If mobjPubPatient Is Nothing Then Exit Sub
    If mobjPubPatient.PatiImageGatherer(Me, mstr采集图片) = False Then Exit Sub
    imgPatient.Picture = LoadPicture(mstr采集图片)
    mlng图像操作 = 2
End Sub

Private Sub cmdPicFile_Click()
    '问题号:56599
    Dim strFileDir As String
On Error GoTo errHanl:
    With cmdialog
        .CancelError = True
        .flags = cdlOFNHideReadOnly
        .Filter = "(*.bmp)|*.bmp"
        .FilterIndex = 2
        .ShowOpen
        strFileDir = .FileName
        If strFileDir = "" Then Exit Sub
        imgPatient.Picture = LoadPicture(strFileDir)
    End With
    mlng图像操作 = 1
    Exit Sub
errHanl:
     
End Sub

Private Sub cmd区域_Click()
    If zl_SelectAndNotAddItem(Me, txt区域, "", "区域", "区域选择", True, False) = False Then
        Exit Sub
    End If
End Sub

Private Sub lblICCard_Click()
    If txt卡号.Enabled = False Or txt卡号.Locked Then Exit Sub
    If gCurSendCard.bln就诊卡 And gCurSendCard.str卡名称 <> "二代身份证" Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txt卡号.Text = mobjICCard.Read_Card()
            If txt卡号.Text <> "" Then mfrmMain.mblnICCard = True
        End If
        Exit Sub
    End If
    '读取其卡信息
    '刘兴洪
    If zlLoadInfor = False Then
        If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
        Exit Sub
    End If
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    txt卡号.Text = strCardNo
    If txt卡号.Text <> "" Then
        '问题号:56599
        If strXmlCardInfor <> "" Then Call LoadPati(strXmlCardInfor)
        If txt密码.Enabled And txt密码.Visible Then txt密码.SetFocus
    Else
        If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind        As Long
    Dim lng病人ID           As Long
    Dim strPassWord         As String
    Dim strErrMsg           As String
    Dim strTmp              As String
    Dim bln允许签约         As Boolean
    Dim bln是否签约         As Boolean
    Dim blnUserCancel       As Boolean
    Static blnIsPatient     As Boolean '上一次触发事件的控件是否是,病人姓名文本框
    '在姓名处刷卡需填充病人信息
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
'获取病人标签
GetPatientTab:
        
        If mbln扫描身份证签约 Then mbln扫描身份证 = True
        Set mrsPatiInfo = Nothing
        '获取病人ID
        blnUserCancel = False
        If gobjSquare.objSquareCard.zlGetPatiID("身份证", strID, False, lng病人ID, strPassWord, strErrMsg, , , , InStr(gstrPrivs, ";合并病人信息;") > 0, , , blnUserCancel) = False Then lng病人ID = 0
             '操作员取消获取新病人
        bln允许签约 = True
        mblnNewPatient = True
        
        If blnUserCancel = True And lng病人ID = 0 Then Exit Sub
        
        If lng病人ID = 0 Then
            txt身份证号.Text = strID
            txtPatient.Text = strName
            Call zlControl.CboLocate(cbo性别, strSex)
            Call zlControl.CboLocate(cbo民族, strNation)
            txt出生日期.Text = Format(datBirthDay, "yyyy-MM-dd")
            txt出生时间.Text = "00:00"
            cbo家庭地址.Text = IIf(Trim(cbo家庭地址.Text) = "", strAddress, cbo家庭地址.Text)
            txtRegLocation.Text = strAddress
            '89242:李南春,2015/12/10,获取病人地址信息
            padd家庭地址.Value = IIf(Trim(padd家庭地址.Value) = "", strAddress, padd家庭地址.Value)
            padd户口地址.Value = strAddress
            
            '74430,冉俊明,2014-7-7,挂号中的病人信息编辑功能中提供采集照片功能
            Call LoadIDImage
        Else
            
            Set mrsPatiInfo = GetPatiByID("病人ID", CStr(lng病人ID))
            If mrsPatiInfo.EOF = False Then
                If (Nvl(mrsPatiInfo!姓名) <> Trim(strName) Or Nvl(mrsPatiInfo!性别) <> strSex Or Format(Nvl(mrsPatiInfo!出生日期, "00:00:00"), "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd")) Then
                    bln允许签约 = False
                    mbln扫描身份证 = False
                    txt支付密码.Text = ""
                    txt支付密码.Tag = ""
                    txt验证密码.Text = ""
                    txt验证密码.Tag = ""
                    If gCurSendCard.str卡名称 = "二代身份证" Then
                        MsgBox "身份证信息与HIS中病人信息不一致,不能进行签约操作！", vbInformation, gstrSysName
                        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                    Else
                        '非二代身份证绑定
                        zl加载病人信息 mrsPatiInfo
                        If txtRegLocation.Text = "" Then
                            txtRegLocation.Text = strAddress
                            padd户口地址.Value = strAddress
                        End If
                    End If
                    Set mrsPatiInfo = Nothing
                Else
                    bln允许签约 = True
                    mbln扫描身份证 = True
                    zl加载病人信息 mrsPatiInfo
                    If imgPatient.Picture = 0 Then
                        '74430,冉俊明,2014-7-7,挂号中的病人信息编辑功能中提供采集照片功能
                        Call LoadIDImage
                    End If
                    If txtRegLocation.Text = "" Then
                        txtRegLocation.Text = strAddress
                        padd户口地址.Value = strAddress
                    End If
                End If
                
                '是否扫描身份证签约
                '检查需要发放的身份证卡是否已经被签约
                If mbln扫描身份证签约 And bln允许签约 Then
                    bln是否签约 = 是否已经签约(strID)
                    If bln是否签约 Then
                        If gCurSendCard.str卡名称 = "二代身份证" Then
                            MsgBox "当前病人已经进行签约病人,无需进行再次签约！", vbInformation, gstrSysName
                            Set mrsPatiInfo = Nothing
                            txt支付密码.Text = ""
                            txt支付密码.Tag = ""
                            txt验证密码.Text = ""
                            txt验证密码.Tag = ""
                            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                        End If
                        bln允许签约 = False
                    End If
                    mbln扫描身份证 = Not bln是否签约
                End If
                
            End If
        End If
        If mbln扫描身份证 And bln允许签约 Then txt支付密码.Tag = Get医疗卡类别ID("二代身份证")
        '从卡上获取病人信息后,调用EMPI
        Call zlQueryEMPIPatiInfo
        blnIsPatient = True
        SetCtrVisibleAndMove
        If gblnNewCardNoPop Then Call cmdOK_Click
        Exit Sub
    End If
    
    If Not ActiveControl Is txt身份证号 Then Exit Sub
    
    '加载病人信息（如果存在）
    '问题号:54197
    
    If txtPatient.Text = "" Or mrsPatiInfo Is Nothing Then GoTo GetPatientTab: Exit Sub  '如果没有病人信息
   
    '如果已存在病人信息与新刷出来的病人身份证不一致,限制执行
    If Nvl(mrsPatiInfo!身份证号) <> strID Then '身份证号不一致 ,不能够继续
        Call MsgBox("当前病人身份证号与身份证上的信息不一致,不能继续!", vbInformation, Me.Caption)
        Exit Sub
    End If
    
    '如果已经存在了病人信息,
    
    '1.当前病人,没有填写身份证号的情况
    '  --检查姓名,性别,以及出生日期  如果不一致,提示是否继续
    '2.当前病人与身份证病人不一致的情况
    '
    '病人信息检查
    
    If Nvl(mrsPatiInfo!姓名) <> Trim(strName) Or Nvl(mrsPatiInfo!性别) <> strSex Or Format(txt出生日期.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
      
        If Nvl(mrsPatiInfo!姓名) <> Trim(strName) Then
             strErrMsg = strErrMsg & " 姓名:" & (mrsPatiInfo!姓名) & " 姓名(身份证):" & strName & vbCrLf
             strTmp = strTmp & "," & "姓名"
        End If
        If Nvl(mrsPatiInfo!性别) <> strSex Then
             strErrMsg = strErrMsg & " 性别:" & Nvl(mrsPatiInfo!性别) & " 性别(身份证):" & strSex & vbCrLf
             strTmp = strTmp & "," & "性别"
        End If
        If Format(txt出生日期.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
             strErrMsg = strErrMsg & " 出生日期:" & Format(txt出生日期.Text, "yyyy-MM-dd") & " 出生日期(身份证):" & Format(datBirthDay, "yyyy-MM-dd") & vbCrLf
             strTmp = strTmp & "," & "出生日期"
        End If
        strTmp = Mid(strTmp, 2)
        strErrMsg = "当前病人信息与身份证上的[" & strTmp & "]等信息不一致," & vbCrLf & strErrMsg
        strErrMsg = strErrMsg & "是否以身份证上的[" & strTmp & "]信息替换当前病人的相应信息?" & vbCrLf
        If MsgBox(strErrMsg, vbYesNo + vbDefaultButton2 + vbQuestion, Me.Caption) = vbYes Then
             txtPatient.Text = strName
             txt身份证号.Text = strID
             Call zlControl.CboLocate(cbo性别, strSex)
             txt出生日期.Text = Format(datBirthDay, "yyyy-MM-dd")
        End If
    End If
    
    cbo年龄单位.Tag = cbo年龄单位.Text
    
    '是否扫描身份证签约
    '检查需要发放的身份证卡是否已经被签约
    If mbln扫描身份证签约 Then mbln扫描身份证 = Not 是否已经签约(strID)
    If mbln扫描身份证 Then txt支付密码.Tag = Get医疗卡类别ID("二代身份证")
    SetCtrVisibleAndMove
    If gblnNewCardNoPop Then Call cmdOK_Click
End Sub

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
    If cbo付款方式.ListIndex = -1 And cbo付款方式.ListCount > 0 Then cbo付款方式.ListIndex = 0
End Sub

Private Sub cbo国籍_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo国籍.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo国籍.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo国籍.ListIndex = lngIdx
    If cbo国籍.ListIndex = -1 And cbo国籍.ListCount > 0 Then cbo国籍.ListIndex = 0
End Sub

Private Sub cbo婚姻_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo婚姻.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo婚姻.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo婚姻.ListIndex = lngIdx
    If cbo婚姻.ListIndex = -1 And cbo婚姻.ListCount > 0 Then cbo婚姻.ListIndex = 0
End Sub

Private Sub cbo民族_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo民族.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo民族.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo民族.ListIndex = lngIdx
    If cbo民族.ListIndex = -1 And cbo民族.ListCount > 0 Then cbo民族.ListIndex = 0
End Sub

Private Sub cbo年龄单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub


Private Sub cbo年龄单位_LostFocus()
    Dim strBirth As String
    If cbo年龄单位.Locked Then Exit Sub
    '更正出生日期
    '69026,冉俊明,2014-8-8,检查输入年龄
    If cbo年龄单位.Text <> cbo年龄单位.Tag Then
        mblnChange = False
        If mblnGetBirth Then
            If mobjPubPatient.ReCalcBirthDay(Trim(txt年龄.Text) & cbo年龄单位.Text, strBirth) Then
                txt出生日期.Text = Format(strBirth, "yyyy-mm-dd")
                txt出生时间.Text = Format(strBirth, "hh:mm")
            End If
        End If
        mblnChange = True
        cbo年龄单位.Tag = cbo年龄单位.Text
    End If
    
    If Trim(txt年龄.Text) <> "" Then
        If mobjPubPatient Is Nothing Then Exit Sub
        If mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text) & cbo年龄单位.Text) = False Then
            If txt年龄.Visible And txt年龄.Enabled And Not txt年龄.Locked Then
                txt年龄.SetFocus: Exit Sub
            End If
        End If
    End If
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
    If cbo职业.ListIndex = -1 And cbo职业.ListCount > 0 Then cbo职业.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
   '相关的输入检查
    mblnCancel = True
    mstrPlugChange = ""
    If mbytFun = 0 And mlng病人ID <> 0 Then
        '35233
'        If CheckValied = False Then Exit Sub
    
        Call CloseIDCard    '47007
        Me.Hide: Exit Sub
    Else
        '必须检查证件合法性
        If IsCertificateCard(mlng病人ID) = False Then Exit Sub
    End If

    Call CloseIDCard
    If Me.Visible Then Me.Hide
    Exit Sub
ErrOther:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function GetmblnCancel() As Boolean
    GetmblnCancel = mblnCancel
End Function

Private Function CheckValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入是否正确
    '编制:刘兴洪
    '日期:2011-01-07 18:13:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long, strSimilar As String, i As Integer, strMCAccount As String
    Dim strSQL As String, rsTmp As ADODB.Recordset, intQuery As Integer
    Dim blnPlugInCheck As Boolean, str出生时间 As String
    Dim strbirthday As String, strAge As String, strSex As String, strErrInfo As String, strInfo As String
    
    '82859:李南春,2015/4/8,病人基本信息调整
    If mblnCancel = False And mbytFun = 1 And mstrCard <> "" And gbln卡费仅划价 Or (mblnCancel = False And mbytFun = 0 And mlng病人ID <> 0) Then
        If mlng病人ID > 0 And mbln基本信息调整 And (mstr年龄 & mstr年龄单位 <> IIf(IsNumeric(txt年龄.Text), txt年龄.Text & cbo年龄单位.Text, txt年龄.Text) Or mstr性别 <> NeedName(cbo性别.Text) Or mstr姓名 <> txtPatient.Text Or _
            mstr身份证号 <> txt身份证号.Text Or mstr出生日期 <> txt出生日期.Text Or mstr出生时间 <> txt出生时间.Text) Then
            If MsgBox("病人基本信息已发生改变，是否继续？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                '记录病人原始信息
                txtPatient.Text = mstr姓名:  cbo性别.ListIndex = cbo.FindIndex(cbo性别, mstr性别, True)
                txt年龄.Text = mstr年龄
                If mstr年龄单位 <> "" Then cbo年龄单位.ListIndex = cbo.FindIndex(cbo年龄单位, mstr年龄单位, True): cbo年龄单位.Visible = True: txt年龄.Width = 690
                If mstr出生日期 <> "" Then txt出生日期.Text = mstr出生日期
                If mstr出生时间 <> "" Then txt出生时间.Text = mstr出生时间
                txt身份证号.Text = mstr身份证号
                Exit Function
            Else
                '记录病人新的信息
                mstr姓名 = txtPatient.Text: mstr性别 = NeedName(cbo性别.Text)
                mstr年龄 = txt年龄.Text: mstr年龄单位 = NeedName(cbo年龄单位.Text)
                mstr出生日期 = txt出生日期.Text: mstr出生时间 = txt出生时间.Text
                mstr身份证号 = txt身份证号.Text
            End If
        End If
    End If
    mblnCancel = False
    
    If txt门诊号.Text = "" And txt门诊号.Enabled Then
        MsgBox "请输入病人的门诊号！", vbInformation, gstrSysName
        If txt门诊号.Visible And txt门诊号.Enabled Then txt门诊号.SetFocus
        Exit Function
    End If
    
    If cbo费别.ListIndex = -1 Then
        MsgBox "请选择病人的费别！", vbInformation, gstrSysName
        If cbo费别.Visible Then cbo费别.SetFocus
        Exit Function
    End If
    If mbytFun = 1 And mstrCard = "" And Trim(txt卡号.Text) <> "" Then
        If CheckPatiValid(Trim(txt卡号.Text)) = False Then
            If txt卡号.Visible Then txt卡号.SetFocus
            Exit Function
        End If
    End If
    If txtPatient.Text = "" Then
        MsgBox "请输入病人的姓名！", vbInformation, gstrSysName
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Function
    End If
    
    If txtPatiMCNO(0).Text <> "" Or txtPatiMCNO(1).Text <> "" Then
        If txtPatiMCNO(0).Text <> txtPatiMCNO(1).Text Then
            MsgBox "请检查,两次输入的医保号不一致！", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
        If zlCommFun.ActualLen(txtPatiMCNO(0).Text) > txtPatiMCNO(0).MaxLength Then
            MsgBox "请检查,医保号最大长度不能超过" & txtPatiMCNO(0).MaxLength & "个字符！", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
    End If
    
    If txtMobile.Text <> "" And IsMobileNO(txtMobile.Text) = False Then
        MsgBox "输入的手机号格式不正确，请重新录入！", vbInformation, gstrSysName
        If txtMobile.Visible And txtMobile.Enabled Then txtMobile.SetFocus
        Exit Function
    End If
    
    If CheckTextLength("姓名", txtPatient) = False Then Exit Function
    If CheckTextLength("出生地点", txtBirthLocation) = False Then Exit Function
    If CheckTextLength("年龄", txt年龄) = False Then Exit Function
    '89242:李南春,2015/12/10,地址信息检查
    If mblnStructAdress Then
        If Not CheckStructAddr(padd家庭地址, padd家庭地址.MaxLength) Then Exit Function
        If Not CheckStructAddr(padd户口地址, padd户口地址.MaxLength) Then Exit Function
    Else
        If zlCommFun.ActualLen(cbo家庭地址.Text) > glngMax家庭地址 Then
            MsgBox "现住址输入过长，只允许输入" & glngMax家庭地址 & "个字符或" & glngMax家庭地址 \ 2 & "个汉字，请检查!", vbInformation, gstrSysName
            cbo家庭地址.SetFocus: Exit Function
        End If
        If CheckTextLength("户口地址", txtRegLocation) = False Then Exit Function
    End If
    If CheckTextLength("现住址邮编", txt家庭邮编) = False Then Exit Function
    If CheckTextLength("户口邮编", txt户口地址邮编) = False Then Exit Function
    If CheckTextLength("出生地点", txtBirthLocation) = False Then Exit Function
    '83062
    For i = 1 To msh过敏.Rows - 1
        If zlCommFun.ActualLen(msh过敏.TextMatrix(i, 1)) > 100 Then
            MsgBox "病人过敏药物反应输入过长，只允许输入100个字符或50个汉字，请检查！", vbInformation, gstrSysName
            If msh过敏.Enabled And msh过敏.Visible Then msh过敏.SetFocus
            Exit Function
        End If
        If zlCommFun.ActualLen(msh过敏.TextMatrix(i, 0)) > 60 Then
            MsgBox "病人过敏药物名称输入过长，只允许输入60个字符或30个汉字，请检查！", vbInformation, gstrSysName
            If msh过敏.Enabled And msh过敏.Visible Then msh过敏.SetFocus
            Exit Function
        End If
    Next i
    '69026,冉俊明,2014-8-11,年龄有效性检查
    '76703,冉俊明,2014-8-15
    
    If mbln禁止输入年龄 Then
        '禁止输入年龄的情况,检查是否录入出生日期
        If txt出生日期.Enabled And IsDate(txt出生日期.Text) = False And Not (gblnAutoAddName And txtPatient.Text = "新病人") Then
            MsgBox "必须输入病人出生日期！", vbInformation, gstrSysName
            txt出生日期.SetFocus: Exit Function
        End If
        If mobjPubPatient Is Nothing Then Exit Function
        If mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), _
                IIf(txt出生日期.Text = "____-__-__", "", txt出生日期.Text) & _
                IIf(txt出生时间.Text = "__:__", "", " " & txt出生时间.Text)) = False Then
            If txt出生日期.Enabled And txt出生日期.Visible Then txt出生日期.SetFocus
            Exit Function
        End If
    End If
    
    If txt年龄.Enabled And txt年龄.Visible Then
        If mobjPubPatient Is Nothing Then Exit Function
        If mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), _
                IIf(txt出生日期.Text = "____-__-__", "", txt出生日期.Text) & _
                IIf(txt出生时间.Text = "__:__", "", " " & txt出生时间.Text)) = False Then
            txt年龄.SetFocus:  Exit Function
        End If
    End If
    
    If IsDate(zlCommFun.GetIDCardDate(txt身份证号.Text)) Then
        If Format(zlCommFun.GetIDCardDate(txt身份证号.Text), "yyyy-mm-dd") <> Format(txt出生日期.Text, "yyyy-mm-dd") Then
            intQuery = MsgBox("输入的身份证号与输入的出生日期不一致，使用身份证号获取的出生日期吗？", vbQuestion + vbYesNoCancel, gstrSysName)
            If intQuery = 6 Then
                txt出生日期.Text = zlCommFun.GetIDCardDate(txt身份证号.Text)
            ElseIf intQuery = 2 Then
                CheckValied = False
                Exit Function
            End If
        End If
    End If
    
    If IsDate(txt出生日期.Text) Then
        '76669，李南春,2014-8-15,年龄与出生日期检查
        str出生时间 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        If CDate(str出生时间) > zlDatabase.Currentdate Then
            If MsgBox("出生时间：" & str出生时间 & " 超过了当前系统时间。" & _
                vbCrLf & vbCrLf & "请检查年龄或出生日期的正确性 ，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If txt出生日期.Enabled And txt出生日期.Visible Then txt出生日期.SetFocus
                Exit Function
            End If
        End If
        If mbln监护人录入 And Trim(txt监护人.Text) = "" Then
            '61945 监护人录入 检查
            strSQL = "Select Floor(Months_Between(Sysdate, [1]) / 12) as 年龄 From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(txt出生日期.Text))
            If Not rsTmp Is Nothing Then
                If Val(Nvl(rsTmp!年龄)) <= mlng监护人年龄 And mlng监护人年龄 <> 0 Then
                    MsgBox "病人在[" & mlng监护人年龄 & "岁]下必须录入监护人,请检查!"
                    Set rsTmp = Nothing
                    If txt监护人.Enabled And txt监护人.Visible Then txt监护人.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    
    strMCAccount = Trim(txtPatiMCNO(0).Text)
    If mlngOutModeMC = 920 And strMCAccount <> txtPatiMCNO(0).Tag And strMCAccount <> "" Then
        strMCAccount = UCase(strMCAccount)
        If CheckExistsMCNO(strMCAccount) Then
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
    End If
    
    '104238:李南春，2017/2/15，检查卡号是否满足发卡控制限制
    If txt卡号.Text <> "" And Len(txt卡号.Text) <> gCurSendCard.lng卡号长度 And Not gCurSendCard.bln严格控制 Then
        Select Case gCurSendCard.byt发卡控制
            Case 0
                MsgBox "输入的卡号小于" & gCurSendCard.str卡名称 & "设定的卡号长度，请重新输入！", vbExclamation, gstrSysName
                If txt卡号.Visible And txt卡号.Enabled Then txt卡号.SetFocus
                    Exit Function
            Case 2
                If MsgBox("输入的卡号小于" & gCurSendCard.str卡名称 & "设定的卡号长度，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txt卡号.Visible And txt卡号.Enabled Then txt卡号.SetFocus
                    Exit Function
                End If
        End Select
    End If
    
     '不直接收取卡费则产生为划价单,同时建立病案    '否则与挂号费用一起收取,在挂号保存时才建立病案
    '绑定卡模式调用时,此处不建档
    If txt密码.Text <> txt验证.Text And txt验证.Enabled And txt验证.Visible Then    '
        MsgBox "两次输入的密码不一致,请重新输入！", vbInformation, gstrSysName
        txt密码.Text = "": txt验证.Text = ""
        txt密码.SetFocus: Exit Function
    End If
    If txt卡号.Text <> "" And Trim(txt密码.Text) <> "" And txt密码.Visible Then
        Select Case gCurSendCard.int密码长度限制
        Case 0
        Case 1
            If Len(txt密码.Text) <> gCurSendCard.int密码长度 Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & gCurSendCard.int密码长度 & "位", vbOKOnly + vbInformation
                If txt密码.Enabled Then txt密码.SetFocus
                Exit Function
             End If
        Case Else
            If Len(txt密码.Text) < Abs(gCurSendCard.int密码长度限制) Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & Abs(gCurSendCard.int密码长度限制) & "位以上.", vbOKOnly + vbInformation
                If txt密码.Enabled Then txt密码.SetFocus
                Exit Function
             End If
        End Select
    End If
  
    '81103,冉俊明,2014-12-26,录入身份证号后,出生日期、年龄、性别的同步关联检查和调整
    If Trim(txt身份证号.Text) <> "" And Not mobjPubPatient Is Nothing Then
        'CheckPatiIdcard(ByVal strIdcard As String, Optional strBirthday As String, _
        '    Optional strAge As String, Optional strSex As String, Optional strErrInfo As String) As Boolean
        '功能：身份证号码合法性校验
        '入参：strIdCard 身份证号码
        '出参：strBirthday  函数返回True为出生日期
        '         strAge 函数返回True为年龄
        '         strSex 函数返回True为性别
        '         strErrInfo 函数返回False为错误信息
        '返回：True/False  身份证合法返回True(可从strBirthday，strSex获取出生日期和性别)，
        '       否则返回False(可从strErrInfo获取详细错误信息)
        If mobjPubPatient.CheckPatiIdcard(Trim(txt身份证号.Text), strbirthday, strAge, strSex, strErrInfo) Then
            If strSex <> NeedName(cbo性别.Text) Then strInfo = "性别"
            If strAge <> Trim(txt年龄.Text) & cbo年龄单位 Then strInfo = strInfo & IIf(strInfo = "", "年龄", "、年龄")
            If Format(strbirthday, "yyyy-mm-dd") <> txt出生日期.Text Then strInfo = strInfo & IIf(strInfo = "", "出生日期", "、出生日期")
            
            If strInfo <> "" Then
                If MsgBox("输入的" & strInfo & "与身份证号的" & strInfo & "不一致，" & _
                        "将根据身份证号修改" & strInfo & "，是否继续？", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                    Call zlControl.CboLocate(cbo性别, strSex)
                    txt出生日期.Text = Format(strbirthday, "yyyy-mm-dd")
                    mstr性别 = NeedName(cbo性别.Text)
                    mstr年龄 = txt年龄.Text: mstr年龄单位 = NeedName(cbo年龄单位.Text)
                    mstr出生日期 = txt出生日期.Text: mstr出生时间 = txt出生时间.Text
                    mstr身份证号 = txt身份证号.Text
                Else
                    Exit Function
                End If
            End If
        Else
            MsgBox strErrInfo, vbInformation, gstrSysName
            If txt身份证号.Enabled And txt身份证号.Visible Then txt身份证号.SetFocus
            Exit Function
        End If
    End If
        
     '检查相似病人信息(新增之前检查,以免加入了重复信息！！！)
    If Trim(txt身份证号.Text) <> "" And cmdOK.Caption Like "确定*" And mlng病人ID = 0 Then
        strSimilar = SimilarIDs(Trim(txt身份证号.Text))
        If strSimilar <> "" Then
            i = UBound(Split(strSimilar, "|")) + 1
            strSimilar = Replace(strSimilar, "|", vbCrLf)
            If i > 20 Then strSimilar = Mid(strSimilar, 1, 200) & "..."
            
            If MsgBox("在已有的病人信息中发现 " & i & " 个信息相似的病人(身份证号相同): " & vbCrLf & vbCrLf & _
                strSimilar & vbCrLf & vbCrLf & "确实要登记为新病人吗？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    '问题号:53408
    If IIf(zlDatabase.GetPara("扫描身份证签约", glngSys, mlngModul) = "1", 1, 0) = 0 And ((gCurSendCard.str卡名称 = "二代身份证" And Trim(txt卡号.Text) <> "") Or Trim(txt支付密码.Text) <> "") Then
         MsgBox "您没有权限进行签约操作,请到参数设置中设置【扫描身份证签约】！", vbOKOnly + vbInformation, gstrSysName
         txt卡号.Text = ""
         txt密码.Text = ""
         txt验证.Text = ""
         If txt卡号.Visible = True And txt卡号.Enabled = True Then txt卡号.SetFocus
         Exit Function
    End If
    
    If Trim(txt支付密码.Text) <> "" And Trim(txt身份证号.Text) <> "" Then
           If 是否已经签约(txt身份证号.Text) Then
                 MsgBox "身份证号码为:" & txt身份证号.Text & "已经签约不能重复签约！", vbOKOnly + vbInformation, gstrSysName
                 txt支付密码.Text = ""
                 If txt支付密码.Visible = True And txt支付密码.Enabled = True Then txt支付密码.SetFocus
                 Exit Function
           End If
    End If
    
    If mbln扫描身份证 = False And gCurSendCard.str卡名称 = "二代身份证" And txt卡号.Text <> "" Then
            MsgBox "绑定身份证只能以刷卡的方式进行，不允许手动输入身份证进行绑定!", vbOKOnly + vbInformation, gstrSysName
            txt卡号.Text = ""
            txt密码.Text = ""
            txt验证.Text = ""
            txt支付密码.Text = ""
            txt验证密码.Text = ""
            '74894:李南春,2014-07-08,取消绑定二代身份证的卡号信息
            mstrCard = ""
            If txt卡号.Visible = True And txt卡号.Enabled = True Then txt卡号.SetFocus
            Exit Function
    End If
    
    If mbln扫描身份证 = False And gCurSendCard.str卡名称 <> "二代身份证" And txt支付密码.Text <> "" Then
            MsgBox "绑定身份证只能以刷卡的方式进行，不允许手动输入身份证进行绑定!", vbOKOnly + vbInformation, gstrSysName
            txt身份证号.Text = ""
            txt支付密码.Text = ""
            txt验证密码.Text = ""
            If txt身份证号.Visible = True And txt身份证号.Enabled = True Then txt身份证号.SetFocus
        Exit Function
    End If
    
    If Trim(txt支付密码.Text) <> Trim(txt验证密码.Text) And (Trim(txt支付密码.Text) <> "" Or Trim(txt验证密码.Text) <> "") Then
        MsgBox "两次输入的密码不一致,请重新输入", vbOKOnly + vbInformation, gstrSysName
        txt支付密码.Text = "": txt验证密码.Text = ""
        If txt支付密码.Visible = True And txt支付密码.Enabled = True Then txt支付密码.SetFocus
        Exit Function
    End If
    
    '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    If CreatePlugInOK(mlngModul) And mlngPlugInHwnd <> 0 Then  '保存插件附加信息前的数据有效性检查
        On Error Resume Next
        blnPlugInCheck = gobjPlugIn.PatiInfoSaveBefore(mlng病人ID)
        Call zlPlugInErrH(Err, "PatiInfoSaveBefore")
        If Err = 0 And blnPlugInCheck = False Then
            Exit Function '检查未通过终止保存
        End If
        Err.Clear
    End If '
        
    '84672:李南春，长度检查以及联系人检查
    If CheckTextLength("其他关系", txt其他关系) = False Then Exit Function
    If txt联系人姓名.Text = "" And (txt联系人电话.Text <> "" Or txt联系人身份证.Text <> "" Or cbo联系人关系.Text <> "") Then
        If MsgBox("没有输入联系人姓名，联系人信息不会保存，是否继续？", vbYesNo + vbInformation, gstrSysName) = vbNo Then
            Exit Function
        Else
            txt联系人身份证.Text = "": txt联系人电话.Text = ""
            cbo联系人关系.ListIndex = -1: txt其他关系.Text = "": txt其他关系.Visible = False
        End If
    End If
    With vsLinkMan
        If .Rows >= 3 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) = "" And (.TextMatrix(i, 1) <> "" Or .TextMatrix(i, 2) <> "" Or .TextMatrix(i, 3) <> "") Then
                    If MsgBox("联系人列表第" & i & "行没有输入联系人姓名，此行的联系人信息不会保存，是否继续？", vbYesNo + vbInformation, gstrSysName) = vbNo Then
                        Exit Function
                    Else
                        .TextMatrix(i, 1) = "": .TextMatrix(i, 2) = "": .TextMatrix(i, 3) = "": .TextMatrix(i, 4) = ""
                    End If
                End If
            Next
        End If
    End With
    '90875:李南春,2016/11/8,医疗卡证件类型
    If IsCertificateCard(mlng病人ID) = False Then Exit Function
    CheckValied = True
End Function

Private Sub cmdOK_Click()
    Dim strPati As String, strCard As String, strMCAccount As String, strTmp As String, str其他关系 As String
    Dim rsCard As ADODB.Recordset, blnTrans As Boolean, strErrMsg As String, blnNewPati As Boolean
    Dim lng病人ID As Long, strNO As String
    Dim lngDept As Long, strDate As String, str出生日期 As String
    Dim str门诊号 As String, byt类型 As Byte, i As Integer
    Dim Datsys As Date, str就诊卡 As String, blnBound As Boolean
    Dim str家庭地址 As String, str户口地址 As String
    Dim strYLKNo As String, colPro As Collection, blnCard As Boolean
    
    txtPatient.Text = Trim(txtPatient.Text)
    txt年龄.Text = Trim(txt年龄.Text)
    
    Set mobjProPati = New Collection
    '相关的输入检查
    Set colPro = New Collection
    If CheckValied = False Then Exit Sub
    '问题号:51072
    If Len(Trim(txt密码.Text)) <= 0 And Len(Trim(txt卡号.Text)) > 0 Then    '没有输入密码
        If zl_Get设置默认发卡密码 = False Then Exit Sub
    End If

    strMCAccount = Trim(txtPatiMCNO(0).Text)
    If mlngOutModeMC = 920 And strMCAccount <> txtPatiMCNO(0).Tag And strMCAccount <> "" Then
        strMCAccount = UCase(strMCAccount)
    End If
    If txt出生时间 = "__:__" Then
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & " " & txt出生时间.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
   
    If Len(txt门诊号.Text) > mintNOLength + 1 And mintNOLength > 0 And mblnCheckNOValidity Then
        MsgBox "注意,输入的门诊号过大,请确认是否输入正常!", vbInformation, gstrSysName
        txt门诊号.SetFocus
        txt门诊号.SelStart = 0: txt门诊号.SelLength = Len(txt门诊号.Text)
        Exit Sub
    End If
    '问题号:57326
    If mlng病人ID <> 0 And mbytFun <> 0 And txt卡号.Text <> "" Then
        If Check发卡性质(mlng病人ID, gCurSendCard.lng卡类别ID) = False Then
            txt卡号.Text = ""
            txt密码.Text = ""
            txt验证.Text = ""
        End If
    End If

    If mbytFun = 1 And (mstrCard <> "" Or txtPatient.Text <> "") And gbln卡费仅划价 Or (mbytFun = 0 And mlng病人ID <> 0) Then
        Datsys = zlDatabase.Currentdate
        strDate = "To_Date('" & Format(Datsys, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        If Exist门诊号(txt门诊号.Text, IIf(mlng病人ID <> 0, mlng病人ID, 0)) Then
            str门诊号 = zlDatabase.GetNextNo(3)
            If Len(str门诊号) > txt门诊号.MaxLength Then
                MsgBox "当前门诊号已经被其它病人使用,系统自动更换门诊号为:" & str门诊号 & _
                       vbCrLf & "但超过了允许的最大门诊号长度:" & txt门诊号.MaxLength & "位,请输入一个门诊号!", vbInformation, gstrSysName
                If txt门诊号.Enabled Then txt门诊号.SetFocus
                Exit Sub
            End If
            txt门诊号.Text = str门诊号
        End If
        If mstrCard = "" Then mstrCard = txt卡号.Text
        If mlng病人ID <> 0 Then
            lng病人ID = mlng病人ID
            byt类型 = 3
            If mbytFun = 1 And mstrCard <> "" Then blnCard = True
        Else
            lng病人ID = zlDatabase.GetNextNo(1)
            byt类型 = 1: blnNewPati = True
        End If
        mlng病人ID = lng病人ID
        '问题:区域:38663
        Dim strCardNo As String, strPassWord As String
        If gCurSendCard.bln就诊卡 Then
            strCardNo = Trim(txt卡号.Text)
            strPassWord = zlCommFun.zlStringEncode(txt密码.Text)
        End If
        '问题号:51071
        '73609:李南春，2014-8-1，病人信息保存
        '84313,李南春,2015/4/27,联系人关系以及其他关系
        strPati = _
        "zl_挂号病人病案_INSERT(" & byt类型 & "," & lng病人ID & "," & txt门诊号.Text & "," & _
                  "'" & strCardNo & "','" & strPassWord & "'," & _
                  "'" & txtPatient.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & cbo年龄单位.Text & "'," & _
                  "'" & NeedName(cbo费别.Text) & "','" & NeedName(cbo付款方式.Text) & "'," & _
                  "'" & NeedName(cbo国籍.Text) & "','" & NeedName(cbo民族.Text) & "','" & NeedName(cbo婚姻.Text) & "'," & _
                  "'" & NeedName(cbo职业.Text, True) & "','" & txt身份证号.Text & "','" & txt单位名称.Text & "'," & _
                  Val(txt单位名称.Tag) & ",'" & txt单位电话.Text & "','" & txt单位邮编.Text & "'," & _
                  "'" & IIf(mblnStructAdress, Trim(padd家庭地址.Value), cbo家庭地址.Text) & "'," & _
                  "'" & txt家庭电话.Text & "','" & txt家庭邮编.Text & "'," & strDate & ",''," & str出生日期 & ",'" & strMCAccount & "','" & IIf(mfrmMain.mblnICCard, txt卡号.Text, "") & "'," & _
                  "NULL," & IIf(Trim(txt区域.Text) = "", "NULL,", "'" & Trim(txt区域.Text) & "',") & _
                   "'" & IIf(mblnStructAdress, Trim(padd户口地址.Value), Trim(txtRegLocation.Text)) & "'," & _
                   "'" & Trim(txt户口地址邮编.Text) & "'," & IIf(Trim(txt联系人身份证.Text) = "", "NULL,", "'" & Trim(txt联系人身份证.Text) & "',") & _
                  IIf(Trim(txt联系人姓名.Text) = "", "NULL,", "'" & Trim(txt联系人姓名.Text) & "',") & _
                  IIf(Trim(txt联系人电话.Text) = "", "NULL,", "'" & Trim(txt联系人电话.Text) & "',") & _
                  IIf(NeedName(cbo联系人关系.Text) = "", "NULL,", "'" & NeedName(cbo联系人关系.Text) & "',")    '问题号:40005
        '监护人_In         In 病人信息.监护人%Type := Null
        strPati = strPati & IIf(Trim(txt监护人.Text) = "", "NULL,", "'" & Trim(txt监护人.Text) & "',")  'lgf
        '54601:刘尔旋,2013-11-27,新增出生地点和户口地址
        strPati = strPati & IIf(Trim(txtBirthLocation.Text) = "", "NULL,", "'" & Trim(txtBirthLocation.Text) & "',")
        '手机号_In         In 病人信息.手机号%Type := Null
        strPati = strPati & "'" & txtMobile.Text & "')"
        
        '89242:李南春,2015/12/10,更新病人地址信息
        If mblnStructAdress Then
            If padd家庭地址.Value <> "" Then
               str家庭地址 = "zl_病人地址信息_update(1," & lng病人ID & ",NULL,3,'" & padd家庭地址.value省 & "','" & _
                   padd家庭地址.value市 & "','" & padd家庭地址.value区县 & "','" & padd家庭地址.value乡镇 & "','" & _
                   padd家庭地址.value详细地址 & "','" & padd家庭地址.Code & "')"
            Else
               str家庭地址 = "zl_病人地址信息_update(2," & lng病人ID & ",NULL,3)"
            End If
            
            If padd户口地址.Value <> "" Then
               str户口地址 = "zl_病人地址信息_update(1," & lng病人ID & ",NULL,4,'" & padd户口地址.value省 & "','" & _
                   padd户口地址.value市 & "','" & padd户口地址.value区县 & "','" & padd户口地址.value乡镇 & "','" & _
                   padd户口地址.value详细地址 & "','" & padd户口地址.Code & "')"
            Else
               str户口地址 = "zl_病人地址信息_update(2," & lng病人ID & ",NULL,4)"
            End If
        End If
        
        'str其他关系
        If cbo联系人关系.Text <> "" And txt其他关系.Visible Then
            str其他关系 = "Zl_病人信息从表_Update("
            '病人ID_In 病人信息从表.病人Id%Type
            str其他关系 = str其他关系 & "" & lng病人ID & ","
            '信息名_In 病人信息从表.信息名%Type0
            str其他关系 = str其他关系 & "'联系人附加信息',"
            '信息值_In 病人信息从表.信息值%Type
            str其他关系 = str其他关系 & "'" & txt其他关系.Text & "',"
            '就诊Id_In 病人信息从表.就诊Id%Type
            str其他关系 = str其他关系 & "'')"
        End If
        '90875:李南春,2016/11/8,医疗卡证件类型
        Call AddCertificate(lng病人ID, colPro, Datsys)
        mstrFirstCode = ""
    
        If bln发卡(True) Then
            mblnInRange = True
        Else
            mblnInRange = False
        End If
        If byt类型 <> 3 Or blnCard Then
            '不为发卡,就是绑定卡
            If bln发卡(True) = False Then
                blnBound = True
            Else
                blnBound = False
            End If
            Call ReLoadCardFee
            Set rsCard = zlGetSpecialItemFee(gCurSendCard.str特准项目, mstrPriceGrade, gCurSendCard.lng收费细目ID)
            If rsCard Is Nothing Then
                MsgBox "不能正确提取" & gCurSendCard.str卡名称 & "费用信息！", vbInformation, gstrSysName
                Exit Sub
            End If
            '变价并且最低限价为零时,不收卡费
            '98364:李南春,2016/7/7,不管卡费是否为0，发卡都应该正常产生费用记录。
            If Me.txt卡号.Text <> "" And blnBound = False Then
                '就诊卡费用:不是变价,屏蔽费别
                '类别保持不变,收费处以"收费特定项目"作为条件搜索。
                Select Case rsCard!科室标志
                Case 0  '无明确执行科室
                    lngDept = UserInfo.部门ID
                Case 1  '病人所在科室
                    lngDept = UserInfo.部门ID
                Case 2  '病人所在病区
                    lngDept = UserInfo.部门ID
                Case 3  '开单人所在科室
                    lngDept = UserInfo.部门ID
                Case 4  '指定科室
                    lngDept = GetOneDept(rsCard!收费细目ID)
                Case Else
                    lngDept = UserInfo.部门ID
                End Select

                strNO = zlDatabase.GetNextNo(13)
                strYLKNo = zlDatabase.GetNextNo(16)  '医疗卡
                strCard = "zl_门诊划价记录_Insert('" & strNO & "',1," & lng病人ID & ",NULL," & txt门诊号.Text & "," & _
                          "NULL,'" & txtPatient.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & cbo年龄单位.Text & "'," & _
                          "'" & NeedName(cbo费别.Text) & "',0," & UserInfo.部门ID & "," & _
                          UserInfo.部门ID & ",'" & UserInfo.姓名 & "',NULL," & rsCard!收费细目ID & "," & _
                          "'" & rsCard!收费类别 & "','" & rsCard!计算单位 & "',NULL,1,1,0," & lngDept & ",NULL," & _
                          rsCard!收入项目ID & ",'" & rsCard!收据费目 & "'," & Format(rsCard!现价, "0.000") & "," & _
                          Format(rsCard!现价, "0.00") & "," & Format(rsCard!现价, "0.00") & "," & strDate & "," & _
                          strDate & ",NULL,'" & UserInfo.姓名 & "','" & strYLKNo & "')"
                
                '存在卡费需要生成住院费用记录
                str就诊卡 = zlGetSaveCardFeeSQL(gCurSendCard.lng卡类别ID, 0, strYLKNo, lng病人ID, 0, UserInfo.部门ID, UserInfo.部门ID, 0, _
                zlStr.NeedName(cbo费别.Text), mstrCard, Trim(txtPatient.Text), zlStr.NeedName(cbo性别.Text), txt年龄.Text & cbo年龄单位.Text, _
                txt卡号.Text, zlCommFun.zlStringEncode(txt密码.Text), "挂号发卡", 0, 0, "", Datsys, mlng磁卡领用ID, rsCard, _
                IIf(mfrmMain.mblnICCard = True, txt卡号.Text, ""), , , , , strNO)
            ElseIf Me.txt卡号.Text <> "" Then
                str就诊卡 = GetCardDataSql(11, lng病人ID, gCurSendCard.lng卡类别ID, mstrCard, Me.txt卡号.Text, Me.txt密码.Text, _
                                    Datsys, "", "挂号绑定卡")
            End If
        End If
        If strPati <> "" Then zlAddArray mobjProPati, strPati
        If str家庭地址 <> "" Then zlAddArray mobjProPati, str家庭地址
        If str户口地址 <> "" Then zlAddArray mobjProPati, str户口地址
        If strCard <> "" Then zlAddArray mobjProPati, strCard
        If str就诊卡 <> "" Then zlAddArray mobjProPati, str就诊卡
    
    End If
    If Not mblnInRange And (byt类型 <> 3 Or blnCard) Or (str就诊卡 <> "") Then
        '发卡状态下，直接更新信息。
        Call SaveAfterArrList
    End If

    Call CloseIDCard
    If Me.Visible Then Me.Hide
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
ErrOther:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function SaveAfterArrList() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存病人信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-06 16:07:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTrans As Boolean
    Dim blnNewPati As Boolean
    Dim strErrMsg As String
    Dim i As Long
    
    On Error GoTo errH
    '无需保存数据
    If mobjProPati Is Nothing Then SaveAfterArrList = True: Exit Function
    '120737,焦博,2018-2-1,新建档病人，在挂号病人信息(就诊卡绑定卡)界面绑卡后，保存报错
    If mobjProPati.Count = 0 Then SaveAfterArrList = True: Exit Function
    blnNewPati = mlng病人ID = 0
    
    blnTrans = True
    zlExecuteProcedureArrAy mobjProPati, Me.Caption, True
    
    '101170:李南春,2017/5/3,保存HIS数据要提交EMPI数据，失败后所有数据都要回退
    If zlSaveEMPIPatiInfo(blnNewPati, mlng病人ID, 0, strErrMsg) = False Then
        gcnOracle.RollbackTrans
        If strErrMsg = "" Then strErrMsg = "向EMPI平台上传病人信息失败！"
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    Set mobjProPati = Nothing
    mstrPlugChange = ""
    If txt卡号.Text <> "" Then
        If gCurSendCard.bln是否写卡 Then Call WriteCard(mlng病人ID)
        Call mfrmMain.SetCardDisplay(gCurSendCard.str短名称 & ":" & Me.txt卡号.Text & "(" & IIf(Not (bln发卡(True)), "绑定卡", "发卡") & ")")
    End If
    '74430,冉俊明,2014-7-7,挂号中的病人信息编辑功能中提供采集照片功能
    Call SavePatiPic(mlng病人ID)
    '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    If CreatePlugInOK(mlngModul) And mlngPlugInHwnd <> 0 Then  '保存插件附加信息
        On Error Resume Next
        Call gobjPlugIn.PatiInfoSaveAfter(mlng病人ID)
        Call zlPlugInErrH(Err, "PatiInfoSaveAfter")
        Err.Clear: On Error GoTo 0
    End If
    SaveAfterArrList = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Function
End Function

Private Sub txtMobile_GotFocus()
    Call zlControl.TxtSelAll(txtMobile)
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txtMobile, KeyAscii, m数字式)
End Sub

Private Function WriteCard(lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写卡
    '入参:lng病人ID - 病人ID
    '编制:王吉
    '问题:56599
    '日期:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    On Error GoTo ErrHandl:
    
    WriteCard = gobjSquare.objSquareCard.zlBandCardArfter(Me, mlngModul, gCurSendCard.lng卡类别ID, lng病人ID, strExpend)
    Exit Function
ErrHandl:
    WriteCard = False
    If ErrCenter() = 1 Then Resume
End Function

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmd单位名称_Click()
    Call SearchUnit("", txt单位名称)
End Sub

Private Sub cmd过敏_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    Dim i As Integer
    
    strSQL = _
        " Select -1 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'西成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试 From Dual Union ALL" & _
        " Select -2 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试 From Dual Union ALL" & _
        " Select -3 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中草药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试 From Dual Union ALL" & _
        " Select ID,nvl(上级ID,-类型) as 上级ID,0 as 末级,NULL as 编码,名称," & _
        " NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试" & _
        " From 诊疗分类目录 Where 类型 IN (1,2,3) And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
        " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        " Union All" & _
        " Select Distinct A.ID,A.分类ID as 上级ID,1 as 末级,A.编码," & _
        " A.名称,A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类," & _
        " Decode(B.是否新药,1,'√','') as 新药,Decode(B.是否皮试,1,'√','') as 皮试" & _
        " From 诊疗项目目录 A,药品特性 B" & _
        " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"

    Set rsTmp = frmPubSel.ShowSelect(Me, strSQL, 2, "过敏药物", , msh过敏.TextMatrix(msh过敏.Row, 0), "请从下面的药品中选择一项作为病人过敏药物。")
    If Not rsTmp Is Nothing Then
        For i = 1 To msh过敏.Rows - 1
            If i <> msh过敏.Row Then
                If msh过敏.RowData(i) = rsTmp!ID Then
                    MsgBox "第 " & i & " 行的药物已经与你选择的药物种类相同,请重新选择！", vbInformation, gstrSysName
                    msh过敏.SetFocus
                    msh过敏_EnterCell
                    Exit Sub
                End If
            End If
        Next
        msh过敏.RowData(msh过敏.Row) = rsTmp!ID
        msh过敏.TextMatrix(msh过敏.Row, 0) = Trim(rsTmp!名称)
    End If
    msh过敏.SetFocus
    msh过敏_EnterCell
    
End Sub

Private Sub cmd家庭地址_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select Distinct Substr(名称,1,2) as ID,NULL as 上级ID,0 as 末级,NULL as 编码," & _
            " Substr(名称,1,2) as 名称 From 地区" & _
            " Union All" & _
            " Select 编码 as ID,Substr(名称,1,2) as 上级ID,1 as 末级,编码,名称 " & _
            " From 地区 Order by 编码", 2, "地区", , cbo家庭地址.Text)
    If Not rsTmp Is Nothing Then
        cbo家庭地址.Text = rsTmp!名称
        cbo家庭地址.SelStart = Len(cbo家庭地址.Text)
    End If
    cbo家庭地址.SetFocus
End Sub

Private Sub Form_Activate()
    If mbytFun = 1 Then
        picCard.Visible = True
        tbcPage.Top = picCard.Top + picCard.Height
        txt卡号.Locked = False
        txt卡号.PasswordChar = IIf(gCurSendCard.str卡号密文 <> "", "*", "")
        txt卡号.MaxLength = gCurSendCard.lng卡号长度
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txt卡号.IMEMode = 0
    Else
        If txt卡号.Text <> "" Then
            picCard.Visible = True
        Else
            picCard.Visible = False
        End If
        txt卡号.Locked = True
    End If
    mblnCancel = False
    tbcPage.Left = 0
    tbcPage.Width = Me.ScaleWidth
    
    Select Case mbytFun
        Case 0
            Me.Caption = "挂号病人信息编辑"
        Case 1
            Me.Caption = "挂号发卡"
    End Select
    
    If (mbytFun = 0 And mlng病人ID = 0) Or (mbytFun = 1 And mstrCard = "") Then '绑定就诊卡模式不提供取消按钮,以防Unload窗体,因为之前提取病人身份时加载的信息会被清除
        cmdOK.Caption = "确定(&O)"
        cmdCancel.Visible = True
        cmdCancel.Left = tbcPage.Left + tbcPage.Width - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    Else
        cmdOK.Caption = "确定(&O)"
        If (mbytFun = 0 And mlng病人ID <> 0) Then
            cmdCancel.Caption = "返回(&X)"
        End If
        cmdCancel.Visible = True
        cmdCancel.Left = tbcPage.Left + tbcPage.Width - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    End If
    
    '78408:李南春,2014/10/9,光标跳转
    If Not Me.ActiveControl Is msh过敏 Then
        If mbytFun = 1 And Not gCurSendCard.str卡名称 Like "二代身份证" Then
            If mstrCard = "" Then
                If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
            Else
                If txt密码.Enabled And txt密码.Visible Then txt密码.SetFocus
            End If
        Else
            If txtPatient.Visible = True And txtPatient.Enabled Then
                txtPatient.SetFocus
            ElseIf txt门诊号.Enabled And txt门诊号.Visible Then
                txt门诊号.SetFocus
            ElseIf txt出生日期.Enabled And txt出生日期.Visible Then
                txt出生日期.SetFocus
            End If
        End If
    End If
    
    mbln扫描身份证 = False
    mbln扫描身份证签约 = IIf(zlDatabase.GetPara("扫描身份证签约", glngSys, mlngModul) = "1", 1, 0) = "1"
    mbln禁止输入年龄 = Val(zlDatabase.GetPara("禁止输入年龄", glngSys, mlngModul, 0)) = 1
    If mbln禁止输入年龄 Then txt年龄.Enabled = False: cbo年龄单位.Enabled = False
    SetCtrVisibleAndMove
    '问题号:56599
    Me.Caption = "挂号病人信息【" & gCurSendCard.str卡名称 & IIf(bln发卡, "发卡", "绑定卡") & "】"
    If Not mfrmMain Is Nothing Then
        If mfrmMain.SendCard Then Me.Caption = "挂号病人信息【" & gCurSendCard.str卡名称 & "发卡" & "】"
    End If
    gsngStartTime = Timer
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If txt过敏.Visible Then
            txt过敏.Visible = False
            msh过敏_EnterCell
            msh过敏.SetFocus
        ElseIf lvwItems.Visible Then
            lvwItems.Visible = False
            txt过敏.Visible = True
            txt过敏.SetFocus
        ElseIf Not cmdCancel.Visible Then
            cmdOK_Click
        Else
            cmdCancel_Click
        End If
    ElseIf KeyCode = vbKeyF2 Then
        Call cmdOK_Click
    ElseIf KeyCode = vbKeyF4 And Shift = vbCtrlMask Then
        If txt卡号.Enabled And txt卡号.Visible Then
            Call lblICCard_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        '89242:李南春,2015/12/10,PatiAddress控件内部处理了跳转，外部不再处理
        If UCase(TypeName(Me.ActiveControl)) = UCase("PatiAddress") Then Exit Sub
        If InStr(1, "txtPatient,txt密码,lvwItems,txt年龄,cbo年龄单位,txt出生日期,msh过敏,txt过敏,txtPatiMCNO,txt区域,vsInoculate,vsCertificate,cbo家庭地址", Me.ActiveControl.Name) <= 0 Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub


Private Sub InitData()
'功能：初始化必要数据
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, lngTmp As Long
    Dim lngCardType As Long
        
    mDateSys = zlDatabase.Currentdate
    
    
    SetCtrVisibleAndMove

    If mrsBaseDict Is Nothing Then
        Set mrsBaseDict = GetBaseDict
    End If
    Set rsTmp = mrsBaseDict
    If rsTmp Is Nothing Then Exit Sub
    
    '国籍
    rsTmp.Filter = "类别='国籍'"
    cbo国籍.Clear
    For i = 1 To rsTmp.RecordCount
        cbo国籍.AddItem rsTmp!编码 & "-" & rsTmp!名称
        If rsTmp!缺省 = 1 Then
            cbo国籍.ItemData(cbo国籍.NewIndex) = 1
            cbo国籍.ListIndex = cbo国籍.NewIndex
        End If
        rsTmp.MoveNext
    Next

    '民族
    rsTmp.Filter = "类别='民族'"
    cbo民族.Clear
    For i = 1 To rsTmp.RecordCount
        cbo民族.AddItem rsTmp!编码 & "-" & rsTmp!名称
        If rsTmp!缺省 = 1 Then
            cbo民族.ItemData(cbo民族.NewIndex) = 1
            cbo民族.ListIndex = cbo民族.NewIndex
        End If
        rsTmp.MoveNext
    Next

    '婚姻状况
    rsTmp.Filter = "类别='婚姻状况'"
    cbo婚姻.Clear
    For i = 1 To rsTmp.RecordCount
        cbo婚姻.AddItem rsTmp!编码 & "-" & rsTmp!名称
        If rsTmp!缺省 = 1 Then
            cbo婚姻.ItemData(cbo婚姻.NewIndex) = 1
            cbo婚姻.ListIndex = cbo婚姻.NewIndex
        End If
        rsTmp.MoveNext
    Next

    '职业
    rsTmp.Filter = "类别='职业'"
    cbo职业.Clear
    For i = 1 To rsTmp.RecordCount
        cbo职业.AddItem rsTmp!编码 & "-" & rsTmp!名称
        If rsTmp!缺省 = 1 Then
            cbo职业.ItemData(cbo职业.NewIndex) = 1
            cbo职业.ListIndex = cbo职业.NewIndex
        End If
        rsTmp.MoveNext
    Next
    
    '84313,李南春,2015/4/27,联系人关系以及其他关系
    '社会关系
    rsTmp.Filter = "类别='社会关系'"
    cbo联系人关系.Clear
    For i = 1 To rsTmp.RecordCount
        cbo联系人关系.AddItem rsTmp!编码 & "-" & rsTmp!名称
        If rsTmp!缺省 = 1 Then
            cbo联系人关系.ItemData(cbo联系人关系.NewIndex) = 1
            cbo联系人关系.ListIndex = cbo联系人关系.NewIndex
        End If
        rsTmp.MoveNext
    Next
        
    '过敏药物查找结果列表初始化
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 1400, 0
        .Add , "编码", "编码", 900
        .Add , "单位", "单位", 600
        .Add , "剂型", "剂型", 600
        .Add , "毒理分类", "毒理分类", 900
        .Add , "新药", "新药", 600
        .Add , "皮试", "皮试", 600
    End With
    
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").index - 1
        .SortOrder = lvwAscending
        .Visible = False
    End With
    '问题号:56599
    Call Init过敏药物
    
    If cbo年龄单位.Tag = "" Then
        cbo年龄单位.Tag = cbo年龄单位.Text
    End If
End Sub

Public Sub ShowMe(bytMode As Byte, frmParent As Object)
    
    Set mfrmMain = frmParent
    
    If gblnAutoAddName And mbytFun = 1 And mstrCard <> "" Then '刷卡自动产生"新病人"
        txtPatient.Text = "新病人"
        mbln基本信息调整 = False
        Call cmdOK_Click
    Else
        If mlngOutModeMC > 0 Then
            If mlngOutModeMC = 920 Then
                txtPatiMCNO(0).MaxLength = 12
            Else
                txtPatiMCNO(0).MaxLength = 30
            End If
            txtPatiMCNO(0).ToolTipText = "最大长度" & txtPatiMCNO(0).MaxLength & "位"
            txtPatiMCNO(1).MaxLength = txtPatiMCNO(0).MaxLength
        End If
        Call NewCardObject  '47007
        If txt门诊号.Text <> "" Then
            mintNOLength = Len(txt门诊号.Text)
        End If
        '82859:李南春,2015/4/8,病人基本信息调整
        mbln基本信息调整 = Not (mlng病人ID <> 0 And InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";基本信息调整;") = 0)
        txtPatient.Enabled = mbln基本信息调整: txt出生日期.Enabled = mbln基本信息调整: txt出生时间.Enabled = mbln基本信息调整
        txt年龄.Enabled = mbln基本信息调整 And Not mbln禁止输入年龄: cbo年龄单位.Enabled = mbln基本信息调整 And Not mbln禁止输入年龄: cbo性别.Enabled = mbln基本信息调整
        txt身份证号.Enabled = mbln基本信息调整
        'Call SetPatiBaseModiPropertyEanbled
        Me.Show bytMode, frmParent
    End If
    
    Call CloseIDCard    '47007
End Sub

Private Sub Form_Load()
    mblnChange = True
    txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
    mblnNewPatient = False
    
    Call InitData
    Call CreateObjectKeyboard
    '创建病人信息公共部件
    '69026,冉俊明,2014-8-8,检查输入年龄
    Call CreatePublicPatient
    mbln家庭地址输入 = Val(Nvl(zlDatabase.GetPara("家庭地址输入方式", glngSys, mlngModul, 1), 1)) = 1
    mblnCheckNOValidity = Val(Nvl(zlDatabase.GetPara("门诊号有效性检查", glngSys, mlngModul, 1), 1)) = 1
    
    mblnStructAdress = Val(zlDatabase.GetPara(251, glngSys)) <> 0 '病人地址结构化录入
    mblnShowTown = Val(zlDatabase.GetPara(252, glngSys)) <> 0 '乡镇地址结构化录入
    
    Call InitTagPage
    Call InitTaskPanelOther
    
    txtRegLocation.MaxLength = glngMax户口地址
    txtBirthLocation.MaxLength = glngMax出生地点
    '初始化地址控件
    If Not mblnStructAdress Then Exit Sub
    padd家庭地址.Visible = True: padd户口地址.Visible = True
    padd家庭地址.ShowTown = mblnShowTown: padd户口地址.ShowTown = mblnShowTown
    cbo家庭地址.Visible = False: cmd家庭地址.Visible = False
    padd家庭地址.Top = cbo家庭地址.Top: padd家庭地址.Left = cbo家庭地址.Left
    txtRegLocation.Visible = False: cmdRegLocation.Visible = False
    padd户口地址.Top = txtRegLocation.Top: padd户口地址.Left = txtRegLocation.Left
    
    padd家庭地址.MaxLength = glngMax家庭地址
    padd户口地址.MaxLength = glngMax户口地址
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrCard = ""
    mblnChange = False
    Set mdic医疗卡属性 = Nothing
    Set mobjKeyboard = Nothing
    Call CloseIDCard
    mblnPlugin = False
    mlngPlugInHwnd = 0: mblnSavePati = False
    '74430,冉俊明,2014-7-7,挂号中的病人信息编辑功能中提供采集照片功能
    mlng图像操作 = 0: mstr采集图片 = ""
    If Not mobjPubPatient Is Nothing Then Set mobjPubPatient = Nothing
    mblnGetBirth = False
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim i As Integer
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    For i = 1 To msh过敏.Rows - 1
        If i <> msh过敏.Row Then
            If msh过敏.RowData(i) = Replace(lvwItems.SelectedItem.Key, "_", "") Then
                MsgBox "第 " & i & " 行的药物已经与你选择的药物种类相同,请重新选择！", vbInformation, gstrSysName
                lvwItems.SetFocus
                Exit Sub
            End If
        End If
    Next
    lvwItems.Visible = False
    msh过敏.RowData(msh过敏.Row) = Replace(lvwItems.SelectedItem.Key, "_", "")
    msh过敏.TextMatrix(msh过敏.Row, 0) = Trim(lvwItems.SelectedItem.Text)
    msh过敏.SetFocus
    msh过敏_EnterCell
    
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    Case vbKeyEscape
        lvwItems.Visible = False
        txt过敏.Visible = True
        txt过敏.SetFocus
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msh过敏_Click()
    msh过敏_EnterCell
End Sub

Private Sub msh过敏_GotFocus()
    msh过敏_EnterCell
End Sub

Private Sub msh过敏_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If KeyCode = vbKeyF4 Then msh过敏_DblClick
    If KeyCode = vbKeyF3 And cmd过敏.Visible Then cmd过敏_Click
    If KeyCode = vbKeyDelete Then
        msh过敏.TextMatrix(msh过敏.Row, 0) = ""
        msh过敏.RowData(msh过敏.Row) = 0
        For i = msh过敏.Row + 1 To msh过敏.Rows - 1
            msh过敏.TextMatrix(i - 1, 0) = msh过敏.TextMatrix(i, 0)
            msh过敏.RowData(i - 1) = msh过敏.RowData(i)
            msh过敏.TextMatrix(i, 0) = ""
            msh过敏.RowData(i) = 0
        Next
        msh过敏_EnterCell
    End If
End Sub

Private Sub msh过敏_DblClick()
    cmd过敏.Visible = False
    txt过敏.Visible = False
    
    'If msh过敏.Row > 1 And msh过敏.TextMatrix(msh过敏.Row - 1, 0) = "" Or msh过敏.RowData(msh过敏.Row) = 0 Then Exit Sub
    Select Case msh过敏.Col
        Case 0 '过敏药物
            txt过敏.Top = msh过敏.CellTop + msh过敏.Top + (msh过敏.CellHeight - txt过敏.Height) / 2 - 15
            txt过敏.Left = msh过敏.Left + msh过敏.CellLeft + 30
            txt过敏.Width = msh过敏.CellWidth - 60
            
            txt过敏.Text = msh过敏.TextMatrix(msh过敏.Row, msh过敏.Col)
            txt过敏.ZOrder
            Call zlControl.TxtSelAll(txt过敏)
            txt过敏.Visible = True
            If txt过敏.Visible Then txt过敏.SetFocus
        Case 1 '过敏反应
            txt过敏反应.Top = msh过敏.CellTop + msh过敏.Top + (msh过敏.CellHeight - txt过敏反应.Height) / 2 - 15
            txt过敏反应.Left = msh过敏.Left + msh过敏.CellLeft + 30
            '75446:李南春,2014-7-16,过敏反应文本框不够
            txt过敏反应.Width = msh过敏.CellWidth - 60
            
            txt过敏反应.Text = msh过敏.TextMatrix(msh过敏.Row, msh过敏.Col)
            txt过敏反应.ZOrder
            Call zlControl.TxtSelAll(txt过敏反应)
            txt过敏反应.Visible = True
            If txt过敏反应.Visible Then txt过敏反应.SetFocus
    End Select
End Sub

Private Sub msh过敏_EnterCell()
    cmd过敏.Visible = False
    txt过敏.Visible = False
    
    '问题号:56599
    If msh过敏.Row > 1 And msh过敏.TextMatrix(msh过敏.Row - 1, 0) = "" Or msh过敏.Col = 1 Then Exit Sub
    
    cmd过敏.Top = msh过敏.CellTop + msh过敏.Top - 15
    If msh过敏.Rows < 5 Then
        cmd过敏.Left = msh过敏.Left + msh过敏.CellWidth - cmd过敏.Width + 45
    Else
        cmd过敏.Left = msh过敏.Left + msh过敏.CellWidth - cmd过敏.Width + 45
    End If
    
    cmd过敏.ZOrder
    cmd过敏.Visible = True
End Sub

Private Sub msh过敏_KeyPress(KeyAscii As Integer)
        If KeyAscii <> 13 Then
            'If msh过敏.Row > 1 And msh过敏.TextMatrix(msh过敏.Row - 1, 0) = "" Or msh过敏.RowData(msh过敏.Row) = 0 Then Exit Sub
            msh过敏_DblClick
            If msh过敏.Col = 0 Then msh过敏.RowData(msh过敏.Row) = 0
            If msh过敏.Col = 0 Then txt过敏.Text = Chr(KeyAscii)
            If msh过敏.Col = 0 Then txt过敏.SelStart = Len(txt过敏.Text)
            '75446:李南春,2014-7-16,编辑过敏记录时重新激活文本框
            If msh过敏.Col = 1 Then txt过敏反应.Text = txt过敏反应.Text & Chr(KeyAscii)
            If msh过敏.Col = 1 Then txt过敏反应.SelStart = Len(txt过敏反应.Text)
        Else
             If msh过敏.Row = msh过敏.Rows - 1 And msh过敏.TextMatrix(msh过敏.Row, 0) <> "" Then
                msh过敏.Rows = msh过敏.Rows + 1
                msh过敏.Row = msh过敏.Rows - 1
                '问题号:56599
                txt过敏反应.Text = ""
                txt过敏反应.Visible = False
                
                msh过敏_EnterCell
            ElseIf msh过敏.TextMatrix(msh过敏.Row, 0) <> "" Then
                msh过敏.Row = msh过敏.Row + 1
                msh过敏_EnterCell
            Else
                cmdOK.SetFocus
            End If
        End If
End Sub
Private Sub msh过敏_Scroll()
    cmd过敏.Visible = False
    '问题号:56599
    txt过敏.Visible = False
    txt过敏反应.Visible = False
End Sub

Private Sub padd户口地址_Change()
    If mblnStructAdress Then mstrPlugChange = mstrPlugChange & ",户口地址"
End Sub

Private Sub padd家庭地址_Change()
    If mblnStructAdress Then mstrPlugChange = mstrPlugChange & ",现住址"
End Sub

Private Sub PicHealth_Resize()
    On Error Resume Next
    With vsOtherInfo
        .Width = PicHealth.ScaleWidth - 30
        .Height = PicHealth.ScaleHeight - .Top - 15
    End With
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    With msh过敏
        .Top = fraUnit.Top + fraUnit.Height + 45
        .Width = picInfo.ScaleWidth - 30
        .Height = picInfo.ScaleHeight - .Top - 15
    End With
End Sub

Private Sub picTaskPanelOther_Resize()
    wndTaskPanelOther.Move 0, 0, picTaskPanelOther.Width, picTaskPanelOther.Height
End Sub

Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    wndTaskPanelOther.Reposition
End Sub

Private Sub txtBirthLocation_Change()
    mstrPlugChange = mstrPlugChange & ",出生地址"
    txtBirthLocation.Tag = ""
End Sub

Private Sub txtBirthLocation_GotFocus()
    Call zlControl.TxtSelAll(txtBirthLocation)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub SearchAddress(ByVal strInput As String, txtInput As Object)
    '--------------------------------------------------------------
    '功能:模糊查找，弹出地区选择列表
    '编制:冉俊明
    '日期:2014-5-23
    '参数:
    '   strInput:输入文本，若为空表示点击按钮进入
    '   txtInput:文本框对象
    '--------------------------------------------------------------
    Dim strSQL As String, strWhere As String
    Dim strKey As String, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset, vRect As RECT
    
    On Error GoTo Errhand
    If strInput <> "" And txtInput.Tag <> "" Then Exit Sub
    vRect = zlControl.GetControlRect(txtInput.Hwnd)
    If strInput = "" Then '点击按钮
        strSQL = "" & _
            "Select ID, 上级id, 编码, 名称, 末级 " & _
            "From (With 地区_t As" & _
            "    (Select Rownum As 行号, ID, 上级id, 末级, 编码, 名称" & _
            "     From (Select Distinct Substr(名称, 1, 2) As ID, Null As 上级id, 0 As 末级, Null As 编码, Substr(名称, 1, 2) As 名称" & _
            "            From 地区" & _
            "            Union All" & _
            "            Select 编码 As ID, Substr(名称, 1, 2) As 上级id, 1 As 末级, 编码, 名称 From 地区))" & _
            "   Select 行号 As ID, To_Number(上级id) As 上级id, 编码, 名称, 末级 From 地区_t Where 上级id Is Null" & _
            "   Union All" & _
            "   Select b.行号, a.行号, b.编码, b.名称, b.末级 From 地区_t A, 地区_t B Where a.Id = b.上级id Order By 编码)"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "地区", False, _
                       "", "", False, False, False, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False)
    Else
        '去掉"'"
        strInput = Replace(strInput, "'", " ")
        strKey = GetMatchingSting(strInput, False)
        If strInput <> "" Then
            If IsNumeric(strInput) Then '输入全是数字时只匹配编码
                strWhere = " Where 编码 Like Upper([1])"
            ElseIf zlCommFun.IsCharAlpha(strInput) Then '输入全是字母时只匹配简码
                strWhere = " Where 简码 Like Upper([1])"
            Else
                strWhere = " Where 编码 Like Upper([1]) Or 名称 Like [1] Or 简码 Like Upper([1])"
            End If
        End If
        
        strSQL = "" & _
            "Select Rownum As ID, 编码, 名称 From 地区 " & strWhere & " Order By 编码"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "地区", False, _
                       "", "", False, False, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False, strKey)
    End If
    If blnCancel Then txtInput.SetFocus: Exit Sub

    If rsTemp Is Nothing Then txtInput.SetFocus: Exit Sub
    If rsTemp.State <> 1 Then txtInput.SetFocus: Exit Sub
    
    txtInput.Text = Nvl(rsTemp!名称)
    txtInput.Tag = Nvl(rsTemp!ID)
    txtInput.SelStart = Len(Nvl(txtInput.Text))
    txtInput.SetFocus
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtBirthLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    '73022,冉俊明,2014-5-20,在单位名称、出生地点、户口地址加上模糊查找功能
    If KeyCode = vbKeyReturn And Trim(txtBirthLocation.Text) <> "" Then
        Call SearchAddress(Trim(txtBirthLocation.Text), txtBirthLocation)
    End If
End Sub

Private Sub txtBirthLocation_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtMobile_Validate(Cancel As Boolean)
    If txtMobile.Text <> "" And IsMobileNO(txtMobile.Text) = False Then
        MsgBox "输入的手机号格式不正确，请重新录入！", vbInformation, gstrSysName
        Cancel = True
        Exit Sub
    End If
    If Exist手机号(txtMobile.Text, IIf(mlng病人ID <> 0, mlng病人ID, 0)) Then
        If MsgBox("输入的手机号与其他病人重复，是否确定录入？", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Cancel = True
    End If
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If mblnNameChange = True And mlng病人ID = 0 Then zlQueryEMPIPatiInfo
    mblnNameChange = False
End Sub

Private Sub txtPatiMCNO_Change(index As Integer)
    If index = 0 Then mstrPlugChange = mstrPlugChange & ",医保号"
End Sub

Private Sub txtRegLocation_Change()
    If Not mblnStructAdress Then mstrPlugChange = mstrPlugChange & ",户口地址"
    txtRegLocation.Tag = ""
End Sub

Private Sub txtRegLocation_GotFocus()
    Call zlControl.TxtSelAll(txtRegLocation)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtRegLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    '73022,冉俊明,2014-5-20,在单位名称、出生地点、户口地址加上模糊查找功能
    If KeyCode = vbKeyReturn And Trim(txtRegLocation.Text) <> "" Then
        Call SearchAddress(Trim(txtRegLocation.Text), txtRegLocation)
    End If
End Sub

Private Sub txtRegLocation_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtPatient_Change()
    mstrPlugChange = mstrPlugChange & ",姓名"
    If mobjIDCard Is Nothing Then Exit Sub
    If Not mobjIDCard Is Nothing And Not txtPatient.Locked Then mobjIDCard.SetEnabled (txtPatient.Text = "")
End Sub

Private Sub txtPatiMCNO_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtPatiMCNO_Validate(index As Integer, Cancel As Boolean)
    txtPatiMCNO(index).Text = UCase(Trim(txtPatiMCNO(index).Text))
    If cbo付款方式.ListCount > 0 Then cbo付款方式.ListIndex = 0

    If index = 1 Then
        If txtPatiMCNO(1).Text <> txtPatiMCNO(0).Text Then
            MsgBox "请检查,两次输入的医保号不一致！", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
    End If
    
    If mlngOutModeMC = 920 And txtPatiMCNO(0).Text <> txtPatiMCNO(0).Tag And txtPatiMCNO(0).Text <> "" Then
        If CheckExistsMCNO(txtPatiMCNO(0).Text) Then
            Cancel = True
        End If
    End If
End Sub

Private Sub txt出生日期_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If txt出生日期.Text = "____-__-__" Then
           zlCommFun.PressKey (vbKeyTab) '跳过时间
           zlCommFun.PressKey (vbKeyTab)
       Else
           zlCommFun.PressKey (vbKeyTab)
       End If
    End If

End Sub

Private Sub txt出生时间_Change()
    Dim str出生时间 As String
    If txt出生时间.Text <> txt出生时间.Tag And InStr(mstrPlugChange, "出生日期") = 0 Then mstrPlugChange = mstrPlugChange & ",出生日期"
    '76669，李南春,2014-8-18,病人年龄更新
    If IsDate(txt出生日期.Text) Then
        str出生时间 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        txt年龄.Text = ReCalcOld(CDate(str出生时间), cbo年龄单位)
        txt年龄.Tag = txt年龄.Text
    End If
End Sub

Private Sub txt出生时间_GotFocus()
    zlControl.TxtSelAll txt出生时间
End Sub

Private Sub txt出生时间_KeyPress(KeyAscii As Integer)
    If Not IsDate(txt出生日期.Text) Then
        KeyAscii = 0
        txt出生时间.Text = "__:__"
    End If
End Sub


Private Sub txt出生时间_Validate(Cancel As Boolean)
    If txt出生时间.Text <> "__:__" And Not IsDate(txt出生时间.Text) Then
        txt出生时间.SetFocus
        Cancel = True
    End If
End Sub

Private Sub txt出生日期_Change()
    Dim str出生时间 As String
    If txt出生日期.Text <> txt出生日期.Tag And InStr(mstrPlugChange, "出生日期") = 0 Then mstrPlugChange = mstrPlugChange & ",出生日期"
    If IsDate(txt出生日期.Text) And mblnChange Then
        mblnChange = False
        txt出生日期.Text = Format(CDate(txt出生日期.Text), "yyyy-mm-dd") '0002-02-02自动转换为2002-02-02,否则,看到的是2002,实际值却是0002
        mblnChange = True
        
        str出生时间 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        txt年龄.Text = ReCalcOld(CDate(str出生时间), cbo年龄单位)
        txt年龄.Tag = txt年龄.Text
        cbo年龄单位.Tag = cbo年龄单位.Text
        mblnGetBirth = False
    End If
End Sub
Private Sub txt出生日期_GotFocus()
    zlControl.TxtSelAll txt出生日期
End Sub

Private Sub txt出生日期_LostFocus()
    If txt出生日期.Text <> "____-__-__" And Not IsDate(txt出生日期.Text) Then
      If txt出生日期.Enabled And txt出生日期.Visible Then txt出生日期.SetFocus
    End If
End Sub


Private Sub txt单位电话_GotFocus()
    Call zlControl.TxtSelAll(txt单位电话)
End Sub

Private Sub txt单位电话_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt单位电话, KeyAscii
End Sub

Private Sub txt单位名称_Change()
     mstrPlugChange = mstrPlugChange & ",单位名称"
    txt单位名称.Tag = ""
End Sub

Private Sub txt单位名称_GotFocus()
    Call zlControl.TxtSelAll(txt单位名称)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt单位名称_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And cmd单位名称.Enabled And cmd单位名称.Visible Then cmd单位名称_Click
    '73022,冉俊明,2014-5-20,在单位名称、出生地点、户口地址加上模糊查找功能
    If KeyCode = vbKeyReturn And Trim(txt单位名称.Text) <> "" Then
        Call SearchUnit(Trim(txt单位名称.Text), txt单位名称)
    End If
End Sub

Private Sub SearchUnit(ByVal strInput As String, txtInput As Object)
    '--------------------------------------------------------------
    '功能:模糊查找，弹出合约单位选择列表
    '编制:冉俊明
    '日期:2014-5-23
    '参数:
    '   strInput:输入文本，若为空表示点击按钮进入
    '   txtInput:文本框对象
    '--------------------------------------------------------------
    Dim strSQL As String, strWhere As String
    Dim strKey As String, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset, vRect As RECT
    
    On Error GoTo Errhand
    If strInput <> "" And txtInput.Tag <> "" Then Exit Sub
    vRect = zlControl.GetControlRect(txtInput.Hwnd)
    If strInput = "" Then '点击按钮
        strSQL = "" & _
        "       Select ID,上级ID,末级,编码,名称,地址,电话,开户银行,帐号,联系人 From  合约单位" & _
        "       Where 撤档时间 Is Null Or 撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD')" & _
        "       Start With 上级ID is NULL" & _
        "       Connect by Prior ID=上级ID"
        '75888,冉俊明,2014-7-28
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "单位", False, _
                       "", "", False, True, False, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False)
    Else
        '去掉"'"
        strInput = Replace(strInput, "'", " ")
        strKey = GetMatchingSting(strInput, False)
        If strInput <> "" Then
            If IsNumeric(strInput) Then '输入全是数字时只匹配编码
                strWhere = " Where 编码 Like Upper([1])"
            ElseIf zlCommFun.IsCharAlpha(strInput) Then '输入全是字母时只匹配简码
                strWhere = " Where 简码 Like Upper([1])"
            Else
                strWhere = " Where 编码 Like Upper([1]) Or 名称 Like [1] Or 简码 Like Upper([1])"
            End If
        End If
        
        strSQL = "" & _
        "       Select ID,上级ID,末级,编码,名称,地址,电话,开户银行,帐号,联系人 From  合约单位" & strWhere & _
        "       And (撤档时间 Is Null Or 撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD'))"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "单位", False, _
                       "", "", False, False, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False, strKey)
    End If
    If blnCancel Then txtInput.SetFocus: Exit Sub

    If rsTemp Is Nothing Then txtInput.SetFocus: Exit Sub
    If rsTemp.State <> 1 Then txtInput.SetFocus: Exit Sub
    
    txtInput.Text = Nvl(rsTemp!名称)
    txtInput.Tag = Nvl(rsTemp!ID)
    txtInput.SelStart = Len(Nvl(txtInput.Text))
    txtInput.SetFocus
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt单位名称_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt单位名称, KeyAscii
End Sub

Private Sub txt单位名称_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt单位邮编_Change()
    mstrPlugChange = mstrPlugChange & ",单位邮编"
End Sub

Private Sub txt单位邮编_GotFocus()
    Call zlControl.TxtSelAll(txt单位邮编)
End Sub

Private Sub txt单位邮编_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt单位邮编, KeyAscii
End Sub

Private Sub txt过敏_Change()
    '75286:李南春，2014-7-16，自由录入过敏药物
    msh过敏.TextMatrix(msh过敏.Row, 0) = txt过敏.Text
End Sub

Private Sub txt过敏_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim ObjItem As ListItem
    Dim strSQL As String
            
    If KeyAscii <> 13 Then
        If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If KeyAscii <> vbKeyEscape Then msh过敏.RowData(msh过敏.Row) = 0
    Else
        KeyAscii = 0

        strSQL = " Select Distinct A.ID,A.编码," & _
        " A.名称,A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类," & _
        " Decode(B.是否新药,1,'√','') as 新药,Decode(B.是否皮试,1,'√','') as 皮试" & _
        " From 诊疗项目目录 A,药品特性 B,诊疗项目别名 C" & _
        " Where A.类别 IN('5','6','7') And A.ID=B.药名ID And A.Id=C.诊疗项目id" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
        " And (C.名称 like [1] OR A.编码 like [1] OR C.简码 like [1])"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrLike & UCase(txt过敏.Text) & "%")
        
        With rsTmp
            If .BOF Or .EOF Then
                msh过敏.SetFocus: msh过敏_EnterCell
                Exit Sub
            Else
                Me.lvwItems.ListItems.Clear
                Do While Not .EOF
                    Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称, , IIf(!皮试 <> "", 1, 2))
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").index - 1) = !编码
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("单位").index - 1) = IIf(IsNull(!单位), "", !单位)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("剂型").index - 1) = IIf(IsNull(!剂型), "", !剂型)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("毒理分类").index - 1) = IIf(IsNull(!毒理分类), "", !毒理分类)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("新药").index - 1) = IIf(IsNull(!新药), "", !新药)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("皮试").index - 1) = IIf(IsNull(!皮试), "", !皮试)
                    .MoveNext
                Loop
                Me.lvwItems.ListItems(1).Selected = True
            End If
        End With
        
        With Me.lvwItems
            .Left = msh过敏.Left
            .Width = msh过敏.Width
            .Height = msh过敏.Height + 300
            If msh过敏.Rows < 5 Then
                .Top = msh过敏.Top + msh过敏.RowHeight(msh过敏.Row) * (msh过敏.Row) - .Height
            Else
                .Top = msh过敏.Top + msh过敏.RowHeight(4) * (3) - .Height
            End If
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
'75446:李南春,2014-7-16,编辑过敏记录时重新激活文本框
Private Sub txt过敏_LostFocus()
    txt过敏.Visible = False
End Sub

Private Sub txt过敏反应_Change()
   '问题号:56599
   msh过敏.TextMatrix(msh过敏.Row, 1) = txt过敏反应.Text
End Sub
'75446:李南春,2014-7-16,编辑过敏记录时重新激活文本框
Private Sub txt过敏反应_LostFocus()
    txt过敏反应.Visible = False
End Sub

Private Sub txt户口地址邮编_Change()
    mstrPlugChange = mstrPlugChange & ",户口邮编"
    mblnChange = True
End Sub

Private Sub txt户口地址邮编_GotFocus()
    Call zlControl.TxtSelAll(txt户口地址邮编)
End Sub

Private Sub txt户口地址邮编_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt家庭电话_Change()
    mstrPlugChange = mstrPlugChange & ",电话"
End Sub

Private Sub txt家庭邮编_Change()
    mstrPlugChange = mstrPlugChange & ",现住址邮编"
End Sub

Private Sub txt家庭电话_GotFocus()
    Call zlControl.TxtSelAll(txt家庭电话)
End Sub

Private Sub txt家庭电话_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt家庭电话, KeyAscii
End Sub

Private Sub txt家庭邮编_GotFocus()
    Call zlControl.TxtSelAll(txt家庭邮编)
End Sub

Private Sub txt家庭邮编_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt家庭邮编, KeyAscii
End Sub

Private Sub txt联系人姓名_Change()
    mstrPlugChange = mstrPlugChange & ",联系人姓名"
End Sub

Private Sub txt监护人_GotFocus()
    zlCommFun.OpenIme (True)
End Sub

Private Sub txt监护人_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txt卡号_GotFocus()
    '72686:李南春,2015/3/25,将卡号全选
    Call zlControl.TxtSelAll(txt卡号)
    Call SetBrushCardObject(True)
End Sub

Private Sub txt卡号_KeyPress(KeyAscii As Integer)
    
    Dim blnCard  As Boolean
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    blnCard = zlCommFun.InputIsCard(txt卡号, KeyAscii, gCurSendCard.str卡号密文 <> "")
    If blnCard And Len(txt卡号.Text) = gCurSendCard.lng卡号长度 - 1 And KeyAscii <> 8 Then
        txt卡号.Text = txt卡号.Text & Chr(KeyAscii): txt卡号.SelStart = Len(txt卡号.Text)
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt卡号_LostFocus()
    Call SetBrushCardObject(False)
End Sub

Private Sub txt卡号_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '72686:李南春,2015/3/25,将卡号全选
    Call zlControl.TxtSelAll(txt卡号)
End Sub

Private Sub txt卡号_Validate(Cancel As Boolean)
    Dim lngPatientID As Long
    Dim lng变动类型 As Long
    Dim blnCardBind As Boolean  '卡是否进行绑定
    If gCurSendCard.lng卡号长度 = Len(Trim(txt卡号.Text)) Then
        If Bln已发卡(txt卡号.Text, gCurSendCard.lng卡类别ID, lngPatientID) Then
            If gCurSendCard.bln自制卡 And gCurSendCard.bln重复使用 And lngPatientID > 0 Then
                lng变动类型 = GetCardLastChangeType(txt卡号.Text, gCurSendCard.lng卡类别ID, lngPatientID)
                If lng变动类型 = 11 Then
                    '如果是绑定
                    If MsgBox("卡号为【" & txt卡号.Text & "】的{" & gCurSendCard.str卡名称 & "}的卡已经与病人标识为【" & lngPatientID & "】的进行了绑定！" & vbCrLf & "是否取消该卡的绑定?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                        Cancel = True
                        txt卡号.Text = ""
                        Exit Sub
                    End If
                    If BlandCancel(gCurSendCard.lng卡类别ID, Trim(txt卡号.Text), lngPatientID) Then
                        Exit Sub
                    End If
                End If
            End If

            MsgBox "该卡号已经被绑定,不能绑定该卡号.", vbInformation, gstrSysName
            Cancel = True
            txt卡号.Text = ""
            Exit Sub
        End If
    End If

    If mbytFun = 1 And mstrCard = "" And Trim(txt卡号.Text) <> "" Then
        If GetPatientState(txt卡号.Text) <> 0 Then
            MsgBox "该卡号的持卡病人正在就诊或等待就诊,不能绑定该卡号.", vbInformation, gstrSysName
            Cancel = True
        End If
        If Not mfrmMain Is Nothing Then
            Call mfrmMain.zlReadPlugInPati(Trim(txt卡号))
        End If
        If Not gCurSendCard.bln自制卡 Then
            '42947
            If zlLoadInfor = False Then Cancel = True
        End If
    End If

    If gCurSendCard.str卡名称 = "二代身份证" Then Exit Sub
End Sub

Private Function CheckPatiValid(ByVal strCard As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查指定输入的卡号是否合法
    '入参：strCard-指定的卡号
    '返回：合法,返回True,否则返回False
    '编制：刘兴洪
    '日期：2010-07-19 10:14:31
    '说明：31182
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String, lng病人ID As Long
    
    '69954:刘尔旋,2014-02-19,门诊挂号管理无法发已经被退卡或者取消绑定卡的问题
    '72324:刘尔旋,2014-04-24,发卡挂号失败时重新挂号不能发卡的问题
    '74894:李南春,2014-07-08,卡号信息检索错误
    strSQL = "Select Nvl(a.就诊状态, 0) 就诊状态, a.病人id, a.姓名, a.性别" & vbNewLine & _
             "From 病人信息 A, 病人医疗卡信息 B, 医疗卡类别 C" & vbNewLine & _
             "Where a.就诊卡号 = b.卡号 And c.特定项目 = '就诊卡' And b.卡类别id = c.Id And b.卡类别id=[1] And b.卡号 = [2]"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gCurSendCard.lng卡类别ID, strCard)
    If rsTmp.RecordCount = 0 Then CheckPatiValid = True: Exit Function
    
    '1.检查状态:原来主要是在输就诊卡时进行检查的,由于txt卡号_Validate事情,不一定能检查到,因此,本增加在按确定时,增加该检查
    If Val(Nvl(rsTmp!就诊状态)) <> 0 Then
        MsgBox "卡号为" & strCard & "的病人正在就诊或等待就诊,不能绑定该卡号.", vbInformation, gstrSysName
        Exit Function
    End If
    
    '2.检查是否病人姓名相同
    If Nvl(rsTmp!姓名) <> Trim(txtPatient.Text) And Val(txt卡号.Tag) = 0 Then
       If MsgBox("持卡病人『" & Nvl(rsTmp!姓名) & "』与输入的病人『" & Trim(txtPatient.Text) & "』不一致,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    '3.挂号病人与刷就诊卡得出的病人是两个不同建档的病人
    lng病人ID = Val(Nvl(rsTmp!病人ID))
    If Val(txt卡号.Tag) <> lng病人ID And Val(txt卡号.Tag) <> 0 Then
        If Nvl(rsTmp!姓名) <> Trim(txtPatient.Text) Then
            If MsgBox("注意: " & vbCrLf & _
                             "     持卡病人『" & Nvl(rsTmp!姓名) & "』与输入的病人『" & Trim(txtPatient.Text) & "』不一致," & vbCrLf & _
                             "     但同时都是建档病人,是否将病人『" & Trim(txtPatient.Text) & "』合并到病人『" & Nvl(rsTmp!姓名) & "』中?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            '合并
            If zlPatiMerge(Val(txt卡号.Tag), lng病人ID, True) = False Then Exit Function
        Else '病人姓名相同,自动进行合并
            '自动合并
            If zlPatiMerge(Val(txt卡号.Tag), lng病人ID, False) = False Then Exit Function
        End If
        '重新刷新相关的数据
        RaiseEvent PatiMerged(lng病人ID)
        
    End If
    CheckPatiValid = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPatientState(strCard As String, Optional lng病人ID As Long) As Long
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查病人的当前状态
    '入参：strCard-卡号
    '出参：lng病人ID-返回病人ID
    '返回：返回病人状态
    '编制：刘兴洪
    '日期：2010-07-19 09:55:25
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strPassWord As String, strErrMsg As String
    '42947
    If gobjSquare.objSquareCard.zlGetPatiID(gCurSendCard.lng卡类别ID, strCard, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
    If lng病人ID = 0 Then Exit Function
    
    strSQL = "Select Nvl(就诊状态,0) 就诊状态,病人ID From 病人信息 Where 病人ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    lng病人ID = 0
    If rsTmp.RecordCount > 0 Then
        lng病人ID = Val(Nvl(rsTmp!病人ID))
        GetPatientState = rsTmp!就诊状态
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt联系人电话_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("电话") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("电话")) = txt联系人电话.Text
    End If
End Sub

Private Sub txt联系人身份证_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt联系人身份证_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("身份证号") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("身份证号")) = txt联系人身份证.Text
    End If
End Sub

Private Sub txt联系人姓名_GotFocus()
    zlCommFun.OpenIme (True)
End Sub

Private Sub txt联系人姓名_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txt联系人姓名_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("姓名") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("姓名")) = txt联系人姓名.Text
        If vsLinkMan.Rows = vsLinkMan.FixedRows + 1 And txt联系人姓名.Text <> "" Then
            vsLinkMan.Rows = vsLinkMan.Rows + 1
        End If
    End If
End Sub

Private Sub txt门诊号_Change()
    mstrPlugChange = mstrPlugChange & ",门诊号"
End Sub

Private Sub txt门诊号_GotFocus()
    Call zlControl.TxtSelAll(txt门诊号)
End Sub

Private Sub txt门诊号_KeyPress(KeyAscii As Integer)
     
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If txt门诊号.Enabled And txt门诊号.Visible And mintNOLength > 0 Then
        '如果手工输入了异常的门诊号则提示
            If Len(txt门诊号.Text) > mintNOLength + 1 Then
                MsgBox "注意,输入的门诊号过大,请确认是否输入正常!", vbInformation, gstrSysName
                txt门诊号.SetFocus
                txt门诊号.SelStart = 0: txt门诊号.SelLength = Len(txt门诊号.Text)
                Exit Sub
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = 32 Then
        KeyAscii = 0
        If txt门诊号.Text = "" Then
            txt门诊号.Text = zlDatabase.GetNextNo(3)
            mintNOLength = Len(Trim(txt门诊号.Text))
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Or InStr(";" & mstrPrivs & ";", ";允许修改门诊号;") = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt密码_Change()
    txt验证.Enabled = txt密码.Text <> ""
    If txt密码.Text = "" Then txt验证.Text = ""
    If gCurSendCard.str卡名称 = "二代身份证" Then
        txt支付密码.Text = txt密码.Text
    End If
End Sub

Private Sub txt密码_GotFocus()
    Call zlControl.TxtSelAll(txt密码)
    Call OpenPassKeyboard(txt密码, False)
End Sub

Private Sub txt密码_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, gCurSendCard.int密码规则 = 1)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt密码.Text = "" Then
            txt验证.Text = ""
            If tbcPage.Selected.index = 0 Then
                '104243:李南春,2016/12/29,焦点定位时检查是否可用
                If txtPatient.Visible And txtPatient.Enabled Then
                    txtPatient.SetFocus
                Else
                    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
                End If
            End If
        Else
            txt验证.SetFocus
        End If
    End If
End Sub
Private Sub txt密码_LostFocus()
    Call ClosePassKeyboard(txt密码)
End Sub
Private Sub CheckInputPassWord(KeyAscii As Integer, Optional ByVal blnOnlyNum As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查密码输入
    '编制:刘兴洪
    '日期:2011-07-07 00:40:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If KeyAscii = 8 Or KeyAscii = 13 Then Exit Sub
    If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If blnOnlyNum Then
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            KeyAscii = 0
        End If
        Exit Sub
    End If
    If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
       If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
            If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                 If InStr(1, "!@#$%^&*()_+-=><?,:;~`./", Asc(KeyAscii)) = 0 Then KeyAscii = 0
            End If
       End If
    End If
End Sub

Private Sub txt年龄_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt年龄.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt年龄.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt年龄_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt年龄.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt年龄_GotFocus()
    Call zlCommFun.OpenIme
    Call zlControl.TxtSelAll(txt年龄)
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    Dim blnTab As Boolean
    
    If KeyAscii = vbKeyReturn Then
        If cbo年龄单位.Visible = False And IsNumeric(txt年龄.Text) Then
            Call txt年龄_Validate(False)
            Call cbo年龄单位.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt年龄.Text) And cbo年龄单位.Visible Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        '仅仅限制几个 指定的特殊的字符 问题:49908
        If InStr("~·！@#￥%……&*（）——-+=|、？、。，~`!#$%^&*()-_=+|\/?<>,/<>", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt年龄_Validate(Cancel As Boolean)
    Dim strBirth As String
    txt年龄.Text = Trim(txt年龄.Text)
    If Not IsNumeric(txt年龄.Text) And Trim(txt年龄.Text) <> "" Then
        cbo年龄单位.ListIndex = -1: cbo年龄单位.Visible = False: txt年龄.Width = 1485
    ElseIf cbo年龄单位.Visible = False Then
        cbo年龄单位.ListIndex = 0: cbo年龄单位.Visible = True: txt年龄.Width = 690
    End If
    If txt年龄.Text <> txt年龄.Tag Then
        mblnChange = False
        If Not IsDate(txt出生日期.Text) Then mblnGetBirth = True
'        txt出生日期.Text = ReCalcBirth(Trim(txt年龄.Text), IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""))
        If mblnGetBirth Then
            If mobjPubPatient.ReCalcBirthDay(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), strBirth) Then
                txt出生日期.Text = Format(strBirth, "yyyy-mm-dd")
                txt出生时间.Text = Format(strBirth, "hh:mm")
            End If
        End If
        mblnChange = True
        txt年龄.Tag = txt年龄.Text
    End If
    '69026,冉俊明,2014-8-8,检查输入年龄
    '76703,冉俊明,2014-8-15
    If cbo年龄单位.Visible Then Exit Sub
    If mobjPubPatient Is Nothing Then Exit Sub
    If mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text), _
            IIf(txt出生日期.Text = "____-__-__", "", txt出生日期.Text) & _
            IIf(txt出生时间.Text = "__:__", "", " " & txt出生时间.Text)) = False Then
        Cancel = True
    End If
End Sub

Private Sub txt其他关系_GotFocus()
    Call zlControl.TxtSelAll(txt单位名称)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt其他关系_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt其他关系_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("附加信息") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("关系")) = NeedName(cbo联系人关系.Text)
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("附加信息")) = txt其他关系.Text
    End If
End Sub

Private Sub txt区域_Change()
    txt区域.Tag = ""
End Sub

Private Sub txt区域_GotFocus()
    zlCommFun.OpenIme (True)
End Sub

Private Sub txt区域_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If (txt区域.Tag <> "" Or Trim(txt区域.Text) = "") Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If zl_SelectAndNotAddItem(Me, txt区域, Trim(txt区域.Text), "区域", "区域选择", True, False) = False Then
        Exit Sub
    End If
End Sub

Private Sub txt区域_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txt身份证号_Change()
    mstrPlugChange = mstrPlugChange & ",身份证号"
    If mbln扫描身份证签约 And ActiveControl Is txt身份证号 And Not mobjIDCard Is Nothing Then
            mobjIDCard.SetEnabled txt身份证号.Text = ""
    End If
End Sub

Private Sub txt身份证号_GotFocus()
    Call zlControl.TxtSelAll(txt身份证号)
'    If gCurSendCard.str卡名称 <> "二代身份证" Then
'        Call OpenIDCard
'    End If
    If mbln扫描身份证签约 = True And txt身份证号.Text = "" Then
        OpenIDCard
    End If
End Sub
Private Sub txt身份证号_KeyPress(KeyAscii As Integer)
    
'    If gCurSendCard.str卡名称 <> "二代身份证" Then
'        If zl当前用户身份证是否绑定(Trim(txt身份证号.Text), Trim(txtPatient.Text), Trim(txt门诊号.Text)) = True Then
'            MsgBox "当前用户的身份证号已经绑定，不允许修改其身份证号", vbInformation, gstrSysName
'            KeyAscii = 0
'        End If
'    End If
    
    If zl当前用户身份证是否绑定(Trim(txt身份证号.Text), Trim(txtPatient.Text), Trim(txt门诊号.Text)) = True Then
        MsgBox "当前用户的身份证号已经绑定，不允许修改其身份证号", vbInformation, gstrSysName
        KeyAscii = 0
    End If

    mbln扫描身份证 = False
    txt支付密码.Text = ""
    txt验证密码.Text = ""
    SetCtrVisibleAndMove
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt身份证号, KeyAscii
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPatient_GotFocus()
    Call zlControl.TxtSelAll(txtPatient)
    Call zlCommFun.OpenIme(True)
    
    If mbln扫描身份证签约 = True And txt身份证号.Text = "" Then
        OpenIDCard
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        '新病人才调用
        If mblnNameChange = True And mlng病人ID = 0 Then zlQueryEMPIPatiInfo
        mblnNameChange = False
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        mblnNameChange = True
    End If
    CheckLen txtPatient, KeyAscii
End Sub

Private Sub txtPatient_LostFocus()
    Call zlCommFun.OpenIme
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)

End Sub

Private Sub txt身份证号_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
    
End Sub

Private Sub txt身份证号_Validate(Cancel As Boolean)
    '65663:刘尔旋,2014-02-20,根据身份证号计算出生日期
    If IsDate(zlCommFun.GetIDCardDate(txt身份证号.Text)) = False Then Exit Sub
    If Format(zlCommFun.GetIDCardDate(txt身份证号.Text), "yyyy-mm-dd") <> Format(txt出生日期.Text, "yyyy-mm-dd") Then
        If IsDate(txt出生日期.Text) Then MsgBox "输入的身份证号与输入的出生日期不一致，将使用身份证号获取的日期替换！", vbInformation, gstrSysName
        txt出生日期.Text = zlCommFun.GetIDCardDate(txt身份证号.Text)
    End If
End Sub

Private Sub txt验证_Change()
    If gCurSendCard.str卡名称 = "二代身份证" Then
        txt验证密码.Text = txt验证.Text
    End If
End Sub

Private Sub txt验证_GotFocus()
    Call zlControl.TxtSelAll(txt验证)
    Call OpenPassKeyboard(txt验证, True)
End Sub

Private Function GetOneDept(lng收费细目ID As Long) As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.执行科室ID From 收费项目目录 A,收费执行科室 B Where B.收费细目ID=A.ID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng收费细目ID)
    If Not rsTmp.EOF Then
        GetOneDept = rsTmp!执行科室ID '默认取第一个(如有多个)
    Else
        GetOneDept = UserInfo.部门ID '如没有指定，则取操作员所在科室
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建密码创建
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional bln确认密码 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, bln确认密码) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Sub txt验证_LostFocus()
    Call ClosePassKeyboard(txt验证)
End Sub

Private Function GetCardDataSql(ByVal byt变动类型 As Byte, ByVal lng病人ID As Long, ByVal lng卡类别ID As Long, _
   ByVal str原卡号 As String, ByVal strCard As String, ByVal str密码 As String, ByVal dtCurdate As Date, _
   ByVal strICCard As String, Optional ByVal str变动原因 As String = "")
    Dim strSQL As String
    Dim strPassWord As String
    strPassWord = zlCommFun.zlStringEncode(str密码)
    'Zl_医疗卡变动_Insert
     strSQL = "Zl_医疗卡变动_Insert("
    '      变动类型_In   Number,
    '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定);
    '５-密码调整(只记录);6-挂失(16取消挂失)
    strSQL = strSQL & "" & byt变动类型 & ","
    '      病人id_In     住院费用记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
    strSQL = strSQL & "" & lng卡类别ID & ","
    '      原卡号_In     病人医疗卡信息.卡号%Type,
    strSQL = strSQL & "'" & str原卡号 & "',"
    '      医疗卡号_In   病人医疗卡信息.卡号%Type,
    strSQL = strSQL & "'" & strCard & "',"
    '      变动原因_In   病人医疗卡变动.变动原因%Type,
    '      --变动原因_In:如果密码调整，变动原因为密码.加密的
    strSQL = strSQL & "'" & str变动原因 & "',"
    '      密码_In       病人信息.卡验证码%Type,
    strSQL = strSQL & "'" & strPassWord & "',"
    '      操作员姓名_In 住院费用记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '      变动时间_In   住院费用记录.登记时间%Type,
    strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
    strSQL = strSQL & "'" & strICCard & "',"
    '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
    strSQL = strSQL & IIf(str变动原因 = "", "NULL)", "'" & str变动原因 & "')")
    GetCardDataSql = strSQL
End Function


Private Function zlLoadInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人信息
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-11-04 10:50:46
    '问题:42947
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strOutCardNO As String, strOutPatiInforXML As String
    Dim strExpand As String
    Dim strPatiXml As String
    On Error GoTo errHandle
    
    If gCurSendCard.lng卡类别ID = 0 Then zlLoadInfor = True: Exit Function
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
   '问题号:53408
    If gCurSendCard.str卡名称 <> "二代身份证" Then
         mbln扫描身份证 = True
         If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, gCurSendCard.lng卡类别ID, False, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Function
    End If
    If strOutCardNO = "" Then Exit Function
    txt卡号.Text = strOutCardNO
    If strOutPatiInforXML = "" Then zlLoadInfor = True: Exit Function
    zlLoadInfor = LoadPati(strOutPatiInforXML)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadPati(strXmlCardInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人信息
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-11-04 10:50:46
    '问题:42947
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strOutCardNO As String
    Dim objNode As MSXML2.IXMLDOMElement, strExpand As String
    Dim objTempNode As MSXML2.IXMLDOMElement
    Dim strPatiXml As String
    Dim strTmp As String, strValue As String
    Dim i As Long, j As Long, lngCount As Long, lngChildCount As Long '问题号:56599
    Dim str过敏药物 As String, str过敏反应 As String '问题号:56599
    Dim str接种日期 As String, str接种名称 As String '问题号:56599
    Dim strABO血型 As String '问题号:56599
    Dim str信息名 As String, str信息值 As String '问题号:56599
    Dim xmlChildNodes As IXMLDOMNodeList, xmlChildNode As IXMLDOMNode '问题号:56599
    Dim str姓名 As String, str关系 As String, str电话 As String, str身份证号 As String, str地址 As String '问题号:56599
    Dim str其他关系 As String
    On Error GoTo errHandle
    
    If strXmlCardInfor = "" Then LoadPati = True: Exit Function
    '加载病人信息
    If zlXML_Init = False Then Exit Function
    
    '101269:李南春,2016/10/8,传入变量错误
    If zlXML_LoadXMLToDOMDocument(strXmlCardInfor, False) = False Then Exit Function
    '    标识    数据类型    长度    精度    说明
    '    卡号    Varchar2    20
    Call zlXML_GetNodeValue("卡号", , strValue)
    '    姓名    Varchar2    100
    Call zlXML_GetNodeValue("姓名", , strValue)
    txtPatient.Text = strValue
    '    性别    Varchar2    4
    Call zlXML_GetNodeValue("性别", , strValue)
    If strValue <> "" Then
        Call zlControl.CboLocate(cbo性别, strValue)
        If cbo性别.ListIndex = -1 Then
            cbo性别.AddItem strValue
            cbo性别.ListIndex = cbo性别.NewIndex
        End If
    End If
    '    年龄    Varchar2    10
    Call zlXML_GetNodeValue("年龄", , strValue)
    If strValue <> "" Then
        Call LoadOldData(strValue, txt年龄, cbo年龄单位)
    End If
    '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    Call zlXML_GetNodeValue("出生日期", , strValue)
    If strValue <> "" Then
        txt出生日期.Text = Format(IIf(IsDate(strValue) = False, "____-__-__", strValue), "YYYY-MM-DD")
        If IsDate(strValue) Then txt出生时间 = Format(CDate(strValue), "HH:MM")
        txt年龄.Text = ReCalcOld(CDate(txt出生日期.Text), cbo年龄单位)      '修改的时候,根据出生日期重算年龄
        txt年龄.Tag = txt年龄.Text
    Else
         txt出生时间.Text = "__:__"
         txt出生日期.Text = ReCalcBirth(Val(txt年龄.Text), cbo年龄单位.Text)
    End If
    cbo年龄单位.Tag = cbo年龄单位.Text
    
    '    出生地点    Varchar2    50
    Call zlXML_GetNodeValue("出生地点", , strValue)
    '    身份证号    VARCHAR2    18
    Call zlXML_GetNodeValue("身份证号", , strValue)
    If strValue <> "" Then
        txt身份证号.Text = strValue
        If InStr(1, txt出生日期.Text, "__") > 0 Then
            strTmp = zlCommFun.GetIDCardDate(txt身份证号.Text)
            If IsDate(strTmp) Then txt出生日期.Text = strTmp
        End If
    End If
    '    其他证件    Varchar2    20
    Call zlXML_GetNodeValue("其他证件", , strValue)
   ' If strValue <> "" Then txt其他证件.Text = strValue
    '    职业    Varchar2    80
    Call zlXML_GetNodeValue("职业", , strValue)
    If strValue <> "" Then
        cbo职业.ListIndex = cbo.FindIndex(cbo职业, strValue)
        If cbo职业.ListIndex = -1 Then
            cbo职业.AddItem strValue, 0
            cbo职业.ListIndex = cbo职业.NewIndex
        End If
    End If
    '    民族    Varchar2    20
    Call zlXML_GetNodeValue("民族", , strValue)
    cbo民族.ListIndex = cbo.FindIndex(cbo民族, strValue, True)
     If cbo民族.ListIndex = -1 And strValue <> "" Then
         cbo民族.AddItem strValue, 0
         cbo民族.ListIndex = cbo民族.NewIndex
     End If
    '    国籍    Varchar2    30
    Call zlXML_GetNodeValue("国籍", , strValue)
    cbo国籍.ListIndex = cbo.FindIndex(cbo国籍, strValue, True)
     If cbo国籍.ListIndex = -1 And strValue <> "" Then
         cbo国籍.AddItem strValue, 0
         cbo国籍.ListIndex = cbo国籍.NewIndex
     End If
    '    学历    Varchar2    10
    Call zlXML_GetNodeValue("学历", , strValue)
    'cbo学历.ListIndex = GetCboIndex(cbo学历, strValue)
'    If cbo学历.ListIndex = -1 And strValue <> "" Then
'        cbo学历.AddItem strValue, 0
'        cbo学历.ListIndex = cbo学历.NewIndex
'    End If
    '    婚姻状况    Varchar2    4
    Call zlXML_GetNodeValue("婚姻状况", , strValue)
    cbo婚姻.ListIndex = cbo.FindIndex(cbo婚姻, strValue, True)
     If cbo婚姻.ListIndex = -1 And strValue <> "" Then
         cbo婚姻.AddItem strValue, 0
         cbo婚姻.ListIndex = cbo婚姻.NewIndex
     End If
    '    区域    Varchar2    30
    Call zlXML_GetNodeValue("区域", , strValue)
    txt区域.Text = strValue
    '    家庭地址    Varchar2    50
    Call zlXML_GetNodeValue("家庭地址", , strValue)
   cbo家庭地址.Text = strValue
   padd家庭地址.Value = strValue
    '    户口地址    Varchar2    50
    Call zlXML_GetNodeValue("户口地址", , strValue)
    txtRegLocation.Text = strValue
    padd户口地址.Value = strValue
    '    家庭电话    Varchar2    20
    Call zlXML_GetNodeValue("家庭电话", , strValue)
   txt家庭电话.Text = strValue
    '    家庭地址邮编    Varchar2    6
    Call zlXML_GetNodeValue("家庭地址邮编", , strValue)
   txt家庭邮编.Text = strValue
    '    监护人  Varchar2    64
    Call zlXML_GetNodeValue("监护人", , strValue)
   'txt监护人.Text = strValue
'    '    联系人姓名  Varchar2    64
'    Call zlXML_GetNodeValue("联系人姓名", , strValue)
'    txt联系人姓名.Text = strValue '问题号:40005
'    '    联系人关系  Varchar2    30
'    Call zlXML_GetNodeValue("联系人关系", , strValue)
'    txt联系人关系.Text = strValue '问题号:40005
'    '    联系人地址  Varchar2    50
'    Call zlXML_GetNodeValue("联系人地址", , strValue)
'    '    联系人电话  Varchar2    20
'    Call zlXML_GetNodeValue("联系人电话", , strValue)
'    txt联系人电话.Text = strValue '问题号:40005
    '    工作单位    Varchar2    100
    Call zlXML_GetNodeValue("工作单位", , strValue)
    txt单位名称.Text = strValue
    lbl单位名称.Tag = ""
    '    单位电话    Varchar2    20
    Call zlXML_GetNodeValue("单位电话", , strValue)
   txt单位电话.Text = strValue
    '    单位邮编    Varchar2    6
    Call zlXML_GetNodeValue("单位邮编", , strValue)
   txt单位邮编.Text = strValue
    '    单位开户行  Varchar2    50
    Call zlXML_GetNodeValue("单位开户行", , strValue)
   'txt单位开户行.Text = strValue
    '    单位帐号    Varchar2    20
    Call zlXML_GetNodeValue("单位帐号", , strValue)
   'txt单位帐号.Text = strValue
   '问题号:56599
    '过敏情况
    Call zlXML_GetRows("药物名称", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("药物名称", i, str过敏药物)
        Call zlXML_GetNodeValue("药物反应", i, str过敏反应)
        SetDrugAllergy str过敏药物, str过敏反应
    Next
    lngCount = 0
    '免疫记录
    Call zlXML_GetRows("疫苗名称", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("疫苗名称", i, str接种名称)
        Call zlXML_GetNodeValue("接种时间", i, str接种日期)
        SetInoculate str接种日期, str接种名称
    Next
    lngCount = 0
    'ABO血型
    Call zlXML_GetNodeValue("ABO血型", , strABO血型)
    If strABO血型 <> "" Then
        For i = 0 To cboBloodType.ListCount - 1
            '76314,李南春，2014-08-06，病人信息正确获取
            If NeedName(cboBloodType.List(i), , ".") = NeedName(strABO血型) Then cboBloodType.ListIndex = i
        Next
    End If
    'RH
    Call zlXML_GetNodeValue("RH", , strValue)
    If strValue <> "" Then
        For i = 0 To cboBH.ListCount - 1
            If cboBH.List(i) = strValue Then cboBH.ListIndex = i
        Next
    End If
    '医学警示
    strValue = ""
    Set xmlChildNodes = zlXML_GetChildNodes("临床基本信息")
    If Not xmlChildNodes Is Nothing Then
        If xmlChildNodes.length > 0 Then
            For i = 0 To xmlChildNodes.length - 1
                Set xmlChildNode = xmlChildNodes(i)
                If xmlChildNode.Text = "1" Then
                    strValue = strValue & "," & Replace(xmlChildNode.nodeName, "标志", "")
                End If
            Next
        End If
    End If
    If strValue <> "" Then txtMedicalWarning.Text = Mid(strValue, 2)
   
    
    '其他医学警示
    Call zlXML_GetNodeValue("其他医学警示", , strValue)
    If strValue <> "" Then txtOtherWaring.Text = strValue
    '联系信息
    '    联系人地址  Varchar2    50
    Call zlXML_GetNodeValue("联系人地址", , str地址)
    'txt联系人地址.Text = str地址
     '    联系人姓名  Varchar2    64
    Call zlXML_GetNodeValue("联系人姓名", , str姓名)
    '    联系人关系  Varchar2    30
    Call zlXML_GetNodeValue("联系人关系", , str关系)
    '    联系人电话  Varchar2    20
    Call zlXML_GetNodeValue("联系人电话", , str电话)
    '    联系人身份证 Varchar2   20
    Call zlXML_GetNodeValue("联系人身份证号", , str身份证号)
    '84313,李南春,2015/4/27,联系人关系以及其他关系
    Call zlXML_GetNodeValue("联系人附加信息", , str其他关系)
    SetLinkInfo str姓名, str关系, str电话, str身份证号, str其他关系
    
    Call zlXML_GetRows("联系信息", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("联系信息", "姓名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("联系信息", "姓名", i, j, str姓名)
                Call zlXML_GetChildNodeValue("联系信息", "关系", i, j, str关系)
                Call zlXML_GetChildNodeValue("联系信息", "电话", i, j, str电话)
                Call zlXML_GetChildNodeValue("联系信息", "身份证号", i, j, str身份证号)
                Call zlXML_GetChildNodeValue("联系信息", "附加信息", i, j, str其他关系)
                SetLinkInfo str姓名, str关系, str电话, str身份证号, str其他关系
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0

    '其他信息
    '健康档案编号
    Call zlXML_GetNodeValue("健康档案编号", , strValue)
    SetOtherInfo "健康档案编号", strValue
    
    '新农合证号
    Call zlXML_GetNodeValue("新农合证号", , strValue)
    SetOtherInfo "新农合证号", strValue

    '其他证件
    Call zlXML_GetRows("其他证件", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("其他证件", "信息名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("其他证件", "信息名", i, j, str信息名)
                Call zlXML_GetChildNodeValue("其他证件", "信息值", i, j, str信息值)
                SetOtherInfo str信息名, str信息值
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    '其他信息
    Call zlXML_GetRows("其他信息", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("其他信息", "信息名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("其他信息", "信息名", i, j, str信息名)
                Call zlXML_GetChildNodeValue("其他信息", "信息值", i, j, str信息值)
                SetOtherInfo str信息名, str信息值
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    '医疗卡属性
    Call zlXML_GetRows("医疗卡属性", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("医疗卡属性", "信息名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("医疗卡属性", "信息名", i, j, str信息名)
                Call zlXML_GetChildNodeValue("医疗卡属性", "信息值", i, j, str信息值)
                If mdic医疗卡属性.Exists(str信息名) Then
                    mdic医疗卡属性.Item(str信息名) = str信息值
                Else
                    mdic医疗卡属性.Add str信息名, str信息值
                End If
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    
    '从卡上获取病人信息后,调用EMPI
    Call zlQueryEMPIPatiInfo
    LoadPati = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub CloseIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:关闭自助读卡功能
    '编制:刘兴洪
    '日期:2012-03-09 16:26:40
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
Private Sub NewCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化新的卡对象
    '编制:刘兴洪
    '日期:2012-03-09 16:28:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gblnNewCardNoPop Then Exit Sub
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.Hwnd)
    End If
    If Not mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.setParaent(Me.Hwnd)
    End If
End Sub
Private Sub OpenIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开身份证读卡器
    '编制:王吉
    '日期:2012-08-31 16:28:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '初始化对卡对象
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.Hwnd)
    End If
    '打开读卡器
    mobjIDCard.SetEnabled (True)
End Sub

Public Function zl_Get设置默认发卡密码() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置默认发卡密码
    '返回:是否继续发卡操作
    '编制:王吉
    '日期:2012-07-06 15:53:14
    '问题号:51072
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCardType As clsCard
    Dim msgResult As VbMsgBoxResult
    Dim arr() As String
    arr = zl_Get医疗卡类型(gCurSendCard.lng卡类别ID)
    If Val(arr(2)) = 0 Then '无限制
        Select Case Val(arr(1))
            Case 0 '无限制
                zl_Get设置默认发卡密码 = True
                Exit Function
            Case 1 '未输入提醒
               msgResult = MsgBox("未输入密码将会影响帐户的使用安全,是否继续？", vbQuestion + vbYesNo, gstrSysName)
               zl_Get设置默认发卡密码 = IIf(msgResult = vbYes, True, False)
               Exit Function
            Case 2 '为输入禁止
                 MsgBox "未输入卡密码,不能进行发卡！", vbExclamation, gstrSysName
                zl_Get设置默认发卡密码 = False
                Exit Function
        End Select
    ElseIf Val(arr(2)) = 1 Then '缺省身份证后N位
        If Len(Trim(txt身份证号.Text)) > 0 Or Len(Trim(txt联系人身份证.Text)) > 0 Then '输入了身份证或联系人身份证号
            If Len(Trim(txt身份证号.Text)) > 0 Then '有身份证优先用身份证
                   txt密码.Text = Right(Trim(txt身份证号.Text), Val(arr(0)))
            Else '否则就用代办人身份证作为密码
                   txt密码.Text = Right(Trim(txt联系人身份证.Text), Val(arr(0)))
            End If
        Else '身份证与联系人身份证都没输入
            Select Case Val(arr(1))
                Case 0 '无限制
                    zl_Get设置默认发卡密码 = True
                    Exit Function
                Case 1 '未输入提醒
                    msgResult = MsgBox("未输入密码将会影响帐户的使用安全,是否继续！", vbQuestion + vbYesNo, gstrSysName)
                    zl_Get设置默认发卡密码 = IIf(msgResult = vbYes, True, False)
                    Exit Function
                Case 2 '为输入禁止
                    MsgBox "未输入卡密码,不能进行发卡！", vbExclamation, gstrSysName
                    zl_Get设置默认发卡密码 = False
                    Exit Function
            End Select
        End If
    End If
    zl_Get设置默认发卡密码 = True
End Function

Private Function zl_Get缺省发卡类别() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取缺省发卡类别
    '返回:缺省发卡类别名称
    '编制:王吉
    '日期:2012-08-31 11:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lngCardTypeID As Long
    Dim rsTemp As Recordset
    
    On Error GoTo ErrHandl:

    strSQL = "" & _
    "   Select Id, 编码, 名称, 短名, 前缀文本, 卡号长度, 缺省标志, 是否固定, 是否严格控制, " & _
    "           nvl(是否自制,0) as 是否自制, nvl(是否存在帐户,0) as 是否存在帐户, " & _
    "           nvl(是否全退,0) as 是否全退,nvl(是否重复使用,0) as 是否重复使用 , " & _
    "           nvl(密码长度,10) as 密码长度,nvl(密码长度限制,0) as 密码长度限制,nvl(密码规则,0) as 密码规则," & _
    "           nvl(是否退现,0) as 是否退现,部件, 备注, 特定项目, 结算方式, 是否启用, 卡号密文,Nvl(密码输入限制,0) as 密码输入限制,Nvl(是否缺省密码,0) as 是否缺省密码," & _
    "           nvl(是否模糊查找,0) as 是否模糊查找,nvl(读卡性质,'1000') as 读卡性质 " & _
    "    From 医疗卡类别" & _
    "    Where ID = [1]" & _
    "    Order by 编码"

    lngCardTypeID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, , , True))
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCardTypeID)
    If rsTemp Is Nothing Then zl_Get缺省发卡类别 = "": Exit Function
    If rsTemp.RecordCount <= 0 Then zl_Get缺省发卡类别 = "": Exit Function
    zl_Get缺省发卡类别 = rsTemp!名称
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub SetCtrVisibleAndMove()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控制显示和位置
    '编制:王吉
    '日期:2012-08-31 11:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str卡类别名称 As String
    Dim lng变化长度 As Long
    Dim lng间隔 As Long
    
    lng间隔 = 100
    
    '按钮区域
    cmdHelp.Top = Me.ScaleHeight - 500
    cmdCancel.Top = cmdHelp.Top
    cmdOK.Top = cmdHelp.Top
    
    '默认发卡是二代身份证时,绑卡控件不可用
    If gCurSendCard.str卡名称 Like "二代身份证" Then
        txt卡号.Enabled = False: txt密码.Enabled = False: txt验证.Enabled = False
        lblICCard.Enabled = False: lbl密码.Enabled = False: lbl验证.Enabled = False
    End If
    
    If picCard.Visible Then
        tbcPage.Top = picCard.Top + picCard.Height
    Else
        tbcPage.Top = picCard.Top
    End If
    tbcPage.Height = Me.ScaleHeight - tbcPage.Top - (Me.ScaleHeight - cmdHelp.Top + 45)
       
    If mlngOutModeMC = 0 Then
        lblPatiMCNO(0).Enabled = False: lblPatiMCNO(1).Enabled = False
        txtPatiMCNO(0).Enabled = False: txtPatiMCNO(1).Enabled = False
    Else
        lblPatiMCNO(0).Enabled = True: lblPatiMCNO(1).Enabled = True
        txtPatiMCNO(0).Enabled = True: txtPatiMCNO(1).Enabled = True
    End If
        
     '扫描的身份证与扫描身份证签约为True的情况下才能绑定身份证
    If mbln扫描身份证 = True And mbln扫描身份证签约 Then
        lbl支付密码.Enabled = True: txt支付密码.Enabled = True
        lbl验证密码.Enabled = True: txt验证密码.Enabled = True
    Else
        '设置支付密码与验证密码不可用
        lbl支付密码.Enabled = False: txt支付密码.Enabled = False: txt支付密码.Text = ""
        lbl验证密码.Enabled = False: txt验证密码.Enabled = False: txt验证密码.Text = "": txt验证密码.Tag = ""
    End If
End Sub

Private Sub txt验证密码_GotFocus()
    Call zlControl.TxtSelAll(txt验证密码)
    Call OpenPassKeyboard(txt验证密码, False)
End Sub

Private Sub txt验证密码_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, gCurSendCard.int密码规则 = 1)
End Sub

Private Sub txt验证密码_LostFocus()
    Call ClosePassKeyboard(txt验证密码)
End Sub
Private Sub txt支付密码_GotFocus()
    Call zlControl.TxtSelAll(txt支付密码)
    Call OpenPassKeyboard(txt支付密码, False)
End Sub

Private Sub txt支付密码_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, gCurSendCard.int密码规则 = 1)
End Sub

Private Sub zl加载病人信息(rsPatiInfo As Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人信息到窗体控件中
    '编制:王吉
    '日期:2012-08-31 11:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    '病人姓名
    txtPatient.Text = Nvl(rsPatiInfo!姓名)
    mlng病人ID = Nvl(Val(rsPatiInfo!病人ID))
    If Nvl(rsPatiInfo!性别) <> "" Then
        Call zlControl.CboLocate(cbo性别, rsPatiInfo!性别)
        If cbo性别.ListIndex = -1 Then
            cbo性别.AddItem rsPatiInfo!性别
            cbo性别.ListIndex = cbo性别.NewIndex
        End If
    End If
    '年龄
    If Nvl(rsPatiInfo!年龄) <> "" Then
        Call LoadOldData(rsPatiInfo!年龄, txt年龄, cbo年龄单位)
    End If
    '出生日期
    If Nvl(rsPatiInfo!出生日期) <> "" Then
        txt出生日期.Text = Format(IIf(IsDate(rsPatiInfo!出生日期) = False, "____-__-__", rsPatiInfo!出生日期), "YYYY-MM-DD")
        If IsDate(rsPatiInfo!出生日期) Then txt出生时间 = Format(CDate(rsPatiInfo!出生日期), "HH:MM")
        txt年龄.Text = ReCalcOld(CDate(txt出生日期.Text), cbo年龄单位)      '修改的时候,根据出生日期重算年龄
        txt年龄.Tag = txt年龄.Text
    Else
         txt出生时间.Text = "__:__"
         txt出生日期.Text = ReCalcBirth(Val(txt年龄.Text), cbo年龄单位.Text)
    End If
    cbo年龄单位.Tag = cbo年龄单位.Text

    '身份证号
    If Nvl(rsPatiInfo!身份证号) <> "" Then
        txt身份证号.Text = rsPatiInfo!身份证号
        If InStr(1, txt出生日期.Text, "__") > 0 Then
            strTmp = zlCommFun.GetIDCardDate(txt身份证号.Text)
            If IsDate(strTmp) Then txt出生日期.Text = strTmp
        End If
    End If
    '职业
    If Nvl(rsPatiInfo!职业) <> "" Then
        cbo职业.ListIndex = cbo.FindIndex(cbo职业, rsPatiInfo!职业)
        If cbo职业.ListIndex = -1 Then
            cbo职业.AddItem rsPatiInfo!职业, 0
            cbo职业.ListIndex = cbo职业.NewIndex
        End If
    End If
    '民族
    cbo民族.ListIndex = cbo.FindIndex(cbo民族, Nvl(rsPatiInfo!民族), True)
     If cbo民族.ListIndex = -1 And Nvl(rsPatiInfo!民族) <> "" Then
         cbo民族.AddItem rsPatiInfo!民族, 0
         cbo民族.ListIndex = cbo民族.NewIndex
     End If
    '国籍
    cbo国籍.ListIndex = cbo.FindIndex(cbo国籍, Nvl(rsPatiInfo!国籍), True)
     If cbo国籍.ListIndex = -1 And Nvl(rsPatiInfo!国籍) <> "" Then
         cbo国籍.AddItem rsPatiInfo!国籍, 0
         cbo国籍.ListIndex = cbo国籍.NewIndex
     End If
    '婚姻状况
    cbo婚姻.ListIndex = cbo.FindIndex(cbo婚姻, Nvl(rsPatiInfo!婚姻状况), True)
     If cbo婚姻.ListIndex = -1 And Nvl(rsPatiInfo!婚姻状况) <> "" Then
         cbo婚姻.AddItem rsPatiInfo!婚姻状况, 0
         cbo婚姻.ListIndex = cbo婚姻.NewIndex
     End If
    txt区域.Text = Nvl(rsPatiInfo!区域)
    '家庭地址
    cbo家庭地址.Text = Nvl(rsPatiInfo!家庭地址)
    Call zlReadAddrInfo(padd家庭地址, Val(Nvl(rsPatiInfo!病人ID)), 0, 3, cbo家庭地址.Text)
    '家庭电话
    txt家庭电话.Text = Nvl(rsPatiInfo!家庭电话)
    '家庭地址邮编
    txt家庭邮编.Text = Nvl(rsPatiInfo!家庭地址邮编)
    '户口地址
    txtRegLocation.Text = Nvl(rsPatiInfo!户口地址)
    Call zlReadAddrInfo(padd户口地址, Val(Nvl(rsPatiInfo!病人ID)), 0, 4, txtRegLocation.Text)
    '户口地址邮编
    txt户口地址邮编.Text = Nvl(rsPatiInfo!户口地址邮编)
    '工作单位
    txt单位名称.Text = Nvl(rsPatiInfo!工作单位)
    lbl单位名称.Tag = ""
    '单位电话
    txt单位电话.Text = Nvl(rsPatiInfo!单位电话)
    '单位邮编
    txt单位邮编.Text = Nvl(rsPatiInfo!单位邮编)
    '门诊号
    txt门诊号.Text = Nvl(rsPatiInfo!门诊号)
    '问题号:40005
    '联系人姓名
    txt联系人姓名.Text = Nvl(rsPatiInfo!联系人姓名)
    '联系人电话
    txt联系人电话.Text = Nvl(rsPatiInfo!联系人电话)
    '84313,李南春,2015/4/27,联系人关系以及其他关系
    '联系人关系
    txt其他关系.Text = ""
    cbo联系人关系.ListIndex = cbo.FindIndex(cbo联系人关系, Nvl(rsPatiInfo!联系人关系), True)
    If cbo联系人关系.ListIndex = -1 And Nvl(rsPatiInfo!联系人关系) <> "" Then
        cbo联系人关系.ListIndex = 8: txt其他关系.Text = Nvl(rsPatiInfo!联系人关系)
    End If
    '手机号
    txtMobile.Text = Nvl(rsPatiInfo!手机号)
    '问题号:56599
    Load健康卡相关信息 (Val(Nvl(rsPatiInfo!病人ID, "0")))
    '90875:李南春,2016/11/8,医疗卡证件类型
    LoadCertificate (Val(Nvl(rsPatiInfo!病人ID)))
    
    mstr年龄 = txt年龄.Text & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, "")
    mstr性别 = NeedName(cbo性别.Text)
    mstr姓名 = txtPatient.Text
    mstr身份证号 = txt身份证号.Text
    mstr出生日期 = txt出生日期.Text
    mstr出生时间 = txt出生时间.Text
End Sub

Private Sub txt支付密码_LostFocus()
    Call ClosePassKeyboard(txt支付密码)
End Sub
Public Function zl当前用户身份证是否绑定(str身份证号 As String, strName As String, str门诊号 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断当前用户身份证是否已被绑定
    '入参:str身份证号:病人身份证 str门诊号:门诊号
    '返回:True 已绑定 false 未绑定
    '编制:王吉
    '日期:2012-08-31 04:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo Errhand
    strSQL = "" & _
    " Select  姓名,门诊号 From 病人信息 A,病人医疗卡信息 B Where A.病人ID=B.病人ID And B.卡号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "医疗卡绑定", str身份证号)
    If rsTemp Is Nothing Then zl当前用户身份证是否绑定 = False: Exit Function
    If rsTemp.RecordCount <= 0 Then zl当前用户身份证是否绑定 = False: Exit Function
    
    If IIf(IsNull(rsTemp!姓名), "", rsTemp!姓名) = strName And IIf(IsNull(rsTemp!门诊号), "", rsTemp!门诊号) = str门诊号 Then
        zl当前用户身份证是否绑定 = True
    Else
        zl当前用户身份证是否绑定 = False
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function
Private Sub Init过敏药物()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化过敏药物FlexGrid
    '入参:
    '返回:
    '编制:王吉
    '日期:2012-12-20 04:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '设置列以及列标题
    With msh过敏
        .Cols = 2
        .TextMatrix(0, 0) = "过敏药物"
        .TextMatrix(0, 1) = "过敏反应"
        .ColWidth(0) = 5000
        .ColWidth(1) = .Width - 4900
        '75286:李南春，2014-7-16，表格对齐方式
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
    End With
End Sub

Private Sub InitTagPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化分页控件
    '编制:56599
    '日期:2012-12-20 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    
    Err = 0: On Error GoTo Errhand:

        Set ObjItem = tbcPage.InsertItem(mPageIndex.基本, "基本", picInfo.Hwnd, 0)
        ObjItem.Tag = mPageIndex.基本
    
        Set ObjItem = tbcPage.InsertItem(mPageIndex.健康档案, "健康档案", PicHealth.Hwnd, 0)
        ObjItem.Tag = mPageIndex.健康档案
        Call InitVsInoculate
        Call InitVsOtherInfo
        Call InitCombox
        Call InitCertificate
        
        '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
        If CreatePlugInOK(mlngModul) Then
            On Error Resume Next
            mlngPlugInHwnd = gobjPlugIn.GetFormHwnd
            Call zlPlugInErrH(Err, "GetFormHwnd")
            Err.Clear: On Error GoTo 0
            If mlngPlugInHwnd <> 0 Then
                picTaskPanelOther.Visible = True
                Set ObjItem = tbcPage.InsertItem(mPageIndex.附加信息, "附加信息", picTaskPanelOther.Hwnd, 0)
                ObjItem.Tag = mPageIndex.附加信息
            End If
        End If
            
        With tbcPage
            tbcPage.Item(0).Selected = True
            .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
            Set .PaintManager.Font = lblBirthLocation.Font
            .PaintManager.BoldSelected = True
            .PaintManager.Layout = xtpTabLayoutAutoSize
            .PaintManager.StaticFrame = True
            .PaintManager.ClientFrame = xtpTabFrameBorder
            .Height = Me.ScaleHeight - 900
        End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Add健康卡相关信息(ByVal lng病人ID As Long, ByRef colPro As Collection, Optional ByVal lng就诊ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:健康卡数据处理
    '入参:
    '编制:56599
    '日期:2012-12-13 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim strTemp() As String
    Dim strSQL As String
    Dim varKey As Variant
    Dim intCount As Integer
    '过敏药物
    With msh过敏
        If .Rows > 1 Then
            '清除该病人所有记录
            strSQL = " Zl_病人过敏药物_Delete(" & lng病人ID & ")"
            zlAddArray colPro, strSQL
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    '病人过敏药物
                    strSQL = "Zl_病人过敏药物_Update("
                    '病人ID_In 病人过敏药物.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '过敏药物ID_In 病人过敏药物.过敏药物ID%Type
                    strSQL = strSQL & "'" & IIf(.RowData(i) <= 0, "", .RowData(i)) & "',"
                    '过敏药物_In  病人过敏药物.过敏药物%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 0) = "", "", .TextMatrix(i, 0)) & "',"
                    '过敏反应_In 病人过敏反应.过敏反应%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"

                    zlAddArray colPro, strSQL
                End If
            Next
        End If
    End With
    '接种信息
    With vsInoculate
        If .Rows > 1 Then
            '清除该病人所有记录
            strSQL = " Zl_病人免疫记录_Delete(" & lng病人ID & ")"
            zlAddArray colPro, strSQL

            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    '病人过敏药物
                    strSQL = "Zl_病人免疫记录_Update("
                    '病人ID_In 病人免疫记录.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '接种时间_In 病人免疫记录.接种时间%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 0) = "", "''", "to_date('" & .TextMatrix(i, 0) & "','yyyy-mm-dd')") & ","
                    '接种名称_In  病人免疫记录.接种名称%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"
                    zlAddArray colPro, strSQL
                End If
                If .TextMatrix(i, 3) <> "" Then
                    '病人过敏药物
                    strSQL = "Zl_病人免疫记录_Update("
                    '病人ID_In 病人免疫记录.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '接种时间_In 病人免疫记录.接种时间%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 2) = "", "''", "to_date('" & .TextMatrix(i, 2) & "','yyyy-mm-dd')") & ","
                    '接种名称_In  病人免疫记录.接种名称%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "''", .TextMatrix(i, 3)) & "')"
                    zlAddArray colPro, strSQL
                End If
            Next
        End If
    End With
    '其他信息
    'ABO血型
    '病人信息从表
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'血型',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & NeedName(cboBloodType.Text, , ".") & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    'RH
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'RH',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & cboBH.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    '医学警示
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'医学警示',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & txtMedicalWarning.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    '其他医学警示
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'其他医学警示',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & txtOtherWaring.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
        
    '84313:李南春,2015/4/29, 第一条联系人信息已保存在病人信息中，从表中不再重复保存
    '联系人相关信息
    intCount = 0
    With vsLinkMan
        If .Rows >= 3 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then '联系人姓名不能为空
                    intCount = intCount + 1
                    For j = 0 To .Cols - 1
                        strSQL = "Zl_病人信息从表_Update("
                        '病人ID_In 病人信息从表.病人Id%Type
                        strSQL = strSQL & "" & lng病人ID & ","
                        '信息名_In 病人信息从表.信息名%Type
                        strSQL = strSQL & "'联系人" & .TextMatrix(0, j) & intCount & "',"
                        '信息值_In 病人信息从表.信息值%Type
                        strSQL = strSQL & "'" & IIf(.TextMatrix(i, j) = "", "", .TextMatrix(i, j)) & "',"
                        '就诊Id_In 病人信息从表.就诊Id%Type
                        strSQL = strSQL & "'')"

                        zlAddArray colPro, strSQL
                    Next
                End If
            Next
        End If
    End With
    '其他信息
     With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    strSQL = "Zl_病人信息从表_Update("
                    '病人ID_In 病人信息从表.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '信息名_In 病人信息从表.信息名%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 0) & "',"
                    '信息值_In 病人信息从表.信息值%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "',"
                    '就诊Id_In 病人信息从表.就诊Id%Type
                    strSQL = strSQL & "'')"

                    zlAddArray colPro, strSQL
                End If
                If .TextMatrix(i, 2) <> "" Then
                    strSQL = "Zl_病人信息从表_Update("
                    '病人ID_In 病人信息从表.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '信息名_In 病人信息从表.信息名%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 2) & "',"
                    '信息值_In 病人信息从表.信息值%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "", .TextMatrix(i, 3)) & "',"
                    '就诊Id_In 病人信息从表.就诊Id%Type
                    strSQL = strSQL & "'')"

                    zlAddArray colPro, strSQL
                End If
            Next
        End If
     End With
     '医疗卡属性
     If Not mdic医疗卡属性 Is Nothing Then
        For Each varKey In mdic医疗卡属性.Keys
            strSQL = "Zl_病人医疗卡属性_Update("
            strSQL = strSQL & lng病人ID & ","
            strSQL = strSQL & gCurSendCard.lng卡类别ID & ","
            strSQL = strSQL & "'" & Trim(txt卡号.Text) & "',"
            strSQL = strSQL & "'" & varKey & "',"
            strSQL = strSQL & "'" & mdic医疗卡属性(varKey) & "')"
            zlAddArray colPro, strSQL
        Next
     End If
     If lng就诊ID = 0 Then Exit Sub
     'ABO血型
    '病人信息从表
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'血型',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & NeedName(cboBloodType.Text, , ".") & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & lng就诊ID & ")"
    zlAddArray colPro, strSQL
    'RH
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'RH',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & cboBH.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & lng就诊ID & ")"
    zlAddArray colPro, strSQL
End Sub
Private Sub InitVsInoculate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsInoculate
    '初始化列表属性
     vsInoculate.Editable = flexEDKbdMouse
    '设置列头
        SetColumHeader vsInoculate, C_InoculateHeader
    '设置选择按钮
        .ColDataType(0) = flexDTDate
        .ColEditMask(0) = "####-##-##"
        .ColDataType(2) = flexDTDate
        .ColEditMask(2) = "####-##-##"
    End With

End Sub
Private Sub InitVsOtherInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, str关系 As String
    
    With vsLinkMan
    '初始化列表属性
        .Editable = flexEDKbd
    '设置列头
        SetColumHeader vsLinkMan, C_LinkManColumHeader
        For i = 0 To cbo联系人关系.ListCount - 1
            str关系 = str关系 & "|" & NeedName(cbo联系人关系.List(i))
        Next
        str关系 = Mid(str关系, 2)
        If str关系 <> "" Then .ColComboList(.ColIndex("关系")) = str关系
    End With
    With vsOtherInfo
         .Editable = flexEDKbd
    '设置列头
        SetColumHeader vsOtherInfo, C_OtherInfoColumHeader
    End With
End Sub

Private Sub InitCombox()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化ComBox控件
    '编制:56599
    '日期:2012-12-07 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '66743:刘尔旋,2013-11-25,血型与RH默认值的问题
    'ComboBox cboBloodType, C_血型
    zlComboxLoadFromSQL "Select 编码,名称,缺省标志 From 血型", cboBloodType
    mintDefaultBlood = cboBloodType.ListIndex
    ComboBox cboBH, C_BH
    If cboBH.ListCount <> 0 Then cboBH.ListIndex = -1
End Sub

Private Sub ComboBox(objCbo As ComboBox, strSet As String)
    Dim varTemp As Variant
    Dim i As Long
    varTemp = Split(strSet, ",")
    With objCbo
        For i = LBound(varTemp) To UBound(varTemp)
            .AddItem varTemp(i)
        Next
    End With
    If objCbo.ListCount <> 0 Then objCbo.ListIndex = 0
End Sub

Private Sub SetColumHeader(objList As Object, strColumHeader As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置列头
    '参数:objList - 设置对象,strColumHeader - 列表设置字符串
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varSet As Variant
    Dim varColum As Variant
    Dim i As Long
    varSet = Split(strColumHeader, ";")
    If UBound(varSet) = 0 Then Exit Sub
        
    For i = LBound(varSet) To UBound(varSet)
        varColum = Split(varSet(i), ",")
        Select Case TypeName(objList)
            Case "VSFlexGrid"
                With objList
                    .Cols = UBound(varSet) + 1
                    .Cell(flexcpText, 0, i) = varColum(0)
                    .ColKey(i) = varColum(0)
                    .ColAlignment(i) = varColum(1)
                    .ColWidth(i) = varColum(2)
                    .ColHidden(i) = Not (varColum(3) = 1)
                End With
            Case Else
            '暂不处理
        End Select
    Next
End Sub
Private Sub SetDrugAllergy(str过敏药物 As String, str过敏反应 As String, Optional lng过敏ID = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置过敏药物
    '编制:56599
    '日期:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    With msh过敏
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = str过敏药物 Then
                    .TextMatrix(i, 1) = str过敏反应
                    If lng过敏ID <> 0 Then .RowData(i) = lng过敏ID
                    Exit Sub
                End If
            Next
        End If
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str过敏药物
        .TextMatrix(.Rows - 1, 1) = str过敏反应
        If lng过敏ID <> 0 Then .RowData(.Rows - 1) = lng过敏ID
        .Rows = .Rows + 1
    End With
End Sub
Private Sub SetInoculate(str接种日期 As String, str接种名称 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置接种情况
    '编制:56599
    '日期:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    With vsInoculate
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 1 To .Cols - 1 Step 2
                    If .TextMatrix(i, j) = str接种名称 Then
                        .TextMatrix(i, j - 1) = str接种日期
                        Exit Sub
                    End If
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = str接种日期
                .TextMatrix(.Rows - 1, j + 1) = str接种名称
                Exit Sub
            End If
        Next
        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        
    End With
End Sub
Private Sub SetLinkInfo(str姓名 As String, str关系 As String, str电话 As String, str身份证号 As String, str其他关系 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置联系人相关信息
    '编制:56599
    '日期:2012-12-12 09:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    '84313,李南春,2015/4/27,联系人关系以及其他关系
    With vsLinkMan
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = str姓名 And .TextMatrix(i, 2) = str身份证号 Then
                    .TextMatrix(i, 1) = str关系: .TextMatrix(i, 3) = str电话
                    If i = 1 Then
                        txt联系人身份证.Text = str身份证号
                        txt联系人姓名.Text = str姓名
                        cbo联系人关系.ListIndex = cbo.FindIndex(cbo联系人关系, str关系, True)
                        If cbo联系人关系.ListIndex = -1 And str关系 <> "" Then
                            cbo联系人关系.ListIndex = 8: txt其他关系.Text = str关系
                        ElseIf cbo联系人关系.ListIndex = 8 Then
                            txt其他关系.Text = str其他关系
                        End If
                        txt联系人电话.Text = str电话
                    End If
                    Exit Sub
                End If
            Next
        End If
        
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str姓名
        If cbo.FindIndex(cbo联系人关系, str关系, True) = -1 And str关系 <> "" Then
            .TextMatrix(.Rows - 1, 1) = "其他": .TextMatrix(.Rows - 1, 4) = str关系
        Else
            .TextMatrix(.Rows - 1, 1) = str关系
            .TextMatrix(.Rows - 1, 4) = str其他关系
        End If
        .TextMatrix(.Rows - 1, 3) = str电话
        .TextMatrix(.Rows - 1, 2) = str身份证号
        If .Rows - 1 = 1 Then
            txt联系人身份证.Text = str身份证号
            txt联系人姓名.Text = str姓名
            cbo联系人关系.ListIndex = cbo.FindIndex(cbo联系人关系, str关系, True)
            If cbo联系人关系.ListIndex = -1 And str关系 <> "" Then
                cbo联系人关系.ListIndex = 8: txt其他关系.Text = str关系
            ElseIf cbo联系人关系.ListIndex = 8 Then
                txt其他关系.Text = str其他关系
            End If
            txt联系人电话.Text = str电话
        End If
        .Rows = .Rows + 1
    End With
End Sub
Private Sub SetOtherInfo(str信息名 As String, str信息值 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置其他情况
    '编制:56599
    '日期:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1 Step 2
                    If .TextMatrix(i, j) = str信息名 Then
                        .TextMatrix(i, j + 1) = str信息值
                        Exit Sub
                    End If
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = str信息名
                .TextMatrix(.Rows - 1, j + 1) = str信息值
                Exit Sub
            End If
        Next
        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        
    End With
End Sub
Public Sub Load健康卡相关信息(lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人健康卡信息
    '入参:lng病人ID - 病人ID
    '编制:56599
    '日期:2012-12-12 14:55:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs过敏药物 As Recordset
    Dim rs免疫记录 As Recordset
    Dim rsABO血型 As Recordset
    Dim rsRH As Recordset
    Dim rs医学警示 As Recordset
    Dim rs其他医学警示 As Recordset
    Dim rs病人信息 As Recordset
    Dim rs联系人 As Recordset
    Dim rs其他信息 As Recordset
    Dim str医学警示 As String
    Dim str联系人姓名 As String
    Dim str联系人关系 As String
    Dim str联系人电话 As String
    Dim str联系人身份证号 As String
    Dim str附加信息 As String
    Dim lng联系人数量 As Long
    Dim i As Long
    On Error GoTo ErrHandl:
    
    '74430,冉俊明,2014-7-7,挂号中的病人信息编辑功能中提供采集照片功能
    Call ReadPatPricture(lng病人ID)
    
    '获取过敏药物
    strSQL = "" & _
    "   Select 病人ID,过敏药物ID,过敏药物,过敏反应 From 病人过敏药物 Where 病人ID=[1]"
    Set rs过敏药物 = zlDatabase.OpenSQLRecord(strSQL, "病人过敏药物", lng病人ID)
    While rs过敏药物.EOF = False
        SetDrugAllergy Nvl(rs过敏药物!过敏药物), Nvl(rs过敏药物!过敏反应), Nvl(rs过敏药物!过敏药物ID, 0)
        rs过敏药物.MoveNext
    Wend
    '获取免疫记录
    strSQL = "" & _
    "   Select 病人ID,接种时间,接种名称 From 病人免疫记录 Where 病人ID=[1]"
    Set rs免疫记录 = zlDatabase.OpenSQLRecord(strSQL, "病人免疫记录", lng病人ID)
    While rs免疫记录.EOF = False
        SetInoculate Nvl(rs免疫记录!接种时间), Nvl(rs免疫记录!接种名称)
        rs免疫记录.MoveNext
    Wend
    '血型
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='血型' And 就诊ID Is NULL"
    Set rsABO血型 = zlDatabase.OpenSQLRecord(strSQL, "ABO血型", lng病人ID)
    While rsABO血型.EOF = False
        For i = 0 To cboBloodType.ListCount - 1
            '76314,李南春，2014-08-06，病人信息正确获取
            If NeedName(cboBloodType.List(i), , ".") = NeedName(Nvl(rsABO血型!信息值)) Then cboBloodType.ListIndex = i
        Next
        rsABO血型.MoveNext
    Wend
    'RH
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='RH' And 就诊ID Is NULL"
    Set rsRH = zlDatabase.OpenSQLRecord(strSQL, "RH", lng病人ID)
    While rsRH.EOF = False
        For i = 0 To cboBH.ListCount - 1
            If cboBH.List(i) = Nvl(rsRH!信息值) Then cboBH.ListIndex = i
        Next
        rsRH.MoveNext
    Wend
    '医学警示
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='医学警示'"
    Set rs医学警示 = zlDatabase.OpenSQLRecord(strSQL, "医学警示", lng病人ID)
    While rs医学警示.EOF = False
        str医学警示 = str医学警示 & "," & Nvl(rs医学警示!信息值)
        rs医学警示.MoveNext
    Wend
    If str医学警示 <> "" Then str医学警示 = Mid(str医学警示, 2)
    txtMedicalWarning.Text = str医学警示
    '其他医学警示
    strSQL = "" & _
    "  Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='其他医学警示'"
    Set rs其他医学警示 = zlDatabase.OpenSQLRecord(strSQL, "其他医学警示", lng病人ID)
    While rs其他医学警示.EOF = False
        txtOtherWaring.Text = Nvl(rs其他医学警示!信息值)
        rs其他医学警示.MoveNext
    Wend
    '联系人相关信息
    '取病人信息表中的联系人信息
    '84313,李南春,2015/4/27,联系人关系以及其他关系
    strSQL = "" & _
    "   Select  A.联系人姓名,A.联系人关系,A.联系人电话,A.联系人身份证号,B.信息值 as 附加信息 From 病人信息 A,病人信息从表 B " & _
    "   Where A.病人ID=B.病人ID(+) And A.病人ID=[1] And B.信息名(+)='联系人附加信息' And Not A.联系人姓名 is Null"
    Set rs病人信息 = zlDatabase.OpenSQLRecord(strSQL, "病人信息联系人信息", lng病人ID)
    If rs病人信息.EOF = False Then
        txt联系人身份证.Text = Nvl(rs病人信息!联系人身份证号)
        txt联系人姓名.Text = Nvl(rs病人信息!联系人姓名)
        txt联系人电话.Text = Nvl(rs病人信息!联系人电话)
        cbo联系人关系.ListIndex = cbo.FindIndex(cbo联系人关系, Nvl(rs病人信息!联系人关系), True)
        If cbo联系人关系.ListIndex = -1 And Nvl(rs病人信息!联系人关系) <> "" Then
            cbo联系人关系.ListIndex = 8: txt其他关系.Text = rs病人信息!联系人关系
        ElseIf cbo联系人关系.ListIndex = 8 Then
            txt其他关系.Text = Nvl(rs病人信息!附加信息)
        End If
        SetLinkInfo Nvl(rs病人信息!联系人姓名), Nvl(rs病人信息!联系人关系), Nvl(rs病人信息!联系人电话), Nvl(rs病人信息!联系人身份证号), txt其他关系.Text
    End If
    '取病人信息从表中的联系人信息
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名 like '联系人%' order by 信息名 Asc"
    Set rs联系人 = zlDatabase.OpenSQLRecord(strSQL, "联系人相关信息", lng病人ID)
    If rs联系人.EOF = False Then
        rs联系人.Filter = "信息名 like '联系人姓名%'"
        lng联系人数量 = rs联系人.RecordCount
        rs联系人.Filter = ""
        For i = 1 To lng联系人数量 + 1
            While rs联系人.EOF = False
                Select Case Nvl(rs联系人!信息名)
                    Case "联系人姓名" & i
                        str联系人姓名 = Nvl(rs联系人!信息值)
                    Case "联系人关系" & i
                        str联系人关系 = Nvl(rs联系人!信息值)
                    Case "联系人电话" & i
                        str联系人电话 = Nvl(rs联系人!信息值)
                    Case "联系人身份证号" & i
                        str联系人身份证号 = Nvl(rs联系人!信息值)
                    Case "联系人附加信息" & i
                        str附加信息 = Nvl(rs联系人!信息值)
                End Select
                rs联系人.MoveNext
            Wend
            SetLinkInfo str联系人姓名, str联系人关系, str联系人电话, str联系人身份证号, str附加信息
            rs联系人.MoveFirst
        Next
    End If
    '其他信息
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名 Not in ('血型','ABO','RH','医学警示','其他医学警示') And 信息名 Not like '联系人%'"
    Set rs其他信息 = zlDatabase.OpenSQLRecord(strSQL, "联系人其他信息", lng病人ID)
    '问题号:115886,焦博,2017/11/08,挂号提取该病人信息时，程序报错
    While rs其他信息.EOF = False
        If Nvl(rs其他信息!信息名) <> "" Then
            SetOtherInfo Nvl(rs其他信息!信息名), Nvl(rs其他信息!信息值)
        End If
        rs其他信息.MoveNext
    Wend
    '医疗卡属性
    Set mdic医疗卡属性 = Nothing
    
    Exit Sub
ErrHandl:
     If ErrCenter() = 1 Then Resume
End Sub

Private Function bln发卡(Optional ByVal blnCardNo As Boolean = False) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:判断当前是否为卡发操作 (不是发卡操作既是绑定卡操作)
'入参:
'编制:56599
'日期:2012-12-12 14:55:36
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln是否发卡 As Boolean
    If gCurSendCard.bln严格控制 = True Then
        mlng磁卡领用ID = CheckUsedBill(5, IIf(mlng磁卡领用ID > 0, mlng磁卡领用ID, gCurSendCard.lng共用批次), IIf(blnCardNo, mstrCard, UCase(txtPatient.Text)), gCurSendCard.lng卡类别ID)
        bln是否发卡 = IIf(mlng磁卡领用ID <= 0, False, True)
        If gCurSendCard.bln自制卡 = False Then
            bln是否发卡 = (gCurSendCard.bln是否发卡 = True)
        End If
    Else
        bln是否发卡 = mbln发卡
        If gCurSendCard.bln自制卡 = False Then
            bln是否发卡 = (gCurSendCard.bln是否发卡 = True)
        End If
    End If
    bln发卡 = bln是否发卡
    mbln发卡 = bln是否发卡
End Function

Public Sub Clear健康档案()
    '---------------------------------------------------------------------------------------------------------------------------------------------
'功能:清除界面信息
'入参:
'编制:56599
'日期:2012-12-25 14:55:36
'---------------------------------------------------------------------------------------------------------------------------------------------
    '68214:刘尔旋,2013-12-02,再次挂号时,血型值初始化
    cboBloodType.ListIndex = mintDefaultBlood
    'RH
    If cboBH.ListCount > 0 Then cboBH.ListIndex = -1
    '医学警示
    txtMedicalWarning.Text = ""
    '其他医学警示
    txtOtherWaring.Text = ""
    '联系人信息
    With vsLinkMan
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
        .TextMatrix(1, 4) = ""
    End With
    '接种情况
    With vsInoculate
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '其他信息
    With vsOtherInfo
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    
    '病人证件
    With vsCertificate
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
End Sub

Private Sub VsInoculate_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '问题号:56599
    If Col = 1 Or Col = 3 Then '接种名称列编辑时需判断是否字数超过了100
        With vsInoculate
           If Len(.TextMatrix(Row, Col)) > 100 Then
                MsgBox "接种名称输入字符超出最大字符数100,多出的字符将被自动截除！", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = Mid(.TextMatrix(Row, Col), 1, 100)
           End If
        End With
        If Col = 3 And vsInoculate.Rows - 1 = Row And vsInoculate.TextMatrix(Row, Col) <> "" Then
                vsInoculate.Rows = vsInoculate.Rows + 1
        End If
    Else
        With vsInoculate
           If IsDate(.TextMatrix(Row, Col)) = False And .TextMatrix(Row, Col) <> "    -  -  " Then
                MsgBox "输入的日期格式不对或不是正确的日期！", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = ""
           ElseIf .TextMatrix(Row, Col) = "    -  -  " Then
                .TextMatrix(Row, Col) = ""
           End If
        End With
    End If
End Sub

Private Sub vsInoculate_KeyDown(KeyCode As Integer, Shift As Integer)
    '问题号:56599
    If KeyCode = 27 And vsInoculate.Rows = 2 Then
        If vsInoculate.TextMatrix(1, 2) <> "    -  -  " And vsInoculate.TextMatrix(1, 3) <> "" Then
            vsInoculate.TextMatrix(1, 2) = "": vsInoculate.TextMatrix(1, 3) = ""
        Else
            vsInoculate.TextMatrix(1, 0) = "": vsInoculate.TextMatrix(1, 1) = ""
        End If
    End If
    If KeyCode = 27 And vsInoculate.Rows > 2 Then 'Esc
        If vsInoculate.TextMatrix(vsInoculate.Rows - 1, 2) <> "    -  -  " And vsInoculate.TextMatrix(vsInoculate.Rows - 1, 2) <> "" Or vsInoculate.TextMatrix(vsInoculate.Rows - 1, 3) <> "" Then
            vsInoculate.TextMatrix(vsInoculate.Rows - 1, 2) = "": vsInoculate.TextMatrix(vsInoculate.Rows - 1, 3) = ""
        Else
            vsInoculate.Rows = vsInoculate.Rows - 1
        End If
    End If
End Sub

Private Sub vsInoculate_KeyPress(KeyAscii As Integer)
    '78408:李南春,2014/10/9,光标跳转
    If KeyAscii = 13 Then
        If vsInoculate.Col = 3 And vsInoculate.Rows - 1 = vsInoculate.Row Then
            zlCommFun.PressKey vbKeyTab
        ElseIf vsInoculate.Col = 3 Then
            vsInoculate.Col = 0: vsInoculate.Row = vsInoculate.Row + 1
            zlCommFun.PressKey vbKeyReturn
        Else
            zlCommFun.PressKey vbKeyRight
        End If
    End If
End Sub

Private Function BlandCancel(ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal lngPatientID As Long) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:取消绑定卡
'入参:intType:0-当前卡号;1-当前类别;2-当前病人所有
'返回:取消成功,返回true,否则返回False
'编制:刘兴洪
'日期:2011-07-29 11:18:05
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Curdate As Date
    Dim strSQL As String, strPassWord As String

    On Error GoTo errHandle

    Curdate = zlDatabase.Currentdate
    
    'Zl_医疗卡变动_Insert
    strSQL = "Zl_医疗卡变动_Insert("
    '      变动类型_In   Number,
    '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
    strSQL = strSQL & "" & 14 & ","
    '      病人id_In     住院费用记录.病人id%Type,
    strSQL = strSQL & "" & lngPatientID & ","
    '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
    strSQL = strSQL & "" & lngCardTypeID & ","
    '      原卡号_In     病人医疗卡信息.卡号%Type,
    strSQL = strSQL & "NULL,"
    '      医疗卡号_In   病人医疗卡信息.卡号%Type,
    strSQL = strSQL & "'" & strCardNo & "'" & ","
    '      变动原因_In   病人医疗卡变动.变动原因%Type,
    strSQL = strSQL & "'挂号绑定卡自动取消绑定',"
    '      密码_In       病人信息.卡验证码%Type,
    strSQL = strSQL & "NULL,"
    '      操作员姓名_In 住院费用记录.操作员姓名%Type,
    strSQL = strSQL & "NULL,"
    '      变动时间_In   住院费用记录.登记时间%Type,
    strSQL = strSQL & "to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
    strSQL = strSQL & "NULL)"

     
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    BlandCancel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdBirthLocation_Click()
    Call SearchAddress("", txtBirthLocation)
End Sub

Private Sub cmdRegLocation_Click()
    Call SearchAddress("", txtRegLocation)
End Sub

Private Sub vsLinkMan_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsLinkMan
        If NewCol = .ColIndex("附加信息") Then
            If .TextMatrix(NewRow, .ColIndex("关系")) = "其他" Then
                .Editable = flexEDKbd
            Else
                .Editable = flexEDNone
            End If
        Else
            .Editable = flexEDKbd
        End If
    End With
End Sub

Private Sub vsLinkMan_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsLinkMan
        If KeyCode = 13 And .ColSel = .Cols - 1 Then
            .Rows = .Rows + 1
            .Select .Rows - 1, 0
            KeyCode = 0
        End If
        If KeyCode = 13 Then
            .Select .RowSel, .ColSel + 1
            KeyCode = 0
        End If
    End With
End Sub

Private Sub vsLinkMan_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Integer
    
    With vsLinkMan
        If Not Row = .FixedRows Then Exit Sub
        Select Case Col
            Case .ColIndex("姓名")
                txt联系人姓名.Text = Trim(.EditText)
            Case .ColIndex("关系")
                For i = 0 To cbo联系人关系.ListCount - 1
                    If NeedName(cbo联系人关系.List(i)) = Trim(.EditText) Then Exit For
                Next
                If i < cbo联系人关系.ListCount Then
                    cbo联系人关系.ListIndex = i
                Else
                    cbo联系人关系.ListIndex = -1
                End If
                txt其他关系.Visible = IIf(cbo联系人关系.ListIndex = 8, True, False)
                If cbo联系人关系.ListIndex = 8 Then
                    txt其他关系.Visible = True
                    cbo联系人关系.Width = 1225
                Else
                    txt其他关系.Visible = False: txt其他关系.Text = ""
                    .TextMatrix(Row, .ColIndex("附加信息")) = ""
                    cbo联系人关系.Width = 2425
                End If
            Case .ColIndex("身份证号")
                txt联系人身份证.Text = Trim(.EditText)
            Case .ColIndex("电话")
                txt联系人电话.Text = Trim(.EditText)
            Case .ColIndex("附加信息")
                txt其他关系.Text = Trim(.EditText)
        End Select
    End With
End Sub

Private Sub vsOtherInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsOtherInfo
        If KeyCode = 13 And .ColSel = .Cols - 1 Then
            .Rows = .Rows + 1
            .Select .Rows - 1, 0
            KeyCode = 0
        End If
        If KeyCode = 13 Then
            .Select .RowSel, .ColSel + 1
            KeyCode = 0
        End If
    End With
End Sub

Private Function InitTaskPanelOther() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载附加信息页面
    '返回:
    '问题号:73935
    '编制:冉俊明
    '日期:2014-07-3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup, Item As TaskPanelGroupItem
    Dim lngHwnd As Long
    
    Err = 0: On Error GoTo Errhand
    If CreatePlugInOK(mlngModul) Then
        If mlngPlugInHwnd <> 0 Then
            With wndTaskPanelOther
                Call .SetGroupInnerMargins(0, 0, 0, 0)
                Call .SetGroupOuterMargins(-1, -24, -1, -1)
                
                Set tkpGroup = .Groups.Add(1, "附加信息")
                tkpGroup.CaptionVisible = False
                tkpGroup.Expandable = False
                tkpGroup.Expanded = True
                
                Set Item = tkpGroup.Items.Add(1, "", xtpTaskItemTypeControl)
                Call HideFormCaption(mlngPlugInHwnd, False)
                Item.Handle = mlngPlugInHwnd
                
                .HotTrackStyle = xtpTaskPanelHighlightItem
                .Reposition
                .DrawFocusRect = True
            End With
        End If
    End If

    InitTaskPanelOther = True
    
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub DeletePatPicture(lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除病人照片
    '入参:lng病人ID - 病人ID
    '编制:56599
    '日期:2012-12-14 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo Errhand:
    strSQL = strSQL & "Zl_病人照片_Delete("
    strSQL = strSQL & lng病人ID & ")"
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SavePatPicture(lng病人ID As Long, strFile As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存病人照片
    '入参:lng病人ID - 病人ID
    '编制:56599
    '日期:2012-12-13 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
        
    If strFile = "" Then Exit Sub
    If Sys.SaveLob(glngSys, 27, lng病人ID, strFile, 0) = False Then
        ShowMsgbox "保存照片有误,请确认文件是否被删除!"
        Exit Sub
    End If
End Sub

Private Function ReadPatPricture(lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人照片
    '入参:lng病人ID - 病人ID
    '编制:56599
    '日期:2012-12-13 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    
    '67776:刘尔旋,2013-11-20,提取无照片的病人信息，照片没有清除
    Set imgPatient.Picture = Nothing
    
    strTmp = Sys.ReadLob(glngSys, 27, lng病人ID)
    mstr采集图片 = strTmp
    imgPatient.Picture = LoadPicture(strTmp)
    If strTmp <> "" Then Kill strTmp
End Function

Private Sub LoadIDImage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载身份证图像
    '编制:刘兴洪
    '日期:2014-06-30 16:20:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim objStdPic As StdPicture
    
    If mobjIDCard Is Nothing Then Exit Sub
    Call mobjIDCard.GetPhotoAsStdPicture(objStdPic)
    imgPatient.Picture = objStdPic
    mlng图像操作 = 4
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function SavePatiPic(ByVal lng病人ID As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------------------------
    '功能：保存病人照片
    '参数：
    '   lng病人ID - 病人ID
    '编制：冉俊明
    '时间：2014-7-8
    '--------------------------------------------------------------------------------------------------------------------------------------------
    Select Case mlng图像操作
        Case 1 '文件
            SavePatPicture lng病人ID, cmdialog.FileName
        Case 2 '采集
            SavePatPicture lng病人ID, mstr采集图片
            mstr采集图片 = ""
        Case 4 '二代身份证
            If imgPatient.Picture <> 0 Then
                mstrIDImageFile = App.Path & "\SFZIMG.bmp"
                SavePicture imgPatient.Picture, mstrIDImageFile
                SavePatPicture lng病人ID, mstrIDImageFile
            End If
        Case 3 '消除
            DeletePatPicture lng病人ID
    End Select
    
    mlng图像操作 = 0: mstr采集图片 = ""
End Function

Public Sub HideFormCaption(ByVal lngHwnd As Long, Optional ByVal blnBorder As Boolean = True)
'功能：隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(lngHwnd, vRect)
    lngStyle = GetWindowLong(lngHwnd, GWL_STYLE)

    If blnBorder Then
        lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
    Else
        lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
    End If
    SetWindowLong lngHwnd, GWL_STYLE, lngStyle
    SetWindowPos lngHwnd, 0, 0, 0, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Private Function CreatePublicPatient() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建zlPublicPatient部件
    '返回:创建成功,返回True,否则返回False
    '编制:冉俊明
    '日期:2014-07-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubPatient Is Nothing Then
        On Error Resume Next
        Set mobjPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If mobjPubPatient Is Nothing Then
        MsgBox "病人信息公共部件（zlPublicPatient）创建失败！", vbInformation, gstrSysName
        Exit Function
    Else
        If mobjPubPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser) = False Then
            MsgBox "病人信息公共部件（zlPublicPatient）初始化失败！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreatePublicPatient = True
End Function

Private Function SetBrushCardObject(ByVal blnComm As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置刷卡接口
    '返回: true-成功，false-失败
    '编制:李南春
    '日期:2016/6/20 13:54:56
    '问题:97634
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    
    Err = 0: On Error Resume Next
    SetBrushCardObject = True
    If txt卡号.Locked Then Exit Function
    If gCurSendCard.lng卡类别ID = 0 Or Val(gCurSendCard.str读卡性质) < 99 Then Exit Function
    If gobjSquare.objSquareCard.zlSetBrushCardObject(gCurSendCard.lng卡类别ID, IIf(blnComm, txt卡号, Nothing), strExpend) Then
        If mobjCommEvents Is Nothing Then Set mobjCommEvents = New clsCommEvents
        Call gobjSquare.objSquareCard.zlInitEvents(Me.Hwnd, mobjCommEvents)
    End If
End Function

Private Sub zlQueryEMPIPatiInfo()
    '功能：从EMPI平台获取病人信息
    '日期：2016/10/9 10:47:13
    '编制：李南春
    '说明：101170
    Dim rsTmp As ADODB.Recordset, strDiff As String, strMsgInfo As String
    Dim rsPatiInfo As ADODB.Recordset
    If CreatePlugInOK(mlngModul) = False Then Exit Sub
    If Trim(txtPatient.Text) = "" Then Exit Sub
    On Error GoTo Errhand
    If zlInitMEPIPati(rsTmp) = False Then Exit Sub
    
    With rsTmp
        .AddNew
        !病人ID = mlng病人ID
        !门诊号 = txt门诊号.Text
        !医保号 = txtPatiMCNO(0).Text
        !身份证号 = txt身份证号.Text
        !姓名 = txtPatient.Text
        !性别 = zlStr.NeedName(cbo性别.Text)
        If IsDate(txt出生日期.Text) Then
            !出生日期 = Format(txt出生日期.Text & " " & IIf(IsDate(txt出生时间.Text), txt出生时间.Text, "00:00"), "YYYY-MM-DD HH:MM")
        Else
            !出生日期 = ""
        End If
        !出生地点 = txtBirthLocation.Text
        !国籍 = zlStr.NeedName(cbo国籍.Text)
        !民族 = zlStr.NeedName(cbo民族.Text)
        !职业 = zlStr.NeedName(cbo职业.Text)
        !工作单位 = txt单位名称.Text
        !婚姻状况 = zlStr.NeedName(cbo婚姻.Text)
        !家庭电话 = txt家庭电话.Text
        !联系人电话 = txt联系人电话.Text
        !单位电话 = txt单位电话.Text
        !家庭地址 = cbo家庭地址.Text
        !家庭地址邮编 = txt家庭邮编.Text
        !户口地址 = txtRegLocation.Text
        !户口地址邮编 = txt户口地址邮编.Text
        !单位邮编 = txt单位邮编.Text
        !联系人姓名 = txt联系人姓名.Text
        !联系人关系 = zlStr.NeedName(cbo联系人关系.Text)
        .Update
    End With
    'EMPI没有找到病人信息,直接返回
    Dim rsOut As New ADODB.Recordset
    On Error Resume Next
    If gobjPlugIn.EMPI_QueryPatiInfo(glngSys, mlngModul, rsTmp, rsOut) = False Then
        Call zlPlugInErrH(Err, "EMPI_QueryPatiInfo")
        Err.Clear: Set mrsEMPIOut = Nothing: Exit Sub
    End If
    Err.Clear: On Error GoTo Errhand
    Set mrsEMPIOut = rsOut
    If mrsEMPIOut Is Nothing Then Exit Sub
    If mrsEMPIOut.RecordCount = 0 Then Exit Sub
    mrsEMPIOut.MoveFirst
    On Error Resume Next
    With mrsEMPIOut
        '104905:李南春，2017/1/12，根据接口返回的病人ID重新加载病人信息
        If mlng病人ID <> Val(Nvl(!病人ID)) And Val(Nvl(!病人ID)) <> 0 Then
            If mlng病人ID = 0 Then
                Set rsPatiInfo = GetPatiByID("病人ID", CStr(Nvl(!病人ID)))
                If rsPatiInfo.EOF Then
                    mlng病人ID = 0
                Else
                    zl加载病人信息 rsPatiInfo
                    
                    mbln基本信息调整 = Not (mlng病人ID <> 0 And InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";基本信息调整;") = 0)
                    txtPatient.Enabled = mbln基本信息调整: txt出生日期.Enabled = mbln基本信息调整: txt出生时间.Enabled = mbln基本信息调整
                    txt年龄.Enabled = mbln基本信息调整: cbo年龄单位.Enabled = mbln基本信息调整: cbo性别.Enabled = mbln基本信息调整
                    txt身份证号.Enabled = mbln基本信息调整
                    SetCtrVisibleAndMove
                End If
            Else
                MsgBox "EMPI返回的病人信息与His病人信息不一致，请在挂号界面重新查询确认。", vbInformation, gstrSysName
                Call cmdCancel_Click
                Exit Sub
            End If
        End If
        
        mstrPlugChange = ""
        If Nvl(!医保号) <> "" Then
            txtPatiMCNO(0).Text = Nvl(!医保号)
            txtPatiMCNO(1).Text = txtPatiMCNO(0).Text
        End If
        If mbln基本信息调整 Or mlng病人ID = 0 Then
            If Nvl(!身份证号) <> "" Then txt身份证号.Text = Nvl(!身份证号)
            If Nvl(!姓名) <> "" Then txtPatient.Text = Nvl(!姓名)
            If Nvl(!性别) <> "" Then cbo性别.ListIndex = cbo.FindIndex(cbo性别, Nvl(!性别), True)
            If Nvl(!出生日期) <> Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS") Then
                txt出生日期.Text = Format(Nvl(!出生日期), "YYYY-MM-DD")
                txt出生时间.Text = Format(Nvl(!出生日期), "HH:MM")
            End If
        Else
            If Nvl(!姓名) <> "" And txtPatient.Text <> Nvl(!姓名) Then strDiff = ",姓名"
            If Nvl(!性别) <> "" And cbo性别.ListIndex <> cbo.FindIndex(cbo性别, Nvl(!性别), True) Then strDiff = strDiff & ",性别"
            If Nvl(!出生日期) <> "" And Format(Nvl(!出生日期), "YYYY-MM-DD HH:MM:SS") <> Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS") Then strDiff = strDiff & ",出生日期"
            If Nvl(!身份证号) <> "" And txt身份证号.Text <> Nvl(!身份证号) Then strDiff = strDiff & ",身份证号"
        End If
        If txt门诊号.Enabled And Exist门诊号(Nvl(!门诊号), mlng病人ID) = False Then
            If Nvl(!门诊号) <> "" Then txt门诊号.Text = Nvl(!门诊号)
        Else
            If Nvl(!门诊号) <> "" And txt门诊号.Text <> Nvl(!门诊号) Then strDiff = strDiff & ",门诊号"
        End If
        If Nvl(!出生地点) <> "" Then txtBirthLocation.Text = Nvl(!出生地点)
        If Nvl(!国籍) <> "" Then cbo国籍.ListIndex = cbo.FindIndex(cbo国籍, Nvl(!国籍), True)
        If Nvl(!民族) <> "" Then cbo民族.ListIndex = cbo.FindIndex(cbo民族, Nvl(!民族), True)
        If Nvl(!职业) <> "" Then cbo职业.ListIndex = cbo.FindIndex(cbo职业, Nvl(!职业))
        If Nvl(!工作单位) <> "" Then txt单位名称.Text = Nvl(!工作单位)
        If Nvl(!婚姻状况) <> "" Then cbo婚姻.ListIndex = cbo.FindIndex(cbo婚姻, Nvl(!婚姻状况), True)
        If Nvl(!家庭电话) <> "" Then txt家庭电话.Text = Nvl(!家庭电话)
        If Nvl(!联系人电话) <> "" Then txt联系人电话.Text = Nvl(!联系人电话)
        If Nvl(!单位电话) <> "" Then txt单位电话.Text = Nvl(!单位电话)
        If Nvl(!家庭地址) <> "" Then cbo家庭地址.Text = Nvl(!家庭地址): padd家庭地址.Value = Nvl(!家庭地址)
        If Nvl(!家庭地址邮编) <> "" Then txt家庭邮编.Text = Nvl(!家庭地址邮编)
        If Nvl(!户口地址) <> "" Then txtRegLocation.Text = Nvl(!户口地址): padd户口地址.Value = Nvl(!户口地址)
        If Nvl(!户口地址邮编) <> "" Then txt户口地址邮编.Text = Nvl(!户口地址邮编)
        If Nvl(!单位邮编) <> "" Then txt单位邮编.Text = Nvl(!单位邮编)
        If Nvl(!联系人姓名) <> "" Then txt联系人姓名.Text = Nvl(!联系人姓名)
        If Nvl(!联系人关系) <> "" Then cbo联系人关系.ListIndex = cbo.FindIndex(cbo联系人关系, Nvl(!联系人关系), True)
    End With
    Err = 0: On Error GoTo 0
    '建档病人才进行提醒
    If mlng病人ID <> 0 Then
        If strDiff <> "" Then strDiff = Mid(strDiff, 2)
        If mstrPlugChange <> "" Then mstrPlugChange = Mid(mstrPlugChange, 2)
        If strDiff <> "" Then
            strMsgInfo = "病人的 " & strDiff & " 与EMPI信息不一致，因不具有调整基本信息的权限或与其他病人信息冲突，本次不会进行更新。"
        End If
        If strMsgInfo <> "" Then MsgBox strMsgInfo, vbInformation, gstrSysName
        mstrPlugChange = ""
    End If
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function zlSaveEMPIPatiInfo(ByVal blnNewPati As Boolean, ByVal lngPatiID As Long, ByVal lngClinicID As Long, ByRef strErrMsg As String) As Boolean
    '功能:上传病人信息到EMPI平台,如果平台信息保存失败，连同HIS数据一起回退
    '参数: In-lngPatiID 病人ID,lngClinicID 挂号ID
    '      Out-strErrMsg 错误信息，函数返回失败有效
    '返回:True-EMPI平台保存信息成功,False-保存失败
    '编制:李南春
    '说明:101170
    Dim blnCharge As Boolean, lngRet As Long
    If CreatePlugInOK(mlngModul) = False Then zlSaveEMPIPatiInfo = True: Exit Function
    
    On Error GoTo Errhand
    If mrsEMPIOut Is Nothing Then
        'EMPI没有病人信息，需要新建
        On Error Resume Next
        lngRet = gobjPlugIn.EMPI_AddPatiInfo(glngSys, mlngModul, lngPatiID, 0, lngClinicID, strErrMsg)
        Call zlPlugInErrH(Err, "EMPI_AddPatiInfo")
        If lngRet = 0 And Err.Number <> 438 Then Err.Clear: Exit Function
        Err.Clear: On Error GoTo Errhand
    Else
        '判断平台回传的信息是否发生改变
        With mrsEMPIOut
            If txt门诊号.Enabled And Exist门诊号(Nvl(!门诊号), lngPatiID) = False Then
                If txt门诊号.Text <> Nvl(!门诊号) Then blnCharge = True: GoTo EMPIModify
            End If
            If txtPatiMCNO(0).Text <> Nvl(!医保号) Then blnCharge = True: GoTo EMPIModify
            If mbln基本信息调整 Or blnNewPati Then
                If txt身份证号.Text <> Nvl(!身份证号) Then blnCharge = True: GoTo EMPIModify
                If txtPatient.Text <> Nvl(!姓名) Then blnCharge = True: GoTo EMPIModify
                If cbo性别.ListIndex <> cbo.FindIndex(cbo性别, Nvl(!性别), True) Then blnCharge = True: GoTo EMPIModify
                If Format(txt出生日期.Text, "YYYY-MM-DD") <> Format(Nvl(!出生日期), "YYYY-MM-DD") Then blnCharge = True: GoTo EMPIModify
                If Format(txt出生时间.Text, "HH:MM") <> Format(Nvl(!出生日期), "HH:MM") Then blnCharge = True: GoTo EMPIModify
            End If
            If txtBirthLocation.Text <> Nvl(!出生地点) Then blnCharge = True: GoTo EMPIModify
            If cbo国籍.ListIndex <> cbo.FindIndex(cbo国籍, Nvl(!国籍), True) Then blnCharge = True: GoTo EMPIModify
            If cbo民族.ListIndex <> cbo.FindIndex(cbo民族, Nvl(!民族), True) Then blnCharge = True: GoTo EMPIModify
            If cbo职业.ListIndex <> cbo.FindIndex(cbo职业, Nvl(!职业)) Then blnCharge = True: GoTo EMPIModify
            If txt单位名称.Text <> Nvl(!工作单位) Then blnCharge = True: GoTo EMPIModify
            If cbo婚姻.ListIndex <> cbo.FindIndex(cbo婚姻, Nvl(!婚姻状况), True) Then blnCharge = True: GoTo EMPIModify
            If txt家庭电话.Text <> Nvl(!家庭电话) Then blnCharge = True: GoTo EMPIModify
            If txt联系人电话.Text <> Nvl(!联系人电话) Then blnCharge = True: GoTo EMPIModify
            If txt单位电话.Text <> Nvl(!单位电话) Then blnCharge = True: GoTo EMPIModify
            If cbo家庭地址.Text <> Nvl(!家庭地址) Then blnCharge = True: GoTo EMPIModify
            If txt家庭邮编.Text <> Nvl(!家庭地址邮编) Then blnCharge = True: GoTo EMPIModify
            If txtRegLocation.Text <> Nvl(!户口地址) Then blnCharge = True: GoTo EMPIModify
            If txt户口地址邮编.Text <> Nvl(!户口地址邮编) Then blnCharge = True: GoTo EMPIModify
            If txt单位邮编.Text <> Nvl(!单位邮编) Then blnCharge = True: GoTo EMPIModify
            If txt联系人姓名.Text <> Nvl(!联系人姓名) Then blnCharge = True: GoTo EMPIModify
            If cbo联系人关系.ListIndex <> cbo.FindIndex(cbo联系人关系, Nvl(!联系人关系), True) Then blnCharge = True: GoTo EMPIModify
        End With
    End If
EMPIModify:
    If blnCharge Then
        On Error Resume Next
        lngRet = gobjPlugIn.EMPI_ModifyPatiInfo(glngSys, mlngModul, lngPatiID, 0, lngClinicID, strErrMsg)
        Call zlPlugInErrH(Err, "EMPI_AddPatiInfo")
        If lngRet = 0 And Err.Number <> 438 Then Err.Clear: Exit Function
        Err.Clear: On Error GoTo Errhand
    End If
    zlSaveEMPIPatiInfo = True
    Exit Function
Errhand:
    strErrMsg = Err.Description
    Call zlPlugInErrH(Err, "zlSaveEMPIPatiInfo")
    Call SaveErrLog
End Function

Private Sub vsCertificate_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long, lngCol As Long
    If Row < 1 Or Col < 0 Then Exit Sub
    '问题号:90875

    With vsCertificate
        If Col = 1 Or Col = 3 Then '证件号码不能超过30
            If Len(.TextMatrix(Row, Col)) > 30 Then
                 MsgBox "证件输入字符超出最大字符数30,多出的字符将被自动截除！", vbInformation, gstrSysName
                 .TextMatrix(Row, Col) = Mid(.TextMatrix(Row, Col), 1, 30)
            End If
            If Col = 3 And .Rows - 1 = Row And .TextMatrix(Row, Col) <> "" Then
                .Rows = .Rows + 1
            End If
        ElseIf Col = 0 Or Col = 2 Then '检查是否选择了重复的证件类型
            For lngRow = 1 To .Rows - 1
                For lngCol = 0 To .Cols - 1 Step 2
                    If (lngRow <> Row Or lngCol <> Col) And .TextMatrix(lngRow, lngCol) = .TextMatrix(Row, Col) And .TextMatrix(Row, Col) <> "" Then
                        MsgBox .TextMatrix(lngRow, lngCol) & "已存在，不能重复选择。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        .Select Row, Col
                        Exit Sub
                    End If
                Next
            Next
        End If
    End With
End Sub
Private Sub vsCertificate_KeyDown(KeyCode As Integer, Shift As Integer)
    '问题号:90875
    If KeyCode = 27 And vsCertificate.Rows = 2 Then
        If vsCertificate.TextMatrix(1, 3) <> "" Then
            vsCertificate.TextMatrix(1, 2) = "": vsCertificate.TextMatrix(1, 3) = ""
        Else
            vsCertificate.TextMatrix(1, 0) = "": vsCertificate.TextMatrix(1, 1) = ""
        End If
    End If
    If KeyCode = 27 And vsCertificate.Rows > 2 Then 'Esc
        If vsCertificate.TextMatrix(vsCertificate.Rows - 1, 2) <> "" Or vsCertificate.TextMatrix(vsCertificate.Rows - 1, 3) <> "" Then
            vsCertificate.TextMatrix(vsCertificate.Rows - 1, 2) = "": vsCertificate.TextMatrix(vsCertificate.Rows - 1, 3) = ""
        Else
            vsCertificate.Rows = vsCertificate.Rows - 1
        End If
    End If
End Sub

Private Sub vsCertificate_KeyPress(KeyAscii As Integer)
    '78408:李南春,2014/10/9,光标跳转
    If KeyAscii = 13 Then
        If vsCertificate.Col = 3 And vsCertificate.Rows - 1 = vsCertificate.Row Then
            zlCommFun.PressKey vbKeyTab
        ElseIf vsCertificate.Col = 3 Then
            vsCertificate.Col = 0: vsCertificate.Row = vsCertificate.Row + 1
            zlCommFun.PressKey vbKeyReturn
        Else
            zlCommFun.PressKey vbKeyRight
        End If
    End If
End Sub

Private Sub InitCertificate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:90875
    '日期:2015/12/17 16:59:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Errhand
    Dim strSQL As String, rsTemp As ADODB.Recordset, str关系 As String, i As Integer
    With vsCertificate
    '初始化列表属性
    vsCertificate.Editable = flexEDKbdMouse
    '设置列头
    SetColumHeader vsCertificate, C_CertificateHeader
    '设置列信息
    strSQL = "Select 名称,缺省标志 from 证件类型  Where  名称 Not Like '其他%' and 名称 Not Like '%身份证'" & vbNewLine & _
            " And Not 名称 in (Select 名称 from  医疗卡类别 Where Nvl(是否证件,0)=0 or Nvl(是否启用,0)=0)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTemp.RecordCount = 0 Then .Editable = flexEDNone: Exit Sub
        Do While Not rsTemp.EOF
            str关系 = str关系 & "|" & Nvl(rsTemp!名称)
            rsTemp.MoveNext
        Loop
        str关系 = Mid(str关系, 2)
        If str关系 <> "" Then .ColComboList(0) = str关系: .ColComboList(2) = str关系
    End With
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub LoadCertificate(ByVal lng病人ID As Long)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人的证件信息到界面
    '编制:李南春
    '时间:2015/12/17 17:37:27
    '问题:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngRow As Integer, lngCol As Integer
    
    On Error GoTo Errhand
    strSQL = "Select  A.名称,A.ID,B.卡号 from 医疗卡类别 A, 病人医疗卡信息 B " & _
            "Where A.ID= B.卡类别ID And A.是否启用=1 And A.是否证件=1 And B.状态=0  And  B.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    If rsTemp.RecordCount = 0 Then Exit Sub
    With vsCertificate
        .Clear 1
        .Rows = 2
        lngRow = 1: lngCol = 0
        While Not rsTemp.EOF
            .TextMatrix(lngRow, lngCol) = Nvl(rsTemp!名称)
            .TextMatrix(lngRow, lngCol + 1) = Nvl(rsTemp!卡号)
            lngCol = lngCol + 2
            If lngCol > 2 Then .Rows = .Rows + 1: lngRow = lngRow + 1: lngCol = 0
            rsTemp.MoveNext
        Wend
    End With
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub AddCardTypeSQL(ByVal intOper As Integer, ByVal lng卡类别ID As Long, ByVal strCode As String, ByVal str全名 As String, ByVal str短名 As String, _
                           ByVal lng卡号长度 As Long, ByRef colPro As Collection)
    Dim strSQL As String

    ' Zl_医疗卡类别_Update
    strSQL = "Zl_医疗卡类别_Update("
    '  Id_In           In 医疗卡类别.ID%Type,
    strSQL = strSQL & "" & lng卡类别ID & ","
    '  编码_In         In 医疗卡类别.编码%Type,
    strSQL = strSQL & "'" & strCode & "',"
    '  名称_In         In 医疗卡类别.名称%Type,
    strSQL = strSQL & "'" & str全名 & "',"
    '  短名_In         In 医疗卡类别.短名%Type,
    strSQL = strSQL & "'" & str短名 & "',"
    '  前缀文本_In     In 医疗卡类别.前缀文本%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  卡号长度_In     In 医疗卡类别.卡号长度%Type,
    strSQL = strSQL & "" & lng卡号长度 & ","
    '  缺省标志_In     In 医疗卡类别.缺省标志%Type,
    strSQL = strSQL & "" & 0 & ","
    '  是否固定_In     In 医疗卡类别.是否固定%Type,
    strSQL = strSQL & "1,"
    '  是否严格控制_In In 医疗卡类别.是否严格控制%Type,
    strSQL = strSQL & "" & 0 & ","
    '  是否自制_In     In 医疗卡类别.是否自制%Type,
    strSQL = strSQL & "" & 0 & ","
    '  是否存在帐户_In In 医疗卡类别.是否存在帐户%Type,
    strSQL = strSQL & "" & 0 & ","
    '  是否全退_In     In 医疗卡类别.是否全退%Type,
    strSQL = strSQL & "0,"
    '  部件_In         In 医疗卡类别.部件%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  备注_In         In 医疗卡类别.备注%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  特定项目_In     In 医疗卡类别.特定项目%Type,
    strSQL = strSQL & "'" & strCode & "',"
    '    收费细目id_In   In 收费项目目录.ID%Type,
    strSQL = strSQL & "" & "0" & ","
    '  结算方式_In     In 医疗卡类别.结算方式%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  是否启用_In     In 医疗卡类别.是否启用%Type,
    strSQL = strSQL & "1,"
    '  卡号密文_In     In 医疗卡类别.卡号密文%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  是否重复使用_In In 医疗卡类别.是否重复使用%Type,
    strSQL = strSQL & "" & 1 & ","
    '密码长度_In     In 医疗卡类别.密码长度%Type,
    strSQL = strSQL & "" & 10 & ","
    '密码长度限制_In In 医疗卡类别.密码长度限制%Type,
    strSQL = strSQL & "" & 0 & ","
    '密码规则_In     In 医疗卡类别.密码规则%Type,
    strSQL = strSQL & "" & 0 & ","
    strSQL = strSQL & "" & 1 & ","
    '  操作方式_In     In Integer := 0
    strSQL = strSQL & "" & intOper & ","
    '是否模糊查找_In     In 医疗卡类别.是否模糊查找%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '问题号:51072
    '密码输入限制_In     In 医疗卡类别.密码输入限制%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '是否缺省密码_In     In 医疗卡类别.是否缺省密码%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '问题号:56508
    '是否制卡_In
    strSQL = strSQL & "" & 0 & ","
    '是否发卡_In
    strSQL = strSQL & "" & 0 & ","
    '是否写卡_In
    strSQL = strSQL & "" & 0 & ","
    '问题号:57697
    '险类_In
    strSQL = strSQL & "" & 0 & ","
    '问题号:57326
    strSQL = strSQL & "" & 1 & ","
    '77872,李南春,2014/12/3:是否支持转帐及代扣
    '是否转帐及代扣_In  In 医疗卡类别.是否转帐及代扣%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '读卡性质_In       In 医疗卡类别.读卡性质%Type := '1000',
    strSQL = strSQL & "" & "1000" & ","
    '键盘控制方式_In   In 医疗卡类别.键盘控制方式%Type := 0,
    strSQL = strSQL & "" & 0 & ","
    '90875:李南春,2015/12/16,增加医疗卡证件类型
    '是否证件_In  In 医疗卡类别.是否证件%Type:=0
    strSQL = strSQL & "" & 1 & ")"
    
    zlAddArray colPro, strSQL
End Sub

Public Sub AddCertificate(ByVal lng病人ID As Long, ByRef colPro As Collection, ByVal dtCurdate As Date)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:建立证件卡类信息，如果是第一次建立卡类别
    '编制:李南春
    '时间:2015/12/17 17:37:27
    '问题:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, rsPatiCard As ADODB.Recordset
    Dim lngRow As Integer, lngCol As Integer
    Dim lngID As Long, strCode As String
    
    On Error GoTo Errhand
    '绑定卡前要判断卡类别是否存在
    strSQL = "Select B.ID,B.编码,B.卡号长度,B.名称,A.卡号,A.病人ID,Decode(A.卡号 ,NULL,1,0) as 标识 from 病人医疗卡信息 A,医疗卡类别 B " & _
            "Where A.卡类别ID(+)=B.ID And B.是否证件=1 And A.状态(+)=0 And A.病人ID(+)=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    Set rsPatiCard = zlDatabase.CopyNewRec(rsTemp)
    With vsCertificate
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) <> "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    lngID = 0: strCode = ""
                    rsTemp.Filter = "名称='" & .TextMatrix(lngRow, lngCol) & "'"
                    If rsTemp.RecordCount = 0 Then
                        lngID = zlDatabase.GetNextId("医疗卡类别")
                        If mstrFirstCode = "" Then
                            strCode = zlDatabase.GetMax("医疗卡类别", "编码", 4)
                            mstrFirstCode = strCode
                        Else
                            strCode = CStr(Val(mstrFirstCode) + 1)
                            strCode = Format(strCode, String(4, "0"))
                            mstrFirstCode = strCode
                        End If
                        Call AddCardTypeSQL(0, lngID, strCode, .TextMatrix(lngRow, lngCol), Left(.TextMatrix(lngRow, lngCol), 1), Len(.TextMatrix(lngRow, lngCol + 1)), colPro)
                    ElseIf Len(.TextMatrix(lngRow, lngCol + 1)) > Val(Nvl(rsTemp!卡号长度)) Then
                        Call AddCardTypeSQL(1, Val(Nvl(rsTemp!ID)), Nvl(rsTemp!编码), .TextMatrix(lngRow, lngCol), Left(.TextMatrix(lngRow, lngCol), 1), Len(.TextMatrix(lngRow, lngCol + 1)), colPro)
                    End If
                    
                    '进行证件卡绑定
                    rsPatiCard.Filter = "名称='" & .TextMatrix(lngRow, lngCol) & "' And 卡号='" & .TextMatrix(lngRow, lngCol + 1) & "'"
                    If rsPatiCard.RecordCount = 0 Then
                        'Zl_医疗卡变动_Insert
                         strSQL = "Zl_医疗卡变动_Insert("
                        '      变动类型_In   Number,
                        '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
                        strSQL = strSQL & "" & 11 & ","
                        '      病人id_In     住院费用记录.病人id%Type,
                        strSQL = strSQL & "" & lng病人ID & ","
                        '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
                        strSQL = strSQL & "" & IIf(lngID = 0, rsTemp!ID, lngID) & ","
                        '      原卡号_In     病人医疗卡信息.卡号%Type,
                        strSQL = strSQL & "'" & "" & "',"
                        '      医疗卡号_In   病人医疗卡信息.卡号%Type,
                        strSQL = strSQL & "'" & .TextMatrix(lngRow, lngCol + 1) & "',"
                        '      变动原因_In   病人医疗卡变动.变动原因%Type,
                        '      --变动原因_In:如果密码调整，变动原因为密码.加密的
                        strSQL = strSQL & "'" & "证件卡绑定" & "',"
                        '      密码_In       病人信息.卡验证码%Type,
                        strSQL = strSQL & "'" & "" & "',"
                        '      操作员姓名_In 住院费用记录.操作员姓名%Type,
                        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                        '      变动时间_In   住院费用记录.登记时间%Type,
                        strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                        '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
                        strSQL = strSQL & "'" & "" & "',"
                        '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
                        strSQL = strSQL & "NULL)"
                    
                        zlAddArray colPro, strSQL
                    Else
                        rsPatiCard!标识 = 1
                        rsPatiCard.Update
                    End If
                End If
            Next
        Next
    End With
    '卡号列表中没有证件号，要解除绑定
    rsPatiCard.Filter = "标识=0"
    If rsPatiCard.RecordCount > 0 Then
        rsPatiCard.MoveFirst
        Do While Not rsPatiCard.EOF
            'Zl_医疗卡变动_Insert
             strSQL = "Zl_医疗卡变动_Insert("
            '      变动类型_In   Number,
            '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
            strSQL = strSQL & "" & 14 & ","
            '      病人id_In     住院费用记录.病人id%Type,
            strSQL = strSQL & "" & lng病人ID & ","
            '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
            strSQL = strSQL & "" & rsPatiCard!ID & ","
            '      原卡号_In     病人医疗卡信息.卡号%Type,
            strSQL = strSQL & "'" & "" & "',"
            '      医疗卡号_In   病人医疗卡信息.卡号%Type,
            strSQL = strSQL & "'" & rsPatiCard!卡号 & "',"
            '      变动原因_In   病人医疗卡变动.变动原因%Type,
            '      --变动原因_In:如果密码调整，变动原因为密码.加密的
            strSQL = strSQL & "'" & "证件卡取消绑定" & "',"
            '      密码_In       病人信息.卡验证码%Type,
            strSQL = strSQL & "'" & "" & "',"
            '      操作员姓名_In 住院费用记录.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '      变动时间_In   住院费用记录.登记时间%Type,
            strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
            strSQL = strSQL & "'" & "" & "',"
            '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
            strSQL = strSQL & "NULL)"
        
            zlAddArray colPro, strSQL
            rsPatiCard.MoveNext
        Loop
    End If
    rsPatiCard.Close
    Exit Sub
Errhand:
    rsPatiCard.Close
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function IsCertificateCard(ByVal lng病人ID As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:证件卡类检查
    '编制:李南春
    '时间:2015/12/17 17:37:27
    '问题:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Long, str证件类型 As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strCardName As String
    
    On Error GoTo Errhand
    With vsCertificate
        '检查输入是否完整
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) = "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    MsgBox "请选择卡号" & .TextMatrix(lngRow, lngCol + 1) & "的证件类型", vbInformation, gstrSysName
                    .Select lngRow, lngCol
                    Exit Function
                End If
                If .TextMatrix(lngRow, lngCol) <> "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    strSQL = "Select 1 from 病人医疗卡信息 A,医疗卡类别 B " & _
                            "Where A.卡类别ID=B.ID And B.名称=[1] And B.是否证件=1 And A.卡号=[2] And  A.病人ID<>[3]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .TextMatrix(lngRow, lngCol), Trim(.TextMatrix(lngRow, lngCol + 1)), lng病人ID)
                    If rsTmp.RecordCount > 0 Then
                        MsgBox .TextMatrix(lngRow, lngCol) & ":" & .TextMatrix(lngRow, lngCol + 1) & "正在被使用,请检查!", vbInformation, gstrSysName
                        .Select lngRow, lngCol
                        Exit Function
                    End If
                    str证件类型 = str证件类型 & ",'" & .TextMatrix(lngRow, lngCol) & "'"
                End If
            Next
        Next
        
        '检查证件类型是否与非证件的医疗卡类别重复，重复则不保存信息
        str证件类型 = Mid(str证件类型, 2)
        If str证件类型 = "" Then IsCertificateCard = True: Exit Function
        strSQL = "Select 名称 From 医疗卡类别 where 名称 in (" & str证件类型 & ") And Nvl(是否证件,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                strCardName = strCardName & "," & Nvl(rsTmp!名称)
            Loop
            
            strCardName = Mid(strCardName, 2)
            MsgBox "医疗卡类别【" & strCardName & "】名称重复,不能继续添加。", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    IsCertificateCard = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function IsMobileNO(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------
    '功能:判断传入的是否为手机号
    '参数:strinput-11位手机号
    '编制:刘尔旋
    '日期:2017-1-25
    '---------------------------------------------------------------------------------------------
    Dim strMobileRange As String
    If Not IsNumeric(strInput) Then Exit Function
    If Len(strInput) <> 11 Then Exit Function
    '中国移动
    strMobileRange = ",139,138,137,136,135,134,159,158,157,150,151,152,147,188,187,182,183,184,178"
    '中国联通
    strMobileRange = strMobileRange & ",130,131,132,156,155,186,185,145,176"
    '中国电信
    strMobileRange = strMobileRange & ",133,153,189,180,181,177,173"
    '虚拟运营商
    strMobileRange = strMobileRange & ",170,"
    If InStr(strMobileRange, "," & Mid(strInput, 1, 3) & ",") = 0 Then Exit Function
    IsMobileNO = True
End Function

Private Sub ReLoadCardFee()
    '离开检查卡费
    Dim lng病人ID As Long, lng收费细目ID As Long
    Dim strSQL As String, str年龄 As String
    Dim rsTmp As ADODB.Recordset
    
    gCurSendCard.lng收费细目ID = 0
    If gCurSendCard.rs卡费 Is Nothing Then Exit Sub
    If gCurSendCard.rs卡费.RecordCount = 0 Then Exit Sub
    If gCurSendCard.lng卡类别ID = 0 Then Exit Sub
    If Trim(txtPatient.Text) = "" Or Trim(txt卡号.Text) = "" Then Exit Sub
    
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    gCurSendCard.rs卡费.MoveFirst
    
    strSQL = "Select Zl1_Ex_CardFee([1],[2],[3],[4],[5],[6],[7],[8],[9]) as 收费细目ID From Dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "卡费", mlngModul, gCurSendCard.lng卡类别ID, Trim(txt卡号.Text), mlng病人ID, _
                Trim(txtPatient.Text), NeedName(cbo性别.Text), str年龄, txt身份证号.Text, Val(Nvl(gCurSendCard.rs卡费!收费细目ID)))
    If rsTmp.EOF Then Exit Sub
    
    lng收费细目ID = Val(Nvl(rsTmp!收费细目ID))
    Set rsTmp = zlGetSpecialItemFee("", mstrPriceGrade, lng收费细目ID)
    If Not rsTmp Is Nothing Then
        Set gCurSendCard.rs卡费 = rsTmp
        gCurSendCard.lng收费细目ID = lng收费细目ID
    End If
End Sub

