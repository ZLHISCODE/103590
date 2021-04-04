VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
   Begin VB.PictureBox picInfo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6885
      Left            =   135
      ScaleHeight     =   6885
      ScaleWidth      =   11490
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   795
      Width           =   11490
      Begin ZlPatiAddress.PatiAddress padd户口地址 
         Height          =   360
         Left            =   1170
         TabIndex        =   17
         Tag             =   "户口地址"
         Top             =   2504
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
      Begin ZlPatiAddress.PatiAddress padd家庭地址 
         Height          =   360
         Left            =   1170
         TabIndex        =   14
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
      Begin VB.TextBox txtMobile 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   8670
         MaxLength       =   20
         TabIndex        =   28
         Top             =   4110
         Width           =   2760
      End
      Begin VB.ComboBox cbo年龄单位 
         Height          =   360
         Left            =   7965
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   459
         Width           =   720
      End
      Begin VB.TextBox txt区域 
         Height          =   360
         Left            =   1125
         MaxLength       =   50
         TabIndex        =   27
         Top             =   4110
         Width           =   5955
      End
      Begin VB.TextBox txt过敏 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1995
         MaxLength       =   50
         TabIndex        =   74
         Top             =   7980
         Visible         =   0   'False
         Width           =   990
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
         TabIndex        =   73
         TabStop         =   0   'False
         ToolTipText     =   "热键:F3"
         Top             =   6540
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton cmd家庭地址 
         Caption         =   "…"
         Height          =   360
         Left            =   8070
         TabIndex        =   72
         ToolTipText     =   "热键F3"
         Top             =   2100
         Width           =   360
      End
      Begin VB.TextBox txt家庭邮编 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   10140
         MaxLength       =   6
         TabIndex        =   15
         Top             =   2095
         Width           =   1290
      End
      Begin VB.TextBox txt家庭电话 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5790
         MaxLength       =   20
         TabIndex        =   8
         Top             =   855
         Width           =   2880
      End
      Begin VB.TextBox txt身份证号 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1155
         MaxLength       =   18
         TabIndex        =   7
         Top             =   868
         Width           =   2880
      End
      Begin VB.ComboBox cbo职业 
         Height          =   360
         Left            =   8670
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2913
         Width           =   2775
      End
      Begin VB.ComboBox cbo婚姻 
         Height          =   360
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3322
         Width           =   2775
      End
      Begin VB.ComboBox cbo民族 
         Height          =   360
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2913
         Width           =   2775
      End
      Begin VB.ComboBox cbo国籍 
         Height          =   360
         Left            =   4710
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2913
         Width           =   2775
      End
      Begin VB.ComboBox cbo性别 
         Height          =   360
         IMEMode         =   3  'DISABLE
         ItemData        =   "frmPatiInfo.frx":0E42
         Left            =   1170
         List            =   "frmPatiInfo.frx":0E44
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   459
         Width           =   825
      End
      Begin VB.TextBox txt年龄 
         Height          =   360
         IMEMode         =   2  'OFF
         Left            =   7275
         TabIndex        =   5
         Top             =   459
         Width           =   690
      End
      Begin VB.TextBox txtPatient 
         Height          =   360
         Left            =   1170
         MaxLength       =   100
         TabIndex        =   0
         Top             =   50
         Width           =   2880
      End
      Begin VB.TextBox txt门诊号 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5790
         MaxLength       =   18
         TabIndex        =   1
         Top             =   50
         Width           =   2880
      End
      Begin VB.TextBox txtPatiMCNO 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1170
         MaxLength       =   30
         TabIndex        =   11
         Top             =   1686
         Width           =   2880
      End
      Begin VB.TextBox txtPatiMCNO 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   5790
         MaxLength       =   30
         TabIndex        =   12
         Top             =   1686
         Width           =   2880
      End
      Begin VB.ComboBox cbo费别 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4710
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3322
         Width           =   2775
      End
      Begin VB.ComboBox cbo付款方式 
         Height          =   360
         ItemData        =   "frmPatiInfo.frx":0E46
         Left            =   8670
         List            =   "frmPatiInfo.frx":0E48
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3322
         Width           =   2775
      End
      Begin VB.ComboBox cbo家庭地址 
         Height          =   360
         Left            =   1170
         TabIndex        =   13
         Top             =   2100
         Width           =   6900
      End
      Begin VB.CommandButton cmd区域 
         Caption         =   "…"
         Height          =   360
         Left            =   7080
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "热键：F3"
         Top             =   4110
         Width           =   375
      End
      Begin VB.TextBox txt支付密码 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1170
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1277
         Width           =   2880
      End
      Begin VB.TextBox txt验证密码 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5790
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1277
         Width           =   2880
      End
      Begin VB.TextBox txt过敏反应 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4230
         MaxLength       =   200
         TabIndex        =   68
         Top             =   6645
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.TextBox txt监护人 
         Height          =   360
         IMEMode         =   2  'OFF
         Left            =   8670
         MaxLength       =   20
         TabIndex        =   26
         Top             =   3720
         Width           =   2775
      End
      Begin VB.TextBox txtBirthLocation 
         Height          =   360
         Left            =   1125
         MaxLength       =   100
         TabIndex        =   25
         Top             =   3720
         Width           =   5955
      End
      Begin VB.CommandButton cmdBirthLocation 
         Caption         =   "…"
         Height          =   360
         Left            =   7080
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox txtRegLocation 
         Height          =   360
         Left            =   1170
         MaxLength       =   100
         TabIndex        =   16
         Top             =   2504
         Width           =   6900
      End
      Begin VB.CommandButton cmdRegLocation 
         Caption         =   "…"
         Height          =   360
         Left            =   8070
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2504
         Width           =   360
      End
      Begin VB.PictureBox picPatient 
         Height          =   1620
         Left            =   9090
         ScaleHeight     =   1560
         ScaleWidth      =   2025
         TabIndex        =   65
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
         TabIndex        =   64
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
         TabIndex        =   63
         Top             =   1665
         Width           =   585
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
         TabIndex        =   62
         Top             =   1665
         Width           =   600
      End
      Begin VB.TextBox txt户口地址邮编 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   10140
         MaxLength       =   6
         TabIndex        =   18
         Top             =   2504
         Width           =   1290
      End
      Begin VB.Frame fraContact 
         Caption         =   "联系人信息"
         Height          =   720
         Left            =   30
         TabIndex        =   56
         Top             =   4500
         Width           =   11415
         Begin VB.TextBox txt其他关系 
            Height          =   360
            Left            =   6705
            MaxLength       =   30
            TabIndex        =   57
            Top             =   285
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.ComboBox cbo联系人关系 
            Height          =   360
            Left            =   6540
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   270
            Width           =   1695
         End
         Begin VB.TextBox txt联系人姓名 
            Height          =   360
            Left            =   630
            MaxLength       =   64
            TabIndex        =   29
            Top             =   270
            Width           =   2160
         End
         Begin VB.TextBox txt联系人电话 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   3450
            MaxLength       =   18
            TabIndex        =   30
            Top             =   270
            Width           =   2460
         End
         Begin VB.TextBox txt联系人身份证 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   9120
            MaxLength       =   18
            TabIndex        =   32
            Top             =   270
            Width           =   2205
         End
         Begin VB.Label lbl联系人关系 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "关系"
            Height          =   240
            Left            =   6015
            TabIndex        =   61
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl联系人电话 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "电话"
            Height          =   240
            Left            =   2925
            TabIndex        =   60
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl联系人姓名 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名"
            Height          =   240
            Left            =   135
            TabIndex        =   59
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl联系人身份证 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "身份证"
            Height          =   240
            Left            =   8355
            TabIndex        =   58
            Top             =   330
            Width           =   720
         End
      End
      Begin VB.Frame fraUnit 
         Caption         =   "单位信息"
         Height          =   750
         Left            =   30
         TabIndex        =   51
         Top             =   5250
         Width           =   11415
         Begin VB.TextBox txt单位电话 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   9120
            MaxLength       =   20
            TabIndex        =   35
            Top             =   270
            Width           =   2205
         End
         Begin VB.TextBox txt单位邮编 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   6540
            MaxLength       =   6
            TabIndex        =   34
            Top             =   270
            Width           =   1680
         End
         Begin VB.TextBox txt单位名称 
            Height          =   360
            Left            =   660
            MaxLength       =   100
            TabIndex        =   33
            Top             =   270
            Width           =   4860
         End
         Begin VB.CommandButton cmd单位名称 
            Caption         =   "…"
            Height          =   360
            Left            =   5520
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   270
            Width           =   360
         End
         Begin VB.Label lbl单位电话 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "电话"
            Height          =   240
            Left            =   8580
            TabIndex        =   55
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl单位邮编 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "邮编"
            Height          =   240
            Left            =   6015
            TabIndex        =   54
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl单位名称 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "名称"
            Height          =   240
            Left            =   135
            TabIndex        =   53
            Top             =   330
            Width           =   480
         End
      End
      Begin XtremeSuiteControls.TaskPanel TaskPanel1 
         Height          =   30
         Left            =   1680
         TabIndex        =   69
         Top             =   375
         Width           =   30
         _Version        =   589884
         _ExtentX        =   53
         _ExtentY        =   53
         _StockProps     =   64
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   1515
         Left            =   2850
         TabIndex        =   71
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh过敏 
         Height          =   1215
         Left            =   30
         TabIndex        =   36
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
         TabIndex        =   4
         Top             =   465
         Width           =   840
         _ExtentX        =   1482
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
         TabIndex        =   3
         Top             =   465
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "YYYY-MM-DD"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblMobile 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手机号"
         Height          =   240
         Left            =   7920
         TabIndex        =   112
         Top             =   4170
         Width           =   720
      End
      Begin VB.Label lbl出生时间 
         AutoSize        =   -1  'True
         Caption         =   "时间"
         Height          =   240
         Left            =   4920
         TabIndex        =   111
         Top             =   525
         Width           =   480
      End
      Begin VB.Label lbl家庭邮编 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "现住址邮编"
         Height          =   240
         Left            =   8910
         TabIndex        =   98
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lbl家庭电话 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "电话"
         Height          =   240
         Left            =   5280
         TabIndex        =   97
         Top             =   915
         Width           =   480
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "现住址"
         Height          =   240
         Left            =   390
         TabIndex        =   96
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label lbl身份证 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   240
         Left            =   150
         TabIndex        =   95
         Top             =   930
         Width           =   960
      End
      Begin VB.Label lbl国籍 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍"
         Height          =   240
         Left            =   4170
         TabIndex        =   94
         Top             =   2970
         Width           =   480
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         Height          =   240
         Left            =   660
         TabIndex        =   93
         Top             =   2973
         Width           =   480
      End
      Begin VB.Label lbl职业 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   240
         Left            =   8160
         TabIndex        =   92
         Top             =   2970
         Width           =   480
      End
      Begin VB.Label lbl婚姻 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   240
         Left            =   150
         TabIndex        =   91
         Top             =   3382
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   240
         Left            =   6765
         TabIndex        =   90
         Top             =   525
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   240
         Left            =   630
         TabIndex        =   89
         Top             =   519
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   240
         Left            =   630
         TabIndex        =   88
         Top             =   110
         Width           =   480
      End
      Begin VB.Label lbl门诊号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   240
         Left            =   5040
         TabIndex        =   87
         Top             =   105
         Width           =   720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   -150
         X2              =   7695
         Y1              =   7785
         Y2              =   7785
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         Height          =   240
         Left            =   2430
         TabIndex        =   86
         Top             =   525
         Width           =   960
      End
      Begin VB.Label lblPatiMCNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         Height          =   240
         Index           =   0
         Left            =   390
         TabIndex        =   85
         Top             =   1746
         Width           =   720
      End
      Begin VB.Label lblPatiMCNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "验证医保号"
         Height          =   240
         Index           =   1
         Left            =   4560
         TabIndex        =   84
         Top             =   1740
         Width           =   1200
      End
      Begin VB.Label lbl付款方式 
         BackStyle       =   0  'Transparent
         Caption         =   "付款方式"
         Height          =   300
         Left            =   7680
         TabIndex        =   83
         Top             =   3352
         Width           =   960
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   240
         Left            =   4170
         TabIndex        =   82
         Top             =   3375
         Width           =   480
      End
      Begin VB.Label lbl区域 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "区域"
         Height          =   240
         Left            =   630
         TabIndex        =   81
         Top             =   4170
         Width           =   480
      End
      Begin VB.Label lbl支付密码 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "支付密码"
         Height          =   240
         Left            =   150
         TabIndex        =   80
         Top             =   1337
         Width           =   960
      End
      Begin VB.Label lbl验证密码 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "验证密码"
         Height          =   240
         Left            =   4800
         TabIndex        =   79
         Top             =   1335
         Width           =   960
      End
      Begin VB.Label lbl监护人 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  监护人"
         Height          =   240
         Left            =   7680
         TabIndex        =   78
         Top             =   3780
         Width           =   960
      End
      Begin VB.Label lblBirthLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点"
         Height          =   240
         Left            =   150
         TabIndex        =   77
         Top             =   3780
         Width           =   960
      End
      Begin VB.Label lblRegLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口地址"
         Height          =   240
         Left            =   150
         TabIndex        =   76
         Top             =   2564
         Width           =   960
      End
      Begin VB.Label lbl户口地址邮编 
         Alignment       =   1  'Right Justify
         Caption         =   "户口地址邮编"
         Height          =   240
         Left            =   8595
         TabIndex        =   75
         Top             =   2564
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   2010
      Top             =   7620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picTaskPanelOther 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   8190
      ScaleHeight     =   825
      ScaleWidth      =   1755
      TabIndex        =   48
      Top             =   7440
      Visible         =   0   'False
      Width           =   1755
      Begin XtremeSuiteControls.TaskPanel wndTaskPanelOther 
         Height          =   435
         Left            =   330
         TabIndex        =   49
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
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   420
      Left            =   90
      TabIndex        =   46
      Top             =   7485
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "返回(&X)"
      Height          =   420
      Left            =   6450
      TabIndex        =   44
      ToolTipText     =   "热键：F2"
      Top             =   7455
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   4875
      TabIndex        =   45
      Top             =   7485
      Visible         =   0   'False
      Width           =   1500
   End
   Begin XtremeSuiteControls.TabControl tbcPage 
      Height          =   6780
      Left            =   -15
      TabIndex        =   47
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
      Left            =   855
      ScaleHeight     =   7230
      ScaleWidth      =   11400
      TabIndex        =   99
      Top             =   300
      Width           =   11400
      Begin VB.Frame Frame2 
         Height          =   105
         Left            =   1050
         TabIndex        =   103
         Top             =   2535
         Width           =   10290
      End
      Begin VB.Frame Frame1 
         Height          =   105
         Left            =   1050
         TabIndex        =   102
         Top             =   4050
         Width           =   10275
      End
      Begin VB.Frame frameLinkMan 
         Height          =   105
         Left            =   1320
         TabIndex        =   101
         Top             =   1020
         Width           =   10020
      End
      Begin VB.TextBox txtOtherWaring 
         Height          =   360
         Left            =   1725
         MaxLength       =   100
         TabIndex        =   40
         Top             =   525
         Width           =   9630
      End
      Begin VB.TextBox txtMedicalWarning 
         Height          =   360
         Left            =   6135
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   120
         Width           =   4860
      End
      Begin VB.ComboBox cboBH 
         Height          =   360
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   120
         Width           =   1410
      End
      Begin VB.ComboBox cboBloodType 
         Height          =   360
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   120
         Width           =   1410
      End
      Begin VB.CommandButton cmdMedicalWarning 
         Caption         =   "…"
         Height          =   330
         Left            =   10995
         TabIndex        =   100
         Top             =   135
         Width           =   330
      End
      Begin VSFlex8Ctl.VSFlexGrid vsLinkMan 
         Height          =   975
         Left            =   30
         TabIndex        =   41
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
         TabIndex        =   43
         Top             =   4380
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
         Left            =   30
         TabIndex        =   42
         Top             =   2880
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
      Begin VB.Label lblInoculate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "接种情况"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -420
         TabIndex        =   110
         Top             =   2475
         Width           =   1860
      End
      Begin VB.Label lblOtherInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "其他信息"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   -420
         TabIndex        =   109
         Top             =   4005
         Width           =   1860
      End
      Begin VB.Label lblLinkman 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "联系人信息"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -300
         TabIndex        =   108
         Top             =   945
         Width           =   1860
      End
      Begin VB.Label lblOtherWaring 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "其他医学警示"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   15
         TabIndex        =   107
         Top             =   585
         Width           =   1875
      End
      Begin VB.Label lblMedicalWarning 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "医学警示"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4665
         TabIndex        =   106
         Top             =   173
         Width           =   1860
      End
      Begin VB.Label lblRH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "RH"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2940
         TabIndex        =   105
         Top             =   173
         Width           =   885
      End
      Begin VB.Label lblBloodType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "血型"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   870
         TabIndex        =   104
         Top             =   150
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmPatiInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'--------------------------------------------------------------------------------------------
'程序入口相关变量
Private mbytFun As Byte  '0-编辑或查看病人信息;2-新建档病人
Private mfrmMain As Object
Private mstrPrivs As String
Private mlng病人ID As Long '传入的病人ID
Private mlng科室ID As Long '传入的科室ID
'--------------------------------------------------------------------------------------------
'模块参数
Private Type Ty_Para
    bln家庭地址输入    As Boolean      '家庭地址输入是否联想
    bln门诊号有效性检查 As Boolean
    bln自动门诊号 As Boolean
    bln结构化地址录入 As Boolean
    bln乡镇地址结构化 As Boolean
    
    bln监护人录入 As Boolean
    int监护人年龄 As Integer
End Type
Private mty_Para As Ty_Para
'--------------------------------------------------------------------------------------------
'相关对象变量定义
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private mobjKeyboard As Object
Private mobjPlugIn As Object '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
Private mlngPlugInHwnd As Long
Private mblnPlugin As Boolean '插建是否创建成功
Private mstrPlugChange As String
Private mobjPubPatient As Object
Private mblnNewPatient As Boolean
Private mblnOK As Boolean '是否确认成功

'--------------------------------------------------------------------------------------------
'模块级变量
Private mblnChange As Boolean
Private mrs家庭地址 As ADODB.Recordset  '缓存家庭地址,初始时读取地区表
Private mrsBaseDict As ADODB.Recordset '国籍,民族,婚姻状况,职业
Private mrsEMPIOut As ADODB.Recordset 'EMPI返回的数据

Private mrs费别 As ADODB.Recordset
Private mintNOLength As Integer '门诊号长度

Private mbln扫描身份证 As Boolean '判断病人信息是否是通过扫描身份证得到
Private mbln扫描身份证签约 As Boolean
Private mintDefaultBlood As Integer '默认血型序号
Private Enum mPageIndex
    基本 = 1
    健康档案 = 2
    附加信息 = 3
End Enum
Private Const C_InoculateHeader = "接种日期,4,2400,1;接种名称,4,2400,1;接种日期,4,2400,1;接种名称,4,2400,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_LinkManColumHeader = "姓名,4,1200,1;关系,4,2400,1;身份证号,4,2400,1;电话,4,1200,1;附加信息,4,2400,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_OtherInfoColumHeader = "信息名,4,2400,1;信息值,4,2400,1;信息名,4,2400,1;信息值,4,2400,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_BH = "阴,阳,不详,未查"
Public Event ReturnVisitClick()     '点击复诊复选框改变对应的费别显示
'74430,冉俊明,2014-7-7,挂号中的病人信息编辑功能中提供采集照片功能
Private mstr采集图片 As String '采集图片本地保存路径
Public mlng图像操作 As Long '指明当前对病人图像操作的类型(1-文件 2-采集 3-清除 4-身份证提取)
Private mstrIDImageFile As String
Public mblnSavePati As Boolean '病人照片信息或附加信息是否已保存
Private mblnNameChange As Boolean
Private mblnGetBirth As Boolean '判断是否允许通过年龄计算生日

Private Sub cbo国籍_Change()
    mstrPlugChange = mstrPlugChange & ",国籍"
End Sub

Private Sub cbo婚姻_Change()
    mstrPlugChange = mstrPlugChange & ",婚姻状况"
End Sub

Private Sub cbo家庭地址_Change()
    mstrPlugChange = mstrPlugChange & ",现住址"
End Sub

Private Sub cbo家庭地址_GotFocus()
    Call gobjCommFun.OpenIme(True)
End Sub
Private Sub cbo家庭地址_LostFocus()
    Call gobjCommFun.OpenIme
End Sub

Private Sub cbo家庭地址_KeyDown(KeyCode As Integer, Shift As Integer)
    '此过程处理本机缓存数据的删除,以及按下拉键时弹出下拉列表
    '下拉列表弹出时,如果按下删除键时,则删除缓存记录
    
    Dim str家庭地址 As String
    
    If KeyCode = vbKeyDelete Then
        str家庭地址 = cbo家庭地址.Text
        If Not mrs家庭地址 Is Nothing And mty_Para.bln家庭地址输入 Then
            If mrs家庭地址.State = 1 And str家庭地址 <> "" Then
                If cbo家庭地址.SelText = str家庭地址 And SendMessage(cbo家庭地址.hWnd, CB_GETDROPPEDSTATE, 0, 0) = True Then
                    mrs家庭地址.Filter = "名称='" & str家庭地址 & "'"
                    If Not mrs家庭地址.EOF Then
                        mrs家庭地址.Delete adAffectCurrent
                        mrs家庭地址.Update
                    End If
                End If
            End If
        End If
    ElseIf KeyCode = vbKeyDown And cbo家庭地址.Text <> "" Then
        If SendMessage(cbo家庭地址.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo家庭地址.hWnd, CB_SHOWDROPDOWN, True, 0&)
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
        If mrs家庭地址 Is Nothing Or mty_Para.bln家庭地址输入 = False Then Exit Sub
        
        str家庭地址 = cbo家庭地址.Text                      '此时,如果选择了部分文字,则选择的文字已经被删除
        lng位置 = cbo家庭地址.SelStart
        
        If mrs家庭地址.State = 1 And Len(str家庭地址) > 1 Then
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(str家庭地址, 1))) > 0 Then
                mrs家庭地址.Filter = "简码 like '" & UCase(str家庭地址) & "*'"
            Else
                mrs家庭地址.Filter = "名称 Like '" & str家庭地址 & "*'"
            End If
            
            If Not mrs家庭地址.EOF Then
                
                If mrs家庭地址.RecordCount <> cbo家庭地址.ListCount Then
                    Call SendMessage(cbo家庭地址.hWnd, CB_RESETCONTENT, 0, 0)
                    mrs家庭地址.Sort = "次数 Desc,名称"
                    For i = 1 To mrs家庭地址.RecordCount
                        AddComboItem cbo家庭地址.hWnd, CB_ADDSTRING, 0, mrs家庭地址!名称
                        mrs家庭地址.MoveNext
                    Next
                    If SendMessage(cbo家庭地址.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo家庭地址.hWnd, CB_SHOWDROPDOWN, True, 0&)
                                        
                    cbo家庭地址.Text = str家庭地址
                    cbo家庭地址.SelStart = lng位置
                End If
            Else
                Call SendMessage(cbo家庭地址.hWnd, CB_SHOWDROPDOWN, False, 0&)
            End If
        ElseIf str家庭地址 = "" Then
            cbo家庭地址.Clear
            Call SendMessage(cbo家庭地址.hWnd, CB_SHOWDROPDOWN, False, 0&)
        End If
    End If
End Sub

Private Sub cbo家庭地址_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim str简码 As String
    Dim str家庭地址 As String
    Dim lng中间输入点 As Long
    
    If (mrs家庭地址 Is Nothing Or mty_Para.bln家庭地址输入 = False) And KeyAscii <> 13 Then Exit Sub
    
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
            mrs家庭地址.Filter = "简码 like '" & UCase(str家庭地址) & "*'"
        Else
            mrs家庭地址.Filter = "名称 Like '" & str家庭地址 & "*'"
        End If
        
        If Not mrs家庭地址.EOF Then
            If mrs家庭地址.RecordCount <> cbo家庭地址.ListCount Then
                Call SendMessage(cbo家庭地址.hWnd, CB_RESETCONTENT, 0, 0)
                mrs家庭地址.Sort = "次数 Desc,名称"
                For i = 1 To mrs家庭地址.RecordCount
                    AddComboItem cbo家庭地址.hWnd, CB_ADDSTRING, 0, mrs家庭地址!名称
                    mrs家庭地址.MoveNext
                Next
                If SendMessage(cbo家庭地址.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo家庭地址.hWnd, CB_SHOWDROPDOWN, True, 0&)
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
            Call SendMessage(cbo家庭地址.hWnd, CB_RESETCONTENT, 0, 0)
            If SendMessage(cbo家庭地址.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then Call SendMessage(cbo家庭地址.hWnd, CB_SHOWDROPDOWN, False, 0&)
            KeyAscii = 0
            cbo家庭地址.Text = str家庭地址
            cbo家庭地址.SelStart = Len(cbo家庭地址.Text)
        End If
        
        If lng中间输入点 > 0 Then cbo家庭地址.SelStart = lng中间输入点: cbo家庭地址.SelText = ""
        
    ElseIf KeyAscii = 13 Then
        'a.在没有选中任何文字,且输入内容为空,光标为于末端时,确认输入,并保存信息到本地缓存
        Call SendMessage(cbo家庭地址.hWnd, CB_SHOWDROPDOWN, False, 0&)
        
        If cbo家庭地址.Text = "" Then
            If txtPatient.Text <> "" Then
                Exit Sub
            Else
                Call gobjCommFun.PressKey(vbKeyTab): Exit Sub
            End If
        End If
        
        '下拉列表弹出时按回车,则定位到末尾
        If cbo家庭地址.SelText = cbo家庭地址.Text Then cbo家庭地址.SelStart = Len(cbo家庭地址.Text): Exit Sub
        
        If mrs家庭地址 Is Nothing Then Call gobjCommFun.PressKey(vbKeyTab): Exit Sub
        If mrs家庭地址.State = 0 Then Call gobjCommFun.PressKey(vbKeyTab): Exit Sub
        If gobjCommFun.ActualLen(cbo家庭地址.Text) > 100 Then Call gobjCommFun.PressKey(vbKeyTab): Exit Sub
       
        'a.非下拉状态下按回车,没有选择文本
        If cbo家庭地址.SelText = "" Then
            str家庭地址 = cbo家庭地址.Text
            mrs家庭地址.Filter = "名称='" & str家庭地址 & "'"
            If mrs家庭地址.EOF Then
                str简码 = Mid(gobjCommFun.zlGetSymbol(str家庭地址), 1, 10)
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
                
                If gobjCommFun.IsCharAlpha(str家庭地址) Then
                    If mrs家庭地址.RecordCount = 1 Then
                        cbo家庭地址.Text = mrs家庭地址!名称
                    Else
                        Call SendMessage(cbo家庭地址.hWnd, CB_SHOWDROPDOWN, True, 0&)
                        Exit Sub
                    End If
                End If
            End If
            
            Call gobjCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo联系人关系_Change()
    mstrPlugChange = mstrPlugChange & ",联系人关系"
End Sub

Private Sub cbo联系人关系_Click()
    With cbo联系人关系
        If .ListIndex = 8 And txt其他关系.Visible = False Then
            .Width = 1225: txt其他关系.Visible = True
        ElseIf .ListIndex <> 8 And txt其他关系.Visible Then
            .Width = 2425: txt其他关系.Visible = False
        ElseIf .ListIndex = -1 Then
            .Width = 2425
        End If
    End With
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("关系") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("关系")) = gobjCommFun.GetNeedName(cbo联系人关系.Text)
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("附加信息")) = gobjCommFun.GetNeedName(txt其他关系.Text)
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
    Dim strSql As String
    Dim vRect As RECT
    Dim strTemp As String
    Dim blnCancel As Boolean
    
'    vRect = gobjControl.GetControlRect(txtMedicalWarning.hWnd)
    
    strSql = "" & _
    "       Select 编码 as ID,名称,简码 From 医学警示 Where 名称 Not Like '其他%'"
    Set rsTemp = gobjDatabase.ShowSQLMultiSelect(Me, strSql, 0, "医学警示", False, txtMedicalWarning.Text, "", False, False, False, vRect.Left, vRect.Top - 180, 500, blnCancel, False, True)
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
        .Flags = cdlOFNHideReadOnly
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
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd区域_Click()
    If zl_SelectAndNotAddItem(Me, txt区域, "", "区域", "区域选择", True, False) = False Then
        Exit Sub
    End If
End Sub

Private Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As PointAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

Public Function zl_AutoAddBaseItem(ByVal strTable As String, str编码 As String, str名称 As String, _
    Optional strTittle As String = "增加项目", Optional blnMsg As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:自动增加项目信息(只针对有编码,名称的信息增加(只增加：编码和名称,简码)
    '--入参数:
    '--出参数:
    '--返  回:增加成功,返回true,否则返回false
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, strSql As String
    Dim int编码 As Integer, strCode As String, strSpecify As String
    zl_AutoAddBaseItem = False
    If blnMsg = True Then
        If MsgBox("没有找到你输入的" & strTable & "，你要把它加入" & strTable & "中吗？", vbYesNo + vbQuestion, strTittle) = vbNo Then
            Exit Function
        End If
    End If
    
    Err = 0: On Error GoTo Errhand:
    
    strSql = "SELECT Nvl(MAX(LENGTH(编码)), 2) As Length FROM  " & strTable
    gobjDatabase.OpenRecordset rsTemp, strSql, strTittle
    
    int编码 = rsTemp!length
    
    strSql = "SELECT Nvl(MAX(LPAD(编码," & int编码 & ",'0')),'00') As Code FROM  " & strTable
    gobjDatabase.OpenRecordset rsTemp, strSql, strTittle
    strCode = rsTemp!Code
    
    int编码 = Len(strCode)
    strCode = strCode + 1
    
    If int编码 >= Len(strCode) Then
    strCode = String(int编码 - Len(strCode), "0") & strCode
    End If
    strSpecify = gobjCommFun.SpellCode(str名称)
    
    
    strSql = "ZL_" & strTable & "_INSERT('" & strCode & "','" & str名称 & "','" & strSpecify & "')"
    gobjDatabase.ExecuteProcedure strSql, strTittle
    str编码 = strCode
    zl_AutoAddBaseItem = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function zl_SelectAndNotAddItem(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    ByVal strTable As String, ByVal strTittle As String, Optional blnOnlyName As Boolean = False, _
    Optional bln未找到增加 As Boolean = False, Optional strOra过程 As String, Optional strWhere As String, _
    Optional bln站点 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:多功能选择器
    '参数:objCtl-文本框控件
    '     strKey-要搜索的值
    '     strTable-表名
    '     strTittle-选择器名称
    '     bln站点-是否进行站点限制
    '返回:
    '编制:刘兴宏
    '日期:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, str编码 As String, str名称 As String
    Dim vRect As RECT, sngX As Single, sngY As Single, strSql As String
    Dim rsTemp  As ADODB.Recordset
    'gobjDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    str名称 = strKey
    
    If strTable = "区域" Then
        strSql = "Select rownum as ID,a.* From " & strTable & " a where 1=1 And Nvl(级数,0) <3 "
    Else
        strSql = "Select rownum as ID,a.* From " & strTable & " a where 1=1 "
    End If
    
    If strKey <> "" Then
        strSql = strSql & _
        "   And ((名称) like [1] or  编码  like [1] or  简码  like  upper([1]))  "
    End If
    strSql = strSql & strWhere & _
    "   order by 编码"
    strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        If UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
            Call CalcPosition(sngX, sngY, objCtl.MsfObj)
            lngH = objCtl.MsfObj.CellHeight
        Else
            Call CalcPosition(sngX, sngY, objCtl)
            lngH = objCtl.CellHeight
        End If
        sngY = sngY - lngH
    Else
'        vRect = gobjControl.GetControlRect(objCtl.hWnd)
        lngH = objCtl.Height
        sngX = gobjControl.GetControlRect(objCtl.hWnd).Left - 15
        sngY = gobjControl.GetControlRect(objCtl.hWnd).Top
    End If
    
    Set rsTemp = gobjDatabase.ShowSQLSelect(frmMain, strSql, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey)
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then gobjControl.TxtSelAll objCtl
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        If bln未找到增加 Then
            If gobjCommFun.IsCharChinese(str名称) = False Then GoTo NOAdd::
            If MsgBox("注意:" & vbCrLf & _
                   "     未找到相关的" & strTable & ",是否增加“" & str名称 & "”？", vbQuestion + vbYesNo + vbDefaultButton2, strTable) = vbNo Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then gobjControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If zl_AutoAddBaseItem(strTable, str编码, str名称, strTable & "增加", False) = False Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then gobjControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
                With objCtl
                    .TextMatrix(.Row, .Col) = IIf(blnOnlyName, str名称, str编码 & "-" & str名称)
                    If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                        .Cell(flexcpData, .Row, .Col) = str名称
                    End If
                End With
            Else
                If gobjControl.IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
                objCtl.Text = IIf(blnOnlyName, str名称, str编码 & "-" & str名称)
                objCtl.Tag = str名称
                gobjCommFun.PressKey vbKeyTab
            End If
            zl_SelectAndNotAddItem = True
            Exit Function
        Else
NOAdd:
            ShowMsgBox "没有找到满足条件的" & strTable & ",请检查!"
            If objCtl.Enabled Then objCtl.SetFocus
            If UCase(TypeName(objCtl)) = UCase("TextBox") Then gobjControl.TxtSelAll objCtl
            Exit Function
        End If
    End If
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        With objCtl
            .TextMatrix(.Row, .Col) = IIf(blnOnlyName, Nvl(rsTemp!名称), Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称))
            If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                .Cell(flexcpData, .Row, .Col) = Nvl(rsTemp!名称)
            Else
                .Text = IIf(blnOnlyName, Nvl(rsTemp!名称), Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称))
            End If
        End With
    Else
        If gobjControl.IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
        objCtl.Text = Nvl(rsTemp!名称)
        objCtl.Tag = Nvl(rsTemp!名称)
        gobjCommFun.PressKey vbKeyTab
    End If
    zl_SelectAndNotAddItem = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter = 1 Then Resume
End Function

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind        As Long
    Dim lng病人ID           As Long
    Dim strPassWord         As String
    Dim strErrMsg           As String
    Dim strTmp              As String

    '获取病人ID
    If gobjSquare.objSquareCard.zlGetPatiID("身份证", strID, False, lng病人ID, strPassWord, strErrMsg, , , , False) = False Then lng病人ID = 0

    If mbytFun = 2 Then
        '建档
        txt身份证号.Text = strID
        txtPatient.Text = strName
        Call gobjControl.CboLocate(cbo性别, strSex)
        Call gobjControl.CboLocate(cbo民族, strNation)
        txt出生日期.Text = Format(datBirthDay, "yyyy-MM-dd")
        txt出生时间.Text = "00:00"
        cbo家庭地址.Text = strAddress
        txtRegLocation.Text = strAddress
        padd家庭地址.Value = IIf(Trim(padd家庭地址.Value) = "", strAddress, padd家庭地址.Value)
        padd户口地址.Value = strAddress
        '74430,冉俊明,2014-7-7,挂号中的病人信息编辑功能中提供采集照片功能
        Call LoadIDImage
        Call zlQueryEMPIPatiInfo
    Else
        If MsgBox("是否使用身份证扫描信息更新当前病人信息？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbYes Then
            txt身份证号.Text = strID
            txtPatient.Text = strName
            Call gobjControl.CboLocate(cbo性别, strSex)
            Call gobjControl.CboLocate(cbo民族, strNation)
            txt出生日期.Text = Format(datBirthDay, "yyyy-MM-dd")
            txt出生时间.Text = "00:00"
            cbo家庭地址.Text = strAddress
            txtRegLocation.Text = strAddress
            padd家庭地址.Value = IIf(Trim(padd家庭地址.Value) = "", strAddress, padd家庭地址.Value)
            padd户口地址.Value = strAddress
            '74430,冉俊明,2014-7-7,挂号中的病人信息编辑功能中提供采集照片功能
            Call LoadIDImage
            Call zlQueryEMPIPatiInfo
        End If
    End If

End Sub

Private Function GetPatiByID(str类型 As String, strValue As String) As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据类型来获取不同条件下的病人信息
    '入参:str类型：查询条件类型 strValue 条件值
    '返回:病人信息集合
    '编制:王吉
    '日期:2012-08-31 04:36:33
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo ErrHandl
    strSql = "" & _
    "   Select 病人ID,门诊号,住院号,就诊卡号,卡验证码,费别,医疗付款方式,姓名,性别,年龄,出生日期,出生地点,身份证号,其他证件,身份,职业,民族,国籍,籍贯,区域,学历,婚姻状况,家庭地址,家庭电话,家庭地址邮编,监护人," & _
    "   联系人姓名,联系人关系,联系人地址,联系人电话,户口地址,户口地址邮编,Email,QQ,合同单位ID,工作单位,单位电话,单位邮编,单位开户行,单位帐号,担保人,担保性质,就诊时间,就诊状态,就诊诊室,住院次数,当前科室ID,当前床号," & _
    "   入院时间,出院时间,在院,IC卡号,健康号,医保号,险类,查询密码,登记时间,停用时间,锁定,联系人身份证号,结算模式 " & _
    "   From 病人信息 " & _
    "   Where " & str类型 & "=[1]"
    
    Set GetPatiByID = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, strValue)
    Exit Function
ErrHandl:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Sub cbo费别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 And cbo费别.ListIndex <> -1 Then Call gobjCommFun.PressKey(vbKeyTab)
    
    If SendMessage(cbo费别.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo费别.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo费别.ListIndex = lngIdx
    If cbo费别.ListIndex = -1 And cbo费别.ListCount > 0 Then cbo费别.ListIndex = 0
End Sub

Private Sub cbo付款方式_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo付款方式.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo付款方式.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo付款方式.ListIndex = lngIdx
    If cbo付款方式.ListIndex = -1 And cbo付款方式.ListCount > 0 Then cbo付款方式.ListIndex = 0
End Sub

Private Sub cbo国籍_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo国籍.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo国籍.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo国籍.ListIndex = lngIdx
    If cbo国籍.ListIndex = -1 And cbo国籍.ListCount > 0 Then cbo国籍.ListIndex = 0
End Sub

Private Sub cbo婚姻_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo婚姻.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo婚姻.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo婚姻.ListIndex = lngIdx
    If cbo婚姻.ListIndex = -1 And cbo婚姻.ListCount > 0 Then cbo婚姻.ListIndex = 0
End Sub

Private Sub cbo民族_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo民族.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo民族.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo民族.ListIndex = lngIdx
    If cbo民族.ListIndex = -1 And cbo民族.ListCount > 0 Then cbo民族.ListIndex = 0
End Sub

Private Sub cbo年龄单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call gobjCommFun.PressKey(vbKeyTab)
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
                txt出生日期.Text = Format(strBirth, "YYYY-MM-DD")
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
    If SendMessage(cbo性别.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo性别.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo性别.ListIndex = lngIdx
    If cbo性别.ListIndex = -1 And cbo性别.ListCount > 0 Then cbo性别.ListIndex = 0
    
End Sub

Private Sub cbo职业_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo职业.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo职业.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo职业.ListIndex = lngIdx
    If cbo职业.ListIndex = -1 And cbo职业.ListCount > 0 Then cbo职业.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    If txtPatient.Text <> "" And mbytFun <> 0 Then
        If MsgBox("是否终止新病人录入?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Sub
    End If
    mlng病人ID = 0
    mstrPlugChange = ""
    Unload Me
    Exit Sub
ErrOther:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Function CheckTextLength(strName As String, txtObj As TextBox) As Boolean
'功能:检查并提示文本框输入长度是否超限
    CheckTextLength = gobjControl.TxtCheckInput(txtObj, strName, , True)
End Function

Public Function CheckExistsMCNO(ByVal strMCNO As String) As Boolean
'功能:检查医保号是否已存在
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    strSql = "Select 1 From 病人信息 Where 医保号 = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, strMCNO)
    If rsTmp.RecordCount > 0 Then
        MsgBox "请检查,输入的医保号已存在!", vbInformation, gstrSysName
        CheckExistsMCNO = True
    End If
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function


Private Function CheckValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入是否正确
    '编制:刘兴洪
    '日期:2011-01-07 18:13:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long, strSimilar As String, i As Integer, strMCAccount As String
    Dim strSql As String, rsTmp As ADODB.Recordset, intQuery As Integer
    Dim blnPlugInCheck As Boolean, str出生时间 As String
    Dim strBirthDay As String, strAge As String, strSex As String, strErrInfo As String, strInfo As String
   
    
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

    If txtPatient.Text = "" Then
        MsgBox "请输入病人的姓名！", vbInformation, gstrSysName
        If txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    
    If txtPatiMCNO(0).Text <> "" Or txtPatiMCNO(1).Text <> "" Then
        If txtPatiMCNO(0).Text <> txtPatiMCNO(1).Text Then
            MsgBox "请检查,两次输入的医保号不一致！", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
        If gobjCommFun.ActualLen(txtPatiMCNO(0).Text) > txtPatiMCNO(0).MaxLength Then
            MsgBox "请检查,医保号最大长度不能超过" & txtPatiMCNO(0).MaxLength & "个字符！", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
    End If
    
    If CheckTextLength("姓名", txtPatient) = False Then Exit Function
    If CheckTextLength("出生地点", txtBirthLocation) = False Then Exit Function
    If mty_Para.bln结构化地址录入 Then
        If Not CheckStructAddr(padd家庭地址, padd家庭地址.MaxLength) Then Exit Function
        If Not CheckStructAddr(padd户口地址, padd户口地址.MaxLength) Then Exit Function
    Else
        If gobjCommFun.ActualLen(cbo家庭地址.Text) > glngMax家庭地址 Then
            MsgBox "家庭住址输入过长，只允许输入" & glngMax家庭地址 & "个字符或" & glngMax家庭地址 \ 2 & "个汉字，请检查!", vbInformation, gstrSysName
            cbo家庭地址.SetFocus: Exit Function
        End If
        If CheckTextLength("户口地址", txtRegLocation) = False Then Exit Function
    End If
    If CheckTextLength("户口地址邮编", txt户口地址邮编) = False Then Exit Function
    If CheckTextLength("年龄", txt年龄) = False Then Exit Function
    If CheckTextLength("出生地点", txtBirthLocation) = False Then Exit Function
    '83062
    For i = 1 To msh过敏.Rows - 1
        If gobjCommFun.ActualLen(msh过敏.TextMatrix(i, 1)) > 100 Then
            MsgBox "病人过敏药物反应输入过长，只允许输入100个字符或50个汉字，请检查！", vbInformation, gstrSysName
            If msh过敏.Enabled And msh过敏.Visible Then msh过敏.SetFocus
            Exit Function
        End If
        If gobjCommFun.ActualLen(msh过敏.TextMatrix(i, 0)) > 60 Then
            MsgBox "病人过敏药物名称输入过长，只允许输入60个字符或30个汉字，请检查！", vbInformation, gstrSysName
            If msh过敏.Enabled And msh过敏.Visible Then msh过敏.SetFocus
            Exit Function
        End If
    Next i
    '69026,冉俊明,2014-8-11,年龄有效性检查
    '76703,冉俊明,2014-8-15
    If txt年龄.Enabled And txt年龄.Visible Then
        If mobjPubPatient Is Nothing Then Exit Function
        If mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), _
                IIf(txt出生日期.Text = "____-__-__", "", txt出生日期.Text) & _
                IIf(txt出生时间.Text = "__:__", "", " " & txt出生时间.Text)) = False Then
            txt年龄.SetFocus:  Exit Function
        End If
    End If
    
    If IsDate(gobjCommFun.GetIDCardDate(txt身份证号.Text)) Then
        If Format(gobjCommFun.GetIDCardDate(txt身份证号.Text), "yyyy-mm-dd") <> Format(txt出生日期.Text, "yyyy-mm-dd") Then
            intQuery = MsgBox("输入的身份证号与输入的出生日期不一致，使用身份证号获取的出生日期吗？", vbQuestion + vbYesNoCancel, gstrSysName)
            If intQuery = 6 Then
                txt出生日期.Text = gobjCommFun.GetIDCardDate(txt身份证号.Text)
            ElseIf intQuery = 2 Then
                CheckValied = False
                Exit Function
            End If
        End If
    End If
    
    If IsDate(txt出生日期.Text) Then
        '76669，李南春,2014-8-15,年龄与出生日期检查
        str出生时间 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        If CDate(str出生时间) > gobjDatabase.Currentdate Then
            If MsgBox("出生时间：" & str出生时间 & " 超过了当前系统时间。" & _
                vbCrLf & vbCrLf & "请检查年龄或出生日期的正确性 ，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If txt出生日期.Enabled And txt出生日期.Visible Then txt出生日期.SetFocus
                Exit Function
            End If
        End If
        If mty_Para.bln监护人录入 And Trim(txt监护人.Text) = "" Then
            '61945 监护人录入 检查
            strSql = "Select Floor(Months_Between(Sysdate, [1]) / 12) as 年龄 From Dual"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(txt出生日期.Text))
            If Not rsTmp Is Nothing Then
                If Val(Nvl(rsTmp!年龄)) <= mty_Para.int监护人年龄 And mty_Para.int监护人年龄 <> 0 Then
                    MsgBox "病人在[" & mty_Para.int监护人年龄 & "岁]下必须录入监护人,请检查!"
                    Set rsTmp = Nothing
                    If txt监护人.Enabled And txt监护人.Visible Then txt监护人.SetFocus
                    Exit Function
                End If
            End If
        End If
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
        If mobjPubPatient.CheckPatiIdcard(Trim(txt身份证号.Text), strBirthDay, strAge, strSex, strErrInfo) Then
            '新病人或调整无业务数据的已有病人信息时提示是否调整不一致的基本信息
            If strSex <> NeedName(cbo性别.Text) Then strInfo = "性别"
            If strAge <> Trim(txt年龄.Text) & cbo年龄单位 Then strInfo = strInfo & IIf(strInfo = "", "年龄", "、年龄")
            If Format(strBirthDay, "yyyy-mm-dd") <> txt出生日期.Text Then strInfo = strInfo & IIf(strInfo = "", "出生日期", "、出生日期")
            
            If strInfo <> "" Then
                If MsgBox("输入的" & strInfo & "与身份证号的" & strInfo & "不一致，" & _
                        "将根据身份证号修改" & strInfo & "，是否继续？", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                    Call gobjControl.CboLocate(cbo性别, strSex)
                    txt出生日期.Text = Format(strBirthDay, "yyyy-mm-dd")
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
    
    
    If Trim(txt支付密码.Text) <> Trim(txt验证密码.Text) And (Trim(txt支付密码.Text) <> "" Or Trim(txt验证密码.Text) <> "") Then
        MsgBox "两次输入的密码不一致,请重新输入", vbOKOnly + vbInformation, gstrSysName
        txt支付密码.Text = "": txt验证密码.Text = ""
        If txt支付密码.Visible = True And txt支付密码.Enabled = True Then txt支付密码.SetFocus
        Exit Function
    End If
    
    '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    If Not mobjPlugIn Is Nothing And mlngPlugInHwnd <> 0 Then '保存插件附加信息前的数据有效性检查
        On Error Resume Next
        blnPlugInCheck = mobjPlugIn.PatiInfoSaveBefore(mlng病人ID)
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
    
    CheckValied = True
End Function

Private Function SimilarIDs(str身份证号 As String) As String
'功能：检查病人是否存在相似信息
'返回：相似记录的病人ID串,如"234,235,236"
    On Error GoTo errH
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Integer
    
    strSql = _
        " Select 病人ID,姓名,Nvl(身份证号,'未登记') 身份证号,门诊号,Nvl(家庭地址,'未登记') 地址,To_Char(登记时间,'YYYY-MM-DD') 登记时间 " & _
        " From 病人信息 Where 身份证号=[1]" & _
        " Order by 病人ID Desc"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", str身份证号)
    
    For i = 1 To rsTmp.RecordCount
        SimilarIDs = SimilarIDs & "|ID:" & rsTmp!病人ID & ",姓名:" & rsTmp!姓名 & ",门诊号:" & Nvl(rsTmp!门诊号, "无") & ",身份证号:" & rsTmp!身份证号 & ",地址:" & rsTmp!地址 & ",登记日期:" & rsTmp!登记时间
        rsTmp.MoveNext
    Next
    SimilarIDs = Mid(SimilarIDs, 2)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Function Exist门诊号(str门诊号 As String, Optional lng病人ID As Long) As Boolean
'功能：判断指定门诊号是否已经存在于数据库中
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select 病人ID From 病人信息 Where 门诊号=[1] And 病人ID<>[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", str门诊号, lng病人ID)
    If rsTmp.RecordCount > 0 Then Exist门诊号 = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Function Exist手机号(str手机号 As String, Optional lng病人ID As Long) As Boolean
'功能：判断指定手机号是否已经存在于数据库中
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select 病人ID From 病人信息 Where 手机号=[1] And 病人ID<>[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", str手机号, lng病人ID)
    If rsTmp.RecordCount > 0 Then Exist手机号 = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Sub cmdOK_Click()
    Dim lng病人ID As Long
    If SaveData(lng病人ID) = False Then Exit Sub
    Call CloseIDCard
    mblnOK = True
    Unload Me
End Sub
Private Function SaveData(ByRef lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存数据
    '出参:lng病人ID-返回当前病人ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-10-27 14:13:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtSysDate As Date, strMCAccount As String, strCardNO As String, strPassWord As String
    Dim strNO As String, lngDept As Long, strDate As String, str出生日期 As String
    Dim strTmp As String, blnTrans As Boolean, strErrMsg As String, blnNewPati As Boolean
    Dim str门诊号 As String, byt类型 As Byte, i As Integer
    Dim strSql As String
    Dim cllPro As Collection
    
    Err = 0: On Error GoTo errHandle
    txtPatient.Text = Trim(txtPatient.Text)
    txt年龄.Text = Trim(txt年龄.Text)
    txt年龄.Tag = txt年龄.Text

    '相关的输入检查
    If CheckValied = False Then Exit Function
    If Not ((mbytFun = 0 And mlng病人ID <> 0) Or mbytFun = 2) Then SaveData = True: Exit Function

    strMCAccount = Trim(txtPatiMCNO(0).Text)
    If txt出生时间 = "__:__" Then
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & " " & txt出生时间.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
   
    If Len(txt门诊号.Text) > mintNOLength + 1 And mintNOLength > 0 And mty_Para.bln门诊号有效性检查 Then
        MsgBox "注意,输入的门诊号过大,请确认是否输入正常!", vbInformation, gstrSysName
        txt门诊号.SetFocus
        txt门诊号.SelStart = 0: txt门诊号.SelLength = Len(txt门诊号.Text)
        Exit Function
    End If
    
   
    dtSysDate = gobjDatabase.Currentdate
    strDate = "To_Date('" & Format(dtSysDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    If Exist门诊号(txt门诊号.Text, IIf(mlng病人ID <> 0 And mbytFun = 0, mlng病人ID, 0)) Then
        str门诊号 = gobjDatabase.GetNextNo(3)
        If Len(str门诊号) > txt门诊号.MaxLength Then
            MsgBox "当前门诊号已经被其它病人使用,系统自动更换门诊号为:" & str门诊号 & _
                   vbCrLf & "但超过了允许的最大门诊号长度:" & txt门诊号.MaxLength & "位,请输入一个门诊号!", vbInformation, gstrSysName
            If txt门诊号.Enabled Then txt门诊号.SetFocus
            Exit Function
        End If
        txt门诊号.Text = str门诊号
    End If

    If mbytFun = 0 And mlng病人ID <> 0 Then
        lng病人ID = mlng病人ID
        byt类型 = 3
    Else
        lng病人ID = gobjDatabase.GetNextNo(1)
        byt类型 = 1: blnNewPati = True
    End If
    Set cllPro = New Collection
    mlng病人ID = lng病人ID
    strSql = _
    "zl_挂号病人病案_INSERT(" & byt类型 & "," & lng病人ID & "," & txt门诊号.Text & "," & _
    "'" & strCardNO & "','" & strPassWord & "'," & _
    "'" & txtPatient.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, "") & "'," & _
    "'" & NeedName(cbo费别.Text) & "','" & NeedName(cbo付款方式.Text) & "'," & _
    "'" & NeedName(cbo国籍.Text) & "','" & NeedName(cbo民族.Text) & "','" & NeedName(cbo婚姻.Text) & "'," & _
    "'" & NeedName(cbo职业.Text) & "','" & txt身份证号.Text & "','" & txt单位名称.Text & "'," & _
    Val(txt单位名称.Tag) & ",'" & txt单位电话.Text & "','" & txt单位邮编.Text & "','" & IIf(mty_Para.bln结构化地址录入, Trim(padd家庭地址.Value), cbo家庭地址.Text) & "'," & _
    "'" & txt家庭电话.Text & "','" & txt家庭邮编.Text & "'," & strDate & ",''," & str出生日期 & ",'" & strMCAccount & "','" & "" & "'," & _
    "NULL," & IIf(Trim(txt区域.Text) = "", "NULL,", "'" & Trim(txt区域.Text) & "',") & _
     "'" & IIf(mty_Para.bln结构化地址录入, Trim(padd户口地址.Value), Trim(txtRegLocation.Text)) & "','" & Trim(txt户口地址邮编.Text) & "'," & IIf(Trim(txt联系人身份证.Text) = "", "NULL,", "'" & Trim(txt联系人身份证.Text) & "',") & _
    IIf(Trim(txt联系人姓名.Text) = "", "NULL,", "'" & Trim(txt联系人姓名.Text) & "',") & _
    IIf(Trim(txt联系人电话.Text) = "", "NULL,", "'" & Trim(txt联系人电话.Text) & "',") & _
    IIf(NeedName(cbo联系人关系.Text) = "", "NULL,", "'" & NeedName(cbo联系人关系.Text) & "',")    '问题号:40005
        
    '监护人_In         In 病人信息.监护人%Type := Null
    strSql = strSql & IIf(Trim(txt监护人.Text) = "", "NULL,", "'" & Trim(txt监护人.Text) & "',")  'lgf
    '54601:刘尔旋,2013-11-27,新增出生地点和户口地址
    strSql = strSql & IIf(Trim(txtBirthLocation.Text) = "", "NULL,", "'" & Trim(txtBirthLocation.Text) & "',")
    strSql = strSql & "'" & txtMobile.Text & "')"
    zlAddArray cllPro, strSql
        
    If mty_Para.bln结构化地址录入 Then
        If padd家庭地址.Value <> "" Then
           strSql = "zl_病人地址信息_update(1," & lng病人ID & ",NULL,3,'" & padd家庭地址.value省 & "','" & _
               padd家庭地址.value市 & "','" & padd家庭地址.value区县 & "','" & padd家庭地址.value乡镇 & "','" & _
               padd家庭地址.value详细地址 & "','" & padd家庭地址.Code & "')"
        Else
           strSql = "zl_病人地址信息_update(2," & lng病人ID & ",NULL,3)"
        End If
        zlAddArray cllPro, strSql
        If padd户口地址.Value <> "" Then
           strSql = "zl_病人地址信息_update(1," & lng病人ID & ",NULL,4,'" & padd户口地址.value省 & "','" & _
               padd户口地址.value市 & "','" & padd户口地址.value区县 & "','" & padd户口地址.value乡镇 & "','" & _
               padd户口地址.value详细地址 & "','" & padd户口地址.Code & "')"
               
               
        Else
           strSql = "zl_病人地址信息_update(2," & lng病人ID & ",NULL,4)"
        End If
        zlAddArray cllPro, strSql
    End If
    
    'str其他关系
    If cbo联系人关系.Text <> "" And txt其他关系.Visible Then
        strSql = "Zl_病人信息从表_Update("
        '病人ID_In 病人信息从表.病人Id%Type
        strSql = strSql & "" & lng病人ID & ","
        '信息名_In 病人信息从表.信息名%Type0
        strSql = strSql & "'联系人附加信息',"
        '信息值_In 病人信息从表.信息值%Type
        strSql = strSql & "'" & txt其他关系.Text & "',"
        '就诊Id_In 病人信息从表.就诊Id%Type
        strSql = strSql & "'')"
        zlAddArray cllPro, strSql
    End If
    Call Add健康卡相关信息(lng病人ID, cllPro)
       
    blnTrans = True
    Call zlExecuteProcedureArrAy(cllPro, Me.Caption, True, False)
    
    '110269:李南春,2016/10/13,保存HIS数据要提交EMPI数据，失败后所有数据都要回退
    If zlSaveEMPIPatiInfo(blnNewPati, mlng病人ID, 0, strErrMsg) = False Then
        gcnOracle.RollbackTrans
        If strErrMsg = "" Then strErrMsg = "向EMPI平台上传病人信息失败！"
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    SaveData = True
    mblnSavePati = True
    
    mstrPlugChange = ""
    '74430,冉俊明,2014-7-7,挂号中的病人信息编辑功能中提供采集照片功能
    Call SavePatiPic(mlng病人ID)
    '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    If Not mobjPlugIn Is Nothing And mlngPlugInHwnd <> 0 Then '保存插件附加信息
        On Error Resume Next
        Call mobjPlugIn.PatiInfoSaveAfter(mlng病人ID)
        Call zlPlugInErrH(Err, "PatiInfoSaveAfter")
        Err.Clear: On Error GoTo 0
    End If
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub padd户口地址_Change()
    If mty_Para.bln结构化地址录入 Then mstrPlugChange = mstrPlugChange & ",户口地址"
End Sub

Private Sub padd家庭地址_Change()
    If mty_Para.bln结构化地址录入 Then mstrPlugChange = mstrPlugChange & ",现住址"
End Sub


Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmd单位名称_Click()
    Call SearchUnit("", txt单位名称)
End Sub

Private Sub cmd过敏_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    Dim i As Integer
    
    strSql = _
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

    Set rsTmp = frmPubSel.ShowSelect(Me, strSql, 2, "过敏药物", , msh过敏.TextMatrix(msh过敏.Row, 0), "请从下面的药品中选择一项作为病人过敏药物。")
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
    tbcPage.Left = 0
    tbcPage.Width = Me.ScaleWidth
    If (mbytFun = 0 And mlng病人ID = 0) Then '绑定就诊卡模式不提供取消按钮,以防Unload窗体,因为之前提取病人身份时加载的信息会被清除
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
        If txtPatient.Visible = True And txtPatient.Enabled Then
            txtPatient.SetFocus
        ElseIf txt门诊号.Enabled And txt门诊号.Visible Then
            txt门诊号.SetFocus
        ElseIf txt出生日期.Enabled And txt出生日期.Visible Then
            txt出生日期.SetFocus
        End If
    End If
    
    mbln扫描身份证 = False
    mbln扫描身份证签约 = True
    SetCtrVisibleAndMove
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
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If InStr(1, "txtPatient,txt密码,lvwItems,txt年龄,cbo年龄单位,txt出生日期,msh过敏,txt过敏,txtPatiMCNO,txt区域,vsInoculate,cbo家庭地址", Me.ActiveControl.Name) <= 0 Then
            KeyAscii = 0
            Call gobjCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Function GetBaseDict() As ADODB.Recordset
'功能：从字典中读取数据
    Dim strSql As String, strTmp As String, arrTmp As Variant, i As Integer
    strTmp = "国籍,民族,婚姻状况,职业,社会关系"
    arrTmp = Split(strTmp, ",")
    For i = 0 To UBound(arrTmp)
        strTmp = arrTmp(i)
        If strSql = "" Then
            strSql = "Select '" & strTmp & "' 类别,编码,名称,Nvl(缺省标志,0) as 缺省 From " & strTmp
        Else
            strSql = strSql & " Union all Select '" & strTmp & "' 类别,编码,名称,Nvl(缺省标志,0) as 缺省 From " & strTmp
        End If
    Next
    strSql = strSql & " Order by 类别,编码"
    
    On Error GoTo errH
    Set GetBaseDict = gobjDatabase.OpenSQLRecord(strSql, "获取国籍,民族,婚姻状况,职业,社会关系")
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function


Private Sub InitData()
'功能：初始化必要数据
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, lngTmp As Long
    Dim lngCardType As Long
    Dim strSql As String
        
       
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
    
    cbo费别.Clear
    Call Init费别(True, True)
    
    cbo年龄单位.AddItem "岁"
    cbo年龄单位.AddItem "月"
    cbo年龄单位.AddItem "天"
    cbo年龄单位.ListIndex = 0
    
    '性别
    strSql = "Select '性别' as 类别,编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 性别 Union All " & _
             " Select '医疗付款方式' as 类别,编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 医疗付款方式 " & _
             " Order by 类别,编码"
    Set rsTmp = New ADODB.Recordset
    Call gobjDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
    rsTmp.Filter = "类别='性别'"
    cbo性别.Clear
    Do While Not rsTmp.EOF
        cbo性别.AddItem rsTmp!编码 & "-" & rsTmp!名称
'        If rsTmp!名称 = gstr性别 Then
'            For i = 0 To cbo性别.ListCount - 1
'                cbo性别.ItemData(i) = 0
'            Next
'            cbo性别.ItemData(cbo性别.NewIndex) = 1
'            cbo性别.ListIndex = cbo性别.NewIndex
'        End If
        
        If rsTmp!缺省 = 1 And cbo性别.ListIndex = -1 Then
            cbo性别.ItemData(cbo性别.NewIndex) = 1
            cbo性别.ListIndex = cbo性别.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    
    '问题号:110155
    rsTmp.Filter = "类别='医疗付款方式'"
    cbo付款方式.Clear
    Do While Not rsTmp.EOF
        cbo付款方式.AddItem rsTmp!编码 & "-" & rsTmp!名称

        If Val(Nvl(rsTmp!缺省)) = 1 And cbo付款方式.ListIndex = -1 Then
            cbo付款方式.ItemData(cbo付款方式.NewIndex) = 1
            cbo付款方式.ListIndex = cbo付款方式.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    
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
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
        .Visible = False
    End With
    '问题号:56599
    Call Init过敏药物
    
    If cbo年龄单位.Tag = "" Then
        cbo年龄单位.Tag = cbo年龄单位.Text
    End If
End Sub

Private Function Init费别(bln初诊 As Boolean, Optional blnKeepIndex As Boolean) As Boolean
'参数：bln初诊=是否允许仅限初诊的项目
'      blnKeepIndex=是否保持原有的费别选择
    Dim strSql As String, i As Integer
    Dim strKeep As String
    Dim str缺省费别 As String
    
    On Error GoTo errH
    
    strKeep = cbo费别.Text      '病人以前的费别,有可能现在的系统中已没有该费别了
    If strKeep <> "" Then strKeep = Mid(strKeep, InStr(1, strKeep, "-") + 1)
    
    '72168,冉俊明,2014/4/22,挂号时通过挂号科室确定可选费别
    If mrs费别 Is Nothing Then '首次调用该函数时[bln初诊]为true
        Set mrs费别 = New ADODB.Recordset
        '费别:身份唯一性项目(包含了缺省费别),可以是初诊,不管有效期间及科室
        strSql = "Select a.编码, a.名称, a.简码, Nvl(a.仅限初诊, 0) As 初诊," & _
                "       Nvl(a.缺省标志, 0) As 缺省, Nvl(b.科室id, 0) As 科室id" & _
                " From 费别 A, 费别适用科室 B" & _
                " Where a.名称 = b.费别(+) And a.属性 = 1" & _
                "      And Trunc(Sysdate) Between Nvl(a.有效开始, To_Date('1900-01-01', 'YYYY-MM-DD'))" & _
                "                         And Nvl(a.有效结束, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                "      And Nvl(a.服务对象, 3) In (1, 3)" & _
                " Order By a.编码"
        Call gobjDatabase.OpenRecordset(mrs费别, strSql, Me.Caption)
    End If
    
    If mrs费别 Is Nothing Then Exit Function
    If bln初诊 Then
        mrs费别.Filter = "科室id=0 or 科室id=" & mlng科室ID
    Else                        '不允许仅限初诊的项目
        mrs费别.Filter = "(初诊=0 and 科室id=0) or (初诊=0 and 科室id=" & mlng科室ID & ")"
    End If
    If mrs费别.RecordCount > 0 Then mrs费别.MoveFirst
    
    cbo费别.Clear
    Do While Not mrs费别.EOF
        cbo费别.AddItem mrs费别!编码 & "-" & mrs费别!名称
        '记录初诊项目:不会是本地缺省及系统缺省
        cbo费别.ItemData(cbo费别.NewIndex) = IIf(mrs费别!初诊 = 1, 2, 0)
        
        If str缺省费别 = "" Then    '没有本地缺省时取系统缺省
            If mrs费别!缺省 = 1 Then str缺省费别 = mrs费别!名称
        End If
        mrs费别.MoveNext
    Loop
    
    If blnKeepIndex And strKeep <> "" Then Call gobjControl.CboLocate(cbo费别, strKeep)

    If cbo费别.ListIndex = -1 Then Call gobjControl.CboLocate(cbo费别, str缺省费别)
    
    If cbo费别.ListIndex = -1 Then If cbo费别.ListCount > 0 Then cbo费别.ListIndex = 0
    If cbo费别.ListIndex <> -1 Then cbo费别.ItemData(cbo费别.ListIndex) = 1
            
    Init费别 = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function ShowMe(frmMain As Object, bytFun As Byte, ByVal lng病人ID As Long, ByRef lngOut病人ID As Long, _
                       Optional lng科室id As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:编辑病人信息档案入口
    '入参:frmMain-调用的主窗体
    '     bytFun-0-编辑或查看病人信息;
    '     bln复诊-复诊
    '     lng病人ID-0-新建病人;>0表示编辑和查看病人信息档案
    '出参:lngOut病人ID-返回新建档或修改病人档案的病人ID
    '返回:修改成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-10-27 11:10:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFun = bytFun: Set mfrmMain = frmMain
    txtPatiMCNO(0).ToolTipText = "最大长度" & txtPatiMCNO(0).MaxLength & "位"
    mlng病人ID = lng病人ID
    mlng科室ID = lng科室id
    '初始化卡结算部件
    If gobjSquare Is Nothing Then CreateSquareCardObject Me, glngModul
    Call NewCardObject   '47007
    If txt门诊号.Text <> "" Then mintNOLength = Len(txt门诊号.Text)
    mblnOK = False
    txtPatient.Enabled = True: txt出生日期.Enabled = True: txt出生时间.Enabled = True
    txt年龄.Enabled = True: cbo年龄单位.Enabled = True: cbo性别.Enabled = True
    txt身份证号.Enabled = True
    Call InitFact
    Me.Show 1, frmMain
    ShowMe = mblnOK
    lngOut病人ID = mlng病人ID
    Call CloseIDCard    '47007
End Function
Private Sub InitParaValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化参数值
    '编制:刘兴洪
    '日期:2017-10-27 13:55:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intTemp As Integer
    On Error GoTo errHandle
    With mty_Para
        .bln家庭地址输入 = Val(Nvl(gobjDatabase.GetPara("家庭地址输入方式", glngSys, 1111, 1), 1)) = 1
        .bln自动门诊号 = gobjDatabase.GetPara("自动门诊号", glngSys, 1111) = "1"
        .bln门诊号有效性检查 = Val(Nvl(gobjDatabase.GetPara("门诊号有效性检查", glngSys, 1111, 1), 1)) = 1
        .bln结构化地址录入 = Val(gobjDatabase.GetPara(251, glngSys)) <> 0 '病人地址结构化录入
        .bln乡镇地址结构化 = Val(gobjDatabase.GetPara(252, glngSys)) <> 0 '乡镇地址结构化录入
        
        intTemp = Val(gobjDatabase.GetPara("N岁以下必须录入监护人", glngSys, 1111, 0))
        .bln监护人录入 = intTemp > 0
        .int监护人年龄 = intTemp
    End With
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function InitFact() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化面版信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-10-29 22:48:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
        
    Call InitParaValue
    mblnChange = True
    mblnNewPatient = False
    
    Call InitData
    Call CreateObjectPlugIn  '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    Call CreateObjectKeyboard
    '创建病人信息公共部件
    '69026,冉俊明,2014-8-8,检查输入年龄
    Call CreatePublicPatient
    
    Call InitTagPage
    Call InitTaskPanelOther
    If mbytFun = 0 And mlng病人ID <> 0 Then
        Me.Caption = "病人详细信息"
        Call LoadPatientInfo
    ElseIf mbytFun = 2 Then
        Me.Caption = "新增病人信息"
    End If
    
    If mty_Para.bln自动门诊号 And txt门诊号.Text = "" Then txt门诊号.Text = gobjDatabase.GetNextNo(3)
    
    '初始化地址控件
    If mty_Para.bln结构化地址录入 Then
        padd家庭地址.MaxLength = glngMax家庭地址: padd户口地址.MaxLength = glngMax户口地址
        padd家庭地址.Visible = True: padd户口地址.Visible = True
        padd家庭地址.ShowTown = mty_Para.bln乡镇地址结构化: padd户口地址.ShowTown = mty_Para.bln乡镇地址结构化
        cbo家庭地址.Visible = False: cmd家庭地址.Visible = False
        padd家庭地址.Top = cbo家庭地址.Top: padd家庭地址.Left = cbo家庭地址.Left
        txtRegLocation.Visible = False: cmdRegLocation.Visible = False
        padd户口地址.Top = txtRegLocation.Top: padd户口地址.Left = txtRegLocation.Left
    End If
    txtRegLocation.MaxLength = glngMax户口地址
    txtBirthLocation.MaxLength = glngMax出生地点
    InitFact = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Sub LoadPatientInfo()
    Dim rsTmp As ADODB.Recordset, strSql As String
    strSql = " Select 姓名,门诊号,出生日期,年龄,性别,身份证号,医保号,家庭地址,家庭电话,家庭地址邮编,户口地址,户口地址邮编," & _
             "        民族,国籍,职业,婚姻状况,费别,医疗付款方式,监护人,联系人姓名,联系人电话,联系人关系,单位邮编,工作单位," & _
             "        单位电话,出生地点,区域,手机号,病人ID " & _
             " From 病人信息 Where 病人ID = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, mlng病人ID)
    If rsTmp.EOF Then
        MsgBox "找不到病人信息,请检查输入的病人信息是否正确!", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    Else
        txtPatient.Text = Nvl(rsTmp!姓名)
        txtPatient.Locked = True
        txtPatient.Enabled = False
        txtMobile.Text = Nvl(rsTmp!手机号)
        txt门诊号.Text = Nvl(rsTmp!门诊号)
        If InStr(gstrPrivs, ";允许修改门诊号;") = 0 Then
            txt门诊号.Locked = True
            txt门诊号.Enabled = False
        End If
        txt年龄.Text = Nvl(rsTmp!年龄)
        txt年龄.Tag = Nvl(rsTmp!年龄)
        txt年龄.Locked = True
        txt年龄.Enabled = False
        cbo年龄单位.Visible = False
        cbo年龄单位.Enabled = False
        txt年龄.Width = 1395
        If IsDate(Nvl(rsTmp!出生日期)) Then
            txt出生日期.Text = Format(Nvl(rsTmp!出生日期), "YYYY-MM-DD")
            txt出生时间.Text = Format(Nvl(rsTmp!出生日期), "HH:MM")
        End If
        txt出生日期.Enabled = False
        txt出生时间.Enabled = False
        Call gobjControl.Cbo.Locate(cbo性别, Nvl(rsTmp!性别))
        cbo性别.Locked = True
        cbo性别.Enabled = False
        txt身份证号.Text = Nvl(rsTmp!身份证号)
        txt身份证号.Locked = True
        txt身份证号.Enabled = False
        txtPatiMCNO(0).Text = Nvl(rsTmp!医保号)
        txtPatiMCNO(1).Text = Nvl(rsTmp!医保号)
        cbo家庭地址.Text = Nvl(rsTmp!家庭地址)
        Call zlReadAddrInfo(padd家庭地址, Val(Nvl(rsTmp!病人ID)), 0, 3, cbo家庭地址.Text)
        
        txt家庭电话.Text = Nvl(rsTmp!家庭电话)
        txt家庭邮编.Text = Nvl(rsTmp!家庭地址邮编)
        txtRegLocation.Text = Nvl(rsTmp!户口地址)
        Call zlReadAddrInfo(padd户口地址, Val(Nvl(rsTmp!病人ID)), 0, 4, txtRegLocation.Text)
        
        txt户口地址邮编.Text = Nvl(rsTmp!户口地址邮编)
        Call gobjControl.Cbo.Locate(cbo民族, Nvl(rsTmp!民族))
        Call gobjControl.Cbo.Locate(cbo国籍, Nvl(rsTmp!国籍))
        Call gobjControl.Cbo.Locate(cbo职业, Nvl(rsTmp!职业))
        If cbo职业.Text = "" Then
            cbo职业.AddItem Nvl(rsTmp!职业)
            cbo职业.ListIndex = cbo职业.NewIndex
        End If
        Call gobjControl.Cbo.Locate(cbo婚姻, Nvl(rsTmp!婚姻状况))
        If cbo婚姻.Text = "" Then
            cbo婚姻.AddItem Nvl(rsTmp!婚姻状况)
            cbo婚姻.ListIndex = cbo婚姻.NewIndex
        End If
'        cbo婚姻.Locked = True
        Call gobjControl.Cbo.Locate(cbo费别, Nvl(rsTmp!费别))
        If InStr(gstrPrivs, ";允许修改费别;") = 0 Then
            cbo费别.Enabled = False
        Else
            cbo费别.Enabled = True
        End If
        Call gobjControl.Cbo.Locate(cbo付款方式, Nvl(rsTmp!医疗付款方式))
        If cbo付款方式.Text = "" Then
            cbo付款方式.AddItem Nvl(rsTmp!医疗付款方式)
            cbo付款方式.ListIndex = cbo付款方式.NewIndex
        End If
'        cbo付款方式.Locked = True
        txt监护人.Text = Nvl(rsTmp!监护人)
'        txt监护人.Locked = True
        txt联系人姓名.Text = Nvl(rsTmp!联系人姓名)
'        txt联系人姓名.Locked = True
        txt联系人电话.Text = Nvl(rsTmp!联系人电话)
'        txt联系人电话.Locked = True
        txt联系人身份证.Text = ""
'        txt联系人身份证.Locked = True
        Call gobjControl.Cbo.Locate(cbo联系人关系, Nvl(rsTmp!联系人关系))
'        cbo联系人关系.Locked = True
        txt单位邮编.Text = Nvl(rsTmp!单位邮编)
'        txt单位邮编.Locked = True
        txt单位名称.Text = Nvl(rsTmp!工作单位)
'        txt单位名称.Locked = True
        txt单位电话.Text = Nvl(rsTmp!单位电话)
'        txt单位电话.Locked = True
        txtBirthLocation.Text = Nvl(rsTmp!出生地点)
'        txtBirthLocation.Locked = True
        txt区域.Text = Nvl(rsTmp!区域)
'        txt区域.Locked = True
'        cmdOK.Visible = False
        cmd家庭地址.Visible = True
        cmdRegLocation.Visible = True
        cmd单位名称.Visible = True
        cmdBirthLocation.Visible = True
        cmd区域.Visible = True
        cmdMedicalWarning.Visible = True
        Call Load健康卡相关信息(mlng病人ID)
        Call ReadPatPricture(mlng病人ID)
        Call zlQueryEMPIPatiInfo
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mblnChange = False
    Set mobjKeyboard = Nothing
    Call CloseIDCard
    '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    If Not mobjPlugIn Is Nothing Then Set mobjPlugIn = Nothing
    mblnPlugin = False
    mlngPlugInHwnd = 0: mblnSavePati = False
    '74430,冉俊明,2014-7-7,挂号中的病人信息编辑功能中提供采集照片功能
    mlng图像操作 = 0: mstr采集图片 = ""
    If Not mobjPubPatient Is Nothing Then Set mobjPubPatient = Nothing
    mblnGetBirth = False
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
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
            Call gobjControl.TxtSelAll(txt过敏)
            txt过敏.Visible = True
            If txt过敏.Visible Then txt过敏.SetFocus
        Case 1 '过敏反应
            txt过敏反应.Top = msh过敏.CellTop + msh过敏.Top + (msh过敏.CellHeight - txt过敏反应.Height) / 2 - 15
            txt过敏反应.Left = msh过敏.Left + msh过敏.CellLeft + 30
            '75446:李南春,2014-7-16,过敏反应文本框不够
            txt过敏反应.Width = msh过敏.CellWidth - 60
            
            txt过敏反应.Text = msh过敏.TextMatrix(msh过敏.Row, msh过敏.Col)
            txt过敏反应.ZOrder
            Call gobjControl.TxtSelAll(txt过敏反应)
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
    txt过敏反应.Visible = False
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
    Call gobjControl.TxtSelAll(txtBirthLocation)
    Call gobjCommFun.OpenIme(True)
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
    Dim strSql As String, strWhere As String
    Dim strKey As String, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset, vRect As RECT
    
    On Error GoTo Errhand
    If strInput <> "" And txtInput.Tag <> "" Then Exit Sub
'    vRect = gobjControl.GetControlRect(txtInput.hWnd)
    If strInput = "" Then '点击按钮
        strSql = "" & _
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
        Set rsTemp = gobjDatabase.ShowSQLSelect(Me, strSql, 2, "地区", False, _
                       "", "", False, False, False, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False)
    Else
        '去掉"'"
        strInput = Replace(strInput, "'", " ")
        strKey = GetMatchingSting(strInput, False)
        If strInput <> "" Then
            If IsNumeric(strInput) Then '输入全是数字时只匹配编码
                strWhere = " Where 编码 Like Upper([1])"
            ElseIf gobjCommFun.IsCharAlpha(strInput) Then '输入全是字母时只匹配简码
                strWhere = " Where 简码 Like Upper([1])"
            Else
                strWhere = " Where 编码 Like Upper([1]) Or 名称 Like [1] Or 简码 Like Upper([1])"
            End If
        End If
        
        strSql = "" & _
            "Select Rownum As ID, 编码, 名称 From 地区 " & strWhere & " Order By 编码"
        Set rsTemp = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "地区", False, _
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
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub txtBirthLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    '73022,冉俊明,2014-5-20,在单位名称、出生地点、户口地址加上模糊查找功能
    If KeyCode = vbKeyReturn And Trim(txtBirthLocation.Text) <> "" Then
        Call SearchAddress(Trim(txtBirthLocation.Text), txtBirthLocation)
    End If
End Sub

Private Sub txtBirthLocation_LostFocus()
    Call gobjCommFun.OpenIme(False)
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If mblnNameChange = True And mlng病人ID = 0 Then zlQueryEMPIPatiInfo
    mblnNameChange = False
End Sub

Private Sub txtPatiMCNO_Change(Index As Integer)
    mstrPlugChange = mstrPlugChange & ",医保号"
End Sub

Private Sub txtRegLocation_Change()
    mstrPlugChange = mstrPlugChange & ",户口地址"
    txtRegLocation.Tag = ""
End Sub

Private Sub txtRegLocation_GotFocus()
    Call gobjControl.TxtSelAll(txtRegLocation)
    Call gobjCommFun.OpenIme(True)
End Sub

Private Sub txtRegLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    '73022,冉俊明,2014-5-20,在单位名称、出生地点、户口地址加上模糊查找功能
    If KeyCode = vbKeyReturn And Trim(txtRegLocation.Text) <> "" Then
        Call SearchAddress(Trim(txtRegLocation.Text), txtRegLocation)
    End If
End Sub

Private Sub txtMobile_Validate(Cancel As Boolean)
    If Exist手机号(txtMobile.Text, IIf(mlng病人ID <> 0, mlng病人ID, 0)) Then
        If MsgBox("输入的手机号与其他病人重复，是否确定录入？", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Cancel = True
    End If
End Sub

Private Sub txtRegLocation_LostFocus()
    Call gobjCommFun.OpenIme(False)
End Sub

Private Sub txtPatient_Change()
    If mobjIDCard Is Nothing And Visible Then Exit Sub
    If Not mobjIDCard Is Nothing And Not txtPatient.Locked Then mobjIDCard.SetEnabled (txtPatient.Text = "")
End Sub

Private Sub txtPatiMCNO_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call gobjCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtPatiMCNO_Validate(Index As Integer, Cancel As Boolean)
    txtPatiMCNO(Index).Text = UCase(Trim(txtPatiMCNO(Index).Text))
    If cbo付款方式.ListCount > 0 Then cbo付款方式.ListIndex = 0

    If Index = 1 Then
        If txtPatiMCNO(1).Text <> txtPatiMCNO(0).Text Then
            MsgBox "请检查,两次输入的医保号不一致！", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub txt出生日期_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If txt出生日期.Text = "____-__-__" Then
           gobjCommFun.PressKey (vbKeyTab) '跳过时间
           gobjCommFun.PressKey (vbKeyTab)
       Else
           gobjCommFun.PressKey (vbKeyTab)
       End If
    End If

End Sub

Private Sub txt出生时间_Change()
    Dim str出生时间 As String
    '76669，李南春,2014-8-18,病人年龄更新
    If IsDate(txt出生日期.Text) Then
        str出生时间 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        txt年龄.Text = ReCalcOld(CDate(str出生时间), cbo年龄单位)
        If cbo年龄单位.Visible Then
            txt年龄.Width = 690
        Else
            txt年龄.Width = 1395
        End If
        txt年龄.Tag = txt年龄.Text
    End If
End Sub

Private Sub txt出生时间_GotFocus()
    gobjControl.TxtSelAll txt出生时间
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

Private Function ReCalcBirth(ByVal strOld As String, ByVal str年龄单位 As String) As String
'功能:根据年龄和年龄单位估算病人的出生日期,年龄单位为岁时,出年月日假定为1月1号,年龄单位为月时,出生日期假定为1号
'返回:出生日期
    Dim strTmp As String, strFormat As String, lngDays As Long
    
    strTmp = "____-__-__"
    If str年龄单位 = "" Then
        strFormat = "YYYY-MM-DD"
        If strOld Like "*岁*月" Or strOld Like "*岁*个月" Then
            strFormat = "YYYY-MM-01"
            lngDays = 365 * Val(strOld) + 30 * Val(Mid(strOld, InStr(1, strOld, "岁") + 1))
        ElseIf strOld Like "*月*天" Or strOld Like "*个月*天" Then
            lngDays = 30 * Val(strOld) + Val(Mid(strOld, InStr(1, strOld, "月") + 1))
        ElseIf strOld Like "*岁" Or IsNumeric(strOld) Then
            strFormat = "YYYY-01-01"
            lngDays = 365 * Val(strOld)
        ElseIf strOld Like "*月" Or strOld Like "*个月" Then
            strFormat = "YYYY-MM-01"
            lngDays = 30 * Val(strOld)
        ElseIf strOld Like "*天" Then
            lngDays = Val(strOld)
        End If
        If lngDays <> 0 Then strTmp = Format(DateAdd("d", lngDays * -1, gobjDatabase.Currentdate), strFormat)
    ElseIf strOld <> "" Then
        Select Case str年龄单位
            Case "岁"
                If Val(strOld) > 200 Then lngDays = -1
            Case "月"
                If Val(strOld) > 2400 Then lngDays = -1
            Case "天"
                If Val(strOld) > 73000 Then lngDays = -1
        End Select
        
        If lngDays = 0 Then
            strTmp = Switch(str年龄单位 = "岁", "yyyy", str年龄单位 = "月", "m", str年龄单位 = "天", "d")
            strTmp = Format(DateAdd(strTmp, Val(strOld) * -1, gobjDatabase.Currentdate), "YYYY-MM-DD")
            
            If str年龄单位 = "岁" Then
                strTmp = Format(strTmp, "YYYY-01-01")
            ElseIf str年龄单位 = "月" Then
                strTmp = Format(strTmp, "YYYY-MM-01")
            End If
        End If
    End If
    ReCalcBirth = strTmp
End Function

Private Function CheckOldData(ByRef txt年龄 As TextBox, ByRef cbo年龄单位 As ComboBox) As Boolean
'功能：检查年龄输入值的有效性
'返回：
    If Not IsNumeric(txt年龄.Text) Then CheckOldData = True: Exit Function
    
    Select Case cbo年龄单位.Text
        Case "岁"
            If Val(txt年龄.Text) > 200 Then
                MsgBox "年龄不能大于200岁!", vbInformation, gstrSysName
                If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "月"
            If Val(txt年龄.Text) > 2400 Then
                MsgBox "年龄不能大于2400月!", vbInformation, gstrSysName
                If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "天"
            If Val(txt年龄.Text) > 73000 Then
                MsgBox "年龄不能大于73000天!", vbInformation, gstrSysName
                If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                CheckOldData = False: Exit Function
            End If
    End Select
    CheckOldData = True
End Function

Private Sub txt出生日期_Change()
    Dim str出生时间 As String
    If IsDate(txt出生日期.Text) And mblnChange Then
        mblnChange = False
        txt出生日期.Text = Format(CDate(txt出生日期.Text), "yyyy-mm-dd") '0002-02-02自动转换为2002-02-02,否则,看到的是2002,实际值却是0002
        mblnChange = True
        
        str出生时间 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        txt年龄.Text = ReCalcOld(CDate(str出生时间), cbo年龄单位)
        If cbo年龄单位.Visible Then
            txt年龄.Width = 690
        Else
            txt年龄.Width = 1395
        End If
        txt年龄.Tag = txt年龄.Text
        cbo年龄单位.Tag = cbo年龄单位.Text
        mblnGetBirth = False
    End If
End Sub
Private Sub txt出生日期_GotFocus()
    gobjControl.TxtSelAll txt出生日期
End Sub

Private Sub txt出生日期_LostFocus()
    If txt出生日期.Text <> "____-__-__" And Not IsDate(txt出生日期.Text) Then
      If txt出生日期.Enabled And txt出生日期.Visible Then txt出生日期.SetFocus
    End If
End Sub


Private Sub txt单位电话_GotFocus()
    Call gobjControl.TxtSelAll(txt单位电话)
End Sub

Private Sub txt单位电话_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt单位电话, KeyAscii
End Sub

Private Sub txt单位名称_Change()
    txt单位名称.Tag = ""
End Sub

Private Sub txt单位名称_GotFocus()
    Call gobjControl.TxtSelAll(txt单位名称)
    Call gobjCommFun.OpenIme(True)
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
    Dim strSql As String, strWhere As String
    Dim strKey As String, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset, vRect As RECT
    
    On Error GoTo Errhand
    If strInput <> "" And txtInput.Tag <> "" Then Exit Sub
'    vRect = gobjControl.GetControlRect(txtInput.hWnd)
    If strInput = "" Then '点击按钮
        strSql = "" & _
        "       Select ID,上级ID,末级,编码,名称,地址,电话,开户银行,帐号,联系人 From  合约单位" & _
        "       Where 撤档时间 Is Null Or 撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD')" & _
        "       Start With 上级ID is NULL" & _
        "       Connect by Prior ID=上级ID"
        '75888,冉俊明,2014-7-28
        Set rsTemp = gobjDatabase.ShowSQLSelect(Me, strSql, 2, "单位", False, _
                       "", "", False, True, False, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False)
    Else
        '去掉"'"
        strInput = Replace(strInput, "'", " ")
        strKey = GetMatchingSting(strInput, False)
        If strInput <> "" Then
            If IsNumeric(strInput) Then '输入全是数字时只匹配编码
                strWhere = " Where 编码 Like Upper([1])"
            ElseIf gobjCommFun.IsCharAlpha(strInput) Then '输入全是字母时只匹配简码
                strWhere = " Where 简码 Like Upper([1])"
            Else
                strWhere = " Where 编码 Like Upper([1]) Or 名称 Like [1] Or 简码 Like Upper([1])"
            End If
        End If
        
        strSql = "" & _
        "       Select ID,上级ID,末级,编码,名称,地址,电话,开户银行,帐号,联系人 From  合约单位" & strWhere & _
        "       And (撤档时间 Is Null Or 撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD'))"
        Set rsTemp = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "单位", False, _
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
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub txt单位名称_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt单位名称, KeyAscii
End Sub

Private Sub txt单位名称_LostFocus()
    Call gobjCommFun.OpenIme
End Sub

Private Sub txt单位邮编_GotFocus()
    Call gobjControl.TxtSelAll(txt单位邮编)
End Sub

Private Sub txt单位邮编_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt单位邮编, KeyAscii
End Sub

Private Sub txt过敏_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim ObjItem As ListItem
    Dim strSql As String
            
    If KeyAscii <> 13 Then
        If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    Else
        KeyAscii = 0
        '75286:李南春，2014-7-16，自由录入过敏药物
        msh过敏.TextMatrix(msh过敏.Row, 0) = txt过敏.Text '问题号:56599

        strSql = " Select Distinct A.ID,A.编码," & _
        " A.名称,A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类," & _
        " Decode(B.是否新药,1,'√','') as 新药,Decode(B.是否皮试,1,'√','') as 皮试" & _
        " From 诊疗项目目录 A,药品特性 B,诊疗项目别名 C" & _
        " Where A.类别 IN('5','6','7') And A.ID=B.药名ID And A.Id=C.诊疗项目id" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
        " And (C.名称 like [1] OR A.编码 like [1] OR C.简码 like [1])"
        
        On Error GoTo errH
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(txt过敏.Text) & "%")
        
        With rsTmp
            If .BOF Or .EOF Then
                msh过敏.SetFocus: msh过敏_EnterCell
                Exit Sub
            Else
                Me.lvwItems.ListItems.Clear
                Do While Not .EOF
                    Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称, , IIf(!皮试 <> "", 1, 2))
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("剂型").Index - 1) = IIf(IsNull(!剂型), "", !剂型)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("毒理分类").Index - 1) = IIf(IsNull(!毒理分类), "", !毒理分类)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("新药").Index - 1) = IIf(IsNull(!新药), "", !新药)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("皮试").Index - 1) = IIf(IsNull(!皮试), "", !皮试)
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
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
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
    mblnChange = True
End Sub

Private Sub txt户口地址邮编_GotFocus()
    Call gobjControl.TxtSelAll(txt户口地址邮编)
End Sub

Private Sub txt户口地址邮编_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt家庭电话_GotFocus()
    Call gobjControl.TxtSelAll(txt家庭电话)
End Sub

Private Sub txt家庭电话_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt家庭电话, KeyAscii
End Sub

Private Sub txt家庭邮编_GotFocus()
    Call gobjControl.TxtSelAll(txt家庭邮编)
End Sub

Private Sub txt家庭邮编_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt家庭邮编, KeyAscii
End Sub
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

Private Sub txt联系人姓名_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("姓名") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("姓名")) = txt联系人姓名.Text
        If vsLinkMan.Rows = vsLinkMan.FixedRows + 1 And txt联系人姓名.Text <> "" Then
            vsLinkMan.Rows = vsLinkMan.Rows + 1
        End If
    End If
End Sub

Private Sub txt门诊号_GotFocus()
    Call gobjControl.TxtSelAll(txt门诊号)
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
        Call gobjCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = 32 Then
        KeyAscii = 0
        If txt门诊号.Text = "" Then
            txt门诊号.Text = gobjDatabase.GetNextNo(3)
            mintNOLength = Len(Trim(txt门诊号.Text))
        End If
        Call gobjCommFun.PressKey(vbKeyTab)
    ElseIf InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Or InStr(gstrPrivs, ";允许修改门诊号;") = 0 Then
        KeyAscii = 0
    End If
End Sub
 
Private Sub txt年龄_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt年龄.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt年龄.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt年龄_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt年龄.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt年龄_GotFocus()
    Call gobjCommFun.OpenIme
    Call gobjControl.TxtSelAll(txt年龄)
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    Dim blnTab As Boolean
    
    If KeyAscii = vbKeyReturn Then
        If cbo年龄单位.Visible = False And IsNumeric(txt年龄.Text) Then
            Call txt年龄_Validate(False)
            Call cbo年龄单位.SetFocus
        Else
            Call gobjCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt年龄.Text) And cbo年龄单位.Visible Then Call gobjCommFun.PressKey(vbKeyTab)
    Else
        '仅仅限制几个 指定的特殊的字符 问题:49908
        If InStr("~·！@#￥%……&*（）——-+=|、？、。，~`!#$%^&*()-_=+|\/?<>,/<>", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt年龄_Validate(Cancel As Boolean)
    Dim strBirth As String
    txt年龄.Text = Trim(txt年龄.Text)
    If Not IsNumeric(txt年龄.Text) And Trim(txt年龄.Text) <> "" Then
        cbo年龄单位.ListIndex = -1: cbo年龄单位.Visible = False: txt年龄.Width = 1395
    ElseIf cbo年龄单位.Visible = False Then
        cbo年龄单位.ListIndex = 0: cbo年龄单位.Visible = True: txt年龄.Width = 690
    End If
    If txt年龄.Text <> txt年龄.Tag Then
        mblnChange = False
        If Not IsDate(txt出生日期.Text) Then mblnGetBirth = True
        '125451：控制是否允许通过年龄计算出生日期
        If mblnGetBirth Then
            If mobjPubPatient.ReCalcBirthDay(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), strBirth) Then
                txt出生日期.Text = Format(strBirth, "YYYY-MM-DD")
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
    Call gobjControl.TxtSelAll(txt单位名称)
    Call gobjCommFun.OpenIme(True)
End Sub

Private Sub txt其他关系_LostFocus()
    Call gobjCommFun.OpenIme
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

Private Sub txt区域_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If (txt区域.Tag <> "" Or Trim(txt区域.Text) = "") Then gobjCommFun.PressKey vbKeyTab: Exit Sub
    If zl_SelectAndNotAddItem(Me, txt区域, Trim(txt区域.Text), "区域", "区域选择", True, False) = False Then
        Exit Sub
    End If
End Sub

Private Sub txt身份证号_Change()
    If mbln扫描身份证签约 And ActiveControl Is txt身份证号 And Not mobjIDCard Is Nothing Then
            mobjIDCard.SetEnabled txt身份证号.Text = ""
    End If
End Sub

Private Sub txt身份证号_GotFocus()
    Call gobjControl.TxtSelAll(txt身份证号)

    If mbln扫描身份证签约 = True And txt身份证号.Text = "" Then
        OpenIDCard
    End If
End Sub
Private Sub txt身份证号_KeyPress(KeyAscii As Integer)
    
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
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPatient_GotFocus()
    Call gobjControl.TxtSelAll(txtPatient)
    Call gobjCommFun.OpenIme(True)
    
    If mobjIDCard Is Nothing And Visible Then Call NewCardObject
    If mobjIDCard Is Nothing And Visible Then Exit Sub
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjIDCard.SetEnabled (True)
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        '新病人才调用
        If mblnNameChange = True And mlng病人ID = 0 Then zlQueryEMPIPatiInfo
        mblnNameChange = False
        Call gobjCommFun.PressKey(vbKeyTab)
    Else
        mblnNameChange = True
    End If
    CheckLen txtPatient, KeyAscii
End Sub

Public Sub CheckLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If gobjCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Private Sub txtPatient_LostFocus()
    Call gobjCommFun.OpenIme
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txt身份证号_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
End Sub

Private Sub txt身份证号_Validate(Cancel As Boolean)
    '65663:刘尔旋,2014-02-20,根据身份证号计算出生日期
    If IsDate(gobjCommFun.GetIDCardDate(txt身份证号.Text)) = False Then Exit Sub
    If Format(gobjCommFun.GetIDCardDate(txt身份证号.Text), "yyyy-mm-dd") <> Format(txt出生日期.Text, "yyyy-mm-dd") Then
        If IsDate(txt出生日期.Text) Then MsgBox "输入的身份证号与输入的出生日期不一致，将使用身份证号获取的日期替换！", vbInformation, gstrSysName
        txt出生日期.Text = gobjCommFun.GetIDCardDate(txt身份证号.Text)
    End If
End Sub

Private Function GetOneDept(lng收费细目ID As Long) As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select B.执行科室ID From 收费项目目录 A,收费执行科室 B Where B.收费细目ID=A.ID And A.ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng收费细目ID)
    If Not rsTmp.EOF Then
        GetOneDept = rsTmp!执行科室ID '默认取第一个(如有多个)
    Else
        GetOneDept = UserInfo.部门ID '如没有指定，则取操作员所在科室
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
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
    If gobjComlib.ErrCenter() = 1 Then
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
    If gobjComlib.ErrCenter() = 1 Then Resume
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
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function
Private Function zlComboxLoadFromSQL(ByVal strSql As String, cboControl As Variant, Optional ByVal blnID As Boolean = False) As Boolean
'本函数的功能是从数据库中读出列表值并装到下拉框中
    Dim rsTemp As New ADODB.Recordset
    Dim intCount As Long
    Dim cmbArray As Variant
    
    On Error GoTo errHandle
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取Cbo数据")
    '下拉框数组
    If IsArray(cboControl) Then
        cmbArray = cboControl
    Else
        '强行组成一个数组
        cmbArray = Array(cboControl)
    End If
    
    For intCount = LBound(cmbArray) To UBound(cmbArray)
        cmbArray(intCount).Clear
        Do Until rsTemp.EOF
            If IsNull(rsTemp("编码")) Then
                cmbArray(intCount).AddItem rsTemp.AbsolutePosition & "." & rsTemp("名称")
            Else
                cmbArray(intCount).AddItem rsTemp("编码") & "." & rsTemp("名称")
            End If
            If blnID = True Then cmbArray(intCount).ItemData(cmbArray(intCount).NewIndex) = rsTemp("ID")
            If rsTemp("缺省标志") = 1 Then
                cmbArray(intCount).ListIndex = cmbArray(intCount).NewIndex
                cmbArray(intCount).ItemData(cmbArray(intCount).NewIndex) = 1
            End If
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
        If blnID = True Then cmbArray(intCount).ListIndex = 0
    Next
    
    zlComboxLoadFromSQL = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    zlComboxLoadFromSQL = False
End Function

Private Function GetCardDataSql(ByVal byt变动类型 As Byte, ByVal lng病人ID As Long, ByVal lng卡类别ID As Long, _
   ByVal str原卡号 As String, ByVal strCard As String, ByVal str密码 As String, ByVal dtCurDate As Date, _
   ByVal strICCard As String, Optional ByVal str变动原因 As String = "")
    Dim strSql As String
    Dim strPassWord As String
    strPassWord = gobjCommFun.zlStringEncode(str密码)
    'Zl_医疗卡变动_Insert
     strSql = "Zl_医疗卡变动_Insert("
    '      变动类型_In   Number,
    '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定);
    '５-密码调整(只记录);6-挂失(16取消挂失)
    strSql = strSql & "" & byt变动类型 & ","
    '      病人id_In     住院费用记录.病人id%Type,
    strSql = strSql & "" & lng病人ID & ","
    '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
    strSql = strSql & "" & lng卡类别ID & ","
    '      原卡号_In     病人医疗卡信息.卡号%Type,
    strSql = strSql & "'" & str原卡号 & "',"
    '      医疗卡号_In   病人医疗卡信息.卡号%Type,
    strSql = strSql & "'" & strCard & "',"
    '      变动原因_In   病人医疗卡变动.变动原因%Type,
    '      --变动原因_In:如果密码调整，变动原因为密码.加密的
    strSql = strSql & "'" & str变动原因 & "',"
    '      密码_In       病人信息.卡验证码%Type,
    strSql = strSql & "'" & strPassWord & "',"
    '      操作员姓名_In 住院费用记录.操作员姓名%Type,
    strSql = strSql & "'" & UserInfo.姓名 & "',"
    '      变动时间_In   住院费用记录.登记时间%Type,
    strSql = strSql & "to_date('" & Format(dtCurDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
    strSql = strSql & "'" & strICCard & "',"
    '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
    strSql = strSql & IIf(str变动原因 = "", "NULL)", "'" & str变动原因 & "')")
    GetCardDataSql = strSql
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
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    If Not mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.setParaent(Me.hWnd)
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
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    '打开读卡器
    mobjIDCard.SetEnabled (True)
End Sub

Private Function zl_Get缺省发卡类别() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取缺省发卡类别
    '返回:缺省发卡类别名称
    '编制:王吉
    '日期:2012-08-31 11:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim lngCardTypeID As Long
    Dim rsTemp As Recordset
    
    On Error GoTo ErrHandl:

    strSql = "" & _
    "   Select Id, 编码, 名称, 短名, 前缀文本, 卡号长度, 缺省标志, 是否固定, 是否严格控制, " & _
    "           nvl(是否自制,0) as 是否自制, nvl(是否存在帐户,0) as 是否存在帐户, " & _
    "           nvl(是否全退,0) as 是否全退,nvl(是否重复使用,0) as 是否重复使用 , " & _
    "           nvl(密码长度,10) as 密码长度,nvl(密码长度限制,0) as 密码长度限制,nvl(密码规则,0) as 密码规则," & _
    "           nvl(是否退现,0) as 是否退现,部件, 备注, 特定项目, 结算方式, 是否启用, 卡号密文,Nvl(密码输入限制,0) as 密码输入限制,Nvl(是否缺省密码,0) as 是否缺省密码," & _
    "           nvl(是否模糊查找,0) as 是否模糊查找,nvl(读卡性质,'1000') as 读卡性质 " & _
    "    From 医疗卡类别" & _
    "    Where ID = [1]" & _
    "    Order by 编码"

    lngCardTypeID = Val(gobjDatabase.GetPara("缺省医疗卡类别", glngSys, glngModul, , , True))
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lngCardTypeID)
    If rsTemp Is Nothing Then zl_Get缺省发卡类别 = "": Exit Function
    If rsTemp.RecordCount <= 0 Then zl_Get缺省发卡类别 = "": Exit Function
    zl_Get缺省发卡类别 = rsTemp!名称
    Exit Function
ErrHandl:
    If gobjComlib.ErrCenter() = 1 Then Resume
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
    
    tbcPage.Top = ScaleTop + 50
    tbcPage.Height = Me.ScaleHeight - tbcPage.Top - (Me.ScaleHeight - cmdHelp.Top + 45)
       
    lblPatiMCNO(0).Enabled = False: lblPatiMCNO(1).Enabled = False
    txtPatiMCNO(0).Enabled = False: txtPatiMCNO(1).Enabled = False
     
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
    Call gobjControl.TxtSelAll(txt验证密码)
    Call OpenPassKeyboard(txt验证密码, False)
End Sub

Private Sub txt验证密码_LostFocus()
    Call ClosePassKeyboard(txt验证密码)
End Sub
Private Sub txt支付密码_GotFocus()
    Call gobjControl.TxtSelAll(txt支付密码)
    Call OpenPassKeyboard(txt支付密码, False)
End Sub

Private Sub LoadOldData(ByVal strOld As String, ByRef txt年龄 As TextBox, ByRef cbo年龄单位 As ComboBox)
'功能:将数据库中保存的年龄按规范的格式加载到界面,不规范的原样显示
    Call gobjControl.LoadOldData(strOld, txt年龄, cbo年龄单位)
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
    If Nvl(rsPatiInfo!性别) <> "" Then
        Call gobjControl.CboLocate(cbo性别, rsPatiInfo!性别)
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
        If cbo年龄单位.Visible Then
            txt年龄.Width = 690
        Else
            txt年龄.Width = 1395
        End If
        txt年龄.Tag = txt年龄.Text
    Else
         txt出生时间.Text = "__:__"
         txt出生日期.Text = ReCalcBirth(Val(txt年龄.Text), cbo年龄单位.Text)
    End If

    '身份证号
    If Nvl(rsPatiInfo!身份证号) <> "" Then
        txt身份证号.Text = rsPatiInfo!身份证号
        If InStr(1, txt出生日期.Text, "__") > 0 Then
            strTmp = gobjCommFun.GetIDCardDate(txt身份证号.Text)
            If IsDate(strTmp) Then txt出生日期.Text = strTmp
        End If
    End If
    '职业
    If Nvl(rsPatiInfo!职业) <> "" Then
        cbo职业.ListIndex = gobjControl.Cbo.FindIndex(cbo职业, rsPatiInfo!职业)
        If cbo职业.ListIndex = -1 Then
            cbo职业.AddItem rsPatiInfo!职业, 0
            cbo职业.ListIndex = cbo职业.NewIndex
        End If
    End If
    '民族
    cbo民族.ListIndex = gobjControl.Cbo.FindIndex(cbo民族, Nvl(rsPatiInfo!民族), True)
     If cbo民族.ListIndex = -1 And Nvl(rsPatiInfo!民族) <> "" Then
         cbo民族.AddItem rsPatiInfo!民族, 0
         cbo民族.ListIndex = cbo民族.NewIndex
     End If
    '国籍
    cbo国籍.ListIndex = gobjControl.Cbo.FindIndex(cbo国籍, Nvl(rsPatiInfo!国籍), True)
     If cbo国籍.ListIndex = -1 And Nvl(rsPatiInfo!国籍) <> "" Then
         cbo国籍.AddItem rsPatiInfo!国籍, 0
         cbo国籍.ListIndex = cbo国籍.NewIndex
     End If
    '婚姻状况
    cbo婚姻.ListIndex = gobjControl.Cbo.FindIndex(cbo婚姻, Nvl(rsPatiInfo!婚姻状况), True)
     If cbo婚姻.ListIndex = -1 And Nvl(rsPatiInfo!婚姻状况) <> "" Then
         cbo婚姻.AddItem rsPatiInfo!婚姻状况, 0
         cbo婚姻.ListIndex = cbo婚姻.NewIndex
     End If
    txt区域.Text = Nvl(rsPatiInfo!区域)
    '家庭地址
    cbo家庭地址.Text = Nvl(rsPatiInfo!家庭地址)
    '家庭电话
    txt家庭电话.Text = Nvl(rsPatiInfo!家庭电话)
    '家庭地址邮编
    txt家庭邮编.Text = Nvl(rsPatiInfo!家庭地址邮编)
    '户口地址
    txtRegLocation.Text = Nvl(rsPatiInfo!户口地址)
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
    cbo联系人关系.ListIndex = gobjControl.Cbo.FindIndex(cbo联系人关系, Nvl(rsPatiInfo!联系人关系), True)
    If cbo联系人关系.ListIndex = -1 And Nvl(rsPatiInfo!联系人关系) <> "" Then
        cbo联系人关系.ListIndex = 8: txt其他关系.Text = Nvl(rsPatiInfo!联系人关系)
    End If
    '问题号:56599
    Load健康卡相关信息 (Val(Nvl(rsPatiInfo!病人ID, "0")))
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
    Dim strSql As String
    Dim rsTemp As Recordset
    On Error GoTo Errhand
    strSql = "" & _
    " Select  姓名,门诊号 From 病人信息 A,病人医疗卡信息 B Where A.病人ID=B.病人ID And B.卡号=[1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "医疗卡绑定", str身份证号)
    If rsTemp Is Nothing Then zl当前用户身份证是否绑定 = False: Exit Function
    If rsTemp.RecordCount <= 0 Then zl当前用户身份证是否绑定 = False: Exit Function
    
    If IIf(IsNull(rsTemp!姓名), "", rsTemp!姓名) = strName And IIf(IsNull(rsTemp!门诊号), "", rsTemp!门诊号) = str门诊号 Then
        zl当前用户身份证是否绑定 = True
    Else
        zl当前用户身份证是否绑定 = False
    End If
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then Resume
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

        Set ObjItem = tbcPage.InsertItem(mPageIndex.基本, "基本", picInfo.hWnd, 0)
        ObjItem.Tag = mPageIndex.基本
    
        Set ObjItem = tbcPage.InsertItem(mPageIndex.健康档案, "健康档案", PicHealth.hWnd, 0)
        ObjItem.Tag = mPageIndex.健康档案
        Call InitVsInoculate
        Call InitVsOtherInfo
        Call InitCombox
        
        '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
        If Not mobjPlugIn Is Nothing Then
            On Error Resume Next
            mlngPlugInHwnd = mobjPlugIn.GetFormHwnd
            Call zlPlugInErrH(Err, "GetFormHwnd")
            Err.Clear: On Error GoTo 0
            If mlngPlugInHwnd <> 0 Then
                picTaskPanelOther.Visible = True
                Set ObjItem = tbcPage.InsertItem(mPageIndex.附加信息, "附加信息", picTaskPanelOther.hWnd, 0)
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
    If gobjComlib.ErrCenter = 1 Then
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
    Dim strSql As String
    Dim varKey As Variant
    Dim intCount As Integer
    '过敏药物
    With msh过敏
        If .Rows > 1 Then
            '清除该病人所有记录
            strSql = " Zl_病人过敏药物_Delete(" & lng病人ID & ")"
            zlAddArray colPro, strSql
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    '病人过敏药物
                    strSql = "Zl_病人过敏药物_Update("
                    '病人ID_In 病人过敏药物.病人Id%Type
                    strSql = strSql & "" & lng病人ID & ","
                    '过敏药物ID_In 病人过敏药物.过敏药物ID%Type
                    strSql = strSql & "'" & IIf(.RowData(i) <= 0, "", .RowData(i)) & "',"
                    '过敏药物_In  病人过敏药物.过敏药物%Type
                    strSql = strSql & "'" & IIf(.TextMatrix(i, 0) = "", "", .TextMatrix(i, 0)) & "',"
                    '过敏反应_In 病人过敏反应.过敏反应%Type
                    strSql = strSql & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"

                    zlAddArray colPro, strSql
                End If
            Next
        End If
    End With
    '接种信息
    With vsInoculate
        If .Rows > 1 Then
            '清除该病人所有记录
            strSql = " Zl_病人免疫记录_Delete(" & lng病人ID & ")"
            zlAddArray colPro, strSql

            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    '病人过敏药物
                    strSql = "Zl_病人免疫记录_Update("
                    '病人ID_In 病人免疫记录.病人Id%Type
                    strSql = strSql & "" & lng病人ID & ","
                    '接种时间_In 病人免疫记录.接种时间%Type
                    strSql = strSql & "" & IIf(.TextMatrix(i, 0) = "", "''", "to_date('" & .TextMatrix(i, 0) & "','yyyy-mm-dd')") & ","
                    '接种名称_In  病人免疫记录.接种名称%Type
                    strSql = strSql & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"
                    zlAddArray colPro, strSql
                End If
                If .TextMatrix(i, 3) <> "" Then
                    '病人过敏药物
                    strSql = "Zl_病人免疫记录_Update("
                    '病人ID_In 病人免疫记录.病人Id%Type
                    strSql = strSql & "" & lng病人ID & ","
                    '接种时间_In 病人免疫记录.接种时间%Type
                    strSql = strSql & "" & IIf(.TextMatrix(i, 2) = "", "''", "to_date('" & .TextMatrix(i, 2) & "','yyyy-mm-dd')") & ","
                    '接种名称_In  病人免疫记录.接种名称%Type
                    strSql = strSql & "'" & IIf(.TextMatrix(i, 3) = "", "''", .TextMatrix(i, 3)) & "')"
                    zlAddArray colPro, strSql
                End If
            Next
        End If
    End With
    '其他信息
    'ABO血型
    '病人信息从表
    strSql = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSql = strSql & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSql = strSql & "'血型',"
    '信息值_In 病人信息从表.信息值%Type
    strSql = strSql & "'" & gobjCommFun.GetNeedName(cboBloodType.Text, ".") & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSql = strSql & "'')"
    zlAddArray colPro, strSql
    'RH
    strSql = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSql = strSql & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSql = strSql & "'RH',"
    '信息值_In 病人信息从表.信息值%Type
    strSql = strSql & "'" & cboBH.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSql = strSql & "'')"
    zlAddArray colPro, strSql
    '医学警示
    strSql = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSql = strSql & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSql = strSql & "'医学警示',"
    '信息值_In 病人信息从表.信息值%Type
    strSql = strSql & "'" & txtMedicalWarning.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSql = strSql & "'')"
    zlAddArray colPro, strSql
    '其他医学警示
    strSql = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSql = strSql & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSql = strSql & "'其他医学警示',"
    '信息值_In 病人信息从表.信息值%Type
    strSql = strSql & "'" & txtOtherWaring.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSql = strSql & "'')"
    zlAddArray colPro, strSql
        
    '84313:李南春,2015/4/29, 第一条联系人信息已保存在病人信息中，从表中不再重复保存
    '联系人相关信息
    intCount = 0
    With vsLinkMan
        If .Rows >= 3 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then '联系人姓名不能为空
                    intCount = intCount + 1
                    For j = 0 To .Cols - 1
                        strSql = "Zl_病人信息从表_Update("
                        '病人ID_In 病人信息从表.病人Id%Type
                        strSql = strSql & "" & lng病人ID & ","
                        '信息名_In 病人信息从表.信息名%Type
                        strSql = strSql & "'联系人" & .TextMatrix(0, j) & intCount & "',"
                        '信息值_In 病人信息从表.信息值%Type
                        strSql = strSql & "'" & IIf(.TextMatrix(i, j) = "", "", .TextMatrix(i, j)) & "',"
                        '就诊Id_In 病人信息从表.就诊Id%Type
                        strSql = strSql & "'')"

                        zlAddArray colPro, strSql
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
                    strSql = "Zl_病人信息从表_Update("
                    '病人ID_In 病人信息从表.病人Id%Type
                    strSql = strSql & "" & lng病人ID & ","
                    '信息名_In 病人信息从表.信息名%Type
                    strSql = strSql & "'" & .TextMatrix(i, 0) & "',"
                    '信息值_In 病人信息从表.信息值%Type
                    strSql = strSql & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "',"
                    '就诊Id_In 病人信息从表.就诊Id%Type
                    strSql = strSql & "'')"

                    zlAddArray colPro, strSql
                End If
                If .TextMatrix(i, 2) <> "" Then
                    strSql = "Zl_病人信息从表_Update("
                    '病人ID_In 病人信息从表.病人Id%Type
                    strSql = strSql & "" & lng病人ID & ","
                    '信息名_In 病人信息从表.信息名%Type
                    strSql = strSql & "'" & .TextMatrix(i, 2) & "',"
                    '信息值_In 病人信息从表.信息值%Type
                    strSql = strSql & "'" & IIf(.TextMatrix(i, 3) = "", "", .TextMatrix(i, 3)) & "',"
                    '就诊Id_In 病人信息从表.就诊Id%Type
                    strSql = strSql & "'')"

                    zlAddArray colPro, strSql
                End If
            Next
        End If
     End With
     If lng就诊ID = 0 Then Exit Sub
     'ABO血型
    '病人信息从表
    strSql = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSql = strSql & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSql = strSql & "'血型',"
    '信息值_In 病人信息从表.信息值%Type
    strSql = strSql & "'" & gobjCommFun.GetNeedName(cboBloodType.Text, ".") & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSql = strSql & lng就诊ID & ")"
    zlAddArray colPro, strSql
    'RH
    strSql = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSql = strSql & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSql = strSql & "'RH',"
    '信息值_In 病人信息从表.信息值%Type
    strSql = strSql & "'" & cboBH.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSql = strSql & lng就诊ID & ")"
    zlAddArray colPro, strSql
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
                        cbo联系人关系.ListIndex = gobjControl.Cbo.FindIndex(cbo联系人关系, str关系, True)
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
        If gobjControl.Cbo.FindIndex(cbo联系人关系, str关系, True) = -1 And str关系 <> "" Then
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
            cbo联系人关系.ListIndex = gobjControl.Cbo.FindIndex(cbo联系人关系, str关系, True)
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
    Dim strSql As String
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
    strSql = "" & _
    "   Select 病人ID,过敏药物ID,过敏药物,过敏反应 From 病人过敏药物 Where 病人ID=[1]"
    Set rs过敏药物 = gobjDatabase.OpenSQLRecord(strSql, "病人过敏药物", lng病人ID)
    While rs过敏药物.EOF = False
        SetDrugAllergy Nvl(rs过敏药物!过敏药物), Nvl(rs过敏药物!过敏反应), Nvl(rs过敏药物!过敏药物ID, 0)
        rs过敏药物.MoveNext
    Wend
    '获取免疫记录
    strSql = "" & _
    "   Select 病人ID,接种时间,接种名称 From 病人免疫记录 Where 病人ID=[1]"
    Set rs免疫记录 = gobjDatabase.OpenSQLRecord(strSql, "病人免疫记录", lng病人ID)
    While rs免疫记录.EOF = False
        SetInoculate Nvl(rs免疫记录!接种时间), Nvl(rs免疫记录!接种名称)
        rs免疫记录.MoveNext
    Wend
    '血型
    strSql = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='血型' And 就诊ID Is NULL"
    Set rsABO血型 = gobjDatabase.OpenSQLRecord(strSql, "ABO血型", lng病人ID)
    While rsABO血型.EOF = False
        For i = 0 To cboBloodType.ListCount - 1
            '76314,李南春，2014-08-06，病人信息正确获取
            If gobjCommFun.GetNeedName(cboBloodType.List(i), ".") = NeedName(Nvl(rsABO血型!信息值)) Then cboBloodType.ListIndex = i
        Next
        rsABO血型.MoveNext
    Wend
    'RH
    strSql = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='RH' And 就诊ID Is NULL"
    Set rsRH = gobjDatabase.OpenSQLRecord(strSql, "RH", lng病人ID)
    While rsRH.EOF = False
        For i = 0 To cboBH.ListCount - 1
            If cboBH.List(i) = Nvl(rsRH!信息值) Then cboBH.ListIndex = i
        Next
        rsRH.MoveNext
    Wend
    '医学警示
    strSql = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='医学警示'"
    Set rs医学警示 = gobjDatabase.OpenSQLRecord(strSql, "医学警示", lng病人ID)
    While rs医学警示.EOF = False
        str医学警示 = str医学警示 & "," & Nvl(rs医学警示!信息值)
        rs医学警示.MoveNext
    Wend
    If str医学警示 <> "" Then str医学警示 = Mid(str医学警示, 2)
    txtMedicalWarning.Text = str医学警示
    '其他医学警示
    strSql = "" & _
    "  Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='其他医学警示'"
    Set rs其他医学警示 = gobjDatabase.OpenSQLRecord(strSql, "其他医学警示", lng病人ID)
    While rs其他医学警示.EOF = False
        txtOtherWaring.Text = Nvl(rs其他医学警示!信息值)
        rs其他医学警示.MoveNext
    Wend
    '联系人相关信息
    '取病人信息表中的联系人信息
    '84313,李南春,2015/4/27,联系人关系以及其他关系
    strSql = "" & _
    "   Select  A.联系人姓名,A.联系人关系,A.联系人电话,A.联系人身份证号,B.信息值 as 附加信息 From 病人信息 A,病人信息从表 B " & _
    "   Where A.病人ID=B.病人ID(+) And A.病人ID=[1] And B.信息名(+)='联系人附加信息' And Not A.联系人姓名 is Null"
    Set rs病人信息 = gobjDatabase.OpenSQLRecord(strSql, "病人信息联系人信息", lng病人ID)
    If rs病人信息.EOF = False Then
        txt联系人身份证.Text = Nvl(rs病人信息!联系人身份证号)
        txt联系人姓名.Text = Nvl(rs病人信息!联系人姓名)
        txt联系人电话.Text = Nvl(rs病人信息!联系人电话)
        cbo联系人关系.ListIndex = gobjControl.Cbo.FindIndex(cbo联系人关系, Nvl(rs病人信息!联系人关系), True)
        If cbo联系人关系.ListIndex = -1 And Nvl(rs病人信息!联系人关系) <> "" Then
            cbo联系人关系.ListIndex = 8: txt其他关系.Text = rs病人信息!联系人关系
        ElseIf cbo联系人关系.ListIndex = 8 Then
            txt其他关系.Text = Nvl(rs病人信息!附加信息)
        End If
        SetLinkInfo Nvl(rs病人信息!联系人姓名), Nvl(rs病人信息!联系人关系), Nvl(rs病人信息!联系人电话), Nvl(rs病人信息!联系人身份证号), txt其他关系.Text
    End If
    '取病人信息从表中的联系人信息
    strSql = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名 like '联系人%' order by 信息名 Asc"
    Set rs联系人 = gobjDatabase.OpenSQLRecord(strSql, "联系人相关信息", lng病人ID)
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
    strSql = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名 Not in ('血型','ABO','RH','医学警示','其他医学警示') And 信息名 Not like '联系人%'"
    Set rs其他信息 = gobjDatabase.OpenSQLRecord(strSql, "联系人其他信息", lng病人ID)
    '问题号:115886,焦博,2017/11/08,挂号提取该病人信息时，程序报错
    While rs其他信息.EOF = False
        If Nvl(rs其他信息!信息名) <> "" Then
            SetOtherInfo Nvl(rs其他信息!信息名), Nvl(rs其他信息!信息值)
        End If
        rs其他信息.MoveNext
    Wend
    
    Exit Sub
ErrHandl:
     If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

'Private Function bln发卡(Optional ByVal blnCardNo As Boolean = False) As Boolean
''---------------------------------------------------------------------------------------------------------------------------------------------
''功能:判断当前是否为卡发操作 (不是发卡操作既是绑定卡操作)
''入参:
''编制:56599
''日期:2012-12-12 14:55:36
''---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim lng磁卡领用ID As Long
'    Dim bln是否发卡 As Boolean
'    If gCurSendCard.bln严格控制 = True Then
'        lng磁卡领用ID = CheckUsedBill(5, IIf(lng磁卡领用ID > 0, lng磁卡领用ID, gCurSendCard.lng共用批次), IIf(blnCardNo, mstrCard, UCase(txtPatient.Text)), gCurSendCard.lng卡类别ID)
'        bln是否发卡 = IIf(lng磁卡领用ID = -3, False, True)
'        If gCurSendCard.bln自制卡 = False Then
'            bln是否发卡 = (gCurSendCard.bln是否发卡 = True)
'        End If
'    Else
'        bln是否发卡 = mbln发卡
'        If gCurSendCard.bln自制卡 = False Then
'            bln是否发卡 = (gCurSendCard.bln是否发卡 = True)
'        End If
'    End If
'    bln发卡 = bln是否发卡
'    mbln发卡 = bln是否发卡
'End Function

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
            gobjCommFun.PressKey vbKeyTab
        ElseIf vsInoculate.Col = 3 Then
            vsInoculate.Col = 0: vsInoculate.Row = vsInoculate.Row + 1
            gobjCommFun.PressKey vbKeyReturn
        Else
            gobjCommFun.PressKey vbKeyRight
        End If
    End If
End Sub

Private Function BlandCancel(ByVal lngCardTypeID As Long, ByVal strCardNO As String, ByVal lngPatientID As Long) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:取消绑定卡
'入参:intType:0-当前卡号;1-当前类别;2-当前病人所有
'返回:取消成功,返回true,否则返回False
'编制:刘兴洪
'日期:2011-07-29 11:18:05
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim curDate As Date
    Dim strSql As String, strPassWord As String

    On Error GoTo errHandle

    curDate = gobjDatabase.Currentdate
    
    'Zl_医疗卡变动_Insert
    strSql = "Zl_医疗卡变动_Insert("
    '      变动类型_In   Number,
    '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
    strSql = strSql & "" & 14 & ","
    '      病人id_In     住院费用记录.病人id%Type,
    strSql = strSql & "" & lngPatientID & ","
    '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
    strSql = strSql & "" & lngCardTypeID & ","
    '      原卡号_In     病人医疗卡信息.卡号%Type,
    strSql = strSql & "NULL,"
    '      医疗卡号_In   病人医疗卡信息.卡号%Type,
    strSql = strSql & "'" & strCardNO & "'" & ","
    '      变动原因_In   病人医疗卡变动.变动原因%Type,
    strSql = strSql & "'挂号绑定卡自动取消绑定',"
    '      密码_In       病人信息.卡验证码%Type,
    strSql = strSql & "NULL,"
    '      操作员姓名_In 住院费用记录.操作员姓名%Type,
    strSql = strSql & "NULL,"
    '      变动时间_In   住院费用记录.登记时间%Type,
    strSql = strSql & "to_date('" & Format(curDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
    strSql = strSql & "NULL,"
    '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
    strSql = strSql & "NULL)"

     
    Call gobjDatabase.ExecuteProcedure(strSql, Me.Caption)
    BlandCancel = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
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

Private Function CreateObjectPlugIn() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建渠道附加信息插件
    '返回:创建成功,返回True,否则返回False
    '问题号:73935
    '编制:冉俊明
    '日期:2014-07-3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnPlugin = False
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear: On Error GoTo 0
    End If
    
    If Not mobjPlugIn Is Nothing Then
        On Error Resume Next
        Call mobjPlugIn.Initialize(gcnOracle, glngSys, 1111)
        mblnPlugin = Err = 0
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
    End If
    CreateObjectPlugIn = True
End Function

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
    If Not mobjPlugIn Is Nothing Then
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
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Sub DeletePatPicture(lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除病人照片
    '入参:lng病人ID - 病人ID
    '编制:56599
    '日期:2012-12-14 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo Errhand:
    strSql = strSql & "Zl_病人照片_Delete("
    strSql = strSql & lng病人ID & ")"
    
    gobjDatabase.ExecuteProcedure strSql, Me.Caption
    
    Exit Sub
Errhand:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub SavePatPicture(lng病人ID As Long, strFile As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存病人照片
    '入参:lng病人ID - 病人ID
    '编制:56599
    '日期:2012-12-13 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
        
    If strFile = "" Then Exit Sub

    If gobjComlib.Sys.SaveLob(glngSys, 27, lng病人ID, strFile, 0) = False Then
        ShowMsgBox "保存照片有误,请确认文件是否被删除!"
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

    strTmp = gobjComlib.Sys.ReadLob(glngSys, 27, lng病人ID)
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
    If gobjComlib.ErrCenter() = 1 Then Resume
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

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

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

Private Sub zlQueryEMPIPatiInfo()
    '功能：从EMPI平台获取病人信息
    '日期：2016/10/9 10:47:13
    '编制：李南春
    '说明：101170
    Dim rsTmp As ADODB.Recordset, strDiff As String, strMsgInfo As String
    If mblnPlugin = False Then Exit Sub
    If mobjPlugIn Is Nothing Then Exit Sub
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
        !性别 = NeedName(cbo性别.Text)
        If IsDate(txt出生日期.Text) Then
            !出生日期 = Format(txt出生日期.Text & " " & IIf(IsDate(txt出生时间.Text), txt出生时间.Text, "00:00"), "YYYY-MM-DD HH:MM")
        Else
            !出生日期 = ""
        End If
        !出生地点 = txtBirthLocation.Text
        !国籍 = NeedName(cbo国籍.Text)
        !民族 = NeedName(cbo民族.Text)
        !职业 = NeedName(cbo职业.Text)
        !工作单位 = txt单位名称.Text
        !婚姻状况 = NeedName(cbo婚姻.Text)
        !家庭电话 = txt家庭电话.Text
        !联系人电话 = txt联系人电话.Text
        !单位电话 = txt单位电话.Text
        !家庭地址 = cbo家庭地址.Text
        !家庭地址邮编 = txt家庭邮编.Text
        !户口地址 = txtRegLocation.Text
        !户口地址邮编 = txt户口地址邮编.Text
        !单位邮编 = txt单位邮编.Text
        !联系人姓名 = txt联系人姓名.Text
        !联系人关系 = NeedName(cbo联系人关系.Text)
        .Update
    End With
    'EMPI没有找到病人信息,直接返回
    Dim rsOut As New ADODB.Recordset
    On Error Resume Next
    If mobjPlugIn.EMPI_QueryPatiInfo(glngSys, glngModul, rsTmp, rsOut) = False Then
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
        mstrPlugChange = ""
        If Nvl(!医保号) <> "" Then txtPatiMCNO(0).Text = Nvl(!医保号): txtPatiMCNO(1).Text = Nvl(!医保号)
        If Nvl(!身份证号) <> "" Then txt身份证号.Text = Nvl(!身份证号)
        If mlng病人ID = 0 Then
            If Nvl(!姓名) <> "" Then txtPatient.Text = Nvl(!姓名)
            If Nvl(!性别) <> "" Then cbo性别.ListIndex = gobjControl.Cbo.FindIndex(cbo性别, Nvl(!性别), True)
            If Nvl(!出生日期) <> Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS") Then
                txt出生日期.Text = Format(Nvl(!出生日期), "YYYY-MM-DD")
                txt出生时间.Text = Format(Nvl(!出生日期), "HH:MM")
            End If
        Else
            If Nvl(!姓名) <> "" And txtPatient.Text <> Nvl(!姓名) Then strDiff = ",姓名"
            If Nvl(!性别) <> "" And cbo性别.ListIndex <> gobjControl.Cbo.FindIndex(cbo性别, Nvl(!性别), True) Then strDiff = strDiff & ",性别"
            If Nvl(!出生日期) <> "" And Format(Nvl(!出生日期), "YYYY-MM-DD HH:MM:SS") <> Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS") Then strDiff = strDiff & ",出生日期"
            If Nvl(!身份证号) <> "" And txt身份证号.Text <> Nvl(!身份证号) Then strDiff = strDiff & ",身份证号"
        End If
        
        If InStr(gstrPrivs, ";允许修改门诊号;") > 0 Or mlng病人ID = 0 Then
            If Nvl(!门诊号) <> "" Then txt门诊号.Text = Nvl(!门诊号)
        Else
            If Nvl(!门诊号) <> "" Then txt门诊号.Text = Nvl(!门诊号)
        End If
        
        If Nvl(!出生地点) <> "" Then txtBirthLocation.Text = Nvl(!出生地点)
        If Nvl(!国籍) <> "" Then cbo国籍.ListIndex = gobjControl.Cbo.FindIndex(cbo国籍, Nvl(!国籍), True)
        If Nvl(!民族) <> "" Then cbo民族.ListIndex = gobjControl.Cbo.FindIndex(cbo民族, Nvl(!民族), True)
        If Nvl(!职业) <> "" Then cbo职业.ListIndex = gobjControl.Cbo.FindIndex(cbo职业, Nvl(!职业))
        If Nvl(!工作单位) <> "" Then txt单位名称.Text = Nvl(!工作单位)
        If Nvl(!婚姻状况) <> "" Then cbo婚姻.ListIndex = gobjControl.Cbo.FindIndex(cbo婚姻, Nvl(!婚姻状况), True)
        If Nvl(!家庭电话) <> "" Then txt家庭电话.Text = Nvl(!家庭电话)
        If Nvl(!联系人电话) <> "" Then txt联系人电话.Text = Nvl(!联系人电话)
        If Nvl(!单位电话) <> "" Then txt单位电话.Text = Nvl(!单位电话)
        If Nvl(!家庭地址) <> "" Then cbo家庭地址.Text = Nvl(!家庭地址): padd家庭地址.Value = Nvl(!家庭地址)
        If Nvl(!家庭地址邮编) <> "" Then txt家庭邮编.Text = Nvl(!家庭地址邮编)
        If Nvl(!户口地址) <> "" Then txtRegLocation.Text = Nvl(!户口地址): padd户口地址.Value = Nvl(!户口地址)
        If Nvl(!户口地址邮编) <> "" Then txt户口地址邮编.Text = Nvl(!户口地址邮编)
        If Nvl(!单位邮编) <> "" Then txt单位邮编.Text = Nvl(!单位邮编)
        If Nvl(!联系人姓名) <> "" Then txt联系人姓名.Text = Nvl(!联系人姓名)
        If Nvl(!联系人关系) <> "" Then cbo联系人关系.ListIndex = gobjControl.Cbo.FindIndex(cbo联系人关系, Nvl(!联系人关系), True)
    End With
    Err = 0: On Error GoTo 0
    If mlng病人ID <> 0 Then
        If strDiff <> "" Then strDiff = Mid(strDiff, 2)
        If strDiff <> "" Then
            strMsgInfo = "病人的 " & strDiff & " 与EMPI信息不一致，因您不具有相应的权限，本次不会进行更新。"
        End If
        If strMsgInfo <> "" Then MsgBox strMsgInfo, vbInformation, gstrSysName
    End If
    Exit Sub
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Public Function zlSaveEMPIPatiInfo(ByVal blnNewPati As Boolean, ByVal lngPatiID As Long, ByVal lngClinicID As Long, ByRef strErrMsg As String) As Boolean
    '功能:上传病人信息到EMPI平台,如果平台信息保存失败，连同HIS数据一起回退
    '参数: In-lngPatiID 病人ID,lngClinicID 挂号ID
    '      Out-strErrMsg 错误信息，函数返回失败有效
    '返回:True-EMPI平台保存信息成功,False-保存失败
    '编制:李南春
    '说明:100915
    Dim blnCharge As Boolean, lngRet As Long
    If mblnPlugin = False Then zlSaveEMPIPatiInfo = True: Exit Function
    If mobjPlugIn Is Nothing Then zlSaveEMPIPatiInfo = True: Exit Function
    
    On Error GoTo Errhand
    If mrsEMPIOut Is Nothing Then
        'EMPI没有病人信息，需要新建
        On Error Resume Next
        lngRet = mobjPlugIn.EMPI_AddPatiInfo(glngSys, glngModul, lngPatiID, 0, lngClinicID, strErrMsg)
        Call zlPlugInErrH(Err, "EMPI_AddPatiInfo")
        If lngRet = 0 And Err.Number <> 438 Then Err.Clear: Exit Function
        Err.Clear: On Error GoTo Errhand
    Else
        '判断平台回传的信息是否发生改变
        With mrsEMPIOut
            If txtPatiMCNO(0).Text <> Nvl(!医保号) Then blnCharge = True: GoTo EMPIModify
            If blnNewPati Then
                If txt身份证号.Text <> Nvl(!身份证号) Then blnCharge = True: GoTo EMPIModify
                If txtPatient.Text <> Nvl(!姓名) Then blnCharge = True: GoTo EMPIModify
                If cbo性别.ListIndex <> gobjControl.Cbo.FindIndex(cbo性别, Nvl(!性别), True) Then blnCharge = True: GoTo EMPIModify
                If Format(txt出生日期.Text, "YYYY-MM-DD") <> Format(Nvl(!出生日期), "YYYY-MM-DD") Then blnCharge = True: GoTo EMPIModify
                If Format(txt出生时间.Text, "HH:MM") <> Format(Nvl(!出生日期), "HH:MM") Then blnCharge = True: GoTo EMPIModify
            End If
            
            If InStr(gstrPrivs, ";允许修改门诊号;") > 0 Or blnNewPati Then
                If txt门诊号.Text <> Nvl(!门诊号) Then blnCharge = True: GoTo EMPIModify
            End If
            If txtBirthLocation.Text <> Nvl(!出生地点) Then blnCharge = True: GoTo EMPIModify
            If cbo国籍.ListIndex <> gobjControl.Cbo.FindIndex(cbo国籍, Nvl(!国籍), True) Then blnCharge = True: GoTo EMPIModify
            If cbo民族.ListIndex <> gobjControl.Cbo.FindIndex(cbo民族, Nvl(!民族), True) Then blnCharge = True: GoTo EMPIModify
            If cbo职业.ListIndex <> gobjControl.Cbo.FindIndex(cbo职业, Nvl(!职业)) Then blnCharge = True: GoTo EMPIModify
            If txt单位名称.Text <> Nvl(!工作单位) Then blnCharge = True: GoTo EMPIModify
            If cbo婚姻.ListIndex <> gobjControl.Cbo.FindIndex(cbo婚姻, Nvl(!婚姻状况), True) Then blnCharge = True: GoTo EMPIModify
            If txt家庭电话.Text <> Nvl(!家庭电话) Then blnCharge = True: GoTo EMPIModify
            If txt联系人电话.Text <> Nvl(!联系人电话) Then blnCharge = True: GoTo EMPIModify
            If txt单位电话.Text <> Nvl(!单位电话) Then blnCharge = True: GoTo EMPIModify
            If cbo家庭地址.Text <> Nvl(!家庭地址) Then blnCharge = True: GoTo EMPIModify
            If txt家庭邮编.Text <> Nvl(!家庭地址邮编) Then blnCharge = True: GoTo EMPIModify
            If txtRegLocation.Text <> Nvl(!户口地址) Then blnCharge = True: GoTo EMPIModify
            If txt户口地址邮编.Text <> Nvl(!户口地址邮编) Then blnCharge = True: GoTo EMPIModify
            If txt单位邮编.Text <> Nvl(!单位邮编) Then blnCharge = True: GoTo EMPIModify
            If txt联系人姓名.Text <> Nvl(!联系人姓名) Then blnCharge = True: GoTo EMPIModify
            If cbo联系人关系.ListIndex <> gobjControl.Cbo.FindIndex(cbo联系人关系, Nvl(!联系人关系), True) Then blnCharge = True: GoTo EMPIModify
        End With
    End If
EMPIModify:
    If blnCharge Then
        On Error Resume Next
        lngRet = mobjPlugIn.EMPI_ModifyPatiInfo(glngSys, glngModul, lngPatiID, 0, lngClinicID, strErrMsg)
        Call zlPlugInErrH(Err, "EMPI_AddPatiInfo")
        If lngRet = 0 And Err.Number <> 438 Then Err.Clear: Exit Function
        Err.Clear: On Error GoTo Errhand
    End If
    zlSaveEMPIPatiInfo = True
    Exit Function
Errhand:
    strErrMsg = Err.Description
    Call zlPlugInErrH(Err, "zlSaveEMPIPatiInfo")
    Call gobjComlib.SaveErrLog
End Function
