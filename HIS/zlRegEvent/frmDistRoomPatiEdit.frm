VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{D01C2596-4FE0-4EA9-9EE8-D97BE62A1165}#4.3#0"; "ZlPatiAddress.ocx"
Begin VB.Form frmDistRoomPatiEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病案信息编辑"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9885
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPatiInfo 
      BorderStyle     =   0  'None
      Height          =   7620
      Left            =   0
      ScaleHeight     =   7620
      ScaleWidth      =   9885
      TabIndex        =   38
      Top             =   0
      Width           =   9885
      Begin ZlPatiAddress.PatiAddress padd户口地址 
         Height          =   360
         Left            =   1080
         TabIndex        =   13
         Tag             =   "户口地址"
         Top             =   1485
         Visible         =   0   'False
         Width           =   6270
         _ExtentX        =   11060
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
         Left            =   1080
         TabIndex        =   10
         Tag             =   "现住址"
         Top             =   1035
         Visible         =   0   'False
         Width           =   6270
         _ExtentX        =   11060
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
      Begin VB.TextBox txt过敏反应 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4740
         MaxLength       =   50
         TabIndex        =   30
         Top             =   6030
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.TextBox txt户口邮编 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   8775
         MaxLength       =   6
         TabIndex        =   14
         Tag             =   "户口地址邮编"
         Top             =   1485
         Width           =   960
      End
      Begin VB.TextBox txtPatiMCNO 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   6135
         MaxLength       =   30
         TabIndex        =   16
         Top             =   1944
         Width           =   2205
      End
      Begin VB.CommandButton cmd户口地址 
         Caption         =   "…"
         Height          =   330
         Left            =   6975
         TabIndex        =   68
         TabStop         =   0   'False
         Tag             =   "户口地址"
         Top             =   1503
         Width           =   360
      End
      Begin VB.CommandButton cmd家庭地址 
         Caption         =   "…"
         Height          =   330
         Left            =   6975
         TabIndex        =   0
         ToolTipText     =   "热键F3"
         Top             =   1047
         Width           =   360
      End
      Begin VB.Frame Frame3 
         Height          =   35
         Left            =   -30
         TabIndex        =   67
         Top             =   5490
         Width           =   10005
      End
      Begin VB.TextBox txtPatiMCNO 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   15
         Top             =   1944
         Width           =   2325
      End
      Begin VB.ComboBox cbo付款方式 
         Height          =   360
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2400
         Width           =   2325
      End
      Begin VB.TextBox txtPatient 
         Height          =   360
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   1
         Top             =   120
         Width           =   2145
      End
      Begin VB.TextBox txt年龄 
         Height          =   360
         IMEMode         =   2  'OFF
         Left            =   4560
         TabIndex        =   6
         Top             =   570
         Width           =   1515
      End
      Begin VB.Frame Frame1 
         Height          =   35
         Left            =   -60
         TabIndex        =   66
         Top             =   2880
         Width           =   10005
      End
      Begin VB.TextBox txt门诊号 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4560
         TabIndex        =   2
         Top             =   120
         Width           =   2175
      End
      Begin VB.ComboBox cbo性别 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7710
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   1185
      End
      Begin VB.ComboBox cbo费别 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7815
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox cbo国籍 
         Height          =   360
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3015
         Width           =   2325
      End
      Begin VB.ComboBox cbo民族 
         Height          =   360
         Left            =   4545
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3015
         Width           =   2085
      End
      Begin VB.ComboBox cbo婚姻 
         Height          =   360
         Left            =   7815
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3015
         Width           =   1935
      End
      Begin VB.ComboBox cbo职业 
         Height          =   360
         Left            =   4545
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3472
         Width           =   2085
      End
      Begin VB.TextBox txt身份证号 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   23
         Top             =   3472
         Width           =   2325
      End
      Begin VB.TextBox txt单位电话 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5700
         MaxLength       =   20
         TabIndex        =   27
         Top             =   3930
         Width           =   1695
      End
      Begin VB.TextBox txt单位邮编 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   8535
         MaxLength       =   6
         TabIndex        =   28
         Top             =   3930
         Width           =   1215
      End
      Begin VB.TextBox txt家庭电话 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7710
         MaxLength       =   20
         TabIndex        =   8
         Top             =   570
         Width           =   2025
      End
      Begin VB.TextBox txt家庭邮编 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   8775
         MaxLength       =   6
         TabIndex        =   11
         Top             =   1032
         Width           =   960
      End
      Begin VB.CommandButton cmd单位名称 
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
         Height          =   330
         Left            =   4110
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3945
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
         Left            =   6270
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "热键:F3"
         Top             =   6090
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txt过敏 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   930
         MaxLength       =   50
         TabIndex        =   35
         Top             =   6090
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.ComboBox cbo年龄单位 
         Height          =   360
         Left            =   6075
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   570
         Width           =   675
      End
      Begin VB.ComboBox cbo医疗类别 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4545
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2400
         Width           =   2085
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   0
         Left            =   7815
         MaxLength       =   64
         TabIndex        =   25
         Top             =   3472
         Width           =   1935
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   1515
         Left            =   2010
         TabIndex        =   37
         Top             =   5640
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
         NumItems        =   0
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   2820
         Top             =   7320
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
               Picture         =   "frmDistRoomPatiEdit.frx":0000
               Key             =   "Itemps"
               Object.Tag             =   "Itemgm"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistRoomPatiEdit.frx":059A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox txt出生时间 
         Height          =   360
         Left            =   2460
         TabIndex        =   5
         Top             =   570
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt出生日期 
         Height          =   360
         Left            =   1080
         TabIndex        =   4
         Top             =   570
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
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
         Format          =   "YYYY-MM-DD"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin zl9RegEvent.UCPatiVitalSigns UCPatiVitalSigns 
         Height          =   945
         Left            =   510
         TabIndex        =   29
         Top             =   4500
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   1667
         TextBackColor   =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         XDis            =   100
         YDis            =   120
         LabToTxt        =   120
      End
      Begin VB.TextBox txt单位名称 
         Height          =   360
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   26
         Top             =   3960
         Width           =   3045
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh过敏 
         Height          =   1215
         Left            =   60
         TabIndex        =   69
         ToolTipText     =   "F2:修改,F3:选择"
         Top             =   5670
         Width           =   9705
         _ExtentX        =   17119
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
      Begin VB.ComboBox cbo家庭地址 
         Height          =   360
         Left            =   1080
         TabIndex        =   9
         Top             =   1032
         Width           =   5895
      End
      Begin VB.TextBox txt户口地址 
         Height          =   360
         Left            =   1080
         TabIndex        =   12
         Tag             =   "户口地址"
         Top             =   1488
         Width           =   5895
      End
      Begin VB.Label lbl户口地址 
         Alignment       =   1  'Right Justify
         Caption         =   "户口地址"
         Height          =   270
         Left            =   90
         TabIndex        =   65
         Top             =   1533
         Width           =   960
      End
      Begin VB.Label lbl户口邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口邮编"
         Height          =   240
         Left            =   7755
         TabIndex        =   64
         Top             =   1530
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   240
         Left            =   3810
         TabIndex        =   60
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   240
         Left            =   570
         TabIndex        =   33
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   240
         Left            =   7215
         TabIndex        =   59
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   240
         Left            =   4050
         TabIndex        =   58
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   240
         Left            =   6840
         TabIndex        =   57
         Top             =   3075
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   240
         Left            =   4050
         TabIndex        =   56
         Top             =   3525
         Width           =   480
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         Height          =   240
         Left            =   4050
         TabIndex        =   55
         Top             =   3075
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍"
         Height          =   240
         Left            =   570
         TabIndex        =   54
         Top             =   3075
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   240
         Left            =   90
         TabIndex        =   53
         Top             =   3532
         Width           =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称"
         Height          =   240
         Left            =   90
         TabIndex        =   52
         Top             =   3990
         Width           =   960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位电话"
         Height          =   240
         Left            =   4710
         TabIndex        =   51
         Top             =   3990
         Width           =   960
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位邮编"
         Height          =   240
         Left            =   7560
         TabIndex        =   50
         Top             =   3990
         Width           =   960
      End
      Begin VB.Label lbl家庭地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "现住址"
         Height          =   240
         Left            =   330
         TabIndex        =   49
         Top             =   1092
         Width           =   720
      End
      Begin VB.Label lbl电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "电话"
         Height          =   240
         Left            =   7200
         TabIndex        =   48
         Top             =   630
         Width           =   480
      End
      Begin VB.Label lbl邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "邮编"
         Height          =   240
         Left            =   8235
         TabIndex        =   47
         Top             =   1095
         Width           =   480
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   240
         Left            =   7290
         TabIndex        =   46
         Top             =   2460
         Width           =   480
      End
      Begin VB.Label lbl付款方式 
         BackStyle       =   0  'Transparent
         Caption         =   "付款方式"
         Height          =   240
         Left            =   75
         TabIndex        =   45
         Top             =   2460
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "定位到过敏药物处,F2修改,F3选择.如果当前行有内容,则输入文字可修改过敏药物名称,否则以输入内容为关键字按简码、名称、编码查找过敏药物."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   7170
         Width           =   9465
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         Height          =   240
         Left            =   75
         TabIndex        =   43
         Top             =   630
         Width           =   975
      End
      Begin VB.Label lblPatiMCNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         Height          =   240
         Index           =   0
         Left            =   330
         TabIndex        =   42
         Top             =   2004
         Width           =   720
      End
      Begin VB.Label lblPatiMCNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "验证医保号"
         Height          =   240
         Index           =   1
         Left            =   4890
         TabIndex        =   41
         Top             =   2004
         Width           =   1200
      End
      Begin VB.Label lbl医疗类别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医疗类别"
         Height          =   240
         Left            =   3570
         TabIndex        =   40
         Top             =   2460
         Width           =   960
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "监护人"
         Height          =   300
         Index           =   22
         Left            =   7020
         TabIndex        =   39
         Top             =   3495
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&R)"
      Height          =   420
      Left            =   5430
      TabIndex        =   31
      Top             =   7725
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   7350
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7725
      Width           =   1500
   End
   Begin VB.PictureBox picAddInfo 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   11190
      ScaleHeight     =   825
      ScaleWidth      =   1755
      TabIndex        =   62
      Top             =   3690
      Visible         =   0   'False
      Width           =   1755
      Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
         Height          =   435
         Left            =   330
         TabIndex        =   63
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
   Begin XtremeSuiteControls.TabControl tabPage 
      Height          =   1905
      Left            =   11400
      TabIndex        =   61
      Top             =   510
      Visible         =   0   'False
      Width           =   1755
      _Version        =   589884
      _ExtentX        =   3096
      _ExtentY        =   3360
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmDistRoomPatiEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mlngModul As Long
Public mstrNO As String '挂号单号
Public mlng挂号ID As Long  '挂号ID
Public mlng病人ID As Long
Public mlng险类 As Long
Public mstr医保号 As String  '存在险类时,不加载病人医保号,信息会导致,修改病案后医保号丢失
Public m就诊卡号 As String
Public m验证密码 As String
Public mbytType As Byte         '传递给存储过程,以确定保存类型
Public mlngOutModeMC As Long '本地医保设置的外挂式医保险类
Public mstrPrivs As String
Public mblnChange As Boolean
Public mstr姓名 As String
Public mstr性别 As String
Public mstr年龄 As String
Public mstr病人_姓名 As String '病人信息中的姓名
Public mstr病人_性别 As String '病人信息中的性别
Public mstr病人_年龄 As String '病人信息中的年龄
Public mstr出生日期 As String '出生日期
Public mbln医嘱业务 As Boolean  '是否发生了医嘱业务
Private mbln基本信息调整 As Boolean '是否调整有医嘱业务的病人基本信息
Public mstrPrivsPubPatient As String
Public mblnStructAdress As Boolean  '病人地址结构化录入
Public mblnShowTown As Boolean      '乡镇地址结构化录入

Private mrs家庭地址 As ADODB.Recordset  '缓存家庭地址,初始时读取地区表
Private mstrSQL As String
Private mDateSys As Date

Private Enum mIndex
    idx_监护人 = 0
    idx_身高 = 1
    idx_体重 = 2
    idx_体温 = 3
End Enum

Private mobjPlugIn As Object '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
Private mlngPlugInHwnd As Long
Private Enum mPageIndex
    病人信息 = 1
    附加信息 = 2
End Enum
Private mobjPubPatient As Object
Private mblnGetBirth As Boolean '判断是否允许通过年龄计算生日

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
Private Sub Load家庭地址()
    Dim strSQL As String, strFile As String
    Dim fld As Field
    Dim fso As Scripting.FileSystemObject
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    strFile = App.Path & "\ZLAddressForRegEvent.Adtg"
    
    Set mrs家庭地址 = New ADODB.Recordset
    
    On Error Resume Next
    If fso.FileExists(strFile) Then
        mrs家庭地址.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '仅Update时才锁定
    End If
    Err.Clear
    On Error GoTo errH
    
    If mrs家庭地址.State = 0 Then
        strSQL = "Select '系统' as 类别,名称,简码,1 as 次数 From 地区"
        Call zlDatabase.OpenRecordset(mrs家庭地址, strSQL, Me.Caption)            '必须是adUseClient才能建索引
        
        If Not mrs家庭地址.EOF Then
            '创建索引:名称,简码
            Set fld = mrs家庭地址.Fields(1)
            fld.Properties("Optimize") = True
            Set fld = mrs家庭地址.Fields(2)
            fld.Properties("Optimize") = True
            
            If fso.FileExists(strFile) Then
                Kill strFile
            End If
            mrs家庭地址.Save strFile, adPersistADTG
        End If
        mrs家庭地址.Close
        mrs家庭地址.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '仅Update时才锁定
        
    End If
    
    lbl家庭地址.ToolTipText = "请定期备份本机[家庭地址]数据文件:" & strFile
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
        If Not mrs家庭地址 Is Nothing Then
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
        If mrs家庭地址 Is Nothing Then Exit Sub
        
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
    
    If mrs家庭地址 Is Nothing Then Exit Sub
    
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

Private Sub cbo民族_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo民族.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo民族.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo民族.ListIndex = lngIdx
End Sub

Private Sub cbo年龄单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo年龄单位_LostFocus()
    If cbo年龄单位.Tag <> cbo年龄单位.Text Then
        mblnChange = False
        If mblnGetBirth Then
            txt出生日期.Text = ReCalcBirth(Trim(txt年龄.Text), IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""))
        End If
        mblnChange = True
    End If
    '69026,冉俊明,2014-8-8,检查输入年龄
    '76703,冉俊明,2014-8-15
    If mobjPubPatient Is Nothing Then Exit Sub
    If mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text) & cbo年龄单位.Text, _
            IIf(txt出生日期.Text = "____-__-__", "", txt出生日期.Text) & _
            IIf(txt出生时间.Text = "__:__", "", " " & txt出生时间.Text)) = False Then
        If txt年龄.Visible And txt年龄.Enabled Then txt年龄.SetFocus
    End If
End Sub

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo性别.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo性别.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo性别.ListIndex = lngIdx
    If cbo性别.ListIndex = -1 And cbo性别.ListCount > 0 Then cbo性别.ListIndex = 0
    
End Sub

Private Sub cbo医疗类别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo医疗类别.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo医疗类别.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo医疗类别.ListIndex = lngIdx
End Sub

Private Sub cbo职业_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo职业.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo职业.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo职业.ListIndex = lngIdx
End Sub

Private Sub cmdCancel_Click()
    gblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, blnTrans As Boolean, lngTmp As Long
    Dim strDate As String, str就诊 As String, strMCAccount As String, strTmp As String
    Dim str年龄 As String, str出生日期 As String
    Dim cllPro As Collection
    Dim blnPlugInCheck As Boolean
    Dim strErrMsg As String
    On Error GoTo errH

    txtPatient.Text = Trim(txtPatient.Text)
    txt年龄.Text = Trim(txt年龄.Text)
    
    If CheckValied = False Then Exit Sub

    strMCAccount = Trim(txtPatiMCNO(0).Text)
    If mlngOutModeMC = 920 And strMCAccount <> txtPatiMCNO(0).Tag And strMCAccount <> "" Then
        strMCAccount = UCase(strMCAccount)
        If CheckExistsMCNO(strMCAccount) Then
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Sub
        End If
    End If
    If mlng险类 > 0 And strMCAccount = "" Then
        strMCAccount = mstr医保号
    End If
    
    If txt出生时间 = "__:__" Then
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & " " & txt出生时间.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If

    str年龄 = Trim(IIf(IsNumeric(txt年龄.Text), txt年龄.Text & cbo年龄单位.Text, txt年龄.Text))
    
    If Me.Caption Like "创建挂号*" Then
        '更新原挂号单据上的病人ID为新的ID
        mlng病人ID = zlDatabase.GetNextNo(1)
        mstr病人_姓名 = txtPatient.Text
        mstr病人_性别 = NeedName(cbo性别.Text)
        mstr病人_年龄 = txt年龄.Text & IIf(cbo年龄单位.Visible, cbo年龄单位, "")
    End If
    
    '更新病人信息缓存值
    If Not mbln医嘱业务 Or InStr(mstrPrivsPubPatient, ";基本信息调整;") > 0 Then
        mstr病人_姓名 = txtPatient.Text
        mstr病人_性别 = NeedName(cbo性别.Text)
        mstr病人_年龄 = txt年龄.Text & IIf(cbo年龄单位.Visible, cbo年龄单位, "")
    End If
    
    '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    If Not mobjPlugIn Is Nothing And mlngPlugInHwnd <> 0 Then '保存插件附加信息前的数据有效性检查
        On Error Resume Next
        blnPlugInCheck = mobjPlugIn.PatiInfoSaveBefore(mlng病人ID)
        Call zlPlugInErrH(Err, "PatiInfoSaveBefore")
        If Err = 0 And blnPlugInCheck = False Then
            Exit Sub '检查未通过终止保存
        End If
        Err.Clear: On Error GoTo errH
    End If
    
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')"
    Set cllPro = New Collection
    '-----------------------------
    mstrSQL = "zl_挂号病人病案_INSERT(" & mbytType & "," & mlng病人ID & "," & txt门诊号.Text & "," & _
              "'" & m就诊卡号 & "','" & m验证密码 & "'," & _
              "'" & mstr病人_姓名 & "','" & mstr病人_性别 & "','" & mstr病人_年龄 & "'," & _
              "'" & NeedName(cbo费别.Text) & "','" & NeedName(cbo付款方式.Text) & "'," & _
              "'" & NeedName(cbo国籍.Text) & "','" & NeedName(cbo民族.Text) & "','" & NeedName(cbo婚姻.Text) & "'," & _
              "'" & NeedName(cbo职业.Text, True) & "','" & txt身份证号.Text & "','" & txt单位名称.Text & "'," & _
              Val(txt单位名称.Tag) & ",'" & txt单位电话.Text & "','" & txt单位邮编.Text & "'," & _
              "'" & IIf(mblnStructAdress, padd家庭地址.Value, cbo家庭地址.Text) & "'," & _
              "'" & txt家庭电话.Text & "','" & txt家庭邮编.Text & "'," & strDate & _
              ",'" & mstrNO & "'," & str出生日期 & ",'" & strMCAccount & "',Null," & IIf(mlng险类 = 0, "NULL", mlng险类) & ","
    '            区域_In           病人信息.区域%Type := Null,
    mstrSQL = mstrSQL & "Null,"
    '  户口地址_In       病人信息.户口地址%Type := Null,
    mstrSQL = mstrSQL & "'" & IIf(mblnStructAdress, padd户口地址.Value, Trim(txt户口地址.Text)) & "',"
    '  户口邮编_In   病人信息.户口邮编%Type := Null,
    mstrSQL = mstrSQL & "'" & Trim(txt户口邮编.Text) & "',"
    '  联系人身份证号_In In 病人信息.联系人身份证号%Type := Null,
    mstrSQL = mstrSQL & "Null,"
    '  联系人姓名_In     In 病人信息.联系人姓名%Type := Null,
    mstrSQL = mstrSQL & "Null,"
    '  联系人电话_In     In 病人信息.联系人电话%Type := Null,
    mstrSQL = mstrSQL & "Null,"
    '  联系人关系_In     In 病人信息.联系人关系%Type := Null,
    mstrSQL = mstrSQL & "Null,"
    '  监护人_In         In 病人信息.监护人%Type := Null
    mstrSQL = mstrSQL & "'" & Trim(txtEdit(idx_监护人).Text) & "'" & ")"
    zlAddArray cllPro, mstrSQL
    
    '89242:李南春,2015/12/10,更新病人地址信息
    If mblnStructAdress Then
        If padd家庭地址.Value <> "" Then
           mstrSQL = "zl_病人地址信息_update(1," & mlng病人ID & ",NULL,3,'" & padd家庭地址.value省 & "','" & _
               padd家庭地址.value市 & "','" & padd家庭地址.value区县 & "','" & padd家庭地址.value乡镇 & "','" & _
               padd家庭地址.value详细地址 & "','" & padd家庭地址.Code & "')"
        Else
           mstrSQL = "zl_病人地址信息_update(2," & mlng病人ID & ",NULL,3)"
        End If
        zlAddArray cllPro, mstrSQL
        
        If padd户口地址.Value <> "" Then
           mstrSQL = "zl_病人地址信息_update(1," & mlng病人ID & ",NULL,4,'" & padd户口地址.value省 & "','" & _
               padd户口地址.value市 & "','" & padd户口地址.value区县 & "','" & padd户口地址.value乡镇 & "','" & _
               padd户口地址.value详细地址 & "','" & padd户口地址.Code & "')"
        Else
           mstrSQL = "zl_病人地址信息_update(2," & mlng病人ID & ",NULL,4)"
        End If
        zlAddArray cllPro, mstrSQL
    End If

    mstrSQL = "ZL_挂号费用信息_Update('" & mstrNO & "'," & mlng病人ID & "," & txt门诊号.Text & "," & _
              "'" & txtPatient.Text & "','" & NeedName(cbo性别.Text) & "','" & str年龄 & "'," & _
              "'" & NeedName(cbo费别.Text) & "')"
    zlAddArray cllPro, mstrSQL
    If mlngOutModeMC > 0 And cbo医疗类别.ListIndex > 0 Then
        If IsDate(cbo医疗类别.Tag) Then strDate = "To_Date('" & cbo医疗类别.Tag & "','YYYY-MM-DD HH24:MI:SS')"
        str就诊 = cbo医疗类别.Text
        str就诊 = Mid(str就诊, 1, InStr(1, str就诊, "-") - 1)
        mstrSQL = "zl_就诊登记记录_UPDATE(" & mlngOutModeMC & "," & mlng病人ID & ",0," & strDate & ",0,'" & str就诊 & "')"
        zlAddArray cllPro, mstrSQL
    End If
    '67070:刘尔旋,2013-11-04,获取写入病人体征记录的SQL
    mstrSQL = UCPatiVitalSigns.GetSaveSQL(mlng病人ID, mlng挂号ID)
    If mstrSQL <> "" Then zlAddArray cllPro, mstrSQL
    
    '89149:李南春,2015/10/12,病人过敏药物记录
    '过敏药物
    With msh过敏
        If .Rows > 1 Then
            '清除该病人所有记录
            mstrSQL = " Zl_病人过敏药物_Delete(" & mlng病人ID & ")"
            zlAddArray cllPro, mstrSQL
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    '病人过敏药物
                    mstrSQL = "Zl_病人过敏药物_Update("
                    '病人ID_In 病人过敏药物.病人Id%Type
                    mstrSQL = mstrSQL & "" & mlng病人ID & ","
                    '过敏药物ID_In 病人过敏药物.过敏药物ID%Type
                    mstrSQL = mstrSQL & "'" & IIf(.RowData(i) <= 0, "", .RowData(i)) & "',"
                    '过敏药物_In  病人过敏药物.过敏药物%Type
                    mstrSQL = mstrSQL & "'" & IIf(.TextMatrix(i, 0) = "", "", .TextMatrix(i, 0)) & "',"
                    '过敏反应_In 病人过敏反应.过敏反应%Type
                    mstrSQL = mstrSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"

                    zlAddArray cllPro, mstrSQL
                End If
            Next
        End If
    End With
    
    '81103,冉俊明,2014-12-26,录入身份证号后,出生日期、年龄、性别的同步关联检查和调整
    If mbln基本信息调整 Then
        Call mobjPubPatient.SavePatiBaseInfo(mlng病人ID, mlng挂号ID, Trim(txtPatient.Text), _
                                    NeedName(cbo性别.Text), str年龄, txt出生日期.Text, "门诊分诊", 1, strErrMsg, , True)
    End If
    '执行存储过程
    zlExecuteProcedureArrAy cllPro, Me.Caption
    
    '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    If Not mobjPlugIn Is Nothing And mlngPlugInHwnd <> 0 Then  '保存插件附加信息
        On Error Resume Next
        Call mobjPlugIn.PatiInfoSaveAfter(mlng病人ID)
        Call zlPlugInErrH(Err, "PatiInfoSaveAfter")
        Err.Clear: On Error GoTo 0
    End If
    
    '只有正确保存后才刷新
    gblnOk = True
    Unload Me
    Exit Sub
errSQL:
    gcnOracle.RollbackTrans
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
 
Private Sub cmd单位名称_Click()
    Call SearchUnit("", txt单位名称)
End Sub


Private Sub cmd户口地址_Click()
    Call SearchAddress("", txt户口地址)
End Sub

Private Sub Form_Activate()
    '78408:李南春,2014/10/9,光标跳转
    If Me.ActiveControl Is msh过敏 Then Exit Sub
    If txtPatient.Enabled Then txtPatient.SetFocus
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
        Else
            cmdCancel_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        '89242:李南春,2015/12/10,PatiAddress控件内部处理了跳转，外部不再处理
        If UCase(TypeName(Me.ActiveControl)) = UCase("PatiAddress") Then Exit Sub
        If InStr(1, "lvwItems,txt年龄,cbo年龄单位,txt出生日期,msh过敏,txt过敏,txtPatiMCNO", Me.ActiveControl.Name) <= 0 Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Resize()
    If tabPage.Visible Then
        tabPage.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 600
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    If Not mobjPlugIn Is Nothing Then Set mobjPlugIn = Nothing
    mlngPlugInHwnd = 0
    
    If Not mrs家庭地址 Is Nothing Then
        If mrs家庭地址.State = 1 Then
            On Error Resume Next
            Kill App.Path & "\ZLAddressForRegEvent.Adtg"
            Err.Clear
            mrs家庭地址.Filter = ""
            mrs家庭地址.Save App.Path & "\ZLAddressForRegEvent.Adtg"
        End If
    End If
    Set mrs家庭地址 = Nothing
    
    mlng病人ID = 0
    mstrNO = ""
    mlng险类 = 0
    mstr医保号 = ""
    m就诊卡号 = ""
    m验证密码 = ""
    mbytType = 0
    mblnChange = False
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
    
    If KeyCode = vbKeyF2 Then msh过敏_DblClick
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
            zlControl.TxtSelAll txt过敏
            txt过敏.Visible = True
            If txt过敏.Visible And txt过敏.Enabled Then txt过敏.SetFocus
        Case 1 '过敏反应
            txt过敏反应.Top = msh过敏.CellTop + msh过敏.Top + (msh过敏.CellHeight - txt过敏反应.Height) / 2 - 15
            txt过敏反应.Left = msh过敏.Left + msh过敏.CellLeft + 30
            txt过敏反应.Width = 3000
            
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

    If msh过敏.Row > 1 And msh过敏.TextMatrix(msh过敏.Row - 1, 0) = "" Or msh过敏.Col = 1 Then Exit Sub
    
    cmd过敏.Top = msh过敏.CellTop + msh过敏.Top - 15
    cmd过敏.Left = msh过敏.Left + msh过敏.CellWidth - cmd过敏.Width + 45
    
    cmd过敏.ZOrder
    cmd过敏.Visible = True
End Sub

Private Sub msh过敏_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        'If msh过敏.Row > 1 And msh过敏.TextMatrix(msh过敏.Row - 1, 0) = "" Or msh过敏.RowData(msh过敏.Row) = 0 Then Exit Sub
        msh过敏_DblClick
        If msh过敏.Col = 0 Then
            msh过敏.RowData(msh过敏.Row) = 0
            txt过敏.Text = Chr(KeyAscii)
            txt过敏.SelStart = Len(txt过敏.Text)
        Else
            txt过敏反应.Text = Chr(KeyAscii)
            txt过敏反应.SelStart = Len(txt过敏反应.Text)
        End If
    Else
        If msh过敏.Row = msh过敏.Rows - 1 And msh过敏.TextMatrix(msh过敏.Row, msh过敏.Col) <> "" Then
            msh过敏.Rows = msh过敏.Rows + 1
            msh过敏.Row = msh过敏.Rows - 1
            
            msh过敏_EnterCell
        ElseIf msh过敏.TextMatrix(msh过敏.Row, msh过敏.Col) <> "" Then
            msh过敏.Row = msh过敏.Row + 1
            msh过敏_EnterCell
        Else
            cmdOK.SetFocus
        End If
    End If
End Sub

Private Sub cmd过敏_Click()
On Error GoTo errH
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
    
    Set rsTmp = frmPubSel.ShowSelect(Me, strSQL, 2, "过敏药物", , msh过敏.Text, "请从下面的药品中选择一项作为病人过敏药物。")
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
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd家庭地址_Click()
On Error GoTo errH
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
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function InitData() As Boolean
On Error GoTo errH
'功能：初始化必要数据
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim objCtl As Control
    Dim strSQL As String, lngTmp As Long
    
    If mlng病人ID = 0 Then
        Me.Caption = "创建挂号病人信息"
    End If
    Me.txt门诊号.Enabled = False
    
    '费别
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省" & vbCrLf & _
        " From 费别 Where 属性=1 And Nvl(服务对象,3) IN(1,3)" & vbCrLf & _
        " Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            cbo费别.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo费别.ItemData(cbo费别.NewIndex) = 1
                cbo费别.ListIndex = cbo费别.NewIndex
            End If
            rsTmp.MoveNext
        Next
        cbo费别.Enabled = False
    End If
    
    
    If mlngOutModeMC > 0 Then
        Set rsTmp = GetDictData("医疗类别")
        cbo医疗类别.Clear
        cbo医疗类别.AddItem " "
        For i = 1 To rsTmp.RecordCount
            cbo医疗类别.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo医疗类别.ItemData(cbo医疗类别.NewIndex) = 1
            End If
            rsTmp.MoveNext
        Next
        cbo医疗类别.ListIndex = 0
        Call zlControl.CboSetWidth(cbo付款方式.Hwnd, txtPatiMCNO(0).Width)
        
    Else
        lblPatiMCNO(0).Visible = False: lblPatiMCNO(1).Visible = False
        txtPatiMCNO(0).Visible = False: txtPatiMCNO(1).Visible = False
        lbl医疗类别.Visible = False: cbo医疗类别.Visible = False
        
        lngTmp = txtPatiMCNO(0).Height / 6
        '79352,冉俊明,2015-1-6,病人地址(户口地址与现住址)调整
        lbl年龄.Top = lbl年龄.Top + lngTmp: txt年龄.Top = txt年龄.Top + lngTmp: cbo年龄单位.Top = cbo年龄单位.Top + lngTmp
        lbl电话.Top = lbl电话.Top + lngTmp: txt家庭电话.Top = txt家庭电话.Top + lngTmp
'        lbl性别.Top = lbl性别.Top + lngTmp: cbo性别.Top = cbo性别.Top + lngTmp
        lbl出生日期.Top = lbl出生日期.Top + lngTmp: txt出生日期.Top = txt出生日期.Top + lngTmp: txt出生时间.Top = txt出生时间.Top + lngTmp
        
        lbl家庭地址.Top = lbl家庭地址.Top + 2 * lngTmp: cbo家庭地址.Top = cbo家庭地址.Top + 2 * lngTmp
        cmd家庭地址.Top = cmd家庭地址.Top + 2 * lngTmp
        lbl邮编.Top = lbl邮编.Top + 2 * lngTmp: txt家庭邮编.Top = txt家庭邮编.Top + 2 * lngTmp
        
        lbl户口地址.Top = lbl户口地址.Top + 3 * lngTmp: txt户口地址.Top = txt户口地址.Top + 3 * lngTmp
        cmd户口地址.Top = cmd户口地址.Top + 3 * lngTmp
        lbl户口邮编.Top = lbl户口邮编.Top + 3 * lngTmp: txt户口邮编.Top = txt户口邮编.Top + 3 * lngTmp
        
        lbl付款方式.Top = lbl付款方式.Top - 3 * lngTmp: cbo付款方式.Top = cbo付款方式.Top - 3 * lngTmp
        lbl费别.Top = lbl费别.Top - 3 * lngTmp: cbo费别.Top = cbo费别.Top - 3 * lngTmp
        lbl费别.Left = lbl民族.Left: cbo费别.Left = cbo民族.Left
        
        Frame1.Top = Frame1.Top - lngTmp
    End If
    
    '性别
    Set rsTmp = GetDictData("性别")
    cbo性别.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo性别.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo性别.ItemData(cbo性别.NewIndex) = 1
                cbo性别.ListIndex = cbo性别.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    
    '年龄单位
    cbo年龄单位.Clear
    cbo年龄单位.AddItem "岁"
    cbo年龄单位.AddItem "月"
    cbo年龄单位.AddItem "天"
    cbo年龄单位.ListIndex = 0
    mDateSys = zlDatabase.Currentdate
    
    If Not mblnStructAdress Then Call Load家庭地址

    '医疗付款方式
    Set rsTmp = GetDictData("医疗付款方式")
    cbo付款方式.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo付款方式.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo付款方式.ItemData(cbo付款方式.NewIndex) = 1
                cbo付款方式.ListIndex = cbo付款方式.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    cbo付款方式.Enabled = False

    '国籍
    Set rsTmp = GetDictData("国籍")
    cbo国籍.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo国籍.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo国籍.ItemData(cbo国籍.NewIndex) = 1
                cbo国籍.ListIndex = cbo国籍.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    '民族
    Set rsTmp = GetDictData("民族")
    cbo民族.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo民族.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo民族.ItemData(cbo民族.NewIndex) = 1
                cbo民族.ListIndex = cbo民族.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    '婚姻状况
    Set rsTmp = GetDictData("婚姻状况")
    cbo婚姻.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo婚姻.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo婚姻.ItemData(cbo婚姻.NewIndex) = 1
                cbo婚姻.ListIndex = cbo婚姻.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    '职业
    Set rsTmp = GetDictData("职业")
    cbo职业.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo职业.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo职业.ItemData(cbo职业.NewIndex) = 1
                cbo职业.ListIndex = cbo职业.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
     
    If mlng病人ID = 0 Then
        Me.cbo付款方式.Enabled = True
        Me.cbo费别.Enabled = True
    End If
    Call SetPatiBaseInforEnabled
    
    '初始化地址控件
    If mblnStructAdress Then
        padd家庭地址.Visible = mblnStructAdress: padd家庭地址.ShowTown = mblnShowTown
        cbo家庭地址.Visible = False: cmd家庭地址.Visible = False
        padd家庭地址.Top = cbo家庭地址.Top: padd家庭地址.Left = cbo家庭地址.Left
        padd户口地址.Visible = mblnStructAdress: padd户口地址.ShowTown = mblnShowTown
        txt户口地址.Visible = False: cmd户口地址.Visible = False
        padd户口地址.Top = txt户口地址.Top: padd户口地址.Left = txt户口地址.Left
    End If
    
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub SetPatiBaseInforEnabled()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置病人的基本信息(姓名,性别,年龄,出生日期)的Eanbeld
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-11 10:40:42
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean
    Dim lngColor As Long
    
    blnEdit = mlng病人ID = 0
    If mlng挂号ID <> 0 Then
        '发生了医嘱业务数据,不能调整病人基本信息
        blnEdit = Not mbln医嘱业务
        'Not zlExistOperationData(mlng病人ID, mstrNO, mlng挂号ID)
    End If
    lngColor = IIf(blnEdit = False, Me.BackColor, Me.txt门诊号.BackColor)
    
    txtPatient.Enabled = blnEdit
    cbo性别.Enabled = blnEdit
    txt年龄.Enabled = blnEdit
    cbo年龄单位.Enabled = blnEdit
    txt出生日期.Enabled = blnEdit
    txt出生时间.Enabled = blnEdit
    txtPatient.BackColor = lngColor
    cbo性别.BackColor = lngColor
    txt年龄.BackColor = lngColor
    cbo年龄单位.BackColor = lngColor
    txt出生日期.BackColor = lngColor
    txt出生时间.BackColor = lngColor
End Sub



Private Sub msh过敏_Scroll()
    cmd过敏.Visible = False
    txt过敏.Visible = False
    txt过敏反应.Visible = False
End Sub

 

Private Sub picAddInfo_Resize()
    wndTaskPanel.Move 0, 0, picAddInfo.Width, picAddInfo.Height
End Sub

Private Sub txtEdit_GotFocus(index As Integer)
    Call zlControl.TxtSelAll(txtEdit(index))
End Sub

Private Sub txtEdit_KeyPress(index As Integer, KeyAscii As Integer)
        Dim strMask As String
        If KeyAscii = 8 Then Exit Sub
        If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab: Exit Sub
        Select Case index
            Case idx_身高, idx_体温, idx_体重
                strMask = "1234567890."
            Case Else
                strMask = ""
        End Select
        If strMask <> "" Then
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
End Sub

Private Sub txtPatiMCNO_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtPatiMCNO_Validate(index As Integer, Cancel As Boolean)
    
    txtPatiMCNO(index).Text = Trim(txtPatiMCNO(index).Text)
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

Private Sub txt出生时间_Change()
    Dim str出生时间 As String
    '76669，李南春,2014-8-18,病人年龄更新
    If IsDate(txt出生日期.Text) Then
        str出生时间 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        txt年龄.Text = ReCalcOld(CDate(str出生时间), cbo年龄单位)
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
    If IsDate(txt出生日期.Text) And mblnChange Then
        mblnChange = False
        txt出生日期.Text = Format(CDate(txt出生日期.Text), "yyyy-mm-dd") '0002-02-02自动转换为2002-02-02,否则,看到的是2002,实际值却是0002
        mblnChange = True
        
        str出生时间 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        txt年龄.Text = ReCalcOld(CDate(str出生时间), cbo年龄单位)
        mblnGetBirth = False
    End If
End Sub
Private Sub txt出生日期_GotFocus()
    zlControl.TxtSelAll txt出生日期
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

Private Sub txt出生日期_LostFocus()
    If txt出生日期.Text <> "____-__-__" And Not IsDate(txt出生日期.Text) Then
        txt出生日期.SetFocus
    End If
End Sub

Private Sub txt单位电话_GotFocus()
    zlControl.TxtSelAll txt单位电话
End Sub

Private Sub txt单位电话_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt单位电话, KeyAscii
End Sub

Private Sub txt单位名称_Change()
    txt单位名称.Tag = ""
End Sub

Private Sub txt单位名称_GotFocus()
    zlControl.TxtSelAll txt单位名称
    OpenIme gstrIme
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
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Sub txt单位邮编_GotFocus()
    zlControl.TxtSelAll txt单位邮编
End Sub

Private Sub txt单位邮编_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt单位邮编, KeyAscii
End Sub

Private Sub txt过敏_Change()
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

        '1.如果是新输入的,显示查找结果的下拉列表
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
            If .BOF Or .EOF Then Exit Sub
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
'        Else
'            '2.如果当前单元格已有内容,则作为编辑处理
'            msh过敏.TextMatrix(msh过敏.Row, 0) = txt过敏.Text
'            If msh过敏.Row + 1 <= msh过敏.Rows - 1 Then msh过敏.Row = msh过敏.Row + 1
'            msh过敏.SetFocus
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt过敏_LostFocus()
    txt过敏.Visible = False
End Sub

Private Sub txt过敏反应_Change()
   '问题号:56599
   msh过敏.TextMatrix(msh过敏.Row, 1) = txt过敏反应.Text
End Sub

Private Sub txt过敏反应_LostFocus()
    txt过敏反应.Visible = False
End Sub

Private Sub Form_Load()
        
    mblnChange = True
    txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
    
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
    
    '73935,冉俊明,20114-7-3,将渠道定制的界面嵌入到病人信息编辑中
    Call CreateObjectPlugIn
    If Not mobjPlugIn Is Nothing Then
        On Error Resume Next
        mlngPlugInHwnd = mobjPlugIn.GetFormHwnd
        Call zlPlugInErrH(Err, "GetFormHwnd")
        Err.Clear: On Error GoTo 0
        If mlngPlugInHwnd <> 0 Then
            tabPage.Visible = True: Me.Height = Me.Height + 350
            cmdOK.Top = cmdOK.Top + 330: cmdCancel.Top = cmdOK.Top
            Call InitTagPage
            Call InitTaskPanel
        End If
    End If
    
    '创建病人信息公共部件
    '69026,冉俊明,2014-8-8,检查输入年龄
    Call CreatePublicPatient
    '获取病人信息公共模块权限
    mstrPrivsPubPatient = ";" & GetPrivFunc(glngSys, 9003) & ";"
    mbln基本信息调整 = False
    padd家庭地址.MaxLength = glngMax家庭地址
    padd户口地址.MaxLength = glngMax户口地址
    txt户口地址.MaxLength = glngMax户口地址
End Sub

Private Sub txt户口地址_Change()
    mblnChange = True
    txt户口地址.Tag = ""
End Sub

Private Sub txt户口地址_GotFocus()
    Call zlControl.TxtSelAll(txt户口地址)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt户口地址_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Trim(txt户口地址.Text) <> "" Then
        Call SearchAddress(Trim(txt户口地址.Text), txt户口地址)
    End If
End Sub

Private Sub txt户口地址_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub SearchAddress(ByVal strInput As String, txtInput As Object)
    '--------------------------------------------------------------
    '功能:模糊查找，弹出地区选择列表
    '编制:冉俊明
    '日期:2015-1-6
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
    txtInput.SelStart = Len(txtInput.Text)
    txtInput.SetFocus
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt户口邮编_Change()
    mblnChange = True
End Sub

Private Sub txt户口邮编_GotFocus()
    Call zlControl.TxtSelAll(txt户口邮编)
End Sub

Private Sub txt户口邮编_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
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
    zlControl.TxtSelAll txt年龄
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
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub



Private Sub txt年龄_Validate(Cancel As Boolean)
    txt年龄.Text = Trim(txt年龄.Text)
    
    If Not IsNumeric(txt年龄.Text) And Trim(txt年龄.Text) <> "" Then
        cbo年龄单位.ListIndex = -1: cbo年龄单位.Visible = False
    ElseIf cbo年龄单位.Visible = False Then
        cbo年龄单位.ListIndex = 0: cbo年龄单位.Visible = True
    End If
    If Not IsDate(txt出生日期.Text) Then mblnGetBirth = True
    mblnChange = False
    If mblnGetBirth Then
        txt出生日期.Text = ReCalcBirth(Trim(txt年龄.Text), IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""))
    End If
    mblnChange = True
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

Private Sub txt身份证号_GotFocus()
    zlControl.TxtSelAll txt身份证号
End Sub

Private Sub txt身份证号_KeyPress(KeyAscii As Integer)
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
    zlControl.TxtSelAll txtPatient
    OpenIme gstrIme
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txtPatient, KeyAscii
End Sub

Public Sub ClearFace()
    Dim i As Integer
    
    txt门诊号.Text = ""
    SetCboDefault cbo费别
    SetCboDefault cbo性别
    If mlngOutModeMC > 0 Then
        txtPatiMCNO(0).Text = ""
        txtPatiMCNO(0).Tag = "" '用于修改时判断是否已存在
        txtPatiMCNO(1).Text = ""
        If cbo医疗类别.ListIndex >= 0 Then cbo医疗类别.ListIndex = 0
    End If
    
    txtPatient.Text = ""
    txt年龄.Text = ""
    Call zlControl.CboLocate(cbo年龄单位, "岁")
    
    SetCboDefault cbo付款方式
    SetCboDefault cbo国籍
    SetCboDefault cbo民族
    SetCboDefault cbo婚姻
    SetCboDefault cbo职业
    
    txt身份证号.Text = ""
    
    txt单位名称.Text = ""
    txt单位名称.Tag = ""
    txt单位电话.Text = ""
    txt单位邮编.Text = ""
    
    cbo家庭地址.Text = ""
    txt家庭邮编.Text = ""
    padd家庭地址.Value = ""
    txt家庭电话.Text = ""
    txt户口地址.Text = ""
    padd户口地址.Value = ""
    For i = 1 To msh过敏.Rows - 1
        msh过敏.TextMatrix(i, 0) = ""
        msh过敏.RowData(i) = 0
    Next
End Sub

Private Sub txtPatient_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Function GetDictData(strDict As String) As ADODB.Recordset
'功能：从指定的字典中读取数据
'参数：strDict=字典对应的表名
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub Set医保付款()
    Dim i As Integer
    For i = 0 To cbo付款方式.ListCount - 1
        If Left(cbo付款方式.List(i), InStr(cbo付款方式.List(i), "-") - 1) = "1" Then
            cbo付款方式.ListIndex = i: Exit Sub
        End If
    Next
End Sub

Public Function GetRegBillID() As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
     On Error GoTo Errhand
     If mstrNO = "" Then Exit Function
        '基本信息
    strSQL = "Select A.病人ID,B.ID as 挂号ID,B.摘要,B.复诊,a.籍贯,A.病人类型," & _
        " Nvl(Nvl(B.续诊科室ID,Decode(B.转诊状态,1,B.转诊科室ID,NULL)),B.执行部门ID) as 科室ID," & _
        " B.传染病上传,B.发病时间,A.险类,A.门诊号,A.姓名,A.性别,A.年龄,A.出生日期,A.医疗付款方式," & _
        " A.国籍,A.民族,A.婚姻状况,A.职业,A.身份证号,A.出生地点,A.监护人,A.家庭地址,A.家庭电话," & _
        " A.区域,A.家庭地址邮编,A.工作单位,A.合同单位id,A.单位电话,A.单位邮编,B.社区,C.社区号,A.其他证件,A.户口地址,a.户口地址邮编" & _
        " From 病人信息 A,病人挂号记录 B,病人社区信息 C" & _
        " Where A.病人ID=B.病人ID And B.病人ID=C.病人ID(+) And B.社区=C.社区(+) And B.NO=[1] And B.记录性质=1 And B.记录状态=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNO)
    If rsTmp.EOF Then Exit Function
    mlng挂号ID = Val(Nvl(rsTmp!挂号ID))
    '74428：李南春，2014-7-8，病人姓名显示颜色处理
    Call SetPatiColor(txtPatient, Nvl(rsTmp!病人类型), IIf(IsNull(rsTmp!险类), Me.ForeColor, vbRed))
    Set rsTmp = Nothing
    GetRegBillID = True
Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub txt身份证号_Validate(Cancel As Boolean)
    '65663:刘尔旋,2014-02-20,根据身份证号计算出生日期
'    If IsDate(zlCommFun.GetIDCardDate(txt身份证号.Text)) = False Then Exit Sub
'    If Format(zlCommFun.GetIDCardDate(txt身份证号.Text), "yyyy-mm-dd") <> Format(txt出生日期.Text, "yyyy-mm-dd") Then
'        MsgBox "输入的身份证号与输入的出生日期不一致，将使用身份证号获取的日期替换！", vbInformation, gstrSysName
'        txt出生日期.Text = zlCommFun.GetIDCardDate(txt身份证号.Text)
'    End If
End Sub

Private Function CreateObjectPlugIn() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建渠道附加信息插件
    '返回:创建成功,返回True,否则返回False
    '问题号:73935
    '编制:冉俊明
    '日期:2014-07-3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear: On Error GoTo 0
    End If
    
    If Not mobjPlugIn Is Nothing Then
        On Error Resume Next
        Call mobjPlugIn.Initialize(gcnOracle, glngSys, 1113)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
    End If
    CreateObjectPlugIn = True
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

Private Sub InitTagPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化分页控件
    '问题号:73935
    '编制:冉俊明
    '日期:2014-07-4
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    
    Err = 0: On Error GoTo Errhand:

    Set ObjItem = tabPage.InsertItem(mPageIndex.病人信息, "病人信息", picPatiInfo.Hwnd, 0)
    ObjItem.Tag = mPageIndex.病人信息

    If Not mobjPlugIn Is Nothing Then
        If mlngPlugInHwnd <> 0 Then
            picAddInfo.Visible = True
            Set ObjItem = tabPage.InsertItem(mPageIndex.附加信息, "附加信息", picAddInfo.Hwnd, 0)
            ObjItem.Tag = mPageIndex.附加信息
        End If
    End If
        
    With tabPage
        tabPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        Set .PaintManager.Font = lbl出生日期.Font
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function InitTaskPanel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载附加信息页面
    '返回:
    '问题号:73935
    '编制:冉俊明
    '日期:2014-07-3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup, Item As TaskPanelGroupItem
    
    Err = 0: On Error GoTo Errhand
    If Not mobjPlugIn Is Nothing Then
        If mlngPlugInHwnd <> 0 Then
            With wndTaskPanel
                Call .SetGroupInnerMargins(0, 0, 0, 0)
                Call .SetGroupOuterMargins(-1, -24, -1, -1)
                
                Set tkpGroup = .Groups.Add(1, "附加信息")
                tkpGroup.CaptionVisible = False
                tkpGroup.Expandable = False
                tkpGroup.Expanded = True
                
                Set Item = tkpGroup.Items.Add(1, "", xtpTaskItemTypeControl)
                Call HideFormCaption(mlngPlugInHwnd, False) '隐藏窗体边框
                Item.Handle = mlngPlugInHwnd
                
                .HotTrackStyle = xtpTaskPanelHighlightItem
                .Reposition
                .DrawFocusRect = True
            End With
        End If
    End If

    InitTaskPanel = True
    
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Private Function CreatePublicPatient() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建zlPublicPatient部件
    '返回:创建成功,返回True,否则返回False
    '编制:冉俊明
    '日期:2014-08-08
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

Public Sub Init过敏药物()
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
        .ColWidth(1) = .Width - 5100
        '75286:李南春，2014-7-16，表格对齐方式
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
    End With
End Sub

Private Function CheckValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人信息是否合法
    '返回:病人信息合法,返回True,否则返回False
    '编制:焦博
    '日期:2017-07-03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSimilar As String
    Dim str出生日期 As String
    Dim strbirthday As String, strAge As String, strSex As String, strErrInfo As String, strInfo As String
    On Error GoTo Errhand
    
    If CheckTextLength("姓名", txtPatient) = False Then Exit Function
    If CheckTextLength("年龄", txt年龄) = False Then Exit Function
    If mblnStructAdress Then
        If Not CheckStructAddr(padd家庭地址, padd家庭地址.MaxLength) Then Exit Function
        If Not CheckStructAddr(padd户口地址, padd户口地址.MaxLength) Then Exit Function
    Else
         If zlCommFun.ActualLen(cbo家庭地址.Text) > glngMax家庭地址 Then
            MsgBox "现住址输入过长，只允许输入" & glngMax家庭地址 & "个字符或" & glngMax家庭地址 \ 2 & "个汉字，请检查!", vbInformation, gstrSysName
            cbo家庭地址.SetFocus: Exit Function
        End If
        If CheckTextLength("户口地址", txt户口地址) = False Then Exit Function
    End If
    
    If Trim(txtPatient.Text) = "" Then
        MsgBox "必须输入病人姓名，请检查！", vbInformation, gstrSysName
        Call zlControl.ControlSetFocus(txtPatient): Exit Function
    End If
    If Trim(txtPatient.Text) <> mstr姓名 Then
        If MsgBox("您即将把病人名字【" & mstr姓名 & "】修改为【" & Trim(txtPatient.Text) & "】,是否继续?", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Call zlControl.ControlSetFocus(txtPatient): Exit Function
        End If
    End If
    
    If mbytType = 1 Then
        '检查相似病人信息(新增之前检查,以免加入了重复信息！！！)
        If Trim(txt身份证号.Text) <> "" Then
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
            '新病人或调整本次挂号无业务数据的已有病人信息时提示是否调整不一致的基本信息
            If strSex <> NeedName(cbo性别.Text) Then strInfo = "性别"
            If strAge <> Trim(txt年龄.Text) & cbo年龄单位 Then strInfo = strInfo & IIf(strInfo = "", "年龄", "、年龄")
            If Format(strbirthday, "yyyy-mm-dd") <> txt出生日期.Text Then strInfo = strInfo & IIf(strInfo = "", "出生日期", "、出生日期")
            
            If strInfo <> "" Then
                If Not mbln医嘱业务 Or InStr(mstrPrivsPubPatient, ";基本信息调整;") > 0 Then
                    If MsgBox("输入的" & strInfo & "与身份证号的" & strInfo & "不一致，" & _
                            "将根据身份证号修改" & strInfo & "，是否继续？", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                        Call zlControl.CboLocate(cbo性别, strSex)
                        txt年龄.Text = ReCalcOld(CDate(strbirthday), cbo年龄单位)
                        txt出生日期.Text = Format(strbirthday, "yyyy-mm-dd")
                        '只有病人发生医嘱业务，操作员有“基本信息调整”权限，且基础信息不一致时操作员选择继续，才单独调用SavePatiBaseInfo接口
                        mbln基本信息调整 = mbln医嘱业务 And InStr(mstrPrivsPubPatient, ";基本信息调整;") > 0
                    Else
                        Exit Function
                    End If
                Else
                    If MsgBox("输入的" & strInfo & "与身份证号的" & strInfo & "不一致，是否继续？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            End If
        Else
            MsgBox strErrInfo, vbInformation, gstrSysName
            If txt身份证号.Enabled And txt身份证号.Visible Then txt身份证号.SetFocus
            Exit Function
        End If
    End If
    
    If txtPatiMCNO(0).Text <> "" Or txtPatiMCNO(1).Text <> "" Then
        If txtPatiMCNO(0).Text <> txtPatiMCNO(1).Text Then
            MsgBox "请检查,两次输入的医保号不一致！", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
        If zlCommFun.ActualLen(txtPatiMCNO(0).Text) > txtPatiMCNO(0).MaxLength Then
            MsgBox "请检查,医保号最大长度不能超过" & txtPatiMCNO(0).MaxLength & "个字符！", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
        If cbo医疗类别.ListIndex <= 0 Then
            MsgBox "请确定医保病人的医疗类别！", vbInformation, gstrSysName
            If cbo医疗类别.Visible Then cbo医疗类别.SetFocus
            Exit Function
        End If
    Else
        If cbo医疗类别.ListIndex > 0 Then
            MsgBox "选定医疗类别时必须输入医保号！", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(txt出生日期.Text) Then
        '76669，李南春,2014-8-15,年龄与出生日期检查
        str出生日期 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        If CDate(str出生日期) > zlDatabase.Currentdate Then
            If MsgBox("出生时间：" & str出生日期 & " 超过了当前系统时间。" & _
                vbCrLf & vbCrLf & "请检查年龄或出生日期的正确性 ，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If txt出生日期.Enabled And txt出生日期.Visible Then txt出生日期.SetFocus
                Exit Function
            End If
        End If
    End If
    
    '69026,冉俊明,2014-8-11,年龄有效性检查
    '76703,冉俊明,2014-8-15
    If txt年龄.Enabled And txt年龄.Visible Then
        If mobjPubPatient Is Nothing Then Exit Function
        If mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), _
                IIf(txt出生日期.Text = "____-__-__", "", txt出生日期.Text) & _
                IIf(txt出生时间.Text = "__:__", "", " " & txt出生时间.Text)) = False Then
            txt年龄.SetFocus: Exit Function
        End If
    End If
    
    If Not Me.Caption Like "创建挂号*" Then
        '75909
        If mlng挂号ID <> 0 And mbln医嘱业务 And InStr(mstrPrivsPubPatient, ";基本信息调整;") = 0 Then
            If mstr姓名 <> txtPatient.Text _
                Or mstr性别 <> NeedName(cbo性别.Text) _
                Or mstr年龄 <> txt年龄.Text & cbo年龄单位 _
                Or mstr出生日期 <> txt出生日期.Text Then
                MsgBox "该病人已经产生了医嘱数据,不允许调整病人的基本信息(姓名,性别,年龄等),请在『病人信息管理』中进行调整。", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    If mlng挂号ID = 0 Then
        If Not GetRegBillID() Then
            MsgBox "无法获取挂号ID", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CheckValied = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
