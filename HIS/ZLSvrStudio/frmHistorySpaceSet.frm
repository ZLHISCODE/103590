VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHistorySpaceSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "历史数据空间设置"
   ClientHeight    =   4890
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7800
   Icon            =   "frmHistorySpaceSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraSetup 
      Height          =   4170
      Index           =   0
      Left            =   0
      TabIndex        =   55
      Top             =   -120
      Visible         =   0   'False
      Width           =   8280
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   0
         Left            =   0
         TabIndex        =   57
         Top             =   465
         Width           =   8385
      End
      Begin VB.Frame fra 
         Caption         =   "历史数据空间的用户"
         Height          =   2955
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   1080
         Width           =   6570
         Begin VB.CommandButton cmd连接 
            Caption         =   "测试(&T)"
            Height          =   350
            Left            =   4080
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2470
            Width           =   1635
         End
         Begin VB.TextBox txtDBLink 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1185
            TabIndex        =   12
            Text            =   "ZLHDLink"
            Top             =   1770
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.OptionButton optServer 
            Caption         =   "远程服务器"
            Height          =   255
            Index           =   1
            Left            =   1185
            TabIndex        =   8
            Top             =   1100
            Width           =   1215
         End
         Begin VB.OptionButton optServer 
            Caption         =   "当前服务器"
            Height          =   255
            Index           =   0
            Left            =   1185
            TabIndex        =   7
            Top             =   800
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.TextBox txtDbaServer 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1185
            MaxLength       =   200
            TabIndex        =   10
            ToolTipText     =   $"frmHistorySpaceSet.frx":058A
            Top             =   1410
            Width           =   1635
         End
         Begin VB.TextBox txtDba口令 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4065
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   360
            Width           =   1635
         End
         Begin VB.TextBox txtDba用户 
            Height          =   300
            Left            =   1185
            MaxLength       =   100
            TabIndex        =   3
            Top             =   360
            Width           =   1635
         End
         Begin VB.CommandButton cmd升级 
            Caption         =   "历史数据升级(&U)"
            Height          =   350
            Left            =   1080
            TabIndex        =   15
            Top             =   2470
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label lblDBLinkPrompt 
            Caption         =   "Oracle不支持通过DBLink操作含有XMLType等对象类型或用户定义类型字段的表，所以，不支持直接转出到远程历史库"
            ForeColor       =   &H00404040&
            Height          =   855
            Left            =   4000
            TabIndex        =   114
            Top             =   1200
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label lblPWDPrompt 
            Caption         =   "登录数据库的密码，无需转换"
            Height          =   255
            Index           =   1
            Left            =   4080
            TabIndex        =   113
            Top             =   795
            Width           =   2415
         End
         Begin VB.Label lblServerName 
            AutoSize        =   -1  'True
            Caption         =   "DBLink名称"
            Height          =   180
            Index           =   1
            Left            =   165
            TabIndex        =   11
            Top             =   1830
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lblServer 
            Caption         =   "位置"
            Height          =   255
            Left            =   705
            TabIndex        =   6
            Top             =   830
            Width           =   375
         End
         Begin VB.Label lblIniModi 
            AutoSize        =   -1  'True
            Caption         =   "修改…"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   5640
            TabIndex        =   14
            Top             =   2160
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lblSetupIni 
            AutoSize        =   -1  'True
            Caption         =   "安装配置文件：C:\Appsoft\ZLHIS10\应用脚本\ZLSETUP.INI"
            Height          =   180
            Left            =   240
            TabIndex        =   13
            Top             =   2160
            Visible         =   0   'False
            Width           =   4770
         End
         Begin VB.Label lblServerName 
            AutoSize        =   -1  'True
            Caption         =   "服务器名"
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   9
            Top             =   1470
            Width           =   720
         End
         Begin VB.Label lblDba 
            AutoSize        =   -1  'True
            Caption         =   "口令"
            Height          =   180
            Index           =   1
            Left            =   3580
            TabIndex        =   4
            Top             =   420
            Width           =   360
         End
         Begin VB.Label lblDba 
            AutoSize        =   -1  'True
            Caption         =   "用户名"
            Height          =   180
            Index           =   2
            Left            =   540
            TabIndex        =   2
            Top             =   420
            Width           =   540
         End
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmHistorySpaceSet.frx":0632
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lbl 
         Caption         =   "   服务器是指本机连接到指定的历史数据空间的服务连接串,该串需在本机Oracle的TnsNames文件中配置。"
         Height          =   405
         Index           =   13
         Left            =   2400
         TabIndex        =   90
         Top             =   2520
         Width           =   4380
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "    配置数据库服务器连接信息。"
         Height          =   180
         Index           =   0
         Left            =   780
         TabIndex        =   0
         Top             =   750
         Width           =   2700
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "第一步 指定DBA用户"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   56
         Top             =   225
         Width           =   2055
      End
   End
   Begin VB.Frame fraImport 
      Height          =   4155
      Left            =   -30
      TabIndex        =   78
      Top             =   -120
      Visible         =   0   'False
      Width           =   8340
      Begin VB.CheckBox chk当前 
         Caption         =   "置为当前历数据空间(&D)"
         Height          =   270
         Left            =   2160
         TabIndex        =   115
         Top             =   3525
         Width           =   2325
      End
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   3
         Left            =   -30
         TabIndex        =   79
         Top             =   570
         Width           =   8415
      End
      Begin VB.TextBox txtMoveName 
         Height          =   300
         Left            =   2160
         TabIndex        =   85
         Top             =   1755
         Width           =   2460
      End
      Begin VB.TextBox txtMoveCode 
         Height          =   300
         Left            =   2160
         TabIndex        =   83
         Top             =   1260
         Width           =   1305
      End
      Begin VB.TextBox txtMoveUser 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   87
         Top             =   2250
         Width           =   2460
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   360
         Picture         =   "frmHistorySpaceSet.frx":29A4
         Top             =   840
         Width           =   480
      End
      Begin VB.Label lblBakVer 
         AutoSize        =   -1  'True
         Caption         =   "备份:"
         Height          =   180
         Left            =   5205
         TabIndex        =   89
         Top             =   1995
         Width           =   450
      End
      Begin VB.Label lblDataVer 
         AutoSize        =   -1  'True
         Caption         =   "在线:"
         Height          =   180
         Left            =   5205
         TabIndex        =   88
         Top             =   1560
         Width           =   450
      End
      Begin VB.Shape shap 
         BorderStyle     =   3  'Dot
         Height          =   1320
         Left            =   4965
         Top             =   1215
         Width           =   2535
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "空间名称"
         Height          =   180
         Index           =   9
         Left            =   1365
         TabIndex        =   84
         Top             =   1830
         Width           =   720
      End
      Begin VB.Label lblNoteImport 
         Caption         =   "    设置被植入的历史数据空间的编号信息及空间名称。"
         Height          =   330
         Left            =   1425
         TabIndex        =   81
         Top             =   855
         Width           =   5955
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "空间用户"
         Height          =   180
         Index           =   5
         Left            =   1365
         TabIndex        =   86
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   180
         Index           =   6
         Left            =   1725
         TabIndex        =   82
         Top             =   1305
         Width           =   360
      End
      Begin VB.Label lblStepImport 
         AutoSize        =   -1  'True
         Caption         =   "第二步 设置历史数据空间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   80
         Top             =   240
         Width           =   2595
      End
   End
   Begin VB.Frame fraSetup 
      Height          =   3960
      Index           =   1
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Visible         =   0   'False
      Width           =   8280
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   1
         Left            =   15
         TabIndex        =   62
         Top             =   570
         Width           =   8415
      End
      Begin TabDlg.SSTab tbHistory 
         Height          =   3015
         Left            =   270
         TabIndex        =   17
         Top             =   810
         Width           =   7230
         _ExtentX        =   12753
         _ExtentY        =   5318
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "常规"
         TabPicture(0)   =   "frmHistorySpaceSet.frx":5D86
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblNewLab"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblNewPwd"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lbl(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lbl(1)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lbl(3)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblIn"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblLinkName"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lbl(12)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label5"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lbl(15)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Image3"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "lblBakPrompt"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "lblPWDPrompt(0)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txtHD"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txt编号"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "txtOwnerLab"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txtOwnerPwd"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "txtOwnerUsr"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "chkCreate当前"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).ControlCount=   19
         TabCaption(1)   =   "数据文件"
         TabPicture(1)   =   "frmHistorySpaceSet.frx":5DA2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblSpaceExtentSize"
         Tab(1).Control(1)=   "lblSpaceExtend"
         Tab(1).Control(2)=   "lblDataFile"
         Tab(1).Control(3)=   "lblFileSize"
         Tab(1).Control(4)=   "lblBakSpace"
         Tab(1).Control(5)=   "Image2"
         Tab(1).Control(6)=   "lblBakSpaceIdx"
         Tab(1).Control(7)=   "lblFileAmount(0)"
         Tab(1).Control(8)=   "lblFileAmount(1)"
         Tab(1).Control(9)=   "lblBakSpaceLob"
         Tab(1).Control(10)=   "lblFileAmount(2)"
         Tab(1).Control(11)=   "txtSpaceExtentSize"
         Tab(1).Control(12)=   "cboSpaceExtentType"
         Tab(1).Control(13)=   "txtDataFile"
         Tab(1).Control(14)=   "chkSpaceExtd"
         Tab(1).Control(15)=   "txtSpaceSize"
         Tab(1).Control(16)=   "txtBakSpace"
         Tab(1).Control(17)=   "txtBakSpaceIdx"
         Tab(1).Control(18)=   "txtFileAmount(0)"
         Tab(1).Control(19)=   "txtFileAmount(1)"
         Tab(1).Control(20)=   "txtBakSpaceLob"
         Tab(1).Control(21)=   "txtFileAmount(2)"
         Tab(1).ControlCount=   22
         Begin VB.TextBox txtFileAmount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   -69855
            MaxLength       =   2
            TabIndex        =   51
            Text            =   "3"
            Top             =   2460
            Width           =   300
         End
         Begin VB.TextBox txtBakSpaceLob 
            BackColor       =   &H00F0F0E0&
            Height          =   300
            Left            =   -72735
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   2460
            Width           =   2160
         End
         Begin VB.TextBox txtFileAmount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   -69855
            MaxLength       =   2
            TabIndex        =   47
            Text            =   "3"
            Top             =   2040
            Width           =   300
         End
         Begin VB.TextBox txtFileAmount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   -72735
            MaxLength       =   2
            TabIndex        =   36
            Text            =   "4"
            Top             =   1200
            Width           =   300
         End
         Begin VB.TextBox txtBakSpaceIdx 
            BackColor       =   &H00F0F0E0&
            Height          =   300
            Left            =   -72735
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   2040
            Width           =   2160
         End
         Begin VB.CheckBox chkCreate当前 
            Caption         =   "创建后置为当前空间(&C)"
            Height          =   270
            Left            =   1440
            TabIndex        =   29
            Top             =   2400
            Width           =   2295
         End
         Begin VB.TextBox txtBakSpace 
            Height          =   300
            Left            =   -72735
            TabIndex        =   32
            Top             =   450
            Width           =   2160
         End
         Begin VB.TextBox txtOwnerUsr 
            BorderStyle     =   0  'None
            Height          =   220
            Left            =   1900
            MaxLength       =   27
            TabIndex        =   21
            Text            =   "201312"
            Top             =   1120
            Width           =   1500
         End
         Begin VB.TextBox txtOwnerPwd 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1485
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   23
            Top             =   1491
            Width           =   1560
         End
         Begin VB.TextBox txtOwnerLab 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1485
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   25
            Top             =   1944
            Width           =   1560
         End
         Begin VB.TextBox txtSpaceSize 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   -70650
            MaxLength       =   6
            TabIndex        =   38
            Text            =   "500"
            Top             =   1185
            Width           =   750
         End
         Begin VB.CheckBox chkSpaceExtd 
            Caption         =   "自动扩展空间"
            Height          =   270
            Left            =   -69480
            TabIndex        =   39
            ToolTipText     =   "AUTOEXTEND ON Next (表空间大小/10)M"
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1425
         End
         Begin VB.TextBox txtDataFile 
            Height          =   300
            Left            =   -72735
            TabIndex        =   34
            Top             =   825
            Width           =   4680
         End
         Begin VB.ComboBox cboSpaceExtentType 
            Height          =   300
            Left            =   -72735
            Style           =   2  'Dropdown List
            TabIndex        =   41
            ToolTipText     =   "AUTOALLOCATE 或 UNIFORM Size nM"
            Top             =   1605
            Width           =   2160
         End
         Begin VB.TextBox txtSpaceExtentSize 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   270
            Left            =   -70200
            MaxLength       =   2
            TabIndex        =   42
            Text            =   "1"
            Top             =   1620
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.TextBox txt编号 
            Height          =   300
            Left            =   1485
            MaxLength       =   5
            TabIndex        =   19
            Top             =   675
            Width           =   840
         End
         Begin VB.TextBox txtHD 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1485
            TabIndex        =   77
            Text            =   "ZLHD"
            Top             =   1083
            Width           =   2160
         End
         Begin VB.Label lblPWDPrompt 
            Caption         =   "登录数据库的密码，无需转换"
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   112
            Top             =   1605
            Width           =   3150
         End
         Begin VB.Label lblBakPrompt 
            Caption         =   "建议按转出截止年月命名,例如:201412"
            Height          =   255
            Left            =   3720
            TabIndex        =   111
            Top             =   1106
            Width           =   3150
         End
         Begin VB.Label lblFileAmount 
            AutoSize        =   -1  'True
            Caption         =   "共创建     个文件"
            Height          =   180
            Index           =   2
            Left            =   -70440
            TabIndex        =   50
            Top             =   2520
            Width           =   1530
         End
         Begin VB.Label lblBakSpaceLob 
            Alignment       =   1  'Right Justify
            Caption         =   "大对象表空间名"
            Height          =   225
            Left            =   -74205
            TabIndex        =   48
            Top             =   2505
            Width           =   1365
         End
         Begin VB.Label lblFileAmount 
            AutoSize        =   -1  'True
            Caption         =   "共创建     个文件"
            Height          =   180
            Index           =   1
            Left            =   -70440
            TabIndex        =   46
            Top             =   2100
            Width           =   1530
         End
         Begin VB.Label lblFileAmount 
            AutoSize        =   -1  'True
            Caption         =   "共创建     个文件"
            Height          =   180
            Index           =   0
            Left            =   -73380
            TabIndex        =   35
            Top             =   1260
            Width           =   1530
         End
         Begin VB.Label lblBakSpaceIdx 
            Alignment       =   1  'Right Justify
            Caption         =   "索引表空间名"
            Height          =   225
            Left            =   -73965
            TabIndex        =   44
            Top             =   2100
            Width           =   1125
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   -74760
            Picture         =   "frmHistorySpaceSet.frx":5DBE
            Stretch         =   -1  'True
            Top             =   600
            Width           =   510
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   240
            Picture         =   "frmHistorySpaceSet.frx":6E40
            Top             =   600
            Width           =   480
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "说明"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   15
            Left            =   90
            TabIndex        =   92
            Top             =   3870
            Width           =   390
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "   服务器是指在线服务器连接到备份数据服务器中的连接串,在线服务器的机器上必需配置该串!"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   60
            TabIndex        =   91
            Top             =   3960
            Width           =   7650
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "服务器"
            Height          =   180
            Index           =   12
            Left            =   1545
            TabIndex        =   28
            Top             =   4335
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lblBakSpace 
            Alignment       =   1  'Right Justify
            Caption         =   "数据表空间名"
            Height          =   225
            Left            =   -73965
            TabIndex        =   31
            Top             =   510
            Width           =   1125
         End
         Begin VB.Label lblLinkName 
            AutoSize        =   -1  'True
            Caption         =   "@"
            ForeColor       =   &H8000000C&
            Height          =   180
            Left            =   5310
            TabIndex        =   27
            Top             =   3960
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblIn 
            Caption         =   "lt"
            ForeColor       =   &H8000000D&
            Height          =   390
            Left            =   60
            TabIndex        =   30
            Top             =   4125
            Width           =   7500
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "远程连接名"
            Height          =   180
            Index           =   3
            Left            =   1185
            TabIndex        =   26
            Top             =   3975
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "用户名"
            Height          =   180
            Index           =   1
            Left            =   870
            TabIndex        =   20
            Top             =   1143
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "编号"
            Height          =   180
            Index           =   0
            Left            =   1050
            TabIndex        =   18
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lblNewPwd 
            AutoSize        =   -1  'True
            Caption         =   "口令"
            Height          =   180
            Left            =   1050
            TabIndex        =   22
            Top             =   1551
            Width           =   360
         End
         Begin VB.Label lblNewLab 
            AutoSize        =   -1  'True
            Caption         =   "验证"
            Height          =   180
            Left            =   1035
            TabIndex        =   24
            Top             =   2004
            Width           =   360
         End
         Begin VB.Label lblFileSize 
            AutoSize        =   -1  'True
            Caption         =   "初始大小          M"
            Height          =   180
            Left            =   -71430
            TabIndex        =   37
            Top             =   1245
            Width           =   1710
         End
         Begin VB.Label lblDataFile 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "第一个文件"
            Height          =   180
            Left            =   -73740
            TabIndex        =   33
            Top             =   885
            Width           =   900
         End
         Begin VB.Label lblSpaceExtend 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "区尺寸"
            Height          =   180
            Left            =   -73380
            TabIndex        =   40
            Top             =   1665
            Width           =   540
         End
         Begin VB.Label lblSpaceExtentSize 
            Caption         =   "M"
            Height          =   255
            Left            =   -69800
            TabIndex        =   43
            Top             =   1680
            Visible         =   0   'False
            Width           =   135
         End
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "第二步 设置历史数据空间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   63
         Top             =   225
         Width           =   2595
      End
   End
   Begin VB.Frame fraTrans 
      Height          =   4065
      Left            =   -30
      TabIndex        =   93
      Top             =   -120
      Visible         =   0   'False
      Width           =   8250
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   5
         Left            =   30
         TabIndex        =   94
         Top             =   450
         Width           =   8415
      End
      Begin VB.TextBox txtBakPWD 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3720
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   109
         Top             =   3480
         Width           =   1530
      End
      Begin MSComctlLib.ImageList imgSys 
         Left            =   240
         Top             =   1560
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistorySpaceSet.frx":A222
               Key             =   "Other"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistorySpaceSet.frx":B2B4
               Key             =   "Run"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistorySpaceSet.frx":E6A6
               Key             =   "Lock"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistorySpaceSet.frx":11A98
               Key             =   "LockAndRun"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwHistory 
         Height          =   1665
         Left            =   1080
         TabIndex        =   108
         Top             =   720
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   2937
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgSys"
         SmallIcons      =   "imgSys"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "编号"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "名称"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "当前"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "编号"
            Text            =   "只读"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "所有者"
            Text            =   "所有者"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "版本号"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "最后转储日期"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "最后复制日期"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblBakPWD 
         AutoSize        =   -1  'True
         Caption         =   "目标库新建的历史空间用户密码"
         Height          =   180
         Left            =   1080
         TabIndex        =   110
         Top             =   3540
         Width           =   2520
      End
      Begin VB.Label Label2 
         Caption         =   "说明："
         Height          =   375
         Left            =   360
         TabIndex        =   97
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblStepTrans 
         AutoSize        =   -1  'True
         Caption         =   "第二步：选择源服务器上的历史数据空间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   96
         Top             =   225
         Width           =   4050
      End
      Begin VB.Label lblNoteTrans 
         Caption         =   "源历史数据空间的限制条件"
         Height          =   930
         Left            =   960
         TabIndex        =   95
         Top             =   2520
         Width           =   6465
      End
      Begin VB.Image Image7 
         Height          =   525
         Left            =   180
         Picture         =   "frmHistorySpaceSet.frx":14E8A
         Stretch         =   -1  'True
         Top             =   720
         Width           =   540
      End
   End
   Begin MSComDlg.CommonDialog cdgPub 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   30
      TabIndex        =   60
      Top             =   4095
      Width           =   1100
   End
   Begin MSComctlLib.ProgressBar pgbState 
      Height          =   150
      Left            =   2655
      TabIndex        =   59
      Top             =   4650
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一步(&B)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4125
      TabIndex        =   54
      Top             =   4095
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6405
      TabIndex        =   53
      Top             =   4095
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   58
      Top             =   4515
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10186
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "18:06"
            Key             =   "STANUM"
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
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一步(&N)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5250
      TabIndex        =   52
      Top             =   4095
      Width           =   1100
   End
   Begin VB.Frame fraDelete 
      Height          =   4065
      Left            =   -30
      TabIndex        =   64
      Top             =   -120
      Visible         =   0   'False
      Width           =   8250
      Begin VB.Frame fra 
         Height          =   1680
         Index           =   1
         Left            =   870
         TabIndex        =   68
         Top             =   1140
         Width           =   6585
         Begin VB.OptionButton optDele 
            Caption         =   "剥离历史数据空间(&1)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   270
            TabIndex        =   70
            Top             =   945
            Width           =   2220
         End
         Begin VB.OptionButton optDele 
            Caption         =   "删除历史数据空间(&2)"
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   69
            Top             =   300
            Value           =   -1  'True
            Width           =   2385
         End
         Begin VB.Label lblDelInfor 
            Caption         =   "只清除所选择的历史数据空间名称，相关的历史数据不从数据库删除。"
            ForeColor       =   &H8000000C&
            Height          =   285
            Index           =   1
            Left            =   465
            TabIndex        =   72
            Top             =   1245
            Width           =   5775
         End
         Begin VB.Label lblDelInfor 
            Caption         =   "彻底从数据库中删除相关的历史数据"
            ForeColor       =   &H8000000C&
            Height          =   330
            Index           =   0
            Left            =   465
            TabIndex        =   71
            Top             =   615
            Width           =   3060
         End
      End
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   2
         Left            =   30
         TabIndex        =   65
         Top             =   450
         Width           =   8415
      End
      Begin VB.Image Image4 
         Height          =   525
         Left            =   120
         Picture         =   "frmHistorySpaceSet.frx":15F0C
         Stretch         =   -1  'True
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lblSpaceOwner 
         AutoSize        =   -1  'True
         Caption         =   "所有者:"
         Height          =   180
         Left            =   3780
         TabIndex        =   76
         Top             =   3285
         Width           =   630
      End
      Begin VB.Label lblDbLink 
         Caption         =   "DB连接：X23423"
         Height          =   180
         Left            =   1080
         TabIndex        =   75
         Top             =   3510
         Width           =   4335
      End
      Begin VB.Label lblSpace 
         Caption         =   "名称：zlbak0701"
         Height          =   180
         Left            =   1080
         TabIndex        =   74
         Top             =   3270
         Width           =   4335
      End
      Begin VB.Label lblCode 
         Caption         =   "编号：200"
         Height          =   180
         Left            =   1080
         TabIndex        =   73
         Top             =   3045
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   3  'Dot
         FillColor       =   &H00C0C0C0&
         Height          =   960
         Left            =   825
         Top             =   2880
         Width           =   6615
      End
      Begin VB.Label lblStepDelete 
         AutoSize        =   -1  'True
         Caption         =   "第一步：选择是删除或剥离历史数据空间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   67
         Top             =   225
         Width           =   4050
      End
      Begin VB.Label lblNoteDelete 
         Caption         =   "    删除历史数据空间可以减少每次备份的耗时及数据量，执行前请确保这些数据已移植到其他数据库，存在完整有效的备份。"
         Height          =   450
         Left            =   870
         TabIndex        =   66
         Top             =   735
         Width           =   6585
      End
   End
   Begin VB.Frame fraMerge 
      Height          =   4065
      Left            =   -30
      TabIndex        =   98
      Top             =   -120
      Visible         =   0   'False
      Width           =   8340
      Begin VB.TextBox txtMergeSpace 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   102
         Top             =   2130
         Width           =   4980
      End
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   4
         Left            =   -30
         TabIndex        =   101
         Top             =   570
         Width           =   8415
      End
      Begin VB.TextBox txtKeepSpaceNO 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   100
         Top             =   1260
         Width           =   1260
      End
      Begin VB.TextBox txtKeepSpaceName 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   99
         Top             =   1635
         Width           =   1260
      End
      Begin VB.Label lblStepMerge 
         AutoSize        =   -1  'True
         Caption         =   "第一步 检查合并的历史数据空间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   107
         Top             =   240
         Width           =   3270
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "保留空间编号"
         Height          =   180
         Index           =   16
         Left            =   1005
         TabIndex        =   106
         Top             =   1305
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "合并空间"
         Height          =   180
         Index           =   10
         Left            =   1365
         TabIndex        =   105
         Top             =   2190
         Width           =   720
      End
      Begin VB.Label lblNoteMerge 
         Caption         =   "    合并的历史数据空间的编号信息及空间名称。"
         Height          =   450
         Left            =   960
         TabIndex        =   104
         Top             =   840
         Width           =   5955
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "保留空间名称"
         Height          =   180
         Index           =   8
         Left            =   1005
         TabIndex        =   103
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   360
         Picture         =   "frmHistorySpaceSet.frx":16F8E
         Top             =   840
         Width           =   480
      End
   End
   Begin ComctlLib.ImageList ist 
      Left            =   1440
      Top             =   4875
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmHistorySpaceSet.frx":1A370
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHistorySpaceSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSys  As Long                   '系统编号
Private mstrSysName   As String                 '系统名称
Private mstrOwnerName As String
Private mstrVersion   As String                 '版本号
Private mstrOwnerPass As String
Private mstrDBLink As String                    '历史空间的DBLink

Private mcnDBA As New ADODB.Connection      'DBA用户或历史空间用户的连接
Private mcnOracle As New ADODB.Connection   '当前登录用户的连接对象

Private mblnFirst As Boolean
Private mblnSucced As Boolean
Private mlngOracleVer As Long

Private Enum ENUFT
    F0创建 = 0
    F1拆卸 = 1
    F2再植 = 2
    F3复制 = 3
    F4切换 = 4
    F5合并 = 5
    F6转移 = 6
End Enum
Private mintFunType        As ENUFT  '0-创建历史数据空间,1-拆卸历史数据空间,2-再植历史数据空间,3-复制非转储数据,
                                     '4－切换在当前的历史数据空间,5-合并历史数据空间,6-转移历史数据空间

Private Enum ENUCOL
    C0编号 = 0
    C1名称 = 1
    C2当前 = 2
    C3只读 = 3
    C4所有者 = 4
    C5版本号 = 5
    C6最后转储日期 = 6
    C7最后复制日期 = 7
End Enum

Private mlng空间编号          As Long
Private mblnMustInstall As Boolean  '必需安装空间
Private mstr合并空间编号 As String
Private mrsMergeSpace As ADODB.Recordset
Private mblnSysUpdateCall As Boolean '是否系统升级调用


Public Function ShowInstall(ByVal frmMain As Form, ByVal cnOracle As ADODB.Connection, _
    ByVal strOwner As String, ByVal strOwnerPass As String, _
    ByVal lng系统 As Long, ByVal intFunType As Integer, _
    ByVal lng空间编号 As Long, Optional ByVal str合并空间编号 As String, _
    Optional ByVal blnSysUpdateCall As Boolean) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '功能:历史数据空间管理接口
    '参数:cnOracle-系统连接
    '     strOwner-所有者用户名
    '     strOwnerPass-所有者密码
    '     lng系统-系统号
    '     intFunType-功能类型,见mintFunType
    '               其中4-切换,5-合并，进入后就开始执行，只是用这个窗体来显示进度。
    '     lng空间编号-intFunType=1时=撤卸数据空间的空间编号
    '             intFunType=5时=保留的数据空间的编号
    '             intFunType=6时=传0（不需要）
    '             blnSysUpdateCall=是否是系统升级调用，如果是，则出错可以继续往下走。
    '返回:安装成功,返回true,否则返回False
    '----------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    mintFunType = intFunType
    
    Set mcnOracle = cnOracle
    mstrOwnerPass = strOwnerPass
    mblnSysUpdateCall = blnSysUpdateCall
    mblnMustInstall = False
    mstr合并空间编号 = str合并空间编号
    
    If mintFunType = F0创建 Then
        If IsHavingHistoryTable(lng系统) = False Then
            If frmMain Is frmHistoryDataMgr Then
                MsgBox "不存在历史数据空间的相关数据表，不能创建历史数据空间,请检查!", vbInformation, gstrSysName
                Exit Function
            Else
                ShowInstall = True
                Exit Function
            End If
        Else
            If Not frmMain Is frmHistoryDataMgr Then
                mblnMustInstall = True
                cmdCancel.Enabled = False
            End If
        End If
    ElseIf mintFunType = F1拆卸 Then
        If IsHavingHistoryTable(lng系统) = False Then
            ShowInstall = True
            Exit Function
        End If
        If Not frmMain Is frmHistoryDataMgr Then
            mblnMustInstall = True
        End If
    ElseIf mintFunType = F6转移 Then
        If gstrUserName <> "SYSTEM" Then
            MsgBox "历史空间转移功能必须以SYSTEM的用户执行，当前用户不是SYSTEM，请重新登录!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    mlngSys = lng系统
    
    If mintFunType = F1拆卸 And Not frmMain Is frmHistoryDataMgr Then
        '系统撤除
              
    Else
        gstrSQL = "select 所有者,版本号,名称 from zlSystems where 编号=" & mlngSys
        Call OpenRecordset(rsTemp, gstrSQL, "读取所有者")
        If Not rsTemp.EOF Then
            mstrOwnerName = Nvl(rsTemp!所有者)
            mstrVersion = Nvl(rsTemp!版本号)
            mstrSysName = Nvl(rsTemp!名称)
        Else
            MsgBox "系统不存在,可能被他人拆卸,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        If strOwner <> mstrOwnerName And mintFunType <> F6转移 Then
            MsgBox "你不是当前应用程序的所有者,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If
    If mstrVersion <> "" Then
        If Val(Split(mstrVersion, ".")(0)) < 10 Then
                MsgBox "不支持9以下的版本,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
        End If
    End If
    mlng空间编号 = lng空间编号
        
    Me.Show 1
    ShowInstall = mblnSucced
End Function



Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = True
    
    If InitCtronl = False Then Unload Me: Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
        
    mblnFirst = True
    Call ApplyOEM(stbThis)
        
    pgbState.Top = stbThis.Top + stbThis.Height / 3
          
    lblDataVer.Caption = "在线:" & mstrVersion
    lblDataVer.Tag = mstrVersion
    
    mlngOracleVer = GetOracleVersion(True, True)
    
End Sub


Private Function InitCtronl() As Boolean
'功能:设置控件可见性和可用性及文字说明，缺省fraSetup等卡片都是不可见状态
    Dim bytErr As Byte, strErrMsg As String, strDbLink As String
    
    cmdPrevious.Enabled = False
    
    txtDbaServer.Text = gstrServer  '目前只有复制、再植、转移功能允许指定远程服务器
    txtDbaServer.Enabled = False
    
    Select Case mintFunType
    Case F0创建   '-创建空间
        Me.Caption = "创建历史数据空间"
        
        If CheckIsDBA(mcnOracle) Then
            Set mcnDBA = mcnOracle
            
            fraSetup(0).Visible = False '直接用当前连接来创建历史数据空间，不必显示DBA连接界面
            fraSetup(1).Visible = True
            
            Call InitCreateTbs
        Else
            fraSetup(0).Visible = True
            fra(0).Caption = "用于创建历史数据空间的DBA用户"
        
            If mblnMustInstall Then
                txtDba用户.Text = gstrUserName
                txtDba口令.Text = gstrPassword
            End If
            If fraSetup(0).Visible Then
                If txtDba用户.Text <> "" Then
                    If txtDba口令.Enabled And txtDba口令.Visible Then txtDba口令.SetFocus
                Else
                    If txtDba用户.Enabled And txtDba用户.Visible Then txtDba用户.SetFocus
                End If
            End If
            
            'Oracle不支持通过DBLink操作含有XMLType等对象类型或用户定义类型字段的表，所以，不支持直接操作远程历史库
            optServer(1).Enabled = False
            lblDBLinkPrompt.Visible = True
            
            '切换、再植时才需要历史库升级所需的配置文件
            lblSetupIni.Visible = False
            lblIniModi.Visible = False
        
        End If
        
        
        If mblnMustInstall Then
            chkCreate当前.value = 1
            chkCreate当前.Enabled = False
        Else
            chkCreate当前.Enabled = True
        End If
        
        '表空间区分配类型
        cboSpaceExtentType.Clear
        cboSpaceExtentType.addItem "自动分配区尺寸"
        cboSpaceExtentType.addItem "统一分配区尺寸"
        cboSpaceExtentType.ListIndex = 0
        
        txtSpaceExtentSize.Text = 1
        txtSpaceExtentSize.Enabled = (cboSpaceExtentType.ListIndex = 1)
    
        InitCtronl = True
        
    Case F1拆卸   '撤卸空间
        Me.Caption = "卸载历史数据空间"
        fraDelete.Visible = True
        
        lblDBLinkPrompt.Visible = False
                        
        If LoadSpaceData Then
            optDele(0).value = True '缺省为删除模式
            lblStep(0).Caption = "第二步:指定DBA用户"
            lblNote(0).Caption = "    设置在远程服务器上的历史数据空间的DBA用户。"
            InitCtronl = True
        End If
    Case F2再植   '再植空间
        Me.Caption = "再植历史数据空间"
        fraSetup(0).Visible = True
        
        lblStep(0).Caption = "第一步 指定历史数据空间用户"
        fra(0).Caption = "历史数据空间用户的连接信息"
        lblNote(0).Caption = "    将剥离或其他原因导致当前没有纳入管理的历史数据空间重新纳入管理。"
        lblServerName(1).Caption = "创建DBLink"
        
        txtDba用户.Text = "ZLHD"
        If txtDba用户.Enabled And txtDba用户.Visible Then txtDba用户.SetFocus

        InitCtronl = True
    Case F3复制   '复制非转出数据表及数据
        Me.Caption = "复制非转储数据"
        fraSetup(0).Visible = True
        optServer(1).Enabled = False
        
        If LoadSpaceData Then
            txtDba用户.Text = lblSpaceOwner.Tag
            If txtDba用户.Enabled And txtDba用户.Visible Then txtDba用户.SetFocus
            txtDbaServer.Enabled = True     '允许将数据复制到远程数据库，因为使用Copy命令，不存在分布式事务
                        
            lblStep(0).Caption = "第一步 指定历史数据空间用户"
            lblNote(0).Caption = "    将所有非转储表的数据复制到以下历史数据空间中(可以是远程)。"
            fra(0).Caption = "历史数据空间的用户"
            cmdNext.Caption = "复制(&F)"
                     
            InitCtronl = True
        End If
    Case F4切换
        Me.Caption = "切换当前历史数据空间"
        fraSetup(0).Visible = True
        
        If LoadSpaceData Then
            txtDba用户.Text = lblSpaceOwner.Tag
            txtDba用户.Enabled = False
            txtDba口令.Enabled = False
            cmd连接.Enabled = False
            optServer(0).Enabled = False
            optServer(1).Enabled = False
            txtDbaServer.Enabled = False
            txtDBLink.Enabled = False
                        
            If optServer(1).value = True Then
                strDbLink = Trim(txtDBLink.Text)
            End If
                         
            lblStep(0).Caption = "    切换当前历史数据空间为" & lblSpace.Tag & "。"
            fra(0).Caption = "历史数据空间的用户"
            
            '执行切换
            Call SetControlEnable(False)
            If ExeFuncChange(Trim(txtDba用户.Text), mstrOwnerName, mlngSys, bytErr, strErrMsg, strDbLink) = False Then
                Call SetControlEnable(True)
                '1-连接失效,2-系统不存在,3-在线版本大于历史版本,4-在线版本小于历史版本
                Select Case bytErr
                Case 2, 4
                    MsgBox strErrMsg, vbInformation + vbDefaultButton1, gstrSysName
                    Unload Me
                    Exit Function
                Case 3
                    MsgBox strErrMsg, vbInformation + vbDefaultButton1, gstrSysName
                    Call ReadSetupIni(1)
                    cmd升级.Visible = True
                    lblIniModi.Visible = cmd升级.Visible: lblSetupIni.Visible = cmd升级.Visible
                    cmd连接.Visible = False
                    txtDba用户.Enabled = False
                    txtDba口令.Enabled = True
                    
                    Me.cmdNext.Caption = "切换(&Q)"
                    InitCtronl = True
                    If cmd升级.Enabled And cmd升级.Visible Then
                        lblIniModi.Visible = True: lblSetupIni.Visible = True
                        cmd升级.SetFocus
                    End If
                    Exit Function
                Case 1
                    '连接不成功,需要重新输入相关的用户名和密码
                    MsgBox strErrMsg, vbInformation + vbDefaultButton1, gstrSysName
                    txtDba用户.Enabled = True
                    txtDba口令.Enabled = True
                    
                    cmd连接.Visible = False
                    
                    If optServer(1).value = True Then txtDBLink.Enabled = True
              
                End Select
            Else
                Call SetControlEnable(True)
                mblnSucced = True
            End If
            
            Me.cmdNext.Caption = "切换(&Q)"
            InitCtronl = True
            Unload Me
        End If
    Case F5合并
        Me.Caption = "合并历史数据空间"
        fraSetup(0).Visible = True
        
        'Oracle不支持通过DBLink操作含有XMLType等对象类型或用户定义类型字段的表，所以，不支持直接操作远程历史库
        optServer(1).Enabled = False
        lblDBLinkPrompt.Visible = True
        
        If LoadSpaceData Then
            lblStep(0).Caption = "第一步 指定DBA用户"
            lblNote(0).Caption = "    指定DBA用户连接信息，用于删除历史数据空间的所有者及表空间文件。"
            fra(0).Caption = "DBA用户连接信息"
        
            cmdNext.Caption = "合并(&Q)"
            
            If txtDba用户.Text <> "" Then
                If txtDba口令.Enabled And txtDba口令.Visible Then txtDba口令.SetFocus
            Else
                If txtDba用户.Enabled And txtDba用户.Visible Then txtDba用户.SetFocus
            End If
            InitCtronl = True
        End If
    Case F6转移
        Me.Caption = "转移历史数据空间"
        fraSetup(0).Visible = True
        
        lblStep(0).Caption = "第一步 指定源服务器的DBA用户"
        lblNote(0).Caption = "    连接源服务器(服务名在本机Tnsnames中存在)，版本是否相符，平台是否支持。"
        fra(0).Caption = "远程服务器连接信息"
        
        txtDba用户.Text = "SYSTEM"
        txtDba用户.Enabled = False
        txtDba用户.BackColor = &H8000000F
        If txtDba口令.Enabled And txtDba口令.Visible Then txtDba口令.SetFocus
        txtDbaServer.Enabled = True
        optServer(1).Enabled = False
        
        InitCtronl = True
    End Select
End Function

Private Sub ReadSetupIni(ByVal intIndex As Integer)
'功能：读取系统安装配置文件
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strSetup As String, strTmp As String
    
    
    '获取安装配置文件
    strSQL = "Select A.文件名 From Zlsysfiles a Where  A.操作=1 And 系统=" & mlngSys
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    If Not rsTmp.EOF Then
        If gobjFile.FileExists(rsTmp!文件名 & "") Then
            strSetup = rsTmp!文件名 & ""
        End If
    End If
    If strSetup = "" Then
        strTmp = gobjFile.GetParentFolderName(App.Path) & "\" & Decode(mlngSys \ 100, 1, "ZLHIS10", 3, "ZLMEDREC10", 4, "ZLMATERIAL10", _
                                                                                6, "ZLDEVICE10", 21, "ZLPEIS10", 22, "ZLBLOOD10", _
                                                                                23, "ZLINFECT10", 24, "ZLOPER10", _
                                                                                25, "ZLLIS10", 26, "ZLPSS10", 27, "ZLHEC10") & "\应用脚本\ZLSETUP.INI"
        If gobjFile.FileExists(strTmp) Then
            strSetup = strTmp
        End If
    End If
    If strSetup <> "" Then
        If Not CheckInitFile(mlngSys, strSetup) Then
            strSetup = ""
        End If
    End If
    lblSetupIni.Caption = "安装配置文件：" & strSetup
    lblSetupIni.Tag = strSetup
    lblSetupIni.ToolTipText = strSetup
    Call SetCtrlPosOnLine(False, 0, lblSetupIni, 60, lblIniModi)
    lblSetupIni.Refresh
    If lblSetupIni.Width >= IIf(intIndex = 0, 5500, 5100) Then
        lblSetupIni.Width = IIf(intIndex = 0, 5500, 5100)
    End If
End Sub

Private Function LoadSpaceData() As Boolean
    '-------------------------------------------------------------------------------
    '功能:加载空间数据信息
    '-------------------------------------------------------------------------------
    Dim i As Long, str空间名称 As String
    Dim rsbakspaces As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim strImgKey As String
    Dim objItem As ListItem
    Dim lngMaxLen As Long
    
    If mintFunType = F1拆卸 And mblnMustInstall Then
        gstrSQL = "Select 编号,名称,所有者,DB连接,当前 From ZLTOOLS.zlbakspaces where 系统=" & mlngSys & " and  当前=1"
    
    ElseIf mintFunType = F5合并 Then
        gstrSQL = "Select 编号,名称,所有者,DB连接,当前 From ZLTOOLS.zlbakspaces where 编号 in(" & mstr合并空间编号 & "," & mlng空间编号 & ") Order by 编号"
    
    ElseIf mintFunType = F6转移 Then
        
        gstrSQL = "Select max(length(编号)) as MaxLen From zltools.zlbakspaces where 系统=" & mlngSys
        OpenRecordset rsbakspaces, gstrSQL, "获取历史数据空间", , , mcnDBA
        lngMaxLen = Val(Nvl(rsbakspaces!MaxLen))
    
        gstrSQL = "Select 编号,名称,所有者,DB连接,当前,只读 From ZLTOOLS.zlbakspaces where 系统=" & mlngSys
    Else
        gstrSQL = "Select 编号,名称,所有者,DB连接,当前 From ZLTOOLS.zlbakspaces where 系统=" & mlngSys & " and  编号=" & mlng空间编号
    End If
    
    If mintFunType = F6转移 Then
        OpenRecordset rsbakspaces, gstrSQL, Me.Caption, , , mcnDBA
    Else
        OpenRecordset rsbakspaces, gstrSQL, Me.Caption
    End If
    
    
    
    
    If mintFunType = F5合并 Then
        If rsbakspaces.RecordCount = 0 Then
           MsgBox "历史数据空间可能已经被他人删除,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
           Exit Function
        End If
        
        For i = 1 To rsbakspaces.RecordCount
            If rsbakspaces!编号 = mlng空间编号 Then
                txtKeepSpaceName.Text = rsbakspaces!名称
                txtKeepSpaceNO.Text = rsbakspaces!编号
            Else
                str空间名称 = str空间名称 & "," & rsbakspaces!名称
            End If
            If IsNull(rsbakspaces!DB连接) = False Then
                MsgBox "Oracle不支持通过DBLink操作含有XMLType等对象类型或用户定义类型字段的表，所以，不支持对远程历史数据空间进行合并。", vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
            End If
            
            rsbakspaces.MoveNext
        Next
        txtMergeSpace.Text = Mid(str空间名称, 2)
        
        rsbakspaces.MoveFirst
        Set mrsMergeSpace = rsbakspaces
        
    ElseIf mintFunType = F6转移 Then
        If rsbakspaces.RecordCount = 0 Then
           MsgBox "历史数据空间可能已经被他人删除或者仅存在一个当前历史间,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
           Exit Function
        End If
        
        With lvwHistory
            .ListItems.Clear
            Do While Not rsbakspaces.EOF
                Set objItem = .ListItems.Add(, "K" & Nvl(rsbakspaces!编号), Lpad(Nvl(rsbakspaces!编号), lngMaxLen), 0, 0)
                
                If .SelectedItem Is Nothing Then objItem.Selected = True
       
                objItem.SubItems(C1名称) = Nvl(rsbakspaces!名称)
                objItem.SubItems(C2当前) = IIf(Val(Nvl(rsbakspaces!当前)) = 1, "√", "")
                objItem.SubItems(C3只读) = IIf(Val(Nvl(rsbakspaces!只读)) = 1, "√", "")
                objItem.SubItems(C4所有者) = Nvl(rsbakspaces!所有者)
                
                If Val(Nvl(rsbakspaces!只读)) = 1 Then
                    strImgKey = "Lock"
                Else
                    strImgKey = "Other"
                End If
                
                objItem.SmallIcon = strImgKey
                objItem.Icon = strImgKey
                
                err.Clear: On Error Resume Next
                gstrSQL = "select 系统,版本号,更新日期,最后转储日期,最后复制日期 from " & rsbakspaces!所有者 & ".ZLBAKINFO where 系统=" & mlngSys
                Set rsTmp = New ADODB.Recordset
                rsTmp.Open gstrSQL, mcnDBA, adOpenKeyset, adLockReadOnly
                '创建历史空间后，所有权限都授给了应用系统的所有者的，所以应该能访问
                If err <> 0 Then
                    MsgBox "警告:" & vbCrLf & "  历史数据空间" & rsbakspaces!名称 & "不能正常连接,请检查权限是否正常。" & vbCrLf & err.Description, vbInformation + vbDefaultButton1
                Else
                    If Not rsTmp.EOF Then
                        objItem.SubItems(C5版本号) = Nvl(rsTmp!版本号)
                        objItem.SubItems(C6最后转储日期) = Format(rsTmp!最后转储日期, "yyyy-mm-dd")
                        objItem.SubItems(C7最后复制日期) = Format(rsTmp!最后复制日期, "yyyy-mm-dd")
                    End If
                End If

                rsbakspaces.MoveNext
            Loop
            If rsbakspaces.RecordCount = 1 And err.Number <> 0 Then
                Exit Function
            End If
        End With
    
    Else
        If rsbakspaces.EOF Then
            MsgBox "历史数据空间编号为:" & mlng空间编号 & " 已经被他人删除,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
            lblCode.Caption = "编号:"
            lblSpace.Caption = "名称:"
            lblDbLink.Caption = "DB连接:"
            lblSpaceOwner.Caption = "所有者:"
            Exit Function
        End If
        
        mlng空间编号 = Val(Nvl(rsbakspaces!编号))
        lblCode.Caption = "编号:" & Nvl(rsbakspaces!编号)
        lblSpace.Caption = "名称:" & Nvl(rsbakspaces!名称)
        lblSpace.Tag = Nvl(rsbakspaces!名称)
        lblDbLink.Caption = "DB连接:" & rsbakspaces!DB连接
        mstrDBLink = "" & rsbakspaces!DB连接
    
        lblSpaceOwner.Caption = "所有者:" & Nvl(rsbakspaces!所有者)
        lblSpaceOwner.Tag = Nvl(rsbakspaces!所有者)
        
        If mintFunType = F4切换 Then
            If mstrDBLink = "" Then
                optServer(0).value = True
            Else
                optServer(1).value = True
                txtDBLink.Text = mstrDBLink
                
                On Error Resume Next
                mcnOracle.Errors.Clear
                gstrSQL = "Select 1 from dual@" & mstrDBLink
                OpenRecordset rsbakspaces, gstrSQL, Me.Caption, , , mcnOracle
                
                If err.Number <> 0 Then
                    MsgBox "历史空间的数据库链路" & txtDBLink.Text & "无法正常连接,请人工删除后重新创建。", vbExclamation, gstrSysName
                    Exit Function
                Else
                    gstrSQL = "Select HOST From All_Db_Links Where Db_Link||'.' Like '" & UCase(mstrDBLink) & ".%'"
                    OpenRecordset rsbakspaces, gstrSQL, Me.Caption, , , mcnOracle
                
                    If rsbakspaces.RecordCount > 0 Then txtDbaServer.Text = rsbakspaces!HOST
                End If
            End If
        End If
    End If
    
    LoadSpaceData = True
End Function

Private Function CheckTransCondition() As Boolean
'功能：检查传输的源和目标数据库是否符合条件
    Dim rsFrom As ADODB.Recordset, rsTo As ADODB.Recordset
    Dim strErr As String
    Dim strTemp As String
    
    '1.要求主次版本相同，修正版本允许不同(10.2.0.4-->10.2.0.1)
    gstrSQL = "Select Substr(Banner, 6, 4) As 版本 From V$version Where Banner Like 'CORE%'"
    Set rsFrom = New ADODB.Recordset
    Set rsTo = New ADODB.Recordset
    OpenRecordset rsFrom, gstrSQL, Me.Caption, , , mcnDBA
    OpenRecordset rsTo, gstrSQL, Me.Caption, , , mcnOracle
    If rsFrom!版本 <> rsTo!版本 Then
        strErr = strErr & vbCrLf & "版本差异太大,源库:" & rsFrom!版本 & ",目标库:" & rsTo!版本 & "。"
    End If
    
    '2.检查兼容版本
    gstrSQL = "Select Substr(Value, 1, 4) As 兼容版本 From V$parameter Where Name = 'compatible'"
    Set rsFrom = New ADODB.Recordset
    Set rsTo = New ADODB.Recordset
    OpenRecordset rsFrom, gstrSQL, Me.Caption, , , mcnDBA
    OpenRecordset rsTo, gstrSQL, Me.Caption, , , mcnOracle
    If rsFrom!兼容版本 <> rsTo!兼容版本 Then
        strErr = strErr & vbCrLf & "兼容版本差异太大,源库:" & rsFrom!兼容版本 & ",目标库:" & rsTo!兼容版本 & "。"
    End If
    
    '3.检查字符集
    gstrSQL = "SELECT PROPERTY_NAME, PROPERTY_VALUE" & vbNewLine & _
                "FROM DATABASE_PROPERTIES" & vbNewLine & _
                "WHERE PROPERTY_NAME ='NLS_CHARACTERSET' or PROPERTY_NAME ='NLS_NCHAR_CHARACTERSET'"
    Set rsFrom = New ADODB.Recordset
    Set rsTo = New ADODB.Recordset
    OpenRecordset rsFrom, gstrSQL, Me.Caption, , , mcnDBA
    OpenRecordset rsTo, gstrSQL, Me.Caption, , , mcnOracle
    
    rsFrom.Filter = "PROPERTY_NAME='NLS_CHARACTERSET'"
    If rsFrom!PROPERTY_VALUE <> rsTo!PROPERTY_VALUE Then
        If MsgBox("数据库字符集不同,可能导致传输失败。" & vbCrLf & "源库:" & rsFrom!PROPERTY_VALUE & ",目标库:" & rsTo!PROPERTY_VALUE & "。" & vbCrLf & "你确定要继续吗？", vbQuestion + vbOKCancel + vbDefaultButton1) = vbCancel Then
            Exit Function
        End If
    End If
    '因现有系统无NVARCHAR和NCHAR数据类型，暂不检查国家字符集
'    rsFrom.Filter = "PROPERTY_NAME='NLS_NCHAR_CHARACTERSET'"
'    If rsFrom!PROPERTY_VALUE <> rsTo!PROPERTY_VALUE Then
'        strErr = strErr & vbCrLf & "国家字符集不同,源库:" & rsFrom!PROPERTY_VALUE & ",目标库:" & rsTo!PROPERTY_VALUE & "。"
'    End If
    
    '4.检查支持转换的平台
    '目标库平台信息
    gstrSQL = "Select d.Platform_Name, Endian_Format" & vbNewLine & _
                "From V$transportable_Platform Tp, V$database D" & vbNewLine & _
                "Where Tp.Platform_Name = d.Platform_Name"
    Set rsFrom = New ADODB.Recordset
    Set rsTo = New ADODB.Recordset
    OpenRecordset rsTo, gstrSQL, Me.Caption, , , mcnOracle
    
    '查源库平台是否支持转换
    strTemp = rsTo!Platform_Name
    If InStr(strTemp, "Linux x86") > 0 Then
        If InStr(strTemp, "64") > 0 Then
            strTemp = "Linux IA (64-bit)"
        Else
            strTemp = "Linux IA (32-bit)"
        End If
    End If
    gstrSQL = "Select Platform_Id From V$transportable_Platform Where Platform_Name = '" & strTemp & "' And Endian_Format = '" & rsTo!Endian_Format & "'"
    OpenRecordset rsFrom, gstrSQL, Me.Caption, , , mcnDBA
    If rsFrom.RecordCount = 0 Then
        If MsgBox("源库不支持转换数据文件到目标库平台（" & rsTo!Platform_Name & "," & rsTo!Endian_Format & "）。" & vbCrLf & "可能导致传输失败，你确定要继续吗？", vbQuestion + vbOKCancel + vbDefaultButton1) = vbCancel Then
            Exit Function
        End If
    End If
    
    If strErr <> "" Then
        MsgBox "检查发现以下原因导致无法进行传输：" & strErr, vbExclamation, gstrSysName
        Exit Function
    End If
    
    CheckTransCondition = True
End Function

Private Sub cboSpaceExtentType_Click()
    txtSpaceExtentSize.Enabled = (cboSpaceExtentType.ListIndex = 1)
    txtSpaceExtentSize.Visible = txtSpaceExtentSize.Enabled
    lblSpaceExtentSize.Visible = txtSpaceExtentSize.Enabled
End Sub

Private Sub cmdCancel_Click()
    Dim strKey As String
    
    If mintFunType = F0创建 Then
        strKey = "未完成历史数据空间创建,真的取消吗？"
    ElseIf mintFunType = F1拆卸 Then
        strKey = "未完成历史数据空间拆卸,真的取消吗？"
    ElseIf mintFunType = F2再植 Then
        strKey = "未完成历史数据空间的再植,真的取消吗？"
    End If
    
    If mblnMustInstall And mintFunType = F0创建 Then
        MsgBox "当前系统必需安装历史数据空间后，才能正常" & vbCrLf & "使用该系统,因此不能取消操作!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    
    If strKey <> "" Then
        If MsgBox(strKey, vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    mblnSucced = False
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name
End Sub

Private Function ExeFuncCreate(ByVal strDbaName As String, ByVal strDbaPass As String, ByVal strServer As String, _
    ByVal strBakUserName As String, ByVal strBakUserPwd As String, ByVal strDbLink As String, _
    ByVal strTableSpace As String, ByVal strDataFile As String, ByVal lngSize As Long, _
    ByVal blnAutoExpent As Boolean, ByVal blnAutoAllocate As Boolean, ByVal intExtentSize As Integer, ByVal blnHaveUser As Boolean, _
    ByVal strTbsNameIdx As String, ByVal strTbsNameLob As String, _
    ByVal lngFileAmount As Long, ByVal lngFileIdxAmount As Long, ByVal lngFileLobAmount As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能:创建历史数据空间
    '参数:strDbaName-远程的dba用户名
    '     strDbaPass-远程的dba用户名的密码
    '     strServer-远程服务器
    '     strBakUserName-历史空间名
    '     strBakUserPwd-用户密码(未加密的)
    '     strDb_Link-连接名
    '     strtablespace-表空间名
    '     strDataFile-数据文件
    '     ExtentSize:统一区尺寸
    '     blnHaveUser:是否已存在用户
    '     strTbsNameIdx,strTbsNameLob:索引表空间和大对象表空间名称
    '     lngFileAmount,lngFileIdxAmount,lngFileLobAmount:数据文件、索引文件、大对象文件的数量
    '返回;成功返回true,否则返回false
    '--------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intCreate As Integer
    Dim strDba_Link As String
    Dim strFileHead As String, strFileTail As String, strError As String
    
    
    gstrSQL = "Insert Into zltools.zlbakspaces(系统, 编号, 名称, 所有者, db连接, 当前, 只读)　Values("
    gstrSQL = gstrSQL & "" & mlngSys & ","
    gstrSQL = gstrSQL & "" & Val(txt编号.Text) & ","
    gstrSQL = gstrSQL & "'" & strTableSpace & "',"
    gstrSQL = gstrSQL & "'" & strBakUserName & "',"
    gstrSQL = gstrSQL & IIf(strDbLink = "", "NULL", "'" & strDbLink & "'") & ","
    gstrSQL = gstrSQL & "0,0)"
    
    On Error Resume Next
    mcnOracle.Execute gstrSQL
    If err <> 0 Then
        MsgBox "历史空间名或编号已经存在,请检查!", vbInformation + vbDefaultButton1, mstrSysName
        Exit Function
    End If
    
    On Error GoTo errHand
    SetPromptText "正在创建表空间…"
    If blnHaveUser = False Then
        '第一步:创建历史数据空间的表空间
        '1-创建成功；2-表空间已经存在；3-创建失败
        
        intCreate = CreateTbs(strTableSpace, strDataFile, lngSize, blnAutoExpent, blnAutoAllocate, intExtentSize, lngFileAmount)
        If intCreate = 2 Or intCreate = 3 Or intCreate = 4 Then
            GoTo ErrDropLink
            Exit Function
        End If
        
        strFileHead = Mid(strDataFile, 1, InStrRev(strDataFile, ".") - 1)
        strFileTail = Mid(strDataFile, InStrRev(strDataFile, "."))
        
        strDataFile = strFileHead & "_IDX" & strFileTail
        intCreate = CreateTbs(strTbsNameIdx, strDataFile, lngSize, blnAutoExpent, blnAutoAllocate, intExtentSize, lngFileIdxAmount)
        If intCreate = 2 Or intCreate = 3 Or intCreate = 4 Then
            GoTo ErrDropLink
            Exit Function
        End If
        
        strDataFile = strFileHead & "_LOB" & strFileTail
        intCreate = CreateTbs(strTbsNameLob, strDataFile, lngSize, blnAutoExpent, blnAutoAllocate, intExtentSize, lngFileLobAmount)
        If intCreate = 2 Or intCreate = 3 Or intCreate = 4 Then
            GoTo ErrDropLink
            Exit Function
        End If
        
        '第二步:设置历史数据空间用户
        SetPromptText "正在设置历史空间用户" & strBakUserName
        
        gstrSQL = "alter user " & strBakUserName & " DEFAULT TABLESPACE " & strTableSpace
        mcnDBA.Execute gstrSQL
        
        gstrSQL = "Grant Connect,Resource,UNLIMITED TABLESPACE," & _
                " Create Table,Create Sequence,Create Role,Create User,Drop User,Create Public Synonym,Drop Public Synonym," & _
                " Alter Session,Create Session,Create Synonym,Create View,Create Database Link,Create Cluster" & _
                " to " & strBakUserName & " With Admin Option"
        mcnDBA.Execute gstrSQL
    End If
    
    '第三步:创建相关的转储数据结构
    SetPromptText "创建数据结构" & strBakUserName
    
    If CreateHistoryStru(strDbLink, strBakUserName, strTableSpace, mstrOwnerName, strTbsNameIdx, strTbsNameLob) = False Then
        '删除创建的临时连接
        GoTo ErrDropLink
        Exit Function
    End If
    
    '第四步:授权(远程数据库使用的DBA用户连接，所以无需授权)
    If strDbLink = "" Then
        Dim cnnbak As ADODB.Connection
        
        Set cnnbak = gobjRegister.GetConnection(strServer, strBakUserName, strBakUserPwd, False, MSODBC, strError, False)
        If cnnbak.State = adStateClosed Then
             '删除创建的临时连接
            MsgBox strError, vbInformation, gstrSysName
            GoTo ErrDropLink
            Exit Function
        End If
        
        If GrantBakToUser(cnnbak, mstrOwnerName) = False Then
            cnnbak.Close
            Exit Function
        End If
        cnnbak.Close
    End If
    
    ExeFuncCreate = True
    
    Exit Function
ErrDropLink:
    Exit Function
errHand:
   If MsgBox("安装失败,请检查!" & vbCrLf & "错误号:" & err.Number & vbCrLf & "错误描述:" & err.Description & vbCrLf & gstrSQL, vbRetryCancel + vbDefaultButton2 + vbQuestion) = vbRetry Then Resume
End Function

Private Function CreateDbLink(ByVal cnOracle As ADODB.Connection, ByVal strDbLinkName As String, _
            strUserName As String, strPassword As String, strServer As String, _
            strOwner As String, Optional blnDropLink As Boolean = True, Optional blnCheckLink As Boolean = True) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '功能:创建远程连接
    '参数:cnOracle-oracle连接对象
    '     strDbLinkName-远程连接名
    '     strUserName-远程用户名
    '     strPassWord-远程用户名密码
    '     strSerVer-远程连接服务串
    '     strOwner-创建连接的所有者
    '     blnDropLink-创建连接前是否选删除原连接
    '     blnCheckLink-检查连接是否正常
    '返回:连接成功,返回true,否则返回False
    '----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    If blnDropLink Then
        gstrSQL = "Select 1 From All_Db_Links Where Db_Link||'.' Like '" & UCase(strDbLinkName) & ".%'"
        OpenRecordset rsTemp, gstrSQL, Me.Caption, , , cnOracle
        
        If rsTemp.RecordCount > 0 Then
            On Error Resume Next
            gstrSQL = "drop Database Link " & strDbLinkName
            cnOracle.Execute gstrSQL
        End If
    End If
    
    cnOracle.Errors.Clear
    On Error Resume Next
        
    gstrSQL = "Create Database Link " & strDbLinkName & " Connect to " & strUserName & " Identified by " & strPassword & " Using '" & strServer & "'"
    cnOracle.Execute gstrSQL
    If err <> 0 Then
        MsgBox "创建远程连接时出错,错误信息如下:" & vbCrLf & "(" & err.Number & ") " & err.Description & vbCrLf & gstrSQL, vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    ElseIf cnOracle.Errors.Count > 0 Then
        MsgBox "创建远程连接时出错,错误信息如下:" & vbCrLf & cnOracle.Errors(0).Description & vbCrLf & gstrSQL, vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    If blnCheckLink = True Then
        '检查创建的连接是否用效
        On Error Resume Next
        cnOracle.Errors.Clear
        gstrSQL = "Select 1 from dual@" & strDbLinkName
        OpenRecordset rsTemp, gstrSQL, Me.Caption, , , cnOracle
        
        If err.Number <> 0 Then
            If InStr(1, err.Description, "ORA-02085") > 0 Then
                If CheckGlobal_name(cnOracle) = True Then
                    MsgBox "创建的远程连接名一定要与目标数据库的全局名称" & vbCrLf & "一致(请检查Global_name),创建失败!", vbInformation + vbDefaultButton1, gstrSysName
                Else
                    '不能正常连接
                    MsgBox "创建的远程连接不能正常使用,创建失败!" & vbCrLf & "(" & err.Number & ") " & err.Description, vbInformation + vbDefaultButton1, gstrSysName
                End If
            Else
                '不能正常连接
                MsgBox "创建的远程连接不能正常使用,创建失败!" & vbCrLf & "(" & err.Number & ") " & err.Description, vbInformation + vbDefaultButton1, gstrSysName
            End If
            
            On Error Resume Next
            gstrSQL = "drop Database Link " & strDbLinkName
            cnOracle.Execute gstrSQL
            Exit Function
        ElseIf cnOracle.Errors.Count > 0 Then
            If InStr(1, cnOracle.Errors(0).Description, "ORA-02085") > 0 Then
                If CheckGlobal_name(cnOracle) = True Then
                    MsgBox "创建的远程连接名一定要与目标数据库的全局名称" & vbCrLf & "一致(请检查Global_name),创建失败!", vbInformation + vbDefaultButton1, gstrSysName
                Else
                    '不能正常连接
                    MsgBox "创建的远程连接不能正常使用,创建失败!" & vbCrLf & cnOracle.Errors(0).Description, vbInformation + vbDefaultButton1, gstrSysName
                End If
            Else
                '不能正常连接
                MsgBox "创建的远程连接不能正常使用,创建失败!" & vbCrLf & cnOracle.Errors(0).Description, vbInformation + vbDefaultButton1, gstrSysName
            End If
            
            On Error Resume Next
            gstrSQL = "drop Database Link " & strDbLinkName
            cnOracle.Execute gstrSQL
            Exit Function
        
        End If
    End If
    CreateDbLink = True
    Exit Function
    
errh:
    MsgBox "创建DBLink失败：" & err.Description, vbExclamation, gstrSysName
    
End Function

Private Function CheckGlobal_name(ByVal cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------
    '功能:检查全局参数是否为true
    '参数:cnOracle-检查的数据库
    '返回:参数为true,返回true,否则false
    '---------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select VALUE  from v$parameter where name = 'global_names'"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption, , , cnOracle)
    If rsTemp.EOF Then
        CheckGlobal_name = False
    ElseIf UCase(rsTemp!value) = "TRUE" Then
        CheckGlobal_name = True
    Else
        CheckGlobal_name = False
    End If
End Function

Private Function UpdateBakInfor(ByRef cnBakOracle As ADODB.Connection, ByVal lngSys As Long, ByVal strVer As String) As Boolean
    '--------------------------------------------------------------------------------------------
    '功能:更新历史数据空间数据
    '参数:cnBakOracle-历史数据空间连接
    '     lngSys-系統号
    '     strVer-版本号
    '返回:列新成功,返回true,否则返回False
    '--------------------------------------------------------------------------------------------
    If strVer = "" Then
        gstrSQL = "Update zlbakinfo set 最后复制日期=sysdate where 系统=" & lngSys
    Else
        gstrSQL = "Update zlbakinfo set 版本号='" & strVer & "',更新日期=sysdate where 系统=" & lngSys
    End If
    
    err = 0: On Error GoTo errHand:
    cnBakOracle.Execute gstrSQL
    UpdateBakInfor = True
    Exit Function
errHand:
    MsgBox "更新历史数据空间版本号时出错,错误信息如下:" & vbCrLf & err.Description
End Function

Private Function UpdateZlBakSpace(ByRef cnOracle As ADODB.Connection, ByVal lngBakCode As Long, ByVal lngSys As Long, _
    Optional blnDelete As Boolean = False, Optional ByVal blnReadonly As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '功能:更新当前标志或删除相关历史数据空间信息
    '参数:cnOracle-更新zlBakSpaces表中的连接
    '     lngBakCode-编号
    '     lngSys-系统
    '     blnDelete-是否删除
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------
    err = 0: On Error GoTo errHand:
    
    If blnDelete Then
        gstrSQL = "delete  zltools.zlbakspaces where 系统=" & lngSys & " and 编号=" & lngBakCode
        cnOracle.Execute gstrSQL
    Else
        gstrSQL = "Update zltools.zlbakspaces set 当前=0 where 系统=" & lngSys
        cnOracle.Execute gstrSQL
        
        gstrSQL = "Update zltools.zlbakspaces set 当前=1" & IIf(blnReadonly, ",只读=1", "") & " where 系统=" & lngSys & " and 编号=" & lngBakCode
        cnOracle.Execute gstrSQL
    End If
    UpdateZlBakSpace = True
    Exit Function
errHand:
    If MsgBox(IIf(blnDelete, "删除", "更新") & " 历史数据空间出错,详细错误信息如下:" & vbCrLf & " (" & err.Number & ") " & err.Description, vbRetryCancel + vbDefaultButton1 + vbQuestion, gstrSysName) = vbRetry Then Resume
End Function

Private Function ExeFuncImport(ByRef cnOracle As ADODB.Connection, _
        ByVal lngBakCode As Long, ByVal strBakName As String, _
        ByVal strBakOwner As String, ByVal lngSys As Long, _
        Optional ByVal strDbLink As String, Optional ByVal strDBLinkUser As String, Optional ByVal strDBLinkPwd As String, Optional ByVal strDBLinkServer As String) As Boolean
    '---------------------------------------------------------------------------------------------------
    '功能:植入已经存在了的历史数据空间
    '参数:cnBakOracle-历史数据空间连接
    '     cnOracle-在线数据连接
    '     lngBakCode-历史数据空间编号
    '     strBakName-历史数据空间名称
    '     strBakOwner-历史数据空间所有者
    '     lngSys-系统号
    '     strDbLink-DBLink名称
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------
    '首先检查数据的有效性
    Dim rsTemp As New ADODB.Recordset
        
    'zlbakinfo(系统,版本号,更新日期)
    gstrSQL = "Insert Into zltools.zlbakspaces(系统, 编号, 名称, 所有者, db连接, 当前, 只读)　Values("
    gstrSQL = gstrSQL & "" & mlngSys & ","
    gstrSQL = gstrSQL & "" & lngBakCode & ","
    gstrSQL = gstrSQL & "'" & strBakName & "',"
    gstrSQL = gstrSQL & "'" & strBakOwner & "',"
    gstrSQL = gstrSQL & IIf(strDbLink = "", "NULL", "'" & strDbLink & "'") & ","
    gstrSQL = gstrSQL & "0,"
    gstrSQL = gstrSQL & "1) "
    err = 0: On Error Resume Next
    cnOracle.Execute gstrSQL
    If err <> 0 Then
        MsgBox "历史空间名或编号已经存在,植入失败,请检查!", vbInformation + vbDefaultButton1, mstrSysName
        Exit Function
    End If
    
    
    ExeFuncImport = True
End Function

Public Function CreateHistoryStru(ByVal strDb_Link As String, ByVal strBakUserName As String, ByVal strBakTableSpace As String, _
    ByVal strSourceUserName As String, ByVal strTbsNameIdx As String, ByVal strTbsNameLob As String) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:创建历史表空间的空间数据结构
    '参数:strDb_link-远程连接名
    '     strBakUserName-备份用户名
    '     strBakTablespace-备份表空间,strTbsNameIdx-索引表空间,strTbsNameLob-大对象表空间
    '     strSourceUserName-拷贝数据结构的源用户
    '返回:成功返回true,否则返回false
    '--------------------------------------------------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strSysIn As String, blnFeeView As Boolean
    
    '需要创建zlBakInfo表
    On Error GoTo errHand
    If CheckTable("zlBakInfo", mcnDBA, strBakUserName) = False Then
        gstrSQL = "Create Table " & strBakUserName & ".zlBakInfo(系统 number(5),版本号 varchar2(20),更新日期 date,最后转储日期 date,最后复制日期 date,中止语句 varchar2(500),提前执行 number(1),提前中止语句 VarChar2(500)) Tablespace " & strBakTableSpace
        mcnDBA.Execute gstrSQL
        gstrSQL = "Alter Table " & strBakUserName & ".zlBakInfo Add Constraint zlBakInfo_PK Primary Key (系统) USING INDEX PCTFREE 5"
        mcnDBA.Execute gstrSQL
    End If
    
    '--插入相关版本信息
    gstrSQL = "select 编号,版本号,sysdate from  zltools.zlsystems  where 编号=" & mlngSys
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , mcnOracle
    If rsTemp.EOF = False Then
        gstrSQL = "insert into " & strBakUserName & ".zlbakinfo(系统,版本号,更新日期)  values(" & mlngSys & ",'" & Nvl(rsTemp!版本号) & "',sysdate) "
        mcnDBA.Execute gstrSQL
    End If

    gstrSQL = "Select 表名 From zlbakTables a where a.系统=" & mlngSys
    
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , mcnOracle
    
    Call SetProgressVisible(True)
    pgbState.Max = rsTemp.RecordCount + 1
    pgbState.value = 0
    SetPromptText "复制结构"
     
    With rsTemp
        Do While Not .EOF
            If "" & !表名 = "病人费用记录" Then  '特殊处理
                blnFeeView = Not CheckTable(Nvl(!表名), mcnDBA, strBakUserName, 1)
            Else
                '检查表是否存在,存在将不创建
                If CheckTable(Nvl(!表名), mcnDBA, strBakUserName) = False Then
                    '创建表结构
                    If CreateTable(mcnOracle, strSourceUserName, strBakTableSpace, strBakUserName, !表名, strTbsNameLob, mcnDBA) = "" Then Call SetProgressVisible(False): Exit Function
            
                    '创建表结构相关的PK、UQ
                    If CreateConstraint(rsTemp!表名, strTbsNameIdx, strSourceUserName, strBakUserName) = False Then Call SetProgressVisible(False): Exit Function
                    '创建表结构索引IX
                    If CreateIndex(rsTemp!表名, strTbsNameIdx, strSourceUserName, strBakUserName) = False Then Call SetProgressVisible(False): Exit Function
                End If
            End If
             
            pgbState.value = pgbState.value + 1
            .MoveNext
        Loop
        If blnFeeView Then
            If mblnSysUpdateCall Then On Error Resume Next
            '历史表空间中的视图，创建后固定不变，当门诊费用记录或住院费用记录增加字段时，此视图不变，仅仅用于兼容旧的程序查询。
            strSQL = "CREATE OR REPLACE VIEW " & strBakUserName & ".病人费用记录 AS" & vbNewLine & _
                    "SELECT ID,记录性质,NO,实际票号,记录状态,序号,从属父号,价格父号,多病人单,记帐单ID,病人ID,主页ID,医嘱序号,门诊标志,记帐费用," & vbNewLine & _
                    "  姓名,性别,年龄,标识号,床号,病人病区ID,病人科室ID,费别,收费类别,收费细目ID,计算单位,付数,发药窗口,数次,加班标志,附加标志,婴儿费," & vbNewLine & _
                    "  收入项目ID,收据费目,标准单价,应收金额,实收金额,划价人,开单部门ID,开单人,发生时间,登记时间,执行部门ID,执行人,执行状态,执行时间,结论," & vbNewLine & _
                    "  操作员编号,操作员姓名,结帐ID,结帐金额,保险大类ID,保险项目否,保险编码,费用类型,统筹金额,是否上传,摘要,是否急诊" & vbNewLine & _
                    "FROM " & strBakUserName & ".住院费用记录" & vbNewLine & _
                    "UNION ALL" & vbNewLine & _
                    "SELECT ID,记录性质,NO,实际票号,记录状态,序号,从属父号,价格父号,-Null,记帐单ID,病人ID,-Null,医嘱序号,门诊标志,记帐费用," & vbNewLine & _
                    "  姓名,性别,年龄,标识号,付款方式,-Null,病人科室ID,费别,收费类别,收费细目ID,计算单位,付数,发药窗口,数次,加班标志,附加标志,婴儿费," & vbNewLine & _
                    "  收入项目ID,收据费目,标准单价,应收金额,实收金额,划价人,开单部门ID,开单人,发生时间,登记时间,执行部门ID,执行人,执行状态,执行时间,结论," & vbNewLine & _
                    "  操作员编号,操作员姓名,结帐ID,结帐金额,保险大类ID,保险项目否,保险编码,费用类型,统筹金额,是否上传,摘要,是否急诊" & vbNewLine & _
                    "FROM " & strBakUserName & ".门诊费用记录"
            mcnDBA.Execute strSQL
        End If
    End With
    Call SetProgressVisible(False)
    CreateHistoryStru = True
    Exit Function
errHand:
    If MsgBox("创建空间结构错误,详细的错误信息如下:" & vbCrLf & "(" & err.Number & ")" & err.Description & vbCrLf & strSQL, vbQuestion + vbRetryCancel + vbDefaultButton2) = vbRetry Then
        Resume
    End If
End Function

Private Function CheckTable(ByVal strTable As String, ByRef cnOracle As ADODB.Connection, ByVal strOwner As String, Optional ByVal bytType = 0) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------------
    '功能:检查表或视图是否存在
    '参数:strTable-表名
    '     cnoracle-数据库连接名
    '     strOwNer-所有者
    '     bytType=0:检查表，1-检查视图
    '返回:存在该对象则返回true,否则False
    '-----------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select 1 from all_objects where OBJECT_TYPE ='" & IIf(bytType = 0, "TABLE", "VIEW") & _
            "' and OWNER=[1] and object_name=[2]"
    Set rsTemp = gclsBase.OpenSQLRecord(cnOracle, gstrSQL, Me.Caption, UCase(strOwner), UCase(strTable))
    
    If rsTemp.EOF Then
        CheckTable = False
    Else
        CheckTable = True
    End If
End Function


Private Function CreateIndex(ByVal strTable As String, ByVal strBakTableSpace As String, ByVal strSourceUserName As String, ByVal strBakUserName As String) As Boolean
    '-------------------------------------------------------------------------
    '功能:创建相关表的约束
    '参数:strTable-表名
    '     strSourceUserName-拷贝数据结构的源用户
    '     strBakUserName-备份用户名
    '     strBakTableSpace-备份空间
    '返回:成功返回true,否则false
    '-------------------------------------------------------------------------
    Dim intType As VbMsgBoxResult
    
    Dim rsUserIndex As New ADODB.Recordset
    Dim rsColumn As New ADODB.Recordset
    Dim strTemp As String
    
    gstrSQL = "Select Table_Name,index_name, Column_Name " & _
             "   From Sys.All_Ind_Columns " & _
             "   Where Index_Owner = [1]  and table_name=[2] And Index_name not like '%_IX_待转出'" & _
             "   Order By index_name,Column_Position"
    Set rsColumn = gclsBase.OpenSQLRecord(mcnOracle, gstrSQL, Me.Caption, strSourceUserName, strTable)
     
    gstrSQL = "Select Index_name,table_name,tablespace_name,Pct_free,Temporary " & _
             "   From all_indexes a " & _
             "   where  table_owner = [1] and   table_name=[2] And index_type='NORMAL' And Index_name not like '%_IX_待转出'" & _
             "          And  Not Exists(Select 1 From All_Constraints b Where a.index_name=b.constraint_name  And Constraint_Type In ('P', 'U') And a.table_owner=b.Owner) " & _
             "   order by index_name"
    Set rsUserIndex = gclsBase.OpenSQLRecord(mcnOracle, gstrSQL, Me.Caption, strSourceUserName, strTable)
    
    On Error GoTo errHand
    With rsUserIndex
        Do While Not .EOF
            rsColumn.Filter = "index_name ='" & !Index_Name & "'"
            If rsColumn.EOF Then
                MsgBox "索引:" & !Index_Name & "不存在!", vbInformation + vbDefaultButton1
                If MsgBox("索引:" & !Index_Name & "不存在,是否继续!", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
            strTemp = ""
            Do While Not rsColumn.EOF
                strTemp = strTemp & "," & rsColumn!Column_Name
                rsColumn.MoveNext
            Loop
            
            If strTemp <> "" Then
                strTemp = Mid(strTemp, 2)
                If Nvl(!Temporary) = "Y" Then
                    gstrSQL = "CREATE INDEX " & strBakUserName & "." & !Index_Name & " ON  " & strBakUserName & "." & Nvl(!Table_Name) & "(" & strTemp & ") "
                Else
                    '由于是只读数据，为提高存储效率，固定pctfree为5
                    gstrSQL = "CREATE INDEX " & strBakUserName & "." & !Index_Name & " ON  " & strBakUserName & "." & Nvl(!Table_Name) & "(" & strTemp & ") PCTFREE 5" & IIf(strBakTableSpace = "", "", " TABLESPACE " & strBakTableSpace)
                End If
                mcnDBA.Execute gstrSQL
            End If
            DoEvents
            .MoveNext
        Loop
    End With
    
    CreateIndex = True
    Exit Function
errHand:
    intType = MsgBox("错误描述:" & err.Description & vbCrLf & gstrSQL & vbCrLf & "是否重试?", vbQuestion + vbAbortRetryIgnore + vbDefaultButton2, gstrSysName)
    If intType = vbAbort Then
        Exit Function
    ElseIf intType = vbIgnore Then
        Resume Next
    ElseIf intType = vbRetry Then
        Resume
    End If
End Function
Private Function CreateConstraint(ByVal strTable As String, ByVal strBakTableSpace As String, ByVal strSourceUserName As String, ByVal strBakUserName As String) As Boolean
    '-------------------------------------------------------------------------
    '功能:创建相关表的约束
    '参数:strTable-表名
    '     strBakTableSpace-历史空间
    '     strSourceUserName-拷贝数据结构的源用户
    '     strBakUserName-历史空间用户名
    '返回:成功返回true,否则false
    '-------------------------------------------------------------------------
    
    Dim rsObject As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    
    err = 0: On Error GoTo errHand:
    gstrSQL = "Select Constraint_Name, Constraint_Type, Table_Name, Delete_Rule, r_Constraint_Name, Deferrable " & _
          "   From Sys.All_Constraints " & _
          "   Where Generated = 'USER NAME' And Owner = [1] And Constraint_Type In ('P', 'U')  and table_name=[2]" & _
          "   Order By Decode(Constraint_Type, 'P', 0, 'U', 1, 2) "
    Set rsObject = gclsBase.OpenSQLRecord(mcnOracle, gstrSQL, Me.Caption, strSourceUserName, strTable)
    
    With rsObject
        Do While Not .EOF
            gstrSQL = " select table_name,column_name from sys.all_cons_columns " & _
                " where owner = [1]  and table_name = [2]  and constraint_name = [3]  order by position"
            Set rsTemp = gclsBase.OpenSQLRecord(mcnOracle, gstrSQL, Me.Caption, strSourceUserName, strTable, !Constraint_Name)
            
            strTemp = ""
            Do While Not rsTemp.EOF
                strTemp = strTemp & "," & rsTemp!Column_Name
                rsTemp.MoveNext
            Loop
            If strTemp <> "" Then
                strTemp = Mid(strTemp, 2)
                
                strSQL = "ALTER TABLE " & strBakUserName & "." & strTable & " ADD CONSTRAINT " & !Constraint_Name
                If !constraint_type = "U" Then
                    strSQL = strSQL & " UNIQUE (" & strTemp & ") "
                Else
                    strSQL = strSQL & " PRIMARY KEY  (" & strTemp & ") "
                End If
                If Nvl(!DEFERRABLE) = "DEFERRABLE" Then
                      strSQL = strSQL & " DEFERRABLE INITIALLY DEFERRED "
                End If
                
                '由于是只读数据，为提高存储效率，固定pctfree为5
                strSQL = strSQL & " USING INDEX PCTFREE 5 TABLESPACE " & strBakTableSpace
                
                '创建约束
                mcnDBA.Execute strSQL
            End If
            .MoveNext
        Loop
    End With
    CreateConstraint = True
    
    Exit Function
errHand:
    If MsgBox("创建约束失败，错误:" & vbCrLf & err.Description & vbCrLf & strSQL & vbCrLf & "是否跳过创建此约束？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        CreateConstraint = True
    End If
End Function


Private Function GetMaxHistory(Optional ByRef strMax名称 As String = "") As Integer
    '-------------------------------------------------------------------------------------------------------------------
    '功能:获取历史数据空间的最大编号
    '出参:str名称-产生的历史数据空间的最大号码
    '返回:最大编号
    '-------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select max(编号) as 编号,max(所有者) as 所有者,to_Char(sysdate-3*365,'yyyy') as 年   from zltools.zlBakSpaces where 系统=" & mlngSys
    OpenRecordset rsTemp, gstrSQL, "获取最大号码"
    If rsTemp.EOF Then
        strMax名称 = Format(DateAdd("yyyy", -3, Now), "YYYY")
        GetMaxHistory = 1
    Else
        If IsNull(rsTemp!所有者) Then
            strMax名称 = Nvl(rsTemp!年)
        Else
            strMax名称 = Replace(Replace(UCase(Nvl(rsTemp!所有者)), "ZLBAK", ""), "ZLHD", "")
            If IsNumeric(strMax名称) Then
                If mintFunType = F0创建 Then
                    strMax名称 = Val(strMax名称) + 1
                End If
            Else
                strMax名称 = Nvl(rsTemp!年)
            End If
        End If
        GetMaxHistory = Val(Nvl(rsTemp!编号)) + 1
    End If
End Function

Private Function SetHistoryInfor(ByRef cnOracle As ADODB.Connection, ByVal strBakUserName As String, ByVal strDbLink As String) As Boolean
    '-------------------------------------------------------------------------------------
    '功能:获取历史表空间的版本信息
    '-------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    If strDbLink <> "" Then strDbLink = "@" & strDbLink
    
    gstrSQL = "Select 1 from all_Tables" & strDbLink & " where table_name='ZLBAKINFO' and OWNER='" & UCase(strBakUserName) & "'"
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , cnOracle
    If rsTemp.EOF Then
        MsgBox "被植入的历史数据空间不是合法的历史数据空间(当前连接的用户模式下不存在ZLBAKINFO表),不能继续!", vbInformation + vbDefaultButton1, gstrSysName
         
        Exit Function
    End If
    
    '确定相关版本.
    gstrSQL = "Select 版本号,更新日期,最后转储日期,最后复制日期 From " & strBakUserName & ".ZLBAKINFO" & strDbLink & " where 系统=" & mlngSys
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , cnOracle
    If rsTemp.EOF Then
        MsgBox "被植入的历史数据空间不是" & mstrSysName & "的历史数据空间!", vbInformation + vbDefaultButton1, gstrSysName
        txtMoveName.Text = ""
        Exit Function
    End If
    
    lblBakVer.Caption = "备份:" & Nvl(rsTemp!版本号)
    lblBakVer.Tag = Nvl(rsTemp!版本号)
    
    If lblBakVer.Tag > lblDataVer.Tag Then
        If MsgBox("被植入的历史数据空间的版本大于了在线数据库的版本,你确定要继续吗?", vbQuestion + vbOKCancel + vbDefaultButton1) = vbCancel Then

            If txtMoveName.Enabled And txtMoveName.Visible Then txtMoveName.SetFocus
            lblBakVer.ForeColor = vbBlue
            lblDataVer.ForeColor = vbBlue
            shap.BorderColor = vbBlue
            Exit Function
        Else
            lblBakVer.ForeColor = vbRed
            lblDataVer.ForeColor = vbRed
            shap.BorderColor = vbRed
        End If
    ElseIf lblBakVer.Tag < lblDataVer.Tag Then
        lblBakVer.ForeColor = vbRed
        lblDataVer.ForeColor = vbRed
        shap.BorderColor = vbRed
    Else
        lblBakVer.ForeColor = &H80000008
        lblDataVer.ForeColor = &H80000008
        shap.BorderColor = &H80000008
    End If
    SetHistoryInfor = True
End Function

Private Function CheckCopyObject(ByRef cnBakOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lng编号 As Long, ByVal strBakOwnerName As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------
    '功能:检查复制非转存储数据空间
    '参数:cnOracle-历史数据连接
    '    lngSys-系统
    '    lng编号 -空间编号
    '    strBakOwnerName-空间用户名
    '返回:数据合法,返回true,否则False
    '-----------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strErr As String
    
   
    strErr = ""
    '检查版本
    gstrSQL = "Select 版本号,更新日期,最后转储日期,最后复制日期 From " & strBakOwnerName & ".ZLBAKINFO where 系统=" & mlngSys
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , cnBakOracle
    If rsTemp.EOF Then
        MsgBox "指定的历史数据空间不是" & mstrSysName & "的历史数据空间，不能复制非转储数据!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    If mstrVersion < Nvl(rsTemp!版本号) Then
        strErr = "历史数据空间的版本(" & Nvl(rsTemp!版本号) & "比在线库版本(" & mstrVersion & ") 还要大," & vbCrLf & " 是否继续复制？"
    ElseIf mstrVersion > Nvl(rsTemp!版本号) Then
        If MsgBox("历史数据空间的版本(" & Nvl(rsTemp!版本号) & "比在线库版本(" & mstrVersion & ") 要小," & vbCrLf & " 是否继续复制？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        strErr = "再次提醒：历史数据空间的版本(" & Nvl(rsTemp!版本号) & "比在线库版本(" & mstrVersion & ") 要小," & vbCrLf & " 是否继续复制？"
    Else
        strErr = ""
    End If
    If strErr <> "" Then
        If MsgBox(strErr, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    CheckCopyObject = True
End Function

Private Sub InitCreateTbs()
'功能：初始化创建页面的内容
    Dim rsTemp As New ADODB.Recordset, strDataFile As String, strTbsName As String
    
    On Error GoTo errh
    '根据当前系统的数据文件确定缺省的表空间文件路径
    gstrSQL = "Select File_Name as Name From Dba_Data_Files Where Tablespace_Name In ('ZL9BASEITEM', 'ZLTOOLSTBS') Order By File_Name"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption, , , mcnDBA)
    With rsTemp
        If .EOF Or .BOF Then
            strDataFile = "C:\"
        Else
            If InStr(1, StrReverse(!name), "\") > 0 Then
                strDataFile = Mid(!name, 1, Len(!name) - InStr(1, StrReverse(!name), "\") + 1)
            ElseIf InStr(1, StrReverse(!name), "/") > 0 Then
                strDataFile = Mid(!name, 1, Len(!name) - InStr(1, StrReverse(!name), "/") + 1)
            Else
                strDataFile = "C:\"
            End If
        End If
    End With
    txtDataFile.Text = strDataFile
    txtDataFile.Tag = strDataFile
    Call txtBakSpace_Change     '执行上一步，再执行下一次时，由于文本框内容未变，所以需要调用一次该事件
    
    '求最大编号
    If Trim(txt编号) = "" Then
        txt编号.Text = GetMaxHistory(strTbsName)
        If Trim(txtOwnerUsr.Text) = "" Then
            txtOwnerUsr.Text = strTbsName
        Else
            txtBakSpace.Text = txtHD.Text & txtOwnerUsr.Text
        End If
    End If
    
    tbHistory.Tab = 0
    If txtOwnerUsr.Enabled And txtOwnerUsr.Visible Then txtOwnerUsr.SetFocus
         
    cmdNext.Caption = "完成(&O)"

    Exit Sub
    
errh:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub cmdNext_Click()
    Dim strUserName As String, strPassword As String, strServer As String, strErrMsg As String
    Dim rsTemp As New ADODB.Recordset
    Dim bytErr As Byte, strError As String
    Dim strTbsName As String, strTbsNameIdx As String, strTbsNameLob As String, strDataFile As String
    Dim blnHaveUser As Boolean
    Dim lngFileAmount As Long, lngFileIdxAmount As Long, lngFileLobAmount As Long
    
    Dim strBakUserName As String, strBakUserPwd As String, strDbLink As String
    Dim lngSize As Long, blnAutoExpent As Boolean, blnAutoLocate As Boolean, intExpentSize As Integer
    Dim strRemarks As String
   
    On Error GoTo errHand
    mblnSucced = False
    SetPromptText ""
    
    If fraSetup(0).Visible Then
        '------------------------------------------------------------
        '第一步(拆卸,删除是第二步)：确定远程历史空间的DBA用户是否存在
        '------------------------------------------------------------
        SetControlEnable False
        
        strUserName = txtDba用户.Text
        strPassword = txtDba口令.Text
        strServer = Trim(txtDbaServer.Text)
        
        If CheckUser(strUserName, strPassword, strServer, strErrMsg) = False Then
            MsgBox strErrMsg, vbExclamation, gstrSysName
            Call SetControlEnable(True)
            Exit Sub
        End If
        txtDba用户.Text = strUserName
        txtDba口令.Text = strPassword
        txtDbaServer.Text = strServer
        
        If optServer(1).value = True Then strDbLink = Trim(txtDBLink.Text)
        
        
        If mintFunType = F6转移 Then
            If strServer = gstrServer Then
                MsgBox "源数据库和目标数据库不能是同一个，请重新指定服务器", vbInformation, gstrSysName
                Call SetControlEnable(True)
                If txtDbaServer.Visible And txtDbaServer.Enabled Then txtDbaServer.SetFocus
                Exit Sub
            End If
            
        ElseIf mintFunType = F3复制 Then
            If Trim(txtDbaServer.Text) = "" Then
                MsgBox "历史空间连接的服务名为空，复制操作要求必须指定服务名，请重新输入。", vbInformation, gstrSysName
                Call SetControlEnable(True)
                If txtDbaServer.Visible And txtDbaServer.Enabled Then txtDbaServer.SetFocus
                Exit Sub
            End If
        End If
        
        If mintFunType <> F4切换 And mintFunType <> F2再植 Then
                        
            '创建时，如果当前用户是DBA，则没有显示数据库连接配置界面，直接用当前连接用户来操作。
            '拆卸时，可能要删除远程历史库，所以，用输入的连接信息来打开新连接（即使执行了测试操作）
            If mintFunType <> F0创建 Or (mintFunType = F0创建 And mcnDBA.State = adStateClosed) Then
                Set mcnDBA = gobjRegister.GetConnection(strServer, strUserName, strPassword, False, MSODBC, strError, False)
            End If
            
            If mcnDBA.State = adStateClosed Then
                MsgBox strError, vbInformation, gstrSysName
                Call SetControlEnable(True)
                If txtDba口令.Visible And txtDba口令.Enabled Then txtDba口令.SetFocus
                Exit Sub
            Else
                Call SetSQLTrace(strServer, strUserName, mcnDBA)
                
                If Not (mintFunType = F3复制) Then
                    If CheckIsDBA(mcnDBA) = False Then
                        MsgBox "不是DBA用户,不能继续！", vbExclamation, gstrSysName
                        Call SetControlEnable(True)
                        If txtDba用户.Visible And txtDba用户.Enabled Then txtDba用户.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        If mintFunType = F0创建 Then
            Call InitCreateTbs
                    
            Call SetControlEnable(True)
            
            fraSetup(0).Visible = False
            fraSetup(1).Visible = True
            cmdPrevious.Enabled = True
            
        ElseIf mintFunType = F1拆卸 Then '拆卸(删除模式，第二步)
            '验证身份并输入操作说明
            If Not CheckAuditStatus("0201", "拆卸", strRemarks) Then
                Call SetControlEnable(True)
                Exit Sub
            End If
            If ExeFuncUnInstall(lblSpace.Tag, UCase(lblSpaceOwner.Tag), mlng空间编号, False) Then
                MsgBox "历史数据空间" & lblSpace.Tag & "拆卸(删除)成功。", vbInformation + vbDefaultButton1, gstrSysName
                '插入重要操作日志
                Call SaveAuditLog(3, "拆卸", "拆卸(删除)历史数据空间“" & lblSpace.Tag & "”", strRemarks)
                mblnSucced = True
            End If
            Call SetControlEnable(True)
            Unload Me
            
        ElseIf mintFunType = F2再植 Then '再植
            
            If strDbLink <> "" Then
                '创建远程连接
                If CreateDbLink(mcnOracle, strDbLink, strUserName, strPassword, strServer, mstrOwnerName) = False Then
                    Call SetControlEnable(True)
                    Exit Sub
                End If
            End If
            
            If SetHistoryInfor(mcnOracle, strUserName, strDbLink) = False Then
                Call SetControlEnable(True)
                Exit Sub
            End If
            fraSetup(0).Visible = False
            fraImport.Visible = True
            
            '求最大编号
            If Trim(txtMoveCode) = "" Then
                txtMoveCode.Text = GetMaxHistory(strTbsName)
                If Trim(txtMoveName.Text) = "" Then
                    txtMoveName.Text = strTbsName
                End If
            End If
            
            Call SetControlEnable(True)
            cmdPrevious.Enabled = True
            cmdNext.Caption = "再植(&Z)"
            
            txtMoveName.Text = strUserName  '默认为历史数据空间用户
            txtMoveUser.Text = strUserName
                        
            If txtMoveName.Enabled And txtMoveName.Visible Then txtMoveName.SetFocus
            
        ElseIf mintFunType = F3复制 Then
            '检查版本及数据对象是否正确.
            If MsgBox("复制非转储数据可能要花费较长的时间，并且历史空间中的同名表的数据将会被覆盖，你确定要继续吗？", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Call SetControlEnable(True)
                If txtDba用户.Visible And txtDba用户.Enabled Then txtDba用户.SetFocus
                Exit Sub
            End If
                        
            If CheckCopyObject(mcnDBA, mlngSys, mlng空间编号, strUserName) = False Then
                Call SetControlEnable(True)
                If txtDba用户.Visible And txtDba用户.Enabled Then txtDba用户.SetFocus
                Exit Sub
            End If
            
            strTbsName = GetBakTableSpace(mcnDBA, strUserName)
            If ExeFuncCopy(strUserName, strPassword, strServer, strTbsName) = False Then
                Call SetControlEnable(True)
                If txtDba用户.Visible And txtDba用户.Enabled Then txtDba用户.SetFocus
                Exit Sub
            End If
                        
            Call SetControlEnable(True)
            If Image1.ToolTipText = "保留临时文件，名称复制到剪切板" Then
                MsgBox "已将复制脚本产生到文件" & Clipboard.GetText, vbInformation + vbDefaultButton1, gstrSysName
            Else
                MsgBox "复制成功!", vbInformation + vbDefaultButton1, gstrSysName
            End If
            Unload Me
            
        ElseIf mintFunType = F4切换 Then
            
            
            If strDbLink <> "" Then
                If CreateDbLink(mcnOracle, strDbLink, strUserName, strPassword, strServer, mstrOwnerName) = False Then
                    '创建建接失败
                    Call SetControlEnable(True)
                    Exit Sub
                End If
            Else
                gstrSQL = "Update ZLTOOLS.zlbakspaces set DB连接='" & strDbLink & "' where 编号 = " & mlng空间编号
                mcnOracle.Execute gstrSQL
            End If
            
            If ExeFuncChange(strUserName, mstrOwnerName, mlngSys, bytErr, strErrMsg, strDbLink) = False Then
                '1-连接失效,2-系统不存在,3-在线版本大于历史版本,4-在线版本小于历史版本
                Call SetControlEnable(True)
                Select Case bytErr
                Case 1, 2, 4
                    MsgBox strErrMsg, vbInformation + vbDefaultButton1, gstrSysName
                Case 3
                    MsgBox strErrMsg, vbInformation + vbDefaultButton1, gstrSysName
                    cmd升级.Visible = True
                    lblSetupIni.Visible = True
                    lblIniModi.Visible = True
                End Select
            Else
                Call SetControlEnable(True)
                mblnSucced = True
                MsgBox "切换成功!", vbInformation + vbDefaultButton1, gstrSysName
                Unload Me
            End If
        ElseIf mintFunType = F5合并 Then
            
            mblnSucced = ExeFuncMerge
            
            Call SetControlEnable(True)
            Unload Me
            
        ElseIf mintFunType = F6转移 Then
            Call SetControlEnable(True)
            
            If CheckTransCondition = False Then
                Unload Me
            ElseIf LoadSpaceData = False Then
                Unload Me
            Else
                lblNoteTrans.Caption = "1.不能选择[当前]历史数据空间，如果仅有一个历史数据空间，请先在源服务器上创建新的历史空间且置为[当前]；" & vbCrLf & _
                            "2.目标服务器不能存在与源服务器同名的历史空间名称及数据文件；" & vbCrLf & _
                            "3.源服务器历史空间中的对象要求是自包含的，即表的索引与表存储在同一表空间，如果存在存储在其他表空间的索引，将会自动重建。"
                            
                cmdNext.Caption = "转移(&Z)"
                cmdPrevious.Enabled = True
                fraSetup(0).Visible = False
                fraTrans.Visible = True
            End If
        End If
        
    ElseIf fraSetup(1).Visible Then '创建
        '------------------------------------------------------------
        '第二步：确定相应的历史数据空间的名称及数据文件
        '------------------------------------------------------------
        SetPromptText "正在检查数据的有效性..."
        
        blnHaveUser = False
        If CheckCreateBakInput = False Then SetPromptText "": Exit Sub
                        
        If MsgBox("你确定现在开始创建历史数据表空间吗？", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
            Exit Sub
        End If
                        
        SetPromptText ""
        strUserName = txtDba用户.Text
        strPassword = txtDba口令.Text
        strServer = txtDbaServer.Text
        If optServer(1).value = True Then strDbLink = Trim(txtDBLink.Text)
        mstrDBLink = strDbLink  '创建出错时，删除过程会用到mstrDBLink
        
        '创建Dblink
        If strDbLink <> "" Then
            If CreateDbLink(mcnOracle, strDbLink, strUserName, strPassword, strServer, mstrOwnerName) = False Then SetPromptText "": Exit Sub
        End If
        
        '创建历史库用户，可能用户存在，下面一段变量初始化放后面，因为历史库存在时，需要获取已经存在的历史库表空间等信息
        If CheckBakUser(blnHaveUser, strDbLink) = False Then SetPromptText "": Exit Sub
        
        strBakUserName = txtHD.Text & txtOwnerUsr.Text
        strBakUserPwd = Trim(txtOwnerPwd)
        strTbsName = Trim(txtBakSpace.Text)
        strTbsNameIdx = Trim(txtBakSpaceIdx.Text)
        strTbsNameLob = Trim(txtBakSpaceLob.Text)
        
        strDataFile = Trim(txtDataFile.Text)
        lngSize = Val(txtSpaceSize.Text)
        blnAutoExpent = chkSpaceExtd.value = 1
        blnAutoLocate = cboSpaceExtentType.ListIndex = 0
        intExpentSize = Val(txtSpaceExtentSize.Text)
        
        lngFileAmount = Val(txtFileAmount(0).Text)
        lngFileIdxAmount = Val(txtFileAmount(1).Text)
        lngFileLobAmount = Val(txtFileAmount(2).Text)
        

        
        Call SetControlEnable(False)
  
        If ExeFuncCreate(strUserName, strPassword, strServer, strBakUserName, strBakUserPwd, strDbLink, strTbsName, strDataFile, lngSize, _
                    blnAutoExpent, blnAutoLocate, intExpentSize, blnHaveUser, strTbsNameIdx, strTbsNameLob, lngFileAmount, lngFileIdxAmount, lngFileLobAmount) = False Then
            MsgBox "安装失败，系统将自动清除已经安装的内容…", vbInformation, gstrSysName
            DoEvents
            Call ExeFuncUnInstall(strBakUserName, UCase(strBakUserName), Val(txt编号), True)
            Call SetControlEnable(True)
            SetPromptText ""
        
            Exit Sub
        Else
            
            '置为当前的数据空间
            If chkCreate当前.value = 1 Then
                SetPromptText "正在创建视图..."
                Call SetProgressVisible(True)
                If CreateAppView(mstrOwnerName, strBakUserName, mlngSys, IIf(strDbLink = "", "", "@" & strDbLink), pgbState) = False Then
                    Call SetProgressVisible(False)
                    MsgBox "置为当前历史数据空间失败,请先检查后，再用[再植]功能植入!", vbInformation, gstrSysName
                    Call SetControlEnable(True)
                    Call UpdateZlBakSpace(mcnOracle, Val(txt编号), mlngSys, True)
                    DoEvents
                    Unload Me
                    Exit Sub
                End If
                Call SetProgressVisible(False)
                
                mcnOracle.BeginTrans
                gstrSQL = "Update zltools.zlbakspaces set 当前=0 where 系统=" & mlngSys & ""
                mcnOracle.Execute gstrSQL
                gstrSQL = "Update zltools.zlbakspaces set 当前=1 where 系统=" & mlngSys & " and 编号=" & Val(txt编号.Text)
                mcnOracle.Execute gstrSQL
                mcnOracle.CommitTrans
                '编译无效对象:
                Call ReCompileObjects(mcnOracle)
            Else
                '对象检查
                If mblnMustInstall = True Then
                    Call ReCompileObjects(mcnOracle)
                    mblnMustInstall = False
                End If
            End If
        End If
        
        MsgBox "已成功创建历史数据空间!", vbInformation, gstrSysName
        
        mblnSucced = True
        Call SetControlEnable(True)
        Unload Me
        
    ElseIf fraDelete.Visible Then   '拆卸(第一步)
        
        lblNote(0).Caption = "指定DBA用户，用于删除表空间及数据文件。"
        fraDelete.Visible = False
        
        cmdNext.Caption = "拆卸(&O)"
        cmdPrevious.Enabled = True
        
        '剥离模式
        If optDele(1).value Then
            '验证身份并输入操作说明
            If Not CheckAuditStatus("0201", "拆卸", strRemarks) Then Exit Sub
            gstrSQL = "delete ZLTOOLS.zlbakSpaces where  nvl(当前,0)<>1 and 系统=" & mlngSys & " and 编号=" & mlng空间编号
            mcnOracle.Execute gstrSQL
            
            Me.Hide
            MsgBox "历史数据空间拆卸(剥离)成功。" & vbCrLf & "你可以通过“再植”功能将剥离的历史空间重新植入!", vbInformation + vbDefaultButton1, gstrSysName
            '插入重要操作日志
            Call SaveAuditLog(3, "拆卸", "拆卸(剥离)历史数据空间“" & lblSpace.Tag & "”", strRemarks)
            mblnSucced = True
            Unload Me
        Else
            fraSetup(0).Visible = True
            
            If txtDba用户.Text = "" And txtDba用户.Enabled Then txtDba用户.Text = "SYS"
            If txtDba口令.Enabled And txtDba口令.Visible Then txtDba口令.SetFocus
            
            If mstrDBLink <> "" Then
                optServer(1).value = True
                optServer(0).Enabled = False
                
                txtDBLink.Text = mstrDBLink '禁止修改
                txtDBLink.Enabled = False
            Else
                optServer(0).value = True
                optServer(1).Enabled = False
            End If
        End If
        
    ElseIf fraImport.Visible Then   '再植
        '0-检查输入的数据是否合法
        SetPromptText "正在检查数据的有效性..."
        
        If optServer(1).value = True Then strDbLink = Trim(txtDBLink.Text)
        
        If CheckMoveInPutValid(strDbLink) = False Then SetPromptText "": Exit Sub
        SetPromptText ""

        '1.确定是否存在结构升级
        Call SetControlEnable(False)
        
        
        '历史空间不存在结构等升迁或本次只植入
        If ExeFuncImport(mcnOracle, Val(txtMoveCode.Text), Trim(txtMoveName), Trim(txtMoveUser), mlngSys, _
            strDbLink, strUserName, strPassword, Trim(txtDbaServer.Text)) = False Then
            '删除相关信息:
            Call UpdateZlBakSpace(mcnOracle, Val(txtMoveCode.Text), mlngSys, True)
            Call SetControlEnable(True)
            Exit Sub
        End If
        mblnSucced = True
        If chk当前.value = 1 Then
            '需要创建相应的视图
            '-----------------------------------------------------
            SetPromptText ("正在创建视图")
            Call SetProgressVisible(True)
            If CreateAppView(mstrOwnerName, Trim(txtMoveUser), mlngSys, IIf(strDbLink = "", "", "@" & strDbLink), pgbState) = False Then
                Call SetProgressVisible(False)
                MsgBox "植入当前系统时失败,请在管理界面重置!", vbInformation + vbDefaultButton1, gstrSysName
                Call SetControlEnable(True)
                Unload Me
                Exit Sub
            End If
            Call SetProgressVisible(False)
            '编译无效对象:
            Call ReCompileObjects(mcnOracle)
            MsgBox "植入成功!", vbInformation + vbDefaultButton1, gstrSysName
            
            Call UpdateZlBakSpace(mcnOracle, Val(txtMoveCode.Text), mlngSys, False, strDbLink <> "")
        End If
        
        Call SetControlEnable(True)
        Unload Me
    ElseIf fraTrans.Visible Then    '传输
        If lvwHistory.SelectedItem Is Nothing Then
             MsgBox "请选择要传输的表空间!", vbInformation + vbDefaultButton1, gstrSysName
             Exit Sub
        Else
            If txtBakPWD.Text = "" Then
                MsgBox "请输入目标库新建的历史空间用户密码。", vbInformation + vbDefaultButton1, gstrSysName
                txtBakPWD.SetFocus
                Exit Sub
            End If
            
            If lvwHistory.SelectedItem.SubItems(C2当前) = "√" Then
                MsgBox "请选择非当前历史空间。因为传输后会删除该空间，所以不能选择当前历史空间。", vbInformation + vbDefaultButton1, gstrSysName
                lvwHistory.SetFocus
                Exit Sub
            End If
            
            strTbsName = lvwHistory.SelectedItem.SubItems(C1名称)
            strBakUserName = lvwHistory.SelectedItem.SubItems(C4所有者)
            
            '1.检查用户名
            gstrSQL = "Select Decode(Trunc(Created), Trunc(Sysdate), 1, 0) Todaycreate From Dba_Users Where Username = '" & strBakUserName & "'"
            Set rsTemp = New ADODB.Recordset
            Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                '如果是当天刚创建的（例如上次传输失败），则提供删除选择
                If rsTemp!Todaycreate = 1 Then
                    If MsgBox("选择的历史空间用户" & strBakUserName & "在目标数据库中已存在，你确定要删除后重新创建吗？(该用户下的所有对象将会被一起删除)", vbOKCancel + vbQuestion + vbDefaultButton1, gstrSysName) = vbOK Then
                        
                        SetPromptText "正在删除用户" & strBakUserName & "及对象…"
                        DoEvents
                        gstrSQL = "drop user " & strBakUserName & " cascade"
                        mcnOracle.Execute gstrSQL
                    Else
                        lvwHistory.SelectedItem.ForeColor = &HC0C0C0
                        Exit Sub
                    End If
                Else
                    MsgBox "选择的历史空间用户" & strBakUserName & "在目标数据库中已存在，传输之前请先删除同名用户!", vbInformation + vbDefaultButton1, gstrSysName
                    lvwHistory.SelectedItem.ForeColor = &HC0C0C0
                    Exit Sub
                End If
            End If
            
            '2.检查表空间名
            '如果是当天刚创建的（例如上次传输失败），则提供删除选择
            gstrSQL = "Select Decode(Trunc(Creation_Time), Trunc(Sysdate), 1, 0) Todaycreate" & vbNewLine & _
                    "From Dba_Data_Files A, V$datafile B" & vbNewLine & _
                    "Where a.File_Id = b.File# And Tablespace_Name = '" & strTbsName & "' Order by Creation_Time"
            Set rsTemp = New ADODB.Recordset
            Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                If rsTemp!Todaycreate = 1 Then
                    If MsgBox("选择的历史空间" & strTbsName & "在目标数据库中已存在，你确定要删除后重新创建吗？(该表空间下的所有对象将会被一起删除)", vbOKCancel + vbQuestion + vbDefaultButton1, gstrSysName) = vbOK Then
                        SetPromptText "正在删除表空间" & strTbsName & "及对象…"
                        DoEvents
                        gstrSQL = "alter tablespace " & strTbsName & " offline"
                        mcnOracle.Execute gstrSQL
                        gstrSQL = "drop tablespace " & strTbsName & " including contents and datafiles cascade constraints"
                        mcnOracle.Execute gstrSQL
                    Else
                        lvwHistory.SelectedItem.ForeColor = &HC0C0C0
                        Exit Sub
                    End If
                Else
                    MsgBox "选择的历史空间" & strTbsName & "已在目标数据库中存在，传输之前请先删除同名表空间!", vbInformation + vbDefaultButton1, gstrSysName
                    lvwHistory.SelectedItem.ForeColor = &HC0C0C0
                    Exit Sub
                End If
                
            End If
            
            '3.检查表空间文件名
            'ASM上的文件可能是zlbak2.263.832000313这样的，忽略不计
            gstrSQL = "Select 1 From dba_data_files Where file_Name like '%/" & strTbsName & ".DBF' or file_Name like '%\" & strTbsName & ".DBF'"
            Set rsTemp = New ADODB.Recordset
            Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                MsgBox "选择的历史空间文件" & strTbsName & ".DBF已在目标数据库中存在，传输之前请先删除同名数据文件!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Sub
            End If
        End If
        
        Call SetControlEnable(False)
        
        mblnSucced = ExeFuncTrans
    
        Call SetControlEnable(True)
        Unload Me
    End If
    
    Exit Sub
errHand:
    MsgBox err.Description, vbInformation, gstrSysName
    Call SetControlEnable(True)

    Unload Me
End Sub

Private Sub cmdPrevious_Click()
    
    Select Case mintFunType
    Case F1拆卸
        fraDelete.Visible = True
        fraSetup(0).Visible = False
        fraSetup(1).Visible = False
        Call optDele_Click(0)
    Case F0创建
        If fraSetup(1).Visible Then
            fraSetup(1).Visible = False
            fraSetup(0).Visible = True
            cmdPrevious.Enabled = False
            cmdNext.Caption = "下一步(&N)"
            If txtDba口令.Enabled Then txtDba口令.SetFocus
        End If
        
    Case F2再植, F5合并, F6转移
        If mintFunType = F2再植 Then
            fraImport.Visible = False
        ElseIf mintFunType = F5合并 Then
            fraMerge.Visible = False
        ElseIf mintFunType = F6转移 Then
            fraTrans.Visible = False
        End If
                
        fraSetup(0).Visible = True
        cmdPrevious.Enabled = False
        cmdNext.Caption = "下一步(&N)"
        cmdNext.Enabled = True
        If txtDba用户.Enabled And txtDba用户.Visible Then txtDba用户.SetFocus
    End Select
End Sub

Private Sub cmd连接_Click()
    Dim strUserName As String, strPassword As String, strServer As String, strError As String
           
    strUserName = txtDba用户.Text
    strPassword = txtDba口令.Text
    strServer = txtDbaServer.Text
    
    If CheckUser(strUserName, strPassword, strServer, strError) = False Then
        MsgBox strError, vbExclamation, gstrSysName
        Exit Sub
    End If
    txtDba用户.Text = strUserName
    txtDba口令.Text = strPassword
    txtDbaServer.Text = strServer
    
    
    '下面这种加了ADDRESS_LIST的写法，在ODBC下，只支持SID，不支持SERVICE_NAME;OLEDB则两种都支持
    'strServer = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strIP & ")(PORT=" & strPort & ")))(CONNECT_DATA=(SID=" & strSID & ")))"
    Set mcnDBA = gobjRegister.GetConnection(strServer, strUserName, strPassword, False, MSODBC, strError, False)
   
    If mcnDBA.State = adStateClosed Then
        MsgBox "打开数据库连接出错。" & strError, vbExclamation, gstrSysName
        If txtDba口令.Visible And txtDba口令.Enabled Then txtDba口令.SetFocus
        Exit Sub
    ElseIf mintFunType <> F3复制 And mintFunType <> F2再植 Then
        
        If CheckIsDBA(mcnDBA) = False Then
            MsgBox "不是DBA用户,不能继续！", vbExclamation, gstrSysName
            If txtDba用户.Visible And txtDba用户.Enabled Then txtDba用户.SetFocus
            Exit Sub
        End If
    End If
    MsgBox "测试成功!", vbInformation + vbDefaultButton1, gstrSysName
End Sub

Private Function CheckIsDBA(ByRef connThis As ADODB.Connection) As Boolean
'功能：判断当前用户是否为DBA角色
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errh
    gstrSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTemp = gclsBase.OpenSQLRecord(connThis, gstrSQL, "判断当前连接用户是否具有DBA角色")
    CheckIsDBA = rsTemp.RecordCount > 0
    
    Exit Function
errh:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Function ExeFuncChange(strBakUserName As String, _
    strOwner As String, lngSys As Long, bytErr As Byte, strErr As String, ByVal strDbLink As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------
    '功能:当前空间切换
    '参数:
    '     strBakUserName-历史数据空间用户
    '     strOwner-创建的所有者
    '出参:bytErr:1-连接失效,2-系统不存在,3-在线版本大于历史版本,4-在线版本小于历史版本
    '     strErr-错误描述
    
    '返回:设置成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    
    '确定当前的连接是否存功
    On Error Resume Next
    If strDbLink = "" Then
        gstrSQL = "Select 版本号 from " & strBakUserName & ".zlbakinfo  where 系统=" & lngSys
    Else
        gstrSQL = "Select 版本号 from " & strBakUserName & ".zlbakinfo@" & strDbLink & "  where 系统=" & lngSys
    End If
    
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , mcnOracle
    bytErr = 0
    If err <> 0 Then
        strErr = "查找相关版本信息错误" & "详细的错误信息:" & vbCrLf & "(" & err.Number & ")" & err.Description
        bytErr = 1
        Exit Function
    End If
    err.Clear: err = 0
    
    '检查相关的版本是否正确
    If rsTemp.EOF Then
        strErr = "历史数据空间的系统(" & mstrSysName & ") 不存在,请检查!"
        bytErr = 2
        Exit Function
    End If
    If Nvl(rsTemp!版本号) < mstrVersion Then
        '当前版本号小于在线版本号,需升级
        strErr = "历史数据空间的系统版本(" & Nvl(rsTemp!版本号) & ") 小于了在线版本(" & mstrVersion & ")," & vbCrLf & " 请升级历史数据!"
        bytErr = 3
        Exit Function
    ElseIf Nvl(rsTemp!版本号) > mstrVersion Then
        '大于在线版本号
        strErr = "历史数据空间的系统版本(" & Nvl(rsTemp!版本号) & ") 大于了在线版本(" & mstrVersion & ")," & vbCrLf & " 请升级在线版本后再切换,请检查!"
        bytErr = 4
        Exit Function
    End If
        
    '可以正常切换了
    SetPromptText ("正在创建视图")
    Call SetProgressVisible(True)
    If CreateAppView(mstrOwnerName, strBakUserName, mlngSys, IIf(strDbLink = "", "", "@" & strDbLink), pgbState) = False Then
        Call SetProgressVisible(False)
        '失败
        Exit Function
    End If
    Call SetProgressVisible(False)
    
    '更新标志
    If UpdateZlBakSpace(mcnOracle, mlng空间编号, lngSys, False, strDbLink <> "") = False Then
        MsgBox "更新标志失败,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    '编译无效对象:
    Call ReCompileObjects(mcnOracle)
    ExeFuncChange = True
End Function

Private Sub cmd升级_Click()
    Dim objfrmUpSys As frmAppUpgradeNew
    '处理升级
    '历史数据空间升级
    Dim strUserName As String, strPassword As String, strServer As String, strErrMsg As String
    
    strUserName = txtDba用户.Text
    strPassword = txtDba口令.Text
    strServer = txtDbaServer.Text
    
    
    If CheckUser(strUserName, strPassword, strServer, strErrMsg) = False Then
        MsgBox strErrMsg, vbExclamation, gstrSysName
        Exit Sub
    End If
    txtDba用户.Text = strUserName
    txtDba口令.Text = strPassword
    txtDbaServer.Text = strServer
    strPassword = strPassword
    
    Set objfrmUpSys = New frmAppUpgradeNew '用来清除模块变量
    If objfrmUpSys.HistoryUp(Me, stbThis.Panels(2), mlngSys, lblSpace.Tag, lblSetupIni.Tag, strUserName, strPassword, strServer, mstrVersion, mstrDBLink) Then
        '不需要刷新界面
        SetPromptText "升级成功！"
    Else
        SetPromptText "升级失败！"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If mblnSysUpdateCall Then
        If mblnMustInstall And mintFunType = F0创建 And mblnSucced = False Then
            If MsgBox("当前系统必需安装历史数据空间后，才能正常" & vbCrLf & "使用该系统,如果现在不创建，稍后请前往【数据转移管理】模块" & vbCrLf & "进行创建!是否现在创建？", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Cancel = 1
                Exit Sub
            End If
        End If
    Else
        If cmdNext.Enabled = False Then
            Cancel = 1
            Exit Sub
        End If
    
        If mblnMustInstall And mintFunType = F0创建 And mblnSucced = False Then
            MsgBox "当前系统必需安装历史数据空间后，才能正常" & vbCrLf & "使用该系统,因此不能取消操作!", vbInformation + vbDefaultButton1, gstrSysName
            Cancel = 1
            Exit Sub
        End If
    End If
    
    Set mrsMergeSpace = Nothing
    
    '不关闭连接对象mcnDBA,因为该对象可能是传入的，多个功能连续使用时会导致传入的连接被关闭
End Sub

Private Sub Image1_DblClick()
    If mintFunType = F3复制 Then
        Image1.ToolTipText = "保留临时文件，名称复制到剪切板"
        MsgBox "仅产生复制脚本文件，不实际执行脚本", vbInformation
    End If
End Sub

Private Sub lblIniModi_Click()
    Dim strFile As String
    
    With cdgPub
        .DialogTitle = "选择应用安装配置文件"
        .Filter = "应用安装配置文件(zlSetup.ini)|zlSetup.ini"
        .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        strFile = IIf(lblSetupIni.Tag = "", "", lblSetupIni.Tag)
        If gobjFile.FileExists(strFile) Then
            .InitDir = gobjFile.GetParentFolderName(strFile)
            .Filename = gobjFile.GetFileName(strFile)
        Else
            .InitDir = "": .Filename = ""
        End If
        On Error Resume Next
        .CancelError = True
        .ShowOpen
        err.Clear: On Error GoTo errh
        If .Filename <> "" Then
            If .Filename <> lblSetupIni.Tag Then
                '配置文件改变，检查配置文件
                If CheckInitFile(mlngSys, .Filename) Then
                    lblSetupIni.Caption = "安装配置文件：" & .Filename
                    lblSetupIni.Tag = .Filename
                    lblSetupIni.ToolTipText = .Filename
                    Call SetCtrlPosOnLine(False, 0, lblSetupIni, 60, lblIniModi)
                    lblSetupIni.Refresh
                    If lblSetupIni.Width >= 5100 Then
                        lblSetupIni.Width = 5100
                    End If
                End If
            End If
        End If
        On Error GoTo 0
    End With
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Private Sub optDele_Click(Index As Integer)
    If optDele(0).value Then
        optDele(0).FontBold = True
        optDele(1).FontBold = False
        cmdNext.Caption = "下一步(&N)"
    Else
        optDele(1).FontBold = True
        optDele(0).FontBold = False
        cmdNext.Caption = "剥离(&O)"
    End If
End Sub
Private Sub optDele_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub optServer_Click(Index As Integer)
    txtDBLink.Enabled = Index = 1
    txtDbaServer.Enabled = txtDBLink.Enabled
    If Index = 1 Then
        txtDbaServer.Text = ""
    Else
        txtDbaServer.Text = gstrServer
    End If
    
    lblServerName(1).Visible = Index = 1
    txtDBLink.Visible = Index = 1
End Sub

Private Sub tbHistory_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        If txtDataFile.Enabled And txtDataFile.Visible Then txtDataFile.SetFocus
    Else
        If txtOwnerUsr.Enabled And txtOwnerUsr.Visible Then txtOwnerUsr.SetFocus
    End If
End Sub

 

Private Sub txtBakSpace_Change()
    Dim strFileBase As String
    
    strFileBase = txtDataFile.Tag & txtBakSpace.Text
    
    txtDataFile.Text = strFileBase & ".dbf"
    txtBakSpaceIdx.Text = txtBakSpace.Text & "_IDX"
    txtBakSpaceLob.Text = txtBakSpace.Text & "_LOB"
End Sub

Private Sub txtBakSpace_GotFocus()
    Call SelAll(txtBakSpace)
End Sub

Private Sub txtBakSpace_KeyPress(KeyAscii As Integer)
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
                If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
                    If InStr(1, "_", Chr(KeyAscii)) = 0 Then
                        If KeyAscii <> 13 And KeyAscii <> 8 Then
                            KeyAscii = 0
                        End If
                    End If
                End If
            End If
        End If
End Sub
 
Private Sub txtDbaServer_GotFocus()
  Call SelAll(txtDbaServer)
End Sub
 

Private Sub txtDbaServer_KeyPress(KeyAscii As Integer)
    If InStr(1, ",.-+~!#$%^&*()|\/>'<" & """") > 0 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtDba口令_GotFocus()
    Call SelAll(txtDba口令)
End Sub

Private Sub txtDba用户_GotFocus()
    Call SelAll(txtDba用户)
End Sub


Private Sub txtDBLink_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
            If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
                If InStr(1, "_", Chr(KeyAscii)) = 0 Then
                    If KeyAscii <> 13 And KeyAscii <> 8 Then
                        KeyAscii = 0
                    End If
                End If
            End If
        End If
    End If
End Sub


Private Sub txtFileAmount_GotFocus(Index As Integer)
    Call SelAll(txtFileAmount(Index))
End Sub

Private Sub txtFileAmount_KeyPress(Index As Integer, KeyAscii As Integer)
    Call LimitInputNumber(KeyAscii)
End Sub

Private Sub txtFileAmount_Validate(Index As Integer, Cancel As Boolean)
    If Not IsNumeric(txtFileAmount(Index).Text) Then Cancel = True
End Sub


Private Sub txtMoveCode_GotFocus()
    Call SelAll(txtMoveCode)
End Sub

Private Sub txtMoveName_GotFocus()
    Call SelAll(txtMoveName)
End Sub

Private Sub txtMoveName_KeyPress(KeyAscii As Integer)
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
                If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
                    If InStr(1, "_", Chr(KeyAscii)) = 0 Then
                        If KeyAscii <> 13 And KeyAscii <> 8 Then
                            KeyAscii = 0
                        End If
                    End If
                End If
            End If
        End If
End Sub

Private Sub txtMoveUser_GotFocus()
    Call SelAll(txtMoveUser)
End Sub

Private Sub txtMoveUser_KeyPress(KeyAscii As Integer)
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
                If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
                    If InStr(1, "_", Chr(KeyAscii)) = 0 Then
                        If KeyAscii <> 13 And KeyAscii <> 8 Then
                            KeyAscii = 0
                        End If
                    End If
                End If
            End If
        End If
End Sub

Private Sub txtOwnerLab_GotFocus()
    Call SelAll(txtOwnerLab)

End Sub

Private Sub txtOwnerPwd_GotFocus()
    Call SelAll(txtOwnerPwd)

End Sub

Private Sub txtOwnerUsr_Change()
    txtBakSpace.Text = txtHD.Text & txtOwnerUsr.Text
    
End Sub

Private Sub txtOwnerUsr_GotFocus()
    Call SelAll(txtOwnerUsr)
End Sub

Private Sub txtOwnerUsr_KeyPress(KeyAscii As Integer)
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
                If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
                    If InStr(1, "_", Chr(KeyAscii)) = 0 Then
                        If KeyAscii <> 13 And KeyAscii <> 8 Then
                            KeyAscii = 0
                        End If
                    End If
                End If
            End If
        End If
End Sub

Private Sub txtOwnerUsr_LostFocus()
        txtOwnerUsr.Text = UCase(txtOwnerUsr.Text)
End Sub

Private Sub txtSpaceExtentSize_KeyPress(KeyAscii As Integer)
    Call LimitInputNumber(KeyAscii)
End Sub

Private Sub SetProgressVisible(ByVal blnVisible As Boolean)
    If blnVisible = True Then
        If stbThis.Panels.Count = 3 Then
            '增加一个窗格
            stbThis.Panels.Add 3
            stbThis.Panels(3).AutoSize = sbrSpring
            stbThis.Panels(2).AutoSize = sbrNoAutoSize
            stbThis.Panels(2).MinWidth = 2440
        End If
        pgbState.Left = stbThis.Panels(3).Left + 30
        pgbState.Width = stbThis.Panels(4).Left - pgbState.Left - 150
        pgbState.Top = stbThis.Top + stbThis.Height / 3
        pgbState.Visible = True
    Else
        If stbThis.Panels.Count = 4 Then
            stbThis.Panels(2).AutoSize = sbrSpring
            stbThis.Panels.Remove 3
        End If
        pgbState.Visible = False
    End If
End Sub
Private Sub SetPromptText(ByVal strText As String)
    stbThis.Panels(2).Text = strText
    stbThis.Panels(2).ToolTipText = strText
End Sub

Private Function DropDBLinkOfUser(ByRef cnOracle As ADODB.Connection, ByVal strUserName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '功能:删除所有的指向指定用户的远程连接
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = " select Owner,DB_LINK from all_db_links  where USERNAME='" & strUserName & "'"
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , cnOracle
    With rsTemp
        Do While Not .EOF
            '如果指向的是同一个用户的Db_Link，则设置正确
             On Error Resume Next
             gstrSQL = " Drop Database Link " & !Owner & "." & Nvl(!DB_LINK)
             cnOracle.Execute gstrSQL
             err.Clear: err = 0
            .MoveNext
        Loop
    End With
End Function

Private Sub DropTablespace(ByVal strTableSpace As String)
'功能：删除指定的表空间
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 1 From Dba_Tablespaces Where Tablespace_Name = '" & strTableSpace & "'"
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , mcnDBA
    If rsTemp.RecordCount > 0 Then
        gstrSQL = "alter tablespace " & strTableSpace & " offline"
        mcnDBA.Execute gstrSQL
        gstrSQL = "drop tablespace " & strTableSpace & " including contents and datafiles cascade constraints"
        mcnDBA.Execute gstrSQL
    End If
End Sub

Private Function ExeFuncUnInstall(ByVal strTableSpace As String, ByVal strUserName As String, ByVal lng编号 As Long, Optional blnErrResume As Boolean = False) As Boolean
'功能：删除已经安装的历史数据空间
    Dim rsTemp As New ADODB.Recordset
    Dim strErrInfo As String, strTBS As String
        
    strErrInfo = ""
    If blnErrResume = False Then
        On Error GoTo errHand   '拆卸时调用
    Else
        On Error Resume Next    '安装失败时调用
    End If
                  
         
    '1.删除指当前用户的远程连接对象
    Call DropDBLinkOfUser(mcnDBA, strUserName)
    
    '2.删除本系统所有者及对象(表、索引等)
    SetPromptText "正在删除历史数据空间用户及相关对象…"
    DoEvents
    gstrSQL = "drop user " & strUserName & " cascade"
    mcnDBA.Execute gstrSQL
    
    
    '3.删除本系统数据表空间
    SetPromptText "正在删除历史数据表空间和数据文件…"
    If CheckTableSpaceIsUse("表空间", strTableSpace, strUserName, mcnDBA) = False Then
        '没有其他用户使用，可以删除
        DropTablespace (strTableSpace)
        DropTablespace (strTableSpace & "_LOB")
        DropTablespace (strTableSpace & "_IDX")
    End If
        
    '4.删除历史数据空间目录
    gstrSQL = "delete zltools.zlbakspaces where 系统= " & mlngSys & " and 编号=" & lng编号
    mcnOracle.Execute gstrSQL

    If mstrDBLink <> "" Then
        '确定是否有不同系统指向同一连接的.则不能删,否则删除
        gstrSQL = "Select 1 From ZLTOOLS.zlbakSpaces where upper(DB连接)=upper('" & mstrDBLink & "') and 系统<>" & mlngSys
        OpenRecordset rsTemp, gstrSQL, Me.Caption, , , mcnOracle
        If rsTemp.EOF Then
            On Error Resume Next
            gstrSQL = "Drop DataBase Link  " & mstrDBLink
            mcnOracle.Execute gstrSQL
            If err <> 0 Then
                  Call MsgBox("删除远程连接名出错,详细请况如下:" & vbCrLf & "(" & err.Number & ") " & vbCrLf & err.Description, vbInformation)
            End If
        End If
    End If

    ExeFuncUnInstall = True
    
    Exit Function
errHand:
    MsgBox err.Description & vbCrLf & "SQL语句：" & gstrSQL, vbInformation, gstrSysName
End Function

Private Function CreateTbs(ByVal TbsName As String, ByVal TbsFile As String, ByVal TbsSize As Long, ByVal AutoExtend As Boolean, _
     ByVal AutoAllocate As Boolean, ByVal ExtentSize As Integer, ByVal lngFileAmount As Long) As Byte
    '----------------------------------------------
    '功能：系统用户,根据参数创建表空间,固定为本地管理类型(8i以前不支持,那时只能创建字典管理类型)
    '参数：
    '   TbsName:表空间名称
    '   TbsFile:表空间文件
    '   TbsSize:表空间大小(M为单位)
    '   Extend:是否自动管理区,否则统一范围尺寸
    '   ExtentSize:统一区尺寸,临时表空间必须指定尺寸(Oracle缺省为1M)
    '   Temp:是否为临时表空间
    '   lngFileAmount:数据文件的数量
    '返回：1-创建成功；2-表空间已经存在；3-创建失败,4-磁盘空间不够
    '----------------------------------------------
    Dim strFileHead As String, strFileTail As String, i As Long
    Dim strFile As String
    
    strFile = "'" & TbsFile & "' Size " & TbsSize & "M " & IIf(AutoExtend, "AUTOEXTEND ON", "")
    If lngFileAmount > 1 Then
        strFileHead = Mid(TbsFile, 1, InStrRev(TbsFile, ".") - 1)
        strFileTail = Mid(TbsFile, InStrRev(TbsFile, "."))
        
        For i = 1 To lngFileAmount - 1
            strFile = strFile & ",'" & strFileHead & "_" & i & strFileTail & "' SIZE " & TbsSize & "M " & IIf(AutoExtend, "AUTOEXTEND ON", "")
        Next
    End If
        
    gstrSQL = "CREATE TABLESPACE " & TbsName & " DATAFILE " & strFile & _
            " EXTENT MANAGEMENT LOCAL " & _
            IIf(AutoAllocate, " AUTOALLOCATE", " UNIFORM SIZE " & IIf(ExtentSize = 0, "1", ExtentSize) & "M") & " Nologging"
            
    err = 0
    On Error Resume Next
    mcnDBA.Execute gstrSQL
    
    
    If err = 0 Then
        CreateTbs = 1
    ElseIf mcnDBA.Errors.Count > 0 Then
        If mcnDBA.Errors.Item(0).NativeError = 1144 Then
            MsgBox "创建的表空间（" & TbsName & "）的磁盘空间不足,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
            CreateTbs = 4
        ElseIf mcnDBA.Errors.Item(0).NativeError = 1119 Then
            Call MsgBox("数据文件(" & TbsFile & ")设置错误，不能继续!" & vbCrLf & "错误信息:" & mcnDBA.Errors(0).Description & vbCrLf & gstrSQL, vbInformation Or vbDefaultButton2, gstrSysName)
            CreateTbs = 3
        Else
            If MsgBox("出现下述错误，是否跳过继续？" & vbCrLf & vbTab & mcnDBA.Errors(0).Description & vbCrLf & gstrSQL, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                CreateTbs = 2
            Else
                CreateTbs = 1
            End If
        End If
    Else
        MsgBox "创建表空间（" & TbsName & "）失败:" & vbCrLf & gstrSQL & vbCrLf & err.Description, vbInformation + vbDefaultButton1, gstrSysName
        CreateTbs = 3
    End If
End Function

Private Function CheckUser(ByRef strUserName As String, ByRef strPassword As String, ByRef strServer As String, ByRef strErrMsg As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:检查用户名、密码和服务器输入是否正确
    '入参:strUsername-用户名,strPassWord-密码,strServer-服务器
    '出参:与入参相同
    '返回:用户合法，返回true,否则返回False
    '-------------------------------------------------------------------------------------------------------------------------
    '------检验用户是否oracle合法用户----------------
    '有效字符串效验
    If Len(Trim(strUserName)) = 0 Then
        strErrMsg = "请输入用户名。"
        If txtDba用户.Enabled And txtDba用户.Visible Then txtDba用户.SetFocus
        Exit Function
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            strErrMsg = "用户名错误。"
            If txtDba用户.Enabled And txtDba用户.Visible Then txtDba用户.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            strErrMsg = "口令错误。"
            If txtDba口令.Enabled And txtDba口令.Visible Then txtDba口令.SetFocus
            
            Exit Function
        End If
    End If
    
    If Trim(strServer) <> "" Then
        If Mid(strServer, Len(strServer) - 1, 1) = "/" Or Mid(strServer, Len(strServer) - 1, 1) = "@" Or Mid(strServer, 1, 1) = "/" Or Mid(strServer, 1, 1) = "@" Then
            strErrMsg = "服务器连接串错误。"
            If txtDbaServer.Enabled And txtDbaServer.Visible Then txtDbaServer.SetFocus
            Exit Function
        End If
    End If
    
    '分离字符串
    Dim intPos As Integer
    
    intPos = InStr(1, strUserName, "@", vbTextCompare)
    If intPos > 0 Then
        strServer = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strUserName, "/", vbTextCompare)
    If intPos > 0 Then
        strPassword = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strPassword, "@", vbTextCompare)
    If intPos > 0 Then
        strServer = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    
    If Len(Trim(strPassword)) = 0 Then
        strErrMsg = "未输入密码!"
        If txtDba口令.Enabled And txtDba口令.Visible Then txtDba口令.SetFocus
        Exit Function
    End If
    
    strUserName = UCase(strUserName)
    
    CheckUser = True
     
End Function

Private Function CheckMoveInPutValid(ByVal strDbLink As String) As Boolean
    '----------------------------------------------------------------------------------
    '功能:检查植入历史数据空间输入的相关项是否合法
    '参数:
    '返回;成功返回true,否则返回false
    '----------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset

    If Val(Trim(txtMoveCode.Text)) <= 0 Then
        MsgBox "请输入正确的空间编号。", vbExclamation, gstrSysName
        If txtMoveCode.Enabled And txtMoveCode.Visible Then txtMoveCode.SetFocus
        Exit Function
    End If
    If Val(Trim(txtMoveCode.Text)) > 999 Then
        MsgBox "空间编号不能大于999。", vbExclamation, gstrSysName
        If txtMoveCode.Enabled And txtMoveCode.Visible Then txtMoveCode.SetFocus
         Exit Function
    End If

    If Trim(txtMoveName.Text) = "" Then
        MsgBox "空间名称无效,请检查。", vbExclamation, gstrSysName
        If txtMoveName.Enabled And txtMoveName.Visible Then txtMoveName.SetFocus
         Exit Function
    End If

    If ActualLen(Trim(txtMoveName.Text)) > 30 Then
        MsgBox "空间名称的长度不能大于30个字符。", vbExclamation, gstrSysName
        If txtMoveName.Enabled And txtMoveName.Visible Then txtMoveName.SetFocus
         Exit Function
    End If
    

    gstrSQL = "Select 1 From zlBakSpaces where 系统=" & mlngSys & " and (编号=" & Val(txtMoveCode.Text) & " or upper(名称)=upper('" & txtMoveName & "'))"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    If Not rsTemp.EOF Then
        MsgBox "输入的编号或名称已经存在,请重新设置编号或名称!", vbInformation + vbDefaultButton1, mstrSysName
        If txtMoveCode.Visible And txtMoveCode.Enabled Then txtMoveCode.SetFocus
        rsTemp.Close
        Exit Function
    End If
  
    Dim bytType As Byte, strErrMsg As String
    
    '检查是否存在升级
    If lblDataVer.Tag > lblBakVer.Tag Then
        If chk当前.value = 1 Then
            MsgBox "该历史数据空间的版本与在线不符，请对历史空间升迁后才能再植为当前!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    ElseIf lblDataVer.Tag < lblBakVer.Tag Then
        If chk当前.value = 1 Then
            MsgBox "该历史数据空间的版本大于了在线版本，请对在线数据库升迁后才能再植为当前!", vbInformation + vbDefaultButton1, gstrSysName
            If txtMoveUser.Enabled And txtMoveUser.Visible Then txtMoveUser.SetFocus
            Exit Function
        End If
        bytType = 0
    Else
        bytType = 2
    End If
    
    '检查相关的数据结构是否合法
    If CheckHistoryObject(mcnOracle, strDbLink, mlngSys, txtMoveUser.Text, bytType, strErrMsg) = False Then
        MsgBox "进行对象检查时，发现有如下错误:" & vbCrLf & strErrMsg
        If chk当前.value = 1 Then
            If txtMoveUser.Enabled And txtMoveUser.Visible Then txtMoveUser.SetFocus
            Exit Function
        End If
    End If
    
    CheckMoveInPutValid = True
End Function

Private Function CheckCreateBakInput() As Boolean
    '----------------------------------------------------------------------------------
    '功能:检查创建历史数据空间输入的相关项是否合法
    '参数:
    '返回;成功返回true,否则返回false
    '----------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    If Val(Trim(txt编号.Text)) <= 0 Then
        MsgBox "请输入正确的空间编号。", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txt编号.Enabled Then txt编号.SetFocus
        Exit Function
    End If
    If Val(Trim(txt编号.Text)) > 999 Then
        MsgBox "空间编号不能大于999。", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txt编号.Enabled Then txt编号.SetFocus
         Exit Function
    End If
    
    '表空间检查
    If Trim(txtBakSpace.Text) = "" Then
        MsgBox "请输入正确的空间名称。", vbExclamation, gstrSysName
        tbHistory.Tab = 1
        If txtBakSpace.Enabled And txtBakSpace.Visible = True Then txtBakSpace.SetFocus
         Exit Function
    End If
    If ActualLen(Trim(txtBakSpace.Text)) > 30 Then
        MsgBox "空间名的长度不能大于30个字符。", vbExclamation, gstrSysName
        tbHistory.Tab = 1
        If txtBakSpace.Enabled And txtBakSpace.Visible Then txtBakSpace.SetFocus
         Exit Function
    End If
    
    If Val(txtSpaceSize.Text) > 100000 Then
        MsgBox "表空间超过100G了。", vbExclamation, gstrSysName
        tbHistory.Tab = 1
        If txtBakSpace.Enabled And txtBakSpace.Visible Then txtBakSpace.SetFocus
        Exit Function
    End If
    If Val(txtSpaceSize.Text) <= 0 Then
        MsgBox "表空间必需大于零。", vbExclamation, gstrSysName
        tbHistory.Tab = 1
        If txtBakSpace.Enabled And txtBakSpace.Visible Then txtBakSpace.SetFocus
        Exit Function
    End If
    
    '数据文件检查
    If InStr(txtDataFile.Text, ".") = 0 Then
        MsgBox "数据文件缺少扩展名。", vbExclamation, gstrSysName
        tbHistory.Tab = 1
        If txtDataFile.Enabled And txtDataFile.Visible Then txtDataFile.SetFocus
        Exit Function
    End If
    
    If Val(txtFileAmount(0).Text) <= 0 Or Val(txtFileAmount(1).Text) <= 0 Or Val(txtFileAmount(2).Text) <= 0 Then
        MsgBox "数据文件数量须大于零。", vbExclamation, gstrSysName
        tbHistory.Tab = 1
        If txtFileAmount(0).Enabled And txtFileAmount(0).Visible Then txtFileAmount(0).SetFocus
        Exit Function
    End If
    
    
    If Trim(txtOwnerUsr.Text) = "" Then
        MsgBox "请输入正确的用户名。", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txtOwnerUsr.Enabled Then txtOwnerUsr.SetFocus
         Exit Function
    End If
    If ActualLen(Trim(txtOwnerUsr.Text)) > 30 Then
        MsgBox "用户名的长度不能大于30个字符。", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txtOwnerUsr.Enabled Then txtOwnerUsr.SetFocus
         Exit Function
    End If
    
    If Trim(txtOwnerPwd.Text) = "" Then
        MsgBox "请输入口令。", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txtOwnerPwd.Enabled Then txtOwnerPwd.SetFocus
         Exit Function
    End If
    If Trim(txtOwnerLab.Text) = "" Then
        MsgBox "请输入验证口令。", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txtOwnerLab.Enabled Then txtOwnerLab.SetFocus
         Exit Function
    End If
    
    If Trim(txtOwnerLab.Text) <> Trim(txtOwnerPwd.Text) Then
        MsgBox "输入的口令与验证口令不致，请重输!", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txtOwnerLab.Enabled Then txtOwnerLab.SetFocus
         Exit Function
    End If
    
    If optServer(1).value = True Then
        If Trim(txtDbaServer.Text) = "" Then
            MsgBox "请重输入服务器名!", vbExclamation, gstrSysName
            tbHistory.Tab = 0
            If txtDbaServer.Visible And txtDbaServer.Enabled Then txtDbaServer.SetFocus
            Exit Function
        End If
        
        If Trim(txtDBLink.Text) = "" Then
            MsgBox "请重输入DBLink名称!", vbExclamation, gstrSysName
            tbHistory.Tab = 0
            If txtDBLink.Enabled Then txtDBLink.SetFocus
            Exit Function
        End If
    End If
        
    
    On Error GoTo errh
     
    gstrSQL = "Select 1 From zlBakSpaces where 系统=" & mlngSys & " and (编号=" & Val(txt编号.Text) & " or upper(名称)=upper('" & txtHD.Text & txtOwnerUsr.Text & "'))"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    If Not rsTemp.EOF Then
        MsgBox "输入的编号或名称重复,请重输!", vbInformation + vbDefaultButton1, mstrSysName
        tbHistory.Tab = 0
        If txt编号.Visible And txt编号.Enabled Then txt编号.SetFocus
        rsTemp.Close
        Exit Function
    End If
    
    CheckCreateBakInput = True
    Exit Function
errh:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Function CheckBakUser(ByRef blnHaveUser As Boolean, ByVal strDbLink As String) As Boolean
'功能：检查并创建历史空间用户
    Dim rsTemp As New ADODB.Recordset, cnTemp As ADODB.Connection
    Dim strUser As String, strPass As String, strServer As String
    Dim strError As String, strTbsName As String, strSQL As String
    
    SetPromptText "正在检查用户的有效性..."
    strUser = Trim(txtHD.Text & txtOwnerUsr.Text)
    strPass = Trim(txtOwnerPwd.Text)
    strServer = Trim(txtDbaServer.Text)
    
    On Error GoTo errh
    gstrSQL = "select 1 from dba_users where username='" & strUser & "'"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption, , , mcnDBA)
    
    blnHaveUser = rsTemp.RecordCount > 0
    If rsTemp.RecordCount > 0 Then
        If MsgBox("用户名为“" & strUser & "”的历史数据空间已经存在,是否将本系统的历史数据空间添加到该用户下?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            tbHistory.Tab = 0
            If txtOwnerUsr.Enabled And txtOwnerUsr.Visible Then txtOwnerUsr.SetFocus
            Exit Function
        End If
    
        '检查密码是否合法
        Set cnTemp = gobjRegister.GetConnection(strServer, strUser, strPass, False, MSODBC, strError, False)
        If cnTemp.State = adStateClosed Then
            MsgBox strError, vbInformation, gstrSysName
            tbHistory.Tab = 0
            If txtOwnerUsr.Enabled And txtOwnerUsr.Visible Then txtOwnerUsr.SetFocus
            Exit Function
        End If
        '由于历史库名称和所有者不同，因此需要重新获取
        '历史库表空间获取
        strSQL = "Select a.Tablespace_Name From User_Tables A Where a.Table_Name = 'ZLBAKINFO'"
        Set rsTemp = gclsBase.OpenSQLRecord(cnTemp, strSQL, "获取历史库表空间")
        If Not rsTemp.EOF Then
            txtBakSpace.Text = rsTemp!Tablespace_Name
        End If
        strTbsName = UCase(Trim(txtBakSpace.Text))
        '获取索引表空间与LOB表空间
        strSQL = "Select a.Tablespace_Name" & vbNewLine & _
                "From User_Tablespaces A" & vbNewLine & _
                "Where a.Tablespace_Name In ('" & strTbsName & "', '" & strTbsName & "_IDX', '" & strTbsName & "_LOB')"
        Set rsTemp = gclsBase.OpenSQLRecord(cnTemp, strSQL, "获取历史库表空间", strTbsName)
        rsTemp.Filter = "Tablespace_Name='" & strTbsName & "_IDX'"
        If Not rsTemp.EOF Then
            txtBakSpaceIdx.Text = rsTemp!Tablespace_Name
        Else
            txtBakSpaceIdx.Text = strTbsName
        End If
        rsTemp.Filter = "Tablespace_Name='" & strTbsName & "_LOB'"
        If Not rsTemp.EOF Then
            txtBakSpaceLob.Text = rsTemp!Tablespace_Name
        Else
            txtBakSpaceLob.Text = strTbsName
        End If
        cnTemp.Close
        Set cnTemp = Nothing
        
        If CheckDiffUserStru(mcnOracle, mcnDBA, strUser, mlngSys, strDbLink) = False Then
            tbHistory.Tab = 0
            If txtOwnerUsr.Enabled And txtOwnerUsr.Visible Then txtOwnerUsr.SetFocus
            Exit Function
        End If
    Else
        On Error Resume Next
        gstrSQL = "create user " & strUser & " identified by " & strPass
        mcnDBA.Execute gstrSQL
        
        If err.Number <> 0 Then
            MsgBox "历史数据空间用户名(" & strUser & ")或口令不符合数据库要求，请重新定义。" & vbCrLf & err.Description, vbExclamation, gstrSysName
            tbHistory.Tab = 0
            If txtOwnerUsr.Enabled And txtOwnerUsr.Visible Then txtOwnerUsr.SetFocus
            Exit Function
        End If
    End If
    
    CheckBakUser = True
    Exit Function
errh:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Function CheckDiffUserStru(ByVal cnOracle As ADODB.Connection, ByVal cnOracleBak As ADODB.Connection, ByVal strBakUserName As String, _
    ByVal lngSys As Long, ByVal strDbLink As String) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能:检查不同系统指定到同一备份用户下的数据结构
    '入参:cnOracle-在线数据库
    '     cnOracleBak-历史数据库连接
    '     strBakUserName-历史数据空间的所有者
    '     lngSys-当前系统的系统号
    '--------------------------------------------------------------------------------------------------------------
    Dim rsTemp  As New ADODB.Recordset
    Dim strTemp As String, strSysIn As String
    
    On Error GoTo errHandle
    
    gstrSQL = "Select 1 From ALL_tables" & IIf(strDbLink = "", "", "@" & strDbLink) & " where Owner='" & strBakUserName & "' And Table_name = '" & UCase("zlBakInfo") & "'"
    OpenRecordset rsTemp, gstrSQL, "检查历史库信息", , , cnOracle
    
    If rsTemp.RecordCount > 0 Then
        gstrSQL = "Select 系统 From " & strBakUserName & ".zlBakInfo"
        OpenRecordset rsTemp, gstrSQL, "获取历史库系统号", , , cnOracleBak
        strSysIn = ""
        With rsTemp
            Do While Not .EOF
                If lngSys = Val(Nvl(!系统)) Then
                    MsgBox "指定的历史空间中已经存在该系统了,不必新创建历史数据空间,请用[再植]功能!", vbInformation + vbDefaultButton1, gstrSysName
                    Exit Function
                End If
                strSysIn = strSysIn & "," & Val(Nvl(!系统))
                .MoveNext
            Loop
        End With
        If strSysIn = "" Then
            CheckDiffUserStru = True
            Exit Function
        End If
    Else
        CheckDiffUserStru = True
        Exit Function
    End If
    
    gstrSQL = "select a.表名 from zlbaktables a,zlbaktables b where a.表名=b.表名 and a.系统 IN (" & Mid(strSysIn, 2) & ") and b.系统=" & lngSys
    OpenRecordset rsTemp, gstrSQL, "获取指定系统是否存在共享表", , , cnOracle
    If rsTemp.EOF Then
        CheckDiffUserStru = True
        Exit Function
    End If
    
    strTemp = ""
    With rsTemp
        Do While Not .EOF
            strTemp = strTemp & "    " & Nvl(!表名) & vbCrLf
            .MoveNext
        Loop
    End With
    gstrSQL = "Select 1 from zlsystems where 编号=" & lngSys & " and nvl(共享号,0) in (" & Mid(strSysIn, 2) & ")"
    OpenRecordset rsTemp, gstrSQL, "获取指定系统是否存在共享", , , cnOracle
    
    '存在共享,不用在判断
    If rsTemp.EOF = False Then
        CheckDiffUserStru = True
        Exit Function
    End If
    MsgBox "你选择的历史数据空间有如下表存在:" & vbCrLf & strTemp & vbCrLf & " 不能指定该历史数据空间!", vbInformation + vbDefaultButton1, gstrSysName
    
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & "最近执行的SQL：" & gstrSQL, vbExclamation, Me.Caption
End Function

Private Function CheckTableSpaceIsUse(ByVal strType As String, ByVal strName As String, ByVal strOwner As String, cnOracle As Connection) As Boolean
    '功能：检查表空间或数据文件是否由其它用户使用
    '参数：strType    表空间 数据文件
    '      strName          表空间或数据文件的名字
    '      strOwner         以区别其它用户的所有者名
    Dim rsTemp As New ADODB.Recordset
    
    If strType = "表空间" Then
        gstrSQL = "select owner from all_tables where tablespace_name='" & UCase(strName) & "' and owner<>'" & UCase(strOwner) & "' AND ROWNUM<2"
    Else
        gstrSQL = "select O.owner  from V$TABLESPACE T,V$DATAFILE F,all_tables O " & _
                  "Where T.TS# = F.TS# And T.name = O.TABLESPACE_NAME " & _
                  "    and F.name='" & UCase(strName) & "' and O.owner<>'" & UCase(strOwner) & "' AND ROWNUM<2 "
    End If
    
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , cnOracle
    
    If rsTemp.RecordCount = 0 Then
        '没有其他用户使用，可以删除
        CheckTableSpaceIsUse = False
    Else
        '有用户使用
        CheckTableSpaceIsUse = True
    End If
End Function

'检查是否存在历史数据空间
Private Function IsHavingHistoryTable(ByVal lngSys As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否存在历史数据空间
    '返回:存在历史空间数据表.返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    IsHavingHistoryTable = False
    gstrSQL = "Select 表名 From zlBakTables where 系统=" & lngSys
    OpenRecordset rsTemp, gstrSQL, "检查表是否存在!", , , mcnOracle
    If rsTemp.EOF Then
        Exit Function
    End If
    IsHavingHistoryTable = True
End Function


Private Sub txtSpaceExtentSize_Validate(Cancel As Boolean)
    If Not IsNumeric(txtSpaceExtentSize.Text) Then Cancel = True
End Sub

Private Sub txtSpaceSize_KeyPress(KeyAscii As Integer)
    Call LimitInputNumber(KeyAscii)
End Sub

Private Sub LimitInputNumber(ByRef KeyAscii As Integer)
'功能：限制只能输入数字
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> 13 And KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtSpaceSize_Validate(Cancel As Boolean)
    If Not IsNumeric(txtSpaceSize.Text) Then Cancel = True
End Sub

Private Sub txt编号_GotFocus()
    Call SelAll(txt编号)

End Sub

Private Sub txt编号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt编号_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Function LogTime() As String
    LogTime = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] "
End Function
 

Private Function CheckHistoryObject(ByVal cnOracle As ADODB.Connection, ByVal strDbLink As String, ByVal lngSys As Long, _
    ByVal strBakOwnerName As String, Optional bytCheckSys As Byte, Optional ByRef strErrMsg As String) As Boolean
    '--------------------------------------------------------------------------------------------------------------------------
    '功能:检查历史数据空间的相对象是否完整
    '参数:strDbLink-历史数据库DBLink连接
    '     cnOracle-在线数据库连接
    '     lngSys-系统号
    '     strBakOwnerName-历史数据空间
    '     bytCheckSys-0-仅仅检查是否在zlbakInfor表中存在系统(不检查表）,1-仅仅检查表对象,>1表示全检查:主要是检查对象和表
    '出参:strErrMsg-返回相关的错误信息
    '返回:如果检查合法,返回true,否则返回False
    '--------------------------------------------------------------------------------------------------------------------------
    Dim rsBakObject As New ADODB.Recordset
    Dim rsObject As New ADODB.Recordset
    Dim strErrInfor As String
    
    On Error GoTo errHand
    If strDbLink <> "" Then strDbLink = "@" & strDbLink

    gstrSQL = "select table_name as 表名  from all_tables" & strDbLink & " where  owner = upper('" & strBakOwnerName & "') "
    OpenRecordset rsBakObject, gstrSQL, Me.Caption, , , cnOracle
    
    '检查zlBakInfo表是否存在
    rsBakObject.Filter = "表名='" & UCase("zlBakInfo") & "'"
    If rsBakObject.EOF Then
        strErrInfor = strErrInfor & vbCrLf & Space(4) & "不存在:zlBakInfo表"
        strErrMsg = strErrInfor
        Exit Function
    End If
    If (bytCheckSys = 0 Or bytCheckSys > 1) Then
        
        gstrSQL = "Select 1 From " & strBakOwnerName & ".zlBakInfo" & strDbLink & " where 系统=" & lngSys
        OpenRecordset rsObject, gstrSQL, Me.Caption, , , cnOracle
        If rsObject.EOF Then
            strErrInfor = strErrInfor & vbCrLf & Space(4) & "在历史数据空间中该系统不存在,请检查!"
            Set rsObject = Nothing
            strErrMsg = strErrInfor
            Exit Function
        End If
        rsObject.Close
    End If
     
    
    If bytCheckSys >= 1 Then
        strErrInfor = ""
        gstrSQL = "Select 表名 from zlbakTables where 系统=" & lngSys
        OpenRecordset rsObject, gstrSQL, Me.Caption, , , cnOracle  '当前库的转出表定义
        With rsObject
            Do While Not .EOF
                rsBakObject.Filter = "表名='" & Nvl(!表名) & "'"
                If rsBakObject.EOF Then '记录历史库中不存在的转出表
                    strErrInfor = strErrInfor & vbCrLf & Space(4) & Nvl(!表名)
                End If
                .MoveNext
            Loop
        End With
    End If
    rsBakObject.Close
    Set rsBakObject = Nothing

    If strErrInfor <> "" Then
        strErrInfor = "以下表在历史库中不存在:" & Mid(strErrInfor, 2) & vbCrLf & "可能是历史库的版本太低。" & _
            IIf(chk当前.value = 1, vbCrLf & "可以在植入之后再切换为当前历史库（同时升级）。", "")
    Else
        CheckHistoryObject = True
    End If
    If bytCheckSys = 0 Then Exit Function
            
    If strErrInfor <> "" Then
        If Mid(strErrInfor, 1, 1) = vbCrLf Then
            strErrMsg = Mid(strErrInfor, 2)
        Else
            strErrMsg = strErrInfor
        End If
    End If
    
    Exit Function
errHand:
    strErrInfor = "(" & err.Number & ")" & err.Description
    strErrMsg = strErrInfor
End Function

Private Sub SetControlEnable(ByVal blnEnable As Boolean)
    '-----------------------------------------------------------------------------
    '功能:设置相关控件的Eanble属性
    '-----------------------------------------------------------------------------
    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Frame" Then
            ctl.Enabled = blnEnable
        ElseIf ctl Is cmdPrevious Or ctl Is cmdNext Or ctl Is cmdHelp Or ctl Is cmdCancel Then
            If blnEnable = False Then
                ctl.Tag = IIf(ctl.Enabled = True, 1, 0)
                ctl.Enabled = blnEnable
            Else
                ctl.Enabled = Val(ctl.Tag) = 1
                ctl.Tag = ""
                
            End If
        End If
    Next
End Sub

Private Sub ReCompileObjects(ByRef cnThis As ADODB.Connection)
'功能：编译指定连接所有者的无效对象
'参数：cnThis=所有者连接,本函数可针对不同所有者调用
    Dim strErrInfor As String
    
    strErrInfor = ""
    
    Call SetProgressVisible(True)
    Call CompileAllInvalidObject(cnThis, strErrInfor, stbThis.Panels(2), pgbState)
    Call SetProgressVisible(False)
        
    If strErrInfor <> "" Then
        If Len(strErrInfor) > 300 Then strErrInfor = Mid(strErrInfor, 1, 300) & "..."
        MsgBox strErrInfor, vbInformation + vbDefaultButton1, gstrSysName
    End If
End Sub

Private Function ExeFuncCopy(ByVal strBakUserName As String, ByVal strBakUserPwd As String, ByVal strBakServer As String, ByVal strBakTBS As String) As Boolean
'功能：通过SQLPlus的Copy命令将远程数据库的非转储数据复制到当前历史表空间
'说明：创建本地临时文件，每张非转储表生成一条copy脚本命令，然后通过shell方式调用sqlplus来执行临时文件中的多条脚本。
'参数：strBakTBS=历史库用户的表空间名称
    Dim rsUnHistory As New ADODB.Recordset
    Dim objFSO As New FileSystemObject
    Dim objScript As Scripting.TextStream
    Dim strScript As String, strFile As String, strErrInfo As String
    Dim lngErrNum As Long, lngCommand As Long, i As Long, lngProcess As Long
    
    gstrSQL = "Select Table_Name From All_Tables Where Owner = '" & mstrOwnerName & _
            "' Minus Select 表名 From zlBakTables Order By Table_Name"
    'gstrSQL = "Select Table_Name From (" & gstrSQL & ") Where Table_Name like '人员表'"
    OpenRecordset rsUnHistory, gstrSQL, Me.Caption
    
    If rsUnHistory.RecordCount = 0 Then
        MsgBox "当前所有者中没有找到非转储表", vbInformation, gstrSysName
        Exit Function
    End If
    
    '生成临时脚本文件
    strFile = objFSO.GetSpecialFolder(TemporaryFolder).Path & "\" & objFSO.GetTempName
    
    Set objScript = objFSO.OpenTextFile(strFile, ForWriting, True)
    strScript = "set arraysize 5000"
    objScript.WriteLine strScript
    
    strScript = "copy from " & mstrOwnerName & "/" & mstrOwnerPass & "@" & gstrServer & _
                " to " & strBakUserName & "/" & strBakUserPwd & "@" & strBakServer & _
                " Replace Table_Name Using select * from Table_Name;"
                
    For i = 1 To rsUnHistory.RecordCount
        objScript.WriteLine Replace(strScript, "Table_Name", rsUnHistory!Table_Name)
        rsUnHistory.MoveNext
    Next
    objScript.WriteLine "exit;"
    objScript.Close

    '生成SQLPlus命令
    strScript = "sqlplus " & mstrOwnerName & "/" & mstrOwnerPass & "@" & gstrServer & " @" & strFile

    '执行Shell命令
    err.Clear: On Error Resume Next
    
    SetPromptText "正在通过sqlplus的Copy命令复制" & rsUnHistory.RecordCount & "张表的数据，请稍等。"
    
    If Not Image1.ToolTipText = "保留临时文件，名称复制到剪切板" Then
        lngCommand = Shell(strScript, vbHide)
    End If
    
    If err.Number <> 0 Then
        lngErrNum = err.Number '53:文件未找到
        strErrInfo = err.Description & IIf(lngErrNum = 53, ",请检查 sqlplus.exe 是否正确安装", "")
        err.Clear
        SetPromptText ""
        Call MsgBox("错误:" & lngErrNum & vbCrLf & strErrInfo, vbInformation, gstrSysName)
        Exit Function
    Else
        If lngCommand <> 0 Then
            lngProcess = OpenProcess(Process_Query_Information, False, lngCommand)
            Do
                Sleep 50
                GetExitCodeProcess lngProcess, lngCommand
                DoEvents
            Loop While lngCommand = Still_Active
            CloseHandle lngProcess
        End If
        SetPromptText "复制非转储表数据完成"
        ExeFuncCopy = True
    End If
    
    
    
    If Image1.ToolTipText = "保留临时文件，名称复制到剪切板" Then
        Call Clipboard.SetText(strFile)
    Else
        objFSO.DeleteFile strFile
        Set objFSO = Nothing
        
        rsUnHistory.MoveFirst
        For i = 1 To rsUnHistory.RecordCount
            SetPromptText "正在创建非转储表的约束和索引(" & i & "/" & rsUnHistory.RecordCount & ")：" & rsUnHistory!Table_Name
   
            '创建表结构相关的PK、UQ
            Call CreateConstraint(rsUnHistory!Table_Name, strBakTBS, mstrOwnerName, strBakUserName)
            '创建表结构索引IX
            Call CreateIndex(rsUnHistory!Table_Name, strBakTBS, mstrOwnerName, strBakUserName)
            
            DoEvents
            rsUnHistory.MoveNext
        Next
    End If
End Function

Private Function GetBakTableSpace(ByRef cnBakOracle As ADODB.Connection, ByVal strBakUser As String) As String
'功能：根据历史空间连接，返回指定历史空间用户的表空间名称
    Dim rsHistory As New ADODB.Recordset

    gstrSQL = "select 名称 from zlbakspaces where 所有者='" & strBakUser & "'"
    OpenRecordset rsHistory, gstrSQL, Me.Caption, , , cnBakOracle
    
    If rsHistory.RecordCount > 0 Then
        GetBakTableSpace = rsHistory!名称
    End If
End Function


Private Function ExeFuncMerge() As Boolean
'功能：合并列表中选择的空间，仅保留编号最小的空间。
'说明：1.先禁用保留空间上的约束和索引
'      2.然后从空间编号最小的开始，插入数据(仅zlbaktables中定义的表)到保留空间中
'      3.每插入完成一个空间，则删除一个空间的用户及表空间文件、zlbakspaces中的记录
'      4.所有空间的数据合并完成后，重建保留表空间的约束和索引
    Dim i As Long, lngLoop As Long
    Dim strKeepVersion As String, strKeepOwner As String, strMergeOwner As String
    Dim strPreTable As String, strTableSpace As String
    Dim strError As String, strTables As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsBakTables As New ADODB.Recordset
    Dim rsDelSpace As New ADODB.Recordset
    Dim blnDisibled As Boolean
    
    On Error GoTo errHandle
    '1.检查
    SetPromptText "正在检查要合并的表空间。"
    '1.1检查版本
    '------------------------------------------------------------------------------
    gstrSQL = ""
    For i = 1 To mrsMergeSpace.RecordCount
        If i = mrsMergeSpace.RecordCount Then
            gstrSQL = gstrSQL & "Select '" & mrsMergeSpace!所有者 & "' As 所有者, 版本号," & mrsMergeSpace!编号 & " as 编号 From " & mrsMergeSpace!所有者 & ".Zlbakinfo"
        Else
            gstrSQL = gstrSQL & "Select '" & mrsMergeSpace!所有者 & "' As 所有者, 版本号," & mrsMergeSpace!编号 & " as 编号 From " & mrsMergeSpace!所有者 & ".Zlbakinfo Union All" & vbCrLf
        End If
        mrsMergeSpace.MoveNext
    Next
    OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    rsTemp.Filter = "编号=" & mlng空间编号
    strKeepVersion = rsTemp!版本号
    strKeepOwner = rsTemp!所有者
    rsTemp.Filter = "编号<>" & mlng空间编号
    
    For i = 1 To rsTemp.RecordCount
        If strKeepVersion <> rsTemp!版本号 Then
            strError = rsTemp!所有者 & ":" & rsTemp!版本号
        End If
        strMergeOwner = strMergeOwner & ",'" & rsTemp!所有者 & "'"
        rsTemp.MoveNext
    Next
    If strError <> "" Then
        MsgBox "要保留的历史空间版本为" & strKeepVersion & ",与要合并的历史空间版本不一致:" & vbCrLf & strError & _
                vbCrLf & "可以通过[切换]操作来升级历史数据空间。"
        Exit Function
    End If
    strMergeOwner = Mid(strMergeOwner, 2)
    
    
    '1.2被合并的表空间中是否存在zlbaktables以外的表，如果存在，则提示这些数据将会在合并后被删除。
    '------------------------------------------------------------------------------
    Set rsBakTables = New ADODB.Recordset
    gstrSQL = "Select 表名 From zlBakTables Where 系统 = " & mlngSys & " Order By 表名"
    OpenRecordset rsBakTables, gstrSQL, Me.Caption
    
    gstrSQL = "Select Owner, Table_Name From All_Tables Where Owner In (" & strMergeOwner & ") And Table_Name<>'ZLBAKINFO' Order By Owner, Table_Name"
    Set rsDelSpace = New ADODB.Recordset
    OpenRecordset rsDelSpace, gstrSQL, Me.Caption
    
    '检查每个要合并的表空间，是否存在zlbaktables以外的表
    mrsMergeSpace.MoveFirst
    mrsMergeSpace.Filter = "编号<>" & mlng空间编号  '保留空间不用检查
    
    strError = ""
    For lngLoop = 1 To mrsMergeSpace.RecordCount
        rsDelSpace.Filter = "Owner='" & mrsMergeSpace!所有者 & "'"
        strTables = ""
        For i = 1 To rsDelSpace.RecordCount
            rsBakTables.Filter = "表名='" & rsDelSpace!Table_Name & "'"
            If rsBakTables.RecordCount = 0 Then
                '存在zlbaktables以外的表
                strTables = strTables & "," & rsDelSpace!Table_Name
            End If
            rsDelSpace.MoveNext
        Next
        If strTables <> "" Then
            strError = strError & mrsMergeSpace!所有者 & ":" & Mid(strTables, 2) & vbCrLf
        End If
        mrsMergeSpace.MoveNext
    Next
    If strError <> "" Then
        If MsgBox("检查发现合并的数据空间中存在非转出表，合并后这些表及数据将会被删除。" & vbCrLf & strError & vbCrLf _
            & "请确保已存在有效备份，或者这些数据不再需要。" & vbCrLf & "你确定要继续吗？", vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
            Exit Function
        End If
    End If
    
    
    '1.3检查字段的一致性(保留空间中的表比删除空间的表多，字段数量相同，类型相同，精度可容纳（保留>合并）)
    '------------------------------------------------------------------------------
    Set rsTemp = New ADODB.Recordset
    gstrSQL = "Select Owner, Table_Name, Column_Name, Data_Type, Decode(Data_Type, 'VARCHAR2', Data_Length, Data_Precision) 长度," & vbNewLine & _
                "       Data_Scale 数字精度" & vbNewLine & _
                "From All_Tab_Columns" & vbNewLine & _
                "Where Owner In ('" & strKeepOwner & "') And Table_Name In(Select 表名 From zlBakTables Where 系统 = " & mlngSys & ")" & vbNewLine & _
                "Order By Table_Name"
    OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    Set rsDelSpace = New ADODB.Recordset
    gstrSQL = "Select Owner, Table_Name, Column_Name, Data_Type, Decode(Data_Type, 'VARCHAR2', Data_Length, Data_Precision) 长度," & vbNewLine & _
                "       Data_Scale 数字精度" & vbNewLine & _
                "From All_Tab_Columns" & vbNewLine & _
                "Where Owner In (" & strMergeOwner & ") And Table_Name In(Select 表名 From zlBakTables Where 系统 = " & mlngSys & ")" & vbNewLine & _
                "Order By Owner, Table_Name"
    OpenRecordset rsDelSpace, gstrSQL, Me.Caption
    
    strTables = ""
    strError = ""
    strPreTable = ""
    mrsMergeSpace.MoveFirst
    mrsMergeSpace.Filter = "编号<>" & mlng空间编号  '保留空间不用检查
    
    For lngLoop = 1 To mrsMergeSpace.RecordCount
        rsDelSpace.Filter = "Owner='" & mrsMergeSpace!所有者 & "'"
        '按表名排了序
        strPreTable = ""
        
        For i = 1 To rsDelSpace.RecordCount
            SetPromptText "正在检查" & mrsMergeSpace!所有者 & "的表结构：" & rsDelSpace!Table_Name
            rsTemp.Filter = "Table_Name='" & rsDelSpace!Table_Name & "'"
            If rsTemp.RecordCount = 0 Then
                 '删除空间中的转出表在保留空间中不存在(下面逗号前的空格，用于后面Mid(xx,4)))
                strTables = strTables & "  , 缺少[" & rsDelSpace!Table_Name & "]"
            Else
                '检查字段
                rsTemp.Filter = "Table_Name='" & rsDelSpace!Table_Name & "' And Column_Name='" & rsDelSpace!Column_Name & "'"
                If rsTemp.RecordCount = 0 Then
                    '保留空间中缺字段
                    If strPreTable <> rsDelSpace!Table_Name Then
                        strError = strError & vbCrLf & rsDelSpace!Owner & "." & rsDelSpace!Table_Name & ":" & rsDelSpace!Column_Name
                        strPreTable = rsDelSpace!Table_Name
                    Else
                        strError = strError & "," & rsDelSpace!Column_Name
                    End If
                Else
                    '字段类型，长度，精度
                    If rsTemp!DATA_TYPE <> rsDelSpace!DATA_TYPE Then
                        If strPreTable <> rsDelSpace!Table_Name Then
                            strError = strError & vbCrLf & rsDelSpace!Owner & "." & rsDelSpace!Table_Name & ":" & rsDelSpace!Column_Name & " " & rsDelSpace!DATA_TYPE
                            strPreTable = rsDelSpace!Table_Name
                        Else
                            strError = strError & "," & rsDelSpace!Column_Name & " " & rsDelSpace!DATA_TYPE
                        End If
                    ElseIf rsDelSpace!DATA_TYPE = "VARCHAR2" Then
                        If rsTemp!长度 < rsDelSpace!长度 Then
                            If strPreTable <> rsDelSpace!Table_Name Then
                                strError = strError & vbCrLf & rsDelSpace!Owner & "." & rsDelSpace!Table_Name & ":" & rsDelSpace!Column_Name & " " & rsDelSpace!DATA_TYPE & "(" & rsDelSpace!长度 & ")"
                                strPreTable = rsDelSpace!Table_Name
                            Else
                                strError = strError & "," & rsDelSpace!Column_Name & " " & rsDelSpace!DATA_TYPE & "(" & rsDelSpace!长度 & ")"
                            End If
                        End If
                   ElseIf rsDelSpace!DATA_TYPE = "NUMBER" Then
                        If rsTemp!长度 < rsDelSpace!长度 Or rsTemp!数字精度 <> rsDelSpace!数字精度 Then
                            If strPreTable <> rsDelSpace!Table_Name Then
                                strError = strError & vbCrLf & rsDelSpace!Owner & "." & rsDelSpace!Table_Name & ":" & rsDelSpace!Column_Name & " " & rsDelSpace!DATA_TYPE & "(" & rsDelSpace!长度 & "," & rsDelSpace!数字精度 & ")"
                                strPreTable = rsDelSpace!Table_Name
                            Else
                                strError = strError & "," & rsDelSpace!Column_Name & " " & rsDelSpace!DATA_TYPE & "(" & rsDelSpace!长度 & "," & rsDelSpace!数字精度 & ")"
                            End If
                        End If
                    End If
                End If
            End If
            
            rsDelSpace.MoveNext
        Next
        mrsMergeSpace.MoveNext
    Next
    If strError <> "" Then
        If Len(strError) > 300 Then strError = Mid(strError, 4, 300) & "..."
        MsgBox "因合并空间以下结构差异导致不能继续，请先执行数据结构修正：" & strError, vbInformation, gstrSysName
        Exit Function
    End If
        
    
    '2.合并前的处理：禁用约束和索引(汇总表的除外，因为查询需要那些索引)
    '------------------------------------------------------------------------------------------------------------------
    DoEvents
    blnDisibled = True
    SetPromptText "正在禁用历史数据空间" & strKeepOwner & "的主键和唯一键约束…"
    Call SetConstraintStatus(mcnOracle, False, strKeepOwner)
    SetPromptText "正在禁用历史数据空间" & strKeepOwner & "的索引…"
    Call SetIndexStatus(mcnOracle, False, strKeepOwner)
    
    
    '3.执行合并
    '普通表的插入与删除(包括多个产品系统的)
    '汇总表的更新
    mrsMergeSpace.MoveFirst
    mrsMergeSpace.Filter = "编号<>" & mlng空间编号  '保留空间除外
    For lngLoop = 1 To mrsMergeSpace.RecordCount
        strTableSpace = mrsMergeSpace!名称
        strMergeOwner = mrsMergeSpace!所有者
        
        '3.1数据处理
        SetPromptText "正在合并历史数据空间" & strTableSpace & "的数据…"
        
        gstrSQL = "Zl1_Datamove_Merge(" & strKeepOwner & "," & strMergeOwner & ")"
        Call ExecuteProcedure(gstrSQL, Me.Caption)
        
        
        '3.2.删除指向当前用户的远程连接对象
        Call DropDBLinkOfUser(mcnOracle, strMergeOwner)
        
        '3.3.删除本系统所有者及对象(表、索引等)
        SetPromptText "正在删除历史数据空间" & strTableSpace & "的用户及相关对象…"
        gstrSQL = "drop user " & strMergeOwner & " cascade"
        mcnDBA.Execute gstrSQL
        
        
        '3.4.删除本系统数据表空间
        SetPromptText "正在删除历史数据表空间" & strTableSpace & "和数据文件…"
        If CheckTableSpaceIsUse("表空间", strTableSpace, strMergeOwner, mcnDBA) = False Then
            '没有其他用户使用，可以删除
            gstrSQL = "alter tablespace " & strTableSpace & " offline"
            mcnDBA.Execute gstrSQL
            gstrSQL = "drop tablespace " & strTableSpace & " including contents and datafiles cascade constraints"
            mcnDBA.Execute gstrSQL
        Else
            MsgBox "表空间" & strTableSpace & "存在其他用户的对象，请移动这些对象后手工删除表空间及文件。", vbInformation
        End If
            
        '3.5.删除历史数据空间目录(可能多个系统共用一个历史数据空间)
        gstrSQL = "delete zltools.zlbakspaces where 名称= '" & strTableSpace & "'"
        mcnOracle.Execute gstrSQL
        
        mrsMergeSpace.MoveNext
    Next
    
    
    '4.启用约束和索引，删除历史空间用户及表空间数据文件
    SetPromptText "正在启用历史数据空间" & strKeepOwner & "的索引…"
    Call SetIndexStatus(mcnOracle, True, strKeepOwner)
    SetPromptText "正在启用历史数据空间" & strKeepOwner & "的主键和唯一键约束…"
    Call SetConstraintStatus(mcnOracle, True, strKeepOwner)

    
    SetPromptText "合并处理完成。"
    MsgBox "合并历史数据空间成功！", vbInformation, gstrSysName
    ExeFuncMerge = True
    
    Exit Function
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
    
    If blnDisibled Then
        SetPromptText "正在启用历史数据空间" & strKeepOwner & "的索引…"
        Call SetIndexStatus(mcnOracle, True, strKeepOwner)
        SetPromptText "正在启用历史数据空间" & strKeepOwner & "的主键和唯一键约束…"
        Call SetConstraintStatus(mcnOracle, True, strKeepOwner)
    End If
End Function


Private Sub SetIndexStatus(ByRef cnThis As ADODB.Connection, ByVal blnEnable As Boolean, ByVal strOwner As String)
'功能:禁用或启用索引，禁用后提高历史库的数据插入速度
'     启用时，该过程执行要先于SetConstraintStatus，否则主键或唯一键字段存在无效索引会引发错误,ORA-14063
'参数:cnThis-连接对象
'     blnEnable-索引可用性，true-启用索引 false -禁用索引
'     strOwner=历史空间的所有者
    Dim rsTmp As New ADODB.Recordset
    Dim cmdTmp As New ADODB.Command
    Dim strSQL As String
    

    '基于规则优化加快SQL执行
    If blnEnable Then
        strSQL = "Select /*+ rule*/" & vbNewLine & _
                " 'alter index " & strOwner & ".' || a.Index_Name || ' Rebuild' Sql" & vbNewLine & _
                "From All_Indexes A, Zltools.Zlbaktables T" & vbNewLine & _
                "Where a.Owner = '" & strOwner & "' And a.Table_Name = t.表名 And t.系统 = " & mlngSys & " And t.直接转出 = 1 And a.Status = 'UNUSABLE' And a.Index_Type = 'NORMAL' And" & vbNewLine & _
                "      Not Exists" & vbNewLine & _
                " (Select 1 From All_Constraints C Where c.Owner = a.Owner And c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U'))"
    Else
        strSQL = "Select /*+ rule*/" & vbNewLine & _
                " 'alter index " & strOwner & ".' || a.Index_Name || ' unusable' Sql" & vbNewLine & _
                "From All_Indexes A, Zltools.Zlbaktables T" & vbNewLine & _
                "Where a.Owner = '" & strOwner & "' And a.Table_Name = t.表名 And t.系统 = " & mlngSys & " And t.直接转出 = 1 And a.Status = 'VALID' And a.Index_Type = 'NORMAL' And Not Exists" & vbNewLine & _
                " (Select 1 From All_Constraints C Where c.Owner = a.Owner And c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U'))"
    End If
    OpenRecordset rsTmp, strSQL, Me.Caption, , , cnThis
       
    Set cmdTmp.ActiveConnection = cnThis
    cmdTmp.CommandType = adCmdText
    
    On Error Resume Next
    While Not rsTmp.EOF
        strSQL = rsTmp!SQL
        cmdTmp.CommandText = strSQL
        cmdTmp.Execute
        
        If err.Number > 0 And blnEnable Then
            '如果该索引正在使用，则只能在线重建，比较慢
            If InStr(err.Description, "ORA-00054") > 0 Then
                err.Clear
                strSQL = Replace(rsTmp!SQL, "Rebuild", "Rebuild Online")
                cmdTmp.CommandText = strSQL
                cmdTmp.Execute
            Else
                Call MsgBox("错误:" & err.Description & vbCrLf & strSQL, vbInformation, "索引重建")
                err.Clear
            End If
        End If
        
        rsTmp.MoveNext
    Wend
End Sub

Private Sub SetConstraintStatus(ByRef cnThis As ADODB.Connection, ByVal blnEnable As Boolean, ByVal strOwner As String)
'功能:禁用或启用的约束，禁用后提高历史库的数据插入速度
'     禁用主键或唯一键则会删除对应的索引
'参数:cnThis-连接对象
'     blnEnable=true-启用约束,false-禁用约束
'     strOwner=历史空间的所有者

    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim cmdTmp As New ADODB.Command
        
    '历史库没有外键和其他约束，所以，这里全是主键或唯一键
    If blnEnable Then
        '注意：对于主键和唯一键，不能使用novalidate方式，这将导致对应的索引不创建
        strSQL = "Select " & vbNewLine & _
                " 'ALTER TABLE " & strOwner & ".'|| a.Table_Name || ' enable constraint ' || a.Constraint_Name Sql" & vbNewLine & _
                "From All_Constraints A, Zltools.Zlbaktables T" & vbNewLine & _
                "Where a.Owner = '" & strOwner & "' And a.Table_Name = t.表名 And t.系统 = " & mlngSys & " And t.直接转出 = 1 And a.Status = 'DISABLED'"
    Else
        strSQL = "Select " & vbNewLine & _
                " 'ALTER TABLE " & strOwner & ".' || a.Table_Name || ' disable constraint ' || a.Constraint_Name || Decode(a.Constraint_Type,'P',' Cascade drop index','U',' Cascade drop index','') Sql" & vbNewLine & _
                "From All_Constraints A, Zltools.Zlbaktables T, All_Tables b" & vbNewLine & _
                "Where a.Owner = '" & strOwner & "' And a.Table_Name = t.表名 And t.系统 = " & mlngSys & " And t.直接转出 = 1 And a.Status = 'ENABLED' And a.Table_Name = b.Table_Name And a.Owner = b.Owner And b.Iot_Type Is Null"
    End If
    OpenRecordset rsTmp, strSQL, Me.Caption, , , cnThis
    
    Set cmdTmp.ActiveConnection = cnThis
    cmdTmp.CommandType = adCmdText
    
    On Error Resume Next
    While Not rsTmp.EOF
        strSQL = rsTmp!SQL
        cmdTmp.CommandText = strSQL
        cmdTmp.Execute
        If err.Number > 0 Then
            Call MsgBox("错误:" & err.Description & vbCrLf & strSQL, vbInformation, "约束启用")
            err.Clear
        End If
        rsTmp.MoveNext
    Wend
End Sub

Private Function ExeFuncTrans() As Boolean
'功能：检查并执行表空间传输
    Dim strTbsName As String, strBakUserName As String, strBakNO As String, strBakUserPwd As String
    Dim strPath As String, strSplit As String
    Dim strServerFrom As String, strPWDFrom As String
    Dim rsTmp As ADODB.Recordset
    Dim lngLoop As Long
    
    On Error GoTo errHandle
    strTbsName = lvwHistory.SelectedItem.SubItems(C1名称)
    strBakUserName = lvwHistory.SelectedItem.SubItems(C4所有者)
    strBakNO = Mid(lvwHistory.SelectedItem.Key, 2)
    strBakUserPwd = txtBakPWD.Text
    strServerFrom = txtDbaServer.Text
    strPWDFrom = txtDba口令.Text
    
    '1.在源数据库建目录对象（要求必须指向表空间文件的位置）
    '--------------------------------------------------------------
    SetPromptText "在源库创建目录ZLTRANSFROM"
    gstrSQL = "Select 1 From Dba_Directories Where Directory_Name = 'ZLTRANSFROM'"
    Set rsTmp = New ADODB.Recordset
    OpenRecordset rsTmp, gstrSQL, Me.Caption, , , mcnDBA
    If rsTmp.RecordCount > 0 Then
        gstrSQL = "Drop DIRECTORY ZLTRANSFROM"
        mcnDBA.Execute gstrSQL
    End If
    
    gstrSQL = "Select File_Name From Dba_Data_Files Where Tablespace_Name = '" & strTbsName & "'"
    Set rsTmp = New ADODB.Recordset
    OpenRecordset rsTmp, gstrSQL, Me.Caption, , , mcnDBA
    
    strPath = rsTmp!file_name
    If InStr(strPath, "\") > 0 Then
        strSplit = "\"
    Else 'linux等平台是/
        strSplit = "/"
    End If
    strPath = Mid(strPath, 1, InStrRev(strPath, strSplit) - 1)

    gstrSQL = "CREATE DIRECTORY ZLTRANSFROM AS '" & strPath & "'"
    mcnDBA.Execute gstrSQL
'    gstrSQL = "GRANT READ, WRITE ON DIRECTORY ZLTRANSFROM TO SYSTEM"   '不用给自己授权
'    mcnDBA.Execute gstrSQL
    
        
    '2.源数据库传输表空间的索引是否存储在其他表空间
    gstrSQL = "Select Index_Name From Dba_Indexes Where Table_Owner = '" & strBakUserName & "' And Tablespace_Name <> '" & strTbsName & "'"
    Set rsTmp = New ADODB.Recordset
    Call OpenRecordset(rsTmp, gstrSQL, Me.Caption, , , mcnDBA)
    If rsTmp.RecordCount > 0 Then
        SetPromptText "源库有" & rsTmp.RecordCount & "个索引在其他表空间，正在重建到历史表空间" & strTbsName
        DoEvents
        For lngLoop = 1 To rsTmp.RecordCount
            gstrSQL = "alter index " & rsTmp!Index_Name & " rebuild tablespace " & strTbsName
            mcnDBA.Execute gstrSQL
            rsTmp.MoveNext
        Next
    End If
            
    
    
    '3.在目标库建传输目录，用户，数据库链路
    '----------------------------------------------------
     '3.1在目标库建传输目录
    SetPromptText "在目标库创建目录对象ZLTRANSTO"
    gstrSQL = "Select 1 From Dba_Directories Where Directory_Name = 'ZLTRANSTO'"
    Set rsTmp = New ADODB.Recordset
    OpenRecordset rsTmp, gstrSQL, Me.Caption, , , mcnOracle
    
    If rsTmp.RecordCount > 0 Then
        gstrSQL = "Drop DIRECTORY ZLTRANSTO"
        mcnOracle.Execute gstrSQL
    End If
    gstrSQL = "Select File_Name From Dba_Data_Files Where Tablespace_Name = 'ZLTOOLSTBS'"
    Set rsTmp = New ADODB.Recordset
    OpenRecordset rsTmp, gstrSQL, Me.Caption, , , mcnOracle
    
    strPath = rsTmp!file_name
    If InStr(strPath, "\") > 0 Then
        strSplit = "\"
    Else 'linux等平台是/
        strSplit = "/"
    End If
    strPath = Mid(strPath, 1, InStrRev(strPath, strSplit) - 1)
    
    gstrSQL = "CREATE DIRECTORY ZLTRANSTO AS '" & strPath & "'"
    mcnOracle.Execute gstrSQL
    
    
     '3.2在目标库建用户
    SetPromptText "在目标库创建历史空间用户" & strBakUserName
    '在cmdnext中已判断是否有同名用户
    gstrSQL = "create user " & strBakUserName & " identified by " & strBakUserPwd '& _
              '" DEFAULT TABLESPACE " & strTbsName   '表空间现在还不存在
    mcnOracle.Execute gstrSQL
    
    gstrSQL = "Grant Connect,Resource,UNLIMITED TABLESPACE," & _
            " Create Table,Create Sequence,Create Role,Create User,Drop User,Create Public Synonym,Drop Public Synonym," & _
            " Alter Session,Create Session,Create Synonym,Create View,Create Database Link,Create Cluster" & _
            " to " & strBakUserName & " With Admin Option"
    mcnOracle.Execute gstrSQL
        
        
     '3.3在目标库建数据库链路
    SetPromptText "在目标库创建数据库链路ZLTRANSTBS"
    gstrSQL = "Select 1 From Dba_Db_Links Where Db_Link||'.' Like 'ZLTRANSTBS.%' And Owner = '" & gstrUserName & "'"
    Set rsTmp = New ADODB.Recordset
    OpenRecordset rsTmp, gstrSQL, Me.Caption, , , mcnOracle
    If rsTmp.RecordCount > 0 Then
        gstrSQL = "Drop DATABASE LINK ZLTRANSTBS"
        mcnOracle.Execute gstrSQL
    End If
    
    gstrSQL = "CREATE DATABASE LINK ZLTRANSTBS CONNECT TO SYSTEM IDENTIFIED BY " & strPWDFrom & " USING '" & strServerFrom & "'"
    mcnOracle.Execute gstrSQL
    
    
    '4.1在目标库执行传输
    '--------------------------------------------------------------
    SetPromptText "正在传输历史数据空间" & strTbsName
    DoEvents
    gstrSQL = "DBMS_STREAMS_TABLESPACE_ADM.PULL_SIMPLE_TABLESPACE('" & strTbsName & "', 'ZLTRANSTBS', 'ZLTRANSTO')"
    'exec DBMS_STREAMS_TABLESPACE_ADM.pull_simple_tablespace(tablespace_name => ,database_link => ,directory_object => ,conversion_extension => )
    mcnOracle.Execute gstrSQL
    
    '4.2修改缺省表空间
    gstrSQL = "alter user " & strBakUserName & " DEFAULT TABLESPACE " & strTbsName
    mcnOracle.Execute gstrSQL
    
    '4.3在目标库植入
    gstrSQL = "Delete zltools.zlbakspaces where 系统=" & mlngSys & " And 编号=" & strBakNO  '为了支持失败后的再次转移
    mcnOracle.Execute gstrSQL
    Call ExeFuncImport(mcnOracle, strBakNO, strTbsName, strBakUserName, mlngSys)
    
    '4.4在目标库删除创建的链路
    gstrSQL = "Drop DATABASE LINK ZLTRANSTBS"
    mcnOracle.Execute gstrSQL
    
    '4.5在目标库删除目录
    gstrSQL = "Drop DIRECTORY ZLTRANSTO"
    mcnOracle.Execute gstrSQL
    
    '4.6在源库删除目录
    gstrSQL = "Drop DIRECTORY ZLTRANSFROM"
    mcnDBA.Execute gstrSQL
    
    
    
    '5.在源库中删除已传输的表空间
    '-----------------------------------------------------------------
    DoEvents
    '5.1.删除本系统所有者及对象(表、索引等)
    SetPromptText "正在删除源库历史数据空间" & strTbsName & "的用户及相关对象…"
    gstrSQL = "drop user " & strBakUserName & " cascade"
    mcnDBA.Execute gstrSQL
    
    
    '5.2.删除本系统数据表空间
    SetPromptText "正在删除历史数据表空间" & strTbsName & "和数据文件…"
    gstrSQL = "alter tablespace " & strTbsName & " offline"
    mcnDBA.Execute gstrSQL
    gstrSQL = "drop tablespace " & strTbsName & " including contents and datafiles cascade constraints"
    mcnDBA.Execute gstrSQL
        
    '5.3.删除历史数据空间管理记录
    gstrSQL = "delete zltools.zlbakspaces where 系统= " & mlngSys & " and 编号=" & strBakNO
    mcnDBA.Execute gstrSQL
    
    
    ExeFuncTrans = True
    Exit Function
errHandle:
    If InStr(err.Description, "ORA-00900") > 0 Then
        Call MsgBox("空间传输失败，在目标服务器的SQLPlus中执行以下SQL，查看详细的错误信息：" & vbCrLf & gstrSQL, vbInformation, "空间传输")
    Else
        Call MsgBox("错误:" & err.Description & vbCrLf & gstrSQL, vbInformation, "空间传输")
    End If
    
End Function

