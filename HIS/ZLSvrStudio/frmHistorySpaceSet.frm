VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHistorySpaceSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ʷ���ݿռ�����"
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
   StartUpPosition =   2  '��Ļ����
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
         Caption         =   "��ʷ���ݿռ���û�"
         Height          =   2955
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   1080
         Width           =   6570
         Begin VB.CommandButton cmd���� 
            Caption         =   "����(&T)"
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
            Caption         =   "Զ�̷�����"
            Height          =   255
            Index           =   1
            Left            =   1185
            TabIndex        =   8
            Top             =   1100
            Width           =   1215
         End
         Begin VB.OptionButton optServer 
            Caption         =   "��ǰ������"
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
         Begin VB.TextBox txtDba���� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4065
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   360
            Width           =   1635
         End
         Begin VB.TextBox txtDba�û� 
            Height          =   300
            Left            =   1185
            MaxLength       =   100
            TabIndex        =   3
            Top             =   360
            Width           =   1635
         End
         Begin VB.CommandButton cmd���� 
            Caption         =   "��ʷ��������(&U)"
            Height          =   350
            Left            =   1080
            TabIndex        =   15
            Top             =   2470
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label lblDBLinkPrompt 
            Caption         =   "Oracle��֧��ͨ��DBLink��������XMLType�ȶ������ͻ��û����������ֶεı����ԣ���֧��ֱ��ת����Զ����ʷ��"
            ForeColor       =   &H00404040&
            Height          =   855
            Left            =   4000
            TabIndex        =   114
            Top             =   1200
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label lblPWDPrompt 
            Caption         =   "��¼���ݿ�����룬����ת��"
            Height          =   255
            Index           =   1
            Left            =   4080
            TabIndex        =   113
            Top             =   795
            Width           =   2415
         End
         Begin VB.Label lblServerName 
            AutoSize        =   -1  'True
            Caption         =   "DBLink����"
            Height          =   180
            Index           =   1
            Left            =   165
            TabIndex        =   11
            Top             =   1830
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lblServer 
            Caption         =   "λ��"
            Height          =   255
            Left            =   705
            TabIndex        =   6
            Top             =   830
            Width           =   375
         End
         Begin VB.Label lblIniModi 
            AutoSize        =   -1  'True
            Caption         =   "�޸ġ�"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��װ�����ļ���C:\Appsoft\ZLHIS10\Ӧ�ýű�\ZLSETUP.INI"
            Height          =   180
            Left            =   240
            TabIndex        =   13
            Top             =   2160
            Visible         =   0   'False
            Width           =   4770
         End
         Begin VB.Label lblServerName 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   9
            Top             =   1470
            Width           =   720
         End
         Begin VB.Label lblDba 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   1
            Left            =   3580
            TabIndex        =   4
            Top             =   420
            Width           =   360
         End
         Begin VB.Label lblDba 
            AutoSize        =   -1  'True
            Caption         =   "�û���"
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
         Caption         =   "   ��������ָ�������ӵ�ָ������ʷ���ݿռ�ķ������Ӵ�,�ô����ڱ���Oracle��TnsNames�ļ������á�"
         Height          =   405
         Index           =   13
         Left            =   2400
         TabIndex        =   90
         Top             =   2520
         Width           =   4380
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "    �������ݿ������������Ϣ��"
         Height          =   180
         Index           =   0
         Left            =   780
         TabIndex        =   0
         Top             =   750
         Width           =   2700
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "��һ�� ָ��DBA�û�"
         BeginProperty Font 
            Name            =   "����"
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
      Begin VB.CheckBox chk��ǰ 
         Caption         =   "��Ϊ��ǰ�����ݿռ�(&D)"
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
         Caption         =   "����:"
         Height          =   180
         Left            =   5205
         TabIndex        =   89
         Top             =   1995
         Width           =   450
      End
      Begin VB.Label lblDataVer 
         AutoSize        =   -1  'True
         Caption         =   "����:"
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
         Caption         =   "�ռ�����"
         Height          =   180
         Index           =   9
         Left            =   1365
         TabIndex        =   84
         Top             =   1830
         Width           =   720
      End
      Begin VB.Label lblNoteImport 
         Caption         =   "    ���ñ�ֲ�����ʷ���ݿռ�ı����Ϣ���ռ����ơ�"
         Height          =   330
         Left            =   1425
         TabIndex        =   81
         Top             =   855
         Width           =   5955
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ռ��û�"
         Height          =   180
         Index           =   5
         Left            =   1365
         TabIndex        =   86
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "���"
         Height          =   180
         Index           =   6
         Left            =   1725
         TabIndex        =   82
         Top             =   1305
         Width           =   360
      End
      Begin VB.Label lblStepImport 
         AutoSize        =   -1  'True
         Caption         =   "�ڶ��� ������ʷ���ݿռ�"
         BeginProperty Font 
            Name            =   "����"
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
         TabCaption(0)   =   "����"
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
         Tab(0).Control(14)=   "txt���"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "txtOwnerLab"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txtOwnerPwd"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "txtOwnerUsr"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "chkCreate��ǰ"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).ControlCount=   19
         TabCaption(1)   =   "�����ļ�"
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
         Begin VB.CheckBox chkCreate��ǰ 
            Caption         =   "��������Ϊ��ǰ�ռ�(&C)"
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
            Caption         =   "�Զ���չ�ռ�"
            Height          =   270
            Left            =   -69480
            TabIndex        =   39
            ToolTipText     =   "AUTOEXTEND ON Next (��ռ��С/10)M"
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
            ToolTipText     =   "AUTOALLOCATE �� UNIFORM Size nM"
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
         Begin VB.TextBox txt��� 
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
            Caption         =   "��¼���ݿ�����룬����ת��"
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   112
            Top             =   1605
            Width           =   3150
         End
         Begin VB.Label lblBakPrompt 
            Caption         =   "���鰴ת����ֹ��������,����:201412"
            Height          =   255
            Left            =   3720
            TabIndex        =   111
            Top             =   1106
            Width           =   3150
         End
         Begin VB.Label lblFileAmount 
            AutoSize        =   -1  'True
            Caption         =   "������     ���ļ�"
            Height          =   180
            Index           =   2
            Left            =   -70440
            TabIndex        =   50
            Top             =   2520
            Width           =   1530
         End
         Begin VB.Label lblBakSpaceLob 
            Alignment       =   1  'Right Justify
            Caption         =   "������ռ���"
            Height          =   225
            Left            =   -74205
            TabIndex        =   48
            Top             =   2505
            Width           =   1365
         End
         Begin VB.Label lblFileAmount 
            AutoSize        =   -1  'True
            Caption         =   "������     ���ļ�"
            Height          =   180
            Index           =   1
            Left            =   -70440
            TabIndex        =   46
            Top             =   2100
            Width           =   1530
         End
         Begin VB.Label lblFileAmount 
            AutoSize        =   -1  'True
            Caption         =   "������     ���ļ�"
            Height          =   180
            Index           =   0
            Left            =   -73380
            TabIndex        =   35
            Top             =   1260
            Width           =   1530
         End
         Begin VB.Label lblBakSpaceIdx 
            Alignment       =   1  'Right Justify
            Caption         =   "������ռ���"
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
            Caption         =   "˵��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "   ��������ָ���߷��������ӵ��������ݷ������е����Ӵ�,���߷������Ļ����ϱ������øô�!"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   60
            TabIndex        =   91
            Top             =   3960
            Width           =   7650
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������"
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
            Caption         =   "���ݱ�ռ���"
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
            Caption         =   "Զ��������"
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
            Caption         =   "�û���"
            Height          =   180
            Index           =   1
            Left            =   870
            TabIndex        =   20
            Top             =   1143
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "���"
            Height          =   180
            Index           =   0
            Left            =   1050
            TabIndex        =   18
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lblNewPwd 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   1050
            TabIndex        =   22
            Top             =   1551
            Width           =   360
         End
         Begin VB.Label lblNewLab 
            AutoSize        =   -1  'True
            Caption         =   "��֤"
            Height          =   180
            Left            =   1035
            TabIndex        =   24
            Top             =   2004
            Width           =   360
         End
         Begin VB.Label lblFileSize 
            AutoSize        =   -1  'True
            Caption         =   "��ʼ��С          M"
            Height          =   180
            Left            =   -71430
            TabIndex        =   37
            Top             =   1245
            Width           =   1710
         End
         Begin VB.Label lblDataFile 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��һ���ļ�"
            Height          =   180
            Left            =   -73740
            TabIndex        =   33
            Top             =   885
            Width           =   900
         End
         Begin VB.Label lblSpaceExtend 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "���ߴ�"
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
         Caption         =   "�ڶ��� ������ʷ���ݿռ�"
         BeginProperty Font 
            Name            =   "����"
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
            Text            =   "���"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "��ǰ"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "���"
            Text            =   "ֻ��"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "������"
            Text            =   "������"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "�汾��"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "���ת������"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "���������"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblBakPWD 
         AutoSize        =   -1  'True
         Caption         =   "Ŀ����½�����ʷ�ռ��û�����"
         Height          =   180
         Left            =   1080
         TabIndex        =   110
         Top             =   3540
         Width           =   2520
      End
      Begin VB.Label Label2 
         Caption         =   "˵����"
         Height          =   375
         Left            =   360
         TabIndex        =   97
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblStepTrans 
         AutoSize        =   -1  'True
         Caption         =   "�ڶ�����ѡ��Դ�������ϵ���ʷ���ݿռ�"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "Դ��ʷ���ݿռ����������"
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
      Caption         =   "����(&H)"
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
      Caption         =   "��һ��(&B)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4125
      TabIndex        =   54
      Top             =   4095
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
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
            Text            =   "����"
            TextSave        =   "18:06"
            Key             =   "STANUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ��(&N)"
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
            Caption         =   "������ʷ���ݿռ�(&1)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "ɾ����ʷ���ݿռ�(&2)"
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   69
            Top             =   300
            Value           =   -1  'True
            Width           =   2385
         End
         Begin VB.Label lblDelInfor 
            Caption         =   "ֻ�����ѡ�����ʷ���ݿռ����ƣ���ص���ʷ���ݲ������ݿ�ɾ����"
            ForeColor       =   &H8000000C&
            Height          =   285
            Index           =   1
            Left            =   465
            TabIndex        =   72
            Top             =   1245
            Width           =   5775
         End
         Begin VB.Label lblDelInfor 
            Caption         =   "���״����ݿ���ɾ����ص���ʷ����"
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
         Caption         =   "������:"
         Height          =   180
         Left            =   3780
         TabIndex        =   76
         Top             =   3285
         Width           =   630
      End
      Begin VB.Label lblDbLink 
         Caption         =   "DB���ӣ�X23423"
         Height          =   180
         Left            =   1080
         TabIndex        =   75
         Top             =   3510
         Width           =   4335
      End
      Begin VB.Label lblSpace 
         Caption         =   "���ƣ�zlbak0701"
         Height          =   180
         Left            =   1080
         TabIndex        =   74
         Top             =   3270
         Width           =   4335
      End
      Begin VB.Label lblCode 
         Caption         =   "��ţ�200"
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
         Caption         =   "��һ����ѡ����ɾ���������ʷ���ݿռ�"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "    ɾ����ʷ���ݿռ���Լ���ÿ�α��ݵĺ�ʱ����������ִ��ǰ��ȷ����Щ��������ֲ���������ݿ⣬����������Ч�ı��ݡ�"
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
         Caption         =   "��һ�� ���ϲ�����ʷ���ݿռ�"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�����ռ���"
         Height          =   180
         Index           =   16
         Left            =   1005
         TabIndex        =   106
         Top             =   1305
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ϲ��ռ�"
         Height          =   180
         Index           =   10
         Left            =   1365
         TabIndex        =   105
         Top             =   2190
         Width           =   720
      End
      Begin VB.Label lblNoteMerge 
         Caption         =   "    �ϲ�����ʷ���ݿռ�ı����Ϣ���ռ����ơ�"
         Height          =   450
         Left            =   960
         TabIndex        =   104
         Top             =   840
         Width           =   5955
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�����ռ�����"
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
Private mlngSys  As Long                   'ϵͳ���
Private mstrSysName   As String                 'ϵͳ����
Private mstrOwnerName As String
Private mstrVersion   As String                 '�汾��
Private mstrOwnerPass As String
Private mstrDBLink As String                    '��ʷ�ռ��DBLink

Private mcnDBA As New ADODB.Connection      'DBA�û�����ʷ�ռ��û�������
Private mcnOracle As New ADODB.Connection   '��ǰ��¼�û������Ӷ���

Private mblnFirst As Boolean
Private mblnSucced As Boolean
Private mlngOracleVer As Long

Private Enum ENUFT
    F0���� = 0
    F1��ж = 1
    F2��ֲ = 2
    F3���� = 3
    F4�л� = 4
    F5�ϲ� = 5
    F6ת�� = 6
End Enum
Private mintFunType        As ENUFT  '0-������ʷ���ݿռ�,1-��ж��ʷ���ݿռ�,2-��ֲ��ʷ���ݿռ�,3-���Ʒ�ת������,
                                     '4���л��ڵ�ǰ����ʷ���ݿռ�,5-�ϲ���ʷ���ݿռ�,6-ת����ʷ���ݿռ�

Private Enum ENUCOL
    C0��� = 0
    C1���� = 1
    C2��ǰ = 2
    C3ֻ�� = 3
    C4������ = 4
    C5�汾�� = 5
    C6���ת������ = 6
    C7��������� = 7
End Enum

Private mlng�ռ���          As Long
Private mblnMustInstall As Boolean  '���谲װ�ռ�
Private mstr�ϲ��ռ��� As String
Private mrsMergeSpace As ADODB.Recordset
Private mblnSysUpdateCall As Boolean '�Ƿ�ϵͳ��������


Public Function ShowInstall(ByVal frmMain As Form, ByVal cnOracle As ADODB.Connection, _
    ByVal strOwner As String, ByVal strOwnerPass As String, _
    ByVal lngϵͳ As Long, ByVal intFunType As Integer, _
    ByVal lng�ռ��� As Long, Optional ByVal str�ϲ��ռ��� As String, _
    Optional ByVal blnSysUpdateCall As Boolean) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------
    '����:��ʷ���ݿռ����ӿ�
    '����:cnOracle-ϵͳ����
    '     strOwner-�������û���
    '     strOwnerPass-����������
    '     lngϵͳ-ϵͳ��
    '     intFunType-��������,��mintFunType
    '               ����4-�л�,5-�ϲ��������Ϳ�ʼִ�У�ֻ���������������ʾ���ȡ�
    '     lng�ռ���-intFunType=1ʱ=��ж���ݿռ�Ŀռ���
    '             intFunType=5ʱ=���������ݿռ�ı��
    '             intFunType=6ʱ=��0������Ҫ��
    '             blnSysUpdateCall=�Ƿ���ϵͳ�������ã�����ǣ��������Լ��������ߡ�
    '����:��װ�ɹ�,����true,���򷵻�False
    '----------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    mintFunType = intFunType
    
    Set mcnOracle = cnOracle
    mstrOwnerPass = strOwnerPass
    mblnSysUpdateCall = blnSysUpdateCall
    mblnMustInstall = False
    mstr�ϲ��ռ��� = str�ϲ��ռ���
    
    If mintFunType = F0���� Then
        If IsHavingHistoryTable(lngϵͳ) = False Then
            If frmMain Is frmHistoryDataMgr Then
                MsgBox "��������ʷ���ݿռ��������ݱ����ܴ�����ʷ���ݿռ�,����!", vbInformation, gstrSysName
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
    ElseIf mintFunType = F1��ж Then
        If IsHavingHistoryTable(lngϵͳ) = False Then
            ShowInstall = True
            Exit Function
        End If
        If Not frmMain Is frmHistoryDataMgr Then
            mblnMustInstall = True
        End If
    ElseIf mintFunType = F6ת�� Then
        If gstrUserName <> "SYSTEM" Then
            MsgBox "��ʷ�ռ�ת�ƹ��ܱ�����SYSTEM���û�ִ�У���ǰ�û�����SYSTEM�������µ�¼!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    mlngSys = lngϵͳ
    
    If mintFunType = F1��ж And Not frmMain Is frmHistoryDataMgr Then
        'ϵͳ����
              
    Else
        gstrSQL = "select ������,�汾��,���� from zlSystems where ���=" & mlngSys
        Call OpenRecordset(rsTemp, gstrSQL, "��ȡ������")
        If Not rsTemp.EOF Then
            mstrOwnerName = Nvl(rsTemp!������)
            mstrVersion = Nvl(rsTemp!�汾��)
            mstrSysName = Nvl(rsTemp!����)
        Else
            MsgBox "ϵͳ������,���ܱ����˲�ж,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        If strOwner <> mstrOwnerName And mintFunType <> F6ת�� Then
            MsgBox "�㲻�ǵ�ǰӦ�ó����������,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If
    If mstrVersion <> "" Then
        If Val(Split(mstrVersion, ".")(0)) < 10 Then
                MsgBox "��֧��9���µİ汾,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
        End If
    End If
    mlng�ռ��� = lng�ռ���
        
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
          
    lblDataVer.Caption = "����:" & mstrVersion
    lblDataVer.Tag = mstrVersion
    
    mlngOracleVer = GetOracleVersion(True, True)
    
End Sub


Private Function InitCtronl() As Boolean
'����:���ÿؼ��ɼ��ԺͿ����Լ�����˵����ȱʡfraSetup�ȿ�Ƭ���ǲ��ɼ�״̬
    Dim bytErr As Byte, strErrMsg As String, strDbLink As String
    
    cmdPrevious.Enabled = False
    
    txtDbaServer.Text = gstrServer  'Ŀǰֻ�и��ơ���ֲ��ת�ƹ�������ָ��Զ�̷�����
    txtDbaServer.Enabled = False
    
    Select Case mintFunType
    Case F0����   '-�����ռ�
        Me.Caption = "������ʷ���ݿռ�"
        
        If CheckIsDBA(mcnOracle) Then
            Set mcnDBA = mcnOracle
            
            fraSetup(0).Visible = False 'ֱ���õ�ǰ������������ʷ���ݿռ䣬������ʾDBA���ӽ���
            fraSetup(1).Visible = True
            
            Call InitCreateTbs
        Else
            fraSetup(0).Visible = True
            fra(0).Caption = "���ڴ�����ʷ���ݿռ��DBA�û�"
        
            If mblnMustInstall Then
                txtDba�û�.Text = gstrUserName
                txtDba����.Text = gstrPassword
            End If
            If fraSetup(0).Visible Then
                If txtDba�û�.Text <> "" Then
                    If txtDba����.Enabled And txtDba����.Visible Then txtDba����.SetFocus
                Else
                    If txtDba�û�.Enabled And txtDba�û�.Visible Then txtDba�û�.SetFocus
                End If
            End If
            
            'Oracle��֧��ͨ��DBLink��������XMLType�ȶ������ͻ��û����������ֶεı����ԣ���֧��ֱ�Ӳ���Զ����ʷ��
            optServer(1).Enabled = False
            lblDBLinkPrompt.Visible = True
            
            '�л�����ֲʱ����Ҫ��ʷ����������������ļ�
            lblSetupIni.Visible = False
            lblIniModi.Visible = False
        
        End If
        
        
        If mblnMustInstall Then
            chkCreate��ǰ.value = 1
            chkCreate��ǰ.Enabled = False
        Else
            chkCreate��ǰ.Enabled = True
        End If
        
        '��ռ�����������
        cboSpaceExtentType.Clear
        cboSpaceExtentType.addItem "�Զ��������ߴ�"
        cboSpaceExtentType.addItem "ͳһ�������ߴ�"
        cboSpaceExtentType.ListIndex = 0
        
        txtSpaceExtentSize.Text = 1
        txtSpaceExtentSize.Enabled = (cboSpaceExtentType.ListIndex = 1)
    
        InitCtronl = True
        
    Case F1��ж   '��ж�ռ�
        Me.Caption = "ж����ʷ���ݿռ�"
        fraDelete.Visible = True
        
        lblDBLinkPrompt.Visible = False
                        
        If LoadSpaceData Then
            optDele(0).value = True 'ȱʡΪɾ��ģʽ
            lblStep(0).Caption = "�ڶ���:ָ��DBA�û�"
            lblNote(0).Caption = "    ������Զ�̷������ϵ���ʷ���ݿռ��DBA�û���"
            InitCtronl = True
        End If
    Case F2��ֲ   '��ֲ�ռ�
        Me.Caption = "��ֲ��ʷ���ݿռ�"
        fraSetup(0).Visible = True
        
        lblStep(0).Caption = "��һ�� ָ����ʷ���ݿռ��û�"
        fra(0).Caption = "��ʷ���ݿռ��û���������Ϣ"
        lblNote(0).Caption = "    �����������ԭ���µ�ǰû������������ʷ���ݿռ������������"
        lblServerName(1).Caption = "����DBLink"
        
        txtDba�û�.Text = "ZLHD"
        If txtDba�û�.Enabled And txtDba�û�.Visible Then txtDba�û�.SetFocus

        InitCtronl = True
    Case F3����   '���Ʒ�ת�����ݱ�����
        Me.Caption = "���Ʒ�ת������"
        fraSetup(0).Visible = True
        optServer(1).Enabled = False
        
        If LoadSpaceData Then
            txtDba�û�.Text = lblSpaceOwner.Tag
            If txtDba�û�.Enabled And txtDba�û�.Visible Then txtDba�û�.SetFocus
            txtDbaServer.Enabled = True     '�������ݸ��Ƶ�Զ�����ݿ⣬��Ϊʹ��Copy��������ڷֲ�ʽ����
                        
            lblStep(0).Caption = "��һ�� ָ����ʷ���ݿռ��û�"
            lblNote(0).Caption = "    �����з�ת��������ݸ��Ƶ�������ʷ���ݿռ���(������Զ��)��"
            fra(0).Caption = "��ʷ���ݿռ���û�"
            cmdNext.Caption = "����(&F)"
                     
            InitCtronl = True
        End If
    Case F4�л�
        Me.Caption = "�л���ǰ��ʷ���ݿռ�"
        fraSetup(0).Visible = True
        
        If LoadSpaceData Then
            txtDba�û�.Text = lblSpaceOwner.Tag
            txtDba�û�.Enabled = False
            txtDba����.Enabled = False
            cmd����.Enabled = False
            optServer(0).Enabled = False
            optServer(1).Enabled = False
            txtDbaServer.Enabled = False
            txtDBLink.Enabled = False
                        
            If optServer(1).value = True Then
                strDbLink = Trim(txtDBLink.Text)
            End If
                         
            lblStep(0).Caption = "    �л���ǰ��ʷ���ݿռ�Ϊ" & lblSpace.Tag & "��"
            fra(0).Caption = "��ʷ���ݿռ���û�"
            
            'ִ���л�
            Call SetControlEnable(False)
            If ExeFuncChange(Trim(txtDba�û�.Text), mstrOwnerName, mlngSys, bytErr, strErrMsg, strDbLink) = False Then
                Call SetControlEnable(True)
                '1-����ʧЧ,2-ϵͳ������,3-���߰汾������ʷ�汾,4-���߰汾С����ʷ�汾
                Select Case bytErr
                Case 2, 4
                    MsgBox strErrMsg, vbInformation + vbDefaultButton1, gstrSysName
                    Unload Me
                    Exit Function
                Case 3
                    MsgBox strErrMsg, vbInformation + vbDefaultButton1, gstrSysName
                    Call ReadSetupIni(1)
                    cmd����.Visible = True
                    lblIniModi.Visible = cmd����.Visible: lblSetupIni.Visible = cmd����.Visible
                    cmd����.Visible = False
                    txtDba�û�.Enabled = False
                    txtDba����.Enabled = True
                    
                    Me.cmdNext.Caption = "�л�(&Q)"
                    InitCtronl = True
                    If cmd����.Enabled And cmd����.Visible Then
                        lblIniModi.Visible = True: lblSetupIni.Visible = True
                        cmd����.SetFocus
                    End If
                    Exit Function
                Case 1
                    '���Ӳ��ɹ�,��Ҫ����������ص��û���������
                    MsgBox strErrMsg, vbInformation + vbDefaultButton1, gstrSysName
                    txtDba�û�.Enabled = True
                    txtDba����.Enabled = True
                    
                    cmd����.Visible = False
                    
                    If optServer(1).value = True Then txtDBLink.Enabled = True
              
                End Select
            Else
                Call SetControlEnable(True)
                mblnSucced = True
            End If
            
            Me.cmdNext.Caption = "�л�(&Q)"
            InitCtronl = True
            Unload Me
        End If
    Case F5�ϲ�
        Me.Caption = "�ϲ���ʷ���ݿռ�"
        fraSetup(0).Visible = True
        
        'Oracle��֧��ͨ��DBLink��������XMLType�ȶ������ͻ��û����������ֶεı����ԣ���֧��ֱ�Ӳ���Զ����ʷ��
        optServer(1).Enabled = False
        lblDBLinkPrompt.Visible = True
        
        If LoadSpaceData Then
            lblStep(0).Caption = "��һ�� ָ��DBA�û�"
            lblNote(0).Caption = "    ָ��DBA�û�������Ϣ������ɾ����ʷ���ݿռ�������߼���ռ��ļ���"
            fra(0).Caption = "DBA�û�������Ϣ"
        
            cmdNext.Caption = "�ϲ�(&Q)"
            
            If txtDba�û�.Text <> "" Then
                If txtDba����.Enabled And txtDba����.Visible Then txtDba����.SetFocus
            Else
                If txtDba�û�.Enabled And txtDba�û�.Visible Then txtDba�û�.SetFocus
            End If
            InitCtronl = True
        End If
    Case F6ת��
        Me.Caption = "ת����ʷ���ݿռ�"
        fraSetup(0).Visible = True
        
        lblStep(0).Caption = "��һ�� ָ��Դ��������DBA�û�"
        lblNote(0).Caption = "    ����Դ������(�������ڱ���Tnsnames�д���)���汾�Ƿ������ƽ̨�Ƿ�֧�֡�"
        fra(0).Caption = "Զ�̷�����������Ϣ"
        
        txtDba�û�.Text = "SYSTEM"
        txtDba�û�.Enabled = False
        txtDba�û�.BackColor = &H8000000F
        If txtDba����.Enabled And txtDba����.Visible Then txtDba����.SetFocus
        txtDbaServer.Enabled = True
        optServer(1).Enabled = False
        
        InitCtronl = True
    End Select
End Function

Private Sub ReadSetupIni(ByVal intIndex As Integer)
'���ܣ���ȡϵͳ��װ�����ļ�
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strSetup As String, strTmp As String
    
    
    '��ȡ��װ�����ļ�
    strSQL = "Select A.�ļ��� From Zlsysfiles a Where  A.����=1 And ϵͳ=" & mlngSys
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    If Not rsTmp.EOF Then
        If gobjFile.FileExists(rsTmp!�ļ��� & "") Then
            strSetup = rsTmp!�ļ��� & ""
        End If
    End If
    If strSetup = "" Then
        strTmp = gobjFile.GetParentFolderName(App.Path) & "\" & Decode(mlngSys \ 100, 1, "ZLHIS10", 3, "ZLMEDREC10", 4, "ZLMATERIAL10", _
                                                                                6, "ZLDEVICE10", 21, "ZLPEIS10", 22, "ZLBLOOD10", _
                                                                                23, "ZLINFECT10", 24, "ZLOPER10", _
                                                                                25, "ZLLIS10", 26, "ZLPSS10", 27, "ZLHEC10") & "\Ӧ�ýű�\ZLSETUP.INI"
        If gobjFile.FileExists(strTmp) Then
            strSetup = strTmp
        End If
    End If
    If strSetup <> "" Then
        If Not CheckInitFile(mlngSys, strSetup) Then
            strSetup = ""
        End If
    End If
    lblSetupIni.Caption = "��װ�����ļ���" & strSetup
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
    '����:���ؿռ�������Ϣ
    '-------------------------------------------------------------------------------
    Dim i As Long, str�ռ����� As String
    Dim rsbakspaces As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim strImgKey As String
    Dim objItem As ListItem
    Dim lngMaxLen As Long
    
    If mintFunType = F1��ж And mblnMustInstall Then
        gstrSQL = "Select ���,����,������,DB����,��ǰ From ZLTOOLS.zlbakspaces where ϵͳ=" & mlngSys & " and  ��ǰ=1"
    
    ElseIf mintFunType = F5�ϲ� Then
        gstrSQL = "Select ���,����,������,DB����,��ǰ From ZLTOOLS.zlbakspaces where ��� in(" & mstr�ϲ��ռ��� & "," & mlng�ռ��� & ") Order by ���"
    
    ElseIf mintFunType = F6ת�� Then
        
        gstrSQL = "Select max(length(���)) as MaxLen From zltools.zlbakspaces where ϵͳ=" & mlngSys
        OpenRecordset rsbakspaces, gstrSQL, "��ȡ��ʷ���ݿռ�", , , mcnDBA
        lngMaxLen = Val(Nvl(rsbakspaces!MaxLen))
    
        gstrSQL = "Select ���,����,������,DB����,��ǰ,ֻ�� From ZLTOOLS.zlbakspaces where ϵͳ=" & mlngSys
    Else
        gstrSQL = "Select ���,����,������,DB����,��ǰ From ZLTOOLS.zlbakspaces where ϵͳ=" & mlngSys & " and  ���=" & mlng�ռ���
    End If
    
    If mintFunType = F6ת�� Then
        OpenRecordset rsbakspaces, gstrSQL, Me.Caption, , , mcnDBA
    Else
        OpenRecordset rsbakspaces, gstrSQL, Me.Caption
    End If
    
    
    
    
    If mintFunType = F5�ϲ� Then
        If rsbakspaces.RecordCount = 0 Then
           MsgBox "��ʷ���ݿռ�����Ѿ�������ɾ��,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
           Exit Function
        End If
        
        For i = 1 To rsbakspaces.RecordCount
            If rsbakspaces!��� = mlng�ռ��� Then
                txtKeepSpaceName.Text = rsbakspaces!����
                txtKeepSpaceNO.Text = rsbakspaces!���
            Else
                str�ռ����� = str�ռ����� & "," & rsbakspaces!����
            End If
            If IsNull(rsbakspaces!DB����) = False Then
                MsgBox "Oracle��֧��ͨ��DBLink��������XMLType�ȶ������ͻ��û����������ֶεı����ԣ���֧�ֶ�Զ����ʷ���ݿռ���кϲ���", vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
            End If
            
            rsbakspaces.MoveNext
        Next
        txtMergeSpace.Text = Mid(str�ռ�����, 2)
        
        rsbakspaces.MoveFirst
        Set mrsMergeSpace = rsbakspaces
        
    ElseIf mintFunType = F6ת�� Then
        If rsbakspaces.RecordCount = 0 Then
           MsgBox "��ʷ���ݿռ�����Ѿ�������ɾ�����߽�����һ����ǰ��ʷ��,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
           Exit Function
        End If
        
        With lvwHistory
            .ListItems.Clear
            Do While Not rsbakspaces.EOF
                Set objItem = .ListItems.Add(, "K" & Nvl(rsbakspaces!���), Lpad(Nvl(rsbakspaces!���), lngMaxLen), 0, 0)
                
                If .SelectedItem Is Nothing Then objItem.Selected = True
       
                objItem.SubItems(C1����) = Nvl(rsbakspaces!����)
                objItem.SubItems(C2��ǰ) = IIf(Val(Nvl(rsbakspaces!��ǰ)) = 1, "��", "")
                objItem.SubItems(C3ֻ��) = IIf(Val(Nvl(rsbakspaces!ֻ��)) = 1, "��", "")
                objItem.SubItems(C4������) = Nvl(rsbakspaces!������)
                
                If Val(Nvl(rsbakspaces!ֻ��)) = 1 Then
                    strImgKey = "Lock"
                Else
                    strImgKey = "Other"
                End If
                
                objItem.SmallIcon = strImgKey
                objItem.Icon = strImgKey
                
                err.Clear: On Error Resume Next
                gstrSQL = "select ϵͳ,�汾��,��������,���ת������,��������� from " & rsbakspaces!������ & ".ZLBAKINFO where ϵͳ=" & mlngSys
                Set rsTmp = New ADODB.Recordset
                rsTmp.Open gstrSQL, mcnDBA, adOpenKeyset, adLockReadOnly
                '������ʷ�ռ������Ȩ�޶��ڸ���Ӧ��ϵͳ�������ߵģ�����Ӧ���ܷ���
                If err <> 0 Then
                    MsgBox "����:" & vbCrLf & "  ��ʷ���ݿռ�" & rsbakspaces!���� & "������������,����Ȩ���Ƿ�������" & vbCrLf & err.Description, vbInformation + vbDefaultButton1
                Else
                    If Not rsTmp.EOF Then
                        objItem.SubItems(C5�汾��) = Nvl(rsTmp!�汾��)
                        objItem.SubItems(C6���ת������) = Format(rsTmp!���ת������, "yyyy-mm-dd")
                        objItem.SubItems(C7���������) = Format(rsTmp!���������, "yyyy-mm-dd")
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
            MsgBox "��ʷ���ݿռ���Ϊ:" & mlng�ռ��� & " �Ѿ�������ɾ��,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
            lblCode.Caption = "���:"
            lblSpace.Caption = "����:"
            lblDbLink.Caption = "DB����:"
            lblSpaceOwner.Caption = "������:"
            Exit Function
        End If
        
        mlng�ռ��� = Val(Nvl(rsbakspaces!���))
        lblCode.Caption = "���:" & Nvl(rsbakspaces!���)
        lblSpace.Caption = "����:" & Nvl(rsbakspaces!����)
        lblSpace.Tag = Nvl(rsbakspaces!����)
        lblDbLink.Caption = "DB����:" & rsbakspaces!DB����
        mstrDBLink = "" & rsbakspaces!DB����
    
        lblSpaceOwner.Caption = "������:" & Nvl(rsbakspaces!������)
        lblSpaceOwner.Tag = Nvl(rsbakspaces!������)
        
        If mintFunType = F4�л� Then
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
                    MsgBox "��ʷ�ռ�����ݿ���·" & txtDBLink.Text & "�޷���������,���˹�ɾ�������´�����", vbExclamation, gstrSysName
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
'���ܣ���鴫���Դ��Ŀ�����ݿ��Ƿ��������
    Dim rsFrom As ADODB.Recordset, rsTo As ADODB.Recordset
    Dim strErr As String
    Dim strTemp As String
    
    '1.Ҫ�����ΰ汾��ͬ�������汾����ͬ(10.2.0.4-->10.2.0.1)
    gstrSQL = "Select Substr(Banner, 6, 4) As �汾 From V$version Where Banner Like 'CORE%'"
    Set rsFrom = New ADODB.Recordset
    Set rsTo = New ADODB.Recordset
    OpenRecordset rsFrom, gstrSQL, Me.Caption, , , mcnDBA
    OpenRecordset rsTo, gstrSQL, Me.Caption, , , mcnOracle
    If rsFrom!�汾 <> rsTo!�汾 Then
        strErr = strErr & vbCrLf & "�汾����̫��,Դ��:" & rsFrom!�汾 & ",Ŀ���:" & rsTo!�汾 & "��"
    End If
    
    '2.�����ݰ汾
    gstrSQL = "Select Substr(Value, 1, 4) As ���ݰ汾 From V$parameter Where Name = 'compatible'"
    Set rsFrom = New ADODB.Recordset
    Set rsTo = New ADODB.Recordset
    OpenRecordset rsFrom, gstrSQL, Me.Caption, , , mcnDBA
    OpenRecordset rsTo, gstrSQL, Me.Caption, , , mcnOracle
    If rsFrom!���ݰ汾 <> rsTo!���ݰ汾 Then
        strErr = strErr & vbCrLf & "���ݰ汾����̫��,Դ��:" & rsFrom!���ݰ汾 & ",Ŀ���:" & rsTo!���ݰ汾 & "��"
    End If
    
    '3.����ַ���
    gstrSQL = "SELECT PROPERTY_NAME, PROPERTY_VALUE" & vbNewLine & _
                "FROM DATABASE_PROPERTIES" & vbNewLine & _
                "WHERE PROPERTY_NAME ='NLS_CHARACTERSET' or PROPERTY_NAME ='NLS_NCHAR_CHARACTERSET'"
    Set rsFrom = New ADODB.Recordset
    Set rsTo = New ADODB.Recordset
    OpenRecordset rsFrom, gstrSQL, Me.Caption, , , mcnDBA
    OpenRecordset rsTo, gstrSQL, Me.Caption, , , mcnOracle
    
    rsFrom.Filter = "PROPERTY_NAME='NLS_CHARACTERSET'"
    If rsFrom!PROPERTY_VALUE <> rsTo!PROPERTY_VALUE Then
        If MsgBox("���ݿ��ַ�����ͬ,���ܵ��´���ʧ�ܡ�" & vbCrLf & "Դ��:" & rsFrom!PROPERTY_VALUE & ",Ŀ���:" & rsTo!PROPERTY_VALUE & "��" & vbCrLf & "��ȷ��Ҫ������", vbQuestion + vbOKCancel + vbDefaultButton1) = vbCancel Then
            Exit Function
        End If
    End If
    '������ϵͳ��NVARCHAR��NCHAR�������ͣ��ݲ��������ַ���
'    rsFrom.Filter = "PROPERTY_NAME='NLS_NCHAR_CHARACTERSET'"
'    If rsFrom!PROPERTY_VALUE <> rsTo!PROPERTY_VALUE Then
'        strErr = strErr & vbCrLf & "�����ַ�����ͬ,Դ��:" & rsFrom!PROPERTY_VALUE & ",Ŀ���:" & rsTo!PROPERTY_VALUE & "��"
'    End If
    
    '4.���֧��ת����ƽ̨
    'Ŀ���ƽ̨��Ϣ
    gstrSQL = "Select d.Platform_Name, Endian_Format" & vbNewLine & _
                "From V$transportable_Platform Tp, V$database D" & vbNewLine & _
                "Where Tp.Platform_Name = d.Platform_Name"
    Set rsFrom = New ADODB.Recordset
    Set rsTo = New ADODB.Recordset
    OpenRecordset rsTo, gstrSQL, Me.Caption, , , mcnOracle
    
    '��Դ��ƽ̨�Ƿ�֧��ת��
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
        If MsgBox("Դ�ⲻ֧��ת�������ļ���Ŀ���ƽ̨��" & rsTo!Platform_Name & "," & rsTo!Endian_Format & "����" & vbCrLf & "���ܵ��´���ʧ�ܣ���ȷ��Ҫ������", vbQuestion + vbOKCancel + vbDefaultButton1) = vbCancel Then
            Exit Function
        End If
    End If
    
    If strErr <> "" Then
        MsgBox "��鷢������ԭ�����޷����д��䣺" & strErr, vbExclamation, gstrSysName
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
    
    If mintFunType = F0���� Then
        strKey = "δ�����ʷ���ݿռ䴴��,���ȡ����"
    ElseIf mintFunType = F1��ж Then
        strKey = "δ�����ʷ���ݿռ��ж,���ȡ����"
    ElseIf mintFunType = F2��ֲ Then
        strKey = "δ�����ʷ���ݿռ����ֲ,���ȡ����"
    End If
    
    If mblnMustInstall And mintFunType = F0���� Then
        MsgBox "��ǰϵͳ���谲װ��ʷ���ݿռ�󣬲�������" & vbCrLf & "ʹ�ø�ϵͳ,��˲���ȡ������!", vbInformation + vbDefaultButton1, gstrSysName
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
    '����:������ʷ���ݿռ�
    '����:strDbaName-Զ�̵�dba�û���
    '     strDbaPass-Զ�̵�dba�û���������
    '     strServer-Զ�̷�����
    '     strBakUserName-��ʷ�ռ���
    '     strBakUserPwd-�û�����(δ���ܵ�)
    '     strDb_Link-������
    '     strtablespace-��ռ���
    '     strDataFile-�����ļ�
    '     ExtentSize:ͳһ���ߴ�
    '     blnHaveUser:�Ƿ��Ѵ����û�
    '     strTbsNameIdx,strTbsNameLob:������ռ�ʹ�����ռ�����
    '     lngFileAmount,lngFileIdxAmount,lngFileLobAmount:�����ļ��������ļ���������ļ�������
    '����;�ɹ�����true,���򷵻�false
    '--------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intCreate As Integer
    Dim strDba_Link As String
    Dim strFileHead As String, strFileTail As String, strError As String
    
    
    gstrSQL = "Insert Into zltools.zlbakspaces(ϵͳ, ���, ����, ������, db����, ��ǰ, ֻ��)��Values("
    gstrSQL = gstrSQL & "" & mlngSys & ","
    gstrSQL = gstrSQL & "" & Val(txt���.Text) & ","
    gstrSQL = gstrSQL & "'" & strTableSpace & "',"
    gstrSQL = gstrSQL & "'" & strBakUserName & "',"
    gstrSQL = gstrSQL & IIf(strDbLink = "", "NULL", "'" & strDbLink & "'") & ","
    gstrSQL = gstrSQL & "0,0)"
    
    On Error Resume Next
    mcnOracle.Execute gstrSQL
    If err <> 0 Then
        MsgBox "��ʷ�ռ��������Ѿ�����,����!", vbInformation + vbDefaultButton1, mstrSysName
        Exit Function
    End If
    
    On Error GoTo errHand
    SetPromptText "���ڴ�����ռ䡭"
    If blnHaveUser = False Then
        '��һ��:������ʷ���ݿռ�ı�ռ�
        '1-�����ɹ���2-��ռ��Ѿ����ڣ�3-����ʧ��
        
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
        
        '�ڶ���:������ʷ���ݿռ��û�
        SetPromptText "����������ʷ�ռ��û�" & strBakUserName
        
        gstrSQL = "alter user " & strBakUserName & " DEFAULT TABLESPACE " & strTableSpace
        mcnDBA.Execute gstrSQL
        
        gstrSQL = "Grant Connect,Resource,UNLIMITED TABLESPACE," & _
                " Create Table,Create Sequence,Create Role,Create User,Drop User,Create Public Synonym,Drop Public Synonym," & _
                " Alter Session,Create Session,Create Synonym,Create View,Create Database Link,Create Cluster" & _
                " to " & strBakUserName & " With Admin Option"
        mcnDBA.Execute gstrSQL
    End If
    
    '������:������ص�ת�����ݽṹ
    SetPromptText "�������ݽṹ" & strBakUserName
    
    If CreateHistoryStru(strDbLink, strBakUserName, strTableSpace, mstrOwnerName, strTbsNameIdx, strTbsNameLob) = False Then
        'ɾ����������ʱ����
        GoTo ErrDropLink
        Exit Function
    End If
    
    '���Ĳ�:��Ȩ(Զ�����ݿ�ʹ�õ�DBA�û����ӣ�����������Ȩ)
    If strDbLink = "" Then
        Dim cnnbak As ADODB.Connection
        
        Set cnnbak = gobjRegister.GetConnection(strServer, strBakUserName, strBakUserPwd, False, MSODBC, strError, False)
        If cnnbak.State = adStateClosed Then
             'ɾ����������ʱ����
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
   If MsgBox("��װʧ��,����!" & vbCrLf & "�����:" & err.Number & vbCrLf & "��������:" & err.Description & vbCrLf & gstrSQL, vbRetryCancel + vbDefaultButton2 + vbQuestion) = vbRetry Then Resume
End Function

Private Function CreateDbLink(ByVal cnOracle As ADODB.Connection, ByVal strDbLinkName As String, _
            strUserName As String, strPassword As String, strServer As String, _
            strOwner As String, Optional blnDropLink As Boolean = True, Optional blnCheckLink As Boolean = True) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '����:����Զ������
    '����:cnOracle-oracle���Ӷ���
    '     strDbLinkName-Զ��������
    '     strUserName-Զ���û���
    '     strPassWord-Զ���û�������
    '     strSerVer-Զ�����ӷ���
    '     strOwner-�������ӵ�������
    '     blnDropLink-��������ǰ�Ƿ�ѡɾ��ԭ����
    '     blnCheckLink-��������Ƿ�����
    '����:���ӳɹ�,����true,���򷵻�False
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
        MsgBox "����Զ������ʱ����,������Ϣ����:" & vbCrLf & "(" & err.Number & ") " & err.Description & vbCrLf & gstrSQL, vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    ElseIf cnOracle.Errors.Count > 0 Then
        MsgBox "����Զ������ʱ����,������Ϣ����:" & vbCrLf & cnOracle.Errors(0).Description & vbCrLf & gstrSQL, vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    If blnCheckLink = True Then
        '��鴴���������Ƿ���Ч
        On Error Resume Next
        cnOracle.Errors.Clear
        gstrSQL = "Select 1 from dual@" & strDbLinkName
        OpenRecordset rsTemp, gstrSQL, Me.Caption, , , cnOracle
        
        If err.Number <> 0 Then
            If InStr(1, err.Description, "ORA-02085") > 0 Then
                If CheckGlobal_name(cnOracle) = True Then
                    MsgBox "������Զ��������һ��Ҫ��Ŀ�����ݿ��ȫ������" & vbCrLf & "һ��(����Global_name),����ʧ��!", vbInformation + vbDefaultButton1, gstrSysName
                Else
                    '������������
                    MsgBox "������Զ�����Ӳ�������ʹ��,����ʧ��!" & vbCrLf & "(" & err.Number & ") " & err.Description, vbInformation + vbDefaultButton1, gstrSysName
                End If
            Else
                '������������
                MsgBox "������Զ�����Ӳ�������ʹ��,����ʧ��!" & vbCrLf & "(" & err.Number & ") " & err.Description, vbInformation + vbDefaultButton1, gstrSysName
            End If
            
            On Error Resume Next
            gstrSQL = "drop Database Link " & strDbLinkName
            cnOracle.Execute gstrSQL
            Exit Function
        ElseIf cnOracle.Errors.Count > 0 Then
            If InStr(1, cnOracle.Errors(0).Description, "ORA-02085") > 0 Then
                If CheckGlobal_name(cnOracle) = True Then
                    MsgBox "������Զ��������һ��Ҫ��Ŀ�����ݿ��ȫ������" & vbCrLf & "һ��(����Global_name),����ʧ��!", vbInformation + vbDefaultButton1, gstrSysName
                Else
                    '������������
                    MsgBox "������Զ�����Ӳ�������ʹ��,����ʧ��!" & vbCrLf & cnOracle.Errors(0).Description, vbInformation + vbDefaultButton1, gstrSysName
                End If
            Else
                '������������
                MsgBox "������Զ�����Ӳ�������ʹ��,����ʧ��!" & vbCrLf & cnOracle.Errors(0).Description, vbInformation + vbDefaultButton1, gstrSysName
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
    MsgBox "����DBLinkʧ�ܣ�" & err.Description, vbExclamation, gstrSysName
    
End Function

Private Function CheckGlobal_name(ByVal cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------
    '����:���ȫ�ֲ����Ƿ�Ϊtrue
    '����:cnOracle-�������ݿ�
    '����:����Ϊtrue,����true,����false
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
    '����:������ʷ���ݿռ�����
    '����:cnBakOracle-��ʷ���ݿռ�����
    '     lngSys-ϵ�y��
    '     strVer-�汾��
    '����:���³ɹ�,����true,���򷵻�False
    '--------------------------------------------------------------------------------------------
    If strVer = "" Then
        gstrSQL = "Update zlbakinfo set ���������=sysdate where ϵͳ=" & lngSys
    Else
        gstrSQL = "Update zlbakinfo set �汾��='" & strVer & "',��������=sysdate where ϵͳ=" & lngSys
    End If
    
    err = 0: On Error GoTo errHand:
    cnBakOracle.Execute gstrSQL
    UpdateBakInfor = True
    Exit Function
errHand:
    MsgBox "������ʷ���ݿռ�汾��ʱ����,������Ϣ����:" & vbCrLf & err.Description
End Function

Private Function UpdateZlBakSpace(ByRef cnOracle As ADODB.Connection, ByVal lngBakCode As Long, ByVal lngSys As Long, _
    Optional blnDelete As Boolean = False, Optional ByVal blnReadonly As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '����:���µ�ǰ��־��ɾ�������ʷ���ݿռ���Ϣ
    '����:cnOracle-����zlBakSpaces���е�����
    '     lngBakCode-���
    '     lngSys-ϵͳ
    '     blnDelete-�Ƿ�ɾ��
    '����:�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------
    err = 0: On Error GoTo errHand:
    
    If blnDelete Then
        gstrSQL = "delete  zltools.zlbakspaces where ϵͳ=" & lngSys & " and ���=" & lngBakCode
        cnOracle.Execute gstrSQL
    Else
        gstrSQL = "Update zltools.zlbakspaces set ��ǰ=0 where ϵͳ=" & lngSys
        cnOracle.Execute gstrSQL
        
        gstrSQL = "Update zltools.zlbakspaces set ��ǰ=1" & IIf(blnReadonly, ",ֻ��=1", "") & " where ϵͳ=" & lngSys & " and ���=" & lngBakCode
        cnOracle.Execute gstrSQL
    End If
    UpdateZlBakSpace = True
    Exit Function
errHand:
    If MsgBox(IIf(blnDelete, "ɾ��", "����") & " ��ʷ���ݿռ����,��ϸ������Ϣ����:" & vbCrLf & " (" & err.Number & ") " & err.Description, vbRetryCancel + vbDefaultButton1 + vbQuestion, gstrSysName) = vbRetry Then Resume
End Function

Private Function ExeFuncImport(ByRef cnOracle As ADODB.Connection, _
        ByVal lngBakCode As Long, ByVal strBakName As String, _
        ByVal strBakOwner As String, ByVal lngSys As Long, _
        Optional ByVal strDbLink As String, Optional ByVal strDBLinkUser As String, Optional ByVal strDBLinkPwd As String, Optional ByVal strDBLinkServer As String) As Boolean
    '---------------------------------------------------------------------------------------------------
    '����:ֲ���Ѿ������˵���ʷ���ݿռ�
    '����:cnBakOracle-��ʷ���ݿռ�����
    '     cnOracle-������������
    '     lngBakCode-��ʷ���ݿռ���
    '     strBakName-��ʷ���ݿռ�����
    '     strBakOwner-��ʷ���ݿռ�������
    '     lngSys-ϵͳ��
    '     strDbLink-DBLink����
    '����:�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------
    '���ȼ�����ݵ���Ч��
    Dim rsTemp As New ADODB.Recordset
        
    'zlbakinfo(ϵͳ,�汾��,��������)
    gstrSQL = "Insert Into zltools.zlbakspaces(ϵͳ, ���, ����, ������, db����, ��ǰ, ֻ��)��Values("
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
        MsgBox "��ʷ�ռ��������Ѿ�����,ֲ��ʧ��,����!", vbInformation + vbDefaultButton1, mstrSysName
        Exit Function
    End If
    
    
    ExeFuncImport = True
End Function

Public Function CreateHistoryStru(ByVal strDb_Link As String, ByVal strBakUserName As String, ByVal strBakTableSpace As String, _
    ByVal strSourceUserName As String, ByVal strTbsNameIdx As String, ByVal strTbsNameLob As String) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����:������ʷ��ռ�Ŀռ����ݽṹ
    '����:strDb_link-Զ��������
    '     strBakUserName-�����û���
    '     strBakTablespace-���ݱ�ռ�,strTbsNameIdx-������ռ�,strTbsNameLob-������ռ�
    '     strSourceUserName-�������ݽṹ��Դ�û�
    '����:�ɹ�����true,���򷵻�false
    '--------------------------------------------------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strSysIn As String, blnFeeView As Boolean
    
    '��Ҫ����zlBakInfo��
    On Error GoTo errHand
    If CheckTable("zlBakInfo", mcnDBA, strBakUserName) = False Then
        gstrSQL = "Create Table " & strBakUserName & ".zlBakInfo(ϵͳ number(5),�汾�� varchar2(20),�������� date,���ת������ date,��������� date,��ֹ��� varchar2(500),��ǰִ�� number(1),��ǰ��ֹ��� VarChar2(500)) Tablespace " & strBakTableSpace
        mcnDBA.Execute gstrSQL
        gstrSQL = "Alter Table " & strBakUserName & ".zlBakInfo Add Constraint zlBakInfo_PK Primary Key (ϵͳ) USING INDEX PCTFREE 5"
        mcnDBA.Execute gstrSQL
    End If
    
    '--������ذ汾��Ϣ
    gstrSQL = "select ���,�汾��,sysdate from  zltools.zlsystems  where ���=" & mlngSys
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , mcnOracle
    If rsTemp.EOF = False Then
        gstrSQL = "insert into " & strBakUserName & ".zlbakinfo(ϵͳ,�汾��,��������)  values(" & mlngSys & ",'" & Nvl(rsTemp!�汾��) & "',sysdate) "
        mcnDBA.Execute gstrSQL
    End If

    gstrSQL = "Select ���� From zlbakTables a where a.ϵͳ=" & mlngSys
    
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , mcnOracle
    
    Call SetProgressVisible(True)
    pgbState.Max = rsTemp.RecordCount + 1
    pgbState.value = 0
    SetPromptText "���ƽṹ"
     
    With rsTemp
        Do While Not .EOF
            If "" & !���� = "���˷��ü�¼" Then  '���⴦��
                blnFeeView = Not CheckTable(Nvl(!����), mcnDBA, strBakUserName, 1)
            Else
                '�����Ƿ����,���ڽ�������
                If CheckTable(Nvl(!����), mcnDBA, strBakUserName) = False Then
                    '������ṹ
                    If CreateTable(mcnOracle, strSourceUserName, strBakTableSpace, strBakUserName, !����, strTbsNameLob, mcnDBA) = "" Then Call SetProgressVisible(False): Exit Function
            
                    '������ṹ��ص�PK��UQ
                    If CreateConstraint(rsTemp!����, strTbsNameIdx, strSourceUserName, strBakUserName) = False Then Call SetProgressVisible(False): Exit Function
                    '������ṹ����IX
                    If CreateIndex(rsTemp!����, strTbsNameIdx, strSourceUserName, strBakUserName) = False Then Call SetProgressVisible(False): Exit Function
                End If
            End If
             
            pgbState.value = pgbState.value + 1
            .MoveNext
        Loop
        If blnFeeView Then
            If mblnSysUpdateCall Then On Error Resume Next
            '��ʷ��ռ��е���ͼ��������̶����䣬��������ü�¼��סԺ���ü�¼�����ֶ�ʱ������ͼ���䣬�������ڼ��ݾɵĳ����ѯ��
            strSQL = "CREATE OR REPLACE VIEW " & strBakUserName & ".���˷��ü�¼ AS" & vbNewLine & _
                    "SELECT ID,��¼����,NO,ʵ��Ʊ��,��¼״̬,���,��������,�۸񸸺�,�ಡ�˵�,���ʵ�ID,����ID,��ҳID,ҽ�����,�����־,���ʷ���," & vbNewLine & _
                    "  ����,�Ա�,����,��ʶ��,����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,����,��ҩ����,����,�Ӱ��־,���ӱ�־,Ӥ����," & vbNewLine & _
                    "  ������ĿID,�վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,������,��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,ִ����,ִ��״̬,ִ��ʱ��,����," & vbNewLine & _
                    "  ����Ա���,����Ա����,����ID,���ʽ��,���մ���ID,������Ŀ��,���ձ���,��������,ͳ����,�Ƿ��ϴ�,ժҪ,�Ƿ���" & vbNewLine & _
                    "FROM " & strBakUserName & ".סԺ���ü�¼" & vbNewLine & _
                    "UNION ALL" & vbNewLine & _
                    "SELECT ID,��¼����,NO,ʵ��Ʊ��,��¼״̬,���,��������,�۸񸸺�,-Null,���ʵ�ID,����ID,-Null,ҽ�����,�����־,���ʷ���," & vbNewLine & _
                    "  ����,�Ա�,����,��ʶ��,���ʽ,-Null,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,����,��ҩ����,����,�Ӱ��־,���ӱ�־,Ӥ����," & vbNewLine & _
                    "  ������ĿID,�վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,������,��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,ִ����,ִ��״̬,ִ��ʱ��,����," & vbNewLine & _
                    "  ����Ա���,����Ա����,����ID,���ʽ��,���մ���ID,������Ŀ��,���ձ���,��������,ͳ����,�Ƿ��ϴ�,ժҪ,�Ƿ���" & vbNewLine & _
                    "FROM " & strBakUserName & ".������ü�¼"
            mcnDBA.Execute strSQL
        End If
    End With
    Call SetProgressVisible(False)
    CreateHistoryStru = True
    Exit Function
errHand:
    If MsgBox("�����ռ�ṹ����,��ϸ�Ĵ�����Ϣ����:" & vbCrLf & "(" & err.Number & ")" & err.Description & vbCrLf & strSQL, vbQuestion + vbRetryCancel + vbDefaultButton2) = vbRetry Then
        Resume
    End If
End Function

Private Function CheckTable(ByVal strTable As String, ByRef cnOracle As ADODB.Connection, ByVal strOwner As String, Optional ByVal bytType = 0) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------------
    '����:�������ͼ�Ƿ����
    '����:strTable-����
    '     cnoracle-���ݿ�������
    '     strOwNer-������
    '     bytType=0:����1-�����ͼ
    '����:���ڸö����򷵻�true,����False
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
    '����:������ر��Լ��
    '����:strTable-����
    '     strSourceUserName-�������ݽṹ��Դ�û�
    '     strBakUserName-�����û���
    '     strBakTableSpace-���ݿռ�
    '����:�ɹ�����true,����false
    '-------------------------------------------------------------------------
    Dim intType As VbMsgBoxResult
    
    Dim rsUserIndex As New ADODB.Recordset
    Dim rsColumn As New ADODB.Recordset
    Dim strTemp As String
    
    gstrSQL = "Select Table_Name,index_name, Column_Name " & _
             "   From Sys.All_Ind_Columns " & _
             "   Where Index_Owner = [1]  and table_name=[2] And Index_name not like '%_IX_��ת��'" & _
             "   Order By index_name,Column_Position"
    Set rsColumn = gclsBase.OpenSQLRecord(mcnOracle, gstrSQL, Me.Caption, strSourceUserName, strTable)
     
    gstrSQL = "Select Index_name,table_name,tablespace_name,Pct_free,Temporary " & _
             "   From all_indexes a " & _
             "   where  table_owner = [1] and   table_name=[2] And index_type='NORMAL' And Index_name not like '%_IX_��ת��'" & _
             "          And  Not Exists(Select 1 From All_Constraints b Where a.index_name=b.constraint_name  And Constraint_Type In ('P', 'U') And a.table_owner=b.Owner) " & _
             "   order by index_name"
    Set rsUserIndex = gclsBase.OpenSQLRecord(mcnOracle, gstrSQL, Me.Caption, strSourceUserName, strTable)
    
    On Error GoTo errHand
    With rsUserIndex
        Do While Not .EOF
            rsColumn.Filter = "index_name ='" & !Index_Name & "'"
            If rsColumn.EOF Then
                MsgBox "����:" & !Index_Name & "������!", vbInformation + vbDefaultButton1
                If MsgBox("����:" & !Index_Name & "������,�Ƿ����!", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
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
                    '������ֻ�����ݣ�Ϊ��ߴ洢Ч�ʣ��̶�pctfreeΪ5
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
    intType = MsgBox("��������:" & err.Description & vbCrLf & gstrSQL & vbCrLf & "�Ƿ�����?", vbQuestion + vbAbortRetryIgnore + vbDefaultButton2, gstrSysName)
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
    '����:������ر��Լ��
    '����:strTable-����
    '     strBakTableSpace-��ʷ�ռ�
    '     strSourceUserName-�������ݽṹ��Դ�û�
    '     strBakUserName-��ʷ�ռ��û���
    '����:�ɹ�����true,����false
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
                
                '������ֻ�����ݣ�Ϊ��ߴ洢Ч�ʣ��̶�pctfreeΪ5
                strSQL = strSQL & " USING INDEX PCTFREE 5 TABLESPACE " & strBakTableSpace
                
                '����Լ��
                mcnDBA.Execute strSQL
            End If
            .MoveNext
        Loop
    End With
    CreateConstraint = True
    
    Exit Function
errHand:
    If MsgBox("����Լ��ʧ�ܣ�����:" & vbCrLf & err.Description & vbCrLf & strSQL & vbCrLf & "�Ƿ�����������Լ����", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        CreateConstraint = True
    End If
End Function


Private Function GetMaxHistory(Optional ByRef strMax���� As String = "") As Integer
    '-------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ʷ���ݿռ�������
    '����:str����-��������ʷ���ݿռ��������
    '����:�����
    '-------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select max(���) as ���,max(������) as ������,to_Char(sysdate-3*365,'yyyy') as ��   from zltools.zlBakSpaces where ϵͳ=" & mlngSys
    OpenRecordset rsTemp, gstrSQL, "��ȡ������"
    If rsTemp.EOF Then
        strMax���� = Format(DateAdd("yyyy", -3, Now), "YYYY")
        GetMaxHistory = 1
    Else
        If IsNull(rsTemp!������) Then
            strMax���� = Nvl(rsTemp!��)
        Else
            strMax���� = Replace(Replace(UCase(Nvl(rsTemp!������)), "ZLBAK", ""), "ZLHD", "")
            If IsNumeric(strMax����) Then
                If mintFunType = F0���� Then
                    strMax���� = Val(strMax����) + 1
                End If
            Else
                strMax���� = Nvl(rsTemp!��)
            End If
        End If
        GetMaxHistory = Val(Nvl(rsTemp!���)) + 1
    End If
End Function

Private Function SetHistoryInfor(ByRef cnOracle As ADODB.Connection, ByVal strBakUserName As String, ByVal strDbLink As String) As Boolean
    '-------------------------------------------------------------------------------------
    '����:��ȡ��ʷ��ռ�İ汾��Ϣ
    '-------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    If strDbLink <> "" Then strDbLink = "@" & strDbLink
    
    gstrSQL = "Select 1 from all_Tables" & strDbLink & " where table_name='ZLBAKINFO' and OWNER='" & UCase(strBakUserName) & "'"
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , cnOracle
    If rsTemp.EOF Then
        MsgBox "��ֲ�����ʷ���ݿռ䲻�ǺϷ�����ʷ���ݿռ�(��ǰ���ӵ��û�ģʽ�²�����ZLBAKINFO��),���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
         
        Exit Function
    End If
    
    'ȷ����ذ汾.
    gstrSQL = "Select �汾��,��������,���ת������,��������� From " & strBakUserName & ".ZLBAKINFO" & strDbLink & " where ϵͳ=" & mlngSys
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , cnOracle
    If rsTemp.EOF Then
        MsgBox "��ֲ�����ʷ���ݿռ䲻��" & mstrSysName & "����ʷ���ݿռ�!", vbInformation + vbDefaultButton1, gstrSysName
        txtMoveName.Text = ""
        Exit Function
    End If
    
    lblBakVer.Caption = "����:" & Nvl(rsTemp!�汾��)
    lblBakVer.Tag = Nvl(rsTemp!�汾��)
    
    If lblBakVer.Tag > lblDataVer.Tag Then
        If MsgBox("��ֲ�����ʷ���ݿռ�İ汾�������������ݿ�İ汾,��ȷ��Ҫ������?", vbQuestion + vbOKCancel + vbDefaultButton1) = vbCancel Then

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

Private Function CheckCopyObject(ByRef cnBakOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lng��� As Long, ByVal strBakOwnerName As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------
    '����:��鸴�Ʒ�ת�洢���ݿռ�
    '����:cnOracle-��ʷ��������
    '    lngSys-ϵͳ
    '    lng��� -�ռ���
    '    strBakOwnerName-�ռ��û���
    '����:���ݺϷ�,����true,����False
    '-----------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strErr As String
    
   
    strErr = ""
    '���汾
    gstrSQL = "Select �汾��,��������,���ת������,��������� From " & strBakOwnerName & ".ZLBAKINFO where ϵͳ=" & mlngSys
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , cnBakOracle
    If rsTemp.EOF Then
        MsgBox "ָ������ʷ���ݿռ䲻��" & mstrSysName & "����ʷ���ݿռ䣬���ܸ��Ʒ�ת������!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    If mstrVersion < Nvl(rsTemp!�汾��) Then
        strErr = "��ʷ���ݿռ�İ汾(" & Nvl(rsTemp!�汾��) & "�����߿�汾(" & mstrVersion & ") ��Ҫ��," & vbCrLf & " �Ƿ�������ƣ�"
    ElseIf mstrVersion > Nvl(rsTemp!�汾��) Then
        If MsgBox("��ʷ���ݿռ�İ汾(" & Nvl(rsTemp!�汾��) & "�����߿�汾(" & mstrVersion & ") ҪС," & vbCrLf & " �Ƿ�������ƣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        strErr = "�ٴ����ѣ���ʷ���ݿռ�İ汾(" & Nvl(rsTemp!�汾��) & "�����߿�汾(" & mstrVersion & ") ҪС," & vbCrLf & " �Ƿ�������ƣ�"
    Else
        strErr = ""
    End If
    If strErr <> "" Then
        If MsgBox(strErr, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    CheckCopyObject = True
End Function

Private Sub InitCreateTbs()
'���ܣ���ʼ������ҳ�������
    Dim rsTemp As New ADODB.Recordset, strDataFile As String, strTbsName As String
    
    On Error GoTo errh
    '���ݵ�ǰϵͳ�������ļ�ȷ��ȱʡ�ı�ռ��ļ�·��
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
    Call txtBakSpace_Change     'ִ����һ������ִ����һ��ʱ�������ı�������δ�䣬������Ҫ����һ�θ��¼�
    
    '�������
    If Trim(txt���) = "" Then
        txt���.Text = GetMaxHistory(strTbsName)
        If Trim(txtOwnerUsr.Text) = "" Then
            txtOwnerUsr.Text = strTbsName
        Else
            txtBakSpace.Text = txtHD.Text & txtOwnerUsr.Text
        End If
    End If
    
    tbHistory.Tab = 0
    If txtOwnerUsr.Enabled And txtOwnerUsr.Visible Then txtOwnerUsr.SetFocus
         
    cmdNext.Caption = "���(&O)"

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
        '��һ��(��ж,ɾ���ǵڶ���)��ȷ��Զ����ʷ�ռ��DBA�û��Ƿ����
        '------------------------------------------------------------
        SetControlEnable False
        
        strUserName = txtDba�û�.Text
        strPassword = txtDba����.Text
        strServer = Trim(txtDbaServer.Text)
        
        If CheckUser(strUserName, strPassword, strServer, strErrMsg) = False Then
            MsgBox strErrMsg, vbExclamation, gstrSysName
            Call SetControlEnable(True)
            Exit Sub
        End If
        txtDba�û�.Text = strUserName
        txtDba����.Text = strPassword
        txtDbaServer.Text = strServer
        
        If optServer(1).value = True Then strDbLink = Trim(txtDBLink.Text)
        
        
        If mintFunType = F6ת�� Then
            If strServer = gstrServer Then
                MsgBox "Դ���ݿ��Ŀ�����ݿⲻ����ͬһ����������ָ��������", vbInformation, gstrSysName
                Call SetControlEnable(True)
                If txtDbaServer.Visible And txtDbaServer.Enabled Then txtDbaServer.SetFocus
                Exit Sub
            End If
            
        ElseIf mintFunType = F3���� Then
            If Trim(txtDbaServer.Text) = "" Then
                MsgBox "��ʷ�ռ����ӵķ�����Ϊ�գ����Ʋ���Ҫ�����ָ�������������������롣", vbInformation, gstrSysName
                Call SetControlEnable(True)
                If txtDbaServer.Visible And txtDbaServer.Enabled Then txtDbaServer.SetFocus
                Exit Sub
            End If
        End If
        
        If mintFunType <> F4�л� And mintFunType <> F2��ֲ Then
                        
            '����ʱ�������ǰ�û���DBA����û����ʾ���ݿ��������ý��棬ֱ���õ�ǰ�����û���������
            '��жʱ������Ҫɾ��Զ����ʷ�⣬���ԣ��������������Ϣ���������ӣ���ʹִ���˲��Բ�����
            If mintFunType <> F0���� Or (mintFunType = F0���� And mcnDBA.State = adStateClosed) Then
                Set mcnDBA = gobjRegister.GetConnection(strServer, strUserName, strPassword, False, MSODBC, strError, False)
            End If
            
            If mcnDBA.State = adStateClosed Then
                MsgBox strError, vbInformation, gstrSysName
                Call SetControlEnable(True)
                If txtDba����.Visible And txtDba����.Enabled Then txtDba����.SetFocus
                Exit Sub
            Else
                Call SetSQLTrace(strServer, strUserName, mcnDBA)
                
                If Not (mintFunType = F3����) Then
                    If CheckIsDBA(mcnDBA) = False Then
                        MsgBox "����DBA�û�,���ܼ�����", vbExclamation, gstrSysName
                        Call SetControlEnable(True)
                        If txtDba�û�.Visible And txtDba�û�.Enabled Then txtDba�û�.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        If mintFunType = F0���� Then
            Call InitCreateTbs
                    
            Call SetControlEnable(True)
            
            fraSetup(0).Visible = False
            fraSetup(1).Visible = True
            cmdPrevious.Enabled = True
            
        ElseIf mintFunType = F1��ж Then '��ж(ɾ��ģʽ���ڶ���)
            '��֤��ݲ��������˵��
            If Not CheckAuditStatus("0201", "��ж", strRemarks) Then
                Call SetControlEnable(True)
                Exit Sub
            End If
            If ExeFuncUnInstall(lblSpace.Tag, UCase(lblSpaceOwner.Tag), mlng�ռ���, False) Then
                MsgBox "��ʷ���ݿռ�" & lblSpace.Tag & "��ж(ɾ��)�ɹ���", vbInformation + vbDefaultButton1, gstrSysName
                '������Ҫ������־
                Call SaveAuditLog(3, "��ж", "��ж(ɾ��)��ʷ���ݿռ䡰" & lblSpace.Tag & "��", strRemarks)
                mblnSucced = True
            End If
            Call SetControlEnable(True)
            Unload Me
            
        ElseIf mintFunType = F2��ֲ Then '��ֲ
            
            If strDbLink <> "" Then
                '����Զ������
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
            
            '�������
            If Trim(txtMoveCode) = "" Then
                txtMoveCode.Text = GetMaxHistory(strTbsName)
                If Trim(txtMoveName.Text) = "" Then
                    txtMoveName.Text = strTbsName
                End If
            End If
            
            Call SetControlEnable(True)
            cmdPrevious.Enabled = True
            cmdNext.Caption = "��ֲ(&Z)"
            
            txtMoveName.Text = strUserName  'Ĭ��Ϊ��ʷ���ݿռ��û�
            txtMoveUser.Text = strUserName
                        
            If txtMoveName.Enabled And txtMoveName.Visible Then txtMoveName.SetFocus
            
        ElseIf mintFunType = F3���� Then
            '���汾�����ݶ����Ƿ���ȷ.
            If MsgBox("���Ʒ�ת�����ݿ���Ҫ���ѽϳ���ʱ�䣬������ʷ�ռ��е�ͬ��������ݽ��ᱻ���ǣ���ȷ��Ҫ������", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Call SetControlEnable(True)
                If txtDba�û�.Visible And txtDba�û�.Enabled Then txtDba�û�.SetFocus
                Exit Sub
            End If
                        
            If CheckCopyObject(mcnDBA, mlngSys, mlng�ռ���, strUserName) = False Then
                Call SetControlEnable(True)
                If txtDba�û�.Visible And txtDba�û�.Enabled Then txtDba�û�.SetFocus
                Exit Sub
            End If
            
            strTbsName = GetBakTableSpace(mcnDBA, strUserName)
            If ExeFuncCopy(strUserName, strPassword, strServer, strTbsName) = False Then
                Call SetControlEnable(True)
                If txtDba�û�.Visible And txtDba�û�.Enabled Then txtDba�û�.SetFocus
                Exit Sub
            End If
                        
            Call SetControlEnable(True)
            If Image1.ToolTipText = "������ʱ�ļ������Ƹ��Ƶ����а�" Then
                MsgBox "�ѽ����ƽű��������ļ�" & Clipboard.GetText, vbInformation + vbDefaultButton1, gstrSysName
            Else
                MsgBox "���Ƴɹ�!", vbInformation + vbDefaultButton1, gstrSysName
            End If
            Unload Me
            
        ElseIf mintFunType = F4�л� Then
            
            
            If strDbLink <> "" Then
                If CreateDbLink(mcnOracle, strDbLink, strUserName, strPassword, strServer, mstrOwnerName) = False Then
                    '��������ʧ��
                    Call SetControlEnable(True)
                    Exit Sub
                End If
            Else
                gstrSQL = "Update ZLTOOLS.zlbakspaces set DB����='" & strDbLink & "' where ��� = " & mlng�ռ���
                mcnOracle.Execute gstrSQL
            End If
            
            If ExeFuncChange(strUserName, mstrOwnerName, mlngSys, bytErr, strErrMsg, strDbLink) = False Then
                '1-����ʧЧ,2-ϵͳ������,3-���߰汾������ʷ�汾,4-���߰汾С����ʷ�汾
                Call SetControlEnable(True)
                Select Case bytErr
                Case 1, 2, 4
                    MsgBox strErrMsg, vbInformation + vbDefaultButton1, gstrSysName
                Case 3
                    MsgBox strErrMsg, vbInformation + vbDefaultButton1, gstrSysName
                    cmd����.Visible = True
                    lblSetupIni.Visible = True
                    lblIniModi.Visible = True
                End Select
            Else
                Call SetControlEnable(True)
                mblnSucced = True
                MsgBox "�л��ɹ�!", vbInformation + vbDefaultButton1, gstrSysName
                Unload Me
            End If
        ElseIf mintFunType = F5�ϲ� Then
            
            mblnSucced = ExeFuncMerge
            
            Call SetControlEnable(True)
            Unload Me
            
        ElseIf mintFunType = F6ת�� Then
            Call SetControlEnable(True)
            
            If CheckTransCondition = False Then
                Unload Me
            ElseIf LoadSpaceData = False Then
                Unload Me
            Else
                lblNoteTrans.Caption = "1.����ѡ��[��ǰ]��ʷ���ݿռ䣬�������һ����ʷ���ݿռ䣬������Դ�������ϴ����µ���ʷ�ռ�����Ϊ[��ǰ]��" & vbCrLf & _
                            "2.Ŀ����������ܴ�����Դ������ͬ������ʷ�ռ����Ƽ������ļ���" & vbCrLf & _
                            "3.Դ��������ʷ�ռ��еĶ���Ҫ�����԰����ģ�������������洢��ͬһ��ռ䣬������ڴ洢��������ռ�������������Զ��ؽ���"
                            
                cmdNext.Caption = "ת��(&Z)"
                cmdPrevious.Enabled = True
                fraSetup(0).Visible = False
                fraTrans.Visible = True
            End If
        End If
        
    ElseIf fraSetup(1).Visible Then '����
        '------------------------------------------------------------
        '�ڶ�����ȷ����Ӧ����ʷ���ݿռ�����Ƽ������ļ�
        '------------------------------------------------------------
        SetPromptText "���ڼ�����ݵ���Ч��..."
        
        blnHaveUser = False
        If CheckCreateBakInput = False Then SetPromptText "": Exit Sub
                        
        If MsgBox("��ȷ�����ڿ�ʼ������ʷ���ݱ�ռ���", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
            Exit Sub
        End If
                        
        SetPromptText ""
        strUserName = txtDba�û�.Text
        strPassword = txtDba����.Text
        strServer = txtDbaServer.Text
        If optServer(1).value = True Then strDbLink = Trim(txtDBLink.Text)
        mstrDBLink = strDbLink  '��������ʱ��ɾ�����̻��õ�mstrDBLink
        
        '����Dblink
        If strDbLink <> "" Then
            If CreateDbLink(mcnOracle, strDbLink, strUserName, strPassword, strServer, mstrOwnerName) = False Then SetPromptText "": Exit Sub
        End If
        
        '������ʷ���û��������û����ڣ�����һ�α�����ʼ���ź��棬��Ϊ��ʷ�����ʱ����Ҫ��ȡ�Ѿ����ڵ���ʷ���ռ����Ϣ
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
            MsgBox "��װʧ�ܣ�ϵͳ���Զ�����Ѿ���װ�����ݡ�", vbInformation, gstrSysName
            DoEvents
            Call ExeFuncUnInstall(strBakUserName, UCase(strBakUserName), Val(txt���), True)
            Call SetControlEnable(True)
            SetPromptText ""
        
            Exit Sub
        Else
            
            '��Ϊ��ǰ�����ݿռ�
            If chkCreate��ǰ.value = 1 Then
                SetPromptText "���ڴ�����ͼ..."
                Call SetProgressVisible(True)
                If CreateAppView(mstrOwnerName, strBakUserName, mlngSys, IIf(strDbLink = "", "", "@" & strDbLink), pgbState) = False Then
                    Call SetProgressVisible(False)
                    MsgBox "��Ϊ��ǰ��ʷ���ݿռ�ʧ��,���ȼ�������[��ֲ]����ֲ��!", vbInformation, gstrSysName
                    Call SetControlEnable(True)
                    Call UpdateZlBakSpace(mcnOracle, Val(txt���), mlngSys, True)
                    DoEvents
                    Unload Me
                    Exit Sub
                End If
                Call SetProgressVisible(False)
                
                mcnOracle.BeginTrans
                gstrSQL = "Update zltools.zlbakspaces set ��ǰ=0 where ϵͳ=" & mlngSys & ""
                mcnOracle.Execute gstrSQL
                gstrSQL = "Update zltools.zlbakspaces set ��ǰ=1 where ϵͳ=" & mlngSys & " and ���=" & Val(txt���.Text)
                mcnOracle.Execute gstrSQL
                mcnOracle.CommitTrans
                '������Ч����:
                Call ReCompileObjects(mcnOracle)
            Else
                '������
                If mblnMustInstall = True Then
                    Call ReCompileObjects(mcnOracle)
                    mblnMustInstall = False
                End If
            End If
        End If
        
        MsgBox "�ѳɹ�������ʷ���ݿռ�!", vbInformation, gstrSysName
        
        mblnSucced = True
        Call SetControlEnable(True)
        Unload Me
        
    ElseIf fraDelete.Visible Then   '��ж(��һ��)
        
        lblNote(0).Caption = "ָ��DBA�û�������ɾ����ռ估�����ļ���"
        fraDelete.Visible = False
        
        cmdNext.Caption = "��ж(&O)"
        cmdPrevious.Enabled = True
        
        '����ģʽ
        If optDele(1).value Then
            '��֤��ݲ��������˵��
            If Not CheckAuditStatus("0201", "��ж", strRemarks) Then Exit Sub
            gstrSQL = "delete ZLTOOLS.zlbakSpaces where  nvl(��ǰ,0)<>1 and ϵͳ=" & mlngSys & " and ���=" & mlng�ռ���
            mcnOracle.Execute gstrSQL
            
            Me.Hide
            MsgBox "��ʷ���ݿռ��ж(����)�ɹ���" & vbCrLf & "�����ͨ������ֲ�����ܽ��������ʷ�ռ�����ֲ��!", vbInformation + vbDefaultButton1, gstrSysName
            '������Ҫ������־
            Call SaveAuditLog(3, "��ж", "��ж(����)��ʷ���ݿռ䡰" & lblSpace.Tag & "��", strRemarks)
            mblnSucced = True
            Unload Me
        Else
            fraSetup(0).Visible = True
            
            If txtDba�û�.Text = "" And txtDba�û�.Enabled Then txtDba�û�.Text = "SYS"
            If txtDba����.Enabled And txtDba����.Visible Then txtDba����.SetFocus
            
            If mstrDBLink <> "" Then
                optServer(1).value = True
                optServer(0).Enabled = False
                
                txtDBLink.Text = mstrDBLink '��ֹ�޸�
                txtDBLink.Enabled = False
            Else
                optServer(0).value = True
                optServer(1).Enabled = False
            End If
        End If
        
    ElseIf fraImport.Visible Then   '��ֲ
        '0-�������������Ƿ�Ϸ�
        SetPromptText "���ڼ�����ݵ���Ч��..."
        
        If optServer(1).value = True Then strDbLink = Trim(txtDBLink.Text)
        
        If CheckMoveInPutValid(strDbLink) = False Then SetPromptText "": Exit Sub
        SetPromptText ""

        '1.ȷ���Ƿ���ڽṹ����
        Call SetControlEnable(False)
        
        
        '��ʷ�ռ䲻���ڽṹ����Ǩ�򱾴�ֲֻ��
        If ExeFuncImport(mcnOracle, Val(txtMoveCode.Text), Trim(txtMoveName), Trim(txtMoveUser), mlngSys, _
            strDbLink, strUserName, strPassword, Trim(txtDbaServer.Text)) = False Then
            'ɾ�������Ϣ:
            Call UpdateZlBakSpace(mcnOracle, Val(txtMoveCode.Text), mlngSys, True)
            Call SetControlEnable(True)
            Exit Sub
        End If
        mblnSucced = True
        If chk��ǰ.value = 1 Then
            '��Ҫ������Ӧ����ͼ
            '-----------------------------------------------------
            SetPromptText ("���ڴ�����ͼ")
            Call SetProgressVisible(True)
            If CreateAppView(mstrOwnerName, Trim(txtMoveUser), mlngSys, IIf(strDbLink = "", "", "@" & strDbLink), pgbState) = False Then
                Call SetProgressVisible(False)
                MsgBox "ֲ�뵱ǰϵͳʱʧ��,���ڹ����������!", vbInformation + vbDefaultButton1, gstrSysName
                Call SetControlEnable(True)
                Unload Me
                Exit Sub
            End If
            Call SetProgressVisible(False)
            '������Ч����:
            Call ReCompileObjects(mcnOracle)
            MsgBox "ֲ��ɹ�!", vbInformation + vbDefaultButton1, gstrSysName
            
            Call UpdateZlBakSpace(mcnOracle, Val(txtMoveCode.Text), mlngSys, False, strDbLink <> "")
        End If
        
        Call SetControlEnable(True)
        Unload Me
    ElseIf fraTrans.Visible Then    '����
        If lvwHistory.SelectedItem Is Nothing Then
             MsgBox "��ѡ��Ҫ����ı�ռ�!", vbInformation + vbDefaultButton1, gstrSysName
             Exit Sub
        Else
            If txtBakPWD.Text = "" Then
                MsgBox "������Ŀ����½�����ʷ�ռ��û����롣", vbInformation + vbDefaultButton1, gstrSysName
                txtBakPWD.SetFocus
                Exit Sub
            End If
            
            If lvwHistory.SelectedItem.SubItems(C2��ǰ) = "��" Then
                MsgBox "��ѡ��ǵ�ǰ��ʷ�ռ䡣��Ϊ������ɾ���ÿռ䣬���Բ���ѡ��ǰ��ʷ�ռ䡣", vbInformation + vbDefaultButton1, gstrSysName
                lvwHistory.SetFocus
                Exit Sub
            End If
            
            strTbsName = lvwHistory.SelectedItem.SubItems(C1����)
            strBakUserName = lvwHistory.SelectedItem.SubItems(C4������)
            
            '1.����û���
            gstrSQL = "Select Decode(Trunc(Created), Trunc(Sysdate), 1, 0) Todaycreate From Dba_Users Where Username = '" & strBakUserName & "'"
            Set rsTemp = New ADODB.Recordset
            Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                '����ǵ���մ����ģ������ϴδ���ʧ�ܣ������ṩɾ��ѡ��
                If rsTemp!Todaycreate = 1 Then
                    If MsgBox("ѡ�����ʷ�ռ��û�" & strBakUserName & "��Ŀ�����ݿ����Ѵ��ڣ���ȷ��Ҫɾ�������´�����(���û��µ����ж��󽫻ᱻһ��ɾ��)", vbOKCancel + vbQuestion + vbDefaultButton1, gstrSysName) = vbOK Then
                        
                        SetPromptText "����ɾ���û�" & strBakUserName & "������"
                        DoEvents
                        gstrSQL = "drop user " & strBakUserName & " cascade"
                        mcnOracle.Execute gstrSQL
                    Else
                        lvwHistory.SelectedItem.ForeColor = &HC0C0C0
                        Exit Sub
                    End If
                Else
                    MsgBox "ѡ�����ʷ�ռ��û�" & strBakUserName & "��Ŀ�����ݿ����Ѵ��ڣ�����֮ǰ����ɾ��ͬ���û�!", vbInformation + vbDefaultButton1, gstrSysName
                    lvwHistory.SelectedItem.ForeColor = &HC0C0C0
                    Exit Sub
                End If
            End If
            
            '2.����ռ���
            '����ǵ���մ����ģ������ϴδ���ʧ�ܣ������ṩɾ��ѡ��
            gstrSQL = "Select Decode(Trunc(Creation_Time), Trunc(Sysdate), 1, 0) Todaycreate" & vbNewLine & _
                    "From Dba_Data_Files A, V$datafile B" & vbNewLine & _
                    "Where a.File_Id = b.File# And Tablespace_Name = '" & strTbsName & "' Order by Creation_Time"
            Set rsTemp = New ADODB.Recordset
            Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                If rsTemp!Todaycreate = 1 Then
                    If MsgBox("ѡ�����ʷ�ռ�" & strTbsName & "��Ŀ�����ݿ����Ѵ��ڣ���ȷ��Ҫɾ�������´�����(�ñ�ռ��µ����ж��󽫻ᱻһ��ɾ��)", vbOKCancel + vbQuestion + vbDefaultButton1, gstrSysName) = vbOK Then
                        SetPromptText "����ɾ����ռ�" & strTbsName & "������"
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
                    MsgBox "ѡ�����ʷ�ռ�" & strTbsName & "����Ŀ�����ݿ��д��ڣ�����֮ǰ����ɾ��ͬ����ռ�!", vbInformation + vbDefaultButton1, gstrSysName
                    lvwHistory.SelectedItem.ForeColor = &HC0C0C0
                    Exit Sub
                End If
                
            End If
            
            '3.����ռ��ļ���
            'ASM�ϵ��ļ�������zlbak2.263.832000313�����ģ����Բ���
            gstrSQL = "Select 1 From dba_data_files Where file_Name like '%/" & strTbsName & ".DBF' or file_Name like '%\" & strTbsName & ".DBF'"
            Set rsTemp = New ADODB.Recordset
            Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                MsgBox "ѡ�����ʷ�ռ��ļ�" & strTbsName & ".DBF����Ŀ�����ݿ��д��ڣ�����֮ǰ����ɾ��ͬ�������ļ�!", vbInformation + vbDefaultButton1, gstrSysName
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
    Case F1��ж
        fraDelete.Visible = True
        fraSetup(0).Visible = False
        fraSetup(1).Visible = False
        Call optDele_Click(0)
    Case F0����
        If fraSetup(1).Visible Then
            fraSetup(1).Visible = False
            fraSetup(0).Visible = True
            cmdPrevious.Enabled = False
            cmdNext.Caption = "��һ��(&N)"
            If txtDba����.Enabled Then txtDba����.SetFocus
        End If
        
    Case F2��ֲ, F5�ϲ�, F6ת��
        If mintFunType = F2��ֲ Then
            fraImport.Visible = False
        ElseIf mintFunType = F5�ϲ� Then
            fraMerge.Visible = False
        ElseIf mintFunType = F6ת�� Then
            fraTrans.Visible = False
        End If
                
        fraSetup(0).Visible = True
        cmdPrevious.Enabled = False
        cmdNext.Caption = "��һ��(&N)"
        cmdNext.Enabled = True
        If txtDba�û�.Enabled And txtDba�û�.Visible Then txtDba�û�.SetFocus
    End Select
End Sub

Private Sub cmd����_Click()
    Dim strUserName As String, strPassword As String, strServer As String, strError As String
           
    strUserName = txtDba�û�.Text
    strPassword = txtDba����.Text
    strServer = txtDbaServer.Text
    
    If CheckUser(strUserName, strPassword, strServer, strError) = False Then
        MsgBox strError, vbExclamation, gstrSysName
        Exit Sub
    End If
    txtDba�û�.Text = strUserName
    txtDba����.Text = strPassword
    txtDbaServer.Text = strServer
    
    
    '�������ּ���ADDRESS_LIST��д������ODBC�£�ֻ֧��SID����֧��SERVICE_NAME;OLEDB�����ֶ�֧��
    'strServer = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strIP & ")(PORT=" & strPort & ")))(CONNECT_DATA=(SID=" & strSID & ")))"
    Set mcnDBA = gobjRegister.GetConnection(strServer, strUserName, strPassword, False, MSODBC, strError, False)
   
    If mcnDBA.State = adStateClosed Then
        MsgBox "�����ݿ����ӳ���" & strError, vbExclamation, gstrSysName
        If txtDba����.Visible And txtDba����.Enabled Then txtDba����.SetFocus
        Exit Sub
    ElseIf mintFunType <> F3���� And mintFunType <> F2��ֲ Then
        
        If CheckIsDBA(mcnDBA) = False Then
            MsgBox "����DBA�û�,���ܼ�����", vbExclamation, gstrSysName
            If txtDba�û�.Visible And txtDba�û�.Enabled Then txtDba�û�.SetFocus
            Exit Sub
        End If
    End If
    MsgBox "���Գɹ�!", vbInformation + vbDefaultButton1, gstrSysName
End Sub

Private Function CheckIsDBA(ByRef connThis As ADODB.Connection) As Boolean
'���ܣ��жϵ�ǰ�û��Ƿ�ΪDBA��ɫ
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errh
    gstrSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTemp = gclsBase.OpenSQLRecord(connThis, gstrSQL, "�жϵ�ǰ�����û��Ƿ����DBA��ɫ")
    CheckIsDBA = rsTemp.RecordCount > 0
    
    Exit Function
errh:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Function ExeFuncChange(strBakUserName As String, _
    strOwner As String, lngSys As Long, bytErr As Byte, strErr As String, ByVal strDbLink As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------
    '����:��ǰ�ռ��л�
    '����:
    '     strBakUserName-��ʷ���ݿռ��û�
    '     strOwner-������������
    '����:bytErr:1-����ʧЧ,2-ϵͳ������,3-���߰汾������ʷ�汾,4-���߰汾С����ʷ�汾
    '     strErr-��������
    
    '����:���óɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    
    'ȷ����ǰ�������Ƿ�湦
    On Error Resume Next
    If strDbLink = "" Then
        gstrSQL = "Select �汾�� from " & strBakUserName & ".zlbakinfo  where ϵͳ=" & lngSys
    Else
        gstrSQL = "Select �汾�� from " & strBakUserName & ".zlbakinfo@" & strDbLink & "  where ϵͳ=" & lngSys
    End If
    
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , mcnOracle
    bytErr = 0
    If err <> 0 Then
        strErr = "������ذ汾��Ϣ����" & "��ϸ�Ĵ�����Ϣ:" & vbCrLf & "(" & err.Number & ")" & err.Description
        bytErr = 1
        Exit Function
    End If
    err.Clear: err = 0
    
    '�����صİ汾�Ƿ���ȷ
    If rsTemp.EOF Then
        strErr = "��ʷ���ݿռ��ϵͳ(" & mstrSysName & ") ������,����!"
        bytErr = 2
        Exit Function
    End If
    If Nvl(rsTemp!�汾��) < mstrVersion Then
        '��ǰ�汾��С�����߰汾��,������
        strErr = "��ʷ���ݿռ��ϵͳ�汾(" & Nvl(rsTemp!�汾��) & ") С�������߰汾(" & mstrVersion & ")," & vbCrLf & " ��������ʷ����!"
        bytErr = 3
        Exit Function
    ElseIf Nvl(rsTemp!�汾��) > mstrVersion Then
        '�������߰汾��
        strErr = "��ʷ���ݿռ��ϵͳ�汾(" & Nvl(rsTemp!�汾��) & ") ���������߰汾(" & mstrVersion & ")," & vbCrLf & " ���������߰汾�����л�,����!"
        bytErr = 4
        Exit Function
    End If
        
    '���������л���
    SetPromptText ("���ڴ�����ͼ")
    Call SetProgressVisible(True)
    If CreateAppView(mstrOwnerName, strBakUserName, mlngSys, IIf(strDbLink = "", "", "@" & strDbLink), pgbState) = False Then
        Call SetProgressVisible(False)
        'ʧ��
        Exit Function
    End If
    Call SetProgressVisible(False)
    
    '���±�־
    If UpdateZlBakSpace(mcnOracle, mlng�ռ���, lngSys, False, strDbLink <> "") = False Then
        MsgBox "���±�־ʧ��,����!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    '������Ч����:
    Call ReCompileObjects(mcnOracle)
    ExeFuncChange = True
End Function

Private Sub cmd����_Click()
    Dim objfrmUpSys As frmAppUpgradeNew
    '��������
    '��ʷ���ݿռ�����
    Dim strUserName As String, strPassword As String, strServer As String, strErrMsg As String
    
    strUserName = txtDba�û�.Text
    strPassword = txtDba����.Text
    strServer = txtDbaServer.Text
    
    
    If CheckUser(strUserName, strPassword, strServer, strErrMsg) = False Then
        MsgBox strErrMsg, vbExclamation, gstrSysName
        Exit Sub
    End If
    txtDba�û�.Text = strUserName
    txtDba����.Text = strPassword
    txtDbaServer.Text = strServer
    strPassword = strPassword
    
    Set objfrmUpSys = New frmAppUpgradeNew '�������ģ�����
    If objfrmUpSys.HistoryUp(Me, stbThis.Panels(2), mlngSys, lblSpace.Tag, lblSetupIni.Tag, strUserName, strPassword, strServer, mstrVersion, mstrDBLink) Then
        '����Ҫˢ�½���
        SetPromptText "�����ɹ���"
    Else
        SetPromptText "����ʧ�ܣ�"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If mblnSysUpdateCall Then
        If mblnMustInstall And mintFunType = F0���� And mblnSucced = False Then
            If MsgBox("��ǰϵͳ���谲װ��ʷ���ݿռ�󣬲�������" & vbCrLf & "ʹ�ø�ϵͳ,������ڲ��������Ժ���ǰ��������ת�ƹ���ģ��" & vbCrLf & "���д���!�Ƿ����ڴ�����", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Cancel = 1
                Exit Sub
            End If
        End If
    Else
        If cmdNext.Enabled = False Then
            Cancel = 1
            Exit Sub
        End If
    
        If mblnMustInstall And mintFunType = F0���� And mblnSucced = False Then
            MsgBox "��ǰϵͳ���谲װ��ʷ���ݿռ�󣬲�������" & vbCrLf & "ʹ�ø�ϵͳ,��˲���ȡ������!", vbInformation + vbDefaultButton1, gstrSysName
            Cancel = 1
            Exit Sub
        End If
    End If
    
    Set mrsMergeSpace = Nothing
    
    '���ر����Ӷ���mcnDBA,��Ϊ�ö�������Ǵ���ģ������������ʹ��ʱ�ᵼ�´�������ӱ��ر�
End Sub

Private Sub Image1_DblClick()
    If mintFunType = F3���� Then
        Image1.ToolTipText = "������ʱ�ļ������Ƹ��Ƶ����а�"
        MsgBox "���������ƽű��ļ�����ʵ��ִ�нű�", vbInformation
    End If
End Sub

Private Sub lblIniModi_Click()
    Dim strFile As String
    
    With cdgPub
        .DialogTitle = "ѡ��Ӧ�ð�װ�����ļ�"
        .Filter = "Ӧ�ð�װ�����ļ�(zlSetup.ini)|zlSetup.ini"
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
                '�����ļ��ı䣬��������ļ�
                If CheckInitFile(mlngSys, .Filename) Then
                    lblSetupIni.Caption = "��װ�����ļ���" & .Filename
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
        cmdNext.Caption = "��һ��(&N)"
    Else
        optDele(1).FontBold = True
        optDele(0).FontBold = False
        cmdNext.Caption = "����(&O)"
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

Private Sub txtDba����_GotFocus()
    Call SelAll(txtDba����)
End Sub

Private Sub txtDba�û�_GotFocus()
    Call SelAll(txtDba�û�)
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
            '����һ������
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
    '����:ɾ�����е�ָ��ָ���û���Զ������
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = " select Owner,DB_LINK from all_db_links  where USERNAME='" & strUserName & "'"
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , cnOracle
    With rsTemp
        Do While Not .EOF
            '���ָ�����ͬһ���û���Db_Link����������ȷ
             On Error Resume Next
             gstrSQL = " Drop Database Link " & !Owner & "." & Nvl(!DB_LINK)
             cnOracle.Execute gstrSQL
             err.Clear: err = 0
            .MoveNext
        Loop
    End With
End Function

Private Sub DropTablespace(ByVal strTableSpace As String)
'���ܣ�ɾ��ָ���ı�ռ�
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

Private Function ExeFuncUnInstall(ByVal strTableSpace As String, ByVal strUserName As String, ByVal lng��� As Long, Optional blnErrResume As Boolean = False) As Boolean
'���ܣ�ɾ���Ѿ���װ����ʷ���ݿռ�
    Dim rsTemp As New ADODB.Recordset
    Dim strErrInfo As String, strTBS As String
        
    strErrInfo = ""
    If blnErrResume = False Then
        On Error GoTo errHand   '��жʱ����
    Else
        On Error Resume Next    '��װʧ��ʱ����
    End If
                  
         
    '1.ɾ��ָ��ǰ�û���Զ�����Ӷ���
    Call DropDBLinkOfUser(mcnDBA, strUserName)
    
    '2.ɾ����ϵͳ�����߼�����(��������)
    SetPromptText "����ɾ����ʷ���ݿռ��û�����ض���"
    DoEvents
    gstrSQL = "drop user " & strUserName & " cascade"
    mcnDBA.Execute gstrSQL
    
    
    '3.ɾ����ϵͳ���ݱ�ռ�
    SetPromptText "����ɾ����ʷ���ݱ�ռ�������ļ���"
    If CheckTableSpaceIsUse("��ռ�", strTableSpace, strUserName, mcnDBA) = False Then
        'û�������û�ʹ�ã�����ɾ��
        DropTablespace (strTableSpace)
        DropTablespace (strTableSpace & "_LOB")
        DropTablespace (strTableSpace & "_IDX")
    End If
        
    '4.ɾ����ʷ���ݿռ�Ŀ¼
    gstrSQL = "delete zltools.zlbakspaces where ϵͳ= " & mlngSys & " and ���=" & lng���
    mcnOracle.Execute gstrSQL

    If mstrDBLink <> "" Then
        'ȷ���Ƿ��в�ͬϵͳָ��ͬһ���ӵ�.����ɾ,����ɾ��
        gstrSQL = "Select 1 From ZLTOOLS.zlbakSpaces where upper(DB����)=upper('" & mstrDBLink & "') and ϵͳ<>" & mlngSys
        OpenRecordset rsTemp, gstrSQL, Me.Caption, , , mcnOracle
        If rsTemp.EOF Then
            On Error Resume Next
            gstrSQL = "Drop DataBase Link  " & mstrDBLink
            mcnOracle.Execute gstrSQL
            If err <> 0 Then
                  Call MsgBox("ɾ��Զ������������,��ϸ�������:" & vbCrLf & "(" & err.Number & ") " & vbCrLf & err.Description, vbInformation)
            End If
        End If
    End If

    ExeFuncUnInstall = True
    
    Exit Function
errHand:
    MsgBox err.Description & vbCrLf & "SQL��䣺" & gstrSQL, vbInformation, gstrSysName
End Function

Private Function CreateTbs(ByVal TbsName As String, ByVal TbsFile As String, ByVal TbsSize As Long, ByVal AutoExtend As Boolean, _
     ByVal AutoAllocate As Boolean, ByVal ExtentSize As Integer, ByVal lngFileAmount As Long) As Byte
    '----------------------------------------------
    '���ܣ�ϵͳ�û�,���ݲ���������ռ�,�̶�Ϊ���ع�������(8i��ǰ��֧��,��ʱֻ�ܴ����ֵ��������)
    '������
    '   TbsName:��ռ�����
    '   TbsFile:��ռ��ļ�
    '   TbsSize:��ռ��С(MΪ��λ)
    '   Extend:�Ƿ��Զ�������,����ͳһ��Χ�ߴ�
    '   ExtentSize:ͳһ���ߴ�,��ʱ��ռ����ָ���ߴ�(OracleȱʡΪ1M)
    '   Temp:�Ƿ�Ϊ��ʱ��ռ�
    '   lngFileAmount:�����ļ�������
    '���أ�1-�����ɹ���2-��ռ��Ѿ����ڣ�3-����ʧ��,4-���̿ռ䲻��
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
            MsgBox "�����ı�ռ䣨" & TbsName & "���Ĵ��̿ռ䲻��,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
            CreateTbs = 4
        ElseIf mcnDBA.Errors.Item(0).NativeError = 1119 Then
            Call MsgBox("�����ļ�(" & TbsFile & ")���ô��󣬲��ܼ���!" & vbCrLf & "������Ϣ:" & mcnDBA.Errors(0).Description & vbCrLf & gstrSQL, vbInformation Or vbDefaultButton2, gstrSysName)
            CreateTbs = 3
        Else
            If MsgBox("�������������Ƿ�����������" & vbCrLf & vbTab & mcnDBA.Errors(0).Description & vbCrLf & gstrSQL, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                CreateTbs = 2
            Else
                CreateTbs = 1
            End If
        End If
    Else
        MsgBox "������ռ䣨" & TbsName & "��ʧ��:" & vbCrLf & gstrSQL & vbCrLf & err.Description, vbInformation + vbDefaultButton1, gstrSysName
        CreateTbs = 3
    End If
End Function

Private Function CheckUser(ByRef strUserName As String, ByRef strPassword As String, ByRef strServer As String, ByRef strErrMsg As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:����û���������ͷ����������Ƿ���ȷ
    '���:strUsername-�û���,strPassWord-����,strServer-������
    '����:�������ͬ
    '����:�û��Ϸ�������true,���򷵻�False
    '-------------------------------------------------------------------------------------------------------------------------
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    '��Ч�ַ���Ч��
    If Len(Trim(strUserName)) = 0 Then
        strErrMsg = "�������û�����"
        If txtDba�û�.Enabled And txtDba�û�.Visible Then txtDba�û�.SetFocus
        Exit Function
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            strErrMsg = "�û�������"
            If txtDba�û�.Enabled And txtDba�û�.Visible Then txtDba�û�.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            strErrMsg = "�������"
            If txtDba����.Enabled And txtDba����.Visible Then txtDba����.SetFocus
            
            Exit Function
        End If
    End If
    
    If Trim(strServer) <> "" Then
        If Mid(strServer, Len(strServer) - 1, 1) = "/" Or Mid(strServer, Len(strServer) - 1, 1) = "@" Or Mid(strServer, 1, 1) = "/" Or Mid(strServer, 1, 1) = "@" Then
            strErrMsg = "���������Ӵ�����"
            If txtDbaServer.Enabled And txtDbaServer.Visible Then txtDbaServer.SetFocus
            Exit Function
        End If
    End If
    
    '�����ַ���
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
        strErrMsg = "δ��������!"
        If txtDba����.Enabled And txtDba����.Visible Then txtDba����.SetFocus
        Exit Function
    End If
    
    strUserName = UCase(strUserName)
    
    CheckUser = True
     
End Function

Private Function CheckMoveInPutValid(ByVal strDbLink As String) As Boolean
    '----------------------------------------------------------------------------------
    '����:���ֲ����ʷ���ݿռ������������Ƿ�Ϸ�
    '����:
    '����;�ɹ�����true,���򷵻�false
    '----------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset

    If Val(Trim(txtMoveCode.Text)) <= 0 Then
        MsgBox "��������ȷ�Ŀռ��š�", vbExclamation, gstrSysName
        If txtMoveCode.Enabled And txtMoveCode.Visible Then txtMoveCode.SetFocus
        Exit Function
    End If
    If Val(Trim(txtMoveCode.Text)) > 999 Then
        MsgBox "�ռ��Ų��ܴ���999��", vbExclamation, gstrSysName
        If txtMoveCode.Enabled And txtMoveCode.Visible Then txtMoveCode.SetFocus
         Exit Function
    End If

    If Trim(txtMoveName.Text) = "" Then
        MsgBox "�ռ�������Ч,���顣", vbExclamation, gstrSysName
        If txtMoveName.Enabled And txtMoveName.Visible Then txtMoveName.SetFocus
         Exit Function
    End If

    If ActualLen(Trim(txtMoveName.Text)) > 30 Then
        MsgBox "�ռ����Ƶĳ��Ȳ��ܴ���30���ַ���", vbExclamation, gstrSysName
        If txtMoveName.Enabled And txtMoveName.Visible Then txtMoveName.SetFocus
         Exit Function
    End If
    

    gstrSQL = "Select 1 From zlBakSpaces where ϵͳ=" & mlngSys & " and (���=" & Val(txtMoveCode.Text) & " or upper(����)=upper('" & txtMoveName & "'))"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    If Not rsTemp.EOF Then
        MsgBox "����ı�Ż������Ѿ�����,���������ñ�Ż�����!", vbInformation + vbDefaultButton1, mstrSysName
        If txtMoveCode.Visible And txtMoveCode.Enabled Then txtMoveCode.SetFocus
        rsTemp.Close
        Exit Function
    End If
  
    Dim bytType As Byte, strErrMsg As String
    
    '����Ƿ��������
    If lblDataVer.Tag > lblBakVer.Tag Then
        If chk��ǰ.value = 1 Then
            MsgBox "����ʷ���ݿռ�İ汾�����߲����������ʷ�ռ���Ǩ�������ֲΪ��ǰ!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    ElseIf lblDataVer.Tag < lblBakVer.Tag Then
        If chk��ǰ.value = 1 Then
            MsgBox "����ʷ���ݿռ�İ汾���������߰汾������������ݿ���Ǩ�������ֲΪ��ǰ!", vbInformation + vbDefaultButton1, gstrSysName
            If txtMoveUser.Enabled And txtMoveUser.Visible Then txtMoveUser.SetFocus
            Exit Function
        End If
        bytType = 0
    Else
        bytType = 2
    End If
    
    '�����ص����ݽṹ�Ƿ�Ϸ�
    If CheckHistoryObject(mcnOracle, strDbLink, mlngSys, txtMoveUser.Text, bytType, strErrMsg) = False Then
        MsgBox "���ж�����ʱ�����������´���:" & vbCrLf & strErrMsg
        If chk��ǰ.value = 1 Then
            If txtMoveUser.Enabled And txtMoveUser.Visible Then txtMoveUser.SetFocus
            Exit Function
        End If
    End If
    
    CheckMoveInPutValid = True
End Function

Private Function CheckCreateBakInput() As Boolean
    '----------------------------------------------------------------------------------
    '����:��鴴����ʷ���ݿռ������������Ƿ�Ϸ�
    '����:
    '����;�ɹ�����true,���򷵻�false
    '----------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    If Val(Trim(txt���.Text)) <= 0 Then
        MsgBox "��������ȷ�Ŀռ��š�", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txt���.Enabled Then txt���.SetFocus
        Exit Function
    End If
    If Val(Trim(txt���.Text)) > 999 Then
        MsgBox "�ռ��Ų��ܴ���999��", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txt���.Enabled Then txt���.SetFocus
         Exit Function
    End If
    
    '��ռ���
    If Trim(txtBakSpace.Text) = "" Then
        MsgBox "��������ȷ�Ŀռ����ơ�", vbExclamation, gstrSysName
        tbHistory.Tab = 1
        If txtBakSpace.Enabled And txtBakSpace.Visible = True Then txtBakSpace.SetFocus
         Exit Function
    End If
    If ActualLen(Trim(txtBakSpace.Text)) > 30 Then
        MsgBox "�ռ����ĳ��Ȳ��ܴ���30���ַ���", vbExclamation, gstrSysName
        tbHistory.Tab = 1
        If txtBakSpace.Enabled And txtBakSpace.Visible Then txtBakSpace.SetFocus
         Exit Function
    End If
    
    If Val(txtSpaceSize.Text) > 100000 Then
        MsgBox "��ռ䳬��100G�ˡ�", vbExclamation, gstrSysName
        tbHistory.Tab = 1
        If txtBakSpace.Enabled And txtBakSpace.Visible Then txtBakSpace.SetFocus
        Exit Function
    End If
    If Val(txtSpaceSize.Text) <= 0 Then
        MsgBox "��ռ��������㡣", vbExclamation, gstrSysName
        tbHistory.Tab = 1
        If txtBakSpace.Enabled And txtBakSpace.Visible Then txtBakSpace.SetFocus
        Exit Function
    End If
    
    '�����ļ����
    If InStr(txtDataFile.Text, ".") = 0 Then
        MsgBox "�����ļ�ȱ����չ����", vbExclamation, gstrSysName
        tbHistory.Tab = 1
        If txtDataFile.Enabled And txtDataFile.Visible Then txtDataFile.SetFocus
        Exit Function
    End If
    
    If Val(txtFileAmount(0).Text) <= 0 Or Val(txtFileAmount(1).Text) <= 0 Or Val(txtFileAmount(2).Text) <= 0 Then
        MsgBox "�����ļ�����������㡣", vbExclamation, gstrSysName
        tbHistory.Tab = 1
        If txtFileAmount(0).Enabled And txtFileAmount(0).Visible Then txtFileAmount(0).SetFocus
        Exit Function
    End If
    
    
    If Trim(txtOwnerUsr.Text) = "" Then
        MsgBox "��������ȷ���û�����", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txtOwnerUsr.Enabled Then txtOwnerUsr.SetFocus
         Exit Function
    End If
    If ActualLen(Trim(txtOwnerUsr.Text)) > 30 Then
        MsgBox "�û����ĳ��Ȳ��ܴ���30���ַ���", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txtOwnerUsr.Enabled Then txtOwnerUsr.SetFocus
         Exit Function
    End If
    
    If Trim(txtOwnerPwd.Text) = "" Then
        MsgBox "��������", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txtOwnerPwd.Enabled Then txtOwnerPwd.SetFocus
         Exit Function
    End If
    If Trim(txtOwnerLab.Text) = "" Then
        MsgBox "��������֤���", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txtOwnerLab.Enabled Then txtOwnerLab.SetFocus
         Exit Function
    End If
    
    If Trim(txtOwnerLab.Text) <> Trim(txtOwnerPwd.Text) Then
        MsgBox "����Ŀ�������֤����£�������!", vbExclamation, gstrSysName
        tbHistory.Tab = 0
        If txtOwnerLab.Enabled Then txtOwnerLab.SetFocus
         Exit Function
    End If
    
    If optServer(1).value = True Then
        If Trim(txtDbaServer.Text) = "" Then
            MsgBox "���������������!", vbExclamation, gstrSysName
            tbHistory.Tab = 0
            If txtDbaServer.Visible And txtDbaServer.Enabled Then txtDbaServer.SetFocus
            Exit Function
        End If
        
        If Trim(txtDBLink.Text) = "" Then
            MsgBox "��������DBLink����!", vbExclamation, gstrSysName
            tbHistory.Tab = 0
            If txtDBLink.Enabled Then txtDBLink.SetFocus
            Exit Function
        End If
    End If
        
    
    On Error GoTo errh
     
    gstrSQL = "Select 1 From zlBakSpaces where ϵͳ=" & mlngSys & " and (���=" & Val(txt���.Text) & " or upper(����)=upper('" & txtHD.Text & txtOwnerUsr.Text & "'))"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    If Not rsTemp.EOF Then
        MsgBox "����ı�Ż������ظ�,������!", vbInformation + vbDefaultButton1, mstrSysName
        tbHistory.Tab = 0
        If txt���.Visible And txt���.Enabled Then txt���.SetFocus
        rsTemp.Close
        Exit Function
    End If
    
    CheckCreateBakInput = True
    Exit Function
errh:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Function CheckBakUser(ByRef blnHaveUser As Boolean, ByVal strDbLink As String) As Boolean
'���ܣ���鲢������ʷ�ռ��û�
    Dim rsTemp As New ADODB.Recordset, cnTemp As ADODB.Connection
    Dim strUser As String, strPass As String, strServer As String
    Dim strError As String, strTbsName As String, strSQL As String
    
    SetPromptText "���ڼ���û�����Ч��..."
    strUser = Trim(txtHD.Text & txtOwnerUsr.Text)
    strPass = Trim(txtOwnerPwd.Text)
    strServer = Trim(txtDbaServer.Text)
    
    On Error GoTo errh
    gstrSQL = "select 1 from dba_users where username='" & strUser & "'"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption, , , mcnDBA)
    
    blnHaveUser = rsTemp.RecordCount > 0
    If rsTemp.RecordCount > 0 Then
        If MsgBox("�û���Ϊ��" & strUser & "������ʷ���ݿռ��Ѿ�����,�Ƿ񽫱�ϵͳ����ʷ���ݿռ���ӵ����û���?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            tbHistory.Tab = 0
            If txtOwnerUsr.Enabled And txtOwnerUsr.Visible Then txtOwnerUsr.SetFocus
            Exit Function
        End If
    
        '��������Ƿ�Ϸ�
        Set cnTemp = gobjRegister.GetConnection(strServer, strUser, strPass, False, MSODBC, strError, False)
        If cnTemp.State = adStateClosed Then
            MsgBox strError, vbInformation, gstrSysName
            tbHistory.Tab = 0
            If txtOwnerUsr.Enabled And txtOwnerUsr.Visible Then txtOwnerUsr.SetFocus
            Exit Function
        End If
        '������ʷ�����ƺ������߲�ͬ�������Ҫ���»�ȡ
        '��ʷ���ռ��ȡ
        strSQL = "Select a.Tablespace_Name From User_Tables A Where a.Table_Name = 'ZLBAKINFO'"
        Set rsTemp = gclsBase.OpenSQLRecord(cnTemp, strSQL, "��ȡ��ʷ���ռ�")
        If Not rsTemp.EOF Then
            txtBakSpace.Text = rsTemp!Tablespace_Name
        End If
        strTbsName = UCase(Trim(txtBakSpace.Text))
        '��ȡ������ռ���LOB��ռ�
        strSQL = "Select a.Tablespace_Name" & vbNewLine & _
                "From User_Tablespaces A" & vbNewLine & _
                "Where a.Tablespace_Name In ('" & strTbsName & "', '" & strTbsName & "_IDX', '" & strTbsName & "_LOB')"
        Set rsTemp = gclsBase.OpenSQLRecord(cnTemp, strSQL, "��ȡ��ʷ���ռ�", strTbsName)
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
            MsgBox "��ʷ���ݿռ��û���(" & strUser & ")�����������ݿ�Ҫ�������¶��塣" & vbCrLf & err.Description, vbExclamation, gstrSysName
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
    '����:��鲻ͬϵͳָ����ͬһ�����û��µ����ݽṹ
    '���:cnOracle-�������ݿ�
    '     cnOracleBak-��ʷ���ݿ�����
    '     strBakUserName-��ʷ���ݿռ��������
    '     lngSys-��ǰϵͳ��ϵͳ��
    '--------------------------------------------------------------------------------------------------------------
    Dim rsTemp  As New ADODB.Recordset
    Dim strTemp As String, strSysIn As String
    
    On Error GoTo errHandle
    
    gstrSQL = "Select 1 From ALL_tables" & IIf(strDbLink = "", "", "@" & strDbLink) & " where Owner='" & strBakUserName & "' And Table_name = '" & UCase("zlBakInfo") & "'"
    OpenRecordset rsTemp, gstrSQL, "�����ʷ����Ϣ", , , cnOracle
    
    If rsTemp.RecordCount > 0 Then
        gstrSQL = "Select ϵͳ From " & strBakUserName & ".zlBakInfo"
        OpenRecordset rsTemp, gstrSQL, "��ȡ��ʷ��ϵͳ��", , , cnOracleBak
        strSysIn = ""
        With rsTemp
            Do While Not .EOF
                If lngSys = Val(Nvl(!ϵͳ)) Then
                    MsgBox "ָ������ʷ�ռ����Ѿ����ڸ�ϵͳ��,�����´�����ʷ���ݿռ�,����[��ֲ]����!", vbInformation + vbDefaultButton1, gstrSysName
                    Exit Function
                End If
                strSysIn = strSysIn & "," & Val(Nvl(!ϵͳ))
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
    
    gstrSQL = "select a.���� from zlbaktables a,zlbaktables b where a.����=b.���� and a.ϵͳ IN (" & Mid(strSysIn, 2) & ") and b.ϵͳ=" & lngSys
    OpenRecordset rsTemp, gstrSQL, "��ȡָ��ϵͳ�Ƿ���ڹ����", , , cnOracle
    If rsTemp.EOF Then
        CheckDiffUserStru = True
        Exit Function
    End If
    
    strTemp = ""
    With rsTemp
        Do While Not .EOF
            strTemp = strTemp & "    " & Nvl(!����) & vbCrLf
            .MoveNext
        Loop
    End With
    gstrSQL = "Select 1 from zlsystems where ���=" & lngSys & " and nvl(�����,0) in (" & Mid(strSysIn, 2) & ")"
    OpenRecordset rsTemp, gstrSQL, "��ȡָ��ϵͳ�Ƿ���ڹ���", , , cnOracle
    
    '���ڹ���,�������ж�
    If rsTemp.EOF = False Then
        CheckDiffUserStru = True
        Exit Function
    End If
    MsgBox "��ѡ�����ʷ���ݿռ������±����:" & vbCrLf & strTemp & vbCrLf & " ����ָ������ʷ���ݿռ�!", vbInformation + vbDefaultButton1, gstrSysName
    
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & "���ִ�е�SQL��" & gstrSQL, vbExclamation, Me.Caption
End Function

Private Function CheckTableSpaceIsUse(ByVal strType As String, ByVal strName As String, ByVal strOwner As String, cnOracle As Connection) As Boolean
    '���ܣ�����ռ�������ļ��Ƿ��������û�ʹ��
    '������strType    ��ռ� �����ļ�
    '      strName          ��ռ�������ļ�������
    '      strOwner         �����������û�����������
    Dim rsTemp As New ADODB.Recordset
    
    If strType = "��ռ�" Then
        gstrSQL = "select owner from all_tables where tablespace_name='" & UCase(strName) & "' and owner<>'" & UCase(strOwner) & "' AND ROWNUM<2"
    Else
        gstrSQL = "select O.owner  from V$TABLESPACE T,V$DATAFILE F,all_tables O " & _
                  "Where T.TS# = F.TS# And T.name = O.TABLESPACE_NAME " & _
                  "    and F.name='" & UCase(strName) & "' and O.owner<>'" & UCase(strOwner) & "' AND ROWNUM<2 "
    End If
    
    OpenRecordset rsTemp, gstrSQL, Me.Caption, , , cnOracle
    
    If rsTemp.RecordCount = 0 Then
        'û�������û�ʹ�ã�����ɾ��
        CheckTableSpaceIsUse = False
    Else
        '���û�ʹ��
        CheckTableSpaceIsUse = True
    End If
End Function

'����Ƿ������ʷ���ݿռ�
Private Function IsHavingHistoryTable(ByVal lngSys As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ������ʷ���ݿռ�
    '����:������ʷ�ռ����ݱ�.����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    IsHavingHistoryTable = False
    gstrSQL = "Select ���� From zlBakTables where ϵͳ=" & lngSys
    OpenRecordset rsTemp, gstrSQL, "�����Ƿ����!", , , mcnOracle
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
'���ܣ�����ֻ����������
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> 13 And KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtSpaceSize_Validate(Cancel As Boolean)
    If Not IsNumeric(txtSpaceSize.Text) Then Cancel = True
End Sub

Private Sub txt���_GotFocus()
    Call SelAll(txt���)

End Sub

Private Sub txt���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt���_KeyPress(KeyAscii As Integer)
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
    '����:�����ʷ���ݿռ��������Ƿ�����
    '����:strDbLink-��ʷ���ݿ�DBLink����
    '     cnOracle-�������ݿ�����
    '     lngSys-ϵͳ��
    '     strBakOwnerName-��ʷ���ݿռ�
    '     bytCheckSys-0-��������Ƿ���zlbakInfor���д���ϵͳ(������,1-�����������,>1��ʾȫ���:��Ҫ�Ǽ�����ͱ�
    '����:strErrMsg-������صĴ�����Ϣ
    '����:������Ϸ�,����true,���򷵻�False
    '--------------------------------------------------------------------------------------------------------------------------
    Dim rsBakObject As New ADODB.Recordset
    Dim rsObject As New ADODB.Recordset
    Dim strErrInfor As String
    
    On Error GoTo errHand
    If strDbLink <> "" Then strDbLink = "@" & strDbLink

    gstrSQL = "select table_name as ����  from all_tables" & strDbLink & " where  owner = upper('" & strBakOwnerName & "') "
    OpenRecordset rsBakObject, gstrSQL, Me.Caption, , , cnOracle
    
    '���zlBakInfo���Ƿ����
    rsBakObject.Filter = "����='" & UCase("zlBakInfo") & "'"
    If rsBakObject.EOF Then
        strErrInfor = strErrInfor & vbCrLf & Space(4) & "������:zlBakInfo��"
        strErrMsg = strErrInfor
        Exit Function
    End If
    If (bytCheckSys = 0 Or bytCheckSys > 1) Then
        
        gstrSQL = "Select 1 From " & strBakOwnerName & ".zlBakInfo" & strDbLink & " where ϵͳ=" & lngSys
        OpenRecordset rsObject, gstrSQL, Me.Caption, , , cnOracle
        If rsObject.EOF Then
            strErrInfor = strErrInfor & vbCrLf & Space(4) & "����ʷ���ݿռ��и�ϵͳ������,����!"
            Set rsObject = Nothing
            strErrMsg = strErrInfor
            Exit Function
        End If
        rsObject.Close
    End If
     
    
    If bytCheckSys >= 1 Then
        strErrInfor = ""
        gstrSQL = "Select ���� from zlbakTables where ϵͳ=" & lngSys
        OpenRecordset rsObject, gstrSQL, Me.Caption, , , cnOracle  '��ǰ���ת������
        With rsObject
            Do While Not .EOF
                rsBakObject.Filter = "����='" & Nvl(!����) & "'"
                If rsBakObject.EOF Then '��¼��ʷ���в����ڵ�ת����
                    strErrInfor = strErrInfor & vbCrLf & Space(4) & Nvl(!����)
                End If
                .MoveNext
            Loop
        End With
    End If
    rsBakObject.Close
    Set rsBakObject = Nothing

    If strErrInfor <> "" Then
        strErrInfor = "���±�����ʷ���в�����:" & Mid(strErrInfor, 2) & vbCrLf & "��������ʷ��İ汾̫�͡�" & _
            IIf(chk��ǰ.value = 1, vbCrLf & "������ֲ��֮�����л�Ϊ��ǰ��ʷ�⣨ͬʱ��������", "")
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
    '����:������ؿؼ���Eanble����
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
'���ܣ�����ָ�����������ߵ���Ч����
'������cnThis=����������,����������Բ�ͬ�����ߵ���
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
'���ܣ�ͨ��SQLPlus��Copy���Զ�����ݿ�ķ�ת�����ݸ��Ƶ���ǰ��ʷ��ռ�
'˵��������������ʱ�ļ���ÿ�ŷ�ת��������һ��copy�ű����Ȼ��ͨ��shell��ʽ����sqlplus��ִ����ʱ�ļ��еĶ����ű���
'������strBakTBS=��ʷ���û��ı�ռ�����
    Dim rsUnHistory As New ADODB.Recordset
    Dim objFSO As New FileSystemObject
    Dim objScript As Scripting.TextStream
    Dim strScript As String, strFile As String, strErrInfo As String
    Dim lngErrNum As Long, lngCommand As Long, i As Long, lngProcess As Long
    
    gstrSQL = "Select Table_Name From All_Tables Where Owner = '" & mstrOwnerName & _
            "' Minus Select ���� From zlBakTables Order By Table_Name"
    'gstrSQL = "Select Table_Name From (" & gstrSQL & ") Where Table_Name like '��Ա��'"
    OpenRecordset rsUnHistory, gstrSQL, Me.Caption
    
    If rsUnHistory.RecordCount = 0 Then
        MsgBox "��ǰ��������û���ҵ���ת����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '������ʱ�ű��ļ�
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

    '����SQLPlus����
    strScript = "sqlplus " & mstrOwnerName & "/" & mstrOwnerPass & "@" & gstrServer & " @" & strFile

    'ִ��Shell����
    err.Clear: On Error Resume Next
    
    SetPromptText "����ͨ��sqlplus��Copy�����" & rsUnHistory.RecordCount & "�ű�����ݣ����Եȡ�"
    
    If Not Image1.ToolTipText = "������ʱ�ļ������Ƹ��Ƶ����а�" Then
        lngCommand = Shell(strScript, vbHide)
    End If
    
    If err.Number <> 0 Then
        lngErrNum = err.Number '53:�ļ�δ�ҵ�
        strErrInfo = err.Description & IIf(lngErrNum = 53, ",���� sqlplus.exe �Ƿ���ȷ��װ", "")
        err.Clear
        SetPromptText ""
        Call MsgBox("����:" & lngErrNum & vbCrLf & strErrInfo, vbInformation, gstrSysName)
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
        SetPromptText "���Ʒ�ת�����������"
        ExeFuncCopy = True
    End If
    
    
    
    If Image1.ToolTipText = "������ʱ�ļ������Ƹ��Ƶ����а�" Then
        Call Clipboard.SetText(strFile)
    Else
        objFSO.DeleteFile strFile
        Set objFSO = Nothing
        
        rsUnHistory.MoveFirst
        For i = 1 To rsUnHistory.RecordCount
            SetPromptText "���ڴ�����ת�����Լ��������(" & i & "/" & rsUnHistory.RecordCount & ")��" & rsUnHistory!Table_Name
   
            '������ṹ��ص�PK��UQ
            Call CreateConstraint(rsUnHistory!Table_Name, strBakTBS, mstrOwnerName, strBakUserName)
            '������ṹ����IX
            Call CreateIndex(rsUnHistory!Table_Name, strBakTBS, mstrOwnerName, strBakUserName)
            
            DoEvents
            rsUnHistory.MoveNext
        Next
    End If
End Function

Private Function GetBakTableSpace(ByRef cnBakOracle As ADODB.Connection, ByVal strBakUser As String) As String
'���ܣ�������ʷ�ռ����ӣ�����ָ����ʷ�ռ��û��ı�ռ�����
    Dim rsHistory As New ADODB.Recordset

    gstrSQL = "select ���� from zlbakspaces where ������='" & strBakUser & "'"
    OpenRecordset rsHistory, gstrSQL, Me.Caption, , , cnBakOracle
    
    If rsHistory.RecordCount > 0 Then
        GetBakTableSpace = rsHistory!����
    End If
End Function


Private Function ExeFuncMerge() As Boolean
'���ܣ��ϲ��б���ѡ��Ŀռ䣬�����������С�Ŀռ䡣
'˵����1.�Ƚ��ñ����ռ��ϵ�Լ��������
'      2.Ȼ��ӿռ�����С�Ŀ�ʼ����������(��zlbaktables�ж���ı�)�������ռ���
'      3.ÿ�������һ���ռ䣬��ɾ��һ���ռ���û�����ռ��ļ���zlbakspaces�еļ�¼
'      4.���пռ�����ݺϲ���ɺ��ؽ�������ռ��Լ��������
    Dim i As Long, lngLoop As Long
    Dim strKeepVersion As String, strKeepOwner As String, strMergeOwner As String
    Dim strPreTable As String, strTableSpace As String
    Dim strError As String, strTables As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsBakTables As New ADODB.Recordset
    Dim rsDelSpace As New ADODB.Recordset
    Dim blnDisibled As Boolean
    
    On Error GoTo errHandle
    '1.���
    SetPromptText "���ڼ��Ҫ�ϲ��ı�ռ䡣"
    '1.1���汾
    '------------------------------------------------------------------------------
    gstrSQL = ""
    For i = 1 To mrsMergeSpace.RecordCount
        If i = mrsMergeSpace.RecordCount Then
            gstrSQL = gstrSQL & "Select '" & mrsMergeSpace!������ & "' As ������, �汾��," & mrsMergeSpace!��� & " as ��� From " & mrsMergeSpace!������ & ".Zlbakinfo"
        Else
            gstrSQL = gstrSQL & "Select '" & mrsMergeSpace!������ & "' As ������, �汾��," & mrsMergeSpace!��� & " as ��� From " & mrsMergeSpace!������ & ".Zlbakinfo Union All" & vbCrLf
        End If
        mrsMergeSpace.MoveNext
    Next
    OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    rsTemp.Filter = "���=" & mlng�ռ���
    strKeepVersion = rsTemp!�汾��
    strKeepOwner = rsTemp!������
    rsTemp.Filter = "���<>" & mlng�ռ���
    
    For i = 1 To rsTemp.RecordCount
        If strKeepVersion <> rsTemp!�汾�� Then
            strError = rsTemp!������ & ":" & rsTemp!�汾��
        End If
        strMergeOwner = strMergeOwner & ",'" & rsTemp!������ & "'"
        rsTemp.MoveNext
    Next
    If strError <> "" Then
        MsgBox "Ҫ��������ʷ�ռ�汾Ϊ" & strKeepVersion & ",��Ҫ�ϲ�����ʷ�ռ�汾��һ��:" & vbCrLf & strError & _
                vbCrLf & "����ͨ��[�л�]������������ʷ���ݿռ䡣"
        Exit Function
    End If
    strMergeOwner = Mid(strMergeOwner, 2)
    
    
    '1.2���ϲ��ı�ռ����Ƿ����zlbaktables����ı�������ڣ�����ʾ��Щ���ݽ����ںϲ���ɾ����
    '------------------------------------------------------------------------------
    Set rsBakTables = New ADODB.Recordset
    gstrSQL = "Select ���� From zlBakTables Where ϵͳ = " & mlngSys & " Order By ����"
    OpenRecordset rsBakTables, gstrSQL, Me.Caption
    
    gstrSQL = "Select Owner, Table_Name From All_Tables Where Owner In (" & strMergeOwner & ") And Table_Name<>'ZLBAKINFO' Order By Owner, Table_Name"
    Set rsDelSpace = New ADODB.Recordset
    OpenRecordset rsDelSpace, gstrSQL, Me.Caption
    
    '���ÿ��Ҫ�ϲ��ı�ռ䣬�Ƿ����zlbaktables����ı�
    mrsMergeSpace.MoveFirst
    mrsMergeSpace.Filter = "���<>" & mlng�ռ���  '�����ռ䲻�ü��
    
    strError = ""
    For lngLoop = 1 To mrsMergeSpace.RecordCount
        rsDelSpace.Filter = "Owner='" & mrsMergeSpace!������ & "'"
        strTables = ""
        For i = 1 To rsDelSpace.RecordCount
            rsBakTables.Filter = "����='" & rsDelSpace!Table_Name & "'"
            If rsBakTables.RecordCount = 0 Then
                '����zlbaktables����ı�
                strTables = strTables & "," & rsDelSpace!Table_Name
            End If
            rsDelSpace.MoveNext
        Next
        If strTables <> "" Then
            strError = strError & mrsMergeSpace!������ & ":" & Mid(strTables, 2) & vbCrLf
        End If
        mrsMergeSpace.MoveNext
    Next
    If strError <> "" Then
        If MsgBox("��鷢�ֺϲ������ݿռ��д��ڷ�ת�����ϲ�����Щ�����ݽ��ᱻɾ����" & vbCrLf & strError & vbCrLf _
            & "��ȷ���Ѵ�����Ч���ݣ�������Щ���ݲ�����Ҫ��" & vbCrLf & "��ȷ��Ҫ������", vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
            Exit Function
        End If
    End If
    
    
    '1.3����ֶε�һ����(�����ռ��еı��ɾ���ռ�ı�࣬�ֶ�������ͬ��������ͬ�����ȿ����ɣ�����>�ϲ���)
    '------------------------------------------------------------------------------
    Set rsTemp = New ADODB.Recordset
    gstrSQL = "Select Owner, Table_Name, Column_Name, Data_Type, Decode(Data_Type, 'VARCHAR2', Data_Length, Data_Precision) ����," & vbNewLine & _
                "       Data_Scale ���־���" & vbNewLine & _
                "From All_Tab_Columns" & vbNewLine & _
                "Where Owner In ('" & strKeepOwner & "') And Table_Name In(Select ���� From zlBakTables Where ϵͳ = " & mlngSys & ")" & vbNewLine & _
                "Order By Table_Name"
    OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    Set rsDelSpace = New ADODB.Recordset
    gstrSQL = "Select Owner, Table_Name, Column_Name, Data_Type, Decode(Data_Type, 'VARCHAR2', Data_Length, Data_Precision) ����," & vbNewLine & _
                "       Data_Scale ���־���" & vbNewLine & _
                "From All_Tab_Columns" & vbNewLine & _
                "Where Owner In (" & strMergeOwner & ") And Table_Name In(Select ���� From zlBakTables Where ϵͳ = " & mlngSys & ")" & vbNewLine & _
                "Order By Owner, Table_Name"
    OpenRecordset rsDelSpace, gstrSQL, Me.Caption
    
    strTables = ""
    strError = ""
    strPreTable = ""
    mrsMergeSpace.MoveFirst
    mrsMergeSpace.Filter = "���<>" & mlng�ռ���  '�����ռ䲻�ü��
    
    For lngLoop = 1 To mrsMergeSpace.RecordCount
        rsDelSpace.Filter = "Owner='" & mrsMergeSpace!������ & "'"
        '������������
        strPreTable = ""
        
        For i = 1 To rsDelSpace.RecordCount
            SetPromptText "���ڼ��" & mrsMergeSpace!������ & "�ı�ṹ��" & rsDelSpace!Table_Name
            rsTemp.Filter = "Table_Name='" & rsDelSpace!Table_Name & "'"
            If rsTemp.RecordCount = 0 Then
                 'ɾ���ռ��е�ת�����ڱ����ռ��в�����(���涺��ǰ�Ŀո����ں���Mid(xx,4)))
                strTables = strTables & "  , ȱ��[" & rsDelSpace!Table_Name & "]"
            Else
                '����ֶ�
                rsTemp.Filter = "Table_Name='" & rsDelSpace!Table_Name & "' And Column_Name='" & rsDelSpace!Column_Name & "'"
                If rsTemp.RecordCount = 0 Then
                    '�����ռ���ȱ�ֶ�
                    If strPreTable <> rsDelSpace!Table_Name Then
                        strError = strError & vbCrLf & rsDelSpace!Owner & "." & rsDelSpace!Table_Name & ":" & rsDelSpace!Column_Name
                        strPreTable = rsDelSpace!Table_Name
                    Else
                        strError = strError & "," & rsDelSpace!Column_Name
                    End If
                Else
                    '�ֶ����ͣ����ȣ�����
                    If rsTemp!DATA_TYPE <> rsDelSpace!DATA_TYPE Then
                        If strPreTable <> rsDelSpace!Table_Name Then
                            strError = strError & vbCrLf & rsDelSpace!Owner & "." & rsDelSpace!Table_Name & ":" & rsDelSpace!Column_Name & " " & rsDelSpace!DATA_TYPE
                            strPreTable = rsDelSpace!Table_Name
                        Else
                            strError = strError & "," & rsDelSpace!Column_Name & " " & rsDelSpace!DATA_TYPE
                        End If
                    ElseIf rsDelSpace!DATA_TYPE = "VARCHAR2" Then
                        If rsTemp!���� < rsDelSpace!���� Then
                            If strPreTable <> rsDelSpace!Table_Name Then
                                strError = strError & vbCrLf & rsDelSpace!Owner & "." & rsDelSpace!Table_Name & ":" & rsDelSpace!Column_Name & " " & rsDelSpace!DATA_TYPE & "(" & rsDelSpace!���� & ")"
                                strPreTable = rsDelSpace!Table_Name
                            Else
                                strError = strError & "," & rsDelSpace!Column_Name & " " & rsDelSpace!DATA_TYPE & "(" & rsDelSpace!���� & ")"
                            End If
                        End If
                   ElseIf rsDelSpace!DATA_TYPE = "NUMBER" Then
                        If rsTemp!���� < rsDelSpace!���� Or rsTemp!���־��� <> rsDelSpace!���־��� Then
                            If strPreTable <> rsDelSpace!Table_Name Then
                                strError = strError & vbCrLf & rsDelSpace!Owner & "." & rsDelSpace!Table_Name & ":" & rsDelSpace!Column_Name & " " & rsDelSpace!DATA_TYPE & "(" & rsDelSpace!���� & "," & rsDelSpace!���־��� & ")"
                                strPreTable = rsDelSpace!Table_Name
                            Else
                                strError = strError & "," & rsDelSpace!Column_Name & " " & rsDelSpace!DATA_TYPE & "(" & rsDelSpace!���� & "," & rsDelSpace!���־��� & ")"
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
        MsgBox "��ϲ��ռ����½ṹ���쵼�²��ܼ���������ִ�����ݽṹ������" & strError, vbInformation, gstrSysName
        Exit Function
    End If
        
    
    '2.�ϲ�ǰ�Ĵ�������Լ��������(���ܱ�ĳ��⣬��Ϊ��ѯ��Ҫ��Щ����)
    '------------------------------------------------------------------------------------------------------------------
    DoEvents
    blnDisibled = True
    SetPromptText "���ڽ�����ʷ���ݿռ�" & strKeepOwner & "��������Ψһ��Լ����"
    Call SetConstraintStatus(mcnOracle, False, strKeepOwner)
    SetPromptText "���ڽ�����ʷ���ݿռ�" & strKeepOwner & "��������"
    Call SetIndexStatus(mcnOracle, False, strKeepOwner)
    
    
    '3.ִ�кϲ�
    '��ͨ��Ĳ�����ɾ��(���������Ʒϵͳ��)
    '���ܱ�ĸ���
    mrsMergeSpace.MoveFirst
    mrsMergeSpace.Filter = "���<>" & mlng�ռ���  '�����ռ����
    For lngLoop = 1 To mrsMergeSpace.RecordCount
        strTableSpace = mrsMergeSpace!����
        strMergeOwner = mrsMergeSpace!������
        
        '3.1���ݴ���
        SetPromptText "���ںϲ���ʷ���ݿռ�" & strTableSpace & "�����ݡ�"
        
        gstrSQL = "Zl1_Datamove_Merge(" & strKeepOwner & "," & strMergeOwner & ")"
        Call ExecuteProcedure(gstrSQL, Me.Caption)
        
        
        '3.2.ɾ��ָ��ǰ�û���Զ�����Ӷ���
        Call DropDBLinkOfUser(mcnOracle, strMergeOwner)
        
        '3.3.ɾ����ϵͳ�����߼�����(��������)
        SetPromptText "����ɾ����ʷ���ݿռ�" & strTableSpace & "���û�����ض���"
        gstrSQL = "drop user " & strMergeOwner & " cascade"
        mcnDBA.Execute gstrSQL
        
        
        '3.4.ɾ����ϵͳ���ݱ�ռ�
        SetPromptText "����ɾ����ʷ���ݱ�ռ�" & strTableSpace & "�������ļ���"
        If CheckTableSpaceIsUse("��ռ�", strTableSpace, strMergeOwner, mcnDBA) = False Then
            'û�������û�ʹ�ã�����ɾ��
            gstrSQL = "alter tablespace " & strTableSpace & " offline"
            mcnDBA.Execute gstrSQL
            gstrSQL = "drop tablespace " & strTableSpace & " including contents and datafiles cascade constraints"
            mcnDBA.Execute gstrSQL
        Else
            MsgBox "��ռ�" & strTableSpace & "���������û��Ķ������ƶ���Щ������ֹ�ɾ����ռ估�ļ���", vbInformation
        End If
            
        '3.5.ɾ����ʷ���ݿռ�Ŀ¼(���ܶ��ϵͳ����һ����ʷ���ݿռ�)
        gstrSQL = "delete zltools.zlbakspaces where ����= '" & strTableSpace & "'"
        mcnOracle.Execute gstrSQL
        
        mrsMergeSpace.MoveNext
    Next
    
    
    '4.����Լ����������ɾ����ʷ�ռ��û�����ռ������ļ�
    SetPromptText "����������ʷ���ݿռ�" & strKeepOwner & "��������"
    Call SetIndexStatus(mcnOracle, True, strKeepOwner)
    SetPromptText "����������ʷ���ݿռ�" & strKeepOwner & "��������Ψһ��Լ����"
    Call SetConstraintStatus(mcnOracle, True, strKeepOwner)

    
    SetPromptText "�ϲ�������ɡ�"
    MsgBox "�ϲ���ʷ���ݿռ�ɹ���", vbInformation, gstrSysName
    ExeFuncMerge = True
    
    Exit Function
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
    
    If blnDisibled Then
        SetPromptText "����������ʷ���ݿռ�" & strKeepOwner & "��������"
        Call SetIndexStatus(mcnOracle, True, strKeepOwner)
        SetPromptText "����������ʷ���ݿռ�" & strKeepOwner & "��������Ψһ��Լ����"
        Call SetConstraintStatus(mcnOracle, True, strKeepOwner)
    End If
End Function


Private Sub SetIndexStatus(ByRef cnThis As ADODB.Connection, ByVal blnEnable As Boolean, ByVal strOwner As String)
'����:���û��������������ú������ʷ������ݲ����ٶ�
'     ����ʱ���ù���ִ��Ҫ����SetConstraintStatus������������Ψһ���ֶδ�����Ч��������������,ORA-14063
'����:cnThis-���Ӷ���
'     blnEnable-���������ԣ�true-�������� false -��������
'     strOwner=��ʷ�ռ��������
    Dim rsTmp As New ADODB.Recordset
    Dim cmdTmp As New ADODB.Command
    Dim strSQL As String
    

    '���ڹ����Ż��ӿ�SQLִ��
    If blnEnable Then
        strSQL = "Select /*+ rule*/" & vbNewLine & _
                " 'alter index " & strOwner & ".' || a.Index_Name || ' Rebuild' Sql" & vbNewLine & _
                "From All_Indexes A, Zltools.Zlbaktables T" & vbNewLine & _
                "Where a.Owner = '" & strOwner & "' And a.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And t.ֱ��ת�� = 1 And a.Status = 'UNUSABLE' And a.Index_Type = 'NORMAL' And" & vbNewLine & _
                "      Not Exists" & vbNewLine & _
                " (Select 1 From All_Constraints C Where c.Owner = a.Owner And c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U'))"
    Else
        strSQL = "Select /*+ rule*/" & vbNewLine & _
                " 'alter index " & strOwner & ".' || a.Index_Name || ' unusable' Sql" & vbNewLine & _
                "From All_Indexes A, Zltools.Zlbaktables T" & vbNewLine & _
                "Where a.Owner = '" & strOwner & "' And a.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And t.ֱ��ת�� = 1 And a.Status = 'VALID' And a.Index_Type = 'NORMAL' And Not Exists" & vbNewLine & _
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
            '�������������ʹ�ã���ֻ�������ؽ����Ƚ���
            If InStr(err.Description, "ORA-00054") > 0 Then
                err.Clear
                strSQL = Replace(rsTmp!SQL, "Rebuild", "Rebuild Online")
                cmdTmp.CommandText = strSQL
                cmdTmp.Execute
            Else
                Call MsgBox("����:" & err.Description & vbCrLf & strSQL, vbInformation, "�����ؽ�")
                err.Clear
            End If
        End If
        
        rsTmp.MoveNext
    Wend
End Sub

Private Sub SetConstraintStatus(ByRef cnThis As ADODB.Connection, ByVal blnEnable As Boolean, ByVal strOwner As String)
'����:���û����õ�Լ�������ú������ʷ������ݲ����ٶ�
'     ����������Ψһ�����ɾ����Ӧ������
'����:cnThis-���Ӷ���
'     blnEnable=true-����Լ��,false-����Լ��
'     strOwner=��ʷ�ռ��������

    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim cmdTmp As New ADODB.Command
        
    '��ʷ��û�����������Լ�������ԣ�����ȫ��������Ψһ��
    If blnEnable Then
        'ע�⣺����������Ψһ��������ʹ��novalidate��ʽ���⽫���¶�Ӧ������������
        strSQL = "Select " & vbNewLine & _
                " 'ALTER TABLE " & strOwner & ".'|| a.Table_Name || ' enable constraint ' || a.Constraint_Name Sql" & vbNewLine & _
                "From All_Constraints A, Zltools.Zlbaktables T" & vbNewLine & _
                "Where a.Owner = '" & strOwner & "' And a.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And t.ֱ��ת�� = 1 And a.Status = 'DISABLED'"
    Else
        strSQL = "Select " & vbNewLine & _
                " 'ALTER TABLE " & strOwner & ".' || a.Table_Name || ' disable constraint ' || a.Constraint_Name || Decode(a.Constraint_Type,'P',' Cascade drop index','U',' Cascade drop index','') Sql" & vbNewLine & _
                "From All_Constraints A, Zltools.Zlbaktables T, All_Tables b" & vbNewLine & _
                "Where a.Owner = '" & strOwner & "' And a.Table_Name = t.���� And t.ϵͳ = " & mlngSys & " And t.ֱ��ת�� = 1 And a.Status = 'ENABLED' And a.Table_Name = b.Table_Name And a.Owner = b.Owner And b.Iot_Type Is Null"
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
            Call MsgBox("����:" & err.Description & vbCrLf & strSQL, vbInformation, "Լ������")
            err.Clear
        End If
        rsTmp.MoveNext
    Wend
End Sub

Private Function ExeFuncTrans() As Boolean
'���ܣ���鲢ִ�б�ռ䴫��
    Dim strTbsName As String, strBakUserName As String, strBakNO As String, strBakUserPwd As String
    Dim strPath As String, strSplit As String
    Dim strServerFrom As String, strPWDFrom As String
    Dim rsTmp As ADODB.Recordset
    Dim lngLoop As Long
    
    On Error GoTo errHandle
    strTbsName = lvwHistory.SelectedItem.SubItems(C1����)
    strBakUserName = lvwHistory.SelectedItem.SubItems(C4������)
    strBakNO = Mid(lvwHistory.SelectedItem.Key, 2)
    strBakUserPwd = txtBakPWD.Text
    strServerFrom = txtDbaServer.Text
    strPWDFrom = txtDba����.Text
    
    '1.��Դ���ݿ⽨Ŀ¼����Ҫ�����ָ���ռ��ļ���λ�ã�
    '--------------------------------------------------------------
    SetPromptText "��Դ�ⴴ��Ŀ¼ZLTRANSFROM"
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
    Else 'linux��ƽ̨��/
        strSplit = "/"
    End If
    strPath = Mid(strPath, 1, InStrRev(strPath, strSplit) - 1)

    gstrSQL = "CREATE DIRECTORY ZLTRANSFROM AS '" & strPath & "'"
    mcnDBA.Execute gstrSQL
'    gstrSQL = "GRANT READ, WRITE ON DIRECTORY ZLTRANSFROM TO SYSTEM"   '���ø��Լ���Ȩ
'    mcnDBA.Execute gstrSQL
    
        
    '2.Դ���ݿ⴫���ռ�������Ƿ�洢��������ռ�
    gstrSQL = "Select Index_Name From Dba_Indexes Where Table_Owner = '" & strBakUserName & "' And Tablespace_Name <> '" & strTbsName & "'"
    Set rsTmp = New ADODB.Recordset
    Call OpenRecordset(rsTmp, gstrSQL, Me.Caption, , , mcnDBA)
    If rsTmp.RecordCount > 0 Then
        SetPromptText "Դ����" & rsTmp.RecordCount & "��������������ռ䣬�����ؽ�����ʷ��ռ�" & strTbsName
        DoEvents
        For lngLoop = 1 To rsTmp.RecordCount
            gstrSQL = "alter index " & rsTmp!Index_Name & " rebuild tablespace " & strTbsName
            mcnDBA.Execute gstrSQL
            rsTmp.MoveNext
        Next
    End If
            
    
    
    '3.��Ŀ��⽨����Ŀ¼���û������ݿ���·
    '----------------------------------------------------
     '3.1��Ŀ��⽨����Ŀ¼
    SetPromptText "��Ŀ��ⴴ��Ŀ¼����ZLTRANSTO"
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
    Else 'linux��ƽ̨��/
        strSplit = "/"
    End If
    strPath = Mid(strPath, 1, InStrRev(strPath, strSplit) - 1)
    
    gstrSQL = "CREATE DIRECTORY ZLTRANSTO AS '" & strPath & "'"
    mcnOracle.Execute gstrSQL
    
    
     '3.2��Ŀ��⽨�û�
    SetPromptText "��Ŀ��ⴴ����ʷ�ռ��û�" & strBakUserName
    '��cmdnext�����ж��Ƿ���ͬ���û�
    gstrSQL = "create user " & strBakUserName & " identified by " & strBakUserPwd '& _
              '" DEFAULT TABLESPACE " & strTbsName   '��ռ����ڻ�������
    mcnOracle.Execute gstrSQL
    
    gstrSQL = "Grant Connect,Resource,UNLIMITED TABLESPACE," & _
            " Create Table,Create Sequence,Create Role,Create User,Drop User,Create Public Synonym,Drop Public Synonym," & _
            " Alter Session,Create Session,Create Synonym,Create View,Create Database Link,Create Cluster" & _
            " to " & strBakUserName & " With Admin Option"
    mcnOracle.Execute gstrSQL
        
        
     '3.3��Ŀ��⽨���ݿ���·
    SetPromptText "��Ŀ��ⴴ�����ݿ���·ZLTRANSTBS"
    gstrSQL = "Select 1 From Dba_Db_Links Where Db_Link||'.' Like 'ZLTRANSTBS.%' And Owner = '" & gstrUserName & "'"
    Set rsTmp = New ADODB.Recordset
    OpenRecordset rsTmp, gstrSQL, Me.Caption, , , mcnOracle
    If rsTmp.RecordCount > 0 Then
        gstrSQL = "Drop DATABASE LINK ZLTRANSTBS"
        mcnOracle.Execute gstrSQL
    End If
    
    gstrSQL = "CREATE DATABASE LINK ZLTRANSTBS CONNECT TO SYSTEM IDENTIFIED BY " & strPWDFrom & " USING '" & strServerFrom & "'"
    mcnOracle.Execute gstrSQL
    
    
    '4.1��Ŀ���ִ�д���
    '--------------------------------------------------------------
    SetPromptText "���ڴ�����ʷ���ݿռ�" & strTbsName
    DoEvents
    gstrSQL = "DBMS_STREAMS_TABLESPACE_ADM.PULL_SIMPLE_TABLESPACE('" & strTbsName & "', 'ZLTRANSTBS', 'ZLTRANSTO')"
    'exec DBMS_STREAMS_TABLESPACE_ADM.pull_simple_tablespace(tablespace_name => ,database_link => ,directory_object => ,conversion_extension => )
    mcnOracle.Execute gstrSQL
    
    '4.2�޸�ȱʡ��ռ�
    gstrSQL = "alter user " & strBakUserName & " DEFAULT TABLESPACE " & strTbsName
    mcnOracle.Execute gstrSQL
    
    '4.3��Ŀ���ֲ��
    gstrSQL = "Delete zltools.zlbakspaces where ϵͳ=" & mlngSys & " And ���=" & strBakNO  'Ϊ��֧��ʧ�ܺ���ٴ�ת��
    mcnOracle.Execute gstrSQL
    Call ExeFuncImport(mcnOracle, strBakNO, strTbsName, strBakUserName, mlngSys)
    
    '4.4��Ŀ���ɾ����������·
    gstrSQL = "Drop DATABASE LINK ZLTRANSTBS"
    mcnOracle.Execute gstrSQL
    
    '4.5��Ŀ���ɾ��Ŀ¼
    gstrSQL = "Drop DIRECTORY ZLTRANSTO"
    mcnOracle.Execute gstrSQL
    
    '4.6��Դ��ɾ��Ŀ¼
    gstrSQL = "Drop DIRECTORY ZLTRANSFROM"
    mcnDBA.Execute gstrSQL
    
    
    
    '5.��Դ����ɾ���Ѵ���ı�ռ�
    '-----------------------------------------------------------------
    DoEvents
    '5.1.ɾ����ϵͳ�����߼�����(��������)
    SetPromptText "����ɾ��Դ����ʷ���ݿռ�" & strTbsName & "���û�����ض���"
    gstrSQL = "drop user " & strBakUserName & " cascade"
    mcnDBA.Execute gstrSQL
    
    
    '5.2.ɾ����ϵͳ���ݱ�ռ�
    SetPromptText "����ɾ����ʷ���ݱ�ռ�" & strTbsName & "�������ļ���"
    gstrSQL = "alter tablespace " & strTbsName & " offline"
    mcnDBA.Execute gstrSQL
    gstrSQL = "drop tablespace " & strTbsName & " including contents and datafiles cascade constraints"
    mcnDBA.Execute gstrSQL
        
    '5.3.ɾ����ʷ���ݿռ�����¼
    gstrSQL = "delete zltools.zlbakspaces where ϵͳ= " & mlngSys & " and ���=" & strBakNO
    mcnDBA.Execute gstrSQL
    
    
    ExeFuncTrans = True
    Exit Function
errHandle:
    If InStr(err.Description, "ORA-00900") > 0 Then
        Call MsgBox("�ռ䴫��ʧ�ܣ���Ŀ���������SQLPlus��ִ������SQL���鿴��ϸ�Ĵ�����Ϣ��" & vbCrLf & gstrSQL, vbInformation, "�ռ䴫��")
    Else
        Call MsgBox("����:" & err.Description & vbCrLf & gstrSQL, vbInformation, "�ռ䴫��")
    End If
    
End Function

