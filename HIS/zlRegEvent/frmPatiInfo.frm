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
   Caption         =   "�ҺŲ�����Ϣ"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
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
   StartUpPosition =   1  '����������
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
      Caption         =   "����(&H)"
      Height          =   420
      Left            =   90
      TabIndex        =   53
      Top             =   7815
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����(&X)"
      Height          =   420
      Left            =   6450
      TabIndex        =   51
      ToolTipText     =   "�ȼ���F2"
      Top             =   7785
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
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
      Begin VB.TextBox txt��֤ 
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
      Begin VB.TextBox txt���� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3795
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   120
         Width           =   1725
      End
      Begin VB.TextBox txt���� 
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
      Begin VB.Label lbl��֤ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��֤"
         Height          =   240
         Left            =   5790
         TabIndex        =   75
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
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
         Caption         =   "����"
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
      Begin ZlPatiAddress.PatiAddress padd��ͥ��ַ 
         Height          =   360
         Left            =   1170
         TabIndex        =   18
         Tag             =   "��סַ"
         Top             =   2100
         Visible         =   0   'False
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   100
      End
      Begin ZlPatiAddress.PatiAddress padd���ڵ�ַ 
         Height          =   360
         Left            =   1170
         TabIndex        =   21
         Tag             =   "���ڵ�ַ"
         Top             =   2505
         Visible         =   0   'False
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Caption         =   "��λ��Ϣ"
         Height          =   750
         Left            =   30
         TabIndex        =   115
         Top             =   5250
         Width           =   11415
         Begin VB.CommandButton cmd��λ���� 
            Caption         =   "��"
            Height          =   360
            Left            =   5520
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   270
            Width           =   360
         End
         Begin VB.TextBox txt��λ���� 
            Height          =   360
            Left            =   660
            MaxLength       =   100
            TabIndex        =   39
            Top             =   270
            Width           =   4860
         End
         Begin VB.TextBox txt��λ�ʱ� 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   6540
            MaxLength       =   6
            TabIndex        =   40
            Top             =   270
            Width           =   1680
         End
         Begin VB.TextBox txt��λ�绰 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   9120
            MaxLength       =   20
            TabIndex        =   41
            Top             =   270
            Width           =   2205
         End
         Begin VB.Label lbl��λ���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   240
            Left            =   135
            TabIndex        =   119
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl��λ�ʱ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ʱ�"
            Height          =   240
            Left            =   6015
            TabIndex        =   117
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl��λ�绰 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�绰"
            Height          =   240
            Left            =   8580
            TabIndex        =   116
            Top             =   330
            Width           =   480
         End
      End
      Begin VB.Frame fraContact 
         Caption         =   "��ϵ����Ϣ"
         Height          =   720
         Left            =   30
         TabIndex        =   110
         Top             =   4500
         Width           =   11415
         Begin VB.TextBox txt��ϵ�����֤ 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   9120
            MaxLength       =   18
            TabIndex        =   38
            Top             =   270
            Width           =   2205
         End
         Begin VB.TextBox txt��ϵ�˵绰 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   3450
            MaxLength       =   18
            TabIndex        =   35
            Top             =   270
            Width           =   1590
         End
         Begin VB.TextBox txt��ϵ������ 
            Height          =   360
            Left            =   630
            MaxLength       =   64
            TabIndex        =   34
            Top             =   270
            Width           =   2160
         End
         Begin VB.TextBox txt������ϵ 
            Height          =   360
            Left            =   6975
            MaxLength       =   30
            TabIndex        =   37
            Top             =   270
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.ComboBox cbo��ϵ�˹�ϵ 
            Height          =   360
            Left            =   5790
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   270
            Width           =   2445
         End
         Begin VB.Label lbl��ϵ�����֤ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���֤"
            Height          =   240
            Left            =   8355
            TabIndex        =   114
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lbl��ϵ������ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   240
            Left            =   135
            TabIndex        =   113
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl��ϵ�˵绰 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�绰"
            Height          =   240
            Left            =   2925
            TabIndex        =   112
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl��ϵ�˹�ϵ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ"
            Height          =   240
            Left            =   5250
            TabIndex        =   111
            Top             =   330
            Width           =   480
         End
      End
      Begin VB.TextBox txt���ڵ�ַ�ʱ� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   10140
         MaxLength       =   6
         TabIndex        =   22
         Top             =   2504
         Width           =   1290
      End
      Begin VB.CommandButton cmdPicCollect 
         Caption         =   "�ɼ�"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�ļ�"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��"
         Height          =   360
         Left            =   8070
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   2504
         Width           =   360
      End
      Begin VB.CommandButton cmdBirthLocation 
         Caption         =   "��"
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
      Begin VB.TextBox txt�໤�� 
         Height          =   360
         IMEMode         =   2  'OFF
         Left            =   8670
         MaxLength       =   20
         TabIndex        =   30
         Top             =   3720
         Width           =   2775
      End
      Begin VB.TextBox txt������Ӧ 
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
      Begin VB.TextBox txt��֤���� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5790
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   1277
         Width           =   2895
      End
      Begin VB.TextBox txt֧������ 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1170
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1277
         Width           =   2895
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   360
         Left            =   7080
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F3"
         Top             =   4110
         Width           =   375
      End
      Begin VB.ComboBox cbo���ʽ 
         Height          =   360
         ItemData        =   "frmPatiInfo.frx":0E42
         Left            =   8670
         List            =   "frmPatiInfo.frx":0E44
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3322
         Width           =   2775
      End
      Begin VB.ComboBox cbo�ѱ� 
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
      Begin VB.ComboBox cbo���䵥λ 
         Height          =   360
         Left            =   7920
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   465
         Width           =   780
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����"
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
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.TextBox txt����� 
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
      Begin VB.TextBox txt���� 
         Height          =   360
         IMEMode         =   2  'OFF
         Left            =   7185
         MaxLength       =   5
         TabIndex        =   9
         Top             =   465
         Width           =   690
      End
      Begin VB.ComboBox cbo�Ա� 
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
      Begin VB.ComboBox cbo���� 
         Height          =   360
         Left            =   4710
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2913
         Width           =   2775
      End
      Begin VB.ComboBox cbo���� 
         Height          =   360
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2913
         Width           =   2775
      End
      Begin VB.ComboBox cbo���� 
         Height          =   360
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3322
         Width           =   2775
      End
      Begin VB.ComboBox cboְҵ 
         Height          =   360
         Left            =   8670
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2913
         Width           =   2775
      End
      Begin VB.TextBox txt���֤�� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1170
         MaxLength       =   18
         TabIndex        =   11
         Top             =   868
         Width           =   2895
      End
      Begin VB.TextBox txt��ͥ�绰 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5790
         MaxLength       =   20
         TabIndex        =   12
         Top             =   855
         Width           =   2895
      End
      Begin VB.TextBox txt��ͥ�ʱ� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   10140
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2095
         Width           =   1290
      End
      Begin VB.CommandButton cmd��ͥ��ַ 
         Caption         =   "��"
         Height          =   360
         Left            =   8070
         TabIndex        =   0
         ToolTipText     =   "�ȼ�F3"
         Top             =   2085
         Width           =   360
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
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
         ToolTipText     =   "�ȼ�:F3"
         Top             =   6540
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txt���� 
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
         Height          =   1215
         Left            =   30
         TabIndex        =   42
         ToolTipText     =   "F4:�޸�,F3:ѡ��"
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
         FormatString    =   "<����ҩ��                            "
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin MSMask.MaskEdBox txt����ʱ�� 
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
      Begin MSMask.MaskEdBox txt�������� 
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
      Begin VB.TextBox txt���� 
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
      Begin VB.ComboBox cbo��ͥ��ַ 
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
         Caption         =   "�ֻ���"
         Height          =   240
         Left            =   7920
         TabIndex        =   123
         Top             =   4170
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ��"
         Height          =   240
         Left            =   4875
         TabIndex        =   120
         Top             =   510
         Width           =   480
      End
      Begin VB.Label lbl���ڵ�ַ�ʱ� 
         Alignment       =   1  'Right Justify
         Caption         =   "�����ʱ�"
         Height          =   240
         Left            =   8595
         TabIndex        =   109
         Top             =   2564
         Width           =   1515
      End
      Begin VB.Label lblRegLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڵ�ַ"
         Height          =   240
         Left            =   150
         TabIndex        =   101
         Top             =   2564
         Width           =   960
      End
      Begin VB.Label lblBirthLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ص�"
         Height          =   240
         Left            =   150
         TabIndex        =   100
         Top             =   3780
         Width           =   960
      End
      Begin VB.Label lbl�໤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  �໤��"
         Height          =   240
         Left            =   7680
         TabIndex        =   98
         Top             =   3780
         Width           =   960
      End
      Begin VB.Label lbl��֤���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��֤����"
         Height          =   240
         Left            =   4800
         TabIndex        =   82
         Top             =   1335
         Width           =   960
      End
      Begin VB.Label lbl֧������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "֧������"
         Height          =   240
         Left            =   150
         TabIndex        =   81
         Top             =   1337
         Width           =   960
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   630
         TabIndex        =   54
         Top             =   4170
         Width           =   480
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   240
         Left            =   4170
         TabIndex        =   80
         Top             =   3375
         Width           =   480
      End
      Begin VB.Label lbl���ʽ 
         BackStyle       =   0  'Transparent
         Caption         =   "���ʽ"
         Height          =   300
         Left            =   7680
         TabIndex        =   60
         Top             =   3352
         Width           =   960
      End
      Begin VB.Label lblPatiMCNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��֤ҽ����"
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
         Caption         =   "ҽ����"
         Height          =   240
         Index           =   0
         Left            =   390
         TabIndex        =   78
         Top             =   1746
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
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
      Begin VB.Label lbl����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   240
         Left            =   5040
         TabIndex        =   72
         Top             =   105
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   630
         TabIndex        =   71
         Top             =   110
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   240
         Left            =   630
         TabIndex        =   70
         Top             =   519
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   6660
         TabIndex        =   69
         Top             =   525
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         Height          =   240
         Left            =   150
         TabIndex        =   68
         Top             =   3382
         Width           =   960
      End
      Begin VB.Label lblְҵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ"
         Height          =   240
         Left            =   8160
         TabIndex        =   67
         Top             =   2970
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   660
         TabIndex        =   66
         Top             =   2973
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   4170
         TabIndex        =   65
         Top             =   2970
         Width           =   480
      End
      Begin VB.Label lbl���֤ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   240
         Left            =   150
         TabIndex        =   64
         Top             =   930
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��סַ"
         Height          =   240
         Left            =   390
         TabIndex        =   63
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label lbl��ͥ�绰 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�绰"
         Height          =   240
         Left            =   5280
         TabIndex        =   62
         Top             =   915
         Width           =   480
      End
      Begin VB.Label lbl��ͥ�ʱ� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��סַ�ʱ�"
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
         Caption         =   "��"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "֤����Ϣ"
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
         Caption         =   "Ѫ��"
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
         Caption         =   "ҽѧ��ʾ"
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
         Caption         =   "����ҽѧ��ʾ"
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
         Caption         =   "��ϵ����Ϣ"
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
         Caption         =   "������Ϣ"
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
         Caption         =   "�������"
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
Option Explicit 'Ҫ���������
Public mbytFun As Byte  '0-�༭��鿴������Ϣ,1-���￨����,��󶨾��￨
Public mstrCard As String '��:����,���ɹ�:����,�󶨾��￨ʱ������
Public mblnChange As Boolean
Public mblnInRange As Boolean 'һ��ͨģʽˢ��ʱ,�����Ƿ����������η�Χ��
Public mrs��ͥ��ַ As ADODB.Recordset  '�����ͥ��ַ,��ʼʱ��ȡ������
Public mlngOutModeMC As Long '����ҽ�����õ����ʽҽ������
Public mrsBaseDict As ADODB.Recordset '����,����,����״��,ְҵ
Public mintNOLength As Integer '����ų���
Public mbln���� As Boolean '�����:56599
Public mstrPrivs As String
Public mlngModul As Long
Public mstr���� As String 'ԭ����
Public mstr�Ա� As String 'ԭ�Ա�
Public mstr���� As String 'ԭ����
Public mstr���䵥λ As String
Public mstr�������� As String
Public mstr����ʱ�� As String
Public mstr���֤�� As String
Public mstrFirstCode As String '��һ��֤�����͵ı���
Private mbln������Ϣ���� As Boolean '�Ƿ�����������˻�����Ϣ
Private mblnCancel As Boolean
Private mlng�ſ�����ID As Long
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private mblnStructAdress As Boolean  '���˵�ַ�ṹ��¼��
Private mblnShowTown As Boolean      '�����ַ�ṹ��¼��
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private mfrmMain As Object
Private mDateSys As Date
Private mblnCheckNOValidity As Boolean
Private mobjKeyboard As Object
Private mblnɨ�����֤ As Boolean '�жϲ�����Ϣ�Ƿ���ͨ��ɨ�����֤�õ�
Private mbln��ֹ�������� As Boolean
Public mrsPatiInfo As Recordset '��ǰHis������Ϣ
Public mlng����ID As Long '��ǰ����ID   :        '35233
Private mbln��ͥ��ַ����    As Boolean      '��ͥ��ַ�����Ƿ�����
Public Event PatiMerged(����ID As Long)     '���˺ϲ��������¼�
Private mblnɨ�����֤ǩԼ As Boolean
Public mbln�໤��¼�� As Boolean
Public mlng�໤������ As Long
Private mintDefaultBlood As Integer 'Ĭ��Ѫ�����
Private Enum mPageIndex
    ���� = 1
    �������� = 2
    ������Ϣ = 3
End Enum
Private mdicҽ�ƿ����� As New Dictionary '�����56599
Private Const C_InoculateHeader = "��������,4,2400,1;��������,4,2400,1;��������,4,2400,1;��������,4,2400,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_LinkManColumHeader = "����,4,1200,1;��ϵ,4,2400,1;���֤��,4,2400,1;�绰,4,1200,1;������Ϣ,4,2400,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_OtherInfoColumHeader = "��Ϣ��,4,2400,1;��Ϣֵ,4,2400,1;��Ϣ��,4,2400,1;��Ϣֵ,4,2400,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_CertificateHeader = "֤������,4,2400,1;֤������,4,2400,1;֤������,4,2400,1;֤������,4,2400,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
'Private Const C_Ѫ�� = "A��,B��,O��,AB��,����"
Private Const C_BH = "��,��,����,δ��"
Public Event ReturnVisitClick()     '������︴ѡ��ı��Ӧ�ķѱ���ʾ
Public mlngPlugInHwnd As Long
Public mblnPlugin As Boolean '�彨�Ƿ񴴽��ɹ�
Public mrsEMPIOut As ADODB.Recordset 'EMPI���ص�����
Public mstrPlugChange As String

'74430,Ƚ����,2014-7-7,�Һ��еĲ�����Ϣ�༭�������ṩ�ɼ���Ƭ����
Private mstr�ɼ�ͼƬ As String '�ɼ�ͼƬ���ر���·��
Public mlngͼ����� As Long 'ָ����ǰ�Բ���ͼ�����������(1-�ļ� 2-�ɼ� 3-��� 4-���֤��ȡ)
Private mstrIDImageFile As String
Public mblnSavePati As Boolean '������Ƭ��Ϣ�򸽼���Ϣ�Ƿ��ѱ���
Public mobjPubPatient As Object
Public mblnNewPatient As Boolean
Private mblnNameChange As Boolean
Public Event ���ʽClick(index As Long)     '������ʽ
Public mstrPriceGrade As String

Private mobjProPati As Collection '�ڹҺ�ǰ�����没����Ϣ��
Private mblnGetBirth As Boolean '�ж��Ƿ�����ͨ�������������

Private Sub cbo���ʽ_Click()
    RaiseEvent ���ʽClick(cbo���ʽ.ListIndex)
End Sub

Private Sub cbo����_Change()
    mstrPlugChange = mstrPlugChange & ",����"
End Sub

Private Sub cbo����_Change()
    mstrPlugChange = mstrPlugChange & ",����״��"
End Sub

Private Sub cbo��ͥ��ַ_Change()
    If Not mblnStructAdress Then mstrPlugChange = mstrPlugChange & ",��סַ"
End Sub

Private Sub cbo��ͥ��ַ_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub
Private Sub cbo��ͥ��ַ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub cbo��ͥ��ַ_KeyDown(KeyCode As Integer, Shift As Integer)
    '�˹��̴������������ݵ�ɾ��,�Լ���������ʱ���������б�
    '�����б���ʱ,�������ɾ����ʱ,��ɾ�������¼
    
    Dim str��ͥ��ַ As String
    
    If KeyCode = vbKeyDelete Then
        str��ͥ��ַ = cbo��ͥ��ַ.Text
        If Not mrs��ͥ��ַ Is Nothing And mbln��ͥ��ַ���� Then
            If mrs��ͥ��ַ.State = 1 And str��ͥ��ַ <> "" Then
                If cbo��ͥ��ַ.SelText = str��ͥ��ַ And SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = True Then
                    mrs��ͥ��ַ.Filter = "����='" & str��ͥ��ַ & "'"
                    If Not mrs��ͥ��ַ.EOF Then
                        mrs��ͥ��ַ.Delete adAffectCurrent
                        mrs��ͥ��ַ.Update
                    End If
                End If
            End If
        End If
    ElseIf KeyCode = vbKeyDown And cbo��ͥ��ַ.Text <> "" Then
        If SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, True, 0&)
    ElseIf KeyCode = vbKeyF3 Then
        cmd��ͥ��ַ.SetFocus
        Call cmd��ͥ��ַ_Click
    End If
End Sub

Private Sub cbo��ͥ��ַ_KeyUp(KeyCode As Integer, Shift As Integer)
    '��ʱtext���ѽ����������Ϣ
    '���¼�����ɾ�����˸��,ɾ������������Ŀ��,�����б�����������Ӧ������ɸѡ
    '���ȫ�����ֶ�ɾ����,����������б�����
        
    Dim str��ͥ��ַ As String, i As Long
    Dim lngλ�� As Long
    
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If mrs��ͥ��ַ Is Nothing Or mbln��ͥ��ַ���� = False Then Exit Sub
        
        str��ͥ��ַ = cbo��ͥ��ַ.Text                      '��ʱ,���ѡ���˲�������,��ѡ��������Ѿ���ɾ��
        lngλ�� = cbo��ͥ��ַ.SelStart
        
        If mrs��ͥ��ַ.State = 1 And Len(str��ͥ��ַ) > 1 Then
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(str��ͥ��ַ, 1))) > 0 Then
                mrs��ͥ��ַ.Filter = "���� like '" & gstrLike & UCase(str��ͥ��ַ) & "*'"
            Else
                mrs��ͥ��ַ.Filter = "���� Like '" & gstrLike & str��ͥ��ַ & "*'"
            End If
            
            If Not mrs��ͥ��ַ.EOF Then
                
                If mrs��ͥ��ַ.RecordCount <> cbo��ͥ��ַ.ListCount Then
                    Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_RESETCONTENT, 0, 0)
                    mrs��ͥ��ַ.Sort = "���� Desc,����"
                    For i = 1 To mrs��ͥ��ַ.RecordCount
                        AddComboItem cbo��ͥ��ַ.Hwnd, CB_ADDSTRING, 0, mrs��ͥ��ַ!����
                        mrs��ͥ��ַ.MoveNext
                    Next
                    If SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, True, 0&)
                                        
                    cbo��ͥ��ַ.Text = str��ͥ��ַ
                    cbo��ͥ��ַ.SelStart = lngλ��
                End If
            Else
                Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, False, 0&)
            End If
        ElseIf str��ͥ��ַ = "" Then
            cbo��ͥ��ַ.Clear
            Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, False, 0&)
        End If
    End If
End Sub

Private Sub cbo��ͥ��ַ_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim str���� As String
    Dim str��ͥ��ַ As String
    Dim lng�м������ As Long
    
    If (mrs��ͥ��ַ Is Nothing Or mbln��ͥ��ַ���� = False) And KeyAscii <> 13 Then Exit Sub
    
    '�ñ��ػ���ƥ������
    If KeyAscii <> 13 And KeyAscii <> vbKeyF4 And KeyAscii <> vbKeyEscape And _
        KeyAscii <> vbKeyBack And KeyAscii <> 26 And KeyAscii <> 3 And KeyAscii <> 22 Then   '26��ʾctrl+z,3-ctrl+c,22-ctrl+v
            
        If mrs��ͥ��ַ.State = 0 Or cbo��ͥ��ַ.Text = "" Then  '���һ����ʱ��ƥ��
            Exit Sub
        End If
       
        'ѡ���м䲿���ı�����������
        If cbo��ͥ��ַ.SelText <> "" And (cbo��ͥ��ַ.SelStart + cbo��ͥ��ַ.SelLength) <> Len(cbo��ͥ��ַ.Text) Then
            lng�м������ = cbo��ͥ��ַ.SelStart + 1
            cbo��ͥ��ַ.Text = Mid(cbo��ͥ��ַ.Text, 1, cbo��ͥ��ַ.SelStart) & Chr(KeyAscii) & Mid(cbo��ͥ��ַ.Text, cbo��ͥ��ַ.SelStart + cbo��ͥ��ַ.SelLength + 1)
            cbo��ͥ��ַ.SelText = ""
            str��ͥ��ַ = cbo��ͥ��ַ.Text
        Else
            '�������β��,�����м�ʱ,�������ѡ��
            If cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text) Or (cbo��ͥ��ַ.SelStart + cbo��ͥ��ַ.SelLength) = Len(cbo��ͥ��ַ.Text) Then
                str��ͥ��ַ = Mid(cbo��ͥ��ַ.Text, 1, cbo��ͥ��ַ.SelStart) & Chr(KeyAscii)
            Else
                str��ͥ��ַ = Mid(cbo��ͥ��ַ.Text, 1, cbo��ͥ��ַ.SelStart) & Chr(KeyAscii) & Mid(cbo��ͥ��ַ.Text, cbo��ͥ��ַ.SelStart + 1)
                lng�м������ = cbo��ͥ��ַ.SelStart + 1
            End If
        End If
        
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(str��ͥ��ַ, 1))) > 0 Then
            mrs��ͥ��ַ.Filter = "���� like '" & gstrLike & UCase(str��ͥ��ַ) & "*'"
        Else
            mrs��ͥ��ַ.Filter = "���� Like '" & gstrLike & str��ͥ��ַ & "*'"
        End If
        
        If Not mrs��ͥ��ַ.EOF Then
            If mrs��ͥ��ַ.RecordCount <> cbo��ͥ��ַ.ListCount Then
                Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_RESETCONTENT, 0, 0)
                mrs��ͥ��ַ.Sort = "���� Desc,����"
                For i = 1 To mrs��ͥ��ַ.RecordCount
                    AddComboItem cbo��ͥ��ַ.Hwnd, CB_ADDSTRING, 0, mrs��ͥ��ַ!����
                    mrs��ͥ��ַ.MoveNext
                Next
                If SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, True, 0&)
            End If
            
            i = KeyAscii    '���������ж��Ƿ��ǰ��˸�ɾ����
            KeyAscii = 0
            cbo��ͥ��ַ.Text = str��ͥ��ַ
            cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text)

            mrs��ͥ��ַ.MoveFirst   '�����������ļ���,��ͬ��ȡ��һ�������
            If mrs��ͥ��ַ!���� = str��ͥ��ַ And i <> vbKeyBack Then
                mrs��ͥ��ַ.MoveNext
            End If
            If Not mrs��ͥ��ַ.EOF Then
                If InStr(1, mrs��ͥ��ַ!����, str��ͥ��ַ) > 0 Or mrs��ͥ��ַ!���� = UCase(str��ͥ��ַ) Then    '�������������������ݵ�һ����,��ѡ�л����������
                    i = Len(cbo��ͥ��ַ.Text)
                    cbo��ͥ��ַ.Text = mrs��ͥ��ַ!����
                    cbo��ͥ��ַ.SelStart = i
                    cbo��ͥ��ַ.SelLength = Len(cbo��ͥ��ַ.Text) - cbo��ͥ��ַ.SelStart
                    
                    If mrs��ͥ��ַ.RecordCount = 1 Then Exit Sub
                End If
            End If
            
        'û���ҵ�ƥ��Ļ�������ʱ,����������б�����
        Else
            Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_RESETCONTENT, 0, 0)
            If SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, False, 0&)
            KeyAscii = 0
            cbo��ͥ��ַ.Text = str��ͥ��ַ
            cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text)
        End If
        
        If lng�м������ > 0 Then cbo��ͥ��ַ.SelStart = lng�м������: cbo��ͥ��ַ.SelText = ""
        
    ElseIf KeyAscii = 13 Then
        'a.��û��ѡ���κ�����,����������Ϊ��,���Ϊ��ĩ��ʱ,ȷ������,��������Ϣ�����ػ���
        Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, False, 0&)
        
        If cbo��ͥ��ַ.Text = "" Then
            If gbln��ͥ��ַ And txtPatient.Text <> "" Then
                Exit Sub
            Else
                Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            End If
        End If
        
        '�����б���ʱ���س�,��λ��ĩβ
        If cbo��ͥ��ַ.SelText = cbo��ͥ��ַ.Text Then cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text): Exit Sub
        
        If mrs��ͥ��ַ.State = 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If zlCommFun.ActualLen(cbo��ͥ��ַ.Text) > 100 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
       
        'a.������״̬�°��س�,û��ѡ���ı�
        If cbo��ͥ��ַ.SelText = "" Then
            str��ͥ��ַ = cbo��ͥ��ַ.Text
            mrs��ͥ��ַ.Filter = "����='" & str��ͥ��ַ & "'"
            If mrs��ͥ��ַ.EOF Then
                str���� = Mid(zlCommFun.zlGetSymbol(str��ͥ��ַ), 1, 10)
                If str���� <> UCase(str��ͥ��ַ) Then
                    With mrs��ͥ��ַ
                        .AddNew
                        !��� = "�û�"
                        !���� = str��ͥ��ַ
                        !���� = str����
                        !���� = 1
                        .Update                 '�ڴ���Unload��save
                    End With
                End If
            Else
                mrs��ͥ��ַ!���� = mrs��ͥ��ַ!���� + 1
                mrs��ͥ��ַ.Update
                
                If zlCommFun.IsCharAlpha(str��ͥ��ַ) Then
                    If mrs��ͥ��ַ.RecordCount = 1 Then
                        cbo��ͥ��ַ.Text = mrs��ͥ��ַ!����
                    Else
                        Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, True, 0&)
                        Exit Sub
                    End If
                End If
            End If
            
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo��ϵ�˹�ϵ_Change()
    mstrPlugChange = mstrPlugChange & ",��ϵ�˹�ϵ"
End Sub

Private Sub cbo��ϵ�˹�ϵ_Click()
    With cbo��ϵ�˹�ϵ
        If .ListIndex = 8 And txt������ϵ.Visible = False Then
            .Width = 1200: txt������ϵ.Visible = True
        ElseIf .ListIndex <> 8 And txt������ϵ.Visible Then
            .Width = 2445: txt������ϵ.Visible = False
        ElseIf .ListIndex = -1 Then
            .Width = 2445
        End If
    End With
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("��ϵ") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("��ϵ")) = zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.Text)
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("������Ϣ")) = zlCommFun.GetNeedName(txt������ϵ.Text)
    End If
End Sub

Private Sub cbo����_Change()
    mstrPlugChange = mstrPlugChange & ",����"
End Sub

Private Sub cbo�Ա�_Change()
    mstrPlugChange = mstrPlugChange & ",�Ա�"
End Sub

Private Sub cboְҵ_Change()
    mstrPlugChange = mstrPlugChange & ",ְҵ"
End Sub

Private Sub chk����_Click()
    RaiseEvent ReturnVisitClick
End Sub

Private Sub cmdMedicalWarning_Click()
'�����:56599
    Dim rsTemp As Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txtMedicalWarning.Hwnd)
    strSQL = "" & _
    "       Select ���� as ID,����,���� From ҽѧ��ʾ Where ���� Not Like '����%'"
    Set rsTemp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "ҽѧ��ʾ", False, txtMedicalWarning.Text, "", False, False, False, vRect.Left, vRect.Top - 180, 500, blnCancel, False, True)
    If blnCancel Then Exit Sub
    If Not rsTemp Is Nothing Then
      While rsTemp.EOF = False
        strTemp = strTemp & "," & rsTemp!����
        rsTemp.MoveNext
      Wend
    End If
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    If strTemp <> "" Then txtMedicalWarning.Text = strTemp
End Sub
''Private Sub SetPatiBaseModiPropertyEanbled()
''   '---------------------------------------------------------------------------------------------------------------------------------------------
''    '����:���ò��˵Ļ�����Ϣ�ı༭����
''    '����:���˺�
''    '����:2013-11-04 11:59:46
''    '---------------------------------------------------------------------------------------------------------------------------------------------
''    Dim blnEnabled As Boolean
''    Dim lngColor As Long
''    On Error GoTo errHandle
''
''    blnEnabled = mbytFun <> 0 Or mlng����ID = 0
''
''    txtPatient.Enabled = blnEnabled
''    cbo�Ա�.Enabled = blnEnabled
''    txt����.Enabled = blnEnabled
''    cbo���䵥λ.Enabled = blnEnabled
''    txt��������.Enabled = blnEnabled
''
''    lngColor = IIf(blnEnabled, vbWhite, Me.BackColor)
''    txtPatient.BackColor = lngColor
''    cbo�Ա�.BackColor = lngColor
''    txt����.BackColor = lngColor
''    cbo���䵥λ.BackColor = lngColor
''    txt��������.BackColor = lngColor
''
''    Exit Sub
''errHandle:
''    If ErrCenter() = 1 Then
''        Resume
''    End If
'' End Sub
''

Private Sub cmdPicClear_Click()
    '�����:56599
    imgPatient.Picture = Nothing
    mlngͼ����� = 3
End Sub

Private Sub cmdPicCollect_Click()
    If mobjPubPatient Is Nothing Then Exit Sub
    If mobjPubPatient.PatiImageGatherer(Me, mstr�ɼ�ͼƬ) = False Then Exit Sub
    imgPatient.Picture = LoadPicture(mstr�ɼ�ͼƬ)
    mlngͼ����� = 2
End Sub

Private Sub cmdPicFile_Click()
    '�����:56599
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
    mlngͼ����� = 1
    Exit Sub
errHanl:
     
End Sub

Private Sub cmd����_Click()
    If zl_SelectAndNotAddItem(Me, txt����, "", "����", "����ѡ��", True, False) = False Then
        Exit Sub
    End If
End Sub

Private Sub lblICCard_Click()
    If txt����.Enabled = False Or txt����.Locked Then Exit Sub
    If gCurSendCard.bln���￨ And gCurSendCard.str������ <> "�������֤" Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txt����.Text = mobjICCard.Read_Card()
            If txt����.Text <> "" Then mfrmMain.mblnICCard = True
        End If
        Exit Sub
    End If
    '��ȡ�俨��Ϣ
    '���˺�
    If zlLoadInfor = False Then
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
        Exit Sub
    End If
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    txt����.Text = strCardNo
    If txt����.Text <> "" Then
        '�����:56599
        If strXmlCardInfor <> "" Then Call LoadPati(strXmlCardInfor)
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    Else
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind        As Long
    Dim lng����ID           As Long
    Dim strPassWord         As String
    Dim strErrMsg           As String
    Dim strTmp              As String
    Dim bln����ǩԼ         As Boolean
    Dim bln�Ƿ�ǩԼ         As Boolean
    Dim blnUserCancel       As Boolean
    Static blnIsPatient     As Boolean '��һ�δ����¼��Ŀؼ��Ƿ���,���������ı���
    '��������ˢ������䲡����Ϣ
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
'��ȡ���˱�ǩ
GetPatientTab:
        
        If mblnɨ�����֤ǩԼ Then mblnɨ�����֤ = True
        Set mrsPatiInfo = Nothing
        '��ȡ����ID
        blnUserCancel = False
        If gobjSquare.objSquareCard.zlGetPatiID("���֤", strID, False, lng����ID, strPassWord, strErrMsg, , , , InStr(gstrPrivs, ";�ϲ�������Ϣ;") > 0, , , blnUserCancel) = False Then lng����ID = 0
             '����Աȡ����ȡ�²���
        bln����ǩԼ = True
        mblnNewPatient = True
        
        If blnUserCancel = True And lng����ID = 0 Then Exit Sub
        
        If lng����ID = 0 Then
            txt���֤��.Text = strID
            txtPatient.Text = strName
            Call zlControl.CboLocate(cbo�Ա�, strSex)
            Call zlControl.CboLocate(cbo����, strNation)
            txt��������.Text = Format(datBirthDay, "yyyy-MM-dd")
            txt����ʱ��.Text = "00:00"
            cbo��ͥ��ַ.Text = IIf(Trim(cbo��ͥ��ַ.Text) = "", strAddress, cbo��ͥ��ַ.Text)
            txtRegLocation.Text = strAddress
            '89242:���ϴ�,2015/12/10,��ȡ���˵�ַ��Ϣ
            padd��ͥ��ַ.Value = IIf(Trim(padd��ͥ��ַ.Value) = "", strAddress, padd��ͥ��ַ.Value)
            padd���ڵ�ַ.Value = strAddress
            
            '74430,Ƚ����,2014-7-7,�Һ��еĲ�����Ϣ�༭�������ṩ�ɼ���Ƭ����
            Call LoadIDImage
        Else
            
            Set mrsPatiInfo = GetPatiByID("����ID", CStr(lng����ID))
            If mrsPatiInfo.EOF = False Then
                If (Nvl(mrsPatiInfo!����) <> Trim(strName) Or Nvl(mrsPatiInfo!�Ա�) <> strSex Or Format(Nvl(mrsPatiInfo!��������, "00:00:00"), "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd")) Then
                    bln����ǩԼ = False
                    mblnɨ�����֤ = False
                    txt֧������.Text = ""
                    txt֧������.Tag = ""
                    txt��֤����.Text = ""
                    txt��֤����.Tag = ""
                    If gCurSendCard.str������ = "�������֤" Then
                        MsgBox "���֤��Ϣ��HIS�в�����Ϣ��һ��,���ܽ���ǩԼ������", vbInformation, gstrSysName
                        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                    Else
                        '�Ƕ������֤��
                        zl���ز�����Ϣ mrsPatiInfo
                        If txtRegLocation.Text = "" Then
                            txtRegLocation.Text = strAddress
                            padd���ڵ�ַ.Value = strAddress
                        End If
                    End If
                    Set mrsPatiInfo = Nothing
                Else
                    bln����ǩԼ = True
                    mblnɨ�����֤ = True
                    zl���ز�����Ϣ mrsPatiInfo
                    If imgPatient.Picture = 0 Then
                        '74430,Ƚ����,2014-7-7,�Һ��еĲ�����Ϣ�༭�������ṩ�ɼ���Ƭ����
                        Call LoadIDImage
                    End If
                    If txtRegLocation.Text = "" Then
                        txtRegLocation.Text = strAddress
                        padd���ڵ�ַ.Value = strAddress
                    End If
                End If
                
                '�Ƿ�ɨ�����֤ǩԼ
                '�����Ҫ���ŵ����֤���Ƿ��Ѿ���ǩԼ
                If mblnɨ�����֤ǩԼ And bln����ǩԼ Then
                    bln�Ƿ�ǩԼ = �Ƿ��Ѿ�ǩԼ(strID)
                    If bln�Ƿ�ǩԼ Then
                        If gCurSendCard.str������ = "�������֤" Then
                            MsgBox "��ǰ�����Ѿ�����ǩԼ����,��������ٴ�ǩԼ��", vbInformation, gstrSysName
                            Set mrsPatiInfo = Nothing
                            txt֧������.Text = ""
                            txt֧������.Tag = ""
                            txt��֤����.Text = ""
                            txt��֤����.Tag = ""
                            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                        End If
                        bln����ǩԼ = False
                    End If
                    mblnɨ�����֤ = Not bln�Ƿ�ǩԼ
                End If
                
            End If
        End If
        If mblnɨ�����֤ And bln����ǩԼ Then txt֧������.Tag = Getҽ�ƿ����ID("�������֤")
        '�ӿ��ϻ�ȡ������Ϣ��,����EMPI
        Call zlQueryEMPIPatiInfo
        blnIsPatient = True
        SetCtrVisibleAndMove
        If gblnNewCardNoPop Then Call cmdOK_Click
        Exit Sub
    End If
    
    If Not ActiveControl Is txt���֤�� Then Exit Sub
    
    '���ز�����Ϣ��������ڣ�
    '�����:54197
    
    If txtPatient.Text = "" Or mrsPatiInfo Is Nothing Then GoTo GetPatientTab: Exit Sub  '���û�в�����Ϣ
   
    '����Ѵ��ڲ�����Ϣ����ˢ�����Ĳ������֤��һ��,����ִ��
    If Nvl(mrsPatiInfo!���֤��) <> strID Then '���֤�Ų�һ�� ,���ܹ�����
        Call MsgBox("��ǰ�������֤�������֤�ϵ���Ϣ��һ��,���ܼ���!", vbInformation, Me.Caption)
        Exit Sub
    End If
    
    '����Ѿ������˲�����Ϣ,
    
    '1.��ǰ����,û����д���֤�ŵ����
    '  --�������,�Ա�,�Լ���������  �����һ��,��ʾ�Ƿ����
    '2.��ǰ���������֤���˲�һ�µ����
    '
    '������Ϣ���
    
    If Nvl(mrsPatiInfo!����) <> Trim(strName) Or Nvl(mrsPatiInfo!�Ա�) <> strSex Or Format(txt��������.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
      
        If Nvl(mrsPatiInfo!����) <> Trim(strName) Then
             strErrMsg = strErrMsg & " ����:" & (mrsPatiInfo!����) & " ����(���֤):" & strName & vbCrLf
             strTmp = strTmp & "," & "����"
        End If
        If Nvl(mrsPatiInfo!�Ա�) <> strSex Then
             strErrMsg = strErrMsg & " �Ա�:" & Nvl(mrsPatiInfo!�Ա�) & " �Ա�(���֤):" & strSex & vbCrLf
             strTmp = strTmp & "," & "�Ա�"
        End If
        If Format(txt��������.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
             strErrMsg = strErrMsg & " ��������:" & Format(txt��������.Text, "yyyy-MM-dd") & " ��������(���֤):" & Format(datBirthDay, "yyyy-MM-dd") & vbCrLf
             strTmp = strTmp & "," & "��������"
        End If
        strTmp = Mid(strTmp, 2)
        strErrMsg = "��ǰ������Ϣ�����֤�ϵ�[" & strTmp & "]����Ϣ��һ��," & vbCrLf & strErrMsg
        strErrMsg = strErrMsg & "�Ƿ������֤�ϵ�[" & strTmp & "]��Ϣ�滻��ǰ���˵���Ӧ��Ϣ?" & vbCrLf
        If MsgBox(strErrMsg, vbYesNo + vbDefaultButton2 + vbQuestion, Me.Caption) = vbYes Then
             txtPatient.Text = strName
             txt���֤��.Text = strID
             Call zlControl.CboLocate(cbo�Ա�, strSex)
             txt��������.Text = Format(datBirthDay, "yyyy-MM-dd")
        End If
    End If
    
    cbo���䵥λ.Tag = cbo���䵥λ.Text
    
    '�Ƿ�ɨ�����֤ǩԼ
    '�����Ҫ���ŵ����֤���Ƿ��Ѿ���ǩԼ
    If mblnɨ�����֤ǩԼ Then mblnɨ�����֤ = Not �Ƿ��Ѿ�ǩԼ(strID)
    If mblnɨ�����֤ Then txt֧������.Tag = Getҽ�ƿ����ID("�������֤")
    SetCtrVisibleAndMove
    If gblnNewCardNoPop Then Call cmdOK_Click
End Sub

Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 And cbo�ѱ�.ListIndex <> -1 Then Call zlCommFun.PressKey(vbKeyTab)
    
    If SendMessage(cbo�ѱ�.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo�ѱ�.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo�ѱ�.ListIndex = lngIdx
    If cbo�ѱ�.ListIndex = -1 And cbo�ѱ�.ListCount > 0 Then cbo�ѱ�.ListIndex = 0
End Sub

Private Sub cbo���ʽ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo���ʽ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo���ʽ.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo���ʽ.ListIndex = lngIdx
    If cbo���ʽ.ListIndex = -1 And cbo���ʽ.ListCount > 0 Then cbo���ʽ.ListIndex = 0
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    If cbo����.ListIndex = -1 And cbo����.ListCount > 0 Then cbo����.ListIndex = 0
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    If cbo����.ListIndex = -1 And cbo����.ListCount > 0 Then cbo����.ListIndex = 0
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    If cbo����.ListIndex = -1 And cbo����.ListCount > 0 Then cbo����.ListIndex = 0
End Sub

Private Sub cbo���䵥λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub


Private Sub cbo���䵥λ_LostFocus()
    Dim strBirth As String
    If cbo���䵥λ.Locked Then Exit Sub
    '������������
    '69026,Ƚ����,2014-8-8,�����������
    If cbo���䵥λ.Text <> cbo���䵥λ.Tag Then
        mblnChange = False
        If mblnGetBirth Then
            If mobjPubPatient.ReCalcBirthDay(Trim(txt����.Text) & cbo���䵥λ.Text, strBirth) Then
                txt��������.Text = Format(strBirth, "yyyy-mm-dd")
                txt����ʱ��.Text = Format(strBirth, "hh:mm")
            End If
        End If
        mblnChange = True
        cbo���䵥λ.Tag = cbo���䵥λ.Text
    End If
    
    If Trim(txt����.Text) <> "" Then
        If mobjPubPatient Is Nothing Then Exit Sub
        If mobjPubPatient.CheckPatiAge(Trim(txt����.Text) & cbo���䵥λ.Text) = False Then
            If txt����.Visible And txt����.Enabled And Not txt����.Locked Then
                txt����.SetFocus: Exit Sub
            End If
        End If
    End If
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo�Ա�.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo�Ա�.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo�Ա�.ListIndex = lngIdx
    If cbo�Ա�.ListIndex = -1 And cbo�Ա�.ListCount > 0 Then cbo�Ա�.ListIndex = 0
    
End Sub

Private Sub cboְҵ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboְҵ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cboְҵ.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cboְҵ.ListIndex = lngIdx
    If cboְҵ.ListIndex = -1 And cboְҵ.ListCount > 0 Then cboְҵ.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
   '��ص�������
    mblnCancel = True
    mstrPlugChange = ""
    If mbytFun = 0 And mlng����ID <> 0 Then
        '35233
'        If CheckValied = False Then Exit Sub
    
        Call CloseIDCard    '47007
        Me.Hide: Exit Sub
    Else
        '������֤���Ϸ���
        If IsCertificateCard(mlng����ID) = False Then Exit Sub
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
    '����:��������Ƿ���ȷ
    '����:���˺�
    '����:2011-01-07 18:13:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long, strSimilar As String, i As Integer, strMCAccount As String
    Dim strSQL As String, rsTmp As ADODB.Recordset, intQuery As Integer
    Dim blnPlugInCheck As Boolean, str����ʱ�� As String
    Dim strbirthday As String, strAge As String, strSex As String, strErrInfo As String, strInfo As String
    
    '82859:���ϴ�,2015/4/8,���˻�����Ϣ����
    If mblnCancel = False And mbytFun = 1 And mstrCard <> "" And gbln���ѽ����� Or (mblnCancel = False And mbytFun = 0 And mlng����ID <> 0) Then
        If mlng����ID > 0 And mbln������Ϣ���� And (mstr���� & mstr���䵥λ <> IIf(IsNumeric(txt����.Text), txt����.Text & cbo���䵥λ.Text, txt����.Text) Or mstr�Ա� <> NeedName(cbo�Ա�.Text) Or mstr���� <> txtPatient.Text Or _
            mstr���֤�� <> txt���֤��.Text Or mstr�������� <> txt��������.Text Or mstr����ʱ�� <> txt����ʱ��.Text) Then
            If MsgBox("���˻�����Ϣ�ѷ����ı䣬�Ƿ������", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                '��¼����ԭʼ��Ϣ
                txtPatient.Text = mstr����:  cbo�Ա�.ListIndex = cbo.FindIndex(cbo�Ա�, mstr�Ա�, True)
                txt����.Text = mstr����
                If mstr���䵥λ <> "" Then cbo���䵥λ.ListIndex = cbo.FindIndex(cbo���䵥λ, mstr���䵥λ, True): cbo���䵥λ.Visible = True: txt����.Width = 690
                If mstr�������� <> "" Then txt��������.Text = mstr��������
                If mstr����ʱ�� <> "" Then txt����ʱ��.Text = mstr����ʱ��
                txt���֤��.Text = mstr���֤��
                Exit Function
            Else
                '��¼�����µ���Ϣ
                mstr���� = txtPatient.Text: mstr�Ա� = NeedName(cbo�Ա�.Text)
                mstr���� = txt����.Text: mstr���䵥λ = NeedName(cbo���䵥λ.Text)
                mstr�������� = txt��������.Text: mstr����ʱ�� = txt����ʱ��.Text
                mstr���֤�� = txt���֤��.Text
            End If
        End If
    End If
    mblnCancel = False
    
    If txt�����.Text = "" And txt�����.Enabled Then
        MsgBox "�����벡�˵�����ţ�", vbInformation, gstrSysName
        If txt�����.Visible And txt�����.Enabled Then txt�����.SetFocus
        Exit Function
    End If
    
    If cbo�ѱ�.ListIndex = -1 Then
        MsgBox "��ѡ���˵ķѱ�", vbInformation, gstrSysName
        If cbo�ѱ�.Visible Then cbo�ѱ�.SetFocus
        Exit Function
    End If
    If mbytFun = 1 And mstrCard = "" And Trim(txt����.Text) <> "" Then
        If CheckPatiValid(Trim(txt����.Text)) = False Then
            If txt����.Visible Then txt����.SetFocus
            Exit Function
        End If
    End If
    If txtPatient.Text = "" Then
        MsgBox "�����벡�˵�������", vbInformation, gstrSysName
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Function
    End If
    
    If txtPatiMCNO(0).Text <> "" Or txtPatiMCNO(1).Text <> "" Then
        If txtPatiMCNO(0).Text <> txtPatiMCNO(1).Text Then
            MsgBox "����,���������ҽ���Ų�һ�£�", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
        If zlCommFun.ActualLen(txtPatiMCNO(0).Text) > txtPatiMCNO(0).MaxLength Then
            MsgBox "����,ҽ������󳤶Ȳ��ܳ���" & txtPatiMCNO(0).MaxLength & "���ַ���", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
    End If
    
    If txtMobile.Text <> "" And IsMobileNO(txtMobile.Text) = False Then
        MsgBox "������ֻ��Ÿ�ʽ����ȷ��������¼�룡", vbInformation, gstrSysName
        If txtMobile.Visible And txtMobile.Enabled Then txtMobile.SetFocus
        Exit Function
    End If
    
    If CheckTextLength("����", txtPatient) = False Then Exit Function
    If CheckTextLength("�����ص�", txtBirthLocation) = False Then Exit Function
    If CheckTextLength("����", txt����) = False Then Exit Function
    '89242:���ϴ�,2015/12/10,��ַ��Ϣ���
    If mblnStructAdress Then
        If Not CheckStructAddr(padd��ͥ��ַ, padd��ͥ��ַ.MaxLength) Then Exit Function
        If Not CheckStructAddr(padd���ڵ�ַ, padd���ڵ�ַ.MaxLength) Then Exit Function
    Else
        If zlCommFun.ActualLen(cbo��ͥ��ַ.Text) > glngMax��ͥ��ַ Then
            MsgBox "��סַ���������ֻ��������" & glngMax��ͥ��ַ & "���ַ���" & glngMax��ͥ��ַ \ 2 & "�����֣�����!", vbInformation, gstrSysName
            cbo��ͥ��ַ.SetFocus: Exit Function
        End If
        If CheckTextLength("���ڵ�ַ", txtRegLocation) = False Then Exit Function
    End If
    If CheckTextLength("��סַ�ʱ�", txt��ͥ�ʱ�) = False Then Exit Function
    If CheckTextLength("�����ʱ�", txt���ڵ�ַ�ʱ�) = False Then Exit Function
    If CheckTextLength("�����ص�", txtBirthLocation) = False Then Exit Function
    '83062
    For i = 1 To msh����.Rows - 1
        If zlCommFun.ActualLen(msh����.TextMatrix(i, 1)) > 100 Then
            MsgBox "���˹���ҩ�ﷴӦ���������ֻ��������100���ַ���50�����֣����飡", vbInformation, gstrSysName
            If msh����.Enabled And msh����.Visible Then msh����.SetFocus
            Exit Function
        End If
        If zlCommFun.ActualLen(msh����.TextMatrix(i, 0)) > 60 Then
            MsgBox "���˹���ҩ���������������ֻ��������60���ַ���30�����֣����飡", vbInformation, gstrSysName
            If msh����.Enabled And msh����.Visible Then msh����.SetFocus
            Exit Function
        End If
    Next i
    '69026,Ƚ����,2014-8-11,������Ч�Լ��
    '76703,Ƚ����,2014-8-15
    
    If mbln��ֹ�������� Then
        '��ֹ������������,����Ƿ�¼���������
        If txt��������.Enabled And IsDate(txt��������.Text) = False And Not (gblnAutoAddName And txtPatient.Text = "�²���") Then
            MsgBox "�������벡�˳������ڣ�", vbInformation, gstrSysName
            txt��������.SetFocus: Exit Function
        End If
        If mobjPubPatient Is Nothing Then Exit Function
        If mobjPubPatient.CheckPatiAge(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""), _
                IIf(txt��������.Text = "____-__-__", "", txt��������.Text) & _
                IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text)) = False Then
            If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus
            Exit Function
        End If
    End If
    
    If txt����.Enabled And txt����.Visible Then
        If mobjPubPatient Is Nothing Then Exit Function
        If mobjPubPatient.CheckPatiAge(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""), _
                IIf(txt��������.Text = "____-__-__", "", txt��������.Text) & _
                IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text)) = False Then
            txt����.SetFocus:  Exit Function
        End If
    End If
    
    If IsDate(zlCommFun.GetIDCardDate(txt���֤��.Text)) Then
        If Format(zlCommFun.GetIDCardDate(txt���֤��.Text), "yyyy-mm-dd") <> Format(txt��������.Text, "yyyy-mm-dd") Then
            intQuery = MsgBox("��������֤��������ĳ������ڲ�һ�£�ʹ�����֤�Ż�ȡ�ĳ���������", vbQuestion + vbYesNoCancel, gstrSysName)
            If intQuery = 6 Then
                txt��������.Text = zlCommFun.GetIDCardDate(txt���֤��.Text)
            ElseIf intQuery = 2 Then
                CheckValied = False
                Exit Function
            End If
        End If
    End If
    
    If IsDate(txt��������.Text) Then
        '76669�����ϴ�,2014-8-15,������������ڼ��
        str����ʱ�� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        If CDate(str����ʱ��) > zlDatabase.Currentdate Then
            If MsgBox("����ʱ�䣺" & str����ʱ�� & " �����˵�ǰϵͳʱ�䡣" & _
                vbCrLf & vbCrLf & "���������������ڵ���ȷ�� ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus
                Exit Function
            End If
        End If
        If mbln�໤��¼�� And Trim(txt�໤��.Text) = "" Then
            '61945 �໤��¼�� ���
            strSQL = "Select Floor(Months_Between(Sysdate, [1]) / 12) as ���� From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(txt��������.Text))
            If Not rsTmp Is Nothing Then
                If Val(Nvl(rsTmp!����)) <= mlng�໤������ And mlng�໤������ <> 0 Then
                    MsgBox "������[" & mlng�໤������ & "��]�±���¼��໤��,����!"
                    Set rsTmp = Nothing
                    If txt�໤��.Enabled And txt�໤��.Visible Then txt�໤��.SetFocus
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
    
    '104238:���ϴ���2017/2/15����鿨���Ƿ����㷢����������
    If txt����.Text <> "" And Len(txt����.Text) <> gCurSendCard.lng���ų��� And Not gCurSendCard.bln�ϸ���� Then
        Select Case gCurSendCard.byt��������
            Case 0
                MsgBox "����Ŀ���С��" & gCurSendCard.str������ & "�趨�Ŀ��ų��ȣ����������룡", vbExclamation, gstrSysName
                If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                    Exit Function
            Case 2
                If MsgBox("����Ŀ���С��" & gCurSendCard.str������ & "�趨�Ŀ��ų��ȣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                    Exit Function
                End If
        End Select
    End If
    
     '��ֱ����ȡ���������Ϊ���۵�,ͬʱ��������    '������Һŷ���һ����ȡ,�ڹҺű���ʱ�Ž�������
    '�󶨿�ģʽ����ʱ,�˴�������
    If txt����.Text <> txt��֤.Text And txt��֤.Enabled And txt��֤.Visible Then    '
        MsgBox "������������벻һ��,���������룡", vbInformation, gstrSysName
        txt����.Text = "": txt��֤.Text = ""
        txt����.SetFocus: Exit Function
    End If
    If txt����.Text <> "" And Trim(txt����.Text) <> "" And txt����.Visible Then
        Select Case gCurSendCard.int���볤������
        Case 0
        Case 1
            If Len(txt����.Text) <> gCurSendCard.int���볤�� Then
                MsgBox "ע��:" & vbCrLf & "�����������" & gCurSendCard.int���볤�� & "λ", vbOKOnly + vbInformation
                If txt����.Enabled Then txt����.SetFocus
                Exit Function
             End If
        Case Else
            If Len(txt����.Text) < Abs(gCurSendCard.int���볤������) Then
                MsgBox "ע��:" & vbCrLf & "�����������" & Abs(gCurSendCard.int���볤������) & "λ����.", vbOKOnly + vbInformation
                If txt����.Enabled Then txt����.SetFocus
                Exit Function
             End If
        End Select
    End If
  
    '81103,Ƚ����,2014-12-26,¼�����֤�ź�,�������ڡ����䡢�Ա��ͬ���������͵���
    If Trim(txt���֤��.Text) <> "" And Not mobjPubPatient Is Nothing Then
        'CheckPatiIdcard(ByVal strIdcard As String, Optional strBirthday As String, _
        '    Optional strAge As String, Optional strSex As String, Optional strErrInfo As String) As Boolean
        '���ܣ����֤����Ϸ���У��
        '��Σ�strIdCard ���֤����
        '���Σ�strBirthday  ��������TrueΪ��������
        '         strAge ��������TrueΪ����
        '         strSex ��������TrueΪ�Ա�
        '         strErrInfo ��������FalseΪ������Ϣ
        '���أ�True/False  ���֤�Ϸ�����True(�ɴ�strBirthday��strSex��ȡ�������ں��Ա�)��
        '       ���򷵻�False(�ɴ�strErrInfo��ȡ��ϸ������Ϣ)
        If mobjPubPatient.CheckPatiIdcard(Trim(txt���֤��.Text), strbirthday, strAge, strSex, strErrInfo) Then
            If strSex <> NeedName(cbo�Ա�.Text) Then strInfo = "�Ա�"
            If strAge <> Trim(txt����.Text) & cbo���䵥λ Then strInfo = strInfo & IIf(strInfo = "", "����", "������")
            If Format(strbirthday, "yyyy-mm-dd") <> txt��������.Text Then strInfo = strInfo & IIf(strInfo = "", "��������", "����������")
            
            If strInfo <> "" Then
                If MsgBox("�����" & strInfo & "�����֤�ŵ�" & strInfo & "��һ�£�" & _
                        "���������֤���޸�" & strInfo & "���Ƿ������", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                    Call zlControl.CboLocate(cbo�Ա�, strSex)
                    txt��������.Text = Format(strbirthday, "yyyy-mm-dd")
                    mstr�Ա� = NeedName(cbo�Ա�.Text)
                    mstr���� = txt����.Text: mstr���䵥λ = NeedName(cbo���䵥λ.Text)
                    mstr�������� = txt��������.Text: mstr����ʱ�� = txt����ʱ��.Text
                    mstr���֤�� = txt���֤��.Text
                Else
                    Exit Function
                End If
            End If
        Else
            MsgBox strErrInfo, vbInformation, gstrSysName
            If txt���֤��.Enabled And txt���֤��.Visible Then txt���֤��.SetFocus
            Exit Function
        End If
    End If
        
     '������Ʋ�����Ϣ(����֮ǰ���,����������ظ���Ϣ������)
    If Trim(txt���֤��.Text) <> "" And cmdOK.Caption Like "ȷ��*" And mlng����ID = 0 Then
        strSimilar = SimilarIDs(Trim(txt���֤��.Text))
        If strSimilar <> "" Then
            i = UBound(Split(strSimilar, "|")) + 1
            strSimilar = Replace(strSimilar, "|", vbCrLf)
            If i > 20 Then strSimilar = Mid(strSimilar, 1, 200) & "..."
            
            If MsgBox("�����еĲ�����Ϣ�з��� " & i & " ����Ϣ���ƵĲ���(���֤����ͬ): " & vbCrLf & vbCrLf & _
                strSimilar & vbCrLf & vbCrLf & "ȷʵҪ�Ǽ�Ϊ�²�����", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    '�����:53408
    If IIf(zlDatabase.GetPara("ɨ�����֤ǩԼ", glngSys, mlngModul) = "1", 1, 0) = 0 And ((gCurSendCard.str������ = "�������֤" And Trim(txt����.Text) <> "") Or Trim(txt֧������.Text) <> "") Then
         MsgBox "��û��Ȩ�޽���ǩԼ����,�뵽�������������á�ɨ�����֤ǩԼ����", vbOKOnly + vbInformation, gstrSysName
         txt����.Text = ""
         txt����.Text = ""
         txt��֤.Text = ""
         If txt����.Visible = True And txt����.Enabled = True Then txt����.SetFocus
         Exit Function
    End If
    
    If Trim(txt֧������.Text) <> "" And Trim(txt���֤��.Text) <> "" Then
           If �Ƿ��Ѿ�ǩԼ(txt���֤��.Text) Then
                 MsgBox "���֤����Ϊ:" & txt���֤��.Text & "�Ѿ�ǩԼ�����ظ�ǩԼ��", vbOKOnly + vbInformation, gstrSysName
                 txt֧������.Text = ""
                 If txt֧������.Visible = True And txt֧������.Enabled = True Then txt֧������.SetFocus
                 Exit Function
           End If
    End If
    
    If mblnɨ�����֤ = False And gCurSendCard.str������ = "�������֤" And txt����.Text <> "" Then
            MsgBox "�����ֻ֤����ˢ���ķ�ʽ���У��������ֶ��������֤���а�!", vbOKOnly + vbInformation, gstrSysName
            txt����.Text = ""
            txt����.Text = ""
            txt��֤.Text = ""
            txt֧������.Text = ""
            txt��֤����.Text = ""
            '74894:���ϴ�,2014-07-08,ȡ���󶨶������֤�Ŀ�����Ϣ
            mstrCard = ""
            If txt����.Visible = True And txt����.Enabled = True Then txt����.SetFocus
            Exit Function
    End If
    
    If mblnɨ�����֤ = False And gCurSendCard.str������ <> "�������֤" And txt֧������.Text <> "" Then
            MsgBox "�����ֻ֤����ˢ���ķ�ʽ���У��������ֶ��������֤���а�!", vbOKOnly + vbInformation, gstrSysName
            txt���֤��.Text = ""
            txt֧������.Text = ""
            txt��֤����.Text = ""
            If txt���֤��.Visible = True And txt���֤��.Enabled = True Then txt���֤��.SetFocus
        Exit Function
    End If
    
    If Trim(txt֧������.Text) <> Trim(txt��֤����.Text) And (Trim(txt֧������.Text) <> "" Or Trim(txt��֤����.Text) <> "") Then
        MsgBox "������������벻һ��,����������", vbOKOnly + vbInformation, gstrSysName
        txt֧������.Text = "": txt��֤����.Text = ""
        If txt֧������.Visible = True And txt֧������.Enabled = True Then txt֧������.SetFocus
        Exit Function
    End If
    
    '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    If CreatePlugInOK(mlngModul) And mlngPlugInHwnd <> 0 Then  '������������Ϣǰ��������Ч�Լ��
        On Error Resume Next
        blnPlugInCheck = gobjPlugIn.PatiInfoSaveBefore(mlng����ID)
        Call zlPlugInErrH(Err, "PatiInfoSaveBefore")
        If Err = 0 And blnPlugInCheck = False Then
            Exit Function '���δͨ����ֹ����
        End If
        Err.Clear
    End If '
        
    '84672:���ϴ������ȼ���Լ���ϵ�˼��
    If CheckTextLength("������ϵ", txt������ϵ) = False Then Exit Function
    If txt��ϵ������.Text = "" And (txt��ϵ�˵绰.Text <> "" Or txt��ϵ�����֤.Text <> "" Or cbo��ϵ�˹�ϵ.Text <> "") Then
        If MsgBox("û��������ϵ����������ϵ����Ϣ���ᱣ�棬�Ƿ������", vbYesNo + vbInformation, gstrSysName) = vbNo Then
            Exit Function
        Else
            txt��ϵ�����֤.Text = "": txt��ϵ�˵绰.Text = ""
            cbo��ϵ�˹�ϵ.ListIndex = -1: txt������ϵ.Text = "": txt������ϵ.Visible = False
        End If
    End If
    With vsLinkMan
        If .Rows >= 3 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) = "" And (.TextMatrix(i, 1) <> "" Or .TextMatrix(i, 2) <> "" Or .TextMatrix(i, 3) <> "") Then
                    If MsgBox("��ϵ���б��" & i & "��û��������ϵ�����������е���ϵ����Ϣ���ᱣ�棬�Ƿ������", vbYesNo + vbInformation, gstrSysName) = vbNo Then
                        Exit Function
                    Else
                        .TextMatrix(i, 1) = "": .TextMatrix(i, 2) = "": .TextMatrix(i, 3) = "": .TextMatrix(i, 4) = ""
                    End If
                End If
            Next
        End If
    End With
    '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
    If IsCertificateCard(mlng����ID) = False Then Exit Function
    CheckValied = True
End Function

Private Sub cmdOK_Click()
    Dim strPati As String, strCard As String, strMCAccount As String, strTmp As String, str������ϵ As String
    Dim rsCard As ADODB.Recordset, blnTrans As Boolean, strErrMsg As String, blnNewPati As Boolean
    Dim lng����ID As Long, strNO As String
    Dim lngDept As Long, strDate As String, str�������� As String
    Dim str����� As String, byt���� As Byte, i As Integer
    Dim Datsys As Date, str���￨ As String, blnBound As Boolean
    Dim str��ͥ��ַ As String, str���ڵ�ַ As String
    Dim strYLKNo As String, colPro As Collection, blnCard As Boolean
    
    txtPatient.Text = Trim(txtPatient.Text)
    txt����.Text = Trim(txt����.Text)
    
    Set mobjProPati = New Collection
    '��ص�������
    Set colPro = New Collection
    If CheckValied = False Then Exit Sub
    '�����:51072
    If Len(Trim(txt����.Text)) <= 0 And Len(Trim(txt����.Text)) > 0 Then    'û����������
        If zl_Get����Ĭ�Ϸ������� = False Then Exit Sub
    End If

    strMCAccount = Trim(txtPatiMCNO(0).Text)
    If mlngOutModeMC = 920 And strMCAccount <> txtPatiMCNO(0).Tag And strMCAccount <> "" Then
        strMCAccount = UCase(strMCAccount)
    End If
    If txt����ʱ�� = "__:__" Then
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & " " & txt����ʱ��.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
   
    If Len(txt�����.Text) > mintNOLength + 1 And mintNOLength > 0 And mblnCheckNOValidity Then
        MsgBox "ע��,���������Ź���,��ȷ���Ƿ���������!", vbInformation, gstrSysName
        txt�����.SetFocus
        txt�����.SelStart = 0: txt�����.SelLength = Len(txt�����.Text)
        Exit Sub
    End If
    '�����:57326
    If mlng����ID <> 0 And mbytFun <> 0 And txt����.Text <> "" Then
        If Check��������(mlng����ID, gCurSendCard.lng�����ID) = False Then
            txt����.Text = ""
            txt����.Text = ""
            txt��֤.Text = ""
        End If
    End If

    If mbytFun = 1 And (mstrCard <> "" Or txtPatient.Text <> "") And gbln���ѽ����� Or (mbytFun = 0 And mlng����ID <> 0) Then
        Datsys = zlDatabase.Currentdate
        strDate = "To_Date('" & Format(Datsys, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        If Exist�����(txt�����.Text, IIf(mlng����ID <> 0, mlng����ID, 0)) Then
            str����� = zlDatabase.GetNextNo(3)
            If Len(str�����) > txt�����.MaxLength Then
                MsgBox "��ǰ������Ѿ�����������ʹ��,ϵͳ�Զ����������Ϊ:" & str����� & _
                       vbCrLf & "��������������������ų���:" & txt�����.MaxLength & "λ,������һ�������!", vbInformation, gstrSysName
                If txt�����.Enabled Then txt�����.SetFocus
                Exit Sub
            End If
            txt�����.Text = str�����
        End If
        If mstrCard = "" Then mstrCard = txt����.Text
        If mlng����ID <> 0 Then
            lng����ID = mlng����ID
            byt���� = 3
            If mbytFun = 1 And mstrCard <> "" Then blnCard = True
        Else
            lng����ID = zlDatabase.GetNextNo(1)
            byt���� = 1: blnNewPati = True
        End If
        mlng����ID = lng����ID
        '����:����:38663
        Dim strCardNo As String, strPassWord As String
        If gCurSendCard.bln���￨ Then
            strCardNo = Trim(txt����.Text)
            strPassWord = zlCommFun.zlStringEncode(txt����.Text)
        End If
        '�����:51071
        '73609:���ϴ���2014-8-1��������Ϣ����
        '84313,���ϴ�,2015/4/27,��ϵ�˹�ϵ�Լ�������ϵ
        strPati = _
        "zl_�ҺŲ��˲���_INSERT(" & byt���� & "," & lng����ID & "," & txt�����.Text & "," & _
                  "'" & strCardNo & "','" & strPassWord & "'," & _
                  "'" & txtPatient.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & cbo���䵥λ.Text & "'," & _
                  "'" & NeedName(cbo�ѱ�.Text) & "','" & NeedName(cbo���ʽ.Text) & "'," & _
                  "'" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "'," & _
                  "'" & NeedName(cboְҵ.Text, True) & "','" & txt���֤��.Text & "','" & txt��λ����.Text & "'," & _
                  Val(txt��λ����.Tag) & ",'" & txt��λ�绰.Text & "','" & txt��λ�ʱ�.Text & "'," & _
                  "'" & IIf(mblnStructAdress, Trim(padd��ͥ��ַ.Value), cbo��ͥ��ַ.Text) & "'," & _
                  "'" & txt��ͥ�绰.Text & "','" & txt��ͥ�ʱ�.Text & "'," & strDate & ",''," & str�������� & ",'" & strMCAccount & "','" & IIf(mfrmMain.mblnICCard, txt����.Text, "") & "'," & _
                  "NULL," & IIf(Trim(txt����.Text) = "", "NULL,", "'" & Trim(txt����.Text) & "',") & _
                   "'" & IIf(mblnStructAdress, Trim(padd���ڵ�ַ.Value), Trim(txtRegLocation.Text)) & "'," & _
                   "'" & Trim(txt���ڵ�ַ�ʱ�.Text) & "'," & IIf(Trim(txt��ϵ�����֤.Text) = "", "NULL,", "'" & Trim(txt��ϵ�����֤.Text) & "',") & _
                  IIf(Trim(txt��ϵ������.Text) = "", "NULL,", "'" & Trim(txt��ϵ������.Text) & "',") & _
                  IIf(Trim(txt��ϵ�˵绰.Text) = "", "NULL,", "'" & Trim(txt��ϵ�˵绰.Text) & "',") & _
                  IIf(NeedName(cbo��ϵ�˹�ϵ.Text) = "", "NULL,", "'" & NeedName(cbo��ϵ�˹�ϵ.Text) & "',")    '�����:40005
        '�໤��_In         In ������Ϣ.�໤��%Type := Null
        strPati = strPati & IIf(Trim(txt�໤��.Text) = "", "NULL,", "'" & Trim(txt�໤��.Text) & "',")  'lgf
        '54601:������,2013-11-27,���������ص�ͻ��ڵ�ַ
        strPati = strPati & IIf(Trim(txtBirthLocation.Text) = "", "NULL,", "'" & Trim(txtBirthLocation.Text) & "',")
        '�ֻ���_In         In ������Ϣ.�ֻ���%Type := Null
        strPati = strPati & "'" & txtMobile.Text & "')"
        
        '89242:���ϴ�,2015/12/10,���²��˵�ַ��Ϣ
        If mblnStructAdress Then
            If padd��ͥ��ַ.Value <> "" Then
               str��ͥ��ַ = "zl_���˵�ַ��Ϣ_update(1," & lng����ID & ",NULL,3,'" & padd��ͥ��ַ.valueʡ & "','" & _
                   padd��ͥ��ַ.value�� & "','" & padd��ͥ��ַ.value���� & "','" & padd��ͥ��ַ.value���� & "','" & _
                   padd��ͥ��ַ.value��ϸ��ַ & "','" & padd��ͥ��ַ.Code & "')"
            Else
               str��ͥ��ַ = "zl_���˵�ַ��Ϣ_update(2," & lng����ID & ",NULL,3)"
            End If
            
            If padd���ڵ�ַ.Value <> "" Then
               str���ڵ�ַ = "zl_���˵�ַ��Ϣ_update(1," & lng����ID & ",NULL,4,'" & padd���ڵ�ַ.valueʡ & "','" & _
                   padd���ڵ�ַ.value�� & "','" & padd���ڵ�ַ.value���� & "','" & padd���ڵ�ַ.value���� & "','" & _
                   padd���ڵ�ַ.value��ϸ��ַ & "','" & padd���ڵ�ַ.Code & "')"
            Else
               str���ڵ�ַ = "zl_���˵�ַ��Ϣ_update(2," & lng����ID & ",NULL,4)"
            End If
        End If
        
        'str������ϵ
        If cbo��ϵ�˹�ϵ.Text <> "" And txt������ϵ.Visible Then
            str������ϵ = "Zl_������Ϣ�ӱ�_Update("
            '����ID_In ������Ϣ�ӱ�.����Id%Type
            str������ϵ = str������ϵ & "" & lng����ID & ","
            '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type0
            str������ϵ = str������ϵ & "'��ϵ�˸�����Ϣ',"
            '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
            str������ϵ = str������ϵ & "'" & txt������ϵ.Text & "',"
            '����Id_In ������Ϣ�ӱ�.����Id%Type
            str������ϵ = str������ϵ & "'')"
        End If
        '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
        Call AddCertificate(lng����ID, colPro, Datsys)
        mstrFirstCode = ""
    
        If bln����(True) Then
            mblnInRange = True
        Else
            mblnInRange = False
        End If
        If byt���� <> 3 Or blnCard Then
            '��Ϊ����,���ǰ󶨿�
            If bln����(True) = False Then
                blnBound = True
            Else
                blnBound = False
            End If
            Call ReLoadCardFee
            Set rsCard = zlGetSpecialItemFee(gCurSendCard.str��׼��Ŀ, mstrPriceGrade, gCurSendCard.lng�շ�ϸĿID)
            If rsCard Is Nothing Then
                MsgBox "������ȷ��ȡ" & gCurSendCard.str������ & "������Ϣ��", vbInformation, gstrSysName
                Exit Sub
            End If
            '��۲�������޼�Ϊ��ʱ,���տ���
            '98364:���ϴ�,2016/7/7,���ܿ����Ƿ�Ϊ0��������Ӧ�������������ü�¼��
            If Me.txt����.Text <> "" And blnBound = False Then
                '���￨����:���Ǳ��,���ηѱ�
                '��𱣳ֲ���,�շѴ���"�շ��ض���Ŀ"��Ϊ����������
                Select Case rsCard!���ұ�־
                Case 0  '����ȷִ�п���
                    lngDept = UserInfo.����ID
                Case 1  '�������ڿ���
                    lngDept = UserInfo.����ID
                Case 2  '�������ڲ���
                    lngDept = UserInfo.����ID
                Case 3  '���������ڿ���
                    lngDept = UserInfo.����ID
                Case 4  'ָ������
                    lngDept = GetOneDept(rsCard!�շ�ϸĿID)
                Case Else
                    lngDept = UserInfo.����ID
                End Select

                strNO = zlDatabase.GetNextNo(13)
                strYLKNo = zlDatabase.GetNextNo(16)  'ҽ�ƿ�
                strCard = "zl_���ﻮ�ۼ�¼_Insert('" & strNO & "',1," & lng����ID & ",NULL," & txt�����.Text & "," & _
                          "NULL,'" & txtPatient.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & cbo���䵥λ.Text & "'," & _
                          "'" & NeedName(cbo�ѱ�.Text) & "',0," & UserInfo.����ID & "," & _
                          UserInfo.����ID & ",'" & UserInfo.���� & "',NULL," & rsCard!�շ�ϸĿID & "," & _
                          "'" & rsCard!�շ���� & "','" & rsCard!���㵥λ & "',NULL,1,1,0," & lngDept & ",NULL," & _
                          rsCard!������ĿID & ",'" & rsCard!�վݷ�Ŀ & "'," & Format(rsCard!�ּ�, "0.000") & "," & _
                          Format(rsCard!�ּ�, "0.00") & "," & Format(rsCard!�ּ�, "0.00") & "," & strDate & "," & _
                          strDate & ",NULL,'" & UserInfo.���� & "','" & strYLKNo & "')"
                
                '���ڿ�����Ҫ����סԺ���ü�¼
                str���￨ = zlGetSaveCardFeeSQL(gCurSendCard.lng�����ID, 0, strYLKNo, lng����ID, 0, UserInfo.����ID, UserInfo.����ID, 0, _
                zlStr.NeedName(cbo�ѱ�.Text), mstrCard, Trim(txtPatient.Text), zlStr.NeedName(cbo�Ա�.Text), txt����.Text & cbo���䵥λ.Text, _
                txt����.Text, zlCommFun.zlStringEncode(txt����.Text), "�Һŷ���", 0, 0, "", Datsys, mlng�ſ�����ID, rsCard, _
                IIf(mfrmMain.mblnICCard = True, txt����.Text, ""), , , , , strNO)
            ElseIf Me.txt����.Text <> "" Then
                str���￨ = GetCardDataSql(11, lng����ID, gCurSendCard.lng�����ID, mstrCard, Me.txt����.Text, Me.txt����.Text, _
                                    Datsys, "", "�ҺŰ󶨿�")
            End If
        End If
        If strPati <> "" Then zlAddArray mobjProPati, strPati
        If str��ͥ��ַ <> "" Then zlAddArray mobjProPati, str��ͥ��ַ
        If str���ڵ�ַ <> "" Then zlAddArray mobjProPati, str���ڵ�ַ
        If strCard <> "" Then zlAddArray mobjProPati, strCard
        If str���￨ <> "" Then zlAddArray mobjProPati, str���￨
    
    End If
    If Not mblnInRange And (byt���� <> 3 Or blnCard) Or (str���￨ <> "") Then
        '����״̬�£�ֱ�Ӹ�����Ϣ��
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
    '����:���没����Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-11-06 16:07:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTrans As Boolean
    Dim blnNewPati As Boolean
    Dim strErrMsg As String
    Dim i As Long
    
    On Error GoTo errH
    '���豣������
    If mobjProPati Is Nothing Then SaveAfterArrList = True: Exit Function
    '120737,����,2018-2-1,�½������ˣ��ڹҺŲ�����Ϣ(���￨�󶨿�)����󿨺󣬱��汨��
    If mobjProPati.Count = 0 Then SaveAfterArrList = True: Exit Function
    blnNewPati = mlng����ID = 0
    
    blnTrans = True
    zlExecuteProcedureArrAy mobjProPati, Me.Caption, True
    
    '101170:���ϴ�,2017/5/3,����HIS����Ҫ�ύEMPI���ݣ�ʧ�ܺ��������ݶ�Ҫ����
    If zlSaveEMPIPatiInfo(blnNewPati, mlng����ID, 0, strErrMsg) = False Then
        gcnOracle.RollbackTrans
        If strErrMsg = "" Then strErrMsg = "��EMPIƽ̨�ϴ�������Ϣʧ�ܣ�"
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    Set mobjProPati = Nothing
    mstrPlugChange = ""
    If txt����.Text <> "" Then
        If gCurSendCard.bln�Ƿ�д�� Then Call WriteCard(mlng����ID)
        Call mfrmMain.SetCardDisplay(gCurSendCard.str������ & ":" & Me.txt����.Text & "(" & IIf(Not (bln����(True)), "�󶨿�", "����") & ")")
    End If
    '74430,Ƚ����,2014-7-7,�Һ��еĲ�����Ϣ�༭�������ṩ�ɼ���Ƭ����
    Call SavePatiPic(mlng����ID)
    '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    If CreatePlugInOK(mlngModul) And mlngPlugInHwnd <> 0 Then  '������������Ϣ
        On Error Resume Next
        Call gobjPlugIn.PatiInfoSaveAfter(mlng����ID)
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
    Call zlControl.TxtCheckKeyPress(txtMobile, KeyAscii, m����ʽ)
End Sub

Private Function WriteCard(lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:д��
    '���:lng����ID - ����ID
    '����:����
    '����:56599
    '����:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    On Error GoTo ErrHandl:
    
    WriteCard = gobjSquare.objSquareCard.zlBandCardArfter(Me, mlngModul, gCurSendCard.lng�����ID, lng����ID, strExpend)
    Exit Function
ErrHandl:
    WriteCard = False
    If ErrCenter() = 1 Then Resume
End Function

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmd��λ����_Click()
    Call SearchUnit("", txt��λ����)
End Sub

Private Sub cmd����_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    Dim i As Integer
    
    strSQL = _
        " Select -1 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'����ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select -2 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�г�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select -3 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�в�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select ID,nvl(�ϼ�ID,-����) as �ϼ�ID,0 as ĩ��,NULL as ����,����," & _
        " NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ��" & _
        " From ���Ʒ���Ŀ¼ Where ���� IN (1,2,3) And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
        " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        " Union All" & _
        " Select Distinct A.ID,A.����ID as �ϼ�ID,1 as ĩ��,A.����," & _
        " A.����,A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������," & _
        " Decode(B.�Ƿ���ҩ,1,'��','') as ��ҩ,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
        " From ������ĿĿ¼ A,ҩƷ���� B" & _
        " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"

    Set rsTmp = frmPubSel.ShowSelect(Me, strSQL, 2, "����ҩ��", , msh����.TextMatrix(msh����.Row, 0), "��������ҩƷ��ѡ��һ����Ϊ���˹���ҩ�")
    If Not rsTmp Is Nothing Then
        For i = 1 To msh����.Rows - 1
            If i <> msh����.Row Then
                If msh����.RowData(i) = rsTmp!ID Then
                    MsgBox "�� " & i & " �е�ҩ���Ѿ�����ѡ���ҩ��������ͬ,������ѡ��", vbInformation, gstrSysName
                    msh����.SetFocus
                    msh����_EnterCell
                    Exit Sub
                End If
            End If
        Next
        msh����.RowData(msh����.Row) = rsTmp!ID
        msh����.TextMatrix(msh����.Row, 0) = Trim(rsTmp!����)
    End If
    msh����.SetFocus
    msh����_EnterCell
    
End Sub

Private Sub cmd��ͥ��ַ_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select Distinct Substr(����,1,2) as ID,NULL as �ϼ�ID,0 as ĩ��,NULL as ����," & _
            " Substr(����,1,2) as ���� From ����" & _
            " Union All" & _
            " Select ���� as ID,Substr(����,1,2) as �ϼ�ID,1 as ĩ��,����,���� " & _
            " From ���� Order by ����", 2, "����", , cbo��ͥ��ַ.Text)
    If Not rsTmp Is Nothing Then
        cbo��ͥ��ַ.Text = rsTmp!����
        cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text)
    End If
    cbo��ͥ��ַ.SetFocus
End Sub

Private Sub Form_Activate()
    If mbytFun = 1 Then
        picCard.Visible = True
        tbcPage.Top = picCard.Top + picCard.Height
        txt����.Locked = False
        txt����.PasswordChar = IIf(gCurSendCard.str�������� <> "", "*", "")
        txt����.MaxLength = gCurSendCard.lng���ų���
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txt����.IMEMode = 0
    Else
        If txt����.Text <> "" Then
            picCard.Visible = True
        Else
            picCard.Visible = False
        End If
        txt����.Locked = True
    End If
    mblnCancel = False
    tbcPage.Left = 0
    tbcPage.Width = Me.ScaleWidth
    
    Select Case mbytFun
        Case 0
            Me.Caption = "�ҺŲ�����Ϣ�༭"
        Case 1
            Me.Caption = "�Һŷ���"
    End Select
    
    If (mbytFun = 0 And mlng����ID = 0) Or (mbytFun = 1 And mstrCard = "") Then '�󶨾��￨ģʽ���ṩȡ����ť,�Է�Unload����,��Ϊ֮ǰ��ȡ�������ʱ���ص���Ϣ�ᱻ���
        cmdOK.Caption = "ȷ��(&O)"
        cmdCancel.Visible = True
        cmdCancel.Left = tbcPage.Left + tbcPage.Width - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    Else
        cmdOK.Caption = "ȷ��(&O)"
        If (mbytFun = 0 And mlng����ID <> 0) Then
            cmdCancel.Caption = "����(&X)"
        End If
        cmdCancel.Visible = True
        cmdCancel.Left = tbcPage.Left + tbcPage.Width - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    End If
    
    '78408:���ϴ�,2014/10/9,�����ת
    If Not Me.ActiveControl Is msh���� Then
        If mbytFun = 1 And Not gCurSendCard.str������ Like "�������֤" Then
            If mstrCard = "" Then
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
            Else
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
            End If
        Else
            If txtPatient.Visible = True And txtPatient.Enabled Then
                txtPatient.SetFocus
            ElseIf txt�����.Enabled And txt�����.Visible Then
                txt�����.SetFocus
            ElseIf txt��������.Enabled And txt��������.Visible Then
                txt��������.SetFocus
            End If
        End If
    End If
    
    mblnɨ�����֤ = False
    mblnɨ�����֤ǩԼ = IIf(zlDatabase.GetPara("ɨ�����֤ǩԼ", glngSys, mlngModul) = "1", 1, 0) = "1"
    mbln��ֹ�������� = Val(zlDatabase.GetPara("��ֹ��������", glngSys, mlngModul, 0)) = 1
    If mbln��ֹ�������� Then txt����.Enabled = False: cbo���䵥λ.Enabled = False
    SetCtrVisibleAndMove
    '�����:56599
    Me.Caption = "�ҺŲ�����Ϣ��" & gCurSendCard.str������ & IIf(bln����, "����", "�󶨿�") & "��"
    If Not mfrmMain Is Nothing Then
        If mfrmMain.SendCard Then Me.Caption = "�ҺŲ�����Ϣ��" & gCurSendCard.str������ & "����" & "��"
    End If
    gsngStartTime = Timer
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If txt����.Visible Then
            txt����.Visible = False
            msh����_EnterCell
            msh����.SetFocus
        ElseIf lvwItems.Visible Then
            lvwItems.Visible = False
            txt����.Visible = True
            txt����.SetFocus
        ElseIf Not cmdCancel.Visible Then
            cmdOK_Click
        Else
            cmdCancel_Click
        End If
    ElseIf KeyCode = vbKeyF2 Then
        Call cmdOK_Click
    ElseIf KeyCode = vbKeyF4 And Shift = vbCtrlMask Then
        If txt����.Enabled And txt����.Visible Then
            Call lblICCard_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        '89242:���ϴ�,2015/12/10,PatiAddress�ؼ��ڲ���������ת���ⲿ���ٴ���
        If UCase(TypeName(Me.ActiveControl)) = UCase("PatiAddress") Then Exit Sub
        If InStr(1, "txtPatient,txt����,lvwItems,txt����,cbo���䵥λ,txt��������,msh����,txt����,txtPatiMCNO,txt����,vsInoculate,vsCertificate,cbo��ͥ��ַ", Me.ActiveControl.Name) <= 0 Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub


Private Sub InitData()
'���ܣ���ʼ����Ҫ����
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
    
    '����
    rsTmp.Filter = "���='����'"
    cbo����.Clear
    For i = 1 To rsTmp.RecordCount
        cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!ȱʡ = 1 Then
            cbo����.ItemData(cbo����.NewIndex) = 1
            cbo����.ListIndex = cbo����.NewIndex
        End If
        rsTmp.MoveNext
    Next

    '����
    rsTmp.Filter = "���='����'"
    cbo����.Clear
    For i = 1 To rsTmp.RecordCount
        cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!ȱʡ = 1 Then
            cbo����.ItemData(cbo����.NewIndex) = 1
            cbo����.ListIndex = cbo����.NewIndex
        End If
        rsTmp.MoveNext
    Next

    '����״��
    rsTmp.Filter = "���='����״��'"
    cbo����.Clear
    For i = 1 To rsTmp.RecordCount
        cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!ȱʡ = 1 Then
            cbo����.ItemData(cbo����.NewIndex) = 1
            cbo����.ListIndex = cbo����.NewIndex
        End If
        rsTmp.MoveNext
    Next

    'ְҵ
    rsTmp.Filter = "���='ְҵ'"
    cboְҵ.Clear
    For i = 1 To rsTmp.RecordCount
        cboְҵ.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!ȱʡ = 1 Then
            cboְҵ.ItemData(cboְҵ.NewIndex) = 1
            cboְҵ.ListIndex = cboְҵ.NewIndex
        End If
        rsTmp.MoveNext
    Next
    
    '84313,���ϴ�,2015/4/27,��ϵ�˹�ϵ�Լ�������ϵ
    '����ϵ
    rsTmp.Filter = "���='����ϵ'"
    cbo��ϵ�˹�ϵ.Clear
    For i = 1 To rsTmp.RecordCount
        cbo��ϵ�˹�ϵ.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!ȱʡ = 1 Then
            cbo��ϵ�˹�ϵ.ItemData(cbo��ϵ�˹�ϵ.NewIndex) = 1
            cbo��ϵ�˹�ϵ.ListIndex = cbo��ϵ�˹�ϵ.NewIndex
        End If
        rsTmp.MoveNext
    Next
        
    '����ҩ����ҽ���б��ʼ��
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 1400, 0
        .Add , "����", "����", 900
        .Add , "��λ", "��λ", 600
        .Add , "����", "����", 600
        .Add , "�������", "�������", 900
        .Add , "��ҩ", "��ҩ", 600
        .Add , "Ƥ��", "Ƥ��", 600
    End With
    
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").index - 1
        .SortOrder = lvwAscending
        .Visible = False
    End With
    '�����:56599
    Call Init����ҩ��
    
    If cbo���䵥λ.Tag = "" Then
        cbo���䵥λ.Tag = cbo���䵥λ.Text
    End If
End Sub

Public Sub ShowMe(bytMode As Byte, frmParent As Object)
    
    Set mfrmMain = frmParent
    
    If gblnAutoAddName And mbytFun = 1 And mstrCard <> "" Then 'ˢ���Զ�����"�²���"
        txtPatient.Text = "�²���"
        mbln������Ϣ���� = False
        Call cmdOK_Click
    Else
        If mlngOutModeMC > 0 Then
            If mlngOutModeMC = 920 Then
                txtPatiMCNO(0).MaxLength = 12
            Else
                txtPatiMCNO(0).MaxLength = 30
            End If
            txtPatiMCNO(0).ToolTipText = "��󳤶�" & txtPatiMCNO(0).MaxLength & "λ"
            txtPatiMCNO(1).MaxLength = txtPatiMCNO(0).MaxLength
        End If
        Call NewCardObject  '47007
        If txt�����.Text <> "" Then
            mintNOLength = Len(txt�����.Text)
        End If
        '82859:���ϴ�,2015/4/8,���˻�����Ϣ����
        mbln������Ϣ���� = Not (mlng����ID <> 0 And InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";������Ϣ����;") = 0)
        txtPatient.Enabled = mbln������Ϣ����: txt��������.Enabled = mbln������Ϣ����: txt����ʱ��.Enabled = mbln������Ϣ����
        txt����.Enabled = mbln������Ϣ���� And Not mbln��ֹ��������: cbo���䵥λ.Enabled = mbln������Ϣ���� And Not mbln��ֹ��������: cbo�Ա�.Enabled = mbln������Ϣ����
        txt���֤��.Enabled = mbln������Ϣ����
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
    '����������Ϣ��������
    '69026,Ƚ����,2014-8-8,�����������
    Call CreatePublicPatient
    mbln��ͥ��ַ���� = Val(Nvl(zlDatabase.GetPara("��ͥ��ַ���뷽ʽ", glngSys, mlngModul, 1), 1)) = 1
    mblnCheckNOValidity = Val(Nvl(zlDatabase.GetPara("�������Ч�Լ��", glngSys, mlngModul, 1), 1)) = 1
    
    mblnStructAdress = Val(zlDatabase.GetPara(251, glngSys)) <> 0 '���˵�ַ�ṹ��¼��
    mblnShowTown = Val(zlDatabase.GetPara(252, glngSys)) <> 0 '�����ַ�ṹ��¼��
    
    Call InitTagPage
    Call InitTaskPanelOther
    
    txtRegLocation.MaxLength = glngMax���ڵ�ַ
    txtBirthLocation.MaxLength = glngMax�����ص�
    '��ʼ����ַ�ؼ�
    If Not mblnStructAdress Then Exit Sub
    padd��ͥ��ַ.Visible = True: padd���ڵ�ַ.Visible = True
    padd��ͥ��ַ.ShowTown = mblnShowTown: padd���ڵ�ַ.ShowTown = mblnShowTown
    cbo��ͥ��ַ.Visible = False: cmd��ͥ��ַ.Visible = False
    padd��ͥ��ַ.Top = cbo��ͥ��ַ.Top: padd��ͥ��ַ.Left = cbo��ͥ��ַ.Left
    txtRegLocation.Visible = False: cmdRegLocation.Visible = False
    padd���ڵ�ַ.Top = txtRegLocation.Top: padd���ڵ�ַ.Left = txtRegLocation.Left
    
    padd��ͥ��ַ.MaxLength = glngMax��ͥ��ַ
    padd���ڵ�ַ.MaxLength = glngMax���ڵ�ַ
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrCard = ""
    mblnChange = False
    Set mdicҽ�ƿ����� = Nothing
    Set mobjKeyboard = Nothing
    Call CloseIDCard
    mblnPlugin = False
    mlngPlugInHwnd = 0: mblnSavePati = False
    '74430,Ƚ����,2014-7-7,�Һ��еĲ�����Ϣ�༭�������ṩ�ɼ���Ƭ����
    mlngͼ����� = 0: mstr�ɼ�ͼƬ = ""
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
    
    For i = 1 To msh����.Rows - 1
        If i <> msh����.Row Then
            If msh����.RowData(i) = Replace(lvwItems.SelectedItem.Key, "_", "") Then
                MsgBox "�� " & i & " �е�ҩ���Ѿ�����ѡ���ҩ��������ͬ,������ѡ��", vbInformation, gstrSysName
                lvwItems.SetFocus
                Exit Sub
            End If
        End If
    Next
    lvwItems.Visible = False
    msh����.RowData(msh����.Row) = Replace(lvwItems.SelectedItem.Key, "_", "")
    msh����.TextMatrix(msh����.Row, 0) = Trim(lvwItems.SelectedItem.Text)
    msh����.SetFocus
    msh����_EnterCell
    
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    Case vbKeyEscape
        lvwItems.Visible = False
        txt����.Visible = True
        txt����.SetFocus
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msh����_Click()
    msh����_EnterCell
End Sub

Private Sub msh����_GotFocus()
    msh����_EnterCell
End Sub

Private Sub msh����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If KeyCode = vbKeyF4 Then msh����_DblClick
    If KeyCode = vbKeyF3 And cmd����.Visible Then cmd����_Click
    If KeyCode = vbKeyDelete Then
        msh����.TextMatrix(msh����.Row, 0) = ""
        msh����.RowData(msh����.Row) = 0
        For i = msh����.Row + 1 To msh����.Rows - 1
            msh����.TextMatrix(i - 1, 0) = msh����.TextMatrix(i, 0)
            msh����.RowData(i - 1) = msh����.RowData(i)
            msh����.TextMatrix(i, 0) = ""
            msh����.RowData(i) = 0
        Next
        msh����_EnterCell
    End If
End Sub

Private Sub msh����_DblClick()
    cmd����.Visible = False
    txt����.Visible = False
    
    'If msh����.Row > 1 And msh����.TextMatrix(msh����.Row - 1, 0) = "" Or msh����.RowData(msh����.Row) = 0 Then Exit Sub
    Select Case msh����.Col
        Case 0 '����ҩ��
            txt����.Top = msh����.CellTop + msh����.Top + (msh����.CellHeight - txt����.Height) / 2 - 15
            txt����.Left = msh����.Left + msh����.CellLeft + 30
            txt����.Width = msh����.CellWidth - 60
            
            txt����.Text = msh����.TextMatrix(msh����.Row, msh����.Col)
            txt����.ZOrder
            Call zlControl.TxtSelAll(txt����)
            txt����.Visible = True
            If txt����.Visible Then txt����.SetFocus
        Case 1 '������Ӧ
            txt������Ӧ.Top = msh����.CellTop + msh����.Top + (msh����.CellHeight - txt������Ӧ.Height) / 2 - 15
            txt������Ӧ.Left = msh����.Left + msh����.CellLeft + 30
            '75446:���ϴ�,2014-7-16,������Ӧ�ı��򲻹�
            txt������Ӧ.Width = msh����.CellWidth - 60
            
            txt������Ӧ.Text = msh����.TextMatrix(msh����.Row, msh����.Col)
            txt������Ӧ.ZOrder
            Call zlControl.TxtSelAll(txt������Ӧ)
            txt������Ӧ.Visible = True
            If txt������Ӧ.Visible Then txt������Ӧ.SetFocus
    End Select
End Sub

Private Sub msh����_EnterCell()
    cmd����.Visible = False
    txt����.Visible = False
    
    '�����:56599
    If msh����.Row > 1 And msh����.TextMatrix(msh����.Row - 1, 0) = "" Or msh����.Col = 1 Then Exit Sub
    
    cmd����.Top = msh����.CellTop + msh����.Top - 15
    If msh����.Rows < 5 Then
        cmd����.Left = msh����.Left + msh����.CellWidth - cmd����.Width + 45
    Else
        cmd����.Left = msh����.Left + msh����.CellWidth - cmd����.Width + 45
    End If
    
    cmd����.ZOrder
    cmd����.Visible = True
End Sub

Private Sub msh����_KeyPress(KeyAscii As Integer)
        If KeyAscii <> 13 Then
            'If msh����.Row > 1 And msh����.TextMatrix(msh����.Row - 1, 0) = "" Or msh����.RowData(msh����.Row) = 0 Then Exit Sub
            msh����_DblClick
            If msh����.Col = 0 Then msh����.RowData(msh����.Row) = 0
            If msh����.Col = 0 Then txt����.Text = Chr(KeyAscii)
            If msh����.Col = 0 Then txt����.SelStart = Len(txt����.Text)
            '75446:���ϴ�,2014-7-16,�༭������¼ʱ���¼����ı���
            If msh����.Col = 1 Then txt������Ӧ.Text = txt������Ӧ.Text & Chr(KeyAscii)
            If msh����.Col = 1 Then txt������Ӧ.SelStart = Len(txt������Ӧ.Text)
        Else
             If msh����.Row = msh����.Rows - 1 And msh����.TextMatrix(msh����.Row, 0) <> "" Then
                msh����.Rows = msh����.Rows + 1
                msh����.Row = msh����.Rows - 1
                '�����:56599
                txt������Ӧ.Text = ""
                txt������Ӧ.Visible = False
                
                msh����_EnterCell
            ElseIf msh����.TextMatrix(msh����.Row, 0) <> "" Then
                msh����.Row = msh����.Row + 1
                msh����_EnterCell
            Else
                cmdOK.SetFocus
            End If
        End If
End Sub
Private Sub msh����_Scroll()
    cmd����.Visible = False
    '�����:56599
    txt����.Visible = False
    txt������Ӧ.Visible = False
End Sub

Private Sub padd���ڵ�ַ_Change()
    If mblnStructAdress Then mstrPlugChange = mstrPlugChange & ",���ڵ�ַ"
End Sub

Private Sub padd��ͥ��ַ_Change()
    If mblnStructAdress Then mstrPlugChange = mstrPlugChange & ",��סַ"
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
    With msh����
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
    mstrPlugChange = mstrPlugChange & ",������ַ"
    txtBirthLocation.Tag = ""
End Sub

Private Sub txtBirthLocation_GotFocus()
    Call zlControl.TxtSelAll(txtBirthLocation)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub SearchAddress(ByVal strInput As String, txtInput As Object)
    '--------------------------------------------------------------
    '����:ģ�����ң���������ѡ���б�
    '����:Ƚ����
    '����:2014-5-23
    '����:
    '   strInput:�����ı�����Ϊ�ձ�ʾ�����ť����
    '   txtInput:�ı������
    '--------------------------------------------------------------
    Dim strSQL As String, strWhere As String
    Dim strKey As String, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset, vRect As RECT
    
    On Error GoTo Errhand
    If strInput <> "" And txtInput.Tag <> "" Then Exit Sub
    vRect = zlControl.GetControlRect(txtInput.Hwnd)
    If strInput = "" Then '�����ť
        strSQL = "" & _
            "Select ID, �ϼ�id, ����, ����, ĩ�� " & _
            "From (With ����_t As" & _
            "    (Select Rownum As �к�, ID, �ϼ�id, ĩ��, ����, ����" & _
            "     From (Select Distinct Substr(����, 1, 2) As ID, Null As �ϼ�id, 0 As ĩ��, Null As ����, Substr(����, 1, 2) As ����" & _
            "            From ����" & _
            "            Union All" & _
            "            Select ���� As ID, Substr(����, 1, 2) As �ϼ�id, 1 As ĩ��, ����, ���� From ����))" & _
            "   Select �к� As ID, To_Number(�ϼ�id) As �ϼ�id, ����, ����, ĩ�� From ����_t Where �ϼ�id Is Null" & _
            "   Union All" & _
            "   Select b.�к�, a.�к�, b.����, b.����, b.ĩ�� From ����_t A, ����_t B Where a.Id = b.�ϼ�id Order By ����)"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "����", False, _
                       "", "", False, False, False, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False)
    Else
        'ȥ��"'"
        strInput = Replace(strInput, "'", " ")
        strKey = GetMatchingSting(strInput, False)
        If strInput <> "" Then
            If IsNumeric(strInput) Then '����ȫ������ʱֻƥ�����
                strWhere = " Where ���� Like Upper([1])"
            ElseIf zlCommFun.IsCharAlpha(strInput) Then '����ȫ����ĸʱֻƥ�����
                strWhere = " Where ���� Like Upper([1])"
            Else
                strWhere = " Where ���� Like Upper([1]) Or ���� Like [1] Or ���� Like Upper([1])"
            End If
        End If
        
        strSQL = "" & _
            "Select Rownum As ID, ����, ���� From ���� " & strWhere & " Order By ����"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����", False, _
                       "", "", False, False, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False, strKey)
    End If
    If blnCancel Then txtInput.SetFocus: Exit Sub

    If rsTemp Is Nothing Then txtInput.SetFocus: Exit Sub
    If rsTemp.State <> 1 Then txtInput.SetFocus: Exit Sub
    
    txtInput.Text = Nvl(rsTemp!����)
    txtInput.Tag = Nvl(rsTemp!ID)
    txtInput.SelStart = Len(Nvl(txtInput.Text))
    txtInput.SetFocus
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtBirthLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    '73022,Ƚ����,2014-5-20,�ڵ�λ���ơ������ص㡢���ڵ�ַ����ģ�����ҹ���
    If KeyCode = vbKeyReturn And Trim(txtBirthLocation.Text) <> "" Then
        Call SearchAddress(Trim(txtBirthLocation.Text), txtBirthLocation)
    End If
End Sub

Private Sub txtBirthLocation_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtMobile_Validate(Cancel As Boolean)
    If txtMobile.Text <> "" And IsMobileNO(txtMobile.Text) = False Then
        MsgBox "������ֻ��Ÿ�ʽ����ȷ��������¼�룡", vbInformation, gstrSysName
        Cancel = True
        Exit Sub
    End If
    If Exist�ֻ���(txtMobile.Text, IIf(mlng����ID <> 0, mlng����ID, 0)) Then
        If MsgBox("������ֻ��������������ظ����Ƿ�ȷ��¼�룿", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Cancel = True
    End If
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If mblnNameChange = True And mlng����ID = 0 Then zlQueryEMPIPatiInfo
    mblnNameChange = False
End Sub

Private Sub txtPatiMCNO_Change(index As Integer)
    If index = 0 Then mstrPlugChange = mstrPlugChange & ",ҽ����"
End Sub

Private Sub txtRegLocation_Change()
    If Not mblnStructAdress Then mstrPlugChange = mstrPlugChange & ",���ڵ�ַ"
    txtRegLocation.Tag = ""
End Sub

Private Sub txtRegLocation_GotFocus()
    Call zlControl.TxtSelAll(txtRegLocation)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtRegLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    '73022,Ƚ����,2014-5-20,�ڵ�λ���ơ������ص㡢���ڵ�ַ����ģ�����ҹ���
    If KeyCode = vbKeyReturn And Trim(txtRegLocation.Text) <> "" Then
        Call SearchAddress(Trim(txtRegLocation.Text), txtRegLocation)
    End If
End Sub

Private Sub txtRegLocation_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtPatient_Change()
    mstrPlugChange = mstrPlugChange & ",����"
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
    If cbo���ʽ.ListCount > 0 Then cbo���ʽ.ListIndex = 0

    If index = 1 Then
        If txtPatiMCNO(1).Text <> txtPatiMCNO(0).Text Then
            MsgBox "����,���������ҽ���Ų�һ�£�", vbInformation, gstrSysName
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

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If txt��������.Text = "____-__-__" Then
           zlCommFun.PressKey (vbKeyTab) '����ʱ��
           zlCommFun.PressKey (vbKeyTab)
       Else
           zlCommFun.PressKey (vbKeyTab)
       End If
    End If

End Sub

Private Sub txt����ʱ��_Change()
    Dim str����ʱ�� As String
    If txt����ʱ��.Text <> txt����ʱ��.Tag And InStr(mstrPlugChange, "��������") = 0 Then mstrPlugChange = mstrPlugChange & ",��������"
    '76669�����ϴ�,2014-8-18,�����������
    If IsDate(txt��������.Text) Then
        str����ʱ�� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        txt����.Text = ReCalcOld(CDate(str����ʱ��), cbo���䵥λ)
        txt����.Tag = txt����.Text
    End If
End Sub

Private Sub txt����ʱ��_GotFocus()
    zlControl.TxtSelAll txt����ʱ��
End Sub

Private Sub txt����ʱ��_KeyPress(KeyAscii As Integer)
    If Not IsDate(txt��������.Text) Then
        KeyAscii = 0
        txt����ʱ��.Text = "__:__"
    End If
End Sub


Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    If txt����ʱ��.Text <> "__:__" And Not IsDate(txt����ʱ��.Text) Then
        txt����ʱ��.SetFocus
        Cancel = True
    End If
End Sub

Private Sub txt��������_Change()
    Dim str����ʱ�� As String
    If txt��������.Text <> txt��������.Tag And InStr(mstrPlugChange, "��������") = 0 Then mstrPlugChange = mstrPlugChange & ",��������"
    If IsDate(txt��������.Text) And mblnChange Then
        mblnChange = False
        txt��������.Text = Format(CDate(txt��������.Text), "yyyy-mm-dd") '0002-02-02�Զ�ת��Ϊ2002-02-02,����,��������2002,ʵ��ֵȴ��0002
        mblnChange = True
        
        str����ʱ�� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        txt����.Text = ReCalcOld(CDate(str����ʱ��), cbo���䵥λ)
        txt����.Tag = txt����.Text
        cbo���䵥λ.Tag = cbo���䵥λ.Text
        mblnGetBirth = False
    End If
End Sub
Private Sub txt��������_GotFocus()
    zlControl.TxtSelAll txt��������
End Sub

Private Sub txt��������_LostFocus()
    If txt��������.Text <> "____-__-__" And Not IsDate(txt��������.Text) Then
      If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus
    End If
End Sub


Private Sub txt��λ�绰_GotFocus()
    Call zlControl.TxtSelAll(txt��λ�绰)
End Sub

Private Sub txt��λ�绰_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt��λ�绰, KeyAscii
End Sub

Private Sub txt��λ����_Change()
     mstrPlugChange = mstrPlugChange & ",��λ����"
    txt��λ����.Tag = ""
End Sub

Private Sub txt��λ����_GotFocus()
    Call zlControl.TxtSelAll(txt��λ����)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��λ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And cmd��λ����.Enabled And cmd��λ����.Visible Then cmd��λ����_Click
    '73022,Ƚ����,2014-5-20,�ڵ�λ���ơ������ص㡢���ڵ�ַ����ģ�����ҹ���
    If KeyCode = vbKeyReturn And Trim(txt��λ����.Text) <> "" Then
        Call SearchUnit(Trim(txt��λ����.Text), txt��λ����)
    End If
End Sub

Private Sub SearchUnit(ByVal strInput As String, txtInput As Object)
    '--------------------------------------------------------------
    '����:ģ�����ң�������Լ��λѡ���б�
    '����:Ƚ����
    '����:2014-5-23
    '����:
    '   strInput:�����ı�����Ϊ�ձ�ʾ�����ť����
    '   txtInput:�ı������
    '--------------------------------------------------------------
    Dim strSQL As String, strWhere As String
    Dim strKey As String, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset, vRect As RECT
    
    On Error GoTo Errhand
    If strInput <> "" And txtInput.Tag <> "" Then Exit Sub
    vRect = zlControl.GetControlRect(txtInput.Hwnd)
    If strInput = "" Then '�����ť
        strSQL = "" & _
        "       Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From  ��Լ��λ" & _
        "       Where ����ʱ�� Is Null Or ����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD')" & _
        "       Start With �ϼ�ID is NULL" & _
        "       Connect by Prior ID=�ϼ�ID"
        '75888,Ƚ����,2014-7-28
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "��λ", False, _
                       "", "", False, True, False, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False)
    Else
        'ȥ��"'"
        strInput = Replace(strInput, "'", " ")
        strKey = GetMatchingSting(strInput, False)
        If strInput <> "" Then
            If IsNumeric(strInput) Then '����ȫ������ʱֻƥ�����
                strWhere = " Where ���� Like Upper([1])"
            ElseIf zlCommFun.IsCharAlpha(strInput) Then '����ȫ����ĸʱֻƥ�����
                strWhere = " Where ���� Like Upper([1])"
            Else
                strWhere = " Where ���� Like Upper([1]) Or ���� Like [1] Or ���� Like Upper([1])"
            End If
        End If
        
        strSQL = "" & _
        "       Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From  ��Լ��λ" & strWhere & _
        "       And (����ʱ�� Is Null Or ����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD'))"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��λ", False, _
                       "", "", False, False, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False, strKey)
    End If
    If blnCancel Then txtInput.SetFocus: Exit Sub

    If rsTemp Is Nothing Then txtInput.SetFocus: Exit Sub
    If rsTemp.State <> 1 Then txtInput.SetFocus: Exit Sub
    
    txtInput.Text = Nvl(rsTemp!����)
    txtInput.Tag = Nvl(rsTemp!ID)
    txtInput.SelStart = Len(Nvl(txtInput.Text))
    txtInput.SetFocus
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt��λ����_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt��λ����, KeyAscii
End Sub

Private Sub txt��λ����_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt��λ�ʱ�_Change()
    mstrPlugChange = mstrPlugChange & ",��λ�ʱ�"
End Sub

Private Sub txt��λ�ʱ�_GotFocus()
    Call zlControl.TxtSelAll(txt��λ�ʱ�)
End Sub

Private Sub txt��λ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt��λ�ʱ�, KeyAscii
End Sub

Private Sub txt����_Change()
    '75286:���ϴ���2014-7-16������¼�����ҩ��
    msh����.TextMatrix(msh����.Row, 0) = txt����.Text
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim ObjItem As ListItem
    Dim strSQL As String
            
    If KeyAscii <> 13 Then
        If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If KeyAscii <> vbKeyEscape Then msh����.RowData(msh����.Row) = 0
    Else
        KeyAscii = 0

        strSQL = " Select Distinct A.ID,A.����," & _
        " A.����,A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������," & _
        " Decode(B.�Ƿ���ҩ,1,'��','') as ��ҩ,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
        " From ������ĿĿ¼ A,ҩƷ���� B,������Ŀ���� C" & _
        " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID And A.Id=C.������Ŀid" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        " And (C.���� like [1] OR A.���� like [1] OR C.���� like [1])"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrLike & UCase(txt����.Text) & "%")
        
        With rsTmp
            If .BOF Or .EOF Then
                msh����.SetFocus: msh����_EnterCell
                Exit Sub
            Else
                Me.lvwItems.ListItems.Clear
                Do While Not .EOF
                    Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����, , IIf(!Ƥ�� <> "", 1, 2))
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").index - 1) = !����
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("��λ").index - 1) = IIf(IsNull(!��λ), "", !��λ)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").index - 1) = IIf(IsNull(!����), "", !����)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("�������").index - 1) = IIf(IsNull(!�������), "", !�������)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("��ҩ").index - 1) = IIf(IsNull(!��ҩ), "", !��ҩ)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("Ƥ��").index - 1) = IIf(IsNull(!Ƥ��), "", !Ƥ��)
                    .MoveNext
                Loop
                Me.lvwItems.ListItems(1).Selected = True
            End If
        End With
        
        With Me.lvwItems
            .Left = msh����.Left
            .Width = msh����.Width
            .Height = msh����.Height + 300
            If msh����.Rows < 5 Then
                .Top = msh����.Top + msh����.RowHeight(msh����.Row) * (msh����.Row) - .Height
            Else
                .Top = msh����.Top + msh����.RowHeight(4) * (3) - .Height
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
'75446:���ϴ�,2014-7-16,�༭������¼ʱ���¼����ı���
Private Sub txt����_LostFocus()
    txt����.Visible = False
End Sub

Private Sub txt������Ӧ_Change()
   '�����:56599
   msh����.TextMatrix(msh����.Row, 1) = txt������Ӧ.Text
End Sub
'75446:���ϴ�,2014-7-16,�༭������¼ʱ���¼����ı���
Private Sub txt������Ӧ_LostFocus()
    txt������Ӧ.Visible = False
End Sub

Private Sub txt���ڵ�ַ�ʱ�_Change()
    mstrPlugChange = mstrPlugChange & ",�����ʱ�"
    mblnChange = True
End Sub

Private Sub txt���ڵ�ַ�ʱ�_GotFocus()
    Call zlControl.TxtSelAll(txt���ڵ�ַ�ʱ�)
End Sub

Private Sub txt���ڵ�ַ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt��ͥ�绰_Change()
    mstrPlugChange = mstrPlugChange & ",�绰"
End Sub

Private Sub txt��ͥ�ʱ�_Change()
    mstrPlugChange = mstrPlugChange & ",��סַ�ʱ�"
End Sub

Private Sub txt��ͥ�绰_GotFocus()
    Call zlControl.TxtSelAll(txt��ͥ�绰)
End Sub

Private Sub txt��ͥ�绰_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt��ͥ�绰, KeyAscii
End Sub

Private Sub txt��ͥ�ʱ�_GotFocus()
    Call zlControl.TxtSelAll(txt��ͥ�ʱ�)
End Sub

Private Sub txt��ͥ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt��ͥ�ʱ�, KeyAscii
End Sub

Private Sub txt��ϵ������_Change()
    mstrPlugChange = mstrPlugChange & ",��ϵ������"
End Sub

Private Sub txt�໤��_GotFocus()
    zlCommFun.OpenIme (True)
End Sub

Private Sub txt�໤��_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txt����_GotFocus()
    '72686:���ϴ�,2015/3/25,������ȫѡ
    Call zlControl.TxtSelAll(txt����)
    Call SetBrushCardObject(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    
    Dim blnCard  As Boolean
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    blnCard = zlCommFun.InputIsCard(txt����, KeyAscii, gCurSendCard.str�������� <> "")
    If blnCard And Len(txt����.Text) = gCurSendCard.lng���ų��� - 1 And KeyAscii <> 8 Then
        txt����.Text = txt����.Text & Chr(KeyAscii): txt����.SelStart = Len(txt����.Text)
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt����_LostFocus()
    Call SetBrushCardObject(False)
End Sub

Private Sub txt����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '72686:���ϴ�,2015/3/25,������ȫѡ
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim lngPatientID As Long
    Dim lng�䶯���� As Long
    Dim blnCardBind As Boolean  '���Ƿ���а�
    If gCurSendCard.lng���ų��� = Len(Trim(txt����.Text)) Then
        If Bln�ѷ���(txt����.Text, gCurSendCard.lng�����ID, lngPatientID) Then
            If gCurSendCard.bln���ƿ� And gCurSendCard.bln�ظ�ʹ�� And lngPatientID > 0 Then
                lng�䶯���� = GetCardLastChangeType(txt����.Text, gCurSendCard.lng�����ID, lngPatientID)
                If lng�䶯���� = 11 Then
                    '����ǰ�
                    If MsgBox("����Ϊ��" & txt����.Text & "����{" & gCurSendCard.str������ & "}�Ŀ��Ѿ��벡�˱�ʶΪ��" & lngPatientID & "���Ľ����˰󶨣�" & vbCrLf & "�Ƿ�ȡ���ÿ��İ�?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                        Cancel = True
                        txt����.Text = ""
                        Exit Sub
                    End If
                    If BlandCancel(gCurSendCard.lng�����ID, Trim(txt����.Text), lngPatientID) Then
                        Exit Sub
                    End If
                End If
            End If

            MsgBox "�ÿ����Ѿ�����,���ܰ󶨸ÿ���.", vbInformation, gstrSysName
            Cancel = True
            txt����.Text = ""
            Exit Sub
        End If
    End If

    If mbytFun = 1 And mstrCard = "" And Trim(txt����.Text) <> "" Then
        If GetPatientState(txt����.Text) <> 0 Then
            MsgBox "�ÿ��ŵĳֿ��������ھ����ȴ�����,���ܰ󶨸ÿ���.", vbInformation, gstrSysName
            Cancel = True
        End If
        If Not mfrmMain Is Nothing Then
            Call mfrmMain.zlReadPlugInPati(Trim(txt����))
        End If
        If Not gCurSendCard.bln���ƿ� Then
            '42947
            If zlLoadInfor = False Then Cancel = True
        End If
    End If

    If gCurSendCard.str������ = "�������֤" Then Exit Sub
End Sub

Private Function CheckPatiValid(ByVal strCard As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ָ������Ŀ����Ƿ�Ϸ�
    '��Σ�strCard-ָ���Ŀ���
    '���أ��Ϸ�,����True,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-19 10:14:31
    '˵����31182
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String, lng����ID As Long
    
    '69954:������,2014-02-19,����ҺŹ����޷����Ѿ����˿�����ȡ���󶨿�������
    '72324:������,2014-04-24,�����Һ�ʧ��ʱ���¹ҺŲ��ܷ���������
    '74894:���ϴ�,2014-07-08,������Ϣ��������
    strSQL = "Select Nvl(a.����״̬, 0) ����״̬, a.����id, a.����, a.�Ա�" & vbNewLine & _
             "From ������Ϣ A, ����ҽ�ƿ���Ϣ B, ҽ�ƿ���� C" & vbNewLine & _
             "Where a.���￨�� = b.���� And c.�ض���Ŀ = '���￨' And b.�����id = c.Id And b.�����id=[1] And b.���� = [2]"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gCurSendCard.lng�����ID, strCard)
    If rsTmp.RecordCount = 0 Then CheckPatiValid = True: Exit Function
    
    '1.���״̬:ԭ����Ҫ��������￨ʱ���м���,����txt����_Validate����,��һ���ܼ�鵽,���,�������ڰ�ȷ��ʱ,���Ӹü��
    If Val(Nvl(rsTmp!����״̬)) <> 0 Then
        MsgBox "����Ϊ" & strCard & "�Ĳ������ھ����ȴ�����,���ܰ󶨸ÿ���.", vbInformation, gstrSysName
        Exit Function
    End If
    
    '2.����Ƿ���������ͬ
    If Nvl(rsTmp!����) <> Trim(txtPatient.Text) And Val(txt����.Tag) = 0 Then
       If MsgBox("�ֿ����ˡ�" & Nvl(rsTmp!����) & "��������Ĳ��ˡ�" & Trim(txtPatient.Text) & "����һ��,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    '3.�ҺŲ�����ˢ���￨�ó��Ĳ�����������ͬ�����Ĳ���
    lng����ID = Val(Nvl(rsTmp!����ID))
    If Val(txt����.Tag) <> lng����ID And Val(txt����.Tag) <> 0 Then
        If Nvl(rsTmp!����) <> Trim(txtPatient.Text) Then
            If MsgBox("ע��: " & vbCrLf & _
                             "     �ֿ����ˡ�" & Nvl(rsTmp!����) & "��������Ĳ��ˡ�" & Trim(txtPatient.Text) & "����һ��," & vbCrLf & _
                             "     ��ͬʱ���ǽ�������,�Ƿ񽫲��ˡ�" & Trim(txtPatient.Text) & "���ϲ������ˡ�" & Nvl(rsTmp!����) & "����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            '�ϲ�
            If zlPatiMerge(Val(txt����.Tag), lng����ID, True) = False Then Exit Function
        Else '����������ͬ,�Զ����кϲ�
            '�Զ��ϲ�
            If zlPatiMerge(Val(txt����.Tag), lng����ID, False) = False Then Exit Function
        End If
        '����ˢ����ص�����
        RaiseEvent PatiMerged(lng����ID)
        
    End If
    CheckPatiValid = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPatientState(strCard As String, Optional lng����ID As Long) As Long
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���鲡�˵ĵ�ǰ״̬
    '��Σ�strCard-����
    '���Σ�lng����ID-���ز���ID
    '���أ����ز���״̬
    '���ƣ����˺�
    '���ڣ�2010-07-19 09:55:25
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strPassWord As String, strErrMsg As String
    '42947
    If gobjSquare.objSquareCard.zlGetPatiID(gCurSendCard.lng�����ID, strCard, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
    If lng����ID = 0 Then Exit Function
    
    strSQL = "Select Nvl(����״̬,0) ����״̬,����ID From ������Ϣ Where ����ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    lng����ID = 0
    If rsTmp.RecordCount > 0 Then
        lng����ID = Val(Nvl(rsTmp!����ID))
        GetPatientState = rsTmp!����״̬
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt��ϵ�˵绰_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("�绰") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("�绰")) = txt��ϵ�˵绰.Text
    End If
End Sub

Private Sub txt��ϵ�����֤_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt��ϵ�����֤_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("���֤��") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("���֤��")) = txt��ϵ�����֤.Text
    End If
End Sub

Private Sub txt��ϵ������_GotFocus()
    zlCommFun.OpenIme (True)
End Sub

Private Sub txt��ϵ������_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txt��ϵ������_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("����") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("����")) = txt��ϵ������.Text
        If vsLinkMan.Rows = vsLinkMan.FixedRows + 1 And txt��ϵ������.Text <> "" Then
            vsLinkMan.Rows = vsLinkMan.Rows + 1
        End If
    End If
End Sub

Private Sub txt�����_Change()
    mstrPlugChange = mstrPlugChange & ",�����"
End Sub

Private Sub txt�����_GotFocus()
    Call zlControl.TxtSelAll(txt�����)
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
     
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If txt�����.Enabled And txt�����.Visible And mintNOLength > 0 Then
        '����ֹ��������쳣�����������ʾ
            If Len(txt�����.Text) > mintNOLength + 1 Then
                MsgBox "ע��,���������Ź���,��ȷ���Ƿ���������!", vbInformation, gstrSysName
                txt�����.SetFocus
                txt�����.SelStart = 0: txt�����.SelLength = Len(txt�����.Text)
                Exit Sub
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = 32 Then
        KeyAscii = 0
        If txt�����.Text = "" Then
            txt�����.Text = zlDatabase.GetNextNo(3)
            mintNOLength = Len(Trim(txt�����.Text))
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Or InStr(";" & mstrPrivs & ";", ";�����޸������;") = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt����_Change()
    txt��֤.Enabled = txt����.Text <> ""
    If txt����.Text = "" Then txt��֤.Text = ""
    If gCurSendCard.str������ = "�������֤" Then
        txt֧������.Text = txt����.Text
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
    Call OpenPassKeyboard(txt����, False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, gCurSendCard.int������� = 1)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����.Text = "" Then
            txt��֤.Text = ""
            If tbcPage.Selected.index = 0 Then
                '104243:���ϴ�,2016/12/29,���㶨λʱ����Ƿ����
                If txtPatient.Visible And txtPatient.Enabled Then
                    txtPatient.SetFocus
                Else
                    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
                End If
            End If
        Else
            txt��֤.SetFocus
        End If
    End If
End Sub
Private Sub txt����_LostFocus()
    Call ClosePassKeyboard(txt����)
End Sub
Private Sub CheckInputPassWord(KeyAscii As Integer, Optional ByVal blnOnlyNum As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:���˺�
    '����:2011-07-07 00:40:53
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

Private Sub txt����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt����.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt����.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt����_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt����.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlCommFun.OpenIme
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim blnTab As Boolean
    
    If KeyAscii = vbKeyReturn Then
        If cbo���䵥λ.Visible = False And IsNumeric(txt����.Text) Then
            Call txt����_Validate(False)
            Call cbo���䵥λ.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt����.Text) And cbo���䵥λ.Visible Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        '�������Ƽ��� ָ����������ַ� ����:49908
        If InStr("~����@#��%����&*��������-+=|����������~`!#$%^&*()-_=+|\/?<>,/<>", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim strBirth As String
    txt����.Text = Trim(txt����.Text)
    If Not IsNumeric(txt����.Text) And Trim(txt����.Text) <> "" Then
        cbo���䵥λ.ListIndex = -1: cbo���䵥λ.Visible = False: txt����.Width = 1485
    ElseIf cbo���䵥λ.Visible = False Then
        cbo���䵥λ.ListIndex = 0: cbo���䵥λ.Visible = True: txt����.Width = 690
    End If
    If txt����.Text <> txt����.Tag Then
        mblnChange = False
        If Not IsDate(txt��������.Text) Then mblnGetBirth = True
'        txt��������.Text = ReCalcBirth(Trim(txt����.Text), IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""))
        If mblnGetBirth Then
            If mobjPubPatient.ReCalcBirthDay(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""), strBirth) Then
                txt��������.Text = Format(strBirth, "yyyy-mm-dd")
                txt����ʱ��.Text = Format(strBirth, "hh:mm")
            End If
        End If
        mblnChange = True
        txt����.Tag = txt����.Text
    End If
    '69026,Ƚ����,2014-8-8,�����������
    '76703,Ƚ����,2014-8-15
    If cbo���䵥λ.Visible Then Exit Sub
    If mobjPubPatient Is Nothing Then Exit Sub
    If mobjPubPatient.CheckPatiAge(Trim(txt����.Text), _
            IIf(txt��������.Text = "____-__-__", "", txt��������.Text) & _
            IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text)) = False Then
        Cancel = True
    End If
End Sub

Private Sub txt������ϵ_GotFocus()
    Call zlControl.TxtSelAll(txt��λ����)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt������ϵ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt������ϵ_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("������Ϣ") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("��ϵ")) = NeedName(cbo��ϵ�˹�ϵ.Text)
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("������Ϣ")) = txt������ϵ.Text
    End If
End Sub

Private Sub txt����_Change()
    txt����.Tag = ""
End Sub

Private Sub txt����_GotFocus()
    zlCommFun.OpenIme (True)
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If (txt����.Tag <> "" Or Trim(txt����.Text) = "") Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If zl_SelectAndNotAddItem(Me, txt����, Trim(txt����.Text), "����", "����ѡ��", True, False) = False Then
        Exit Sub
    End If
End Sub

Private Sub txt����_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txt���֤��_Change()
    mstrPlugChange = mstrPlugChange & ",���֤��"
    If mblnɨ�����֤ǩԼ And ActiveControl Is txt���֤�� And Not mobjIDCard Is Nothing Then
            mobjIDCard.SetEnabled txt���֤��.Text = ""
    End If
End Sub

Private Sub txt���֤��_GotFocus()
    Call zlControl.TxtSelAll(txt���֤��)
'    If gCurSendCard.str������ <> "�������֤" Then
'        Call OpenIDCard
'    End If
    If mblnɨ�����֤ǩԼ = True And txt���֤��.Text = "" Then
        OpenIDCard
    End If
End Sub
Private Sub txt���֤��_KeyPress(KeyAscii As Integer)
    
'    If gCurSendCard.str������ <> "�������֤" Then
'        If zl��ǰ�û����֤�Ƿ��(Trim(txt���֤��.Text), Trim(txtPatient.Text), Trim(txt�����.Text)) = True Then
'            MsgBox "��ǰ�û������֤���Ѿ��󶨣��������޸������֤��", vbInformation, gstrSysName
'            KeyAscii = 0
'        End If
'    End If
    
    If zl��ǰ�û����֤�Ƿ��(Trim(txt���֤��.Text), Trim(txtPatient.Text), Trim(txt�����.Text)) = True Then
        MsgBox "��ǰ�û������֤���Ѿ��󶨣��������޸������֤��", vbInformation, gstrSysName
        KeyAscii = 0
    End If

    mblnɨ�����֤ = False
    txt֧������.Text = ""
    txt��֤����.Text = ""
    SetCtrVisibleAndMove
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt���֤��, KeyAscii
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
    
    If mblnɨ�����֤ǩԼ = True And txt���֤��.Text = "" Then
        OpenIDCard
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        '�²��˲ŵ���
        If mblnNameChange = True And mlng����ID = 0 Then zlQueryEMPIPatiInfo
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

Private Sub txt���֤��_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
    
End Sub

Private Sub txt���֤��_Validate(Cancel As Boolean)
    '65663:������,2014-02-20,�������֤�ż����������
    If IsDate(zlCommFun.GetIDCardDate(txt���֤��.Text)) = False Then Exit Sub
    If Format(zlCommFun.GetIDCardDate(txt���֤��.Text), "yyyy-mm-dd") <> Format(txt��������.Text, "yyyy-mm-dd") Then
        If IsDate(txt��������.Text) Then MsgBox "��������֤��������ĳ������ڲ�һ�£���ʹ�����֤�Ż�ȡ�������滻��", vbInformation, gstrSysName
        txt��������.Text = zlCommFun.GetIDCardDate(txt���֤��.Text)
    End If
End Sub

Private Sub txt��֤_Change()
    If gCurSendCard.str������ = "�������֤" Then
        txt��֤����.Text = txt��֤.Text
    End If
End Sub

Private Sub txt��֤_GotFocus()
    Call zlControl.TxtSelAll(txt��֤)
    Call OpenPassKeyboard(txt��֤, True)
End Sub

Private Function GetOneDept(lng�շ�ϸĿID As Long) As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.ִ�п���ID From �շ���ĿĿ¼ A,�շ�ִ�п��� B Where B.�շ�ϸĿID=A.ID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�շ�ϸĿID)
    If Not rsTmp.EOF Then
        GetOneDept = rsTmp!ִ�п���ID 'Ĭ��ȡ��һ��(���ж��)
    Else
        GetOneDept = UserInfo.����ID '��û��ָ������ȡ����Ա���ڿ���
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������봴��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:59:39
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

Private Function OpenPassKeyboard(ctlText As Control, Optional blnȷ������ As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, blnȷ������) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Sub txt��֤_LostFocus()
    Call ClosePassKeyboard(txt��֤)
End Sub

Private Function GetCardDataSql(ByVal byt�䶯���� As Byte, ByVal lng����ID As Long, ByVal lng�����ID As Long, _
   ByVal strԭ���� As String, ByVal strCard As String, ByVal str���� As String, ByVal dtCurdate As Date, _
   ByVal strICCard As String, Optional ByVal str�䶯ԭ�� As String = "")
    Dim strSQL As String
    Dim strPassWord As String
    strPassWord = zlCommFun.zlStringEncode(str����)
    'Zl_ҽ�ƿ��䶯_Insert
     strSQL = "Zl_ҽ�ƿ��䶯_Insert("
    '      �䶯����_In   Number,
    '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����);
    '��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
    strSQL = strSQL & "" & byt�䶯���� & ","
    '      ����id_In     סԺ���ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
    strSQL = strSQL & "" & lng�����ID & ","
    '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
    strSQL = strSQL & "'" & strԭ���� & "',"
    '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
    strSQL = strSQL & "'" & strCard & "',"
    '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
    '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
    strSQL = strSQL & "'" & str�䶯ԭ�� & "',"
    '      ����_In       ������Ϣ.����֤��%Type,
    strSQL = strSQL & "'" & strPassWord & "',"
    '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '      Ic����_In     ������Ϣ.Ic����%Type := Null,
    strSQL = strSQL & "'" & strICCard & "',"
    '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
    strSQL = strSQL & IIf(str�䶯ԭ�� = "", "NULL)", "'" & str�䶯ԭ�� & "')")
    GetCardDataSql = strSQL
End Function


Private Function zlLoadInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ϣ
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-11-04 10:50:46
    '����:42947
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strOutCardNO As String, strOutPatiInforXML As String
    Dim strExpand As String
    Dim strPatiXml As String
    On Error GoTo errHandle
    
    If gCurSendCard.lng�����ID = 0 Then zlLoadInfor = True: Exit Function
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
   '�����:53408
    If gCurSendCard.str������ <> "�������֤" Then
         mblnɨ�����֤ = True
         If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, gCurSendCard.lng�����ID, False, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Function
    End If
    If strOutCardNO = "" Then Exit Function
    txt����.Text = strOutCardNO
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
    '����:���ز�����Ϣ
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-11-04 10:50:46
    '����:42947
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strOutCardNO As String
    Dim objNode As MSXML2.IXMLDOMElement, strExpand As String
    Dim objTempNode As MSXML2.IXMLDOMElement
    Dim strPatiXml As String
    Dim strTmp As String, strValue As String
    Dim i As Long, j As Long, lngCount As Long, lngChildCount As Long '�����:56599
    Dim str����ҩ�� As String, str������Ӧ As String '�����:56599
    Dim str�������� As String, str�������� As String '�����:56599
    Dim strABOѪ�� As String '�����:56599
    Dim str��Ϣ�� As String, str��Ϣֵ As String '�����:56599
    Dim xmlChildNodes As IXMLDOMNodeList, xmlChildNode As IXMLDOMNode '�����:56599
    Dim str���� As String, str��ϵ As String, str�绰 As String, str���֤�� As String, str��ַ As String '�����:56599
    Dim str������ϵ As String
    On Error GoTo errHandle
    
    If strXmlCardInfor = "" Then LoadPati = True: Exit Function
    '���ز�����Ϣ
    If zlXML_Init = False Then Exit Function
    
    '101269:���ϴ�,2016/10/8,�����������
    If zlXML_LoadXMLToDOMDocument(strXmlCardInfor, False) = False Then Exit Function
    '    ��ʶ    ��������    ����    ����    ˵��
    '    ����    Varchar2    20
    Call zlXML_GetNodeValue("����", , strValue)
    '    ����    Varchar2    100
    Call zlXML_GetNodeValue("����", , strValue)
    txtPatient.Text = strValue
    '    �Ա�    Varchar2    4
    Call zlXML_GetNodeValue("�Ա�", , strValue)
    If strValue <> "" Then
        Call zlControl.CboLocate(cbo�Ա�, strValue)
        If cbo�Ա�.ListIndex = -1 Then
            cbo�Ա�.AddItem strValue
            cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
        End If
    End If
    '    ����    Varchar2    10
    Call zlXML_GetNodeValue("����", , strValue)
    If strValue <> "" Then
        Call LoadOldData(strValue, txt����, cbo���䵥λ)
    End If
    '    ��������    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    Call zlXML_GetNodeValue("��������", , strValue)
    If strValue <> "" Then
        txt��������.Text = Format(IIf(IsDate(strValue) = False, "____-__-__", strValue), "YYYY-MM-DD")
        If IsDate(strValue) Then txt����ʱ�� = Format(CDate(strValue), "HH:MM")
        txt����.Text = ReCalcOld(CDate(txt��������.Text), cbo���䵥λ)      '�޸ĵ�ʱ��,���ݳ���������������
        txt����.Tag = txt����.Text
    Else
         txt����ʱ��.Text = "__:__"
         txt��������.Text = ReCalcBirth(Val(txt����.Text), cbo���䵥λ.Text)
    End If
    cbo���䵥λ.Tag = cbo���䵥λ.Text
    
    '    �����ص�    Varchar2    50
    Call zlXML_GetNodeValue("�����ص�", , strValue)
    '    ���֤��    VARCHAR2    18
    Call zlXML_GetNodeValue("���֤��", , strValue)
    If strValue <> "" Then
        txt���֤��.Text = strValue
        If InStr(1, txt��������.Text, "__") > 0 Then
            strTmp = zlCommFun.GetIDCardDate(txt���֤��.Text)
            If IsDate(strTmp) Then txt��������.Text = strTmp
        End If
    End If
    '    ����֤��    Varchar2    20
    Call zlXML_GetNodeValue("����֤��", , strValue)
   ' If strValue <> "" Then txt����֤��.Text = strValue
    '    ְҵ    Varchar2    80
    Call zlXML_GetNodeValue("ְҵ", , strValue)
    If strValue <> "" Then
        cboְҵ.ListIndex = cbo.FindIndex(cboְҵ, strValue)
        If cboְҵ.ListIndex = -1 Then
            cboְҵ.AddItem strValue, 0
            cboְҵ.ListIndex = cboְҵ.NewIndex
        End If
    End If
    '    ����    Varchar2    20
    Call zlXML_GetNodeValue("����", , strValue)
    cbo����.ListIndex = cbo.FindIndex(cbo����, strValue, True)
     If cbo����.ListIndex = -1 And strValue <> "" Then
         cbo����.AddItem strValue, 0
         cbo����.ListIndex = cbo����.NewIndex
     End If
    '    ����    Varchar2    30
    Call zlXML_GetNodeValue("����", , strValue)
    cbo����.ListIndex = cbo.FindIndex(cbo����, strValue, True)
     If cbo����.ListIndex = -1 And strValue <> "" Then
         cbo����.AddItem strValue, 0
         cbo����.ListIndex = cbo����.NewIndex
     End If
    '    ѧ��    Varchar2    10
    Call zlXML_GetNodeValue("ѧ��", , strValue)
    'cboѧ��.ListIndex = GetCboIndex(cboѧ��, strValue)
'    If cboѧ��.ListIndex = -1 And strValue <> "" Then
'        cboѧ��.AddItem strValue, 0
'        cboѧ��.ListIndex = cboѧ��.NewIndex
'    End If
    '    ����״��    Varchar2    4
    Call zlXML_GetNodeValue("����״��", , strValue)
    cbo����.ListIndex = cbo.FindIndex(cbo����, strValue, True)
     If cbo����.ListIndex = -1 And strValue <> "" Then
         cbo����.AddItem strValue, 0
         cbo����.ListIndex = cbo����.NewIndex
     End If
    '    ����    Varchar2    30
    Call zlXML_GetNodeValue("����", , strValue)
    txt����.Text = strValue
    '    ��ͥ��ַ    Varchar2    50
    Call zlXML_GetNodeValue("��ͥ��ַ", , strValue)
   cbo��ͥ��ַ.Text = strValue
   padd��ͥ��ַ.Value = strValue
    '    ���ڵ�ַ    Varchar2    50
    Call zlXML_GetNodeValue("���ڵ�ַ", , strValue)
    txtRegLocation.Text = strValue
    padd���ڵ�ַ.Value = strValue
    '    ��ͥ�绰    Varchar2    20
    Call zlXML_GetNodeValue("��ͥ�绰", , strValue)
   txt��ͥ�绰.Text = strValue
    '    ��ͥ��ַ�ʱ�    Varchar2    6
    Call zlXML_GetNodeValue("��ͥ��ַ�ʱ�", , strValue)
   txt��ͥ�ʱ�.Text = strValue
    '    �໤��  Varchar2    64
    Call zlXML_GetNodeValue("�໤��", , strValue)
   'txt�໤��.Text = strValue
'    '    ��ϵ������  Varchar2    64
'    Call zlXML_GetNodeValue("��ϵ������", , strValue)
'    txt��ϵ������.Text = strValue '�����:40005
'    '    ��ϵ�˹�ϵ  Varchar2    30
'    Call zlXML_GetNodeValue("��ϵ�˹�ϵ", , strValue)
'    txt��ϵ�˹�ϵ.Text = strValue '�����:40005
'    '    ��ϵ�˵�ַ  Varchar2    50
'    Call zlXML_GetNodeValue("��ϵ�˵�ַ", , strValue)
'    '    ��ϵ�˵绰  Varchar2    20
'    Call zlXML_GetNodeValue("��ϵ�˵绰", , strValue)
'    txt��ϵ�˵绰.Text = strValue '�����:40005
    '    ������λ    Varchar2    100
    Call zlXML_GetNodeValue("������λ", , strValue)
    txt��λ����.Text = strValue
    lbl��λ����.Tag = ""
    '    ��λ�绰    Varchar2    20
    Call zlXML_GetNodeValue("��λ�绰", , strValue)
   txt��λ�绰.Text = strValue
    '    ��λ�ʱ�    Varchar2    6
    Call zlXML_GetNodeValue("��λ�ʱ�", , strValue)
   txt��λ�ʱ�.Text = strValue
    '    ��λ������  Varchar2    50
    Call zlXML_GetNodeValue("��λ������", , strValue)
   'txt��λ������.Text = strValue
    '    ��λ�ʺ�    Varchar2    20
    Call zlXML_GetNodeValue("��λ�ʺ�", , strValue)
   'txt��λ�ʺ�.Text = strValue
   '�����:56599
    '�������
    Call zlXML_GetRows("ҩ������", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("ҩ������", i, str����ҩ��)
        Call zlXML_GetNodeValue("ҩ�ﷴӦ", i, str������Ӧ)
        SetDrugAllergy str����ҩ��, str������Ӧ
    Next
    lngCount = 0
    '���߼�¼
    Call zlXML_GetRows("��������", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("��������", i, str��������)
        Call zlXML_GetNodeValue("����ʱ��", i, str��������)
        SetInoculate str��������, str��������
    Next
    lngCount = 0
    'ABOѪ��
    Call zlXML_GetNodeValue("ABOѪ��", , strABOѪ��)
    If strABOѪ�� <> "" Then
        For i = 0 To cboBloodType.ListCount - 1
            '76314,���ϴ���2014-08-06��������Ϣ��ȷ��ȡ
            If NeedName(cboBloodType.List(i), , ".") = NeedName(strABOѪ��) Then cboBloodType.ListIndex = i
        Next
    End If
    'RH
    Call zlXML_GetNodeValue("RH", , strValue)
    If strValue <> "" Then
        For i = 0 To cboBH.ListCount - 1
            If cboBH.List(i) = strValue Then cboBH.ListIndex = i
        Next
    End If
    'ҽѧ��ʾ
    strValue = ""
    Set xmlChildNodes = zlXML_GetChildNodes("�ٴ�������Ϣ")
    If Not xmlChildNodes Is Nothing Then
        If xmlChildNodes.length > 0 Then
            For i = 0 To xmlChildNodes.length - 1
                Set xmlChildNode = xmlChildNodes(i)
                If xmlChildNode.Text = "1" Then
                    strValue = strValue & "," & Replace(xmlChildNode.nodeName, "��־", "")
                End If
            Next
        End If
    End If
    If strValue <> "" Then txtMedicalWarning.Text = Mid(strValue, 2)
   
    
    '����ҽѧ��ʾ
    Call zlXML_GetNodeValue("����ҽѧ��ʾ", , strValue)
    If strValue <> "" Then txtOtherWaring.Text = strValue
    '��ϵ��Ϣ
    '    ��ϵ�˵�ַ  Varchar2    50
    Call zlXML_GetNodeValue("��ϵ�˵�ַ", , str��ַ)
    'txt��ϵ�˵�ַ.Text = str��ַ
     '    ��ϵ������  Varchar2    64
    Call zlXML_GetNodeValue("��ϵ������", , str����)
    '    ��ϵ�˹�ϵ  Varchar2    30
    Call zlXML_GetNodeValue("��ϵ�˹�ϵ", , str��ϵ)
    '    ��ϵ�˵绰  Varchar2    20
    Call zlXML_GetNodeValue("��ϵ�˵绰", , str�绰)
    '    ��ϵ�����֤ Varchar2   20
    Call zlXML_GetNodeValue("��ϵ�����֤��", , str���֤��)
    '84313,���ϴ�,2015/4/27,��ϵ�˹�ϵ�Լ�������ϵ
    Call zlXML_GetNodeValue("��ϵ�˸�����Ϣ", , str������ϵ)
    SetLinkInfo str����, str��ϵ, str�绰, str���֤��, str������ϵ
    
    Call zlXML_GetRows("��ϵ��Ϣ", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("��ϵ��Ϣ", "����", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "����", i, j, str����)
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "��ϵ", i, j, str��ϵ)
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "�绰", i, j, str�绰)
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "���֤��", i, j, str���֤��)
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "������Ϣ", i, j, str������ϵ)
                SetLinkInfo str����, str��ϵ, str�绰, str���֤��, str������ϵ
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0

    '������Ϣ
    '�����������
    Call zlXML_GetNodeValue("�����������", , strValue)
    SetOtherInfo "�����������", strValue
    
    '��ũ��֤��
    Call zlXML_GetNodeValue("��ũ��֤��", , strValue)
    SetOtherInfo "��ũ��֤��", strValue

    '����֤��
    Call zlXML_GetRows("����֤��", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("����֤��", "��Ϣ��", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("����֤��", "��Ϣ��", i, j, str��Ϣ��)
                Call zlXML_GetChildNodeValue("����֤��", "��Ϣֵ", i, j, str��Ϣֵ)
                SetOtherInfo str��Ϣ��, str��Ϣֵ
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    '������Ϣ
    Call zlXML_GetRows("������Ϣ", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("������Ϣ", "��Ϣ��", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("������Ϣ", "��Ϣ��", i, j, str��Ϣ��)
                Call zlXML_GetChildNodeValue("������Ϣ", "��Ϣֵ", i, j, str��Ϣֵ)
                SetOtherInfo str��Ϣ��, str��Ϣֵ
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    'ҽ�ƿ�����
    Call zlXML_GetRows("ҽ�ƿ�����", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("ҽ�ƿ�����", "��Ϣ��", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("ҽ�ƿ�����", "��Ϣ��", i, j, str��Ϣ��)
                Call zlXML_GetChildNodeValue("ҽ�ƿ�����", "��Ϣֵ", i, j, str��Ϣֵ)
                If mdicҽ�ƿ�����.Exists(str��Ϣ��) Then
                    mdicҽ�ƿ�����.Item(str��Ϣ��) = str��Ϣֵ
                Else
                    mdicҽ�ƿ�����.Add str��Ϣ��, str��Ϣֵ
                End If
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    
    '�ӿ��ϻ�ȡ������Ϣ��,����EMPI
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
    '����:�ر�������������
    '����:���˺�
    '����:2012-03-09 16:26:40
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
    '����:��ʼ���µĿ�����
    '����:���˺�
    '����:2012-03-09 16:28:23
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
    '����:�����֤������
    '����:����
    '����:2012-08-31 16:28:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '��ʼ���Կ�����
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.Hwnd)
    End If
    '�򿪶�����
    mobjIDCard.SetEnabled (True)
End Sub

Public Function zl_Get����Ĭ�Ϸ�������() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ĭ�Ϸ�������
    '����:�Ƿ������������
    '����:����
    '����:2012-07-06 15:53:14
    '�����:51072
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCardType As clsCard
    Dim msgResult As VbMsgBoxResult
    Dim arr() As String
    arr = zl_Getҽ�ƿ�����(gCurSendCard.lng�����ID)
    If Val(arr(2)) = 0 Then '������
        Select Case Val(arr(1))
            Case 0 '������
                zl_Get����Ĭ�Ϸ������� = True
                Exit Function
            Case 1 'δ��������
               msgResult = MsgBox("δ�������뽫��Ӱ���ʻ���ʹ�ð�ȫ,�Ƿ������", vbQuestion + vbYesNo, gstrSysName)
               zl_Get����Ĭ�Ϸ������� = IIf(msgResult = vbYes, True, False)
               Exit Function
            Case 2 'Ϊ�����ֹ
                 MsgBox "δ���뿨����,���ܽ��з�����", vbExclamation, gstrSysName
                zl_Get����Ĭ�Ϸ������� = False
                Exit Function
        End Select
    ElseIf Val(arr(2)) = 1 Then 'ȱʡ���֤��Nλ
        If Len(Trim(txt���֤��.Text)) > 0 Or Len(Trim(txt��ϵ�����֤.Text)) > 0 Then '���������֤����ϵ�����֤��
            If Len(Trim(txt���֤��.Text)) > 0 Then '�����֤���������֤
                   txt����.Text = Right(Trim(txt���֤��.Text), Val(arr(0)))
            Else '������ô��������֤��Ϊ����
                   txt����.Text = Right(Trim(txt��ϵ�����֤.Text), Val(arr(0)))
            End If
        Else '���֤����ϵ�����֤��û����
            Select Case Val(arr(1))
                Case 0 '������
                    zl_Get����Ĭ�Ϸ������� = True
                    Exit Function
                Case 1 'δ��������
                    msgResult = MsgBox("δ�������뽫��Ӱ���ʻ���ʹ�ð�ȫ,�Ƿ������", vbQuestion + vbYesNo, gstrSysName)
                    zl_Get����Ĭ�Ϸ������� = IIf(msgResult = vbYes, True, False)
                    Exit Function
                Case 2 'Ϊ�����ֹ
                    MsgBox "δ���뿨����,���ܽ��з�����", vbExclamation, gstrSysName
                    zl_Get����Ĭ�Ϸ������� = False
                    Exit Function
            End Select
        End If
    End If
    zl_Get����Ĭ�Ϸ������� = True
End Function

Private Function zl_Getȱʡ�������() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡȱʡ�������
    '����:ȱʡ�����������
    '����:����
    '����:2012-08-31 11:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lngCardTypeID As Long
    Dim rsTemp As Recordset
    
    On Error GoTo ErrHandl:

    strSQL = "" & _
    "   Select Id, ����, ����, ����, ǰ׺�ı�, ���ų���, ȱʡ��־, �Ƿ�̶�, �Ƿ��ϸ����, " & _
    "           nvl(�Ƿ�����,0) as �Ƿ�����, nvl(�Ƿ�����ʻ�,0) as �Ƿ�����ʻ�, " & _
    "           nvl(�Ƿ�ȫ��,0) as �Ƿ�ȫ��,nvl(�Ƿ��ظ�ʹ��,0) as �Ƿ��ظ�ʹ�� , " & _
    "           nvl(���볤��,10) as ���볤��,nvl(���볤������,0) as ���볤������,nvl(�������,0) as �������," & _
    "           nvl(�Ƿ�����,0) as �Ƿ�����,����, ��ע, �ض���Ŀ, ���㷽ʽ, �Ƿ�����, ��������,Nvl(������������,0) as ������������,Nvl(�Ƿ�ȱʡ����,0) as �Ƿ�ȱʡ����," & _
    "           nvl(�Ƿ�ģ������,0) as �Ƿ�ģ������,nvl(��������,'1000') as �������� " & _
    "    From ҽ�ƿ����" & _
    "    Where ID = [1]" & _
    "    Order by ����"

    lngCardTypeID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, , , True))
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCardTypeID)
    If rsTemp Is Nothing Then zl_Getȱʡ������� = "": Exit Function
    If rsTemp.RecordCount <= 0 Then zl_Getȱʡ������� = "": Exit Function
    zl_Getȱʡ������� = rsTemp!����
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub SetCtrVisibleAndMove()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿ�����ʾ��λ��
    '����:����
    '����:2012-08-31 11:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��������� As String
    Dim lng�仯���� As Long
    Dim lng��� As Long
    
    lng��� = 100
    
    '��ť����
    cmdHelp.Top = Me.ScaleHeight - 500
    cmdCancel.Top = cmdHelp.Top
    cmdOK.Top = cmdHelp.Top
    
    'Ĭ�Ϸ����Ƕ������֤ʱ,�󿨿ؼ�������
    If gCurSendCard.str������ Like "�������֤" Then
        txt����.Enabled = False: txt����.Enabled = False: txt��֤.Enabled = False
        lblICCard.Enabled = False: lbl����.Enabled = False: lbl��֤.Enabled = False
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
        
     'ɨ������֤��ɨ�����֤ǩԼΪTrue������²��ܰ����֤
    If mblnɨ�����֤ = True And mblnɨ�����֤ǩԼ Then
        lbl֧������.Enabled = True: txt֧������.Enabled = True
        lbl��֤����.Enabled = True: txt��֤����.Enabled = True
    Else
        '����֧����������֤���벻����
        lbl֧������.Enabled = False: txt֧������.Enabled = False: txt֧������.Text = ""
        lbl��֤����.Enabled = False: txt��֤����.Enabled = False: txt��֤����.Text = "": txt��֤����.Tag = ""
    End If
End Sub

Private Sub txt��֤����_GotFocus()
    Call zlControl.TxtSelAll(txt��֤����)
    Call OpenPassKeyboard(txt��֤����, False)
End Sub

Private Sub txt��֤����_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, gCurSendCard.int������� = 1)
End Sub

Private Sub txt��֤����_LostFocus()
    Call ClosePassKeyboard(txt��֤����)
End Sub
Private Sub txt֧������_GotFocus()
    Call zlControl.TxtSelAll(txt֧������)
    Call OpenPassKeyboard(txt֧������, False)
End Sub

Private Sub txt֧������_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, gCurSendCard.int������� = 1)
End Sub

Private Sub zl���ز�����Ϣ(rsPatiInfo As Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ϣ������ؼ���
    '����:����
    '����:2012-08-31 11:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    '��������
    txtPatient.Text = Nvl(rsPatiInfo!����)
    mlng����ID = Nvl(Val(rsPatiInfo!����ID))
    If Nvl(rsPatiInfo!�Ա�) <> "" Then
        Call zlControl.CboLocate(cbo�Ա�, rsPatiInfo!�Ա�)
        If cbo�Ա�.ListIndex = -1 Then
            cbo�Ա�.AddItem rsPatiInfo!�Ա�
            cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
        End If
    End If
    '����
    If Nvl(rsPatiInfo!����) <> "" Then
        Call LoadOldData(rsPatiInfo!����, txt����, cbo���䵥λ)
    End If
    '��������
    If Nvl(rsPatiInfo!��������) <> "" Then
        txt��������.Text = Format(IIf(IsDate(rsPatiInfo!��������) = False, "____-__-__", rsPatiInfo!��������), "YYYY-MM-DD")
        If IsDate(rsPatiInfo!��������) Then txt����ʱ�� = Format(CDate(rsPatiInfo!��������), "HH:MM")
        txt����.Text = ReCalcOld(CDate(txt��������.Text), cbo���䵥λ)      '�޸ĵ�ʱ��,���ݳ���������������
        txt����.Tag = txt����.Text
    Else
         txt����ʱ��.Text = "__:__"
         txt��������.Text = ReCalcBirth(Val(txt����.Text), cbo���䵥λ.Text)
    End If
    cbo���䵥λ.Tag = cbo���䵥λ.Text

    '���֤��
    If Nvl(rsPatiInfo!���֤��) <> "" Then
        txt���֤��.Text = rsPatiInfo!���֤��
        If InStr(1, txt��������.Text, "__") > 0 Then
            strTmp = zlCommFun.GetIDCardDate(txt���֤��.Text)
            If IsDate(strTmp) Then txt��������.Text = strTmp
        End If
    End If
    'ְҵ
    If Nvl(rsPatiInfo!ְҵ) <> "" Then
        cboְҵ.ListIndex = cbo.FindIndex(cboְҵ, rsPatiInfo!ְҵ)
        If cboְҵ.ListIndex = -1 Then
            cboְҵ.AddItem rsPatiInfo!ְҵ, 0
            cboְҵ.ListIndex = cboְҵ.NewIndex
        End If
    End If
    '����
    cbo����.ListIndex = cbo.FindIndex(cbo����, Nvl(rsPatiInfo!����), True)
     If cbo����.ListIndex = -1 And Nvl(rsPatiInfo!����) <> "" Then
         cbo����.AddItem rsPatiInfo!����, 0
         cbo����.ListIndex = cbo����.NewIndex
     End If
    '����
    cbo����.ListIndex = cbo.FindIndex(cbo����, Nvl(rsPatiInfo!����), True)
     If cbo����.ListIndex = -1 And Nvl(rsPatiInfo!����) <> "" Then
         cbo����.AddItem rsPatiInfo!����, 0
         cbo����.ListIndex = cbo����.NewIndex
     End If
    '����״��
    cbo����.ListIndex = cbo.FindIndex(cbo����, Nvl(rsPatiInfo!����״��), True)
     If cbo����.ListIndex = -1 And Nvl(rsPatiInfo!����״��) <> "" Then
         cbo����.AddItem rsPatiInfo!����״��, 0
         cbo����.ListIndex = cbo����.NewIndex
     End If
    txt����.Text = Nvl(rsPatiInfo!����)
    '��ͥ��ַ
    cbo��ͥ��ַ.Text = Nvl(rsPatiInfo!��ͥ��ַ)
    Call zlReadAddrInfo(padd��ͥ��ַ, Val(Nvl(rsPatiInfo!����ID)), 0, 3, cbo��ͥ��ַ.Text)
    '��ͥ�绰
    txt��ͥ�绰.Text = Nvl(rsPatiInfo!��ͥ�绰)
    '��ͥ��ַ�ʱ�
    txt��ͥ�ʱ�.Text = Nvl(rsPatiInfo!��ͥ��ַ�ʱ�)
    '���ڵ�ַ
    txtRegLocation.Text = Nvl(rsPatiInfo!���ڵ�ַ)
    Call zlReadAddrInfo(padd���ڵ�ַ, Val(Nvl(rsPatiInfo!����ID)), 0, 4, txtRegLocation.Text)
    '���ڵ�ַ�ʱ�
    txt���ڵ�ַ�ʱ�.Text = Nvl(rsPatiInfo!���ڵ�ַ�ʱ�)
    '������λ
    txt��λ����.Text = Nvl(rsPatiInfo!������λ)
    lbl��λ����.Tag = ""
    '��λ�绰
    txt��λ�绰.Text = Nvl(rsPatiInfo!��λ�绰)
    '��λ�ʱ�
    txt��λ�ʱ�.Text = Nvl(rsPatiInfo!��λ�ʱ�)
    '�����
    txt�����.Text = Nvl(rsPatiInfo!�����)
    '�����:40005
    '��ϵ������
    txt��ϵ������.Text = Nvl(rsPatiInfo!��ϵ������)
    '��ϵ�˵绰
    txt��ϵ�˵绰.Text = Nvl(rsPatiInfo!��ϵ�˵绰)
    '84313,���ϴ�,2015/4/27,��ϵ�˹�ϵ�Լ�������ϵ
    '��ϵ�˹�ϵ
    txt������ϵ.Text = ""
    cbo��ϵ�˹�ϵ.ListIndex = cbo.FindIndex(cbo��ϵ�˹�ϵ, Nvl(rsPatiInfo!��ϵ�˹�ϵ), True)
    If cbo��ϵ�˹�ϵ.ListIndex = -1 And Nvl(rsPatiInfo!��ϵ�˹�ϵ) <> "" Then
        cbo��ϵ�˹�ϵ.ListIndex = 8: txt������ϵ.Text = Nvl(rsPatiInfo!��ϵ�˹�ϵ)
    End If
    '�ֻ���
    txtMobile.Text = Nvl(rsPatiInfo!�ֻ���)
    '�����:56599
    Load�����������Ϣ (Val(Nvl(rsPatiInfo!����ID, "0")))
    '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
    LoadCertificate (Val(Nvl(rsPatiInfo!����ID)))
    
    mstr���� = txt����.Text & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, "")
    mstr�Ա� = NeedName(cbo�Ա�.Text)
    mstr���� = txtPatient.Text
    mstr���֤�� = txt���֤��.Text
    mstr�������� = txt��������.Text
    mstr����ʱ�� = txt����ʱ��.Text
End Sub

Private Sub txt֧������_LostFocus()
    Call ClosePassKeyboard(txt֧������)
End Sub
Public Function zl��ǰ�û����֤�Ƿ��(str���֤�� As String, strName As String, str����� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϵ�ǰ�û����֤�Ƿ��ѱ���
    '���:str���֤��:�������֤ str�����:�����
    '����:True �Ѱ� false δ��
    '����:����
    '����:2012-08-31 04:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo Errhand
    strSQL = "" & _
    " Select  ����,����� From ������Ϣ A,����ҽ�ƿ���Ϣ B Where A.����ID=B.����ID And B.����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ҽ�ƿ���", str���֤��)
    If rsTemp Is Nothing Then zl��ǰ�û����֤�Ƿ�� = False: Exit Function
    If rsTemp.RecordCount <= 0 Then zl��ǰ�û����֤�Ƿ�� = False: Exit Function
    
    If IIf(IsNull(rsTemp!����), "", rsTemp!����) = strName And IIf(IsNull(rsTemp!�����), "", rsTemp!�����) = str����� Then
        zl��ǰ�û����֤�Ƿ�� = True
    Else
        zl��ǰ�û����֤�Ƿ�� = False
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function
Private Sub Init����ҩ��()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ҩ��FlexGrid
    '���:
    '����:
    '����:����
    '����:2012-12-20 04:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '�������Լ��б���
    With msh����
        .Cols = 2
        .TextMatrix(0, 0) = "����ҩ��"
        .TextMatrix(0, 1) = "������Ӧ"
        .ColWidth(0) = 5000
        .ColWidth(1) = .Width - 4900
        '75286:���ϴ���2014-7-16�������뷽ʽ
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
    End With
End Sub

Private Sub InitTagPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ҳ�ؼ�
    '����:56599
    '����:2012-12-20 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    
    Err = 0: On Error GoTo Errhand:

        Set ObjItem = tbcPage.InsertItem(mPageIndex.����, "����", picInfo.Hwnd, 0)
        ObjItem.Tag = mPageIndex.����
    
        Set ObjItem = tbcPage.InsertItem(mPageIndex.��������, "��������", PicHealth.Hwnd, 0)
        ObjItem.Tag = mPageIndex.��������
        Call InitVsInoculate
        Call InitVsOtherInfo
        Call InitCombox
        Call InitCertificate
        
        '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
        If CreatePlugInOK(mlngModul) Then
            On Error Resume Next
            mlngPlugInHwnd = gobjPlugIn.GetFormHwnd
            Call zlPlugInErrH(Err, "GetFormHwnd")
            Err.Clear: On Error GoTo 0
            If mlngPlugInHwnd <> 0 Then
                picTaskPanelOther.Visible = True
                Set ObjItem = tbcPage.InsertItem(mPageIndex.������Ϣ, "������Ϣ", picTaskPanelOther.Hwnd, 0)
                ObjItem.Tag = mPageIndex.������Ϣ
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

Public Sub Add�����������Ϣ(ByVal lng����ID As Long, ByRef colPro As Collection, Optional ByVal lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ݴ���
    '���:
    '����:56599
    '����:2012-12-13 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim strTemp() As String
    Dim strSQL As String
    Dim varKey As Variant
    Dim intCount As Integer
    '����ҩ��
    With msh����
        If .Rows > 1 Then
            '����ò������м�¼
            strSQL = " Zl_���˹���ҩ��_Delete(" & lng����ID & ")"
            zlAddArray colPro, strSQL
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    '���˹���ҩ��
                    strSQL = "Zl_���˹���ҩ��_Update("
                    '����ID_In ���˹���ҩ��.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '����ҩ��ID_In ���˹���ҩ��.����ҩ��ID%Type
                    strSQL = strSQL & "'" & IIf(.RowData(i) <= 0, "", .RowData(i)) & "',"
                    '����ҩ��_In  ���˹���ҩ��.����ҩ��%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 0) = "", "", .TextMatrix(i, 0)) & "',"
                    '������Ӧ_In ���˹�����Ӧ.������Ӧ%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"

                    zlAddArray colPro, strSQL
                End If
            Next
        End If
    End With
    '������Ϣ
    With vsInoculate
        If .Rows > 1 Then
            '����ò������м�¼
            strSQL = " Zl_�������߼�¼_Delete(" & lng����ID & ")"
            zlAddArray colPro, strSQL

            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    '���˹���ҩ��
                    strSQL = "Zl_�������߼�¼_Update("
                    '����ID_In �������߼�¼.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '����ʱ��_In �������߼�¼.����ʱ��%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 0) = "", "''", "to_date('" & .TextMatrix(i, 0) & "','yyyy-mm-dd')") & ","
                    '��������_In  �������߼�¼.��������%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"
                    zlAddArray colPro, strSQL
                End If
                If .TextMatrix(i, 3) <> "" Then
                    '���˹���ҩ��
                    strSQL = "Zl_�������߼�¼_Update("
                    '����ID_In �������߼�¼.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '����ʱ��_In �������߼�¼.����ʱ��%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 2) = "", "''", "to_date('" & .TextMatrix(i, 2) & "','yyyy-mm-dd')") & ","
                    '��������_In  �������߼�¼.��������%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "''", .TextMatrix(i, 3)) & "')"
                    zlAddArray colPro, strSQL
                End If
            Next
        End If
    End With
    '������Ϣ
    'ABOѪ��
    '������Ϣ�ӱ�
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'Ѫ��',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & NeedName(cboBloodType.Text, , ".") & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    'RH
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'RH',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & cboBH.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    'ҽѧ��ʾ
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'ҽѧ��ʾ',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & txtMedicalWarning.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    '����ҽѧ��ʾ
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'����ҽѧ��ʾ',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & txtOtherWaring.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
        
    '84313:���ϴ�,2015/4/29, ��һ����ϵ����Ϣ�ѱ����ڲ�����Ϣ�У��ӱ��в����ظ�����
    '��ϵ�������Ϣ
    intCount = 0
    With vsLinkMan
        If .Rows >= 3 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then '��ϵ����������Ϊ��
                    intCount = intCount + 1
                    For j = 0 To .Cols - 1
                        strSQL = "Zl_������Ϣ�ӱ�_Update("
                        '����ID_In ������Ϣ�ӱ�.����Id%Type
                        strSQL = strSQL & "" & lng����ID & ","
                        '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
                        strSQL = strSQL & "'��ϵ��" & .TextMatrix(0, j) & intCount & "',"
                        '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
                        strSQL = strSQL & "'" & IIf(.TextMatrix(i, j) = "", "", .TextMatrix(i, j)) & "',"
                        '����Id_In ������Ϣ�ӱ�.����Id%Type
                        strSQL = strSQL & "'')"

                        zlAddArray colPro, strSQL
                    Next
                End If
            Next
        End If
    End With
    '������Ϣ
     With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    strSQL = "Zl_������Ϣ�ӱ�_Update("
                    '����ID_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 0) & "',"
                    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "',"
                    '����Id_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "'')"

                    zlAddArray colPro, strSQL
                End If
                If .TextMatrix(i, 2) <> "" Then
                    strSQL = "Zl_������Ϣ�ӱ�_Update("
                    '����ID_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 2) & "',"
                    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "", .TextMatrix(i, 3)) & "',"
                    '����Id_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "'')"

                    zlAddArray colPro, strSQL
                End If
            Next
        End If
     End With
     'ҽ�ƿ�����
     If Not mdicҽ�ƿ����� Is Nothing Then
        For Each varKey In mdicҽ�ƿ�����.Keys
            strSQL = "Zl_����ҽ�ƿ�����_Update("
            strSQL = strSQL & lng����ID & ","
            strSQL = strSQL & gCurSendCard.lng�����ID & ","
            strSQL = strSQL & "'" & Trim(txt����.Text) & "',"
            strSQL = strSQL & "'" & varKey & "',"
            strSQL = strSQL & "'" & mdicҽ�ƿ�����(varKey) & "')"
            zlAddArray colPro, strSQL
        Next
     End If
     If lng����ID = 0 Then Exit Sub
     'ABOѪ��
    '������Ϣ�ӱ�
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'Ѫ��',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & NeedName(cboBloodType.Text, , ".") & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & lng����ID & ")"
    zlAddArray colPro, strSQL
    'RH
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'RH',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & cboBH.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & lng����ID & ")"
    zlAddArray colPro, strSQL
End Sub
Private Sub InitVsInoculate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSGrid�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsInoculate
    '��ʼ���б�����
     vsInoculate.Editable = flexEDKbdMouse
    '������ͷ
        SetColumHeader vsInoculate, C_InoculateHeader
    '����ѡ��ť
        .ColDataType(0) = flexDTDate
        .ColEditMask(0) = "####-##-##"
        .ColDataType(2) = flexDTDate
        .ColEditMask(2) = "####-##-##"
    End With

End Sub
Private Sub InitVsOtherInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSGrid�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, str��ϵ As String
    
    With vsLinkMan
    '��ʼ���б�����
        .Editable = flexEDKbd
    '������ͷ
        SetColumHeader vsLinkMan, C_LinkManColumHeader
        For i = 0 To cbo��ϵ�˹�ϵ.ListCount - 1
            str��ϵ = str��ϵ & "|" & NeedName(cbo��ϵ�˹�ϵ.List(i))
        Next
        str��ϵ = Mid(str��ϵ, 2)
        If str��ϵ <> "" Then .ColComboList(.ColIndex("��ϵ")) = str��ϵ
    End With
    With vsOtherInfo
         .Editable = flexEDKbd
    '������ͷ
        SetColumHeader vsOtherInfo, C_OtherInfoColumHeader
    End With
End Sub

Private Sub InitCombox()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ComBox�ؼ�
    '����:56599
    '����:2012-12-07 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '66743:������,2013-11-25,Ѫ����RHĬ��ֵ������
    'ComboBox cboBloodType, C_Ѫ��
    zlComboxLoadFromSQL "Select ����,����,ȱʡ��־ From Ѫ��", cboBloodType
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
    '����:������ͷ
    '����:objList - ���ö���,strColumHeader - �б������ַ���
    '����:56599
    '����:2012-12-05 11:39:36
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
            '�ݲ�����
        End Select
    Next
End Sub
Private Sub SetDrugAllergy(str����ҩ�� As String, str������Ӧ As String, Optional lng����ID = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���ҩ��
    '����:56599
    '����:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    With msh����
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = str����ҩ�� Then
                    .TextMatrix(i, 1) = str������Ӧ
                    If lng����ID <> 0 Then .RowData(i) = lng����ID
                    Exit Sub
                End If
            Next
        End If
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str����ҩ��
        .TextMatrix(.Rows - 1, 1) = str������Ӧ
        If lng����ID <> 0 Then .RowData(.Rows - 1) = lng����ID
        .Rows = .Rows + 1
    End With
End Sub
Private Sub SetInoculate(str�������� As String, str�������� As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ý������
    '����:56599
    '����:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    With vsInoculate
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 1 To .Cols - 1 Step 2
                    If .TextMatrix(i, j) = str�������� Then
                        .TextMatrix(i, j - 1) = str��������
                        Exit Sub
                    End If
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = str��������
                .TextMatrix(.Rows - 1, j + 1) = str��������
                Exit Sub
            End If
        Next
        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        
    End With
End Sub
Private Sub SetLinkInfo(str���� As String, str��ϵ As String, str�绰 As String, str���֤�� As String, str������ϵ As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ϵ�������Ϣ
    '����:56599
    '����:2012-12-12 09:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    '84313,���ϴ�,2015/4/27,��ϵ�˹�ϵ�Լ�������ϵ
    With vsLinkMan
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = str���� And .TextMatrix(i, 2) = str���֤�� Then
                    .TextMatrix(i, 1) = str��ϵ: .TextMatrix(i, 3) = str�绰
                    If i = 1 Then
                        txt��ϵ�����֤.Text = str���֤��
                        txt��ϵ������.Text = str����
                        cbo��ϵ�˹�ϵ.ListIndex = cbo.FindIndex(cbo��ϵ�˹�ϵ, str��ϵ, True)
                        If cbo��ϵ�˹�ϵ.ListIndex = -1 And str��ϵ <> "" Then
                            cbo��ϵ�˹�ϵ.ListIndex = 8: txt������ϵ.Text = str��ϵ
                        ElseIf cbo��ϵ�˹�ϵ.ListIndex = 8 Then
                            txt������ϵ.Text = str������ϵ
                        End If
                        txt��ϵ�˵绰.Text = str�绰
                    End If
                    Exit Sub
                End If
            Next
        End If
        
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str����
        If cbo.FindIndex(cbo��ϵ�˹�ϵ, str��ϵ, True) = -1 And str��ϵ <> "" Then
            .TextMatrix(.Rows - 1, 1) = "����": .TextMatrix(.Rows - 1, 4) = str��ϵ
        Else
            .TextMatrix(.Rows - 1, 1) = str��ϵ
            .TextMatrix(.Rows - 1, 4) = str������ϵ
        End If
        .TextMatrix(.Rows - 1, 3) = str�绰
        .TextMatrix(.Rows - 1, 2) = str���֤��
        If .Rows - 1 = 1 Then
            txt��ϵ�����֤.Text = str���֤��
            txt��ϵ������.Text = str����
            cbo��ϵ�˹�ϵ.ListIndex = cbo.FindIndex(cbo��ϵ�˹�ϵ, str��ϵ, True)
            If cbo��ϵ�˹�ϵ.ListIndex = -1 And str��ϵ <> "" Then
                cbo��ϵ�˹�ϵ.ListIndex = 8: txt������ϵ.Text = str��ϵ
            ElseIf cbo��ϵ�˹�ϵ.ListIndex = 8 Then
                txt������ϵ.Text = str������ϵ
            End If
            txt��ϵ�˵绰.Text = str�绰
        End If
        .Rows = .Rows + 1
    End With
End Sub
Private Sub SetOtherInfo(str��Ϣ�� As String, str��Ϣֵ As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:56599
    '����:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1 Step 2
                    If .TextMatrix(i, j) = str��Ϣ�� Then
                        .TextMatrix(i, j + 1) = str��Ϣֵ
                        Exit Sub
                    End If
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = str��Ϣ��
                .TextMatrix(.Rows - 1, j + 1) = str��Ϣֵ
                Exit Sub
            End If
        Next
        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        
    End With
End Sub
Public Sub Load�����������Ϣ(lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز��˽�������Ϣ
    '���:lng����ID - ����ID
    '����:56599
    '����:2012-12-12 14:55:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs����ҩ�� As Recordset
    Dim rs���߼�¼ As Recordset
    Dim rsABOѪ�� As Recordset
    Dim rsRH As Recordset
    Dim rsҽѧ��ʾ As Recordset
    Dim rs����ҽѧ��ʾ As Recordset
    Dim rs������Ϣ As Recordset
    Dim rs��ϵ�� As Recordset
    Dim rs������Ϣ As Recordset
    Dim strҽѧ��ʾ As String
    Dim str��ϵ������ As String
    Dim str��ϵ�˹�ϵ As String
    Dim str��ϵ�˵绰 As String
    Dim str��ϵ�����֤�� As String
    Dim str������Ϣ As String
    Dim lng��ϵ������ As Long
    Dim i As Long
    On Error GoTo ErrHandl:
    
    '74430,Ƚ����,2014-7-7,�Һ��еĲ�����Ϣ�༭�������ṩ�ɼ���Ƭ����
    Call ReadPatPricture(lng����ID)
    
    '��ȡ����ҩ��
    strSQL = "" & _
    "   Select ����ID,����ҩ��ID,����ҩ��,������Ӧ From ���˹���ҩ�� Where ����ID=[1]"
    Set rs����ҩ�� = zlDatabase.OpenSQLRecord(strSQL, "���˹���ҩ��", lng����ID)
    While rs����ҩ��.EOF = False
        SetDrugAllergy Nvl(rs����ҩ��!����ҩ��), Nvl(rs����ҩ��!������Ӧ), Nvl(rs����ҩ��!����ҩ��ID, 0)
        rs����ҩ��.MoveNext
    Wend
    '��ȡ���߼�¼
    strSQL = "" & _
    "   Select ����ID,����ʱ��,�������� From �������߼�¼ Where ����ID=[1]"
    Set rs���߼�¼ = zlDatabase.OpenSQLRecord(strSQL, "�������߼�¼", lng����ID)
    While rs���߼�¼.EOF = False
        SetInoculate Nvl(rs���߼�¼!����ʱ��), Nvl(rs���߼�¼!��������)
        rs���߼�¼.MoveNext
    Wend
    'Ѫ��
    strSQL = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='Ѫ��' And ����ID Is NULL"
    Set rsABOѪ�� = zlDatabase.OpenSQLRecord(strSQL, "ABOѪ��", lng����ID)
    While rsABOѪ��.EOF = False
        For i = 0 To cboBloodType.ListCount - 1
            '76314,���ϴ���2014-08-06��������Ϣ��ȷ��ȡ
            If NeedName(cboBloodType.List(i), , ".") = NeedName(Nvl(rsABOѪ��!��Ϣֵ)) Then cboBloodType.ListIndex = i
        Next
        rsABOѪ��.MoveNext
    Wend
    'RH
    strSQL = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='RH' And ����ID Is NULL"
    Set rsRH = zlDatabase.OpenSQLRecord(strSQL, "RH", lng����ID)
    While rsRH.EOF = False
        For i = 0 To cboBH.ListCount - 1
            If cboBH.List(i) = Nvl(rsRH!��Ϣֵ) Then cboBH.ListIndex = i
        Next
        rsRH.MoveNext
    Wend
    'ҽѧ��ʾ
    strSQL = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='ҽѧ��ʾ'"
    Set rsҽѧ��ʾ = zlDatabase.OpenSQLRecord(strSQL, "ҽѧ��ʾ", lng����ID)
    While rsҽѧ��ʾ.EOF = False
        strҽѧ��ʾ = strҽѧ��ʾ & "," & Nvl(rsҽѧ��ʾ!��Ϣֵ)
        rsҽѧ��ʾ.MoveNext
    Wend
    If strҽѧ��ʾ <> "" Then strҽѧ��ʾ = Mid(strҽѧ��ʾ, 2)
    txtMedicalWarning.Text = strҽѧ��ʾ
    '����ҽѧ��ʾ
    strSQL = "" & _
    "  Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='����ҽѧ��ʾ'"
    Set rs����ҽѧ��ʾ = zlDatabase.OpenSQLRecord(strSQL, "����ҽѧ��ʾ", lng����ID)
    While rs����ҽѧ��ʾ.EOF = False
        txtOtherWaring.Text = Nvl(rs����ҽѧ��ʾ!��Ϣֵ)
        rs����ҽѧ��ʾ.MoveNext
    Wend
    '��ϵ�������Ϣ
    'ȡ������Ϣ���е���ϵ����Ϣ
    '84313,���ϴ�,2015/4/27,��ϵ�˹�ϵ�Լ�������ϵ
    strSQL = "" & _
    "   Select  A.��ϵ������,A.��ϵ�˹�ϵ,A.��ϵ�˵绰,A.��ϵ�����֤��,B.��Ϣֵ as ������Ϣ From ������Ϣ A,������Ϣ�ӱ� B " & _
    "   Where A.����ID=B.����ID(+) And A.����ID=[1] And B.��Ϣ��(+)='��ϵ�˸�����Ϣ' And Not A.��ϵ������ is Null"
    Set rs������Ϣ = zlDatabase.OpenSQLRecord(strSQL, "������Ϣ��ϵ����Ϣ", lng����ID)
    If rs������Ϣ.EOF = False Then
        txt��ϵ�����֤.Text = Nvl(rs������Ϣ!��ϵ�����֤��)
        txt��ϵ������.Text = Nvl(rs������Ϣ!��ϵ������)
        txt��ϵ�˵绰.Text = Nvl(rs������Ϣ!��ϵ�˵绰)
        cbo��ϵ�˹�ϵ.ListIndex = cbo.FindIndex(cbo��ϵ�˹�ϵ, Nvl(rs������Ϣ!��ϵ�˹�ϵ), True)
        If cbo��ϵ�˹�ϵ.ListIndex = -1 And Nvl(rs������Ϣ!��ϵ�˹�ϵ) <> "" Then
            cbo��ϵ�˹�ϵ.ListIndex = 8: txt������ϵ.Text = rs������Ϣ!��ϵ�˹�ϵ
        ElseIf cbo��ϵ�˹�ϵ.ListIndex = 8 Then
            txt������ϵ.Text = Nvl(rs������Ϣ!������Ϣ)
        End If
        SetLinkInfo Nvl(rs������Ϣ!��ϵ������), Nvl(rs������Ϣ!��ϵ�˹�ϵ), Nvl(rs������Ϣ!��ϵ�˵绰), Nvl(rs������Ϣ!��ϵ�����֤��), txt������ϵ.Text
    End If
    'ȡ������Ϣ�ӱ��е���ϵ����Ϣ
    strSQL = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ�� like '��ϵ��%' order by ��Ϣ�� Asc"
    Set rs��ϵ�� = zlDatabase.OpenSQLRecord(strSQL, "��ϵ�������Ϣ", lng����ID)
    If rs��ϵ��.EOF = False Then
        rs��ϵ��.Filter = "��Ϣ�� like '��ϵ������%'"
        lng��ϵ������ = rs��ϵ��.RecordCount
        rs��ϵ��.Filter = ""
        For i = 1 To lng��ϵ������ + 1
            While rs��ϵ��.EOF = False
                Select Case Nvl(rs��ϵ��!��Ϣ��)
                    Case "��ϵ������" & i
                        str��ϵ������ = Nvl(rs��ϵ��!��Ϣֵ)
                    Case "��ϵ�˹�ϵ" & i
                        str��ϵ�˹�ϵ = Nvl(rs��ϵ��!��Ϣֵ)
                    Case "��ϵ�˵绰" & i
                        str��ϵ�˵绰 = Nvl(rs��ϵ��!��Ϣֵ)
                    Case "��ϵ�����֤��" & i
                        str��ϵ�����֤�� = Nvl(rs��ϵ��!��Ϣֵ)
                    Case "��ϵ�˸�����Ϣ" & i
                        str������Ϣ = Nvl(rs��ϵ��!��Ϣֵ)
                End Select
                rs��ϵ��.MoveNext
            Wend
            SetLinkInfo str��ϵ������, str��ϵ�˹�ϵ, str��ϵ�˵绰, str��ϵ�����֤��, str������Ϣ
            rs��ϵ��.MoveFirst
        Next
    End If
    '������Ϣ
    strSQL = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ�� Not in ('Ѫ��','ABO','RH','ҽѧ��ʾ','����ҽѧ��ʾ') And ��Ϣ�� Not like '��ϵ��%'"
    Set rs������Ϣ = zlDatabase.OpenSQLRecord(strSQL, "��ϵ��������Ϣ", lng����ID)
    '�����:115886,����,2017/11/08,�Һ���ȡ�ò�����Ϣʱ�����򱨴�
    While rs������Ϣ.EOF = False
        If Nvl(rs������Ϣ!��Ϣ��) <> "" Then
            SetOtherInfo Nvl(rs������Ϣ!��Ϣ��), Nvl(rs������Ϣ!��Ϣֵ)
        End If
        rs������Ϣ.MoveNext
    Wend
    'ҽ�ƿ�����
    Set mdicҽ�ƿ����� = Nothing
    
    Exit Sub
ErrHandl:
     If ErrCenter() = 1 Then Resume
End Sub

Private Function bln����(Optional ByVal blnCardNo As Boolean = False) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:�жϵ�ǰ�Ƿ�Ϊ�������� (���Ƿ����������ǰ󶨿�����)
'���:
'����:56599
'����:2012-12-12 14:55:36
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln�Ƿ񷢿� As Boolean
    If gCurSendCard.bln�ϸ���� = True Then
        mlng�ſ�����ID = CheckUsedBill(5, IIf(mlng�ſ�����ID > 0, mlng�ſ�����ID, gCurSendCard.lng��������), IIf(blnCardNo, mstrCard, UCase(txtPatient.Text)), gCurSendCard.lng�����ID)
        bln�Ƿ񷢿� = IIf(mlng�ſ�����ID <= 0, False, True)
        If gCurSendCard.bln���ƿ� = False Then
            bln�Ƿ񷢿� = (gCurSendCard.bln�Ƿ񷢿� = True)
        End If
    Else
        bln�Ƿ񷢿� = mbln����
        If gCurSendCard.bln���ƿ� = False Then
            bln�Ƿ񷢿� = (gCurSendCard.bln�Ƿ񷢿� = True)
        End If
    End If
    bln���� = bln�Ƿ񷢿�
    mbln���� = bln�Ƿ񷢿�
End Function

Public Sub Clear��������()
    '---------------------------------------------------------------------------------------------------------------------------------------------
'����:���������Ϣ
'���:
'����:56599
'����:2012-12-25 14:55:36
'---------------------------------------------------------------------------------------------------------------------------------------------
    '68214:������,2013-12-02,�ٴιҺ�ʱ,Ѫ��ֵ��ʼ��
    cboBloodType.ListIndex = mintDefaultBlood
    'RH
    If cboBH.ListCount > 0 Then cboBH.ListIndex = -1
    'ҽѧ��ʾ
    txtMedicalWarning.Text = ""
    '����ҽѧ��ʾ
    txtOtherWaring.Text = ""
    '��ϵ����Ϣ
    With vsLinkMan
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
        .TextMatrix(1, 4) = ""
    End With
    '�������
    With vsInoculate
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '������Ϣ
    With vsOtherInfo
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    
    '����֤��
    With vsCertificate
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
End Sub

Private Sub VsInoculate_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '�����:56599
    If Col = 1 Or Col = 3 Then '���������б༭ʱ���ж��Ƿ�����������100
        With vsInoculate
           If Len(.TextMatrix(Row, Col)) > 100 Then
                MsgBox "�������������ַ���������ַ���100,������ַ������Զ��س���", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = Mid(.TextMatrix(Row, Col), 1, 100)
           End If
        End With
        If Col = 3 And vsInoculate.Rows - 1 = Row And vsInoculate.TextMatrix(Row, Col) <> "" Then
                vsInoculate.Rows = vsInoculate.Rows + 1
        End If
    Else
        With vsInoculate
           If IsDate(.TextMatrix(Row, Col)) = False And .TextMatrix(Row, Col) <> "    -  -  " Then
                MsgBox "��������ڸ�ʽ���Ի�����ȷ�����ڣ�", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = ""
           ElseIf .TextMatrix(Row, Col) = "    -  -  " Then
                .TextMatrix(Row, Col) = ""
           End If
        End With
    End If
End Sub

Private Sub vsInoculate_KeyDown(KeyCode As Integer, Shift As Integer)
    '�����:56599
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
    '78408:���ϴ�,2014/10/9,�����ת
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
'����:ȡ���󶨿�
'���:intType:0-��ǰ����;1-��ǰ���;2-��ǰ��������
'����:ȡ���ɹ�,����true,���򷵻�False
'����:���˺�
'����:2011-07-29 11:18:05
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Curdate As Date
    Dim strSQL As String, strPassWord As String

    On Error GoTo errHandle

    Curdate = zlDatabase.Currentdate
    
    'Zl_ҽ�ƿ��䶯_Insert
    strSQL = "Zl_ҽ�ƿ��䶯_Insert("
    '      �䶯����_In   Number,
    '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
    strSQL = strSQL & "" & 14 & ","
    '      ����id_In     סԺ���ü�¼.����id%Type,
    strSQL = strSQL & "" & lngPatientID & ","
    '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
    strSQL = strSQL & "" & lngCardTypeID & ","
    '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
    strSQL = strSQL & "NULL,"
    '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
    strSQL = strSQL & "'" & strCardNo & "'" & ","
    '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
    strSQL = strSQL & "'�ҺŰ󶨿��Զ�ȡ����',"
    '      ����_In       ������Ϣ.����֤��%Type,
    strSQL = strSQL & "NULL,"
    '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
    strSQL = strSQL & "NULL,"
    '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '      Ic����_In     ������Ϣ.Ic����%Type := Null,
    strSQL = strSQL & "NULL,"
    '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
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
        If NewCol = .ColIndex("������Ϣ") Then
            If .TextMatrix(NewRow, .ColIndex("��ϵ")) = "����" Then
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
            Case .ColIndex("����")
                txt��ϵ������.Text = Trim(.EditText)
            Case .ColIndex("��ϵ")
                For i = 0 To cbo��ϵ�˹�ϵ.ListCount - 1
                    If NeedName(cbo��ϵ�˹�ϵ.List(i)) = Trim(.EditText) Then Exit For
                Next
                If i < cbo��ϵ�˹�ϵ.ListCount Then
                    cbo��ϵ�˹�ϵ.ListIndex = i
                Else
                    cbo��ϵ�˹�ϵ.ListIndex = -1
                End If
                txt������ϵ.Visible = IIf(cbo��ϵ�˹�ϵ.ListIndex = 8, True, False)
                If cbo��ϵ�˹�ϵ.ListIndex = 8 Then
                    txt������ϵ.Visible = True
                    cbo��ϵ�˹�ϵ.Width = 1225
                Else
                    txt������ϵ.Visible = False: txt������ϵ.Text = ""
                    .TextMatrix(Row, .ColIndex("������Ϣ")) = ""
                    cbo��ϵ�˹�ϵ.Width = 2425
                End If
            Case .ColIndex("���֤��")
                txt��ϵ�����֤.Text = Trim(.EditText)
            Case .ColIndex("�绰")
                txt��ϵ�˵绰.Text = Trim(.EditText)
            Case .ColIndex("������Ϣ")
                txt������ϵ.Text = Trim(.EditText)
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
    '����:���ظ�����Ϣҳ��
    '����:
    '�����:73935
    '����:Ƚ����
    '����:2014-07-3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup, Item As TaskPanelGroupItem
    Dim lngHwnd As Long
    
    Err = 0: On Error GoTo Errhand
    If CreatePlugInOK(mlngModul) Then
        If mlngPlugInHwnd <> 0 Then
            With wndTaskPanelOther
                Call .SetGroupInnerMargins(0, 0, 0, 0)
                Call .SetGroupOuterMargins(-1, -24, -1, -1)
                
                Set tkpGroup = .Groups.Add(1, "������Ϣ")
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

Private Sub DeletePatPicture(lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��������Ƭ
    '���:lng����ID - ����ID
    '����:56599
    '����:2012-12-14 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo Errhand:
    strSQL = strSQL & "Zl_������Ƭ_Delete("
    strSQL = strSQL & lng����ID & ")"
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SavePatPicture(lng����ID As Long, strFile As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���没����Ƭ
    '���:lng����ID - ����ID
    '����:56599
    '����:2012-12-13 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
        
    If strFile = "" Then Exit Sub
    If Sys.SaveLob(glngSys, 27, lng����ID, strFile, 0) = False Then
        ShowMsgbox "������Ƭ����,��ȷ���ļ��Ƿ�ɾ��!"
        Exit Sub
    End If
End Sub

Private Function ReadPatPricture(lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ƭ
    '���:lng����ID - ����ID
    '����:56599
    '����:2012-12-13 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    
    '67776:������,2013-11-20,��ȡ����Ƭ�Ĳ�����Ϣ����Ƭû�����
    Set imgPatient.Picture = Nothing
    
    strTmp = Sys.ReadLob(glngSys, 27, lng����ID)
    mstr�ɼ�ͼƬ = strTmp
    imgPatient.Picture = LoadPicture(strTmp)
    If strTmp <> "" Then Kill strTmp
End Function

Private Sub LoadIDImage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������֤ͼ��
    '����:���˺�
    '����:2014-06-30 16:20:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim objStdPic As StdPicture
    
    If mobjIDCard Is Nothing Then Exit Sub
    Call mobjIDCard.GetPhotoAsStdPicture(objStdPic)
    imgPatient.Picture = objStdPic
    mlngͼ����� = 4
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function SavePatiPic(ByVal lng����ID As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ����没����Ƭ
    '������
    '   lng����ID - ����ID
    '���ƣ�Ƚ����
    'ʱ�䣺2014-7-8
    '--------------------------------------------------------------------------------------------------------------------------------------------
    Select Case mlngͼ�����
        Case 1 '�ļ�
            SavePatPicture lng����ID, cmdialog.FileName
        Case 2 '�ɼ�
            SavePatPicture lng����ID, mstr�ɼ�ͼƬ
            mstr�ɼ�ͼƬ = ""
        Case 4 '�������֤
            If imgPatient.Picture <> 0 Then
                mstrIDImageFile = App.Path & "\SFZIMG.bmp"
                SavePicture imgPatient.Picture, mstrIDImageFile
                SavePatPicture lng����ID, mstrIDImageFile
            End If
        Case 3 '����
            DeletePatPicture lng����ID
    End Select
    
    mlngͼ����� = 0: mstr�ɼ�ͼƬ = ""
End Function

Public Sub HideFormCaption(ByVal lngHwnd As Long, Optional ByVal blnBorder As Boolean = True)
'���ܣ�����һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
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
    '����:����zlPublicPatient����
    '����:�����ɹ�,����True,���򷵻�False
    '����:Ƚ����
    '����:2014-07-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubPatient Is Nothing Then
        On Error Resume Next
        Set mobjPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If mobjPubPatient Is Nothing Then
        MsgBox "������Ϣ����������zlPublicPatient������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    Else
        If mobjPubPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser) = False Then
            MsgBox "������Ϣ����������zlPublicPatient����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreatePublicPatient = True
End Function

Private Function SetBrushCardObject(ByVal blnComm As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ���ӿ�
    '����: true-�ɹ���false-ʧ��
    '����:���ϴ�
    '����:2016/6/20 13:54:56
    '����:97634
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    
    Err = 0: On Error Resume Next
    SetBrushCardObject = True
    If txt����.Locked Then Exit Function
    If gCurSendCard.lng�����ID = 0 Or Val(gCurSendCard.str��������) < 99 Then Exit Function
    If gobjSquare.objSquareCard.zlSetBrushCardObject(gCurSendCard.lng�����ID, IIf(blnComm, txt����, Nothing), strExpend) Then
        If mobjCommEvents Is Nothing Then Set mobjCommEvents = New clsCommEvents
        Call gobjSquare.objSquareCard.zlInitEvents(Me.Hwnd, mobjCommEvents)
    End If
End Function

Private Sub zlQueryEMPIPatiInfo()
    '���ܣ���EMPIƽ̨��ȡ������Ϣ
    '���ڣ�2016/10/9 10:47:13
    '���ƣ����ϴ�
    '˵����101170
    Dim rsTmp As ADODB.Recordset, strDiff As String, strMsgInfo As String
    Dim rsPatiInfo As ADODB.Recordset
    If CreatePlugInOK(mlngModul) = False Then Exit Sub
    If Trim(txtPatient.Text) = "" Then Exit Sub
    On Error GoTo Errhand
    If zlInitMEPIPati(rsTmp) = False Then Exit Sub
    
    With rsTmp
        .AddNew
        !����ID = mlng����ID
        !����� = txt�����.Text
        !ҽ���� = txtPatiMCNO(0).Text
        !���֤�� = txt���֤��.Text
        !���� = txtPatient.Text
        !�Ա� = zlStr.NeedName(cbo�Ա�.Text)
        If IsDate(txt��������.Text) Then
            !�������� = Format(txt��������.Text & " " & IIf(IsDate(txt����ʱ��.Text), txt����ʱ��.Text, "00:00"), "YYYY-MM-DD HH:MM")
        Else
            !�������� = ""
        End If
        !�����ص� = txtBirthLocation.Text
        !���� = zlStr.NeedName(cbo����.Text)
        !���� = zlStr.NeedName(cbo����.Text)
        !ְҵ = zlStr.NeedName(cboְҵ.Text)
        !������λ = txt��λ����.Text
        !����״�� = zlStr.NeedName(cbo����.Text)
        !��ͥ�绰 = txt��ͥ�绰.Text
        !��ϵ�˵绰 = txt��ϵ�˵绰.Text
        !��λ�绰 = txt��λ�绰.Text
        !��ͥ��ַ = cbo��ͥ��ַ.Text
        !��ͥ��ַ�ʱ� = txt��ͥ�ʱ�.Text
        !���ڵ�ַ = txtRegLocation.Text
        !���ڵ�ַ�ʱ� = txt���ڵ�ַ�ʱ�.Text
        !��λ�ʱ� = txt��λ�ʱ�.Text
        !��ϵ������ = txt��ϵ������.Text
        !��ϵ�˹�ϵ = zlStr.NeedName(cbo��ϵ�˹�ϵ.Text)
        .Update
    End With
    'EMPIû���ҵ�������Ϣ,ֱ�ӷ���
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
        '104905:���ϴ���2017/1/12�����ݽӿڷ��صĲ���ID���¼��ز�����Ϣ
        If mlng����ID <> Val(Nvl(!����ID)) And Val(Nvl(!����ID)) <> 0 Then
            If mlng����ID = 0 Then
                Set rsPatiInfo = GetPatiByID("����ID", CStr(Nvl(!����ID)))
                If rsPatiInfo.EOF Then
                    mlng����ID = 0
                Else
                    zl���ز�����Ϣ rsPatiInfo
                    
                    mbln������Ϣ���� = Not (mlng����ID <> 0 And InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";������Ϣ����;") = 0)
                    txtPatient.Enabled = mbln������Ϣ����: txt��������.Enabled = mbln������Ϣ����: txt����ʱ��.Enabled = mbln������Ϣ����
                    txt����.Enabled = mbln������Ϣ����: cbo���䵥λ.Enabled = mbln������Ϣ����: cbo�Ա�.Enabled = mbln������Ϣ����
                    txt���֤��.Enabled = mbln������Ϣ����
                    SetCtrVisibleAndMove
                End If
            Else
                MsgBox "EMPI���صĲ�����Ϣ��His������Ϣ��һ�£����ڹҺŽ������²�ѯȷ�ϡ�", vbInformation, gstrSysName
                Call cmdCancel_Click
                Exit Sub
            End If
        End If
        
        mstrPlugChange = ""
        If Nvl(!ҽ����) <> "" Then
            txtPatiMCNO(0).Text = Nvl(!ҽ����)
            txtPatiMCNO(1).Text = txtPatiMCNO(0).Text
        End If
        If mbln������Ϣ���� Or mlng����ID = 0 Then
            If Nvl(!���֤��) <> "" Then txt���֤��.Text = Nvl(!���֤��)
            If Nvl(!����) <> "" Then txtPatient.Text = Nvl(!����)
            If Nvl(!�Ա�) <> "" Then cbo�Ա�.ListIndex = cbo.FindIndex(cbo�Ա�, Nvl(!�Ա�), True)
            If Nvl(!��������) <> Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS") Then
                txt��������.Text = Format(Nvl(!��������), "YYYY-MM-DD")
                txt����ʱ��.Text = Format(Nvl(!��������), "HH:MM")
            End If
        Else
            If Nvl(!����) <> "" And txtPatient.Text <> Nvl(!����) Then strDiff = ",����"
            If Nvl(!�Ա�) <> "" And cbo�Ա�.ListIndex <> cbo.FindIndex(cbo�Ա�, Nvl(!�Ա�), True) Then strDiff = strDiff & ",�Ա�"
            If Nvl(!��������) <> "" And Format(Nvl(!��������), "YYYY-MM-DD HH:MM:SS") <> Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS") Then strDiff = strDiff & ",��������"
            If Nvl(!���֤��) <> "" And txt���֤��.Text <> Nvl(!���֤��) Then strDiff = strDiff & ",���֤��"
        End If
        If txt�����.Enabled And Exist�����(Nvl(!�����), mlng����ID) = False Then
            If Nvl(!�����) <> "" Then txt�����.Text = Nvl(!�����)
        Else
            If Nvl(!�����) <> "" And txt�����.Text <> Nvl(!�����) Then strDiff = strDiff & ",�����"
        End If
        If Nvl(!�����ص�) <> "" Then txtBirthLocation.Text = Nvl(!�����ص�)
        If Nvl(!����) <> "" Then cbo����.ListIndex = cbo.FindIndex(cbo����, Nvl(!����), True)
        If Nvl(!����) <> "" Then cbo����.ListIndex = cbo.FindIndex(cbo����, Nvl(!����), True)
        If Nvl(!ְҵ) <> "" Then cboְҵ.ListIndex = cbo.FindIndex(cboְҵ, Nvl(!ְҵ))
        If Nvl(!������λ) <> "" Then txt��λ����.Text = Nvl(!������λ)
        If Nvl(!����״��) <> "" Then cbo����.ListIndex = cbo.FindIndex(cbo����, Nvl(!����״��), True)
        If Nvl(!��ͥ�绰) <> "" Then txt��ͥ�绰.Text = Nvl(!��ͥ�绰)
        If Nvl(!��ϵ�˵绰) <> "" Then txt��ϵ�˵绰.Text = Nvl(!��ϵ�˵绰)
        If Nvl(!��λ�绰) <> "" Then txt��λ�绰.Text = Nvl(!��λ�绰)
        If Nvl(!��ͥ��ַ) <> "" Then cbo��ͥ��ַ.Text = Nvl(!��ͥ��ַ): padd��ͥ��ַ.Value = Nvl(!��ͥ��ַ)
        If Nvl(!��ͥ��ַ�ʱ�) <> "" Then txt��ͥ�ʱ�.Text = Nvl(!��ͥ��ַ�ʱ�)
        If Nvl(!���ڵ�ַ) <> "" Then txtRegLocation.Text = Nvl(!���ڵ�ַ): padd���ڵ�ַ.Value = Nvl(!���ڵ�ַ)
        If Nvl(!���ڵ�ַ�ʱ�) <> "" Then txt���ڵ�ַ�ʱ�.Text = Nvl(!���ڵ�ַ�ʱ�)
        If Nvl(!��λ�ʱ�) <> "" Then txt��λ�ʱ�.Text = Nvl(!��λ�ʱ�)
        If Nvl(!��ϵ������) <> "" Then txt��ϵ������.Text = Nvl(!��ϵ������)
        If Nvl(!��ϵ�˹�ϵ) <> "" Then cbo��ϵ�˹�ϵ.ListIndex = cbo.FindIndex(cbo��ϵ�˹�ϵ, Nvl(!��ϵ�˹�ϵ), True)
    End With
    Err = 0: On Error GoTo 0
    '�������˲Ž�������
    If mlng����ID <> 0 Then
        If strDiff <> "" Then strDiff = Mid(strDiff, 2)
        If mstrPlugChange <> "" Then mstrPlugChange = Mid(mstrPlugChange, 2)
        If strDiff <> "" Then
            strMsgInfo = "���˵� " & strDiff & " ��EMPI��Ϣ��һ�£��򲻾��е���������Ϣ��Ȩ�޻�������������Ϣ��ͻ�����β�����и��¡�"
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
    '����:�ϴ�������Ϣ��EMPIƽ̨,���ƽ̨��Ϣ����ʧ�ܣ���ͬHIS����һ�����
    '����: In-lngPatiID ����ID,lngClinicID �Һ�ID
    '      Out-strErrMsg ������Ϣ����������ʧ����Ч
    '����:True-EMPIƽ̨������Ϣ�ɹ�,False-����ʧ��
    '����:���ϴ�
    '˵��:101170
    Dim blnCharge As Boolean, lngRet As Long
    If CreatePlugInOK(mlngModul) = False Then zlSaveEMPIPatiInfo = True: Exit Function
    
    On Error GoTo Errhand
    If mrsEMPIOut Is Nothing Then
        'EMPIû�в�����Ϣ����Ҫ�½�
        On Error Resume Next
        lngRet = gobjPlugIn.EMPI_AddPatiInfo(glngSys, mlngModul, lngPatiID, 0, lngClinicID, strErrMsg)
        Call zlPlugInErrH(Err, "EMPI_AddPatiInfo")
        If lngRet = 0 And Err.Number <> 438 Then Err.Clear: Exit Function
        Err.Clear: On Error GoTo Errhand
    Else
        '�ж�ƽ̨�ش�����Ϣ�Ƿ����ı�
        With mrsEMPIOut
            If txt�����.Enabled And Exist�����(Nvl(!�����), lngPatiID) = False Then
                If txt�����.Text <> Nvl(!�����) Then blnCharge = True: GoTo EMPIModify
            End If
            If txtPatiMCNO(0).Text <> Nvl(!ҽ����) Then blnCharge = True: GoTo EMPIModify
            If mbln������Ϣ���� Or blnNewPati Then
                If txt���֤��.Text <> Nvl(!���֤��) Then blnCharge = True: GoTo EMPIModify
                If txtPatient.Text <> Nvl(!����) Then blnCharge = True: GoTo EMPIModify
                If cbo�Ա�.ListIndex <> cbo.FindIndex(cbo�Ա�, Nvl(!�Ա�), True) Then blnCharge = True: GoTo EMPIModify
                If Format(txt��������.Text, "YYYY-MM-DD") <> Format(Nvl(!��������), "YYYY-MM-DD") Then blnCharge = True: GoTo EMPIModify
                If Format(txt����ʱ��.Text, "HH:MM") <> Format(Nvl(!��������), "HH:MM") Then blnCharge = True: GoTo EMPIModify
            End If
            If txtBirthLocation.Text <> Nvl(!�����ص�) Then blnCharge = True: GoTo EMPIModify
            If cbo����.ListIndex <> cbo.FindIndex(cbo����, Nvl(!����), True) Then blnCharge = True: GoTo EMPIModify
            If cbo����.ListIndex <> cbo.FindIndex(cbo����, Nvl(!����), True) Then blnCharge = True: GoTo EMPIModify
            If cboְҵ.ListIndex <> cbo.FindIndex(cboְҵ, Nvl(!ְҵ)) Then blnCharge = True: GoTo EMPIModify
            If txt��λ����.Text <> Nvl(!������λ) Then blnCharge = True: GoTo EMPIModify
            If cbo����.ListIndex <> cbo.FindIndex(cbo����, Nvl(!����״��), True) Then blnCharge = True: GoTo EMPIModify
            If txt��ͥ�绰.Text <> Nvl(!��ͥ�绰) Then blnCharge = True: GoTo EMPIModify
            If txt��ϵ�˵绰.Text <> Nvl(!��ϵ�˵绰) Then blnCharge = True: GoTo EMPIModify
            If txt��λ�绰.Text <> Nvl(!��λ�绰) Then blnCharge = True: GoTo EMPIModify
            If cbo��ͥ��ַ.Text <> Nvl(!��ͥ��ַ) Then blnCharge = True: GoTo EMPIModify
            If txt��ͥ�ʱ�.Text <> Nvl(!��ͥ��ַ�ʱ�) Then blnCharge = True: GoTo EMPIModify
            If txtRegLocation.Text <> Nvl(!���ڵ�ַ) Then blnCharge = True: GoTo EMPIModify
            If txt���ڵ�ַ�ʱ�.Text <> Nvl(!���ڵ�ַ�ʱ�) Then blnCharge = True: GoTo EMPIModify
            If txt��λ�ʱ�.Text <> Nvl(!��λ�ʱ�) Then blnCharge = True: GoTo EMPIModify
            If txt��ϵ������.Text <> Nvl(!��ϵ������) Then blnCharge = True: GoTo EMPIModify
            If cbo��ϵ�˹�ϵ.ListIndex <> cbo.FindIndex(cbo��ϵ�˹�ϵ, Nvl(!��ϵ�˹�ϵ), True) Then blnCharge = True: GoTo EMPIModify
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
    '�����:90875

    With vsCertificate
        If Col = 1 Or Col = 3 Then '֤�����벻�ܳ���30
            If Len(.TextMatrix(Row, Col)) > 30 Then
                 MsgBox "֤�������ַ���������ַ���30,������ַ������Զ��س���", vbInformation, gstrSysName
                 .TextMatrix(Row, Col) = Mid(.TextMatrix(Row, Col), 1, 30)
            End If
            If Col = 3 And .Rows - 1 = Row And .TextMatrix(Row, Col) <> "" Then
                .Rows = .Rows + 1
            End If
        ElseIf Col = 0 Or Col = 2 Then '����Ƿ�ѡ�����ظ���֤������
            For lngRow = 1 To .Rows - 1
                For lngCol = 0 To .Cols - 1 Step 2
                    If (lngRow <> Row Or lngCol <> Col) And .TextMatrix(lngRow, lngCol) = .TextMatrix(Row, Col) And .TextMatrix(Row, Col) <> "" Then
                        MsgBox .TextMatrix(lngRow, lngCol) & "�Ѵ��ڣ������ظ�ѡ��", vbInformation, gstrSysName
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
    '�����:90875
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
    '78408:���ϴ�,2014/10/9,�����ת
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
    '����:��ʼ��VSGrid�ؼ�
    '����:90875
    '����:2015/12/17 16:59:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Errhand
    Dim strSQL As String, rsTemp As ADODB.Recordset, str��ϵ As String, i As Integer
    With vsCertificate
    '��ʼ���б�����
    vsCertificate.Editable = flexEDKbdMouse
    '������ͷ
    SetColumHeader vsCertificate, C_CertificateHeader
    '��������Ϣ
    strSQL = "Select ����,ȱʡ��־ from ֤������  Where  ���� Not Like '����%' and ���� Not Like '%���֤'" & vbNewLine & _
            " And Not ���� in (Select ���� from  ҽ�ƿ���� Where Nvl(�Ƿ�֤��,0)=0 or Nvl(�Ƿ�����,0)=0)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTemp.RecordCount = 0 Then .Editable = flexEDNone: Exit Sub
        Do While Not rsTemp.EOF
            str��ϵ = str��ϵ & "|" & Nvl(rsTemp!����)
            rsTemp.MoveNext
        Loop
        str��ϵ = Mid(str��ϵ, 2)
        If str��ϵ <> "" Then .ColComboList(0) = str��ϵ: .ColComboList(2) = str��ϵ
    End With
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub LoadCertificate(ByVal lng����ID As Long)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:���ز��˵�֤����Ϣ������
    '����:���ϴ�
    'ʱ��:2015/12/17 17:37:27
    '����:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngRow As Integer, lngCol As Integer
    
    On Error GoTo Errhand
    strSQL = "Select  A.����,A.ID,B.���� from ҽ�ƿ���� A, ����ҽ�ƿ���Ϣ B " & _
            "Where A.ID= B.�����ID And A.�Ƿ�����=1 And A.�Ƿ�֤��=1 And B.״̬=0  And  B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rsTemp.RecordCount = 0 Then Exit Sub
    With vsCertificate
        .Clear 1
        .Rows = 2
        lngRow = 1: lngCol = 0
        While Not rsTemp.EOF
            .TextMatrix(lngRow, lngCol) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, lngCol + 1) = Nvl(rsTemp!����)
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

Private Sub AddCardTypeSQL(ByVal intOper As Integer, ByVal lng�����ID As Long, ByVal strCode As String, ByVal strȫ�� As String, ByVal str���� As String, _
                           ByVal lng���ų��� As Long, ByRef colPro As Collection)
    Dim strSQL As String

    ' Zl_ҽ�ƿ����_Update
    strSQL = "Zl_ҽ�ƿ����_Update("
    '  Id_In           In ҽ�ƿ����.ID%Type,
    strSQL = strSQL & "" & lng�����ID & ","
    '  ����_In         In ҽ�ƿ����.����%Type,
    strSQL = strSQL & "'" & strCode & "',"
    '  ����_In         In ҽ�ƿ����.����%Type,
    strSQL = strSQL & "'" & strȫ�� & "',"
    '  ����_In         In ҽ�ƿ����.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '  ǰ׺�ı�_In     In ҽ�ƿ����.ǰ׺�ı�%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  ���ų���_In     In ҽ�ƿ����.���ų���%Type,
    strSQL = strSQL & "" & lng���ų��� & ","
    '  ȱʡ��־_In     In ҽ�ƿ����.ȱʡ��־%Type,
    strSQL = strSQL & "" & 0 & ","
    '  �Ƿ�̶�_In     In ҽ�ƿ����.�Ƿ�̶�%Type,
    strSQL = strSQL & "1,"
    '  �Ƿ��ϸ����_In In ҽ�ƿ����.�Ƿ��ϸ����%Type,
    strSQL = strSQL & "" & 0 & ","
    '  �Ƿ�����_In     In ҽ�ƿ����.�Ƿ�����%Type,
    strSQL = strSQL & "" & 0 & ","
    '  �Ƿ�����ʻ�_In In ҽ�ƿ����.�Ƿ�����ʻ�%Type,
    strSQL = strSQL & "" & 0 & ","
    '  �Ƿ�ȫ��_In     In ҽ�ƿ����.�Ƿ�ȫ��%Type,
    strSQL = strSQL & "0,"
    '  ����_In         In ҽ�ƿ����.����%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  ��ע_In         In ҽ�ƿ����.��ע%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  �ض���Ŀ_In     In ҽ�ƿ����.�ض���Ŀ%Type,
    strSQL = strSQL & "'" & strCode & "',"
    '    �շ�ϸĿid_In   In �շ���ĿĿ¼.ID%Type,
    strSQL = strSQL & "" & "0" & ","
    '  ���㷽ʽ_In     In ҽ�ƿ����.���㷽ʽ%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  �Ƿ�����_In     In ҽ�ƿ����.�Ƿ�����%Type,
    strSQL = strSQL & "1,"
    '  ��������_In     In ҽ�ƿ����.��������%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  �Ƿ��ظ�ʹ��_In In ҽ�ƿ����.�Ƿ��ظ�ʹ��%Type,
    strSQL = strSQL & "" & 1 & ","
    '���볤��_In     In ҽ�ƿ����.���볤��%Type,
    strSQL = strSQL & "" & 10 & ","
    '���볤������_In In ҽ�ƿ����.���볤������%Type,
    strSQL = strSQL & "" & 0 & ","
    '�������_In     In ҽ�ƿ����.�������%Type,
    strSQL = strSQL & "" & 0 & ","
    strSQL = strSQL & "" & 1 & ","
    '  ������ʽ_In     In Integer := 0
    strSQL = strSQL & "" & intOper & ","
    '�Ƿ�ģ������_In     In ҽ�ƿ����.�Ƿ�ģ������%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '�����:51072
    '������������_In     In ҽ�ƿ����.������������%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '�Ƿ�ȱʡ����_In     In ҽ�ƿ����.�Ƿ�ȱʡ����%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '�����:56508
    '�Ƿ��ƿ�_In
    strSQL = strSQL & "" & 0 & ","
    '�Ƿ񷢿�_In
    strSQL = strSQL & "" & 0 & ","
    '�Ƿ�д��_In
    strSQL = strSQL & "" & 0 & ","
    '�����:57697
    '����_In
    strSQL = strSQL & "" & 0 & ","
    '�����:57326
    strSQL = strSQL & "" & 1 & ","
    '77872,���ϴ�,2014/12/3:�Ƿ�֧��ת�ʼ�����
    '�Ƿ�ת�ʼ�����_In  In ҽ�ƿ����.�Ƿ�ת�ʼ�����%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '��������_In       In ҽ�ƿ����.��������%Type := '1000',
    strSQL = strSQL & "" & "1000" & ","
    '���̿��Ʒ�ʽ_In   In ҽ�ƿ����.���̿��Ʒ�ʽ%Type := 0,
    strSQL = strSQL & "" & 0 & ","
    '90875:���ϴ�,2015/12/16,����ҽ�ƿ�֤������
    '�Ƿ�֤��_In  In ҽ�ƿ����.�Ƿ�֤��%Type:=0
    strSQL = strSQL & "" & 1 & ")"
    
    zlAddArray colPro, strSQL
End Sub

Public Sub AddCertificate(ByVal lng����ID As Long, ByRef colPro As Collection, ByVal dtCurdate As Date)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:����֤��������Ϣ������ǵ�һ�ν��������
    '����:���ϴ�
    'ʱ��:2015/12/17 17:37:27
    '����:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, rsPatiCard As ADODB.Recordset
    Dim lngRow As Integer, lngCol As Integer
    Dim lngID As Long, strCode As String
    
    On Error GoTo Errhand
    '�󶨿�ǰҪ�жϿ�����Ƿ����
    strSQL = "Select B.ID,B.����,B.���ų���,B.����,A.����,A.����ID,Decode(A.���� ,NULL,1,0) as ��ʶ from ����ҽ�ƿ���Ϣ A,ҽ�ƿ���� B " & _
            "Where A.�����ID(+)=B.ID And B.�Ƿ�֤��=1 And A.״̬(+)=0 And A.����ID(+)=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    Set rsPatiCard = zlDatabase.CopyNewRec(rsTemp)
    With vsCertificate
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) <> "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    lngID = 0: strCode = ""
                    rsTemp.Filter = "����='" & .TextMatrix(lngRow, lngCol) & "'"
                    If rsTemp.RecordCount = 0 Then
                        lngID = zlDatabase.GetNextId("ҽ�ƿ����")
                        If mstrFirstCode = "" Then
                            strCode = zlDatabase.GetMax("ҽ�ƿ����", "����", 4)
                            mstrFirstCode = strCode
                        Else
                            strCode = CStr(Val(mstrFirstCode) + 1)
                            strCode = Format(strCode, String(4, "0"))
                            mstrFirstCode = strCode
                        End If
                        Call AddCardTypeSQL(0, lngID, strCode, .TextMatrix(lngRow, lngCol), Left(.TextMatrix(lngRow, lngCol), 1), Len(.TextMatrix(lngRow, lngCol + 1)), colPro)
                    ElseIf Len(.TextMatrix(lngRow, lngCol + 1)) > Val(Nvl(rsTemp!���ų���)) Then
                        Call AddCardTypeSQL(1, Val(Nvl(rsTemp!ID)), Nvl(rsTemp!����), .TextMatrix(lngRow, lngCol), Left(.TextMatrix(lngRow, lngCol), 1), Len(.TextMatrix(lngRow, lngCol + 1)), colPro)
                    End If
                    
                    '����֤������
                    rsPatiCard.Filter = "����='" & .TextMatrix(lngRow, lngCol) & "' And ����='" & .TextMatrix(lngRow, lngCol + 1) & "'"
                    If rsPatiCard.RecordCount = 0 Then
                        'Zl_ҽ�ƿ��䶯_Insert
                         strSQL = "Zl_ҽ�ƿ��䶯_Insert("
                        '      �䶯����_In   Number,
                        '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
                        strSQL = strSQL & "" & 11 & ","
                        '      ����id_In     סԺ���ü�¼.����id%Type,
                        strSQL = strSQL & "" & lng����ID & ","
                        '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
                        strSQL = strSQL & "" & IIf(lngID = 0, rsTemp!ID, lngID) & ","
                        '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
                        strSQL = strSQL & "'" & "" & "',"
                        '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
                        strSQL = strSQL & "'" & .TextMatrix(lngRow, lngCol + 1) & "',"
                        '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
                        '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
                        strSQL = strSQL & "'" & "֤������" & "',"
                        '      ����_In       ������Ϣ.����֤��%Type,
                        strSQL = strSQL & "'" & "" & "',"
                        '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                        strSQL = strSQL & "'" & UserInfo.���� & "',"
                        '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
                        strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                        '      Ic����_In     ������Ϣ.Ic����%Type := Null,
                        strSQL = strSQL & "'" & "" & "',"
                        '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
                        strSQL = strSQL & "NULL)"
                    
                        zlAddArray colPro, strSQL
                    Else
                        rsPatiCard!��ʶ = 1
                        rsPatiCard.Update
                    End If
                End If
            Next
        Next
    End With
    '�����б���û��֤���ţ�Ҫ�����
    rsPatiCard.Filter = "��ʶ=0"
    If rsPatiCard.RecordCount > 0 Then
        rsPatiCard.MoveFirst
        Do While Not rsPatiCard.EOF
            'Zl_ҽ�ƿ��䶯_Insert
             strSQL = "Zl_ҽ�ƿ��䶯_Insert("
            '      �䶯����_In   Number,
            '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
            strSQL = strSQL & "" & 14 & ","
            '      ����id_In     סԺ���ü�¼.����id%Type,
            strSQL = strSQL & "" & lng����ID & ","
            '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
            strSQL = strSQL & "" & rsPatiCard!ID & ","
            '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
            strSQL = strSQL & "'" & "" & "',"
            '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
            strSQL = strSQL & "'" & rsPatiCard!���� & "',"
            '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
            '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
            strSQL = strSQL & "'" & "֤����ȡ����" & "',"
            '      ����_In       ������Ϣ.����֤��%Type,
            strSQL = strSQL & "'" & "" & "',"
            '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
            strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            '      Ic����_In     ������Ϣ.Ic����%Type := Null,
            strSQL = strSQL & "'" & "" & "',"
            '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
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

Public Function IsCertificateCard(ByVal lng����ID As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:֤��������
    '����:���ϴ�
    'ʱ��:2015/12/17 17:37:27
    '����:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Long, str֤������ As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strCardName As String
    
    On Error GoTo Errhand
    With vsCertificate
        '��������Ƿ�����
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) = "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    MsgBox "��ѡ�񿨺�" & .TextMatrix(lngRow, lngCol + 1) & "��֤������", vbInformation, gstrSysName
                    .Select lngRow, lngCol
                    Exit Function
                End If
                If .TextMatrix(lngRow, lngCol) <> "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    strSQL = "Select 1 from ����ҽ�ƿ���Ϣ A,ҽ�ƿ���� B " & _
                            "Where A.�����ID=B.ID And B.����=[1] And B.�Ƿ�֤��=1 And A.����=[2] And  A.����ID<>[3]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .TextMatrix(lngRow, lngCol), Trim(.TextMatrix(lngRow, lngCol + 1)), lng����ID)
                    If rsTmp.RecordCount > 0 Then
                        MsgBox .TextMatrix(lngRow, lngCol) & ":" & .TextMatrix(lngRow, lngCol + 1) & "���ڱ�ʹ��,����!", vbInformation, gstrSysName
                        .Select lngRow, lngCol
                        Exit Function
                    End If
                    str֤������ = str֤������ & ",'" & .TextMatrix(lngRow, lngCol) & "'"
                End If
            Next
        Next
        
        '���֤�������Ƿ����֤����ҽ�ƿ�����ظ����ظ��򲻱�����Ϣ
        str֤������ = Mid(str֤������, 2)
        If str֤������ = "" Then IsCertificateCard = True: Exit Function
        strSQL = "Select ���� From ҽ�ƿ���� where ���� in (" & str֤������ & ") And Nvl(�Ƿ�֤��,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                strCardName = strCardName & "," & Nvl(rsTmp!����)
            Loop
            
            strCardName = Mid(strCardName, 2)
            MsgBox "ҽ�ƿ����" & strCardName & "�������ظ�,���ܼ�����ӡ�", vbInformation, gstrSysName
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
    '����:�жϴ�����Ƿ�Ϊ�ֻ���
    '����:strinput-11λ�ֻ���
    '����:������
    '����:2017-1-25
    '---------------------------------------------------------------------------------------------
    Dim strMobileRange As String
    If Not IsNumeric(strInput) Then Exit Function
    If Len(strInput) <> 11 Then Exit Function
    '�й��ƶ�
    strMobileRange = ",139,138,137,136,135,134,159,158,157,150,151,152,147,188,187,182,183,184,178"
    '�й���ͨ
    strMobileRange = strMobileRange & ",130,131,132,156,155,186,185,145,176"
    '�й�����
    strMobileRange = strMobileRange & ",133,153,189,180,181,177,173"
    '������Ӫ��
    strMobileRange = strMobileRange & ",170,"
    If InStr(strMobileRange, "," & Mid(strInput, 1, 3) & ",") = 0 Then Exit Function
    IsMobileNO = True
End Function

Private Sub ReLoadCardFee()
    '�뿪��鿨��
    Dim lng����ID As Long, lng�շ�ϸĿID As Long
    Dim strSQL As String, str���� As String
    Dim rsTmp As ADODB.Recordset
    
    gCurSendCard.lng�շ�ϸĿID = 0
    If gCurSendCard.rs���� Is Nothing Then Exit Sub
    If gCurSendCard.rs����.RecordCount = 0 Then Exit Sub
    If gCurSendCard.lng�����ID = 0 Then Exit Sub
    If Trim(txtPatient.Text) = "" Or Trim(txt����.Text) = "" Then Exit Sub
    
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    gCurSendCard.rs����.MoveFirst
    
    strSQL = "Select Zl1_Ex_CardFee([1],[2],[3],[4],[5],[6],[7],[8],[9]) as �շ�ϸĿID From Dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����", mlngModul, gCurSendCard.lng�����ID, Trim(txt����.Text), mlng����ID, _
                Trim(txtPatient.Text), NeedName(cbo�Ա�.Text), str����, txt���֤��.Text, Val(Nvl(gCurSendCard.rs����!�շ�ϸĿID)))
    If rsTmp.EOF Then Exit Sub
    
    lng�շ�ϸĿID = Val(Nvl(rsTmp!�շ�ϸĿID))
    Set rsTmp = zlGetSpecialItemFee("", mstrPriceGrade, lng�շ�ϸĿID)
    If Not rsTmp Is Nothing Then
        Set gCurSendCard.rs���� = rsTmp
        gCurSendCard.lng�շ�ϸĿID = lng�շ�ϸĿID
    End If
End Sub

