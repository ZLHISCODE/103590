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
      Begin ZlPatiAddress.PatiAddress padd���ڵ�ַ 
         Height          =   360
         Left            =   1170
         TabIndex        =   17
         Tag             =   "���ڵ�ַ"
         Top             =   2504
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
      Begin ZlPatiAddress.PatiAddress padd��ͥ��ַ 
         Height          =   360
         Left            =   1170
         TabIndex        =   14
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
      Begin VB.TextBox txtMobile 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   8670
         MaxLength       =   20
         TabIndex        =   28
         Top             =   4110
         Width           =   2760
      End
      Begin VB.ComboBox cbo���䵥λ 
         Height          =   360
         Left            =   7965
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   459
         Width           =   720
      End
      Begin VB.TextBox txt���� 
         Height          =   360
         Left            =   1125
         MaxLength       =   50
         TabIndex        =   27
         Top             =   4110
         Width           =   5955
      End
      Begin VB.TextBox txt���� 
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
         TabIndex        =   73
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F3"
         Top             =   6540
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton cmd��ͥ��ַ 
         Caption         =   "��"
         Height          =   360
         Left            =   8070
         TabIndex        =   72
         ToolTipText     =   "�ȼ�F3"
         Top             =   2100
         Width           =   360
      End
      Begin VB.TextBox txt��ͥ�ʱ� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   10140
         MaxLength       =   6
         TabIndex        =   15
         Top             =   2095
         Width           =   1290
      End
      Begin VB.TextBox txt��ͥ�绰 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5790
         MaxLength       =   20
         TabIndex        =   8
         Top             =   855
         Width           =   2880
      End
      Begin VB.TextBox txt���֤�� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1155
         MaxLength       =   18
         TabIndex        =   7
         Top             =   868
         Width           =   2880
      End
      Begin VB.ComboBox cboְҵ 
         Height          =   360
         Left            =   8670
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2913
         Width           =   2775
      End
      Begin VB.ComboBox cbo���� 
         Height          =   360
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3322
         Width           =   2775
      End
      Begin VB.ComboBox cbo���� 
         Height          =   360
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2913
         Width           =   2775
      End
      Begin VB.ComboBox cbo���� 
         Height          =   360
         Left            =   4710
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2913
         Width           =   2775
      End
      Begin VB.ComboBox cbo�Ա� 
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
      Begin VB.TextBox txt���� 
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
      Begin VB.TextBox txt����� 
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
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4710
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3322
         Width           =   2775
      End
      Begin VB.ComboBox cbo���ʽ 
         Height          =   360
         ItemData        =   "frmPatiInfo.frx":0E46
         Left            =   8670
         List            =   "frmPatiInfo.frx":0E48
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3322
         Width           =   2775
      End
      Begin VB.ComboBox cbo��ͥ��ַ 
         Height          =   360
         Left            =   1170
         TabIndex        =   13
         Top             =   2100
         Width           =   6900
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   360
         Left            =   7080
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F3"
         Top             =   4110
         Width           =   375
      End
      Begin VB.TextBox txt֧������ 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1170
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1277
         Width           =   2880
      End
      Begin VB.TextBox txt��֤���� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5790
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1277
         Width           =   2880
      End
      Begin VB.TextBox txt������Ӧ 
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
      Begin VB.TextBox txt�໤�� 
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
         Caption         =   "��"
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
         Caption         =   "��"
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
         TabIndex        =   64
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
         TabIndex        =   63
         Top             =   1665
         Width           =   585
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
         TabIndex        =   62
         Top             =   1665
         Width           =   600
      End
      Begin VB.TextBox txt���ڵ�ַ�ʱ� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   10140
         MaxLength       =   6
         TabIndex        =   18
         Top             =   2504
         Width           =   1290
      End
      Begin VB.Frame fraContact 
         Caption         =   "��ϵ����Ϣ"
         Height          =   720
         Left            =   30
         TabIndex        =   56
         Top             =   4500
         Width           =   11415
         Begin VB.TextBox txt������ϵ 
            Height          =   360
            Left            =   6705
            MaxLength       =   30
            TabIndex        =   57
            Top             =   285
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.ComboBox cbo��ϵ�˹�ϵ 
            Height          =   360
            Left            =   6540
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   270
            Width           =   1695
         End
         Begin VB.TextBox txt��ϵ������ 
            Height          =   360
            Left            =   630
            MaxLength       =   64
            TabIndex        =   29
            Top             =   270
            Width           =   2160
         End
         Begin VB.TextBox txt��ϵ�˵绰 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   3450
            MaxLength       =   18
            TabIndex        =   30
            Top             =   270
            Width           =   2460
         End
         Begin VB.TextBox txt��ϵ�����֤ 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   9120
            MaxLength       =   18
            TabIndex        =   32
            Top             =   270
            Width           =   2205
         End
         Begin VB.Label lbl��ϵ�˹�ϵ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ"
            Height          =   240
            Left            =   6015
            TabIndex        =   61
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl��ϵ�˵绰 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�绰"
            Height          =   240
            Left            =   2925
            TabIndex        =   60
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl��ϵ������ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   240
            Left            =   135
            TabIndex        =   59
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl��ϵ�����֤ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���֤"
            Height          =   240
            Left            =   8355
            TabIndex        =   58
            Top             =   330
            Width           =   720
         End
      End
      Begin VB.Frame fraUnit 
         Caption         =   "��λ��Ϣ"
         Height          =   750
         Left            =   30
         TabIndex        =   51
         Top             =   5250
         Width           =   11415
         Begin VB.TextBox txt��λ�绰 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   9120
            MaxLength       =   20
            TabIndex        =   35
            Top             =   270
            Width           =   2205
         End
         Begin VB.TextBox txt��λ�ʱ� 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   6540
            MaxLength       =   6
            TabIndex        =   34
            Top             =   270
            Width           =   1680
         End
         Begin VB.TextBox txt��λ���� 
            Height          =   360
            Left            =   660
            MaxLength       =   100
            TabIndex        =   33
            Top             =   270
            Width           =   4860
         End
         Begin VB.CommandButton cmd��λ���� 
            Caption         =   "��"
            Height          =   360
            Left            =   5520
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   270
            Width           =   360
         End
         Begin VB.Label lbl��λ�绰 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�绰"
            Height          =   240
            Left            =   8580
            TabIndex        =   55
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl��λ�ʱ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ʱ�"
            Height          =   240
            Left            =   6015
            TabIndex        =   54
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl��λ���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
         Height          =   1215
         Left            =   30
         TabIndex        =   36
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
      Begin MSMask.MaskEdBox txt�������� 
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
         Caption         =   "�ֻ���"
         Height          =   240
         Left            =   7920
         TabIndex        =   112
         Top             =   4170
         Width           =   720
      End
      Begin VB.Label lbl����ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "ʱ��"
         Height          =   240
         Left            =   4920
         TabIndex        =   111
         Top             =   525
         Width           =   480
      End
      Begin VB.Label lbl��ͥ�ʱ� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��סַ�ʱ�"
         Height          =   240
         Left            =   8910
         TabIndex        =   98
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lbl��ͥ�绰 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�绰"
         Height          =   240
         Left            =   5280
         TabIndex        =   97
         Top             =   915
         Width           =   480
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��סַ"
         Height          =   240
         Left            =   390
         TabIndex        =   96
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label lbl���֤ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   240
         Left            =   150
         TabIndex        =   95
         Top             =   930
         Width           =   960
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   4170
         TabIndex        =   94
         Top             =   2970
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   660
         TabIndex        =   93
         Top             =   2973
         Width           =   480
      End
      Begin VB.Label lblְҵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ"
         Height          =   240
         Left            =   8160
         TabIndex        =   92
         Top             =   2970
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         Height          =   240
         Left            =   150
         TabIndex        =   91
         Top             =   3382
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   6765
         TabIndex        =   90
         Top             =   525
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   240
         Left            =   630
         TabIndex        =   89
         Top             =   519
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   630
         TabIndex        =   88
         Top             =   110
         Width           =   480
      End
      Begin VB.Label lbl����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
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
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   240
         Left            =   2430
         TabIndex        =   86
         Top             =   525
         Width           =   960
      End
      Begin VB.Label lblPatiMCNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����"
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
         Caption         =   "��֤ҽ����"
         Height          =   240
         Index           =   1
         Left            =   4560
         TabIndex        =   84
         Top             =   1740
         Width           =   1200
      End
      Begin VB.Label lbl���ʽ 
         BackStyle       =   0  'Transparent
         Caption         =   "���ʽ"
         Height          =   300
         Left            =   7680
         TabIndex        =   83
         Top             =   3352
         Width           =   960
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   240
         Left            =   4170
         TabIndex        =   82
         Top             =   3375
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   630
         TabIndex        =   81
         Top             =   4170
         Width           =   480
      End
      Begin VB.Label lbl֧������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "֧������"
         Height          =   240
         Left            =   150
         TabIndex        =   80
         Top             =   1337
         Width           =   960
      End
      Begin VB.Label lbl��֤���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��֤����"
         Height          =   240
         Left            =   4800
         TabIndex        =   79
         Top             =   1335
         Width           =   960
      End
      Begin VB.Label lbl�໤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  �໤��"
         Height          =   240
         Left            =   7680
         TabIndex        =   78
         Top             =   3780
         Width           =   960
      End
      Begin VB.Label lblBirthLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ص�"
         Height          =   240
         Left            =   150
         TabIndex        =   77
         Top             =   3780
         Width           =   960
      End
      Begin VB.Label lblRegLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڵ�ַ"
         Height          =   240
         Left            =   150
         TabIndex        =   76
         Top             =   2564
         Width           =   960
      End
      Begin VB.Label lbl���ڵ�ַ�ʱ� 
         Alignment       =   1  'Right Justify
         Caption         =   "���ڵ�ַ�ʱ�"
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
      Caption         =   "����(&H)"
      Height          =   420
      Left            =   90
      TabIndex        =   46
      Top             =   7485
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����(&X)"
      Height          =   420
      Left            =   6450
      TabIndex        =   44
      ToolTipText     =   "�ȼ���F2"
      Top             =   7455
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
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
         Caption         =   "��"
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
         TabIndex        =   43
         Top             =   4380
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
      Begin VB.Label lblInoculate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "�������"
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
         Caption         =   "������Ϣ"
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
         Caption         =   "��ϵ����Ϣ"
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
         Caption         =   "����ҽѧ��ʾ"
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
         Caption         =   "ҽѧ��ʾ"
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
         Caption         =   "Ѫ��"
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
Option Explicit 'Ҫ���������
'--------------------------------------------------------------------------------------------
'���������ر���
Private mbytFun As Byte  '0-�༭��鿴������Ϣ;2-�½�������
Private mfrmMain As Object
Private mstrPrivs As String
Private mlng����ID As Long '����Ĳ���ID
Private mlng����ID As Long '����Ŀ���ID
'--------------------------------------------------------------------------------------------
'ģ�����
Private Type Ty_Para
    bln��ͥ��ַ����    As Boolean      '��ͥ��ַ�����Ƿ�����
    bln�������Ч�Լ�� As Boolean
    bln�Զ������ As Boolean
    bln�ṹ����ַ¼�� As Boolean
    bln�����ַ�ṹ�� As Boolean
    
    bln�໤��¼�� As Boolean
    int�໤������ As Integer
End Type
Private mty_Para As Ty_Para
'--------------------------------------------------------------------------------------------
'��ض����������
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private mobjKeyboard As Object
Private mobjPlugIn As Object '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
Private mlngPlugInHwnd As Long
Private mblnPlugin As Boolean '�彨�Ƿ񴴽��ɹ�
Private mstrPlugChange As String
Private mobjPubPatient As Object
Private mblnNewPatient As Boolean
Private mblnOK As Boolean '�Ƿ�ȷ�ϳɹ�

'--------------------------------------------------------------------------------------------
'ģ�鼶����
Private mblnChange As Boolean
Private mrs��ͥ��ַ As ADODB.Recordset  '�����ͥ��ַ,��ʼʱ��ȡ������
Private mrsBaseDict As ADODB.Recordset '����,����,����״��,ְҵ
Private mrsEMPIOut As ADODB.Recordset 'EMPI���ص�����

Private mrs�ѱ� As ADODB.Recordset
Private mintNOLength As Integer '����ų���

Private mblnɨ�����֤ As Boolean '�жϲ�����Ϣ�Ƿ���ͨ��ɨ�����֤�õ�
Private mblnɨ�����֤ǩԼ As Boolean
Private mintDefaultBlood As Integer 'Ĭ��Ѫ�����
Private Enum mPageIndex
    ���� = 1
    �������� = 2
    ������Ϣ = 3
End Enum
Private Const C_InoculateHeader = "��������,4,2400,1;��������,4,2400,1;��������,4,2400,1;��������,4,2400,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_LinkManColumHeader = "����,4,1200,1;��ϵ,4,2400,1;���֤��,4,2400,1;�绰,4,1200,1;������Ϣ,4,2400,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_OtherInfoColumHeader = "��Ϣ��,4,2400,1;��Ϣֵ,4,2400,1;��Ϣ��,4,2400,1;��Ϣֵ,4,2400,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_BH = "��,��,����,δ��"
Public Event ReturnVisitClick()     '������︴ѡ��ı��Ӧ�ķѱ���ʾ
'74430,Ƚ����,2014-7-7,�Һ��еĲ�����Ϣ�༭�������ṩ�ɼ���Ƭ����
Private mstr�ɼ�ͼƬ As String '�ɼ�ͼƬ���ر���·��
Public mlngͼ����� As Long 'ָ����ǰ�Բ���ͼ�����������(1-�ļ� 2-�ɼ� 3-��� 4-���֤��ȡ)
Private mstrIDImageFile As String
Public mblnSavePati As Boolean '������Ƭ��Ϣ�򸽼���Ϣ�Ƿ��ѱ���
Private mblnNameChange As Boolean
Private mblnGetBirth As Boolean '�ж��Ƿ�����ͨ�������������

Private Sub cbo����_Change()
    mstrPlugChange = mstrPlugChange & ",����"
End Sub

Private Sub cbo����_Change()
    mstrPlugChange = mstrPlugChange & ",����״��"
End Sub

Private Sub cbo��ͥ��ַ_Change()
    mstrPlugChange = mstrPlugChange & ",��סַ"
End Sub

Private Sub cbo��ͥ��ַ_GotFocus()
    Call gobjCommFun.OpenIme(True)
End Sub
Private Sub cbo��ͥ��ַ_LostFocus()
    Call gobjCommFun.OpenIme
End Sub

Private Sub cbo��ͥ��ַ_KeyDown(KeyCode As Integer, Shift As Integer)
    '�˹��̴������������ݵ�ɾ��,�Լ���������ʱ���������б�
    '�����б���ʱ,�������ɾ����ʱ,��ɾ�������¼
    
    Dim str��ͥ��ַ As String
    
    If KeyCode = vbKeyDelete Then
        str��ͥ��ַ = cbo��ͥ��ַ.Text
        If Not mrs��ͥ��ַ Is Nothing And mty_Para.bln��ͥ��ַ���� Then
            If mrs��ͥ��ַ.State = 1 And str��ͥ��ַ <> "" Then
                If cbo��ͥ��ַ.SelText = str��ͥ��ַ And SendMessage(cbo��ͥ��ַ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = True Then
                    mrs��ͥ��ַ.Filter = "����='" & str��ͥ��ַ & "'"
                    If Not mrs��ͥ��ַ.EOF Then
                        mrs��ͥ��ַ.Delete adAffectCurrent
                        mrs��ͥ��ַ.Update
                    End If
                End If
            End If
        End If
    ElseIf KeyCode = vbKeyDown And cbo��ͥ��ַ.Text <> "" Then
        If SendMessage(cbo��ͥ��ַ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo��ͥ��ַ.hWnd, CB_SHOWDROPDOWN, True, 0&)
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
        If mrs��ͥ��ַ Is Nothing Or mty_Para.bln��ͥ��ַ���� = False Then Exit Sub
        
        str��ͥ��ַ = cbo��ͥ��ַ.Text                      '��ʱ,���ѡ���˲�������,��ѡ��������Ѿ���ɾ��
        lngλ�� = cbo��ͥ��ַ.SelStart
        
        If mrs��ͥ��ַ.State = 1 And Len(str��ͥ��ַ) > 1 Then
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(str��ͥ��ַ, 1))) > 0 Then
                mrs��ͥ��ַ.Filter = "���� like '" & UCase(str��ͥ��ַ) & "*'"
            Else
                mrs��ͥ��ַ.Filter = "���� Like '" & str��ͥ��ַ & "*'"
            End If
            
            If Not mrs��ͥ��ַ.EOF Then
                
                If mrs��ͥ��ַ.RecordCount <> cbo��ͥ��ַ.ListCount Then
                    Call SendMessage(cbo��ͥ��ַ.hWnd, CB_RESETCONTENT, 0, 0)
                    mrs��ͥ��ַ.Sort = "���� Desc,����"
                    For i = 1 To mrs��ͥ��ַ.RecordCount
                        AddComboItem cbo��ͥ��ַ.hWnd, CB_ADDSTRING, 0, mrs��ͥ��ַ!����
                        mrs��ͥ��ַ.MoveNext
                    Next
                    If SendMessage(cbo��ͥ��ַ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo��ͥ��ַ.hWnd, CB_SHOWDROPDOWN, True, 0&)
                                        
                    cbo��ͥ��ַ.Text = str��ͥ��ַ
                    cbo��ͥ��ַ.SelStart = lngλ��
                End If
            Else
                Call SendMessage(cbo��ͥ��ַ.hWnd, CB_SHOWDROPDOWN, False, 0&)
            End If
        ElseIf str��ͥ��ַ = "" Then
            cbo��ͥ��ַ.Clear
            Call SendMessage(cbo��ͥ��ַ.hWnd, CB_SHOWDROPDOWN, False, 0&)
        End If
    End If
End Sub

Private Sub cbo��ͥ��ַ_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim str���� As String
    Dim str��ͥ��ַ As String
    Dim lng�м������ As Long
    
    If (mrs��ͥ��ַ Is Nothing Or mty_Para.bln��ͥ��ַ���� = False) And KeyAscii <> 13 Then Exit Sub
    
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
            mrs��ͥ��ַ.Filter = "���� like '" & UCase(str��ͥ��ַ) & "*'"
        Else
            mrs��ͥ��ַ.Filter = "���� Like '" & str��ͥ��ַ & "*'"
        End If
        
        If Not mrs��ͥ��ַ.EOF Then
            If mrs��ͥ��ַ.RecordCount <> cbo��ͥ��ַ.ListCount Then
                Call SendMessage(cbo��ͥ��ַ.hWnd, CB_RESETCONTENT, 0, 0)
                mrs��ͥ��ַ.Sort = "���� Desc,����"
                For i = 1 To mrs��ͥ��ַ.RecordCount
                    AddComboItem cbo��ͥ��ַ.hWnd, CB_ADDSTRING, 0, mrs��ͥ��ַ!����
                    mrs��ͥ��ַ.MoveNext
                Next
                If SendMessage(cbo��ͥ��ַ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo��ͥ��ַ.hWnd, CB_SHOWDROPDOWN, True, 0&)
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
            Call SendMessage(cbo��ͥ��ַ.hWnd, CB_RESETCONTENT, 0, 0)
            If SendMessage(cbo��ͥ��ַ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then Call SendMessage(cbo��ͥ��ַ.hWnd, CB_SHOWDROPDOWN, False, 0&)
            KeyAscii = 0
            cbo��ͥ��ַ.Text = str��ͥ��ַ
            cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text)
        End If
        
        If lng�м������ > 0 Then cbo��ͥ��ַ.SelStart = lng�м������: cbo��ͥ��ַ.SelText = ""
        
    ElseIf KeyAscii = 13 Then
        'a.��û��ѡ���κ�����,����������Ϊ��,���Ϊ��ĩ��ʱ,ȷ������,��������Ϣ�����ػ���
        Call SendMessage(cbo��ͥ��ַ.hWnd, CB_SHOWDROPDOWN, False, 0&)
        
        If cbo��ͥ��ַ.Text = "" Then
            If txtPatient.Text <> "" Then
                Exit Sub
            Else
                Call gobjCommFun.PressKey(vbKeyTab): Exit Sub
            End If
        End If
        
        '�����б���ʱ���س�,��λ��ĩβ
        If cbo��ͥ��ַ.SelText = cbo��ͥ��ַ.Text Then cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text): Exit Sub
        
        If mrs��ͥ��ַ Is Nothing Then Call gobjCommFun.PressKey(vbKeyTab): Exit Sub
        If mrs��ͥ��ַ.State = 0 Then Call gobjCommFun.PressKey(vbKeyTab): Exit Sub
        If gobjCommFun.ActualLen(cbo��ͥ��ַ.Text) > 100 Then Call gobjCommFun.PressKey(vbKeyTab): Exit Sub
       
        'a.������״̬�°��س�,û��ѡ���ı�
        If cbo��ͥ��ַ.SelText = "" Then
            str��ͥ��ַ = cbo��ͥ��ַ.Text
            mrs��ͥ��ַ.Filter = "����='" & str��ͥ��ַ & "'"
            If mrs��ͥ��ַ.EOF Then
                str���� = Mid(gobjCommFun.zlGetSymbol(str��ͥ��ַ), 1, 10)
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
                
                If gobjCommFun.IsCharAlpha(str��ͥ��ַ) Then
                    If mrs��ͥ��ַ.RecordCount = 1 Then
                        cbo��ͥ��ַ.Text = mrs��ͥ��ַ!����
                    Else
                        Call SendMessage(cbo��ͥ��ַ.hWnd, CB_SHOWDROPDOWN, True, 0&)
                        Exit Sub
                    End If
                End If
            End If
            
            Call gobjCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo��ϵ�˹�ϵ_Change()
    mstrPlugChange = mstrPlugChange & ",��ϵ�˹�ϵ"
End Sub

Private Sub cbo��ϵ�˹�ϵ_Click()
    With cbo��ϵ�˹�ϵ
        If .ListIndex = 8 And txt������ϵ.Visible = False Then
            .Width = 1225: txt������ϵ.Visible = True
        ElseIf .ListIndex <> 8 And txt������ϵ.Visible Then
            .Width = 2425: txt������ϵ.Visible = False
        ElseIf .ListIndex = -1 Then
            .Width = 2425
        End If
    End With
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("��ϵ") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("��ϵ")) = gobjCommFun.GetNeedName(cbo��ϵ�˹�ϵ.Text)
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("������Ϣ")) = gobjCommFun.GetNeedName(txt������ϵ.Text)
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
    Dim strSql As String
    Dim vRect As RECT
    Dim strTemp As String
    Dim blnCancel As Boolean
    
'    vRect = gobjControl.GetControlRect(txtMedicalWarning.hWnd)
    
    strSql = "" & _
    "       Select ���� as ID,����,���� From ҽѧ��ʾ Where ���� Not Like '����%'"
    Set rsTemp = gobjDatabase.ShowSQLMultiSelect(Me, strSql, 0, "ҽѧ��ʾ", False, txtMedicalWarning.Text, "", False, False, False, vRect.Left, vRect.Top - 180, 500, blnCancel, False, True)
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
        .Flags = cdlOFNHideReadOnly
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
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd����_Click()
    If zl_SelectAndNotAddItem(Me, txt����, "", "����", "����ѡ��", True, False) = False Then
        Exit Sub
    End If
End Sub

Private Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
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

Public Function zl_AutoAddBaseItem(ByVal strTable As String, str���� As String, str���� As String, _
    Optional strTittle As String = "������Ŀ", Optional blnMsg As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�Զ�������Ŀ��Ϣ(ֻ����б���,���Ƶ���Ϣ����(ֻ���ӣ����������,����)
    '--�����:
    '--������:
    '--��  ��:���ӳɹ�,����true,���򷵻�false
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, strSql As String
    Dim int���� As Integer, strCode As String, strSpecify As String
    zl_AutoAddBaseItem = False
    If blnMsg = True Then
        If MsgBox("û���ҵ��������" & strTable & "����Ҫ��������" & strTable & "����", vbYesNo + vbQuestion, strTittle) = vbNo Then
            Exit Function
        End If
    End If
    
    Err = 0: On Error GoTo Errhand:
    
    strSql = "SELECT Nvl(MAX(LENGTH(����)), 2) As Length FROM  " & strTable
    gobjDatabase.OpenRecordset rsTemp, strSql, strTittle
    
    int���� = rsTemp!length
    
    strSql = "SELECT Nvl(MAX(LPAD(����," & int���� & ",'0')),'00') As Code FROM  " & strTable
    gobjDatabase.OpenRecordset rsTemp, strSql, strTittle
    strCode = rsTemp!Code
    
    int���� = Len(strCode)
    strCode = strCode + 1
    
    If int���� >= Len(strCode) Then
    strCode = String(int���� - Len(strCode), "0") & strCode
    End If
    strSpecify = gobjCommFun.SpellCode(str����)
    
    
    strSql = "ZL_" & strTable & "_INSERT('" & strCode & "','" & str���� & "','" & strSpecify & "')"
    gobjDatabase.ExecuteProcedure strSql, strTittle
    str���� = strCode
    zl_AutoAddBaseItem = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function zl_SelectAndNotAddItem(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    ByVal strTable As String, ByVal strTittle As String, Optional blnOnlyName As Boolean = False, _
    Optional blnδ�ҵ����� As Boolean = False, Optional strOra���� As String, Optional strWhere As String, _
    Optional blnվ�� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�๦��ѡ����
    '����:objCtl-�ı���ؼ�
    '     strKey-Ҫ������ֵ
    '     strTable-����
    '     strTittle-ѡ��������
    '     blnվ��-�Ƿ����վ������
    '����:
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, str���� As String, str���� As String
    Dim vRect As RECT, sngX As Single, sngY As Single, strSql As String
    Dim rsTemp  As ADODB.Recordset
    'gobjDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    str���� = strKey
    
    If strTable = "����" Then
        strSql = "Select rownum as ID,a.* From " & strTable & " a where 1=1 And Nvl(����,0) <3 "
    Else
        strSql = "Select rownum as ID,a.* From " & strTable & " a where 1=1 "
    End If
    
    If strKey <> "" Then
        strSql = strSql & _
        "   And ((����) like [1] or  ����  like [1] or  ����  like  upper([1]))  "
    End If
    strSql = strSql & strWhere & _
    "   order by ����"
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
        If blnδ�ҵ����� Then
            If gobjCommFun.IsCharChinese(str����) = False Then GoTo NOAdd::
            If MsgBox("ע��:" & vbCrLf & _
                   "     δ�ҵ���ص�" & strTable & ",�Ƿ����ӡ�" & str���� & "����", vbQuestion + vbYesNo + vbDefaultButton2, strTable) = vbNo Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then gobjControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If zl_AutoAddBaseItem(strTable, str����, str����, strTable & "����", False) = False Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then gobjControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
                With objCtl
                    .TextMatrix(.Row, .Col) = IIf(blnOnlyName, str����, str���� & "-" & str����)
                    If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                        .Cell(flexcpData, .Row, .Col) = str����
                    End If
                End With
            Else
                If gobjControl.IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
                objCtl.Text = IIf(blnOnlyName, str����, str���� & "-" & str����)
                objCtl.Tag = str����
                gobjCommFun.PressKey vbKeyTab
            End If
            zl_SelectAndNotAddItem = True
            Exit Function
        Else
NOAdd:
            ShowMsgBox "û���ҵ�����������" & strTable & ",����!"
            If objCtl.Enabled Then objCtl.SetFocus
            If UCase(TypeName(objCtl)) = UCase("TextBox") Then gobjControl.TxtSelAll objCtl
            Exit Function
        End If
    End If
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        With objCtl
            .TextMatrix(.Row, .Col) = IIf(blnOnlyName, Nvl(rsTemp!����), Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����))
            If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                .Cell(flexcpData, .Row, .Col) = Nvl(rsTemp!����)
            Else
                .Text = IIf(blnOnlyName, Nvl(rsTemp!����), Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����))
            End If
        End With
    Else
        If gobjControl.IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
        objCtl.Text = Nvl(rsTemp!����)
        objCtl.Tag = Nvl(rsTemp!����)
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
    Dim lng����ID           As Long
    Dim strPassWord         As String
    Dim strErrMsg           As String
    Dim strTmp              As String

    '��ȡ����ID
    If gobjSquare.objSquareCard.zlGetPatiID("���֤", strID, False, lng����ID, strPassWord, strErrMsg, , , , False) = False Then lng����ID = 0

    If mbytFun = 2 Then
        '����
        txt���֤��.Text = strID
        txtPatient.Text = strName
        Call gobjControl.CboLocate(cbo�Ա�, strSex)
        Call gobjControl.CboLocate(cbo����, strNation)
        txt��������.Text = Format(datBirthDay, "yyyy-MM-dd")
        txt����ʱ��.Text = "00:00"
        cbo��ͥ��ַ.Text = strAddress
        txtRegLocation.Text = strAddress
        padd��ͥ��ַ.Value = IIf(Trim(padd��ͥ��ַ.Value) = "", strAddress, padd��ͥ��ַ.Value)
        padd���ڵ�ַ.Value = strAddress
        '74430,Ƚ����,2014-7-7,�Һ��еĲ�����Ϣ�༭�������ṩ�ɼ���Ƭ����
        Call LoadIDImage
        Call zlQueryEMPIPatiInfo
    Else
        If MsgBox("�Ƿ�ʹ�����֤ɨ����Ϣ���µ�ǰ������Ϣ��", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbYes Then
            txt���֤��.Text = strID
            txtPatient.Text = strName
            Call gobjControl.CboLocate(cbo�Ա�, strSex)
            Call gobjControl.CboLocate(cbo����, strNation)
            txt��������.Text = Format(datBirthDay, "yyyy-MM-dd")
            txt����ʱ��.Text = "00:00"
            cbo��ͥ��ַ.Text = strAddress
            txtRegLocation.Text = strAddress
            padd��ͥ��ַ.Value = IIf(Trim(padd��ͥ��ַ.Value) = "", strAddress, padd��ͥ��ַ.Value)
            padd���ڵ�ַ.Value = strAddress
            '74430,Ƚ����,2014-7-7,�Һ��еĲ�����Ϣ�༭�������ṩ�ɼ���Ƭ����
            Call LoadIDImage
            Call zlQueryEMPIPatiInfo
        End If
    End If

End Sub

Private Function GetPatiByID(str���� As String, strValue As String) As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������ȡ��ͬ�����µĲ�����Ϣ
    '���:str���ͣ���ѯ�������� strValue ����ֵ
    '����:������Ϣ����
    '����:����
    '����:2012-08-31 04:36:33
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo ErrHandl
    strSql = "" & _
    "   Select ����ID,�����,סԺ��,���￨��,����֤��,�ѱ�,ҽ�Ƹ��ʽ,����,�Ա�,����,��������,�����ص�,���֤��,����֤��,���,ְҵ,����,����,����,����,ѧ��,����״��,��ͥ��ַ,��ͥ�绰,��ͥ��ַ�ʱ�,�໤��," & _
    "   ��ϵ������,��ϵ�˹�ϵ,��ϵ�˵�ַ,��ϵ�˵绰,���ڵ�ַ,���ڵ�ַ�ʱ�,Email,QQ,��ͬ��λID,������λ,��λ�绰,��λ�ʱ�,��λ������,��λ�ʺ�,������,��������,����ʱ��,����״̬,��������,סԺ����,��ǰ����ID,��ǰ����," & _
    "   ��Ժʱ��,��Ժʱ��,��Ժ,IC����,������,ҽ����,����,��ѯ����,�Ǽ�ʱ��,ͣ��ʱ��,����,��ϵ�����֤��,����ģʽ " & _
    "   From ������Ϣ " & _
    "   Where " & str���� & "=[1]"
    
    Set GetPatiByID = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, strValue)
    Exit Function
ErrHandl:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 And cbo�ѱ�.ListIndex <> -1 Then Call gobjCommFun.PressKey(vbKeyTab)
    
    If SendMessage(cbo�ѱ�.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo�ѱ�.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo�ѱ�.ListIndex = lngIdx
    If cbo�ѱ�.ListIndex = -1 And cbo�ѱ�.ListCount > 0 Then cbo�ѱ�.ListIndex = 0
End Sub

Private Sub cbo���ʽ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo���ʽ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo���ʽ.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo���ʽ.ListIndex = lngIdx
    If cbo���ʽ.ListIndex = -1 And cbo���ʽ.ListCount > 0 Then cbo���ʽ.ListIndex = 0
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    If cbo����.ListIndex = -1 And cbo����.ListCount > 0 Then cbo����.ListIndex = 0
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    If cbo����.ListIndex = -1 And cbo����.ListCount > 0 Then cbo����.ListIndex = 0
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    If cbo����.ListIndex = -1 And cbo����.ListCount > 0 Then cbo����.ListIndex = 0
End Sub

Private Sub cbo���䵥λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call gobjCommFun.PressKey(vbKeyTab)
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
                txt��������.Text = Format(strBirth, "YYYY-MM-DD")
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
    If SendMessage(cbo�Ա�.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo�Ա�.hWnd, KeyAscii)
    If lngIdx <> -2 Then cbo�Ա�.ListIndex = lngIdx
    If cbo�Ա�.ListIndex = -1 And cbo�Ա�.ListCount > 0 Then cbo�Ա�.ListIndex = 0
    
End Sub

Private Sub cboְҵ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboְҵ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cboְҵ.hWnd, KeyAscii)
    If lngIdx <> -2 Then cboְҵ.ListIndex = lngIdx
    If cboְҵ.ListIndex = -1 And cboְҵ.ListCount > 0 Then cboְҵ.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    If txtPatient.Text <> "" And mbytFun <> 0 Then
        If MsgBox("�Ƿ���ֹ�²���¼��?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Sub
    End If
    mlng����ID = 0
    mstrPlugChange = ""
    Unload Me
    Exit Sub
ErrOther:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Function CheckTextLength(strName As String, txtObj As TextBox) As Boolean
'����:��鲢��ʾ�ı������볤���Ƿ���
    CheckTextLength = gobjControl.TxtCheckInput(txtObj, strName, , True)
End Function

Public Function CheckExistsMCNO(ByVal strMCNO As String) As Boolean
'����:���ҽ�����Ƿ��Ѵ���
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    strSql = "Select 1 From ������Ϣ Where ҽ���� = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, strMCNO)
    If rsTmp.RecordCount > 0 Then
        MsgBox "����,�����ҽ�����Ѵ���!", vbInformation, gstrSysName
        CheckExistsMCNO = True
    End If
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function


Private Function CheckValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ���ȷ
    '����:���˺�
    '����:2011-01-07 18:13:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long, strSimilar As String, i As Integer, strMCAccount As String
    Dim strSql As String, rsTmp As ADODB.Recordset, intQuery As Integer
    Dim blnPlugInCheck As Boolean, str����ʱ�� As String
    Dim strBirthDay As String, strAge As String, strSex As String, strErrInfo As String, strInfo As String
   
    
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

    If txtPatient.Text = "" Then
        MsgBox "�����벡�˵�������", vbInformation, gstrSysName
        If txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    
    If txtPatiMCNO(0).Text <> "" Or txtPatiMCNO(1).Text <> "" Then
        If txtPatiMCNO(0).Text <> txtPatiMCNO(1).Text Then
            MsgBox "����,���������ҽ���Ų�һ�£�", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
        If gobjCommFun.ActualLen(txtPatiMCNO(0).Text) > txtPatiMCNO(0).MaxLength Then
            MsgBox "����,ҽ������󳤶Ȳ��ܳ���" & txtPatiMCNO(0).MaxLength & "���ַ���", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
    End If
    
    If CheckTextLength("����", txtPatient) = False Then Exit Function
    If CheckTextLength("�����ص�", txtBirthLocation) = False Then Exit Function
    If mty_Para.bln�ṹ����ַ¼�� Then
        If Not CheckStructAddr(padd��ͥ��ַ, padd��ͥ��ַ.MaxLength) Then Exit Function
        If Not CheckStructAddr(padd���ڵ�ַ, padd���ڵ�ַ.MaxLength) Then Exit Function
    Else
        If gobjCommFun.ActualLen(cbo��ͥ��ַ.Text) > glngMax��ͥ��ַ Then
            MsgBox "��ͥסַ���������ֻ��������" & glngMax��ͥ��ַ & "���ַ���" & glngMax��ͥ��ַ \ 2 & "�����֣�����!", vbInformation, gstrSysName
            cbo��ͥ��ַ.SetFocus: Exit Function
        End If
        If CheckTextLength("���ڵ�ַ", txtRegLocation) = False Then Exit Function
    End If
    If CheckTextLength("���ڵ�ַ�ʱ�", txt���ڵ�ַ�ʱ�) = False Then Exit Function
    If CheckTextLength("����", txt����) = False Then Exit Function
    If CheckTextLength("�����ص�", txtBirthLocation) = False Then Exit Function
    '83062
    For i = 1 To msh����.Rows - 1
        If gobjCommFun.ActualLen(msh����.TextMatrix(i, 1)) > 100 Then
            MsgBox "���˹���ҩ�ﷴӦ���������ֻ��������100���ַ���50�����֣����飡", vbInformation, gstrSysName
            If msh����.Enabled And msh����.Visible Then msh����.SetFocus
            Exit Function
        End If
        If gobjCommFun.ActualLen(msh����.TextMatrix(i, 0)) > 60 Then
            MsgBox "���˹���ҩ���������������ֻ��������60���ַ���30�����֣����飡", vbInformation, gstrSysName
            If msh����.Enabled And msh����.Visible Then msh����.SetFocus
            Exit Function
        End If
    Next i
    '69026,Ƚ����,2014-8-11,������Ч�Լ��
    '76703,Ƚ����,2014-8-15
    If txt����.Enabled And txt����.Visible Then
        If mobjPubPatient Is Nothing Then Exit Function
        If mobjPubPatient.CheckPatiAge(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""), _
                IIf(txt��������.Text = "____-__-__", "", txt��������.Text) & _
                IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text)) = False Then
            txt����.SetFocus:  Exit Function
        End If
    End If
    
    If IsDate(gobjCommFun.GetIDCardDate(txt���֤��.Text)) Then
        If Format(gobjCommFun.GetIDCardDate(txt���֤��.Text), "yyyy-mm-dd") <> Format(txt��������.Text, "yyyy-mm-dd") Then
            intQuery = MsgBox("��������֤��������ĳ������ڲ�һ�£�ʹ�����֤�Ż�ȡ�ĳ���������", vbQuestion + vbYesNoCancel, gstrSysName)
            If intQuery = 6 Then
                txt��������.Text = gobjCommFun.GetIDCardDate(txt���֤��.Text)
            ElseIf intQuery = 2 Then
                CheckValied = False
                Exit Function
            End If
        End If
    End If
    
    If IsDate(txt��������.Text) Then
        '76669�����ϴ�,2014-8-15,������������ڼ��
        str����ʱ�� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        If CDate(str����ʱ��) > gobjDatabase.Currentdate Then
            If MsgBox("����ʱ�䣺" & str����ʱ�� & " �����˵�ǰϵͳʱ�䡣" & _
                vbCrLf & vbCrLf & "���������������ڵ���ȷ�� ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus
                Exit Function
            End If
        End If
        If mty_Para.bln�໤��¼�� And Trim(txt�໤��.Text) = "" Then
            '61945 �໤��¼�� ���
            strSql = "Select Floor(Months_Between(Sysdate, [1]) / 12) as ���� From Dual"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(txt��������.Text))
            If Not rsTmp Is Nothing Then
                If Val(Nvl(rsTmp!����)) <= mty_Para.int�໤������ And mty_Para.int�໤������ <> 0 Then
                    MsgBox "������[" & mty_Para.int�໤������ & "��]�±���¼��໤��,����!"
                    Set rsTmp = Nothing
                    If txt�໤��.Enabled And txt�໤��.Visible Then txt�໤��.SetFocus
                    Exit Function
                End If
            End If
        End If
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
        If mobjPubPatient.CheckPatiIdcard(Trim(txt���֤��.Text), strBirthDay, strAge, strSex, strErrInfo) Then
            '�²��˻������ҵ�����ݵ����в�����Ϣʱ��ʾ�Ƿ������һ�µĻ�����Ϣ
            If strSex <> NeedName(cbo�Ա�.Text) Then strInfo = "�Ա�"
            If strAge <> Trim(txt����.Text) & cbo���䵥λ Then strInfo = strInfo & IIf(strInfo = "", "����", "������")
            If Format(strBirthDay, "yyyy-mm-dd") <> txt��������.Text Then strInfo = strInfo & IIf(strInfo = "", "��������", "����������")
            
            If strInfo <> "" Then
                If MsgBox("�����" & strInfo & "�����֤�ŵ�" & strInfo & "��һ�£�" & _
                        "���������֤���޸�" & strInfo & "���Ƿ������", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                    Call gobjControl.CboLocate(cbo�Ա�, strSex)
                    txt��������.Text = Format(strBirthDay, "yyyy-mm-dd")
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
    
    
    If Trim(txt֧������.Text) <> Trim(txt��֤����.Text) And (Trim(txt֧������.Text) <> "" Or Trim(txt��֤����.Text) <> "") Then
        MsgBox "������������벻һ��,����������", vbOKOnly + vbInformation, gstrSysName
        txt֧������.Text = "": txt��֤����.Text = ""
        If txt֧������.Visible = True And txt֧������.Enabled = True Then txt֧������.SetFocus
        Exit Function
    End If
    
    '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    If Not mobjPlugIn Is Nothing And mlngPlugInHwnd <> 0 Then '������������Ϣǰ��������Ч�Լ��
        On Error Resume Next
        blnPlugInCheck = mobjPlugIn.PatiInfoSaveBefore(mlng����ID)
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
    
    CheckValied = True
End Function

Private Function SimilarIDs(str���֤�� As String) As String
'���ܣ���鲡���Ƿ����������Ϣ
'���أ����Ƽ�¼�Ĳ���ID��,��"234,235,236"
    On Error GoTo errH
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Integer
    
    strSql = _
        " Select ����ID,����,Nvl(���֤��,'δ�Ǽ�') ���֤��,�����,Nvl(��ͥ��ַ,'δ�Ǽ�') ��ַ,To_Char(�Ǽ�ʱ��,'YYYY-MM-DD') �Ǽ�ʱ�� " & _
        " From ������Ϣ Where ���֤��=[1]" & _
        " Order by ����ID Desc"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", str���֤��)
    
    For i = 1 To rsTmp.RecordCount
        SimilarIDs = SimilarIDs & "|ID:" & rsTmp!����ID & ",����:" & rsTmp!���� & ",�����:" & Nvl(rsTmp!�����, "��") & ",���֤��:" & rsTmp!���֤�� & ",��ַ:" & rsTmp!��ַ & ",�Ǽ�����:" & rsTmp!�Ǽ�ʱ��
        rsTmp.MoveNext
    Next
    SimilarIDs = Mid(SimilarIDs, 2)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Function Exist�����(str����� As String, Optional lng����ID As Long) As Boolean
'���ܣ��ж�ָ��������Ƿ��Ѿ����������ݿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select ����ID From ������Ϣ Where �����=[1] And ����ID<>[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", str�����, lng����ID)
    If rsTmp.RecordCount > 0 Then Exist����� = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Function Exist�ֻ���(str�ֻ��� As String, Optional lng����ID As Long) As Boolean
'���ܣ��ж�ָ���ֻ����Ƿ��Ѿ����������ݿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select ����ID From ������Ϣ Where �ֻ���=[1] And ����ID<>[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", str�ֻ���, lng����ID)
    If rsTmp.RecordCount > 0 Then Exist�ֻ��� = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Sub cmdOK_Click()
    Dim lng����ID As Long
    If SaveData(lng����ID) = False Then Exit Sub
    Call CloseIDCard
    mblnOK = True
    Unload Me
End Sub
Private Function SaveData(ByRef lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:lng����ID-���ص�ǰ����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-10-27 14:13:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtSysDate As Date, strMCAccount As String, strCardNO As String, strPassWord As String
    Dim strNO As String, lngDept As Long, strDate As String, str�������� As String
    Dim strTmp As String, blnTrans As Boolean, strErrMsg As String, blnNewPati As Boolean
    Dim str����� As String, byt���� As Byte, i As Integer
    Dim strSql As String
    Dim cllPro As Collection
    
    Err = 0: On Error GoTo errHandle
    txtPatient.Text = Trim(txtPatient.Text)
    txt����.Text = Trim(txt����.Text)
    txt����.Tag = txt����.Text

    '��ص�������
    If CheckValied = False Then Exit Function
    If Not ((mbytFun = 0 And mlng����ID <> 0) Or mbytFun = 2) Then SaveData = True: Exit Function

    strMCAccount = Trim(txtPatiMCNO(0).Text)
    If txt����ʱ�� = "__:__" Then
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & " " & txt����ʱ��.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
   
    If Len(txt�����.Text) > mintNOLength + 1 And mintNOLength > 0 And mty_Para.bln�������Ч�Լ�� Then
        MsgBox "ע��,���������Ź���,��ȷ���Ƿ���������!", vbInformation, gstrSysName
        txt�����.SetFocus
        txt�����.SelStart = 0: txt�����.SelLength = Len(txt�����.Text)
        Exit Function
    End If
    
   
    dtSysDate = gobjDatabase.Currentdate
    strDate = "To_Date('" & Format(dtSysDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    If Exist�����(txt�����.Text, IIf(mlng����ID <> 0 And mbytFun = 0, mlng����ID, 0)) Then
        str����� = gobjDatabase.GetNextNo(3)
        If Len(str�����) > txt�����.MaxLength Then
            MsgBox "��ǰ������Ѿ�����������ʹ��,ϵͳ�Զ����������Ϊ:" & str����� & _
                   vbCrLf & "��������������������ų���:" & txt�����.MaxLength & "λ,������һ�������!", vbInformation, gstrSysName
            If txt�����.Enabled Then txt�����.SetFocus
            Exit Function
        End If
        txt�����.Text = str�����
    End If

    If mbytFun = 0 And mlng����ID <> 0 Then
        lng����ID = mlng����ID
        byt���� = 3
    Else
        lng����ID = gobjDatabase.GetNextNo(1)
        byt���� = 1: blnNewPati = True
    End If
    Set cllPro = New Collection
    mlng����ID = lng����ID
    strSql = _
    "zl_�ҺŲ��˲���_INSERT(" & byt���� & "," & lng����ID & "," & txt�����.Text & "," & _
    "'" & strCardNO & "','" & strPassWord & "'," & _
    "'" & txtPatient.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, "") & "'," & _
    "'" & NeedName(cbo�ѱ�.Text) & "','" & NeedName(cbo���ʽ.Text) & "'," & _
    "'" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "'," & _
    "'" & NeedName(cboְҵ.Text) & "','" & txt���֤��.Text & "','" & txt��λ����.Text & "'," & _
    Val(txt��λ����.Tag) & ",'" & txt��λ�绰.Text & "','" & txt��λ�ʱ�.Text & "','" & IIf(mty_Para.bln�ṹ����ַ¼��, Trim(padd��ͥ��ַ.Value), cbo��ͥ��ַ.Text) & "'," & _
    "'" & txt��ͥ�绰.Text & "','" & txt��ͥ�ʱ�.Text & "'," & strDate & ",''," & str�������� & ",'" & strMCAccount & "','" & "" & "'," & _
    "NULL," & IIf(Trim(txt����.Text) = "", "NULL,", "'" & Trim(txt����.Text) & "',") & _
     "'" & IIf(mty_Para.bln�ṹ����ַ¼��, Trim(padd���ڵ�ַ.Value), Trim(txtRegLocation.Text)) & "','" & Trim(txt���ڵ�ַ�ʱ�.Text) & "'," & IIf(Trim(txt��ϵ�����֤.Text) = "", "NULL,", "'" & Trim(txt��ϵ�����֤.Text) & "',") & _
    IIf(Trim(txt��ϵ������.Text) = "", "NULL,", "'" & Trim(txt��ϵ������.Text) & "',") & _
    IIf(Trim(txt��ϵ�˵绰.Text) = "", "NULL,", "'" & Trim(txt��ϵ�˵绰.Text) & "',") & _
    IIf(NeedName(cbo��ϵ�˹�ϵ.Text) = "", "NULL,", "'" & NeedName(cbo��ϵ�˹�ϵ.Text) & "',")    '�����:40005
        
    '�໤��_In         In ������Ϣ.�໤��%Type := Null
    strSql = strSql & IIf(Trim(txt�໤��.Text) = "", "NULL,", "'" & Trim(txt�໤��.Text) & "',")  'lgf
    '54601:������,2013-11-27,���������ص�ͻ��ڵ�ַ
    strSql = strSql & IIf(Trim(txtBirthLocation.Text) = "", "NULL,", "'" & Trim(txtBirthLocation.Text) & "',")
    strSql = strSql & "'" & txtMobile.Text & "')"
    zlAddArray cllPro, strSql
        
    If mty_Para.bln�ṹ����ַ¼�� Then
        If padd��ͥ��ַ.Value <> "" Then
           strSql = "zl_���˵�ַ��Ϣ_update(1," & lng����ID & ",NULL,3,'" & padd��ͥ��ַ.valueʡ & "','" & _
               padd��ͥ��ַ.value�� & "','" & padd��ͥ��ַ.value���� & "','" & padd��ͥ��ַ.value���� & "','" & _
               padd��ͥ��ַ.value��ϸ��ַ & "','" & padd��ͥ��ַ.Code & "')"
        Else
           strSql = "zl_���˵�ַ��Ϣ_update(2," & lng����ID & ",NULL,3)"
        End If
        zlAddArray cllPro, strSql
        If padd���ڵ�ַ.Value <> "" Then
           strSql = "zl_���˵�ַ��Ϣ_update(1," & lng����ID & ",NULL,4,'" & padd���ڵ�ַ.valueʡ & "','" & _
               padd���ڵ�ַ.value�� & "','" & padd���ڵ�ַ.value���� & "','" & padd���ڵ�ַ.value���� & "','" & _
               padd���ڵ�ַ.value��ϸ��ַ & "','" & padd���ڵ�ַ.Code & "')"
               
               
        Else
           strSql = "zl_���˵�ַ��Ϣ_update(2," & lng����ID & ",NULL,4)"
        End If
        zlAddArray cllPro, strSql
    End If
    
    'str������ϵ
    If cbo��ϵ�˹�ϵ.Text <> "" And txt������ϵ.Visible Then
        strSql = "Zl_������Ϣ�ӱ�_Update("
        '����ID_In ������Ϣ�ӱ�.����Id%Type
        strSql = strSql & "" & lng����ID & ","
        '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type0
        strSql = strSql & "'��ϵ�˸�����Ϣ',"
        '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
        strSql = strSql & "'" & txt������ϵ.Text & "',"
        '����Id_In ������Ϣ�ӱ�.����Id%Type
        strSql = strSql & "'')"
        zlAddArray cllPro, strSql
    End If
    Call Add�����������Ϣ(lng����ID, cllPro)
       
    blnTrans = True
    Call zlExecuteProcedureArrAy(cllPro, Me.Caption, True, False)
    
    '110269:���ϴ�,2016/10/13,����HIS����Ҫ�ύEMPI���ݣ�ʧ�ܺ��������ݶ�Ҫ����
    If zlSaveEMPIPatiInfo(blnNewPati, mlng����ID, 0, strErrMsg) = False Then
        gcnOracle.RollbackTrans
        If strErrMsg = "" Then strErrMsg = "��EMPIƽ̨�ϴ�������Ϣʧ�ܣ�"
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    SaveData = True
    mblnSavePati = True
    
    mstrPlugChange = ""
    '74430,Ƚ����,2014-7-7,�Һ��еĲ�����Ϣ�༭�������ṩ�ɼ���Ƭ����
    Call SavePatiPic(mlng����ID)
    '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    If Not mobjPlugIn Is Nothing And mlngPlugInHwnd <> 0 Then '������������Ϣ
        On Error Resume Next
        Call mobjPlugIn.PatiInfoSaveAfter(mlng����ID)
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

Private Sub padd���ڵ�ַ_Change()
    If mty_Para.bln�ṹ����ַ¼�� Then mstrPlugChange = mstrPlugChange & ",���ڵ�ַ"
End Sub

Private Sub padd��ͥ��ַ_Change()
    If mty_Para.bln�ṹ����ַ¼�� Then mstrPlugChange = mstrPlugChange & ",��סַ"
End Sub


Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmd��λ����_Click()
    Call SearchUnit("", txt��λ����)
End Sub

Private Sub cmd����_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    Dim i As Integer
    
    strSql = _
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

    Set rsTmp = frmPubSel.ShowSelect(Me, strSql, 2, "����ҩ��", , msh����.TextMatrix(msh����.Row, 0), "��������ҩƷ��ѡ��һ����Ϊ���˹���ҩ�")
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
    tbcPage.Left = 0
    tbcPage.Width = Me.ScaleWidth
    If (mbytFun = 0 And mlng����ID = 0) Then '�󶨾��￨ģʽ���ṩȡ����ť,�Է�Unload����,��Ϊ֮ǰ��ȡ�������ʱ���ص���Ϣ�ᱻ���
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
        If txtPatient.Visible = True And txtPatient.Enabled Then
            txtPatient.SetFocus
        ElseIf txt�����.Enabled And txt�����.Visible Then
            txt�����.SetFocus
        ElseIf txt��������.Enabled And txt��������.Visible Then
            txt��������.SetFocus
        End If
    End If
    
    mblnɨ�����֤ = False
    mblnɨ�����֤ǩԼ = True
    SetCtrVisibleAndMove
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
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If InStr(1, "txtPatient,txt����,lvwItems,txt����,cbo���䵥λ,txt��������,msh����,txt����,txtPatiMCNO,txt����,vsInoculate,cbo��ͥ��ַ", Me.ActiveControl.Name) <= 0 Then
            KeyAscii = 0
            Call gobjCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Function GetBaseDict() As ADODB.Recordset
'���ܣ����ֵ��ж�ȡ����
    Dim strSql As String, strTmp As String, arrTmp As Variant, i As Integer
    strTmp = "����,����,����״��,ְҵ,����ϵ"
    arrTmp = Split(strTmp, ",")
    For i = 0 To UBound(arrTmp)
        strTmp = arrTmp(i)
        If strSql = "" Then
            strSql = "Select '" & strTmp & "' ���,����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strTmp
        Else
            strSql = strSql & " Union all Select '" & strTmp & "' ���,����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strTmp
        End If
    Next
    strSql = strSql & " Order by ���,����"
    
    On Error GoTo errH
    Set GetBaseDict = gobjDatabase.OpenSQLRecord(strSql, "��ȡ����,����,����״��,ְҵ,����ϵ")
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function


Private Sub InitData()
'���ܣ���ʼ����Ҫ����
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
    
    cbo�ѱ�.Clear
    Call Init�ѱ�(True, True)
    
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.ListIndex = 0
    
    '�Ա�
    strSql = "Select '�Ա�' as ���,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �Ա� Union All " & _
             " Select 'ҽ�Ƹ��ʽ' as ���,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From ҽ�Ƹ��ʽ " & _
             " Order by ���,����"
    Set rsTmp = New ADODB.Recordset
    Call gobjDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
    rsTmp.Filter = "���='�Ա�'"
    cbo�Ա�.Clear
    Do While Not rsTmp.EOF
        cbo�Ա�.AddItem rsTmp!���� & "-" & rsTmp!����
'        If rsTmp!���� = gstr�Ա� Then
'            For i = 0 To cbo�Ա�.ListCount - 1
'                cbo�Ա�.ItemData(i) = 0
'            Next
'            cbo�Ա�.ItemData(cbo�Ա�.NewIndex) = 1
'            cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
'        End If
        
        If rsTmp!ȱʡ = 1 And cbo�Ա�.ListIndex = -1 Then
            cbo�Ա�.ItemData(cbo�Ա�.NewIndex) = 1
            cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    
    '�����:110155
    rsTmp.Filter = "���='ҽ�Ƹ��ʽ'"
    cbo���ʽ.Clear
    Do While Not rsTmp.EOF
        cbo���ʽ.AddItem rsTmp!���� & "-" & rsTmp!����

        If Val(Nvl(rsTmp!ȱʡ)) = 1 And cbo���ʽ.ListIndex = -1 Then
            cbo���ʽ.ItemData(cbo���ʽ.NewIndex) = 1
            cbo���ʽ.ListIndex = cbo���ʽ.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    
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
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
        .Visible = False
    End With
    '�����:56599
    Call Init����ҩ��
    
    If cbo���䵥λ.Tag = "" Then
        cbo���䵥λ.Tag = cbo���䵥λ.Text
    End If
End Sub

Private Function Init�ѱ�(bln���� As Boolean, Optional blnKeepIndex As Boolean) As Boolean
'������bln����=�Ƿ�������޳������Ŀ
'      blnKeepIndex=�Ƿ񱣳�ԭ�еķѱ�ѡ��
    Dim strSql As String, i As Integer
    Dim strKeep As String
    Dim strȱʡ�ѱ� As String
    
    On Error GoTo errH
    
    strKeep = cbo�ѱ�.Text      '������ǰ�ķѱ�,�п������ڵ�ϵͳ����û�и÷ѱ���
    If strKeep <> "" Then strKeep = Mid(strKeep, InStr(1, strKeep, "-") + 1)
    
    '72168,Ƚ����,2014/4/22,�Һ�ʱͨ���Һſ���ȷ����ѡ�ѱ�
    If mrs�ѱ� Is Nothing Then '�״ε��øú���ʱ[bln����]Ϊtrue
        Set mrs�ѱ� = New ADODB.Recordset
        '�ѱ�:���Ψһ����Ŀ(������ȱʡ�ѱ�),�����ǳ���,������Ч�ڼ估����
        strSql = "Select a.����, a.����, a.����, Nvl(a.���޳���, 0) As ����," & _
                "       Nvl(a.ȱʡ��־, 0) As ȱʡ, Nvl(b.����id, 0) As ����id" & _
                " From �ѱ� A, �ѱ����ÿ��� B" & _
                " Where a.���� = b.�ѱ�(+) And a.���� = 1" & _
                "      And Trunc(Sysdate) Between Nvl(a.��Ч��ʼ, To_Date('1900-01-01', 'YYYY-MM-DD'))" & _
                "                         And Nvl(a.��Ч����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                "      And Nvl(a.�������, 3) In (1, 3)" & _
                " Order By a.����"
        Call gobjDatabase.OpenRecordset(mrs�ѱ�, strSql, Me.Caption)
    End If
    
    If mrs�ѱ� Is Nothing Then Exit Function
    If bln���� Then
        mrs�ѱ�.Filter = "����id=0 or ����id=" & mlng����ID
    Else                        '��������޳������Ŀ
        mrs�ѱ�.Filter = "(����=0 and ����id=0) or (����=0 and ����id=" & mlng����ID & ")"
    End If
    If mrs�ѱ�.RecordCount > 0 Then mrs�ѱ�.MoveFirst
    
    cbo�ѱ�.Clear
    Do While Not mrs�ѱ�.EOF
        cbo�ѱ�.AddItem mrs�ѱ�!���� & "-" & mrs�ѱ�!����
        '��¼������Ŀ:�����Ǳ���ȱʡ��ϵͳȱʡ
        cbo�ѱ�.ItemData(cbo�ѱ�.NewIndex) = IIf(mrs�ѱ�!���� = 1, 2, 0)
        
        If strȱʡ�ѱ� = "" Then    'û�б���ȱʡʱȡϵͳȱʡ
            If mrs�ѱ�!ȱʡ = 1 Then strȱʡ�ѱ� = mrs�ѱ�!����
        End If
        mrs�ѱ�.MoveNext
    Loop
    
    If blnKeepIndex And strKeep <> "" Then Call gobjControl.CboLocate(cbo�ѱ�, strKeep)

    If cbo�ѱ�.ListIndex = -1 Then Call gobjControl.CboLocate(cbo�ѱ�, strȱʡ�ѱ�)
    
    If cbo�ѱ�.ListIndex = -1 Then If cbo�ѱ�.ListCount > 0 Then cbo�ѱ�.ListIndex = 0
    If cbo�ѱ�.ListIndex <> -1 Then cbo�ѱ�.ItemData(cbo�ѱ�.ListIndex) = 1
            
    Init�ѱ� = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function ShowMe(frmMain As Object, bytFun As Byte, ByVal lng����ID As Long, ByRef lngOut����ID As Long, _
                       Optional lng����id As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�༭������Ϣ�������
    '���:frmMain-���õ�������
    '     bytFun-0-�༭��鿴������Ϣ;
    '     bln����-����
    '     lng����ID-0-�½�����;>0��ʾ�༭�Ͳ鿴������Ϣ����
    '����:lngOut����ID-�����½������޸Ĳ��˵����Ĳ���ID
    '����:�޸ĳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-10-27 11:10:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFun = bytFun: Set mfrmMain = frmMain
    txtPatiMCNO(0).ToolTipText = "��󳤶�" & txtPatiMCNO(0).MaxLength & "λ"
    mlng����ID = lng����ID
    mlng����ID = lng����id
    '��ʼ�������㲿��
    If gobjSquare Is Nothing Then CreateSquareCardObject Me, glngModul
    Call NewCardObject   '47007
    If txt�����.Text <> "" Then mintNOLength = Len(txt�����.Text)
    mblnOK = False
    txtPatient.Enabled = True: txt��������.Enabled = True: txt����ʱ��.Enabled = True
    txt����.Enabled = True: cbo���䵥λ.Enabled = True: cbo�Ա�.Enabled = True
    txt���֤��.Enabled = True
    Call InitFact
    Me.Show 1, frmMain
    ShowMe = mblnOK
    lngOut����ID = mlng����ID
    Call CloseIDCard    '47007
End Function
Private Sub InitParaValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ֵ
    '����:���˺�
    '����:2017-10-27 13:55:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intTemp As Integer
    On Error GoTo errHandle
    With mty_Para
        .bln��ͥ��ַ���� = Val(Nvl(gobjDatabase.GetPara("��ͥ��ַ���뷽ʽ", glngSys, 1111, 1), 1)) = 1
        .bln�Զ������ = gobjDatabase.GetPara("�Զ������", glngSys, 1111) = "1"
        .bln�������Ч�Լ�� = Val(Nvl(gobjDatabase.GetPara("�������Ч�Լ��", glngSys, 1111, 1), 1)) = 1
        .bln�ṹ����ַ¼�� = Val(gobjDatabase.GetPara(251, glngSys)) <> 0 '���˵�ַ�ṹ��¼��
        .bln�����ַ�ṹ�� = Val(gobjDatabase.GetPara(252, glngSys)) <> 0 '�����ַ�ṹ��¼��
        
        intTemp = Val(gobjDatabase.GetPara("N�����±���¼��໤��", glngSys, 1111, 0))
        .bln�໤��¼�� = intTemp > 0
        .int�໤������ = intTemp
    End With
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function InitFact() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�������Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-10-29 22:48:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
        
    Call InitParaValue
    mblnChange = True
    mblnNewPatient = False
    
    Call InitData
    Call CreateObjectPlugIn  '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    Call CreateObjectKeyboard
    '����������Ϣ��������
    '69026,Ƚ����,2014-8-8,�����������
    Call CreatePublicPatient
    
    Call InitTagPage
    Call InitTaskPanelOther
    If mbytFun = 0 And mlng����ID <> 0 Then
        Me.Caption = "������ϸ��Ϣ"
        Call LoadPatientInfo
    ElseIf mbytFun = 2 Then
        Me.Caption = "����������Ϣ"
    End If
    
    If mty_Para.bln�Զ������ And txt�����.Text = "" Then txt�����.Text = gobjDatabase.GetNextNo(3)
    
    '��ʼ����ַ�ؼ�
    If mty_Para.bln�ṹ����ַ¼�� Then
        padd��ͥ��ַ.MaxLength = glngMax��ͥ��ַ: padd���ڵ�ַ.MaxLength = glngMax���ڵ�ַ
        padd��ͥ��ַ.Visible = True: padd���ڵ�ַ.Visible = True
        padd��ͥ��ַ.ShowTown = mty_Para.bln�����ַ�ṹ��: padd���ڵ�ַ.ShowTown = mty_Para.bln�����ַ�ṹ��
        cbo��ͥ��ַ.Visible = False: cmd��ͥ��ַ.Visible = False
        padd��ͥ��ַ.Top = cbo��ͥ��ַ.Top: padd��ͥ��ַ.Left = cbo��ͥ��ַ.Left
        txtRegLocation.Visible = False: cmdRegLocation.Visible = False
        padd���ڵ�ַ.Top = txtRegLocation.Top: padd���ڵ�ַ.Left = txtRegLocation.Left
    End If
    txtRegLocation.MaxLength = glngMax���ڵ�ַ
    txtBirthLocation.MaxLength = glngMax�����ص�
    InitFact = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Sub LoadPatientInfo()
    Dim rsTmp As ADODB.Recordset, strSql As String
    strSql = " Select ����,�����,��������,����,�Ա�,���֤��,ҽ����,��ͥ��ַ,��ͥ�绰,��ͥ��ַ�ʱ�,���ڵ�ַ,���ڵ�ַ�ʱ�," & _
             "        ����,����,ְҵ,����״��,�ѱ�,ҽ�Ƹ��ʽ,�໤��,��ϵ������,��ϵ�˵绰,��ϵ�˹�ϵ,��λ�ʱ�,������λ," & _
             "        ��λ�绰,�����ص�,����,�ֻ���,����ID " & _
             " From ������Ϣ Where ����ID = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, mlng����ID)
    If rsTmp.EOF Then
        MsgBox "�Ҳ���������Ϣ,��������Ĳ�����Ϣ�Ƿ���ȷ!", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    Else
        txtPatient.Text = Nvl(rsTmp!����)
        txtPatient.Locked = True
        txtPatient.Enabled = False
        txtMobile.Text = Nvl(rsTmp!�ֻ���)
        txt�����.Text = Nvl(rsTmp!�����)
        If InStr(gstrPrivs, ";�����޸������;") = 0 Then
            txt�����.Locked = True
            txt�����.Enabled = False
        End If
        txt����.Text = Nvl(rsTmp!����)
        txt����.Tag = Nvl(rsTmp!����)
        txt����.Locked = True
        txt����.Enabled = False
        cbo���䵥λ.Visible = False
        cbo���䵥λ.Enabled = False
        txt����.Width = 1395
        If IsDate(Nvl(rsTmp!��������)) Then
            txt��������.Text = Format(Nvl(rsTmp!��������), "YYYY-MM-DD")
            txt����ʱ��.Text = Format(Nvl(rsTmp!��������), "HH:MM")
        End If
        txt��������.Enabled = False
        txt����ʱ��.Enabled = False
        Call gobjControl.Cbo.Locate(cbo�Ա�, Nvl(rsTmp!�Ա�))
        cbo�Ա�.Locked = True
        cbo�Ա�.Enabled = False
        txt���֤��.Text = Nvl(rsTmp!���֤��)
        txt���֤��.Locked = True
        txt���֤��.Enabled = False
        txtPatiMCNO(0).Text = Nvl(rsTmp!ҽ����)
        txtPatiMCNO(1).Text = Nvl(rsTmp!ҽ����)
        cbo��ͥ��ַ.Text = Nvl(rsTmp!��ͥ��ַ)
        Call zlReadAddrInfo(padd��ͥ��ַ, Val(Nvl(rsTmp!����ID)), 0, 3, cbo��ͥ��ַ.Text)
        
        txt��ͥ�绰.Text = Nvl(rsTmp!��ͥ�绰)
        txt��ͥ�ʱ�.Text = Nvl(rsTmp!��ͥ��ַ�ʱ�)
        txtRegLocation.Text = Nvl(rsTmp!���ڵ�ַ)
        Call zlReadAddrInfo(padd���ڵ�ַ, Val(Nvl(rsTmp!����ID)), 0, 4, txtRegLocation.Text)
        
        txt���ڵ�ַ�ʱ�.Text = Nvl(rsTmp!���ڵ�ַ�ʱ�)
        Call gobjControl.Cbo.Locate(cbo����, Nvl(rsTmp!����))
        Call gobjControl.Cbo.Locate(cbo����, Nvl(rsTmp!����))
        Call gobjControl.Cbo.Locate(cboְҵ, Nvl(rsTmp!ְҵ))
        If cboְҵ.Text = "" Then
            cboְҵ.AddItem Nvl(rsTmp!ְҵ)
            cboְҵ.ListIndex = cboְҵ.NewIndex
        End If
        Call gobjControl.Cbo.Locate(cbo����, Nvl(rsTmp!����״��))
        If cbo����.Text = "" Then
            cbo����.AddItem Nvl(rsTmp!����״��)
            cbo����.ListIndex = cbo����.NewIndex
        End If
'        cbo����.Locked = True
        Call gobjControl.Cbo.Locate(cbo�ѱ�, Nvl(rsTmp!�ѱ�))
        If InStr(gstrPrivs, ";�����޸ķѱ�;") = 0 Then
            cbo�ѱ�.Enabled = False
        Else
            cbo�ѱ�.Enabled = True
        End If
        Call gobjControl.Cbo.Locate(cbo���ʽ, Nvl(rsTmp!ҽ�Ƹ��ʽ))
        If cbo���ʽ.Text = "" Then
            cbo���ʽ.AddItem Nvl(rsTmp!ҽ�Ƹ��ʽ)
            cbo���ʽ.ListIndex = cbo���ʽ.NewIndex
        End If
'        cbo���ʽ.Locked = True
        txt�໤��.Text = Nvl(rsTmp!�໤��)
'        txt�໤��.Locked = True
        txt��ϵ������.Text = Nvl(rsTmp!��ϵ������)
'        txt��ϵ������.Locked = True
        txt��ϵ�˵绰.Text = Nvl(rsTmp!��ϵ�˵绰)
'        txt��ϵ�˵绰.Locked = True
        txt��ϵ�����֤.Text = ""
'        txt��ϵ�����֤.Locked = True
        Call gobjControl.Cbo.Locate(cbo��ϵ�˹�ϵ, Nvl(rsTmp!��ϵ�˹�ϵ))
'        cbo��ϵ�˹�ϵ.Locked = True
        txt��λ�ʱ�.Text = Nvl(rsTmp!��λ�ʱ�)
'        txt��λ�ʱ�.Locked = True
        txt��λ����.Text = Nvl(rsTmp!������λ)
'        txt��λ����.Locked = True
        txt��λ�绰.Text = Nvl(rsTmp!��λ�绰)
'        txt��λ�绰.Locked = True
        txtBirthLocation.Text = Nvl(rsTmp!�����ص�)
'        txtBirthLocation.Locked = True
        txt����.Text = Nvl(rsTmp!����)
'        txt����.Locked = True
'        cmdOK.Visible = False
        cmd��ͥ��ַ.Visible = True
        cmdRegLocation.Visible = True
        cmd��λ����.Visible = True
        cmdBirthLocation.Visible = True
        cmd����.Visible = True
        cmdMedicalWarning.Visible = True
        Call Load�����������Ϣ(mlng����ID)
        Call ReadPatPricture(mlng����ID)
        Call zlQueryEMPIPatiInfo
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mblnChange = False
    Set mobjKeyboard = Nothing
    Call CloseIDCard
    '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    If Not mobjPlugIn Is Nothing Then Set mobjPlugIn = Nothing
    mblnPlugin = False
    mlngPlugInHwnd = 0: mblnSavePati = False
    '74430,Ƚ����,2014-7-7,�Һ��еĲ�����Ϣ�༭�������ṩ�ɼ���Ƭ����
    mlngͼ����� = 0: mstr�ɼ�ͼƬ = ""
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
            Call gobjControl.TxtSelAll(txt����)
            txt����.Visible = True
            If txt����.Visible Then txt����.SetFocus
        Case 1 '������Ӧ
            txt������Ӧ.Top = msh����.CellTop + msh����.Top + (msh����.CellHeight - txt������Ӧ.Height) / 2 - 15
            txt������Ӧ.Left = msh����.Left + msh����.CellLeft + 30
            '75446:���ϴ�,2014-7-16,������Ӧ�ı��򲻹�
            txt������Ӧ.Width = msh����.CellWidth - 60
            
            txt������Ӧ.Text = msh����.TextMatrix(msh����.Row, msh����.Col)
            txt������Ӧ.ZOrder
            Call gobjControl.TxtSelAll(txt������Ӧ)
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
    txt������Ӧ.Visible = False
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
    Call gobjControl.TxtSelAll(txtBirthLocation)
    Call gobjCommFun.OpenIme(True)
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
    Dim strSql As String, strWhere As String
    Dim strKey As String, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset, vRect As RECT
    
    On Error GoTo Errhand
    If strInput <> "" And txtInput.Tag <> "" Then Exit Sub
'    vRect = gobjControl.GetControlRect(txtInput.hWnd)
    If strInput = "" Then '�����ť
        strSql = "" & _
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
        Set rsTemp = gobjDatabase.ShowSQLSelect(Me, strSql, 2, "����", False, _
                       "", "", False, False, False, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False)
    Else
        'ȥ��"'"
        strInput = Replace(strInput, "'", " ")
        strKey = GetMatchingSting(strInput, False)
        If strInput <> "" Then
            If IsNumeric(strInput) Then '����ȫ������ʱֻƥ�����
                strWhere = " Where ���� Like Upper([1])"
            ElseIf gobjCommFun.IsCharAlpha(strInput) Then '����ȫ����ĸʱֻƥ�����
                strWhere = " Where ���� Like Upper([1])"
            Else
                strWhere = " Where ���� Like Upper([1]) Or ���� Like [1] Or ���� Like Upper([1])"
            End If
        End If
        
        strSql = "" & _
            "Select Rownum As ID, ����, ���� From ���� " & strWhere & " Order By ����"
        Set rsTemp = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "����", False, _
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
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub txtBirthLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    '73022,Ƚ����,2014-5-20,�ڵ�λ���ơ������ص㡢���ڵ�ַ����ģ�����ҹ���
    If KeyCode = vbKeyReturn And Trim(txtBirthLocation.Text) <> "" Then
        Call SearchAddress(Trim(txtBirthLocation.Text), txtBirthLocation)
    End If
End Sub

Private Sub txtBirthLocation_LostFocus()
    Call gobjCommFun.OpenIme(False)
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If mblnNameChange = True And mlng����ID = 0 Then zlQueryEMPIPatiInfo
    mblnNameChange = False
End Sub

Private Sub txtPatiMCNO_Change(Index As Integer)
    mstrPlugChange = mstrPlugChange & ",ҽ����"
End Sub

Private Sub txtRegLocation_Change()
    mstrPlugChange = mstrPlugChange & ",���ڵ�ַ"
    txtRegLocation.Tag = ""
End Sub

Private Sub txtRegLocation_GotFocus()
    Call gobjControl.TxtSelAll(txtRegLocation)
    Call gobjCommFun.OpenIme(True)
End Sub

Private Sub txtRegLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    '73022,Ƚ����,2014-5-20,�ڵ�λ���ơ������ص㡢���ڵ�ַ����ģ�����ҹ���
    If KeyCode = vbKeyReturn And Trim(txtRegLocation.Text) <> "" Then
        Call SearchAddress(Trim(txtRegLocation.Text), txtRegLocation)
    End If
End Sub

Private Sub txtMobile_Validate(Cancel As Boolean)
    If Exist�ֻ���(txtMobile.Text, IIf(mlng����ID <> 0, mlng����ID, 0)) Then
        If MsgBox("������ֻ��������������ظ����Ƿ�ȷ��¼�룿", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Cancel = True
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
    If cbo���ʽ.ListCount > 0 Then cbo���ʽ.ListIndex = 0

    If Index = 1 Then
        If txtPatiMCNO(1).Text <> txtPatiMCNO(0).Text Then
            MsgBox "����,���������ҽ���Ų�һ�£�", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If txt��������.Text = "____-__-__" Then
           gobjCommFun.PressKey (vbKeyTab) '����ʱ��
           gobjCommFun.PressKey (vbKeyTab)
       Else
           gobjCommFun.PressKey (vbKeyTab)
       End If
    End If

End Sub

Private Sub txt����ʱ��_Change()
    Dim str����ʱ�� As String
    '76669�����ϴ�,2014-8-18,�����������
    If IsDate(txt��������.Text) Then
        str����ʱ�� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        txt����.Text = ReCalcOld(CDate(str����ʱ��), cbo���䵥λ)
        If cbo���䵥λ.Visible Then
            txt����.Width = 690
        Else
            txt����.Width = 1395
        End If
        txt����.Tag = txt����.Text
    End If
End Sub

Private Sub txt����ʱ��_GotFocus()
    gobjControl.TxtSelAll txt����ʱ��
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

Private Function ReCalcBirth(ByVal strOld As String, ByVal str���䵥λ As String) As String
'����:������������䵥λ���㲡�˵ĳ�������,���䵥λΪ��ʱ,�������ռٶ�Ϊ1��1��,���䵥λΪ��ʱ,�������ڼٶ�Ϊ1��
'����:��������
    Dim strTmp As String, strFormat As String, lngDays As Long
    
    strTmp = "____-__-__"
    If str���䵥λ = "" Then
        strFormat = "YYYY-MM-DD"
        If strOld Like "*��*��" Or strOld Like "*��*����" Then
            strFormat = "YYYY-MM-01"
            lngDays = 365 * Val(strOld) + 30 * Val(Mid(strOld, InStr(1, strOld, "��") + 1))
        ElseIf strOld Like "*��*��" Or strOld Like "*����*��" Then
            lngDays = 30 * Val(strOld) + Val(Mid(strOld, InStr(1, strOld, "��") + 1))
        ElseIf strOld Like "*��" Or IsNumeric(strOld) Then
            strFormat = "YYYY-01-01"
            lngDays = 365 * Val(strOld)
        ElseIf strOld Like "*��" Or strOld Like "*����" Then
            strFormat = "YYYY-MM-01"
            lngDays = 30 * Val(strOld)
        ElseIf strOld Like "*��" Then
            lngDays = Val(strOld)
        End If
        If lngDays <> 0 Then strTmp = Format(DateAdd("d", lngDays * -1, gobjDatabase.Currentdate), strFormat)
    ElseIf strOld <> "" Then
        Select Case str���䵥λ
            Case "��"
                If Val(strOld) > 200 Then lngDays = -1
            Case "��"
                If Val(strOld) > 2400 Then lngDays = -1
            Case "��"
                If Val(strOld) > 73000 Then lngDays = -1
        End Select
        
        If lngDays = 0 Then
            strTmp = Switch(str���䵥λ = "��", "yyyy", str���䵥λ = "��", "m", str���䵥λ = "��", "d")
            strTmp = Format(DateAdd(strTmp, Val(strOld) * -1, gobjDatabase.Currentdate), "YYYY-MM-DD")
            
            If str���䵥λ = "��" Then
                strTmp = Format(strTmp, "YYYY-01-01")
            ElseIf str���䵥λ = "��" Then
                strTmp = Format(strTmp, "YYYY-MM-01")
            End If
        End If
    End If
    ReCalcBirth = strTmp
End Function

Private Function CheckOldData(ByRef txt���� As TextBox, ByRef cbo���䵥λ As ComboBox) As Boolean
'���ܣ������������ֵ����Ч��
'���أ�
    If Not IsNumeric(txt����.Text) Then CheckOldData = True: Exit Function
    
    Select Case cbo���䵥λ.Text
        Case "��"
            If Val(txt����.Text) > 200 Then
                MsgBox "���䲻�ܴ���200��!", vbInformation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "��"
            If Val(txt����.Text) > 2400 Then
                MsgBox "���䲻�ܴ���2400��!", vbInformation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "��"
            If Val(txt����.Text) > 73000 Then
                MsgBox "���䲻�ܴ���73000��!", vbInformation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                CheckOldData = False: Exit Function
            End If
    End Select
    CheckOldData = True
End Function

Private Sub txt��������_Change()
    Dim str����ʱ�� As String
    If IsDate(txt��������.Text) And mblnChange Then
        mblnChange = False
        txt��������.Text = Format(CDate(txt��������.Text), "yyyy-mm-dd") '0002-02-02�Զ�ת��Ϊ2002-02-02,����,��������2002,ʵ��ֵȴ��0002
        mblnChange = True
        
        str����ʱ�� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        txt����.Text = ReCalcOld(CDate(str����ʱ��), cbo���䵥λ)
        If cbo���䵥λ.Visible Then
            txt����.Width = 690
        Else
            txt����.Width = 1395
        End If
        txt����.Tag = txt����.Text
        cbo���䵥λ.Tag = cbo���䵥λ.Text
        mblnGetBirth = False
    End If
End Sub
Private Sub txt��������_GotFocus()
    gobjControl.TxtSelAll txt��������
End Sub

Private Sub txt��������_LostFocus()
    If txt��������.Text <> "____-__-__" And Not IsDate(txt��������.Text) Then
      If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus
    End If
End Sub


Private Sub txt��λ�绰_GotFocus()
    Call gobjControl.TxtSelAll(txt��λ�绰)
End Sub

Private Sub txt��λ�绰_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt��λ�绰, KeyAscii
End Sub

Private Sub txt��λ����_Change()
    txt��λ����.Tag = ""
End Sub

Private Sub txt��λ����_GotFocus()
    Call gobjControl.TxtSelAll(txt��λ����)
    Call gobjCommFun.OpenIme(True)
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
    Dim strSql As String, strWhere As String
    Dim strKey As String, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset, vRect As RECT
    
    On Error GoTo Errhand
    If strInput <> "" And txtInput.Tag <> "" Then Exit Sub
'    vRect = gobjControl.GetControlRect(txtInput.hWnd)
    If strInput = "" Then '�����ť
        strSql = "" & _
        "       Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From  ��Լ��λ" & _
        "       Where ����ʱ�� Is Null Or ����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD')" & _
        "       Start With �ϼ�ID is NULL" & _
        "       Connect by Prior ID=�ϼ�ID"
        '75888,Ƚ����,2014-7-28
        Set rsTemp = gobjDatabase.ShowSQLSelect(Me, strSql, 2, "��λ", False, _
                       "", "", False, True, False, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False)
    Else
        'ȥ��"'"
        strInput = Replace(strInput, "'", " ")
        strKey = GetMatchingSting(strInput, False)
        If strInput <> "" Then
            If IsNumeric(strInput) Then '����ȫ������ʱֻƥ�����
                strWhere = " Where ���� Like Upper([1])"
            ElseIf gobjCommFun.IsCharAlpha(strInput) Then '����ȫ����ĸʱֻƥ�����
                strWhere = " Where ���� Like Upper([1])"
            Else
                strWhere = " Where ���� Like Upper([1]) Or ���� Like [1] Or ���� Like Upper([1])"
            End If
        End If
        
        strSql = "" & _
        "       Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From  ��Լ��λ" & strWhere & _
        "       And (����ʱ�� Is Null Or ����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD'))"
        Set rsTemp = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "��λ", False, _
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
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub txt��λ����_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt��λ����, KeyAscii
End Sub

Private Sub txt��λ����_LostFocus()
    Call gobjCommFun.OpenIme
End Sub

Private Sub txt��λ�ʱ�_GotFocus()
    Call gobjControl.TxtSelAll(txt��λ�ʱ�)
End Sub

Private Sub txt��λ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt��λ�ʱ�, KeyAscii
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim ObjItem As ListItem
    Dim strSql As String
            
    If KeyAscii <> 13 Then
        If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    Else
        KeyAscii = 0
        '75286:���ϴ���2014-7-16������¼�����ҩ��
        msh����.TextMatrix(msh����.Row, 0) = txt����.Text '�����:56599

        strSql = " Select Distinct A.ID,A.����," & _
        " A.����,A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������," & _
        " Decode(B.�Ƿ���ҩ,1,'��','') as ��ҩ,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
        " From ������ĿĿ¼ A,ҩƷ���� B,������Ŀ���� C" & _
        " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID And A.Id=C.������Ŀid" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        " And (C.���� like [1] OR A.���� like [1] OR C.���� like [1])"
        
        On Error GoTo errH
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(txt����.Text) & "%")
        
        With rsTmp
            If .BOF Or .EOF Then
                msh����.SetFocus: msh����_EnterCell
                Exit Sub
            Else
                Me.lvwItems.ListItems.Clear
                Do While Not .EOF
                    Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����, , IIf(!Ƥ�� <> "", 1, 2))
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("��λ").Index - 1) = IIf(IsNull(!��λ), "", !��λ)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("�������").Index - 1) = IIf(IsNull(!�������), "", !�������)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("��ҩ").Index - 1) = IIf(IsNull(!��ҩ), "", !��ҩ)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("Ƥ��").Index - 1) = IIf(IsNull(!Ƥ��), "", !Ƥ��)
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
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
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
    mblnChange = True
End Sub

Private Sub txt���ڵ�ַ�ʱ�_GotFocus()
    Call gobjControl.TxtSelAll(txt���ڵ�ַ�ʱ�)
End Sub

Private Sub txt���ڵ�ַ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt��ͥ�绰_GotFocus()
    Call gobjControl.TxtSelAll(txt��ͥ�绰)
End Sub

Private Sub txt��ͥ�绰_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt��ͥ�绰, KeyAscii
End Sub

Private Sub txt��ͥ�ʱ�_GotFocus()
    Call gobjControl.TxtSelAll(txt��ͥ�ʱ�)
End Sub

Private Sub txt��ͥ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt��ͥ�ʱ�, KeyAscii
End Sub
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

Private Sub txt��ϵ������_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("����") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("����")) = txt��ϵ������.Text
        If vsLinkMan.Rows = vsLinkMan.FixedRows + 1 And txt��ϵ������.Text <> "" Then
            vsLinkMan.Rows = vsLinkMan.Rows + 1
        End If
    End If
End Sub

Private Sub txt�����_GotFocus()
    Call gobjControl.TxtSelAll(txt�����)
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
        Call gobjCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = 32 Then
        KeyAscii = 0
        If txt�����.Text = "" Then
            txt�����.Text = gobjDatabase.GetNextNo(3)
            mintNOLength = Len(Trim(txt�����.Text))
        End If
        Call gobjCommFun.PressKey(vbKeyTab)
    ElseIf InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Or InStr(gstrPrivs, ";�����޸������;") = 0 Then
        KeyAscii = 0
    End If
End Sub
 
Private Sub txt����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt����.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt����.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt����_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt����.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt����_GotFocus()
    Call gobjCommFun.OpenIme
    Call gobjControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim blnTab As Boolean
    
    If KeyAscii = vbKeyReturn Then
        If cbo���䵥λ.Visible = False And IsNumeric(txt����.Text) Then
            Call txt����_Validate(False)
            Call cbo���䵥λ.SetFocus
        Else
            Call gobjCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt����.Text) And cbo���䵥λ.Visible Then Call gobjCommFun.PressKey(vbKeyTab)
    Else
        '�������Ƽ��� ָ����������ַ� ����:49908
        If InStr("~����@#��%����&*��������-+=|����������~`!#$%^&*()-_=+|\/?<>,/<>", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim strBirth As String
    txt����.Text = Trim(txt����.Text)
    If Not IsNumeric(txt����.Text) And Trim(txt����.Text) <> "" Then
        cbo���䵥λ.ListIndex = -1: cbo���䵥λ.Visible = False: txt����.Width = 1395
    ElseIf cbo���䵥λ.Visible = False Then
        cbo���䵥λ.ListIndex = 0: cbo���䵥λ.Visible = True: txt����.Width = 690
    End If
    If txt����.Text <> txt����.Tag Then
        mblnChange = False
        If Not IsDate(txt��������.Text) Then mblnGetBirth = True
        '125451�������Ƿ�����ͨ����������������
        If mblnGetBirth Then
            If mobjPubPatient.ReCalcBirthDay(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""), strBirth) Then
                txt��������.Text = Format(strBirth, "YYYY-MM-DD")
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
    Call gobjControl.TxtSelAll(txt��λ����)
    Call gobjCommFun.OpenIme(True)
End Sub

Private Sub txt������ϵ_LostFocus()
    Call gobjCommFun.OpenIme
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

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If (txt����.Tag <> "" Or Trim(txt����.Text) = "") Then gobjCommFun.PressKey vbKeyTab: Exit Sub
    If zl_SelectAndNotAddItem(Me, txt����, Trim(txt����.Text), "����", "����ѡ��", True, False) = False Then
        Exit Sub
    End If
End Sub

Private Sub txt���֤��_Change()
    If mblnɨ�����֤ǩԼ And ActiveControl Is txt���֤�� And Not mobjIDCard Is Nothing Then
            mobjIDCard.SetEnabled txt���֤��.Text = ""
    End If
End Sub

Private Sub txt���֤��_GotFocus()
    Call gobjControl.TxtSelAll(txt���֤��)

    If mblnɨ�����֤ǩԼ = True And txt���֤��.Text = "" Then
        OpenIDCard
    End If
End Sub
Private Sub txt���֤��_KeyPress(KeyAscii As Integer)
    
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
        '�²��˲ŵ���
        If mblnNameChange = True And mlng����ID = 0 Then zlQueryEMPIPatiInfo
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

Private Sub txt���֤��_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
End Sub

Private Sub txt���֤��_Validate(Cancel As Boolean)
    '65663:������,2014-02-20,�������֤�ż����������
    If IsDate(gobjCommFun.GetIDCardDate(txt���֤��.Text)) = False Then Exit Sub
    If Format(gobjCommFun.GetIDCardDate(txt���֤��.Text), "yyyy-mm-dd") <> Format(txt��������.Text, "yyyy-mm-dd") Then
        If IsDate(txt��������.Text) Then MsgBox "��������֤��������ĳ������ڲ�һ�£���ʹ�����֤�Ż�ȡ�������滻��", vbInformation, gstrSysName
        txt��������.Text = gobjCommFun.GetIDCardDate(txt���֤��.Text)
    End If
End Sub

Private Function GetOneDept(lng�շ�ϸĿID As Long) As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select B.ִ�п���ID From �շ���ĿĿ¼ A,�շ�ִ�п��� B Where B.�շ�ϸĿID=A.ID And A.ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng�շ�ϸĿID)
    If Not rsTmp.EOF Then
        GetOneDept = rsTmp!ִ�п���ID 'Ĭ��ȡ��һ��(���ж��)
    Else
        GetOneDept = UserInfo.����ID '��û��ָ������ȡ����Ա���ڿ���
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
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
    If gobjComlib.ErrCenter() = 1 Then
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
    If gobjComlib.ErrCenter() = 1 Then Resume
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
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function
Private Function zlComboxLoadFromSQL(ByVal strSql As String, cboControl As Variant, Optional ByVal blnID As Boolean = False) As Boolean
'�������Ĺ����Ǵ����ݿ��ж����б�ֵ��װ����������
    Dim rsTemp As New ADODB.Recordset
    Dim intCount As Long
    Dim cmbArray As Variant
    
    On Error GoTo errHandle
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "��ȡCbo����")
    '����������
    If IsArray(cboControl) Then
        cmbArray = cboControl
    Else
        'ǿ�����һ������
        cmbArray = Array(cboControl)
    End If
    
    For intCount = LBound(cmbArray) To UBound(cmbArray)
        cmbArray(intCount).Clear
        Do Until rsTemp.EOF
            If IsNull(rsTemp("����")) Then
                cmbArray(intCount).AddItem rsTemp.AbsolutePosition & "." & rsTemp("����")
            Else
                cmbArray(intCount).AddItem rsTemp("����") & "." & rsTemp("����")
            End If
            If blnID = True Then cmbArray(intCount).ItemData(cmbArray(intCount).NewIndex) = rsTemp("ID")
            If rsTemp("ȱʡ��־") = 1 Then
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

Private Function GetCardDataSql(ByVal byt�䶯���� As Byte, ByVal lng����ID As Long, ByVal lng�����ID As Long, _
   ByVal strԭ���� As String, ByVal strCard As String, ByVal str���� As String, ByVal dtCurDate As Date, _
   ByVal strICCard As String, Optional ByVal str�䶯ԭ�� As String = "")
    Dim strSql As String
    Dim strPassWord As String
    strPassWord = gobjCommFun.zlStringEncode(str����)
    'Zl_ҽ�ƿ��䶯_Insert
     strSql = "Zl_ҽ�ƿ��䶯_Insert("
    '      �䶯����_In   Number,
    '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����);
    '��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
    strSql = strSql & "" & byt�䶯���� & ","
    '      ����id_In     סԺ���ü�¼.����id%Type,
    strSql = strSql & "" & lng����ID & ","
    '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
    strSql = strSql & "" & lng�����ID & ","
    '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
    strSql = strSql & "'" & strԭ���� & "',"
    '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
    strSql = strSql & "'" & strCard & "',"
    '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
    '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
    strSql = strSql & "'" & str�䶯ԭ�� & "',"
    '      ����_In       ������Ϣ.����֤��%Type,
    strSql = strSql & "'" & strPassWord & "',"
    '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
    strSql = strSql & "'" & UserInfo.���� & "',"
    '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
    strSql = strSql & "to_date('" & Format(dtCurDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '      Ic����_In     ������Ϣ.Ic����%Type := Null,
    strSql = strSql & "'" & strICCard & "',"
    '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
    strSql = strSql & IIf(str�䶯ԭ�� = "", "NULL)", "'" & str�䶯ԭ�� & "')")
    GetCardDataSql = strSql
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
    '����:�����֤������
    '����:����
    '����:2012-08-31 16:28:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '��ʼ���Կ�����
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    '�򿪶�����
    mobjIDCard.SetEnabled (True)
End Sub

Private Function zl_Getȱʡ�������() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡȱʡ�������
    '����:ȱʡ�����������
    '����:����
    '����:2012-08-31 11:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim lngCardTypeID As Long
    Dim rsTemp As Recordset
    
    On Error GoTo ErrHandl:

    strSql = "" & _
    "   Select Id, ����, ����, ����, ǰ׺�ı�, ���ų���, ȱʡ��־, �Ƿ�̶�, �Ƿ��ϸ����, " & _
    "           nvl(�Ƿ�����,0) as �Ƿ�����, nvl(�Ƿ�����ʻ�,0) as �Ƿ�����ʻ�, " & _
    "           nvl(�Ƿ�ȫ��,0) as �Ƿ�ȫ��,nvl(�Ƿ��ظ�ʹ��,0) as �Ƿ��ظ�ʹ�� , " & _
    "           nvl(���볤��,10) as ���볤��,nvl(���볤������,0) as ���볤������,nvl(�������,0) as �������," & _
    "           nvl(�Ƿ�����,0) as �Ƿ�����,����, ��ע, �ض���Ŀ, ���㷽ʽ, �Ƿ�����, ��������,Nvl(������������,0) as ������������,Nvl(�Ƿ�ȱʡ����,0) as �Ƿ�ȱʡ����," & _
    "           nvl(�Ƿ�ģ������,0) as �Ƿ�ģ������,nvl(��������,'1000') as �������� " & _
    "    From ҽ�ƿ����" & _
    "    Where ID = [1]" & _
    "    Order by ����"

    lngCardTypeID = Val(gobjDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, glngModul, , , True))
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lngCardTypeID)
    If rsTemp Is Nothing Then zl_Getȱʡ������� = "": Exit Function
    If rsTemp.RecordCount <= 0 Then zl_Getȱʡ������� = "": Exit Function
    zl_Getȱʡ������� = rsTemp!����
    Exit Function
ErrHandl:
    If gobjComlib.ErrCenter() = 1 Then Resume
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
    
    tbcPage.Top = ScaleTop + 50
    tbcPage.Height = Me.ScaleHeight - tbcPage.Top - (Me.ScaleHeight - cmdHelp.Top + 45)
       
    lblPatiMCNO(0).Enabled = False: lblPatiMCNO(1).Enabled = False
    txtPatiMCNO(0).Enabled = False: txtPatiMCNO(1).Enabled = False
     
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
    Call gobjControl.TxtSelAll(txt��֤����)
    Call OpenPassKeyboard(txt��֤����, False)
End Sub

Private Sub txt��֤����_LostFocus()
    Call ClosePassKeyboard(txt��֤����)
End Sub
Private Sub txt֧������_GotFocus()
    Call gobjControl.TxtSelAll(txt֧������)
    Call OpenPassKeyboard(txt֧������, False)
End Sub

Private Sub LoadOldData(ByVal strOld As String, ByRef txt���� As TextBox, ByRef cbo���䵥λ As ComboBox)
'����:�����ݿ��б�������䰴�淶�ĸ�ʽ���ص�����,���淶��ԭ����ʾ
    Call gobjControl.LoadOldData(strOld, txt����, cbo���䵥λ)
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
    If Nvl(rsPatiInfo!�Ա�) <> "" Then
        Call gobjControl.CboLocate(cbo�Ա�, rsPatiInfo!�Ա�)
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
        If cbo���䵥λ.Visible Then
            txt����.Width = 690
        Else
            txt����.Width = 1395
        End If
        txt����.Tag = txt����.Text
    Else
         txt����ʱ��.Text = "__:__"
         txt��������.Text = ReCalcBirth(Val(txt����.Text), cbo���䵥λ.Text)
    End If

    '���֤��
    If Nvl(rsPatiInfo!���֤��) <> "" Then
        txt���֤��.Text = rsPatiInfo!���֤��
        If InStr(1, txt��������.Text, "__") > 0 Then
            strTmp = gobjCommFun.GetIDCardDate(txt���֤��.Text)
            If IsDate(strTmp) Then txt��������.Text = strTmp
        End If
    End If
    'ְҵ
    If Nvl(rsPatiInfo!ְҵ) <> "" Then
        cboְҵ.ListIndex = gobjControl.Cbo.FindIndex(cboְҵ, rsPatiInfo!ְҵ)
        If cboְҵ.ListIndex = -1 Then
            cboְҵ.AddItem rsPatiInfo!ְҵ, 0
            cboְҵ.ListIndex = cboְҵ.NewIndex
        End If
    End If
    '����
    cbo����.ListIndex = gobjControl.Cbo.FindIndex(cbo����, Nvl(rsPatiInfo!����), True)
     If cbo����.ListIndex = -1 And Nvl(rsPatiInfo!����) <> "" Then
         cbo����.AddItem rsPatiInfo!����, 0
         cbo����.ListIndex = cbo����.NewIndex
     End If
    '����
    cbo����.ListIndex = gobjControl.Cbo.FindIndex(cbo����, Nvl(rsPatiInfo!����), True)
     If cbo����.ListIndex = -1 And Nvl(rsPatiInfo!����) <> "" Then
         cbo����.AddItem rsPatiInfo!����, 0
         cbo����.ListIndex = cbo����.NewIndex
     End If
    '����״��
    cbo����.ListIndex = gobjControl.Cbo.FindIndex(cbo����, Nvl(rsPatiInfo!����״��), True)
     If cbo����.ListIndex = -1 And Nvl(rsPatiInfo!����״��) <> "" Then
         cbo����.AddItem rsPatiInfo!����״��, 0
         cbo����.ListIndex = cbo����.NewIndex
     End If
    txt����.Text = Nvl(rsPatiInfo!����)
    '��ͥ��ַ
    cbo��ͥ��ַ.Text = Nvl(rsPatiInfo!��ͥ��ַ)
    '��ͥ�绰
    txt��ͥ�绰.Text = Nvl(rsPatiInfo!��ͥ�绰)
    '��ͥ��ַ�ʱ�
    txt��ͥ�ʱ�.Text = Nvl(rsPatiInfo!��ͥ��ַ�ʱ�)
    '���ڵ�ַ
    txtRegLocation.Text = Nvl(rsPatiInfo!���ڵ�ַ)
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
    cbo��ϵ�˹�ϵ.ListIndex = gobjControl.Cbo.FindIndex(cbo��ϵ�˹�ϵ, Nvl(rsPatiInfo!��ϵ�˹�ϵ), True)
    If cbo��ϵ�˹�ϵ.ListIndex = -1 And Nvl(rsPatiInfo!��ϵ�˹�ϵ) <> "" Then
        cbo��ϵ�˹�ϵ.ListIndex = 8: txt������ϵ.Text = Nvl(rsPatiInfo!��ϵ�˹�ϵ)
    End If
    '�����:56599
    Load�����������Ϣ (Val(Nvl(rsPatiInfo!����ID, "0")))
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
    Dim strSql As String
    Dim rsTemp As Recordset
    On Error GoTo Errhand
    strSql = "" & _
    " Select  ����,����� From ������Ϣ A,����ҽ�ƿ���Ϣ B Where A.����ID=B.����ID And B.����=[1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "ҽ�ƿ���", str���֤��)
    If rsTemp Is Nothing Then zl��ǰ�û����֤�Ƿ�� = False: Exit Function
    If rsTemp.RecordCount <= 0 Then zl��ǰ�û����֤�Ƿ�� = False: Exit Function
    
    If IIf(IsNull(rsTemp!����), "", rsTemp!����) = strName And IIf(IsNull(rsTemp!�����), "", rsTemp!�����) = str����� Then
        zl��ǰ�û����֤�Ƿ�� = True
    Else
        zl��ǰ�û����֤�Ƿ�� = False
    End If
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then Resume
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

        Set ObjItem = tbcPage.InsertItem(mPageIndex.����, "����", picInfo.hWnd, 0)
        ObjItem.Tag = mPageIndex.����
    
        Set ObjItem = tbcPage.InsertItem(mPageIndex.��������, "��������", PicHealth.hWnd, 0)
        ObjItem.Tag = mPageIndex.��������
        Call InitVsInoculate
        Call InitVsOtherInfo
        Call InitCombox
        
        '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
        If Not mobjPlugIn Is Nothing Then
            On Error Resume Next
            mlngPlugInHwnd = mobjPlugIn.GetFormHwnd
            Call zlPlugInErrH(Err, "GetFormHwnd")
            Err.Clear: On Error GoTo 0
            If mlngPlugInHwnd <> 0 Then
                picTaskPanelOther.Visible = True
                Set ObjItem = tbcPage.InsertItem(mPageIndex.������Ϣ, "������Ϣ", picTaskPanelOther.hWnd, 0)
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
    If gobjComlib.ErrCenter = 1 Then
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
    Dim strSql As String
    Dim varKey As Variant
    Dim intCount As Integer
    '����ҩ��
    With msh����
        If .Rows > 1 Then
            '����ò������м�¼
            strSql = " Zl_���˹���ҩ��_Delete(" & lng����ID & ")"
            zlAddArray colPro, strSql
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    '���˹���ҩ��
                    strSql = "Zl_���˹���ҩ��_Update("
                    '����ID_In ���˹���ҩ��.����Id%Type
                    strSql = strSql & "" & lng����ID & ","
                    '����ҩ��ID_In ���˹���ҩ��.����ҩ��ID%Type
                    strSql = strSql & "'" & IIf(.RowData(i) <= 0, "", .RowData(i)) & "',"
                    '����ҩ��_In  ���˹���ҩ��.����ҩ��%Type
                    strSql = strSql & "'" & IIf(.TextMatrix(i, 0) = "", "", .TextMatrix(i, 0)) & "',"
                    '������Ӧ_In ���˹�����Ӧ.������Ӧ%Type
                    strSql = strSql & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"

                    zlAddArray colPro, strSql
                End If
            Next
        End If
    End With
    '������Ϣ
    With vsInoculate
        If .Rows > 1 Then
            '����ò������м�¼
            strSql = " Zl_�������߼�¼_Delete(" & lng����ID & ")"
            zlAddArray colPro, strSql

            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    '���˹���ҩ��
                    strSql = "Zl_�������߼�¼_Update("
                    '����ID_In �������߼�¼.����Id%Type
                    strSql = strSql & "" & lng����ID & ","
                    '����ʱ��_In �������߼�¼.����ʱ��%Type
                    strSql = strSql & "" & IIf(.TextMatrix(i, 0) = "", "''", "to_date('" & .TextMatrix(i, 0) & "','yyyy-mm-dd')") & ","
                    '��������_In  �������߼�¼.��������%Type
                    strSql = strSql & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"
                    zlAddArray colPro, strSql
                End If
                If .TextMatrix(i, 3) <> "" Then
                    '���˹���ҩ��
                    strSql = "Zl_�������߼�¼_Update("
                    '����ID_In �������߼�¼.����Id%Type
                    strSql = strSql & "" & lng����ID & ","
                    '����ʱ��_In �������߼�¼.����ʱ��%Type
                    strSql = strSql & "" & IIf(.TextMatrix(i, 2) = "", "''", "to_date('" & .TextMatrix(i, 2) & "','yyyy-mm-dd')") & ","
                    '��������_In  �������߼�¼.��������%Type
                    strSql = strSql & "'" & IIf(.TextMatrix(i, 3) = "", "''", .TextMatrix(i, 3)) & "')"
                    zlAddArray colPro, strSql
                End If
            Next
        End If
    End With
    '������Ϣ
    'ABOѪ��
    '������Ϣ�ӱ�
    strSql = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSql = strSql & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSql = strSql & "'Ѫ��',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSql = strSql & "'" & gobjCommFun.GetNeedName(cboBloodType.Text, ".") & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSql = strSql & "'')"
    zlAddArray colPro, strSql
    'RH
    strSql = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSql = strSql & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSql = strSql & "'RH',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSql = strSql & "'" & cboBH.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSql = strSql & "'')"
    zlAddArray colPro, strSql
    'ҽѧ��ʾ
    strSql = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSql = strSql & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSql = strSql & "'ҽѧ��ʾ',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSql = strSql & "'" & txtMedicalWarning.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSql = strSql & "'')"
    zlAddArray colPro, strSql
    '����ҽѧ��ʾ
    strSql = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSql = strSql & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSql = strSql & "'����ҽѧ��ʾ',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSql = strSql & "'" & txtOtherWaring.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSql = strSql & "'')"
    zlAddArray colPro, strSql
        
    '84313:���ϴ�,2015/4/29, ��һ����ϵ����Ϣ�ѱ����ڲ�����Ϣ�У��ӱ��в����ظ�����
    '��ϵ�������Ϣ
    intCount = 0
    With vsLinkMan
        If .Rows >= 3 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then '��ϵ����������Ϊ��
                    intCount = intCount + 1
                    For j = 0 To .Cols - 1
                        strSql = "Zl_������Ϣ�ӱ�_Update("
                        '����ID_In ������Ϣ�ӱ�.����Id%Type
                        strSql = strSql & "" & lng����ID & ","
                        '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
                        strSql = strSql & "'��ϵ��" & .TextMatrix(0, j) & intCount & "',"
                        '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
                        strSql = strSql & "'" & IIf(.TextMatrix(i, j) = "", "", .TextMatrix(i, j)) & "',"
                        '����Id_In ������Ϣ�ӱ�.����Id%Type
                        strSql = strSql & "'')"

                        zlAddArray colPro, strSql
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
                    strSql = "Zl_������Ϣ�ӱ�_Update("
                    '����ID_In ������Ϣ�ӱ�.����Id%Type
                    strSql = strSql & "" & lng����ID & ","
                    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
                    strSql = strSql & "'" & .TextMatrix(i, 0) & "',"
                    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
                    strSql = strSql & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "',"
                    '����Id_In ������Ϣ�ӱ�.����Id%Type
                    strSql = strSql & "'')"

                    zlAddArray colPro, strSql
                End If
                If .TextMatrix(i, 2) <> "" Then
                    strSql = "Zl_������Ϣ�ӱ�_Update("
                    '����ID_In ������Ϣ�ӱ�.����Id%Type
                    strSql = strSql & "" & lng����ID & ","
                    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
                    strSql = strSql & "'" & .TextMatrix(i, 2) & "',"
                    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
                    strSql = strSql & "'" & IIf(.TextMatrix(i, 3) = "", "", .TextMatrix(i, 3)) & "',"
                    '����Id_In ������Ϣ�ӱ�.����Id%Type
                    strSql = strSql & "'')"

                    zlAddArray colPro, strSql
                End If
            Next
        End If
     End With
     If lng����ID = 0 Then Exit Sub
     'ABOѪ��
    '������Ϣ�ӱ�
    strSql = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSql = strSql & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSql = strSql & "'Ѫ��',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSql = strSql & "'" & gobjCommFun.GetNeedName(cboBloodType.Text, ".") & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSql = strSql & lng����ID & ")"
    zlAddArray colPro, strSql
    'RH
    strSql = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSql = strSql & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSql = strSql & "'RH',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSql = strSql & "'" & cboBH.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSql = strSql & lng����ID & ")"
    zlAddArray colPro, strSql
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
                        cbo��ϵ�˹�ϵ.ListIndex = gobjControl.Cbo.FindIndex(cbo��ϵ�˹�ϵ, str��ϵ, True)
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
        If gobjControl.Cbo.FindIndex(cbo��ϵ�˹�ϵ, str��ϵ, True) = -1 And str��ϵ <> "" Then
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
            cbo��ϵ�˹�ϵ.ListIndex = gobjControl.Cbo.FindIndex(cbo��ϵ�˹�ϵ, str��ϵ, True)
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
    Dim strSql As String
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
    strSql = "" & _
    "   Select ����ID,����ҩ��ID,����ҩ��,������Ӧ From ���˹���ҩ�� Where ����ID=[1]"
    Set rs����ҩ�� = gobjDatabase.OpenSQLRecord(strSql, "���˹���ҩ��", lng����ID)
    While rs����ҩ��.EOF = False
        SetDrugAllergy Nvl(rs����ҩ��!����ҩ��), Nvl(rs����ҩ��!������Ӧ), Nvl(rs����ҩ��!����ҩ��ID, 0)
        rs����ҩ��.MoveNext
    Wend
    '��ȡ���߼�¼
    strSql = "" & _
    "   Select ����ID,����ʱ��,�������� From �������߼�¼ Where ����ID=[1]"
    Set rs���߼�¼ = gobjDatabase.OpenSQLRecord(strSql, "�������߼�¼", lng����ID)
    While rs���߼�¼.EOF = False
        SetInoculate Nvl(rs���߼�¼!����ʱ��), Nvl(rs���߼�¼!��������)
        rs���߼�¼.MoveNext
    Wend
    'Ѫ��
    strSql = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='Ѫ��' And ����ID Is NULL"
    Set rsABOѪ�� = gobjDatabase.OpenSQLRecord(strSql, "ABOѪ��", lng����ID)
    While rsABOѪ��.EOF = False
        For i = 0 To cboBloodType.ListCount - 1
            '76314,���ϴ���2014-08-06��������Ϣ��ȷ��ȡ
            If gobjCommFun.GetNeedName(cboBloodType.List(i), ".") = NeedName(Nvl(rsABOѪ��!��Ϣֵ)) Then cboBloodType.ListIndex = i
        Next
        rsABOѪ��.MoveNext
    Wend
    'RH
    strSql = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='RH' And ����ID Is NULL"
    Set rsRH = gobjDatabase.OpenSQLRecord(strSql, "RH", lng����ID)
    While rsRH.EOF = False
        For i = 0 To cboBH.ListCount - 1
            If cboBH.List(i) = Nvl(rsRH!��Ϣֵ) Then cboBH.ListIndex = i
        Next
        rsRH.MoveNext
    Wend
    'ҽѧ��ʾ
    strSql = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='ҽѧ��ʾ'"
    Set rsҽѧ��ʾ = gobjDatabase.OpenSQLRecord(strSql, "ҽѧ��ʾ", lng����ID)
    While rsҽѧ��ʾ.EOF = False
        strҽѧ��ʾ = strҽѧ��ʾ & "," & Nvl(rsҽѧ��ʾ!��Ϣֵ)
        rsҽѧ��ʾ.MoveNext
    Wend
    If strҽѧ��ʾ <> "" Then strҽѧ��ʾ = Mid(strҽѧ��ʾ, 2)
    txtMedicalWarning.Text = strҽѧ��ʾ
    '����ҽѧ��ʾ
    strSql = "" & _
    "  Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='����ҽѧ��ʾ'"
    Set rs����ҽѧ��ʾ = gobjDatabase.OpenSQLRecord(strSql, "����ҽѧ��ʾ", lng����ID)
    While rs����ҽѧ��ʾ.EOF = False
        txtOtherWaring.Text = Nvl(rs����ҽѧ��ʾ!��Ϣֵ)
        rs����ҽѧ��ʾ.MoveNext
    Wend
    '��ϵ�������Ϣ
    'ȡ������Ϣ���е���ϵ����Ϣ
    '84313,���ϴ�,2015/4/27,��ϵ�˹�ϵ�Լ�������ϵ
    strSql = "" & _
    "   Select  A.��ϵ������,A.��ϵ�˹�ϵ,A.��ϵ�˵绰,A.��ϵ�����֤��,B.��Ϣֵ as ������Ϣ From ������Ϣ A,������Ϣ�ӱ� B " & _
    "   Where A.����ID=B.����ID(+) And A.����ID=[1] And B.��Ϣ��(+)='��ϵ�˸�����Ϣ' And Not A.��ϵ������ is Null"
    Set rs������Ϣ = gobjDatabase.OpenSQLRecord(strSql, "������Ϣ��ϵ����Ϣ", lng����ID)
    If rs������Ϣ.EOF = False Then
        txt��ϵ�����֤.Text = Nvl(rs������Ϣ!��ϵ�����֤��)
        txt��ϵ������.Text = Nvl(rs������Ϣ!��ϵ������)
        txt��ϵ�˵绰.Text = Nvl(rs������Ϣ!��ϵ�˵绰)
        cbo��ϵ�˹�ϵ.ListIndex = gobjControl.Cbo.FindIndex(cbo��ϵ�˹�ϵ, Nvl(rs������Ϣ!��ϵ�˹�ϵ), True)
        If cbo��ϵ�˹�ϵ.ListIndex = -1 And Nvl(rs������Ϣ!��ϵ�˹�ϵ) <> "" Then
            cbo��ϵ�˹�ϵ.ListIndex = 8: txt������ϵ.Text = rs������Ϣ!��ϵ�˹�ϵ
        ElseIf cbo��ϵ�˹�ϵ.ListIndex = 8 Then
            txt������ϵ.Text = Nvl(rs������Ϣ!������Ϣ)
        End If
        SetLinkInfo Nvl(rs������Ϣ!��ϵ������), Nvl(rs������Ϣ!��ϵ�˹�ϵ), Nvl(rs������Ϣ!��ϵ�˵绰), Nvl(rs������Ϣ!��ϵ�����֤��), txt������ϵ.Text
    End If
    'ȡ������Ϣ�ӱ��е���ϵ����Ϣ
    strSql = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ�� like '��ϵ��%' order by ��Ϣ�� Asc"
    Set rs��ϵ�� = gobjDatabase.OpenSQLRecord(strSql, "��ϵ�������Ϣ", lng����ID)
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
    strSql = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ�� Not in ('Ѫ��','ABO','RH','ҽѧ��ʾ','����ҽѧ��ʾ') And ��Ϣ�� Not like '��ϵ��%'"
    Set rs������Ϣ = gobjDatabase.OpenSQLRecord(strSql, "��ϵ��������Ϣ", lng����ID)
    '�����:115886,����,2017/11/08,�Һ���ȡ�ò�����Ϣʱ�����򱨴�
    While rs������Ϣ.EOF = False
        If Nvl(rs������Ϣ!��Ϣ��) <> "" Then
            SetOtherInfo Nvl(rs������Ϣ!��Ϣ��), Nvl(rs������Ϣ!��Ϣֵ)
        End If
        rs������Ϣ.MoveNext
    Wend
    
    Exit Sub
ErrHandl:
     If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

'Private Function bln����(Optional ByVal blnCardNo As Boolean = False) As Boolean
''---------------------------------------------------------------------------------------------------------------------------------------------
''����:�жϵ�ǰ�Ƿ�Ϊ�������� (���Ƿ����������ǰ󶨿�����)
''���:
''����:56599
''����:2012-12-12 14:55:36
''---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim lng�ſ�����ID As Long
'    Dim bln�Ƿ񷢿� As Boolean
'    If gCurSendCard.bln�ϸ���� = True Then
'        lng�ſ�����ID = CheckUsedBill(5, IIf(lng�ſ�����ID > 0, lng�ſ�����ID, gCurSendCard.lng��������), IIf(blnCardNo, mstrCard, UCase(txtPatient.Text)), gCurSendCard.lng�����ID)
'        bln�Ƿ񷢿� = IIf(lng�ſ�����ID = -3, False, True)
'        If gCurSendCard.bln���ƿ� = False Then
'            bln�Ƿ񷢿� = (gCurSendCard.bln�Ƿ񷢿� = True)
'        End If
'    Else
'        bln�Ƿ񷢿� = mbln����
'        If gCurSendCard.bln���ƿ� = False Then
'            bln�Ƿ񷢿� = (gCurSendCard.bln�Ƿ񷢿� = True)
'        End If
'    End If
'    bln���� = bln�Ƿ񷢿�
'    mbln���� = bln�Ƿ񷢿�
'End Function

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
'����:ȡ���󶨿�
'���:intType:0-��ǰ����;1-��ǰ���;2-��ǰ��������
'����:ȡ���ɹ�,����true,���򷵻�False
'����:���˺�
'����:2011-07-29 11:18:05
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim curDate As Date
    Dim strSql As String, strPassWord As String

    On Error GoTo errHandle

    curDate = gobjDatabase.Currentdate
    
    'Zl_ҽ�ƿ��䶯_Insert
    strSql = "Zl_ҽ�ƿ��䶯_Insert("
    '      �䶯����_In   Number,
    '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
    strSql = strSql & "" & 14 & ","
    '      ����id_In     סԺ���ü�¼.����id%Type,
    strSql = strSql & "" & lngPatientID & ","
    '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
    strSql = strSql & "" & lngCardTypeID & ","
    '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
    strSql = strSql & "NULL,"
    '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
    strSql = strSql & "'" & strCardNO & "'" & ","
    '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
    strSql = strSql & "'�ҺŰ󶨿��Զ�ȡ����',"
    '      ����_In       ������Ϣ.����֤��%Type,
    strSql = strSql & "NULL,"
    '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
    strSql = strSql & "NULL,"
    '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
    strSql = strSql & "to_date('" & Format(curDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '      Ic����_In     ������Ϣ.Ic����%Type := Null,
    strSql = strSql & "NULL,"
    '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
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

Private Function CreateObjectPlugIn() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������������Ϣ���
    '����:�����ɹ�,����True,���򷵻�False
    '�����:73935
    '����:Ƚ����
    '����:2014-07-3
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
    '����:���ظ�����Ϣҳ��
    '����:
    '�����:73935
    '����:Ƚ����
    '����:2014-07-3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup, Item As TaskPanelGroupItem
    Dim lngHwnd As Long
    
    Err = 0: On Error GoTo Errhand
    If Not mobjPlugIn Is Nothing Then
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
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Sub DeletePatPicture(lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��������Ƭ
    '���:lng����ID - ����ID
    '����:56599
    '����:2012-12-14 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo Errhand:
    strSql = strSql & "Zl_������Ƭ_Delete("
    strSql = strSql & lng����ID & ")"
    
    gobjDatabase.ExecuteProcedure strSql, Me.Caption
    
    Exit Sub
Errhand:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub SavePatPicture(lng����ID As Long, strFile As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���没����Ƭ
    '���:lng����ID - ����ID
    '����:56599
    '����:2012-12-13 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
        
    If strFile = "" Then Exit Sub

    If gobjComlib.Sys.SaveLob(glngSys, 27, lng����ID, strFile, 0) = False Then
        ShowMsgBox "������Ƭ����,��ȷ���ļ��Ƿ�ɾ��!"
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

    strTmp = gobjComlib.Sys.ReadLob(glngSys, 27, lng����ID)
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
    If gobjComlib.ErrCenter() = 1 Then Resume
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

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

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

Private Sub zlQueryEMPIPatiInfo()
    '���ܣ���EMPIƽ̨��ȡ������Ϣ
    '���ڣ�2016/10/9 10:47:13
    '���ƣ����ϴ�
    '˵����101170
    Dim rsTmp As ADODB.Recordset, strDiff As String, strMsgInfo As String
    If mblnPlugin = False Then Exit Sub
    If mobjPlugIn Is Nothing Then Exit Sub
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
        !�Ա� = NeedName(cbo�Ա�.Text)
        If IsDate(txt��������.Text) Then
            !�������� = Format(txt��������.Text & " " & IIf(IsDate(txt����ʱ��.Text), txt����ʱ��.Text, "00:00"), "YYYY-MM-DD HH:MM")
        Else
            !�������� = ""
        End If
        !�����ص� = txtBirthLocation.Text
        !���� = NeedName(cbo����.Text)
        !���� = NeedName(cbo����.Text)
        !ְҵ = NeedName(cboְҵ.Text)
        !������λ = txt��λ����.Text
        !����״�� = NeedName(cbo����.Text)
        !��ͥ�绰 = txt��ͥ�绰.Text
        !��ϵ�˵绰 = txt��ϵ�˵绰.Text
        !��λ�绰 = txt��λ�绰.Text
        !��ͥ��ַ = cbo��ͥ��ַ.Text
        !��ͥ��ַ�ʱ� = txt��ͥ�ʱ�.Text
        !���ڵ�ַ = txtRegLocation.Text
        !���ڵ�ַ�ʱ� = txt���ڵ�ַ�ʱ�.Text
        !��λ�ʱ� = txt��λ�ʱ�.Text
        !��ϵ������ = txt��ϵ������.Text
        !��ϵ�˹�ϵ = NeedName(cbo��ϵ�˹�ϵ.Text)
        .Update
    End With
    'EMPIû���ҵ�������Ϣ,ֱ�ӷ���
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
        If Nvl(!ҽ����) <> "" Then txtPatiMCNO(0).Text = Nvl(!ҽ����): txtPatiMCNO(1).Text = Nvl(!ҽ����)
        If Nvl(!���֤��) <> "" Then txt���֤��.Text = Nvl(!���֤��)
        If mlng����ID = 0 Then
            If Nvl(!����) <> "" Then txtPatient.Text = Nvl(!����)
            If Nvl(!�Ա�) <> "" Then cbo�Ա�.ListIndex = gobjControl.Cbo.FindIndex(cbo�Ա�, Nvl(!�Ա�), True)
            If Nvl(!��������) <> Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS") Then
                txt��������.Text = Format(Nvl(!��������), "YYYY-MM-DD")
                txt����ʱ��.Text = Format(Nvl(!��������), "HH:MM")
            End If
        Else
            If Nvl(!����) <> "" And txtPatient.Text <> Nvl(!����) Then strDiff = ",����"
            If Nvl(!�Ա�) <> "" And cbo�Ա�.ListIndex <> gobjControl.Cbo.FindIndex(cbo�Ա�, Nvl(!�Ա�), True) Then strDiff = strDiff & ",�Ա�"
            If Nvl(!��������) <> "" And Format(Nvl(!��������), "YYYY-MM-DD HH:MM:SS") <> Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS") Then strDiff = strDiff & ",��������"
            If Nvl(!���֤��) <> "" And txt���֤��.Text <> Nvl(!���֤��) Then strDiff = strDiff & ",���֤��"
        End If
        
        If InStr(gstrPrivs, ";�����޸������;") > 0 Or mlng����ID = 0 Then
            If Nvl(!�����) <> "" Then txt�����.Text = Nvl(!�����)
        Else
            If Nvl(!�����) <> "" Then txt�����.Text = Nvl(!�����)
        End If
        
        If Nvl(!�����ص�) <> "" Then txtBirthLocation.Text = Nvl(!�����ص�)
        If Nvl(!����) <> "" Then cbo����.ListIndex = gobjControl.Cbo.FindIndex(cbo����, Nvl(!����), True)
        If Nvl(!����) <> "" Then cbo����.ListIndex = gobjControl.Cbo.FindIndex(cbo����, Nvl(!����), True)
        If Nvl(!ְҵ) <> "" Then cboְҵ.ListIndex = gobjControl.Cbo.FindIndex(cboְҵ, Nvl(!ְҵ))
        If Nvl(!������λ) <> "" Then txt��λ����.Text = Nvl(!������λ)
        If Nvl(!����״��) <> "" Then cbo����.ListIndex = gobjControl.Cbo.FindIndex(cbo����, Nvl(!����״��), True)
        If Nvl(!��ͥ�绰) <> "" Then txt��ͥ�绰.Text = Nvl(!��ͥ�绰)
        If Nvl(!��ϵ�˵绰) <> "" Then txt��ϵ�˵绰.Text = Nvl(!��ϵ�˵绰)
        If Nvl(!��λ�绰) <> "" Then txt��λ�绰.Text = Nvl(!��λ�绰)
        If Nvl(!��ͥ��ַ) <> "" Then cbo��ͥ��ַ.Text = Nvl(!��ͥ��ַ): padd��ͥ��ַ.Value = Nvl(!��ͥ��ַ)
        If Nvl(!��ͥ��ַ�ʱ�) <> "" Then txt��ͥ�ʱ�.Text = Nvl(!��ͥ��ַ�ʱ�)
        If Nvl(!���ڵ�ַ) <> "" Then txtRegLocation.Text = Nvl(!���ڵ�ַ): padd���ڵ�ַ.Value = Nvl(!���ڵ�ַ)
        If Nvl(!���ڵ�ַ�ʱ�) <> "" Then txt���ڵ�ַ�ʱ�.Text = Nvl(!���ڵ�ַ�ʱ�)
        If Nvl(!��λ�ʱ�) <> "" Then txt��λ�ʱ�.Text = Nvl(!��λ�ʱ�)
        If Nvl(!��ϵ������) <> "" Then txt��ϵ������.Text = Nvl(!��ϵ������)
        If Nvl(!��ϵ�˹�ϵ) <> "" Then cbo��ϵ�˹�ϵ.ListIndex = gobjControl.Cbo.FindIndex(cbo��ϵ�˹�ϵ, Nvl(!��ϵ�˹�ϵ), True)
    End With
    Err = 0: On Error GoTo 0
    If mlng����ID <> 0 Then
        If strDiff <> "" Then strDiff = Mid(strDiff, 2)
        If strDiff <> "" Then
            strMsgInfo = "���˵� " & strDiff & " ��EMPI��Ϣ��һ�£�������������Ӧ��Ȩ�ޣ����β�����и��¡�"
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
    '����:�ϴ�������Ϣ��EMPIƽ̨,���ƽ̨��Ϣ����ʧ�ܣ���ͬHIS����һ�����
    '����: In-lngPatiID ����ID,lngClinicID �Һ�ID
    '      Out-strErrMsg ������Ϣ����������ʧ����Ч
    '����:True-EMPIƽ̨������Ϣ�ɹ�,False-����ʧ��
    '����:���ϴ�
    '˵��:100915
    Dim blnCharge As Boolean, lngRet As Long
    If mblnPlugin = False Then zlSaveEMPIPatiInfo = True: Exit Function
    If mobjPlugIn Is Nothing Then zlSaveEMPIPatiInfo = True: Exit Function
    
    On Error GoTo Errhand
    If mrsEMPIOut Is Nothing Then
        'EMPIû�в�����Ϣ����Ҫ�½�
        On Error Resume Next
        lngRet = mobjPlugIn.EMPI_AddPatiInfo(glngSys, glngModul, lngPatiID, 0, lngClinicID, strErrMsg)
        Call zlPlugInErrH(Err, "EMPI_AddPatiInfo")
        If lngRet = 0 And Err.Number <> 438 Then Err.Clear: Exit Function
        Err.Clear: On Error GoTo Errhand
    Else
        '�ж�ƽ̨�ش�����Ϣ�Ƿ����ı�
        With mrsEMPIOut
            If txtPatiMCNO(0).Text <> Nvl(!ҽ����) Then blnCharge = True: GoTo EMPIModify
            If blnNewPati Then
                If txt���֤��.Text <> Nvl(!���֤��) Then blnCharge = True: GoTo EMPIModify
                If txtPatient.Text <> Nvl(!����) Then blnCharge = True: GoTo EMPIModify
                If cbo�Ա�.ListIndex <> gobjControl.Cbo.FindIndex(cbo�Ա�, Nvl(!�Ա�), True) Then blnCharge = True: GoTo EMPIModify
                If Format(txt��������.Text, "YYYY-MM-DD") <> Format(Nvl(!��������), "YYYY-MM-DD") Then blnCharge = True: GoTo EMPIModify
                If Format(txt����ʱ��.Text, "HH:MM") <> Format(Nvl(!��������), "HH:MM") Then blnCharge = True: GoTo EMPIModify
            End If
            
            If InStr(gstrPrivs, ";�����޸������;") > 0 Or blnNewPati Then
                If txt�����.Text <> Nvl(!�����) Then blnCharge = True: GoTo EMPIModify
            End If
            If txtBirthLocation.Text <> Nvl(!�����ص�) Then blnCharge = True: GoTo EMPIModify
            If cbo����.ListIndex <> gobjControl.Cbo.FindIndex(cbo����, Nvl(!����), True) Then blnCharge = True: GoTo EMPIModify
            If cbo����.ListIndex <> gobjControl.Cbo.FindIndex(cbo����, Nvl(!����), True) Then blnCharge = True: GoTo EMPIModify
            If cboְҵ.ListIndex <> gobjControl.Cbo.FindIndex(cboְҵ, Nvl(!ְҵ)) Then blnCharge = True: GoTo EMPIModify
            If txt��λ����.Text <> Nvl(!������λ) Then blnCharge = True: GoTo EMPIModify
            If cbo����.ListIndex <> gobjControl.Cbo.FindIndex(cbo����, Nvl(!����״��), True) Then blnCharge = True: GoTo EMPIModify
            If txt��ͥ�绰.Text <> Nvl(!��ͥ�绰) Then blnCharge = True: GoTo EMPIModify
            If txt��ϵ�˵绰.Text <> Nvl(!��ϵ�˵绰) Then blnCharge = True: GoTo EMPIModify
            If txt��λ�绰.Text <> Nvl(!��λ�绰) Then blnCharge = True: GoTo EMPIModify
            If cbo��ͥ��ַ.Text <> Nvl(!��ͥ��ַ) Then blnCharge = True: GoTo EMPIModify
            If txt��ͥ�ʱ�.Text <> Nvl(!��ͥ��ַ�ʱ�) Then blnCharge = True: GoTo EMPIModify
            If txtRegLocation.Text <> Nvl(!���ڵ�ַ) Then blnCharge = True: GoTo EMPIModify
            If txt���ڵ�ַ�ʱ�.Text <> Nvl(!���ڵ�ַ�ʱ�) Then blnCharge = True: GoTo EMPIModify
            If txt��λ�ʱ�.Text <> Nvl(!��λ�ʱ�) Then blnCharge = True: GoTo EMPIModify
            If txt��ϵ������.Text <> Nvl(!��ϵ������) Then blnCharge = True: GoTo EMPIModify
            If cbo��ϵ�˹�ϵ.ListIndex <> gobjControl.Cbo.FindIndex(cbo��ϵ�˹�ϵ, Nvl(!��ϵ�˹�ϵ), True) Then blnCharge = True: GoTo EMPIModify
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
