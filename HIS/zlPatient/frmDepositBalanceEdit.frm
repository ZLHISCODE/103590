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
   Caption         =   "Ԥ�����"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13200
   BeginProperty Font 
      Name            =   "����"
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
   StartUpPosition =   1  '����������
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
         Caption         =   "��    ע "
         Height          =   240
         Left            =   6345
         TabIndex        =   30
         Tag             =   "��    ע "
         Top             =   1995
         Width           =   1080
      End
      Begin VB.Label lblWorkUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������λ"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6345
         TabIndex        =   29
         Tag             =   "������λ "
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
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   3050
         TabIndex        =   28
         Tag             =   "���֤�� "
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
      Begin VB.Label lbl�ֻ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� ��"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   27
         Tag             =   "�� �� �� "
         Top             =   1950
         Width           =   960
      End
      Begin VB.Label lblδ�ɷ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ�ɷ��� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   3060
         TabIndex        =   26
         Tag             =   "δ�ɷ��� "
         ToolTipText     =   "δ�ɿ�Ļ��۵����úϼ�"
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
      Begin VB.Label lblҽ��Ԥ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ��Ԥ�� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6345
         TabIndex        =   25
         Tag             =   "ҽ��Ԥ�� "
         ToolTipText     =   "ҽ��Ԥ����"
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
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ���� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   9390
         TabIndex        =   24
         Tag             =   "סԺ���� "
         Top             =   471
         Width           =   1080
      End
      Begin VB.Label lblδ����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ����� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   23
         Tag             =   "δ����� "
         ToolTipText     =   "δ��˵Ļ��ۼ��˷��úϼ�"
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Label lblӦ�տ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӧ �� �� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   9390
         TabIndex        =   22
         Tag             =   "Ӧ �� �� "
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Label lblҽ�Ƹ��ʽ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ƹ��ʽ "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   8910
         TabIndex        =   21
         Tag             =   "ҽ�Ƹ��ʽ "
         Top             =   75
         Width           =   1560
      End
      Begin VB.Label lbl��ͥ��ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Tag             =   "��ͥ��ַ "
         Top             =   471
         Width           =   1080
      End
      Begin VB.Label lbl������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   3060
         TabIndex        =   19
         Tag             =   "������� "
         Top             =   1605
         Width           =   1080
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� �� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Tag             =   "�� �� �� "
         Top             =   1605
         Width           =   1080
      End
      Begin VB.Label lbl�ѱ�ȼ� 
         AutoSize        =   -1  'True
         Caption         =   "�ѱ� "
         Height          =   240
         Left            =   5925
         TabIndex        =   17
         Tag             =   "�ѱ� "
         Top             =   90
         Width           =   600
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   2040
         TabIndex        =   16
         Tag             =   "���� "
         Top             =   90
         Width           =   600
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Tag             =   "�Ա� "
         Top             =   105
         Width           =   600
      End
      Begin VB.Label lblԤ����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ����� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   3060
         TabIndex        =   14
         Tag             =   "Ԥ����� "
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   3840
         TabIndex        =   13
         Tag             =   "���� "
         Top             =   90
         Width           =   600
      End
      Begin VB.Label lblʣ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʣ���� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6345
         TabIndex        =   12
         Tag             =   "ʣ���� "
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Label lbl������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ����� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   9390
         TabIndex        =   11
         Tag             =   "δ����� "
         ToolTipText     =   "δ��˵Ļ��ۼ��˷��úϼ�"
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lbl�ʻ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ʻ���� "
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Tag             =   "�ʻ���� "
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
         Caption         =   "ƾ����ӡ����(&V)"
         Height          =   420
         Left            =   3960
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   60
         Width           =   2025
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   420
         Left            =   150
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   60
         Width           =   1500
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   420
         Left            =   11625
         TabIndex        =   60
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   45
         Width           =   1500
      End
      Begin VB.CommandButton cmdSetup 
         Caption         =   "�վݴ�ӡ����(&S)"
         Height          =   420
         Left            =   1770
         TabIndex        =   63
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F10"
         Top             =   60
         Width           =   2025
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   420
         Left            =   10050
         TabIndex        =   58
         ToolTipText     =   "�ȼ���F2"
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
         ToolTipText     =   "�ȼ���F3"
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
         ToolTipText     =   "�ȼ���F12"
         Top             =   90
         Width           =   2100
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ�ݺ�"
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
         Caption         =   "Ԥ�����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "���ݺ�"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
         Begin VB.TextBox txt������ 
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
         Begin VB.TextBox txt�ʺ� 
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
         Begin MSMask.MaskEdBox txt�տ� 
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
               Name            =   "����"
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
            Caption         =   "ժҪ"
            Height          =   240
            Left            =   540
            TabIndex        =   76
            Top             =   3975
            Width           =   480
         End
         Begin VB.Label lblCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������"
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
            Caption         =   "�˿�"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl�ɿλ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ɿλ"
            Height          =   240
            Left            =   60
            TabIndex        =   73
            Top             =   3540
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   240
            Left            =   300
            TabIndex        =   72
            Top             =   2160
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ʺ�"
            Height          =   240
            Left            =   540
            TabIndex        =   71
            Top             =   2595
            Width           =   480
         End
         Begin VB.Label lblUnit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ɿ����"
            Height          =   240
            Left            =   60
            TabIndex        =   70
            Top             =   3060
            Width           =   960
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�˿�ϼ�"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "���ֺϼ�"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "�����˿�"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl�Ҳ� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ҳ�"
            BeginProperty Font 
               Name            =   "����"
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
         ToolTipText     =   "�ȼ���F11"
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
            Name            =   "����"
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
            Caption         =   "ȫ��(&A)"
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
               Name            =   "����"
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
               Name            =   "����"
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
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "����"
         IDKind          =   -1
         BackColor       =   -2147483633
      End
      Begin VB.Label lblPatientNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   6060
         TabIndex        =   61
         Top             =   180
         Width           =   840
      End
      Begin VB.Label lblPatiPage 
         AutoSize        =   -1  'True
         Caption         =   "סԺ����"
         Height          =   240
         Left            =   10730
         TabIndex        =   59
         Top             =   180
         Width           =   960
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   120
         TabIndex        =   42
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lblԤ������ 
         AutoSize        =   -1  'True
         Caption         =   "Ԥ������"
         Height          =   240
         Left            =   3615
         TabIndex        =   41
         Top             =   180
         Width           =   960
      End
      Begin VB.Label lblThirdSummary 
         Caption         =   "���������˿����"
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
Option Explicit 'Ҫ���������
'��ڲ���----------------------------------------------------------------------------------
Private mstrPrivs As String
Private mlngModul As Long
Private mbytCallObject As Byte '���õĶ���(0-Ԥ����������;1-���˷��ò�ѯ����;2-ҽ�ƿ��������;3-�Һ�ģ�����
Private mlng����ID As Long, mlng��ҳID As Long, mdblDefPreMoney As Double
Private mbytPrepayType As Byte   ' 1-����Ԥ��;2-סԺԤ��(4ʱ,1,����תסԺ;2ʱסԺת����)
Private mblnNotClick As Boolean
'�������----------------------------------------------------------------------------------
Private mblnUnLoad  As Boolean '���ڿ��ƴ���ֱ���˳�
Private mdblʣ���� As Double
Private mdblԤ����� As Double
Private mdbl������� As Double
Private mlng����ID As Long, mstrCardPrivs As String
Private mstrȱʡ���㷽ʽ As String
Private mblnOK As Boolean, mstr�˿����Ա As String
Private mblnδ��Ʋ���Ԥ�� As Boolean '51628
Private mblnסԺ��Ԥ����֤ As Boolean   '63113:������,2013-10-29,סԺԤ���˿���֤
Private mbln������Ժ��������˿� As Boolean
Private mblnNurseCall As Boolean
Private mblnFirst As Boolean
Private Enum BalanceType
    C1�ֽ� = 1
    C2���ֽ� = 2
    C3�����ʻ� = 3
    C4ҽ��ͳ�� = 4
    C5���տ� = 5
End Enum

'ҽ������----------------------
Private mcur�ʻ���� As Currency '�����ʻ����
Private mstr�����ʻ� As String '�����ʻ����㷽ʽ
Private mstr�������� As String
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
'���ڽ��㿨�ĵĴ������
Private Type Ty_SquareCard
    blnExistsObjects As Boolean     '��װ�˽��㿨�ĵ�
    dblˢ���ܶ� As Double
    bln������ As Boolean '��ǰ��ȡ�ĵ����ǿ�����
End Type
Private mtySquareCard As Ty_SquareCard

Private mstrBrushCardNo As String

Private Type Ty_BillInfor
    lngԤ��ID As Long
    strNO As String
    lng�����ID As Long
    bln���ѿ� As Boolean
    str���� As String
    str���� As String
    str������ˮ�� As String
    str����˵�� As String
    str������λ As String
    dbl��� As Double
    blnת�� As Boolean
    bln�˿��鿨 As Boolean
    dt�տ�ʱ�� As Date
    lng���ѿ�ID As Long
End Type
Private mcurBill As Ty_BillInfor
Private mFactProperty As Ty_FactProperty
Private mblnStartFactUseType As Boolean '�Ƿ����õ���ص���������
Private mrsDepositBalance As ADODB.Recordset    '��ǰ���˵�Ԥ�����
Private mbytBackMoneyType As Byte '�˿ʽ:1-��ֹ;0-��ʾ
Private mbytOracleBackType As Byte '�˿���_In;0-�����˿����Ƿ�����˲�����1-����˿���
Private mblnClearWinInfor As Boolean  '�ɿ��,�Ƿ����������Ϣ
Private mblnCheckPass As Boolean 'ˢ��ʱҪ����������,'0000000000'��λ˳���ʾ��������,�ֱ�Ϊ:1.����Һ�,2.���ﻮ��,3.�����շ�,4.�������,5.��Ժ�Ǽ�,6.סԺ����,7.���˽���,8.����Ԥ����,9.���鼼ʦվ,10.Ӱ��ҽ��վ.'
Private mbln�ų�δ�ɼ�δ�� As Boolean 'ʣ����ų�δ�ɼ�δ����
'�������������
Private mobjPlugIn As Object
Private mstrPatiOld As String
Private mstrPatiSex As String
Private mblnOneCard As Boolean  '�Ƿ�ֻ��һ�ž��￨
Private mlngFactModule As Long '��Ʊ��ز���ģ���
Private mblnOptErrBill As Boolean '�շ�ģʽ�´����쳣����
Private mobjThridSwap As clsThirdSwap
Private mobjPtDelItems As clsBalanceItems
Private Enum pg_Page
    pg_Ԥ������˿� = 1
    pg_Ԥ����ʷ��¼ = 2
End Enum

Private Enum PaneId
    EM_Head = 1
    EM_PatiInfo = 2
    EM_BillList = 3
    EM_Cmd = 4
End Enum
Private mpatiInfo As New clsPatientInfo
Private mobjCards As New Cards  'zlOneCardComLib.Cards
Private mobjEInvoice As clsEinvoice '����Ʊ�ݲ���
Private mintPrintType As Integer

Private Sub zlInitBalanceGrid(Optional bln�鿴 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�������б�
    '����:���˺�
    '����:2015-01-23 14:14:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    With vsBlance
    
        For i = 1 To .Rows - 1
            .RowData(i) = ""
        Next
        .Clear: .Rows = 2: i = 0: .Cols = 23
        .TextMatrix(0, i) = "�����ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "���ѿ�ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "��������": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�༭״̬": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 600: i = i + 1
        .TextMatrix(0, i) = "����״̬": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�����": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ȫ��": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "У�Ա�־": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�����": .ColWidth(i) = 0: i = i + 1
        
        .TextMatrix(0, i) = "���ݺ�": .ColWidth(i) = 1500: i = i + 1
        .TextMatrix(0, i) = "�˿ʽ": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "Ԥ�����": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "�˿���": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "�������": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "���������": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "������ˮ��": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "����˵��": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "��ע": .ColWidth(i) = 2500: i = i + 1
        .TextMatrix(0, i) = "��������ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ת��": .ColWidth(i) = 0: i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = flexAlignLeftCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            
            'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
            Select Case .ColKey(i)
            Case "�Ƿ�ת��", "��������ID", "��������", "����", "�Ƿ񱣴�", "�Ƿ�����", "У�Ա�־", "�༭״̬", "�Ƿ�����", "�Ƿ�ȫ��", "����״̬", "�Ƿ���֤"
                .ColHidden(i) = True
                .ColData(i) = "-1||1"
            Case "����"
                .ColAlignment(i) = flexAlignRightCenter
                .ColData(i) = "1||0"
                .ColDataType(i) = flexDTBoolean
            Case "�˿���"
                .ColAlignment(i) = flexAlignRightCenter
                .ColData(i) = "1||0"
            Case "Ԥ�����"
                .ColAlignment(i) = flexAlignRightCenter
                .ColData(i) = "0||0"
            Case .ColIndex("�˿ʽ")
                .ColData(i) = """1||0"
            Case "���������"
                .ColData(i) = "1||2"
            Case .ColIndex("�������")
                .ColData(i) = "1||0"
            Case Else
                .ColData(i) = "1||" & IIf(bln�鿴, "0", "2")
            End Select
            If bln�鿴 Then .ColData(i) = ""
        Next
        If Not bln�鿴 Then .Editable = flexEDKbdMouse
        .ExplorerBar = flexExMove
    End With
    zl_vsGrid_Para_Restore mlngModul, vsBlance, Me.Name, "�����б�"
    vsBlance.ColWidth(vsBlance.ColIndex("����")) = 600
    vsBlance.ColHidden(vsBlance.ColIndex("����")) = False
    
    With vsDepositHistory
        .Clear: .Rows = 2: i = 0: .Cols = 7
        .TextMatrix(0, i) = "����": .ColWidth(i) = 1350: i = i + 1
        .TextMatrix(0, i) = "���ݺ�": .ColWidth(i) = 1110: i = i + 1
        .TextMatrix(0, i) = "Ʊ�ݺ�": .ColWidth(i) = 1110: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 1200: i = i + 1
        .TextMatrix(0, i) = "�ɿ���": .ColWidth(i) = 1200: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 1000: i = i + 1
        .TextMatrix(0, i) = "�տ���": .ColWidth(i) = 1000: i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            Select Case .ColKey(i)
            Case "����", "���ݺ�", "Ʊ�ݺ�"
                .ColAlignment(i) = flexAlignCenterCenter
            Case "�ɿ���"
                 .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
    End With
    zl_vsGrid_Para_Restore mlngModul, vsDepositHistory, Me.Name, "Ԥ���嵥"
    zl_vsGrid_Para_Restore mlngModul, vsThirdTotal, Me.Name, "�����˿����"
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlShowEdit(ByVal frmMain As Object, ByVal bytCallObject As Byte, ByVal objEInvoice As clsEinvoice, _
    ByVal strPrivs As String, ByVal lngModule As Long, Optional ByVal bytPrepayType As Byte = 0, _
    Optional ByVal lng����id As Long = 0, Optional lng��ҳID As Long = 0, Optional ByVal blnNurseCall As Boolean = False, _
    Optional blnOneCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������,���ڲ���Ԥ������Ϣ�༭��鿴
    '���:frmMain-���õ�������
    '        bytCallObject:���õĶ���(0-Ԥ����������;1-���˷��ò�ѯ����;2-ҽ�ƿ�����,3-����Һŵ���)...
    '        bytPrepayType-Ԥ������(0-�����סԺ;1-����;2-סԺ)
    '        strInNo:Ҫ������˿�ĵ��ݺ�(mbytInState=1��3ʱ��Ч),�Ӳ�����Ϣ�Ǽ��е����˿�ʱΪ��
    '        blnNurseCall-��ʿվ����
    '����:
    '����:Ԥ����ֻ��һ�γɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-02-17 16:11:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    mbytCallObject = bytCallObject:   mstrPrivs = strPrivs: mlngModul = lngModule
    mlng����ID = lng����id: mlng��ҳID = lng��ҳID
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
    If txtPatient.Tag <> "" And Not mpatiInfo.����ID = 0 Then
        If cboPatiPage.ItemData(cboPatiPage.ListIndex) <> Val(cboPatiPage.Tag) Then
            cboPatiPage.Tag = cboPatiPage.ItemData(cboPatiPage.ListIndex)
            Call ShowPatiInfoFromPage
            Call ShowPremayBalance(True, mpatiInfo.����ID)
            Call LoadThirdDelDeposit(Val(cboPatiPage.ItemData(cboPatiPage.ListIndex)))
        End If
    End If
    Call ShowHistoryPrepay
End Sub

Private Sub ShowPatiInfoFromPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ�����ҳ��ʾ������Ϣ
    '����:���˺�
    '����:2018-11-29 09:46:39
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim lng��ҳID As Long
    
    If cboType.ListIndex < 0 Then Exit Sub
    If cboType.ItemData(cboType.ListIndex) = 1 Then Exit Sub    '����Ԥ������������ҳ
    If cboPatiPage.ListIndex < 0 Then Exit Sub  '����ҳʱ��������
    lng��ҳID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    
    '���ݵڼ�����Ժ������Ϣ
    Call GetPatient(IDKind.GetfaultCard, txtPatient.Tag, False, False, txtPatient.Tag, lng��ҳID)
    '���ز�����Ϣ���ؼ�
    Call LoadPatiInforToContronl
    
 End Sub
Private Sub LoadPatiInforToContronl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ϣ���ؼ�
    '����:���˺�
    '����:2018-11-29 09:51:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str������ As String, dbl������ As Double
     
    On Error GoTo errHandle
    
    If mpatiInfo.����ID = 0 Then Exit Sub
    
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(mpatiInfo.סԺ�� = "", "", "סԺ��:" & mpatiInfo.סԺ�� & "  ") & _
                       IIf(mpatiInfo.����� = "", "", "�����:" & mpatiInfo.�����)
                       
    txtPatient.IMEMode = 0: txtPatient.PasswordChar = ""    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.Text = mpatiInfo.����: txtPatient.Tag = mpatiInfo.����ID
    
    lblSex.Caption = lblSex.Tag & mpatiInfo.�Ա�: mstrPatiSex = mpatiInfo.�Ա�
    lblOld.Caption = lblOld.Tag & mpatiInfo.����: mstrPatiOld = mpatiInfo.����
     
    lblҽ�Ƹ��ʽ.Caption = lblҽ�Ƹ��ʽ.Tag & mpatiInfo.ҽ�Ƹ��ʽ
    
    
    lbl����.Caption = lbl����.Tag
    If mpatiInfo.��ǰ����ID <> 0 Then
        lbl����.Caption = lbl����.Tag & IIf(mpatiInfo.���� = "", "��ͥ", mpatiInfo.����)
    End If
    
    lbl����.Caption = lbl����.Tag & GET��������(mpatiInfo.��Ժ����ID)
    
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(mpatiInfo.��ǰ����ID = 0, mpatiInfo.��Ժ����ID, mpatiInfo.��ǰ����ID))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    
    'ҽ���Ķ�-��Ժ����ת�����ʻ�
    If mpatiInfo.���� > 0 And InStr(mstrPrivs, ";����ת��;") > 0 And mstr�����ʻ� <> "" Then
        If cbo.FindIndex(cboStyle, mstr�����ʻ�, True) = -1 Then
            cboStyle.AddItem mstr�����ʻ�
            cboStyle.ItemData(cboStyle.NewIndex) = 3
        End If
        
        'ҽ���ӿ�
        mcur�ʻ���� = gclsInsure.SelfBalance(mpatiInfo.����ID, mpatiInfo.ҽ����, 30, , mpatiInfo.����)
        lbl�ʻ����.Caption = lbl�ʻ����.Tag & Format(mcur�ʻ����, "0.00")
        lbl�ʻ����.Visible = True
        lblԤ�����.Left = lblδ�ɷ���.Left
        If lbl�ʻ����.Visible Then
            Line2(14).Visible = True: Line2(11).x2 = Line2(7).x2
        Else
            Line2(14).Visible = False: Line2(11).x2 = Line2(14).x2
        End If
    End If
    
    lbl�ѱ�ȼ�.Caption = lbl�ѱ�ȼ�.Tag & mpatiInfo.�ѱ�
    Call Get������Ϣ(mpatiInfo.����ID, mpatiInfo.��ҳID, dbl������, str������)
    lbl������.Caption = lbl������.Tag & str������
    lbl�������.Caption = lbl�������.Tag & Format(dbl������, "##,##0.00;-##,##0.00; ;")
    
    lbl�ֻ���.Caption = lbl�ֻ���.Tag & mpatiInfo.�ֻ���
    lbl���֤��.Caption = lbl���֤��.Tag & mpatiInfo.���֤��
    
    lblMemo.Caption = lblMemo.Tag & mpatiInfo.���˱�ע
    lblWorkUnit.Caption = lblWorkUnit.Tag & mpatiInfo.������λ
    lbl��ͥ��ַ.Caption = lbl��ͥ��ַ.Tag & mpatiInfo.��ͥ��ַ
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
    
    mlng����ID = 0
    mFactProperty = zl_GetInvoicePreperty(mlngFactModule, 2, cboType.ItemData(cboType.ListIndex))
    Call GetFact
    If cboType.Text = "סԺԤ��" Then
        If cboPatiPage.ListCount > 0 Then cboPatiPage.Tag = "0": cboPatiPage.ListIndex = 0
    End If
    Call ShowPremayBalance(True, 0)
    '���¼��ص�ǰ����˿���Ϣ
    If Not mblnNotClick Then
        Call ShowHistoryPrepay
        Call LoadThirdDelDeposit
    End If
    Call SetCtrlEnabled
    
    lblPatiPage.Visible = cboType.Text = "סԺԤ��": cboPatiPage.Visible = cboType.Text = "סԺԤ��"
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
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXml As String
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
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
     
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
    
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text = "" Then Exit Sub
    Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As Card)
    Call txtPatient_GotFocus
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub
Private Sub IDKind_ReadCard(ByVal objCard As Card, objPatiInfor As clsPatientInfo, blnCancel As Boolean)
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text = "" Then Exit Sub
    Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("���֤", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub
Private Sub SetcmdOkEnabled()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����cmdOk��neable����
    '���ƣ����˺�
    '���ڣ�2010-07-09 16:24:53
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    cmdOK.Enabled = mpatiInfo.����ID > 0
End Sub


Private Sub SetCtrlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Enabled����
    '����:���˺�
    '����:2011-07-24 09:30:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean, objCtl As Control
    Dim int���� As Integer
    
    If cboStyle.ListIndex >= 0 Then int���� = cboStyle.ItemData(cboStyle.ListIndex)
    blnEdit = True
    cboType.Enabled = blnEdit
    cboUnit.Enabled = blnEdit
    txtUnit.Enabled = blnEdit And int���� = 2
    cboStyle.Enabled = blnEdit
    txtCode.Enabled = blnEdit And int���� = 2
    txt������.Enabled = blnEdit And int���� = 2
    txt�ʺ�.Enabled = blnEdit And int���� = 2
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
    '��ѡ��֧Ʊʱ�Ŵ����ϴνɿ���Ϣ
    Dim strInfo As String
    Dim lngIndex As Long
    
    If cboStyle.ListIndex = -1 Then Exit Sub
        
    '�����:111657,����,2017/07/25,ʹ���ֽ�֧��Ԥ����ʱ,�λ������������
    mstrBrushCardNo = ""     '�����������ʱ����Ŀ���
    mcurBill.blnת�� = False
    mcurBill.lngԤ��ID = 0
    lngIndex = cboStyle.ListIndex + 1
    Call SetCtrlEnabled
    txtMoney.Enabled = True
    
    Select Case cboStyle.ItemData(cboStyle.ListIndex)
    Case 3, 1
        txtUnit.Text = "": txt������.Text = "": txt�ʺ�.Text = ""
    Case 2
        If cboStyle.Text Like "*Ʊ*" Or cboStyle.Text Like "*��*" Then
            If mpatiInfo.����ID = 0 Then Exit Sub
            strInfo = GetLastInfo(mpatiInfo.����ID)
            If strInfo <> "" Then
                txtUnit.Text = IIf(Split(strInfo, "|")(0) = "", txtUnit.Text, Split(strInfo, "|")(0))
                txt������.Text = IIf(Split(strInfo, "|")(1) = "", txt������.Text, Split(strInfo, "|")(1))
                txt�ʺ�.Text = IIf(Split(strInfo, "|")(2) = "", txt�ʺ�.Text, Split(strInfo, "|")(2))
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
    
    If InStr(1, mstrPrivs, ";Ԥ���˿�;") = 0 Then
        MsgBox "��û��Ȩ�޽���Ԥ���˿������", vbInformation, gstrSysName
        If cbo.Locate(cboStyle, BalanceType.C5���տ�, True) Then Cancel = True
    End If
End Sub


Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If SendMessage(cboUnit.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cboUnit.hwnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cboUnit.ListIndex = lngIdx
    'ǿ��Ҫѡ��һ��(��һ��)
    If cboUnit.ListIndex = -1 And cboUnit.ListCount <> 0 Then cboUnit.ListIndex = 0
End Sub
Private Sub cmdCancel_Click()
    If Not mblnOK Then Unload Me: Exit Sub
    If mpatiInfo.����ID > 0 Then
        If MsgBox("�ò��˵���δ�����˿����,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If MsgBox("ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Function CheckDataValied(ByRef objSetFocus As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���������Ƿ�Ϸ�
    '���أ��Ϸ�����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-18 16:38:39
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng����id As Long
    Dim intԤ����� As Integer, dblCashTotal As Double, dblCash As Double, dblPt As Double
    Dim objItem As clsBalanceItem, i As Long

    On Error GoTo errHandle
        
    If mpatiInfo.����ID = 0 Then
        lng����id = Val(txtPatient.Tag)
    Else
        lng����id = mpatiInfo.����ID
    End If
    
    '�˿����
    If InStr(1, mstrPrivs, ";Ԥ���˿�;") = 0 Then
        MsgBox "��û��Ȩ�޽���Ԥ���˿������", vbInformation, gstrSysName: Exit Function
    End If
    
    If mpatiInfo.����ID = 0 Then
        MsgBox "û��ȷ����Ԥ����Ĳ���,�����˿", vbExclamation, gstrSysName
        Set objSetFocus = txtPatient
        Exit Function
    End If
     
    If LenB(StrConv(txtUnit.Text, vbFromUnicode)) > 50 Then
        MsgBox "�ɿλ����ֻ���� 50 ���ַ��� 25 ������,���޸ģ�", vbInformation, App.Title
        Set objSetFocus = txtUnit
        Exit Function
    End If
    
    If LenB(StrConv(txt������.Text, vbFromUnicode)) > 50 Then
        MsgBox "����������ֻ���� 50 ���ַ��� 25 ������,���޸ģ�", vbInformation, App.Title
        Set objSetFocus = txt������
        Exit Function
    End If
    
    If LenB(StrConv(cboNote.Text, vbFromUnicode)) > 50 Then
        MsgBox "�ɿ�ժҪֻ���� 50 ���ַ��� 25 ������,���޸ģ�", vbInformation, App.Title
        Set objSetFocus = cboNote
        Exit Function
    End If
     
    If CCur(StrToNum(txtCashTotal.Text)) = 0 And CCur(StrToNum(txtThirdTotal.Text)) = 0 Then
        MsgBox "�˿����Ϊ�ջ���,�����룡", vbExclamation, gstrSysName
        Set objSetFocus = txtCashTotal
        Exit Function
    End If
    
    If StrToNum(txtCashTotal.Text) < 0 Then
        Call MsgBox("�ò�����������,���ܽ�������˿������", vbInformation + vbOKOnly, gstrSysName)
        Set objSetFocus = txtCashTotal
        Exit Function
    End If
    
    If Val(lblCashTotal.Tag) <> StrToNum(txtCashTotal.Text) Then
         If MsgBox("�㵱ǰ������˿������˿��б��е����ֽ�һ��,�Ƿ��Զ��������ֽ�" & vbCrLf & vbCrLf & _
               "���ֺϼ�:" & Format(Val(lblCashTotal.Tag), "###0.00###") & vbCrLf & _
               "������:" & Format(StrToNum(txtCashTotal.Text), "###0.00###") & vbCrLf & _
               "", vbQuestion + vbYes + vbDefaultButton1, gstrSysName) = vbNo Then Set objSetFocus = txtCashTotal: Exit Function
        '�Զ���̯
        Call AutoShareBalanceMoney(StrToNum(txtCashTotal.Text))
        If Val(lblCashTotal.Tag) <> StrToNum(txtCashTotal.Text) Then
            MsgBox "δ��̯��ɣ�����!", vbInformation + vbOKOnly, gstrSysName
            Set objSetFocus = txtCashTotal: Exit Function
        End If
    Else
        '���ܴ����������ֽ��¼���С����ɵģ���������ʾ�����¼���
        With vsBlance
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("�˿ʽ")) <> "" Then
                    If zlGetBalanceItemFromBalanceGrid(i, objItem) = False Then Exit Function
                    If GetVsGridBoolColVal(vsBlance, i, .ColIndex("����")) Then
                          dblCashTotal = roundEx(dblCashTotal + objItem.ʣ����, 5)
                          dblCash = roundEx(dblCash + objItem.������, 5)
                    End If
                End If
            Next
        End With
        
        If dblCash < StrToNum(txtCashTotal.Text) Then
           If MsgBox("��������˿�����������ֺϼƣ��Ƿ����¼����˿" & vbCrLf & "������:" & txtCashTotal.Text & vbCrLf & "���ֺϼ�:" & Format(dblPt + dblCash, "###0.00"), vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call ReCalePtBalanceMoney(1) 'ֻ�������ֲ���
           End If
           Exit Function
        End If
    End If

    If Val(lblCashTotal.Tag) <> StrToNum(txtMoney.Text) - StrToNum(txt�տ�.Text) Then
        Call MsgBox("��ǰ�˿���(" & Format(CCur(StrToNum(txtMoney.Text) - StrToNum(txt�տ�.Text)), "0.00") & ")�뱾�����ֺϼ�(" & Format(Val(lblCashTotal.Tag), "0.00") & ")��һ��,�����˿�!", vbInformation + vbOKOnly, gstrSysName)
        Set objSetFocus = txtMoney
        Exit Function
    End If
    
    If mdblʣ���� - CCur(StrToNum(txtCashTotal.Text)) - CCur(StrToNum(txtThirdTotal.Text)) < 0 Then
        If mbytBackMoneyType = 1 Then
            Call MsgBox("�˿���(" & Format(CCur(StrToNum(txtCashTotal.Text)) + CCur(StrToNum(txtThirdTotal.Text)), "0.00") & ")�����˲��˵�ǰ��ʣ���(" & Format(mdblʣ����, "0.00") & "),�����˿�!", vbInformation + vbOKOnly, gstrSysName)
            Set objSetFocus = txtCashTotal
            Exit Function
        Else
            If MsgBox("�˿���(" & Format(CCur(StrToNum(txtCashTotal.Text)) + CCur(StrToNum(txtThirdTotal.Text)), "0.00") & ")�����˲��˵�ǰ��ʣ���(" & Format(mdblʣ����, "0.00") & "),������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Set objSetFocus = txtCashTotal
                Exit Function
            End If
            mbytOracleBackType = 0
        End If
    End If

    If cboStyle.ListIndex = -1 And CCur(StrToNum(txtMoney.Text)) <> 0 Then
        MsgBox "δȷ����ǰ�˿ʽ�����ܽ�������˿", vbExclamation, gstrSysName
        Set objSetFocus = cboType
        Exit Function
    End If
    
    If cboStyle.ListIndex >= 0 Then
        If mobjThridSwap.objPayCards(cboStyle.ItemData(cboStyle.ListIndex)).�������� = 3 Then
            MsgBox "ҽ�����˸����ʻ�ת�ʽ��ܽ�������˿������", vbInformation, gstrSysName
            Set objSetFocus = txtMoney
            Exit Function
        End If
    End If
    
    If cboType.ListIndex >= 0 Then intԤ����� = cboType.ItemData(cboType.ListIndex)
    Select Case intԤ�����
    Case 1 '����Ԥ��
        If InStr(1, mstrPrivs, ";���ﲡ������˿�;") = 0 Then
           MsgBox "��û��Ȩ�޽�������˿����,�������Ա��ϵ��������˿�Ȩ�ޣ�", vbInformation, gstrSysName: Exit Function
        End If
                
        If gbytԤ��������鿨 <> 0 Then
            If CreatePublicExpense() Then
                If Not gobjPublicExpense.zlPatiIdentify(mlngModul, Me, lng����id, Val(StrToNum(txtCashTotal.Text)), True) Then
                    Set objSetFocus = cboType
                    Exit Function
                End If
            End If
        End If

    Case 2 'סԺԤ��
    
        If mbln������Ժ��������˿� = False And mpatiInfo.��Ժ Then
            MsgBox "������Ժ,���ܽ�������˿�,���飡", vbInformation, gstrSysName
            Set objSetFocus = txtMoney
            Exit Function
        End If
    
         If Not mblnNurseCall And InStr(1, mstrPrivs, ";��Ժ��������˿�;") = 0 And mpatiInfo.��Ժ Then
            MsgBox "��û��Ȩ�޶���Ժ���˽�������˿����,�������Ա��ϵ��������˿�Ȩ�ޣ�", vbInformation, gstrSysName: Exit Function
         End If
         
         If Not mblnNurseCall And InStr(1, mstrPrivs, ";��Ժ��������˿�;") = 0 And Not mpatiInfo.��Ժ Then
            MsgBox "��û��Ȩ�޶Գ�Ժ���˽�������˿����,�������Ա��ϵ��������˿�Ȩ�ޣ�", vbInformation, gstrSysName: Exit Function
         End If
         
        If gbytԤ��������鿨 <> 0 And mblnסԺ��Ԥ����֤ Then
            If CreatePublicExpense() Then
                If Not gobjPublicExpense.zlPatiIdentify(mlngModul, Me, lng����id, Val(StrToNum(txtCashTotal.Text)), True) Then
                    Set objSetFocus = cboType
                    Exit Function
                End If
            End If
        End If
        
    Case Else
        MsgBox "δѡ������˿��Ԥ������,������", vbExclamation, gstrSysName
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
    '����:��鷢Ʊ�Ƿ���Ч(ͬʱ���ɷ�Ʊ��)
    '����:��Ʊ�Ϸ���true,���򷵻�False
    '����:���˺�
    '����:2018-08-31 09:32:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
     
    Dim intԤ������ As Integer
    
    On Error GoTo errHandle
    If cboType.ListIndex <> -1 Then intԤ������ = cboType.ItemData(cboType.ListIndex)
    If mobjEInvoice.zlIsStartEInvoice(0, intԤ������) Then
        If mobjEInvoice.zlGetTranPaperInvoiceModule = 0 Then CheckFactIsValied = True: Exit Function
        If Trim(txtFact.Text) = "" Then
            MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
            zlControl.ControlSetFocus txtFact: Exit Function
        End If
        CheckFactIsValied = True: Exit Function
    End If
  
    If mFactProperty.intInvoicePrint = 0 Then CheckFactIsValied = True: Exit Function
    If Trim(txtFact.Text) = "" Then Call GetFact
    
    'Ʊ�ݺ�����
    If gblnBillԤ�� Then
        If Trim(txtFact.Text) = "" Then
            MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
            zlControl.ControlSetFocus txtFact: Exit Function
        End If
        
        mlng����ID = CheckUsedBill(2, IIf(mlng����ID > 0, mlng����ID, mFactProperty.lngShareUseID), txtFact.Text, cboType.ItemData(cboType.ListIndex))
        If mlng����ID <= 0 Then
            Select Case mlng����ID
                Case 0 '����ʧ��
                Case -1
                    MsgBox "��û�����ú͹��õ�Ԥ��Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Case -3
                    MsgBox "Ʊ�ݺ��벻�ڵ�ǰ��Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
                    txtFact.SetFocus
            End Select
            txtFact.Text = ""
            Exit Function
        End If
        CheckFactIsValied = True: Exit Function
    End If
    
    If Len(txtFact.Text) <> gbytԤ�� And txtFact.Text <> "" Then
        MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytԤ�� & " λ��", vbInformation, gstrSysName
        txtFact.SetFocus: Exit Function
    End If
    CheckFactIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetItemsFromRecord(ByVal intԤ������ As Integer, ByVal dblMoney As Double, ByVal rsMoney As ADODB.Recordset, ByRef objItems_out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݼ�¼������ȡ���еĽ�����
    '���:dblMoney-��ǰ��̯���
    '     rsMoney-��ǰ��¼��
    '����:objItems_out-���ط�̯����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-09-11 11:38:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, bln���ѿ� As Boolean
    Dim dblTemp As Double, objItem As clsBalanceItem
    Dim objCard  As Card
    
    On Error GoTo errHandle
    If dblMoney = 0 Then GetItemsFromRecord = True: Exit Function
    
    If objItems_out Is Nothing Then Set objItems_out = New clsBalanceItems
    If rsMoney.RecordCount = 0 Then GetItemsFromRecord = True: Exit Function
    
    rsMoney.MoveFirst
    rsMoney.Sort = "�տ�ʱ��"
    With rsMoney
        Do While Not .EOF
            lngCardTypeID = Val(Nvl(!�����ID))
            bln���ѿ� = Val(Nvl(!���ѿ�)) = 1
            dblTemp = roundEx(Val(Nvl(!��Ԥ��)), 6)
            If dblTemp <> 0 Then
                
                If dblMoney > dblTemp Then
                    dblMoney = roundEx(dblMoney - dblTemp, 6)
                Else
                    dblTemp = dblMoney: dblMoney = 0
                End If
                rsMoney!��Ԥ�� = roundEx(Nvl(!��Ԥ��, 0) - dblTemp, 6)
                rsMoney.Update
                Set objCard = mobjThridSwap.zlGetCardFromCardType(lngCardTypeID, bln���ѿ�, Nvl(!���㷽ʽ))
                Set objItem = New clsBalanceItem
                Set objItem.objCard = objCard
                objItem.���㷽ʽ = Nvl(!���㷽ʽ)
                If objItem.���㷽ʽ = "" Then objItem.���㷽ʽ = objCard.���㷽ʽ
                objItem.���� = Nvl(!����)
                objItem.�����ID = lngCardTypeID
                objItem.�Ƿ�����ɾ�� = True
                objItem.Ԥ��ID = Val(Nvl(!Ԥ��ID))
                objItem.���ѿ� = bln���ѿ�
                objItem.У�Ա�־ = 1
                objItem.�Ƿ��˿�ֽ��� = True
                objItem.�Ƿ�Ԥ�� = True
                objItem.�Ƿ����� = Val(objCard.�������Ĺ���) <> 0
                objItem.�������� = objCard.��������
                objItem.������ = roundEx(dblTemp, 2)
                objItem.ʣ���� = roundEx(Val(Nvl(!Ԥ�����)), 2)
                objItem.ԭʼ��� = roundEx(Val(Nvl(!���)), 2)
                objItem.��������ID = Val(Nvl(!��������ID))
                objItem.������ˮ�� = Trim(Nvl(!������ˮ��))
                objItem.����˵�� = Trim(Nvl(!����˵��))
                objItem.������� = Trim(Nvl(!�������))
                objItem.����ժҪ = Trim(Nvl(!ժҪ))
                objItem.����Ԥ�� = intԤ������ = 1
                 
                objItems_out.AddItem objItem
                objItems_out.������ = roundEx(objItems_out.������ + objItem.������, 6)
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

Private Function LoadThirdDelDeposit(Optional int��ҳID As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������˿���Ϣ
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-21 18:01:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItemsTemp As clsBalanceItems, objItem As clsBalanceItem
    Dim objFsItems As clsBalanceItems '����Ԥ����
    Dim rsMoney As ADODB.Recordset, strSQL As String
    Dim intԤ������ As Integer, objCard As Card
    Dim strWhere As String, lng����id As Long, strDefaultBalance As String, strTemp As String
    Dim lngCardTypeID As Long, bln���ѿ� As Boolean, blnDelCash As Boolean
    Dim lngRow As Long, dblThirdMoney As Double, dblCashMoney As Double
    Dim blnAdd As Boolean, dblMoney As Double
    Dim intKind As Integer
    Dim i As Integer
     
    On Error GoTo errHandle
    
    If mpatiInfo.����ID = 0 Then Exit Function
    
    lng����id = mpatiInfo.����ID
    
    strWhere = " And Not Exists(Select 1 From ���㷽ʽ   Where B.���㷽ʽ= ���� And ����=5)"    '�������տ�
    If int��ҳID > 0 Then strWhere = strWhere & " And b.��ҳID=[3] "
    
    intԤ������ = cboType.ItemData(cboType.ListIndex)
    strSQL = "" & _
    "    Select b.no,a.Ԥ��id, a.����id, a.Ԥ�����, nvl(a.Ԥ�����,0) as Ԥ�����,b.���,a.Ԥ����� as ��Ԥ��,b.���㷽ʽ, Nvl(b.�����id, b.���㿨���) As �����id, " & vbCrLf & _
    "           Decode(Nvl(b.���㿨���, 0), 0, 0, 1) As ���ѿ�, b.����, " & vbCrLf & _
    "           b.������ˮ��, b.����˵��, b.��������id, b.�տ�ʱ��,b.�������,b.ժҪ,nvl(c.�Ƿ�����,0) as ���ѿ��Ƿ�����,M.���� " & vbCrLf & _
    "    From Ԥ��������� A, ����Ԥ����¼ B,���ѿ����Ŀ¼ C,���㷽ʽ M " & vbCrLf & _
    "    Where a.����id = [1] And a.Ԥ����� = [2] And a.Ԥ��id = b.Id and B.���㿨���=C.���(+) and nvl(a.Ԥ�����,0)<>0 And B.���㷽ʽ=M.����(+) " & strWhere & vbCrLf & _
    "    Order By b.�տ�ʱ��"
    
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����id, intԤ������, int��ҳID)
    
    Set rsMoney = zlDatabase.CopyNewRec(rsMoney)
    
    Set objFsItems = New clsBalanceItems
    
    '�ȴ�����Ԥ������
    rsMoney.Filter = "Ԥ�����<0"
    rsMoney.Sort = "�տ�ʱ��"
    With rsMoney
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            lngCardTypeID = Val(Nvl(!�����ID))
            bln���ѿ� = Val(Nvl(!���ѿ�)) = 1
            dblMoney = roundEx(Val(Nvl(!Ԥ�����)), 6)
            
            Set objCard = mobjThridSwap.zlGetCardFromCardType(lngCardTypeID, bln���ѿ�, Nvl(!���㷽ʽ))
            Set objItem = New clsBalanceItem
            Set objItem.objCard = objCard
            objItem.���㷽ʽ = Nvl(!���㷽ʽ)
            If objItem.���㷽ʽ = "" Then objItem.���㷽ʽ = objCard.���㷽ʽ
            objItem.���ݺ� = Nvl(!NO)
            objItem.���� = Nvl(!����)
            objItem.�����ID = lngCardTypeID
            objItem.�Ƿ�����ɾ�� = True
            objItem.Ԥ��ID = Val(Nvl(!Ԥ��ID))
            objItem.���ѿ� = bln���ѿ�
            objItem.У�Ա�־ = 1
            objItem.�Ƿ��˿�ֽ��� = True
            objItem.�Ƿ�Ԥ�� = True
            objItem.�Ƿ����� = Val(objCard.�������Ĺ���) <> 0
            objItem.�������� = objCard.��������
            objItem.������ = roundEx(Val(Nvl(!Ԥ�����)), 2)
            objItem.ʣ���� = objItem.������
            objItem.ԭʼ��� = roundEx(Val(Nvl(!���)), 2)
            objItem.��������ID = Val(Nvl(!��������ID))
            objItem.������ˮ�� = Trim(Nvl(!������ˮ��))
            objItem.����˵�� = Trim(Nvl(!����˵��))
            objItem.������� = Trim(Nvl(!�������))
            objItem.����ժҪ = Trim(Nvl(!ժҪ))
            objItem.����Ԥ�� = intԤ������ = 1
            objItem.�Ƿ�ת�� = objCard.�Ƿ�ת�ʼ�����
            objFsItems.AddItem objItem
            objFsItems.������ = roundEx(objFsItems.������ + objItem.������, 6)
            .MoveNext
        Loop
    End With
    
    '��δ�����ڹ�������ID����Ԥ��
    For Each objItem In objFsItems
         If objItem.�����ID > 0 And objItem.���ѿ� = False Then
            '������
            rsMoney.Filter = "�����ID=" & objItem.�����ID & " And ���ѿ�=0 And ��������ID=" & objItem.��������ID & " And ��Ԥ��>0"
            If objItem.objTag Is Nothing Then Set objItem.objTag = New clsBalanceItems
            Set objItemsTemp = objItem.objTag
            dblMoney = roundEx(-1 * objItem.������ - objItemsTemp.������, 6)
            If dblMoney >= 0 Then
                Call GetItemsFromRecord(intԤ������, dblMoney, rsMoney, objItemsTemp)
                Set objItem.objTag = objItemsTemp
            End If
         End If
    Next
    
    '�ٴ����������ID���ڣ��������ڶ�Ӧ�ĵļ�¼
    For Each objItem In objFsItems
         If objItem.�����ID > 0 And objItem.���ѿ� = False Then
            '������
            rsMoney.Filter = "�����ID=" & objItem.�����ID & "  And ���ѿ�=0 And ��������ID=0 And ��Ԥ��>0"
            If objItem.objTag Is Nothing Then Set objItem.objTag = New clsBalanceItems
            Set objItemsTemp = objItem.objTag
            dblMoney = roundEx(-1 * objItem.������ - objItemsTemp.������, 6)
            
            If dblMoney >= 0 Then
                Call GetItemsFromRecord(intԤ������, dblMoney, rsMoney, objItemsTemp)
                Set objItem.objTag = objItemsTemp
            End If
         End If
    Next
         
    '�������ͨ�ķ�̯����
    For Each objItem In objFsItems
        '������
        rsMoney.Filter = "��Ԥ��>0"
        If objItem.objTag Is Nothing Then Set objItem.objTag = New clsBalanceItems
        Set objItemsTemp = objItem.objTag
        dblMoney = roundEx(-1 * objItem.������ - objItemsTemp.������, 6)
        If dblMoney >= 0 Then
            Call GetItemsFromRecord(intԤ������, dblMoney, rsMoney, objItemsTemp)
            Set objItem.objTag = objItemsTemp
        End If
    Next
    
    Call SaveAutoRelevanceData(lng����id, objFsItems)
    
    Call ClearVsBalance
    lngRow = 1
    dblCashMoney = 0: dblThirdMoney = 0: strDefaultBalance = ""
    vsBlance.Redraw = flexRDNone
    
    rsMoney.Filter = "��Ԥ��>0"
    rsMoney.Sort = "�տ�ʱ��"
    With rsMoney
        Do While Not .EOF
            lngCardTypeID = Val(Nvl(!�����ID))
            bln���ѿ� = Val(Nvl(!���ѿ�)) = 1
            dblMoney = roundEx(Val(Nvl(!��Ԥ��)), 6)
    
            Set objCard = mobjThridSwap.zlGetCardFromCardType(lngCardTypeID, bln���ѿ�, Nvl(!���㷽ʽ))
            Set objItem = New clsBalanceItem
            Set objItem.objCard = objCard
            objItem.���㷽ʽ = Nvl(!���㷽ʽ)
            If objItem.���㷽ʽ = "" Then objItem.���㷽ʽ = objCard.���㷽ʽ
            objItem.���ݺ� = Nvl(!NO)
            objItem.���� = Nvl(!����)
            objItem.�����ID = lngCardTypeID
            objItem.�Ƿ�����ɾ�� = True
            objItem.Ԥ��ID = Val(Nvl(!Ԥ��ID))
            objItem.���ѿ� = bln���ѿ�
            objItem.У�Ա�־ = 1
            objItem.�Ƿ��˿�ֽ��� = True
            objItem.�Ƿ�Ԥ�� = True
            objItem.�Ƿ����� = Val(objCard.�������Ĺ���) <> 0
            objItem.�������� = objCard.��������
            objItem.������ = dblMoney
            objItem.δ�˽�� = dblMoney
            objItem.ʣ���� = roundEx(Val(Nvl(!Ԥ�����)), 2)
            objItem.ԭʼ��� = roundEx(Val(Nvl(!���)), 2)
            objItem.��������ID = Val(Nvl(!��������ID))
            objItem.������ˮ�� = Trim(Nvl(!������ˮ��))
            objItem.����˵�� = Trim(Nvl(!����˵��))
            objItem.������� = Trim(Nvl(!�������))
            objItem.����ժҪ = Trim(Nvl(!ժҪ))
            objItem.����Ԥ�� = intԤ������ = 1
            objItem.�Ƿ�ת�� = objCard.�Ƿ�ת�ʼ�����
            If objCard Is Nothing Then
                objItem.�������� = Val(Nvl(!����))
            Else
                objItem.�������� = objCard.��������
            End If
        
            Set objItemsTemp = New clsBalanceItems
            objItemsTemp.AddItem objItem
            objItemsTemp.������ = objItem.������
            objItemsTemp.�շ����� = 1
            blnAdd = False
            
            If bln���ѿ� Then
                objItem.�Ƿ��������� = Val(Nvl(!���ѿ��Ƿ�����)) = 1
                objItem.�Ƿ�ǿ������ = True
                objItem.�Ƿ�����ɾ�� = True
                objItem.Tag = IIf(objItem.�Ƿ���������, "ȱʡ����", "")
                blnDelCash = IIf(objItem.�Ƿ���������, True, False)
            ElseIf lngCardTypeID > 0 Then
                If Not mobjThridSwap.zlThirdReturnCashCheck(objCard, objItemsTemp, blnDelCash, strTemp) Then
                    '1.��ֹ����
                    objItem.�Ƿ��������� = False
                    objItem.�Ƿ�ǿ������ = blnDelCash
                    objItem.�Ƿ�����ɾ�� = True
                    blnDelCash = False
                    blnAdd = True
                Else
                    objItem.�Ƿ�����༭ = False
                    objItem.�Ƿ�����ɾ�� = True
                    objItem.�Ƿ�ǿ������ = True
                    objItem.�Ƿ��������� = True
                    
                    If blnDelCash = False Then  '�Ƿ�ȱʡ����
                        '�������֣�����ɾ��
                        objItem.Tag = ""
                    Else
                        objItem.Tag = "ȱʡ����"
                        If strTemp <> "" And strDefaultBalance = "" Then strDefaultBalance = strTemp
                    End If
                End If
            End If
            
            If lngCardTypeID <> 0 Then
                objItem.�������� = IIf(bln���ѿ�, 5, 3)
            ElseIf objCard.�������� = 7 Then
                objItem.�������� = 4 '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                blnDelCash = IIf(objItem.�Ƿ���������, True, False)
            Else
                objItem.�������� = 0
                objItem.�Ƿ�����ɾ�� = True
                objItem.�Ƿ�ǿ������ = True
                objItem.�Ƿ��������� = True
                blnDelCash = True
            End If
            
            With vsBlance
                .TextMatrix(lngRow, .ColIndex("����")) = IIf(objItem.�Ƿ��������� And blnDelCash, 1, 0)
                .TextMatrix(lngRow, .ColIndex("����")) = objItem.��������
                .TextMatrix(lngRow, .ColIndex("�����ID")) = objItem.�����ID
                .TextMatrix(lngRow, .ColIndex("���ѿ�ID")) = objItem.���ѿ�ID
                .TextMatrix(lngRow, .ColIndex("��������")) = objItem.��������
                .TextMatrix(lngRow, .ColIndex("�༭״̬")) = IIf(objItem.�Ƿ�����༭, "1", "0") & "|" & IIf(objItem.�Ƿ�����ɾ��, "1", "0")      '�Ƿ�����༭|�Ƿ�����ɾ��
                .TextMatrix(lngRow, .ColIndex("�Ƿ�����")) = IIf(objCard.�Ƿ�����, 1, 0)
                .TextMatrix(lngRow, .ColIndex("�Ƿ�ȫ��")) = IIf(objCard.�Ƿ�ȫ��, 1, 0)
                .TextMatrix(lngRow, .ColIndex("У�Ա�־")) = objItem.У�Ա�־
                .TextMatrix(lngRow, .ColIndex("�Ƿ�����")) = IIf(objItem.�Ƿ�����, 1, 0)
                .TextMatrix(lngRow, .ColIndex("���������")) = objCard.����
                .TextMatrix(lngRow, .ColIndex("���ݺ�")) = objItem.���ݺ�
                .TextMatrix(lngRow, .ColIndex("�˿ʽ")) = objItem.���㷽ʽ
                .TextMatrix(lngRow, .ColIndex("Ԥ�����")) = IIf(objItem.�������� = 9, Format(objItem.ʣ����, "###0.00#####"), Format(objItem.ʣ����, "0.00"))
                .TextMatrix(lngRow, .ColIndex("�˿���")) = IIf(objItem.�������� = 9, Format(objItem.������, "###0.00#####"), Format(objItem.������, "0.00"))
                .TextMatrix(lngRow, .ColIndex("�������")) = objItem.�������
                .TextMatrix(lngRow, .ColIndex("��ע")) = objItem.����ժҪ
                .TextMatrix(lngRow, .ColIndex("������ˮ��")) = objItem.������ˮ��
                .TextMatrix(lngRow, .ColIndex("����˵��")) = objItem.����˵��
                .TextMatrix(lngRow, .ColIndex("����")) = IIf(objItem.�Ƿ�����, String(Len(objItem.����), "*"), objItem.����)
                .RowData(lngRow) = objItem
                .Rows = .Rows + 1
                lngRow = .Rows - 1
            End With
            dblThirdMoney = roundEx(dblThirdMoney + IIf(objItem.�Ƿ��������� And blnDelCash, 0, objItem.������), 2)
            dblCashMoney = roundEx(dblCashMoney + IIf(objItem.�Ƿ��������� And blnDelCash, objItem.������, 0), 2)
            .MoveNext
        Loop

    End With
    With vsBlance
        If .Rows > 2 Then
            If .TextMatrix(.Rows - 1, .ColIndex("�˿ʽ")) = "" Then
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
    If mdblԤ����� <> mdblʣ���� Then Call AutoShareBalanceMoney(mdblʣ����, True)
    'ȱʡ��λ����ǰȱʡ�Ľ��㷽ʽ�ϣ���һ��)
    For i = 0 To cboStyle.ListCount - 1
        intKind = cboStyle.ItemData(i)
        If mobjThridSwap.objPayCards(intKind).���㷽ʽ = strDefaultBalance Then
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

Private Sub SaveAutoRelevanceData(ByVal lng����id As Long, ByVal objItems As clsBalanceItems)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���渺��Ԥ���Զ�����������Ŀ
    '���:objItems-��Ŀ��Ϣ
     '����:���˺�
    '����:2018-09-11 17:40:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strHead As String, strTemp As String, str������Ϣ As String
    Dim objItem As clsBalanceItem, objItemsTemp As clsBalanceItems, objItemTemp As clsBalanceItem
    Dim cllPro As Collection, blnTrans As Boolean
    Dim strDate As String, lng����ID As Long
    
    On Error GoTo errHandle
    
    If objItems Is Nothing Then Exit Sub
    If objItems.Count = 0 Then Exit Sub
    Set cllPro = New Collection
    strDate = "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')"
    
    For Each objItem In objItems
        Set objItemsTemp = objItem.objTag
        If Not objItemsTemp Is Nothing Then
            '    Zl_����Ԥ����¼_Relevance
            strHead = "Zl_����Ԥ����¼_Relevance("
            '    ����id_In     ����Ԥ����¼.����id%Type,
            strHead = strHead & "" & lng����id & ","
            '    Ԥ��id_In     ����Ԥ����¼.Id%Type,
            strHead = strHead & "" & objItem.Ԥ��ID & ","
            str������Ϣ = ""
            For Each objItemTemp In objItemsTemp
                'ԭԤ��ID|���||....
                strTemp = "||" & objItemTemp.Ԥ��ID & "|" & objItemTemp.������
                If zlCommFun.ActualLen(str������Ϣ & strTemp) > 4000 Then
                    lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
                    str������Ϣ = Mid(str������Ϣ, 3)
                    strSQL = strHead
                    '    ������Ϣ_In   Varchar2 := Null,
                    strSQL = strSQL & "'" & str������Ϣ & "',"
                    '   ����id_In     ����Ԥ����¼.����id%Type,
                    strSQL = strSQL & "" & lng����ID & ","
                    '    ����Ա���_In ����Ԥ����¼.����Ա���%Type,
                    strSQL = strSQL & "'" & UserInfo.��� & "',"
                    '    ����Ա����_In ����Ԥ����¼.����Ա����%Type,
                    strSQL = strSQL & "'" & UserInfo.���� & "',"
                    '    �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type := Null,
                    strSQL = strSQL & "" & strDate & ","
                    '    У�Ա�־_In   ����Ԥ����¼.У�Ա�־%Type := 0,
                    strSQL = strSQL & "" & 0 & ")"
                    '    �ɿ���id_In   ����Ԥ����¼.�ɿ���id%Type := -1
                    zlAddArray cllPro, strSQL
                    str������Ϣ = ""
                End If
                str������Ϣ = str������Ϣ & strTemp
            Next
            If str������Ϣ <> "" Then
                lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
                str������Ϣ = Mid(str������Ϣ, 3)
                strSQL = strHead
                '    ������Ϣ_In   Varchar2 := Null,
                strSQL = strSQL & "'" & str������Ϣ & "',"
                '   ����id_In     ����Ԥ����¼.����id%Type,
                strSQL = strSQL & "" & lng����ID & ","
                '    ����Ա���_In ����Ԥ����¼.����Ա���%Type,
                strSQL = strSQL & "'" & UserInfo.��� & "',"
                '    ����Ա����_In ����Ԥ����¼.����Ա����%Type,
                strSQL = strSQL & "'" & UserInfo.���� & "',"
                '    �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type := Null,
                strSQL = strSQL & "" & strDate & ","
                '    У�Ա�־_In   ����Ԥ����¼.У�Ա�־%Type := 0,
                strSQL = strSQL & "" & 0 & ")"
                '    �ɿ���id_In   ����Ԥ����¼.�ɿ���id%Type := -1
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
    '����:���¼�����ͨ�˿���
    '���:intCalceType ��=0��ʾֻ�������ֺϼƼ������˿�ϼ�
    '                               =1��ʾ:ֻ�������ֲ��֣���ʣ�����Ϊ�����˿�
    '                               =2��ʾ:�����������㣬ʣ�����Ϊ�����˿�
    '����:���˺�
    '����:2018-09-07 09:42:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblCashMoney As Double, dblThirdDelMoney As Double
    Dim objItem As clsBalanceItem
    
    On Error GoTo errHandle

    With vsBlance
        dblCashMoney = 0: dblThirdDelMoney = 0
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("�˿ʽ")) <> "" Then
                If zlGetBalanceItemFromBalanceGrid(i, objItem) Then
                   If GetVsGridBoolColVal(vsBlance, i, .ColIndex("����")) Then
                        If intCalceType = 1 Or intCalceType = 2 Then
                            objItem.������ = objItem.ʣ����
                            .TextMatrix(i, .ColIndex("�˿���")) = Format(objItem.������, "###0.00" & IIf(objItem.�������� = 9, "###", ""))
                        End If
                        dblCashMoney = roundEx(dblCashMoney + objItem.������, 6)
                   Else
                        If intCalceType = 2 Then
                            objItem.������ = objItem.ʣ����
                            .TextMatrix(i, .ColIndex("�˿���")) = Format(objItem.������, "###0.00" & IIf(objItem.�������� = 9, "###", ""))
                        End If
                        dblThirdDelMoney = roundEx(dblThirdDelMoney + objItem.������, 6)
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
    '����:������㷽ʽ�б�(����)
    '����:���˺�
    '����:2018-08-30 13:36:45
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
    '����:������������ֹ����Ա�ظ�����
    '���:blnLocked-��ʾ����
    '����:���˺�
    '����:2018-08-31 09:59:07
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
 
    Dim objPati As clsPatiInfo, lngԤ��ID As Long
    Dim objSefocus As Object
    Dim objDelCashThird As clsBalanceItems
    Dim bytP As Byte, blnVocherPrint As Boolean
    Dim strNos As String, intԤ������ As Integer

    If cmdOK.Enabled = False Then Exit Sub '��ֹ�ظ�ִ��
    
     
    Call LockScreen(True)
    
    If Not Checkδ��Ʋ���Ԥ�� Then Call LockScreen(False):  Exit Sub
    
    If CheckDataValied(objSefocus) = False Then
        Call LockScreen(False):
        If Not objSefocus Is Nothing Then Call zlControl.ControlSetFocus(objSefocus)
        Exit Sub
    End If
    
    If Check�˿� = False Then
        Call LockScreen(False): Call zlControl.ControlSetFocus(txtMoney): Exit Sub
    End If
    
    If GetPatiObject(objPati) = False Then
        Call LockScreen(False): Call zlControl.ControlSetFocus(txtPatient): Exit Sub
    End If
    
    If cboType.ListIndex <> -1 Then intԤ������ = cboType.ItemData(cboType.ListIndex)
    '���ϵ���Ʊ��
    If zlCancelEInvoiceBat(mpatiInfo, strNos) = False Then
        MsgBox "���ϵ���Ʊ��ʧ�ܣ���ֹ����˿�", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    bytP = Val(zlDatabase.GetPara("ƾ����ӡ��ʽ", glngSys, mlngModul))
    Select Case bytP
    Case 0 '����ӡԤ����Ʊ
       blnVocherPrint = False
    Case 1 '�Զ���ӡ
       blnVocherPrint = True
    Case 2 '��ӡ����
        If MsgBox("�Ƿ���Ҫ��ӡԤ��ƾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then blnVocherPrint = True
    End Select
    '��ʱ�ֵ��ݽ��г�Ԥ��
    'str�տ�ʱ�� = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    'str����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    'str������� = "-" & str����ID
    
    '1.�ٱ�����������Ԥ�����ԭ����
    Set objDelCashThird = New clsBalanceItems
    If Excute_BalanceList_ReturnMoney(objPati, lngԤ��ID, objDelCashThird, blnVocherPrint) = False Then
        Call LockScreen(False): Call zlControl.ControlSetFocus(vsBlance): Exit Sub
    End If
    
    '2.�ȱ�����ͨ������˿�
    If Excute_CashAndOther_ReturnMoney(objPati, objDelCashThird, lngԤ��ID, blnVocherPrint) = False Then
        Call LockScreen(False): Call zlControl.ControlSetFocus(txtMoney): Exit Sub
    End If
    
    '���¿��ߵ���Ʊ��
    Call zlCreateEInvoiceBat(mpatiInfo, strNos)
    
    '3.��ɺ󣬰�������
    Call Excute_Plug_PatiPrePayAfter(objPati, lngԤ��ID)
    
    '4.����
    Call LockScreen(False)
    '����:48249
    If mbytCallObject <> 0 Then '����ģ�����Ԥ���ɿ�ʱ,ֱ���˳�
        mblnOK = True: txtPatient.Tag = "": Unload Me: Exit Sub
    End If
    
    If mblnClearWinInfor Then
        Call ClearBill
        Call InitFace(True)
        Call cboStyle_Click
    Else
        SetMoneyInfo False
        Set mpatiInfo = New clsPatientInfo
        Call GetFact  '���»�ȡ��Ʊ��
        txtPatient.Tag = ""
    End If
    
    Call SetcmdOkEnabled
    If txtPatient.Enabled Then txtPatient.SetFocus
    mblnOK = True
End Sub

Private Function DelDepositErrBill(ByVal strNO As String, Optional ByVal bytOpt As Byte) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��Ԥ���쳣���ݼ�¼
    '���: strno-���ݺţ�Optype-(0-ɾ���쳣��ֵ���ݣ�1-ɾ���쳣�˿�ݣ�2-ɾ���쳣����˿��)
    '����:
    '����:2018-06-29
    '˵��:
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
    '����:�����ؽ��������
    '����:���˺�
    '����:2018-11-29 10:03:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    If gblnLED Then zl9LedVoice.DisplayPatient ""
    
    Set mpatiInfo = New clsPatientInfo '���������Ϣ
    txtPatient.Text = "": txtPatient.Locked = False
    txtPatient.Tag = ""
    cboUnit.ListIndex = 0
    txtUnit.Tag = ""
    txtUnit.Text = ""
    mstr�˿����Ա = ""
    
    txt������.Text = ""
    txt�ʺ�.Text = ""
    SetMoneyInfo True
    
    txtMoney.Text = "0.00": txt�տ�.Text = "0.00"
    lblCashTotal.Tag = "": txtCashTotal.Text = "0.00"
    txtTotal.Text = "0.00": txtThirdTotal.Text = "0.00"
    
    If cboStyle.ListCount <> 0 And cboStyle.Tag <> "" Then cboStyle.ListIndex = Val(cboStyle.Tag) '�ָ�ȱʡ���㷽ʽ
    txtCode.Text = "": txtCode.Locked = False
    
    cboNote.Text = ""
    
    Call ClearVsBalance
    'ҽ���Ķ�
    Call Clear�����ʻ�
    
    '�µ�һ��Ԥ�����
    cboNO.Text = "": cboNO.Locked = True
    
    vsBlance.Rows = 1: vsBlance.Rows = 2
    vsDepositHistory.Rows = 1: vsDepositHistory.Rows = 2
    vsThirdTotal.Rows = 1: vsThirdTotal.Rows = 2
    
    txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ�ݺ�") And gblnBillԤ�� '89302
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
    
    If mlng����ID <> 0 And Trim(txtPatient.Text) = "" Then
        txtPatient.Text = "-" & mlng����ID
        Call txtPatient_KeyPress(13)
        If mdblDefPreMoney <> 0 And StrToNum(txtMoney.Text) = 0 Then
            txtMoney.Text = Format(mdblDefPreMoney, "###0.00;-###0.00;;")
        End If
    End If
    If gblnLED And Trim(txtPatient.Text) = "" Then
        zl9LedVoice.DisplayPatient ""    '˫����ʾ��������ڵ�ǰ������ʾ֮�������ʾ�����ƶ�����
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
                intIndex = IDKind.GetKindIndex("IC����")
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
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub GetFact(Optional blnFirst As Boolean = False, Optional ByVal int���� As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ͬ���ķ�Ʊ
    '����:���˺�
    '����:2011-07-19 17:47:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPati As Collection
    Dim strFactNO As String, intԤ������ As Integer
    
    'Ʊ�����ü�鼰��ʼ
    '����Ʊ�ݴ���
    If mobjEInvoice Is Nothing Then Exit Sub
    txtFact.Text = ""
    If cboType.ListIndex <> -1 Then intԤ������ = cboType.ItemData(cboType.ListIndex)
    If mobjEInvoice.zlIsStartEInvoice(int����, intԤ������) Then
        If blnFirst Then Exit Sub
        If mobjEInvoice.zlGetTranPaperInvoiceModule = 0 Then Exit Sub
        If mobjEInvoice.zlIsHisManagerInvoice = False Then
            Call mobjEInvoice.zlGetPatiCollectFromPatiObject(mpatiInfo, cllPati)
            Call mobjEInvoice.zlGetNextInvoiceNo(Me, strFactNO, cllPati, mlng����ID)
            If strFactNO <> "" Then txtFact.Text = strFactNO
            Exit Sub
        End If
    End If
    
    If mFactProperty.intInvoicePrint = 0 Then Exit Sub
   'Ʊ�����ü�鼰��ʼ
    If gblnBillԤ�� Then
        mlng����ID = CheckUsedBill(2, IIf(mlng����ID > 0, mlng����ID, mFactProperty.lngShareUseID), "", mFactProperty.strUseType)
        If mlng����ID <= 0 Then
            Select Case mlng����ID
                Case 0 '����ʧ��
                Case -1
                    MsgBox "��û�����ú͹��õ�Ԥ��Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End Select
            If blnFirst Then mblnUnLoad = True: Exit Sub
        End If
        '�ϸ�ȡ��һ������
        txtFact.Text = GetNextBill(mlng����ID)
    Else
        '��ɢ��ȡ��һ������
        txtFact.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, mlngFactModule, "")))
    End If
End Sub

Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ģ�����
    '����:���˺�
    '����:2012-02-27 11:23:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
        
    On Error GoTo errHandle
    
    mstrȱʡ���㷽ʽ = zlDatabase.GetPara("ȱʡԤ�����㷽ʽ", glngSys, mlngModul)
    mbytBackMoneyType = Val(zlDatabase.GetPara("�˿��ֹ��ʽ", glngSys, mlngModul))
    '���㷽ʽ:���|���㷽ʽ:���....
    mblnClearWinInfor = IIf(zlDatabase.GetPara("��Ԥ���������Ϣ", glngSys, glngModul) <> "1", True, False)
    mblnδ��Ʋ���Ԥ�� = zlDatabase.GetPara("����δ��Ʋ�׼��Ԥ��", glngSys, mlngModul, , , InStr(mstrPrivs, ";��������;") > 0) = "1"
    gblnSeekName = Nvl(zlDatabase.GetPara("����ģ������", glngSys, mlngModul, 1)) = 1
    mblnסԺ��Ԥ����֤ = zlDatabase.GetPara("סԺ��Ԥ����֤", glngSys, mlngModul, "0") = "1"
    mbln������Ժ��������˿� = zlDatabase.GetPara("������Ժ��������˿�", glngSys, mlngModul, "1") = "1"
    'ˢ��Ҫ����������
    mblnCheckPass = Mid(zlDatabase.GetPara(46, glngSys, , "0000000000"), 8, 1) = "1"
    mbln�ų�δ�ɼ�δ�� = zlDatabase.GetPara("ʣ����ų�δ�ɼ�δ����", glngSys, mlngModul, "0") = "1"

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
    
    'Ʊ�����ü�鼰��ʼ
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
    
    lblTitle.Caption = gstrUnitName & "����˿�"
    mstrCardPrivs = ";" & GetPrivFunc(glngSys, 1151) & ";"
    
    If gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.��� & "��Ϊ������", mlngModul, gcnOracle
    End If
    
    Call zlCheckFactIsEnough
    
    IDKind.IDKind = Val(zlDatabase.GetPara("�ϴ����뷽ʽ", glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0))
    
    '81693:���ϴ�,2015/4/21,������
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
        If MsgBox("�㵱ǰ���ڽ�������˿���Ƿ����Ҫ�˳�?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    mblnUnLoad = False
    mlng����ID = 0: mstr�����ʻ� = ""
    mstr�˿����Ա = "": mblnOptErrBill = False
    
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
    zlDatabase.SetPara "�ϴ����뷽ʽ", IDKind.IDKind, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    Set mobjThridSwap = Nothing
    Set mobjCards = Nothing
    Set mpatiInfo = Nothing
End Sub

Private Sub InitPrepayType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Ԥ������
    '����:���˺�
    '����:2011-07-14 18:50:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    With cboType
        .Clear
        cboType.Tag = "-1"
        mblnNotClick = True
        If InStr(1, mstrPrivs, ";����Ԥ��;") > 0 Then
            .AddItem "����Ԥ��": .ItemData(.NewIndex) = 1
            If mbytPrepayType = 1 Then .ListIndex = .NewIndex
        End If
        
        If InStr(1, mstrPrivs, ";סԺԤ��;") > 0 Then
            .AddItem "סԺԤ��": .ItemData(.NewIndex) = 2
            If mbytPrepayType = 2 Then .ListIndex = .NewIndex
        End If
        
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
        If cboType.ListCount = 0 Then
            MsgBox "�㲻�߱�����Ԥ����סԺԤ��Ȩ�ޣ�����ϵͳ����Ա��ϵ!", vbInformation + vbOKOnly, gstrSysName
            mblnUnLoad = True
        End If
        mblnNotClick = False
     End With
End Sub


Private Sub InitFace(Optional blnSave As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ڲ������ô�����漰����״̬
    '����:���˺�
    '����:2011-07-17 10:36:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    If Not gobjSquare.objSquareCard Is Nothing And blnSave = False Then
        IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
    End If
    
    On Error GoTo errHandle
    
    strSQL = "Select ����, ����, ����, ȱʡ��־ From ����Ԥ��ժҪ Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboNote.Clear
    If rsTmp.RecordCount > 0 Then
        While Not rsTmp.EOF
            cboNote.AddItem Nvl(rsTmp!����)
            rsTmp.MoveNext
        Wend
    End If
    
    cboNote.ListIndex = -1: Call InitPrepayType
    If mblnUnLoad Then Exit Sub
    
    IDKind.Enabled = True
    
    '����������
    Call CreateIDAndICCardObject
    cboNO.Text = ""
    
    Call Load֧����ʽ
    
    lblMoney.Caption = "�˿���": lblMoney.FontBold = True: lblMoney.ForeColor = vbRed
    txtMoney.ForeColor = vbRed: txtMoney.Font.Bold = True
    
    txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ�ݺ�") And gblnBillԤ�� '89302
     
    If lbl�ʻ����.Visible = False Then lblԤ�����.Left = lbl�ʻ����.Left
    
    If lbl�ʻ����.Visible Then
        Line2(14).Visible = True: Line2(11).x2 = 2415
    Else
        Line2(14).Visible = False: Line2(11).x2 = Line2(14).x2
    End If
    
    Call mobjThridSwap.zlInitCompents(Me, mlngModul, mobjICCard)
    If mbln�ų�δ�ɼ�δ�� Then
        lblʣ����.ToolTipText = "ʣ��� = Ԥ����� + ҽ��Ԥ���� - δ����� - δ�ɷ��� - δ�����"
    Else
        lblʣ����.ToolTipText = "ʣ��� = Ԥ����� + ҽ��Ԥ���� - δ�����"
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
    '����:����IC��ID����
    '����:���˺�
    '����:2018-08-30 14:39:06
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
        .ColWidth(.ColIndex("�˿���")) = IIf(.Rows * .RowHeight(0) <= .Height, 1855, 1620)
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
            Case pg_Page.pg_Ԥ������˿�
                Set objTemp = picDepositBack
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            Case pg_Page.pg_Ԥ����ʷ��¼
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
    '����27363
    If KeyAscii = 13 Then
        If txtCashTotal.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    '�˿�ʱ���������븺��
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
    If mpatiInfo.����ID = 0 Then Exit Sub
    dblMoney = StrToNum(txtCashTotal)
    
    If Val(lblCashTotal.Tag) <> dblMoney Then
        If MsgBox("�㵱ǰ������˿������˿��б��е����ֽ�һ��,�Ƿ��Զ��������ֽ�" & vbCrLf & vbCrLf & _
               "���ֺϼ�:" & Format(Val(lblCashTotal.Tag), "###0.00###") & vbCrLf & _
               "������:" & Format(dblMoney, "###0.00###") & vbCrLf & _
               "", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Call zlControl.ControlSetFocus(txtCashTotal): Exit Sub
        End If
        '�Զ���̯
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
    Dim dbl�˿��� As Double, dbl��� As Double
    dbl�˿��� = Val(lblCashTotal.Tag)
    dbl��� = dbl�˿��� - Val(txtMoney.Text)
    If dbl��� < 0 Then
        txt�տ�.Text = Format(-1 * dbl���, "0.00")
    Else
        txt�տ�.Text = "0.00"
    End If
End Sub

Private Sub txtMoney_GotFocus()
    txtMoney.SelStart = 0: txtMoney.SelLength = Len(txtMoney.Text)
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    '����27363
    If KeyAscii = 13 Then
        If txtMoney.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    '�˿�ʱ���������븺��
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
    '����27363
    Dim dblMoney  As Double
    If Not IsNumeric(StrToNum(txtMoney.Text)) Then txtMoney.SetFocus: Exit Sub
    If mpatiInfo.����ID = 0 Or IsNumeric(StrToNum(txtMoney.Text)) = False Then Exit Sub
    txtMoney.Text = Format(StrToNum(txtMoney.Text), "##,##0.00;-##,##0.00; ;")
    If txtMoney.MaxLength > 12 Then txtMoney.MaxLength = 12
    '108813:���ϴ�,2017/5/8,������������
    If gblnLED Then
        '#22 1234.56   --Ԥ��һǧ������ʮ�ĵ�����Ԫ Y
        '#23 1234.56   --����һǧ������ʮ�ĵ�����Ԫ Z
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
    If mpatiInfo.����ID = 0 Then Exit Sub
    dblMoney = StrToNum(txtMoney.Text)
    If Val(lblCashTotal.Tag) > dblMoney Then
        MsgBox "�����˿��" & Format(dblMoney, "###0.00###") & "��С�ڱ���Ӧ�˽�" & Format(Val(lblCashTotal.Tag), "###0.00###") & _
                     "�����������˿���Ϊ" & Format(Val(lblCashTotal.Tag), "###0.00###") & "Ԫ��" & vbCrLf & "", vbInformation, gstrSysName
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
        
    '�����ַ�������Form_KeyPress�н���
    If IDKind.GetCurCard.���� = "����" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.GetCurCard.���� = "�����" Or IDKind.GetCurCard.���� = "סԺ��" Or IDKind.GetCurCard.���� = "�ֻ���" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
    End If
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        Set frmPatiSelect.mfrmParent = Me
        frmPatiSelect.mbytSize = 1 '������(С��)
        frmPatiSelect.Show 1, Me
        blnSel = True
    End If
    
    Me.Refresh
    '����27379
    mstr�������� = ""
    txtPatient.ForeColor = &HFF0000
    
    'ˢ����ϻ���������س�
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
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnCancel As Boolean, bytPrepayType As Byte
    
    Call ClearBill
    '��ȡ������Ϣ
    SetMoneyInfo True
    sta.Panels(2) = ""
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If Not GetPatient(objCard, strInput, blnCancel, blnCard) Then
        '�����쳣��������
        If mblnOptErrBill = False Then
            If blnCancel Then 'ȡ������
                Call zlControl.TxtSelAll(txtPatient): txtPatient.SetFocus: Exit Sub
            End If
            sta.Panels(2) = "δ�ҵ��ò��ˣ�������������!"
            If blnCard = True Then
                txtPatient.PasswordChar = "": txtPatient.Text = ""
                '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
                txtPatient.IMEMode = 0
            Else
                txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
            End If
            Set mpatiInfo = New clsPatientInfo
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        End If
        Exit Sub
    End If
    
    '���ò��˷�����Ϣ
    Call SetMoneyInfo(False, mpatiInfo.����ID)
    Call LoadPatiPage(mpatiInfo.����ID)
    
    '79361:���ϴ�,2014/11/18,ȱʡ���˵�Ԥ������
    bytPrepayType = IIf(mpatiInfo.��Ժ, 2, 1)
    If bytPrepayType <> mbytPrepayType Then
        mbytPrepayType = bytPrepayType: Call InitPrepayType
    End If
    Call LoadPatiInforToContronl '���ز�����Ϣ
    
    Call Led��ӭ��Ϣ
    Call SetcmdOkEnabled
    Call zlCommFun.PressKey(vbKeyTab)
    
    '���¼��ص�ǰ����˿���Ϣ
    Call LoadThirdDelDeposit
    '������ʷԤ����¼
    Call ShowHistoryPrepay
End Sub

Private Sub Led��ӭ��Ϣ()
    Dim strInfo As String, lngPatient As Long
    'LED��ʼ��
    If Not gblnLED Then Exit Sub
    If gblnLedWelcome Then
        zl9LedVoice.Reset com
        zl9LedVoice.Speak "#1"
        zl9LedVoice.Init UserInfo.��� & "��Ϊ������", mlngModul, gcnOracle
    End If
    strInfo = Trim(txtPatient.Text)
    If mpatiInfo.����ID > 0 Then strInfo = strInfo & " " & mpatiInfo.�Ա� & " " & mpatiInfo.����: lngPatient = mpatiInfo.����ID
    zl9LedVoice.DisplayPatient strInfo, lngPatient

End Sub

Private Sub Clear�����ʻ�()
    '���ܣ���������ʻ���Ϣ
    Dim i As Integer
    
    On Error GoTo errHandle
    
    For i = 0 To cboStyle.ListCount - 1
        If cboStyle.ItemData(i) = 3 Then
            cboStyle.RemoveItem i: Exit For
        End If
    Next
    mcur�ʻ���� = 0
    lbl�ʻ����.Caption = lbl�ʻ����.Tag
    lbl�ʻ����.Visible = False: Line2(14).Visible = False
    Line2(11).x2 = Line2(14).x2
    lblԤ�����.Left = lbl�ʻ����.Left
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
                                           Optional ByRef blnCancel As Boolean, _
                                           Optional ByVal blnCard As Boolean, _
                                           Optional ByVal lng����id As Long, _
                                           Optional ByVal lng��ҳID As Long = -1) As Boolean
    '���ܣ���ȡ������Ϣ
    '������strInput=[ˢ��]|[A����ID]|[BסԺ��]
    '          lng��ҳID=-1��ʾ���ﲡ�˻��������סԺ����;lng��ҳID=0��ʾԤ��Ժ����;lng��ҳID>0��ʾסԺ����
    '˵����
    '     1.�����ڲ���Ԥ����
    '     2.�Զ�ʶ������Ժ״̬,����(����ID,��ҳID,����,�Ա�,����,סԺ��,����,��Ժ��־)
    '����:�Ƿ��ȡ�ɹ�,�ɹ�ʱmPatiInfo�а���������Ϣ,ʧ��ʱ���mPatiInfo
    Dim lng�����ID As Long
    Dim strWhere As String, strPassWord As String, strErrMsg As String
    Dim blnHavePassWord As Boolean, blnIsMobileNO As Boolean
    
    blnCancel = False: mstr�˿����Ա = ""

    If lng����id > 0 Then GoTo ReadPati

    blnIsMobileNO = IDKind.IsMobileNo(strInput)
    Call Clear�����ʻ� '��������ʻ���Ϣ
    
    If (blnCard And objCard.���� Like "����*") And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then   'ˢ����ȱʡ�Ŀ�
        lng�����ID = IDKind.GetDefaultCardTypeID
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����id, strPassWord, strErrMsg) = False Then
            If Not blnIsMobileNO Then GoTo NotFoundPati
            If gobjSquare.objSquareCard.zlGetPatiID("�ֻ���", strInput, False, lng����id, strPassWord) = False Then GoTo NotFoundPati
        Else
            blnHavePassWord = True
        End If
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then  '����ID
        lng����id = Mid(strInput, 2)
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then  'סԺ��(��ס(��)Ժ�Ĳ���)
        If Val(Mid(strInput, 2)) = 0 Then GoTo NotFoundPati
        If zlGetPatiIDByInNo(Mid(strInput, 2), lng����id, lng��ҳID) = False Then GoTo NotFoundPati
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(�������ﲡ��)
        If Val(Mid(strInput, 2)) = 0 Then GoTo NotFoundPati
        If GetPatiID("�����", Mid(strInput, 2), lng����id) = False Then GoTo NotFoundPati
    Else '��������
        Select Case objCard.����
            Case "����", "��������￨"
                '����ģ���鳤��,��������ղ��һ�Ӱ������
                If (Not gblnSeekName) Or (gblnSeekName And Len(strInput) < 2) Then GoTo NotFoundPati
                If GetPatiIdFromPatiName(txtPatient, strInput, lng����id, Me, , , , , blnCancel) = False Then GoTo NotFoundPati
            Case "ҽ����"
                strInput = UCase(strInput)
                If GetPatiID("ҽ����", strInput, lng����id) = False Then GoTo NotFoundPati
            Case "�����"
                If Not IsNumeric(strInput) Then GoTo NotFoundPati
                If Val(strInput) = 0 Then GoTo NotFoundPati
                If GetPatiID("�����", strInput, lng����id) = False Then GoTo NotFoundPati
            Case "סԺ��"
                If Not IsNumeric(strInput) Then GoTo NotFoundPati
                If Val(strInput) = 0 Then GoTo NotFoundPati
                If zlGetPatiIDByInNo(Val(strInput), lng����id, lng��ҳID) = False Then GoTo NotFoundPati
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����id, strPassWord, strErrMsg, lng�����ID) = False Then GoTo NotFoundPati
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����id, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati
                End If
                blnHavePassWord = True
        End Select
    End If

ReadPati:
    If lng����id <= 0 Then GoTo NotFoundPati
    If GetPatiInfo(lng����id, lng��ҳID, mpatiInfo) = False Then GoTo NotFoundPati
    If mpatiInfo.����ID = 0 Then GoTo NotFoundPati
    
    On Error GoTo Errhand
    '�����쳣����
    If OptOthersErrBill(mpatiInfo.����ID) Then
        Exit Function
    End If
    '��Ҫ��������
    If mblnCheckPass And (blnCard Or IDKind.GetCurCard.�ӿ���� <> 0) Then
        If Not blnHavePassWord Then
            strPassWord = mpatiInfo.����֤��
        End If
        If strPassWord <> "" Then
            If CreatePublicExpense() Then
                If gobjPublicExpense.zlVerifyPassWord(Me, strPassWord, mpatiInfo.����, mpatiInfo.�Ա�, mpatiInfo.����) = False Then GoTo NotFoundPati
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
    '����:��ȡ������Ϣ����
    '����:objPati_Out-���ز�����Ϣ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-30 15:14:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng��ҳID As Long
    On Error GoTo errHandle
    If mpatiInfo.����ID = 0 Then Exit Function
    
    Set objPati_Out = New clsPatiInfo
    With objPati_Out
        .���� = mpatiInfo.����
        .�Ա� = mpatiInfo.�Ա�
        .���� = mpatiInfo.����
        .��ҳID = mpatiInfo.��ҳID
        .����ID = mpatiInfo.����ID
        .����� = mpatiInfo.�����
        .סԺ�� = mpatiInfo.סԺ��
        .ҽ�Ƹ��ʽ = mpatiInfo.ҽ�Ƹ��ʽ
    End With
    lng��ҳID = IIf(cboType.ItemData(cboType.ListIndex) = 2, mpatiInfo.��ҳID, 0)
    If cboPatiPage.Visible And cboPatiPage.ListIndex > 0 And cboType.ItemData(cboType.ListIndex) = 2 Then
        lng��ҳID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    objPati_Out.��ҳID = lng��ҳID
        
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
    '����:ִ���б��е������˿�
    '���:objPati-��ǰ������Ϣ
    '     objCurItem-��ǰ�˿���
    '     blnVocherPrint-�Ƿ��ӡԤ��ƾ��
    '����:objdelItem_Out-��ǰ��Ч���˿���
    '     blnSave_out-�Ƿ��Ѿ�����������
    '     str����ID-ʹ����ɺ󣬷���"",���򷵻�ԭ����ID
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-31 10:30:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItemTemp As clsBalanceItem
    Dim blnChangeMoney As Boolean, cllPro As Collection
    Dim intԤ������ As Integer, bln����Ʊ�� As Boolean, int���� As Integer
    
    On Error GoTo errHandle
    
    blnSave_out = False
    If cboType.ListIndex <> -1 Then intԤ������ = cboType.ItemData(cboType.ListIndex)
    int���� = IIf(objCurItem.�������� = 3, mpatiInfo.����, 0)
    bln����Ʊ�� = mobjEInvoice.zlIsStartEInvoice(int����, intԤ������)
    objPati.���� = int����
    Call GetFact(False, int����)
    objCurItem.��Ʊ�� = IIf(bln����Ʊ��, "", txtFact.Text)
    objCurItem.����ID = mlng����ID
    
    If mobjThridSwap.zlThird_ReturnMoney_IsValied(objPati, objCurItem, 2, objItemTemp, False) = False Then
        Exit Function
    End If
    
    Set cllPro = New Collection
    If mobjThridSwap.zlThird_ReturnMoney(objPati, objCurItem, cllPro, objDelItem_Out, blnSave_out, False, blnChangeMoney, , int����, bln����Ʊ��) = False Then
        Exit Function
    End If
       
    Excute_Third_ReturnMoney = True
    '����NO
    Call AddComboxNoFromNo(objDelItem_Out.���ݺ�)
    '��ӡ���Ʊ��
    Call PrintDepostBill(objDelItem_Out.���ݺ�, blnVocherPrint, int����)
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
    '����:ִ���б��е�����Ϊ2���������㷽ʽ�˿�
    '���:objPati-��ǰ������Ϣ
    '     objCurItem-��ǰ�˿���
    '     blnVocherPrint-�Ƿ��ӡԤ��ƾ��
    '����:objdelItem_Out-��ǰ��Ч���˿���
    '     blnSave_out-�Ƿ��Ѿ�����������
    '     str����ID-ʹ����ɺ󣬷���"",���򷵻�ԭ����ID
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-31 10:30:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllSquare As Collection
    Dim intԤ������ As Integer, bln����Ʊ�� As Boolean, int���� As Integer
    On Error GoTo errHandle
    blnSave_out = False
    
    If cboType.ListIndex <> -1 Then intԤ������ = cboType.ItemData(cboType.ListIndex)
    int���� = IIf(objCurItem.�������� = 3, mpatiInfo.����, 0)
    bln����Ʊ�� = mobjEInvoice.zlIsStartEInvoice(int����, intԤ������)
    objPati.���� = int����
    Call GetFact(False, int����)
    objCurItem.��Ʊ�� = IIf(bln����Ʊ��, "", txtFact.Text)
    objCurItem.����ID = mlng����ID
    
    If Not mobjThridSwap.zlSquare_ReturnMoneySQL(objPati, cllSquare, objCurItem, , int����, bln����Ʊ��) Then Exit Function
    
    blnSave_out = True
    Excute_Square_ReturnMoney = True
    objCurItem.�Ƿ񱣴� = True
    objCurItem.�Ƿ���� = True
    objCurItem.�Ƿ�Ԥ�� = True
    objCurItem.�Ƿ�����༭ = False
    objCurItem.�Ƿ�����ɾ�� = False
    objCurItem.�Ƿ��������� = False
    
    Set objDelItem_Out = objCurItem
    Call AddComboxNoFromNo(objCurItem.���ݺ�)
    '��ӡ���Ʊ��
    Call PrintDepostBill(objCurItem.���ݺ�, blnVocherPrint, int����)
        
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
    '����:ִ���б��е�����Ϊ2���������㷽ʽ�˿�
    '���:objPati-��ǰ������Ϣ
    '     objCurItem-��ǰ�˿���
    '     blnVocherPrint-�Ƿ��ӡԤ��ƾ��
    '����:objdelItem_Out-��ǰ��Ч���˿���
    '     blnSave_out-�Ƿ��Ѿ�����������
    '     str����ID-ʹ����ɺ󣬷���"",���򷵻�ԭ����ID
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-31 10:30:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, blnTrans As Boolean
    Dim intԤ������ As Integer, bln����Ʊ�� As Boolean, int���� As Integer
    
    On Error GoTo errHandle
    blnSave_out = False
    
    '�����������ֲ���
    If cboType.ListIndex <> -1 Then intԤ������ = cboType.ItemData(cboType.ListIndex)
    int���� = IIf(objCurItem.�������� = 3, mpatiInfo.����, 0)
    bln����Ʊ�� = mobjEInvoice.zlIsStartEInvoice(int����, intԤ������)
    objPati.���� = int����
    Call GetFact(False, int����)
    objCurItem.��Ʊ�� = IIf(bln����Ʊ��, "", txtFact.Text)
    objCurItem.����ID = mlng����ID
    
    If mobjThridSwap.zlGetSaveSQLfromItem(objPati, objCurItem, 0, cllPro, True, , int����, bln����Ʊ��) = False Then Exit Function
    
    blnTrans = True
    Call zlExecuteProcedureArrAy(cllPro, Me.Caption)
    blnTrans = False
    
    
    blnSave_out = True
    objCurItem.�Ƿ񱣴� = True
    objCurItem.�Ƿ���� = True
    objCurItem.�Ƿ�Ԥ�� = True
    objCurItem.�Ƿ�����༭ = False
    objCurItem.�Ƿ�����ɾ�� = False
    objCurItem.�Ƿ��������� = False
    
    Set objDelItem_Out = objCurItem
    Excute_ListOther_ReturnMoney = True '����NO
    Call AddComboxNoFromNo(objCurItem.���ݺ�)
    '��ӡ���Ʊ��
    Call PrintDepostBill(objCurItem.���ݺ�, blnVocherPrint, int����)
    
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
    '����:ִ���б��е������˿�
    '���:objPati-��ǰ������Ϣ
    '     objCurItem-��ǰ�˿���
    '     blnVocherPrint-�Ƿ��ӡԤ��ƾ��
    '����:objdelItem_Out-��ǰ��Ч���˿���
    '     blnSave_out-�Ƿ��Ѿ�����������
    '     str����ID-ʹ����ɺ󣬷���"",���򷵻�ԭ����ID
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-31 10:30:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String, i As Long
    Dim objDelItems As clsBalanceItems, cllPro As Collection
    Dim intԤ������ As Integer, bln����Ʊ�� As Boolean, int���� As Integer
    
    On Error GoTo errHandle
    
    blnSave_out = False
    Set cllPro = New Collection
    Set objDelItems = objCurItem.objTag
    
    If cboType.ListIndex <> -1 Then intԤ������ = cboType.ItemData(cboType.ListIndex)
    int���� = IIf(objCurItem.�������� = 3, mpatiInfo.����, 0)
    bln����Ʊ�� = mobjEInvoice.zlIsStartEInvoice(int����, intԤ������)
    objPati.���� = int����
    Call GetFact(False, int����)
    objCurItem.��Ʊ�� = IIf(bln����Ʊ��, "", txtFact.Text)
    objCurItem.����ID = mlng����ID
    
    If mobjThridSwap.zlThird_TransferAccounts(objPati, objCurItem, cllPro, strErrMsg, blnSave_out, False, int����, bln����Ʊ��) = False Then
        If blnSave_out Then
            For i = 1 To objDelItems.Count
                vsBlance.Cell(flexcpForeColor, objDelItems(i).�к�, 0, objDelItems(i).�к�, vsBlance.Cols - 1) = vbRed
                vsBlance.RowData(objDelItems(i).�к�) = objDelItems(i)
            Next
        End If
        Exit Function
    End If
    
    Set objDelItem_Out = objCurItem
    For i = 1 To objDelItems.Count
        vsBlance.Cell(flexcpForeColor, objDelItems(i).�к�, 0, objDelItems(i).�к�, vsBlance.Cols - 1) = vbGrayed
        vsBlance.RowData(objDelItems(i).�к�) = objDelItems(i)
    Next
    Excute_Third_TransferAccounts = True
    
    '����NO
    Call AddComboxNoFromNo(objDelItem_Out.���ݺ�)
    
    '��ӡ���Ʊ��
    Call PrintDepostBill(objDelItem_Out.���ݺ�, blnVocherPrint, int����)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Excute_BalanceList_ReturnMoney(ByVal objPati As clsPatiInfo, ByRef lngԤ��ID_out As Long, _
    ByRef objDelCashItems_out As clsBalanceItems, ByVal blnVocherPrint As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���б��е������˿�
    '���:objPati-��ǰ������Ϣ
    '   blnVocherPrint-�Ƿ��ӡԤ��ƾ��
    '����:lngԤ��ID_Out-���һ��Ԥ��ID
    '     objDelCashItems_out-��ǰ�����б������ֵ���Ŀ
    '     str����ID-ʹ����ɺ󣬷���"",���򷵻�ԭ����ID
    '����:ִ�гɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-30 15:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTranItems As clsBalanceItems, objItems As clsBalanceItems, objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim blnFind As Boolean
    Dim i As Long, blnSaveed As Boolean
    Dim bln���� As Boolean
    
    On Error GoTo errHandle
    
    Set objTranItems = New clsBalanceItems
    If objDelCashItems_out Is Nothing Then Set objDelCashItems_out = New clsBalanceItems
    With vsBlance
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("�˿ʽ")) <> "" Then
                
                Set objItem = Nothing
                Call zlGetBalanceItemFromBalanceGrid(i, objItem)
                If objItem Is Nothing Then
                    MsgBox "�ڵ�" & i & "���е��˿���Ϣ��������!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                objItem.�к� = i
                objItem.�ɿλ = Trim(txtUnit.Text)
                objItem.����ID = IIf(cboUnit.ItemData(cboUnit.ListIndex) = 0, 0, cboUnit.ItemData(cboUnit.ListIndex))
                bln���� = GetVsGridBoolColVal(vsBlance, i, .ColIndex("����"))
                
                If bln���� Then
                    '���ִ���
                    objDelCashItems_out.AddItem objItem
                    objDelCashItems_out.������ = roundEx(objDelCashItems_out.������ + objItem.������, 6)
                    
                ElseIf objItem.�Ƿ�ת�� And objItem.���ѿ� = False Then
                    If objItem.�Ƿ���� = False Then
                        '�ȼ���ת��
                        blnFind = False
                        For Each objItemTemp In objTranItems
                            If objItemTemp.�����ID = objItem.�����ID Then
                                'ͬһ�����ģ������һ��ת
                                Set objItems = objItemTemp.objTag
                                If objItems Is Nothing Then Set objItems = New clsBalanceItems
                                objItems.AddItem objItem
                                objItems.������ = roundEx(objItems.������ + objItem.������, 2)
                                Set objItemTemp.objTag = objItems
                                
                                objItemTemp.������ = objItems.������
                                blnFind = True
                                objTranItems.������ = roundEx(objTranItems.������ + objItem.������, 2)
                                Exit For
                            End If
                        Next
                        If blnFind = False Then
                            Set objItemTemp = mobjThridSwap.zlCopyNewItemFromBalanceItem(objItem)
                            Set objItems = New clsBalanceItems
                            objItems.AddItem objItemTemp
                            objItems.������ = roundEx(objItems.������ + objItemTemp.������, 2)
                            Set objItem.objTag = objItems
                            objTranItems.AddItem objItem
                            objTranItems.������ = roundEx(objTranItems.������ + objItem.������, 2)
                        End If
                    End If
                ElseIf objItem.���ѿ� Then '���ѿ��˿�
                    '���ѿ���ش���
                    If objItem.�Ƿ���� = False Then
                        Set objItems = New clsBalanceItems
                        objItems.AddItem objItem
                        objItems.������ = roundEx(objItems.������ + objItem.������, 2)
                        Set objItem.objTag = objItems
                        If Excute_Square_ReturnMoney(objPati, objItem, objItemTemp, blnSaveed, blnVocherPrint) = False Then
                            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                            If blnSaveed Then
                                objItemTemp.�к� = i
                                Call zlSetVsBalanceEditStatus(objItemTemp, True)
                                .RowData(i) = objItemTemp
                            End If
                            Exit Function
                        End If
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbGrayed
                        lngԤ��ID_out = objItemTemp.ID
                        objItemTemp.�к� = i
                        Call zlSetVsBalanceEditStatus(objItemTemp, True)
                    End If
                Else
                    If objItem.�Ƿ���� = False Then
                        If objItem.�����ID <= 0 Then
                            '֧Ʊ�Ȳ�ԭ����
                            If Excute_ListOther_ReturnMoney(objPati, objItem, objItemTemp, blnSaveed, blnVocherPrint) = False Then
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                                If blnSaveed Then
                                    objItemTemp.�к� = i
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
                                    objItemTemp.�к� = i
                                    Call zlSetVsBalanceEditStatus(objItemTemp, True)
                                    .RowData(i) = objItemTemp
                                End If
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                                Exit Function
                            Else
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbGrayed
                            End If
                        End If
                        lngԤ��ID_out = objItemTemp.ID
                        objItemTemp.�к� = i
                        Call zlSetVsBalanceEditStatus(objItemTemp, True)
                    End If
                End If
            End If
        Next
    End With
    
    'ִ��ת�ʲ���
    For Each objItem In objTranItems
        If Excute_Third_TransferAccounts(objPati, objItem, objItemTemp, blnSaveed, blnVocherPrint) = False Then Exit Function
        lngԤ��ID_out = objItemTemp.ID
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
    ByRef objDelCashThird As clsBalanceItems, Optional lngԤ��ID_out As Long, _
    Optional ByVal blnVocherPrint As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˿�(��ͨ���㷽ʽ)
    '����:lngԤ��ID_out-�����Ԥ��ID
    '     ��ǰ���ֵ�������
    '   blnVocherPrint-�Ƿ��ӡԤ��ƾ��
    '����:
    '   str����ID-ʹ����ɺ󣬷���"",���򷵻�ԭ����ID
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-17 11:15:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCurItem As clsBalanceItem, objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim cllPro As Collection, objCard As Card
    Dim blnTrans As Boolean, dblMoney As Double, dblTotal As Double, dblTemp As Double
    Dim objDelCashItems As clsBalanceItems, int���� As Integer
    Dim intԤ������ As Integer, bln����Ʊ�� As Boolean
    
    On Error GoTo errH
    
    '�ȼ�鷢Ʊ���Ƿ�Ϸ�
    If CheckFactIsValied = False Then Exit Function
    
    Set cllPro = New Collection
    Set objCurItem = New clsBalanceItem
    Set objCard = mobjThridSwap.zlGetCardFromBalanceName(cboStyle.Text)
    
    If objDelCashThird Is Nothing Then Set objDelCashThird = New clsBalanceItems
    Set objDelCashItems = New clsBalanceItems
    
    If StrToNum(txtCashTotal.Text) = 0 Then Excute_CashAndOther_ReturnMoney = True: Exit Function
    
    If cboType.ListIndex <> -1 Then intԤ������ = cboType.ItemData(cboType.ListIndex)
    int���� = IIf(objCurItem.�������� = 3, mpatiInfo.����, 0)
    Call GetFact(False, int����)
    bln����Ʊ�� = mobjEInvoice.zlIsStartEInvoice(int����, intԤ������)
    objPati.���� = int����
    
    With objCurItem
        Set .objCard = objCard
        .������ = StrToNum(txtCashTotal.Text)
        .�������� = 0
        .���㷽ʽ = objCard.���㷽ʽ
        .����ժҪ = Trim(cboNote.Text)
        .������� = Trim(txtCode.Text)
        .������ = Trim(txt������.Text)
        .����ID = IIf(cboUnit.ItemData(cboUnit.ListIndex) = 0, 0, cboUnit.ItemData(cboUnit.ListIndex))
        .�ɿλ = Trim(txtUnit.Text)
        .����Ԥ�� = IIf(cboType.ItemData(cboType.ListIndex) = 1, True, False)
        .��Ʊ�� = IIf(bln����Ʊ��, "", txtFact.Text)
        .����ID = mlng����ID
    End With

    '�����������ֲ���
    For Each objItemTemp In objDelCashThird
        objDelCashItems.AddItem objItemTemp
        objDelCashItems.������ = roundEx(objDelCashItems.������ + objItemTemp.������, 2)
    Next
    
    If roundEx(objCurItem.������, 2) - roundEx(objDelCashThird.������, 2) > 0 Then
        If MsgBox("��ǰ�˿���Ȳ��˿�������(���������ʻ������˿���Ƿ����?" & vbCrLf & _
            "�������:" & Format(roundEx(objDelCashThird.������, 2), "0.00") & vbCrLf & _
            "�����˿�:" & Format(objCurItem.������, "0.00") & vbCrLf & "ע�� �������=��ͨ�������+�����ʻ��������ֺϼ� ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    Set objCurItem.objTag = objDelCashItems
    If mobjThridSwap.zlGetSaveSQLfromItem(objPati, objCurItem, 0, cllPro, True, , int����, bln����Ʊ��) = False Then Exit Function
    
    blnTrans = True
    Call zlExecuteProcedureArrAy(cllPro, Me.Caption)
    blnTrans = False
    
    '����NO
    Call AddComboxNoFromNo(objCurItem.���ݺ�)
    
    '��ӡ���Ʊ��
    Call PrintDepostBill(objCurItem.���ݺ�, blnVocherPrint, int����)
    
    lngԤ��ID_out = objCurItem.ID
    Excute_CashAndOther_ReturnMoney = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If Err.Description Like "*�˿�����ڲ���ʣ��Ԥ�����*" And mbytOracleBackType = 1 Then
        If MsgBox("�˿���Ȳ��˵�ǰ������,�Ƿ���ԣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
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
    '����:���ñ༭״̬
    '���:blnSetRowData-�Ƿ�objItem���ø�Rowdata����
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-09-03 10:06:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    If objItem Is Nothing Then Exit Sub
    
    lngRow = objItem.�к�
    With vsBlance
        If lngRow < 0 Or lngRow > .Rows - 1 Then Exit Sub
        .TextMatrix(lngRow, .ColIndex("����״̬")) = IIf(objItem.�Ƿ����, 1, 0)
        .TextMatrix(lngRow, .ColIndex("�༭״̬")) = IIf(objItem.�Ƿ�����༭, 1, 0) & "|" & IIf(objItem.�Ƿ�����ɾ��, 1, 0)
        If blnSetRowData Then .RowData(lngRow) = objItem
    End With
End Sub


Private Sub AddComboxNoFromNo(ByVal strDepositNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݺţ������ݺż��뵥�ݺ���������
    '����:���˺�
    '����:2018-08-31 09:50:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, i As Long
    
    
    On Error GoTo errHandle
    '���뵥����ʷ��¼(�������͵���)
    strNO = strDepositNo
    For i = 0 To cboNO.ListCount - 1
        strNO = strNO & "," & cboNO.List(i)
    Next
    cboNO.Clear
    For i = 0 To UBound(Split(strNO, ","))
        cboNO.AddItem Split(strNO, ",")(i)
        If i = 9 Then Exit For 'ֻ��ʾ10��
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub PrintDepostBill(ByVal strNO As String, ByVal blnVocherPrint As Boolean, Optional ByVal int���� As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡԤ��Ʊ��
    '���:blnPrint-�Ƿ��ӡԤ��Ʊ��
    '       blnVocherPrint-�Ƿ��ӡԤ��ƾ��
    '       strNO-���ݺ�
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-31 09:43:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intԤ������ As Integer
    
    On Error GoTo errHandle
    If cboType.ListIndex <> -1 Then intԤ������ = cboType.ItemData(cboType.ListIndex)
    If mobjEInvoice.zlIsStartEInvoice(int����, intԤ������) Then Exit Sub
    
    If blnVocherPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1103_2", Me, "NO=" & strNO, 2)
    End If
    
    If mintPrintType < 0 Then
        Select Case mFactProperty.intInvoicePrint
            Case 0 '����ӡԤ����Ʊ
               mintPrintType = 0
            Case 1 '�Զ���ӡ
               mintPrintType = 1
            Case 2 '��ӡ����
                If MsgBox("�Ƿ���Ҫ��ӡԤ��Ʊ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mintPrintType = 1
                Else
                    mintPrintType = 0
                End If
        End Select
    End If
    
    If mintPrintType = 0 Then Exit Sub

    If Not gblnBillԤ�� And Trim(txtFact.Text) <> "" Then
        '��ɢ�����浱ǰ����
        zlDatabase.SetPara "��ǰԤ��Ʊ�ݺ�", Trim(txtFact.Text), glngSys, mlngFactModule
    End If
    
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & cboNO.List(0), "����ID=" & mpatiInfo.����ID, "�տ�ʱ��=" & Format(Now, "yyyy-mm-dd HH:MM:SS"), _
    IIf(mFactProperty.intInvoiceFormat = 0, "", "ReportFormat=" & mFactProperty.intInvoiceFormat), 2)
    
    Call zlCheckFactIsEnough
'    Call GetFact
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub GetDepositData(ByVal lng����id As Long, Optional ByVal int��ҳID As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¶�ȡԤ������
    '���:lng����ID-����ID��
    '����:���˺�
    '����:2011-07-22 17:02:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strWhere As String
    
    On Error GoTo errHandle
    If lng����id = 0 Then
        If mpatiInfo.����ID = 0 Then Set mrsDepositBalance = Nothing: Exit Sub
        lng����id = mpatiInfo.����ID
    End If
    mdbl������� = 0: mdblԤ����� = 0: mdblʣ���� = 0
     '������Ȼ���,���������
    Set mrsDepositBalance = GetMoneyInfo(lng����id)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ShowPremayBalance(ByVal blnreReadData As Boolean, ByVal lng����id As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������صĽ��㷽ʽ����������,��ʾԤ�����
    '���:blnReRead-�ض�����
    '       lng����ID-��ȡָ���Ĳ���ID(0ʱ,��mPatiInfo�ж�ȡ����ID)
    '����:���˺�
    '����:2011-07-21 15:44:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset, strSQL As String
    Dim intԤ������ As Integer
    Dim dblδ�� As Double, dblδ�� As Double, dblYB As Double
    Dim lng��ҳID As Long, dblʣ���� As Double, int��ҳID As Integer
    
    On Error GoTo errHandle
    If lng����id = 0 Then
        If mpatiInfo.����ID = 0 Then Exit Sub
        lng����id = mpatiInfo.����ID
    End If
    If cboPatiPage.Visible And cboPatiPage.ListIndex >= 0 Then
        int��ҳID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    If blnreReadData Then Call GetDepositData(lng����id, int��ҳID)
    If cboType.ListIndex < 0 Then Exit Sub
    
    sta.Panels(2).Text = ""
    mdbl������� = 0: mdblԤ����� = 0: mdblʣ���� = 0
    intԤ������ = cboType.ItemData(cboType.ListIndex)
    
    If Not mrsDepositBalance Is Nothing Then
        With mrsDepositBalance
            .Filter = "����=" & intԤ������
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                mdbl������� = mdbl������� + Val(Nvl(!�������))
                mdblԤ����� = mdblԤ����� + Val(Nvl(!Ԥ�����))
                .MoveNext
            Loop
        End With
    End If
    
    'ҽ��Ԥ�����
    If cboPatiPage.Visible And cboPatiPage.ListIndex >= 0 And cboType.ItemData(cboType.ListIndex) = 2 Then
        lng��ҳID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    
    Set rsMoney = New ADODB.Recordset
    If lng��ҳID = 0 Then
        strSQL = "Select Sum(���) As ҽ��Ԥ�� From ����ģ����� Where ����ID = [1] And ��ҳID Is Null"
        Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ��Ԥ��", lng����id)
    Else
        strSQL = "Select Sum(���) As ҽ��Ԥ�� From ����ģ����� Where ����ID = [1] And ��ҳID = [2]"
        Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ��Ԥ��", lng����id, lng��ҳID)
    End If
    
    If Not rsMoney.EOF Then
        If Val(Nvl(rsMoney!ҽ��Ԥ��, 0)) > 0 Then
            dblYB = Val(Nvl(rsMoney!ҽ��Ԥ��, 0))
            lblҽ��Ԥ��.Caption = lblҽ��Ԥ��.Tag & Format(rsMoney!ҽ��Ԥ��, "##,##0.00;-##,##0.00; ;")
        Else
            lblҽ��Ԥ��.Caption = lblҽ��Ԥ��.Tag
        End If
    Else
        lblҽ��Ԥ��.Caption = lblҽ��Ԥ��.Tag
    End If
    
    mdblʣ���� = Format(mdblԤ����� - mdbl�������, "0.00")
    '����27363
    lbl�������.Caption = lbl�������.Tag & Format(mdbl�������, "##,##0.00;-##,##0.00; ;")
    lblԤ�����.Caption = lblԤ�����.Tag & Format(mdblԤ�����, "##,##0.00;-##,##0.00; ;")
    dblδ�� = GetUnAuditedFee(lng����id, , intԤ������)
    dblδ�� = GetUnAuditedFee(lng����id, False, intԤ������)
    lblδ�����.Caption = lblδ�����.Tag & Format(dblδ��, "##,##0.00;-##,##0.00; ;")
    lblδ�ɷ���.Caption = lblδ�ɷ���.Tag & Format(dblδ��, "##,##0.00;-##,##0.00; ;")
    dblʣ���� = IIf(mbln�ų�δ�ɼ�δ��, mdblʣ���� - dblδ�� - dblδ�� + dblYB, mdblʣ���� + dblYB)
    lblʣ����.Caption = lblʣ����.Tag & Format(dblʣ����, "##,##0.00;-##,##0.00; ;")
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub SetMoneyInfo(blnClear As Boolean, Optional lng����id As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ϣ
    '���:blnClear-���
    '     lng����ID-ָ������ID
    '����:���˺�
    '����:2011-07-21 15:40:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset
    Dim strSQL As String
    
    If blnClear Then
        lblSex.Caption = lblSex.Tag: mstrPatiSex = ""
        lblOld.Caption = lblOld.Tag: mstrPatiOld = ""
        lblPatientNO.Caption = lblPatientNO.Tag
        lbl����.Caption = lbl����.Tag
        lbl����.Caption = lbl����.Tag
        lbl��ͥ��ַ.Caption = lbl��ͥ��ַ.Tag
        lblҽ�Ƹ��ʽ.Caption = lblҽ�Ƹ��ʽ.Tag
        lbl������.Caption = lbl������.Tag
        lbl�������.Caption = lbl�������.Tag
        lblWorkUnit.Caption = lblWorkUnit.Tag
        
        lblδ�����.Caption = lblδ�����.Tag
        lblδ�ɷ���.Caption = lblδ�ɷ���.Tag
        lbl�������.Caption = lbl�������.Tag
        lblԤ�����.Caption = lblԤ�����.Tag
        lblʣ����.Caption = lblʣ����.Tag
        lblҽ��Ԥ��.Caption = lblҽ��Ԥ��.Tag
        lbl�ֻ���.Caption = lbl�ֻ���.Tag
        lbl���֤��.Caption = lbl���֤��.Tag
        lblӦ�տ�.Caption = lblӦ�տ�.Tag
        lblӦ�տ�.ForeColor = &H80000007
        
        mdbl������� = 0
        mdblԤ����� = 0
        mdblʣ���� = 0
    Else
        On Error GoTo errHandle
        '��ʾԤ�����
        Call ShowPremayBalance(True, lng����id)
        '����Ƿ���Ӧ�տ�
        strSQL = "Select Zl_Patientdue([1]) ʣ��Ӧ�� From dual"
        Set rsMoney = New ADODB.Recordset
        Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, "��ȡӦ�տ�", lng����id)
        If Not rsMoney.EOF Then
            If Nvl(rsMoney!ʣ��Ӧ��, 0) > 0 Then
                MsgBox "��ע�⣬�ò������� " & rsMoney!ʣ��Ӧ�� & "Ԫ Ӧ�տ�δ�ɣ�", vbInformation, gstrSysName
                lblӦ�տ�.Caption = lblӦ�տ�.Tag & Format(rsMoney!ʣ��Ӧ��, "##,##0.00;-##,##0.00; ;")
                lblӦ�տ�.ForeColor = &HFF&
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
    '����27379 by lesfeng 2010-01-18
    If mpatiInfo.����ID = 0 Then
        mstr�������� = mpatiInfo.��������
    End If
    If mstr�������� = "" Then
        If mpatiInfo.����ID > 0 Then
            If mpatiInfo.���� > 0 Then
                txtPatient.ForeColor = vbRed
            Else
                txtPatient.ForeColor = &HFF0000
            End If
        Else
            txtPatient.ForeColor = &HFF0000
        End If
    Else
        Call SetPatiColor(txtPatient, mstr��������)
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

Private Sub txt������_GotFocus()
    zlControl.TxtSelAll txt������
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("~!%^""'|`", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    Else
        CheckInputLen txt������, KeyAscii
    End If
End Sub

Private Sub txt�տ�_GotFocus()
    Call zlControl.TxtSelAll(txt�տ�)
End Sub

Private Sub txt�ʺ�_GotFocus()
    zlControl.TxtSelAll txt�ʺ�
End Sub

Private Sub txt�ʺ�_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("~!%^""'|`", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    Else
        CheckInputLen txt�ʺ�, KeyAscii
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
    '����:��ʼ�����סԺ�ٴ�������Ϣ
    '����:���˺�
    '����:2018-11-29 10:33:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,A.����,B.������� " & _
        "from ���ű� A,��������˵�� B " & _
        "Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        "and B.����ID=A.ID and B.������� IN(1,2,3) AND B.�������� IN('�ٴ�','����') " & _
        "Order by B.�������,A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboUnit.Clear
    cboUnit.AddItem "��"
    cboUnit.ItemData(0) = 0
    cboUnit.ListIndex = 0
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & IIf(IsNull(rsTmp!����), "", rsTmp!����)
            cboUnit.ItemData(cboUnit.ListCount - 1) = rsTmp!ID
            rsTmp.MoveNext
        Next
    End If
    
    If Not gbln�ɿ���� Then
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
    '����:��ʼ��IDKind�ؼ���ʶ����
    '����:���˺�
    '����:2018-11-29 10:36:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKind As String
    
    On Error GoTo errHandle
    
    strKind = "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;ס|סԺ��|0;��|���ۺ�|0;��|���￨|0;��|�ֻ���|0"
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKind, txtPatient)
    mtySquareCard.blnExistsObjects = Not gobjSquare.objSquareCard Is Nothing
'    gobjSquare.objSquareCard.mblnYLMgr = mbytCallObject = 2
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load֧����ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:���˺�
    '����:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objPayCard As Cards, str���� As String
    Dim objCard As Card
    
    '���㷽ʽ:���ò�ѯ��ҽ�ƿ�����ʱ��һ��ֻ֧��Ԥ����,�����ڴ��յ����
    'mbytCallObject:���õĶ���(0-Ԥ����������;1-���˷��ò�ѯ����;2-ҽ�ƿ��������;3-����Һŵ���...
    If InStr(1, mstrPrivs, ";Ԥ���տ�;") > 0 Or _
        InStr(1, mstrPrivs, ";Ԥ���տ�;") > 0 Or _
        InStr(1, mstrPrivs, ";Ԥ�������˿�;") > 0 Or _
        InStr(1, mstrPrivs, ";����Ԥ��תסԺ;") > 0 _
        Or InStr(1, mstrPrivs, ";סԺԤ��ת����;") > 0 Or mbytCallObject > 0 Then
        str���� = ",1,2,7,8,3"
    End If
    
    If str���� = "" Then str���� = ",1,2,7,8,3"
    
    str���� = Mid(str����, 2)
    
    If mblnNurseCall Then
        str���� = "7,8"
    End If
    
    If mobjThridSwap.zlReGetPayCards(str����, "Ԥ����", objPayCard) = False Then
        MsgBox "Ԥ������û�п��õĽ��㷽ʽ,���ȵ����㷽ʽ���������á�", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    If objPayCard.Count = 0 Then
        MsgBox "Ԥ������û�п��õĽ��㷽ʽ,���ȵ����㷽ʽ���������á�", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    
    '����˿ֻ������ͨ���㷽ʽ������˿�
    For i = 1 To objPayCard.Count
        Set objCard = objPayCard(i)
        If objCard.�ӿ���� <= 0 And objCard.�������� <> 3 Then
             cboStyle.AddItem objCard.���㷽ʽ
             cboStyle.ItemData(cboStyle.NewIndex) = objCard.��������
             If objCard.ȱʡ��־ And cboStyle.ListIndex < 0 Then cboStyle.ListIndex = cboStyle.NewIndex
             If objCard.���㷽ʽ = mstrȱʡ���㷽ʽ Then cboStyle.ListIndex = cboStyle.NewIndex
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
    '����:��鵱ǰƱ���Ƿ�����
    '����:���˺�
    '����:2012-09-06 15:41:52
    '˵��:
    '����:37372
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngʣ������ As Long, strType As String
    
    '��Ҫ���ʣ�������Ƿ����:
    If cboType.ListIndex < 0 Then
        strType = ""
    Else
        strType = cboType.ItemData(cboType.ListIndex)
    End If
    
    If zlCheckInvoiceOverplusEnough(2, gint����ʣ��Ʊ������, lngʣ������, mlng����ID, strType) = False Then
        MsgBox "ע��:" & vbCrLf & _
               "    ��ǰʣ��Ʊ��(" & lngʣ������ & ") С���˱���������(" & gint����ʣ��Ʊ������ & "),��ע�������Ʊ!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        zlCheckFactIsEnough = False: Exit Function
    End If
    zlCheckFactIsEnough = True
End Function

Private Sub LoadPatiPage(ByVal lng����id As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز��˵�סԺ����
    '����:���˺�
    '����:2012-12-11 10:19:58
    '˵��:
    '����:51628
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim bln���� As Boolean
    On Error GoTo errHandle
        
    cboPatiPage.Clear
    With cboPatiPage
        .AddItem "����סԺ": .Tag = 0
        .ItemData(.NewIndex) = -1
        
        If GetPatiPageNum(lng����id, rsTemp) = False Then Exit Sub
        If rsTemp.State = 0 Then Exit Sub
        Do While Not rsTemp.EOF
            If bln���� = False And Val(Nvl(rsTemp!��������, 0)) <> 0 Then bln���� = True
            If Val(Nvl(rsTemp!��ҳID)) = 0 And Val(Nvl(rsTemp!��������)) = 0 Then
                .AddItem "ԤԼ��Ժ"
            Else
                .AddItem "��" & rsTemp!��ҳID & "��" & IIf(Val("" & rsTemp!��������) = 1, "(��������)", IIf(Val("" & rsTemp!��������) = 2, "(סԺ����)", ""))
            End If
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!��ҳID))
            If .ListIndex < 0 Then .ListIndex = .NewIndex
            If mblnNurseCall Then
                If Val(Nvl(rsTemp!��ҳID)) = mlng��ҳID Then
                    .ListIndex = .NewIndex
                End If
                cboPatiPage.Enabled = False
            Else
                If Val(Nvl(rsTemp!��ҳID)) = mpatiInfo.��ҳID Then
                    .ListIndex = .NewIndex
                End If
            End If
            rsTemp.MoveNext
        Loop
        If .ListCount > 0 Then .ListIndex = 0
        If bln���� = True Then Call cbo.SetListWidth(cboPatiPage.hwnd, 2000)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function Checkδ��Ʋ���Ԥ��() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡���Ƿ����,δ���,����Ԥ��
    '����:���˺�
    '����:2012-12-11 10:19:58
    '˵��:
    '����:51628
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����id As Long, lng��ҳID As Long
    Dim str����id As String, PatiPageInfo As clsPatientInfo
    
    On Error GoTo errHandle
    
    If mblnδ��Ʋ���Ԥ�� = False Then Checkδ��Ʋ���Ԥ�� = True: Exit Function
    
    '����Ԥ�������
    If cboType.ItemData(cboType.ListIndex) <> 2 Then Checkδ��Ʋ���Ԥ�� = True: Exit Function
    
    '��ǰסԺ������Ϊ��Ժ��,Ҳ�����
    If mpatiInfo.��Ժ = False Then Checkδ��Ʋ���Ԥ�� = True: Exit Function
    
    lng����id = mpatiInfo.����ID
    '������סԺ������,Ҳ�ܽ�Ԥ��,��˲����
    If cboPatiPage.ListIndex < 0 Then Checkδ��Ʋ���Ԥ�� = True: Exit Function
    
    lng��ҳID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    str����id = lng����id & ":" & lng��ҳID
    Call GetPatiPageInforByID(str����id, PatiPageInfo, False)
    If PatiPageInfo.����� = 0 Then
        MsgBox "ע��" & vbCrLf & "   ���ˡ�" & mpatiInfo.���� & "��δ���,�������Ԥ����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Checkδ��Ʋ���Ԥ�� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function Check�˿�() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡���˿�ǰ����Ƿ���ڱ仯
    '����:���ϴ�
    '����:2016/2/25 10:21:39
    '˵��:
    '����:93144
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����id As Long
    Dim dblԤ����� As Double, dbl������� As Double, dblʣ����� As Double
    Dim intIndex As Integer
    Dim objCard As Card
    On Error GoTo errHandle
    
    If mpatiInfo.����ID = 0 Then Exit Function
    
    If cboType.ListIndex < 0 Then
        If StrToNum(txtMoney.Text) <> 0 Then
            MsgBox "δѡ��ָ���Ľ�����Ϣ!", vbInformation + vbOKOnly, gstrSysName
        End If
        Exit Function
    End If
    Check�˿� = True
    
    Exit Function
    
    intIndex = cboType.ItemData(cboType.ListIndex)
    Set objCard = mobjThridSwap.objPayCards(intIndex)
    lng����id = mpatiInfo.����ID
    Set mrsDepositBalance = GetMoneyInfo(lng����id)
    If Not mrsDepositBalance Is Nothing Then
        With mrsDepositBalance
            .Filter = "����=" & objCard.��������
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                dbl������� = dbl������� + Val(Nvl(!�������))
                dblԤ����� = dblԤ����� + Val(Nvl(!Ԥ�����))
                .MoveNext
            Loop
        End With
    End If
    dblʣ����� = Format(dblԤ����� - dbl�������, "0.00")
    If mdblʣ���� <> dblʣ����� Then
        MsgBox "���˵�ʣ������ѷ����仯,������ȷ���˿���!", vbInformation + vbOKOnly, gstrSysName
        Call ShowPremayBalance(False, 0)
        Exit Function
    End If
    
    Check�˿� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OptOthersErrBill(ByVal lng����id As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:�տ��������ⲡ���Ƿ�����������Ա�������쳣���ݣ�������
    '���: lng����ID
    '����:
    '����:2018-08-07
    '˵��:
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsErrBills As ADODB.Recordset
    Dim str����Ա���� As String, strTittle As String
    Dim strNO As String
    
    On Error GoTo errHandle
    '��Ȩ�ޣ���Ϊ�շ�״̬
    'type: 1-�쳣��ֵ��2-�쳣���ʣ�3-�쳣����˿�
    strSQL = "Select Type, No , ���� ,����Ա����" & vbNewLine & _
            "From (Select 2 Type, a.No, a.����, a.����Ա����" & vbNewLine & _
            "       From ����Ԥ����¼ a" & vbNewLine & _
            "       Where Nvl(У�Ա�־, 0) <> 0 And ��¼���� = 1 And ����id = [1] And ��¼״̬ = 2 " & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 3 Type, a.No, a.����, a.����Ա����" & vbNewLine & _
            "       From ����Ԥ����¼ a" & vbNewLine & _
            "       Where Nvl(У�Ա�־, 0) <>0 And ��¼���� = 1 And ����id = [1] And ��¼״̬ = 0 And A.���ӱ�־=1)" & vbNewLine & _
            "Order By Decode(����Ա����, [2], 0, 1), Type"
    Set rsErrBills = zlDatabase.OpenSQLRecord(strSQL, "�����쳣���ݲ�ѯ", lng����id, UserInfo.����)
    If rsErrBills.EOF Then Exit Function
    
    str����Ա���� = Nvl(rsErrBills!����Ա����)
    If Nvl(rsErrBills!type) = 2 Then
        strTittle = "����"
    Else
        strTittle = "����˿�"
    End If
    '��������Ա�ж�Ȩ��
    If str����Ա���� <> UserInfo.���� Then
        If InStr(mstrPrivs, ";�����������쳣����;") = 0 Then Exit Function
        If MsgBox("ע��:" & vbCrLf & _
            "       �ò��˴����ɲ���Ա��" & str����Ա���� & "�����������쳣" & strTittle & "���ݣ�" & vbCrLf & vbCrLf & _
            "       �Ƿ�Ըõ��ݽ���" & strTittle & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    Else
        If MsgBox("ע��:" & vbCrLf & _
            "       �ò��˴����쳣" & strTittle & "���ݣ��Ƿ����ڶԸõ��ݽ��д���", _
                    vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    strNO = Nvl(rsErrBills!NO)
    '�����´�����
    If mobjEInvoice Is Nothing Then Exit Function
    If frmDeposit.zlShowEdit(Me, mbytCallObject, 7, mobjEInvoice, mstrPrivs, mlngModul, mbytPrepayType, strNO) = False Then Exit Function
    OptOthersErrBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub Excute_Plug_PatiPrePayAfter(ByRef objPati As clsPatiInfo, ByVal lngԤ��ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ҵ��������ӿ�
    '���:objpati-������Ϣ����
    '����:���˺�
    '����:2018-08-31 10:25:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPlugIn Is Nothing Then Exit Sub
    '81693:���ϴ�,2015/4/21,������
    On Error Resume Next
    Call mobjPlugIn.PatiPrePayAfter(objPati.����ID, IIf(mbytPrepayType = 2, 1, 0), lngԤ��ID)
    Err.Clear
End Sub

Private Sub vsBlance_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    With vsBlance
       If .ColIndex("����") = Position Or Col = .ColIndex("����") Then
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
        Case .ColIndex("�˿ʽ")
        Case .ColIndex("����")
            Call ReCalePtBalanceMoney '���¼����˿���
        Case .ColIndex("�˿���")
            If Not zlGetBalanceItemFromBalanceGrid(Row, objItem) Then Exit Sub
             objItem.������ = roundEx(Val(.TextMatrix(Row, Col)), 2)
             .RowData(Row) = objItem
             Call ReCalePtBalanceMoney '���¼����˿���
        Case Else
        End Select
    End With
End Sub

Private Sub vsBlance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBlance, Me.Name, "�����б�"
End Sub

Private Sub vsBlance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vsBlance, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsBlance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBlance, Me.Name, "�����б�"
End Sub

Private Sub vsBlance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim objItem As clsBalanceItem
    
    If mpatiInfo.����ID = 0 Then Cancel = True: Exit Sub
    
    With vsBlance
        If Val(.TextMatrix(Row, .ColIndex("����״̬"))) = 1 Then Cancel = True: Exit Sub
        
        varTemp = Split(.TextMatrix(Row, .ColIndex("�༭״̬")) & "|||", "|")
        .ComboList = ""
        
        Select Case Col
        Case .ColIndex("�˿ʽ")
            If Not zlGetBalanceItemFromBalanceGrid(Row, objItem) Then Cancel = True: Exit Sub
            If Val(varTemp(1)) <> 1 Then Cancel = True: Exit Sub
            '�Ƿ�����༭|�Ƿ�����ɾ��
             .ColComboList(.ColIndex("�˿ʽ")) = ""
             .ComboList = "..."
             .CellButtonPicture = imgDel
        Case .ColIndex("����")
            If .TextMatrix(Row, .ColIndex("��������")) = 1 Then Cancel = True: Exit Sub
            If Not CheckDelCashColIsEdit(Row) Then Cancel = True: Exit Sub
        Case .ColIndex("�˿���")
            If ChecklDelMoneyIsEdit(.Row) = False Then Cancel = True: Exit Sub
        Case Else
            Cancel = True: Exit Sub
        End Select
    End With
End Sub

Private Sub vsBlance_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsBlance.ColIndex("����") Then Cancel = True
End Sub

Private Sub vsBlance_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim varData As Variant
    With vsBlance
        '�Ƿ�����༭|�Ƿ�����ɾ��
        varData = Split(.TextMatrix(Row, .ColIndex("�༭״̬")) & "||", "|")
        If varData(1) <> 1 Then Exit Sub
    End With
    
    Call DeletePayInfor(Row)
    
    Call ReCalePtBalanceMoney '���¼����˿���Ϣ
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
        
        'ɾ��
        If KeyCode = vbKeyDelete Then
            Call vsBlance_CellButtonClick(.Row, .Col)
            Exit Sub
        End If
    End With
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsBlance
        Select Case .Col
        Case .ColIndex("�˿ʽ")
            If Trim(.TextMatrix(.Row, .ColIndex("�˿ʽ"))) = "" And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        Case .ColIndex("�˿���")
            If (Trim(.TextMatrix(.Row, .ColIndex("�˿ʽ"))) = "" Or Val(.TextMatrix(.Row, .ColIndex("�˿���"))) = 0) And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        Case Else
            If (Trim(.TextMatrix(.Row, .ColIndex("�˿ʽ"))) = "" Or Val(.TextMatrix(.Row, .ColIndex("�˿���"))) = 0) And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        End Select
        Call zlVsMoveGridCell(vsBlance, .ColIndex("�˿ʽ"), , False, lngRow)
    End With
End Sub

Private Sub vsBlance_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim strKey As String, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsBlance
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "") '�ݲ���������
        Select Case Col
        Case .ColIndex("�˿ʽ")
        Case .ColIndex("�˿���")
        Case Else
        End Select
        Call zlVsMoveGridCell(vsBlance, .ColIndex("�˿ʽ"), -1, False, lngRow)
    End With

End Sub
Private Sub vsBlance_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsBlance
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "") '�ݲ���������
        Select Case Col
        Case .ColIndex("�˿ʽ")
        Case .ColIndex("�˿���")
        Case Else
        End Select
    End With
End Sub


Private Sub vsBlance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Err = 0: On Error GoTo Errhand:
    With vsBlance
        If .MouseRow < 1 Or .MouseRow > .Rows - 1 Then Exit Sub
        If .MouseCol < 0 Or .MouseCol > .Cols - 1 Then Exit Sub
        If .MouseCol = .ColIndex("����") Then .ToolTipText = "": Exit Sub
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
    Dim strInput As String, str���㷽ʽ As String
    
    With vsBlance
        If Row <= 0 Then Exit Sub
        
        Select Case Col
        Case .ColIndex("����")
            If CheckIsAllowBackCash(Row) = False Then Cancel = True: Exit Sub
        Case .ColIndex("�˿���")
            If Trim(.EditText) = "" Then .EditText = 0
            strInput = Trim(.EditText): strInput = Replace(strInput, Chr(vbKeyReturn), ""): strInput = Replace(strInput, Chr(10), "")
            If Not zlGetBalanceItemFromBalanceGrid(Row, objItem) Then Exit Sub
            str���㷽ʽ = Trim(.TextMatrix(.Row, .ColIndex("�˿ʽ")))
            If Val(Abs(strInput)) > Abs(objItem.ʣ����) Then
                MsgBox "�����""" & str���㷽ʽ & """�˿���ܳ��� " & Format(objItem.ʣ����, "0.00") & " ��", vbInformation, gstrSysName
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
    '����:���ݽ��������е����ݣ���ȡָ���е�BalanceItem����
    '���:lngRow-ָ������
    '����:objBalanceItem-
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-03-30 15:22:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String, lng�����ID As Long, lng���ѿ�ID As Long
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
        
        str���㷽ʽ = .TextMatrix(lngRow, .ColIndex("�˿ʽ"))
        If str���㷽ʽ = "" Then Exit Function
        lng�����ID = Val(.TextMatrix(lngRow, .ColIndex("�����ID")))
        lng���ѿ�ID = Val(.TextMatrix(lngRow, .ColIndex("���ѿ�ID")))
        
        If lng�����ID = 0 Then
            Set objCard = mobjThridSwap.zlGetCardFromBalanceName(str���㷽ʽ)
        Else
            Call gobjSquare.objSquareCard.zlGetCard(lng�����ID, lng���ѿ�ID <> 0, objCard)
        End If
        
        varTemp = Split(.TextMatrix(lngRow, .ColIndex("�༭״̬")) & "|", "|")
        Set objBalanceItem_Out = New clsBalanceItem
        With objBalanceItem_Out
            Set .objCard = objCard
            .��������ID = Val(vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("��������ID")))
            .������ˮ�� = vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("������ˮ��"))
            .����˵�� = vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("����˵��"))
            .������� = vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("�������"))
            .����ժҪ = vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("��ע"))
            .���� = vsBlance.Cell(flexcpData, lngRow, vsBlance.ColIndex("����"))
            .�Ƿ����� = Val(vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("�Ƿ�����"))) = 1
            .������ = Val(vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("�˿���")))
            .�Ƿ�����༭ = Val(varTemp(0)) = 1
            .�Ƿ�����ɾ�� = Val(varTemp(1)) = 1
            .������� = CStr(vsBlance.Cell(flexcpData, lngRow, vsBlance.ColIndex("�����ID")))
            .���ѿ� = lng���ѿ�ID <> 0
            .���ѿ�ID = lng���ѿ�ID
            .�����ID = lng�����ID
            .���� = ""
            .У�Ա�־ = Val(vsBlance.TextMatrix(lngRow, vsBlance.ColIndex("У�Ա�־")))
            .�������� = objCard.��������
            .�Ƿ�ת�� = IsTransfer(lng�����ID)
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

Private Function IsTransfer(ByVal lng�����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���ҽ�ƿ��Ƿ�֧��ת��
    ' ���� :lng�����ID-ҽ�ƿ����.id
    ' ���� : 2019/01/22
    ' ˵�� :
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    If lng�����ID = 0 Then Exit Function
    If mobjCards("K" & lng�����ID) Is Nothing Then Exit Function
    IsTransfer = mobjCards("K" & lng�����ID).�Ƿ�ת�ʼ�����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDelCashColIsEdit(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������Ƿ��������
    '���:lngRow-ָ������
    '����:����༭����true,���򷵻�False
    '����:���˺�
    '����:2018-08-31 11:00:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    
    If GetVsGridBoolColVal(vsBlance, lngRow, vsBlance.ColIndex("����")) = True Then CheckDelCashColIsEdit = True: Exit Function '�������֣��ٸ�Ϊ�����֣�������༭
    
    If Not zlGetBalanceItemFromBalanceGrid(lngRow, objItem) Then Exit Function
    If objItem.�Ƿ��������� = False And objItem.�Ƿ�ǿ������ = False Then Exit Function
    If objItem.�Ƿ���� Then Exit Function
    
    If objItem.�Ƿ񱣴� Then '����Ѿ������˵�,��Ҫ�����жϽ����Ƿ�ɹ��Ľ���
        'If mobjThridSwap.zlThird_IsSwapIsSucces(objItem, intSwapStatu, strErrMsg) Then Exit Function '���׳ɹ�������������
        'If intSwapStatu <> 0 Then
        '    strNotes = "ע��:" & vbCrLf & _
        '    "    " & objCard.���� & " �������ڽ����У����ܽ������ֲ���"
        '    If strErrMsg <> "" Then strNotes = strNotes & "����ϸ������Ϣ���£�" & vbCrLf & strErrMsg
        '    If strErrMsg = "" Then strNotes = strNotes & "��"
        '    MsgBox strNotes, vbInformation + vbOKOnly, gstrSysName
        '    Exit Function
        'End If
        ''��ɾ����Ȼ���ٿ��ܷ�����
        '
        Exit Function
    End If
    CheckDelCashColIsEdit = True: Exit Function
End Function

Private Function CheckIsAllowBackCash(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ���������
    '���:lngRow-ָ������
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-31 11:30:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem, intSwapStatu As Integer, strErrMsg As String, strNotes As String
    Dim str����Ա���� As String
    On Error GoTo errHandle
        
    If GetVsGridBoolColVal(vsBlance, lngRow, vsBlance.ColIndex("����")) = True Then CheckIsAllowBackCash = True: Exit Function '�������֣��ٸ�Ϊ�����֣�������༭
    If zlGetBalanceItemFromBalanceGrid(lngRow, objItem) = False Then Exit Function
    
    If objItem.�������� = 2 Then CheckIsAllowBackCash = True: Exit Function
    
    If objItem.�Ƿ񱣴� Then '����Ѿ������˵�,��Ҫ�����жϽ����Ƿ�ɹ��Ľ���
        If mobjThridSwap.zlThird_IsSwapIsSucces(objItem, intSwapStatu, strErrMsg) Then Exit Function '���׳ɹ�������������
        If intSwapStatu <> 0 Then
            strNotes = "ע��:" & vbCrLf & _
            "    " & objItem.objCard.���� & " �������ڽ����У����ܽ������ֲ���"
            If strErrMsg <> "" Then strNotes = strNotes & "����ϸ������Ϣ���£�" & vbCrLf & strErrMsg
            If strErrMsg = "" Then strNotes = strNotes & "��"
            MsgBox strNotes, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        '��ȷʧ��,��ɾ����Ȼ���ٿ��ܷ�����
        If DelDepositErrBill(objItem.���ݺ�, 2) Then Exit Function
        objItem.�Ƿ񱣴� = False
        objItem.�Ƿ���� = False
        '����Ƿ���������
        Exit Function
    End If
    
    If objItem.�Ƿ��������� Then CheckIsAllowBackCash = True: Exit Function

    If InStr(";" & mstrCardPrivs & ";", ";�����˿�ǿ������;") > 0 Then
        '�߱�ǿ������Ȩ��
        If MsgBox(objItem.objCard.���� & " ��֧�����֣����Ƿ�ǿ�����֣�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Function
        objItem.Tag = UserInfo.����
        CheckIsAllowBackCash = True
        Exit Function
    End If
    
    '�Ѿ���֤���ģ�������֤
    str����Ա���� = zlDatabase.UserIdentifyByUser(Me, "ǿ��������֤", glngSys, 1151, "�����˿�ǿ������")
    If str����Ա���� = "" Then
        MsgBox "¼��Ĳ���Ա��֤ʧ�ܻ���¼��Ĳ���Ա���߱�ǿ������Ȩ�ޣ�����ǿ�����֣�", vbInformation, gstrSysName
        Exit Function
    End If
    objItem.Tag = str����Ա����



    CheckIsAllowBackCash = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub DeletePayInfor(ByVal lngDelRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ����
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-31 11:38:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, objItem   As clsBalanceItem
    
    On Error GoTo errHandle
    With vsBlance
        If lngDelRow > .Rows - 1 Or lngDelRow < 1 Then Exit Sub
        If zlGetBalanceItemFromBalanceGrid(lngDelRow, objItem) = False Then Exit Sub
        
        If objItem.�Ƿ���� Then Exit Sub
        If objItem.�Ƿ񱣴� Then
            If mobjThridSwap.zlThird_IsCancelFromItems(objItem) = False Then Exit Sub
            '��ȷʧ��,��ɾ����Ȼ���ٿ��ܷ�����
            If DelDepositErrBill(objItem.���ݺ�, 2) = False Then Exit Sub
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
    '����:�Զ���̯����
    '���:dblMoney-�˿���
    '       blnAllMoney-�Ƿ��̯���з��ã��ֽ�+������
    '����:���˺�
    '����:2018-09-07 09:42:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblCashMoney As Double, dblTotal As Double, dblThirdDelMoney As Double
    Dim objItem As clsBalanceItem
    
    On Error GoTo errHandle
    If dblMoney < 0 And blnAllMoney Then Exit Sub
    If dblMoney < 0 Then dblMoney = 0
    dblTotal = dblMoney
    With vsBlance
        dblCashMoney = 0: dblThirdDelMoney = 0
        '�Ⱥϲ���������
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("�˿ʽ")) <> "" Then
                If zlGetBalanceItemFromBalanceGrid(i, objItem) Then
                   If GetVsGridBoolColVal(vsBlance, i, .ColIndex("����")) Or blnAllMoney Then
                        If objItem.������ < 0 Then dblTotal = roundEx(dblTotal - objItem.������, 5)
                   End If
                End If
            End If
        Next
        
        '�ٷ�̯���
        dblThirdDelMoney = 0: dblCashMoney = 0
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("�˿ʽ")) <> "" Then
                If zlGetBalanceItemFromBalanceGrid(i, objItem) Then
                   If GetVsGridBoolColVal(vsBlance, i, .ColIndex("����")) Then
                        If objItem.ʣ���� > 0 Then
                            If dblTotal > objItem.ʣ���� Then
                                dblTotal = roundEx(dblTotal - objItem.ʣ����, 5)
                                objItem.������ = objItem.ʣ����
                            Else
                                objItem.������ = dblTotal
                                dblTotal = 0
                            End If
                            .RowData(i) = objItem
                            .TextMatrix(i, .ColIndex("�˿���")) = Format(objItem.������, "####0.00" & IIf(objItem.�������� = 9, "####", ""))
                        End If
                        dblCashMoney = roundEx(dblCashMoney + objItem.������, 5)
                   Else
                        If objItem.ʣ���� > 0 And blnAllMoney Then
                            If dblTotal > objItem.ʣ���� Then
                                dblTotal = roundEx(dblTotal - objItem.ʣ����, 5)
                                objItem.������ = objItem.ʣ����
                            Else
                                objItem.������ = dblTotal
                                dblTotal = 0
                            End If
                            .RowData(i) = objItem
                            .TextMatrix(i, .ColIndex("�˿���")) = Format(objItem.������, "####0.00" & IIf(objItem.�������� = 9, "####", ""))
                        End If
                        dblThirdDelMoney = roundEx(dblThirdDelMoney + objItem.������, 5)
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
      If MsgBox("��ǰ����Ľ��δ��̯��ɣ��Ƿ��Է�̯�����ֽ��Ϊ�����˿���?" & vbCrLf & vbCrLf & _
               "��ǰ���룺" & Format(dblMoney, "0.00") & vbCrLf & _
               "��̯���֣�" & Format(dblCashMoney, "0.00") & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
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
    '���ܣ���ʼ����ҳ�ؼ�
    Dim objItem As TabControlItem
    
    Err = 0: On Error GoTo Errhand:
    With tbPage
        picDeposit.BorderStyle = 0
        picDepositBack.BorderStyle = 0
        picDepositHistory.BorderStyle = 0
        picDeposit.BackColor = &H8000000F
        picDepositBack.BackColor = &H8000000F
        picDepositHistory.BackColor = &H8000000F
        
        Set objItem = .InsertItem(pg_Page.pg_Ԥ������˿�, "�˿��б�", picDepositBack.hwnd, 0)
        objItem.Tag = pg_Page.pg_Ԥ������˿�
        Set objItem = .InsertItem(pg_Page.pg_Ԥ����ʷ��¼, "��ʷ��¼", picDepositHistory.hwnd, 0)
        objItem.Tag = pg_Page.pg_Ԥ����ʷ��¼
        
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
    '����:��ʾ��ʷ��Ԥ������
    '����:���˺�
    '����:2011-09-16 10:17:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, int���� As Integer, lngRow As Long, strWhere As String
    Dim rsMoney As ADODB.Recordset
    Dim lng����id As Long, i As Integer
    
    If mpatiInfo.����ID = 0 Then
        lng����id = mlng����ID
    Else
        lng����id = mpatiInfo.����ID
    End If
    
    If cboType.ListIndex < 0 Then
         int���� = 1
    Else
        int���� = cboType.ItemData(cboType.ListIndex)
    End If
    
    On Error GoTo errHandle
    '84217,���ϴ�,2015/4/22,��ʾָ����סԺ�ڼ���ɵ�Ԥ��
    If cboType.Text = "סԺԤ��" And cboPatiPage.ListIndex > 0 Then
        strWhere = " And A.��ҳID= " & cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    
    If gbln��Ժ����ʾ Then
        strWhere = strWhere & _
                " And Exists (Select 1 From ��Ա�� C, ������Ա D, ���ű� E " & _
                " Where C.���� =A.����Ա���� And C.Id = D.��Աid And D.����id = E.Id And (E.վ�� = '" & gstrNodeNo & "' Or E.վ�� Is Null))"
    End If
            
    If gblnShowHave Then
        'ֻ��ʾ��ʣ�����ʷ�ɿ�
        '���Ӳ�������������һ�ν���ʱ��һ��һ��
        strSQL = _
        "   Select NO,Sum(Nvl(A.���,0)) as ���  " & _
        "    From ����Ԥ����¼ A" & _
        "   Where A.����ID Is Null And Nvl(A.���, 0)<>0 And A.����ID=[1] And A.Ԥ�����=[2] " & _
        "   Group by NO " & _
        "   Having Sum(Nvl(A.���,0))<>0"
        
        strSQL = _
        " Select LTrim(To_Char(A.�տ�ʱ��,'YYYY-MM-DD')) as ����,A.NO as ���ݺ�,A.ʵ��Ʊ�� as Ʊ�ݺ�," & _
        "           C.���� as ����,Ltrim(To_Char(Nvl(A.���,0),'9,999,999,990.00')) as �ɿ���,A.���㷽ʽ as ����,A.����Ա���� as �տ���" & _
        " From ����Ԥ����¼ A,(" & strSQL & ") B,���ű� C" & _
        " Where A.����ID Is Null And A.Ԥ�����=[2]  And Nvl(A.���,0)<>0 And A.����ID=C.ID(+)" & _
        "       And A.���㷽ʽ Not IN(Select ���� From ���㷽ʽ Where ����=5)" & _
        "       And A.NO=B.NO And A.����ID=[1] And Not Exists (Select 1 From ����Ԥ����¼ Where No = a.No And Nvl(У�Ա�־, 0) <> 0 And ����ID=[1]) " & strWhere & _
        " Union All" & _
        " Select Min(LTrim(To_Char(A.�տ�ʱ��,'YYYY-MM-DD'))) as ����,A.NO as ���ݺ�,Max(A.ʵ��Ʊ��) as Ʊ�ݺ�," & _
        "           B.���� as ����,Ltrim(To_Char(Sum(Nvl(A.���,0)-Nvl(A.��Ԥ��,0)),'9,999,999,990.00')) as �ɿ���,A.���㷽ʽ as ����,A.����Ա���� as �տ���" & _
        " From ����Ԥ����¼ A,���ű� B" & _
        " Where A.��¼���� IN(1,11) And A.����ID is Not NULL And A.����ID=B.ID(+) And A.Ԥ�����=[2] " & _
        "       And Nvl(A.���,0)<>Nvl(A.��Ԥ��,0) And A.����ID=[1] And Not Exists (Select 1 From ����Ԥ����¼ Where No = a.No And Nvl(У�Ա�־, 0) <> 0 And ����ID=[1]) " & strWhere & _
        " Having Sum(Nvl(A.���,0)-Nvl(A.��Ԥ��,0))<>0" & _
        " Group by A.NO,B.����,A.���㷽ʽ,A.����Ա����" & _
        " Order by ����,���ݺ�,����"
    Else
        '������ʷ�ɿ���ϸ�嵥
        strSQL = _
        " Select Ltrim(To_Char(A.�տ�ʱ��,'YYYY-MM-DD')) as ����,A.NO as ���ݺ�,A.ʵ��Ʊ�� as Ʊ�ݺ�,B.���� as ����, " & _
        " Ltrim(To_Char(A.���,'9,999,999,990.00')) as �ɿ���,A.���㷽ʽ as ����,A.����Ա���� as �տ��� " & _
        " From ����Ԥ����¼ A,���ű� B" & _
        " Where A.����ID=B.ID(+) And A.��¼����=1 And A.����ID=[1]  And A.Ԥ�����=[2] " & _
        " And Not Exists (Select 1 From ����Ԥ����¼ Where No = a.No And Nvl(У�Ա�־, 0) <> 0 And ����ID=[1]) " & strWhere & _
        " Order by A.�տ�ʱ�� Desc"
    End If
    
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����id, int����)
    If Not rsMoney.EOF Then
        With vsDepositHistory
            If gblnShowHave Then
                .TextMatrix(0, .ColIndex("�ɿ���")) = "ʣ����"
            Else
                .TextMatrix(0, .ColIndex("�ɿ���")) = "�ɿ���"
            End If
            .Rows = rsMoney.RecordCount + 1
            For i = 1 To rsMoney.RecordCount
                .TextMatrix(i, .ColIndex("����")) = Nvl(rsMoney!����)
                .TextMatrix(i, .ColIndex("���ݺ�")) = Nvl(rsMoney!���ݺ�)
                .TextMatrix(i, .ColIndex("Ʊ�ݺ�")) = Nvl(rsMoney!Ʊ�ݺ�)
                .TextMatrix(i, .ColIndex("����")) = Nvl(rsMoney!����)
                .TextMatrix(i, .ColIndex("�ɿ���")) = Nvl(rsMoney!�ɿ���)
                .TextMatrix(i, .ColIndex("����")) = Nvl(rsMoney!����)
                .TextMatrix(i, .ColIndex("�տ���")) = Nvl(rsMoney!�տ���)
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
    '����:����˿����Ƿ�����༭
    '���:lngRow-ָ������
    '����:����༭����true,���򷵻�False
    '����:���˺�
    '����:2018-08-31 11:00:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    If Not zlGetBalanceItemFromBalanceGrid(lngRow, objItem) Then Exit Function
    If objItem.objCard.�Ƿ�ȫ�� Then Exit Function
    If objItem.�Ƿ���� Then Exit Function
    ChecklDelMoneyIsEdit = True: Exit Function
End Function

Private Sub vsDepositHistory_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsDepositHistory, Me.Name, "Ԥ���嵥"
End Sub

Private Sub vsDepositHistory_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsDepositHistory, Me.Name, "Ԥ���嵥"
End Sub

Private Sub LoadThirdTotal()
    '����:���������˿�����б�
    Dim str���㷽ʽ As String, strThirdMoney As String, strTmp As String
    Dim i As Integer, j As Integer, dblThird As Double
    Dim var���㷽ʽ As Variant, varData As Variant, varTmp As Variant
    
    On Error GoTo errHandle
    
    vsThirdTotal.Rows = 2: vsThirdTotal.Cell(flexcpText, 1, 0, 1, 1) = ""
    With vsBlance
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("�˿ʽ")) <> "" And Val(.TextMatrix(i, .ColIndex("�˿���"))) <> 0 Then
                If .TextMatrix(i, .ColIndex("����")) = 0 Then
                    If InStr("," & str���㷽ʽ & ",", "," & .TextMatrix(i, .ColIndex("�˿ʽ")) & ",") = 0 Then
                        str���㷽ʽ = str���㷽ʽ & "," & .TextMatrix(i, .ColIndex("�˿ʽ"))
                    End If
                    strThirdMoney = strThirdMoney & "|" & .TextMatrix(i, .ColIndex("�˿ʽ")) & "," & Val(.TextMatrix(i, .ColIndex("�˿���")))
                End If
            End If
        Next
        
        str���㷽ʽ = Mid(str���㷽ʽ, 2)
        strThirdMoney = Mid(strThirdMoney, 2)
        If str���㷽ʽ = "" Or strThirdMoney = "" Then Exit Sub
        var���㷽ʽ = Split(str���㷽ʽ, ",")
        varData = Split(strThirdMoney, "|")
        For i = 0 To UBound(var���㷽ʽ)
            dblThird = 0
            For j = 0 To UBound(varData)
                varTmp = Split(varData(j), ",")
                If var���㷽ʽ(i) = varTmp(0) Then
                    dblThird = dblThird + Val(varTmp(1))
                End If
            Next
            strTmp = strTmp & "|" & var���㷽ʽ(i) & "," & dblThird
        Next
        
        strTmp = Mid(strTmp, 2)
        If strTmp = "" Then Exit Sub
    End With
    
    With vsThirdTotal
        varData = Split(strTmp, "|")
        .Rows = UBound(varData) + 2
        For i = 1 To UBound(varData) + 1
            varTmp = Split(varData(i - 1), ",")
            .TextMatrix(i, .ColIndex("�˿ʽ")) = varTmp(0)
            .TextMatrix(i, .ColIndex("�˿���")) = Format(varTmp(1), "0.00")
        Next
        .ColWidth(.ColIndex("�˿���")) = IIf(.Rows * .RowHeight(0) <= .Height, 1855, 1620)
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
    zl_vsGrid_Para_Save mlngModul, vsThirdTotal, Me.Name, "�����˿����"
End Sub

Private Sub vsThirdTotal_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsThirdTotal, Me.Name, "�����˿����"
End Sub

Private Sub vsThirdTotal_GotFocus()
     vsThirdTotal.BackColorSel = &HFFEBD7
End Sub

Private Sub vsThirdTotal_LostFocus()
     vsThirdTotal.BackColorSel = &HE0E0E0
End Sub

Private Sub InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
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
        .Options.UseSplitterTracker = True 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetPatiInfo(ByVal lng����id As Long, ByVal lng��ҳID As Long, ByRef patiinfo As clsPatientInfo) As Boolean
    '���ܣ����ݲ���id����ҳid��ȡ������Ϣ�Ͳ�����ҳ�е���Ϣ
    '��Σ�lng����id-����id
    '          lng��ҳid-��ҳid=-1ʱ��ʾ��ѯ���һ��סԺ����Ϣ,�����ʾ��ȡָ��סԺ��������Ϣ����ҳid=0��ʾԤ��Ժ��
    '���PatiInfo-������Ϣ�е���Ϣ
    '       ��PatiPageInfo-������ҳ�е���Ϣ
    '���أ���ȡ�ɹ�����true,���򷵻�false
    Dim PatiPageInfo As New clsPatientInfo
    Dim str����id As String, blnLastTime As Boolean
    On Error GoTo errHandle
    
    If GetPatiInforFromPatiID(lng����id, patiinfo) = False Then Exit Function
    If patiinfo.����ID = 0 Then Exit Function
    blnLastTime = lng��ҳID = -1
    If blnLastTime Then
        '��ȡ���һ��סԺ����Ϣ
        str����id = lng����id
    Else
        '��ȡָ��סԺ����סԺ����Ϣ
        str����id = lng����id & ":" & lng��ҳID
    End If
    If GetPatiPageInforByID(str����id, PatiPageInfo, blnLastTime) = False Then GetPatiInfo = True: Exit Function
    If PatiPageInfo.����ID > 0 Then
        patiinfo.��ǰ����ID = PatiPageInfo.��ǰ����ID
        patiinfo.��Ժ����ID = PatiPageInfo.��Ժ����ID
        patiinfo.ҽ�Ƹ��ʽ = IIf(Val(PatiPageInfo.��ҳID) = 0, patiinfo.ҽ�Ƹ��ʽ, PatiPageInfo.ҽ�Ƹ��ʽ)
        patiinfo.��ҳID = PatiPageInfo.��ҳID
        If patiinfo.�������� = "" Then patiinfo.�������� = PatiPageInfo.��������
        patiinfo.���� = IIf(PatiPageInfo.���� = "", patiinfo.����, PatiPageInfo.����)
        patiinfo.�Ա� = IIf(PatiPageInfo.�Ա� = "", patiinfo.�Ա�, PatiPageInfo.�Ա�)
        patiinfo.���� = PatiPageInfo.����
        patiinfo.�ѱ� = IIf(PatiPageInfo.�ѱ� = "", patiinfo.�ѱ�, PatiPageInfo.�ѱ�)
        patiinfo.�������� = IIf(PatiPageInfo.�������� = 0, patiinfo.��������, PatiPageInfo.��������)
        patiinfo.���˱�ע = IIf(PatiPageInfo.���˱�ע = "", patiinfo.���˱�ע, PatiPageInfo.���˱�ע)
        patiinfo.��Ժ����ID = PatiPageInfo.��Ժ����ID
        patiinfo.����� = PatiPageInfo.�����
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
    '����:�������ϵ��ӷ�Ʊ������˿
    '���:objPati-��ǰ������Ϣ
    '       strNos-Ԥ�����ݺţ�����ö��ŷָ�
    '����:ִ�гɹ�����true,���򷵻�False
    '����:����
    '����:2020-04-07 17:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, int���� As Integer, intԤ����� As Integer
    
    On Error GoTo errHandle
    If mobjEInvoice Is Nothing Then Exit Function
    intԤ����� = cboType.ItemData(cboType.ListIndex)
    With vsBlance
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("�˿ʽ")) <> "" And Val(.TextMatrix(i, .ColIndex("�˿���"))) <> 0 Then
                int���� = IIf(Val(.TextMatrix(i, .ColIndex("��������"))) = 3, mpatiInfo.����, 0)
                objPati.���� = int����
                strNos = strNos & "," & .TextMatrix(i, .ColIndex("���ݺ�"))
                If mobjEInvoice.zlCancelEInvoiceFromBalanceInfor(Me, objPati, .TextMatrix(i, .ColIndex("���ݺ�"))) = False Then Exit Function
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
    '����:�������ߵ��ӷ�Ʊ������˿
    '���:objPati-��ǰ������Ϣ
    '       strNos-Ԥ�����ţ�����ö��ŷָ�
    '����:ִ�гɹ�����true,���򷵻�False
    '����:����
    '����:2020-04-07 17:37:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllSwapData As Collection, strDate As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim intԤ����� As Integer, int���� As Integer
    
    On Error GoTo errHandle
    If mobjEInvoice Is Nothing Then Exit Function
    If strNos = "" Then Exit Function
    strDate = "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')"
    intԤ����� = cboType.ItemData(cboType.ListIndex)
    strSQL = "" & _
            "Select a.No, a.Id As ����id, c.ԭԤ��id, a.����id, a.Ԥ�����, -1 * a.��Ԥ�� As �������, a.���㷽ʽ, d.����" & vbNewLine & _
            "From ����Ԥ����¼ A," & vbNewLine & _
            "     (Select ����id" & vbNewLine & _
            "              From ����Ԥ����¼" & vbNewLine & _
            "              Where ����id = [1]  And ��¼���� = 1 And ���ӱ�־ = 1) B," & vbNewLine & _
            "     (Select a.No, a.Id As ԭԤ��id" & vbNewLine & _
            "       From ����Ԥ����¼ A, Table(f_Str2List([2])) B" & vbNewLine & _
            "       Where a.��¼���� = 1 And a.��¼״̬ = 1 And a.No = b.Column_Value) C, ���㷽ʽ D " & vbNewLine & _
            "Where a.����id = [1]  And a.��¼���� = 11 And a.����id = b.����id And Nvl(a.��Ԥ��, 0) > 0 And a.No = c.No And a.���㷽ʽ = d.����(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡԤ�����", objPati.����ID, strNos, intԤ�����)
    If rsTemp.EOF Then zlCreateEInvoiceBat = True: Exit Function
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            int���� = IIf(Val(Nvl(!����)) = 3, mpatiInfo.����, 0)
            If mobjEInvoice.zlIsStartEInvoice(int����, intԤ�����) And Nvl(!���㷽ʽ) <> "" Then
                Set cllSwapData = Nothing
                Call GetFact(False, int����)
                objPati.���� = int����
                If mobjEInvoice.zlGetEinvoiceSwapCollect(objPati, Nvl(!ԭԤ��ID), Nvl(!NO), Val(Nvl(!�������)), strDate, txtFact.Text, cllSwapData, Nvl(!����ID), mlng����ID) Then
                    '���ߵ���Ʊ��
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


