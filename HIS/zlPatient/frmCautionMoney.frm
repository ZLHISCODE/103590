VERSION 5.00
Object = "{CC0839AF-B32F-436B-8884-BE2BB3B4C73F}#4.1#0"; "zlIDKind.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCautionMoney 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ѻ�𵥾�"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCautionMoney.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   11910
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   75
      ScaleHeight     =   2655
      ScaleWidth      =   11775
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1050
      Width           =   11775
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   20
         X1              =   6105
         X2              =   11640
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   5100
         TabIndex        =   68
         Tag             =   "���֤�� "
         Top             =   1950
         Width           =   960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   18
         X1              =   1260
         X2              =   4845
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
         TabIndex        =   67
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
         Left            =   2640
         TabIndex        =   66
         Tag             =   "δ�ɷ��� "
         ToolTipText     =   "δ�ɿ�Ļ��۵����úϼ�"
         Top             =   1170
         Width           =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   19
         X1              =   3660
         X2              =   4845
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Label lblҽ��Ԥ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ��Ԥ�� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   5100
         TabIndex        =   65
         Tag             =   "ҽ��Ԥ�� "
         ToolTipText     =   "ҽ��Ԥ����"
         Top             =   795
         Width           =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   17
         X1              =   6105
         X2              =   7680
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblWorkUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������λ"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   5100
         TabIndex        =   63
         Tag             =   "������λ "
         Top             =   1560
         Width           =   960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   16
         X1              =   1260
         X2              =   11640
         Y1              =   2550
         Y2              =   2550
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   15
         X1              =   1260
         X2              =   2430
         Y1              =   1815
         Y2              =   1815
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   14
         X1              =   1245
         X2              =   4845
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   13
         X1              =   8895
         X2              =   11640
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   12
         X1              =   6105
         X2              =   7680
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   10
         X1              =   1260
         X2              =   7725
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   9
         X1              =   6105
         X2              =   11640
         Y1              =   1815
         Y2              =   1815
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   8
         X1              =   3660
         X2              =   4845
         Y1              =   1815
         Y2              =   1815
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   7
         X1              =   1245
         X2              =   2415
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   6
         X1              =   8895
         X2              =   11640
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   5
         X1              =   8895
         X2              =   11640
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   4
         X1              =   8895
         X2              =   11640
         Y1              =   345
         Y2              =   345
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   3
         X1              =   5505
         X2              =   7080
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   2
         X1              =   3660
         X2              =   4380
         Y1              =   345
         Y2              =   345
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   0
         X1              =   2115
         X2              =   2835
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   1
         X1              =   780
         X2              =   1500
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label lblMemo 
         AutoSize        =   -1  'True
         Caption         =   "��    ע "
         Height          =   240
         Left            =   240
         TabIndex        =   60
         Tag             =   "��    ע "
         Top             =   2310
         Width           =   1080
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ���� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7890
         TabIndex        =   59
         Tag             =   "סԺ���� "
         Top             =   405
         Width           =   1080
      End
      Begin VB.Label lblδ����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ����� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   58
         Tag             =   "δ����� "
         ToolTipText     =   "δ��˵Ļ��ۼ��˷��úϼ�"
         Top             =   1170
         Width           =   1080
      End
      Begin VB.Label lblӦ�տ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӧ �� �� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7875
         TabIndex        =   57
         Tag             =   "Ӧ �� �� "
         Top             =   1170
         Width           =   1080
      End
      Begin VB.Label lblҽ�Ƹ��ʽ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ƹ��ʽ "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7410
         TabIndex        =   56
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
         TabIndex        =   55
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
         Left            =   2640
         TabIndex        =   53
         Tag             =   "������� "
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� �� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   52
         Tag             =   "�� �� �� "
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label lbl�ѱ�ȼ� 
         AutoSize        =   -1  'True
         Caption         =   "�ѱ� "
         Height          =   240
         Left            =   4965
         TabIndex        =   51
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
         Left            =   1560
         TabIndex        =   50
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
         TabIndex        =   49
         Tag             =   "�Ա� "
         Top             =   105
         Width           =   600
      End
      Begin VB.Label lblѺ����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ѻ����� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   48
         Tag             =   "Ѻ����� "
         Top             =   795
         Width           =   1080
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   3120
         TabIndex        =   47
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
         Left            =   5100
         TabIndex        =   46
         Tag             =   "ʣ���� "
         Top             =   1170
         Width           =   1080
      End
      Begin VB.Label lbl������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ����� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7890
         TabIndex        =   45
         Tag             =   "δ����� "
         Top             =   795
         Width           =   1080
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   0
      ScaleHeight     =   1650
      ScaleWidth      =   11910
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6960
      Width           =   11910
      Begin VB.CheckBox chk����ʾ����Ѻ�� 
         Caption         =   "����ʾ����Ѻ��"
         Height          =   240
         Left            =   9360
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ѻ���嵥"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   -30
         TabIndex        =   42
         Top             =   0
         Width           =   12015
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
         Height          =   1335
         Left            =   135
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   270
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   2355
         _Version        =   393216
         ForeColor       =   -2147483641
         FixedCols       =   0
         RowHeightMin    =   250
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   11910
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   8610
      Width           =   11910
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   420
         Left            =   150
         TabIndex        =   32
         Top             =   60
         Width           =   1500
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   420
         Left            =   10335
         TabIndex        =   31
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   45
         Width           =   1500
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   420
         Left            =   8760
         TabIndex        =   27
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
      Height          =   990
      Left            =   75
      ScaleHeight     =   990
      ScaleWidth      =   11760
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   0
      Width           =   11755
      Begin VB.TextBox txtFact 
         ForeColor       =   &H00C00000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   6300
         MaxLength       =   50
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F3"
         Top             =   570
         Width           =   2370
      End
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9540
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F12"
         Top             =   570
         Width           =   1830
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11355
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F8"
         Top             =   555
         Width           =   420
      End
      Begin VB.Label lblPatientNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   345
         TabIndex        =   54
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ�ݺ�"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   5565
         TabIndex        =   33
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lblFlag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   11355
         TabIndex        =   39
         Top             =   570
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ѻ�𵥾�"
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
         TabIndex        =   43
         Top             =   45
         Width           =   1875
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8760
         TabIndex        =   38
         Top             =   630
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   34
      Top             =   9210
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmCautionMoney.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16034
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
      Height          =   3225
      Left            =   75
      ScaleHeight     =   3225
      ScaleWidth      =   11775
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3800
      Width           =   11775
      Begin VB.ComboBox cboѺ����� 
         Height          =   360
         Left            =   7995
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   3690
      End
      Begin VB.ComboBox cboPatiPage 
         Height          =   360
         Left            =   3675
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   570
         Width           =   1335
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   585
         TabIndex        =   61
         Top             =   135
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
      Begin VB.ComboBox cboType 
         Height          =   360
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   570
         Width           =   1380
      End
      Begin VB.TextBox txtMan 
         Enabled         =   0   'False
         Height          =   360
         Left            =   7980
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2760
         Width           =   3705
      End
      Begin VB.TextBox txtCode 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7995
         MaxLength       =   30
         TabIndex        =   16
         Top             =   1440
         Width           =   3690
      End
      Begin VB.TextBox txtUnit 
         Height          =   360
         Left            =   7995
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1005
         Width           =   3690
      End
      Begin VB.TextBox txt�ʺ� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7980
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1890
         Width           =   3705
      End
      Begin VB.ComboBox cboUnit 
         Height          =   360
         Left            =   7995
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   585
         Width           =   3690
      End
      Begin VB.ComboBox cboNote 
         Height          =   360
         Left            =   1230
         TabIndex        =   22
         Text            =   "cboNote"
         Top             =   2325
         Width           =   10485
      End
      Begin VB.TextBox txt������ 
         Height          =   360
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1890
         Width           =   3765
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1230
         MaxLength       =   100
         TabIndex        =   1
         ToolTipText     =   "�ȼ���F11"
         Top             =   135
         Width           =   3765
      End
      Begin VB.ComboBox cboStyle 
         Height          =   360
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1440
         Width           =   3765
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   360
         Left            =   1230
         TabIndex        =   24
         Top             =   2760
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin MSCommLib.MSComm com 
         Left            =   -330
         Top             =   2070
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSMask.MaskEdBox txtMoney 
         Height          =   360
         Left            =   1230
         TabIndex        =   10
         Top             =   1005
         Width           =   3765
         _ExtentX        =   6641
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
      Begin zlIDKind.ucQRCodePayButton btQRCodePay 
         Height          =   360
         Left            =   5010
         TabIndex        =   70
         ToolTipText     =   "ɨ�븶����ʹ�ÿ����F6�����п���֧��"
         Top             =   1425
         Visible         =   0   'False
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   635
         Appearance      =   2
         ToolTipString   =   "ɨ�븶����ʹ�ÿ����F6�����п���֧��"
      End
      Begin VB.Label lblѺ����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ѻ�����"
         Height          =   240
         Left            =   6960
         TabIndex        =   69
         Top             =   180
         Width           =   960
      End
      Begin VB.Label lblPatiPage 
         AutoSize        =   -1  'True
         Caption         =   "סԺ����"
         Height          =   240
         Left            =   2685
         TabIndex        =   5
         Top             =   615
         Width           =   960
      End
      Begin VB.Label lblRepairMoney 
         Caption         =   "������:"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   5010
         TabIndex        =   62
         Top             =   1050
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label lblѺ������ 
         AutoSize        =   -1  'True
         Caption         =   "Ѻ������"
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   630
         Width           =   960
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ɿ����"
         Height          =   240
         Left            =   6960
         TabIndex        =   7
         Top             =   645
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ʺ�"
         Height          =   240
         Left            =   7440
         TabIndex        =   19
         Top             =   1950
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   435
         TabIndex        =   17
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label lbl�ɿλ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ɿλ"
         Height          =   240
         Left            =   6960
         TabIndex        =   11
         Top             =   1065
         Width           =   960
      End
      Begin VB.Line Line1 
         X1              =   -135
         X2              =   7755
         Y1              =   -30
         Y2              =   -30
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   120
         TabIndex        =   0
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblMoney 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
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
         Left            =   675
         TabIndex        =   9
         Top             =   1065
         Width           =   510
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   240
         Left            =   6960
         TabIndex        =   15
         Top             =   1500
         Width           =   960
      End
      Begin VB.Label lblStyle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "֧����ʽ"
         Height          =   240
         Left            =   195
         TabIndex        =   13
         Top             =   1500
         Width           =   960
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ"
         Height          =   240
         Left            =   645
         TabIndex        =   21
         Top             =   2385
         Width           =   480
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�տ�ʱ��"
         Height          =   240
         Left            =   195
         TabIndex        =   23
         Top             =   2820
         Width           =   960
      End
      Begin VB.Label lblMan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�տ�Ա"
         Height          =   240
         Left            =   7200
         TabIndex        =   25
         Top             =   2820
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCautionMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'˵����
'1.�˿������ַ�ʽ,ȱʡ�ķ�ʽ���ڹ�������ָ���ĵ���ִ���˿�ܣ����������տ�״̬��ʹ���˿�ܣ���һ�ַ�ʽ
'���������տ�״̬�տ�,���տ����Ը�����ʾ�˿��ʱ(�˿���<=�������)�������ַ�ʽ����Ӱ�첡��Ԥ�����ͳ��
Private Enum InStateType
    EM_��Ѻ�� = 0
    EM_������� = 1
    EM_��Ѻ�� = 2
    EM_�쳣���� = 5
    EM_�쳣���� = 6
    EM_�쳣���� = 7
End Enum
'��ڲ���----------------------------------------------------------------------------------
Private mbytInState As Byte '0-��Ѻ��(ȱʡ,���л�����),1-�������(1),2-��Ѻ��(1);5-�쳣����,6-�����쳣����,7-�쳣�˿�����
Private mstrInNO As String 'Ҫ������˿�ĵ��ݺ�(mbytInState=1��3ʱ��Ч),�Ӳ�����Ϣ�Ǽ��е����˿�ʱΪ��
Private mblnNOMoved As Boolean '����ϸʱ��¼��ǰѡ��ĵ����Ƿ����������ݱ���,����������ʱ�������ж�
Private mblnViewCancel As Boolean '�Ƿ�����˿��(mbytInState=1ʱ��Ч)
Private mstrPrivs As String
Private mlngModul As Long
Private mblnNotClick As Boolean
Private mstrbrPassWord As String
'�������----------------------------------------------------------------------------------
Private mblnUnLoad  As Boolean '���ڿ��ƴ���ֱ���˳�
Private mdblʣ���� As Double
Private mdblԤ����� As Double
Private mdbl������� As Double
Private mlng����ID As Long, mstrCardPrivs As String
Private mstrRedFact As String
Private mstrȱʡ���㷽ʽ As String
Private mblnOK As Boolean
Private mblnδ��Ʋ���Ԥ�� As Boolean '51628
Private mblnסԺ��Ԥ����֤ As Boolean   '63113:������,2013-10-29,סԺԤ���˿���֤

'ҽ������----------------------
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
Private mobjPayMode As Collection   '���㷽ʽ
Private mlngCardTypeID  As Long
Private mstr���㷽ʽ      As String
Private mstrBrushCardNo As String

Private Type Ty_YJInfor
    lngѺ��ID As Long
    strNO As String
    lng�����ID As Long
    str���� As String
    str���� As String
    str������ˮ�� As String
    str����˵�� As String
    str������λ As String
    dbl��� As Double
    bln�˿��鿨 As Boolean
    dt�տ�ʱ�� As Date
End Type
Private mYJinfo As Ty_YJInfor
Private mFactProperty As Ty_FactProperty
Private mblnStartFactUseType As Boolean '�Ƿ����õ���ص�ʹ������
Private mrsDepositBalance As ADODB.Recordset    '��ǰ���˵�Ԥ�����
Private mbytBackMoneyType As Byte '�˿ʽ:1-��ֹ;0-��ʾ
Private mbytOracleBackType As Byte '�˿���_In;0-�����˿����Ƿ�����˲�����1-����˿���
Private mblnClearWinInfor As Boolean  '�ɿ��,�Ƿ����������Ϣ
Private mblnCheckPass As Boolean 'ˢ��ʱҪ����������,'0000000000'��λ˳���ʾ��������,�ֱ�Ϊ:1.����Һ�,2.���ﻮ��,3.�����շ�,4.�������,5.��Ժ�Ǽ�,6.סԺ����,7.���˽���,8.����Ԥ����,9.���鼼ʦվ,10.Ӱ��ҽ��վ.'
'�������������
Private mobjPlugIn As Object
Private mstrPatiOld As String
Private mstrPatiSex As String
Private mlngFactModule As Long '��Ʊ��ز���ģ���
Private mblnOptErrBill As Boolean '�շ�ģʽ�´����쳣����
Private mbln�ų�δ�ɼ�δ�� As Boolean 'ʣ����ų�δ�ɼ�δ����
Private mstrQRcode  As String    'ɨ��֧���ӿڷ��صĶ�ά�봮
Private mstr�˿����Ա As String
Private mblnCheckSwapFailed As Boolean '�쳣�ؽ�ʱ,�Ƿ��齻��ʧ����(zlSwapIsSucces)
                                                                 'True-��齻��ʧ�ܣ�False-��齻�׳ɹ�
Private mbnQRPay   As Boolean  '�Ƿ���ɨ�븶��
Private mpatiInfo As New clsPatientInfo 'zlOneCardComLib.clsPatientInfo

Public Function zlShowEdit(ByVal frmMain As Object, ByVal bytInState As Byte, _
                                          ByVal strPrivs As String, ByVal lngModule As Long, _
                                          Optional strInNo As String = "", _
                                          Optional ByVal blnViewCancel As Boolean = False, _
                                          Optional blnNOMoved As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������,���ڲ���Ѻ����Ϣ�༭��鿴
    '���:frmMain-���õ�������
    '        bytInState:0-��Ѻ��(ȱʡ,���л�����),1-�������(1),2-��Ѻ��(1)
    '        strInNo:Ҫ������˿�ĵ��ݺ�(mbytInState=1��3ʱ��Ч)
    '         blnViewCancel:�Ƿ�����˿��(mbytInState=1ʱ��Ч)
    '        blnNOMoved:����ϸʱ��¼��ǰѡ��ĵ����Ƿ����������ݱ���,����������ʱ�������ж�
    '����:
    '����:Ѻ��ֻ��һ�γɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-02-17 16:11:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    mbytInState = bytInState: mstrPrivs = strPrivs: mlngModul = lngModule
    mstrInNO = strInNo: mblnViewCancel = blnViewCancel: mblnNOMoved = blnNOMoved
    mlngFactModule = mlngModul
    mblnOK = False
    If frmMain Is Nothing Then
        frmCautionMoney.Show
    Else
        frmCautionMoney.Show 1, frmMain
    End If
    zlShowEdit = mblnOK
End Function

Private Sub btQRCodePay_zlErrShow(ByVal strErrMsg As String, ByVal lngErrNum As Long)
    Call RestorePayStyle '�ָ��ϴ�ѡ����
    If strErrMsg = "" Then Exit Sub
    MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
End Sub

Private Sub btQRCodePay_zlGetPayMoney(dblMoney As Double, strExpend As String, blnCancel As Boolean)
    Err = 0: On Error GoTo errHandle:
    
    If Not (mbytInState = EM_��Ѻ�� Or mbytInState = EM_�쳣����) Then blnCancel = True: Exit Sub

    lblStyle.Tag = cboStyle.ListIndex     '�ȼ�¼��ǰѡ���֧����ʽ
    '��λ��ָ�������
    If btQRCodePay.Tag = "" Then
        MsgBox "δ�ҵ���Ч��ɨ�븶���,����!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True
        Exit Sub
    End If
    
    '���´�������
    dblMoney = StrToNum(txtMoney.Text)

    If dblMoney <> 0 Then
        txtMoney.Text = Format(dblMoney, "0.00")
    End If
    
    If dblMoney < 0 Then
        MsgBox "��ǰΪ�˿ɨ�븶��֧���˿����!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True
        Exit Sub
    End If
    
    If dblMoney = 0 Then
        MsgBox "δ���뱾��Ӧ�ɽ�����Ҫ����ɨ�븶������������ɨ�븶!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True
        zlControl.ControlSetFocus txtMoney
        Exit Sub
    End If
    If CheckDataValied = False Then blnCancel = True:  Exit Sub
    If Not Checkδ��Ʋ���Ԥ�� Then blnCancel = True:  Exit Sub
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    blnCancel = True
End Sub

Private Sub btQRCodePay_zlQRCodePayment(ByVal lngCardTypeID As Long, ByVal strPayMentQRCode As String, ByVal strExpendXML As String, blnCancel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ɨ�븶��
    '���:lngCardTypeID-�����ID
    '     strPayMentQRCode-��ά�븶������
    '     strExpendXML-����
    '����:strExpendXML-����
    '     blnCancel-true��ʾȡ������ɨ�븶,False-��ʾ����ɨ�븶�ɹ�
    '����:���˺�
    '����:2019-03-07 11:34:19
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle

    If lngCardTypeID = 0 Or blnCancel Then
        blnCancel = True
        Call RestorePayStyle '�ָ��ϴ�ѡ���֧����ʽ
        Exit Sub
    End If

    blnCancel = False
    If LocatePayStyle(lngCardTypeID) = False Then  '��λ��ɨ�븶��ָ�������
        blnCancel = True
        MsgBox "������Чʶ��ǰɨ�븶����𣬿��ܱ�����֧�ָ�����ɨ�븶���������Ա��ϵ��", vbInformation + vbOKOnly, gstrSysName
        Call RestorePayStyle '�ָ��ϴ�ѡ���֧����ʽ
        Exit Sub
    End If
    mstrQRcode = strPayMentQRCode
    mbnQRPay = True
    Call cmdOK_Click
    mbnQRPay = False
    mstrQRcode = ""
    If Not mblnClearWinInfor Then Call RestorePayStyle  '�ָ��ϴ�ѡ���֧����ʽ

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    blnCancel = True
    Call RestorePayStyle '�ָ��ϴ�ѡ���֧����ʽ
End Sub

Private Sub cboPatiPage_Click()
    If txtPatient.Tag <> "" And mbytInState = 0 And mpatiInfo.����ID > 0 Then
        If cboPatiPage.ItemData(cboPatiPage.ListIndex) <> mpatiInfo.��ҳID Then
            Call ShowPatiPageInfo
        End If
    End If
    Call ShowHistoryPrepay("")
End Sub

Private Sub ShowPatiPageInfo()
    Dim lng��ҳID As Long
    lng��ҳID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    '���ݵڼ�����Ժ������Ϣ
    Call GetPatient(IDKind.GetfaultCard, txtPatient.Tag, False, False, txtPatient.Tag, lng��ҳID)
    If mpatiInfo.����ID > 0 Then Exit Sub
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(mpatiInfo.סԺ�� = "", "", "סԺ��:" & mpatiInfo.סԺ�� & "   ") & _
                       IIf(mpatiInfo.����� = "", "", "�����:" & mpatiInfo.�����)
    lbl�ѱ�ȼ�.Caption = lbl�ѱ�ȼ�.Tag & mpatiInfo.�ѱ�
    txtPatient.Text = mpatiInfo.����
    txtPatient.Tag = mpatiInfo.����ID
    lblSex.Caption = lblSex.Tag & mpatiInfo.�Ա�
    lblOld.Caption = lblOld.Tag & mpatiInfo.����
    lblҽ�Ƹ��ʽ.Caption = lblҽ�Ƹ��ʽ.Tag & mpatiInfo.ҽ�Ƹ��ʽ
    lbl����.Caption = lbl����.Tag & GET��������(mpatiInfo.��Ժ����ID)
    lbl����.Caption = lbl����.Tag & IIf(mpatiInfo.���� = "", "��ͥ", mpatiInfo.����)
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(mpatiInfo.��ǰ����ID = 0, mpatiInfo.��Ժ����ID, mpatiInfo.��ǰ����ID))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    Call Loadҽ��Ԥ��(mpatiInfo.����ID, lng��ҳID)
End Sub

Private Sub cboPatiPage_KeyDown(KeyCode As Integer, Shift As Integer)
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboType_Click()
    If cboType.ListIndex < 0 Then Exit Sub
    
    '88657:���ϴ���2015/9/17,�л�Ԥ������ˢ��Ԥ�����
    If mbytInState = EM_��Ѻ�� And chkCancel.Value = 0 Or mbytInState = EM_�쳣���� Then
        mlng����ID = 0
        '�����:112784,����,2017/10/13,��ȡ��ȷ��Ʊ�ݸ�ʽ
        mFactProperty = zl_GetInvoicePreperty(mlngFactModule, 21, cboType.ItemData(cboType.ListIndex))
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact
        Call ShowPremayBalance(True, 0)
        Call SetCtrlEnabled
        Call ShowHistoryPrepay("")
    ElseIf mbytInState = EM_��Ѻ�� Or chkCancel.Value = 1 Then
        mlng����ID = 0
        '�����:112784,����,2017/10/13,��ȡ��ȷ��Ʊ�ݸ�ʽ
        mFactProperty = zl_GetInvoicePreperty(mlngFactModule, 22, cboType.ItemData(cboType.ListIndex))
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact(False, True)
    End If
    
     '�����:45666
    If mbytInState = EM_��Ѻ�� And cboType.Text = "סԺѺ��" Then
        chk����ʾ����Ѻ��.Visible = True
        chk����ʾ����Ѻ��.Value = IIf(zlDatabase.GetPara("����ʾ����Ԥ��", glngSys, mlngModul, , Array(chk����ʾ����Ѻ��), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
    Else
        chk����ʾ����Ѻ��.Visible = False
    End If
    lblPatiPage.Visible = cboType.Text = "סԺѺ��": cboPatiPage.Visible = cboType.Text = "סԺѺ��"
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboѺ�����_Click()
    If mblnNotClick Then Exit Sub
    If Not (mbytInState = EM_�쳣���� Or mbytInState = EM_��Ѻ��) Then Exit Sub
    With cboѺ�����
        txtMoney.Text = ""
        If Val(.ItemData(.ListIndex)) > 0 Then
            txtMoney.Text = Format(Val(.ItemData(.ListIndex)), "###0.00;-###0.00;;")
        End If
    End With
End Sub

Private Sub cboѺ�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Function zlThirdReturnCashCheck(Optional ByRef blnChange As Boolean) As Boolean
    '����:���������ּ��
    Dim dblMoney As Double, strTKList As String
    Dim strBalanceIDs As String, strXMLExpend As String
    Dim strValue As String, bln�������� As Boolean
    Dim int����״̬ As Integer, strȱʡ���ַ�ʽ As String
    
    On Error GoTo errHandle
    If cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then zlThirdReturnCashCheck = True: Exit Function
    If mYJinfo.lng�����ID = 0 Or mYJinfo.str���� = "" Then zlThirdReturnCashCheck = True: Exit Function
    cboStyle.Enabled = False: cboStyle.Locked = True
    dblMoney = roundEx(mYJinfo.dbl���, 6)
    strBalanceIDs = "8" & "|" & mYJinfo.lngѺ��ID
    
    strTKList = strTKList & Space(8) & "<TK>" & vbCrLf
    strTKList = strTKList & Space(8) & "    <TKFS>" & mYJinfo.str���� & "</TKFS>" & vbCrLf
    strTKList = strTKList & Space(8) & "    <TKJE>" & mYJinfo.dbl��� & "</TKJE>" & vbCrLf
    strTKList = strTKList & Space(8) & "    <JYLSH>" & mYJinfo.str������ˮ�� & "</JYLSH>" & vbCrLf
    strTKList = strTKList & Space(8) & "    <JYSM>" & mYJinfo.str����˵�� & "</JYSM>" & vbCrLf
    strTKList = strTKList & Space(8) & "    <KH>" & mYJinfo.str���� & "</KH>" & vbCrLf
    strTKList = strTKList & Space(8) & "</TK>" & vbCrLf
        
    strXMLExpend = "<INPUT>" & vbCrLf
    strXMLExpend = strXMLExpend & "    <TKLIST>" & vbCrLf
    strXMLExpend = strXMLExpend & strTKList
    strXMLExpend = strXMLExpend & "    </TKLIST>" & vbCrLf
    strXMLExpend = strXMLExpend & "</INPUT>"

    bln�������� = gobjSquare.objSquareCard.zlReturnCashCheck(Me, mlngModul, mYJinfo.lng�����ID, mYJinfo.str����, _
                          strBalanceIDs, dblMoney, mYJinfo.str������ˮ��, mYJinfo.str����˵��, strXMLExpend)
    If zlXML_Init() Then
        If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) Then
            Call zlXML_GetNodeValue("TXZT", , strValue): int����״̬ = Val(strValue)
            Call zlXML_GetNodeValue("QSTKFS", , strValue): strȱʡ���ַ�ʽ = Nvl(strValue)
        End If
    End If
    '�ӿڷ���ΪTrue-��������.
    If bln�������� Then
        blnChange = True
        Call Load֧����ʽ(True) '��������Ϊ1,2�Ľ��㷽ʽ
        If int����״̬ = 1 Then  'ȱʡ����
            Call LoadOriginReturnMoneyStyle(True) '����ԭʼ�˿ʽ
            cboStyle.ListIndex = cbo.FindIndex(cboStyle, strȱʡ���ַ�ʽ, True)
        Else                               '��������
            Call LoadOriginReturnMoneyStyle '����ԭʼ�˿ʽ
        End If
        zlThirdReturnCashCheck = True: Exit Function
    End If
    
    '�ӿڷ���ΪFalse-����ͨ����ǿ�����֡�Ȩ��������.
    If int����״̬ = 1 Then    '����ǿ������
        '��ǿ������Ȩ��
        If InStr(";" & mstrCardPrivs & ";", ";�����˿�ǿ������;") > 0 Then
            blnChange = True
            Call Load֧����ʽ(True)                 '��������Ϊ1,2�Ľ��㷽ʽ
            Call LoadOriginReturnMoneyStyle '����ԭʼ�˿ʽ
            zlThirdReturnCashCheck = True: Exit Function
        End If
        
        'û��ǿ������Ȩ��
        mstr�˿����Ա = zlDatabase.UserIdentifyByUser(Me, "ǿ��������֤", glngSys, 1151, "�����˿�ǿ������")
        If mstr�˿����Ա = "" Then
            MsgBox "¼��Ĳ���Ա��֤ʧ�ܻ���¼��Ĳ���Ա���߱�ǿ������Ȩ�ޣ�����ǿ�����֣� " & vbCrLf & _
                         "���Ҫǿ�����֣����������߱���ǿ�����֡��Ĳ���Ա������", vbInformation, gstrSysName
        Else
            blnChange = True
            Call Load֧����ʽ(True)                 '��������Ϊ1,2�Ľ��㷽ʽ
            Call LoadOriginReturnMoneyStyle '����ԭʼ�˿ʽ
        End If
    End If
    zlThirdReturnCashCheck = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub chk����ʾ����Ѻ��_Click()
    Call ShowHistoryPrepay("")
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
    zlControl.ControlSetFocus txtPatient
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
    If mpatiInfo.����ID = 0 Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
    chk����ʾ����Ѻ��.Enabled = cmdOK.Enabled
End Sub

Private Sub SetCtrlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Enabled����
    '����:���˺�
    '����:2011-07-24 09:30:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean, objCtl As Control
    Dim int���� As Integer
    blnEdit = mbytInState <> EM_�������
    If cboStyle.ListIndex >= 0 Then int���� = cboStyle.ItemData(cboStyle.ListIndex)
    Select Case mbytInState
    Case EM_��Ѻ��
        If chkCancel.Value = Checked Then GoTo goEnd:
        blnEdit = True
        cboѺ�����.Enabled = blnEdit
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
     Case EM_�쳣����  '�쳣����
        blnEdit = mblnCheckSwapFailed
        cboѺ�����.Enabled = False
        cboType.Enabled = False
        cboUnit.Enabled = blnEdit
        txtUnit.Enabled = blnEdit And int���� = 2
        cboStyle.Enabled = blnEdit
        txtCode.Enabled = blnEdit And int���� = 2
        txt������.Enabled = blnEdit And int���� = 2
        txt�ʺ�.Enabled = blnEdit And int���� = 2
        cboNote.Enabled = blnEdit
        cboPatiPage.Enabled = False
        txtPatient.Enabled = False
        txtMoney.Enabled = blnEdit
    Case Else
        If cboStyle.ListIndex < 0 Then GoTo goEnd:
        Select Case cboStyle.ItemData(cboStyle.ListIndex)
        Case 3 '�����ӿ�
            txtUnit.Enabled = False: txt������.Enabled = False
            txt�ʺ�.Enabled = False
        Case 1 '�ֽ�
            '�ֽ�
            txtUnit.Enabled = False: txt������.Enabled = False
            txt�ʺ�.Enabled = False: txtCode.Enabled = False
        Case 2
            blnEdit = cboStyle.Text Like "*Ʊ*" Or cboStyle.Text Like "*��*"
            txtCode.Enabled = blnEdit
            txtUnit.Enabled = True: txt������.Enabled = True: txt�ʺ�.Enabled = True
        Case Else
            txtUnit.Enabled = True: txt������.Enabled = True: txt�ʺ�.Enabled = True
        End Select
    End Select
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
    Dim strStyle As String
    If mbytInState = EM_��Ѻ�� Or chkCancel.Value = 1 Then Exit Sub
    
    If cboStyle.ListIndex = -1 Then Exit Sub
        
    '�����:111657,����,2017/07/25,ʹ���ֽ�֧��Ԥ����ʱ,�λ������������
    mstrBrushCardNo = ""     '�����������ʱ����Ŀ���
    mYJinfo.lngѺ��ID = 0
    strStyle = "_" & cboStyle.List(cboStyle.ListIndex)
  
   If Not mobjPayMode Is Nothing Then
        If CollectionExitsValue(mobjPayMode, strStyle) Then
            mlngCardTypeID = mobjPayMode(strStyle).�ӿ����
            mstr���㷽ʽ = mobjPayMode(strStyle).���㷽ʽ
        End If
        Call ShowPremayBalance(False, 0)
    End If
    Call SetCtrlEnabled
    Select Case cboStyle.ItemData(cboStyle.ListIndex)
    Case 1
        txtUnit.Text = "": txt������.Text = "": txt�ʺ�.Text = "": txtCode.Text = ""
    Case 2
        If cboStyle.Text Like "*Ʊ*" Or cboStyle.Text Like "*��*" Then
            '��֧Ʊ���ֽ�������,����������
            '����:36611
            If mpatiInfo.����ID = 0 Then Exit Sub
            strInfo = GetLastInfo(mpatiInfo.����ID)
            If strInfo <> "" Then
                txtUnit.Text = IIf(Split(strInfo, "|")(0) = "", txtUnit.Text, Split(strInfo, "|")(0))
                txt������.Text = IIf(Split(strInfo, "|")(1) = "", txt������.Text, Split(strInfo, "|")(1))
                txt�ʺ�.Text = IIf(Split(strInfo, "|")(2) = "", txt�ʺ�.Text, Split(strInfo, "|")(2))
                txtCode.Text = IIf(Split(strInfo, "|")(3) = "", txtCode.Text, Split(strInfo, "|")(3))
            End If
        End If
    Case -1
        If CheckParaConfig(mlngCardTypeID) = False Then
            mblnUnLoad = mbytInState = EM_�쳣����: Exit Sub
        End If
        If CCur(StrToNum(txtMoney.Text)) < 0 And mbytInState = EM_��Ѻ�� Then
            MsgBox "���������������븺����", vbInformation, gstrSysName
            txtMoney.Text = "": zlControl.ControlSetFocus txtMoney
        End If
    End Select
End Sub

Private Function CheckParaConfig(ByVal lngCardTypeID As Long) As Boolean
    Dim i As Integer
    If mlngCardTypeID = 0 Then
        CheckParaConfig = Not mbnQRPay: Exit Function
    End If
    If ZlGetParaConfig(lngCardTypeID, 6) = False Then
        MsgBox "����Ԥ�������е�Ѻ�𲿷ֲ�֧��ʹ��" & mstr���㷽ʽ & "���нɿ��ʹ���������㷽ʽ�ɿ" & vbCrLf & _
                     "������" & mstr���㷽ʽ & "���нɿ����ϵ�ӿڹ���Ա���������ӿڡ�", vbInformation, gstrSysName
        If mbytInState = EM_�쳣���� And Not mblnCheckSwapFailed Then Exit Function
        With cboStyle
            For i = 0 To .ListCount - 1
                If .ItemData(i) = 1 Then .ListIndex = i
            Next
            If .ItemData(.ListIndex) = -1 Then txtMoney = ""
        End With
        Exit Function
    End If
    CheckParaConfig = True
End Function
    
Private Sub cboStyle_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then
        If cboStyle.ListIndex = -1 Then
            Beep
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        If cboStyle.Locked Then Exit Sub
        If KeyAscii >= 32 Then
            lngIdx = cbo.MatchIndex(cboStyle.hwnd, KeyAscii)
            If lngIdx = -1 And cboStyle.ListCount > 0 Then lngIdx = 0
            cboStyle.ListIndex = lngIdx
        End If
    End If
End Sub

Private Sub cboStyle_Validate(Cancel As Boolean)
    If cboStyle.Locked Then Exit Sub
    If Not (cboStyle.ListIndex > -1 And mbytInState = EM_��Ѻ��) Then Exit Sub
    If mbytInState = EM_��Ѻ�� Then
         If InStr(1, mstrPrivs, ";Ѻ���տ�;") = 0 Then
             MsgBox "��û��Ȩ�޽���Ѻ���տ������", vbInformation, gstrSysName
         End If
     Else
         If InStr(1, mstrPrivs, ";Ѻ���˿�;") = 0 Then
             MsgBox "��û��Ȩ�޽���Ѻ���˿������", vbInformation, gstrSysName
         End If
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

Private Sub chkCancel_Click()
    Dim ctlTmp As Control
    Dim strTmp As String
    
    IDKind.Enabled = (chkCancel.Value <> Checked)
    
    If chkCancel.Value = Checked Then
        '����
        cmdOK.Enabled = True
        chkCancel.ForeColor = &HFF&
        btQRCodePay.Visible = False
        '�����ؽ��������
        Set mpatiInfo = New clsPatientInfo '���������Ϣ
        txtPatient.Text = "": txtPatient.Locked = True
        Call SetMoneyInfo(True)
                
        txtMoney.Text = "" '�������˲��ֿ�
        cboStyle.ListIndex = -1: cboStyle.Locked = True
        txtCode.Text = "": txtCode.Locked = True
        txtMan.Text = ""
        txtDate.Text = "____-__-__": txtDate.Enabled = False
        cboNote.ListIndex = cboNote.ListCount - 1
                
        picFace.Enabled = True '���������������˿����
        For Each ctlTmp In Me.Controls
           If ctlTmp.Name <> "com" Then
                If ctlTmp.Container.Name = "picFace" Then
                     If InStr(1, "cboNote,lblNote,txtMan,txtDate", ctlTmp.Name) <= 0 Then
                         strTmp = UCase(TypeName(ctlTmp))
                         If strTmp <> "LABEL" And strTmp <> "LINE" Then
                             On Error Resume Next     'MASKEDBOX��֧��locked����
                             ctlTmp.Enabled = False
                             If strTmp <> "MASKEDBOX" Then ctlTmp.Locked = True    '������locked����Ϊreadbill��cboStyle����listindexʱ����Click�����enabled����Ϊtrue
                             On Error GoTo 0
                         End If
                     End If
                End If
           End If
        Next ctlTmp
                
        '�������˿�ĵ��ݺ�
        cboNO.Text = "": cboNO.Tag = ""
        cboNO.Locked = False
        txtFact.Text = ""
        txtFact.Locked = True
        zlControl.ControlSetFocus cboNO
    Else
        '����
        chkCancel.ForeColor = 0
        btQRCodePay.Visible = btQRCodePay.Tag <> ""
        picFace.Enabled = True
        Call Load֧����ʽ
        For Each ctlTmp In Me.Controls
           If ctlTmp.Name <> "com" Then
                If ctlTmp.Container.Name = "picFace" Then
                     If InStr(1, "cboNote,lblNote,txtMan,txtDate", ctlTmp.Name) <= 0 Then
                         strTmp = UCase(TypeName(ctlTmp))
                         If strTmp <> "LABEL" And strTmp <> "LINE" Then
                             On Error Resume Next       'MASKEDBOX��֧��locked����
                             ctlTmp.Enabled = True
                             If strTmp <> "MASKEDBOX" Then ctlTmp.Locked = False
                             On Error GoTo 0
                         End If
                     End If
                End If
           End If
        Next ctlTmp
        
        Call ClearBill
    End If
    Call SetCtrlEnabled
End Sub

Private Sub cmdCancel_Click()
    If Not mblnOK Then Unload Me: Exit Sub
    If mbytInState = EM_��Ѻ�� Then
        If chkCancel.Value = Checked Then
            If MsgBox("ȷʵҪ�����˿��˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If mpatiInfo.����ID > 0 Then
                If MsgBox("�ò��˵�Ѻ�����δ��ȡ,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
    End If
    Unload Me
End Sub
Private Sub zlBackDepositYJ()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���Ѻ�����
    '���ƣ����˺�
    '���ڣ�2010-06-18 16:34:59
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim blnCanDel As Boolean, intInsure As Integer
    Dim bln��ӡ As Boolean, strSQL As String
    Dim msgBoxResult As String, strErrMsg As String
    Dim dblѺ���� As Double, rsTmp As New ADODB.Recordset
    Dim blnCancel As Boolean
    
    mbytOracleBackType = 1
    '�˿�
    If cboNO.Tag = "" Then
        MsgBox "�õ���δ��ȷ��ȡ,�����˿", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    '�����
    If txtMoney.Text = "" Then
        MsgBox "�˿����Ϊ��,�����룡", vbExclamation, gstrSysName
        Exit Sub
    ElseIf CCur(StrToNum(txtMoney.Text)) = 0 Then
        MsgBox "�˿����Ϊ��,�����룡", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    '���㷽ʽ���
    If cboStyle.ListIndex = -1 Then
        MsgBox "��ȷ�����㷽ʽ��", vbExclamation, gstrSysName
        zlControl.ControlSetFocus cboStyle: Exit Sub
    End If
    '��鵥���Ƿ����ˣ���Ϊ�쳣����
    If Not CheckBackErrBill(cboNO.Text, strErrMsg) Then
        MsgBox strErrMsg, vbExclamation, gstrSysName
        Exit Sub
    End If

    '����27363 by lesfeng 2010-01-13
    If MsgBox("ȷʵҪ������ " & cboNO.Text & " ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Select Case mFactProperty.intInvoicePrint
        Case 0 '����ӡԤ����Ʊ
           bln��ӡ = False
        Case 1 '�Զ���ӡ
           bln��ӡ = True
        Case 2 '��ӡ����
            msgBoxResult = MsgBox("�Ƿ���Ҫ��ӡѺ���Ʊ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
            bln��ӡ = (msgBoxResult = vbYes)
        End Select
    
        If mYJinfo.lng�����ID = 0 Then
            If gbytԤ��������鿨 <> 0 Then
                If mblnסԺ��Ԥ����֤ Or cboType.ItemData(cboType.ListIndex) = 1 Then
                    If CreatePublicExpense() Then
                        If Not gobjPublicExpense.zlPatiIdentify(mlngModul, Me, Val(txtPatient.Tag), Val(StrToNum(txtMoney.Text)), False) Then Exit Sub
                    End If
                End If
            End If
        End If
        'ҽ����ؼ��
        blnCanDel = True 'ȱʡΪ֧��,���ǹ��̵�һ�㻯����
        intInsure = ExistInsure(cboNO.Text)
        If intInsure > 0 Then
            'ȥ����ҽ������ƥ����
            blnCanDel = gclsInsure.GetCapability(supportԤ���˸����ʻ�, Val(txtPatient.Tag), intInsure)
        End If
        strSQL = "Select Nvl(���, 0) as Ѻ���� From ����Ѻ���¼ Where NO = [1] And ��¼״̬=1 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡѺ�����", cboNO.Text)
        If rsTmp.EOF Then
            MsgBox "û�з���Ҫ�˿��Ѻ���¼,�õ��ݿ����Ѿ����ˣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        dblѺ���� = Val(rsTmp!Ѻ����)
        If CCur(StrToNum(txtMoney.Text)) > dblѺ���� Then
            If mbytBackMoneyType = 1 Then
                Call MsgBox("�ñ�Ѻ����˿����Ѻ����࣬�㲻���������ŵ��ݣ�", vbInformation + vbOKOnly, gstrSysName)
                Exit Sub
            Else
                If MsgBox("�ñ�Ѻ����˿����Ѻ����࣬������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                mbytOracleBackType = 0
            End If
        End If
        
        cmdOK.Enabled = False   '��ҽ����ʱ
        
        '��������ӿڽ����Ƿ�Ϸ�
        '108666:���ϴ���2017/5/9���ָ�ȷ�ϰ�ť����״̬
        If zlCheckDepositDelValied(Val(cboNO.Tag), StrToNum(txtMoney.Text)) = False Then cmdOK.Enabled = True: Exit Sub
        
        'ִ�����ϲ���
        If Not CancelBill(CLng(cboNO.Tag), cboNO.Text, blnCanDel, intInsure, bln��ӡ, cboNote.Text) Then '�˿�
'            MsgBox "����ʧ��,�����Ըò���������������,����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            cmdOK.Enabled = True
            Exit Sub
        End If
        
        If bln��ӡ Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_3", Me, "NO=" & cboNO.Text, 2)
            Call zlCheckFactIsEnough
        End If
        
        cmdOK.Enabled = True
        
        'ҽ���Ķ�
        For i = 0 To cboStyle.ListCount - 1
            If cboStyle.ItemData(i) = 3 Then
                cboStyle.RemoveItem i: Exit For
            End If
        Next
     Else
        blnCancel = True
    End If
    If mbytInState <> EM_��Ѻ�� Then
        chkCancel.Value = Unchecked '(�������¼�)
    Else
        mblnOK = Not blnCancel
        Unload Me: Exit Sub '�˿�ģʽ�������˳�
    End If
    mblnOK = Not blnCancel
    Call ClearBill
End Sub

Private Function CheckBackErrBill(ByVal strNO As String, ByRef strErrMsg As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:�˿�ʱ��鵥���Ƿ�����
    '���:
    '����:
    '����:2018-07-20
    '˵��:
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select У�Ա�־ From ����Ѻ���¼ Where  ��¼״̬=3 And No=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If Not rsTmp.EOF Then
        strErrMsg = "����[" & strNO & "]���˿�����ظ�������"
        If Nvl(rsTmp!У�Ա�־, 0) <> 0 Then
            strErrMsg = "����[" & strNO & "]������Ϊ�쳣�˿�ݣ����˳����¶�ȡ��"
        End If
        Exit Function
    End If
    CheckBackErrBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckDataValied(Optional ByVal bln��ӡ As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���������Ƿ�Ϸ�
    '���أ��Ϸ�����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-18 16:38:39
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
   '�µ�����
  If mpatiInfo.����ID = 0 Then
        MsgBox "û��ȷ����ȡѺ��Ĳ���,���ܽ���Ѻ���ֵ��", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtPatient: Exit Function
  End If
          
  If LenB(StrConv(txtUnit.Text, vbFromUnicode)) > 50 Then
      MsgBox "�ɿλ����ֻ���� 50 ���ַ��� 25 ������,���޸ģ�", vbInformation, App.Title
      zlControl.ControlSetFocus txtUnit: Exit Function
  End If
  If LenB(StrConv(txt������.Text, vbFromUnicode)) > 50 Then
      MsgBox "����������ֻ���� 50 ���ַ��� 25 ������,���޸ģ�", vbInformation, App.Title
      zlControl.ControlSetFocus txt������: Exit Function
  End If
  If LenB(StrConv(cboNote.Text, vbFromUnicode)) > 50 Then
      MsgBox "�ɿ�ժҪֻ���� 50 ���ַ��� 25 ������,���޸ģ�", vbInformation, App.Title
      zlControl.ControlSetFocus cboNote: Exit Function
  End If
  If CheckParaConfig(mlngCardTypeID) = False Then Exit Function
  If mbytInState = EM_��Ѻ�� Then
    If cboType.ListIndex < 0 Then Exit Function
    '����:44963
    If mpatiInfo Is Nothing Then Exit Function
    If mpatiInfo.����ID = 0 Then Exit Function
    If cboType.ItemData(cboType.ListIndex) = 2 Then
        If Not mpatiInfo.��Ժ And gblnAllowOut = False Then
            If Not (mpatiInfo.�������� = 0 And mpatiInfo.��ҳID = 0 And mpatiInfo.סԺ״̬ = 0) Then
                MsgBox "���˻�δסԺ,���ܽ�סԺѺ��,����!", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Else
        If mpatiInfo.��Ժ And gblnBanIn = True Then
            MsgBox "���˻�δ��Ժ,���ܽ�����Ѻ��,����!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
  End If
  '�����
  '����27363 by lesfeng 2010-01-13
    If txtMoney.Text = "" Then
      MsgBox "�տ����Ϊ��,�����룡", vbExclamation, gstrSysName
      zlControl.ControlSetFocus txtMoney: Exit Function
    ElseIf CCur(StrToNum(txtMoney.Text)) = 0 Then
      MsgBox "�տ����Ϊ��,�����룡", vbExclamation, gstrSysName
      zlControl.ControlSetFocus txtMoney: Exit Function
    End If

    mbytOracleBackType = 1

    If cboѺ�����.ListIndex = -1 Then
        MsgBox "��ȷ��Ѻ�����", vbExclamation, gstrSysName
        zlControl.ControlSetFocus cboѺ�����: Exit Function
    End If
    
    If cboStyle.ListIndex = -1 Then
        MsgBox "��ȷ�����㷽ʽ��", vbExclamation, gstrSysName
        zlControl.ControlSetFocus cboStyle: Exit Function
    End If
    
    If InStr(mstrPrivs, ";Ѻ���տ�;") = 0 Then
        MsgBox "��û��Ȩ�޽���Ѻ���տ������", vbInformation, gstrSysName
        Exit Function
    End If
  
    If bln��ӡ = False Then CheckDataValied = True: Exit Function
    If CheckInvoicePrint = False Then Exit Function

    CheckDataValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInvoicePrint() As Boolean
    '����:���Ʊ�����
    
    On Error GoTo errHandle
    If mFactProperty.intInvoicePrint = 0 Then CheckInvoicePrint = True: Exit Function
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
                    zlControl.ControlSetFocus txtFact
            End Select
            txtFact.Text = ""
            Exit Function
        End If
    Else
        If Len(txtFact.Text) <> gbytԤ�� And txtFact.Text <> "" Then
            MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytԤ�� & " λ��", vbInformation, gstrSysName
            zlControl.ControlSetFocus txtFact: Exit Function
        End If
    End If
    CheckInvoicePrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckSwapIsSucces(ByVal lngѺ��ID As Long, ByVal dblMoney As Double, ByVal intSwapType As Integer, _
                                    ByRef strErrMsg As String, ByRef intState As Integer) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:��齻���Ƿ�ɹ�
    '���:intState(0-����ʧ�ܣ�1-�������ڽ���)
    '����:
    '����:2018-06-28 15:06:20
    '˵��:
    '-------------------------------------------------------------------------------------------------------------------------
    Dim intSwapStatus_Out As Integer, strSwapExtendInfor As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϽ����Ƿ�ɹ���10.35.90��
    '���:  frmMain-���õ�������
    '       lngModule-ģ���
    '       intSwapType-0-�ۿ�;1-�˿2-ת��
    '       lngCardTypeID-�����ID
    '       strCardNO-����
    '       dblSwapMoney-���׽��
    '       strBalanceIDs-����֧�����漰�Ľ���ID ��ʽ:�շ�����|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn �շ�����: 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ�
    '       strExpend-��չ����:�˿������ʱ���Ŵ���,��ʽ���� ��
    '        <INPUT>
    '            <TKLIST>
    '                    <TK>
    '                       <JYLSH>������ˮ��</JYLSH>
    '                       <KH>����</KH>
    '                       <JE>���</JE>
    '                    </TK>
    '            </TKLIST>
    '        </INPUT>
    '����:intSwapStatus_Out-�ӿڷ���Falseʱ���˲�����Ч:����״̬: 0-���׵���ʧ��;1-�������ڴ�����
    '     strErrMsg- ���صĴ�����Ϣ:  Ϊ�գ�������ʾ,��Ϊ��ʱ��������ʾ����Ϣ
    '     strXMLExpend-���Ժ���չ
    '���أ��ӿڵ��óɹ�����true,���򷵻�Flase
    '����:2013-06-15 20:22:51
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If intSwapType = 1 Then strSwapExtendInfor = GetExpendInfo(lngѺ��ID, False, dblMoney)
    CheckSwapIsSucces = gobjSquare.objSquareCard.zlSwapIsSucces(Me, mlngModul, intSwapType, mlngCardTypeID, "8|" & lngѺ��ID, mstrBrushCardNo, _
        dblMoney, intSwapStatus_Out, strErrMsg, strSwapExtendInfor)
        
    intState = intSwapStatus_Out
End Function

Private Function GetExpendInfo(ByVal lngѺ��ID As Long, Optional ByVal blnReturn As Boolean, Optional ByVal dblMoney As Double) As String
    '-------------------------------------------------------------------------------------------------------------------------
    '����:��������ӿ�zlSwapIsSucces��չ���
    '���:  lngѺ��ID
    '       blnReturn-(true-�˿�׵����,false-����״̬������)
    '       dblMoney-�˿�׵��˿���
    '����:
    '����:2018-07-20
    '˵��:
    '       strExpend-��չ����:�˿������ʱ���Ŵ���,��ʽ���� ��
    '        <INPUT>
    '            <TKLIST>
    '                    <TK>
    '                       <JYLSH>������ˮ��</JYLSH>
    '                       <KH>����</KH>
    '                       <JE>���</JE>
    '                    </TK>
    '            </TKLIST>
    '        </INPUT>
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strExpend As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    If lngѺ��ID = 0 Then Exit Function
    strSQL = "Select No,���㷽ʽ,����˵��,Decode(��¼״̬, 2, -1 * ���, ���) As ���,����,������ˮ�� From ����Ѻ���¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngѺ��ID)
    If rsTmp.EOF Then Exit Function
    If blnReturn Then
        strExpend = "<INPUT>" & vbCrLf & _
                    "   <TKLIST>" & vbCrLf & _
                    "      <TK>" & vbCrLf & _
                    "        <JYLSH>" & Nvl(rsTmp!������ˮ��) & "</JYLSH>" & vbCrLf & _
                    "        <TKFS>" & Nvl(rsTmp!���㷽ʽ) & "</TKFS>" & vbCrLf & _
                    "        <JYSM>" & Nvl(rsTmp!����˵��) & "</JYSM>" & vbCrLf & _
                    "        <DJH>" & Nvl(rsTmp!NO) & "</DJH>" & vbCrLf & _
                    "        <TKJE>" & dblMoney & "</TKJE>" & vbCrLf & _
                    "      </TK>" & vbCrLf & _
                    "   </TKLIST>" & vbCrLf & _
                    "</INPUT>"
    Else
        strExpend = "<INPUT>" & vbCrLf & _
                    "   <TKLIST>" & vbCrLf & _
                    "     <TK>" & vbCrLf & _
                    "       <JYLSH>" & Nvl(rsTmp!������ˮ��) & "</JYLSH>" & vbCrLf & _
                    "       <KH>" & Nvl(rsTmp!����) & "</KH>" & vbCrLf & _
                    "       <JE>" & dblMoney & "</JE>" & vbCrLf & _
                    "     </TK>" & vbCrLf & _
                    "  </TKLIST>" & vbCrLf & _
                    "</INPUT>"
    End If

    GetExpendInfo = strExpend
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckBrushCard(Optional ByRef blnUnolad As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ˢ��
    '���Σ��Ƿ�رմ���
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lng����id As String, strXmlIn As String
    Dim dblMoney As Double
    Dim strExpand As String '�����:55666
    Dim dbl�˻���� As Double '�����:55666
    Dim strBrushNo As String
    
    On Error GoTo errHandle
    dblMoney = 1 * StrToNum(txtMoney.Text)
    If cboStyle.ItemData(cboStyle.ListIndex) >= 0 Then CheckBrushCard = True: Exit Function
     '����ˢ������
    'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln���ѿ� As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByVal dbl��� As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln�˷� As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln���� As Boolean = False, _
        Optional ByVal bln�����ֹ As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal blnתԤ�� As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "", _
        Optional ByVal str������Դ As String, _
        Optional ByVal lng����ID As Long) As Boolean
    '       strXmlIn-XML���,Ŀǰ��ʽ����:
    '       <IN>
    '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
    '       </IN>
    '       str������Դ - ��ǰ֧�����õķ�����Դ�������ö��ŷָ�(ʹ�����ѿ�֧��ʱ����)
    '       lng����ID - ����ID(ʹ�����ѿ�֧��ʱ����)
    '�����:55666
    If mpatiInfo.����ID > 0 Then lng����id = mpatiInfo.����ID

    strBrushNo = mstrBrushCardNo
    
    strXmlIn = "" & _
    "<IN>" & vbCrLf & _
    "   <CZLX>0</CZLX>" & vbCrLf & _
    "   <QRCODE>" & mstrQRcode & "</QRCODE>" & vbCrLf & _
    "</IN>"
    
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, mlngCardTypeID, False, _
        Nvl(mpatiInfo.����), mpatiInfo.�Ա�, mpatiInfo.����, dblMoney, mstrBrushCardNo, mstrbrPassWord, _
        False, True, False, False, Nothing, False, False, strXmlIn, _
        cboType.ItemData(cboType.ListIndex), lng����id) = False Then Exit Function
        
    '����ǰ,һЩ���ݼ��
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    
    strXmlIn = "" & _
    "<IN>" & vbCrLf & _
    "   <QRCODE>" & mstrQRcode & "</QRCODE>" & vbCrLf & _
    "   <SFYJ>" & 1 & "</SFYJ>" & vbCrLf & _
    "</IN>"
    
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, mlngCardTypeID, _
        False, mstrBrushCardNo, dblMoney, "", strXmlIn) = False Then
        If gbln���ý����첽���� Then
            'ɾ��ԭʼ����
            If mbytInState = EM_�쳣���� Then
                strSQL = GetDeleteSQL(mstrInNO)
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                MsgBox "����������ʧ�ܣ���ɾ�����쳣���ݡ�", vbInformation, gstrSysName
                blnUnolad = True
            End If
            Exit Function
        End If
    End If
    '�����:55666,55851
    gobjSquare.objSquareCard.zlGetAccountMoney Me, mlngModul, mlngCardTypeID, mstrBrushCardNo, strExpand, dbl�˻����, False
    If dbl�˻���� <> 0 Then
        sta.Panels(2).Text = "�˻����:" & Format(dbl�˻����, "0.00")
        If dbl�˻���� < dblMoney Then
            MsgBox "ע��:" & vbCrLf & _
                         "�˻����Ϊ" & Format(dbl�˻����, "0.00") & "Ԫ��С��ԭ�ɿ���" & Format(dblMoney, "0.00") & _
                         "�����νɿ�" & Format(dbl�˻����, "0.00") & "Ԫ��", vbInformation, gstrSysName
            lblMoney.Tag = dblMoney
            dblMoney = Format(dbl�˻����, "0.00")
            lblRepairMoney.Visible = True
        End If
    End If
    
    '�ж�Ԥ�����Ƿ񳬳�ˢ�������
    If lblRepairMoney.Visible Then
        lblRepairMoney.Caption = "������:" & Format((CDbl(txtMoney.Text) - dblMoney), "###0.00;-###0.00;;")
        txtMoney.Text = Format(dblMoney, "###0.00;-###0.00;;")
    End If
    CheckBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ReCancelBill(ByVal strNO As String, Optional ByVal cllStatusUpdate As Collection) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:�����쳣�˿��
    '���:
    '����:
    '����:2018-07-03
    '˵��:
    '-------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lng����ID As Long, rsBill As ADODB.Recordset
    Dim strSQL As String, blnTrans As Boolean
    Dim msgBoxResult As VbMsgBoxResult
    Dim bln��ӡ As Boolean
    
    On Error GoTo errHandle
    strSQL = "Select Id From ����Ѻ���¼ Where ��¼״̬=2 And Nvl(У�Ա�־,0)<>0  And No=[1]"
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsBill.RecordCount = 0 Then
        MsgBox "����[" & strNO & "]Ϊ���쳣���ݣ���ˢ�º����ԣ�", vbInformation, gstrSysName
        Exit Function
    End If
    lng����ID = Nvl(rsBill!ID, 0)
    strSQL = "Select Id From ����Ѻ���¼ Where ��¼״̬=3  And No=[1]"
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsBill.EOF Then
        MsgBox "û�в��ҵ�ԭʼѺ�𵥾ݣ��޷������˿������", vbInformation, gstrSysName
        Exit Function
    End If
    lngID = Nvl(rsBill!ID, 0)
    
    Select Case mFactProperty.intInvoicePrint
    Case 0 '����ӡԤ����Ʊ
       bln��ӡ = False
    Case 1 '�Զ���ӡ
       bln��ӡ = True
    Case 2 '��ӡ����
        msgBoxResult = MsgBox("�Ƿ���Ҫ��ӡѺ���Ʊ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        bln��ӡ = (msgBoxResult = vbYes)
    End Select
    If bln��ӡ Then Call GetFact  '���»�ȡ��Ʊ��
    Set cllStatusUpdate = New Collection
    '����У�Ա�־�������˿�
    strSQL = "zl_����Ѻ���¼_DELETE(" & lngID & ",'" & cboNote.Text & "','" & _
        UserInfo.��� & "','" & UserInfo.���� & "'," & lng����ID & "," & _
        IIf(bln��ӡ, "'" & txtFact.Text & "'", "NULL") & "," & IIf(bln��ӡ, IIf(mlng����ID > 0, mlng����ID, "Null"), "Null") & ",2)"
    zlAddArray cllStatusUpdate, strSQL
    
    '���������ӿ�
    If zlDepositDel(lngID, lng����ID, StrToNum(txtMoney.Text), strNO, cllStatusUpdate, blnTrans, , True) = False Then
        Exit Function
    End If
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    If bln��ӡ Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_3", Me, "NO=" & mstrInNO, 2)
        Call zlCheckFactIsEnough
    End If
    ReCancelBill = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim msgBoxResult As VbMsgBoxResult '�����:50656
    Dim bln��ӡ As Boolean  '�����:57624
    Dim lngѺ��ID As Long, intState As Integer
    Dim strErrMsg As String, blnBeenErr As Boolean
    Dim cllStatusUpdate As Collection
    Dim blnVocherPrint As Boolean, bytP As Byte '��ӡƾ��
    Dim strSavedDate As String '�տ����ڣ����ڴ�ӡ
    Dim blnUnload As Boolean

    If chkCancel.Value = Checked Then
        If mbytInState = EM_��Ѻ�� Or mbytInState = EM_��Ѻ�� Then
            '��Ѻ��
            Call zlBackDepositYJ: Exit Sub
        ElseIf mbytInState = EM_�쳣���� Then
            'EM_�쳣����
            '�����쳣�˿��

            If CheckSwapIsSucces(Val(cboNO.Tag), StrToNum(txtMoney.Text), 1, strErrMsg, intState) = False Then
                If intState = 0 Then
                    '�˿��ʧ�ܣ��ָ�Ϊ��������
                    If Not DelDepositErrBill(mstrInNO, 1) Then
                        MsgBox "�������˿��쳣����ɾ��ʧ�ܣ����Ժ����ԣ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    MsgBox "����[" & mstrInNO & "]�������˿����ʧ�ܣ������տ��б���������˿�" & IIf(strErrMsg <> "", "��" & vbCrLf & "������Ϣ���£�" & vbCrLf, "��") & _
                    strErrMsg, vbInformation, gstrSysName
                    mblnOK = True
                Else
                    MsgBox "�����������ڽ����У��޷������˿���Ժ�����" & IIf(strErrMsg <> "", "��" & vbCrLf & "������Ϣ���£�" & vbCrLf, "��") & _
                        strErrMsg, vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                mblnOK = ReCancelBill(mstrInNO, cllStatusUpdate)
            End If
            '�����շ�ģʽ�����쳣���ݺ󣬻ص��շ�״̬
            If mblnOptErrBill Then
                Call RestoreStatue: Exit Sub
            Else
                Unload Me: Exit Sub
            End If
        End If
    End If
    
    If mbytInState <> EM_�쳣���� Then
        Select Case mFactProperty.intInvoicePrint
        Case 0 '����ӡԤ����Ʊ
           bln��ӡ = False
        Case 1 '�Զ���ӡ
           bln��ӡ = True
        Case 2 '��ӡ����
            msgBoxResult = MsgBox("�Ƿ���Ҫ��ӡѺ��Ʊ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
            bln��ӡ = (msgBoxResult = vbYes)
        End Select
    End If
    
    If (mbytInState = EM_��Ѻ�� Or mbytInState = EM_�쳣����) Then
        bytP = Val(zlDatabase.GetPara("Ѻ��ƾ����ӡ��ʽ", glngSys, mlngModul))
        Select Case bytP
        Case 0 '����ӡԤ����Ʊ
           blnVocherPrint = False
        Case 1 '�Զ���ӡ
           blnVocherPrint = True
        Case 2 '��ӡ����
            msgBoxResult = MsgBox("�Ƿ���Ҫ��ӡѺ��ƾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
            blnVocherPrint = (msgBoxResult = vbYes)
        End Select
    End If
    If mbytInState = EM_�쳣���� Then
        '�쳣�տ�ݽ�������
        If bln��ӡ Then Call GetFact  '���»�ȡ��Ʊ��
        If CheckDataValied(bln��ӡ) = False Then Exit Sub
        If mblnCheckSwapFailed Then
            If CheckBrushCard(blnUnload) = False Then
                If Not blnUnload Then Exit Sub
                'ɾ���쳣���ݺ�رմ���
                If mblnOptErrBill Then
                    Call RestoreStatue: Exit Sub
                Else
                    mblnOK = True: Unload Me: Exit Sub 'ˢ������
                End If
            End If
            mblnOK = ReDepositErrBill(mstrInNO, bln��ӡ, blnVocherPrint)
        Else
            If CheckSwapIsSucces(Val(cboNO.Tag), StrToNum(txtMoney.Text), 0, strErrMsg, intState) = False Then
                If intState = 0 Then
                    If CheckBrushCard(blnUnload) = False Then
                        If blnUnload Then 'ɾ���쳣���ݺ�رմ���
                            If mblnOptErrBill Then
                                Call RestoreStatue: Exit Sub
                            Else
                                mblnOK = True: Unload Me: Exit Sub 'ˢ������
                            End If
                        Else
                            mblnCheckSwapFailed = True
                            btQRCodePay.Visible = btQRCodePay.Tag <> ""
                            Call SetCtrlEnabled: Exit Sub
                        End If
                    End If
                    mblnOK = ReDepositErrBill(mstrInNO, bln��ӡ, blnVocherPrint)
                Else
                    MsgBox "�����������ڽ����У��޷��������գ����Ժ�����" & IIf(strErrMsg <> "", "��" & vbCrLf & "������Ϣ���£�" & vbCrLf, "��") & _
                            strErrMsg, vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                mblnOK = ReDepositErrBill(mstrInNO, bln��ӡ, blnVocherPrint)
            End If
        End If
        '�����շ�ģʽ�����쳣���ݺ󣬻ص��շ�״̬
        If mblnOptErrBill Then
            Call RestoreStatue: Exit Sub
        Else
            Unload Me: Exit Sub
        End If
    ElseIf mbytInState = EM_�쳣���� Then
        '�쳣����
        If CheckSwapIsSucces(Val(cboNO.Tag), StrToNum(txtMoney.Text), 0, strErrMsg, intState) = False Then
            If intState = 0 Then
                If DelDepositErrBill(mstrInNO) = False Then
                    MsgBox "��������ʧ�ܣ����Ժ����ԣ�", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                MsgBox "�����������ڽ����У��޷����ϵ��ݣ����Ժ�����" & IIf(strErrMsg <> "", "��" & vbCrLf & "������Ϣ���£�" & vbCrLf, "��") & _
                        strErrMsg, vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            MsgBox "���������ѳɹ������������ϵ��ݣ��������շѣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        mblnOK = True: Unload Me: Exit Sub
    End If
    
    If Not Checkδ��Ʋ���Ԥ�� Then Exit Sub
    If CheckDataValied(bln��ӡ) = False Then Exit Sub
    If CheckBrushCard = False Then Exit Sub
    '����
    cmdOK.Enabled = False
    
    '�м䲻���е����࣬���ⳤʱ�������ɲ���
    If Not SaveBill(bln��ӡ, lngѺ��ID, blnBeenErr, strSavedDate) Then
        If blnBeenErr Then
            Call SetcmdOkEnabled
            zlControl.ControlSetFocus txtPatient
        Else
            '������ʱ�����ݽӿ���Ϣ��ʾ
            If cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then
                MsgBox "Ѻ�𵥾ݱ���ʧ��,�����Ըò����������������,����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            End If
            cmdOK.Enabled = True: Exit Sub
        End If
    Else
        '�����:57624
        '�����:50656
        If bln��ӡ Then 'Ʊ�ݺ�Ϊ�վͱ�ʾ����ӡ��Ʊ
            '78751:���ϴ�,2014/10/20,����Ԥ��Ʊ�ݴ�ӡ��ʽ
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_2", Me, "NO=" & cboNO.List(0), "����ID=" & mpatiInfo.����ID, "�տ�ʱ��=" & Format(strSavedDate, "yyyy-mm-dd HH:MM:SS"), 2)
            Call zlCheckFactIsEnough
        End If
        If blnVocherPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_4", Me, "NO=" & cboNO.List(0), 2)
        End If
        '81693:���ϴ�,2015/4/21,������
        If Not mobjPlugIn Is Nothing Then
            On Error Resume Next
            Call mobjPlugIn.PatiPrePayAfter(mpatiInfo.����ID, cboType.ItemData(cboType.ListIndex), lngѺ��ID)
            Err.Clear
        End If
    End If
    '�����:55666
    '���ڲ����������
    If UBound(Split(lblRepairMoney.Caption, ":")) = 1 And Split(lblRepairMoney.Caption, ":")(1) <> "" Then
        txtPatient.Tag = ""
        lblRepairMoney.Tag = Split(lblRepairMoney.Caption, ":")(1)
        IDKind.IDKind = IDKind.GetKindIndex("����")
        txtPatient.Text = "-" & mpatiInfo.����ID
        txtPatient_KeyPress 13
        'ˢ��Ʊ��
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact
        '��λ֧����ʽ
        '84751:���ϴ�,2015/5/14,������λԽ��
        For i = 0 To cboStyle.ListCount - 1
            If cboStyle.List(i) = mstrȱʡ���㷽ʽ Then
                cboStyle.ListIndex = i
            End If
        Next
        txtMoney.Text = Format(lblRepairMoney.Tag, "0.00")
        lblRepairMoney.Tag = ""
        lblRepairMoney.Visible = False: lblRepairMoney.Caption = "������:"
        cmdOK.Enabled = True
        Exit Sub
    End If
    
    If mblnClearWinInfor Then
        Call ClearBill
        Call InitFace(True)
        Call cboStyle_Click
    Else
        '�����:44732
        SetMoneyInfo False
        Set mpatiInfo = New clsPatientInfo
        
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact  '���»�ȡ��Ʊ��
    End If
    Call SetcmdOkEnabled
    zlControl.ControlSetFocus txtPatient
    mblnOK = True
End Sub

Private Sub RestoreStatue()
    '���ܣ��ָ���Ѻ��״̬
    mbytInState = EM_��Ѻ��
    Call ClearBill: Call InitFace
    mblnOptErrBill = False
    cmdOK.Caption = "ȷ��(&O)"
    chkCancel.Value = Unchecked
    Call SetCtrlEnabled
End Sub

Private Function DelDepositErrBill(ByVal strNO As String, Optional ByVal bytOpt As Byte) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��Ѻ���쳣���ݼ�¼
    '���: strno-���ݺţ�Optype-(0-ɾ���쳣��ֵ���ݣ�1-ɾ���쳣�˿��)
    '����:
    '����:2018-06-29
    '˵��:
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle

    strSQL = GetDeleteSQL(strNO, bytOpt)
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    DelDepositErrBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ReDepositErrBill(ByVal strNO As String, ByVal blnPrintInvoice As Boolean, ByVal blnPrintVocher As Boolean) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:�쳣��������
    '���:  strNO-���ݺ�
    '       blnPrintInvoice-�Ƿ��ӡƱ�� ��Ϊtrue��ӡ
    '       blnPrintVocher-�Ƿ��ӡƾ�� ��Ϊtrue��ӡ
    '����:
    '����:2018-06-28 16:11:16
    '˵��:
    '-------------------------------------------------------------------------------------------------------------------------
    Dim rsBill As ADODB.Recordset, strSQL As String, blnTrans As Boolean
    Dim dbl��� As Double
    Dim strCurDate As String, cllStatusUpdate As Collection
    
    On Error GoTo errHandle
    
    strSQL = "Select Id,��ҳID,���,�տ�ʱ�� From ����Ѻ���¼ Where ��¼״̬=0 And Nvl(У�Ա�־,0)<>0  And No=[1]"
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsBill.RecordCount = 0 Then
        MsgBox "����[" & strNO & "]Ϊ���쳣���ݣ���ˢ�º����ԣ�", vbInformation, gstrSysName
        Exit Function
    End If
    strSQL = ""
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    Set cllStatusUpdate = New Collection
    dbl��� = IIf(mblnCheckSwapFailed, StrToNum(txtMoney.Text), Nvl(rsBill!���))
    Call zlGetDepositYJSQL(cllStatusUpdate, rsBill!ID, Nvl(rsBill!��ҳID, 0), strNO, dbl���, blnPrintInvoice, strCurDate, 2)

    '��������֧��
    If zlInterfacePrayMoney(rsBill!ID, strNO, StrToNum(txtMoney.Text), cllStatusUpdate, blnTrans) = False Then
        Exit Function
    End If
    If blnTrans Then
        gcnOracle.CommitTrans: blnTrans = False
    Else
        If cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then
            zlDatabase.ExecuteProcedure cllStatusUpdate(1), Me.Caption
        End If
    End If
    If blnPrintInvoice Then 'Ʊ�ݺ�Ϊ�վͱ�ʾ����ӡ��Ʊ
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_2", Me, "NO=" & cboNO.Text, "����ID=" & mpatiInfo.����ID, _
                                 "�տ�ʱ��=" & Format(strCurDate, "yyyy-mm-dd HH:MM:SS"), 2)
        Call zlCheckFactIsEnough
    End If
    If blnPrintVocher Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_4", Me, "NO=" & cboNO.Text, 2)
    End If
    ReDepositErrBill = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ClearBill()
'����:�����ؽ��������
    If mbytInState = EM_��Ѻ�� And gblnLED Then
        zl9LedVoice.DisplayPatient ""
    End If
    
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
    
    txtMoney.Text = ""
    If Val(cboѺ�����.ItemData(cboѺ�����.ListIndex)) > 0 Then
            txtMoney.Text = Format(Val(cboѺ�����.ItemData(cboѺ�����.ListIndex)), "###0.00;-###0.00;;")
    End If
    If cboStyle.ListCount <> 0 And cboStyle.Tag <> "" Then cboStyle.ListIndex = Val(cboStyle.Tag) '�ָ�ȱʡ���㷽ʽ
    txtCode.Text = "": txtCode.Locked = False
    
    txtMan.Text = UserInfo.����
    txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    cboNote.Text = ""
    
    '�µ�һ��Ѻ�𵥾�
    cboNO.Text = "": cboNO.Locked = True
    
    txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ�ݺ�") And gblnBillԤ�� '89302
    If mFactProperty.intInvoicePrint <> 0 Then Call GetFact
    zlControl.ControlSetFocus txtPatient
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Activate()
    If mblnUnLoad Then Unload Me: Exit Sub
    If mbytInState = EM_��Ѻ�� Then
        If gblnLED And Trim(txtPatient.Text) = "" Then
            zl9LedVoice.DisplayPatient ""    '˫����ʾ��������ڵ�ǰ������ʾ֮�������ʾ�����ƶ�����
        End If
    ElseIf mbytInState = EM_������� Then
        zlControl.ControlSetFocus cmdCancel
    ElseIf mbytInState = EM_��Ѻ�� Or mbytInState = EM_�쳣���� Then
        If mstrInNO = "" Then
            zlControl.ControlSetFocus cboNO
        Else
           zlControl.ControlSetFocus cmdOK
        End If
        If mbytInState = EM_�쳣���� Then txtMoney.Text = Abs(txtMoney.Text): Call InitPatientInfo(mstrInNO)
    ElseIf mbytInState = EM_�쳣���� Or mbytInState = EM_�쳣���� Then
        '��ʼ��������Ϣ
        Call InitPatientInfo(mstrInNO)
        txtMoney.Enabled = False
    End If
    '�����:45666
    If mbytInState = EM_��Ѻ�� And cboType.Text = "סԺѺ��" Then '��Ѻ��
        chk����ʾ����Ѻ��.Visible = True
        chk����ʾ����Ѻ��.Value = IIf(zlDatabase.GetPara("����ʾ����Ԥ��", glngSys, mlngModul, , Array(chk����ʾ����Ѻ��), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
    End If
    
End Sub

Private Sub InitPatientInfo(ByVal strNO As String)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:�����쳣���ݺų�ʼ��������Ϣ
    '����:2018-06-28 17:50:31
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsInfo As ADODB.Recordset
    
    On Error GoTo errHandle

    strSQL = "Select ����id, �����id, ���㷽ʽ From ����Ѻ���¼ Where NO = [1]"
    Set rsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsInfo.RecordCount = 0 Then Exit Sub
    If mlngCardTypeID = 0 Then mlngCardTypeID = Nvl(rsInfo!�����ID, 0)
    mstr���㷽ʽ = Nvl(rsInfo!���㷽ʽ)
    If GetPatiInfo(rsInfo!����ID, -1, mpatiInfo) = False Then
        Set mpatiInfo = New clsPatientInfo
        Exit Sub
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
            zlControl.ControlSetFocus txtFact
        Case vbKeyF4
            If Shift = vbCtrlMask And IDKind.Enabled Then
                Dim intIndex As Integer
                intIndex = IDKind.GetKindIndex("IC����")
                If intIndex <= 0 Then Exit Sub
                 IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
            End If
        Case vbKeyF6
            If btQRCodePay.Visible = False Or btQRCodePay.Enabled = False Then Exit Sub
            Call btQRCodePay.zlReReadQRCode
        Case vbKeyF11
            zlControl.ControlSetFocus txtPatient
        Case vbKeyF12
            zlControl.ControlSetFocus cboNO
        Case vbKeyF8
            If chkCancel.Visible And picNO.Enabled Then chkCancel.Value = IIf(chkCancel.Value = 1, 0, 1)
        Case vbKeyEscape
            Call cmdCancel_Click
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub GetFact(Optional blnFirst As Boolean = False, Optional blnRed As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ͬ���ķ�Ʊ
    '����:���˺�
    '����:2011-07-19 17:47:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
        If Not blnRed Then
            txtFact.Text = GetNextBill(mlng����ID)
        Else
            mstrRedFact = GetNextBill(mlng����ID)
        End If
    Else
        '��ɢ��ȡ��һ������
        If Not blnRed Then
            txtFact.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, mlngFactModule, "")))
        Else
            mstrRedFact = zlCommFun.IncStr(UCase(zlDatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, mlngFactModule, "")))
        End If
    End If
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ģ�����
    '����:���˺�
    '����:2012-02-27 11:23:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrȱʡ���㷽ʽ = zlDatabase.GetPara("ȱʡԤ�����㷽ʽ", glngSys, mlngModul)
    mbytBackMoneyType = Val(zlDatabase.GetPara("�˿��ֹ��ʽ", glngSys, mlngModul))
    '���㷽ʽ:���|���㷽ʽ:���....
    mblnClearWinInfor = IIf(zlDatabase.GetPara("��Ԥ���������Ϣ", glngSys, glngModul) <> "1", True, False)
    mblnδ��Ʋ���Ԥ�� = zlDatabase.GetPara("����δ��Ʋ�׼��Ԥ��", glngSys, mlngModul, , , InStr(mstrPrivs, ";��������;") > 0) = "1"
    gblnSeekName = Nvl(zlDatabase.GetPara("����ģ������", glngSys, mlngModul, 1)) = 1
    mblnסԺ��Ԥ����֤ = zlDatabase.GetPara("סԺ��Ԥ����֤", glngSys, mlngModul, "0") = "1"
    'ˢ��Ҫ����������
    mblnCheckPass = Mid(zlDatabase.GetPara(46, glngSys, , "0000000000"), 8, 1) = "1"
    mbln�ų�δ�ɼ�δ�� = zlDatabase.GetPara("ʣ����ų�δ�ɼ�δ����", glngSys, mlngModul, "0") = "1"
    
End Sub

Private Sub Form_Load()
    
    Call InitPara
    mblnOK = False: mblnUnLoad = False

    'Ʊ�����ü�鼰��ʼ
    If mbytInState = EM_��Ѻ�� Or mbytInState = EM_��Ѻ�� Then
        mblnStartFactUseType = zlStartFactUseType(2)
        If mblnStartFactUseType = False Then
            If mFactProperty.intInvoicePrint <> 0 Then Call GetFact(True, mbytInState = EM_��Ѻ��)
        End If
    End If
    
    zlControl.PicShowFlat picInfo, -1
    zlControl.PicShowFlat picFace, -1

    If Not InitUnit Then Unload Me: Exit Sub
    
    Call InitIDKind
    
    mstrCardPrivs = GetPrivFunc(glngSys, 1151)
    Call InitFace
    If mblnUnLoad Then Exit Sub
    
    lblTitle.Caption = gstrUnitName & "Ѻ�𵥾�"
    
    If (mbytInState = EM_��Ѻ�� Or mbytInState = EM_��Ѻ��) And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.��� & "��Ϊ������", mlngModul, gcnOracle
        
        Call zlCheckFactIsEnough
    End If

    If mbytInState = EM_��Ѻ�� Then
        IDKind.IDKind = Val(zlDatabase.GetPara("�ϴ����뷽ʽ", glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0))
    End If
    
    '81693:���ϴ�,2015/4/21,������
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear: On Error GoTo 0
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInState = EM_��Ѻ��: mstrInNO = ""
    mblnViewCancel = False: mblnUnLoad = False
    mlng����ID = 0: mblnNOMoved = False
    mblnOptErrBill = False
    mstr�˿����Ա = ""
    
    If (mbytInState = EM_��Ѻ�� Or mbytInState = EM_��Ѻ��) And gblnLED Then
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
    Set mpatiInfo = Nothing
    Set mobjPayMode = Nothing
    mblnCheckSwapFailed = False

    If mbytInState = EM_��Ѻ�� Then
        zlDatabase.SetPara "�ϴ����뷽ʽ", IDKind.IDKind, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    End If
    '�����:45666
    If mbytInState = EM_��Ѻ�� And cboType.Text = "סԺѺ��" Then
        zlDatabase.SetPara "����ʾ����Ԥ��", chk����ʾ����Ѻ��.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    End If
End Sub

Private Sub InitPrepayType(Optional bytPrepayType As Byte = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Ѻ������
    '����:���˺�
    '����:2011-07-14 18:50:56
    '---------------------------------------------------------------------------------------------------------------------------------------------

    With cboType
        .Clear
        .AddItem "����Ѻ��": .ItemData(.NewIndex) = 1
        If bytPrepayType = 1 Then .ListIndex = .NewIndex
        .AddItem "סԺѺ��": .ItemData(.NewIndex) = 2
        If bytPrepayType = 2 Then .ListIndex = .NewIndex
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
     End With
End Sub

Private Sub InitFace(Optional blnSave As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ڲ������ô�����漰����״̬
    '����:���˺�
    '����:2011-07-17 10:36:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strValue As String, varData As Variant, varTemp As Variant
    Dim i As Integer, j As Integer, blnChange As Boolean
    
    If Not gobjSquare.objSquareCard Is Nothing And blnSave = False Then
        IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
    End If
    
    On Error GoTo errHandle
    
    strSQL = "Select ����, ����, ����, ȱʡ��־  From ����Ԥ��ժҪ"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    With cboNote
        .Clear
        If rsTmp.RecordCount > 0 Then
            While Not rsTmp.EOF
                .AddItem Nvl(rsTmp!����)
                If Nvl(rsTmp!ȱʡ��־) = 1 Then .ListIndex = .NewIndex
                rsTmp.MoveNext
            Wend
        End If
        .ListIndex = -1
    End With
    
    strSQL = "Select ����, ����, ����, ȱʡ��־  From Ѻ����� Order By ����"
    strSQL = " Select Distinct ����" & _
                  " From (Select b.����, b.����, b.����, Decode(b.ȱʡ��־, 1, 1, 0) As ȱʡ��־" & _
                  "           From ���㷽ʽӦ�� A, ���㷽ʽ B" & _
                  "           Where a.Ӧ�ó��� = 'Ԥ����' And b.���� = a.���㷽ʽ And Nvl(b.����, 1) = 5 " & _
                  "           Union All" & _
                  " Select ����, ����, ����, Decode(ȱʡ��־, 1, 2, 0) As ȱʡ��־ From Ѻ����� Order By ȱʡ��־ Desc, ����)"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cboѺ�����
        .Clear
        If rsTmp.RecordCount > 0 Then
            While Not rsTmp.EOF
                .AddItem Nvl(rsTmp!����)
                rsTmp.MoveNext
            Wend
        Else
            MsgBox "δ�ҵ���Ч��Ѻ����������ֵ�������С����ù����������µġ�Ѻ����������ã�", vbExclamation, gstrSysName
            mblnUnLoad = True
            Exit Sub
        End If
    End With
    
    '����Ѻ�����ȱʡ���
    strValue = zlDatabase.GetPara("���տ�����", glngSys, mlngModul)
    varData = Split(strValue, "|")
    
    With cboѺ�����
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i), ":")
            For j = 0 To cboѺ�����.ListCount - 1
                If varTemp(0) = cboѺ�����.List(j) Then
                    cboѺ�����.ItemData(j) = Val(varTemp(1)): Exit For
                End If
            Next
        Next
        mblnNotClick = True
        .ListIndex = 0
        mblnNotClick = False
    End With
    
    Call InitPrepayType
    If mblnUnLoad Then Exit Sub

    IDKind.Enabled = mbytInState = EM_��Ѻ��
    Select Case mbytInState
        Case EM_��Ѻ�� '��ȡѺ��
            '����������
            Call CreateMobjCard
            cboNO.Text = ""
            txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
            txtMan.Text = UserInfo.����
            
            Call Load֧����ʽ
            '�˿�Ȩ��
            If InStr(mstrPrivs, ";Ѻ���˿�;") = 0 Then
                chkCancel.Visible = False
            End If
            txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ�ݺ�") And gblnBillԤ�� '89302
        Case EM_������� 'ָ���������
            picList.Visible = False
            Me.Height = Me.Height - picList.Height
            If mblnViewCancel Then lblFlag.Visible = True
            chkCancel.Visible = False
            cmdOK.Visible = False
            
            cmdCancel.Caption = "�˳�(&X)"
            
            picNO.Enabled = False
            picFace.Enabled = False
            cboNote.Locked = True
            txtFact.Locked = True
            txtUnit.Locked = True
            txt������.Locked = True
            txt�ʺ�.Locked = True
            
            '��ʾ��������
            If Not ReadBill(mstrInNO) Then
                MsgBox "������ȷ��ȡ�õ������ݣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
                mblnUnLoad = True
            End If
        Case EM_��Ѻ��, EM_�쳣���� 'ָ�������˿�
            chkCancel.Value = Checked   '�ڵ��õ�click�¼��д��� picFace.Enabled = True '���������������˿����
            txtFact.Locked = True
            If mstrInNO <> "" Then  '������Ϣ��������Ԥ��,û��ָ�����ݺ�
                picNO.Enabled = False
                '��ʾ��������
                Dim intBill As Integer
                intBill = ReadBill(mstrInNO)
                If intBill <> -1 Then
                    If intBill <> 3 Then
                        MsgBox "������ȷ��ȡ�õ������ݣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
                    End If
                    mblnUnLoad = True
                Else
                    If mbytInState = EM_��Ѻ�� Then
                        If zlThirdReturnCashCheck(blnChange) Then
                            If cboStyle.ListCount > 1 And blnChange Then
                                cboStyle.Enabled = True: cboStyle.Locked = False
                                cboStyle.BackColor = &H80000005
                            End If
                        End If
                    End If
                End If
            End If
            If mbytInState = EM_�쳣���� Then cmdOK.Caption = "����(&R)": cmdOK.Enabled = True
        Case EM_�쳣����
            cmdOK.Caption = "����(&R)"
            cmdOK.Enabled = True
            chkCancel.Visible = False
            Call Load֧����ʽ
            
            If mstrInNO <> "" Then  '������Ϣ��������Ԥ��,û��ָ�����ݺ�
                '��ʾ��������
                Dim intBillErr As Integer
                intBillErr = ReadBill(mstrInNO)
                If intBillErr <> -1 Then
                    If intBillErr <> 3 Then
                        MsgBox "������ȷ��ȡ�õ������ݣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
                    End If
                    mblnUnLoad = True
                End If
            End If
        Case EM_�쳣����
            If mblnViewCancel Then lblFlag.Visible = True
            chkCancel.Visible = False
            
            cmdOK.Caption = "����(&Z)"
            
            picNO.Enabled = False
            picFace.Enabled = False
            txtPatient.Enabled = False
            cboStyle.Enabled = False
            cboPatiPage.Enabled = False
            cboType.Enabled = False
            cboUnit.Enabled = False
            '��ʾ��������
            If Not ReadBill(mstrInNO) Then
                MsgBox "������ȷ��ȡ�õ������ݣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
                mblnUnLoad = True
            End If
    End Select

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

Public Sub CreateMobjCard()
    '����������
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hwnd)
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hwnd)
    Set mobjICCard.gcnOracle = gcnOracle
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

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0: txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If IsDate(txtDate.Text) And KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtMoney_Change()
    '����27363
    If IsNumeric(StrToNum(txtMoney.Text)) Then
        txtMoney.ForeColor = IIf(CCur(StrToNum(txtMoney.Text)) >= 0, vbBlue, vbRed)
    End If
End Sub

Private Sub txtMoney_GotFocus()
    txtMoney.SelStart = 0: txtMoney.SelLength = Len(txtMoney.Text)
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    '����27363
    If KeyAscii <> 13 Then
        If KeyAscii = Asc(".") And InStr(txtMoney.Text, ".") > 0 Then KeyAscii = 0: Beep: Exit Sub
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
    Else
        If txtMoney.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtMoney_LostFocus()
    '����27363
    Dim dblMoney  As Double
    If Not IsNumeric(StrToNum(txtMoney.Text)) Then zlControl.ControlSetFocus txtMoney: Exit Sub
    If mpatiInfo.����ID > 0 And IsNumeric(StrToNum(txtMoney.Text)) Then
        txtMoney.Text = Format(StrToNum(txtMoney.Text), "##,##0.00;-##,##0.00; ;")
        If txtMoney.MaxLength > 12 Then txtMoney.MaxLength = 12
        '108813:���ϴ�,2017/5/8,������������
        If gblnLED Then
            '#22 1234.56   --Ԥ��һǧ������ʮ�ĵ�����Ԫ Y
            '#23 1234.56   --����һǧ������ʮ�ĵ�����Ԫ Z
            dblMoney = StrToNum(txtMoney.Text)
            zl9LedVoice.Speak "#22 " & dblMoney
        End If
    End If
End Sub

Private Sub cboNO_GotFocus()
    If Not cboNO.Locked Then cboNO.SelStart = 0: cboNO.SelLength = Len(cboNO.Text)
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim strOper As String, vDate As Date
    Dim blnChange As Boolean
    
    If cboNO.Locked Then Exit Sub
    
    'ת���ɴ�д(���ֲ��ɴ���)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    ElseIf cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 11)
        
        '�Ƿ���ת������ݱ���,��¼����Ϊ1��ʾ�����Ԥ��
        If zlDatabase.NOMoved("����Ѻ���¼", cboNO.Text, "", "", Me.Caption) Then
            If Not ReturnMovedExes(cboNO.Text, 6, Me.Caption) Then Exit Sub
            mblnNOMoved = False
        End If
        
        '����Ȩ��
        If Not ReadBillInfo(0, cboNO.Text, -21, strOper, vDate) Then
            cboNO.Text = "": zlControl.ControlSetFocus cboNO: Exit Sub
        End If
        If Not BillOperCheck(6, strOper, vDate, "�˿�") Then
            cboNO.Text = "": zlControl.ControlSetFocus cboNO: Exit Sub
        End If
        '����27363
        '��ȡҪ���ϵ�Ԥ�����
        Select Case ReadBill(cboNO.Text)
            Case -1
                If InStr(mstrPrivs, ";Ѻ���˿�;") = 0 Then
                    MsgBox "��û��Ȩ�޽���Ѻ���˿������", vbInformation, gstrSysName
                    chkCancel.Value = 0
                Else
                    If Val(StrToNum(txtMoney.Text)) < 0 Then
                        MsgBox "�ñ�Ԥ�����Ϊ��,��ʾ�˿�,����ִ�иò�����", vbExclamation, gstrSysName
                        chkCancel.Value = 0
                    Else
                        zlControl.ControlSetFocus cmdOK
                    End If
                End If
                If chkCancel.Value <> 0 Then
                    If zlThirdReturnCashCheck(blnChange) Then
                        If cboStyle.ListCount > 1 And blnChange Then
                            cboStyle.Enabled = True: cboStyle.Locked = False
                            cboStyle.BackColor = &H80000005
                        End If
                    End If
                End If
            Case 0
                MsgBox "��ȡ��Ѻ�𵥾�ʧ�ܣ�", vbExclamation, gstrSysName
                cboNO.Text = "": zlControl.ControlSetFocus cboNO
            Case 1
                MsgBox "��Ѻ�𵥾ݲ����ڣ�", vbExclamation, gstrSysName
                cboNO.Text = "": zlControl.ControlSetFocus cboNO
            Case 2
                MsgBox "��Ѻ�𵥾��Ѿ��˿", vbExclamation, gstrSysName
                cboNO.Text = "": zlControl.ControlSetFocus cboNO
            Case 3
                cboNO.Text = "": zlControl.ControlSetFocus cboNO
        End Select
    End If
End Sub

Private Sub cboNote_GotFocus()
    cboNote.SelStart = 0: cboNote.SelLength = Len(cboNote.Text)
End Sub

Private Sub cboNote_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtPatient_Change()
    If Not Me.ActiveControl Is txtPatient Or txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjIDCard.SetEnabled(True)
    If Not mobjICCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjICCard.SetEnabled(True)
    txtPatient.Tag = ""
End Sub

Private Sub ClearWinInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������Ϣ
    '����:���˺�
    '����:2012-02-27 11:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '�����:55666
    lblRepairMoney.Caption = "������:": lblRepairMoney.Visible = False
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    If txtPatient.Locked Then Exit Sub
    
    Call ClearWinInfor
        
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
    
    If txtPatient.Tag <> "" Then Exit Sub
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        Set frmPatiSelect.mfrmParent = Me
        frmPatiSelect.mbytSize = 1 '������(С��)
        frmPatiSelect.Show 1, Me
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
        End If
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnCancel As Boolean, bytPrepayType As Byte
    Dim str������ As String, dbl������ As Double
    
    Call ClearBill
    '��ȡ������Ϣ
    SetMoneyInfo True
    sta.Panels(2) = ""
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If Not GetPatient(objCard, strInput, blnCancel, blnCard) Then
        '�����쳣��������
        If mblnOptErrBill = False Then
            If blnCancel Then 'ȡ������
                zlControl.ControlSetFocus txtPatient: Exit Sub
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
            zlControl.ControlSetFocus txtPatient
        End If
        Exit Sub
    End If
    
    '���ز��˵�סԺ����
    Call LoadPatiPage(mpatiInfo.����ID)
    '���ò��˷�����Ϣ
    Call SetMoneyInfo(False, mpatiInfo.����ID)
    
    'ȱʡ���˵�Ѻ������
    bytPrepayType = IIf(mpatiInfo.��Ժ, 2, 1)
    If bytPrepayType <> cboType.ItemData(cboType.ListIndex) Then
        Call InitPrepayType(bytPrepayType)
    End If
    
    If mpatiInfo.��ǰ����ID <> 0 Then
        lbl����.Caption = lbl����.Tag & IIf(mpatiInfo.���� = "", "��ͥ", mpatiInfo.����)
    End If
            
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(mpatiInfo.סԺ�� = "", "", "סԺ��:" & mpatiInfo.סԺ�� & "   ") & _
                           IIf(mpatiInfo.����� = "", "", "�����:" & mpatiInfo.�����)
    lbl����.Caption = lbl����.Tag & GET��������(mpatiInfo.��Ժ����ID)
    '46764
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(mpatiInfo.��ǰ����ID = 0, mpatiInfo.��Ժ����ID, mpatiInfo.��ǰ����ID))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    
    lbl�ѱ�ȼ�.Caption = lbl�ѱ�ȼ�.Tag & mpatiInfo.�ѱ�
    Call Get������Ϣ(mpatiInfo.����ID, mpatiInfo.��ҳID, dbl������, str������)
    lbl������.Caption = lbl������.Tag & str������
    lbl�������.Caption = lbl�������.Tag & Format(dbl������, "##,##0.00;-##,##0.00; ;")
    '�����:116059,����,2017/12/7,Ԥ��������ʾ�����ֻ��ţ���ȡ������Ϣ�еġ��ֻ��š�
    lbl�ֻ���.Caption = lbl�ֻ���.Tag & mpatiInfo.�ֻ���
    lbl���֤��.Caption = lbl���֤��.Tag & mpatiInfo.���֤��
    lblMemo.Caption = lblMemo.Tag & mpatiInfo.���˱�ע
    '72828,Ƚ����,2014-5-9,���ӹ�����λ��Ϣ����ʾ
    lblWorkUnit.Caption = lblWorkUnit.Tag & mpatiInfo.������λ
    
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    txtPatient.Text = mpatiInfo.����
    txtPatient.Tag = mpatiInfo.����ID
    '-----------------------------------------------------------------------------------------
    lblSex.Caption = lblSex.Tag & mpatiInfo.�Ա�
    mstrPatiSex = mpatiInfo.�Ա�
    lblOld.Caption = lblOld.Tag & mpatiInfo.����
    mstrPatiOld = mpatiInfo.����
    lbl��ͥ��ַ.Caption = lbl��ͥ��ַ.Tag & mpatiInfo.��ͥ��ַ
    lblҽ�Ƹ��ʽ.Caption = lblҽ�Ƹ��ʽ.Tag & mpatiInfo.ҽ�Ƹ��ʽ
    Call Led��ӭ��Ϣ
    Call SetcmdOkEnabled
    
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Led��ӭ��Ϣ()
    Dim strInfo As String, lngPatient As Long
    'LED��ʼ��
    If mbytInState = EM_��Ѻ�� And gblnLED Then
        If gblnLedWelcome Then
            zl9LedVoice.Reset com
            zl9LedVoice.Speak "#1"
            zl9LedVoice.Init UserInfo.��� & "��Ϊ������", mlngModul, gcnOracle
        End If
        strInfo = Trim(txtPatient.Text)
        If mpatiInfo.��ǰ����ID > 0 Then strInfo = strInfo & " " & mpatiInfo.�Ա� & " " & mpatiInfo.����: lngPatient = mpatiInfo.����ID
        zl9LedVoice.DisplayPatient strInfo, lngPatient
    End If
End Sub
 
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
                                           Optional ByRef blnCancel As Boolean, _
                                           Optional ByVal blnCard As Boolean, _
                                           Optional ByVal lng����id As Long, _
                                           Optional ByVal lng��ҳID As Long = -1) As Boolean
    '���ܣ���ȡ������Ϣ
    '������strInput=[ˢ��]|[A����ID]|[BסԺ��]
    '          lng����ID=����id,����סԺ��������Ѻ���¼ʱ����
    '          lng��ҳID=-1��ʾ���ﲡ�˻��������סԺ����;lng��ҳID=0��ʾԤ��Ժ����;lng��ҳID>0��ʾסԺ����
    '˵����
    '     1.�����ڲ���Ԥ����
    '     2.�Զ�ʶ������Ժ״̬,����(����ID,��ҳID,����,�Ա�,����,סԺ��,����,��Ժ��־)
    '����:�Ƿ��ȡ�ɹ�,�ɹ�ʱmPatiInfo�а���������Ϣ
    Dim lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim blnHavePassWord As Boolean, blnIsMobileNO As Boolean
    
    blnCancel = False: mstr�˿����Ա = ""
    If lng����id > 0 Then GoTo ReadPati
    
    blnIsMobileNO = IDKind.IsMobileNo(strInput)
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
                If GetPatiID("�����", Mid(strInput, 2), lng����id) = False Then GoTo NotFoundPati
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
    If mbytInState = EM_��Ѻ�� Then
        If OptOthersErrBill(mpatiInfo.����ID) Then
            Exit Function
        End If
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

Private Function SaveBill(Optional blnPrintInvoice As Boolean = False, Optional ByRef lngѺ��ID As Long, Optional ByRef blnBeenErr As Boolean, Optional ByRef strCurDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ե�ǰ�����Ԥ����ݴ���
    '����: blnBeenErr-�Ƿ����쳣������true-���쳣��false-��
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-17 11:15:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, i As Integer
    Dim blnTrans As Boolean, dblMoney As Double
    Dim lng��ҳID As Long
    Dim blnThirdCard As Boolean
    Dim cllDeposit As Collection, cllStatusUpdate As Collection
    
    '��ǰ�����Ƿ�������
    blnThirdCard = cboStyle.ItemData(cboStyle.ListIndex) = -1 And mlngCardTypeID <> 0
    
    
    strNO = zlDatabase.GetNextNo(11)
    lngѺ��ID = zlDatabase.GetNextId("����Ԥ����¼")
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    '����27363
    dblMoney = 1 * StrToNum(txtMoney.Text)
    
Once:
    
    lng��ҳID = IIf(cboType.ItemData(cboType.ListIndex) = 2, mpatiInfo.��ҳID, 0)
    If cboPatiPage.Visible And cboPatiPage.ListIndex > 0 And cboType.ItemData(cboType.ListIndex) = 2 Then
        lng��ҳID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If

    '��ȡԤ��SQL
    Set cllDeposit = New Collection
    Call zlGetDepositYJSQL(cllDeposit, lngѺ��ID, lng��ҳID, strNO, dblMoney, blnPrintInvoice, strCurDate, _
                        IIf(blnThirdCard And mbytInState = EM_��Ѻ�� And gbln���ý����첽����, 1, 0))
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    zlExecuteProcedureArrAy cllDeposit, Me.Caption, True, True
    
    If gbln���ý����첽���� And blnThirdCard Then
        Set cllStatusUpdate = New Collection
        Call zlGetDepositYJSQL(cllStatusUpdate, lngѺ��ID, lng��ҳID, strNO, dblMoney, blnPrintInvoice, strCurDate, 2)
    End If

    If zlInterfacePrayMoney(lngѺ��ID, strNO, StrToNum(txtMoney.Text), cllStatusUpdate, blnTrans, blnBeenErr) = False Then
        If blnBeenErr Then mblnOK = True
        Exit Function
    End If
  
    If blnTrans Then
        gcnOracle.CommitTrans: blnTrans = False
    End If
    
    '���뵥����ʷ��¼(�������͵���)
    For i = 0 To cboNO.ListCount - 1
        strNO = strNO & "," & cboNO.List(i)
    Next
    cboNO.Clear
    For i = 0 To UBound(Split(strNO, ","))
        cboNO.AddItem Split(strNO, ",")(i)
        If i = 9 Then Exit For 'ֻ��ʾ10��
    Next
    
    If Not gblnBillԤ�� And blnPrintInvoice And Trim(txtFact.Text) <> "" Then
        '��ɢ�����浱ǰ����
        zlDatabase.SetPara "��ǰԤ��Ʊ�ݺ�", Trim(txtFact.Text), glngSys, mlngFactModule
    End If
    SaveBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If Err.Description Like "*�˿�����ڲ���Ѻ�����*" And mbytOracleBackType = 1 Then
        If MsgBox("�˿���Ȳ���Ѻ������,�Ƿ���ԣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
        mbytOracleBackType = 0
        GoTo Once
    End If
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetDeleteSQL(ByVal strNO As String, Optional ByVal bytOpt As Byte) As String
    Dim strSQL As String
    
    strSQL = "Zl_����Ѻ���쳣��¼_Delete("
    '  ���ݺ�_In       ����Ѻ���¼.No%Type
    strSQL = strSQL & "'" & strNO & "'," & bytOpt & ")"
    GetDeleteSQL = strSQL
End Function

Private Function zlGetDepositYJSQL(cllDeposit As Collection, ByVal lngѺ��ID As Long, ByVal lng��ҳID As Long, _
                            ByVal strNO As String, ByVal dblMoney As Double, ByVal blnPrintInvoice As Boolean, _
                            ByVal strCurDate As String, ByVal byt����״̬ As Byte) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : -��ȡѺ���ֵSQL
    ' ���� : ����-cllDeposit����Σ�strNo-���ݺţ�dblMoney-��blnPrintInvoice-�Ƿ��ӡ��Ʊ
    '                               strCurDate-�տ����ڣ�btyOptType-��������
    '                               bytУ�Ա�־-У�Ա�־��byt����״̬-����״̬
    '                           --����״̬:0-�������㣬1-����Ϊ�쳣���ݣ�2-����쳣����
    '---------------------------------------------------------------------------------------
    
    Dim strSQL As String
    
    'Zl_����Ѻ���¼_Insert_S
    strSQL = "Zl_����Ѻ���¼_Insert_S("
    '  Id_In         ����Ѻ���¼.ID%Type,
    strSQL = strSQL & "" & lngѺ��ID & ","
    '  ���ݺ�_In     ����Ѻ���¼.NO%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  Ʊ�ݺ�_In     Ʊ��ʹ����ϸ.����%Type,
    If blnPrintInvoice Then
        strSQL = strSQL & "'" & txtFact.Text & "',"
    Else
        strSQL = strSQL & "NULL,"
    End If
    '  Ѻ�����_In     ����Ѻ���¼.Ѻ�����%Type,
    strSQL = strSQL & "'" & cboѺ�����.Text & "',"
    '  ����id_In     ����Ѻ���¼.����id%Type,
    strSQL = strSQL & "" & mpatiInfo.����ID & ","
    '  ��ҳid_In     ����Ѻ���¼.��ҳid%Type,
    strSQL = strSQL & "" & ZVal(lng��ҳID) & ","
    '  ����_In         ����Ѻ���¼.����%Type,
    strSQL = strSQL & "'" & mpatiInfo.���� & "',"
    '  �Ա�_In         ����Ѻ���¼.�Ա�%Type,
    strSQL = strSQL & "'" & mpatiInfo.�Ա� & "',"
    '  ����_In         ����Ѻ���¼.����%Type,
    strSQL = strSQL & "'" & mpatiInfo.���� & "',"
    '  �����_In       ����Ѻ���¼.�����%Type,
    strSQL = strSQL & ZVal(mpatiInfo.�����) & ","
    '  סԺ��_In       ����Ѻ���¼.סԺ��%Type,
    strSQL = strSQL & ZVal(mpatiInfo.סԺ��) & ","
    '  ���ʽ����_In ����Ѻ���¼.���ʽ����%Type,
    strSQL = strSQL & "'" & mpatiInfo.ҽ�Ƹ��ʽ & "',"
    '  ����id_In     ����Ѻ���¼.����id%Type,
    strSQL = strSQL & "" & ZVal(cboUnit.ItemData(cboUnit.ListIndex)) & ","
    '  �ɿλ_In   ����Ѻ���¼.�ɿλ%Type,
    strSQL = strSQL & "'" & Trim(txtUnit.Text) & "',"
    '  ��λ������_In ����Ѻ���¼.��λ������%Type,
    strSQL = strSQL & "'" & Trim(txt������.Text) & "',"
    '  ��λ�ʺ�_In   ����Ѻ���¼.��λ�ʺ�%Type,
    strSQL = strSQL & "'" & Trim(txt�ʺ�.Text) & "',"
    '  ժҪ_In       ����Ѻ���¼.ժҪ%Type,
    strSQL = strSQL & "'" & Trim(cboNote.Text) & "',"
    '  ���_In       ����Ѻ���¼.���%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  ���㷽ʽ_In   ����Ѻ���¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & mstr���㷽ʽ & "',"
    '  �������_In   ����Ѻ���¼.�������%Type,
    strSQL = strSQL & "'" & txtCode.Text & "',"
    '  �Ƿ�����_In   ����Ѻ���¼.Ԥ�����%Type := Null,
    strSQL = strSQL & "" & IIf(cboType.ItemData(cboType.ListIndex) = 1, 1, 0) & ","
    '  ����id_In     ����Ѻ���¼.����id%Type,
    strSQL = strSQL & "" & ZVal(mlng����ID) & ","
    '  ����Ա���_In ����Ѻ���¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ����Ѻ���¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����_In       ����Ѻ���¼.����%Type := Null,
    strSQL = strSQL & "" & IIf(mstrBrushCardNo = "", "NULL", "'" & mstrBrushCardNo & "'") & ","
    '  �տ�ʱ��_In   ����Ѻ���¼.�տ�ʱ��%Type := Null
    strSQL = strSQL & "to_date('" & strCurDate & "','yyyy-mm-dd hh24:mi:ss'),"
    '  �����id_In   ����Ѻ���¼.�����id%Type := Null,
    strSQL = strSQL & "" & ZVal(mlngCardTypeID) & ","
    '  ������ˮ��_In ����Ѻ���¼.������ˮ��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ����˵��_In   ����Ѻ���¼.����˵��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '   ����״̬_In     Number :=0
    strSQL = strSQL & byt����״̬ & ")"
   
    zlAddArray cllDeposit, strSQL
    
    zlGetDepositYJSQL = True

End Function

Private Function ReadBill(strNO As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡԤ�����(����ġ��˿��),����д���漰����mpatiInfo(������Ϣ),��������Tag��
    '���:strNO-Ѻ�𵥾ݺ�
    '����:
    '����: -1-�ɹ�;0-ʧ��;1-�õ��ݲ�����;2:�õ����Ѿ��˿�(���ʱ��Ч);3-Ȩ�޲���(������)
    '����:���˺�
    '����:2011-07-15 11:45:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngPrepayType As Long, rsTemp As New ADODB.Recordset, strFullNO As String
    Dim strWhere As String, i As Long, strTmp As String
    Dim rs��Ժ����ʾ As New ADODB.Recordset
    Dim str������ As String, dbl������ As Double
    Dim lng�����ID, str��������� As String, bln�˿��鿨 As Boolean
    Dim str���㷽ʽ As String, objCards As New Cards  'zlOneCardComLib.Cards
    On Error GoTo errH
    
    strFullNO = GetFullNO(strNO, 11)
    
    strWhere = IIf(mbytInState = EM_������� And mblnViewCancel, "And A.��¼״̬=2", " And A.��¼״̬ IN(0,1,3) ")
    gstrSQL = "" & _
    "Select a.Id, a.Ѻ�����, a.ʵ��Ʊ��, a.����id, a.��ҳid, a.����id As ��ǰ����ID, a.��¼״̬, a.ժҪ, a. ���, a.���㷽ʽ, a.�������, a.�տ�ʱ��, a.����Ա����, a.�ɿλ," & vbNewLine & _
    "       a.��λ������, a.��λ�ʺ�, a.�����id, a.����, a.������ˮ��, a.����˵��, b.���� As ��������, a.�Ƿ�����" & vbNewLine & _
    "From " & IIf(mblnNOMoved, "H", "") & "����Ѻ���¼ A, ���㷽ʽ B" & vbNewLine & _
    "Where a.No = [1]  And a.��¼״̬ In (0, 1, 3) And a.���㷽ʽ = b.����(+)" & strWhere
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFullNO)
    If rsTemp.RecordCount = 0 Then ReadBill = 1: Exit Function
    If Val(Nvl(rsTemp!�����ID)) > 0 Then
        If gOneCardData.zlGetYLCardObjs(objCards) = False Then Exit Function
        lng�����ID = Val(Nvl(rsTemp!�����ID))
        If objCards("K" & lng�����ID) Is Nothing Then
            MsgBox "δ�ҵ������idΪ" & lng�����ID & "��ҽ�ƿ���Ϣ,�����Ƿ�������!", vbOKOnly + vbInformation, gstrSysName
        Else
            str��������� = objCards("K" & lng�����ID).����
            bln�˿��鿨 = objCards("K" & lng�����ID).�Ƿ��˿��鿨 = 1
        End If
    End If
    If GetPatiInfo(Val(Nvl(rsTemp!����ID)), IIf(Val(Nvl(rsTemp!��ҳID)) = 0, -1, Val(Nvl(rsTemp!��ҳID))), mpatiInfo) = False Then Exit Function
    If mpatiInfo.����ID = 0 Then
        MsgBox "δ�ҵ�������Ϣ,����!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If

    If mbytInState = EM_��Ѻ�� Or chkCancel.Value = 1 Then
        '�˿�,��Ҫ����Ƿ���ھ�����˿�Ȩ��
        
        If InStr(1, mstrPrivs, ";Ѻ���˿�;") = 0 Then
            MsgBox "�㲻�߱���Ѻ�𵥾ݽ����˿��Ȩ��,����ϵͳ����Ա��ϵ!", vbOKOnly + vbInformation, gstrSysName
            ReadBill = 3
            Exit Function
        End If
        
        If gbln��Ժ����ʾ Then
            strTmp = "Select 1 From ���ű� A, ������Ա B, ��Ա�� C" & vbNewLine & _
                    " Where a.Id = b.����id And b.��Աid = c.Id And (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null) And c.���� =[1]  And Rownum < 2"
    
            Set rs��Ժ����ʾ = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, Nvl(rsTemp!����Ա����))
            
            If rs��Ժ����ʾ.RecordCount = 0 Then
                MsgBox "��Ѻ�𵥾ݲ����ڱ�վ��,�������˿�!", vbOKOnly + vbInformation, gstrSysName
                ReadBill = 3: Exit Function
            End If
        End If
    End If
    
    With mYJinfo
        .lngѺ��ID = Val(Nvl(rsTemp!ID))
        .strNO = strFullNO
        .lng�����ID = Val(Nvl(rsTemp!�����ID))
        .str���� = Nvl(rsTemp!����)
        .str���� = str���������
        .str������ˮ�� = Nvl(rsTemp!������ˮ��)
        .dbl��� = Val(Nvl(rsTemp!���))
        .str����˵�� = Nvl(rsTemp!����˵��)
        .bln�˿��鿨 = bln�˿��鿨
        .dt�տ�ʱ�� = Format(rsTemp!�տ�ʱ��, "yyyy-MM-dd hh:mm:ss")
    End With
    
    cboNO.Text = strFullNO
    cboNO.Tag = rsTemp!ID '�Դ�IDΪ׼�˿�
    txtPatient.Text = mpatiInfo.����
    txtPatient.Tag = rsTemp!����ID
    '74426:���ϴ�,2014-7-9,����������ʾ��ɫ����
    Call SetPatiColor(txtPatient, Nvl(mpatiInfo.��������), IIf(Val(mpatiInfo.����) = 0, &HFF0000, vbRed))
    lbl�ѱ�ȼ�.Caption = lbl�ѱ�ȼ�.Tag & mpatiInfo.�ѱ�
    
    Call Get������Ϣ(rsTemp!����ID, Val(Nvl(rsTemp!��ҳID)), dbl������, str������)
    lbl������.Caption = lbl������.Tag & str������
    lbl�������.Caption = lbl�������.Tag & dbl������
    lbl�ֻ���.Caption = lbl�ֻ���.Tag & mpatiInfo.�ֻ���
    lbl���֤��.Caption = lbl���֤��.Tag & mpatiInfo.���֤��

    '72828,Ƚ����,2014-5-9,���ӹ�����λ��Ϣ����ʾ
    lblWorkUnit.Caption = lblWorkUnit.Tag & Nvl(mpatiInfo.������λ)
    
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(IsNull(rsTemp!��ǰ����ID), 0, rsTemp!��ǰ����ID))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    cboType.ListIndex = -1
    lngPrepayType = Val(Nvl(rsTemp!�Ƿ�����))
    For i = 0 To cboType.ListCount - 1
         If cboType.ItemData(i) = IIf(lngPrepayType = 1, 1, 2) Then
            cboType.ListIndex = i: Exit For
         End If
     Next
     
     With cboType
        If cboType.ListIndex < 0 Then
           .AddItem IIf(lngPrepayType = 1, "����Ѻ��", "סԺѺ��")
           .ItemData(.NewIndex) = IIf(lngPrepayType = 1, 1, 2)
           .ListIndex = .NewIndex
        End If
     End With
     
     With cboPatiPage
        .Clear
        .Visible = lngPrepayType <> 1
        If Val(Nvl(rsTemp!��ҳID)) <> 0 Then
            .AddItem "��" & rsTemp!��ҳID & "��"
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!��ҳID))
            .ListIndex = .NewIndex
        Else
            .AddItem "ԤԼ��Ժ"
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!��ҳID))
            .ListIndex = .NewIndex
        End If
     End With
    
    txtFact.Tag = txtFact.Text
    txtFact.Text = IIf(IsNull(rsTemp!ʵ��Ʊ��), "", rsTemp!ʵ��Ʊ��)
    If mbytInState = EM_�쳣���� Then txtFact.Text = txtFact.Tag
    txtUnit.Text = IIf(IsNull(rsTemp!�ɿλ), "", rsTemp!�ɿλ)
    txt������.Text = IIf(IsNull(rsTemp!��λ������), "", rsTemp!��λ������)
    txt�ʺ�.Text = IIf(IsNull(rsTemp!��λ�ʺ�), "", rsTemp!��λ�ʺ�)
    
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(Val(Nvl(mpatiInfo.סԺ��)) = 0, "", "סԺ��:" & mpatiInfo.סԺ�� & "   ") & _
                           IIf(Val(Nvl(mpatiInfo.�����)) = 0, "", "�����:" & mpatiInfo.�����)
    lblSex.Caption = lblSex.Tag & mpatiInfo.�Ա�
    mstrPatiSex = mpatiInfo.�Ա�
    lblOld.Caption = lblOld.Tag & mpatiInfo.����
    mstrPatiOld = mpatiInfo.����
    lbl����.Caption = lbl����.Tag & mpatiInfo.����
    lbl����.Caption = lbl����.Tag & GET��������(Val(Nvl(rsTemp!��ǰ����ID)))
    lbl��ͥ��ַ.Caption = lbl��ͥ��ַ.Tag & Nvl(mpatiInfo.��ͥ��ַ)
    lblҽ�Ƹ��ʽ.Caption = lblҽ�Ƹ��ʽ.Tag & Nvl(mpatiInfo.ҽ�Ƹ��ʽ)
    txtMoney.Text = Format(rsTemp!���, "##,##0.00;-##,##0.00;;")
    txtMoney.Tag = rsTemp!���
    If mYJinfo.lng�����ID <> 0 Then
        cboStyle.ListIndex = cbo.FindIndex(cboStyle, mYJinfo.str����, True)
    Else
        cboStyle.ListIndex = cbo.FindIndex(cboStyle, IIf(IsNull(rsTemp!���㷽ʽ), "", rsTemp!���㷽ʽ), True)
     End If
    If cboStyle.ListIndex = -1 Then
        str���㷽ʽ = IIf(IsNull(rsTemp!���㷽ʽ), "", rsTemp!���㷽ʽ)
        If mYJinfo.lng�����ID <> 0 Then
            cboStyle.AddItem mYJinfo.str����
            cboStyle.ItemData(cboStyle.NewIndex) = -1
            Call MakeCardsFrom���㷽ʽ(mYJinfo.str����, str���㷽ʽ, Val(Nvl(rsTemp!�����ID)))
        Else
            cboStyle.AddItem str���㷽ʽ
            cboStyle.ItemData(cboStyle.NewIndex) = Val("" & rsTemp!��������)
            Call MakeCardsFrom���㷽ʽ(str���㷽ʽ, str���㷽ʽ)
        End If
        cboStyle.ListIndex = cboStyle.NewIndex
    End If
    
    txtCode.Text = IIf(IsNull(rsTemp!�������), "", rsTemp!�������)
    txtMan.Text = IIf(IsNull(rsTemp!����Ա����), "", rsTemp!����Ա����)
    txtDate.Text = Format(rsTemp!�տ�ʱ��, "yyyy-MM-dd")
    cboNote.Text = IIf(IsNull(rsTemp!ժҪ), "", rsTemp!ժҪ)
    mblnNotClick = True
    If Nvl(rsTemp!Ѻ�����) <> "" Then cboѺ�����.ListIndex = cbo.FindIndex(cboѺ�����, Nvl(rsTemp!Ѻ�����), True)

    mblnNotClick = False
    '��ȡ���˷�����Ϣ
    Call SetMoneyInfo(False, rsTemp!����ID, strNO)
    ReadBill = -1
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetDepositData(ByVal lng����id As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¶�ȡԤ������
    '���:lng����ID-����ID��
    '����:���˺�
    '����:2011-07-22 17:02:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    '����:������صĽ��㷽ʽ����������,��ʾ�������ͷ�����Ϣ
    '���:blnReRead-�ض�����
    '       lng����ID-��ȡָ���Ĳ���ID(0ʱ,��mPatiInfo��¼�ж�ȡ����ID)
    '����:���˺�
    '����:2011-07-21 15:44:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim intѺ������ As Integer, bln�����ӿ� As Boolean
    Dim strWhere As String
    Dim dblδ�� As Double, dblδ�� As Double, dblYB As Double
    Dim lng��ҳID As Long, dblʣ���� As Double
    Dim rsYJMoney As New ADODB.Recordset, dblѺ����� As Double
    
    On Error GoTo errHandle
    If lng����id = 0 Then
        If mpatiInfo.����ID = 0 Then Exit Sub
        lng����id = mpatiInfo.����ID
    End If
    
    If blnreReadData Then Call GetDepositData(lng����id)
    sta.Panels(2).Text = ""
    mdbl������� = 0: mdblԤ����� = 0: mdblʣ���� = 0
    intѺ������ = cboType.ItemData(cboType.ListIndex)
    bln�����ӿ� = cboStyle.ItemData(cboStyle.ListIndex) = -1
    strWhere = "And nvl(�����ID,0)<>0 or nvl(���㿨���,0)<>0 "

    If Not mrsDepositBalance Is Nothing Then
    With mrsDepositBalance
        .Filter = "����=" & intѺ������
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
    
    Call Loadҽ��Ԥ��(lng����id, lng��ҳID, dblYB)
    strSQL = "Select Sum(Nvl(���, 0)) as Ѻ����� From ����Ѻ���¼ Where nvl(У�Ա�־,0) =0 And ����id = [1] " & IIf(lng��ҳID = 0, "", "And ��ҳid = [2]")
    Set rsYJMoney = zlDatabase.OpenSQLRecord(strSQL, "��ȡѺ�����", lng����id, lng��ҳID)
    
    dblѺ����� = Val(Nvl(rsYJMoney!Ѻ�����, 0))
    mdblʣ���� = mdblԤ����� - mdbl�������
    '����27363
    lbl�������.Caption = lbl�������.Tag & Format(mdbl�������, "##,##0.00;-##,##0.00; ;")
    lblѺ�����.Caption = lblѺ�����.Tag & Format(dblѺ�����, "##,##0.00;-##,##0.00; ;")
    dblδ�� = GetUnAuditedFee(lng����id, , intѺ������)
    dblδ�� = GetUnAuditedFee(lng����id, False, intѺ������)
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

Private Sub Loadҽ��Ԥ��(ByVal lng����id As Long, ByVal lng��ҳID As Long, Optional ByRef dblYB As Double)
    '����:��ȡ���˵�ҽ��Ԥ����
    Dim rsMoney As ADODB.Recordset, strSQL As String

    On Error GoTo errHandle
    
    If lng����id = 0 Then Exit Sub
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
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetMoneyInfo(blnClear As Boolean, Optional lng����id As Long, _
    Optional strBackNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ϣ
    '���:blnClear-���
    '     lng����ID-ָ������ID
    '     strBackNO-ָ����Ԥ������(�˿�ʱ����,��Ҫ���Ƕ�λ���嵥����ȥ)
    '����:���˺�
    ' �޸�:���˺�(�˺�ʱ,���Ӷ�λ����),���Ӳ���;strBackNo
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
        '72828,Ƚ����,2014-5-9,���ӹ�����λ��Ϣ����ʾ
        lblWorkUnit.Caption = lblWorkUnit.Tag
        
        lblδ�����.Caption = lblδ�����.Tag
        lblδ�ɷ���.Caption = lblδ�ɷ���.Tag
        lbl�������.Caption = lbl�������.Tag
        lblѺ�����.Caption = lblѺ�����.Tag
        lblʣ����.Caption = lblʣ����.Tag
        lblҽ��Ԥ��.Caption = lblҽ��Ԥ��.Tag
        lbl�ֻ���.Caption = lbl�ֻ���.Tag
        lbl���֤��.Caption = lbl���֤��.Tag
        lblӦ�տ�.Caption = lblӦ�տ�.Tag
        lblӦ�տ�.ForeColor = &H80000007
        
        mdbl������� = 0
        mdblԤ����� = 0
        mdblʣ���� = 0
        
        mshList.Redraw = False
        mshList.Clear
        mshList.Rows = 2
        mshList.Cols = 2
        mshList.Redraw = True
    Else
        On Error GoTo errHandle
        '��ʾ�������ͷ�����Ϣ
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
        Call ShowHistoryPrepay(strBackNo)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowHistoryPrepay(ByVal strBackNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ʷ�Ľ�Ѻ������
    '����:���˺�
    '����:2011-09-16 10:17:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, int���� As Integer, lngRow As Long, strWhere As String
    Dim rsMoney As ADODB.Recordset
    Dim lng����id As Long
    
    If mpatiInfo.����ID = 0 Then
        lng����id = 0
    Else
        lng����id = mpatiInfo.����ID
    End If
    
    If cboType.ListIndex < 0 Then
        int���� = 1
    Else
        int���� = IIf(cboType.ItemData(cboType.ListIndex) = 1, 1, 0)
    End If
    
    On Error GoTo errHandle
    '84217,���ϴ�,2015/4/22,��ʾָ����סԺ�ڼ���ɵ�Ԥ��
    If cboType.Text = "סԺѺ��" And chk����ʾ����Ѻ��.Value = 1 And cboPatiPage.ListIndex >= 0 Then
        strWhere = " And A.��ҳID= " & cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    
    If gbln��Ժ����ʾ Then
        strWhere = strWhere & _
                " And Exists (Select 1 From ��Ա�� C, ������Ա D, ���ű� E " & _
                " Where C.���� =A.����Ա���� And C.Id = D.��Աid And D.����id = E.Id And (E.վ�� = '" & gstrNodeNo & "' Or E.վ�� Is Null))"
    End If
            
    '������ʷ�ɿ���ϸ�嵥
    strSQL = _
    " Select Ltrim(To_Char(A.�տ�ʱ��,'YYYY-MM-DD')) as ����,A.NO as ���ݺ�,B.���� as ����, " & _
    " Ltrim(To_Char(A.���,'9,999,999,990.00')) as �ɿ���,A.���㷽ʽ as ����,A.����Ա���� as �տ��� " & _
    " From " & IIf(mblnNOMoved, "H", "") & "����Ѻ���¼ A,���ű� B" & _
    " Where A.����ID=B.ID(+) And  A.����ID=[1]  And A.�Ƿ�����=[2] " & _
    " And  Nvl(A.У�Ա�־, 0) = 0   " & strWhere & _
    " Order by A.�տ�ʱ�� Desc"
    
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����id, int����)
    mshList.Rows = 2: mshList.Cols = 2: mshList.Clear
    If Not rsMoney.EOF Then
        Set mshList.DataSource = rsMoney
        mshList.ColWidth(0) = 1350: mshList.ColAlignment(0) = 4
        mshList.ColWidth(1) = 1110: mshList.ColAlignment(1) = 4
        mshList.ColWidth(2) = 1200: mshList.ColAlignment(2) = 1
        mshList.ColWidth(3) = 1600: mshList.ColAlignment(3) = 7
        mshList.ColWidth(4) = 1000: mshList.ColAlignment(4) = 4
        mshList.ColWidth(5) = 1000: mshList.ColAlignment(5) = 1
    End If
    If mshList.Rows > 1 Then
        mshList.Row = 1: mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    End If
    
    '���˺�:24386,����һ����λ�Ĺ���
    If strBackNo <> "" Then
        lngRow = zlControl.MshGrdFindRow(mshList, strBackNo, 1)
        If lngRow > 0 Then
            mshList.Row = lngRow: mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
            If (Not mshList.RowIsVisible(mshList.Row)) Or ((mshList.Row + 1) * mshList.RowHeight(0)) + 50 > mshList.Height Then mshList.TopRow = mshList.Row
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtFact_GotFocus()
    zlControl.ControlSetFocus txtFact
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
    If mpatiInfo.����ID > 0 Then
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
        If CreatePublicPatient() Then
            txtPatient.ForeColor = gobjPublicPatient.GetPatiColor(mstr��������, True)
        End If
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
    zlControl.ControlSetFocus txtUnit
End Sub

Private Sub txt������_GotFocus()
    zlControl.ControlSetFocus txt������
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

Private Sub txt�ʺ�_GotFocus()
    zlControl.ControlSetFocus txt�ʺ�
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
'��ģ�
'���ܣ���ʼ�����סԺ�ٴ�������Ϣ
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

Private Function CancelBill(ByVal lngID As Long, ByVal strNO As String, ByVal blnCanDel As Boolean, _
                                        ByVal intInsure As Integer, ByVal bln��ӡ As Boolean, ByVal strNote As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ��ID��Ѻ�𵥾�ִ���˿��
    '���:lngID=����ID
    '        blnCanDel=�Ƿ�֧���˸����ʻ�
    '        intInsure=��������ʹ�õĸ����ʻ��ı������,��Ϊ0
    '        strNo=���ݺ�
    '        strNote=ժҪ
    '����:
    '����:
    '����:���˺�
    '����:2011-07-19 09:28:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, blnTrans As Boolean
    Dim lng����ID As Long
    Dim blnThirdCard As Boolean
    Dim cllStatusUpdate As Collection
    
    If mYJinfo.lng�����ID <> 0 Then
        lng����ID = zlDatabase.GetNextId("����Ԥ����¼")
        blnThirdCard = cboStyle.ItemData(cboStyle.ListIndex) = -1
        If Not blnThirdCard And mstr�˿����Ա <> "" Then
            strNote = mstr�˿����Ա & "ǿ������:" & Format(txtMoney.Text, "0.00") & "Ԫ"
        End If
    Else
    End If
    On Error GoTo errH
    'Id_In         ����Ѻ���¼.Id%Type,
    'ժҪ_In       ����Ѻ���¼.ժҪ%Type,
    '����Ա���_In ����Ѻ���¼.����Ա���%Type,
    '����Ա����_In ����Ѻ���¼.����Ա����%Type,
    '����id_In     ����Ѻ���¼.Id%Type := Null,
    'Ʊ�ݺ�_In     ����Ѻ���¼.ʵ��Ʊ��%Type := Null,
    '����id_In     Ʊ�����ü�¼.Id%Type := Null,
    '����״̬_In   Number := 0,
    '��������_In   Number := 0,
    '���ַ�ʽ_In   ����Ѻ���¼.���㷽ʽ%Type := Null
    strSQL = "zl_����Ѻ���¼_DELETE(" & lngID & ",'" & strNote & "','" & _
        UserInfo.��� & "','" & UserInfo.���� & "'," & IIf(lng����ID = 0, "NULL", lng����ID) & "," & _
        "'" & IIf(bln��ӡ, mstrRedFact, "") & "'," & IIf(bln��ӡ, IIf(mlng����ID > 0, mlng����ID, "Null"), 0) & _
        IIf(blnThirdCard And gbln���ý����첽����, ",1)", ",0," & IIf(cboStyle.ItemData(cboStyle.ListIndex) <> -1, 1, 0) & ",'" & cboStyle.Text & "')")
    gcnOracle.BeginTrans: blnTrans = True
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If blnThirdCard And gbln���ý����첽���� Then
        '����У�Ա�־�������˿�
        strSQL = "zl_����Ѻ���¼_DELETE(" & lngID & ",'" & strNote & "','" & _
            UserInfo.��� & "','" & UserInfo.���� & "'," & IIf(lng����ID = 0, "NULL", lng����ID) & "," & _
            "'" & IIf(bln��ӡ, mstrRedFact, "") & "'," & IIf(bln��ӡ, IIf(mlng����ID > 0, mlng����ID, "Null"), 0) & ",2)"
        Set cllStatusUpdate = New Collection
        zlAddArray cllStatusUpdate, strSQL
    End If
    
    '����ҽ���ӿ�
    If intInsure <> 0 And blnCanDel Then
        If Not gclsInsure.TransferDelSwap(lngID, intInsure) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    End If

    If blnThirdCard Then
        If zlDepositDel(lngID, lng����ID, StrToNum(txtMoney.Text), strNO, cllStatusUpdate, blnTrans) = False Then
    
            Exit Function
        End If
    End If
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    
    
    If Not gblnBillԤ�� And bln��ӡ And mstrRedFact <> "" Then
        '��ɢ�����浱ǰ����
        zlDatabase.SetPara "��ǰԤ��Ʊ�ݺ�", mstrRedFact, glngSys, mlngFactModule
    End If
    CancelBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitIDKind()
    Dim strKind As String
    strKind = "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;ס|סԺ��|0;��|���ۺ�|0;��|���￨|0;��|�ֻ���|0"
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKind, txtPatient)
    mtySquareCard.blnExistsObjects = Not gobjSquare.objSquareCard Is Nothing
End Sub

Private Function zlCheckDepositDelValied(ByRef lngѺ��ID As Long, _
    ByVal dbl�˿��� As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷ѽ��׽ӿ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExend As String
    Dim cllSquareBalance As Collection
    
    If mYJinfo.lng�����ID = 0 Or cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then zlCheckDepositDelValied = True: Exit Function
    If Not mtySquareCard.blnExistsObjects Or gobjSquare.objSquareCard Is Nothing Then
            MsgBox "ע��:" & vbCrLf & _
                         "      ��ǰ��Ѻ��" & mYJinfo.str���� & " �����,�������ڲ�������ز���,�����˿�,����ϵͳ����Ա��ϵ!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
            Exit Function
    End If
    
    Set cllSquareBalance = New Collection
    'Array(�����ID,���ѿ�ID,ˢ�����, ����,����,�������,�Ƿ�����,ʣ��δ�˽��)
    cllSquareBalance.Add Array(mYJinfo.lng�����ID, 0, 0, mYJinfo.str����, "", "", False, dbl�˿���)
    'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, bln���ѿ� As Boolean, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, ByVal strSwapNo As String, _
    ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ����˽���ǰ�ļ��
    '���:frmMain-���õ�������
    '       lngModule-���õ�ģ���
    '       lngCardTypeID-�����ID
    '       strCardNo-����
    '       strBalanceIDs   String  In  ����֧�����漰�Ľ���ID ��ʽ:�շ�����|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                                   �շ�����: 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '       dblMoney-�˿���
    '       strSwapNo-������ˮ��(�˿�ʱ���)
    '       strSwapMemo-����˵��(�˿�ʱ����)
    '       strXMLExpend    XML IN  ��ѡ����(��չ��).��δ����
    '����:�˿�Ϸ�,����true,���򷵻�Flase
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, mYJinfo.lng�����ID, False, mYJinfo.str����, _
        "8|" & lngѺ��ID, dbl�˿���, mYJinfo.str������ˮ��, mYJinfo.str����˵��, strXMLExend) = False Then
          zlCheckDepositDelValied = False
          Exit Function
     End If
     '100610:���ϴ�,2016/10/13��Ԥ���˿��Ƿ���֤ˢ��
     If mYJinfo.bln�˿��鿨 Then
        '   zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln���ѿ� As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl��� As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln�˷� As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln���� As Boolean = False, _
        Optional ByVal bln�����ֹ As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal blnתԤ�� As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "") As Boolean
        '       strXmlIn-����������XML���,Ŀǰ��ʽ����:
        '       <IN>
        '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
        '       </IN>
        
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, mYJinfo.lng�����ID, False, _
            Trim(txtPatient.Text), mstrPatiSex, mstrPatiOld, dbl�˿���, mstrBrushCardNo, mstrbrPassWord, _
            True, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
        If mYJinfo.str���� <> mstrBrushCardNo Then
            MsgBox "ע��:" & vbCrLf & _
                         "      ��ǰ����[" & mstrBrushCardNo & "]��ԭ���׿���[" & mYJinfo.str���� & "]��һ�£���ʹ��ԭ������!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If
     
goEnd:
    zlCheckDepositDelValied = True
    Exit Function
End Function

Private Function zlDepositDel(ByRef lngѺ��ID As Long, ByRef lng����ID As Long, ByVal dblMoney As Double, ByVal strNO As String, ByVal cllStatusUpdate As Collection, _
                                ByRef blnTrans As Boolean, Optional blnBeenErr As Boolean, Optional ByVal blnReCancel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ���Ԥ������
    '��Σ� lngѺ��ID-����Ѻ���¼.ID��blnReCancel-���ˣ�blnTrans-��ǰ����״̬��blnBeenErr-�Ƿ�����쳣
    '���أ��ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strXMLExpend As String, strErrMsg As String, intState As Integer
    Dim rsBalance_Out As ADODB.Recordset, rsExpend_Out As ADODB.Recordset, cllThird As Collection, cllThirdExpend As Collection
    
    Err = 0: On Error GoTo Errhand:
    If mYJinfo.lng�����ID = 0 Or cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then zlDepositDel = True: Exit Function
    
    If gbln���ý����첽���� Then
        If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
        'ִ��״̬���¹���(���������ύ)
        If Not blnTrans Then gcnOracle.BeginTrans: blnTrans = True
        zlExecuteProcedureArrAy cllStatusUpdate, Me.Caption, blnTrans, blnTrans
    End If
    
    'Public Function zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long,bln���ѿ� as boolean ByVal strCardNo As String, _
        ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
        ByRef strSwapExtendInfor As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '����:�ʻ��ۿ���˽���
        '���:frmMain-���õ�������
        '       lngModule-���õ�ģ���
        '       lngCardTypeID-�����ID:ҽ�ƿ����.ID
        '       strCardNo-����
        '       strBalanceIDs-����֧�����漰�Ľ���ID(����ԭ����ID):
        '                           ��ʽ:�շ�����(|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
        '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
        '       dblMoney-�˿���
        '       strSwapNo-������ˮ��(�ۿ�ʱ�Ľ�����ˮ��)
        '       strSwapMemo-����˵��(�ۿ�ʱ�Ľ���˵��)
        '       strSwapExtendInfor-�����˷ѵĳ���ID��
        '                           ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
        '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
        '    blnResolveXMLToRecord-�Ƿ����XML������¼��(rsBalance_Out,rsExpend_Out��
        '����: strSwapNo-������ˮ��(�˿����ˮ��)
        '      strSwapMemo-����˵��(�˿��˵��)
        '    intStatus_Out-����״̬:�ӿڷ���Falseʱ���˲�����Ч: 0-���׵���ʧ��;1-�������ڴ�����
        '    strErrMsg_Out-������Ϣ:Ϊ��ʱ������ʾ���ǿ�ʱ����ʾ
        '       strSwapExtendInfor-���׵���չ��Ϣ
        '           ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n ÿ����Ŀ�в��ܰ���|�ַ�
        '����:��������    True:���óɹ�,False:����ʧ��
     strSwapNO = mYJinfo.str������ˮ��: strSwapMemo = mYJinfo.str����˵��
     '81489,Ƚ����,2015-4-29,�˷Ѵ������ID
     strSwapExtendInfor = "8|" & lng����ID
     strXMLExpend = GetExpendInfo(lngѺ��ID, True, dblMoney)
     If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, mYJinfo.lng�����ID, False, mYJinfo.str����, _
        "8|" & lngѺ��ID, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor, strXMLExpend, True, rsBalance_Out, rsExpend_Out, intState, strErrMsg) = False Then
        
        'ɾ����Ч��Ԥ������
        'intStateΪ0ʱ����������ɾ��ԭʼ��¼��Ϊ1ʱ����������
        If gbln���ý����첽���� Then
            If intState = 1 Then
                gcnOracle.RollbackTrans: blnTrans = False
                MsgBox "�������������ڽ����У��������쳣�˿��[" & strNO & "]" & IIf(strErrMsg <> "", "��" & vbCrLf & "������Ϣ���£�" & vbCrLf, "��") & _
                    strErrMsg, vbInformation, gstrSysName
                 blnBeenErr = True: Exit Function
            Else
                '���˲���Ѻ���¼��
                gcnOracle.RollbackTrans: blnTrans = False
                'ɾ��ԭʼ����
                strSQL = GetDeleteSQL(strNO, 1)
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                
                If blnReCancel Then
                    MsgBox "����[" & strNO & "]�������˿��ʧ�ܣ�������ִ���˿����" & IIf(strErrMsg <> "", "��" & vbCrLf & "������Ϣ���£�" & vbCrLf, "��") & _
                            strErrMsg, vbInformation, gstrSysName
                Else
                    MsgBox "����������ʧ�ܣ����Ժ�����" & IIf(strErrMsg <> "", "��" & vbCrLf & "������Ϣ���£�" & vbCrLf, "��") & _
                        strErrMsg, vbInformation, gstrSysName
                End If
                Exit Function
            End If
        Else
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
    End If
    
    If Not rsBalance_Out Is Nothing Then
         If rsBalance_Out.RecordCount = 0 Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "����ʧ�ܣ��ӿڵ���ʧ�ܣ�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        If rsBalance_Out.RecordCount > 1 Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "����ʧ�ܣ���֧�ֶ��ֽ��㷽ʽ֧����", vbInformation, gstrSysName
            Exit Function
        End If
        If Val(rsBalance_Out!���׽��) <> dblMoney Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "����ʧ�ܣ�ʵ���˿����뱾���˿������", vbInformation, gstrSysName
            Exit Function
        End If
        '��ȡ����Sql
        Set cllThird = New Collection
        If GetYJThirdUpdateSQL(lng����ID, "", Nvl(rsBalance_Out!���㷽ʽ), 0, "", "", "", "", Nvl(rsBalance_Out!�Ƿ���ͨ����, 0), cllThird, True) Then
    
            zlExecuteProcedureArrAy cllThird, Me.Caption, True, True
        End If
    End If
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    zlDepositDel = True

    If rsExpend_Out Is Nothing Then
        If Save��������(lng����ID, mYJinfo.lng�����ID, mYJinfo.str����, strSwapNO, strSwapMemo, _
            strSwapExtendInfor, blnTrans, True, "8|" & lng����ID) = False Then Exit Function
    Else
        '��ȡ��չ��ϢSql
        If zlGetThreeSwapExpendSQL(mYJinfo.lng�����ID, lng����ID, mYJinfo.str����, rsExpend_Out, cllThirdExpend) Then
            '�������񣬲�Ӱ���������ݱ���
            On Error GoTo ErrExpend:
            zlExecuteProcedureArrAy cllThirdExpend, Me.Caption
        End If
    End If

    Exit Function
Errhand:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter = 1 Then
        Resume
    End If
    Exit Function
ErrExpend:
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub Load֧����ʽ(Optional ByVal bln�������ַ�ʽ As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:���˺�
    '����:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objCard As Card
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim str���� As String, strErrMsg As String
    Dim strRQCardTypeIDs As String, objPayMode As Cards
    
    If InStr(1, mstrPrivs, ";Ѻ���տ�;") > 0 Then
        str���� = "1,2,7,8"
    End If
    If bln�������ַ�ʽ Then str���� = "1,2"
    If str���� = "" Then str���� = "1,2,7,8"
    
    On Error GoTo errHandle
    Set rsTemp = Get���㷽ʽ("Ԥ����", str����)
    'zlGetCards����ȡ��Ч�Ŀ�����
    '���:bytType
    '                   0-����ҽ�ƿ�
    '                   1-���õ�ҽ�ƿ�
    '                   2-���д��������˻���������
    '                   3-���õ������˻���ҽ�ƿ�
    Set objPayMode = gobjSquare.objSquareCard.zlGetCards(3)
    Set mobjPayMode = New Collection
    
    With cboStyle
        .Clear: strRQCardTypeIDs = ""
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 1 To objPayMode.Count
                If objPayMode(i).���㷽ʽ = Nvl(rsTemp!����) Then
                    blnFind = True
                    Exit For
                End If
            Next
            '104083:���ϴ���2016/12/21�������˻��������̬����
            '����Ϊ8�ĸ�������ҽ�ƿ�������
            If Not blnFind And InStr(",3,8,", "," & rsTemp!���� & ",") = 0 Then
                .AddItem Nvl(rsTemp!����)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!����))
                Call MakeCardsFrom���㷽ʽ(Nvl(rsTemp!����), Nvl(rsTemp!����))
                If rsTemp!ȱʡ = 1 Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                If mstrȱʡ���㷽ʽ = Nvl(rsTemp!����) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 1 To objPayMode.Count
            rsTemp.Filter = "���� ='" & objPayMode(i).���㷽ʽ & "'"
            If Not rsTemp.EOF Then
                If str���� <> 5 And Not objPayMode(i).���ѿ� Then
                    .AddItem objPayMode(i).����: .ItemData(.NewIndex) = -1
                    Call MakeCardsFrom���㷽ʽ(objPayMode(i).����, objPayMode(i).���㷽ʽ, objPayMode(i).�ӿ����)
                    If mstrȱʡ���㷽ʽ = objPayMode(i).���� Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                    If objPayMode(i).�Ƿ�֧��ɨ�븶 Then
                        strRQCardTypeIDs = strRQCardTypeIDs & "," & objPayMode(i).�ӿ����
                    End If
                End If
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    If bln�������ַ�ʽ Then Exit Sub
    If cboStyle.ListCount = 0 Then
        MsgBox "Ԥ������û�п��õĽ��㷽ʽ,���ȵ����㷽ʽ���������á�", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    If strRQCardTypeIDs <> "" Then strRQCardTypeIDs = Mid(strRQCardTypeIDs, 2)
    '��ʼ��ɨ��ؼ�
    If btQRCodePay.zlInit(Me, strRQCardTypeIDs, glngSys, mlngModul, gcnOracle, gstrDBUser, strErrMsg) = False Then strRQCardTypeIDs = ""
    btQRCodePay.Tag = strRQCardTypeIDs
    btQRCodePay.Visible = strRQCardTypeIDs <> "" And mbytInState = EM_��Ѻ��
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Save��������(ByVal lngѺ��ID As Long, ByVal lng�����ID As Long, _
    ByVal str���� As String, str������ˮ�� As String, str����˵�� As String, strExpend As String, _
    blnTrans As Boolean, Optional bln��Ѻ�� As Boolean = False, Optional strExpendOld As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, varData As Variant, varTemp As Variant, cllPro As Collection, i As Long
     
    Err = 0: On Error GoTo Errhand:
    If bln��Ѻ�� = False Then
        '�˷�ʱ,�����Ľ���
        '���½�����Ϣ
         '    Zl_�����ӿڸ���_Update
        strSQL = "Zl_�����ӿڸ���_Update("
        '  �����id_In   ����Ԥ����¼.�����id%Type,
        strSQL = strSQL & "" & lng�����ID & ","
        '  ���ѿ�_In     Number,
        strSQL = strSQL & "0,"
        '  ����_In       ����Ԥ����¼.����%Type,
        strSQL = strSQL & "'" & str���� & "',"
        '  ����ids_In    Varchar2,
        strSQL = strSQL & "'" & lngѺ��ID & "',"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
        strSQL = strSQL & "'" & str������ˮ�� & "',"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type
        strSQL = strSQL & "'" & str����˵�� & "',"
        'Ԥ����ɿ�_In Number := 0
        strSQL = strSQL & "" & 2 & ","
        '�˷ѱ�־ :1-�˷�;0-����
        strSQL = strSQL & "" & IIf(bln��Ѻ��, 1, 0) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    '���ύ,�����������,�ٸ�����صĽ�����Ϣ
    'strExpend:������չ��Ϣ,��ʽ:��Ŀ����|��Ŀ����||...
    varData = Split(strExpend, "||")
    Dim str������Ϣ As String, strTemp As String
    Set cllPro = New Collection
    If strExpendOld <> strExpend Then
        For i = 0 To UBound(varData)
            If Trim(varData(i)) <> "" Then
                varTemp = Split(varData(i) & "|", "|")
                If varTemp(0) <> "" Then
                    strTemp = varTemp(0) & "|" & varTemp(1)
                    If zlCommFun.ActualLen(str������Ϣ & "||" & strTemp) > 2000 Then
                        str������Ϣ = Mid(str������Ϣ, 3)
                        'Zl_�������㽻��_Insert
                        strSQL = "Zl_�������㽻��_Insert("
                        '�����id_In ����Ԥ����¼.�����id%Type,
                        strSQL = strSQL & "" & lng�����ID & ","
                        '���ѿ�_In   Number,
                        strSQL = strSQL & "0,"
                        '����_In     ����Ԥ����¼.����%Type,
                        strSQL = strSQL & "'" & str���� & "',"
                        '����ids_In  Varchar2,
                        strSQL = strSQL & "'" & lngѺ��ID & "',"
                        '������Ϣ_In Varchar2:������Ŀ|��������||...
                        strSQL = strSQL & "'" & str������Ϣ & "',"
                        'Ԥ����ɿ�_In Number := 0
                        strSQL = strSQL & "2,"
                        '���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null
                        strSQL = strSQL & "Null,"
                        'Ԥ��id_In     ����Ԥ����¼.Id%Type := Null
                        strSQL = strSQL & "Null,"
                        '����_In       �������㽻��.����%Type := Nul
                        strSQL = strSQL & "2)"
                        zlAddArray cllPro, strSQL
                        str������Ϣ = ""
                    End If
                    str������Ϣ = str������Ϣ & "||" & strTemp
                End If
            End If
        Next
        
        If str������Ϣ <> "" Then
            str������Ϣ = Mid(str������Ϣ, 3)
            'Zl_�������㽻��_Insert
            strSQL = "Zl_�������㽻��_Insert("
            '�����id_In ����Ԥ����¼.�����id%Type,
            strSQL = strSQL & "" & lng�����ID & ","
            '���ѿ�_In   Number,
            strSQL = strSQL & "0,"
            '����_In     ����Ԥ����¼.����%Type,
            strSQL = strSQL & "'" & str���� & "',"
            '����ids_In  Varchar2,
            strSQL = strSQL & "'" & lngѺ��ID & "',"
            '������Ϣ_In Varchar2:������Ŀ|��������||...
            strSQL = strSQL & "'" & str������Ϣ & "',"
            'Ԥ����ɿ�_In Number := 0
            strSQL = strSQL & "2,"
            '���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null
            strSQL = strSQL & "Null,"
            'Ԥ��id_In     ����Ԥ����¼.Id%Type := Null
            strSQL = strSQL & "Null,"
            '����_In       �������㽻��.����%Type := Nul
            strSQL = strSQL & "2)"
            zlAddArray cllPro, strSQL
        End If
    End If
    Err = 0: On Error GoTo ErrOthers: blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption
    blnTrans = False
    Save�������� = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrOthers:
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    '    �ܱ������,�����
     Call ErrCenter
End Function

Private Function zlInterfacePrayMoney(ByVal lngѺ��ID As Long, ByVal strNO As String, ByVal dblMoney As Double, _
                                       ByVal cllStatusUpdate As Collection, ByRef blnTrans As Boolean, Optional blnBeenErr As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ӿ�֧�����
    '����: intState-���׽ӿڷ���ʧ�ܺ�Ľ���״̬��0-ʧ�ܣ�1-���ڽ���
    '����:֧���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSQL As String
    Dim rsBalance_Out As ADODB.Recordset, rsExpend_Out As ADODB.Recordset, strErrMsg As String, intState As Integer
    Dim dbl���׽�� As Double, strTmp As String
    Dim cllThird As Collection, cllThirdExpend As Collection
    
    'intState 0-ʧ�ܣ�1-���ڽ���
    If cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    If mlngCardTypeID = 0 Then zlInterfacePrayMoney = True: Exit Function
    
    If gbln���ý����첽���� Then
        If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
        'ִ��״̬���¹���(���������ύ)
        If Not blnTrans Then gcnOracle.BeginTrans: blnTrans = True
        zlExecuteProcedureArrAy cllStatusUpdate, Me.Caption, blnTrans, blnTrans
    End If
    
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln���ѿ� As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ��
    '���:frmMain-���õ�������
    '        lngModule-����ģ���
    '        strBalanceIDs-����ID,����ö��ŷ���
    '        strPrepayNos-��Ԥ��ʱ��Ч. Ԥ�����ݺ�,����ö��ŷ���
    '       strCardNo-����
    '       dblMoney-֧�����
    '����:strSwapGlideNO-������ˮ��
    '       strSwapMemo-����˵��
    '       strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '����:�ۿ�ɹ�,����true,���򷵻�Flase
    '˵��:
    '   ��������Ҫ�ۿ�ĵط����øýӿ�,Ŀǰ�滮��:�շ��ң��Һ���;������ѯ��;ҽ������վ��ҩ���ȡ�
    '   һ����˵���ɹ��ۿ�󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
    '   �ڿۿ�ɹ��󣬷��ؽ�����ˮ�ź���ر�ע˵���������������������Ϣ�����Է��ڽ���˵�����Ա��˷�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strSwapExtendInfor = "" & _
                                "<IN>" & vbCrLf & _
                                "       <QRCODE>" & mstrQRcode & "</QRCODE>" & vbCrLf & _
                                "       <SFYJ>" & 1 & "</SFYJ>" & vbCrLf & _
                                "</IN>"
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, mlngCardTypeID, False, mstrBrushCardNo, "", strNO, _
        dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor, True, rsBalance_Out, rsExpend_Out, intState, strErrMsg) = False Then
         
        'ɾ����Ч��Ԥ������
        'intStateΪ0ʱ����������ɾ��ԭʼ��¼��Ϊ1ʱ����������
        If gbln���ý����첽���� Then
            If intState = 1 Then
                gcnOracle.RollbackTrans: blnTrans = False
                MsgBox "�������������ڽ����У��������쳣����[" & strNO & "]" & IIf(strErrMsg <> "", "��" & vbCrLf & "������Ϣ���£�" & vbCrLf, "��") & _
                    strErrMsg, vbInformation, gstrSysName
                blnBeenErr = True
                Exit Function
            Else
                '������Ա��Ԥ����������
                gcnOracle.RollbackTrans: blnTrans = False
                'ɾ��ԭʼ����
                If mbytInState = EM_��Ѻ�� Then
                    strSQL = GetDeleteSQL(strNO)
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                End If
                If mbytInState = EM_�쳣���� Then
                    strTmp = "����������ʧ�ܣ���ɾ�����쳣���ݡ�"
                Else
                    strTmp = "����������ʧ�ܣ����Ժ�����" & IIf(strErrMsg <> "", "��" & vbCrLf & "������Ϣ���£�" & vbCrLf, "��") & _
                                    strErrMsg
                End If
                MsgBox strTmp, vbInformation, gstrSysName
                Exit Function
            End If
        Else
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
    End If
    
    '����ʵ�ʷ��ؽ�����Ϣ����
    If rsBalance_Out Is Nothing Then
        'ԭ�нӿڷ�ʽ
        If Save��������(lngѺ��ID, mlngCardTypeID, mstrBrushCardNo, strSwapGlideNO, strSwapMemo, _
            strSwapExtendInfor, blnTrans) = False Then Exit Function
        zlInterfacePrayMoney = True
        Exit Function
    Else
        If rsBalance_Out.RecordCount = 0 Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "����ʧ�ܣ��ӿڵ���ʧ�ܣ�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        If rsBalance_Out.RecordCount > 1 Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "����ʧ�ܣ���֧�ֶ��ֽ��㷽ʽ֧����", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        dbl���׽�� = Nvl(rsBalance_Out!���׽��, 0)
        If dbl���׽�� = 0 Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "����ʧ�ܣ���ǰ���׽��Ϊ0������ ��", vbInformation, gstrSysName:
            Exit Function
        End If
        If dbl���׽�� <> 0 Then sta.Panels(2).Text = "���ν��׽��:" & dbl���׽��
        
        If dbl���׽�� < dblMoney Then
            MsgBox "ע��:" & vbCrLf & _
                                 "���ν��׽��Ϊ" & Format(dbl���׽��, "0.00") & "Ԫ����ԭ�ɿ���" & Format(dblMoney, "0.00") & "��һ�£�", vbInformation, gstrSysName
            dblMoney = Format(dbl���׽��, "0.00")
            lblRepairMoney.Visible = True
        End If
        
        '�ж�Ԥ�����Ƿ񳬳�ˢ�������
        If lblRepairMoney.Visible Then
            lblRepairMoney.Caption = "������:" & Format((CDbl(txtMoney.Text) - dblMoney), "###0.00;-###0.00;;")
            If lblMoney.Tag <> "" Then lblRepairMoney.Caption = "������:" & Format((CDbl(lblMoney.Tag) - dblMoney), "###0.00;-###0.00;;")
            lblMoney.Tag = ""
            txtMoney.Text = Format(dblMoney, "###0.00;-###0.00;;")
        End If
        
        '��ȡ����Sql
        Set cllThird = New Collection
        If GetYJThirdUpdateSQL(lngѺ��ID, IIf(Nvl(rsBalance_Out!����) = "", mstrBrushCardNo, Nvl(rsBalance_Out!����)), Nvl(rsBalance_Out!���㷽ʽ), Nvl(rsBalance_Out!���׽��, 0), Nvl(rsBalance_Out!�������), _
                            strSwapGlideNO, strSwapMemo, Nvl(rsBalance_Out!����ժҪ), Nvl(rsBalance_Out!�Ƿ���ͨ����, 0), cllThird) Then
            
            '�ύ����
            zlExecuteProcedureArrAy cllThird, Me.Caption, False, True
            blnTrans = False
            zlInterfacePrayMoney = True
        End If
        '��ȡ��չ��ϢSql
        If zlGetThreeSwapExpendSQL(mlngCardTypeID, CStr(lngѺ��ID), IIf(Nvl(rsBalance_Out!����) = "", mstrBrushCardNo, Nvl(rsBalance_Out!����)), rsExpend_Out, cllThirdExpend) Then
            '�������񣬲�Ӱ���������ݱ���
            On Error GoTo ErrExpend:
            zlExecuteProcedureArrAy cllThirdExpend, Me.Caption
        End If
        If dblMoney <> Nvl(rsBalance_Out!���׽��, 0) Then txtMoney.Text = Nvl(rsBalance_Out!���׽��, 0)
    End If

    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrExpend:
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetThreeSwapExpendSQL(ByVal lng�����ID As Long, ByVal strѺ��IDs As String, ByVal str���� As String, _
                                        ByVal rsExpend As ADODB.Recordset, ByRef cllTirdExpend As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��չ��Ϣ�����SQL������
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:
    '����:2018-03-27 17:33:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str������Ϣ As String, strTemp As String
    On Error GoTo errHandle
    
    If cllTirdExpend Is Nothing Then Set cllTirdExpend = New Collection
    If rsExpend Is Nothing Then zlGetThreeSwapExpendSQL = True: Exit Function
    If rsExpend.State <> 1 Then zlGetThreeSwapExpendSQL = True: Exit Function
    
    With rsExpend
        rsExpend.Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        str������Ϣ = ""
        Do While Not .EOF
            If Nvl(!��Ŀ����) <> "" Then
                strTemp = Nvl(!��Ŀ����) & "|" & Nvl(!��Ŀ����)
                If zlCommFun.ActualLen(str������Ϣ & "||" & strTemp) > 2000 Then
                        str������Ϣ = Mid(str������Ϣ, 3)
                        'Zl_�������㽻��_Insert
                        strSQL = "Zl_�������㽻��_Insert("
                        '�����id_In ����Ԥ����¼.�����id%Type,
                        strSQL = strSQL & "" & lng�����ID & ","
                        '���ѿ�_In   Number,
                        strSQL = strSQL & "" & 0 & ","
                        '����_In     ����Ԥ����¼.����%Type,
                        strSQL = strSQL & "'" & str���� & "',"
                        '����ids_In  Varchar2,
                        strSQL = strSQL & "'" & strѺ��IDs & "',"
                        '������Ϣ_In Varchar2:������Ŀ|��������||...
                        strSQL = strSQL & "'" & str������Ϣ & "',"
                        'Ԥ����ɿ�_In Number := 0
                        strSQL = strSQL & "2,"
                        '���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null
                        strSQL = strSQL & "Null,"
                        'Ԥ��id_In     ����Ԥ����¼.Id%Type := Null
                        strSQL = strSQL & "Null,"
                        '����_In       �������㽻��.����%Type := Nul
                        strSQL = strSQL & "2)"
                        zlAddArray cllTirdExpend, strSQL
                        str������Ϣ = ""
                End If
                str������Ϣ = str������Ϣ & "||" & strTemp
            End If
            .MoveNext
        Loop
        
    End With
    If str������Ϣ <> "" Then
        str������Ϣ = Mid(str������Ϣ, 3)
        'Zl_�������㽻��_Insert
        strSQL = "Zl_�������㽻��_Insert("
        '�����id_In ����Ԥ����¼.�����id%Type,
        strSQL = strSQL & "" & lng�����ID & ","
        '���ѿ�_In   Number,
        strSQL = strSQL & "" & 0 & ","
        '����_In     ����Ԥ����¼.����%Type,
        strSQL = strSQL & "'" & str���� & "',"
        '����ids_In  Varchar2,
        strSQL = strSQL & "'" & strѺ��IDs & "',"
        '������Ϣ_In Varchar2:������Ŀ|��������||...
        strSQL = strSQL & "'" & str������Ϣ & "',"
        'Ԥ����ɿ�_In Number := 0
        strSQL = strSQL & "2,"
        '���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null
        strSQL = strSQL & "Null,"
        'Ԥ��id_In     ����Ԥ����¼.Id%Type := Null
        strSQL = strSQL & "Null,"
        '����_In       �������㽻��.����%Type := Nul
        strSQL = strSQL & "2)"
        zlAddArray cllTirdExpend, strSQL
    End If
    zlGetThreeSwapExpendSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub zlCheckFactIsEnough()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰƱ���Ƿ�����
    '����:���˺�
    '����:2012-09-06 15:41:52
    '˵��:
    '����:37372
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngʣ������ As Long, strType As String
    If mbytInState = EM_������� Or mbytInState = EM_��Ѻ�� Then Exit Sub
    '��Ҫ���ʣ�������Ƿ����:
    If cboType.ListIndex < 0 Then
        strType = ""
    Else
        strType = cboType.ItemData(cboType.ListIndex)
    End If
    If zlCheckInvoiceOverplusEnough(2, gint����ʣ��Ʊ������, lngʣ������, mlng����ID, strType) = False Then
        MsgBox "ע��:" & vbCrLf & _
               "    ��ǰʣ��Ʊ��(" & lngʣ������ & ") С���˱���������(" & gint����ʣ��Ʊ������ & "),��ע�������Ʊ!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
    End If
End Sub
Private Sub LoadPatiPage(ByVal lng����id As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز��˵�סԺ����
    '����:���˺�
    '����:2012-12-11 10:19:58
    '˵��:
    '����:51628
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim bln���� As Boolean
    On Error GoTo errHandle
        
    cboPatiPage.Clear
    With cboPatiPage
        
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
            If Val(Nvl(rsTemp!��ҳID)) = mpatiInfo.��ҳID Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
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
    Dim str����id As String, PatiPageInfo As New clsPatientInfo
    Dim lng����id As Long, lng��ҳID As Long
    
    On Error GoTo errHandle
    If mblnδ��Ʋ���Ԥ�� = False Then Checkδ��Ʋ���Ԥ�� = True: Exit Function
    '����Ԥ�������
    If cboType.ItemData(cboType.ListIndex) <> 2 Then Checkδ��Ʋ���Ԥ�� = True: Exit Function
    '��ǰסԺ������Ϊ��Ժ��,Ҳ�����
    If Not mpatiInfo.��Ժ Then Checkδ��Ʋ���Ԥ�� = True: Exit Function
    lng����id = mpatiInfo.����ID
    '������סԺ������,Ҳ�ܽ�Ԥ��,��˲����
    If cboPatiPage.ListIndex < 0 Then Checkδ��Ʋ���Ԥ�� = True: Exit Function
    lng��ҳID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    str����id = lng����id & ":" & lng��ҳID
    Call GetPatiPageInforByID(str����id, PatiPageInfo, False)
    If PatiPageInfo.����� = False Then
        MsgBox "ע��" & vbCrLf & "   ���ˡ�" & mpatiInfo.���� & "��δ���,�������Ѻ��!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Checkδ��Ʋ���Ԥ�� = True
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
    Dim str����Ա���� As String, strTittle
    On Error GoTo errHandle
    '��Ȩ�ޣ���Ϊ�շ�״̬
    If mbytInState <> EM_��Ѻ�� Then Exit Function
    'type: 1-�쳣��ֵ��2-�쳣����
    strSQL = "Select Type, No , ���� ,����Ա����" & vbNewLine & _
            "From (Select 1 Type, a.No, a.����, a.����Ա����" & vbNewLine & _
            "       From ����Ѻ���¼ a" & vbNewLine & _
            "       Where Nvl(У�Ա�־, 0) <> 0 And ��¼״̬ = 0 And ����id = [1]" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 2 Type, a.No, a.����, a.����Ա����" & vbNewLine & _
            "       From ����Ѻ���¼ a" & vbNewLine & _
            "       Where Nvl(У�Ա�־, 0) <> 0 And ����id = [1] And ��¼״̬ = 2)" & vbNewLine & _
            "Order By Decode(����Ա����, [2], 0, 1), Type"
    Set rsErrBills = zlDatabase.OpenSQLRecord(strSQL, "�����쳣���ݲ�ѯ", lng����id, UserInfo.����)
    If rsErrBills.EOF Then Exit Function
    
    str����Ա���� = Nvl(rsErrBills!����Ա����)
    If Nvl(rsErrBills!type) = 1 Then
        strTittle = "�տ�"
    ElseIf Nvl(rsErrBills!type) = 2 Then
        strTittle = "����"
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
    mstrInNO = Nvl(rsErrBills!NO)
    mstrBrushCardNo = Nvl(rsErrBills!����)
    If Nvl(rsErrBills!type) = 1 Then
        mbytInState = EM_�쳣����
    Else
        mbytInState = EM_�쳣����
    End If
    Call InitFace
    '��ʼ��������Ϣ
    Call InitPatientInfo(mstrInNO)
    If mbytInState = EM_�쳣���� Then txtMoney.Text = Abs(txtMoney.Text)
    Call SetCtrlEnabled
    mblnOptErrBill = True
    OptOthersErrBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RestorePayStyle()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ָ����ϴ�ѡ���֧����ʽ
    '˵��:lblStyle.Tag��¼�����ϴ�ѡ���֧����ʽ
    '       cboStyle.Tag��¼����ȱʡ��֧����ʽ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    On Error GoTo errHandle
    
    If lblStyle.Tag = "" Then Exit Sub
    '���ϴ�ѡ���֧����ʽ,�ָ�
    intIndex = Val(lblStyle.Tag)
    lblStyle.Tag = ""
    If intIndex > cboStyle.ListCount - 1 Then cboStyle.ListIndex = Val(cboStyle.Tag): Exit Sub
    cboStyle.ListIndex = intIndex

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LocatePayStyle(ByVal lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݿ����ID,��λ��ָ����֧�������
    '���:lngCardTypeID-�����ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnFind As Boolean, i As Integer
    If lngCardTypeID = 0 Then Exit Function
    If mobjPayMode Is Nothing Then Exit Function
    If mobjPayMode.Count = 0 Then Exit Function
    With cboStyle
        For i = 1 To mobjPayMode.Count
            If mobjPayMode(i).�ӿ���� = lngCardTypeID Then
                cboStyle.ListIndex = cbo.FindIndex(cboStyle, mobjPayMode(i).����)
                blnFind = True: Exit For
            End If
        Next
    End With
    LocatePayStyle = blnFind
End Function

Private Sub LoadOriginReturnMoneyStyle(Optional ByVal blnȱʡ���� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز���ԭʼ�˿ʽ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mYJinfo.lng�����ID = 0 Or mYJinfo.str���� = "" Then Exit Sub
    cboStyle.AddItem mYJinfo.str����
    cboStyle.ItemData(cboStyle.NewIndex) = -1
    If Not blnȱʡ���� Then cboStyle.ListIndex = cboStyle.NewIndex
End Sub

Private Function ZlGetParaConfig(ByVal lng�����ID As Long, ByVal intPara As Long, _
                                                    Optional strErrMsg As String, Optional strExpend As String) As Boolean
    ZlGetParaConfig = gobjSquare.objSquareCard.ZlGetParaConfig(Me, lng�����ID, False, intPara, strErrMsg, strExpend)
End Function

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
    patiinfo.סԺ״̬ = 9
    If GetPatiPageInforByID(str����id, PatiPageInfo, blnLastTime) = False Then GetPatiInfo = True: Exit Function
    If PatiPageInfo.����ID > 0 Then
        patiinfo.סԺ״̬ = 0
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

Private Sub MakeCardsFrom���㷽ʽ(ByVal str���� As String, ByVal str���㷽ʽ As String, _
                 Optional ByVal lng�����ID As Long)
    '���ܣ����ݽ��㷽ʽ����Cards����
    Dim objCard As Card
    Set objCard = New Card
    If mobjPayMode Is Nothing Then Set mobjPayMode = New Collection
    objCard.���� = str����
    objCard.���㷽ʽ = str���㷽ʽ
    objCard.�ӿ���� = lng�����ID
    mobjPayMode.Add objCard, "_" & str����
End Sub


