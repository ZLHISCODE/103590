VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeposit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ԥ�����"
   ClientHeight    =   9270
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
   Icon            =   "frmDeposit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   11910
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2325
      Left            =   75
      ScaleHeight     =   2325
      ScaleWidth      =   11775
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1050
      Width           =   11775
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   18
         X1              =   1260
         X2              =   4845
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label lbl�ֻ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� ��"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   70
         Tag             =   "�� �� �� "
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label lblδ�ɷ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ�ɷ��� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   2640
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   66
         Tag             =   "������λ "
         Top             =   1560
         Width           =   960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         Index           =   16
         X1              =   6105
         X2              =   11640
         Y1              =   2160
         Y2              =   2160
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
         X1              =   3660
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
         Index           =   11
         X1              =   1245
         X2              =   2415
         Y1              =   1080
         Y2              =   1080
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
         Left            =   5100
         TabIndex        =   63
         Tag             =   "��    ע "
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ���� "
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7890
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   56
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
         TabIndex        =   55
         Tag             =   "�� �� �� "
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label lbl�ѱ�ȼ� 
         AutoSize        =   -1  'True
         Caption         =   "�ѱ� "
         Height          =   240
         Left            =   4965
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         Left            =   2640
         TabIndex        =   51
         Tag             =   "Ԥ����� "
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
         Tag             =   "δ����� "
         ToolTipText     =   "δ��˵Ļ��ۼ��˷��úϼ�"
         Top             =   795
         Width           =   1080
      End
      Begin VB.Label lbl�ʻ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ʻ���� "
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   47
         Tag             =   "�ʻ���� "
         Top             =   795
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.PictureBox picList 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   0
      ScaleHeight     =   1650
      ScaleWidth      =   11910
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   6660
      Width           =   11910
      Begin VB.CheckBox chk����ʾ����Ԥ�� 
         Caption         =   "����ʾ����Ԥ��"
         Height          =   240
         Left            =   9360
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ԥ���嵥"
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
         TabIndex        =   44
         Top             =   0
         Width           =   12015
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
         Height          =   1335
         Left            =   135
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
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
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   8310
      Width           =   11910
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   420
         Left            =   150
         TabIndex        =   34
         Top             =   60
         Width           =   1500
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   420
         Left            =   10335
         TabIndex        =   32
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   45
         Width           =   1500
      End
      Begin VB.CommandButton cmdSetup 
         Caption         =   "��ӡ����(&S)"
         Height          =   420
         Left            =   1770
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F10"
         Top             =   60
         Width           =   1620
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   420
         Left            =   8760
         TabIndex        =   28
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
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   0
      Width           =   11755
      Begin VB.TextBox txtFact 
         ForeColor       =   &H00C00000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   6300
         MaxLength       =   50
         TabIndex        =   29
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
         TabIndex        =   30
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
         TabIndex        =   39
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
         TabIndex        =   57
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
         TabIndex        =   35
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
         TabIndex        =   41
         Top             =   570
         Visible         =   0   'False
         Width           =   360
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
         TabIndex        =   45
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
         TabIndex        =   40
         Top             =   630
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   36
      Top             =   8910
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
            Picture         =   "frmDeposit.frx":08CA
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
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3480
      Width           =   11775
      Begin VB.CheckBox chkAllCash 
         Caption         =   "�����˻�ǿ������"
         Enabled         =   0   'False
         Height          =   240
         Left            =   9360
         TabIndex        =   3
         Top             =   195
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.ComboBox cboPatiPage 
         Height          =   360
         Left            =   3675
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   570
         Width           =   1335
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   585
         TabIndex        =   64
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
      Begin VB.CheckBox chk����temp 
         Caption         =   "��ʱ����"
         Enabled         =   0   'False
         Height          =   240
         Left            =   7995
         TabIndex        =   2
         Top             =   195
         Width           =   1335
      End
      Begin VB.ComboBox cboType 
         Height          =   360
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   570
         Width           =   1380
      End
      Begin VB.TextBox txtMan 
         Enabled         =   0   'False
         Height          =   360
         Left            =   7980
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2760
         Width           =   3705
      End
      Begin VB.TextBox txtCode 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7995
         MaxLength       =   30
         TabIndex        =   17
         Top             =   1440
         Width           =   3690
      End
      Begin VB.TextBox txtUnit 
         Height          =   360
         Left            =   7995
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1005
         Width           =   3690
      End
      Begin VB.TextBox txt�ʺ� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7980
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1890
         Width           =   3705
      End
      Begin VB.ComboBox cboUnit 
         Height          =   360
         Left            =   7995
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   585
         Width           =   3690
      End
      Begin VB.ComboBox cboNote 
         Height          =   360
         Left            =   1230
         TabIndex        =   23
         Text            =   "cboNote"
         Top             =   2325
         Width           =   10485
      End
      Begin VB.TextBox txt������ 
         Height          =   360
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   19
         Top             =   1890
         Width           =   3765
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1230
         TabIndex        =   1
         ToolTipText     =   "�ȼ���F11"
         Top             =   135
         Width           =   3765
      End
      Begin VB.ComboBox cboStyle 
         Height          =   360
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1440
         Width           =   3765
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   360
         Left            =   1230
         TabIndex        =   25
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
         TabIndex        =   11
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
      Begin VB.Label lblPatiPage 
         AutoSize        =   -1  'True
         Caption         =   "סԺ����"
         Height          =   240
         Left            =   2685
         TabIndex        =   6
         Top             =   615
         Width           =   960
      End
      Begin VB.Label lblRepairMoney 
         Caption         =   "������:"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   5010
         TabIndex        =   65
         Top             =   1050
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label lblԤ������ 
         AutoSize        =   -1  'True
         Caption         =   "Ԥ������"
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   630
         Width           =   960
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ɿ����"
         Height          =   240
         Left            =   6960
         TabIndex        =   8
         Top             =   645
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ʺ�"
         Height          =   240
         Left            =   7440
         TabIndex        =   20
         Top             =   1950
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   435
         TabIndex        =   18
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label lbl�ɿλ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ɿλ"
         Height          =   240
         Left            =   6960
         TabIndex        =   12
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
         TabIndex        =   10
         Top             =   1065
         Width           =   510
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   240
         Left            =   6960
         TabIndex        =   16
         Top             =   1500
         Width           =   960
      End
      Begin VB.Label lblStyle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "֧����ʽ"
         Height          =   240
         Left            =   195
         TabIndex        =   14
         Top             =   1500
         Width           =   960
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ"
         Height          =   240
         Left            =   645
         TabIndex        =   22
         Top             =   2385
         Width           =   480
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�տ�ʱ��"
         Height          =   240
         Left            =   195
         TabIndex        =   24
         Top             =   2820
         Width           =   960
      End
      Begin VB.Label lblMan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�տ�Ա"
         Height          =   240
         Left            =   7200
         TabIndex        =   26
         Top             =   2820
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'˵����
'1.�˿������ַ�ʽ,ȱʡ�ķ�ʽ���ڹ�������ָ���ĵ���ִ���˿�ܣ����������տ�״̬��ʹ���˿�ܣ���һ�ַ�ʽ
'���������տ�״̬�տ�,���տ����Ը�����ʾ�˿��ʱ(�˿���<=�������)�������ַ�ʽ����Ӱ�첡��Ԥ�����ͳ��

'��ڲ���----------------------------------------------------------------------------------
Private mbytInState As Byte '0-��Ԥ����(ȱʡ,���л�����),1-�������(1),2-����״̬(1); 3-����˿�(37770), 4-תԤ��
Private mstrInNO As String 'Ҫ������˿�ĵ��ݺ�(mbytInState=1��3ʱ��Ч),�Ӳ�����Ϣ�Ǽ��е����˿�ʱΪ��
Private mblnNOMoved As Boolean '����ϸʱ��¼��ǰѡ��ĵ����Ƿ����������ݱ���,����������ʱ�������ж�
Private mblnViewCancel As Boolean '�Ƿ�����˿��(mbytInState=1ʱ��Ч)
Private mstrPrivs As String
Private mlngModul As Long
Private mbytCallObject As Byte '���õĶ���(0-Ԥ����������;1-���˷��ò�ѯ����;2-ҽ�ƿ�...
Private mlng����ID As Long, mlng��ҳID As Long, mdblDefPreMoney As Double
Private mbytPrepayType As Byte   ' 1-����Ԥ��;2-סԺԤ��(4ʱ,1,����תסԺ;2ʱסԺת����)
Private mblnNotClick As Boolean
Private mstrbrPassWord As String
'�������----------------------------------------------------------------------------------
Private mblnUnLoad  As Boolean '���ڿ��ƴ���ֱ���˳�
Private mrsInfo As New ADODB.Recordset '������Ϣ(����ID,����,�Ա�,����,סԺ��,����,��Ժ��־)
Private mdblʣ���� As Double
Private mdblԤ����� As Double
Private mdbl������� As Double
Private mdblԤ�����_���� As Double, mdblԤ�����_�������� As Double
Private mlng����ID As Long, mstrCardPrivs As String
Private mstr���տ� As String
Private mstrRedFact As String
Private mstrȱʡ���㷽ʽ As String
Private mblnOK As Boolean, mstr�˿����Ա As String
Private mstrPrintDate As String
Private mblnδ��Ʋ���Ԥ�� As Boolean '51628
Private mblnסԺ��Ԥ����֤ As Boolean   '63113:������,2013-10-29,סԺԤ���˿���֤
Private mbln������Ժ��������˿� As Boolean
Private mblnNurseCall As Boolean

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

'Private mobjSquareCard As Object
Private mblnClickSquareCtrl As Boolean
'��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
Private mcolPayMode As Collection   '��֧����ʽ
Private mlngCardTypeID  As Long
Private mbln���ѿ�     As Boolean
Private mstr���㷽ʽ      As String
Private mstrBrushCardNo As String
 
Private mlngҽ�ƿ����� As Long
Private Type Ty_BillInfor
    lngԤ��ID As Long
    strNo As String
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
Private mblnPassInputCardNo As Boolean  '�Ƿ��������뿨��
Private mblnDefaultPassInputCardNo As Boolean 'ȱʡˢ���Ƿ��������뿨��
Private mrsDepositBalance As ADODB.Recordset    '��ǰ���˵�Ԥ�����
Private mrsDepositInfor As ADODB.Recordset    '��ǰ����Ԥ�����(�������ͼ���ص���ˮ�ŷ������)
Private mbytBackMoneyType As Byte '�˿ʽ:1-��ֹ;0-��ʾ
Private mbytOracleBackType As Byte '�˿���_In;0-�����˿����Ƿ�����˲�����1-����˿���
Private mblnClearWinInfor As Boolean  '�ɿ��,�Ƿ����������Ϣ
Private mblnCheckPass As Boolean 'ˢ��ʱҪ����������,'0000000000'��λ˳���ʾ��������,�ֱ�Ϊ:1.����Һ�,2.���ﻮ��,3.�����շ�,4.�������,5.��Ժ�Ǽ�,6.סԺ����,7.���˽���,8.����Ԥ����,9.���鼼ʦվ,10.Ӱ��ҽ��վ.'
'�������������
Private mobjPlugIn As Object
Private mstrPatiOld As String
Private mstrPatiSex As String
Private mblnOneCard As Boolean  '�Ƿ�ֻ��һ�ž��￨
Private mlngFactModule As Long '��Ʊ��ز���ģ���

Public Function zlShowEdit(ByVal frmMain As Object, ByVal bytCallObject As Byte, _
    ByVal bytInState As Byte, _
    ByVal strPrivs As String, ByVal lngModule As Long, Optional ByVal bytPrepayType As Byte = 0, Optional strInNo As String = "", _
    Optional ByVal blnViewCancel As Boolean = False, Optional blnNOMoved As Boolean = False, _
    Optional ByVal lng����ID As Long = 0, Optional lng��ҳID As Long = 0, Optional dblDefPreMoney As Double = 0, _
    Optional ByVal blnNurseCall As Boolean = False, _
    Optional blnOneCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������,���ڲ���Ԥ������Ϣ�༭��鿴
    '���:frmMain-���õ�������
    '        bytCallObject:���õĶ���(0-Ԥ����������;1-���˷��ò�ѯ����;2-ҽ�ƿ�����)...
    '        bytInState:0-��Ԥ����(ȱʡ,���л�����),1-�������(1),2-����״̬(1);3-����˿�(37770)
    '        bytPrepayType-Ԥ������(0-�����סԺ;1-����;2-סԺ)
    '        strInNo:Ҫ������˿�ĵ��ݺ�(mbytInState=1��3ʱ��Ч),�Ӳ�����Ϣ�Ǽ��е����˿�ʱΪ��
    '         blnViewCancel:�Ƿ�����˿��(mbytInState=1ʱ��Ч)
    '        blnNOMoved:����ϸʱ��¼��ǰѡ��ĵ����Ƿ����������ݱ���,����������ʱ�������ж�
    '        dblDefPreMoney-ȱʡ�Ľɿ���(Ŀǰֻ�в��˷��ò�ѯ�е���ʱ����Ч)
    '        blnNurseCall-��ʿվ����
    '����:
    '����:Ԥ����ֻ��һ�γɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-02-17 16:11:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    mbytCallObject = bytCallObject: mbytInState = bytInState: mstrPrivs = strPrivs: mlngModul = lngModule
    mstrInNO = strInNo: mblnViewCancel = blnViewCancel: mblnNOMoved = blnNOMoved
    mlng����ID = lng����ID: mlng��ҳID = lng��ҳID: mdblDefPreMoney = dblDefPreMoney
    mbytPrepayType = bytPrepayType
    mblnNurseCall = blnNurseCall
    mblnOneCard = blnOneCard
    mlngFactModule = IIf(mbytCallObject = 2, 1107, mlngModul)
    
    mblnOK = False
    If frmMain Is Nothing Then
        frmDeposit.Show
    Else
        frmDeposit.Show 1, frmMain
    End If
    zlShowEdit = mblnOK
End Function
 
Private Sub cboPatiPage_Click()
    If txtPatient.Tag <> "" And mbytInState = 0 And Not mrsInfo Is Nothing And mrsInfo.State = 1 Then
        If cboPatiPage.ItemData(cboPatiPage.ListIndex) <> Nvl(mrsInfo!��ҳID, 0) Then
            Call ShowPatiPageInfo
        End If
    End If
    Call ShowHistoryPrepay("")
End Sub

Private Sub ShowPatiPageInfo()
    Dim lng��ҳID As Long
    lng��ҳID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    '���ݵڼ�����Ժ������Ϣ
    Call GetPatient(IDKind.GetfaultCard, txtPatient.Tag, False, False, lng��ҳID)
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(Val(Nvl(mrsInfo!סԺ��)) = 0, "", "סԺ��:" & mrsInfo!סԺ�� & "   ") & _
                       IIf(Val(Nvl(mrsInfo!�����)) = 0, "", "�����:" & mrsInfo!�����)
    lbl�ѱ�ȼ�.Caption = lbl�ѱ�ȼ�.Tag & mrsInfo!�ѱ�
    txtPatient.Text = mrsInfo!����
    txtPatient.Tag = mrsInfo!����ID
    lblSex.Caption = lblSex.Tag & IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
    lblOld.Caption = lblOld.Tag & IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
    lblҽ�Ƹ��ʽ.Caption = lblҽ�Ƹ��ʽ.Tag & Nvl(mrsInfo!ҽ�Ƹ��ʽ)
    lbl����.Caption = lbl����.Tag & GET��������(mrsInfo!����ID)
    lbl����.Caption = lbl����.Tag & IIf(mrsInfo!���� = 0, "��ͥ", mrsInfo!����)
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(Val(Nvl(mrsInfo!��ǰ����id)) = 0, Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ǰ����id))))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
End Sub
Private Sub cboPatiPage_KeyDown(KeyCode As Integer, Shift As Integer)
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboType_Click()
    If cboType.ListIndex < 0 Then Exit Sub
    
    '88657:���ϴ���2015/9/17,�л�Ԥ������ˢ��Ԥ�����
    If mbytInState = 0 And chkCancel.Value = 0 Or mbytInState = 3 Then
        mlng����ID = 0
        '�����:112784,����,2017/10/13,��ȡ��ȷ��Ʊ�ݸ�ʽ
        mFactProperty = zl_GetInvoicePreperty(mlngFactModule, 2, cboType.ItemData(cboType.ListIndex))
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact
        Call ShowPremayBalance(True, 0)
        Call SetCtrlEnabled
        Call ShowHistoryPrepay("")
    ElseIf mbytInState = 4 Then
        mlng����ID = 0
        '�����:112784,����,2017/10/13,��ȡ��ȷ��Ʊ�ݸ�ʽ
        mFactProperty = zl_GetInvoicePreperty(mlngFactModule, 2, IIf(cboType.ItemData(cboType.ListIndex) = 1, 2, 1))
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact
        Call ShowPremayBalance(True, 0)
        Call SetCtrlEnabled
    '�����:114482,����,2017/10/10,�û��ڽɿ�������Ԥ��ʱ���������Ͻǡ��ˡ���ť��ȷ���Ƿ��ӡ��Ʊ��
    ElseIf mbytInState = 2 Or chkCancel.Value = 1 Then
        mlng����ID = 0
        '�����:112784,����,2017/10/13,��ȡ��ȷ��Ʊ�ݸ�ʽ
        mFactProperty = zl_GetInvoicePreperty(mlngFactModule, 12, cboType.ItemData(cboType.ListIndex))
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact(False, True)
    End If
    
     '�����:45666
    If mbytInState = 0 And cboType.Text = "סԺԤ��" Then '��Ԥ����
        chk����ʾ����Ԥ��.Visible = True
        chk����ʾ����Ԥ��.Value = IIf(zldatabase.GetPara("����ʾ����Ԥ��", glngSys, mlngModul, , Array(chk����ʾ����Ԥ��), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
    Else
        chk����ʾ����Ԥ��.Visible = False
    End If
    lblPatiPage.Visible = cboType.Text = "סԺԤ��": cboPatiPage.Visible = cboType.Text = "סԺԤ��"
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

 
Private Sub chkAllCash_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk����temp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk����ʾ����Ԥ��_Click()
    Call ShowHistoryPrepay("")
End Sub

Private Sub IDKind_Click(objcard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXml As String
    
    If objcard.���� Like "IC��*" And objcard.ϵͳ Then
        If mobjICCard Is Nothing Then
               Set mobjICCard = New clsICCard
               Call mobjICCard.SetParent(Me.hWnd)
               Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text = "" Then Exit Sub
        Call FindPati(objcard, False, txtPatient.Text)
        Exit Sub
    End If
     
    lng�����ID = objcard.�ӿ����
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
    Call FindPati(objcard, False, txtPatient.Text)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objcard As zlIDKind.Card)
    Call txtPatient_GotFocus
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub
Private Sub IDKind_ReadCard(ByVal objcard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text = "" Then Exit Sub
    Call FindPati(objcard, True, txtPatient.Text)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objcard As Card
    Set objcard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objcard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objcard, True, txtPatient.Text)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objcard As Card
    Set objcard = IDKind.GetIDKindCard("���֤", CardTypeName)
    If objcard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objcard, True, txtPatient.Text)
End Sub
Private Sub SetcmdOkEnabled()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����cmdOk��neable����
    '���ƣ����˺�
    '���ڣ�2010-07-09 16:24:53
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    If mrsInfo Is Nothing Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = mrsInfo.State = adStateOpen
    End If
    chk����ʾ����Ԥ��.Enabled = cmdOK.Enabled
End Sub
Private Sub SetCtrlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Enabled����
    '����:���˺�
    '����:2011-07-24 09:30:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean, objCtl As Control
    blnEdit = mbytInState <> 1
    Select Case mbytInState
    Case 4  'תԤ��
        cboType.Enabled = True
        blnEdit = False
        cboUnit.Enabled = blnEdit
        txtUnit.Enabled = blnEdit
        cboStyle.Enabled = blnEdit: cboStyle.ListIndex = -1
        txtCode.Enabled = blnEdit: txt������.Enabled = blnEdit
        txt�ʺ�.Enabled = blnEdit: cboNote.Enabled = blnEdit
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
    Dim strInfo As String, dbl�̶���� As Double
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long
    Dim blnFind As Boolean '�����:55666
    If mbytInState = 2 Or chkCancel.Value = 1 Then Exit Sub
    If mbytInState = 4 Then Exit Sub
    
    If cboStyle.ListIndex = -1 Then Exit Sub
        
    '�����:111657,����,2017/07/25,ʹ���ֽ�֧��Ԥ����ʱ,�λ������������
    mstrBrushCardNo = ""     '�����������ʱ����Ŀ���
    mcurBill.blnת�� = False
    mcurBill.lngԤ��ID = 0
    lngIndex = cboStyle.ListIndex + 1
''    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
   If Not mcolPayMode Is Nothing Then
        '����:56478
        If zlCommFun.GetNeedName(cboStyle.Text) = zlCommFun.GetNeedName(mstr�����ʻ�) Then
            mlngCardTypeID = 0
            mbln���ѿ� = False
            mstr���㷽ʽ = zlCommFun.GetNeedName(mstr�����ʻ�)
        Else
            mlngCardTypeID = Val(mcolPayMode(lngIndex)(3))
            mbln���ѿ� = Val(mcolPayMode(lngIndex)(5)) = 1
            mstr���㷽ʽ = Trim(mcolPayMode(lngIndex)(6))
        End If
        Call ShowPremayBalance(False, 0)
    End If
    Call SetCtrlEnabled
    txtMoney.Enabled = True
    Select Case cboStyle.ItemData(cboStyle.ListIndex)
    Case 3, 1
        txtUnit.Text = "": txt������.Text = "": txt�ʺ�.Text = ""
    Case 2
        If cboStyle.Text Like "*Ʊ*" Or cboStyle.Text Like "*��*" Then
            '��֧Ʊ���ֽ�������,����������
            '����:36611
            If mrsInfo Is Nothing Then Exit Sub
            If mrsInfo.State = adStateClosed Then Exit Sub
            If mrsInfo.EOF Then Exit Sub
            strInfo = GetLastInfo(mrsInfo!����ID)
            If strInfo <> "" Then
                txtUnit.Text = IIf(Split(strInfo, "|")(0) = "", txtUnit.Text, Split(strInfo, "|")(0))
                txt������.Text = IIf(Split(strInfo, "|")(1) = "", txt������.Text, Split(strInfo, "|")(1))
                txt�ʺ�.Text = IIf(Split(strInfo, "|")(2) = "", txt�ʺ�.Text, Split(strInfo, "|")(2))
            End If
        End If
    Case 5 ''ȱʡ���:34705
        varData = Split(mstr���տ�, "|"): dbl�̶���� = 0
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i), ":")
            If varTemp(0) = Split(cboStyle.Text & "-", "-")(0) Then
                dbl�̶���� = Val(varTemp(1)): Exit For
            End If
        Next
        If dbl�̶���� <> 0 Then
            txtMoney.Text = Format(dbl�̶����, "##,##0.00;-##,##0.00; ;"): txtMoney.Enabled = False
             txtMoney.Tag = dbl�̶����:
        End If
    End Select
'    '�����:55666
'     '�µ�����
'    If mrsInfo.State = adStateClosed Then
'        If txtPatient.Visible And cboStyle.Text Like "*��*" Then
'            MsgBox "û��ȷ����ȡԤ����Ĳ���,���ܽ���ˢ��������", vbExclamation, gstrSysName
'            '����Ĭ�ϻ�ԭ��Ϊ�ֽ�֧��
'            For i = 0 To cboStyle.ListCount
'                If cboStyle.List(i) = "�ֽ�" Then
'                    blnFind = True
'                    cboStyle.ListIndex = i
'                End If
'            Next
'            If blnFind And cboStyle.ListCount > 0 Then cboStyle.ListIndex = 0: blnFind = False
'        End If
'        If txtPatient.Visible Then txtPatient.SetFocus: Exit Sub
'    End If
'    If IIf(Trim(txtMoney.Text) = "", "0", Trim(txtMoney.Text)) = "0" And txtMoney.Visible And Not mrsInfo Is Nothing And Not txtMoney Is ActiveControl Then
'        MsgBox "û�������ֵ���,���ܽ���ˢ��������", vbExclamation, gstrSysName
'        txtMoney.SetFocus
'        For i = 0 To cboStyle.ListCount
'                If cboStyle.List(i) = "�ֽ�" Then
'                    blnFind = True
'                    cboStyle.ListIndex = i
'                End If
'            Next
'            If blnFind And cboStyle.ListCount > 0 Then cboStyle.ListIndex = 0: blnFind = False
'        Exit Sub
'    End If
'    'ˢ��
'    CheckBrushCard
End Sub

Private Sub cboStyle_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then
        If cboStyle.ListIndex = -1 Then
            Beep
        Else
            'Call cboStyle_Click
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        If cboStyle.Locked Then Exit Sub
        If KeyAscii >= 32 Then
            lngIdx = cbo.MatchIndex(cboStyle.hWnd, KeyAscii)
            If lngIdx = -1 And cboStyle.ListCount > 0 Then lngIdx = 0
            cboStyle.ListIndex = lngIdx
        End If
    End If
End Sub

Private Sub cboStyle_Validate(Cancel As Boolean)
    If cboStyle.Locked Then Exit Sub
    If Not (cboStyle.ListIndex > -1 And (mbytInState = 0 Or mbytInState = 3)) Then Exit Sub
    If cboStyle.ItemData(cboStyle.ListIndex) = BalanceType.C5���տ� Then
         If mbytInState = 0 Then
             If InStr(mstrPrivs, ";���տ���ȡ;") = 0 Then
                 MsgBox "��û��Ȩ�޽��д��տ���ȡ������", vbInformation, gstrSysName
                 If cbo.Locate(cboStyle, BalanceType.C1�ֽ�, True) Then Cancel = True
             End If
         Else
             If InStr(1, mstrPrivs, ";���տ��˿�;") = 0 Then
                 MsgBox "��û��Ȩ�޽��д��տ���˿������", vbInformation, gstrSysName
                 If cbo.Locate(cboStyle, BalanceType.C1�ֽ�, True) Then Cancel = True
             End If
         End If
     ElseIf mbytInState = 0 Then
         If InStr(1, mstrPrivs, ";Ԥ���տ�;") = 0 Then
             MsgBox "��û��Ȩ�޽���Ԥ���տ������", vbInformation, gstrSysName
             If cbo.Locate(cboStyle, BalanceType.C5���տ�, True) Then Cancel = True
         End If
     Else
         If InStr(1, mstrPrivs, ";Ԥ���˿�;") = 0 Then
             MsgBox "��û��Ȩ�޽���Ԥ���˿������", vbInformation, gstrSysName
             If cbo.Locate(cboStyle, BalanceType.C5���տ�, True) Then Cancel = True
         End If
     End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If SendMessage(cboUnit.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cboUnit.hWnd, KeyAscii, 0.5)
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
        chkCancel.ForeColor = &HFF&
        '�����ؽ��������
        Set mrsInfo = New ADODB.Recordset '���������Ϣ
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
        If cboNO.Visible Then cboNO.SetFocus
    Else
        '����
        chkCancel.ForeColor = 0
        
        picFace.Enabled = True
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
    If mbytInState = 0 Then
        If chkCancel.Value = Checked Then
            If MsgBox("ȷʵҪ�����˿��˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If mrsInfo.State = adStateOpen Then
                If MsgBox("�ò��˵�Ԥ������δ��ȡ,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
    End If
    If mbytInState = 3 Then
        If mrsInfo.State = adStateOpen Then
            If MsgBox("�ò��˵���δ�����˿����,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("δ�����˿����,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    Unload Me
End Sub
Private Sub zlBackDeposit()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���Ԥ������
    '���ƣ����˺�
    '���ڣ�2010-06-18 16:34:59
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, blnExistsSquare As Boolean '�Ƿ���ڽ��㿨
    Dim blnCanDel As Boolean, intInsure As Integer
    Dim bln��ӡ As Boolean
    Dim msgBoxResult As String
    
    mbytOracleBackType = 1
'   �˿�
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
    
    
    '����27363 by lesfeng 2010-01-13
    If MsgBox("ȷʵҪ������ " & cboNO.Text & " ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Select Case mFactProperty.intInvoicePrint
        Case 0 '����ӡԤ����Ʊ
           bln��ӡ = False
        Case 1 '�Զ���ӡ
           bln��ӡ = True
        Case 2 '��ӡ����
            msgBoxResult = MsgBox("�Ƿ���Ҫ��ӡԤ����Ʊ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
            bln��ӡ = (msgBoxResult = vbYes)
        End Select
        
        If mcurBill.lng�����ID = 0 Then
            If Not is���տ�(cboNO.Text) And gbytԤ��������鿨 <> 0 Then
                If mblnסԺ��Ԥ����֤ Or cboType.ItemData(cboType.ListIndex) = 1 Then
                    If Not zldatabase.PatiIdentify(Me, glngSys, Val(txtPatient.Tag), Val(StrToNum(txtMoney.Text)), _
                        , , , , , , , (gbytԤ��������鿨 = 2)) Then Exit Sub
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
        
        '���������ж�
        If cboStyle.ItemData(cboStyle.ListIndex) <> BalanceType.C5���տ� Then
            Dim dblԤ����� As Double
            dblԤ����� = HaveSpare(cboNO.Text)
            If dblԤ����� = 0 And InStr(mstrPrivs, ";Ԥ�������˿�;") = 0 Then
                MsgBox "�ò�����û��Ԥ������û��Ȩ���������ŵ��ݣ�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If HaveBalance(cboNO.Text) <> 0 Then 'And InStr(mstrPrivs, ";Ԥ�������˿�;") = 0 'ɾ�� Ԥ�������˿�Ȩ�� 54779
                MsgBox "�ñ�Ԥ���Ѿ��������ڽ���ʱʹ�ã��㲻���������ŵ��ݣ�", vbInformation, gstrSysName
                Exit Sub
            End If
            '87858
            If CCur(StrToNum(txtMoney.Text)) > dblԤ����� Then
                '46067
                If mbytBackMoneyType = 1 Then
                    '�����˿�,���ܴ�������������:37375
                    Call MsgBox("�ñ�Ԥ���Ľ��Ȳ��˵�ǰ�����࣬�㲻���������ŵ��ݣ�", vbInformation + vbOKOnly, gstrSysName)
                    Exit Sub
                Else
                    If MsgBox("�ñ�Ԥ���Ľ��Ȳ��˵�ǰ�����࣬������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                    mbytOracleBackType = 0
                End If
            End If
        End If
        
        cmdOK.Enabled = False   '��ҽ����ʱ
        
        '��������ӿڽ����Ƿ�Ϸ�
        '108666:���ϴ���2017/5/9���ָ�ȷ�ϰ�ť����״̬
        If zlCheckDepositDelValied(Val(cboNO.Tag), StrToNum(txtMoney.Text)) = False Then cmdOK.Enabled = True: Exit Sub
        
        'ִ�����ϲ���
        If Not CancelBill(CLng(cboNO.Tag), blnCanDel, intInsure, bln��ӡ) Then '�˿�
            MsgBox "����ʧ��,�����Ըò���������������,����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            cmdOK.Enabled = True
            Exit Sub
        End If
        
        If bln��ӡ Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103_1", Me, "NO=" & mstrInNO, "�տ�ʱ��=" & Format(Now, "yyyy-mm-dd HH:MM:SS"), _
                            IIf(mFactProperty.intInvoiceFormat = 0, "", "ReportFormat=" & mFactProperty.intInvoiceFormat), 2)
            Call zlCheckFactIsEnough
        End If
        
        
        Call RePrintBill '���´�ӡԤ����Ʊ
        
        cmdOK.Enabled = True
        
        'ҽ���Ķ�
        For i = 0 To cboStyle.ListCount - 1
            If cboStyle.ItemData(i) = 3 Then
                cboStyle.RemoveItem i: Exit For
            End If
        Next
    End If
    If mbytInState <> 2 Then
        chkCancel.Value = Unchecked '(�������¼�)
    Else
        mblnOK = True
        Unload Me: Exit Sub '�˿�ģʽ�������˳�
    End If
    mblnOK = True
    Call ClearBill
End Sub

Private Sub RePrintBill()
    '���Ϻ����´�ӡԤ����Ʊ
    Dim blnRePrint As Boolean, strNotDelNos As String, strSQL As String
    Dim objFactProperty As Ty_FactProperty
    Dim intInvoiceFormat As Integer, str�տ�ʱ�� As String
    
    On Error GoTo errHandle
    strNotDelNos = GetTurnMZToZYMultiNOs(cboNO.Text, mblnNOMoved)
    If strNotDelNos = "" Then Exit Sub
    
    objFactProperty = zl_GetInvoicePreperty(mlngModul, 2, cboType.ItemData(cboType.ListIndex))
    Select Case objFactProperty.intInvoicePrint
    Case 0 '����ӡԤ����Ʊ
       blnRePrint = False
    Case 1 '�Զ���ӡ
       blnRePrint = True
    Case 2 '��ӡ����
        blnRePrint = MsgBox("��ǰԤ������Ϊ�������תסԺ���ɵģ��Ҹõ������ڵķ�Ʊͬʱ��ӡ�˶���Ԥ�����ݣ�" & _
            "�Ƿ��ʣ�൥�����´�ӡԤ��Ʊ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End Select
    If blnRePrint = False Then Exit Sub
    
    intInvoiceFormat = Val(zldatabase.GetPara(284, glngSys, , "0"))
    
    Call GetFact '���»�ȡ��Ʊ��,��Ϊ��ǰ��Ʊ�ſ���Ҳ����Ʊ��ӡʹ��

    'Ʊ�ݺż��
    If gblnBillԤ�� Then
        If Trim(txtFact.Text) = "" Then
            MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
        Else
            mlng����ID = CheckUsedBill(2, IIf(mlng����ID > 0, mlng����ID, objFactProperty.lngShareUseID), _
                txtFact.Text, cboType.ItemData(cboType.ListIndex))
            If mlng����ID <= 0 Then
                Select Case mlng����ID
                Case 0 '����ʧ��
                Case -1
                    MsgBox "��û�����ú͹��õ�Ԥ��Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Case -3
                    MsgBox "Ʊ�ݺ��벻�ڵ�ǰ��Ч���÷�Χ�ڣ�", vbInformation, gstrSysName
                End Select
                txtFact.Text = ""
            End If
        End If
    Else
        If Len(txtFact.Text) <> gbytԤ�� And txtFact.Text <> "" Then
            MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytԤ�� & " λ��", vbInformation, gstrSysName
            txtFact.Text = ""
        End If
    End If
    
    If Trim(txtFact.Text) <> "" Then '��Ʊ����Ч����ӡ
        'ִ�����ݴ���
        'Zl_����Ԥ����¼_Reprint
        strSQL = "Zl_����Ԥ����¼_Reprint("
        '  ���ݺ�_In Varchar2,
        strSQL = strSQL & "'" & strNotDelNos & "',"
        '  Ʊ�ݺ�_In Ʊ��ʹ����ϸ.����%Type,
        strSQL = strSQL & "'" & Trim(txtFact.Text) & "',"
        '  ����id_In Ʊ��ʹ����ϸ.����id%Type,
        strSQL = strSQL & "" & IIf(mlng����ID = 0, "NULL", mlng����ID) & ","
        '  ʹ����_In Ʊ��ʹ����ϸ.ʹ����%Type
        strSQL = strSQL & "'" & UserInfo.���� & "')"
        zldatabase.ExecuteProcedure strSQL, Me.Caption
        
        If Not gblnBillԤ�� Then
            '��ɢ�����浱ǰ����
            zldatabase.SetPara "��ǰԤ��Ʊ�ݺ�", Trim(txtFact.Text), glngSys, mlngFactModule
        End If
        
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & strNotDelNos, _
            "�տ�ʱ��=" & Format(mcurBill.dt�տ�ʱ��, "yyyy-mm-dd HH:MM:SS"), _
            IIf(intInvoiceFormat = 0, "", "ReportFormat=" & intInvoiceFormat), 2)
        Call zlCheckFactIsEnough
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetTurnMZToZYMultiNOs(ByVal strNo As String, Optional ByVal blnNOMoved As Boolean) As String
    '���ܣ���ȡ����תסԺ������Ԥ�����ݣ�������һ�δ�ӡ�Ķ��ŵ��ݺ�
    '���:strNo-��Ҫ�ش�NO
    '     blnNOMoved-�Ƿ�ת����ʷ��ռ�
    '����:
    '����:һ�δ�ӡ�Ķ��ŵ��ݺţ���ʽ��A001,A002,A003,...
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strNos As String
    
    On Error GoTo errHandle
    'Ӧ�������һ�δ�ӡ���������
    strSQL = _
        "Select a.NO" & vbNewLine & _
        "From Ʊ�ݴ�ӡ���� A" & vbNewLine & _
        "Where a.�������� = 2" & vbNewLine & _
        "      And a.ID In (Select ID" & vbNewLine & _
        "                From (Select b.Id" & vbNewLine & _
        "                      From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B" & vbNewLine & _
        "                      Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 2 And b.No = [1]" & vbNewLine & _
        "                      Order By a.ʹ��ʱ�� Desc)" & vbNewLine & _
        "                Where Rownum < 2)" & vbNewLine & _
        "      And Not Exists(Select 1 From ����Ԥ����¼ Where ��¼���� = 1 And ��¼״̬ = 2 And No = a.No)" & vbNewLine & _
        "Order By No"
    If blnNOMoved Then
        strSQL = Replace(strSQL, "Ʊ�ݴ�ӡ����", "HƱ�ݴ�ӡ����")
        strSQL = Replace(strSQL, "Ʊ��ʹ����ϸ", "HƱ��ʹ����ϸ")
    End If
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "", strNo)
    If rsTemp.EOF Then Exit Function
    
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & Nvl(rsTemp!NO)
            .MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    GetTurnMZToZYMultiNOs = strNos
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckDataValied() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���������Ƿ�Ϸ�
    '���أ��Ϸ�����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-18 16:38:39
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
   '�µ�����
  If mrsInfo.State = adStateClosed Then
      If mbytInState = 3 Then
            MsgBox "û��ȷ����Ԥ����Ĳ���,�����˿", vbExclamation, gstrSysName
      Else
            MsgBox "û��ȷ����ȡԤ����Ĳ���,���ܴ��̣�", vbExclamation, gstrSysName
      End If
      txtPatient.SetFocus: Exit Function
  End If
  
    If mbytInState = 3 And chkAllCash.Value = 1 Then
        If Val(cboStyle.ItemData(cboStyle.ListIndex)) <> 1 And Val(cboStyle.ItemData(cboStyle.ListIndex)) <> 2 Then
            MsgBox "�����˻�ǿ�����ֵ�����£�ֻ��ѡ���ֽ����֧Ʊ��Ľ��㷽ʽ��", vbInformation, gstrSysName
            If cboStyle.Enabled And cboStyle.Visible Then cboStyle.SetFocus
            Exit Function
        End If
    End If
          
  If LenB(StrConv(txtUnit.Text, vbFromUnicode)) > 50 Then
      MsgBox "�ɿλ����ֻ���� 50 ���ַ��� 25 ������,���޸ģ�", vbInformation, App.Title
      txtUnit.SetFocus: Exit Function
  End If
  If LenB(StrConv(txt������.Text, vbFromUnicode)) > 50 Then
      MsgBox "����������ֻ���� 50 ���ַ��� 25 ������,���޸ģ�", vbInformation, App.Title
      txt������.SetFocus: Exit Function
  End If
  If LenB(StrConv(cboNote.Text, vbFromUnicode)) > 50 Then
      MsgBox "�ɿ�ժҪֻ���� 50 ���ַ��� 25 ������,���޸ģ�", vbInformation, App.Title
      cboNote.SetFocus: Exit Function
  End If
  If mbytInState = 0 Then
    If cboType.ListIndex < 0 Then Exit Function
    '����:44963
    If mrsInfo Is Nothing Then Exit Function
    If cboType.ItemData(cboType.ListIndex) = 2 Then
        If Val(Nvl(mrsInfo!��Ժ)) = 0 And gblnAllowOut = False Then
            strSQL = "Select 1 From ������ҳ Where ����ID=[1] And Nvl(��ҳID,0)=0 And Nvl(��������,0)=0" 'Ԥ��Ժ
            Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsInfo!����ID)))
            If rsTemp.EOF Then
                MsgBox "���˻�δסԺ,���ܽ�סԺԤ��,����!", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Else
        If Val(Nvl(mrsInfo!��Ժ)) = 1 And gblnBanIn = True Then
            MsgBox "���˻�δ��Ժ,���ܽ�����Ԥ��,����!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
  End If
  '�����
  '����27363 by lesfeng 2010-01-13
  If txtMoney.Text = "" And mblnClickSquareCtrl = False Then
      MsgBox IIf(mbytInState = 3, "�˿���", "�տ���") & "����Ϊ��,�����룡", vbExclamation, gstrSysName
      txtMoney.SetFocus: Exit Function
  ElseIf CCur(StrToNum(txtMoney.Text)) = 0 And mblnClickSquareCtrl = False Then
      MsgBox IIf(mbytInState = 3, "�˿���", "�տ���") & "����Ϊ��,�����룡", vbExclamation, gstrSysName
      txtMoney.SetFocus: Exit Function
  End If

  If InStr(mstrPrivs, ";�����ɿ�;") = 0 And StrToNum(txtMoney.Text) < 0 Then
      MsgBox IIf(mbytInState = 3, "�˿���", "�տ���") & "����Ϊ����,������", vbExclamation, gstrSysName
      txtMoney.SetFocus: Exit Function
  End If
  mbytOracleBackType = 1
  
  If mbytInState = 3 Then
        If mbln������Ժ��������˿� = False And cboType.ItemData(cboType.ListIndex) = 2 Then
            If Val(Nvl(mrsInfo!��Ժ)) = 1 Then
                MsgBox "������Ժ,���ܽ�������˿�,����!", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If mdblʣ���� - IIf(cboStyle.ItemData(cboStyle.ListIndex) < 0, 0, mdblԤ�����_����) - CCur(StrToNum(txtMoney.Text)) < 0 Then
            '46067
            If mbytBackMoneyType = 1 Then
                Call MsgBox("�˿���Ȳ��˵�ǰ������,�����˿�!", vbInformation + vbOKOnly, gstrSysName)
                txtMoney.SetFocus: Exit Function
            Else
                If MsgBox("�˿���Ȳ��˵�ǰ������,������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    txtMoney.SetFocus: Exit Function
                End If
                mbytOracleBackType = 0
            End If
        End If
        '�����:112995,����,2017/10/13,�˿��˷�ʱ��ʾ�����˷ѽ��
        If mblnOneCard Then
            If mdblʣ���� - IIf(cboStyle.ItemData(cboStyle.ListIndex) < 0, 0, mdblԤ�����_����) - CCur(StrToNum(txtMoney.Text)) > 0 Then
                MsgBox "�˿���Ȳ��˵�ǰ�������,������˿����!", vbInformation + vbOKOnly, gstrSysName
                txtMoney.SetFocus: Exit Function
            End If
        End If
        If cboStyle.ItemData(cboStyle.ListIndex) < 0 Then
            If mdblԤ�����_���� - CCur(StrToNum(txtMoney.Text)) < 0 Then
              Call MsgBox("" & cboStyle.Text & "���ֻ����" & Format(mdblԤ�����_����, "###0.00;-###0.00;;") & "!", vbInformation + vbOKOnly, gstrSysName)
              txtMoney.SetFocus: Exit Function
            End If
        End If
        If gbytԤ��������鿨 <> 0 Then
            If mrsInfo Is Nothing Then
                lng����ID = Val(txtPatient.Tag)
            ElseIf mrsInfo.State <> 1 Then
                lng����ID = Val(txtPatient.Tag)
            Else
                lng����ID = mrsInfo!����ID
            End If
            If mblnסԺ��Ԥ����֤ Or cboType.ItemData(cboType.ListIndex) = 1 Then
                If Not zldatabase.PatiIdentify(Me, glngSys, lng����ID, Val(StrToNum(txtMoney.Text)), _
                    , , , , , , , (gbytԤ��������鿨 = 2)) Then Exit Function
            End If
        End If
        
  Else
        If CCur(StrToNum(txtMoney.Text)) < 0 And Abs(CCur(StrToNum(txtMoney.Text))) > mdblʣ���� Then
            '46067
            If mbytBackMoneyType = 1 Then
                    '�����˿�,���ܴ�������������:37375
                    Call MsgBox("�˿���Ȳ��˵�ǰ������,�����˿�!", vbInformation + vbOKOnly, gstrSysName)
                    txtMoney.SetFocus: Exit Function
            Else
                If MsgBox("�˿���Ȳ��˵�ǰ������,������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    txtMoney.SetFocus: Exit Function
                End If
                mbytOracleBackType = 0
            End If
        End If
  End If
  
  If cboStyle.ListIndex = -1 Then
      MsgBox "��ȷ�����㷽ʽ��", vbExclamation, gstrSysName
      cboStyle.SetFocus: Exit Function
  End If
  
  If cboStyle.ItemData(cboStyle.ListIndex) = BalanceType.C5���տ� Then
      If mbytInState = 3 Then
           If InStr(1, mstrPrivs, ";���տ��˿�;") = 0 Then
                MsgBox "��û��Ȩ�޽��д��տ��˿������", vbInformation, gstrSysName
                Exit Function
           End If
      Else
            If InStr(mstrPrivs, ";���տ���ȡ;") = 0 Then
                MsgBox "��û��Ȩ�޽��д��տ���ȡ������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
  ElseIf mbytInState = 3 Then
        '�˿����
        If InStr(1, mstrPrivs, ";Ԥ���˿�;") = 0 Then
            MsgBox "��û��Ȩ�޽���Ԥ���˿������", vbInformation, gstrSysName: Exit Function
        End If
  ElseIf InStr(mstrPrivs, ";Ԥ���տ�;") = 0 Then
      MsgBox "��û��Ȩ�޽���Ԥ���տ������", vbInformation, gstrSysName
      Exit Function
  End If
  'ҽ���Ķ�
  If cboStyle.ItemData(cboStyle.ListIndex) = 3 Then
      If mbytInState = 3 Then
            MsgBox "ҽ�����˸����ʻ�ת�ʽ��ܽ����˿", vbInformation, gstrSysName
            txtMoney.SetFocus: Exit Function
      Else
            If CCur(StrToNum(txtMoney.Text)) < 0 Then
                MsgBox "ҽ�����˸����ʻ�ת�ʽ���Ϊ����", vbInformation, gstrSysName
                txtMoney.SetFocus: Exit Function
            End If
            If CCur(StrToNum(txtMoney.Text)) > mcur�ʻ���� Then
                MsgBox "ҽ�����˸����ʻ�ת�ʽ��ܳ������:" & Format(mcur�ʻ����, "0.00"), vbInformation, gstrSysName
                txtMoney.SetFocus: Exit Function
            End If
        End If
  End If
  
  If mblnClickSquareCtrl Then
      If CCur(StrToNum(txtMoney.Text)) < 0 Then
          MsgBox "���ʿ�תԤ������Ϊ����", vbInformation, gstrSysName
          txtMoney.SetFocus: Exit Function
      End If
  End If
  
  '�����:50656
  If mFactProperty.intInvoicePrint = 0 Then CheckDataValied = True: Exit Function
  'Ʊ�ݺ�����
  If gblnBillԤ�� Then
      If Trim(txtFact.Text) = "" Then
          MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
          txtFact.SetFocus: Exit Function
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
  Else
      If Len(txtFact.Text) <> gbytԤ�� And txtFact.Text <> "" Then
          MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytԤ�� & " λ��", vbInformation, gstrSysName
          txtFact.SetFocus: Exit Function
      End If
  End If
    CheckDataValied = True
End Function

Private Function Select�����˿�() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˿�ѡ��
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-21 18:01:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset, strSQL As String
    Dim intԤ������ As Integer, bln�����ӿ� As Boolean
    Dim strWhere As String, lng����ID As Long
    Dim vRect  As RECT, blnCancel As Boolean
    
    On Error GoTo errHandle
    If mbytInState = 4 Then Select�����˿� = True: Exit Function
    
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    
    If Me.cboStyle.ItemData(cboStyle.ListIndex) >= 0 Then Select�����˿� = True: Exit Function
    
    mcurBill.lng�����ID = -1
    
    lng����ID = Val(Nvl(mrsInfo!����ID))
    mdbl������� = 0: mdblԤ����� = 0: mdblʣ���� = 0
    intԤ������ = cboType.ItemData(cboType.ListIndex)
    
    '����˿�ʱ,��Ҫ��鵱ǰ֧�����Ƿ����
    '˵���� And nvl(A.����˵��,' ') Like '%%' ��һ���Ŀ����ʹ�ü�¼ֻ��һ��ʱ������ѡ����
    strWhere = IIf(mbln���ѿ�, " And A.���㿨���=[3] ", " And nvl(A.�����ID,0)=[3]")
    strSQL = _
        "Select a.�����id, a.���㿨���, Min(a.�տ�ʱ��) As �տ�ʱ��, a.����, a.������ˮ��, a.����˵��," & vbNewLine & _
        "       Max(Decode(Sign(a.���), -1, 0, Decode(a.��¼����, 11, 0, ID))) As Ԥ��id," & vbNewLine & _
        "       Sum(Nvl(���, 0)) - Sum(Nvl(��Ԥ��, 0)) As Ԥ�����" & vbNewLine & _
        "From ����Ԥ����¼ A" & vbNewLine & _
        "Where a.����id = [1] And a.Ԥ����� = [2] And a.��¼���� In (1, 11) And nvl(A.����˵��,' ') Like '%%'" & strWhere & vbNewLine & _
        "Group By a.�����id, a.���㿨���, a.����, a.������ˮ��, a.����˵��"
    '108376,���� 2017,06/13 �б��ϼ�һ��"��������"(Ԥ������տ�ʱ��)
    strSQL = _
        "Select Distinct a.Ԥ��id, a.�����id, a.���㿨��� As ���ѽӿ�id, Nvl(b.����, c.���) As ����," & vbNewLine & _
        "       Nvl(b.����, c.����) As ����, a.����, a.������ˮ��, a.����˵��, Nvl(b.�Ƿ�ת�ʼ�����, 0) As ת��," & vbNewLine & _
        "       a.Ԥ�����, a.�տ�ʱ�� As ��������, Nvl(b.�Ƿ��˿��鿨, 0) As �Ƿ��˿��鿨, n.���ѿ�id" & vbNewLine & _
        "From (" & strSQL & ") A, ҽ�ƿ���� B, ���ѿ����Ŀ¼ C, ���˿������¼ N" & vbNewLine & _
        "Where a.�����id = b.Id(+) And a.���㿨��� = c.���(+) And a.Ԥ��id = n.����id(+)" & vbNewLine & _
        "      And Nvl(a.Ԥ�����, 0) > 0" & vbNewLine & _
        "      And Not Exists (Select 1 From ���ѿ���Ϣ Where �ӿڱ�� = a.���㿨��� And ���� = a.���� And Nvl(��ǰ״̬, 1) <> 1" & vbNewLine & _
        "           And ��� = (Select Max(���) From ���ѿ���Ϣ Where �ӿڱ�� = a.���㿨��� And ���� = a.����))" & vbNewLine & _
        "Order By ����"
    
    strSQL = _
        "Select Rownum As ID, Ԥ��id, �����id, ���ѽӿ�id, ����, ����, ����, ������ˮ��, ����˵��, " & vbNewLine & _
        "       Ԥ�����, ��������, ת�� As ת��_ID, �Ƿ��˿��鿨 As �Ƿ��˿��鿨_ID, ���ѿ�id" & vbNewLine & _
        "From (" & strSQL & ")"
    Set rsMoney = zldatabase.ShowSQLSelect(Me, strSQL, 0, cboStyle.Text & "�˿�", False, "", "��ѡ����Ҫ�˿�Ľ���", _
        False, False, False, vRect.Left, vRect.Top, cboStyle.Height, blnCancel, True, True, lng����ID, intԤ������, _
        mlngCardTypeID)
    If blnCancel Then Exit Function
    If rsMoney Is Nothing Then
        MsgBox cboStyle.Text & "�����ڿ������,�����˿�!", vbOKOnly + vbInformation, gstrSysName
        txtMoney.Text = "0.00"
        Exit Function
    End If
    
    With rsMoney
        mcurBill.lngԤ��ID = Val(Nvl(!Ԥ��ID))
        mcurBill.bln���ѿ� = mbln���ѿ�
        mcurBill.lng�����ID = mlngCardTypeID
        mcurBill.str������ˮ�� = Nvl(!������ˮ��)
        mcurBill.str����˵�� = Nvl(!����˵��)
        mcurBill.str���� = Nvl(!����)
        mcurBill.blnת�� = Val(Nvl(!ת��_ID)) = 1
        mcurBill.bln�˿��鿨 = Val(Nvl(!�Ƿ��˿��鿨_ID)) = 1
        mcurBill.dbl��� = Val(Val(Nvl(rsMoney!Ԥ�����)))
        mcurBill.lng���ѿ�ID = Val(Val(Nvl(rsMoney!���ѿ�ID)))
        txtMoney.Text = Format(Val(Nvl(rsMoney!Ԥ�����)), "#,###0.00;-#,###0.00;;")
        mdblԤ�����_���� = Val(Val(Nvl(rsMoney!Ԥ�����)))
    End With
    Select�����˿� = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function
Private Function CheckBrushCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ˢ��
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset
    Dim lng����ID As String
    Dim dblMoney As Double
    Dim strExpand As String '�����:55666
    Dim dbl�˻���� As Double '�����:55666
    Dim bln���ѿ�ˢ�� As Boolean '�����:55666
    Dim bln������ˢ�� As Boolean '�����:55666
    
    On Error GoTo errHandle
    dblMoney = IIf(mbytInState = 3, -1, 1) * StrToNum(txtMoney.Text)
    If cboStyle.ItemData(cboStyle.ListIndex) >= 0 Then CheckBrushCard = True: Exit Function
    If mbytInState = 3 Then
        If mcurBill.lng�����ID < 0 Then
            If Select�����˿� = False Then
                Exit Function
            End If
        End If
        dblMoney = StrToNum(txtMoney.Text)
         If zlCheckDepositDelValied(mcurBill.lngԤ��ID, dblMoney) = False Then Exit Function
         CheckBrushCard = True: Exit Function
    End If
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
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = adStateOpen Then lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, mlngCardTypeID, mbln���ѿ�, _
        Nvl(mrsInfo!����), Nvl(mrsInfo!�Ա�), Nvl(mrsInfo!����), dblMoney, mstrBrushCardNo, mstrbrPassWord, _
        False, True, False, False, Nothing, False, False, "<IN><CZLX>0</CZLX></IN>", _
        cboType.ItemData(cboType.ListIndex), lng����ID) = False Then Exit Function
    '����ǰ,һЩ���ݼ��
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, mlngCardTypeID, _
        mbln���ѿ�, mstrBrushCardNo, dblMoney, "", "") = False Then Exit Function
    '�����:55666,55851
    gobjSquare.objSquareCard.zlGetAccountMoney Me, mlngModul, mlngCardTypeID, mstrBrushCardNo, strExpand, dbl�˻����, mbln���ѿ�
    If dbl�˻���� <> 0 Then sta.Panels(2).Text = "�˻����:" & dbl�˻����
    '�ж�Ԥ�����Ƿ񳬳�ˢ�������
    lblRepairMoney.Visible = CDbl(txtMoney.Text) > dblMoney
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
Private Function CheckChangDepositType() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���תԤ��������
    '����:���ݺϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 09:54:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    If mrsInfo.State = adStateClosed Then
        MsgBox "û��ȷ��תԤ����Ĳ���,����תԤ����", vbExclamation, gstrSysName
       If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus: Exit Function
    End If
    If LenB(StrConv(cboNote.Text, vbFromUnicode)) > 50 Then
        MsgBox "�ɿ�ժҪֻ���� 50 ���ַ��� 25 ������,���޸ģ�", vbInformation, App.Title
        If cboNote.Enabled And cboNote.Visible Then cboNote.SetFocus: Exit Function
    End If
    If txtMoney.Text = "" Then
        MsgBox "תԤ�����Ϊ��,�����룡", vbExclamation, gstrSysName
        If txtMoney.Enabled And txtMoney.Visible Then txtMoney.SetFocus: Exit Function
    ElseIf CCur(StrToNum(txtMoney.Text)) = 0 Then
        MsgBox "תԤ�����Ϊ��,�����룡", vbExclamation, gstrSysName
        If txtMoney.Enabled And txtMoney.Visible Then txtMoney.SetFocus: Exit Function
    End If
    
    If StrToNum(txtMoney.Text) < 0 Then
        MsgBox "תԤ�����Ϊ����,������", vbExclamation, gstrSysName
        If txtMoney.Enabled And txtMoney.Visible Then txtMoney.SetFocus: Exit Function
    End If
  
    If mdblʣ���� - CCur(StrToNum(txtMoney.Text)) < 0 Then
        Call MsgBox("תԤ�����Ȳ��˵�ǰ������,�����˿�!", vbInformation + vbOKOnly, gstrSysName)
        If txtMoney.Enabled And txtMoney.Visible Then txtMoney.SetFocus: Exit Function
    End If
    
    '112999
    If cboType.ListIndex < 0 Then Exit Function
    If cboType.ItemData(cboType.ListIndex) = 1 Then
        If Val(Nvl(mrsInfo!��Ժ)) = 0 And gblnAllowOut = False Then
            strSQL = "Select 1 From ������ҳ Where ����ID=[1] And Nvl(��ҳID,0)=0 And Nvl(��������,0)=0" 'Ԥ��Ժ
            Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsInfo!����ID)))
            If rsTemp.EOF Then
                MsgBox "���˻�δסԺ����������Ԥ��תסԺ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Else
        If Val(Nvl(mrsInfo!��Ժ)) = 1 And gblnBanIn = True Then
            MsgBox "���˻�δ��Ժ������סԺԤ��ת���", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    'Ʊ�ݺ�����
    If gblnBillԤ�� Then
        If Trim(txtFact.Text) = "" Then
            MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
            txtFact.SetFocus: Exit Function
        End If
        If mbytInState = 4 Then
            mlng����ID = CheckUsedBill(2, IIf(mlng����ID > 0, mlng����ID, mFactProperty.lngShareUseID), txtFact.Text, IIf(cboType.ItemData(cboType.ListIndex) = 1, 2, 1))
        Else
            mlng����ID = CheckUsedBill(2, IIf(mlng����ID > 0, mlng����ID, mFactProperty.lngShareUseID), txtFact.Text, cboType.ItemData(cboType.ListIndex))
        End If
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
    Else
        If Len(txtFact.Text) <> gbytԤ�� And txtFact.Text <> "" Then
            MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytԤ�� & " λ��", vbInformation, gstrSysName
            txtFact.SetFocus: Exit Function
        End If
    End If
    CheckChangDepositType = True
End Function
Private Function SaveChageDepositType() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ������ת��
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 10:02:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String

    On Error GoTo errHandle
    mstrPrintDate = Format(zldatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    'Zl_����Ԥ����¼_תԤ��
    strSQL = "Zl_����Ԥ����¼_תԤ��("
    '  Ʊ�ݺ�_In     Ʊ��ʹ����ϸ.����%Type,
    strSQL = strSQL & "'" & txtFact.Text & "',"
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & Val(Nvl(mrsInfo!����ID)) & ","
    '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
    strSQL = strSQL & "" & IIf(mrsInfo!��ǰ����id <> 0 And mrsInfo!��Ժ <> 0, mrsInfo!��ҳID, "Null") & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & IIf(cboUnit.ItemData(cboUnit.ListIndex) = 0, "NULL", cboUnit.ItemData(cboUnit.ListIndex)) & ","
    '  ���_In       ����Ԥ����¼.���%Type,
    strSQL = strSQL & "" & StrToNum(txtMoney.Text) & ","
    '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
   strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
   strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
   strSQL = strSQL & "to_Date('" & mstrPrintDate & "','yyyy-mm-dd hh24:mi:ss'),"
    '  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "" & IIf(mlng����ID = 0, "NULL", mlng����ID) & ","
    '  Ԥ�����_In   ����Ԥ����¼.Ԥ�����%Type,
    strSQL = strSQL & "" & cboType.ItemData(cboType.ListIndex) & ","
    '  ժҪ_In       ����Ԥ����¼.ժҪ%Type
   strSQL = strSQL & "'" & cboNote.Text & "')"
    zldatabase.ExecuteProcedure strSQL, Me.Caption
    SaveChageDepositType = True
    
    If Not gblnBillԤ�� And Trim(txtFact.Text) <> "" Then
        '��ɢ�����浱ǰ����
        zldatabase.SetPara "��ǰԤ��Ʊ�ݺ�", Trim(txtFact.Text), glngSys, mlngFactModule
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    Dim i As Integer, blnCardEnable As Boolean, lng�ӿڱ�� As Long, strBlanceInfor As String
    Dim varData As Variant, blnHave���㷽ʽ As Boolean
    Dim blnCanDel As Boolean, intInsure As Integer
    Dim msgBoxResult As VbMsgBoxResult '�����:50656
    Dim bln��ӡ As Boolean  '�����:57624
    Dim lngԤ��ID As Long
    
    If chkCancel.Value = Checked Then
         Call zlBackDeposit: Exit Sub
    End If
    
    '�����:57624
    '�����:50565
    Select Case mFactProperty.intInvoicePrint
    Case 0 '����ӡԤ����Ʊ
       bln��ӡ = False
    Case 1 '�Զ���ӡ
       bln��ӡ = True
    Case 2 '��ӡ����
        msgBoxResult = MsgBox("�Ƿ���Ҫ��ӡԤ��Ʊ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        bln��ӡ = (msgBoxResult = vbYes)
    End Select
    
    If mbytInState = 4 Then
        '--����תסԺ��סԺת����
        If CheckChangDepositType = False Then Exit Sub
        If SaveChageDepositType = False Then Exit Sub
        If bln��ӡ Then    '120271
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "�տ�ʱ��=" & mstrPrintDate, "NO='��' ", "����ID=" & mrsInfo!����ID, "ReportFormat=2", 2)
        End If
        mblnOK = True: Unload Me: Exit Sub
    End If
    
    If Not Checkδ��Ʋ���Ԥ�� Then Exit Sub
    If CheckDataValied = False Then Exit Sub
    If CheckBrushCard = False Then Exit Sub
    '����
    cmdOK.Enabled = False
    
    If Check�˿� = False Then cmdOK.Enabled = True: Exit Sub
    '�м䲻���е����࣬���ⳤʱ�������ɲ���
    If Not SaveBill(bln��ӡ, lngԤ��ID) Then
        MsgBox "Ԥ����ݱ���ʧ��,�����Ըò����������������,����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
        cmdOK.Enabled = True: Exit Sub
    Else
        '�����:57624
        '�����:50656
        If bln��ӡ Then 'Ʊ�ݺ�Ϊ�վͱ�ʾ����ӡ��Ʊ
            '78751:���ϴ�,2014/10/20,����Ԥ��Ʊ�ݴ�ӡ��ʽ
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & cboNO.List(0), "����ID=" & mrsInfo!����ID, "�տ�ʱ��=" & Format(Now, "yyyy-mm-dd HH:MM:SS"), _
                            IIf(mFactProperty.intInvoiceFormat = 0, "", "ReportFormat=" & mFactProperty.intInvoiceFormat), 2)
            Call zlCheckFactIsEnough
        End If
        
        '81693:���ϴ�,2015/4/21,������
        If Not mobjPlugIn Is Nothing Then
            On Error Resume Next
            Call mobjPlugIn.PatiPrePayAfter(mrsInfo!����ID, IIf(mbytPrepayType = 2, 1, 0), lngԤ��ID)
            Err.Clear
        End If
    End If
    '�����:55666
    '���ڲ����������
    If UBound(Split(lblRepairMoney.Caption, ":")) = 1 And Split(lblRepairMoney.Caption, ":")(1) <> "" Then
        txtPatient.Tag = ""
        txtMoney.Text = Split(lblRepairMoney.Caption, ":")(1)
        IDKind.IDKind = IDKind.GetKindIndex("����")
        txtPatient.Text = "-" & mrsInfo!����ID
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
    
        lblRepairMoney.Visible = False: lblRepairMoney.Caption = "������:"
        cmdOK.Enabled = True
        Exit Sub
    End If
    
    '����:48249
    If mbytCallObject = 1 Or mbytCallObject = 2 Then
        '���ò�ѯʱ,ֱ���˳�
        mblnOK = True: Unload Me: Exit Sub
    End If
    
    If mblnClearWinInfor Then
        Call ClearBill
        Call InitFace(True)
        Call cboStyle_Click
    Else
        '�����:44732
        SetMoneyInfo False, , , True
        Set mrsInfo = New ADODB.Recordset
        
        If mFactProperty.intInvoicePrint <> 0 Then Call GetFact  '���»�ȡ��Ʊ��
    End If
    Call SetcmdOkEnabled
    If txtPatient.Enabled Then txtPatient.SetFocus
    mblnOK = True
End Sub

Private Sub ClearBill()
'����:�����ؽ��������
    If (mbytInState = 0 Or mbytInState = 3) And gblnLED Then
        zl9LedVoice.DisplayPatient ""
    End If
    
    Set mrsInfo = New ADODB.Recordset '���������Ϣ
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
    If cboStyle.ListCount <> 0 And cboStyle.Tag <> "" Then cboStyle.ListIndex = Val(cboStyle.Tag) '�ָ�ȱʡ���㷽ʽ
    txtCode.Text = "": txtCode.Locked = False
    
    txtMan.Text = UserInfo.����
    txtDate.Text = Format(zldatabase.Currentdate(), "yyyy-MM-dd")
    cboNote.Text = ""
    
    'ҽ���Ķ�
    Call Clear�����ʻ�
    
    '�µ�һ��Ԥ�����
    cboNO.Text = "": cboNO.Locked = True
    
    txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ�ݺ�") And gblnBillԤ�� '89302
    If mFactProperty.intInvoicePrint <> 0 Then Call GetFact
    txtPatient.SetFocus
End Sub

Private Sub cmdSetup_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub chkAllCash_Click()
    Dim dblMoney As Double, lngRow As Long
    Dim strDBUser As String
    Dim strPrivs As String
    
    If chkAllCash.Value = 1 Then
        If mstr�˿����Ա <> "" Then
            mdblԤ�����_���� = 0
            GoTo ResetMoney
        End If
        If InStr(";" & mstrCardPrivs & ";", ";�����˿�ǿ������;") = 0 Then
            mstr�˿����Ա = zldatabase.UserIdentifyByUser(Me, "ǿ��������֤", glngSys, 1151, "�����˿�ǿ������")
            If mstr�˿����Ա = "" Then
                MsgBox "¼��Ĳ���Ա��֤ʧ�ܻ���¼��Ĳ���Ա���߱�ǿ������Ȩ�ޣ�����ǿ�����֣�", vbInformation, gstrSysName
                chkAllCash.Value = 0
                GoTo ResetMoney
            End If
        
            mdblԤ�����_���� = 0
        Else
            If MsgBox("���ڲ�֧�����ֵ�������,�Ƿ�����ǿ�ƽ������֣�", _
                                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then chkAllCash.Value = 0: GoTo ResetMoney
            mstr�˿����Ա = UserInfo.����
            mdblԤ�����_���� = 0
        End If
    Else
        mdblԤ�����_���� = mdblԤ�����_��������
    End If
ResetMoney:
    If mdblʣ���� - mdblԤ�����_���� > 0 Then
        txtMoney.Text = Format(mdblʣ���� - mdblԤ�����_����, "#,##0.00;-#,##0.00;;")
    Else
        txtMoney.Text = ""
    End If
End Sub

Private Sub Form_Activate()
    If mblnUnLoad Then Unload Me: Exit Sub
    If mbytInState = 0 Or mbytInState = 3 Or mbytInState = 4 Then
        ' mbytInState=3:��ʾ����˿�,4-��ʾתԤ��
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
    ElseIf mbytInState = 1 Then
        cmdCancel.SetFocus
    ElseIf mbytInState = 2 Then
        If mstrInNO = "" Then
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        Else
           If cmdOK.Enabled Then cmdOK.SetFocus
        End If
    End If
    '�����:45666
    If mbytInState = 0 And cboType.Text = "סԺԤ��" Then '��Ԥ����
        chk����ʾ����Ԥ��.Visible = True
        chk����ʾ����Ԥ��.Value = IIf(zldatabase.GetPara("����ʾ����Ԥ��", glngSys, mlngModul, , Array(chk����ʾ����Ԥ��), InStr(mstrPrivs, ";��������;") > 0) = "1", 1, 0)
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
            txtFact.Text = zlCommFun.IncStr(UCase(zldatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, mlngFactModule, "")))
        Else
            mstrRedFact = zlCommFun.IncStr(UCase(zldatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, mlngFactModule, "")))
        End If
    End If
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ģ�����
    '����:���˺�
    '����:2012-02-27 11:23:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrȱʡ���㷽ʽ = zldatabase.GetPara("ȱʡԤ�����㷽ʽ", glngSys, mlngModul)
    mbytBackMoneyType = Val(zldatabase.GetPara("�˿��ֹ��ʽ", glngSys, mlngModul))
    '���㷽ʽ:���|���㷽ʽ:���....
    mstr���տ� = zldatabase.GetPara("���տ�����", glngSys, mlngModul)
    mblnClearWinInfor = IIf(zldatabase.GetPara("��Ԥ���������Ϣ", glngSys, glngModul) <> "1", True, False)
    mblnδ��Ʋ���Ԥ�� = zldatabase.GetPara("����δ��Ʋ�׼��Ԥ��", glngSys, mlngModul, , , InStr(mstrPrivs, ";��������;") > 0) = "1"
    gblnSeekName = Nvl(zldatabase.GetPara("����ģ������", glngSys, mlngModul, 1)) = 1
    mblnסԺ��Ԥ����֤ = zldatabase.GetPara("סԺ��Ԥ����֤", glngSys, mlngModul, "0") = "1"
    mbln������Ժ��������˿� = zldatabase.GetPara("������Ժ��������˿�", glngSys, mlngModul, "1") = "1"
    'ˢ��Ҫ����������
    mblnCheckPass = Mid(zldatabase.GetPara(46, glngSys, , "0000000000"), 8, 1) = "1"
End Sub
Private Sub Form_Load()
    Dim lngH As Long
    
    Call InitPara
    mblnOK = False: mblnUnLoad = False
    
    'Ʊ�����ü�鼰��ʼ
    If mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3 Then
        mblnStartFactUseType = zlStartFactUseType(2)
        If mblnStartFactUseType = False Then
            If mFactProperty.intInvoicePrint <> 0 Then Call GetFact(True, mbytInState = 2)
        End If
    End If
    
    zlControl.PicShowFlat picInfo, -1
    zlControl.PicShowFlat picFace, -1
    
    Set mrsInfo = New ADODB.Recordset
    
    If Not InitUnit Then Unload Me: Exit Sub

    Call InitIDKind
    
    Call InitFace
    If mblnUnLoad Then Exit Sub
    
    lblTitle.Caption = gstrUnitName & "Ԥ�����"
    mstrCardPrivs = GetPrivFunc(glngSys, 1151)
    
    If (mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3) And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.��� & "��Ϊ������", mlngModul, gcnOracle
        
        Call zlCheckFactIsEnough
    End If
    If mbytInState = 0 Or mbytInState = 3 Then
        IDKind.IDKind = Val(zldatabase.GetPara("�ϴ����뷽ʽ", glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0))
    End If
    
    '81693:���ϴ�,2015/4/21,������
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear: On Error GoTo 0
    End If
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInState = 0: mstrInNO = ""
    mblnViewCancel = False: mblnUnLoad = False
    mlng����ID = 0: mstr�����ʻ� = "": mblnNOMoved = False
    mstr�˿����Ա = ""
    
    If (mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3) And gblnLED Then
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
'    Call zlCardSquareObject(True)
    Call SaveWinState(Me, App.ProductName)
    If mbytInState = 0 Or mbytInState = 3 Then
        zldatabase.SetPara "�ϴ����뷽ʽ", IDKind.IDKind, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    End If
    '�����:45666
    If mbytInState = 0 And cboType.Text = "סԺԤ��" Then
        zldatabase.SetPara "����ʾ����Ԥ��", chk����ʾ����Ԥ��.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    End If
End Sub

Private Sub InitPrepayType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Ԥ������
    '����:���˺�
    '����:2011-07-14 18:50:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytInState = 4 Then
        With cboType
            .Clear
            If InStr(1, mstrPrivs, ";����Ԥ��תסԺ;") > 0 Then
                .AddItem "����תסԺ": .ItemData(.NewIndex) = 1
                If mbytPrepayType = 1 Then .ListIndex = .NewIndex
            End If
            If InStr(1, mstrPrivs, ";סԺԤ��ת����;") > 0 Then
                .AddItem "סԺת����": .ItemData(.NewIndex) = 2
                If mbytPrepayType = 2 Then .ListIndex = .NewIndex
            End If
            
        End With
        lblԤ������.Caption = "תԤ��"
        If cboType.ListCount = 0 Then
            MsgBox "�㲻�߱�����Ԥ��תסԺ��סԺԤ��ת����Ȩ�ޣ�����ϵͳ����Ա��ϵ!", vbInformation + vbOKOnly, gstrSysName
            mblnUnLoad = True
        End If
        
        Exit Sub
    End If
    With cboType
        .Clear
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
        
     End With
End Sub

Private Sub InitFace(Optional blnSave As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ڲ������ô�����漰����״̬
    '����:���˺�
    '����:2011-07-17 10:36:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, i As Integer, strSQL As String
    Dim ctlTmp As Control
    
    If Not gobjSquare.objSquareCard Is Nothing And blnSave = False Then
        IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
    End If
    
    On Error GoTo errHandle
    
    strSQL = "Select ����, ����, ����, ȱʡ��־ From ����Ԥ��ժҪ Order by ����"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboNote.Clear
    If rsTmp.RecordCount > 0 Then
        While Not rsTmp.EOF
            cboNote.AddItem Nvl(rsTmp!����)
            rsTmp.MoveNext
        Wend
    End If
    
    cboNote.ListIndex = -1: Call InitPrepayType
    If mblnUnLoad Then Exit Sub
    
    IDKind.Enabled = (mbytInState = 0 Or mbytInState = 3 Or mbytInState = 4)
    Select Case mbytInState
        Case 0, 3 '��ȡԤ����,����˿�
            '����������
            Call CreateMobjCard
            cboNO.Text = ""
            txtDate.Text = Format(zldatabase.Currentdate(), "yyyy-MM-dd")
            txtMan.Text = UserInfo.����
            
            Call Load֧����ʽ
            '�˿�Ȩ��
            If InStr(mstrPrivs, ";Ԥ���˿�;") = 0 And InStr(mstrPrivs, ";���տ��˿�;") = 0 Or mbytInState = 3 Or mbytInState = 4 Then
                chkCancel.Visible = False
            End If
            'ֻ�д��տ���ȡȨ��
            If InStr(mstrPrivs, ";Ԥ���տ�;") = 0 Then Call cbo.Locate(cboStyle, 5, True)
            If mbytInState = 3 Then
                lblMoney.Caption = "�˿���": lblMoney.FontBold = True: lblMoney.ForeColor = vbRed
                txtMoney.ForeColor = vbRed: txtMoney.Font.Bold = True
            End If
            txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ�ݺ�") And gblnBillԤ�� '89302
        Case 1 'ָ���������
            picList.Visible = False
            Me.Height = Me.Height - picList.Height
            
            If mblnViewCancel Then lblFlag.Visible = True
            cmdSetup.Visible = False
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
        Case 2 'ָ�������˿�
            
            chkCancel.Value = Checked   '�ڵ��õ�click�¼��д��� picFace.Enabled = True '���������������˿����
            cmdSetup.Visible = False
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
                End If
            End If
        Case 4
            '����������
            Call CreateMobjCard
            chkCancel.Visible = False
            txtFact.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ�ݺ�") And gblnBillԤ�� '89302
    End Select
    
    If lbl�ʻ����.Visible = False Then lblԤ�����.Left = lbl�ʻ����.Left
    If lbl�ʻ����.Visible Then
        Line2(14).Visible = True: Line2(11).X2 = 2415
    Else
        Line2(14).Visible = False: Line2(11).X2 = Line2(14).X2
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub CreateMobjCard()
    '����������
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hWnd)
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

Private Sub txtFact_LostFocus()
'    If Not txtFact.Locked And txtFact.Text <> "" Then
'        txtFact.Text = Format(Left(txtFact.Text, gbytԤ��), String(gbytԤ��, "0"))
'    End If
End Sub

Private Sub txtMoney_Change()
    '����27363
    If IsNumeric(StrToNum(txtMoney.Text)) Then
        If mbytInState = 3 Then
            txtMoney.ForeColor = vbRed
        Else
            txtMoney.ForeColor = IIf(CCur(StrToNum(txtMoney.Text)) >= 0, vbBlue, vbRed)
        End If
    End If
End Sub

Private Sub txtMoney_GotFocus()
    txtMoney.SelStart = 0: txtMoney.SelLength = Len(txtMoney.Text)
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    '����27363
    If KeyAscii <> 13 Then
        If chkCancel.Value = Checked Or mbytInState = 3 Or mbytInState = 4 Then
            '�˿�ʱ���������븺��
            If KeyAscii = Asc(".") And InStr(txtMoney.Text, ".") > 0 Then KeyAscii = 0: Beep: Exit Sub
            If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        Else
            '�տ�ʱ����ͨ�������˿�
            If KeyAscii = Asc(".") And InStr(txtMoney.Text, ".") > 0 Then KeyAscii = 0: Beep: Exit Sub
            'Ȩ������
            If InStr(mstrPrivs, ";Ԥ���˿�;") = 0 Then
                If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            Else
                If (txtMoney.Text <> "" And txtMoney.SelLength <> Len(txtMoney.Text)) And KeyAscii = Asc("-") Then KeyAscii = 0: Beep: Exit Sub
                If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            End If
        End If
        If (txtMoney.Text <> "" And txtMoney.SelLength <> Len(Format(StrToNum(txtMoney.Text), "##,##0.00;-##,##0.00; ;"))) And _
            (Len(Format(StrToNum(txtMoney.Text), "##,##0.00;-##,##0.00; ;")) >= txtMoney.MaxLength) And _
            InStr(Chr(8), Chr(KeyAscii)) = 0 Then
            If txtMoney.SelLength > 0 And txtMoney.SelLength <= txtMoney.MaxLength Then
            Else
                KeyAscii = 0: Beep: Exit Sub
            End If
        End If
    Else
        If txtMoney.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtMoney_LostFocus()
    '����27363
    Dim dblMoney  As Double
    If Not IsNumeric(StrToNum(txtMoney.Text)) Then txtMoney.SetFocus: Exit Sub
    If mrsInfo.State = 1 And IsNumeric(StrToNum(txtMoney.Text)) Then
        txtMoney.Text = Format(StrToNum(txtMoney.Text), "##,##0.00;-##,##0.00; ;")
        If txtMoney.MaxLength > 12 Then txtMoney.MaxLength = 12
        '108813:���ϴ�,2017/5/8,������������
        If mbytInState = 4 Then Exit Sub
        If gblnLED Then
            '#22 1234.56   --Ԥ��һǧ������ʮ�ĵ�����Ԫ Y
            '#23 1234.56   --����һǧ������ʮ�ĵ�����Ԫ Z
            dblMoney = StrToNum(txtMoney.Text)
            If mbytInState = 3 Then dblMoney = -1 * dblMoney
            zl9LedVoice.Speak "#22 " & dblMoney
        End If
    End If
End Sub

Private Sub cboNO_GotFocus()
    If Not cboNO.Locked Then cboNO.SelStart = 0: cboNO.SelLength = Len(cboNO.Text)
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim strOper As String, vDate As Date
    
    If cboNO.Locked Then Exit Sub
    
    'ת���ɴ�д(���ֲ��ɴ���)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    ElseIf cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 11)
        
        '�Ƿ���ת������ݱ���,��¼����Ϊ1��ʾ�����Ԥ��
        If zldatabase.NOMoved("����Ԥ����¼", cboNO.Text, , "1", Me.Caption) Then
            If Not ReturnMovedExes(cboNO.Text, 6, Me.Caption) Then Exit Sub
            mblnNOMoved = False
        End If
        
        '����Ȩ��
        If Not ReadBillInfo(0, cboNO.Text, -2, strOper, vDate) Then
            cboNO.Text = "": cboNO.SetFocus: Exit Sub
        End If
        If Not BillOperCheck(6, strOper, vDate, "�˿�") Then
            cboNO.Text = "": cboNO.SetFocus: Exit Sub
        End If
        '����27363
        '��ȡҪ���ϵ�Ԥ�����
        Select Case ReadBill(cboNO.Text)
            Case -1
                If cboStyle.ItemData(cboStyle.ListIndex) = BalanceType.C5���տ� Then
                    If InStr(mstrPrivs, ";���տ��˿�;") = 0 Then
                        MsgBox "��û��Ȩ�޽��д��տ��˿������", vbInformation, gstrSysName
                        chkCancel.Value = 0
                    End If
                ElseIf InStr(mstrPrivs, ";Ԥ���˿�;") = 0 Then
                    MsgBox "��û��Ȩ�޽���Ԥ���˿������", vbInformation, gstrSysName
                    chkCancel.Value = 0
                Else
                    If HaveSpare(cboNO.Text) = 0 And InStr(mstrPrivs, ";Ԥ�������˿�;") = 0 Then
                        MsgBox "�ò�����û��Ԥ�����,��û��Ȩ���������ŵ��ݣ�", vbInformation, gstrSysName
                        chkCancel.Value = 0
                    ElseIf HaveBalance(cboNO.Text) <> 0 Then
                        MsgBox "�ñ�Ԥ���Ѿ��������ڽ���ʱʹ��,�㲻���������ŵ��ݣ�", vbInformation, gstrSysName
                        chkCancel.Value = 0
                    ElseIf Val(StrToNum(txtMoney.Text)) < 0 Then
                        MsgBox "�ñ�Ԥ�����Ϊ��,��ʾ�˿�,����ִ�иò�����", vbExclamation, gstrSysName
                        chkCancel.Value = 0
                    Else
                        If cmdOK.Enabled Then cmdOK.SetFocus
                    End If
                End If
            Case 0
                MsgBox "��ȡ��Ԥ�����ʧ�ܣ�", vbExclamation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus
            Case 1
                MsgBox "��Ԥ����ݲ����ڣ�", vbExclamation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus
            Case 2
                MsgBox "��Ԥ������Ѿ��˿", vbExclamation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus
            Case 3
                cboNO.Text = "": cboNO.SetFocus
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
    Dim blnCancel As Boolean
    Dim blnCard As Boolean, blnICCard As Boolean
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
Private Sub FindPati(ByVal objcard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnCancel As Boolean, bytPrepayType As Byte
    
    '��ȡ������Ϣ
    SetMoneyInfo True
    sta.Panels(2) = ""
    If objcard.���� Like "IC��*" And objcard.ϵͳ Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If Not GetPatient(objcard, strInput, blnCancel, blnCard) Then
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
        Set mrsInfo = New ADODB.Recordset
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Sub
    End If
    '���ò��˷�����Ϣ
    Call SetMoneyInfo(False, mrsInfo!����ID)
    Call LoadPatiPage(Val(Nvl(mrsInfo!����ID)))
    
    '79361:���ϴ�,2014/11/18,ȱʡ���˵�Ԥ������
    '���ò�ѯ�����ʿվ����ʱ���Զ��л�Ԥ�����ͣ��Դ����Ϊ׼
    bytPrepayType = IIf(Val(Nvl(mrsInfo!��Ժ)) = 1, 2, 1)
    If bytPrepayType <> mbytPrepayType And Not (mbytCallObject = 1 Or mblnNurseCall) Then
        mbytPrepayType = bytPrepayType: Call InitPrepayType
    End If
    
    If mrsInfo!��ǰ����id <> 0 Then
        lbl����.Caption = lbl����.Tag & IIf(mrsInfo!���� = 0, "��ͥ", mrsInfo!����)
    End If
            
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(Val(Nvl(mrsInfo!סԺ��)) = 0, "", "סԺ��:" & mrsInfo!סԺ�� & "   ") & _
                           IIf(Val(Nvl(mrsInfo!�����)) = 0, "", "�����:" & mrsInfo!�����)
    lbl����.Caption = lbl����.Tag & GET��������(mrsInfo!����ID)
    '46764
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(Val(Nvl(mrsInfo!��ǰ����id)) = 0, Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ǰ����id))))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
                
    'ҽ���Ķ�-��Ժ����ת�����ʻ�
    If Not IsNull(mrsInfo!����) And InStr(mstrPrivs, ";����ת��;") > 0 And mstr�����ʻ� <> "" Then
        If cbo.FindIndex(cboStyle, mstr�����ʻ�, True) = -1 Then
            cboStyle.AddItem mstr�����ʻ�
            cboStyle.ItemData(cboStyle.NewIndex) = 3
        End If
        'ҽ���ӿ�
        mcur�ʻ���� = gclsInsure.SelfBalance(mrsInfo!����ID, mrsInfo!ҽ����, 30, , mrsInfo!����)
        lbl�ʻ����.Caption = lbl�ʻ����.Tag & Format(mcur�ʻ����, "0.00")
        lbl�ʻ����.Visible = True
        lblԤ�����.Left = 2640
        If lbl�ʻ����.Visible Then
            Line2(14).Visible = True: Line2(11).X2 = 2415
        Else
            Line2(14).Visible = False: Line2(11).X2 = Line2(14).X2
        End If
    End If
    
    lbl�ѱ�ȼ�.Caption = lbl�ѱ�ȼ�.Tag & mrsInfo!�ѱ�
    lbl������.Caption = lbl������.Tag & mrsInfo!������
    lbl�������.Caption = lbl�������.Tag & mrsInfo!������
    '�����:116059,����,2017/12/7,Ԥ��������ʾ�����ֻ��ţ���ȡ������Ϣ�еġ��ֻ��š�
    lbl�ֻ���.Caption = lbl�ֻ���.Tag & mrsInfo!�ֻ���
    chk����temp.Value = mrsInfo!��������
    lblMemo.Caption = lblMemo.Tag & Nvl(mrsInfo!��ע)
    '72828,Ƚ����,2014-5-9,���ӹ�����λ��Ϣ����ʾ
    lblWorkUnit.Caption = lblWorkUnit.Tag & Nvl(mrsInfo!������λ)
    
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    txtPatient.Text = mrsInfo!����
    txtPatient.Tag = mrsInfo!����ID
    '-----------------------------------------------------------------------------------------
    lblSex.Caption = lblSex.Tag & IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
    mstrPatiSex = IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
    lblOld.Caption = lblOld.Tag & IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
    mstrPatiOld = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
    lbl��ͥ��ַ.Caption = lbl��ͥ��ַ.Tag & Nvl(mrsInfo!��ͥ��ַ)
    lblҽ�Ƹ��ʽ.Caption = lblҽ�Ƹ��ʽ.Tag & Nvl(mrsInfo!ҽ�Ƹ��ʽ)
    Call Led��ӭ��Ϣ
    Call SetcmdOkEnabled
    
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Led��ӭ��Ϣ()
    Dim strInfo As String, lngPatient As Long
    'LED��ʼ��
    If (mbytInState = 0 Or mbytInState = 3) And gblnLED Then
        If gblnLedWelcome Then
            zl9LedVoice.Reset com
            zl9LedVoice.Speak "#1"
            zl9LedVoice.Init UserInfo.��� & "��Ϊ������", mlngModul, gcnOracle
        End If
        strInfo = Trim(txtPatient.Text)
        If mrsInfo.State = 1 Then strInfo = strInfo & " " & mrsInfo!�Ա� & " " & mrsInfo!����: lngPatient = Val("" & mrsInfo!����ID)
        zl9LedVoice.DisplayPatient strInfo, lngPatient
    End If
End Sub
Private Sub Clear�����ʻ�()
'���ܣ���������ʻ���Ϣ
    Dim i As Integer
    For i = 0 To cboStyle.ListCount - 1
        If cboStyle.ItemData(i) = 3 Then
            cboStyle.RemoveItem i: Exit For
        End If
    Next
    mcur�ʻ���� = 0
    lbl�ʻ����.Caption = lbl�ʻ����.Tag
    lbl�ʻ����.Visible = False: Line2(14).Visible = False
    Line2(11).X2 = Line2(14).X2
    lblԤ�����.Left = lbl�ʻ����.Left
End Sub
 

Private Function GetPatient(ByVal objcard As Card, ByVal strInput As String, blnCancel As Boolean, Optional blnCard As Boolean = False, Optional lng��ҳID As Long) As Boolean
    '���ܣ���ȡ������Ϣ
    '������strInput=[ˢ��]|[A����ID]|[BסԺ��]
    '˵����
    '     1.�����ڲ���Ԥ����
    '     2.�Զ�ʶ������Ժ״̬,����(����ID,��ҳID,����,�Ա�,����,סԺ��,����,��Ժ��־)
    '����:�Ƿ��ȡ�ɹ�,�ɹ�ʱmrsInfo�а���������Ϣ,ʧ��ʱmrsInfo=Close
    Dim rsTmp As ADODB.Recordset, strPati As String, strSQL As String
    Dim vRect As RECT, i As Integer, lng�����ID As Long, bln�����ʻ� As Boolean, lng����ID As Long, strPassWord As String, strErrMsg As String
    Dim strWhere As String, blnICCard As Boolean
    Dim blnHavePassWord As Boolean
    Dim rsTemp As ADODB.Recordset, str��Ժ���� As String
    Dim blnIsMobileNO As Boolean
    Dim strRecent As String     '��ȡ���һ�β�����Ϣ����
    
    blnCancel = False
    strWhere = ""
    strRecent = " And Nvl(A.��ҳID,0)=C.��ҳID(+) "
    If lng��ҳID <> 0 Then
        strWhere = strWhere & " And A.����ID=[2] And C.��ҳID=[3]"
        strRecent = ""
        GoTo PatiPage
    End If
    blnIsMobileNO = IDKind.IsMobileNo(strInput)
    Call Clear�����ʻ� '��������ʻ���Ϣ
    mdblԤ�����_�������� = 0
    chkAllCash.Value = 0: chkAllCash.Visible = False
    mstr�˿����Ա = ""
    
    '112999
    If gblnAllowOut = False Then '�������Ժ���˽�סԺԤ��
        If mbytInState = 0 And cboType.ItemData(cboType.ListIndex) = 2 _
            Or mbytInState = 4 And cboType.ItemData(cboType.ListIndex) = 1 Then
            
            '��Ժ��Ԥ��Ժ
            str��Ժ���� = " And (Nvl(a.��Ժ, 0) = 1" & vbNewLine & _
                        "       Or Exists (Select 1 From ������ҳ" & vbNewLine & _
                        "                  Where ����id = a.����id And Nvl(��ҳid, 0) = 0 And Nvl(��������, 0) = 0)) "
        End If
    End If
    
    If (blnCard And objcard.���� Like "����*") _
        And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then   'ˢ����ȱʡ�Ŀ�
        lng�����ID = IDKind.GetDefaultCardTypeID
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then
            If blnIsMobileNO Then
                If gobjSquare.objSquareCard.zlGetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord) = False Then
                    GoTo NotFoundPati:
                End If
            Else
                GoTo NotFoundPati:
            End If
        End If
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.����ID=[1]"
        strInput = "-" & lng����ID
        blnHavePassWord = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then  '����ID
        strWhere = strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then  'סԺ��(��ס(��)Ժ�Ĳ���)
        strWhere = strWhere & " And (A.����ID,C.��ҳID) In (Select Max(����id),Max(��ҳID) From ������ҳ Where סԺ�� = [1])"
        strRecent = ""
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(�������ﲡ��)
        strWhere = strWhere & " And A.�����=[1]"
    Else '��������
        Select Case objcard.����
            Case "����", "��������￨"
                
                '����ģ���鳤��,��������ղ��һ�Ӱ������
                If (Not gblnSeekName) Or (gblnSeekName And Len(strInput) < 2) Then
                    Set mrsInfo = New ADODB.Recordset: Exit Function
                End If
                
                strPati = _
                " Select A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����," & _
                "           A.סԺ��,B.���� as ����,A.��ǰ���� as ����," & _
                "           A.��������,A.���֤��,A.��ͥ��ַ,A.����֤��,Nvl(A.��Ժ,0) As ��Ժ��־ " & _
                " From ������Ϣ A,���ű� B " & _
                " Where A.ͣ��ʱ�� is NULL And A.��ǰ����ID=B.ID(+) And A.���� Like [1] " & str��Ժ���� & _
                "   Order by A.����"
                
                
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTmp = zldatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", "bytSize=1")
                If Not rsTmp Is Nothing Then
                    strInput = rsTmp!����ID
                    strWhere = strWhere & " And A.����ID=[2]"
                Else
                    Set mrsInfo = New ADODB.Recordset: Exit Function
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                strWhere = strWhere & " And A.ҽ����=[2]"
            Case "IC����"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
                blnICCard = (InStr(1, "-+*.", Left(strInput, 1)) = 0) And objcard.ϵͳ
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And (A.����ID,C.��ҳID) In (Select Max(����id),Max(��ҳID) From ������ҳ Where סԺ�� = [2])"
                strRecent = ""
            Case Else
                '��������,��ȡ��صĲ���ID
                If objcard.�ӿ���� > 0 Then
                    lng�����ID = objcard.�ӿ����
                    bln�����ʻ� = objcard.�Ƿ�����ʻ�
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objcard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
PatiPage:
    '����27379
    '72828,Ƚ����,2014-5-9,���ӹ�����λ��Ϣ����ʾ������ʽ������A.������λ�ֶ�
    strSQL = _
    " Select A.����ID,Nvl(C.��ҳID,0) as ��ҳID,Nvl(C.��ǰ����ID,0) as ����ID,Nvl(c.��Ժ����ID,0) as ����ID,Nvl(A.��ǰ����ID,0) as ��ǰ����ID, Nvl(a.��Ժ,0) as ��Ժ," & _
    "           Decode(Nvl(A.��ҳID,0),0,A.ҽ�Ƹ��ʽ,C.ҽ�Ƹ��ʽ) ҽ�Ƹ��ʽ,Nvl(A.��������,C.��������) as ��������," & _
    "            Nvl(C.����, a.����) As ����, Nvl(C.�Ա�, a.�Ա�) As �Ա�,A.����,Nvl(A.�����,0) as �����,Nvl(C.סԺ��,0) as סԺ��,Nvl(C.��Ժ����,0) as ����,A.��ͥ��ַ,A.����֤��," & _
    "           B.����,B.����,Nvl(B.ҽ����,A.ҽ����) ҽ����,B.����,Nvl(C.�ѱ�,A.�ѱ�) �ѱ�,A.������,A.������,Nvl(A.��������,0) as ��������, A.������λ,A.�ֻ���,C.��ע,Nvl(A.��Ժ,0) As ��Ժ��־" & _
    " From ������Ϣ A,ҽ�����˵��� B,������ҳ C,ҽ�����˹����� E" & _
    " Where A.ͣ��ʱ�� is NULL" & _
    "       And A.����ID=C.����ID(+) " & strRecent & _
    "       And C.����ID=E.����ID(+) And E.��־(+)=1  " & str��Ժ���� & _
    "       And E.ҽ����=B.ҽ����(+) And E.����=B.����(+) And E.���� = B.����(+) " & strWhere
    
    On Error GoTo errH
    Set mrsInfo = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, lng��ҳID)
    If mrsInfo.EOF Then
        Set mrsInfo = New ADODB.Recordset: Exit Function
    End If
    '��Ҫ��������
    If mblnCheckPass And (blnCard Or blnICCard Or IDKind.GetCurCard.�ӿ���� <> 0) Then
        If Not blnHavePassWord Then
            strPassWord = Nvl(mrsInfo!����֤��)
        End If
        If strPassWord <> "" Then
            If zlCommFun.VerifyPassWord(Me, strPassWord, mrsInfo!����, mrsInfo!�Ա�, mrsInfo!����) = False Then
                 Set mrsInfo = New ADODB.Recordset: Exit Function
            End If
        End If
    End If
    GetPatient = True
    Exit Function
errH:
     If ErrCenter() = 1 Then Resume
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
End Function

Private Function SaveBill(Optional blnPrintInvoice As Boolean = False, Optional ByRef lngԤ��ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ե�ǰ�����Ԥ����ݴ���
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-17 11:15:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String, strSQL As String, i As Integer
    Dim blnInsure As Boolean, strCurDate As String
    Dim blnTrans As Boolean, dblMoney As Double
    Dim lng��ҳID As Long
    
    strNo = zldatabase.GetNextNo(11)
    lngԤ��ID = zldatabase.GetNextId("����Ԥ����¼")
    strCurDate = Format(zldatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    blnInsure = cboStyle.ItemData(cboStyle.ListIndex) = 3 And Not IsNull(mrsInfo!����)
    '����27363
    dblMoney = IIf(mbytInState = 3, -1, 1) * StrToNum(txtMoney.Text)
    
Once:
    'Zl_����Ԥ����¼_Insert
    strSQL = "Zl_����Ԥ����¼_Insert("
    '  Id_In         ����Ԥ����¼.ID%Type,
    strSQL = strSQL & "" & lngԤ��ID & ","
    '  ���ݺ�_In     ����Ԥ����¼.NO%Type,
    strSQL = strSQL & "'" & strNo & "',"
    '  Ʊ�ݺ�_In     Ʊ��ʹ����ϸ.����%Type,
    '60669
    If blnPrintInvoice Then
        strSQL = strSQL & "'" & txtFact.Text & "',"
    Else
        strSQL = strSQL & "NULL,"
    End If
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & mrsInfo!����ID & ","
    '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,:42329
    '����:44963
    'mrsInfo!��ǰ����id <> 0 And mrsInfo!��Ժ <> 0 And
    
    lng��ҳID = IIf(cboType.ItemData(cboType.ListIndex) = 2, Val(Nvl(mrsInfo!��ҳID)), 0)
    If cboPatiPage.Visible And cboPatiPage.ListIndex > 0 And cboType.ItemData(cboType.ListIndex) = 2 Then
        lng��ҳID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    strSQL = strSQL & "" & IIf(lng��ҳID = 0, "NULL", lng��ҳID) & ","
    
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & IIf(cboUnit.ItemData(cboUnit.ListIndex) = 0, "NULL", cboUnit.ItemData(cboUnit.ListIndex)) & ","
    '  ���_In       ����Ԥ����¼.���%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & mstr���㷽ʽ & "',"
    '  �������_In   ����Ԥ����¼.�������%Type,
    strSQL = strSQL & "'" & txtCode.Text & "',"
    '  �ɿλ_In   ����Ԥ����¼.�ɿλ%Type,
    If blnInsure Then
        strSQL = strSQL & "'" & Nvl(mrsInfo!����) & "',"
    Else
        strSQL = strSQL & "'" & Trim(txtUnit.Text) & "',"
    End If
    '  ��λ������_In ����Ԥ����¼.��λ������%Type,
    If blnInsure Then
        strSQL = strSQL & "'" & Nvl(mrsInfo!����) & "',"
    Else
        strSQL = strSQL & "'" & Trim(txt������.Text) & "',"
    End If
    '  ��λ�ʺ�_In   ����Ԥ����¼.��λ�ʺ�%Type,
    If blnInsure Then
        strSQL = strSQL & "'" & Nvl(mrsInfo!ҽ����) & "',"
    Else
        strSQL = strSQL & "'" & Trim(txt�ʺ�.Text) & "',"
    End If
    '  ժҪ_In       ����Ԥ����¼.ժҪ%Type,
    strSQL = strSQL & "'" & Trim(cboNote.Text) & "',"
    '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "" & IIf(mlng����ID = 0, "NULL", mlng����ID) & ","
    '  Ԥ�����_In   ����Ԥ����¼.Ԥ�����%Type := Null,
    strSQL = strSQL & "" & cboType.ItemData(cboType.ListIndex) & ","
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(mlngCardTypeID = 0 Or mbln���ѿ� Or chkAllCash.Value = 1, "NULL", mlngCardTypeID) & ","
   '  ���㿨���_in ����Ԥ����¼.���㿨���%type:=NULL,
    strSQL = strSQL & "" & IIf(mlngCardTypeID = 0 Or Not mbln���ѿ�, "NULL", mlngCardTypeID) & ","
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    If mbytInState = 3 Then
        strSQL = strSQL & "" & IIf(mbytInState = 3 And chkAllCash.Value = 0, "'" & mcurBill.str���� & "'", "NULL") & ","
    Else
        strSQL = strSQL & "" & IIf(mstrBrushCardNo = "", "NULL", "'" & mstrBrushCardNo & "'") & ","
    End If
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "" & IIf(mbytInState = 3 And chkAllCash.Value = 0, "'" & mcurBill.str������ˮ�� & "'", "NULL") & ","
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "" & IIf(mbytInState = 3, "'" & IIf(chkAllCash.Value = 1, mstr�˿����Ա & "ǿ������:" & Format(IIf(dblMoney < mdblԤ�����_��������, dblMoney, mdblԤ�����_��������), "0.00") & "Ԫ", mcurBill.str����˵��) & "'", "NULL") & ","
    '  ������λ_In   ����Ԥ����¼.������λ%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type := Null
    strSQL = strSQL & "to_date('" & strCurDate & "','yyyy-mm-dd hh24:mi:ss'),"
    '   ��������_In Integer:=0 :0-������Ԥ��;1-��Ϊ���۵�;3-����˿�
    strSQL = strSQL & IIf(mbytInState = 3, 3, 0) & ","
    '   ����id_In     ����Ԥ����¼.����id%Type := Null
    strSQL = strSQL & "NULL,"
    '   ��������_In   ����Ԥ����¼.��������%Type := Null,
    strSQL = strSQL & "NULL,"
    '   �˿���_In   Number := 0
    strSQL = strSQL & mbytOracleBackType & ","
    '   ǿ������_In   Number := 0
    strSQL = strSQL & IIf(mbytInState = 3 And chkAllCash.Value = 1, 1, 0) & ","
    '   ���½������_In Number := 1,
    strSQL = strSQL & "1,"
    '   �Ƿ�ת��_In     Number := 0
    strSQL = strSQL & IIf(mbytInState = 3 And mcurBill.blnת��, 1, 0) & ")"
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    zldatabase.ExecuteProcedure strSQL, Me.Caption
    If blnInsure Then
        'ҽ���ӿ�
        If Not gclsInsure.TransferSwap(lngԤ��ID, CCur(dblMoney), mrsInfo!����) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    End If
    If mbytInState = 3 Then
        If zlDepositDel(mcurBill.lngԤ��ID, lngԤ��ID, StrToNum(txtMoney.Text)) = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    Else
        If zlInterfacePrayMoney(lngԤ��ID, strNo, StrToNum(txtMoney.Text)) = False Then
            'ɾ����Ч��Ԥ������
            gcnOracle.RollbackTrans: Exit Function
            'Call DeletePrepay(strNO): Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    '���뵥����ʷ��¼(�������͵���)
    For i = 0 To cboNO.ListCount - 1
        strNo = strNo & "," & cboNO.List(i)
    Next
    cboNO.Clear
    For i = 0 To UBound(Split(strNo, ","))
        cboNO.AddItem Split(strNo, ",")(i)
        If i = 9 Then Exit For 'ֻ��ʾ10��
    Next
    
    If Not gblnBillԤ�� And blnPrintInvoice And Trim(txtFact.Text) <> "" Then
        '��ɢ�����浱ǰ����
        zldatabase.SetPara "��ǰԤ��Ʊ�ݺ�", Trim(txtFact.Text), glngSys, mlngFactModule
    End If
    SaveBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If Err.Description Like "*�˿�����ڲ���ʣ��Ԥ�����*" And mbytOracleBackType = 1 Then
        If MsgBox("�˿���Ȳ��˵�ǰ������,�Ƿ���ԣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
        mbytOracleBackType = 0
        GoTo Once
    End If
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ReadBill(strNo As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡԤ�����(����ġ��˿��),����д���漰����mrsInfo(������Ϣ),��������Tag��
    '���:strNO-Ԥ�����ݺ�
    '����:
    '����: -1-�ɹ�;0-ʧ��;1-�õ��ݲ�����;2:�õ����Ѿ��˿�(���ʱ��Ч);3-Ȩ�޲���(������)
    '����:���˺�
    '����:2011-07-15 11:45:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngPrepayType As Long, rsTemp As New ADODB.Recordset, strFullNO As String
    Dim strWhere As String, lngԤ����� As Long
    Dim i As Long, blnHave As Boolean, strTmp As String
    Dim rs��վ����ʾ As New ADODB.Recordset
    
    On Error GoTo errH
    strFullNO = GetFullNO(strNo, 11)
    If cboType.ListIndex >= 0 Then
        lngԤ����� = cboType.ItemData(cboType.ListIndex)
    End If
    
    strWhere = IIf(mbytInState = 1, IIf(mblnViewCancel, "And A.��¼״̬=2", "And A.��¼״̬ IN(1,3)"), "")
    If mbytCallObject = 1 Or mbytCallObject = 2 Or mbytInState = 3 Then
        strWhere = strWhere & " And Not Exists(Select 1 From ���㷽ʽ   Where A.���㷽ʽ= ���� And ����=5)"
    End If
    
    gstrSQL = "" & _
    "   Select   A.ID,A.Ԥ�����,A.ʵ��Ʊ��,A.����ID,A.��ҳID,A.����ID,A.��¼״̬,A.ժҪ,A.���, " & _
    "               A.���㷽ʽ,A.�������,A.�տ�ʱ��,A.����Ա����,A.�ɿλ,A.��λ������," & _
    "               A.��λ�ʺ�,A.�����ID,nvl(A.���㿨���,C.�ӿڱ��) as ���㿨���, " & _
    "               nvl(A.����,C.����) as ����,nvl(A.������ˮ��,C.������ˮ��) as ������ˮ��,A.����˵��,A.������λ," & _
    "               M.���� as ���������, nvl(J.����,Q.����) as ���ѿ�����,Nvl(M.�Ƿ��˿��鿨,0) as �Ƿ��˿��鿨,c.���ѿ�ID " & _
    "   From " & IIf(mblnNOMoved, "H", "") & "����Ԥ����¼ A, " & _
    "          ���˿������¼ C,���ѿ����Ŀ¼ J,���ѿ����Ŀ¼ Q,ҽ�ƿ���� M " & _
    "   Where  A.��¼����=1 And A.No=[1]   " & strWhere & _
    "          And A.ID=c.����ID(+) and C.�ӿڱ��=Q.���(+)" & _
    "          And A.�����ID=M.ID(+) And A.���㿨���=J.���(+)"
    
    '72828,Ƚ����,2014-5-9,���ӹ�����λ��Ϣ����ʾ������ʽ������B.������λ�ֶ�
    gstrSQL = _
    "Select A.ʵ��Ʊ�� as Ʊ�ݺ�,A.����ID,A.��ҳID,A.����ID,B.�����,B.סԺ��,nvl(D.����,B.����) as ����,nvl(D.�Ա�,B.�Ա�) as �Ա�,nvl(D.����,B.����) as ����," & _
    "           A.����ID As ��ǰ����ID,B.��ǰ����,B.��ͥ��ַ,A.ID,A.��¼״̬,A.ժҪ,A.���," & _
    "           A.���㷽ʽ,C.����,A.�������,A.�տ�ʱ��,A.����Ա����,B.��ͬ��λID," & _
    "           Decode(Nvl(A.��ҳID,0),0,B.ҽ�Ƹ��ʽ,D.ҽ�Ƹ��ʽ) ҽ�Ƹ��ʽ," & _
    "           Decode(Nvl(C.����,1),3,NULL,A.�ɿλ) as �ɿλ," & _
    "           Decode(Nvl(C.����,1),3,NULL,A.��λ������) as ��λ������," & _
    "           Decode(Nvl(C.����,1),3,NULL,A.��λ�ʺ�) as ��λ�ʺ�,Nvl(D.�ѱ�,B.�ѱ�) �ѱ�," & _
    "           B.������,B.������,Nvl(B.��������,0) as ��������, B.������λ,B.�ֻ���," & _
    "           B.��������,B.����," & _
    "           NVL(A.Ԥ�����,0) as Ԥ�����, " & _
    "           A.�����ID,A.���㿨���,A.����,A.������ˮ��,A.����˵��,A.������λ,A.���������, A.���ѿ�����,A.�Ƿ��˿��鿨,a.���ѿ�ID " & _
    " From (" & gstrSQL & ") A, ������Ϣ B,���㷽ʽ C,������ҳ D" & _
    " Where A.����ID=B.����ID  And B.����ID=D.����ID(+) And nvl(B.��ҳID,0)=D.��ҳID(+) And A.���㷽ʽ=C.����(+)"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFullNO)
    If rsTemp.RecordCount = 0 Then ReadBill = 1: Exit Function
    If mbytInState = 2 Or chkCancel.Value = 1 Then
        '�˿�,��Ҫ����Ƿ���ھ�����˿�Ȩ��
        lngPrepayType = Val(Nvl(rsTemp!Ԥ�����))
        If InStr(1, mstrPrivs, IIf(lngPrepayType = 1, ";����Ԥ��;", ";סԺԤ��;")) = 0 Then
            MsgBox "�㲻�߱���Ԥ�����ݽ����˿��Ȩ��,����ϵͳ����Ա��ϵ!", vbOKOnly + vbInformation, gstrSysName
            ReadBill = 3
            Exit Function
        End If
        
        If gbln��վ����ʾ Then
            strTmp = "Select 1 From ���ű� A, ������Ա B, ��Ա�� C" & vbNewLine & _
                    " Where a.Id = b.����id And b.��Աid = c.Id And (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� Is Null) And c.���� =[1]  And Rownum < 2"
    
            Set rs��վ����ʾ = zldatabase.OpenSQLRecord(strTmp, Me.Caption, Nvl(rsTemp!����Ա����))
            
            If rs��վ����ʾ.RecordCount = 0 Then
                MsgBox "��Ԥ�����ݲ����ڱ�վ��,�������˿�!", vbOKOnly + vbInformation, gstrSysName
                ReadBill = 3: Exit Function
            End If
        End If
    End If
    
    With mcurBill
        .strNo = strFullNO
        .lngԤ��ID = Val(Nvl(rsTemp!ID))
        .lng�����ID = IIf(Val(Nvl(rsTemp!�����ID)) = 0, Val(Nvl(rsTemp!���㿨���)), Val(Nvl(rsTemp!�����ID)))
        .bln���ѿ� = Val(Nvl(rsTemp!���㿨���)) <> 0
        .str���� = IIf(.bln���ѿ�, Nvl(rsTemp!���ѿ�����), Nvl(rsTemp!���������))
        .str���� = Nvl(rsTemp!����)
        .bln�˿��鿨 = Val(Nvl(rsTemp!�Ƿ��˿��鿨)) = 1
        .str������ˮ�� = Nvl(rsTemp!������ˮ��)
        .str����˵�� = Nvl(rsTemp!����˵��)
        .str������λ = Nvl(rsTemp!������λ)
        .dt�տ�ʱ�� = Format(rsTemp!�տ�ʱ��, "yyyy-MM-dd hh:mm:ss")
        .lng���ѿ�ID = Val(Nvl(rsTemp!���ѿ�ID))
    End With
    
    cboNO.Text = strFullNO
    cboNO.Tag = rsTemp!ID '�Դ�IDΪ׼�˿�
    txtPatient.Text = rsTemp!����
    txtPatient.Tag = rsTemp!����ID
    '74426:���ϴ�,2014-7-9,����������ʾ��ɫ����
    Call SetPatiColor(txtPatient, Nvl(rsTemp!��������), IIf(IsNull(rsTemp!����), &HFF0000, vbRed))
    lbl�ѱ�ȼ�.Caption = lbl�ѱ�ȼ�.Tag & rsTemp!�ѱ�
    
    lbl������.Caption = lbl������.Tag & rsTemp!������
    lbl�������.Caption = lbl�������.Tag & rsTemp!������
    lbl�ֻ���.Caption = lbl�ֻ���.Tag & rsTemp!�ֻ���
    chk����temp.Value = rsTemp!��������
    '72828,Ƚ����,2014-5-9,���ӹ�����λ��Ϣ����ʾ
    lblWorkUnit.Caption = lblWorkUnit.Tag & Nvl(rsTemp!������λ)
    
    cboUnit.ListIndex = cbo.FindIndex(cboUnit, IIf(IsNull(rsTemp!��ǰ����id), 0, rsTemp!��ǰ����id))
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    cboType.ListIndex = -1
    lngPrepayType = Val(Nvl(rsTemp!Ԥ�����))
    For i = 0 To cboType.ListCount - 1
         If cboType.ItemData(i) = lngPrepayType Then
            cboType.ListIndex = i: Exit For
         End If
     Next
     
     With cboType
        If cboType.ListIndex < 0 Then
           .AddItem IIf(lngPrepayType = 1, "����Ԥ��", "סԺԤ��")
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
     
    txtFact.Text = IIf(IsNull(rsTemp!Ʊ�ݺ�), "", rsTemp!Ʊ�ݺ�)
    txtUnit.Text = IIf(IsNull(rsTemp!�ɿλ), "", rsTemp!�ɿλ)
    txt������.Text = IIf(IsNull(rsTemp!��λ������), "", rsTemp!��λ������)
    txt�ʺ�.Text = IIf(IsNull(rsTemp!��λ�ʺ�), "", rsTemp!��λ�ʺ�)
    
    lblPatientNO.Caption = lblPatientNO.Tag & IIf(Val(Nvl(rsTemp!סԺ��)) = 0, "", "סԺ��:" & rsTemp!סԺ�� & "   ") & _
                           IIf(Val(Nvl(rsTemp!�����)) = 0, "", "�����:" & rsTemp!�����)
    lblSex.Caption = lblSex.Tag & IIf(IsNull(rsTemp!�Ա�), "", rsTemp!�Ա�)
    mstrPatiSex = IIf(IsNull(rsTemp!�Ա�), "", rsTemp!�Ա�)
    lblOld.Caption = lblOld.Tag & IIf(IsNull(rsTemp!����), "", rsTemp!����)
    mstrPatiOld = IIf(IsNull(rsTemp!����), "", rsTemp!����)
    lbl����.Caption = lbl����.Tag & IIf(IsNull(rsTemp!��ǰ����), "", rsTemp!��ǰ����)
    lbl����.Caption = lbl����.Tag & GET��������(IIf(IsNull(rsTemp!��ǰ����id), 0, rsTemp!��ǰ����id))
    lbl��ͥ��ַ.Caption = lbl��ͥ��ַ.Tag & Nvl(rsTemp!��ͥ��ַ)
    lblҽ�Ƹ��ʽ.Caption = lblҽ�Ƹ��ʽ.Tag & Nvl(rsTemp!ҽ�Ƹ��ʽ)
    txtMoney.Text = Format(rsTemp!���, "##,##0.00;-##,##0.00;;")
    txtMoney.Tag = rsTemp!���
    If mcurBill.lng�����ID <> 0 Then
        cboStyle.ListIndex = cbo.FindIndex(cboStyle, mcurBill.str����, True)
    Else
        cboStyle.ListIndex = cbo.FindIndex(cboStyle, IIf(IsNull(rsTemp!���㷽ʽ), "", rsTemp!���㷽ʽ), True)
     End If
    If cboStyle.ListIndex = -1 Then
        If mcurBill.lng�����ID <> 0 Then
            cboStyle.AddItem mcurBill.str����
            cboStyle.ItemData(cboStyle.NewIndex) = -1
        Else
            cboStyle.AddItem IIf(IsNull(rsTemp!���㷽ʽ), "", rsTemp!���㷽ʽ)
            cboStyle.ItemData(cboStyle.NewIndex) = Val("" & rsTemp!����)
        End If
        cboStyle.ListIndex = cboStyle.NewIndex
        
    End If
    
    txtCode.Text = IIf(IsNull(rsTemp!�������), "", rsTemp!�������)
    txtMan.Text = IIf(IsNull(rsTemp!����Ա����), "", rsTemp!����Ա����)
    txtDate.Text = Format(rsTemp!�տ�ʱ��, "yyyy-MM-dd")
    cboNote.Text = IIf(IsNull(rsTemp!ժҪ), "", rsTemp!ժҪ)
    '��ȡ���˷�����Ϣ
    Call SetMoneyInfo(False, rsTemp!����ID, strNo)
    ReadBill = -1
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub GetDepositData(ByVal lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¶�ȡԤ������
    '���:lng����ID-����ID��
    '����:���˺�
    '����:2011-07-22 17:02:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, intԤ������ As Integer
    Dim strWhere As String, strTittle As String
    
    On Error GoTo errHandle
    If lng����ID = 0 Then
        If mrsInfo Is Nothing Then Set mrsDepositBalance = Nothing: Exit Sub
        If mrsInfo.State <> 1 Then Set mrsDepositBalance = Nothing: Exit Sub
        lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    
    mdbl������� = 0: mdblԤ����� = 0: mdblʣ���� = 0
    '������Ȼ���,���������
    Set mrsDepositBalance = GetMoneyInfo(lng����ID, , , , True)
    '�����,�ֱ����ͳ��,ֻ���˿�ʱ,�Żᷢ��
    If mbytInState <> 3 Then Exit Sub
    strSQL = "" & _
    "   Select A.Ԥ�����,nvl(A.�����ID,0) as �����ID,nvl(A.���㿨���,0) as ���㿨���, " & _
    "           A.����,A.������ˮ��,A.����˵��," & _
    "           max(decode(sign(���),-1, decode(A.��¼״̬,1,0,2,0,ID),ID)) as Ԥ��ID," & _
    "           nvl(sum(���),0)-nvl(sum(nvl(��Ԥ��,0)),0) as Ԥ����� " & _
    "   From ����Ԥ����¼ A " & _
    "   Where   A.����ID=[1] and (nvl(A.���㿨���,0)<>0 or nvl(�����ID,0)<>0) " & _
    "   Group by A.Ԥ�����,nvl(A.�����ID,0),nvl(A.���㿨���,0),A.����,A.������ˮ��,A.����˵��" & _
    "   Having nvl(sum(���),0)-nvl(sum(nvl(��Ԥ��,0)),0)  <>0"
        
    strSQL = "" & _
    "   Select RowNum as ID,A.Ԥ�����, A.Ԥ��ID, " & _
    "           A.�����ID,A.���㿨��� as ���ѽӿ�ID, " & _
    "          nvl(B.����,C.���) as ����,nvl(B.����,C.����) as ����, " & _
    "          Decode(B.����,NULL,C.�Ƿ�ȫ��,B.�Ƿ�ȫ��) as �Ƿ�ȫ��," & _
    "          Decode(B.����,NULL,C.�Ƿ�����,B.�Ƿ�����) as �Ƿ�����," & _
    "          A.����,A.������ˮ��,A.����˵��," & _
    "          A.Ԥ����� " & _
    "   From (" & strSQL & ") A,ҽ�ƿ���� B,���ѿ����Ŀ¼ C" & _
    "   Where   A.�����ID=B.ID(+)  and A.���㿨���=C.���(+)  and nvl(A.Ԥ�����,0)>0" & _
    "   Order by ����,A.����,A.������ˮ��,A.����˵��"
    Set mrsDepositInfor = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ShowPremayBalance(ByVal blnreReadData As Boolean, ByVal lng����ID As Long, _
    Optional ByVal blnNotSelect�����˿� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������صĽ��㷽ʽ����������,��ʾԤ�����
    '���:blnReRead-�ض�����
    '       lng����ID-��ȡָ���Ĳ���ID(0ʱ,��mrsInfo��¼�ж�ȡ����ID)
    '����:���˺�
    '����:2011-07-21 15:44:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset, strSQL As String
    Dim intԤ������ As Integer, bln�����ӿ� As Boolean
    Dim strWhere As String, strTittle As String, strPrevName As String
    Dim dbl��� As Double, strTemp As String, strBlance As String, str���� As String
    Dim strNotBalance As String, strCardBalance As String, blnPrevCash As Boolean
    Dim dbl���� As Double, dbl�������� As Double, dblΪ�� As Double, dblδ�� As Double, dblYB As Double
    Dim lng��ҳID As Long
    
    On Error GoTo errHandle
    If lng����ID = 0 Then
        If mrsInfo Is Nothing Then Exit Sub
        If mrsInfo.State <> 1 Then Exit Sub
        lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    
    If blnreReadData Then Call GetDepositData(lng����ID)
    If cboStyle.ListIndex < 0 And mbytInState <> 4 Then Exit Sub
    sta.Panels(2).Text = ""
    mdbl������� = 0: mdblԤ����� = 0: mdblʣ���� = 0: mdblԤ�����_���� = 0
    intԤ������ = cboType.ItemData(cboType.ListIndex)
    If mbytInState = 4 Then
        bln�����ӿ� = False
    Else
        bln�����ӿ� = cboStyle.ItemData(cboStyle.ListIndex) = -1
    End If
    strWhere = "And nvl(�����ID,0)<>0 or nvl(���㿨���,0)<>0 "
    If mbytInState = 3 Then
            mcurBill.bln���ѿ� = False
            mcurBill.lng�����ID = 0
            mcurBill.str������ˮ�� = ""
            mcurBill.str����˵�� = ""
            mcurBill.str���� = ""
            mcurBill.bln�˿��鿨 = False
            If bln�����ӿ� Then
                If mbln���ѿ� Then
                    mrsDepositInfor.Filter = "Ԥ�����=" & intԤ������ & " and  ���ѽӿ�ID=" & mlngCardTypeID
                Else
                    mrsDepositInfor.Filter = "Ԥ�����=" & intԤ������ & " and �����ID=" & mlngCardTypeID
                End If
                sta.Panels(2).Text = ""
                If mrsDepositInfor.RecordCount <> 0 Then
                    strTemp = "": strTittle = "": strBlance = ""
                    With mrsDepositInfor
                        .Sort = "����,������ˮ��,����˵��"
                        dbl��� = 0
                        Do While .EOF = False
                            'A.����,A.������ˮ��,A.����˵��
                            str���� = Nvl(!����): strTemp = Nvl(!����)
                            If strTemp <> strBlance Then
                                If strBlance <> "" Then
                                    strTittle = strTittle & strBlance & ":" & Format(dbl���, "###0.00;-###0.00;;") & Space(2)
                                End If
                                strBlance = strTemp: dbl��� = 0
                            End If
                            dbl��� = dbl��� + Val(Nvl(!Ԥ�����, 0))
                            .MoveNext
                        Loop
                        If strBlance <> "" Then
                            strTittle = strTittle & strBlance & ":" & Format(dbl���, "###0.00;-###0.00;;") & Space(2)
                        End If
                        
                        sta.Panels(2).Text = str���� & ":" & strTittle
                    End With
                End If
                If blnNotSelect�����˿� = False Then Call Select�����˿�
            Else
                mrsDepositInfor.Filter = "Ԥ�����=" & intԤ������
                If mrsDepositInfor.RecordCount <> 0 Then
                    strTemp = "": strTittle = "": strBlance = ""
                    With mrsDepositInfor
                        .Sort = "�����ID,���ѽӿ�ID,����"
                        dbl��� = 0
                        Do While .EOF = False
                            'A.����,A.������ˮ��,A.����˵��
                            strTemp = Nvl(!�����ID) & "-" & Nvl(!���ѽӿ�ID) & "-" & Nvl(!����)
                            If strTemp <> strBlance Then
                                If strBlance <> "" Then
                                    If blnPrevCash Then
                                        strCardBalance = strCardBalance & strPrevName & ":" & Format(dbl���, "###0.00;-###0.00;;") & Space(2)
                                    Else
                                        strNotBalance = strNotBalance & strPrevName & ":" & Format(dbl���, "###0.00;-###0.00;;") & Space(2)
                                    End If
                                End If
                                blnPrevCash = Val(Nvl(!�Ƿ�����)) = 1
                                strPrevName = Nvl(!����)
                                strBlance = strTemp: dbl��� = 0
                            End If
                            str���� = Nvl(!����)
                            dbl��� = dbl��� + Val(Nvl(!Ԥ�����, 0))
                            If Nvl(!�Ƿ�����) <> 1 Then
                                mdblԤ�����_���� = mdblԤ�����_���� + Val(Nvl(!Ԥ�����, 0))
                                dbl�������� = dbl�������� + Val(Nvl(!Ԥ�����, 0))
                            Else
                                dbl���� = dbl���� + Val(Nvl(!Ԥ�����, 0))
                            End If
                            .MoveNext
                        Loop
                        If dbl��� <> 0 Then
                            If blnPrevCash Then
                                strCardBalance = strCardBalance & strPrevName & ":" & Format(dbl���, "###0.00;-###0.00;;") & Space(5)
                            Else
                                strNotBalance = strNotBalance & strPrevName & ":" & Format(dbl���, "###0.00;-###0.00;;") & Space(5)
                            End If
                        End If
                        If strBlance <> "" Then strTittle = strTittle & _
                            IIf(strCardBalance = "", "", "��������" & Format(dbl����, "###0.00;-###0.00;;") & "Ԫ,����:" & strCardBalance) & _
                            IIf(strNotBalance = "", "", "����������" & Format(dbl��������, "###0.00;-###0.00;;") & "Ԫ,����:" & strNotBalance)
                        sta.Panels(2).Text = strTittle
                    End With
                End If
                If mdblԤ�����_���� <> 0 Then
                    mdblԤ�����_�������� = mdblԤ�����_����
                    chkAllCash.Visible = True
                    chkAllCash.Enabled = True
                End If
            End If
    End If
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
        Set rsMoney = zldatabase.OpenSQLRecord(strSQL, "��ȡҽ��Ԥ��", lng����ID)
    Else
        strSQL = "Select Sum(���) As ҽ��Ԥ�� From ����ģ����� Where ����ID = [1] And ��ҳID = [2]"
        Set rsMoney = zldatabase.OpenSQLRecord(strSQL, "��ȡҽ��Ԥ��", lng����ID, lng��ҳID)
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

    mdblʣ���� = mdblԤ����� - mdbl�������
    '����27363
    lbl�������.Caption = lbl�������.Tag & Format(mdbl�������, "##,##0.00;-##,##0.00; ;")
    lblԤ�����.Caption = lblԤ�����.Tag & Format(mdblԤ�����, "##,##0.00;-##,##0.00; ;")
    dblΪ�� = GetUnAuditedFee(lng����ID, , intԤ������)
    dblδ�� = GetUnAuditedFee(lng����ID, False, intԤ������)
    lblδ�����.Caption = lblδ�����.Tag & Format(dblΪ��, "##,##0.00;-##,##0.00; ;")
    lblδ�ɷ���.Caption = lblδ�ɷ���.Tag & Format(dblδ��, "##,##0.00;-##,##0.00; ;")
    lblʣ����.Caption = lblʣ����.Tag & Format(mdblʣ���� - dblδ�� - dblΪ�� + dblYB, "##,##0.00;-##,##0.00; ;")
    If mbytInState = 3 Then
        If bln�����ӿ� Then
            If mdblʣ���� - mdblԤ�����_���� >= 0 Then
                txtMoney.Text = IIf(mdblԤ�����_���� > 0, Format(mdblԤ�����_����, "#,##0.00;-#,##0.00;;"), "0.00")
            Else
                txtMoney.Text = IIf(mdblʣ���� > 0, Format(mdblʣ����, "#,##0.00;-#,##0.00;;"), "0.00")
            End If
        Else
            If mdblʣ���� - mdblԤ�����_���� >= 0 Then
                txtMoney.Text = Format(mdblʣ���� - mdblԤ�����_����, "#,##0.00;-#,##0.00;;")
            Else
                txtMoney.Text = "0.00"
            End If
        End If
    Else
        If mbytInState = 4 Then
            txtMoney.Text = Format(mdblʣ����, "#,##0.00;-#,##0.00;;")
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetMoneyInfo(blnClear As Boolean, Optional lng����ID As Long, _
    Optional strBackNo As String = "", Optional ByVal blnNotSelect�����˿� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ϣ
    '���:blnClear-���
    '     lng����ID-ָ������ID
    '     strBackNO-ָ����Ԥ������(�˿�ʱ����,��Ҫ���Ƕ�λ���嵥����ȥ)
    '����:���˺�
    ' �޸�:���˺�(�˺�ʱ,���Ӷ�λ����),���Ӳ���;strBackNo
    '����:2011-07-21 15:40:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset, lng��ҳID As Long
    Dim strSQL As String, lngRow As Long
    
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
        chk����temp.Value = 0
        '72828,Ƚ����,2014-5-9,���ӹ�����λ��Ϣ����ʾ
        lblWorkUnit.Caption = lblWorkUnit.Tag
        
        lblδ�����.Caption = lblδ�����.Tag
        lblδ�ɷ���.Caption = lblδ�ɷ���.Tag
        lbl�������.Caption = lbl�������.Tag
        lblԤ�����.Caption = lblԤ�����.Tag
        lblʣ����.Caption = lblʣ����.Tag
        lblҽ��Ԥ��.Caption = lblҽ��Ԥ��.Tag
        lbl�ֻ���.Caption = lbl�ֻ���.Tag
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
        '��ʾԤ�����
        Call ShowPremayBalance(True, lng����ID, blnNotSelect�����˿�)
        '����Ƿ���Ӧ�տ�
        strSQL = "Select Zl_Patientdue([1]) ʣ��Ӧ�� From dual"
        Set rsMoney = New ADODB.Recordset
        Set rsMoney = zldatabase.OpenSQLRecord(strSQL, "��ȡӦ�տ�", lng����ID)
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
    '����:��ʾ��ʷ��Ԥ������
    '����:���˺�
    '����:2011-09-16 10:17:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, int���� As Integer, lngRow As Long, strWhere As String
    Dim rsMoney As ADODB.Recordset
    Dim lng����ID As Long
    If mrsInfo Is Nothing Then
        lng����ID = mlng����ID
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = mlng����ID
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    
    If cboType.ListIndex < 0 Then
         int���� = 1 'cboType.ItemData(cboType.ListIndex)
    Else
        int���� = cboType.ItemData(cboType.ListIndex)
    End If
    
    On Error GoTo errHandle
    '84217,���ϴ�,2015/4/22,��ʾָ����סԺ�ڼ���ɵ�Ԥ��
    If cboType.Text = "סԺԤ��" And chk����ʾ����Ԥ��.Value = 1 And cboPatiPage.ListIndex >= 0 Then
        strWhere = " And A.��ҳID= " & cboPatiPage.ItemData(cboPatiPage.ListIndex)
    End If
    
    If gbln��վ����ʾ Then
        strWhere = strWhere & _
                " And Exists (Select 1 From ��Ա�� C, ������Ա D, ���ű� E " & _
                " Where C.���� =A.����Ա���� And C.Id = D.��Աid And D.����id = E.Id And (E.վ�� = '" & gstrNodeNo & "' Or E.վ�� Is Null))"
    End If
    
    If gblnShowHave Then
        'ֻ��ʾ��ʣ�����ʷ�ɿ�
        '���Ӳ�������������һ�ν���ʱ��һ��һ��
        strSQL = _
        "   Select NO,Sum(Nvl(A.���,0)) as ���  " & _
        "    From " & IIf(mblnNOMoved, "H", "") & "����Ԥ����¼ A" & _
        "   Where A.����ID Is Null And Nvl(A.���, 0)<>0 And A.����ID=[1] And A.Ԥ�����=[2] " & _
        "   Group by NO " & _
        "   Having Sum(Nvl(A.���,0))<>0"
        
        strSQL = _
        " Select LTrim(To_Char(A.�տ�ʱ��,'YYYY-MM-DD')) as ����,A.NO as ���ݺ�," & _
        "           C.���� as ����,Ltrim(To_Char(Nvl(A.���,0),'9,999,999,990.00')) as ʣ����,A.���㷽ʽ as ����,A.����Ա���� as �տ���" & _
        " From " & IIf(mblnNOMoved, "H", "") & "����Ԥ����¼ A,(" & strSQL & ") B,���ű� C" & _
        " Where A.����ID Is Null And A.Ԥ�����=[2]  And Nvl(A.���,0)<>0 And A.����ID=C.ID(+)" & _
        "       And A.���㷽ʽ Not IN(Select ���� From ���㷽ʽ Where ����=5)" & _
        "       And A.NO=B.NO And A.����ID=[1] " & strWhere & _
        " Union All" & _
        " Select Min(LTrim(To_Char(A.�տ�ʱ��,'YYYY-MM-DD'))) as ����,A.NO as ���ݺ�," & _
        "           B.���� as ����,Ltrim(To_Char(Sum(Nvl(A.���,0)-Nvl(A.��Ԥ��,0)),'9,999,999,990.00')) as ʣ����,A.���㷽ʽ as ����,A.����Ա���� as �տ���" & _
        " From " & IIf(mblnNOMoved, "H", "") & "����Ԥ����¼ A,���ű� B" & _
        " Where A.��¼���� IN(1,11) And A.����ID is Not NULL And A.����ID=B.ID(+) And A.Ԥ�����=[2] " & _
        "       And Nvl(A.���,0)<>Nvl(A.��Ԥ��,0) And A.����ID=[1] " & strWhere & _
        " Having Sum(Nvl(A.���,0)-Nvl(A.��Ԥ��,0))<>0" & _
        " Group by A.NO,B.����,A.���㷽ʽ,A.����Ա����" & _
        " Order by ����,���ݺ�,����"
    Else
        '������ʷ�ɿ���ϸ�嵥
        strSQL = _
        " Select Ltrim(To_Char(A.�տ�ʱ��,'YYYY-MM-DD')) as ����,A.NO as ���ݺ�,B.���� as ����, " & _
        " Ltrim(To_Char(A.���,'9,999,999,990.00')) as �ɿ���,A.���㷽ʽ as ����,A.����Ա���� as �տ��� " & _
        " From " & IIf(mblnNOMoved, "H", "") & "����Ԥ����¼ A,���ű� B" & _
        " Where A.����ID=B.ID(+) And A.��¼����=1 And A.����ID=[1]  And A.Ԥ�����=[2] " & strWhere & _
        " Order by A.�տ�ʱ�� Desc"
    End If
    
    Set rsMoney = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, int����)
    mshList.Clear
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
    If mrsInfo.State = 1 Then
        mstr�������� = IIf(IsNull(mrsInfo!��������), "", mrsInfo!��������)
    End If
    If mstr�������� = "" Then
        If mrsInfo.State = 1 Then
            If GetOutPatient(mrsInfo!����ID) Then
                txtPatient.ForeColor = vbRed
            Else
                txtPatient.ForeColor = &HFF0000
            End If
        Else
            txtPatient.ForeColor = &HFF0000
        End If
    Else
        txtPatient.ForeColor = zldatabase.GetPatiColor(mstr��������, True)
    End If
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
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
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

Private Function CancelBill(lngID As Long, blnCanDel As Boolean, intInsure As Integer, bln��ӡ As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ��ID��Ԥ�����ִ���˿��
    '���:lngID=����ID
    '        blnCanDel=�Ƿ�֧���˸����ʻ�
    '        intInsure=��������ʹ�õĸ����ʻ��ı������,��Ϊ0
    '����:
    '����:
    '����:���˺�
    '����:2011-07-19 09:28:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, blnTrans As Boolean
    Dim lng��Ԥ��ID As Long
    If mcurBill.lng�����ID <> 0 Then
        lng��Ԥ��ID = zldatabase.GetNextId("����Ԥ����¼")
    End If
    On Error GoTo errH
    '111209:л�٣�2017/07/24��Ԥ��������ϴ�ӡ��Ʊʱ,������Ԥ����
    strSQL = "zl_����Ԥ����¼_DELETE(" & lngID & ",'" & cboNote.Text & "','" & _
        UserInfo.��� & "','" & UserInfo.���� & "'," & IIf(blnCanDel, 1, 0) & "," & IIf(lng��Ԥ��ID = 0, "NULL", lng��Ԥ��ID) & "," & _
        "'" & IIf(bln��ӡ, mstrRedFact, "") & "'," & IIf(bln��ӡ, IIf(mlng����ID > 0, mlng����ID, "Null"), 0) & ")"
    gcnOracle.BeginTrans: blnTrans = True
    zldatabase.ExecuteProcedure strSQL, Me.Caption
    '����ҽ���ӿ�
    If intInsure <> 0 And blnCanDel Then
        If Not gclsInsure.TransferDelSwap(lngID, intInsure) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    End If
    If zlDepositDel(lngID, lng��Ԥ��ID, StrToNum(txtMoney.Text)) = False Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    
    If Not gblnBillԤ�� And bln��ӡ And mstrRedFact <> "" Then
        '��ɢ�����浱ǰ����
        zldatabase.SetPara "��ǰԤ��Ʊ�ݺ�", mstrRedFact, glngSys, mlngFactModule
    End If
    CancelBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOutPatient(ByVal lngID As Long) As Boolean
'���ܣ��ж����ﲡ���Ƿ�����ҽ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim int���� As Integer
    
    GetOutPatient = False
    On Error GoTo errH
    
    strSQL = _
        "Select ���� " & _
        "from ������Ϣ " & _
        "Where ����id = [1] and rownum <= 1 "

    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    
    If Not rsTmp.EOF Then
        int���� = IIf(IsNull(rsTmp!����), -1, rsTmp!����)
        GetOutPatient = int���� <> -1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:������رս��㿨����
'    '���:blnClosed:�رն���
'    '����:���˺�
'    '����:2010-01-05 14:51:23
'    '����:
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strExpend As String
'
'   If mbytInState = 1 Then Exit Sub
'    'ֻ��:ִ�л��˷�ʱ,�ſ��ܹܽ��㿨��
'    If blnClosed Then
'FromClose:
'        If Not mobjSquareCard Is Nothing Then
'            Call mobjSquareCard.CloseWindows
'            Set mobjSquareCard = Nothing
'        End If
'        Exit Sub
'    End If
'    '��������
'    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
'    Err = 0: On Error Resume Next
'    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
'    If Err <> 0 Then
'        mtySquareCard.blnExistsObjects = False
'        Exit Sub
'    End If
'    Dim strKind As String
'
'    '��װ�˽��㿨�Ĳ���
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    '����:zlInitComponents (��ʼ���ӿڲ���)
'    '    ByVal frmMain As Object, _
'    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
'    '        ByVal cnOracle As ADODB.Connection, _
'    '        Optional blnDeviceSet As Boolean = False, _
'    '        Optional strExpand As String
'    '����:
'    '����:   True:���óɹ�,False:����ʧ��
'    '����:���˺�
'    '����:2009-12-15 15:16:22
'    'HIS����˵��.
'    '   1.���������շ�ʱ���ñ��ӿ�
'    '   2.����סԺ����ʱ���ñ��ӿ�
'    '   3.����Ԥ����ʱ
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    If mobjSquareCard.zlInitComponents(Me, mlngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend) Then
'        mtySquareCard.blnExistsObjects = True
'        mobjSquareCard.mblnYLMgr = mbytCallObject = 2
'    End If
'    strKind = "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;ס|סԺ��|0;��|���￨|0"
'    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, mobjSquareCard, strKind, txtPatient)
'End Sub

Private Sub InitIDKind()
    Dim strKind As String
    strKind = "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;ס|סԺ��|0;��|���￨|0;��|�ֻ���|0"
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKind, txtPatient)
    mtySquareCard.blnExistsObjects = Not gobjSquare.objSquareCard Is Nothing
    gobjSquare.objSquareCard.mblnYLMgr = mbytCallObject = 2
End Sub

Private Function zlCheckDepositDelValied(ByRef lngԤ��ID As Long, _
    ByVal dbl�˿��� As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷ѽ��׽ӿ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExend As String
    Dim cllSquareBalance As Collection
    
    If mcurBill.lng�����ID = 0 Then zlCheckDepositDelValied = True: Exit Function
    
    If Not mtySquareCard.blnExistsObjects Or gobjSquare.objSquareCard Is Nothing Then
            MsgBox "ע��:" & vbCrLf & _
                         "      ��ǰ��Ԥ���" & mcurBill.str���� & " �����,�������ڲ�������ز���,�����˿�,����ϵͳ����Ա��ϵ!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
            Exit Function
    End If
    
    If mbytInState = 3 And mcurBill.blnת�� Then
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, mcurBill.lng�����ID, False, Nvl(mrsInfo!����), Nvl(mrsInfo!�Ա�), Nvl(mrsInfo!����), dbl�˿���, mstrBrushCardNo, mstrbrPassWord, False, False, False, False) = False Then Exit Function
        mcurBill.str���� = mstrBrushCardNo
        zlXML.ClearXmlText
        zlXML.AppendNode "IN"
            zlXML.appendData "CZLX", "4"
        zlXML.AppendNode "IN", True
        strXMLExend = zlXML.XmlText
        zlXML.ClearXmlText
        If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, mcurBill.lng�����ID, _
            mcurBill.str����, dbl�˿���, "", strXMLExend) = False Then
            zlCheckDepositDelValied = False
            Exit Function
        End If
    Else
        Set cllSquareBalance = New Collection
        'Array(�����ID,���ѿ�ID,ˢ�����, ����,����,�������,�Ƿ�����,ʣ��δ�˽��)
        cllSquareBalance.Add Array(mcurBill.lng�����ID, mcurBill.lng���ѿ�ID, 0, mcurBill.str����, "", "", False, dbl�˿���)
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
        If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, mcurBill.lng�����ID, mcurBill.bln���ѿ�, mcurBill.str����, _
            "1|" & lngԤ��ID, dbl�˿���, mcurBill.str������ˮ��, mcurBill.str����˵��, strXMLExend) = False Then
              zlCheckDepositDelValied = False
              Exit Function
         End If
         '100610:���ϴ�,2016/10/13��Ԥ���˿������˿��Ƿ���֤ˢ��
         If mcurBill.bln���ѿ� = False And mcurBill.bln�˿��鿨 _
            Or mcurBill.bln���ѿ� And gbln���ѿ��˷��鿨 Then
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
            
            If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, mcurBill.lng�����ID, mcurBill.bln���ѿ�, _
                Trim(txtPatient.Text), mstrPatiSex, mstrPatiOld, dbl�˿���, mstrBrushCardNo, mstrbrPassWord, _
                True, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
            mcurBill.str���� = mstrBrushCardNo
        End If
    End If
     
goEnd:
    zlCheckDepositDelValied = True
    Exit Function
End Function

Private Function zlDepositDel(ByRef lngԤ��ID As Long, ByRef lng��Ԥ��ID As Long, ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ���Ԥ������
    '��Σ� lngԤ��ID-Ԥ��ID
    '���أ��ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dblCurMoney As Double, dblMoneySum As Double
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strXMLExpend As String, strԤ��IDs As String
    
    Err = 0: On Error GoTo Errhand:
    If mcurBill.lng�����ID = 0 Then zlDepositDel = True: Exit Function
    
    If mcurBill.bln���ѿ� Then
        '�������ѿ����
        strSQL = _
            "Select �ӿڱ��, ���ѿ�id, ����, -1 * Sum(Ӧ�ս��) As Ӧ�ս��" & vbNewLine & _
            "From ���˿������¼" & vbNewLine & _
            "Where ��¼���� = 4 And ����id = [1]" & vbNewLine & _
            "Group By �ӿڱ��, ���ѿ�id, ����"
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngԤ��ID)
        
        '����ʹ���˶������ѿ�
        dblMoneySum = dblMoney
        Do While Not rsTemp.EOF
             If Val(Nvl(rsTemp!Ӧ�ս��)) < dblMoneySum Then
              dblCurMoney = Val(Nvl(rsTemp!Ӧ�ս��))
              dblMoneySum = Round(dblMoneySum - Val(Nvl(rsTemp!Ӧ�ս��)), 6)
            Else
              dblCurMoney = dblMoneySum
              dblMoneySum = 0
            End If
            
            'Zl_���˿������¼_�˿�
            strSQL = "Zl_���˿������¼_�˿�("
            '  �ӿڱ��_In   ���ѿ����Ŀ¼.���%Type,
            strSQL = strSQL & "" & Val(Nvl(rsTemp!�ӿڱ��)) & ","
            '  ����_In       ���ѿ���Ϣ.����%Type,
            strSQL = strSQL & "'" & Nvl(rsTemp!����) & "',"
            '  ���ѿ�id_In   ���ѿ���Ϣ.Id%Type,
            strSQL = strSQL & "" & Val(Nvl(rsTemp!���ѿ�ID)) & ","
            '  ������_In   ���˿������¼.Ӧ�ս��%Type,
            strSQL = strSQL & "" & dblCurMoney & ","
            '  ԭԤ��id_In   ���˿������¼.����id%Type,
            strSQL = strSQL & "" & lngԤ��ID & ","
            '  ��Ԥ��id_In   ���˿������¼.����id%Type,
            strSQL = strSQL & "" & lng��Ԥ��ID & ","
            '  ����Ա���_In ���˿������¼.����Ա���%Type,
            strSQL = strSQL & "'" & UserInfo.��� & "',"
            '  ����Ա����_In ���˿������¼.����Ա����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  �˿�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type
            strSQL = strSQL & "To_Date('" & Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'))"

            zldatabase.ExecuteProcedure strSQL, Me.Caption
            
            If dblMoneySum = 0 Then Exit Do
            rsTemp.MoveNext
        Loop
        If dblMoneySum > 0 Then
            MsgBox "ʣ����˽��(" & Format(dblMoney - dblMoneySum, "0.00") & ")���㱾���˿���(" & _
                Format(dblMoney, "0.00") & ")�������˷ѣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If mbytInState = 3 And mcurBill.blnת�� Then
        strXMLExpend = "<IN><CZLX>4</CZLX></IN>"
        strSwapNO = mcurBill.str������ˮ��: strSwapMemo = mcurBill.str����˵��
        strSwapExtendInfor = "1|" & lng��Ԥ��ID
        If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModul, mcurBill.lng�����ID, mcurBill.str����, _
            lngԤ��ID, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then Exit Function
    Else
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
            '����: strSwapNo-������ˮ��(�˿����ˮ��)
            '         strSwapMemo-����˵��(�˿��˵��)
            '       strSwapExtendInfor-���׵���չ��Ϣ
            '           ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n ÿ����Ŀ�в��ܰ���|�ַ�
            '����:��������    True:���óɹ�,False:����ʧ��
         strSwapNO = mcurBill.str������ˮ��: strSwapMemo = mcurBill.str����˵��
         '81489,Ƚ����,2015-4-29,�˷Ѵ������ID
         strSwapExtendInfor = "1|" & lng��Ԥ��ID
         If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, mcurBill.lng�����ID, mcurBill.bln���ѿ�, mcurBill.str����, _
            "1|" & lngԤ��ID, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    End If
    '127450:���ϴ�,2018/6/20������˿�ʱ����Ҫ��ȡ��Ӧ�ĳ�Ԥ����¼
    If mbytInState = 3 Then
        strSQL = "Select ID From ����Ԥ����¼ Where ��¼���� In (1, 11) And NO = (Select NO from ����Ԥ����¼ Where ID = [1])"
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "����˿��¼", lng��Ԥ��ID)
        Do While Not rsTemp.EOF
            strԤ��IDs = strԤ��IDs & "," & rsTemp!ID
            rsTemp.MoveNext
        Loop
    End If
    If strԤ��IDs <> "" Then
        strԤ��IDs = Mid(strԤ��IDs, 2)
    Else
        strԤ��IDs = lng��Ԥ��ID
    End If
    If Save��������(strԤ��IDs, mcurBill.lng�����ID, mcurBill.bln���ѿ�, mcurBill.str����, strSwapNO, strSwapMemo, _
        strSwapExtendInfor, True, "1|" & lng��Ԥ��ID) = False Then Exit Function
    zlDepositDel = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 

Private Sub Load֧����ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:���˺�
    '����:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim str���� As String
    
                        
    '���㷽ʽ:���ò�ѯ��ҽ�ƿ�����ʱ��һ��ֻ֧��Ԥ����,�����ڴ��յ����
    'mbytCallObject:���õĶ���(0-Ԥ����������;1-���˷��ò�ѯ����;2-ҽ�ƿ�...
    If InStr(1, mstrPrivs, ";Ԥ���տ�;") > 0 Or _
        InStr(1, mstrPrivs, ";Ԥ���տ�;") > 0 Or _
        InStr(1, mstrPrivs, ";Ԥ�������˿�;") > 0 Or _
        InStr(1, mstrPrivs, ";����Ԥ��תסԺ;") > 0 _
        Or InStr(1, mstrPrivs, ";סԺԤ��ת����;") > 0 Or mbytCallObject > 0 Then
        str���� = ",1,2,7,8,3"
    End If
    'ֻ�д��տ�Ȩ��ʱ,���ܴ����������ʵ�Ԥ����
    '����:45471
    If InStr(1, mstrPrivs, ";���տ��˿�;") > 0 Or InStr(1, mstrPrivs, ";���տ���ȡ;") > 0 Then
        If mbytCallObject = 0 Then str���� = str���� & ",5"
    End If
    If str���� = "" Then str���� = ",1,2,7,8,3"
    str���� = Mid(str����, 2)
    
    If mblnNurseCall Then
        str���� = "7,8"
    End If
    
    On Error GoTo errHandle
    Set rsTemp = Get���㷽ʽ("Ԥ����", str����)
    Set mcolPayMode = New Collection
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType: varData = Split(strPayType, ";")
    With cboStyle
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!����) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If rsTemp!���� = 3 And InStr(mstrPrivs, ";����ת��;") > 0 Then
                mstr�����ʻ� = rsTemp!���� & "-" & rsTemp!���� '���ݲ��˶�̬����
            End If
            '104083:���ϴ���2016/12/21�������˻��������̬����
            '����Ϊ8�ĸ�������ҽ�ƿ�������
            If Not blnFind And InStr(",3,8,", "," & rsTemp!���� & ",") = 0 Then
                .AddItem Nvl(rsTemp!����)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!����))
                mcolPayMode.Add Array("", Nvl(rsTemp!����), 0, 0, 0, 0, Nvl(rsTemp!����), 0, 0), "K" & j
                If rsTemp!ȱʡ = 1 Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                If mstrȱʡ���㷽ʽ = Nvl(rsTemp!����) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        For i = 0 To UBound(varData)
            '�����:116175��������2017/12/8����ҽ�ƿ��Ľɿʽ���Ƶ���Ϊ�ܽ��㷽ʽ������豸���ù�ͬ����
            rsTemp.Filter = "���� ='" & Split(varData(i), "|")(6) & "'"
            If Not rsTemp.EOF Then
                If InStr(1, varData(i), "|") <> 0 And str���� <> 5 Then
                    varTemp = Split(varData(i), "|")
                    mcolPayMode.Add varTemp, "K" & j
                    .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                    If mstrȱʡ���㷽ʽ = varTemp(1) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                    j = j + 1
                End If
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    If cboStyle.ListCount = 0 Then
        MsgBox "Ԥ������û�п��õĽ��㷽ʽ,���ȵ����㷽ʽ���������á�", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Save��������(ByVal strԤ��IDs As String, _
    ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    ByVal str���� As String, str������ˮ�� As String, str����˵�� As String, _
    strExpend As String, Optional bln��Ԥ�� As Boolean = False, Optional strExpendOld As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSQL As String, varData As Variant, varTemp As Variant, cllPro As Collection, i As Long
     
    Err = 0: On Error GoTo Errhand:
    If bln��Ԥ�� = False Then
        '�˷�ʱ,�����Ľ���
        '���½�����Ϣ
        '    Zl_�����ӿڸ���_Update
        strSQL = "Zl_�����ӿڸ���_Update("
        '  �����id_In   ����Ԥ����¼.�����id%Type,
        strSQL = strSQL & "" & lng�����ID & ","
        '  ���ѿ�_In     Number,
        strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
        '  ����_In       ����Ԥ����¼.����%Type,
        strSQL = strSQL & "'" & str���� & "',"
        '  ����ids_In    Varchar2,
        strSQL = strSQL & "'" & strԤ��IDs & "',"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
        strSQL = strSQL & "'" & str������ˮ�� & "',"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type
        strSQL = strSQL & "'" & str����˵�� & "',"
        'Ԥ����ɿ�_In Number := 0
        strSQL = strSQL & "" & 1 & ","
        '�˷ѱ�־ :1-�˷�;0-����
        strSQL = strSQL & "" & IIf(bln��Ԥ��, 1, 0) & ")"
        Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    gcnOracle.CommitTrans
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
                        strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
                        '����_In     ����Ԥ����¼.����%Type,
                        strSQL = strSQL & "'" & str���� & "',"
                        '����ids_In  Varchar2,
                        strSQL = strSQL & "'" & strԤ��IDs & "',"
                        '������Ϣ_In Varchar2:������Ŀ|��������||...
                        strSQL = strSQL & "'" & str������Ϣ & "',"
                        'Ԥ����ɿ�_In Number := 0
                        strSQL = strSQL & "1)"
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
            strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
            '����_In     ����Ԥ����¼.����%Type,
            strSQL = strSQL & "'" & str���� & "',"
            '����ids_In  Varchar2,
            strSQL = strSQL & "'" & strԤ��IDs & "',"
            '������Ϣ_In Varchar2:������Ŀ|��������||...
            strSQL = strSQL & "'" & str������Ϣ & "',"
            'Ԥ����ɿ�_In Number := 0
            strSQL = strSQL & "1)"
            zlAddArray cllPro, strSQL
        End If
    End If
    Err = 0: On Error GoTo ErrOthers:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    Save�������� = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrOthers:
    '    �ܱ������,������
     Call ErrCenter
End Function


Private Function zlInterfacePrayMoney(ByVal lngԤ��ID As Long, ByVal strNo As String, ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ӿ�֧�����
    '����:֧���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    
    If cboStyle.ItemData(cboStyle.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    If mlngCardTypeID = 0 Then zlInterfacePrayMoney = True: Exit Function
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
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, mlngCardTypeID, mbln���ѿ�, mstrBrushCardNo, "", strNo, _
        dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    If Save��������(lngԤ��ID, mlngCardTypeID, mbln���ѿ�, mstrBrushCardNo, strSwapGlideNO, strSwapMemo, _
        strSwapExtendInfor) = False Then Exit Function
    zlInterfacePrayMoney = True
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
    If mbytInState = 1 Or mbytInState = 2 Then Exit Sub
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
Private Sub LoadPatiPage(ByVal lng����ID As Long)
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
    strSQL = "" & _
    "   Select ��ҳID,��������,��Ժ����,��Ժ����  " & _
    "   From ������ҳ" & _
    "   Where ����ID=[1]  " & _
    "   Order By Nvl(��ҳID,0) Desc"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    With cboPatiPage
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
                If Val(Nvl(rsTemp!��ҳID)) = Val(Nvl(mrsInfo!��ҳID)) Then
                    .ListIndex = .NewIndex
                End If
            End If
            rsTemp.MoveNext
        Loop
        If bln���� = True Then Call cbo.SetListWidth(cboPatiPage.hWnd, 2000)
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
    Dim strSQL As String
    Dim lng����ID As Long, lng��ҳID As Long
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If mblnδ��Ʋ���Ԥ�� = False Then Checkδ��Ʋ���Ԥ�� = True: Exit Function
    '����Ԥ�������
    If cboType.ItemData(cboType.ListIndex) <> 2 Then Checkδ��Ʋ���Ԥ�� = True: Exit Function
    '��ǰסԺ������Ϊ��Ժ��,Ҳ�����
    If Val(Nvl(mrsInfo!��Ժ)) <> 1 Then Checkδ��Ʋ���Ԥ�� = True: Exit Function
    lng����ID = Val(Nvl(mrsInfo!����ID))
    '������סԺ������,Ҳ�ܽ�Ԥ��,��˲����
    If cboPatiPage.ListIndex < 0 Then Checkδ��Ʋ���Ԥ�� = True: Exit Function
    lng��ҳID = cboPatiPage.ItemData(cboPatiPage.ListIndex)
    strSQL = "Select ID From ���˱䶯��¼ Where ����id = [1] And ��ҳid = [2] And ��ʼԭ�� = 2 And ���� Is Not Null "
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "ע��" & vbCrLf & "   ���ˡ�" & mrsInfo!���� & "��δ���,�������Ԥ����!", vbInformation + vbOKOnly, gstrSysName
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
    Dim lng����ID As Long
    Dim dblԤ����� As Double, dbl������� As Double, dblʣ����� As Double
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then lng����ID = 0
    If mrsInfo.State <> 1 Then lng����ID = 0
    lng����ID = Val(Nvl(mrsInfo!����ID))
    If lng����ID <> 0 Then
        Set mrsDepositBalance = GetMoneyInfo(lng����ID, , , , True)
        If Not mrsDepositBalance Is Nothing Then
            With mrsDepositBalance
                .Filter = "����=" & cboType.ItemData(cboType.ListIndex)
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not .EOF
                    dbl������� = dbl������� + Val(Nvl(!�������))
                    dblԤ����� = dblԤ����� + Val(Nvl(!Ԥ�����))
                    .MoveNext
                Loop
            End With
        End If
        dblʣ����� = dblԤ����� - dbl�������
        If mdblʣ���� <> dblʣ����� Then
            MsgBox "���˵�ʣ������ѷ����仯,������ȷ���˿���!", vbInformation + vbOKOnly, gstrSysName
            Call ShowPremayBalance(False, 0)
            txtMoney.SetFocus: Exit Function
        End If
    End If
    Check�˿� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
