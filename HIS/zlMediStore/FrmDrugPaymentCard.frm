VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form FrmDrugPaymentCard 
   Caption         =   "ҩƷ���"
   ClientHeight    =   6975
   ClientLeft      =   600
   ClientTop       =   2550
   ClientWidth     =   11400
   Icon            =   "FrmDrugPaymentCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FrmDrugPaymentCard.frx":0E42
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
      Height          =   3495
      Left            =   780
      TabIndex        =   46
      Top             =   1800
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6165
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Fra2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   6720
      TabIndex        =   31
      Top             =   -90
      Width           =   4935
      Begin VB.PictureBox picdown 
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   240
         ScaleHeight     =   1155
         ScaleWidth      =   4455
         TabIndex        =   45
         Top             =   4290
         Width           =   4455
         Begin VB.TextBox Txt����˵�� 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   870
            MaxLength       =   50
            TabIndex        =   20
            Top             =   0
            Width           =   3585
         End
         Begin VB.TextBox Txt�������� 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   420
            Width           =   1875
         End
         Begin VB.TextBox Txt������� 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   810
            Width           =   1875
         End
         Begin VB.TextBox Txt����� 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   630
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   810
            Width           =   1005
         End
         Begin VB.TextBox Txt������ 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   630
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   420
            Width           =   1005
         End
         Begin VB.Label Lbl����˵�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����˵��:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   19
            Top             =   60
            Width           =   810
         End
         Begin VB.Label Lbl������� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������"
            Height          =   180
            Left            =   1770
            TabIndex        =   27
            Top             =   885
            Width           =   720
         End
         Begin VB.Label Lbl����� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����"
            Height          =   180
            Left            =   30
            TabIndex        =   25
            Top             =   870
            Width           =   540
         End
         Begin VB.Label Lbl�������� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Left            =   1800
            TabIndex        =   23
            Top             =   480
            Width           =   720
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   180
            Left            =   30
            TabIndex        =   21
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.PictureBox picup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   240
         ScaleHeight     =   1725
         ScaleWidth      =   4455
         TabIndex        =   37
         Top             =   720
         Width           =   4455
         Begin VB.TextBox TxtNo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   2925
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   38
            Top             =   0
            Width           =   1275
         End
         Begin VB.Label lbl�����ʺ� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�����ʺ�:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   54
            Top             =   1260
            Width           =   810
         End
         Begin VB.Label txt�����ʺ� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   960
            TabIndex        =   53
            Top             =   1260
            Width           =   90
         End
         Begin VB.Label txt������ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   3840
            TabIndex        =   51
            Top             =   1035
            Width           =   90
         End
         Begin VB.Label txt˰��� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   960
            TabIndex        =   50
            Top             =   1530
            Width           =   90
         End
         Begin VB.Label txt������ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   960
            TabIndex        =   49
            Top             =   990
            Width           =   90
         End
         Begin VB.Label txt�绰��ַ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   960
            TabIndex        =   48
            Top             =   720
            Width           =   90
         End
         Begin VB.Label txt��λ���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   960
            TabIndex        =   47
            Top             =   450
            Width           =   90
         End
         Begin VB.Label LblNo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   2595
            TabIndex        =   44
            Top             =   45
            Width           =   240
         End
         Begin VB.Label Lbl��λ���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��λ����:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   43
            Top             =   450
            Width           =   810
         End
         Begin VB.Label Lbl�绰��ַ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ַ�绰:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   42
            Top             =   720
            Width           =   810
         End
         Begin VB.Label Lbl������ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   41
            Top             =   990
            Width           =   810
         End
         Begin VB.Label Lbl������ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   3180
            TabIndex        =   40
            Top             =   1035
            Width           =   630
         End
         Begin VB.Label lbl˰��� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "˰��ǼǺ�:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   39
            Top             =   1530
            Width           =   990
         End
      End
      Begin ZL9BillEdit.BillEdit mshPaymentList 
         Height          =   1665
         Left            =   240
         TabIndex        =   18
         Top             =   2535
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2937
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ����֪ͨ��"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1155
         TabIndex        =   35
         Top             =   330
         Width           =   2100
      End
   End
   Begin VB.Frame Fra1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   0
      TabIndex        =   32
      Top             =   -90
      Width           =   6735
      Begin MSComctlLib.TreeView tvwProvider 
         Height          =   3585
         Left            =   1320
         TabIndex        =   36
         Top             =   2160
         Visible         =   0   'False
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   6324
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgTree"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.TextBox Txt��ҩ��λ 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   930
         TabIndex        =   1
         Top             =   210
         Width           =   4275
      End
      Begin VB.CommandButton Cmd��Ӧ�� 
         Caption         =   "��"
         Height          =   300
         Left            =   5220
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   210
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPurchaseList 
         DragIcon        =   "FrmDrugPaymentCard.frx":1184
         Height          =   2745
         Left            =   30
         TabIndex        =   12
         Top             =   1425
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   4842
         _Version        =   393216
         BackColor       =   -2147483624
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ImageList imgTree 
         Left            =   6120
         Top             =   90
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDrugPaymentCard.frx":12CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDrugPaymentCard.frx":2FDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDrugPaymentCard.frx":4CE4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ZL9BillEdit.BillEdit mshImprest 
         Height          =   885
         Left            =   120
         TabIndex        =   14
         Top             =   4560
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1561
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.TextBox Txt��Ʊ�� 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   930
         TabIndex        =   4
         ToolTipText     =   "�����ʽ:��ʼ��Ʊ��-������Ʊ��"
         Top             =   600
         Width           =   1995
      End
      Begin VB.TextBox Txt���� 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         ToolTipText     =   "�����ʽ:��ʼNO-����NO"
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NO"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3240
         TabIndex        =   5
         Top             =   660
         Width           =   180
      End
      Begin VB.Label Lbl��Ʊ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   3
         Top             =   667
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ���Ʊ�嵥"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   4320
         Width           =   1260
      End
      Begin VB.Label Lbl������ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3720
         TabIndex        =   10
         Top             =   1110
         Width           =   810
      End
      Begin VB.Label Txt������ 
         Height          =   180
         Left            =   4650
         TabIndex        =   11
         Top             =   1110
         Width           =   2010
      End
      Begin VB.Label Lbl��Ӧ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ��λ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Lbl�嵥 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "δ���Ʊ�嵥"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   1110
         Width           =   1260
      End
      Begin VB.Label Lbl�ۼƺϼ� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ۼ�Ӧ��:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1740
         TabIndex        =   8
         Top             =   1110
         Width           =   810
      End
      Begin VB.Label Lbl�ϼ� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2730
         TabIndex        =   9
         Top             =   1110
         Width           =   1065
      End
   End
   Begin VB.Frame Fra4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   34
      Top             =   5370
      Width           =   6735
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   240
         Picture         =   "FrmDrugPaymentCard.frx":5B36
         TabIndex        =   52
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton Cmd��ѡ�񸶿� 
         Caption         =   "��ѡ�񸶿(&U)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   5040
         Picture         =   "FrmDrugPaymentCard.frx":5C80
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Cmdȫѡ 
         Caption         =   "ȫѡ(&A)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   1800
         Picture         =   "FrmDrugPaymentCard.frx":5DCA
         TabIndex        =   15
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton Cmd��� 
         Caption         =   "ȫ��(&C)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   3210
         Picture         =   "FrmDrugPaymentCard.frx":5F14
         TabIndex        =   16
         Top             =   240
         Width           =   1100
      End
   End
   Begin VB.Frame Fra3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6720
      TabIndex        =   33
      Top             =   5370
      Width           =   4935
      Begin VB.CommandButton Cmdȡ������ 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   2700
         Picture         =   "FrmDrugPaymentCard.frx":605E
         TabIndex        =   30
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton Cmd���� 
         Caption         =   "ȷ��(&O)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   900
         Picture         =   "FrmDrugPaymentCard.frx":61A8
         TabIndex        =   29
         Top             =   240
         Width           =   1100
      End
   End
   Begin VB.Timer LimitTime 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   6660
      Top             =   0
   End
   Begin MSComctlLib.ImageList imlTbrClr 
      Left            =   705
      Top             =   -1035
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":62F2
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":650E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":672A
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":6946
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":6B62
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":6D7E
            Key             =   "Annul"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":6F9A
            Key             =   "Store"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":71B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":73D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":76EE
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":790A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":7B26
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmDrugPaymentCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CurMoney As Currency            '��ǰ�����ܽ��
Private CurLastMoney As Currency        '�ϴ�ѡ�񸶿���
Private mintUnit As Integer             '0:ҩ�ⵥλ 1:���ﵥλ 2:סԺ��λ 3:�ۼ۵�λ

Private mblnSave As Boolean
Private mblnSuccess As Boolean
Private mstr���ݺ� As String
Private mint�༭״̬ As Integer         '�༭���� 1:��ʾ����;2:��ʾ�޸�,3:��ʾ���;4:��ʾȡ��
Private mint��¼״̬ As Integer
Private mblnChange As Boolean
Private mintParallelRecord As Integer   '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mintpurchaseclick As Boolean
Dim mstrPrivs As String                     'Ȩ��

Private Const mconintCol��־ As Integer = 0
Private Const mconintcol��Ʊ��  As Integer = 1
Private Const mconintcol��ⵥ�� As Integer = 2
Private Const mconintcolҩƷ��Ϣ As Integer = 3
Private Const mconIntCol��� As Integer = 4
Private Const mconIntCol��λ As Integer = 5
Private Const mconintcol��Ʊ��� As Integer = 6
Private Const mconIntCol���� As Integer = 7
Private Const mconIntCol�ɹ��� As Integer = 8
Private Const mconintcol������ As Integer = 9
Private Const mconintcol������� As Integer = 10
Private Const mconIntCol�ۼ� As Integer = 11
Private Const mconIntCol�ۼ۽�� As Integer = 12
Private Const mconIntCol���� As Integer = 13
Private Const mconIntCol���� As Integer = 14
Private Const mconintcol������� As Integer = 15

Private Const mconIntColS As Integer = 16

Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim rs���㷽ʽ As New Recordset
    Dim intLop As Integer
    
    On Error GoTo errHandle
    GetDepend = False
    With rsDepend
        If .State = 1 Then .Close
        gstrSQL = "Select ID,�ϼ�ID,����,����,����,ĩ��,��ַ||�绰 as �绰��ַ,��������,�ʺ�,˰��ǼǺ� From ҩƷ��Ӧ�� Where " & _
              " To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' Start with �ϼ�ID is Null Connect by prior ID=�ϼ�ID"
        Call SQLTest(App.Title, "ҩƷ���", gstrSQL)
        Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "GetDepend")
        Call SQLTest
        
        If .EOF Then
            MsgBox "ҩƷ��Ӧ�̵���Ϣ��ȫ�����ڹ�ҩ��λ�����н������ã�", vbInformation, gstrSysName
            Exit Function
        End If
        
    End With
        
    
    With rs���㷽ʽ
        If .State = 1 Then .Close
        gstrSQL = "Select * From ���㷽ʽӦ�� Where Ӧ�ó���='��ҩ��' Order by ȱʡ��־ desc"
        Call SQLTest(App.Title, "ҩƷ���", gstrSQL)
        Set rs���㷽ʽ = zldatabase.OpenSQLRecord(gstrSQL, "GetDepend")
        Call SQLTest
        
        If .EOF Then
            MsgBox "���㷽ʽӦ����Ϣ��ȫ,���ڽ��㷽ʽ�����н������ã�", vbInformation, gstrSysName
            Exit Function
        End If
        mshPaymentList.Clear
        For intLop = 1 To .RecordCount
            mshPaymentList.AddItem !���㷽ʽ
            .MoveNext
        Next
        mshPaymentList.ListIndex = 0
        
        .Close
    End With
    
    With rsDepend
        tvwProvider.Nodes.Clear
        tvwProvider.Nodes.Add , , "R", "���й�Ӧ��", 1, 1
        tvwProvider.Nodes("R").Tag = 0
        .MoveFirst
        
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                If !ĩ�� = 1 Then
                    tvwProvider.Nodes.Add "R", 4, "K_" & !Id, "[" & !���� & "]" & !����, 3, 3
                Else
                    tvwProvider.Nodes.Add "R", 4, "K_" & !Id, "[" & !���� & "]" & !����, 2, 2
                End If
            Else
                If !ĩ�� = 1 Then
                    tvwProvider.Nodes.Add "K_" & !�ϼ�ID, 4, "K_" & !Id, "[" & !���� & "]" & !����, 3, 3
                Else
                    tvwProvider.Nodes.Add "K_" & !�ϼ�ID, 4, "K_" & !Id, "[" & !���� & "]" & !����, 2, 2
                End If
            End If
            tvwProvider.Nodes("K_" & !Id).Tag = !ĩ��
            .MoveNext
        Loop
        tvwProvider.Nodes("R").Selected = True
        tvwProvider.Nodes("R").Expanded = True
        
    End With
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, _
        Optional int��¼״̬ As Integer = 1, Optional BlnSuccess As Boolean = False)
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mblnSuccess = BlnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1320)
    
    mintUnit = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ��������", "ҩƷ��λ", "0")
    If Not GetDepend Then Exit Sub
    
    If mint�༭״̬ = 1 Then
        mstr���ݺ� = NextNo(31)
        TxtNo = mstr���ݺ�
        
    ElseIf mint�༭״̬ = 2 Then
'        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then
        'mblnEdit = False
        Cmd����.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 4 Then
        'mblnEdit = False
        Cmd����.Caption = "��ӡ(&P)"
        If InStr(mstrPrivs, "����֪ͨ����ӡ") = 0 Then
            Cmd����.Visible = False
        Else
            Cmd����.Visible = True
        End If
    End If
    
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub


Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Cmd��ѡ�񸶿�_Click()
    Dim Array��Ʊ��() As String
    Dim Str��Ʊ�� As String
    Dim curPayment As Currency
    Dim intRow As Integer
    Dim strTemp As String
    Dim rs���㷽ʽ As New Recordset
    
    Str��Ʊ�� = ""
    
'    With mshImprest
'        curImprest = 0
'        For introw = 1 To .Rows - 1
'            If .TextMatrix(introw, 0) = "��" Then
'                curImprest = curImprest + .TextMatrix(introw, 2)
'            End If
'        Next
'    End With
    
    curPayment = CurLastMoney
    
    With mshPaymentList
        .ClearMsf
        .Cols = 3
        .rows = 3

        .TextMatrix(0, 0) = "���ʽ"
        .TextMatrix(0, 1) = "������"
        .TextMatrix(0, 2) = "�������"
    
        With rs���㷽ʽ
            gstrSQL = "Select * From ���㷽ʽӦ�� Where Ӧ�ó���='��ҩ��' Order by ȱʡ��־ desc"
            Call OpenRecordset(rs���㷽ʽ, "���㷽ʽӦ��")
            
            If .EOF Then
                MsgBox "���㷽ʽӦ����Ϣ��ȫ��", vbInformation, gstrSysName
                Exit Sub
            End If
            mshPaymentList.Clear
            For intRow = 1 To .RecordCount
                mshPaymentList.AddItem !���㷽ʽ
                .MoveNext
            Next
            mshPaymentList.ListIndex = 0
            .Close
        End With
            
        .TextMatrix(1, 0) = mshPaymentList.CboText
        .TextMatrix(1, 1) = GetFormat(curPayment, 2)
    End With
    
    Cmd��ѡ�񸶿�.Enabled = False
    Cmd����.Enabled = True
    mshPaymentList.Active = True
    
    'ͳ����ⵥID
    With mshPurchaseList
        For intRow = 1 To .rows - 2
            If .TextMatrix(intRow, mconintCol��־) <> "" Then
                strTemp = "'" & String(8 - Len(.TextMatrix(intRow, GetCol(mshPurchaseList, "��Ʊ��"))), "0") & .TextMatrix(intRow, GetCol(mshPurchaseList, "��Ʊ��")) & "'"
                
                If Str��Ʊ�� = "" Then
                    Str��Ʊ�� = strTemp
                Else
                    If InStr(1, Str��Ʊ��, strTemp) = 0 Then
                        Str��Ʊ�� = Str��Ʊ�� & "," & strTemp
                    End If
                End If
            End If
        Next
    End With
    Array��Ʊ�� = Split(Str��Ʊ��, ",")
    Me.txt������ = UBound(Array��Ʊ��) + 1
End Sub

Private Sub Cmd��Ӧ��_Click()
    tvwProvider.Visible = tvwProvider.Visible Xor True
    If tvwProvider.Visible Then
        tvwProvider.Top = Txt��ҩ��λ.Top + Txt��ҩ��λ.Height
        tvwProvider.SetFocus
    End If
    Cmd����.Enabled = False
End Sub

Private Function SaveVerify() As Boolean
    Dim intRow As Integer
    Dim NO_IN As String
    Dim ������_IN As Double
    Dim ��λID_IN As Long
    Dim �����_IN As String
    
    SaveVerify = False
    
    NO_IN = TxtNo
    ��λID_IN = Txt��ҩ��λ.Tag
    �����_IN = UserInfo.�û�����
    ������_IN = 0
    On Error GoTo errHandle:
    
    With mshPaymentList
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" And IIf(.TextMatrix(intRow, 1) = "", 0, .TextMatrix(intRow, 1)) <> 0 Then
                ������_IN = ������_IN + Val(.TextMatrix(intRow, 1))
            End If
        Next
    End With
    'zl_ҩƷ�������_VERIFY( /*NO_IN*/, /*��λID_IN*/, /*������_IN*/, /*�����_IN*/ );
    gstrSQL = "zl_ҩƷ�������_VERIFY('" & NO_IN & "'," & ��λID_IN & "," & ������_IN _
        & ",'" & �����_IN & "')"
    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
    
    
    SaveVerify = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function


Private Function SaveCard() As Boolean
    Dim IntTotalRows As Integer
    Dim intRow As Integer
    Dim intLop As Integer
    Dim Cur��� As Currency
    Dim curImprest As Currency
    Dim NO_IN As String
    Dim ���_IN As Integer
    Dim Ԥ����_IN As Integer
    Dim ��λID_IN As Long
    Dim ���_IN As Double
    Dim ���㷽ʽ_IN As String
    Dim �������_IN As String
    Dim ������_IN As String
    Dim ��������_IN As String
    Dim �������_IN As Long
    Dim ժҪ_IN As String
    
    SaveCard = False
    With mshPaymentList
        For intRow = 1 To .rows - 1
            Cur��� = Cur��� + Val(.TextMatrix(intRow, 1))
        Next
    End With
    
    
    If Cur��� <> CurLastMoney Then
        MsgBox "�����ƽ,���鸶��������ⵥ��Ʊ����Ԥ����֮���Ƿ���ͬ!", vbInformation, gstrSysName
        mshPaymentList.SetFocus
        Exit Function
    End If
    
    IntTotalRows = IIf(LTrim(RTrim(mshPaymentList.TextMatrix(1, 1))) = "", 0, 1)
    If IntTotalRows < 1 Then Exit Function
    IntTotalRows = IIf(LTrim(RTrim(mshPaymentList.TextMatrix(mshPaymentList.rows - 1, 1))) = "", mshPaymentList.rows - 2, mshPaymentList.rows - 1)
    If IntTotalRows < 1 Then Exit Function
    If CheckData(IntTotalRows) = False Then Exit Function
    
    NO_IN = TxtNo
    Ԥ����_IN = 0
    ��λID_IN = Txt��ҩ��λ.Tag
    ������_IN = UserInfo.�û�����
    ��������_IN = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    ժҪ_IN = Txt����˵��
    �������_IN = zldatabase.GetNextId("ҩƷ�����¼")
    
    On Error GoTo errHandle:
    
    '��ʼ����
    gcnOracle.BeginTrans
    
    If mint�༭״̬ = 2 Then
        gstrSQL = "zl_ҩƷ�������_delete('" & NO_IN & "')"
            
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        Call SQLTest
    End If
        
    'ѭ������ÿ������
    With mshPaymentList
        'zl_ҩƷ�������_INSERT( /*NO_IN*/, /*���_IN*/, /*Ԥ����_IN*/, /*��λID_IN*/,
            '/*���_IN*/, /*���㷽ʽ_IN*/, /*�������_IN*/, /*������_IN*/, /*��������_IN*/,
            '/*�������_IN*/, /*ժҪ_IN*/ );
        For intRow = 1 To IntTotalRows
            'Modified by zyb 2002-11-08
            'If Val(.TextMatrix(intRow, 1)) > 0 Then
            If Val(.TextMatrix(intRow, 1)) <> 0 Then
                ���_IN = intRow
                ���_IN = .TextMatrix(intRow, 1)
                ���㷽ʽ_IN = .TextMatrix(intRow, 0)
                �������_IN = .TextMatrix(intRow, 2)
                gstrSQL = "zl_ҩƷ�������_INSERT('" & NO_IN & "'," & ���_IN & "," & Ԥ����_IN & "," & ��λID_IN _
                    & "," & ���_IN & ",'" & ���㷽ʽ_IN & "','" & �������_IN & "','" & ������_IN & "',to_date('" _
                    & ��������_IN & "','yyyy-mm-dd HH24:MI:SS')," & �������_IN & ",'" & ժҪ_IN & "')"
                Call SQLTest(App.Title, Me.Caption, gstrSQL)
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                Call SQLTest
            End If
        Next
    End With
                        
    '��Ӧ�ɹ��嵥
    With mshPurchaseList
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, mconintCol��־) <> "" Then
                gstrSQL = "Update ҩƷӦ����¼ Set �������=" & �������_IN & " where �շ�id=" & .RowData(intRow)
                Call SQLTest(App.Title, Me.Caption, gstrSQL)
                gcnOracle.Execute gstrSQL
                Call SQLTest
                
            End If
        Next
    End With
    '����Ԥ����
    With mshImprest
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                gstrSQL = "Update ҩƷ�����¼ Set �������=" & �������_IN & " where id=" & .RowData(intRow)
                Call SQLTest(App.Title, Me.Caption, gstrSQL)
                gcnOracle.Execute gstrSQL
                Call SQLTest
                
            End If
        Next
    End With
    '�ύ����
    gcnOracle.CommitTrans
    SaveCard = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Cmd����_Click()
    Dim BlnSuccess As Boolean
    
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_1", "zl8_bill_1320_1"), mint��¼״̬, 0, 1320, "ҩƷ���", TxtNo.Text
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 3 Then        '���
        If SaveVerify = True Then
            mblnChange = False
            mblnSave = False
            mblnSuccess = True

            If GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ��������", "��˴�ӡ", "0") = "1" Then
                '��ӡ
                If InStr(mstrPrivs, "����֪ͨ����ӡ") <> 0 Then
                    ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_1", "zl8_bill_1320_1"), Me, "���ݱ��=" & TxtNo.Text, "��¼״̬=" & mint��¼״̬, 2
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
            
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
        mblnChange = False
        mblnSave = False
        mblnSuccess = True
        If GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ��������", "���̴�ӡ", "0") = "1" Then
            '��ӡ
            If InStr(mstrPrivs, "����֪ͨ����ӡ") <> 0 Then
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_1", "zl8_bill_1320_1"), Me, "���ݱ��=" & TxtNo.Text, "��¼״̬=" & mint��¼״̬, 2
            End If
        End If
        If mint�༭״̬ = 2 Then   '�޸�
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    mstr���ݺ� = NextNo(31)
    TxtNo = mstr���ݺ�
    mblnSave = False
'    mblnEdit = True
    
    initGrid
    Txt������ = ""
    Lbl�ϼ� = ""
    Txt��ҩ��λ = ""
    Txt��ҩ��λ.Tag = 0
    tvwProvider.Tag = 0
    txt��λ���� = ""
    txt�绰��ַ = ""
    
    txt������ = ""
    txt˰��� = ""
    txt�����ʺ� = ""
    Txt����˵�� = ""
    Txt��Ʊ�� = ""
    Txt���� = ""
    Cmd����.Enabled = False
    
End Sub

Private Sub Cmd���_Click()
    Dim IntChk As Integer
    For IntChk = 1 To mshPurchaseList.rows - 2
        mshPurchaseList.TextMatrix(IntChk, 0) = ""
    Next
    For IntChk = 1 To mshImprest.rows - 2
        If mshImprest.TextMatrix(IntChk, 0) <> "" Then
            mshImprest.TextMatrix(IntChk, 0) = ""
        End If
    Next
    
    BanlanceMoney
End Sub

Private Sub Cmdȡ������_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub Cmdȫѡ_Click()
    Dim IntChk As Integer
    Cmd���_Click
    For IntChk = 1 To mshPurchaseList.rows - 2
        mshPurchaseList.Row = IntChk
        mshPurchaseList_KeyDown vbKeySpace, 0
    Next
End Sub


Private Sub Form_Activate()
    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            '�����ѱ�ɾ��
            MsgBox "�õ����ѱ�ɾ�������飡", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 3
            '�޸ĵĵ����ѱ����
            MsgBox "�õ����ѱ���������ˣ����飡", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then zlcommfun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub initGrid()
    Dim IntCol As Integer
    
    With mshPurchaseList
        .Clear
        .rows = 2
        .Cols = mconIntColS
        .TextMatrix(0, mconintCol��־) = "��־"
        .TextMatrix(0, mconintcol��Ʊ��) = "��Ʊ��"
        .TextMatrix(0, mconintcol��ⵥ��) = "��ⵥ��"
        .TextMatrix(0, mconintcolҩƷ��Ϣ) = "ҩƷ��Ϣ"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconintcol��Ʊ���) = "��Ʊ���"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol�ɹ���) = "�ɹ���"
        .TextMatrix(0, mconintcol������) = "������"
        .TextMatrix(0, mconintcol�������) = "�������"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconintcol�������) = "�������"
        
        .ColAlignment(mconintcol��Ʊ��) = flexAlignLeftCenter
        .ColAlignment(mconintcol��ⵥ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconintcol��Ʊ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ���) = flexAlignRightCenter
        .ColAlignment(mconintcol������) = flexAlignRightCenter
        .ColAlignment(mconintcol�������) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
        
        '��ͷ�м����
        For IntCol = 0 To .Cols - 1
            .ColAlignmentFixed(IntCol) = flexAlignCenterCenter
        Next
        
        .ColWidth(mconintCol��־) = 450
        .ColWidth(mconintcol��Ʊ��) = 800
        .ColWidth(mconintcol��ⵥ��) = 800
        .ColWidth(mconintcolҩƷ��Ϣ) = 2000
        .ColWidth(mconIntCol���) = 800
        .ColWidth(mconIntCol��λ) = 450
        .ColWidth(mconintcol��Ʊ���) = 1000
        .ColWidth(mconIntCol����) = 1000
        .ColWidth(mconIntCol�ɹ���) = 1000
        .ColWidth(mconintcol������) = 1000
        .ColWidth(mconintcol�������) = 1000
        .ColWidth(mconIntCol�ۼ�) = 1000
        .ColWidth(mconIntCol�ۼ۽��) = 1000
        .ColWidth(mconIntCol����) = 2000
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconintcol�������) = 1000
    End With
    
    With mshPaymentList
        .ClearMsf
        .Cols = 3
        .rows = 3
        .TextMatrix(0, 0) = "���ʽ"
        .TextMatrix(0, 1) = "������"
        .TextMatrix(0, 2) = "�������"
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 1600
        
        .ColData(0) = 3
        .ColData(1) = 4
        .ColData(2) = 4
    End With
    
    With mshImprest
        .ClearMsf
        
        .Cols = 4
        .rows = 4
        .Active = True
        
        .TextMatrix(0, 0) = "ѡ��"
        .TextMatrix(0, 1) = "���㷽ʽ"
        .TextMatrix(0, 2) = "������"
        .TextMatrix(0, 3) = "�������"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 800
        .ColWidth(2) = 1200
        .ColWidth(3) = 1600
        
        .ColData(0) = -1
        .ColData(1) = 5
        .ColData(2) = 5
        .ColData(3) = 5
        .LocateCol = 0
    End With
End Sub

Private Sub initCard()
    initGrid
    On Error GoTo errHandle
    If mint�༭״̬ = 1 Then
        Txt��Ʊ��.Enabled = True
        Txt����.Enabled = True
        Txt������ = UserInfo.�û�����
        Txt�������� = Format(zldatabase.Currentdate, "yyyy-MM-dd")
        Exit Sub
    Else
        Dim rsPayment As New Recordset
        Dim intRecord As Integer
        Dim intLop As Integer
        
        gstrSQL = "SELECT a.���, a.���, a.���㷽ʽ, a.�������, a.ժҪ,a.�������,a.������,a.��������,a.�����,a.�������, " _
                & " b.����, b.id,b.��ַ || b.�绰 as �绰��ַ,��������,�ʺ�,˰��ǼǺ� " _
                & " FROM ҩƷ�����¼ a, ҩƷ��Ӧ�� b " _
                & "Where a.��λid = b.ID " _
                & "  AND no = '" & mstr���ݺ� _
                & "' AND ��¼״̬ = " & mint��¼״̬ _
               & " order by a.��� "
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        Set rsPayment = zldatabase.OpenSQLRecord(gstrSQL, "initCard")
        Call SQLTest
        
        If Not rsPayment.EOF Then
            intRecord = rsPayment.RecordCount
            rsPayment.MoveFirst
            Txt��ҩ��λ.Text = rsPayment!����
            Txt��ҩ��λ.Tag = rsPayment!Id
            Txt����˵�� = IIf(IsNull(rsPayment!ժҪ), "", rsPayment!ժҪ)
            txt������ = Get������(IIf(IsNull(rsPayment!�������), 0, rsPayment!�������))
            Txt����˵��.Tag = IIf(IsNull(rsPayment!�������), 0, rsPayment!�������)
            Txt������ = rsPayment!������
            If mint�༭״̬ = 2 Then
                Txt������ = UserInfo.�û�����
            End If
            Txt�������� = Format(rsPayment!��������, "yyyy-mm-dd hh:mm:ss")
            Txt����� = IIf(IsNull(rsPayment!�����), "", rsPayment!�����)
            Txt������� = IIf(IsNull(rsPayment!�������), "", Format(rsPayment!�������, "yyyy-mm-dd hh:mm:ss"))
                        
            tvwProvider.Tag = "1"
            txt��λ����.Caption = rsPayment!����
            txt�绰��ַ = IIf(IsNull(rsPayment!�绰��ַ), "", rsPayment!�绰��ַ)
            
            txt������ = IIf(IsNull(rsPayment!��������), "", rsPayment!��������)
            txt˰��� = IIf(IsNull(rsPayment!˰��ǼǺ�), "", rsPayment!˰��ǼǺ�)
            txt�����ʺ� = IIf(IsNull(rsPayment!�ʺ�), "", rsPayment!�ʺ�)
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            RefreshPurchaseList
            RefreshImprest
            BanlanceMoney
            
            With mshPaymentList
                For intLop = 1 To intRecord
                    .TextMatrix(intLop, 0) = IIf(IsNull(rsPayment!���㷽ʽ), "", rsPayment!���㷽ʽ)
                    .TextMatrix(intLop, 1) = rsPayment!���
                    .TextMatrix(intLop, 2) = IIf(IsNull(rsPayment!�������), "", rsPayment!�������)
                    If intLop = .rows - 1 Then .rows = .rows + 1
                    rsPayment.MoveNext
                Next
            End With
            
            
            Cmd��ѡ�񸶿�.Enabled = False
            mshPaymentList.Active = False
            If mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
                Txt��ҩ��λ.Enabled = False
                Cmd��Ӧ��.Enabled = False
                Cmdȫѡ.Enabled = False
                Cmd���.Enabled = False
                Cmd����.Enabled = True
                mshImprest.Active = False
                Txt����˵��.Enabled = False
            Else
                Txt��ҩ��λ.Enabled = True
                Cmd��Ӧ��.Enabled = True
                Cmdȫѡ.Enabled = True
                Cmd���.Enabled = True
                Cmd����.Enabled = False
            End If
            
        Else
            mintParallelRecord = 2
            Exit Sub
        End If
        
        'LockUserCons
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    
    CurLastMoney = 0
    TxtNo = mstr���ݺ�
    
    Me.Txt��ҩ��λ.Tag = 0
    tvwProvider.Tag = "0" '���Ϊ1,���ʾ��ѡ��;����Ϊδѡ��
    
    initCard
    RestoreWinState Me
    mshPurchaseList.Enabled = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
'    If Me.Height < 6135 Then
'        Me.Height = 6135
'    End If
'    If Me.Width < 11760 Then
'        Me.Width = 11760
'    End If
    
    With Fra2
        '.Top = (Me.ScaleHeight - .Height - Fra3.Height) / 2 + 30
        '.Top = (Me.ScaleHeight - .Height - Fra3.Height) + 30
        .Left = Me.ScaleWidth - .Width + 20
        .Height = Me.ScaleHeight - Fra3.Height
    End With
    
    With Lbl����
        .Left = Fra2.Width / 2 - .Width / 2
    End With
    
    With Fra3
        .Top = Fra2.Top + Fra2.Height - 120
        .Left = Fra2.Left
    End With
    
    With picdown
        .Top = Fra2.Height - .Height - 50
    End With
    
    With mshPaymentList
        .Height = picdown.Top - .Top - 50
    End With
        
    
    With Fra4
        .Top = Me.ScaleHeight - .Height + 20
        .Width = Fra2.Left
    End With
    
    With Fra1
        .Width = Fra2.Left
        .Height = Me.ScaleHeight - Fra4.Height + 220
    End With
    
    With mshImprest
        .Top = Fra1.Height - 1900
        .Height = 1800
        .Left = mshPurchaseList.Left   ' + 50
        .Width = Fra1.Width - 200
    End With
    
    With Label1
        .Left = Lbl�嵥.Left
        .Top = mshImprest.Top - .Height - 100
    End With
    
    With mshPurchaseList
        .Height = Label1.Top - .Top - 100
        .Width = Fra1.Width - 50
    End With
    
    With Cmd��ѡ�񸶿�
        .Left = Fra4.Width - .Width - 250
    End With
    
    With Txt������
        .Left = Fra1.Width - .Width - 100
    End With
    
    With Lbl������
        .Left = Txt������.Left - .Width - 100
    End With
    
    With Lbl�ϼ�
        .Left = Lbl������.Left - .Width - 100
    End With
    
    With Lbl�ۼƺϼ�
        .Left = Lbl�ϼ�.Left - .Width   '- 100
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me
End Sub

'ȡָ����ͷ����λ��
Private Function GetCol(mshFlex As MSHFlexGrid, ByVal ColName As String) As Integer
    Dim i As Integer
    
    GetCol = -1
    With mshFlex
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = ColName Then
                GetCol = i
                Exit Function
            End If
        Next
        
    End With
End Function


Private Function BanlanceMoney()
    Dim intRow As Integer
    Dim cur�ۼƺϼ� As Currency
    Dim curImprest As Currency
    Dim IntCol As Integer
    
    IntCol = GetCol(mshPurchaseList, "��Ʊ���")
    
    cur�ۼƺϼ� = 0
    For intRow = 1 To mshPurchaseList.rows - 1
        If mshPurchaseList.TextMatrix(intRow, 0) <> "" Then
            cur�ۼƺϼ� = cur�ۼƺϼ� + Val(mshPurchaseList.TextMatrix(intRow, IntCol))
        End If
    Next
    
    If cur�ۼƺϼ� <> 0 Then
        Txt������ = "[��" & GetFormat(cur�ۼƺϼ�, 2) & "]"
    Else
        Txt������ = ""
        If mint�༭״̬ = 1 Then
            With mshPaymentList
                .ClearMsf
                .Cols = 3
                .rows = 3
                .TextMatrix(0, 0) = "���ʽ"
                .TextMatrix(0, 1) = "������"
                .TextMatrix(0, 2) = "�������"
            End With
        End If
    End If
    
    curImprest = 0
    With mshImprest
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                curImprest = curImprest + .TextMatrix(intRow, 2)
            End If
        Next
        If curImprest > cur�ۼƺϼ� And cur�ۼƺϼ� > 0 Then
            Cmd��ѡ�񸶿�.Enabled = False
            Cmd����.Enabled = False
            If mintpurchaseclick = False Then
                MsgBox "�Բ���,��ǰѡ���Ԥ�����������Ӧ�����������ѡ��Ԥ���", vbOKOnly, gstrSysName
                mshImprest.SetFocus
            End If
            Exit Function
        End If
    End With
    
    If CurLastMoney <> (cur�ۼƺϼ� - curImprest) And cur�ۼƺϼ� <> 0 Then
        Cmd��ѡ�񸶿�.Enabled = True
        Cmd����.Enabled = False
        mshPaymentList.Active = False
    Else
        Cmd��ѡ�񸶿�.Enabled = False
        Cmd����.Enabled = False
    End If
    CurLastMoney = cur�ۼƺϼ� - curImprest
End Function


Private Sub Label2_Click()

End Sub

Private Sub mshImprest_DblClick(Cancel As Boolean)
    If mint�༭״̬ > 2 Then Exit Sub
    If mshImprest.TextMatrix(mshImprest.Row, 1) = "" Then
        Cancel = True
        Exit Sub
    End If
    With mshImprest
        If .TextMatrix(.Row, 0) = "" Then
            .TextMatrix(.Row, 0) = "��"
        Else
            .TextMatrix(.Row, 0) = ""
        End If
        Cancel = True
        BanlanceMoney
    End With
    
End Sub

Private Sub mshPaymentList_AfterDeleteRow()
    Dim Cur��� As Currency
    Dim intLop As Integer
    
    Cur��� = 0
    
    For intLop = 1 To mshPaymentList.rows - 1
        If intLop <> mshPaymentList.Row Then
            Cur��� = Cur��� + Val(mshPaymentList.TextMatrix(intLop, 1))
        End If
    Next
    Cur��� = CurLastMoney - Cur���
    
    If Cur��� <> 0 Then
        mshPaymentList.TextMatrix(mshPaymentList.Row, 1) = Format(Cur���, "#####0.00;-#####0.00; ;")
        mshPaymentList.TextMatrix(mshPaymentList.Row, 0) = mshPaymentList.CboText
    End If
End Sub

Private Sub mshPaymentList_cboClick(ListIndex As Long)
    With mshPaymentList
        If .Col <> 0 Then Exit Sub
        .TextMatrix(.Row, .Col) = .CboText
    End With
End Sub

Private Sub mshPaymentList_cboKeyDown(KeyCode As Integer, Shift As Integer)
    With mshPaymentList
        .TextMatrix(.Row, .Col) = .CboText
    End With
End Sub

Private Sub mshPaymentList_EnterCell(Row As Long, Col As Long)
    With mshPaymentList
    Select Case Col
        Case 1
            .TxtCheck = True
            .MaxLength = 16
            .TextMask = ".1234567890"
        Case 2
            .TxtCheck = True
            .MaxLength = 10
    End Select
    End With
    
End Sub

Private Sub mshPaymentList_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim intLop As Integer
    
    If mshPaymentList.Col = 2 Then
        If KeyCode <> vbKeyReturn Then
            mshPaymentList.ColData(2) = 4
            mshPaymentList.TxtCheck = False
        Else
            mshPaymentList.ColData(2) = 0
            mshPaymentList.TxtCheck = True
            mshPaymentList.TextLen = 10
        End If
    End If
    If mshPaymentList.Col = 1 _
            And mshPaymentList.Row = mshPaymentList.rows - 1 _
            And KeyCode = vbKeyReturn Then
        Txt����˵��.SetFocus
    End If
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mshPaymentList.TxtVisible = False Then Exit Sub
'    If ChkԤ����.Value = 1 Then Exit Sub
    Dim Cur��� As Currency
    Dim curImprest As Currency
    
    
    If mshPaymentList.Col = 1 Then
        Cur��� = 0
        For intLop = 1 To mshPaymentList.rows - 1
            If intLop <> mshPaymentList.Row Then
                Cur��� = Cur��� + Val(mshPaymentList.TextMatrix(intLop, 1))
            End If
        Next
        
        Cur��� = CurLastMoney - Cur���
        
        
        
        If Val(mshPaymentList.Text) = 0 And Cur��� > 0 Then
            MsgBox "�������Ϊ��!", vbInformation, gstrSysName
            Cancel = True
            mshPaymentList.TxtSetFocus
            Exit Sub
        End If
        If Not IsNumeric(mshPaymentList.Text) And Trim(mshPaymentList.Text) <> "" Then
            MsgBox "�������к��зǷ��ַ�!", vbInformation, gstrSysName
            Cancel = True
            mshPaymentList.TxtSetFocus
            Exit Sub
        End If
        If Val(mshPaymentList.Text) < 0 Then
            MsgBox "�����¼����Ϊ����!", vbInformation, gstrSysName
            Cancel = True
            mshPaymentList.TxtSetFocus
            Exit Sub
        End If
        If Val(mshPaymentList.Text) >= 10 ^ 14 - 1 Then
            MsgBox "���������С��" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
            Cancel = True
            mshPaymentList.TxtSetFocus
            Exit Sub
        End If
        If Trim(mshPaymentList.Text) = "" Then Exit Sub
        Cur��� = Cur��� - IIf(Trim(mshPaymentList.Text) = "", 0, mshPaymentList.Text)
        If Cur��� < 0 Then
            MsgBox "��������ܶ�!", vbInformation, gstrSysName
            Cancel = True
            mshPaymentList.TxtSetFocus
            Exit Sub
        End If
        If mshPaymentList.Row >= mshPaymentList.rows - 1 And Cur��� > 0 Then
            mshPaymentList.rows = mshPaymentList.rows + 1
        End If
                
        mshPaymentList.Text = GetFormat(mshPaymentList.Text, 2)
        If Cur��� > 0 Then
            mshPaymentList.TextMatrix(mshPaymentList.Row + 1, 1) = GetFormat(Cur���, 2)
            mshPaymentList.TextMatrix(mshPaymentList.Row + 1, 0) = mshPaymentList.CboText
        End If
    End If
End Sub

Private Sub mshPurchaseList_DblClick()
    If mint�༭״̬ > 2 Then Exit Sub
    
    If mshPurchaseList.TextMatrix(mshPurchaseList.Row, 1) = "" Then Exit Sub
    
    If mshPurchaseList.TextMatrix(mshPurchaseList.Row, 0) <> "" Then
        mshPurchaseList.TextMatrix(mshPurchaseList.Row, 0) = ""
    Else
        mshPurchaseList.TextMatrix(mshPurchaseList.Row, 0) = "��"
    End If
    mintpurchaseclick = True
    BanlanceMoney
    mintpurchaseclick = False
End Sub

Private Sub mshPurchaseList_DragDrop(Source As Control, x As Single, y As Single)
    If mshPurchaseList.Tag = "" Then Exit Sub
    If mshPurchaseList.MouseCol = 0 Then Exit Sub
    mshPurchaseList.Redraw = False
    mshPurchaseList.ColPosition(Val(mshPurchaseList.Tag)) = mshPurchaseList.MouseCol
    DoSort
    mshPurchaseList.Redraw = True
End Sub

Private Sub mshPurchaseList_GotFocus()
    tvwProvider.Visible = False
End Sub

Private Sub mshPurchaseList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then mshPurchaseList_DblClick
End Sub


Private Sub mshPurchaseList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mshPurchaseList.Tag = ""
    If mshPurchaseList.MouseRow <> 0 Then Exit Sub
    If mshPurchaseList.MouseCol = 0 Then Exit Sub
    mshPurchaseList.Tag = Str(mshPurchaseList.MouseCol)
    mshPurchaseList.Drag 1
End Sub

Private Sub tvwProvider_DblClick()
    Dim rsProvider As New Recordset
    
    On Error GoTo errHandle
    If tvwProvider.SelectedItem.Children <> 0 Then Exit Sub
    If tvwProvider.SelectedItem.Tag = 0 Then Exit Sub
    
    Txt��ҩ��λ = tvwProvider.SelectedItem
    Txt��ҩ��λ.Tag = Mid(tvwProvider.SelectedItem.Key, 3)
    tvwProvider.Tag = "1"
    tvwProvider.Visible = False
    
    With rsProvider
        gstrSQL = "Select ����,����,��ַ||�绰 as �绰��ַ,��������,�ʺ�,˰��ǼǺ� " _
            & " From ҩƷ��Ӧ��  " _
            & "Where id=" & Txt��ҩ��λ.Tag
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, "tvwProvider_DblClick")
        Call SQLTest
        
        If .EOF Then Exit Sub
        
        txt��λ���� = "[" & !���� & "]" & !����
        txt�绰��ַ = IIf(IsNull(!�绰��ַ), "", !�绰��ַ)
        txt������ = IIf(IsNull(!��������), "", !��������)
        txt�����ʺ� = IIf(IsNull(!�ʺ�), "", !�ʺ�)
        txt˰��� = IIf(IsNull(!˰��ǼǺ�), "", !˰��ǼǺ�)
    End With

    Call RefreshPurchaseList
    Call RefreshImprest
    Call BanlanceMoney
    
'    ChkԤ����.Enabled = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'ˢ�¸����嵥
Private Function RefreshPurchaseList()
    Dim Rec�ۼƺϼ� As New ADODB.Recordset
    Dim rsPayment As New Recordset
    Dim strLevel As String
    Dim strUnitName As String
    Dim str��װϵ�� As String
    Dim intLop As Integer
    
    Txt���� = ""
    Txt��Ʊ�� = ""
    On Error GoTo errHandle
    With rsPayment
        If .State = 1 Then .Close
        
        If glngSys \ 100 = 8 Then
            strUnitName = Choose(mintUnit + 1, "b.ҩ�ⵥλ", "b.�ۼ۵�λ")
            str��װϵ�� = Choose(mintUnit + 1, "b.ҩ���װ", "1")
        Else
            strUnitName = Choose(mintUnit + 1, "b.ҩ�ⵥλ", "b.���ﵥλ", "b.סԺ��λ", "b.�ۼ۵�λ")
            str��װϵ�� = Choose(mintUnit + 1, "b.ҩ���װ", "b.�����װ", "b.סԺ��װ", "1")
        End If
        
        
        If mint�༭״̬ = 1 Then
            
            gstrSQL = "SELECT distinct a.������� AS �������, a.no, c.��Ʊ��, c.��Ʊ���," _
                & "('[' || b.���� || ']' || decode(e.����,null,d.ͨ������,e.����)) AS ҩƷ��Ϣ, b.���, " _
                & "a.�ɱ���*" & str��װϵ�� & " AS �ɹ���, b.ָ�������� * a.ʵ������ AS �������, " _
                & "b.ָ��������*" & str��װϵ�� & " as ָ��������, a.����, a.����," & strUnitName & " AS ��λ, a.ʵ������/" & str��װϵ�� & "  AS ����," _
                & "a.���ۼ�*" & str��װϵ�� & " AS ����, a.ʵ������ * a.���ۼ� AS �ۼ۽��, c.�շ�id,c.������� " _
                & " FROM (SELECT * From ҩƷ�շ���¼ Where ���� = 1 AND ��ҩ��λid =" & Txt��ҩ��λ.Tag _
                        & " AND ����� IS NOT NULL) a," _
                    & " ҩƷĿ¼ b," _
                    & " ҩƷӦ����¼ c, " _
                    & " ҩƷ��Ϣ d," _
                    & " ҩƷ���� e " _
               & " Where c.�շ�id = a.ID " _
                 & " AND a.ҩƷid = b.ҩƷid " _
                 & " AND b.ҩƷid = e.ҩƷid (+) " _
                 & " AND b.ҩ��id = d.ҩ��id " _
                 & " AND c.��Ʊ�� IS NOT NULL " _
                 & " AND c.��Ʊ��� <> 0 " _
                 & " AND c.������� IS NULL " _
                 & " AND c.��ҩ��λid IS NOT NULL " _
                 & " AND c.��ҩ��λid =" & Txt��ҩ��λ.Tag _
               & " ORDER BY c.��Ʊ��, a.no "
                
        ElseIf mint�༭״̬ = 2 Then
            '�޸ĸ��
            gstrSQL = "SELECT distinct a.������� AS �������, a.no, c.��Ʊ��, c.��Ʊ���," _
                & "('[' || b.���� || ']' || decode(e.����,null,d.ͨ������,e.����)) AS ҩƷ��Ϣ, b.���, " _
                & "a.�ɱ���*" & str��װϵ�� & " AS �ɹ���, b.ָ�������� * a.ʵ������ AS �������, " _
                & "b.ָ��������*" & str��װϵ�� & " as ָ��������, a.����, a.����," & strUnitName & " AS ��λ, a.ʵ������/" & str��װϵ�� & "  AS ����," _
                & "a.���ۼ�*" & str��װϵ�� & " AS ����, a.ʵ������ * a.���ۼ� AS �ۼ۽��, c.�շ�id,c.������� " _
                & " FROM (SELECT * From ҩƷ�շ���¼ Where ���� = 1 AND ��ҩ��λid =" & Txt��ҩ��λ.Tag _
                        & " AND ����� IS NOT NULL) a," _
                    & " ҩƷĿ¼ b," _
                    & " ҩƷӦ����¼ c, " _
                    & " ҩƷ��Ϣ d," _
                    & " ҩƷ���� e " _
               & " Where c.�շ�id = a.ID " _
                 & " AND a.ҩƷid = b.ҩƷid " _
                 & " AND b.ҩƷid = e.ҩƷid (+) " _
                 & " AND b.ҩ��id = d.ҩ��id " _
                 & " AND c.��Ʊ�� IS NOT NULL " _
                 & " AND c.��Ʊ��� <> 0 " _
                 & " AND c.������� IS NULL " _
                 & " AND c.��ҩ��λid =" & Txt��ҩ��λ.Tag  '_
               '& " ORDER BY c.��Ʊ��, a.no "
               
             gstrSQL = gstrSQL & _
                 " union " _
                & "SELECT distinct a.������� AS �������, a.no, c.��Ʊ��, c.��Ʊ���," _
                & "('[' || b.���� || ']' || decode(e.����,null,d.ͨ������,e.����)) AS ҩƷ��Ϣ, b.���, " _
                & "a.�ɱ���*" & str��װϵ�� & " AS �ɹ���, b.ָ�������� * a.ʵ������ AS �������, " _
                & "b.ָ��������*" & str��װϵ�� & " as ָ��������, a.����, a.����," & strUnitName & " AS ��λ, a.ʵ������/" & str��װϵ�� & "  AS ����," _
                & "a.���ۼ�*" & str��װϵ�� & " AS ����, a.ʵ������ * a.���ۼ� AS �ۼ۽��, c.�շ�id,c.������� " _
                & " FROM (SELECT * From ҩƷ�շ���¼ Where ���� = 1 AND ��ҩ��λid =" & Txt��ҩ��λ.Tag _
                        & " AND ����� IS NOT NULL) a," _
                    & " ҩƷĿ¼ b," _
                    & " ҩƷӦ����¼ c, " _
                    & " ҩƷ��Ϣ d," _
                    & " ҩƷ���� e " _
               & " Where c.�շ�id = a.ID " _
                 & " AND a.ҩƷid = b.ҩƷid " _
                 & " AND b.ҩƷid = e.ҩƷid (+) " _
                 & " AND b.ҩ��id = d.ҩ��id " _
                 & " AND c.��Ʊ�� IS NOT NULL " _
                 & " AND c.��Ʊ��� <> 0 " _
                 & " AND c.������� =" & Txt����˵��.Tag _
                 & " AND c.��ҩ��λid =" & Txt��ҩ��λ.Tag
                    
            '   & " ORDER BY c.��Ʊ��, a.no "
        Else
            gstrSQL = "SELECT distinct a.������� AS �������, a.no, c.��Ʊ��, c.��Ʊ���," _
                & "('[' || b.���� || ']' || decode(e.����,null,d.ͨ������,e.����)) AS ҩƷ��Ϣ, b.���, " _
                & "a.�ɱ���*" & str��װϵ�� & " AS �ɹ���, b.ָ�������� * a.ʵ������ AS �������, " _
                & "b.ָ��������*" & str��װϵ�� & " as ָ��������, a.����, a.����," & strUnitName & " AS ��λ, a.ʵ������/" & str��װϵ�� & "  AS ����," _
                & "a.���ۼ�*" & str��װϵ�� & " AS ����, a.ʵ������ * a.���ۼ� AS �ۼ۽��, c.�շ�id,c.������� " _
                & " FROM (SELECT * From ҩƷ�շ���¼ Where ���� = 1 AND ��ҩ��λid =" & Txt��ҩ��λ.Tag _
                        & " AND ����� IS NOT NULL) a," _
                    & " ҩƷĿ¼ b," _
                    & " ҩƷӦ����¼ c, " _
                    & " ҩƷ��Ϣ d," _
                    & " ҩƷ���� e " _
               & " Where c.�շ�id = a.ID " _
                 & " AND a.ҩƷid = b.ҩƷid " _
                 & " AND b.ҩƷid = e.ҩƷid (+) " _
                 & " AND b.ҩ��id = d.ҩ��id " _
                 & " AND c.��Ʊ�� IS NOT NULL " _
                 & " AND c.��Ʊ��� <> 0 " _
                 & " AND c.������� =" & Txt����˵��.Tag _
                 & " AND c.��ҩ��λid =" & Txt��ҩ��λ.Tag _
               & " ORDER BY c.��Ʊ��, a.no "
        End If
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        Set rsPayment = zldatabase.OpenSQLRecord(gstrSQL, "RefreshPurchaseList")
        Call SQLTest
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            initGrid
        End If
        
        If .RecordCount > 0 And InStr(1, "12", mint�༭״̬) <> 0 Then
            Cmdȫѡ.Enabled = True
            Cmd���.Enabled = True
        Else
            Cmdȫѡ.Enabled = False
            Cmd���.Enabled = False
        End If
        If .EOF Then Exit Function
        .MoveFirst
        For intLop = 1 To .RecordCount
            mshPurchaseList.TextMatrix(intLop, mconintCol��־) = IIf(IsNull(!�������), "", "��")
            mshPurchaseList.TextMatrix(intLop, mconintcol��Ʊ��) = !��Ʊ��
            mshPurchaseList.TextMatrix(intLop, mconintcol��ⵥ��) = !No
            mshPurchaseList.TextMatrix(intLop, mconintcolҩƷ��Ϣ) = !ҩƷ��Ϣ
            mshPurchaseList.TextMatrix(intLop, mconIntCol���) = IIf(IsNull(!���), "", !���)
            mshPurchaseList.TextMatrix(intLop, mconIntCol��λ) = !��λ
            mshPurchaseList.TextMatrix(intLop, mconintcol��Ʊ���) = GetFormat(!��Ʊ���, 2)
            mshPurchaseList.TextMatrix(intLop, mconIntCol����) = GetFormat(!����, 3)
            
            mshPurchaseList.TextMatrix(intLop, mconIntCol�ɹ���) = GetFormat(!�ɹ���, 4)
            mshPurchaseList.TextMatrix(intLop, mconintcol������) = GetFormat(!ָ��������, 4)
            mshPurchaseList.TextMatrix(intLop, mconintcol�������) = GetFormat(!�������, 2)
            
            mshPurchaseList.TextMatrix(intLop, mconIntCol�ۼ�) = GetFormat(!����, 4)
            mshPurchaseList.TextMatrix(intLop, mconIntCol�ۼ۽��) = GetFormat(!�ۼ۽��, 4)
            
            mshPurchaseList.TextMatrix(intLop, mconIntCol����) = IIf(IsNull(!����), "", !����)
            mshPurchaseList.TextMatrix(intLop, mconIntCol����) = IIf(IsNull(!����), "", !����)
            
            mshPurchaseList.TextMatrix(intLop, mconintcol�������) = Format(IIf(IsNull(!�������), "", !�������), "yyyy-MM-dd")
            
            mshPurchaseList.RowData(intLop) = !�շ�id
            If intLop = mshPurchaseList.rows - 1 Then mshPurchaseList.rows = mshPurchaseList.rows + 1
            .MoveNext
        Next
        
    End With
    
    With Rec�ۼƺϼ�
        If .State = 1 Then .Close
        If mint�༭״̬ = 1 Then
            gstrSQL = "Select Sum(��Ʊ���) as �ϼ� From ҩƷӦ����¼  " _
                   & " Where ��Ʊ�� is Not Null And ��Ʊ���<>0 " _
                   & "   And ������� is Null " _
                   & "   and ��ҩ��λID=" & Txt��ҩ��λ.Tag
            Call SQLTest(App.Title, Me.Caption, gstrSQL)
            Set Rec�ۼƺϼ� = zldatabase.OpenSQLRecord(gstrSQL, "cmd����_Click")
            Call SQLTest
            
        Else
            gstrSQL = "Select sum(�ϼ�) as �ϼ� " _
                   & "  From (" & _
                          " Select Sum(��Ʊ���) as �ϼ� From ҩƷӦ����¼ Where ��Ʊ�� is Not Null And ��Ʊ���<>0  And ��ҩ��λID=" & Txt��ҩ��λ.Tag & " And ������� is Null" _
                        & " Union Select Sum(��Ʊ���) as �ϼ� From ҩƷӦ����¼ Where ��Ʊ�� is Not Null And ��Ʊ���<>0 And ��ҩ��λID=" & Txt��ҩ��λ.Tag & " And �������=" & Txt����˵��.Tag & ")"
            Call SQLTest(App.Title, Me.Caption, gstrSQL)
            Set Rec�ۼƺϼ� = zldatabase.OpenSQLRecord(gstrSQL, "cmd����_Click")
            Call SQLTest
            
        End If
        Lbl�ϼ� = GetFormat(!�ϼ�, 2)
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'ˢ��Ԥ����
Private Sub RefreshImprest()
    Dim rsImprest As New Recordset
    Dim intRow As Integer
    Dim intRecord As Integer
    
    On Error GoTo errHandle
    If mint�༭״̬ = 1 Then
        gstrSQL = "select id,���㷽ʽ,�������,���,������� " _
                & " from ҩƷ�����¼ " _
               & " where ��λid=" & Txt��ҩ��λ.Tag _
                & "  and Ԥ����=1 " _
                & "  and nvl(�������,0)=0 " _
                & "  and ������� is not null "
                '& "  and ��¼״̬=1 "
        gstrSQL = gstrSQL _
               & " union all " _
               & "select id,���㷽ʽ,�������,���,������� " _
                & " from ҩƷ�����¼ " _
               & " where ��λid=" & Txt��ҩ��λ.Tag _
                & "  and nvl(Ԥ����,0)=0 " _
                & "  and nvl(�������,0)=0 " _
                & "  and ������� is not null " _
                & "  and ��¼״̬=2 "
        
                
    ElseIf mint�༭״̬ = 2 Then
        gstrSQL = "select id,���㷽ʽ,�������,���,������� " _
                & " from ҩƷ�����¼ " _
               & " where ��λid=" & Txt��ҩ��λ.Tag _
                & "  and Ԥ����=1 " _
                & "  and nvl(�������,0)=0 " _
                & "  and ������� is not null "
                '& "  and ��¼״̬=1 "
        gstrSQL = gstrSQL _
            & " union all " _
               & "select id,���㷽ʽ,�������,���,������� " _
                & " from ҩƷ�����¼ " _
               & " where ��λid=" & Txt��ҩ��λ.Tag _
                & "  and nvl(Ԥ����,0)=0 " _
                & "  and nvl(�������,0)=0 " _
                & "  and ������� is not null " _
                & "  and ��¼״̬=2 " _
            & " union " _
            & "select id,���㷽ʽ,�������,��� ,�������" _
                & " from ҩƷ�����¼ " _
               & " where ��λid=" & Txt��ҩ��λ.Tag _
                & "  and Ԥ����=1 " _
                & "  and nvl(�������,0)=" & Txt����˵��.Tag _
                & "  and ������� is not null " _
            & "union  select id,���㷽ʽ,�������,��� ,�������" _
                & " from ҩƷ�����¼ " _
               & " where ��λid=" & Txt��ҩ��λ.Tag _
                & "  and nvl(Ԥ����,0)=0 " _
                & "  and nvl(�������,0)=" & Txt����˵��.Tag _
                & "  and ������� is not null " _
                & "  and (��¼״̬=2) "
    Else
        gstrSQL = "select id,���㷽ʽ,�������,���,������� " _
                & " from ҩƷ�����¼ " _
               & " where ��λid=" & Txt��ҩ��λ.Tag _
                & "  and Ԥ����=1 " _
                & "  and nvl(�������,0)=" & Txt����˵��.Tag _
                & "  and ������� is not null " _
                & "union  select id,���㷽ʽ,�������,��� ,�������" _
                & " from ҩƷ�����¼ " _
               & " where ��λid=" & Txt��ҩ��λ.Tag _
                & "  and nvl(Ԥ����,0)=0 " _
                & "  and nvl(�������,0)=" & Txt����˵��.Tag _
                & "  and ������� is not null " _
                & "  and (��¼״̬=2) "
                
                '& "  and (��¼״̬=1 or ��¼״̬=3)  "
                

    End If
    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    Set rsImprest = zldatabase.OpenSQLRecord(gstrSQL, "RefreshImprest")
    Call SQLTest
    
    If rsImprest.EOF Then Exit Sub
    intRecord = rsImprest.RecordCount
    rsImprest.MoveFirst
    With mshImprest
        For intRow = 1 To intRecord
            .TextMatrix(intRow, 0) = IIf(IIf(IsNull(rsImprest!�������), 0, rsImprest!�������) > 0, "��", "")
            .TextMatrix(intRow, 1) = rsImprest!���㷽ʽ
            .TextMatrix(intRow, 2) = rsImprest!���
            .TextMatrix(intRow, 3) = IIf(IsNull(rsImprest!�������), "", rsImprest!�������)
            .RowData(intRow) = rsImprest!Id
            If intRow = .rows - 1 Then .rows = .rows + 1
            rsImprest.MoveNext
        Next
        rsImprest.Close
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub tvwProvider_LostFocus()
'    tvwProvider.Visible = False
End Sub

Private Sub TxtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Txt����_GotFocus()
    With Txt����
        .SelStart = 0
        .SelLength = 100
    End With
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then mshPurchaseList.SetFocus
End Sub

Private Sub Txt����_Validate(Cancel As Boolean)
    Call SelAccord
End Sub

Private Sub Txt��Ʊ��_GotFocus()
    With Txt��Ʊ��
        .SelStart = 0
        .SelLength = 100
    End With
End Sub

Private Sub Txt��Ʊ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then TxtNo.SetFocus
End Sub

Private Sub Txt��Ʊ��_Validate(Cancel As Boolean)
    Call SelAccord
End Sub

Private Sub txt����˵��_GotFocus()
    tvwProvider.Visible = False
End Sub

Private Sub txt����˵��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then If Cmd����.Enabled Then Cmd����.SetFocus
End Sub

Private Sub Txt��ҩ��λ_GotFocus()
    tvwProvider.Visible = False
End Sub

Private Sub txt��ҩ��λ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String
    Dim rec��Ӧ�� As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If LTrim(RTrim(Txt��ҩ��λ)) = "" Then Exit Sub
    If InStr(1, Txt��ҩ��λ, "[") <> 0 Then
        If InStr(2, Txt��ҩ��λ, "]") <> 0 Then
            strInput = Mid(Txt��ҩ��λ.Text, 2, InStr(2, Txt��ҩ��λ, "]") - 2)
        Else
            strInput = Mid(Txt��ҩ��λ.Text, 2)
        End If
    Else
        strInput = Txt��ҩ��λ.Text
    End If
    
    With rec��Ӧ��
        gstrSQL = "Select ID,����,����,����,��ַ||�绰 as �绰��ַ,��������,�ʺ�,˰��ǼǺ�  From ҩƷ��Ӧ�� Where (���� like '" & UCase(strInput) & "%' Or ���� like '" & UCase(strInput) & "%' Or ���� like '" & UCase(strInput) & "%') And To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' And ĩ��=1"
        Call OpenRecordset(rec��Ӧ��, "ҩƷ��Ӧ��")
        
        If .EOF Then
            MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
            KeyCode = 0
            Txt��ҩ��λ = ""
            tvwProvider.Tag = "0"
            Exit Sub
        End If
        If .RecordCount > 1 Then
            Set mshProvider.Recordset = rec��Ӧ��
            SetProviderWidth Txt��ҩ��λ.Left, Txt��ҩ��λ.Top + Txt��ҩ��λ.Height + Fra1.Top
            Exit Sub
        Else
            Txt��ҩ��λ = "[" & !���� & "]" & !����
            Txt��ҩ��λ.Tag = !Id
            tvwProvider.Tag = "1"
        End If
    End With
    
    txt��λ���� = rec��Ӧ��!����
    txt�绰��ַ = IIf(IsNull(rec��Ӧ��!�绰��ַ), "", rec��Ӧ��!�绰��ַ)
    txt������ = IIf(IsNull(rec��Ӧ��!��������), "", rec��Ӧ��!��������)
    txt�����ʺ� = IIf(IsNull(rec��Ӧ��!�ʺ�), "", rec��Ӧ��!�ʺ�)
    txt˰��� = IIf(IsNull(rec��Ӧ��!˰��ǼǺ�), "", rec��Ӧ��!˰��ǼǺ�)
    Call RefreshPurchaseList
    Call RefreshImprest
    
    Call BanlanceMoney
    
End Sub


Private Function CheckData(ByVal ������ As Integer) As Boolean
    Dim IntCheck As Integer
    
    CheckData = False
    With mshPaymentList
        For IntCheck = 1 To ������
            If Val(.TextMatrix(IntCheck, 1)) = 0 And LTrim(RTrim(.TextMatrix(IntCheck, 1))) = "" Then
                MsgBox "��" & IntCheck & "�еĸ������Ϊ�㣡", vbInformation, gstrSysName
                Exit Function
            End If
            If Not IsNumeric(.TextMatrix(IntCheck, 1)) Then
                MsgBox "��" & IntCheck & "�еĸ������к��зǷ��ַ���", vbInformation, gstrSysName
                Exit Function
            End If
            If Val(.TextMatrix(IntCheck, 1)) > 10 ^ 11 - 1 Then
                MsgBox "��" & IntCheck & "�еĸ���������ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            If LenB(StrConv(.TextMatrix(IntCheck, 2), vbFromUnicode)) > 10 Then
                MsgBox "��" & IntCheck & "�еĽ�����볤�ȳ���!(���10���ַ�)", vbInformation, gstrSysName
                Exit Function
            End If
        Next
        If LenB(StrConv(Txt����˵��.Text, vbFromUnicode)) > 50 Then
            MsgBox "����˵���ĳ��ȳ���!(���Ϊ50���ַ���25������)", vbInformation, gstrSysName
            Txt����˵��.SetFocus
            Exit Function
        End If
        
        CheckData = True
    End With
End Function

Sub DoSort()
    
    mshPurchaseList.Col = 0
    mshPurchaseList.ColSel = mshPurchaseList.Cols - 1
    mshPurchaseList.Sort = 2 ' ��׼����
    
End Sub

Private Function Get������(ByVal LngPay As Long) As Long
    Dim Rec������ As New ADODB.Recordset
    If LngPay = 0 Then Get������ = 0: Exit Function
    With Rec������
        gstrSQL = "Select distinct ��Ʊ�� as PayCount From ҩƷӦ����¼ Where �������=" & LngPay
        Call OpenRecordset(Rec������, "������")
        
        If .EOF Then
            Get������ = 0
        Else
            Get������ = .RecordCount
        End If
    End With
End Function



Private Sub mshProvider_DblClick()
    mshProvider_KeyPress 13
End Sub

Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
    
    With mshProvider
        If KeyCode = vbKeyRight Then
            If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
                
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If .LeftCol <> 0 Then
                .LeftCol = .LeftCol - 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyHome Then
            If .LeftCol <> 0 Then
                .LeftCol = 0
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyEnd Then
            For i = .Cols - 1 To 0 Step -1
                sngWidth = sngWidth + .ColWidth(i)
                If sngWidth > .Width Then
                    .LeftCol = i + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                    Exit For
                End If
            Next
        ElseIf KeyCode = vbKeyReturn Then
            Call mshProvider_KeyPress(13)
        End If
    End With
End Sub

Private Sub mshProvider_KeyPress(KeyAscii As Integer)
    With mshProvider
        If KeyAscii = 13 Then
            Txt��ҩ��λ.Text = "[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2)
            Txt��ҩ��λ.Tag = .TextMatrix(.Row, 0)
            tvwProvider.Tag = "1"
            txt��λ����.Caption = .TextMatrix(.Row, 2)
            txt�绰��ַ = .TextMatrix(.Row, 4)
            txt�����ʺ� = .TextMatrix(.Row, 6)
            txt������ = .TextMatrix(.Row, 5)
            txt˰��� = .TextMatrix(.Row, 7)
            
            .Visible = False
            Call RefreshPurchaseList
            Call RefreshImprest
            Call BanlanceMoney
            mshPurchaseList.SetFocus
        End If
    End With
End Sub

Private Sub mshProvider_LostFocus()
    SaveFlexState mshProvider, Me.Caption
    If mshProvider.Visible Then mshProvider.Visible = False
End Sub


'���ù�Ӧ��ѡ�����Ŀ�ȼ��������
Private Sub SetProviderWidth(ByVal sngLeft As Single, ByVal sngTop As Single)
    
    With mshProvider
        .Visible = True
        .Left = sngLeft
        .Top = sngTop
'        If RestoreFlexState(mshProvider, Me.Caption) = False Then
            'Select ID,����,����,����,��ַ||�绰 as �绰��ַ,��������,�ʺ�,˰��ǼǺ�
            
            .ColWidth(0) = 0
            .ColWidth(1) = 1000
            .ColWidth(2) = 2500
            .ColWidth(3) = 1000
            
            .ColWidth(4) = 1500
            .ColWidth(5) = 1500
            .ColWidth(6) = 1000
            .ColWidth(7) = 1000
            
'        End If
        
        .SetFocus
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub SelAccord()
    '�����û�����ķ�Ʊ�ż�NO��ѡ����Ӧ�ļ�¼
    Dim strInvoice As String, strBill As String, StrTmp As String
    Dim intDo As Integer, lngRow As Long, lngRows As Long, lngLastRow As Long
    Dim intYear As Integer, strYear As String
    Dim arrInvoice, arrBill, blnFind As Boolean
    
    lngLastRow = mshPurchaseList.Row
    strInvoice = Trim(Txt��Ʊ��)
    strBill = Trim(Txt����)
    If strInvoice = "" And strBill = "" Then Exit Sub
    
    '��������ʽ
    arrInvoice = Split(strInvoice, "-")
    arrBill = Split(strBill, "-")
    If UBound(arrInvoice) > 1 Then
        MsgBox "�����ʽ���ԣ�123��123-300�������������룡", vbInformation, gstrSysName
        Txt��Ʊ��.SetFocus
        Exit Sub
    End If
    If UBound(arrBill) > 1 Then
        MsgBox "�����ʽ���ԣ�C0000001��C0000001-C0000020�������������룡", vbInformation, gstrSysName
        Txt����.SetFocus
        Exit Sub
    End If
    
    '--���������λ,�򰴹������--
    Txt���� = ""
    For intDo = 0 To UBound(arrBill)
        StrTmp = UCase(LTrim(arrBill(intDo)))
        If Len(StrTmp) < 8 Then
            intYear = Format(zldatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            StrTmp = strYear & String(7 - Len(StrTmp), "0") & StrTmp
        End If
        arrBill(intDo) = StrTmp
        Txt���� = Txt���� & IIf(Txt���� = "", "", "-") & StrTmp
    Next
    
    'ѭ��ѡ��
    Call Cmd���_Click
    lngRows = mshPurchaseList.rows - 1
    mshPurchaseList.Redraw = False
    For lngRow = 1 To lngRows
        blnFind = False
        If strInvoice <> "" And strBill <> "" Then
            '����Ϊ��
            StrTmp = mshPurchaseList.TextMatrix(lngRow, mconintcol��Ʊ��)
            If UBound(arrInvoice) = 1 Then
                If arrInvoice(0) <= StrTmp Then
                    blnFind = (StrTmp <= arrInvoice(1))
                End If
            Else
                blnFind = (StrTmp = arrInvoice(0))
            End If
            If blnFind Then
                blnFind = False
                StrTmp = mshPurchaseList.TextMatrix(lngRow, mconintcol��ⵥ��)
                If UBound(arrBill) = 1 Then
                    If arrBill(0) <= StrTmp Then
                        blnFind = (StrTmp <= arrBill(1))
                    End If
                Else
                    blnFind = (arrBill(0) = StrTmp)
                End If
            End If
        ElseIf strInvoice <> "" Then
            '�����뷢Ʊ��
            StrTmp = mshPurchaseList.TextMatrix(lngRow, mconintcol��Ʊ��)
            If UBound(arrInvoice) = 1 Then
                If arrInvoice(0) <= StrTmp Then
                    blnFind = (StrTmp <= arrInvoice(1))
                End If
            Else
                blnFind = (StrTmp = arrInvoice(0))
            End If
        Else
            '�����뵥�ݺ�
            StrTmp = mshPurchaseList.TextMatrix(lngRow, mconintcol��ⵥ��)
            If UBound(arrBill) = 1 Then
                If arrBill(0) <= StrTmp Then
                    blnFind = (StrTmp <= arrBill(1))
                End If
            Else
                blnFind = (arrBill(0) = StrTmp)
            End If
        End If
        
        '����ҵ���ִ��˫���¼�
        If blnFind And Trim(mshPurchaseList.TextMatrix(lngRow, mconintCol��־)) = "" Then
            With mshPurchaseList
                .Row = lngRow
                .Col = 1
            End With
            Call mshPurchaseList_DblClick
        End If
    Next
    mshPurchaseList.Row = lngLastRow
    mshPurchaseList.Redraw = True
End Sub
