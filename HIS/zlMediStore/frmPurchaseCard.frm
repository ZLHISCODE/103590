VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmPurchaseCard 
   Caption         =   "ҩƷ�⹺��ⵥ"
   ClientHeight    =   7320
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12495
   Icon            =   "frmPurchaseCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdAddProducer 
      Caption         =   "����������(&P)"
      Height          =   350
      Left            =   2520
      TabIndex        =   54
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "��������(&C)"
      Height          =   350
      Left            =   6600
      TabIndex        =   53
      ToolTipText     =   "���Ƶ�ǰ�з�Ʊ��ϢӦ���������޷�Ʊ��Ϣ��"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdALLDel 
      Caption         =   "ȫ��(&D)"
      Height          =   350
      Left            =   4680
      TabIndex        =   52
      ToolTipText     =   "��������еķ�Ʊ�������"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.PictureBox picInputCost 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   5760
      ScaleHeight     =   1635
      ScaleWidth      =   5415
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   5445
      Begin VB.CommandButton cmdGetData 
         Caption         =   "��ȡ(&G)"
         Height          =   300
         Left            =   4550
         TabIndex        =   51
         Top             =   0
         Width           =   855
      End
      Begin VB.ComboBox cboInputDate 
         Height          =   300
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   0
         Width           =   1440
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfInputCost 
         Height          =   1300
         Left            =   0
         TabIndex        =   48
         Top             =   300
         Width           =   5415
         _cx             =   9551
         _cy             =   2293
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPurchaseCard.frx":014A
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʱ��"
         Height          =   180
         Left            =   0
         TabIndex        =   50
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdGetInputCost 
      Caption         =   "��"
      Height          =   300
      Left            =   1320
      TabIndex        =   46
      Top             =   6120
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CheckBox chkת���ƿ� 
      Caption         =   "������ⵥҩƷ�ƿ⵽"
      Height          =   270
      Left            =   4680
      TabIndex        =   41
      Top             =   5700
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.ComboBox cboEnterStock 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "frmPurchaseCard.frx":0204
      Left            =   6915
      List            =   "frmPurchaseCard.frx":020D
      TabIndex        =   40
      Text            =   "cboEnterStock"
      Top             =   5685
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.PictureBox PicInput 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   210
      ScaleHeight     =   1635
      ScaleWidth      =   2775
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   2805
      Begin VB.TextBox Txt�Ӽ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         MaxLength       =   8
         TabIndex        =   37
         Text            =   "15.0000"
         Top             =   690
         Width           =   1725
      End
      Begin VB.CommandButton CmdNO 
         Caption         =   "ȡ��"
         Height          =   345
         Left            =   1800
         TabIndex        =   39
         Top             =   1140
         Width           =   855
      End
      Begin VB.CommandButton CmdYes 
         Caption         =   "ȷ��"
         Height          =   345
         Left            =   810
         TabIndex        =   38
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "    ������ӳ��ʣ����ۼ۵ļ��㹫ʽ�����ۼ�=�ɱ���*(1+�ӳ���%)"
         ForeColor       =   &H00400000&
         Height          =   585
         Left            =   0
         TabIndex        =   35
         Top             =   150
         Width           =   2805
      End
      Begin VB.Label Lbl�Ӽ��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ӳ���(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   90
         TabIndex        =   36
         Top             =   750
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   9975
      TabIndex        =   33
      Top             =   6135
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   8655
      TabIndex        =   32
      Top             =   6135
      Visible         =   0   'False
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
      Height          =   2175
      Left            =   2520
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3030
      TabIndex        =   14
      Top             =   5685
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   1290
      TabIndex        =   13
      Top             =   5640
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   -15
      TabIndex        =   12
      Top             =   5640
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8655
      TabIndex        =   10
      Top             =   5655
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9975
      TabIndex        =   11
      Top             =   5655
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   5325
      Left            =   0
      ScaleHeight     =   5265
      ScaleWidth      =   12375
      TabIndex        =   15
      Top             =   0
      Width           =   12435
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
         Height          =   1815
         Left            =   4800
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3201
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   32768
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   7
         Top             =   1020
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   9
         Top             =   4080
         Width           =   10410
      End
      Begin VB.CommandButton cmdProvider 
         Caption         =   "��"
         Height          =   300
         Left            =   11010
         TabIndex        =   20
         Top             =   660
         Width           =   300
      End
      Begin VB.TextBox txtProvider 
         Height          =   300
         Left            =   8055
         TabIndex        =   2
         Top             =   660
         Width           =   2895
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.TextBox txtNO 
         Height          =   315
         IMEMode         =   2  'OFF
         Left            =   9870
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   210
         Width           =   1425
      End
      Begin VB.Label lbl�޸����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�޸�����"
         Height          =   180
         Left            =   3240
         TabIndex        =   58
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label lbl�޸��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  �޸���"
         Height          =   180
         Left            =   3285
         TabIndex        =   57
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Txt�޸����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4050
         TabIndex        =   56
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt�޸��� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4050
         TabIndex        =   55
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label txt�˲��� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7290
         TabIndex        =   45
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label txt�˲����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7290
         TabIndex        =   44
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label lbl�˲��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  �˲���"
         Height          =   180
         Left            =   6525
         TabIndex        =   43
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label lbl�˲����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�˲�����"
         Height          =   180
         Left            =   6480
         TabIndex        =   42
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   29
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   2040
         TabIndex        =   28
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "������ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   27
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10410
         TabIndex        =   24
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10410
         TabIndex        =   23
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   22
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   21
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9390
         TabIndex        =   1
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ�⹺��ⵥ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   0
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ(&S)"
         Height          =   180
         Left            =   195
         TabIndex        =   3
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   19
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   9765
         TabIndex        =   17
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   9600
         TabIndex        =   16
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label LblProvider 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ��λ(&G)"
         Height          =   180
         Left            =   7035
         TabIndex        =   5
         Top             =   720
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":021B
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0435
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":064F
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0869
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0A83
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0C9D
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0EB7
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":10D1
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":12EB
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1505
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":171F
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1939
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1B53
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1D6D
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1F87
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":21A1
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   30
      Top             =   6960
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPurchaseCard.frx":23BB
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15690
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmPurchaseCard.frx":2C4F
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmPurchaseCard.frx":3151
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin VB.Label lblCode 
      Caption         =   "����"
      Height          =   255
      Left            =   2550
      TabIndex        =   25
      Top             =   5730
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuCol 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(���������)"
         Index           =   0
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(������)"
         Index           =   1
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(������)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmPurchaseCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng��ҩ��λID As Long              '��ҩ��λID
Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5���޸ķ�Ʊ��6��������
                                            '7��������ˣ������������µ��ݲ���ˣ��Ѹ���ĵ��ݲ������
                                            '����ˣ�ͬ����������˺�ĵ��ݲ����������;8-ҩ���˻�;9-�˲�
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼��7����ȫ������
Private mstrPrivs As String                 'Ȩ��
Private mintBatchNoLen As Integer           '���ݿ������Ŷ��峤��
Private mstr��� As String                  '������¼Ĭ����۵�ֵ
Private mstr���ս��� As String                  '������¼Ĭ�����ս���
Private mbln��ʾ As Boolean                 '��ҩƷѡ������ѡ���ҩƷ��������������ݵıȽϿ��Ƿ��ظ��������ظ�������ֻ��ʾһ�Σ�true �Ѿ���ʾ�ˣ�false��û����ʾ
Private mrs�ֶμӳ� As ADODB.Recordset      '�ֶμӳɼ���
'Private mblnʱ��ȡ�ϴ��ۼ� As Boolean        'ʱ��ҩƷֱ��ȥ�ϴ��ۼ�
Private mlng��װϵ�� As Long                '��¼��װϵ��
Private mbln��ʾ��ʽ As Boolean             '��ʾ��ʽ true-ֻ��ʾһ�Σ�false-������ʾ
Private mblnЧ����ʾ As Boolean             '�Ƿ���ʾʧЧ�ڵ�ҩƷ,��Ҫ�����ڼ��ص���ʱ�����Ĺ���ҩƷ��ʾ��true-��ʾ;false-����ʾ

Private marrFrom As Variant                   '��¼�û��ָ�������и���
Private marrInitGrid As Variant                '��¼��ʼ��������и���

Private mblnEnter As Boolean                '�Ƿ���뵥Ԫ��
Private mstr������� As String              '�������

Private mbln�޸������� As Boolean           '�����޸�������
Private mdbl�Ӽ��� As Double
Private mblnUpdate As Boolean               '��ʾ�Ƿ��Ѹ������¼۸���µ�������
Private mbln�˻� As Boolean                 '��ʾ�Ƿ����˻���
Private mbln�б�ҩƷ��ѡ����б굥λ��� As Boolean      '���ز�������
Private mstr�洢�ⷿ��ʾ As String
Private mbln����ҩƷ�޴洢�ⷿ As Boolean
Private mintȡ�ϴβɹ��۷�ʽ As Integer     '0-���ȴ�ҩƷ���ȡ;1-���ȴ�ҩƷ���ȡ
Private mbln��Ӧ��У�� As Boolean

Private mbln�Ӽ��� As Boolean               'ʱ��ҩƷ�Ƿ��������Ӽ���
Private mintʱ������ۼۼӳɷ�ʽ As Integer 'ϵͳ������ʱ��ҩƷ�⹺���ʱ�ۼۼ��㷽ʽ��0�����ۿۺ�Ĳɹ��ۼ����ۼ�;1�����ۿ�ǰ�Ĳɹ��ۼ����ۼۡ�
Private mintʱ�۷ֶμӳɷ�ʽ As Integer     ' 0-�����ֶμӳɣ�Ĭ�ϣ� 1-���ֶμӳ�
Private mint����� As Integer             '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ

Private mbln�����ֹ�����ӳ��� As Boolean
Private mstrColumn_UnSelected As String     '��¼��Щ�б�����Ϊ����ʾ
Public RecReturn As Recordset
Private mblnViewCost As Boolean             '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���

Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

Private mstrTime_Start As String                      '���뵥�ݱ༭����ʱ�����༭���ݵ�����޸�ʱ��
Private mstrTime_End As String                        '�˿̸ñ༭���ݵ�����޸�ʱ��

Private mlng���ⷿ As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

Private mstrControlItem As String           '�˲顢��ˡ�������˻��������޸ĵ���Ŀ

Private mintLastCol As Integer              '�û����������е����ɼ��е��к�

Private mintDrugNameShow As Integer         'ҩƷ��ʾ��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����

Private mblnMSH_GetFocus As Boolean         '����ֻһ����ʾ
Private mlng�����̳��� As Long                 '�������ֶγ���
Private mlngԭ���س��� As Long                 'ԭ�����ֶγ���

Private mblnȡĿ¼�в�����Ϣ As Boolean

Private Const MStrCaption As String = "ҩƷ�⹺������"

Private Enum ����

    �˲� = 1
    ��� = 2
    ������� = 3
End Enum

Private mblnLoad As Boolean              '��¼�Ƿ�ִ�����Form_Load�¼�

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintNumberDigit As Integer      '����С��λ��
Private mintMoneyDigit As Integer       '���С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4


'=========================================================================================
Private mconIntCol�к� As Integer
Private mconIntColҩ�� As Integer
Private mconIntCol��Ʒ�� As Integer
Private mconIntCol��Դ As Integer
Private mconIntCol����ҩ�� As Integer
Private mconIntCol��� As Integer
Private mconIntCol��� As Integer
Private mconIntColҩ�ۼ��� As Integer
Private mconIntColԭ������ As Integer
Private mconIntColԭ���� As Integer
Private mconIntCol����ϵ�� As Integer
Private mconintcol���� As Integer
Private mconIntCol���� As Integer
Private mconIntColԭ���� As Integer
Private mconIntCol��λ As Integer
Private mconIntCol���� As Integer
Private mconIntCol�������� As Integer
Private mconIntColЧ�� As Integer
Private mconIntCol���� As Integer
Private mconIntCol�������� As Integer
Private mconIntCol���� As Integer
Private mconIntColָ�������� As Integer
Private mconIntCol���� As Integer
Private mconIntCol�ɱ��� As Integer
Private mconIntCol�ɱ���� As Integer
Private mconIntCol�ۼ� As Integer
Private mconIntCol�ۼ۽�� As Integer
Private mconintCol��� As Integer
Private mconintCol���ۼ� As Integer
Private mconintCol���۵�λ As Integer
Private mconintCol���۽�� As Integer
Private mconintCol���۲�� As Integer
Private mconIntCol��׼�ĺ� As Integer
Private mconIntCol��� As Integer
Private mconIntCol���ս��� As Integer
Private mconintcol��Ʒ�ϸ�֤ As Integer
Private mconintcol������� As Integer
Private mconintcol��Ʊ�� As Integer
Private mconintcol��Ʊ���� As Integer
Private mconIntCol��Ʊ���� As Integer
Private mconintcol��Ʊ��� As Integer
Private mconIntCol�ɹ��� As Integer
Private mconIntCol�������� As Integer
Private mconIntCol�Ƿ����� As Integer
Private mconIntcol�ӳ��� As Integer
Private mconIntColҩƷ��������� As Integer
Private mconIntColҩƷ���� As Integer
Private mconIntColҩƷ���� As Integer
Private mconIntCol�����־ As Integer
Private mconIntCol�ƻ�id As Integer
Private mconintcol������� As Integer
Private Const mconIntColS As Integer = 52
'=========================================================================================



Private Function CheckQualifications(ByVal strInput As String) As Boolean
    'У�鹩Ӧ����Ϣ������Ч��
    'strInput���ַ���ʱΪ���ƣ�����ʱΪID
    Dim rsTmp As ADODB.Recordset
    Dim strMsgInfo As String
    Dim strMsgDate As String
    Dim dateCurrent As Date
    Dim strMsg As String
    
    Dim intCheckType As Integer
    Dim arrColumn
    Dim strCheck As String
    Dim strCheck_��Ӧ�� As String
    Dim n As Integer
    Dim strTmp As String
    
    On Error GoTo errHandle
    If strInput = "" Then
        CheckQualifications = True
        Exit Function
    End If
        
    '����У����Ŀ�ͷ�ʽ�ı����ʽ��У�鷽ʽ|���1,��Ŀ1,�Ƿ�У��;���1,��Ŀ2,�Ƿ�У��;���2,��Ŀ1,�Ƿ�У��;���2,��Ŀ2....
    strCheck = zlDataBase.GetPara("����У��", glngSys, 1300, "")
    
    '����Ĳ�����ʽ����ȷʱ�˳�
    If InStr(1, strCheck, "|") = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    'ȡУ�鷽ʽ��0-����飻1�����ѣ�2����ֹ
    intCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))
    
    '�����ʱ�˳�
    If intCheckType = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    'ȡУ�����ݣ�
    strCheck = Mid(strCheck, InStr(1, strCheck, "|") + 1)

    If strCheck = "" Then
        CheckQualifications = True
        Exit Function
    End If

    '�ֱ�ȡ���ģ������̣���Ӧ����ҪУ�������
    strCheck = strCheck & ";"
    arrColumn = Split(strCheck, ";")
    For n = 0 To UBound(arrColumn)
        If arrColumn(n) <> "" Then
'            If Split(arrColumn(n), ",")(0) = "����" And Split(arrColumn(n), ",")(2) = 1 Then
'                strCheck_���� = IIf(strCheck_���� = "", "", strCheck_���� & ";") & Split(arrColumn(n), ",")(1)
'            End If
'
'            If Split(arrColumn(n), ",")(0) = "����������" And Split(arrColumn(n), ",")(2) = 1 Then
'                strCheck_������ = IIf(strCheck_������ = "", "", strCheck_������ & ";") & Split(arrColumn(n), ",")(1)
'            End If

            If Split(arrColumn(n), ",")(0) = "ҩƷ��Ӧ��" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_��Ӧ�� = IIf(strCheck_��Ӧ�� = "", "", strCheck_��Ӧ�� & ";") & Split(arrColumn(n), ",")(1)
            End If
        End If
    Next
    
    '��У������ʱ�˳�
    If strCheck_��Ӧ�� = "" Then
        CheckQualifications = True
        Exit Function
    End If
    
    dateCurrent = CDate(Format(Sys.Currentdate, "yyyy-mm-dd"))
    
    gstrSQL = "Select ('[' || ���� || ']' || ����) AS ��Ӧ��, ˰��ǼǺ�, ���֤��, ִ�պ�, ��Ȩ��, ������֤��, ������֤����, ҩ��ֱ�����, ҩ��ֱ�������, ���֤Ч��, ִ��Ч��, ��Ȩ�� " & _
              "From ��Ӧ�� " & _
              "Where (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And ID = [1] "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "��Ӧ����Ϣ", Val(strInput))
    
    strTmp = ""
    
    If Not rsTmp.EOF Then
        If nvl(rsTmp!˰��ǼǺ�) = "" And InStr(strCheck_��Ӧ��, "˰��ǼǺ�") > 0 Then
            strTmp = rsTmp!��Ӧ�� & "��" & "��˰��ǼǺ�"
        End If
        
        If nvl(rsTmp!���֤��) = "" And InStr(strCheck_��Ӧ��, "���֤��") > 0 Then
            strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "�����֤��"
        End If
        
        If nvl(rsTmp!ִ�պ�) = "" And InStr(strCheck_��Ӧ��, "ִ�պ�") > 0 Then
            strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��ִ�պ�"
        End If
        
        If nvl(rsTmp!��Ȩ��) = "" And InStr(strCheck_��Ӧ��, "��Ȩ��") > 0 Then
            strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "����Ȩ��"
        End If
        
        If nvl(rsTmp!������֤��) = "" And InStr(strCheck_��Ӧ��, "������֤��") > 0 Then
            strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��������֤��"
        End If
        
        If nvl(rsTmp!������֤����) <> "" Then
            If DateDiff("d", rsTmp!������֤����, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "������֤����") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "������֤���ѹ���"
            End If
        End If
        
        If nvl(rsTmp!ҩ��ֱ�����) = "" And InStr(strCheck_��Ӧ��, "ҩ��ֱ�����") > 0 Then
            strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��ҩ��ֱ�����"
        End If
        
        If nvl(rsTmp!ҩ��ֱ�������) <> "" Then
            If DateDiff("d", rsTmp!ҩ��ֱ�������, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "ҩ��ֱ�������") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "ҩ��ֱ������ѹ���"
            End If
        End If
        
        If nvl(rsTmp!���֤Ч��) <> "" Then
            If DateDiff("d", rsTmp!���֤Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "���֤Ч��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "���֤�ѹ���"
            End If
        End If
        
        If nvl(rsTmp!ִ��Ч��) <> "" Then
            If DateDiff("d", rsTmp!ִ��Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "ִ��Ч��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "ִ���ѹ���"
            End If
        End If
        
        If nvl(rsTmp!��Ȩ��) <> "" Then
            If DateDiff("d", rsTmp!ִ��Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "��Ȩ��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��Ȩ�ѹ���"
            End If
        End If
    End If
    
    '��ʾ���ֹ
    If strTmp <> "" Then
        If intCheckType = 1 Then
            If MsgBox("δͨ������У�飬�Ƿ������" & vbCrLf & strTmp, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                CheckQualifications = True
                Exit Function
            Else
                Exit Function
            End If
        ElseIf intCheckType = 2 Then
            MsgBox "δͨ������У�飬������⣡" & vbCrLf & strTmp, vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    CheckQualifications = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check��Ա����(lngUserId As Long, lngDeptId As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From ������Ա Where ��Աid=[1] And ����id=[2] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�ж��û��Ƿ���������ⲿ��]", lngUserId, lngDeptId)
    Check��Ա���� = (rsTemp.RecordCount > 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check��ͬ��λ() As Boolean
    Dim n As Integer
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errHandle
    For n = 1 To mshBill.rows - 1
        If Trim(mshBill.TextMatrix(n, 0)) <> "" Then
            gstrSQL = "select nvl(��ͬ��λid,0) ��ͬ��λid from ҩƷ��� where ҩƷid=[1] "
            Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�ж��Ƿ��Ǵ��ں�ͬ��λ]", Val(mshBill.TextMatrix(n, 0)))
            
            If Not rs.EOF Then
                gstrSQL = "select id,���� from ��Ӧ�� " & _
                          "where (վ�� = [2] Or վ�� is Null) And ĩ��=1 And substr(����,1,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
                          "  and id=(select nvl(��ͬ��λid,0) id from ҩƷ��� where ҩƷid=[1]) "
                Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�ж��Ƿ����ں�ͬ��λ���ɹ�]", Val(mshBill.TextMatrix(n, 0)), gstrNodeNo)
                
                If Not rs.EOF Then
                    If rs!id <> txtProvider.Tag Then
                        strTmp = strTmp & mshBill.TextMatrix(n, mconIntColҩ��) & "[" & rs!���� & "]" & vbCrLf
                    End If
                End If
            End If
        End If
    Next
    
    If strTmp <> "" Then
        MsgBox "�ù�ҩ��λ��������ҩƷ�ĺ�ͬ��λ��" & vbCrLf & strTmp, vbInformation, gstrSysName
        Exit Function
    End If
    
    Check��ͬ��λ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check�洢�ⷿ() As Boolean
    Dim n As Integer
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
        
    mbln����ҩƷ�޴洢�ⷿ = True
    On Error GoTo errHandle
    For n = 1 To mshBill.rows - 1
        If Trim(mshBill.TextMatrix(n, 0)) <> "" Then
            gstrSQL = "select �շ�ϸĿID from �շ�ִ�п��� where �շ�ϸĿID=[1] and ִ�п���ID=[2]  "
            Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�ж�ҩƷ�洢�ⷿ]", Val(mshBill.TextMatrix(n, 0)), cboEnterStock.ItemData(cboEnterStock.ListIndex))
            
            If rs.RecordCount = 0 Then
                 strTmp = strTmp & mshBill.TextMatrix(n, mconIntColҩ��) & vbCrLf
            Else
                mbln����ҩƷ�޴洢�ⷿ = False
            End If
        End If
    Next
    
    If strTmp <> "" Then
        If mbln����ҩƷ�޴洢�ⷿ Then
            mstr�洢�ⷿ��ʾ = "�������ҩƷû�����ô洢�ⷿ���������ƿ⵽[" & cboEnterStock.Text & "]��"
        Else
            mstr�洢�ⷿ��ʾ = "����ҩƷû�����ô洢�ⷿ���������ƿ⵽[" & cboEnterStock.Text & "] ��" & vbCrLf & strTmp & vbCrLf & "����ҩƷ���Ե����ƿ⡣"
        End If
        Check�洢�ⷿ = False
    Else
        Check�洢�ⷿ = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetNextEnableCol(ByVal intCurrCol As Integer) As Integer
    '������һ���ɼ������õ��к�
    Dim n As Integer
    Dim intNextCol As Integer
    
    If mshBill.TextMatrix(mshBill.Row, 0) <> "" Then
        If intCurrCol > mshBill.Cols Or intCurrCol + 1 >= mintLastCol Then
            If mshBill.Row = mshBill.rows - 1 Then
                mshBill.rows = mshBill.rows + 1
            End If
            
            mshBill.Row = mshBill.Row + 1
            GetNextEnableCol = 2
            Exit Function
        End If
        
        With mshBill
            For n = intCurrCol + 1 To .Cols - 1
                If .ColWidth(n) > 0 And .ColData(n) <> 5 Then
                    intNextCol = n
                    Exit For
                End If
            Next
        End With
        
        GetNextEnableCol = IIf(intNextCol = 0, mintLastCol, intNextCol)
    End If
End Function
Private Sub GetSysParm()
    Dim int���� As Integer
    
    'mint�༭״̬��1.������2���޸ģ�3�����գ�4���鿴��5���޸ķ�Ʊ��6��������
    '7��������ˣ������������µ��ݲ���ˣ��Ѹ���ĵ��ݲ������
    '����ˣ�ͬ����������˺�ĵ��ݲ����������;8-ҩ���˻�;9-�˲�

    If mint�༭״̬ = 9 Then
        int���� = ����.�˲�
    ElseIf mint�༭״̬ = 3 Then
        int���� = ����.���
    ElseIf mint�༭״̬ = 7 Then
        int���� = ����.�������
    End If
    
    If int���� > 0 Then
        mstrControlItem = "," & GetControlItem(1, int����) & ","
    End If
End Sub

Private Sub GetҩƷ��������(ByVal intBillRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim int�������� As Integer      '0-������;1-����
    Dim intҩ����� As Integer      '0-������;1-����
    Dim intҩ������ As Integer      '0-������;1-����
    Dim bln�Ƿ����ҩ������ As Boolean  'True-����ҩ������;False-������ҩ������
    
    If Val(mshBill.TextMatrix(intBillRow, 0)) = 0 Then Exit Sub
    On Error GoTo errHandle
    strSQL = "SELECT NVL(ҩ�����, 0) ҩ�����,NVL(ҩ������, 0) ҩ������ " & _
            " From ҩƷ��� WHERE ҩƷID = [1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "ȡҩƷ�ⷿ��������", Val(mshBill.TextMatrix(intBillRow, 0)))
    
    If rsTemp.RecordCount > 0 Then
        intҩ����� = rsTemp!ҩ�����
        intҩ������ = rsTemp!ҩ������
    End If
    
    If intҩ������ = 1 Then     '���ҩ�����������������Ϊ1
        int�������� = 1
    Else
        If intҩ����� = 1 Then
            strSQL = "SELECT ����ID From ��������˵�� " & _
                    " WHERE ((�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���')) AND ����ID = [1] "
            Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "ȡ��������", cboStock.ItemData(Me.cboStock.ListIndex))
            
            bln�Ƿ����ҩ������ = (rsTemp.RecordCount > 0)
                    
            If bln�Ƿ����ҩ������ Then
                int�������� = 0
            Else
                int�������� = 1
            End If
        End If
    End If
    
    mshBill.TextMatrix(intBillRow, mconIntCol��������) = int��������
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetժҪ(ByVal strNo As String, ByVal int�༭״̬ As Integer) As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    Select Case int�༭״̬
        Case 6          '����(ȡ���һ�γ�����ժҪ)
            gstrSQL = "Select ժҪ From ҩƷ�շ���¼ Where ����=1 And No=[1] Order By ������� Desc "
        Case 5, 7       '�޸ķ�Ʊ���������
            gstrSQL = "Select ժҪ From ҩƷ�շ���¼ Where ���� = 1 And NO = [1] And (Mod(��¼״̬, 3) = 0 Or ��¼״̬ = 1) "
    End Select
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡժҪ��Ϣ", strNo)
    
    If Not rsTemp.EOF Then
        GetժҪ = nvl(rsTemp!ժҪ)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub Refresh�����־()
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    
    On Error GoTo errHandle
    For n = 1 To mshBill.rows - 1
        If mshBill.TextMatrix(n, mconIntCol���) <> "" Then
            gstrSQL = "Select Nvl(Max(�������), 0) ������� From Ӧ����¼ " & _
                " where �շ�id=(Select Id From ҩƷ�շ���¼ Where ����=1 And No=[1] And (Mod(��¼״̬,3)=0 Or ��¼״̬=1) " & _
                " And ���=[2]) "
            Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ�������]", txtNO.Text, Val(mshBill.TextMatrix(n, mconIntCol���)))
            
            If rs.EOF Then
                mshBill.RowData(n) = 0
            Else
                mshBill.RowData(n) = rs!�������
            End If
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub


'�������������
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    GetDepend = False
    strSQL = "SELECT B.Id " _
            & "FROM ҩƷ�������� A, ҩƷ������ B " _
            & "Where A.���id = B.ID AND A.���� = 1 and rownum=1 "
    Set rsDepend = zlDataBase.OpenSQLRecord(strSQL, MStrCaption)
    If rsDepend.EOF Then
        MsgBox "û������ҩƷ�⹺��������������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
        
    strSQL = "Select (id) from ��Ӧ�� " & _
              "Where (վ�� = [1] Or վ�� is Null) And (����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or ����ʱ�� is null) " & _
              "  And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) and rownum=1 "
    Set rsDepend = zlDataBase.OpenSQLRecord(strSQL, MStrCaption & "-��Ӧ��", gstrNodeNo)
    If rsDepend.EOF Then
        MsgBox "û������ҩƷ��ҩ��λ������ҩƷ��ҩ��λ����", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
        
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function


Private Sub SetDrugName(ByVal intType As Integer)
    'ҩƷ������ʾ��
    'intType��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
    Dim lngRow As Long
    
    If intType = mintDrugNameShow Then Exit Sub
    
    mintDrugNameShow = intType
    
    With mshBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntColҩ��) <> "" Then
                If mintDrugNameShow = 1 Then
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ����)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ����)
                Else
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ���������)
                End If
            End If
        Next
    End With
End Sub

Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.rows < 2 Then Exit Sub
    If mshBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "�к�", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !�к� = n
                !��� = IIf(Val(mshBill.TextMatrix(n, mconIntCol���)) = 0, n, Val(mshBill.TextMatrix(n, mconIntCol���)))
                !ҩƷID = Val(mshBill.TextMatrix(n, 0))
                !���� = Val(mshBill.TextMatrix(n, mconIntCol����))
                
                .Update
            End If
        Next
        
    End With
End Sub

Private Sub Setʱ�۷���ҩƷ���ۼ�(ByVal intRow As Integer, ByVal dblPrice As Double)
    Dim Dbl���� As Double

    With mshBill
        If .TextMatrix(intRow, mconIntColԭ����) = "" Then Exit Sub
        If Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) <> 1 Or Val(.TextMatrix(intRow, mconIntCol��������)) <> 1 Then Exit Sub
        
       .TextMatrix(intRow, mconintCol���ۼ�) = zlStr.FormatEx(dblPrice, gtype_UserDrugDigits.Digit_���ۼ�, , True) '���ۼ��ֶα���������С��λ����˰������λ��������ʾ
        
        If mint�༭״̬ = 6 Then
            Dbl���� = Val(.TextMatrix(intRow, mconIntCol��������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��))
        Else
            Dbl���� = Val(.TextMatrix(intRow, mconIntCol����)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��))
        End If
        
        If Val(.TextMatrix(intRow, mconIntCol�ɱ���)) = Val(.TextMatrix(intRow, mconIntCol�ۼ�)) Then
            'ͨ�������ֶε�����������������������ۼۺ��ۼ۲��ȵ����
            .TextMatrix(intRow, mconintCol���۽��) = .TextMatrix(intRow, mconIntCol�ۼ۽��)
        Else
            .TextMatrix(intRow, mconintCol���۽��) = zlStr.FormatEx(Dbl���� * Val(.TextMatrix(intRow, mconintCol���ۼ�)), mintMoneyDigit, , True)
        End If
        .TextMatrix(intRow, mconintCol���۲��) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol���۽��)) - Val(.TextMatrix(intRow, mconIntCol�ɱ����)), mintMoneyDigit, , True)
    End With
End Sub

Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, Optional int��¼״̬ As Integer = 1, Optional BlnSuccess As Boolean = False)
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    
    mint��¼״̬ = int��¼״̬
    mblnSuccess = BlnSuccess
    mblnChange = False
    mblnЧ����ʾ = False
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1300)

    mbln�޸������� = (Val(zlDataBase.GetPara("�޸Ĳɹ��޼�", glngSys, ģ���.�⹺���)) = 1)
    mbln�б�ҩƷ��ѡ����б굥λ��� = (Val(zlDataBase.GetPara("�б�ҩƷ��ѡ����б굥λ���", glngSys, ģ���.�⹺���)) = 1)
    mintȡ�ϴβɹ��۷�ʽ = Val(zlDataBase.GetPara("ȡ�ϴβɹ��۷�ʽ", glngSys, ģ���.�⹺���))
    mbln��Ӧ��У�� = (Val(zlDataBase.GetPara("У�鹩Ӧ������", glngSys, ģ���.�⹺���)) = 1)
    mintʱ�۷ֶμӳɷ�ʽ = gtype_UserSysParms.P181_ҩƷ��ⰴ�ֶμӳ�
    Set mfrmMain = FrmMain
    mblnEdit = False
    If mint�༭״̬ = 1 Or mint�༭״̬ = 8 Then
        mblnEdit = True

        txtNO.Locked = True
        txtNO.TabStop = True
        If Not GetDepend Then Exit Sub
    ElseIf mint�༭״̬ = 2 Or mint�༭״̬ = 7 Then
        mblnEdit = True
        If mint�༭״̬ = 2 Then
            txtNO.Locked = True
            txtNO.TabStop = True
        End If
    ElseIf mint�༭״̬ = 3 Then
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
        
        If Not mbln�˻� Then
            Me.chkת���ƿ�.Visible = True
            Me.cboEnterStock.Visible = True
        End If
    ElseIf mint�༭״̬ = 4 Then
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If Not zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint�༭״̬ = 5 Then
        mblnEdit = False
        
    ElseIf mint�༭״̬ = 6 Then
        mblnEdit = False
        CmdSave.Caption = "����(&O)"
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
    End If
    
    
    LblTitle.Caption = GetUnitName & IIf(mint�༭״̬ = 8, "ҩƷ�˻���", LblTitle.Caption)
    mblnЧ����ʾ = True
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
End Sub

Private Sub zlPrintBill_Check()
    Dim lng�ϴ�ҩƷID As Long
    
    If Val(zlDataBase.GetPara("��˴�ӡ", glngSys, ģ���.�⹺���)) = 1 Then
        '��ӡ
        If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
            printbill
            
            If Val(zlDataBase.GetPara("��ӡҩƷ����", glngSys, ģ���.�⹺���)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
                '��ҩƷID˳���������
                recSort.Sort = "ҩƷid"
                recSort.MoveFirst
                '��ӡҩƷ����
                Do While Not recSort.EOF
                    If lng�ϴ�ҩƷID <> Val(recSort!ҩƷID) Then
                        ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1300_1", Me, "ҩƷ=" & Val(recSort!ҩƷID), 2
                        lng�ϴ�ҩƷID = recSort!ҩƷID
                    End If
                    recSort.MoveNext
                Loop
            End If
                
        End If
    End If
End Sub

Private Sub cboEnterStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If cboEnterStock.ListIndex >= 0 Then
        If Val(cboEnterStock.Tag) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
            Exit Sub
        End If
    End If

    Dim rsEnterDept As ADODB.Recordset
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim i As Long
        
    vRect = zlControl.GetControlRect(cboEnterStock.hWnd)
    gstrSQL = "Select Distinct a.Id, a.����, a.���� " & vbNewLine & _
              "From ��������˵�� C, �������ʷ��� B, ���ű� A, " & vbNewLine & _
              "     (Select �Է��ⷿid ID " & vbNewLine & _
              "       From ҩƷ������� " & vbNewLine & _
              "       Where ���ڿⷿid = [1] And ���� In (1, 3) " & vbNewLine & _
              "       Union " & vbNewLine & _
              "       Select ���ڿⷿid ID From ҩƷ������� Where �Է��ⷿid = [1] And ���� In (2, 3)) D " & vbNewLine & _
              "Where c.�������� = b.���� And b.���� || '' In ('H', 'I', 'J', 'K', 'L', 'M', 'N') " & vbNewLine & _
              "    And a.Id = c.����id And a.Id = d.Id " & vbNewLine & _
              "    And To_Char(a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' " & vbNewLine & _
              "    And (A.���� like [2] or A.���� like [2] or A.���� like [2] ) " & vbNewLine & _
              "Order By a.���� "
    Set rsEnterDept = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, MStrCaption, False, "", "", _
            False, False, True, vRect.Left - 15, vRect.Top, 3000, blnCancel, False, False, _
            cboStock.ItemData(cboStock.ListIndex), _
            IIf(gstrMatchMethod = "0", "%", "") & UCase(Trim(cboEnterStock.Text)) & "%")
    If blnCancel = False Then
        If rsEnterDept Is Nothing Then Exit Sub
        If Not rsEnterDept.EOF Then
            For i = 0 To cboEnterStock.ListCount - 1
                If cboEnterStock.ItemData(i) = nvl(rsEnterDept!id, -1) Then
                    cboEnterStock.ListIndex = i
                    cboEnterStock.Tag = nvl(rsEnterDept!id, 0)
                    Exit For
                End If
            Next
        End If
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_Click()
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    Dim str������ As String
    
    On Error GoTo errHandle
    
    str�ⷿ���� = ""
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        gstrSQL = "Select �������� From ��������˵�� Where ����id =[1]"
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "�жϿⷿ����", cboStock.ItemData(cboStock.ListIndex))
        Do While Not rsDetail.EOF
            str�ⷿ���� = str�ⷿ���� & "," & rsDetail!��������
            rsDetail.MoveNext
        Loop
        If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True

        str������ = zlDataBase.GetPara("������", glngSys, ģ���.�⹺���)
        
        If InStr(1, "|" & str������ & "|", "|ԭ����|") = 0 Then mshBill.ColWidth(mconIntColԭ����) = IIf(bln��ҩ�ⷿ, 800, 0)
        
        If mblnLoad = True Then Call SetSelectorRS(IIf(mint�༭״̬ = 8 Or mbln�˻�, 2, 1), MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , IIf(mint�༭״̬ = 8 Or mbln�˻�, Val(txtProvider.Tag), 0))
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboStock_Validate False
        OS.PressKey (vbKeyTab)
    End If
End Sub


Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
        
    With cboStock
        If .ListIndex <> mintcboIndex Then
            For i = 1 To mshBill.rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.rows Then
                If MsgBox("����ı�ⷿ���п���Ҫ�ı���ӦҩƷ�ĵ�λ����Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '����ҩƷ��λ�ı�
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                    mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
                    
                    mlng���ⷿ = Me.cboStock.ItemData(Me.cboStock.ListIndex)
                    Call GetDrugDigit(mlng���ⷿ, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        
        
    End With
End Sub

Private Sub chkת���ƿ�_Click()
    If chkת���ƿ�.Value = 1 Then
        cboEnterStock.Enabled = True
    Else
        cboEnterStock.Enabled = False
    End If
End Sub

Private Sub cmdAddProducer_Click()
    frmDrugProducer.Show 1, Me
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(0, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol�ɱ����) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                If Trim(.TextMatrix(intRow, mconintcol��Ʊ��)) <> "" Then
                    .TextMatrix(intRow, mconintcol��Ʊ���) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                End If
                
                Call Setʱ�۷���ҩƷ���ۼ�(intRow, Val(.TextMatrix(intRow, mconintCol���ۼ�)))
            End If
        Next
    End With
    mblnChange = False
End Sub

Private Sub cmdAllSel_Click()
    Dim rsDrug As New Recordset
    Dim intRow As Integer
    
    On Error GoTo errHandle
    For intRow = 1 To mshBill.rows - 1
        If (gtype_UserSysParms.P173_������Ǹ������ܽ��и������ = 1 And mshBill.TextMatrix(intRow, mconIntCol�����־) <> "��") Or gtype_UserSysParms.P173_������Ǹ������ܽ��и������ <> 1 Then
            If mshBill.TextMatrix(intRow, 0) <> "" And mshBill.RowData(intRow) = 0 Then
                mshBill.TextMatrix(intRow, mconIntCol��������) = mshBill.TextMatrix(intRow, mconIntCol����)
                mshBill.TextMatrix(intRow, mconIntCol�ɱ����) = zlStr.FormatEx(mshBill.TextMatrix(intRow, mconIntCol����) * mshBill.TextMatrix(intRow, mconIntCol�ɱ���), mintMoneyDigit, , True)
                mshBill.TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(mshBill.TextMatrix(intRow, mconIntCol����) * mshBill.TextMatrix(intRow, mconIntCol�ۼ�), mintMoneyDigit, , True)
                mshBill.TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(mshBill.TextMatrix(intRow, mconIntCol�ۼ۽��) - mshBill.TextMatrix(intRow, mconIntCol�ɱ����), mintMoneyDigit, , True)
                  
                If Trim(mshBill.TextMatrix(intRow, mconintcol��Ʊ��)) <> "" Then
                    gstrSQL = "select sum(nvl(��Ʊ���,0)) as ��Ʊ��� " _
                        & " From ҩƷ�շ���¼ x,(Select �շ�id,��Ʊ��� From Ӧ����¼ Where ϵͳ��ʶ=1 And ��¼����=0) y " _
                        & " WHERE x.id=y.�շ�id(+) and X.NO=[1] AND ����=1 " _
                        & " and x.ҩƷid+0=[2] " _
                        & " and x.���=[3] "
                    Set rsDrug = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�, Val(mshBill.TextMatrix(intRow, 0)), Val(mshBill.TextMatrix(intRow, mconIntCol���)))
                    
                    If rsDrug.EOF Then
                        mshBill.TextMatrix(intRow, mconintcol��Ʊ���) = mshBill.TextMatrix(intRow, mconIntCol�ɱ����)
                    Else
                        mshBill.TextMatrix(intRow, mconintcol��Ʊ���) = zlStr.FormatEx(rsDrug.Fields(0), mintMoneyDigit, , True)
                    End If
                End If
                
                Call Setʱ�۷���ҩƷ���ۼ�(intRow, Val(mshBill.TextMatrix(intRow, mconintCol���ۼ�)))
            End If
        End If
    Next
    mblnChange = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    Dim i As Integer
    
    With mshBill
        '1�����з�Ʊ��������2��
        For i = 1 To .rows - 1
           If Trim(.TextMatrix(i, mconintcol��Ʊ��)) = "" Or .TextMatrix(i, 0) = "" Then Exit For
        Next
        
        If i = .rows - 1 Then Exit Sub
        
        '2����Ʊ�����Ʊ����Ϊ�գ�����ʾ
        If Trim(.TextMatrix(.Row, mconintcol��Ʊ����)) = "" Or .TextMatrix(.Row, mconIntCol��Ʊ����) = "" Then
            If MsgBox("��Ʊ�����Ʊ����Ϊ�գ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("�Ƿ񽫸��еķ�Ʊ��Ϣ�������Ƶ���Ʊ��Ϊ�յ��У�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        '3������
        For i = 1 To .rows - 1
            If i <> .Row And Trim(.TextMatrix(i, mconintcol��Ʊ��)) = "" And .TextMatrix(i, 0) <> "" Then    '���Ǳ༭���ҷ�Ʊ��Ϊ�յ������޸�
                
                .TextMatrix(i, mconintcol��Ʊ��) = .TextMatrix(.Row, mconintcol��Ʊ��)
                .TextMatrix(i, mconintcol��Ʊ����) = .TextMatrix(.Row, mconintcol��Ʊ����)
                .TextMatrix(i, mconIntCol��Ʊ����) = .TextMatrix(.Row, mconIntCol��Ʊ����)
                If mint��¼״̬ = 1 Then .TextMatrix(i, mconintcol��Ʊ���) = .TextMatrix(i, mconIntCol�ɱ����)
                
            End If
        Next
    End With
End Sub

Private Sub cmdALLDel_Click()
    Dim i As Integer
    
    With mshBill
        '1�����޷�Ʊ��������2��
        For i = 1 To .rows - 1
           If Trim(.TextMatrix(i, mconintcol��Ʊ��)) <> "" Or .TextMatrix(i, 0) = "" Then Exit For
        Next
        
        If i = .rows - 1 Then Exit Sub
    
        If MsgBox("�ò�������������еķ�Ʊ������ݣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            For i = 1 To .rows - 1
            
                If Trim(.TextMatrix(i, mconintcol��Ʊ��)) <> "" And .TextMatrix(i, 0) <> "" Then
                    .TextMatrix(i, mconintcol��Ʊ��) = ""
                    .TextMatrix(i, mconintcol��Ʊ����) = ""
                    .TextMatrix(i, mconIntCol��Ʊ����) = ""
                    .TextMatrix(i, mconintcol��Ʊ���) = ""
                End If
                
            Next
            
            cmdCopy.Enabled = False
        End If
    End With
End Sub

'����
Private Sub cmdFind_Click()
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRow mshBill, mconIntColҩƷ���������, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
        
    End If
    
    Form_Resize
End Sub

Private Sub cmdGetData_Click()
    If cboInputDate.Text <> "��������" Then
        If MsgBox("��ѯʱ��������ܺ������Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call getInputData
        End If
    End If
End Sub

Private Sub cmdGetInputCost_Click()
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim dblRowStation As Double
    Dim lngCurentRow As Long
    Dim lngҩƷID As Long
    Dim dbl����ϵ�� As Double
    
    dblLeft = mshBill.Left + mshBill.MsfObj.CellLeft
    dblTop = mshBill.Top + mshBill.MsfObj.CellTop
    'ͨ���ؼ��߶Ȼ�ȡ��λ��
    dblRowStation = mshBill.MsfObj.CellTop
    dblRowStation = dblRowStation / mshBill.MsfObj.CellHeight
    lngCurentRow = CLng(dblRowStation) 'Clng��֤ȡ������Ϊ����
    
    If mshBill.TextMatrix(lngCurentRow, 0) <> "" Then
        cboInputDate.Clear
        '��ʼ�������б�
        cboInputDate.AddItem "��������"
        cboInputDate.AddItem "������"
        cboInputDate.AddItem "һ����"
        cboInputDate.ListIndex = 0
        
        picInputCost.Visible = True
        vsfInputCost.SetFocus
        picInputCost.Top = dblTop
        picInputCost.Left = dblLeft
        
        lngҩƷID = mshBill.TextMatrix(lngCurentRow, 0)
        dbl����ϵ�� = mshBill.TextMatrix(lngCurentRow, mconIntCol����ϵ��)
        picInputCost.Tag = lngҩƷID
        cmdGetData.Tag = dbl����ϵ��
        lblDate.Tag = mintCostDigit
        vsfInputCost.Tag = lngCurentRow
        
        Call getInputData
    End If
End Sub

Private Sub getInputData()
    Dim dbBeginDate As Date
    Dim dbEndDate As Date
    Dim rsTemp As ADODB.Recordset
    
    If cboInputDate.Text = "��������" Then
        dbBeginDate = CDate(Format(DateAdd("M", -3, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "������" Then
        dbBeginDate = CDate(Format(DateAdd("M", -6, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "һ����" Then
        dbBeginDate = CDate(Format(DateAdd("yyyy", -1, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    
    gstrSQL = "Select a.No, a.�������, a.�ɱ���, b.��Ʊ��, b.��Ʊ����" & vbNewLine & _
                "From ҩƷ�շ���¼ A, Ӧ����¼ B" & vbNewLine & _
                "Where a.Id = b.�շ�id(+) And a.���� = 1 And a.ҩƷid + 0 = [1] And a.�ⷿid+0=[2]  And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0) And" & vbNewLine & _
                "      Nvl(a.����id, 0) = 0 And nvl(a.��ҩ��ʽ,0)=0 And a.������� Between [3] And [4]" & vbNewLine & _
                "Order By a.������� Desc"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�����Ϣ��ѯ", picInputCost.Tag, cboStock.ItemData(cboStock.ListIndex), dbBeginDate, dbEndDate)
    vsfInputCost.rows = 1
    Do While Not rsTemp.EOF
        With vsfInputCost
            .rows = .rows + 1
            .TextMatrix(.rows - 1, .ColIndex("no")) = rsTemp!NO
            .TextMatrix(.rows - 1, .ColIndex("���ʱ��")) = Format(rsTemp!�������, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.rows - 1, .ColIndex("�ɱ���")) = zlStr.FormatEx(rsTemp!�ɱ��� * cmdGetData.Tag, lblDate.Tag, , True)
            .TextMatrix(.rows - 1, .ColIndex("��Ʊ��")) = IIf(IsNull(rsTemp!��Ʊ��), "", rsTemp!��Ʊ��)
            .TextMatrix(.rows - 1, .ColIndex("��Ʊ����")) = IIf(IsNull(rsTemp!��Ʊ����), "", rsTemp!��Ʊ����)
            
            rsTemp.MoveNext
        End With
    Loop
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdNo_Click()
    Dim mdbl�Ӽ��� As Double
    Dim dblTemp�ۼ� As Double
    
    With mshBill
        mdbl�Ӽ��� = Val(Txt�Ӽ���.Tag)
                
        '���¼������ۼۡ����
        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then
            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ɱ���)), mdbl�Ӽ��� / 100, Val(.TextMatrix(.Row, mconIntCol�ɱ���)) * (1 + (mdbl�Ӽ��� / 100))), mintPriceDigit, , True)
        End If
        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ�)) * Val(.TextMatrix(.Row, mconIntCol����)), mintMoneyDigit, , True)
        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
    
        Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / Val(.TextMatrix(.Row, mconIntCol����ϵ��)))
    End With
    PicInput.Visible = False
End Sub

Private Sub CmdNO_LostFocus()
    Call PicInput_LostFocus
End Sub

Private Sub cmdYes_Click()
    Dim dbl�ɱ��� As Double
    
    If Val(Txt�Ӽ���) > 9900 Or Val(Txt�Ӽ���) < 0 Then
        MsgBox "������Ϸ��ļӳ��ʣ���0-9900��", vbInformation, gstrSysName
        Txt�Ӽ���.SetFocus
        Exit Sub
    End If
    
    With mshBill
        '���ݲ�������ʱ��ҩƷ�ۼ۹�ʽ�гɱ��۵��㷨
        dbl�ɱ��� = IIf(mintʱ������ۼۼӳɷ�ʽ = 0, Val(.TextMatrix(.Row, mconIntCol�ɱ���)), Val(.TextMatrix(.Row, mconIntCol�ɹ���)))
                        
        '���¼������ۼۡ����
        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then
            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), dbl�ɱ���, Val(Txt�Ӽ���) / 100, dbl�ɱ��� * (1 + (Val(Txt�Ӽ���) / 100))), mintPriceDigit, , True)
        End If
        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ�)) * Val(.TextMatrix(.Row, mconIntCol����)), mintMoneyDigit, , True)
        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
        .TextMatrix(.Row, mconIntcol�ӳ���) = zlStr.FormatEx(Val(Txt�Ӽ���), 2) & "%"
        
        Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / Val(.TextMatrix(.Row, mconIntCol����ϵ��)))
    End With
    
    PicInput.Visible = False
    mshBill.SetFocus
End Sub
Private Sub CmdYes_LostFocus()
    Call PicInput_LostFocus
End Sub

Private Sub Form_Activate()
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            '�����ѱ�ɾ��
            If mint�༭״̬ = 5 Then
                MsgBox "�õ����ѱ������������޸ķ�Ʊ��Ϣ�����飡", vbOKOnly, gstrSysName
            ElseIf mint�༭״̬ = 6 Then
                MsgBox "�õ�����û�п��Գ�����ҩƷ�����飡", vbOKOnly, gstrSysName
            Else
                MsgBox "�õ����ѱ�ɾ�������飡", vbOKOnly, gstrSysName
            End If
            Unload Me
            Exit Sub
        Case 3
            '�޸ĵĵ����ѱ����
            MsgBox "�õ����ѱ���������ˣ����飡", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 4
            MsgBox "�õ����ѱ������˸���������޸ķ�Ʊ��Ϣ�����飡", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 5
            MsgBox "�õ����ѱ������˸�������ܽ��в�����ˣ�", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 6
            MsgBox "�㲻��[" & cboStock.Text & "]��Ա�����ܽ��в�����ˡ�", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 7
            MsgBox "�õ�����ȫ�����߲��ָ�����ܳ�����", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
    End Select
    '��ʼ�����뷽ʽ
    If (mint�༭״̬ = 1 Or mint�༭״̬ = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint���뷽ʽ = Val(zlDataBase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram staThis, gint���뷽ʽ
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single, sngTop As Single
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRow mshBill, mconIntColҩ��, txtCode.Text, False
    ElseIf KeyCode = vbKeyF4 Then
        '���ϵͳ����Ϊ�棬����ʾ�û�����Ӽ���
        If mbln�Ӽ��� And (mint�༭״̬ = 1 Or mint�༭״̬ = 2) Then
            If PicInput.Visible Then PicInput.SetFocus: Exit Sub
            If mshBill.TextMatrix(mshBill.Row, mconIntColҩ��) = "" Then Exit Sub
            If Split(mshBill.TextMatrix(mshBill.Row, mconIntColԭ����), "||")(2) <> 1 Then Exit Sub
            sngLeft = Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
            sngTop = Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
            If sngTop + 1700 > Screen.Height Then
                sngTop = sngTop - mshBill.MsfObj.CellHeight - 1700
            End If
            
            With PicInput
                .Top = sngTop
                .Left = sngLeft
                .Visible = True
            End With
            Txt�Ӽ��� = "15.00000"
            With mshBill
                If Val(.TextMatrix(.Row, mconIntCol�ۼ�)) <> 0 And Val(.TextMatrix(.Row, mconIntCol�ɱ���)) <> 0 Then
                    Txt�Ӽ��� = zlStr.FormatEx((Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / Val(.TextMatrix(.Row, mconIntCol�ɱ���)) - 1) * 100, 5, , True)
                End If
            End With
            Txt�Ӽ���.Tag = Txt�Ӽ���
            Txt�Ӽ���.SetFocus
        End If
    ElseIf KeyCode = vbKeyF7 Then
        If staThis.Panels("PY").Bevel = sbrRaised Then
            Logogram staThis, 0
        Else
            Logogram staThis, 1
        End If
    End If
End Sub

Private Sub cmdProvider_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txtProvider.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select id,�ϼ�ID,ĩ��,����,����,���� From ��Ӧ�� " & _
              "Where (վ�� = [1] Or վ�� is Null) And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
              "  And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
              "Start with �ϼ�ID is null connect by prior ID =�ϼ�ID order by level,ID"
    'Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "-ҩƷ��Ӧ��", gstrNodeNo)
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 1, "��ҩ��λ", True, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, False, False, True, gstrNodeNo)
    
    If rsProvider.State = 0 Then
        txtProvider.SetFocus
        Exit Sub
    End If
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    
    Me.txtProvider.Tag = rsProvider!id
    Me.txtProvider = rsProvider!����
    mblnChange = True
    
    mshBill.SetFocus
    
    If CheckQualifications(Val(txtProvider.Tag)) = False Then
        txtProvider.Text = ""
        txtProvider.Tag = "0"
        mblnChange = True
        Exit Sub
    End If
    
    mblnChange = True
    If Val(txtProvider.Tag) <> mlng��ҩ��λID And (mint�༭״̬ = 8 Or mbln�˻�) Then
        mlng��ҩ��λID = Val(txtProvider.Tag)
        mshBill.ClearBill
        mshBill.TextMatrix(1, mconIntCol�к�) = "1"
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ӡ����
Private Sub printbill()
    Dim int��λϵ�� As Integer
    Dim strNo As String
    
    Select Case mintUnit
        Case mconint�ۼ۵�λ
            int��λϵ�� = 4
        Case mconint���ﵥλ
            int��λϵ�� = 2
        Case mconintסԺ��λ
            int��λϵ�� = 1
        Case mconintҩ�ⵥλ
            int��λϵ�� = 3
    End Select
    
    strNo = txtNO.Tag
    FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1300", "zl8_bill_1300"), _
        mint��¼״̬, int��λϵ��, "1300", IIf(mint�༭״̬ = 8 Or mbln�˻�, "ҩƷ�˻���", "ҩƷ�⹺��ⵥ"), strNo
End Sub

Private Sub CmdSave_Click()
    On Error GoTo ErrHand
    Dim strNewNO As String
    Dim BlnSuccess As Boolean, blnTrans As Boolean, bln�˻��� As Boolean
    Dim strҩƷ As String
    Dim intLop As Integer
    Dim lng�ϴ�ҩƷID As Long
    
    '�����������ݼ�
    Call SetSortRecord
    
    mstr������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    '��������ҩƷ����Ԥ���۴���
    For intLop = 1 To Me.mshBill.rows - 1
        If mshBill.TextMatrix(intLop, 0) <> "" Then '��ҩƷ
            Call AutoAdjustPrice_ByID(Val(mshBill.TextMatrix(intLop, 0)))
        End If
    Next
    
    If mint�༭״̬ = 9 Then    '�˲�
        mstrTime_End = GetBillInfo(1, mstr���ݺ�)
        If mstrTime_End = "" Then
            MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
   
        If mstrTime_End > mstrTime_Start Then
            MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
            Exit Sub
        End If

        If Not SaveCard Then Exit Sub
        
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 7 Then    '�������
        '�ȳ��������������ݲ����
        gcnOracle.BeginTrans
        blnTrans = True
        '������NO
        strNewNO = Sys.GetNextNo(21, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(strNewNO) Then Exit Sub
        '����δ����µ���
        BlnSuccess = SaveNewCard(strNewNO)
        '�������˼�¼���в�������
        If BlnSuccess Then BlnSuccess = SaveVerifyCard(strNewNO)
        '����ԭ����
        If BlnSuccess Then BlnSuccess = SaveStrike
        '����µ���
        If BlnSuccess Then BlnSuccess = SaveCheck(strNewNO)
        
        '���½���ʾ��ʽ����Ϊfalse
        mbln��ʾ��ʽ = False
        
        If BlnSuccess Then
            gcnOracle.CommitTrans
        Else
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        blnTrans = False
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 3 Then        '���
        Dim rsTemp As New ADODB.Recordset
        
        If ValidData = False Then Exit Sub '��Ҫ�Ǽ��ƻ������ɵ���ⵥ
        
        mstrTime_End = GetBillInfo(1, mstr���ݺ�)
        If mstrTime_End = "" Then
            MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
   
        If mstrTime_End > mstrTime_Start Then
            MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
            Exit Sub
        End If
       
        If chkת���ƿ�.Value = 1 Then
            If cboEnterStock.ListIndex < 0 Then
                MsgBox "Ҫ�ƿ�Ĳ��Ų���ȷ��", vbInformation, gstrSysName
                cboEnterStock.SetFocus
                Exit Sub
            End If
            If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
                MsgBox "���벿�����Ƴ����Ų�����ͬ��", vbInformation, gstrSysName
                cboEnterStock.SetFocus
                Exit Sub
            End If
        End If
       
        'Modified by ZYB 2004-05-16 ��������
        '�����ʱ���µļ۸�����Ч�������Ҫ������ɾ�������²���
        '��ΪֻӰ�����ҩƷ���ݣ���Ӧ�������������Ӱ��
        If Not ��鵥��(1, txtNO, False) And Not mblnUpdate Then
            '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
            MsgBox "�м�¼δʹ�������ۼۣ������Զ���ɸ��£��ۼۡ��ۼ۽���ۣ������º����飡", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        If Not ҩƷ�������(Txt������.Caption) Then Exit Sub
        
        '���۹�������Ƿ���ڲ��������۵�ҩƷ
        For intLop = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(intLop, 0) <> "" And gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                If IsPriceAdjustMod(Val(mshBill.TextMatrix(intLop, 0))) = True Then
                    If Val(mshBill.TextMatrix(intLop, mconIntCol�ɱ���)) <> Val(mshBill.TextMatrix(intLop, mconIntCol�ۼ�)) Then
                        MsgBox "��" & intLop & "��ҩƷ���������۹�������ⵥ���ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        mshBill.Row = intLop
                        mshBill.MsfObj.TopRow = intLop
                        Exit Sub
                    End If
                End If
            End If
        Next
                
        '��鱾�����Ƿ�Ϊ�˻�������ҩ��ʽ=1��
        gstrSQL = "Select nvl(��ҩ��ʽ,0) �˻� " & _
                  "From ҩƷ�շ���¼ " & _
                  "Where ���� =1 and NO=[1] AND ROWNUM<2 "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[����Ƿ����˻���]", txtNO.Text)
        If Not rsTemp.EOF Then
            bln�˻��� = (rsTemp!�˻� = 1)
        End If
        
        'ֻ�в����˻���ʱ���Ž������²�������Ϊ�������ʱ�޸ķ�Ʊ��Ϣ��������ֱ�����
        If Not bln�˻��� Then
            blnTrans = True
            gcnOracle.BeginTrans
            '������ʱ�޸��˵��ݣ����������ɵ��ݱ���
            If mblnChange Then
                If Not SaveCard(True) Then
                    gcnOracle.RollbackTrans: Exit Sub
                End If
            End If
            If Not SaveCheck() Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
            
            'ת���ƿⴰ��
            If chkת���ƿ�.Value = 1 And Me.cboEnterStock.ListIndex >= 0 Then
                If Check�Ƿ���ڸ����� Then
                    If MsgBox("���ҩƷ�������ڸ���������ʹ�õ����ƿ�Ĺ��ܡ�ȷ����⣬ѡ��<��>��������ˣ�ѡ��<��>��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        gcnOracle.CommitTrans
                        
                        zlPrintBill_Check
                    Else
                        gcnOracle.RollbackTrans: Exit Sub
                    End If
                Else
                    If Not Check�洢�ⷿ Then
                        If mbln����ҩƷ�޴洢�ⷿ Then
                            gcnOracle.CommitTrans
                            
                            MsgBox mstr�洢�ⷿ��ʾ
                            
                            zlPrintBill_Check
                        Else
                            gcnOracle.CommitTrans
                        
                            MsgBox mstr�洢�ⷿ��ʾ
                            
                            zlPrintBill_Check
                            
                            frmTransferCard.ShowCard Me, txtNO.Text, 11, , BlnSuccess
                        End If
                    Else
                        gcnOracle.CommitTrans
                        
                        zlPrintBill_Check
                        
                        frmTransferCard.ShowCard Me, txtNO.Text, 11, , BlnSuccess
                    End If
                End If
            Else
                gcnOracle.CommitTrans
                
                zlPrintBill_Check
            End If
        Else
            If SaveCheck Then
                If Val(zlDataBase.GetPara("��˴�ӡ", glngSys, ģ���.�⹺���)) = 1 Then
                    '��ӡ
                    If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                        printbill
                        
                        If Val(zlDataBase.GetPara("��ӡҩƷ����", glngSys, ģ���.�⹺���)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
                            '��ҩƷID˳���������
                            recSort.Sort = "ҩƷid"
                            recSort.MoveFirst
                            '��ӡҩƷ����
                            Do While Not recSort.EOF
                                If lng�ϴ�ҩƷID <> Val(recSort!ҩƷID) Then
                                    ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1300_1", Me, "ҩƷ=" & Val(recSort!ҩƷID), 2
                                    lng�ϴ�ҩƷID = recSort!ҩƷID
                                End If
                                recSort.MoveNext
                            Loop
                        End If
                        
                    End If
                End If
            End If
        End If
        blnTrans = False
        Unload Me
        Exit Sub
    End If
            
    If mint�༭״̬ = 5 Then      '�޸ķ�Ʊ��Ϣ
        If SaveRecipe = True Then
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 6 Then
        If mblnChange = False Then
            MsgBox "��¼�����������", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("��ȷʵҪ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If SaveStrike = True Then
                Unload Me
            End If
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 8 Then
        If SaveRestore Then
            If Val(zlDataBase.GetPara("���̴�ӡ", glngSys, ģ���.�⹺���)) = 1 Then
                '��ӡ
                If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                    printbill
                    
                    If Val(zlDataBase.GetPara("��ӡҩƷ����", glngSys, ģ���.�⹺���)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
                        '��ҩƷID˳���������
                        recSort.Sort = "ҩƷid"
                        recSort.MoveFirst
                        '��ӡҩƷ����
                        Do While Not recSort.EOF
                            If lng�ϴ�ҩƷID <> Val(recSort!ҩƷID) Then
                                ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1300_1", Me, "ҩƷ=" & Val(recSort!ҩƷID), 2
                                lng�ϴ�ҩƷID = recSort!ҩƷID
                            End If
                            recSort.MoveNext
                        Loop
                    End If
                    
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 2 Then
        If Not ��鵥��(1, txtNO, False) And Not mblnUpdate Then
            '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
            MsgBox "�м�¼δʹ�������ۼۣ������Զ���ɸ��£��ۼۡ��ۼ۽���ۣ������º����飡", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
    
    If mint�༭״̬ = 1 Then '��������ʱ���ж��ۼ��Ƿ��Ѿ�����

        If ����ۼ� Then
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
            
    If ValidData = False Then Exit Sub
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
        If Val(zlDataBase.GetPara("���̴�ӡ", glngSys, ģ���.�⹺���)) = 1 Then
            '��ӡ
            If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                printbill
                
                If Val(zlDataBase.GetPara("��ӡҩƷ����", glngSys, ģ���.�⹺���)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
                    '��ҩƷID˳���������
                    recSort.Sort = "ҩƷid"
                    recSort.MoveFirst
                    '��ӡҩƷ����
                    Do While Not recSort.EOF
                        If lng�ϴ�ҩƷID <> Val(recSort!ҩƷID) Then
                            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1300_1", Me, "ҩƷ=" & Val(recSort!ҩƷID), 2
                            lng�ϴ�ҩƷID = recSort!ҩƷID
                        End If
                        recSort.MoveNext
                    Loop
                End If
                    
            End If
        End If
        If mint�༭״̬ = 2 Then   '�޸�
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    SetEdit
    
'    txtProvider.Text = ""
'    txtProvider.Tag = "0"
    txtժҪ.Text = ""
    txtProvider.SetFocus
    mblnChange = False
    If txtNO.Tag <> "" Then Me.staThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNO.Tag
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Sub

Private Function ����ۼ�() As Boolean
    '���ܣ��⹺����ʱ���ж϶���ҩƷ�Ƿ��������ۼۣ������޸ĺ���ʾ
    Dim strMsg As String '������ʾ��Ϣ
    Dim i As Integer, intSum As Integer, intPriceDigit As Integer
    Dim rsPrice As New ADODB.Recordset
    Dim Dbl���� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    
    On Error GoTo errHandle
    
    ����ۼ� = False
    
    With mshBill
        For i = 1 To .rows - 1
            If mshBill.TextMatrix(i, 0) <> "" Then
                
                If Val(Split(.TextMatrix(i, mconIntColԭ����), "||")(2)) = 0 Then '�ж϶���

                    dbl���ۼ� = zlStr.FormatEx(Get�ۼ�(False, Val(.TextMatrix(i, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, mconIntCol����))) * Val(.TextMatrix(i, mconIntCol����ϵ��)), mintPriceDigit)
                    
                    If .TextMatrix(i, mconIntCol�ۼ�) <> dbl���ۼ� Then
                        intSum = intSum + 1 '��¼�����˼�������
                        
                        dbl�ɱ��� = Val(.TextMatrix(i, mconIntCol�ɱ���))
                        Dbl���� = Val(.TextMatrix(i, mconIntCol����))
                        dbl�ɱ���� = dbl�ɱ��� * Dbl����
                        dbl���۽�� = dbl���ۼ� * Dbl����
                        dbl��� = dbl���۽�� - dbl�ɱ����
                        
                        '�����ۼ��������
                        .TextMatrix(i, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ�, mintPriceDigit, , True)
                        .TextMatrix(i, mconIntCol�ۼ۽��) = zlStr.FormatEx(dbl���۽��, mintMoneyDigit, , True)
                        .TextMatrix(i, mconintCol���) = zlStr.FormatEx(dbl���, mintMoneyDigit, , True)
                        .TextMatrix(i, mconIntcol�ӳ���) = zlStr.FormatEx((Val(.TextMatrix(i, mconIntCol�ۼ�)) / dbl�ɱ��� - 1) * 100, 2) & "%"
                        
                    End If
                End If
            End If
        Next
        
        If intSum > 0 Then
            MsgBox "�м�¼δʹ�������ۼۣ��������Զ���ɸ��£��ۼۡ��ۼ۽���ۣ������º����飡", vbInformation, gstrSysName
            ����ۼ� = True
        End If
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Load()
    Dim rsTemp As ADODB.Recordset
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    Dim i As Integer, j As Integer
    Dim str������ As String
    
    On Error GoTo errHandle
    mblnLoad = False
    marrFrom = Array()
    marrInitGrid = Array()
    mblnUpdate = False
    mintBatchNoLen = GetBatchNoLen()
    mbln�Ӽ��� = Get�Ӽ���()
'    mblnʱ��ȡ�ϴ��ۼ� = IIf(Val(zldatabase.GetPara(183, 100)) = 0, False, True)
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    
    mblnȡĿ¼�в�����Ϣ = gtype_UserSysParms.P294_����ȡĿ¼�в�����Ϣ = 1
    
    txtNO = mstr���ݺ�
    txtNO.Tag = txtNO
    Call Getʱ��ҩƷֱ��ȷ���ۼ�
    Call GetSysParm
    Call GetDefineSize
    mblnEnter = True
        
    If glngModul = 1300 Then '�⹺��ⵥ�˻�
        gstrSQL = "select ��ҩ��ʽ from ҩƷ�շ���¼ where no=[1] and ��¼״̬=[2] and ����=1 and rownum=1 "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�ж��Ƿ����˿ⵥ", mstr���ݺ�, mint��¼״̬)
        
        If rsTemp.RecordCount > 0 Then
            mbln�˻� = IIf(IsNull(rsTemp!��ҩ��ʽ), False, rsTemp!��ҩ��ʽ)
        Else
            mbln�˻� = False
        End If
    Else
        mbln�˻� = False
    End If
        
    Set mrs�ֶμӳ� = Nothing
    If mintʱ�۷ֶμӳɷ�ʽ = 1 Then
        gstrSQL = "select ���, ��ͼ�, ��߼�, �ӳ���, ��۶�, ˵��, ���� from ҩƷ�ӳɷ��� order by ���"
        Set mrs�ֶμӳ� = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯ�ֶμӳ�")
    End If
    mshBill.Value = Format(Sys.Currentdate, "YYYY-MM-DD")
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�⹺������", "ҩƷ������ʾ��ʽ", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    mlng���ⷿ = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    'ȡ��ⵥλ��С��λ��
    Call GetDrugDigit(mlng���ⷿ, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)

    Call initCard
    
    mstrTime_Start = GetBillInfo(1, mstr���ݺ�)
    mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    
    'ֻ����ҩ��ⷿ����ʾ"ԭ����"��
    str�ⷿ���� = ""
    gstrSQL = "Select �������� From ��������˵�� Where ����id =[1]"
    Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "�жϿⷿ����", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsDetail.EOF
        str�ⷿ���� = str�ⷿ���� & "," & rsDetail!��������
        rsDetail.MoveNext
    Loop
    If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True
    str������ = zlDataBase.GetPara("������", glngSys, ģ���.�⹺���)
    If InStr(1, "|" & str������ & "|", "|ԭ����|") = 0 Then mshBill.ColWidth(mconIntColԭ����) = IIf(bln��ҩ�ⷿ, 800, 0)
    
    For i = 1 To mconIntColS - 1
        ReDim Preserve marrInitGrid(UBound(marrInitGrid) + 1)
        marrInitGrid(UBound(marrInitGrid)) = mshBill.TextMatrix(0, i) & "|" & mshBill.ColWidth(i)
    Next
    
    RestoreWinState Me, App.ProductName, MStrCaption
    
    For i = 1 To mconIntColS - 1
        ReDim Preserve marrFrom(UBound(marrFrom) + 1)
        marrFrom(UBound(marrFrom)) = mshBill.TextMatrix(0, i) & "|" & mshBill.ColWidth(i)
    Next
    
    For i = 0 To UBound(marrInitGrid)
        For j = 0 To UBound(marrFrom)
            If Split(marrInitGrid(i), "|")(0) = Split(marrFrom(j), "|")(0) And Split(marrInitGrid(i), "|")(1) * Split(marrFrom(j), "|")(1) = 0 Then
                mshBill.ColWidth(i + 1) = Split(marrInitGrid(i), "|")(1)
            End If
        Next
    Next

    If mint�༭״̬ <> 6 Then
        If mshBill.ColWidth(mconIntCol��������) > 0 Then
            mshBill.ColWidth(mconIntCol��������) = 0
        End If
    Else
        If mshBill.ColWidth(mconIntCol��������) = 0 Then
            mshBill.ColWidth(mconIntCol��������) = 1000
        End If
    End If

    mshBill.ColWidth(mconIntCol���) = 0
    mshBill.ColWidth(mconIntCol��������) = 0

    If mint�༭״̬ = 8 Or mbln�˻� = True Or mintUnit = mconint�ۼ۵�λ Then
        mshBill.ColWidth(mconintCol���ۼ�) = 0
        mshBill.ColWidth(mconintCol���۵�λ) = 0
        mshBill.ColWidth(mconintCol���۽��) = 0
        mshBill.ColWidth(mconintCol���۲��) = 0
    Else
        mshBill.ColWidth(mconintCol���ۼ�) = 0
        mshBill.ColWidth(mconintCol���۵�λ) = 0
        mshBill.ColWidth(mconintCol���۽��) = 0
        mshBill.ColWidth(mconintCol���۲��) = 0
        
        If InStr(1, "|" & mstrColumn_UnSelected & "|", "|���ۼ�|") = 0 Then mshBill.ColWidth(mconintCol���ۼ�) = 1000
        If InStr(1, "|" & mstrColumn_UnSelected & "|", "|���۵�λ|") = 0 Then mshBill.ColWidth(mconintCol���۵�λ) = 1000
        If InStr(1, "|" & mstrColumn_UnSelected & "|", "|���۽��|") = 0 Then mshBill.ColWidth(mconintCol���۽��) = 1000
        If InStr(1, "|" & mstrColumn_UnSelected & "|", "|���۲��|") = 0 Then mshBill.ColWidth(mconintCol���۲��) = 1000
    End If
    
'    ����ϵͳ��������ҩ����Ա�鿴����ʱ���Ƿ���ʾ�ɱ���
    If InStr(1, "|" & mstrColumn_UnSelected & "|", "|ָ��������|") = 0 Then mshBill.ColWidth(mconIntColָ��������) = IIf((mblnViewCost Or mint�༭״̬ = 7), 1000, 0)
    If InStr(1, "|" & mstrColumn_UnSelected & "|", "|�ɱ���|") = 0 Then mshBill.ColWidth(mconIntCol�ɱ���) = IIf(mblnViewCost Or mint�༭״̬ = 7, 1000, 0)
    If InStr(1, "|" & mstrColumn_UnSelected & "|", "|�ɹ���|") = 0 Then mshBill.ColWidth(mconIntCol�ɹ���) = IIf((mblnViewCost Or mint�༭״̬ = 7), 1000, 0)
    If InStr(1, "|" & mstrColumn_UnSelected & "|", "|�ɱ����|") = 0 Then mshBill.ColWidth(mconIntCol�ɱ����) = IIf(mblnViewCost Or mint�༭״̬ = 7, 900, 0)
    If InStr(1, "|" & mstrColumn_UnSelected & "|", "|���|") = 0 Then mshBill.ColWidth(mconintCol���) = IIf(mblnViewCost, 900, 0)
    If InStr(1, "|" & mstrColumn_UnSelected & "|", "|���۲��|") = 0 Then mshBill.ColWidth(mconintCol���۲��) = IIf(mblnViewCost, 1000, 0)
    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = IIf(mshBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, mshBill.ColWidth(mconIntCol��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = 0
    End If
    
    mblnLoad = True
    If Check��Ա����(UserInfo.�û�ID, cboStock.ItemData(cboStock.ListIndex)) = False And mint�༭״̬ = 7 Then
        mintParallelRecord = 6
        Exit Sub
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim num��װϵ�� As String
    Dim strOrder As String, strCompare As String
    Dim i As Long, j As Long
    Dim rsEnterStock As New ADODB.Recordset
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim str���� As String
    Dim strArray As String
    Dim dbl�ɱ��� As Double
    Dim intCostDigit As Integer        '�ɱ���С��λ��
    Dim intPriceDigit As Integer       '�ۼ�С��λ��
    Dim intNumberDigit As Integer      '����С��λ��
    Dim intMoneyDigit As Integer       '���С��λ��
    Dim blnAllPay As Boolean
    Dim strҩ�� As String
    Dim strSqlOrder As String
    Dim rs As ADODB.Recordset
    
    blnAllPay = True
    On Error GoTo errHandle
    strOrder = zlDataBase.GetPara("����", glngSys, ģ���.�⹺���)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "���"
    
    If strCompare = "0" Then
        strSqlOrder = "���"
    ElseIf strCompare = "1" Then
        strSqlOrder = "ҩƷ����"
    ElseIf strCompare = "2" Then
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strSqlOrder = "ͨ����"
        Else
            strSqlOrder = "Nvl(��Ʒ��, ͨ����)"
        End If
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")
    
    intCostDigit = mintCostDigit
    intPriceDigit = mintPriceDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
        
    If mint�༭״̬ <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            For i = 0 To .ListCount - 1
                cboStock.AddItem .List(i)
                cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
            Next
            mintcboIndex = .ListIndex
            cboStock.ListIndex = .ListIndex
            cboStock.Enabled = .Enabled
        End With
    End If
    
    If mint�༭״̬ = 3 Then
        With cboEnterStock
            .Clear
            Set rsEnterStock = ReturnSQL(cboStock.ItemData(cboStock.ListIndex), MStrCaption & "[ҩƷ������ȡ�ƿ�ⷿ]", True)
            
            Do While Not rsEnterStock.EOF
                .AddItem rsEnterStock.Fields(2)
                .ItemData(.NewIndex) = rsEnterStock.Fields(0)
                rsEnterStock.MoveNext
            Loop
    
            If .ListCount > 0 Then
                .ListIndex = 0
            End If
                                
        End With
    End If
    
    Select Case mint�༭״̬
        Case 1, 8
            Txt������ = UserInfo.�û�����
            Txt�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
'            Txt�޸��� = UserInfo.�û�����
'            Txt�޸����� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4, 5, 6, 7, 9
            initGrid
            If mint�༭״̬ = 4 Then
                gstrSQL = "select b.id,b.���� from ҩƷ�շ���¼ a,���ű� b where a.�ⷿid=b.id and A.���� = 1 and a.no=[1] "
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�)
                
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                With cboStock
                    .AddItem rsInitCard!����
                    .ItemData(.NewIndex) = rsInitCard!id
                    .ListIndex = 0
                End With
                rsInitCard.Close
            End If
            
            Select Case mintUnit
                Case mconint�ۼ۵�λ
                    strUnitQuantity = "D.���㵥λ AS �ۼ۵�λ,D.���㵥λ AS ��λ, A.��д���� AS ����,'1' as ����ϵ��, "
                    num��װϵ�� = "1"
                Case mconint���ﵥλ
                    strUnitQuantity = "D.���㵥λ AS �ۼ۵�λ,B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ����,B.�����װ as ����ϵ��,"
                    num��װϵ�� = "B.�����װ"
                Case mconintסԺ��λ
                    strUnitQuantity = "D.���㵥λ AS �ۼ۵�λ,B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ����,B.סԺ��װ as ����ϵ��,"
                    num��װϵ�� = "B.סԺ��װ"
                Case mconintҩ�ⵥλ
                    strUnitQuantity = "D.���㵥λ AS �ۼ۵�λ,B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ����,B.ҩ���װ as ����ϵ��,"
                    num��װϵ�� = "B.ҩ���װ"
            End Select
            
            Select Case mint�༭״̬
            Case 5, 7     '�޸ķ�Ʊ,�������
                If mint��¼״̬ = 1 Then
                    gstrSQL = "SELECT * FROM (SELECT DISTINCT A.ҩƷID,A.���,'[' || D.���� || ']' As ҩƷ����, D.���� As ͨ����, E.���� As ��Ʒ��," & _
                        " B.ҩƷ��Դ,B.����ҩ��,D.���,D.���� AS ԭ������,A.����, A.ԭ����,A.����,NVL(A.����,0) ����," & _
                        " NVL(B.�б�ҩƷ,0) �б�ҩƷ,NVL(B.���������,0) ���������,B.���Ч��,A.Ч��," & strUnitQuantity & _
                        " nvl(A.����,b.ָ��������)*" & num��װϵ�� & " AS ָ�������� ,A.�ɱ���*" & num��װϵ�� & " AS �ɹ���, " & _
                        " A.�ɱ���� AS �ɹ����,D.�Ƿ���,B.ҩ������ ҩ����������," & _
                        " DECODE(A.����, NULL, 0, A.����) AS ����, A.���ۼ�*" & num��װϵ�� & " AS ���ۼ�,A.���۽��,A.���, " & _
                        " A.��׼�ĺ�,C.�������,C.�������, C.��Ʊ�� ,c.��Ʊ����, C.��Ʊ����, C.��Ʊ���,A.��ҩ��λID,F.���� AS ��Ӧ��, A.������,A.��������," & _
                        " A.�޸���,A.�޸�����,A.�����,A.�������,A.�ⷿID,G.���� AS ����,NVL(C.�������,0) AS �������,Nvl(A.��ҩ��ʽ,0) �˻�,A.���,A.���ս���," & _
                        " A.��Ʒ�ϸ�֤,A.��������,A.��ҩ�� As �˲���,A.��ҩ���� As �˲�����,B.ҩ�ۼ���, Nvl(A.�÷�, 0) As ����,A.Ƶ�� As �ӳ���,c.�����־,a.�ƻ�id " & _
                        " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E ,Ӧ����¼ C,��Ӧ�� F,���ű� G " & _
                        " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=D.ID AND A.�ⷿID=G.ID" & _
                        " AND A.��ҩ��λID=F.ID AND SUBSTR(F.����,1,1)=1" & _
                        " AND A.ID = C.�շ�ID(+) AND C.ϵͳ��ʶ(+)=1 AND C.��¼����(+)=0 " & _
                        " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                        " AND A.��¼״̬ =[2] " & _
                        " AND A.���� = 1 AND A.NO =[1] " & _
                        " ) ORDER BY " & strSqlOrder
                Else
                    gstrSQL = "SELECT * FROM (SELECT DISTINCT A.ҩƷID,A.���,'[' || D.���� || ']' As ҩƷ����, D.���� As ͨ����, E.���� As ��Ʒ��," & _
                        " B.ҩƷ��Դ,B.����ҩ��,D.���,D.���� AS ԭ������,A.����,A.ԭ����,A.����,A.����," & _
                        " NVL(B.�б�ҩƷ,0) �б�ҩƷ,NVL(B.���������,0) ���������,B.���Ч��,A.Ч��," & strUnitQuantity & _
                        " nvl(A.����,b.ָ��������)*" & num��װϵ�� & " AS ָ�������� ,A.�ɱ���*" & num��װϵ�� & " AS �ɹ���," & _
                        " A.�ɱ���� AS �ɹ����,D.�Ƿ���,B.ҩ������ ҩ����������,  " & _
                        " DECODE(A.����, NULL, 0, A.����) AS ����, A.���ۼ�*" & num��װϵ�� & " AS ���ۼ� ,A.���۽��,A.���," & _
                        " A.��׼�ĺ�, A.�������, A.�������, A.��Ʊ��,a.��Ʊ����,A.��Ʊ����,A.��Ʊ���,A.��ҩ��λID,F.���� AS ��Ӧ��, A.�ⷿID,G.���� AS ����,NVL(A.�������,0) AS �������,A.�˻�,A.��������,A.��ҩ�� As �˲���,A.��ҩ���� As �˲�����,B.ҩ�ۼ���,A.����,A.�ӳ���,�����־,a.�ƻ�id " & _
                        " FROM " & _
                        "     (SELECT X.NO, SUM(ʵ������) AS ��д����,SUM(�ɱ����) AS �ɱ����,�������,�������,��Ʊ��,��Ʊ����,��Ʊ����," & _
                        "      X.ҩƷID,X.���,X.����, X.ԭ����,X.����,NVL(X.����,0) ����,X.Ч��,X.����,X.�ɱ���,X.���ۼ�,X.����," & _
                        "      X.��ҩ��λID,�ⷿID,NVL(Y.�������,0) AS �������,Nvl(X.��ҩ��ʽ,0) As �˻�,X.��������,X.��׼�ĺ�,��Ʊ���,X.��ҩ��,X.��ҩ����,Sum(���۽��) ���۽��,Sum(���) ���,Sum(To_Number(Nvl(�÷�, 0))) As ����,Ƶ�� As �ӳ���,y.�����־,x.�ƻ�id " & _
                        "      FROM ҩƷ�շ���¼ X,(Select ���,��ĿID,�������,�������,��Ʊ��,��Ʊ����,��Ʊ����,�������,Sum(��Ʊ���) ��Ʊ���,�����־  From Ӧ����¼ " & _
                        "      Where ϵͳ��ʶ = 1 And ��¼���� =0 And ��ⵥ�ݺ�=[1] Group By ���,��ĿID,�������,�������,��Ʊ��,��Ʊ����,��Ʊ����,�������,�����־) Y " & _
                        "      WHERE X.��� = Y.���(+) And X.ҩƷID=Y.��ĿID(+) AND X.NO=[1] AND ����=1 " & _
                        "      GROUP BY X.NO,X.ҩƷID,X.���,X.����,X.ԭ����,X.����,NVL(X.����,0),X.Ч��,X.����,X.�ɱ���,X.���ۼ�,X.����," & _
                        "      X.��ҩ��λID,X.�ⷿID,�������,�������,��Ʊ��,��Ʊ����,��Ʊ����,NVL(Y.�������,0),Nvl(X.��ҩ��ʽ,0),X.��������,X.��׼�ĺ�,��Ʊ���,X.��ҩ��,X.��ҩ����,X.Ƶ��,y.�����־,x.�ƻ�id " & _
                        "      HAVING SUM(ʵ������)<>0 ) A," & _
                        "      ҩƷ��� B,�շ���Ŀ���� E ,�շ���ĿĿ¼ D,��Ӧ�� F,���ű� G " & _
                        " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=D.ID AND A.�ⷿID=G.ID" & _
                        " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                        " AND A.��ҩ��λID=F.ID AND SUBSTR(F.����,1,1)=1 " & _
                        " ) ORDER BY " & strSqlOrder
                End If
                
            Case 6      '����
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.ҩƷID,A.���,'[' || D.���� || ']' As ҩƷ����, D.���� As ͨ����, E.���� As ��Ʒ��," & _
                    " B.ҩƷ��Դ,B.����ҩ��,D.���,D.���� AS ԭ������,A.����,A.ԭ����,A.����," & _
                    " NVL(B.�б�ҩƷ,0) �б�ҩƷ,NVL(B.���������,0) ���������,B.���Ч��,A.Ч��," & strUnitQuantity & _
                    " nvl(A.����,b.ָ��������)*" & num��װϵ�� & " AS ָ�������� ,A.�ɱ���*" & num��װϵ�� & " AS �ɹ���," & _
                    " A.�ɱ���� AS �ɹ����,D.�Ƿ���,B.ҩ������ ҩ����������,  " & _
                    " DECODE(A.����, NULL, 0, A.����) AS ����, A.���ۼ�*" & num��װϵ�� & " AS ���ۼ� ,0 AS ���۽��,0 AS ���,A.����, " & _
                    " A.��׼�ĺ�,A.�������,A.�������, A.��Ʊ��,a.��Ʊ����,A.��Ʊ����,A.��Ʊ���,A.��ҩ��λID,F.���� AS ��Ӧ��, A.�ⷿID,G.���� AS ����,NVL(A.�������,0) AS �������,A.�˻�,A.��������,A.����,A.��ҩ�� As �˲���,A.��ҩ���� As �˲�����,B.ҩ�ۼ���,A.�ӳ���,a.�����־,a.�ƻ�id,a.��� " & _
                    " FROM " & _
                    "     (SELECT MIN(X.ID) AS ID, SUM(ʵ������) AS ��д����,SUM(�ɱ����) AS �ɱ����,�������,�������,��Ʊ��,��Ʊ����,��Ʊ����,Sum(��Ʊ���) As ��Ʊ���," & _
                    "      X.ҩƷID,X.���,X.����, X.ԭ����,X.����,X.Ч��,X.����,X.�ɱ���,X.���ۼ�,X.����," & _
                    "      X.��ҩ��λID,�ⷿID,NVL(Y.�������,0) AS �������,Nvl(X.��ҩ��ʽ,0) As �˻�,X.��������,X.��׼�ĺ�,NVL(X.����,0) ����,X.��ҩ��,X.��ҩ����,Sum(To_Number(Nvl(�÷�, 0))) As ����,Ƶ�� As �ӳ���,y.�����־,x.�ƻ�id,x.���  " & _
                    "      FROM ҩƷ�շ���¼ X,(SELECT �շ�id,�������,�������,�������,��Ʊ��,��Ʊ����,��Ʊ����,��Ʊ���,�����־ FROM Ӧ����¼ WHERE ϵͳ��ʶ=1 AND ��¼����=0) Y " & _
                    "      WHERE X.ID=Y.�շ�ID(+) AND X.NO=[1] AND ����=1 " & _
                    "      GROUP BY X.ҩƷID,X.���,X.����,X.ԭ����,X.����,X.Ч��,X.����,X.�ɱ���,X.���ۼ�,X.����," & _
                    "      X.��ҩ��λID,X.�ⷿID,�������,�������,��Ʊ��,��Ʊ����,��Ʊ����,NVL(Y.�������,0),Nvl(X.��ҩ��ʽ,0),X.��������,X.��׼�ĺ�,NVL(X.����,0),X.��ҩ��,X.��ҩ����,X.Ƶ��,�����־,x.�ƻ�id,x.��� " & _
                    "      HAVING SUM(ʵ������)<>0 ) A," & _
                    "      ҩƷ��� B,�շ���Ŀ���� E ,�շ���ĿĿ¼ D,��Ӧ�� F,���ű� G " & _
                    " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=D.ID AND A.�ⷿID=G.ID" & _
                    " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                    " AND A.��ҩ��λID=F.ID AND SUBSTR(F.����,1,1)=1 " & _
                    " ) ORDER BY " & strSqlOrder
            Case Else
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.ҩƷID,A.���,'[' || D.���� || ']' As ҩƷ����, D.���� As ͨ����, E.���� As ��Ʒ��," & _
                    " B.ҩƷ��Դ,B.����ҩ��,D.���,D.���� AS ԭ������,A.����, A.ԭ����,A.����,NVL(A.����,0) ����," & _
                    " NVL(B.�б�ҩƷ,0) �б�ҩƷ,NVL(B.���������,0) ���������,B.���Ч��,A.Ч��," & strUnitQuantity & _
                    " nvl(A.����,b.ָ��������)*" & num��װϵ�� & " AS ָ�������� ,A.�ɱ���*" & num��װϵ�� & " AS �ɹ���, " & _
                    " A.�ɱ���� AS �ɹ����,D.�Ƿ���,B.ҩ������ ҩ����������," & _
                    " DECODE(A.����, NULL, 0, A.����) AS ����, A.���ۼ�*" & num��װϵ�� & " AS ���ۼ�,A.���۽��,A.���, " & _
                    " A.��׼�ĺ�,C.�������,C.�������, C.��Ʊ�� ,c.��Ʊ����, C.��Ʊ����, C.��Ʊ���,A.��ҩ��λID,F.���� AS ��Ӧ��, A.ժҪ,A.������,A.��������," & _
                    " A.�޸���,A.�޸�����,A.�����,A.�������,A.�ⷿID,G.���� AS ����,NVL(C.�������,0) AS �������,Nvl(A.��ҩ��ʽ,0) �˻�,A.���,A.���ս���,A.��Ʒ�ϸ�֤," & _
                    " A.��������,A.��ҩ�� As �˲���,A.��ҩ���� As �˲�����,B.ҩ�ۼ���, Nvl(A.�÷�, 0) As ����,A.Ƶ�� As �ӳ���,A.�Է�����ID,a.�ƻ�id " & _
                    " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E ,Ӧ����¼ C,��Ӧ�� F,���ű� G " & _
                    " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=D.ID AND A.�ⷿID=G.ID" & _
                    " AND A.��ҩ��λID=F.ID AND SUBSTR(F.����,1,1)=1" & _
                    " AND A.ID = C.�շ�ID(+) AND C.ϵͳ��ʶ(+)=1 AND C.��¼����(+)=0 " & _
                    " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                    " AND A.��¼״̬ =[2] " & _
                    " AND A.���� = 1 AND A.NO = [1] " & _
                    " ) ORDER BY " & strSqlOrder
            End Select
             
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�, mint��¼״̬)
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Select Case mint�༭״̬
                Case 2, 6, 9 '�޸ġ��������˲�
                    If mint�༭״̬ = 2 Then
                        Txt������ = rsInitCard!������
                        Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
                        Txt�޸��� = IIf(IsNull(rsInitCard!�޸���), "", rsInitCard!�޸���)
                        Txt�޸����� = IIf(IsNull(rsInitCard!�޸�����), "", Format(rsInitCard!�޸�����, "yyyy-mm-dd hh:mm:ss"))
                        Txt����� = ""
                        Txt������� = ""
                    Else
                        Txt������ = UserInfo.�û�����
                        Txt�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
'                        Txt�޸��� = UserInfo.�û�����
'                        Txt�޸����� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                        Txt����� = UserInfo.�û�����
                        Txt������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                        If mint�༭״̬ = 9 Then
                            Txt������ = rsInitCard!������
                            Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
                            Txt�޸��� = IIf(IsNull(rsInitCard!�޸���), "", rsInitCard!�޸���)
                            Txt�޸����� = IIf(IsNull(rsInitCard!�޸�����), "", Format(rsInitCard!�޸�����, "yyyy-mm-dd hh:mm:ss"))
                            Lbl�����.Caption = "�˲���"
                            Lbl�������.Caption = "�˲�����"
                            lbl�˲���.Visible = False
                            txt�˲���.Visible = False
                            lbl�˲�����.Visible = False
                            txt�˲�����.Visible = False
                        End If
                    End If
                Case Else '3�����գ�4���鿴��5:�޸ķ�Ʊ��7���������
                    If (mint�༭״̬ = 5 Or mint�༭״̬ = 7) And mint��¼״̬ <> 1 Then
                        Txt������ = UserInfo.�û�����
                        Txt�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                        Txt�޸��� = UserInfo.�û�����
                        Txt�޸����� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    Else
                        Txt������ = rsInitCard!������
                        Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
                        Txt�޸��� = IIf(IsNull(rsInitCard!�޸���), "", rsInitCard!�޸���)
                        Txt�޸����� = IIf(IsNull(rsInitCard!�޸�����), "", Format(rsInitCard!�޸�����, "yyyy-mm-dd hh:mm:ss"))
                        Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
                        Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
                    End If
            End Select
            
            '��ҽ���������Զ������ƿⲿ��
            If mint�༭״̬ = 3 Then        '���
                If nvl(rsInitCard!�Է�����id, 0) > 0 Then
                    chkת���ƿ�.Tag = rsInitCard!�Է�����id
                    chkת���ƿ�.Value = 1
                    chkת���ƿ�.Enabled = False
                    For i = 0 To cboEnterStock.ListCount
                        If Val(cboEnterStock.ItemData(i)) = rsInitCard!�Է�����id Then
                            cboEnterStock.ListIndex = i
                            Exit For
                        End If
                    Next
                    cboEnterStock.Enabled = False
                End If
            ElseIf mint�༭״̬ = 9 Then    '�˲�
                chkת���ƿ�.Tag = nvl(rsInitCard!�Է�����id)
            End If
            
            txt�˲���.Caption = IIf(IsNull(rsInitCard!�˲���), "", rsInitCard!�˲���)
            txt�˲�����.Caption = IIf(IsNull(rsInitCard!�˲�����), "", Format(rsInitCard!�˲�����, "yyyy-mm-dd hh:mm:ss"))

            
            txtProvider.Tag = rsInitCard!��ҩ��λID
            txtProvider.Text = rsInitCard!��Ӧ��
            
            If mint�༭״̬ = 5 Or mint�༭״̬ = 6 Or mint�༭״̬ = 7 Then
                txtժҪ.Text = GetժҪ(mstr���ݺ�, mint�༭״̬)
            Else
                txtժҪ.Text = IIf(IsNull(rsInitCard!ժҪ), "", rsInitCard!ժҪ)
            End If
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            If mint�༭״̬ = 7 Then
                'ֻҪ��һ�ʸ��˿�Ͳ�������в������
                With rsInitCard
                    Do While Not .EOF
                        If !������� <> 0 Then
                            mintParallelRecord = 5        '�ѱ������˸���
                            Exit Sub
                        Else
                            '����Ƿ���ڲ��ָ�������
                            gstrSQL = "Select Nvl(Max(�������), 0) ������� From Ӧ����¼ " & _
                                " where �շ�id=(Select Id From ҩƷ�շ���¼ Where ����=1 And No=[1] And (Mod(��¼״̬,3)=0 Or ��¼״̬=1) " & _
                                " And ���=[2]) "
                            strOrder = rsInitCard!���
                            Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ�������]", txtNO.Text, strOrder)
                            
                            If rs!������� <> 0 Then
                                mintParallelRecord = 5
                                Exit Sub
                            End If
                        End If
                        .MoveNext
                    Loop
                    If .RecordCount <> 0 Then .MoveFirst
                End With
            End If
            
            intRow = 0
            mbln�˻� = (rsInitCard!�˻� = 1)
            If mbln�˻� Then LblTitle.Caption = Mid(LblTitle.Caption, 1, Len(LblTitle.Caption) - 5) & "�˻���"
            With mshBill
                Do While Not rsInitCard.EOF
                    intRow = intRow + 1
                    .rows = .rows + 1
                    
                    .TextMatrix(intRow, 0) = rsInitCard.Fields(0)
                    
                    If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                        strҩ�� = rsInitCard!ͨ����
                    Else
                        strҩ�� = IIf(IsNull(rsInitCard!��Ʒ��), rsInitCard!ͨ����, rsInitCard!��Ʒ��)
                    End If
                    
                    .TextMatrix(intRow, mconIntColҩƷ���������) = rsInitCard!ҩƷ���� & strҩ��
                    .TextMatrix(intRow, mconIntColҩƷ����) = rsInitCard!ҩƷ����
                    .TextMatrix(intRow, mconIntColҩƷ����) = strҩ��
                    .TextMatrix(intRow, 0) = rsInitCard!ҩƷID
                    If mintDrugNameShow = 1 Then
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
                    ElseIf mintDrugNameShow = 2 Then
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
                    Else
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ���������)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol��Ʒ��) = IIf(IsNull(rsInitCard!��Ʒ��), "", rsInitCard!��Ʒ��)
                    
                    .TextMatrix(intRow, mconIntCol��Դ) = nvl(rsInitCard!ҩƷ��Դ)
                    .TextMatrix(intRow, mconIntCol����ҩ��) = nvl(rsInitCard!����ҩ��)
                    .TextMatrix(intRow, mconIntColҩ�ۼ���) = IIf(IsNull(rsInitCard!ҩ�ۼ���), "", rsInitCard!ҩ�ۼ���)
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(rsInitCard!ԭ����), "", rsInitCard!ԭ����)
                    .TextMatrix(intRow, mconIntCol��λ) = rsInitCard!��λ
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsInitCard!Ч��), "", rsInitCard!Ч��)
                    
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
                        '����Ϊ��Ч��
                        .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
                    End If
                
                    If mbln�˻� Then
                        .TextMatrix(intRow, mconIntCol����) = zlStr.FormatEx(rsInitCard!���� * IIf(mint�༭״̬ = 6, 1, -1), intNumberDigit, , True)
                        .TextMatrix(intRow, mconIntCol�ɱ����) = zlStr.FormatEx(rsInitCard!�ɹ���� * IIf(mint�༭״̬ = 6, 1, -1), intMoneyDigit, , True)
                        .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(rsInitCard!���۽�� * IIf(mint�༭״̬ = 6, 1, -1), intMoneyDigit, , True)
                        .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(rsInitCard!��� * IIf(mint�༭״̬ = 6, 1, -1), intMoneyDigit, , True)
                        .TextMatrix(intRow, mconintcol��Ʊ���) = IIf(zlStr.FormatEx(nvl(rsInitCard!��Ʊ���, 0) * IIf(mint�༭״̬ = 6, 1, -1), intMoneyDigit) = "0.00", "", zlStr.FormatEx(nvl(rsInitCard!��Ʊ���, 0) * IIf(mint�༭״̬ = 6, 1, -1), intMoneyDigit, , True))
                    Else
                        .TextMatrix(intRow, mconIntCol����) = zlStr.FormatEx(rsInitCard!����, intNumberDigit, , True)
                        .TextMatrix(intRow, mconIntCol�ɱ����) = zlStr.FormatEx(IIf(mint�༭״̬ = 6, 0, rsInitCard!�ɹ����), intMoneyDigit, , True)
                        .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(rsInitCard!���۽��, intMoneyDigit, , True)
                        .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(rsInitCard!���, intMoneyDigit, , True)
                        .TextMatrix(intRow, mconintcol��Ʊ���) = IIf(zlStr.FormatEx(IIf(IsNull(rsInitCard!��Ʊ���), "0", rsInitCard!��Ʊ���), intMoneyDigit) = "0.00", "", zlStr.FormatEx(IIf(IsNull(rsInitCard!��Ʊ���), "0", rsInitCard!��Ʊ���), intMoneyDigit, , True))
                    End If
                    .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx(rsInitCard!�ɹ���, intCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!���ۼ�, intPriceDigit, , True)
                    .TextMatrix(intRow, mconIntCol����) = rsInitCard!����
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsInitCard!��׼�ĺ�), "", rsInitCard!��׼�ĺ�)
                    .TextMatrix(intRow, mconintcol�������) = IIf(IsNull(rsInitCard!�������), "", rsInitCard!�������)
                    .TextMatrix(intRow, mconintcol�������) = IIf(IsNull(rsInitCard!�������), "", rsInitCard!�������)
                    .TextMatrix(intRow, mconintcol��Ʊ��) = IIf(IsNull(rsInitCard!��Ʊ��), "", rsInitCard!��Ʊ��)
                    .TextMatrix(intRow, mconintcol��Ʊ����) = IIf(IsNull(rsInitCard!��Ʊ����), "", rsInitCard!��Ʊ����)
                    .TextMatrix(intRow, mconIntCol��Ʊ����) = IIf(IsNull(rsInitCard!��Ʊ����), "", rsInitCard!��Ʊ����)
                    .TextMatrix(intRow, mconIntColָ��������) = zlStr.FormatEx(rsInitCard!ָ��������, intCostDigit, , True)
                    .TextMatrix(intRow, mconIntColԭ������) = IIf(IsNull(rsInitCard!ԭ������), "!", rsInitCard!ԭ������)
                    
                    .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(rsInitCard!���Ч��), "0", rsInitCard!���Ч��) & "||" & rsInitCard!�ӳ��� * 100 & "||" & IIf(IsNull(rsInitCard!�Ƿ���), 0, rsInitCard!�Ƿ���) & "||" & IIf(IsNull(rsInitCard!ҩ����������), 0, rsInitCard!ҩ����������)
                    
                    '��������
                    Call GetҩƷ��������(intRow)
                    
                    'ʱ�۷���ҩƷ������Ҫ���������ۼۡ��ۼ۽����
                    If .TextMatrix(intRow, mconIntColԭ����) <> "" Then
                        If Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol��������)) = 1 Then
                            .TextMatrix(intRow, mconintCol���۵�λ) = rsInitCard!�ۼ۵�λ
                            .TextMatrix(intRow, mconintCol���ۼ�) = zlStr.FormatEx(rsInitCard!���ۼ� / Val(rsInitCard!����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�, , True)

                            If mbln�˻� Then
                                .TextMatrix(intRow, mconintCol���۽��) = zlStr.FormatEx(-1 * rsInitCard!���۽��, intMoneyDigit, , True)
                                .TextMatrix(intRow, mconintCol���۲��) = zlStr.FormatEx(-1 * rsInitCard!���, intMoneyDigit, , True)
                            Else
                                .TextMatrix(intRow, mconintCol���۽��) = zlStr.FormatEx(rsInitCard!���۽��, intMoneyDigit, , True)
                                .TextMatrix(intRow, mconintCol���۲��) = zlStr.FormatEx(rsInitCard!���, intMoneyDigit, , True)
                            End If

                            If mint�༭״̬ <> 6 Then   '���ǳ���ʱ
                                If mbln�˻� Then
                                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(-1 * (rsInitCard!���۽�� - rsInitCard!����), intMoneyDigit, , True)
                                    .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(-1 * (rsInitCard!��� - rsInitCard!����), intMoneyDigit, , True)
                                Else
                                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol���۽��)) - Val(rsInitCard!����), intMoneyDigit, , True)
                                    .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol���۲��)) - Val(rsInitCard!����), intMoneyDigit, , True)
                                End If
                                If Val(.TextMatrix(intRow, mconIntCol����)) <> 0 And rsInitCard!���� <> 0 Then
                                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)) / Val(.TextMatrix(intRow, mconIntCol����)), intPriceDigit, , True)
                                Else
                                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!���ۼ�, intPriceDigit, , True)
                                End If
                            Else
                                '����ʱ
                                .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(0, intMoneyDigit, , True)
                                .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(0, intMoneyDigit, , True)
                                .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconintCol���ۼ�)) * Val(rsInitCard!����ϵ��) * Val(rsInitCard!����) - Val(rsInitCard!����)) / Val(rsInitCard!����), intPriceDigit, , True)
                            End If
                        End If
                    End If
                    
                    .TextMatrix(intRow, mconintcol����) = ""
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mconIntCol���) = rsInitCard!���
                    .TextMatrix(intRow, mconIntCol��������) = IIf(IsNull(rsInitCard!��������), "", rsInitCard!��������)
                    .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ɱ���)) * 100 / IIf(Val(.TextMatrix(intRow, mconIntCol����)) = 0, 1, Val(.TextMatrix(intRow, mconIntCol����))), intCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol�Ƿ�����) = "��"
                    If mint�༭״̬ = 6 Then
                        .TextMatrix(intRow, mconIntCol���) = nvl(rsInitCard!���)
                    End If
                    If mint�༭״̬ <> 6 Then
                        If (mint�༭״̬ = 5 Or mint�༭״̬ = 7) And mint��¼״̬ <> 1 Then
                            .TextMatrix(intRow, mconIntCol���) = ""
                            .TextMatrix(intRow, mconIntCol���ս���) = ""
                            .TextMatrix(intRow, mconintcol��Ʒ�ϸ�֤) = ""
                        Else
                            .TextMatrix(intRow, mconIntCol���) = nvl(rsInitCard!���)
                            .TextMatrix(intRow, mconIntCol���ս���) = nvl(rsInitCard!���ս���)
                            .TextMatrix(intRow, mconintcol��Ʒ�ϸ�֤) = nvl(rsInitCard!��Ʒ�ϸ�֤)
                        End If
                    End If
                    
                    '���ݲ�������ʱ��ҩƷ�ۼ۹�ʽ�гɱ��۵��㷨
                    dbl�ɱ��� = IIf(mintʱ������ۼۼӳɷ�ʽ = 0, Val(.TextMatrix(intRow, mconIntCol�ɱ���)), Val(.TextMatrix(intRow, mconIntCol�ɹ���)))
                    
                    '����ӳ���
                    If Val(.TextMatrix(intRow, mconIntCol�ۼ�)) <> 0 And dbl�ɱ��� <> 0 Then
                        If IIf(IsNull(rsInitCard!�ӳ���), "", rsInitCard!�ӳ���) <> "" Then
                            .TextMatrix(intRow, mconIntcol�ӳ���) = zlStr.FormatEx(Val(rsInitCard!�ӳ���) * 100, 2) & "%"
                        Else
                            .TextMatrix(intRow, mconIntcol�ӳ���) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol�ۼ�)) / dbl�ɱ��� - 1) * 100, 2) & "%"
                        End If
                    End If
                    .TextMatrix(intRow, mconIntCol�ƻ�id) = IIf(IsNull(rsInitCard!�ƻ�id), "", rsInitCard!�ƻ�id)

                    '�б�ҩƷ��Ҫ��ɫ
                    mblnEnter = False
                    .Row = intRow
                    For i = mconIntColҩ�� To .Cols - 1
                        j = .ColData(i)
                        If .ColData(i) = 5 Then .ColData(i) = 0
                        .Col = i
                        If rsInitCard!�б�ҩƷ = 1 Then
                            .MsfObj.CellForeColor = IIf(rsInitCard!��������� = 0, &H800000, &H800080)
                        Else
                            .MsfObj.CellForeColor = IIf(rsInitCard!��������� = 0, &H0, &H40&)     ' &H40C0&
                        End If
                        .ColData(i) = j
                    Next
                    mblnEnter = True
                    
                    If mint�༭״̬ = 6 Then
                        .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(0, intNumberDigit, , True)
                        .RowData(intRow) = rsInitCard!�������
                        .TextMatrix(intRow, mconIntCol����) = rsInitCard!����
                        
                        If rsInitCard!������� = 0 Then
                            '����Ƿ���ڲ��ָ�������
                            gstrSQL = "Select Nvl(Max(�������), 0) ������� From Ӧ����¼ " & _
                                " where �շ�id=(Select Id From ҩƷ�շ���¼ Where ����=1 And No=[1] And (Mod(��¼״̬,3)=0 Or ��¼״̬=1) " & _
                                " And ���=[2]) "
                            Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ�������]", txtNO.Text, Val(.TextMatrix(intRow, mconIntCol���)))
                            
                            If rs!������� = 0 Then
                                blnAllPay = False
                            End If
                        End If
                    Else
'                        If (mint�༭״̬ = 5 Or mint�༭״̬ = 7) And mint��¼״̬ <> 1 Then
'                            .TextMatrix(intRow, mconIntCol����) = 0
'                        Else
                            .TextMatrix(intRow, mconIntCol����) = rsInitCard!����
'                        End If
                        .RowData(intRow) = nvl(rsInitCard!�������, 0)
                    End If
                    
                    If mint�༭״̬ = 5 Or mint�༭״̬ = 6 Or mint�༭״̬ = 7 Then
                        .TextMatrix(intRow, mconIntCol�����־) = IIf(IsNull(rsInitCard!�����־), "��", IIf(rsInitCard!�����־ = 0, "��", "��"))
                    End If
                    
                    rsInitCard.MoveNext
                Loop
                .Col = mconIntColҩ��
                .CmdVisible = False
            End With
            rsInitCard.Close
    End Select
    SetEdit         '���ñ༭����
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    Call ��ʾ�ϼƽ��
    
    If mint�༭״̬ = 6 And blnAllPay = True Then
        mintParallelRecord = 7
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            
            cboStock.Enabled = False
            txtProvider.Enabled = False
            cmdProvider.Enabled = False
            txtժҪ.Enabled = False
            
            For intCol = 0 To .Cols - 1
                .ColData(intCol) = 0
            Next
            If mint�༭״̬ = 3 Then
                .ColData(mconIntCol�ɱ���) = IIf(InStr(1, mstrControlItem, ",�ɱ���,") > 0, 4, 5)
                .ColData(mconIntCol�ɹ���) = IIf(InStr(1, mstrControlItem, ",�ɹ���,") > 0, 4, 5)
                .ColData(mconIntCol�ۼ�) = IIf(InStr(1, mstrControlItem, ",�ۼ�,") > 0, 4, 5)
                .ColData(mconIntCol����) = IIf(InStr(1, mstrControlItem, ",����,") > 0, 4, 5)
                .ColData(mconIntCol�ɱ����) = IIf(InStr(1, mstrControlItem, ",�ɱ����,") > 0, 4, 5)
                .ColData(mconIntCol���) = IIf(InStr(1, mstrControlItem, ",���,") > 0, 1, 5)
                .ColData(mconIntCol���ս���) = IIf(InStr(1, mstrControlItem, ",���ս���,") > 0, 1, 5)
                .ColData(mconintcol��Ʊ��) = IIf(InStr(1, mstrControlItem, ",��Ʊ��,") > 0, 4, 5)
                .ColData(mconintcol��Ʊ����) = IIf(InStr(1, mstrControlItem, ",��Ʊ����,") > 0, 4, 5)
                .ColData(mconIntCol��Ʊ����) = IIf(InStr(1, mstrControlItem, ",��Ʊ����,") > 0, 2, 5)
                .ColData(mconintcol��Ʊ���) = IIf(InStr(1, mstrControlItem, ",��Ʊ���,") > 0, 4, 5)

                txtProvider.Enabled = True
                cmdProvider.Enabled = True

                If mint��¼״̬ <> 1 Then
                    txtProvider.Enabled = False
                    cmdProvider.Enabled = False
                End If
            ElseIf mint�༭״̬ = 5 Then
                .ColData(mconintcol��Ʊ��) = 4
                .ColData(mconintcol��Ʊ����) = 4
                .ColData(mconIntCol��Ʊ����) = 2

                txtProvider.Enabled = True
                cmdProvider.Enabled = True

                If mint��¼״̬ <> 1 Then
                    txtProvider.Enabled = False
                    cmdProvider.Enabled = False
                End If
            ElseIf mint�༭״̬ = 6 Then
                For intCol = 0 To .Cols - 1
                    .ColData(intCol) = 5
                Next
                mshBill.ColData(mconIntColҩ��) = 0
                mshBill.ColData(mconIntCol��������) = 4
                txtժҪ.Enabled = True
            ElseIf mint�༭״̬ = 9 Then
                .ColData(mconIntCol�ɱ���) = IIf(InStr(1, mstrControlItem, ",�ɱ���,") > 0, 4, 5)
                .ColData(mconIntCol�ɹ���) = IIf(InStr(1, mstrControlItem, ",�ɹ���,") > 0, 4, 5)
                .ColData(mconIntCol�ۼ�) = IIf(InStr(1, mstrControlItem, ",�ۼ�,") > 0, 4, 5)
                .ColData(mconIntCol����) = IIf(InStr(1, mstrControlItem, ",����,") > 0, 4, 5)
                .ColData(mconIntCol�ɱ����) = IIf(InStr(1, mstrControlItem, ",�ɱ����,") > 0, 4, 5)
                .ColData(mconIntCol���) = IIf(InStr(1, mstrControlItem, ",���,") > 0, 1, 5)
                .ColData(mconIntCol���ս���) = IIf(InStr(1, mstrControlItem, ",���ս���,") > 0, 1, 5)
                .ColData(mconintcol��Ʊ��) = IIf(InStr(1, mstrControlItem, ",��Ʊ��,") > 0, 4, 5)
                .ColData(mconintcol��Ʊ����) = IIf(InStr(1, mstrControlItem, ",��Ʊ����,") > 0, 4, 5)
                .ColData(mconIntCol��Ʊ����) = IIf(InStr(1, mstrControlItem, ",��Ʊ����,") > 0, 2, 5)
                .ColData(mconintcol��Ʊ���) = IIf(InStr(1, mstrControlItem, ",��Ʊ���,") > 0, 4, 5)

                'Modifed by ZYB 20050104
                .ColData(mconIntColָ��������) = IIf(mbln�޸�������, 4, 0)
                'Modifed by ZYB 20050104 END
'                .ColData(mconintcol���) = 4
'                .ColData(mconintcol��Ʒ�ϸ�֤) = 4
                txtժҪ.Enabled = True
            End If
            
            If mbln�˻� Then
                txtProvider.Enabled = False
                cmdProvider.Enabled = False
            End If
        Else
            If mint�༭״̬ = 7 Then
                txtProvider.Enabled = False
                cmdProvider.Enabled = False
                txtժҪ.Enabled = False
                cboStock.Enabled = False
                
                For intCol = 0 To .Cols - 1
                    .ColData(intCol) = 5
                Next
                .ColData(mconIntCol�ɱ���) = IIf(InStr(1, mstrControlItem, ",�ɱ���,") > 0, 4, 5)
                .ColData(mconIntCol�ɹ���) = IIf(InStr(1, mstrControlItem, ",�ɹ���,") > 0, 4, 5)
                .ColData(mconIntCol�ۼ�) = IIf(InStr(1, mstrControlItem, ",�ۼ�,") > 0, 4, 5)
                .ColData(mconIntCol����) = IIf(InStr(1, mstrControlItem, ",����,") > 0, 4, 5)
                .ColData(mconIntCol�ɱ����) = IIf(InStr(1, mstrControlItem, ",�ɱ����,") > 0, 4, 5)
                .ColData(mconIntCol���) = IIf(InStr(1, mstrControlItem, ",���,") > 0, 1, 5)
                .ColData(mconIntCol���ս���) = IIf(InStr(1, mstrControlItem, ",���ս���,") > 0, 1, 5)
                .ColData(mconintcol��Ʊ��) = IIf(InStr(1, mstrControlItem, ",��Ʊ��,") > 0, 4, 5)
                .ColData(mconintcol��Ʊ����) = IIf(InStr(1, mstrControlItem, ",��Ʊ����,") > 0, 4, 5)
                .ColData(mconIntCol��Ʊ����) = IIf(InStr(1, mstrControlItem, ",��Ʊ����,") > 0, 2, 5)
                .ColData(mconintcol��Ʊ���) = IIf(InStr(1, mstrControlItem, ",��Ʊ���,") > 0, 4, 5)

'                .LocateCol = mconIntCol�ɱ���
                Exit Sub
            ElseIf mint�༭״̬ = 8 Or mbln�˻� Then
                .ColData(mconIntCol����) = 5
                .ColData(mconIntCol��������) = 5
                .ColData(mconIntColЧ��) = 5
                .ColData(mconIntCol����) = 5
                .ColData(mconIntColָ��������) = IIf(mbln�޸�������, 4, 5)
                '.ColData(mconIntCol�ɱ���) = 5
                .ColData(mconIntCol�ɱ����) = 5
                If mbln�˻� Then
                    txtProvider.Enabled = False
                    cmdProvider.Enabled = False
                End If
                '�˻���������ѡ��ⷿ
                cboStock.Enabled = False
                Exit Sub
            End If
            .ColData(0) = 5
            .ColData(mconIntColҩ��) = 1
            .ColData(mconIntCol���) = 5
            .ColData(mconIntCol���) = 5
            If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                .ColData(mconIntCol����) = 1
                .ColData(mconIntColԭ����) = 1
            Else
                .ColData(mconIntCol����) = 5
                .ColData(mconIntColԭ����) = 5
            End If
            .ColData(mconIntCol��λ) = 5
            .ColData(mconIntCol����) = 4
            .ColData(mconIntCol��������) = 2
            .ColData(mconIntColЧ��) = 5
            .ColData(mconIntCol����) = 4
            
            .ColData(mconIntCol�ۼ�) = 5
            .ColData(mconIntCol�ۼ۽��) = 5
            .ColData(mconintCol���) = 5
            
            .ColData(mconintcol��Ʊ��) = 4
            .ColData(mconintcol��Ʊ����) = 4
            .ColData(mconIntCol��Ʊ����) = 2
            
            .ColData(mconIntColָ��������) = IIf(mbln�޸�������, 4, 5)
            .ColData(mconIntColԭ������) = 5
            .ColData(mconIntColԭ����) = 5
            .ColData(mconintcol����) = 5
            .ColData(mconIntCol����ϵ��) = 5
            .ColData(mconIntCol��׼�ĺ�) = 4
            .ColData(mconIntCol�����־) = 5

            .ColData(mconIntCol�ɱ���) = 4
            .ColData(mconIntCol�ɹ���) = 4
            .ColData(mconIntCol�ɱ����) = 4
            .ColData(mconIntCol����) = 4
            .ColData(mconintcol��Ʊ���) = 4
            
            .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
            .ColAlignment(mconIntCol���) = flexAlignLeftCenter
            .ColAlignment(mconIntCol����) = flexAlignLeftCenter
            .ColAlignment(mconIntColԭ����) = flexAlignLeftCenter
            .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
            .ColAlignment(mconIntCol����) = flexAlignLeftCenter
            .ColAlignment(mconIntCol��������) = flexAlignLeftCenter
            .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
            .ColAlignment(mconIntCol����) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ɱ���) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ɱ����) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
            .ColAlignment(mconintCol���) = flexAlignRightCenter
            .ColAlignment(mconIntCol����) = flexAlignRightCenter
            .ColAlignment(mconintcol��Ʊ��) = flexAlignLeftCenter
            .ColAlignment(mconintcol��Ʊ����) = flexAlignLeftCenter
            .ColAlignment(mconIntCol��Ʊ����) = flexAlignLeftCenter
            .ColAlignment(mconintcol��Ʊ���) = flexAlignRightCenter
            .ColAlignment(mconIntCol�����־) = flexAlignLeftCenter
            
            cboStock.Enabled = True
           
            txtProvider.Enabled = True
            cmdProvider.Enabled = True
            txtժҪ.Enabled = True
        End If
    End With
End Sub


Private Sub initGrid()
    '����ʼ������ʼ��ժҪ�ı���ĳ���
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        Call SetColumnByUserDefine
        .TextMatrix(0, mconIntCol�к�) = ""
        .TextMatrix(0, mconIntColҩ��) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mconIntCol��Դ) = "ҩƷ��Դ"
        .TextMatrix(0, mconIntCol����ҩ��) = "����ҩ��"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntColҩ�ۼ���) = "ҩ�ۼ���"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol����) = "������"
        .TextMatrix(0, mconIntColԭ����) = "ԭ����"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntColЧ��) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol�ɱ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol�ɱ����) = "�ɱ����"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconintCol���) = "���"
        .TextMatrix(0, mconintCol���ۼ�) = "���ۼ�"
        .TextMatrix(0, mconintCol���۵�λ) = "���۵�λ"
        .TextMatrix(0, mconintCol���۽��) = "���۽��"
        .TextMatrix(0, mconintCol���۲��) = "���۲��"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mconintcol�������) = "�������"
        .TextMatrix(0, mconintcol�������) = "�������"
        .TextMatrix(0, mconintcol��Ʊ��) = "��Ʊ��"
        .TextMatrix(0, mconintcol��Ʊ����) = "��Ʊ����"
        .TextMatrix(0, mconIntCol��Ʊ����) = "��Ʊ����"
        .TextMatrix(0, mconintcol��Ʊ���) = "��Ʊ���"
        .TextMatrix(0, mconIntColָ��������) = "�ɹ��޼�"
        .TextMatrix(0, mconIntColԭ������) = "ԭ������"
        .TextMatrix(0, mconIntColԭ����) = "ԭЧ��"
        .TextMatrix(0, mconintcol����) = "����"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol���ս���) = "���ս���"
        .TextMatrix(0, mconintcol��Ʒ�ϸ�֤) = "��Ʒ�ϸ�֤"
        .TextMatrix(0, mconIntCol�ɹ���) = "�ɹ���"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol�Ƿ�����) = "�Ƿ�����"
        .TextMatrix(0, mconIntcol�ӳ���) = "�ӳ���"
        .TextMatrix(0, mconIntColҩƷ���������) = "ҩƷ���������"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntCol�����־) = "�����־"
        .TextMatrix(0, mconIntCol�ƻ�id) = "�ƻ�id"
                
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol�к�) = 300
        .ColWidth(mconIntColҩ��) = 2000
        .ColWidth(mconIntCol��Ʒ��) = 2000
        .ColWidth(mconIntCol��Դ) = 900
        .ColWidth(mconIntCol����ҩ��) = 900
        .ColWidth(mconIntCol���) = 0
        .ColWidth(mconIntColҩ�ۼ���) = 1200
        .ColWidth(mconIntCol���) = 900
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntColԭ����) = 0
        .ColWidth(mconIntCol��λ) = 500
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntCol��������) = 1000
        .ColWidth(mconIntColЧ��) = 1000
        .ColWidth(mconIntCol����) = 1100
        .ColWidth(mconIntCol��������) = IIf(mint�༭״̬ = 6, 1100, 0)
        .ColWidth(mconIntCol����) = 0
        .ColWidth(mconIntCol�ɱ���) = 1000
        .ColWidth(mconIntCol�ɱ����) = 900
        .ColWidth(mconIntCol�ۼ�) = 1000
        .ColWidth(mconIntCol�ۼ۽��) = 900
        .ColWidth(mconintCol���) = 800
        .ColWidth(mconintCol���ۼ�) = IIf(mintUnit = mconint�ۼ۵�λ, 0, 1000)
        .ColWidth(mconintCol���۵�λ) = IIf(mintUnit = mconint�ۼ۵�λ, 0, 1000)
        .ColWidth(mconintCol���۽��) = IIf(mintUnit = mconint�ۼ۵�λ, 0, 1000)
        .ColWidth(mconintCol���۲��) = IIf(mintUnit = mconint�ۼ۵�λ, 0, 1000)
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntCol��׼�ĺ�) = 1000
        .ColWidth(mconintcol�������) = 1200
        .ColWidth(mconintcol�������) = 1000
        .ColWidth(mconintcol��Ʊ��) = 800
        .ColWidth(mconintcol��Ʊ����) = 1000
        .ColWidth(mconIntCol��Ʊ����) = 1000
        .ColWidth(mconintcol��Ʊ���) = 900
        .ColWidth(mconIntColָ��������) = 1000
        .ColWidth(mconIntColԭ������) = 0
        .ColWidth(mconIntColԭ����) = 0
        .ColWidth(mconintcol����) = 0
        .ColWidth(mconIntCol����ϵ��) = 0
        .ColWidth(mconIntCol���) = 1000
        .ColWidth(mconIntCol���ս���) = 4500
        .ColWidth(mconintcol��Ʒ�ϸ�֤) = 1000
        .ColWidth(mconIntCol�ɹ���) = 1000
        .ColWidth(mconIntCol��������) = 0
        .ColWidth(mconIntCol�Ƿ�����) = 0
        .ColWidth(mconIntcol�ӳ���) = 1000
        .ColWidth(mconIntColҩƷ���������) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        If mint�༭״̬ = 6 Then
            .ColWidth(mconIntCol�����־) = 800
        Else
            .ColWidth(mconIntCol�����־) = 0
        End If
        .ColWidth(mconIntCol�ƻ�id) = 0
                
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mconIntCol�к�) = 5
        .ColData(mconIntColҩ��) = 1
        .ColData(mconIntCol��Ʒ��) = 5
        .ColData(mconIntCol��Դ) = 5
        .ColData(mconIntCol����ҩ��) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntColҩ�ۼ���) = 5
        .ColData(mconIntCol���) = 5
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            .ColData(mconIntCol����) = 1
            .ColData(mconIntColԭ����) = 1
        Else
            .ColData(mconIntCol����) = 5
            .ColData(mconIntColԭ����) = 5
        End If
        .ColData(mconIntCol��λ) = 5
        .ColData(mconIntCol����) = 4
        .ColData(mconIntCol��������) = 2
        .ColData(mconIntColЧ��) = 5
        .ColData(mconIntCol����) = 4
        .ColData(mconIntCol��������) = 5
        .ColData(mconIntCol����) = 5
        
        .ColData(mconIntCol�ۼ�) = 5
        .ColData(mconIntCol�ۼ۽��) = 5
        .ColData(mconintCol���) = 5
        
        .ColData(mconintCol���ۼ�) = 5
        .ColData(mconintCol���۵�λ) = 5
        .ColData(mconintCol���۽��) = 5
        .ColData(mconintCol���۲��) = 5
        
        .ColData(mconIntCol��׼�ĺ�) = 5
        .ColData(mconintcol�������) = 4
        .ColData(mconintcol�������) = 2
        .ColData(mconintcol��Ʊ��) = 4
        .ColData(mconintcol��Ʊ����) = 4
        .ColData(mconIntCol��Ʊ����) = 2
        
        .ColData(mconIntColָ��������) = IIf(mbln�޸�������, 4, 5)
        .ColData(mconIntColԭ������) = 5
        .ColData(mconIntColԭ����) = 5
        .ColData(mconintcol����) = 5
        .ColData(mconIntCol����ϵ��) = 5
        .ColData(mconIntCol���) = 1
        .ColData(mconIntCol���ս���) = 1
        .ColData(mconintcol��Ʒ�ϸ�֤) = 4
        .ColData(mconIntCol�Ƿ�����) = 5
        .ColData(mconIntcol�ӳ���) = 5
        .ColData(mconIntColҩƷ���������) = 5
        .ColData(mconIntColҩƷ����) = 5
        .ColData(mconIntColҩƷ����) = 5

        .ColData(mconIntCol�ɱ���) = 4
        .ColData(mconIntCol�ɱ����) = 4
        .ColData(mconIntCol����) = 4
        .ColData(mconintcol��Ʊ���) = 4
        .ColData(mconIntCol�ɹ���) = 4
        
        .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Դ) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����ҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntColҩ�ۼ���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColԭ����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��������) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignRightCenter
        .ColAlignment(mconIntCol��������) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɱ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɱ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol���) = flexAlignRightCenter
        .ColAlignment(mconintCol���ۼ�) = flexAlignRightCenter
        .ColAlignment(mconintCol���۵�λ) = flexAlignRightCenter
        .ColAlignment(mconintCol���۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol���۲��) = flexAlignRightCenter
        .ColAlignment(mconIntCol����) = flexAlignRightCenter
        .ColAlignment(mconIntCol��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mconintcol�������) = flexAlignLeftCenter
        .ColAlignment(mconintcol�������) = flexAlignLeftCenter
        .ColAlignment(mconintcol��Ʊ��) = flexAlignLeftCenter
        .ColAlignment(mconintcol��Ʊ����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʊ����) = flexAlignLeftCenter
        .ColAlignment(mconintcol��Ʊ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ���) = flexAlignRightCenter
        .ColAlignment(mconIntcol�ӳ���) = flexAlignRightCenter
        
        .PrimaryCol = mconIntColҩ��
        .LocateCol = mconIntColҩ��
    End With
    
    Call SetColumnByUserDefine
    txtժҪ.MaxLength = Sys.FieldsLength("ҩƷ�շ���¼", "ժҪ")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.Width < 12735 Then Me.Width = 12735
    
    With Pic����
        
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200

    End With
    

    
    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic����.Width
    End With
    
    
    With mshBill
        .Left = 200
        .Width = Pic����.Width - .Left * 2
    End With
    With txtNO
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    cmdProvider.Left = mshBill.Left + mshBill.Width - cmdProvider.Width
    txtProvider.Left = cmdProvider.Left - txtProvider.Width
    LblProvider.Left = txtProvider.Left - LblProvider.Width - 100
    
    
    With Lbl��������
        .Top = Pic����.Height - 200 - .Height
        .Left = mshBill.Left + 100
    End With
    
    With Txt��������
        .Top = Lbl��������.Top - 60
        .Left = Lbl��������.Left + Lbl��������.Width + 100
    End With
    
    With Lbl������
        .Top = Lbl��������.Top - Lbl��������.Height - 180
        .Left = Lbl��������.Left
    End With
    
    With Txt������
        .Top = Lbl������.Top - 60
        .Left = Txt��������.Left
    End With
    
    With lbl�޸�����
        .Top = Lbl��������.Top
        .Left = Txt��������.Left + Txt��������.Width + 400
    End With
    
    With Txt�޸�����
        .Top = Txt��������.Top
        .Left = lbl�޸�����.Left + lbl�޸�����.Width + 100
    End With
    
    With lbl�޸���
        .Top = Lbl������.Top
        .Left = lbl�޸�����.Left
    End With
    
    With Txt�޸���
        .Top = Txt������.Top
        .Left = Txt�޸�����.Left
    End With

    With Txt�������
        .Top = Txt��������.Top
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With Lbl�������
        .Top = Lbl��������.Top
        .Left = Txt�������.Left - 100 - .Width
    End With
    
    With Txt�����
        .Top = Txt������.Top
        .Left = Txt�������.Left
    End With
    
    With Lbl�����
        .Top = Lbl������.Top
        .Left = Lbl�������.Left
    End With
    
    With txt�˲�����
        .Top = Txt��������.Top
        .Left = Lbl�������.Left - 400 - .Width
    End With
    
    With lbl�˲�����
        .Top = Lbl��������.Top
        .Left = txt�˲�����.Left - .Width - 100
    End With
    
    With txt�˲���
        .Top = Txt������.Top
        .Left = txt�˲�����.Left
    End With
    
    With lbl�˲���
        .Top = Lbl������.Top
        .Left = lbl�˲�����.Left
    End With
    
    With txtժҪ
        .Top = Lbl������.Top - 140 - .Height
        .Left = Txt������.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With
    
    With lblժҪ
        .Top = txtժҪ.Top + 50
        .Left = txtժҪ.Left - .Width - 180
    End With
    
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 4
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnViewCost = False Then
        lblPurchasePrice.Visible = False
        lblDifference.Visible = False
    End If
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With CmdCancel
        .Left = Pic����.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 100
    End With
    
    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdAllCls
        .Left = CmdSave.Left - .Width - 500
        .Top = CmdCancel.Top
    End With
    
    With cmdAllSel
        .Left = cmdAllCls.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    
    With cmdHelp
        .Left = Pic����.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With
    
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    cmdAddProducer.Visible = (mint�༭״̬ = 1 Or mint�༭״̬ = 2)
    With cmdAddProducer
        .Top = cmdFind.Top
        .Left = IIf(lblCode.Visible, txtCode.Left + txtCode.Width + 50, cmdFind.Left + cmdFind.Width + 50)
    End With
    
    If mint�༭״̬ = 5 Then '�޸ķ�Ʊ��Ϣ�ð�ť�ſ���
        With cmdCopy
            cmdCopy.Visible = True
            .Top = cmdFind.Top
            If txtCode.Visible Then
                .Left = txtCode.Left + txtCode.Width + 100
            Else
                .Left = cmdFind.Left + cmdFind.Width + 100
            End If
        End With
        
        With cmdALLDel
            .Visible = True
            .Left = cmdCopy.Left + cmdCopy.Width + 100
            .Top = cmdCopy.Top
        End With
    End If
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
    Me.chkת���ƿ�.Top = txtCode.Top
    Me.cboEnterStock.Top = txtCode.Top
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�⹺������", "ҩƷ������ʾ��ʽ", mintDrugNameShow)
    
    If msh����.Visible = True Then
        msh����.Visible = False
        mshBill.SetFocus
        mshBill.Col = mconIntCol����
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS  'ж�����ݼ�
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
    mlng��ҩ��λID = 0
    Call ReleaseSelectorRS 'ж�����ݼ�
End Sub

Private Function SaveCheck(Optional ByVal strNo As String = "", Optional ByVal blnTrans As Boolean = False) As Boolean
    'blnTrans:��ʾ�Ƿ�ʼ�������������
    mblnSave = False
    SaveCheck = False
    
    Dim n As Integer
    Dim m As Integer
    Dim dbl�ϼ����� As Double
    Dim lngҩƷID As Long
    Dim lng���� As Long
    Dim strҩƷ As String
    Dim intNumCol As Integer
    
    '�����,ֻ�г���ҵ��ż�飬��������˿�,������ʱ�ۿ�治��Ҳ��ֹ
    If mint�༭״̬ = 6 Or mint�༭״̬ = 8 Or mbln�˻� = True Then
        If mint�༭״̬ = 6 Then
            intNumCol = mconIntCol��������
        Else
            intNumCol = mconIntCol����
        End If
        strҩƷ = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol����, intNumCol, mconIntCol����ϵ��, IIf(mint�༭״̬ = 3, IIf(mbln�˻� = True, 3, 1), 3), , mintNumberDigit)
        If strҩƷ <> "" Then
            If mbln��ʾ��ʽ = False Then
                If mint����� = 1 Then '��������
                    If MsgBox("ҩƷ��" & strҩƷ & "����治�㣬�Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                ElseIf mint����� = 2 Then '�����ֹ
                    MsgBox "ҩƷ��" & strҩƷ & "����治�㣬������ˣ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If
    
    gstrSQL = "zl_ҩƷ�⹺_Verify('" & IIf(mint�༭״̬ = 7, strNo, txtNO.Tag) & "'," & IIf(mint�༭״̬ = 7, "'" & txtNO.Tag & "'", "Null") & ",'" & UserInfo.�û����� & "',to_date('" & Format(mstr�������, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'))"
    On Error GoTo errHandle
    If blnTrans Then gcnOracle.BeginTrans
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
    
    If blnTrans Then gcnOracle.CommitTrans
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Private Sub mnuColDrug_Click(Index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(Index).Checked = True
        
        Call SetDrugName(Index)
    End With
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol�к�, Row)
    mshBill.Value = Format(Sys.Currentdate, "YYYY-MM-DD")
    mshBill.TextMatrix(Row, mconIntCol�Ƿ�����) = "��"
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call ��ʾ�ϼƽ��
    Call RefreshRowNO(mshBill, mconIntCol�к�, mshBill.Row)
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "345679", mint�༭״̬) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("��ȷʵҪɾ������ҩƷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
            
End Sub

Private Sub mshbill_CommandClick()
    Dim strҩƷID As String
    Dim i As Integer
    Dim intRow As Integer
    Dim intOldRow As Integer
    
    On Error GoTo errHandle
    intOldRow = mshBill.Row
    Select Case mshBill.Col
    Case mconIntColҩ��
        Dim RecReturn As Recordset
        
        mblnChange = True
        mshBill.CmdEnable = False
'        Set RecReturn = FrmҩƷѡ����.ShowME(Me, IIf(mint�༭״̬ = 8 Or mbln�˻�, 2, 1), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , True, True, False, False, True, IIf(mint�༭״̬ = 8 Or mbln�˻�, Val(txtProvider.Tag), 0))
        If grsMaster.State = adStateClosed Then
            Call SetSelectorRS(IIf(mint�༭״̬ = 8 Or mbln�˻�, 2, 1), MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , IIf(mint�༭״̬ = 8 Or mbln�˻�, Val(txtProvider.Tag), 0))
        End If
        Set RecReturn = frmSelector.ShowME(Me, 0, IIf(mint�༭״̬ = 8 Or mbln�˻�, 2, 1), , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , IIf(mint�༭״̬ = 8 Or mbln�˻�, Val(txtProvider.Tag), 0), True, True, True, , , mstrPrivs)
        If RecReturn.RecordCount > 0 And (mint�༭״̬ = 8 Or mbln�˻� = True) Then
            Set RecReturn = CheckRedo(RecReturn) '����ظ���¼�����ظ��ļ�¼���˵�Ȼ�󷵻ع��˺�����ݼ�
        End If
        
        mshBill.CmdEnable = True
        
        mshBill.Redraw = False
        If RecReturn.RecordCount > 0 Then
            
            RecReturn.MoveFirst
            For i = 1 To RecReturn.RecordCount
                With mshBill
                    .Redraw = False
                    mlng��װϵ�� = Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ)
                    intRow = .Row
                    .TextMatrix(intRow, mconIntCol�к�) = .Row
                    SetColValue .Row, RecReturn!ҩƷID, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                        nvl(RecReturn!ҩƷ��Դ), "" & RecReturn!����ҩ��, _
                        IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                        Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                        IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�) * Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                        RecReturn!ָ�������� * Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                        IIf(IsNull(RecReturn!����), "!", RecReturn!����), RecReturn!���Ч��, "", _
                        Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, _
                        RecReturn!ҩ������, RecReturn!�ӳ��� / 100, IIf(IsNull(RecReturn!��������), "", Format(RecReturn!��������, "yyyy-mm-dd")), _
                        RecReturn!�ۼ۵�λ, RecReturn!ԭ����
                    If .TextMatrix(.Row, mconIntColԭ������) = "!" Then
                        .Col = mconIntCol����
                    Else
                        .Col = mconIntCol����
                    End If
                                            
                    If (.TextMatrix(intRow, 0) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                        .rows = .rows + 1
                    End If

                    .Row = .rows - 1
                    RecReturn.MoveNext
                End With
            Next
            mshBill.Row = intOldRow
            RecReturn.Close
        End If
        mshBill.Redraw = True
    Case mconIntCol����
        Dim rsProvider As Recordset
        Dim vRect As RECT, blnCancel As Boolean
        vRect = zlControl.GetControlRect(mshBill.MsfObj.hWnd)
        
        gstrSQL = "Select ���� as id,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null Order By ����"
        Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
        True, vRect.Left + 8200, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
        
        If rsProvider Is Nothing Then
            Exit Sub
        End If
        If Not rsProvider.EOF Then
            mshBill.TextMatrix(mshBill.Row, mconIntCol����) = rsProvider!����
            gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
                        Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mconIntCol����), mshBill.TextMatrix(mshBill.Row, 0))
            If Not rsProvider.EOF Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol��׼�ĺ�) = IIf(IsNull(rsProvider!��׼�ĺ�), "", rsProvider!��׼�ĺ�)
            Else
                mshBill.TextMatrix(mshBill.Row, mconIntCol��׼�ĺ�) = ""
            End If
        End If
    Case mconIntColԭ����
        Dim vRects As RECT, blnCancels As Boolean
        vRects = zlControl.GetControlRect(mshBill.MsfObj.hWnd)
        
        gstrSQL = "Select ���� as id,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null Order By ����"
        Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
        True, vRects.Left + 9000, vRects.Top, 300, blnCancels, False, True, gstrNodeNo)
        
        If rsProvider Is Nothing Then
            Exit Sub
        End If
        If Not rsProvider.EOF Then
            mshBill.TextMatrix(mshBill.Row, mconIntColԭ����) = rsProvider!����
        End If
    Case mconIntCol���
        Dim rs��� As New Recordset
                    
        gstrSQL = "Select ����,����,���� From ҩƷ��� Order By ����"
        Set rs��� = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ���")
                
        If rs���.EOF Then
            rs���.Close
            Exit Sub
        End If
        With FrmSelect
            Set .TreeRec = rs���
            .StrNode = "����ҩƷ���"
            .lngMode = 1
            .Show 1, Me
            If .BlnSuccess = True Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol���) = .CurrentName
                If mconIntCol��� <> mintLastCol And mconIntCol��� < mconintcol��Ʊ�� Then
                    mshBill.Col = mconintcol��Ʊ��
                End If
            End If
        End With
        Unload FrmSelect
    Case mconIntCol���ս���
        Dim rs���ս��� As New Recordset
                    
        gstrSQL = "Select ����,���� From ������ս��� Order By ����"
        Set rs���ս��� = zlDataBase.OpenSQLRecord(gstrSQL, "������ս���")
                
        If rs���ս���.EOF Then
            rs���ս���.Close
            Exit Sub
        End If
        With FrmSelect
            Set .TreeRec = rs���ս���
            .StrNode = "�������ս���"
            .lngMode = 1
            .Show 1, Me
            If .BlnSuccess = True Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol���ս���) = .CurrentName
                If mconIntCol���ս��� <> mintLastCol And mconIntCol���ս��� < mconintcol��Ʊ�� Then
                    mshBill.Col = mconintcol��Ʊ��
                End If
            End If
        End With
        Unload FrmSelect
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mconIntCol���� Or .Col = mconIntCol�������� Or .Col = mconIntCol�ɹ��� Or .Col = mconIntCol�ɱ��� Or .Col = mconIntCol�ɱ���� Or .Col = mconintcol��Ʊ��� Or .Col = mconIntCol�ۼ� Or .Col = mconintCol���ۼ� Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mconIntCol����, mconIntCol��������
                    intDigit = mintNumberDigit
                Case mconIntCol�ɹ���, mconIntCol�ɱ���
                   intDigit = mintCostDigit
                Case mconIntCol�ɱ����, mconintcol��Ʊ���
                    intDigit = mintMoneyDigit
                Case mconIntCol�ۼ�
                    intDigit = mintPriceDigit
                Case mconintCol���ۼ�
                    intDigit = gtype_UserDrugDigits.Digit_���ۼ� '���ۼ۱�������С��λ����˰������λ������ʾ������
            End Select
            
            If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                KeyAscii = 0
                Exit Sub
            End If
            
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                If .SelLength = Len(strKey) Then Exit Sub
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    Dim lngRow As Long
    Dim strxq As String
    Dim dbl�ۼ� As Double
    Dim dblLeft As Double
    Dim dblTop As Double
    
    If mint�༭״̬ = 5 And Trim(mshBill.TextMatrix(mshBill.Row, mconintcol��Ʊ��)) <> "" Then
        cmdCopy.Enabled = True
    Else
        cmdCopy.Enabled = False
    End If
    
    
    If mint�༭״̬ = 8 Then
        cmdGetInputCost.Visible = False
        picInputCost.Visible = False
    End If
    If Not mblnEnter Then Exit Sub
    
    If Trim(txtProvider.Text) = "" And (mint�༭״̬ = 8 Or mbln�˻�) Then
        If mblnMSH_GetFocus Then
            mblnMSH_GetFocus = False
            MsgBox "����ѡ��Ӧ�̣�", vbInformation, gstrSysName
        End If
        SendMessage txtProvider.hWnd, 7, 0, 0   'ֱ����txtprovider.setfocus�ᱨ��
        Exit Sub
    End If
    
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        If .Row <> .LastRow Then
            lngRow = .LastRow
            If PicInput.Visible Then
                '���¼������ۼۡ����
                dbl�ۼ� = Val(.TextMatrix(lngRow, mconIntCol�ɱ���)) * (1 + (Val(Txt�Ӽ���) / 100))
                .TextMatrix(lngRow, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(lngRow, 0)), Val(.TextMatrix(lngRow, mconIntCol�ɱ���)), Val(Txt�Ӽ���) / 100, dbl�ۼ�, lngRow), mintPriceDigit, , True)
                .TextMatrix(lngRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconIntCol�ۼ�)) * Val(.TextMatrix(lngRow, mconIntCol����)), mintMoneyDigit, , True)
                .TextMatrix(lngRow, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(lngRow, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(lngRow, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(lngRow, mconIntCol�ɱ����) = "", 0, .TextMatrix(lngRow, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                PicInput.Visible = False
            End If
        End If
        SetInputFormat .Row
        
        'Modified by zyb 2002-10-30
        If mbln�����ֹ�����ӳ��� = False Then
            PicInput.Visible = False
        ElseIf PicInput.Visible = True Then
            If Txt�Ӽ���.Visible And Txt�Ӽ���.Enabled Then
                Txt�Ӽ���.SetFocus
            End If
            Exit Sub
        End If
                
        Select Case .Col
            Case mconIntColҩ��
                .txtCheck = False
                .MaxLength = 40
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
                
            Case mconIntCol����
                OS.OpenIme True
    
                .txtCheck = False
                .MaxLength = mlng�����̳���
                .TxtSetFocus
                
            Case mconIntColԭ����
                OS.OpenIme True
    
                .txtCheck = False
                .MaxLength = mlngԭ���س���
                .TxtSetFocus
                
            Case mconIntCol����
                .txtCheck = False
                '.TextMask = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
                .MaxLength = mintBatchNoLen
            Case mconIntCol��������
                .txtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mconIntCol����) <> "" And Len(.TextMatrix(.Row, mconIntCol����)) = 8 Then
                    strxq = UCase(.TextMatrix(.Row, mconIntCol����))
                    If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                        If IsNumeric(strxq) Then
                            If Trim(.TextMatrix(.Row, mconIntCol��������)) = "" Then
                                strxq = TranNumToDate(strxq)
                                If Trim(strxq) = "" Then Exit Sub
                                .TextMatrix(.Row, mconIntCol��������) = Format(strxq, "yyyy-mm-dd")
                            End If
                         End If
                    End If
                End If
            Case mconIntColЧ��
                .txtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If Trim(.TextMatrix(.Row, mconIntColԭ����)) = "" Then
                    Exit Sub
                End If
                If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(0) = "0" Then
                    Exit Sub
                End If
                If .TextMatrix(.Row, mconIntCol��������) <> "" Then
                    If Trim(.TextMatrix(.Row, mconIntColЧ��)) = "" Then
                        strxq = UCase(.TextMatrix(.Row, mconIntCol��������))
                    End If
                ElseIf .TextMatrix(.Row, mconIntCol����) <> "" And Len(.TextMatrix(.Row, mconIntCol����)) = 8 Then
                    strxq = UCase(.TextMatrix(.Row, mconIntCol����))
                    If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                        If IsNumeric(strxq) Then
                            If Trim(.TextMatrix(.Row, mconIntColЧ��)) = "" Then
                                strxq = TranNumToDate(strxq)
                            Else
                                Exit Sub
                            End If
                        Else
                            strxq = ""
                        End If
                    Else
                        strxq = ""
                    End If
                End If
                If Trim(strxq) = "" Then Exit Sub
                
                .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntColԭ����), "||")(0), strxq), "yyyy-mm-dd")
                
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 Then
                    '����Ϊ��Ч��
                    .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntColЧ��)), "yyyy-mm-dd")
                End If
                
                Call CheckLapse(.TextMatrix(.Row, mconIntColЧ��))
            Case mconIntCol����
                .txtCheck = True
                .MaxLength = 5
                .TextMask = ".1234567890"
                staThis.Panels.Item(2) = .TextMatrix(.Row, mconIntColҩ��) & "��ָ��������Ϊ��" & .TextMatrix(.Row, mconIntColָ��������)
                
                If mint�༭״̬ = 7 Then
                    Call SetState
                End If
            Case mconIntCol�ɱ���, mconIntColָ��������, mconIntCol�ɹ���, mconintCol���ۼ�
                .txtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
                If mint�༭״̬ = 7 Then
                    Call SetState
                ElseIf mint�༭״̬ = 8 And .Col = mconIntCol�ɹ��� Then
                    cmdGetInputCost.Visible = True
                    dblLeft = mshBill.Left + mshBill.MsfObj.CellLeft + mshBill.MsfObj.CellWidth - cmdGetInputCost.Width + 20
                    dblTop = mshBill.Top + mshBill.MsfObj.CellTop
                    cmdGetInputCost.Top = dblTop
                    cmdGetInputCost.Left = dblLeft
                End If
            Case mconIntCol�ɱ����
                .txtCheck = True
                .MaxLength = 14
                .TextMask = "-.1234567890"
                If mint�༭״̬ = 7 Then
                    Call SetState
                End If
            Case mconIntCol����
                .txtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mconIntCol��������
                .txtCheck = True
                .MaxLength = 11
                If mint�༭״̬ = 6 And mbln�˻� = True Then
                    .TextMask = "-.1234567890"
                Else
                    .TextMask = ".1234567890"
                End If
                
                If .TextMatrix(.Row, mconIntCol�����־) = "��" And mint�༭״̬ = 6 And gtype_UserSysParms.P173_������Ǹ������ܽ��и������ = 1 Then
                    .ColData(mconIntCol��������) = 5
                ElseIf mint�༭״̬ = 6 Then
                    .ColData(mconIntCol��������) = 4
                End If
            Case mconintcol��Ʊ��
                .txtCheck = False
                .MaxLength = 200
                
                If .TextMatrix(.Row, mconIntCol�����־) = "��" And (mint�༭״̬ = 5 Or mint�༭״̬ = 7) And .Col = mconintcol��Ʊ�� And gtype_UserSysParms.P173_������Ǹ������ܽ��и������ = 1 Then
                    .ColData(mconintcol��Ʊ��) = 5
                ElseIf mint�༭״̬ = 5 Then
                    .ColData(mconintcol��Ʊ��) = 4
                End If
            Case mconintcol��Ʊ����
                .txtCheck = True
                .MaxLength = 20
                .TextMask = "1234567890"
                
                If .TextMatrix(.Row, mconIntCol�����־) = "��" And (mint�༭״̬ = 5 Or mint�༭״̬ = 7) And .Col = mconintcol��Ʊ���� And gtype_UserSysParms.P173_������Ǹ������ܽ��и������ = 1 Then
                    .ColData(mconintcol��Ʊ����) = 5
                ElseIf mint�༭״̬ = 5 And Trim(.TextMatrix(.Row, mconintcol��Ʊ��)) <> "" Then
                    .ColData(mconintcol��Ʊ����) = 4
                End If
            Case mconintcol��Ʊ���
                .txtCheck = True
                .MaxLength = 14
                .TextMask = "-.1234567890"
                
                If .TextMatrix(.Row, mconIntCol�����־) = "��" And (mint�༭״̬ = 5 Or mint�༭״̬ = 7) And .Col = mconintcol��Ʊ��� And gtype_UserSysParms.P173_������Ǹ������ܽ��и������ = 1 Then
                    .ColData(mconintcol��Ʊ���) = 5
                ElseIf mint�༭״̬ = 5 Then
                    .ColData(mconintcol��Ʊ���) = 4
                End If
            Case mconIntCol��Ʊ����
                .txtCheck = True
                .TextMask = "1234567890-"
                .Value = Sys.Currentdate
                .MaxLength = 10
                
                If .TextMatrix(.Row, mconIntCol�����־) = "��" And (mint�༭״̬ = 5 Or mint�༭״̬ = 7) And .Col = mconIntCol��Ʊ���� And gtype_UserSysParms.P173_������Ǹ������ܽ��и������ = 1 Then
                    .ColData(mconIntCol��Ʊ����) = 5
                ElseIf mint�༭״̬ = 5 Then
                    .ColData(mconIntCol��Ʊ����) = 2
                End If
            Case mconIntCol���, mconintcol��Ʒ�ϸ�֤, mconIntCol���ս���
                .txtCheck = True
                .MaxLength = 100
            Case mconIntCol�ۼ�
                .txtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
                If mint�༭״̬ = 7 Then
                    Call SetState
                End If
            Case mconIntCol��׼�ĺ�
                .txtCheck = False
                .MaxLength = 40
            Case mconintcol�������
                .txtCheck = False
                .MaxLength = 200
            Case mconintcol�������
                .txtCheck = True
                .TextMask = "1234567890-"
                .Value = Sys.Currentdate
                .MaxLength = 10
                
                If .TextMatrix(.Row, mconintcol�������) <> "" And (mint�༭״̬ = 1 Or mint�༭״̬ = 2) Then
                    .ColData(mconintcol�������) = 2
                Else
                    .ColData(mconintcol�������) = 5
                End If
            End Select
    End With
End Sub

Private Sub mshBill_GotFocus()
    
    With mshBill
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strUnitQuantity As String
    Dim dbl�ӳ��� As Double, dblָ�����ۼ� As Double
    Dim dbl�ɱ��� As Double
    Dim strxq As String
    Dim intRow As Integer
    Dim i As Integer
    Dim strҩƷID As String
    Dim intOldRow As Integer
    Dim dbl�ۼ� As Double
    Dim dblTemp�ۼ� As Double
    Dim strҩƷ As String
    Dim rsProvider As Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsMaxs As New Recordset
    Dim ints���� As Integer, strCodes As String
                    
    intOldRow = mshBill.Row
        
    If KeyCode <> vbKeyReturn Then Exit Sub
    On Error GoTo errHandle
    
    With mshBill
'        .Text = UCase(Trim(.Text))
        strKey = Trim(.Text)
        
        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col
            Case mconIntColҩ��
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    .Redraw = False
                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 3630
                    End If
                    
                    If grsMaster.State = adStateClosed Then '��ȡ���ݼ�
                        Call SetSelectorRS(IIf(mint�༭״̬ = 8 Or mbln�˻�, 2, 1), MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , IIf(mint�༭״̬ = 8 Or mbln�˻�, Val(txtProvider.Tag), 0))
                    End If
'                    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, IIf(mint�༭״̬ = 8 Or mbln�˻�, 2, 1), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , strkey, sngLeft, sngTop, True, True, False, False, True, IIf(mint�༭״̬ = 8 Or mbln�˻�, Val(txtProvider.Tag), 0))
                    Set RecReturn = frmSelector.ShowME(Me, 1, IIf(mint�༭״̬ = 8 Or mbln�˻�, 2, 1), strKey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , IIf(mint�༭״̬ = 8 Or mbln�˻�, Val(txtProvider.Tag), 0), True, True, True, , , mstrPrivs)
                                                        
                    If RecReturn.RecordCount > 0 And (mint�༭״̬ = 8 Or mbln�˻� = True) Then
                        Set RecReturn = CheckRedo(RecReturn) '����ظ���¼ �����ظ���¼��ҩƷid���ػ���
                    End If
'                    If strҩƷid <> "" And mint�༭״̬ = 8 Then
'                        mbln��ʾ = False
'                        Set RecReturn = GetRs(strҩƷid, RecReturn) '�����ظ�������
'                    End If
                                                        
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        
                        For i = 1 To RecReturn.RecordCount
                            mlng��װϵ�� = Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ)
                            
                            intRow = .Row
                            .TextMatrix(intRow, mconIntCol�к�) = .Row
                            If SetColValue(.Row, RecReturn!ҩƷID, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                                 nvl(RecReturn!ҩƷ��Դ), "" & RecReturn!����ҩ��, IIf(IsNull(RecReturn!���), "", RecReturn!���), _
                                 IIf(IsNull(RecReturn!����), "", RecReturn!����), Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                                 IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�) * Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                                 RecReturn!ָ�������� * Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                                 IIf(IsNull(RecReturn!����), "!", RecReturn!����), RecReturn!���Ч��, "", _
                                 Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, _
                                 RecReturn!ҩ������, RecReturn!�ӳ��� / 100, IIf(IsNull(RecReturn!��������), "", Format(RecReturn!��������, "yyyy-mm-dd")), RecReturn!�ۼ۵�λ, RecReturn!ԭ����) = False Then ' RecReturn!����
                                 Cancel = True
                                 Exit Sub
                             End If
                            .Text = .TextMatrix(.Row, .Col)
                            If (.TextMatrix(intRow, 0) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                    Else
                        Cancel = True
                    End If
                    Call ��ʾ�����
                End If
                .Redraw = True
            Case mconIntCol����
                '�޴���
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol����) = ""
                    End If
                    
                    If mconIntCol���� <> mintLastCol And mconIntCol���� < mconIntCol���� Then
'                        .Col = mconIntCol����
                        .Col = GetNextEnableCol(mconIntCol����)
                        Cancel = True
                    End If
                    Exit Sub
                Else
                    vRect = zlControl.GetControlRect(mshBill.MsfObj.hWnd)
                    
                    .Text = UCase(Trim(.Text))
                    strKey = Trim(.Text)
                    
                    If Trim(.Text) = "" Then Exit Sub
                    
                    gstrSQL = "Select ���� as id,���� ,����,���� From ҩƷ������ " _
                            & "Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And (upper(����) like '" & IIf(gstrMatchMethod = "0", "%", "") & strKey & "%' or Upper(����) like '" & IIf(gstrMatchMethod = "0", "%", "") & strKey & "%' or Upper(����) like '" & strKey & "%') " _
                                & "Order By ���� "
                                
                    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "������", False, "", "������ѡ��", False, False, _
                    True, vRect.Left + 8200, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
                    
                    If blnCancel = True Then mshBill.Text = "": .TextMatrix(.Row, mconIntCol����) = "": Exit Sub '��ѡ����ʱ����Esc�������´���
                    
                    If rsProvider Is Nothing Then
                        If MsgBox("ҩƷ������û���ҵ�������������̣���Ҫ��������ҩƷ����������", vbYesNo + vbQuestion, MStrCaption) = vbNo Then
                            mshBill.TextMatrix(mshBill.Row, mconIntCol����) = ""
                            mshBill.Text = ""
                            Cancel = True
                            Exit Sub
                        Else
                            If LenB(StrConv(strKey, vbFromUnicode)) > mlng�����̳��� Then
                                MsgBox "���������ƹ���(���" & mlng�����̳��� & "���ַ���" & Int(mlng�����̳��� / 2) & "������)!", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                            
                            If rsMaxs.State = 1 Then rsMaxs.Close
                            gstrSQL = "SELECT Nvl(MAX(LENGTH(����)),2) As Length FROM ҩƷ������"
                            Set rsMaxs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-ҩƷ�����̱��볤��")
                            ints���� = rsMaxs!length
                            
                            gstrSQL = "SELECT Nvl(MAX(LPAD(����," & ints���� & ",'0')),'00') As Code FROM ҩƷ������"
                            Set rsMaxs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-ҩƷ�����̱���")
                            strCodes = rsMaxs!Code
                            
                            ints���� = Len(strCodes)
                            strCodes = strCodes + 1
                            If ints���� >= Len(strCodes) Then
                                strCodes = String(ints���� - Len(strCodes), "0") & strCodes
                            End If

                            gstrSQL = "ZL_ҩƷ������_INSERT('" & strCodes & "','" & strKey & "',zlSpellCode('" & strKey & "',10))"
                            
                            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        End If
                    Else
                        mshBill.TextMatrix(mshBill.Row, mconIntCol����) = rsProvider!����
                        mshBill.Text = rsProvider!����
                        
                        gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
                        Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mconIntCol����), mshBill.TextMatrix(mshBill.Row, 0))
                        If Not rsProvider.EOF Then
                            mshBill.TextMatrix(mshBill.Row, mconIntCol��׼�ĺ�) = IIf(IsNull(rsProvider!��׼�ĺ�), "", rsProvider!��׼�ĺ�)
                        Else
                            mshBill.TextMatrix(mshBill.Row, mconIntCol��׼�ĺ�) = ""
                        End If
                    End If
                End If
                OS.OpenIme
            Case mconIntColԭ����
                '�޴���
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntColԭ����) = ""
                    End If
                    
                    If mconIntColԭ���� <> mintLastCol And mconIntColԭ���� < mconIntCol���� Then
'                        .Col = mconIntCol����
                        .Col = GetNextEnableCol(mconIntColԭ����)
                        Cancel = True
                    End If
                    Exit Sub
                Else
                
                    vRect = zlControl.GetControlRect(mshBill.MsfObj.hWnd)
                    .Text = UCase(Trim(.Text))
                    strKey = Trim(.Text)
                    
                    gstrSQL = "Select ���� as id,���� ,����,���� From ҩƷ������ " _
                            & "Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And (upper(����) like '" & IIf(gstrMatchMethod = "0", "%", "") & strKey & "%' or Upper(����) like '" & IIf(gstrMatchMethod = "0", "%", "") & strKey & "%' or Upper(����) like '" & strKey & "%') " _
                                & "Order By ���� "
                                
                    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "ԭ����", False, "", "ԭ����ѡ��", False, False, _
                    True, vRect.Left + 9000, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
                    
                    If blnCancel = True Then .Text = "": .TextMatrix(.Row, mconIntColԭ����) = "": Exit Sub '��ѡ����ʱ����Esc�������´���
                    
                    If rsProvider Is Nothing Then
                        If MsgBox("ҩƷ������û���ҵ��������ԭ���أ���Ҫ��������ҩƷ����������", vbYesNo + vbQuestion, MStrCaption) = vbNo Then
                            mshBill.TextMatrix(mshBill.Row, mconIntColԭ����) = ""
                            mshBill.Text = ""
                            Cancel = True
                            Exit Sub
                        Else
                            If LenB(StrConv(strKey, vbFromUnicode)) > mlngԭ���س��� Then
                                MsgBox "ԭ�������ƹ���(���" & mlngԭ���س��� & "���ַ���" & Int(mlngԭ���س��� / 2) & "������)!", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                        
                            If rsMaxs.State = 1 Then rsMaxs.Close
                            gstrSQL = "SELECT Nvl(MAX(LENGTH(����)),2) As Length FROM ҩƷ������"
                            Set rsMaxs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-ҩƷ�����̱��볤��")
                            ints���� = rsMaxs!length
                            
                            gstrSQL = "SELECT Nvl(MAX(LPAD(����," & ints���� & ",'0')),'00') As Code FROM ҩƷ������"
                            Set rsMaxs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-ҩƷ�����̱���")
                            strCodes = rsMaxs!Code
                            
                            ints���� = Len(strCodes)
                            strCodes = strCodes + 1
                            If ints���� >= Len(strCodes) Then
                                strCodes = String(ints���� - Len(strCodes), "0") & strCodes
                            End If
                            
                            gstrSQL = "ZL_ҩƷ������_INSERT('" & strCodes & "','" & strKey & "',zlSpellCode('" & strKey & "',10))"
 
                            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        End If
                    Else
                        mshBill.TextMatrix(mshBill.Row, mconIntColԭ����) = rsProvider!����
                        mshBill.Text = rsProvider!����
                    End If
                End If
                OS.OpenIme
            Case mconIntCol���ս���
                If Trim(.TextMatrix(.Row, mconIntColҩ��)) = "" Then
                    .Text = ""
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol���ս���) = ""
                    End If
                    .Col = GetNextEnableCol(mconIntCol���ս���)
                    Cancel = True
                    Exit Sub
                Else
                    Dim rs���� As New Recordset
                    
                    gstrSQL = "" & _
                        "   Select ����,���� From ������ս��� " & _
                        "   Where upper(����) like [1] or Upper(����) like [1] "
                    Set rs���� = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%")
                    
                    If rs����.EOF Then
                        .TextMatrix(.Row, mconIntCol���ս���) = .Text
                        .Col = GetNextEnableCol(mconIntCol���ս���)
                        Cancel = True
                        Exit Sub
                    Else
                        If rs����.RecordCount = 1 Then
                            .TextMatrix(.Row, mconIntCol���ս���) = rs����.Fields("����")
                            .Text = rs����.Fields("����")
                            .Col = GetNextEnableCol(mconIntCol���ս���)
                        Else
                            Set msh����.Recordset = rs����
                            With msh����
                                .Redraw = False
                                .Left = Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                                .Top = Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight
                                .Visible = True
                                .SetFocus
                                .ColWidth(0) = 800
                                .ColWidth(1) = 5000
                                .Row = 1
                                .Col = 0
                                .TopRow = 1
                                .ColSel = .Cols - 1
                                .Redraw = True
                                Cancel = True
                                Exit Sub
                            End With
                        End If
                    End If
                End If
                OS.OpenIme
            Case mconIntCol����
                '�޴���
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol����) = ""
                    End If
                    If mconIntCol���� <> mintLastCol And mconIntCol���� < mconIntCol�������� Then
                        .Col = mconIntCol��������
                        Cancel = True
                    End If
                    Exit Sub
                End If
            Case mconIntCol��������
                '�д���
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "�Բ����������ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                        .TextMatrix(.Row, mconIntCol��������) = .Text
                        
                        '����Ч��
                        If Trim(.TextMatrix(.Row, mconIntColԭ����)) = "" Then
                            Exit Sub
                        End If
                        If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(0) = "0" Then
                            Exit Sub
                        End If
                        If .TextMatrix(.Row, mconIntCol��������) <> "" Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol��������))
                        ElseIf .TextMatrix(.Row, mconIntCol����) <> "" And Len(.TextMatrix(.Row, mconIntCol����)) = 8 Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol����))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                If IsNumeric(strxq) Then
                                    If Trim(.TextMatrix(.Row, mconIntColЧ��)) = "" Then
                                        strxq = TranNumToDate(strxq)
                                    Else
                                        Exit Sub
                                    End If
                                Else
                                    strxq = ""
                                End If
                            Else
                                strxq = ""
                            End If
                        End If
                        If Trim(strxq) = "" Then Exit Sub
                        
                        .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntColԭ����), "||")(0), strxq), "yyyy-mm-dd")
                        
                        If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 Then
                            '����Ϊ��Ч��
                            .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntColЧ��)), "yyyy-mm-dd")
                        End If
                        
                        Call CheckLapse(.TextMatrix(.Row, mconIntColЧ��))
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "�Բ����������ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mconIntCol��������) Then
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    If .ColData(mconIntColЧ��) = 2 Then
                        If mconIntCol�������� <> mintLastCol And mconIntCol�������� < mconIntColЧ�� Then
                            .Col = mconIntColЧ��
                        End If
                    Else
                        If mconIntCol�������� <> mintLastCol And mconIntCol�������� < mconIntCol���� Then
                            .Col = mconIntCol����
                        End If
                    End If
                    Exit Sub
                End If
            Case mconIntColЧ��
                '�д���
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "�Բ���ʧЧ�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "�Բ���ʧЧ�ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mconIntColЧ��) Then
                
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    
                    Exit Sub
                End If
            Case mconIntCol����
                If Trim(.TextMatrix(.Row, mconIntColҩ��)) = "" Then
                    .Text = ""
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�Բ��𣬿��ʱ���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol����) Then
                    SetDisCount .Row, strKey
                End If
                
                Call ���ɱ���
                Call ��ʾ�ϼƽ��
                Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / IIf(Val(.TextMatrix(.Row, mconIntCol����ϵ��)) = 0, 1, Val(.TextMatrix(.Row, mconIntCol����ϵ��))))
            Case mconIntColָ��������
                If .TxtVisible Then
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "�Բ��𣬲ɹ��޼۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If mintUnit = mconintҩ�ⵥλ Then
                        If Val(strKey) < 0.01 Then
                            MsgBox "�Բ��𣬲ɹ��޼۱������0.01,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    Else
                        If Val(strKey) < 0.001 Then
                            MsgBox "�Բ��𣬲ɹ��޼۱������0.001,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntColָ��������) Then
                        strKey = zlStr.FormatEx(strKey, mintCostDigit, , True)
                        .Text = strKey
                        'Modifed by ZYB 20050104
                        .TextMatrix(.Row, mconIntColָ��������) = .Text
                        'Modifed by ZYB 20050104 END
                        SetDisCount .Row, strKey
                    End If
                    
                    Call ���ɱ���
                    Call ��ʾ�ϼƽ��
                End If
            Case mconIntCol�ɹ���
                If Trim(.TextMatrix(.Row, mconIntColҩ��)) = "" Then
                    .Text = ""
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�Բ��𣬲ɹ��۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "�Բ��𣬲ɹ��۲���Ϊ����,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "�ɹ��۱���С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = zlStr.FormatEx(strKey, 7, , True)
                    .TextMatrix(.Row, .Col) = .Text
                End If
                
                If Val(strKey) > Val(.TextMatrix(.Row, mconIntColָ��������)) Then
                    MsgBox "������Ĳɹ��۴����˲ɹ��޼ۡ�", vbInformation + vbOKOnly, gstrSysName
                End If
                
                '�������ÿ���
                If strKey <> "" Then
                    strKey = zlStr.FormatEx(strKey, mintCostDigit, , True)
                    .Text = strKey
                    .TextMatrix(.Row, mconIntCol�ɹ���) = .Text
                End If
               
                '����ɱ��ۣ��ɱ���=�ɹ���*����
                .TextMatrix(.Row, mconIntCol�ɱ���) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ɹ���)) * Val(.TextMatrix(.Row, mconIntCol����)) / 100, mintCostDigit, , True)
                
                '��ʱ��ҩƷ�Ĵ���
                If strKey <> "" And .TextMatrix(.Row, mconIntColԭ����) <> "" And mint�༭״̬ <> 8 And mbln�˻� = False Then
                    '���ݲ�������ʱ��ҩƷ�ۼ۹�ʽ�гɱ��۵��㷨
                    dbl�ɱ��� = IIf(mintʱ������ۼۼӳɷ�ʽ = 0, Val(.TextMatrix(.Row, mconIntCol�ɱ���)), Val(.TextMatrix(.Row, mconIntCol�ɹ���)))
                    
                    If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 Then
                        '���ۿ��ƣ�ʱ��ҩƷ���ۼ�ֱ�ӵ��ڳɱ���
                        If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ɱ���)), mintPriceDigit, , True)
                            If .TextMatrix(.Row, mconIntCol����) <> "" Then
                                .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                            End If
                        Else
                            '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                            '���ϵͳ����Ϊ�棬����ʾ�û�����Ӽ���
                            If mbln�Ӽ��� And mintʱ������ۼۼӳɷ�ʽ = 1 Then
                                mbln�����ֹ�����ӳ��� = True
                                sngLeft = Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                                sngTop = Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                                If sngTop + 1700 > Screen.Height Then
                                    sngTop = sngTop - mshBill.MsfObj.CellHeight - 1700
                                End If
                                
                                With PicInput
                                    .Top = sngTop
                                    .Left = sngLeft
                                    .Visible = True
                                End With
                                 Txt�Ӽ��� = Val(Replace(.TextMatrix(.Row, mconIntcol�ӳ���), "%", "")) '"15.00000"
                                 .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), dbl�ɱ���, Val(Txt�Ӽ���) / 100, dbl�ɱ��� * (1 + (Val(Txt�Ӽ���) / 100))), mintPriceDigit)
                                 
'                                If Val(.TextMatrix(.Row, mconIntCol�ۼ�)) <> 0 And dbl�ɱ��� <> 0 Then
'                                    Txt�Ӽ��� = zlStr.FormatEx(����ӳ���(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ۼ�)), dbl�ɱ���), 5, , True)
'                                End If
                                Txt�Ӽ���.Tag = Txt�Ӽ���
                                Txt�Ӽ���.SetFocus
                            Else
                                If mintʱ�۷ֶμӳɷ�ʽ = 1 Then
                                    If get�ֶμӳ��ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����ϵ��)), dbl�ɱ���, dbl�ӳ���, dblTemp�ۼ�) = False Then
                                        Cancel = True
                                        .TxtSetFocus
                                        Exit Sub
                                    End If
                                Else
                                    dbl�ӳ��� = Val(Replace(.TextMatrix(.Row, mconIntcol�ӳ���), "%", "")) / 100
                                    dblTemp�ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                                End If
                                If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  'û�й�ѡʱ��ȡ�ϴ��ۼ۲���
                                    .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), dbl�ɱ���, dbl�ӳ���, dblTemp�ۼ�), mintPriceDigit, , True)
                                End If
                                
                                .TextMatrix(.Row, mconIntcol�ӳ���) = zlStr.FormatEx(dbl�ӳ��� * 100, 2) & "%"
                                If .TextMatrix(.Row, mconIntCol����) <> "" Then
                                    .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                                End If
                            End If
                        End If
                    Else
                        '����ҩƷ����ӳ��ʣ���ʵ�����壬����ʾ
                        If Val(.TextMatrix(.Row, mconIntCol�ۼ�)) <> 0 And dbl�ɱ��� <> 0 Then
                            .TextMatrix(.Row, mconIntcol�ӳ���) = zlStr.FormatEx((Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / dbl�ɱ��� - 1) * 100, 2) & "%"
                        End If
                        
                        '���ۿ��ƣ�����ҩƷ�����¼��ĳɱ����Ƿ�����ۼ�
                        If gtype_UserSysParms.P275_���۹���ģʽ = 2 And Val(.TextMatrix(.Row, mconIntCol�ɱ���)) <> Val(.TextMatrix(.Row, mconIntCol�ۼ�)) Then
                            If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                MsgBox "�ö���ҩƷ���������۹���ģʽ�����ɱ���Ӧ���ۼ�(" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�), mintPriceDigit, , True) & ")��ȣ�", vbInformation + vbOKOnly, gstrSysName
                                
                                .TextMatrix(.Row, mconIntCol�ɱ���) = .TextMatrix(.Row, mconIntCol�ۼ�)
                                strKey = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ɱ���)) / (Val(.TextMatrix(.Row, mconIntCol����)) / 100), mintPriceDigit, , True)
                                .TextMatrix(.Row, mconIntCol�ɹ���) = strKey
                                .Text = strKey
                                .TextMatrix(.Row, mconIntcol�ӳ���) = "0%"
                            End If
                        End If
                    End If
                End If
                
                '�˻�ʱ
                If strKey <> "" And (mint�༭״̬ = 8 Or mbln�˻� = True) Then
                    If .TextMatrix(.Row, mconIntColԭ����) <> "" Then
                        If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 Then
                            '���ۿ��ƣ�ʱ��ҩƷ���ۼ�ֱ�ӵ��ڳɱ���
                            If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ɱ���)), mintPriceDigit, , True)
                                If .TextMatrix(.Row, mconIntCol����) <> "" Then
                                    .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                                End If
                            End If
                        Else
                            '���ۿ��ƣ�����ҩƷ�����¼��ĳɱ����Ƿ�����ۼ�
                            If gtype_UserSysParms.P275_���۹���ģʽ = 2 And Val(.TextMatrix(.Row, mconIntCol�ɱ���)) <> Val(.TextMatrix(.Row, mconIntCol�ۼ�)) Then
                                If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                MsgBox "�ö���ҩƷ���������۹���ģʽ�����ɱ���Ӧ���ۼ�(" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�), mintPriceDigit, , True) & ")��ȣ�", vbInformation + vbOKOnly, gstrSysName
                                
                                .TextMatrix(.Row, mconIntCol�ɱ���) = .TextMatrix(.Row, mconIntCol�ۼ�)
                                strKey = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ɱ���)) / (Val(.TextMatrix(.Row, mconIntCol����)) / 100), mintPriceDigit, , True)
                                .TextMatrix(.Row, mconIntCol�ɹ���) = strKey
                                .Text = strKey
                                .TextMatrix(.Row, mconIntcol�ӳ���) = "0%"
                                End If
                            End If
                        End If
                    End If
                 End If
                
                '���ý��
                If strKey <> "" And .TextMatrix(.Row, mconIntCol����) <> "" Then
                    .TextMatrix(.Row, mconIntCol�ɱ����) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * Val(.TextMatrix(.Row, mconIntCol�ɱ���)), mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * Val(.TextMatrix(.Row, mconIntCol�ۼ�)), mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconintcol��Ʊ���) = IIf(Trim(.TextMatrix(.Row, mconintcol��Ʊ��)) = "", "", .TextMatrix(.Row, mconIntCol�ɱ����))
                    .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                End If
                
                Call ���ɱ���
                Call ��ʾ�ϼƽ��
                Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / Val(.TextMatrix(.Row, mconIntCol����ϵ��)))
            Case mconIntCol�ɱ���
                If Trim(.TextMatrix(.Row, mconIntColҩ��)) = "" Then
                    .Text = ""
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�Բ��𣬳ɱ��۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "�Բ��𣬳ɱ��۲���Ϊ����,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "�ɱ��۱���С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = zlStr.FormatEx(strKey, 7, , True)
                    .TextMatrix(.Row, .Col) = .Text
                End If
                
                If strKey <> "" Then
                    strKey = zlStr.FormatEx(strKey, mintCostDigit, , True)
                    .Text = strKey
                    .TextMatrix(.Row, mconIntCol�ɱ���) = .Text
                End If
                          
                If Val(strKey) > Val(.TextMatrix(.Row, mconIntColָ��������)) Then
                    MsgBox "������ĳɱ��۴����˲ɹ��޼ۡ�", vbInformation + vbOKOnly, gstrSysName
                End If
                           
                If Val(.TextMatrix(.Row, mconIntCol����)) = 0 Then
                    .TextMatrix(.Row, mconIntCol����) = "100"
                End If
                
                '������ʣ�����=�ɱ���/�ɹ���
                If Val(.TextMatrix(.Row, mconIntCol�ɹ���)) <> 0 Then
                    .TextMatrix(.Row, mconIntCol����) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ɱ���)) / Val(.TextMatrix(.Row, mconIntCol�ɹ���)) * 100, 7, , True)
                Else
                    .TextMatrix(.Row, mconIntCol����) = "100"
                End If
                
                '��ʱ��ҩƷ�Ĵ���
                If strKey <> "" And .TextMatrix(.Row, mconIntColԭ����) <> "" And mint�༭״̬ <> 8 And mbln�˻� = False Then
                    '���ݲ�������ʱ��ҩƷ�ۼ۹�ʽ�гɱ��۵��㷨
                    dbl�ɱ��� = IIf(mintʱ������ۼۼӳɷ�ʽ = 0, Val(.TextMatrix(.Row, mconIntCol�ɱ���)), Val(.TextMatrix(.Row, mconIntCol�ɹ���)))
                        
                    If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 Then
                        '���ۿ��ƣ�ʱ��ҩƷ���ۼ�ֱ�ӵ��ڳɱ���
                        If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(strKey, mintPriceDigit, , True)
                            If .TextMatrix(.Row, mconIntCol����) <> "" Then
                                .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                            End If
                            
                            .TextMatrix(.Row, mconIntcol�ӳ���) = "0%"
                        Else
                            '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                            '���ϵͳ����Ϊ�棬����ʾ�û�����Ӽ���
                            If mbln�Ӽ��� And mintʱ������ۼۼӳɷ�ʽ = 0 Then
                                mbln�����ֹ�����ӳ��� = True
                                If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  '���δ��ѡȡ�ϴ��ۼۣ��ҹ�ѡ���ֹ�¼��ӳ��ʲ����򵯳��ӳ��ʿ����û�ѡ��
                                    sngLeft = Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                                    sngTop = Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                                    If sngTop + 1700 > Screen.Height Then
                                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 1700
                                    End If
                                    
                                    With PicInput
                                        .Top = sngTop
                                        .Left = sngLeft
                                        .Visible = True
                                    End With
                                    Txt�Ӽ��� = Val(Replace(.TextMatrix(.Row, mconIntcol�ӳ���), "%", "")) '"15.00000"
                                    .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), dbl�ɱ���, Val(Txt�Ӽ���) / 100, dbl�ɱ��� * (1 + (Val(Txt�Ӽ���) / 100))), mintPriceDigit)


'                                    If Val(.TextMatrix(.Row, mconIntCol�ۼ�)) <> 0 And dbl�ɱ��� <> 0 Then
'                                        Txt�Ӽ��� = zlStr.FormatEx(����ӳ���(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ۼ�)), dbl�ɱ���), 5, , True)
'                                    End If
                                    Txt�Ӽ���.Tag = Txt�Ӽ���
                                    Txt�Ӽ���.SetFocus
                                End If
                            Else
                                If mintʱ�۷ֶμӳɷ�ʽ = 1 Then
                                    If get�ֶμӳ��ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����ϵ��)), dbl�ɱ���, dbl�ӳ���, dblTemp�ۼ�) = False Then
                                        Cancel = True
                                        .TxtSetFocus
                                        Exit Sub
                                    End If
                                Else
                                    dbl�ӳ��� = Val(Replace(.TextMatrix(.Row, mconIntcol�ӳ���), "%", "")) / 100
                                    dblTemp�ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                                End If
                                
                                If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  'û�й�ѡʱ��ȡ�ϴ��ۼ۲���
                                    .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), dbl�ɱ���, dbl�ӳ���, dblTemp�ۼ�), mintPriceDigit, , True)
                                End If
                                .TextMatrix(.Row, mconIntcol�ӳ���) = zlStr.FormatEx(dbl�ӳ��� * 100, 2) & "%"
                                If .TextMatrix(.Row, mconIntCol����) <> "" Then
                                    .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                                End If
                            End If
                        End If
                    Else
                        '����ҩƷ����ӳ��ʣ���ʵ�����壬����ʾ
                        If Val(.TextMatrix(.Row, mconIntCol�ۼ�)) <> 0 And dbl�ɱ��� <> 0 Then
                            .TextMatrix(.Row, mconIntcol�ӳ���) = zlStr.FormatEx((Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / dbl�ɱ��� - 1) * 100, 2) & "%"
                        End If
                        
                        '���ۿ��ƣ�����ҩƷ�����¼��ĳɱ����Ƿ�����ۼ�
                        If gtype_UserSysParms.P275_���۹���ģʽ = 2 And Val(strKey) <> Val(.TextMatrix(.Row, mconIntCol�ۼ�)) Then
                            If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                MsgBox "�ö���ҩƷ���������۹���ģʽ�����ɱ���Ӧ���ۼ�(" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�), mintPriceDigit, , True) & ")��ȣ�", vbInformation + vbOKOnly, gstrSysName
                                strKey = .TextMatrix(.Row, mconIntCol�ۼ�)
                                .TextMatrix(.Row, mconIntCol�ɱ���) = zlStr.FormatEx(strKey, mintPriceDigit, , True)
                                .Text = strKey
                                .TextMatrix(.Row, mconIntcol�ӳ���) = "0%"
                                .TextMatrix(.Row, mconIntCol����) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ɱ���)) / Val(.TextMatrix(.Row, mconIntCol�ɹ���)) * 100, 7, , True)
'                                Cancel = True
'                                .TxtSetFocus
'                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                '�˻�ʱ
                If strKey <> "" And (mint�༭״̬ = 8 Or mbln�˻� = True) Then
                    If .TextMatrix(.Row, mconIntColԭ����) <> "" Then
                        If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 Then
                            '���ۿ��ƣ�ʱ��ҩƷ���ۼ�ֱ�ӵ��ڳɱ���
                            If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(strKey, mintPriceDigit, , True)
                                If .TextMatrix(.Row, mconIntCol����) <> "" Then
                                    .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                                End If

                                .TextMatrix(.Row, mconIntcol�ӳ���) = "0%"
                            End If
                        Else
                            '���ۿ��ƣ�����ҩƷ�����¼��ĳɱ����Ƿ�����ۼ�
                            If gtype_UserSysParms.P275_���۹���ģʽ = 2 And Val(strKey) <> Val(.TextMatrix(.Row, mconIntCol�ۼ�)) Then
                                If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                    MsgBox "�ö���ҩƷ���������۹���ģʽ�����ɱ���Ӧ���ۼ�(" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�), mintPriceDigit, , True) & ")��ȣ�", vbInformation + vbOKOnly, gstrSysName
                                    strKey = .TextMatrix(.Row, mconIntCol�ۼ�)
                                    .TextMatrix(.Row, mconIntCol�ɱ���) = zlStr.FormatEx(strKey, mintPriceDigit, , True)
                                    .Text = strKey
                                    .TextMatrix(.Row, mconIntcol�ӳ���) = "0%"
                                    .TextMatrix(.Row, mconIntCol����) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ɱ���)) / Val(.TextMatrix(.Row, mconIntCol�ɹ���)) * 100, 7, , True)
                                End If
                            End If
                        End If
                    End If
                 End If
                
                '���ý��
                If strKey <> "" And .TextMatrix(.Row, mconIntCol����) <> "" Then
                    .TextMatrix(.Row, mconIntCol�ɱ����) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * Val(.TextMatrix(.Row, mconIntCol�ɱ���)), mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * Val(.TextMatrix(.Row, mconIntCol�ۼ�)), mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconintcol��Ʊ���) = IIf(Trim(.TextMatrix(.Row, mconintcol��Ʊ��)) = "", "", .TextMatrix(.Row, mconIntCol�ɱ����))
                    .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                End If
                
                Call ���ɱ���
                Call ��ʾ�ϼƽ��
                Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / Val(.TextMatrix(.Row, mconIntCol����ϵ��)))
            Case mconIntCol�ۼ�
                '������ۼ۲��ܴ���ָ�����ۼ�
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�ۼ۱���Ϊ�����ͣ������䣡", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .TxtVisible = False Then strKey = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�), mintPriceDigit, , True)
                
                '�ж���������ۼ���ָ�����ۼ�
                gstrSQL = "Select ָ�����ۼ� From ҩƷĿ¼ Where ҩƷID=[1] "
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", Val(.TextMatrix(.Row, 0)))
                
                dblָ�����ۼ� = Round(rsTemp!ָ�����ۼ� * Val(.TextMatrix(.Row, mconIntCol����ϵ��)), mintPriceDigit)
                strKey = Round(strKey, 5)
                If Val(strKey) > dblָ�����ۼ� Then
                    MsgBox "�ۼ۲��ܴ���ָ�����ۼۣ�ָ�����ۼۣ���" & dblָ�����ۼ� & "��", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                '���ۿ��ƣ�ʱ��ҩƷ���ۼ�ֱ�ӵ��ڳɱ��ۣ�ֻ��ʱ��ҩƷ�����޸��ۼ�
                If gtype_UserSysParms.P275_���۹���ģʽ = 2 And Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 And Val(strKey) <> Val(.TextMatrix(.Row, mconIntCol�ɱ���)) Then
                    If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                        MsgBox "��ʱ��ҩƷ���������۹���ģʽ���ۼ�Ӧ�ͳɱ���(" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ɱ���), mintPriceDigit, , True) & ")��ȣ�", vbInformation + vbOKOnly, gstrSysName
                        strKey = .TextMatrix(.Row, mconIntCol�ɱ���)
'                        Cancel = True
'                        .TxtSetFocus
'                        Exit Sub
                    End If
                End If
                
                .Text = zlStr.FormatEx(strKey, mintPriceDigit, , True)
                .TextMatrix(.Row, .Col) = .Text
                
'                If Len(Mid(.Text, InStr(1, .Text, ".") + 1)) > Get����(2, mintUnit) Then
'                    MsgBox "�ۼ۾��ȴ��������õļ��㾫�ȣ������䣡", vbInformation, gstrSysName
'                    Cancel = True
'                    .TxtSetFocus
'                    Exit Sub
'                End If
                                                
                dbl�ɱ��� = IIf(mintʱ������ۼۼӳɷ�ʽ = 0, Val(.TextMatrix(.Row, mconIntCol�ɱ���)), Val(.TextMatrix(.Row, mconIntCol�ɹ���)))
                .TextMatrix(.Row, mconIntcol�ӳ���) = zlStr.FormatEx(����ӳ���(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, .Col)), dbl�ɱ���), 2) & "%"
                '������
                .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ�)) * Val(.TextMatrix(.Row, mconIntCol����)), mintMoneyDigit, , True)
                .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) - Val(.TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                
                Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / Val(.TextMatrix(.Row, mconIntCol����ϵ��)))
            Case mconintCol���ۼ�
                '����ʱ�۷���ҩƷ�����ۼ����
                '��������ۼ۲��ܴ���ָ�����ۼ�
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "���ۼ۱���Ϊ�����ͣ������䣡", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .TxtVisible = False Then strKey = zlStr.FormatEx(.TextMatrix(.Row, mconintCol���ۼ�), gtype_UserDrugDigits.Digit_���ۼ�, , True)
                
                '�ж���������ۼ���ָ�����ۼ�
                gstrSQL = "Select ָ�����ۼ� From ҩƷĿ¼ Where ҩƷID=[1] "
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", Val(.TextMatrix(.Row, 0)))
                
                dblָ�����ۼ� = Round(rsTemp!ָ�����ۼ�, gtype_UserDrugDigits.Digit_���ۼ�)
                
                If Val(strKey) <> 0 Then
                    strKey = Round(strKey, gtype_UserDrugDigits.Digit_���ۼ�)
                End If
                If Val(strKey) > dblָ�����ۼ� Then
                    MsgBox "���ۼ۲��ܴ���ָ�����ۼۣ�ָ�����ۼۣ���" & dblָ�����ۼ� & "��", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                                
                .Text = zlStr.FormatEx(strKey, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                .TextMatrix(.Row, .Col) = .Text
                
                .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(Val(.TextMatrix(.Row, .Col)) * Val(.TextMatrix(.Row, mconIntCol����ϵ��)), mintPriceDigit, , True)
                .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ�)) * Val(.TextMatrix(.Row, mconIntCol����)), mintMoneyDigit, , True)
                
                If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                    .TextMatrix(.Row, mconIntCol�ɱ���) = .TextMatrix(.Row, mconIntCol�ۼ�)
                    .TextMatrix(.Row, mconIntCol�ɱ����) = .TextMatrix(.Row, mconIntCol�ۼ۽��)
                End If
                
                .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) - Val(.TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                
                dbl�ɱ��� = IIf(mintʱ������ۼۼӳɷ�ʽ = 0, Val(.TextMatrix(.Row, mconIntCol�ɱ���)), Val(.TextMatrix(.Row, mconIntCol�ɹ���)))
                .TextMatrix(.Row, mconIntcol�ӳ���) = zlStr.FormatEx(����ӳ���(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ۼ�)), dbl�ɱ���), 2) & "%"
                
                Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.Text))
                Call ��ʾ�ϼƽ��
            Case mconIntCol�ɱ����
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�Բ��𣬳ɱ�������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) >= 10 ^ 14 - 1 Then
                        MsgBox "�ɱ�������С��" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) * Val(.TextMatrix(.Row, mconIntCol����)) < 0 Then
                        MsgBox "�ɱ�������Ӧ����������һ�£�", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                
                '��ʽ�����
                If strKey <> "" Then
                    strKey = zlStr.FormatEx(strKey, mintMoneyDigit, , True)
                    .Text = strKey
                End If
                
                If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol�ɱ����) Then
                    If .TextMatrix(.Row, mconIntCol����) <> "" Then
                        '���ۿ��ƣ�����ҩƷ�����ܵ����ɱ�����Ϊ�ۼ۹̶����ۼ۽��Ҳ�̶���
                        If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                            If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 0 And strKey <> .TextMatrix(.Row, mconIntCol�ۼ۽��) Then
                                MsgBox "�ö���ҩƷ���������۹���ģʽ�����ܵ����ɱ���", vbInformation + vbOKOnly, gstrSysName
                                strKey = .TextMatrix(.Row, mconIntCol�ۼ۽��)
                                .Text = strKey
                                Cancel = True
'                                .TxtSetFocus
                                Exit Sub
                            End If
                        Else
                            If mbln�Ӽ��� Then
                                'ȡ�øı�ɹ����ǰ�ļӼ���
                                dbl�ɱ��� = IIf(mintʱ������ۼۼӳɷ�ʽ = 0, Val(.TextMatrix(.Row, mconIntCol�ɱ���)), Val(.TextMatrix(.Row, mconIntCol�ɹ���)))
                                mdbl�Ӽ��� = 15
                                If Val(.TextMatrix(.Row, mconIntCol�ۼ�)) <> 0 And dbl�ɱ��� <> 0 Then
                                    mdbl�Ӽ��� = ����ӳ���(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ۼ�)), dbl�ɱ���)
                                End If
                            End If
                            
                            '����ɱ��ۡ��ɹ��ۣ��ɱ���=�ɹ����/����;�ɹ���=(�ɹ����/����)/����
                            If Val(.TextMatrix(.Row, mconIntCol����)) = 0 Then
                                .TextMatrix(.Row, mconIntCol����) = "100"
                            End If
                            .TextMatrix(.Row, mconIntCol�ɱ���) = zlStr.FormatEx(strKey / .TextMatrix(.Row, mconIntCol����), mintCostDigit, , True)
                            .TextMatrix(.Row, mconIntCol�ɹ���) = zlStr.FormatEx((strKey / .TextMatrix(.Row, mconIntCol����)) * 100 / Val(.TextMatrix(.Row, mconIntCol����)), mintCostDigit, , True)
                            
                            '���ݲ�������ʱ��ҩƷ�ۼ۹�ʽ�гɱ��۵��㷨
                            dbl�ɱ��� = IIf(mintʱ������ۼۼӳɷ�ʽ = 0, Val(.TextMatrix(.Row, mconIntCol�ɱ���)), Val(.TextMatrix(.Row, mconIntCol�ɹ���)))
                            
                            '��ʱ��ҩƷ�Ĵ���
                            If .TextMatrix(.Row, mconIntColԭ����) <> "" Then
                                '���¼������ۼۡ����
                                If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 Then
                                    '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                                    If mbln�Ӽ��� Then
                                        dbl�ӳ��� = (mdbl�Ӽ��� / 100)
                                        dblTemp�ۼ� = dbl�ɱ��� * (1 + (mdbl�Ӽ��� / 100))
                                        
                                        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  'û�й�ѡʱ��ȡ�ϴ��ۼ۲���
                                            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), dbl�ɱ���, dbl�ӳ���, dblTemp�ۼ�), mintPriceDigit, , True)
                                        End If
                                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ�)) * Val(.TextMatrix(.Row, mconIntCol����)), mintMoneyDigit, , True)
                                        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                                    Else
                                        If mintʱ�۷ֶμӳɷ�ʽ = 1 Then
                                            If get�ֶμӳ��ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����ϵ��)), dbl�ɱ���, dbl�ӳ���, dblTemp�ۼ�) = False Then
                                                Cancel = True
                                                .TxtSetFocus
                                                Exit Sub
                                            End If
                                        Else
                                            dbl�ӳ��� = Val(Replace(.TextMatrix(.Row, mconIntcol�ӳ���), "%", "")) / 100
                                            dblTemp�ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                                        End If
                                                                            
                                        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  'û�й�ѡʱ��ȡ�ϴ��ۼ۲���
                                            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), dbl�ɱ���, dbl�ӳ���, dblTemp�ۼ�), mintPriceDigit, , True)
                                        End If
                                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                                        .TextMatrix(.Row, mconIntcol�ӳ���) = zlStr.FormatEx(dbl�ӳ��� * 100, 2) & "%"
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(.Row, mconIntCol����)) <> 0 Then
                        .TextMatrix(.Row, mconIntCol�ɱ���) = zlStr.FormatEx(strKey / Val(.TextMatrix(.Row, mconIntCol����)), mintCostDigit, , True)
                        .TextMatrix(.Row, mconIntCol�ɹ���) = zlStr.FormatEx(strKey / Val(.TextMatrix(.Row, mconIntCol����)) * 100 / Val(.TextMatrix(.Row, mconIntCol����)), mintCostDigit, , True)
                    End If
                    .TextMatrix(.Row, mconintcol��Ʊ���) = IIf(Trim(.TextMatrix(.Row, mconintcol��Ʊ��)) = "", "", zlStr.FormatEx(strKey, mintMoneyDigit, , True))
                    
                    '���ۿ��ƣ�����ҩƷ�����ܵ����ɱ�����Ϊ�ۼ۹̶����ۼ۽��Ҳ�̶���
                    If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                        .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(strKey / Val(.TextMatrix(.Row, mconIntCol����)), mintCostDigit, , True)
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                    End If
                    
                    .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - strKey, mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconIntCol�ɱ����) = zlStr.FormatEx(strKey, mintMoneyDigit, , True)
                End If
                
                Call ���ɱ���
                Call ��ʾ�ϼƽ��
                Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / IIf(Val(.TextMatrix(.Row, mconIntCol����ϵ��)) = 0, 1, Val(.TextMatrix(.Row, mconIntCol����ϵ��))))
            Case mconIntCol����
                If .TextMatrix(.Row, 0) = "" Then
                    .Text = ""
                    Exit Sub
                End If
                
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "�Բ��������������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�Բ�����������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) = 0 Then
                        MsgBox "�Բ�����������Ϊ��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If mint�༭״̬ = 2 And Val(.TextMatrix(.Row, mconIntCol����)) <> 0 And .TextMatrix(.Row, mconIntCol�Ƿ�����) = "��" Then
                        If Not ��ͬ����(Val(strKey), Val(.TextMatrix(.Row, mconIntCol����))) Then
                            MsgBox "�Բ��������ķ���Ӧ����ԭ���������ķ���һ�£�", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    If Val(strKey) < 0 Then
                        If mint�༭״̬ = 8 Or mbln�˻� Then
                            MsgBox "�˿ⵥ�������븺���������䣡", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        If Not zlStr.IsHavePrivs(mstrPrivs, "��������") Then
                            MsgBox "�Բ�����û�и���������Ȩ�ޣ������䣡", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        If .TextMatrix(.Row, mconIntCol��������) = 1 Then
                            MsgBox "����ҩƷ����������⣬������", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "��������С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                    '1 ����Ƿ����㹻�Ŀ������˻�;2 ��鸺���˿�ʱ����Ƿ��㹻
                    If mint�༭״̬ = 8 Or mbln�˻� Or Val(strKey) < 0 Then
                        If Not CheckUsableNum(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)), Val(.Text), Val(.TextMatrix(.Row, mconIntCol����ϵ��)), Trim(txtNO.Text), 1, mint�����, mintNumberDigit) Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    .Text = strKey
                    
'                    If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) > Get����(3, mintUnit) Then
'                        MsgBox "�������ȴ��������õļ��㾫�ȣ������䣡", vbInformation, gstrSysName
'                        Cancel = True
'                        .TxtSetFocus
'                        Exit Sub
'                    End If
                    If .TextMatrix(.Row, mconIntCol�ɱ���) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ɱ����) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ɱ���) * strKey, mintMoneyDigit, , True)
                        
                        '���ۿ���
                        If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                            '������������۹������������ۼ�
                        Else
                            '���ݲ�������ʱ��ҩƷ�ۼ۹�ʽ�гɱ��۵��㷨
                            dbl�ɱ��� = IIf(mintʱ������ۼۼӳɷ�ʽ = 0, Val(.TextMatrix(.Row, mconIntCol�ɱ���)), Val(.TextMatrix(.Row, mconIntCol�ɹ���)))
                            
                            'ʱ��ҩƷ�Ĵ���
                            If .TextMatrix(.Row, mconIntColԭ����) <> "" And mint�༭״̬ <> 8 And mbln�˻� <> True Then
                                If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 Then
                                    '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                                    If mbln�Ӽ��� Then
                                        mdbl�Ӽ��� = 15
                                        If Val(.TextMatrix(.Row, mconIntCol�ۼ�)) <> 0 And dbl�ɱ��� <> 0 Then
                                            mdbl�Ӽ��� = ����ӳ���(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ۼ�)), dbl�ɱ���)
                                        End If
                                        
    '                                    mdbl�Ӽ��� = mdbl�Ӽ��� / 100
                                        dblTemp�ۼ� = dbl�ɱ��� * (1 + (mdbl�Ӽ��� / 100))
                                        
                                        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  'û�й�ѡʱ��ȡ�ϴ��ۼ۲���
                                            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), dbl�ɱ���, mdbl�Ӽ���, dblTemp�ۼ�), mintPriceDigit, , True)
                                        End If
                                        
                                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ�)) * strKey, mintMoneyDigit, , True)
                                        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                                    Else
                                        If mintʱ�۷ֶμӳɷ�ʽ = 1 Then
                                            If get�ֶμӳ��ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����ϵ��)), dbl�ɱ���, dbl�ӳ���, dblTemp�ۼ�) = False Then
                                                Cancel = True
                                                .TxtSetFocus
                                                Exit Sub
                                            End If
                                        Else
                                            dbl�ӳ��� = Val(Replace(.TextMatrix(.Row, mconIntcol�ӳ���), "%", "")) / 100
                                            dblTemp�ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                                        End If
                                                                            
                                        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  'û�й�ѡʱ��ȡ�ϴ��ۼ۲���
                                            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), dbl�ɱ���, dbl�ӳ���, dblTemp�ۼ�), mintPriceDigit, , True)
                                        End If
                                        .TextMatrix(.Row, mconIntcol�ӳ���) = zlStr.FormatEx(dbl�ӳ��� * 100, 2) & "%"
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�) * strKey, mintMoneyDigit, , True)
                    End If
                    .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconintcol��Ʊ���) = .TextMatrix(.Row, mconIntCol�ɱ����)
                                    
                    .TextMatrix(.Row, mconIntCol����) = strKey
                    Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / Val(.TextMatrix(.Row, mconIntCol����ϵ��)))
                End If
                ��ʾ�ϼƽ��
            Case mconIntCol��������
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "�Բ��������������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�Բ�����������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        If Not zlStr.IsHavePrivs(mstrPrivs, "��������") Then
                            MsgBox "�Բ�����û�и���������Ȩ�ޣ������䣡", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If Val(strKey) <> 0 And Not ��ͬ����(Val(strKey), Val(.TextMatrix(.Row, mconIntCol����))) Then
                        MsgBox "�Բ��𣬳��������ķ���Ӧ����ԭ������һ�£�", vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 0 Then
                        If Val(strKey) > Val(.TextMatrix(.Row, mconIntCol����)) Then
                            MsgBox "�Բ��𣬳����������ܴ���ԭ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    Else
                        If Val(strKey) < Val(.TextMatrix(.Row, mconIntCol����)) Then
                            MsgBox "�Բ��𣬳����������ܴ���ԭ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                                        
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "������������С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                    .Text = strKey
                    
                    If .TextMatrix(.Row, mconIntCol�ɱ���) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ɱ����) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ɱ���) * strKey, mintMoneyDigit, , True)
                    End If
                    If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�) * strKey, mintMoneyDigit, , True)
                    End If
                    .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                    
                    gstrSQL = "select sum(nvl(��Ʊ���,0)) as ��Ʊ��� " _
                        & " From ҩƷ�շ���¼ x,(Select �շ�id,��Ʊ��� From Ӧ����¼ Where ϵͳ��ʶ=1 And ��¼����=0)  y " _
                        & " WHERE x.id=y.�շ�id(+) and x.NO=[1] AND ����=1 " _
                        & " and x.ҩƷid=[2] " _
                        & " and x.���=[3] "
                    Set rsDrug = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�, Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol���)))
                    
                    If rsDrug.EOF Then
                        .TextMatrix(.Row, mconintcol��Ʊ���) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                    Else
                        .TextMatrix(.Row, mconintcol��Ʊ���) = zlStr.FormatEx(strKey / .TextMatrix(.Row, mconIntCol����) * rsDrug.Fields(0), mintMoneyDigit, , True)
                    End If
                    
                    .TextMatrix(.Row, mconIntCol��������) = strKey
                    Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconintCol���ۼ�)))
                End If
                ��ʾ�ϼƽ��
            Case mconintcol��Ʊ��
                If Trim(.TextMatrix(.Row, mconIntColҩ��)) = "" Then
                    .Text = ""
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .ColData(mconIntCol��Ʊ����) = 5
                        .ColData(mconintcol��Ʊ���) = 5
                        .ColData(mconintcol��Ʊ����) = 5
                        .TextMatrix(.Row, mconintcol��Ʊ����) = ""
                        .TextMatrix(.Row, mconintcol��Ʊ���) = ""
                        .TextMatrix(.Row, mconIntCol��Ʊ����) = ""
                        .TextMatrix(.Row, .Col) = " "
                        .Text = " "
                    ElseIf .TxtVisible = False Then
                        If Trim(.TextMatrix(.Row, mconintcol��Ʊ��)) = "" Then
                           .ColData(mconIntCol��Ʊ����) = 5
                           .ColData(mconintcol��Ʊ���) = 5
                           .ColData(mconintcol��Ʊ����) = 5
                           .TextMatrix(.Row, mconintcol��Ʊ����) = ""
                           .TextMatrix(.Row, mconintcol��Ʊ���) = ""
                           .TextMatrix(.Row, mconIntCol��Ʊ����) = ""
                           .TextMatrix(.Row, .Col) = " "
                           .Text = " "
                        Else
                            .Text = .TextMatrix(.Row, .Col)
                           
                            If mint�༭״̬ = 9 Or mint�༭״̬ = 3 Or mint�༭״̬ = 7 Then
                                .ColData(mconIntCol��Ʊ����) = IIf(InStr(1, mstrControlItem, ",��Ʊ����,") > 0, 2, 5)
                                .ColData(mconintcol��Ʊ����) = IIf(InStr(1, mstrControlItem, ",��Ʊ����,") > 0, 4, 5)
                                .ColData(mconintcol��Ʊ���) = IIf(InStr(1, mstrControlItem, ",��Ʊ���,") > 0, 4, 5)
                            Else
                                .ColData(mconIntCol��Ʊ����) = 2
                                .ColData(mconintcol��Ʊ����) = 4
                                .ColData(mconintcol��Ʊ���) = 4
                           End If
                        End If
                    End If
                ElseIf mint��¼״̬ = 1 Then
                    If mint�༭״̬ = 9 Or mint�༭״̬ = 3 Or mint�༭״̬ = 7 Then
                         .ColData(mconIntCol��Ʊ����) = IIf(InStr(1, mstrControlItem, ",��Ʊ����,") > 0, 2, 5)
                         .ColData(mconintcol��Ʊ����) = IIf(InStr(1, mstrControlItem, ",��Ʊ����,") > 0, 4, 5)
                         .ColData(mconintcol��Ʊ���) = IIf(InStr(1, mstrControlItem, ",��Ʊ���,") > 0, 4, 5)
                     Else
                         .ColData(mconIntCol��Ʊ����) = 2
                         .ColData(mconintcol��Ʊ����) = 4
                         .ColData(mconintcol��Ʊ���) = 4
                    End If
                    .TextMatrix(.Row, mconintcol��Ʊ���) = .TextMatrix(.Row, mconIntCol�ɱ����)
                End If
                    
                Exit Sub
            Case mconintcol��Ʊ����
                If Trim(.Text) = "" Then
                   If mconintcol��Ʊ���� <> mintLastCol Then
                       .Col = GetNextEnableCol(mconintcol��Ʊ����)
                       .Text = ""
                       Cancel = True
                       Exit Sub
                   End If
                End If
            Case mconintcol��Ʊ���
                If Trim(.TextMatrix(.Row, mconIntColҩ��)) = "" Then
                    .Text = ""
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�Բ��𣬷�Ʊ������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Abs(Val(strKey)) < 0.001 Then
                        MsgBox "�Բ��𣬷�Ʊ���������0.001,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 14 - 1 Then
                        MsgBox "��Ʊ������С��" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                
                If strKey <> "" Then
                    strKey = zlStr.FormatEx(strKey, 2, , True)
                    .Text = strKey
                ElseIf .TxtVisible = True Then
                    .Text = " "
                ElseIf .TxtVisible = False Then
                    If .TextMatrix(.Row, .Col) = "" Then
                        .Text = " "
                    Else
                        .Text = .TextMatrix(.Row, .Col)
                    End If
                    
                End If
            Case mconIntCol��Ʊ����
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        
                        If strKey = "" Then
                            MsgBox "�Բ���Ч�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "�Բ��𣬷�Ʊ���ڱ���Ϊ��������(2000-10-10) �� ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            Case mconIntCol��׼�ĺ�
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol��׼�ĺ�) = ""
                    End If
                    If mconIntCol��׼�ĺ� <> mintLastCol Then
                        .Col = GetNextEnableCol(mconIntCol��׼�ĺ�)
                        Cancel = True
                    End If
                    Exit Sub
                End If
            Case mconIntCol���
                '�޴���
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol���) = ""
                    End If
                    If mconIntCol��� <> mintLastCol Then
                        .Col = GetNextEnableCol(mconIntCol���)
                        Cancel = True
                        Exit Sub
                    End If
                Else
                    Dim rs��� As New Recordset
                    
                    gstrSQL = "Select ����,����,���� From ҩƷ��� " _
                            & "Where upper(����) like [1] or Upper(����) like [1] or Upper(����) like [2] "
                    Set rs��� = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%", strKey & "%")
                    
                    If rs���.EOF Then
                        .TextMatrix(.Row, mconIntCol���) = .Text
                        If mconIntCol��� <> mintLastCol And mconIntCol��� < mconintcol��Ʊ�� Then
                            .Col = mconintcol��Ʊ��
                            Cancel = True
                            Exit Sub
                        End If
                    Else
                        If rs���.RecordCount = 1 Then
                            .TextMatrix(.Row, mconIntCol���) = rs���.Fields("����")
                            .Text = rs���.Fields("����")
                        Else
                            Set msh����.Recordset = rs���
                            With msh����
                                .Redraw = False
                                .Left = Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                                .Top = Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight
                                .Visible = True
                                .SetFocus
                                .ColWidth(0) = 800
                                .ColWidth(1) = 800
                                .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                                .Row = 1
                                .Col = 0
                                .TopRow = 1
                                .ColSel = .Cols - 1
                                .Redraw = True
                                Cancel = True
                                Exit Sub
                            End With
                        End If
                    End If
                End If
            Case mconintcol�������
                '�޴���
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .ColData(mconintcol�������) = 5
                        .TextMatrix(.Row, mconintcol�������) = ""
                        .TextMatrix(.Row, .Col) = ""
                        .Text = ""
                    ElseIf .TxtVisible = False Then
                        If Trim(.TextMatrix(.Row, mconintcol�������)) = "" Then
                           .ColData(mconintcol�������) = 5
                           .TextMatrix(.Row, mconintcol�������) = ""
                           .TextMatrix(.Row, .Col) = ""
                           .Text = ""
                        Else
                            .Text = .TextMatrix(.Row, .Col)
                            .ColData(mconintcol�������) = 2
                        End If
                    End If
                    
                Else
                    .TextMatrix(.Row, .Col) = .Text
                    .ColData(mconintcol�������) = 2
                End If
                
                If mconintcol������� <> mintLastCol Then
                    .Col = GetNextEnableCol(mconintcol�������)
                    Cancel = True
                    Exit Sub
                End If
            Case mconintcol�������
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconintcol�������) = ""
                    End If
                    If mconintcol������� <> mintLastCol Then
                        .Col = GetNextEnableCol(mconintcol�������)
                        Cancel = True
                        Exit Sub
                    End If
                Else
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        
                        If strKey = "" Then
                            MsgBox "�Բ���������ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "�Բ���������ڱ���Ϊ��������(2000-10-10) �� ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            Case mconintcol��Ʒ�ϸ�֤
                '�޴���
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconintcol��Ʒ�ϸ�֤) = ""
                    End If
                    If mconintcol��Ʒ�ϸ�֤ <> mintLastCol Then
                        .Col = GetNextEnableCol(mconintcol��Ʒ�ϸ�֤)
                        Cancel = True
                        Exit Sub
                    End If
                End If
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ҩƷĿ¼��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal lngҩƷID As Long, ByVal strҩƷ���� As String, ByVal strͨ���� As String, _
    ByVal str��Ʒ�� As String, ByVal strҩƷ��Դ As String, ByVal str����ҩ�� As String, _
    ByVal str��� As String, ByVal str���� As String, ByVal str��λ As String, ByVal num�ۼ� As Double, _
    ByVal numָ�������� As Double, ByVal strԭ������ As String, ByVal intԭЧ�� As Integer, _
    ByVal str���� As String, ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
    ByVal int�Ƿ��� As Integer, ByVal intҩ������ As Integer, ByVal dbl�ӳ��� As Double, ByVal str�������� As String, _
    ByVal str�ۼ۵�λ As String, ByVal strԭ���� As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim rsPrice As New Recordset
    Dim lngDepartid As Long
    Dim dblRate As Double, dbl�ɱ��� As Double
    Dim bln�б�ҩƷ As Boolean, dbl��������� As Double
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long, j As Long
    Dim dblʱ�۳ɱ��� As Double
    Dim strҩ�� As String
    Dim rsRecord As ADODB.Recordset
    Dim rsProvider As ADODB.Recordset
    Dim dblTemp�ۼ� As Double
    Dim rs�ۼ� As ADODB.Recordset
    
    SetColValue = False
    On Error GoTo errHandle
'    If mint�༭״̬ = 8 Then
'        '����Ƿ��ظ�
'        If Not CheckRepeatMedicine(mshBill, lngҩƷID & "," & "0" & "|" & lng���� & "," & mconIntCol����, introw) Then
'            Exit Function
'        End If
'    End If
    
    With mshBill
        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol�к� Then .TextMatrix(intRow, intCol) = ""
        Next
        
        gstrSQL = "SELECT Nvl(a.���������,0) ���������,nvl(a.����,0) ����,Nvl(a.�б�ҩƷ,0) �б�ҩƷ,nvl(a.�ɱ���,0) �ɱ���, a.��׼�ĺ�,a.�ϴ���׼�ĺ�,a.�ϴβ���,b.����,a.ԭ����,a.�ϴ���������,a.ҩ�ۼ��� " & _
                  "from ҩƷ��� a,�շ���ĿĿ¼ b  where a.ҩƷid=b.id and ҩƷid=[1]"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ����]", lngҩƷID)
            
        If rsTemp!���� = 0 Then
            dblRate = 100
        Else
            dblRate = rsTemp!����
        End If
        bln�б�ҩƷ = (rsTemp!�б�ҩƷ = 1)
        dbl��������� = rsTemp!���������
        dbl�ɱ��� = rsTemp!�ɱ���
        
        .TextMatrix(intRow, 0) = lngҩƷID
        
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = strͨ����
        Else
            strҩ�� = IIf(str��Ʒ�� <> "", str��Ʒ��, strͨ����)
        End If
        
        .TextMatrix(intRow, mconIntColҩƷ���������) = strҩƷ���� & strҩ��
        .TextMatrix(intRow, mconIntColҩƷ����) = strҩƷ����
        .TextMatrix(intRow, mconIntColҩƷ����) = strҩ��
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        Else
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ���������)
        End If
        
        .TextMatrix(intRow, mconIntCol��Ʒ��) = str��Ʒ��
        
        .TextMatrix(intRow, mconIntCol��Դ) = strҩƷ��Դ
        .TextMatrix(intRow, mconIntCol����ҩ��) = str����ҩ��
        .TextMatrix(intRow, mconIntColҩ�ۼ���) = IIf(IsNull(rsTemp!ҩ�ۼ���), "", rsTemp!ҩ�ۼ���)
        .TextMatrix(intRow, mconIntCol���) = str���
        .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(str����), "", str����)
        .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(strԭ����), nvl(rsTemp!ԭ����), strԭ����)
        
        '���ء���׼�ĺš��������ڹ��򣬸��ݲ�������ȡ
        '���������ȴ��ϴ����ȡ
        '���أ�ֱ�Ӵӹ�����ȡ�ϴβ��أ����û������շ���Ŀ��ȡ���أ�û���������
        '��׼�ĺţ����ȴӹ�����ȡ�ϴ���׼�ĺţ����û����ӹ�����ȡ��׼�ĺţ���û��������׼�ĺ�
        '�������ڣ����ȴӹ�����ȡ�ϴ��������ڣ����û������
        '�ɱ��ۣ��ӹ�����ȡ�ɱ���
        
        '���������ȴ�����������ȡ
        '���أ����ȴӿ������������ȡ���أ����û������շ���Ŀ��ȡ���أ�û���������
        '��׼�ĺţ����ȴӿ������������ȡ��׼�ĺţ����û����ӹ�����ȡ��׼�ĺţ���û��������׼�ĺ�
        '�������ڣ����ȴӿ������������ȡ�������ڣ����û������
        '�ɱ��ۣ����ȴ�ҩƷ�������������ȡ�ϴβɹ��ۣ�û����ӹ�����ȡ�ɱ���
        If IIf(IsNull(rsTemp!�ϴβ���), "", rsTemp!�ϴβ���) <> "" Then
            .TextMatrix(intRow, mconIntCol����) = rsTemp!�ϴβ���
        Else
            .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
        End If
        If IIf(IsNull(rsTemp!�ϴ���׼�ĺ�), "", rsTemp!�ϴ���׼�ĺ�) <> "" Then
            .TextMatrix(intRow, mconIntCol��׼�ĺ�) = rsTemp!�ϴ���׼�ĺ�
        Else
            .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
        End If
        
        If IIf(IsNull(rsTemp!�ϴ���������), "", rsTemp!�ϴ���������) <> "" Then
            .TextMatrix(intRow, mconIntCol��������) = Format(rsTemp!�ϴ���������, "yyyy-mm-dd")
        Else
            .TextMatrix(intRow, mconIntCol��������) = ""
        End If
        
        .TextMatrix(intRow, mconIntCol��λ) = str��λ
        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(num�ۼ�, mintPriceDigit, , True)
        .TextMatrix(intRow, mconIntColָ��������) = zlStr.FormatEx(numָ��������, mintCostDigit, , True)
        .TextMatrix(intRow, mconIntColԭ������) = IIf(IsNull(strԭ������), "", strԭ������)
        .TextMatrix(intRow, mconIntCol����) = lng����
        
        'ȡ����ҩƷ�����ż�Ч�ڣ��Լ�ԭ�ɹ���
        If mint�༭״̬ = 8 Or mbln�˻� Then
            gstrSQL = " Select �ϴ����� ����,Ч��,�ϴ���������,�ϴβ���,ԭ����,��׼�ĺ�,�ϴβɹ��� From ҩƷ���" & _
                    " Where �ⷿID=[1] And ҩƷID=[2] " & _
                    " And ����=1 And nvl(����,0)=[3] "
            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ��ҩƷ�����ź�Ч��]", cboStock.ItemData(cboStock.ListIndex), lngҩƷID, lng����)
            If rsPrice.RecordCount <> 0 Then
                .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsPrice!����), "", rsPrice!����)
                .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsPrice!Ч��), "", rsPrice!Ч��)
                .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsPrice!�ϴβ���), "", rsPrice!�ϴβ���)
                .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(rsPrice!ԭ����), "", rsPrice!ԭ����)
                .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsPrice!��׼�ĺ�), "", rsPrice!��׼�ĺ�)
                .TextMatrix(intRow, mconIntCol��������) = Format(rsPrice!�ϴ���������, "yyyy-mm-dd")
                
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
                    '����Ϊ��Ч��
                    .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
                End If
                
                dbl�ɱ��� = nvl(rsPrice!�ϴβɹ���, 0)
                If dbl�ɱ��� > 0 Then
                    .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(dbl�ɱ��� * num����ϵ��, mintCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx(dbl�ɱ��� * num����ϵ�� * dblRate / 100, mintCostDigit, , True)
                End If
            End If
        End If
        
        'ԭЧ���ֶ����汣��ԭЧ�ڣ�ָ����ۣ��Ƿ��ۣ�ҩ�������ȣ���ʽΪ��ԭЧ��||ָ�������||�Ƿ���||ҩ������
        .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(intԭЧ��), "0", intԭЧ��) & "||" & dbl�ӳ��� & "||" & int�Ƿ��� & "||" & intҩ������
       
        .TextMatrix(intRow, mconintcol����) = str����
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        If intRow > 1 Then
            .TextMatrix(intRow, mconintcol�������) = .TextMatrix(intRow - 1, mconintcol�������)
            .TextMatrix(intRow, mconintcol�������) = .TextMatrix(intRow - 1, mconintcol�������)
            .TextMatrix(intRow, mconintcol��Ʊ��) = .TextMatrix(intRow - 1, mconintcol��Ʊ��)
            .TextMatrix(intRow, mconintcol��Ʊ����) = .TextMatrix(intRow - 1, mconintcol��Ʊ����)
            .TextMatrix(intRow, mconIntCol��Ʊ����) = .TextMatrix(intRow - 1, mconIntCol��Ʊ����)
        End If
        
        SetInputFormat intRow
        SetDisCount intRow, dblRate
        lngDepartid = cboStock.ItemData(cboStock.ListIndex)
        
        '��������
        Call GetҩƷ��������(intRow)
        
        '˵�����������ַ�������Ͳ����������Ŀ������������ٶȡ�
        '�������Բ�����Щ��ֱ���õ�һ��SQL���ʵ�֣�����������ҩƷ�Ͷ������ݿ���ɨ��һ�Ρ�
        
        '�Զ��۲ɹ�������ȡ�ϴεĲɹ��ۺͿ���
        If Not (mint�༭״̬ = 8 Or mbln�˻�) Then
            If mintȡ�ϴβɹ��۷�ʽ = 0 Then
                If Val(.TextMatrix(intRow, mconIntCol��������)) = 1 Then
                    gstrSQL = "select �ϴβɹ���,�ϴβ���,��׼�ĺ�,�ϴ��������� from ҩƷ��� where ����=1 and �ⷿid=[1] and ҩƷid=[2] " & _
                            " and nvl(����,0) =(select max(nvl(����,0)) from ҩƷ��� where ����=1 and �ⷿid=[1] )"
                Else
                    gstrSQL = "select �ϴβɹ���,�ϴβ���,��׼�ĺ�,�ϴ��������� from ҩƷ��� where ����=1 and �ⷿid=[1] and ҩƷid=[2]"
                End If
                Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ�ϴβɹ���]", lngDepartid, lngҩƷID)
                
                If Not rsPrice.EOF Then
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsPrice!�ϴβ���), IIf(IsNull(rsTemp!����), "", rsTemp!����), rsPrice!�ϴβ���)
                    'mintʱ������ۼۼӳɷ�ʽ
                    If nvl(rsPrice.Fields(0), 0) = 0 Then
                        If dbl�ɱ��� > 0 Then
                            .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(dbl�ɱ��� * num����ϵ��, mintCostDigit, , True)
                            .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx(dbl�ɱ��� * num����ϵ�� * dblRate / 100, mintCostDigit, , True)
                        End If
                    Else
                        .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(rsPrice.Fields(0) * num����ϵ��, mintCostDigit, , True)
                        .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx(rsPrice.Fields(0) * num����ϵ�� * dblRate / 100, mintCostDigit, , True)
                    End If
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsPrice!��׼�ĺ�), IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�), rsPrice!��׼�ĺ�)
                    .TextMatrix(intRow, mconIntCol��������) = IIf(IsNull(rsPrice!�ϴ���������), "", Format(rsPrice!�ϴ���������, "yyyy-mm-dd"))
                Else
                    .TextMatrix(intRow, mconIntCol��������) = ""
                    If dbl�ɱ��� > 0 Then
                        .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(dbl�ɱ��� * num����ϵ��, mintCostDigit, , True)
                        .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx(dbl�ɱ��� * num����ϵ�� * dblRate / 100, mintCostDigit, , True)
                    End If
                End If
                If Val(.TextMatrix(intRow, mconIntCol�ɹ���)) <> 0 Then
                    .TextMatrix(intRow, mconIntCol����) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ɱ���)) / Val(.TextMatrix(intRow, mconIntCol�ɹ���)) * 100, 7, , True)
                End If
            Else
                If dbl�ɱ��� > 0 Then
                    .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(dbl�ɱ��� * num����ϵ��, mintCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx(dbl�ɱ��� * num����ϵ�� * dblRate / 100, mintCostDigit, , True)
                End If
            End If
        End If
        
        If mblnȡĿ¼�в�����Ϣ = True Then 'ȡҩƷĿ¼�еĲ�����׼�ĺ�
            If IIf(IsNull(rsTemp!����), "", rsTemp!����) <> "" Then
                .TextMatrix(intRow, mconIntCol����) = rsTemp!����
            End If
            If IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�) <> "" Then
                .TextMatrix(intRow, mconIntCol��׼�ĺ�) = rsTemp!��׼�ĺ�
            End If
        End If
        
        '���ݲ�������ʱ��ҩƷ�ۼ۹�ʽ�гɱ��۵��㷨
        dblʱ�۳ɱ��� = IIf(mintʱ������ۼۼӳɷ�ʽ = 0, Val(.TextMatrix(intRow, mconIntCol�ɱ���)), Val(.TextMatrix(intRow, mconIntCol�ɹ���)))
        
        'ʱ��ҩƷ����
        If int�Ƿ��� = 1 Then
            '���ۿ��ƣ�ʱ��ҩƷ���ۼ�ֱ�ӵ��ڳɱ���
            If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(intRow, 0))) = True Then
                .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ɱ���), mintPriceDigit, , True)
                If .TextMatrix(intRow, mconIntCol����) <> "" Then
                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                End If
            Else
                If mint�༭״̬ <> 8 And mbln�˻� = False Then
                    dblTemp�ۼ� = dblʱ�۳ɱ��� * (1 + dbl�ӳ���)
                    '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                    If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� = 1 Then
                        gstrSQL = "select nvl(�ϴ��ۼ�,0) �ϴ��ۼ� from ҩƷ��� where ҩƷid=[1]"
                                         
                        Set rs�ۼ� = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯ�ۼ�", lngҩƷID)
                        If rs�ۼ�!�ϴ��ۼ� > 0 Then
                            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rs�ۼ�!�ϴ��ۼ� * mlng��װϵ��, mintPriceDigit, , True)
                        Else
                            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(intRow, 0)), dblʱ�۳ɱ���, dbl�ӳ���, dblTemp�ۼ�), mintPriceDigit, , True)
                        End If
                    Else
                        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(intRow, 0)), dblʱ�۳ɱ���, dbl�ӳ���, dblTemp�ۼ�), mintPriceDigit, , True)
                    End If
                    .TextMatrix(intRow, mconIntcol�ӳ���) = zlStr.FormatEx(dbl�ӳ��� * 100, 2) & "%"
                Else
                    '���������ʽ�����ۼ�
                    gstrSQL = " Select Decode(Nvl(����,0),0,nvl(ʵ�ʽ��,0)/Nvl(ʵ������,0),Nvl(���ۼ�,nvl(ʵ�ʽ��,0)/Nvl(ʵ������,0))) �ۼ� From ҩƷ���" & _
                              " Where �ⷿID=[1] And ҩƷID=[2] And ����=1 And NVL(����,0)=[3]"
                    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[���������ʽ�����ۼ�]", cboStock.ItemData(cboStock.ListIndex), lngҩƷID, lng����)
                    
                    If Not rsTemp.EOF Then
                        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsTemp!�ۼ� * num����ϵ��, mintPriceDigit, , True)
                    Else
                        gstrSQL = "select nvl(�ϴ��ۼ�,0) �ۼ� from ҩƷ��� where ҩƷid=[1]"
                        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯ�ۼ�", lngҩƷID)
                        
                        If Not rsTemp.EOF Then
                            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsTemp!�ۼ� * num����ϵ��, mintPriceDigit, , True)
                        End If
                    End If
                    
                    If Val(.TextMatrix(intRow, mconIntCol�ۼ�)) <> 0 And dblʱ�۳ɱ��� <> 0 Then
                        .TextMatrix(intRow, mconIntcol�ӳ���) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol�ۼ�)) / dblʱ�۳ɱ��� - 1) * 100, 2) & "%"
                    End If
                End If
            End If
        Else
            '����ҩƷ��ʾ�ӳ��ʣ���ʵ�����壬����ʾ
            If Val(.TextMatrix(intRow, mconIntCol�ۼ�)) <> 0 And dblʱ�۳ɱ��� <> 0 Then
                .TextMatrix(intRow, mconIntcol�ӳ���) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol�ۼ�)) / dblʱ�۳ɱ��� - 1) * 100, 2) & "%"
            End If
                        
            '���ۿ��ƣ�����ҩƷ���ɱ���Ĭ�ϵ����ۼ�
            If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(intRow, 0))) = True Then
                .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�), mintPriceDigit, , True)
                .TextMatrix(intRow, mconIntcol�ӳ���) = "0%"
                .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ɱ���) / .TextMatrix(intRow, mconIntCol����) * 100, mintPriceDigit, , True)
            End If
        End If
        
        If mstr��� = "" Then
            gstrSQL = "Select ����  From ҩƷ��� where ȱʡ��־=1"
            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, "SetColValue")
            
            If Not rsPrice.EOF Then
                .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsPrice!����), "", rsPrice!����)
                mstr��� = rsPrice!����
            End If
        Else
            .TextMatrix(intRow, mconIntCol���) = mstr���
        End If
        
        If mstr���ս��� = "" Then
            gstrSQL = "Select ����  From ������ս��� where ȱʡ��־=1"
            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, "SetColValue")
            
            If Not rsPrice.EOF Then
                .TextMatrix(intRow, mconIntCol���ս���) = IIf(IsNull(rsPrice!����), "", rsPrice!����)
                mstr���ս��� = rsPrice!����
            End If
        Else
            .TextMatrix(intRow, mconIntCol���ս���) = mstr���ս���
        End If
        
        If .TextMatrix(intRow, mconIntColԭ����) <> "" Then
            If mintUnit <> mconint�ۼ۵�λ And Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol��������)) = 1 Then
                .TextMatrix(intRow, mconintCol���۵�λ) = str�ۼ۵�λ
            End If
        End If
        
        '�б�ҩƷ��Ҫ��ɫ
        mblnEnter = False
        intCol = .Col
        For i = mconIntColҩ�� To .Cols - 1
            j = .ColData(i)
            If .ColData(i) = 5 Then .ColData(i) = 0
            .Col = i
            If bln�б�ҩƷ Then
                mshBill.MsfObj.CellForeColor = IIf(dbl��������� = 0, &H800000, &H800080)
            Else
                mshBill.MsfObj.CellForeColor = IIf(dbl��������� = 0, &H0, &H40&)     ' &H40C0&
            End If
            .ColData(i) = j
        Next
        .Col = intCol
        
        If (.TextMatrix(intRow, mconIntCol����) <> "" And .TextMatrix(intRow, mconIntCol��׼�ĺ�) <> "") Then
        Else
            If .TextMatrix(intRow, mconIntCol����) <> "" And .TextMatrix(intRow, mconIntCol��׼�ĺ�) = "" Then  '���ز�Ϊ�գ���׼�ĺ�Ϊ��ʱ
                gstrSQL = "select ��׼�ĺ�,�������� from ҩƷ�����̶��� where  ҩƷid=[1] and ��������=[2]"
                Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, 0), mshBill.TextMatrix(mshBill.Row, mconIntCol����))
                Do While Not rsProvider.EOF
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsProvider!��׼�ĺ�), "", rsProvider!��׼�ĺ�)
                    Exit Do
                Loop
            ElseIf (.TextMatrix(intRow, mconIntCol����) = "" And .TextMatrix(intRow, mconIntCol��׼�ĺ�) <> "") Then '����Ϊ�գ���׼�ĺŲ�Ϊ��ʱ
                gstrSQL = "select ��׼�ĺ�,�������� from ҩƷ�����̶��� where  ҩƷid=[1]"
                Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, 0))
                Do While Not rsProvider.EOF
                    .TextMatrix(mshBill.Row, mconIntCol����) = IIf(IsNull(rsProvider!��������), "", rsProvider!��������)
                    Exit Do
                Loop
            ElseIf .TextMatrix(intRow, mconIntCol����) = "" And .TextMatrix(intRow, mconIntCol��׼�ĺ�) = "" Then '����Ϊ�գ���׼�ĺ�Ϊ��ʱ
                gstrSQL = "select ��׼�ĺ�,�������� from ҩƷ�����̶��� where  ҩƷid=[1]"
                Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, 0))
                Do While Not rsProvider.EOF
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsProvider!��׼�ĺ�), "", rsProvider!��׼�ĺ�)
                    .TextMatrix(mshBill.Row, mconIntCol����) = IIf(IsNull(rsProvider!��������), "", rsProvider!��������)
                    Exit Do
                Loop
            End If

        End If
        
        If mint�༭״̬ = 8 Then
            If Val(.TextMatrix(intRow, mconIntCol�ɹ���)) = 0 Then
                MsgBox "��" & intRow & "��ҩƷ�ɱ���Ϊ���ˣ���ע��ȷ�ϣ�", vbInformation, gstrSysName
            End If
        End If
        mblnEnter = True
    End With
    SetColValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function get�ֶμӳ��ۼ�(ByVal lngҩƷID As Long, ByVal lng����ϵ�� As Long, ByVal dbl�ɹ��� As Double, ByRef dblR�ӳ��� As Double, ByRef dbl�ۼ� As Double) As Boolean
    '����:������ʱ��ҩƷ�ֶμӳ����󣬸��ݲɹ��ۼ������Ӧ���ۼ�
    '�ۼۼ��㹫ʽ�������۸���2000Ԫ/֧��ƿ��У���2000Ԫ�����µ�ҩƷ��������ۼ۸�=ʵ�ʹ����ۡ���1+����ʣ�+��۶
    '               �����۸���2000Ԫ/֧��ƿ��У�����2000Ԫ�����ϵ�ҩƷ��������ۼ۸� = ʵ�ʹ����� + ��۶�˶��Ѿ��������������ã�

    '�������ɹ���
    Dim dbl�ӳ��� As Double
    Dim dbl��۶� As Double
    Dim blnData As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    dbl�ӳ��� = 0
    dbl��۶� = 0
    
    gstrSQL = "select ��� from  �շ���ĿĿ¼ a where a.id=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ��ҩƷ���ʷ���", lngҩƷID)
    If rsTemp!��� = 7 Then
        mrs�ֶμӳ�.Filter = "����=1"
    Else
        mrs�ֶμӳ�.Filter = "����=0"
    End If
      
    If mrs�ֶμӳ�.RecordCount <> 0 Then
        mrs�ֶμӳ�.MoveFirst
        Do While Not mrs�ֶμӳ�.EOF
            With mrs�ֶμӳ�
                If dbl�ɹ��� > !��ͼ� And dbl�ɹ��� <= !��߼� Then
                    dbl�ӳ��� = IIf(IsNull(!�ӳ���), 0, !�ӳ���) / 100
                    dblR�ӳ��� = dbl�ӳ���
                    dbl��۶� = IIf(IsNull(!��۶�), 0, !��۶�)
                    blnData = True
                    Exit Do
                End If
            End With
            mrs�ֶμӳ�.MoveNext
        Loop
    End If
    
    If blnData = False Then
        If rsTemp!��� = 7 Then
            MsgBox "����ҩ��δ���ý���Ϊ��" & dbl�ɹ��� & " " & "�ķֶμӳ����ݣ��뵽ҩƷĿ¼�����зֶμӳ������ã�", vbInformation, gstrSysName
        Else
            MsgBox "����ҩ/��ҩ��δ���ý���Ϊ��" & dbl�ɹ��� & " " & "�ķֶμӳ����ݣ��뵽ҩƷĿ¼�����зֶμӳ������ã�", vbInformation, gstrSysName
        End If
        get�ֶμӳ��ۼ� = False
    End If
    
    dbl�ۼ� = dbl�ɹ��� * (1 + dbl�ӳ���) + dbl��۶�
    
    Set rsTemp = Nothing
    gstrSQL = "Select ָ�����ۼ� From ҩƷ��� Where ҩƷID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", lngҩƷID)
    If rsTemp!ָ�����ۼ� * lng����ϵ�� < dbl�ۼ� Then
        dbl�ۼ� = rsTemp!ָ�����ۼ� * lng����ϵ��
    End If
    
    get�ֶμӳ��ۼ� = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub SetInputFormat(ByVal intRow As Integer)
    If mint�༭״̬ = 5 Then '�޸ķ�Ʊ��Ϣ
        'δ����ļ�¼�����޸�
        mshBill.ColData(mconintcol��Ʊ��) = IIf(mshBill.RowData(intRow) = 0, 4, 0)
        mshBill.ColData(mconintcol��Ʊ����) = IIf(mshBill.RowData(intRow) = 0, 4, 0)
        mshBill.ColData(mconIntCol��Ʊ����) = IIf(mshBill.RowData(intRow) = 0, 2, 0)
        mshBill.ColData(mconintcol��Ʊ���) = IIf(mshBill.RowData(intRow) = 0, 4, 0)

        Exit Sub
    End If
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        If mshBill.TextMatrix(intRow, mconIntColԭ����) <> "" Then
            mshBill.ColData(mconintCol���ۼ�) = 5
            If Val(Split(mshBill.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 1 And Val(mshBill.TextMatrix(intRow, mconIntCol��������)) = 1 Then
                mshBill.ColData(mconintCol���ۼ�) = 4
            End If
        End If
    End If
    
    If mint�༭״̬ = 9 Or mint�༭״̬ = 3 Or mint�༭״̬ = 7 Then
        If mshBill.TextMatrix(intRow, mconIntColԭ����) <> "" Then
            mshBill.ColData(mconIntColЧ��) = 2                '���������
            '�����ʱ��ҩƷ�������������ۼ�
            If InStr(1, mstrControlItem, ",�ۼ�,") > 0 Then
                If Split(mshBill.TextMatrix(intRow, mconIntColԭ����), "||")(2) = 1 Then
                    mshBill.ColData(mconIntCol�ۼ�) = IIf(Getʱ��ҩƷֱ��ȷ���ۼ�, 4, 5)
                Else
                    mshBill.ColData(mconIntCol�ۼ�) = 5
                End If
            End If
        Else
            mshBill.ColData(mconIntColЧ��) = 5
        End If
    End If
    
    If mblnEdit = False Then Exit Sub
    With mshBill
        If mint�༭״̬ = 9 Or mint�༭״̬ = 3 Or mint�༭״̬ = 7 Or mint�༭״̬ = 8 Or mbln�˻� Then Exit Sub
        
        If mint�༭״̬ = 1 Then
            .ColData(mconIntCol����) = 1
            .ColData(mconIntColԭ����) = 1
        End If
        
        If .TextMatrix(intRow, mconIntColԭ����) <> "" Then
            .ColData(mconIntColЧ��) = 2                '���������
            '�����ʱ��ҩƷ�������������ۼ�
            If Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2) = 1 Then
                .ColData(mconIntCol�ۼ�) = IIf(Getʱ��ҩƷֱ��ȷ���ۼ�, 4, 5)
            Else
                .ColData(mconIntCol�ۼ�) = 5
            End If
        Else
            .ColData(mconIntColЧ��) = 5
        End If
        
        If Trim(.TextMatrix(intRow, mconintcol��Ʊ��)) = "" Then
            .ColData(mconintcol��Ʊ����) = 5
            .ColData(mconIntCol��Ʊ����) = 5
            .ColData(mconintcol��Ʊ���) = 5
        Else
            .ColData(mconintcol��Ʊ����) = 4
            .ColData(mconIntCol��Ʊ����) = 2
            .ColData(mconintcol��Ʊ���) = 4
        End If
        
    End With
End Sub

'�����ۿ�
Private Sub SetDisCount(ByVal intRow As Integer, ByVal intDisCount As Double)
    Dim dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    
    With mshBill
        'ȡԭ���ɱ���
        dbl�ɱ��� = IIf(mintʱ������ۼۼӳɷ�ʽ = 0, Val(.TextMatrix(.Row, mconIntCol�ɱ���)), Val(.TextMatrix(.Row, mconIntCol�ɹ���)))
        
        If mbln�Ӽ��� Then
            mdbl�Ӽ��� = 15
            If Val(.TextMatrix(.Row, mconIntCol�ۼ�)) <> 0 And dbl�ɱ��� <> 0 Then
                mdbl�Ӽ��� = ����ӳ���(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ۼ�)), dbl�ɱ���)
            End If
        End If
        If mshBill.Col = mconIntColָ�������� Then
            intDisCount = Val(.TextMatrix(intRow, mconIntCol����))
        Else
            .TextMatrix(intRow, mconIntCol����) = intDisCount
        End If
        
        If .TextMatrix(intRow, mconIntColָ��������) <> "" Then
'            If .TextMatrix(intRow, mconIntCol�ɹ���) = "" Then
'                .TextMatrix(intRow, mconIntCol�ɹ���) = .TextMatrix(intRow, mconIntColָ��������)
'            End If
            If Not (mint�༭״̬ = 8 Or mbln�˻�) Then
                .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol�ɹ���)) * intDisCount / 100), mintCostDigit, , True)
            End If
            If .TextMatrix(intRow, mconIntCol����) <> "" Then
               .TextMatrix(intRow, mconIntCol�ɱ����) = zlStr.FormatEx((.TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol�ɱ���)), mintMoneyDigit, , True)
               .TextMatrix(intRow, mconintcol��Ʊ���) = IIf(Trim(.TextMatrix(intRow, mconintcol��Ʊ��)) = "", "", .TextMatrix(intRow, mconIntCol�ɱ����))
            End If
            .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(intRow, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(intRow, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(intRow, mconIntCol�ɱ����) = "", 0, .TextMatrix(intRow, mconIntCol�ɱ����)), mintMoneyDigit, , True)
            
            '���ݲ�������ʱ��ҩƷ�ۼ۹�ʽ�гɱ��۵��㷨
            dbl�ɱ��� = IIf(mintʱ������ۼۼӳɷ�ʽ = 0, Val(.TextMatrix(.Row, mconIntCol�ɱ���)), Val(.TextMatrix(.Row, mconIntCol�ɹ���)))
            
            '��ʱ��ҩƷ�Ĵ���
            If .TextMatrix(intRow, mconIntColԭ����) <> "" Then
                If Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2) = 1 Then
                    '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                    If mbln�Ӽ��� Then
                        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  'û�й�ѡʱ��ȡ�ϴ��ۼ۲���
                            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), dbl�ɱ���, (mdbl�Ӽ��� / 100), dbl�ɱ��� * (1 + (mdbl�Ӽ��� / 100))), mintPriceDigit, , True)
                        End If
                    Else
                        dbl�ӳ��� = Val(Replace(.TextMatrix(intRow, mconIntcol�ӳ���), "%", "")) / 100
                        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  'û�й�ѡʱ��ȡ�ϴ��ۼ۲���
                            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), dbl�ɱ���, dbl�ӳ���, dbl�ɱ��� * (1 + dbl�ӳ���)), mintPriceDigit, , True)
                        End If
                    End If
                    If .TextMatrix(intRow, mconIntCol����) <> "" Then
                        .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol�ۼ�), mintMoneyDigit, , True)
                        'Modified by ZYB  ##2002-10-24
                        '###########################################################
                        .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(intRow, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(intRow, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(intRow, mconIntCol�ɱ����) = "", 0, .TextMatrix(intRow, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                        '###########################################################
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub mshBill_LeaveCell(Row As Long, Col As Long)
    If mblnEnter Then OS.OpenIme
    mshBill.Redraw = False
    If mblnЧ����ʾ Then
        With mshBill
            If .Col = mconIntColЧ�� Then
                CheckLapse (mshBill.TextMatrix(mshBill.Row, mconIntColЧ��))
            End If
        End With
    End If
End Sub

Private Sub mshBill_LostFocus()
    OS.OpenIme
End Sub

Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With mshBill
           If .Col = mconIntColҩ�� Then
                PopupMenu mnuCol, 2
            End If
        End With
    End If
End Sub



Private Sub msh����_DblClick()
    msh����_KeyDown vbKeyReturn, 0
End Sub

Private Sub msh����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As ADODB.Recordset
    
    On Error GoTo errHandle
    With mshBill
        If KeyCode = vbKeyEscape Then
            msh����.Visible = False
            .SetFocus
        End If
        
        If KeyCode = vbKeyReturn Then
            If .Col = mconIntCol���ս��� Then
                .TextMatrix(.Row, .Col) = msh����.TextMatrix(msh����.Row, 1)
                msh����.Visible = False
                .SetFocus
                .Col = GetNextEnableCol(.Col)
                Exit Sub
            End If

            .TextMatrix(.Row, .Col) = msh����.TextMatrix(msh����.Row, 2)
            
            gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
            Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mconIntCol����), mshBill.TextMatrix(mshBill.Row, 0))
            If Not rsProvider.EOF Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol��׼�ĺ�) = IIf(IsNull(rsProvider!��׼�ĺ�), "", rsProvider!��׼�ĺ�)
            Else
                mshBill.TextMatrix(mshBill.Row, mconIntCol��׼�ĺ�) = ""
            End If
            msh����.Visible = False
            .Col = mconIntCol����
            .SetFocus
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh����_LostFocus()
    If msh����.Visible Then
        msh����.Visible = False
    End If
End Sub

Private Sub PicInput_LostFocus()
    Dim strActive As String
    strActive = UCase(Me.ActiveControl.Name)
    
    If InStr(1, "CMDYES,CMDNO,TXT�Ӽ���", strActive) <> 0 Then
        Exit Sub
    Else
        If strActive = "MSHBILL" Then
            If mbln�����ֹ�����ӳ��� = True Then Exit Sub
        End If
    End If
    mbln�����ֹ�����ӳ��� = False
    PicInput.Visible = False
End Sub

Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And staThis.Tag <> "PY" Then
        Logogram staThis, 0
        staThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And staThis.Tag <> "WB" Then
        Logogram staThis, 1
        staThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub


Private Sub txtNO_Change()
    If txtNO.Locked = True Then
        If mstr���ݺ� <> "" And mstr���ݺ� <> txtNO.Text Then
            txtNO.Text = mstr���ݺ�
        End If
    End If
End Sub

Private Sub txtNO_GotFocus()
    If txtNO.Locked = False Then
        txtNO.SelStart = 0
        txtNO.SelLength = Len(txtNO.Text)
    End If
End Sub

Private Sub TxtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

'------------------------------------------------------------------
'------------------------------------------------------------------
'-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
' 0����ʾ���п���ѡ�񣬵������޸�
' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
'4:  ��ʾ����Ϊ�������ı����û�����
'5:  ��ʾ���в�����ѡ��
'-----------------------------------------------------------------
'-----------------------------------------------------------------
Private Sub txtProvider_Change()
    With txtProvider
        .Text = UCase(.Text)
        .SelStart = Len(.Text)
    End With
End Sub

Private Sub txtProvider_GotFocus()
    txtProvider.SelStart = 0
    txtProvider.SelLength = Len(txtProvider.Text)
End Sub

Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String
    Dim adoProvider As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txtProvider.hWnd)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint�༭״̬ = 4 Then Exit Sub
    
    On Error GoTo errHandle
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(.Text)
        gstrSQL = "Select id,����,����,���� From ��Ӧ�� " & _
                  "Where (վ�� = [2] Or վ�� is Null) And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
                  "  And ĩ��=1 And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
                  "  And (���� like [1] Or ���� like [1] or ���� like [1] )"
        'Set adoProvider = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, _
                            IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)
        Set adoProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "��ҩ��λ", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)

        If blnCancel = True Then .SetFocus: Exit Sub  '��ѡ����ʱ����Esc�������´���
        
        If adoProvider.State = 0 Then
            MsgBox "û��������Ĺ�ҩ��λ�������䣡", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If

        .Text = adoProvider!����
        .Tag = adoProvider!id
        mblnChange = True
        
        
        adoProvider.Close
        mshBill.SetFocus
        mshBill.Col = 1
        mshBill.Row = 1
        
        If CheckQualifications(Val(txtProvider.Tag)) = False Then
            txtProvider.Text = ""
            txtProvider.Tag = "0"
            Exit Sub
        End If
        
        If Val(.Tag) <> mlng��ҩ��λID And (mint�༭״̬ = 8 Or mbln�˻�) Then
            mlng��ҩ��λID = Val(txtProvider.Tag)
            mshBill.ClearBill
            mshBill.TextMatrix(1, mconIntCol�к�) = "1"
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Function ValidData() As Boolean
    Dim intLop As Integer
    Dim rsStock As New Recordset
    Dim blnStock As Boolean
    Dim strNo As String
    
    ValidData = False
    On Error GoTo errHandle
    gstrSQL = "SELECT count(*) " _
            & "From ��������˵�� " _
            & "WHERE ((�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���')) AND ����id =[1] "
    Set rsStock = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[���]", cboStock.ItemData(cboStock.ListIndex))
               
    If rsStock.Fields(0) > 0 Then
        blnStock = False
    Else
        blnStock = True
    End If
    
    If txtNO.Locked = False Then
        '�������������޸ĵ��ݺ�
        strNo = txtNO.Text
        If strNo = "" Then
            MsgBox "�����뵥�ݺš�", vbInformation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
        
        If InStr(strNo, "'") > 0 Then
            MsgBox "���ݺ����������к��зǷ��ַ���", vbExclamation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
        If LenB(StrConv(strNo, vbFromUnicode)) > 8 Then
            MsgBox "���ݺų��Ȳ��ܳ���8����ĸ��", vbExclamation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
    Else
        '��ֹ�û�ǿ���޸�
'        If mstr���ݺ� <> "" And mstr���ݺ� <> txtNO.Text Then
'            txtNO.Text = mstr���ݺ�
'        End If
    End If
    
    If CheckQualifications(Val(txtProvider.Tag)) = False Then Exit Function
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            If Val(txtProvider.Tag) = 0 Then
                MsgBox "�Բ��𣬹�ҩ��λ����Ϊ�գ�", vbOKOnly + vbInformation, gstrSysName
                txtProvider.SetFocus
                Exit Function
            End If
            
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntColҩ��)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol����))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ������Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol�ɱ���))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ�Ĳɹ���Ϊ���ˣ����飡", vbInformation, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol�ɱ���
                        Exit Function
                    End If
                    
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mconIntCol����))), vbFromUnicode)) > mintBatchNoLen Then
                        MsgBox "��" & intLop & "��ҩƷ�����ų���,���������" & Int(mintBatchNoLen / 2) & "�����ֻ�" & mintBatchNoLen & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol�ɱ����))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ�Ĳɹ����Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol�ɱ����
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol����))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ�Ŀ���Ϊ���ˣ����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If Val(Trim(Trim(.TextMatrix(intLop, mconIntCol����)))) >= 1000# Then
                        MsgBox "��" & intLop & "��ҩƷ�Ŀ���̫���ˣ����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If Split(.TextMatrix(intLop, mconIntColԭ����), "||")(0) <> "0" Then
                        If Trim(.TextMatrix(intLop, mconIntCol����)) = "" Or Trim(.TextMatrix(intLop, mconIntColЧ��)) = "" Then
                            MsgBox "��" & intLop & "�е�ҩƷ��Ч��ҩƷ,����������ż�Ч��" & vbCrLf & "��Ϣ�������뵥���У�", vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            If .TextMatrix(intLop, mconIntCol����) = "" Then
                                .Col = mconIntCol����
                            Else
                                .Col = mconIntColЧ��
                            End If
                            Exit Function
                        End If
                    End If
                    
                    '����ҩƷ����¼����غ�����
                    If Val(.TextMatrix(intLop, mconIntCol��������)) = 1 And (.TextMatrix(intLop, mconIntCol����) = "" Or .TextMatrix(intLop, mconIntCol����) = "") Then
                        MsgBox "��" & intLop & "�е�ҩƷ�Ƿ���ҩƷ,������������̺�����" & vbCrLf & "��Ϣ���뵥���У�", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        If .TextMatrix(intLop, mconIntCol����) = "" Then
                            .Col = mconIntCol����
                        Else
                            .Col = mconIntCol����
                        End If
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol�ɱ���)) > 9999999999# Then
                        MsgBox "  ��" & intLop & "��ҩƷ�Ĳɹ��۴��������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol�ɱ���
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol����)) > 9999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ���������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol�ɱ����)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ�Ĳɹ������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol�ɱ����
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol�ۼ۽��)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ���ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconintcol��Ʊ���)) > 1E+15 Then
                        MsgBox "��" & intLop & "��ҩƷ���ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ999999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintcol��Ʊ���
                        Exit Function
                    End If
                    
                    If Trim(.TextMatrix(intLop, mconintcol��Ʊ��)) <> "" Then
                        If Trim(.TextMatrix(intLop, mconIntCol��Ʊ����)) = "" Then
                            MsgBox "��" & intLop & "��ҩƷû�����뷢Ʊ���ڣ����飡", vbInformation + vbOKOnly, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mconIntCol��Ʊ����
                            .ColData(mconIntCol��Ʊ����) = 2
                            Exit Function
                        End If
                    End If
                    
                    '���۹�������Ƿ���ڲ��������۵�ҩƷ
                    If gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                        If IsPriceAdjustMod(Val(.TextMatrix(intLop, 0))) = True Then
                            If Val(.TextMatrix(intLop, mconIntCol�ɱ���)) <> Val(.TextMatrix(intLop, mconIntCol�ۼ�)) Then
                                MsgBox "��" & intLop & "��ҩƷ���������۹�������ⵥ���ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next
        Else
            Exit Function
        End If
    End With
    
    
    ValidData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtProvider_LostFocus()
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub txtProvider_Validate(Cancel As Boolean)
    If Trim(txtProvider.Text) = "" Then
        If mint�༭״̬ = 8 Or mbln�˻� Then
            mblnMSH_GetFocus = True
        End If
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If CheckQualifications(Val(txtProvider.Tag)) = False Then
        txtProvider.Text = ""
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If Val(txtProvider.Tag) <> mlng��ҩ��λID And (mint�༭״̬ = 8 Or mbln�˻�) Then
        mlng��ҩ��λID = Val(txtProvider.Tag)
        mshBill.ClearBill
        mshBill.TextMatrix(1, mconIntCol�к�) = "1"
    End If
End Sub
Private Function SaveVerifyCard(ByVal strNo As String) As Boolean
    '���ܣ��������ʱ�������˼�¼���в�������
    '����ֵ:true-ִ�гɹ� false-ִ��ʧ��
    Dim str������� As String
    
    On Error GoTo ErrHand
    
    SaveVerifyCard = False
    str������� = Format(Txt�������.Caption, "yyyy-mm-dd hh:mm:ss")
    gstrSQL = "zl_ҩƷ�������_insert("
    '�ⷿid
    gstrSQL = gstrSQL & cboStock.ItemData(cboStock.ListIndex)
    '����
    gstrSQL = gstrSQL & ",1"
    '����no
    gstrSQL = gstrSQL & ",'" & txtNO.Text & "'"
    'newNO
    gstrSQL = gstrSQL & ",'" & strNo & "'"
    '�����
    gstrSQL = gstrSQL & ",'" & UserInfo.�û����� & "'"
    '�������
    gstrSQL = gstrSQL & ",to_date('" & Format(mstr�������, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
    '��ע
    If Trim(txtժҪ.Text) = "" Then
        gstrSQL = gstrSQL & "," & "Null" & ")"
    Else
        gstrSQL = gstrSQL & ",'" & txtժҪ.Text & "')"
    End If
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
    SaveVerifyCard = True
    Exit Function
    
ErrHand:
If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveNewCard(ByVal strNo As String) As Boolean
    '���ܣ�������˲����µ����õ��ĺ���
    '����strNO���µ��ݵ�No
    '����ֵ��true-�µ��ݲ����ɹ� false-�µ��ݲ���ʧ��
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lng�Է�����id As Long
    Dim lngProviderId As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim lngBatchNO As Long
    Dim strProducingArea As String
    Dim strOldProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblDiscount As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim strBrief As String
    Dim strBooker As String
    Dim datBookDate As String
    Dim strModifier As String
    Dim datModifyDate As String
    Dim strAssessor As String
    Dim datAssessDate As String
    Dim dbl���� As Double
    Dim dbl�ӳ��� As Double
    
    Dim str��� As String
    Dim str���ս��� As String
    Dim str��Ʒ�ϸ�֤ As String
    Dim Str��Ʊ�� As String
    Dim dat��Ʊ���� As String
    Dim dbl��Ʊ��� As Double
    Dim strָ�������� As String
    Dim intRow As Integer
    Dim str��׼�ĺ� As String
    Dim str������� As String
    Dim str������� As String
    Dim str��Ʊ���� As String
    
    Dim str�˲��� As String
    Dim str�˲����� As String
    
    Dim datTimeProduct As String
    
    Dim lngRow As Integer
    Dim m As Integer
    Dim lngҩƷID As Long
    Dim lng���� As Long
    Dim i As Integer
    Dim arrSql As Variant
    
    SaveNewCard = False
    arrSql = Array()
    
    On Error GoTo errHandle
    With mshBill
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        lng�Է�����id = Val(chkת���ƿ�.Tag)
        lngProviderId = txtProvider.Tag
        strBrief = txtժҪ.Text
        strBooker = Txt������.Caption
        datBookDate = Format(Txt��������.Caption, "yyyy-mm-dd hh:mm:ss")
'        strModifier = Txt�޸���.Caption
'        datModifyDate = Format(Txt�޸�����.Caption, "yyyy-mm-dd hh:mm:ss")
        str�˲��� = txt�˲���.Caption
        str�˲����� = Format(txt�˲�����.Caption, "yyyy-mm-dd hh:mm:ss")
    
        '��ҩƷID˳���������
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lngSerial = .TextMatrix(intRow, mconIntCol���)
                lngDrugID = .TextMatrix(intRow, 0)
                strProducingArea = .TextMatrix(intRow, mconIntCol����)
                strOldProducingArea = .TextMatrix(intRow, mconIntColԭ����)
                strBatchNo = .TextMatrix(intRow, mconIntCol����)
                lngBatchNO = Val(.TextMatrix(intRow, mconIntCol����))
                datTimeProduct = IIf(Trim(.TextMatrix(intRow, mconIntCol��������)) = "", "", .TextMatrix(intRow, mconIntCol��������))
                datTimeLimit = IIf(Trim(.TextMatrix(intRow, mconIntColЧ��)) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datTimeLimit <> "" Then
                    '����ΪʧЧ��������
                    datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                End If
                
                dblQuantity = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����, , True)
                dblDiscount = .TextMatrix(intRow, mconIntCol����)
                dbl�ӳ��� = Val(Replace(.TextMatrix(intRow, mconIntcol�ӳ���), "%", "")) / 100
                dblPurchasePrice = Round(.TextMatrix(intRow, mconIntCol�ɱ���) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                dblPurchaseMoney = .TextMatrix(intRow, mconIntCol�ɱ����)
                dblSalePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�)
                dblSaleMoney = .TextMatrix(intRow, mconIntCol�ۼ۽��)
                dblMistakePrice = .TextMatrix(intRow, mconintCol���)
                
                str��� = Trim(.TextMatrix(intRow, mconIntCol���))
                str���ս��� = Trim(.TextMatrix(intRow, mconIntCol���ս���))
                str��Ʒ�ϸ�֤ = Trim(.TextMatrix(intRow, mconintcol��Ʒ�ϸ�֤))
                Str��Ʊ�� = Trim(.TextMatrix(intRow, mconintcol��Ʊ��))
                str��Ʊ���� = Trim(.TextMatrix(intRow, mconintcol��Ʊ����))
                dat��Ʊ���� = IIf(.TextMatrix(intRow, mconIntCol��Ʊ����) = "", "", .TextMatrix(intRow, mconIntCol��Ʊ����))
                dbl��Ʊ��� = IIf(Trim(.TextMatrix(intRow, mconintcol��Ʊ���)) = "", 0, .TextMatrix(intRow, mconintcol��Ʊ���))
                str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", .TextMatrix(intRow, mconIntCol��׼�ĺ�))
                str������� = IIf(Trim(.TextMatrix(intRow, mconintcol�������)) = "", "", .TextMatrix(intRow, mconintcol�������))
                str������� = IIf(Trim(.TextMatrix(intRow, mconintcol�������)) = "", "", .TextMatrix(intRow, mconintcol�������))
                
                'ʱ�۷���ҩƷ����
                If Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol��������)) = 1 And Trim(.TextMatrix(intRow, mconintCol���ۼ�)) <> "" Then
                    dblSalePrice = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol���ۼ�)), gtype_UserDrugDigits.Digit_���ۼ�)
                    dblSaleMoney = Val(.TextMatrix(intRow, mconintCol���۽��))
                    dblMistakePrice = Val(.TextMatrix(intRow, mconintCol���۲��))
                    dbl���� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol���۽��)) - Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)), mintMoneyDigit, , True)
                End If
                
                '����ҩƷĿ¼�е�ָ��������
                If mbln�޸������� Then
                    strָ�������� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntColָ��������)) / Val(.TextMatrix(intRow, mconIntCol����ϵ��)), gtype_UserDrugDigits.Digit_�ɱ���)
                    gstrSQL = "zl_ҩƷĿ¼_UpdateCustom(" & lngDrugID & ",'ָ��������=" & strָ�������� & "')"
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
                
                gstrSQL = "zl_ҩƷ�⹺_INSERT("
                'NO
                gstrSQL = gstrSQL & "'" & strNo & "'"
                '���
                gstrSQL = gstrSQL & "," & lngSerial
                '�ⷿID
                gstrSQL = gstrSQL & "," & lngStockid
                '�Է�����ID
                gstrSQL = gstrSQL & "," & IIf(lng�Է�����id <= 0, "null", lng�Է�����id)
                '��ҩ��λID
                gstrSQL = gstrSQL & "," & lngProviderId
                'ҩƷID
                gstrSQL = gstrSQL & "," & lngDrugID
                '����
                gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                '����
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                'Ч��
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                'ʵ������
                gstrSQL = gstrSQL & "," & dblQuantity
                '�ɱ���
                gstrSQL = gstrSQL & "," & dblPurchasePrice
                '�ɱ����
                gstrSQL = gstrSQL & "," & dblPurchaseMoney
                '����
                gstrSQL = gstrSQL & "," & dblDiscount
                '���ۼ�
                gstrSQL = gstrSQL & "," & dblSalePrice
                '���۽��
                gstrSQL = gstrSQL & "," & dblSaleMoney
                '���
                gstrSQL = gstrSQL & "," & dblMistakePrice
                'ժҪ
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '������
                gstrSQL = gstrSQL & ",'" & strBooker & "'"
                '��Ʊ��
                gstrSQL = gstrSQL & ",'" & Str��Ʊ�� & "'"
                '��Ʊ����
                gstrSQL = gstrSQL & "," & IIf(dat��Ʊ���� = "", "Null", "to_date('" & Format(dat��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '��Ʊ���
                gstrSQL = gstrSQL & "," & dbl��Ʊ���
                '��������
                gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                '���
                gstrSQL = gstrSQL & ",'" & str��� & "'"
                '��Ʒ�ϸ�֤
                gstrSQL = gstrSQL & ",'" & str��Ʒ�ϸ�֤ & "'"
                '�˲���
                gstrSQL = gstrSQL & "," & IIf(str�˲��� <> "", "'" & str�˲��� & "'", "NULL")
                '�˲�����
                gstrSQL = gstrSQL & "," & IIf(str�˲��� <> "", "to_date('" & str�˲����� & "','yyyy-mm-dd HH24:MI:SS')", "NULL")
                '����
                gstrSQL = gstrSQL & "," & lngBatchNO
                '�Ƿ��˻�
                gstrSQL = gstrSQL & "," & IIf(mbln�˻�, -1, 1)
                '��������
                gstrSQL = gstrSQL & "," & IIf(datTimeProduct = "", "Null", "to_date('" & Format(datTimeProduct, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '��׼�ĺ�
                gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
                '�������
                gstrSQL = gstrSQL & ",'" & str������� & "'"
                '����
                gstrSQL = gstrSQL & "," & IIf(dbl���� <> 0, dbl����, "NULL")
                '�ӳ���
                gstrSQL = gstrSQL & "," & dbl�ӳ���
                '��Ʊ����
                gstrSQL = gstrSQL & "," & IIf(str��Ʊ���� = "", "NULL", "'" & str��Ʊ���� & "'")
                '�ƻ�id
                gstrSQL = gstrSQL & ",NULL"
                '�������
                gstrSQL = gstrSQL & ",2"
                'ԭ����
                gstrSQL = gstrSQL & ",'" & strOldProducingArea & "'"
                '�������
                gstrSQL = gstrSQL & "," & IIf(str������� <> "", "to_date('" & str������� & "','yyyy-mm-dd HH24:MI:SS')", "Null")
                '���ս���
                gstrSQL = gstrSQL & ",'" & str���ս��� & "'"
                '�޸���
                gstrSQL = gstrSQL & ",'" & strModifier & "'"
                '�޸�����
                gstrSQL = gstrSQL & "," & IIf(datModifyDate = "", "Null", "to_date('" & datModifyDate & "','yyyy-mm-dd HH24:MI:SS')")
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        Next
        
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveNewCard")
        Next
        
        SaveNewCard = True
        mstr���ݺ� = chrNo
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveCard(Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lng�Է�����id As Long
    Dim lngProviderId As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim lngBatchNO As Long
    Dim strProducingArea As String
    Dim strOldProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblDiscount As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim strBrief As String
    Dim strBooker As String
    Dim datBookDate As String
    Dim strModifier As String
    Dim datModifyDate As String
    Dim strAssessor As String
    Dim datAssessDate As String
    Dim dbl���� As Double
    Dim dbl�ӳ��� As Double
    
    Dim str��� As String
    Dim str���ս��� As String
    Dim str��Ʒ�ϸ�֤ As String
    Dim Str��Ʊ�� As String
    Dim dat��Ʊ���� As String
    Dim dbl��Ʊ��� As Double
    Dim strָ�������� As String
    Dim intRow As Integer
    Dim str��׼�ĺ� As String
    Dim str������� As String
    Dim str������� As String
    Dim str��Ʊ���� As String
    Dim lng�ƻ�id As Long
    
    Dim str�˲��� As String
    Dim str�˲����� As String
    Dim rsTemp As New ADODB.Recordset
    
    Dim datTimeProduct As String
    
    Dim n As Integer
    Dim m As Integer
    Dim dbl�ϼ����� As Double
    Dim lngҩƷID As Long
    Dim lng���� As Long
    Dim i As Integer
    Dim arrSql As Variant
    
    SaveCard = False
    arrSql = Array()
    If Not Check��ͬ��λ Then Exit Function
    If Not CheckProvider Then Exit Function
    
    On Error GoTo errHandle
    With mshBill
        chrNo = Trim(txtNO)
        If chrNo = "" Then chrNo = Sys.GetNextNo(21, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(chrNo) Then Exit Function
        
        If mint�༭״̬ = 1 Then
            If CheckNOExists(1, chrNo) Then
                MsgBox "������ͬ���ݺŵ��⹺��ⵥ�����鵥�ݺ��Ƿ���ȷ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        Me.txtNO.Tag = chrNo
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        lngProviderId = txtProvider.Tag
        
        strBrief = Trim(txtժҪ.Text)
        strBooker = Txt������
        

        If mint�༭״̬ = 9 Then '9-�˲�
            '�޸���Ϣ
            strModifier = Txt�޸���
            datModifyDate = Format(Txt�޸�����, "yyyy-mm-dd hh:mm:ss")
            
            If IsDate(Txt��������) Then
                datBookDate = Format(Txt��������.Caption, "yyyy-mm-dd hh:mm:ss")
            Else
                datBookDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            End If
        Else
            datBookDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        End If
        
        
        strAssessor = Txt�����
        
        'ȡԭʼ���ݵĺ˲���
        If mint�༭״̬ <> 9 Then
            gstrSQL = "Select ��ҩ��,to_Char(��ҩ����,'yyyy-MM-dd hh24:mi:ss') ��ҩ���� " & _
                " From ҩƷ�շ���¼ Where ����=1 And NO=[1] "
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡԭʼ���ݵĺ˲���]", chrNo)
                
            If Not rsTemp.EOF Then
                str�˲��� = nvl(rsTemp!��ҩ��)
                str�˲����� = nvl(rsTemp!��ҩ����)
            End If
            If mint�༭״̬ = 2 Then
                '�޸ĵ���������˲��ˣ���Ҫ�ٴκ˲�
                str�˲��� = ""
                str�˲����� = ""
            End If
        Else
            str�˲��� = Txt�����.Caption
            str�˲����� = Txt�������.Caption
        End If
                
        If mint�༭״̬ = 2 Or mint�༭״̬ = 9 Or blnǿ�Ʊ��� Then        '�޸�
            gstrSQL = "zl_ҩƷ�⹺_Delete('" & mstr���ݺ� & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            strBooker = Txt������
            datBookDate = Format(Txt��������.Caption, "yyyy-mm-dd hh:mm:ss")
            '�޸���Ϣ
            If mint�༭״̬ = 9 Then
                strModifier = Txt�޸���
                datModifyDate = Format(Txt�޸�����.Caption, "yyyy-mm-dd hh:mm:ss")
            Else
                strModifier = UserInfo.�û�����
                datModifyDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            End If
        End If
            
        lng�Է�����id = Val(chkת���ƿ�.Tag)
    
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
'                If Val(.TextMatrix(intRow, mconIntCol���)) = 0 Then
'                    lngSerial = intRow
'                Else
'                    lngSerial = Val(.TextMatrix(intRow, mconIntCol���))
'                End If
                lngSerial = intRow
                
                lngDrugID = .TextMatrix(intRow, 0)
                strProducingArea = .TextMatrix(intRow, mconIntCol����)
                strOldProducingArea = .TextMatrix(intRow, mconIntColԭ����)
                strBatchNo = .TextMatrix(intRow, mconIntCol����)
                lngBatchNO = Val(.TextMatrix(intRow, mconIntCol����))
                datTimeProduct = IIf(Trim(.TextMatrix(intRow, mconIntCol��������)) = "", "", .TextMatrix(intRow, mconIntCol��������))
                datTimeLimit = IIf(Trim(.TextMatrix(intRow, mconIntColЧ��)) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datTimeLimit <> "" Then
                    '����ΪʧЧ��������
                    datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                End If
                
                dblQuantity = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����, , True)
                dblDiscount = .TextMatrix(intRow, mconIntCol����)
                dbl�ӳ��� = Val(Replace(.TextMatrix(intRow, mconIntcol�ӳ���), "%", "")) / 100
                dblPurchasePrice = Round(.TextMatrix(intRow, mconIntCol�ɱ���) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                dblPurchaseMoney = .TextMatrix(intRow, mconIntCol�ɱ����)
                dblSalePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�)
                dblSaleMoney = .TextMatrix(intRow, mconIntCol�ۼ۽��)
                dblMistakePrice = .TextMatrix(intRow, mconintCol���)
                
                If Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 0 And mintUnit <> 4 Then
                    '����Ƕ���ҩƷ�����ۼ�ȡԭʼ�۸񱣴�
                    dblSalePrice = Get�ۼ�(Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 1, lngDrugID, lngStockid, 0)
                                    
                    If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(lngDrugID) = True Then
                        '�����ʵ�����۹����ҩƷ���ɱ���ҲҪ���ۼ�һ��
                        dblPurchasePrice = dblSalePrice
                    End If
                End If
                
                str��� = Trim(.TextMatrix(intRow, mconIntCol���))
                str���ս��� = Trim(.TextMatrix(intRow, mconIntCol���ս���))
                str��Ʒ�ϸ�֤ = Trim(.TextMatrix(intRow, mconintcol��Ʒ�ϸ�֤))
                Str��Ʊ�� = Trim(.TextMatrix(intRow, mconintcol��Ʊ��))
                str��Ʊ���� = Trim(.TextMatrix(intRow, mconintcol��Ʊ����))
                dat��Ʊ���� = IIf(.TextMatrix(intRow, mconIntCol��Ʊ����) = "", "", .TextMatrix(intRow, mconIntCol��Ʊ����))
                dbl��Ʊ��� = IIf(Trim(.TextMatrix(intRow, mconintcol��Ʊ���)) = "", 0, .TextMatrix(intRow, mconintcol��Ʊ���))
                str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", .TextMatrix(intRow, mconIntCol��׼�ĺ�))
                str������� = IIf(Trim(.TextMatrix(intRow, mconintcol�������)) = "", "", .TextMatrix(intRow, mconintcol�������))
                str������� = IIf(Trim(.TextMatrix(intRow, mconintcol�������)) = "", "", .TextMatrix(intRow, mconintcol�������))
                lng�ƻ�id = IIf(Trim(.TextMatrix(intRow, mconIntCol�ƻ�id)) = "", 0, Val(.TextMatrix(intRow, mconIntCol�ƻ�id)))
                
                'ʱ�۷���ҩƷ����
                If Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol��������)) = 1 And Trim(.TextMatrix(intRow, mconintCol���ۼ�)) <> "" Then
                    dblSalePrice = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol���ۼ�)), gtype_UserDrugDigits.Digit_���ۼ�)
                    dblSaleMoney = Val(.TextMatrix(intRow, mconintCol���۽��))
                    dblMistakePrice = Val(.TextMatrix(intRow, mconintCol���۲��))
                    dbl���� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol���۽��)) - Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)), mintMoneyDigit, , True)
                End If
  
                '����ҩƷĿ¼�е�ָ��������
                If mbln�޸������� Then
                    strָ�������� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntColָ��������)) / Val(.TextMatrix(intRow, mconIntCol����ϵ��)), gtype_UserDrugDigits.Digit_�ɱ���)
                    gstrSQL = "zl_ҩƷĿ¼_UpdateCustom(" & lngDrugID & ",'ָ��������=" & strָ�������� & "')"
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
                
                gstrSQL = "zl_ҩƷ�⹺_INSERT("
                'NO
                gstrSQL = gstrSQL & "'" & chrNo & "'"
                '���
                gstrSQL = gstrSQL & "," & lngSerial
                '�ⷿID
                gstrSQL = gstrSQL & "," & lngStockid
                '�Է�����ID
                gstrSQL = gstrSQL & "," & IIf(lng�Է�����id <= 0, "null", lng�Է�����id)
                '��ҩ��λID
                gstrSQL = gstrSQL & "," & lngProviderId
                'ҩƷID
                gstrSQL = gstrSQL & "," & lngDrugID
                '����
                gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                '����
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                'Ч��
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                'ʵ������
                gstrSQL = gstrSQL & "," & dblQuantity
                '�ɱ���
                gstrSQL = gstrSQL & "," & dblPurchasePrice
                '�ɱ����
                gstrSQL = gstrSQL & "," & dblPurchaseMoney
                '����
                gstrSQL = gstrSQL & "," & dblDiscount
                '���ۼ�
                gstrSQL = gstrSQL & "," & dblSalePrice
                '���۽��
                gstrSQL = gstrSQL & "," & dblSaleMoney
                '���
                gstrSQL = gstrSQL & "," & dblMistakePrice
                'ժҪ
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '������
                gstrSQL = gstrSQL & ",'" & strBooker & "'"
                '��Ʊ��
                gstrSQL = gstrSQL & ",'" & Str��Ʊ�� & "'"
                '��Ʊ����
                gstrSQL = gstrSQL & "," & IIf(dat��Ʊ���� = "", "Null", "to_date('" & Format(dat��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '��Ʊ���
                gstrSQL = gstrSQL & "," & dbl��Ʊ���
                '��������
                gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                '���
                gstrSQL = gstrSQL & ",'" & str��� & "'"
                '��Ʒ�ϸ�֤
                gstrSQL = gstrSQL & ",'" & str��Ʒ�ϸ�֤ & "'"
                '�˲���
                gstrSQL = gstrSQL & "," & IIf(str�˲��� <> "", "'" & str�˲��� & "'", "NULL")
                '�˲�����
                gstrSQL = gstrSQL & "," & IIf(str�˲��� <> "", "to_date('" & str�˲����� & "','yyyy-mm-dd HH24:MI:SS')", "NULL")
                '����
                gstrSQL = gstrSQL & "," & lngBatchNO
                '�Ƿ��˻�
                gstrSQL = gstrSQL & "," & IIf(mbln�˻�, -1, 1)
                '��������
                gstrSQL = gstrSQL & "," & IIf(datTimeProduct = "", "Null", "to_date('" & Format(datTimeProduct, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '��׼�ĺ�
                gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
                '�������
                gstrSQL = gstrSQL & ",'" & str������� & "'"
                '����
                gstrSQL = gstrSQL & "," & IIf(dbl���� <> 0, dbl����, "NULL")
                '�ӳ���
                gstrSQL = gstrSQL & "," & dbl�ӳ���
                '��Ʊ����
                gstrSQL = gstrSQL & "," & IIf(str��Ʊ���� = "", "NULL", "'" & str��Ʊ���� & "'")
                '�ƻ�id
                gstrSQL = gstrSQL & "," & IIf(lng�ƻ�id = 0, "NULL", lng�ƻ�id)
                '�������
                gstrSQL = gstrSQL & "," & 0
                'ԭ����
                gstrSQL = gstrSQL & ",'" & strOldProducingArea & "'"
                '�������
                gstrSQL = gstrSQL & "," & IIf(str������� <> "", "to_date('" & str������� & "','yyyy-mm-dd HH24:MI:SS')", "NULL")
                '���ս���
                gstrSQL = gstrSQL & ",'" & str���ս��� & "'"
                '�޸���
                gstrSQL = gstrSQL & ",'" & strModifier & "'"
                '�޸�����
                gstrSQL = gstrSQL & "," & IIf(datModifyDate = "", "Null", "to_date('" & datModifyDate & "','yyyy-mm-dd HH24:MI:SS')")
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        If Not blnǿ�Ʊ��� Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        If Not blnǿ�Ʊ��� Then gcnOracle.CommitTrans
        mstr���ݺ� = chrNo
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    If Not blnǿ�Ʊ��� Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'�˻�
Private Function SaveRestore() As Boolean
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lngProviderId As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim strProducingArea As String
    Dim strOldProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblDiscount As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim strBrief As String
    Dim strBooker As String
    Dim datBookDate As String
    Dim strModifier As String
    Dim datModifyDate As String
    Dim strAssessor As String
    Dim datAssessDate As String
    Dim str��� As String
    Dim str���ս��� As String
    Dim str��Ʒ�ϸ�֤ As String
    Dim Str��Ʊ�� As String
    Dim str��Ʊ���� As String
    Dim dat��Ʊ���� As String
    Dim dbl��Ʊ��� As Double
    Dim strָ�������� As String
    Dim intRow As Integer
    Dim dbl���� As Double
    Dim dbl�ӳ��� As Double
    Dim str��׼�ĺ� As String
    Dim str�˲��� As String
    Dim str�˲����� As String
    Dim rsTemp As New ADODB.Recordset
    
    Dim datTimeProduct As String
    Dim n As Integer
    Dim m As Integer
    Dim dbl�ϼ����� As Double
    Dim lngҩƷID As Long
    Dim lng���� As Long
    Dim str������� As String
    Dim str������� As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim blnTran As Boolean  '�Ƿ�ʼ������
    Dim intLop As Integer
    
    On Error GoTo errHandle
    
    SaveRestore = False
    'ֻ��ҩ�������ʹ���˻�����
    arrSql = Array()
    If Val(txtProvider.Tag) = 0 Then
        MsgBox "��ѡ��Ӧ�̣�", vbInformation, gstrSysName
        txtProvider.SetFocus
        Exit Function
    End If
    
    With mshBill
        If .TextMatrix(1, 0) = "" Then Exit Function
        
        chrNo = Trim(txtNO)
        If chrNo = "" Then chrNo = Sys.GetNextNo(21, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(chrNo) Then Exit Function
        Me.txtNO.Tag = chrNo
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        lngProviderId = Val(txtProvider.Tag)
        strBrief = Trim(txtժҪ.Text)
        strBooker = Txt������
        datBookDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
'        strModifier = Txt�޸���
'        datModifyDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        strAssessor = Txt�����
        
        'ȡԭ�˲���
        gstrSQL = "Select ��ҩ��,to_Char(��ҩ����,'yyyy-MM-dd hh24:mi:ss') ��ҩ���� From ҩƷ�շ���¼ Where ����=1 And NO=[1] "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡԭʼ���ݵĺ˲���]", chrNo)
        
        If Not rsTemp.EOF Then
            str�˲��� = nvl(rsTemp!��ҩ��)
            str�˲����� = nvl(rsTemp!��ҩ����)
        End If
        
        On Error GoTo errHandle
            
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                lngDrugID = .TextMatrix(intRow, 0)
                strProducingArea = .TextMatrix(intRow, mconIntCol����)
                strOldProducingArea = .TextMatrix(intRow, mconIntColԭ����)
                strBatchNo = .TextMatrix(intRow, mconIntCol����)
                datTimeProduct = IIf(Trim(.TextMatrix(intRow, mconIntCol��������)) = "", "", .TextMatrix(intRow, mconIntCol��������))
                datTimeLimit = IIf(Trim(.TextMatrix(intRow, mconIntColЧ��)) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", .TextMatrix(intRow, mconIntCol��׼�ĺ�))
                dblQuantity = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol����)) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����, , True)
                dblDiscount = .TextMatrix(intRow, mconIntCol����)
                dbl�ӳ��� = Val(Replace(.TextMatrix(intRow, mconIntcol�ӳ���), "%", "")) / 100
                dblPurchasePrice = Round(Val(.TextMatrix(intRow, mconIntCol�ɱ���)) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                dblPurchaseMoney = Val(.TextMatrix(intRow, mconIntCol�ɱ����))
                dblSalePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�)
                dblSaleMoney = Val(.TextMatrix(intRow, mconIntCol�ۼ۽��))
                dblMistakePrice = Val(.TextMatrix(intRow, mconintCol���))
                lngSerial = intRow
                
                str��� = Trim(.TextMatrix(intRow, mconIntCol���))
                str���ս��� = Trim(.TextMatrix(intRow, mconIntCol���ս���))
                str��Ʒ�ϸ�֤ = Trim(.TextMatrix(intRow, mconintcol��Ʒ�ϸ�֤))
                Str��Ʊ�� = Trim(.TextMatrix(intRow, mconintcol��Ʊ��))
                str��Ʊ���� = Trim(.TextMatrix(intRow, mconintcol��Ʊ����))
                dat��Ʊ���� = IIf(.TextMatrix(intRow, mconIntCol��Ʊ����) = "", "", .TextMatrix(intRow, mconIntCol��Ʊ����))
                dbl��Ʊ��� = IIf(Trim(.TextMatrix(intRow, mconintcol��Ʊ���)) = "", 0, .TextMatrix(intRow, mconintcol��Ʊ���))
                str������� = IIf(Trim(.TextMatrix(intRow, mconintcol�������)) = "", "", .TextMatrix(intRow, mconintcol�������))
                str������� = IIf(Trim(.TextMatrix(intRow, mconintcol�������)) = "", "", .TextMatrix(intRow, mconintcol�������))
                
                'ʱ�۷���ҩƷ����
                If Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol��������)) = 1 And Trim(.TextMatrix(intRow, mconintCol���ۼ�)) <> "" Then
                    dblSalePrice = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol���ۼ�)), gtype_UserDrugDigits.Digit_���ۼ�)
                    dblSaleMoney = Val(.TextMatrix(intRow, mconintCol���۽��))
                    dblMistakePrice = Val(.TextMatrix(intRow, mconintCol���۲��))
                    dbl���� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol���۽��)) - Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)), mintMoneyDigit, , True)
                End If
                
                '����ҩƷĿ¼�е�ָ��������
                If mbln�޸������� Then
                    strָ�������� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntColָ��������)) / Val(.TextMatrix(intRow, mconIntCol����ϵ��)), gtype_UserDrugDigits.Digit_�ɱ���)
                    gstrSQL = "zl_ҩƷĿ¼_UpdateCustom(" & lngDrugID & ",'ָ��������=" & strָ�������� & "')"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
                
                If dblQuantity = 0 Then
                    MsgBox "��" & lngSerial & "�е��˻�����Ϊ�㣬�������浥�ݣ�", vbInformation, gstrSysName
                    Exit Function
                End If
                
                gstrSQL = "zl_ҩƷ�⹺_INSERT("
                'NO
                gstrSQL = gstrSQL & "'" & chrNo & "'"
                '���
                gstrSQL = gstrSQL & "," & lngSerial
                '�ⷿID
                gstrSQL = gstrSQL & "," & lngStockid
                '�Է�����ID
                gstrSQL = gstrSQL & ",null"
                '��ҩ��λID
                gstrSQL = gstrSQL & "," & lngProviderId
                'ҩƷID
                gstrSQL = gstrSQL & "," & lngDrugID
                '����
                gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                '����
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                'Ч��
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                'ʵ������
                gstrSQL = gstrSQL & "," & dblQuantity
                '�ɱ���
                gstrSQL = gstrSQL & "," & dblPurchasePrice
                '�ɱ����
                gstrSQL = gstrSQL & "," & dblPurchaseMoney
                '����
                gstrSQL = gstrSQL & "," & dblDiscount
                '���ۼ�
                gstrSQL = gstrSQL & "," & dblSalePrice
                '���۽��
                gstrSQL = gstrSQL & "," & dblSaleMoney
                '���
                gstrSQL = gstrSQL & "," & dblMistakePrice
                'ժҪ
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '������
                gstrSQL = gstrSQL & ",'" & strBooker & "'"
                '��Ʊ��
                gstrSQL = gstrSQL & ",'" & Str��Ʊ�� & "'"
                '��Ʊ����
                gstrSQL = gstrSQL & "," & IIf(dat��Ʊ���� = "", "Null", "to_date('" & Format(dat��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '��Ʊ���
                gstrSQL = gstrSQL & "," & dbl��Ʊ���
                '��������
                gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                '���
                gstrSQL = gstrSQL & ",'" & str��� & "'"
                '��Ʒ�ϸ�֤
                gstrSQL = gstrSQL & ",'" & str��Ʒ�ϸ�֤ & "'"
                '�˲���
                gstrSQL = gstrSQL & "," & IIf(str�˲��� <> "", "'" & str�˲��� & "'", "NULL")
                '�˲�����
                gstrSQL = gstrSQL & "," & IIf(str�˲��� <> "", "to_date('" & str�˲����� & "','yyyy-mm-dd HH24:MI:SS')", "NULL")
                '����
                gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, mconIntCol����))
                '�Ƿ��˻�
                gstrSQL = gstrSQL & ",-1"
                '��������
                gstrSQL = gstrSQL & "," & IIf(datTimeProduct = "", "Null", "to_date('" & Format(datTimeProduct, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '��׼�ĺ�
                gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
                '�������
                gstrSQL = gstrSQL & ",'" & str������� & "'"
                '����
                gstrSQL = gstrSQL & "," & IIf(dbl���� <> 0, dbl����, "NULL")
                '�ӳ���
                gstrSQL = gstrSQL & "," & dbl�ӳ���
                '��Ʊ����
                gstrSQL = gstrSQL & "," & IIf(str��Ʊ���� = "", "NULL", "'" & str��Ʊ���� & "'")
                '�ƻ�id
                gstrSQL = gstrSQL & ",NULL"
                '�������
                gstrSQL = gstrSQL & ",0"
                'ԭ����
                gstrSQL = gstrSQL & ",'" & strOldProducingArea & "'"
                '�������
                gstrSQL = gstrSQL & ",NULL"
                '���ս���
                gstrSQL = gstrSQL & ",NULL"
                '�޸���
                gstrSQL = gstrSQL & ",'" & strModifier & "'"
                '�޸�����
                gstrSQL = gstrSQL & "," & IIf(datModifyDate = "", "Null", "to_date('" & datModifyDate & "','yyyy-mm-dd HH24:MI:SS')")
                gstrSQL = gstrSQL & ")"
                    
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        gcnOracle.BeginTrans
        blnTran = True
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveRestore")
        Next
        gcnOracle.CommitTrans
        
        mstr���ݺ� = chrNo
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveRestore = True
    Exit Function
errHandle:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'�������
Private Function SaveStrike() As Boolean
    Dim �д�_IN As Integer
    Dim ԭ��¼״̬_IN As Integer
    Dim NO_IN As String
    Dim ���_IN As Integer
    Dim ҩƷID_IN As Long
    Dim ��������_IN As Double
    Dim ������_IN As String
    Dim ��������_IN  As String
    Dim ��Ʊ��_IN As String
    Dim ��Ʊ����_In As String
    Dim ��Ʊ����_IN As String
    Dim ��Ʊ���_IN As Double
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim n As Integer
    Dim intȫ������ As Integer
    Dim ժҪ_IN As String
    Dim strҩƷID As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim strҩƷ As String
    Dim intNumCol As Integer
    
    arrSql = Array()
    SaveStrike = False
    With mshBill
        'Ϊ���Ⲣ�����������¸��¸����־
        Call Refresh�����־
        
        '���������������ű�����ԭʼ������ͬ���Ѹ���ļ�¼�����������������˵ĵ���Ҳ�����������
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, mconIntCol��������)) <> 0 Then
                If Not ��ͬ����(Val(.TextMatrix(intRow, mconIntCol����)), Val(.TextMatrix(intRow, mconIntCol��������))) Then
                    MsgBox "������Ϸ��ĳ�����������" & intRow & "�У���", vbInformation, gstrSysName
                    .MsfObj.TopRow = intRow
                    Exit Function
                End If
                If .RowData(intRow) <> 0 Then
                    MsgBox "��" & intRow & "�е�ҩƷ�Ѿ���������������", vbInformation, gstrSysName
                    .MsfObj.TopRow = intRow
                    Exit Function
                End If
            End If
        Next
        
        If mint�༭״̬ = 6 Then '����
            intNumCol = mconIntCol��������
        Else
            intNumCol = mconIntCol����
        End If
        '�����
        If mint����� <> 0 And mint�༭״̬ <> 7 And mbln�˻� = False Then '�˻������Ͳ�����˲��ü��
            strҩƷ = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol����, intNumCol, mconIntCol����ϵ��, 2, , mintNumberDigit)
            If strҩƷ <> "" Then
                If mbln��ʾ��ʽ = False Then
                    If mint����� = 1 Then '��������
                        If MsgBox("ҩƷ��" & strҩƷ & "����治�㣬�Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Function
                        Else
                            mbln��ʾ��ʽ = True
                        End If
                    ElseIf mint����� = 2 Then '�����ֹ
                        MsgBox "ҩƷ��" & strҩƷ & "����治�㣬������ˣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
        
        NO_IN = Trim(txtNO.Tag)
        ������_IN = UserInfo.�û�����
        ԭ��¼״̬_IN = mint��¼״̬
        ժҪ_IN = Trim(txtժҪ.Text)
        
        On Error GoTo errHandle
        
        �д�_IN = 0
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" And (Val(.TextMatrix(intRow, mconIntCol��������)) <> 0 Or mint�༭״̬ = 7) Then
                �д�_IN = �д�_IN + 1
                ҩƷID_IN = .TextMatrix(intRow, 0)
                strҩƷID = IIf(strҩƷID = "", "", strҩƷID & ",") & ҩƷID_IN
                If Val(.TextMatrix(intRow, mconIntCol��������)) = Val(.TextMatrix(intRow, mconIntCol����)) Then
                    intȫ������ = 1
                Else
                    intȫ������ = 0
                End If
                ��������_IN = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol��������)) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����, , True)
                ��Ʊ��_IN = Trim(.TextMatrix(intRow, mconintcol��Ʊ��))
                ��Ʊ����_In = Trim(.TextMatrix(intRow, mconintcol��Ʊ����))
                ��Ʊ����_IN = IIf(.TextMatrix(intRow, mconIntCol��Ʊ����) = "", "", .TextMatrix(intRow, mconIntCol��Ʊ����))
                ��Ʊ���_IN = IIf(Trim(.TextMatrix(intRow, mconintcol��Ʊ���)) = "", 0, .TextMatrix(intRow, mconintcol��Ʊ���))
                ���_IN = .TextMatrix(intRow, mconIntCol���)
                
                gstrSQL = "ZL_ҩƷ�⹺_STRIKE("
                '�д�
                gstrSQL = gstrSQL & �д�_IN
                'ԭ��¼״̬
                gstrSQL = gstrSQL & "," & ԭ��¼״̬_IN
                'NO
                gstrSQL = gstrSQL & ",'" & NO_IN & "'"
                '���
                gstrSQL = gstrSQL & "," & ���_IN
                'ҩƷID
                gstrSQL = gstrSQL & "," & ҩƷID_IN
                '��������
                gstrSQL = gstrSQL & "," & ��������_IN
                '������
                gstrSQL = gstrSQL & ",'" & ������_IN & "'"
                '��������
                gstrSQL = gstrSQL & ",to_date('" & Format(mstr�������, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
                '��Ʊ��
                gstrSQL = gstrSQL & "," & IIf(��Ʊ��_IN = "", "Null", "'" & ��Ʊ��_IN & "'")
                '��Ʊ���
                gstrSQL = gstrSQL & "," & ��Ʊ���_IN
                '�Ƿ�ȫ������
                gstrSQL = gstrSQL & "," & IIf(mint�༭״̬ = 7 Or intȫ������ = 1, 1, 0)
                '�Ƿ�������
                gstrSQL = gstrSQL & "," & IIf(mint�༭״̬ = 7, 1, 0)
                'ժҪ
                gstrSQL = gstrSQL & ",'" & ժҪ_IN & "'"
                '��Ʊ����
                gstrSQL = gstrSQL & "," & IIf(��Ʊ����_In = "", "NULL", "'" & ��Ʊ����_In & "'")
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        If mint�༭״̬ <> 7 Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        If mint�༭״̬ <> 7 Then gcnOracle.CommitTrans
        
        If �д�_IN = 0 Then
            MsgBox "û��ѡ��һ��ҩƷ����������¼�����������", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        '��ʾͣ��ҩƷ
        If strҩƷID <> "" Then
            Call CheckStopMedi(strҩƷID)
        End If
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveStrike = True
    Exit Function
errHandle:
    If mint�༭״̬ <> 7 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveRecipe(Optional ByVal strNewNO As String = "") As Boolean
    Dim chrNo As String
    Dim lng��� As Long
    Dim Str��Ʊ�� As String
    Dim str��Ʊ���� As String
    Dim dat��Ʊ���� As String
    Dim dbl��Ʊ��� As Double
    Dim int������־ As Integer '1��δ���������޸ķ�Ʊ��Ϣ; 2�����ֳ��������޸ķ�Ʊ��Ϣ
    Dim intRow As Integer
    Dim n As Integer
    Dim i As Integer
    Dim arrSql As Variant
    
    arrSql = Array()
    SaveRecipe = False
    '����Ƿ����빩ҩ��λ
    If Val(txtProvider.Tag) = 0 Then
        MsgBox "��ѡ��ҩƷ��Ӧ�̣�", vbInformation, gstrSysName
        txtProvider.SetFocus
        Exit Function
    End If
    If Not Check��ͬ��λ Then Exit Function
    If Not CheckProvider Then Exit Function
        
    With mshBill
        If strNewNO = "" Then
            chrNo = Trim(txtNO)
        Else
            chrNo = strNewNO
        End If
        
        On Error GoTo errHandle
        
        'Ϊ���Ⲣ�����������¸��¸����־
        Call Refresh�����־
        
        '�����ⵥ���Ƿ��Ѹ���
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, mconIntCol����)) <> 0 Then
                If .RowData(intRow) <> 0 Then
                    MsgBox "��" & intRow & "�е�ҩƷ�Ѿ���������޸ĸ�ҩƷ�ķ�Ʊ��Ϣ��", vbInformation, gstrSysName
                End If
            End If
        Next
                
        
        If mint�༭״̬ = 5 Then
            If mint��¼״̬ = 1 Then
                int������־ = 1
            Else
                int������־ = 2
            End If
        ElseIf mint�༭״̬ = 7 Then
            int������־ = 1
        End If
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                If .RowData(intRow) = 0 Then
'                    If strNewNO = "" Then
'                        lng��� = Val(.TextMatrix(intRow, mconIntCol���))
'                    Else
'                        lng��� = intRow
'                    End If
                    lng��� = Val(.TextMatrix(intRow, mconIntCol���))
                    Str��Ʊ�� = Trim(.TextMatrix(intRow, mconintcol��Ʊ��))
                    str��Ʊ���� = Trim(.TextMatrix(intRow, mconintcol��Ʊ����))
                    dat��Ʊ���� = IIf(.TextMatrix(intRow, mconIntCol��Ʊ����) = "", "", .TextMatrix(intRow, mconIntCol��Ʊ����))
                    dbl��Ʊ��� = IIf(mbln�˻�, -1, 1) * IIf(Trim(.TextMatrix(intRow, mconintcol��Ʊ���)) = "", 0, .TextMatrix(intRow, mconintcol��Ʊ���))
                    
                    gstrSQL = "zl_ҩƷ�⹺��Ʊ��Ϣ_UPDATE("
                    'NO
                    gstrSQL = gstrSQL & "'" & chrNo & "'"
                    '���
                    gstrSQL = gstrSQL & "," & lng���
                    '��Ʊ��
                    gstrSQL = gstrSQL & ",'" & Str��Ʊ�� & "'"
                    '��Ʊ����
                    gstrSQL = gstrSQL & "," & IIf(dat��Ʊ���� = "", "Null", "to_date('" & Format(dat��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                    '��Ʊ���
                    gstrSQL = gstrSQL & "," & dbl��Ʊ���
                    '��ҩ��λID
                    gstrSQL = gstrSQL & "," & Val(txtProvider.Tag)
                    '������־
                    gstrSQL = gstrSQL & "," & int������־
                    '��Ʊ����
                    gstrSQL = gstrSQL & ",'" & str��Ʊ���� & "'"
                    gstrSQL = gstrSQL & ")"

                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
            End If
            
            recSort.MoveNext
        Next
        
        If mint�༭״̬ <> 7 Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        If mint�༭״̬ <> 7 Then gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveRecipe = True
    Exit Function
errHandle:
    If mint�༭״̬ <> 7 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ��ʾ�ϼƽ��()
    Dim curTotal As Double, Cur���ʽ�� As Double, Cur���ʲ�� As Double
    Dim intLop As Integer
    Dim dblʱ�۷��� As Boolean
    
    curTotal = 0: Cur���ʽ�� = 0: Cur���ʲ�� = 0:
    
    With mshBill
        For intLop = 1 To .rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mconIntCol�ɱ����))
'            Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
            If .TextMatrix(intLop, mconIntColԭ����) <> "" Then
                If Val(Split(.TextMatrix(intLop, mconIntColԭ����), "||")(2)) = 1 And Val(.TextMatrix(intLop, mconIntCol��������)) = 1 Then
                    dblʱ�۷��� = True
                    Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconintCol���۽��))
                Else
                    Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
                End If
            Else
                Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
            End If
        Next
    End With
    
    Cur���ʲ�� = Cur���ʽ�� - curTotal
    lblPurchasePrice.Caption = "�ɱ����ϼƣ�" & zlStr.FormatEx(curTotal, mintMoneyDigit, , True)
    
    If dblʱ�۷��� = True Then
        lblSalePrice.Caption = "�ۼ۽��(ʱ�۷��������۽��)�ϼƣ�" & zlStr.FormatEx(Cur���ʽ��, mintMoneyDigit, , True)
        lblDifference.Caption = "���(ʱ�۷��������۲��)�ϼƣ�" & zlStr.FormatEx(Cur���ʲ��, mintMoneyDigit, , True)
    Else
        lblDifference.Caption = "��ۺϼƣ�" & zlStr.FormatEx(Cur���ʲ��, mintMoneyDigit, , True)
        lblSalePrice.Caption = "�ۼ۽��ϼƣ�" & zlStr.FormatEx(Cur���ʽ��, mintMoneyDigit, , True)
    End If
        
End Sub
Private Sub ��ʾ�����()
    Dim RecTmp As New ADODB.Recordset
    Dim Dbl���� As Double
    Dim str��λ As String
    Dim intID As Long
    Dim strUnit As String
    Dim strQuantity As String
    Dim bln��ʾ���ο�� As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    If mint�༭״̬ = 6 Then
        bln��ʾ���ο�� = True
    End If
    
    If mshBill.TextMatrix(mshBill.Row, mconIntColҩ��) = "" Then
        staThis.Panels(2).Text = ""
        Exit Sub
    End If
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
    intID = mshBill.TextMatrix(mshBill.Row, 0)
 
    If RecTmp.State = 1 Then RecTmp.Close
    
    Select Case mintUnit
        Case mconint�ۼ۵�λ
            strUnit = "C.���㵥λ"
            strQuantity = "��������"
        Case mconint���ﵥλ
            strUnit = "B.���ﵥλ"
            strQuantity = "��������/�����װ"
        Case mconintסԺ��λ
            strUnit = "B.סԺ��λ"
            strQuantity = "��������/סԺ��װ"
        Case mconintҩ�ⵥλ
            strUnit = "B.ҩ�ⵥλ"
            strQuantity = "��������/ҩ���װ"
    End Select

    gstrSQL = " SELECT B.ҩƷID," & strUnit & " AS ��λ,SUM(" & strQuantity & ") AS ���� " & _
              " FROM ҩƷ��� A,ҩƷ��� B,�շ���ĿĿ¼ C " & _
              " WHERE A.����=1 AND A.��������<>0 AND A.�ⷿID=[1] " & _
              " AND A.ҩƷID=B.ҩƷID AND B.ҩƷID=C.ID AND B.ҩƷID=[2] "
    '����ǳ�����������ˣ��˿⣬����ʾ�����εĿ��
    If bln��ʾ���ο�� = True Then
        gstrSQL = gstrSQL & " AND NVL(A.����,0)=[3] "
    End If
    gstrSQL = gstrSQL & " GROUP BY B.ҩƷID," & strUnit
    Set RecTmp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ʾ�����]", cboStock.ItemData(cboStock.ListIndex), intID, Val(mshBill.TextMatrix(mshBill.Row, mconIntCol����)))
    
    If RecTmp.EOF Then
        staThis.Panels(2).Text = "��" & IIf(bln��ʾ���ο�� = True, "����", "") & "ҩƷ��ǰ�����Ϊ[0]"
        Exit Sub
    End If
    Dbl���� = IIf(IsNull(RecTmp!����), 0, RecTmp!����)
    
    With mshBill
        strSQL = ""
        If .TextMatrix(.Row, mconIntCol�����־) = "��" And mint�༭״̬ = 6 And gtype_UserSysParms.P173_������Ǹ������ܽ��и������ = 1 Then
            strSQL = "���á���Ǻ���ܸ�������������Ѿ���ǵ�ҩƷ���ܳ�����"
        End If
    End With
    
    staThis.Panels(2).Text = "��" & IIf(bln��ʾ���ο�� = True, "����", "") & "ҩƷ��ǰ�����Ϊ[" & FormatEx(Dbl����, mintNumberDigit) & "]" & RecTmp!��λ & "  " & strSQL
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt�Ӽ���_GotFocus()
    Txt�Ӽ���.SelStart = 0
    Txt�Ӽ���.SelLength = Len(Txt�Ӽ���)
End Sub

Private Sub Txt�Ӽ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cmdYes_Click
End Sub

Private Sub Txt�Ӽ���_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then KeyAscii = 0
End Sub

Private Sub Txt�Ӽ���_LostFocus()
    Call PicInput_LostFocus
End Sub


Private Sub txtժҪ_Change()
    mblnChange = True
End Sub

Private Sub txtժҪ_GotFocus()
    OS.OpenIme True
    With txtժҪ
        .SelStart = 0
        .SelLength = Len(txtժҪ.Text)
    End With
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txtժҪ_LostFocus()
    OS.OpenIme
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


'ȡָ�������۶��۵�λ������ֵ��ȱʡΪ0-���ۼ۵�λ���ۣ���ѡΪ1-��ҩ�ⵥλ���ۣ�
Private Function GetUnit() As Integer
   GetUnit = gtype_UserSysParms.P29_ָ�������۶��۵�λ
    
End Function

'ȡʱ��ҩƷ���ʱ���Ƿ��������Ӽ���
Private Function Getʱ��ҩƷֱ��ȷ���ۼ�() As Boolean
    Getʱ��ҩƷֱ��ȷ���ۼ� = (gtype_UserSysParms.P76_ʱ��ҩƷֱ��ȷ���ۼ� = 1)
    mintʱ������ۼۼӳɷ�ʽ = gtype_UserSysParms.P126_ʱ��ҩƷ�ۼۼӳɷ�ʽ
End Function

'ȡʱ��ҩƷ���ʱ���Ƿ��������Ӽ���
Private Function Get�Ӽ���() As Boolean
    Get�Ӽ��� = (gtype_UserSysParms.P54_ʱ��ҩƷ�ԼӼ������ = 1)
End Function

'ȡ���ݿ������ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select ���� from ҩƷ�շ���¼ where rownum<1 "
    Set rsBatchNolen = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-ȡ���ų���")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ���ɱ���()
    Dim dbl�ɱ��� As Double, dbl���ۼ� As Double
    '����ɱ��۱����ۼۻ��ߣ���ʾ�û�
    With mshBill
        If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
        dbl�ɱ��� = Format(Val(.TextMatrix(.Row, mconIntCol�ɱ���)), "#####0.00000;-#####0.00000;0;")
        dbl���ۼ� = Format(Val(.TextMatrix(.Row, mconIntCol�ۼ�)), "#####0.00000;-#####0.00000;0;")
    End With
    If dbl�ɱ��� > dbl���ۼ� Then
        MsgBox "���ѣ���ҩƷ�ĳɱ��۱����ۼۻ��ߣ�", vbInformation, gstrSysName
    End If
End Sub

Private Function CopyCard() As String
    Dim intRow As Integer, intUpdate As Integer
    Dim sinԭ���� As Double, sin������ As Double
    Dim dbl�ɹ��� As Double, dbl�ɹ���� As Double, dbl��� As Double, dbl���۽�� As Double, dbl���� As Double
    Dim strNo As Variant
    Dim dbl�ۼ� As Double
    On Error GoTo ErrHand
    
    strNo = Sys.GetNextNo(21, Me.cboStock.ItemData(Me.cboStock.ListIndex))
    If IsNull(strNo) Then Exit Function
    intUpdate = 0
    CopyCard = ""
    
    '���Ʋ����µ���
    gstrSQL = "zl_billcopy(1,'" & txtNO.Tag & "','" & strNo & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
    
'    �޸Ĳɹ��ۡ��ɹ�����ۣ�Ҫ���ǵ�������˳����ĵ��ݣ���ʱ��Ҫ�޸Ĳɹ��ۡ��ɹ�����ۣ�
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                dbl�ɹ��� = Val(.TextMatrix(intRow, mconIntCol�ɱ���))
                dbl�ɹ���� = Val(.TextMatrix(intRow, mconIntCol�ɱ����))
                dbl��� = Val(.TextMatrix(intRow, mconintCol���))
                dbl���۽�� = Val(.TextMatrix(intRow, mconIntCol�ۼ۽��))
                dbl���� = Val(.TextMatrix(intRow, mconIntCol����))
                dbl�ۼ� = Val(.TextMatrix(intRow, mconIntCol�ۼ�))
                Call Get����(txtNO.Tag, Val(.TextMatrix(intRow, mconIntCol���)), sinԭ����)
                If Get����(strNo, Val(.TextMatrix(intRow, mconIntCol���)), sin������) Then
                    If Abs(sin������) > 0 Then
                        '��������
                        dbl�ɹ��� = Round(dbl�ɹ��� / Val(.TextMatrix(intRow, mconIntCol����ϵ��)), gtype_UserDrugDigits.Digit_�ɱ���)
                        dbl�ۼ� = Round(dbl�ۼ� / Val(.TextMatrix(intRow, mconIntCol����ϵ��)), gtype_UserDrugDigits.Digit_���ۼ�)
                        dbl�ɹ���� = Val(IIf(mbln�˻�, -1, 1)) * dbl�ɹ����
                        dbl��� = Val(IIf(mbln�˻�, -1, 1)) * dbl���
                        dbl���۽�� = Val(IIf(mbln�˻�, -1, 1)) * dbl���۽��
                        
                        '����ҩƷ�շ���¼
                        gstrSQL = "zl_Bill_������Ϣ(1,'" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol���)) & ",'�ɱ���','" & dbl�ɹ��� & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        gstrSQL = "zl_Bill_������Ϣ(1,'" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol���)) & ",'�ɱ����','" & dbl�ɹ���� & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        gstrSQL = "zl_Bill_������Ϣ(1,'" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol���)) & ",'���','" & dbl��� & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        gstrSQL = "zl_Bill_������Ϣ(1,'" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol���)) & ",'���۽��','" & dbl���۽�� & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        gstrSQL = "zl_Bill_������Ϣ(1,'" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol���)) & ",'����','" & dbl���� & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        gstrSQL = "zl_Bill_������Ϣ(1,'" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol���)) & ",'���ۼ�','" & dbl�ۼ� & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        '����Ӧ����¼
                        gstrSQL = "zl_Bill_����Ӧ����¼('" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol���)) & ",'�ɹ���','" & dbl�ɹ��� & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        gstrSQL = "zl_Bill_����Ӧ����¼('" & strNo & "'," & Val(.TextMatrix(intRow, mconIntCol���)) & ",'�ɹ����','" & dbl�ɹ���� & "')"
                        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)

                        intUpdate = intUpdate + 1
                    End If
                End If
            End If
        Next
    End With

    If intUpdate = 0 Then
        MsgBox "�޷���ɲ�����ˣ���Ϊ�õ����ѱ�ȫ��������", vbInformation, gstrSysName
        Exit Function
    End If
    
    CopyCard = strNo
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function Get����(ByVal strNo As String, ByVal int��� As Integer, sin���� As Double) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(ʵ������,0) ���� From ҩƷ�շ���¼" & _
              " Where ����=1 And NO=[1] And ���=[2] ANd (��¼״̬=1 Or Mod(��¼״̬,3)=0)"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ����]", strNo, int���)
    If rsTemp.EOF Then Exit Function
    sin���� = rsTemp!����
    Get���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckProvider() As Boolean
    Dim lngRow As Long
    Dim strҩƷ As String
    Dim str�б�ҩƷ As String
    Dim rsTemp As New ADODB.Recordset
    '��鹩Ӧ���Ƿ����б�ҩƷ���б굥λ
    strҩƷ = ""
    
    On Error GoTo errHandle
    With mshBill
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, 0)) <> 0 Then
                strҩƷ = strҩƷ & "," & Val(.TextMatrix(lngRow, 0))
            End If
        Next
        If strҩƷ <> "" Then strҩƷ = Mid(strҩƷ, 2)
    End With
    
    '���б�ҩƷ����ȥ����ͬһ���б굥λ���б�ҩƷ��������޼�¼����˵����ȷ�����򰴼�¼�е�ҩƷID��ʾ���ǺϷ����б굥λ
    gstrSQL = " Select a.ҩƷID From ҩƷ��� a " & _
              " Where a.ҩƷID in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) And Nvl(a.�б�ҩƷ,0)=1" & _
              " Minus" & _
              " Select A.ҩƷID From " & _
              "     (Select a.ҩƷID From ҩƷ��� a " & _
              "     Where a.ҩƷID in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) And Nvl(a.�б�ҩƷ,0)=1) A,ҩƷ�б굥λ B" & _
              " Where A.ҩƷID=B.ҩƷID And B.��λID=[1] And (B.����ʱ�� is null or B.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    gstrSQL = " Select '['||A.����||']'||A.���� ҩƷ���� " & _
              " From " & _
              "     (Select A.ҩƷID,C.����,Nvl(B.����,C.����) ����" & _
              "     From (" & gstrSQL & ") A,�շ���Ŀ���� B,�շ���ĿĿ¼ C" & _
              "     Where A.ҩƷID=B.�շ�ϸĿID(+) and A.ҩƷID=C.ID" & _
              "     and B.����(+)=3 and B.����(+)=1) A"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�ж��Ƿ����б굥λ���ɹ�]", Val(txtProvider.Tag), strҩƷ)
    
    With rsTemp
        strҩƷ = ""
        Do While Not .EOF
            strҩƷ = strҩƷ & "��" & rsTemp!ҩƷ����
            .MoveNext
        Loop
        If strҩƷ <> "" Then strҩƷ = Mid(strҩƷ, 2)
    End With
    
    If strҩƷ <> "" Then
        If mbln�б�ҩƷ��ѡ����б굥λ��� = True Then
            If MsgBox("�ù�ҩ��λ���������б�ҩƷ���б굥λ���Ƿ������" & vbCrLf & strҩƷ, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            MsgBox "�ù�ҩ��λ���������б�ҩƷ���б굥λ��" & vbCrLf & strҩƷ, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CheckProvider = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ʱ��ҩƷ���ۼ�(ByVal lngҩƷID As Long, ByVal sin�ɹ��� As Double, ByVal sin�ӳ��� As Double, ByVal sin�ۼ� As Double, Optional ByVal lngLastRow As Long = -1) As Double
    Dim sin���ۼ� As Double, sinָ�����ۼ� As Double, sin��������� As Double
    Dim sinTempָ�����ۼ� As Double
    Dim rsTemp As New ADODB.Recordset
    Dim sin������� As Double
    'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
    '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
    '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)+�ۼ۲���
    If lngLastRow = -1 Then lngLastRow = mshBill.Row
    On Error GoTo errHandle
    gstrSQL = "Select ָ�����ۼ�,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", lngҩƷID)
    
    sinָ�����ۼ� = rsTemp!ָ�����ۼ�
    sinTempָ�����ۼ� = rsTemp!ָ�����ۼ� * Val(mshBill.TextMatrix(lngLastRow, mconIntCol����ϵ��))
    sin��������� = rsTemp!���������
    
    ʱ��ҩƷ���ۼ� = 0
    If sin��������� = 100 Then
        ʱ��ҩƷ���ۼ� = sin�ۼ�
        Exit Function
    End If
    
    If (mint�༭״̬ = 8 Or mbln�˻�) Then
        '������˻����򰴳���ķ�ʽ�����ۼ�
        gstrSQL = " Select Nvl(ʵ������,0) ʵ������,Nvl(ʵ�ʽ��,0) ʵ�ʽ�� From ҩƷ��� " & _
                " Where ����=1 And ҩƷID=[2] And �ⷿID=[1] And Nvl(����,0)=[3] "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�������ۼ�]", cboStock.ItemData(cboStock.ListIndex), lngҩƷID, Val(mshBill.TextMatrix(mshBill.Row, mconIntCol����)))
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "ҩƷ������ݴ���δ�ҵ�ָ��ҩƷ�Ŀ���¼����", vbInformation, gstrSysName
            ʱ��ҩƷ���ۼ� = sin�ۼ�
            Exit Function
        End If
        '�϶���������û�������Ļ����޷�����˴�
        sin������� = rsTemp!ʵ�ʽ�� / rsTemp!ʵ������ * Val(mshBill.TextMatrix(lngLastRow, mconIntCol����ϵ��))
    Else
        sin���ۼ� = sin�ɹ��� * (1 + sin�ӳ���)
        If sin���ۼ� / Val(mshBill.TextMatrix(lngLastRow, mconIntCol����ϵ��)) >= sinָ�����ۼ� Then
            ʱ��ҩƷ���ۼ� = sin�ۼ�
            Exit Function
        End If
        sinָ�����ۼ� = sinָ�����ۼ� * Val(mshBill.TextMatrix(lngLastRow, mconIntCol����ϵ��))
        sin������� = (sinָ�����ۼ� - sin���ۼ�) * (1 - sin��������� / 100)
    End If
    
    ʱ��ҩƷ���ۼ� = IIf(sin������� + sin�ۼ� > sinTempָ�����ۼ�, sinTempָ�����ۼ�, sin������� + sin�ۼ�)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ����ӳ���(ByVal lngҩƷID As Long, ByVal sin���ۼ� As Double, ByVal sin�ɱ��� As Double) As Double
    Dim sinָ�����ۼ� As Double, sin��������� As Double
    Dim rsTemp As New ADODB.Recordset
    '�������ۼ۷���ɱ���,����ʱ��ҩƷ��ʽ�ı仯,����ԭ������ӳ��ʵĹ�ʽ��Ч,�����¼���
    'ԭ��ʽ:(���ۼ�/�ɱ���-1)*100
    '�ֹ�ʽ������:�������ۼ��ǰ��ӳ����������,�ټ������������ǲ��ֽ��,���ʵ�ʰ��ӳ�����������ۼ�=ָ�����ۼ�-(ָ�����ۼ�-���ۼ�)/���������
    '������ԭ��ʽ���ʵ�ʵļӳ���
    ����ӳ��� = 0.15
    
    On Error GoTo errHandle
    gstrSQL = " Select A.ָ�����ۼ�,Nvl(A.���������,100) ���������,Nvl(B.�Ƿ���,0) ʱ�� " & _
          " From ҩƷ��� A,�շ���ĿĿ¼ B " & _
          " Where A.ҩƷID=B.ID AND A.ҩƷID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", lngҩƷID)
    
    sinָ�����ۼ� = rsTemp!ָ�����ۼ�
    sin��������� = rsTemp!���������
    If rsTemp!ʱ�� = 0 Then Exit Function
    
    'ָ�����ۼ�-(ָ�����ۼ�-���ۼ�)/���������
    sinָ�����ۼ� = sinָ�����ۼ� * Val(mshBill.TextMatrix(mshBill.Row, mconIntCol����ϵ��))
    If sin��������� <> 100 And sin��������� > 0 Then
        sin���ۼ� = sinָ�����ۼ� - (sinָ�����ۼ� - sin���ۼ�) / sin��������� * 100
    Else
        sin���ۼ� = sinָ�����ۼ� - (sinָ�����ۼ� - sin���ۼ�)
    End If
    If sin�ɱ��� = 0 Then
        
        ����ӳ��� = (sin���ۼ� / IIf(sin�ɱ��� = 0, 1, sin�ɱ���)) * 100
    Else
        ����ӳ��� = (Val(sin���ۼ�) / Val(sin�ɱ���) - 1) * 100
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function У�����ۼ�(ByVal sin���ۼ� As Double, Optional ByVal lngLastRow As Long = -1) As Double
    '���ܣ��õ�����ǰ��λϵ�����������ָ�����ۼۣ����ʱ��ҩƷ������������ۼ۴���ָ�����ۼۣ���ָ�����ۼ�Ϊ׼
    Dim sinָ�����ۼ� As Double
    Dim rsTemp As New ADODB.Recordset
       
    On Error GoTo errHandle
    If lngLastRow = -1 Then lngLastRow = mshBill.Row
    
    gstrSQL = " Select ָ�����ۼ�,Nvl(���������,100) ��������� " & _
              " From ҩƷ��� Where ҩƷID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", Val(mshBill.TextMatrix(lngLastRow, 0)))
    
    sinָ�����ۼ� = rsTemp!ָ�����ۼ�
    sinָ�����ۼ� = sinָ�����ۼ� * Val(mshBill.TextMatrix(lngLastRow, mconIntCol����ϵ��))
    
    У�����ۼ� = IIf(sin���ۼ� > sinָ�����ۼ�, sinָ�����ۼ�, sin���ۼ�)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetColumnByUserDefine()
    Dim intColumns As Integer
    Dim strColumn_Selected As String
    Dim strColumn_All As String
    Dim arrColumn_All, arrColumn_Selected, arrColumn_UnSelected, arr����, arr��������
    Dim intCol As Integer, intCols As Integer
    Dim strAllCol As String
    Dim strChange As String
    Dim strOldColName As String, strNewColName As String
    On Error GoTo ErrHand
    
    strColumn_Selected = zlDataBase.GetPara("ѡ����", glngSys, ģ���.�⹺���)
    mstrColumn_UnSelected = zlDataBase.GetPara("������", glngSys, ģ���.�⹺���)
    strColumn_All = "ҩ��,2|ҩƷ��Դ,4|����ҩ��,5|ҩ�ۼ���,7|���,8|������,13|ԭ����,14|��λ,15|����,16|��������,17|Ч��,18|����,19|ָ��������,22|�ɹ���,23|����,24|" & _
                    "�ɱ���,25|�ɱ����,26|�ӳ���,27|�ۼ�,28|�ۼ۽��,29|���,30|���ۼ�,31|���۵�λ,32|���۽��,33|���۲��,34|��׼�ĺ�, 35|���,36|" & _
                    "��Ʒ�ϸ�֤,37|�������,38|�������,39|���ս���,40|��Ʊ��,41|��Ʊ����,42|��Ʊ����,43|��Ʊ���,44"
    If strColumn_Selected <> "" Then
        '�����ϰ汾���������Ʊ仯����ʽ��������,������|������,������...
        strChange = "����,������|�����,�ɱ���|������,�ɱ����"
        
        For intCol = 0 To UBound(Split(strChange, "|"))
            strOldColName = Split(Split(strChange, "|")(intCol), ",")(0)
            strNewColName = Split(Split(strChange, "|")(intCol), ",")(1)
            
            If InStr(1, "|" & strColumn_Selected & "|", "|" & strOldColName & "|") <> 0 Then
                strColumn_Selected = Replace("|" & strColumn_Selected & "|", "|" & strOldColName & "|", "|" & strNewColName & "|")
                strColumn_Selected = Left(strColumn_Selected, Len(strColumn_Selected) - 1)
                strColumn_Selected = Mid(strColumn_Selected, 2)
            End If
            
            If InStr(1, "|" & mstrColumn_UnSelected & "|", "|" & strOldColName & "|") <> 0 Then
                mstrColumn_UnSelected = Replace("|" & mstrColumn_UnSelected & "|", "|" & strOldColName & "|", "|" & strNewColName & "|")
                mstrColumn_UnSelected = Left(mstrColumn_UnSelected, Len(mstrColumn_UnSelected) - 1)
                mstrColumn_UnSelected = Mid(mstrColumn_UnSelected, 2)
            End If
        Next
        
        If mstrColumn_UnSelected <> "" Then
            strAllCol = strColumn_Selected & "|" & mstrColumn_UnSelected
        Else
            strAllCol = strColumn_Selected
        End If
        arr���� = Split(strColumn_All, "|")
        arr�������� = Split(strAllCol, "|")
        If UBound(arr����) <> UBound(arr��������) Or InStr(1, "|" & strColumn_Selected & "|", "|������|") = 0 Or InStr(1, "|" & mstrColumn_UnSelected & "|", "|������|") <> 0 Then
            strColumn_Selected = "ҩ��|ҩƷ��Դ|����ҩ��|ҩ�ۼ���|���|������|ԭ����|����|��������|Ч��|��λ|����|ָ��������|�ɹ���|����|�ɱ���|�ɱ����|�ӳ���|�ۼ�|�ۼ۽��|���|��׼�ĺ�|���|��Ʒ�ϸ�֤|�������|�������|���ս���|��Ʊ��|��Ʊ����|��Ʊ����|��Ʊ���"
            mstrColumn_UnSelected = "���ۼ�|���۵�λ|���۽��|���۲��"
            zlDataBase.SetPara "ѡ����", strColumn_Selected, glngSys, ģ���.�⹺���
            zlDataBase.SetPara "������", mstrColumn_UnSelected, glngSys, ģ���.�⹺���
        End If
    Else
        strColumn_Selected = "ҩ��|ҩƷ��Դ|����ҩ��|ҩ�ۼ���|���|������|ԭ����|����|��������|Ч��|��λ|����|ָ��������|�ɹ���|����|�ɱ���|�ɱ����|�ӳ���|�ۼ�|�ۼ۽��|���|��׼�ĺ�|���|��Ʒ�ϸ�֤|�������|�������|���ս���|��Ʊ��|��Ʊ����|��Ʊ����|��Ʊ���"
        mstrColumn_UnSelected = "���ۼ�|���۵�λ|���۽��|���۲��"
        zlDataBase.SetPara "ѡ����", strColumn_Selected, glngSys, ģ���.�⹺���
        zlDataBase.SetPara "������", mstrColumn_UnSelected, glngSys, ģ���.�⹺���
    End If
    
    '��װ��ȱʡ����
    mconIntCol�к� = 1
    mconIntColҩ�� = 2
    mconIntCol��Ʒ�� = 3
    mconIntCol��Դ = 4
    mconIntCol����ҩ�� = 5
    mconIntCol��� = 6
    mconIntColҩ�ۼ��� = 7
    mconIntCol��� = 8
    mconIntColԭ������ = 9
    mconIntColԭ���� = 10
    mconIntCol����ϵ�� = 11
    mconintcol���� = 12
    mconIntCol���� = 13
    mconIntColԭ���� = 14
    mconIntCol��λ = 15
    mconIntCol���� = 16
    mconIntCol�������� = 17
    mconIntColЧ�� = 18
    mconIntCol���� = 19
    mconIntCol�������� = 20
    mconIntCol���� = 21
    mconIntColָ�������� = 22
    mconIntCol�ɹ��� = 23
    mconIntCol���� = 24
    mconIntCol�ɱ��� = 25
    mconIntCol�ɱ���� = 26
    mconIntcol�ӳ��� = 27
    mconIntCol�ۼ� = 28
    mconIntCol�ۼ۽�� = 29
    mconintCol��� = 30
    mconintCol���ۼ� = 31
    mconintCol���۵�λ = 32
    mconintCol���۽�� = 33
    mconintCol���۲�� = 34
    mconIntCol��׼�ĺ� = 35
    mconIntCol��� = 36
    mconintcol��Ʒ�ϸ�֤ = 37
    mconintcol������� = 38
    mconintcol������� = 39
    mconIntCol���ս��� = 40
    mconintcol��Ʊ�� = 41
    mconintcol��Ʊ���� = 42
    mconIntCol��Ʊ���� = 43
    mconintcol��Ʊ��� = 44
    mconIntCol�������� = 45
    mconIntCol�Ƿ����� = 46
    mconIntColҩƷ��������� = 47
    mconIntColҩƷ���� = 48
    mconIntColҩƷ���� = 49
    mconIntCol�����־ = 50
    mconIntCol�ƻ�id = 51
    
    mintLastCol = 51

    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = IIf(mshBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, mshBill.ColWidth(mconIntCol��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = 0
    End If
    
    '�����û����õ�����˳��
    arrColumn_All = Split(strColumn_All, "|")
    arrColumn_Selected = Split(strColumn_Selected, "|")
    intCols = UBound(arrColumn_Selected)
    For intCol = 0 To intCols
        Call SetColumnValue(arrColumn_Selected(intCol), Split(arrColumn_All(intCol), ",")(1))
    Next
    
    '��δѡ����е��п�����Ϊ�㣬��������Ϊ5��������ѡ��
    If mstrColumn_UnSelected = "" Then Exit Sub
    intCol = intCols + 1
    intColumns = 0
    arrColumn_UnSelected = Split(mstrColumn_UnSelected, "|")
    intCols = UBound(arrColumn_All)
    For intCol = intCol To intCols
        If UBound(arrColumn_UnSelected) >= intColumns Then
            Call SetColumnValue(arrColumn_UnSelected(intColumns), Split(arrColumn_All(intCol), ",")(1), False)
            intColumns = intColumns + 1
        Else
            Call SetColumnValue(Split(arrColumn_All(intCol), ",")(0), Split(arrColumn_All(intCol), ",")(1), False)
        End If
    Next
    Exit Sub
ErrHand:
    MsgBox "�ָ�������ʱ�������������½��������ã�", vbInformation, gstrSysName
End Sub

Private Sub SetColumnValue(ByVal str���� As String, ByVal intValue As Integer, Optional ByVal blnShow As Boolean = True)
    Select Case str����
    Case "�к�"
        mconIntCol�к� = intValue
    Case "ҩ��"
        mconIntColҩ�� = intValue
    Case "ҩƷ��Դ"
        mconIntCol��Դ = intValue
    Case "����ҩ��"
        mconIntCol����ҩ�� = intValue
    Case "���"
        mconIntCol��� = intValue
    Case "���"
        mconIntCol��� = intValue
    Case "ҩ�ۼ���"
        mconIntColҩ�ۼ��� = intValue
    Case "ԭ������"
        mconIntColԭ������ = intValue
    Case "ԭ����"
        mconIntColԭ���� = intValue
    Case "����ϵ��"
        mconIntCol����ϵ�� = intValue
    Case "����"
        mconintcol���� = intValue
    Case "������"
        mconIntCol���� = intValue
    Case "ԭ����"
        mconIntColԭ���� = intValue
    Case "��λ"
        mconIntCol��λ = intValue
    Case "����"
        mconIntCol���� = intValue
    Case "��������"
        mconIntCol�������� = intValue
    Case "Ч��"
        mconIntColЧ�� = intValue
    Case "����"
        mconIntCol���� = intValue
    Case "��������"
        mconIntCol�������� = intValue
    Case "ָ��������"
        mconIntColָ�������� = intValue
    Case "����"
        mconIntCol���� = intValue
    Case "�ɱ���"
        mconIntCol�ɱ��� = intValue
    Case "�ɱ����"
        mconIntCol�ɱ���� = intValue
    Case "�ۼ�"
        mconIntCol�ۼ� = intValue
    Case "�ۼ۽��"
        mconIntCol�ۼ۽�� = intValue
    Case "���"
        mconintCol��� = intValue
    Case "���ۼ�"
        mconintCol���ۼ� = intValue
    Case "���۵�λ"
        mconintCol���۵�λ = intValue
    Case "���۽��"
        mconintCol���۽�� = intValue
    Case "���۲��"
        mconintCol���۲�� = intValue
    Case "��׼�ĺ�"
        mconIntCol��׼�ĺ� = intValue
    Case "���"
        mconIntCol��� = intValue
    Case "��Ʒ�ϸ�֤"
        mconintcol��Ʒ�ϸ�֤ = intValue
    Case "�������"
        mconintcol������� = intValue
    Case "�������"
        mconintcol������� = intValue
    Case "��Ʊ��"
        mconintcol��Ʊ�� = intValue
    Case "��Ʊ����"
        mconintcol��Ʊ���� = intValue
    Case "��Ʊ����"
        mconIntCol��Ʊ���� = intValue
    Case "��Ʊ���"
        mconintcol��Ʊ��� = intValue
    Case "�ɹ���"
        mconIntCol�ɹ��� = intValue
    Case "�ӳ���"
        mconIntcol�ӳ��� = intValue
    Case "���ս���"
        mconIntCol���ս��� = intValue
    End Select
    
    If Not blnShow Then
        mshBill.ColWidth(intValue) = 0
        mshBill.ColData(intValue) = 5
    Else
        mintLastCol = intValue
    End If
End Sub

Private Function Check�Ƿ���ڸ�����() As Boolean
    Dim n As Integer
    
    With mshBill
        For n = 1 To .rows - 1
            If Val(.TextMatrix(n, 0)) <> 0 Then
                If Val(.TextMatrix(n, mconIntCol����)) < 0 Then
                    Check�Ƿ���ڸ����� = True
                    Exit Function
                End If
            End If
        Next

    End With
End Function

Private Sub RefreshBill()
    '�����¼۸����µ���������ݣ����ڵ������ʱ
    Dim lngRow As Long, lngRows As Long, lngҩƷID As Long
    Dim Dbl���� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    Dim rsPrice As New ADODB.Recordset
    Dim intPriceDigit As Integer
        
    On Error GoTo errHandle
    intPriceDigit = GetDigit(0, 1, 2, 1)
        
    gstrSQL = " Select �շ�ϸĿID,nvl(�ּ�,0) �ּ� From �շѼ�Ŀ " & _
            " Where (��ֹ���� Is NULL Or sysdate Between ִ������ And nvl(��ֹ����,to_date('3000-01-01','yyyy-MM-dd')))" & _
            GetPriceClassString("")
    gstrSQL = "Select A.���,A.ҩƷID,B.�ּ� From ҩƷ�շ���¼ A,(" & gstrSQL & ") B,�շ���ĿĿ¼ C" & _
            " Where A.����=1 And A.NO=[1] And A.ҩƷID=B.�շ�ϸĿID And C.ID=B.�շ�ϸĿID And Round(A.���ۼ�," & intPriceDigit & ")<>Round(B.�ּ�," & intPriceDigit & ") And Nvl(C.�Ƿ���,0)=0" & _
            " Union All " & _
            " Select A.���, A.ҩƷid, decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�) �ּ� " & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C, ҩƷ��� D , " & _
            "      (Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 1 and (x.��ֹ���� Is Null Or Sysdate Between x.ִ������ And Nvl(x.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where A.���� = 1 And A.NO = [1] And C.ID = A.ҩƷid And Round(A.���ۼ�, " & intPriceDigit & ") <> Round(decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�), " & intPriceDigit & ") And " & _
            " Nvl(C.�Ƿ���, 0) = 1 And D.ҩƷid = A.ҩƷid And B.���� = 1 And B.�ⷿid = A.�ⷿid And B.ҩƷid = A.ҩƷid And " & _
            " a.ҩƷid = x.ҩƷid(+) And a.�ⷿid = x.�ⷿid(+) And Nvl(a.����, 0) = Nvl(x.����(+), 0) AND " & _
            " Nvl(B.����, 0) = Nvl(A.����, 0) And NVL(b.ʵ������, 0) <> 0 And decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�) > 0 " & _
            " Order by ҩƷid,���"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ��ǰ�۸�]", txtNO.Text)
    
    If rsPrice.EOF Then Exit Sub
    
    lngRows = mshBill.rows - 1
    For lngRow = 1 To lngRows
        lngҩƷID = Val(mshBill.TextMatrix(lngRow, 0))
        If lngҩƷID <> 0 Then
            rsPrice.Filter = "ҩƷID=" & lngҩƷID
            If rsPrice.RecordCount <> 0 Then
                '�Ե�ǰ���¼۸����µ���������ݣ����ۡ����۽���ۣ�
                dbl���ۼ� = rsPrice!�ּ� * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��))
                dbl�ɱ��� = Val(mshBill.TextMatrix(lngRow, mconIntCol�ɱ���))
                Dbl���� = Val(mshBill.TextMatrix(lngRow, mconIntCol����))
                dbl�ɱ���� = dbl�ɱ��� * Dbl����
                dbl���۽�� = dbl���ۼ� * Dbl����
                dbl��� = dbl���۽�� - dbl�ɱ����
                
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ�, intPriceDigit, , True)
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(dbl���۽��, mintMoneyDigit, , True)
                mshBill.TextMatrix(lngRow, mconintCol���) = zlStr.FormatEx(dbl���, mintMoneyDigit, , True)
                mshBill.TextMatrix(lngRow, mconIntcol�ӳ���) = zlStr.FormatEx((Val(mshBill.TextMatrix(lngRow, mconIntCol�ۼ�)) / dbl�ɱ��� - 1) * 100, 2) & "%"
            End If
        End If
    Next
    rsPrice.Filter = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetState()
    '����ĳ����Ԫ���Ƿ���Խ����޸�
    Dim strTemp As String
    Dim i As Integer
        
    With mshBill
        If .TextMatrix(.Row, mconIntCol�����־) = "��" And gtype_UserSysParms.P173_������Ǹ������ܽ��и������ = 1 Then
            If InStr(1, mstrControlItem, "�ɹ���") > 0 Then
                .ColData(mconIntCol�ɹ���) = 5
            Else
                .ColData(mconIntCol�ɹ���) = 4
            End If
            
            If InStr(1, mstrControlItem, "����") > 0 Then
                .ColData(mconIntCol����) = 5
            Else
                .ColData(mconIntCol����) = 4
            End If
            
            If InStr(1, mstrControlItem, "�ɱ���") > 0 Then
                .ColData(mconIntCol�ɱ���) = 5
            Else
                .ColData(mconIntCol�ɱ���) = 4
            End If
            
            If InStr(1, mstrControlItem, "�ɱ����") > 0 Then
                .ColData(mconIntCol�ɱ����) = 5
            Else
                .ColData(mconIntCol�ɱ����) = 4
            End If
            
            If InStr(1, mstrControlItem, "�ۼ�") > 0 Then
                .ColData(mconIntCol�ۼ�) = 5
            Else
                .ColData(mconIntCol�ۼ�) = 4
            End If
        End If
    End With
End Sub

Private Function CheckRedo(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '���ܣ����ظ��ļ�¼���˵��������ع��˺�����ݼ���

    Dim i As Integer
    Dim strTemp As String
    Dim str���� As String
    Dim strҩƷID As String
    Dim str�ظ�ҩ�� As String
    Dim strDub As String
    Dim strSQL As String
    
    rsTemp.MoveFirst
    str���� = ""
    Do While Not rsTemp.EOF
        str���� = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
        If InStr(1, strTemp, rsTemp!ҩƷID & "," & str����) = 0 Then
            strTemp = strTemp & rsTemp!ҩƷID & "," & str���� & "|"
        End If
        rsTemp.MoveNext
    Loop
    
    With mshBill
        For i = 1 To .rows - 1
            If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol����)) > 0 And .TextMatrix(i, 0) <> "" Then
                strҩƷID = strҩƷID & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntColҩ��) & "|"
            End If
        Next
        
        If strҩƷID <> "" Then   'Ϊ��������ƴ��sql
            strDub = ""
            For i = 0 To UBound(Split(strҩƷID, "|")) - 1
                strDub = strDub & "ҩƷid<>" & Split(Split(strҩƷID, "|")(i), ",")(0) & " and "
                If UBound(Split(str�ظ�ҩ��, ",")) <= 2 Then
                    str�ظ�ҩ�� = str�ظ�ҩ�� & Split(Split(strҩƷID, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        
        If str�ظ�ҩ�� <> "" Then
            MsgBox str�ظ�ҩ�� & "�б����Ѿ������ˣ�" & vbCrLf & "����ҩƷ������ӣ�", vbInformation, gstrSysName
            strSQL = strDub
        End If
        rsTemp.Filter = strSQL
        Set CheckRedo = rsTemp
    End With
End Function

Private Sub vsfInputCost_DblClick()
    If vsfInputCost.rows = 1 Then Exit Sub
    With mshBill
        .SetFocus
        .Row = vsfInputCost.Tag
        .Col = mconIntCol�ɹ���
        .TextMatrix(vsfInputCost.Tag, mconIntCol�ɹ���) = vsfInputCost.TextMatrix(vsfInputCost.Row, vsfInputCost.ColIndex("�ɱ���"))
        .TextMatrix(vsfInputCost.Tag, mconIntCol�ɱ���) = zlStr.FormatEx(Val(.TextMatrix(vsfInputCost.Tag, mconIntCol�ɹ���)) * Val(.TextMatrix(vsfInputCost.Tag, mconIntCol����)) / 100, mintCostDigit, , True)
        '���ý��
        If .TextMatrix(vsfInputCost.Tag, mconIntCol����) <> "" Then
            .TextMatrix(vsfInputCost.Tag, mconIntCol�ɱ����) = zlStr.FormatEx(.TextMatrix(vsfInputCost.Tag, mconIntCol����) * Val(.TextMatrix(vsfInputCost.Tag, mconIntCol�ɱ���)), mintMoneyDigit, , True)
            .TextMatrix(vsfInputCost.Tag, mconintcol��Ʊ���) = IIf(Trim(.TextMatrix(vsfInputCost.Tag, mconintcol��Ʊ��)) = "", "", .TextMatrix(vsfInputCost.Tag, mconIntCol�ɱ����))
            .TextMatrix(vsfInputCost.Tag, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(vsfInputCost.Tag, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(vsfInputCost.Tag, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(vsfInputCost.Tag, mconIntCol�ɱ����) = "", 0, .TextMatrix(vsfInputCost.Tag, mconIntCol�ɱ����)), mintMoneyDigit, , True)
            .TextMatrix(vsfInputCost.Tag, mconintCol���۲��) = zlStr.FormatEx(Val(.TextMatrix(vsfInputCost.Tag, mconintCol���۽��)) - Val(.TextMatrix(vsfInputCost.Tag, mconIntCol�ɱ����)), mintMoneyDigit, , True)
        End If
        
        Call ��ʾ�ϼƽ��
        picInputCost.Visible = False
    End With
End Sub

Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
     
    gstrSQL = "Select t.�ϴβ��� as ������, t.ԭ���� as ԭ���� From ҩƷ��� T Where Rownum < 1"
    Call zlDataBase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    mlng�����̳��� = rsTmp.Fields("������").DefinedSize
    mlngԭ���س��� = rsTmp.Fields("ԭ����").DefinedSize
   
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
