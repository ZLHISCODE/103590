VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmPurchaseCard 
   Caption         =   "�����⹺��ⵥ"
   ClientHeight    =   6990
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11430
   Icon            =   "frmPurchaseCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdALLDel 
      Caption         =   "ȫ��(&D)"
      Height          =   350
      Left            =   4080
      TabIndex        =   63
      ToolTipText     =   "��������еķ�Ʊ�������"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdBulkCopy 
      Caption         =   "��������(&C)"
      Height          =   350
      Left            =   6000
      TabIndex        =   62
      ToolTipText     =   "���Ƶ�ǰ�з�Ʊ��ϢӦ���������޷�Ʊ��Ϣ��"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "��������(&C)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6000
      TabIndex        =   60
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtCopy 
      Enabled         =   0   'False
      Height          =   270
      Left            =   7320
      MaxLength       =   100
      TabIndex        =   59
      Text            =   "1"
      Top             =   5805
      Width           =   600
   End
   Begin VB.CommandButton cmdExtractData 
      Caption         =   "��ȡ����(&E)"
      Height          =   350
      Left            =   1440
      TabIndex        =   58
      Top             =   5760
      Width           =   1215
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
      TabIndex        =   42
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
         TabIndex        =   45
         Text            =   "15.0000"
         Top             =   690
         Width           =   1725
      End
      Begin VB.CommandButton CmdNO 
         Caption         =   "ȡ��"
         Height          =   345
         Left            =   1800
         TabIndex        =   47
         Top             =   1140
         Width           =   855
      End
      Begin VB.CommandButton CmdYes 
         Caption         =   "ȷ��"
         Height          =   345
         Left            =   810
         TabIndex        =   46
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "    ������ӳ��ʣ����ۼ۵ļ��㹫ʽ�����ۼ�=�ɱ���*(1+�ӳ���%)"
         ForeColor       =   &H00400000&
         Height          =   585
         Left            =   0
         TabIndex        =   43
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
         TabIndex        =   44
         Top             =   750
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   10095
      TabIndex        =   41
      Top             =   6225
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   8775
      TabIndex        =   16
      Top             =   6225
      Visible         =   0   'False
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
      Height          =   2175
      Left            =   2520
      TabIndex        =   40
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
      Left            =   4425
      TabIndex        =   23
      Top             =   5775
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2760
      TabIndex        =   22
      Top             =   5745
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   225
      TabIndex        =   21
      Top             =   5745
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8760
      TabIndex        =   15
      Top             =   5745
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   10095
      TabIndex        =   34
      Top             =   5745
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   5565
      Left            =   0
      ScaleHeight     =   5505
      ScaleWidth      =   11655
      TabIndex        =   24
      Top             =   0
      Width           =   11715
      Begin VB.PictureBox picCostly 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6120
         ScaleHeight     =   495
         ScaleWidth      =   5175
         TabIndex        =   55
         Top             =   3120
         Width           =   5175
         Begin VB.TextBox txtTypeVar 
            Height          =   270
            Left            =   3720
            TabIndex        =   57
            Top             =   80
            Width           =   2000
         End
         Begin VB.Label lblType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Caption         =   "�������͡�"
            Height          =   180
            Left            =   2400
            TabIndex        =   56
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label lblCostly 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "��ֵ������Ϣ(&1)"
            Height          =   180
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   1350
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
         Height          =   1815
         Left            =   4800
         TabIndex        =   35
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
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   1000
         TabIndex        =   6
         Top             =   4080
         Width           =   10410
      End
      Begin VB.CommandButton cmdProvider 
         Caption         =   "��"
         Height          =   300
         Left            =   11160
         TabIndex        =   4
         Top             =   660
         Width           =   300
      End
      Begin VB.TextBox txtProvider 
         Height          =   300
         Left            =   8280
         TabIndex        =   3
         Top             =   660
         Width           =   2895
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   660
         Width           =   1515
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9960
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   158
         Width           =   1425
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   240
         TabIndex        =   5
         Top             =   1005
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
      Begin VSFlex8Ctl.VSFlexGrid vsfCostlyInfo 
         Height          =   615
         Left            =   1200
         TabIndex        =   54
         Top             =   5200
         Visible         =   0   'False
         Width           =   3375
         _cx             =   5953
         _cy             =   1085
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
      End
      Begin VB.Label lbl�˲����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�˲�����"
         Height          =   180
         Left            =   4035
         TabIndex        =   51
         Top             =   4920
         Width           =   720
      End
      Begin VB.Label lbl�˲��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  �˲���"
         Height          =   180
         Left            =   4080
         TabIndex        =   50
         Top             =   4515
         Width           =   720
      End
      Begin VB.Label txt�˲����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4845
         TabIndex        =   49
         Top             =   4860
         Width           =   1890
      End
      Begin VB.Label txt�˲��� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4845
         TabIndex        =   48
         Top             =   4455
         Width           =   1890
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   38
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   2040
         TabIndex        =   37
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɹ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   36
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9450
         TabIndex        =   32
         Top             =   4500
         Width           =   1890
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9420
         TabIndex        =   31
         Top             =   4905
         Width           =   1890
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   30
         Top             =   4830
         Width           =   1890
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   29
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
         Left            =   9480
         TabIndex        =   18
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "���������⹺��ⵥ"
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
         TabIndex        =   17
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
         TabIndex        =   0
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  ������"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   4485
         Width           =   720
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   135
         TabIndex        =   27
         Top             =   4890
         Width           =   720
      End
      Begin VB.Label lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  �����"
         Height          =   180
         Left            =   8610
         TabIndex        =   26
         Top             =   4560
         Width           =   720
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   8610
         TabIndex        =   25
         Top             =   4965
         Width           =   720
      End
      Begin VB.Label LblProvider 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������λ(&G)"
         Height          =   180
         Left            =   7200
         TabIndex        =   2
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
            Picture         =   "frmPurchaseCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1000
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
            Picture         =   "frmPurchaseCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   39
      Top             =   6630
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPurchaseCard.frx":22EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13811
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmPurchaseCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmPurchaseCard.frx":3080
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
   Begin VB.Frame fraMoveNO 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1200
      TabIndex        =   53
      Top             =   6000
      Width           =   7605
      Begin VB.ComboBox cboType 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   15
         Width           =   885
      End
      Begin VB.CheckBox chkת���ƿ� 
         Caption         =   "������ⵥ�ƿ⵽"
         Height          =   270
         Left            =   90
         TabIndex        =   7
         Top             =   30
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ComboBox cboEnterStock 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "frmPurchaseCard.frx":3582
         Left            =   2160
         List            =   "frmPurchaseCard.frx":358B
         TabIndex        =   9
         Text            =   "cboEnterStock"
         Top             =   15
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "��"
         Height          =   300
         Left            =   4575
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   15
         Width           =   300
      End
      Begin VB.TextBox txtDraw 
         Height          =   300
         Left            =   2190
         TabIndex        =   10
         Top             =   15
         Width           =   2415
      End
      Begin VB.CommandButton cmdDrawPerson 
         Caption         =   "��"
         Height          =   300
         Left            =   7230
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   15
         Width           =   300
      End
      Begin VB.TextBox txtDrawPerson 
         Height          =   300
         Left            =   5790
         TabIndex        =   13
         Top             =   15
         Width           =   1425
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������(&L)"
         Height          =   180
         Left            =   4920
         TabIndex        =   12
         Top             =   75
         Width           =   825
      End
   End
   Begin VB.Label lblCopy 
      AutoSize        =   -1  'True
      Caption         =   "(���9999)"
      Enabled         =   0   'False
      Height          =   180
      Left            =   7920
      TabIndex        =   61
      Top             =   5850
      Width           =   900
   End
   Begin VB.Label lblCode 
      Caption         =   "����"
      Height          =   255
      Left            =   3945
      TabIndex        =   33
      Top             =   5865
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "��������"
      Visible         =   0   'False
      Begin VB.Menu mnuSearch01 
         Caption         =   "����ID"
      End
      Begin VB.Menu mnuSearch02 
         Caption         =   "��������"
      End
      Begin VB.Menu mnuSearch03 
         Caption         =   "סԺ��"
      End
      Begin VB.Menu mnuSearch04 
         Caption         =   "�����"
      End
      Begin VB.Menu mnuSearch05 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "frmPurchaseCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private mrs���ڿ��� As ADODB.Recordset
Private mlng������λID As Long              '��ҩ��λID
Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5���޸ķ�Ʊ��6��������
                                            '7��������ˣ������������µ��ݲ���ˣ��Ѹ���ĵ��ݲ����������ˣ�ͬ����������˺�ĵ��ݲ����������;
                                            '8�����Ŀ��˻�,9-�˲�,10-�޸�ע��֤��
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mstrPrivs As String                 'Ȩ��
Private mintBatchNoLen As Integer           '���ݿ������Ŷ��峤��
Private mint��Ʊ��Len As Integer            '���ݿ��еķ�Ʊ�ų���

Private mbln�޸������� As Boolean               '�����޸�������
Private mbln�Ӽ��� As Boolean                   'ʱ�������Ƿ��������Ӽ���
Private mdbl�Ӽ��� As Double
Private mbln��������    As Boolean              '����ʱ���ݺ��ۼ�1
Private mintUnit  As Integer                    '��ʾ��λ:0-ɢװ��λ,1-��װ��λ
Private mbln�˻� As Boolean
Private mblnFirst As Boolean
Private mbln�ⷿ  As Boolean                    '�ÿⷿ�Ƿ�Ϊ���Ŀ�!
Private mstr������� As String                  '�������ʱ��
Private mstr���ս��� As String                  '������¼Ĭ�����ս���

Public mrsReturn As Recordset
Private mbln��ǿ�ƿ���ָ���۸� As Boolean       '�����Ƿ�ǿ�Ʋ���ָ���۸�
Private mbln�ֶμӳ��� As Boolean               '�Էֶμӳ���Ϊ����
Private mblnʱ�۹�ǰ���� As Boolean             '�⹺��ʱ��ʱ�۰���ǰ�ӳ����ۼ�
Private mblnʱ������ֱ��ȷ���ۼ� As Boolean     '�⹺���ʱ,ʱ������ֱ��ȷ���ۼۣ�ֱ��ȷ���ۼ۵���˼�ǿ����ֶ�����
Private mblnUpdate As Boolean                   '�Ƿ��µ��ۼ۸��µ��ݣ���Ҫ��������ʱ���۸��µ����
Private mblnCheckPrice As Boolean
Private mbln���б굥λ��� As Boolean           '���Ϊtrue,�б��������Ͽ����ڷ��б굥λ�����
Private mCllBillData As Collection              '��ʹ�õ���������,Ŀǰ��Ҫ���˻������޸�,�Բ���ID & "-" & ����Ϊ����
Private mbln��Ҫ�˲� As Boolean
Private mbln��Ӧ��У�� As Boolean
Private mintCheckType As Integer                '����У�鷽ʽ��0-����飻1�����ѣ�2����ֹ
Private mintProduceDate As Integer              '��������ע��֤Ч�ڼ�飬1-��飬0-�����

Private mrsCostlyInfo As ADODB.Recordset        '��ֵ����
Private mlngLastRow As Long
Private mblnʱ������ȡ�ϴ��ۼ� As Boolean        '�⹺���ʱ��ʱ������ȡ�ϴ��ۼ� true-ȡ�ϴ��ۼ� false-������ʽ����
Private mblnCostView As Boolean                 '�鿴�ɱ��� true-����鿴 false-������鿴
Private Const mstrCaption As String = "�����⹺��ⵥ"
Private mintLastCol As Integer                  '�û����������е����ɼ��е��к�
Private mbln��ֵ���� As Boolean                 'true-��ǰ���Ǹ�ֵ���ģ�false-���Ǹ�ֵ����

Private mbln�����������Ų��ؿ��� As Boolean  '�Ƿ�������������Ų����Ƿ�¼��

'���˺�:2007/06/10:����10813
Private mstrTime_Start As String                '���뵥�ݱ༭�ĵ���ʱ�� ,��Ҫ�ж��Ƿ񵥾ݱ����˸��Ĺ�,����༭��,���ܽ������
Private mstrTime_End As String
Private Const mlngModule = 1712

Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

'=========================================================================================
Private mCol�к� As Integer
Private mCol���� As Integer
Private mCol��� As Integer
Private mCol��Ʒ�� As Integer
Private mCol��� As Integer
Private mColԭ���� As Integer
Private mColԭ���� As Integer
Private mCol����ϵ�� As Integer
Private mCol���� As Integer
Private mCol���� As Integer
Private mcol��׼�ĺ� As Integer
Private mCol��λ As Integer
Private mCol���� As Integer
Private mcol�������� As Integer
Private mColЧ�� As Integer
Private mCol���� As Integer
Private mCol�������� As Integer
Private mCol���� As Integer
Private mColָ�������� As Integer
Private mCol�ɹ��� As Integer
Private mCol���� As Integer
Private mcol�ӳ��� As Integer
Private mCol����� As Integer
Private mCol������ As Integer
Private mCol�ۼ� As Integer
Private mCol�ۼ۽�� As Integer
Private mCol��� As Integer
Private mcol���ۼ� As Integer
Private mcol���۵�λ As Integer
Private mcol���۽�� As Integer
Private mcol���۲�� As Integer
Private mCol������� As Integer
Private mCol���ս��� As Integer
Private mCol��Ʊ�� As Integer
Private mcol��Ʊ���� As Integer
Private mCol��Ʊ���� As Integer
Private mCol��Ʊ��� As Integer
Private mcolһ���Բ��� As Integer
Private mcol���Ч�� As Integer
Private mcol�������   As Integer
Private mcol���ʧЧ�� As Integer
Private mcolע��֤�� As Integer
Private mcol��Ʒ���� As Integer
Private mcol������� As Integer
Private mcol�ڲ����� As Integer
Private mcol����ID As Integer
Private mcolע��֤��Ч�� As Integer
Private Const mCols As Integer = 48
'=========================================================================================
Private Function CheckValuePrice(ByVal int���� As Integer, ByVal strNo As String) As Boolean
    '����ֵ������������������ⵥ�ļ۸��м۸�䶯ʱ���½���۸񣬽��
    '��������ⵥ�������ں��Ƿ����ͬ���εĵ��ۼ�¼������е��ۼ�¼��������ĵ��ۼ�¼�͵�ǰ��ⵥ�ļ۸���бȽ�
    'ֻ���ʱ�����ĵ��ۼۺͳɱ���
    '���أ�true-���ͨ��,false-�м۸�䶯
    Dim rsData As ADODB.Recordset
    Dim rsprice As ADODB.Recordset
    Dim lng����ID As Long
    Dim lng���� As Long
    Dim str�������� As String
    Dim dblԭ�� As Double
    Dim dbl���ۼ� As Double
    Dim dbl�ֳɱ��� As Double
    Dim strAdjustList As String '��Ҫ�䶯���嵥������id,����,���ۼ�(Ϊ0��ʾ�۸��ޱ仯),�ֳɱ���(Ϊ0��ʾ�۸��ޱ仯)
    Dim lngRow As Long
    Dim lngRows As Long
    Dim dbl���� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl���۽�� As Double
    Dim dbl��� As Double
    Dim blnUpdate As Boolean
    
    gstrSQL = "Select '�ۼ�' As ����, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, a.���ۼ� As ԭ��, a.�������� " & vbNewLine & _
        " From ҩƷ�շ���¼ A, �շ���ĿĿ¼ C" & vbNewLine & _
        " Where a.���� = [1] And a.No = [2] And c.Id = a.ҩƷid And Nvl(c.�Ƿ���, 0) = 1 And a.����id > 0 And Exists" & vbNewLine & _
        " (Select 1" & vbNewLine & _
        "       From ҩƷ�շ���¼ B" & vbNewLine & _
        "       Where a.ҩƷid = b.ҩƷid And a.���� = b.���� And b.���� = 13 And b.������� > a.�������� And b.ժҪ = '���ĵ���')" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select '�ɱ���' As ����, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, a.�ɱ��� As ԭ��, a.�������� " & vbNewLine & _
        " From ҩƷ�շ���¼ A" & vbNewLine & _
        " Where a.���� = [1] And a.No = [2] And a.����id > 0 And Exists" & vbNewLine & _
        " (Select 1" & vbNewLine & _
        "       From ҩƷ�շ���¼ B" & vbNewLine & _
        "       Where a.ҩƷid = b.ҩƷid And a.���� = b.���� And b.���� = 18 And b.������� > a.�������� And b.ժҪ = '�������ϳɱ��۵���') "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", int����, strNo)
        
    If rsData.RecordCount = 0 Then
        CheckValuePrice = True
        Exit Function
    End If
    
    '��鵽�е��ۼ�¼��Ƚϼ۸����ڵ��ۼ�¼�����ж�����ȡ���һ���۸����Ƚ�
    Do While Not rsData.EOF
        lng����ID = rsData!����ID
        lng���� = rsData!����
        str�������� = Format(rsData!��������, "yyyy-mm-dd hh:mm:ss")
        dblԭ�� = rsData!ԭ��
        
        dbl���ۼ� = 0
        dbl�ֳɱ��� = 0
        
        If rsData!���� = "�ۼ�" Then
            gstrSQL = "Select ���ۼ� As �ּ� " & _
                " From ҩƷ�շ���¼ " & _
                " Where ID = (Select Max(ID) " & _
                " From ҩƷ�շ���¼ B " & _
                " Where b.ҩƷid = [1] And b.���� = [2] And b.���� = 13 And b.������� > [3] And b.ժҪ = '���ĵ���') "
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", lng����ID, lng����, CDate(str��������))
            
            If rsprice.RecordCount > 0 Then
                If Round(rsprice!�ּ�, 2) <> Round(dblԭ��, 2) Then
                    dbl���ۼ� = rsprice!�ּ�
                    blnUpdate = True
                End If
            End If
        End If
        
        If rsData!���� = "�ɱ���" Then
            gstrSQL = "Select ���� As �ּ� " & _
                " From ҩƷ�շ���¼ " & _
                " Where ID = (Select Max(ID) " & _
                " From ҩƷ�շ���¼ B " & _
                " Where b.ҩƷid = [1] And b.���� = [2] And b.���� = 18 And b.������� > [3] And b.ժҪ = '�������ϳɱ��۵���') "
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", lng����ID, lng����, CDate(str��������))
            
            If rsprice.RecordCount > 0 Then
                If Round(rsprice!�ּ�, 2) <> Round(dblԭ��, 2) Then
                    dbl�ֳɱ��� = rsprice!�ּ�
                    blnUpdate = True
                End If
            End If
        End If
        
        '�Ե�ǰ���¼۸����µ���������ݣ����ۡ����۽���ۣ�
        lngRows = mshBill.Rows - 1
        For lngRow = 1 To lngRows
            If lng����ID = Val(mshBill.TextMatrix(lngRow, 0)) And (dbl���ۼ� <> 0 Or dbl�ֳɱ��� <> 0) Then
                dbl���� = Val(mshBill.TextMatrix(lngRow, mCol����))
                If dbl���ۼ� <> 0 Then
                    dbl���ۼ� = Val(Format(dbl���ۼ� * Val(mshBill.TextMatrix(lngRow, mCol����ϵ��)), mFMT.FM_���ۼ�))
                    dbl���۽�� = dbl���ۼ� * dbl����
                Else
                    dbl���ۼ� = Val(mshBill.TextMatrix(lngRow, mCol�ۼ�))
                    dbl���۽�� = Val(mshBill.TextMatrix(lngRow, mCol�ۼ۽��))
                End If
                
                If dbl�ֳɱ��� <> 0 Then
                    dbl�ֳɱ��� = Val(Format(dbl�ֳɱ��� * Val(mshBill.TextMatrix(lngRow, mCol����ϵ��)), mFMT.FM_�ɱ���))
                    dbl�ɱ���� = dbl�ֳɱ��� * dbl����
                Else
                    dbl�ֳɱ��� = Val(mshBill.TextMatrix(lngRow, mCol�����))
                    dbl�ɱ���� = Val(mshBill.TextMatrix(lngRow, mCol������))
                End If
                
                dbl��� = dbl���۽�� - dbl�ɱ����
                
                mshBill.TextMatrix(lngRow, mCol�����) = Format(dbl�ֳɱ���, mFMT.FM_�ɱ���)
                mshBill.TextMatrix(lngRow, mCol������) = Format(dbl�ɱ����, mFMT.FM_���)
                mshBill.TextMatrix(lngRow, mCol�ۼ�) = Format(dbl���ۼ�, mFMT.FM_���ۼ�)
                mshBill.TextMatrix(lngRow, mCol�ۼ۽��) = Format(dbl���۽��, mFMT.FM_���)
                mshBill.TextMatrix(lngRow, mCol���) = Format(dbl���, mFMT.FM_���)
                
                ''���˺�:���ۼ۴���
                Call �������ۼۼ����۲��(lngRow)
            End If
        Next
        
        rsData.MoveNext
    Loop
    
    CheckValuePrice = Not blnUpdate
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'�������������
Private Function GetDepend() As Boolean
    Dim rstemp As New Recordset
    
    On Error GoTo ErrHandle
    GetDepend = False
    
    gstrSQL = "" & _
        "   SELECT B.Id " & _
        "   FROM ҩƷ�������� A, ҩƷ������ B " & _
        "   Where A.���id = B.ID " & _
        "           AND A.���� = 30 and rownum=1 "
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ����������������")
    If rstemp.EOF Then
        ShowMsgBox "û���������������⹺����������������������������࣡"
        rstemp.Close
        Exit Function
    End If
    rstemp.Close
   
    gstrSQL = "" & _
        "   Select id " & _
        "   From ��Ӧ�� " & _
        "   where (����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or ����ʱ�� is null)  " & _
        "           And substr(����,5,1)=1 and Nvl(ĩ��,0)=1 and (վ��=[1] or վ�� is null) and rownum=1 "
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�湩Ӧ��", gstrNodeNo)
    If rstemp.EOF Then
        ShowMsgBox "û�������������Ϲ�Ӧ��λ�����ڹ�Ӧ�̹��������ã�"
        rstemp.Close
        Exit Function
    End If
    rstemp.Close
    GetDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, _
    Optional int��¼״̬ As Integer = 1, Optional strPrivs As String, Optional blnSuccess As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������
    '--�����:frmnMain-���õĴ���
    '--       str���ݺ�-���ݺ�;int�༭״̬;int��¼״̬;strPrivs-Ȩ�޴�
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    
    Set mfrmMain = FrmMain
    
    mblnSave = False
    mblnSuccess = False
    
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    
    mint��¼״̬ = int��¼״̬
        
    mblnSuccess = blnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = strPrivs
     
    mbln�޸������� = IIf(Val(zlDatabase.GetPara("�޸Ĳɹ��޼�", glngSys, mlngModule, "0")) = 1, 1, 0) = 1
    mbln��Ӧ��У�� = (Val(zlDatabase.GetPara("У�鹩Ӧ������", glngSys, mlngModule, "0")) = 1)
   
    
    Call GetRegInFor(g˽��ģ��, "�����⹺������", "���ݺ��ۼ�", strReg)
    mbln�������� = IIf(strReg = "", True, Val(strReg) = 1)
    
    If Not GetDepend Then Exit Sub
    
    cmdCopy.Visible = (mint�༭״̬ = 1 Or mint�༭״̬ = 2)
    txtCopy.Visible = (mint�༭״̬ = 1 Or mint�༭״̬ = 2)
    lblCopy.Visible = (mint�༭״̬ = 1 Or mint�༭״̬ = 2)
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 8 Then '�������˻�
        mblnEdit = True

        txtNO.Locked = True
        txtNO.TabStop = True
        txtNO = mstr���ݺ�
        txtNO.Tag = txtNO
    ElseIf mint�༭״̬ = 2 Or mint�༭״̬ = 7 Then    '�޸Ļ�������
        mblnEdit = True
        
        If mint�༭״̬ = 2 Then
            txtNO.Locked = True
            txtNO.TabStop = True
        End If
    ElseIf mint�༭״̬ = 3 Then            '���
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
        chkת���ƿ�.Visible = True
    ElseIf mint�༭״̬ = 4 Then            '����
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If InStr(mstrPrivs, "���ݴ�ӡ") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
        vsfCostlyInfo.Editable = flexEDNone
    ElseIf mint�༭״̬ = 5 Then        '�޸ķ�Ʊ
        mblnEdit = False
    ElseIf mint�༭״̬ = 6 Then        '����
        mblnEdit = False
        CmdSave.Caption = "����(&O)"
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
    ElseIf mint�༭״̬ = 9 Then
        '�˲鹦��
        mblnEdit = False
    End If
    fraMoveNO.Visible = mint�༭״̬ = 3    '���

    LblTitle.Caption = GetUnitName & IIf(mint�༭״̬ = 8, "���������˻���", LblTitle.Caption)
    Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub

Private Sub cboEnterStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If cboEnterStock.ListIndex >= 0 Then
        If Val(cboEnterStock.Tag) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
            Exit Sub
        End If
    End If
    
    Dim blnOptionerPrivs As Boolean
    
    blnOptionerPrivs = Not zlStr.IsHavePrivs(mstrPrivs, "���пⷿ")
    If Select����ѡ����(Me, cboEnterStock, Trim(cboEnterStock.Text), "V,W,K,12", blnOptionerPrivs) = False Then
        Exit Sub
    End If
    If cboEnterStock.ListIndex >= 0 Then
        cboEnterStock.Tag = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    End If
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_Click()
    Call ��ǰ��Ϊ�ⷿ
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
            For i = 1 To mshBill.Rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.Rows Then
                If MsgBox("����ı�ⷿ���п���Ҫ�ı���Ӧ�������ϵĵ�λ����Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '�����������ϵ�λ�ı�
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
    End With
End Sub

Private Sub cboType_Change()
    mblnChange = True
End Sub

Private Sub cboType_Click()
    Dim bln�ƿ� As Boolean
        
    bln�ƿ� = cboType.ItemData(cboType.ListIndex) = 0

    cboEnterStock.Visible = bln�ƿ�
    txtDraw.Visible = Not bln�ƿ�
    cmdDraw.Visible = Not bln�ƿ�
    txtDrawPerson.Visible = Not bln�ƿ�
    cmdDrawPerson.Visible = Not bln�ƿ�
    lbl������.Visible = Not bln�ƿ�
    lbl������.Enabled = lbl������.Visible
    
End Sub

Private Sub cboType_GotFocus()
    OS.OpenIme False
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub chkת���ƿ�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mCol��������) = Format(0, mFMT.FM_����)
                .TextMatrix(intRow, mCol������) = Format(0, mFMT.FM_���)
                .TextMatrix(intRow, mCol�ۼ۽��) = Format(0, mFMT.FM_���)
                .TextMatrix(intRow, mCol���) = Format(0, mFMT.FM_���)
                '���˺�:���ۼ۴���
                .TextMatrix(intRow, mcol���۽��) = Format(0, mFMT.FM_���)
                .TextMatrix(intRow, mcol���۲��) = Format(0, mFMT.FM_���)
                If Trim(.TextMatrix(intRow, mCol��Ʊ��)) <> "" Then
                    .TextMatrix(intRow, mCol��Ʊ���) = Format(0, mFMT.FM_���)
                End If
            End If
        Next
    End With
End Sub

Private Sub cmdAllSel_Click()
    Dim rstemp As New Recordset
    Dim intRow As Integer, dbl��Ʊ��� As Double
    
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And .RowData(intRow) = 0 Then
                .TextMatrix(intRow, mCol��������) = .TextMatrix(intRow, mCol����)
                .TextMatrix(intRow, mCol������) = Format(Val(.TextMatrix(intRow, mCol����)) * Val(.TextMatrix(intRow, mCol�����)), mFMT.FM_���)
                .TextMatrix(intRow, mCol�ۼ۽��) = Format(Val(.TextMatrix(intRow, mCol����)) * Val(.TextMatrix(intRow, mCol�ۼ�)), mFMT.FM_���)
                .TextMatrix(intRow, mCol���) = Format(Val(.TextMatrix(intRow, mCol�ۼ۽��)) - Val(.TextMatrix(intRow, mCol������)), mFMT.FM_���)
                
                '���˺�:���ۼ۴���,��Ҫ��ȷ��ʱ�۶�������
                Call �������ۼۼ����۲��(intRow, False)
                If Trim(.TextMatrix(intRow, mCol��Ʊ��)) <> "" Or Trim(.TextMatrix(intRow, mCol�������)) <> "" Then
                
                    dbl��Ʊ��� = GetTotale��Ʊ���(mstr���ݺ�, Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(.Row, mCol���)))
                    If dbl��Ʊ��� = 0 Then dbl��Ʊ��� = Val(.TextMatrix(intRow, mCol������))
                    .TextMatrix(intRow, mCol��Ʊ���) = Format(dbl��Ʊ���, mFMT.FM_���)
                End If
            End If
        Next
    End With
End Sub
Public Function GetTotale��Ʊ���(ByVal strNo As String, ByVal lng����ID As Long, ByVal lng��� As Long) As Double
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-19 14:43:30
    '-----------------------------------------------------------------------------------------------------------
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select sum(nvl(Q.��Ʊ���,0)) as ��Ʊ��� " & _
        "   From ҩƷ�շ���¼ x," & _
        "        ( Select B.ID, B.�շ�id,Sum(B.��Ʊ���) as ��Ʊ���, Max(B.��Ʊ��) As ��Ʊ��,Max(B.�������) as �������, Max(B.��Ʊ����) As ��Ʊ����, Max(B.�������) As ������� " & _
        "          From ҩƷ�շ���¼ A,Ӧ����¼ B " & _
        "          Where A.ID = B.�շ�id And A.NO =[1] and A.ҩƷID=[2] And A.���=[3]  And A.���� = 15 And B.ϵͳ��ʶ = 5 And B.��¼���� In (0, -1) " & _
        "          Group By B.ID,B.�շ�id ) Q " & _
        "   WHERE x.id=q.�շ�id(+) AND X.����=15" & _
        "         and X.NO=[1] and X.ҩƷid=[2] and x.���=[3]"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strNo, lng����ID, lng���)
    If rstemp.EOF Then
        GetTotale��Ʊ��� = 0
    Else
        GetTotale��Ʊ��� = IIf(mbln�˻�, -1, 1) * Val(zlStr.Nvl(rstemp!��Ʊ���))
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub cmdBulkCopy_Click()
    Dim i As Integer
    
    With mshBill
        '1�����з�Ʊ��������2��
        For i = 1 To .Rows - 1
           If Trim(.TextMatrix(i, mCol��Ʊ��)) = "" Or .TextMatrix(i, 0) = "" Then Exit For
        Next
        
        If i = .Rows - 1 Then Exit Sub
        
        '2����Ʊ�����Ʊ����Ϊ�գ�����ʾ
        If Trim(.TextMatrix(.Row, mcol��Ʊ����)) = "" Or .TextMatrix(.Row, mCol��Ʊ����) = "" Then
            If MsgBox("��Ʊ�����Ʊ����Ϊ�գ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("�Ƿ񽫸��еķ�Ʊ��Ϣ�������Ƶ���Ʊ��Ϊ�յ��У�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        '3������
        For i = 1 To .Rows - 1
            If i <> .Row And Trim(.TextMatrix(i, mCol��Ʊ��)) = "" And .TextMatrix(i, 0) <> "" Then    '���Ǳ༭���ҷ�Ʊ��Ϊ�յ������޸�
                
                .TextMatrix(i, mCol��Ʊ��) = .TextMatrix(.Row, mCol��Ʊ��)
                .TextMatrix(i, mcol��Ʊ����) = .TextMatrix(.Row, mcol��Ʊ����)
                .TextMatrix(i, mCol��Ʊ����) = .TextMatrix(.Row, mCol��Ʊ����)
                .TextMatrix(i, mCol��Ʊ���) = .TextMatrix(i, mCol������)
                
            End If
        Next
    End With
End Sub

Private Sub cmdALLDel_Click()
    Dim i As Integer
    
    With mshBill
        '1�����޷�Ʊ��������2��
        For i = 1 To .Rows - 1
           If Trim(.TextMatrix(i, mCol��Ʊ��)) <> "" Or .TextMatrix(i, 0) = "" Then Exit For
        Next
        
        If i = .Rows - 1 Then Exit Sub
    
        If MsgBox("�ò�������������еķ�Ʊ������ݣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            For i = 1 To .Rows - 1
            
                If Trim(.TextMatrix(i, mCol��Ʊ��)) <> "" And .TextMatrix(i, 0) <> "" Then
                    .TextMatrix(i, mCol��Ʊ��) = ""
                    .TextMatrix(i, mcol��Ʊ����) = ""
                    .TextMatrix(i, mCol��Ʊ����) = ""
                    .TextMatrix(i, mCol��Ʊ���) = ""
                End If
                
            Next
            
            cmdBulkCopy.Enabled = False
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    '��¼���������Ƶ�ǰ�����ݵ��������ݺ���
    Dim lngCopyNum As Long
    Dim lngMoveRowStart As Long, lngMoveRowEnd As Long
    Dim i As Long
    Dim intCol As Integer
    Dim lngRow As Long
    Dim str���� As String
    Dim str��λ As String
    
    With mshBill
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        If mbln��ֵ���� = False Then Exit Sub
        
        lngCopyNum = Val(Trim(txtCopy.Text))
        If lngCopyNum = 0 Then
            MsgBox "��¼�븴�Ƶ�������", vbInformation, gstrSysName
            txtCopy.SetFocus
            Exit Sub
        End If
        
        str���� = .TextMatrix(.Row, mCol����)
        str��λ = .TextMatrix(.Row, mCol��λ)
        
        '����
        If MsgBox("�Ƿ�����������Ϊ��" & str���� & "�������Ĺ���" & lngCopyNum & "" & str��λ & " ?", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                
        lngRow = .Row
        
        '��¼��ǰ�к������ݵ���ʼ�к�
        If .Row = .Rows - 1 Then
            '��ǰ�������һ��ʱ
            lngMoveRowStart = 0
            lngMoveRowEnd = 0
        Else
            '��ǰ�в������һ��ʱ
            lngMoveRowStart = .Row + 1
            lngMoveRowEnd = .Rows - 1
        End If
        
        '��������
        .Rows = .Rows + lngCopyNum
                
        '�ѵ�ǰ�к�����������������ƶ�
        If lngMoveRowStart <> 0 Then
            For i = lngMoveRowEnd To lngMoveRowStart Step -1
                For intCol = 0 To .Cols - 1
                    If mCol�к� = intCol Then
                        .TextMatrix(i + lngCopyNum, intCol) = i + lngCopyNum
                    Else
                        .TextMatrix(i + lngCopyNum, intCol) = .TextMatrix(i, intCol)
                    End If
                Next
            Next
        End If
        
        '���Ƶ�ǰ��
        For i = 1 To lngCopyNum
            For intCol = 0 To .Cols - 1
                If mCol�к� = intCol Then
                    .TextMatrix(i + lngRow, intCol) = i + lngRow
                Else
                    .TextMatrix(i + lngRow, intCol) = .TextMatrix(lngRow, intCol)
                End If
            Next
        Next
    End With
End Sub

Private Sub cmdDrawPerson_Click()
    If SelectItem(txtDrawPerson, "", True) = False Then Exit Sub
End Sub

Private Sub cmdExtractData_Click()
    If Val(txtProvider.Tag) <= 0 Then
        txtProvider.SetFocus
        MsgBox "��¼�빩�̵�λ��Ϣ��", vbInformation, gstrSysName
        Exit Sub
    End If
    With frmPurchaseCardExtract
        .EntryPort cboStock.ItemData(cboStock.ListIndex) & ";" & cboStock.Text, txtProvider.Tag
        .Show vbModal, Me
    End With
    With mshBill
        .Row = 1
        .SetFocus
    End With
End Sub

'����
Private Sub cmdFind_Click()
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRownew mshBill, mCol����, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
    
    If mint�༭״̬ = 5 Then '�޸ķ�Ʊ��Ϣ�ð�ť�ſ���
        With cmdALLDel
            cmdALLDel.Visible = True
            .Top = cmdFind.Top
            If txtCode.Visible Then
                .Left = txtCode.Left + txtCode.Width + 100
            Else
                .Left = cmdFind.Left + cmdFind.Width + 100
            End If
        End With
    End If
    
    With cmdCopy
        .Left = IIf(txtCode.Visible, txtCode.Left + txtCode.Width, cmdFind.Left + cmdFind.Width) + 100
        .Top = cmdFind.Top
    End With
    
    With txtCopy
        .Left = cmdCopy.Left + cmdCopy.Width + 50
        .Top = txtCode.Top
    End With
    
    With lblCopy
        .Left = txtCopy.Left + txtCopy.Width + 25
        .Top = txtCopy.Top + 45
    End With
        
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub CmdNO_Click()

    Dim mdbl�Ӽ��� As Double
    Dim dbl�ɱ��� As Double
    
    With mshBill
        mdbl�Ӽ��� = Val(Txt�Ӽ���.Tag)
        If mblnʱ�۹�ǰ���� Then
            dbl�ɱ��� = Val(.TextMatrix(.Row, mCol�ɹ���))
        Else
            dbl�ɱ��� = Val(.TextMatrix(.Row, mCol�����))
        End If
        
        '���¼������ۼۡ����
        If mint�༭״̬ = 8 And Val(.TextMatrix(.Row, mCol�ۼ�)) <> 0 Then
        Else
            .TextMatrix(.Row, mCol�ۼ�) = Format(У�����ۼ�(dbl�ɱ��� * (1 + (mdbl�Ӽ��� / 100)) + _
             ʱ�۲������ۼ�(Val(.TextMatrix(.Row, 0)), dbl�ɱ���, mdbl�Ӽ��� / 100)), mFMT.FM_���ۼ�)
        End If
        .TextMatrix(.Row, mCol�ۼ۽��) = Format(Val(.TextMatrix(.Row, mCol�ۼ�)) * Val(.TextMatrix(.Row, mCol����)), mFMT.FM_���)
        .TextMatrix(.Row, mCol���) = Format(IIf(.TextMatrix(.Row, mCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mCol������) = "", 0, .TextMatrix(.Row, mCol������)), mFMT.FM_���)
        
        '���˺�:���ۼ۴���
        Call �������ۼۼ����۲��(.Row, True)
    End With
    PicInput.Visible = False
End Sub

Private Sub CmdNO_LostFocus()
    Call PicInput_LostFocus
End Sub

Private Sub CmdYes_Click()
    If Val(Txt�Ӽ���) > 9900 Or Val(Txt�Ӽ���) < 0 Then
        MsgBox "������Ϸ��ļӳ��ʣ���0-9900��", vbInformation, gstrSysName
        Txt�Ӽ���.SetFocus
        Exit Sub
    End If
    Dim dbl�ɱ��� As Double
    With mshBill
        If mblnʱ�۹�ǰ���� Then
            dbl�ɱ��� = Val(.TextMatrix(.Row, mCol�ɹ���))
        Else
            dbl�ɱ��� = Val(.TextMatrix(.Row, mCol�����))
        End If
        '���¼������ۼۡ����
        If mint�༭״̬ = 8 And Val(.TextMatrix(.Row, mCol�ۼ�)) <> 0 Then
        Else
            .TextMatrix(.Row, mCol�ۼ�) = Format(У�����ۼ�(dbl�ɱ��� * (1 + (Val(Txt�Ӽ���) / 100)) + _
            ʱ�۲������ۼ�(Val(.TextMatrix(.Row, 0)), dbl�ɱ���, Val(Txt�Ӽ���) / 100)), mFMT.FM_���ۼ�)
        End If
         
        .TextMatrix(.Row, mcol�ӳ���) = zlStr.FormatEx(Val(Txt�Ӽ���), 2) & "%"
        .TextMatrix(.Row, mCol�ۼ۽��) = Format(Val(.TextMatrix(.Row, mCol�ۼ�)) * Val(.TextMatrix(.Row, mCol����)), mFMT.FM_���)
        .TextMatrix(.Row, mCol���) = Format(IIf(.TextMatrix(.Row, mCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mCol������) = "", 0, .TextMatrix(.Row, mCol������)), mFMT.FM_���)
        
        '���˺�:���ۼ۴���
        Call �������ۼۼ����۲��(.Row, True)
    End With
    
    PicInput.Visible = False
    mshBill.SetFocus
End Sub

Private Sub CmdYes_LostFocus()
    Call PicInput_LostFocus
End Sub

Private Sub Form_Activate()
'    mblnChange = False
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            '�����ѱ�ɾ��
            If mint�༭״̬ = 5 Then
               MsgBox "�õ����ѱ�ȫ�������������޸ķ�Ʊ��Ϣ�����飡", vbOKOnly, gstrSysName
            ElseIf mint�༭״̬ = 6 Then
                MsgBox "�õ�����û�п��Գ������������ϣ����飡", vbOKOnly, gstrSysName
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
        Case 6  '����
            MsgBox "�õ����Ѹ�������ܽ��г�����", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
    End Select
    SetEdit
    If IsCtrlSetFocus(txtProvider) Then
        zlControl.ControlSetFocus txtProvider
    Else
        If mint�༭״̬ = 3 And IsCtrlSetFocus(chkת���ƿ�) Then
             zlControl.ControlSetFocus chkת���ƿ�
        Else
            zlControl.ControlSetFocus mshBill
        End If
    End If
    '��ʼ�����뷽ʽ
    If (mint�༭״̬ = 1 Or mint�༭״̬ = 2) And gbytSimpleCodeTrans = 1 Then
        stbThis.Panels("PY").Visible = True
        stbThis.Panels("WB").Visible = True
        gSystem_Para.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram stbThis, gSystem_Para.int���뷽ʽ
    Else
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single, sngTop As Single
    
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRownew mshBill, mCol����, txtCode.Text, False
    ElseIf KeyCode = vbKeyF4 Then
        '���ϵͳ����Ϊ�棬����ʾ�û�����Ӽ���
        If mbln�Ӽ��� And (mint�༭״̬ = 1 Or mint�༭״̬ = 2) Then
            If PicInput.Visible Then PicInput.SetFocus: Exit Sub
            If mshBill.TextMatrix(mshBill.Row, mCol����) = "" Then Exit Sub
            '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
            If Split(mshBill.TextMatrix(mshBill.Row, mColԭ����), "||")(2) <> 1 Then Exit Sub
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
            Txt�Ӽ��� = "15.0000"
            With mshBill
                If Val(.TextMatrix(.Row, mCol�ۼ�)) <> 0 And Val(.TextMatrix(.Row, mCol�����)) <> 0 Then
                    Txt�Ӽ��� = Format((Val(.TextMatrix(.Row, mCol�ۼ�)) / Val(.TextMatrix(.Row, mCol�����)) - 1) * 100, "####0.0000000;-####0.0000000;0;0")
                End If
            End With
            Txt�Ӽ���.Tag = Txt�Ӽ���
            Txt�Ӽ���.SetFocus
        End If
    ElseIf KeyCode = vbKeyF7 Then
        If stbThis.Panels("PY").Bevel = sbrRaised Then
            Logogram stbThis, 0
        Else
            Logogram stbThis, 1
        End If
    End If
End Sub

Private Sub cmdProvider_Click()
    Dim rstemp As New Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txtProvider.hwnd)
    
    gstrSQL = "" & _
        "   Select id,�ϼ�ID,����,����,����,ĩ�� " & _
        "   From ��Ӧ�� " & _
        "   Where  (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
        "       And (substr(����,5,1)=1 And (վ��=[1] or վ�� is null)  Or Nvl(ĩ��,0)=0) " & _
        "   Start with �ϼ�ID is null connect by prior ID =�ϼ�ID " & _
        "   Order by level,ID"
    
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
    
    Set rstemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 2, "��Ӧ��ѡ��", True, "", "��ѡ������������ϵĹ�Ӧ��", True, True, False, vRect.Left - 15, vRect.Top, txtProvider.Height, blnCancel, False, False, gstrNodeNo)
        
    If rstemp Is Nothing Or blnCancel Then Exit Sub
    If rstemp.State <> 1 Then Exit Sub
    
    With rstemp
        Me.txtProvider = "[" & zlStr.Nvl(!����) & "] " & zlStr.Nvl(!����)
        Me.txtProvider.Tag = zlStr.Nvl(!Id)
    End With
    If mshBill.Col = 1 Then mshBill.Col = mCol����
    mshBill.SetFocus
    
    If CheckQualifications(mlngModule, 2, Val(txtProvider.Tag)) = False Then
        txtProvider.Text = ""
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If Val(txtProvider.Tag) <> mlng������λID And (mint�༭״̬ = 8 Or mbln�˻�) Then     '�˻�
        mlng������λID = Val(txtProvider.Tag)
        mshBill.ClearBill
        mshBill.TextMatrix(1, mCol�к�) = "1"
    End If
End Sub

'��ӡ����
Private Sub printbill()
    Dim strUnit As String
    Dim strNo As String
    strNo = txtNO.Tag
    
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1712", _
        mint��¼״̬, mintUnit, "1712", IIf(mint�༭״̬ = 8 Or mbln�˻�, "�����˻���", "�����⹺��ⵥ"), strNo
End Sub

Private Function SaveNewCard(ByVal strNo As String) As Boolean
    '���ܣ�������˲����µ���
    '����strNO ����������µ���no
    Dim chrNo As Variant
    Dim lng��� As Long
    Dim lngStockID As Long
    Dim lng������λid As Long
    Dim lng����ID As Long
    Dim str���� As String
    Dim str���� As String
    Dim strЧ�� As String
    Dim dblʵ������ As Double
    Dim dbl�ɱ��� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl���� As Double
    Dim dbl���ۼ� As Double
    Dim dbl���۽�� As Double
    Dim dbl��� As Double
    Dim str���۲�� As String '�շ���¼�����÷��ֶα�������⹺����۲�����÷��ֶ��������ַ�������������double���ͻ���� -.00x����
    Dim strժҪ As String
    Dim str������ As String
    Dim str�������� As String
    Dim str����� As String
    Dim datAssessDate As String
    Dim str��Ʊ�� As String
    Dim str��Ʊ���� As String
    Dim str��Ʊ���� As String
    Dim str������� As String
    Dim str���ʧЧ�� As String
    Dim dbl��Ʊ��� As Double
    Dim str��������  As String
    Dim str�˲��� As String
    Dim str�˲����� As String
    Dim strע��֤�� As String
    Dim intUnit As Integer
    Dim strUnit As String
    Dim strָ�������� As String
    Dim str������� As String
    Dim str���ս��� As String
    Dim str��Ʒ���� As String
    Dim str�ڲ����� As String
    Dim str��׼�ĺ� As String
    Dim lng����ID As Long
    Dim intRow As Integer
    Dim i As Integer
    Dim arrSQL As Variant
    Dim n As Long
    
    SaveNewCard = False
    arrSQL = Array()
    With mshBill

        chrNo = Trim(txtNO)
        lngStockID = cboStock.ItemData(cboStock.ListIndex)
        lng������λid = txtProvider.Tag
        strժҪ = Trim(txtժҪ.Text)

        str������ = Txt������
        str�������� = Txt��������.Caption
        str�˲��� = txt�˲���
        str�˲����� = txt�˲�����.Caption
        str����� = Txt�����

        On Error GoTo ErrHandle

        'ȡ�ÿⷿ�ĵ�λ������ָ��������ʱʹ��
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng����ID = .TextMatrix(intRow, 0)
                str���� = .TextMatrix(intRow, mCol����)
                str���� = .TextMatrix(intRow, mCol����)
                str��׼�ĺ� = .TextMatrix(intRow, mcol��׼�ĺ�)
                strЧ�� = IIf(.TextMatrix(intRow, mColЧ��) = "", "", .TextMatrix(intRow, mColЧ��))
                dblʵ������ = GetFormat(.TextMatrix(intRow, mCol����) * .TextMatrix(intRow, mCol����ϵ��), g_С��λ��.obj_���С��.����С��)
                dbl���� = Val(.TextMatrix(intRow, mCol����))
                dbl�ɱ��� = GetFormat(Val(.TextMatrix(intRow, mCol�����)) / .TextMatrix(intRow, mCol����ϵ��), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl�ɱ���� = GetFormat(Val(.TextMatrix(intRow, mCol������)), g_С��λ��.obj_���С��.���С��)

                'dbl���ۼ� = Round(Val(.TextMatrix(intRow, mCol�ۼ�)) / .TextMatrix(intRow, mCol����ϵ��), g_С��λ��.obj_ɢװС��.���ۼ�С��)
                'dbl���۽�� = Round(Val(.TextMatrix(intRow, mCol�ۼ۽��)), g_С��λ��.obj_ɢװС��.���С��)
                '���ݿ��е�:��� = ���۽�� - ������
                '���ݿ��е�:�÷� = ���۽��-�ۼ۽������۲��-���(�ⷿ��λ�Ĳ��)

                dbl���ۼ� = GetFormat(Val(.TextMatrix(intRow, mcol���ۼ�)), g_С��λ��.obj_���С��.���ۼ�С��)
                dbl���۽�� = GetFormat(Val(.TextMatrix(intRow, mcol���۽��)), g_С��λ��.obj_���С��.���ۼ�С��)
                dbl��� = GetFormat(Val(.TextMatrix(intRow, mcol���۲��)), g_С��λ��.obj_���С��.���ۼ�С��)
                str���۲�� = GetFormat(Val(.TextMatrix(intRow, mcol���۲��)) - Val(.TextMatrix(intRow, mCol���)), g_С��λ��.obj_���С��.���ۼ�С��)
                'dbl��� = Round(Val(.TextMatrix(intRow, mCol���)), g_С��λ��.obj_ɢװС��.���С��)
                lng��� = .TextMatrix(intRow, mCol���)

                str������� = Trim(.TextMatrix(intRow, mCol�������))
                str���ս��� = Trim(.TextMatrix(intRow, mCol���ս���))
                str��Ʊ�� = Trim(.TextMatrix(intRow, mCol��Ʊ��))
                str��Ʊ���� = Trim(.TextMatrix(intRow, mcol��Ʊ����))
                str��Ʊ���� = Trim(IIf(.TextMatrix(intRow, mCol��Ʊ����) = "", "", .TextMatrix(intRow, mCol��Ʊ����)))
                dbl��Ʊ��� = Round(Val(.TextMatrix(intRow, mCol��Ʊ���)), g_С��λ��.obj_ɢװС��.���С��)

                str������� = Trim(IIf(.TextMatrix(intRow, mcol�������) = "", "", .TextMatrix(intRow, mcol�������)))
                str���ʧЧ�� = Trim(IIf(.TextMatrix(intRow, mcol���ʧЧ��) = "", "", .TextMatrix(intRow, mcol���ʧЧ��)))
                str�������� = Trim(IIf(.TextMatrix(intRow, mcol��������) = "", "", .TextMatrix(intRow, mcol��������)))
                strע��֤�� = Trim(.TextMatrix(intRow, mcolע��֤��))

                str�ڲ����� = Trim(.TextMatrix(intRow, mcol�ڲ�����))
                '��������µ��ݷ���id����2
                lng����ID = 2

                str��Ʒ���� = Trim(.TextMatrix(intRow, mcol��Ʒ����))

                ' Zl_�����⹺_Insert
                gstrSQL = "zl_�����⹺_INSERT("
                '  No_In         In ҩƷ�շ���¼.NO%Type,
                gstrSQL = gstrSQL & "'" & strNo & "',"
                '  ���_In       In ҩƷ�շ���¼.���%Type,
                gstrSQL = gstrSQL & "" & lng��� & ","
                '  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
                gstrSQL = gstrSQL & "" & lngStockID & ","
                '  ��ҩ��λid_In In ҩƷ�շ���¼.��ҩ��λid%Type,
                gstrSQL = gstrSQL & "" & lng������λid & ","
                '  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
                gstrSQL = gstrSQL & "" & lng����ID & ","
                '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "'" & str���� & "',"
                '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "'" & str���� & "',"
                '  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str�������� = "", "Null", "to_date('" & Format(str��������, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  �������_In   In ҩƷ�շ���¼.�������%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str������� = "", "Null", "to_date('" & Format(str�������, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  ���Ч��_In   In ҩƷ�շ���¼.���Ч��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str���ʧЧ�� = "", "Null", "to_date('" & Format(str���ʧЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  ʵ������_In   In ҩƷ�շ���¼.ʵ������%Type := Null,
                gstrSQL = gstrSQL & "" & dblʵ������ & ","
                '  �ɱ���_In     In ҩƷ�շ���¼.�ɱ���%Type := Null,
                gstrSQL = gstrSQL & "" & dbl�ɱ��� & ","
                '  �ɱ����_In   In ҩƷ�շ���¼.�ɱ����%Type := Null,
                gstrSQL = gstrSQL & "" & dbl�ɱ���� & ","
                '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "" & dbl���� & ","
                '  ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type := Null,
                gstrSQL = gstrSQL & "" & dbl���ۼ� & ","
                '  ���۽��_In   In ҩƷ�շ���¼.���۽��%Type := Null,
                gstrSQL = gstrSQL & "" & dbl���۽�� & ","
                '  ���_In       In ҩƷ�շ���¼.���%Type := Null,
                gstrSQL = gstrSQL & "" & dbl��� & ","
                '  ���۲��_In   In ҩƷ�շ���¼.���%Type := Null,Ŀǰ������÷��ֶ�
                gstrSQL = gstrSQL & "" & str���۲�� & ","
                '  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(strժҪ = "", "NULL", "'" & strժҪ & "'") & ","
                '   ע��֤��_In   In ҩƷ�շ���¼.ע��֤��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(strע��֤�� = "", "NULL", "'" & strע��֤�� & "'") & ","
                '  ������_In     In ҩƷ�շ���¼.������%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str������ = "", "NULL", "'" & str������ & "'") & ","
                '  �������_In   In Ӧ����¼.�������%Type := Null
                gstrSQL = gstrSQL & "" & IIf(str������� = "", "NULL", "'" & str������� & "'") & ","
                '  ��Ʊ��_In     In Ӧ����¼.��Ʊ��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str��Ʊ�� = "", "NULL", "'" & str��Ʊ�� & "'") & ","
                '  ��Ʊ����_In   In Ӧ����¼.��Ʊ����%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str��Ʊ���� = "", "Null", "to_date('" & Format(str��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  ��Ʊ���_In   In Ӧ����¼.��Ʊ���%Type := Null,
                gstrSQL = gstrSQL & "" & dbl��Ʊ��� & ","
                '  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
                gstrSQL = gstrSQL & "to_date('" & str�������� & "','yyyy-mm-dd HH24:MI:SS'),"
                '  �˲���_In     In ҩƷ�շ���¼.��ҩ��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str�˲��� = "", "NULL", "'" & str�˲��� & "'") & ","
                '  �˲�����_In   In ҩƷ�շ���¼.��ҩ����%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str�˲����� = "", "Null", "to_date('" & str�˲����� & "','yyyy-mm-dd hh24:mi:ss')") & ","
                '  ����_In       In ҩƷ�շ���¼.����%Type := 0,
                gstrSQL = gstrSQL & "" & Val(.TextMatrix(intRow, mCol����)) & ","
                '  �˻�_In       In Number := 1
                gstrSQL = gstrSQL & "" & IIf(mbln�˻�, -1, 1) & ","
                '  ��ֵ����_In   In varchar2(250)
                gstrSQL = gstrSQL & "'" & GetCostlyInfoStr(intRow) & "'" & ","
                '  ��Ʒ����_In   In ҩƷ�շ���¼.��Ʒ����%Type :=Null
                gstrSQL = gstrSQL & "" & IIf(str��Ʒ���� = "", "NULL", "'" & str��Ʒ���� & "'") & ","
                '  �ڲ�����
                gstrSQL = gstrSQL & IIf(str�ڲ����� = "", "Null", "'" & str�ڲ����� & "'") & ","
                '  ����ID
                gstrSQL = gstrSQL & lng����ID & ","
                '  ��Ʊ����
                gstrSQL = gstrSQL & IIf(str��Ʊ���� = "", "NULL", "'" & str��Ʊ���� & "'")
                '  �������
                gstrSQL = gstrSQL & ",1,"
                '  ��׼�ĺ�
                gstrSQL = gstrSQL & IIf(str��׼�ĺ� = "", "NULL", "'" & str��׼�ĺ� & "'") & ","
                '  ���ս���
                gstrSQL = gstrSQL & "'" & str���ս��� & "'"
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
            End If
            
            recSort.MoveNext
        Next
        
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "SaveCard")
        Next

        mstr���ݺ� = chrNo
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveNewCard = True
    Exit Function

ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveVerifyCard(ByVal strNo As String) As Boolean
    '���ܣ��������ʱ�������˼�¼���в�������
    '����ֵ:true-ִ�гɹ� false-ִ��ʧ��
    Dim str������� As String
    
    On Error GoTo ErrHand
    
    SaveVerifyCard = False
    
    gstrSQL = "Zl_ҩƷ�������_Insert("
    '�ⷿid
    gstrSQL = gstrSQL & cboStock.ItemData(cboStock.ListIndex)
    '����
    gstrSQL = gstrSQL & ",15"
    '����no
    gstrSQL = gstrSQL & ",'" & txtNO.Text & "'"
    'newNO
    gstrSQL = gstrSQL & ",'" & strNo & "'"
    '�����
    gstrSQL = gstrSQL & ",'" & UserInfo.�û��� & "'"
    '�������
    gstrSQL = gstrSQL & ",to_date('" & Format(mstr�������, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
    '��ע
    If Trim(txtժҪ.Text) = "" Then
        gstrSQL = gstrSQL & "," & "Null" & ")"
    Else
        gstrSQL = gstrSQL & ",'" & txtժҪ.Text & "')"
    End If
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
    SaveVerifyCard = True
    Exit Function
    
ErrHand:
If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub CmdSave_Click()
    On Error GoTo ErrHand
    Dim strNewNO As String
    Dim strReg As String
    Dim blnSuccess As Boolean, blnTrans As Boolean
    
    '�����������ݼ�
    Call SetSortRecord
    
    mstr������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 7 Then
        ' '������ˣ������������µ��ݲ���ˣ��Ѹ���ĵ��ݲ����������ˣ�ͬ����������˺�ĵ��ݲ����������
        '�ȳ��������������ݲ����
        gcnOracle.BeginTrans
        
        blnTrans = True
        '�����µ�no
        strNewNO = sys.GetNextNo(68, cboStock.ItemData(cboStock.ListIndex))
        blnSuccess = (strNewNO <> "")
        '�����µ�δ��˵Ĳ�����˵���
        If blnSuccess Then blnSuccess = SaveNewCard(strNewNO)
        '����ԭʼ����
        If blnSuccess Then blnSuccess = SaveStrike
        '����²����Ĳ�����˵���
        If blnSuccess Then blnSuccess = SaveCheck(strNewNO)
        '�������˼�¼���в�������
        If blnSuccess Then blnSuccess = SaveVerifyCard(strNewNO)
        
        If blnSuccess Then
            gcnOracle.CommitTrans
        Else
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        blnTrans = False
        Unload Me
        Exit Sub
    End If
    
    '���˺�:2007/05/14:���Ӻ˲���
    If mint�༭״̬ = 9 Then    '�˲�
        '���˺�:�����Ƕ���,��Ҫ�ȶԶ��۵ļ۸���е���.
        If mblnUpdate = False Then
            If Not ��鵥��(15, txtNO.Tag, False) Then
                '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
                ShowMsgBox "�м�¼δʹ�������ۼۣ������Զ���ɸ��£��ۼۡ��ۼ۽���ۣ������º����飡"
                Call RefreshBill
                mblnUpdate = True
                Exit Sub
            End If
        End If
        
        If mblnCheckPrice = False Then
            '�����ֵ������������������ⵥ���۸�
            If Not CheckValuePrice(15, txtNO.Tag) Then
                ShowMsgBox "��ֵ������ⵥ�м۸��ѵ��ۣ��������Զ���ɸ��£��ۼۡ��ۼ۽����ɱ��ۡ��ɱ�����ۣ�,���飡"
                mblnCheckPrice = True
                Exit Sub
            End If
        End If
        
        '���˺�:2007/06/10:����10813
        mstrTime_End = GetBillInfo(15, txtNO.Tag)
        If mstrTime_End = "" Then
            MsgBox "ע��:" & vbCrLf & "  �õ����Ѿ�����������Աɾ��,���ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        If mstrTime_End <> mstrTime_Start Then
            If MsgBox("ע��:" & vbCrLf & "  �õ����Ѿ�����������Ա�༭�����ܼ���!" & vbCrLf & "  �Ƿ�����ˢ�µ���?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                mshBill.ClearBill
                Call initCard
            End If
            Exit Sub
        End If
        
        If Not SaveCard Then Exit Sub
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 3 Then      '���
        If chkת���ƿ�.Value = 1 And mbln�˻� = False Then
            If cboType.ItemData(cboType.ListIndex) <> 0 Then
                If Val(txtDraw.Tag) = 0 Then
                    ShowMsgBox "δ������ص����ò���,����!"
                    zlControl.ControlSetFocus txtDraw, True
                    Exit Sub
                End If
                If Trim(txtDrawPerson.Tag) = "" And Trim(txtDrawPerson.Text) <> "" Then
                    If MsgBox("�����˲��ǵ�ǰ�������ŵ������Ա,�Ƿ����?", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        zlControl.ControlSetFocus txtDrawPerson, True
                        Exit Sub
                    End If
                End If
                If Trim(txtDrawPerson.Tag) = "" And Trim(txtDrawPerson.Text) = "" Then
                    If MsgBox("δ������ص�������,�Ƿ����?", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        zlControl.ControlSetFocus txtDrawPerson, True
                        Exit Sub
                    End If
                End If
            Else
                If cboEnterStock.ListIndex < 0 Then
                    ShowMsgBox "Ҫ�ƿ�Ĳ��Ų���ȷ��"
                    cboEnterStock.SetFocus
                    Exit Sub
                End If
                If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
                    ShowMsgBox "���벿�����Ƴ����Ų�����ͬ��"
                    cboEnterStock.SetFocus
                    Exit Sub
                End If
            End If
        End If
        If mblnUpdate = False Then
            If Not ��鵥��(15, txtNO.Tag, False) Then
                '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
                ShowMsgBox "�м�¼δʹ�������ۼۣ������Զ���ɸ��£��ۼۡ��ۼ۽���ۣ������º����飡"
                Call RefreshBill
                mblnUpdate = True
                Exit Sub
            End If
        End If
        
        If mblnCheckPrice = False Then
            '�����ֵ������������������ⵥ���۸�
            If Not CheckValuePrice(15, txtNO.Tag) Then
                ShowMsgBox "��ֵ������ⵥ�м۸��ѵ��ۣ��������Զ���ɸ��£��ۼۡ��ۼ۽����ɱ��ۡ��ɱ�����ۣ�,���飡"
                mblnCheckPrice = True
                Exit Sub
            End If
        End If
        
        If Not ���ϵ������(Txt������.Caption) Then Exit Sub
                
        If mbln�˻� = False Then
            '�����˻�����Ҫ���¸��¼۸�,���Ҽ۸���ԭ���۸�һ��ʱ,���ø��¼ۼۣ����û��Ҫ����
            blnTrans = True
            gcnOracle.BeginTrans
            
            If mblnUpdate Or mblnChange Or mblnCheckPrice Then
                '���˺�:2007/05/15
                '0.��ԭ���۸�һ�£������±��浥��
                '1.������ԭ����,Ҳ��Ҫ���±��浥��
                If Not SaveCard(True) Then
                    gcnOracle.RollbackTrans: Exit Sub
                End If
                If mbln��Ҫ�˲� Then
                    '���˺�:2007/06/10:����10813
                    mstrTime_Start = GetBillInfo(15, mstr���ݺ�, False, True)
                Else
                    '���˺�:2007/06/10:����10813
                    mstrTime_Start = GetBillInfo(15, mstr���ݺ�)
                End If
            End If
            
            '���˺�:2007/06/10:����10813
            If mbln��Ҫ�˲� Then
                mstrTime_End = GetBillInfo(15, txtNO.Tag, False, True)
            Else
                mstrTime_End = GetBillInfo(15, txtNO.Tag)
            End If
            
            If mstrTime_End = "" Then
                gcnOracle.RollbackTrans
                MsgBox "ע��:" & vbCrLf & "  �õ����Ѿ�����������Աɾ��,���ܼ�����", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If mstrTime_End <> mstrTime_Start Then
                If MsgBox("ע��:" & vbCrLf & "  �õ����Ѿ�����������Ա�༭�����ܼ���!" & vbCrLf & "  �Ƿ�����ˢ�µ���?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    gcnOracle.RollbackTrans
                    mshBill.ClearBill
                    Call initCard
                    Exit Sub
                Else
                    gcnOracle.RollbackTrans: Exit Sub
                End If
            End If
            
            If SaveCheck = True Then
                Dim blnTemp As Boolean
                strReg = IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
                
                If chkת���ƿ�.Value = 1 Then
                    If cboType.ItemData(cboType.ListIndex) <> 0 Then
                        '��������
                        gcnOracle.CommitTrans
                        blnTemp = False
                        If mbln�˻� = False Then
                            frmDrawCard.ShowCard Me, txtNO.Text, 7, , , blnSuccess, Val(txtDraw.Tag), Trim(txtDrawPerson.Text)
                        End If
                    Else
                        If Check�ƿ�(blnTemp) = False Then
                            If blnTemp = True Then gcnOracle.RollbackTrans: Exit Sub
                            gcnOracle.CommitTrans
                            blnTemp = False
                        Else
                            gcnOracle.CommitTrans
                            blnTemp = True
                        End If
                    End If
                Else
                    gcnOracle.CommitTrans
                    blnTemp = False
                End If
                
                If Val(strReg) = 1 Then
                    '��ӡ
                    If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                        printbill
                    End If
                End If
                If blnTemp Then
                    frmTransferCard.ShowCard Me, txtNO.Text, 11, , , blnSuccess
                End If
                Unload Me
                Exit Sub
            Else
                gcnOracle.RollbackTrans
            End If
        
        Else
            If Not �����_�˻� Then Exit Sub
            If SaveCheck = True Then
                strReg = IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
                If Val(strReg) = 1 Then
                    '��ӡ
                    If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                        printbill
                    End If
                End If
                Unload Me
                Exit Sub
            End If
        End If
        blnTrans = False
        mblnUpdate = False
        mblnCheckPrice = False
        Exit Sub
    End If
            
    If mint�༭״̬ = 5 Then      '�޸ķ�Ʊ��Ϣ
        If SaveRecipe = True Then
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 10 Then      '�޸�ע��֤��
        If SaveRegist = True Then
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 6 Then
        If SaveStrike = True Then
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 8 Then    '�˻�
        If ValidData = False Then Exit Sub
        If SaveRestore Then
            strReg = IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
            If Val(strReg) = 1 Then
                '��ӡ
                If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                    printbill
                End If
            End If
            Unload Me
            Exit Sub
        End If
    End If
            
    If ValidData = False Then Exit Sub
    If Not CheckProvider Then Exit Sub

    blnSuccess = SaveCard
        
    If blnSuccess = True Then
        '��ձ������ݼ�
        If mrsCostlyInfo.RecordCount > 0 Then
            mrsCostlyInfo.MoveFirst
            Do While Not mrsCostlyInfo.EOF
                mrsCostlyInfo.Delete
                mrsCostlyInfo.MoveNext
            Loop
        End If
        strReg = IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
        If Val(strReg) = 1 Then
            '��ӡ
            If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                printbill
            End If
        End If
        If mint�༭״̬ = 2 Then   '�޸�
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
'    If mbln�������� Then
'        mstr���ݺ� = NextNo(68)
'    End If
    txtNO = ""
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    
    Call RefreshRowNO(mshBill, mCol�к�, 1)
    
    SetEdit
    
    txtProvider.Text = ""
    txtProvider.Tag = "0"
    txtժҪ.Text = ""
    txtProvider.SetFocus
    mblnChange = False
    vsfCostlyInfo.Visible = False: Call Form_Resize
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNO.Tag
    Exit Sub
ErrHand:
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function ProduceDateCheck(ByVal strDate As String) As Boolean
    '���ܣ��������ڣ�ע��֤Ч�ڼ��
    'strdate ��������
    '����ֵ��true-���ͨ����false-���δͨ��
    If mintProduceDate = 1 Then
        With mshBill
            If .TextMatrix(.Row, mcolע��֤��Ч��) = "" Then
                ProduceDateCheck = True 'ע��֤��Ч��Ϊ���򲻼��
                Exit Function
            Else
                If CDate(strDate) > CDate(.TextMatrix(.Row, mcolע��֤��Ч��)) Then
                    If mintCheckType = 1 Then
                        If MsgBox("�������ڴ���ע��֤��Ч�ڣ�������Ϊ��֤�������ģ��Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                            ProduceDateCheck = True
                        Else
                            ProduceDateCheck = False
                        End If
                    ElseIf mintCheckType = 2 Then
                        MsgBox "�������ڴ���ע��֤��Ч�ڣ�������Ϊ��֤�������ģ�", vbInformation, gstrSysName
                        ProduceDateCheck = False
                    Else
                        ProduceDateCheck = True
                    End If
                Else
                    ProduceDateCheck = True
                End If
            End If
            
        End With
    Else
        ProduceDateCheck = True
    End If
End Function

Private Sub Form_Load()
    Dim strReg As String
    Dim strCheck As String
    
    strCheck = zlDatabase.GetPara("����У��", glngSys, mlngModule, 0)
    If InStr(1, strCheck, "|") > 0 Then
        'У�鷽ʽ��0-����飻1�����ѣ�2����ֹ
        mintCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))
    End If
    mintProduceDate = Val(zlDatabase.GetPara("��������Ч�ڼ��", glngSys, mlngModule, "0"))
    
    Me.lblType.Caption = "����ID��": Me.lblType.Tag = 1
    
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    strReg = Val(zlDatabase.GetPara("��˲�������", glngSys, mlngModule, "0"))
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    mblnFirst = True
    With cboType
        .AddItem "�ƿ⵽"
        .ItemData(.NewIndex) = 0
        If Val(strReg) = 0 Then .ListIndex = 0
        .AddItem "���õ�"
        .ItemData(.NewIndex) = 1
        If Val(strReg) = 1 Then .ListIndex = 1
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    Call chkת���ƿ�_Click
    
    mbln���б굥λ��� = IIf(Val(zlDatabase.GetPara("�б����Ŀ�ѡ����б굥λ���", glngSys, mlngModule, "0")) = 1, 1, 0) = 1
   
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
        .FM_ɢװ���ۼ� = GetFmtString(2, g_�ۼ�)
    End With
    
    '���˺�:���Ӻ˲鹦��,2007/05/13
    '�⹺��⣬��Ҫȷ���Ƿ���Ҫ�˲鹦��
    mbln��Ҫ�˲� = Val(zlDatabase.GetPara("�����⹺��Ҫ�˲�", glngSys, 0)) = 1
     
    lbl�˲���.Visible = mbln��Ҫ�˲�
    lbl�˲�����.Visible = mbln��Ҫ�˲�
    txt�˲���.Visible = mbln��Ҫ�˲�
    txt�˲�����.Visible = mbln��Ҫ�˲�
    
    mintBatchNoLen = GetBatchNoLen()
    mint��Ʊ��Len = Get��Ʊ��Len
    
    mbln�Ӽ��� = Get�Ӽ���()
    mbln�ֶμӳ��� = IS�ֶμӳ���()
    mblnʱ�۹�ǰ���� = ISCHECK�⹺��ǰ����()
    mbln��ǿ�ƿ���ָ���۸� = ISCHECK��ǿ�ƿ���ָ���۸�()
    mblnʱ������ֱ��ȷ���ۼ� = isʱ������ֱ��ȷ���ۼ�()
    mblnʱ������ȡ�ϴ��ۼ� = isʱ������ȡ�ϴ��ۼ�()
    
    mbln�����������Ų��ؿ��� = Val(zlDatabase.GetPara(305, glngSys, 0)) = 1
    
    mbln�˻� = False
    
    txtNO = mstr���ݺ�
    txtNO.Tag = txtNO.Text
    '�ָ����Ի���������
    RestoreWinState Me, App.ProductName, mstrCaption
    
    Call initCard
    If mint�༭״̬ <> 6 Then
        If mshBill.ColWidth(mCol��������) > 0 Then
            mshBill.ColWidth(mCol��������) = 0
        End If
    Else
        If mshBill.ColWidth(mCol��������) = 0 Then
            mshBill.ColWidth(mCol��������) = 800
        End If
    End If
    Call RestoreBILLWidthSet
    mblnUpdate = False
    mblnChange = False
    mblnCheckPrice = False
    mshBill.ColWidth(mCol���) = 0
    '�ָ����Ի��������ú󣬻���Ҫ��Ȩ�޿��Ƶ��н�һ������
    With mshBill
        .ColWidth(mCol�����) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mCol�ɹ���) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mCol������) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mCol���) = IIf(mblnCostView = False, 0, 900)
    End With
End Sub

Private Sub initCard()
    '------------------------------------------------------------------------------------------------
    '����:��ʼ����Ƭ��Ϣ
    '------------------------------------------------------------------------------------------------
    Dim rstemp As New Recordset
    Dim strUnit As String, strUnitQuantity As String, str���� As String
    Dim num��װϵ�� As String, strOrder As String, strCompare As String, strReg As String
    Dim dblSum As Double, intRow As Integer, i As Integer
    Dim varStuff As Variant
    Dim lngProviderID As Long
    Dim strDateBegin As String, strDateEnd As String, strIVNO As String
    Dim dtIVDate As Date
    Dim rs As ADODB.Recordset
    Dim dbl�ɹ��� As Double
     
    strReg = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strOrder = IIf(strReg = "", "00", strReg)
    On Error GoTo ErrHandle
    '�ⷿ
    strCompare = Mid(strOrder, 1, 1)
    
    '�����˻�
    If mint�༭״̬ = 8 Then
        cmdExtractData.Visible = True
    Else
        cmdExtractData.Visible = False
        lblCode.Left = lblCode.Left - cmdExtractData.Width - 100
        txtCode.Left = txtCode.Left - cmdExtractData.Width - 100
        cmdFind.Left = cmdFind.Left - cmdExtractData.Width - 100
    End If
    
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
            Set rstemp = ReturnSQL(cboStock.ItemData(cboStock.ListIndex), mstrCaption, True)
            Do While Not rstemp.EOF
                .AddItem rstemp.Fields(2)
                .ItemData(.NewIndex) = rstemp.Fields(0)
                rstemp.MoveNext
            Loop
    
            If .ListCount > 0 Then
                .ListIndex = 0
            End If
                                
        End With
    End If
    
    Select Case mint�༭״̬
        Case 1, 8           '�������˻�
        
            Txt������ = UserInfo.�û���
            Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4, 5, 6, 7, 9, 10
                
            initGrid
            
            If mint�༭״̬ = 4 Then
                gstrSQL = "" & _
                    "   Select b.id,b.���� " & _
                    "   From ҩƷ�շ���¼ a,���ű� b " & _
                    "   where a.�ⷿid=b.id and A.���� = 15 and a.no=[1] "
                
                Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�)
                If rstemp.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                With cboStock
                    .AddItem rstemp!����
                    .ItemData(.NewIndex) = rstemp!Id
                    .ListIndex = 0
                    mintcboIndex = 0
                End With
                rstemp.Close
            End If
            
             
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "D.���㵥λ AS ��λ,D.���㵥λ  as ɢװ��λ, decode(nvl(A.��ҩ��ʽ,0),1,-1,1)* A.��д����  AS ����,1 as ����ϵ��,"
                    num��װϵ�� = "1"
                Case 1
                    strUnitQuantity = "B.��װ��λ AS ��λ,D.���㵥λ  as  ɢװ��λ, decode(nvl(A.��ҩ��ʽ,0),1,-1,1)* (A.��д���� / B.����ϵ��) AS ����,B.����ϵ�� ,"
                    num��װϵ�� = "B.����ϵ��"
            End Select
            
            
            Select Case mint�༭״̬
            Case 5
                '�޸ķ�Ʊ��Ϣ
                If mint��¼״̬ = 1 Then GoTo Go����:
                    gstrSQL = "" & _
                        "   Select * " & _
                        "   From (  SELECT distinct a.ҩƷid as ����id,A.���,('[' || D.���� || ']' || D.����) AS ������Ϣ,zlSpellCode(D.����) ����,E.���� ��Ʒ��,D.���,D.���� as ԭ����,A.����,a.��׼�ĺ�, A.����,A.����,to_char(A.��������,'yyyy-mm-dd') ��������," & _
                        "                   b.���Ч��,A.Ч��,b.һ���Բ���,Nvl(b.�Ƿ��������,0) as �������,b.�ⷿ����,b.���Ч��,A.�������,a.���Ч�� as ���ʧЧ��,Nvl(b.�ӳ���,0)/100 as �ӳ���," & strUnitQuantity & _
                        "                   b.ָ��������*" & num��װϵ�� & " as ָ�������� ,a.�ɱ���*" & num��װϵ�� & " AS �����, decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*a.�ɱ���� as ������,b.ָ�������/100 as ָ�������,d.�Ƿ���,b.���÷���,nvl(a.��ҩ��ʽ,0) �˻�," & _
                        "                   DECODE(A.����, NULL, 0, A.����) AS ����, A.���ۼ� ,a.���۽��,a.���,a.���۲��," & _
                        "                   a.�������,a.��Ʊ�� ,a.��Ʊ����, a.��Ʊ����,a.��Ʊ���,a.��ҩ��λid,f.���� as ��Ӧ��,a.ע��֤��,a.��Ʒ����,a.������,a.��������," & _
                        "                   a.�����,a.�������,a.�ⷿid,g.���� as ����,nvl(a.�������,0) as �������,A.�˲���,A.�˲�����,a.�ڲ�����,a.����ID,b.ע��֤��Ч�� " & _
                        "           FROM (  select min(X.id) as id,max(Nvl(x.��ҩ��ʽ,0)) ��ҩ��ʽ, " & _
                        "                           sum(x.ʵ������) as ��д����,sum(x.�ɱ����) as �ɱ���� ,sum(x.���۽��) as ���۽��,sum(x.���) as ���,sum(to_number(nvl(to_char(x.�÷�," & gOraFmt_Max.FM_��� & " ),0), " & gOraFmt_Max.FM_��� & ")) as ���۲��," & _
                        "                           y.�������,y.��Ʊ��,y.��Ʊ����,y.��Ʊ����,sum(y.��Ʊ���) as ��Ʊ���," & _
                        "                           X.ҩƷID,X.���,X.����,X.��׼�ĺ�, X.����,NVL(X.����,0) ����,X.��������,X.Ч��,X.���Ч��,X.�������,X.����,X.�ɱ���,X.���ۼ�," & _
                        "                           x.��ҩ��λID,X.ע��֤��,X.��Ʒ����,x.�ⷿID,max(x.������) as ������,max(x.��������) as ��������,max(x.�����) as �����," & _
                        "                           max(x.�������) as �������,max(x.��ҩ��) as �˲���,x.�ڲ�����,x.����ID," & _
                        "                           max(x.��ҩ����) as �˲�����,Nvl(Y.�������,0) as ������� " & _
                        "                   From ҩƷ�շ���¼ x,Ӧ����¼ y " & _
                        "                   WHERE x.id=y.�շ�id(+)  and y.ϵͳ��ʶ(+)=5 and y.��¼����(+)=0 and X.NO=[1] AND ����=15  " & _
                        "                   group by X.ҩƷID,X.���,X.����,X.��׼�ĺ�,X.����,NVL(X.����,0)  ,x.��������,X.Ч��,X.���Ч��,X.�������,X.����,X.�ɱ���,X.���ۼ�," & _
                        "                            x.��ҩ��λID,X.ע��֤��,X.��Ʒ����,X.�ⷿID,X.�ڲ�����,X.����ID,y.�������,y.��Ʊ��,y.��Ʊ����,y.��Ʊ����,NVL(Y.�������,0) " & _
                        "                   having sum(ʵ������)<>0 ) A," & _
                        "                   �������� B,�շ���ĿĿ¼ D,��Ӧ�� f,���ű� g,�շ���Ŀ���� e  " & _
                        "           Where A.ҩƷid = B.����id and a.ҩƷid=D.id and a.��ҩ��λid=f.id and a.�ⷿid=g.id And d.id  = e.�շ�ϸĿid(+) And e.����(+) = 3  " & _
                        "          ) " & _
                        "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            Case 6
                '����
                gstrSQL = "" & _
                    "   Select * " & _
                    "   From (  SELECT distinct a.ҩƷid as ����id,A.���,('[' || D.���� || ']' || D.����) AS ������Ϣ,zlSpellCode(D.����) ����,e.���� ��Ʒ��,D.���,D.���� as ԭ����,A.����,a.��׼�ĺ�, A.����,to_char(A.��������,'yyyy-mm-dd') ��������," & _
                    "                   b.���Ч��,A.Ч��,b.һ���Բ���,nvl(b.�Ƿ��������,0) as �������,b.�ⷿ����,b.���Ч��,A.�������,a.���Ч�� as ���ʧЧ��,Nvl(b.�ӳ���,0)/100 as �ӳ���," & strUnitQuantity & _
                    "                   b.ָ��������*" & num��װϵ�� & " as ָ�������� ,a.�ɱ���*" & num��װϵ�� & " AS �����, " & _
                    "                   decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*a.�ɱ���� as ������,b.ָ�������/100 as ָ�������,d.�Ƿ���,b.���÷���,nvl(a.��ҩ��ʽ,0) �˻�," & _
                    "                   DECODE(A.����, NULL, 0, A.����) AS ����,A.���ۼ�,decode(nvl(A.��ҩ��ʽ,0),1,-1,1)* nvl(A.���۽��,0) as ���۽��,decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*nvl(A.���,0)  as ���,decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*nvl(A.���۲��,0) as ���۲��, " & _
                    "                   a.�������,a.��Ʊ�� ,a.��Ʊ����, a.��Ʊ����,0 as ��Ʊ���,a.��ҩ��λid,f.���� as ��Ӧ��,a.ע��֤��,a.��Ʒ����,a.�ⷿid,g.���� as ����," & _
                    "                   nvl(a.�������,0) as �������,A.�˲���,A.�˲�����,a.�ڲ�����,a.����ID,b.ע��֤��Ч�� " & _
                    "           FROM (  select min(X.id) as id,max(Nvl(x.��ҩ��ʽ,0)) ��ҩ��ʽ, sum(��д����) as ��д����,sum(�ɱ����) as �ɱ����," & _
                    "                           y.�������,y.��Ʊ��,y.��Ʊ����,y.��Ʊ����,X.ҩƷID,X.���,X.����,X.��׼�ĺ�, X.����,X.��������,X.Ч��,X.���Ч��,X.�������,X.����,X.�ɱ���,X.���ۼ�," & _
                    "                           sum(x.���۽��) as ���۽��,sum(x.���) as ���,sum(to_number(nvl(to_char(x.�÷�," & gOraFmt_Max.FM_��� & " ),0), " & gOraFmt_Max.FM_��� & ")) as ���۲��," & _
                    "                           x.��ҩ��λID,X.ע��֤��,X.��Ʒ����,x.�ⷿID,max(x.��ҩ��) as �˲���,max(x.��ҩ����) as �˲�����," & _
                    "                           x.�ڲ�����, x.����ID, Nvl(Y.�������,0) as ������� " & _
                    "                   From ҩƷ�շ���¼ x,(Select Id, ��¼����, ��¼״̬, No, ��Ŀid, ���, �շ�id, ��λid, Ʒ��, ���, ����, ����, ������λ, ��ⵥ�ݺ�, ���ݽ��, ����, �ɹ���, �ɹ����, �������, ��Ʊ��,��Ʊ����, ��Ʊ����, ��Ʊ���, �ƶ�����, �ƻ����, �ƻ���, �ƻ�����, ������, ��������, �����, �������, �������, �ƻ����, ϵͳ��ʶ From Ӧ����¼ Where ϵͳ��ʶ=5 And ��¼����=0) y " & _
                    "                   WHERE x.id=y.�շ�id(+) and X.NO=[1] AND ����=15  " & _
                    "                   group by X.ҩƷID,X.���,X.����,X.��׼�ĺ�,X.����,x.��������,X.Ч��,X.���Ч��,X.�������,X.����,X.�ɱ���,X.���ۼ�," & _
                    "                            x.��ҩ��λID,X.ע��֤��,X.��Ʒ����,X.�ⷿID,x.�ڲ�����,x.����ID,y.�������,y.��Ʊ��,y.��Ʊ����,y.��Ʊ����,NVL(Y.�������,0) " & _
                    "                   having sum(��д����)<>0 ) A," & _
                    "                   �������� B,�շ���ĿĿ¼ D,��Ӧ�� f,���ű� g,�շ���Ŀ���� e " & _
                    "           Where A.ҩƷid = B.����id and a.ҩƷid=D.id and a.��ҩ��λid=f.id and a.�ⷿid=g.id And d.id  = e.�շ�ϸĿid(+) And e.����(+) = 3  " & _
                    "          ) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            Case 10
                '�޸�ע��֤��
                gstrSQL = "" & _
                    "   Select * " & _
                    "   From (  SELECT distinct a.ҩƷid ����id,A.���,b.һ���Բ���,nvl(b.�Ƿ��������,0) as �������,b.�ⷿ����,('[' || D.���� || ']' || D.����) AS ������Ϣ,zlSpellCode(D.����) ����,e.���� ��Ʒ��, D.���,D.���� as ԭ����,A.����,A.���ս���,A.��׼�ĺ�, A.����,Nvl(A.����,0) ����,to_char(A.��������,'yyyy-mm-dd') ��������," & _
                    "                   b.���Ч��,A.Ч��,b.���Ч��,A.�������,a.���Ч�� as ���ʧЧ��,Nvl(b.�ӳ���,0)/100 as �ӳ���," & strUnitQuantity & _
                    "                   b.ָ��������*" & num��װϵ�� & " as ָ�������� ,A.�ɱ���*" & num��װϵ�� & " AS �����, decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*A.�ɱ���� AS ������,Nvl(A.��ҩ��ʽ,0) �˻�,b.ָ�������/100 as ָ�������,D.�Ƿ���,b.���÷���," & _
                    "                   DECODE(A.����, NULL, 0, A.����) AS ����, A.���ۼ�, decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*A.���۽�� ���۽��, " & _
                    "                   decode(nvl(A.��ҩ��ʽ,0),1,-1,1)* A.��� ���,  decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*to_number(nvl(to_char(A.�÷�," & gOraFmt_Max.FM_��� & " ),0), " & gOraFmt_Max.FM_��� & ")   as ���۲��, " & _
                    "                   C.�������,C.��Ʊ�� ,c.��Ʊ����, C.��Ʊ����, decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*C.��Ʊ��� ��Ʊ���,a.��ҩ��λid,f.���� as ��Ӧ��,a.ע��֤��,a.��Ʒ����, a.ժҪ,A.������,A.��������,A.��ҩ�� as �˲���,A.��ҩ���� as �˲�����,A.�����,A.�������," & _
                    "                   a.�ⷿid,g.���� as ����,nvl(c.�������,0) as �������, a.�ڲ�����, a.����id,b.ע��֤��Ч�� " & _
                    "           FROM ҩƷ�շ���¼ A, �������� B,�շ���ĿĿ¼ D, (Select Id, ��¼����, ��¼״̬, No, ��Ŀid, ���, �շ�id, ��λid, Ʒ��, ���, ����, ����, ������λ, ��ⵥ�ݺ�, ���ݽ��, ����, �ɹ���, �ɹ����, �������, ��Ʊ��,��Ʊ����, ��Ʊ����, ��Ʊ���, �ƶ�����, �ƻ����, �ƻ���, �ƻ�����, ������, ��������, �����, �������, ժҪ, �������, �ƻ����, ϵͳ��ʶ From Ӧ����¼ Where ϵͳ��ʶ=5 And ��¼����=0) C,��Ӧ�� f,���ű� g,�շ���Ŀ���� e  " & _
                    "           Where A.ҩƷid = B.����id  and a.ҩƷid=d.ID AND A.Id = C.�շ�id (+)   and a.��ҩ��λid=f.id and a.�ⷿid=g.id And d.id  = e.�շ�ϸĿid(+) And e.����(+) = 3  " & _
                    "                   AND (A.��¼״̬ = 1 Or Mod(A.��¼״̬, 3) = 0) And A.������� Is Not Null " & _
                    "                   AND A.���� = 15 AND A.No = [1] " & _
                    "       ) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            Case Else
            
Go����:
                gstrSQL = "" & _
                    "   Select * " & _
                    "   From (  SELECT distinct a.ҩƷid ����id,A.���,b.һ���Բ���,nvl(b.�Ƿ��������,0) as �������,b.�ⷿ����,('[' || D.���� || ']' || D.����) AS ������Ϣ,zlSpellCode(D.����) ����,E.���� ��Ʒ��,D.���,D.���� as ԭ����,A.����,A.���ս���,A.��׼�ĺ�, A.����,Nvl(A.����,0) ����,to_char(A.��������,'yyyy-mm-dd') ��������," & _
                    "               b.���Ч��,A.Ч��,b.���Ч��,A.�������,a.���Ч�� as ���ʧЧ��,Nvl(b.�ӳ���,0)/100 as �ӳ���," & strUnitQuantity & _
                    "               b.ָ��������*" & num��װϵ�� & " as ָ�������� ,A.�ɱ���*" & num��װϵ�� & " AS �����, decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*A.�ɱ���� AS ������,Nvl(A.��ҩ��ʽ,0) �˻�,b.ָ�������/100 as ָ�������,D.�Ƿ���,b.���÷���," & _
                    "               DECODE(A.����, NULL, 0, A.����) AS ����, A.���ۼ�, decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*A.���۽�� ���۽��, " & _
                    "               decode(nvl(A.��ҩ��ʽ,0),1,-1,1)* A.��� ���,  decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*to_number(nvl(to_char(A.�÷�," & gOraFmt_Max.FM_��� & " ),0), " & gOraFmt_Max.FM_��� & ")   as ���۲��, " & _
                    "               C.�������,C.��Ʊ�� ,c.��Ʊ����, C.��Ʊ����, decode(nvl(A.��ҩ��ʽ,0),1,-1,1)*C.��Ʊ��� ��Ʊ���,a.��ҩ��λid,f.���� as ��Ӧ��,a.ע��֤��,a.��Ʒ����, a.ժҪ,A.������,A.��������,A.��ҩ�� as �˲���,A.��ҩ���� as �˲�����,A.�����,A.�������," & _
                    "               a.�ⷿid,g.���� as ����,nvl(c.�������,0) as �������, a.�ڲ�����, a.����id,b.ע��֤��Ч�� " & _
                    "           FROM ҩƷ�շ���¼ A, �������� B,�շ���ĿĿ¼ D, (Select Id, ��¼����, ��¼״̬, No, ��Ŀid, ���, �շ�id, ��λid, Ʒ��, ���, ����, ����, ������λ, ��ⵥ�ݺ�, ���ݽ��, ����, �ɹ���, �ɹ����, �������, ��Ʊ��,��Ʊ����, ��Ʊ����, ��Ʊ���, �ƶ�����, �ƻ����, �ƻ���, �ƻ�����, ������, ��������, �����, �������, ժҪ, �������, �ƻ����, ϵͳ��ʶ From Ӧ����¼ Where ϵͳ��ʶ=5 And ��¼����=0) C,��Ӧ�� f,���ű� g,�շ���Ŀ���� e " & _
                    "           Where A.ҩƷid = B.����id  and a.ҩƷid=d.ID AND A.Id = C.�շ�id (+)   and a.��ҩ��λid=f.id and a.�ⷿid=g.id And d.id  = e.�շ�ϸĿid(+) And e.����(+) = 3 " & _
                    "                   AND A.��¼״̬ =[2] " & _
                    "                   AND A.���� = 15 AND A.No = [1] " & _
                    "       ) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            End Select
            
            
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�, mint��¼״̬, _
                            cboStock.ItemData(cboStock.ListIndex), _
                            lngProviderID, strDateBegin, strDateEnd)
            
            If rstemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            If mint�༭״̬ = 9 Then
                '���˺�:2007/06/10:����10813
                mstrTime_Start = GetBillInfo(15, mstr���ݺ�)
            ElseIf mint�༭״̬ = 3 And mbln��Ҫ�˲� Then
                '���˺�:2007/06/10:����10813
                mstrTime_Start = GetBillInfo(15, mstr���ݺ�, False, True)
            Else
                '���˺�:2007/06/10:����10813
                mstrTime_Start = GetBillInfo(15, mstr���ݺ�)
            End If
            Select Case mint�༭״̬
                Case 2, 6
                    Txt������ = UserInfo.�û���
                    Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    If mint�༭״̬ = 2 Then
                        Txt����� = ""
                        Txt������� = ""
                        txt�˲��� = ""
                        txt�˲����� = ""
                    Else
                        Txt����� = UserInfo.�û���
                        Txt������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                        txt�˲��� = UserInfo.�û���
                        txt�˲����� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    End If
                Case 9
                    Txt������ = zlStr.Nvl(rstemp!������)
                    Txt�������� = IIf(zlStr.Nvl(rstemp!��������) = "", "", Format(rstemp!��������, "yyyy-mm-dd hh:mm:ss"))
                    txt�˲��� = UserInfo.�û���
                    txt�˲����� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    Txt����� = ""
                    Txt������� = ""
                Case Else
                    Txt������ = zlStr.Nvl(rstemp!������)
                    Txt�������� = Format(rstemp!��������, "yyyy-mm-dd hh:mm:ss")
                    txt�˲��� = zlStr.Nvl(rstemp!�˲���) 'UserInfo.�û�����
                    txt�˲����� = IIf(IsNull(rstemp!�˲�����), "", Format(rstemp!�˲�����, "yyyy-mm-dd hh:mm:ss"))
                    Txt����� = IIf(IsNull(rstemp!�����), "", rstemp!�����)
                    Txt������� = IIf(IsNull(rstemp!�������), "", Format(rstemp!�������, "yyyy-mm-dd hh:mm:ss"))
            End Select
            txtProvider.Text = rstemp!��Ӧ��
            txtProvider.Tag = rstemp!��ҩ��λID
            mbln�˻� = (rstemp!�˻� = 1)
'            txtժҪ.Text = IIf(IsNull(rsTemp!ժҪ), "", rsTemp!ժҪ)
            If mint�༭״̬ = 5 Or mint�༭״̬ = 6 Then
                txtժҪ.Text = GetժҪ(mstr���ݺ�, mint�༭״̬)
            Else
                txtժҪ.Text = IIf(IsNull(rstemp!ժҪ), "", rstemp!ժҪ)
            End If
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            If mint�༭״̬ = 5 Or mint�༭״̬ = 7 Then
                If rstemp!������� <> 0 Then
                    mintParallelRecord = IIf(mint�༭״̬ = 5, 4, 5)        '�ѱ������˸���
                    Exit Sub
                ElseIf mint�༭״̬ = 7 Then
                    '����Ƿ���ڲ��ָ�������
                    gstrSQL = "Select Nvl(Max(�������), 0) ������� From Ӧ����¼ " & _
                        " where �շ�id=(Select Id From ҩƷ�շ���¼ Where ����=15 And No=[1] And (Mod(��¼״̬,3)=0 Or ��¼״̬=1) " & _
                        " And ���=[2]) "
                    strOrder = rstemp!���
                    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[ȡ�������]", txtNO.Text, strOrder)
                    
                    If rs!������� <> 0 Then
                        mintParallelRecord = 5
                    End If
                End If
            End If
            
            If mbln�˻� Then LblTitle.Caption = "���������˻���"
            intRow = 0
            If mbln�˻� Or mint�༭״̬ = 3 Then
                Set mCllBillData = New Collection
            End If
                        
            With mshBill
                Do While Not rstemp.EOF
                    intRow = intRow + 1
                    .Rows = .Rows + 1
                    
                    .TextMatrix(intRow, 0) = rstemp.Fields(0)
                    .TextMatrix(intRow, mCol����) = rstemp!������Ϣ
                    .TextMatrix(intRow, mCol��Ʒ��) = IIf(IsNull(rstemp!��Ʒ��), "", rstemp!��Ʒ��)
                    .TextMatrix(intRow, mCol���) = IIf(IsNull(rstemp!���), "", rstemp!���)
                    .TextMatrix(intRow, mCol����) = IIf(IsNull(rstemp!����), "", rstemp!����)
                    .TextMatrix(intRow, mCol��λ) = rstemp!��λ
                    .TextMatrix(intRow, mCol����) = IIf(IsNull(rstemp!����), "", rstemp!����)
                    .TextMatrix(intRow, mcol��׼�ĺ�) = IIf(IsNull(rstemp!��׼�ĺ�), "", rstemp!��׼�ĺ�)
                    .TextMatrix(intRow, mColЧ��) = IIf(IsNull(rstemp!Ч��), "", rstemp!Ч��)
                    .TextMatrix(intRow, mCol����) = Format(rstemp!����, mFMT.FM_����)
                    .TextMatrix(intRow, mCol�����) = Format(rstemp!�����, mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mCol�ɹ���) = Format(Val(.TextMatrix(intRow, mCol�����)) * 100 / IIf(Val(zlStr.Nvl(rstemp!����)) = 0, 1, Val(zlStr.Nvl(rstemp!����))), mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mCol������) = Format(IIf(mint�༭״̬ = 6, 0, rstemp!������), mFMT.FM_���)
                    
                    '���˺�:���ۼ۴���:���ۼ�-->���ۼ�;���۽��-->���۽��;���-->���۲��;��;-->�ⷿ��λ���
                    ' ���۽�������������ۼۣ�
                    ' ���۲��"���ۼ۽����۽�������ⵥλ����Ľ��Ͱ����۵�λ����Ľ��Ĳ�ֵ��
                    .TextMatrix(intRow, mcol���ۼ�) = Format(Val(zlStr.Nvl(rstemp!���ۼ�)), mFMT.FM_ɢװ���ۼ�)          'If Val(.TextMatrix(.row, mcol���ۼ�)) = 0 Then
                    .TextMatrix(intRow, mCol�ۼ�) = Format(Val(zlStr.Nvl(rstemp!���ۼ�)) * Val(zlStr.Nvl(rstemp!����ϵ��)), mFMT.FM_���ۼ�)
                    
                    '�����ۼ�
'                    .TextMatrix(intRow, mCol�ۼ�) = Format((Val(NVL(rsTemp!���۽��)) - Val(NVL(rsTemp!���۲��))) / Val(NVL(rsTemp!����)), mFMT.FM_���ۼ�)
                    .TextMatrix(intRow, mcol���۵�λ) = zlStr.Nvl(rstemp!ɢװ��λ)
                    
                    If mint�༭״̬ = 6 Then
                        '����û����صĲ��
                        .TextMatrix(intRow, mcol���۲��) = ""
                        .TextMatrix(intRow, mcol���۽��) = ""
                        .TextMatrix(intRow, mCol���) = ""
                        .TextMatrix(intRow, mCol�ۼ۽��) = ""
                    Else
                        .TextMatrix(intRow, mcol���۲��) = Format(Val(zlStr.Nvl(rstemp!���)), mFMT.FM_���)
                        .TextMatrix(intRow, mcol���۽��) = Format(Val(zlStr.Nvl(rstemp!���۽��)), mFMT.FM_���)
                        '�����ۼۼ��ۼ۽��
'                        .TextMatrix(intRow, mCol���) = Format(Val(NVL(rsTemp!���)) - Val(NVL(rsTemp!���۲��)), mFMT.FM_���)
'                        .TextMatrix(intRow, mCol�ۼ۽��) = Format(Val(NVL(rsTemp!���۽��)) - Val(NVL(rsTemp!���۲��)), mFMT.FM_���)
                        .TextMatrix(intRow, mCol���) = Format(Val(zlStr.Nvl(rstemp!���)), mFMT.FM_���)
                        .TextMatrix(intRow, mCol�ۼ۽��) = Format(Val(zlStr.Nvl(rstemp!���۽��)), mFMT.FM_���)
                    End If
                    
                    .TextMatrix(intRow, mCol����) = rstemp!����
                    .TextMatrix(intRow, mCol�������) = IIf(IsNull(rstemp!�������), "", rstemp!�������)
                    
                    If mint�༭״̬ <> 6 Then '��������ʾ���ս���
                        If (mint�༭״̬ = 5 And mint��¼״̬ <> 1) Then '�޸ķ�Ʊ��Ϣ�����߼�¼����ʾ���ս���
                        Else
                            .TextMatrix(intRow, mCol���ս���) = IIf(IsNull(rstemp!���ս���), "", rstemp!���ս���)
                        End If
                    End If
                    
                    .TextMatrix(intRow, mCol��Ʊ��) = IIf(IsNull(rstemp!��Ʊ��), "", rstemp!��Ʊ��)
                    .TextMatrix(intRow, mcol��Ʊ����) = IIf(IsNull(rstemp!��Ʊ����), "", rstemp!��Ʊ����)
                    .TextMatrix(intRow, mCol��Ʊ����) = IIf(IsNull(rstemp!��Ʊ����), "", rstemp!��Ʊ����)
                    .TextMatrix(intRow, mCol��Ʊ���) = IIf(Format(IIf(IsNull(rstemp!��Ʊ���), "0", rstemp!��Ʊ���), mFMT.FM_���) = "0.00", "", Format(IIf(IsNull(rstemp!��Ʊ���), "0", rstemp!��Ʊ���), mFMT.FM_���))
                    
                    .TextMatrix(intRow, mcolһ���Բ���) = zlStr.Nvl(rstemp!һ���Բ���)
                    .TextMatrix(intRow, mcol�������) = zlStr.Nvl(rstemp!�������)
                    .TextMatrix(intRow, mcol���Ч��) = zlStr.Nvl(rstemp!���Ч��)
                    .TextMatrix(intRow, mcol�������) = zlStr.Nvl(rstemp!�������)
                    .TextMatrix(intRow, mcol���ʧЧ��) = zlStr.Nvl(rstemp!���ʧЧ��)
                    .TextMatrix(intRow, mcol��������) = zlStr.Nvl(rstemp!��������)
                    .TextMatrix(intRow, mcolע��֤��Ч��) = IIf(IsNull(rstemp!ע��֤��Ч��), "", Format(rstemp!ע��֤��Ч��, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mColָ��������) = Format(rstemp!ָ��������, mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mColԭ����) = IIf(IsNull(rstemp!ԭ����), "!", rstemp!ԭ����)
                    
                    '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                    .TextMatrix(intRow, mColԭ����) = IIf(IsNull(rstemp!���Ч��), "0", rstemp!���Ч��) & "||" & rstemp!ָ������� & "||" & IIf(IsNull(rstemp!�Ƿ���), 0, rstemp!�Ƿ���) & "||" & IIf(IsNull(rstemp!���÷���), 0, rstemp!���÷���) & "||" & zlStr.Nvl(rstemp!�ⷿ����, 0)
                    .TextMatrix(intRow, mCol����) = ""
                    .TextMatrix(intRow, mCol����ϵ��) = zlStr.Nvl(rstemp!����ϵ��)
                    .TextMatrix(intRow, mcolע��֤��) = zlStr.Nvl(rstemp!ע��֤��)
                    .TextMatrix(intRow, mcol��Ʒ����) = zlStr.Nvl(rstemp!��Ʒ����)
                    .TextMatrix(intRow, mCol���) = zlStr.Nvl(rstemp!���)
                    
                    .TextMatrix(intRow, mcol�ڲ�����) = zlStr.Nvl(rstemp!�ڲ�����)
                    .TextMatrix(intRow, mcol����ID) = zlStr.Nvl(rstemp!����ID)
                    
                    If (mbln�˻� Or mint�༭״̬ = 3) And mint�༭״̬ <> 6 Then
                        dblSum = 0
                        For Each varStuff In mCllBillData
                            If varStuff(0) = CStr(rstemp!����ID & "_" & IIf(IsNull(rstemp!����), "0", rstemp!����)) Then
                                dblSum = varStuff(1)
                                mCllBillData.Remove varStuff(0)
                                Exit For
                            End If
                        Next
                        str���� = rstemp!����ID & "_" & IIf(IsNull(rstemp!����), "0", rstemp!����)
                        dblSum = dblSum + Val(zlStr.Nvl(rstemp!����)) * Val(zlStr.Nvl(rstemp!����ϵ��))
                        mCllBillData.Add Array(str����, dblSum), str����
                    End If
                    
                    If mint�༭״̬ = 6 Then
                        .TextMatrix(intRow, mCol��������) = Format(0, mFMT.FM_����)
                        .RowData(intRow) = rstemp!�������
                        
                        '����Ƿ���ڲ��ָ�������
                        gstrSQL = "Select Nvl(Max(�������), 0) ������� From Ӧ����¼ " & _
                            " where �շ�id=(Select Id From ҩƷ�շ���¼ Where ����=15 And No=[1] And (Mod(��¼״̬,3)=0 Or ��¼״̬=1) " & _
                            " And ���=[2]) "
                        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[ȡ�������]", txtNO.Text, Val(.TextMatrix(intRow, mCol���)))
                        
                        If rs!������� <> 0 Then
                            mintParallelRecord = 6
                        End If
                    Else
                        .TextMatrix(intRow, mCol����) = rstemp!����
                    End If
                    
                    '����ӳ���
                    If mblnʱ�۹�ǰ���� Then
                        dbl�ɹ��� = Val(.TextMatrix(intRow, mCol�ɹ���))
                    Else
                        dbl�ɹ��� = Val(.TextMatrix(intRow, mCol�����))
                    End If
                    
                    If Val(rstemp!�Ƿ���) = 1 Then 'ʱ��
                        If Val(.TextMatrix(intRow, mCol�ۼ�)) <> 0 And dbl�ɹ��� <> 0 Then
                            .TextMatrix(intRow, mcol�ӳ���) = zlStr.FormatEx((Val(.TextMatrix(intRow, mCol�ۼ�)) / dbl�ɹ��� - 1) * 100, 2) & "%"
                        End If
                    Else '����
                        .TextMatrix(intRow, mcol�ӳ���) = zlStr.FormatEx(Val(rstemp!�ӳ���) * 100, 2) & "%"
                    End If
                    
                    rstemp.MoveNext
                Loop
            End With
            rstemp.Close
    End Select
    SetEdit         '���ñ༭����
    Call RefreshRowNO(mshBill, mCol�к�, 1)
    Call ��ʾ�ϼƽ��
    '��ֵ����VsflexGrid�ؼ���ʼ��
    With vsfCostlyInfo
        .Editable = flexEDKbd
        .BackColorBkg = vbWhite
        .AllowUserResizing = flexResizeColumns
        .Rows = 2
        .Cols = 6
        .RowHeight(0) = 300
        '.RowHeight(1) = 245
        .TextMatrix(0, 0) = "SN"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "��������"
        .TextMatrix(0, 3) = "סԺ��"
        .TextMatrix(0, 4) = "����"
        .TextMatrix(0, 5) = "����ID"
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColKey(0) = "SN"
        .ColKey(1) = "����"
        .ColKey(2) = "��������"
        .ColKey(3) = "סԺ��"
        .ColKey(4) = "����"
        .ColKey(5) = "����ID"
        .ColHidden(0) = True
        .ColHidden(5) = True
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 1000
        .Visible = False
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetժҪ(ByVal strNo As String, ByVal int�༭״̬ As Integer) As String
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    Select Case int�༭״̬
        Case 6          '����(ȡ���һ�γ�����ժҪ)
            gstrSQL = "Select ժҪ From ҩƷ�շ���¼ Where ����=15 And No=[1] Order By ������� asc "
        Case 5, 7       '�޸ķ�Ʊ���������
            gstrSQL = "Select ժҪ From ҩƷ�շ���¼ Where ���� = 15 And NO = [1] And (Mod(��¼״̬, 3) = 0 Or ��¼״̬ = 1) order by ������� asc"
    End Select
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡժҪ��Ϣ", strNo)
    
    If Not rstemp.EOF Then
        GetժҪ = zlStr.Nvl(rstemp!ժҪ)
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetEdit()
    Dim intCol As Integer
    Dim intRow As Integer
    
    With mshBill
        If mblnEdit = False Then
            
            cboStock.Enabled = False
            txtProvider.Enabled = False
            cmdProvider.Enabled = False
            txtժҪ.Enabled = False
            
            For intCol = 0 To .Cols - 1
                .ColData(intCol) = 0
            Next
            
            If mint�༭״̬ = 5 Then
                '�޸ķ�Ʊ��Ϣ
                mshBill.ColData(mCol��Ʊ��) = 4
                mshBill.ColData(mcol��Ʊ����) = 4
                mshBill.ColData(mCol��Ʊ����) = 2
                .ColData(mCol��Ʊ���) = 4
  
                txtProvider.Enabled = True
                cmdProvider.Enabled = True
            ElseIf mint�༭״̬ = 10 Then
                .ColData(mcolע��֤��) = 4
            ElseIf mint�༭״̬ = 6 Then
                '����
                For intCol = 0 To .Cols - 1
                    .ColData(intCol) = 5
                Next
                mshBill.ColData(mCol����) = 0
                mshBill.ColData(mCol��������) = 4
                txtժҪ.Enabled = True
                
            ElseIf mint�༭״̬ = 9 Then
                '�˲�
                '���˺�:���Ӻ˲���2007/05/13:10557
                '���˺�:20070530,���Ӻ˲�����ʱ���ݸ�����������Ӧ�ı༭��Ŀ
                Call Set��������Update
                txtժҪ.Enabled = True
            ElseIf mint�༭״̬ = 3 Then
                '���
                '���,Ҫ����ĳɱ���
                '���˺�:20070530,���Ӻ˲�����ʱ���ݸ�����������Ӧ�ı༭��Ŀ
                Call Set��������Update
                For intRow = 1 To .Rows - 1
                    If Val(.TextMatrix(intRow, mcol����ID)) > 0 Then    '����Ǳ������ģ��Ѿ��ڹ����д����˴˴����ٴ������Բ���ʾ
                        chkת���ƿ�.Visible = False
                        cboEnterStock.Visible = False
                        cboType.Visible = False
                        Exit For
                    End If
                Next
            End If
        Else
            '1.������2���޸ģ�3�����գ�4���鿴��5���޸ķ�Ʊ��6��������
            '7��������ˣ������������µ��ݲ���ˣ��Ѹ���ĵ��ݲ����������ˣ�ͬ����������˺�ĵ��ݲ����������;
            '8�����Ŀ��˻�,9-�˲�
            If mint�༭״̬ = 7 Then
            
                txtProvider.Enabled = False
                cmdProvider.Enabled = False
                txtժҪ.Enabled = False
                cboStock.Enabled = False
                '���˺�:20070530,���Ӻ˲�����ʱ���ݸ�����������Ӧ�ı༭��Ŀ
                Call Set��������Update
                '                For intCol = 0 To .Cols - 1
                '                    .ColData(intCol) = 5
                '                Next
                '                .ColData(mCol�����) = 4
                '                .ColData(mCol������) = 4
                '                .LocateCol = mCol�����
                Exit Sub
            ElseIf mint�༭״̬ = 8 Or mbln�˻� Then
                .ColData(mCol����) = 5
                .ColData(mCol�ɹ���) = 5
                .ColData(mColЧ��) = 5
                .ColData(3) = 5
                .ColData(mCol�����) = 5
                .ColData(mCol������) = 5
                .ColData(mcolע��֤��) = 5
                .ColData(mcol��Ʒ����) = 5
                .ColData(mcol���Ч��) = 5
                .ColData(mcol�������) = 5
                
                .ColData(mcol��������) = 5
                .ColData(mCol����) = 5
                
                If mbln�˻� Then
                    txtProvider.Enabled = False
                    cmdProvider.Enabled = False
                End If
                '�˻���������ѡ��ⷿ
                cboStock.Enabled = False
                Exit Sub
            ElseIf mint�༭״̬ = 3 Then
                For intRow = 1 To .Rows - 1
                    If Val(.TextMatrix(intRow, mcol����ID)) > 0 Then    '����Ǳ������ģ��Ѿ��ڹ����д����˴˴����ٴ������Բ���ʾ
                        chkת���ƿ�.Visible = False
                        cboEnterStock.Visible = False
                        cboType.Visible = False
                        Exit For
                    End If
                Next
            End If
            .ColData(0) = 5
            .ColData(mCol�к�) = 5
            .ColData(mCol����) = 1
            .ColData(mCol���) = 5
            .ColData(mCol���) = 5
            .ColData(mCol����) = 5
            .ColData(mCol��λ) = 5
            .ColData(mCol����) = 4
            .ColData(mcol��������) = 2
            .ColData(mcolע��֤��) = 4
            .ColData(mColЧ��) = 5
            .ColData(mCol����) = 4
            
            .ColData(mCol�ۼ�) = 5
            .ColData(mCol�ۼ۽��) = 5
            .ColData(mCol���) = 5
            '���˺�:���ۼ۴���
            .ColData(mcol���ۼ�) = 5
            .ColData(mcol���۽��) = 5
            .ColData(mcol���۲��) = 5
 
            .ColData(mCol���ս���) = IIf(mint�༭״̬ = 1 Or mint�༭״̬ = 2, 1, 5)
            .ColData(mCol�������) = 4
            .ColData(mCol��Ʊ��) = 4
            .ColData(mcol��Ʊ����) = 4
            .ColData(mCol��Ʊ����) = 2
            
            If mbln��ǿ�ƿ���ָ���۸� Then
                .ColData(mColָ��������) = 5
            Else
                .ColData(mColָ��������) = IIf(mbln�޸�������, 4, 5)
            End If
            
            .ColData(mColԭ����) = 5
            .ColData(mColԭ����) = 5
            .ColData(mCol����) = 5
            .ColData(mCol����ϵ��) = 5
            
            .ColData(mcolһ���Բ���) = 5
            .ColData(mcol�������) = 5
            .ColData(mcol���Ч��) = 5
            .ColData(mcol�������) = 2
            .ColData(mcol���ʧЧ��) = 5
            .ColData(mcolע��֤��) = 4
            .ColData(mcolע��֤��Ч��) = 5

            .ColData(mCol�ɹ���) = 4
            .ColData(mCol�����) = 4
            .ColData(mCol������) = 4
            
            .ColData(mCol�ɹ���) = 4
            .ColData(mCol����) = 4
            .ColData(mCol��Ʊ���) = 4
                  
            .ColAlignment(mCol����) = flexAlignLeftCenter
            .ColAlignment(mCol���) = flexAlignLeftCenter
            .ColAlignment(mCol����) = flexAlignLeftCenter
            .ColAlignment(mCol��λ) = flexAlignCenterCenter
            .ColAlignment(mCol����) = flexAlignLeftCenter
            .ColAlignment(mColЧ��) = flexAlignLeftCenter
            .ColAlignment(mCol����) = flexAlignRightCenter
            .ColAlignment(mCol�ɹ���) = flexAlignRightCenter
            .ColAlignment(mCol�����) = flexAlignRightCenter
            .ColAlignment(mCol������) = flexAlignRightCenter
            .ColAlignment(mCol�ۼ�) = flexAlignRightCenter
            .ColAlignment(mCol�ۼ۽��) = flexAlignRightCenter
            .ColAlignment(mCol���) = flexAlignRightCenter

            
            '���˺�:���ۼ۴���
            .ColAlignment(mcol���ۼ�) = flexAlignRightCenter
            .ColAlignment(mcol���۽��) = flexAlignRightCenter
            .ColAlignment(mcol���۲��) = flexAlignRightCenter
            .ColAlignment(mcol���۵�λ) = flexAlignCenterCenter
            
            .ColAlignment(mCol����) = flexAlignRightCenter
            .ColAlignment(mCol�������) = flexAlignLeftCenter
            .ColAlignment(mCol��Ʊ��) = flexAlignLeftCenter
            .ColAlignment(mcol��Ʊ����) = flexAlignLeftCenter
            .ColAlignment(mCol��Ʊ����) = flexAlignLeftCenter
            .ColAlignment(mCol��Ʊ���) = flexAlignRightCenter
            
            .ColAlignment(mcol�������) = flexAlignCenterCenter
            .ColAlignment(mcol���ʧЧ��) = flexAlignCenterCenter
            .ColAlignment(mcolע��֤��) = flexAlignLeftCenter
            .ColAlignment(mcol��Ʒ����) = flexAlignLeftCenter
            
            cboStock.Enabled = True
                        
            txtProvider.Enabled = True
            cmdProvider.Enabled = True
            txtժҪ.Enabled = True
        End If
    End With
End Sub


Private Sub initGrid()
    '����ʼ������ʼ��ժҪ�ı���ĳ���
    On Error GoTo ErrHandle
    Dim intCol As Integer
    
    With mshBill
        .Active = True
        .Cols = mCols
        .Value = Format(sys.Currentdate, "yyyy-mm-dd")
        .MsfObj.FixedCols = 1
        Call SetColumnByUserDefine
        .TextMatrix(0, mCol�к�) = ""
        .TextMatrix(0, mCol����) = "���������"
        .TextMatrix(0, mCol���) = "���"
        .TextMatrix(0, mCol��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mCol���) = "���"
        .TextMatrix(0, mCol����) = "����"
        .TextMatrix(0, mcol��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mCol��λ) = "��λ"
        .TextMatrix(0, mCol����) = "����"
        .TextMatrix(0, mcol��������) = "��������"
        .TextMatrix(0, mColЧ��) = "ʧЧ��"
        .TextMatrix(0, mCol����) = "����"
        .TextMatrix(0, mCol��������) = "��������"
        .TextMatrix(0, mCol����) = "����"
        
        .TextMatrix(0, mCol�ɹ���) = "�ɹ���"
        .TextMatrix(0, mCol�����) = "�����"
        .TextMatrix(0, mCol������) = "������"
        .TextMatrix(0, mCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mCol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mCol���) = "���"
        
        .TextMatrix(0, mcol���ۼ�) = "���ۼ�"
        .TextMatrix(0, mcol���۵�λ) = "���۵�λ"
        .TextMatrix(0, mcol���۽��) = "���۽��"
        .TextMatrix(0, mcol���۲��) = "���۲��"
        
        .TextMatrix(0, mCol����) = "����"
        .TextMatrix(0, mcol�ӳ���) = "�ӳ���"
        
        .TextMatrix(0, mCol���ս���) = "���ս���"
        .TextMatrix(0, mCol�������) = "�������"
        .TextMatrix(0, mCol��Ʊ��) = "��Ʊ��"
        .TextMatrix(0, mcol��Ʊ����) = "��Ʊ����"
        .TextMatrix(0, mCol��Ʊ����) = "��Ʊ����"
        .TextMatrix(0, mCol��Ʊ���) = "��Ʊ���"
        .TextMatrix(0, mColָ��������) = "�ɹ��޼�"
        .TextMatrix(0, mColԭ����) = "ԭ����"
        .TextMatrix(0, mColԭ����) = "ԭЧ��"
        .TextMatrix(0, mCol����) = "����"
        .TextMatrix(0, mCol����ϵ��) = "����ϵ��"
        
        .TextMatrix(0, mcolһ���Բ���) = "һ���Բ���"
        .TextMatrix(0, mcol�������) = "�������"
        .TextMatrix(0, mcol���Ч��) = "���Ч��"
        .TextMatrix(0, mcol�������) = "�������"
        .TextMatrix(0, mcol���ʧЧ��) = "���ʧЧ��"
        .TextMatrix(0, mcolע��֤��) = "ע��֤��"
        .TextMatrix(0, mcol��Ʒ����) = "��Ʒ����"
        .TextMatrix(0, mcol�ڲ�����) = "�ڲ�����"
        .TextMatrix(0, mcol����ID) = "����ID"
        .TextMatrix(0, mcolע��֤��Ч��) = "ע��֤��Ч��"

        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mCol�к�) = "1"
        
        .ColWidth(0) = 0
        
        .ColWidth(mCol�к�) = 300
        .ColWidth(mCol����) = 2000
        .ColWidth(mCol���) = 0
        .ColWidth(mCol��Ʒ��) = 900
        .ColWidth(mCol���) = 900
        .ColWidth(mCol����) = 800
        .ColWidth(mcol��׼�ĺ�) = 1000
        .ColWidth(mCol��λ) = 500
        .ColWidth(mCol����) = 800
        .ColWidth(mcol��������) = 1000
        .ColWidth(mColЧ��) = 1000
        .ColWidth(mCol����) = 800
        .ColWidth(mCol��������) = IIf(mint�༭״̬ = 6, 800, 0)
        .ColWidth(mCol����) = 0
        
        .ColWidth(mCol�����) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mCol�ɹ���) = IIf(mblnCostView = False, 0, 900)
        
        .ColWidth(mCol������) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mCol�ۼ�) = 900
        .ColWidth(mCol�ۼ۽��) = 900
        .ColWidth(mCol���) = IIf(mblnCostView = False, 0, 900)
            
        '���˺�:���ۼ۴���
        .ColWidth(mcol���ۼ�) = 900
        .ColWidth(mcol���۽��) = 900
        .ColWidth(mcol���۲��) = 800
        .ColWidth(mcol���۵�λ) = 800
        If mbln��ǿ�ƿ���ָ���۸� Then
            .ColWidth(mColָ��������) = 0
        Else
            .ColWidth(mColָ��������) = 900
        End If
        
        .ColWidth(mCol����) = 800
        
        If mbln�ֶμӳ��� = True Then
            .ColWidth(mcol�ӳ���) = 0
        Else
            .ColWidth(mcol�ӳ���) = 1200
        End If
        .ColWidth(mCol���ս���) = 4500
        .ColWidth(mCol�������) = 800
        .ColWidth(mCol��Ʊ��) = 800
        .ColWidth(mcol��Ʊ����) = 1000
        .ColWidth(mCol��Ʊ����) = 1000
        .ColWidth(mCol��Ʊ���) = 900
        .ColWidth(mColԭ����) = 0
        .ColWidth(mColԭ����) = 0
        .ColWidth(mCol����) = 0
        .ColWidth(mCol����ϵ��) = 0
        .ColWidth(mcolһ���Բ���) = 0
        .ColWidth(mcol�������) = 0
        .ColWidth(mcol���Ч��) = 0
        .ColWidth(mcol�������) = 1200
        .ColWidth(mcol���ʧЧ��) = 1200
        .ColWidth(mcolע��֤��) = 1600
        .ColWidth(mcol��Ʒ����) = IIf(gblnCode = True, 2000, 0)
        .ColWidth(mcol�ڲ�����) = 0
        .ColWidth(mcol����ID) = 0
        .ColWidth(mcolע��֤��Ч��) = 0
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mCol�к�) = 5
        .ColData(mCol����) = 1
        .ColData(mCol���) = 5
        .ColData(mCol��Ʒ��) = 5
        .ColData(mCol���) = 5
        .ColData(mCol����) = 5
        .ColData(mCol��λ) = 5
        .ColData(mCol����) = 4
        .ColData(mcol��������) = 2
        .ColData(mcol��׼�ĺ�) = 4
        .ColData(mColЧ��) = 5
        .ColData(mCol����) = 4
        .ColData(mCol��������) = 5
        .ColData(mCol����) = 5
        .ColData(mcol�ӳ���) = 5
        .ColData(mCol�ۼ�) = 5
        .ColData(mCol�ۼ۽��) = 5
        .ColData(mCol���) = 5
        '���˺�:���ۼ۴���
        .ColData(mcol���۵�λ) = 5
        .ColData(mcol���ۼ�) = IIf(mbln�˻� Or mint�༭״̬ = 8, 5, 4)
        .ColData(mcol���۽��) = 5
        .ColData(mcol���۲��) = 5
        
        .ColData(mCol���ս���) = IIf(mint�༭״̬ = 1 Or mint�༭״̬ = 2, 1, 5)
        .ColData(mCol��Ʊ��) = 4
        .ColData(mcol��Ʊ����) = 5
        .ColData(mCol�������) = 4
        .ColData(mCol��Ʊ����) = 2
        
        If mbln��ǿ�ƿ���ָ���۸� Then
            .ColData(mColָ��������) = 5
        Else
            .ColData(mColָ��������) = IIf(mbln�޸�������, 4, 5)
        End If
        .ColData(mColԭ����) = 5
        .ColData(mColԭ����) = 5
        .ColData(mCol����) = 5
        .ColData(mCol����ϵ��) = 5
        

        .ColData(mcolһ���Բ���) = 5
        .ColData(mcol�������) = 5
        .ColData(mcol���Ч��) = 5
        .ColData(mcol�������) = 2
        .ColData(mcol���ʧЧ��) = 5
        .ColData(mcolע��֤��) = 5
        .ColData(mcol��Ʒ����) = 4
        .ColData(mcol�ڲ�����) = 5
        .ColData(mcol����ID) = 5
        .ColData(mcolע��֤��Ч��) = 5

         .ColData(mCol�����) = 4
         .ColData(mCol������) = 4
        
        .ColData(mCol����) = 4
        .ColData(mCol�ɹ���) = 4
        .ColData(mCol��Ʊ���) = 4
        
        .ColAlignment(mCol����) = flexAlignLeftCenter
        .ColAlignment(mCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mCol���) = flexAlignLeftCenter
        .ColAlignment(mCol����) = flexAlignLeftCenter
        .ColAlignment(mCol���ս���) = flexAlignLeftCenter
        .ColAlignment(mCol��λ) = flexAlignCenterCenter
        .ColAlignment(mCol����) = flexAlignLeftCenter
        .ColAlignment(mColЧ��) = flexAlignLeftCenter
        .ColAlignment(mCol����) = flexAlignRightCenter
        .ColAlignment(mCol��������) = flexAlignRightCenter
        .ColAlignment(mCol�����) = flexAlignRightCenter
        .ColAlignment(mCol������) = flexAlignRightCenter
        .ColAlignment(mCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mCol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mCol���) = flexAlignRightCenter
        '���˺�:���ۼ۴���
        .ColAlignment(mcol���۵�λ) = flexAlignCenterCenter
        .ColAlignment(mcol���ۼ�) = flexAlignRightCenter
        .ColAlignment(mcol���۽��) = flexAlignRightCenter
        .ColAlignment(mcol���۲��) = flexAlignRightCenter
        
        .ColAlignment(mCol����) = flexAlignRightCenter
        .ColAlignment(mCol��Ʊ��) = flexAlignLeftCenter
        .ColAlignment(mcol��Ʊ����) = flexAlignLeftCenter
        .ColAlignment(mCol��Ʊ����) = flexAlignLeftCenter
        .ColAlignment(mCol��Ʊ���) = flexAlignRightCenter
                 

        .ColAlignment(mcol�������) = flexAlignLeftCenter
        .ColAlignment(mcol��������) = flexAlignLeftCenter
        .ColAlignment(mcol���ʧЧ��) = flexAlignLeftCenter
        .ColAlignment(mcolע��֤��) = flexAlignLeftCenter
        .ColAlignment(mcol��Ʒ����) = flexAlignLeftCenter
        
        .PrimaryCol = mCol����
        .LocateCol = mCol����
        Call SetColumnByUserDefine
        '���˺�:�������С��λ,��Ҫ������:
        If mintUnit = 0 Then
            .ColWidth(mcol���۵�λ) = 0
            .ColWidth(mcol���ۼ�) = 0
            .ColWidth(mcol���۽��) = 0
            .ColWidth(mcol���۲��) = 0
            .ColWidth(mcol���۵�λ) = 5
            .ColWidth(mcol���ۼ�) = 5
            .ColWidth(mcol���۽��) = 5
            .ColWidth(mcol���۲��) = 5
        End If
    End With
    
    
    txtժҪ.MaxLength = sys.FieldsLength("ҩƷ�շ���¼", "ժҪ")
    
    '��ֵ����
    Select Case mint�༭״̬
    Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
        '1.����; 2.�༭; 3.���; 4:��ѯ; 6:����; 7:�������; 8:�˻�; 9:�˲�; 10:�޸�ע��֤��
        Set mrsCostlyInfo = New ADODB.Recordset
        With mrsCostlyInfo
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .CursorType = adOpenStatic
            .Fields.Append "SN", adInteger, , adFldIsNullable
            .Fields.Append "id", adInteger, , adFldIsNullable
            .Fields.Append "����", adLongVarChar, 100, adFldIsNullable
            .Fields.Append "��������", adVarChar, 64, adFldIsNullable
            .Fields.Append "סԺ��", adVarChar, 20, adFldIsNullable
            .Fields.Append "����", adVarChar, 10, adFldIsNullable
            .Open
        End With
        If mint�༭״̬ <> 1 Then
            Dim rsTmp As ADODB.Recordset
            Dim strTmp As String
            
            strTmp = "select a.��� SN, c.id ����id, b.����, b.��������, b.סԺ��, b.���� " _
                   & "from ҩƷ�շ���¼ a, �շ���¼������Ϣ b, ���ű� c " _
                   & "where a.id=b.�շ�id and b.����=c.����(+) and a.no=[1] " _
                   & "order by a.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, mstrCaption, mstr���ݺ�)
            Do While Not rsTmp.EOF
                mrsCostlyInfo.AddNew
                mrsCostlyInfo!sn = rsTmp!sn
                mrsCostlyInfo!Id = rsTmp!����id
                mrsCostlyInfo!���� = rsTmp!����
                mrsCostlyInfo!�������� = rsTmp!��������
                mrsCostlyInfo!סԺ�� = rsTmp!סԺ��
                mrsCostlyInfo!���� = rsTmp!����
                mrsCostlyInfo.Update
                rsTmp.MoveNext
            Loop
            rsTmp.Close
        End If
       
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    
    With Pic����
        
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200

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
        .Top = Lbl��������.Top - 80
        .Left = Lbl��������.Left + Lbl��������.Width + 100
    End With
    
    With Lbl������
        .Top = Lbl��������.Top - .Height - 140
        .Left = mshBill.Left + 100
    End With
    
    With Txt������
        .Top = Lbl������.Top - 80
        .Left = Lbl������.Left + Lbl������.Width + 100
    End With
    
    With lbl�˲���
        .Top = Lbl������.Top
        .Left = Abs(mshBill.Width - .Width - txt�˲���.Width - 100) / 2
    End With
    With txt�˲���
        .Top = lbl�˲���.Top - 80
        .Left = lbl�˲���.Left + lbl�˲���.Width + 100
    End With
    
    With lbl�˲�����
        .Top = Lbl��������.Top
        .Left = lbl�˲���.Left
    End With
    With txt�˲�����
        .Top = Txt��������.Top
        .Left = txt�˲���.Left
    End With
    
    
    With Txt�������
        .Top = Lbl��������.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With Lbl�������
        .Top = Lbl��������.Top
        .Left = Txt�������.Left - 100 - .Width
    End With
    
    With Txt�����
        .Top = Lbl������.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With lbl�����
        .Top = Lbl������.Top
        .Left = Txt�����.Left - 100 - .Width
    End With
    
    With txtժҪ
        .Top = Lbl������.Top - 140 - .Height
        .Left = Txt������.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With
    
    With lblժҪ
        .Top = txtժҪ.Top + 50
        .Left = txtժҪ.Left - .Width - 100
    End With
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    If mblnCostView = False Then
        lblPurchasePrice.Visible = False
    End If
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 3
    End With
    
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnCostView = False Then
        lblDifference.Visible = False
    End If
    
    With mshBill
        '��ֵ����
        picCostly.Visible = vsfCostlyInfo.Visible
        If vsfCostlyInfo.Visible Then
            picCostly.Height = 400
            picCostly.Left = .Left
            picCostly.Width = .Width
            vsfCostlyInfo.Height = 650
            vsfCostlyInfo.Left = .Left
            vsfCostlyInfo.Width = .Width
            .Height = lblPurchasePrice.Top - .Top - 60 - vsfCostlyInfo.Height - picCostly.Height
            picCostly.Top = .Top + .Height + 40
            vsfCostlyInfo.Top = picCostly.Top + picCostly.Height + 10
        Else
            .Height = lblPurchasePrice.Top - .Top - 60
        End If
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
        
    With cmdExtractData
        .Top = CmdCancel.Top
    End With
    
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    If mint�༭״̬ = 5 Then '�޸ķ�Ʊ��Ϣ�ð�ť�ſ���
        With cmdBulkCopy
            cmdBulkCopy.Visible = True
            .Top = cmdFind.Top
            If txtCode.Visible Then
                .Left = txtCode.Left + txtCode.Width + 100
            Else
                .Left = cmdFind.Left + cmdFind.Width + 100
            End If
        End With
        
        With cmdALLDel
            .Visible = True
            .Left = cmdBulkCopy.Left + cmdBulkCopy.Width + 100
            .Top = cmdBulkCopy.Top
        End With
    End If
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    Me.fraMoveNO.Top = txtCode.Top
    Me.fraMoveNO.Left = txtCode.Left + txtCode.Width + 50
    
    With cmdCopy
        .Left = IIf(txtCode.Visible, txtCode.Left + txtCode.Width, cmdFind.Left + cmdFind.Width) + 100
        .Top = cmdFind.Top
    End With
    
    With txtCopy
        .Left = cmdCopy.Left + cmdCopy.Width + 50
        .Top = txtCode.Top
    End With
    
    With lblCopy
        .Left = txtCopy.Left + txtCopy.Width + 25
        .Top = txtCopy.Top + 45
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mshProvider.Visible = True Then
        mshProvider.Visible = False
        txtProvider.SetFocus
        txtProvider.SelLength = Len(txtProvider.Text)
        txtProvider.SelStart = 0
        Cancel = True
        Exit Sub
    End If
    
    If msh����.Visible = True Then
        msh����.Visible = False
        mshBill.SetFocus
        mshBill.Col = mCol����
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, mstrCaption
        Call SaveBILLWidth
        Call zlDatabase.SetPara("��˲�������", cboType.ItemData(cboType.ListIndex), glngSys, mlngModule)
        Exit Sub
    End If
    
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, mstrCaption
    End If
    Call SaveBILLWidth
    Call zlDatabase.SetPara("��˲�������", cboType.ItemData(cboType.ListIndex), glngSys, mlngModule)
    '��ֵ����
    If mrsCostlyInfo Is Nothing Then Exit Sub
    If mrsCostlyInfo.State = adStateOpen Then mrsCostlyInfo.Close
End Sub

Private Function SaveCheck(Optional ByVal strNo As String = "") As Boolean
    mblnSave = False
    SaveCheck = False
    
    gstrSQL = "zl_�����⹺_Verify('" & IIf(mint�༭״̬ = 7, strNo, txtNO.Tag) & "','" & UserInfo.�û��� & "',to_date('" & Format(mstr�������, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'))"
    
    On Error GoTo ErrHandle
    'If mint�༭״̬ <> 7 Then gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
   ' If mint�༭״̬ <> 7 Then gcnOracle.CommitTrans
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    'If mint�༭״̬ <> 7 Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function
 




Private Sub lblType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.PopupMenu mnuSearch
End Sub

Private Sub mnuSearch01_Click()
    lblType.Caption = "����ID��"
    If lblType.Tag <> 1 Then txtTypeVar.Text = ""
    lblType.Tag = 1
End Sub

Private Sub mnuSearch02_Click()
    lblType.Caption = "����������"
    If lblType.Tag <> 2 Then txtTypeVar.Text = ""
    lblType.Tag = 2
End Sub

Private Sub mnuSearch03_Click()
    lblType.Caption = "סԺ�š�"
    If lblType.Tag <> 3 Then txtTypeVar.Text = ""
    lblType.Tag = 3
End Sub

Private Sub mnuSearch04_Click()
    lblType.Caption = "����š�"
    If lblType.Tag <> 4 Then txtTypeVar.Text = ""
    lblType.Tag = 4
End Sub

Private Sub mnuSearch05_Click()
    lblType.Caption = "���š�"
    If lblType.Tag <> 5 Then txtTypeVar.Text = ""
    lblType.Tag = 5
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mCol�к�, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call ��ʾ�ϼƽ��
    Call RefreshRowNO(mshBill, mCol�к�, mshBill.Row)
    If mshBill.TextMatrix(mshBill.Row, 0) <> "" Then
        CostlyInfo_Refresh Val(mshBill.TextMatrix(mshBill.Row, 1)), IsCostly(mshBill.TextMatrix(mshBill.Row, 0))
    End If
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, ",3,4,5,6,7,9,10,", "," & mint�༭״̬ & ",") <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("��ȷʵҪɾ�������������ϣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
                Exit Sub
            End If
            CostlyInfo_Refresh Val(mshBill.TextMatrix(mshBill.Row, 1)), False
            '������Ӧ��ֵ���ϵ�SN
            RecountSN mshBill.Row
        End If
    End With
End Sub
Private Sub ColMoveNextCol(ByVal lngCol As Long)
    '------------------------------------------------------------------------------
    '����:���ƶ�
    '����:
    '����:���˺�
    '����:2007/08/14
    '------------------------------------------------------------------------------
    Dim i As Long
    With mshBill
        For i = lngCol + 1 To .Cols - 1
            Select Case .ColData(i)
            Case -1, 1, 2, 3, 4
                '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
                ' 0����ʾ���п���ѡ�񣬵������޸�
                ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
                ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
                ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
                '4:  ��ʾ����Ϊ�������ı����û�����
                '5:  ��ʾ���в�����ѡ��
                .Col = i
                Exit For
            End Select
        Next
        
        If i - 1 = .Cols - 1 Then
            If .Row = .Rows - 1 Then
                .Rows = .Rows + 1
            End If
            .Row = .Row + 1
            .Col = mCol����
        End If
    End With
    
End Sub

Private Sub mshbill_CommandClick()
    Dim i As Integer
    Dim int����� As Integer
    Dim rs���ս��� As Recordset
    
    On Error GoTo ErrHandle
    
    int����� = mshBill.Row

     If mshBill.Col = mCol���� Then
        Dim mrsReturn As Recordset
        If mint�༭״̬ = 8 Or mbln�˻� Then
            If Val(txtProvider.Tag) = 0 Then
                ShowMsgBox "δѡ���˻���λ!"
                If txtProvider.Enabled Then txtProvider.SetFocus
                Exit Sub
            End If
        End If
        Set mrsReturn = Frm����ѡ����.ShowMe(Me, IIf(mint�༭״̬ = 8 Or mbln�˻�, 2, 1), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , True, True, False, False, True, IIf(mint�༭״̬ = 8 Or mbln�˻�, Val(txtProvider.Tag), 0), IIf(mintUnit = 0, True, False), , , , 1712, , mstrPrivs, , False)
        
        If mrsReturn.RecordCount > 0 And (mint�༭״̬ = 8 Or mbln�˻� = True) Then
            Set mrsReturn = CheckRedo(mrsReturn) '����ظ���¼�����ظ��ļ�¼���˵�Ȼ�󷵻ع��˺�����ݼ�
        End If
        
        If mrsReturn.RecordCount > 0 Then
            With mshBill
                mrsReturn.MoveFirst
                For i = 1 To mrsReturn.RecordCount
                    If CheckQualifications(mlngModule, 0, Val(mrsReturn!����ID)) = False Then Exit Sub
                    
                    SetColValue .Row, mrsReturn!����ID, "[" & mrsReturn!���� & "]" & mrsReturn!����, IIf(IsNull(mrsReturn!���), "", mrsReturn!���), _
                        IIf(IsNull(mrsReturn!����), "", mrsReturn!����), _
                        IIf(mintUnit = 0, mrsReturn!ɢװ��λ, mrsReturn!��װ��λ), _
                        IIf(IsNull(mrsReturn!�ۼ�), 0, mrsReturn!�ۼ�) * IIf(mintUnit = 0, 1, mrsReturn!����ϵ��), _
                         mrsReturn!ָ�������� * IIf(mintUnit = 0, 1, mrsReturn!����ϵ��), _
                        IIf(IsNull(mrsReturn!����), "!", mrsReturn!����), mrsReturn!���Ч��, "", _
                        IIf(mintUnit = 0, 1, mrsReturn!����ϵ��), IIf(IsNull(mrsReturn!����), 0, mrsReturn!����), mrsReturn!ʱ��, _
                        mrsReturn!���÷���, mrsReturn!ָ������� / 100, IIf(IsNull(mrsReturn!��׼�ĺ�), "", mrsReturn!��׼�ĺ�), IIf(IsNull(mrsReturn!��Ʒ��), "", mrsReturn!��Ʒ��)
                    
                    Call ColMoveNextCol(.Col)
                    
                    '��ֵ��������
                    If .TextMatrix(.Row, 0) = "" Then
                        vsfCostlyInfo.Visible = False
                    Else
                        vsfCostlyInfo.Visible = IsCostly(.TextMatrix(.Row, 0))
                    End If
                    CostlyInfo_Refresh Val(.TextMatrix(.Row, 1)), vsfCostlyInfo.Visible
                    Call Form_Resize
                    
                    mblnChange = True

                    If .Row = .Rows - 1 Then .Rows = .Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
                    .Row = .Row + 1
                    
                    mrsReturn.MoveNext
                Next
                
                mshBill.Row = int�����
            
                
'                If mrsReturn.RecordCount = 1 Then
'                    If CheckQualifications(mlngModule, 0, Val(mrsReturn!����ID)) = False Then Exit Sub
'
'                    SetColValue .Row, mrsReturn!����ID, "[" & mrsReturn!���� & "]" & mrsReturn!����, IIf(IsNull(mrsReturn!���), "", mrsReturn!���), _
'                        IIf(IsNull(mrsReturn!����), "", mrsReturn!����), _
'                        IIf(mintUnit = 0, mrsReturn!ɢװ��λ, mrsReturn!��װ��λ), _
'                        IIf(IsNull(mrsReturn!�ۼ�), 0, mrsReturn!�ۼ�) * IIf(mintUnit = 0, 1, mrsReturn!����ϵ��), _
'                         mrsReturn!ָ�������� * IIf(mintUnit = 0, 1, mrsReturn!����ϵ��), _
'                        IIf(IsNull(mrsReturn!����), "!", mrsReturn!����), mrsReturn!���Ч��, "", _
'                        IIf(mintUnit = 0, 1, mrsReturn!����ϵ��), IIf(IsNull(mrsReturn!����), 0, mrsReturn!����), mrsReturn!ʱ��, _
'                        mrsReturn!���÷���, mrsReturn!ָ������� / 100, IIf(IsNull(mrsReturn!��׼�ĺ�), "", mrsReturn!��׼�ĺ�)
'
'                    Call ColMoveNextCol(.Col)
'
'                    '��ֵ��������
'                    If .TextMatrix(.Row, 0) = "" Then
'                        vsfCostlyInfo.Visible = False
'                    Else
'                        vsfCostlyInfo.Visible = IsCostly(.TextMatrix(.Row, 0))
'                    End If
'                    CostlyInfo_Refresh Val(.TextMatrix(.Row, 1)), vsfCostlyInfo.Visible
'                    Call Form_Resize
'
'                    mblnChange = True
'                End If
            End With
            mrsReturn.Close
        End If
    ElseIf mshBill.Col = mCol���ս��� Then
        gstrSQL = "Select ���� as id,null as �ϼ�id,����,����,1 as ĩ�� From ������ս��� Order By ���� "
        Set rs���ս��� = zlDatabase.ShowSelect(Me, gstrSQL, 1, "������ս���", True, , "ѡ��������ս���")
        If rs���ս��� Is Nothing Then Exit Sub
        If rs���ս���.State <> 1 Then Exit Sub
        
        With rs���ս���
            mshBill.TextMatrix(mshBill.Row, mCol���ս���) = zlStr.Nvl(!����)
        End With
        
        Call ColMoveNextCol(mshBill.Col)
    Else
        Dim rstemp As New Recordset
        
        gstrSQL = "Select rownum as id,null as �ϼ�id,����,����,����,1 as ĩ�� From ���������� "
        Set rstemp = zlDatabase.ShowSelect(Me, gstrSQL, 1, "����������ѡ��", True, , "ѡ���������������̻���")
        
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
        If rstemp Is Nothing Then Exit Sub
        If rstemp.State <> 1 Then Exit Sub
        
        With rstemp
            If CheckQualifications(mlngModule, 1, CStr(zlStr.Nvl(!����))) = False Then Exit Sub
            mshBill.TextMatrix(mshBill.Row, mCol����) = zlStr.Nvl(!����)
        End With
        
        gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mCol����), mshBill.TextMatrix(mshBill.Row, 0))
        If rstemp.RecordCount > 0 Then
            mshBill.TextMatrix(mshBill.Row, mcol��׼�ĺ�) = IIf(IsNull(rstemp!��׼�ĺ�), "", rstemp!��׼�ĺ�)
        Else
            mshBill.TextMatrix(mshBill.Row, mcol��׼�ĺ�) = ""
        End If
        Call ColMoveNextCol(mshBill.Col)
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub
Private Function �������ۼۼ����۲��(ByVal lngRow As Long, Optional bln���ۼ� As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ݿⷿ��λ����ɢװ��λ�����ۼۼ����
    '���:lngRow -ָ���������
    '     bln���ۼ�-���ۼ�Ϊ�ۼ�
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-28 12:09:04
    '-----------------------------------------------------------------------------------------------------------
    Dim dbl����ϵ�� As Double, arrSplit As Variant, dbl���� As Double
    
    
    With mshBill
    
        dbl����ϵ�� = Val(.TextMatrix(lngRow, mCol����ϵ��))
        
        dbl���� = IIf(mint�༭״̬ = 6, IIf(Val(.TextMatrix(lngRow, mCol��������)) = 0, Val(.Text), Val(.TextMatrix(lngRow, mCol��������))), Val(.TextMatrix(lngRow, mCol����)))
        If dbl���� = 0 Or Val(.TextMatrix(lngRow, 0)) = 0 Then
            .TextMatrix(lngRow, mcol���۽��) = 0
            .TextMatrix(lngRow, mcol���۲��) = 0
            .TextMatrix(lngRow, mCol���) = 0
            .TextMatrix(lngRow, mCol�ۼ۽��) = 0
            .TextMatrix(lngRow, mcol���ۼ�) = Format(Val(.TextMatrix(lngRow, mCol�ۼ�)) / IIf(dbl����ϵ�� = 0, 1, dbl����ϵ��), mFMT.FM_ɢװ���ۼ�)
            Exit Function
        End If
        
        '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
        If .TextMatrix(lngRow, mColԭ����) <> "" Then
           arrSplit = Split(.TextMatrix(lngRow, mColԭ����), "||")
           If Val(arrSplit(2)) = 1 And (IIf(mbln�ⷿ, arrSplit(4) = 1, arrSplit(3) = 1)) Then
                'ʵ������
                '���˺�:���ۼ۴���
                If bln���ۼ� Then
                    .TextMatrix(lngRow, mcol���ۼ�) = Format(Val(.TextMatrix(lngRow, mCol�ۼ�)) / dbl����ϵ��, mFMT.FM_ɢװ���ۼ�)
                End If
                .TextMatrix(lngRow, mcol���۽��) = Format(Val(.TextMatrix(lngRow, mcol���ۼ�)) * (dbl���� * dbl����ϵ��), mFMT.FM_���)
                '���۲��=���۽��-������
                .TextMatrix(lngRow, mcol���۲��) = Format(Val(.TextMatrix(lngRow, mcol���۽��)) - Val(.TextMatrix(lngRow, mCol������)), mFMT.FM_���)
           Else '����
                '���˺�:���ۼ۴���
                .TextMatrix(lngRow, mcol���ۼ�) = Format(Val(.TextMatrix(lngRow, mCol�ۼ�)) / dbl����ϵ��, mFMT.FM_ɢװ���ۼ�)
                .TextMatrix(lngRow, mcol���۽��) = Format(Val(.TextMatrix(lngRow, mcol���ۼ�)) * (dbl���� * dbl����ϵ��), mFMT.FM_���)
                '���۲��=���۽��-������
                .TextMatrix(lngRow, mcol���۲��) = Format(Val(.TextMatrix(lngRow, mcol���۽��)) - Val(.TextMatrix(lngRow, mCol������)), mFMT.FM_���)
           End If
        Else
            .TextMatrix(lngRow, mcol���ۼ�) = Format(Val(.TextMatrix(lngRow, mCol�ۼ�)) / dbl����ϵ��, mFMT.FM_ɢװ���ۼ�)
            .TextMatrix(lngRow, mcol���۽��) = Format(Val(.TextMatrix(lngRow, mcol���ۼ�)) * (dbl���� * dbl����ϵ��), mFMT.FM_���)
            '���۲��=���۽��-������
            .TextMatrix(lngRow, mcol���۲��) = Format(Val(.TextMatrix(lngRow, mcol���۽��)) - Val(.TextMatrix(lngRow, mCol������)), mFMT.FM_���)
        End If
    End With
    �������ۼۼ����۲�� = True
End Function


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        strKey = .Text
        If strKey = "" Then
            strKey = .TextMatrix(.Row, .Col)
        End If
        Select Case .Col
            Case mcol��Ʒ����
                Select Case KeyAscii
                    Case vbKeyBack, vbKeyEscape, 3, 22
                        Exit Sub
                    Case vbKeyReturn
'                        Call OS.PressKey(vbKeyTab)
                        Exit Sub
                    Case Else
                        '����¼�����ֺ���ĸ
                        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (InStr(1, "`������������������!@#$��%^&*()_-����������?'+{}����<>��������~[]:;'\|,./", Chr(KeyAscii)) > 0) Then Exit Sub
                End Select
                KeyAscii = 0
                Exit Sub
            Case mCol����, mCol��������
                intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.����С��, g_С��λ��.obj_ɢװС��.����С��)
            Case mCol�ɹ���, mCol�����
               intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.�ɱ���С��, g_С��λ��.obj_ɢװС��.�ɱ���С��)
            Case mCol������, mCol��Ʊ���
                intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.���С��, g_С��λ��.obj_ɢװС��.���С��)
            Case mCol�ۼ�, mcol���ۼ�
                intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.���ۼ�С��, g_С��λ��.obj_ɢװС��.���ۼ�С��)
        End Select
        
        If .Col = mCol���� Or .Col = mCol�������� Or .Col = mCol�ɹ��� Or .Col = mCol����� Or .Col = mCol������ Or .Col = mCol��Ʊ��� Or .Col = mCol�ۼ� Or .Col = mcol���ۼ� Then
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
    Dim str���� As String
    Dim strxq As String
    
    If mint�༭״̬ = 5 And Trim(mshBill.TextMatrix(mshBill.Row, mCol��Ʊ��)) <> "" Then '��ǰ�з�Ʊ�Ų�Ϊ�ղſ���
        cmdBulkCopy.Enabled = True
    Else
        cmdBulkCopy.Enabled = False
    End If

    
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        If .Row <> .LastRow Then
            lngRow = .LastRow
            If PicInput.Visible Then
                '���¼������ۼۡ����
                .TextMatrix(lngRow, mCol�ۼ�) = Format(У�����ۼ�(Val(.TextMatrix(lngRow, mCol�����)) * (1 + (Val(Txt�Ӽ���) / 100)) + _
                ʱ�۲������ۼ�(Val(.TextMatrix(lngRow, 0)), Val(.TextMatrix(lngRow, mCol�����)), Val(Txt�Ӽ���) / 100, lngRow), lngRow), mFMT.FM_�ɱ���)
                .TextMatrix(lngRow, mCol�ۼ۽��) = Format(Val(.TextMatrix(lngRow, mCol�ۼ�)) * Val(.TextMatrix(lngRow, mCol����)), mFMT.FM_���)
                .TextMatrix(lngRow, mCol���) = Format(IIf(.TextMatrix(lngRow, mCol�ۼ۽��) = "", 0, .TextMatrix(lngRow, mCol�ۼ۽��)) - IIf(.TextMatrix(lngRow, mCol������) = "", 0, .TextMatrix(lngRow, mCol������)), mFMT.FM_���)
                Call �������ۼۼ����۲��(lngRow)
                PicInput.Visible = False
            End If
        End If
        
        SetInputFormat .Row
        
        If Not (.Col = mCol����� Or .Col = mCol�ɹ��� Or .Col = mCol���� Or .Col = mCol������) Then PicInput.Visible = False
        
        If .Col = mCol������ And PicInput.Visible Then Txt�Ӽ���.SetFocus: Exit Sub
        If .Col = mCol���� And PicInput.Visible Then Txt�Ӽ���.SetFocus: Exit Sub
        
        Select Case .Col
            Case mCol����
                .TxtCheck = False
                .MaxLength = 80
                
                'ֻ�������в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
                
            Case mCol����
                ImeLanguage True
                .TxtCheck = False
                .MaxLength = 30
                .TxtSetFocus
                
            Case mCol����
                .TxtCheck = False
                .MaxLength = mintBatchNoLen
            Case mcolע��֤��
                .TxtCheck = False
                .MaxLength = 50
            Case mcol��Ʒ����
                .TxtCheck = False
                .MaxLength = 50
            Case mColЧ��
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                .Value = Format(sys.Currentdate, "yyyy-mm-dd")
                
'                If Trim(.TextMatrix(.Row, mCol����)) = "" Or IsNumeric(.TextMatrix(.Row, mCol����)) = False Then
                    If Not IsDate(Trim(.TextMatrix(.Row, mcol��������))) Then
                        str���� = ""
                    Else
                        str���� = Format(.TextMatrix(.Row, mcol��������), "yyyymmdd")
                    End If
'                Else
'                    str���� = Trim(.TextMatrix(.Row, mCol����))
'                End If
                
                If str���� <> "" And Trim(.TextMatrix(.Row, mColԭ����)) <> "" Then
                    '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����

                    If IsNumeric(str����) And Split(.TextMatrix(.Row, mColԭ����), "||")(0) <> "0" Then
                        strxq = UCase(str����)
                        If Trim(.TextMatrix(.Row, mColЧ��)) = "" Then
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq, True)
                                If strxq = "" Then Exit Sub
                                
                                .TextMatrix(.Row, mColЧ��) = Format(DateAdd("M", Split(.TextMatrix(.Row, mColԭ����), "||")(0), strxq), "yyyy-mm-dd")
                                Call CheckLapse(.TextMatrix(.Row, mColЧ��))
                            End If
                        End If
                    End If
                End If
            Case mcol��������
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                .Value = Format(sys.Currentdate, "yyyy-mm-dd")
                If Trim(.TextMatrix(.Row, mCol����)) = "" Or IsNumeric(.TextMatrix(.Row, mCol����)) = False Then
                    If Not IsDate(Trim(.TextMatrix(.Row, mcol��������))) Then
                        str���� = ""
                    Else
                        str���� = Format(.TextMatrix(.Row, mcol��������), "yyyymmdd")
                    End If
                Else
                    str���� = Trim(.TextMatrix(.Row, mCol����))
                End If
                
                If str���� <> "" Then
                    
                    '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                    If IsNumeric(str����) And Split(.TextMatrix(.Row, mColԭ����), "||")(0) <> "0" Then
                        strxq = UCase(str����)
                        If Trim(.TextMatrix(.Row, mcol��������)) = "" Then
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq = "" Then Exit Sub
                                .TextMatrix(.Row, mcol��������) = Format(strxq, "yyyy-mm-dd")
                            End If
                        End If
                    End If
                End If
            Case mcol�������
                .TxtCheck = True
                .Value = Format(sys.Currentdate, "yyyy-mm-dd")
                .TextMask = "1234567890-"
                .MaxLength = 10
            Case mcol���ʧЧ��
                .TxtCheck = True
                .Value = Format(sys.Currentdate, "yyyy-mm-dd")
                .TextMask = "1234567890-"
                .MaxLength = 10
            Case mCol����
                .TxtCheck = True
                .MaxLength = 3
                .TextMask = ".1234567890"
                stbThis.Panels.Item(2) = .TextMatrix(.Row, mCol����) & "��ָ��������Ϊ��" & .TextMatrix(.Row, mColָ��������)
                
            Case mCol�����, mColָ��������, mCol�ɹ���
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mCol������
                .TxtCheck = True
                .MaxLength = 14
                .TextMask = ".1234567890"
            Case mcol���ۼ�
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
                
            Case mCol�ۼ�
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mCol����
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mCol��������
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mCol�������
                .TxtCheck = False
                .MaxLength = 200
            Case mCol���ս���
                .MaxLength = 100
            Case mCol��Ʊ��
                .TxtCheck = False
                .MaxLength = mint��Ʊ��Len
            Case mcol��Ʊ����
                .TxtCheck = True
                .MaxLength = 20
                .TextMask = "1234567890"
            Case mCol��Ʊ���
                .TxtCheck = True
                .MaxLength = 14
                .TextMask = "-.1234567890"
            Case mCol��Ʊ����
                .TxtCheck = True
                .TextMask = "1234567890-"
                .Value = sys.Currentdate
                .MaxLength = 10
        End Select
        
        '��ֵ����
        Select Case mint�༭״̬
            Case 3, 4, 5, 6, 7, 10
                vsfCostlyInfo.Enabled = False
                picCostly.Enabled = False
        End Select
        
        If .Row <> .LastRow Then
            '״̬�л�
            If .TextMatrix(.Row, 0) = "" Then
                vsfCostlyInfo.Visible = False
                Call Form_Resize
                Exit Sub
            Else
                vsfCostlyInfo.Visible = IsCostly(.TextMatrix(.Row, 0))
                '��ʾ����
                If vsfCostlyInfo.Visible Then
                    '��λ
                    If mrsCostlyInfo Is Nothing Then Exit Sub
                    If mrsCostlyInfo.RecordCount > 0 Then mrsCostlyInfo.MoveFirst
                    mrsCostlyInfo.Find "SN=" & .TextMatrix(.Row, 1)
                    If Not mrsCostlyInfo.EOF Then
                        vsfCostlyInfo.TextMatrix(1, 1) = IIf(IsNull(mrsCostlyInfo!����), "", mrsCostlyInfo!����)
                        vsfCostlyInfo.TextMatrix(1, 2) = IIf(IsNull(mrsCostlyInfo!��������), "", mrsCostlyInfo!��������)
                        vsfCostlyInfo.TextMatrix(1, 3) = IIf(IsNull(mrsCostlyInfo!סԺ��), "", mrsCostlyInfo!סԺ��)
                        vsfCostlyInfo.TextMatrix(1, 4) = IIf(IsNull(mrsCostlyInfo!����), "", mrsCostlyInfo!����)
                        vsfCostlyInfo.TextMatrix(1, 5) = IIf(IsNull(mrsCostlyInfo!Id), "", mrsCostlyInfo!Id)
                    Else
                        vsfCostlyInfo.TextMatrix(1, 1) = ""
                        vsfCostlyInfo.TextMatrix(1, 2) = ""
                        vsfCostlyInfo.TextMatrix(1, 3) = ""
                        vsfCostlyInfo.TextMatrix(1, 4) = ""
                        vsfCostlyInfo.TextMatrix(1, 5) = ""
                    End If
                    vsfCostlyInfo.Col = vsfCostlyInfo.ColIndex("����")
                End If
            End If
            Call Form_Resize
        End If
        
    End With
    
End Sub

Private Sub mshBill_GotFocus()
'    If mint�༭״̬ <> 1 Then
        '��ֵ����
        '״̬�л�
        With mshBill
            If .TextMatrix(.Row, 0) = "" Then
                vsfCostlyInfo.Visible = False
                Call Form_Resize
                Exit Sub
            Else
                vsfCostlyInfo.Visible = IsCostly(.TextMatrix(.Row, 0))
                '��ʾ����
                If vsfCostlyInfo.Visible Then
                    '��λ
                    Call Form_Resize
                    If mrsCostlyInfo Is Nothing Then
                        Exit Sub
                    End If
                    If mrsCostlyInfo.RecordCount > 0 Then
                        mrsCostlyInfo.MoveFirst
                    Else
                        Exit Sub
                    End If
                    mrsCostlyInfo.Find "SN=" & Val(.TextMatrix(.Row, 1))
                    If Not mrsCostlyInfo.EOF Then
                        vsfCostlyInfo.TextMatrix(1, 1) = IIf(IsNull(mrsCostlyInfo!����), "", mrsCostlyInfo!����)
                        vsfCostlyInfo.TextMatrix(1, 2) = IIf(IsNull(mrsCostlyInfo!��������), "", mrsCostlyInfo!��������)
                        vsfCostlyInfo.TextMatrix(1, 3) = IIf(IsNull(mrsCostlyInfo!סԺ��), "", mrsCostlyInfo!סԺ��)
                        vsfCostlyInfo.TextMatrix(1, 4) = IIf(IsNull(mrsCostlyInfo!����), "", mrsCostlyInfo!����)
                        vsfCostlyInfo.TextMatrix(1, 5) = IIf(IsNull(mrsCostlyInfo!Id), "", mrsCostlyInfo!Id)
                    End If
                    vsfCostlyInfo.Col = vsfCostlyInfo.ColIndex("����")
                End If
            End If
        End With
        Call Form_Resize
'    End If
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rstemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim dbl�ӳ��� As Double, dbl��Ʊ��� As Double
    Dim dblָ�����ۼ� As Double
    Dim dbl�ۼ� As Double, dbl�ɹ��� As Double, dbl����� As Double, dbl���� As Double, dbl���� As Double
    Dim lng����ID As Long
    Dim sng�ֶ��ۼ� As Double
    Dim dblCostPrice As Double, dblPrice As Double
    Dim strBidMess As String
    Dim dbl�ɱ��� As Double
    Dim intCol As Integer
    Dim i As Integer
    Dim int����� As Integer
    
    int����� = mshBill.Row

    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With mshBill
        .Text = Trim(.Text)
        strKey = Trim(.Text)
        
        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col
            Case mCol����
                If strKey <> "" Then
                    Dim mrsReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    
'                    If sngTop + 3630 > Screen.Height Then
'                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 3630
'                    End If
                    If mint�༭״̬ = 8 Or mbln�˻� Then
                        If Val(txtProvider.Tag) = 0 Then
                            ShowMsgBox "δѡ���˻���λ!"
                            Cancel = True
                            If txtProvider.Enabled Then txtProvider.SetFocus
                            Exit Sub
                        End If
                    End If
                    Set mrsReturn = FrmMulitSel.ShowSelect(Me, IIf(mint�༭״̬ = 8 Or mbln�˻�, 2, 1) _
                                  , cboStock.ItemData(cboStock.ListIndex) _
                                  , cboStock.ItemData(cboStock.ListIndex) _
                                  , cboStock.ItemData(cboStock.ListIndex) _
                                  , strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight _
                                  , True, True, False, False, True _
                                  , IIf(mint�༭״̬ = 8 Or mbln�˻�, Val(txtProvider.Tag), 0), IIf(mintUnit = 0, True, False), , , 1712, , mstrPrivs, , False)
                    
                    If mrsReturn.RecordCount > 0 And (mint�༭״̬ = 8 Or mbln�˻� = True) Then
                        Set mrsReturn = CheckRedo(mrsReturn) '����ظ���¼�����ظ��ļ�¼���˵�Ȼ�󷵻ع��˺�����ݼ�
                    End If
                    
                    If mrsReturn.RecordCount <= 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                    
                    mrsReturn.MoveFirst
                    For i = 1 To mrsReturn.RecordCount
                        If CheckQualifications(mlngModule, 0, Val(mrsReturn!����ID)) = False Then Exit Sub
                        
                        If SetColValue(.Row, mrsReturn!����ID, "[" & mrsReturn!���� & "]" & mrsReturn!����, IIf(IsNull(mrsReturn!���), "", mrsReturn!���), _
                                    IIf(IsNull(mrsReturn!����), "", mrsReturn!����), _
                                    IIf(mintUnit = 0, mrsReturn!ɢװ��λ, mrsReturn!��װ��λ), _
                                    IIf(IsNull(mrsReturn!�ۼ�), 0, mrsReturn!�ۼ�) * IIf(mintUnit = 0, 1, mrsReturn!����ϵ��), _
                                    mrsReturn!ָ�������� * IIf(mintUnit = 0, 1, mrsReturn!����ϵ��), _
                                    IIf(IsNull(mrsReturn!����), "!", mrsReturn!����), mrsReturn!���Ч��, "", _
                                    IIf(mintUnit = 0, 1, mrsReturn!����ϵ��), IIf(IsNull(mrsReturn!����), 0, mrsReturn!����), mrsReturn!ʱ��, _
                                    mrsReturn!���÷���, mrsReturn!ָ������� / 100, IIf(IsNull(mrsReturn!��׼�ĺ�), "", mrsReturn!��׼�ĺ�), IIf(IsNull(mrsReturn!��Ʒ��), "", mrsReturn!��Ʒ��)) Then
                            
                            If .Row = .Rows - 1 Then .Rows = .Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
                            .Row = .Row + 1
                            
                            .Text = .TextMatrix(.Row, .Col)
                        Else
                            Cancel = True
                        End If
                        
                        mrsReturn.MoveNext
                    Next
                    
                    mshBill.Row = int�����
                    
'                    If mrsReturn.RecordCount = 1 Then
'                        If CheckQualifications(mlngModule, 0, Val(mrsReturn!����ID)) = False Then Exit Sub
'
'                        If SetColValue(.Row, mrsReturn!����ID, "[" & mrsReturn!���� & "]" & mrsReturn!����, IIf(IsNull(mrsReturn!���), "", mrsReturn!���), _
'                                    IIf(IsNull(mrsReturn!����), "", mrsReturn!����), _
'                                    IIf(mintUnit = 0, mrsReturn!ɢװ��λ, mrsReturn!��װ��λ), _
'                                    IIf(IsNull(mrsReturn!�ۼ�), 0, mrsReturn!�ۼ�) * IIf(mintUnit = 0, 1, mrsReturn!����ϵ��), _
'                                    mrsReturn!ָ�������� * IIf(mintUnit = 0, 1, mrsReturn!����ϵ��), _
'                                    IIf(IsNull(mrsReturn!����), "!", mrsReturn!����), mrsReturn!���Ч��, "", _
'                                    IIf(mintUnit = 0, 1, mrsReturn!����ϵ��), IIf(IsNull(mrsReturn!����), 0, mrsReturn!����), mrsReturn!ʱ��, _
'                                    mrsReturn!���÷���, mrsReturn!ָ������� / 100, IIf(IsNull(mrsReturn!��׼�ĺ�), "", mrsReturn!��׼�ĺ�)) = False Then ' mrsReturn!����
'                             Cancel = True
'                             Exit Sub
'                         End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'                        Cancel = True
'                    End If
                    Call ��ʾ�����
                    '��ֵ��������
                    If .TextMatrix(.Row, 0) = "" Then
                        vsfCostlyInfo.Visible = False
                    Else
                        vsfCostlyInfo.Visible = IsCostly(.TextMatrix(.Row, 0))
                    End If
                    CostlyInfo_Refresh Val(.TextMatrix(.Row, 1)), vsfCostlyInfo.Visible
                    Call Form_Resize
                End If
            Case mCol����
                '�޴���
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mCol����) = ""
                    End If
                    Call ColMoveNextCol(.Col)
                    '.Col = mCol����
                    Cancel = True
                    Exit Sub
                Else
                    Dim rs���� As New Recordset
                    
                    .Text = UCase(Trim(.Text))
                    strKey = Trim(.Text)
                    
                    gstrSQL = "" & _
                        "   Select ����,����,���� From ���������� " & _
                        "   Where upper(����) like [1] or Upper(����) like [1] or Upper(����) like [1]"
                    
                    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%")
                    
                    
                    If rs����.EOF Then
                        If MsgBox("��������������û���ҵ�������Ĳ��أ���Ҫ��������������������������", vbYesNo + vbQuestion, mstrCaption) = vbNo Then
                            Cancel = True
                            Exit Sub
                        Else
                            Dim rsMax As New Recordset
                            Dim int���� As Integer, strCode As String, strSpecify As String
                            
                            If rsMax.State = 1 Then rsMax.Close
                            gstrSQL = "SELECT Nvl(MAX(LENGTH(����)),2) As Length FROM ����������"
                            zlDatabase.OpenRecordset rsMax, gstrSQL, mstrCaption
                            int���� = rsMax!Length
                            
                            gstrSQL = "SELECT Nvl(MAX(LPAD(����," & int���� & ",'0')),'00') As Code FROM ����������"
                            rsMax.Close
                            zlDatabase.OpenRecordset rsMax, gstrSQL, mstrCaption
                            strCode = rsMax!Code
                            
                            int���� = Len(strCode)
                            strCode = strCode + 1
                            
                            If int���� >= Len(strCode) Then
                                strCode = String(int���� - Len(strCode), "0") & strCode
                            End If
                            strSpecify = zlStr.GetCodeByVB(strKey)
                            
                            
                            gstrSQL = "ZL_����������_INSERT('" & strCode & "','" & strKey & "','" & strSpecify & "')"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        End If
                    Else
                        If rs����.RecordCount = 1 Then
                            If CheckQualifications(mlngModule, 1, rs����.Fields("����")) = False Then
                                Exit Sub
                            End If
                            
                            .TextMatrix(.Row, mCol����) = rs����.Fields("����")
                            .Text = rs����.Fields("����")
                            
                            gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
                            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(.Row, mCol����), .TextMatrix(mshBill.Row, 0))
                            If rstemp.RecordCount > 0 Then
                                .TextMatrix(.Row, mcol��׼�ĺ�) = IIf(IsNull(rstemp!��׼�ĺ�), "", rstemp!��׼�ĺ�)
                            Else
                                .TextMatrix(.Row, mcol��׼�ĺ�) = ""
                            End If
                        Else
                            Set msh����.Recordset = rs����
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
                OS.OpenIme False
            
            Case mCol���ս���
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mCol���ս���) = ""
                    End If
                    Call ColMoveNextCol(.Col)
                    Cancel = True
                    Exit Sub
                Else
                    Dim rs���� As New Recordset
                    
                    gstrSQL = "" & _
                        "   Select ����,���� From ������ս��� " & _
                        "   Where upper(����) like [1] or Upper(����) like [1] "
                    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%")
                    
                    If rs����.EOF Then
                        MsgBox "������ս���û���ҵ������������룡", vbInformation, mstrCaption
                        Cancel = True
                        Exit Sub
                    Else
                        If rs����.RecordCount = 1 Then
                            .TextMatrix(.Row, mCol���ս���) = rs����.Fields("����")
                            .Text = rs����.Fields("����")
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
                OS.OpenIme False
            Case mcol��׼�ĺ�
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mcol��׼�ĺ�) = ""
                    End If
                    Call ColMoveNextCol(.Col)
                    Cancel = True
                    Exit Sub
                End If
            Case mCol����
                If Len(strKey) > mintBatchNoLen Then
                    ShowMsgBox "���Ų��ܴ���" & mintBatchNoLen & "λ"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
               
                '�޴���
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mCol����) = ""
                    End If
                    Call ColMoveNextCol(.Col)
                    '.Col = mcol��������
                    Cancel = True
                    Exit Sub
                End If
            Case mcolע��֤��
                If LenB(StrConv(strKey, vbFromUnicode)) > 50 Then
                    ShowMsgBox "ע��֤�Ų��ܴ���50���ַ���25������,����!"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                '�޴���
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .Text = " "
                    Else
                        If Trim(.TextMatrix(.Row, mcolע��֤��)) = "" Then
                            .TextMatrix(.Row, mcolע��֤��) = " "
                        End If
                        
                    End If
                    Exit Sub
                End If
            Case mcol��Ʒ����
                If Len(Trim(strKey)) > 50 Then
                    ShowMsgBox "��Ʒ���벻�ܴ���50���ַ�,����!"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                '�޴���
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .Text = " "
                    Else
                        If Trim(.TextMatrix(.Row, mcol��Ʒ����)) = "" Then
                            .TextMatrix(.Row, mcol��Ʒ����) = " "
                        End If
                        
                    End If
                    Exit Sub
                End If
                .Text = UCase(.Text)
            Case mColЧ��
                '�д���
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "ʧЧ�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "ʧЧ�ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mColЧ��) Then
                
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    
                    Exit Sub
                End If
            Case mcol��������
                '�д���
                Dim str���� As String, strxq As String
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "�������ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "�������ڱ���Ϊ�������磨2000-10-10����20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                
                    If .ColData(mColЧ��) = 5 Then
                        If Not IsDate(Trim(strKey)) Then
                            str���� = ""
                        Else
                            str���� = Format(strKey, "yyyymmdd")
                        End If
                        If str���� <> "" And Trim(.TextMatrix(.Row, mColԭ����)) <> "" Then
                            '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                            If IsNumeric(str����) And Split(.TextMatrix(.Row, mColԭ����), "||")(0) <> "0" Then
                                strxq = UCase(str����)
                                If Trim(.TextMatrix(.Row, mColЧ��)) = "" Then
                                    If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                        strxq = TranNumToDate(strxq, True)
                                        If strxq = "" Then Exit Sub
                                        
                                        .TextMatrix(.Row, mColЧ��) = Format(DateAdd("M", Split(.TextMatrix(.Row, mColԭ����), "||")(0), strxq), "yyyy-mm-dd")
                                        Call CheckLapse(.TextMatrix(.Row, mColЧ��))
                                    End If
                                End If
                            End If
                        End If
                    End If
                
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mcol��������) Then
                
                    If .ColData(mColЧ��) = 5 And .TextMatrix(.Row, mcol��������) <> "" Then
                        If Not IsDate(Trim(.TextMatrix(.Row, mcol��������))) Then
                            str���� = ""
                        Else
                            str���� = Format(.TextMatrix(.Row, mcol��������), "yyyymmdd")
                        End If
                        If str���� <> "" And Trim(.TextMatrix(.Row, mColԭ����)) <> "" Then
                            '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                            If IsNumeric(str����) And Split(.TextMatrix(.Row, mColԭ����), "||")(0) <> "0" Then
                                strxq = UCase(str����)
                                If Trim(.TextMatrix(.Row, mColЧ��)) = "" Then
                                    If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                        strxq = TranNumToDate(strxq, True)
                                        If strxq = "" Then Exit Sub
                                        
                                        .TextMatrix(.Row, mColЧ��) = Format(DateAdd("M", Split(.TextMatrix(.Row, mColԭ����), "||")(0), strxq), "yyyy-mm-dd")
                                        Call CheckLapse(.TextMatrix(.Row, mColЧ��))
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    Else
                        
                    End If
                    
                    Exit Sub
                End If

            Case mCol����
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "���ʱ���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" And strKey <> .TextMatrix(.Row, mCol����) Then
                    SetDisCount .Row, strKey
                End If
                
                Call ���ɱ���
                Call ��ʾ�ϼƽ��
            Case mcol�������
                '�д���
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "������ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                        'Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "������ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                                
                        Exit Sub
                    End If
                    If Format(sys.Currentdate, "yyyy-mm-dd") >= Format(DateAdd("m", Val(.TextMatrix(.Row, mcol���Ч��)), CDate(strKey)), "yyyy-mm-dd") Then
                        If MsgBox("�����������Ѿ��������ʧЧ��(" & Format(DateAdd("m", Val(.TextMatrix(.Row, mcol���Ч��)), CDate(strKey)), "yyyy-mm-dd") & "),�Ƿ�Ҫ�������!", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    .Text = strKey
                    '����ʧЧ��
                    .TextMatrix(.Row, mcol���ʧЧ��) = Format(DateAdd("m", Val(.TextMatrix(.Row, mcol���Ч��)), CDate(strKey)), "yyyy-mm-dd")
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mcol�������) Then
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    Exit Sub
                End If
                
            Case mcol���ʧЧ��
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "���ʧЧ�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                        'Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "���ʧЧ�ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Format(sys.Currentdate, "yyyy-mm-dd") >= Format(DateAdd("m", Val(.TextMatrix(.Row, mcol���Ч��)), CDate(strKey)), "yyyy-mm-dd") Then
                        If MsgBox("�����������Ѿ��������ʧЧ��(" & Format(DateAdd("m", Val(.TextMatrix(.Row, mcol���Ч��)), CDate(strKey)), "yyyy-mm-dd") & "),�Ƿ�Ҫ�������!", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If

                    .Text = strKey
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mcol���ʧЧ��) Then
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    Exit Sub
                End If
                
            Case mColָ��������
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "ָ�������۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" And strKey <> .TextMatrix(.Row, mColָ��������) Then
                    SetDisCount .Row, strKey
                End If
                
                Call ���ɱ���
                Call ��ʾ�ϼƽ��
            Case mCol�ɹ���
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�ɹ��۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "�ɹ��۱������0,�����䣡", vbInformation + vbOKOnly, gstrSysName
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
                    .Text = Format(strKey, mFMT.FM_�ɱ���)
                    .TextMatrix(.Row, .Col) = .Text
                    If mbln��ǿ�ƿ���ָ���۸� = False Then
                        If Val(strKey) > Val(.TextMatrix(.Row, mColָ��������)) Then
                            MsgBox "������Ĳɹ���(" & strKey & ")�����˲ɹ��޼�(" & .TextMatrix(.Row, mColָ��������) & ")��", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    '�б�ɱ����ж�
                    dblCostPrice = Get�б굥λ�ɱ���(.TextMatrix(.Row, lng����ID))
                    dblPrice = CDbl(IIf(.Text <> "", .Text, IIf(.TextMatrix(.Row, mCol�ɹ���) = "", 0, .TextMatrix(.Row, mCol�ɹ���))))
                    If dblCostPrice < dblPrice And dblCostPrice <> 0 Then
                        strBidMess = zlDatabase.GetPara("��ⵥ�۳��б굥��", glngSys, mlngModule)
                        If Val(strBidMess) = 0 Then     '��ֹ��ⵥ�۳��б굥��
                            MsgBox "��ֹ�ɹ��ۣ�" & dblPrice & "���� �б굥�ۣ�" & dblCostPrice & "����", vbCritical, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        ElseIf Val(strBidMess) = 1 Then '��ʾ
                            MsgBox "�ɹ��ۣ�" & dblPrice & "���� �б굥�ۣ�" & dblCostPrice & "����", vbInformation, gstrSysName
                        End If
                    End If
                End If
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_�ɱ���)
                    .Text = strKey
                    .TextMatrix(.Row, mCol�ɹ���) = .Text
                End If
                .TextMatrix(.Row, mCol�����) = Format(Val(.TextMatrix(.Row, mCol�ɹ���)) * Val(.TextMatrix(.Row, mCol����)) / 100, mFMT.FM_�ɱ���)
                
                If ISCheckScalc�ۼ�(False, .Row) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                End If
                Call ���ɱ���
                Call ��ʾ�ϼƽ��
            Case mCol�����
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "����۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "����۱������0,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "����۱���С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = Format(strKey, mFMT.FM_�ɱ���)
                    .TextMatrix(.Row, .Col) = .Text
                End If
                '�б�ɱ����ж�
                dblCostPrice = Get�б굥λ�ɱ���(.TextMatrix(.Row, lng����ID))
                dblPrice = CDbl(IIf(.Text <> "", .Text, IIf(.TextMatrix(.Row, mCol�ɹ���) = "", 0, .TextMatrix(.Row, mCol�ɹ���))))
                If dblCostPrice < dblPrice And dblCostPrice <> 0 Then
                    strBidMess = zlDatabase.GetPara("��ⵥ�۳��б굥��", glngSys, mlngModule)
                    If Val(strBidMess) = 0 Then     '��ֹ��ⵥ�۳��б굥��
                        MsgBox "��ֹ�ɹ��ۣ�" & dblPrice & "���� �б굥�ۣ�" & dblCostPrice & "����", vbCritical, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    ElseIf Val(strBidMess) = 1 Then '��ʾ
                        MsgBox "�ɹ��ۣ�" & dblPrice & "���� �б굥�ۣ�" & dblCostPrice & "����", vbInformation, gstrSysName
                    End If
                End If
                
                '�������ÿ���
                If Val(.TextMatrix(.Row, mCol����)) = 0 Then
                
                    If strKey <> "" And Val(.TextMatrix(.Row, mColָ��������)) <> 0 Then
                        .TextMatrix(.Row, mCol����) = Format((strKey / .TextMatrix(.Row, mColָ��������)) * 100, mFMT.FM_�ɱ���)
                    Else
                        .TextMatrix(.Row, mCol����) = 100
                    End If
                End If
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_�ɱ���)
                    .Text = strKey
                    .TextMatrix(.Row, mCol�����) = .Text
                End If
                If Val(.TextMatrix(.Row, mCol�ɹ���)) <> 0 Then
                    .TextMatrix(.Row, mCol�ɹ���) = Format((Val(.TextMatrix(.Row, mCol�����)) / .TextMatrix(.Row, mCol����)) * 100, mFMT.FM_�ɱ���)
                Else
                    .TextMatrix(.Row, mCol�ɹ���) = Format(Val(.TextMatrix(.Row, mCol�����)), mFMT.FM_�ɱ���)
                    .TextMatrix(.Row, mCol����) = "100"
                End If
        
                If ISCheckScalc�ۼ�(True, .Row) = False Then
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If

                Call ���ɱ���
                Call ��ʾ�ϼƽ��
            Case mCol������
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "���������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Abs(Val(strKey)) < 0 Then
                        MsgBox "������ı������0,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 14 - 1 Then
                        MsgBox "���������С��" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                
                dbl���� = Val(.TextMatrix(.Row, mCol����))
                dbl���� = Val(.TextMatrix(.Row, mCol����))
                dbl�ۼ� = Val(.TextMatrix(.Row, mCol�ۼ�))
                
                lng����ID = Val(.TextMatrix(.Row, 0))
                If strKey <> "" And strKey <> .TextMatrix(.Row, mCol������) Then
                    If dbl���� <> 0 Then
                        dbl����� = Val(strKey) / dbl����
                        dbl�ɹ��� = dbl����� * 100 / IIf(dbl���� = 0, 1, dbl����)
                        .TextMatrix(.Row, mCol�����) = Format(dbl�����, mFMT.FM_�ɱ���)
                        .TextMatrix(.Row, mCol�ɹ���) = Format(dbl�ɹ���, mFMT.FM_�ɱ���)
                        dbl�ɱ��� = IIf(mblnʱ�۹�ǰ���� = True, dbl�ɹ���, dbl�����)
                        
                        If mbln�Ӽ��� = True Then
                            'ȡ�øı������ǰ�ļӼ���
                            mdbl�Ӽ��� = 15
                            If dbl�ۼ� <> 0 And dbl����� <> 0 Then
                                If Val(Replace(.TextMatrix(.Row, mcol�ӳ���), "%", "")) >= 0 Then
                                    mdbl�Ӽ��� = Val(Replace(.TextMatrix(.Row, mcol�ӳ���), "%", ""))
                                Else
                                    mdbl�Ӽ��� = ����ӳ���(lng����ID, dbl�ۼ�, dbl�ɱ���)
                                End If
                            End If
                        End If
                        
                        '��ʱ�۲��ϵĴ���
                        If .TextMatrix(.Row, mColԭ����) <> "" Then
                            '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                            
                            '���¼������ۼۡ����
                            If Split(.TextMatrix(.Row, mColԭ����), "||")(2) = 1 Then
                                '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                                If mbln�Ӽ��� = True Then
                                    If mint�༭״̬ = 8 And dbl�ۼ� <> 0 Then
                                    Else
                                        .TextMatrix(.Row, mCol�ۼ�) = Format(У�����ۼ�(dbl�ɱ��� * (1 + (mdbl�Ӽ��� / 100)) + _
                                        ʱ�۲������ۼ�(lng����ID, dbl�ɱ���, (mdbl�Ӽ��� / 100))), mFMT.FM_���ۼ�)
                                    End If
                                    .TextMatrix(.Row, mCol�ۼ۽��) = Format(Val(.TextMatrix(.Row, mCol�ۼ�)) * dbl����, mFMT.FM_���)
                                    .TextMatrix(.Row, mCol���) = Format(IIf(.TextMatrix(.Row, mCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mCol������) = "", 0, .TextMatrix(.Row, mCol������)), mFMT.FM_���)
                                    '���˺�:���ۼ۴���
                                    Call �������ۼۼ����۲��(.Row)
                                ElseIf mbln�ֶμӳ��� = True Then
                                    dbl�ӳ��� = 0
                                    If mint�༭״̬ = 8 And dbl�ۼ� <> 0 Then
                                    Else
                                        If Get�ֶμӳ��ۼ�(dbl�ɱ���, Val(.TextMatrix(.Row, mCol����ϵ��)), mstrCaption, sng�ֶ��ۼ�) = False Then
                                            Cancel = True
                                            .TxtSetFocus
                                            Exit Sub
                                        End If
                                        .TextMatrix(.Row, mCol�ۼ�) = Format(У�����ۼ�(sng�ֶ��ۼ� + _
                                                                      ʱ�۲������ۼ�(lng����ID, dbl�ɱ���, dbl�ӳ���, -1, sng�ֶ��ۼ�)) _
                                                                      , mFMT.FM_���ۼ�)
                                    End If
                                    .TextMatrix(.Row, mCol�ۼ۽��) = Format(dbl���� * Val(.TextMatrix(.Row, mCol�ۼ�)), mFMT.FM_���)
                                    '���˺�:���ۼ۴���
                                    Call �������ۼۼ����۲��(.Row)
                                Else 'mblnʱ������ȡ�ϴ��ۼ� = True����3��ȡ�ۼ۷�ʽ��û������ʱ�����ȴ��ϴ�ȡ�����û�����ռӳ��ʷ�ʽȡ
                                    If mblnʱ������ȡ�ϴ��ۼ� = True Then
                                        gstrSQL = "Select Nvl(�ϴ��ۼ�, 0) As �ϴ��ۼ� From �������� Where ����id = [1]"
                                        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID)
                                        If rstemp!�ϴ��ۼ� > 0 Then
                                            .TextMatrix(.Row, mCol�ۼ�) = Format(zlStr.Nvl(rstemp!�ϴ��ۼ�, 0) * Val(.TextMatrix(.Row, mCol����ϵ��)), mFMT.FM_���ۼ�)
                                            If dbl�ɱ��� <> 0 Then
                                                .TextMatrix(.Row, mcol�ӳ���) = Format((Val(.TextMatrix(.Row, mCol�ۼ�)) / dbl�ɱ��� - 1) * 100, "###0.00") & "%"
                                            End If
                                        Else
                                            '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                                            If dbl�ɱ��� <> 0 Then
                                                mdbl�Ӽ��� = Val(Replace(.TextMatrix(.Row, mcol�ӳ���), "%", "")) '����ӳ���(lng����ID, dbl�ۼ�, dbl�ɱ���)
                                                .TextMatrix(.Row, mcol�ӳ���) = Format(mdbl�Ӽ���, "####0.00") & "%"
                                            End If
                                            If mint�༭״̬ = 8 And dbl�ۼ� <> 0 Then
                                            Else
                                                .TextMatrix(.Row, mCol�ۼ�) = Format(У�����ۼ�(dbl�ɱ��� * (1 + (mdbl�Ӽ��� / 100)) + _
                                                ʱ�۲������ۼ�(lng����ID, dbl�ɱ���, (mdbl�Ӽ��� / 100))), mFMT.FM_���ۼ�)
                                            End If
                                        End If
                                    Else
                                        If dbl�ɱ��� <> 0 Then
                                            mdbl�Ӽ��� = Val(Replace(.TextMatrix(.Row, mcol�ӳ���), "%", "")) '����ӳ���(lng����ID, dbl�ۼ�, dbl�ɱ���)
                                            .TextMatrix(.Row, mcol�ӳ���) = Format(mdbl�Ӽ���, "####0.00") & "%"
                                        End If
                                        If mint�༭״̬ = 8 And dbl�ۼ� <> 0 Then
                                        Else
                                            .TextMatrix(.Row, mCol�ۼ�) = Format(У�����ۼ�(dbl�ɱ��� * (1 + (mdbl�Ӽ��� / 100)) + _
                                            ʱ�۲������ۼ�(lng����ID, dbl�ɱ���, (mdbl�Ӽ��� / 100))), mFMT.FM_���ۼ�)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                        
                    If .TextMatrix(.Row, mColָ��������) <> "" Then
                        If Val(.TextMatrix(.Row, mColָ��������)) = 0 Then
                             .TextMatrix(.Row, mCol����) = 100
'                        Else
'                             .TextMatrix(.Row, mCol����) = Format(Val(.TextMatrix(.Row, mCol�����)) / Val(.TextMatrix(.Row, mColָ��������)) * 100, mFMT.FM_�ɱ���)
                        End If
                    End If
                    If strKey <> "" Then
                         .Text = Format(strKey, mFMT.FM_���)

                    End If
                    .TextMatrix(.Row, mCol��Ʊ���) = IIf(Trim(.TextMatrix(.Row, mCol��Ʊ��)) = "" And Trim(.TextMatrix(.Row, mcol��Ʊ����)) = "", "", Format(strKey, mFMT.FM_���))
                    .TextMatrix(.Row, mCol���) = Format(IIf(.TextMatrix(.Row, mCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mCol�ۼ۽��)) - strKey, mFMT.FM_���)
                    .TextMatrix(.Row, mCol������) = Format(strKey, mFMT.FM_���)
                    '���˺�:���ۼ۴���
                    Call �������ۼۼ����۲��(.Row)
                End If
                
                Call ���ɱ���
                Call ��ʾ�ϼƽ��
                
            Case mCol����
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "�����������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "��������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) = 0 Then
                        MsgBox "��������Ϊ��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Abs(Val(strKey)) < 0.001 Then
                            MsgBox "�����ı������0.001,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "��������С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = Format(strKey, mFMT.FM_����)
                    '����Ƿ����㹻�Ŀ������˻�
                    If mint�༭״̬ = 8 Or mbln�˻� Then
                        If Not CheckStock(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mCol����)), Val(.Text) * Val(.TextMatrix(.Row, mCol����ϵ��))) Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    .Text = strKey
                    lng����ID = Val(.TextMatrix(.Row, 0))
                    dbl���� = Val(strKey)
                    dbl����� = Val(.TextMatrix(.Row, mCol�����))
                    dbl�ɹ��� = Val(.TextMatrix(.Row, mCol�ɹ���))
                    dbl�ۼ� = Val(.TextMatrix(.Row, mCol�ۼ�))
                    dbl�ɱ��� = IIf(mblnʱ�۹�ǰ���� = True, dbl�ɹ���, dbl�����)
                    
                    .TextMatrix(.Row, mCol������) = Format(dbl����� * Val(strKey), mFMT.FM_���)
                    
                    'ʱ�۲��ϵĴ���
                    If .TextMatrix(.Row, mColԭ����) <> "" Then
                        '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                        If Split(.TextMatrix(.Row, mColԭ����), "||")(2) = 1 Then
                            '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                            If mbln�Ӽ��� = True Then
                                mdbl�Ӽ��� = 15
                                
                                If dbl�ɱ��� <> 0 Then
                                    mdbl�Ӽ��� = ����ӳ���(lng����ID, dbl�ۼ�, dbl�ɱ���)
                                    .TextMatrix(.Row, mcol�ӳ���) = zlStr.FormatEx(mdbl�Ӽ���, 2) & "%"
                                End If
                                If mint�༭״̬ = 8 And dbl�ۼ� <> 0 Then
                                Else
                                    .TextMatrix(.Row, mCol�ۼ�) = Format(У�����ۼ�(dbl�ɱ��� * (1 + (mdbl�Ӽ��� / 100)) + _
                                    ʱ�۲������ۼ�(lng����ID, dbl�ɱ���, (mdbl�Ӽ��� / 100))), mFMT.FM_���ۼ�)
                                End If
                                .TextMatrix(.Row, mCol�ۼ۽��) = Format(Val(.TextMatrix(.Row, mCol�ۼ�)) * strKey, mFMT.FM_���)
                                .TextMatrix(.Row, mCol���) = Format(IIf(.TextMatrix(.Row, mCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mCol������) = "", 0, .TextMatrix(.Row, mCol������)), mFMT.FM_���)
                                '���˺�:���ۼ۴���
                                Call �������ۼۼ����۲��(.Row)
                            ElseIf mbln�ֶμӳ��� = True Then
                                dbl�ӳ��� = 0
                                If mint�༭״̬ = 8 And dbl�ۼ� <> 0 Then
                                Else
                                    If Get�ֶμӳ��ۼ�(dbl�ɱ���, Val(.TextMatrix(.Row, mCol����ϵ��)), mstrCaption, sng�ֶ��ۼ�) = False Then
                                        Cancel = True
                                        .TxtSetFocus
                                        Exit Sub
                                    End If
                                    .TextMatrix(.Row, mCol�ۼ�) = Format(У�����ۼ�(sng�ֶ��ۼ� + _
                                                                  ʱ�۲������ۼ�(lng����ID, dbl�ɱ���, dbl�ӳ���, -1, sng�ֶ��ۼ�)) _
                                                                  , mFMT.FM_���ۼ�)
                                End If
                                .TextMatrix(.Row, mcol�ӳ���) = Format(dbl�ӳ��� * 100, "####0.00") & "%" '��Ϊ�Ƿֶμӳɵ����Լӳ��ʲ�׼ȷ��ȡһ��ģ��ֵ����
                            Else  'mblnʱ������ȡ�ϴ��ۼ� = True����3��ȡ�ۼ۷�ʽ��û������ʱ�����ȴ��ϴ�ȡ�����û�����ռӳ��ʷ�ʽȡ
                                If mblnʱ������ȡ�ϴ��ۼ� = True Then
                                    gstrSQL = "Select Nvl(�ϴ��ۼ�, 0) As �ϴ��ۼ� From �������� Where ����id = [1]"
                                    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID)
                                    If rstemp!�ϴ��ۼ� > 0 Then
                                        .TextMatrix(.Row, mCol�ۼ�) = Format(zlStr.Nvl(rstemp!�ϴ��ۼ�, 0) * Val(.TextMatrix(.Row, mCol����ϵ��)), mFMT.FM_���ۼ�)
                                        If dbl�ɱ��� <> 0 Then
                                            .TextMatrix(.Row, mcol�ӳ���) = Format((Val(.TextMatrix(.Row, mCol�ۼ�)) / dbl�ɱ��� - 1) * 100, "###0.00") & "%"
                                        End If
                                    Else
                                        '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                                        If dbl�ɱ��� <> 0 Then
                                            mdbl�Ӽ��� = Val(Replace(.TextMatrix(.Row, mcol�ӳ���), "%", "")) '����ӳ���(lng����ID, dbl�ۼ�, dbl�ɱ���)
                                            .TextMatrix(.Row, mcol�ӳ���) = Format(mdbl�Ӽ���, "####0.00") & "%"
                                        End If
                                        If mint�༭״̬ = 8 And dbl�ۼ� <> 0 Then
                                        Else
                                            .TextMatrix(.Row, mCol�ۼ�) = Format(У�����ۼ�(dbl�ɱ��� * (1 + (mdbl�Ӽ��� / 100)) + _
                                            ʱ�۲������ۼ�(lng����ID, dbl�ɱ���, (mdbl�Ӽ��� / 100))), mFMT.FM_���ۼ�)
                                        End If
                                    End If
                                Else
                                    If dbl�ɱ��� <> 0 Then
                                        mdbl�Ӽ��� = Val(Replace(.TextMatrix(.Row, mcol�ӳ���), "%", "")) '����ӳ���(lng����ID, dbl�ۼ�, dbl�ɱ���)
                                        .TextMatrix(.Row, mcol�ӳ���) = Format(mdbl�Ӽ���, "####0.00") & "%"
                                    End If
                                    If mint�༭״̬ = 8 And dbl�ۼ� <> 0 Then
                                    Else
                                        .TextMatrix(.Row, mCol�ۼ�) = Format(У�����ۼ�(dbl�ɱ��� * (1 + (mdbl�Ӽ��� / 100)) + _
                                        ʱ�۲������ۼ�(lng����ID, dbl�ɱ���, (mdbl�Ӽ��� / 100))), mFMT.FM_���ۼ�)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(.Row, mCol�ۼ�)) <> 0 Then
                        .TextMatrix(.Row, mCol�ۼ۽��) = Format(Val(.TextMatrix(.Row, mCol�ۼ�)) * Val(strKey), mFMT.FM_���)
                    End If
                    .TextMatrix(.Row, mCol���) = Format(IIf(.TextMatrix(.Row, mCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mCol������) = "", 0, .TextMatrix(.Row, mCol������)), mFMT.FM_���)
                    .TextMatrix(.Row, .Col) = strKey
                    '���˺�:���ۼ۴���
                    Call �������ۼۼ����۲��(.Row)
                    If mint�༭״̬ = 8 Or (mint�༭״̬ = 2 And mbln�˻� = True) Then
                        .TextMatrix(.Row, mCol��Ʊ���) = IIf(Trim(.TextMatrix(.Row, mCol��Ʊ��)) = "" And Trim(.TextMatrix(.Row, mcol��Ʊ����)) = "", "", .TextMatrix(.Row, mCol�ۼ۽��))
                    Else
                        .TextMatrix(.Row, mCol��Ʊ���) = IIf(Trim(.TextMatrix(.Row, mCol��Ʊ��)) = "" And Trim(.TextMatrix(.Row, mcol��Ʊ����)) = "", "", .TextMatrix(.Row, mCol������))
                    End If
                End If
                ��ʾ�ϼƽ��
            Case mCol��������
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "�����������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "��������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Abs(Val(strKey)) > Abs(Val(.TextMatrix(.Row, mCol����))) Then
                        MsgBox "���������ľ���ֵ���ܴ���ԭ�������ľ���ֵ,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "������������С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = Format(strKey, mFMT.FM_����)
                    .Text = strKey
                    
                    If Val(.TextMatrix(.Row, mCol�����)) <> 0 Then
                        .TextMatrix(.Row, mCol������) = Format(Val(.TextMatrix(.Row, mCol�����)) * strKey, mFMT.FM_���)
                    End If
                    If .TextMatrix(.Row, mCol�ۼ�) <> "" Then
                        .TextMatrix(.Row, mCol�ۼ۽��) = Format(Val(.TextMatrix(.Row, mCol�ۼ�)) * strKey, mFMT.FM_���)
                    End If
                    .TextMatrix(.Row, mCol���) = Format(IIf(.TextMatrix(.Row, mCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mCol������) = "", 0, .TextMatrix(.Row, mCol������)), mFMT.FM_���)
                    '���˺�:���ۼ۴���
                    Call �������ۼۼ����۲��(.Row, False)
                    If Trim(.TextMatrix(.Row, mCol��Ʊ��)) <> "" Or Trim(.TextMatrix(.Row, mCol�������)) <> "" Or Trim(.TextMatrix(.Row, mcol��Ʊ����)) <> "" Then
                    
                        dbl��Ʊ��� = GetTotale��Ʊ���(mstr���ݺ�, Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mCol���)))
                        If Val(.TextMatrix(.Row, mCol����)) = 0 Then
                            .TextMatrix(.Row, mCol��Ʊ���) = Format(0, mFMT.FM_���)
                        Else
                            .TextMatrix(.Row, mCol��Ʊ���) = Format(Val(strKey) / Val(.TextMatrix(.Row, mCol����)) * dbl��Ʊ���, mFMT.FM_���)
                        End If
                    End If
                End If
                
                ��ʾ�ϼƽ��
                
            Case mCol��Ʊ��
                
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .ColData(mCol��Ʊ����) = 5
                        .ColData(mCol��Ʊ���) = 5
                        .ColData(mcol��Ʊ����) = 5
                        .TextMatrix(.Row, mCol��Ʊ���) = ""
                        .TextMatrix(.Row, mCol��Ʊ����) = ""
                        .TextMatrix(.Row, mcol��Ʊ����) = ""
                        .TextMatrix(.Row, .Col) = " "
                        .Text = " "
                    ElseIf .TxtVisible = False Then
                        If Trim(.TextMatrix(.Row, mCol��Ʊ��)) = "" Then
                            .ColData(mCol��Ʊ����) = 5
                            .ColData(mCol��Ʊ���) = 5
                            .ColData(mcol��Ʊ����) = 5
                            .TextMatrix(.Row, mCol��Ʊ���) = ""
                            .TextMatrix(.Row, mCol��Ʊ����) = ""
                            .TextMatrix(.Row, mcol��Ʊ����) = ""
                            .TextMatrix(.Row, .Col) = " "
                            .Text = " "
                        Else
                           .Text = .TextMatrix(.Row, .Col)
                           .ColData(mCol��Ʊ����) = 2
                           .ColData(mcol��Ʊ����) = 4
                           .ColData(mCol��Ʊ���) = 4
                                    
                        End If
                    End If
                Else
                    If zlCommFun.ActualLen(.Text) > mint��Ʊ��Len Then
                        ShowMsgBox "��Ʊ�����ֻ������" & mint��Ʊ��Len & "���ַ�!"
                            Cancel = True
                        Exit Sub
                    End If

                    .ColData(mCol��Ʊ����) = 2
                    .ColData(mcol��Ʊ����) = 4
                    .ColData(mCol��Ʊ���) = 4
                   
                    If mint�༭״̬ = 8 Or (mint�༭״̬ = 3 And mbln�˻� = True) Then
                        .TextMatrix(.Row, mCol��Ʊ���) = .TextMatrix(.Row, mCol������)
                    Else
                        If mint��¼״̬ <> 1 Then
                            If Val(.TextMatrix(.Row, mCol��Ʊ���)) = 0 Then
                                .TextMatrix(.Row, mCol��Ʊ���) = .TextMatrix(.Row, mCol������)
                            End If
                        Else
                            .TextMatrix(.Row, mCol��Ʊ���) = .TextMatrix(.Row, mCol������)
                        End If
                    End If
                End If
                ��ʾ�ϼƽ��
                Exit Sub
            Case mcol��Ʊ����
                If Trim(.Text) = "" Then
                    If mcol��Ʊ���� <> mintLastCol Then
                        .Col = GetNextEnableCol(mcol��Ʊ����)
                        .Text = ""
                        Cancel = True
                        Exit Sub
                    End If
                Else
                    If zlCommFun.ActualLen(.Text) > 20 Then
                        ShowMsgBox "��Ʊ�������ֻ������" & 20 & "���ַ�!"
                            Cancel = True
                        Exit Sub
                    End If
                End If
                Exit Sub
            Case mCol��Ʊ���
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "��Ʊ������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Abs(Val(strKey)) < 0 Then
                        MsgBox "��Ʊ���������0,�����䣡", vbInformation + vbOKOnly, gstrSysName
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
                    strKey = Format(Val(strKey), mFMT.FM_���)
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
                ��ʾ�ϼƽ��
            Case mcol���ۼ�
                '�������:
                ' 1.�ۼ۲��ܴ���ָ�����ۼ�(���ݲ���:��ǿ�ƿ���ָ���۸����)
                ' 2.����˽�������ۼ�
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If strKey <> "" Then
                    If Not IsNumeric(strKey) Then
                        ShowMsgBox "���ۼ۱���Ϊ�����ͣ������䣡"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) < 0 Then
                        ShowMsgBox "���ۼ۱�����ڵ���0,�����䣡"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        ShowMsgBox "���ۼ۱���С��" & (10 ^ 11 - 1)
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If mbln��ǿ�ƿ���ָ���۸� = False Then
                        '�ж���������ۼ���ָ�����ۼ�
                        gstrSQL = "Select ָ�����ۼ� From �������� Where ����ID=[1] "
                        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[��ȡָ�����ۼ�]", Val(.TextMatrix(.Row, 0)))
                        dblָ�����ۼ� = Val(zlStr.Nvl(rstemp!ָ�����ۼ�))
                        dblָ�����ۼ� = Val(Format(dblָ�����ۼ�, mFMT.FM_ɢװ���ۼ�))
                        If Val(Format(Val(strKey), mFMT.FM_ɢװ���ۼ�)) > dblָ�����ۼ� Then
                            ShowMsgBox "���ۼ۲��ܴ���ָ�����ۼۣ�ָ�����ۼۣ���" & dblָ�����ۼ� & "��"
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If Val(.TextMatrix(.Row, mCol����ϵ��)) = 0 Then
                        dbl�ɹ��� = Val(.TextMatrix(.Row, mCol�����))
                    Else
                        dbl�ɹ��� = Val(.TextMatrix(.Row, mCol�����)) / Val(.TextMatrix(.Row, mCol����ϵ��))
                    End If
                    
                    If Val(strKey) < dbl�ɹ��� Then
                        If MsgBox("ע�⣺" & vbCrLf & "     ���ۼ�(��" & Format(Val(strKey), mFMT.FM_ɢװ���ۼ�) & " С����" & vbCrLf & "     ����ۣ���" & Format(dbl�ɹ���, mFMT.FM_�ɱ���) & "��,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                    End If
                End If
                
                If strKey <> "" Then
                    strKey = Format(Val(strKey), mFMT.FM_ɢװ���ۼ�)
                    .Text = strKey
                    .TextMatrix(.Row, .Col) = strKey
                ElseIf .TxtVisible = True Then
                    .Text = " "
                ElseIf .TxtVisible = False Then
                    If .TextMatrix(.Row, .Col) = "" Then
                        .Text = " "
                    Else
                        .Text = .TextMatrix(.Row, .Col)
                    End If
                End If
                '���˺�:���ۼ۴���
                Call �������ۼۼ����۲��(.Row, False)
                If strKey <> "" Then
                    .TextMatrix(.Row, mCol�ۼ�) = Format(Val(strKey) * Val(.TextMatrix(.Row, mCol����ϵ��)), mFMT.FM_���ۼ�)
                    .TextMatrix(.Row, mCol�ۼ۽��) = Format(Val(.TextMatrix(.Row, mCol�ۼ�)) * Val(.TextMatrix(.Row, mCol����)), mFMT.FM_���)
                    .TextMatrix(.Row, mCol���) = Format(Val(.TextMatrix(.Row, mCol�ۼ۽��)) - Val(.TextMatrix(.Row, mCol������)), mFMT.FM_���)



                End If
                ��ʾ�ϼƽ��
                
            Case mCol�ۼ�
                '�������:
                ' 1.�ۼ۲��ܴ���ָ�����ۼ�(���ݲ���:��ǿ�ƿ���ָ���۸����)
                ' 2.����˽�������ۼ�
                            
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If strKey <> "" Then
                    If Not IsNumeric(strKey) Then
                        ShowMsgBox "�ۼ۱���Ϊ�����ͣ������䣡"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) < 0 Then
                        ShowMsgBox "�ۼ۱�����ڵ���0,�����䣡"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        ShowMsgBox "�ۼ۱���С��" & (10 ^ 11 - 1)
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If mbln��ǿ�ƿ���ָ���۸� = False Then
                        '�ж���������ۼ���ָ�����ۼ�
                        gstrSQL = "Select ָ�����ۼ� From �������� Where ����ID=[1] "
                        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[��ȡָ�����ۼ�]", Val(.TextMatrix(.Row, 0)))
                        dblָ�����ۼ� = Val(zlStr.Nvl(rstemp!ָ�����ۼ�))
                        dblָ�����ۼ� = Val(Format(dblָ�����ۼ� * Val(.TextMatrix(.Row, mCol����ϵ��)), mFMT.FM_���ۼ�))
                        
                        If Val(Format(Val(strKey), mFMT.FM_���ۼ�)) > dblָ�����ۼ� Then
                            ShowMsgBox "�ۼ۲��ܴ���ָ�����ۼۣ�ָ�����ۼۣ���" & dblָ�����ۼ� & "��"
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    If Val(strKey) < Val(.TextMatrix(.Row, mCol�����)) Then
                        If MsgBox("ע�⣺" & vbCrLf & "     �ۼ�(��" & Format(Val(strKey), mFMT.FM_���ۼ�) & " С����" & vbCrLf & "     ����ۣ���" & Format(Val(.TextMatrix(.Row, mCol�����)), mFMT.FM_�ɱ���) & "��,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                    End If
                End If
                
                If strKey <> "" Then
                    strKey = Format(Val(strKey), mFMT.FM_���ۼ�)
                    .Text = strKey
                    .TextMatrix(.Row, .Col) = strKey
                ElseIf .TxtVisible = True Then
                    .Text = " "
                    .TextMatrix(.Row, .Col) = " "
                ElseIf .TxtVisible = False Then
                    If .TextMatrix(.Row, .Col) = "" Then
                        .Text = " "
                        .TextMatrix(.Row, .Col) = " "
                    Else
                        .Text = .TextMatrix(.Row, .Col)
                    End If
                End If
                If mblnʱ�۹�ǰ���� Then
                    dbl�ɹ��� = Val(.TextMatrix(.Row, mCol�ɹ���))
                Else
                    dbl�ɹ��� = Val(.TextMatrix(.Row, mCol�����))
                End If
                
                '������
                If strKey <> "" Then
                    .TextMatrix(.Row, mcol�ӳ���) = zlStr.FormatEx(����ӳ���(Val(.TextMatrix(.Row, 0)), Val(strKey), dbl�ɹ���), 2) & "%"
                    .TextMatrix(.Row, mCol�ۼ۽��) = Format(Val(strKey) * Val(.TextMatrix(.Row, mCol����)), mFMT.FM_���)
                    .TextMatrix(.Row, mCol���) = Format(Val(.TextMatrix(.Row, mCol�ۼ۽��)) - Val(.TextMatrix(.Row, mCol������)), mFMT.FM_���)
                End If
                '���˺�:���ۼ۴���
                Call �������ۼۼ����۲��(.Row)
                ��ʾ�ϼƽ��

            Case mCol��Ʊ����
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        
                        If strKey = "" Then
                            MsgBox "Ч�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "��Ʊ���ڱ���Ϊ��������(2000-10-10) �� ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            Case mCol�������
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, .Col) = " "
                        .Text = " "
                    ElseIf .TxtVisible = False Then
                        If Trim(.TextMatrix(.Row, mCol�������)) = "" Then
                            .TextMatrix(.Row, .Col) = " "
                            .Text = " "
                        Else
                           .Text = .TextMatrix(.Row, .Col)
                        End If
                    End If
                Else
                    If zlCommFun.ActualLen(.Text) > 200 Then
                        ShowMsgBox "����������ֻ������" & 200 & "���ַ�!"
                        Cancel = True
                        Exit Sub
                    End If
                End If
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetNextEnableCol(ByVal intCurrCol As Integer) As Integer
    '������һ���ɼ������õ��к�
    Dim n As Integer
    Dim intNextCol As Integer

    If intCurrCol > mshBill.Cols Or intCurrCol + 1 >= mintLastCol Then
        GetNextEnableCol = mintLastCol
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
End Function

'�Ӳ���Ŀ¼��ȡֵ��������Ӧ����
Public Function SetColValue(ByVal intRow As Integer, ByVal lng����ID As Long, ByVal str���� As String, ByVal str��� As String, _
    ByVal str���� As String, ByVal str��λ As String, ByVal num�ۼ� As Double, _
    ByVal numָ�������� As Double, ByVal strԭ���� As String, ByVal intԭЧ�� As Integer, _
    ByVal str���� As String, ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
    ByVal int�Ƿ��� As Integer, ByVal int���÷��� As Integer, ByVal dblָ������� As Double, ByVal str��׼�ĺ� As String, ByVal str��Ʒ�� As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim rsprice As New Recordset
    Dim lngDepartid As Long
    Dim dblrate As Double, dbl��������� As Double, dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    Dim rstemp As New ADODB.Recordset
    Dim int�ⷿ���� As Integer
    Dim strɢװ��λ As String
    Dim bln����ⷿ As Boolean
    Dim bln��ֵ���� As Boolean
    Dim bln���ٲ��� As Boolean
    Dim bln�������� As Boolean
    Dim bln���÷��� As Boolean
    Dim strMsg As String
    Dim sng�ֶ��ۼ� As Double
    Dim dbl���ۼӳ��� As Double
    
    On Error GoTo ErrHandle
    SetColValue = False
    With mshBill
        For intCol = 0 To .Cols - 1
            If intCol <> mCol�к� Then .TextMatrix(intRow, intCol) = ""
        Next
        
        gstrSQL = "SELECT a.�ӳ��� from �������� a where a.����id=[1]"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ӳ���", lng����ID)
        dbl�ӳ��� = Nvl(rstemp!�ӳ���, 0) / 100
        
        gstrSQL = "select count(*) rec from ��������˵�� where ����id=[1] and ��������='����ⷿ'"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��������˵��", cboStock.ItemData(cboStock.ListIndex))
        If rstemp!rec = 1 Then
            bln����ⷿ = True
        End If
        rstemp.Close
        
        gstrSQL = "SELECT nvl(A.����,0) ����,nvl(a.�ӳ���,0)/100 as �ӳ���,A.���Ч��,A.һ���Բ���,A.�ɱ���,A.�ⷿ����,A.���÷���,A.ע��֤��,B.���㵥λ ɢװ��λ" & _
                  ",Nvl(A.�Ƿ��������,0) As �������, a.��ֵ����, a.���ٲ���, a.��������,a.ע��֤��Ч�� " & _
                  "From �������� A,�շ���ĿĿ¼ B " & _
                  "Where a.����ID=b.id and A.����id=[1] "
        Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", lng����ID)
        
        dbl���ۼӳ��� = Val(rsprice!�ӳ���)
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            bln��ֵ���� = (zlStr.Nvl(rsprice!��ֵ����, 0) = 1)
            bln���ٲ��� = (zlStr.Nvl(rsprice!���ٲ���, 0) = 1)
            bln�������� = (zlStr.Nvl(rsprice!��������, 0) = 1)
            bln���÷��� = (zlStr.Nvl(rsprice!���÷���, 0) = 1)
            
            strMsg = ""
            If bln����ⷿ Then
                If bln��ֵ���� = False Then
                    strMsg = IIf(strMsg = "", "", strMsg & "��") & """��ֵ����"""
                End If
                If bln���ٲ��� = False Then
                    strMsg = IIf(strMsg = "", "", strMsg & "��") & """���ٲ���"""
                End If
                If bln�������� = False Then
                    strMsg = IIf(strMsg = "", "", strMsg & "��") & """��������"""
                End If
                If bln���÷��� = False Then
                    strMsg = IIf(strMsg = "", "", strMsg & "��") & """���÷���"""
                End If
                
                If strMsg <> "" Then
                    MsgBox "(" & str���� & ")��������������߱�" & strMsg & "�����ԡ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        .TextMatrix(intRow, mcolע��֤��Ч��) = IIf(IsNull(rsprice!ע��֤��Ч��), "", Format(rsprice!ע��֤��Ч��, "yyyy-mm-dd"))
        
        If rsprice!���� = 0 Then
            dblrate = 100
        Else
            dblrate = rsprice!����
        End If
        int�ⷿ���� = Val(zlStr.Nvl(rsprice!�ⷿ����))
        dbl�ɱ��� = rsprice!�ɱ���
        
        strɢװ��λ = zlStr.Nvl(rsprice!ɢװ��λ)
        
        .TextMatrix(intRow, mcol���Ч��) = zlStr.Nvl(rsprice!���Ч��, 0)
        .TextMatrix(intRow, mcolһ���Բ���) = zlStr.Nvl(rsprice!һ���Բ���, 0)
        .TextMatrix(intRow, mcol�������) = zlStr.Nvl(rsprice!�������, 0)
        
        .TextMatrix(intRow, mCol�к�) = intRow
        .TextMatrix(intRow, 0) = lng����ID
        .TextMatrix(intRow, mCol����) = str����
        .TextMatrix(intRow, mCol��Ʒ��) = str��Ʒ��
        .TextMatrix(intRow, mCol���) = str���
        
        If CheckQualifications(mlngModule, 1, IIf(IsNull(str����), "", str����)) = False Then
            .TextMatrix(intRow, mCol����) = ""
        Else
            .TextMatrix(intRow, mCol����) = IIf(IsNull(str����), "", str����)
        End If
        .TextMatrix(intRow, mcol��׼�ĺ�) = IIf(IsNull(str��׼�ĺ�), "", str��׼�ĺ�)
        
        .TextMatrix(intRow, mCol��λ) = str��λ
        
        .TextMatrix(intRow, mCol�ۼ�) = Format(num�ۼ�, mFMT.FM_���ۼ�)
        .TextMatrix(intRow, mColָ��������) = Format(numָ��������, mFMT.FM_�ɱ���)
        
        .TextMatrix(intRow, mColԭ����) = IIf(IsNull(strԭ����), "", strԭ����)
        .TextMatrix(intRow, mCol����) = lng����
        .TextMatrix(intRow, mcolע��֤��) = zlStr.Nvl(rsprice!ע��֤��)   'ȡĬ��ֵ
        
        
        'ȡ���ò��ϵ����ż�Ч��
        If mint�༭״̬ = 8 Or mbln�˻� Then
            gstrSQL = "" & _
                " Select �ϴ����� ����,Ч��,�ϴ���������,�ϴβɹ��� From ҩƷ���" & _
                " Where �ⷿID=[1] And ҩƷID=[2]" & _
                "       And ����=1 And nvl(����,0)=[3]"
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng����ID, lng����)
            
            If rsprice.RecordCount <> 0 Then
                .TextMatrix(intRow, mCol����) = IIf(IsNull(rsprice!����), "", rsprice!����)
                .TextMatrix(intRow, mColЧ��) = IIf(IsNull(rsprice!Ч��), "", rsprice!Ч��)
                If IsNull(rsprice!�ϴ���������) Then
                    .TextMatrix(intRow, mcol��������) = ""
                Else
                    .TextMatrix(intRow, mcol��������) = Format(rsprice!�ϴ���������, "yyyy-mm-dd")
                End If
                
                dbl�ɱ��� = zlStr.Nvl(rsprice!�ϴβɹ���, 0)
                
                If dbl�ɱ��� > 0 Then
                    .TextMatrix(intRow, mCol�ɹ���) = Format(dbl�ɱ��� * num����ϵ��, mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mCol�����) = Format(dbl�ɱ��� * num����ϵ�� * dblrate / 100, mFMT.FM_�ɱ���)
                End If
            End If
        End If
        
        'ԭЧ���ֶ����汣��ԭЧ�ڣ�ָ����ۣ��Ƿ��ۣ����÷����ȣ���ʽΪ�����Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
        .TextMatrix(intRow, mColԭ����) = IIf(IsNull(intԭЧ��), "0", intԭЧ��) & "||" & dblָ������� & "||" & int�Ƿ��� & "||" & int���÷��� & "||" & int�ⷿ����
        
        .TextMatrix(intRow, mCol����) = str����
        .TextMatrix(intRow, mCol����ϵ��) = num����ϵ��
        If intRow > 1 Then
            .TextMatrix(intRow, mCol�������) = .TextMatrix(intRow - 1, mCol�������)
            .TextMatrix(intRow, mCol��Ʊ��) = .TextMatrix(intRow - 1, mCol��Ʊ��)
            .TextMatrix(intRow, mcol��Ʊ����) = .TextMatrix(intRow - 1, mcol��Ʊ����)
            .TextMatrix(intRow, mCol��Ʊ����) = .TextMatrix(intRow - 1, mCol��Ʊ����)
        End If
        
        SetInputFormat intRow
        SetDisCount intRow, dblrate
        lngDepartid = cboStock.ItemData(cboStock.ListIndex)
        
        '˵�����������ַ�������Ͳ����������Ŀ������������ٶȡ�
        '�������Բ�����Щ��ֱ���õ�һ��SQL���ʵ�֣��������������ľͶ������ݿ���ɨ��һ�Ρ�
        
        If Not (mint�༭״̬ = 8 Or mbln�˻�) Then
            '�Զ��۲ɹ�������ȡ�ϴεĽ���ۺͿ���
                        
            '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
            If Val(Split(.TextMatrix(intRow, mColԭ����), "||")(4)) > 0 Then
                gstrSQL = "" & _
                    "   Select �ϴβɹ���,�ϴβ���,�ϴ��������� " & _
                    "   From ҩƷ��� " & _
                    "   where ����=1 and �ⷿid=[1] and ҩƷid=" & lng����ID & _
                    "       and nvl(����,0) =(  Select max(nvl(����,0)) " & _
                    "                            From ҩƷ��� " & _
                    "                           Where ����=1 and �ⷿid=[1]" & _
                    "                               and ҩƷid=[2] )"
            Else
                gstrSQL = "select �ϴβɹ���,�ϴβ���,�ϴ��������� from ҩƷ��� where ����=1 and �ⷿid= [1] and ҩƷid=[2]"
            End If
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngDepartid, lng����ID)
            
            If Not rsprice.EOF Then
                If .TextMatrix(intRow, mCol����) = "" Then
                    If CheckQualifications(mlngModule, 1, IIf(IsNull(rsprice.Fields(1)), "", rsprice.Fields(1))) = False Then
                        .TextMatrix(intRow, mCol����) = ""
                    Else
                        .TextMatrix(intRow, mCol����) = IIf(IsNull(rsprice.Fields(1)), "", rsprice.Fields(1))
                    End If
                End If
                If IsNull(rsprice!�ϴ���������) Then
                    .TextMatrix(intRow, mcol��������) = ""
                Else
                    .TextMatrix(intRow, mcol��������) = Format(rsprice!�ϴ���������, "yyyy-mm-dd")
                    If ProduceDateCheck(.TextMatrix(intRow, mcol��������)) = False Then
                        '���ϸ�����������
                        For intCol = 0 To .Cols - 1
                            .TextMatrix(intRow, intCol) = ""
                        Next
                        .Row = intRow
                        .Col = mCol����
                        Exit Function
                    End If
                End If
                If Val(zlStr.Nvl(rsprice.Fields(0))) = 0 Then
                    .TextMatrix(intRow, mCol�ɹ���) = Format(dbl�ɱ��� * num����ϵ��, mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mCol�����) = Format(dbl�ɱ��� * num����ϵ�� * dblrate / 100, mFMT.FM_�ɱ���)
                Else
                    .TextMatrix(intRow, mCol�ɹ���) = Format(Val(zlStr.Nvl(rsprice.Fields(0)) * num����ϵ��), mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mCol�����) = Format(Val(zlStr.Nvl(rsprice.Fields(0)) * num����ϵ��) * dblrate / 100, mFMT.FM_�ɱ���)
                End If
            Else
                If dbl�ɱ��� > 0 Then
                    .TextMatrix(intRow, mCol�ɹ���) = Format(dbl�ɱ��� * num����ϵ��, mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mCol�����) = Format(dbl�ɱ��� * num����ϵ�� * dblrate / 100, mFMT.FM_�ɱ���)
                End If
            End If
        End If
        
        If .TextMatrix(intRow, mCol����) <> "" Then
            gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(mshBill.Row, mCol����), lng����ID)
            If Not rstemp.EOF Then
               .TextMatrix(intRow, mcol��׼�ĺ�) = IIf(IsNull(rstemp!��׼�ĺ�), "", rstemp!��׼�ĺ�)
            End If
        End If
        
        Dim dbl����� As Double, dbl�ɹ��� As Double
        
        dbl����� = Val(.TextMatrix(intRow, mCol�����))
        dbl�ɹ��� = Val(.TextMatrix(intRow, mCol�ɹ���))
        dbl�ɱ��� = IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����)
        'ʱ�۲��ϴ���
        If int�Ƿ��� = 1 Then
            If mint�༭״̬ = 8 Or mbln�˻� Then
                gstrSQL = "" & _
                "   Select ʵ�ʽ��/ʵ������*" & num����ϵ�� & " as  �ۼ� " & _
                "   From ҩƷ��� " & _
                "   Where �ⷿid=[1]" & _
                "           and ҩƷid=[2]" & _
                "           and ����=1 and ʵ������>0 and " & _
                "           nvl(����,0)=[3]"
                Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng����ID, lng����)
                If rstemp.EOF Then
                    MsgBox "ʱ�۲���û�п�棬���ܳ��⣬���飡", vbOKOnly, gstrSysName
                    Exit Function
                End If
               .TextMatrix(intRow, mCol�ۼ�) = Format(Nvl(rstemp!�ۼ�, 0), mFMT.FM_���ۼ�)
'               .TextMatrix(intRow, mcol�ӳ���) = Format(Val(.TextMatrix(intRow, mCol�ۼ�)) / Val(), "###0.00") & "%"
            Else
                If mbln�Ӽ��� = True Then
                    '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                    .TextMatrix(intRow, mCol�ۼ�) = Format(У�����ۼ�(dbl�ɱ��� * (1 + dbl�ӳ���) + _
                                                    ʱ�۲������ۼ�(lng����ID, dbl�ɱ���, dbl�ӳ���)) _
                                                    , mFMT.FM_���ۼ�)
                    .TextMatrix(intRow, mcol�ӳ���) = Format(dbl�ӳ��� * 100, "###0.00") & "%"
                ElseIf mbln�ֶμӳ��� = True Then
                    dbl�ӳ��� = 0
                    If Get�ֶμӳ��ۼ�(dbl�ɱ���, Val(.TextMatrix(intRow, mCol����ϵ��)), mstrCaption, sng�ֶ��ۼ�) = False Then
                        .TextMatrix(intRow, mCol�ۼ�) = Format(У�����ۼ�(dbl�ɱ��� * (1 + dbl�ӳ���) + _
                                                        ʱ�۲������ۼ�(lng����ID, dbl�ɱ���, dbl�ӳ���)) _
                                                        , mFMT.FM_���ۼ�)
                        .TextMatrix(intRow, mcol�ӳ���) = Format(dbl�ӳ��� * 100, "###0.00") & "%"
                    Else
                        .TextMatrix(intRow, mCol�ۼ�) = Format(У�����ۼ�(sng�ֶ��ۼ� + _
                                                      ʱ�۲������ۼ�(lng����ID, dbl�ɱ���, dbl�ӳ���, -1, sng�ֶ��ۼ�)) _
                                                      , mFMT.FM_���ۼ�)
                        .TextMatrix(intRow, mcol�ӳ���) = Format(dbl�ӳ��� * 100, "####0.00") & "%" '��Ϊ�Ƿֶμӳɵ����Լӳ��ʲ�׼ȷ��ȡһ��ģ��ֵ����
                    End If
                Else 'ȡ�ϴ��ۼ�ģʽ��û�й�ѡ�κ�ȡ�ۼ۷�ʽ
                    If mblnʱ������ȡ�ϴ��ۼ� = True Then
                        gstrSQL = "Select Nvl(�ϴ��ۼ�, 0) As �ϴ��ۼ� From �������� Where ����id = [1]"
                        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID)
                        If rstemp!�ϴ��ۼ� > 0 Then
                            .TextMatrix(intRow, mCol�ۼ�) = Format(zlStr.Nvl(rstemp!�ϴ��ۼ�, 0) * num����ϵ��, mFMT.FM_���ۼ�)
                            If dbl�ɱ��� <> 0 Then
                                .TextMatrix(intRow, mcol�ӳ���) = Format((Val(.TextMatrix(intRow, mCol�ۼ�)) / dbl�ɱ��� - 1) * 100, "###0.00") & "%"
                            End If
                        Else
                            '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                            .TextMatrix(intRow, mCol�ۼ�) = Format(У�����ۼ�(dbl�ɱ��� * (1 + dbl�ӳ���) + _
                                                            ʱ�۲������ۼ�(lng����ID, dbl�ɱ���, dbl�ӳ���)) _
                                                            , mFMT.FM_���ۼ�)
                            .TextMatrix(intRow, mcol�ӳ���) = Format(dbl�ӳ��� * 100, "###0.00") & "%"
                        End If
                    Else
                        '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                        .TextMatrix(intRow, mCol�ۼ�) = Format(У�����ۼ�(dbl�ɱ��� * (1 + dbl�ӳ���) + _
                                                        ʱ�۲������ۼ�(lng����ID, dbl�ɱ���, dbl�ӳ���)) _
                                                        , mFMT.FM_���ۼ�)
                        .TextMatrix(intRow, mcol�ӳ���) = Format(dbl�ӳ��� * 100, "###0.00") & "%"
                    End If
                End If
            End If
        Else
            .TextMatrix(intRow, mcol�ӳ���) = Format(dbl���ۼӳ��� * 100, "###0.00") & "%" '"15.00%"'����ȡ����мӳ���
        End If
        .TextMatrix(intRow, mcol���۵�λ) = strɢװ��λ
        '���˺�:���ۼ۴���
        Call �������ۼۼ����۲��(intRow)
        mshBill.MsfObj.CellForeColor = IIf(int�Ƿ��� = 0, &H0, &H40&)     ' &H40C0&
        
        If mstr���ս��� = "" Then
            gstrSQL = "Select ����  From ������ս��� where ȱʡ��־=1"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "SetColValue")
            
            If Not rstemp.EOF Then
                .TextMatrix(intRow, mCol���ս���) = IIf(IsNull(rstemp!����), "", rstemp!����)
                mstr���ս��� = rstemp!����
            End If
        Else
            .TextMatrix(intRow, mCol���ս���) = mstr���ս���
        End If
    End With
    Call ��ʾ�����
    SetColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetInputFormat(ByVal intRow As Integer)
    '--------------------------------------------------------------------------------------------------------
    '����:���õ�ǰ�еı༭��ʽ
    '����:introw-��ǰ��
    '����:
    '����:���˺�
    '����:2007/05/15
    '--------------------------------------------------------------------------------------------------------
    
    With mshBill
    
        '1.������2���޸ģ�3�����գ�4���鿴��5���޸ķ�Ʊ��6��������
        '7��������ˣ������������µ��ݲ���ˣ��Ѹ���ĵ��ݲ����������ˣ�ͬ����������˺�ĵ��ݲ����������;
        '8�����Ŀ��˻�,9-�˲�
        If mint�༭״̬ = 9 Or mint�༭״̬ = 3 Or mint�༭״̬ = 7 Then
            '���˺�:2007/05/30:�������̿���ʱ������ص�����
            Call Set��������Update(True)
        End If
        If mint�༭״̬ = 7 Then
                '�����ʱ�����ģ������������ۼ�
                '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                If Split(.TextMatrix(intRow, mColԭ����) & "||||||||", "||")(2) = 1 Then
                    If Split(.TextMatrix(intRow, mColԭ����) & "||||||||", "||")(4) = 1 Then
                        .ColData(mcol���ۼ�) = IIf(mbln�˻� Or mint�༭״̬ = 8, 5, 4)
                    Else
                        .ColData(mcol���ۼ�) = 5
                    End If
                Else
                    .ColData(mcol���ۼ�) = 5
                End If
        End If
        If mblnEdit = False Then Exit Sub
        
        If mint�༭״̬ = 7 Or mint�༭״̬ = 8 Or mbln�˻� Then Exit Sub
'        If .TextMatrix(intRow, mColԭ����) = "!" Then
            .ColData(mCol����) = 1              '���ı�����
'        Else
'            .ColData(mCol����) = 5              '��ֹ
'        End If

        If .TextMatrix(intRow, mcolһ���Բ���) = "1" Then
            .ColData(mcol�������) = 2
            .ColData(mcol���ʧЧ��) = 2
        Else
            .ColData(mcol�������) = 5              '��ֹ
            .ColData(mcol���ʧЧ��) = 5
        End If
        
        If .TextMatrix(intRow, mcol�������) = "1" Then
            .ColData(mcol��Ʒ����) = 4
        Else
            .ColData(mcol��Ʒ����) = 5
        End If

        .ColData(mColЧ��) = 2

        '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
        If .TextMatrix(intRow, mColԭ����) <> "" Then
            If mint�༭״̬ <> 9 And mint�༭״̬ <> 3 And mint�༭״̬ <> 7 Then
                '�����ʱ�����ģ������������ۼ�
                '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                If Split(.TextMatrix(intRow, mColԭ����), "||")(2) = 1 Then
                    If Split(.TextMatrix(intRow, mColԭ����), "||")(4) = 1 Then
                        .ColData(mcol���ۼ�) = IIf(mbln�˻� Or mint�༭״̬ = 8, 5, 4)
                    Else
                        .ColData(mcol���ۼ�) = 5
                    End If
                    .ColData(mCol�ۼ�) = IIf(mblnʱ������ֱ��ȷ���ۼ�, 4, 5)
                Else
                    .ColData(mCol�ۼ�) = 5
                    .ColData(mcol���ۼ�) = 5
                End If
            Else
                '�˲�\���\��������Ѿ������˸��ۼ۵�����
                '20070530:���˺�
            End If
            
        Else
            .ColData(mCol�ۼ�) = 5
            .ColData(mcol���ۼ�) = 5
        End If
        
        If Trim(.TextMatrix(intRow, mCol��Ʊ��)) = "" Then
            .ColData(mCol��Ʊ����) = 5
            .ColData(mCol��Ʊ���) = 5
            .ColData(mcol��Ʊ����) = 5
        Else
            .ColData(mCol��Ʊ����) = 2
            .ColData(mcol��Ʊ����) = 4
            .ColData(mCol��Ʊ���) = 4
        End If
        
    End With
End Sub


'�����ۿ�
Private Sub SetDisCount(ByVal intRow As Integer, ByVal intDisCount As Double)
    Dim dbl�ӳ��� As Double, dbl�ۼ� As Double, dbl�ɹ��� As Double, dbl���� As Double, dbl����� As Double
    Dim lng����ID  As Long
    Dim bln�Ƿ���� As Boolean
    
    With mshBill
        dbl�ۼ� = Val(.TextMatrix(intRow, mCol�ۼ�))
        dbl����� = Val(.TextMatrix(intRow, mCol�����))
        dbl�ɹ��� = Val(.TextMatrix(intRow, mCol�ɹ���))
        lng����ID = Val(.TextMatrix(intRow, 0))
        
        If mbln�Ӽ��� Then
            mdbl�Ӽ��� = 15
            If mblnʱ�۹�ǰ���� Then
                If dbl�ۼ� <> 0 And dbl�ɹ��� <> 0 Then
                    mdbl�Ӽ��� = ����ӳ���(lng����ID, dbl�ۼ�, dbl�ɹ���)
                    bln�Ƿ���� = True
                End If
            Else
                If dbl�ۼ� <> 0 And dbl����� <> 0 Then
                    mdbl�Ӽ��� = ����ӳ���(lng����ID, dbl�ۼ�, dbl�����)
                    bln�Ƿ���� = True
                End If
            End If
        End If
        
        If mshBill.Col = mColָ�������� Then
            .Text = Format(intDisCount, mFMT.FM_���)
            
            .TextMatrix(intRow, mColָ��������) = .Text
            If Val(.TextMatrix(intRow, mCol�ɹ���)) = 0 Then
                .TextMatrix(intRow, mCol�ɹ���) = .Text
                dbl�ɹ��� = Val(.TextMatrix(intRow, mCol�ɹ���))
            End If
            intDisCount = Val(.TextMatrix(intRow, mCol����))
        Else
            .TextMatrix(intRow, mCol����) = intDisCount
        End If
        
        If .TextMatrix(intRow, mColָ��������) <> "" Then
            If .TextMatrix(intRow, mCol�ɹ���) = "" Then
                .TextMatrix(intRow, mCol�ɹ���) = .TextMatrix(intRow, mColָ��������)
                dbl�ɹ��� = Val(.TextMatrix(intRow, mCol�ɹ���))
            End If
            If Not (mint�༭״̬ = 8 Or mbln�˻�) Then
                .TextMatrix(intRow, mCol�����) = Format((Val(.TextMatrix(intRow, mCol�ɹ���)) * intDisCount / 100), mFMT.FM_�ɱ���)
            End If
            dbl����� = Val(.TextMatrix(intRow, mCol�����))
            If .TextMatrix(intRow, mCol����) <> "" Then
               .TextMatrix(intRow, mCol������) = Format((Val(.TextMatrix(intRow, mCol����)) * Val(.TextMatrix(intRow, mCol�����))), mFMT.FM_���)
               .TextMatrix(intRow, mCol��Ʊ���) = IIf(Trim(.TextMatrix(intRow, mCol��Ʊ��)) = "" And Trim(.TextMatrix(intRow, mcol��Ʊ����)) = "", "", .TextMatrix(intRow, mCol������))
            End If
            .TextMatrix(intRow, mCol���) = Format(IIf(.TextMatrix(intRow, mCol�ۼ۽��) = "", 0, .TextMatrix(intRow, mCol�ۼ۽��)) - IIf(.TextMatrix(intRow, mCol������) = "", 0, .TextMatrix(intRow, mCol������)), mFMT.FM_���)
            '��ʱ�����ĵĴ���
            '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
            If .TextMatrix(intRow, mColԭ����) <> "" Then
                If Split(.TextMatrix(intRow, mColԭ����), "||")(2) = 1 Then
                    '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                    If mbln�Ӽ��� Then
                        If bln�Ƿ���� Then
                            .TextMatrix(intRow, mCol�ۼ�) = Format(У�����ۼ�(IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����) * (1 + (mdbl�Ӽ��� / 100)) + _
                                ʱ�۲������ۼ�(lng����ID, IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����), (mdbl�Ӽ��� / 100))), mFMT.FM_���ۼ�)
                        Else
                            dbl�ӳ��� = Val(Replace(.TextMatrix(intRow, mcol�ӳ���), "%", "")) / 100
                            .TextMatrix(intRow, mCol�ۼ�) = Format(У�����ۼ�(IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����) * (1 + dbl�ӳ���) + _
                                ʱ�۲������ۼ�(lng����ID, IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����), dbl�ӳ���)), mFMT.FM_���ۼ�)
                        End If
                    Else
                        dbl�ӳ��� = Val(Replace(.TextMatrix(intRow, mcol�ӳ���), "%", "")) / 100
                        .TextMatrix(intRow, mCol�ۼ�) = Format(У�����ۼ�(IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����) * (1 + dbl�ӳ���) + _
                            ʱ�۲������ۼ�(lng����ID, IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����), dbl�ӳ���)), mFMT.FM_���ۼ�)
                    End If
                    If .TextMatrix(intRow, mCol����) <> "" Then
                        .TextMatrix(intRow, mCol�ۼ۽��) = Format(.TextMatrix(intRow, mCol����) * Val(.TextMatrix(intRow, mCol�ۼ�)), mFMT.FM_���)
                        .TextMatrix(intRow, mCol���) = Format(IIf(.TextMatrix(intRow, mCol�ۼ۽��) = "", 0, .TextMatrix(intRow, mCol�ۼ۽��)) - IIf(.TextMatrix(intRow, mCol������) = "", 0, .TextMatrix(intRow, mCol������)), mFMT.FM_���)
                    End If
                End If
            End If
            Call �������ۼۼ����۲��(intRow)
        End If
    End With
End Sub

Private Sub mshBill_LeaveCell(Row As Long, Col As Long)
     ImeLanguage False
     
     With mshBill
        If .Col = mcol�������� And .TextMatrix(.Row, mcol��������) <> "" Then
            If ProduceDateCheck(.TextMatrix(.Row, mcol��������)) = False Then
                .TextMatrix(.Row, mcol��������) = ""
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub mshBill_LostFocus()
     ImeLanguage False
End Sub

Private Sub mshBill_Validate(Cancel As Boolean)
    mshBill.LastRow = 0
End Sub

Private Sub mshProvider_DblClick()
    mshProvider_KeyDown vbKeyReturn, 0
End Sub

Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        mshProvider.Visible = False
        txtProvider.SetFocus
        txtProvider.SelStart = 0
        txtProvider.SelLength = Len(txtProvider.Text)
    End If
    
    If KeyCode = vbKeyReturn Then
        txtProvider.Text = mshProvider.TextMatrix(mshProvider.Row, 2)
        txtProvider.Tag = mshProvider.TextMatrix(mshProvider.Row, 0)
        mshProvider.Visible = False
        mshBill.SetFocus
    End If
    
    If CheckQualifications(mlngModule, 2, Val(txtProvider.Tag)) = False Then
        txtProvider.Text = ""
        txtProvider.Tag = "0"
        Exit Sub
    End If

    If Val(txtProvider.Tag) <> mlng������λID And (mint�༭״̬ = 8 Or mbln�˻�) Then
        mshBill.ClearBill
        mlng������λID = Val(txtProvider.Tag)
        mshBill.TextMatrix(1, mCol�к�) = "1"
    End If
End Sub

Private Sub mshProvider_LostFocus()
    If mshProvider.Visible Then
        mshProvider.Visible = False
    End If
End Sub

Private Sub msh����_DblClick()
    msh����_KeyDown vbKeyReturn, 0
End Sub

Private Sub msh����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As ADODB.Recordset
    
    With mshBill
    
        If KeyCode = vbKeyEscape Then
            msh����.Visible = False
            .SetFocus
        End If
        
        If .Col = mCol���ս��� And KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, .Col) = msh����.TextMatrix(msh����.Row, 1)
            .Text = .TextMatrix(.Row, .Col)
            msh����.Visible = False
            .SetFocus
            Call ColMoveNextCol(.Col)
            Exit Sub
        End If
        
        If CheckQualifications(mlngModule, 1, msh����.TextMatrix(msh����.Row, 2)) = False Then
            Exit Sub
        End If
        
        If KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, .Col) = msh����.TextMatrix(msh����.Row, 2)
            msh����.Visible = False
            
            gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
            Set rsProvider = zlDatabase.OpenSQLRecord(gstrSQL, "msh����_KeyDown", .TextMatrix(.Row, .Col), .TextMatrix(.Row, 0))
            If rsProvider.RecordCount > 0 Then
                .TextMatrix(.Row, mcol��׼�ĺ�) = IIf(IsNull(rsProvider!��׼�ĺ�), "", rsProvider!��׼�ĺ�)
            Else
                .TextMatrix(.Row, mcol��׼�ĺ�) = ""
            End If
            
            .Col = mCol����
            .SetFocus
        End If
    
    End With
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
            If mshBill.Col = mCol����� Or mshBill.Col = mCol������ Then Exit Sub
        End If
    End If
    PicInput.Visible = False
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And stbThis.Tag <> "PY" Then
        Logogram stbThis, 0
        stbThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And stbThis.Tag <> "WB" Then
        Logogram stbThis, 1
        stbThis.Tag = Panel.Key
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


Private Sub txtCopy_Change()
    txtCopy.Text = Val(txtCopy.Text)
    If Val(txtCopy.Text) > 9999 Then txtCopy.Text = 9999
End Sub

Private Sub txtCopy_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtDrawPerson_Change()
    mblnChange = True
    txtDrawPerson.Tag = ""
End Sub

Private Sub txtDrawPerson_GotFocus()
    OS.OpenIme False
    zlControl.TxtSelAll txtDrawPerson
End Sub

Private Sub txtDrawPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtDrawPerson.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
    If SelectItem(txtDrawPerson, Trim(txtDrawPerson.Text), True) = False Then Exit Sub
End Sub

Private Sub txtNO_Change()
    If txtNO.Locked = True Then
'        If mstr���ݺ� <> "" And mstr���ݺ� <> txtNO.Text Then
'            txtNO.Text = mstr���ݺ�
'        End If
    End If
End Sub

Private Sub TxtNo_GotFocus()
    If txtNO.Locked = False Then
        txtNO.SelStart = 0
        txtNO.SelLength = Len(txtNO.Text)
    End If
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
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
    mblnChange = True
End Sub

Private Sub txtProvider_GotFocus()
    txtProvider.SelStart = 0
    txtProvider.SelLength = Len(txtProvider.Text)
End Sub

Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String
    Dim adoProvider As New Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then Exit Sub
    
    On Error GoTo ErrHandle
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = GetMatchingSting(UCase(.Text))
        
        
        gstrSQL = "" & _
            "   Select id,����,����,���� " & _
            "   From ��Ӧ�� " & _
            "   Where (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
            "       And (վ��=[2] or վ�� is null) And ĩ��=1 And (substr(����,5,1)=1 ) " & _
            "       And (���� like [1] Or ���� like [1] or ���� like [1]) "
        Set adoProvider = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strProviderText, gstrNodeNo)
        
        If adoProvider.EOF Then
            MsgBox "û��������Ĺ�����λ�������䣡", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If
        
        If adoProvider.RecordCount > 1 Then
            Set mshProvider.Recordset = adoProvider
            Dim intCol As Integer
            Dim intRow As Integer
            
            With mshProvider
                If .Visible = False Then .Visible = True
                .Redraw = False
                .SetFocus
                
                For intRow = 0 To .Rows - 1
                    .Row = intRow
                    For intCol = 0 To .Cols - 1
                        .Col = intCol
                        If .Row = 0 Then
                            .CellFontBold = True
                        Else
                            .CellFontBold = False
                        End If
                    Next
                Next
                .Font.Bold = False
                .FontFixed.Bold = True
                .ColWidth(0) = 0
                .ColWidth(1) = 1000
                .ColWidth(2) = 2700
                .ColWidth(3) = 1200
                .Row = 1
                .TopRow = 1
                .Col = 0
                .ColSel = .Cols - 1
                
                .Top = txtProvider.Top + txtProvider.Height
                .Left = cmdProvider.Left + cmdProvider.Width - .Width
                .Redraw = True
                Exit Sub
            End With
        Else
            .Text = adoProvider!����
            .Tag = adoProvider!Id
        End If
        adoProvider.Close
        mshBill.SetFocus
        mshBill.Col = 1
        mshBill.Row = 1
        
        If CheckQualifications(mlngModule, 2, Val(txtProvider.Tag)) = False Then
            txtProvider.Text = ""
            txtProvider.Tag = "0"
            Exit Sub
        End If
        
        If Val(.Tag) <> mlng������λID And mint�༭״̬ = 8 Then
            mlng������λID = Val(txtProvider.Tag)
            mshBill.ClearBill
            mshBill.TextMatrix(1, mCol�к�) = "1"
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Function ValidData() As Boolean
    Dim intLop As Integer
    Dim rsStock As New Recordset
    Dim blnStock As Boolean
    Dim bln��ֵ����¼�� As Boolean
    Dim strNo As String
    Dim bln��ֵ���� As Boolean
    
    On Error GoTo ErrHandle
    ValidData = False
    
    gstrSQL = "" & _
        "   SELECT count(*)" & _
        "   From ��������˵�� " & _
        "   WHERE ((�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���')) " & _
        "           AND ����id =[1]"
        
    Set rsStock = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
    
    If rsStock.Fields(0) > 0 Then
        blnStock = False
    Else
        blnStock = True
    End If
    
    If txtNO.Locked = False Then
        '�������������޸ĵ��ݺ�
        strNo = txtNO.Text
        If strNo = "" Then
            ShowMsgBox "�����뵥�ݺš�"
            txtNO.SetFocus
            Exit Function
        End If
        
        If InStr(strNo, "'") > 0 Then
            ShowMsgBox "���ݺ����������к��зǷ��ַ���"
            txtNO.SetFocus
            Exit Function
        End If
        
        If LenB(StrConv(strNo, vbFromUnicode)) > 8 Then
            ShowMsgBox "���ݺų��Ȳ��ܳ���8����ĸ��"
            txtNO.SetFocus
            Exit Function
        End If
    Else
'        '��ֹ�û�ǿ���޸�
'        If mstr���ݺ� <> "" And mstr���ݺ� <> txtNO.Text Then
'            txtNO.Text = mstr���ݺ�
'        End If
    End If
    
    If CheckQualifications(mlngModule, 2, Val(txtProvider.Tag)) = False Then Exit Function
    
    bln��ֵ����¼�� = IIf(Val(zlDatabase.GetPara("��ֵ���ı�����д��ϸ��Ϣ", glngSys, mlngModule)) = 1, True, False)
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            If Val(txtProvider.Tag) = 0 Then
                ShowMsgBox "������λ����Ϊ�գ�"
                txtProvider.SetFocus
                Exit Function
            End If
            
            '29679 ����취
            If Val(.TextMatrix(.Row, mCol����)) < 0 Then
                ShowMsgBox "��������С��0��"
                .Col = mCol����
                .SetFocus
                Exit Function
            End If
            
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                ShowMsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!"
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mCol����)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mCol����))) = "" Then
                        ShowMsgBox "��" & intLop & "���������ϵ�����Ϊ���ˣ����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol����
                        Exit Function
                    End If
                    
'                    If Val(Trim(.TextMatrix(intLop, mCol�����))) = 0 Then
'                        ShowMsgbox "��" & intLop & "���������ϵĽ����Ϊ���ˣ����飡"
'                        .SetFocus
'                        .Row = intLop
'                        .MsfObj.TopRow = intLop
'                        .Col = mCol�����
'                        Exit Function
'                    End If
                    If mbln��ǿ�ƿ���ָ���۸� = False Then
                        If Val(.TextMatrix(intLop, mColָ��������)) < 0 Then
                            ShowMsgBox "��" & intLop & "���������ϵĲɹ��޼۱�����ڵ���0�����飡"
                            .SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mCol�����
                            Exit Function
                        End If
                    End If
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mCol����))), vbFromUnicode)) > mintBatchNoLen Then
                        ShowMsgBox "��" & intLop & "���������ϵ����ų���,���������" & Int(mintBatchNoLen / 2) & "�����ֻ�" & mintBatchNoLen & "���ַ�!"
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol����
                        Exit Function
                    End If
                    
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mcolע��֤��))), vbFromUnicode)) > 50 Then
                        ShowMsgBox "��" & intLop & "���������ϵ�ע��֤�ų���,���������25�����ֻ�50���ַ�!"
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mcolע��֤��
                        Exit Function
                    End If
                    
                    If Len(Trim(.TextMatrix(intLop, mcol��Ʒ����))) > 50 Then
                        ShowMsgBox "��" & intLop & "���������ϵ���Ʒ���볬��,���������50���ַ�!"
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mcol��Ʒ����
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mCol����))) = "" Then
                        ShowMsgBox "��" & intLop & "���������ϵĿ���Ϊ���ˣ����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol����
                        Exit Function
                    End If
                    
                    If Val(Trim(Trim(.TextMatrix(intLop, mCol����)))) >= 1000# Then
                        ShowMsgBox "��" & intLop & "���������ϵĿ���̫���ˣ����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol����
                        Exit Function
                    End If
                    
                    If blnStock = True Then
                        '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                        If Split(.TextMatrix(intLop, mColԭ����), "||")(0) <> "0" Then
                                        
                            If Trim(.TextMatrix(intLop, mCol����)) = "" Or Trim(.TextMatrix(intLop, mColЧ��)) = "" Then
                                ShowMsgBox "��" & intLop & "�е�����������Ч�ڲ���,����������ż�Ч��" & vbCrLf & "��Ϣ�������뵥���У�"
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                If .TextMatrix(intLop, mCol����) = "" Then
                                    .Col = mCol����
                                Else
                                    .Col = mColЧ��
                                End If
                                Exit Function
                            End If
                        End If
                        
                        '����ҩƷ����¼����غ�����
                        If mbln�����������Ų��ؿ��� = True And Not (mint�༭״̬ = 8 Or mbln�˻� = True) Then '�˻������
                            '����ҩƷ����¼����غ�����
                            If Split(.TextMatrix(intLop, mColԭ����), "||")(4) <> "0" And (.TextMatrix(intLop, mCol����) = "" Or .TextMatrix(intLop, mCol����) = "") Then
                                MsgBox "��" & intLop & "�е������Ƿ�������,������Ĳ��غ�����" & vbCrLf & "��Ϣ���뵥���У�", vbInformation, gstrSysName
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                If .TextMatrix(intLop, mCol����) = "" Then
                                    .Col = mCol����
                                Else
                                    .Col = mCol����
                                End If
                                Exit Function
                            End If
                        End If
                    Else '�����ǡ����ϲ��š�
                        If Split(.TextMatrix(intLop, mColԭ����), "||")(3) <> "0" Then
                            '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                            If Split(.TextMatrix(intLop, mColԭ����), "||")(0) <> "0" Then
                                            
                                If Trim(.TextMatrix(intLop, mCol����)) = "" Or Trim(.TextMatrix(intLop, mColЧ��)) = "" Then
                                    ShowMsgBox "��" & intLop & "�е�����������Ч�ڲ���,����������ż�Ч��" & vbCrLf & "��Ϣ�������뵥���У�"
                                    mshBill.SetFocus
                                    .Row = intLop
                                    .MsfObj.TopRow = intLop
                                    If .TextMatrix(intLop, mCol����) = "" Then
                                        .Col = mCol����
                                    Else
                                        .Col = mColЧ��
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                        
                        '����ҩƷ����¼����غ�����
                        If mbln�����������Ų��ؿ��� = True And Not (mint�༭״̬ = 8 Or mbln�˻� = True) Then '�˻������
                        '����ҩƷ����¼����غ�����
                            If Split(.TextMatrix(intLop, mColԭ����), "||")(3) <> "0" And (.TextMatrix(intLop, mCol����) = "" Or .TextMatrix(intLop, mCol����) = "") Then
                                MsgBox "��" & intLop & "�е������Ƿ�������,������Ĳ��غ�����" & vbCrLf & "��Ϣ���뵥���У�", vbInformation, gstrSysName
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                If .TextMatrix(intLop, mCol����) = "" Then
                                    .Col = mCol����
                                Else
                                    .Col = mCol����
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mCol�����)) > 9999999999# Then
                        ShowMsgBox "  ��" & intLop & "���������ϵĽ���۴��������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol�����
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mCol�����)) < 0 Then
                        ShowMsgBox "  ��" & intLop & "���������ϵĽ���۱�����ڵ���0�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol�����
                        Exit Function
                    End If
                                        
                    If Val(.TextMatrix(intLop, mCol����)) > 9999999999# Then
                        ShowMsgBox "��" & intLop & "���������ϵ��������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mCol������)) > 9999999999999# Then
                        ShowMsgBox "��" & intLop & "���������ϵĽ�������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol������
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mCol�ۼ۽��)) > 9999999999999# Then
                        ShowMsgBox "��" & intLop & "���������ϵ��ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol����
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mcol���۽��)) > 9999999999999# Then
                        ShowMsgBox "��" & intLop & "���������ϵ����۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mcol���۽��
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mcol���ۼ�)) > 9999999999999# Then
                        ShowMsgBox "��" & intLop & "���������ϵ����ۼ۴��������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mcol���ۼ�
                        Exit Function
                    End If
                    
                    
                    If Val(.TextMatrix(intLop, mcol���۲��)) > 9999999999999# Then
                        ShowMsgBox "��" & intLop & "���������ϵ����۲�۴��������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mcol���۲��
                        Exit Function
                    End If
                    
                    If LenB(StrConv(.TextMatrix(intLop, mCol�������), vbFromUnicode)) > 200 Then
                        MsgBox "��" & intLop & "�е�������������Ų��ܴ���200���ַ���100�����֣�", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol�������
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mCol��Ʊ���)) > 1E+15 Then
                        ShowMsgBox "��" & intLop & "���������ϵ��ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ999999999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol��Ʊ���
                        Exit Function
                    End If
                    If zlCommFun.ActualLen(.TextMatrix(intLop, mCol��Ʊ��)) > mint��Ʊ��Len Then
                        ShowMsgBox "��" & intLop & "���������ϵķ�Ʊ�����������" & mint��Ʊ��Len & "���ַ���" & mint��Ʊ��Len / 2 & "�����֣����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mCol��Ʊ��
                        Exit Function
                    End If
                    If zlCommFun.ActualLen(.TextMatrix(intLop, mcol��Ʊ����)) > 20 Then
                        ShowMsgBox "��" & intLop & "���������ϵķ�Ʊ�������������" & 20 & "���ַ���" & 20 / 2 & "�����֣����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mcol��Ʊ����
                        Exit Function
                    End If
                    
                    bln��ֵ���� = IsCostly(.TextMatrix(intLop, 0))
                    '�Ƿ�ǿ��¼���ֵ������Ϣ
                    If bln��ֵ����¼�� = True And bln��ֵ���� = True Then
                        If Trim(.TextMatrix(intLop, mcolע��֤��)) = "" Then
                            ShowMsgBox "��" & intLop & "��δ¼�롰ע��֤�š���Ϣ�����飡"
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mcolע��֤��
                            Exit Function
                        End If
                        If mrsCostlyInfo.RecordCount > 0 Then mrsCostlyInfo.MoveFirst
                        mrsCostlyInfo.Find "SN=" & .TextMatrix(.Row, 1)
                        If mrsCostlyInfo.EOF Then
                            ShowMsgBox "��" & intLop & "��δ¼���ֵ������Ϣ�����飡"
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mCol����
                            Exit Function
                        Else
                            Dim blnCostlyOK As Boolean
                            blnCostlyOK = True
                            If IIf(IsNull(mrsCostlyInfo!����), "", mrsCostlyInfo!����) = "" Then
                                blnCostlyOK = False
                                ShowMsgBox "��" & intLop & "��δ¼���ֵ���ĵġ����ҡ���Ϣ�����飡"
                            End If
                            If blnCostlyOK And IIf(IsNull(mrsCostlyInfo!��������), "", mrsCostlyInfo!��������) = "" Then
                                blnCostlyOK = False
                                ShowMsgBox "��" & intLop & "��δ¼���ֵ���ĵġ�������������Ϣ�����飡"
                            End If
                            If blnCostlyOK And IIf(IsNull(mrsCostlyInfo!סԺ��), "", mrsCostlyInfo!סԺ��) = "" Then
                                blnCostlyOK = False
                                ShowMsgBox "��" & intLop & "��δ¼���ֵ���ĵġ�סԺ�š���Ϣ�����飡"
                            End If
                            If blnCostlyOK And IIf(IsNull(mrsCostlyInfo!����), "", mrsCostlyInfo!����) = "" Then
                                blnCostlyOK = False
                                ShowMsgBox "��" & intLop & "��δ¼���ֵ���ĵġ����š���Ϣ�����飡"
                            End If
                            If blnCostlyOK = False Then
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                .Col = mCol����
                                Exit Function
                            End If
                        End If
                    End If
                    
                    '�б굥���ж�
                    Dim dblCostPrice As Double, dblPrice As Double
                    Dim strBidMess As String
                    dblCostPrice = Get�б굥λ�ɱ���(.TextMatrix(intLop, 0))
                    dblPrice = CDbl(IIf(.TextMatrix(intLop, mCol�ɹ���) <> "", .TextMatrix(intLop, mCol�ɹ���), _
                                    IIf(.TextMatrix(intLop, mCol�ɹ���) = "", 0, .TextMatrix(intLop, mCol�ɹ���))))
                    If dblCostPrice < dblPrice And dblCostPrice <> 0 Then
                        strBidMess = zlDatabase.GetPara("��ⵥ�۳��б굥��", glngSys, mlngModule)
                        If Val(strBidMess) = 0 Then     '��ֹ��ⵥ�۳��б굥��
                            ShowMsgBox "��" & intLop & "�н�ֹ�ɹ��ۣ�" & dblPrice & "���� �б굥�ۣ�" & dblCostPrice & "����"
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mCol�ɹ���
                            Exit Function
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
ErrHandle:
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
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If CheckQualifications(mlngModule, 2, Val(txtProvider.Tag)) = False Then
        txtProvider.Text = ""
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If Val(txtProvider.Tag) <> mlng������λID And (mint�༭״̬ = 8 Or mbln�˻�) Then
        mlng������λID = Val(txtProvider.Tag)
        mshBill.ClearBill
        mshBill.TextMatrix(1, mCol�к�) = "1"
    End If
End Sub

Private Function SaveCard(Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
'----------------------------------------------------------------------------
'�޸ĸù���ʱ��ע�� frmPurchaseVerifyBatch(�������)����Ĺ����Ƿ��漰
'----------------------------------------------------------------------------
    Dim chrNo As Variant
    Dim lng��� As Long
    Dim lngStockID As Long
    Dim lng������λid As Long
    Dim lng����ID As Long
    Dim str���� As String
    Dim str���� As String
    Dim strЧ�� As String
    Dim dblʵ������ As Double
    Dim dbl�ɱ��� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl���� As Double
    Dim dbl���ۼ� As Double
    Dim dbl���۽�� As Double
    Dim dbl��� As Double
    Dim str���۲�� As String '�շ���¼�����÷��ֶα�������⹺����۲�����÷��ֶ��������ַ�������������double���ͻ���� -.00x����
    Dim strժҪ As String
    Dim str������ As String
    Dim str�������� As String
    Dim str����� As String
    Dim datAssessDate As String
    Dim str��Ʊ�� As String
    Dim str��Ʊ���� As String
    Dim str��Ʊ���� As String
    Dim str������� As String
    Dim str���ʧЧ�� As String
    Dim dbl��Ʊ��� As Double
    Dim str��������  As String
    Dim str�˲��� As String
    Dim str�˲����� As String
    Dim strע��֤�� As String
    Dim intUnit As Integer
    Dim strUnit As String
    Dim strָ�������� As String
    Dim str������� As String
    Dim str���ս��� As String
    Dim str��Ʒ���� As String
    Dim str�ڲ����� As String
    Dim str��׼�ĺ� As String
    Dim lng����ID As Long
    Dim intRow As Integer
    Dim i As Integer
    Dim arrSQL As Variant
    Dim blnTrans As Boolean
    Dim n As Long
    
    SaveCard = False
    arrSQL = Array()
    With mshBill
        
        chrNo = Trim(txtNO)
        lngStockID = cboStock.ItemData(cboStock.ListIndex)
        
        If mint�༭״̬ = 1 Then
            If chrNo <> "" Then
                If CheckNOExists(68, chrNo) Then Exit Function
            End If
            If chrNo = "" Then
                chrNo = sys.GetNextNo(68, lngStockID)
            End If
            If IsNull(chrNo) Then Exit Function
        End If
        txtNO.Tag = chrNo
        lng������λid = txtProvider.Tag
        strժҪ = Trim(txtժҪ.Text)
        
        
        '���˺�:2007/05/15:����˲���
        str������ = Txt������
        str�˲��� = IIf(txt�˲���.Visible, txt�˲���, "")
        str����� = Txt�����
        
        If mint�༭״̬ = 9 Then
            str�������� = Trim(Txt��������.Caption)
            str�˲����� = IIf(txt�˲���.Visible, Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss"), "")
        Else
            str�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            If blnǿ�Ʊ��� Then
                str�˲����� = Trim(txt�˲�����.Caption)
            Else
                str�˲����� = ""
            End If
        End If
        
        On Error GoTo ErrHandle
        
        If mint�༭״̬ = 2 Or mint�༭״̬ = 9 Or blnǿ�Ʊ��� Then         '2:�޸�;  9:�˲�;
            gstrSQL = "zl_�����⹺_Delete('" & mstr���ݺ� & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
        End If
        
        'ȡ�ÿⷿ�ĵ�λ������ָ��������ʱʹ��
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng����ID = .TextMatrix(intRow, 0)
                str���� = .TextMatrix(intRow, mCol����)
                str���� = .TextMatrix(intRow, mCol����)
                
                str��׼�ĺ� = .TextMatrix(intRow, mcol��׼�ĺ�)
                strЧ�� = IIf(Trim(.TextMatrix(intRow, mColЧ��)) = "", "", .TextMatrix(intRow, mColЧ��))
                dblʵ������ = GetFormat(.TextMatrix(intRow, mCol����) * .TextMatrix(intRow, mCol����ϵ��), g_С��λ��.obj_���С��.����С��)
                dbl���� = Val(.TextMatrix(intRow, mCol����))
                dbl�ɱ��� = GetFormat(Val(.TextMatrix(intRow, mCol�����)) / .TextMatrix(intRow, mCol����ϵ��), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl�ɱ���� = GetFormat(Val(.TextMatrix(intRow, mCol������)), g_С��λ��.obj_���С��.���С��)
                
                '���˺�:���ۼ۴���
                
                'dbl���ۼ� = Round(Val(.TextMatrix(intRow, mCol�ۼ�)) / .TextMatrix(intRow, mCol����ϵ��), g_С��λ��.obj_ɢװС��.���ۼ�С��)
                'dbl���۽�� = Round(Val(.TextMatrix(intRow, mCol�ۼ۽��)), g_С��λ��.obj_ɢװС��.���С��)
                '���ݿ��е�:��� = ���۽�� - ������
                '���ݿ��е�:�÷� = ���۽��-�ۼ۽������۲��-���(�ⷿ��λ�Ĳ��)

                dbl���ۼ� = GetFormat(Val(.TextMatrix(intRow, mcol���ۼ�)), g_С��λ��.obj_���С��.���ۼ�С��)
                dbl���۽�� = GetFormat(Val(.TextMatrix(intRow, mcol���۽��)), g_С��λ��.obj_���С��.���ۼ�С��)
                dbl��� = GetFormat(Val(.TextMatrix(intRow, mcol���۲��)), g_С��λ��.obj_���С��.���ۼ�С��)
                str���۲�� = GetFormat(Val(.TextMatrix(intRow, mcol���۲��)) - Val(.TextMatrix(intRow, mCol���)), g_С��λ��.obj_���С��.���ۼ�С��)
                'dbl��� = Round(Val(.TextMatrix(intRow, mCol���)), g_С��λ��.obj_ɢװС��.���С��)
                lng��� = intRow
                
                str���ս��� = Trim(.TextMatrix(intRow, mCol���ս���))
                str������� = Trim(.TextMatrix(intRow, mCol�������))
                str��Ʊ�� = Trim(.TextMatrix(intRow, mCol��Ʊ��))
                str��Ʊ���� = Trim(.TextMatrix(intRow, mcol��Ʊ����))
                str��Ʊ���� = Trim(IIf(.TextMatrix(intRow, mCol��Ʊ����) = "", "", .TextMatrix(intRow, mCol��Ʊ����)))
                dbl��Ʊ��� = Round(Val(.TextMatrix(intRow, mCol��Ʊ���)), g_С��λ��.obj_ɢװС��.���С��)
                  
                str������� = Trim(IIf(.TextMatrix(intRow, mcol�������) = "", "", .TextMatrix(intRow, mcol�������)))
                str���ʧЧ�� = Trim(IIf(.TextMatrix(intRow, mcol���ʧЧ��) = "", "", .TextMatrix(intRow, mcol���ʧЧ��)))
                str�������� = Trim(IIf(.TextMatrix(intRow, mcol��������) = "", "", .TextMatrix(intRow, mcol��������)))
                strע��֤�� = Trim(.TextMatrix(intRow, mcolע��֤��))
                
                str�ڲ����� = Trim(.TextMatrix(intRow, mcol�ڲ�����))
                lng����ID = Val(.TextMatrix(intRow, mcol����ID))
                
                If gblnCode = True Then str��Ʒ���� = Trim(.TextMatrix(intRow, mcol��Ʒ����))
                
                '���²��������е�ָ��������
                strָ�������� = Val(.TextMatrix(intRow, mColָ��������)) & "/" & IIf(mintUnit = 0, "1", "����ϵ��")
                
                '����:����ID_IN,SQL_IN
                If mbln��ǿ�ƿ���ָ���۸� = False Then
                    gstrSQL = "zl_��������_UpdateCustom(" & lng����ID & ",'ָ��������=" & strָ�������� & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                End If
                ' Zl_�����⹺_Insert
                gstrSQL = "zl_�����⹺_INSERT("
                '  No_In         In ҩƷ�շ���¼.NO%Type,
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                '  ���_In       In ҩƷ�շ���¼.���%Type,
                gstrSQL = gstrSQL & "" & lng��� & ","
                '  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
                gstrSQL = gstrSQL & "" & lngStockID & ","
                '  ��ҩ��λid_In In ҩƷ�շ���¼.��ҩ��λid%Type,
                gstrSQL = gstrSQL & "" & lng������λid & ","
                '  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
                gstrSQL = gstrSQL & "" & lng����ID & ","
                '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "'" & str���� & "',"
                '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "'" & str���� & "',"
                '  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str�������� = "", "Null", "to_date('" & Format(str��������, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  �������_In   In ҩƷ�շ���¼.�������%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str������� = "", "Null", "to_date('" & Format(str�������, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  ���Ч��_In   In ҩƷ�շ���¼.���Ч��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str���ʧЧ�� = "", "Null", "to_date('" & Format(str���ʧЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  ʵ������_In   In ҩƷ�շ���¼.ʵ������%Type := Null,
                gstrSQL = gstrSQL & "" & dblʵ������ & ","
                '  �ɱ���_In     In ҩƷ�շ���¼.�ɱ���%Type := Null,
                gstrSQL = gstrSQL & "" & dbl�ɱ��� & ","
                '  �ɱ����_In   In ҩƷ�շ���¼.�ɱ����%Type := Null,
                gstrSQL = gstrSQL & "" & dbl�ɱ���� & ","
                '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "" & dbl���� & ","
                '  ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type := Null,
                gstrSQL = gstrSQL & "" & dbl���ۼ� & ","
                '  ���۽��_In   In ҩƷ�շ���¼.���۽��%Type := Null,
                gstrSQL = gstrSQL & "" & dbl���۽�� & ","
                '  ���_In       In ҩƷ�շ���¼.���%Type := Null,
                gstrSQL = gstrSQL & "" & dbl��� & ","
                '  ���۲��_In   In ҩƷ�շ���¼.���%Type := Null,Ŀǰ������÷��ֶ�
                gstrSQL = gstrSQL & "" & str���۲�� & ","
                '  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(strժҪ = "", "NULL", "'" & strժҪ & "'") & ","
                '   ע��֤��_In   In ҩƷ�շ���¼.ע��֤��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(strע��֤�� = "", "NULL", "'" & strע��֤�� & "'") & ","
                '  ������_In     In ҩƷ�շ���¼.������%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str������ = "", "NULL", "'" & str������ & "'") & ","
                '  �������_In   In Ӧ����¼.�������%Type := Null
                gstrSQL = gstrSQL & "" & IIf(str������� = "", "NULL", "'" & str������� & "'") & ","
                '  ��Ʊ��_In     In Ӧ����¼.��Ʊ��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str��Ʊ�� = "", "NULL", "'" & str��Ʊ�� & "'") & ","
                '  ��Ʊ����_In   In Ӧ����¼.��Ʊ����%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str��Ʊ���� = "", "Null", "to_date('" & Format(str��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  ��Ʊ���_In   In Ӧ����¼.��Ʊ���%Type := Null,
                gstrSQL = gstrSQL & "" & dbl��Ʊ��� & ","
                '  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
                gstrSQL = gstrSQL & "to_date('" & str�������� & "','yyyy-mm-dd HH24:MI:SS'),"
                '  �˲���_In     In ҩƷ�շ���¼.��ҩ��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str�˲��� = "", "NULL", "'" & str�˲��� & "'") & ","
                '  �˲�����_In   In ҩƷ�շ���¼.��ҩ����%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str�˲����� = "", "Null", "to_date('" & str�˲����� & "','yyyy-mm-dd hh24:mi:ss')") & ","
                '  ����_In       In ҩƷ�շ���¼.����%Type := 0,
                gstrSQL = gstrSQL & "" & Val(.TextMatrix(intRow, mCol����)) & ","
                '  �˻�_In       In Number := 1
                gstrSQL = gstrSQL & "" & IIf(mbln�˻�, -1, 1) & ","
                '  ��ֵ����_In   In varchar2(250)
                gstrSQL = gstrSQL & "'" & GetCostlyInfoStr(intRow) & "'" & ","
                '  ��Ʒ����_In   In ҩƷ�շ���¼.��Ʒ����%Type :=Null
                gstrSQL = gstrSQL & "" & IIf(str��Ʒ���� = "", "NULL", "'" & str��Ʒ���� & "'") & ","
                '  �ڲ�����
                gstrSQL = gstrSQL & IIf(str�ڲ����� = "", "Null", "'" & str�ڲ����� & "'") & ","
                '  ����ID
                gstrSQL = gstrSQL & IIf(lng����ID = 0, "Null", lng����ID) & ","
                '  ��Ʊ����
                gstrSQL = gstrSQL & IIf(str��Ʊ���� = "", "NULL", "'" & str��Ʊ���� & "'") & ","
                '   �������
                gstrSQL = gstrSQL & "0" & ","
                '   ��׼�ĺ�
                gstrSQL = gstrSQL & IIf(str��׼�ĺ� = "", "NULL", "'" & str��׼�ĺ� & "'") & ","
                '   ���ս���
                gstrSQL = gstrSQL & "'" & str���ս��� & "'"
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
            End If
            
            recSort.MoveNext
        Next
        
        If blnǿ�Ʊ��� = False Then gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "SaveCard")
        Next
         If blnǿ�Ʊ��� = False Then gcnOracle.CommitTrans: blnTrans = False
        
        mstr���ݺ� = chrNo
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
    
ErrHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'�˻�
Private Function SaveRestore() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�˻�����
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-28 14:15:40
    '-----------------------------------------------------------------------------------------------------------

    Dim lng��� As Long, lngStockID As Long, lng������λid As Long, lng����ID As Long
    Dim str���� As String, str���� As String, strЧ�� As String, chrNo As String
    Dim dblʵ������ As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���� As Double
    Dim dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double, dbl���۲�� As Double
    Dim strժҪ As String, str������ As String, str�������� As String, str����� As String
    Dim datAssessDate As String, str�������� As String, strע��֤�� As String
    Dim str��Ʊ�� As String, str��Ʊ���� As String, dbl��Ʊ��� As Double
    Dim intUnit As Integer, strUnit As String, strָ�������� As String
    Dim intRow As Integer, str������� As String, str���ʧЧ�� As String, str�˲��� As String
    Dim str�˲����� As String, str������� As String
    Dim str��Ʒ���� As String
    Dim i As Integer
    Dim arrSQL As Variant
    Dim n As Long
    
    SaveRestore = False
    arrSQL = Array()
    'ֻ�пⷿ������ʹ���˻�����
    
    If Val(txtProvider.Tag) = 0 Then
        MsgBox "��ѡ��Ӧ�̣�", vbInformation, gstrSysName
        txtProvider.SetFocus
        Exit Function
    End If
    
    With mshBill
        chrNo = Trim(txtNO.Tag)
        lngStockID = cboStock.ItemData(cboStock.ListIndex)
        If chrNo = "" Then chrNo = sys.GetNextNo(68, lngStockID)
        If IsNull(chrNo) Then Exit Function
        
        txtNO.Tag = chrNo
        lng������λid = Val(txtProvider.Tag)
        strժҪ = Trim(txtժҪ.Text)
        str������ = Txt������
        str�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        str����� = Txt�����
        str�˲��� = txt�˲���
        str�˲����� = txt�˲�����
        
        On Error GoTo ErrHandle
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng����ID = Val(.TextMatrix(intRow, 0))
                str���� = .TextMatrix(intRow, mCol����)
                str���� = .TextMatrix(intRow, mCol����)
                strЧ�� = IIf(.TextMatrix(intRow, mColЧ��) = "", "", .TextMatrix(intRow, mColЧ��))
                dblʵ������ = Round(Val(.TextMatrix(intRow, mCol����)) * .TextMatrix(intRow, mCol����ϵ��), g_С��λ��.obj_���С��.����С��)
                dbl���� = Val(.TextMatrix(intRow, mCol����))
                dbl�ɱ��� = Round(Val(.TextMatrix(intRow, mCol�����)) / .TextMatrix(intRow, mCol����ϵ��), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl�ɱ���� = Round(Val(.TextMatrix(intRow, mCol������)), g_С��λ��.obj_���С��.���С��)
                
                '���˺�:���ۼ۲�����
                dbl���ۼ� = Round(Val(.TextMatrix(intRow, mCol�ۼ�)) / .TextMatrix(intRow, mCol����ϵ��), g_С��λ��.obj_���С��.���ۼ�С��)
                dbl���۽�� = Round(Val(.TextMatrix(intRow, mCol�ۼ۽��)), g_С��λ��.obj_���С��.���С��)
                dbl��� = Round(Val(.TextMatrix(intRow, mCol���)), g_С��λ��.obj_���С��.���С��)
                
                'dbl���ۼ� = Round(Val(.TextMatrix(intRow, mCol�ۼ�)) / .TextMatrix(intRow, mCol����ϵ��), g_С��λ��.obj_ɢװС��.���ۼ�С��)
'                'dbl���۽�� = Round(Val(.TextMatrix(intRow, mCol�ۼ۽��)), g_С��λ��.obj_ɢװС��.���С��)
'                dbl���ۼ� = Round(Val(.TextMatrix(intRow, mcol���ۼ�)), g_С��λ��.obj_ɢװС��.���ۼ�С��)
'                dbl���۽�� = Round(Val(.TextMatrix(intRow, mcol���۽��)), g_С��λ��.obj_ɢװС��.���ۼ�С��)
'                dbl���۲�� = Round(Val(.TextMatrix(intRow, mcol���۲��)), g_С��λ��.obj_ɢװС��.���ۼ�С��)
  
                dbl��� = Round(Val(.TextMatrix(intRow, mCol���)), g_С��λ��.obj_���С��.���С��)
                lng��� = intRow
                
                str������� = Trim(.TextMatrix(intRow, mCol�������))
                str��Ʊ�� = Trim(.TextMatrix(intRow, mCol��Ʊ��))
                strע��֤�� = Trim(.TextMatrix(intRow, mcolע��֤��))
                str��Ʒ���� = Trim(.TextMatrix(intRow, mcol��Ʒ����))
                str��Ʊ���� = IIf(.TextMatrix(intRow, mCol��Ʊ����) = "", "", .TextMatrix(intRow, mCol��Ʊ����))
                dbl��Ʊ��� = Round(Val(.TextMatrix(intRow, mCol��Ʊ���)), g_С��λ��.obj_���С��.���С��)
                
                str�������� = IIf(.TextMatrix(intRow, mcol��������) = "", "", .TextMatrix(intRow, mcol��������))
                str������� = IIf(.TextMatrix(intRow, mcol�������) = "", "", .TextMatrix(intRow, mcol�������))
                str���ʧЧ�� = IIf(.TextMatrix(intRow, mcol���ʧЧ��) = "", "", .TextMatrix(intRow, mcol���ʧЧ��))
                ' Zl_�����⹺_Insert
                gstrSQL = "Zl_�����⹺_Insert("
                '    No_In         In ҩƷ�շ���¼.NO%Type,
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                '    ���_In       In ҩƷ�շ���¼.���%Type,
                gstrSQL = gstrSQL & "" & lng��� & ","
                '    �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
                gstrSQL = gstrSQL & "" & lngStockID & ","
                '    ��ҩ��λid_In In ҩƷ�շ���¼.��ҩ��λid%Type,
                gstrSQL = gstrSQL & "" & lng������λid & ","
                '    ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
                gstrSQL = gstrSQL & "" & lng����ID & ","
                '    ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "'" & str���� & "',"
                '    ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "'" & str���� & "',"
                '    ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str�������� = "", "Null", "to_date('" & Format(str��������, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '    Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '    �������_In   In ҩƷ�շ���¼.�������%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str������� = "", "Null", "to_date('" & Format(str�������, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '    ���Ч��_In   In ҩƷ�շ���¼.���Ч��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str���ʧЧ�� = "", "Null", "to_date('" & Format(str���ʧЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '    ʵ������_In   In ҩƷ�շ���¼.ʵ������%Type := Null,
                gstrSQL = gstrSQL & "" & dblʵ������ & ","
                '    �ɱ���_In     In ҩƷ�շ���¼.�ɱ���%Type := Null,
                gstrSQL = gstrSQL & "" & dbl�ɱ��� & ","
                '    �ɱ����_In   In ҩƷ�շ���¼.�ɱ����%Type := Null,
                gstrSQL = gstrSQL & "" & dbl�ɱ���� & ","
                '    ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "" & dbl���� & ","
                '    ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type := Null,
                gstrSQL = gstrSQL & "" & dbl���ۼ� & ","
                '    ���۽��_In   In ҩƷ�շ���¼.���۽��%Type := Null,
                gstrSQL = gstrSQL & "" & dbl���۽�� & ","
                '    ���_In       In ҩƷ�շ���¼.���%Type := Null,
                gstrSQL = gstrSQL & "" & dbl��� & ","
                '   ���۲��_In   In ҩƷ�շ���¼.���%Type := Null,Ŀǰ������÷��ֶ�
                gstrSQL = gstrSQL & "" & dbl���۲�� & ","
                '    ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
                gstrSQL = gstrSQL & "'" & strժҪ & "',"
                '    ע��֤��_In   In ҩƷ�շ���¼.ע��֤��%Type := Null,
                gstrSQL = gstrSQL & "'" & strע��֤�� & "',"
                '    ������_In     In ҩƷ�շ���¼.������%Type := Null,
                gstrSQL = gstrSQL & "'" & str������ & "',"
                '    �������_In   In Ӧ����¼.�������%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str������� = "", "NULL", "'" & str������� & "'") & ","
                '    ��Ʊ��_In     In Ӧ����¼.��Ʊ��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str��Ʊ�� = "", "NULL", "'" & str��Ʊ�� & "'") & ","
                '    ��Ʊ����_In   In Ӧ����¼.��Ʊ����%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str��Ʊ���� = "", "Null", "to_date('" & Format(str��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '    ��Ʊ���_In   In Ӧ����¼.��Ʊ���%Type := Null,
                gstrSQL = gstrSQL & "" & dbl��Ʊ��� & ","
                '    ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str�������� = "", "Null", "to_date('" & str�������� & "','yyyy-mm-dd HH24:MI:SS')") & ","
                '    �˲���_In     In ҩƷ�շ���¼.��ҩ��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str�˲��� = "", "NULL", "'" & str�˲��� & "'") & ","
                '    �˲�����_In   In ҩƷ�շ���¼.��ҩ����%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str�˲����� = "", "Null", "to_date('" & str�˲����� & "','yyyy-mm-dd HH24:MI:SS')") & ","
                '    ����_In       In ҩƷ�շ���¼.����%Type := 0,
                gstrSQL = gstrSQL & "" & Val(.TextMatrix(intRow, mCol����)) & ","
                '    �˻�_In       In Number := 1
                gstrSQL = gstrSQL & "-1,"
                '  ��ֵ����_In   In varchar2(250)
                gstrSQL = gstrSQL & "'" & GetCostlyInfoStr(intRow) & "'" & ","
                '  ��Ʒ����_In   In ҩƷ�շ���¼.��Ʒ����%Type :=Null
                gstrSQL = gstrSQL & "" & IIf(str��Ʒ���� = "", "NULL", "'" & str��Ʒ���� & "'")
                gstrSQL = gstrSQL & ")"
               
               ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        mstr���ݺ� = chrNo
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveRestore = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'�������
Private Function SaveStrike() As Boolean
    Dim int�д� As Integer
    Dim intԭ��¼״̬ As Integer
    Dim strNo As String
    Dim int��� As Integer
    Dim lng����ID As Long
    Dim dbl�������� As Double
    Dim str������ As String
    Dim str��������  As String
    Dim str��Ʊ�� As String
    Dim str��Ʊ���� As String
    Dim dbl��Ʊ��� As Double
    Dim str������� As String
    Dim str��Ʊ���� As String
    
    Dim intRow As Integer
    Dim rstemp As New ADODB.Recordset
    Dim blnȫ�� As Boolean
    Dim lng�ⷿid As Long, int����� As Integer, lng���� As Long
    Dim i As Integer
    Dim arrSQL As Variant
    Dim n As Long
    
    arrSQL = Array()
    SaveStrike = False
    With mshBill
        '���������������ű�����ԭʼ������ͬ���Ѹ���ļ�¼�����������������˵ĵ���Ҳ�����������
        strNo = Trim(txtNO.Tag)
        lng�ⷿid = cboStock.ItemData(cboStock.ListIndex)
        int����� = Get������(cboStock.ItemData(cboStock.ListIndex))
        
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, mCol��������)) <> 0 Then
                If Not ��ͬ����(Val(.TextMatrix(intRow, mCol����)), Val(.TextMatrix(intRow, mCol��������))) Then
                    MsgBox "������Ϸ��ĳ�����������" & intRow & "�У���", vbInformation, gstrSysName
                    .MsfObj.TopRow = intRow
                    Exit Function
                End If
                If .RowData(intRow) <> 0 Then
                    MsgBox "��" & intRow & "�е����������Ѿ���������������", vbInformation, gstrSysName
                    .MsfObj.TopRow = intRow
                    Exit Function
                End If
                If int����� <> 0 And mint�༭״̬ <> 7 And mbln�˻� = False Then
                    dbl�������� = Round(Val(.TextMatrix(intRow, mCol��������)) * Val(.TextMatrix(intRow, mCol����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                    lng���� = ȡ��������(15, strNo, Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mCol���)))
                    If Check��������(lng�ⷿid, Val(.TextMatrix(intRow, 0)), lng����, dbl��������, int�����) = False Then Exit Function
                End If
            End If
        Next
        
        str������ = UserInfo.�û���
        str�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        intԭ��¼״̬ = mint��¼״̬
        
        On Error GoTo ErrHandle
        
        int�д� = 0
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�

'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And (Val(.TextMatrix(intRow, mCol��������)) <> 0 Or mint�༭״̬ = 7) Then
                int�д� = int�д� + 1
                int��� = Val(.TextMatrix(intRow, mCol���))
                
                lng����ID = Val(.TextMatrix(intRow, 0))
                dbl�������� = Round(IIf(mbln�˻�, -1, 1) * Val(.TextMatrix(intRow, mCol��������)) * Val(.TextMatrix(intRow, mCol����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                
                str������� = Trim(.TextMatrix(intRow, mCol�������))
                str��Ʊ�� = Trim(.TextMatrix(intRow, mCol��Ʊ��))
                str��Ʊ���� = Trim(.TextMatrix(intRow, mcol��Ʊ����))
                str��Ʊ���� = IIf(.TextMatrix(intRow, mCol��Ʊ����) = "", "", .TextMatrix(intRow, mCol��Ʊ����))
                If mint�༭״̬ = 7 Then
                    '����ԭ���ݵķ�Ʊ
                    dbl��Ʊ��� = GetTotale��Ʊ���(strNo, lng����ID, int���)
                Else
                    dbl��Ʊ��� = Round(IIf(mbln�˻�, -1, 1) * Val(.TextMatrix(intRow, mCol��Ʊ���)), g_С��λ��.obj_ɢװС��.���С��)
                End If
                blnȫ�� = False
                If dbl�������� = Round(IIf(mbln�˻�, -1, 1) * Val(.TextMatrix(intRow, mCol����)) * Val(.TextMatrix(intRow, mCol����ϵ��)), g_С��λ��.obj_ɢװС��.����С��) Then
                    blnȫ�� = True
                End If
                
                If mint�༭״̬ = 7 Then
                    blnȫ�� = True
                End If
                
                
                ' Zl_�����⹺_Strike
                gstrSQL = "ZL_�����⹺_STRIKE("
                '  �д�_In       In Integer,
                gstrSQL = gstrSQL & "" & int�д� & ","
                '  ԭ��¼״̬_In In ҩƷ�շ���¼.��¼״̬%Type,
                gstrSQL = gstrSQL & "" & intԭ��¼״̬ & ","
                '  No_In         In ҩƷ�շ���¼.NO%Type,
                gstrSQL = gstrSQL & "'" & strNo & "',"
                '  ���_In       In ҩƷ�շ���¼.���%Type,
                gstrSQL = gstrSQL & "" & int��� & ","
                '  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
                gstrSQL = gstrSQL & "" & lng����ID & ","
                '  ��������_In   In ҩƷ�շ���¼.ʵ������%Type,
                gstrSQL = gstrSQL & "" & dbl�������� & ","
                '  ������_In     In ҩƷ�շ���¼.������%Type,
                gstrSQL = gstrSQL & "'" & str������ & "',"
                '  ��������_In   In ҩƷ�շ���¼.��������%Type,
                gstrSQL = gstrSQL & "to_date('" & Format(mstr�������, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'),"
                '  �������_In   In Ӧ����¼.�������%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str������� = "", "null", "'" & str������� & "'") & ","
                '  ��Ʊ��_In     In Ӧ����¼.��Ʊ��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str��Ʊ�� = "", "null", "'" & str��Ʊ�� & "'") & ","
                '  ��Ʊ����_In   In Ӧ����¼.��Ʊ����%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str��Ʊ���� = "", "Null", "to_date('" & Format(str��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  ��Ʊ���_In   In Ӧ����¼.��Ʊ���%Type := Null,
                gstrSQL = gstrSQL & "" & dbl��Ʊ��� & ","
                '  ȫ������_In   In ҩƷ�շ���¼.ʵ������%Type := 0 --���ڲ������
                gstrSQL = gstrSQL & "" & IIf(blnȫ��, 1, 0) & ","
                '  �������_In   In Number := 0 --������˱�־:1-�������,0-����
                gstrSQL = gstrSQL & "" & IIf(mint�༭״̬ = 7, 1, 0) & ","
                'ժҪ_in in ҩƷ�շ���¼.ժҪ%type
                gstrSQL = gstrSQL & "'" & txtժҪ.Text & "',"
                '  ��Ʊ����_In      Ӧ����¼.��Ʊ����%type :=null
                gstrSQL = gstrSQL & IIf(str��Ʊ���� = "", "NULL)", "'" & str��Ʊ���� & "')")
                Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
            End If
            recSort.MoveNext
        Next
        If mint�༭״̬ <> 7 Then gcnOracle.BeginTrans
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "SaveCard")
            Next
        If mint�༭״̬ <> 7 Then gcnOracle.CommitTrans
        If int�д� = 0 Then
            ShowMsgBox "û��ѡ��һ���������������������ܳ��������飡"
            Exit Function
        End If
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveStrike = True
    Exit Function
ErrHandle:
    If mint�༭״̬ <> 7 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveRecipe() As Boolean
    Dim chrNo As String
    Dim lng��� As Long
    Dim str��Ʊ�� As String
    Dim str��Ʊ���� As String
    Dim str��Ʊ���� As String
    Dim dbl��Ʊ��� As Double
    Dim cllTemp As New Collection
    Dim intRow As Integer
    Dim n As Long
    
    SaveRecipe = False
    '����Ƿ����빩ҩ��λ
    If Val(txtProvider.Tag) = 0 Then
        MsgBox "��ѡ���������ϵĹ�Ӧ�̣�", vbInformation, gstrSysName
        txtProvider.SetFocus
        Exit Function
    End If

    With mshBill
        chrNo = Trim(txtNO.Tag)
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
            
                lng��� = Val(.TextMatrix(intRow, mCol���))
                
                str��Ʊ�� = .TextMatrix(intRow, mCol��Ʊ��)
                str��Ʊ���� = Trim(.TextMatrix(intRow, mcol��Ʊ����))
                str��Ʊ���� = IIf(.TextMatrix(intRow, mCol��Ʊ����) = "", "", .TextMatrix(intRow, mCol��Ʊ����))
                dbl��Ʊ��� = Round(Val(.TextMatrix(intRow, mCol��Ʊ���)), g_С��λ��.obj_ɢװС��.���С��)
                
                '    NO_IN       IN ҩƷ�շ���¼.NO%TYPE := NULL,
                '    ��¼״̬_IN     IN ҩƷ�շ���¼.��¼״̬%type:=NULL,
                '    ���_IN     IN ҩƷ�շ���¼.���%TYPE:=NULL,
                '    ��Ʊ��_IN       IN Ӧ����¼.��Ʊ��%TYPE := NULL,
                '    ��Ʊ����_IN     IN Ӧ����¼.��Ʊ����%TYPE := NULL,
                '    ��Ʊ���_IN     IN Ӧ����¼.��Ʊ���%TYPE := NULL,
                '    ��ҩ��λ_IN     in Ӧ����¼.��λID%TYPE:=0,
                '    ��Ʊ����_in     in Ӧ����¼.��Ʊ����%type := null
                
                gstrSQL = "zl_�����⹺��Ʊ��Ϣ_UPDATE( "
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                gstrSQL = gstrSQL & "" & mint��¼״̬ & ","
                gstrSQL = gstrSQL & "" & lng��� & ","
                gstrSQL = gstrSQL & "'" & str��Ʊ�� & "',"
                gstrSQL = gstrSQL & "" & IIf(str��Ʊ���� = "", "Null", "to_date('" & Format(str��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                gstrSQL = gstrSQL & "" & dbl��Ʊ��� & ","
                gstrSQL = gstrSQL & "" & Val(txtProvider.Tag) & ","
                gstrSQL = gstrSQL & IIf(str��Ʊ���� = "", "NULL", "'" & str��Ʊ���� & "'") & ")"
                AddArray cllTemp, gstrSQL
            End If
            recSort.MoveNext
        Next
        err = 0: On Error GoTo ErrHandle
        ExecuteProcedureArrAy cllTemp, mstrCaption
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveRecipe = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveRegist() As Boolean
    Dim chrNo As String
    Dim lng����ID As Long
    Dim strע��֤�� As String
    Dim cllTemp As New Collection
    Dim intRow As Integer
    Dim n As Long
    
    SaveRegist = False

    With mshBill
        chrNo = Trim(txtNO.Tag)
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
            
                lng����ID = .TextMatrix(intRow, 0)
                strע��֤�� = .TextMatrix(intRow, mcolע��֤��)
                
                gstrSQL = "Zl_�����⹺_�޸�ע��֤��( "
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                gstrSQL = gstrSQL & "" & lng����ID & ","
                gstrSQL = gstrSQL & "'" & strע��֤�� & "')"
                AddArray cllTemp, gstrSQL
            End If
            recSort.MoveNext
        Next
        err = 0: On Error GoTo ErrHandle
        ExecuteProcedureArrAy cllTemp, mstrCaption
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveRegist = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub ��ʾ�ϼƽ��()
    Dim curTotal As Double, Cur���ʽ�� As Double, Cur���ʲ�� As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur���ʽ�� = 0: Cur���ʲ�� = 0
    
    With mshBill
        For intLop = 1 To .Rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mCol������))
            Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mCol�ۼ۽��))
            Cur���ʲ�� = Cur���ʲ�� + Val(.TextMatrix(intLop, mCol���))
        Next
    End With
    
'    Cur���ʲ�� = Cur���ʽ�� - curTotal
    lblPurchasePrice.Caption = "������ϼƣ�" & Format(curTotal, mFMT.FM_���)
    lblSalePrice.Caption = "�ۼ۽��ϼƣ�" & Format(Cur���ʽ��, mFMT.FM_���)
    lblDifference.Caption = "��ۺϼƣ�" & Format(Cur���ʲ��, mFMT.FM_���)
End Sub

Private Sub ��ʾ�����()
    Dim rstemp As New ADODB.Recordset
    Dim dbl���� As Double
    Dim str��λ As String, strUnit As String, strQuantity As String
    Dim intID As Long, lng���� As Long
    
    On Error GoTo ErrHandle
    If mshBill.TextMatrix(mshBill.Row, mCol����) = "" Then
        stbThis.Panels(2).Text = ""
        Exit Sub
    End If
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
    intID = mshBill.TextMatrix(mshBill.Row, 0)
    lng���� = Val(mshBill.TextMatrix(mshBill.Row, mCol����))
    If mintUnit = 0 Then
            strQuantity = "a.��������"
    Else
            strQuantity = "a.��������/b.����ϵ��"
    End If

    
    gstrSQL = "" & _
    "   Select b.����ID," & IIf(mintUnit = 0, "c.���㵥λ", "b.��װ��λ") & " as ��λ, Sum(nvl(" & strQuantity & ",0)) as ���� " & _
    "   From ҩƷ��� a,�������� b,�շ���ĿĿ¼ c " & _
    "   Where a.����=1 and a.ҩƷid=b.����id and b.����id=c.id " & _
    "         and a.��������<>0 And a.�ⷿID=[1] and b.����ID=[2]  " & IIf(mint�༭״̬ = 8 Or mbln�˻�, " and nvl(����,0)=[3]", "") & _
    "   Group by b.����ID," & _
            IIf(mintUnit = 0, "c.���㵥λ", "b.��װ��λ")
    
   Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), intID, lng����)
   With rstemp
        If .EOF Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        
        dbl���� = IIf(IsNull(!����), 0, !����)
        stbThis.Panels(2).Text = "���������ϵ�ǰ�����Ϊ[" & Format(dbl����, mFMT.FM_����) & "]" & zlStr.Nvl(!��λ)
        
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtTypeVar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Select Case Me.lblType.Tag
        Case 1, 3, 4
            Call Comm_Selecter(Me.txtTypeVar.Text, Me.lblType.Tag + 2)
        Case Else
            Call Comm_Selecter("%" & Me.txtTypeVar.Text & "%", Me.lblType.Tag + 2)
        End Select
        Me.vsfCostlyInfo.SetFocus
    Else
        If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
        If InStr("1,3,4", Me.lblType.Tag) Then
            If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub Txt�Ӽ���_GotFocus()
    Txt�Ӽ���.SelStart = 0
    Txt�Ӽ���.SelLength = Len(Txt�Ӽ���)
End Sub

Private Sub Txt�Ӽ���_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then Call CmdYes_Click
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
    ImeLanguage True
    zlControl.TxtSelAll txtժҪ
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txtժҪ_LostFocus()
    ImeLanguage False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'ȡ���ݿ��з�Ʊ�ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
Private Function Get��Ʊ��Len() As Integer
    Dim rstemp As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select ��Ʊ�� from Ӧ����¼ where rownum<1 "
    zlDatabase.OpenRecordset rstemp, gstrSQL, "ȡ�ֶγ���"
    Get��Ʊ��Len = rstemp.Fields(0).DefinedSize
    rstemp.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




'ȡ���ݿ������ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
Private Function GetBatchNoLen() As Integer
    Dim rstemp As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select ���� from ҩƷ�շ���¼ where rownum<1 "
    zlDatabase.OpenRecordset rstemp, gstrSQL, "ȡ�ֶγ���"
    GetBatchNoLen = rstemp.Fields(0).DefinedSize
    rstemp.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ���ɱ���()
    Dim dbl�ɱ��� As Double, dbl���ۼ� As Double, dbl�ۼ� As Double
    
    '����ɱ��۱����ۼۻ��ߣ���ʾ�û�
    With mshBill
        If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
        dbl�ɱ��� = Format(Val(.TextMatrix(.Row, mCol�����)), "#####0.00;-#####0.00;0;")
        dbl�ۼ� = Format(Val(.TextMatrix(.Row, mCol�ۼ�)), "#####0.00;-#####0.00;0;")
        dbl���ۼ� = Format(Val(.TextMatrix(.Row, mcol���ۼ�)) * Val(.TextMatrix(.Row, mCol����ϵ��)), "#####0.00;-#####0.00;0;")
    End With
    If dbl�ɱ��� > dbl�ۼ� Then
        MsgBox "���ѣ����������ϵĳɱ��۱��ۼۻ��ߣ�", vbInformation, gstrSysName
    End If
    If dbl�ɱ��� > dbl���ۼ� Then
        MsgBox "���ѣ����������ϵĳɱ��۱����ۼۻ��ߣ�", vbInformation, gstrSysName
    End If
End Sub

Private Function CopyCard() As String
    Dim intRow As Integer, intUpdate As Integer, str������� As String
    Dim dblԭ���� As Double, dbl������ As Double
    Dim dbl����� As Double, dbl������ As Double, dbl��� As Double, dbl���۽�� As Double, dbl���� As Double
    Dim dbl�ɹ��� As Double, dbl�ۼ� As Double, dbl��Ʊ��� As Double, dbl���۲�� As Double
    Dim str��Ʊ�� As String, str��Ʊ���� As String, str��Ʊ���� As String
    Dim lng��� As Long
    
    
    Dim strNo As Variant
    On Error GoTo ErrHand
    
    strNo = sys.GetNextNo(68, cboStock.ItemData(cboStock.ListIndex))
    If IsNull(strNo) Then Exit Function
    
    intUpdate = 0
    CopyCard = ""
    
    '���Ʋ����µ���
    ' ����_IN,NO_IN,NewNO_IN
    
'    gstrSQL = "zl_��������_billcopy(15,'" & txtNO.Text & "','" & StrNo & "','" & UserInfo.�û��� & "')"
    gstrSQL = "zl_��������_billcopy(15,'" & txtNO.Tag & "','" & strNo & "')"
    zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
    '�ɹ��ۣ����ʣ�����ۣ�������ۼۣ���Ʊ�ţ���Ʊ���ڣ���Ʊ���
    
    '�޸Ľ���ۡ��������ۣ�Ҫ���ǵ�������˳����ĵ��ݣ���ʱ��Ҫ�޸Ľ���ۡ��������ۣ�
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                dbl����� = Val(.TextMatrix(intRow, mCol�����))
                dbl������ = IIf(mbln�˻� = True, "-" & Val(.TextMatrix(intRow, mCol������)), Val(.TextMatrix(intRow, mCol������)))
                dbl��� = Val(.TextMatrix(intRow, mCol���))
                
                '���˺�:���ۼ۴���
                dbl�ۼ� = Val(.TextMatrix(intRow, mcol���ۼ�))
                dbl���۽�� = IIf(mbln�˻� = True, (-1) * Val(.TextMatrix(intRow, mcol���۽��)), Val(.TextMatrix(intRow, mcol���۽��)))
                dbl���۲�� = IIf(mbln�˻� = True, (-1) * Val(.TextMatrix(intRow, mcol���۲��)), Val(.TextMatrix(intRow, mcol���۲��)))
                dbl��� = Round(IIf(mbln�˻� = True, (-1) * Val(.TextMatrix(intRow, mcol���۲��)), Val(.TextMatrix(intRow, mcol���۲��))), g_С��λ��.obj_ɢװС��.���ۼ�С��)
                dbl���۲�� = Round(dbl���۲�� - dbl���, g_С��λ��.obj_ɢװС��.���ۼ�С��)
                
                dbl���� = Val(.TextMatrix(intRow, mCol����))
                
                str������� = Trim(.TextMatrix(intRow, mCol�������))
                str��Ʊ�� = Trim(.TextMatrix(intRow, mCol��Ʊ��))
                str��Ʊ���� = Trim(.TextMatrix(intRow, mcol��Ʊ����))
                str��Ʊ���� = Trim(IIf(.TextMatrix(intRow, mCol��Ʊ����) = "", "", .TextMatrix(intRow, mCol��Ʊ����)))
                dbl��Ʊ��� = IIf(mbln�˻� = True, (-1) * Val(.TextMatrix(intRow, mCol��Ʊ���)), Val(.TextMatrix(intRow, mCol��Ʊ���)))
                If dbl��Ʊ��� = 0 Then dbl��Ʊ��� = IIf(str������� <> "", dbl������, 0)
                Call Get����(txtNO.Tag, Val(.TextMatrix(intRow, mCol���)), dblԭ����)
                lng��� = intRow
                
                If Get����(strNo, Val(.TextMatrix(intRow, mCol���)), dbl������) Then
                    If Abs(dbl������) > 0 Then
                        '��������
                        dbl����� = Round(dbl����� / Val(.TextMatrix(intRow, mCol����ϵ��)), g_С��λ��.obj_���С��.�ɱ���С��)
                        'dbl�ۼ� = Round(dbl�ۼ� / Val(.TextMatrix(intRow, mCol����ϵ��)), g_С��λ��.obj_ɢװС��.���ۼ�С��)
                        dbl������ = Round(dbl������ * dbl������ / dblԭ����, g_С��λ��.obj_ɢװС��.���С��)
                        dbl��� = Round(dbl��� * dbl������ / dblԭ����, g_С��λ��.obj_ɢװС��.���С��)
                        dbl���۽�� = Round(dbl���۽�� * dbl������ / dblԭ����, g_С��λ��.obj_ɢװС��.���С��)
                        dbl���۲�� = Round(dbl���۲�� * dbl������ / dblԭ����, g_С��λ��.obj_ɢװС��.���С��)
                        
                        dbl��Ʊ��� = Round(dbl��Ʊ��� * dbl������ / dblԭ����, g_С��λ��.obj_ɢװС��.���С��)
                        
                        
                        '����
                        gstrSQL = "zl_Bill_������Ϣ(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol���)) & ",'�ɱ���','" & dbl����� & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_������Ϣ(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol���)) & ",'�ɱ����','" & dbl������ & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_������Ϣ(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol���)) & ",'���','" & dbl��� & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_������Ϣ(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol���)) & ",'���ۼ�','" & dbl�ۼ� & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_������Ϣ(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol���)) & ",'���۽��','" & dbl���۽�� & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_������Ϣ(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol���)) & ",'�÷�','" & dbl���۲�� & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        
                        gstrSQL = "zl_Bill_������Ϣ(15,'" & strNo & "'," & Val(.TextMatrix(intRow, mCol���)) & ",'����','" & dbl���� & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        
                        
                        gstrSQL = "zl_Bill_����Ӧ����¼('" & strNo & "'," & Val(.TextMatrix(intRow, mCol���)) & ",'��Ʊ���','" & dbl��Ʊ��� & "',5)"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_����Ӧ����¼('" & strNo & "'," & Val(.TextMatrix(intRow, mCol���)) & ",'�������'," & IIf(str������� = "", "NULL", "''" & str������� & "''") & ",5)"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_����Ӧ����¼('" & strNo & "'," & Val(.TextMatrix(intRow, mCol���)) & ",'��Ʊ��'," & IIf(str��Ʊ�� = "", "NULL", "''" & str��Ʊ�� & "''") & ",5)"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_����Ӧ����¼('" & strNo & "'," & Val(.TextMatrix(intRow, mCol���)) & ",'��Ʊ����'," & IIf(str��Ʊ�� = "", "NULL", "''" & str��Ʊ���� & "''") & ",5)"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        gstrSQL = "zl_Bill_����Ӧ����¼('" & strNo & "'," & Val(.TextMatrix(intRow, mCol���)) & ",'��Ʊ����'," & IIf(str��Ʊ���� = "", "NULL", "to_date('" & str��Ʊ���� & "','yyyy-mm-dd')") & ",5)"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        
                        intUpdate = intUpdate + 1
                    End If
                End If
            End If
        Next
    End With
    gstrSQL = "zl_���ϲ������_update(15,'" & strNo & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
    
    If intUpdate = 0 Then
        MsgBox "�޷���ɲ�����ˣ���Ϊ�õ����ѱ�ȫ��������", vbInformation, gstrSysName
        Exit Function
    End If
    CopyCard = strNo
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get����(ByVal strNo As String, ByVal int��� As Integer, dbl���� As Double) As Boolean
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select Nvl(ʵ������,0) ���� " & _
        "   From ҩƷ�շ���¼" & _
        "   Where ����=15 And NO=[1]  And ���=[2]"
    
    
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strNo, int���)
    
    If rstemp.EOF Then Exit Function
    dbl���� = rstemp!����
    Get���� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ʱ�۲������ۼ�(ByVal lng����ID As Long, ByVal sin�ɹ��� As Double, ByVal sin�ӳ��� As Double, _
    Optional LngLastRow As Long = -1, Optional sng�ۼ� As Double = -99999999) As Double
    '------------------------------------------------------------------------------------------------------
    '����:����ָ���۸���۱ȼ����ʱ�۲��ϵĲ���������
    '���:lng����ID-����ID
    '     sin�ɹ���-�ɹ��۸�
    '     sin�ӳ���-�ӳ���(�������0,ͬʱ�ִ���dbl���ۼ�,�򽫰���������ۼ۽��м���)
    '     LngLastRow-���ݵ��к�
    '     sng�ۼ�-��������ۼ�
    '����:
    '����:���ۼ�
    '�޸���:���˺�
    '�޸�ʱ��:2007/2/25
    '------------------------------------------------------------------------------------------------------
       'ʱ�۲������ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
    '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
    '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
    
    Dim sin���ۼ� As Double, sinָ�����ۼ� As Double, sin��������� As Double
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ָ�����ۼ�,Nvl(���������,100) ��������� From �������� Where ����ID=[1]"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ�����ۼ�", lng����ID)
    If rstemp.EOF Then Exit Function
    
    sinָ�����ۼ� = rstemp!ָ�����ۼ�
    sin��������� = rstemp!���������
    
    ʱ�۲������ۼ� = 0
    If sin��������� = 100 Then Exit Function
    
    '���δ��ָ���ۣ��Ͳ�������������
    If sinָ�����ۼ� = 0 Then Exit Function
    If LngLastRow = -1 Then LngLastRow = mshBill.Row
    
    If mint�༭״̬ = 8 Or mbln�˻� Then
        '������˻����򰴳���ķ�ʽ�����ۼ�
        gstrSQL = " Select Nvl(ʵ������,0) ʵ������,Nvl(ʵ�ʽ��,0) ʵ�ʽ�� From ҩƷ��� " & _
                  " Where ����=1 And ҩƷID=[1] And �ⷿID=[2] And Nvl(����,0)=[3]"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������ۼ�", lng����ID, cboStock.ItemData(cboStock.ListIndex), Val(mshBill.TextMatrix(LngLastRow, mCol����)))
        
        
        If rstemp.RecordCount = 0 Then
            ShowMsgBox "�������Ͽ�����ݴ���δ�ҵ�ָ���������ϵĿ���¼����"
            Exit Function
        End If
        '�϶���������û�������Ļ����޷�����˴�
        ʱ�۲������ۼ� = rstemp!ʵ�ʽ�� / rstemp!ʵ������ * Val(mshBill.TextMatrix(LngLastRow, mCol����ϵ��))
    Else
        If sng�ۼ� <> -99999999 And sin�ӳ��� = 0 Then
            sin���ۼ� = sng�ۼ�
        Else
            sin���ۼ� = sin�ɹ��� * (1 + sin�ӳ���)
        End If
        
        If sin���ۼ� / Val(mshBill.TextMatrix(LngLastRow, mCol����ϵ��)) >= sinָ�����ۼ� Then Exit Function
        sinָ�����ۼ� = sinָ�����ۼ� * Val(mshBill.TextMatrix(LngLastRow, mCol����ϵ��))
        
        ʱ�۲������ۼ� = (sinָ�����ۼ� - sin���ۼ�) * (1 - sin��������� / 100)
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ����ӳ���(ByVal lng����ID As Long, ByVal sin���ۼ� As Double, ByVal sin�ɱ��� As Double) As Double
    Dim sinָ�����ۼ� As Double, sin��������� As Double
    Dim rstemp As New ADODB.Recordset
    '�������ۼ۷���ɱ���,����ʱ�����Ĺ�ʽ�ı仯,����ԭ������ӳ��ʵĹ�ʽ��Ч,�����¼���
    'ԭ��ʽ:(���ۼ�/�ɱ���-1)*100
    '�ֹ�ʽ������:�������ۼ��ǰ��ӳ����������,�ټ������������ǲ��ֽ��,���ʵ�ʰ��ӳ�����������ۼ�=ָ�����ۼ�-(ָ�����ۼ�-���ۼ�)/���������
    '������ԭ��ʽ���ʵ�ʵļӳ���
    ����ӳ��� = 0.15
    
    On Error GoTo ErrHandle
    gstrSQL = "Select a.ָ�����ۼ�,Nvl(a.���������,100) ���������,Nvl(b.�Ƿ���,0) ʱ�� From �������� A, �շ���ĿĿ¼ b Where a.����ID=b.id  and b.ID=[1]"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ�����ۼ�", lng����ID)
    If rstemp.EOF Then Exit Function
    
    sinָ�����ۼ� = rstemp!ָ�����ۼ�
    sin��������� = rstemp!���������
    If rstemp!ʱ�� = 0 Then Exit Function
    
'    If mbln�ֶμӳ��� Then
'            ����ӳ��� = Get�ֶμӳ���(sin�ɱ���)
'    Else
        
        'ָ�����ۼ�-(ָ�����ۼ�-���ۼ�)/���������
        sinָ�����ۼ� = sinָ�����ۼ� * Val(mshBill.TextMatrix(mshBill.Row, mCol����ϵ��))
        If sin��������� <> 100 And sin��������� > 0 Then
            sin���ۼ� = sinָ�����ۼ� - (sinָ�����ۼ� - sin���ۼ�) / sin��������� * 100
        Else
            sin���ۼ� = sinָ�����ۼ� - (sinָ�����ۼ� - sin���ۼ�)
        End If
        ����ӳ��� = (sin���ۼ� / sin�ɱ��� - 1) * 100
   ' End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function У�����ۼ�(ByVal sin���ۼ� As Double, Optional LngLastRow As Long = -1) As Double
    '�õ�����ǰ��λϵ�����������ָ�����ۼۣ����ʱ�����ļ�����������ۼ۴���ָ�����ۼۣ���ָ�����ۼ�Ϊ׼
    Dim sinָ�����ۼ� As Double
    Dim rstemp As New ADODB.Recordset
    If LngLastRow = -1 Then LngLastRow = mshBill.Row
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ָ�����ۼ�,Nvl(���������,100) ��������� From �������� Where ����ID=[1]"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ�����ۼ�", Val(mshBill.TextMatrix(LngLastRow, 0)))
    If rstemp.EOF Then Exit Function
    
    sinָ�����ۼ� = Val(zlStr.Nvl(rstemp!ָ�����ۼ�))
    sinָ�����ۼ� = sinָ�����ۼ� * Val(mshBill.TextMatrix(LngLastRow, mCol����ϵ��))
    If sinָ�����ۼ� = 0 Then sinָ�����ۼ� = sin���ۼ�
    
    If Val(sin���ۼ�) > Val(sinָ�����ۼ�) And Not mbln��ǿ�ƿ���ָ���۸� Then
'        MsgBox "�ۼۣ�" & sin���ۼ� & "����" & "ָ���ۼۣ�" & sinָ�����ۼ� & "��ǿ�и�Ϊָ���ۼۡ�", vbInformation, gstrSysName
        У�����ۼ� = sinָ�����ۼ�
    Else
        У�����ۼ� = sin���ۼ�
    End If
    'У�����ۼ� = IIf(sin���ۼ� > sinָ�����ۼ�, sinָ�����ۼ�, sin���ۼ�)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetColumnByUserDefine()
    Dim intColumns As Integer
    Dim strColumn_UnSelected As String
    Dim strColumn_Selected As String
    Dim strColumn_All As String
    Dim arrColumn_All, arrColumn_Selected, arrColumn_UnSelected
    Dim strCol As String
    Dim intCol As Integer, intCols As Integer
    Dim i As Integer
    
    strColumn_Selected = zlDatabase.GetPara("ѡ����", glngSys, mlngModule)
    strColumn_UnSelected = zlDatabase.GetPara("������", glngSys, mlngModule)
     
    'strColumn_All = "����,0|���,1|����,1|����,0|��������,1|�������,1|���ʧЧ��,1|Ч��,0|��λ,1|����,0|ָ��������,1|�ɹ���,1|����,1|" & _
                "�ӳ���,0|�����,0|������,0|�ۼ�,0|�ۼ۽��,0|���,0|��Ʊ��,0|��Ʊ����,0|��Ʊ����,0|��Ʊ���,0"
    
    strColumn_All = ""
    '��װ��ȱʡ����
    i = 1: mCol�к� = i:
    i = i + 1: mCol���� = i: strColumn_All = strColumn_All & "����," & i & "|"
    i = i + 1: mCol��� = i:
    i = i + 1: mCol��Ʒ�� = i: strColumn_All = strColumn_All & "��Ʒ��," & i & "|"
    i = i + 1: mCol��� = i: strColumn_All = strColumn_All & "���," & i & "|"
    i = i + 1: mColԭ���� = i:
    i = i + 1: mColԭ���� = i:
    i = i + 1: mCol����ϵ�� = i:
    i = i + 1: mCol���� = i:
    i = i + 1: mCol���� = i: strColumn_All = strColumn_All & "����," & i & "|"
    i = i + 1: mcol��׼�ĺ� = i: strColumn_All = strColumn_All & "��׼�ĺ�," & i & "|"
    i = i + 1: mCol��λ = i: strColumn_All = strColumn_All & "��λ," & i & "|"
    i = i + 1: mCol���� = i: strColumn_All = strColumn_All & "����," & i & "|"
    i = i + 1: mcol�������� = i: strColumn_All = strColumn_All & "��������," & i & "|"
    i = i + 1: mColЧ�� = i: strColumn_All = strColumn_All & "Ч��," & i & "|"
    i = i + 1: mcolһ���Բ��� = i:
    i = i + 1: mcol������� = i:
    i = i + 1: mcol���Ч�� = i:
    i = i + 1: mcol������� = i: strColumn_All = strColumn_All & "�������," & i & "|"
    i = i + 1: mcol���ʧЧ�� = i: strColumn_All = strColumn_All & "���ʧЧ��," & i & "|"
    i = i + 1: mcolע��֤�� = i: strColumn_All = strColumn_All & "ע��֤��," & i & "|"
    i = i + 1: mcolע��֤��Ч�� = i: strColumn_All = strColumn_All & "ע��֤��Ч��," & i & "|"
    i = i + 1: mcol�ڲ����� = i
    i = i + 1: mcol����ID = i
    i = i + 1: mcol��Ʒ���� = i
    If gblnCode = True Then
        strColumn_All = strColumn_All & "��Ʒ����," & i & "|"
    End If
    
    i = i + 1: mCol���� = i: strColumn_All = strColumn_All & "����," & i & "|"
    i = i + 1: mCol�������� = i
    i = i + 1: mCol���� = i:
    i = i + 1: mColָ�������� = i: strColumn_All = strColumn_All & "ָ��������," & i & "|"
    i = i + 1: mCol�ɹ��� = i: strColumn_All = strColumn_All & "�ɹ���," & i & "|"
    i = i + 1: mCol���� = i: strColumn_All = strColumn_All & "����," & i & "|"
    i = i + 1: mCol����� = i: strColumn_All = strColumn_All & "�����," & i & "|"
    i = i + 1: mCol������ = i: strColumn_All = strColumn_All & "������," & i & "|"
    i = i + 1: mcol�ӳ��� = i: strColumn_All = strColumn_All & "�ӳ���," & i & "|"
    i = i + 1: mCol�ۼ� = i: strColumn_All = strColumn_All & "�ۼ�," & i & "|"
    i = i + 1: mCol�ۼ۽�� = i: strColumn_All = strColumn_All & "�ۼ۽��," & i & "|"
    i = i + 1: mCol��� = i: strColumn_All = strColumn_All & "���," & i & "|"
    
    '���˺�:���ۼ۴���
    i = i + 1: mcol���ۼ� = i: strColumn_All = strColumn_All & "���ۼ�," & i & "|"
    i = i + 1: mcol���۵�λ = i: strColumn_All = strColumn_All & "���۵�λ," & i & "|"
    i = i + 1: mcol���۽�� = i: strColumn_All = strColumn_All & "���۽��," & i & "|"
    i = i + 1: mcol���۲�� = i: strColumn_All = strColumn_All & "���۲��," & i & "|"
    i = i + 1: mCol������� = i: strColumn_All = strColumn_All & "�������," & i & "|"
    i = i + 1: mCol���ս��� = i: strColumn_All = strColumn_All & "���ս���," & i & "|"
    i = i + 1: mCol��Ʊ�� = i: strColumn_All = strColumn_All & "��Ʊ��," & i & "|"
    i = i + 1: mcol��Ʊ���� = i: strColumn_All = strColumn_All & "��Ʊ����," & i & "|"
    i = i + 1: mCol��Ʊ���� = i: strColumn_All = strColumn_All & "��Ʊ����," & i & "|"
    i = i + 1: mCol��Ʊ��� = i: strColumn_All = strColumn_All & "��Ʊ���," & i
    
    
    If strColumn_Selected = "" Then Exit Sub
    
    '�����û����õ�����˳��
    arrColumn_All = Split(strColumn_All, "|")
    arrColumn_Selected = Split(strColumn_Selected, "|")
    intCols = UBound(arrColumn_Selected)
    
    For intCol = 0 To intCols
        Call SetColumnValue(arrColumn_Selected(intCol), Split(arrColumn_All(intCol), ",")(1))
    Next
    
    For intCol = 0 To UBound(arrColumn_All)
       strCol = "|" & Split(arrColumn_All(intCol), ",")(0) & "|"
       If InStr("|" & strColumn_Selected & "|", strCol) = 0 Then
           '��ѡ�����в�����,��϶�������,������δѡ������û�е�,ֻ��ʾ��������,��Ҫ����
           If InStr("|" & strColumn_UnSelected & "|", strCol) = 0 Then
               strColumn_UnSelected = strColumn_UnSelected & "|" & Split(arrColumn_All(intCol), ",")(0)
           End If
       End If
    Next
     
    '��δѡ����е��п�����Ϊ�㣬��������Ϊ5��������ѡ��
    If strColumn_UnSelected = "" Then Exit Sub
    If Left(strColumn_UnSelected, 1) = "|" Then strColumn_UnSelected = Mid(strColumn_UnSelected, 2)
    intCol = intCols + 1
    intColumns = 0
    arrColumn_UnSelected = Split(strColumn_UnSelected, "|")
    intCols = UBound(arrColumn_All)
    For intCol = intCol To intCols
        If UBound(arrColumn_UnSelected) >= intColumns Then
            Call SetColumnValue(arrColumn_UnSelected(intColumns), Split(arrColumn_All(intCol), ",")(1), False)
                intColumns = intColumns + 1
        Else
            Call SetColumnValue(Split(arrColumn_All(intCol), ",")(0), Split(arrColumn_All(intCol), ",")(1), False)
        End If
    Next
End Sub

Private Sub SetColumnValue(ByVal str���� As String, ByVal intValue As Integer, Optional ByVal blnShow As Boolean = True)
    Select Case str����
    Case "�к�"
        mCol�к� = intValue
    Case "����", "����"
        mCol���� = intValue
    Case "���"
        mCol��� = intValue
    Case "��Ʒ��"
        mCol��Ʒ�� = intValue
    Case "���"
        mCol��� = intValue
    Case "ԭ����"
        mColԭ���� = intValue
    Case "ԭ����"
        mColԭ���� = intValue
    Case "����ϵ��"
        mCol����ϵ�� = intValue
    Case "����"
        mCol���� = intValue
    Case "����"
        mCol���� = intValue
    Case "��λ"
        mCol��λ = intValue
    Case "����"
        mCol���� = intValue
    Case "��������"
        mcol�������� = intValue
    Case "Ч��"
        mColЧ�� = intValue
    Case "��׼�ĺ�"
        mcol��׼�ĺ� = intValue
    Case "����"
        mCol���� = intValue
    Case "��������"
        mCol�������� = intValue
    Case "ָ��������"
        mColָ�������� = intValue
    Case "����"
        mCol���� = intValue
    Case "�ɹ���"
        mCol�ɹ��� = intValue
    Case "�����"
        mCol����� = intValue
    Case "������"
        mCol������ = intValue
    Case "�ۼ�"
        mCol�ۼ� = intValue
    Case "�ۼ۽��"
        mCol�ۼ۽�� = intValue
    Case "���"
        mCol��� = intValue
    Case "���ۼ�"
        mcol���ۼ� = intValue
    Case "���۵�λ"
        mcol���۵�λ = intValue
    Case "���۽��"
        mcol���۽�� = intValue
    Case "���۲��"
        mcol���۲�� = intValue
    Case "�������"
        mCol������� = intValue
    Case "���ս���"
        mCol���ս��� = intValue
    Case "��Ʊ��"
        mCol��Ʊ�� = intValue
    Case "��Ʊ����"
        mcol��Ʊ���� = intValue
    Case "��Ʊ����"
        mCol��Ʊ���� = intValue
    Case "��Ʊ���"
        mCol��Ʊ��� = intValue
    Case "һ���Բ��� "
        mcolһ���Բ��� = intValue
    Case "�������"
        mcol������� = intValue
    Case "���Ч�� "
        mcol���Ч�� = intValue
    Case "�������"
        mcol������� = intValue
    Case "���ʧЧ��"
        mcol���ʧЧ�� = intValue
    Case "ע��֤��"
        mcolע��֤�� = intValue
    Case "ע��֤��Ч��"
        mcolע��֤��Ч�� = intValue
    Case "��Ʒ����"
        mcol��Ʒ���� = intValue
    Case "�ӳ���"
        mcol�ӳ��� = intValue
    Case Else
        blnShow = False
    End Select
'--    Debug.Print str���� & vbTab & intValue
    
    If Not blnShow Then
        mshBill.ColWidth(intValue) = 0
        mshBill.ColData(intValue) = 5
    Else
        mintLastCol = intValue
    End If
End Sub

Private Function CheckStock(ByVal lng����ID As Long, ByVal lng���� As Long, ByVal dbl���� As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '����:�����Ŀ��������Ƿ����
    '---------------------------------------------------------------------------------------------------------
    Dim lng�ⷿid As Long, intRow As Integer, intLop As Integer
    Dim blnMsg As Boolean
    Dim dblSum As Double, dbltotal As Double
    Dim varStuff As Variant
    
    Dim rsCheck As New ADODB.Recordset
    '�˻�ʱʹ�ñ����������Լ��������˻������Ƿ��㹻
    If mint�༭״̬ <> 8 And mbln�˻� = False Then CheckStock = True: Exit Function
    
    On Error GoTo ErrHandle
    
    '����ԭ�����е�ԭʼ����
    With mshBill
        dblSum = 0
        intRow = .Row
        If mint�༭״̬ <> 8 Then
            For Each varStuff In mCllBillData
                If varStuff(0) = Val(.TextMatrix(.Row, 0)) & "_" & Val(.TextMatrix(.Row, mCol����)) Then
                    dblSum = varStuff(1)
                    Exit For
                End If
            Next
        End If
        dbltotal = 0
        For intLop = 1 To .Rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                If intLop <> intRow And Trim(.TextMatrix(intLop, 0)) = Trim(.TextMatrix(intRow, 0)) And Val(.TextMatrix(intRow, mCol����)) = Val(.TextMatrix(intLop, mCol����)) Then
                    dbltotal = dbltotal + Val(.TextMatrix(intLop, mCol����)) * Val(.TextMatrix(intLop, mCol����ϵ��))
                End If
            End If
        Next
    End With
                
    
    
    lng�ⷿid = cboStock.ItemData(cboStock.ListIndex)
    
    gstrSQL = "Select �������� From ҩƷ��� Where �ⷿID=[1] And Nvl(����,0)=[2]  And ����=1 And ҩƷID=[3] "
    
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "������Ƿ��㹻�����˻�", lng�ⷿid, lng����, lng����ID)
    If rsCheck.RecordCount <> 0 Then
        dblSum = dblSum + Val(zlStr.Nvl(rsCheck!��������))
    Else
    End If
    dbltotal = dbltotal + dbl����
    If dbltotal > dblSum Then
        ShowMsgBox "�˻��������ܴ������еĿ����������ǰ�������Ϊ��" & dblSum / Val(mshBill.TextMatrix(mshBill.Row, mCol����ϵ��)) & "����"
        Exit Function
    End If
    CheckStock = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function �����_�˻�() As Boolean
    Dim lngRow As Long, lngRows As Long, lng����ID As Long, lng�ⷿid As Long, lng���� As Long
    Dim dbl������� As Double, dbl�˻����� As Double
    
    Dim rstemp As New ADODB.Recordset
    Dim blnExit As Boolean
    
    On Error GoTo ErrHandle

    '��������˻�����ʱ
    lngRows = mshBill.Rows - 1
    For lngRow = 1 To lngRows
        blnExit = False
        lng����ID = Val(mshBill.TextMatrix(lngRow, 0))
        
        If lng����ID <> 0 And Val(mshBill.TextMatrix(lngRow, mCol����)) < 0 Then
        
            lng���� = Val(mshBill.TextMatrix(lngRow, mCol����))
            
            lng�ⷿid = cboStock.ItemData(cboStock.ListIndex)
            
            gstrSQL = "" & _
                "   Select Nvl(A.ʵ������,0)/" & Choose(mintUnit + 1, "1", "B.����ϵ��") & " As ���� " & _
                "   From ҩƷ��� A,�������� B " & _
                "   Where A.ҩƷID=[1] And A.����=1 And A.ҩƷID=B.����ID And Nvl(A.����,0)=[2]  And A.�ⷿID=[3]"
            
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "����˻���¼�Ŀ���Ƿ��㹻!", lng����ID, lng����, lng�ⷿid)
            
            If rstemp.EOF Then
                blnExit = True
            Else
                blnExit = (rstemp!���� < Abs(Val(mshBill.TextMatrix(lngRow, mCol����))))
            End If
            
            If blnExit Then
                MsgBox "��" & lngRow & "�е����Ŀ������������������ˣ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Next
    �����_�˻� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ISCheckScalc�ۼ�(ByVal bln����� As Boolean, ByVal lngRow As Long) As Boolean
    '����:��鼰����ʱ�۵��ۼ�
    '����:bln�����:true-�����,False-�ɹ���
    '����:�ɹ�:����ture,���򷵻�False
    Dim dbl���� As Double
    Dim dbl�ۼ� As Double
    Dim dbl�ӳ��� As Double
    Dim dbl���� As Double, dbl���۽�� As Double, dbl������ As Double
    Dim sng�ֶ��ۼ� As Double
    Dim lng����ID As Long
    Dim dblָ������� As Double
    Dim dbl����ϵ�� As Double
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim rstemp As ADODB.Recordset
    
    err = 0: On Error GoTo ErrHand:
    With mshBill
        '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
        If Not (Trim(.TextMatrix(lngRow, mColԭ����)) <> "") Then
            'δ����
            GoTo Calc:
        End If
        If Val(Split(.TextMatrix(lngRow, mColԭ����), "||")(2)) <> 1 Then
            '����ʱ������,�˳�
            GoTo Calc:
        End If
        lng����ID = Val(.TextMatrix(lngRow, 0))
        If lng����ID = 0 Then
            Exit Function
        End If
        If mblnʱ�۹�ǰ���� Then
            If bln����� Then
                Call ���¼����ۼ�(lngRow)
                GoTo Calc:
            End If
        Else
            If bln����� = False Then
                Call ���¼����ۼ�(lngRow)
                GoTo Calc:
            End If
        End If
        
        dbl�ۼ� = Val(.TextMatrix(lngRow, mCol�ۼ�))
        '����ǰ���ǰ�۸�Ļ�,�������ӳ��ʲ�һ��
        If mblnʱ�۹�ǰ���� Then
            dbl���� = Val(.TextMatrix(lngRow, mCol�ɹ���))
        Else
            dbl���� = Val(.TextMatrix(lngRow, mCol�����))
        End If
        dbl����ϵ�� = Val(.TextMatrix(lngRow, mCol����ϵ��))
        
        'ԭЧ���ֶ����汣��ԭЧ�ڣ�ָ����ۣ��Ƿ��ۣ����÷����ȣ���ʽΪ�����Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
        dblָ������� = Val(Split(.TextMatrix(lngRow, mColԭ����), "||")(1))
            
        '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
        '���ϵͳ����Ϊ�棬����ʾ�û�����Ӽ���
        If mbln�Ӽ��� = True Then
            sngLeft = Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
            sngTop = Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
            If sngLeft + PicInput.Width > Screen.Width Then
                sngLeft = sngLeft + mshBill.MsfObj.CellWidth - PicInput.Width
            End If
            
            If sngTop + 1700 > Screen.Height Then
                sngTop = sngTop - mshBill.MsfObj.CellHeight - 1700
            End If
            
            With PicInput
                .Top = sngTop
                .Left = sngLeft
                .Visible = True
                .Tag = IIf(bln�����, "1", "0")
            End With
            Txt�Ӽ��� = Val(Replace(.TextMatrix(.Row, mcol�ӳ���), "%", "")) '"15.0000"
            .TextMatrix(lngRow, mCol�ۼ�) = Format(У�����ۼ�(dbl���� * (1 + (Val(Txt�Ӽ���) / 100)) + ʱ�۲������ۼ�(lng����ID, dbl����, (Val(Txt�Ӽ���) / 100))), mFMT.FM_���ۼ�)
            
'            If dbl�ۼ� <> 0 And dbl���� <> 0 Then
'                Txt�Ӽ��� = Format(����ӳ���(lng����ID, dbl�ۼ�, dbl����), "###0.0000000;-###0.0000000;0;0")
'            End If
            
            Txt�Ӽ���.Tag = Txt�Ӽ���
            Txt�Ӽ���.SetFocus
        ElseIf mbln�ֶμӳ��� = True Then
            dbl�ӳ��� = 0 ' Get�ֶμӳ���(dbl����) / 100
            If mint�༭״̬ = 8 And dbl�ۼ� <> 0 Then
            Else
                If Get�ֶμӳ��ۼ�(dbl����, dbl����ϵ��, mstrCaption, sng�ֶ��ۼ�) = True Then
                    
                    .TextMatrix(lngRow, mCol�ۼ�) = Format(У�����ۼ�(sng�ֶ��ۼ� + _
                                                    ʱ�۲������ۼ�(lng����ID, dbl����, 0, -1, sng�ֶ��ۼ�)) _
                                                    , mFMT.FM_���ۼ�)
                Else
                    ISCheckScalc�ۼ� = False
                    Exit Function
                End If
            End If
        Else 'mblnʱ������ȡ�ϴ��ۼ� = True���ȴ��ϴ�ȡ�����û�����ռӳ��ʷ�ʽȡ
            If mblnʱ������ȡ�ϴ��ۼ� = True Then
                gstrSQL = "Select Nvl(�ϴ��ۼ�, 0) As �ϴ��ۼ� From �������� Where ����id = [1]"
                Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID)
                If rstemp!�ϴ��ۼ� > 0 Then
                    .TextMatrix(lngRow, mCol�ۼ�) = Format(zlStr.Nvl(rstemp!�ϴ��ۼ�, 0) * Val(.TextMatrix(lngRow, mCol����ϵ��)), mFMT.FM_���ۼ�)
                    If dbl���� <> 0 Then
                        .TextMatrix(lngRow, mcol�ӳ���) = Format((Val(.TextMatrix(lngRow, mCol�ۼ�)) / dbl���� - 1) * 100, "###0.00") & "%"
                    End If
                Else
                    '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                    If dbl���� <> 0 Then
                        Txt�Ӽ��� = Val(Replace(.TextMatrix(.Row, mcol�ӳ���), "%", "")) '����ӳ���(lng����ID, dbl�ۼ�, dbl����)
                        .TextMatrix(lngRow, mcol�ӳ���) = Format(Txt�Ӽ���, "####0.00") & "%"
                    End If
                    If mint�༭״̬ = 8 And dbl�ۼ� <> 0 Then
                    Else
                        .TextMatrix(lngRow, mCol�ۼ�) = Format(У�����ۼ�(dbl���� * (1 + (Txt�Ӽ��� / 100)) + _
                        ʱ�۲������ۼ�(lng����ID, dbl����, (Txt�Ӽ��� / 100))), mFMT.FM_���ۼ�)
                    End If
                End If
            Else '3��ȡ�ۼ۷�ʽ��û������ʱ�����ռӳ��ʷ�ʽȡ
                If dbl���� <> 0 Then
                    Txt�Ӽ��� = Val(Replace(.TextMatrix(.Row, mcol�ӳ���), "%", "")) '����ӳ���(lng����ID, dbl�ۼ�, dbl����)
                    .TextMatrix(lngRow, mcol�ӳ���) = Format(Txt�Ӽ���, "####0.00") & "%"
                End If
                If mint�༭״̬ = 8 And dbl�ۼ� <> 0 Then
                Else
                    .TextMatrix(lngRow, mCol�ۼ�) = Format(У�����ۼ�(dbl���� * (1 + (Txt�Ӽ��� / 100)) + _
                    ʱ�۲������ۼ�(lng����ID, dbl����, (Txt�Ӽ��� / 100))), mFMT.FM_���ۼ�)
                End If
            End If
            
        End If
Calc:
        dbl���� = Val(.TextMatrix(lngRow, mCol����))
        dbl�ۼ� = Val(.TextMatrix(lngRow, mCol�ۼ�))
        dbl���� = Val(.TextMatrix(lngRow, mCol�����))
        dbl���۽�� = dbl���� * dbl�ۼ�
        dbl������ = dbl���� * dbl����
        .TextMatrix(lngRow, mCol�ۼ۽��) = Format(dbl���۽��, mFMT.FM_���)
        .TextMatrix(lngRow, mCol������) = Format(dbl������, mFMT.FM_���)
        .TextMatrix(lngRow, mCol��Ʊ���) = IIf(Trim(Trim(.TextMatrix(lngRow, mCol��Ʊ��))) = "", "", .TextMatrix(lngRow, mCol������))
        .TextMatrix(lngRow, mCol���) = Format(dbl���۽�� - dbl������, mFMT.FM_���)
        
        ''���˺�:���ۼ۴���
        Call �������ۼۼ����۲��(lngRow)
    End With
    ISCheckScalc�ۼ� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ���¼����ۼ�(ByVal lngRow As Long) As Boolean
    Dim dbl���� As Double, dbl���� As Double, dbl�ۼ� As Double, lng����ID As Long
    Dim dbl�ɹ��� As Double, dbl����� As Double, dbl�ӳ��� As Double
    Dim sng�ֶ��ۼ� As Double
    Dim rstemp As ADODB.Recordset
    
    With mshBill
            dbl���� = Val(.TextMatrix(lngRow, mCol����))
            dbl���� = Val(.TextMatrix(lngRow, mCol����))
            dbl�ۼ� = Val(.TextMatrix(lngRow, mCol�ۼ�))
            dbl����� = Val(.TextMatrix(lngRow, mCol�����))
            dbl�ɹ��� = Val(.TextMatrix(lngRow, mCol�ɹ���))
            lng����ID = Val(.TextMatrix(lngRow, 0))
                
            If mbln�Ӽ��� Then
                mdbl�Ӽ��� = 15
                If dbl�ۼ� <> 0 And dbl����� <> 0 Then
                    If Val(Replace(.TextMatrix(lngRow, mcol�ӳ���), "%", "")) >= 0 Then '"X.XXX%"С����Ϊ0�ͻᱨ���Ͳ�ƥ�����
                        mdbl�Ӽ��� = Val(Replace(.TextMatrix(lngRow, mcol�ӳ���), "%", "")) 'Val(Split(.TextMatrix(lngRow, mcol�ӳ���), "%")(0))
                    Else
                        mdbl�Ӽ��� = ����ӳ���(lng����ID, dbl�ۼ�, IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����))
                    End If
                End If
             End If
                
            '��ʱ�۲��ϵĴ���
            If .TextMatrix(lngRow, mColԭ����) <> "" Then
                '���¼������ۼۡ����
                '
                '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                If Split(.TextMatrix(lngRow, mColԭ����), "||")(2) = 1 Then
                    '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                    If mbln�Ӽ��� Then
                        If mint�༭״̬ = 8 And dbl�ۼ� <> 0 Then
                        Else
                            If mbln�ֶμӳ��� Then
                                If Get�ֶμӳ��ۼ�(IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����), Val(.TextMatrix(.Row, mCol����ϵ��)), mstrCaption, sng�ֶ��ۼ�) = True Then
                                    .TextMatrix(lngRow, mCol�ۼ�) = Format(У�����ۼ�(sng�ֶ��ۼ� + _
                                                                    ʱ�۲������ۼ�(lng����ID, IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����), 0, -1, sng�ֶ��ۼ�)) _
                                                                    , mFMT.FM_���ۼ�)
                                End If
                            Else
                                .TextMatrix(lngRow, mCol�ۼ�) = Format(У�����ۼ�(IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����) * (1 + (mdbl�Ӽ��� / 100)) + _
                                                                ʱ�۲������ۼ�(lng����ID, IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����), (mdbl�Ӽ��� / 100))) _
                                                                , mFMT.FM_���ۼ�)
                            End If
                        End If
                        .TextMatrix(lngRow, mCol�ۼ۽��) = Format(Val(.TextMatrix(lngRow, mCol�ۼ�)) * dbl����, mFMT.FM_���)
                        .TextMatrix(lngRow, mCol���) = Format(IIf(.TextMatrix(lngRow, mCol�ۼ۽��) = "", 0, .TextMatrix(lngRow, mCol�ۼ۽��)) - IIf(.TextMatrix(lngRow, mCol������) = "", 0, .TextMatrix(lngRow, mCol������)), mFMT.FM_���)
                    Else
                        If mbln�ֶμӳ��� Then
                            dbl�ӳ��� = 0 ' Get�ֶμӳ���(IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����)) / 100
                        Else
                            dbl�ӳ��� = Val(Replace(.TextMatrix(lngRow, mcol�ӳ���), "%", "")) / 100
                        End If
                        If mint�༭״̬ = 8 And dbl�ۼ� <> 0 Then
                        Else
                            If mbln�ֶμӳ��� Then
                                If Get�ֶμӳ��ۼ�(IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����), Val(.TextMatrix(.Row, mCol����ϵ��)), mstrCaption, sng�ֶ��ۼ�) = True Then
                                    .TextMatrix(lngRow, mCol�ۼ�) = Format(У�����ۼ�(sng�ֶ��ۼ� + _
                                                                    ʱ�۲������ۼ�(lng����ID, IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����), 0, -1, sng�ֶ��ۼ�)) _
                                                                    , mFMT.FM_���ۼ�)
                                End If
                            Else
                                If mblnʱ������ȡ�ϴ��ۼ� = True Then 'ȡ�ϴ��ۼ�
                                    gstrSQL = "Select Nvl(�ϴ��ۼ�, 0) As �ϴ��ۼ� From �������� Where ����id = [1]"
                                    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID)
                                    If rstemp!�ϴ��ۼ� > 0 Then
                                        .TextMatrix(lngRow, mCol�ۼ�) = Format(zlStr.Nvl(rstemp!�ϴ��ۼ�, 0) * Val(.TextMatrix(lngRow, mCol����ϵ��)), mFMT.FM_���ۼ�)
                                        If IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����) <> 0 Then
                                            .TextMatrix(lngRow, mcol�ӳ���) = Format((Val(.TextMatrix(lngRow, mCol�ۼ�)) / IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����) - 1) * 100, "###0.00") & "%"
                                        End If
                                    Else
                                        dbl�ӳ��� = Val(Replace(.TextMatrix(lngRow, mcol�ӳ���), "%", "")) / 100
                                        
                                        .TextMatrix(lngRow, mCol�ۼ�) = Format(У�����ۼ�(IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����) * (1 + dbl�ӳ���) + _
                                                                        ʱ�۲������ۼ�(lng����ID, IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����), dbl�ӳ���)) _
                                                                        , mFMT.FM_���ۼ�)
                                    End If
                                Else '������Ϊѡ�񰴼ӳ��ʼ���
                                    dbl�ӳ��� = Val(Replace(.TextMatrix(lngRow, mcol�ӳ���), "%", "")) / 100
                                        
                                    .TextMatrix(lngRow, mCol�ۼ�) = Format(У�����ۼ�(IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����) * (1 + dbl�ӳ���) + _
                                                                    ʱ�۲������ۼ�(lng����ID, IIf(mblnʱ�۹�ǰ����, dbl�ɹ���, dbl�����), dbl�ӳ���)) _
                                                                    , mFMT.FM_���ۼ�)
                                End If
                            End If
                        End If
                        .TextMatrix(lngRow, mCol�ۼ۽��) = Format(dbl���� * Val(.TextMatrix(lngRow, mCol�ۼ�)), mFMT.FM_���)
                    End If
                End If
            End If
            ''���˺�:���ۼ۴���
            Call �������ۼۼ����۲��(lngRow)
    End With
End Function


Private Sub RefreshBill()
    '�����¼۸����µ���������ݣ����ڵ������ʱ
    Dim lngRow As Long, lngRows As Long, lng����ID As Long
    Dim dbl���� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    Dim rsprice As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = " Select �շ�ϸĿID,nvl(�ּ�,0) �ּ� From �շѼ�Ŀ " & _
            " Where (��ֹ���� Is NULL Or sysdate Between ִ������ And nvl(��ֹ����,to_date('3000-01-01','yyyy-MM-dd')))" & _
            GetPriceClassString("")
            
    gstrSQL = "Select A.���,A.ҩƷID ,B.�ּ� From ҩƷ�շ���¼ A,(" & gstrSQL & ") B,�շ���ĿĿ¼ C" & _
            " Where A.����=15  And A.NO=[1] And A.ҩƷID=B.�շ�ϸĿID And C.ID=B.�շ�ϸĿID And Round(A.���ۼ�,7)<>Round(B.�ּ�,7) And Nvl(C.�Ƿ���,0)=0" & _
            " Order by A.���"
    
    Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[ȡ��ǰ�۸�]", txtNO.Text)
    
    If rsprice.EOF Then Exit Sub
    
    lngRows = mshBill.Rows - 1
    For lngRow = 1 To lngRows
        lng����ID = Val(mshBill.TextMatrix(lngRow, 0))
        If lng����ID <> 0 Then
            rsprice.Filter = "ҩƷID=" & lng����ID
            If rsprice.RecordCount <> 0 Then
                '�Ե�ǰ���¼۸����µ���������ݣ����ۡ����۽���ۣ�
                dbl���ۼ� = rsprice!�ּ� * Val(mshBill.TextMatrix(lngRow, mCol����ϵ��))
                dbl�ɱ��� = Val(mshBill.TextMatrix(lngRow, mCol�����))
                dbl���� = Val(mshBill.TextMatrix(lngRow, mCol����))
                
                dbl�ɱ���� = dbl�ɱ��� * dbl����
                dbl���۽�� = dbl���ۼ� * dbl����
                dbl��� = dbl���۽�� - dbl�ɱ����
                
                mshBill.TextMatrix(lngRow, mCol�ۼ�) = Format(dbl���ۼ�, mFMT.FM_���ۼ�)
                mshBill.TextMatrix(lngRow, mCol�ۼ۽��) = Format(dbl���۽��, mFMT.FM_���)
                mshBill.TextMatrix(lngRow, mCol���) = Format(dbl���, mFMT.FM_���)
                ''���˺�:���ۼ۴���
                Call �������ۼۼ����۲��(lngRow)
            End If
        End If
    Next
    rsprice.Filter = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Function CheckProvider() As Boolean
    Dim lngRow As Long
    Dim str���� As String
    Dim str�б���� As String
    Dim rstemp As New ADODB.Recordset
    '��鹩Ӧ���Ƿ����б���ϵ��б굥λ
    On Error GoTo ErrHandle
    str���� = ""
    With mshBill
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, 0)) <> 0 Then
                str���� = str���� & "," & Val(.TextMatrix(lngRow, 0))
            End If
        Next
        If str���� <> "" Then str���� = Mid(str����, 2)
    End With
    
    '���б��������ȥ����ͬһ���б굥λ���б������������޼�¼����˵����ȷ�����򰴼�¼�еĲ���ID��ʾ���ǺϷ����б굥λ
    gstrSQL = " Select a.����ID From �������� a,Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) b " & _
              " Where a.����ID=b.Column_Value And Nvl(a.�б����,0)=1" & _
              " Minus" & _
              " Select A.����ID From " & _
              "     (Select a.����ID From �������� a,Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) b  Where a.����ID=b.Column_Value And Nvl(a.�б����,0)=1) A,�����б굥λ B" & _
              " Where A.����ID=B.����ID And B.��λID=[1]"
    gstrSQL = " Select '['||A.����||']'||A.���� �������� " & _
              " From " & _
              "     (Select A.����ID,C.����,Nvl(B.����,C.����) ����" & _
              "     From (" & gstrSQL & ") A,�շ���Ŀ���� B,�շ���ĿĿ¼ C" & _
              "     Where A.����ID=B.�շ�ϸĿID(+) and A.����ID=C.ID" & _
              "     and B.����(+)=3 and B.����(+)=1) A"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[�ж��Ƿ����б굥λ���ɹ�]", Val(txtProvider.Tag), str����)
    
    With rstemp
        str���� = ""
        Do While Not .EOF
            str���� = str���� & "��" & rstemp!��������
            .MoveNext
        Loop
        If str���� <> "" Then str���� = Mid(str����, 2)
    End With
    
    If str���� <> "" Then
        If mbln���б굥λ��� Then
            If MsgBox("�ù�����λ���������б����ĵ��б굥λ,�Ƿ������" & vbCrLf & str����, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            MsgBox "�ù�����λ���������б����ĵ��б굥λ��" & vbCrLf & str����, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CheckProvider = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub chkת���ƿ�_Click()
    Dim blnEnabled As Boolean
    blnEnabled = (chkת���ƿ�.Value = 1)
    
    cboType.Enabled = blnEnabled
    cboEnterStock.Enabled = blnEnabled
    txtDraw.Enabled = blnEnabled
    cmdDraw.Enabled = blnEnabled
    txtDrawPerson.Enabled = blnEnabled
    cmdDrawPerson.Enabled = blnEnabled
    lbl������.Enabled = blnEnabled
End Sub

Private Function Check�ƿ�(ByRef blnExit As Boolean) As Boolean
    '--------------------------------------------------------------------------------------------------
    '����:����⹺��ⵥת�Ƶ������ⷿʱ���Բ��ϵ��ƿ��������м��:
    '       1)�Դ洢�ⷿ���м��
    '       2)�Ը������ý��м��
    '����:blnExit-�����˳�
    '--------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim rstemp As New ADODB.Recordset
    Dim strTmp As String
    Dim bln���ƿ���� As Boolean
    Dim bln�߱����ٲ��� As Boolean
    
    On Error GoTo ErrHand
    bln���ƿ���� = True
    '����Ƿ��˻���
    If mbln�˻� Then
        If MsgBox("�˿ⵥ������ʹ�õ����ƿ�Ĺ��ܡ�ȷ����⣬ѡ��<��>��������ˣ�ѡ��<��>��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            blnExit = True
        Else
            blnExit = False
        End If
        Check�ƿ� = False
        Exit Function
    End If
    If chkת���ƿ�.Value <> 1 Then Check�ƿ� = False: Exit Function
    If cboEnterStock.ListIndex < 0 Then Check�ƿ� = False: Exit Function
    
    bln�߱����ٲ��� = �ж�ֻ�߱����ϲ���(cboEnterStock.ItemData(cboEnterStock.ListIndex))
    
    strTmp = ""
    With mshBill
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, 0)) <> "" Then
                '��鸺�����ⲿ��,���ܽ��е���
                If Val(.TextMatrix(i, mCol����)) < 0 Then
                    If MsgBox("����Ϊ��" & .TextMatrix(i, mCol����) & "������������Ϊ����������ʹ�õ����ƿ�Ĺ��ܡ�ȷ����⣬ѡ��<��>��������ˣ�ѡ��<��>��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        blnExit = True
                    Else
                        blnExit = False
                    End If
                    Check�ƿ� = False
                    Exit Function
                End If
                
                '����Ƿ����ô洢�ⷿ
                gstrSQL = "select �շ�ϸĿID from �շ�ִ�п��� where �շ�ϸĿID=[1] and ִ�п���ID=[2]  "
                Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[�жϴ洢�ⷿ]", Val(.TextMatrix(i, 0)), cboEnterStock.ItemData(cboEnterStock.ListIndex))
                If rstemp.RecordCount = 0 Then
                     strTmp = strTmp & "����:" & mshBill.TextMatrix(i, mCol����) & " ���:" & mshBill.TextMatrix(i, mCol���) & vbCrLf
                Else
                    If bln�߱����ٲ��� Then
                        '�жϸ��ٲ���
                        gstrSQL = "Select �������� From �������� where ����id=[1]"
                        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[�жϸ�������]", Val(.TextMatrix(i, 0)))
                        If Not rstemp.EOF Then
                            If Val(Nvl(rstemp!��������)) = 1 Then
                                bln���ƿ���� = False
                            Else
                                strTmp = strTmp & "����:" & mshBill.TextMatrix(i, mCol����) & " ���:" & mshBill.TextMatrix(i, mCol���) & vbCrLf
                            End If
                        End If
                    Else
                        bln���ƿ���� = False
                    End If
                End If
            End If
        Next
    End With
    
    If strTmp <> "" Then
        If bln���ƿ���� Then
            ShowMsgBox "����������û�����ô洢�ⷿ����ٲ��ϣ��������ƿ⵽[" & cboEnterStock.Text & "]��"
            Check�ƿ� = False
            Exit Function
        End If
        ShowMsgBox "���²���û�����ô洢�ⷿ���������ƿ⵽[" & cboEnterStock.Text & "] ��" & vbCrLf & strTmp & vbCrLf & "������Ͽ��Ե����ƿ⡣"
    End If
    Check�ƿ� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function RestoreBILLWidthSet() As Boolean
    Dim strWidth As String, strText As String
    Dim arrText As Variant, i As Integer
    Dim arrWidth As Variant
    
    On Error Resume Next
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����", 0, 0)) = 0 Then
        RestoreBILLWidthSet = True: Exit Function
    End If
   
    '����Ƿ���Ҫ�ָ�
    
    strWidth = zlDatabase.GetPara("�����п�", glngSys, mlngModule)
    strText = zlDatabase.GetPara("������ͷ�ı�", glngSys, mlngModule)
    
    If strText <> "" Then
        '�̶��б���,���ָ���ʹ��ȱʡ
        '��������,���ָ���ʹ��ȱʡ
        arrText = Split(strText, ",")
        arrWidth = Split(strWidth, ",")
        For i = 0 To UBound(arrText) + 1
            Call SetBillColWidth(arrText(i), arrWidth(i))
        Next
    End If
    RestoreBILLWidthSet = True
End Function
Private Sub SetBillColWidth(ByVal strColName As String, ByVal lngWidth As Long)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����п��
    '����:strName-��ͷ������
    '     lngwidth-�п�
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Integer, i As Integer
    If lngWidth <= 0 Then Exit Sub
    With mshBill
        intCol = -1
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = strColName Then
                intCol = i
                Exit For
            End If
        Next
        If intCol = -1 Then Exit Sub
        If .ColWidth(intCol) <= 0 Then Exit Sub
        .ColWidth(intCol) = lngWidth
    End With
End Sub
Private Sub SaveBILLWidth()
    Dim strWidth As String, strText As String, i As Integer
        
    On Error Resume Next
    
    strWidth = "": strText = ""
    For i = 0 To mshBill.Cols - 1
        strWidth = strWidth & "," & mshBill.ColWidth(i)
        strText = strText & "," & mshBill.TextMatrix(0, i)
    Next
    zlDatabase.SetPara "�����п�", Mid(strWidth, 2), glngSys, mlngModule
    zlDatabase.SetPara "������ͷ�ı�", Mid(strText, 2), glngSys, mlngModule
End Sub
Private Function Set��������Update(Optional blnRowInput As Boolean = False) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '����:��������,������Ӧ���޸���Ŀ
    '����:blnRowInput-�Ƿ�����������ж�̬�ж��������
    '����:���óɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2007/05/30
    '----------------------------------------------------------------------------------------------------------
    Dim int���� As Integer, intCol As Integer
    Dim mrs���ڿ��� As New ADODB.Recordset
    Dim arr���� As Variant
    Dim str���� As String
    
    On Error GoTo ErrHandle
    '1.������2���޸ģ�3�����գ�4���鿴��5���޸ķ�Ʊ��6��������
    '7��������ˣ������������µ��ݲ���ˣ��Ѹ���ĵ��ݲ����������ˣ�ͬ����������˺�ĵ��ݲ����������;
    '8�����Ŀ��˻�,9-�˲�
    
    If mint�༭״̬ = 3 Or mint�༭״̬ = 9 Or mint�༭״̬ = 7 Then
        '  1.�˲飬2.��ˣ�3.�������
        int���� = Decode(mint�༭״̬, 3, 2, 9, 1, 3)
    Else
        Set��������Update = True
        Exit Function
    End If
    
    gstrSQL = "Select ����,','||����||',' as ���� From ���ݻ��ڿ��� where ����=[1] and ����=[2] order by ����"
    '����:��Ÿû��ڿ��޸ĵ���Ŀ����ʽΪ"��Ŀ1,��Ŀ2,..."����ѡ��ĿΪ"�ɹ��ۣ����ʣ�����ۣ�������ۼۣ���Ʊ�ţ���Ʊ���ڣ���Ʊ���,��Ʊ����"���Ժ�����䡣
    
    If mrs���ڿ��� Is Nothing Then
        Set mrs���ڿ��� = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, 15, int����)
    ElseIf mrs���ڿ���.State <> 1 Then
        Set mrs���ڿ��� = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, 15, int����)
    End If
    '���û�����ݿ����򶼲����޸�
    If mrs���ڿ���.RecordCount = 0 Then
        For intCol = 0 To mshBill.Cols - 1
            mshBill.ColData(intCol) = 0
        Next
        Exit Function
    End If
    With mshBill
        If blnRowInput Then
            str���� = zlStr.Nvl(mrs���ڿ���!����)
            If InStr(1, str����, ",�ۼ�,") > 0 Then
                '�����ʱ�����ģ������������ۼ�
                '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                If .TextMatrix(.Row, mColԭ����) <> "" Then
                    If Split(.TextMatrix(.Row, mColԭ����), "||")(2) = 1 Then
                        .ColData(mCol�ۼ�) = IIf(mblnʱ������ֱ��ȷ���ۼ�, 4, 5)
                        .ColData(mcol���ۼ�) = IIf(mbln�˻� Or mint�༭״̬ = 8, 5, 4)
                    Else
                        .ColData(mCol�ۼ�) = 0
                        .ColData(mcol���ۼ�) = 0
                    End If
                Else
                     .ColData(mCol�ۼ�) = 0
                     .ColData(mcol���ۼ�) = 0
                End If
            Else
                .ColData(mCol�ۼ�) = 0
                .ColData(mcol���ۼ�) = 0
            End If
            If Trim(.TextMatrix(.Row, mCol��Ʊ��)) = "" Then
                .ColData(mCol��Ʊ����) = 0
                .ColData(mCol��Ʊ���) = 0
                .ColData(mcol��Ʊ����) = 0
            Else
                If InStr(1, str����, ",��Ʊ����,") > 0 Then
                    .ColData(mCol��Ʊ����) = 2
                Else
                    .ColData(mCol��Ʊ����) = 0
                End If
                If InStr(1, str����, ",��Ʊ����,") > 0 Then
                    .ColData(mcol��Ʊ����) = 4
                Else
                    .ColData(mcol��Ʊ����) = 0
                End If
                If InStr(1, str����, ",��Ʊ���,") > 0 Then
                    .ColData(mCol��Ʊ���) = 4
                Else
                    .ColData(mCol��Ʊ���) = 0
                End If
            End If
            Set��������Update = True
            Exit Function
        End If
        
        For intCol = 0 To .Cols - 1
            .ColData(intCol) = 0
        Next
        If mrs���ڿ���.EOF = False Then
            '"�ɹ��ۣ����ʣ�����ۣ�������ۼۣ���Ʊ�ţ���Ʊ���ڣ���Ʊ���,��Ʊ����
            arr���� = Split(zlStr.Nvl(mrs���ڿ���!����), ",")
            For intCol = 0 To UBound(arr����)
                Select Case arr����(intCol)
                Case "�ɹ���"
                    If mbln��ǿ�ƿ���ָ���۸� = False Then
                        .ColData(mColָ��������) = IIf(mbln�޸�������, 4, 0)
                    End If
                    .ColData(mCol�ɹ���) = 4
                Case "����"
                    .ColData(mCol����) = 4
                Case "�����"
                    .ColData(mCol�����) = 4
                Case "������"
                    .ColData(mCol������) = 4
                Case "�ۼ�"
                    .ColData(mCol�ۼ�) = 4
                Case "���ۼ�"
                    .ColData(mcol���ۼ�) = IIf(mbln�˻� Or mint�༭״̬ = 8, 5, 4)
                Case "��Ʊ��"
                    mshBill.ColData(mCol��Ʊ��) = 4
                Case "��Ʊ����"
                    mshBill.ColData(mcol��Ʊ����) = 4
                Case "��Ʊ����"
                    mshBill.ColData(mCol��Ʊ����) = 2
                Case "��Ʊ���"
                   .ColData(mCol��Ʊ���) = 4
                End Select
            Next
        End If
        '���¶�λ
        For int���� = 0 To mshBill.Cols - 1
            If .ColData(int����) = 4 Or .ColData(int����) = 2 Then
                .LocateCol = int����
                Exit For
            End If
        Next
    End With
    Set��������Update = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SelectItem(ByVal objCtl As Control, ByVal strKey As String, Optional bln��Ա As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�๦��ѡ����
    '���:objCtl-�ı���ؼ�
    '     strKey-Ҫ������ֵ
    '     bln��Ա-�Ƿ���Աѡ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-27 10:37:40
    '-----------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, strTittle As String
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rstemp  As ADODB.Recordset
    Dim bytStyle As Byte
    Dim strվ������ As String
    
    'zlDatabase.ShowSelect
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
    
    strվ������ = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
    
    bytStyle = 0
    If bln��Ա Then
        strTittle = "��Աѡ����"
        If strKey = "" Then
            gstrSQL = "" & _
                    "   Select ID, ���,����,���� From ��Ա�� a " & _
                    "   Where exists(select 1 from ������Ա where ��Աid=a.id and ����id=[2]) " & _
                    "       and (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
                    "       and (a.վ��=[3] or a.վ�� is null) " & _
                    "   order by ���"
        Else
            gstrSQL = "" & _
                    "   Select ID, ���,����,���� From ��Ա�� a " & _
                    "   Where (���� like [1] or  ���  like [1] or  ����  like  upper([1])) " & _
                    "       and (a.վ��=[3] or a.վ�� is null) " & _
                    "       and exists(select 1 from ������Ա where ��Աid=a.id and ����id=[2]) " & _
                    "       and (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & _
                    "   order by ���"
        End If
    Else
        strTittle = "����ѡ����"
        gstrSQL = "SELECT a.id,a.�ϼ�ID,a.����,a.����,a.���� " & _
                  "FROM ���ű� a " & _
                  "Where ( TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or a.����ʱ�� is null ) " & _
                  IIf(strվ������ <> "", " And a.վ�� = [3] ", "")
    
        If strKey <> "" Then
            gstrSQL = gstrSQL & _
                      " And (a.���� like upper([1]) Or a.���� like upper([1]) or a.���� like [1])" & _
                      " Order by ���� "
        Else
            If gstrNodeNo = "-" Then
                'û��վ���,��������ʾ
                gstrSQL = gstrSQL & " start with �ϼ�id is null connect by prior id=�ϼ�id "
                bytStyle = 1
            Else
                '����վ�㣬��Ҫ�ǿ����ϼ�������վ���ţ����¼�δ���õ���������ֻ�����б�ʽ���д���
                gstrSQL = gstrSQL & " Order by ���� "
            End If
        End If
    End If
    
    strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hwnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rstemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, strTittle, False, "", "", _
                    False, False, True, sngX, sngY, lngH, blnCancel, False, False, _
                    strKey, Val(txtDraw.Tag), strվ������)
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If rstemp Is Nothing Then
        ShowMsgBox "û���ҵ���������������,����!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        With objCtl
            If bln��Ա Then
                .TextMatrix(.Row, .Col) = zlStr.Nvl(rstemp!����)
                .Cell(flexcpData, .Row, .Col) = zlStr.Nvl(rstemp!����)
            Else
                .TextMatrix(.Row, .Col) = zlStr.Nvl(rstemp!����) & "-" & zlStr.Nvl(rstemp!����)
                .Cell(flexcpData, .Row, .Col) = zlStr.Nvl(rstemp!Id)
            End If
        End With
    Else
        If bln��Ա Then
            objCtl.Text = zlStr.Nvl(rstemp!����)
            objCtl.Tag = zlStr.Nvl(rstemp!����)
        Else
            objCtl.Text = zlStr.Nvl(rstemp!����) & "-" & zlStr.Nvl(rstemp!����)
            objCtl.Tag = IIf(bln��Ա, zlStr.Nvl(rstemp!����), zlStr.Nvl(rstemp!Id))
        End If
        zlControl.ControlSetFocus objCtl, True
        OS.PressKey vbKeyTab
    End If
    SelectItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdDraw_Click()
    If SelectItem(txtDraw, "") = False Then Exit Sub
End Sub
Private Sub txtDraw_Change()
    txtDraw.Tag = ""
    txtDrawPerson.Text = ""
    mblnChange = True
End Sub

Private Sub txtDraw_GotFocus()
    OS.OpenIme False
    zlControl.TxtSelAll txtDraw
End Sub
Private Sub txtDraw_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtDraw.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
    If SelectItem(txtDraw, Trim(txtDraw.Text)) = False Then Exit Sub
End Sub

Private Function ��ǰ��Ϊ�ⷿ() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�жϵ�ǰ�ⷿ��Ϊ�ⷿ
    '���:
    '����:
    '����:����true��ʾ��Ϊ�ⷿ,����Ϊ(���ϲ��Ż��Ƽ���)
    '����:���˺�
    '����:2008-12-03 11:23:18
    '-----------------------------------------------------------------------------------------------------------
    Dim rstemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   SELECT count(*)" & _
        "   From ��������˵�� " & _
        "   WHERE ((�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���')) " & _
        "        AND ����id =[1]"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
    If rstemp.Fields(0) > 0 Then
        ��ǰ��Ϊ�ⷿ = False
        mbln�ⷿ = False
    Else
        ��ǰ��Ϊ�ⷿ = True
        mbln�ⷿ = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfCostlyInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfCostlyInfo
        Select Case .Col
            Case .ColIndex("����")
                .ColComboList(.Col) = "..."
            Case .ColIndex("��������")
                .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

Private Sub vsfCostlyInfo_BeforeDataRefresh(Cancel As Boolean)
    '���浽���ݼ�
    If vsfCostlyInfo.Visible = False Then Exit Sub
    If vsfCostlyInfo.Rows < 2 Then Exit Sub
    If vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = "" Then
        Cancel = True
        MsgBox "'����'��Ϣδ¼�룡", vbCritical, gstrSysName
        Exit Sub
    End If
    If vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("��������")) = "" Then
        Cancel = True
        MsgBox "'��������'��Ϣδ¼�룡", vbCritical, gstrSysName
        Exit Sub
    End If
    If vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("סԺ��")) = "" Then
        Cancel = True
        MsgBox "'סԺ��'��Ϣδ¼�룡", vbCritical, gstrSysName
        Exit Sub
    End If
    If vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = "" Then
        Cancel = True
        MsgBox "'����'��Ϣδ¼�룡", vbCritical, gstrSysName
        Exit Sub
    End If
End Sub

Private Sub vsfCostlyInfo_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow > 1 Then Cancel = True
End Sub

Private Sub CostlyInfo_Refresh(ByVal lngId As Long, ByVal blnCostly As Boolean)
'��ֵ������Ϣˢ��
    If lngId < 1 Then Exit Sub
    If mshBill.TextMatrix(lngId, 0) = "" Then Exit Sub
    If mrsCostlyInfo Is Nothing Then Exit Sub
    If mrsCostlyInfo.RecordCount <= 0 Then
        If blnCostly Then
            mrsCostlyInfo.AddNew
            mrsCostlyInfo!sn = lngId
            mrsCostlyInfo.Update
        Else
            Exit Sub
        End If
    End If
    If mrsCostlyInfo.RecordCount > 0 Then mrsCostlyInfo.MoveFirst
    mrsCostlyInfo.Find "SN=" & lngId
    If mrsCostlyInfo.EOF Then
        If blnCostly Then
            mrsCostlyInfo.AddNew
            mrsCostlyInfo!sn = lngId
            mrsCostlyInfo.Update
        End If
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����ID")) = ""
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = ""
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("��������")) = ""
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("סԺ��")) = ""
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = ""
    Else
        If blnCostly = False Then
            mrsCostlyInfo.Delete
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����ID")) = ""
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = ""
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("��������")) = ""
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("סԺ��")) = ""
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = ""
        Else
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����ID")) = IIf(IsNull(mrsCostlyInfo!Id), "", mrsCostlyInfo!Id)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = IIf(IsNull(mrsCostlyInfo!����), "", mrsCostlyInfo!����)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("��������")) = IIf(IsNull(mrsCostlyInfo!��������), "", mrsCostlyInfo!��������)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("סԺ��")) = IIf(IsNull(mrsCostlyInfo!סԺ��), "", mrsCostlyInfo!סԺ��)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = IIf(IsNull(mrsCostlyInfo!����), "", mrsCostlyInfo!����)
        End If
    End If
    
    vsfCostlyInfo.Visible = blnCostly
    Call Form_Resize
End Sub

Private Function IsCostly(ByVal lngMaterialID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    Set rsTmp = zlDatabase.OpenSQLRecord("select count(��ֵ����) ��ֵ���� from �������� where ����id=[1] and nvl(��ֵ����,'')='1'", mstrCaption, lngMaterialID)
    If rsTmp.RecordCount = 1 And rsTmp!��ֵ���� = "1" Then
        IsCostly = True
        mbln��ֵ���� = True
    Else
        mbln��ֵ���� = False
    End If
    rsTmp.Close
    
    cmdCopy.Enabled = mbln��ֵ����
    txtCopy.Enabled = mbln��ֵ����
    lblCopy.Enabled = mbln��ֵ����
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfCostlyInfo_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vsfCostlyInfo
        Select Case Col
            Case .ColIndex("����")
                Call Comm_Selecter("%" & .EditText & "%", 1)
            Case .ColIndex("��������")
                'If .EditText = "" Then
                    Call Comm_Selecter(.TextMatrix(1, .ColIndex("����ID")), 2)
                'End If
        End Select
    End With
End Sub

Private Sub vsfCostlyInfo_EnterCell()
    With vsfCostlyInfo
        Select Case .Col
            Case .ColIndex("����"), .ColIndex("��������")
                .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

Private Sub vsfCostlyInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfCostlyInfo
        Select Case .Col
            Case .ColIndex("����")
                .ColComboList(.Col) = ""
            Case .ColIndex("��������")
                .ColComboList(.Col) = ""
        End Select
    End With
End Sub

Private Sub vsfCostlyInfo_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vsfCostlyInfo
            Select Case Col
                Case .ColIndex("����")
                    Call Comm_Selecter("%" & .EditText & "%", 1)
                Case .ColIndex("��������")
                    'If .EditText = "" Then
                        Call Comm_Selecter(.TextMatrix(1, .ColIndex("����id")), 2)
                    'End If
            End Select
            If Col = .ColIndex("����") Then
                .Col = 1
            Else
                .Col = Col + 1
            End If
        End With
        Exit Sub
    
    ElseIf vsfCostlyInfo.ColIndex("סԺ��") = Col Then
        If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
    Else
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub vsfCostlyInfo_Validate(Cancel As Boolean)
    '��������
    With vsfCostlyInfo
        '��λ
        If mrsCostlyInfo.RecordCount > 0 Then
            mrsCostlyInfo.MoveFirst
            mrsCostlyInfo.Find "SN=" & mshBill.TextMatrix(mshBill.Row, 1)
        End If
        If mrsCostlyInfo.EOF Then
            '����
            mrsCostlyInfo.AddNew
            mrsCostlyInfo!sn = mshBill.TextMatrix(mshBill.Row, 1)
        End If
        
        mrsCostlyInfo!Id = IIf(.TextMatrix(1, 5) = "", 0, .TextMatrix(1, 5))
        mrsCostlyInfo!���� = .TextMatrix(1, 1)
        mrsCostlyInfo!�������� = .TextMatrix(1, 2)
        mrsCostlyInfo!סԺ�� = IIf(.TextMatrix(1, 3) = "", Null, .TextMatrix(1, 3))
        mrsCostlyInfo!���� = .TextMatrix(1, 4)
        mrsCostlyInfo.Update
    End With
End Sub

Private Sub RecountSN(ByVal lngRow As Long)
'������Ӧ��ֵ���ϵ�SN
    Dim i As Long, lngMax As Long
    If mrsCostlyInfo.RecordCount <= 0 Then Exit Sub
    mrsCostlyInfo.MoveFirst
    Do While Not mrsCostlyInfo.EOF
        If lngMax < mrsCostlyInfo!sn Then
            lngMax = mrsCostlyInfo!sn
        End If
        mrsCostlyInfo.MoveNext
    Loop
    
    If lngRow >= lngMax Then Exit Sub
    
    For i = lngRow + 1 To lngMax
        With mrsCostlyInfo
            .MoveFirst
            .Find "SN=" & i
            If Not .EOF Then
                mrsCostlyInfo!sn = mrsCostlyInfo!sn - 1
                mrsCostlyInfo.Update
            End If
        End With
    Next
End Sub

Private Sub Comm_Selecter(ByVal strParam As String, ByVal bytIndex As Byte)
    Dim rstemp As ADODB.Recordset
    Dim sngX As Single, sngY As Single, lngH As Long
    Dim blnCancel As Boolean
    Dim strsql As String
    Dim rectTmp As RECT
    
    If bytIndex < 3 Then
        Call CalcPosition(sngX, sngY, vsfCostlyInfo)
        lngH = vsfCostlyInfo.CellHeight
    Else
        rectTmp = zlControl.GetControlRect(Me.txtTypeVar.hwnd)
        sngX = rectTmp.Left
        sngY = rectTmp.Top + Me.txtTypeVar.Height
        lngH = Me.txtTypeVar.Height
    End If
    sngY = sngY - lngH
    
    On Error GoTo ErrHandle
    Select Case bytIndex
    Case 1
        strsql = "SELECT a.id, a.����, a.����, a.���� " _
                & "FROM ���ű� a " _
                & "Where TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
                & "  and (a.���� like [1] or a.���� like [1])" _
                & "order by a.����"
    Case 2
        strsql = "Select ����ID id, סԺ��, ����, ��ǰ���� From ������Ϣ Where ��Ժʱ�� Is Null And (��ǰ����id = [1]) order by ����,��ǰ����,סԺ��"
    Case 3 '����ID
        strsql = "select rownum ID,a.��ǰ����ID ����ID,b.���� ��������,a.����,a.סԺ��,a.��ǰ���� ���� from ������Ϣ a, ���ű� b " _
               & "where a.��ǰ����id=b.id and a.����id=[1] order by a.����,a.��ǰ����,a.סԺ��"
    Case 4 '��������
        strsql = "select rownum ID,a.��ǰ����ID ����ID,b.���� ��������,a.����,a.סԺ��,a.��ǰ���� ���� from ������Ϣ a, ���ű� b " _
               & "where a.��ǰ����id=b.id and a.���� like [1] order by a.����,a.��ǰ����,a.סԺ��"
    Case 5 'סԺ��
        strsql = "select rownum ID,a.��ǰ����ID ����ID,b.���� ��������,a.����,a.סԺ��,a.��ǰ���� ���� from ������Ϣ a, ���ű� b " _
               & "where a.��ǰ����id=b.id and a.סԺ��=[1] order by a.סԺ��,a.����,a.��ǰ����"
    Case 6 '�����
        strsql = "select rownum ID,a.��ǰ����ID ����ID,b.���� ��������,a.����,a.סԺ��,a.��ǰ���� ���� from ������Ϣ a, ���ű� b " _
               & "where a.��ǰ����id=b.id and a.�����=[1] order by a.�����,a.����,a.��ǰ����,a.סԺ��"
    Case 7 '����
        strsql = "select rownum ID,a.��ǰ����ID ����ID,b.���� ��������,a.����,a.סԺ��,a.��ǰ���� ���� from ������Ϣ a, ���ű� b " _
               & "where a.��ǰ����id=b.id and a.��ǰ���� like [1] order by a.��ǰ����,a.����,a.סԺ��"
    End Select
    Set rstemp = zlDatabase.ShowSQLSelect(Me, strsql, 0, "ѡ����", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strParam)
    
    If blnCancel = True Then
        GoTo gtEmpty
        Exit Sub
    End If
    
    If Not rstemp Is Nothing Then
        Select Case bytIndex
        Case 1
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = zlStr.Nvl(rstemp!����)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����ID")) = zlStr.Nvl(rstemp!Id)
        Case 2
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("��������")) = zlStr.Nvl(rstemp!����)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("סԺ��")) = zlStr.Nvl(rstemp!סԺ��)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = zlStr.Nvl(rstemp!��ǰ����)
        Case Else
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = IIf(IsNull(rstemp!��������), "", rstemp!��������)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����ID")) = IIf(IsNull(rstemp!����id), "", rstemp!����id)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("��������")) = IIf(IsNull(rstemp!����), "", rstemp!����)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("סԺ��")) = IIf(IsNull(rstemp!סԺ��), "", rstemp!סԺ��)
            vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = IIf(IsNull(rstemp!����), "", rstemp!����)
        End Select
        rstemp.Close
    Else
        GoTo gtEmpty
    End If
    Exit Sub
    
gtEmpty:
    If bytIndex <> 2 Then
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = ""
        vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����ID")) = ""
    End If
    vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("��������")) = ""
    vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("סԺ��")) = ""
    vsfCostlyInfo.TextMatrix(1, vsfCostlyInfo.ColIndex("����")) = ""
    Exit Sub
ErrHandle:
    MsgBox "¼�������д�", vbCritical, gstrSysName
End Sub

Private Function GetCostlyInfoStr(ByVal intSN As Integer) As String
'��ֵ�����ַ���
    Dim strTmp As String
    If mrsCostlyInfo Is Nothing Then Exit Function
    With mrsCostlyInfo
        If .RecordCount <= 0 Then Exit Function
        .MoveFirst
        .Find "SN=" & intSN
        If Not .EOF Then
            strTmp = IIf(IsNull(mrsCostlyInfo!����), "", Trim(mrsCostlyInfo!����)) _
                   & "," & IIf(IsNull(mrsCostlyInfo!��������), "", Trim(mrsCostlyInfo!��������)) _
                   & "," & IIf(IsNull(mrsCostlyInfo!סԺ��), "", Trim(mrsCostlyInfo!סԺ��)) _
                   & "," & IIf(IsNull(mrsCostlyInfo!����), "", Trim(mrsCostlyInfo!����))
            If Replace(strTmp, ",", "") = "" Then Exit Function
            GetCostlyInfoStr = strTmp
        End If
    End With
End Function

Private Function Get�б굥λ�ɱ���(ByVal lng����ID As Long) As Double
    '----------------------------------------------------------------------------------------------
    '����:��ȡ�б굥λ�ĳɱ���
    '����:����id
    '����:�ɹ�,���سɱ���
    '����:������
    '����:2010/11/19
    '����:33718
    '----------------------------------------------------------------------------------------------
    Dim lng��Ӧ��λID As Long
    Dim rstemp As New ADODB.Recordset
    lng��Ӧ��λID = Val(txtProvider.Tag)
    err = 0: On Error GoTo ErrHand:
    gstrSQL = "Select �ɱ��� FROM �����б굥λ where ����id=[1] and ��λid=[2] "
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID, lng��Ӧ��λID)
    If rstemp.EOF Then
        Get�б굥λ�ɱ��� = 0
    Else
        Get�б굥λ�ɱ��� = Val(zlStr.Nvl(rstemp!�ɱ���))
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.Rows < 2 Then Exit Sub
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
        
        For n = 1 To mshBill.Rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !�к� = n
                !��� = IIf(Val(mshBill.TextMatrix(n, mCol���)) = 0, n, Val(mshBill.TextMatrix(n, mCol���)))
                !ҩƷid = Val(mshBill.TextMatrix(n, 0))
                !���� = Val(mshBill.TextMatrix(n, mCol����))
                
                .Update
            End If
        Next
        
    End With
End Sub

Private Function CheckRedo(ByVal rstemp As ADODB.Recordset) As ADODB.Recordset
    '���ܣ����ظ��ļ�¼���˵��������ع��˺�����ݼ���

    Dim i As Integer
    Dim strTemp As String
    Dim str���� As String
    Dim str����ID As String
    Dim str�ظ����� As String
    Dim strDub As String
    Dim strsql As String
    
    rstemp.MoveFirst
    str���� = ""
    Do While Not rstemp.EOF
        str���� = IIf(IsNull(rstemp!����), "0", rstemp!����)
        If InStr(1, strTemp, rstemp!����ID & "," & str����) = 0 Then
            strTemp = strTemp & rstemp!����ID & "," & str���� & "|"
        End If
        rstemp.MoveNext
    Loop
    
    With mshBill
        For i = 1 To .Rows - 1
            If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mCol����)) > 0 And .TextMatrix(i, 0) <> "" Then
                str����ID = str����ID & .TextMatrix(i, 0) & "," & .TextMatrix(i, mCol����) & "|"
            End If
        Next
        
        If str����ID <> "" Then   'Ϊ��������ƴ��sql
            strDub = ""
            For i = 0 To UBound(Split(str����ID, "|")) - 1
                strDub = strDub & "����id<>" & Split(Split(str����ID, "|")(i), ",")(0) & " and "
                If UBound(Split(str�ظ�����, ",")) <= 2 Then
                    str�ظ����� = str�ظ����� & Split(Split(str����ID, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        
        If str�ظ����� <> "" Then
            MsgBox str�ظ����� & "�б����Ѿ������ˣ�" & vbCrLf & "���ϲ��ϲ�����ӣ�", vbInformation, gstrSysName
            strsql = strDub
        End If
        rstemp.Filter = strsql
        Set CheckRedo = rstemp
    End With
End Function
