VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.UserControl usrTendFileEditor 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   8565
   Begin VB.PictureBox picBiref 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   120
      ScaleHeight     =   1785
      ScaleWidth      =   4395
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1740
      Visible         =   0   'False
      Width           =   4395
      Begin VB.ComboBox cbo��ʶ 
         Height          =   300
         Left            =   690
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   540
         Width           =   3495
      End
      Begin VB.ComboBox cboС�᷶Χ 
         Height          =   300
         Left            =   690
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   180
         Width           =   1365
      End
      Begin VB.TextBox txtС������ 
         Height          =   300
         Left            =   690
         TabIndex        =   42
         Top             =   900
         Width           =   3495
      End
      Begin VB.ComboBox cboС�� 
         Height          =   300
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   315
         Left            =   3570
         Picture         =   "usrTendFileEditor.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "ȡ��"
         Top             =   1320
         Width           =   450
      End
      Begin VB.CommandButton cmdOk 
         Height          =   315
         Left            =   3030
         Picture         =   "usrTendFileEditor.ctx":058A
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "ȷ��"
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label lbl��ʶ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʶ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   270
         TabIndex        =   39
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lbl��Χ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Χ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   270
         TabIndex        =   35
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������ݺ���ʾ��ȷ����"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   690
         TabIndex        =   45
         Top             =   1380
         Width           =   2010
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblС������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   270
         TabIndex        =   41
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblС�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "С��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2400
         TabIndex        =   37
         Top             =   240
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList imgRow 
      Left            =   6150
      Top             =   510
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
            Picture         =   "usrTendFileEditor.ctx":0B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrTendFileEditor.ctx":0EAE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCloumn 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   3075
      Left            =   60
      ScaleHeight     =   3075
      ScaleWidth      =   5955
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2070
      Visible         =   0   'False
      Width           =   5955
      Begin MSComctlLib.ListView lstColumnItems 
         Height          =   2490
         Left            =   180
         TabIndex        =   24
         Top             =   450
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   4392
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��Ŀ���"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��Ŀ����"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "��λ"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   2460
         Picture         =   "usrTendFileEditor.ctx":1248
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "ȷ��"
         Top             =   2310
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterCancel 
         Height          =   315
         Left            =   3000
         Picture         =   "usrTendFileEditor.ctx":17D2
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "ȡ��"
         Top             =   2310
         Width           =   450
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ѡ��(&S)"
         Height          =   300
         Index           =   0
         Left            =   2430
         TabIndex        =   25
         Top             =   1245
         Width           =   1100
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ɾ��(&E)"
         Height          =   300
         Index           =   1
         Left            =   2430
         TabIndex        =   26
         Top             =   1575
         Width           =   1100
      End
      Begin VB.TextBox txtColumnNo 
         Height          =   300
         Left            =   4545
         MaxLength       =   20
         TabIndex        =   30
         Top             =   120
         Width           =   1185
      End
      Begin MSComctlLib.ListView lstColumnUsed 
         Height          =   2490
         Left            =   3720
         TabIndex        =   31
         Top             =   450
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   4392
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��Ŀ���"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��Ŀ����"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "��λ"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ѷ������ݣ�������������á�"
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2340
         TabIndex        =   32
         Top             =   690
         Width           =   1260
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblColumnItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ�����¼��Ŀ:"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   180
         Width           =   1530
      End
      Begin VB.Label lblColumnNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͷ����:"
         Height          =   180
         Left            =   3735
         TabIndex        =   29
         Top             =   180
         Width           =   810
      End
   End
   Begin VB.TextBox txtLength 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   2340
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   3090
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   60
      ScaleHeight     =   4515
      ScaleWidth      =   8385
      TabIndex        =   10
      Top             =   510
      Width           =   8385
      Begin VB.CheckBox chkSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ȫѡ"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   930
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdWord 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6120
         Picture         =   "usrTendFileEditor.ctx":1D5C
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1290
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox picDouble 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6330
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2910
         Visible         =   0   'False
         Width           =   930
         Begin VB.PictureBox picDnInput 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   540
            ScaleHeight     =   255
            ScaleWidth      =   375
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   375
            Begin VB.Label lblDnInput 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   60
               TabIndex        =   20
               Top             =   30
               Width           =   315
            End
         End
         Begin VB.PictureBox picUpInput 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   435
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   435
            Begin VB.Label lblUpInput 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H80000008&
               Height          =   165
               Left            =   60
               TabIndex        =   19
               Top             =   30
               Width           =   315
            End
         End
         Begin VB.TextBox txtDnInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   525
            MaxLength       =   12
            TabIndex        =   7
            Top             =   30
            Width           =   345
         End
         Begin VB.TextBox txtUpInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   30
            MaxLength       =   12
            TabIndex        =   6
            Top             =   30
            Width           =   375
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Left            =   435
            TabIndex        =   16
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   1
         ItemData        =   "usrTendFileEditor.ctx":209E
         Left            =   6660
         List            =   "usrTendFileEditor.ctx":20B4
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   1590
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5790
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   1290
         Visible         =   0   'False
         Width           =   615
         Begin VB.TextBox txtInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   30
            Width           =   315
         End
         Begin VB.Label lblInput 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Caption         =   "��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   14
            Top             =   30
            Width           =   315
         End
      End
      Begin VB.PictureBox picMutilInput 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   5790
         ScaleHeight     =   435
         ScaleWidth      =   1575
         TabIndex        =   8
         Top             =   3330
         Visible         =   0   'False
         Width           =   1600
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   810
            TabIndex        =   9
            Top             =   90
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������¼"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   15
            TabIndex        =   13
            Top             =   112
            Width           =   720
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   0
         ItemData        =   "usrTendFileEditor.ctx":20EC
         Left            =   5790
         List            =   "usrTendFileEditor.ctx":2102
         TabIndex        =   3
         Top             =   1590
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   2655
         Left            =   0
         TabIndex        =   0
         Top             =   930
         Width           =   4305
         _cx             =   7594
         _cy             =   4683
         Appearance      =   0
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   1
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrTendFileEditor.ctx":213A
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AutoSizeMouse   =   0   'False
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
      Begin VB.Label lblCurPage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "P333"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   7650
         TabIndex        =   46
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "һ�㻤���¼��"
         Height          =   180
         Left            =   3450
         TabIndex        =   12
         Top             =   30
         Width           =   1275
      End
      Begin VB.Label lblSubhead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:##"
         Height          =   180
         Left            =   390
         TabIndex        =   11
         Top             =   540
         Width           =   720
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "usrTendFileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'��������:
'1.�����¼ͬһʱ��ֻ���ܴ���һ����¼
'2.�����¼�в���Ҫ�����µ����� , ��¼�����Ƿ����, �ܲ������, �����˵����ݲż�¼
'3.¼�뻤���¼����ʱ,�����¼������ݴ�����������, ����ȡ����
'4.�����¼���в���Ҫ¼�������¼�������׾����ȷ��Ҫ��¼���ڻ���ժҪ�������͵�����
'#ʵ��ԭ��:
'1.�����û��޸Ĺ�������,�����ṩ�༭״̬ҳ���л��Ĺ���,���û��޸Ĺ���ҳ���ݽ�����ҳ����,���ٳ���ʵ���Ѷ�
'2.���Ӽ�¼����¼��Щҳ��Щ��Ԫ���û��޸Ĺ�
'3.�κα༭(ճ��,�������),����Ҫ���¼���ÿ�����ݵ�ռ����

Public mblnEditable As Boolean
'Public objFileSys As New FileSystemObject
'Public objStream As TextStream

Private mfrmParent As Object
Private mblnInit As Boolean
Private mblnShow As Boolean                 '�Ƿ���ʾ¼���
Private mblnVerify As Boolean               '�Ƿ���ǩģʽ(���޸�,����������и���ճ������Ȳ���,ֻ���޸�)
Private mstrVerify As String                '�ȴ���ǩ��ID��
Private mintVerify As Integer               '��ǰ����Ա����߼���
Private mintVerify_Last As Integer          '��ѡ��ǩ��¼����߼���
Private mblnBlowup As Boolean               '�Ŵ�񣿷Ŵ�1/3��������9�ŷŴ�Ϊ12��
Private mblnChange As Boolean               '�Ƿ��޸�����
Private mstrData As String                  '����༭״̬ǰ����֮ǰ������
Private mintPreDays As Long
Private mstrMaxDate As String

Private mint����ҳ As Integer
Private mintҳ�� As Integer
Private mlng�ļ�ID As Long
Private mlng��ʽID As Long
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng����ID As Long
Private mlng����ID As Long
Private mintӤ�� As Integer
Private mbln���� As Boolean                 '�Ƿ���Ҫ¼������
Private mstrPrivs As String

Private mintSymbol As Integer               '��ǰ�ؼ�����
Private mstrSymbol As String                '�����ַ�
Private mblnClear As Boolean                '���Ϊ��,���mrsDataMap��¼��;����ҳʱӦ����,�����û��޸ĵ������Ա���ʾ������ʹ��
Private mstrCollectItems As String          '������Ŀ����
Private mstrColCollect As String            '������Ŀ�м���:col;1|col;4,5
Private mstrCOLNothing As String            'δ�󶨵��м���+���Ŀ��(���ܻ��Ŀ���Ƿ��)
Private mstrCOLActive As String             '��м���
Private mstrCatercorner As String           '�жԽ��߼���
Private mblnEditAssistant As Boolean        '��ǰѡ�����Ŀ�Ƿ�������дʾ�ѡ��
Private mlngPageRows As Long                '���ļ���ʽһҳ����ʾ��������
Private mlngOverrunRows As Long             '����������
Private mlngRowCount As Long                '��ǰ��¼������
Private mlngRowCurrent As Long              '��ǰ��¼�ڱ�ҳ��ʵ������
Private mlngDate As Long                    '����
Private mlngTime As Long                    'ʱ��
Private mlngChoose As Long                  'ѡ����
Private mlngOperator As Long                '��ʿ
Private mlngSignLevel As Long               'ǩ������
Private mlngSigner As Long                  'ǩ����Ϣ
Private mlngSignName As Long                'ǩ����
Private mlngSignTime As Long                'ǩ��ʱ��
Private mlngRecord As Long                  '��¼ID
Private mlngNoEditor As Long                '��ֹ�༭��,���ڻ�ʿ�����Ի�ʿ��Ϊ׼,�����ڻ�ʿ������ǩ����Ϊ׼
Private mlngCollectType As Long             '�������
Private mlngCollectText As Long             '�����ı�
Private mlngCollectStyle As Long            '���ܱ��
Private mlngCollectDay As Long              '��������:0-����;1-����

Private mblnSign As Boolean                 '�Ƿ�ǩ��
Private mblnArchive As Boolean              '�Ƿ�鵵
Private mintType As Integer                 '��¼��ǰ�ı༭ģʽ
Private mblnDateAd As Boolean               '������д?
Private mstr��ʼʱ�� As String              '��ǰ�ļ��Ŀ�ʼʱ��
Private mstr����ʱ�� As String              '��ǰ�ļ��Ľ���ʱ��
Private CellRect As RECT

Private rsTemp As New ADODB.Recordset
Private mrsItems As New ADODB.Recordset             '���л����¼��Ŀ�嵥
Private mrsSelItems As New ADODB.Recordset          '��ǰ¼��Ļ����¼��Ŀ�嵥
Private mrsDataMap As New ADODB.Recordset           '��ǰ����Ա¼������ݾ���,���¼����ʽһ��,���������ȫ�������Ա�Ѹ�ٻָ�
Private mrsCellMap As New ADODB.Recordset           '�༭�������ݾ���,�ֶ���:ҳ��,�к�,�к�,��¼ID,����,��λ,ɾ��
Private mrsCopyMap As New ADODB.Recordset           '����������

Private Enum ColIcon
    ǩ�� = 1
    ��ǩ = 2
End Enum
Private Enum SignLevel
    ���� = 1
    ���� = 2
    �м� = 3
    ʦ�� = 4
    Աʿ = 5
    δ���� = 9
End Enum

Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_CHILD = &H40000000
Private Const WS_POPUP = &H80000000
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Public Event AfterDataChanged(ByVal blnChange As Boolean)
Public Event AfterRefresh()
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)

Dim strFields As String
Dim strValues As String
Dim blnScroll As Boolean

'��¼�ϴ�ѡ����,����,�Ա�ˢ�º����¶�λ
Dim lngLastRow As Long
Dim lngLastTopRow As Long
Dim lngLastPatientID As Long

Private mstrTag As String           '�ݴ�

'�����ļ���ʽ�������
Private mintTabTiers As Integer     '��ͷ���
Private mintTagFormHour As Integer  '��ʼʱ������
Private mintTagToHour As Integer    '��ֹʱ������
Private mobjTagFont As New StdFont  '������ʽ����
Private mlngTagColor As Long        '������ʽ��ɫ
Private mstrPaperSet As String      '��ʽ
Private mstrPageHead As String      'ҳü
Private mstrPageFoot As String      'ҳ��
Private mblnChildForm As Boolean
Private mstrSubhead As String       '���ϱ�ǩ
Private mstrTabHead As String       '��ͷ��Ԫ
Private mstrColWidth As String      '�п����д�
Private mstrColumns As String       '��ǰ�����ļ����ж�Ӧ����Ŀ
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
'����򿪻����¼�ļ���SQL���������ط�Ҳ��ʹ�ã������޸�
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL���� As String
Private mstrSQL As String

'######################################################################################################################
'**********************************************************************************************************************
'��#�ָ��������ڵĴ��붼���ͼ���,û�±�
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const WHITE_BRUSH = 0    '��ɫ����
Private Const cdblWidth As Double = 6          'һ��Ӣ���ַ��Ŀ��
Private Const cHideCols = 3         'ǰ׺������:����,ʱ��,ѡ��
Private Const cControlFields = 2    '��¼��������:ҳ��,�к�

Private Function GetRBGFromOLEColor(ByVal dwOleColour As Long) As Long
    '��VB����ɫת��ΪRGB��ʾ
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long
    
    OleTranslateColor dwOleColour, 0, clrref
    
    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF
    
    GetRBGFromOLEColor = RGB(r, g, b)
End Function

Private Function GetSymbolWidth(ByVal strPara As String) As Double
    'ȱʡ������9��,�������Сͬ�ȷŴ�
    Dim sinFontSize As Single
    Dim i As Integer, j As Integer
    
    j = Len(strPara)
    sinFontSize = VsfData.FontSize
    For i = 1 To j
        GetSymbolWidth = GetSymbolWidth + IIf(Asc(Mid(strPara, i, 1)) > 0, 1, 2) * cdblWidth * sinFontSize / 9
    Next
End Function

Private Sub DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim strText As String
    Dim strLeft As String
    Dim strRight As String
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim dblWidth As Double
    Dim lngBackColor As Long
    Dim lngForeColor As Long
    Dim blnDraw As Boolean
    '��ͼ���
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim lngBrush As Long
    Dim lngOldBrush As Long
    Dim lpPoint As POINTAPI
    Dim t_ClientRect As RECT
    On Error GoTo errHand
    '******************************************
    '�ڴ��¼��в��ܶԵ�Ԫ����κ����Ը�ֵ,����Celldata,�����������¼�����ѭ��,���¹��������ʱ���޷�����������
    '******************************************
    'ʹ��ƥ��ı���ɫ��ǰ��ɫ����������ı������
    If Not mblnInit Then Exit Sub
    If VsfData.RowHidden(ROW) Then Exit Sub
    Done = False
    
    strText = VsfData.TextMatrix(ROW, COL)
    If IsDiagonal(COL) And InStr(1, strText, "/") <> 0 Then
        blnDraw = True
        '����ֵ
        strLeft = Split(strText, "/")(0)
        strRight = Mid(strText, InStr(1, strText, "/") + 1)
        lngLeft = LenB(StrConv(strLeft, vbFromUnicode))
        lngRight = LenB(StrConv(strRight, vbFromUnicode))
        'ȡ�ַ����
        dblWidth = GetSymbolWidth(strRight)
        '�趨�ͻ������С
        With t_ClientRect
            .Left = Left + 1
            .Top = Top + 1
            .Right = Right - 1
            .Bottom = Bottom - 1
        End With
        
        '1���������
        '�����뱳��ɫ��ͬ��ˢ��
        If ROW < VsfData.FixedRows Then
            lngBackColor = GetRBGFromOLEColor(VsfData.BackColorFixed)
            lngForeColor = GetRBGFromOLEColor(VsfData.ForeColorFixed)
        Else
            If ROW = VsfData.RowSel Then
                lngBackColor = GetRBGFromOLEColor(VsfData.BackColorSel)
                lngForeColor = RGB(0, 0, 0)
            Else
                lngBackColor = RGB(255, 255, 255)
                lngForeColor = GetRBGFromOLEColor(VsfData.Cell(flexcpForeColor, ROW, COL))
            End If

        End If
        lngBrush = CreateSolidBrush(lngBackColor)
        'ʹ�ø�ˢ����䱳��ɫ
        lngOldBrush = SelectObject(hDC, lngBrush)
        Call FillRect(hDC, t_ClientRect, lngBrush)
        '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
        Call SelectObject(hDC, lngOldBrush)
        Call DeleteObject(lngBrush)
        
        '2��׼������
        '�����»���
        Call SetTextColor(hDC, lngForeColor)
        lngPen = CreatePen(0, 1, lngForeColor)
        lngOldPen = SelectObject(hDC, lngPen)
        '����
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Top)
        '����ı�
        Call TextOut(hDC, Left, Top, strLeft, lngLeft)
        Call TextOut(hDC, IIf(Right - dblWidth >= Left, Right - dblWidth, Left), Bottom - 16, strRight, lngRight)
        
        '��ԭ���ʲ�����
        Call SelectObject(hDC, lngOldPen)
        Call DeleteObject(lngPen)
        
        '�������ͼ
        Done = True
    End If
    
    '3������ǻ����У���������⴦��
    If Val(VsfData.TextMatrix(ROW, mlngCollectType)) < 0 And Val(VsfData.TextMatrix(ROW, mlngCollectStyle)) > 0 _
        And (COL >= mlngDate And COL < mlngNoEditor) Then
        Call DrawCollectCell(hDC, ROW, COL, Left, Top, Right, Bottom)
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DrawCollectCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Dim lngPen As Long, lngOldPen As Long
    Dim lpPoint As POINTAPI
    
    '�����»���
    lngPen = CreatePen(0, 1, vbRed)
    lngOldPen = SelectObject(hDC, lngPen)
    
    If Val(VsfData.TextMatrix(ROW, mlngCollectStyle)) = 1 Then  '���»�����
        '����
        Call MoveToEx(hDC, Left, Top, lpPoint)
        Call LineTo(hDC, Right, Top)
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Bottom - 2)
    Else                                                        '��������˫����
        If InStr(1, "|" & mstrColCollect & ";", "|" & COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Then 'And Val(VsfData.TextMatrix(ROW, COL)) <> 0 Then
            '����
            Call MoveToEx(hDC, Left, Bottom - 4, lpPoint)
            Call LineTo(hDC, Right, Bottom - 4)
            Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
            Call LineTo(hDC, Right, Bottom - 2)
        End If
    End If
    
    '��ԭ���ʲ�����
    Call SelectObject(hDC, lngOldPen)
    Call DeleteObject(lngPen)
End Sub

'######################################################################################################################
'**********************************************************************************************************************
'��#�ָ��������ڵĴ��붼��������,û�±�
Private Function GetData(ByVal strInput As String) As Variant
    Dim arrData
    Dim strData As String
    Dim strLine(256) As Byte
    Dim lngRow As Long, lngRows As Long
    
    GetData = ""
    lngRows = SendMessage(txtLength.hwnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngRows
        Call ClearArray(strLine)
        Call SendMessage(txtLength.hwnd, EM_GETLINE, lngRow - 1, strLine(0))
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", "|ZYB.ZLSOFT|") & strData & IIf(lngRow < lngRows, vbCrLf, "")
    Next
    GetData = Split(GetData, "|ZYB.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = 0 To intMax
        strLine(intDo) = 0
    Next
    strLine(1) = 1
End Sub

Private Function TrimStr(ByVal str As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�������ȥ�����˵Ŀո�

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Private Function TruncZero(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

'**********************************************************************************************************************
'######################################################################################################################


Private Sub BoundItems(ByVal intCol As Integer)
    Dim lstItem As ListItem
    Dim rsActive As New ADODB.Recordset
    On Error GoTo errHand
    'ֻ�ṩ������,ѡ����������Ļ��Ŀ
    '�󶨻��Ŀ(��һ����Ŀ������,��������Ŀʱ,��Ŀ���ͱ���=0����Ŀ��ʾֻ������ֵ,ѡ������,��������Ŀ��Ŀ��������Ŀ��ʾ��������һ��)
    
    gstrSQL = "" & _
        " SELECT /*+ RULE */ A.��Ŀ���,A.��λ,A.��Ŀ����,B.��ͷ����,NVL(B.��־,0) AS ��־" & vbNewLine & _
        " FROM" & vbNewLine & _
        "     (SELECT A.��Ŀ���,B.��λ,B.��λ||A.��Ŀ���� AS ��Ŀ����" & vbNewLine & _
        "     FROM �����¼��Ŀ A,���²�λ B" & vbNewLine & _
        "     WHERE A.��Ŀ��� =B.��Ŀ���(+) AND A.��Ŀ����=2 And A.��Ŀ����=0 And A.��Ŀ��ʾ IN (0,4,5)) A," & vbNewLine & _
        "     (SELECT A.��ͷ����,A.��Ŀ���,A.��λ||B.��Ŀ���� AS ��Ŀ����,1 AS ��־" & vbNewLine & _
        "     FROM ���˻���ҳ��_���Ŀ A,�����¼��Ŀ B" & vbNewLine & _
        "     WHERE A.��Ŀ���=B.��Ŀ��� AND A.�ļ�ID=[1] AND A.ҳ��=[2] AND A.�к�=[3]) B" & vbNewLine & _
        " WHERE A.��Ŀ���=B.��Ŀ���(+) AND A.��Ŀ����=B.��Ŀ����(+)" & vbNewLine & _
        " ORDER BY A.��Ŀ���"
    Set rsActive = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡδ���õĻ��Ŀ", mlng�ļ�ID, mintҳ��, intCol)
    If rsActive.RecordCount = 0 Then
        RaiseEvent AfterRowColChange("û�пɹ�ѡ��Ļ��Ŀ�����ڻ�����Ŀ����ģ���н������ã�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '������Ŀ
    lstColumnItems.ListItems.Clear
    lstColumnUsed.ListItems.Clear
    With rsActive
        Do While Not .EOF
            If !��־ = 1 Then
                txtColumnNo.Text = NVL(!��ͷ����)
                Set lstItem = lstColumnUsed.ListItems.Add(, Now & "_" & !��Ŀ��� & "_" & lstColumnUsed.ListItems.Count, !��Ŀ���)
                lstItem.SubItems(1) = !��Ŀ����
                lstItem.SubItems(2) = NVL(!��λ)
            Else
                Set lstItem = lstColumnItems.ListItems.Add(, Now & "_" & !��Ŀ��� & "_" & lstColumnItems.ListItems.Count + 100, !��Ŀ���)
                lstItem.SubItems(1) = !��Ŀ����
                lstItem.SubItems(2) = NVL(!��λ)
            End If
            .MoveNext
        Loop
    End With
    
    '���ÿؼ����꣨��߻��ұ߳�����Ļ��С���һ�����ʾ����������Ϊ������ʾ��
    With picCloumn
        .Left = VsfData.Left + VsfData.CellLeft + VsfData.CellWidth / 2 - .Width / 2
        .Top = picMain.Top + VsfData.Top + VsfData.CellTop
        If .Height + .Top + picMain.Top > ScaleHeight Then
            .Top = ScaleHeight - picMain.Top - .Height
        End If
        If .Left + .Width > ScaleWidth Then
            .Left = ScaleWidth - .Width
        End If
        If .Left < VsfData.Left Then
            .Left = VsfData.Left
        End If
        .Visible = True
    End With
    
    lblNote.Visible = ISColHaveData
    cmdColumn(0).Enabled = Not lblNote.Visible
    cmdColumn(1).Enabled = Not lblNote.Visible
    cmdFilterOK.Enabled = Not lblNote.Visible
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetPeriod() As String
    On Error GoTo errHand
    
    '��ȡָ��ҳ������ݷ���ʱ�䷶Χ
    gstrSQL = " Select /*+ RULE */ MIN(����ʱ��) ��ʼʱ��,MAX(����ʱ��) AS ����ʱ�� From ���˻����ӡ Where �ļ�ID=[1] And (��ʼҳ��=[2] OR ����ҳ��=[2])"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ��ҳ������ݷ���ʱ�䷶Χ", mlng�ļ�ID, mintҳ��)
    If NVL(rsTemp!��ʼʱ��) = "" Then
        If mintӤ�� = 0 Then
            gstrSQL = " Select  /*+ RULE */ ��Ժ���� AS ��ʼʱ��,sysdate AS ����ʱ�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
        Else
            gstrSQL = " Select  /*+ RULE */ ����ʱ�� AS ��ʼʱ��,sysdate AS ����ʱ�� From ������������¼ Where ����ID=[1] And ��ҳID=[2] And ���=[3]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ժ���ڻ��������", mlng����ID, mlng��ҳID, mintӤ��)
    End If
    GetPeriod = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "��" & Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadStruDef() As Boolean
    Dim lngCOL As Long
    On Error GoTo errHand
    
    '��ȡ�ļ�����
    mblnDateAd = False
    Call GetFileProperty
    
    '��ȡ���Ŀ�������ж���(��ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...)
    mstrCOLActive = ""
    mstrCOLNothing = ""
    mstrCollectItems = ""
    mstrColCollect = ""
    gstrSQL = " Select  /*+ RULE */ A.�к�,A.��ͷ����,A.���,A.��Ŀ���,A.��λ From ���˻���ҳ��_���Ŀ A " & _
              " Where A.�ļ�ID=[1] And A.ҳ��=[2] " & _
              " Order by A.�к�,A.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������Զ���Ļ��Ŀ", mlng�ļ�ID, mintҳ��)
    If rsTemp.RecordCount <> 0 Then
        Do While Not rsTemp.EOF
            If lngCOL <> rsTemp!�к� Then
                lngCOL = rsTemp!�к�
                mstrCOLActive = mstrCOLActive & "||" & rsTemp!�к� & ";" & rsTemp!��ͷ���� & "|" & rsTemp!��Ŀ��� & "," & NVL(rsTemp!��λ)
            Else
                mstrCOLActive = mstrCOLActive & ";" & rsTemp!��Ŀ��� & "," & NVL(rsTemp!��λ)
            End If
            rsTemp.MoveNext
        Loop
    End If
    If mstrCOLActive <> "" Then mstrCOLActive = Mid(mstrCOLActive, 3)
    
    '��ȡ�����ļ���ʽ����
    gstrSQL = "Select  /*+ RULE */ d.�������, d.�����ı�, d.Ҫ������" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ���ʽ����", mlng��ʽID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !Ҫ������
            Case "��ͷ����": mintTabTiers = Val("" & !�����ı�)
            Case "������":  VsfData.Cols = Val("" & !�����ı�)
            Case "��С�и�": VsfData.RowHeightMin = BlowUp(Val("" & !�����ı�))
            Case "�ı�����"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set VsfData.Font = objFont
                Set lblSubhead.Font = VsfData.Font
                Set Font = lblSubhead.Font
                
            Case "�ı���ɫ": VsfData.ForeColor = Val("" & !�����ı�)
            Case "�����ɫ": VsfData.GridColor = Val("" & !�����ı�): VsfData.GridColorFixed = VsfData.GridColor
            
            Case "�����ı�"
                lblTitle.Caption = "" & !�����ı�
                lblTitle.AutoSize = True
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set lblTitle.Font = objFont
                lblTitle.AutoSize = False
            
            Case "��ʼʱ��": mintTagFormHour = Val("" & !�����ı�)
            Case "��ֹʱ��": mintTagToHour = Val("" & !�����ı�)
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "������ɫ": mlngTagColor = Val("" & !�����ı�)
            Case "��Ч������"
                mlngOverrunRows = 0
                mlngPageRows = Val(!�����ı�)
            End Select
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select  /*+ RULE */ ��ʽ, ҳü, ҳ��,���� From ����ҳ���ʽ Where ���� = 3 And ��� In (Select ҳ�� From �����ļ��б� Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҳ���ʽ", mlng��ʽID)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!��ʽ: mstrPageHead = "" & rsTemp!ҳü: mstrPageFoot = "" & rsTemp!ҳ��
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select  /*+ RULE */ d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ϱ�ǩ����", mlng��ʽID)
    With rsTemp
        mstrSubhead = ""
        Do While Not .EOF
            mstrSubhead = mstrSubhead & "|" & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
            .MoveNext
        Loop
        If mstrSubhead <> "" Then mstrSubhead = Mid(mstrSubhead, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select  /*+ RULE */ d.�������, d.�����д�, d.�����ı�" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ͷ��Ԫ����", mlng��ʽID)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !�����д� - 1 & "," & !������� & "," & !�����ı�
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '��ѯ�����֯
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql�� As String, str��ʽ As String
    Dim bln���� As Boolean, blnʱ�� As Boolean, bln��ʿ As Boolean
    Dim blnǩ���� As Boolean, blnǩ��ʱ�� As Boolean, blnǩ������ As Boolean
    Dim bln�Խ��� As Boolean, blnѡ���� As Boolean          '�����һ���ǶԽ�����ѡ����,��ֱ����ȡ��������,ƴ��ͷʱ����ֵ�����/
    Dim lngColumn As Long, blnAddCollect As Boolean
    
    gstrSQL = "Select  /*+ RULE */ d.�������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
        " Order By d.�������, d.�����д�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���м��϶���", mlng��ʽID)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = ""
        mstrSQL�� = "": mstrSQL�� = "": strSql�� = "": mstrSQL�� = "": mstrSQL���� = ""
        bln���� = False: blnʱ�� = False: bln��ʿ = False
        blnǩ���� = False: blnǩ��ʱ�� = False: blnǩ������ = False
        Do While Not .EOF
            If lngColumn <> !������� Then
                blnAddCollect = False
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) & "|" & !������� & "'" & !Ҫ������
                mstrColWidth = mstrColWidth & "," & !�������� & "`" & !������� & "`" & !Ҫ�ر�ʾ
                If !Ҫ�ر�ʾ = 1 Then mstrCatercorner = mstrCatercorner & "," & !�������
                str��ʽ = ""
                If !Ҫ������ <> "" Then
                    str��ʽ = "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
                    mstrSQL�� = mstrSQL�� & "," & IIf(Mid(strSql��, 3) = "", "''", Mid(strSql��, 3)) & " As C" & Format(lngColumn, "00")
                Else
                    If strSql�� <> "" Then
                        mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        mstrSQL�� = mstrSQL�� & ",'' As C" & Format(lngColumn, "00")
                    End If
                End If
                strSql�� = ""
                lngColumn = !�������
                bln�Խ��� = (NVL(!Ҫ�ر�ʾ, 0) = 1)
                blnѡ���� = False
                mrsItems.Filter = "��Ŀ����='" & NVL(!Ҫ������) & "'"
                If mrsItems.RecordCount <> 0 Then
                    blnѡ���� = (mrsItems!��Ŀ��ʾ = 5)
                    If mrsItems!��Ŀ��ʾ = 4 Then   '������Ŀ
                        blnAddCollect = True
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!��Ŀ���
                        mstrColCollect = mstrColCollect & "|" & !������� & ";" & mrsItems!��Ŀ���
                    End If
                End If
                mrsItems.Filter = 0
            Else
                mstrColumns = mstrColumns & "," & !Ҫ������
                str��ʽ = str��ʽ & "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
                mrsItems.Filter = "��Ŀ����='" & NVL(!Ҫ������) & "'"
                If mrsItems.RecordCount <> 0 Then
                    If mrsItems!��Ŀ��ʾ = 4 Then   '������Ŀ
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!��Ŀ���
                        If blnAddCollect Then
                            mstrColCollect = mstrColCollect & "," & mrsItems!��Ŀ���
                        Else    '�п���һ�а�������Ŀ,��һ����Ŀ���ǻ�����Ŀ,�ڶ�����Ŀ���ǻ�����Ŀ,���,����Ĵ��뱣֤���������
                            mstrColCollect = mstrColCollect & "|" & !������� & ";" & mrsItems!��Ŀ���
                        End If
                    End If
                End If
                mrsItems.Filter = 0
            End If
            
            Select Case !Ҫ������
            Case "����"
                bln���� = True
                mblnDateAd = (NVL(!Ҫ�ر�ʾ, 0) = 1)
                mstrSQL�� = mstrSQL�� & ",����"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, " & IIf(mblnDateAd, "'dd/MM'", "'yyyy-mm-dd'") & ") As ����"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case "ʱ��"
                blnʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ʱ��"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ����"
                blnǩ���� = True
                mstrSQL�� = mstrSQL�� & ",ǩ����"
                mstrSQL�� = mstrSQL�� & ",l.ǩ����"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ��ʱ��"
                blnǩ��ʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ǩ��ʱ��"
                mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "��ʿ"
                bln��ʿ = True
                mstrSQL�� = mstrSQL�� & ",��ʿ"
                mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case Else
                If !Ҫ������ <> "" Then
                    mstrSQL�� = mstrSQL�� & ",Max(""" & !Ҫ������ & """) As """ & !Ҫ������ & """"
                    mstrSQL���� = mstrSQL���� & " Or """ & !Ҫ������ & """ Is Not Null"
                    
                    If bln�Խ��� And blnѡ���� Then
                        If strSql�� <> "" Then
                            '�ڶ���
                            strSql�� = strSql�� & "||'/'||""" & !Ҫ������ & """"
                        Else
                            '��һ��
                            strSql�� = strSql�� & "||""" & !Ҫ������ & """"
                        End If
                    Else
                        strSql�� = strSql�� & "||""" & !Ҫ������ & """"
                    End If
                    
                    If (Trim("" & !�����ı�) = "" And Trim("" & !Ҫ�ص�λ) = "") Or (bln�Խ��� And blnѡ����) Then
                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,c.��¼����), '') As """ & !Ҫ������ & """"
                    Else
                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "','" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "')), '') As """ & !Ҫ������ & """"
                    End If
                Else
                    'Ϊ�ձ�ʾδ����,ǿ�Ƽ�,��������滻
                    mstrCOLNothing = mstrCOLNothing & "," & Format(!�������, "00")
                    mstrSQL�� = mstrSQL�� & ",Max(""" & "C" & Format(!�������, "00") & """) As C" & Format(!�������, "00")
                    mstrSQL���� = mstrSQL���� & " Or """ & "C" & Format(!�������, "00") & """ Is Not Null"
                    mstrSQL�� = mstrSQL�� & ", C" & Format(!�������, "00") & " AS C" & Format(!�������, "00")
                End If
            End Select
            .MoveNext
        Loop
        
        If mstrCollectItems <> "" Then
            mstrCollectItems = Mid(mstrCollectItems, 2)
            mstrColCollect = Mid(mstrColCollect, 2)
        End If
        mstrCOLNothing = Mid(mstrCOLNothing, 2)
        mstrCatercorner = Mid(mstrCatercorner, 2)
        mstrColWidth = Mid(mstrColWidth, 2)
        '�������һ�еĸ�ʽ
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) '& "|" & !������� & "'" & !Ҫ������
        mstrColumns = Mid(mstrColumns, 2)     '��ʽ��:�к�;��Ŀ����1,��Ŀ����2|�к�...,ʵ��;1;����|2;����|3...
        If Mid(strSql��, 3) <> "" Then
            mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
        Else
            mstrSQL�� = mstrSQL�� & ",'' As C" & Format(lngColumn, "00")
        End If
        
        If mstrSQL���� <> "" Then mstrSQL���� = "(" & Mid(mstrSQL����, 5) & ")"
        
        '���û�г������ڣ�ʱ�䣬��ʿ�����ڲ���Ҫ���䣬�Ա�֤�в�����������
        If bln���� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'yyyy-mm-dd') As ����"
        If blnʱ�� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
        If bln��ʿ = False Then mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
        
        If blnǩ���� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ���� As ǩ����"
        If blnǩ��ʱ�� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"
        
        If Mid(mstrSQL��, 2) = "" Then
            MsgBox "�Բ�����û�ж��嵱ǰ��������ʾ����Ϣ�����ڲ����ļ������ж��壡", vbInformation, gstrSysName
            Exit Function
        End If
        
        '�����ڲ��������ӹ̶���
        mstrSQL�� = UCase(mstrSQL�� & ",MAX(ǩ������) AS ǩ������,MAX(ǩ����Ϣ) AS ǩ����Ϣ,MAX(��¼ID) AS ��¼ID,MAX(����) AS ����,MAX(ʵ������) AS ʵ������,MAX(�������) AS �������,MAX(�����ı�) AS �����ı�,MAX(���ܱ��) AS ���ܱ��,MAX(��������) AS ��������")
        mstrSQL�� = UCase(mstrSQL�� & ",l.ǩ������,l.ǩ���� AS ǩ����Ϣ,C.��¼ID,P.����||'' AS ����,DECODE(SIGN(P.����ҳ��-P.��ʼҳ��),1,DECODE(SIGN([5]-P.��ʼҳ��),1, P.�����к�,P.����-P.�����к� ),P.����) AS ʵ������,NVL(L.�������,0) AS �������,L.�����ı�,L.���ܱ��,to_char(L.����ʱ��,'yyyy-MM-dd hh24:mi:ss')||'' AS ��������")
        mstrSQL�� = UCase(mstrSQL�� & ",ǩ������,ǩ����Ϣ,��¼ID,����,ʵ������,�������,�����ı�,���ܱ��,��������")
        
        '�����Ŀ���뵽SQL��
        Call PreActiveCOL
        Call SQLCombination
    End With
    
    ReadStruDef = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub PreActiveHead()
    Dim arrData
    Dim intCol As Integer
    Dim strName As String
    Dim intDo As Integer, intCount As Integer
    '���±�ͷ
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        strName = Split(Split(arrData(intDo), "|")(0), ";")(1)
        VsfData.TextMatrix(mintTabTiers - 1, intCol + cHideCols + VsfData.FixedCols - 1) = strName
        If mintTabTiers = 3 And VsfData.TextMatrix(1, intCol + cHideCols + VsfData.FixedCols - 1) = "" Then VsfData.TextMatrix(1, intCol + cHideCols + VsfData.FixedCols - 1) = strName
        If mintTabTiers = 2 And VsfData.TextMatrix(0, intCol + cHideCols + VsfData.FixedCols - 1) = "" Then VsfData.TextMatrix(0, intCol + cHideCols + VsfData.FixedCols - 1) = strName
    Next
    
    With chkSwitch
        .Value = 0
        .Top = VsfData.Top + VsfData.Cell(flexcpTop, mintTabTiers - 1, mlngChoose) + VsfData.Cell(flexcpHeight, mintTabTiers - 1, mlngChoose) - .Height
        .Left = VsfData.Left + VsfData.Cell(flexcpLeft, mintTabTiers - 1, mlngChoose) + 50
        .Visible = mblnVerify
    End With
End Sub

Private Sub PreActiveCOL()
    Dim arrData
    Dim arrCol
    Dim intCol As Integer, strName As String
    Dim strColFormat As String, strCOLNames As String, strCOLPart As String, strCOLCOND As String, strCOLDEF As String, strCOLMID As String, strCOLIN As String
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '�����Ŀ���뵽��ѯSQL�У���ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...
    '�󶨶����Ŀ�����о��Զ�תΪ�Խ�����
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        strName = Split(Split(arrData(intDo), "|")(0), ";")(1)
        arrCol = Split(Split(arrData(intDo), "|")(1), ";")
        
        '�����б�ʾ(ÿ������������Ŀ)
        strCOLPart = ""
        strCOLNames = ""
        strColFormat = ""
        strCOLCOND = ""
        strCOLMID = ""
        strCOLIN = ""
        strCOLDEF = ""
        intMax = UBound(arrCol)
        For intIn = 0 To intMax
            strCOLPart = Split(arrCol(intIn), ",")(1)
            mrsItems.Filter = "��Ŀ���=" & Val(Split(arrCol(intIn), ",")(0))
            strCOLNames = strCOLNames & "," & mrsItems!��Ŀ����
            strCOLCOND = strCOLCOND & " OR """ & strCOLPart & mrsItems!��Ŀ���� & """ IS NOT NULL"
            strCOLMID = strCOLMID & ",Max(""" & strCOLPart & mrsItems!��Ŀ���� & """) As """ & strCOLPart & mrsItems!��Ŀ���� & """"
            If intIn = 0 Then
                strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.���²�λ||") & "c.��Ŀ����, '" & strCOLPart & mrsItems!��Ŀ���� & "', Nvl(c.δ��˵��,c.��¼����), '') As """ & strCOLPart & mrsItems!��Ŀ���� & """"
            Else
                strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.���²�λ||") & "c.��Ŀ����, '" & strCOLPart & mrsItems!��Ŀ���� & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,'/','/'||c.��¼����||'')), '') As """ & strCOLPart & mrsItems!��Ŀ���� & """"
            End If
            If intIn = 0 Then
                If intMax = 0 Then
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!��Ŀ���� & """ AS C" & intCol
                Else
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!��Ŀ���� & """||"
                End If
            Else
                strCOLDEF = strCOLDEF & "NVL(" & strCOLPart & mrsItems!��Ŀ���� & ",'/') AS C" & intCol
            End If
            
            strColFormat = strColFormat & "{[" & strCOLPart & mrsItems!��Ŀ���� & "]" & IIf(intMax > 0 And intIn = 0, "/", "") & "}"
        Next
        If strCOLPart <> "" Then
            strCOLPart = Mid(strCOLPart, 2)
        End If
        strCOLNames = Mid(strCOLNames, 2)
        
        '�Խ���
        If intMax > 0 Then
            mstrCatercorner = mstrCatercorner & IIf(mstrCatercorner = "", "", ",") & intCol
        End If
        '�и�ʽ:15'��ʿ'1'{[��ʿ]}
        mstrColumns = Replace(mstrColumns, intCol & "''1'", intCol & "'" & strCOLNames & "'1'" & strColFormat)
        '��
        mstrSQL�� = Replace(mstrSQL��, "'' AS C" & Format(intCol, "00"), strCOLDEF)
        '����
        mstrSQL���� = Replace(UCase(mstrSQL����), " OR """ & "C" & Format(intCol, "00") & """ IS NOT NULL", strCOLCOND)
        '��
        mstrSQL�� = Replace(mstrSQL��, ",MAX(""" & "C" & Format(intCol, "00") & """) AS C" & Format(intCol, "00"), strCOLMID)
        '��
        mstrSQL�� = Replace(mstrSQL��, ", C" & Format(intCol, "00") & " AS C" & Format(intCol, "00"), strCOLIN)
    Next
    mrsItems.Filter = 0
    
    '��δ�󶨵��е�SQL�������
    If mstrCOLNothing = "" Then Exit Sub
    arrData = Split(mstrCOLNothing, ",")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        '��(����Ҫ����)
'        mstrSQL�� = Replace(mstrSQL��, ",'' AS C" & arrData(intDo), "")
        '����
        mstrSQL���� = Replace(UCase(mstrSQL����), " OR """ & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL", "")
        '��
        mstrSQL�� = Replace(mstrSQL��, ",MAX(""" & "C" & Format(arrData(intDo), "00") & """) AS C" & Format(arrData(intDo), "00"), "")
        '��
        mstrSQL�� = Replace(mstrSQL��, ", C" & Format(arrData(intDo), "00") & " AS C" & Format(arrData(intDo), "00"), "")
    Next
End Sub

Private Sub SQLCombination(Optional ByVal lng��¼ID As Long = 0)
    Dim str���� As String
    str���� = mstrSQL���� & IIf(lng��¼ID = 0, "", IIf(mstrSQL���� = "", "", " And") & " ��¼ID=[6]")
    
    mstrSQL = "Select  /*+ RULE */ ����,����ʱ��,'' AS ѡ��," & Mid(mstrSQL��, 12) & vbCrLf & _
                " From (Select ��¼���,ʱ�� as ����,����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "        From (Select c.��¼���,l.����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "               From ���˻������� l, ���˻�����ϸ c,���˻����ļ� f,���˻����ӡ p " & vbCrLf & _
                "               Where l.ID=p.��¼ID And l.Id = c.��¼id And l.�ļ�ID=f.ID And f.ID=p.�ļ�ID " & _
                "               And c.��ֹ�汾 Is Null And c.��¼����<>5  " & _
                "               And f.id=[1] And f.����id = [2] And f.��ҳid = [3] And Nvl(f.Ӥ��,0)=[4] And (P.��ʼҳ��=[5] Or P.����ҳ��=[5]))" & vbCrLf & _
                IIf(str���� <> "", "Where " & str����, "") & _
                "       Group By ����, ʱ��, ����ʱ��,��¼���,��ʿ,ǩ����,ǩ��ʱ��" & _
                                "       Order By ����ʱ��,��¼���,��ʿ,ǩ����,ǩ��ʱ��)"
End Sub

Private Sub zlRefresh()
    Dim aryRow() As String, aryItem() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCOL As Long, lngCount As Long, strCell As String
    Dim strTmpSQL As String
    Dim aryPeriod() As String
    Dim strtmp As String
    
    Err = 0: On Error GoTo errHand
    Call InitCons
    '���ϱ�ǩ��ȡ
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    aryPeriod = Split(GetPeriod, "��")
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as ��Ϣ From Dual"
    aryItem = Split(mstrSubhead, "|")
    
        
    For lngCount = 0 To UBound(aryItem)
        strPrefix = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") - 1)
        strItemName = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") + 1, InStr(1, aryItem(lngCount), "}") - InStr(1, aryItem(lngCount), "{") - 1)
        
        strtmp = strPrefix
        Select Case strItemName
        Case "��ǰ����"
        
            strTmpSQL = "Select  /*+ RULE */ b.����" & vbNewLine & _
                        "From (Select ����id, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a,���ű� b " & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����id Is Not Null And b.ID=a.����id" & vbNewLine & _
                        "Order By a.��ʼʱ��"
                        
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            
        Case "��ǰ����"
        
            strTmpSQL = "Select  /*+ RULE */ a.����" & vbNewLine & _
                        "From (Select ����, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���� Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"

            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case "��ǰ����"
        
            strTmpSQL = "Select  /*+ RULE */ ���� From ���ű� a Where a.ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID)
            
        Case "סԺҽʦ"
            strTmpSQL = "Select  /*+ RULE */ a.����ҽʦ" & vbNewLine & _
                        "From (Select ����ҽʦ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ҽʦ Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "סԺҽʦ", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
        Case "���λ�ʿ"
        
            strTmpSQL = "Select  /*+ RULE */ a.���λ�ʿ" & vbNewLine & _
                        "From (Select ���λ�ʿ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���λ�ʿ Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "���λ�ʿ", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case "����ȼ�"
            strTmpSQL = "Select  /*+ RULE */ b.����" & vbNewLine & _
                        "From (Select ����ȼ�ID, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a,����ȼ� b" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ȼ�ID Is Not Null And b.���=a.����ȼ�ID" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "����ȼ�", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case Else
            strtmp = ""
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", strPrefix, strItemName, mlng����ID, mlng��ҳID, mintӤ��)
        End Select
        
        If rsTemp.BOF = False Then
            If strtmp <> "" Then
                lblSubhead.Tag = lblSubhead.Tag & " " & strtmp & rsTemp.Fields(0).Value
            Else
                lblSubhead.Tag = lblSubhead.Tag & " " & rsTemp.Fields(0).Value
            End If
        End If
    Next
    lblSubhead.Tag = Trim(lblSubhead.Tag)
    
    '���ϱ�ǩ��ɢ����
    Call zlLableBruit
    
    '�����м�¼��
    Call InitRecords
    
    'װ������
    Call SQLCombination
    gstrSQL = mstrSQL
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mintҳ��)
    '�����������¼���ṹ
    Call DataMap_Init(rsTemp)
    '�����ݲ����û����¼���ĸ�ʽ,ͬʱʵ��һ�����ݷ�����ʾ�Ĺ���
    Call PreTendFormat(rsTemp)
    
    lblCurPage.Caption = "P" & mintҳ��
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DataMap_Init(ByVal rsSource As ADODB.Recordset)
    '��ʼ���ڴ����ݼ�
    
    If Not mblnClear Then Exit Sub
    
    '���ݼ�¼��,���ڿ��ٻָ�
    Set mrsDataMap = CopyNewRec(rsSource)
    mrsDataMap.Sort = "ҳ��,�к�"
    '�޸ĵ�Ԫ���¼,���ڱ���
    Call Record_Init(mrsCellMap, "ID," & adLongVarChar & ",50|ҳ��," & adDouble & ",18|�к�," & adDouble & ",18|" & _
            "�к�," & adDouble & ",18|��¼ID," & adDouble & ",18|����," & adLongVarChar & ",4000|��λ," & adLongVarChar & ",100|" & _
            "����," & adDouble & ",1|ɾ��," & adDouble & ",1")
    mrsCellMap.Sort = "ҳ��,�к�,�к�"
    '���Ƽ�¼��
    Set mrsCopyMap = New ADODB.Recordset
    Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
    
    'Ϊ�˲�Ӱ��֮��Ļ�ҳ,���˲�������Ϊ��
    mblnClear = False
End Sub

Private Function DataMap_Save() As Boolean
    '����ǰҳ�����û��༭�������ݱ�������,ҳ���л��򱣴�ǰ����
    Dim blnExit As Boolean
    Dim lngRow As Long, lngCOL As Long, lngRows As Long, lngCols As Long
    On Error GoTo errHand
    
    '�����Ƿ�༭��������
'    '�����ǰҳδ�༭��,�򲻱ر���
'    mrsCellMap.Filter = "ҳ��=" & mintҳ��
'    blnExit = (mrsCellMap.RecordCount = 0)
'    If blnExit Then
'        mrsCellMap.Filter = 0
'        DataMap_Save = True
'        Exit Function
'    End If
'    mrsCellMap.Filter = 0
    If Not CheckFlip Then Exit Function
    
    '��ɾ��ָ��ҳ�ŵ�����������
    mrsDataMap.Filter = "ҳ��=" & mintҳ��
    Do While True
        If mrsDataMap.RecordCount = 0 Then Exit Do
        mrsDataMap.Delete
        mrsDataMap.MoveNext
    Loop
    mrsDataMap.Filter = 0
    
    '����ָ��ҳ�ŵ�����������
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsDataMap.AddNew
        mrsDataMap!ҳ�� = mintҳ��
        mrsDataMap!�к� = lngRow
        mrsDataMap!ɾ�� = IIf(VsfData.RowHidden(lngRow), 1, 0)
        For lngCOL = 0 To lngCols - VsfData.FixedCols
            If lngCOL + VsfData.FixedCols = mlngChoose Then
                mrsDataMap.Fields(cControlFields + lngCOL).Value = VsfData.Cell(flexcpChecked, lngRow, mlngChoose)
            ElseIf InStr(1, "," & mlngCollectType & "," & mlngRecord & ",", "," & lngCOL + VsfData.FixedCols & ",") <> 0 Then
                mrsDataMap.Fields(cControlFields + lngCOL).Value = Val(VsfData.TextMatrix(lngRow, lngCOL + VsfData.FixedCols))
            Else
                mrsDataMap.Fields(cControlFields + lngCOL).Value = IIf(VsfData.TextMatrix(lngRow, lngCOL + VsfData.FixedCols) = "", Null, VsfData.TextMatrix(lngRow, lngCOL + VsfData.FixedCols))
            End If
        Next
        mrsDataMap.Update
    Next
    
    DataMap_Save = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function DataMap_Restore(ByVal rsTemp As ADODB.Recordset) As Boolean
    '��ָ��ҳ������ݻָ��������
    Dim lngRow As Long, lngCOL As Long, lngRows As Long, lngCols As Long
    On Error GoTo errHand
    
    VsfData.Cell(flexcpChecked, VsfData.FixedRows, mlngChoose, VsfData.Rows - 1, mlngChoose) = flexTSUnchecked
    '����ָ��ҳ�ŵ����������е������
    mrsDataMap.Filter = "ҳ��=" & mintҳ��
    lngRows = mrsDataMap.RecordCount
    
    If lngRows = 0 Then
        'û���޸Ĺ���������󶨶�ȡ�ļ�¼��
        mrsDataMap.Filter = 0
        Set VsfData.DataSource = rsTemp
        DataMap_Restore = True
        Exit Function
    Else
        Set VsfData.DataSource = rsTemp
    End If
    
    mrsDataMap.MoveFirst
    lngCols = VsfData.Cols - 1
    For lngRow = 0 To lngRows - 1
        If lngRow > VsfData.Rows - VsfData.FixedRows - 1 Then VsfData.Rows = VsfData.Rows + 1
        For lngCOL = 0 To lngCols - VsfData.FixedCols
            If lngCOL + VsfData.FixedCols = mlngChoose Then
                If InStr(1, "3,4", NVL(mrsDataMap.Fields(cControlFields + lngCOL).Value, 0)) <> 0 Then
                    VsfData.Cell(flexcpChecked, VsfData.FixedRows + lngRow, lngCOL + VsfData.FixedCols) = NVL(mrsDataMap.Fields(cControlFields + lngCOL).Value)
                End If
            Else
                VsfData.TextMatrix(VsfData.FixedRows + lngRow, lngCOL + VsfData.FixedCols) = NVL(mrsDataMap.Fields(cControlFields + lngCOL).Value)
            End If
        Next
        If mrsDataMap!ɾ�� = 1 Then VsfData.RowHidden(VsfData.FixedRows + lngRow) = True
        
        mrsDataMap.MoveNext
    Next
    
    mrsDataMap.Filter = 0
    DataMap_Restore = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CellMap_Update(ByVal lngStart As Long, ByVal lngDeff As Long)
    Dim lngPos As Long
    Dim intCol As Integer
    
    '���µ�ǰҳ�����д�����ʼ�е��к�����
    With mrsCellMap
        If .RecordCount <> 0 Then .MoveLast
        If .BOF Then Exit Sub
        Do While Not .BOF
            If !ҳ�� = mintҳ�� And !�к� > lngStart Then
                intCol = !�к�
                lngPos = .AbsolutePosition
                !�к� = !�к� + lngDeff
                !ID = mintҳ�� & "," & !�к� & "," & !�к�
                .Update
                .MoveFirst
                .Move lngPos - 2
            Else
                .MovePrevious
            End If
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional ByVal blnAddPage As Boolean = True) As ADODB.Recordset
    'ֻ������¼���Ľṹ,ͬʱ����ҳ��,�к��ֶ�
    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer
    
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        If blnAddPage Then
            .Fields.Append "ҳ��", adDouble, 18
            .Fields.Append "�к�", adDouble, 18
        End If
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Name = "��������" Then
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, 50, adFldIsNullable      '0:��ʾ����
            ElseIf rsSource.Fields(intFields).Type = 200 Then       '�����ʹ���Ϊ�ַ���
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:��ʾ����
            Else
                .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:��ʾ����
            End If
        Next
        If blnAddPage Then
            .Fields.Append "ɾ��", adDouble, 1
        End If
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set CopyNewRec = rsTarget
End Function

Private Sub PreTendMutilRows()
    Dim lngRowCount As Long, lngRowCurrent As Long  '��ǰ��¼������,��ǰ��¼�ڱ�ҳ��ʵ������
    Dim lngCOL As Long, lngMax As Long
    Dim lngRow As Long
    On Error GoTo errHand
    
    Dim arrData
    Dim intData As Integer, intDatas As Integer
    '���һ����ʾ�����������ʾ(���ݵ�ǰ����ռ����������ӿհ��в�����������,Ȼ�������δ���ǰ�е�����)
    'ÿҳֻ��ʾʵ�ʵ�������,��'@��ȡ��ע�ͼ���
    
    lngRow = VsfData.FixedRows
    Do While True
        If lngRow > VsfData.Rows - 1 Then Exit Do
        If lngRow >= mlngPageRows + mlngOverrunRows + VsfData.FixedRows Then Exit Do
        If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") <> 0 Then Exit Do
        lngRowCount = Val(VsfData.TextMatrix(lngRow, mlngRowCount))
        '@ʵ��������
'        lngRowCurrent = Val(VsfData.TextMatrix(lngRow, mlngRowCurrent))
        
        If lngRowCount > 1 Then
            '�����ӿ���
            VsfData.Rows = VsfData.Rows + lngRowCount - 1
            '�ӵ�ǰ�е���һ�п�ʼ��ÿ�е�λ��+�����ӵĿհ���������֤�����Ŀհ��дӵ�ǰ�е���һ�п�ʼ
            For intData = VsfData.Rows - lngRowCount To lngRow + 1 Step -1
                VsfData.RowPosition(intData) = intData + lngRowCount - 1
            Next
            
            'ѭ������ǰ������
            For lngCOL = 0 To VsfData.Cols - 1
                If VsfData.ColHidden(lngCOL) And lngCOL <> mlngRowCount Then
                    'ѭ����ֵ
                    For intData = 2 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, lngCOL) = VsfData.TextMatrix(lngRow, lngCOL)
                    Next
                ElseIf (lngCOL < mlngNoEditor And lngCOL <> mlngDate And lngCOL <> mlngTime) Then
                    '׼����ֵ
                    With txtLength
                        .Width = VsfData.ColWidth(lngCOL)
                        .Text = VsfData.TextMatrix(lngRow, lngCOL)
                        .FontName = VsfData.CellFontName
                        .FontSize = VsfData.CellFontSize
                    End With
                    arrData = GetData(txtLength.Text)
                    intDatas = UBound(arrData)
                    
                    If intDatas > 0 Then
                        'ѭ����ֵ
                        For intData = 0 To intDatas
                            If VsfData.Rows <= lngRow + intData Then VsfData.Rows = VsfData.Rows + 1
                            VsfData.TextMatrix(lngRow + intData, lngCOL) = arrData(intData)
                        Next
                    End If
                ElseIf lngCOL = mlngNoEditor Then
                        '����ֵ��Ϊ��1��ʼ,������4������,����4|1
                        For intData = 1 To lngRowCount
                            VsfData.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngRowCount & "|" & intData
                        Next
                    Else
                End If
            Next
            '@ʵ��������
'            '�����ҳ��һ�е����ݲ�ȫ,���Ƚ��ü�¼��һ�е�������(����,ʱ��,ǩ��)��Ϣ���Ƶ�
'            If lngRow = VsfData.FixedRows And lngRowCount <> lngRowCurrent Then
'                '�̶�������ʾ����ʱ����ǩ����
'                lngMax = lngRowCount - lngRowCurrent
'                If mlngDate > -1 Then VsfData.TextMatrix(lngRow + lngMax, mlngDate) = VsfData.TextMatrix(lngRow, mlngDate)
'                If mlngTime > -1 Then VsfData.TextMatrix(lngRow + lngMax, mlngTime) = VsfData.TextMatrix(lngRow, mlngTime)
'                if mlngOperator <>-1 then VsfData.TextMatrix(lngRow + lngMax, mlngOperator) = VsfData.TextMatrix(lngRow, mlngOperator)
'                if mlngSignName <>-1 then VsfData.TextMatrix(lngRow + lngMax, mlngsignname) = VsfData.TextMatrix(lngRow, mlngsignname)
'                'ɾ���������
'                For lngCol = 1 To lngMax
'                    VsfData.RemoveItem lngRow
'                Next
'            End If
'            lngRow = lngRow + lngRowCurrent - 1 '���ϸü�¼�ڱ�ҳʵ�ʵ�����
            '@ʵ��������Ҫע���������д���
            lngRow = lngRow + lngRowCount - 1
        Else
            VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
        End If
        lngRow = lngRow + 1
    Loop
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreTendFormat(ByVal rsTemp As ADODB.Recordset)
    Dim aryItem() As String
    Dim lngRow As Long, lngCOL As Long, lngCount As Long, strCell As String
    On Error GoTo errHand
    
    '���û����¼���ĸ�ʽ
    With VsfData
        .Redraw = flexRDNone
        .Clear
        Call DataMap_Restore(rsTemp)
        
        '��ͷ��д
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        '�����ڲ�����������
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(mlngChoose) = Not mblnVerify
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        .ColHidden(mlngCollectStyle) = True
        .ColHidden(mlngCollectText) = True
        .ColHidden(mlngCollectType) = True
        .ColHidden(mlngCollectDay) = True
        .ColWidth(0) = 250
        .ColWidth(mlngChoose) = 250      'ѡ����
        
        .FrozenCols = mlngTime
        .SheetBorder = &H40C0&
        
        '������ͷ
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCOL = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCOL + cHideCols + .FixedCols - 1) = strCell
        Next
        '���ù̶��м�ѡ����
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        .TextMatrix(2, 0) = " "
        .TextMatrix(0, mlngChoose) = " "
        .TextMatrix(1, mlngChoose) = " "
        .TextMatrix(2, mlngChoose) = " "
        Call PreActiveHead
        
        '�п�����
        Dim blnAlign As Boolean
        aryItem = Split(mstrColWidth, ",")
        For lngCount = cHideCols + .FixedCols To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = BlowUp(Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(0)))
                If InStr(1, aryItem(lngCount - cHideCols - .FixedCols), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(1))
                End If
            End If
        Next
        
        '�̶��и�ʽΪ����
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '�ٰ��кϲ�
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        If blnAlign = False Then
            '��Ϊ�����û���������ʾ�ж��뷽ʽ
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .RowHeight(lngCount) <> .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
        Case 3
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        If .Rows = .FixedRows Then
            mlngOverrunRows = 0
        Else
            '�õ���һ�еĳ�����
            mlngOverrunRows = Val(.TextMatrix(3, mlngRowCount)) - Val(.TextMatrix(3, mlngRowCurrent))
            '�������һ�еĳ�����
            mlngOverrunRows = mlngOverrunRows + Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent))
        End If
        
        Call PreTendMutilRows
        Call FillPage
        
        Call WriteColor
        .Redraw = flexRDDirect
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub WriteColor()
    Dim blnTag As Boolean
    Dim lngCount As Long
    '����Ժ�ɫ��ʾ��ͬʱ������ʼ������ΪNoCheckBox������ͼ��
    With VsfData
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, 1) <> "" Then
                '����Ժ�ɫ��ʾ
                blnTag = False
                If mintTagFormHour < mintTagToHour Then
                    blnTag = (Hour(.TextMatrix(lngCount, 1)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 1)) < mintTagToHour)
                Else
                    blnTag = (Hour(.TextMatrix(lngCount, 1)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 1)) < mintTagToHour)
                End If
                If blnTag Then
                    Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFont
                    .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
                End If
            End If
            
            '������ʼ������ΪNoCheckBox
            If Not VsfData.RowHidden(lngCount) Then
                If Not VsfData.TextMatrix(lngCount, mlngRowCount) Like "*|1" Then
                    VsfData.Cell(flexcpChecked, lngCount, mlngChoose) = flexNoCheckbox
                Else
                    If VsfData.Cell(flexcpChecked, lngCount, mlngChoose) <> flexTSChecked Then
                        VsfData.Cell(flexcpChecked, lngCount, mlngChoose) = flexTSUnchecked
                    End If
                    
                    '����ͼ��
                    If VsfData.TextMatrix(lngCount, mlngSigner) = "" Then
                        VsfData.Cell(flexcpPicture, lngCount, 0) = Nothing
                    Else
                        If InStr(1, VsfData.TextMatrix(lngCount, mlngSigner), "/") <> 0 Then
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(��ǩ).Picture
                        Else
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(ǩ��).Picture
                        End If
                    End If
                
                    '����С�����ʾ
                    If Val(VsfData.TextMatrix(lngCount, mlngCollectType)) <> 0 Then
                        VsfData.TextMatrix(lngCount, mlngDate) = VsfData.TextMatrix(lngCount, mlngCollectText)
                        VsfData.TextMatrix(lngCount, mlngTime) = VsfData.TextMatrix(lngCount, mlngCollectText)
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub zlLableBruit()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    
    lblSubhead.Caption = lblSubhead.Tag
    lblSubhead.Top = lblTitle.Top + lblTitle.Height + 120
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    VsfData.Move lngScaleLeft + 210, lblSubhead.Top + lblSubhead.Height + 45, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Height = picMain.Height - VsfData.Top
End Sub

Private Sub GetFileProperty()
    '��ȡ�ļ�����
    On Error GoTo errHand
    
    gstrSQL = " Select  /*+ RULE */ ��ʼʱ��,����ʱ��,��ʽID,����ID,�鵵�� From ���˻����ļ� " & _
              " Where ����ID=[1] And ��ҳID=[2] And Ӥ��=[3] And ID=[4] And Rownum<2"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�����", mlng����ID, mlng��ҳID, mintӤ��, mlng�ļ�ID)
    If rsTemp.RecordCount <> 0 Then
        mlng��ʽID = rsTemp!��ʽID
        mlng����ID = rsTemp!����ID
        mblnArchive = (NVL(rsTemp!�鵵��) <> "")
        mstr��ʼʱ�� = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss")
        mstr����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    End If
    
    '���ҳ��=-1,˵��ȱʡ��ʾ���һҳ
    If mintҳ�� = -1 Then
        gstrSQL = " Select  /*+ RULE */ MAX(����ҳ��) AS ҳ�� From ���˻����ӡ Where �ļ�ID=[1]"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ��ҳ������ݷ���ʱ�䷶Χ", mlng�ļ�ID)
        mintҳ�� = NVL(rsTemp!ҳ��, 1)
        mint����ҳ = mintҳ��
    End If
    
    RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitEnv()
    Dim Rs As New ADODB.Recordset
    On Error GoTo errHand
    
    glngHours = Val(zlDatabase.GetPara("���ݲ�¼ʱ��", glngSys))
    '���ִ��ڵ����л����¼��Ŀ
    gstrSQL = " Select  /*+ RULE */ ��Ŀ���,��Ŀ����,��Ŀ����,��Ŀ����,��Ŀ����,��ĿС��,��Ŀ��ʾ,��Ŀ��λ,��Ŀֵ��,����ȼ�,Ӧ�÷�ʽ" & _
              " From �����¼��Ŀ B" & _
              " Where B.Ӧ�÷�ʽ<>0 " & _
              " Order by ��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "���ִ��ڵ����л����¼��Ŀ")
    
    'ȡ��ǰ����Ա�ļ���
    mintVerify = δ����
    mintVerify_Last = δ����
    gstrSQL = "select /*+ RULE */ Ƹ�μ���ְ�� from ��Ա�� p where ID=[1]"
    Set Rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", glngUserId)
    If Not Rs.EOF Then
        mintVerify = NVL(Rs("Ƹ�μ���ְ��"), δ����)
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitRecords()
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim lngCOL As Long, lngOrder As Long, strName As String, intImmovable As Integer, strFormat As String
    Dim arrColumn, arrItem, strColumns As String
    Dim blnSet As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    strColumns = mstrColumns
    If Not mblnInit Then
        '��ʼ���ڴ��¼��(δ��Ӧ��Ŀ����Ϊ���Ŀ,�����о�Ϊ�̶���)
        strFields = "��," & adDouble & ",18|���," & adDouble & ",2|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",20|�̶�," & adDouble & ",2|��ʽ," & adLongVarChar & ",2000"
        Call Record_Init(mrsSelItems, strFields)
        strFields = "��|���|��Ŀ���|��Ŀ����|�̶�|��ʽ"
    End If
    
    '�����ж���
    If Not mblnInit Then
        arrColumn = Split(strColumns, "|")
        j = UBound(arrColumn)
        For i = 0 To j
            lngCOL = Split(arrColumn(i), "'")(0)
            arrItem = Split(Split(arrColumn(i), "'")(1), ",")
            blnSet = False   '����������Դ���ֵΪ׼'�����Ҳ�����Ŀ���ǻ��Ŀ
            If UBound(Split(arrColumn(i), "'")) > 1 Then
                blnSet = True
                intImmovable = Split(arrColumn(i), "'")(2)
            End If
            If UBound(Split(arrColumn(i), "'")) > 2 Then
                strFormat = Split(arrColumn(i), "'")(3)
            End If
            
            k = UBound(arrItem)
            For l = 0 To k
                strName = arrItem(l)
                mrsItems.Filter = "��Ŀ����='" & strName & "'"
                If mrsItems.RecordCount <> 0 Then
                    lngOrder = mrsItems!��Ŀ���
                    If Not blnSet Then intImmovable = 1   '�̶��������޸�
                Else
                    lngOrder = 0
                    If Not blnSet Then intImmovable = 0
                    
                    '��¼������
                    Select Case strName
                    Case "����"
                        mlngDate = i + cHideCols + VsfData.FixedCols
                    Case "ʱ��"
                        mlngTime = i + cHideCols + VsfData.FixedCols
                    Case "��ʿ"
                        mlngOperator = i + cHideCols + VsfData.FixedCols
                    Case "ǩ����"
                        mlngSignName = i + cHideCols + VsfData.FixedCols
                    Case "ǩ��ʱ��"
                        mlngSignTime = i + cHideCols + VsfData.FixedCols
                    End Select
                End If
                strValues = lngCOL & "|" & l + 1 & "|" & lngOrder & "|" & strName & "|" & intImmovable & "|" & strFormat
                Call Record_Add(mrsSelItems, strFields, strValues)
            Next
        Next
        
        'Call OutputRsData(mrsSelItems)
        
        '��������ڲ�������(�����ڶ�ȡ���ݺ��ʱ���ӵ�,��ʱֻ��Ԥ������)
        mlngChoose = 2 + VsfData.FixedCols
        mlngSignLevel = VsfData.Cols + cHideCols + VsfData.FixedCols '����������
        mlngSigner = mlngSignLevel + 1
        mlngRecord = mlngSigner + 1
        mlngRowCount = mlngRecord + 1
        mlngRowCurrent = mlngRowCount + 1
        mlngCollectType = mlngRowCurrent + 1
        mlngCollectText = mlngCollectType + 1
        mlngCollectStyle = mlngCollectText + 1
        mlngCollectDay = mlngCollectStyle + 1
        If mlngOperator <> -1 And mlngSignName <> -1 Then
            mlngNoEditor = IIf(mlngOperator < mlngSignName, mlngOperator, mlngSignName)
        Else
            mlngNoEditor = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        End If
    End If
    
    mrsItems.Filter = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub ArchiveMe()
    On Error GoTo errHand
    
    If mlng����ID = 0 Or gblnMoved Then Exit Sub
    If MsgBox("��Ҫ���ò��˱���סԺ���л����ļ��鵵��", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
        Dim strNow As String

        strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        gstrSQL = "ZL_���˻����ļ�_ARCHIVE(" & mlng����ID & "," & mlng��ҳID & "," & mintӤ�� & ",1)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�鵵")

        mblnArchive = True
        RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub UnArchiveMe()
    On Error GoTo errHand
    
    If mlng����ID = 0 Or gblnMoved Then Exit Sub
    If MsgBox("��Ҫȡ���ò��˵Ĺ鵵״̬��", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then

        gstrSQL = "ZL_���˻����ļ�_ARCHIVE(" & mlng����ID & "," & mlng��ҳID & "," & mintӤ�� & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�����鵵")
        
        mblnArchive = False
        RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function SignMe(Optional ByVal bln��ǩ As Boolean = False) As Boolean
    Dim blnSign As Boolean          '�Ƿ�ǩ���ɹ�
    Dim blnRefresh As Boolean
    Dim strSignTime As String       '��֤����ǩ����ǩ��ʱ��һ��,����ȡ��ǩ��ʱ��ǩ��ʱ��ͳһȡ��
    Dim str״̬ As String           '����ǩ��ѡ��,����ѭ��ǩ��ʱ��ͣ�ĵ���ǩ������
    Dim str�д��� As String
    Dim str���� As String
    Dim intRow As Integer, intRows As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '������ʱ��ѭ��������δǩ�����ݽ���ǩ��
    
    If mlng����ID = 0 Then Exit Function
    
    '��ǩ:������δǩ�������ݽ���ǩ��
    '��ǩ:��������ǩ�������ݽ�����ǩ
    If bln��ǩ Then
        If Not mblnVerify Then
            gstrSQL = " Select /*+ RULE */ distinct B.����ʱ�� " & vbNewLine & _
                      " From ���˻�����ϸ A,���˻������� B,���˻����ļ� C" & vbNewLine & _
                      " Where A.��¼ID=B.ID And B.�ļ�ID=C.ID And B.�������=0 And A.������Դ=0 And MOD(A.��¼����,10)=5 AND A.��ֹ�汾 Is NULL And C.ID=[1] "
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID)
            If rsTemp.RecordCount = 0 Then
                RaiseEvent AfterRowColChange("��������ǩ�������ݣ�", True, mblnSign, mblnArchive)
                Exit Function
            End If
        
            '������ǩģʽ,���޸�����,�ɹ�ѡ����
            mblnVerify = True
            chkSwitch.Visible = mblnVerify
            VsfData.ColHidden(mlngChoose) = Not mblnVerify
            VsfData.Cell(flexcpChecked, VsfData.FixedRows, mlngChoose, VsfData.Rows - 1, mlngChoose) = flexTSUnchecked
            Call WriteColor
            RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
            Exit Function
        Else
            '��ȡ����ǩ������
            gstrSQL = " Select /*+ RULE */ distinct B.����ʱ�� " & vbNewLine & _
                      " From ���˻�����ϸ A,���˻������� B,���˻����ļ� C,(SELECT COLUMN_VALUE FROM TABLE(CAST(F_NUM2LIST([2]) AS ZLTOOLS.T_NUMLIST))) G" & vbNewLine & _
                      " Where A.��¼ID=B.ID And B.ID=G.COLUMN_VALUE And B.�ļ�ID=C.ID And B.�������=0 And MOD(A.��¼����,10)=5 AND A.��ֹ�汾 Is NULL And C.ID=[1] "
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID, mstrVerify)
        End If
    Else
        '���Ա����޸ĵ����ݽ���ǩ��(��ȡδǩ������-��ǩ������)
        gstrSQL = "" & _
                "SELECT /*+ RULE */ DISTINCT B.����ʱ��" & vbNewLine & _
                "FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                "WHERE A.��¼ID=B.ID And B.�������=0 And A.������Դ=0 AND A.��ֹ�汾 IS NULL AND A.��¼���� =1 AND instr(NVL(B.ǩ����,'QMR'),'/',1)=0 AND A.��¼��=[2] AND B.�ļ�ID=[1]" & vbNewLine & _
                "MINUS" & vbNewLine & _
                "SELECT DISTINCT B.����ʱ��" & vbNewLine & _
                "FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                "WHERE A.��¼ID=B.ID And B.�������=0 And A.������Դ=0 AND A.��ֹ�汾 IS NULL AND A.��¼���� =5 AND A.��¼��=[2] AND B.�ļ�ID=[1]"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID, gstrUserName)
        If rsTemp.RecordCount = 0 Then
            RaiseEvent AfterRowColChange("û���ҵ���Ҫǩ�������ݣ�ֻ�ܶ��Լ��Ǽǻ��޸ĵ����ݽ���ǩ������", True, mblnSign, mblnArchive)
            Exit Function
        End If
    End If
    
    '׼��ǩ��
    strSignTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    With rsTemp
        Do While Not .EOF
            str�д��� = ""
            blnSign = SignName(Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss"), strSignTime, bln��ǩ, str״̬, str�д���)
            If Not blnSign Then Exit Do
            If Not blnRefresh Then blnRefresh = blnSign
            If str�д��� <> "" Then
                str���� = str���� & vbCrLf & "����ʱ��=[" & Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "]" & str�д���
            End If
            .MoveNext
        Loop
    End With
    
    
    If blnRefresh And Not mblnVerify Then Call ShowMe(mfrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, mintҳ��)
'    If str���� <> "" Then MsgBox "ǩ��ʱ�������´���" & str����, vbInformation, gstrSysName
    SignMe = blnRefresh
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub UnSignMe(Optional ByVal bln��ǩ As Boolean = False)
    Dim intPos As Integer
    Dim lngStart As Long                '��ʼ��
    Dim lngRecord As Long
    Dim blnOK As Boolean
    Dim strSignTime As String           'ǩ��ʱ��
    Dim blnClear As Boolean             'ȡ��ǩ��ʱ�Ƿ�����ð汾�����ݻ��˵��ϴ�ǩ�����״̬
    Dim blnTrans As Boolean
    
    Dim clsSign As Object
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '�������һ���Ǳ��˵�ǩ�������ݵ�ǰѡ�����ݵ�ǩ��ʱ�䣬����ȡ��ǩ��
    
    If mlng����ID = 0 Then Exit Sub
    
    '��Ҫ�Լ��
    '��ǰ��¼���¼�¼���˳�
    lngStart = GetStartRow(VsfData.ROW)
    lngRecord = Val(VsfData.TextMatrix(lngStart, mlngRecord))
    If lngRecord = 0 Then
        RaiseEvent AfterRowColChange("������¼������ȡ��ǩ����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '��ǰ��¼δǩ�����˳�
    If VsfData.TextMatrix(lngStart, mlngSigner) = "" Then
        RaiseEvent AfterRowColChange("��ǰ��¼��δǩ����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '��ǩ����ǰ��¼δ��ǩ���˳���ƽǩ����ǰ��¼����ǩ���˳�
    intPos = InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/")
    If bln��ǩ Then
        If intPos = 0 Then
            RaiseEvent AfterRowColChange("��ǰ��¼δ��ǩ���޷�ִ��ȡ����ǩ������", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    Else
        If intPos <> 0 Then
            RaiseEvent AfterRowColChange("��ǰ��¼����ǩ����ȡ����ǩ���ٲ�����", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    End If
    '��ǰ��¼�����ǩ���˲��Ǳ������˳�
    gstrSQL = "" & _
              " SELECT /*+ RULE */ A.��¼��,A.��Ŀ���� AS ǩ��ʱ��" & vbNewLine & _
              " FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
              " WHERE A.��¼ID=B.ID And B.�������=0 AND B.�ļ�ID=[1] AND A.��¼ID=[2] AND A.��¼����=" & IIf(bln��ǩ, 15, 5) & vbNewLine & _
              " ORDER BY A.��Ŀ���� DESC"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ǰ��¼�����ǩ���˲��Ǳ������˳�", mlng�ļ�ID, lngRecord)
    If rsTemp!��¼�� <> gstrUserName Then
        RaiseEvent AfterRowColChange("���������ǩ���ˣ�����ִ�б�������", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '��ȡ��������׼��ȡ��ǩ������ǩ
    gstrSQL = "" & _
              " SELECT /*+ RULE */ A.��ĿID AS ֤��ID,B.����ʱ��" & vbNewLine & _
              " FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
              " WHERE A.��¼ID=B.ID And B.�������=0 AND B.�ļ�ID=[1] And A.��¼��=[2] And A.��Ŀ����=[3] " & _
              " AND A.��¼����=" & IIf(bln��ǩ, 15, 5) & vbNewLine & _
              " ORDER BY A.��Ŀ���� DESC"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������׼��ȡ��ǩ������ǩ", mlng�ļ�ID, gstrUserName, CStr(rsTemp!ǩ��ʱ��))
    
    'ѯ���Ƿ���Ҫ�������
    If Not bln��ǩ Then
        blnClear = (MsgBox("ȡ��ǩ��ʱ�Ƿ�ð汾�����ݻ��˵��ϴ�ǩ�����״̬��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
    End If
    
    gcnOracle.BeginTrans
    blnTrans = True
    Do While Not rsTemp.EOF
        If NVL(rsTemp!֤��ID, 0) > 0 Then
            '����ǩ����֤��ֻ��֤һ��
            Err.Clear
            On Error Resume Next
            If clsSign Is Nothing Then
                Set clsSign = CreateObject("zl9ESign.clsESign")
                If Err <> 0 Then Err = 0
                
                If Not clsSign Is Nothing Then
                    If clsSign.Initialize(gcnOracle, glngSys) Then
                        If Not clsSign.CheckCertificate(gstrDBUser) Then
                            gcnOracle.RollbackTrans
                            Exit Sub
                        End If
                    Else
                        gcnOracle.RollbackTrans
                        RaiseEvent AfterRowColChange("ȡ��ǩ��ʱ��Ҫ�ٴ���֤����ϵͳû������ǩ����֤���ģ�����ȡ����", True, mblnSign, mblnArchive)
                        Exit Sub
                    End If
                Else
                    gcnOracle.RollbackTrans
                    RaiseEvent AfterRowColChange("ǩ��������ʼ��ʧ�ܣ�", True, mblnSign, mblnArchive)
                    Exit Sub
                End If
            End If
        End If
        
        'ȡ��ǩ��
        gstrSQL = "ZL_���˻�������_UNSIGNNAME("
        gstrSQL = gstrSQL & mlng�ļ�ID & ","
        gstrSQL = gstrSQL & "To_Date('" & Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & IIf(blnClear, "1", "0") & ")"
        
        Debug.Print gstrSQL
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ִ��ȡ��ǩ��")
        
        rsTemp.MoveNext
    Loop
    gcnOracle.CommitTrans
    blnTrans = False
    
    'ˢ������
    Call ShowMe(mfrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, mintҳ��)
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SignName(ByVal strStart As String, ByVal strSignTime As String, ByVal bln��ǩ As Boolean, _
    str״̬ As String, Optional str���� As String) As Boolean
    '******************************************************************************************************************
    '����:
    '
    '
    '******************************************************************************************************************
    Dim oSign As cEPRSign
    Dim strSource As String             '��ǩԴ���ݴ�
    Dim lngLoop As Long
    Dim Rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    strSource = ""
    
    '��ȡҪǩ��������
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = " Select /*+ RULE */ a.id,a.��¼id,a.��¼����,a.��Ŀ����,a.��Ŀid,a.��Ŀ���,a.��Ŀ����,a.��Ŀ����,a.��¼����,a.��Ŀ��λ, " & _
              "     a.��¼���,a.���²�λ,a.��¼���,a.���Ժϸ�,a.δ��˵��,a.��ʼ�汾,a.��ֹ�汾,a.��¼��,a.��¼ʱ��  " & _
              " From ���˻�����ϸ a,���˻������� b,���˻����ļ� c " & _
              " Where a.��¼id=b.ID And B.�������=0 And b.�ļ�ID=c.ID And a.��ֹ�汾 Is Null And C.ID=[1] And b.����ʱ��=[2]"
    Call SQLDIY(gstrSQL)
    Set Rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҪǩ��������", mlng�ļ�ID, CDate(strStart))
    If Rs.BOF = False Then
        Do While Not Rs.EOF
            For lngLoop = 0 To Rs.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(Rs.Fields(lngLoop).Value, ""))
            Next
            Rs.MoveNext
        Loop
    End If
    If strSource = "" Then
        RaiseEvent AfterRowColChange("��ǰû����Ҫǩ������Ϣ��", True, mblnSign, mblnArchive)
        Exit Function
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Err = 0
    Set oSign = frmTendFileSign.ShowMe(Me, mstrPrivs, mlng�ļ�ID, mintVerify_Last, strSource, bln��ǩ, str״̬, str����)
    On Error GoTo errHand
    
    If Not oSign Is Nothing Then
        gstrSQL = "ZL_���˻�������_SIGNNAME("
        gstrSQL = gstrSQL & mlng�ļ�ID & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss')," & IIf(bln��ǩ, 1, 0) & ","
        gstrSQL = gstrSQL & "'" & oSign.���� & "',"
        gstrSQL = gstrSQL & "'" & oSign.ǩ����Ϣ & "'," & oSign.ǩ������ & ","
        gstrSQL = gstrSQL & oSign.֤��ID & ","
        gstrSQL = gstrSQL & oSign.ǩ����ʽ & ",'" & strSignTime & "')"
        
        Debug.Print gstrSQL
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ִ��ǩ��")
        SignName = True
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CancelMe() As Boolean
    CancelMe = True
    mblnVerify = False
    mblnChange = False
    Call ShowMe(mfrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, mintҳ��)
End Function

Public Function SaveME() As Boolean
    If Not CheckData Then Exit Function
    If Not SaveData Then Exit Function
    
    mblnShow = False
    Call InitCons
    SaveME = True
    
    Call ShowMe(mfrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, mintҳ��)
End Function

Public Function ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, _
    ByVal lngDeptID As Long, ByVal intBaby As Integer, Optional ByVal strPrivs As String, Optional ByVal blnEditable As Boolean = True, _
    Optional ByVal intҳ�� As Integer = -1, Optional ByVal blnClear As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ʾ�����¼�ļ�����
    '������ frmParent           �ϼ��������
    '       lngPatiID           ����id
    '       lngPageID           ��ҳid
    '       lngDeptID           Ҫ��ʾ�����¼�Ŀ���
    '       intBaby             Ӥ����־
    '       blnEditable         ���Ϊ��,˵������Ϊ��ѯ�Ӵ�����ʹ��,ȡ����༭��صĹ���
    '       blnClear            ���Ϊ��,���mrsDataMap��¼��;����ҳʱӦ����,�����û��޸ĵ������Ա���ʾ������ʹ��
    '���أ� ��
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    Err = 0
    
    mblnInit = False
    lngLastRow = VsfData.ROW
    lngLastTopRow = VsfData.TopRow
    lngLastPatientID = mlng����ID
    If lngLastRow < 1 Then lngLastRow = 1
    If lngLastTopRow < 1 Then lngLastTopRow = 1
    
    If mblnChange Then
        If MsgBox("��ǰ���˵����ݻ�δ���棬�㡰�ǡ����б��棬�㡰�񡱽����������޸ģ�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call VsfData_EnterCell
            Call SaveData
        End If
    End If
    
    mblnClear = blnClear
    mintҳ�� = intҳ��
    mlng�ļ�ID = lngFileID
    mlng����ID = lngPatiID
    mlng��ҳID = lngPageId
    mlng����ID = lngDeptID
    mintӤ�� = intBaby
    mstrPrivs = strPrivs
    mblnBlowup = (zlDatabase.GetPara("�����ļ���ʾģʽ", glngSys, 1255, 0) = 1)
    Set mfrmParent = frmParent
    
    mintPreDays = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1"))
    mstrMaxDate = Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd")
    
    If mrsItems.State = 0 Then
        Call InitMenuBar
        Call InitEnv            '��ʼ������
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
    End If
    
    Call InitVariable
    Call InitCons
    If Not ReadStruDef Then Exit Function
    Call zlRefresh
    mblnInit = True
    mblnEditable = blnEditable And Not gblnMoved And Not mblnArchive
    
    '�ָ���λ
    If lngLastPatientID <> mlng����ID Then
        lngLastRow = 1
        lngLastTopRow = 1
    End If
    
    If VsfData.Rows - 1 > lngLastRow Then VsfData.ROW = lngLastRow
    If VsfData.RowIsVisible(VsfData.ROW) Then VsfData.TopRow = lngLastTopRow
    
    RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    RaiseEvent AfterRefresh
    
'    Call OutputRsData(mrsSelItems)
    ShowMe = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckFlip() As Boolean
    Dim blnExit As Boolean
    Dim lngRow As Long, lngCOL As Long, lngRows As Long, lngCols As Long
    'ҳ���л�ǰ��飺����ʱ����ȷ����������������ڱ���ʱ�Ͳ����ټ������ҳ��������ˣ�����������¼��ʱ�Ѿ������˼�飬�˴��Թ���
    
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsCellMap.Filter = "ҳ��=" & mintҳ�� & " And �к�=" & lngRow & " And �к�>" & mlngTime
        If mrsCellMap.RecordCount <> 0 Then
            If Not VsfData.RowHidden(lngRow) Then
                blnExit = (VsfData.TextMatrix(lngRow, mlngDate) = "" Or VsfData.TextMatrix(lngRow, mlngTime) = "")
                If blnExit Then
                    mrsCellMap.Filter = 0
                    VsfData.ROW = lngRow
                    If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                    RaiseEvent AfterRowColChange("�벹������ʱ�䣡", True, mblnSign, mblnArchive)
                    CheckFlip = False
                    Exit Function
                End If
            End If
        End If
    Next
    
    mrsCellMap.Filter = 0
    CheckFlip = True
End Function

Private Function CheckData() As Boolean
    Dim intLevel As Integer
    Dim lngPage As Long
    On Error GoTo errHand
    '�������
    
    '����޸������ݶ�����ʱ�䲻ȫ����ʾ�����ݺϷ�����¼��ʱ�Ѿ���飩
    If Not DataMap_Save Then Exit Function
'    Call OutputRsData(mrsCellMap)
'    Call OutputRsData(mrsDataMap)
    
    '�������ǩģʽ,������ѡ�����Ƿ���ڲ�����ǩ�����
    If mblnVerify Then
        mstrVerify = ""
        '��ǩ��������������
        For lngPage = 1 To mint����ҳ
            mrsDataMap.Filter = "ҳ��=" & lngPage
            Do While Not mrsDataMap.EOF
                If NVL(mrsDataMap!ѡ��, 0) = flexTSChecked Then
                    mstrVerify = mstrVerify & "," & mrsDataMap!��¼ID
                    
                    If IsNull(mrsDataMap!ǩ������) Then
                        intLevel = NVL(mrsDataMap!ǩ������, δ����)
                    Else
                        intLevel = Val(mrsDataMap!ǩ������) + 1
                    End If
                    If mintVerify_Last > intLevel Then mintVerify_Last = intLevel
                End If
                mrsDataMap.MoveNext
            Loop
        Next
        mrsDataMap.Filter = 0
        
        If mstrVerify = "" Then
            RaiseEvent AfterRowColChange("����Ҫѡ��һ�����ݲ��������ǩ������", True, mblnSign, mblnArchive)
            Exit Function
        End If
        mstrVerify = Mid(mstrVerify, 2)
    End If
    
    CheckData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    Dim arrValue, arrOrder, arrPart, arrCollect
    Dim strSQL() As String
    Dim intAllow As Integer
    Dim lngRecord As Long
    Dim blnTrans As Boolean, blnSaved As Boolean, blnDel As Boolean
    Dim intPos As Integer, intMax As Integer, intPage As Integer, intRow As Integer, intUsedRows As Integer
    Dim strReturn As String, strCellData As String, strPart As String
    Dim strMonth As String, strDay As String
    Dim strDate As String, strTime As String, strTemp As String
    Dim strDatetime As String, strCurrDate As String, strDays As String
    
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    'ͬ�ж���ѭ�����ã�ZL_���˻�������_UPDATE
    '��һ��ǰ���ã�
    '   1��ZL_���˻�������_SYNCHRO��ͬ�����ݵ����µ��뻤���¼���У���Ҫ��¼ɾ������ϸID��
    '   2��ZL_���˻����ӡ_UPDATE����ɴ�ӡ���ݽ���
    'ɾ����Ŀ���¼��ɾ����Ҳ��Ҫ��¼
    '�޸����ݵ�ͬ���ͽ��������ݶ�Ӧ��������ʱ�䱣�浽mrsCellMap��
    
'    objStream.WriteLine (Now & "��������SQL")
    intAllow = IIf(InStr(mstrPrivs, "���˻����¼") > 0, 1, 0)
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    With mrsCellMap
        '����Ч���ݹ��˳���:��¼ID>0����ʷ����+��������Ч����
        .Filter = "��¼ID>0 or (��¼ID=0 And ɾ��=0)"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If intRow <> !�к� Then
endWork:
                If intRow > 0 Then
                    mrsDataMap.Filter = "ҳ��=" & intPage & " And �к�=" & intRow
                    If mrsDataMap.RecordCount <> 0 Then
                        blnDel = (mrsDataMap!ɾ�� = 1)
                        intUsedRows = Val(Split(NVL(mrsDataMap!���� & "|"), "|")(0))
                    End If
                    mrsDataMap.Filter = 0
                End If

                If blnSaved Then
                    '��ɴ�ӡ���ݽ���
'                    �ļ�ID_IN IN ���˻����ӡ.�ļ�ID%TYPE,
'                    ����ʱ��_IN IN ���˻����ӡ.����ʱ��%TYPE,
'                    ����_IN IN ���˻����ӡ.����%TYPE,
'                    ɾ��_IN Number:=0
                    gstrSQL = "ZL_���˻����ӡ_UPDATE(" & mlng�ļ�ID & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss')," & intUsedRows & "," & IIf(blnDel, "1", "0") & ")"
                    strSQL(ReDimArray(strSQL)) = gstrSQL
                    
                    'ֻҪ�޸Ĺ�����,��Ȼ��ִ�д�ӡ����,�����������л������ڵĴ���
                    If InStr(1, "," & strDays & ",", "," & Mid(strDatetime, 1, 10) & ",") = 0 Then
                        'ͬ����������Ļ���(ҹ��,ȫ����ܿ���Ĵ���)
                        strDays = strDays & "," & Mid(strDatetime, 1, 10)
                        gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & Mid(strDatetime, 1, 10) & "')"
                        strSQL(ReDimArray(strSQL)) = gstrSQL
                        
                        strTemp = Format(DateAdd("d", 1, CDate(strDatetime)), "yyyy-MM-dd")
                        If InStr(1, "," & strDays & ",", "," & strTemp & ",") = 0 Then
                            strDays = strDays & "," & strTemp
                            gstrSQL = "ZL_��������_UPDATE(" & mlng�ļ�ID & ",'" & strTemp & "')"
                            strSQL(ReDimArray(strSQL)) = gstrSQL
                        End If
                    End If
                    
                    blnSaved = False
                    If .EOF Then Exit Do
                End If
                
                '����ֵ
                intPage = !ҳ��
                intRow = !�к�
                strDate = ""
                strDatetime = ""
                lngRecord = NVL(!��¼ID, 0)
            End If
            
            If !�к� = mlngDate Then
                If NVL(!����, 0) = 1 Then
                    arrCollect = Split(!����, ";")
                    strDatetime = arrCollect(3)
                '    �ļ�ID_IN IN ���˻�������.�ļ�ID%TYPE,
                '    ����ʱ��_IN IN ���˻�������.����ʱ��%TYPE,
                '    �������_IN IN ���˻�������.�������%TYPE,
                '    �����ı�_IN IN ���˻�������.�����ı�%TYPE,
                '    ���ܱ��_IN IN ���˻�������.���ܱ��%TYPE,
                '    ɾ��_IN Number:=0
                    gstrSQL = "ZL_���˻�������_COLLECT(" & mlng�ļ�ID & ",to_date('" & arrCollect(3) & "','yyyy-MM-dd hh24:mi:ss')," & _
                            Val(arrCollect(1)) & ",'" & arrCollect(0) & "'," & Val(arrCollect(2)) & "," & !ɾ�� & ")"
                    strSQL(ReDimArray(strSQL)) = gstrSQL
                    blnSaved = True
                Else
                    strDate = NVL(!����)
                    If strDate <> "" Then
                        If mblnDateAd Then
                            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
                        Else
                            strDate = Format(strDate, "yyyy-MM-dd")
                        End If
                    End If
                End If
            ElseIf !�к� = mlngTime Then
                strTime = NVL(!����)
                If strDatetime = "" Then
                    If strDate = "" Then strDate = Mid(strCurrDate, 1, 10)
                    strDatetime = strDate & " " & strTime & ":00"
                End If
                
                If lngRecord <> 0 Then
                    '���·���ʱ��
                    gstrSQL = "Zl_���˻�������_����ʱ��(" & lngRecord & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'))"
                    strSQL(ReDimArray(strSQL)) = gstrSQL
                    blnSaved = True
                End If
            Else
                If !�к� > mlngTime Then
                    'ȡָ����Ԫ�������
                    strCellData = NVL(!����)
                    strPart = NVL(!��λ)
                    strReturn = ShowInput(!�к�, strCellData, True)
                    'strOrders��ʽ����Ŀ���,��Ŀ���...
                    'strValues��ʽ��ֵ'ֵ'ֵ...
                    arrOrder = Split(Split(strReturn, "||")(0), ",")
                    arrValue = Split(Split(strReturn, "||")(1) & "'", "'")
                    arrPart = Split(strPart & "/////", "/")
                    
                    intMax = UBound(arrOrder)
                    For intPos = 0 To intMax
    '                    �ļ�ID_IN IN ���˻�������.�ļ�ID%TYPE,
    '                    ����ʱ��_IN IN ���˻�������.����ʱ��%TYPE,
    '                    ��¼����_IN IN ���˻�����ϸ.��¼����%TYPE,          --������Ŀ=1���ϱ�˵��=2�������ձ��=4��ǩ����¼=5���±�˵��=6�����������=9
    '                    ��Ŀ���_IN IN ���˻�����ϸ.��Ŀ���%TYPE,          --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
    '                    ��¼����_IN IN ���˻�����ϸ.��¼����%TYPE := NULL,  --��¼���ݣ��������Ϊ�գ��������ǰ�����ݣ�37��38/37
    '                    ���²�λ_IN IN ���˻�����ϸ.���²�λ%TYPE := NULL,
    '                    ���˼�¼_IN IN NUMBER := 1,
                        gstrSQL = "ZL_���˻�������_UPDATE(" & mlng�ļ�ID & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'),1," & _
                                arrOrder(intPos) & ",'" & arrValue(intPos) & "','" & arrPart(intPos) & "'," & intAllow & ",0," & IIf(mblnVerify, 1, 0) & ")"
                        strSQL(ReDimArray(strSQL)) = gstrSQL
                        blnSaved = True
                    Next
                    mrsItems.Filter = 0
                End If
            End If
            
            .MoveNext
        Loop
        
        If blnSaved Then GoTo endWork
        mrsDataMap.Filter = 0
    End With

    'ѭ��ִ��SQL��������
    On Error Resume Next
    intMax = UBound(strSQL)
    
    gcnOracle.BeginTrans
    blnTrans = True
    
    On Error GoTo errHand
    If intMax > 0 Then
'        objStream.WriteLine (Now & "׼����������")
        For intPos = 1 To intMax
            If strSQL(intPos) <> "" Then
                Debug.Print strSQL(intPos)
    '            objStream.WriteLine (Now & "��SQL��" & strSQL(intPos))
                Call zlDatabase.ExecuteProcedure(strSQL(intPos), "���滤���¼������")
            End If
        Next
    '    objStream.WriteLine (Now & "�����������")
    End If
    If mblnVerify Then
        If Not SignMe(mblnVerify) Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
    End If
    
    gcnOracle.CommitTrans
    SaveData = True
    blnTrans = False
    mblnChange = False
    mblnVerify = False
    
    RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    RaiseEvent AfterRefresh
    Exit Function
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cboС��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtС������.SetFocus
End Sub

Private Sub cboС�᷶Χ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboС��.SetFocus
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strDate As String, strTime As String
    Dim strLockItem As String                   'ͬ������������,�������޸Ļ�ɾ��
    Dim lngTop As Long, lngHeight As Long
    Dim intMax As Integer                       'ͬ������������ռ�õ��������
    Dim intNULL As Integer, lngStartRow As Long
    Dim lngRow As Long, lngCOL As Long, lngRows As Long, lngCols As Long
    Dim strKey As String, strField As String, strValue As String
    
    Select Case Control.ID
    'ճ��,���ʱ��Ҫͬ��mrsCellMap����
    Case conMenu_Edit_Copy
        '����ָ�������е�����
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        lngRow = GetStartRow(VsfData.ROW)
        If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then Exit Sub
        
        '���Ƽ�¼��
        Set mrsCopyMap = New ADODB.Recordset
        Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
        
        '�õ�ָ�������е���ʼ��,������
        lngCols = VsfData.Cols - 1
        lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngRows = lngRow + lngRows - 1
        For lngRow = lngRow To lngRows
            mrsCopyMap.AddNew
            mrsCopyMap!ҳ�� = mintҳ��
            mrsCopyMap!�к� = lngRow
            For lngCOL = 0 To lngCols - VsfData.FixedCols    '����һ���̶���
                mrsCopyMap.Fields(cControlFields + lngCOL).Value = IIf(VsfData.TextMatrix(lngRow, lngCOL + VsfData.FixedCols) = "", Null, VsfData.TextMatrix(lngRow, lngCOL + VsfData.FixedCols))
            Next
            mrsCopyMap.Update
        Next
    Case conMenu_Edit_PASTE
        'ճ��ʱ����Ŀ�������帲�ǣ�ͬ�������������У���г���
        '���Ŀ���ܲ�ͬҳ����Ŀ��ͬ����λ��ͬ�����Բ����ǻ��Ŀ
        'ͬ������ռ�õ��������䣬�粻������ӿհ��У�����ճ��
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If mrsCopyMap.RecordCount = 0 Then Exit Sub
        
        '��ҳ�����в���������н���ճ��,ɾ��,ֻ�ܱ༭�����Ŀ�����
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        If InStr(1, VsfData.TextMatrix(lngStartRow, mlngRowCount), "|") <> 0 And lngStartRow = 3 Then
            If Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngStartRow, mlngRowCurrent)) Then
                RaiseEvent AfterRowColChange("��ҳ�����в�����ճ�������л�����һҳ���в�����", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
        
        '���Ŀ���������Ƿ����ͬ������������,�����������ͬ���ļ�¼
        strLockItem = GetSynItems(2, intMax)        '1.������Ŀ���;2.�����к�
        
        '�õ�Ŀ�������е���ʼ��,������
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
        lngCols = VsfData.Cols - 1
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
            If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then Exit Sub
            lngStartRow = lngRow
            If mlngDate > -1 Then strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
        Else
            'ɾ�������������,����һ��
            lngRow = GetStartRow(VsfData.ROW)
            If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then Exit Sub
            lngStartRow = lngRow
            strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
            lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0)) - 1
            For intNULL = 1 To lngRows
                VsfData.RemoveItem lngRow + 1
            Next
        End If
        
        '������������,�������������������������ӵ�����
        intNULL = mrsCopyMap.RecordCount - 1
        For lngRow = 1 To mrsCopyMap.RecordCount - 1
            '��֤��ǰ�����������һҳ����ʾȫ
            If lngRow + VsfData.ROW > VsfData.Rows - 1 Then Exit For
            
            If Val(VsfData.TextMatrix(lngRow + VsfData.ROW, mlngRecord)) = 0 And VsfData.TextMatrix(lngRow + VsfData.ROW, mlngRowCount) = "" Then
                intNULL = intNULL - 1
            Else
                Exit For
            End If
        Next
        '�����ӿ���
        If intNULL > 0 Then
            VsfData.Rows = VsfData.Rows + intNULL
            '�ӵ�ǰ�м�¼�Ŀհ��п�ʼ��ÿ�е�λ��+�����ӵĿհ�����
            For lngRow = 1 To intNULL
                VsfData.RowPosition(VsfData.Rows - 1) = lngStartRow + 1
            Next
        End If
        
        '��ԭ���ڣ�ʱ�䣬ǿ�Ʋ������޸�
        VsfData.TextMatrix(lngStartRow, mlngDate) = strDate
        VsfData.TextMatrix(lngStartRow, mlngTime) = strTime
        '��¼�û��޸Ĺ��ĵ�Ԫ��
        If mlngDate <> -1 Then
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngDate) & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        '2\ʱ��
        strKey = mintҳ�� & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngTime) & "|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '�����������
        With mrsCopyMap
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                For lngCOL = 0 To lngCols - VsfData.FixedCols
                    Select Case lngCOL + VsfData.FixedCols
                    Case 1, mlngDate, mlngTime, mlngOperator, mlngSigner, mlngSignTime, mlngRecord
                    Case Else
                        If InStr(1, "," & strLockItem & ",", "," & lngCOL - (cHideCols - 1) & ",") = 0 And InStr(1, "," & mstrCOLNothing & ",", "," & lngCOL - (cHideCols - 1) & ",") = 0 Then
                            VsfData.TextMatrix(lngStartRow + .AbsolutePosition - 1, lngCOL + VsfData.FixedCols) = NVL(.Fields(cControlFields + lngCOL).Value)
                            
                            '�޸ı�־
                            If .AbsolutePosition = 1 Then
                                strKey = mintҳ�� & "," & lngStartRow & "," & lngCOL + VsfData.FixedCols
                                strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCOL + VsfData.FixedCols & "|" & _
                                    Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & GetMutilData(lngStartRow, lngCOL + VsfData.FixedCols, lngTop, lngHeight) & "|0"
                                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                            End If
                        End If
                    End Select
                Next
                .MoveNext
            Loop
        End With
        '�����ɫ
        'Call WriteColor
        mblnChange = True
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    
    Case conMenu_Edit_Clear
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        If VsfData.TextMatrix(VsfData.ROW, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("��ǩ�������ݲ�����ɾ����", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '��ҳ�����в���������н���ճ��,ɾ��,ֻ�ܱ༭�����Ŀ�����
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        If InStr(1, VsfData.TextMatrix(lngStartRow, mlngRowCount), "|") <> 0 And lngStartRow = 3 Then
            If Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngStartRow, mlngRowCurrent)) Then
                RaiseEvent AfterRowColChange("��ҳ�����в�����ɾ�������л�����һҳ���в�����", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
        
        '���Ŀ���������Ƿ����ͬ������������,�����������ͬ���ļ�¼
        strLockItem = GetSynItems(2, intMax)        '1.������Ŀ���;2.�����к�
        
        '׼��ɾ��
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|����|ɾ��"
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
        Else
            lngRow = GetStartRow(VsfData.ROW)
            lngStartRow = lngRow
            If VsfData.TextMatrix(lngStartRow, mlngSigner) <> "" Then
                RaiseEvent AfterRowColChange("��ǩ�������ݲ�����ɾ����", True, mblnSign, mblnArchive)
                Exit Sub
            End If
            
            'ɾ������������
            lngRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
            For intNULL = 2 To lngRows
                VsfData.RowHidden(lngRow + intNULL - 1) = True
            Next
        End If
        
        '��¼�û��޸Ĺ��ĵ�Ԫ��
        If VsfData.TextMatrix(lngStartRow, mlngCollectType) = 0 Then
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngDate) & "|0|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            '2\ʱ��
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngTime
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngTime & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngTime) & "|0|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        Else
            '1\����
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & _
                    VsfData.TextMatrix(lngStartRow, mlngCollectText) & ";" & Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) & ";" & Val(VsfData.TextMatrix(lngStartRow, mlngCollectStyle)) & ";" & VsfData.TextMatrix(lngStartRow, mlngCollectDay) & "|1|1"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        'ɾ����ʼ���з�ͬ��������
        If strLockItem = "" Then
            VsfData.RowHidden(lngRow) = True
            If VsfData.TextMatrix(lngStartRow, mlngCollectType) = 0 Then
                '��д�޸ı�־
                For lngCOL = mlngTime + 1 To mlngNoEditor - 1
                    strKey = mintҳ�� & "," & lngStartRow & "," & lngCOL
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCOL & "|" & _
                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||0|1"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                Next
            End If
        Else
            '��д�޸ı�־(����ͬ������,������ʱ���в��������)``
            For lngCOL = mlngTime + 1 To mlngNoEditor - 1
                If InStr(1, "," & strLockItem & ",", "," & lngCOL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 And lngCOL <> mlngDate And lngCOL <> mlngTime Then
                    VsfData.TextMatrix(lngStartRow, lngCOL) = ""
                    
                    strKey = mintҳ�� & "," & lngStartRow & "," & lngCOL
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCOL & "|" & _
                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||0|1"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                End If
            Next
            VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
        End If
            
        Call FillPage
        mblnChange = True
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
        
    Case conMenu_Edit_SPECIALCHAR
        
        '��鵱ǰ¼��ؼ�
        On Error Resume Next
        Dim objTXT As TextBox
        Dim strText As String
        Dim intPos As Integer, intLen As Integer
        
        mstrSymbol = frmInsSymbol.ShowMe(False, 0)
        If mintSymbol = -1 Then
            Set objTXT = txtInput
        Else
            Set objTXT = txt(mintSymbol)
        End If
        strText = objTXT.Text
        intPos = objTXT.SelStart
        intLen = Len(objTXT)
        objTXT.Text = Mid(strText, 1, intPos) & mstrSymbol & Mid(strText, intPos + 1)
    Case conMenu_Edit_Append
        Call BoundItems(VsfData.COL - (cHideCols + VsfData.FixedCols - 1))
    Case conMenu_Edit_PrevPage
        If mintҳ�� > 1 Then
            If Not DataMap_Save Then Exit Sub
            mintҳ�� = mintҳ�� - 1
            '���²�ѯSQL
            '������ȡ����
            mblnInit = False
            Call InitVariable
            Call InitCons
            Call ReadStruDef
            Call zlRefresh
            mblnInit = True
        End If
    Case conMenu_Edit_NextPage
        If mintҳ�� < mint����ҳ + 1 Then
            If Not DataMap_Save Then Exit Sub
            mintҳ�� = mintҳ�� + 1
            '���²�ѯSQL
            '������ȡ����
            mblnInit = False
            Call InitVariable
            Call InitCons
            Call ReadStruDef
            Call zlRefresh
            mblnInit = True
        End If
    Case conMenu_Edit_Word
        Call cmdWord_Click
    Case conMenu_Edit_Brief
        Call ShowBrief
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim arrData
    Dim blnFind As Boolean
    Dim strItem As String
    Dim intDo  As Integer, intCount As Integer
    
    If Not mblnInit Then Exit Sub
    Select Case Control.ID
    Case conMenu_Edit_Copy
        Control.Enabled = Not mblnShow And Not mblnArchive And Not mblnVerify And mblnEditable
    Case conMenu_Edit_PASTE
        Control.Enabled = False
        If mrsCopyMap.State = 0 Then Exit Sub
        'ǩ�����ݲ�����ճ��
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "" Then
            intDo = GetStartRow(VsfData.ROW)
        Else
            intDo = VsfData.ROW
        End If
        If VsfData.TextMatrix(intDo, mlngSigner) <> "" Then Exit Sub
        If Val(VsfData.TextMatrix(intDo, mlngCollectType)) <> 0 Then Exit Sub
        
        Control.Enabled = Not mblnShow And Not mblnArchive And mblnEditable And mrsCopyMap.RecordCount
    Case conMenu_Edit_Clear
        Control.Enabled = False
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "" Then
            intDo = GetStartRow(VsfData.ROW)
        Else
            intDo = VsfData.ROW
        End If
        If VsfData.TextMatrix(intDo, mlngSigner) <> "" Then Exit Sub
        
        Control.Enabled = Not mblnArchive And Not mblnVerify And mblnEditable
    Case conMenu_Edit_SPECIALCHAR
        Control.Enabled = mblnShow And Not mblnArchive And mblnEditable And (mintType = 0 Or mintType = 6)
    Case conMenu_Edit_Append
        Control.Enabled = (InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0) And Not mblnArchive And mblnEditable
    Case conMenu_Edit_PrevPage
        Control.Enabled = (mintҳ�� > 1)
    Case conMenu_Edit_NextPage
        Control.Enabled = (mintҳ�� < mint����ҳ + 1)
    Case conMenu_Edit_Word
        Control.Enabled = mblnEditAssistant And mblnShow And Not mblnArchive And mblnEditable
    Case conMenu_Edit_Brief
        Control.Enabled = Not mblnArchive And Not mblnVerify And mblnEditable
    End Select
End Sub

Private Sub chkSwitch_Click()
    Dim blnSel As Boolean            '�Ƿ�ȫ��ѡ��
    Dim blnUpdate As Boolean
    Dim intLevel As Integer
    Dim lngRow As Long, lngRows As Long
    Dim strKey As String, strField As String, strValue As String
    '��������ȫ��ѡ�л�ȡ��ѡ�У����������
    
    If Not mblnInit Then Exit Sub
    lngRows = VsfData.Rows - 1
    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
    
    blnSel = chkSwitch.Value
    For lngRow = VsfData.FixedRows To lngRows
        If Not VsfData.RowHidden(lngRow) Then
            If VsfData.TextMatrix(lngRow, mlngRowCount) Like "*|1" Then
                If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) = 0 Then    '�����в�����༭
                    blnUpdate = False
                    If blnSel Then
                        '���,ǩ�����ļ�¼,�ҵ�ǰ����Ա������ϴ�ǩ�������
                        If VsfData.TextMatrix(lngRow, mlngSignLevel) = "" Then
                            intLevel = δ����
                        Else
                            intLevel = Val(VsfData.TextMatrix(lngRow, mlngSignLevel)) + 1
                        End If
                        If mintVerify < intLevel And intLevel <> δ���� Then
                            blnUpdate = (VsfData.Cell(flexcpChecked, lngRow, mlngChoose) <> flexTSChecked)
                            VsfData.Cell(flexcpChecked, lngRow, mlngChoose) = flexTSChecked
                        End If
                    Else
                        blnUpdate = (VsfData.Cell(flexcpChecked, lngRow, mlngChoose) <> flexTSUnchecked)
                        VsfData.Cell(flexcpChecked, lngRow, mlngChoose) = flexTSUnchecked
                    End If
                    
                    If blnUpdate Then
                        '�����޸ļ�¼�Ա�ͬ��
                        strKey = mintҳ�� & "," & lngRow & "," & mlngChoose
                        strValue = strKey & "|" & mintҳ�� & "|" & lngRow & "|" & mlngChoose & "|" & _
                            Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & VsfData.Cell(flexcpChecked, lngRow, mlngChoose) & "|1"
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                End If
            End If
        End If
    Next
End Sub

Private Sub cmdCancel_Click()
    picBiref.Visible = False
End Sub

Private Sub cmdColumn_Click(Index As Integer)
    Dim lstItem As ListItem
    
    If cmdColumn(Index).Enabled = False Then Exit Sub
    If Index = 0 Then
        'add
        If Not lstColumnItems.SelectedItem Is Nothing Then
            Set lstItem = lstColumnUsed.ListItems.Add(, lstColumnItems.SelectedItem.Key, lstColumnItems.SelectedItem.Text)
            lstItem.SubItems(1) = lstColumnItems.SelectedItem.SubItems(1)
            lstItem.SubItems(2) = lstColumnItems.SelectedItem.SubItems(2)
            lstColumnItems.ListItems.Remove lstColumnItems.SelectedItem.Index
        End If
        If txtColumnNo.Text = "" Then
            txtColumnNo.Text = Replace(lstItem.SubItems(1), lstItem.SubItems(2), "")
        End If
    Else
        'del
        If Not lstColumnUsed.SelectedItem Is Nothing Then
            Set lstItem = lstColumnItems.ListItems.Add(, lstColumnUsed.SelectedItem.Key, lstColumnUsed.SelectedItem.Text)
            lstItem.SubItems(1) = lstColumnUsed.SelectedItem.SubItems(1)
            lstItem.SubItems(2) = lstColumnUsed.SelectedItem.SubItems(2)
            lstColumnUsed.ListItems.Remove lstColumnUsed.SelectedItem.Index
            If lstColumnUsed.ListItems.Count = 0 Then txtColumnNo.Text = ""
        End If
    End If
End Sub

Private Sub cmdFilterCancel_Click()
    picCloumn.Visible = False
End Sub

Private Sub cmdFilterOK_Click()
    Dim strPara As String
    Dim strTest As String
    Dim lngCOL As Long, lngRow As Long
    Dim intDo As Integer, intCount As Integer, intFace As Integer
    On Error GoTo errHand
    
    If lstColumnUsed.ListItems.Count > 0 Then
        If Trim(txtColumnNo.Text) = "" Then
            RaiseEvent AfterRowColChange("��ͷ���Ʋ���Ϊ�գ�", True, mblnSign, mblnArchive)
            txtColumnNo.SetFocus
            Exit Sub
        End If
        If LenB(StrConv(txtColumnNo.Text, vbFromUnicode)) > 20 Then
            RaiseEvent AfterRowColChange("��ͷ���Ʋ��ܳ���10�����ֻ�20���ַ���", True, mblnSign, mblnArchive)
            txtColumnNo.SetFocus
            Exit Sub
        End If
    End If
    
    'ƴ������ʽ����ͷ����|��Ŀ���,��λ;��Ŀ���,��λ
    strPara = Trim(txtColumnNo.Text) & "|"
    intCount = lstColumnUsed.ListItems.Count
    If intCount > 2 Then
        RaiseEvent AfterRowColChange("ÿ�а󶨵���Ŀ�����ܳ���2����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '��Ŀ��ʾ����һ��
    For intDo = 1 To intCount
        mrsItems.Filter = "��Ŀ���=" & Val(lstColumnUsed.ListItems(intDo).Text)
        If intDo = 1 Then
            intFace = mrsItems!��Ŀ��ʾ
        Else
            If intFace <> mrsItems!��Ŀ��ʾ Then
                RaiseEvent AfterRowColChange("�󶨵�������Ŀ�ı�ʾ��������һ�£���Ҫô����ѡ���Ҫô������ֵ¼���", True, mblnSign, mblnArchive)
                mrsItems.Filter = 0
                Exit Sub
            End If
        End If
        
        'ƴ��
        strTest = lstColumnUsed.ListItems(intDo).Text
        If lstColumnUsed.ListItems(intDo).SubItems(2) <> "" Then
            strTest = strTest & "," & lstColumnUsed.ListItems(intDo).SubItems(2)
        End If
        If ISActiveUsed(strTest) Then Exit Sub
        
        strPara = strPara & IIf(intDo > 1, ";", "") & strTest
        mrsItems.Filter = 0
    Next
    
    '��������
    gstrSQL = "ZL_���˻���ҳ��_UPDATE(" & mlng�ļ�ID & "," & mintҳ�� & "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",'" & strPara & "','" & gstrUserName & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ŀ������")
    picCloumn.Visible = False
    lngCOL = VsfData.COL
    lngRow = VsfData.ROW
    
    '���²�ѯSQL
    '������ȡ����
    mblnInit = False
    Call InitVariable
    Call InitCons
    Call ReadStruDef
    Call zlRefresh
    mblnInit = True
    
    VsfData.ROW = lngRow
    VsfData.COL = lngCOL
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function ISActiveUsed(ByVal strTest As String) As Boolean
    Dim arrData, arrCol
    Dim lngCOL As Long
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '���ĳ�����Ŀ�Ƿ��ѱ������а�
    ISActiveUsed = True
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        arrCol = Split(Split(arrData(intDo), "|")(1), ";")
        lngCOL = Split(Split(arrData(intDo), "|")(0), ";")(0)
        intMax = UBound(arrCol)
        For intIn = 0 To intMax
            If strTest = arrCol(intIn) And VsfData.COL - (cHideCols + VsfData.FixedCols - 1) <> lngCOL Then
                RaiseEvent AfterRowColChange(Split(strTest, ",")(1) & mrsItems!��Ŀ���� & " �Ѿ����󶨵�" & lngCOL & "�У��������ظ��󶨣�", True, mblnSign, mblnArchive)
                Exit Function
            End If
        Next
    Next
    ISActiveUsed = False
End Function

Private Function GetActivePart(ByVal intFindCol As Integer, ByVal intItem As Integer) As String
    '��ȡָ���еĻ��Ŀ
    Dim arrData
    Dim arrCol
    Dim intCol As Integer, strPart As String
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '�����Ŀ���뵽��ѯSQL�У���ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...
    '�󶨶����Ŀ�����о��Զ�תΪ�Խ�����
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        If intCol = intFindCol - cHideCols Then
            arrCol = Split(Split(arrData(intDo), "|")(1), ";")
            strPart = Split(arrCol(intItem), ",")(1)
            Exit For
        End If
    Next
    GetActivePart = strPart
End Function

Private Function CalcCollect(ByVal lngItem As Long, ByVal strStart As String, ByVal strEnd As String) As Double
    Dim dblCollect As Double
    On Error GoTo errHand
    
    gstrSQL = " SELECT /*+ RULE */ NVL(SUM(NVL(��¼����,0)),0) AS ����" & _
              " From ���˻�����ϸ A,���˻������� B," & vbNewLine & _
              "      (Select ��� From ���������Ŀ Start With ���=[2] Connect By Prior ���=�����) C" & vbNewLine & _
              " Where A.��¼ID=B.ID And A.��ֹ�汾 Is NULL And A.��¼����=1 AND B.�������=0 And A.��Ŀ���=C.���" & vbNewLine & _
              " And B.�ļ�ID=[1] And B.����ʱ�� Between [3] And [4]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��������", mlng�ļ�ID, lngItem, CDate(strStart), CDate(strEnd))
    dblCollect = rsTemp!����
    
    CalcCollect = dblCollect
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    Dim arrTime
    Dim arrItem
    Dim arrValue() As Double
    Dim bln���� As Boolean, blnExit As Boolean
    Dim lngStart As Long
    Dim lngCOL As Long, lngCount As Long, lngRow As Long, lngRows As Long
    Dim strToday As String, str����ʱ�� As String
    Dim strStartDate As String, strEndDate As String
    Dim strStartTime As String, strEndTime As String
    Dim strKey As String, strField As String, strValue As String
    On Error GoTo errHand
    '����һ���µĻ��ܼ�¼
    
    If InStr(1, txtС������.Text, ";") <> 0 Then
        RaiseEvent AfterRowColChange("С�������в��ܺ��зֺţ�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    If InStr(1, txtС������.Text, "'") <> 0 Then
        RaiseEvent AfterRowColChange("С�������в��ܺ��е����ţ�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    If LenB(StrConv(txtС������.Text, vbFromUnicode)) > 50 Then
        RaiseEvent AfterRowColChange("С�����Ʋ��ܳ���50���ַ���25�����֣�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '���ʱ�䷶Χ�Ƿ����
    '����
    '    �� ����
    '    ҹ ���� - ����
    '    ȫ ���� - ����
    '����
    '    �� ����
    '    ҹ ���� - ����
    '    ȫ ���� - ����
    strToday = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    arrTime = Split(cboС��.Tag, ";")   '��ʽ:��ʼʱ��,����ʱ��;��ʼʱ��,����ʱ��
    strStartTime = Split(arrTime(cboС��.ItemData(cboС��.ListIndex) - 1), ",")(0)
    strEndTime = Split(arrTime(cboС��.ItemData(cboС��.ListIndex) - 1), ",")(1)
    If strEndTime < strStartTime Then bln���� = True
    If bln���� = True Then
        strStartDate = strToday & " " & strStartTime & ":00"
        strEndDate = DateAdd("d", 1, CDate(strToday)) & " " & strEndTime & ":59"
    Else
        strStartDate = strToday & " " & strStartTime & ":00"
        strEndDate = strToday & " " & strEndTime & ":59"
    End If
    If cboС�᷶Χ.ListIndex = 1 Then   '����,ȫ����1��
        strStartDate = Format(DateAdd("d", -1, CDate(strStartDate)), "yyyy-MM-dd HH:mm:ss")
        strEndDate = Format(DateAdd("d", -1, CDate(strEndDate)), "yyyy-MM-dd HH:mm:ss")
    End If
    str����ʱ�� = DateAdd("s", 1 + cboС��.ItemData(cboС��.ListIndex), strEndDate)
    
    '����Ƿ��Ѿ����ڸ�����
    blnExit = False
    mrsDataMap.Filter = "ɾ��=0 And �������=" & -1 * cboС��.ItemData(cboС��.ListIndex) & " And ��������='" & str����ʱ�� & "'"    '��¼ID>0������,���ǵ��������
    blnExit = (mrsDataMap.RecordCount)
    mrsDataMap.Filter = 0
    
    If blnExit Then
        RaiseEvent AfterRowColChange("��Ҫ��ӵ�С�������Ѵ��ڣ�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '���ҿհ���
    lngRows = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngRows
        If Val(VsfData.TextMatrix(lngRow, mlngRecord)) = 0 And VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then
            lngStart = lngRow
            Exit For
        End If
    Next
    If lngStart = 0 Then
        '˵��û���ҵ��հ���
        VsfData.Rows = VsfData.Rows + 1
        lngStart = VsfData.Rows - 1
    End If
    
    'ͳ�ƻ�������(�����ݿ��л���,��ǰ����ֻ��¼���Ƿ��޸�,����֪��ԭֵ�Ƕ���,���Ե�ǰδ��������ݲ�����)
    '������Ŀ����
    '������Ŀ�м���:col;1|col;4,5
    arrItem = Split(mstrCollectItems, ",")
    lngRows = UBound(arrItem)
    ReDim Preserve arrValue(lngRows) As Double
    For lngRow = 0 To lngRows
        arrValue(lngRow) = CalcCollect(arrItem(lngRow), strStartDate, strEndDate)
    Next
    
    'ͨ�ò���
    VsfData.TextMatrix(lngStart, mlngDate) = txtС������.Text
    VsfData.TextMatrix(lngStart, mlngTime) = txtС������.Text
    VsfData.TextMatrix(lngStart, mlngRowCount) = "1|1"                          'Ϊ�˱�֤ʱ�䲻�ظ�,��ȡ����ʱ��+��ķ�ʽ
    VsfData.TextMatrix(lngStart, mlngRowCurrent) = "1"
    VsfData.TextMatrix(lngStart, mlngCollectText) = txtС������.Text
    VsfData.TextMatrix(lngStart, mlngCollectType) = -1 * (cboС��.ItemData(cboС��.ListIndex))      '��ʾС��;-1�װ�;-2ҹ��;3-ȫ��
    VsfData.TextMatrix(lngStart, mlngCollectStyle) = cbo��ʶ.ListIndex         '����24Сʱ,���»�����
    VsfData.TextMatrix(lngStart, mlngCollectDay) = str����ʱ��
    
    'ͬ������������ʱ���е�����
    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|����|ɾ��"
    '1\����
    strKey = mintҳ�� & "," & lngStart & "," & mlngDate
    strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngDate & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & _
            txtС������.Text & ";" & -1 * (cboС��.ItemData(cboС��.ListIndex)) & ";" & cbo��ʶ.ListIndex & ";" & str����ʱ�� & "|1|0"
    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    
    'չ��
    arrItem = Split(mstrColCollect, "|")
    lngCount = 0
    lngRows = UBound(arrItem)
    For lngRow = 0 To lngRows
        lngCOL = Split(arrItem(lngRow), ";")(0)
        If UBound(Split(Split(arrItem(lngRow), ";")(1), ",")) = 1 Then
            strValue = arrValue(lngCount) & "/" & arrValue(lngCount + 1)
            lngCount = lngCount + 2
        Else
            strValue = arrValue(lngCount)
            lngCount = lngCount + 1
        End If
        
        VsfData.TextMatrix(lngStart, lngCOL + cHideCols) = strValue
        strKey = mintҳ�� & "," & lngStart & "," & lngCOL + cHideCols
        strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & lngCOL + cHideCols & "|" & _
            Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.TextMatrix(lngStart, lngCOL + cHideCols) & "|1|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    Next
    
'    '�ϲ���Ԫ��
'    lngRows = Split(Split(mstrColCollect, "|")(0), ";")(0) + cHideCols - 1
'    For lngRow = mlngTime + 1 To lngRows
'        VsfData.TextMatrix(lngStart, lngRow) = txtС������.Text
'    Next
'    VsfData.MergeCells = flexMergeRestrictRows          '���ᵥԪ��Ȼ�ǵ����ϲ�,�ϲ�����������ϲ���Ԫ��
'    VsfData.MergeRow(lngStart) = True
    
    mblnChange = True
    picBiref.Visible = False
    
    RaiseEvent AfterDataChanged(mblnChange)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdWord_Click()
    Dim strInput As String
    '�����ʾ�ѡ����
    
    If cmdWord.Tag = -1 Then
        strInput = txtInput.Text
    Else
        strInput = txt(Val(cmdWord.Tag)).Text
    End If
    strInput = frmEditAssistant.ShowMe(Me, mlng����ID, mlng��ҳID, mintӤ��, strInput)
    
    If cmdWord.Tag = -1 Then
        txtInput.Text = strInput
    Else
        txt(Val(cmdWord.Tag)).Text = strInput
    End If
End Sub

Private Sub ShowBrief()
    Dim strStart As String, strEnd As String
    Dim strHave As String, strDate As String
    Dim strTag As String    'cboС���tag�б���ʱ��Σ���ʽ����ʼ,����;��ʼ,����
    Dim rsData As New ADODB.Recordset
    On Error GoTo errHand
    '��ʾС�ᴰ��
    
    If Not DataMap_Save Then Exit Sub       '��������,�Ա�ѡ��С���ʱ��������ݼ��
    
    '����¼���Ƿ���ڻ�����Ŀ�У�������������˳�
    If mstrCollectItems = "" Then
        RaiseEvent AfterRowColChange("��ǰ�ļ���δʹ�û�����Ŀ��", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '��ȡ����ʱ��(���=3Ϊȫ��С��)
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    gstrSQL = "Select ���,����,��ʼ,���� From �������ʱ�� Order by ��� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡС��")
    If rsTemp.RecordCount = 0 Then
        RaiseEvent AfterRowColChange("��δ���û���ʱ��,�����ڻ�����Ŀ����ģ��Ļ�����Ŀ�����ã�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    rsTemp.Filter = "���=3"
    If rsTemp.RecordCount = 0 Then
        rsTemp.Filter = 0
        RaiseEvent AfterRowColChange("ȫ�����ʱ��δ����,�����ڻ�����Ŀ����ģ��Ļ�����Ŀ�����ã�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    strStart = NVL(rsTemp!��ʼ)
    strEnd = NVL(rsTemp!����)
    rsTemp.Filter = 0
    
    '���ػ������
    cboС��.Clear
    With rsTemp
        Do While Not .EOF
            If Not (NVL(!��ʼ) = "" Or NVL(!����) = "") Then
                cboС��.AddItem !����
                cboС��.ItemData(cboС��.NewIndex) = !���
                strTag = strTag & ";" & !��ʼ & "," & !����
            End If
            .MoveNext
        Loop
        If strTag <> "" Then
            cboС��.Tag = Mid(strTag, 2)
            cboС��.ListIndex = 0
        Else
            RaiseEvent AfterRowColChange("����Ļ�����ȫ����ӣ�", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    End With
    
    With cboС�᷶Χ
        .Clear
        .AddItem "����"
        .AddItem "����"
        .ListIndex = 0
    End With
    With cbo��ʶ
        .Clear
        .AddItem "������"
        .AddItem "���º��߱�ʶ"
        .AddItem "���ܽ����˫���߱�ʶ"
        .ListIndex = 0
    End With
    
    '��������
    With picBiref
        .Top = VsfData.Top + VsfData.Cell(flexcpTop, VsfData.FixedRows, mlngTime)
        .Left = (ScaleWidth - .Width) / 2
        .Visible = True
    End With
    
    On Error Resume Next
    cboС��.SetFocus
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub VsfData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Dim lngRow As Long, lngCOL As Long
    Dim dblHeight As Double, dblWidth As Double
    
    If Not mblnInit Then Exit Sub
    Call InitCons
    
'    '����̶��еĸ߶�
'    For lngRow = 0 To 2
'        If Not VsfData.RowHidden(lngRow) Then dblHeight = dblHeight + VsfData.ROWHEIGHT(lngRow)
'    Next
'    '�ӿɼ��п�ʼ���²������һ���ɼ���
'    For lngRow = NewTopRow To VsfData.Rows - 1
'        If Not VsfData.RowIsVisible(lngRow) Then
'            lngRow = lngRow - 1
'            Exit For
'        End If
'    Next
'    '�ӿɼ��п�ʼ�������һ���ɼ���
'    For lngCol = NewLeftCol To VsfData.Cols - 1
'        If Not VsfData.ColIsVisible(lngCol) Then
'            lngCol = lngCol - 1
'            Exit For
'        Else
'            dblWidth = dblWidth + VsfData.ColWidth(lngCol)
'        End If
'    Next
'
'    If Not VsfData.RowIsVisible(VsfData.Row) Then
'        VsfData.Row = VsfData.Row + IIf(OldTopRow < NewTopRow, 1, -1)
'    Else
'        '��ǰ�����еĸ߶�+�̶��еĸ߶�������ڱ��ؼ��ĸ߶�,˵����ǰѡ��������д�����ס���ֵ����
'        If VsfData.Row >= lngRow - 1 And CellRect.Bottom * (lngRow - NewTopRow + 1) + dblHeight >= VsfData.ClientHeight Then
'            '��ס���ֵ������
'            VsfData.Row = VsfData.Row + IIf(OldTopRow < NewTopRow, 1, -1)
'        End If
'    End If
'
'    If Not VsfData.ColIsVisible(VsfData.Col) Then
'        VsfData.Col = VsfData.Col + IIf(OldLeftCol < NewLeftCol, 1, -1)
'    Else
'        '��ǰ�����еĸ߶�+�̶��еĸ߶�������ڱ��ؼ��ĸ߶�,˵����ǰѡ��������д�����ס���ֵ����
'        If VsfData.Col = lngCol And dblWidth >= VsfData.ClientWidth Then
'            '��ס���ֵ������
'            VsfData.Col = VsfData.Col + IIf(OldLeftCol < NewLeftCol, 1, -1)
'        End If
'    End If
'
'    Call VsfData_EnterCell
End Sub

Private Sub VsfData_DblClick()
    Call VsfData_KeyDown(Asc("A"), 0)
End Sub

Private Sub VsfData_EnterCell()
    Dim strCols As String
    Dim intMax As Integer
    Dim lngStart As Long
    On Error Resume Next
    
    '��������ʾ��¼��ؼ�
    Select Case mintType
    Case 0, 3
        picInput.Visible = False
    Case 1, 2
        lstSelect(mintType - 1).Visible = False
    Case 4, 5
        picDouble.Visible = False
    Case 6
        picMutilInput.Visible = False
    End Select
    cmdWord.Visible = False
    
    'δ������в�����¼������
    mintType = -1
    If InStr(1, mstrPrivs, "�����¼�Ǽ�") = 0 Then Exit Sub
    
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        lngStart = VsfData.ROW
    Else
        lngStart = GetStartRow(VsfData.ROW)
    End If
    If Val(VsfData.TextMatrix(lngStart, mlngCollectType)) < 0 Then Exit Sub   '�����в�����༭,ǩ��
    If mblnVerify Then  '�������mblnShow�ж���������
        If VsfData.COL = mlngChoose Then Call VsfData_KeyDown(vbKeySpace, 0): Exit Sub
        If VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then Exit Sub
        If Val(VsfData.TextMatrix(lngStart, mlngRecord)) = 0 Then Exit Sub
        If VsfData.TextMatrix(lngStart, mlngSigner) = "" Then Exit Sub
    Else
        '��ǩ��������ֻ������ǩ״̬���޸�
        If InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/") <> 0 Then Exit Sub
        '�����ǰ����Ա�ļ������ǩ������Ա�ļ����,��������༭����
        If VsfData.TextMatrix(lngStart, mlngSigner) <> "" Then
            If mintVerify > Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1 Then Exit Sub
        End If
    End If
    If mblnArchive Then Exit Sub
    If Not mblnShow Or Not mblnEditable Then Exit Sub
    
    '��ҳ�����в���������н���ճ��,ɾ��,ֻ�ܱ༭�����Ŀ�����
    If InStr(1, VsfData.TextMatrix(lngStart, mlngRowCount), "|") <> 0 And lngStart = 3 Then
        If Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngStart, mlngRowCurrent)) Then
            If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then
                RaiseEvent AfterRowColChange("�������޸Ŀ�ҳ�����еĻ��Ŀ���ݣ�", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
    End If
    'ͬ�������в�����༭
    strCols = GetSynItems(2, intMax)
    If strCols <> "" Then
        '����ͬ�����ݵ���,������ʱ���ǲ������޸ĵ�
        If VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then Exit Sub
        strCols = "," & strCols & ","
        If InStr(1, strCols, "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then Exit Sub
    End If
    
    If VsfData.COL <= mlngNoEditor - 1 Then Call ShowInput
    '�ÿؼ���ý���
    Select Case mintType
    Case 0, 3
        picInput.SetFocus
    Case 1, 2
        lstSelect(mintType - 1).SetFocus
    Case 4, 5
        picDouble.SetFocus
    Case 6
        picMutilInput.SetFocus
    End Select
End Sub

Private Sub vsfData_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    Dim strCols As String
    Dim intMax As Integer
    If mblnInit = False Then Exit Sub
    If mblnEditable = False Then Exit Sub
    If OldRow = NewRow And OldCol = NewCol Then Exit Sub
    
    'ѡ����,ͬ��������ֱ���˳�,����˴������ʾ��Ϣ
    If NewCol = mlngChoose Then Exit Sub
    strCols = GetSynItems(2, intMax)
    If strCols <> "" Then
        strCols = "," & strCols & ","
        If InStr(1, strCols, "," & NewCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then Exit Sub
    End If
    
    '��ʾ��ǰ��Ŀ�������Ϣ
    mrsSelItems.Filter = "��=" & NewCol - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        mrsItems.Filter = "��Ŀ���=" & mrsSelItems!��Ŀ���
        If mrsItems.RecordCount <> 0 Then
            If NVL(mrsItems!��Ŀֵ��) <> "" Then
                If mrsItems!��Ŀ���� = 0 Then
                    strInfo = "��Ч��Χ:" & Split(mrsItems!��Ŀֵ��, ";")(0) & "��" & Split(mrsItems!��Ŀֵ��, ";")(1)
                Else
                    strInfo = "��Ч��Χ:" & mrsItems!��Ŀֵ��
                End If
            Else
                strInfo = ""
            End If
        End If
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    '����Ƿ���ǩ��
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        intMax = VsfData.ROW
    Else
        intMax = GetStartRow(VsfData.ROW)
    End If
    mblnSign = (VsfData.TextMatrix(intMax, mlngSigner) <> "")
    
    RaiseEvent AfterRowColChange(strInfo, False, mblnSign, mblnArchive)
End Sub

Private Sub vsfData_DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Call DrawCell(hDC, ROW, COL, Left, Top, Right, Bottom, Done)
End Sub

Private Sub VsfData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngStart As Long
    Dim intLevel As Integer
    Dim strField As String, strKey As String, strValue As String
    
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    ElseIf KeyCode = vbKeySpace And mblnVerify Then
        'ֻ��ѡ��ʼ��
        lngStart = GetStartRow(VsfData.ROW)
        If VsfData.TextMatrix(lngStart, mlngTime) = "" Then Exit Sub
        
        '��ǩʱ,��ǰ��¼��ǩ��,�Ҳ���Ա��ǩ��������ϴ�ǩ������߲�����
        If VsfData.TextMatrix(lngStart, mlngSignLevel) = "" Then
            RaiseEvent AfterRowColChange("�����ݻ�δǩ�������ܽ�����ǩ��", True, mblnSign, mblnArchive)
            Exit Sub
        Else
            intLevel = Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1
        End If
        If mintVerify >= intLevel Then
            RaiseEvent AfterRowColChange("���ļ���Ҫ���ϴ���ǩ�˵ļ���߲��ܹ�ѡ�ü�¼��", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = IIf(VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = flexTSChecked, flexTSUnchecked, flexTSChecked)
        '�����޸ļ�¼�Ա�ͬ��
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
        strKey = mintҳ�� & "," & lngStart & "," & mlngChoose
        strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngChoose & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.Cell(flexcpChecked, lngStart, mlngChoose) & "|1"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    Else
        If Not (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or Shift <> 0) Then
            mblnShow = True
            Call VsfData_EnterCell
        End If
    End If
End Sub

Private Sub InitVariable()
    '�������
    mlngDate = -1
    mlngTime = -1
    mlngOperator = -1
    mlngSigner = -1
    mlngSignTime = -1
    mlngRecord = -1
    mlngNoEditor = -1
    
    mblnChange = False
    mblnShow = False
    mblnSign = False
    mblnArchive = False
    mblnEditAssistant = False
    
End Sub

Private Sub InitCons()
    '��������ؼ�
    picInput.Visible = False
    lstSelect(0).Visible = False
    lstSelect(1).Visible = False
    picDouble.Visible = False
    picMutilInput.Visible = False
    cmdWord.Visible = False
    
    picBiref.Visible = False
    picCloumn.Visible = False
End Sub

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim Rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "�˵���"
    cbsThis.ActiveMenuBar.Visible = False
    
    cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 16, 16
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
        '------------------------------------------------------------------------------------------------------------------
        '����������
        Set cbrToolBar = cbsThis.Add("��׼", xtpBarTop)
        cbrToolBar.ShowTextBelowIcons = False
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "����"): cbrControl.ToolTipText = "����(Ctrl+C)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PASTE, "ճ��"):  cbrControl.ToolTipText = "ճ��(Ctrl+V)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "���"):   cbrControl.ToolTipText = "���"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SPECIALCHAR, "�������"):  cbrControl.ToolTipText = "�����������(Ctrl+D)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Word, "�ʾ�ѡ��"):  cbrControl.ToolTipText = "�ʾ�ѡ��(Ctrl+W)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Brief, "С��"): cbrControl.ToolTipText = "С��"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "�а�"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "�а�"
        
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrevPage, "��ҳ"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "��ҳ"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextPage, "��ҳ"):   cbrControl.ToolTipText = "��ҳ"
        End With
    
        For Each cbrControl In cbrToolBar.Controls
            If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
                cbrControl.Style = xtpButtonIconAndCaption
            End If
        Next
    
         '�����
        With cbsThis.KeyBindings
            .Add FCONTROL, Asc("C"), conMenu_Edit_Copy
            .Add FCONTROL, Asc("V"), conMenu_Edit_PASTE
            .Add FCONTROL, Asc("D"), conMenu_Edit_SPECIALCHAR
            .Add FCONTROL, Asc("W"), conMenu_Edit_Word
        End With
    
    InitMenuBar = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lng����id As Long, ByVal lng��ҳid As Long, _
    ByVal strTime As String, ByVal strCurTime As String, ByRef strMsg As String) As Boolean
    Dim blnMsg As Boolean
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '���ݷ���ʱ������ڵ�ǰ���ҵ���Чʱ�䷶Χ��
    
    blnMsg = (strMsg <> "")
    
    '����ļ���ʼ,����ʱ��
    If strTime <= Format(mstr��ʼʱ��, "yyyy-MM-dd HH:mm") Then
        strMsg = "����ʱ�䲻��С���ļ���ʼʱ��[" & mstr��ʼʱ�� & "]"
        GoTo exitHand
    End If
    If mstr����ʱ�� <> "" Then
        If strTime > Format(mstr����ʱ��, "yyyy-MM-dd HH:mm") Then
            strMsg = "����ʱ�䲻�ܴ����ļ�����ʱ��[" & mstr����ʱ�� & "]"
            GoTo exitHand
        End If
    End If
    
    '���ݲ��˱䶯��¼���м��
    gstrSQL = " Select  /*+ RULE */ ��ʼԭ��,����ID,to_char(��ʼʱ��,'yyyy-MM-dd hh24:mi') AS ��ʼʱ��,to_char(NVL(��ֹʱ��,sysDate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS ��ֹʱ�� " & _
              " From ���˱䶯��¼ " & _
              " Where ����ID=[1] And ��ҳID=[2]" & _
              " Order by ��ʼʱ��,��ʼԭ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ǰ������Чʱ�䷶Χ", lng����id, lng��ҳid)
    With rsTemp
        .Filter = "����ID=" & mlng����ID
        Do While Not .EOF
            If strTime >= !��ʼʱ�� And strTime <= !��ֹʱ�� Then
                blnExist = True
                Exit Do
            End If
            .MoveNext
        Loop
        .Filter = 0
        '�ҵ��˾��˳�
        If blnExist Then
            If Not IsAllowInput(lng����id, lng��ҳid, strTime, strCurTime) Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[�������ݲ�¼����Чʱ��:" & glngHours & "Сʱ]"
                GoTo exitHand
            End If
            
            CheckTime = True
            Exit Function
        End If
        
        'û�ҵ�,������ԭ�����׼ȷ����ʾ
        .Filter = "��ʼԭ��=1"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 1 And strTime < !��ʼʱ�� Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ�����Ժʱ��:" & !��ʼʱ�� & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "��ʼԭ��=2"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 2 And strTime < !��ʼʱ�� Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ������ʱ��:" & !��ʼʱ�� & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "��ʼԭ��=10"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 10 And strTime > !��ֹʱ�� Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻�ܴ��ڳ�Ժʱ��:" & !��ֹʱ�� & "]"
                GoTo exitHand
            End If
        End If
        .Filter = 0
        '�������˵��
        strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[���ڵ�ǰ��������Чʱ�䷶Χ��]"
        GoTo exitHand
    End With
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
exitHand:
    rsTemp.Filter = 0
    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
End Function

Private Function CheckInput(strReturn As String, strInfo As String) As Boolean
    Dim i As Integer, j As Integer
    Dim strOrders As String, strText As String
    '���¼�����ݵĺϷ���(����Ҳ��Ϊ��һ���ַ�,���ǵ�������Ŀ�ȴ��ڲ���\�������Ϣ)
    '���ص�����,���һ�а󶨶����Ŀ,�Ե�������Ϊ�ָ���
    
    'mintType:0=�ı���¼��;1=��ѡ;2=��ѡ;3=ѡ��;4-Ѫѹ��һ�а���������Ŀ,���ʽ����Ѫѹ��������Ŀ;5=һ�а���������Ŀ�Ҿ���ѡ����Ŀ;
    '6=һ�а�N����Ŀ,�ֹ�¼��
    Select Case mintType
    Case 0
        strText = txtInput.Text
        strOrders = txtInput.Tag
    Case 1, 2   '���
        If mintType = 1 Then
            strText = Mid(lstSelect(mintType - 1).Text, 2)
        Else
            j = lstSelect(mintType - 1).ListCount
            For i = 1 To j
                If lstSelect(mintType - 1).Selected(i - 1) Then
                    strText = strText & "," & Mid(lstSelect(mintType - 1).List(i - 1), 2)
                End If
            Next
            If strText <> "" Then strText = Mid(strText, 2)
        End If
        strOrders = lstSelect(mintType - 1).Tag
    Case 4
        strText = txtUpInput.Text & "'" & txtDnInput.Text
        strOrders = txtUpInput.Tag & "'" & txtDnInput.Tag
    Case 6
        j = txt.Count
        For i = 1 To j
            strText = strText & "'" & txt(i - 1).Text
            strOrders = strOrders & "'" & txt(i - 1).Tag
        Next
        If strText <> "" Then
            strText = Mid(strText, 2)
            strOrders = Mid(strOrders, 2)
        End If
    Case 3      '���
        strText = lblInput.Caption
    Case 5      '���
        strText = lblUpInput.Caption & "/" & lblDnInput.Caption
    End Select
    If Val(strOrders) <> 0 Then
        If Not CheckValid(strText, strOrders, strInfo) Then Exit Function
    ElseIf VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then
        If Not CheckDateTime(strText, strInfo) Then Exit Function
    End If
    
    strReturn = strText
    CheckInput = True
End Function

Private Function CheckDateTime(strText As String, strInfo As String) As Boolean
    Dim blnCheck As Boolean
    Dim strCurrDate As String
    Dim strDate As String
    
    If VsfData.COL = mlngDate Then
        If mblnDateAd Then
            If Trim(strText) = "" Then
                strInfo = "���ڲ���Ϊ�գ�"
                Exit Function
            End If
            If InStr(1, strText, "/") = 0 Then
                strInfo = "���ڸ�ʽ������1��12�գ�12/01"
                Exit Function
            End If
            
            strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strText)
            If Not IsDate(strDate) Then
                strInfo = "¼������ݲ��ǺϷ������ڣ���1��12�գ�12/01"
                Exit Function
            End If
        Else
            If Trim(strText) = "" Then
                strInfo = "���ڲ���Ϊ�գ�"
                Exit Function
            End If
            If Not IsDate(strText) Then
                strInfo = "¼������ݲ��ǺϷ������ڣ���1��12�գ�2011-01-12"
                Exit Function
            End If
            strDate = Format(strText, "yyyy-MM-dd")
        End If
        If strDate > mstrMaxDate Then
            strInfo = "¼��������ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ��"
            Exit Function
        End If
        
        If VsfData.TextMatrix(VsfData.ROW, mlngTime) <> "" Then
            blnCheck = True
            strDate = strDate & " " & VsfData.TextMatrix(VsfData.ROW, mlngTime)
        End If
    Else
        If Trim(strText) = "" Then
            strInfo = "ʱ�䲻��Ϊ�գ�"
            Exit Function
        End If
        If Len(strText) <= 2 Then
            strText = String(2 - Len(strText), "0") & strText
            strText = strText & ":00"
        End If
        If Val(Mid(strText, 1, 2)) < 0 Or Val(Mid(strText, 1, 2)) > 23 Then
            strInfo = "¼���ʱ����Ч��СʱӦ����0-23֮�䣡"
            Exit Function
        End If
        If Mid(strText, 3, 1) <> ":" Then
            strInfo = "¼���ʱ���ʽ����[09:00]��"
            Exit Function
        End If
        If Len(strText) < 5 Then strText = strText & String(5 - Len(strText), "0")
        If Not (Val(Mid(strText, 4, 2)) >= 0 And Val(Mid(strText, 4, 2)) <= 59) Then
            strInfo = "¼���ʱ����Ч������Ӧ����0-59֮�䣡"
            Exit Function
        End If
        If Len(strText) > 5 Then
            strInfo = "¼���ʱ���ʽ����[09:00]��"
            Exit Function
        End If
        
        '���кϷ��Լ��
        If VsfData.TextMatrix(VsfData.ROW, mlngDate) <> "" Then
            strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            strDate = VsfData.TextMatrix(VsfData.ROW, mlngDate)
            If mblnDateAd Then
                strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
            Else
                strDate = Format(VsfData.TextMatrix(VsfData.ROW, mlngDate), "yyyy-MM-dd")
            End If
            strDate = strDate & " " & strText
            blnCheck = True
        End If
    End If
    
    If blnCheck Then
        '���ݷ���ʱ�䲻���ڵ�ǰ����Ա�������ҵ���Чʱ����ǰ
        If Not CheckTime(VsfData.ROW, mlng����ID, mlng��ҳID, strDate, strCurrDate, strInfo) Then
            Exit Function
        End If
    End If
    
    CheckDateTime = True
End Function

Private Function CheckValid(strReturn As String, ByVal strOrders As String, strInfo As String) As Boolean
    Dim arrData, arrOrder
    Dim i As Integer, j As Integer
    Dim dblMin As Double, dblMax As Double
    Dim strText As String, strName As String, strFormat As String
    
    '���и�ʽ��װ����
    mrsSelItems.Filter = "��=" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        '�д��е�δ���ж���
        strFormat = NVL(mrsSelItems!��ʽ)   '{P[����]C}{...}
    End If
    mrsSelItems.Filter = 0
    
    '�������
    arrData = Split(strReturn, "'")
    arrOrder = Split(strOrders, "'")
    j = UBound(arrData)
    For i = 0 To j
        strText = arrData(i)
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & strOrders
            strName = GetActivePart(VsfData.COL, i) & mrsItems!��Ŀ����
            If strText <> "" Then
                If mrsItems!��Ŀ���� = 0 And mrsItems!��Ŀ��ʾ = 0 Then
                    strText = Val(strText)
                    If NVL(mrsItems!��ĿС��, 0) <> 0 Then   '��������ͨ���ؼ���MaxLength�����Ƶ�
                        If InStr(1, strText, ".") <> 0 Then strText = Mid(strText, 1, InStr(1, strText, ".") - 1)
                        If Len(strText) > mrsItems!��Ŀ���� Then
                            mrsItems.Filter = 0
                            strInfo = "[" & strName & "]¼������ݳ����˺Ϸ����ȣ�"
                            Exit Function
                        End If
                        
                        strText = Val(arrData(i))
                        If InStr(1, strText, ".") <> 0 Then
                            strText = Mid(strText, InStr(1, strText, ".") + 1)
                            If Len(strText) > mrsItems!��ĿС�� Then
                                mrsItems.Filter = 0
                                strInfo = "[" & strName & "]¼���С�����ֳ����˺Ϸ����ȣ�"
                                Exit Function
                            End If
                        End If
                        strText = Val(arrData(i))
                    End If
                    If Not IsNull(mrsItems!��Ŀֵ��) Then
                        dblMin = Split(mrsItems!��Ŀֵ��, ";")(0)
                        dblMax = Split(mrsItems!��Ŀֵ��, ";")(1)
                        If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                            mrsItems.Filter = 0
                            strInfo = "[" & strName & "]¼������ݲ���" & Format(dblMin, "#0.00") & "��" & Format(dblMax, "#0.00") & "����Ч��Χ��"
                            Exit Function
                        End If
                    End If
                Else
                    If LenB(StrConv(strText, vbFromUnicode)) > mrsItems!��Ŀ���� Then
                        strInfo = "[" & strName & "]¼������ݳ�������󳤶ȣ�" & mrsItems!��Ŀ���� & "��"
                        mrsItems.Filter = 0
                        Exit Function
                    End If
                End If
                strFormat = Replace(strFormat, "[" & strName & "]", strText)
            Else
                'ɾ������Ŀ
                If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                    Call SubstrPro(strFormat, strName)
                Else
                    '����Ŀ������ʱ,�����ǰ�о��жԽ�������,�����
                    strFormat = Replace(strFormat, "[" & strName & "]", strText)
                End If
            End If
        Else
            strFormat = strReturn
        End If
    Next
    If j = -1 Then
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & strOrders
            strName = mrsItems!��Ŀ����
            strFormat = Replace(strFormat, "[" & strName & "]", strText)
        End If
    End If
    mrsItems.Filter = 0
    
    strFormat = Replace(strFormat, "{", "")
    strFormat = Replace(strFormat, "}", "")
    strReturn = strFormat
    CheckValid = True
End Function

Public Function SubstrVal(ByVal strData As String, ByVal strFormat As String, ByVal strName As String, intPos As Integer) As String
    Dim i As Integer, j As Integer, l As Integer, r As Integer
    Dim strQZ As String, strHZ As String
    '����ǰһ����Ŀ�ĺ�׺����+��ǰ��Ŀ��ǰ׺���ŵ�λ��
    
    If strData = "" Then Exit Function
    strData = UCase(strData)
    j = Len(strFormat)
    l = InStr(1, strFormat, "[" & strName & "]")
    If l = 0 Then Exit Function
    '�õ�ǰ׺
    For i = l To 1 Step -1
        If Mid(strFormat, i, 1) = "{" Then Exit For
    Next
    strQZ = Mid(strFormat, i + 1, l - i - 1)
    '�ҵ�����Ŀ��ʽ���еĽ�������
    i = l + Len(strName) + 2
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    '�õ���׺
    strHZ = Mid(strFormat, i, r - i)
    '�����׺Ϊ��,�������Ѱ����һ����Ŀ��ǰ׺����
    If strHZ = "" And r < j Then
        For r = r + 1 To j
            If Mid(strFormat, r, 1) = "[" Then Exit For
        Next
        strHZ = Mid(strFormat, InStr(i, strFormat, "{") + 1, r - InStr(i, strFormat, "{") - 1)
    End If
    'ȡ��ָ����Ŀ���������ݴ�
    If strHZ <> "" Then
        j = InStr(intPos, strData, strHZ) '��Ϊ������ȡ��,���ǵ��ָ���������ͬ�����,��¼��һ�ε����λ��,�´δ����λ������ȡ����
        If j = 0 Then
            '�п����м���ڻس����з�
            j = InStr(intPos, Replace(strData, vbCrLf, ""), strHZ)
            If j = 0 Then Exit Function
        End If
    End If
    strData = Mid(strData, intPos)
    'ǰ׺Ϊ��,������ǰѰ����һ����Ŀ�ĺ�׺����
'    If strQZ = "" And i > 1 And intPos > 1 Then
'        For i = i - 1 To 1 Step -1
'            If Mid(strFormat, i, 1) = "]" Then Exit For
'        Next
'        strQZ = Mid(strFormat, i + 1, InStr(i, strFormat, "}") - i - 1)
'    End If
    
    SubstrVal = SubstrAnaly(strData, strHZ, strQZ)
    intPos = intPos + Len(strQZ & SubstrVal & strHZ)
    '�������������ȥ���س����з�����,������ַ�����ԭ������
'    If strHZ <> "" Then
'
'        strData = Mid(strData, 1, InStr(1, Replace(strData, vbCrLf, ""), strHZ) - 1) '��������Ŀ�������
'        intPOS = i + Len(strHZ)
'    End If
'    If strQZ <> "" Then strData = Mid(strData, InStr(1, strData, strQZ) + Len(strQZ)) '��������Ŀ�������
'    SubstrVal = strData ' Replace(strData, vbCrLf, "")
End Function

Private Function SubstrAnaly(ByVal strData As String, ByVal strHZ As String, ByVal strQZ As String) As String
    Dim strText As String
    Dim strCompare As String           '�Աȴ�
    Dim intLen As Integer, intActLen As Integer           'ǰ׺/��׺�ĳ���
    Dim intPos As Integer, intEnd As Integer
    Dim lngASC As Long
    Dim blnFind As Boolean
    '�����س����з�����,�ո����±ȶ�
    
    strText = strData
    If strHZ <> "" Then
        '�Ѻ�׺ȥ��
        strHZ = Replace(strHZ, vbCrLf, "")
        intEnd = Len(strText)
        intLen = Len(strHZ)
        For intPos = 1 To intEnd
            lngASC = Asc(Mid(strText, intPos, 1))
            intActLen = intActLen + 1
            If Not (lngASC = 13 Or lngASC = 10) Then
                If lngASC = 32 Then
                    strCompare = ""
                    intActLen = 0
                Else
                    strCompare = strCompare & Mid(strText, intPos, 1)
                End If
                If Len(strCompare) = intLen Then
                    If strCompare = strHZ Then
                        blnFind = True
                        intPos = intPos - intActLen
                    Else
                        strCompare = ""
                        intPos = intPos - intActLen + 1
                        intActLen = 0
                    End If
                End If
            End If
            If blnFind Then Exit For
        Next
        '�϶���
        strText = Mid(strText, 1, intPos)
    End If
    
    '��ȥ��ǰ׺
    If strQZ <> "" Then
        If InStr(1, strText, strQZ) = 0 Then strText = strQZ & strText
        strQZ = Replace(strQZ, vbCrLf, "")
        intEnd = Len(strText)
        intLen = Len(strQZ)
        strCompare = ""
        intActLen = 0
        blnFind = False
        For intPos = 1 To intEnd
            lngASC = Asc(Mid(strText, intPos, 1))
            intActLen = intActLen + 1
            If Not (lngASC = 13 Or lngASC = 10) Then
                If lngASC = 32 Then
                    strCompare = ""
                    intActLen = 0
                Else
                    strCompare = strCompare & Mid(strText, intPos, 1)
                End If
                If Len(strCompare) = intLen Then
                    If strCompare = strQZ Then
                        blnFind = True
                        intPos = intPos + 1
                    Else
                        strCompare = ""
                        intPos = intPos + 1
                        intActLen = 0
                    End If
                End If
            End If
            If blnFind Then Exit For
        Next
        strText = Mid(strText, intPos)
    End If
    
    If IsNumeric(Replace(strText, vbCrLf, "")) Then
        SubstrAnaly = Replace(strText, vbCrLf, "")
    Else
        SubstrAnaly = strText
    End If
End Function

Public Sub SubstrPro(strFormat As String, ByVal strName As String, Optional ByVal intType As Integer = 0)
    Dim i As Integer, j As Integer, l As Integer, r As Integer
    'intType=0-ɾ��ָ����ʽ��;1-�õ�ָ����ʽ��
    j = Len(strFormat)
    i = InStr(1, strFormat, "[" & strName & "]")
    If i = 0 Then Exit Sub
    
    For l = i To 1 Step -1
        If Mid(strFormat, l, 1) = "{" Then Exit For
    Next
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    If intType = 0 Then
        strFormat = Mid(strFormat, 1, l - 1) & Mid(strFormat, r + 1)
    Else
        strFormat = Mid(strFormat, l, r - l + 1)
    End If
End Sub

Private Sub MoveNextCell()
    Dim arrData
    Dim blnNULL As Boolean                      '�Ƿ�Ϊ����
    Dim strReturn As String, strMsg As String, strPart As String
    Dim lngStart As Long, lngMutilRows As Long, lngDeff As Long
    Dim intRow As Integer, intCount As Integer, intNULL As Integer  '����ж��ٿ���
    '��ֵȻ���ƶ�����һ����Ч��Ԫ��
    
    '�������,���ϸ���ٴε���Ҫ��¼��
    If mintType >= 0 Then
        If Not CheckInput(strReturn, strMsg) Then
            RaiseEvent AfterRowColChange(strMsg, True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        lngMutilRows = 1
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
            lngMutilRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        End If
        lngStart = GetStartRow(VsfData.ROW)
        
        '׼����ֵ
        With txtLength
            '������ʱ���еĿ�Ȳ���,Ϊ�˱��ⷵ�ض���,ǿ������Ϊ5000
            .Width = IIf(VsfData.COL = mlngDate Or VsfData.COL = mlngTime, 5000, VsfData.CellWidth)
            .Text = strReturn
            .FontName = VsfData.CellFontName
            .FontSize = VsfData.CellFontSize
        End With
        arrData = GetData(txtLength.Text)
        intCount = UBound(arrData)
        
        If intCount > lngMutilRows - 1 Then
            '������������,�������������������������ӵ�����
            intNULL = intCount - (lngMutilRows - 1)
            For intRow = lngMutilRows To intCount
                '��֤��ǰ�����������һҳ����ʾȫ
                If intRow + lngStart > VsfData.Rows - 1 Then Exit For
                
                If Val(VsfData.TextMatrix(intRow + lngStart, mlngRecord)) = 0 And VsfData.TextMatrix(intRow + lngStart, mlngRowCount) = "" Then
                    intNULL = intNULL - 1
                Else
                    Exit For
                End If
            Next
            '�����ӿ���
            If intNULL > 0 Then
                lngDeff = intNULL
                VsfData.Rows = VsfData.Rows + intNULL
                '�ӵ�ǰ�м�¼�Ŀհ��п�ʼ��ÿ�е�λ��+�����ӵĿհ�����
                For intRow = VsfData.Rows - intNULL - 1 To lngStart + intCount - intNULL + 1 Step -1
                    VsfData.RowPosition(intRow) = intRow + intNULL
                Next
            End If
            'ѭ����ֵ
            intCount = UBound(arrData)
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = arrData(intRow)
                VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intCount + 1 & "|" & intRow + 1
                VsfData.TextMatrix(lngStart + intRow, mlngRowCurrent) = intCount + 1
            Next
            '���������н��и�ֵ
            lngMutilRows = lngStart + intCount
            For intRow = lngStart + 1 To lngMutilRows
                For intCount = 0 To VsfData.Cols - 1
                    VsfData.Cell(flexcpForeColor, intRow, intCount) = VsfData.Cell(flexcpForeColor, lngStart, intCount)
                    If VsfData.ColHidden(intCount) And InStr(1, "," & mlngRowCount & "," & mlngRowCurrent & ",", "," & intCount & ",") = 0 Then
                        VsfData.TextMatrix(intRow, intCount) = VsfData.TextMatrix(lngStart, intCount)
                    End If
                Next
            Next
        Else
            '�Ը������¸�ֵ����ֻ����һ������ʱ����֪Ϊ�λ�����ַ�ASCII��Ϊ1�ķ��ţ�
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(arrData(intRow), Chr(1), "")
            Next
            For intRow = intCount + 1 To lngMutilRows - 1
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = ""
            Next
            
            '����������������д������,intNULL��¼���һ����Ϊ���е��к�
            intNULL = lngStart + lngMutilRows - 1
            For intRow = lngMutilRows To 1 Step -1
                blnNULL = True
                For intCount = 0 To VsfData.Cols - 1
                    If Not VsfData.ColHidden(intCount) Then
                        If VsfData.TextMatrix(intRow + lngStart - 1, intCount) <> "" Then
                            blnNULL = False
                            Exit For
                        End If
                    End If
                Next
                
                If Not blnNULL Then Exit For
                intNULL = intNULL - 1
            Next
            '������д�����
            For intRow = lngStart To intNULL
                VsfData.TextMatrix(intRow, mlngRowCount) = (intNULL - lngStart + 1) & "|" & intRow - lngStart + 1
                VsfData.TextMatrix(intRow, mlngRowCurrent) = (intNULL - lngStart + 1)
            Next
            For intRow = intNULL + 1 To lngStart + lngMutilRows - 1
                VsfData.TextMatrix(intRow, mlngRowCount) = ""
                VsfData.TextMatrix(intRow, mlngRowCurrent) = ""
            Next
        End If
        
        '���кŷ����仯����ͬ������mrsCellMap�д��ڸ��кŵ��к�����
        If lngDeff <> 0 Then Call CellMap_Update(lngStart, lngDeff)
        
        If mstrData <> strReturn Then
            mblnChange = True
            
            'ͬ������������ʱ���е�����
            Dim strKey As String, strField As String, strValue As String
            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
            '1\����
            If mlngDate <> -1 Then
                strKey = mintҳ�� & "," & lngStart & "," & mlngDate
                strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngDate & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.TextMatrix(lngStart, mlngDate) & "|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
            '2\ʱ��
            strKey = mintҳ�� & "," & lngStart & "," & mlngTime
            strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngTime & "|" & _
                Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.TextMatrix(lngStart, mlngTime) & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            
            '��¼�û��޸Ĺ��ĵ�Ԫ��
            If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                strPart = GetActivePart(VsfData.COL, 0)
            Else
                strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
            End If
            
            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
            strKey = mintҳ�� & "," & lngStart & "," & VsfData.COL
            strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & VsfData.COL & "|" & _
                Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(strReturn = "", "1", "0")
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    End If
    
toNextCol:
    If VsfData.COL < mlngNoEditor - 1 Then       '�����¼���϶��л�ʿǩ����
        VsfData.COL = VsfData.COL + 1
        If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Then GoTo toNextCol
    Else
toNextRow:
        '������һ��
        intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
        If VsfData.ROW + intRow < VsfData.Rows Then
            VsfData.ROW = VsfData.ROW + intRow
        End If
        If VsfData.RowHidden(VsfData.ROW) Then GoTo toNextRow
        VsfData.COL = IIf(mlngDate > 0, mlngDate, mlngTime)
    End If
    If VsfData.ColIsVisible(VsfData.COL) = False Then
        VsfData.LeftCol = VsfData.COL
    End If
    If VsfData.RowIsVisible(VsfData.ROW) = False Then
        VsfData.TopRow = VsfData.ROW
    End If
End Sub

Private Function GetStartRow(ByVal lngRow As Long) As Long
    Dim lngStart As Long
    Dim lngCurRows As Long, lngRows As Long
    '��ȡ������ʼ��,������ҳ�򷵻�0
    '�����ҳδ��ʾȫ,��˵��������ҳ,Ҳ����0
    '���������������������в�������
    
    lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))    '������
    lngCurRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1)) '��ǰ��
    If lngCurRows = 1 Then
        GetStartRow = lngRow
        Exit Function
    End If
    
    'Ѱ����ʼ��
    For lngRow = lngRow To 3 Step -1
        If VsfData.TextMatrix(lngRow, mlngRowCount) = lngRows & "|1" Then
            lngStart = lngRow
            Exit For
        End If
    Next
    
    GetStartRow = lngStart
End Function

Private Function GetMutilData(ByVal lngRow As Long, ByVal lngCOL As Long, dblTop As Long, dblHeight As Long) As String
    Dim lngCurRow As Long
    Dim lngCount As Long
    Dim lngStart As Long    '��ʼ��
    Dim lngRecordId As Long
    Dim strReturn As String
    Dim blnAdjust As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '���ص�һ�е�����
    '������ֱ��ȡ������ʱ��������ҳ��ʾȫ��ƴ�ӣ�����ӿ��ж�ȡ
    
    If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then
        GetMutilData = VsfData.TextMatrix(lngRow, lngCOL)
        Exit Function
    End If
    lngRecordId = Val(VsfData.TextMatrix(lngRow, mlngRecord))
    lngCount = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
    lngCurRow = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1))
    
    If lngCount > 1 Then
        lngStart = GetStartRow(lngRow)
    Else
        lngStart = lngRow
    End If
    If lngRecordId <> 0 And (lngStart = 0 Or lngStart + lngCount > VsfData.Rows) Then   'ҳ��Ч��=�̶�������+��ͷ
        '�����ݿ�����ȡ
        Call SQLCombination(lngRecordId)
        gstrSQL = mstrSQL
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mintҳ��, lngRecordId)
        strReturn = NVL(rsTemp.Fields(lngCOL).Value)
        If lngStart = 0 Then lngStart = 3       '���δ�ҵ���ʼ�����趨Ϊ��1��
        blnAdjust = True
    Else
        For lngRow = lngStart To lngStart + lngCount - 1
            strReturn = strReturn & VsfData.TextMatrix(lngRow, lngCOL)
        Next
    End If
    
'    'У���и�(�п���ʵ������ռ5�ж���ǰҳ��ֻ��ʾ��3��,����3����ʾ�������Բ�ȫ,���Ի�����ԭ�����и���ʾ����,���´�������)
'    If blnAdjust Then
'        If lngStart = 3 Then
'            lngCurRow = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(1))
'            lngCount = lngCount - lngCurRow + 1
'        Else
'            lngCount = mlngPageRows +mlngOverrunRows + VsfData.FixedRows - lngStart
'        End If
'    End If
    'ȡ�и�
    VsfData.ROW = lngStart
    dblHeight = lngCount * VsfData.RowHeightMin + 20
    dblTop = VsfData.Top + VsfData.CellTop
    
    GetMutilData = strReturn
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowInput(Optional ByVal intCol As Integer = -1, Optional ByVal strCellData As String = "", Optional ByVal blnAnalyse As Boolean = False) As String
    Dim arrData, arrValue
    Dim lngOrder As Long
    Dim i As Integer, j As Integer, intPos As Integer, intIndex As Integer
    Dim strFormat As String, strText As String, strValue As String  '��ʽ��,���ݴ�,��ֵ��
    Dim strOrders As String, strTypes As String, strBounds As String, strLen As String, strName As String
    Const txtHeight = 300
    On Error GoTo errHand
    
    '�����ļ��������ģ����Ҫ����:
    '1��һ�а�һ����Ŀ�Ĳ��ù�
    '2��һ�а�������Ŀ�ģ�Ѫѹ����ɶԣ�Ҫô����¼�룬Ҫô����ѡ�񣬲���������֣�Ҳ��������ֵ�ѡ����ѡ
    '3��һ�а󶨶����Ŀ�ģ�ֻ����¼����Ŀ
    '���������������ƣ�ֻȡ��һ����Ŀ�����ʼ���
    
    '����Ǳ��洦�����������´���
    If intCol = -1 Then intCol = VsfData.COL
    If blnAnalyse Then
        strText = strCellData
    Else
        'ȡ��ǰ��Ԫ�������
        CellRect.Left = VsfData.CellLeft + VsfData.Left
        CellRect.Top = VsfData.CellTop + VsfData.Top
        CellRect.Bottom = VsfData.CellHeight + 20
        CellRect.Right = VsfData.CellWidth + 20
        strText = GetMutilData(VsfData.ROW, intCol, CellRect.Top, CellRect.Bottom)
    End If
    mstrData = strText
    mintType = 0
    intIndex = 0
    
    'ȡ��ǰ�еİ���Ŀ
    intPos = 1
    mrsSelItems.Filter = "��=" & intCol - cHideCols
    Do While Not mrsSelItems.EOF
        lngOrder = mrsSelItems!��Ŀ���
        If lngOrder = 0 Then
            strLen = 0
            strValue = strText
            Exit Do
        End If
        
        '��Ŀ��ʾ:2��ѡ;3-��ѡ;4-����;5-ѡ��
        '��Ŀֵ��:��Ŀ��ʾΪ0-��ʾ��Сֵ;���ֵ;��Ŀ��ʾΪ2,3-��ʾ��ĿA;��ĿB,ǰ�й��ı�ʾȱʡ��
        strFormat = UCase(NVL(mrsSelItems!��ʽ))
        strOrders = strOrders & "," & lngOrder
        If lngOrder <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & lngOrder
            strName = strName & "," & mrsItems!��Ŀ����
            strLen = strLen & "," & mrsItems!��Ŀ���� & ";" & NVL(mrsItems!��ĿС��)
            strTypes = strTypes & "," & mrsItems!��Ŀ��ʾ
            strBounds = strBounds & "," & mrsItems!��Ŀֵ��
            strValue = strValue & "'" & SubstrVal(strText, strFormat, GetActivePart(intCol, intIndex) & mrsItems!��Ŀ����, intPos)
            
            Select Case mrsItems!��Ŀ��ʾ
            Case 0  '�ı�¼����
                If mrsSelItems.RecordCount = 2 Then
                    mintType = 4
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 2  '��ѡ
                mintType = 1
            Case 3  '��ѡ
                mintType = 2
            Case 4  '����
                If mrsSelItems.RecordCount = 2 Then
                    mintType = 4
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 5  'ѡ��
                If mrsSelItems.RecordCount = 1 Then
                    mintType = 3
                Else
                    mintType = 5
                End If
            End Select
        Else
            strTypes = strTypes & ","
            strBounds = strBounds & ","
            strLen = strLen & ","
            strName = strName & ","
        End If
        
        intIndex = intIndex + 1
        mrsSelItems.MoveNext
    Loop
    If strOrders <> "" Then
        strOrders = Mid(strOrders, 2)
        strName = Mid(strName, 2)
        strLen = Mid(strLen, 2)
        strTypes = Mid(strTypes, 2)
        strBounds = Mid(strBounds, 2)
        strValue = Mid(strValue, 2)
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    If blnAnalyse Then
        ShowInput = strOrders & "||" & strValue
        Exit Function
    End If
    
    '���4����У��,�����ͷ�ı�����/����Ϊ6
    If mintType = 4 Then
        If Not IsDiagonal(intCol) Then
            mintType = 6
        End If
    End If
    
    '�жϵ�ǰ�е�����
    'mintType:0=�ı���¼��;1=��ѡ;2=��ѡ;3=ѡ��;4-Ѫѹ��һ�а���������Ŀ,���ʽ����Ѫѹ��������Ŀ;5=һ�а���������Ŀ�Ҿ���ѡ����Ŀ;
    '6=һ�а�2����������Ŀ,�ֹ�¼��
    arrValue = Split(strValue & "'", "'")
    Select Case mintType
    Case 0, 3
        With picInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
        If mintType = 0 Then
            txtInput.Visible = True
            If Val(strLen) <> 0 And Val(strOrders) <> 10 Then
                txtInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
            Else
                txtInput.MaxLength = 0
            End If
            txtInput.Tag = lngOrder
        Else
            txtInput.Visible = False
        End If
        With txtInput
            .Top = 0
            .Text = strValue
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = .Width - (180 + IIf(mblnBlowup, 180 * 1 / 3, 0)) / 2 '����9��ʱ��ȥ90,����Խ��۳��ı߾�ԽС,�Ա�֤�ı��������ʵ��һ��
        End With
        With lblInput
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = CellRect.Bottom
            .Width = CellRect.Right
            .Top = 50
            .Tag = lngOrder
            .Caption = strValue
            .Visible = (mintType = 3)
        End With
        
        '��������ڻ�ʱ���У��趨�̶�ֵ
        If mintType = 0 And txtInput.Text = "" Then
            If intCol = mlngDate Then
                If mblnDateAd Then
                    txtInput.Text = Format(zlDatabase.Currentdate, "d-M")
                    txtInput.Text = Replace(txtInput.Text, "-", "/")
                Else
                    txtInput.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
                End If
            ElseIf intCol = mlngTime Then
                txtInput.Text = Format(zlDatabase.Currentdate, "HH:mm")
            End If
        End If
    Case 1, 2
        '��������
        lstSelect(mintType - 1).Clear
        arrData = Split(strBounds, ";")
        j = UBound(arrData)
        For i = 0 To j
            If arrData(i) <> "" Then
                If Mid(arrData(i), 1, 1) = "��" Then
                    lstSelect(mintType - 1).AddItem i + 1 & Mid(arrData(i), 2)
                    If strText = "" Then lstSelect(mintType - 1).ListIndex = i
                Else
                    lstSelect(mintType - 1).AddItem i + 1 & arrData(i)
                End If
            End If
        Next
        '��ѡ����¼�����ݵ������
        If strValue <> "" Then
            strValue = Replace(strValue, vbCrLf, "")
            For i = 0 To j
                If InStr(1, "," & strValue & ",", "," & Mid(lstSelect(mintType - 1).List(i), 2) & ",") <> 0 Then
                    lstSelect(mintType - 1).Selected(i) = True
                End If
            Next
        End If
        '��ʾ
        With lstSelect(mintType - 1)
            .Left = CellRect.Left
            .Top = CellRect.Top
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = .ListCount * 300
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '���м���ĳ���Ϊ����
            If .Width < CellRect.Right Then .Width = CellRect.Right
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Tag = lngOrder
            .Visible = True
        End With
    Case 4, 5
        With picDouble
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Height = CellRect.Bottom
            If .Height < 280 Then .Height = 280
            .Width = CellRect.Right
            If .Width < 820 Then .Width = 820
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
        lblSplit.FontName = VsfData.FontName
        lblSplit.FontSize = VsfData.FontSize
        lblSplit.Left = (picDouble.Width - lblSplit.Width) / 2
        If mblnBlowup Then
            lblSplit.Width = 150
        Else
            lblSplit.Width = 105
        End If
        
        With txtUpInput
            .Text = arrValue(0)
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = (picDouble.Width - lblSplit.Width) * 0.4
            .ZOrder IIf(mintType = 4, 0, 1)
            .Locked = Not (mintType = 4)
            .Tag = Split(strOrders, ",")(0)
        End With
        With picUpInput
            .Left = txtUpInput.Left
            .Width = txtUpInput.Width
            .Height = CellRect.Bottom
            .ZOrder IIf(mintType = 5, 0, 1)
            .Tag = Split(strOrders, ",")(0)
        End With
        With lblUpInput
            .Alignment = 2
            .Caption = arrValue(0)
            .Left = 0
            .Top = 50
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = txtUpInput.Width
            .Height = CellRect.Bottom
            .Tag = Split(strOrders, ",")(0)
        End With
        With txtDnInput
            .Text = arrValue(1)
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Left = lblSplit.Left + lblSplit.Width
            .Width = picDouble.Width - .Left
            .ZOrder IIf(mintType = 4, 0, 1)
            .Locked = Not (mintType = 4)
            .Tag = Split(strOrders, ",")(1)
        End With
        With picDnInput
            .Left = txtDnInput.Left
            .Height = CellRect.Bottom
            .Width = txtDnInput.Width
            .ZOrder IIf(mintType = 5, 0, 1)
            .Tag = Split(strOrders, ",")(1)
        End With
        With lblDnInput
            .Alignment = 2
            .Caption = arrValue(1)
            .Left = 0
            .Top = 50
            .Height = CellRect.Bottom
            .Width = txtDnInput.Width
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Tag = Split(strOrders, ",")(1)
        End With
        
        If mintType = 4 Then
            If strLen <> "" Then txtUpInput.MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
            If strLen <> "" Then txtDnInput.MaxLength = Val(Split(Split(strLen, ",")(1), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(1), ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
        End If
    Case 6
        '��ɾ����ǰ�Ŀؼ�
        j = txt.Count - 1
        For i = 1 To j
            Unload lbl(i)
            Unload txt(i)
        Next
        '�趨����
        With picMutilInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = IIf(CellRect.Right < 1600, 1600, CellRect.Right)
        End With
        '��ȱʡ�ؼ���ֵ
        arrData = Split(strOrders, ",")
        j = UBound(arrData)
        lbl(0).Top = 130
        lbl(0).Caption = Split(strName, ",")(0)
        lbl(0).FontName = VsfData.FontName
        lbl(0).FontSize = VsfData.FontSize
        txt(0).Tag = arrData(0)
        txt(0).FontName = VsfData.FontName
        txt(0).FontSize = VsfData.FontSize
        txt(0).Width = picMutilInput.Width - txt(0).Left - 100
        txt(0).MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1)  'С��λ��Ҫ����С����
        txt(0).Text = arrValue(0)
        If Not mblnBlowup Then
            txt(0).Height = 225
        End If
        
        '���ؿؼ�
        For i = 1 To j
            Load lbl(i)
            With lbl(i)
                .Caption = Split(strName, ",")(i)
                .Left = lbl(0).Left + lbl(0).Width - .Width
                .Top = lbl(i - 1).Top + txtHeight + IIf(mblnBlowup, txtHeight * 1 / 3, 0)
                .Visible = True
            End With
            Load txt(i)
            With txt(i)
                .TabIndex = txt(i - 1).TabIndex + 1
                .Left = txt(0).Left
                .Top = txt(i - 1).Top + txtHeight + IIf(mblnBlowup, txtHeight * 1 / 3, 0)
                .Tag = arrData(i)
                If strLen <> "" Then
                    .MaxLength = Val(Split(Split(strLen, ",")(i), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(i), ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
                End If
                .Text = arrValue(i)
                .Visible = True
            End With
        Next
        
        With picMutilInput
            .Height = txt(j).Top + txt(j).Height + 120
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
    End Select
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CheckFormat(ByVal strNames As String, ByVal strFormat As String)
    '�����ʽ��Ѫѹ�ķ�ʽ��ͬ,����ʽ����Ϊ6
    
    'ȥ��ǰ׺����жԱ�
    strFormat = Mid(strFormat, InStr(1, strFormat, "["))
    strFormat = Replace(strFormat, "[", "")
    strFormat = Replace(strFormat, "]", "")
    If Not (strFormat Like Split(strNames, ",")(0) & "/}*" Or strFormat Like "{/*" & Split(strNames, ",")(1)) Then
        mintType = 6
    End If
End Sub

Private Function IsDiagonal(ByVal intCol As Integer) As Boolean
    Dim arrCol, arrData
    Dim intDo As Integer, intCount As Integer
    '�ж�ָ�����Ƿ��������жԽ��ߣ�mstrColWidth�ĸ�ʽ��765`11`1`1,765`11`2`1,...����������`�������`�жԽ��ߣ�
    
    IsDiagonal = (InStr(1, "," & mstrCatercorner & ",", "," & intCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0)
End Function

Private Sub ISAssistant(ByVal lngOrder As Long, ByVal objTXT As TextBox)
    Dim intIndex As Integer
    Dim objParent As Object
    '������Ŀ�ĳ��Ⱦ����Ƿ�������дʾ�ѡ��
    mblnEditAssistant = False
    cmdWord.Visible = mblnEditAssistant
    
    mrsItems.Filter = "��Ŀ���=" & lngOrder
    If mrsItems.RecordCount = 0 Then
        mrsItems.Filter = 0
        Exit Sub
    End If
    mblnEditAssistant = (mrsItems!��Ŀ���� > 100)
    mrsItems.Filter = 0
    
    '�������ʾ�ѡ��,��ʾ����λ
    If mblnEditAssistant Then
        If UCase(objTXT.Name) = "TXTINPUT" Then
            intIndex = -1 '��ʾtxtInput
            Set objParent = picInput
        Else
            intIndex = objTXT.Index
            Set objParent = picMutilInput
        End If
        With cmdWord
            .Tag = intIndex
            .Top = objParent.Top + objTXT.Top + 25
            .Left = objParent.Left + objTXT.Left + objTXT.Width - .Width + 25
            .Visible = True
        End With
    End If
End Sub

Private Sub FillPage()
    Dim lngRow As Long, lngRows As Long, lngCount As Long
    '��֤ÿҳ��Ч������
    
    lngRows = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngRows
        If Not VsfData.RowHidden(lngRow) Then lngCount = lngCount + 1
    Next
    
    If lngCount < mlngPageRows + mlngOverrunRows Then VsfData.Rows = VsfData.Rows + (mlngPageRows + mlngOverrunRows - lngCount)
End Sub

Public Function GetSynItems(ByVal intType As Integer, ByRef intMax As Integer) As String
    Dim arrCols
    Dim strItems As String
    Dim strCols As String
    Dim strNames As String
    Dim lngRecord As Long, lngStartRow As Long, lngEndRow As Long
    Dim intIn As Integer, intOut As Integer, intInMAX As Integer, intOutMax As Integer, intCount As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    'intType��ȷ����������ֵ��1)������Ŀ���;2)�����к�
    'intMAX������ͬ����������ռ�õ��и�
    '����ͬ��������(һ���ļ��в����ܳ����ظ�����Ŀ,����,�ж�ʱ���ؼ���к�)
    
    lngRecord = Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord))
    If lngRecord = 0 Then Exit Function
    
    gstrSQL = "" & _
        " SELECT /*+ RULE */ B.��Ŀ���,B.��Ŀ����,A.������� AS �к�" & vbNewLine & _
        " FROM �����ļ��ṹ A,���˻�����ϸ B" & vbNewLine & _
        " WHERE A.Ҫ������=B.��Ŀ���� AND A.��ID=" & vbNewLine & _
        "      (SELECT A.ID FROM �����ļ��ṹ A,���˻����ļ� B " & vbNewLine & _
        "       WHERE B.ID=[2] And A.�ļ�ID=B.��ʽID AND A.�������=4 AND A.��ID IS NULL)" & vbNewLine & _
        " AND B.������Դ>0 AND B.��¼ID=[1]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ͬ��������", lngRecord, mlng�ļ�ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    
    '��ȡͬ�������Ϣ
    Do While Not rsTemp.EOF
        If InStr(1, "," & strCols & ",", "," & rsTemp!�к� & ",") = 0 Then strCols = strCols & "," & rsTemp!�к�
        strItems = strItems & "," & rsTemp!��Ŀ���
        strNames = strNames & "," & rsTemp!��Ŀ����
        rsTemp.MoveNext
    Loop
    strCols = Mid(strCols, 2)
    strItems = Mid(strItems, 2)
    strNames = Mid(strNames, 2)
    
    '������ѭ�����������ռ�и�
    If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
        lngStartRow = VsfData.ROW
        lngEndRow = VsfData.ROW
        intInMAX = 1
    Else
        lngStartRow = GetStartRow(VsfData.ROW)
        intInMAX = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngEndRow = lngStartRow + intInMAX - 1
    End If
    
    intCount = 1    'ͬ����ֻ������������Ŀ������ռ����ֻ������1�У��������ݲ�����Ҫ���
'    '����ռ�ó���1�вż��
'    If intInMAX > 1 Then
'        arrCols = Split(strCols, ",")
'        intOutMax = UBound(arrCols)
'        For intOut = 0 To intOutMax
'            For intIn = 2 To intInMAX
'                If VsfData.TextMatrix(intIn + lngStartRow - 1, arrCols(intOut) + 1) <> "" Then
'                    If intIn > intCount Then intCount = intIn
'                End If
'            Next
'        Next
'    End If
    
    intMax = intCount
    GetSynItems = IIf(intType = 1, strItems, strCols)
    If strNames <> "" Then
        RaiseEvent AfterRowColChange("������,ʱ����,�Լ� " & strNames & " ��ͬ�����������ݣ��������޸Ļ�ɾ����", True, mblnSign, mblnArchive)
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ISColHaveData() As Boolean
    Dim arrData
    Dim arrCol
    Dim intCol As Integer
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    Dim strCond As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '�����ݿ�����ȡ���ݣ������ǰ���Ŀ�д�������������������Ŀ����
    
    '�����Ŀ���뵽��ѯSQL�У���ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...
    '�󶨶����Ŀ�����о��Զ�תΪ�Խ�����
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        If intCol = VsfData.COL - cHideCols - VsfData.FixedCols + 1 Then
            arrCol = Split(Split(arrData(intDo), "|")(1), ";")
            intMax = UBound(arrCol)
            For intIn = 0 To intMax
                strCond = strCond & " OR (��Ŀ���=" & Split(arrCol(intIn), ",")(0)
                If Split(arrCol(intIn), ",")(1) = "" Then
                    strCond = strCond & ")"
                Else
                    strCond = strCond & " AND NVL(���²�λ,'TWBW')='" & Split(arrCol(intIn), ",")(1) & "')"
                End If
            Next
            
            Exit For
        End If
    Next
    
    If strCond <> "" Then
        strCond = " AND (" & Mid(strCond, 4) & ")"
        '��ѯ���ݿ�
        gstrSQL = " SELECT /*+ RULE */ 1 FROM ���˻�����ϸ A,���˻������� B,���˻����ӡ C" & vbNewLine & _
                  " Where A.��¼ID=B.ID And B.�������=0 And B.ID=C.��¼ID And C.�ļ�ID=B.�ļ�ID " & vbNewLine & _
                  " And C.�ļ�ID=[1] And (C.����ҳ��=[2] OR C.��ʼҳ��=[2])" & strCond & " AND ROWNUM<2"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ���ݿ⵱ǰҳ��ָ������Ƿ���ڻ��Ŀ", mlng�ļ�ID, mintҳ��)
        ISColHaveData = rsTemp.RecordCount
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


'######################################################################################################################
'**********************************************************************************************************************
'�����ǻ������������
Private Sub lblDnInput_Click()
    txtDnInput.SetFocus
End Sub

Private Sub lblUpInput_Click()
    txtUpInput.SetFocus
End Sub

Private Sub lstColumnItems_DblClick()
    Call cmdColumn_Click(0)
End Sub

Private Sub lstColumnItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lstColumnItems_DblClick
End Sub

Private Sub lstColumnUsed_DblClick()
    Call cmdColumn_Click(1)
End Sub

Private Sub lstColumnUsed_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lstColumnUsed_DblClick
End Sub

Private Sub lstSelect_DblClick(Index As Integer)
    Call lstSelect_KeyDown(Index, vbKeyReturn, 0)
End Sub

Private Sub lstSelect_GotFocus(Index As Integer)
    mblnEditAssistant = False
End Sub

Private Sub txtColumnNo_GotFocus()
    txtColumnNo.SelStart = 0
    txtColumnNo.SelLength = 100
End Sub

Private Sub txtColumnNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtDnInput_GotFocus()
    txtDnInput.SelStart = 0
    txtDnInput.SelLength = 100
    Call ISAssistant(Val(txtDnInput.Tag), txtDnInput)
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = 100
    mintSymbol = -1
    Call ISAssistant(Val(txtInput.Tag), txtInput)
End Sub

Private Sub txtUpInput_GotFocus()
    txtUpInput.SelStart = 0
    txtUpInput.SelLength = 100
    Call ISAssistant(Val(txtUpInput.Tag), txtUpInput)
End Sub

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).SelStart = 0
    txt(Index).SelLength = 100
    mintSymbol = Index
    Call ISAssistant(Val(txt(Index).Tag), txt(Index))
End Sub

Private Sub lblUpInput_DblClick()
    lblUpInput.Caption = IIf(lblUpInput.Caption = "", "��", "")
    txtUpInput.SetFocus
End Sub

Private Sub lblDnInput_DblClick()
    lblDnInput.Caption = IIf(lblDnInput.Caption = "", "��", "")
    txtDnInput.SetFocus
End Sub

Private Sub lblInput_DblClick()
    lblInput.Caption = IIf(lblInput.Caption = "", "��", "")
End Sub

Private Sub txtUpInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyRight Then
        If txtUpInput.SelStart = Len(txtUpInput.Text) Then txtDnInput.SetFocus
    ElseIf KeyCode = vbKeySpace And txtUpInput.Locked Then
        lblUpInput.Caption = IIf(lblUpInput.Caption = "", "��", "")
    End If
End Sub

Private Sub txtDnInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call picDouble_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyLeft Then
        If txtDnInput.SelStart = 0 Then txtUpInput.SetFocus
    ElseIf KeyCode = vbKeySpace And txtDnInput.Locked Then
        lblDnInput.Caption = IIf(lblDnInput.Caption = "", "��", "")
    End If
End Sub

Private Sub picMutilInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MoveNextCell
End Sub

Private Sub picDouble_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    End If
End Sub

Private Sub picInput_GotFocus()
    If txtInput.Visible Then
        txtInput.SetFocus
    End If
End Sub

Private Sub picInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not txtInput.Visible Then
        If KeyCode = vbKeySpace Then
            Call lblInput_DblClick
        End If
    End If
    
    If KeyCode = vbKeyReturn Then
        '�ƶ�����һ����Ԫ��
        Call MoveNextCell
    End If
End Sub

Private Sub lstSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MoveNextCell
End Sub

Private Sub picMutilInput_GotFocus()
    On Error Resume Next
    txt(0).SetFocus
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index < txt.Count - 1 Then
            txt(Index + 1).SetFocus
        Else
            Call picMutilInput_KeyDown(KeyCode, Shift)
        End If
    End If
End Sub

Private Sub picDouble_GotFocus()
    If txtUpInput.Visible Then
        txtUpInput.SetFocus
    End If
End Sub

Private Sub picMain_Resize()
    picMain.Left = 0
    VsfData.Width = picMain.Width
    VsfData.Height = picMain.Height - VsfData.Top
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then Exit Sub
    Call picInput_KeyDown(KeyCode, Shift)
End Sub

Private Sub txtUpInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("/") Then
        KeyAscii = 0
        txtDnInput.SetFocus
    End If
End Sub

Private Sub txtС������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = Asc(";") Then KeyAscii = 0
End Sub

Private Sub cboС��_Click()
    If cboС��.Tag = "" Then Exit Sub
    
    txtС������.Text = Format(DateAdd("d", -1 * cboС�᷶Χ.ListIndex, zlDatabase.Currentdate), "MM-DD") & " " & cboС��.Text
End Sub

Private Sub cboС�᷶Χ_Click()
    Call cboС��_Click
End Sub

Private Sub UserControl_GotFocus()
    On Error Resume Next
    VsfData.SetFocus
End Sub

Private Sub UserControl_Initialize()
    mblnShow = False
    mblnChange = False
    mblnInit = False
    
'    Set objStream = objFileSys.OpenTextFile("C:\WORKLOG.txt", ForAppending, True)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    '�����ַ���Ϊ���ݷָ�������¼�¼���ķָ�������˲�����¼��
    If KeyAscii = 39 Or KeyAscii = 13 Or KeyAscii = Asc("|") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyEscape And mblnShow Then
        mblnShow = False
        Call InitCons
    End If
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    Err = 0: On Error Resume Next
    lblTitle.Move lngScaleLeft, lngScaleTop + 120, lngScaleRight - lngScaleLeft
    With lblSubhead
        .Left = lngScaleLeft + 210: .Width = lngScaleRight - lngScaleLeft - 210 * 2
        .Top = lblTitle.Top + lblTitle.Height + 120
    End With
    picMain.Move lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom - lngScaleTop
    VsfData.Move lngScaleLeft + 210, lblSubhead.Top + lblSubhead.Height + 45, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Height = picMain.Height - VsfData.Top
    
    lblCurPage.Top = picMain.Top
    lblCurPage.Left = picMain.Width - lblCurPage.Width
    
    '���ϱ�ǩ��ɢ����
    Call zlLableBruit
End Sub

Private Sub UserControl_Terminate()
'    objStream.Close
End Sub

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position
End Sub

Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '���¼�¼,���������,������
    'strPrimary:�ֶ���|ֵ
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'strPrimary = "RecordID|5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '��λ��ָ����¼
    'strPrimary:����,ֵ
    'blnDelete=True,��ü�¼������"ɾ��"�ֶ�
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !ɾ�� = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub OutputRsData(ByVal rsObj As ADODB.Recordset)
    Dim intCol As Integer, intCols As Integer
    Dim strValues As String
    With rsObj
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strValues = ""
            intCols = .Fields.Count - 1
            For intCol = 0 To intCols
                strValues = strValues & "," & .Fields(intCol).Name & ":" & .Fields(intCol).Value
            Next
            Debug.Print Mid(strValues, 2)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '�Ŵ����壬��Ԫ����
    BlowUp = dblChange
    If Not mblnBlowup Then Exit Function
    BlowUp = dblChange + (dblChange * 1 / 3)
End Function
