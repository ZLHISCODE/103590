VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmFeeRefundment 
   Caption         =   "�������תסԺ����-->�˷�"
   ClientHeight    =   6285
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10860
   Icon            =   "frmFeeRefundment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   10860
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   216
      ScaleHeight     =   495
      ScaleWidth      =   10860
      TabIndex        =   5
      Top             =   648
      Width           =   10860
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   480
         TabIndex        =   17
         Top             =   75
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   $"frmFeeRefundment.frx":058A
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
         ShowPropertySet =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F5;CTRL+A;CTRL+C"
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtPatient 
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
         Height          =   360
         Left            =   1170
         MaxLength       =   100
         TabIndex        =   9
         Top             =   72
         Width           =   2040
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   72
         Width           =   600
      End
      Begin VB.TextBox txtOld 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5100
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   72
         Width           =   585
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6555
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   72
         Width           =   1815
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Index           =   7
         Left            =   15
         TabIndex        =   13
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3252
         TabIndex        =   12
         Top             =   132
         Width           =   480
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4608
         TabIndex        =   11
         Top             =   132
         Width           =   480
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5832
         TabIndex        =   10
         Top             =   132
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5928
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmFeeRefundment.frx":0653
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14076
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.PictureBox picMzToZy 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   1095
      ScaleHeight     =   3705
      ScaleWidth      =   7785
      TabIndex        =   1
      Top             =   2025
      Width           =   7788
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3465
         ScaleHeight     =   375
         ScaleWidth      =   3840
         TabIndex        =   18
         Top             =   1950
         Visible         =   0   'False
         Width           =   3840
         Begin VB.ComboBox cboStyle 
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   615
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   0
            Width           =   1710
         End
         Begin VB.TextBox txtSum 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2295
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   1530
         End
         Begin VB.Label lblBack 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�˿�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   30
            TabIndex        =   21
            Top             =   60
            Width           =   480
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsFee 
         Height          =   2505
         Left            =   120
         TabIndex        =   2
         Top             =   30
         Width           =   5625
         _cx             =   9922
         _cy             =   4419
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483628
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
         Rows            =   10
         Cols            =   0
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFeeRefundment.frx":0EE7
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
         ExplorerBar     =   2
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
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   735
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2880
         Width           =   11160
         _cx             =   19685
         _cy             =   1296
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   360
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFeeRefundment.frx":0EFD
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
         ExplorerBar     =   3
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
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰת���ϼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   2610
         Width           =   1350
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1932
      Left            =   660
      ScaleHeight     =   1935
      ScaleWidth      =   3750
      TabIndex        =   3
      Top             =   1692
      Width           =   3756
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   4992
         Left            =   492
         TabIndex        =   4
         Top             =   36
         Width           =   9516
         _Version        =   589884
         _ExtentX        =   16775
         _ExtentY        =   8811
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picHistory 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2904
      Left            =   3240
      ScaleHeight     =   2910
      ScaleWidth      =   5910
      TabIndex        =   14
      Top             =   1212
      Width           =   5904
      Begin VSFlex8Ctl.VSFlexGrid vsHistory 
         Height          =   2208
         Left            =   108
         TabIndex        =   15
         Top             =   348
         Width           =   5628
         _cx             =   9927
         _cy             =   3895
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483628
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
         Rows            =   10
         Cols            =   0
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFeeRefundment.frx":0FC8
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
         ExplorerBar     =   2
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   -15
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmFeeRefundment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:��������ʷ��ú��շѷ��ý������ʻ��˷Ѵ���
'����:���˺�
'����:2011-03-01 14:29:10
'������:�����շ�(תסԺ�����˷�);���˽��ʹ���(תסԺ��������)
'����:36076
'---------------------------------------------------------------------------------------------------------------------------------------------
Private mlngModule As Long, mstrPrivs As String
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private mrsFeeList As ADODB.Recordset
Private mrsHistoryList As ADODB.Recordset
Private mrsBalance As ADODB.Recordset, mrsBalanceBak As ADODB.Recordset
Private mblnNotClick As Boolean
Private mstr��־ As String   '�˷�;����
Private mintSucces As Integer  '�˷ѳɹ�����
Private mlng����ID As Long, mint���� As Integer '1-�շ�;2-����
Private mblnSel As Boolean  '�Ƿ��Ѿ�ѡ������صĵ���
Private mlngShareUseID As Long
Private mcur��� As Currency
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset
Private mblnValid As Boolean
Private mrsBalanceDup As ADODB.Recordset
Private mstrStyle As String
Private mblnMultiBalance As Boolean
Private mcur�ϼ� As Currency
Private mstrUsedBills As String
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private mobjSquare As Object
 
Private Enum ҽԺҵ��
    support����������� = 33        'ҽ���Ƿ�֧������������ϣ���֧��ֻ�и������ʻ�ԭ����,�����ҽ�����㷽ʽ��Ϊ�ֽ�,֧�ֵ����ж�ÿһ�ֽ��㷽ʽ�Ƿ������˻�
End Enum
Private Enum mPgIndex
    pg_���� = 1
    pg_��ʷ���� = 2
End Enum
Private mbln�������� As Boolean
Private mbln����תסԺ����� As Boolean
Private mstrFindNO As String '���ҵ��ݺ�
Private mstrFindFpNo As String '���ҵķ�Ʊ��
Private mint�շ��嵥 As Integer      '0-����ӡ,1-Ҫ��ӡ,2-ѡ���Ƿ��ӡ
Private mblnҩ����λ As Boolean '����,����,�շ�ʱ�Ƿ������ﵥλ������ʾ������,�շ�Ҳ���ܰ�סԺ��λ
'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String
'-----------------------------------------------------------------------------------

Public Function zlShowEdit(ByVal frmMain As Object, ByVal int���� As Integer, _
    ByVal lngModuel As Long, ByVal strPrivs As String, _
    Optional ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:frmMain-���õ�������
    '       int����-1-�շѵ�;2-���ʵ�
    '       lng����ID-��ָ�����˽����˷�
    '����:
    '����:ֻҪһ�������˷ѳɹ�������true,���򷵻�False
    '����:���˺�
    '����:2011-02-22 16:31:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mintSucces = 0: mstrPrivs = strPrivs: mlngModule = lngModuel
    mlng����ID = lng����ID: mint���� = int����
    mstr��־ = "�˷�": If mint���� = 2 Then mstr��־ = "����"
    Me.Caption = IIf(mint���� = 1, "�����շ�תסԺ����-�˷ѹ���", "�������תסԺ����-���ʹ���")
    
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlShowEdit = mintSucces > 0
End Function

Private Sub cboStyle_Change()
    Call SetBlanceShow
End Sub

Private Function IsYBSingle(ByVal strNO As String, ByVal intInsure As Integer) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset, blnInsureSingle As Boolean
    
    blnInsureSingle = gclsInsure.GetCapability(83, , intInsure)
    If blnInsureSingle = False Then
        IsYBSingle = False
        Exit Function
    Else
        strSQL = "Select 1 From ҽ��������ϸ Where NO = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTmp.EOF Then
            IsYBSingle = False
        Else
            If CheckAllTurn(strNO) Then
                IsYBSingle = False
            Else
                IsYBSingle = True
            End If
        End If
    End If
    
End Function

Private Sub ClsAllNO()
   Dim i As Long
    With vsFee
        If .ColIndex("���ݺ�") >= 0 Then
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" Then
                    .TextMatrix(i, .ColIndex(mstr��־)) = 0
                End If
            Next
            Call CalcSUMMony
            Call SetBlanceShow
            mblnSel = False
        End If
    End With
End Sub
Private Sub SelAllNO()
    Dim i As Long
    With vsFee
        If .ColIndex("���ݺ�") >= 0 Then
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" Then
                    .TextMatrix(i, .ColIndex(mstr��־)) = -1
                End If
            Next
            Call CalcSUMMony
            Call SetBlanceShow
            mblnSel = True
        End If
    End With
End Sub

Private Sub zlSaveData()
    Dim i As Integer
    If SaveData = False Then
        stbThis.Panels(2).Text = IIf(mint���� = 1, "�˷�ʧ��!", "����ʧ��!")
        Exit Sub
    End If
    mstrFindNO = "": mstrFindFpNo = ""
    Call ReadListData
    Call ReadHistoryListData
    If vsFee.TextMatrix(1, vsFee.ColIndex("����")) = "" Then
        picBack.Visible = False
        For i = 1 To vsBalance.Cols - 1
            vsBalance.TextMatrix(0, i) = ""
        Next i
    End If
    mblnChange = False
    stbThis.Panels(2).Text = IIf(mint���� = 1, "�˷ѳɹ�!", "���ʳɹ�!")
End Sub

 
Private Sub cboStyle_Click()
    Call SetBlanceShow
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
          If mobjICCard Is Nothing Then
              Set mobjICCard = CreateObject("zlICCard.clsICCard")
              Set mobjICCard.gcnOracle = gcnOracle
          End If
          If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then
            Call txtPatient_KeyPress(vbKeyReturn)
        End If
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

 

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub


Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    '54896
    Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("���֤", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_Edit_ReBillingButton  '����
            Call zlSaveData
    Case conMenu_View_StatusBar
        stbThis.Visible = Not stbThis.Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Button
        cbsThis(2).Visible = Not cbsThis(2).Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each mcbrControl In cbsThis(2).Controls
            mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
        cbsThis.RecalcLayout
    Case conMenu_Edit_SelAll    'ȫѡ
            Call SelAllNO
    Case conMenu_Edit_ClsAll    'ȫ��
            Call ClsAllNO
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                Call zlCallCustomReprot(Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    Err = 0: On Error Resume Next
    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    With picFilter
        .Top = lngTop
        .Left = lngLeft
        .Width = lngRight - 100
    End With
    With picList
        .Left = lngLeft + 50: .Top = picFilter.Top + picFilter.Height
        .Width = lngRight - 100
        .Height = lngBottom - .Top
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean
    Dim i As Integer
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.pg_���� Then
            Control.Enabled = Trim(vsFee.TextMatrix(1, vsFee.ColIndex("���ݺ�"))) <> ""
        Else
            Control.Enabled = Trim(vsHistory.TextMatrix(1, vsHistory.ColIndex("���ݺ�"))) <> ""
        End If
    Case conMenu_Edit_ReBillingButton ' �˷�
        With vsFee
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("ѡ��")) <> "" Then
                    mblnSel = True
                    Exit For
                Else
                    mblnSel = False
                End If
            Next i
        End With
        Control.Enabled = mblnSel And Val(Me.tbPage.Selected.Tag) = mPgIndex.pg_����
    Case conMenu_Edit_SelAll, conMenu_Edit_ClsAll    'ȫѡ
            Control.Enabled = Val(Me.tbPage.Selected.Tag) = mPgIndex.pg_����
    Case conMenu_View_Refresh
    End Select
End Sub

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2011-01-25 15:22:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
    Set objItem = tbPage.InsertItem(mPgIndex.pg_����, IIf(mint���� = 1, "�˷Ѵ���", "���ʴ���"), picMzToZy.hWnd, 0)
    objItem.Tag = mPgIndex.pg_����
    Set objItem = tbPage.InsertItem(mPgIndex.pg_��ʷ����, IIf(mint���� = 1, "��ʷ�˷Ѽ�¼", "��ʷ���ʼ�¼"), picHistory.hWnd, 0)
    objItem.Tag = mPgIndex.pg_��ʷ����
     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
    stbThis.Top = Me.ScaleHeight - Me.stbThis.Height
End Sub
Private Sub Form_Activate()
    Dim strKey As String
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mblnChange = False
End Sub

Private Sub Form_Load()
    Dim i As Long
    mbln�������� = Val(zlDatabase.GetPara("����ת�������˷�", glngSys, 1131)) = 1
    mbln����תסԺ����� = IIf(Val(zlDatabase.GetPara("����תסԺ�����", glngSys, 1143, 0)) = 1, True, False)
    mint�շ��嵥 = 0: mblnҩ����λ = False
    If mint���� = 1 Then
        mint�շ��嵥 = Val(zlDatabase.GetPara("�շ��嵥��ӡ��ʽ", glngSys, 1121))   '�����շ�
        mblnҩ����λ = zlDatabase.GetPara("ҩƷ��λ", glngSys, 1121) = "1"
    End If
    mblnSel = False
    RestoreWinState Me, App.ProductName
    mstrPrivs = gstrPrivs: mlngModule = glngModul
    If mint���� = 1 Then
        mlngShareUseID = Val(zlDatabase.GetPara("�����շ�Ʊ������", glngSys, mlngModule, "0"))
        IDKind.IDKindStr = "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;ס|סԺ��|0;��|���￨|0;��|���ݺ�|0;��|��Ʊ��|0"
    Else
        IDKind.IDKindStr = "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;ס|סԺ��|0;��|���￨|0;��|���ݺ�|0"
        mlngShareUseID = 0
    End If
    Call initCardSquareData
    IDKind.IDKind = Val(zlDatabase.GetPara("����תסԺIDKIND", glngSys, mlngModule, "0"))
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call zlDefCommandBars
    Call InitPage
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    Set mrsInfo = New ADODB.Recordset
    vsFee.OwnerDraw = flexODContent
    '���ŵ���ʹ�ö��ֽ��㷽ʽģʽ
    mblnMultiBalance = zlDatabase.GetPara(79, glngSys) = "1"
    Call zlCreateObject
    Call LoadStyle
 End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
'    If mblnChange Then
'        If MsgBox("ע��:" & vbCrLf & "    ���޸�������,���㻹δ����,�Ƿ����Ҫ�˳�?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'            Cancel = 1: Exit Sub
'        End If
'    End If
    zlDatabase.SetPara "����תסԺIDKIND", IDKind.IDKind, glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0
    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, IIf(mint���� = 1, "�˷��б�", "�����б�"), True
    zl_vsGrid_Para_Save mlngModule, vsHistory, Me.Caption, IIf(mint���� = 1, "��ʷ�˷��б�", "��ʷ�����б�"), True
    
    SaveWinState Me, App.ProductName
    Call zlCloseObject
     
    Set mrsFeeList = Nothing
    Set mrsInfo = Nothing
    Set mrsHistoryList = Nothing
    Set mrsBalance = Nothing
    
End Sub
Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub

Private Sub picMzToZy_Resize()
    Err = 0: On Error Resume Next
    With picMzToZy
        vsFee.Top = .ScaleTop + 100
        vsFee.Width = .ScaleWidth - vsFee.Left * 2
        'cmdOk.Top = .ScaleHeight - cmdOk.Height - 50
        'cmdOk.Left = .ScaleWidth - cmdOk.Width - vsFee.Left * 2
        vsBalance.Left = vsFee.Left
        vsBalance.Width = IIf(picBack.Visible, vsFee.Width - 3000, vsFee.Width)
        picBack.Left = vsFee.Width - 4000
        vsBalance.Top = .ScaleHeight - vsBalance.Height - 100 - picBack.Height
        picBack.Top = vsBalance.Top + vsBalance.Height + 45
        lblSum.Top = IIf(vsBalance.Visible, vsBalance.Top, .ScaleHeight - stbThis.Height) - lblSum.Height - 20
        
        vsFee.Height = lblSum.Top - vsFee.Top
        'cmdAllCls.Top = .ScaleHeight - cmdAllCls.Height - 50
        'cmdAllSel.Top = cmdAllCls.Top
        'cmdOk.Top = cmdAllCls.Top
    End With
End Sub

Private Function LoadStyle() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    cboStyle.Clear
    On Error GoTo errH
    Set rsTmp = Get���㷽ʽ("�շ�", "1,2")
    For i = 1 To rsTmp.RecordCount
        If InStr(",1,2,", "," & rsTmp!���� & ",") > 0 And Val(Nvl(rsTmp!Ӧ����)) = 0 Then
            cboStyle.AddItem rsTmp!����
            cboStyle.ItemData(cboStyle.NewIndex) = rsTmp!����
            If rsTmp!ȱʡ = 1 And cboStyle.ListIndex = -1 Then
                Call zlControl.CboSetIndex(cboStyle.hWnd, cboStyle.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboStyle.ListIndex = -1 And cboStyle.ListCount > 0 Then Call zlControl.CboSetIndex(cboStyle.hWnd, 0)
    txtSum.ForeColor = vbRed
    strSQL = "" & _
            " Select B.����,B.����,Nvl(B.ȱʡ��־,0) as ȱʡ,Nvl(B.����,1) as ����,Nvl(B.Ӧ����,0) as Ӧ����" & _
            " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
            " Where A.Ӧ�ó���=[1] And B.����=A.���㷽ʽ " & _
            " And B.����<>8 " & _
            " Order by ����,lpad(����,3,' ')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "�շ�")
    For i = 1 To rsTmp.RecordCount
        If InStr(",1,2,7,", "," & rsTmp!���� & ",") > 0 Then
            mstrStyle = mstrStyle & rsTmp!���� & ":"
        End If
        rsTmp.MoveNext
    Next
    LoadStyle = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SetPicBack(ByVal strNos As String) As Boolean
    vsBalance.Width = vsFee.Width
    'picBack.Left = vsBalance.Width + vsBalance.Left + 30
    picBack.Visible = True
    SetPicBack = True
End Function

Private Sub picHistory_Resize()
    Err = 0: On Error Resume Next
    With picHistory
        vsHistory.Top = 100: vsHistory.Left = .ScaleLeft + 50
        vsHistory.Width = .ScaleWidth - vsHistory.Left * 2
        vsHistory.Height = .ScaleHeight - vsHistory.Top - 100
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   If Val(tbPage.Selected.Tag) = mPgIndex.pg_���� Then
        If vsFee.Enabled And vsFee.Visible Then vsFee.SetFocus
    Else
        Exit Sub
    End If
End Sub
Private Function InitBlanceData(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '���:strNos-ָ���ĵ��ݺ�,�Զ��ŷ���:'A0001,A0002
    '����:
    '����:
    '����:���˺�
    '����:2011-02-23 14:45:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Err = 0: On Error GoTo errHandle
    If mint���� = 2 Then
        InitBlanceData = True
        Exit Function
    End If
    If strNos = "" Then InitBlanceData = True: Exit Function
    strSQL = _
    "Select Distinct ����id" & vbNewLine & _
    "From ������ü�¼" & vbNewLine & _
    "Where NO In (Select Distinct NO" & vbNewLine & _
    "             From ������ü�¼" & vbNewLine & _
    "             Where ����id In (Select ����id" & vbNewLine & _
    "                            From ����Ԥ����¼" & vbNewLine & _
    "                            Where ������� In (Select b.�������" & vbNewLine & _
    "                                           From ������ü�¼ A, ����Ԥ����¼ B" & vbNewLine & _
    "                                           Where a.No In (Select Column_Value From Table(f_Str2list([1]))) And" & vbNewLine & _
    "                                                 Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 And a.����id = b.����id))) And" & vbNewLine & _
    "      Mod(��¼����, 10) = 1 And ��¼״̬ <> 0"

    strSQL = _
    " Select /*+ rule */ A.���㷽ʽ,Nvl(B.����,1) as ����,B.Ӧ����,A.���,A.�������" & _
    " From (  Select Decode(A.��¼����,3,A.���㷽ʽ,NULL) as ���㷽ʽ,A.�������," & _
    "               Sum(A.��Ԥ��) as ���" & _
    "         From ����Ԥ����¼ A,(" & strSQL & ") B" & _
    "         Where A.����ID=B.����ID And A.��¼���� IN(1,11,3) And Nvl(A.��Ԥ��,0)<>0" & _
    "         Group by Decode(A.��¼����,3,A.���㷽ʽ,NULL),A.�������" & _
    "       ) A,���㷽ʽ B " & _
    " Where A.���㷽ʽ=B.����(+) " & _
    " "
    Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(strNos, "'", ""))
    Set mrsBalanceBak = mrsBalance
    InitBlanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InitPatialBalance(ByVal strNos As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�������˷ѵĽ�������
    '���:strNos-ָ���ĵ��ݺ�,�Զ��ŷ���:'A0001,A0002
    '����:
    '����:
    '����:������
    '����:2014-06-03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, dblSum As Double, i As Integer, strTable As String
    Dim cur��� As Currency, rsTmp As ADODB.Recordset, rsTx As ADODB.Recordset
    Dim bln���� As Boolean, dblҽ������ As Double, dblҽ������ As Double
    Dim arrNO() As String
    Dim j As Integer
    Dim lngRow As Long
    Dim curOld As Currency
    Dim curOldTotal As Currency
    Dim strOldNOs As String, strNewNos As String
    Err = 0: On Error GoTo errHandle
    If mint���� = 2 Then
        InitPatialBalance = 0
        Exit Function
    End If
    If strNos = "" Then InitPatialBalance = 0: Exit Function
    
    Call InitBlanceData(strNos)
    dblSum = 0
    curOld = 0
    curOldTotal = 0
    
    Set mrsBalance = New ADODB.Recordset
    mrsBalance.Fields.Append "���㷽ʽ", adVarChar, 20
    mrsBalance.Fields.Append "����", adBigInt, 2
    mrsBalance.Fields.Append "Ӧ����", adBigInt, 1
    mrsBalance.Fields.Append "���", adDouble, 30
    mrsBalance.Fields.Append "ժҪ", adVarChar, 50
    mrsBalance.Fields.Append "�������", adVarChar, 30
    mrsBalance.CursorLocation = adUseClient
    mrsBalance.LockType = adLockOptimistic
    mrsBalance.CursorType = adOpenStatic
    mrsBalance.Open
    
    strNos = Replace(strNos, "'", "")
    arrNO = Split(strNos, ",")
    For i = 0 To UBound(arrNO)
        For j = 1 To vsFee.Rows - 1
            If vsFee.TextMatrix(j, vsFee.ColIndex("���ݺ�")) = arrNO(i) Then lngRow = j: Exit For
        Next j
        If CheckAllTurn(arrNO(i)) = True Then
            strOldNOs = strOldNOs & "," & arrNO(i)
        Else
            If Val(vsFee.TextMatrix(lngRow, vsFee.ColIndex("����"))) <> 0 Then
                If IsYBSingle(vsFee.TextMatrix(lngRow, vsFee.ColIndex("���ݺ�")), Val(vsFee.TextMatrix(lngRow, vsFee.ColIndex("����")))) = False Then
                    strOldNOs = strOldNOs & "," & arrNO(i)
                Else
                    strNewNos = strNewNos & "," & arrNO(i)
                End If
            Else
                strNewNos = strNewNos & "," & arrNO(i)
            End If
        End If
    Next i
    If strOldNOs <> "" Then
        strOldNOs = Mid(strOldNOs, 2)
        
        strTable = _
        "Select Distinct ����id" & vbNewLine & _
        "From ������ü�¼" & vbNewLine & _
        "Where NO In" & vbNewLine & _
        "      (Select Distinct NO" & vbNewLine & _
        "       From ������ü�¼" & vbNewLine & _
        "       Where ����id In (Select ����id" & vbNewLine & _
        "                      From ����Ԥ����¼" & vbNewLine & _
        "                      Where ������� In (Select b.�������" & vbNewLine & _
        "                                     From ������ü�¼ A, ����Ԥ����¼ B" & vbNewLine & _
        "                                     Where a.No In (Select Column_Value From Table(f_Str2list([1]))) And b.������� < 0 And" & vbNewLine & _
        "                                           Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 And a.����id = b.����id))) And" & vbNewLine & _
        "      Mod(��¼����, 10) = 1 And ��¼״̬ <> 0" & vbNewLine & _
        "Union" & vbNewLine & _
        "Select Distinct ����id" & vbNewLine & _
        "From ������ü�¼" & vbNewLine & _
        "Where NO In (Select Distinct NO" & vbNewLine & _
        "             From ������ü�¼" & vbNewLine & _
        "             Where ����id In (Select a.����id" & vbNewLine & _
        "                            From ������ü�¼ A, ����Ԥ����¼ B" & vbNewLine & _
        "                            Where a.No In (Select Column_Value From Table(f_Str2list([1]))) And b.������� > 0 And" & vbNewLine & _
        "                                  Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 And a.����id = b.����id))"

        
        strSQL = _
        " Select /*+ rule */ A.���㷽ʽ,0 as ����,Null As Ӧ����,A.���,Null As ժҪ,Null As �������" & _
        " From (  Select '��Ԥ��' as ���㷽ʽ," & _
        "               Sum(A.��Ԥ��) as ���" & _
        "         From ����Ԥ����¼ A,(" & strTable & ") B" & _
        "         Where A.����ID=B.����ID And Mod(A.��¼����,10) = 1 " & _
        "       ) A "

        strSQL = strSQL & _
        " Union " & _
        " Select /*+ rule */ A.���㷽ʽ,Nvl(B.����,1) as ����,B.Ӧ����,A.���,Null As ժҪ,Null As �������" & _
        " From (  Select Decode(A.��¼����,3,A.���㷽ʽ,NULL) as ���㷽ʽ," & _
        "               Sum(A.��Ԥ��) as ���" & _
        "         From ����Ԥ����¼ A,(" & strTable & ") B" & _
        "         Where A.����ID=B.����ID And A.��¼����=3 And Nvl(A.��Ԥ��,0)<>0" & _
        "         Group by Decode(A.��¼����,3,A.���㷽ʽ,NULL)" & _
        "       ) A,���㷽ʽ B " & _
        " Where A.���㷽ʽ=B.���� And B.���� In (3,4)"
        
        strSQL = strSQL & _
        " Union " & _
        " Select /*+ rule */ A.���㷽ʽ,Nvl(B.����,1) as ����,B.Ӧ����,A.���,Null As ժҪ,Null As �������" & _
        " From (  Select Decode(A.��¼����,3,A.���㷽ʽ,NULL) as ���㷽ʽ," & _
        "               Sum(A.��Ԥ��) as ���" & _
        "         From ����Ԥ����¼ A,(" & strTable & ") B" & _
        "         Where A.����ID=B.����ID And A.��¼����=3 And Nvl(A.��Ԥ��,0)<>0" & _
        "         And Exists (Select 1 From ҽ�ƿ���� Where ID=A.�����ID And �Ƿ����� = 0)" & _
        "         Group by Decode(A.��¼����,3,A.���㷽ʽ,NULL)" & _
        "       ) A,���㷽ʽ B " & _
        " Where A.���㷽ʽ=B.���� And B.���� In (7,8)"
        
        strSQL = strSQL & _
        " Union " & _
        " Select /*+ rule */ A.���㷽ʽ,Nvl(B.����,1) as ����,B.Ӧ����,A.���,Null As ժҪ,Null As �������" & _
        " From (  Select Decode(A.��¼����,3,A.���㷽ʽ,NULL) as ���㷽ʽ," & _
        "               Sum(A.��Ԥ��) as ���" & _
        "         From ����Ԥ����¼ A,(" & strTable & ") B" & _
        "         Where A.����ID=B.����ID And A.��¼����=3 And Nvl(A.��Ԥ��,0)<>0" & _
        "         And A.���㿨��� Is Not Null" & _
        "         Group by Decode(A.��¼����,3,A.���㷽ʽ,NULL)" & _
        "       ) A,���㷽ʽ B " & _
        " Where A.���㷽ʽ=B.���� And B.���� = 8"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strOldNOs)
        Do While Not rsTmp.EOF
            If Val(Nvl(rsTmp!���)) <> 0 Then
                With mrsBalance
                    .AddNew
                    !���㷽ʽ = Nvl(rsTmp!���㷽ʽ)
                    !���� = Nvl(rsTmp!����)
                    !Ӧ���� = "0"
                    !��� = Val(Nvl(rsTmp!���))
                    !ժҪ = ""
                    !������� = ""
                    .Update
                End With
                curOldTotal = curOldTotal + Val(Nvl(rsTmp!���))
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    strSQL = "Select Sum(ʵ�ս��) As ���" & vbNewLine & _
            "From ������ü�¼" & vbNewLine & _
            "Where NO In (Select Column_Value From Table(f_Str2list([1]))) And Mod(��¼����, 10) = 1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strOldNOs)
    If rsTmp.RecordCount <> 0 Then
        curOld = Val(Nvl(rsTmp!���))
    End If
    
    dblSum = dblSum + curOld - curOldTotal
    
    If strNewNos <> "" Then strNewNos = Mid(strNewNos, 2)
    
    cur��� = mcur�ϼ� - curOld
    
    strSQL = "Select A.���㷽ʽ, Sum(A.���) As ���,B.����" & vbNewLine & _
            "From ҽ��������ϸ A,���㷽ʽ B" & vbNewLine & _
            "Where A.NO In (Select Column_Value From Table(f_Str2list([1]))) And A.���㷽ʽ=B.����(+)" & vbNewLine & _
            "Group By A.���㷽ʽ,B.����" & vbNewLine & _
            "Having Sum(A.���) <> 0"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNewNos)
    
    Do While Not rsTmp.EOF
        If Val(Nvl(rsTmp!���)) <> 0 Then
            If Val(cur���) > Val(Nvl(rsTmp!���)) Then
                With mrsBalance
                    .AddNew
                    !���㷽ʽ = Nvl(rsTmp!���㷽ʽ)
                    !���� = Nvl(rsTmp!����)
                    !Ӧ���� = "0"
                    !��� = Val(Nvl(rsTmp!���))
                    !ժҪ = ""
                    !������� = ""
                    .Update
                End With
                cur��� = cur��� - Val(Nvl(rsTmp!���))
                If Val(Nvl(rsTmp!����)) = 3 Then dblҽ������ = dblҽ������ + Val(Nvl(rsTmp!���))
                If Val(Nvl(rsTmp!����)) = 4 Then dblҽ������ = dblҽ������ + Val(Nvl(rsTmp!���))
            Else
                With mrsBalance
                    .AddNew
                    !���㷽ʽ = Nvl(rsTmp!���㷽ʽ)
                    !���� = Nvl(rsTmp!����)
                    !Ӧ���� = "0"
                    !��� = cur���
                    !ժҪ = ""
                    !������� = ""
                    .Update
                End With
                InitPatialBalance = Format(dblSum, "0.00")
                Exit Function
            End If
        End If
        rsTmp.MoveNext
    Loop
    

    strTable = _
        "Select Distinct ����id" & vbNewLine & _
        "From ������ü�¼" & vbNewLine & _
        "Where NO In" & vbNewLine & _
        "      (Select Distinct NO" & vbNewLine & _
        "       From ������ü�¼" & vbNewLine & _
        "       Where ����id In (Select ����id" & vbNewLine & _
        "                      From ����Ԥ����¼" & vbNewLine & _
        "                      Where ������� In (Select b.�������" & vbNewLine & _
        "                                     From ������ü�¼ A, ����Ԥ����¼ B" & vbNewLine & _
        "                                     Where a.No In (Select Column_Value From Table(f_Str2list([1]))) And b.������� < 0 And" & vbNewLine & _
        "                                           Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 And a.����id = b.����id))) And" & vbNewLine & _
        "      Mod(��¼����, 10) = 1 And ��¼״̬ <> 0" & vbNewLine & _
        "Union" & vbNewLine & _
        "Select Distinct ����id" & vbNewLine & _
        "From ������ü�¼" & vbNewLine & _
        "Where NO In (Select Distinct NO" & vbNewLine & _
        "             From ������ü�¼" & vbNewLine & _
        "             Where ����id In (Select a.����id" & vbNewLine & _
        "                            From ������ü�¼ A, ����Ԥ����¼ B" & vbNewLine & _
        "                            Where a.No In (Select Column_Value From Table(f_Str2list([1]))) And b.������� > 0 And" & vbNewLine & _
        "                                  Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 And a.����id = b.����id))"
    
    strSQL = _
    " Select /*+ rule */ '��Ԥ��' as ���㷽ʽ," & _
    "               Sum(A.��Ԥ��) as ���" & _
    "         From ����Ԥ����¼ A,(" & strTable & ") B" & _
    "         Where A.����ID=B.����ID And A.��¼���� IN(1,11) And Nvl(A.��Ԥ��,0)<>0"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNewNos)
    
    If rsTmp.RecordCount <> 0 Then
        If Val(Nvl(rsTmp!���)) <> 0 Then
            If Val(cur���) > Val(Nvl(rsTmp!���)) Then
                With mrsBalance
                    .AddNew
                    !���㷽ʽ = Nvl(rsTmp!���㷽ʽ)
                    !���� = 0
                    !Ӧ���� = "0"
                    !��� = Val(Nvl(rsTmp!���))
                    !ժҪ = ""
                    !������� = ""
                    .Update
                End With
                cur��� = cur��� - Val(Nvl(rsTmp!���))
            Else
                With mrsBalance
                    .AddNew
                    !���㷽ʽ = Nvl(rsTmp!���㷽ʽ)
                    !���� = 0
                    !Ӧ���� = "0"
                    !��� = cur���
                    !ժҪ = ""
                    !������� = ""
                    .Update
                End With
                InitPatialBalance = Format(dblSum, "0.00")
                Exit Function
            End If
        End If
    End If
    
    strSQL = "Select a.���㷽ʽ, Sum(a.��Ԥ��) As ���, a.�����id, a.���㿨���, a.����, Min(a.������ˮ��) As ������ˮ��," & vbNewLine & _
            "                        Min(a.����˵��) As ����˵��, Min(a.������λ) As ������λ, b.����" & vbNewLine & _
            "                 From ����Ԥ����¼ A, ���㷽ʽ B" & vbNewLine & _
            "                 Where a.��¼���� = 3 And a.����id In (" & strTable & ") And a.���㷽ʽ = b.���� And" & vbNewLine & _
            "                       b.���� In (1, 2, 7, 8)" & vbNewLine & _
            "                 Group By a.���㷽ʽ, a.�����id, a.���㿨���, a.����, b.����" & vbNewLine & _
            "                 Having Sum(a.��Ԥ��) <> 0" & vbNewLine & _
            "                 Order By �����id,���� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNewNos)
    
    Do While Not rsTmp.EOF
        If Val(Nvl(rsTmp!����)) = 7 Or (Val(Nvl(rsTmp!����)) = 8 And Not IsNull(rsTmp!�����ID)) Then
            strSQL = "Select 1 from ҽ�ƿ���� Where id = [1] And �Ƿ����� = 1"
            Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(rsTmp!�����ID)))
            bln���� = Not rsTx.EOF
            If bln���� Then
                If Val(cur���) > Val(Nvl(rsTmp!���)) Then
                    dblSum = dblSum + Val(Nvl(rsTmp!���))
                    cur��� = cur��� - Val(Nvl(rsTmp!���))
                Else
                    dblSum = dblSum + cur���
                    InitPatialBalance = Format(dblSum, "0.00")
                    Exit Function
                End If
            Else
                If Val(Nvl(rsTmp!���)) <> 0 Then
                    If Val(cur���) > Val(Nvl(rsTmp!���)) Then
                        With mrsBalance
                            .AddNew
                            !���㷽ʽ = Nvl(rsTmp!���㷽ʽ)
                            !���� = Nvl(rsTmp!����)
                            !Ӧ���� = "0"
                            !��� = Val(Nvl(rsTmp!���))
                            !ժҪ = ""
                            !������� = ""
                            .Update
                        End With
                        cur��� = cur��� - Val(Nvl(rsTmp!���))
                    Else
                        With mrsBalance
                            .AddNew
                            !���㷽ʽ = Nvl(rsTmp!���㷽ʽ)
                            !���� = Nvl(rsTmp!����)
                            !Ӧ���� = "0"
                            !��� = cur���
                            !ժҪ = ""
                            !������� = ""
                            .Update
                        End With
                        InitPatialBalance = Format(dblSum, "0.00")
                        Exit Function
                    End If
                End If
            End If
        Else
            If Val(Nvl(rsTmp!����)) = 8 Then
                If Val(Nvl(rsTmp!���)) <> 0 Then
                    If Val(cur���) > Val(Nvl(rsTmp!���)) Then
                        With mrsBalance
                            .AddNew
                            !���㷽ʽ = Nvl(rsTmp!���㷽ʽ)
                            !���� = Nvl(rsTmp!����)
                            !Ӧ���� = "0"
                            !��� = Val(Nvl(rsTmp!���))
                            !ժҪ = ""
                            !������� = ""
                            .Update
                        End With
                        cur��� = cur��� - Val(Nvl(rsTmp!���))
                    Else
                        With mrsBalance
                            .AddNew
                            !���㷽ʽ = Nvl(rsTmp!���㷽ʽ)
                            !���� = Nvl(rsTmp!����)
                            !Ӧ���� = "0"
                            !��� = cur���
                            !ժҪ = ""
                            !������� = ""
                            .Update
                        End With
                        InitPatialBalance = Format(dblSum, "0.00")
                        Exit Function
                    End If
                End If
            Else
                If Val(cur���) > Val(Nvl(rsTmp!���)) Then
                    dblSum = dblSum + Val(Nvl(rsTmp!���))
                    cur��� = cur��� - Val(Nvl(rsTmp!���))
                Else
                    dblSum = dblSum + cur���
                    InitPatialBalance = Format(dblSum, "0.00")
                    Exit Function
                End If
            End If
        End If
        rsTmp.MoveNext
    Loop
    

    If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
    
    InitPatialBalance = Format(dblSum, "0.00")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "ȫѡ(&A)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "ȫ��(&C)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReBillingButton, IIf(mint���� = 1, "�˷�(&X)", "����(&X)")): mcbrControl.BeginGroup = True
    End With
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll
        .Add FCONTROL, vbKeyC, conMenu_Edit_ClsAll
      End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "ȫѡ"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "ȫ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ReBillingButton, IIf(mint���� = 1, "�˷�", "����"))
        mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:���˺�
    '����:2011-01-25 15:14:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsGrid As VSFlexGrid, rsTemp As New ADODB.Recordset, strSQL As String
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    If mint���� = 1 Then
        objPrint.Title.Text = gstrUnitName & "����תסԺ�˷����"
    Else
        objPrint.Title.Text = gstrUnitName & "����תסԺ�������"
    End If
    
    objRow.Add "���ˣ�" & txtPatient.Text
    objRow.Add "�Ա�" & txtSex.Text
    objRow.Add "���䣺" & txtOld.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    If Val(Me.tbPage.Selected.Tag) = mPgIndex.pg_���� Then
        Set vsGrid = vsFee
    Else
        Set vsGrid = vsHistory
    End If
    
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .Cell(flexcpData, 0, intCol) = .ColWidth(intCol)
            If .ColHidden(intCol) Or intCol = .ColIndex(mstr��־) Then .ColWidth(intCol) = 0
        Next
    End With
    
    Set objPrint.Body = vsGrid
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub zlCallCustomReprot(ByVal lngSys As Long, strReprotName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ص��Զ��屨��
    '����:���˺�
    '����:2011-01-25 15:16:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As Variant, lng����ID As Long
    Dim vsGrid As VSFlexGrid
    If Val(Me.tbPage.Selected.Tag) = mPgIndex.pg_���� Then
        Set vsGrid = vsFee
    Else
        Set vsGrid = vsHistory
    End If
    With vsGrid
        If .Row > 0 Then
            strNO = Trim(.TextMatrix(.Row, .ColIndex("���ݺ�")))
        End If
        If strNO <> "" Then
            Call ReportOpen(gcnOracle, lngSys, strReprotName, Me, "NO=" & strNO)
        Else
            Call ReportOpen(gcnOracle, lngSys, strReprotName, Me)
        End If
    End With
End Sub

Private Sub txtPatient_Change()
    txtPatient.Tag = ""
    If txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "")
    stbThis.Panels(2).Text = ""
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")

End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    
    If txtPatient.Locked Then Exit Sub
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
    Else
        If IDKind.GetCurCard.���� Like "����*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        ElseIf IDKind.GetCurCard.���� = "�����" Or IDKind.GetCurCard.���� = "סԺ��" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
         Else
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
        End If
    End If
    
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        ElseIf IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
            KeyAscii = 0
            'ˢ�²�����Ϣ:"-����ID"
            Call GetPatient(IDKind.GetCurCard, txtPatient.Tag, False)
            If mrsInfo.State = 0 Then   '
                txtPatient.Text = "": txtOld.Text = ""
                txtSex.Text = "": txtסԺ��.Text = ""
                Exit Sub
            End If
            Call ReadListData
            Call ReadHistoryListData
            Exit Sub
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, txtPatient.Text)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnMsg As Boolean, blnICCard As Boolean, blnIDCard As Boolean
 
    '54899
    If objCard.���� Like "IC��*" And objCard.ϵͳ = True And mstrPassWord <> "" Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.���� Like "*���֤*" And objCard.ϵͳ = True And mstrPassWord <> "" Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        
    If Not GetPatient(objCard, strInput, blnCard, blnMsg) Then
        If blnCard Then
            If Not blnMsg Then MsgBox "����ȷ��������Ϣ�������Ƿ���ȷˢ����", vbInformation, gstrSysName
            txtPatient.Text = "": txtOld.Text = ""
            txtסԺ��.Text = ""
            vsFee.Clear 1: vsFee.Rows = 2
            vsHistory.Clear 1: vsHistory.Rows = 2
            Exit Sub
        End If
        If Not blnMsg Then MsgBox "���ܶ�ȡ������Ϣ��", vbInformation, gstrSysName
        zlControl.TxtSelAll txtPatient
        txtOld.Text = "": txtSex.Text = "": txtסԺ��.Text = ""
        vsFee.Clear 1: vsFee.Rows = 2
        vsHistory.Clear 1: vsHistory.Rows = 2
        Exit Sub
    End If
    
    '��ȡ�ɹ�
    '���￨������
     If (objCard.���� Like "IC��*" Or objCard.���� Like "*���֤*") And objCard.ϵͳ = True And blnCard Then blnCard = False
    If Mid(gstrCardPass, 6, 1) = "1" And (blnCard Or blnICCard Or blnIDCard) Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
            vsFee.Clear 1: vsFee.Rows = 2
            vsHistory.Clear 1: vsHistory.Rows = 2
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
        End If
    End If
    Call ReadListData
    Call ReadHistoryListData
 
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub
Private Sub txtPatient_Validate(Cancel As Boolean)
    If IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
        mblnValid = True
        Call txtPatient_KeyPress(13)
        mblnValid = False
    End If
End Sub
Private Sub zlClearPatiInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������Ϣ
    '����:���˺�
    '����:2011-02-23 09:39:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
    txtסԺ��.Text = "": Set mrsInfo = New ADODB.Recordset
    vsFee.Clear 1: vsFee.Rows = 2
    vsHistory.Clear 1: vsHistory.Rows = 2
End Sub
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��
    '����: blnOutMsg-�Ѿ���ʾ,�������ⲿ����ʾ
    '����:
    '����:���˺�
    '����:2011-01-25 16:57:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim strIF As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    mstrFindNO = "": mstrFindFpNo = ""
    
    strSQL = _
        "   Select A.����ID,Nvl(B.��ҳID,0) as ��ҳID,A.����� as �����,A.��ǰ����,B.��Ժ����," & _
        "      Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�,Nvl(B.����,A.����)   as ����,A.IC����,A.���￨��,A.����֤��," & _
        "       Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�,C.���� as ��ǰ����,A.��ǰ����ID,D.���� as ��Ժ����,B.��Ժ����ID, A.���� as ����,E.����,E.ҽ����,E.����," & _
        "       A.�Ǽ�ʱ��,Nvl(B.״̬,0) as ״̬,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,Nvl(B.��˱�־,0) as ��˱�־,B.��Ժ����,B.��Ժ����,B.��������,B.��������" & _
        " From ������Ϣ A,������ҳ B,���ű� C,���ű� D,ҽ�����˵��� E,ҽ�����˹����� F" & _
        " Where A.ͣ��ʱ�� is NULL And A.����ID=B.����ID(+) And A.��ҳID=B.��ҳID(+) " & _
        "           And A.����ID=F.����ID(+) And F.��־(+)=1 And F.ҽ����=E.ҽ����(+) And F.����=E.����(+) And F.���� = E.����(+)" & _
        "           And A.��ǰ����ID=C.ID(+) And B.��Ժ����ID=D.ID(+)" & _
        "           And A.ͣ��ʱ�� is NULL "
    
    If blnCard = True And objCard.���� Like "����*" Then  'ˢ��
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.����ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSQL = strSQL & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(������Ժ)
        strSQL = strSQL & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSQL = strSQL & " And A.�����=[1]"
    ElseIf Left(strInput, 1) = "." Or IDKind.IDKind = IDKind.GetKindIndex("���ݺ�") Then
        '���ݺŲ���
        If Left(strInput, 1) = "." Then
            strTemp = UCase(GetFullNO(Mid(strInput, 2), IIf(mint���� = 1, 13, 14)))
        Else
            strTemp = UCase(GetFullNO(strInput, IIf(mint���� = 1, 13, 14)))
        End If
        txtPatient.Text = strTemp
        gstrSQL = "" & _
        "   Select  distinct A.����ID " & _
        "   From ������ü�¼ A " & _
        "   Where A.NO=[1] and Mod(A.��¼����,10)=[2] " & _
        "              And Rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp, mint����)
        If rsTemp.EOF Then
            MsgBox "ע��:" & vbCrLf & "  ���ݺ�Ϊ��" & strTemp & "��������,��������ĵ����Ƿ���ȷ!", vbInformation + vbOKOnly, gstrSysName
            Call zlClearPatiInfor
            Exit Function
        End If
        If Not GetPatient("-" & rsTemp!����ID, False, True) Then
            Call zlClearPatiInfor
            Exit Function
        End If
        mstrFindNO = strTemp
        GetPatient = True
        Exit Function
    
    Else '��������
        Select Case objCard.����
            Case "����"
                If mrsInfo.State = 1 Then
                    If Not mrsInfo.EOF Then
                        If mrsInfo!���� = Trim(txtPatient.Text) Then GetPatient = True: Exit Function
                    End If
                End If
              strSQL = strSQL & " And A.����=[2]"
            Case "ҽ����"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.ҽ����=[2]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
            Case "��Ʊ��"
                strSQL = "" & _
                "   Select distinct A.����ID " & _
                "   From ������ü�¼ A,Ʊ�ݴ�ӡ���� B,Ʊ��ʹ����ϸ C" & _
                "   Where A.NO=B.NO and Mod(A.��¼����,10)=1 and A.��¼״̬=1  " & _
                "               and  B.��������=1 And B.ID=C.��ӡID and C.Ʊ��=1 And C.����=1 And C.����=[1] And Rownum=1 " & _
                "   "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, mint����)
                If rsTemp.EOF Then
                    MsgBox "ע��:" & vbCrLf & "  ��Ʊ��Ϊ��" & strInput & "��������,��������ķ�Ʊ���Ƿ���ȷ!", vbInformation + vbOKOnly, gstrSysName
                    Call zlClearPatiInfor
                    Exit Function
                End If
                If Not GetPatient(objCard, "-" & rsTemp!����ID, False, True) Then
                    Call zlClearPatiInfor
                    Exit Function
                End If
                mstrFindFpNo = strInput
                GetPatient = True
                Exit Function
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
            End Select
    End If
    txtPatient.ForeColor = Me.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    
    If Not mrsInfo.EOF Then
        txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!��������))
        txtPatient.Text = Nvl(mrsInfo!����): txtOld.Text = Nvl(mrsInfo!����): txtSex.Text = Nvl(mrsInfo!�Ա�)
        txtסԺ��.Text = Nvl(mrsInfo!�����)
        If Val(Nvl(mrsInfo!��ҳID)) <> 0 Then
            If zlIsAllowFeeChange(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID))) = False Then
                txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
                txtסԺ��.Text = ""
                If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
                Exit Function
            End If
        End If
        
        txtPatient.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!����֤��)
        GetPatient = True
        Exit Function
    Else
        txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
        txtסԺ��.Text = ""
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
NotFoundPati:
    txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
    txtסԺ��.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Set mrsInfo = New ADODB.Recordset
End Function
Private Function zlGetFpToBIllNOs(ByVal strFpNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���ķ�Ʊ��,�ҳ���Ӧ�ĵ��ݺ�
    '����:���ض�Ӧ�ĵ��ݺ�,�ö��ŷָ�
    '����:���˺�
    '����:2011-02-25 10:50:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, strNos As String
    
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select distinct NO From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B " & _
    "   Where A.��������=1 and A.ID=B.��ӡID and B.Ʊ��=1 And B.����=[1]  " & _
    "   Order by NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFpNo)
    strNos = ""
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & Nvl(rsTemp!NO)
            .MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    zlGetFpToBIllNOs = strNos
 
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ReadListData(Optional blnFilter As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ҫ���ʵ���ϸ����
    '����:��ȡ�ɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-01-25 17:10:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim lngCol As Long, strSQL As String, lngRow As Long
    Dim strFilter As String, strNos As String
    Dim strWhere As String, strTable1 As String
    Dim strALLNOs As String
    If mrsInfo Is Nothing Then
        lng����ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = 0
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    If mstrFindNO <> "" Then
        If mint���� = 1 Then
            strNos = Replace(GetMultiNOs(mstrFindNO), "'", "")
        Else
            strNos = mstrFindNO
        End If
        strTable1 = ",Table( f_Str2list([2])) J "
        strWhere = "  And A.NO=J.Column_Value"
    ElseIf mstrFindFpNo <> "" And mint���� = 1 Then
        strNos = zlGetFpToBIllNOs(mstrFindFpNo)
        If strNos = "" Then
            MsgBox "δ�ҵ���Ӧ��Ʊ�ŵĵ���,����!"
            Exit Function
        End If
        strTable1 = ",Table( f_Str2list([2])) J "
        strWhere = "  And A.NO=J.Column_Value"
    Else
        strTable1 = ""
        strWhere = "  And A.����ID=[1]"
    End If
    mblnSel = False
    On Error GoTo errHandle
    If blnFilter = False Then zlCommFun.ShowFlash "���ڶ�ȡ��������,���Ժ� ..."
    Screen.MousePointer = 11
    DoEvents
    Me.Refresh
    If mint���� = 1 Then
        strTable = " " & _
        "Select a.����, a.ҽ��, b.Id, a.����, a.No, a.ʵ��Ʊ��, a.���, a.�շ����, a.��������, a.�շ�ϸĿid, a.ִ�в���id, a.����, a.����, a.����, a.Ӧ�ս��, a.ʵ�ս��," & vbNewLine & _
        "       a.������, a.����ʱ��, a.�����, a.�������, a.ת����, a.ת��ʱ��, b.����id " & _
        "From (Select Max(����) as ����, Decode(Max(����), 0, '', '��') As ҽ��, '�շѵ�' As ����, " & vbNewLine & _
        "           NO, ʵ��Ʊ��, ���, �շ����, ��������, �շ�ϸĿid, ִ�в���id, " & vbNewLine & _
        "           Avg(Nvl(����, 1)) As ����, Sum(����) ����, ��׼���� As ����, Sum(Ӧ�ս��) As Ӧ�ս��, " & vbNewLine & _
        "           Sum(ʵ�ս��) As ʵ�ս��, ������, To_Char(Max(����ʱ��), 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, " & vbNewLine & _
        "           Max(�����) As �����, Max(�������) As �������, Max(ת����) As ת����,Max(ת��ʱ��) As ת��ʱ�� " & vbNewLine & _
        "      From(Select Row_Number() Over(Partition By a.ID Order By m.���) As Rn, a.ID,Nvl(M.����,0) as ����, A.�۸񸸺�, " & vbNewLine & _
        "               A.NO, A.ʵ��Ʊ��, A.��� As ���, A.�շ����, A.��������, A.�շ�ϸĿid, A.ִ�в���id, " & vbNewLine & _
        "               A.����, A.����, A.��׼����, A.Ӧ�ս��, " & vbNewLine & _
        "               A.ʵ�ս��, A.������, A.����ʱ��, " & vbNewLine & _
        "               Q.�����, Q.�������,Q.ת����, Q.ת��ʱ��,A.��¼״̬" & vbNewLine & _
        "           From ������ü�¼ A, ���ս����¼ M, ������˼�¼ Q " & strTable1 & vbNewLine & _
        "           Where Mod(A.��¼����,10) = 1  " & strWhere & _
        "               And A.��¼״̬ <> 0 And A.����id = M.��¼id(+)" & vbNewLine & _
        "               And  M.����(+) = 1 And A.ID = Q.����id(+) And Nvl(a.���ӱ�־,0) <> 9 " & vbNewLine & _
        "               And a.Id In (Select b.Id " & vbNewLine & _
        "                        From ������ü�¼ B, ������ü�¼ C, ������˼�¼ D" & vbNewLine & _
        "                        Where c.Id = d.����id And d.��¼״̬ = 1 And b.No = c.No))" & vbNewLine & _
        "      Where Rn < 2" & _
        "      Group By NO, ʵ��Ʊ��, ���, �շ����, ��׼����,�շ�ϸĿid, ��������, ִ�в���id,������, ����ʱ��" & _
        "      Having Sum(����) <> 0) A, ������ü�¼ B Where a.No = b.No And Mod(b.��¼����,10) = 1 And a.��� = b.��� And b.��¼״̬ In (1,3) " & _
        "      And b.�Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ������ü�¼ Where NO = a.No And Mod(��¼����, 10) = 1 And ��� = a.��� And ��¼״̬ In (1,3))"
    Else
        '���ʵ�
        strTable = " " & _
        "    Select 0 as ����, Decode(NULL, Null, '', '��') As ҽ��, Max(Decode(A.�۸񸸺�, Null, ID, 0)) As ID, '���ʵ�' As ����, " & vbNewLine & _
        "           A.NO, A.ʵ��Ʊ��, A.��� As ���, A.�շ����, A.��������, A.�շ�ϸĿid, A.ִ�в���id, " & vbNewLine & _
        "           Avg(Nvl(A.����, 1)) As ����, Sum(A.����) ����, A.��׼���� As ����, Sum(A.Ӧ�ս��) As Ӧ�ս��, " & vbNewLine & _
        "           Sum(A.ʵ�ս��) As ʵ�ս��, A.������, To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, " & vbNewLine & _
        "           Max(Q.�����) As �����, Max(Q.�������) As �������, Max(Q.ת����) As ת����, " & vbNewLine & _
        "           Max(Q.ת��ʱ��) As ת��ʱ��,0 AS ����ID " & vbNewLine & _
        "    From ������ü�¼ A,  ������˼�¼ Q " & strTable1 & vbNewLine & _
        "    Where  A.��¼���� = 2 " & strWhere & vbNewLine & _
        "               And A.��¼״̬ <> 0 And A.ID = Q.����id(+) " & vbNewLine & _
        "           And a.Id In (Select b.Id" & vbNewLine & _
        "                        From ������ü�¼ B, ������ü�¼ C, ������˼�¼ D" & vbNewLine & _
        "                        Where c.Id = d.����id And d.��¼״̬ = 1 And b.No = c.No)" & vbNewLine & _
        "    Group By A.NO, A.ʵ��Ʊ��, A.���, A.��׼����, A.�շ����, A.�շ�ϸĿid, A.��������, A.ִ�в���id, " & vbNewLine & _
        "              A.������, A.����ʱ��, ����ID Having Sum(A.����) <> 0"
    End If
    strSQL = "" & _
    " Select  A.ID,'' as " & mstr��־ & ",A.����,A.No as ���ݺ�,A.ʵ��Ʊ�� As Ʊ�ݺ�, " & vbNewLine & _
    "       A.���,A.��������,A.�շ�ϸĿID,A.ִ�в���ID,A.�շ����,P.���, " & vbNewLine & _
    "       C.���� as ����,Nvl(B.����,C.����) as ����,E1.���� as ��Ʒ��,C.���," & vbNewLine & _
    "       A.����, A.����,C.���㵥λ," & vbNewLine & _
    "       ltrim(to_char(A.����,'9999990.00000')) as ����," & vbNewLine & _
    "       ltrim(to_char(A.Ӧ�ս��,'9999990.00')) as Ӧ�ս��," & vbNewLine & _
    "       ltrim(to_char(A.ʵ�ս��,'9999990.00')) as ʵ�ս��," & vbNewLine & _
    "       A.������,A.����ʱ��,A.ҽ��, A.����,A.�����, " & vbNewLine & _
    "       A.�������,A.ת����,A.ת��ʱ��,A.����ID" & vbNewLine & _
    "From (" & strTable & ") A,�շ���ĿĿ¼ C,�շ���Ŀ���� B,�շ���Ŀ���� E1,�շ���� P" & _
    " Where A.�շ�ϸĿID=C.ID And A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    "       and A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
    "       And A.�շ����=P.����(+)" & _
    " Order by A.����,A.NO,A.���"
    If mrsFeeList Is Nothing Or blnFilter = False Then
        Set mrsFeeList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, strNos)
    Else
        mrsFeeList.Filter = 0
    End If
    vsFee.Redraw = flexRDNone
    vsFee.Clear: vsFee.Cols = 0
    Set vsFee.DataSource = mrsFeeList
    If vsFee.Rows <= 1 Then vsFee.Rows = 2
    With vsFee
        For lngCol = 0 To .Cols - 1
             .ColAlignment(lngCol) = flexAlignLeftCenter
             .FixedAlignment(lngCol) = flexAlignCenterCenter
              .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
              If .ColKey(lngCol) Like "*ID" Or InStr(1, ",����,����,���,��������,ת����־,�շ����,", "," & .ColKey(lngCol) & ",") > 0 Then
                    .ColHidden(lngCol) = True
              ElseIf .ColKey(lngCol) Like "*��*" Or .ColKey(lngCol) Like "*��*" Or .ColKey(lngCol) Like "*��" Then
                    .ColAlignment(lngCol) = flexAlignRightCenter
              End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .ColDataType(.ColIndex(mstr��־)) = flexDTBoolean
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsFee, Me.Caption, IIf(mint���� = 1, "�˷��б�", "�����б�"), True
        '����
        Dim strNO As String, str���� As String
        strALLNOs = ""
        For lngRow = 1 To .Rows - 1
            If strNO <> Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) _
                 And strNO <> "" Then
                '�����ָ���
                .Select lngRow, .FixedCols, lngRow, .Cols - 1
                .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
            End If
            .Cell(flexcpData, lngRow, .ColIndex("���ݺ�")) = .TextMatrix(lngRow, .ColIndex("���ݺ�"))
            .Cell(flexcpData, lngRow, .ColIndex(mstr��־)) = Val(.TextMatrix(lngRow, .ColIndex(mstr��־)))
            strNO = Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�")))
            str���� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
            strALLNOs = strALLNOs & "," & strNO
        Next
        .Editable = flexEDKbdMouse
    End With
    If strALLNOs <> "" Then strALLNOs = Mid(strALLNOs, 2)
    If blnFilter = False Then zlCommFun.StopFlash
    vsFee.Redraw = flexRDBuffered
    '���ؽ��㷽ʽ
    Call InitBlanceData(strALLNOs)
    Call CalcSUMMony
    Call SetBlanceShow
    Call StatusShowBillSum
    
    Screen.MousePointer = 0
    ReadListData = True
    Exit Function
errHandle:
    vsFee.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
   If blnFilter = False Then zlCommFun.StopFlash
End Function
Private Function ReadHistoryListData(Optional blnFilter As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ҫ���ʵ���ϸ����
    '����:��ȡ�ɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-01-25 17:10:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim lngCol As Long, strSQL As String, lngRow As Long
    Dim strFilter As String, strNos As String
    Dim strWhere As String
    If mrsInfo Is Nothing Then
        lng����ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = 0
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    On Error GoTo errHandle
    If blnFilter = False Then zlCommFun.ShowFlash "���ڶ�ȡ��ʷת����������,���Ժ� ..."
    Screen.MousePointer = 11
    DoEvents
    Me.Refresh
    If mint���� = 1 Then
        strTable = "" & _
        " Select Max(Nvl(����, 0)) As ����, Decode(Max(Nvl(����, 0)), 0, '', '��') As ҽ��, Max(Decode(�۸񸸺�, Null, ID, 0)) As ID," & vbNewLine & _
        "       '�շѵ�' As ����, NO, ʵ��Ʊ��, Nvl(�۸񸸺�, ���) As ���, �շ����, ��������, �շ�ϸĿid, ִ�в���id, Avg(Nvl(����, 1)) As ����, -1 * Avg(����) ����," & vbNewLine & _
        "       Sum(��׼����) As ����, -1 * Sum(Ӧ�ս��) As Ӧ�ս��, -1 * Sum(ʵ�ս��) As ʵ�ս��, ������," & vbNewLine & _
        "       To_Char(����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, ����id, Max(�����) As �����, Max(�������) As �������, Max(ת����) As ת����," & vbNewLine & _
        "       Max(ת��ʱ��) As ת��ʱ��" & vbNewLine & _
        " From (Select Row_Number() Over(Partition By a.Id Order By m.���) As Rn, a.Id, m.����, a.No, a.ʵ��Ʊ��, a.�۸񸸺�, a.���, a.�շ����," & vbNewLine & _
        "              a.��������, a.�շ�ϸĿid, a.ִ�в���id, a.����, a.����, a.��׼����, a.Ӧ�ս��, a.ʵ�ս��, a.������, a.����ʱ��, a.����id, q.�����, q.�������, q.ת����," & vbNewLine & _
        "              q.ת��ʱ��" & vbNewLine & _
        "       From ������ü�¼ A, ���ս����¼ M, ������˼�¼ Q, ������ü�¼ K" & vbNewLine & _
        "       Where Mod(a.��¼����, 10) = 1 and A.����ID=[1] " & strWhere & _
        "             And a.����id = m.��¼id(+) And m.����(+) = 1 " & vbNewLine & _
        "             And q.����id(+) = k.Id" & vbNewLine & _
        "             And k.Id In (Select Min(ID) From ������ü�¼ Where NO = a.No And Mod(��¼����,10) = 1 And ��� = a.���)" & vbNewLine & _
        "             And a.Id In (Select Max(d.Id)" & vbNewLine & _
        "                      From ������ü�¼ D, ������ü�¼ B, ������˼�¼ C" & vbNewLine & _
        "                      Where b.����id + 0 = a.����id And b.��� = a.��� And b.Id = c.����id And a.No = b.No And d.��¼״̬ = 2 And" & vbNewLine & _
        "                            d.No = b.No And c.��¼״̬ = 2" & vbNewLine & _
        "                      Group By d.���))" & vbNewLine & _
        " Where Rn < 2" & vbNewLine & _
        " Group By NO, ʵ��Ʊ��, Nvl(�۸񸸺�, ���), �շ����, �շ�ϸĿid, ��������, ִ�в���id, ������, ����ʱ��, ����id"
    Else
        '���ʵ�
        strTable = " " & _
        "    Select 0 as ����, Decode(NULL, Null, '', '��') As ҽ��, Max(Decode(A.�۸񸸺�, Null, a.ID, 0)) As ID, '���ʵ�' As ����, " & vbNewLine & _
        "           A.NO, A.ʵ��Ʊ��, Nvl(A.�۸񸸺�, A.���) As ���, A.�շ����, A.��������, A.�շ�ϸĿid, A.ִ�в���id, " & vbNewLine & _
        "           Avg(Nvl(A.����, 1)) As ����, -1 * Avg(A.����) ����, Sum(A.��׼����) As ����, -1 * Sum(A.Ӧ�ս��) As Ӧ�ս��, " & vbNewLine & _
        "           -1 * Sum(A.ʵ�ս��) As ʵ�ս��, A.������, To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, 0 as ����id, " & vbNewLine & _
        "           Max(Q.�����) As �����, Max(Q.�������) As �������, Max(Q.ת����) As ת����, " & vbNewLine & _
        "           Max(Q.ת��ʱ��) As ת��ʱ�� " & vbNewLine & _
        "    From ������ü�¼ A,  ������˼�¼ Q, ������ü�¼ K " & vbNewLine & _
        "    Where  A.��¼���� = 2  and A.����ID=[1] " & strWhere & vbNewLine & _
        "               And q.����id(+) = k.Id " & vbNewLine & _
        "           And k.Id In (Select Min(ID) From ������ü�¼ Where NO = a.No And ��¼���� = 2 And ��� = a.���) " & _
        "           And a.Id In (Select Max(d.Id)" & vbNewLine & _
        "         From ������ü�¼ D, ������ü�¼ B, ������˼�¼ C" & vbNewLine & _
        "         Where b.����id + 0 = a.����id And b.���=a.��� And b.Id = c.����id And a.No = b.No And d.��¼״̬ = 2 And d.No = b.No And c.��¼״̬ = 2" & vbNewLine & _
        "         Group By d.���) " & _
        "    Group By A.NO, A.ʵ��Ʊ��, Nvl(A.�۸񸸺�, A.���), A.�շ����, A.�շ�ϸĿid, A.��������, A.ִ�в���id, " & vbNewLine & _
        "              A.������, A.����ʱ��"
    End If
    strSQL = "" & _
    " Select  A.ID,A.����,A.No as ���ݺ�,A.ʵ��Ʊ�� As Ʊ�ݺ�, " & vbNewLine & _
    "       A.���,A.��������,A.�շ�ϸĿID,A.ִ�в���ID,A.�շ����,P.���, " & vbNewLine & _
    "       C.���� as ����,Nvl(B.����,C.����) as ����,E1.���� as ��Ʒ��,C.���," & vbNewLine & _
    "       A.����, A.����,C.���㵥λ," & vbNewLine & _
    "       ltrim(to_char(A.����,'9999990.00000')) as ����," & vbNewLine & _
    "       ltrim(to_char(A.Ӧ�ս��,'9999990.00')) as Ӧ�ս��," & vbNewLine & _
    "       ltrim(to_char(A.ʵ�ս��,'9999990.00')) as ʵ�ս��," & vbNewLine & _
    "       A.������,A.����ʱ��, A.����ID,A.ҽ��, A.����,A.�����, " & vbNewLine & _
    "       A.�������,A.ת����,A.ת��ʱ��" & vbNewLine & _
    "From (" & strTable & ") A,�շ���ĿĿ¼ C,�շ���Ŀ���� B,�շ���Ŀ���� E1,�շ���� P" & _
    " Where A.�շ�ϸĿID=C.ID And A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    "       and A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
    "       And A.�շ����=P.����(+)" & _
    " Order by A.����,A.ʵ��Ʊ��,A.NO,A.���"
    
    If mrsHistoryList Is Nothing Or blnFilter = False Then
        Set mrsHistoryList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, strNos)
    Else
        mrsHistoryList.Filter = 0
    End If
    vsHistory.Redraw = flexRDNone
    vsHistory.Clear: vsHistory.Cols = 0
    Set vsHistory.DataSource = mrsHistoryList
    If vsHistory.Rows <= 1 Then vsHistory.Rows = 2
    
    With vsHistory
        For lngCol = 0 To .Cols - 1
             .ColAlignment(lngCol) = flexAlignLeftCenter
             .FixedAlignment(lngCol) = flexAlignCenterCenter
              .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
              If .ColKey(lngCol) Like "*ID" Or .ColKey(lngCol) = "��������" Or .ColKey(lngCol) = "ת����־" Or .ColKey(lngCol) = "�շ����" Then
                    .ColHidden(lngCol) = True
              ElseIf .ColKey(lngCol) Like "*��*" Or .ColKey(lngCol) Like "*��*" Or .ColKey(lngCol) Like "*��" Then
                    .ColAlignment(lngCol) = flexAlignRightCenter
              End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsHistory, Me.Caption, IIf(mint���� = 1, "��ʷ�˷��б�", "��ʷ�����б�"), True
        '����
        Dim strNO As String, str���� As String
        For lngRow = 1 To .Rows - 1
            If strNO <> Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) _
                 And strNO <> "" Then
                '�����ָ���
                .Select lngRow, .FixedCols, lngRow, .Cols - 1
                .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
            End If
            .Cell(flexcpData, lngRow, .ColIndex("���ݺ�")) = .TextMatrix(lngRow, .ColIndex("���ݺ�"))
            strNO = Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�")))
            str���� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
        Next
        .Editable = flexEDNone
    End With
    If blnFilter = False Then zlCommFun.StopFlash
    vsHistory.Redraw = flexRDBuffered
    
    Screen.MousePointer = 0
    ReadHistoryListData = True
    Exit Function
errHandle:
    vsHistory.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
   If blnFilter = False Then zlCommFun.StopFlash
End Function

Private Sub vsBalance_DblClick()
    Dim i As Integer
    With vsBalance
        'If .TextMatrix(.Row, 0) = "�տ����" Then Exit Sub
        If .Cell(flexcpFontUnderline, .Row, .Col, .Row, .Col) = True Then
            If Val(.TextMatrix(.Row, .Col + 1)) = 0 Then Exit Sub
            For i = 0 To .Cols - 1
                If IsNumeric(.TextMatrix(0, i)) = False Then
                If InStr(mstrStyle, .TextMatrix(0, i)) > 0 Then
                    txtSum.Text = Val(txtSum.Text) + Val(.TextMatrix(0, i + 1))
                    .TextMatrix(0, i + 1) = "0"
                End If
                End If
            Next i
        End If
    End With
End Sub

Private Sub vsFee_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsFee
        Select Case Col
        Case .ColIndex(mstr��־)
            txtSum.Text = 0
            SetNOBill .TextMatrix(Row, .ColIndex("����")), .TextMatrix(Row, .ColIndex("���ݺ�")), Val(.TextMatrix(Row, .Col)) <> 0
            mblnSel = Val(.TextMatrix(Row, .Col)) <> 0
            Call SetRowSelected(Row)
            mblnChange = True
            Call CalcSUMMony
            Call SetBlanceShow
            If mblnSel = False Then mblnSel = IsCheckSelNo
        Case Else
        End Select
    End With
End Sub

Private Sub vsFee_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, IIf(mint���� = 1, "�˷��б�", "�����б�"), True
End Sub

Private Sub vsFee_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
   Dim cur�ϼ� As Currency, i As Long
    If NewRow <> OldRow Then
'        With vsFee
'            If .TextMatrix(NewRow, .ColIndex("���ݺ�")) <> "" Then
'                For i = NewRow - 1 To .FixedRows Step -1
'                    If .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(NewRow, .ColIndex("���ݺ�")) Then
'                        cur�ϼ� = cur�ϼ� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
'                    Else
'                        Exit For
'                    End If
'                Next
'                For i = NewRow To .Rows - 1
'                    If .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(NewRow, .ColIndex("���ݺ�")) Then
'                        cur�ϼ� = cur�ϼ� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
'                    Else
'                        Exit For
'                    End If
'                Next
'            End If
            Call StatusShowBillSum
            'Me.stbThis.Panels(2).Text = "��ǰ���ݺϼ�:" & Format(cur�ϼ�, gstrDec)
'        End With
    End If
End Sub

Private Sub vsFee_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, IIf(mint���� = 1, "�˷��б�", "�����б�"), True
End Sub

Private Sub vsFee_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsFee
        Select Case Col
        Case .ColIndex(mstr��־)
            If CheckIsInput(Row) = False Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub
 

Private Sub vsFee_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ݻ��ߺ������
    '����:���˺�
    '����:2011-01-26 09:57:32
    '˵��:
    '       1.OwnerDrawҪ����ΪOver(������Ԫ��������)
    '       2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
    '       3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    Dim strText As String
    strText = " "
    With vsFee
        '����������еı��߼�����
        lngLeft = .ColIndex(mstr��־): lngRight = .ColIndex(mstr��־)
        
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        Call GetBillNOStartAndEndRow(Row, lngBegin, lngEnd)
        If lngBegin = lngEnd Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '���б����±���
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, strText, 1, 0
        Done = True
    End With
End Sub

Private Function SysColor2RGB(ByVal lngColor As Long) As Long
'���ܣ���VB��ϵͳ��ɫת��ΪRGBɫ
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

Private Sub GetBillNOStartAndEndRow(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������
    '����:���˺�
    '����:2011-01-26 10:01:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    lngBegin = lngRow: lngEnd = lngRow
    With vsFee
        For i = lngRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(lngRow, .ColIndex("���ݺ�")) Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
        For i = lngRow To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(lngRow, .ColIndex("���ݺ�")) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
    End With
End Sub
Private Function SetNOBill(ByVal str���� As String, ByVal strNO As String, ByVal blnSel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ȫѡ��ȫ�嵥��
    '���:str����-��������(�շѵ�,���ʵ�)
    '       strNO-ָ����NO
    '        blnSel:true��ʾȫѡ,����ȫ��
    '����:
    '����:
    '����:���˺�
    '����:2011-01-24 10:47:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsFee
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" _
                And .TextMatrix(i, .ColIndex("���ݺ�")) = strNO Then
                .TextMatrix(i, .ColIndex(mstr��־)) = IIf(blnSel, -1, 0)
            End If
        Next
    End With
    SetNOBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function CheckMulitBillValied(ByVal strNO As String, ByVal lngInsure As Long, _
    ByRef strOutNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���൥���շѵ��Ƿ�Ϸ�
    '���:strNO-���ݺ�
    '����:strOutNos-���صĶ൥��,����Ϊ:A0001,A002...
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-02-24 14:10:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNos As String, strTemp As String, strNo1 As String
    Dim i As Long, m As Long, varTemp As Variant
    Dim strNOsTemp As String
    On Error GoTo errHandle
    With vsFee
        If mint���� <> 1 Then
            '���ʵ�,������ֱ�ӷ���
            strOutNos = strNO: CheckMulitBillValied = True: Exit Function
        End If
        strNos = Replace(GetMultiNOs(strNO), "'", "") 'һ���շѵ���������
        If InStr(1, strNos, ",") = 0 Then
            '�Ƕ൥���շ�,ֱ�ӷ���
            strOutNos = strNO: CheckMulitBillValied = True: Exit Function
        End If
        strTemp = "": strNOsTemp = ""
        For i = 1 To .Rows - 1
             strNo1 = Trim(.TextMatrix(i, .ColIndex("���ݺ�")))
             If strNo1 <> strNO Then
                '1. ����Ƿ����δ��ѡ���˷ѵ�
                If InStr(1, strTemp & ",", "," & strNo1 & ",") = 0 Then
                    If InStr(1, "," & strNos & ",", "," & strNo1 & ",") > 0 Then
                         If GetVsGridBoolColVal(vsFee, i, .ColIndex(mstr��־)) = False Then
                            MsgBox "ע��:" & vbCrLf & "    ���ݺ�Ϊ" & strNo1 & "�뵥��Ϊ" & strNO & "���շѵ� " & vbCrLf & "    �Ƕ൥���շ�,���Ա���һ����!", vbInformation + vbOKOnly, gstrSysName
                             .Row = i: Exit Function
                        End If
                        strNOsTemp = strNOsTemp & "," & strNo1
                     End If
                    strTemp = strTemp & "," & strNo1
                End If
             Else
                strNOsTemp = strNOsTemp & "," & strNo1
             End If
        Next
        '2.���δ��ȡ�����ĵ���
        varTemp = Split(strNos, ",")
        strTemp = ""
    
        For m = 0 To UBound(varTemp)
            If InStr(1, "," & strNOsTemp & ",", "," & varTemp(m) & ",") = 0 Then
                strTemp = strTemp & "," & varTemp(m)
            End If
        Next
            
            If strTemp <> "" Then
                strTemp = Mid(strTemp, 2)
                If MsgBox("ע��:" & vbCrLf & "����Ϊ" & strNO & "�Ƕ൥���շ�,���������µ���:" & vbCrLf & strTemp & vbCrLf & "    �쳣,��������Ϊ�ϴ��˷�ʱ�쳣,�Ƿ��������?", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then
                    .Row = 1: Exit Function
                End If
                If strNOsTemp <> "" Then
                    strNos = Mid(strNOsTemp, 2)
                Else
                    MsgBox "�����쳣,�����˷�!", vbOKOnly + vbInformation, gstrSysName
                    .Row = 1: Exit Function
                    Exit Function
                End If
            End If
         
        '3.�Ϸ�ɾ��,����
         strOutNos = strNos
    End With
    CheckMulitBillValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExecuteDelBill(ByVal strDelDate As String, ByVal strNos As String, intInsure As Integer, ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ������˷Ѳ���
    '���:strNos-���ݺ�:�����Ƕ൥��
    '       lngInsure-����
    '����:ִ�гɹ�������true,���򷵻�False
    '����:���˺�
    '����:2011-02-24 15:35:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, k As Long, varTemp  As Variant, strAllBalance      As String, strBalance As String, bln������� As Boolean
    Dim blnҽ���ӿڴ�ӡƱ�� As Boolean, bln�൥��һ�ν��� As Boolean, blnYB�������� As Boolean, bln�˷Ѻ��ӡ�ص� As Boolean
    Dim lng����ID As Long, cllPro As Collection, blnTrans As Boolean, lng����ID As Long, str������ˮ�� As String, str����˵�� As String
    Dim lng����ID1 As Long, varBalance As Variant, strAdvance As String, strInvoice As String
    Dim strSQL As String, j As Long, blnTransMedicare As Boolean, rsTmp As ADODB.Recordset, blnҽ�������� As Boolean, blnTurnAll As Boolean
    Dim str���㷽ʽ As String, cur������ As Currency, cur�ɷ���� As Currency, cur����� As Currency, cur��� As Currency, cur�˿�ϼ� As Currency
    Dim strDelNOs As String, lng����ID As Long, blnExecuteThreeSwap As Boolean
    
    If intInsure <> 0 Then
        blnҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, , intInsure, CStr(lng����ID))
        bln�൥��һ�ν��� = Not (gclsInsure.GetCapability(83, , intInsure) Or gclsInsure.GetCapability(85, , intInsure))
        blnYB�������� = gclsInsure.GetCapability(support�����������, , intInsure)
        If blnYB�������� = False Then
            MsgBox "ע��:" & vbCrLf & "   ���ݺ�Ϊ" & strNos & "�ĵ���,��֧��ҽ����������,����"
            Exit Function
        End If
        bln�˷Ѻ��ӡ�ص� = gclsInsure.GetCapability(support�˷Ѻ��ӡ�ص�, , intInsure)
    End If
    
    If intInsure <> 0 And blnҽ���ӿڴ�ӡƱ�� Then
        Dim strUserType As String
        Dim lngShareUseID As Long
        If mrsInfo Is Nothing Then
            lng����ID = mlng����ID
        ElseIf mrsInfo.State <> 1 Then
            lng����ID = mlng����ID
        Else
            lng����ID = Val(Nvl(mrsInfo!����ID))
        End If
        strUserType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
        lngShareUseID = zl_GetInvoiceShareID(1121, strUserType)
         
        lng����ID = GetInvoiceGroupID(1, 1, lng����ID, lngShareUseID)
        Select Case lng����ID
            Case -1
                MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Exit Function
            Case -2
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Exit Function
        End Select
        strInvoice = GetNextBill(lng����ID)
    End If
    
    '��ȡ����ID
    Err = 0: On Error GoTo errHandle
    Set cllPro = New Collection
    varTemp = Split(strNos, ",")
    For i = 0 To UBound(varTemp)
            'Zl_����תסԺ_�շ�ת��
            strSQL = "Zl_����תסԺ_�շ�ת��("
            '     No_In         סԺ���ü�¼.NO%Type,
            strSQL = strSQL & "'" & varTemp(i) & "',"
            '     ����Ա���_In סԺ���ü�¼.����Ա���%Type,
            strSQL = strSQL & "'" & UserInfo.��� & "',"
            '     ����Ա����_In סԺ���ü�¼.����Ա����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '     �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
            strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
            '     �����˷�_In   Number := 0(�����˷�_In:0-����תסԺ��������;1-�����˷�ģʽ:Ϊ1ʱ:��Ժ����id_In����ҳID_IN���Բ���)
            strSQL = strSQL & "1,"
            '     ��Ժ����id_In סԺ���ü�¼.��������id%Type := Null,
            strSQL = strSQL & "Null,"
            '     ��ҳid_In     סԺ���ü�¼.��ҳid%Type := Null
            strSQL = strSQL & "Null,"
            '     ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null
            strSQL = strSQL & IIf(picBack.Visible, "'" & cboStyle.Text & "'", "Null") & ","
           With vsFee
                lng����ID1 = 0
                For j = 1 To .Rows - 1
                    If .TextMatrix(j, .ColIndex("���ݺ�")) = varTemp(i) Then
                        lng����ID1 = Val(.TextMatrix(j, .ColIndex("����ID")))
                        Exit For
                    End If
                Next
           End With
           strAllBalance = strAllBalance & "," & lng����ID1
          cllPro.Add Array(strSQL, lng����ID1, varTemp(0), CStr(varTemp(0)), varTemp(i))
    Next
    
    
     If intInsure <> 0 And bln�൥��һ�ν��� Then
        On Error GoTo errH: blnTrans = True
        gcnOracle.BeginTrans
            '�����һ�ſ�ʼ��
        For i = cllPro.Count To 1 Step -1
            If InStr("," & mstrUsedBills & ",", "," & Val(cllPro(i)(1)) & ",") = 0 Then
                blnExecuteThreeSwap = False
                lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
                If mcur��� <> 0 Then
                    Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng����ID & "," & Val(cllPro(i)(1)) & "," & mcur��� & ")", Me.Caption)
                    mcur��� = 0
                Else
                    Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng����ID & "," & Val(cllPro(i)(1)) & ")", Me.Caption)
                End If
                
                If ExecuteThreeSwap(Val(cllPro(i)(1)), lng����ID, str������ˮ��, str����˵��) = True Then
                    blnExecuteThreeSwap = True
                End If
                
                'Zl_����תסԺ_����������
                strSQL = "Zl_����תסԺ_����������("
                '  No_In         סԺ���ü�¼.NO%Type,
                strSQL = strSQL & "'" & varTemp(i - 1) & "',"
                '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                strSQL = strSQL & "'" & UserInfo.��� & "',"
                '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                strSQL = strSQL & "'" & UserInfo.���� & "',"
                '  �˷�ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ����˷�ʱ,ÿ�ŵ��ݵ��˷�ʱ����ͬ,����ϵͳ��ǰʱ��
                strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
                '  �����˷�_In   Number := 0,
                strSQL = strSQL & "" & 1 & ","
                '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
                strSQL = strSQL & "Null,"
                '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                strSQL = strSQL & "Null,"
                '  �����˷�_In   Number := 0,
                strSQL = strSQL & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
                '  ����ID_In     סԺ���ü�¼.����id%Type)
                strSQL = strSQL & "" & lng����ID & ")"
                Call zlDatabase.ExecuteProcedure(strSQL, "����������")
                mstrUsedBills = mstrUsedBills & "," & Val(cllPro(i)(1))
            End If
        Next
        
        '�Ȳ���Ʊ�ݣ�ҽ���ӿڲ���ȡ��
        If blnҽ���ӿڴ�ӡƱ�� Then
            strSQL = "zl_�����շѼ�¼_RePrint('" & CStr(cllPro(1)(3)) & "','" & strInvoice & "'," & ZVal(lng����ID) & ",'" & UserInfo.���� & "'," & _
                "To_Date('" & Format(strDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
        strAdvance = strAllBalance
        If Not gclsInsure.ClinicDelSwap(Val(cllPro(cllPro.Count)(1)), , intInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            MsgBox "ҽ������ʧ�ܣ��޷������������ת��Ժ������", vbInformation, gstrSysName
            Exit Function
        Else
            blnTransMedicare = True
        End If

        If Not (strAdvance = strAllBalance Or strAdvance = "") Then
            '���ݷ��صĽ�����Ϣ������Ԥ����¼��strAdvance���ظ�ʽ:���㷽ʽ1|���||���㷽ʽ2:���...
            '�ȷ�̯��ÿ�ŵ�����
            Set rsTmp = GetBalanceSet
            varBalance = Split(strAdvance, "||")
            For i = 0 To UBound(varBalance)
                str���㷽ʽ = Split(varBalance(i), "|")(0)
                cur������ = -1 * Val(Split(varBalance(i), "|")(1))
                For k = 0 To UBound(varTemp)
                    cur�ɷ���� = Getʵ�ս��(varTemp(k))
                    rsTmp.Filter = "�������=" & k
                    For j = 1 To rsTmp.RecordCount
                        cur�ɷ���� = cur�ɷ���� - rsTmp!������
                        rsTmp.MoveNext
                    Next
                    If cur�ɷ���� > 0 Then
                        If cur�ɷ���� <= cur������ Then
                            cur������ = cur������ - cur�ɷ����
                        Else
                            cur�ɷ���� = cur������
                            cur������ = 0
                        End If
                        rsTmp.AddNew
                        rsTmp!������� = k
                        rsTmp!���㷽ʽ = str���㷽ʽ
                        rsTmp!������ = cur�ɷ����
                        rsTmp.Update
                        
                        If cur������ = 0 Then Exit For
                    End If
                Next
            Next
            
            For k = 0 To UBound(varTemp)
                strBalance = ""
                cur����� = 0
                cur��� = Getʵ�ս��(varTemp(k))
                
                rsTmp.Filter = "�������=" & k
                For i = 1 To rsTmp.RecordCount
                    strBalance = IIf(strBalance = "", "", strBalance & "||") & rsTmp!���㷽ʽ & "|" & -1 * rsTmp!������
                    cur��� = cur��� - rsTmp!������
                    rsTmp.MoveNext
                Next

                '��Ϊָ���Ľ��㷽ʽ��������ֽ𣬿��ܲ����µ������
                'If cbo�˿ʽ.ItemData(cbo�˿ʽ.ListIndex) = 1 Then
                    cur������ = Format(CentMoney(cur���), "0.00")
                    cur����� = cur������ - cur���
'                Else
'                    cur������ = cur���
'                End If
                cur�˿�ϼ� = cur�˿�ϼ� + cur������
                lng����ID = GetDelBalanceID(varTemp(k))
                strSQL = "zl_�����շѽ���_Update(" & lng����ID & ",'" & "�ֽ�" & "|" & -1 * cur������ & "| ',0,'" & strBalance & "'," & -1 * cur����� & ")"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            Next
        End If
        gcnOracle.CommitTrans: blnTrans = False
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
     Else
         '�����һ�ſ�ʼ��
        For i = cllPro.Count To 1 Step -1
            On Error GoTo errH
            blnExecuteThreeSwap = False
            blnҽ�������� = False: blnTurnAll = False
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
            If intInsure <> 0 Then
                blnҽ�������� = IsYBSingle(CStr(cllPro(i)(4)), intInsure)
            Else
                blnTurnAll = CheckAllTurn(CStr(cllPro(i)(4)))
                If InStr("," & mstrUsedBills & ",", "," & Val(cllPro(i)(1)) & ",") > 0 Then blnTurnAll = True
            End If
            If blnҽ�������� Or (intInsure = 0 And Not blnTurnAll) Then
                If InStr("," & mstrUsedBills & ",", "," & Val(cllPro(i)(1)) & ",") = 0 Then
                    gcnOracle.BeginTrans: blnTrans = True
                    If mcur��� <> 0 Then
                        Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng����ID & ",Null," & mcur��� & ")", Me.Caption)
                        mcur��� = 0
                    Else
                        Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng����ID & ")", Me.Caption)
                    End If
                    
                    blnTransMedicare = False
                    If intInsure <> 0 Then                    '����ҽ���ӿ�
                          If blnYB�������� Then
                                strAdvance = lng����ID & "|" & "0" & "|" & CStr(cllPro(i)(4))
                                If Not gclsInsure.ClinicDelSwap(CStr(cllPro(i)(1)), True, intInsure, strAdvance) Then
                                    gcnOracle.RollbackTrans
                                    MsgBox "ҽ������ʧ�ܣ��޷������������ת��Ժ������", vbInformation, gstrSysName
                                    Exit Function
                                Else
                                    blnTransMedicare = True
                                End If
                            End If
                    End If
                    gcnOracle.CommitTrans: blnTrans = False
                    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
                    
                    If ExecuteThreeSwap(Val(cllPro(i)(1)), lng����ID, str������ˮ��, str����˵��) = True Then
                        blnExecuteThreeSwap = True
                    End If
                    
                    'Zl_����תסԺ_����������
                    strSQL = "Zl_����תסԺ_����������("
                    '  No_In         סԺ���ü�¼.NO%Type,
                    strSQL = strSQL & "'" & varTemp(i - 1) & "',"
                    '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                    strSQL = strSQL & "'" & UserInfo.��� & "',"
                    '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                    strSQL = strSQL & "'" & UserInfo.���� & "',"
                    '  �˷�ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ����˷�ʱ,ÿ�ŵ��ݵ��˷�ʱ����ͬ,����ϵͳ��ǰʱ��
                    strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
                    '  �����˷�_In   Number := 0,
                    strSQL = strSQL & "" & 1 & ","
                    '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
                    strSQL = strSQL & "Null,"
                    '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                    strSQL = strSQL & "Null,"
                    '  �����˷�_In   Number := 0,
                    strSQL = strSQL & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
                    '  ����ID_In     סԺ���ü�¼.����id%Type)
                    strSQL = strSQL & "" & lng����ID & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, "����������")
                    
                    strDelNOs = strDelNOs & IIf(strDelNOs = "", "", ",") & cllPro(i)(0)
                End If
            Else
                If InStr("," & mstrUsedBills & ",", "," & Val(cllPro(i)(1)) & ",") = 0 Then
                    gcnOracle.BeginTrans: blnTrans = True
                    If mcur��� <> 0 Then
                        Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng����ID & "," & Val(cllPro(i)(1)) & "," & mcur��� & ")", Me.Caption)
                        mcur��� = 0
                    Else
                        Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng����ID & "," & Val(cllPro(i)(1)) & ")", Me.Caption)
                    End If
                    
                    blnTransMedicare = False
                    If intInsure <> 0 Then                    '����ҽ���ӿ�
                          If blnYB�������� Then
                                strAdvance = lng����ID & "|" & "0"
                                If Not gclsInsure.ClinicDelSwap(CStr(cllPro(i)(1)), True, intInsure, strAdvance) Then
                                    gcnOracle.RollbackTrans
                                    MsgBox "ҽ������ʧ�ܣ��޷������������ת��Ժ������", vbInformation, gstrSysName
                                    Exit Function
                                Else
                                    blnTransMedicare = True
                                End If
                            End If
                    End If
                    gcnOracle.CommitTrans: blnTrans = False
                    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
                    
                    If ExecuteThreeSwap(Val(cllPro(i)(1)), lng����ID, str������ˮ��, str����˵��) = True Then
                        blnExecuteThreeSwap = True
                    End If
                    
                    'Zl_����תסԺ_����������
                    strSQL = "Zl_����תסԺ_����������("
                    '  No_In         סԺ���ü�¼.NO%Type,
                    strSQL = strSQL & "'" & varTemp(i - 1) & "',"
                    '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                    strSQL = strSQL & "'" & UserInfo.��� & "',"
                    '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                    strSQL = strSQL & "'" & UserInfo.���� & "',"
                    '  �˷�ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ����˷�ʱ,ÿ�ŵ��ݵ��˷�ʱ����ͬ,����ϵͳ��ǰʱ��
                    strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
                    '  �����˷�_In   Number := 0,
                    strSQL = strSQL & "" & 1 & ","
                    '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
                    strSQL = strSQL & "Null,"
                    '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                    strSQL = strSQL & "Null,"
                    '  �����˷�_In   Number := 0,
                    strSQL = strSQL & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
                    '  ����ID_In     סԺ���ü�¼.����id%Type)
                    strSQL = strSQL & "" & lng����ID & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, "����������")
                    
                    strDelNOs = strDelNOs & IIf(strDelNOs = "", "", ",") & cllPro(i)(0)
                    mstrUsedBills = mstrUsedBills & "," & Val(cllPro(i)(1))
                End If
            End If
        Next
     End If
     
    If intInsure <> 0 And bln�˷Ѻ��ӡ�ص� And InStr(1, mstrPrivs, ";ҽ���˷ѻص�;") > 0 Then
        '����:35248
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me, "NO=" & strNos, 2)
    End If
    ExecuteDelBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    
    If blnTrans Then
        'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, intInsure)
    End If
    
    If Err.Number <> 0 Then
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    
    '�ж���ʾ,����ӡ�������˷Ѻ��ٴ�ӡ���Լ�ѡ���ش�
    If strDelNOs <> "" Then
        MsgBox "����[" & strNos & "]�˷�ʧ�ܡ����ǣ�����[" & strDelNOs & "]�ѳɹ��˷ѡ�" & vbCrLf & _
            "����δ��ӡ�����ִ��ʧ�ܵĵ��������˷ѣ�", vbInformation, gstrSysName
    End If
    Exit Function
End Function

Private Function ExecuteThreeSwap(lngBalance As Long, lng����ID As Long, Optional ByRef str������ˮ�� As String, Optional ByRef str����˵�� As String) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset, strBalanceIDs As String, rsTotal As ADODB.Recordset
    Dim dblMoney As Double, strAll As String, strDetail() As String, strItem() As String, strCardNo As String
    Dim i As Integer, lngCardID As Long
    If mobjSquare Is Nothing Then Set mobjSquare = gobjSquare.objSquareCard
    If mobjSquare Is Nothing Then Exit Function
    strSQL = _
        "Select ժҪ" & vbNewLine & _
        "    From ����Ԥ����¼" & vbNewLine & _
        "    Where ���㷽ʽ Is Null And ��¼���� = 3 And ��¼״̬ = 2 And ����id = [1]"
   
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    If rsTemp.RecordCount = 0 Then Exit Function
    strAll = Nvl(rsTemp!ժҪ)
    If strAll = "" Then Exit Function
    
    strDetail = Split(strAll, "|")
    For i = 0 To UBound(strDetail)
        If strDetail(i) <> "" Then
            strItem = Split(strDetail(i), ",")
            If Val(strItem(0)) = 1 Then
                lngCardID = Val(strItem(1))
                dblMoney = -1 * Val(strItem(2))
                strSQL = "Select Distinct a.����id" & vbNewLine & _
                            "From ������ü�¼ A" & vbNewLine & _
                            "Where a.No In (Select Distinct a.No From ������ü�¼ A Where Mod(a.��¼����, 10) = 1 And a.����id = [1]) And Mod(a.��¼����, 10) = 1 And" & vbNewLine & _
                            "      a.��¼״̬ <> 0"
                strSQL = "Select Min(����ID) As ����ID,Min(����) As ���� From ����Ԥ����¼ Where ����ID IN (" & strSQL & ") And �����ID = [2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lngCardID)
                strBalanceIDs = "3|" & Val(Nvl(rsTemp!����ID))
                If mobjSquare.zlReturnCheck(Me, mlngModule, lngCardID, False, Nvl(rsTemp!����), _
                    strBalanceIDs, dblMoney, str������ˮ��, str����˵��, "3|" & lng����ID) = False Then Exit Function
                If mobjSquare.zlReturnMoney(Me, mlngModule, lngCardID, False, Nvl(rsTemp!����), _
                    strBalanceIDs, dblMoney, str������ˮ��, str����˵��, "3|" & lng����ID) = False Then Exit Function
            End If
        End If
    Next i
    
    ExecuteThreeSwap = True
End Function

Public Function GetBalanceSet() As ADODB.Recordset
'���ܣ�����һ�������¼������
    Dim rsTmp As New ADODB.Recordset
       
    rsTmp.Fields.Append "�������", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "���㷽ʽ", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "������", adCurrency, , adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set GetBalanceSet = rsTmp
End Function

Public Function Getʵ�ս��(ByVal strNO As String) As Currency
    Dim i As Long, cur��� As Currency
    With vsFee
        cur��� = 0
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) = strNO Then
                cur��� = cur��� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
            End If
        Next
        Getʵ�ս�� = cur���
    End With
End Function
Private Function ExecuteWirteOff(strDELDae As String, ByVal cllDel As Variant) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�������������
    '����:���˺�
    '����:2011-02-25 10:22:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strSQL As String
    Dim cllPro As Collection
    Set cllPro = New Collection
    For i = 1 To cllDel.Count
        'Zl_����תסԺ_����ת��
        strSQL = "Zl_����תסԺ_����ת��("
        '  No_In         סԺ���ü�¼.NO%Type,
        strSQL = strSQL & "'" & cllDel(i)(0) & "',"
        '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type
        strSQL = strSQL & "To_Date('" & strDELDae & "','yyyy-mm-dd hh24:mi:ss'),"
        '   ��������_In   Number := 0
        '   --��������_In:0-����תסԺ��������;1-��������˷�ģʽ
        strSQL = strSQL & "1)"
        zlAddArray cllPro, strSQL
    Next
    On Error GoTo errHandle
    zlExecuteProcedureArrAy cllPro, Me.Caption
    ExecuteWirteOff = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ʻ��˷�
    '����:�˷ѻ����ʳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-02-23 11:21:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng���� As Long, lng����ID As Long
    Dim strOutNos As String, strTemp As String, strDelDate As String
    Dim m As Long, i As Long, blnHaveData As Boolean, blnPrintList As Boolean '�Ƿ��ӡ�嵥
    Dim cllDelNO As Collection, strDelNOs As String, lngRow As Long, strNO As String
    Dim lng����ID As Long
    
    strDelDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    blnPrintList = False
    If InStr(mstrPrivs, ";��ӡ�嵥;") > 0 And mint���� = 1 Then
        Select Case mint�շ��嵥    '0-����ӡ,1-Ҫ��ӡ,2-ѡ���Ƿ��ӡ
        Case 2
             If MsgBox("Ҫ��ӡ�շ��嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                blnPrintList = True
             End If
        Case 1
            blnPrintList = True
        End Select
    End If
    mstrUsedBills = ""
    With vsFee
        If .Rows <= 1 Then Exit Function
        If .Cols <= 1 Then Exit Function
        Set cllDelNO = New Collection
        strTemp = ""
        For lngRow = 1 To .Rows - 1
            '���ʵ���
            strNO = Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�")))
            If CheckBillExistReplenishData(1, , strNO) And mint���� = 1 Then
                MsgBox "ѡ��ĵ��ݴ��ڲ�������¼���޷������˷ѣ�", vbInformation, gstrSysName
                Exit Function
            End If
            If GetVsGridBoolColVal(vsFee, lngRow, .ColIndex(mstr��־)) _
                And strNO <> "" And InStr(1, "," & strTemp & ",", "," & strNO & ",") = 0 Then
                lng���� = Val(.TextMatrix(lngRow, .ColIndex("����")))
                lng����ID = Val(.TextMatrix(lngRow, .ColIndex("����ID")))
                strOutNos = ""
'                If CheckMulitBillValied(strNo, lng����, strOutNos) = False Then
'                    Exit Function
'                End If
                If lng���� <> 0 And IsYBSingle(strNO, lng����) = False Then
                    If CheckInsureAll(lng����ID) = False Then
                        MsgBox "ѡ��ĵ��ݴ�������δ�˷ѵ��ݣ��޷������˷ѣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                blnHaveData = False
                For i = 1 To cllDelNO.Count
                    If cllDelNO(i)(0) = strNO Then
                        blnHaveData = True: Exit For
                    End If
                    If InStr(1, "," & cllDelNO(i)(1) & ",", "," & strNO & ",") > 0 Then
                        blnHaveData = True: Exit For
                    End If
                    If lng���� <> 0 Then
                        If IsYBSingle(strNO, lng����) = False Then
                            If Val(cllDelNO(i)(3)) = lng����ID Then
                                blnHaveData = True: Exit For
                            End If
                        End If
                    End If
                Next
                If blnHaveData = False Then
                    '�������ʵ���
                    cllDelNO.Add Array(strNO, strOutNos, lng����, lng����ID)
                End If
                strTemp = strTemp & "," & strNO & "," & strOutNos

            End If
        Next
    End With
    'ִ�о������ʻ��˷Ѳ���
    If cllDelNO.Count = 0 Then
        MsgBox "ע��:" & vbCrLf & "    û��ѡ��һ����Ҫ�����˷ѻ����ʵĵ���,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '�˷�
    strDelNOs = ""
    If mint���� = 2 Then
        If ExecuteWirteOff(strDelDate, cllDelNO) = False Then Exit Function
    Else
        For i = 1 To cllDelNO.Count
            If ExecuteDelBill(strDelDate, IIf(cllDelNO(i)(1) <> "", cllDelNO(i)(1), cllDelNO(i)(0)), Val(cllDelNO(i)(2)), Val(cllDelNO(i)(2))) = False Then
                    Exit Function
            End If
            strDelNOs = strDelNOs & "," & IIf(cllDelNO(i)(1) <> "", cllDelNO(i)(1), cllDelNO(i)(0))
        Next
    End If
    If strDelNOs <> "" Then strDelNOs = Mid(strDelNOs, 2)
    '��ӡ�����嵥
    If blnPrintList And mint���� = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & "'" & Replace(strDelNOs, ",", "','") & "'", "ҩƷ��λ=" & IIf(mblnҩ����λ, 1, 0), 2)
    End If
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetLocaleNO(ByVal str���� As String, ByVal strNO As String, ByVal blnSelect As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ����NO
    '����:���˺�
    '����:2011-02-09 14:56:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    With vsFee
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) = strNO Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(blnSelect, -1, 0)
            End If
        Next
    End With
End Sub
Private Function CheckIsInput(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ������������
    '���:lngRow-ָ������
    '����:
    '����:��Ч,����true,���򷵻�False
    '����:���˺�
    '����:2011-02-09 15:04:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strNO As String, i As Long, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant
    Dim lng����ID As Long, str���� As String
    
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    lng����ID = Val(Nvl(mrsInfo!����ID))
    With vsFee
            intInsure = Val(.TextMatrix(lngRow, .ColIndex("����")))
            strNO = .TextMatrix(lngRow, .ColIndex("���ݺ�"))
            str���� = .TextMatrix(lngRow, .ColIndex("����"))
            If intInsure > 0 And str���� = "�շѵ�" Then
                If Not gclsInsure.GetCapability(support�����������, lng����ID, intInsure) Then
                    stbThis.Panels(2).Text = "����[" & strNO & "]�Ĳ������಻֧�������������,���в�����ѡ��ת��!"
                    Exit Function
                Else
                    '���жϸõ��ݵ�ÿ�ֽ��㷽ʽ�Ƿ�֧��,�����˷�ʱ,������Ϊָ�����㷽ʽ,�˴��򻯹���Ϊ�������˷�
                    strTmp = GetBalanceType(strNO)
                    If strTmp <> "" Then
                        arrBalanceType = Split(strTmp, ",")
                        For i = 0 To UBound(arrBalanceType)
                            strBalanceType = arrBalanceType(i)
                            If Not gclsInsure.GetCapability(support�����������, lng����ID, intInsure, strBalanceType) Then
                                stbThis.Panels(2).Text = "����[" & strNO & "]�Ĳ������಻֧��" & strBalanceType & "����,���в�����ѡ��ת��!"
                                Exit Function
                            End If
                        Next
                    End If
                End If
            End If
    End With
    CheckIsInput = True
End Function
Private Function SetRowSelected(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ�е�ѡ��״̬
    '       ����Ƕ��ŵ����е�һ��,����ͬʱ���ö����е���������
    '����:���˺�
    '����:2011-02-09 14:50:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strNO As String, i As Long, strTmp As String
    Dim blnSelect As Boolean, lng����ID As Long, str���� As String
    lng����ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng����ID = Val(Nvl(mrsInfo!����ID))
        End If
    End If
    With vsFee
        intInsure = Val(.TextMatrix(lngRow, .ColIndex("����")))
        blnSelect = GetVsGridBoolColVal(vsFee, lngRow, .ColIndex(mstr��־))
        str���� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
        If intInsure > 0 And str���� = "�շѵ�" Then 'ȫ��ѡ���ȡ��
            If Not IsYBSingle(.TextMatrix(lngRow, .ColIndex("���ݺ�")), intInsure) Then
                If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
            End If
        Else '�ֽ�����Ҫ����൥���շ����
            If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
        End If
    End With
    SetRowSelected = True
End Function

Private Function CheckInsureAll(lngBalance As Long) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, blnFound As Boolean
    strSQL = "Select Distinct a.No" & vbNewLine & _
            "From ������ü�¼ A, ������ü�¼ B" & vbNewLine & _
            "Where b.����id = [1] And a.No = b.No And Mod(a.��¼����,10) = Mod(b.��¼����,10)" & vbNewLine & _
            "Group By a.No" & vbNewLine & _
            "Having Sum(a.ʵ�ս��) <> 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalance)
    Do While Not rsTmp.EOF
        blnFound = False
        With vsFee
            For i = 1 To .Rows - 1
                If GetVsGridBoolColVal(vsFee, i, .ColIndex(mstr��־)) Then
                    If Trim(.TextMatrix(i, .ColIndex("���ݺ�"))) = Trim(rsTmp!NO) Then blnFound = True: Exit For
                End If
            Next i
            If blnFound = False Then
                CheckInsureAll = False
                Exit Function
            End If
        End With
        rsTmp.MoveNext
    Loop
    CheckInsureAll = True
End Function

Private Function GetBalanceType(ByVal strNO As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ�ŵ����е�ҽ�����㷽ʽ��
    '����:ҽ�����㷽ʽ��
    '����:���˺�
    '����:2011-02-09 15:01:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long
    On Error GoTo errH
    strSQL = "Select A.���㷽ʽ From ����Ԥ����¼ A, ���㷽ʽ B" & vbNewLine & _
            "Where A.���㷽ʽ = B.���� And B.���� In (3, 4) And A.NO = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    For i = 1 To rsTmp.RecordCount
        GetBalanceType = GetBalanceType & "," & rsTmp!���㷽ʽ
        rsTmp.MoveNext
    Next
    GetBalanceType = Mid(GetBalanceType, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckAllTurn(ByVal strNO As String) As Boolean
    Dim strSQL As String, rsData As ADODB.Recordset
    strSQL = "Select 1" & vbNewLine & _
            " From ����Ԥ����¼ A," & vbNewLine & _
            "     (Select Distinct ����id" & vbNewLine & _
            "       From ������ü�¼" & vbNewLine & _
            "       Where NO In (Select Distinct NO" & vbNewLine & _
            "                    From ������ü�¼" & vbNewLine & _
            "                    Where ����id In" & vbNewLine & _
            "                          (Select ����id" & vbNewLine & _
            "                           From ����Ԥ����¼" & vbNewLine & _
            "                           Where ������� In (Select b.�������" & vbNewLine & _
            "                                          From ������ü�¼ A, ����Ԥ����¼ B" & vbNewLine & _
            "                                          Where a.No = [1] And a.��¼���� = 1 And a.��¼״̬ <> 0 And a.����id = b.����id))) And" & vbNewLine & _
            "             ��¼���� = 1 And ��¼״̬ <> 0) B" & vbNewLine & _
            " Where a.����id = b.����id And a.��¼���� = 3 And (Exists (Select 1 From ҽ�ƿ���� Where ID = a.�����id And �Ƿ�ȫ�� = 1) Or Exists" & vbNewLine & _
            "       (Select 1 From ���ѿ����Ŀ¼ Where ��� = a.���㿨��� And �Ƿ�ȫ�� = 1))" & vbNewLine & _
            " Group By ���㷽ʽ" & vbNewLine & _
            " Having Sum(��Ԥ��) <> 0"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsData.EOF Then
        CheckAllTurn = False
    Else
        CheckAllTurn = True
    End If
End Function

Private Function SetMultiOther(ByVal lngRow As Long, blnSelect As Boolean, intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ŵ�������ѡ���ȡ��
    '       ���ҽ�����ŵ���Ҫ�������˷�,ѡ������һ��ʱ,ȫѡ����,ȡ��ʱȫȡ��
    '���:lngRow-��ǰ��
    '        blnSelect-�Ƿ�ѡ��
    '        intInsure-����
    '����:
    '����:���˺�
    '����:2011-02-09 15:41:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, k As Long, strNO As String, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant
    Dim lng����ID As Long, str���� As String, blnAllTurn As Boolean
    lng����ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng����ID = Val(Nvl(mrsInfo!����ID))
        End If
    End If
    With vsFee
        str���� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
        If intInsure = 0 Then
            If CheckAllTurn(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) = True Then
                blnAllTurn = True
            Else
                blnAllTurn = False
            End If
            If mblnMultiBalance Or blnAllTurn Then     '   �൥��,���ֽ��㷽ʽ
                '33635:ԭ���Ƕ൥���Ҷ��ֽ��㷽ʽ,���ܲ�����
                strNO = ""
                For k = 1 To .Rows - 1
                      If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
                        And Trim(.TextMatrix(lngRow, .ColIndex("����ID"))) <> "" _
                        And .TextMatrix(k, .ColIndex("����")) = str���� Then
                          If InStr(1, "," & strNO & ",", "," & .TextMatrix(k, .ColIndex("���ݺ�")) & ",") = 0 Then
                                strNO = strNO & "," & .TextMatrix(k, .ColIndex("���ݺ�"))
                          End If
                      End If
                Next
                If strNO <> "" Then strNO = Mid(strNO, 2)
                If InStr(1, strNO, ",") > 0 Then    '֤��Ϊ�൥��
                    'һԺҪ��,ֻҪ�Ƕ൥�ݽ����,��תʱ,������ȫת
                    'If CheckSingleBalance(strNo) = False Then    '�Ƕ��ֽ��㷽ʽ,�������˷�,'ȫѡ
                        For k = 1 To .Rows - 1
                              If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
                                  And Trim(.TextMatrix(lngRow, .ColIndex("����ID"))) <> "" _
                                   And .TextMatrix(k, .ColIndex("����")) = str���� Then
                                    .TextMatrix(k, .ColIndex(mstr��־)) = IIf(blnSelect, -1, 0)
                              End If
                        Next
                    'End If
                End If
            End If
            '����Ƿ�������ѿ��Ľ���,�������,�ֲ�֧���ⲿ�����ݵĴ���
            If strNO = "" Then strNO = .TextMatrix(lngRow, .ColIndex("���ݺ�"))
'            If str���� = "�շѵ�" Then
'                If zlIsExistsSquareCard(strNO) Then
'                    stbThis.Panels(2).Text = "�ݲ�֧�ֶ����ѿ����ݵ��������תסԺ����!"
'                    For k = 1 To .Rows - 1
'                          If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) And Trim(.TextMatrix(lngRow, .ColIndex("����ID"))) <> "" Then
'                                .TextMatrix(k, .ColIndex(mstr��־)) = 0
'                          End If
'                    Next
'                End If
'            End If
            '����Ƿ�������ѿ�,����൥���д������ѿ�,Ҳ����ȫѡ
            SetMultiOther = True
            Exit Function
        End If
        If IsYBSingle(vsFee.TextMatrix(lngRow, .ColIndex("���ݺ�")), intInsure) Then SetMultiOther = True: Exit Function
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
                And i <> lngRow And .TextMatrix(i, .ColIndex("����")) = str���� Then
                If GetVsGridBoolColVal(vsFee, i, .ColIndex(mstr��־)) <> GetVsGridBoolColVal(vsFee, lngRow, .ColIndex(mstr��־)) Then
                   If intInsure <> 0 And str���� = "�շѵ�" And blnSelect Then
                        strNO = .TextMatrix(i, .ColIndex("���ݺ�"))
                        '�жϸõ��ݵ�ÿ�ֽ��㷽ʽ�Ƿ�֧��,�����˷�ʱ,������Ϊָ�����㷽ʽ,�˴��򻯹���Ϊ�������˷�
                         strTmp = GetBalanceType(strNO)
                         If strTmp <> "" Then
                             arrBalanceType = Split(strTmp, ",")
                             For j = 0 To UBound(arrBalanceType)
                                 strBalanceType = arrBalanceType(j)
                                 If Not gclsInsure.GetCapability(support�����������, lng����ID, intInsure, strBalanceType) Then
                                     stbThis.Panels(2).Text = "����[" & strNO & "]�Ĳ������಻֧��" & strBalanceType & "����,���в�����ѡ��ת��!"
                                     For k = 1 To .Rows - 1
                                        If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(i, .ColIndex("����ID")) _
                                            And .TextMatrix(k, .ColIndex("����")) = .TextMatrix(i, .ColIndex("����")) Then
                                            .TextMatrix(k, .ColIndex(mstr��־)) = 0
                                        End If
                                     Next
                                     Exit Function
                                 End If
                             Next
                         End If
                    End If
                    .TextMatrix(i, .ColIndex(mstr��־)) = IIf(blnSelect, -1, 0)
                End If
            End If
        Next
    End With
    SetMultiOther = True
End Function
Private Function IsCheckSelNo() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���Ƿ����ѡ��
    '����:ѡ��,����true,���򷵻�False
    '����:���˺�
    '����:2011-02-23 15:41:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsFee
        For i = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsFee, i, .ColIndex(mstr��־)) Then
                IsCheckSelNo = True: Exit Function
            End If
        Next
    End With
    IsCheckSelNo = False
End Function
Private Function CheckSingleBalance(ByVal strNO As String) As Boolean
'���ܣ��ж�ָ���������Ƿ�ֻ��һ�ַ�ҽ�����㷽ʽ(��Ԥ������)
'       :strNO(��ʽΪ"E01,E02"):����:34035
'������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strNO = Replace(strNO, "'", "")
    CheckSingleBalance = True
    
    strSQL = "" & _
    " Select /*+ rule */ Count(Distinct A.���㷽ʽ) num" & vbNewLine & _
    " From ����Ԥ����¼ A, ���㷽ʽ B,Table( f_Str2list([1])) J" & vbNewLine & _
    " Where   A.��¼���� = 3 And A.��¼״̬ In (1, 3) " & _
    "           And A.���㷽ʽ = B.���� And B.���� In (1, 2)  And A.NO = J.Column_Value"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNO)
    If rsTmp!Num > 1 Then CheckSingleBalance = False
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function zlIsExistsSquareCard(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ����Ƿ�Ϊ�����㵥��
    '���:strNos-���ݺ�(����Ϊ����,�ö��ŷ���)
    '����:
    '����:����,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2010-01-11 12:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNoIns As String
    
    On Error GoTo errHandle
    
    strNoIns = Replace(strNos, "'", "")
    strSQL = "" & _
    "   Select /*+ rule */ A.ID As ������id " & _
    "   From ���˿������¼ A, ����Ԥ����¼ B,Table( f_Str2list([1])) J " & _
    "   Where A.����id = B.ID and B.��¼����=3 And B.NO = J.Column_Value And Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����շѵ��Ƿ����ˢ����¼", strNoIns)
    zlIsExistsSquareCard = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsHistory_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsHistory, Me.Caption, IIf(mint���� = 1, "��ʷ�˷��б�", "��ʷ�����б�"), True
End Sub
Private Sub vsHistory_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsHistory, Me.Caption, IIf(mint���� = 1, "��ʷ�˷��б�", "��ʷ�����б�"), True
End Sub
Private Sub SetBlanceShow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ���㷽ʽ
    '���:blnAllSel-ѡ�����еĵ���
    '����:���˺�
    '����:2011-02-23 14:54:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String, lngRow As Long, i As Long, str���� As String
    Dim blnȫѡ As Boolean, blnδѡ As Boolean, rsTmp As ADODB.Recordset
    Dim strFilter As String, bln�˿� As Boolean, strSQL As String
    Dim strSelNos As String, strNO As String, intCol As Integer
    If mint���� = 2 Then Exit Sub
    With vsFee
        blnȫѡ = True: blnδѡ = True
        For lngRow = 1 To .Rows - 1
            strNO = Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�")))
            If GetVsGridBoolColVal(vsFee, lngRow, .ColIndex(mstr��־)) Then
                If InStr(1, strSelNos & ",", "," & strNO & ",") = 0 Then
                    strSelNos = strSelNos & "," & strNO
                    blnδѡ = False
                End If
            End If
             If InStr(1, strSelNos & ",", "," & strNO & ",") = 0 Then blnȫѡ = False
        Next
    End With
    If strSelNos <> "" Then strSelNos = Mid(strSelNos, 2)
    bln�˿� = False
    '��ʾ����ѡ��ĵ��ݵĽ��㷽ʽ֮��
    If Not mrsBalance Is Nothing Then
        If blnȫѡ Or blnδѡ Then
            mrsBalance.Filter = 0
            If blnȫѡ Then bln�˿� = True
        Else
'            strFilter = Replace(strSelNos, ",", "' Or NO='")
'            strFilter = " NO='" & strFilter & "'"
'            mrsBalance.Filter = strFilter
            bln�˿� = True
        End If
        If SetPicBack(strSelNos) = True Then
            txtSum.Text = InitPatialBalance(strSelNos)
        Else
            Call InitBlanceData(strSelNos)
        End If
        mcur��� = 0
        If Val(cboStyle.ItemData(cboStyle.ListIndex)) = 1 Then
            mcur��� = Val(txtSum.Text) - CentMoney(Val(txtSum.Text))
            If mcur��� <> 0 Then
            With mrsBalance
                .AddNew
                !���㷽ʽ = "����"
                !���� = 1
                !Ӧ���� = "0"
                !��� = Format(mcur���, "0.00")
                !ժҪ = ""
                !������� = ""
                .Update
            End With
            End If
            txtSum.Text = Format(txtSum.Text - mcur���, "0.00")
        Else
            mcur��� = 0
        End If
        
        mrsBalanceBak.Filter = "��� <> 0"
        mrsBalanceBak.Sort = "����,Ӧ����,���㷽ʽ"
        mrsBalance.Sort = "����,Ӧ����,���㷽ʽ"
        vsBalance.Redraw = flexRDNone
        vsBalance.Clear 1
        vsBalance.Cols = 1
        
        If Not mrsBalanceBak.EOF Then
            For i = 1 To mrsBalanceBak.RecordCount
                If Nvl(mrsBalanceBak!���㷽ʽ, "��Ԥ��") <> strBalance Then
                    strBalance = Nvl(mrsBalanceBak!���㷽ʽ, "��Ԥ��")
                    vsBalance.Cols = vsBalance.Cols + 2
                    vsBalance.ColAlignment(vsBalance.Cols - 2) = 7
                    vsBalance.ColAlignment(vsBalance.Cols - 1) = 1
                End If
                If mrsBalanceBak!���� <> 1 Then
                    vsBalance.Cell(flexcpFontBold, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = True  '����
                    vsBalance.Cell(flexcpForeColor, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = vbBlue
                ElseIf bln�˿� Then
                    vsBalance.Cell(flexcpFontBold, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = True  '����
                    vsBalance.Cell(flexcpForeColor, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = vbBlue  '��ɫ:�˿�
                End If
                vsBalance.TextMatrix(0, vsBalance.Cols - 2) = strBalance & ":"
                vsBalance.TextMatrix(0, vsBalance.Cols - 1) = _
                    Val(vsBalance.TextMatrix(0, vsBalance.Cols - 1)) + Nvl(mrsBalanceBak!���, 0)
                    '�൥��ʹ�ö��ֽ���ʱ,���ʽ����û�н��зֱҴ���,���Բ�����formatȡ��λ��
                'vsBalance.ColData(vsBalance.Cols - 2) = "ժҪ:" & mrsBalanceBak!ժҪ
                vsBalance.ColData(vsBalance.Cols - 1) = "�������:" & mrsBalanceBak!�������
                mrsBalanceBak.MoveNext
            Next
        End If
        
        intCol = 0
        strBalance = ""
        If Not mrsBalance.EOF Then
            For i = 1 To mrsBalance.RecordCount
                If Nvl(mrsBalance!���㷽ʽ, "��Ԥ��") <> strBalance Then
                    strBalance = Nvl(mrsBalance!���㷽ʽ, "��Ԥ��")
                    intCol = intCol + 2
                    vsBalance.ColAlignment(intCol - 1) = 7
                    vsBalance.ColAlignment(intCol) = 1
                End If
                If mrsBalance!���� <> 1 Then
                    vsBalance.Cell(flexcpFontBold, 1, intCol, 1, intCol - 1) = True '����
                    vsBalance.Cell(flexcpForeColor, 1, intCol, 1, intCol - 1) = IIf(bln�˿�, vbRed, vbBlue) '��ɫ
                ElseIf bln�˿� Then
                    vsBalance.Cell(flexcpFontBold, 1, intCol, 1, intCol - 1) = True '����
                    vsBalance.Cell(flexcpForeColor, 1, intCol, 1, intCol - 1) = vbRed '��ɫ:�˿�
                End If
                vsBalance.TextMatrix(1, intCol - 1) = strBalance & ":"
                If Nvl(mrsBalance!���㷽ʽ) = "����" Then
                    vsBalance.TextMatrix(1, intCol) = _
                        Format(Val(vsBalance.TextMatrix(1, intCol)) + Nvl(mrsBalance!���, 0), "0.00")
                Else
                    vsBalance.TextMatrix(1, intCol) = _
                        Val(vsBalance.TextMatrix(1, intCol)) + Nvl(mrsBalance!���, 0)
                End If
                vsBalance.ColData(intCol) = "�������:" & mrsBalance!�������
                mrsBalance.MoveNext
            Next
        End If
        If strSelNos = "" Then
            For i = 1 To vsBalance.Cols - 1
                vsBalance.TextMatrix(1, i) = ""
            Next i
        End If
        Call vsBalance.AutoSize(0, vsBalance.Cols - 1)
        vsBalance.Row = vsBalance.FixedRows
        If vsBalance.Cols <> 1 Then vsBalance.Col = vsBalance.FixedCols
        'vsBalance.TextMatrix(0, 0) = IIf(bln�˿�, "�˿����", "�տ����")
        vsBalance.Redraw = flexRDDirect
    End If
End Sub
Public Sub CalcSUMMony()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    '����:���˺�
    '����:2011-03-11 18:09:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cur��� As Currency
    With vsFee
        cur��� = 0
        For i = .FixedRows To .Rows - 1
            If GetVsGridBoolColVal(vsFee, i, .ColIndex(mstr��־)) Then
                cur��� = cur��� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
            End If
        Next
        lblSum.Caption = "��ǰת���ϼ�:" & Format(cur���, "###0.00;-###0.00;0.00;0.00")
        mcur�ϼ� = cur���
    End With
End Sub
Public Sub StatusShowBillSum()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    '����:���˺�
    '����:2011-03-11 18:09:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cur��� As Currency, dbl��Ʊ��� As Double, strNO As String, str��Ʊ�� As String
    Dim strTemp As String
    
    With vsFee
        strTemp = "": dbl��Ʊ��� = 0: cur��� = 0
        If Not (.Row > .Rows - 1 Or .Row < 1) Then
            strNO = .TextMatrix(.Row, .ColIndex("���ݺ�"))
            str��Ʊ�� = .TextMatrix(.Row, .ColIndex("Ʊ�ݺ�"))
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("���ݺ�")) = strNO Then
                        cur��� = cur��� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
                End If
                If .TextMatrix(i, .ColIndex("Ʊ�ݺ�")) = str��Ʊ�� Then
                        dbl��Ʊ��� = dbl��Ʊ��� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
                End If
            Next
            strTemp = "����(" & strNO & ")�ϼ�:" & Format(cur���, "###0.00;-###0.00;0.00;0.00")
            strTemp = strTemp & "  ��Ʊ(" & str��Ʊ�� & ")�ϼ�:" & Format(dbl��Ʊ���, "###0.00;-###0.00;0.00;0.00")
        End If
        stbThis.Panels(2).Text = strTemp
    End With
End Sub

Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨����������Ϣ
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Set objCard = IDKind.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    gobjSquare.bln��ȱʡ������ = IDKind.Cards.��ȱʡ������
End Sub
Private Sub zlCreateObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������¼�����
    '����: �����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-08-28 16:16:00
    '˵��:
    '����:54896
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '������������
    Err = 0: On Error Resume Next
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
         Set mobjICCard.gcnOracle = gcnOracle
    End If
    
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    
End Sub
Private Sub zlCloseObject()
    '�ر���ض���
    Err = 0: On Error Resume Next
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
    End If
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
    End If
    Set mobjIDCard = Nothing
    Set mobjICCard = Nothing
End Sub


