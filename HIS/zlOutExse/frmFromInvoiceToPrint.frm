VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFromInvoiceToPrint 
   Caption         =   "���ݷ�Ʊ���ش�Ʊ"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   Icon            =   "frmFromInvoiceToPrint.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmFromInvoiceToPrint.frx":058A
   ScaleHeight     =   8325
   ScaleWidth      =   11865
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   45
      ScaleHeight     =   660
      ScaleWidth      =   11700
      TabIndex        =   14
      Top             =   7170
      Width           =   11700
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10095
         TabIndex        =   17
         Top             =   225
         Width           =   1440
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "�ش�(&O)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8565
         TabIndex        =   16
         Top             =   195
         Width           =   1440
      End
      Begin VB.TextBox txtInvoice 
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
         IMEMode         =   3  'DISABLE
         Left            =   765
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   180
         Width           =   2175
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ��"
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
         Left            =   270
         TabIndex        =   19
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   3195
         TabIndex        =   18
         Top             =   285
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin VB.PictureBox picCon 
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
      Height          =   675
      Left            =   180
      ScaleHeight     =   675
      ScaleWidth      =   14475
      TabIndex        =   5
      Top             =   135
      Width           =   14475
      Begin VB.CommandButton cmdBrush 
         Caption         =   "ˢ��(&N)"
         Height          =   375
         Left            =   9195
         TabIndex        =   8
         Top             =   150
         Width           =   1245
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
         Left            =   1305
         MaxLength       =   64
         TabIndex        =   7
         Top             =   150
         Width           =   2040
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   375
         Left            =   615
         TabIndex        =   6
         Top             =   150
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   661
         Appearance      =   2
         IDKindStr       =   $"frmFromInvoiceToPrint.frx":0B14
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
         MustSelectItems =   "����"
         BackColor       =   -2147483633
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   345
         Left            =   5505
         TabIndex        =   9
         Top             =   180
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   146800643
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   345
         Left            =   7515
         TabIndex        =   10
         Top             =   180
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   146800643
         CurrentDate     =   36588
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "���Է���ʱ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3690
         TabIndex        =   11
         Top             =   210
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Left            =   7200
         TabIndex        =   13
         Top             =   225
         Width           =   240
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
         Left            =   150
         TabIndex        =   12
         Top             =   195
         Width           =   480
      End
   End
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   5025
      ScaleHeight     =   2295
      ScaleWidth      =   5400
      TabIndex        =   3
      Top             =   2280
      Width           =   5400
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2685
         Left            =   105
         TabIndex        =   4
         Top             =   75
         Width           =   4650
         _cx             =   8202
         _cy             =   4736
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFromInvoiceToPrint.frx":0BC7
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
   Begin VB.PictureBox PicDetail 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   5250
      ScaleHeight     =   2775
      ScaleWidth      =   5535
      TabIndex        =   1
      Top             =   3840
      Width           =   5535
      Begin VSFlex8Ctl.VSFlexGrid vsDetail 
         Height          =   2685
         Left            =   180
         TabIndex        =   2
         Top             =   330
         Width           =   8505
         _cx             =   15002
         _cy             =   4736
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFromInvoiceToPrint.frx":0BDD
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
   Begin VB.PictureBox picInvoice 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   150
      ScaleHeight     =   4695
      ScaleWidth      =   4380
      TabIndex        =   0
      Top             =   1065
      Width           =   4380
      Begin VB.TextBox txtFilterInvoice 
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
         IMEMode         =   3  'DISABLE
         Left            =   900
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Width           =   2175
      End
      Begin VSFlex8Ctl.VSFlexGrid vsInvoiceList 
         Height          =   2685
         Left            =   300
         TabIndex        =   23
         Top             =   660
         Width           =   8505
         _cx             =   15002
         _cy             =   4736
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFromInvoiceToPrint.frx":0BF3
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
      Begin VB.Label lblInVoice 
         AutoSize        =   -1  'True
         Caption         =   "��Ʊ��"
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
         Left            =   255
         TabIndex        =   21
         Top             =   60
         Width           =   630
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   7965
      Width           =   11865
      _ExtentX        =   20929
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
            Picture         =   "frmFromInvoiceToPrint.frx":0C96
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15849
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
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   1500
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFromInvoiceToPrint.frx":152A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFromInvoiceToPrint.frx":187E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   30
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmFromInvoiceToPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mPanel
    Pane_Search = 1
    Pane_List = 2
    Pane_Detail = 3
    Pane_inVoiceList = 4
End Enum
'-----------------------------------------------------------------------------------
'���㿨���
Private mSquareCard As SquareCard '���㿨���
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Private mrsInfo As ADODB.Recordset
Private mstrFindNO As String, mstrFindFpNo As String
Private mrsList As ADODB.Recordset  '�����б�
Private mrsDetail As ADODB.Recordset
Private mrsBalance As ADODB.Recordset
Private mstrNOs As String
Private mlngModule As Long
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mblnValid As Boolean
Private mblnSel As Boolean
Private mstrPrivs As String
Private mintSucces As Integer  '�ɹ���ӡ����
Private mlng����ID As Long
Private mblnFirst As Boolean
Private mblnNotClick As Boolean
Private mlngShareUseID As Long '������������ID
Private mstrUseType As String 'ʹ�����
Private mintInvoiceFormat As Integer  '��ӡ�ķ�Ʊ��ʽ,��Ʊ��ʽ���
Private mintOldInvoiceFormat As Integer '�ɷ�Ʊ��ӡ�ĸ�ʽ
Private mblnStartFactUseType As Boolean   '�Ƿ�������ʹ�����
Private mintInvoicePrint As Integer  '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
Private mlng����ID As Long
Private mintInsure As Integer
Private mrsInVoice As ADODB.Recordset
Private mblnNotChange As Boolean

Public Function zlRePrintBill(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal strPrivs As String, Optional lng����ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ش�Ʊ�����
    '����:��ӡ�ɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-09-04 22:39:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs: mintSucces = 0
    mlng����ID = lng����ID
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlRePrintBill = mintSucces > 0
End Function
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, strKey As String
    Dim objTemp As Object
    With dkpMan
        .ImageList = imlPaneIcons
        Set objPane = .CreatePane(mPanel.Pane_Search, 200, 100, DockLeftOf, Nothing)
        objPane.Tag = mPanel.Pane_Search
        objPane.Title = "��������": objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoCaption
        objPane.MaxTrackSize.Height = 675 \ Screen.TwipsPerPixelY
        objPane.MinTrackSize.Height = 675 \ Screen.TwipsPerPixelY
        objPane.Handle = picCon.hWnd
        
        Set objTemp = .CreatePane(mPanel.Pane_inVoiceList, 100, 100, DockBottomOf, objPane)
        objTemp.Tag = mPanel.Pane_inVoiceList
        objTemp.Title = "��Ʊ��Ϣ�б�": objTemp.Options = PaneNoCloseable Or PaneNoHideable
        objTemp.Handle = picInvoice.hWnd
        
        Set objPane = .CreatePane(mPanel.Pane_List, 300, 100, DockRightOf, objTemp)
        objPane.Tag = mPanel.Pane_List
        objPane.Title = "�շѵ����б�": objPane.Options = PaneNoCloseable Or PaneNoHideable
        objPane.Handle = picList.hWnd
 
        Set objPane = .CreatePane(mPanel.Pane_Detail, 300, 100, DockBottomOf, objPane)
        objPane.Tag = mPanel.Pane_Detail
        objPane.Title = "������ϸ�б�": objPane.Options = PaneNoCloseable Or PaneNoHideable
        objPane.Handle = PicDetail.hWnd
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    dkpMan.RecalcLayout: DoEvents
    zlRestoreDockPanceToReg Me, dkpMan, "����"
    'Call GetRegInFor(g˽��ģ��, Me.Name, "����", strKey)
    'If Val(strKey) = 1 Then mPanSearch.Hide
        
End Function
Private Sub chkDate_Click()
    dtpBegin.Enabled = chkDate.Value <> 1
    dtpEnd.Enabled = chkDate.Value <> 1
End Sub
Private Sub cmdBrush_Click()
        mblnNotChange = True
        Call ReadInVoice(False)
        mblnNotChange = False
End Sub
Private Sub cmdCancel_Click()
     Unload Me
End Sub
 
Private Function zlMakeupPrint(ByVal strNos As String, ByVal strReclaimInvoice As String, _
    Optional ByVal blnMediCare As Boolean, _
    Optional ByVal blnDel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ��
    '���:strReclaimInvoice-���յ�Ʊ�ݺ�
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-04 22:58:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String
    Dim intInsure As Integer, blnVirtualPrint As Boolean, lng����ID As Long, lng����ID As Long
    Dim strUseType  As String, lngShareUseID As Long, intInvoiceFormat As Integer
    Dim intOldInvoiceFormat As Integer
    If strNos = "" Then Exit Function
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    
    lng����ID = Val(Nvl(mrsInfo!����ID))
    If lng����ID = 0 Then Exit Function
    If strNos = "" Then
        MsgBox "��ǰû�е��ݿ����ش�Ʊ�ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    If blnMediCare Then
        intInsure = ChargeExistInsure(strNos, lng����ID, lng����ID, , blnDel)
        If intInsure <> 0 Then
            blnVirtualPrint = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure, CStr(lng����ID))
        End If
    End If
    strUseType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
    lngShareUseID = zl_GetInvoiceShareID(mlngModule, strUseType)
    intInvoiceFormat = zl_GetInvoicePrintFormat(mlngModule, strUseType, intOldInvoiceFormat)
    '������ʣ�������Ĳſ����ش򣬱���ҽ������ʹ������Ҳ�������´�ӡ
    If Not blnVirtualPrint Then
        If Not BillExistMoney(strNos, 1) Then
            MsgBox "�����е���Ŀ�Ѿ�ȫ���˷ѣ����ܽ��д�ӡ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    'Ƚ����,2014-12-17,���������շѵ��ݲ������ش�Ʊ��
    If CheckBillExistReplenishData(1, , strNos) = True Then
        MsgBox "�����е���Ŀ�Ѿ������˱��ղ�����㣬���ܽ��д�ӡ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    Dim dtDate As Date, strPriceGrade As String
    dtDate = zlDatabase.Currentdate
    '��ȡ�۸�ȼ�
    If gintPriceGradeStartType >= 2 Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, lng����ID, 0, "", , , strPriceGrade)
    Else
        strPriceGrade = gstr��ͨ�۸�ȼ�
    End If
    Call frmPrint.ReportPrint(2, strNos, "", strReclaimInvoice, mlng����ID, mlngShareUseID, txtInvoice.Text, dtDate, "", "", _
        gTy_Module_Para.bln�ֱ��ӡ, mintInvoiceFormat, blnVirtualPrint, , mstrUseType, , , , strPriceGrade)
    
    '��ҽһ��ͨд����85950
    Call WriteInforToCard(Me, mlngModule, mstrPrivs, mSquareCard.objSquareCard, 0, strNos)
    
    mintSucces = mintSucces + 1
    zlMakeupPrint = True
End Function

Private Function CheckFP() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鷢Ʊ�Ƿ���ȷ
    '����: ��Ʊ�Ϸ� ����true,���򷵻�False
    '����:���˺�
    '����:2012-07-12 11:30:22
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intNum As Integer, varData As Variant
     On Error GoTo errHandle
     If Not gblnStrictCtrl Then
        If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
            MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytFactLength & " λ��", vbInformation, gstrSysName
            txtInvoice.SetFocus: Exit Function
        End If
        CheckFP = True
        Exit Function
     End If
    If Trim(txtInvoice.Text) = "" Then
        MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
        txtInvoice.SetFocus: Exit Function
    End If
    intNum = 1
    If Not gTy_Module_Para.bln�ֱ��ӡ Then intNum = 1
 
InvoiceHandle:
    If zlCheckInvoiceValied(mlng����ID, intNum, txtInvoice.Text, mlngShareUseID, mstrUseType) = False Then
        Exit Function
    End If
    '�����������,Ʊ���Ƿ�����
    If CheckBillRepeat(mlng����ID, 1, txtInvoice.Text) Then
            'Tag�����⣺24363:���˺飺��Ҫ�ǽ���Զ����ɵĺ��Ƿ��û����ģ���Ҫ�����
            If txtInvoice.Locked = False And txtInvoice.Tag <> Trim(txtInvoice.Text) Then
                MsgBox "Ʊ�ݺ�""" & txtInvoice.Text & """�Ѿ���ʹ�ã����������롣", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Function
            Else
                Call RefreshFact
                If txtInvoice.Text = "" Then
                    txtInvoice.SetFocus: Exit Function
                Else
                    MsgBox "��ǰƱ�ݺ��Ѿ���ʹ�ã������»�ȡƱ�ݺ�:" & txtInvoice.Text, vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Function
                End If
            End If
    End If
   CheckFP = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
     Dim strReclaimInvoice As String '����Ʊ��
     Dim strNos As String, strNo As String, blnYb As Boolean
     Dim lngRow As Long
     
     With vsInvoiceList
        strReclaimInvoice = .TextMatrix(.Row, .ColIndex("��Ʊ��"))
        If strReclaimInvoice = "" Then
            MsgBox "δѡ����Ҫ�ش��Ʊ��", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        If MsgBox("��ȷ��Ҫ�ش�Ʊ��Ϊ��" & strReclaimInvoice & "����Ʊ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
     End With
     With vsList
        For lngRow = 1 To .Rows - 1
            strNo = Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�")))
            If strNo <> "" And InStr(strNos & ",", "," & strNo & ",") = 0 Then
               If blnYb = False Then blnYb = .TextMatrix(lngRow, .ColIndex("ҽ��")) = "��"
                strNos = strNos & "," & strNo
            End If
        Next
     End With
    If Not CheckFP() Then Exit Sub
    Call zlMakeupPrint(strNos, strReclaimInvoice, blnYb, False)
    Call RefreshFact
     Call zlClearPatiInfor
     gblnOK = True
End Sub
 

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True: Exit Sub
    If Action = PaneActionFloating Then Cancel = True: Exit Sub
    If Action = PaneActionPinning Then Cancel = True: Exit Sub
    If Action = PaneActionCollapsing Then Cancel = True: Exit Sub
    If Action = PaneActionAttaching Then Cancel = True: Exit Sub
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pane_Search    '1
        Item.Handle = picCon.hWnd
    Case Pane_List      ' 2
        Item.Handle = picList.hWnd
    Case Pane_Detail    '3
        Item.Handle = PicDetail.hWnd
    Case Pane_inVoiceList  ' 4
        Item.Handle = picInvoice.hWnd
    End Select
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
     Bottom = stbThis.Height + picDown.Height
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    Call zlClearPatiInfor
    mblnFirst = False
    If mlng����ID <> 0 Then
        If GetPatient(IDKind.GetCurCard, "-" & mlng����ID, False) = False Then
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            zlControl.TxtSelAll txtPatient
            Exit Sub
        End If
         vsList.SetFocus: Exit Sub
    End If
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
    Exit Sub
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mblnStartFactUseType = zlStartFactUseType(1)
    mlng����ID = 0
    lblFormat.Alignment = 0

    dtpBegin.MaxDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59")
    dtpEnd.MaxDate = dtpBegin.MaxDate
    dtpBegin.Value = Format(DateAdd("d", -7, dtpBegin.MaxDate), "yyyy-mm-dd")
    dtpEnd.Value = Format(dtpEnd.MaxDate, "yyyy-mm-dd")
    Call InitPanel
    Call zlCardSquareObject
    Call zlCreateObject
End Sub
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With picDown
        .Top = ScaleHeight - stbThis.Height - .Height
        .Width = ScaleWidth
        .Left = ScaleLeft
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call zlCardSquareObject(True)
    zlSaveDockPanceToReg Me, dkpMan, "����"
    zl_vsGrid_Para_Save mlngModule, vsInvoiceList, Me.Caption, "��Ʊ�б�", False
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "�����б�", False
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "������ϸ�б�", False
    Call zlCloseObject
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Then Exit Sub
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
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
    If mSquareCard.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub
 

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

 
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub picCon_Resize()
    Err = 0: On Error Resume Next
    With picCon
        cmdBrush.Left = .ScaleWidth - cmdBrush.Width - 50
    End With
End Sub
Private Sub picDown_Resize()
    With picDown
        cmdCancel.Left = .ScaleLeft + .ScaleWidth - cmdCancel.Width - 50
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        vsList.Left = .ScaleLeft
        vsList.Width = .ScaleWidth
        vsList.Height = .ScaleHeight
        vsList.Top = .ScaleTop
    End With
End Sub
Private Sub PicDetail_Resize()
    Err = 0: On Error Resume Next
    With PicDetail
        vsDetail.Left = .ScaleLeft
        vsDetail.Width = .ScaleWidth
        vsDetail.Height = .ScaleHeight
        vsDetail.Top = .ScaleTop
    End With
End Sub

Private Sub picInvoice_Resize()
    Err = 0: On Error Resume Next
    With picInvoice
        vsInvoiceList.Top = txtFilterInvoice.Top + txtFilterInvoice.Height + 20
        vsInvoiceList.Left = .ScaleLeft
        vsInvoiceList.Width = .ScaleWidth
        vsInvoiceList.Height = .ScaleHeight - vsInvoiceList.Top
    End With
End Sub
 Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("���֤", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub
Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������رս��㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If mSquareCard Is Nothing Then
         Set mSquareCard = New SquareCard
    End If
    'ֻ��:ִ�л��˷�ʱ,�ſ��ܹܽ��㿨��
    If blnClosed Then
       If Not mSquareCard.objSquareCard Is Nothing Then
            Call mSquareCard.objSquareCard.CloseWindows
            Set mSquareCard.objSquareCard = Nothing
        End If
        Set mSquareCard = Nothing
        Exit Sub
    End If
    
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    Set mSquareCard.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        Err = 0: On Error GoTo 0:      Exit Sub
    End If
    If mSquareCard.objSquareCard Is Nothing Then Exit Sub
    Dim objCard As Card
    
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
       
   '��װ�˽��㿨�Ĳ���
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   '����:zlInitComponents (��ʼ���ӿڲ���)
   '    ByVal frmMain As Object, _
   '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
   '        ByVal cnOracle As ADODB.Connection, _
   '        Optional blnDeviceSet As Boolean = False, _
   '        Optional strExpand As String
   '����:
   '����:   True:���óɹ�,False:����ʧ��
   '����:���˺�
   '����:2009-12-15 15:16:22
   'HIS����˵��.
   '   1.���������շ�ʱ���ñ��ӿ�
   '   2.����סԺ����ʱ���ñ��ӿ�
   '   3.����Ԥ����ʱ
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   If mSquareCard.objSquareCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
        '��ʼ�������ɹ�,����Ϊ�����ڴ���
        Exit Sub
   End If
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
        "Select A.����ID,Nvl(B.��ҳID,0) as ��ҳID,A.����� as �����,A.��ǰ����,B.��Ժ����,A.����,A.�Ա�,Nvl(B.����,A.����) as ����,A.IC����,A.���￨��,A.����֤��," & _
        "       Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�,C.���� as ��ǰ����,A.��ǰ����ID,D.���� as ��Ժ����,B.��Ժ����ID, A.���� as ����,E.����,E.ҽ����,E.����," & _
        "       A.�Ǽ�ʱ��,Nvl(B.״̬,0) as ״̬,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,Nvl(B.��˱�־,0) as ��˱�־,B.��Ժ����,B.��Ժ����,B.��������,NVL(A.��������,B.��������) as ��������" & _
        " From ������Ϣ A,������ҳ B,���ű� C,���ű� D,ҽ�����˵��� E,ҽ�����˹����� F" & _
        " Where A.ͣ��ʱ�� is NULL And A.����ID=B.����ID(+) And A.��ҳID=B.��ҳID(+) " & _
        "           And A.����ID=F.����ID(+) And F.��־(+)=1 And F.ҽ����=E.ҽ����(+) And F.����=E.����(+) And F.���� = E.����(+)" & _
        "           And A.��ǰ����ID=C.ID(+) And B.��Ժ����ID=D.ID(+)" & _
        "           And A.ͣ��ʱ�� is NULL "
    
    If blnCard = True And objCard.���� Like "����*" And InStr("-+*", Left(strInput, 1)) = 0 Then  '103563
        lng�����ID = IDKind.GetDefaultCardTypeID
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If mSquareCard.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.����ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSQL = strSQL & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(������Ժ)
        strSQL = strSQL & " And A.סԺ��=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSQL = strSQL & " And A.�����=[1]"
        '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "." Or objCard.���� = "���ݺ�" Then
        '���ݺŲ���
        If Left(strInput, 1) = "." Then
            strTemp = UCase(GetFullNO(Mid(strInput, 2), 13))
        Else
            strTemp = UCase(GetFullNO(strInput, 13))
        End If
        txtPatient.Text = strTemp
        gstrSQL = "" & _
        "   Select  distinct A.����ID " & _
        "   From ������ü�¼ A " & _
        "   Where A.NO=[1] and A.��¼����=[2] " & _
        "              And Rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp, 1)
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
            Case "����", "��������￨"
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
                '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.סԺ��=[2]"
            Case "��Ʊ��"
                strSQL = "" & _
                "   Select distinct A.����ID " & _
                "   From ������ü�¼ A,Ʊ�ݴ�ӡ���� B,Ʊ��ʹ����ϸ C" & _
                "   Where A.NO=B.NO and A.��¼����=1 and A.��¼״̬=1  " & _
                "               and  B.��������=1 And B.ID=C.��ӡID and C.Ʊ��=1 And C.����=1 And C.����=[1] And Rownum=1 " & _
                "   "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput)
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
                txtInvoice.Text = strInput
                GetPatient = True
                Exit Function
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If mSquareCard.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If mSquareCard.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
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
        '75259�����ϴ�,2014-7-10������������ɫ����
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(IsNull(mrsInfo!����), txtPatient.ForeColor, vbRed))
        txtPatient.Text = Nvl(mrsInfo!����)
        txtPatient.PasswordChar = ""
        txtPatient.IMEMode = 0
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!����֤��)
        GetPatient = True
        Exit Function
    Else
        Call zlClearPatiInfor
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
NotFoundPati:
  Call zlClearPatiInfor
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Function

Private Sub zlClearPatiInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������Ϣ
    '����:���˺�
    '����:2011-09-04 18:09:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtPatient.Text = ""
    ': txtOld.Text = "": txtSex.Text = ""
    'txtסԺ��.Text = "":
    Set mrsInfo = New ADODB.Recordset
    vsList.Clear 1: vsList.Rows = 2: vsDetail.Clear 1: vsDetail.Rows = 2
    vsInvoiceList.Clear 1: vsInvoiceList.Rows = 2
End Sub

 
Private Sub txtFilterInvoice_Change()
    If mblnNotChange Then Exit Sub
   Call ReadInVoice(True)

End Sub

Private Sub txtFilterInvoice_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Exit Sub
        If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub
Private Sub txtPatient_Change()
    txtPatient.Tag = ""
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    IDKind.SetAutoReadCard (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    stbThis.Panels(2).Text = ""
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
     If txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "")
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
    '����:51488
    If (IDKind.Cards.������� = "�ո��" Or IDKind.Cards.������� = " ") And Chr(KeyAscii) = " " Then KeyAscii = 0: Exit Sub
        
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
    Else
        If IDKind.GetCurCard.���� Like "����*" Then
            '103563,ֻҪ����ĵ�һ���ַ��ǡ�-+*����������ȫ���֣�����Ϊ����ˢ��
            If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
                blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
            End If
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
                Call zlClearPatiInfor
                Exit Sub
            End If
            mblnNotChange = True
            Call ReadInVoice(False)
            mblnNotChange = False
            Exit Sub
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtPatient.Text))
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
    '����:2012-09-03 09:32:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim strSQL As String, curTotal As Currency, blnIDCard As Boolean
    Dim blnICCard As Boolean, blnMsg As Boolean
    If objCard.���� Like "IC��*" And objCard.ϵͳ And mstrPassWord <> "" Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.���� Like "*���֤*" And objCard.ϵͳ And mstrPassWord <> "" Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    
    If Not GetPatient(objCard, strInput, blnCard, blnMsg) Then
        If blnCard Then
            If Not blnMsg Then MsgBox "����ȷ��������Ϣ�������Ƿ���ȷˢ����", vbInformation, gstrSysName
            Call zlClearPatiInfor
            Exit Sub
        End If
        If Not blnMsg Then MsgBox "���ܶ�ȡ������Ϣ��", vbInformation, gstrSysName
        Call zlClearPatiInfor
        Exit Sub
    End If
    
    '��ȡ�ɹ�
    '���￨������
    If (objCard.���� Like "IC��*" Or objCard.���� Like "*���֤*") And objCard.ϵͳ = True And blnCard Then blnCard = False
    If Mid(gstrCardPass, 3, 1) = "1" And (blnCard Or blnICCard Or blnIDCard Or IDKind.GetCurCard.�ӿ���� <> 0) And mstrPassWord <> "" Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
            Call zlClearPatiInfor
             txtPatient.SetFocus: Exit Sub
        End If
    End If
    mblnNotChange = True
    Call ReadInVoice(False)
    mblnNotChange = False
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
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
Private Sub txtPatient_Validate(Cancel As Boolean)
    If IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
        mblnValid = True
        Call txtPatient_KeyPress(13)
        mblnValid = False
    End If
End Sub
 

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
 
Private Sub ShowDetail(Optional strNo As String, Optional strDate As String, _
            Optional ByVal blnDel As Boolean, Optional blnSort As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ϸ����
    '����:strDate:���ݵĵǼ�ʱ��
    '����:���˺�
    '����:2011-09-04 20:18:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, strSQL As String
    Dim lngCol As Long
    
    On Error GoTo errH
    strSQL = _
    " Select C.���� as ���,Nvl(E.����,B.����) as ����," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & "B.���," & _
            IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
    "       To_Char(Avg(Nvl(A.����,1)*" & IIf(blnDel, "-1*", "") & "A.����)" & _
            IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ",'9999990.00000') as ����, " & _
    "       A.�ѱ�,To_Char(Sum(A.��׼����)" & _
            IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as ����, " & _
    "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
    "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
    "       D.���� as ִ�п���,Max(Nvl(A.��������,B.��������)) as ����," & _
    "       Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��',9,'�쳣�շѵ�','��'||ABS(A.ִ��״̬)||'���˷�') as ˵��," & _
    "       A.��¼״̬" & _
    " From  ������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,�շ���Ŀ���� E,ҩƷ��� X" & _
              IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
    " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
    "       And A.��¼����=1 And A.NO=[1] And A.��¼״̬" & IIf(blnDel, "=2", " IN(1,3)") & _
            IIf(strDate <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
    "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
            IIf(Not gblnShowErr, " And Nvl(A.���ӱ�־,0)<>9", "") & _
    " Group by Nvl(A.�۸񸸺�,A.���),C.����,Nvl(E.����,B.����)," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� ,", "") & " B.���,A.���㵥λ,A.�ѱ�,D.����," & _
    "       A.ִ��״̬,A.��¼״̬,X.ҩƷID,X." & gstrҩ����λ & ",Nvl(X." & gstrҩ����װ & ",1)" & _
    " Order by Nvl(A.�۸񸸺�,A.���)"
    If strDate <> "" Then
        Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(strDate))
    Else
        Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, "")
    End If
    With vsDetail
        .Clear 1
        .Redraw = flexRDNone
        Set .DataSource = mrsDetail
        For lngCol = 0 To .COLS - 1
             .ColAlignment(lngCol) = flexAlignLeftCenter
             .FixedAlignment(lngCol) = flexAlignCenterCenter
              .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
              If .ColKey(lngCol) Like "*ID" Or InStr(1, ",����", "," & .ColKey(lngCol) & ",") > 0 Then
                    .ColHidden(lngCol) = True
              ElseIf .ColKey(lngCol) Like "*��*" Or .ColKey(lngCol) Like "*��*" Or .ColKey(lngCol) Like "*��" Then
                    .ColAlignment(lngCol) = flexAlignRightCenter
              End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        zl_vsGrid_Para_Restore mlngModule, vsDetail, Me.Caption, "������ϸ�б�", False
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function ReadInVoice(ByVal blnFilter As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�Ʊ��Ϣ
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-04-02 15:33:06
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, dtStartDate As Date, dtEndDate As Date
    Dim strWhere As String
    On Error GoTo errHandle
    dtStartDate = CDate("1901-01-01")
    dtEndDate = dtStartDate
    If mrsInfo Is Nothing Then
        lng����ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = 0
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    If chkDate.Value = 0 Then
        strWhere = strWhere & " And A.����ʱ��+0 betWeen [2] and [3]"
        dtStartDate = CDate(Format(dtpBegin.Value, "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpEnd.Value, "yyyy-mm-dd") & " 23:59:59")
    End If
    
    If blnFilter = False Or mrsInVoice Is Nothing Then
        gstrSQL = "" & _
        "   Select  C.���� as ��Ʊ��,C.ʹ����,to_char(C.ʹ��ʱ��,'yyyy-mm-dd hh24:mi:ss') as ʹ��ʱ��," & _
        "               Sum(nvl(Q.ʵ�ս��,0)) as  ��Ʊ���" & _
        "   From (     Select  A.NO,nvl(A.�۸񸸺�,A.���) as ���,sum(ʵ�ս��) as ʵ�ս��" & _
        "                   From ������ü�¼   A" & _
        "                    Where Mod(a.��¼����,0)=1 And A.����ID=[1]  " & strWhere & _
        "                   Group by NO,nvl(A.�۸񸸺�,A.���) " & _
        "               )  Q,Ʊ�ݴ�ӡ��ϸ B,Ʊ��ʹ����ϸ C" & _
        "   Where   q.NO=B.NO  And  instr(','||B.���||',',','||Q.���||',')>0  And B.ʹ��ID=C.ID" & _
        "               And  B.Ʊ��=1 And nvl(B.�Ƿ����,0)=0 " & _
        "    Group by  C.����  ,C.ʹ����,to_char(C.ʹ��ʱ��,'yyyy-mm-dd hh24:mi:ss') " & _
        "   Order by ��Ʊ�� "
        Set mrsInVoice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, dtStartDate, dtEndDate)
    End If
    mrsInVoice.Filter = 0
    With vsInvoiceList
        .Clear 1
        .Rows = 2
        .Row = -1
        If mrsInVoice.RecordCount <> 0 Then mrsInVoice.MoveFirst
        Do While Not mrsInVoice.EOF
            If Nvl(mrsInVoice!��Ʊ��) Like "*" & txtFilterInvoice.Text & "*" _
                Or Trim(txtFilterInvoice.Text) = "" Then
                If .TextMatrix(.Rows - 1, .ColIndex("��Ʊ��")) <> "" Then .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("��Ʊ��")) = Nvl(mrsInVoice!��Ʊ��)
                .TextMatrix(.Rows - 1, .ColIndex("ʹ����")) = Nvl(mrsInVoice!ʹ����)
                .TextMatrix(.Rows - 1, .ColIndex("ʹ������")) = Nvl(mrsInVoice!ʹ��ʱ��)
                .TextMatrix(.Rows - 1, .ColIndex("��Ʊ���")) = Format(Nvl(mrsInVoice!��Ʊ���), "#######" & gstrDec)
            End If
            mrsInVoice.MoveNext
        Loop
        .Row = 1
    End With
    ReadInVoice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Function ReadListData(ByVal strInvoiceNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ҫ��ϸ����
    '���:strInvoiceNO-��Ʊ��
    '����:��ȡ�ɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-01-25 17:10:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim lngCol As Long, strSQL As String, lngRow As Long
    Dim strFilter As String, strNos As String
    Dim strWhere As String, strTable1 As String, dtStartDate As Date, dtEndDate As Date
    Dim strNo As String, i As Long, lng������� As Long
    
    mblnSel = False
    On Error GoTo errHandle
    zlCommFun.ShowFlash "���ڶ�ȡ��������,���Ժ� ..."
    Screen.MousePointer = 11
    strTable1 = " Select distinct NO From Ʊ�ݴ�ӡ��ϸ  Where Ʊ��=1 And Ʊ��=[1] "
    DoEvents
    strTable = "" & _
    "   Select  A.NO,A.ʵ��Ʊ��,A.����,A.�Ա�,A.����, " & _
    "         Decode(A.�����־,2,'',A.��ʶ��) as �����,  " & _
    "         Decode(A.�����־,2,A.��ʶ��,'') as סԺ��, " & _
    "         Min(A.�ѱ�) as �ѱ�,  " & _
    "        Sum(Decode(A.��¼״̬,2,-1,1)*A.Ӧ�ս��) as Ӧ�ս��," & _
    "        Sum(Decode(A.��¼״̬,2,-1,1)*A.ʵ�ս��) as ʵ�ս��, " & _
    "         Max(A.��¼״̬) as ����,A.����ID," & _
    "        A.������,A.��������ID,A.���ʽ,A.������,A.����Ա����,A.�Ǽ�ʱ��" & _
    "   From ������ü�¼ A,( " & strTable1 & ") B" & _
    "   Where A.��¼���� =1 and A.NO=B.NO  " & _
    "   Group by A.NO,A.ʵ��Ʊ��,A.����,A.�Ա�,A.����, " & _
    "               Decode(A.�����־,2,'',A.��ʶ��),Decode(A.�����־,2,A.��ʶ��,'')," & _
    "               A.������,A.��������ID,A.���ʽ,A.������,A.����Ա����,A.�Ǽ�ʱ��,A.����ID"
           
    strSQL = _
    " Select Decode(Nvl(Max(t.����),0),0,NULL,'��') as ҽ��,A.NO as ���ݺ�, " & _
    "       Min(A.ʵ��Ʊ��) as Ʊ�ݺ�,Max(B.����) as ��������," & _
    "       max(A.������) as ������,max(A.�����) as �����,max(A.סԺ��) as סԺ��, " & _
    "       max(C.����) as  ҽ�Ƹ��ʽ, max(A.����) as ����,max(A.�Ա�) as �Ա�,max(A.����) as ����," & _
    "       min(A.�ѱ�) as �ѱ�, " & _
    "       To_Char(max(A.Ӧ�ս��),'9999999" & gstrDec & "')   as Ӧ�ս��," & _
    "       To_Char(max(A.ʵ�ս��),'9999999" & gstrDec & "')  as ʵ�ս��," & _
    "       max(A.������) as ������,max(A.����Ա����) as ����Ա, " & _
    "       To_Char(max(A.�Ǽ�ʱ��) ,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��," & _
    "       max(A.����) as ����,A.����ID,Max(nvl(M.�������,A.����ID)) as �������ID,nvl(Max(t.����),0) as ����ID" & _
    " From (" & strTable & ") A,����Ԥ����¼ M, ���ű� B,ҽ�Ƹ��ʽ C,���ս����¼ T" & _
    " Where  A.��������ID=B.ID And A.���ʽ=C.����(+)  " & _
    "       And A.����id=t.��¼ID(+) And t.����(+)=1 And A.����id=M.����ID(+) " & _
    "       And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
    " Group by  A.No,A.����ID" & _
    " Order by �������ID desc,�Ǽ�ʱ�� Desc,���ݺ� Desc"
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInvoiceNO)
    vsList.Redraw = flexRDNone
    vsList.Clear: vsList.COLS = 2
    Set vsList.DataSource = mrsList
    If vsList.Rows <= 1 Then vsList.Rows = 2
    With vsList
        mstrNOs = ""
        For lngCol = 0 To .COLS - 1
             .ColAlignment(lngCol) = flexAlignLeftCenter
             .FixedAlignment(lngCol) = flexAlignCenterCenter
              .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
              If .ColKey(lngCol) Like "*ID" Or InStr(1, ",����,", "," & .ColKey(lngCol) & ",") > 0 Then
                    .ColHidden(lngCol) = True
              ElseIf .ColKey(lngCol) Like "*��*" Or .ColKey(lngCol) Like "*��*" Or .ColKey(lngCol) Like "*��" Then
                    .ColAlignment(lngCol) = flexAlignRightCenter
              End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "�����б�", False
        .Redraw = flexRDBuffered
        lng������� = 0
        For i = 1 To .Rows - 1
            strNo = Trim(.TextMatrix(i, .ColIndex("���ݺ�")))
            If InStr(1, mstrNOs & ",", "," & strNo & ",") = 0 Then
                mstrNOs = mstrNOs & "," & strNo
            End If
            If lng������� <> Trim(.TextMatrix(lngRow, .ColIndex("�������ID"))) _
                 And lng������� <> 0 Then
                '�����ָ���
                .Select lngRow, .FixedCols, lngRow, .COLS - 1
                .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
            End If
        Next
        vsList_AfterRowColChange 0, 0, .Row, .Col
    End With
    mrsList.Filter = "����ID>0"
    If Not mrsList.EOF Then
        mintInsure = Val(Nvl(mrsList!����ID))
    Else
        mintInsure = 0
    End If
    zlCommFun.StopFlash
    Call ReInitPatiInvoice
    Screen.MousePointer = 0
    ReadListData = True
    Exit Function
errHandle:
    vsList.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
   zlCommFun.StopFlash
End Function
Private Sub vsInvoiceList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsInvoiceList, Me.Caption, "��Ʊ�б�", False
End Sub

Private Sub vsInvoiceList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsInvoiceList, Me.Caption, "��Ʊ�б�", False
End Sub

Private Sub vsDetail_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "������ϸ�б�", False
End Sub

Private Sub vsDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "������ϸ�б�", False
End Sub
 
Private Sub vsInvoiceList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInvoiceNO As String
    If OldRow = NewRow Then Exit Sub
    If NewRow = -1 Then Exit Sub
    With vsInvoiceList
        strInvoiceNO = .TextMatrix(NewRow, .ColIndex("��Ʊ��"))
    End With
    Call ReadListData(strInvoiceNO)
End Sub
 
 

Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "�����б�", False
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strNo As String
    If OldRow = NewRow Then Exit Sub
    With vsList
        If NewRow < 0 Then Exit Sub
        strNo = Trim(.TextMatrix(NewRow, .ColIndex("���ݺ�")))
    End With
    ShowDetail strNo
    
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "�����б�", False
End Sub
 
Private Sub vsList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsList
        Select Case Col
        Case .ColIndex("ѡ��")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True, _
    Optional ByVal intInsure_IN As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���³�ʼ�����˷�Ʊ��Ϣ
    '���:blnFact-�Ƿ�����ȡ��Ʊ��
    '����:���˺�
    '����:2011-04-29 14:17:33
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String, lng����ID As Long
    Dim intInsure As Integer
  
    If Not mrsInfo Is Nothing Then
      If mrsInfo.State = 1 Then lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    intInsure = IIf(intInsure_IN <> 0, intInsure_IN, mintInsure)
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModule, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModule, mstrUseType, mintOldInvoiceFormat)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModule, mstrUseType)
    Call ZlShowBillFormat(mlngModule, lblFormat, mintInvoiceFormat)
    If blnFact Then Call RefreshFact
End Sub

Private Sub RefreshFact()
    '���ܣ�ˢ���շ�Ʊ�ݺ�
  '  If mintInvoicePrint = 0 Then Exit Sub
    If gblnStrictCtrl Then
        'lblFact.tag��Ҫ�Ǽ�鷢Ʊ���Ƿ��ֹ������.�ֹ������,��Ʊ��Ϊ��,�������Զ������ķ�Ʊ��
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            If zlCheckInvoiceValied(mlng����ID, 1, , mlngShareUseID, mstrUseType) = False Then
                txtInvoice.Text = "": txtInvoice.Tag = "": Exit Sub
            End If
            '�ϸ�ȡ��һ������
            txtInvoice.Text = GetNextBill(mlng����ID)
            'Tag�����⣺24363:���˺飺��Ҫ�ǽ���Զ����ɵĺ��Ƿ��û����ģ���Ҫ�����
            '    1.���ĵ�Ʊ�ݺ���Ҫ����Ƿ��ظ����ظ���ֱ�ӷ��ز����ķ�Ʊ��
            '    2.���������������ĵ�����£�����Ƿ��ظ�������ظ����Զ�ȡ��һ�����룡
            txtInvoice.Tag = txtInvoice.Text
            lblFact.Tag = txtInvoice.Tag
            If mblnStartFactUseType Then Call zlCheckFactIsEnough
        End If
    Else
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            '��ɢ��ȡ��һ������
            txtInvoice.Text = zlStr.Increase(UCase(zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, mlngModule)))
            'Tag�����⣺24363:���˺飺��Ҫ�ǽ���Զ����ɵĺ��Ƿ��û����ģ���Ҫ�����
            '    1.���ĵ�Ʊ�ݺ���Ҫ����Ƿ��ظ����ظ���ֱ�ӷ��ز����ķ�Ʊ��
            '    2.���������������ĵ�����£�����Ƿ��ظ�������ظ����Զ�ȡ��һ�����룡
        End If
        txtInvoice.Tag = txtInvoice.Text
        lblFact.Tag = txtInvoice.Tag
    End If
    txtInvoice.SelStart = Len(txtInvoice.Text)
End Sub

Private Sub zlCheckFactIsEnough()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰƱ���Ƿ�����
    '����:���˺�
    '����:2011-05-10 17:54:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngʣ������ As Long
    '���˺� ����:26948 ����:2009-12-28 17:43:00
    '��Ҫ���ʣ�������Ƿ����:
    If zlCheckInvoiceOverplusEnough(1, gTy_Module_Para.int����ʣ��Ʊ������, lngʣ������, mlng����ID, mstrUseType) = False Then
        MsgBox "ע��:" & vbCrLf & _
               "    ��ǰʣ��Ʊ��(" & lngʣ������ & ") С���˱���������(" & gTy_Module_Para.int����ʣ��Ʊ������ & "),��ע�������Ʊ!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
    End If
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



