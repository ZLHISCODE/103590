VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmMakeupPrintBill 
   Caption         =   "סԺ���˲���Ʊ"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   Icon            =   "frmMakeupPrintBill.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   11790
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBalance 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   6480
      ScaleHeight     =   4695
      ScaleWidth      =   4380
      TabIndex        =   17
      Top             =   1140
      Width           =   4380
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   2685
         Left            =   525
         TabIndex        =   18
         Top             =   300
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
         FormatString    =   $"frmMakeupPrintBill.frx":030A
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
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         Caption         =   "����ϼ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   465
         TabIndex        =   19
         Top             =   3165
         Width           =   1155
      End
   End
   Begin VB.PictureBox PicDetail 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   75
      ScaleHeight     =   2775
      ScaleWidth      =   5535
      TabIndex        =   15
      Top             =   4290
      Width           =   5535
      Begin VSFlex8Ctl.VSFlexGrid vsDetail 
         Height          =   2685
         Left            =   180
         TabIndex        =   16
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
         FormatString    =   $"frmMakeupPrintBill.frx":0384
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
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   90
      ScaleHeight     =   2295
      ScaleWidth      =   5400
      TabIndex        =   13
      Top             =   1290
      Width           =   5400
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2685
         Left            =   105
         TabIndex        =   14
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
         FormatString    =   $"frmMakeupPrintBill.frx":039A
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
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   375
         Left            =   615
         TabIndex        =   24
         Top             =   150
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   661
         Appearance      =   2
         IDKindStr       =   "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;ס|סԺ��|0;��|���￨|0"
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
         MaxLength       =   100
         TabIndex        =   9
         Top             =   150
         Width           =   2040
      End
      Begin VB.CommandButton cmdBrush 
         Caption         =   "ˢ��(&N)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9195
         TabIndex        =   6
         Top             =   150
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   345
         Left            =   5520
         TabIndex        =   7
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
         Format          =   93323267
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
         Format          =   93323267
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
         TabIndex        =   8
         Top             =   210
         Value           =   1  'Checked
         Width           =   1890
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
         Left            =   150
         TabIndex        =   12
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblArang 
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
         TabIndex        =   11
         Top             =   225
         Width           =   240
      End
   End
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   45
      ScaleHeight     =   660
      ScaleWidth      =   11700
      TabIndex        =   0
      Top             =   7170
      Width           =   11700
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
         Left            =   2985
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "����(&O)"
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
         Left            =   8955
         TabIndex        =   4
         Top             =   225
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "ȫѡ(&A)"
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
         Left            =   105
         TabIndex        =   3
         ToolTipText     =   "�ȼ���Ctrl+A"
         Top             =   225
         Width           =   1100
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "ȫ��(&R)"
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
         Left            =   1245
         TabIndex        =   2
         ToolTipText     =   "�ȼ���Ctrl+R"
         Top             =   225
         Width           =   1100
      End
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
         TabIndex        =   1
         Top             =   225
         Width           =   1100
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   5325
         TabIndex        =   23
         Top             =   300
         Visible         =   0   'False
         Width           =   120
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
         Left            =   2490
         TabIndex        =   22
         Top             =   300
         Width           =   480
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   8055
      Width           =   11790
      _ExtentX        =   20796
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
            Picture         =   "frmMakeupPrintBill.frx":03B0
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15716
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
            Picture         =   "frmMakeupPrintBill.frx":0C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeupPrintBill.frx":0F98
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
Attribute VB_Name = "frmMakeupPrintBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mPanel
    Pane_Search = 1
    Pane_List = 2
    Pane_Detail = 3
    Pane_Balance = 4
End Enum
'-----------------------------------------------------------------------------------
'���㿨���
Private mSquareCard As SquareCard '���㿨���
Private mstrPassWord As String
Private mbytInvoiceKind As Byte
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
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mblnValid As Boolean
Private mblnSel As Boolean
Private mstrPrivs As String
Private mintSucces As Integer  '�ɹ���ӡ����
Private mlng����ID As Long
Private mblnFirst As Boolean
Private mblnNotClick As Boolean
Private mobjInvoice As clsInvoice
Private mobjFact As clsFactProperty

Private mlng����ID As Long
Private mintInsure As Integer
Private mintPrintNums As Boolean '��ӡƱ������
Private mblnNOMoved As Boolean

Public Function zlRePrintBill(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal strPrivs As String, Optional lng����ID As Long = 0) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ش�Ʊ�����
    '����:��ӡ�ɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2013-01-05 15:21:03
    '����:56283
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mlngModule = lngModule: mstrPrivs = strPrivs: mintSucces = 0
    mlng����ID = lng����ID
    mbytInvoiceKind = IIf(gbytInvoiceKind = 0, 3, 1)
    
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlRePrintBill = mintSucces > 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function InitPanel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2013-01-05 15:22:02
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
        Set objTemp = .CreatePane(mPanel.Pane_List, 300, 100, DockBottomOf, objPane)
        objTemp.Tag = mPanel.Pane_List
        objTemp.Title = "���˵����б�": objTemp.Options = PaneNoCloseable Or PaneNoHideable
        objTemp.Handle = picList.hWnd
        
        Set objPane = .CreatePane(mPanel.Pane_Balance, 100, 100, DockRightOf, objTemp)
        objPane.Tag = mPanel.Pane_Balance
        objPane.Title = "������Ϣ�б�": objPane.Options = PaneNoCloseable Or PaneNoHideable
        objPane.Handle = picBalance.hWnd
       '
        Set objPane = .CreatePane(mPanel.Pane_Detail, 300, 100, DockBottomOf, objTemp)
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
End Function
Private Sub chkDate_Click()
    dtpBegin.Enabled = chkDate.Value <> 1
    dtpEnd.Enabled = chkDate.Value <> 1
End Sub
Private Sub cmdBrush_Click()
        Call ReadListData
End Sub
Private Sub cmdCancel_Click()
     Unload Me
End Sub

Private Sub cmdClear_Click()
    On Error GoTo errHandle
    With vsList
        If .ColIndex("ѡ��") < 0 Then Exit Sub
        .Cell(flexcpChecked, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 0
    End With
    Call SetBlanceShow
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function zlMakeupPrint(ByVal lng����ID As Long, lng��ҳID As Long, _
    ByVal strNO As String, ByVal lng����ID As String, ByVal intInsure As Integer, _
    Optional strInvoice As String = "", Optional ByVal bytFunc As Byte = 1) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ʊ��
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-01-05 15:24:07
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFact As New clsFactProperty
    
    On Error GoTo errHandle
    If lng����ID <= 0 Or lng����ID = 0 Then Exit Function
 
    mobjFact.LastUseID = mlng����ID
    Call frmPrint.ReportPrint(2, strNO, lng����ID, mobjFact, strInvoice, , , , lng����ID, mobjFact.��ӡ��ʽ)
        
    '��ҽһ��ͨд����85950
    Call WriteInforToCard(Me, mlngModule, mstrPrivs, gobjSquare.objSquareCard, 0, bytFunc, lng����ID, lng����ID)
    
    mintSucces = mintSucces + 1
    zlMakeupPrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckFP(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鷢Ʊ�Ƿ���ȷ
    '����: ��Ʊ�Ϸ� ����true,���򷵻�False
    '����:���˺�
    '����:2012-07-12 11:30:22
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intNum As Integer
    
     On Error GoTo errHandle
    intNum = 1
    If lng����ID = 0 Then
        MsgBox "��������Ҫ�����Ʊ��", vbInformation, gstrSysName
        Exit Function
    End If
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
    
InvoiceHandle:
    If zlGetInvoiceGroupUseID(mlng����ID, intNum, txtInvoice.Text) = False Then
        Exit Function
    End If
    '�����������,Ʊ���Ƿ�����
    If CheckBillRepeat(mlng����ID, mbytInvoiceKind, txtInvoice.Text) Then
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

     Dim str����IDs As String, lng����ID As Long, strNO As String
     Dim lngRow As Long, intInsure As Integer
     Dim lng����ID As Long, lng��ҳID As Long
     Dim bytFunc As Byte '��������
     
    On Error GoTo errHandle

     If mrsInfo Is Nothing Then Exit Sub
     If mrsInfo.State <> 1 Then Exit Sub
     If mrsInfo.RecordCount = 0 Then Exit Sub
     
     lng����ID = Val(Nvl(mrsInfo!����ID))
     lng��ҳID = Val(Nvl(mrsInfo!��ҳID))
     
     With vsList
        str����IDs = ""
        For lngRow = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsList, lngRow, .ColIndex("ѡ��")) Then
                 lng����ID = Val(.TextMatrix(lngRow, .ColIndex("����ID")))
                 strNO = Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�")))
                 intInsure = Val(.TextMatrix(lngRow, .ColIndex("����ID")))
                 bytFunc = IIf(.TextMatrix(lngRow, .ColIndex("��������")) = "�������", 0, 1)
                 If lng����ID <> 0 Then
                    If Not CheckFP(lng����ID) Then Exit Sub
                    
                     str����IDs = str����IDs & "," & lng����ID
                    If lng����ID <> Val(.TextMatrix(lngRow, .ColIndex("����ID"))) _
                        Or Val(.TextMatrix(lngRow, .ColIndex("��ҳID"))) <> lng��ҳID _
                        Or intInsure <> mintInsure Then
                        
                        lng����ID = Val(.TextMatrix(lngRow, .ColIndex("����ID"))): lng��ҳID = Val(.TextMatrix(lngRow, .ColIndex("��ҳID")))
                        mintInsure = intInsure
                        
                        '����Ʊ�����Ϣ
                        Call ReInitPatiInvoice(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID)), intInsure)
                        Call RefreshFact    '���´���Ʊ
                    End If
                    Call zlMakeupPrint(lng����ID, lng��ҳID, strNO, lng����ID, intInsure, Trim(txtInvoice.Text), bytFunc)
                    '����ȡ��Ʊ
                    Call RefreshFact
                 End If
            End If
        Next
     End With
     If str����IDs = "" Then
        MsgBox "δѡ��Ҫ�����Ʊ��", vbOKOnly, gstrSysName
        Exit Sub
     End If
     Call zlClearPatiInfor
     If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSelAll_Click()
    On Error GoTo errHandle
    With vsList
        If .ColIndex("ѡ��") < 0 Then Exit Sub
        .Cell(flexcpChecked, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = -1
    End With
    Call SetBlanceShow
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pane_Search    '1
        Item.Handle = picCon.hWnd
    Case Pane_List      ' 2
        Item.Handle = picList.hWnd
    Case Pane_Detail    '3
        Item.Handle = PicDetail.hWnd
    Case Pane_Balance  ' 4
        Item.Handle = picBalance.hWnd
    End Select
End Sub
Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
     Bottom = stbThis.Height + picDown.Height
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    On Error GoTo errHandle
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
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub Form_Load()

    On Error GoTo errHandle
    cmdBrush.Enabled = False
    mblnFirst = True
    lblFormat.Alignment = 0
    dtpBegin.MaxDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59")
    dtpEnd.MaxDate = dtpBegin.MaxDate
    dtpBegin.Value = Format(DateAdd("d", -7, dtpBegin.MaxDate), "yyyy-mm-dd")
    dtpEnd.Value = Format(dtpEnd.MaxDate, "yyyy-mm-dd")
    txtInvoice.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ�ݺ�") And gblnStrictCtrl '89302
    Call InitPanel
    Call zlCardSquareObject
    Call zlCreateObject
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 


Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With picDown
        .Top = ScaleHeight - stbThis.Height - .Height
        .Width = ScaleWidth
        .Left = ScaleLeft
    End With
End Sub
Private Sub UnloadWinSaveInfor()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�رմ���ʱ,���������Ϣ
    '����:���˺�
    '����:2013-01-05 15:27:20
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Call zlCardSquareObject(True)
    zlSaveDockPanceToReg Me, dkpMan, "����"
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Caption, "�����б�", False
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "�����б�", False
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "������ϸ�б�", False
    Call zlCloseObject
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    '���洰����Ի���Ϣ
    Call UnloadWinSaveInfor
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    On Error GoTo errHandle
    
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
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
Private Sub picDetail_Resize()
    Err = 0: On Error Resume Next
    With PicDetail
        vsDetail.Left = .ScaleLeft
        vsDetail.Width = .ScaleWidth
        vsDetail.Height = .ScaleHeight
        vsDetail.Top = .ScaleTop
    End With
End Sub

Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBalance.Left = .ScaleLeft
        vsBalance.Width = .ScaleWidth
        vsBalance.Height = .ScaleHeight - lblSum.Height - 50
        vsBalance.Top = .ScaleTop
        lblSum.Top = .ScaleHeight - lblSum.Height - 10
        lblSum.Left = .ScaleLeft
    End With
End Sub
 Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
                            
    On Error GoTo errHandle
    
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("���֤", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    
    On Error GoTo errHandle
    
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
    '��ʼ�������ɹ�,����Ϊ�����ڴ���
   If mSquareCard.objSquareCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then Exit Sub
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��
    '����:blnOutMsg-�Ѿ���ʾ,�������ⲿ����ʾ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-01-05 15:30:46
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim strIF As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    mstrFindNO = "": mstrFindFpNo = ""
    mintPrintNums = 0
    strSQL = _
        "Select A.����ID,Nvl(B.��ҳID,0) as ��ҳID,A.����� as �����,A.��ǰ����,B.��Ժ����," & _
        "       Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�,Nvl(B.����,A.����) as ����,A.IC����,A.���￨��,A.����֤��," & _
        "       Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�,C.���� as ��ǰ����,A.��ǰ����ID,D.���� as ��Ժ����,B.��Ժ����ID, A.���� as ����,E.����,E.ҽ����,E.����," & _
        "       A.�Ǽ�ʱ��,Nvl(B.״̬,0) as ״̬,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,Nvl(B.��˱�־,0) as ��˱�־,B.��Ժ����,B.��Ժ����,B.��������,B.��������" & _
        " From ������Ϣ A,������ҳ B,���ű� C,���ű� D,ҽ�����˵��� E,ҽ�����˹����� F" & _
        " Where A.ͣ��ʱ�� is NULL And A.����ID=B.����ID(+) And A.��ҳID=B.��ҳID(+) " & _
        "           And A.����ID=F.����ID(+) And F.��־(+)=1 And F.ҽ����=E.ҽ����(+) And F.����=E.����(+) And F.���� = E.����(+)" & _
        "           And A.��ǰ����ID=C.ID(+) And B.��Ժ����ID=D.ID(+)" & _
        "           And A.ͣ��ʱ�� is NULL "
    
    If blnCard = True And objCard.���� Like "����*" Then    'ˢ��
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
        strSQL = strSQL & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSQL = strSQL & " And A.�����=[1]"
    ElseIf Left(strInput, 1) = "." Or objCard.���� = "���ݺ�" Then
        '���ݺŲ���
        If Left(strInput, 1) = "." Then
            strTemp = UCase(GetFullNO(Mid(strInput, 2), 15))
        Else
            strTemp = UCase(GetFullNO(strInput, 15))
        End If
        
        txtPatient.Text = strTemp
        gstrSQL = "" & _
        "   Select  distinct A.����ID " & _
        "   From ���˽��ʼ�¼A " & _
        "   Where A.NO=[1]  And Rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp, 1)
        If rsTemp.EOF Then
            MsgBox "ע��:" & vbCrLf & "  ���ݺ�Ϊ��" & strTemp & "��������,��������ĵ����Ƿ���ȷ!", vbInformation + vbOKOnly, gstrSysName
            Call zlClearPatiInfor
            Exit Function
        End If
        If Val(Nvl(rsTemp!����ID)) = 0 Then
            MsgBox "�ý��˵����Ǻ�Լ��λ����!", vbInformation, gstrSysName
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
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
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
        txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!��������))
        txtPatient.Text = Nvl(mrsInfo!����)
        'txtOld.Text = Nvl(mrsInfo!����): txtSex.Text = Nvl(mrsInfo!�Ա�)
        ' txtסԺ��.Text = Nvl(mrsInfo!�����)
        txtPatient.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
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
    '����:2013-01-05 15:31:23
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    txtPatient.Text = ""
    Set mrsInfo = New ADODB.Recordset
    vsList.Clear 1: vsList.Rows = 2: vsDetail.Clear 1: vsDetail.Rows = 2
    vsBalance.Clear 1: vsBalance.Rows = 2
    txtInvoice.Text = ""
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txtPatient_Change()
    txtPatient.Tag = ""
    cmdBrush.Enabled = False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    IDKind.SetAutoReadCard (txtPatient.Text <> "" And Me.ActiveControl Is txtPatient)
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
                Call zlClearPatiInfor
                Exit Sub
            End If
            Call ReadListData
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
    '����:2013-01-05 15:32:23
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim strSQL As String, curTotal As Currency, blnIDCard As Boolean
    Dim blnICCard As Boolean, blnMsg As Boolean
    
    On Error GoTo errHandle
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ And mstrPassWord <> "" Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.���� Like "*���֤*" And objCard.ϵͳ And mstrPassWord <> "" Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If Not GetPatient(objCard, strInput, blnCard, blnMsg) Then
        cmdBrush.Enabled = False
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
    If Mid(gstrCardPass, 6, 1) = "1" And (blnCard Or blnICCard Or blnIDCard) Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
            Call zlClearPatiInfor
             txtPatient.SetFocus: Exit Sub
        End If
    End If
    cmdBrush.Enabled = True
    Call ReadListData
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
  
Private Sub ShowDetail(ByVal strNO As String, Optional lng����ID As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ϸ����
    '���:strNO-���˵��ݺ�
    '       lng����ID-����ID
    '����:���˺�
    '����:2013-01-05 15:33:04
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim i As Long, j As Long, strSQL As String
    Dim lngCol As Long
    
    On Error GoTo errH
    
    '�Ƿ���ת������ݱ���
    mblnNOMoved = zlDatabase.NOMoved("���˽��ʼ�¼", strNO, , , Me.Caption)
    strSQL = "" & _
    "   Select  '����' as סԺ,A.����ʱ��,A.NO,A.���,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.Ӥ����,A.���ʽ��,A.��������ID " & _
    "   From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A" & _
    "   Where A.����ID=[1]" & _
    "    Union ALL " & _
    "   Select  Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��') as סԺ,A.����ʱ��,A.NO,A.���,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.Ӥ����,A.���ʽ��,A.��������ID " & _
    "   From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A" & _
    "   Where A.����ID=[1] " & _
    "   "
    strSQL = _
    "  Select   A.סԺ," & _
    "            Nvl(B.����,'δ֪') as ����,To_Char(A.����ʱ��,'YYYY-MM-DD') as ʱ��," & _
    "            A.NO as ���ݺ�,Nvl(E.����,D.����) as ��Ŀ,A.�վݷ�Ŀ as ��Ŀ," & _
    "            Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.���ʽ��" & _
    " From (" & strSQL & ") A,���ű� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E" & _
    " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=D.ID" & _
    "           And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    " Order by סԺ Desc,ʱ�� Desc,���ݺ� Desc,���"
    Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If mrsDetail.EOF Then Exit Sub
 
    With vsDetail
        .Clear 1
        .Redraw = flexRDNone
        Set .DataSource = mrsDetail
        For lngCol = 0 To .Cols - 1
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
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsDetail, Me.Caption, "������ϸ�б�", False
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function InitBlanceData(ByVal str����IDs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '���:str����IDs-����ID(����ö��ŷ���)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-01-05 15:33:40
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim strSQL As String
    Err = 0: On Error GoTo errHandle
    
    If str����IDs = "" Then InitBlanceData = True: Exit Function
    
    strSQL = _
    " Select   M.NO,M.ID as ����ID" & _
    " From ���˽��ʼ�¼ M , Table(f_num2list([1]))  J" & _
    " Where  M.ID=J.Column_Value"
    
    strSQL = _
    " Select /*+ rule */ A.NO,A.����ID, A.���㷽ʽ,Nvl(B.����,1) as ����,B.Ӧ����,A.���,A.ժҪ,A.�������" & _
    " From (  Select B.NO ,A.����ID,Decode(A.��¼����,2,A.���㷽ʽ,12,A.���㷽ʽ,NULL) as ���㷽ʽ,A.ժҪ,A.�������," & _
    "               Sum(A.��Ԥ��) as ���" & _
    "         From ����Ԥ����¼ A, (" & strSQL & ")  B" & _
    "         Where A.����ID=B.����ID And A.��¼���� IN(1,11,2,12) And Nvl(A.��Ԥ��,0)<>0" & _
    "         Group by B.NO,A.����ID, Decode(A.��¼����,2,A.���㷽ʽ,12,A.���㷽ʽ,NULL),A.ժҪ,A.�������" & _
    "       ) A,���㷽ʽ B " & _
    " Where A.���㷽ʽ=B.����(+) " & _
    " "
    
    Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(str����IDs, "'", ""))
    InitBlanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetBlanceShow()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ���㷽ʽ
    '���:blnAllSel-ѡ�����еĵ���
    '����:���˺�
    '����:2013-01-05 15:35:36
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String, lngRow As Long, i As Long, str���� As String
    Dim blnȫѡ As Boolean, blnδѡ As Boolean
    Dim strFilter As String, bln�˿� As Boolean
    Dim lng����ID As Long, str����IDs As String
    Dim dblMoney As Double
    
    On Error GoTo errHandle
    
    lblSum.Caption = "����ϼ�:" & Format(0, "0.00")
    If mrsBalance Is Nothing Then Exit Sub
    
    With vsList
        blnȫѡ = True: blnδѡ = True
        For lngRow = 1 To .Rows - 1
            lng����ID = Val(.TextMatrix(lngRow, .ColIndex("����ID")))
            
            If GetVsGridBoolColVal(vsList, lngRow, .ColIndex("ѡ��")) Then
                If InStr(1, str����IDs & ",", "," & lng����ID & ",") = 0 Then
                    str����IDs = str����IDs & "," & lng����ID
                    blnδѡ = False
                End If
            End If
             If InStr(1, str����IDs & ",", "," & lng����ID & ",") = 0 Then blnȫѡ = False
        Next
    End With
    If str����IDs <> "" Then str����IDs = Mid(str����IDs, 2)
    bln�˿� = False
    '��ʾ����ѡ��ĵ��ݵĽ��㷽ʽ֮��
    If blnȫѡ Or blnδѡ Then
        mrsBalance.Filter = 0
        If blnȫѡ Then bln�˿� = True
    Else
        strFilter = Replace(str����IDs, ",", " Or ����ID=")
        strFilter = " ����ID=" & strFilter & ""
        mrsBalance.Filter = strFilter
        bln�˿� = True
    End If
    
    mrsBalance.Sort = "NO,���㷽ʽ"
    With vsBalance
         .Redraw = flexRDNone
        .Rows = IIf(mrsBalance.RecordCount = 0, 1, mrsBalance.RecordCount) + 1
        i = 1
        dblMoney = 0
        Do While Not mrsBalance.EOF
            .TextMatrix(i, .ColIndex("NO")) = Nvl(mrsBalance!NO)
            .TextMatrix(i, .ColIndex("���㷽ʽ")) = Nvl(mrsBalance!���㷽ʽ)
            .TextMatrix(i, .ColIndex("������")) = Format(Val(Nvl(mrsBalance!���)), "0.00")
            dblMoney = dblMoney + Val(Nvl(mrsBalance!���))
            i = i + 1
            mrsBalance.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Caption, "�����б�", False
        .Redraw = flexRDBuffered
        If blnδѡ Then
            lblSum.Caption = "δ��ϼ�:" & Format(dblMoney, "0.00")
        Else
            lblSum.Caption = "����ϼ�:" & Format(dblMoney, "0.00")
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function ReadListData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ҫ��ϸ����
    '����:��ȡ�ɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2013-01-05 15:36:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim lngCol As Long, strSQL As String, lngRow As Long
    Dim strFilter As String, str����IDs As String
    Dim strWhere As String, strTable1 As String, dtStartDate As Date, dtEndDate As Date
    Dim strNO As String, i As Long, lng������� As Long
    Dim lng��ҳID  As Long
    
    On Error GoTo errHandle
    
    dtStartDate = CDate("1901-01-01")
    dtEndDate = dtStartDate
    lng��ҳID = 0
    If mrsInfo Is Nothing Then
        lng����ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = 0
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
        lng��ҳID = Val(Nvl(mrsInfo!��ҳID))
    End If
    If mstrFindNO <> "" Then
        strWhere = "  And A.NO=[2]"
    Else
        strTable1 = ""
        strWhere = "  And A.����ID=[1]"
    End If
    
    If chkDate.Value = 0 Then
        strWhere = strWhere & " And A.�շ�ʱ�� betWeen [3] and [4]"
        dtStartDate = CDate(Format(dtpBegin.Value, "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpEnd.Value, "yyyy-mm-dd") & " 23:59:59")
    End If
    mblnSel = False
    zlCommFun.ShowFlash "���ڶ�ȡ��������,���Ժ� ..."
    Screen.MousePointer = 11
    DoEvents
    strSQL = "" & _
    "   Select  -1 as  ѡ�� ,  max(Decode(Y.����,NULL,NULL,'��')) as ҽ��,a.Id as ����ID, a.No as ���ݺ�, a.ʵ��Ʊ��, a.����id, Max(B.��ҳID) as ��ҳID, a.����Ա���, a.����Ա����, a.��ע, a.ԭ��, To_Char(a.�շ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, " & _
    "          Decode(a.��������, 1, '�������', 'סԺ����') As ��������, To_Char(Sum(b.ʵ�ս��), '99999990.00') As ʵ�ʽ��, " & _
    "          nvl(Max(X.����),0) as ����ID " & _
    "   From ���˽��ʼ�¼ A, סԺ���ü�¼ B,���ս����¼ X,������� Y " & _
    "   Where   a.Id = b.����id And a.ʵ��Ʊ�� Is Null And a.��¼״̬ = 1 " & strWhere & _
    "         And A.id=X.��¼ID(+) And X.����(+)=2 And X.����=Y.���(+) And Nvl(X.���(+),1)=1 " & _
    "   Group By a.Id, a.No, a.ʵ��Ʊ��, a.����id , a.����Ա���, a.����Ա����, a.��ע, a.ԭ��, " & _
    "           To_Char(a.�շ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'),  Decode(a.��������, 1, '�������', 'סԺ����') " & _
    "   order by ����ʱ��,���ݺ�"
     
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, mstrFindNO, dtStartDate, dtEndDate)
    vsList.Redraw = flexRDNone
    vsList.Clear: vsList.Cols = 2
    Set vsList.DataSource = mrsList
    If vsList.Rows <= 1 Then vsList.Rows = 2
    With vsList
        mstrNOs = ""
        For lngCol = 0 To .Cols - 1
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
        .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "�����б�", False
        .Editable = flexEDKbdMouse

        .Redraw = flexRDBuffered
        vsList_AfterRowColChange 0, 0, .Row, .Col
    
    End With
    mrsList.Filter = "����ID>0"
    If Not mrsList.EOF Then
        mintInsure = Val(Nvl(mrsList!����ID))
    End If
    mrsList.Filter = 0
    If mrsList.RecordCount <> 0 Then mrsList.MoveFirst
    str����IDs = ""
    With mrsList
        Do While Not .EOF
            str����IDs = str����IDs & "," & Val(Nvl(!����ID))
            .MoveNext
        Loop
        str����IDs = "-1" & str����IDs
        If str����IDs <> "" Then str����IDs = Mid(str����IDs, 2)
    End With
    '���ؽ��㷽ʽ
    Call InitBlanceData(str����IDs)
    Call SetBlanceShow
    '����Ʊ�����Ϣ
    Call ReInitPatiInvoice(lng����ID, lng��ҳID, mintInsure)
    Call RefreshFact    '���´���Ʊ
    zlCommFun.StopFlash
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
Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Caption, "�����б�", False
End Sub

Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Caption, "�����б�", False
End Sub

Private Sub vsDetail_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "������ϸ�б�", False
End Sub

Private Sub vsDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "������ϸ�б�", False
End Sub
Private Sub vsList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call SetBlanceShow
End Sub

Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "�����б�", False
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strNO As String, lng����ID As Long
    
    On Error GoTo errHandle
    
    If OldRow = NewRow Then Exit Sub
    With vsList
        If NewRow < 0 Then Exit Sub
        strNO = Trim(.TextMatrix(NewRow, .ColIndex("���ݺ�")))
        lng����ID = Val(.TextMatrix(NewRow, .ColIndex("����ID")))
    End With
    ShowDetail strNO, lng����ID
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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

Private Sub ReInitPatiInvoice(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
     Optional ByVal intInsure As Integer = 0)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���³�ʼ�����˷�Ʊ��Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-01-05 15:38:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytInvoiceKind     As Byte, intFormat As Integer
    Dim intPrintMode As Integer, lngShareUseID As Long
    On Error GoTo errHandle

    bytInvoiceKind = Val(zlDatabase.GetPara("סԺ����Ʊ������", glngSys, 1137, "0"))

    Set mobjInvoice = New clsInvoice: Set mobjFact = New clsFactProperty
    mobjInvoice.zlInitCommon glngSys, gcnOracle, gstrDBUser
    
    Call mobjInvoice.zlGetInvoicePreperty(1137, IIf(bytInvoiceKind = 0, 3, 1), 0, 0, 0, mobjFact, , , 2)
    
    mobjFact.ʹ����� = zlDatabase.GetPara("��Լ��λ���ʴ�ӡ", glngSys, 1137)
    mobjFact.Ʊ�� = IIf(bytInvoiceKind = 0, 3, 1)
    Call mobjInvoice.zlGetInvoicePrintFormat(1137, mobjFact.Ʊ��, mobjFact.ʹ�����, intFormat, 2)
    mobjFact.��ӡ��ʽ = intFormat
    If mobjInvoice.zlGetInvoicePrintMode(1137, mobjFact.Ʊ��, mobjFact.ʹ�����, intPrintMode) = False Then Exit Sub
    mobjFact.��ӡ��ʽ = intPrintMode
    If mobjInvoice.zlGetInvoiceShareID(1137, mobjFact.Ʊ��, mobjFact.ʹ�����, lngShareUseID) = False Then Exit Sub
    mobjFact.��������ID = lngShareUseID
    
    Call ZlShowBillFormat(bytInvoiceKind, lblFormat, mobjFact.��ӡ��ʽ)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshFact()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ���շ�Ʊ�ݺ�
    '����:���˺�
    '����:2013-01-05 15:39:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
       Dim strFactNO As String
       
    On Error GoTo errHandle
        
    If mobjFact.��ӡ��ʽ = 0 Then Exit Sub
    If Not mobjFact.�ϸ���� Then
        '���ϸ������
        '��ɢ��ȡ��һ������
        txtInvoice.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("��ǰ����Ʊ�ݺ�", glngSys, 1137, "")))
        txtInvoice.Tag = txtInvoice.Text
        txtInvoice.SelStart = Len(txtInvoice.Text)
        Exit Sub
    End If
    If zlGetInvoiceGroupUseID(mlng����ID, 1, "") = False Then
          txtInvoice.Text = "": txtInvoice.Tag = ""
        Exit Sub
    End If
    
    '�ϸ�ȡ��һ������
    If mobjInvoice.zlGetNextBill(1137, mlng����ID, strFactNO) = False Then strFactNO = ""
    txtInvoice.Text = strFactNO
    'Tag�����⣺24363:���˺飺��Ҫ�ǽ���Զ����ɵĺ��Ƿ��û����ģ���Ҫ�����
    '    1.���ĵ�Ʊ�ݺ���Ҫ����Ƿ��ظ����ظ���ֱ�ӷ��ز����ķ�Ʊ��
    '    2.���������������ĵ�����£�����Ƿ��ظ�������ظ����Զ�ȡ��һ�����룡
    txtInvoice.Tag = txtInvoice.Text
    lblFact.Tag = txtInvoice.Tag
    txtInvoice.SelStart = Len(txtInvoice.Text)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
     
End Sub

Private Function zlGetInvoiceGroupUseID(ByRef lng����ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵ�����ID
    '���:lng����ID-����id
    '       intNum-ҳ��
    '       strInvoiceNO-����ķ�Ʊ��
    '����:lng����ID-����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    If mobjInvoice.zlGetInvoiceGroupID(1137, UserInfo.����, mobjFact.Ʊ��, _
        mobjFact.ʹ�����, lng����ID, mobjFact.��������ID, lng����ID, intNum, strInvoiceNO) = False Then Exit Function
    
    If lng����ID > 0 Then zlGetInvoiceGroupUseID = True: Exit Function
    
    Select Case lng����ID
        Case 0 '����ʧ��
        Case -1
            If Trim(mobjFact.ʹ�����) = "" Then
                MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "��û�����ú͹��õġ�" & mobjFact.ʹ����� & "������Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
            Exit Function
        Case -2
            If Trim(mobjFact.ʹ�����) = "" Then
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "���صĹ���Ʊ�ݵġ�" & mobjFact.ʹ����� & "������Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
            Exit Function
        Case -3
            MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,���������룡", vbInformation, gstrSysName
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus
            Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub zlCreateObject()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������¼�����
    '����:���˺�
    '����:2013-01-05 15:41:19
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
Private Function CheckBillRepeat(lng����ID As Long, bytƱ�� As Byte, strNO As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʹ����Ʊ��֮ǰ������Ƿ��ظ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-01-05 15:42:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���� From Ʊ��ʹ����ϸ" & _
        " Where ����ID=[1] And Ʊ��=[2] And ����=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng����ID, bytƱ��, strNO)
    CheckBillRepeat = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

