VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMakeupPrintBill 
   Caption         =   "�����շѲ���Ʊ"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   Icon            =   "frmMakeupPrintBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   11790
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBalance 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   6525
      ScaleHeight     =   4695
      ScaleWidth      =   4380
      TabIndex        =   17
      Top             =   1185
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
         Cols            =   5
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
      Begin VB.Label lbl�ϼ� 
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
         FormatString    =   $"frmMakeupPrintBill.frx":03CF
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
         FocusRect       =   0
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
         FormatString    =   $"frmMakeupPrintBill.frx":03E5
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
      Begin VB.CheckBox chkRegistFee 
         Caption         =   "���Һŷ�"
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
         Left            =   8850
         TabIndex        =   25
         Top             =   210
         Width           =   1320
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   375
         Left            =   555
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
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
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
         Left            =   1245
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
         Left            =   10245
         TabIndex        =   6
         Top             =   150
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   345
         Left            =   5175
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
         Format          =   146800643
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   345
         Left            =   7125
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
         Left            =   3420
         TabIndex        =   8
         Top             =   210
         Value           =   1  'Checked
         Width           =   1800
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
         Left            =   90
         TabIndex        =   12
         Top             =   210
         Width           =   480
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
         Left            =   6840
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
         Left            =   3780
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
         Left            =   8565
         TabIndex        =   4
         Top             =   210
         Width           =   1440
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
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "�ȼ���Ctrl+A"
         Top             =   225
         Width           =   1440
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
         Left            =   1530
         TabIndex        =   2
         ToolTipText     =   "�ȼ���Ctrl+R"
         Top             =   225
         Width           =   1440
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
         Width           =   1440
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   6105
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
         Left            =   3285
         TabIndex        =   22
         Top             =   330
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
            Picture         =   "frmMakeupPrintBill.frx":03FB
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
            Picture         =   "frmMakeupPrintBill.frx":0C8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeupPrintBill.frx":0FE3
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
'-----------------------------------------------------------------------------------
Private mrsInfo As ADODB.Recordset
Private mrsList As ADODB.Recordset  '�����б�
Private mrsDetail As ADODB.Recordset
Private mrsBalance As ADODB.Recordset
Private mlngModule As Long
Attribute mlngModule.VB_VarHelpID = -1
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
Private mblnStartFactUseType As Boolean   '�Ƿ�������ʹ�����
Private mintInvoicePrint As Integer  '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
Private mlng����ID As Long
Private mintInsure As Integer

'��ز���
Private mbln���ֽ������  As Boolean
Private mintPatiInvoiceFormat As Integer '���ֽ��������ӡ�ķ�Ʊ��ʽ

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
    
    mbln���ֽ������ = Val(zlDatabase.GetPara("�����˲���Ʊ�����ֽ������", glngSys, mlngModule, "")) = 1
  
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
        Set objTemp = .CreatePane(mPanel.Pane_List, 300, 100, DockBottomOf, objPane)
        objTemp.Tag = mPanel.Pane_List
        objTemp.Title = "�����б�": objTemp.Options = PaneNoCloseable Or PaneNoHideable
        objTemp.Handle = picList.hWnd
        
        Set objPane = .CreatePane(mPanel.Pane_Balance, 100, 100, DockRightOf, objTemp)
        objPane.Tag = mPanel.Pane_Balance
        objPane.Title = "������Ϣ": objPane.Options = PaneNoCloseable Or PaneNoHideable
        objPane.Handle = picBalance.hWnd
       '
        Set objPane = .CreatePane(mPanel.Pane_Detail, 300, 100, DockBottomOf, objTemp)
        objPane.Tag = mPanel.Pane_Detail
        objPane.Title = "������ϸ�б�": objPane.Options = PaneNoCloseable Or PaneNoHideable
        objPane.Handle = PicDetail.hWnd
       '  .SetCommandBars Me.cbsThis
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
    Call ReadListData
End Sub

Private Sub cmdCancel_Click()
     Unload Me
End Sub

Private Sub cmdClear_Click()
    With vsList
        If .Rows <= .FixedRows Then Exit Sub
        If .ColIndex("ѡ��") < 0 Then Exit Sub
        .Cell(flexcpChecked, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 0
    End With
    Call SetBlanceShow
    Call InitPatiInsure
End Sub

Private Function zlMakeupPrint(ByVal lng����ID As Long, _
    ByVal strNos As String, _
    ByVal strUseType As String, _
    ByVal strBillNameDemo As String, _
    ByVal intInvoiceFormat As Integer, _
    ByVal blnVirtualPrint As Boolean, _
    ByVal intInusre As Integer, _
    ByRef lng����ID As Long, _
    ByVal lngShareUseID As Long, _
    ByVal strFactNO As String, _
    Optional ByVal str����IDs As String = "", _
    Optional strPriceGrade As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ��
    '���:strNos-��Ҫ��ӡ��NO
    '     strUseType-ʹ�����
    '     strBillNameDemo-Ʊ�ݸ�ʽ˵��
    '     intInvoiceFormat-��Ʊ��ӡ��ʽ
    '     blnVirtualPrint-�Ƿ�ҽֻ�ӿڴ�ӡƱ��
    '     intInusre-����
    '     blnOnePrint-�Ƿ�һ�δ�ӡ(true-��һ�δ�ӡ�����ֽ������,����ֽ��������ӡ)
    '     strFactNo-��Ʊ��
    '     str����IDs-���δ�ӡ�漰�Ľ���IDs,����ö��ŷָ�
    '����:lng����ID-��������ID
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-04 22:58:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    Dim bln�ֱ��ӡ As Boolean, lng��ӡID As Long
    


    If strNos = "" Then Exit Function
    If lng����ID = 0 Then Exit Function
    If strNos = "" Then
        MsgBox "��ǰû�е��ݿ����ش�Ʊ�ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not CheckFP(strNos, strUseType, strBillNameDemo, strFactNO, lng����ID, lngShareUseID) Then Exit Function
           
    '--------------------------------------------------------------------------------------
    '������ʱ����
    If mbln���ֽ������ Then
        If zlSaveTempPrintData(strNos, lng����ID, strFactNO, lng��ӡID) = False Then Exit Function
    End If
    '--------------------------------------------------------------------------------------
    
    '������ʣ�������Ĳſ����ش򣬱���ҽ������ʹ������Ҳ�������´�ӡ
    If Not blnVirtualPrint Then
        If Not BillExistMoney(strNos, 1, , lng��ӡID) Then
            MsgBox "����[" & strNos & "]�е���Ŀ�Ѿ�ȫ���˷ѣ����ܽ��д�ӡ��", vbInformation, gstrSysName
            Call zlDeleteTempPrintData(lng��ӡID)
            Exit Function
        End If
    End If
    
    'Ƚ����,2014-12-17,���������շѵ��ݲ������ش�Ʊ��
    If CheckBillExistReplenishData(2, , Replace(strNos, "'", ""), lng��ӡID) = True Then
        MsgBox "����[" & strNos & "]�д����Ѿ������˱��ղ���������Ŀ�����ܽ��д�ӡ��", vbInformation, gstrSysName
        Call zlDeleteTempPrintData(lng��ӡID)
        Exit Function
    End If
    
    Dim dtDate As Date
    dtDate = zlDatabase.Currentdate
    
    bln�ֱ��ӡ = gTy_Module_Para.bln�ֱ��ӡ
    If mbln���ֽ������ Then bln�ֱ��ӡ = False
    
    strNos = "'" & Replace(strNos, ",", "','") & "'"
    Call frmPrint.ReportPrint(1, strNos, "", "", lng����ID, lngShareUseID, strFactNO, dtDate, "", "", _
        bln�ֱ��ӡ, intInvoiceFormat, blnVirtualPrint, , strUseType, , mbln���ֽ������, lng��ӡID, strPriceGrade)
    
    mintSucces = mintSucces + 1
    zlMakeupPrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckFP(ByVal strNos As String, ByVal strUserType As String, ByVal strBillNameDemo As String, ByRef strFactNO As String, ByRef lng����ID As Long, ByRef lngShareID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鷢Ʊ�Ƿ���ȷ
    '���:strNos-����NO����ȡ��Ʊ��
    '     strUserType-Ʊ��ʹ�����
    '     lngShareID-��ǰ��������
    '     strFactNo-��Ʊ��
    '����:lng����ID-��������ID
    '     lngShareID-���ع���ID
    '     strFactNo-��Ʊ��
    '����: ��Ʊ�Ϸ� ����true,���򷵻�False
    '����:���˺�
    '����:2012-07-12 11:30:22
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intNum As Integer, varData As Variant
    
    On Error GoTo errHandle
    varData = Split(strNos, ",")
    intNum = UBound(varData) + 1
    If strNos = "" Then
        MsgBox "��������Ҫ�����Ʊ��", vbInformation, gstrSysName
        Exit Function
    End If
    If Not gblnStrictCtrl Then
       If Len(strFactNO) <> gbytFactLength And strFactNO <> "" Then
           MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytFactLength & " λ��", vbInformation, gstrSysName
           If InputFactNo(strUserType, strBillNameDemo, lng����ID, lngShareID, strFactNO) Then Exit Function

        End If
       CheckFP = True
       Exit Function
    End If
     
    If Trim(strFactNO) = "" Then
       MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
       Exit Function
    End If
    
    If Not gTy_Module_Para.bln�ֱ��ӡ Or mbln���ֽ������ Then intNum = 1
 
InvoiceHandle:
    If zlCheckInvoiceValied(lng����ID, intNum, strFactNO, lngShareID, strUserType) = False Then Exit Function

    '�����������,Ʊ���Ƿ�����
    If CheckBillRepeat(lng����ID, 1, strFactNO) Then
        MsgBox "Ʊ�ݺ�""" & strFactNO & """�Ѿ���ʹ�ã����������롣", vbInformation, gstrSysName
        If mblnStartFactUseType = False Then
            txtInvoice.Text = GetNextFactNo(strUserType, lng����ID, lngShareID)
        End If
        If InputFactNo(strUserType, strBillNameDemo, lng����ID, lngShareID, strFactNO) Then Exit Function
    End If
   CheckFP = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckPrintValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ӡ�ĺϷ���
    '���:
    '����:
    '����:���˺�
    '����:2016-04-29 11:58:23
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then
        MsgBox "δѡ��ָ���Ĳ���,��ѡ����Ҫ��ӡ��Ʊ�Ĳ���!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    
    If mrsInfo.State <> 1 Then
        MsgBox "δѡ��ָ���Ĳ���,��ѡ����Ҫ��ӡ��Ʊ�Ĳ���!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.EOF Then
        MsgBox "δѡ��ָ���Ĳ���,��ѡ����Ҫ��ӡ��Ʊ�Ĳ���!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    
    CheckPrintValied = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SplitGroupPrint(ByRef cllPrint As Collection, ByRef cllUseType As Collection, _
    ByRef cllRegistNos As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Զ����з���,�Ա�����ӡ
    '����:cllPrint-�����ӡ����()
    '     ��ʽ:array(Key,����IDs,�������s,���ݺ�,ʹ�����,Ʊ�ݸ�ʽ,�Ƿ�ҽ���ӿڴ�ӡ,����),"K_" & Ʊ�ݸ�ʽ & "_" & ���� & "_" & �ӿڴ�ӡ��־ & "_"  & �������
    '     cllUseType-��ǰѡ��Ҫ��ӡ��Ʊ�ݵ�ʹ�����,��ʽ:array(ʹ�����,Ʊ��)����K" & ʹ�����
    '����:����ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2016-04-29 12:00:30
    '˵����95543
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String, strKey As String, lngRow As Long
    Dim blnVirtualPrint As Boolean, blnYb As Boolean
    Dim lng����ID As Long, lng������� As String, intInsure As Integer, lng����ID As Long
    Dim str����IDs As String, str�������s As String, strNos As String
    Dim intPrintFormat As Integer, strUserType As String
    Dim cllUserTypes As New Collection, strInsureIDs As String
    Dim lngTemp As Long, varData As Variant, intGeneralFromat As Integer, strGeneralUserType As String
    Dim strUseType As String, strUseTypes As String, strBillNameDemo As String
    
    On Error GoTo errHandle
    'һ����������˲���Ʊ�������������:
    '1.���ҽ������ͨ����ʹ�õ���ͬ��Ʊ(����ʹ�����)��ͬһ�ַ�Ʊ��ʽ��ͬʱҽ���ӿڲ���ӡ��Ʊ���򲻷�ҽ������ͨ���ˣ�һ���ӡ.
    '2.���ҽ������ͨ����ʹ�ò�ͬ��Ʊ(��ʹ�����)��ͬ��Ʊ��ʽ��ͬʱҽ���ӿڲ���ӡ������Ҫ��ҽ������ͨ���ˣ��ֱ��ӡ.
    '3.���ҽ���ӿڴ�ӡ�����Ǹ��ݽӿڷ��صĵ��������飬ȷ����ӡ����(�ӿڴ�ӡ�ķ���һ�𣬽ӿڲ���ӡ�ķ���һ��)
    '4.�����˲���Ʊʱ���ֵ��ݴ�ӡ��ʧЧ!
    '�����������˴�ӡ��Ʊ����ֽ���������д�ӡ
    Set cllPrint = New Collection
    Set cllUseType = New Collection
    Set cllRegistNos = New Collection
    
    lng����ID = Val(Nvl(mrsInfo!����ID))
    
    '��ͨ��ʽ
    strGeneralUserType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
 
    intGeneralFromat = zl_GetInvoicePrintFormat(mlngModule, strGeneralUserType, , mbln���ֽ������)
    
    Set cllUserTypes = New Collection
    strInsureIDs = "": strUseTypes = ""
    With vsList
        For lngRow = 1 To .Rows - 1
            If .IsSubtotal(lngRow) Then
                '�����в�����
            ElseIf GetVsGridBoolColVal(vsList, lngRow, .ColIndex("ѡ��")) Then
                strNo = Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�")))
                If .TextMatrix(lngRow, .ColIndex("����")) = "�Һŵ�" Then
                    cllRegistNos.Add strNo
                Else
                    lng������� = Val(.TextMatrix(lngRow, .ColIndex("�������ID")))
                    lng����ID = Val(.TextMatrix(lngRow, .ColIndex("����ID")))
                    blnYb = .TextMatrix(lngRow, .ColIndex("ҽ��")) = "��"
                    intInsure = .TextMatrix(lngRow, .ColIndex("����ID"))
                    
                    blnVirtualPrint = False
                    If intInsure <> 0 Then  'InStr(strInsureIDs & ",", "," & intInsure & ",") = 0 And
                        blnVirtualPrint = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure, CStr(lng����ID))
                    End If
                    
                    '�ж�ʹ�����
                    If InStr(strInsureIDs & ",", "," & intInsure & ",") = 0 Then
                        strInsureIDs = strInsureIDs & "," & intInsure
                        strUseType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
                        intPrintFormat = zl_GetInvoicePrintFormat(mlngModule, strUseType, , mbln���ֽ������)
                        If mblnStartFactUseType = False Then strUseType = ""
                        cllUserTypes.Add Array(strUseType, intPrintFormat), "K" & intInsure
                    Else
                        strUseType = cllUserTypes("K" & intInsure)(0)
                        intPrintFormat = cllUserTypes("K" & intInsure)(1)
                    End If
                    
                    '��ȡʹ�����
                    If InStr(1, strUseTypes & ",", "," & IIf(strUseType = "", "-", strUseType) & ",") = 0 Then
                        strBillNameDemo = ZlGetBillFormat(mlngModule, intPrintFormat)
                        cllUseType.Add Array(strUseType, strBillNameDemo), "K" & strUseType
                        strUseTypes = strUseTypes & "," & IIf(strUseType = "", "-", strUseType)
                    End If
                    
                    lngTemp = IIf(mbln���ֽ������, 0, lng�������)
                
                    '104391
                    '1.���ҽ������ͨ����ʹ�õ���ͬ��Ʊ(����ʹ�����)��ͬһ�ַ�Ʊ��ʽ��ͬʱҽ���ӿڲ���ӡ��Ʊ���򲻷�ҽ������ͨ���ˣ�һ���ӡ.
                    If Not blnVirtualPrint And Not mblnStartFactUseType And intPrintFormat = intGeneralFromat And intInsure <> 0 Then
                        'һ���ӡ:1.����ҽ���ӿڴ�ӡ
                        '         2.����ʹ�����������ͨƱ����һ�ָ�ʽ
                        intInsure = 0
                    End If
                    
                    'Key:"K_" & Ʊ�ݸ�ʽ & "_" & ���� & "_" & �ӿڴ�ӡ��־ & "_"  & �������
                    strKey = "K_" & intPrintFormat & "_" & intInsure & "_" & IIf(blnVirtualPrint, 1, 0) & "_" & lngTemp
                    'array(Key,����IDs,�������s,���ݺ�,ʹ�����,Ʊ�ݸ�ʽ,�Ƿ�ҽ���ӿڴ�ӡ,����)
                    If FindCllKeyIsExsits(cllPrint, strKey) Then
                        varData = cllPrint(strKey)
                        str����IDs = varData(1) & "," & lng����ID
                        str�������s = varData(1) & "," & lng�������
                        strNos = varData(3) & "," & strNo
                        cllPrint.Remove strKey
                    Else
                        str����IDs = lng����ID
                        str�������s = lng�������
                        strNos = strNo
                    End If
                    cllPrint.Add Array(strKey, str����IDs, str�������s, strNos, strUseType, intPrintFormat, IIf(blnVirtualPrint, 1, 0), intInsure), strKey
                End If
            End If
        Next
    End With

    SplitGroupPrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function FindCllKeyIsExsits(ByVal cllData As Collection, ByVal strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҽ����е�Keyֵ�Ƿ���ڣ����ڷ���true,���򷵻�False
    '���:cllData-��������
    '     strKey-���ҵ�Keyֵ
    '����:���Key���ڣ�����True,���򷵻�False
    '����:���˺�
    '����:2016-05-03 10:57:45
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrData As Variant
    Err = 0: On Error Resume Next
    arrData = cllData(strKey)
    If Err <> 0 Then Err = 0: Exit Function
    FindCllKeyIsExsits = True
    Exit Function
End Function

Private Sub cmdOK_Click()
    Dim cllPrint As Collection, arrPrint As Variant
    Dim strNos As String, str����IDs As String, lng����ID As Long
    Dim i As Long, j As Long, strPrintUserType As String
    Dim cllUseType As Collection, strFactNO As String
    Dim strUseType As String, strBillNameDemo As String, lngShareUseID As Long, lng����ID As Long
    Dim strPriceGrade As String
    Dim cllRegistNos As Collection
    Dim blnPrintSccess As Boolean
    
    On Error GoTo errHandler
    '1.Ʊ�ݴ�ӡ�ĺϷ��Լ��
    If Not CheckPrintValied Then Exit Sub
    lng����ID = Val(Nvl(mrsInfo!����ID))
      
    '2.�ֽ�Ʊ�ݴ�ӡ����
    '   ��ʽ:array(Key,����IDs,�������s,���ݺ�,ʹ�����,Ʊ�ݸ�ʽ,�Ƿ�ҽ���ӿڴ�ӡ,����),"K_" & Ʊ�ݸ�ʽ & "_" & ���� & "_" & �ӿڴ�ӡ��־ & "_"  & �������
    If SplitGroupPrint(cllPrint, cllUseType, cllRegistNos) = False Then Exit Sub

    If cllPrint.Count = 0 And cllRegistNos.Count = 0 Then
        MsgBox "δѡ����Ҫ��ӡ��Ʊ��,��ѡ����ٲ���Ʊ��", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    '��ȡ�۸�ȼ�
    If gintPriceGradeStartType >= 2 Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, lng����ID, 0, "", , , strPriceGrade)
    Else
        strPriceGrade = gstr��ͨ�۸�ȼ�
    End If
    
    '3.������ص��շѵ�Ʊ�ݴ�ӡ
    If cllPrint.Count > 0 Then
        For j = 1 To cllUseType.Count
            strUseType = cllUseType(j)(0)
            strBillNameDemo = cllUseType(j)(1)
            'ȷ����������
            lngShareUseID = zl_GetInvoiceShareID(mlngModule, strUseType)
            If InputFactNo(strUseType, strBillNameDemo, lng����ID, lngShareUseID, strFactNO) = False Then GoTo PrintEnd:
            
            For i = 1 To cllPrint.Count
                'array(Key,����IDs,�������s,���ݺ�,ʹ�����,Ʊ�ݸ�ʽ,�Ƿ�ҽ���ӿڴ�ӡ,����)
                arrPrint = cllPrint(i)
                If arrPrint(4) = strUseType Then
                    strNos = strNos & "," & arrPrint(3)
                    str����IDs = str����IDs & "," & arrPrint(1)
                    '��ȡƱ��
                    If Not zlMakeupPrint(lng����ID, arrPrint(3), strUseType, strBillNameDemo, Val(arrPrint(5)), _
                        IIf(Val(arrPrint(6)) = 1, True, False), Val(arrPrint(7)), lng����ID, lngShareUseID, _
                        strFactNO, str����IDs, strPriceGrade) Then GoTo PrintEnd:
                    blnPrintSccess = True
                    strFactNO = GetNextFactNo(arrPrint(4), lng����ID, lngShareUseID)
                    txtInvoice.Text = strFactNO
                End If
            Next
        Next
        If str����IDs <> "" Then str����IDs = Mid(str����IDs, 2)
        If strNos <> "" Then strNos = Mid(strNos, 2)
          
        '��ҽһ��ͨд����85950
        Call WriteInforToCard(Me, mlngModule, mstrPrivs, mSquareCard.objSquareCard, 0, strNos)
        
        '81688:���ϴ�,2015/5/18,������
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            Call gobjPlugIn.OutPatiInvoicePrintAfter(lng����ID, str����IDs)
            Err.Clear
        End If
    End If
    
    If cllRegistNos.Count > 0 Then
        If PrintRegistBill(cllRegistNos, lng����ID, blnPrintSccess) = False Then GoTo PrintEnd:
    End If
    GoTo PrintEnd:
    
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
PrintEnd:
    If blnPrintSccess Then Call ReadListData 'ˢ������
End Sub

Private Function PrintRegistBill(ByVal cllRegistNos As Collection, ByVal lng����ID As Long, _
    ByRef blnPrinted As Boolean) As Boolean
    '����Һŵ���
    Dim strCommon As String, intAtom As Integer, blnOK As Boolean
    Dim i As Long, blnFirstNO As Boolean
    
    On Error Resume Next
    If gobjRegist Is Nothing Then
        Set gobjRegist = CreateObject("zl9RegEvent.clsRegEvent")
        If gobjRegist Is Nothing Then Exit Function
    End If
    
    Err.Clear: On Error GoTo 0
    '�������úϷ�������
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    
    For i = 1 To cllRegistNos.Count
        'Public Function PrintRegistBill(frmMain As Object, cnMain As ADODB.Connection, _
         ByVal lngSys As Long, ByVal strDbUser As String, _
         ByVal strNO As String, ByVal lng����ID As Long, _
         Optional ByVal blnFirstNO As Boolean) As Boolean
        blnFirstNO = (i = 1)
        If gobjRegist.PrintRegistBill(Me, gcnOracle, glngSys, gstrDBUser, cllRegistNos(i), lng����ID, blnFirstNO) = False Then
             Call GlobalDeleteAtom(intAtom)
             Exit Function
        End If
        blnPrinted = True
    Next
    
    Call GlobalDeleteAtom(intAtom)
    PrintRegistBill = True
End Function

Private Function InputFactNo(ByVal strUseType As String, ByVal strBillNameDemo As String, ByRef lng����ID As Long, ByRef lngShareUseID As Long, ByRef strFactNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч�ķ�Ʊ��
    '���:strUseType-ʹ�����
    '     strBillNameDemo-Ʊ������˵��
    '     lng����ID-��ǰ������ID
    '     lngShareUseID-��������ID
    '����:���صķ�Ʊ��
    '����:����ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2016-06-08 11:00:28
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnValid As Boolean
    On Error GoTo errHandle
    
    If Not mblnStartFactUseType Then
        '������ʹ������ʱ��ֱ�Ӵ���������¼��ķ�Ʊ����ȡ��
        strFactNO = Trim(txtInvoice.Text)
        If strFactNO = "" Then GoTo ReInput:
        If gblnStrictCtrl Then
            If Not zlCheckInvoiceValied(lng����ID, 1, strFactNO, lngShareUseID, strUseType) Then Exit Function
        End If
        InputFactNo = True
        Exit Function
    End If
    
ReInput:
    Do
        '����Ʊ�����ö�ȡ
        blnValid = False
        'ȷ����������
        strFactNO = GetNextFactNo(strUseType, lng����ID, lngShareUseID)
        
        If frmInputBox.InputBox(Me, "��Ʊ������:" & IIf(strUseType = "", "", "��" & strUseType & "������ʽ:" & strBillNameDemo), "��ȷ�ϲ���ʹ�õĿ�ʼƱ�ݺ��룺", 30, 1, False, False, strFactNO, _
        False, Me.Left + 1500, Me.Top + 1500) = False Then Exit Function
        '�û�ȡ������,����ӡ
        If strFactNO = "" Then Exit Function
        If gblnStrictCtrl Then
            If zlCheckInvoiceValied(lng����ID, 1, strFactNO, lngShareUseID, strUseType) Then blnValid = True
        Else
            blnValid = True
        End If
    Loop While Not blnValid
    
    InputFactNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdSelAll_Click()
    With vsList
        If .Rows <= .FixedRows Then Exit Sub
        If .ColIndex("ѡ��") < 0 Then Exit Sub
        .Cell(flexcpChecked, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = -1
    End With
    Call SetBlanceShow
    Call InitPatiInsure
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
    Call zlClearPatiInfor
    Call ReadListData: Call ShowDetail '����һ�½���
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If cmdSelAll.Visible And cmdSelAll.Enabled Then Call cmdSelAll_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If cmdClear.Visible And cmdClear.Enabled Then Call cmdClear_Click
    End If
End Sub

Private Sub Form_Load()
    Dim intPrintDays As Integer
    
    mblnFirst = True
    mblnStartFactUseType = zlStartFactUseType(1)
    mlng����ID = 0
    lblFormat.Alignment = 0

    dtpBegin.MaxDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59")
    dtpEnd.MaxDate = dtpBegin.MaxDate
    intPrintDays = Val(zlDatabase.GetPara("ȱʡ��Ʊ��ӡ����", glngSys, mlngModule, "0"))
    If intPrintDays <= 0 Then
        chkDate.Value = vbChecked
        intPrintDays = 7
    Else
        chkDate.Value = vbUnchecked
    End If
    dtpBegin.Enabled = (chkDate.Value = vbUnchecked): dtpEnd.Enabled = dtpBegin.Enabled
    dtpBegin.Value = Format(DateAdd("d", -1 * (intPrintDays - 1), dtpBegin.MaxDate), "yyyy-mm-dd")
    dtpEnd.Value = Format(dtpEnd.MaxDate, "yyyy-mm-dd")
    chkRegistFee.Value = Val(zlDatabase.GetPara("�����˲���Ʊ�ݺ��Һŷ�", glngSys, mlngModule, "0"))
    
    'δ����ʹ�����ʱ������������������ʾ
    txtInvoice.Visible = Not mblnStartFactUseType
    lblFact.Visible = Not mblnStartFactUseType
    lblFormat.Visible = Not mblnStartFactUseType
 
    Call InitPanel
    Call zlCardSquareObject
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
    zlDatabase.SetPara "�����˲���Ʊ�ݺ��Һŷ�", chkRegistFee.Value, glngSys, mlngModule
    Call zlCardSquareObject(True)
    zlSaveDockPanceToReg Me, dkpMan, "����"
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Caption, "�����б�", False
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "�����б�", False
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "������ϸ�б�", False
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If txtPatient.Locked Then Exit Sub
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

Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBalance.Left = .ScaleLeft
        vsBalance.Width = .ScaleWidth
        vsBalance.Height = .ScaleHeight - lbl�ϼ�.Height - 50
        vsBalance.Top = .ScaleTop
        lbl�ϼ�.Top = .ScaleHeight - lbl�ϼ�.Height - 10
        lbl�ϼ�.Left = .ScaleLeft
    End With
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
        '75259�����ϴ�,2014-7-10��������������ʾ��ɫ����
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(IsNull(mrsInfo!����), Me.ForeColor, vbRed))
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
    '����:2011-09-04 18:09:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtPatient.Text = ""

    Set mrsInfo = New ADODB.Recordset
    vsList.Clear 1: vsList.Rows = 1: vsDetail.Clear 1: vsDetail.Rows = 1
    vsBalance.Clear 1: vsBalance.Rows = 1
    lbl�ϼ�.Caption = "����ϼ�:" & Format(0, "0.00")
    
    Set mrsList = Nothing: Set mrsDetail = Nothing: Set mrsBalance = Nothing
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
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
    Call ReadListData
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
 
Private Sub ShowDetail(Optional ByVal byt��¼���� As Byte = 1, Optional ByVal strNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ϸ����
    '����:
    '����:���˺�
    '����:2011-09-04 20:18:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, strSQL As String
    Dim lngCol As Long
    
    On Error GoTo errHandler
    strSQL = _
    " Select C.���� as ���,Nvl(E.����,B.����) as ����," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & "B.���," & _
            IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
    "       Trim(To_Char(Avg(Nvl(A.����,1)*A.����)" & _
            IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ",'9999990.00000')) as ����, " & _
    "       Trim(To_Char(Sum(A.��׼����)" & _
            IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "')) as ����, " & _
    "       Trim(To_Char(Sum(A.Ӧ�ս��),'9999999" & gstrDec & "')) as Ӧ�ս��, " & _
    "       Trim(To_Char(Sum(A.ʵ�ս��),'9999999" & gstrDec & "')) as ʵ�ս��, " & _
    "       D.���� as ִ�п���,Nvl(A.��������,B.��������) as ����," & _
    "       Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��',9,'�쳣�շѵ�','��'||ABS(A.ִ��״̬)||'���˷�') as ˵��" & _
    " From  ������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,�շ���Ŀ���� E,ҩƷ��� X" & _
              IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
    " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
    "       And A.��¼����=[1] And A.NO=[2] And A.��¼״̬ IN(1,3)" & _
    "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
            IIf(Not gblnShowErr, " And Nvl(A.���ӱ�־,0)<>9", "") & _
    " Group by Nvl(A.�۸񸸺�,A.���),C.����,Nvl(E.����,B.����)," & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� ,", "") & " B.���,A.���㵥λ,A.�ѱ�,D.����," & _
    "       Nvl(A.��������,B.��������),A.ִ��״̬,A.��¼״̬,X.ҩƷID,X." & gstrҩ����λ & ",Nvl(X." & gstrҩ����װ & ",1)" & _
    " Order by Nvl(A.�۸񸸺�,A.���)"
    Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, byt��¼����, strNo)
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
        .ColHidden(.ColIndex("��λ")) = byt��¼���� = 4 '�Һŵ����ء���λ����
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        zl_vsGrid_Para_Restore mlngModule, vsDetail, Me.Caption, "������ϸ�б�", False
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitBlanceData(ByVal str���ý��� As String, ByVal str�ҺŽ��� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '���:str���ý���-���õ��ݣ�ָ���Ľ�����ţ���ʽ���������,�������,...
    '     str�ҺŽ���-�Һŵ��ݣ�ָ���Ľ���ID����ʽ������ID,����ID,...
    '����:
    '����:���˺�
    '����:2011-09-04 21:32:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strSubSql As String
    
    Err = 0: On Error GoTo errHandle
    If str���ý��� = "" And str�ҺŽ��� = "" Then
        Set mrsBalance = Nothing
        InitBlanceData = True: Exit Function
    End If
    
    If str���ý��� <> "" Then
        If zlStr.ActualLen(str���ý���) <= 4000 Then
            strSubSql = "Select Column_Value As ������� From Table(f_Str2list([1]))"
        Else
            strSubSql = FromStrListBulidSQL(str���ý���, "�������")
        End If
        strSQL = _
            "Select /*+cardinality(c,10)*/'�շѵ�' As ����,Min(a.No)||Decode(Min(a.No),Max(a.No),'','��'||Max(a.No)) As NO, b.������� As �������" & vbNewLine & _
            "From ������ü�¼ A, ����Ԥ����¼ B,(" & strSubSql & ") C" & vbNewLine & _
            "Where a.����id = b.����id And Mod(a.��¼����,10)=1 And (b.������� = c.������� Or b.����id = c.�������)" & vbNewLine & _
            "Group By b.�������" & vbNewLine
    End If
    If str�ҺŽ��� <> "" Then
        If strSQL <> "" Then strSQL = strSQL & "Union All"
        If zlStr.ActualLen(str�ҺŽ���) <= 4000 Then
            strSubSql = "Select Column_Value As ����id From Table(f_Str2list([2]))"
        Else
            strSubSql = FromStrListBulidSQL(str�ҺŽ���, "����id")
        End If
        strSQL = strSQL & vbNewLine & _
            "Select /*+cardinality(c,10)*/'�Һŵ�' As ����,Min(a.No)||Decode(Min(a.No),Max(a.No),'','��'||Max(a.No)) As NO, b.����id As �������" & vbNewLine & _
            "From ������ü�¼ A, ����Ԥ����¼ B,(" & strSubSql & ") C" & vbNewLine & _
            "Where a.����id = b.����id And a.��¼����=4 And b.����id = c.����id" & vbNewLine & _
            "Group By b.����id"
    End If

    '������Ϣ
    strSQL = _
        " Select Max(����) as ����, Max(t.No) As NO," & vbNewLine & _
        "       Decode(Mod(s.��¼����, 10), 1, '��Ԥ��', s.���㷽ʽ) As ���㷽ʽ, Sum(s.��Ԥ��) As ���, t.�������" & vbNewLine & _
        " From ����Ԥ����¼ S, (" & strSQL & ") T" & vbNewLine & _
        " Where s.������� = t.������� Or s.����id = t.�������" & vbNewLine & _
        " Group By t.����, t.�������, Decode(Mod(s.��¼����, 10), 1, '��Ԥ��', s.���㷽ʽ)"
            
    Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str���ý���, str�ҺŽ���)
    InitBlanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FromStrListBulidSQL(ByVal strData As String, _
    Optional ByVal strColumnName As String, _
    Optional ByVal strSplit As String = ",") As String
    '���ܣ���ȡ���ַ����б��SQL,�ַ������ȳ���4000ʱ
    Dim strSQL As String
    Dim varData As Variant, i As Long, strTemp As String
    
    On Error GoTo errHandler
    varData = Split(strData, strSplit)
    For i = 0 To UBound(varData)
        If zlStr.ActualLen(strTemp) > 4000 Then
            strSQL = strSQL & _
                " Union All" & _
                " Select Column_Value" & IIf(strColumnName <> "", " As " & strColumnName, "") & _
                " From Table(f_Str2list('" & Mid(strTemp, 2) & "'))"
            strTemp = ""
        End If
        strTemp = strTemp & "," & varData(i)
    Next
    If strTemp <> "" Then
        strSQL = strSQL & _
            " Union All" & _
            " Select Column_Value" & IIf(strColumnName <> "", " As " & strColumnName, "") & _
            " From Table(f_Str2list('" & Mid(strTemp, 2) & "'))"
    End If
    If strSQL <> "" Then strSQL = Mid(strSQL, 11)
    FromStrListBulidSQL = strSQL
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetBlanceShow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ���㷽ʽ
    '���:blnAllSel-ѡ�����еĵ���
    '����:���˺�
    '����:2011-02-23 14:54:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String, lngRow As Long, i As Long, str���� As String
    Dim blnȫѡ As Boolean, blnδѡ As Boolean
    Dim strFilter As String
    Dim strSelNos As String, dblMoney As Double
    Dim str�������� As String, lng������� As Long
    
    lbl�ϼ�.Caption = "����ϼ�:0.00"
    vsBalance.Clear 1: vsBalance.Rows = 1
    If mrsBalance Is Nothing Then Exit Sub
    
    With vsList
        blnȫѡ = True: blnδѡ = True
        For lngRow = 1 To .Rows - 1
            If .IsSubtotal(lngRow) = False Then
                str�������� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
                If str�������� = "�Һŵ�" Then
                    lng������� = .TextMatrix(lngRow, .ColIndex("����ID"))
                Else
                    lng������� = .TextMatrix(lngRow, .ColIndex("�������ID"))
                End If
                If GetVsGridBoolColVal(vsList, lngRow, .ColIndex("ѡ��")) Then
                    If InStr(1, strSelNos & ",", "," & str�������� & ":" & lng������� & ",") = 0 Then
                        strSelNos = strSelNos & "," & str�������� & ":" & lng�������
                        blnδѡ = False
                        
                        If strFilter <> "" Then strFilter = strFilter & " Or "
                        strFilter = strFilter & "(����='" & str�������� & "' And �������=" & lng������� & ")"
                    End If
                End If
                If InStr(1, strSelNos & ",", "," & str�������� & ":" & lng������� & ",") = 0 Then blnȫѡ = False
            End If
        Next
    End With
    If strSelNos <> "" Then strSelNos = Mid(strSelNos, 2)
    
    '��ʾ����ѡ��ĵ��ݵĽ��㷽ʽ֮��
    If blnȫѡ Or blnδѡ Then
        mrsBalance.Filter = 0
    Else
        mrsBalance.Filter = strFilter
    End If
    mrsBalance.Sort = "����,NO Desc,���㷽ʽ"
    
    With vsBalance
        .Redraw = flexRDNone
        Set .DataSource = mrsBalance
        
        If chkRegistFee.Value = vbChecked Then
            '������ʾ
            .OutlineBar = flexOutlineBarComplete
            .Subtotal flexSTClear
            .Subtotal flexSTNone, .ColIndex("����"), , , &H8000000F
            .OutlineCol = .ColIndex("���ݺ�")
        End If
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) Then
                .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(i, 0)
            Else
                dblMoney = dblMoney + Val(.TextMatrix(i, .ColIndex("������")))
                .TextMatrix(i, .ColIndex("������")) = FormatEx(Val(.TextMatrix(i, .ColIndex("������"))), 6, , , 2)
            End If
        Next
        
        '���кϲ�
        .MergeCol(.ColIndex("���ݺ�")) = True
        .MergeCells = flexMergeRestrictColumns
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Caption, "�����б�", False
        .Redraw = flexRDBuffered
        lbl�ϼ�.Caption = "����ϼ�:" & Format(dblMoney, "0.00")
    End With
End Sub

Private Function ReadListData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ҫ��ϸ����
    '����:��ȡ�ɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-01-25 17:10:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strTable As String
    Dim lngCol As Long, strSQL As String
    Dim strWhere As String, dtStartDate As Date, dtEndDate As Date
    Dim i As Long, lng������� As Long, byt�������� As Byte
    Dim blnRemove As Boolean, blnVirtualPrint As Boolean
    Dim intInsure As Integer
    Dim lng����ID As Long, dblʣ������ As Double
    Dim strPreNo As String, strPreNoType As String
    Dim str���ý��� As String '������ţ���ʽ���������,�������,...
    Dim str�ҺŽ��� As String '����ID����ʽ������ID,����ID,...
    Dim j As Long
    
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
        strWhere = strWhere & " And A.����ʱ�� betWeen [2] and [3]"
        dtStartDate = CDate(Format(dtpBegin.Value, "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpEnd.Value, "yyyy-mm-dd") & " 23:59:59")
    End If
    '�ų����в������ĵ���(�շѽ���ID�ͷ����һ���˷ѵĽ���ID���ڷ��ò����¼���շѽ���id�У������һ���˷ѵĽ���ID����)
    If chkRegistFee.Value = vbChecked Then
        strWhere = strWhere & vbNewLine & _
            " And Mod(a.��¼����, 10) In (1, 4) " & vbNewLine & _
            " And Not Exists(Select 1 From ���ò����¼ Where ��¼���� = 1 " & _
                           "And (Mod(a.��¼����,10)=1 And Nvl(���ӱ�־,0)=0 Or a.��¼����=4 And Nvl(���ӱ�־,0)=1) And �շѽ���id = a.����id)"
    Else
        strWhere = strWhere & _
            " And Mod(a.��¼����, 10) = 1 " & vbNewLine & _
            " And Not Exists(Select 1 From ���ò����¼ Where ��¼���� = 1 And Nvl(���ӱ�־, 0) = 0 And �շѽ���id = a.����id)"
    End If
    strWhere = strWhere & vbNewLine & _
        " And Not Exists(Select 1 From ���ò����¼ M, ����Ԥ����¼ N Where m.������� = n.������� And n.����id = a.����id)"
    
    mblnSel = False
    On Error GoTo errHandle
    zlCommFun.ShowFlash "���ڶ�ȡ��������,���Ժ� ..."
    Screen.MousePointer = 11
    DoEvents
    strTable = "" & _
            " Select Mod(a.��¼����,10) As ��¼����,a.No, Max(a.ʵ��Ʊ��) As ʵ��Ʊ��, Max(a.����) As ����, Max(a.�Ա�) As �Ա�, Max(a.����) As ����," & vbNewLine & _
            "        Max(Decode(a.�����־, 2, '', a.��ʶ��)) As �����, Max(Decode(a.�����־, 2, a.��ʶ��, '')) As סԺ��," & vbNewLine & _
            "        Max(a.�ѱ�) As �ѱ�, Max(a.������) As ������, Max(a.��������id) As ��������id, Max(a.���ʽ) As ���ʽ," & vbNewLine & _
            "        Max(a.������) As ������," & vbNewLine & _
            "        Max(Decode(Decode(a.��¼����,4,1,a.��¼����), 1, Decode(a.��¼״̬, 1, a.����Ա����, 3, a.����Ա����, ''), '')) As ����Ա����," & vbNewLine & _
            "        Max(Decode(Decode(a.��¼����,4,1,a.��¼����), 1, Decode(a.��¼״̬, 1, a.�Ǽ�ʱ��, 3, a.�Ǽ�ʱ��, Null), Null)) As �Ǽ�ʱ��," & vbNewLine & _
            "        Sum(Decode(Decode(a.��¼����,4,1,a.��¼����), 1, Decode(a.��¼״̬, 1, a.Ӧ�ս��, 3, a.Ӧ�ս��, 0), 0)) As Ӧ�ս��," & vbNewLine & _
            "        Sum(Decode(Decode(a.��¼����,4,1,a.��¼����), 1, Decode(a.��¼״̬, 1, a.ʵ�ս��, 3, a.ʵ�ս��, 0), 0)) As ʵ�ս��," & vbNewLine & _
            "        Max(Decode(Decode(a.��¼����,4,1,a.��¼����), 1, Decode(a.��¼״̬, 1, a.����id, 3, a.����id, 0), 0)) As ����id," & vbNewLine & _
            "        Sum(Nvl(a.����, 1) * a.����) As ʣ������" & vbNewLine & _
            " From ������ü�¼ A" & vbNewLine & _
            " Where a.��¼״̬ In (1, 2, 3) And a.����ID=[1]" & strWhere & vbNewLine & _
            "       And Nvl(a.���ӱ�־, 0) <> 9 And Nvl(a.����״̬, 0) <> 1" & vbNewLine & _
            " Group By Mod(a.��¼����,10),a.No"
        
    strSQL = "Select /*+ RULE */" & _
            "  Decode(a.��¼����,4,'�Һŵ�','�շѵ�') As ����, -1 As ѡ��, Decode(Nvl(Max(t.����), 0), 0, Null, '��') As ҽ��, " & _
            "  a.No As ���ݺ�, Max(b.����) As ��������, Max(a.������) As ������, Max(a.�����) As �����," & _
            "  Max(a.סԺ��) As סԺ��, Max(c.����) As ҽ�Ƹ��ʽ, Max(a.����) As ����, Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Min(a.�ѱ�) As �ѱ�," & _
            "  To_Char(Max(a.Ӧ�ս��), '99999990.00') As Ӧ�ս��, To_Char(Max(a.ʵ�ս��), '99999990.00') As ʵ�ս��, Max(a.������) As ������," & _
            "  Max(a.����Ա����) As ����Ա, To_Char(Max(a.�Ǽ�ʱ��), 'YYYY-MM-DD HH24:MI:SS') As �Ǽ�ʱ��, a.����id, " & _
            "  Max(Decode(a.��¼����,4,a.����ID,Nvl(m.�������, a.����id))) As �������id," & _
            "  Nvl(Max(t.����), 0) As ����id, Max(a.ʣ������) As ʣ������" & _
            " From (" & strTable & ") A, ����Ԥ����¼ M, ���ű� B, ҽ�Ƹ��ʽ C, ���ս����¼ T" & _
            " Where a.��������id = b.Id And a.���ʽ = c.����(+) And a.����id = t.��¼id(+) And t.����(+) = 1" & _
            "       And a.����id = m.����id(+) And (b.վ�� = '" & gstrNodeNo & "' Or b.վ�� Is Null)" & _
            "       And a.ʵ��Ʊ�� Is Null " & _
            "       And (a.��¼����=1 And (Nvl(t.����,0)<>0 Or Nvl(t.����,0)=0 And a.ʣ������<>0) " & _
            "            Or a.��¼����=4 And a.ʣ������<>0)" & _
            " Group By a.��¼����,a.No, a.����id" & _
            " Order By ����,����id Desc,���ݺ�"
            
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, dtStartDate, dtEndDate)
    '102113,��ͨ�����Լ���ҽ���ӿڴ�ӡ�ĵ���ȫ���˷ѵĲ���ʾ
    With vsList
        .Redraw = flexRDNone
        Set .DataSource = mrsList
        
        For lngCol = 0 To .COLS - 1
            .ColAlignment(lngCol) = flexAlignLeftCenter
            .FixedAlignment(lngCol) = flexAlignCenterCenter
            .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
            If .ColKey(lngCol) Like "*ID" Then
                .ColHidden(lngCol) = True
            ElseIf .ColKey(lngCol) = "ʣ������" Then
                .ColHidden(lngCol) = True
            ElseIf .ColKey(lngCol) Like "*��*" Or .ColKey(lngCol) Like "*��" Then
                .ColAlignment(lngCol) = flexAlignRightCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
        Call .AutoSize(0, .COLS - 1)
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "�����б�", False

        For i = 1 To .Rows - 1
            If i > .Rows - 1 Then Exit For
            lng����ID = Val(Trim(.TextMatrix(i, .ColIndex("����ID"))))

            If strPreNoType <> .TextMatrix(i, .ColIndex("����")) _
                Or strPreNo <> Trim(.TextMatrix(i, .ColIndex("���ݺ�"))) Then
                blnVirtualPrint = False: blnRemove = False
                
                strPreNoType = .TextMatrix(i, .ColIndex("����"))
                strPreNo = Trim(.TextMatrix(i, .ColIndex("���ݺ�")))
                intInsure = Val(Trim(.TextMatrix(i, .ColIndex("����Id"))))
                dblʣ������ = Val(Trim(.TextMatrix(i, .ColIndex("ʣ������"))))

                If intInsure <> 0 Then
                    blnVirtualPrint = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure, CStr(lng����ID))
                End If

                If blnVirtualPrint = False And RoundEx(dblʣ������, 6) = 0 Then
                    blnRemove = True
                    .RemoveItem i
                    i = i - 1
                End If
            ElseIf blnRemove Then
                .RemoveItem i
                i = i - 1
            End If

            If blnRemove = False Then
                lng������� = Val(Trim(.TextMatrix(i, .ColIndex("�������ID"))))
                byt�������� = IIf(.TextMatrix(i, .ColIndex("����")) = "�Һŵ�", 4, 1)
                If Not (byt�������� = 1 And InStr(1, str���ý��� & ",", "," & lng������� & ",") > 0 _
                    Or byt�������� = 4 And InStr(1, str�ҺŽ��� & ",", "," & lng����ID & ",") > 0) Then

                    If byt�������� = 1 Then
                        str���ý��� = str���ý��� & "," & lng�������
                    Else
                        str�ҺŽ��� = str�ҺŽ��� & "," & lng����ID
                    End If

                    '�����ָ���
                    If i > .FixedRows Then
                        .CellBorderRange i, .FixedCols, i, .COLS - 1, vbBlue, 0, 1, 0, 0, 0, 0
                    End If
                End If
            End If
        Next
        
        If chkRegistFee.Value = vbChecked Then
            '������ʾ
            .OutlineBar = flexOutlineBarComplete
            .Subtotal flexSTClear
            .Subtotal flexSTNone, .ColIndex("����"), , , &H8000000F
            .Outline .ColIndex("ѡ��")
            .OutlineCol = .ColIndex("ѡ��")
            For i = 1 To .Rows - 1
                .MergeRow(i) = False
                If .IsSubtotal(i) Then
                    .TextMatrix(i, .ColIndex("ѡ��")) = "-1"
                    For j = 0 To .COLS - 1
                        If j > .ColIndex("ѡ��") Then
                            .Cell(flexcpText, i, j) = .TextMatrix(i, 0)
                        End If
                    Next
                    .MergeRow(i) = True
                End If
            Next
            .MergeCells = flexMergeRestrictRows
        End If
        .ColHidden(.ColIndex("����")) = True
        
        .Editable = flexEDKbdMouse
        .Redraw = flexRDBuffered
        vsList_AfterRowColChange 0, 0, .Row, .Col
    End With
    If str���ý��� <> "" Then str���ý��� = Mid(str���ý���, 2)
    If str�ҺŽ��� <> "" Then str�ҺŽ��� = Mid(str�ҺŽ���, 2)
    
    If str���ý��� = "" And str�ҺŽ��� = "" Then
        vsDetail.Clear 1: vsDetail.Rows = 1
        vsBalance.Clear 1: vsBalance.Rows = 1
    End If
    
    '���ؽ��㷽ʽ
    Call InitBlanceData(str���ý���, str�ҺŽ���)
    Call SetBlanceShow
    Call InitPatiInsure
    
    zlCommFun.StopFlash
    
    Screen.MousePointer = 0
    ReadListData = True
    Exit Function
errHandle:
    vsList.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
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
    Call SetSelect(Row)
    Call SetBlanceShow
    '����ѡ�񵥾�ȷ��������
    Call InitPatiInsure
End Sub

Private Sub InitPatiInsure()
    '����ѡ�񵥾�ȷ����������
    Dim strNo As String, lngRow As Long
    
    mintInsure = 0
    With vsList
        For lngRow = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsList, lngRow, .ColIndex("ѡ��")) Then
                strNo = Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�")))
                If Val(.TextMatrix(lngRow, .ColIndex("����ID"))) <> 0 Then
                    mintInsure = Val(.TextMatrix(lngRow, .ColIndex("����ID")))
                    Exit For
                End If
            End If
        Next
    End With
    '���³�ʼ�����˷�Ʊ��Ϣ
    Call ReInitPatiInvoice
End Sub

Private Sub SetSelect(ByVal Row As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ���־
    '����:���˺�
    '����:2011-09-04 22:14:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    With vsList
        '73270,Ƚ����,2014-5-23,�����ѡ�����µĸ�ѡ�򣬱�������ʱ����13�����Ͳ�ƥ�䡱
        If Row < 0 Or .ColIndex("�������ID") < 0 Or .ColIndex("ѡ��") < 0 Then Exit Sub
        
        If .IsSubtotal(Row) Then
            For i = Row + 1 To .Rows - 1
                If .IsSubtotal(i) Then Exit For
                .TextMatrix(i, .ColIndex("ѡ��")) = .TextMatrix(Row, .ColIndex("ѡ��"))
            Next
        Else
            For i = Row - 1 To 1 Step -1
                If .IsSubtotal(i) Then Exit For
                If Not (.TextMatrix(i, .ColIndex("����")) = .TextMatrix(Row, .ColIndex("����")) _
                    And Val(.TextMatrix(i, .ColIndex("�������ID"))) = Val(.TextMatrix(Row, .ColIndex("�������ID")))) Then Exit For
                .TextMatrix(i, .ColIndex("ѡ��")) = .TextMatrix(Row, .ColIndex("ѡ��"))
            Next
            For i = Row + 1 To .Rows - 1
                If .IsSubtotal(i) Then Exit For
                If Not (.TextMatrix(i, .ColIndex("����")) = .TextMatrix(Row, .ColIndex("����")) _
                    And Val(.TextMatrix(i, .ColIndex("�������ID"))) = Val(.TextMatrix(Row, .ColIndex("�������ID")))) Then Exit For
                .TextMatrix(i, .ColIndex("ѡ��")) = .TextMatrix(Row, .ColIndex("ѡ��"))
            Next
        End If
    End With
End Sub

Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "�����б�", False
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strNo As String, byt��¼���� As Byte
    
    If OldRow = NewRow Then Exit Sub
    With vsList
        If NewRow < .FixedRows Then Exit Sub
        byt��¼���� = IIf(.TextMatrix(NewRow, .ColIndex("����")) = "�Һŵ�", 4, 1)
        strNo = Trim(.TextMatrix(NewRow, .ColIndex("���ݺ�")))
    End With
    ShowDetail byt��¼����, strNo
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
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModule, mstrUseType)
    mintPatiInvoiceFormat = zl_GetInvoicePrintFormat(mlngModule, mstrUseType, , True)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModule, mstrUseType)
    Call ShowBillFormat
    If blnFact Then Call RefreshFact
End Sub

Private Function ShowBillFormat()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ��¼���շѲ���Ա��ʾ����ʹ���շ�Ʊ�ݸ�ʽ
    '����:���˺�
    '����:2016-06-08 10:06:20
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strName As String, intFormat As Integer
    
    On Error GoTo errHandle
    If mblnStartFactUseType Then Exit Function
    
    If mbln���ֽ������ Then
        intFormat = mintPatiInvoiceFormat
    Else
        intFormat = mintInvoiceFormat
    End If
    Call ZlShowBillFormat(mlngModule, lblFormat, intFormat)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetNextFactNo(ByVal strUseType As String, ByRef lng����ID As Long, ByRef lngShareUseID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��һ�ŷ�Ʊ��
    '���:strUserType-ʹ�����
    '     lng����ID-����ID
    '     lngShareUseID-����ID
    '����:lng����ID-����ID
    '����:��һ�ŷ�Ʊ��
    '����:���˺�
    '����:2016-06-08 10:27:46
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If gblnStrictCtrl Then
        If zlCheckInvoiceValied(lng����ID, 1, , lngShareUseID, strUseType) = False Then Exit Function
        '�ϸ�ȡ��һ������
        GetNextFactNo = GetNextBill(lng����ID)
        Exit Function
    End If
    GetNextFactNo = zlStr.Increase(UCase(zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, mlngModule)))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RefreshFact()
    '���ܣ�ˢ���շ�Ʊ�ݺ�
    If mblnStartFactUseType Then Exit Sub
    
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
            Call zlCheckFactIsEnough
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
