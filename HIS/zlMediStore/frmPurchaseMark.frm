VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmPurchaseMark 
   Caption         =   "�˶Է�Ʊ"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   Icon            =   "frmPurchaseMark.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   9135
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4320
      ScaleHeight     =   255
      ScaleWidth      =   2895
      TabIndex        =   31
      Top             =   480
      Width           =   2895
      Begin VB.PictureBox picCol3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   34
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   33
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   32
         Top             =   0
         Width           =   260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "δ�޸�"
         Height          =   180
         Left            =   2280
         TabIndex        =   37
         Top             =   37
         Width           =   540
      End
      Begin VB.Label lblNotExecute 
         AutoSize        =   -1  'True
         Caption         =   "�Ѹ���"
         Height          =   180
         Left            =   360
         TabIndex        =   36
         Top             =   37
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "���޸�"
         Height          =   180
         Left            =   1320
         TabIndex        =   35
         Top             =   37
         Width           =   540
      End
   End
   Begin VB.PictureBox picDetails 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   3840
      ScaleHeight     =   1335
      ScaleWidth      =   3495
      TabIndex        =   15
      Top             =   1080
      Width           =   3495
      Begin XtremeSuiteControls.TabControl tbcDetails 
         Height          =   975
         Left            =   720
         TabIndex        =   16
         Top             =   240
         Width           =   1935
         _Version        =   589884
         _ExtentX        =   3413
         _ExtentY        =   1720
         _StockProps     =   64
      End
   End
   Begin VB.Frame fra���� 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3015
      Begin VB.CommandButton cmdҩƷ 
         Caption         =   "��"
         Height          =   300
         Left            =   2640
         TabIndex        =   30
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox txtҩƷ���� 
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   28
         Top             =   1182
         Width           =   1725
      End
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   300
         Width           =   1725
      End
      Begin VB.CommandButton Cmd��Ӧ�� 
         Caption         =   "��"
         Height          =   300
         Left            =   2640
         TabIndex        =   12
         Top             =   720
         Width           =   255
      End
      Begin VB.ComboBox cbo������� 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1623
         Width           =   1725
      End
      Begin VB.TextBox txt������Ʊ�� 
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   5
         Top             =   3420
         Width           =   1725
      End
      Begin VB.TextBox txt��ʼ��Ʊ�� 
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   4
         Top             =   2976
         Width           =   1725
      End
      Begin VB.TextBox txt��Ӧ�� 
         Height          =   300
         Left            =   1200
         TabIndex        =   1
         Top             =   741
         Width           =   1485
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   2064
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   57671683
         CurrentDate     =   40848
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   2520
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   57671683
         CurrentDate     =   40848
      End
      Begin VB.Label lblҩƷ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ҩƷ����"
         Height          =   180
         Left            =   180
         TabIndex        =   29
         Top             =   1242
         Width           =   720
      End
      Begin VB.Label lbl�ⷿ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�⹺�ⷿ"
         Height          =   180
         Left            =   180
         TabIndex        =   23
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl��ʼ���� 
         Caption         =   "��ʼ����"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblEnd��Ʊ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������Ʊ��"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label lblStart��Ʊ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��ʼ��Ʊ��"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   3030
         Width           =   900
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   2580
         Width           =   720
      End
      Begin VB.Label lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Left            =   180
         TabIndex        =   3
         Top             =   1710
         Width           =   720
      End
      Begin VB.Label lbl��Ӧ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�� Ӧ ��"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   780
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   6780
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPurchaseMark.frx":6852
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9763
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmPurchaseMark.frx":70E6
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmPurchaseMark.frx":75E8
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
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   3720
      ScaleHeight     =   4095
      ScaleWidth      =   5175
      TabIndex        =   17
      Top             =   2640
      Width           =   5175
      Begin VSFlex8Ctl.VSFlexGrid vsf��ͷ 
         Height          =   3855
         Left            =   2640
         TabIndex        =   27
         Top             =   0
         Width           =   2415
         _cx             =   4260
         _cy             =   6800
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
         AllowUserResizing=   0
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
      Begin VB.CheckBox chkȫѡ 
         Caption         =   "ȫѡ"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   1215
      End
      Begin VB.PictureBox picSplit 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   255
         ScaleWidth      =   4935
         TabIndex        =   21
         Top             =   2760
         Width           =   4935
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf��� 
         Height          =   855
         Left            =   240
         TabIndex        =   19
         Top             =   3120
         Width           =   4455
         _cx             =   7858
         _cy             =   1508
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
         AllowUserResizing=   0
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
      Begin VSFlex8Ctl.VSFlexGrid vsfδ��� 
         Height          =   975
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   5175
         _cx             =   9128
         _cy             =   1720
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
         AllowUserResizing=   0
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
         Begin VB.PictureBox picSetCols 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   220
            Index           =   0
            Left            =   0
            Picture         =   "frmPurchaseMark.frx":7AEA
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf�ѱ�� 
         Height          =   1215
         Left            =   0
         TabIndex        =   20
         Top             =   1440
         Width           =   5295
         _cx             =   9340
         _cy             =   2143
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
         AllowUserResizing=   0
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
         Begin VB.PictureBox picSetCols 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   220
            Index           =   1
            Left            =   0
            Picture         =   "frmPurchaseMark.frx":801C
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpPanel 
      Bindings        =   "frmPurchaseMark.frx":854E
      Left            =   720
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgPicture 
      Left            =   1320
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPurchaseMark.frx":8562
   End
End
Attribute VB_Name = "frmPurchaseMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const menuToolSave As Integer = 101
Private Const menuToolGetData As Integer = 102
Private Const menuToolExit As Integer = 103
Private Const menuTool��� As Integer = 104
Private Const menuTool��ϸ As Integer = 105
Private Const menuToolȫѡ As Integer = 106
Private Const menuToolSave2 As Integer = 107
Private mobjMnu As ICommandBarControl

Private Const CSTCOLOR_UNMODIFY = &HC0C0FF       '�ۺ� ѡ��ҳ��ɫ
Private Const CSTCOLOR_NORECORDS = &HFFFFFF   '

Private Const mColumn As String = "����|NO|ҩƷ����|����|���|��λ|����|�ɹ���|�ɹ����|��Ʊ��|��Ʊ���|��Ʊ����|�����|�������"
Private mintShowCostDigit As Integer            '�ɱ���С��λ��
Private mintShowPriceDigit As Integer           '�ۼ�С��λ��
Private mintShowNumberDigit As Integer          '����С��λ��
Private mintShowMoneyDigit As Integer           '���С��λ��
Private mintUnit As Integer '��λϵ��
Private mblnDo As Boolean
Private mstrSelColumn As String                 '��¼ѡ����Ҫ��ʾ����
Private mstrColumn As String
Private mvMsg As VbMsgBoxResult                 '��ʾ��Ϣ
Private mstrLike As String                      '�ǰ����ַ�ʽ��������ƥ��
Private mstrPrivs As String                     'ģ��Ȩ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4
Private mint��λϵ�� As Integer

Private mStr�ⷿ As String
Private mstr��ǰ�ⷿ As String
Private Const MStrCaption As String = "ҩƷ�⹺������"

Private Enum mPage
    δ��ǵ��� = 0
    �ѱ�ǵ��� = 1
End Enum

Private Enum mColumnMark
    ��� = 0
    id = 1
    ��ĿId
    �����־
    NO
    ҩƷ����
    ����
    ���
    ���㵥λ
    ����
    �ɹ���
    �ɹ����
'    ��ǰ���
'    ȫԺ���
    �������
    ��Ʊ��
    ��Ʊ���
    ��Ʊ����
    �����
    �������
    ����ϵ��
    �����װ
    סԺ��װ
    ҩ���װ
    count = 22
End Enum

Public Sub showMe(ByVal str�ⷿ As String, ByVal str��ǰ�ⷿ As String, ByVal objFrm As frmMainList, ByVal strPrivs As String)
    '�������̣�����������������ñ����壬��������Ӧ����
    
'    mlng�ⷿ = lng�ⷿ
'    mStr�ⷿ = str�ⷿ
    mStr�ⷿ = str�ⷿ
    mstr��ǰ�ⷿ = str��ǰ�ⷿ
    mstrPrivs = strPrivs
    Me.Show vbModal, objFrm
End Sub

Private Sub initComandbar()
    '��ʼ��������
    Dim cbrControlMain As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = Me.imgPicture.Icons
    
    '����������
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagFloating Or xtpFlagAlignAny
    
    With cbrToolBar.Controls    'menuToolSave2
        Set cbrControlMain = .Add(xtpControlButton, menuToolSave, "���")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, menuToolSave2, "ȡ�����")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, menuToolGetData, "��ȡ����")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, menuToolExit, "�˳�")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        cbrControlMain.BeginGroup = True
        
'        Set cbrControlMain = .Add(xtpControlButton, menuTool���, "���")
'            cbrControlMain.flags = xtpFlagRightAlign
'            cbrControlMain.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
'        Set cbrControlMain = .Add(xtpControlButton, menuTool��ϸ, "��ϸ")
'            cbrControlMain.flags = xtpFlagRightAlign
'            cbrControlMain.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
'            cbrControlMain.Checked = True
    End With
    
    cbsMain.Item(1).Delete
End Sub

Private Sub InitTabControl()
    '��ʼ��Tabcontrol�ؼ�
    With Me.tbcDetails
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(mPage.δ��ǵ���, "δ��ǵ���", picList.hWnd, 0).Tag = "δ��ǵ���_"
        .InsertItem(mPage.�ѱ�ǵ���, "�ѱ�ǵ���", picList.hWnd, 0).Tag = "�ѱ�ǵ���_"
        .Item(mPage.�ѱ�ǵ���).Selected = True
        .Item(mPage.δ��ǵ���).Selected = True
    End With
End Sub

Private Sub cbo�ⷿ_Click()
    Call SetSelectorRS(1, "ҩƷ�⹺������", cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    vsf��ͷ.Visible = False
    Select Case Control.id
        Case menuTool���   '����б�
'            Call Simple(Control)
        Case menuTool��ϸ   '��ϸ�б�
'            Call Full(Control)
        Case menuToolGetData   '��ȡ����
            Call checkUpdate
        Case menuToolSave, menuToolSave2  '����
            Call Save
        Case menuToolExit
            Call ExitForm
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    picDetails.Move fra����.Width, lngTop, Me.Width - fra����.Width - fra����.Left, lngBottom - staThis.Height - lngTop
    tbcDetails.Move 0, 0, picDetails.Width, picDetails.Height
    cbo�ⷿ.Move lbl�ⷿ.Left + lbl�ⷿ.Width + 280, lbl�ⷿ.Top, fra����.Width - cbo�ⷿ.Left - 100
    txt��Ӧ��.Move cbo�ⷿ.Left, txt��Ӧ��.Top, fra����.Width - cbo�ⷿ.Left - 100
    Cmd��Ӧ��.Left = txt��Ӧ��.Left + txt��Ӧ��.Width - Cmd��Ӧ��.Width
    txtҩƷ����.Move cbo�ⷿ.Left, txtҩƷ����.Top, fra����.Width - cbo�ⷿ.Left - 100
    cmdҩƷ.Left = txt��Ӧ��.Left + txt��Ӧ��.Width - Cmd��Ӧ��.Width
    cbo�������.Move cbo�ⷿ.Left, cbo�������.Top, txtҩƷ����.Width
    dtp��ʼʱ��.Move cbo�ⷿ.Left, dtp��ʼʱ��.Top, txtҩƷ����.Width
    dtp����ʱ��.Move cbo�ⷿ.Left, dtp����ʱ��.Top, txtҩƷ����.Width
    txt��ʼ��Ʊ��.Move cbo�ⷿ.Left, txt��ʼ��Ʊ��.Top, txtҩƷ����.Width
    txt������Ʊ��.Move cbo�ⷿ.Left, txt������Ʊ��.Top, txtҩƷ����.Width
    
    Call initOtherControl
End Sub

'Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    If Control.Id = menuTool��� Then
'        If Control.Checked Then
'            Control.IconId = 105
'        Else
'            Control.IconId = 104
'        End If
'    ElseIf Control.Id = menuTool��ϸ Then
'        If Control.Checked Then
'            Control.IconId = 105
'        Else
'            Control.IconId = 104
'        End If
'    End If
'End Sub

Private Sub cbo�������_Click()
    If cbo�������.Text = "�Զ�������" Then
        lbl��ʼ����.Visible = True
        dtp��ʼʱ��.Visible = True
        lbl��������.Visible = True
        dtp����ʱ��.Visible = True
        
        lblStart��Ʊ.Top = dtp����ʱ��.Top + dtp����ʱ��.Height + 130
        txt��ʼ��Ʊ��.Top = dtp����ʱ��.Top + dtp����ʱ��.Height + 80
        lblEnd��Ʊ.Top = txt��ʼ��Ʊ��.Top + txt��ʼ��Ʊ��.Height + 130
        txt������Ʊ��.Top = txt��ʼ��Ʊ��.Top + txt��ʼ��Ʊ��.Height + 80
    Else
        lbl��ʼ����.Visible = False
        dtp��ʼʱ��.Visible = False
        lbl��������.Visible = False
        dtp����ʱ��.Visible = False
        
        lblStart��Ʊ.Top = cbo�������.Top + cbo�������.Height + 130
        txt��ʼ��Ʊ��.Top = cbo�������.Top + cbo�������.Height + 80
        lblEnd��Ʊ.Top = txt��ʼ��Ʊ��.Top + txt��ʼ��Ʊ��.Height + 130
        txt������Ʊ��.Top = txt��ʼ��Ʊ��.Top + txt��ʼ��Ʊ��.Height + 80
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If tbcDetails.Item(0).Selected = True Then
'        cbsMain(1).Controls(1).Caption = "���"
        cbsMain.FindControl(xtpControlButton, menuToolSave).Caption = "���"
    Else
        cbsMain.FindControl(xtpControlButton, menuToolSave).Caption = "ȡ�����"
    End If
End Sub

Private Sub chkȫѡ_Click()
    Dim i As Integer
    If tbcDetails.Item(mPage.δ��ǵ���).Selected = True Then
        With vsfδ���
            For i = 1 To .rows - 1
                If chkȫѡ.Value = 1 Then
                    .TextMatrix(i, mColumnMark.�����־) = "��"
                    .Cell(flexcpFontBold, i, mColumnMark.�����־) = True
                    .Cell(flexcpFontSize, i, mColumnMark.�����־) = 10
                    .Cell(flexcpForeColor, i, mColumnMark.�����־) = vbBlue
                Else
                    .TextMatrix(i, mColumnMark.�����־) = ""
                End If
            Next
        End With
    ElseIf tbcDetails.Item(mPage.�ѱ�ǵ���).Selected = True Then
        With vsf�ѱ��
        For i = 1 To .rows - 1
            If Trim(.TextMatrix(i, mColumnMark.�������)) = "δ����" Then '�Ѿ�����ĵ��ݲ����޸ĸ����־
                If chkȫѡ.Value = 1 Then
                    .TextMatrix(i, mColumnMark.�����־) = "��"
                    .Cell(flexcpFontBold, i, mColumnMark.�����־) = True
                    .Cell(flexcpFontSize, i, mColumnMark.�����־) = 10
                    .Cell(flexcpForeColor, i, mColumnMark.�����־) = vbBlue
                Else
                    .TextMatrix(i, mColumnMark.�����־) = ""
                End If
            End If
        Next
    End With
    End If
End Sub

Private Sub Cmd��Ӧ��_Click()
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsRecord As ADODB.Recordset
    
    vRect = zlControl.GetControlRect(txt��Ӧ��.hWnd) '��ȡλ��
    
    gstrSQL = "select id,����,����,���� from ��Ӧ�� Where ĩ�� = 1 and (վ�� = '-' Or վ�� Is Null) And" & _
                " (Substr(����, 1, 1) = 1 Or Nvl(ĩ��, 0) = 0)"
                
    Set rsRecord = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "��Ӧ��", False, "", "", False, False, _
    True, vRect.Left, vRect.Top, txt��Ӧ��.Height, blnCancel, False, True)

    If rsRecord Is Nothing Then
        Exit Sub
    Else
        If txt��Ӧ��.Tag <> rsRecord!id Then
            txtҩƷ����.Tag = ""
            txtҩƷ����.Text = ""
            txt��ʼ��Ʊ��.Text = ""
            txt������Ʊ��.Text = ""
        End If
        txt��Ӧ��.Text = rsRecord!����
        txt��Ӧ��.Tag = rsRecord!id
    End If
    zlControl.TxtSelAll txt��Ӧ��
    OS.OpenIme True
End Sub

Private Sub CmdҩƷ_Click()
    Dim RecReturn As Recordset
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "ҩƷ�⹺������", cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    End If
    
    Set RecReturn = frmSelector.showMe(Me, 0, 1, , , , cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), , , , , , False, mstrPrivs)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        txtҩƷ����.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        txtҩƷ����.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    txtҩƷ����.Tag = RecReturn!ҩƷid
End Sub

Private Sub dkpPanel_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.id = 1 Then
         Item.Handle = fra����.hWnd
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 12000
    mblnDo = Val(zlDataBase.GetPara("ʹ�ø��Ի����")) <> 0
    mstrLike = IIf(Val(zlDataBase.GetPara("����ƥ��")) = 0, "%", "")
    staThis.Panels(2).Picture = picColor
    
    Call initComandbar  '��ʼ��������
    Call initPanel  '��ʼ�����
    Call InitTabControl
    Call initOtherControl   '���������ؼ�λ��
    Call initColumn '��ʼ����
    Call initComboBox
        
    If mblnDo Then
        RestoreWinState Me, App.ProductName, MStrCaption
    End If
    
    mstrColumn = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "������", "����ʾ����", "")
    If mstrColumn <> "" And mblnDo = True Then
        Call SetColumnVisible
    End If
    
    lbl��ʼ����.Visible = False
    dtp��ʼʱ��.Visible = False
    lbl��������.Visible = False
    dtp����ʱ��.Visible = False
    
    Me.Caption = "�˶Է�Ʊ"
    dtp��ʼʱ��.Value = DateAdd("d", -7, Sys.Currentdate)
    dtp����ʱ��.Value = Sys.Currentdate
    
    mintUnit = zlDataBase.GetPara("ҩƷ��λ", glngSys, 1300, 0, 0, True)
    Select Case mintUnit
        Case 4 '�ۼ۵�λ
            mint��λϵ�� = 4
        Case 2 '���ﵥλ
            mint��λϵ�� = 2
        Case 3 'סԺ��λ
            mint��λϵ�� = 3
        Case 1 'ҩ�ⵥλ
            mint��λϵ�� = 1
        Case 0
            mint��λϵ�� = 1
    End Select
    Call GetDrugDigit(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
End Sub

Private Sub SetColumnVisible()
    Dim i As Integer
    Dim strTemp As String
    Dim arrColumn As Variant
    Dim j As Integer
    
    ReDim arrColumn(UBound(Split(mstrColumn, "|"))) As String
    For i = 0 To UBound(arrColumn) - 1
        arrColumn(i) = Split(mstrColumn, "|")(i)
    Next
    With vsf��ͷ
        For i = 1 To .rows - 1
            For j = 0 To UBound(arrColumn) - 1
                If InStr(1, arrColumn(j), .TextMatrix(i, 2)) > 0 Then
                    .TextMatrix(i, 1) = IIf(Split(arrColumn(j), ",")(0) = "0", "", Split(arrColumn(j), ",")(0))
                End If
            Next
        Next
    End With
    
    For i = 1 To vsf��ͷ.rows - 1
        If vsf��ͷ.TextMatrix(i, 1) = "" Then
            vsfδ���.colHidden(vsfδ���.ColIndex(vsf��ͷ.TextMatrix(i, 2))) = True
            vsf�ѱ��.colHidden(vsfδ���.ColIndex(vsf��ͷ.TextMatrix(i, 2))) = True
        End If
    Next
End Sub

Private Sub initPanel()
    '��ʼ�������ؼ�
    'DockingPane
    '-----------------------------------------------------
    Dim objPaneCon As Pane
    Dim objPaneDetail As Pane
    
    Me.dkpPanel.SetCommandBars Me.cbsMain
    Me.dkpPanel.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpPanel.Options.ThemedFloatingFrames = True
    Me.dkpPanel.Options.AlphaDockingContext = True
    
    Set objPaneCon = Me.dkpPanel.CreatePane(1, 200, 0, DockLeftOf, Nothing)
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable
    objPaneCon.Title = "��ȡ����"
End Sub

Private Sub initOtherControl()
    '��ʼ�������ؼ� �����б�
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    picDetails.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - staThis.Height - lngTop
    tbcDetails.Move 0, 0, picDetails.Width, picDetails.Height
    
    If tbcDetails.Item(0).Selected = True Then
        vsf�ѱ��.Visible = False
        vsfδ���.Visible = True
        
        chkȫѡ.Move 0, 0, 1215, 255
        vsfδ���.Move 0, chkȫѡ.Height, picList.Width, (picList.Height / 6) * 5
        picSplit.Move 0, vsfδ���.Top + vsfδ���.Height, picList.ScaleWidth, 50
        vsf���.Move 0, picSplit.Top + picSplit.Height, picList.ScaleWidth, picList.Height - picSplit.Top - picSplit.Height
        
    Else
        vsf�ѱ��.Visible = True
        vsfδ���.Visible = False
        
        chkȫѡ.Move 0, 0, 1215, 255
        vsf�ѱ��.Move 0, chkȫѡ.Height, picList.ScaleWidth, (picList.ScaleHeight / 6) * 5
        picSplit.Move 0, vsf�ѱ��.Top + vsf�ѱ��.Height, picList.ScaleWidth, 50
        vsf���.Move 0, picSplit.Top + picSplit.Height, picSplit.ScaleWidth, picList.Height - picSplit.Top - picSplit.Height
    End If
End Sub

Private Sub initColumn()
    Dim i As Integer
    '��ʼ���������ͷ δ���
    With vsfδ���
        .rows = 1
        .Cols = mColumnMark.count
        .Editable = flexEDNone
        .MergeCells = flexMergeRestrictColumns
        
        .AllowSelection = False '���ܶ�ѡ
        .SelectionMode = flexSelectionByRow '����ѡ��
        
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False '���ܶ�ѡ��Ԫ��
        .RowHeight(0) = 310
    End With
    With vsf�ѱ��
        .rows = 1
        .Cols = mColumnMark.count
        
        .Editable = flexEDNone
        .MergeCells = flexMergeRestrictColumns
        
        .AllowSelection = False '���ܶ�ѡ
        .SelectionMode = flexSelectionByRow '����ѡ��
        
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False '���ܶ�ѡ��Ԫ��
        .RowHeight(0) = 310
    End With
    With vsf���
        .rows = 0
        .Editable = flexEDNone
        .GridLines = flexGridInset
        .GridColor = &H0&
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False '���ܶ�ѡ��Ԫ��
    End With
    With vsf��ͷ
        .Cols = 3
        .ColDataType(1) = flexDTBoolean
        .Editable = flexEDNone
        .MergeCells = flexMergeRestrictAll
        .MergeRow(0) = True
        .TextMatrix(0, 0) = "δѡ�е��н�����"
        .TextMatrix(0, 1) = "δѡ�е��н�����"
        .TextMatrix(0, 2) = "δѡ�е��н�����"
        .colHidden(0) = True
        .rows = UBound(Split(mColumn, "|")) + 1
        For i = 1 To .rows - 1
            .TextMatrix(i, 1) = "1"
            .TextMatrix(i, 2) = Split(mColumn, "|")(i)
            .RowHeight(i) = 300
        Next
        .Visible = False
    End With
    
    VsfGridColFormat vsfδ���, mColumnMark.���, "���", 640, flexAlignCenterCenter, "���"
    VsfGridColFormat vsfδ���, mColumnMark.id, "id", 600, flexAlignCenterCenter, "id"
    VsfGridColFormat vsfδ���, mColumnMark.��ĿId, "��Ŀid", 600, flexAlignCenterCenter, "��Ŀid"
    VsfGridColFormat vsfδ���, mColumnMark.�����־, "����", 1000, flexAlignCenterCenter, "����"
    VsfGridColFormat vsfδ���, mColumnMark.NO, "NO", 1500, flexAlignLeftCenter, "NO"
    VsfGridColFormat vsfδ���, mColumnMark.ҩƷ����, "ҩƷ����", 1500, flexAlignLeftCenter, "ҩƷ����"
    
    VsfGridColFormat vsfδ���, mColumnMark.����, "����", 600, flexAlignLeftCenter, "����"
    VsfGridColFormat vsfδ���, mColumnMark.���, "���", 1500, flexAlignLeftCenter, "���"
    
    VsfGridColFormat vsfδ���, mColumnMark.���㵥λ, "��λ", 600, flexAlignLeftCenter, "��λ"
    VsfGridColFormat vsfδ���, mColumnMark.����, "����", 1000, flexAlignRightCenter, "����"
    VsfGridColFormat vsfδ���, mColumnMark.�ɹ���, "�ɹ���", 1000, flexAlignRightCenter, "�ɹ���"
    VsfGridColFormat vsfδ���, mColumnMark.�ɹ����, "�ɹ����", 1000, flexAlignRightCenter, "�ɹ����"
    
'    VsfGridColFormat vsfδ���, mColumnMark.��ǰ���, "��ǰ���", 1000, flexAlignRightCenter, "��ǰ���"
'    VsfGridColFormat vsfδ���, mColumnMark.ȫԺ���, "ȫԺ���", 1500, flexAlignRightCenter, "ȫԺ���"
    VsfGridColFormat vsfδ���, mColumnMark.�������, "�������", 1000, flexAlignLeftCenter, "�������"
    VsfGridColFormat vsfδ���, mColumnMark.�����, "�����", 1000, flexAlignLeftCenter, "�����"
    VsfGridColFormat vsfδ���, mColumnMark.�������, "�������", 1000, flexAlignLeftCenter, "�������"
    VsfGridColFormat vsfδ���, mColumnMark.��Ʊ��, "��Ʊ��", 1000, flexAlignLeftCenter, "��Ʊ��"
    VsfGridColFormat vsfδ���, mColumnMark.��Ʊ���, "��Ʊ���", 1000, flexAlignRightCenter, "��Ʊ���"
    VsfGridColFormat vsfδ���, mColumnMark.��Ʊ����, "��Ʊ����", 1000, flexAlignLeftCenter, "��Ʊ����"
    VsfGridColFormat vsfδ���, mColumnMark.����ϵ��, "����ϵ��", 1000, flexAlignLeftCenter, "����ϵ��"
    VsfGridColFormat vsfδ���, mColumnMark.�����װ, "�����װ", 1000, flexAlignLeftCenter, "�����װ"
    VsfGridColFormat vsfδ���, mColumnMark.סԺ��װ, "סԺ��װ", 1000, flexAlignRightCenter, "סԺ��װ"
    VsfGridColFormat vsfδ���, mColumnMark.ҩ���װ, "ҩ���װ", 1000, flexAlignLeftCenter, "ҩ���װ"
    vsfδ���.Cell(flexcpPicture, vsfδ���.Row, 0, vsfδ���.Row, 0) = picSetCols(0)
    
    '�ѱ��
    VsfGridColFormat vsf�ѱ��, mColumnMark.���, "���", 640, flexAlignCenterCenter, "���"
    VsfGridColFormat vsf�ѱ��, mColumnMark.id, "id", 600, flexAlignCenterCenter, "id"
    VsfGridColFormat vsf�ѱ��, mColumnMark.��ĿId, "��Ŀid", 600, flexAlignCenterCenter, "��Ŀid"
    VsfGridColFormat vsf�ѱ��, mColumnMark.�����־, "ȡ������", 1000, flexAlignCenterCenter, "ȡ������"
    VsfGridColFormat vsf�ѱ��, mColumnMark.NO, "NO", 1500, flexAlignLeftCenter, "NO"
    VsfGridColFormat vsf�ѱ��, mColumnMark.ҩƷ����, "ҩƷ����", 1500, flexAlignLeftCenter, "ҩƷ����"
    VsfGridColFormat vsf�ѱ��, mColumnMark.����, "����", 600, flexAlignLeftCenter, "����"
    VsfGridColFormat vsf�ѱ��, mColumnMark.���, "���", 1500, flexAlignLeftCenter, "���"
    VsfGridColFormat vsf�ѱ��, mColumnMark.���㵥λ, "��λ", 1000, flexAlignLeftCenter, "��λ"
    VsfGridColFormat vsf�ѱ��, mColumnMark.����, "����", 1000, flexAlignRightCenter, "����"
    VsfGridColFormat vsf�ѱ��, mColumnMark.�ɹ���, "�ɹ���", 1000, flexAlignRightCenter, "�ɹ���"
    VsfGridColFormat vsf�ѱ��, mColumnMark.�ɹ����, "�ɹ����", 1000, flexAlignRightCenter, "�ɹ����"
'    VsfGridColFormat vsf�ѱ��, mColumnMark.��ǰ���, "��ǰ���", 1000, flexAlignRightCenter, "��ǰ���"
'    VsfGridColFormat vsf�ѱ��, mColumnMark.ȫԺ���, "ȫԺ���", 1500, flexAlignRightCenter, "ȫԺ���"
    VsfGridColFormat vsf�ѱ��, mColumnMark.�������, "�������", 1000, flexAlignLeftCenter, "�������"
    VsfGridColFormat vsf�ѱ��, mColumnMark.�����, "�����", 1000, flexAlignLeftCenter, "�����"
    VsfGridColFormat vsf�ѱ��, mColumnMark.�������, "�������", 1000, flexAlignLeftCenter, "�������"
    VsfGridColFormat vsf�ѱ��, mColumnMark.��Ʊ��, "��Ʊ��", 1000, flexAlignLeftCenter, "��Ʊ��"
    VsfGridColFormat vsf�ѱ��, mColumnMark.��Ʊ���, "��Ʊ���", 1000, flexAlignRightCenter, "��Ʊ���"
    VsfGridColFormat vsf�ѱ��, mColumnMark.��Ʊ����, "��Ʊ����", 1000, flexAlignLeftCenter, "��Ʊ����"
    VsfGridColFormat vsf�ѱ��, mColumnMark.����ϵ��, "����ϵ��", 1000, flexAlignLeftCenter, "����ϵ��"
    VsfGridColFormat vsf�ѱ��, mColumnMark.�����װ, "�����װ", 1000, flexAlignLeftCenter, "�����װ"
    VsfGridColFormat vsf�ѱ��, mColumnMark.סԺ��װ, "סԺ��װ", 1000, flexAlignRightCenter, "סԺ��װ"
    VsfGridColFormat vsf�ѱ��, mColumnMark.ҩ���װ, "ҩ���װ", 1000, flexAlignLeftCenter, "ҩ���װ"
    vsf�ѱ��.Cell(flexcpPicture, vsf�ѱ��.Row, 0, vsf�ѱ��.Row, 0) = picSetCols(1)
    
    vsfδ���.colHidden(mColumnMark.id) = True
    vsf�ѱ��.colHidden(mColumnMark.id) = True
    vsfδ���.colHidden(mColumnMark.��ĿId) = True
    vsf�ѱ��.colHidden(mColumnMark.��ĿId) = True
    vsfδ���.colHidden(mColumnMark.�������) = True
    vsf�ѱ��.colHidden(mColumnMark.�������) = True
    
    vsfδ���.colHidden(mColumnMark.����ϵ��) = True
    vsfδ���.colHidden(mColumnMark.�����װ) = True
    vsfδ���.colHidden(mColumnMark.סԺ��װ) = True
    vsfδ���.colHidden(mColumnMark.ҩ���װ) = True
    
    vsf�ѱ��.colHidden(mColumnMark.����ϵ��) = True
    vsf�ѱ��.colHidden(mColumnMark.�����װ) = True
    vsf�ѱ��.colHidden(mColumnMark.סԺ��װ) = True
    vsf�ѱ��.colHidden(mColumnMark.ҩ���װ) = True
End Sub

Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf�����ã��������п��ж��뷽ʽ���̶��ж��뷽ʽ��Ĭ��Ϊ���ж��룩
    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub

Private Sub initComboBox()
    Dim i As Integer
    Dim strTemp As String
    Dim strIndex As String
    Dim arrtemp As Variant
    
    With cbo�������
        .Clear
        .AddItem "����", "0"
        .AddItem "һ������", "1"
        .AddItem "һ������", "2"
        .AddItem "��������", "3"
        .AddItem "�Զ�������", "4"
        .ListIndex = 0
    End With
    
    ReDim arrtemp(UBound(Split(mStr�ⷿ, "|"))) As String
    
    With cbo�ⷿ
        .Clear
        For i = 0 To UBound(arrtemp) - 1
            strIndex = ""
            strTemp = ""
            arrtemp(i) = Split(mStr�ⷿ, "|")(i)
            strIndex = Mid(arrtemp(i), 1, InStr(1, arrtemp(i), ",") - 1)
            strTemp = Mid(arrtemp(i), InStr(1, arrtemp(i), ",") + 1)
            .AddItem strTemp
            .ItemData(.NewIndex) = strIndex
        Next
        .ListIndex = mstr��ǰ�ⷿ
    End With
End Sub

'Private Sub Simple(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    '��ģʽ
'    If Control.Checked = False Then
'        Control.Checked = True
'        cbsMain.Item(1).Controls.Item(5).Checked = False
'    End If
'
'    If cbsMain.Item(1).Controls.Item(4).Checked = True Then '��ģʽ��ѡ�еĻ�
'        With vsfδ���
'            .ColHidden(mColumnMark.��Ʊ��) = True
'            .ColHidden(mColumnMark.��Ʊ���) = True
'            .ColHidden(mColumnMark.��Ʊ����) = True
'        End With
'        With vsf�ѱ��
'            .ColHidden(mColumnMark.��Ʊ��) = True
'            .ColHidden(mColumnMark.��Ʊ���) = True
'            .ColHidden(mColumnMark.��Ʊ����) = True
'        End With
'    End If
'End Sub

'Private Sub Full(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    '����ģʽ
'    If Control.Checked = False Then
'        Control.Checked = True
'        cbsMain.Item(1).Controls.Item(4).Checked = False
'    End If
'
'    If cbsMain.Item(1).Controls.Item(5).Checked = True Then '������ģʽ��ѡ�еĻ�
'        With vsfδ���
'            .ColHidden(mColumnMark.��Ʊ��) = False
'            .ColHidden(mColumnMark.��Ʊ���) = False
'            .ColHidden(mColumnMark.��Ʊ����) = False
'        End With
'
'        With vsf�ѱ��
'            .ColHidden(mColumnMark.��Ʊ��) = False
'            .ColHidden(mColumnMark.��Ʊ���) = False
'            .ColHidden(mColumnMark.��Ʊ����) = False
'        End With
'    End If
'End Sub

Private Sub checkUpdate()
    '����Ƿ��޸��˼�¼
    Dim i As Integer
    Dim blnChange As Boolean
    Dim lngResult As Long
    
    blnChange = False
    For i = 1 To vsfδ���.rows - 1
        If vsfδ���.Cell(flexcpForeColor, i, mColumnMark.�����־) = vbBlue Then
            blnChange = True
        End If
    Next
    
    For i = 1 To vsf�ѱ��.rows - 1
        If vsf�ѱ��.Cell(flexcpForeColor, i, mColumnMark.�����־) = vbBlue Then
            blnChange = True
        End If
    Next
    
    If blnChange = True Then
        lngResult = MsgBox("�������ݱ��޸��ˣ��Ƿ������", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName)
    End If
    
    If lngResult = vbYes Or blnChange = False Then
        Call GetData
    End If
End Sub

Private Sub GetData()
    Dim rsRecord As ADODB.Recordset
    Dim i As Integer
    Dim blnChange As Boolean
    Dim lngResult As Long
    Dim dbDate As Date
    
    On Error GoTo errHandle
'    blnChange = False
'    For i = 1 To vsfδ���.rows - 1
'        If vsfδ���.Cell(flexcpForeColor, i, mColumnMark.�����־) = vbBlue Then
'            blnChange = True
'        End If
'    Next
'
'    For i = 1 To vsf�ѱ��.rows - 1
'        If vsf�ѱ��.Cell(flexcpForeColor, i, mColumnMark.�����־) = vbBlue Then
'            blnChange = True
'        End If
'    Next
'
'    If blnChange = True Then
'        lngResult = MsgBox("�������ݱ��޸��ˣ��Ƿ񱣴棿", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName)
'    End If
'
'    If lngResult = vbYes Then
'        Call Save
'    End If
        
    '��ȡ���ݵķ���
    If Trim(txt��Ӧ��.Text) = "" And Trim(txt��ʼ��Ʊ��.Text) = "" And Trim(txt������Ʊ��.Text) = "" And Trim(cbo�������.Text) = "" Then
        Exit Sub
    End If
    If txt��Ӧ��.Text = "" Then
        txt��Ӧ��.Tag = ""
    End If
    gstrSQL = "Select distinct a.Id, a.��ⵥ�ݺ� no, a.��Ŀid, Decode(a.�����־, Null, 0, 0, 0, 1) �����־, a.����, a.Ʒ�� As ҩƷ����, a.���, a.����, a.������λ, a.�ɹ���, a.�ɹ����," & _
              "     A.������� , A.�����, A.�������, A.��Ʊ��, A.��Ʊ���, A.��Ʊ����, b.����ϵ��, b.�����װ, b.סԺ��װ, b.ҩ���װ,b.���ﵥλ,b.סԺ��λ,b.ҩ�ⵥλ " & _
              "  From Ӧ����¼ A, ҩƷ��� B where a.��Ŀid = b.ҩƷid And a.����� is not null and a.�ⷿid=[1] and a.��Ʊ�� is not null and  a.��Ʊ���� is not null and a.��¼״̬=1 and a.��¼����=0 "
    
    
    If txt��Ӧ��.Tag <> "" Then
        gstrSQL = gstrSQL & " and a.��λid=[2]"
    End If
        
    If Me.txt��ʼ��Ʊ�� <> "" And Me.txt������Ʊ�� <> "" Then gstrSQL = gstrSQL & " And a.��Ʊ�� >= [3] And a.��Ʊ�� <=[4] "
    If Me.txt��ʼ��Ʊ�� <> "" And Me.txt������Ʊ�� = "" Then gstrSQL = gstrSQL & " And a.��Ʊ�� >= [3] "
    If Me.txt��ʼ��Ʊ�� = "" And Me.txt������Ʊ�� <> "" Then gstrSQL = gstrSQL & " And a.��Ʊ�� <= [3] "
    
    If cbo�������.Text = "����" Then
        dbDate = CDate(Format(Date, "yyyy-mm-dd") & " 00:00:00")
        gstrSQL = gstrSQL & " and a.������� between [5] and sysdate"
    End If
    If cbo�������.Text = "һ������" Then
        dbDate = CDate(Format(DateAdd("d", -7, Date), "yyyy-mm-dd") & " 00:00:00")
        gstrSQL = gstrSQL & " and a.������� between [5] and sysdate"
    End If
    If cbo�������.Text = "һ������" Then
        dbDate = CDate(Format(DateAdd("d", -30, Date), "yyyy-mm-dd") & " 00:00:00")
        gstrSQL = gstrSQL & " and a.������� between [5] and sysdate"
    End If
    If cbo�������.Text = "��������" Then
        dbDate = CDate(Format(DateAdd("d", -90, Date), "yyyy-mm-dd") & " 00:00:00")
        gstrSQL = gstrSQL & " and a.������� between [5] and sysdate"
    End If
    If cbo�������.Text = "�Զ�������" Then
        dbDate = CDate(Format(dtp��ʼʱ��, "yyyy-mm-dd") & " 00:00:00")
        gstrSQL = gstrSQL & " and a.������� between [5] and [6]"
    End If
    
    If txtҩƷ����.Tag <> "" Then
        gstrSQL = gstrSQL & " and a.��Ŀid=[7]"
    End If
    
    Set rsRecord = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ����", cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), txt��Ӧ��.Tag, UCase(txt��ʼ��Ʊ��.Text), UCase(txt������Ʊ��.Text), dbDate, CDate(Format(dtp����ʱ��, "yyyy-mm-dd") & " 23:59:59"), txtҩƷ����.Tag)
    If rsRecord Is Nothing Then
        vsfδ���.rows = 1
        vsf�ѱ��.rows = 1
        Exit Sub
    End If
    
    If vsfδ���.rows > 1 Then
        With vsfδ���
            .Cell(flexcpFontBold, 1, 0, .rows - 1, .Cols - 1) = False
            .Cell(flexcpForeColor, 1, 0, .rows - 1, .Cols - 1) = vbBlack
        End With
    End If
    If vsf�ѱ��.rows > 1 Then
        With vsf�ѱ��
            .Cell(flexcpFontBold, 1, 0, .rows - 1, .Cols - 1) = False
            .Cell(flexcpForeColor, 1, 0, .rows - 1, .Cols - 1) = vbBlack
        End With
    End If
    
    Call SetColumn(rsRecord)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get��ǰ���(ByVal lngҩƷID As Long) As String
    '���ܣ�ĳ��ҩƷ��ĳ�����ҵĵ�ǰ���
    '����ֵ�����ز�ѯ���Ŀ����
    '������ҩƷid
    Dim rsRecord As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Sum(a.ʵ������) ��ǰ���,b.���㵥λ  From ҩƷ��� a,�շ���ĿĿ¼ b where a.ҩƷid=b.id and a.ҩƷid = [1] And a.�ⷿid = [2] group by b.���㵥λ"
    Set rsRecord = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ���", lngҩƷID, cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    If rsRecord Is Nothing Or rsRecord.RecordCount = 0 Then
        Get��ǰ��� = "0"
        Exit Function
    Else
        Get��ǰ��� = IIf(IsNull(rsRecord!��ǰ���), "0", rsRecord!��ǰ���) & IIf(IsNull(rsRecord!���㵥λ), "", rsRecord!���㵥λ)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetȫԺ���(ByVal lngҩƷID As Long) As String
    '���ܣ���ѯĳ��ҩƷ��ȫԺ�Ŀ��
    '����ֵ�����ز�ѯ���Ŀ����
    '������ҩƷid
    Dim rsRecord As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Sum(a.ʵ������) ȫԺ���,b.���㵥λ  From ҩƷ��� a,�շ���ĿĿ¼ b where a.ҩƷid=b.id  and a.ҩƷid=[1] group by b.���㵥λ"
    Set rsRecord = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ���", lngҩƷID)
    If rsRecord Is Nothing Or rsRecord.RecordCount = 0 Then
        GetȫԺ��� = "0"
        Exit Function
    Else
        GetȫԺ��� = IIf(IsNull(rsRecord!ȫԺ���), "0", rsRecord!ȫԺ���) & IIf(IsNull(rsRecord!���㵥λ), "", rsRecord!���㵥λ)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetColumn(ByVal rsRecord As ADODB.Recordset)
    '��������ֵ
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    Set rsTemp = rsRecord
    
    rsRecord.Filter = "�����־=0"
    With vsfδ���
        .rows = rsRecord.RecordCount + 1
        For i = 1 To rsRecord.RecordCount
            .TextMatrix(i, mColumnMark.���) = i
            .TextMatrix(i, mColumnMark.����) = IIf(IsNull(rsRecord!����), "", rsRecord!����)
            .TextMatrix(i, mColumnMark.id) = rsRecord!id
            .TextMatrix(i, mColumnMark.��ĿId) = rsRecord!��ĿId
            
            .TextMatrix(i, mColumnMark.NO) = rsRecord!NO
            If rsRecord!�����־ = 0 Then
                .TextMatrix(i, mColumnMark.�����־) = ""
            End If
            
            .TextMatrix(i, mColumnMark.ҩƷ����) = rsRecord!ҩƷ����
            .TextMatrix(i, mColumnMark.���) = rsRecord!���
            
            Select Case mint��λϵ��
                Case 4  '�ۼ۵�λ
                    .TextMatrix(i, mColumnMark.���㵥λ) = IIf(IsNull(rsRecord!������λ), "", rsRecord!������λ)
                Case 1  'ҩ�ⵥλ
                    .TextMatrix(i, mColumnMark.���㵥λ) = IIf(IsNull(rsRecord!ҩ�ⵥλ), "", rsRecord!ҩ�ⵥλ)
                Case 2  '���ﵥλ
                    .TextMatrix(i, mColumnMark.���㵥λ) = IIf(IsNull(rsRecord!���ﵥλ), "", rsRecord!���ﵥλ)
                Case 3  'סԺ��λ
                    .TextMatrix(i, mColumnMark.���㵥λ) = IIf(IsNull(rsRecord!סԺ��λ), "", rsRecord!סԺ��λ)
            End Select
            Select Case mint��λϵ��
                Case 4  '�ۼ۵�λ
                    .TextMatrix(i, mColumnMark.����) = zlStr.FormatEx(IIf(IsNull(rsRecord!����), 0, rsRecord!����), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.�ɹ���) = zlStr.FormatEx(rsRecord!�ɹ���, mintShowCostDigit, , True)
                Case 1  'ҩ�ⵥλ
                    .TextMatrix(i, mColumnMark.����) = zlStr.FormatEx(IIf(IsNull(rsRecord!����), 0, rsRecord!���� / IIf(IsNull(rsRecord!ҩ���װ), 1, rsRecord!ҩ���װ)), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.�ɹ���) = zlStr.FormatEx(rsRecord!�ɹ��� * IIf(IsNull(rsRecord!ҩ���װ), 1, rsRecord!ҩ���װ), mintShowCostDigit, , True)
                Case 2  '���ﵥλ
                    .TextMatrix(i, mColumnMark.����) = zlStr.FormatEx(IIf(IsNull(rsRecord!����), 0, rsRecord!���� / IIf(IsNull(rsRecord!�����װ), 1, rsRecord!�����װ)), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.�ɹ���) = zlStr.FormatEx(rsRecord!�ɹ��� * IIf(IsNull(rsRecord!�����װ), 1, rsRecord!�����װ), mintShowCostDigit, , True)
                Case 3  'סԺ��λ
                    .TextMatrix(i, mColumnMark.����) = zlStr.FormatEx(IIf(IsNull(rsRecord!����), 0, rsRecord!���� / IIf(IsNull(rsRecord!סԺ��װ), 1, rsRecord!סԺ��װ)), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.�ɹ���) = zlStr.FormatEx(rsRecord!�ɹ��� * IIf(IsNull(rsRecord!סԺ��װ), 1, rsRecord!סԺ��װ), mintShowCostDigit, , True)
            End Select
            
            .TextMatrix(i, mColumnMark.�ɹ����) = zlStr.FormatEx(rsRecord!�ɹ����, mintShowMoneyDigit, , True)
'            .TextMatrix(i, mColumnMark.��ǰ���) = Get��ǰ���(rsRecord!��ĿId)
'            .TextMatrix(i, mColumnMark.ȫԺ���) = GetȫԺ���(rsRecord!��ĿId)
            
            If IsNull(rsRecord!�������) Or rsRecord!������� = 0 Then
                .TextMatrix(i, mColumnMark.�������) = "δ����"
            Else
                .TextMatrix(i, mColumnMark.�������) = "�Ѹ���"
            End If
            .TextMatrix(i, mColumnMark.�����) = rsRecord!�����
            .TextMatrix(i, mColumnMark.�������) = Format(rsRecord!�������, "yyyy-mm-dd")
            .TextMatrix(i, mColumnMark.��Ʊ��) = IIf(IsNull(rsRecord!��Ʊ��), "", rsRecord!��Ʊ��)
            .TextMatrix(i, mColumnMark.��Ʊ���) = zlStr.FormatEx(IIf(IsNull(rsRecord!��Ʊ���), "", rsRecord!��Ʊ���), mintShowMoneyDigit, , True)
            .TextMatrix(i, mColumnMark.��Ʊ����) = IIf(IsNull(rsRecord!��Ʊ����), "", Format(rsRecord!��Ʊ����, "yyyy-mm-dd"))
            .TextMatrix(i, mColumnMark.����ϵ��) = rsRecord!����ϵ��
            .TextMatrix(i, mColumnMark.�����װ) = rsRecord!�����װ
            .TextMatrix(i, mColumnMark.סԺ��װ) = rsRecord!סԺ��װ
            .TextMatrix(i, mColumnMark.ҩ���װ) = rsRecord!ҩ���װ
            .RowHeight(i) = 310
            If .TextMatrix(i, mColumnMark.�������) = "�Ѹ���" Then
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
            End If
            rsRecord.MoveNext
        Next
    End With
    
    rsTemp.Filter = "�����־=1"
    With vsf�ѱ��
        .rows = rsRecord.RecordCount + 1
        For i = 1 To rsRecord.RecordCount
            .TextMatrix(i, mColumnMark.���) = i
            .TextMatrix(i, mColumnMark.����) = IIf(IsNull(rsRecord!����), "", rsRecord!����)
            .TextMatrix(i, mColumnMark.id) = rsRecord!id
            .TextMatrix(i, mColumnMark.��ĿId) = rsRecord!��ĿId
            
            .TextMatrix(i, mColumnMark.NO) = rsRecord!NO
            If rsRecord!�����־ = 1 Then
                .TextMatrix(i, mColumnMark.�����־) = ""
            End If
            
            .TextMatrix(i, mColumnMark.ҩƷ����) = rsRecord!ҩƷ����
            .TextMatrix(i, mColumnMark.���) = rsRecord!���
            
            Select Case mint��λϵ��
                Case 4  '�ۼ۵�λ
                    .TextMatrix(i, mColumnMark.���㵥λ) = IIf(IsNull(rsRecord!������λ), "", rsRecord!������λ)
                Case 1  'ҩ�ⵥλ
                    .TextMatrix(i, mColumnMark.���㵥λ) = IIf(IsNull(rsRecord!ҩ�ⵥλ), "", rsRecord!ҩ�ⵥλ)
                Case 2  '���ﵥλ
                    .TextMatrix(i, mColumnMark.���㵥λ) = IIf(IsNull(rsRecord!���ﵥλ), "", rsRecord!���ﵥλ)
                Case 3  'סԺ��λ
                    .TextMatrix(i, mColumnMark.���㵥λ) = IIf(IsNull(rsRecord!סԺ��λ), "", rsRecord!סԺ��λ)
            End Select
            Select Case mint��λϵ��
                Case 4  '�ۼ۵�λ
                    .TextMatrix(i, mColumnMark.����) = zlStr.FormatEx(IIf(IsNull(rsRecord!����), 0, rsRecord!����), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.�ɹ���) = zlStr.FormatEx(rsRecord!�ɹ���, mintShowCostDigit, , True)
                Case 1  'ҩ�ⵥλ
                    .TextMatrix(i, mColumnMark.����) = zlStr.FormatEx(IIf(IsNull(rsRecord!����), 0, rsRecord!���� / IIf(IsNull(rsRecord!ҩ���װ), 1, rsRecord!ҩ���װ)), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.�ɹ���) = zlStr.FormatEx(rsRecord!�ɹ��� * IIf(IsNull(rsRecord!ҩ���װ), 1, rsRecord!ҩ���װ), mintShowCostDigit, , True)
                Case 2  '���ﵥλ
                    .TextMatrix(i, mColumnMark.����) = zlStr.FormatEx(IIf(IsNull(rsRecord!����), 0, rsRecord!���� / IIf(IsNull(rsRecord!�����װ), 1, rsRecord!�����װ)), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.�ɹ���) = zlStr.FormatEx(rsRecord!�ɹ��� * IIf(IsNull(rsRecord!�����װ), 1, rsRecord!�����װ), mintShowCostDigit, , True)
                Case 3  'סԺ��λ
                    .TextMatrix(i, mColumnMark.����) = zlStr.FormatEx(IIf(IsNull(rsRecord!����), 0, rsRecord!���� / IIf(IsNull(rsRecord!סԺ��װ), 1, rsRecord!סԺ��װ)), mintShowNumberDigit, , True)
                    .TextMatrix(i, mColumnMark.�ɹ���) = zlStr.FormatEx(rsRecord!�ɹ��� * IIf(IsNull(rsRecord!סԺ��װ), 1, rsRecord!סԺ��װ), mintShowCostDigit, , True)
            End Select
            
            .TextMatrix(i, mColumnMark.�ɹ����) = zlStr.FormatEx(rsRecord!�ɹ����, mintShowMoneyDigit, , True)
'            .TextMatrix(i, mColumnMark.��ǰ���) = Get��ǰ���(rsRecord!��ĿId)
'            .TextMatrix(i, mColumnMark.ȫԺ���) = GetȫԺ���(rsRecord!��ĿId)
            
            If IsNull(rsRecord!�������) Or rsRecord!������� = 0 Then
                .TextMatrix(i, mColumnMark.�������) = "δ����"
            Else
                .TextMatrix(i, mColumnMark.�������) = "�Ѹ���"
            End If
            .TextMatrix(i, mColumnMark.�����) = rsRecord!�����
            .TextMatrix(i, mColumnMark.�������) = Format(rsRecord!�������, "yyyy-mm-dd")
            .TextMatrix(i, mColumnMark.��Ʊ��) = IIf(IsNull(rsRecord!��Ʊ��), "", rsRecord!��Ʊ��)
            .TextMatrix(i, mColumnMark.��Ʊ���) = zlStr.FormatEx(IIf(IsNull(rsRecord!��Ʊ���), "", rsRecord!��Ʊ���), mintShowMoneyDigit, , True)
            .TextMatrix(i, mColumnMark.��Ʊ����) = IIf(IsNull(rsRecord!��Ʊ����), "", Format(rsRecord!��Ʊ����, "yyyy-mm-dd"))
            .TextMatrix(i, mColumnMark.����ϵ��) = rsRecord!����ϵ��
            .TextMatrix(i, mColumnMark.�����װ) = rsRecord!�����װ
            .TextMatrix(i, mColumnMark.סԺ��װ) = rsRecord!סԺ��װ
            .TextMatrix(i, mColumnMark.ҩ���װ) = rsRecord!ҩ���װ
            .RowHeight(i) = 310
            If .TextMatrix(i, mColumnMark.�������) = "�Ѹ���" Then
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
            End If
            rsRecord.MoveNext
        Next
    End With
    If vsfδ���.rows > 1 Then
        vsfδ���.Select 1, 1
    End If
End Sub

Private Sub Form_Resize()
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - staThis.Panels(3).Width - staThis.Panels(4).Width - staThis.Panels(5).Width - staThis.Panels(6).Width - .Width - 500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    Dim strTemp As String
    
    SaveWinState Me, App.ProductName, MStrCaption
    mblnDo = False
    mvMsg = vbYes
    With vsf��ͷ
        For i = 1 To .rows - 1
            strTemp = strTemp & IIf(.TextMatrix(i, 1) = "", 0, .TextMatrix(i, 1)) & "," & .TextMatrix(i, 2) & "|"
        Next
    End With
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\������", "����ʾ����", strTemp)
    Call ReleaseSelectorRS
End Sub

Private Sub picSetCols_Click(Index As Integer)
    With vsf��ͷ
        .Top = vsfδ���.Top + .CellHeight
        .Left = vsfδ���.Left + 10
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    If tbcDetails.Item(0).Selected = True Then
        If Button = 1 And vsfδ���.Height + y > 200 And picSplit.Top + y < staThis.Top - 1500 Then
            Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
            vsfδ���.Move vsfδ���.Left, chkȫѡ.Height, lngRight - lngLeft, vsfδ���.Height + y
            picSplit.Move vsfδ���.Left, picSplit.Top + y, lngRight - lngLeft
            vsf���.Move vsfδ���.Left, picSplit.Top + picSplit.Height, lngRight - lngLeft, picList.Height - y
        End If
    Else
        If Button = 1 And vsf�ѱ��.Height + y > 200 And picSplit.Top + y < staThis.Top - 1500 Then
            Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
            vsf�ѱ��.Move vsf�ѱ��.Left, chkȫѡ.Height, lngRight - lngLeft, vsf�ѱ��.Height + y
            picSplit.Move vsf�ѱ��.Left, picSplit.Top + y, lngRight - lngLeft
            vsf���.Move vsf�ѱ��.Left, picSplit.Top + picSplit.Height, lngRight - lngLeft, picList.Height - y
        End If
    End If
End Sub

Private Sub tbcDetails_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim cbrToolBar As CommandBar
    
    If tbcDetails.Item(0).Selected = True Then
        cbsMain.FindControl(xtpControlButton, menuToolSave).Visible = True
        cbsMain.FindControl(xtpControlButton, menuToolSave2).Visible = False
        vsf�ѱ��.Visible = False
        vsfδ���.Visible = True
        chkȫѡ.Enabled = True
        vsfδ���.Move 0, chkȫѡ.Height, picList.Width, (picList.Height / 6) * 5
        picSplit.Move 0, vsfδ���.Top + vsfδ���.Height, picList.ScaleWidth, 50
        vsf���.Move 0, picSplit.Top + picSplit.Height, picList.ScaleWidth, picList.Height - picSplit.Top - picSplit.Height
    Else
        cbsMain.FindControl(xtpControlButton, menuToolSave).Visible = False
        cbsMain.FindControl(xtpControlButton, menuToolSave2).Visible = True
        vsf�ѱ��.Visible = True
        vsfδ���.Visible = False
        chkȫѡ.Enabled = False
        vsf�ѱ��.Move 0, chkȫѡ.Height, picList.ScaleWidth, (picList.ScaleHeight / 6) * 5
        picSplit.Move 0, vsf�ѱ��.Top + vsf�ѱ��.Height, picList.ScaleWidth, 50
        vsf���.Move 0, picSplit.Top + picSplit.Height, picSplit.ScaleWidth, picList.Height - picSplit.Top - picSplit.Height
    End If
    Call setTabControlColor(tbcDetails)
End Sub

Private Sub txt��Ӧ��_GotFocus()
    zlControl.TxtSelAll txt��Ӧ��
    OS.OpenIme True
End Sub

Private Sub txt��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsRecord As ADODB.Recordset
    
    If KeyCode = 13 Then
        vRect = zlControl.GetControlRect(txt��Ӧ��.hWnd) '��ȡλ��
        
        gstrSQL = "select id,����,����,���� from ��Ӧ�� where ĩ�� = 1 and (���� like [1] OR ���� like [1] OR ���� like [1])"
        Set rsRecord = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "��Ӧ��", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, txt��Ӧ��.Height, blnCancel, False, True, UCase(txt��Ӧ��.Text) & "%")
        
        If blnCancel Then txt��Ӧ��.SetFocus: Exit Sub
        
        If rsRecord Is Nothing Then
            MsgBox "û��������Ĺ�Ӧ�̣������䣡", vbOKOnly + vbInformation, gstrSysName
            txt��Ӧ��.SelStart = 0
            txt��Ӧ��.SelLength = Len(txt��Ӧ��)
            Exit Sub
        Else
            If txt��Ӧ��.Tag <> rsRecord!id Then
                txtҩƷ����.Tag = ""
                txtҩƷ����.Text = ""
                txt��ʼ��Ʊ��.Text = ""
                txt������Ʊ��.Text = ""
            End If
            txt��Ӧ��.Text = rsRecord!����
            txt��Ӧ��.Tag = rsRecord!id
        End If
    End If
End Sub

Private Sub txt��Ӧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then
        txt��Ӧ��.Text = ""
        txt��Ӧ��.Tag = ""
    End If
End Sub

Private Sub txtҩƷ����_Change()
    If txtҩƷ����.Text = "" Then
         txtҩƷ����.Tag = ""
    End If
End Sub

Private Sub txtҩƷ����_GotFocus()
    zlControl.TxtSelAll txtҩƷ����
End Sub

Private Sub txtҩƷ����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RecReturn As Recordset
    Dim strkey As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtҩƷ����.Text) = "" Then Exit Sub
    vRect = zlControl.GetControlRect(txtҩƷ����.hWnd) '��ȡλ��
    dblLeft = vRect.Left
    dblTop = vRect.Top + txtҩƷ����.Height
    
    strkey = Trim(txtҩƷ����.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "ҩƷ�⹺������", cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    End If
    
    Set RecReturn = frmSelector.showMe(Me, 1, 1, strkey, dblLeft, dblTop, cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), , , , , , False, mstrPrivs)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        txtҩƷ����.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        txtҩƷ����.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    txtҩƷ����.Tag = RecReturn!ҩƷid
End Sub

Private Sub txtҩƷ����_Validate(Cancel As Boolean)
    If Trim(txtҩƷ����.Text) <> "" Then
        Call txtҩƷ����_KeyDown(vbKeyReturn, 1)
    End If
End Sub

Private Sub vsf��ͷ_LostFocus()
    vsf��ͷ.Visible = False
End Sub

Private Sub vsf��ͷ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With vsf��ͷ
        If .Row = 0 Then Exit Sub
        If .Col = 1 And Button = 1 Then
            If .TextMatrix(.Row, .Col) = "1" Then
                .TextMatrix(.Row, .Col) = ""
'                If tbcDetails.Item(0).Selected = True Then
                    vsfδ���.colHidden(vsfδ���.ColIndex(.TextMatrix(.Row, 2))) = True
'                Else
                    vsf�ѱ��.colHidden(vsf�ѱ��.ColIndex(.TextMatrix(.Row, 2))) = True
'                End If
            ElseIf .TextMatrix(.Row, .Col) = "" Then
                .TextMatrix(.Row, .Col) = "1"
'                If tbcDetails.Item(0).Selected = True Then
                    vsfδ���.colHidden(vsfδ���.ColIndex(.TextMatrix(.Row, 2))) = False
'                Else
                    vsf�ѱ��.colHidden(vsf�ѱ��.ColIndex(.TextMatrix(.Row, 2))) = False
'                End If
            End If
        End If
    End With
End Sub

Private Sub vsfδ���_DblClick()
    With vsfδ���
        vsf��ͷ.Visible = False
        If .Row = 0 Then
            Exit Sub
        End If
        If .TextMatrix(.Row, mColumnMark.�����־) = "��" Then
            .TextMatrix(.Row, mColumnMark.�����־) = ""
            .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = False
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbBlack
        Else
            .TextMatrix(.Row, mColumnMark.�����־) = "��"
            .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbBlue
        End If
    End With
End Sub

Private Sub vsfδ���_EnterCell()
    Dim rsTemp As ADODB.Recordset
    Dim rs��λ As ADODB.Recordset
    Dim i As Integer
    Dim intסԺϵ�� As Integer
    Dim int����ϵ�� As Integer
    Dim intҩ��ϵ�� As Integer
    
    On Error GoTo errHandle
    With vsfδ���
        If .rows = 1 Then
            vsf���.rows = 0
            Exit Sub
        End If
        If .TextMatrix(.Row, mColumnMark.��ĿId) <> "" And .Row <> 0 Then
            gstrSQL = "Select ����, Sum(ʵ������) ����" & _
                      "  From (Select ʵ������, b.���� " & _
                              " From ҩƷ��� A, ���ű� B, (Select Distinct ִ�п���id From �շ�ִ�п��� Where �շ�ϸĿid = [1]) D" & _
                              " Where a.�ⷿid = b.Id And a.�ⷿid = d.ִ�п���id And a.ҩƷid = [1]" & _
                              " Union All " & _
                              " Select 0 ����, a.����" & _
                              " From ���ű� A, (Select Distinct ִ�п���id From �շ�ִ�п��� Where �շ�ϸĿid = [1]) B" & _
                              " Where a.Id = b.ִ�п���id)" & _
                      "  Group By ���� "
           Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, .TextMatrix(.Row, mColumnMark.��ĿId))
           If Not rsTemp Is Nothing Then
                With vsf���
                    .Cols = rsTemp.RecordCount + 1
                    .rows = 2
                    For i = 1 To rsTemp.RecordCount
                        .TextMatrix(0, 0) = "�ⷿ"
                        .TextMatrix(1, 0) = "����"
                        VsfGridColFormat vsf���, i, rsTemp!����, 1500, flexAlignCenterCenter, rsTemp!����
                        
                        intסԺϵ�� = Val(vsfδ���.TextMatrix(vsfδ���.Row, mColumnMark.סԺ��װ))
                        int����ϵ�� = Val(vsfδ���.TextMatrix(vsfδ���.Row, mColumnMark.�����װ))
                        intҩ��ϵ�� = Val(vsfδ���.TextMatrix(vsfδ���.Row, mColumnMark.ҩ���װ))
                        Select Case mint��λϵ��
                            Case 4  '�ۼ۵�λ
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!����), 0, rsTemp!����), mintShowNumberDigit, , True)
                            Case 1  'ҩ�ⵥλ
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!����), 0, rsTemp!���� / intҩ��ϵ��), mintShowNumberDigit, , True)
                            Case 2  '���ﵥλ
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!����), 0, rsTemp!���� / int����ϵ��), mintShowNumberDigit, , True)
                            Case 3  'סԺ��λ
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!����), 0, rsTemp!���� / intסԺϵ��), mintShowNumberDigit, , True)
                        End Select
                        rsTemp.MoveNext
                    Next
                    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
                    .Cell(flexcpAlignment, 1, 1, 1, .Cols - 1) = flexAlignRightCenter
                    .ColAlignment(0) = flexAlignCenterCenter
                    .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = &H8000000F
                    .RowHeight(0) = 300
                    .RowHeight(1) = 300
                End With
           End If
        End If
        
        If .Row = 0 Then
            vsf���.Clear
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsf�ѱ��_DblClick()
    With vsf�ѱ��
        vsf��ͷ.Visible = False
        If .Row = 0 Then
            Exit Sub
        End If
        If Trim(.TextMatrix(.Row, mColumnMark.�������)) = "δ����" Then '�Ѿ�����ĵ��ݲ����޸ĸ����־
            If .TextMatrix(.Row, mColumnMark.�����־) = "��" Then
                .TextMatrix(.Row, mColumnMark.�����־) = ""
                .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = False
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbBlack
            Else
                .TextMatrix(.Row, mColumnMark.�����־) = "��"
                .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbBlue
            End If
        End If
    End With
End Sub

'Private Sub vsfδ���_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    With vsfδ���
'        vsf��ͷ.Visible = False
'        If .Row = 0 Then
'            Exit Sub
'        End If
'        If y < .CellHeight * .rows Then
'            If Button = 1 Then
'                If .Col = mColumnMark.�����־ Then
'                    If .TextMatrix(.Row, .Col) = "��" Then
'                        .TextMatrix(.Row, .Col) = ""
'                    Else
'                        .TextMatrix(.Row, .Col) = "��"
'                        .Cell(flexcpFontBold, .Row, .Col) = True
'                        .Cell(flexcpFontSize, .Row, .Col) = 10
'                        .Cell(flexcpForeColor, .Row, .Col) = vbBlue
'                    End If
'                End If
'            End If
'        End If
'    End With
'End Sub


Private Sub vsf�ѱ��_EnterCell()
    Dim rsTemp As ADODB.Recordset
    Dim rs��λ As ADODB.Recordset
    Dim i As Integer
    Dim intסԺϵ�� As Integer
    Dim int����ϵ�� As Integer
    Dim intҩ��ϵ�� As Integer
    
    On Error GoTo errHandle
    With vsf�ѱ��
        If .rows = 1 Then
            vsf���.rows = 0
            Exit Sub
        End If
        If .TextMatrix(.Row, mColumnMark.��ĿId) <> "" And .Row <> 0 Then
            gstrSQL = "Select ����, Sum(ʵ������) ����" & _
                      "  From (Select ʵ������, b.���� " & _
                              " From ҩƷ��� A, ���ű� B, (Select Distinct ִ�п���id From �շ�ִ�п��� Where �շ�ϸĿid = [1]) D" & _
                              " Where a.�ⷿid = b.Id And a.�ⷿid = d.ִ�п���id And a.ҩƷid = [1]" & _
                              " Union All " & _
                              " Select 0 ����, a.����" & _
                              " From ���ű� A, (Select Distinct ִ�п���id From �շ�ִ�п��� Where �շ�ϸĿid = [1]) B" & _
                              " Where a.Id = b.ִ�п���id)" & _
                      "  Group By ���� "
           Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, .TextMatrix(.Row, mColumnMark.��ĿId))
           If Not rsTemp Is Nothing Then
                With vsf���
                    .Cols = rsTemp.RecordCount + 1
                    .rows = 2
                    For i = 1 To rsTemp.RecordCount
                        .TextMatrix(0, 0) = "�ⷿ"
                        .TextMatrix(1, 0) = "����"
                        VsfGridColFormat vsf���, i, rsTemp!����, 1500, flexAlignCenterCenter, rsTemp!����
                        
                        intסԺϵ�� = Val(vsf�ѱ��.TextMatrix(vsf�ѱ��.Row, mColumnMark.סԺ��װ))
                        int����ϵ�� = Val(vsf�ѱ��.TextMatrix(vsf�ѱ��.Row, mColumnMark.�����װ))
                        intҩ��ϵ�� = Val(vsf�ѱ��.TextMatrix(vsf�ѱ��.Row, mColumnMark.ҩ���װ))
                        Select Case mint��λϵ��
                            Case 4  '�ۼ۵�λ
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!����), 0, rsTemp!����), mintShowNumberDigit, , True)
                            Case 1  'ҩ�ⵥλ
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!����), 0, rsTemp!���� / intҩ��ϵ��), mintShowNumberDigit, , True)
                            Case 2  '���ﵥλ
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!����), 0, rsTemp!���� / int����ϵ��), mintShowNumberDigit, , True)
                            Case 3  'סԺ��λ
                                .TextMatrix(1, i) = zlStr.FormatEx(IIf(IsNull(rsTemp!����), 0, rsTemp!���� / intסԺϵ��), mintShowNumberDigit, , True)
                        End Select
                        .ColAlignment(i) = flexAlignRightCenter
                        .ColWidth(i) = 1500
                        rsTemp.MoveNext
                    Next
                    .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = &H8000000F
                End With
           End If
        End If
        
        If .Row = 0 Then
            vsf���.Clear
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'Private Sub vsf�ѱ��_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim blnMsg As Boolean
'
'    With vsf�ѱ��
'        vsf��ͷ.Visible = False
'        If .Row = 0 Then
'            Exit Sub
'        End If
'        If y < .CellHeight * .rows Then
'            If Button = 1 Then
'                If .Col = mColumnMark.�����־ Then
'                    If Trim(.TextMatrix(.Row, mColumnMark.�������)) = "δ����" Then '�Ѿ�����ĵ��ݲ����޸ĸ����־
'                        If .TextMatrix(.Row, .Col) = "��" Then
''                            If mvMsg <> vbCancel And mvMsg <> vbIgnore Then
''                                mvMsg = frmMsgBox.ShowMsgBox("�õ����Ѿ���Ǹ��ȷ��Ҫȡ����ǣ�", Me)
''                                blnMsg = True
''                            End If
''                            If (mvMsg = vbYes Or mvMsg = vbIgnore) And blnMsg = True Then
'                                .TextMatrix(.Row, .Col) = ""
''                            End If
''                            If (mvMsg = vbCancel Or mvMsg = vbIgnore) And blnMsg = False Then
''                                .TextMatrix(.Row, .Col) = ""
''                            End If
'                        Else
'                            .TextMatrix(.Row, .Col) = "��"
'                            .Cell(flexcpFontBold, .Row, .Col) = True
'                            .Cell(flexcpFontSize, .Row, .Col) = 10
'                            .Cell(flexcpForeColor, .Row, .Col) = vbBlue
'                        End If
'                    End If
'                End If
'            End If
'        End If
'    End With
'End Sub



Private Sub ExitForm()
    '�˳�����ķ���
'    Dim i As Integer
'    Dim blnChange As Boolean
'    Dim lngResult As Long
'
'    blnChange = False
'    For i = 1 To vsfδ���.rows - 1
'        If vsfδ���.Cell(flexcpForeColor, i, mColumnMark.�����־) = vbBlue Then
'            blnChange = True
'        End If
'    Next
'
'    For i = 1 To vsf�ѱ��.rows - 1
'        If vsf�ѱ��.Cell(flexcpForeColor, i, mColumnMark.�����־) = vbBlue Then
'            blnChange = True
'        End If
'    Next
'
'    If blnChange = True Then
'        lngResult = MsgBox("�������ݱ��޸��ˣ��Ƿ񱣴棿", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName)
'
'        If lngResult = vbNo Then    '�˳����� ������
'            Unload Me
'        Else    '�˳����壬����
'            Call Save
'            Unload Me
'        End If
'    Else
'        Unload Me
'    End If
    Unload Me
End Sub

Private Sub Save()
    '���淽��
    Dim i As Integer
    Dim intTemp As Integer
    Dim strTemp As String
    Dim blnContinue As Boolean
    
    blnContinue = False
    If tbcDetails.Item(mPage.δ��ǵ���).Selected = True Then
        If MsgBox("����ǵ��ݣ��Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
'            strTemp = "��ǳɹ�,"
            With vsfδ���
                For i = 1 To .rows - 1
                    If .TextMatrix(i, mColumnMark.�����־) = "��" Then
                        intTemp = 1
                    Else
                        intTemp = 0
                    End If
                    gstrSQL = "zl_Ӧ����¼_�����־(" & .TextMatrix(i, mColumnMark.id) & ","
                    gstrSQL = gstrSQL & intTemp & ")"
                    
                    zlDataBase.ExecuteProcedure gstrSQL, MStrCaption
                Next
                    
                For i = 1 To .rows - 1
                    .Cell(flexcpFontBold, i, mColumnMark.�����־) = False
                    .Cell(flexcpFontSize, i, mColumnMark.�����־) = 9
                    .Cell(flexcpForeColor, i, mColumnMark.�����־) = vbBlack
                Next
            End With
            blnContinue = True
        End If
    End If
    
    If tbcDetails.Item(mPage.�ѱ�ǵ���).Selected = True Then
        If MsgBox("��ȡ�����ݱ�ǣ��Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
'            strTemp = "ȡ����ǳɹ�,"
            With vsf�ѱ��
                For i = 1 To .rows - 1
                    If .TextMatrix(i, mColumnMark.�����־) = "��" Then
                        intTemp = 0
                    Else
                        intTemp = 1
                    End If
                    gstrSQL = "zl_Ӧ����¼_�����־(" & .TextMatrix(i, mColumnMark.id) & ","
                    gstrSQL = gstrSQL & intTemp & ")"
                    
                    zlDataBase.ExecuteProcedure gstrSQL, MStrCaption
                Next
                
                For i = 1 To .rows - 1
                    .Cell(flexcpFontBold, i, mColumnMark.�����־) = False
                    .Cell(flexcpFontSize, i, mColumnMark.�����־) = 9
                    .Cell(flexcpForeColor, i, mColumnMark.�����־) = vbBlack
                Next
            End With
            blnContinue = True
        End If
    End If
'    If MsgBox(strTemp & "�Ƿ����������", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
'        txt��Ӧ��.Text = ""
'        txtҩƷ����.Tag = ""
'        txtҩƷ����.Text = ""
'        cbo�������.ListIndex = 0
'        txt��ʼ��Ʊ��.Text = ""
'        txt������Ʊ��.Text = ""
'    End If
    If blnContinue = True Then
        Call GetData
    End If
End Sub

Private Sub setTabControlColor(ByVal objtbc As TabControl)
    '��Tabcontrol�ؼ�������ɫ�ж�
    Dim i As Integer
    
    With objtbc
        For i = 0 To .ItemCount - 1
            If .Item(i).Selected = True Then
                .Item(i).Color = CSTCOLOR_UNMODIFY
            Else
                .Item(i).Color = CSTCOLOR_NORECORDS
            End If
        Next
    End With
End Sub





