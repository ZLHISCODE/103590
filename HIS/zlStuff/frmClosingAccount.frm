VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmClosingAccount 
   Caption         =   "���Ľ�����"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10965
   Icon            =   "frmClosingAccount.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picIni 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1800
      Picture         =   "frmClosingAccount.frx":030A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picShowDetail 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   4200
      ScaleHeight     =   4695
      ScaleWidth      =   9015
      TabIndex        =   8
      Top             =   1920
      Width           =   9015
      Begin VB.CommandButton cmdҩƷ 
         Height          =   300
         Left            =   3600
         Picture         =   "frmClosingAccount.frx":0894
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   270
      End
      Begin VB.ComboBox cbo��λ 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   60
         Width           =   1395
      End
      Begin VB.TextBox txt��ϸ���� 
         Height          =   300
         Left            =   600
         TabIndex        =   9
         Top             =   60
         Width           =   3000
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMistake 
         Height          =   765
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   3600
         Visible         =   0   'False
         Width           =   2895
         _cx             =   5106
         _cy             =   1349
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
         BackColorSel    =   16764622
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   275
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClosingAccount.frx":13CE
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
      Begin VSFlex8Ctl.VSFlexGrid vsfStuff 
         Height          =   765
         Left            =   0
         TabIndex        =   13
         Top             =   2040
         Width           =   4455
         _cx             =   7858
         _cy             =   1349
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
         BackColorSel    =   16764622
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   20
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   275
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClosingAccount.frx":14DB
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   765
         Left            =   0
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   5295
         _cx             =   9340
         _cy             =   1349
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
         BackColorSel    =   16764622
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   275
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClosingAccount.frx":1809
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
      Begin VSFlex8Ctl.VSFlexGrid vsfMistake 
         Height          =   765
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   3000
         Width           =   2895
         _cx             =   5106
         _cy             =   1349
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
         BackColorSel    =   16764622
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   275
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClosingAccount.frx":18EF
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
      Begin VB.Label lbl��ϸ���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lbl��λ 
         AutoSize        =   -1  'True
         Caption         =   "��λ"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   5040
         TabIndex        =   16
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   3015
      TabIndex        =   3
      Top             =   480
      Width           =   3015
      Begin VB.PictureBox picList 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   2055
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
         Begin VSFlex8Ctl.VSFlexGrid vsfList 
            Height          =   1005
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   4815
            _cx             =   8493
            _cy             =   1773
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
            BackColorSel    =   16764622
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
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   275
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmClosingAccount.frx":1986
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
            Begin VB.Image Image1 
               Height          =   15
               Left            =   1080
               Top             =   240
               Width           =   135
            End
         End
      End
      Begin VB.ComboBox cbo�ⷿ 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   600
         TabIndex        =   4
         Text            =   "cbo�ⷿ"
         Top             =   60
         Width           =   1935
      End
      Begin VB.Label lbl�ⷿ 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "�ⷿ"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   4200
      Width           =   2415
      Begin VB.Frame fraLine 
         Height          =   2085
         Left            =   120
         TabIndex        =   1
         Top             =   -120
         Width           =   45
      End
      Begin XtremeSuiteControls.TabControl tbcDetail 
         Height          =   975
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1720
         _StockProps     =   64
         Enabled         =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   7560
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14261
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   720
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Bindings        =   "frmClosingAccount.frx":1B8B
      Left            =   1200
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmClosingAccount.frx":1B9F
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   240
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmClosingAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Ĭ�ϵĴ����С
Private Const mcstlngWinNormalWidth As Long = 12000
Private Const mcstlngWinNormalHeight As Long = 8000

Private mrsAccount As ADODB.Recordset         '���ڻ������¼
Private mrsDetail As ADODB.Recordset
Private mrsMistake As ADODB.Recordset

Private mint��淽ʽ As Integer         '��淽ʽ -1-�ֹ���� >=0�Զ����
Private mint���ʱ�� As Integer         '���ʱ��
Private mlng���ID As Long
Private mstr��ǰ���� As String          '��ǰϵͳ����


Private mblnStart As Boolean

Private Const mconPane_Dept_Condition = 1                     '������

''�༭�˵�
'���Ľ��
Private Const mconMenu_Edit_CA_AddIniAccount = 3300            '��ʼ���
Private Const mconMenu_Edit_CA_AddNewAccount = 3301             '��������¼
Private Const mconMenu_Edit_CA_VerifyAccount = 3302             '��˽���¼
Private Const mconMenu_Edit_CA_CancelAccount = 3303             'ȡ������¼
Private Const mconMenu_Edit_CA_VerifyMistake = 3304             '��˽�����
Private Const mconMenu_Edit_CA_DeleteAccount = 3305             'ɾ������¼

Private Const mconMenu_CA_Refresh = 7001                        'ˢ��

'�����ϸ�б�
Private Const mconTab_CA_Detail = 0                 '�����ϸ�б�
Private Const mconTab_CA_Stuff = 1                '������ϸ�б�
Private Const mconTab_CA_Mistake = 2                '�������б�
Private Const mstrCaption As String = "���Ľ�����"

Private Type Type_Privs
    bln���пⷿ As Boolean
    bln��ʼ�� As Boolean
    bln��� As Boolean
End Type
Private mPrives As Type_Privs

Private mstrPrivs As String


Private Sub cbo��λ_Click()
    Dim intIndex As Integer
    
    If mblnStart = False Then Exit Sub
    With vsfList
        If .Rows = 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("�����")) = "" Then
            intIndex = 0
        Else
            intIndex = 1
        End If
    End With
    
    With cbo��λ
        If Val(.Tag) <> .ListIndex Then
            .Tag = .ListIndex
            If tbcDetail.Selected.Index = mconTab_CA_Detail Then
                LoadInOutList intIndex, mlng���ID
            ElseIf tbcDetail.Selected.Index = mconTab_CA_Stuff Then
                LoadDetailList intIndex, mlng���ID
            ElseIf tbcDetail.Selected.Index = mconTab_CA_Mistake Then
                LoadMistakeList intIndex, mlng���ID
            End If
        End If
    End With
End Sub


Private Sub cbo�ⷿ_Click()
    Dim lng�ⷿID As Long
    Dim Str�ڼ� As String
    Dim strSQL As String
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim rsTemp As ADODB.Recordset
    
    If Val(cbo�ⷿ.Tag) = cbo�ⷿ.ListIndex Then Exit Sub 'ѡ��Ļ����ϴοⷿ����ˢ��
    cbo�ⷿ.Tag = cbo�ⷿ.ListIndex
    
    If mblnStart = True Then
        Call CheckClosAccount
    End If
    
    If mblnStart = False Then Exit Sub
    
    Call GetAccountRecord
    Call RefreshList
    
End Sub

Private Sub cbo�ⷿ_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cbo�ⷿ.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfList): Exit Sub
    
    If cbo�ⷿ.ListIndex >= 0 Then
        If Val(cbo�ⷿ.Tag) = cbo�ⷿ.ListIndex Then
            Call zlControl.ControlSetFocus(vsfList, True)
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cbo�ⷿ, Trim(cbo�ⷿ.Text), "V,W") = False Then
        cbo�ⷿ.ListIndex = Val(cbo�ⷿ.Tag) 'δ�ҵ�����ԭ�ϴοⷿ
        Exit Sub
    End If
    If cbo�ⷿ.ListIndex >= 0 Then
        cbo�ⷿ.Tag = cbo�ⷿ.ListIndex
    End If

End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intIndex As Integer
    
    With vsfList
        If .TextMatrix(.Row, .ColIndex("�����")) = "" Then
            intIndex = 0
        Else
            intIndex = 1
        End If
    End With
    
    Select Case Control.Id
        ''''��ӡ
        Case conMenu_File_PrintSet
            '��ӡ����
            zlPrintSet
        Case conMenu_File_Preview
            '��ӡԤ��
            subPrint intIndex, tbcDetail.Selected.Index, 2
        Case conMenu_File_Print
            '��ӡ
            subPrint intIndex, tbcDetail.Selected.Index, 1
        Case conMenu_File_Excel
            '�����Excel
            subPrint intIndex, tbcDetail.Selected.Index, 3

        ''''����
        Case mconMenu_Edit_CA_VerifyAccount
            '��˽��
            Call MediAccountProcess_VerifyAccount
        Case mconMenu_Edit_CA_AddIniAccount
            '��ʼ���/��ʼ��
            Call MediAccountProcess_AddIniAccount(0)
        Case mconMenu_Edit_CA_AddNewAccount
            '���
            Call MediAccountProcess_AddIniAccount(1)
        Case mconMenu_Edit_CA_DeleteAccount
            'ɾ�����
            Call MediAccountProcess_DeleteAccount
        Case mconMenu_Edit_CA_CancelAccount
            'ȡ�����
            Call MediAccountProcess_CancleAccount
        ''''�鿴
        Case mconMenu_CA_Refresh 'ˢ��
            Call GetAccountRecord
            Call RefreshList
        ''''����
        Case conMenu_Help_Help                         '����
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
        Case conMenu_Help_Web                          'WEB�ϵ�����
        Case conMenu_Help_Web_Home                     '������ҳ
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum                    '������̳
            Call zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail                     '���ͷ���
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About                        '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)

        Case conMenu_File_Exit
            '�˳�
            Unload Me
        Case Else
            If Control.Id > 401 And Control.Id < 499 Then
                'ִ���Զ��屨��
                Call BillPrint_Custom(Control)
            End If
    End Select
End Sub

Private Sub BillPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '��ӡ�Զ��屨��

    'Ĭ�ϲ���������=����id��ҩ��=ҩ��id�����ID=���ID
    Dim strName As String
    Dim intType As Integer
    Dim lng���id As Long
    Dim lng�ⷿID As Long

    strName = Split(Control.Parameter, ",")(1)

    If strName = "ZL" & glngSys \ 100 & "_INSIDE_1332" Then
        Call ReportOpen(gcnOracle, glngSys, strName, Me)
    Else
        If vsfList.Row <> 0 Then
            lng���id = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("���ID")))
            lng�ⷿID = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("�ⷿID")))
        End If

        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strName, Me, _
            "���id=" & lng���id, _
            "�ⷿid=" & lng�ⷿID, _
            "����=" & IIf(Val(txt��ϸ����.Tag) = 0, "", Val(txt��ϸ����.Tag)))
    End If
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub


Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
      
    Me.picDetail.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub



Private Sub CmdҩƷ_Click()
    Dim intIndex As Integer
    
    With vsfList
        If .Rows = 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("�����")) = "" Then
            intIndex = 0
        Else
            intIndex = 1
        End If
    End With
    
    GetSelect ""

    DoEvents

    If tbcDetail.Selected.Index = mconTab_CA_Detail Then
        LoadInOutList intIndex, mlng���ID
    ElseIf tbcDetail.Selected.Index = mconTab_CA_Stuff Then
        LoadDetailList intIndex, mlng���ID
    ElseIf tbcDetail.Selected.Index = mconTab_CA_Mistake Then
        LoadMistakeList intIndex, mlng���ID
    End If
End Sub


Private Sub GetSelect(ByVal strInput As String)
    Dim vRect As RECT
    Dim strReturn As String
    Dim sngX As Single
    Dim sngY As Single
    Dim sngH As Single
    
    vRect = zlControl.GetControlRect(txt��ϸ����.hwnd)
    sngX = vRect.Left + picShowDetail.Left
    sngY = vRect.Top + picShowDetail.Top
    
    strReturn = SelectInput(strInput, sngX, sngY, sngH)
    
    If strReturn = "" Then Exit Sub
            
    txt��ϸ����.Tag = Val(Split(strReturn, ";")(0))
    txt��ϸ����.Text = Split(strReturn, ";")(1)
End Sub


Private Function SelectInput(ByVal strKey As String, ByVal sngX As Single, ByVal sngY As Single, ByVal sngH As Single) As String
    Dim strFindString As String
    Dim blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset
    
    strKey = UCase(Trim(strKey))
    
    If strKey <> "" Then
        strFindString = " And (B.���� Like [1] OR C.���� Like [2] OR C.���� LIKE [2])"
        
        If IsNumeric(strKey) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
             strFindString = " And (B.���� Like [1] Or C.���� Like [2] And C.����=3)"
        ElseIf zlStr.IsCharAlpha(strKey) Then         '01,11.����ȫ����ĸʱֻƥ�����
             strFindString = " And C.���� Like [2] "
        ElseIf zlStr.IsCharChinese(strKey) Then
            strFindString = " And B.���� Like [2] "
        End If
    End If
    
    gstrSQL = "Select Rownum As ID, ����id, ��������, ���, ���� as ������,ɢװ��λ, ��װ��λ, ��װϵ�� " & _
        " From (Select Distinct A.����id, B.����, '['||B.����||']'|| B.���� As ��������, B.���, B.����," & _
        "         B.���㵥λ As ɢװ��λ, A.��װ��λ, A.����ϵ�� ��װϵ�� " & _
        "       From �������� A, " & _
        "      (Select B.ID, B.����, B.����, B.���,B.����,B.���㵥λ From �շ���ĿĿ¼ B, �շ���Ŀ���� C " & _
        "       Where (B.վ�� = [4] Or B.վ�� is Null) And B.ID = C.�շ�ϸĿid And B.��� In ('4') " & strFindString & _
        ") B, �շ���Ŀ���� C "
    
    If Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)) > 0 Then
        gstrSQL = gstrSQL & ", �շ�ִ�п��� D "
    End If
    
    gstrSQL = gstrSQL & " Where A.����id = B.ID And A.����id = C.�շ�ϸĿid(+) And C.����(+) = 3 "
    
    If Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)) > 0 Then
        gstrSQL = gstrSQL & " And A.����ID = D.�շ�ϸĿID And ִ�п���ID = [3] "
    End If

    gstrSQL = gstrSQL & " Order By B.����)"
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ҩƷѡ����", False, "", "ѡ��ҩƷ", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                    strKey & "%", "%" & strKey & "%", _
                    Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)), gstrNodeNo)
    
    If blnCancel = True Then Exit Function
    
    If rsTemp Is Nothing Then
        SelectInput = ""
    Else
        SelectInput = rsTemp!����ID & ";" & rsTemp!�������� & ";" & rsTemp!��װ��λ & "," & rsTemp!��װϵ�� & "|" & rsTemp!ɢװ��λ & "," & "1"
    End If
End Function


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1
            Item.Handle = picMain.hwnd
    End Select
End Sub

Private Sub Form_Load()
    Me.Width = mcstlngWinNormalWidth
    Me.Height = mcstlngWinNormalHeight
    
    mint��淽ʽ = Val(zlDatabase.GetPara(281, 100, , 0))
    mint���ʱ�� = Val(zlDatabase.GetPara(281, 100, , 0))
    mstr��ǰ���� = Format(sys.Currentdate, "yyyy-mm-dd")

    Call GetPrivs
    
    Call initGrid   '��ʼ���б� ��ϲ���Ϣ
    Call InitDetailRec '���������ϸ��¼��
    If GetStockName = False Then
        Unload Me
        Exit Sub
    End If
    Call IniDrugUnit 'Ϊ��λ�����б����ֵ
    
    Call InitComandBars
    Call InitPanes
    Call InitTabControl
  
    '    ����Զ��屨��
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    RestoreWinState Me, App.ProductName, mstrCaption
    
    
    mblnStart = True
    Call CheckClosAccount

    '��������
    Call GetAccountRecord
    Call RefreshList
    
    If mint���ʱ�� = 0 Then
        Me.Caption = "���Ľ�����(ÿ�����һ����)"
    Else
        Me.Caption = "���Ľ�����(ÿ��" & mint���ʱ�� & "�ս��)"
    End If
End Sub


Private Sub initGrid()
    Const cstRowHeight = 300
    
    With vsfList
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .RowHeightMin = cstRowHeight
    End With


    With vsfDetail
        .Rows = 1
        .RowHeightMin = cstRowHeight
        
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True
        .MergeCol(1) = True
    End With

    With vsfStuff
        .Rows = 2
        .RowHeightMin = cstRowHeight
        
        .Cell(flexcpFontBold, 0, 0, 1, .Cols - 1) = True
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
    End With
    
    With vsfMistake(0)
        .RowHeightMin = cstRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
    End With
    
    With vsfMistake(1)
        .RowHeightMin = cstRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
    End With
End Sub


Private Sub GetAccountRecord()
    Dim rsTmp As ADODB.Recordset

    On Error GoTo ErrHandle

    gstrSQL = "Select A.ID, Nvl(A.�ⷿid, 0) As �ⷿid, B.���� As �ⷿ����, A.�ڳ�����, A.��ĩ����, A.������" & _
              "   , A.��������, �����, �������,ȡ����,ȡ������, Nvl(A.�ϴν��id, 0) As �ϴν��id,a.�ڼ�,a.���� " & _
              "From ���Ͻ���¼ A, ���ű� B " & _
              "Where A.�ⷿid = B.ID(+) "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "ȡ����¼")

    Call InitAccountRec

    With mrsAccount
        Do While Not rsTmp.EOF
            .AddNew
            !Id = rsTmp!Id
            !�ⷿid = rsTmp!�ⷿid
            !�ⷿ���� = NVL(rsTmp!�ⷿ����, "ȫԺ")
            !�ڳ����� = IIf(IsNull(rsTmp!�ڳ�����), "", Format(rsTmp!�ڳ�����, "YYYY-MM-DD HH:MM:SS"))
            !��ĩ���� = IIf(IsNull(rsTmp!��ĩ����), "", Format(rsTmp!��ĩ����, "YYYY-MM-DD HH:MM:SS"))
            !������ = NVL(rsTmp!������, "")
            !�������� = IIf(IsNull(rsTmp!��������), "", Format(rsTmp!��������, "YYYY-MM-DD HH:MM:SS"))
            !����� = NVL(rsTmp!�����, "")
            !������� = IIf(IsNull(rsTmp!�������), "", Format(rsTmp!�������, "YYYY-MM-DD HH:MM:SS"))
            !ȡ���� = NVL(rsTmp!ȡ����, "")
            !ȡ������ = IIf(IsNull(rsTmp!ȡ������), "", Format(rsTmp!ȡ������, "YYYY-MM-DD HH:MM:SS"))
            !�ϴν��id = rsTmp!�ϴν��id
            !�ڼ� = NVL(rsTmp!�ڼ�, "")
            !���� = Val(NVL(rsTmp!����, "1"))
            .Update

            rsTmp.MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshList()
    'ˢ�½���¼�б�,Ϊ����¼�������ֵ
    Dim strFilter As String
    Dim Str�ڼ� As String
    Dim strSQL As String
    
    Str�ڼ� = Format(sys.Currentdate, "yyyyMM")

    mrsAccount.Filter = "�ⷿid=" & Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    mrsAccount.Sort = "������� Desc"
    
    With vsfList
        .Redraw = flexRDNone
        .Rows = 1
        
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        
        Do While Not mrsAccount.EOF
            .Rows = .Rows + 1

            .TextMatrix(.Rows - 1, .ColIndex("���ID")) = mrsAccount!Id
            .TextMatrix(.Rows - 1, .ColIndex("�ϴν��ID")) = mrsAccount!�ϴν��id
            .TextMatrix(.Rows - 1, .ColIndex("�ⷿID")) = NVL(mrsAccount!�ⷿid, 0)

            .TextMatrix(.Rows - 1, .ColIndex("�ڳ�����")) = Format(mrsAccount!�ڳ�����, "YYYY-MM-DD HH:MM:SS")
            .TextMatrix(.Rows - 1, .ColIndex("��ĩ����")) = Format(mrsAccount!��ĩ����, "YYYY-MM-DD HH:MM:SS")

            .TextMatrix(.Rows - 1, .ColIndex("������")) = mrsAccount!������
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = Format(mrsAccount!��������, "YYYY-MM-DD HH:MM:SS")
            .TextMatrix(.Rows - 1, .ColIndex("�����")) = mrsAccount!�����
            .TextMatrix(.Rows - 1, .ColIndex("�������")) = Format(mrsAccount!�������, "YYYY-MM-DD HH:MM:SS")
            .TextMatrix(.Rows - 1, .ColIndex("ȡ����")) = mrsAccount!ȡ����
            .TextMatrix(.Rows - 1, .ColIndex("ȡ������")) = Format(mrsAccount!ȡ������, "YYYY-MM-DD HH:MM:SS")
            .TextMatrix(.Rows - 1, .ColIndex("�ڼ�")) = mrsAccount!�ڼ�
            .TextMatrix(.Rows - 1, .ColIndex("����")) = mrsAccount!����
            
            If mrsAccount!���� = 0 Then
                .Cell(flexcpPicture, .Rows - 1, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = picIni.Picture
            End If
            
            '�ڳ�������ɫ��ʶ
            If mrsAccount!���� = 0 Then
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
            End If
            
            'δ��������ú�ɫ��ʶ
            If Format(mrsAccount!�������, "YYYY-MM-DD HH:MM:SS") = "" Then
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
            End If
            
            mrsAccount.MoveNext
        Loop
                  
        If .Rows > 1 Then
            .Row = 1
        End If

        .Redraw = flexRDDirect
    End With
    
    stbThis.Panels(2).Text = ""
    If vsfList.Rows = 1 Then
        stbThis.Panels(2).Text = "[" & cbo�ⷿ.Text & "]" & "���ڳ�������ݣ���ͨ���̵�ȷ�ʽȷ����ǰ�ⷿ������ȷ��" & vbCrLf & "���������ֹ�������������ݻ���ÿ�¹̶������Զ�����������ݣ�"
    End If
End Sub


Private Function IsInString(ByVal strTarget As String, ByVal strOrigin As String, Optional strSplit As String = "") As Boolean
    'ĳ���ַ����Ƿ������һ���ַ���
    'strTarget��Ŀ���ַ���
    'strOrigin��ԭ�ַ���
    'strSplit���ָ�������Ϊ��ʱΪ��ȷƥ�䣩
    '��strTarget���Ƿ����strOrigin
    
    IsInString = InStrB(strSplit & strTarget & strSplit, strSplit & strOrigin & strSplit) > 0
End Function



Private Sub InitComandBars()
    '��ʼ���˵�������ȫ���˵����������������˵���
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
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
    Me.cbsMain.Icons = Me.imgPublic.Icons
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsMain.ActiveMenuBar.Title = "�˵�"
    Me.cbsMain.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.Id = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        cbrControlMain.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.Id = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_AddIniAccount, "��ʼ��(&N)")
        cbrControlMain.Visible = mPrives.bln���
        '
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_AddNewAccount, "���(A)")
        cbrControlMain.Enabled = IIf(mint��淽ʽ = -1, True, False)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_DeleteAccount, "ɾ��(D)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_VerifyAccount, "���(&V)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_CancelAccount, "ȡ��(C)")
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.Id = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_CA_Refresh, "ˢ��(&R)")
    End With

    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.Id = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControlMain = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�����")
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, "������ҳ(&H)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, "������̳(&F)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        
        Set cbrControlMain = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��")
        cbrControlMain.BeginGroup = True
    End With
    
    '�����
    With Me.cbsMain.KeyBindings
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F5, mconMenu_CA_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
  
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_AddIniAccount, "��ʼ��")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_AddNewAccount, "���")
        cbrControlMain.Enabled = IIf(mint��淽ʽ = -1, True, False)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_DeleteAccount, "ɾ��")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_VerifyAccount, "���")
        cbrControlMain.Visible = mPrives.bln���
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_CA_CancelAccount, "ȡ��")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_CA_Refresh, "ˢ��")
        cbrControlMain.BeginGroup = True
         
        Set cbrControlMain = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        cbrControlMain.BeginGroup = True
    End With
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub InitPanes()
    '��ʼ�������ؼ�
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
'    Me.dkpMain.Options.DefaultPaneOptions = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
    
    Dim objPaneCon As Pane
    Dim objPaneList As Pane
    Dim objPaneDetail As Pane
    
    Set objPaneCon = Me.dkpMain.CreatePane(mconPane_Dept_Condition, 250, 100, DockLeftOf, Nothing)
    objPaneCon.Title = "�����ϸ"
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable
'    objPaneCon.MaxTrackSize.SetSize 290, 500
End Sub


Private Sub InitTabControl()
    '��ʼ����ҳ�ؼ�
    
    With Me.tbcDetail
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(mconTab_CA_Detail, "�����ϸ", Me.picShowDetail.hwnd, 0).Tag = "�����ϸ_"
        .InsertItem(mconTab_CA_Stuff, "������ϸ", Me.picShowDetail.hwnd, 0).Tag = "������ϸ_"
        .InsertItem(mconTab_CA_Mistake, "�����ϸ", Me.picShowDetail.hwnd, 0).Tag = "�����ϸ_"
        
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < mcstlngWinNormalWidth Then Me.Width = mcstlngWinNormalWidth
    If Me.Height < mcstlngWinNormalHeight Then Me.Height = mcstlngWinNormalHeight
End Sub



Private Sub Form_Unload(Cancel As Integer)
    mblnStart = False
    
    SaveWinState Me, App.ProductName, mstrCaption
End Sub

Private Sub picDetail_Resize()
    On Error Resume Next
    
    With fraLine
'        .Top = 0
        .Left = 0
        .Height = picDetail.Height + 100
    End With
    
    With tbcDetail
        .Top = 0
        .Left = fraLine.Left + 50
        .Width = picDetail.Width - fraLine.Width
        .Height = picDetail.Height - 50
    End With
End Sub



Private Sub picList_Resize()
    On Error Resume Next
    
    With vsfList
        .Move 0, 0, picList.Width, picList.Height
    End With
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    
    With cbo�ⷿ
        .Width = picMain.Width - .Left - 60
    End With

    With picList
        .Top = cbo�ⷿ.Top + cbo�ⷿ.Height + 120
        .Left = 0
        .Width = picMain.Width
        .Height = picMain.Height - .Top
    End With
End Sub


Private Sub picShowDetail_Resize()
    On Error Resume Next
    
    With vsfDetail
        .Top = txt��ϸ����.Top + txt��ϸ����.Height + 120
        .Left = 0
        .Width = picShowDetail.Width
        .Height = picShowDetail.Height - .Top
    End With
    
    
    With vsfStuff
        .Top = vsfDetail.Top
        .Left = 0
        .Width = vsfDetail.Width
        .Height = vsfDetail.Height
    End With
    
    With vsfMistake(0)
        .Top = vsfDetail.Top
        .Left = 0
        .Width = vsfDetail.Width
        .Height = vsfDetail.Height
    End With
    
    With vsfMistake(1)
        .Top = vsfDetail.Top
        .Left = 0
        .Width = vsfDetail.Width
        .Height = vsfDetail.Height
    End With
End Sub


Private Sub tbcDetail_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim intIndex As Integer
    
    vsfDetail.Visible = False
    vsfStuff.Visible = False
    vsfMistake(0).Visible = False
    vsfMistake(1).Visible = False
    
    With vsfList
        If .TextMatrix(.Row, .ColIndex("�����")) = "" Then
            intIndex = 0
        Else
            intIndex = 1
        End If
    End With
    
    If Item.Index = mconTab_CA_Detail Then
        vsfDetail.Visible = True
        LoadInOutList intIndex, mlng���ID
    ElseIf Item.Index = mconTab_CA_Stuff Then
        vsfStuff.Visible = True
        LoadDetailList intIndex, mlng���ID
    ElseIf Item.Index = mconTab_CA_Mistake Then
         vsfMistake(1).Visible = True
        LoadMistakeList intIndex, mlng���ID
    End If
End Sub


Private Function GetStockName() As Boolean
    'ȡ��ǰ����Ա��������Ŀⷿ
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle

    gstrSQL = "SELECT DISTINCT a.id, a.���� || '-' || a.���� as ���� " _
            & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
            & "Where (a.վ�� = [2] Or a.վ�� is Null) And c.�������� = b.���� " _
            & "  AND Instr('WV',b.����,1) > 0 " _
            & "  AND a.id = c.����id " _
            & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
            & IIf(mPrives.bln���пⷿ = True, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])") _
            & "Order by a.���� || '-' || a.���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����Ա��������Ŀⷿ", UserInfo.Id, gstrNodeNo)
    
    If rsTmp.EOF Then
        MsgBox "��ǰ����Ա�������κοⷿ�����ܽ��н�������", vbInformation, gstrSysName
        Exit Function
    End If
    
    With rsTmp
        cbo�ⷿ.Clear
        
        Do While Not .EOF
            cbo�ⷿ.AddItem !����
            cbo�ⷿ.ItemData(cbo�ⷿ.NewIndex) = !Id
          
            .MoveNext
        Loop
        
        cbo�ⷿ.ListIndex = 0
    End With
    
    GetStockName = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub IniDrugUnit()
    '����ʹ�õĵ�λ
    With cbo��λ
        .Clear
        .AddItem "ɢװ��λ"
        .AddItem "��װ��λ"
        .ListIndex = 0
    End With
End Sub




Private Sub txt��ϸ����_GotFocus()
    zlControl.TxtSelAll txt��ϸ����
End Sub

Private Sub txt��ϸ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub

    txt��ϸ����_Validate True
End Sub

Private Sub txt��ϸ����_KeyPress(KeyAscii As Integer)
    If InStr(" ';", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��ϸ����_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(txt��ϸ����.Text) = "" Then
        txt��ϸ����.Tag = 0
    End If
End Sub

Private Sub txt��ϸ����_Validate(Cancel As Boolean)
    Dim intIndex As Integer
    
    With vsfList
        If .TextMatrix(.Row, .ColIndex("�����")) = "" Then
            intIndex = 0
        Else
            intIndex = 1
        End If
    End With
    
    With txt��ϸ����
        If Trim(.Text) = "" Then
            .Tag = 0
        Else
            GetSelect .Text
        End If

        DoEvents

        If tbcDetail.Selected.Index = mconTab_CA_Detail Then
            LoadInOutList intIndex, mlng���ID
        ElseIf tbcDetail.Selected.Index = mconTab_CA_Stuff Then
            LoadDetailList intIndex, mlng���ID
        ElseIf tbcDetail.Selected.Index = mconTab_CA_Mistake Then
            LoadMistakeList intIndex, mlng���ID
        End If
    End With
End Sub

Private Sub vsfList_EnterCell()
    Dim i As Integer
    Dim intIndex As Integer
    
    With vsfList
        If Val(.Tag) = .Row Then Exit Sub

        .Tag = .Row

        If .Row <= vsfList.FixedRows - 1 Then
            ClearDetailList
            ClearStuffList
            ClearMistakeList
            Exit Sub
        End If

        If Val(.TextMatrix(.Row, .ColIndex("���ID"))) = 0 Then
            ClearDetailList
            ClearStuffList
            ClearMistakeList
            Exit Sub
        End If
        
        With vsfList
            If .TextMatrix(.Row, .ColIndex("�����")) = "" Then
                intIndex = 0
            Else
                intIndex = 1
              End If
        End With
        
        Call LoadData(intIndex)
        
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, .Row, 1)
        .Redraw = flexRDDirect
    End With
 End Sub


Private Sub ClearDetailList()
    vsfDetail.Rows = 1
    vsfDetail.Rows = 2
End Sub

Private Sub ClearStuffList()
    vsfStuff.Rows = 2
    vsfStuff.Rows = 3
End Sub

Private Sub ClearMistakeList()
    vsfMistake(0).Rows = 1
    vsfMistake(0).Rows = 2

    vsfMistake(1).Rows = 1
    vsfMistake(1).Rows = 2
End Sub

Private Sub LoadData(ByVal intType As Integer)
    '��ȡ����
    Dim lng�ϴν��ID As Long
    Dim str�ڳ����� As String
    Dim str��ĩ���� As String
    Dim rsTemp As ADODB.Recordset

    With vsfList
        If .Row = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("���ID"))) = 0 Then Exit Sub

        mlng���ID = Val(.TextMatrix(.Row, .ColIndex("���ID")))
        lng�ϴν��ID = Val(.TextMatrix(.Row, .ColIndex("�ϴν��ID")))
        gstrSQL = "Select �ڳ�����, ��ĩ���� From ���Ͻ���¼ Where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ڲ�ѯ", mlng���ID)
        
        If rsTemp.RecordCount <= 0 Then '��Ӧ����¼�ѱ�ɾ��
            MsgBox "�ü�¼�Ѿ���ɾ����", vbInformation, gstrSysName
            GetAccountRecord
            RefreshList
            Exit Sub
        End If
        
        str�ڳ����� = IIf(IsNull(rsTemp!�ڳ�����), "", rsTemp!�ڳ�����)
        str��ĩ���� = IIf(IsNull(rsTemp!��ĩ����), "", rsTemp!��ĩ����)
    End With

    Call AviShow(Me)

    On Error GoTo ErrHandle

    GetDetailRecord intType, mlng���ID, str�ڳ�����, str��ĩ����

    vsfDetail.Visible = False
    vsfStuff.Visible = False
    vsfMistake(0).Visible = False
    vsfMistake(1).Visible = False

    If tbcDetail.Selected.Index = mconTab_CA_Detail Then
        LoadInOutList intType, mlng���ID
    ElseIf tbcDetail.Selected.Index = mconTab_CA_Stuff Then
        LoadDetailList intType, mlng���ID
    ElseIf tbcDetail.Selected.Index = mconTab_CA_Mistake Then
        LoadMistakeList intType, mlng���ID
    End If

    Call AviShow(Me, False)

    Exit Sub
ErrHandle:
    Call AviShow(Me, False)
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadInOutList(ByVal intType As Integer, ByVal lng���id As Long)
     '�����������
    Dim lng�ⷿID As Long
    Dim lng����ID As Long
    Dim blnShowSubType As Boolean
    Dim strFilter As String
    Dim strOrder As String
    
    Dim strҵ����� As String
    Dim strҵ����� As String
    
    Dim dbl���� As Double
    Dim Dbl��� As Double
    Dim dbl��� As Double
    
    Dim intRow As Integer
    
    
    Dim strTmp As String
    Dim str��λ As String
    Dim dbl��װ As String
    
    Dim intShowNumberDigit As Integer          '����С��λ��
    Dim intShowMoneyDigit As Integer           '���С��λ��
    Dim intUnit As Integer  '1-ɢװ;2-��װ
    
    Call ClearDetailList
    
    'ȡ�۸�������������ʾ����intUnit= 1-ɢװ
    If cbo��λ.ListIndex = 0 Then
        intUnit = 1
    ElseIf cbo��λ.ListIndex = 1 Then
        intUnit = 2
    End If
            
    intShowNumberDigit = newGetDigit(0, 3, intUnit)
    intShowMoneyDigit = newGetDigit(0, 4)
    
    lng�ⷿID = Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    lng����ID = Val(txt��ϸ����.Tag)
    
    strFilter = "����=" & intType & " And ���ID=" & lng���id
    If lng�ⷿID > 0 Then strFilter = strFilter & " And �ⷿID=" & lng�ⷿID
    If lng����ID > 0 Then strFilter = strFilter & " And ����ID=" & lng����ID
    
    strOrder = "ҵ�����,ҵ�����"
    
    vsfDetail.Visible = True
    
    
    With vsfDetail
        mrsDetail.Filter = strFilter
        mrsDetail.Sort = strOrder
        If mrsDetail.RecordCount = 0 Then
            Exit Sub
        End If
        
        .Redraw = flexRDNone
        .Rows = 1
        
        Do While Not mrsDetail.EOF
            If cbo��λ.ListIndex = 0 Then
                str��λ = mrsDetail!ɢװ��λ
                dbl��װ = mrsDetail!ɢװ��װ
            ElseIf cbo��λ.ListIndex = 1 Then
                str��λ = mrsDetail!��װ��λ
                dbl��װ = mrsDetail!��װ��װ
            End If
            
            If strҵ����� & strҵ����� <> mrsDetail!ҵ����� & mrsDetail!ҵ����� Then
                .Rows = .Rows + 1
                intRow = .Rows - 1
                
                strҵ����� = mrsDetail!ҵ�����
                strҵ����� = mrsDetail!ҵ�����
            End If
            
            .TextMatrix(intRow, .ColIndex("ҵ�����")) = mrsDetail!ҵ�����
            .TextMatrix(intRow, .ColIndex("ҵ�����")) = mrsDetail!ҵ�����
            
            
            If mrsDetail!���� <> 0 Then .TextMatrix(intRow, .ColIndex("����")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("����"))) + mrsDetail!���� / dbl��װ, intShowNumberDigit, , True)
            If mrsDetail!��� <> 0 Then .TextMatrix(intRow, .ColIndex("�ۼ۽��")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("�ۼ۽��"))) + mrsDetail!���, intShowMoneyDigit, , True)
            If mrsDetail!��� <> 0 Then .TextMatrix(intRow, .ColIndex("���")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("���"))) + mrsDetail!���, intShowMoneyDigit, , True)
            .TextMatrix(intRow, .ColIndex("�ɱ����")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("�ۼ۽��"))) - Val(.TextMatrix(intRow, .ColIndex("���"))), intShowMoneyDigit, , True)
            
            
            mrsDetail.MoveNext
        Loop
        
        'ָ�����Ķ���������
        If lng����ID > 0 Then .TextMatrix(0, .ColIndex("����")) = "����(" & str��λ & ")"
        .ColHidden(.ColIndex("����")) = Not (lng����ID > 0)
        If .ColHidden(.ColIndex("����")) Then
            .ColWidth(.ColIndex("����")) = 0
        Else
            .ColWidth(.ColIndex("����")) = 1500
        End If

        '������Ϊ��ɫ�������Ϊ��ɫ
        For intRow = 2 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("ҵ�����")) = "3-����" Then
                .Cell(flexcpForeColor, intRow, .ColIndex("����"), intRow, .ColIndex("�ɱ����")) = vbRed
            ElseIf .TextMatrix(intRow, .ColIndex("ҵ�����")) = "2-���" Then
                .Cell(flexcpForeColor, intRow, .ColIndex("����"), intRow, .ColIndex("�ɱ����")) = vbBlue
            End If
        Next

        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub LoadDetailList(ByVal intType As Integer, ByVal lng���id As Long)
    '���������ϸ
    Dim lng�ⷿID As Long
    Dim lng����ID As Long
    Dim blnShowSubType As Boolean
    Dim strFilter As String
    Dim strOrder As String
    
    Dim strҵ����� As String
    Dim strҵ����� As String
    
    Dim str�������� As String
    
    Dim dbl���� As Double
    Dim Dbl��� As Double
    Dim dbl��� As Double
    
    Dim lngRow As Long
    
    Dim strTmp As String
    Dim str��λ As String
    Dim dbl��װ As String
    
    Dim intShowNumberDigit As Integer          '����С��λ��
    Dim intShowMoneyDigit As Integer           '���С��λ��
    Dim intUnit As Integer  '1-ɢװ;2-��װ;
    
    ClearStuffList
    
    vsfStuff.Visible = True
    
    'ȡ�۸�������������ʾ����
    If cbo��λ.ListIndex = 0 Then
        intUnit = 1
    ElseIf cbo��λ.ListIndex = 1 Then
        intUnit = 2
    End If
            
    intShowNumberDigit = newGetDigit(0, 3, intUnit)
    intShowMoneyDigit = newGetDigit(0, 4)
    
    lng�ⷿID = Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    lng����ID = Val(txt��ϸ����.Tag)
    
    strFilter = "����=" & intType & " And ���ID=" & lng���id
    If lng�ⷿID > 0 Then strFilter = strFilter & " And �ⷿID=" & lng�ⷿID
    If lng����ID > 0 Then strFilter = strFilter & " And ����ID=" & lng����ID
    
    strOrder = "��������,ҵ�����,ҵ�����"
    
    With vsfStuff
        mrsDetail.Filter = strFilter
        mrsDetail.Sort = strOrder
        If mrsDetail.RecordCount = 0 Then
            Exit Sub
        End If
        
        .Redraw = flexRDNone
        .Rows = 2
        
        Do While Not mrsDetail.EOF
            If cbo��λ.ListIndex = 0 Then
                str��λ = mrsDetail!ɢװ��λ
                dbl��װ = mrsDetail!ɢװ��װ
            ElseIf cbo��λ.ListIndex = 1 Then
                str��λ = mrsDetail!��װ��λ
                dbl��װ = mrsDetail!��װ��װ
            End If
            
            If lng�ⷿID = 0 And (mrsDetail!ҵ����� = "���Ŀⷿ���" Or mrsDetail!ҵ����� = "���Ŀⷿ����") Then
                'ͳ��ȫԺʱ���������ڲ���ͨ���ƿ⣩
            Else
                If str�������� <> mrsDetail!�������� Then
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                   
                    str�������� = mrsDetail!��������
                End If
                
                .TextMatrix(lngRow, .ColIndex("��������")) = mrsDetail!��������
                .TextMatrix(lngRow, .ColIndex("��Ʒ��")) = NVL(mrsDetail!��Ʒ��)
                .TextMatrix(lngRow, .ColIndex("���")) = mrsDetail!���
                .TextMatrix(lngRow, .ColIndex("��λ")) = str��λ
                
                If mrsDetail!ҵ����� = "1-�ڳ�" Then
                    .TextMatrix(lngRow, .ColIndex("�ڳ�����")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("�ڳ�����"))) + mrsDetail!���� / dbl��װ, intShowNumberDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("�ڳ����")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("�ڳ����"))) + mrsDetail!���, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("�ڳ����")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("�ڳ����"))) + mrsDetail!���, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("�ڳ��ɱ�")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("�ڳ����"))) - Val(.TextMatrix(lngRow, .ColIndex("�ڳ����"))), intShowMoneyDigit, , True)
                ElseIf mrsDetail!ҵ����� = "2-���" Then
                    .TextMatrix(lngRow, .ColIndex("�������")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("�������"))) + mrsDetail!���� / dbl��װ, intShowNumberDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("�����")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("�����"))) + mrsDetail!���, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("�����")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("�����"))) + mrsDetail!���, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("���ɱ�")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("�����"))) - Val(.TextMatrix(lngRow, .ColIndex("�����"))), intShowMoneyDigit, , True)
                ElseIf mrsDetail!ҵ����� = "3-����" Then
                    .TextMatrix(lngRow, .ColIndex("��������")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("��������"))) + mrsDetail!���� / dbl��װ, intShowNumberDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("������")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("������"))) + mrsDetail!���, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("������")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("������"))) + mrsDetail!���, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("����ɱ�")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("������"))) - Val(.TextMatrix(lngRow, .ColIndex("������"))), intShowMoneyDigit, , True)
                ElseIf mrsDetail!ҵ����� = "4-��ĩ" Then
                    .TextMatrix(lngRow, .ColIndex("��ĩ����")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("��ĩ����"))) + mrsDetail!���� / dbl��װ, intShowNumberDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("��ĩ���")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("��ĩ���"))) + mrsDetail!���, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("��ĩ���")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("��ĩ���"))) + mrsDetail!���, intShowMoneyDigit, , True)
                    .TextMatrix(lngRow, .ColIndex("��ĩ�ɱ�")) = zlStr.FormatEx(Val(.TextMatrix(lngRow, .ColIndex("��ĩ���"))) - Val(.TextMatrix(lngRow, .ColIndex("��ĩ���"))), intShowMoneyDigit, , True)
                End If
            End If
            
            mrsDetail.MoveNext
        Loop
                
        '������Ϊ��ɫ�������Ϊ��ɫ
        .Cell(flexcpForeColor, 2, .ColIndex("�������"), .Rows - 1, .ColIndex("���ɱ�")) = vbBlue
        .Cell(flexcpForeColor, 2, .ColIndex("��������"), .Rows - 1, .ColIndex("����ɱ�")) = vbRed
       
        '�ϼƴ����ʾ
'        .Cell(flexcpFontBold, 2, .ColIndex("��ĩ����"), .rows - 1, .ColIndex("��ĩ�ɱ�")) = True
            
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub LoadMistakeList(ByVal intType As Integer, ByVal lng���id As Long)
    Dim lng�ⷿID As Long
    Dim lng����ID As Long
    Dim strFilter As String
    Dim strOrder As String
    
    Dim str�ⷿ As String
    
    Dim dbl������ As Double
    Dim dbl���� As Double
    Dim dbl��۲� As Double
    
    Dim intRow As Integer
    Dim strUnit As String
    
    Dim intShowNumberDigit As Integer          '����С��λ��
    Dim intShowMoneyDigit As Integer           '���С��λ��
    Dim intUnit As Integer  '1-�ۼ�;2-����;3-סԺ;4-ҩ��;
    
    On Error GoTo ErrHandle

    Call ClearMistakeList
    
    'ȡ�۸�������������ʾ����
    If cbo��λ.ListIndex = 0 Then
        intUnit = 1
    ElseIf cbo��λ.ListIndex = 1 Then
        intUnit = 2
    End If
            
    intShowNumberDigit = newGetDigit(0, 3, intUnit)
    intShowMoneyDigit = newGetDigit(0, 4)
    
    GetMistakeRecord intType, lng���id
    
    lng�ⷿID = Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    lng����ID = Val(txt��ϸ����.Tag)
    
    strFilter = "����=" & intType & " And ���ID=" & lng���id
    If lng�ⷿID > 0 Then strFilter = strFilter & " And �ⷿID=" & lng�ⷿID
    If lng����ID > 0 Then strFilter = strFilter & " And ����ID=" & lng����ID
    
    strOrder = "��������,����"
    If lng�ⷿID > 0 Then strOrder = "�ⷿ����"
    
    If lng�ⷿID > 0 Then
        vsfMistake(1).Visible = True
        vsfMistake(0).Visible = False
    Else
        vsfMistake(1).Visible = False
        vsfMistake(0).Visible = True
    End If
    
    mrsMistake.Filter = strFilter
    mrsMistake.Sort = strOrder
    If mrsMistake.RecordCount = 0 Then Exit Sub
            
    If lng�ⷿID > 0 Then
        With vsfMistake(1)
            .Redraw = flexRDNone
            
            .Rows = 1
        
            Do While Not mrsMistake.EOF
                .Rows = .Rows + 1
                intRow = .Rows - 1
                
                .TextMatrix(intRow, .ColIndex("����")) = mrsMistake!��������
                .TextMatrix(intRow, .ColIndex("��Ʒ��")) = NVL(mrsMistake!��Ʒ��, "")
                .TextMatrix(intRow, .ColIndex("���")) = mrsMistake!���
                .TextMatrix(intRow, .ColIndex("����")) = mrsMistake!����
                .TextMatrix(intRow, .ColIndex("����")) = zlStr.FormatEx(mrsMistake!����, intShowMoneyDigit, , True)
                .TextMatrix(intRow, .ColIndex("��۲�")) = zlStr.FormatEx(mrsMistake!��۲�, intShowMoneyDigit, , True)
                
                Select Case intUnit
                Case 2  '"��װ��λ"
                    .TextMatrix(intRow, .ColIndex("��λ")) = mrsMistake!��װ��λ
                    .TextMatrix(intRow, .ColIndex("������")) = zlStr.FormatEx(mrsMistake!������ / mrsMistake!��װ��װ, intShowNumberDigit, , True)
                
                Case Else
                    .TextMatrix(intRow, .ColIndex("��λ")) = mrsMistake!���㵥λ
                    .TextMatrix(intRow, .ColIndex("������")) = zlStr.FormatEx(mrsMistake!������, intShowNumberDigit, , True)
                End Select
                
                mrsMistake.MoveNext
            Loop
            
            .Redraw = flexRDDirect
        End With
    Else
        With vsfMistake(0)
            .Redraw = flexRDNone
            
            .Rows = 1
            
            Do While Not mrsMistake.EOF
                If str�ⷿ <> mrsMistake!�ⷿ���� Then
                    .Rows = .Rows + 1
                    intRow = .Rows - 1
                    
                    str�ⷿ = mrsMistake!�ⷿ����
                End If
                
                .TextMatrix(intRow, .ColIndex("�ⷿ")) = mrsMistake!�ⷿ����
                
                Select Case intUnit
                Case 2  '"��װ��λ"
                    .TextMatrix(intRow, .ColIndex("������")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("������"))) + mrsMistake!������ / mrsMistake!��װ��װ, intShowNumberDigit, , True)
                Case Else
                  .TextMatrix(intRow, .ColIndex("������")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("������"))) + mrsMistake!������, intShowNumberDigit, , True)
                End Select
                
                .TextMatrix(intRow, .ColIndex("����")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("����"))) + mrsMistake!����, intShowMoneyDigit, , True)
                .TextMatrix(intRow, .ColIndex("��۲�")) = zlStr.FormatEx(Val(.TextMatrix(intRow, .ColIndex("��۲�"))) + mrsMistake!��۲�, intShowMoneyDigit, , True)
                
                mrsMistake.MoveNext
            Loop
            
            .Redraw = flexRDDirect
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'
Private Sub GetMistakeRecord(ByVal intType As Integer, ByVal lng���id As Long)
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    '�ж��Ƿ��ѱ���ý������¼
    mrsMistake.Filter = "����=" & intType & " And ���ID=" & lng���id
    If mrsMistake.RecordCount > 0 Then Exit Sub

    mrsMistake.Filter = ""

    'û�ҵ�ʱ�����ݿ��ȡ
    '[' || B.���� || ']' As ����, B.���� As ͨ����, E.���� As ��Ʒ��
    gstrSQL = "Select Distinct A.���id, A.�ⷿid, A.����id, Nvl(A.����, 0) ����, A.������, A.����, A.��۲�, " & _
        " '[' || F.���� || ']' As ����, F.���� As ͨ����, E.���� As ��Ʒ��, F.���, D.���� As �ⷿ����, F.���㵥λ, " & _
        " B.��װ��λ, B.����ϵ�� ��װ��װ " & _
        " From ���Ͻ����� A, �������� B, �շ���ĿĿ¼ F, �շ���Ŀ���� E, ���ű� D " & _
        " Where A.����id = B.����id And B.����id = F.ID And A.�ⷿid = D.ID And B.����id = E.�շ�ϸĿid(+) And " & _
        " E.����(+) = 3 And A.���id = [1] "

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "��ȡ�����ϸ��¼", lng���id)

    '���½����ϸ���ݼ�
    With rsTmp
        Do While Not .EOF
            mrsMistake.AddNew
            mrsMistake!���� = intType
            mrsMistake!���ID = lng���id
            mrsMistake!�ⷿid = !�ⷿid
            mrsMistake!����ID = !����ID
            mrsMistake!���� = !����
            mrsMistake!�ⷿ���� = !�ⷿ����
            
            mrsMistake!�������� = !���� & !ͨ����

            mrsMistake!��Ʒ�� = NVL(!��Ʒ��, "")
            mrsMistake!��� = NVL(!���, "")
            mrsMistake!������ = NVL(!������, 0)
            mrsMistake!���� = NVL(!����, 0)
            mrsMistake!��۲� = NVL(!��۲�, 0)
            mrsMistake!���㵥λ = !���㵥λ
            mrsMistake!��װ��λ = !��װ��λ
            mrsMistake!��װ��װ = !��װ��װ
            mrsMistake.Update

            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub





Private Sub GetDetailRecord(ByVal intType As Integer, ByVal lng���id As Long, ByVal str�ڳ����� As String, ByVal str��ĩ���� As String)
    Dim rsTmp As ADODB.Recordset
    Dim strSqlUnit As String
    Dim strSqlGroup As String

    On Error GoTo ErrHandle
    '�ж��Ƿ��Ѽ�¼�ý����ϸ
    mrsDetail.Filter = "����=" & intType & " And ���ID=" & lng���id
    If mrsDetail.RecordCount > 0 Then Exit Sub

    mrsDetail.Filter = ""

    ''''û�ҵ�ʱ�����ݿ��ȡ
    gstrSQL = "Select Distinct  A.*, E.���� As ��Ʒ�� From ("

    'ȡ���ڽ�����ĩ������Ϊ���ڵ��ڳ�����
    gstrSQL = gstrSQL & "Select A.�ⷿid, A.ҵ�����, A.ҵ�����, '[' || B.���� || ']' As ����, B.���� As ͨ����, B.���, A.����id, Sum(A.����) As ����, Sum(A.���) As ���, Sum(A.���) As ���,C.��װ��λ,C.����ϵ�� ��װ��װ,B.���㵥λ as ɢװ��λ,1 as ɢװ��װ " & _
        " From (Select A.�ⷿid, '1-�ڳ�' As ҵ�����, '' As ҵ�����, A.����id As ����id, Sum(A.�ڳ�����) As ����, Sum(A.�ڳ����) As ���, Sum(A.�ڳ����) As ��� " & _
        "       From ���Ͻ����ϸ A Where ���id = [1] " & _
        "       Group By A.�ⷿid, A.����id "

    'ȡ�ڼ䷢����
    'ע���õ������ͻ�ⷿ�Ĺ���������ȷ��ֻͳ��ҩƷ����
    If str�ڳ����� <> "" Then
        gstrSQL = gstrSQL & _
        "     Union All " & _
        "     Select * From " & _
        "       (Select A.�ⷿid, Decode(B.ϵ��, 1, '2-���', '3-����') As ҵ�����, B.���� As ҵ�����, A.ҩƷid As ҩƷid, Sum(Nvl(A.ʵ������, 0) * Nvl(A.����, 1)) As ����, Sum(Nvl(A.���۽��, 0)) As ���, Sum(Nvl(A.���, 0)) As ��� " & _
        "       From ҩƷ�շ���¼ A, ҩƷ������ B " & _
        "       Where A.������id = B.ID And A.���� In (15, 16, 17, 18, 19, 20, 21, 22, 24, 25, 26) And " & _
        "       A.������� Between [2] And [3] " & _
        "       Group By A.�ⷿid, A.ҩƷid, B.����, Decode(B.ϵ��, 1, '2-���', '3-����')) X " & _
        "     where X.���� <> 0 Or X.��� <> 0 Or X.��� <> 0"
    End If

    'ȡ������ĩ����
    gstrSQL = gstrSQL & _
        "       Union All " & _
        "       Select A.�ⷿid, '4-��ĩ' As ҵ�����, '' As ҵ�����, A.����id, Sum(A.��ĩ����) As ����, Sum(A.��ĩ���) As ���, Sum(A.��ĩ���) As ��� " & _
        "       From ���Ͻ����ϸ A " & _
        "       Where ���id = [1] " & _
        "       Group By A.�ⷿid, A.����id) A, �շ���ĿĿ¼ B, �������� C " & _
        " Where A.����id = B.Id And A.����ID = C.����ID " & _
        " Group By A.ҵ�����, A.ҵ�����, A.�ⷿid, '[' || B.���� || ']' , B.����, B.���, A.����id,C.��װ��λ,C.����ϵ��,B.���㵥λ "

    gstrSQL = gstrSQL & ") A, �շ���Ŀ���� E " & _
        " Where A.����id = E.�շ�ϸĿid(+) And E.����(+) = 3 " & _
        " Order By A.ҵ�����, A.ҵ�����, A.�ⷿid, A.����, A.ͨ����, E.����, A.���, A.����id"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "��ȡ�����ϸ��¼", lng���id, CDate(IIf(str�ڳ����� = "", "1990-01-01", str�ڳ�����)), CDate(str��ĩ����))

    '���½����ϸ���ݼ�
    With rsTmp
        Do While Not .EOF
            mrsDetail.AddNew
            mrsDetail!���� = intType
            mrsDetail!���ID = lng���id
            mrsDetail!ҵ����� = NVL(!ҵ�����, "")
            mrsDetail!ҵ����� = NVL(!ҵ�����, "")
            mrsDetail!�ⷿid = !�ⷿid
            
            mrsDetail!�������� = !���� & !ͨ����

            mrsDetail!��Ʒ�� = NVL(!��Ʒ��, "")
            mrsDetail!��� = NVL(!���, "")
            mrsDetail!����ID = !����ID
            mrsDetail!���� = NVL(!����, 0)
            mrsDetail!��� = NVL(!���, 0)
            mrsDetail!��� = NVL(!���, 0)
            mrsDetail!ɢװ��λ = !ɢװ��λ
            mrsDetail!ɢװ��װ = !ɢװ��װ
            mrsDetail!��װ��λ = !��װ��λ
            mrsDetail!��װ��װ = !��װ��װ
            mrsDetail.Update

            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub InitDetailRec()
    '���������ϸ��¼��
    Set mrsDetail = New ADODB.Recordset
    With mrsDetail
        If .State = 1 Then .Close
        
        .Fields.Append "����", adDouble, 1, adFldIsNullable
        .Fields.Append "���ID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҵ�����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ҵ�����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�ⷿID", adDouble, 18, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "��Ʒ��", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "ɢװ��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ɢװ��װ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��װ��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��װ��װ", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    '����¼��
    Set mrsMistake = New ADODB.Recordset
    With mrsMistake
        If .State = 1 Then .Close
        
        .Fields.Append "����", adDouble, 1, adFldIsNullable
        .Fields.Append "���ID", adDouble, 18, adFldIsNullable
        .Fields.Append "�ⷿID", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "�ⷿ����", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "��Ʒ��", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "������", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "��۲�", adDouble, 18, adFldIsNullable
        .Fields.Append "���㵥λ", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "ɢװ��λ", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "��װ��λ", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "ɢװ��װ", adDouble, 10, adFldIsNullable
        .Fields.Append "��װ��װ", adDouble, 10, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub


Private Sub CheckClosAccount()
    '��������ʱ��˵���Ѿ���ʼ���ˣ��������ֻ�ܳ�ʼ��һ��
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim cbrControlAdd As CommandBarControl
    Dim cbrMenuAdd As CommandBarControl
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    '��ʼ��
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_CA_AddIniAccount, , True)
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_CA_AddIniAccount, , True)
    '��� δ��ʼ��Ҳ��������湦�ܣ�ֻ�����˳�ʼ������������
    Set cbrControlAdd = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_CA_AddNewAccount, , True)
    Set cbrMenuAdd = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_CA_AddNewAccount, , True)
    
    gstrSQL = "select 1 from ���Ͻ���¼ where �ⷿid=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��ʼ��", cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    
    If mPrives.bln��ʼ�� = True Then
        If rsTemp.RecordCount > 0 Then
            cbrMenu.Enabled = False
            cbrControl.Enabled = False
            If mint��淽ʽ = -1 Then
                cbrControlAdd.Enabled = True
                cbrMenuAdd.Enabled = True
            End If
        Else
            cbrMenu.Enabled = True
            cbrControl.Enabled = True
            cbrControlAdd.Enabled = False
            cbrMenuAdd.Enabled = False
        End If
    Else
        cbrMenu.Visible = False
        cbrControl.Visible = False
               
        If rsTemp.RecordCount > 0 Then
            If mint��淽ʽ = -1 Then
                cbrControlAdd.Enabled = True
                cbrMenuAdd.Enabled = True
            End If
        Else
            cbrControlAdd.Enabled = False
            cbrMenuAdd.Enabled = False
        End If
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub GetPrivs()
    'Ȩ��
    mstrPrivs = GetPrivFunc(glngSys, 1729)
    
    With mPrives
        .bln���пⷿ = IsInString(mstrPrivs, "���пⷿ", ";")
        .bln��ʼ�� = IsInString(mstrPrivs, "��ʼ���", ";")
        .bln��� = IsInString(mstrPrivs, "���", ";")
    End With

End Sub


Private Sub InitAccountRec()
    '����¼��¼��
    Set mrsAccount = New ADODB.Recordset
    With mrsAccount
        If .State = 1 Then .Close
        
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "�ⷿID", adDouble, 18, adFldIsNullable
        .Fields.Append "�ⷿ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�ڳ�����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ĩ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ȡ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ȡ������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�ϴν��ID", adDouble, 10, adFldIsNullable
        .Fields.Append "�ڼ�", adLongVarChar, 6, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 2, adFldIsNullable
                
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub subPrint(ByVal intListIndex As Integer, ByVal intDetailindex As Integer, bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    Dim str�ڳ����� As String
    Dim str��ĩ���� As String

    With vsfList
        If .Row = 0 Then Exit Sub

        str�ڳ����� = "�ڳ����ڣ�" & IIf(.TextMatrix(.Row, .ColIndex("�ڳ�����")) = "", "(��ʼ���)", .TextMatrix(.Row, .ColIndex("�ڳ�����")))
        str��ĩ���� = "��ĩ���ڣ�" & .TextMatrix(.Row, .ColIndex("��ĩ����"))
    End With

    str�ڳ����� = Format(str�ڳ�����, "yyyy-mm-dd")
    str��ĩ���� = Format(str��ĩ����, "yyyy-mm-dd")

    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True

    If intDetailindex = mconTab_CA_Detail Then
        objPrint.Title.Text = "���Ľ�����"
    ElseIf intDetailindex = mconTab_CA_Stuff Then
        objPrint.Title.Text = "���Ľ����ϸ"
    ElseIf intDetailindex = mconTab_CA_Mistake Then
        objPrint.Title.Text = "���Ľ�����"
    End If

    objRow.Add str�ڳ����� & "   " & str��ĩ����
    objRow.Add "�ⷿ��" & cbo�ⷿ.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow

    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(sys.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow

    If intDetailindex = mconTab_CA_Detail Then
        Set objPrint.Body = vsfDetail
    ElseIf intDetailindex = mconTab_CA_Stuff Then
        Set objPrint.Body = vsfStuff
    ElseIf intDetailindex = mconTab_CA_Mistake Then
        If vsfMistake(0).Visible Then
            Set objPrint.Body = vsfMistake(0)
        Else
            Set objPrint.Body = vsfMistake(1)
        End If
    End If

    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
              zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub


Private Sub MediAccountProcess_AddIniAccount(ByVal int��淽ʽ As Integer)
    '��������ʼ��
    'int��淽ʽ 0-��ʼ����� 1-���
    Dim lng�ⷿID As Long
    Dim rsData As ADODB.Recordset
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    
    On Error GoTo ErrHandle
    
    'ֻ�г�ʼ������ʾ����治��ʾ
    If int��淽ʽ = 0 Then
        If MsgBox("��ʾ����ʼ�����Ե�ǰ���������Ϊ��ʼ������ݣ�����ͨ���̵�ȷ����ǰ���������ȷ��" & vbCrLf & vbCrLf & "�Ƿ����ڽ��г�ʼ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ���� = 22 And �ⷿid = [1] And ������� Is Null And Rownum = 1 "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "IsAccountTime", Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)))
    
    If Not rsData.EOF Then
        MsgBox "[" & cbo�ⷿ.Text & "]" & "�����̵㵥��δ��ˣ�����˺��ٽ��б���" & IIf(int��淽ʽ = 1, "��棡", "��ʼ����"), vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrSQL = "select 1 ��¼ from ���Ͻ���¼ where �ⷿid=[1] and ������� is null and rownum<=1"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "δ��˽��", Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)))
    
    If Not rsData.EOF Then
        MsgBox "[" & cbo�ⷿ.Text & "]" & "���н�浥��δ��ˣ�����˺��ٽ��б���" & IIf(int��淽ʽ = 1, "��棡", "��ʼ����"), vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call AviShow(Me)
    
    lng�ⷿID = Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    
    gstrSQL = "Zl_���Ͻ���¼_Insert("
    'lng�ⷿID
    gstrSQL = gstrSQL & IIf(lng�ⷿID = 0, "Null", lng�ⷿID)
    '������
    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
    '����
    gstrSQL = gstrSQL & "," & int��淽ʽ
    gstrSQL = gstrSQL & ")"

    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption & "���")

    Call GetAccountRecord
    Call RefreshList
            
    If mblnStart = True Then
        Call CheckClosAccount
    End If

    Call AviShow(Me, False)

    Exit Sub
ErrHandle:
    Call AviShow(Me, False)
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub MediAccountProcess_VerifyAccount()
    '��˽��
    Dim lng���id As Long
    
    On Error GoTo ErrHandle
    
    With vsfList
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("�������")) <> "" Then Exit Sub
        
        lng���id = Val(.TextMatrix(.Row, .ColIndex("���ID")))
    End With
    
    If lng���id = 0 Then Exit Sub

    gstrSQL = "Zl_���Ͻ���¼_Verify("
    '���ID
    gstrSQL = gstrSQL & lng���id
    '�����
    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
    gstrSQL = gstrSQL & ")"

    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption & "��˽���¼")

    Call GetAccountRecord
    Call RefreshList
    
    MsgBox "��������ϣ���鿴��", vbInformation, gstrSysName
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub MediAccountProcess_CancleAccount()
    'ȡ����浥�ݣ�ֻ�ܴ����ʼȡ������;���ݲ���ȡ��
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    With vsfList
        If .TextMatrix(.Row, .ColIndex("�������")) <> "" And .TextMatrix(.Row, .ColIndex("ȡ������")) = "" Then
            gstrSQL = "Select Max(��ĩ����) as ���� From ���Ͻ���¼ Where �ⷿid = [1] And ������� Is Not Null And ȡ���� Is Null"

            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������ѯ", Val(.TextMatrix(.Row, .ColIndex("�ⷿid"))))
            
            If rsTemp.RecordCount > 0 Then
                If rsTemp!���� = CDate(.TextMatrix(.Row, .ColIndex("��ĩ����"))) Then
                    gstrSQL = "Zl_���Ͻ���¼_Cancel("
                    '���id
                    gstrSQL = gstrSQL & .TextMatrix(.Row, .ColIndex("���id"))
                    'ȡ����
                    gstrSQL = gstrSQL & ",'" & gstrUserName & "')"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption & "���ȡ��")
                    
                    Call GetAccountRecord
                    Call RefreshList
                Else
                    MsgBox "��Ӹÿⷿ���һ�ν���¼ȡ�������һ�ν���¼��ĩ����Ϊ��(" & Format(rsTemp!����, "YYYY-MM-DD HH:MM:SS") & ")", vbInformation, gstrSysName
                End If
            End If
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub MediAccountProcess_DeleteAccount()

    On Error GoTo ErrHandle
    
    'ɾ����浥��
    With vsfList
        If .TextMatrix(.Row, .ColIndex("�������")) = "" Then
            gstrSQL = "Zl_���Ͻ���¼_Delete("
            '���id
            gstrSQL = gstrSQL & .TextMatrix(.Row, .ColIndex("���id")) & ")"
            
            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption & "���ɾ��")
            
            Call GetAccountRecord
            Call RefreshList
            
            If mblnStart = True Then
                Call CheckClosAccount
            End If
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfList_RowColChange()
    '�˵�״̬
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
            
    With vsfList
        '�ƶ���һ���ı�ǵ���ǰ�У�
        .Cell(flexcpText, 0, 0, .Rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
        End If
        
        If mPrives.bln��� Then
            '��˲˵�
            Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_CA_VerifyAccount, , True)
            Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_CA_VerifyAccount, , True)
    
            If Not cbrMenu Is Nothing Then cbrMenu.Enabled = (.TextMatrix(.Row, .ColIndex("�������")) = "")
            If Not cbrControl Is Nothing Then cbrControl.Enabled = (.TextMatrix(.Row, .ColIndex("�������")) = "")
        End If
        'ɾ���˵�
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_CA_DeleteAccount, , True)
        Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_CA_DeleteAccount, , True)

        If Not cbrMenu Is Nothing Then cbrMenu.Enabled = (.TextMatrix(.Row, .ColIndex("�������")) = "")
        If Not cbrControl Is Nothing Then cbrControl.Enabled = (.TextMatrix(.Row, .ColIndex("�������")) = "")
        'ȡ���˵�
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_CA_CancelAccount, , True)
        Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_CA_CancelAccount, , True)

        If Not cbrMenu Is Nothing Then cbrMenu.Enabled = (.TextMatrix(.Row, .ColIndex("�������")) <> "" And .TextMatrix(.Row, .ColIndex("ȡ������")) = "")
        If Not cbrControl Is Nothing Then cbrControl.Enabled = (.TextMatrix(.Row, .ColIndex("�������")) <> "" And .TextMatrix(.Row, .ColIndex("ȡ������")) = "")
        If Val(.TextMatrix(.Row, .ColIndex("����"))) = 0 Then '��ʼ�����ݲ���ȡ��
            If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
            If Not cbrControl Is Nothing Then cbrControl.Enabled = False
        End If
    End With
End Sub

'ȡ���Ľ��۸��������С��λ��
Private Function newGetDigit(ByVal int���� As Integer, ByVal int���� As Integer, Optional ByVal int��λ As Integer) As Integer
    'int���ʣ�0-���㾫��;
    'int���ݣ�1-�ɱ���;2-���ۼ�;3-����;4-���
    'int��λ�������ȡ���λ�������Բ�����ò���
    '         ���ĵ�λ:1-ɢװ;2-��װ
    '���أ���С2�����Ϊ���ݿ����С��λ��
    
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    If int���� = 4 Then 'ȡ��� ��λ=5�Ĳ��ǽ��
        int��λ = 5
    End If
    
    gstrSQL = "Select Nvl(����, 0) ���� From ҩƷ���ľ��� Where ���� = [1] And ��� = 2 And ���� = [2] And ��λ = [3] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ" & Choose(int����, "�ɱ���", "���ۼ�", "����") & "С��λ��", int����, int����, int��λ)
    
    If rsTmp.RecordCount > 0 Then
        newGetDigit = rsTmp!����
    End If
    
    If newGetDigit = 0 Then
        '���û�����þ��ȣ���ȡ���ݿ���������λ��
        newGetDigit = Choose(int����, gtype_UserDrugDigits.Digit_�ɱ���, gtype_UserDrugDigits.Digit_���ۼ�, gtype_UserDrugDigits.Digit_����, gtype_UserDrugDigits.Digit_���)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    newGetDigit = Choose(int����, gtype_UserDrugDigits.Digit_�ɱ���, gtype_UserDrugDigits.Digit_���ۼ�, gtype_UserDrugDigits.Digit_����, gtype_UserDrugDigits.Digit_���)
End Function



