VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frm��Ʊ��Ϣ��¼ 
   Caption         =   "��Ʊ��Ϣ��¼"
   ClientHeight    =   11265
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19980
   Icon            =   "frm��Ʊ��Ϣ��¼.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11265
   ScaleWidth      =   19980
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   10440
      ScaleHeight     =   255
      ScaleWidth      =   3375
      TabIndex        =   37
      Top             =   9023
      Width           =   3375
      Begin VB.PictureBox picԤ�� 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1820
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   41
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picͣ�� 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   40
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox pic��Ч�� 
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   910
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   39
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox pic��治�� 
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2730
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   38
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   2100
         TabIndex        =   45
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "ҩƷ"
         Height          =   180
         Index           =   2
         Left            =   285
         TabIndex        =   44
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1200
         TabIndex        =   43
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "�豸"
         Height          =   180
         Index           =   4
         Left            =   3015
         TabIndex        =   42
         Top             =   37
         Width           =   360
      End
   End
   Begin VB.PictureBox pic������Ϣ 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3255
      ScaleWidth      =   3375
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4080
      Width           =   3375
      Begin VB.CommandButton cmdҩƷ 
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2820
         TabIndex        =   18
         Top             =   1035
         Width           =   255
      End
      Begin VB.CheckBox chkDept 
         BackColor       =   &H8000000B&
         Caption         =   "����(&W)"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   14
         Tag             =   "4"
         Top             =   240
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         BackColor       =   &H8000000B&
         Caption         =   "ҩƷ(&D)"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Tag             =   "1"
         Top             =   240
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         BackColor       =   &H8000000B&
         Caption         =   "����(&M)"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Tag             =   "2"
         Top             =   600
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         BackColor       =   &H8000000B&
         Caption         =   "�豸(&S)"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   16
         Tag             =   "4"
         Top             =   600
         Width           =   1035
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   300
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1470
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker dtp���ƽ���ʱ�� 
         Height          =   315
         Left            =   1380
         TabIndex        =   25
         Top             =   2340
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   193003523
         CurrentDate     =   43522
      End
      Begin MSComCtl2.DTPicker dtp���ƿ�ʼʱ�� 
         Height          =   315
         Left            =   1380
         TabIndex        =   23
         Top             =   1905
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   193003523
         CurrentDate     =   43522
      End
      Begin VB.TextBox txt��Ŀ���� 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   19
         Top             =   1035
         Width           =   1725
      End
      Begin VB.Label lblҩƷ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Ʒ    ��"
         Height          =   180
         Left            =   360
         TabIndex        =   17
         Top             =   1095
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   360
         TabIndex        =   20
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl���ƿ�ʼ���� 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����"
         Height          =   180
         Left            =   360
         TabIndex        =   22
         Top             =   1965
         Width           =   720
      End
      Begin VB.Label lbl���ƽ������� 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   360
         TabIndex        =   24
         Top             =   2400
         Width           =   720
      End
   End
   Begin VB.PictureBox pic������Ϣ 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2295
      ScaleWidth      =   3375
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3375
      Begin VB.ComboBox cbo������� 
         Height          =   300
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   570
         Width           =   1725
      End
      Begin VB.CommandButton Cmd��Ӧ�� 
         Caption         =   "��"
         Height          =   300
         Left            =   2820
         TabIndex        =   5
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox txt��Ӧ�� 
         Height          =   300
         Left            =   1380
         TabIndex        =   4
         Top             =   120
         Width           =   1485
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   315
         Left            =   1380
         TabIndex        =   9
         Top             =   1050
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   193003523
         CurrentDate     =   40848
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   315
         Left            =   1380
         TabIndex        =   11
         Top             =   1538
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   193003523
         CurrentDate     =   40848
      End
      Begin VB.Label lbl��ʼ���� 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   360
         TabIndex        =   10
         Top             =   1605
         Width           =   720
      End
      Begin VB.Label lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   360
         TabIndex        =   6
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl��Ӧ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "�� Ӧ ��"
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.PictureBox picDetails 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   3480
      ScaleHeight     =   5295
      ScaleWidth      =   13335
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   13335
      Begin VB.Frame fra������Ϣ 
         Height          =   855
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   13335
         Begin VB.TextBox txt��Ʊ���� 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3390
            TabIndex        =   31
            Top             =   300
            Width           =   1485
         End
         Begin VB.TextBox txt��Ʊ�� 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            TabIndex        =   29
            Top             =   300
            Width           =   1485
         End
         Begin MSComCtl2.DTPicker dtp��Ʊ���� 
            Height          =   315
            Left            =   5955
            TabIndex        =   33
            Top             =   300
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   193003523
            CurrentDate     =   40848
         End
         Begin VB.Label lblǿ�� 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   90
         End
         Begin VB.Label lbl��Ʊ���� 
            AutoSize        =   -1  'True
            Caption         =   "��Ʊ����"
            Height          =   180
            Left            =   5174
            TabIndex        =   32
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lbl��Ʊ���� 
            AutoSize        =   -1  'True
            Caption         =   "��Ʊ����"
            Height          =   180
            Left            =   2617
            TabIndex        =   30
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lbl��Ʊ�� 
            AutoSize        =   -1  'True
            Caption         =   "��Ʊ��"
            Height          =   180
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lbl��ʾ 
            Caption         =   $"frm��Ʊ��Ϣ��¼.frx":6852
            ForeColor       =   &H000000FF&
            Height          =   540
            Left            =   8520
            TabIndex        =   46
            Top             =   187
            Width           =   4695
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2805
         Left            =   0
         TabIndex        =   36
         Top             =   1200
         Width           =   12060
         _cx             =   21272
         _cy             =   4948
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   315
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm��Ʊ��Ϣ��¼.frx":68F3
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
         ExplorerBar     =   5
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
   Begin XtremeSuiteControls.TaskPanel tkpMain 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   3375
      _Version        =   589884
      _ExtentX        =   5953
      _ExtentY        =   1508
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   35
      Top             =   10905
      Width           =   19980
      _ExtentX        =   35243
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   2356
            Picture         =   "frm��Ʊ��Ϣ��¼.frx":6968
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   30154
            MinWidth        =   600
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
   Begin VSFlex8Ctl.VSFlexGrid mshSelect 
      Height          =   2535
      Left            =   3480
      TabIndex        =   26
      Top             =   6240
      Width           =   4695
      _cx             =   8281
      _cy             =   4471
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
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
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
   Begin XtremeCommandBars.ImageManager imgPicture 
      Left            =   1320
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frm��Ʊ��Ϣ��¼.frx":71FC
   End
   Begin XtremeDockingPane.DockingPane dkpPanel 
      Left            =   720
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frm��Ʊ��Ϣ��¼"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const menuToolSave As Integer = 101
Private Const menuToolGetData As Integer = 102
Private Const menuToolExit As Integer = 103
Private Const menuToolCheck As Integer = 104
Private Const menuToolCheckCancel As Integer = 105
Private Const menuToolHelp As Integer = 108

Private Enum mColumn
    ѡ�� = 0
    No = 1
    �ⷿ
    ���
    ��¼״̬
    ��Ŀid
    ��Ŀ��Ϣ
    ���
    ����
    ��λ
    ����
    �ɹ���
    �ɹ����
    ��Ʊ��
    ��Ʊ����
    ��Ʊ���
    ��ʶ
    �շ�ID
    �������
    Count = 19
    
End Enum
'��������ҩƷ�����ġ����ʺ��豸
Private Const glngColor���� As Long = &HC00000
Private Const glngColorҩƷ As Long = &HC0
Private Const glngColor���� As Long = &H8000&
Private Const glngColor�豸 As Long = &HC000C0
'���ڿɷ�༭��ɫ����
Private Const glngColorGray = &H80000004
Private Const glngColorWhite = &H80000005
Private Const gintҩƷIndex = 0
Private Const gint����Index = 1
Private Const gint����Index = 2
Private Const gint�豸Index = 3
Private mfrmMain As Form
Private mstrPrivs As String
Private mstr��Ӧ��Type As String 'ҩƷ�����ġ����ʡ��豸
Private mstrSelectTag As String
Private Const mintShowPriceDigit = 5           '�۸�С��λ��
Private Const mintShowAmountDigit = 5         '���С��λ��
Private mbln������ As Boolean '��ѡ�ķ�Ʊ���ϼ��Ƿ���Ҫ���£��޸ķ���������Ҫ����
Private mbln�����־ As Boolean

Private Sub cbo�������_Click()
    Dim dateCurrentDate As Date
    
    If cbo�������.Text = "�Զ�������" Then
        dtp��ʼʱ��.Enabled = True
        dtp����ʱ��.Enabled = True
        
    Else
        dtp��ʼʱ��.Enabled = False
        dtp����ʱ��.Enabled = False
    End If
    
    '����ѡ��ı�ʱ��
    dateCurrentDate = Sys.Currentdate
    Select Case cbo�������.ListIndex
        Case 0, 1
            dtp��ʼʱ��.Value = CDate(Format(dateCurrentDate, "yyyy-mm-dd") & " 00:00:00")
            dtp����ʱ��.Value = dateCurrentDate
        Case 2
            dtp��ʼʱ��.Value = CDate(Format(DateAdd("d", -7, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp����ʱ��.Value = dateCurrentDate
        Case 3
            dtp��ʼʱ��.Value = CDate(Format(DateAdd("d", -30, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp����ʱ��.Value = dateCurrentDate
        Case 4
            dtp��ʼʱ��.Value = CDate(Format(DateAdd("d", -90, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp����ʱ��.Value = dateCurrentDate
    End Select
End Sub



Private Sub cbo��������_Click()
    Dim dateCurrentDate As Date
    
    If cbo��������.Text = "�Զ�������" Then
        dtp���ƿ�ʼʱ��.Enabled = True
        dtp���ƽ���ʱ��.Enabled = True
        
    Else
        dtp���ƿ�ʼʱ��.Enabled = False
        dtp���ƽ���ʱ��.Enabled = False
    End If
    
    '����ѡ��ı�ʱ��
    dateCurrentDate = Sys.Currentdate
    Select Case cbo��������.ListIndex
        Case 0, 1
            dtp���ƿ�ʼʱ��.Value = CDate(Format(dateCurrentDate, "yyyy-mm-dd") & " 00:00:00")
            dtp���ƽ���ʱ��.Value = dateCurrentDate
        Case 2
            dtp���ƿ�ʼʱ��.Value = CDate(Format(DateAdd("d", -7, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp���ƽ���ʱ��.Value = dateCurrentDate
        Case 3
            dtp���ƿ�ʼʱ��.Value = CDate(Format(DateAdd("d", -30, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp���ƽ���ʱ��.Value = dateCurrentDate
        Case 4
            dtp���ƿ�ʼʱ��.Value = CDate(Format(DateAdd("d", -90, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp���ƽ���ʱ��.Value = dateCurrentDate
    End Select
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case menuToolGetData '��ȡ����
            If ValidData = False Then Exit Sub
            GetData
        Case menuToolSave '��������
            If Not SaveCard Then Exit Sub '����ʧ���˳�
            txt��Ʊ��.SetFocus: txt��Ʊ��.Text = "": txt��Ʊ����.Text = ""
            '������ȡ����
            If ValidData = False Then Exit Sub
            GetData
        Case menuToolCheck  'ȫѡ
            cbsCheck
        Case menuToolCheckCancel   'ȫ��
            cbsCheckCancel
        Case menuToolHelp '����
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
        Case menuToolExit '�˳�
            Unload Me
    End Select
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If mbln������ Then
        AmountSum
        mbln������ = False
    End If
End Sub

Private Sub chkDept_Click(Index As Integer)
    Dim intSum As Integer
    Dim i As Integer
    
    For i = chkDept.LBound To chkDept.UBound
        If chkDept(i).Value = 1 Then intSum = intSum + 1
    Next
    
    If intSum = 1 Then
        txt��Ŀ����.Enabled = True
        cmdҩƷ.Enabled = True
        txt��Ŀ����.BackColor = glngColorWhite
    Else
        txt��Ŀ����.Enabled = False
        cmdҩƷ.Enabled = False
        txt��Ŀ����.BackColor = glngColorGray
    End If
    
    txt��Ŀ����.Text = ""
    txt��Ŀ����.Tag = ""
End Sub

Private Sub cbsCheck()
    Dim i As Integer
    
    With vsfList
        For i = 1 To .Rows - 1
            .TextMatrix(i, mColumn.ѡ��) = "��"
            .TextMatrix(i, mColumn.��Ʊ��) = txt��Ʊ��.Text
            .TextMatrix(i, mColumn.��Ʊ����) = txt��Ʊ����.Text
        Next
    End With
    
    AmountSum
End Sub

Private Sub cbsCheckCancel()
    Dim i As Integer
    
    With vsfList
        For i = 1 To .Rows - 1
            .TextMatrix(i, mColumn.ѡ��) = ""
            .TextMatrix(i, mColumn.��Ʊ��) = ""
            .TextMatrix(i, mColumn.��Ʊ����) = ""
        Next
    End With
    
    AmountSum
End Sub

Private Sub Cmd��Ӧ��_Click()
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    strTemp = frm��Ӧ��ѡ��.SelDept(mstrPrivs)
    If strTemp = "" Then
        Unload frm��Ӧ��ѡ��
        If txt��Ӧ��.Enabled Then txt��Ӧ��.SetFocus
        Exit Sub
    End If
    txt��Ӧ��.Text = Mid(strTemp, InStr(strTemp, ",") + 1)
    txt��Ӧ��.Tag = Val(Left(strTemp, InStr(strTemp, ",") - 1))
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord("select ���� from ��Ӧ�� where id=[1] ", Caption & "-��ȡ��Ӧ������", txt��Ӧ��.Tag)
    If Not rsTemp.EOF Then
        mstr��Ӧ��Type = Nvl(rsTemp!����)
    End If
    rsTemp.Close
    Call SetClass
    
    zlCommFun.PressKey vbKeyTab
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdҩƷ_Click()
    Call GetItem("")
    zlCommFun.PressKey vbKeyTab
End Sub


Private Sub Form_Load()
    Call initComandbar  '��ʼ��������
    Call InitTask  '��ʼ�����
    Call initComboBox
    Call initColumn
    dtp��Ʊ����.Value = Sys.Currentdate
    mbln�����־ = Val(zlDatabase.GetPara("�⹺�����Ҫ������Ǹ������ܽ��и������", glngSys, 0)) = 1
    
    RestoreWinState Me, App.ProductName
    stbThis.Panels(2).Picture = picColor
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
    
    With cbrToolBar.Controls    '
        Set cbrControlMain = .Add(xtpControlButton, menuToolGetData, "��ȡ����")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, menuToolSave, "ȷ��")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, menuToolCheck, "ȫѡ")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, menuToolCheckCancel, "ȫ��")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, menuToolHelp, "����")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, menuToolExit, "�˳�")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        
    End With
    '�����
    With Me.cbsMain.KeyBindings
        .Add 0, VK_F5, menuToolGetData
        .Add 0, VK_ESCAPE, menuToolExit
    End With
    
    
    cbsMain.Item(1).Delete
End Sub

Private Sub InitTask()
'---------------------------------------
'��ʼ���������
'----------------------------------------
    Dim objGroup As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
   
    Call tkpMain.SetMargins(0, 0, 0, 0, 0)
    Call tkpMain.SetItemInnerMargins(0, 0, 0, 0)
    Call tkpMain.SetItemOuterMargins(0, 0, 0, 0)
    Call tkpMain.SetGroupInnerMargins(0, 0, 0, 0)
    Call tkpMain.SetGroupOuterMargins(3, 3, 3, 0)
        
    Set objGroup = tkpMain.Groups.Add(1, "��������")
    objGroup.Expandable = False '��������
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic������Ϣ
    pic������Ϣ.BackColor = objItem.BackColor
   
    Set objGroup = tkpMain.Groups.Add(2, "��������")
    objGroup.Expandable = True
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic������Ϣ
    pic������Ϣ.BackColor = objItem.BackColor
    objGroup.Expanded = False  'û�д�

End Sub

Private Sub initComboBox()
    With cbo��������
        .Clear
        .AddItem ""
        .AddItem "����"
        .AddItem "һ������"
        .AddItem "һ������"
        .AddItem "��������"
        .AddItem "�Զ�������"
    End With
    
    With cbo�������
        .Clear
        .AddItem ""
        .AddItem "����"
        .AddItem "һ������"
        .AddItem "һ������"
        .AddItem "��������"
        .AddItem "�Զ�������"
        .ListIndex = 1
    End With
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.tkpMain.Move 0, 530, Me.tkpMain.Width, Me.ScaleHeight - stbThis.Height - 530
    Me.picDetails.Move tkpMain.Width, 530, Me.ScaleWidth - tkpMain.Width, tkpMain.Height
    fra������Ϣ.Move 0, 0, picDetails.Width, fra������Ϣ.Height
    lbl��ʾ.Left = fra������Ϣ.Width - lbl��ʾ.Width - 50
    
    vsfList.Move 0, fra������Ϣ.Height, picDetails.Width, picDetails.Height - fra������Ϣ.Height
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - stbThis.Panels(3).Width - stbThis.Panels(4).Width - .Width - 400
    End With
End Sub


Public Sub ShowCard(frmMain As Form, ByVal strPrivs As String)
    
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain

    Me.Show vbModal, frmMain
End Sub

Private Sub SetClass()
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    For i = 0 To chkDept.Count - 1
        'ϵͳ
        If i >= 2 Then
            Set rsTemp = zlDatabase.OpenSQLRecord("Select Count(1) Rec From zlSystems Where ��� = [1]", Caption, IIf(i = 2, 400, 600))
            chkDept(i).Enabled = rsTemp!rec > 0
            rsTemp.Close
        Else
            chkDept(i).Enabled = True
        End If
        'Ȩ��
        Select Case i
            Case 0
                chkDept(i).Enabled = chkDept(i).Enabled And InStr(mstrPrivs, ";ҩƷ;") > 0
            Case 1
                chkDept(i).Enabled = chkDept(i).Enabled And InStr(mstrPrivs, ";����;") > 0
            Case 2
                chkDept(i).Enabled = chkDept(i).Enabled And InStr(mstrPrivs, ";����;") > 0
            Case 3
                chkDept(i).Enabled = chkDept(i).Enabled And InStr(mstrPrivs, ";�豸;") > 0
        End Select
    Next
    '��Ӧ��
    If Len(mstr��Ӧ��Type) >= 1 Then  'ҩƷ
        chkDept(gintҩƷIndex).Enabled = chkDept(gintҩƷIndex).Enabled And Mid(mstr��Ӧ��Type, 1, 1) = "1"
    Else
        chkDept(gintҩƷIndex).Enabled = False
    End If
    If Len(mstr��Ӧ��Type) >= 5 Then  '����
        chkDept(gint����Index).Enabled = chkDept(gint����Index).Enabled And Mid(mstr��Ӧ��Type, 5, 1) = "1"
    Else
        chkDept(gint����Index).Enabled = False
    End If
    If Len(mstr��Ӧ��Type) >= 2 Then  '����
        chkDept(gint����Index).Enabled = chkDept(gint����Index).Enabled And Mid(mstr��Ӧ��Type, 2, 1) = "1"
    Else
        chkDept(gint����Index).Enabled = False
    End If
    If Len(mstr��Ӧ��Type) >= 3 Then  '�豸
        chkDept(gint�豸Index).Enabled = chkDept(gint�豸Index).Enabled And Mid(mstr��Ӧ��Type, 3, 1) = "1"
    Else
        chkDept(gint�豸Index).Enabled = False
    End If
    
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub

Private Sub GetData()
    Dim rsRecord As ADODB.Recordset
    Dim str���� As String
    Dim strSQL As String
    Dim str������SQL As String
    Dim str�豸��SQL As String
    Dim strҩƷ������SQL As String
    Dim str���SQL As String
    
    On Error GoTo errHandle
    
    Me.MousePointer = vbHourglass
    
    If cbo��������.Text <> "" And cbo�������.Text <> "" Then '����������ڶ���Ϊ��
        strSQL = " And ((x.�������� between [2] and [3] And x.������� is Null) Or x.������� between [4] and [5] ) "
    ElseIf cbo��������.Text <> "" Then
        strSQL = " And x.�������� between [2] and [3] And x.������� is Null "
    ElseIf cbo�������.Text <> "" Then
        strSQL = " And x.������� between [4] and [5] "
    End If
    
    '��Ҫ����ҩƷ������
    If chkDept(gintҩƷIndex).Value = 1 Or chkDept(gint����Index).Value = 1 Or _
        (chkDept(gintҩƷIndex).Value = 1 And chkDept(gint����Index).Value = 1 And chkDept(gint����Index).Value = 1 And chkDept(gint�豸Index).Value = 1) Or _
        (chkDept(gintҩƷIndex).Value <> 1 And chkDept(gint����Index).Value <> 1 And chkDept(gint����Index).Value <> 1 And chkDept(gint�豸Index).Value <> 1) Then
        
        '��Ҫ��ѯ��Щ����
        If chkDept(gintҩƷIndex).Value = 1 Then str���� = "1"
        If chkDept(gint����Index).Value = 1 Then str���� = IIf(str���� = "", "", str���� & ",") & "15"
        If str���� = "" Then str���� = "1,15" '��δ��ѡ�����Ƕ���ѡ
        
        If txt��Ŀ����.Text <> "" Then strSQL = strSQL & " And x.ҩƷID = [6]"
        '��ȡ��¼״̬���ã��ж��Ƿ񱻳�����������Min(x.��¼״̬)�ķ���
        'ȡԭʼ���ݵ��շ�ID����Min(x.id)�ķ���
        strҩƷ������SQL = "Select Distinct a.No ��ⵥ�ݺ�, a.���, a.��¼״̬, a.ҩƷid ��ĿID, '[' || d.���� || ']' || d.���� As ��Ŀ��Ϣ, d.���, a.����, d.���㵥λ As ��λ," & vbNewLine & _
                        "                       a.��д���� As ����, a.�ɱ��� * 1 As �ɹ���, a.�ɱ���� As �ɹ����, e.���� �ⷿ,decode(a.����,1,1,15,5) ��ʶ,a.�շ�ID,Null �������" & vbNewLine & _
                        "       From (Select x.No, Min(x.��¼״̬) ��¼״̬, Sum(ʵ������) As ��д����, Sum(�ɱ����) As �ɱ����, x.ҩƷid, x.���," & vbNewLine & _
                        "                     x.����, x.�ɱ���,  x.��ҩ��λid, x.�ⷿid,x.����,Min(x.id) �շ�ID " & vbNewLine & _
                        "              From ҩƷ�շ���¼ X" & vbNewLine & _
                        "              Where Not Exists" & vbNewLine & _
                        "               (Select 1" & vbNewLine & _
                        "                     From Ӧ����¼ Y" & vbNewLine & _
                        "                     Where x.Id = y.�շ�id And y.ϵͳ��ʶ In (1, 5) And y.��¼���� = 0 And y.��Ʊ�� Is Not Null)" & vbNewLine & _
                        "                     And ���� in (" & str���� & ") " & vbNewLine & _
                        "             " & strSQL & vbNewLine & _
                        "              Group By x.No, x.ҩƷid, x.���, x.����, x.�ɱ���, x.��ҩ��λid, x.�ⷿid,x.����" & vbNewLine & _
                        "              Having Sum(ʵ������) <> 0) A, �շ���ĿĿ¼ D, ���ű� E, ��Ӧ�� F" & vbNewLine & _
                        "       Where a.ҩƷid = d.Id And a.��ҩ��λid  + 0 = f.Id And a.�ⷿid = e.Id And (Substr(f.����, 1, 1) = 1 or Substr(f.����, 5, 1) = 1) and f.id = [1]"

    End If
    
    '��Ҫ��������:����ϵͳ���밲װ
    If chkDept(gint����Index).Enabled = True And (chkDept(gint����Index).Value = 1 Or (chkDept(gintҩƷIndex).Value = 1 And chkDept(gint����Index).Value = 1 And chkDept(gint����Index).Value = 1 And chkDept(gint�豸Index).Value = 1) Or _
    (chkDept(gintҩƷIndex).Value <> 1 And chkDept(gint����Index).Value <> 1 And chkDept(gint����Index).Value <> 1 And chkDept(gint�豸Index).Value <> 1)) Then
        
        If txt��Ŀ����.Text <> "" Then strSQL = strSQL & " And x.����id = [6]"
        '��ȡ��¼״̬���ã��ж��Ƿ񱻳�����������Min(x.��¼״̬)�ķ���
        'ȡԭʼ���ݵ��շ�ID����Min(x.id)�ķ���
        str������SQL = "Select Distinct a.No ��ⵥ�ݺ�, a.���, a.��¼״̬, a.����id ��ĿID, '[' || d.���� || ']' || d.���� As ��Ŀ��Ϣ, d.���, a.����, d.ɢװ��λ As ��λ," & vbNewLine & _
                    "                       a.��д���� As ����, a.�ɱ��� * 1 As �ɹ���, a.�ɱ���� As �ɹ����, e.���� �ⷿ,2 ��ʶ,a.�շ�ID,Null �������" & vbNewLine & _
                    "       From (Select x.No, Min(x.��¼״̬) ��¼״̬, Sum(ʵ������) As ��д����, Sum(���) As �ɱ����, x.����id, x.���," & vbNewLine & _
                    "                     x.����, x.���� �ɱ���,  x.������λid, x.�ⷿid,x.����,Min(x.id) �շ�ID" & vbNewLine & _
                    "              From �����շ���¼ X " & vbNewLine & _
                    "              Where Not Exists" & vbNewLine & _
                    "               (Select 1" & vbNewLine & _
                    "                     From Ӧ����¼ Y" & vbNewLine & _
                    "                     Where x.Id = y.�շ�id And y.ϵͳ��ʶ = 2 And y.��¼���� In (0, -1) And y.��Ʊ�� Is Not Null)" & vbNewLine & _
                    "                     And x.���� = 1 " & vbNewLine & _
                    "             " & strSQL & vbNewLine & _
                    "              Group By x.No, x.����id, x.���, x.����, x.����, x.������λid, x.�ⷿid,x.����" & vbNewLine & _
                    "              Having Sum(ʵ������) <> 0) A, ����Ŀ¼ D, ���ű� E, ��Ӧ�� F" & vbNewLine & _
                    "       Where a.����id = d.Id And a.������λid + 0 = f.Id And a.�ⷿid = e.Id And  Substr(f.����, 2, 1) = 1 and f.id = [1]"

    End If
    
    '��Ҫ�����豸:�豸ϵͳ���밲װ
    If chkDept(gint�豸Index).Enabled = True And (chkDept(gint�豸Index).Value = 1 Or (chkDept(gintҩƷIndex).Value = 1 And chkDept(gint����Index).Value = 1 And chkDept(gint����Index).Value = 1 And chkDept(gint�豸Index).Value = 1) Or _
    (chkDept(gintҩƷIndex).Value <> 1 And chkDept(gint����Index).Value <> 1 And chkDept(gint����Index).Value <> 1 And chkDept(gint�豸Index).Value <> 1)) Then
        
        If txt��Ŀ����.Text <> "" Then strSQL = strSQL & " And x.�豸id = [6]"
        '��ȡ��¼״̬���ã��ж��Ƿ񱻳�����������Min(x.��¼״̬)�ķ���
        'ȡԭʼ���ݵ��շ�ID����Min(x.id)�ķ���
        str�豸��SQL = "Select Distinct a.No ��ⵥ�ݺ�, a.���, a.��¼״̬, a.�豸id ��Ŀid, '[' || d.���� || ']' || d.���� As ��Ŀ��Ϣ, d.���, a.����, d.��λ ," & vbNewLine & _
                    "                a.��д���� As ����, a.�ɱ��� * 1 As �ɹ���, a.�ɱ���� As �ɹ����, e.���� �ⷿ, 3 ��ʶ, a.�շ�id,a.�������" & vbNewLine & _
                    "From (Select x.No, Min(x.��¼״̬) ��¼״̬, Sum(ʵ������) As ��д����, Sum(���) As �ɱ����, x.�豸id, x.���, Null ����, x.���� �ɱ���, x.������λid, x.�ⷿid," & vbNewLine & _
                    "              x.����, Min(x.Id) �շ�ID,x.�������" & vbNewLine & _
                    "       From �豸�շ���¼ X " & vbNewLine & _
                    "              Where Not Exists" & vbNewLine & _
                    "               (Select 1" & vbNewLine & _
                    "                     From Ӧ����¼ Y" & vbNewLine & _
                    "                     Where x.Id = y.�շ�id And y.ϵͳ��ʶ = 3 And y.��¼���� In (0, -1) And y.��Ʊ�� Is Not Null)" & vbNewLine & _
                    "                     And x.���� =1 " & vbNewLine & _
                    "             " & strSQL & vbNewLine & _
                    "       Group By x.No, x.�豸id, x.���, x.����, x.����, x.������λid, x.�ⷿid, x.����,x.�������" & vbNewLine & _
                    "       Having Sum(ʵ������) <> 0) A, �豸Ŀ¼ D, ���ű� E, ��Ӧ�� F" & vbNewLine & _
                    "Where a.�豸id = d.Id And a.������λid + 0 = f.Id And a.�ⷿid = e.Id And Substr(f.����, 3, 1) = 1 and f.id = [1]"


    End If
    
    If strҩƷ������SQL <> "" And str������SQL <> "" And str�豸��SQL <> "" Then
        str���SQL = strҩƷ������SQL & vbNewLine & " Union All " & vbNewLine & str������SQL & vbNewLine & " Union All" & vbNewLine & str�豸��SQL & vbNewLine
    ElseIf strҩƷ������SQL <> "" And str������SQL <> "" Then
        str���SQL = strҩƷ������SQL & vbNewLine & " Union All " & vbNewLine & str������SQL & vbNewLine
    ElseIf strҩƷ������SQL <> "" And str�豸��SQL <> "" Then
        str���SQL = strҩƷ������SQL & vbNewLine & " Union All" & vbNewLine & str�豸��SQL & vbNewLine
    ElseIf str������SQL <> "" And str�豸��SQL <> "" Then
        str���SQL = str������SQL & vbNewLine & " Union All" & vbNewLine & str�豸��SQL & vbNewLine
    ElseIf strҩƷ������SQL <> "" Then
        str���SQL = strҩƷ������SQL & vbNewLine
    ElseIf str������SQL <> "" Then
        str���SQL = str������SQL & vbNewLine
    ElseIf str�豸��SQL <> "" Then
        str���SQL = str�豸��SQL & vbNewLine
    End If
    
    
    gstrSQL = "Select * " & vbNewLine & _
            "   From ( " & vbNewLine & _
            "  " & str���SQL & _
            ")" & vbNewLine & _
            "Order By ��ⵥ�ݺ�,�ⷿ,��� Asc"
    
    

    Set rsRecord = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", Val(txt��Ӧ��.Tag), CDate(Format(dtp���ƿ�ʼʱ��, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(dtp���ƽ���ʱ��, "yyyy-mm-dd") & " 23:59:59"), CDate(Format(dtp��ʼʱ��, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(dtp����ʱ��, "yyyy-mm-dd") & " 23:59:59"), Val(txt��Ŀ����.Tag))
    
    SetColumn rsRecord
    Me.MousePointer = vbDefault
    
    stbThis.Panels(2).Text = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ValidData() As Boolean
    
    If Val(txt��Ӧ��.Tag) = 0 Then
        ShowMsgbox "��Ӧ��δѡ�񣬲��ܼ�����"
        If txt��Ӧ��.Enabled Then txt��Ӧ��.SetFocus
        Exit Function
    End If
    
    If Trim(cbo�������.Text) = "" And Trim(cbo��������.Text) = "" Then
        ShowMsgbox "������ں��������ڶ�Ϊ�գ����ܼ�����"
        If cbo�������.Enabled Then cbo�������.SetFocus
        Exit Function
    End If
    
    ValidData = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Cancel = 1
        Exit Sub
    End If
    
    SaveWinState Me, App.ProductName
    'ж�ش������
    If Not mfrmMain Is Nothing Then
        Set mfrmMain = Nothing
    End If
End Sub

Private Sub mshSelect_DblClick()
    With mshSelect
        If .Row > 0 And .TextMatrix(.Row, 0) <> "" Then
            mshSelect_KeyPress 13
        End If
    End With
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    With mshSelect
        Select Case mstrSelectTag
            Case "Provide"
                If KeyAscii = vbKeyReturn Then
                    If .Row = 0 Then Exit Sub
                    txt��Ӧ��.Text = "[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2)
                    
                    mstr��Ӧ��Type = .TextMatrix(.Row, 4)
                    Call SetClass
                    
                    txt��Ӧ��.Tag = Val(.TextMatrix(.Row, 0))
                    zlCommFun.PressKey vbKeyTab
                ElseIf KeyAscii = 27 Then
                    If txt��Ӧ��.Enabled Then txt��Ӧ��.SetFocus
                End If
            Case Else
        End Select
        .Visible = False
    End With
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

Private Sub txt��Ʊ����_GotFocus()
    txt��Ʊ����.SelStart = 0
    txt��Ʊ����.SelLength = Len(txt��Ʊ����.Text)
End Sub

Private Sub txt��Ʊ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtp��Ʊ����.SetFocus
End Sub

Private Sub txt��Ʊ����_LostFocus()
    Dim i As Integer
    
    With vsfList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, mColumn.ѡ��) = "��" Then .TextMatrix(i, mColumn.��Ʊ����) = txt��Ʊ����.Text
        Next
    End With
End Sub

Private Sub txt��Ʊ��_GotFocus()
    txt��Ʊ��.SelStart = 0
    txt��Ʊ��.SelLength = Len(txt��Ʊ��.Text)
End Sub

Private Sub txt��Ʊ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt��Ʊ����.SetFocus
End Sub

Private Sub txt��Ʊ��_LostFocus()
    Dim i As Integer
    
    With vsfList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, mColumn.ѡ��) = "��" Then .TextMatrix(i, mColumn.��Ʊ��) = txt��Ʊ��.Text
        Next
    End With
    
End Sub

Private Sub txt��Ӧ��_GotFocus()
    txt��Ӧ��.SelStart = 0
    txt��Ӧ��.SelLength = Len(txt��Ӧ��.Text)
End Sub

Private Sub txt��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If SelMltProvide = False And mshSelect.Visible = False Then
            If txt��Ӧ��.Enabled Then txt��Ӧ��.SetFocus: txt��Ӧ��.SelStart = 0: txt��Ӧ��.SelLength = Len(txt��Ӧ��.Text)
        Else
            If mshSelect.Visible = False Then
                zlCommFun.PressKey vbKeyTab
            End If
        End If
    End If
End Sub

Private Function SelMltProvide() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ��Ӧ������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String
    Dim strȨ�� As String
    
    If Trim(txt��Ӧ��.Text) = "" Then Exit Function
    
    strTmp = GetMatchingSting(UCase(txt��Ӧ��.Text), False)
    
    strȨ�� = " and " & Get����Ȩ��(mstrPrivs)
    
    SelMltProvide = False
    
    strSQL = "" & _
        "  Select   ID,����,����,����,����" & _
        "  From  ��Ӧ�� " & _
        "  Where (����ʱ�� is null or To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01') " & _
        "       " & zl_��ȡվ������ & "  and ĩ��=1  " & _
        "       And ( ���� Like [1] or ���� like [1] or ����  like upper([1])) " & strȨ��
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTmp)
    
    If rsTemp.EOF Then
        ShowMsgbox "δ�ҵ�ָ���Ĺ�Ӧ��!"
        Exit Function
    End If
    With rsTemp
        If .RecordCount > 1 Then
            mstrSelectTag = "Provide"
            Set mshSelect.DataSource = rsTemp
            With mshSelect
                .Top = tkpMain.Top + pic������Ϣ.Top + txt��Ӧ��.Top + txt��Ӧ��.Height + 10
                .Left = pic������Ϣ.Left + txt��Ӧ��.Left
                .Visible = True
                .ColWidth(0) = 0
                .ColWidth(1) = 800
                .ColWidth(2) = 2000
                .ColWidth(3) = 800
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
                .ZOrder
                .SetFocus
                Exit Function
            End With
        Else
            txt��Ӧ��.Text = "[" & Nvl(rsTemp!����) & "]" & rsTemp!����
            txt��Ӧ��.Tag = Nvl(rsTemp!ID, 0)
            mstr��Ӧ��Type = rsTemp!����
            Call SetClass
            SelMltProvide = True
        End If
    End With
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub GetItem(ByVal strkey As String)
    Dim intClass As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    Dim sngX As Single, sngY As Single, sngH As Single
    Dim intSysParam As Integer
    Dim strMatch As String
    
    intClass = GetClassValue()
    vRect = zlControl.GetControlRect(txt��Ŀ����.hwnd)
    sngX = vRect.Left
    sngY = vRect.Bottom
    
    On Error GoTo errHandle
    Select Case intClass
    Case 0
        'ҩƷ
        If strkey = "" Then
            strSQL = "Select ID, �ϼ�id, ����, ����, '' ���, '' ����, '' ҩ�ⵥλ, '' סԺ��λ, '' ���ﵥλ, 0 As ĩ�� " & _
                     "From ���Ʒ���Ŀ¼ " & vbLf & _
                     "Where ���� in ('1','2','3') " & vbLf & _
                     "Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id " & vbLf & _
                     "Union all " & vbLf & _
                     "Select a.Id, c.����id As �ϼ�id, a.����, a.����, a.���, a.����, b.ҩ�ⵥλ, b.סԺ��λ, b.���ﵥλ, 1 As ĩ�� " & vbLf & _
                     "From �շ���ĿĿ¼ A, ҩƷ��� B, ������ĿĿ¼ C " & vbLf & _
                     "Where a.Id = b.ҩƷid And b.ҩ��id = c.Id And a.��� in ('5','6','7') " & vbLf & _
                     "  And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-ҩƷ" _
                    , False, "", "ѡ��", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select Distinct a.ID, null �ϼ�ID, a.����, a.����, a.���, a.����, b.ҩ�ⵥλ, b.סԺ��λ, b.���ﵥλ " & vbLf & _
                     "From �շ���ĿĿ¼ A, ҩƷ��� B, �շ���Ŀ���� C " & vbLf & _
                     "Where a.Id = b.ҩƷid And a.id = c.�շ�ϸĿid And A.��� in ('5','6','7') " & vbLf & _
                     "  And (to_char(A.����ʱ��, 'yyyy-mm-dd') = '3000-01-01' or A.����ʱ�� is null) " & _
                     "  And C.���� = 1 "
            intSysParam = Val(zlDatabase.GetPara("���뷽ʽ"))
            strMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
            
            If IsNumeric(strkey) Then
                strSQL = strSQL & " And (a.���� Like [1] Or C.���� Like [2] And C.����=3) "
            ElseIf zlCommFun.IsCharAlpha(strkey) Then
                strSQL = strSQL & " And C.���� Like [2] and c.����=" & IIf(intSysParam = 0, 1, 2) & " "
            ElseIf zlCommFun.IsCharChinese(strkey) Then
                strSQL = strSQL & " And C.���� Like [2] "
            Else
                strSQL = strSQL & " And (a.���� = [1] And C.���� Like [2] Or C.���� LIKE [2]) and c.����=" & IIf(intSysParam = 0, 1, 2) & " "
            End If
            strSQL = strSQL & vbNewLine & "Order by a.���� "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-ҩƷ" _
                    , False, "", "ѡ��", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strkey & "%" _
                    , strMatch & strkey & "%")
        End If
        
    Case 1
        '����
        If strkey = "" Then
            strSQL = "Select ID, �ϼ�id, ����, ����, '' ���, '' ����, '' As ���㵥λ, 0 As ĩ�� " & _
                     "From ���Ʒ���Ŀ¼ " & vbLf & _
                     "Where ���� = '7' " & vbLf & _
                     "Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id " & vbLf & _
                     "Union all " & vbLf & _
                     "Select i.Id, b.����id As �ϼ�id, i.����, i.����, i.���, i.����, i.���㵥λ, 1 As ĩ�� " & vbLf & _
                     "From �շ���ĿĿ¼ I, �������� T, ������ĿĿ¼ B " & vbLf & _
                     "Where i.Id = t.����id And t.����id = b.Id And i.��� = '4' " & vbLf & _
                     "  And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-����" _
                    , False, "", "ѡ��", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select Distinct i.Id, i.����, i.����, i.���, i.����, i.���㵥λ, 1 As ĩ�� " & vbLf & _
                     "From �շ���ĿĿ¼ I, �������� T, �շ���Ŀ���� B " & vbLf & _
                     "Where i.Id = t.����id And i.Id = b.�շ�ϸĿid And i.��� = '4' " & vbLf & _
                     "  And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) "
            intSysParam = Val(zlDatabase.GetPara("���뷽ʽ"))
            strMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
            
            If IsNumeric(strkey) Then
                strSQL = strSQL & " And (i.���� Like [1] Or b.���� Like [2] And b.����=3) "
            ElseIf zlCommFun.IsCharAlpha(strkey) Then
                strSQL = strSQL & " And b.���� Like [2] And b.���� = [3] "
            ElseIf zlCommFun.IsCharChinese(strkey) Then
                strSQL = strSQL & " And b.���� Like [2] "
            Else
                strSQL = strSQL & " And (i.���� = [1] And b.���� Like [2] Or b.���� LIKE [2]) And b.���� = [3] "
            End If
            strSQL = strSQL & vbLf & "Order by i.���� "
        
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-����" _
                    , False, "", "ѡ��", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strkey & "%" _
                    , strMatch & strkey & "%" _
                    , IIf(intSysParam = 0, 1, 2))
        End If
    Case 2
        '����
        If strkey = "" Then
            strSQL = "Select ID, 0 ĩ��, �ϼ�id, ����, ����, '' ���, '' ����, '' ɢװ��λ, '' ��װ��λ " & _
                     "From ���ʷ��� " & _
                     "Where ������� in ('��ͨ����', 'ҽ������') " & _
                     "Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id " & _
                     "Union All " & _
                     "Select ID, 1 ĩ��, ����id �ϼ�id, ����, ����, ���, ����, ɢװ��λ, ��װ��λ " & _
                     "From ����Ŀ¼ " & _
                     "Where (to_char(����ʱ��,'yyyy-MM-DD') = '3000-01-01' or ����ʱ�� is null) And ������� in ('��ͨ����', 'ҽ������') "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-����" _
                    , False, "", "ѡ��", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select ID, ����, ����, ���, ����, ɢװ��λ, ��װ��λ " & _
                     "From ����Ŀ¼ " & _
                     "Where (to_char(����ʱ��,'yyyy-MM-DD') = '3000-01-01' or ����ʱ�� is null) And ������� in ('��ͨ����', 'ҽ������') "
            
            If IsNumeric(strkey) Then
                strSQL = strSQL & " And (���� Like [1] Or ���� Like [2]) "
            ElseIf zlCommFun.IsCharAlpha(strkey) Then
                strSQL = strSQL & " And ���� Like [2] "
            ElseIf zlCommFun.IsCharChinese(strkey) Then
                strSQL = strSQL & " And ���� Like [2] "
            Else
                strSQL = strSQL & " And (���� = [1] And ���� Like [2] Or ���� LIKE [2]) "
            End If
            strSQL = strSQL & vbLf & "Order by ���� "
        
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-����" _
                    , False, "", "ѡ��", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strkey & "%" _
                    , "%" & strkey & "%")
        End If
    Case 3
        '�豸
        If strkey = "" Then
            strSQL = "Select ID, 0 ĩ��, �ϼ�id, ����, ����, '' ���, '' ����, '' ��λ " & _
                     "From �豸���� " & _
                     "Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id " & _
                     "Union All " & _
                     "Select ID, 1 ĩ��, ����id �ϼ�id, ����, ����, ���, ����, ��λ " & _
                     "From �豸Ŀ¼ " & _
                     "Where (to_char(����ʱ��,'yyyy-MM-DD') = '3000-01-01' or ����ʱ�� is null) "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-�豸" _
                    , False, "", "ѡ��", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select ID, ����, ����, ���, ����, ��λ " & _
                     "From �豸Ŀ¼ " & _
                     "Where (to_char(����ʱ��,'yyyy-MM-DD') = '3000-01-01' or ����ʱ�� is null) "
            
            If IsNumeric(strkey) Then
                strSQL = strSQL & " And (���� Like [1] Or ���� Like [2]) "
            ElseIf zlCommFun.IsCharAlpha(strkey) Then
                strSQL = strSQL & " And ���� Like [2] "
            ElseIf zlCommFun.IsCharChinese(strkey) Then
                strSQL = strSQL & " And ���� Like [2] "
            Else
                strSQL = strSQL & " And (���� = [1] And ���� Like [2] Or ���� LIKE [2]) "
            End If
            strSQL = strSQL & vbLf & "Order by ���� "
        
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-�豸" _
                    , False, "", "ѡ��", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strkey & "%" _
                    , "%" & strkey & "%")
        End If
    End Select
    
    If blnCancel = False And Not rsTemp Is Nothing Then
        txt��Ŀ����.Text = Nvl(rsTemp!����)
        txt��Ŀ����.Tag = Nvl(rsTemp!ID)
    End If
    
    If Not rsTemp Is Nothing Then
        rsTemp.Close
    ElseIf rsTemp Is Nothing And Not blnCancel Then
        MsgBox "δ�ҵ�����Ŀ!", vbInformation + vbDefaultButton1, gstrSysName
        txt��Ŀ����.SetFocus
        txt��Ŀ����.SelStart = 0
        txt��Ŀ����.SelLength = Len(txt��Ŀ����.Text)
    End If
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub

Private Function GetClassValue() As Integer
    Dim i As Integer
    For i = 0 To chkDept.Count - 1
        If chkDept(i).Value And chkDept(i).Enabled Then
            GetClassValue = i
            Exit Function
        End If
    Next
    GetClassValue = -1
End Function


Private Sub initColumn()
    Dim i As Integer
    '��ʼ�����
    With vsfList
        .Rows = 1
        .Cols = mColumn.Count
    End With
    
    'vsf�����ã��������п��ж��뷽ʽ���̶��ж��뷽ʽ��Ĭ��Ϊ���ж���
    VsfGridColFormat vsfList, mColumn.ѡ��, "ѡ��", 500, flexAlignCenterCenter, "ѡ��"
    VsfGridColFormat vsfList, mColumn.No, "��ⵥ�ݺ�", 1500, flexAlignLeftCenter, "��ⵥ�ݺ�"
    VsfGridColFormat vsfList, mColumn.�ⷿ, "�ⷿ", 1000, flexAlignLeftCenter, "�ⷿ"
    VsfGridColFormat vsfList, mColumn.���, "���", 640, flexAlignCenterCenter, "���"
    VsfGridColFormat vsfList, mColumn.��¼״̬, "��¼״̬", 0, flexAlignLeftCenter, "��¼״̬"
    VsfGridColFormat vsfList, mColumn.��Ŀid, "��ĿID", 0, flexAlignLeftCenter, "��ĿID"
    VsfGridColFormat vsfList, mColumn.��Ŀ��Ϣ, "��Ŀ��Ϣ", 2500, flexAlignLeftCenter, "��Ŀ��Ϣ"
    VsfGridColFormat vsfList, mColumn.���, "���", 1500, flexAlignLeftCenter, "���"
    VsfGridColFormat vsfList, mColumn.����, "����", 600, flexAlignLeftCenter, "����"
    VsfGridColFormat vsfList, mColumn.��λ, "��λ", 600, flexAlignLeftCenter, "��λ"
    VsfGridColFormat vsfList, mColumn.����, "����", 1000, flexAlignRightCenter, "����"
    VsfGridColFormat vsfList, mColumn.�ɹ���, "�ɹ���", 1000, flexAlignRightCenter, "�ɹ���"
    VsfGridColFormat vsfList, mColumn.�ɹ����, "�ɹ����", 1000, flexAlignRightCenter, "�ɹ����"
    VsfGridColFormat vsfList, mColumn.��Ʊ��, "��Ʊ��", 1500, flexAlignLeftCenter, "��Ʊ��"
    VsfGridColFormat vsfList, mColumn.��Ʊ����, "��Ʊ����", 1500, flexAlignLeftCenter, "��Ʊ����"
    VsfGridColFormat vsfList, mColumn.��Ʊ���, "��Ʊ���", 1500, flexAlignRightCenter, "��Ʊ���"
    VsfGridColFormat vsfList, mColumn.��ʶ, "��ʶ", 0, flexAlignLeftCenter, "��ʶ"
    VsfGridColFormat vsfList, mColumn.�շ�ID, "�շ�ID", 0, flexAlignLeftCenter, "�շ�ID"
    VsfGridColFormat vsfList, mColumn.�������, "�������", 0, flexAlignLeftCenter, "�������"
    
End Sub

Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf�����ã��������п��ж��뷽ʽ���̶��ж��뷽ʽ��Ĭ��Ϊ���ж��룩
    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth: If lngColWidth = 0 Then .ColHidden(intCol) = True
        .ColAlignment(intCol) = intColAlignment
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub


Private Sub SetColumn(ByVal rsRecord As ADODB.Recordset)
    Dim lngLoop As Long
    With vsfList
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsRecord.RecordCount + 1
        For lngLoop = 1 To rsRecord.RecordCount
            .TextMatrix(lngLoop, mColumn.No) = rsRecord!��ⵥ�ݺ�
            .Cell(flexcpForeColor, lngLoop, mColumn.��Ŀ��Ϣ, lngLoop, mColumn.��Ŀ��Ϣ) = IIf(rsRecord!��ʶ = 1, glngColorҩƷ, IIf(rsRecord!��ʶ = 2, glngColor����, IIf(rsRecord!��ʶ = 3, glngColor�豸, glngColor����)))
            .TextMatrix(lngLoop, mColumn.�ⷿ) = rsRecord!�ⷿ
            .TextMatrix(lngLoop, mColumn.���) = rsRecord!���
            .TextMatrix(lngLoop, mColumn.��¼״̬) = rsRecord!��¼״̬
            .TextMatrix(lngLoop, mColumn.��Ŀid) = rsRecord!��Ŀid
            .TextMatrix(lngLoop, mColumn.��Ŀ��Ϣ) = rsRecord!��Ŀ��Ϣ
            .TextMatrix(lngLoop, mColumn.���) = "" & rsRecord!���
            .TextMatrix(lngLoop, mColumn.����) = "" & rsRecord!����
            .TextMatrix(lngLoop, mColumn.��λ) = rsRecord!��λ
            .TextMatrix(lngLoop, mColumn.����) = zlStr.FormatEx(IIf(IsNull(rsRecord!����), 0, rsRecord!����), mintShowPriceDigit, , True)
            .ColFormat(mColumn.����) = "#0.00000"
            .TextMatrix(lngLoop, mColumn.�ɹ���) = zlStr.FormatEx(rsRecord!�ɹ���, mintShowPriceDigit, , True)
            .ColFormat(mColumn.�ɹ���) = "#0.00000"
            .TextMatrix(lngLoop, mColumn.�ɹ����) = zlStr.FormatEx(rsRecord!�ɹ����, mintShowAmountDigit, , True)
            .ColFormat(mColumn.�ɹ����) = "#0.00000"
            .TextMatrix(lngLoop, mColumn.��Ʊ���) = zlStr.FormatEx(rsRecord!�ɹ����, mintShowAmountDigit, , True)
            .ColFormat(mColumn.��Ʊ���) = "#0.00000"
            .TextMatrix(lngLoop, mColumn.��ʶ) = rsRecord!��ʶ
            .TextMatrix(lngLoop, mColumn.�շ�ID) = rsRecord!�շ�ID
            .TextMatrix(lngLoop, mColumn.�������) = "" & rsRecord!�������
            
            rsRecord.MoveNext
        Next
        
        If .Rows > 1 Then
            .Cell(flexcpFontBold, 1, mColumn.��Ʊ���, .Rows - 1, mColumn.��Ʊ���) = True '��Ʊ���Ӵ�
            .Cell(flexcpFontBold, 1, mColumn.ѡ��, .Rows - 1, mColumn.ѡ��) = True 'ѡ��Ӵ�
        End If
        .Redraw = flexRDDirect
    End With

    If vsfList.Rows > 1 Then
        vsfList.Select 1, 1
    End If
End Sub

Private Sub txt��Ŀ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GetItem UCase(Trim(txt��Ŀ����))
End Sub

Private Sub vsfList_DblClick()
    With vsfList
        If .Rows < 2 Then Exit Sub
        If .MouseRow = 0 Or .MouseCol = mColumn.��Ʊ��� Then Exit Sub
            
        If .TextMatrix(.Row, mColumn.ѡ��) = "��" Then
            .TextMatrix(.Row, mColumn.ѡ��) = ""
            .TextMatrix(.Row, mColumn.��Ʊ��) = ""
            .TextMatrix(.Row, mColumn.��Ʊ����) = ""
        Else
            .TextMatrix(.Row, mColumn.ѡ��) = "��"
            .TextMatrix(.Row, mColumn.��Ʊ��) = txt��Ʊ��.Text
            .TextMatrix(.Row, mColumn.��Ʊ����) = txt��Ʊ����.Text
        End If
    End With
    
    AmountSum
End Sub

Private Sub vsfList_EnterCell()

    With vsfList
        .Editable = flexEDNone
        .FocusRect = flexFocusLight
        
        Select Case .Col
            Case mColumn.��Ʊ���
                If Val(.TextMatrix(.Row, mColumn.��ʶ)) <> 3 Then .Editable = flexEDKbdMouse
        End Select
        
    End With
    
End Sub

Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    With vsfList
        Select Case Col
            Case mColumn.��Ʊ���
                If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                ElseIf KeyAscii = Asc(".") Then
                    If InStr(.EditText, ".") <> 0 Then     'ֻ�ܴ���һ��С����
                        KeyAscii = 0
                    End If
                End If
                
        End Select
    End With
End Sub

Private Function SaveCard() As Boolean
    Dim lngLoop As Long
    Dim strNO As String
    Dim lng��� As Long
    Dim Str��Ʊ�� As String
    Dim str��Ʊ���� As String
    Dim dat��Ʊ���� As String
    Dim dbl��Ʊ��� As Double
    Dim int������־ As Integer '1��δ���������޸ķ�Ʊ��Ϣ; 2�����ֳ��������޸ķ�Ʊ��Ϣ
    Dim arrSql As Variant
    
    
    arrSql = Array()
    SaveCard = False
    If vsfList.Rows < 2 Then Exit Function
    '����Ƿ����빩ҩ��λ
    If Trim(txt��Ʊ��.Text) = "" Then
        MsgBox "��Ʊ�Ų���Ϊ�գ�", vbInformation, gstrSysName
        txt��Ʊ��.SetFocus
        Exit Function
    End If
    
    With vsfList
        '����Ƿ�ȫ��δ��ѡ
        For lngLoop = 1 To .Rows - 1
            If Trim(.TextMatrix(lngLoop, mColumn.ѡ��)) = "��" Then Exit For
        Next
        If lngLoop = .Rows Then
            MsgBox "δѡ�񵥾���Ϣ�����飡", vbInformation, gstrSysName
            vsfList.SetFocus
            Exit Function
        End If
        
        On Error GoTo errHandle
        Str��Ʊ�� = Trim(txt��Ʊ��.Text)
        str��Ʊ���� = Trim(txt��Ʊ����.Text)
        dat��Ʊ���� = dtp��Ʊ����.Value
        
        If MsgBox("��������ѡ��ķ�Ʊ��Ϣ���Ƿ������" & _
        vbCrLf & "��Ʊ�ţ�" & Str��Ʊ�� & "    ��Ʊ���룺" & IIf(str��Ʊ���� = "", "��", str��Ʊ����) _
        , vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Function
        
        For lngLoop = 1 To .Rows - 1
            If Trim(.TextMatrix(lngLoop, mColumn.ѡ��)) = "��" Then '��ѡ��
                strNO = Trim(.TextMatrix(lngLoop, mColumn.No))
                lng��� = Val(.TextMatrix(lngLoop, mColumn.���))
                dbl��Ʊ��� = IIf(Trim(.TextMatrix(lngLoop, mColumn.��Ʊ���)) = "", 0, .TextMatrix(lngLoop, mColumn.��Ʊ���))
                int������־ = IIf(Val(.TextMatrix(lngLoop, mColumn.��¼״̬)) = 1, 1, 2)
                
                If Val(.TextMatrix(lngLoop, mColumn.��ʶ)) = 1 Then 'ҩƷ
                    gstrSQL = "zl_ҩƷ�⹺��Ʊ��Ϣ_UPDATE("
                    'NO
                    gstrSQL = gstrSQL & "'" & strNO & "'"
                    '���
                    gstrSQL = gstrSQL & "," & lng���
                    '��Ʊ��
                    gstrSQL = gstrSQL & ",'" & Str��Ʊ�� & "'"
                    '��Ʊ����
                    gstrSQL = gstrSQL & "," & IIf(dat��Ʊ���� = "", "Null", "to_date('" & Format(dat��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                    '��Ʊ���
                    gstrSQL = gstrSQL & "," & dbl��Ʊ���
                    '��ҩ��λID
                    gstrSQL = gstrSQL & "," & Val(txt��Ӧ��.Tag)
                    '������־
                    gstrSQL = gstrSQL & "," & int������־
                    '��Ʊ����
                    gstrSQL = gstrSQL & ",'" & str��Ʊ���� & "'"
                    '�Զ�������
                    gstrSQL = gstrSQL & "," & IIf(mbln�����־, 1, 0) & ""
                    gstrSQL = gstrSQL & ")"
                ElseIf Val(.TextMatrix(lngLoop, mColumn.��ʶ)) = 5 Then '����
                    gstrSQL = "zl_�����⹺��Ʊ��Ϣ_UPDATE( "
                    gstrSQL = gstrSQL & "'" & strNO & "',"
                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(lngLoop, mColumn.��¼״̬)) & ","
                    gstrSQL = gstrSQL & "" & lng��� & ","
                    gstrSQL = gstrSQL & "'" & Str��Ʊ�� & "',"
                    gstrSQL = gstrSQL & "" & IIf(dat��Ʊ���� = "", "Null", "to_date('" & Format(dat��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                    gstrSQL = gstrSQL & "" & dbl��Ʊ��� & ","
                    gstrSQL = gstrSQL & "" & Val(txt��Ӧ��.Tag) & ","
                    gstrSQL = gstrSQL & IIf(str��Ʊ���� = "", "NULL", "'" & str��Ʊ���� & "'") & ")"
                ElseIf Val(.TextMatrix(lngLoop, mColumn.��ʶ)) = 2 Then '����
                    gstrSQL = "ZL_�����⹺���_Invoice( "
                    '�շ���¼ID���������ĵ���ȡԭʼ�շ�ID��
                    gstrSQL = gstrSQL & "'" & Val(.TextMatrix(lngLoop, mColumn.�շ�ID)) & "',"
                    '��Ʊ��
                    gstrSQL = gstrSQL & "'" & Str��Ʊ�� & "',"
                    '��Ʊ����
                    gstrSQL = gstrSQL & "" & IIf(dat��Ʊ���� = "", "Null", "to_date('" & Format(dat��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                    '��Ʊ���
                    gstrSQL = gstrSQL & "" & dbl��Ʊ��� & ","
                    '��Ʊ����
                    gstrSQL = gstrSQL & "" & IIf(str��Ʊ���� = "", "NULL", "'" & str��Ʊ���� & "'") & ","
                    '��Ӧ��
                    gstrSQL = gstrSQL & Val(txt��Ӧ��.Tag) & ")"
                ElseIf Val(.TextMatrix(lngLoop, mColumn.��ʶ)) = 3 Then '�豸
                    gstrSQL = "ZL_�豸�⹺���_ModifyFP("
                    '   no_in        IN  �豸�շ���¼.no%TYPE := NULL,
                    gstrSQL = gstrSQL & "'" & strNO & "',"
                    '   �豸id_IN        IN  �豸�շ���¼.�豸id%type:=null,
                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(lngLoop, mColumn.��Ŀid)) & ","
                    '   ���_IN      IN  �豸�շ���¼.���%type:=null,
                    gstrSQL = gstrSQL & "" & lng��� & ","
                    '   �������_In
                    gstrSQL = gstrSQL & "" & IIf(Trim(.TextMatrix(lngLoop, mColumn.�������)) = "", "NULL", "'" & Trim(.TextMatrix(lngLoop, mColumn.�������)) & "'") & ","
                    '   ��Ʊ��_IN        IN  �豸�շ���¼.��Ʊ����%TYPE := NULL,
                    gstrSQL = gstrSQL & "" & Str��Ʊ�� & ","
                    '   ��Ʊ����_IN      IN  �豸�շ���¼.��Ʊ����%TYPE := NULL
                    gstrSQL = gstrSQL & "" & IIf(dat��Ʊ���� = "", "Null", "to_date('" & Format(dat��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                    '   ��Ʊ����
                    gstrSQL = gstrSQL & IIf(str��Ʊ���� = "", "NULL", "'" & str��Ʊ���� & "'") & ")"

                End If
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        Next
        
        gcnOracle.BeginTrans
        For lngLoop = 0 To UBound(arrSql)
            Call zlDatabase.ExecuteProcedure(CStr(arrSql(lngLoop)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
    End With
    SaveCard = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AmountSum()
    Dim lngLoop As Long
    Dim dblAmountSum As Double
    
    With vsfList
        For lngLoop = 1 To .Rows - 1
            If Trim(.TextMatrix(lngLoop, mColumn.ѡ��)) = "��" Then '��ѡ��
                dblAmountSum = dblAmountSum + Val(.TextMatrix(lngLoop, mColumn.��Ʊ���))
            End If
        Next
    End With
    
    stbThis.Panels(2).Text = "��ѡ��Ʊ���ϼƣ�" & dblAmountSum
End Sub

Private Sub vsfList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strkey As String
    
    With vsfList
        If Trim(.TextMatrix(.Row, mColumn.ѡ��)) = "��" Then mbln������ = True
        
        If Col = mColumn.��Ʊ��� Then
            .EditText = Trim(.EditText)
            strkey = Trim(.EditText)
        
            If .TextMatrix(Row, Col) = "" Or strkey = "" Then
                MsgBox "�Բ��𣬽��������룡", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
            If Not IsNumeric(strkey) And strkey <> "" Then
                MsgBox "�Բ��𣬽�����Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                Cancel = True
                Exit Sub
            End If
            If Val(strkey) < 0 Then
                MsgBox "�Բ��𣬽���Ϊ����,�����䣡", vbInformation + vbOKOnly, gstrSysName
                Cancel = True
                Exit Sub
            End If
                
           strkey = zlStr.FormatEx(strkey, mintShowAmountDigit, , True)
            .EditText = strkey
        End If
    End With
End Sub
