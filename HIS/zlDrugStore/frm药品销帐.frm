VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmҩƷ���� 
   Caption         =   "ҩƷ��ҩ����"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15510
   Icon            =   "frmҩƷ����.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   15510
   StartUpPosition =   2  '��Ļ����
   Begin VB.OptionButton optListType 
      Caption         =   "��ҩƷ������ʾ"
      Height          =   180
      Index           =   0
      Left            =   7080
      TabIndex        =   36
      Top             =   1650
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton optListType 
      Caption         =   "�����˻�����ʾ"
      Height          =   180
      Index           =   1
      Left            =   8760
      TabIndex        =   37
      Top             =   1650
      Width           =   1815
   End
   Begin VB.Frame fraCondition 
      Height          =   1575
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11415
      Begin VB.ComboBox cboNode 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "����(&T)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   24
         Top             =   638
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "ҽ������(&W)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   23
         Top             =   638
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ��(&R)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   22
         ToolTipText     =   "�ȼ���F2"
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cbo���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7800
         TabIndex        =   21
         Text            =   "cbo����"
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox cbo������ 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   20
         Text            =   "����������"
         Top             =   1020
         Width           =   2775
      End
      Begin VB.TextBox txtPati 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5760
         TabIndex        =   19
         ToolTipText     =   "����סԺ�š�����ID������(ָ���˲���ʱ)�����￨��"
         Top             =   1020
         Width           =   2415
      End
      Begin VB.CommandButton cmdAllSelect 
         Caption         =   "ȫѡ(&S)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   18
         ToolTipText     =   "�ȼ���F2"
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdAllUnSelect 
         Caption         =   "ȫ��(&U)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10320
         TabIndex        =   17
         ToolTipText     =   "�ȼ���F2"
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox chkNoTime 
         Caption         =   "�����ڼ�"
         Height          =   180
         Left            =   1200
         TabIndex        =   16
         Tag             =   "1|0"
         Top             =   255
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dtp��ʼʱ�� 
         Height          =   315
         Left            =   2400
         TabIndex        =   25
         Top             =   188
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   104267779
         CurrentDate     =   36985
      End
      Begin MSComCtl2.DTPicker Dtp����ʱ�� 
         Height          =   315
         Left            =   5640
         TabIndex        =   26
         Top             =   188
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   104267779
         CurrentDate     =   36985
      End
      Begin VB.Label lblDept 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7320
         TabIndex        =   40
         Top             =   660
         Width           =   420
      End
      Begin VB.Label lblNode 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ֲ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4440
         TabIndex        =   39
         Top             =   660
         Width           =   420
      End
      Begin VB.Label lblʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "�����ڼ�"
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
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��"
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
         Left            =   5280
         TabIndex        =   31
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   30
         Top             =   660
         Width           =   840
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4440
         TabIndex        =   28
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label lblPatiInputType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ�š�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4920
         TabIndex        =   27
         Top             =   1080
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   11520
      TabIndex        =   11
      ToolTipText     =   "�ȼ���F2"
      Top             =   480
      Width           =   1335
   End
   Begin TabDlg.SSTab sstabList 
      Height          =   6780
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   11959
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "     δ���(&0)     "
      TabPicture(0)   =   "frmҩƷ����.frx":06EA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl��ʾ(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "picHsc(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "picBatHsc(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "vsfBatch(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "vsfDetail(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "vsfMain(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "vsfList(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "     �����(&1)     "
      TabPicture(1)   =   "frmҩƷ����.frx":0706
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl��ʾ(1)"
      Tab(1).Control(1)=   "vsfBatch(1)"
      Tab(1).Control(2)=   "vsfDetail(1)"
      Tab(1).Control(3)=   "vsfMain(1)"
      Tab(1).Control(4)=   "picHsc(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "picBatHsc(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "vsfList(1)"
      Tab(1).ControlCount=   7
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2295
         Index           =   1
         Left            =   -74520
         TabIndex        =   38
         Tag             =   "��ϸ"
         Top             =   1800
         Visible         =   0   'False
         Width           =   12135
         _cx             =   21405
         _cy             =   4048
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   19
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmҩƷ����.frx":0722
         ScrollTrack     =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2295
         Index           =   0
         Left            =   45
         TabIndex        =   35
         Tag             =   "��ϸ"
         Top             =   1440
         Visible         =   0   'False
         Width           =   15255
         _cx             =   26908
         _cy             =   4048
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   20
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmҩƷ����.frx":099A
         ScrollTrack     =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   2055
         Index           =   0
         Left            =   45
         TabIndex        =   5
         Tag             =   "������"
         Top             =   360
         Width           =   12735
         _cx             =   22463
         _cy             =   3625
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmҩƷ����.frx":0C38
         ScrollTrack     =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   2295
         Index           =   0
         Left            =   45
         TabIndex        =   6
         Tag             =   "��ϸ"
         Top             =   2400
         Width           =   12735
         _cx             =   22463
         _cy             =   4048
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmҩƷ����.frx":0D78
         ScrollTrack     =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid vsfBatch 
         Height          =   1815
         Index           =   0
         Left            =   45
         TabIndex        =   9
         Tag             =   "��ϸ"
         Top             =   4920
         Width           =   12735
         _cx             =   22463
         _cy             =   3201
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmҩƷ����.frx":0FBB
         ScrollTrack     =   -1  'True
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
      Begin VB.PictureBox picBatHsc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   1
         Left            =   -75000
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   12735
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4850
         Width           =   12735
      End
      Begin VB.PictureBox picBatHsc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   0
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   12735
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   4850
         Width           =   12735
      End
      Begin VB.PictureBox picHsc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   1
         Left            =   -75000
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   12855
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2450
         Width           =   12855
      End
      Begin VB.PictureBox picHsc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   0
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   12855
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2450
         Width           =   12855
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   2055
         Index           =   1
         Left            =   -74955
         TabIndex        =   12
         Tag             =   "������"
         Top             =   360
         Width           =   12735
         _cx             =   22463
         _cy             =   3625
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmҩƷ����.frx":118A
         ScrollTrack     =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   2295
         Index           =   1
         Left            =   -74955
         TabIndex        =   13
         Tag             =   "��ϸ"
         Top             =   2520
         Width           =   12735
         _cx             =   22463
         _cy             =   4048
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmҩƷ����.frx":12A5
         ScrollTrack     =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid vsfBatch 
         Height          =   1815
         Index           =   1
         Left            =   -74955
         TabIndex        =   14
         Tag             =   "��ϸ"
         Top             =   4920
         Width           =   12735
         _cx             =   22463
         _cy             =   3201
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   15592924
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmҩƷ����.frx":14E8
         ScrollTrack     =   -1  'True
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
      Begin VB.Label lbl��ʾ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������ʼ�¼�б�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   1
         Left            =   -70200
         TabIndex        =   34
         Top             =   60
         Width           =   2025
      End
      Begin VB.Label lbl��ʾ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "δ������ʼ�¼�б�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   0
         Left            =   4800
         TabIndex        =   33
         Top             =   60
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "��ҩ����(&V)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   11520
      TabIndex        =   2
      ToolTipText     =   "�ȼ���F2"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   11520
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&E)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   11520
      TabIndex        =   0
      ToolTipText     =   "�ȼ���F2"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Menu mnuPati 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuPatiItem 
         Caption         =   "סԺ��(&0)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "ID(&1)"
         Index           =   1
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "����(&2)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "FrmҩƷ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�ӿڲ���
Private mlng�ⷿid As Long              '��ǰ�ⷿID
Private mstrUnit As String              '��ǰ�ⷿ�����ڵİ�װ��λ
Private mintҩƷ���� As Integer         'ҩƷ���ư�������
Private mint����λ�� As Integer
Private mint��ӡ��ҩ�嵥 As Integer
Private mstrReceiveMsg As String        '�����洫�ݵ�����������Ϣ����ʽΪ������ʱ��,����id|����ʱ��,����id...

'��������
Private mrsDetail As ADODB.Recordset            'δ�����ϸ��¼���ݼ�
Private mrsVerifyDetail As ADODB.Recordset      '�������ϸ��¼���ݼ�
Private mrsBatch As ADODB.Recordset             'δ���������ϸ���ݼ�
Private mrsVerifyBatch As ADODB.Recordset       '�����������ϸ���ݼ�

Private mblnDrop As Boolean                     '��KeyDown���ж������б��Ƿ񵯳�

Private mbln��˳�Ժ�������� As Boolean
Private mint�Ƿ�������ʾܾ� As Integer

Private mlngMainRow As Long                  '��ǰѡ�е���Ŀ�б�����
Private mlngDetailRow As Long                '��ǰѡ�е���ϸ�б�����
Private mlngListRow As Long                  '��ǰѡ�е��б�����

Private mblnAllowChange As Boolean              '�Ƿ������޸���������
Private mdblSum As Double

Private mblnStart As Boolean

Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F

Private mstrPrivs As String

Private mstrReturnWriteOffInfo As String    '���ڼ�¼����������˵���Ϣ�������������棺����ʱ��,����id|����ʱ��,����id...

'���ѿ�
Private mstrCardType As String   '���ѿ�/���п���𣬸�ʽ������|ȫ��|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������);��
Private mintCardCount As Integer  '������
Private mobjSquareCard As Object    'һ��ͨ����

Private mobjPlugIn As Object             '��ҽӿڶ���

'ҽ���ӿ�
Private gclsInsure As New clsInsure

Private Type TYPE_MedicarePAR
    �������� As Boolean
    �����ϴ� As Boolean
    ������ɺ��ϴ� As Boolean
    ���������ϴ� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

Private Enum FindType
    סԺ�� = 0
    Id = 1
    ���� = 2
End Enum

Private Sub AutoExpendQuantity()
    '���ǵ�ͬһ����ID��Ӧ����շ�ID���������Ҫ�����������ֽ⵽����շ���¼��
    '�ֽ��ԭ���ǰ���Ŵ�����ȷ��䣨�Ѱ���Ž�������
    Dim n As Integer
    Dim dbl׼������ As Double
    Dim dblʣ������ As Double
    Dim int�շ���� As Integer
    Dim lng����id As Long
    Dim lngҩƷid As Long
    Dim str����ʱ�� As String

    With mrsBatch
        If .RecordCount > 0 Then .MoveFirst
        For n = 1 To .RecordCount
            dbl׼������ = !׼������
            
            If lng����id = !����ID And lngҩƷid = !ҩƷID And str����ʱ�� = !����ʱ�� Then

            Else
                dblʣ������ = !��������
            End If
            
            If dblʣ������ >= dbl׼������ Then
                dblʣ������ = dblʣ������ - dbl׼������
                !�������� = dbl׼������
            Else
                !�������� = dblʣ������
                dblʣ������ = 0
            End If
            
            lng����id = !����ID
            lngҩƷid = !ҩƷID
            str����ʱ�� = !����ʱ��
            
            .Update
            .MoveNext
        Next
    End With
    
    With mrsDetail
        .MoveFirst
        Do While Not .EOF
            mrsBatch.Filter = "����=" & !���� & _
                " And No='" & !NO & "' " & _
                " And ҩƷID=" & !ҩƷID & _
                " And ����ID=" & !����ID & _
                " And ����ʱ��='" & !����ʱ�� & "'"
            
            dbl׼������ = 0
            If mrsBatch.RecordCount > 0 Then
                Do While Not mrsBatch.EOF
                    dbl׼������ = dbl׼������ + mrsBatch!׼������
                    mrsBatch.MoveNext
                Loop
                
                If dbl׼������ < !�������� Then
                    mrsBatch.MoveFirst
                    Do While Not mrsBatch.EOF
                        mrsBatch!��˱�־ = 2
                        mrsBatch.Update
                        mrsBatch.MoveNext
                    Loop
                End If
            End If
            
            !׼������ = dbl׼������
            If dbl׼������ < !�������� Then
                !��˱�־ = 2
            End If
            .Update
            .MoveNext
        Loop
    End With
End Sub

Private Sub AutoExpendQuantityByVerify()
    '���ǵ�ͬһ����ID��Ӧ����շ�ID���������Ҫ�����������ֽ⵽����շ���¼��
    '�ֽ��ԭ���ǰ���Ŵ�����ȷ��䣨�Ѱ���Ž�������
    '�����������֮ǰ�Ѿܾ������˼�¼
    Dim n As Integer
    Dim dbl׼������ As Double
    Dim dblʣ������ As Double
    Dim int�շ���� As Integer
    Dim lng����id As Long
    Dim lngҩƷid As Long
    Dim str����ʱ�� As String
    Dim lng���� As Long


    With mrsVerifyBatch
        If .RecordCount > 0 Then .MoveFirst
        For n = 1 To .RecordCount
            dbl׼������ = !׼������
            
            If lng����id = !����ID And lngҩƷid = !ҩƷID And str����ʱ�� = !����ʱ�� And lng���� = !���� Then
               

            Else
                If (lng����id <> !����ID Or str����ʱ�� <> !����ʱ��) Then
                    dblʣ������ = !��������
                End If
            End If
            
            If dblʣ������ >= dbl׼������ Then
                dblʣ������ = dblʣ������ - dbl׼������
                !�������� = dbl׼������
            Else
                !�������� = dblʣ������
                dblʣ������ = 0
            End If
            
            lng����id = !����ID
            lngҩƷid = !ҩƷID
            str����ʱ�� = !����ʱ��
            lng���� = !����
            
            .Update
            .MoveNext
        Next
    End With
    
    With mrsVerifyDetail
        .Filter = "��˱�־=2"
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            mrsVerifyBatch.Filter = "����=" & !���� & _
                " And No='" & !NO & "' " & _
                " And ҩƷID=" & !ҩƷID & _
                " And ����ID=" & !����ID & _
                " And ����ʱ��='" & !����ʱ�� & "'"
            
            dbl׼������ = 0
            If mrsVerifyBatch.RecordCount > 0 Then
                Do While Not mrsVerifyBatch.EOF
                    dbl׼������ = dbl׼������ + mrsVerifyBatch!׼������
                    mrsVerifyBatch.MoveNext
                Loop
                
                If dbl׼������ < !�������� Then
                    mrsVerifyBatch.MoveFirst
                    Do While Not mrsVerifyBatch.EOF
                        mrsVerifyBatch!��˱�־ = 2
                        mrsVerifyBatch.Update
                        mrsVerifyBatch.MoveNext
                    Loop
                End If
            End If
            
'            !׼������ = dbl׼������
'            If dbl׼������ < !�������� Then
'                !��˱�־ = 2
'            End If
'            .Update
            .MoveNext
        Loop
    End With
End Sub
Private Sub GetPres(ByVal int�������� As Integer)
    'int�������ͣ�0-������1-ҽ������
    Dim rstemp As ADODB.Recordset
    Dim strSqlDept As String
    
    On Error GoTo errHandle
    If cbo����.ListIndex > 0 Then
        strSqlDept = " And B.����id = [1] "
    End If
        
    gstrSQL = "Select Distinct A.ID, A.����||'-'||A.���� As ���� " & _
        " From ��Ա�� A, ������Ա B " & _
        " Where A.ID = B.��Աid " & strSqlDept & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) " & _
        " Order By ����"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ա", Val(cbo����.ItemData(cbo����.ListIndex)))
        
    cbo������.Clear
    
    cbo������.AddItem "����������"
    cbo������.ItemData(cbo������.NewIndex) = 0

    Do While Not rstemp.EOF
        cbo������.AddItem rstemp!����
        cbo������.ItemData(cbo������.NewIndex) = rstemp!Id
        rstemp.MoveNext
    Loop
    
    cbo������.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetStockName()
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select ���� From ���ű� Where ID = [1] "
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ�ⷿ����", mlng�ⷿid)
    
    If Not rstemp.EOF Then
        Me.Caption = Me.Caption & "(" & rstemp!���� & ")"
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadDetailList(ByVal int��� As Integer, ByVal lngҩƷid As Long)
    If int��� = 0 Then
        With mrsDetail
            If mrsDetail Is Nothing Then Exit Sub
            If .RecordCount = 0 Then Exit Sub

            .Filter = "ҩƷID=" & lngҩƷid

            If .EOF Then Exit Sub

            Call IniGrid(int���, 2)
            Do While Not .EOF
                vsfDetail(0).rows = vsfDetail(0).rows + 1
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("��˱�־")) = IIf(!��˱�־ = 1, "��", IIf(!��˱�־ = 2, "��", ""))
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("�������")) = !�������
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("����")) = !����
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("NO")) = !NO
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("ҩƷid")) = !ҩƷID
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("����ʱ��")) = Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("��ʶ��")) = IIf(IsNull(!��ʶ��), "", !��ʶ��)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("����")) = IIf(IsNull(!����), "", !����)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("����")) = IIf(IsNull(!����), "", !����)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("׼������")) = FormatEx(!׼������ / !��װ, 5)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("��������")) = FormatEx(!�������� / !��װ, 5)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("��װ")) = IIf(IsNull(!��װ), "", !��װ)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("��λ")) = IIf(IsNull(!��λ), "", !��λ)
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("����id")) = !����ID
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("����id")) = !����ID
                vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("����ԭ��")) = IIf(IsNull(!����ԭ��), "", !����ԭ��)
                 
                '׼����С����������ʱ��׼�������Ϊ��ɫ
                If Val(vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("׼������"))) < Val(vsfDetail(0).TextMatrix(vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("��������"))) Then
                    vsfDetail(0).Cell(flexcpForeColor, vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("׼������"), vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("׼������")) = vbRed
                End If
                
               .MoveNext
            Loop
            
            '��˱�־�мӴ���ʾ
            vsfDetail(0).Cell(flexcpFontBold, 1, vsfDetail(0).ColIndex("��˱�־"), vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("��˱�־")) = True
            
            '��˱�־����ɫ��ʾ
            vsfDetail(0).Cell(flexcpForeColor, 1, vsfDetail(0).ColIndex("��˱�־"), vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("��˱�־")) = vbBlue
            
            '׼�������мӴ���ʾ
            vsfDetail(0).Cell(flexcpFontBold, 1, vsfDetail(0).ColIndex("׼������"), vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("׼������")) = True
            
            '���������мӴ���ʾ
            vsfDetail(0).Cell(flexcpFontBold, 1, vsfDetail(0).ColIndex("��������"), vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("��������")) = True
            
            '�����������Ϊ��ɫ
            vsfDetail(0).Cell(flexcpForeColor, 1, vsfDetail(0).ColIndex("��������"), vsfDetail(0).rows - 1, vsfDetail(0).ColIndex("��������")) = vbBlue
            
            vsfDetail(0).Row = 1
        End With
    Else
        With mrsVerifyDetail
            Call IniGrid(int���, 2)
            If mrsVerifyDetail Is Nothing Then Exit Sub
            
            .Filter = "ҩƷID=" & lngҩƷid
            If .EOF Then Exit Sub
            Do While Not .EOF
                vsfDetail(1).rows = vsfDetail(1).rows + 1
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("��˱�־")) = IIf(!��˱�־ = 1, "��", IIf(!��˱�־ = 2, "��", ""))
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("�������")) = !�������
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("����")) = !����
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("NO")) = !NO
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("ҩƷid")) = !ҩƷID
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("����ʱ��")) = Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("���ʱ��")) = Format(!���ʱ��, "yyyy-mm-dd hh:mm:ss")
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("�����")) = !�����
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("��ʶ��")) = IIf(IsNull(!��ʶ��), "", !��ʶ��)
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("����")) = IIf(IsNull(!����), "", !����)
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("����")) = IIf(IsNull(!����), "", !����)
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("��������")) = FormatEx(!�������� / !��װ, 5)
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("��װ")) = IIf(IsNull(!��װ), "", !��װ)
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("��λ")) = IIf(IsNull(!��λ), "", !��λ)
                vsfDetail(1).TextMatrix(vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("����id")) = !����ID
                
                If !��˱�־ = 2 Then
                    '��˾ܾ���־�мӴ���ʾ
                    vsfDetail(1).Cell(flexcpFontBold, vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("��˱�־"), vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("��˱�־")) = True
                    '��˾ܾ���־�к�ɫ��ʾ
                    vsfDetail(1).Cell(flexcpForeColor, vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("��˱�־"), vsfDetail(1).rows - 1, vsfDetail(1).ColIndex("��˱�־")) = vbRed
                End If
                
                .MoveNext
            Loop
            
            vsfDetail(1).Row = 1
        End With
    End If
End Sub

Private Sub LoadList(ByVal int��� As Integer)
    
    mdblSum = 0
    If int��� = 0 Then
        With mrsDetail
            If .RecordCount = 0 Then Exit Sub

            .Filter = ""

            If .EOF Then Exit Sub

            Call IniGrid(int���, 4)
            Do While Not .EOF
                vsfList(0).rows = vsfList(0).rows + 1
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("��˱�־")) = IIf(!��˱�־ = 1, "��", IIf(!��˱�־ = 2, "��", ""))
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("�������")) = !�������
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("����")) = !����
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("NO")) = !NO
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("ҩƷid")) = !ҩƷID
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("����ʱ��")) = Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("����")) = IIf(IsNull(!��ʶ��), "", !��ʶ�� & "-") & IIf(IsNull(!����), "", !����)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("����")) = IIf(IsNull(!��ǰ����), "", !��ǰ����)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("׼������")) = FormatEx(!׼������ / !��װ, 5)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("��������")) = FormatEx(!�������� / !��װ, 5)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("���˽��")) = FormatEx(!�������� * !���˽��, 2)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("��װ")) = IIf(IsNull(!��װ), "", !��װ)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("��λ")) = IIf(IsNull(!��λ), "", !��λ)
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("����id")) = !����ID
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("ҩƷ")) = !ҩƷ
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("��Ʒ��")) = !��Ʒ��
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("���")) = !���
                vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("����ԭ��")) = IIf(IsNull(!����ԭ��), "", !����ԭ��)
                
                mdblSum = mdblSum + FormatEx(!�������� * !���˽��, 2)
                
                '׼����С����������ʱ��׼�������Ϊ��ɫ
                If Val(vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("׼������"))) < Val(vsfList(0).TextMatrix(vsfList(0).rows - 1, vsfList(0).ColIndex("��������"))) Then
                    vsfList(0).Cell(flexcpForeColor, vsfList(0).rows - 1, vsfList(0).ColIndex("׼������"), vsfList(0).rows - 1, vsfList(0).ColIndex("׼������")) = vbRed
                End If
                
               .MoveNext
            Loop
            
            '��˱�־�мӴ���ʾ
            vsfList(0).Cell(flexcpFontBold, 1, vsfList(0).ColIndex("��˱�־"), vsfList(0).rows - 1, vsfList(0).ColIndex("��˱�־")) = True
            
            '��˱�־����ɫ��ʾ
            vsfList(0).Cell(flexcpForeColor, 1, vsfList(0).ColIndex("��˱�־"), vsfList(0).rows - 1, vsfList(0).ColIndex("��˱�־")) = vbBlue
            
            '׼�������мӴ���ʾ
            vsfList(0).Cell(flexcpFontBold, 1, vsfList(0).ColIndex("׼������"), vsfList(0).rows - 1, vsfList(0).ColIndex("׼������")) = True
            
            '���������мӴ���ʾ
            vsfList(0).Cell(flexcpFontBold, 1, vsfList(0).ColIndex("��������"), vsfList(0).rows - 1, vsfList(0).ColIndex("��������")) = True
            
            '�����������Ϊ��ɫ
            vsfList(0).Cell(flexcpForeColor, 1, vsfList(0).ColIndex("��������"), vsfList(0).rows - 1, vsfList(0).ColIndex("��������")) = vbBlue
            
            '��ʾ���ʽ��ϼ���Ϣ
            vsfList(0).rows = vsfList(0).rows + 1
            vsfList(0).Cell(flexcpText, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = "���˽��ϼƣ�" & FormatEx(mdblSum, 5)
            vsfList(0).Cell(flexcpFontBold, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = True
            vsfList(0).Cell(flexcpForeColor, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = vbRed
            vsfList(0).Cell(flexcpAlignment, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = flexAlignLeftCenter
            vsfList(0).MergeCells = flexMergeRestrictRows
            vsfList(0).MergeRow(vsfList(0).rows - 1) = True
            
            vsfList(0).Row = 1
        End With
    Else
        With mrsVerifyDetail
            If .RecordCount = 0 Then Exit Sub

            .Filter = ""

            If .EOF Then Exit Sub

            Call IniGrid(int���, 4)
            Do While Not .EOF
                vsfList(1).rows = vsfList(1).rows + 1
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("��˱�־")) = IIf(!��˱�־ = 1, "��", IIf(!��˱�־ = 2, "��", ""))
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("�������")) = !�������
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("����")) = !����
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("NO")) = !NO
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("ҩƷid")) = !ҩƷID
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("����ʱ��")) = Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("���ʱ��")) = Format(!���ʱ��, "yyyy-mm-dd hh:mm:ss")
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("�����")) = !�����
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("����")) = IIf(IsNull(!��ʶ��), "", !��ʶ�� & "-") & IIf(IsNull(!����), "", !����)
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("����")) = IIf(IsNull(!��ǰ����), "", !��ǰ����)
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("��������")) = FormatEx(!�������� / !��װ, 5)
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("��װ")) = IIf(IsNull(!��װ), "", !��װ)
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("��λ")) = IIf(IsNull(!��λ), "", !��λ)
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("����id")) = !����ID
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("ҩƷ")) = !ҩƷ
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("��Ʒ��")) = !��Ʒ��
                vsfList(1).TextMatrix(vsfList(1).rows - 1, vsfList(1).ColIndex("���")) = !���
                 
                If !��˱�־ = 2 Then
                    '��˾ܾ���־�мӴ���ʾ
                    vsfList(1).Cell(flexcpFontBold, vsfList(1).rows - 1, vsfList(1).ColIndex("��˱�־"), vsfList(1).rows - 1, vsfList(1).ColIndex("��˱�־")) = True
                    '��˾ܾ���־�к�ɫ��ʾ
                    vsfList(1).Cell(flexcpForeColor, vsfList(1).rows - 1, vsfList(1).ColIndex("��˱�־"), vsfList(1).rows - 1, vsfList(1).ColIndex("��˱�־")) = vbRed
                End If
                
                .MoveNext
            Loop
            
            vsfList(1).Row = 1
        End With
    End If
End Sub


Private Sub LoadBatchList(ByVal int��� As Integer, ByVal Int���� As Integer, _
                ByVal strNo As String, ByVal lngҩƷid As Long, _
                ByVal strʱ�� As String, ByVal lng����id As Long, _
                ByVal bln���±�־ As Boolean, ByVal int��˱�־ As Integer)
    If int��� = 0 Then
        With mrsBatch
            .Filter = "����=" & Int���� & _
                    " And No='" & strNo & "' " & _
                    " And ҩƷID=" & lngҩƷid & _
                    " And ����ID=" & lng����id & _
                    " And ����ʱ��='" & strʱ�� & "' "
            .Sort = "�շ���� Desc"
            
            If .EOF Then Exit Sub
            
            picBatHsc(0).Visible = True
            
            Call IniGrid(int���, 3)
            Do While Not .EOF
                vsfBatch(0).rows = vsfBatch(0).rows + 1
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("����")) = !����
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("NO")) = !NO
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("ҩƷid")) = !ҩƷID
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("����ʱ��")) = Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("����")) = IIf(IsNull(!����), "", !����)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("����")) = IIf(IsNull(!����), "", !����)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("Ч��")) = Format(!Ч��, "yyyy-mm-dd")
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("׼������")) = FormatEx(!׼������ / !��װ, 5)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("��������")) = FormatEx(!�������� / !��װ, 5)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("��װ")) = IIf(IsNull(!��װ), "", !��װ)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("��λ")) = IIf(IsNull(!��λ), "", !��λ)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("�շ����")) = IIf(IsNull(!�շ����), "", !�շ����)
                vsfBatch(0).TextMatrix(vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("����")) = FormatEx(!���� * !��װ, 5)
                
                If bln���±�־ Then
                    !��˱�־ = int��˱�־
                    .Update
                End If
                
               .MoveNext
            Loop
            vsfBatch(0).Cell(flexcpForeColor, 1, vsfBatch(0).ColIndex("��������"), vsfBatch(0).rows - 1, vsfBatch(0).ColIndex("��������")) = vbBlue
        End With
    Else
        With mrsVerifyBatch
            .Filter = "����=" & Int���� & _
                    " And No='" & strNo & "' " & _
                    " And ҩƷID=" & lngҩƷid & _
                    " And ����ID=" & lng����id & _
                    " And ���ʱ��='" & strʱ�� & "' "
            .Sort = "�շ���� Desc"

            If .EOF Then Exit Sub
        
            picBatHsc(1).Visible = True

            Call IniGrid(int���, 3)
            Do While Not .EOF
                vsfBatch(1).rows = vsfBatch(1).rows + 1
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("����")) = !����
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("NO")) = !NO
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("ҩƷid")) = !ҩƷID
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("���ʱ��")) = Format(!���ʱ��, "yyyy-mm-dd hh:mm:ss")
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("����")) = IIf(IsNull(!����), "", !����)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("����")) = IIf(IsNull(!����), "", !����)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("Ч��")) = Format(!Ч��, "yyyy-mm-dd")
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("׼������")) = FormatEx(!׼������ / !��װ, 5)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("��������")) = FormatEx(!�������� / !��װ, 5)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("��װ")) = IIf(IsNull(!��װ), "", !��װ)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("��λ")) = IIf(IsNull(!��λ), "", !��λ)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("�շ����")) = IIf(IsNull(!�շ����), "", !�շ����)
                vsfBatch(1).TextMatrix(vsfBatch(1).rows - 1, vsfBatch(1).ColIndex("����")) = FormatEx(!���� * !��װ, 5)
                .MoveNext
            Loop
        End With
    End If
End Sub
Private Sub GetRecord(ByVal int��� As Integer)
    Dim rstemp As ADODB.Recordset
    Dim strSubUnit As String
    Dim strSubName As String
    Dim intRow As Integer
    Dim str����ID As String
    Dim strSqlCondition As String
    Dim strNo As String
    Dim strҩ�� As String
    Dim lngSum As Long
    Dim arrExecute As Variant
    Dim i As Integer
    Dim strNos As String
    
'    On Error GoTo errHandle
    
    Call IniRecord(int���)
    vsfMain(int���).rows = 1
    vsfDetail(int���).rows = 1
    vsfBatch(int���).rows = 1

    
    ''''1����ȡ��������
    '�Ƿ����
    If int��� = 0 Then
        strSqlCondition = strSqlCondition & " And A.����� Is Null And A.״̬ = 0  "
        If chkNoTime.Value = 0 Then
            strSqlCondition = strSqlCondition & " And A.����ʱ�� Between [3] And [4] "
        End If
    Else
       strSqlCondition = strSqlCondition & " And A.����� Is Not Null And A.״̬ <> 0  "
        If chkNoTime.Value = 0 Then
            strSqlCondition = strSqlCondition & " And A.���ʱ�� Between [3] And [4] "
        End If
    End If
    
    '��λ����װ����
    Select Case mstrUnit
    Case "�ۼ۵�λ"
        strSubUnit = "X.���㵥λ ��λ,1 ��װ,A.���� As ��������,"
    Case "���ﵥλ"
        strSubUnit = "D.���ﵥλ ��λ,D.�����װ ��װ,A.���� As ��������,"
    Case "סԺ��λ"
        strSubUnit = "D.סԺ��λ ��λ,D.סԺ��װ ��װ,A.���� As ��������,"
    Case "ҩ�ⵥλ"
        strSubUnit = "D.ҩ�ⵥλ ��λ,D.ҩ���װ ��װ,A.���� As ��������,"
    End Select
    
    '����/ҽ������
    If cbo����.ListIndex > 0 Then
        strSqlCondition = strSqlCondition & " And A.���벿��id = [2] "
    End If
    
    '������
    If cbo������.ListIndex > 0 Then
        strSqlCondition = strSqlCondition & " And A.������=[5] "
    End If
    
    '��������
    If Val(txtPati.Tag) <> 0 Then
        strSqlCondition = strSqlCondition & " And B.����ID=[6] "
    End If
    
    If mstrReceiveMsg <> "" Then
        '����������������Ϣ����Ϊ��Ҫ��������ȥ��ԭ����ʱ������
        gstrSQL = "Select /*+ Rule*/ Distinct A.�շ�ϸĿid, X.���, " & strSubUnit & " '['||X.����||']' As ҩƷ����,X.���� As ͨ����,E.���� As ��Ʒ��,D.ҩƷ��Դ" & IIf(int��� = 0, ",A.���ۼ� ", "") & _
            " From (Select A.�շ�ϸĿid, Sum(A.����) As ����" & IIf(int��� = 0, ",B.��׼���� ���ۼ� ", "") & _
            " From סԺ���ü�¼ B, ������ҳ F, ���˷������� A , Table(f_Str2list2([7], '|', ',')) T " & _
            " Where A.�������=1 And A.����id = B.ID And A.��˲���id = [1] And B.����id = F.����id" & IIf(int��� = 1, "(+)", "") & " And B.��ҳid = F.��ҳid" & IIf(int��� = 1, "(+)", "") & _
            " And a.����ʱ�� = To_Date(t.C1, 'yyyy-mm-dd hh24:mi:ss') And b.����id = t.C2 "
            
        gstrSQL = gstrSQL & strSqlCondition
    Else
        gstrSQL = "Select Distinct A.�շ�ϸĿid, X.���, " & strSubUnit & " '['||X.����||']' As ҩƷ����,X.���� As ͨ����,E.���� As ��Ʒ��,D.ҩƷ��Դ" & IIf(int��� = 0, ",A.���ۼ� ", "") & _
            " From (Select A.�շ�ϸĿid, Sum(A.����) As ����" & IIf(int��� = 0, ",B.��׼���� ���ۼ� ", "") & _
            " From סԺ���ü�¼ B, ������ҳ F, ���˷������� A " & _
            " Where A.�������=1 And A.����id = B.ID And A.��˲���id = [1] And B.����id = F.����id" & IIf(int��� = 1, "(+)", "") & " And B.��ҳid = F.��ҳid" & IIf(int��� = 1, "(+)", "")
        gstrSQL = gstrSQL & strSqlCondition
    End If
        
    If mbln��˳�Ժ�������� = False Then
        gstrSQL = gstrSQL & " And F.��Ժ����" & IIf(int��� = 1, "(+)", "") & " Is Null "
    End If
    
'    gstrSQL = gstrSQL & strSqlCondition & _
'        " And Exists (Select 1 From ҩƷ�շ���¼ C " & _
'        " Where B.No = C.No And C.����id = A.����id And C.����� Is Not Null And (C.��¼״̬ = 1 Or Mod(C.��¼״̬, 3) = 0))"

    '�ų�������Һ�������Ĺ����в����ĵ���
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From ��Һ��ҩ���� Y,ҩƷ�շ���¼ S Where Y.�շ�id = S.ID And S.����id=B.ID) "
    
    gstrSQL = gstrSQL & " Group By A.�շ�ϸĿid" & IIf(int��� = 0, ",B.��׼����", "") & ") A,ҩƷ��� D, �շ���Ŀ���� E, �շ���ĿĿ¼ X " & _
        " Where A.�շ�ϸĿid = D.ҩƷid And A.�շ�ϸĿid = X.ID And X.ID = E.�շ�ϸĿid(+) And E.����(+) = 3 " & _
        " Order By ҩƷ����"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ����", mlng�ⷿid, Val(cbo����.ItemData(cbo����.ListIndex)), Dtp��ʼʱ��.Value, Dtp����ʱ��.Value, NeedName(cbo������.Text), Val(txtPati.Tag), mstrReceiveMsg)
    
    If rstemp.EOF Then
        Call IniGrid(int���, 0)
'        MsgBox "û���ҵ�������������ҩ�����¼��", vbInformation, gstrSysName
        cmdAllSelect.Enabled = False
        cmdAllUnSelect.Enabled = False
        Exit Sub
    End If
    
    If sstabList.Tab = 0 Then
        cmdAllSelect.Enabled = True
        cmdAllUnSelect.Enabled = True
    End If
    
    Call IniGrid(int���, 0)
    
    mdblSum = 0
    Do While Not rstemp.EOF
        With vsfMain(int���)
            .rows = .rows + 1
   
            .TextMatrix(.rows - 1, .ColIndex("�շ�ϸĿid")) = rstemp!�շ�ϸĿid
            
            If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                strҩ�� = rstemp!ͨ����
            Else
                strҩ�� = IIf(IsNull(rstemp!��Ʒ��), rstemp!ͨ����, rstemp!��Ʒ��)
            End If
            
            If mintҩƷ���� = 0 Then
                .TextMatrix(.rows - 1, .ColIndex("ҩƷ����")) = rstemp!ҩƷ���� & strҩ��
            ElseIf mintҩƷ���� = 1 Then
                .TextMatrix(.rows - 1, .ColIndex("ҩƷ����")) = rstemp!ҩƷ����
            ElseIf mintҩƷ���� = 2 Then
                .TextMatrix(.rows - 1, .ColIndex("ҩƷ����")) = strҩ��
            End If
            
            .TextMatrix(.rows - 1, .ColIndex("��Ʒ��")) = IIf(IsNull(rstemp!��Ʒ��), "", rstemp!��Ʒ��)
                
            .TextMatrix(.rows - 1, .ColIndex("���")) = IIf(IsNull(rstemp!���), "", rstemp!���)
            .TextMatrix(.rows - 1, .ColIndex("��������")) = FormatEx(rstemp!�������� / rstemp!��װ, 5)
            If int��� = 0 Then
                .TextMatrix(.rows - 1, .ColIndex("���˽��")) = FormatEx(rstemp!�������� * rstemp!���ۼ�, 2)
                mdblSum = mdblSum + FormatEx(rstemp!�������� * rstemp!���ۼ�, 2)
            End If
            .TextMatrix(.rows - 1, .ColIndex("��λ")) = rstemp!��λ
            .TextMatrix(.rows - 1, .ColIndex("��Դ")) = rstemp!ҩƷ��Դ
            rstemp.MoveNext
        End With
    Loop
    
    '���˽��ĺϼ���Ϣ
    If int��� = 0 Then
        vsfMain(int���).rows = vsfMain(int���).rows + 1
        vsfMain(int���).Cell(flexcpText, vsfMain(int���).rows - 1, 1, vsfMain(int���).rows - 1, vsfMain(int���).Cols - 1) = "���˽��ϼƣ�" & FormatEx(mdblSum, 5)
        vsfMain(int���).Cell(flexcpFontBold, vsfMain(int���).rows - 1, 1, vsfMain(int���).rows - 1, vsfMain(int���).Cols - 1) = True
        vsfMain(int���).Cell(flexcpForeColor, vsfMain(int���).rows - 1, 1, vsfMain(int���).rows - 1, vsfMain(int���).Cols - 1) = vbRed
        vsfMain(int���).MergeCells = flexMergeRestrictRows
        vsfMain(int���).MergeRow(vsfMain(int���).rows - 1) = True
    End If
    
    ''''2����ȡ��ϸ����
    '��λ�ִ�
    Select Case mstrUnit
    Case "�ۼ۵�λ"
        strSubUnit = "X.���㵥λ ��λ,1 ��װ, A.���� "
    Case "���ﵥλ"
        strSubUnit = "D.���ﵥλ ��λ,D.�����װ ��װ, A.���� "
    Case "סԺ��λ"
        strSubUnit = "D.סԺ��λ ��λ,D.סԺ��װ ��װ, A.���� "
    Case "ҩ�ⵥλ"
        strSubUnit = "D.ҩ�ⵥλ ��λ,D.ҩ���װ ��װ, A.���� "
    End Select
    
    If int��� = 0 Then
        If mstrReceiveMsg <> "" Then
            gstrSQL = "Select /*+ Rule*/ �������, ����, NO, ���, ҩƷ����, ͨ����, ��Ʒ��,���, ҩƷID, ����id, ����ʱ��, ��ʶ��, ����, ����id,����, ��λ, ��װ, Sum(����) As ��������,���ۼ�,��ǰ����,����ԭ�� " & _
                " From (Select Distinct E.���� As �������, C.����, C.NO,B.��׼���� ���ۼ�,b.���, '[' || x.���� || ']' As ҩƷ����, x.���� As ͨ����, w.���� As ��Ʒ��,X.���, C.ҩƷID, A.����id, A.����ʱ��,A.����ԭ��, B.��ʶ��, nvl(F.����,b.����) ����, B.����id,B.����,G.��ǰ����, " & strSubUnit & " " & _
                " From ���˷������� A, סԺ���ü�¼ B,ҩƷ�շ���¼ C, ҩƷ��� D, �շ���Ŀ���� W, �շ���ĿĿ¼ X, ���ű� P, ������ҳ F, ���ű� E,������Ϣ G  , Table(f_Str2list2([7], '|', ',')) T " & _
                " Where A.�������=1 And A.����id = B.ID And B.����id=G.����id(+) And B.No = C.No And B.ID = C.����id And B.��������id = P.ID And B.�շ�ϸĿid = D.ҩƷid And B.�շ�ϸĿid = X.ID " & _
                " And x.Id = w.�շ�ϸĿid(+) And w.����(+) = 3 And B.����id = F.����id And B.��ҳid = F.��ҳid " & _
                " And a.����ʱ�� = To_Date(t.C1, 'yyyy-mm-dd hh24:mi:ss') And b.����id = t.C2 "
        Else
            gstrSQL = "Select �������, ����, NO, ���, ҩƷ����, ͨ����, ��Ʒ��,���, ҩƷID, ����id, ����ʱ��, ��ʶ��, ����, ����id,����, ��λ, ��װ, Sum(����) As ��������,���ۼ�,��ǰ����,����ԭ�� " & _
                " From (Select Distinct E.���� As �������, C.����, C.NO,B.��׼���� ���ۼ�,b.���, '[' || x.���� || ']' As ҩƷ����, x.���� As ͨ����, w.���� As ��Ʒ��,X.���, C.ҩƷID, A.����id, A.����ʱ��,A.����ԭ��, B.��ʶ��, nvl(F.����,b.����) ����, B.����id,B.����,G.��ǰ����, " & strSubUnit & " " & _
                " From ���˷������� A, סԺ���ü�¼ B,ҩƷ�շ���¼ C, ҩƷ��� D, �շ���Ŀ���� W, �շ���ĿĿ¼ X, ���ű� P, ������ҳ F, ���ű� E,������Ϣ G  " & _
                " Where A.�������=1 And A.����id = B.ID And B.����id=G.����id(+) And B.No = C.No And B.ID = C.����id And B.��������id = P.ID And B.�շ�ϸĿid = D.ҩƷid And B.�շ�ϸĿid = X.ID " & _
                " And x.Id = w.�շ�ϸĿid(+) And w.����(+) = 3 And B.����id = F.����id And B.��ҳid = F.��ҳid "
        End If
        
        If mbln��˳�Ժ�������� = False Then
            gstrSQL = gstrSQL & " And F.��Ժ���� Is Null "
        End If
        
        '�ų�������Һ�������Ĺ����в������ĵ���
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where Y.�շ�id = C.ID) "
        
        gstrSQL = gstrSQL & " And A.���벿��id = E.ID And B.ִ�в���id = [1]  " & _
            " And C.����� Is Not Null And (C.��¼״̬ = 1 Or Mod(C.��¼״̬, 3) = 0) " & IIf(mstrReceiveMsg = "", strSqlCondition, "") & ")" & _
            " Group By �������, ����, NO, ���, ҩƷ����, ͨ����, ��Ʒ��,���, ҩƷID, ����id, ����ʱ��, ��ʶ��, ����, ����id, ����,��ǰ����, ��λ, ��װ,���ۼ�,����ԭ�� " & _
            " Order By �������, ��ʶ��, ����ʱ��, ����, NO, ��� "
    Else
        gstrSQL = "Select �������, ����, NO, ���, ҩƷ����, ͨ����, ��Ʒ��,���, ҩƷID, ����id, ����ʱ��, ���ʱ��, �����, ״̬, ��ʶ��, ����id, ����, ����, ��λ, ��װ, Sum(����) As ��������,��ǰ���� " & _
            " From (Select Distinct E.���� As �������, C.����, C.NO, b.���, '[' || x.���� || ']' As ҩƷ����, x.���� As ͨ����, w.���� As ��Ʒ��,X.���, C.ҩƷID, A.����id, A.����ʱ��, A.���ʱ��, A.�����, A.״̬, B.��ʶ��, nvl(F.����,b.����) ����, B.����id, B.����,G.��ǰ����, " & strSubUnit & " " & _
            " From ���˷������� A, סԺ���ü�¼ B,ҩƷ�շ���¼ C, ҩƷ��� D, �շ���Ŀ���� W, �շ���ĿĿ¼ X, ���ű� P, ������ҳ F, ���ű� E,������Ϣ G " & _
            " Where A.�������=1 And A.����id = B.ID And B.����id=G.����id(+) And B.No = C.No And B.ID = C.����id And B.��������id = P.ID And B.�շ�ϸĿid = D.ҩƷid And B.�շ�ϸĿid = X.ID " & _
            " And x.Id = w.�շ�ϸĿid(+) And w.����(+) = 3 And B.����id = F.����id(+) And B.��ҳid = F.��ҳid(+) " & _
            " And A.���벿��id = E.ID And B.ִ�в���id = [1]  "
            
        '�ų�������Һ�������Ĺ����в������ĵ���
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where Y.�շ�id = C.ID) "
        
        gstrSQL = gstrSQL & " And C.����� Is Not Null And (C.��¼״̬ = 1 Or Mod(C.��¼״̬, 3) = 0) " & strSqlCondition & ")" & _
            " Group By �������, ����, NO, ���, ҩƷ����, ͨ����, ��Ʒ��,���,ҩƷID, ����id, ����ʱ��, ���ʱ��, �����, ״̬, ��ʶ��, ����, ����id, ����,��ǰ����,��λ, ��װ " & _
            " Order By �������, ��ʶ��, ���ʱ��, �����, ����, NO, ��� "
    End If
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", mlng�ⷿid, Val(cbo����.ItemData(cbo����.ListIndex)), Dtp��ʼʱ��.Value, Dtp����ʱ��.Value, NeedName(cbo������.Text), Val(txtPati.Tag), mstrReceiveMsg)
    
    If rstemp.EOF Then
'        MsgBox "û���ҵ����������ĵ�����ϸ��¼��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If int��� = 0 Then
        Do While Not rstemp.EOF
            With mrsDetail
                .AddNew
                
                !��˱�־ = IIf(optListType(0).Value = True, 1, 0)
                !������� = rstemp!�������
                !���� = rstemp!����
                !NO = rstemp!NO
                !ҩƷID = rstemp!ҩƷID
                !����ID = rstemp!����ID
                !����ʱ�� = Format(rstemp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                !��ʶ�� = rstemp!��ʶ��
                !���� = rstemp!����
                !���� = rstemp!����
                !�������� = rstemp!��������
                !���˽�� = rstemp!���ۼ�
                !��װ = rstemp!��װ
                !��λ = rstemp!��λ
                !��� = rstemp!���
                !��ǰ���� = rstemp!��ǰ����
                !����ID = rstemp!����ID
                !����ԭ�� = rstemp!����ԭ��
                
                If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                    strҩ�� = rstemp!ͨ����
                Else
                    strҩ�� = IIf(IsNull(rstemp!��Ʒ��), rstemp!ͨ����, rstemp!��Ʒ��)
                End If
                
                If mintҩƷ���� = 0 Then
                    !ҩƷ = rstemp!ҩƷ���� & strҩ��
                ElseIf mintҩƷ���� = 1 Then
                    !ҩƷ = rstemp!ҩƷ����
                ElseIf mintҩƷ���� = 2 Then
                    !ҩƷ = strҩ��
                End If
                !��Ʒ�� = IIf(IsNull(rstemp!��Ʒ��), "", rstemp!��Ʒ��)
                
                .Update
                
                If InStr(1, strNo, rstemp!NO) = 0 Then
                    strNo = IIf(strNo = "", "", strNo & ",") & rstemp!NO
                End If
                rstemp.MoveNext
            End With
        Loop
    Else
        Do While Not rstemp.EOF
            With mrsVerifyDetail
                .AddNew
                
                !��˱�־ = rstemp!״̬
                !������� = rstemp!�������
                !���� = rstemp!����
                !NO = rstemp!NO
                !ҩƷID = rstemp!ҩƷID
                !����ID = rstemp!����ID
                !����ʱ�� = Format(rstemp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                !���ʱ�� = Format(rstemp!���ʱ��, "yyyy-mm-dd hh:mm:ss")
                !����� = rstemp!�����
                !��ʶ�� = rstemp!��ʶ��
                !���� = rstemp!����
                !���� = rstemp!����
                !�������� = rstemp!��������
                !��װ = rstemp!��װ
                !��λ = rstemp!��λ
                !��� = rstemp!���
                !��ǰ���� = rstemp!��ǰ����
                !����ID = rstemp!����ID
                
                If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                    strҩ�� = rstemp!ͨ����
                Else
                    strҩ�� = IIf(IsNull(rstemp!��Ʒ��), rstemp!ͨ����, rstemp!��Ʒ��)
                End If
                
                If mintҩƷ���� = 0 Then
                    !ҩƷ = rstemp!ҩƷ���� & strҩ��
                ElseIf mintҩƷ���� = 1 Then
                    !ҩƷ = rstemp!ҩƷ����
                ElseIf mintҩƷ���� = 2 Then
                    !ҩƷ = strҩ��
                End If
                !��Ʒ�� = IIf(IsNull(rstemp!��Ʒ��), "", rstemp!��Ʒ��)
                
                .Update
                
                If InStr(1, strNo, rstemp!NO) = 0 Then
                    strNo = IIf(strNo = "", "", strNo & ",") & rstemp!NO
                End If
                
                rstemp.MoveNext
            End With
        Loop
    End If
    
    ''''3����ȡ������ϸ����
    If int��� = 0 Then
        '��λ����װ����
        Select Case mstrUnit
        Case "�ۼ۵�λ"
            strSubUnit = "X.���㵥λ ��λ,1 ��װ,C.ʵ������ As ׼������,A.���� As ��������"
        Case "���ﵥλ"
            strSubUnit = "D.���ﵥλ ��λ,D.�����װ ��װ,C.ʵ������ As ׼������,A.���� As ��������"
        Case "סԺ��λ"
            strSubUnit = "D.סԺ��λ ��λ,D.סԺ��װ ��װ,C.ʵ������ As ׼������,A.���� As ��������"
        Case "ҩ�ⵥλ"
            strSubUnit = "D.ҩ�ⵥλ ��λ,D.ҩ���װ ��װ,C.ʵ������ As ׼������,A.���� As ��������"
        End Select
        
        ' 'Having Sum(ʵ������) > 0
        gstrSQL = "Select /*+ Rule*/ C.ID As �շ�ID, C.ҩƷID, C.����, C.NO, C.��� As �շ����, C.����, C.����, C.Ч��, F.����, P.���� As ��������, " & _
            " A.����id, B.��� As �������, B.��¼����, B.��ҳID, A.����ʱ��, C.���ۼ� As ����, " & strSubUnit & " " & _
            " From ���˷������� A, סԺ���ü�¼ B, " & _
            " (Select A.ID, A.����, A.NO, A.���, A.ҩƷid, A.����, A.����, A.Ч��, A.����id, B.ʵ������, A.���ۼ� " & _
            " From ҩƷ�շ���¼ A, " & _
            " (Select a.����, a.NO, a.���, a.ҩƷid, Sum(Nvl(a.����, 1) * a.ʵ������) As ʵ������ " & _
            " From ҩƷ�շ���¼ a ,Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b " & _
            " Where a.���� In (9, 10) And a.������� Is Not Null And a.No=b.Column_Value "
        
        '�ų�������Һ�������Ĺ����в������ĵ���
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where Y.�շ�id = A.ID) "
            
        gstrSQL = gstrSQL & " Group By ����, NO, ���, ҩƷid " & _
            " ) B" & _
            " Where A.NO = B.NO And A.���� = B.���� And A.ҩƷid + 0 = B.ҩƷid And A.��� = B.��� And A.����� Is Not Null " & _
            " And (A.��¼״̬ = 1 Or Mod(A.��¼״̬, 3) = 0))C, " & _
            " ҩƷ��� D, �շ���ĿĿ¼ X, ���ű� P, ������ҳ F, ���ű� E " & _
            " Where A.�������=1 And A.����id = B.ID And B.No = C.No And B.ID = C.����id And B.��������id = P.ID And B.�շ�ϸĿid = D.ҩƷid And B.�շ�ϸĿid = X.ID And B.����id = F.����id And B.��ҳid = F.��ҳid And A.���벿��id = E.ID " & _
            " And B.ִ�в���id = [1] " & strSqlCondition

        If mbln��˳�Ժ�������� = False Then
            gstrSQL = gstrSQL & " And F.��Ժ���� Is Null "
        End If
        
        gstrSQL = gstrSQL & " Order By A.����ʱ��, C.����, C.NO, C.��� Desc "
    Else
        '��λ����װ����
        '��λ����װ����
        Select Case mstrUnit
        Case "�ۼ۵�λ"
            strSubUnit = "X.���㵥λ ��λ,1 ��װ,C.ʵ������ As ׼������,A.���� As ��������"
        Case "���ﵥλ"
            strSubUnit = "D.���ﵥλ ��λ,D.�����װ ��װ,C.ʵ������ As ׼������, A.���� As ��������"
        Case "סԺ��λ"
            strSubUnit = "D.סԺ��λ ��λ,D.סԺ��װ ��װ,C.ʵ������ As ׼������, A.���� As ��������"
        Case "ҩ�ⵥλ"
            strSubUnit = "D.ҩ�ⵥλ ��λ,D.ҩ���װ ��װ,C.ʵ������ As ׼������, A.���� As ��������"
        End Select
        
        gstrSQL = "Select C.ID As �շ�ID, C.ҩƷID, C.����, C.NO, C.��� As �շ����, C.����, C.����, C.Ч��, F.����, P.���� As ��������,C.����, " & _
            " A.����id, B.��� As �������, B.��¼����, B.��ҳID, A.����ʱ��, A.���ʱ��, C.���ۼ� As ����, " & strSubUnit & " " & _
            " From ���˷������� A, סԺ���ü�¼ B,ҩƷ�շ���¼ C, ҩƷ��� D, �շ���ĿĿ¼ X, ���ű� P, ������ҳ F, ���ű� E " & _
            " Where A.�������=1 And A.����id = B.ID And B.No = C.No And B.ID = C.����id And B.��������id = P.ID And B.�շ�ϸĿid = D.ҩƷid And B.�շ�ϸĿid = X.ID And B.����id = F.����id(+) And B.��ҳid = F.��ҳid(+) And A.���벿��id = E.ID " & _
            " And B.ִ�в���id = [1]  " & strSqlCondition & _
            " And C.������� Is Not Null " & _
            " And ((A.״̬ = 1 And Mod(C.��¼״̬, 3) = 2 And A.���ʱ�� = C.�������)) "
            
        '�ų�������Һ�������Ĺ����в������ĵ���
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where Y.�շ�id = C.ID) "
        
        gstrSQL = gstrSQL & " Union All "
        
        ' 'Having Sum(ʵ������) > 0
        gstrSQL = gstrSQL & "Select /*+ Rule*/ C.ID As �շ�ID, C.ҩƷID, C.����, C.NO, C.��� As �շ����, C.����, C.����, C.Ч��, F.����, P.���� As ��������, C.����," & _
            " A.����id, B.��� As �������, B.��¼����, B.��ҳID, A.����ʱ��, A.���ʱ��,  C.���ۼ� As ����, " & strSubUnit & " " & _
            " From ���˷������� A, סԺ���ü�¼ B, " & _
            " (Select A.ID, A.����, A.NO, A.���, A.ҩƷid, A.����, A.����, A.Ч��, A.����id, B.ʵ������, A.���ۼ�,A.���� " & _
            " From ҩƷ�շ���¼ A, " & _
            " (Select a.����, a.NO, a.���, a.ҩƷid, Sum(Nvl(a.����, 1) * a.ʵ������) As ʵ������ " & _
            " From ҩƷ�շ���¼ a ,Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) b " & _
            " Where a.���� In (9, 10) And a.������� Is Not Null And a.No=b.Column_Value "
        
        '�ų�������Һ�������Ĺ����в������ĵ���
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From ��Һ��ҩ���� Y Where Y.�շ�id = A.ID) "
            
        gstrSQL = gstrSQL & " Group By ����, NO, ���, ҩƷid " & _
            " ) B" & _
            " Where A.NO = B.NO And A.���� = B.���� And A.ҩƷid + 0 = B.ҩƷid And A.��� = B.��� And A.����� Is Not Null " & _
            " And (A.��¼״̬ = 1 Or Mod(A.��¼״̬, 3) = 0))C, " & _
            " ҩƷ��� D, �շ���ĿĿ¼ X, ���ű� P, ������ҳ F, ���ű� E " & _
            " Where A.״̬=2 and A.�������=1 And A.����id = B.ID And B.No = C.No And B.ID = C.����id And B.��������id = P.ID And B.�շ�ϸĿid = D.ҩƷid And B.�շ�ϸĿid = X.ID And B.����id = F.����id And B.��ҳid = F.��ҳid And A.���벿��id = E.ID " & _
            " And B.ִ�в���id = [1] And A.����ʱ�� Between [3] And [4] "

        If mbln��˳�Ժ�������� = False Then
            gstrSQL = gstrSQL & " And F.��Ժ���� Is Null "
        End If
        
        gstrSQL = gstrSQL & " Order By ���ʱ��, ����, NO, �շ���� Desc"
        
'        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", mlng�ⷿid, Val(cbo����.ItemData(cbo����.ListIndex)), Dtp��ʼʱ��.Value, Dtp����ʱ��.Value, NeedName(cbo������.Text), Val(txtPati.Tag), strNo)
    End If
    
    If int��� = 0 Then
        'NO�����ܳ���4K���ֽ��ֱ�ִ��SQL�ٻ������ݼ�
        arrExecute = GetArrayByStr(strNo, 4000, ",")
        For i = 0 To UBound(arrExecute)
            strNos = arrExecute(i)
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", mlng�ⷿid, Val(cbo����.ItemData(cbo����.ListIndex)), Dtp��ʼʱ��.Value, Dtp����ʱ��.Value, NeedName(cbo������.Text), Val(txtPati.Tag), strNos)
                
            Do While Not rstemp.EOF
                With mrsBatch
                    .AddNew
                    !���� = rstemp!����
                    !NO = rstemp!NO
                    !ҩƷID = rstemp!ҩƷID
                    !����ʱ�� = Format(rstemp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                    !�շ���� = rstemp!�շ����
                    !���� = rstemp!����
                    !���� = rstemp!����
                    !Ч�� = rstemp!Ч��
                    
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And zlStr.Nvl(!Ч��) <> "" Then
                        '����Ϊ��Ч��
                        !Ч�� = Format(DateAdd("D", -1, !Ч��), "yyyy-mm-dd")
                    End If
                    
                    !׼������ = rstemp!׼������
                    !�������� = rstemp!��������
                    !��װ = rstemp!��װ
                    !��λ = rstemp!��λ
                    !���� = rstemp!����
                    !�շ�Id = rstemp!�շ�Id
                    !��ҳid = IIf(IsNull(rstemp!��ҳid), 0, rstemp!��ҳid)
                    !������� = rstemp!�������
                    !���� = rstemp!����
                    !����ID = rstemp!����ID
                    !��¼���� = rstemp!��¼����
                    !��˱�־ = IIf(optListType(0).Value = True, 1, 0)
                    .Update
                    
                    rstemp.MoveNext
                End With
            Loop
        Next
        
        Call AutoExpendQuantity
    Else
        arrExecute = GetArrayByStr(strNo, 4000, ",")
        For i = 0 To UBound(arrExecute)
            strNos = arrExecute(i)
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", mlng�ⷿid, Val(cbo����.ItemData(cbo����.ListIndex)), Dtp��ʼʱ��.Value, Dtp����ʱ��.Value, NeedName(cbo������.Text), Val(txtPati.Tag), strNos)
                
            Do While Not rstemp.EOF
                With mrsVerifyBatch
                    .AddNew
                    !���� = rstemp!����
                    !NO = rstemp!NO
                    !ҩƷID = rstemp!ҩƷID
                    !����ʱ�� = Format(rstemp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                    !���ʱ�� = Format(rstemp!���ʱ��, "yyyy-mm-dd hh:mm:ss")
                    !�շ���� = rstemp!�շ����
                    !���� = rstemp!����
                    !���� = rstemp!����
                    !Ч�� = rstemp!Ч��
                    !���� = rstemp!����
                    
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And zlStr.Nvl(!Ч��) <> "" Then
                        '����Ϊ��Ч��
                        !Ч�� = Format(DateAdd("D", -1, !Ч��), "yyyy-mm-dd")
                    End If
                    
                    !׼������ = Abs(rstemp!׼������)
                    !�������� = Abs(rstemp!��������)
                    !��װ = rstemp!��װ
                    !��λ = rstemp!��λ
                    !���� = rstemp!����
                    !�շ�Id = rstemp!�շ�Id
                    !��ҳid = IIf(IsNull(rstemp!��ҳid), 0, rstemp!��ҳid)
                    !������� = rstemp!�������
                    !���� = rstemp!����
                    !����ID = rstemp!����ID
                    !��¼���� = rstemp!��¼����
                    !��˱�־ = 1
                    .Update
                    
                    rstemp.MoveNext
                End With
            Loop
        Next
        Call AutoExpendQuantityByVerify
    End If
    
    ''''''4����λ�����ܵ�һ�У�����ȡ��һ����ϸ����
    If vsfMain(int���).rows > 1 Then
        mlngMainRow = 1
        mlngDetailRow = 1
        Call LoadDetailList(int���, Val(vsfMain(int���).TextMatrix(1, vsfMain(int���).ColIndex("�շ�ϸĿid"))))
        
        mlngListRow = 1
        Call LoadList(int���)
    End If

    cmdAllSelect.Enabled = False
    cmdAllUnSelect.Enabled = False
        
    If sstabList.Tab = 0 Then
        If mrsDetail.RecordCount > 0 Then
            cmdAllSelect.Enabled = True
            cmdAllUnSelect.Enabled = True
        End If
    End If
    
    '��ȡ��һ��������ϸ����
    If optListType(0).Value = True Then
        If vsfDetail(int���).rows > 1 Then
            If int��� = 0 Then
                Call LoadBatchList(int���, Val(vsfDetail(int���).TextMatrix(1, vsfDetail(int���).ColIndex("����"))), vsfDetail(int���).TextMatrix(1, vsfDetail(int���).ColIndex("NO")), Val(vsfDetail(int���).TextMatrix(1, vsfDetail(int���).ColIndex("ҩƷid"))), vsfDetail(int���).TextMatrix(1, vsfDetail(int���).ColIndex("����ʱ��")), Val(vsfDetail(int���).TextMatrix(1, vsfDetail(int���).ColIndex("����id"))), False, IIf(optListType(0).Value = True, 1, 0))
            Else
                Call LoadBatchList(int���, Val(vsfDetail(int���).TextMatrix(1, vsfDetail(int���).ColIndex("����"))), vsfDetail(int���).TextMatrix(1, vsfDetail(int���).ColIndex("NO")), Val(vsfDetail(int���).TextMatrix(1, vsfDetail(int���).ColIndex("ҩƷid"))), vsfDetail(int���).TextMatrix(1, vsfDetail(int���).ColIndex("���ʱ��")), Val(vsfDetail(int���).TextMatrix(1, vsfDetail(int���).ColIndex("����id"))), False, IIf(optListType(0).Value = True, 1, 0))
            End If
        End If
    Else
        If vsfList(int���).rows >= 2 Then
            If int��� = 0 Then
                Call LoadBatchList(int���, Val(vsfList(int���).TextMatrix(1, vsfList(int���).ColIndex("����"))), vsfList(int���).TextMatrix(1, vsfList(int���).ColIndex("NO")), Val(vsfList(int���).TextMatrix(1, vsfList(int���).ColIndex("ҩƷid"))), vsfList(int���).TextMatrix(1, vsfList(int���).ColIndex("����ʱ��")), Val(vsfList(int���).TextMatrix(1, vsfList(int���).ColIndex("����id"))), False, IIf(optListType(0).Value = True, 1, 0))
            Else
                Call LoadBatchList(int���, Val(vsfList(int���).TextMatrix(1, vsfList(int���).ColIndex("����"))), vsfList(int���).TextMatrix(1, vsfList(int���).ColIndex("NO")), Val(vsfList(int���).TextMatrix(1, vsfList(int���).ColIndex("ҩƷid"))), vsfList(int���).TextMatrix(1, vsfList(int���).ColIndex("���ʱ��")), Val(vsfList(int���).TextMatrix(1, vsfList(int���).ColIndex("����id"))), False, IIf(optListType(0).Value = True, 1, 0))
            End If
        End If
    End If
    
    cmdAllSelect.Enabled = False
    cmdAllUnSelect.Enabled = False
        
    If int��� = 0 Then
        If mrsBatch.RecordCount > 0 Then
            cmdAllSelect.Enabled = True
            cmdAllUnSelect.Enabled = True
        End If
        
        If mrsDetail.RecordCount > 0 Then
            If optListType(0).Value = True Then
                If vsfMain(int���).rows > 1 Then
                    vsfMain(int���).Row = 1
                    vsfMain(int���).SetFocus
                End If
            Else
                If vsfList(int���).rows > 1 Then
                    vsfMain(int���).Row = 1
                    vsfList(int���).SetFocus
                End If
            End If
        End If
    Else
        If mrsVerifyDetail.RecordCount > 0 Then
            If optListType(0).Value = True Then
                If vsfMain(int���).rows > 1 Then
                    vsfMain(int���).Row = 1
                    vsfMain(int���).SetFocus
                End If
            Else
                If vsfList(int���).rows > 1 Then
                    vsfMain(int���).Row = 1
                    vsfList(int���).SetFocus
                End If
            End If
        End If
    End If
    
    mstrReceiveMsg = ""
    
    Exit Sub
errHandle:
    mstrReceiveMsg = ""
    If ErrCenter() = 1 Then
        Resume
    End If

    Call SaveErrLog
End Sub


Private Sub IniRecord(ByVal int��� As Integer)
    'δ�����ϸ��¼��
    If int��� = 0 Then
        Set mrsDetail = New ADODB.Recordset
        With mrsDetail
            If .State = 1 Then .Close
            .Fields.Append "��˱�־", adDouble, 18, adFldIsNullable
            .Fields.Append "�������", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
            .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "��ʶ��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "׼������", adDouble, 18, adFldIsNullable
            .Fields.Append "��������", adDouble, 18, adFldIsNullable
            .Fields.Append "���˽��", adDouble, 18, adFldIsNullable
            .Fields.Append "��װ", adDouble, 18, adFldIsNullable
            .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "ҩƷ", adLongVarChar, 200, adFldIsNullable
            .Fields.Append "��Ʒ��", adLongVarChar, 200, adFldIsNullable
            .Fields.Append "���", adLongVarChar, 200, adFldIsNullable
            .Fields.Append "��ǰ����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "����ԭ��", adLongVarChar, 200, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
    Else
        '�������ϸ��¼��
        Set mrsVerifyDetail = New ADODB.Recordset
        With mrsVerifyDetail
            If .State = 1 Then .Close
            .Fields.Append "��˱�־", adDouble, 18, adFldIsNullable
            .Fields.Append "�������", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
            .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "���ʱ��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "�����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "��ʶ��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "��������", adDouble, 18, adFldIsNullable
            .Fields.Append "��װ", adDouble, 18, adFldIsNullable
            .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "ҩƷ", adLongVarChar, 200, adFldIsNullable
            .Fields.Append "��Ʒ��", adLongVarChar, 200, adFldIsNullable
            .Fields.Append "���", adLongVarChar, 200, adFldIsNullable
            .Fields.Append "��ǰ����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
    End If
    
    'δ���������ϸ��¼��
    If int��� = 0 Then
        Set mrsBatch = New ADODB.Recordset
        With mrsBatch
            If .State = 1 Then .Close
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
            .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "�շ����", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "Ч��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "׼������", adDouble, 18, adFldIsNullable
            .Fields.Append "��������", adDouble, 18, adFldIsNullable
            .Fields.Append "��װ", adDouble, 18, adFldIsNullable
            .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable
            .Fields.Append "��ҳID", adDouble, 18, adFldIsNullable
            .Fields.Append "�������", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
            .Fields.Append "��˱�־", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
    Else
        '�����������ϸ��¼��
        Set mrsVerifyBatch = New ADODB.Recordset
        With mrsVerifyBatch
            If .State = 1 Then .Close
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
            .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "���ʱ��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "�շ����", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "Ч��", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "׼������", adDouble, 18, adFldIsNullable
            .Fields.Append "��������", adDouble, 18, adFldIsNullable
            .Fields.Append "��װ", adDouble, 18, adFldIsNullable
            .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "��˱�־", adDouble, 18, adFldIsNullable
            .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable
            .Fields.Append "��ҳID", adDouble, 18, adFldIsNullable
            .Fields.Append "�������", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
    End If
End Sub
Private Sub IniGrid(ByVal int��� As Integer, ByVal intGrid As Integer)
    Dim i As Integer
    Dim strArr As Variant
    Dim strTemp As Variant
    
    'int��ˣ�0��δ��ˣ�1�������
    'intGrid��0����ʼ�����б�1����ʼ�����б�2����ʼ��ϸ�б�3��������ϸ�б�4-���˻����б�
    
    '��ʼ�����б�
    If intGrid = 0 Or intGrid = 1 Then
        With vsfMain(int���)
            .Redraw = flexRDNone
            .rows = 1
            
            If gintҩƷ������ʾ = 2 Then
                .ColWidth(.ColIndex("��Ʒ��")) = IIf(.ColWidth(.ColIndex("��Ʒ��")) = 0, 2000, .ColWidth(.ColIndex("��Ʒ��")))
            Else
                .ColWidth(.ColIndex("��Ʒ��")) = 0
            End If
                
            .Redraw = flexRDDirect
        End With
    End If
    
    '��ʼ��ϸ�б�
    If intGrid = 0 Or intGrid = 2 Then
        With vsfDetail(int���)
            .Redraw = flexRDNone
            .rows = 1
            .ColDataType(.ColIndex("��������")) = flexDTDouble

            .Redraw = flexRDDirect
        End With
    End If
    
    '��ʼ�����б�
    If intGrid = 0 Or intGrid = 3 Then
        With vsfBatch(int���)
            .Redraw = flexRDNone
            .rows = 1
            .ColDataType(.ColIndex("��������")) = flexDTDouble
            .TextMatrix(0, .ColIndex("Ч��")) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")

            .Redraw = flexRDDirect
        End With
    End If
    
    '��ʼ���˻����б�
    If intGrid = 0 Or intGrid = 4 Then
        With vsfList(int���)
            .Redraw = flexRDNone
            .rows = 1
            
            If gintҩƷ������ʾ = 2 Then
                .ColWidth(.ColIndex("��Ʒ��")) = IIf(.ColWidth(.ColIndex("��Ʒ��")) = 0, 2000, .ColWidth(.ColIndex("��Ʒ��")))
            Else
                .ColWidth(.ColIndex("��Ʒ��")) = 0
            End If
                
            .Redraw = flexRDDirect
        End With
    End If
End Sub

Private Sub Oper_ReVerify()
    '����֮ǰ�Ѿܾ������˼�¼
    Dim strCurrent As String
    Dim Int���� As Integer
    Dim strNo As String
    Dim lngҩƷid As Long
    Dim lng����id As Long
    Dim str����ʱ�� As String
    Dim i As Integer
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim bln�Ƿ�����ҩ As Boolean
    Dim str������� As String
    Dim strҩƷid As String
    Dim arrSql As Variant
    Dim blnBeginTrans As Boolean
    Dim Int��ҩ As Integer
    Dim lng����ID As Long
    Dim dbl�������� As Double
    Dim strReturnInfo As String
    Dim strReserve As String
    
    On Error GoTo errHandle
    
    If optListType(0).Value = True Then
        With vsfDetail(1)
            If .Row = 0 Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("����id"))) = 0 Then Exit Sub
            If Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) = "" Then Exit Sub
            
            Int���� = Val(.TextMatrix(.Row, .ColIndex("����")))
            strNo = .TextMatrix(.Row, .ColIndex("NO"))
            lngҩƷid = Val(.TextMatrix(.Row, .ColIndex("ҩƷID")))
            lng����id = Val(.TextMatrix(.Row, .ColIndex("����id")))
            str����ʱ�� = .TextMatrix(.Row, .ColIndex("����ʱ��"))
        End With
    Else
        With vsfList(1)
            If .Row = 0 Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("����id"))) = 0 Then Exit Sub
            If Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) = "" Then Exit Sub
            
            Int���� = Val(.TextMatrix(.Row, .ColIndex("����")))
            strNo = .TextMatrix(.Row, .ColIndex("NO"))
            lngҩƷid = Val(.TextMatrix(.Row, .ColIndex("ҩƷID")))
            lng����id = Val(.TextMatrix(.Row, .ColIndex("����id")))
            str����ʱ�� = .TextMatrix(.Row, .ColIndex("����ʱ��"))
        End With
    End If
    
    mrsVerifyDetail.Filter = "����id=" & lng����id & " And ����ʱ��='" & str����ʱ�� & "' "
    If mrsVerifyDetail.RecordCount = 0 Then Exit Sub
    lng����ID = mrsVerifyDetail!����ID
    dbl�������� = mrsVerifyDetail!��������
    
    '����Ƿ����
    mrsVerifyBatch.Filter = "����=" & Int���� & _
        " And No='" & strNo & "' " & _
        " And ҩƷID=" & lngҩƷid & _
        " And ����ID=" & lng����id & _
        " And ����ʱ��='" & str����ʱ�� & "' " & _
        " And ��˱�־=1 " & _
        " And ��������<>0 "
    If mrsVerifyBatch.RecordCount = 0 Then
        MsgBox "������ʣҩƷС�ڵ�ǰ��Ҫ���ʵ����������ܽ���������˲�����", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If IsOutPatient(mstrPrivs, mrsVerifyBatch!����, mrsVerifyBatch!NO, 2, 2) = False Then Exit Sub
    If IsReceiptBalance_Charge(1, mstrPrivs, mrsVerifyBatch!����, mrsVerifyBatch!NO, mrsVerifyBatch!�������, 2, 2) = False Then Exit Sub
    
    '��ʼ��ҽ������
    gclsInsure.InitOracle gcnOracle
    
    strCurrent = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    arrSql = Array()
    
    '����֮ǰ�Ѿܾ������˼�¼
    gstrSQL = "Zl_���˷�������_Cancel("
    '����ID
    gstrSQL = gstrSQL & lng����id
    '����ʱ��
    gstrSQL = gstrSQL & ",To_Date('" & str����ʱ�� & "','YYYY-MM-DD HH24:MI:SS')"
    '�����
    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
    '���ʱ��
    gstrSQL = gstrSQL & ",To_Date('" & strCurrent & "','yyyy-MM-dd hh24:mi:ss')"
    '��������
    gstrSQL = gstrSQL & ",0"
    gstrSQL = gstrSQL & ")"
    
    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = gstrSQL
     
    '��ҩ����
    Do While Not mrsVerifyBatch.EOF
        gstrSQL = "zl_ҩƷ�շ���¼_������ҩ("
        '�շ�ID
        gstrSQL = gstrSQL & mrsVerifyBatch!�շ�Id
        '�����
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '���ʱ��
        gstrSQL = gstrSQL & ",To_Date('" & strCurrent & "','yyyy-MM-dd hh24:mi:ss')"
        '����
        gstrSQL = gstrSQL & "," & IIf(IsNull(mrsVerifyBatch!����), "NULL", IIf(Mid(mrsVerifyBatch!����, 1, 1) = "(", "NULL", "'" & Mid(mrsVerifyBatch!����, 1, 8) & "'"))
        'Ч��
        gstrSQL = gstrSQL & "," & IIf(IsNull(mrsVerifyBatch!Ч��), "NULL", IIf(mrsVerifyBatch!Ч�� = "", "NULL", "To_Date('" & Format(mrsVerifyBatch!Ч��, "yyyy-MM-dd") & "','yyyy-MM-dd')"))
        '����
        gstrSQL = gstrSQL & "," & IIf(IsNull(mrsVerifyBatch!����), "NULL", "'" & mrsVerifyBatch!���� & "'")
        '��ҩ��
        gstrSQL = gstrSQL & "," & mrsVerifyBatch!��������
        '��ҩ�ⷿ
        gstrSQL = gstrSQL & ",NULL"
        '��ҩ��
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '����λ��
        gstrSQL = gstrSQL & "," & mint����λ��
        '����
        gstrSQL = gstrSQL & ",2"
        '���ܷ�ҩ��
        gstrSQL = gstrSQL & ",Null"
        gstrSQL = gstrSQL & ")"

        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
                
        bln�Ƿ�����ҩ = True
        
        If InStr("," & strҩƷid & ",", "," & mrsVerifyBatch!ҩƷID & ",") = 0 Then
            strҩƷid = IIf(strҩƷid = "", "", strҩƷid & ",") & mrsVerifyBatch!ҩƷID
        End If
        
        strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & Val(mrsVerifyBatch!�շ�Id) & "," & mrsVerifyBatch!��������
        
        '��¼��ǰ������˵ļ�¼������ʱ��Ͳ���ID�����ڷ��ظ�������
        If mstrReturnWriteOffInfo = "" Then
            mstrReturnWriteOffInfo = Format(mrsVerifyBatch!����ʱ��, "yyyy-mm-dd hh:mm:ss") & "," & lng����ID
        ElseIf InStr(mstrReturnWriteOffInfo & "|", Format(mrsVerifyBatch!����ʱ��, "yyyy-mm-dd hh:mm:ss") & "," & lng����ID & "|") = 0 Then
            mstrReturnWriteOffInfo = mstrReturnWriteOffInfo & "|" & Format(mrsVerifyBatch!����ʱ��, "yyyy-mm-dd hh:mm:ss") & "," & lng����ID
        End If
        
        mrsVerifyBatch.MoveNext
    Loop
    
    mrsVerifyBatch.MoveFirst
    str������� = mrsVerifyBatch!������� & ":" & dbl��������
    
    '������ü��˼�¼
    gstrSQL = "ZL_סԺ���ʼ�¼_Delete("
    'NO
    gstrSQL = gstrSQL & "'" & mrsVerifyBatch!NO & "'"
    '��ţ�������
    gstrSQL = gstrSQL & ",'" & str������� & "'"
    '����Ա���
    gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
    '����Ա����
    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
    '��¼����
    gstrSQL = gstrSQL & "," & mrsVerifyBatch!��¼����
    '����״̬
    gstrSQL = gstrSQL & ",1"
    gstrSQL = gstrSQL & ")"

    ReDim Preserve arrSql(UBound(arrSql) + 1)
    arrSql(UBound(arrSql)) = gstrSQL

    'ҽ������
    If Not IsNull(mrsVerifyBatch!����) And InStr(1, strMCNO, mrsVerifyBatch!NO) = 0 Then
        MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, , Val(mrsVerifyBatch!����))
        MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , Val(mrsVerifyBatch!����))
        strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & mrsVerifyBatch!NO & "," & mrsVerifyBatch!���� & _
                "," & IIf(MCPAR.���������ϴ�, "1", "0") & "," & IIf(MCPAR.������ɺ��ϴ�, "1", "0")
    End If
       
    '��ʾͣ��ҩƷ
    If strҩƷid <> "" Then
        Int��ҩ = 1
        Call CheckStopMedi(strҩƷid, Int��ҩ)
        If Int��ҩ = 2 Then Exit Sub
    End If

    '���д�����ҩ��������
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "Oper_ReVerify")
    Next
                
    'ҽ�������������ϴ�������ʱ�ϴ�
    If strMCNO <> "" Then
        arrMCRec = Split(strMCNO, "|")
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    gcnOracle.RollbackTrans:
                    Exit Sub
                End If
            End If
        Next
    End If
                            
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    'ҽ�������������ϴ�����ɺ��ϴ�
    If strMCNO <> "" Then
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    MsgBox "����""" & CStr(arrMCPar(0)) & """������������ҽ������ʧ�ܣ��õ��������ʡ�", vbInformation, gstrSysName
                End If
            End If
        Next
    End If
    
    If bln�Ƿ�����ҩ = True Then
        frm���ŷ�ҩ����New.BlnRefresh = True
        If mint��ӡ��ҩ�嵥 = 2 Then
            If MsgBox("����Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "��ҩʱ��=" & strCurrent, "��װϵ��=" & IIf(mstrUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), 2)
            End If
        ElseIf mint��ӡ��ҩ�嵥 = 1 Then
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "��ҩʱ��=" & strCurrent, "��װϵ��=" & IIf(mstrUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), 2)
        End If
    End If
    
    '������ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing And bln�Ƿ�����ҩ Then
        On Error Resume Next
        mobjPlugIn.DrugReturnByID mlng�ⷿid, strReturnInfo, CDate(strCurrent), strReserve
        err.Clear: On Error GoTo 0
    End If

    Call GetRecord(Val(sstabList.Tab))
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub

Private Sub SetAllSelect(ByVal intType As Integer)
    'intType:1-AllSelect;0-AllUnSelect
    Dim n As Integer
    Dim int��˱�־ As Integer
    
    If sstabList.Tab = 1 Then Exit Sub
    
    If optListType(0).Value = True Then
        With vsfDetail(0)
            If .rows <= 1 Then Exit Sub
            If .TextMatrix(1, .ColIndex("�������")) = "" Then Exit Sub
    
            For n = 1 To .rows - 1
                If .TextMatrix(n, .ColIndex("�������")) <> "" Then
                    If intType = 1 Then
                        If Val(.TextMatrix(n, .ColIndex("׼������"))) >= Val(.TextMatrix(n, .ColIndex("��������"))) Then
                            .TextMatrix(n, .ColIndex("��˱�־")) = "��"
                        Else
                            .TextMatrix(n, .ColIndex("��˱�־")) = "��"
                        End If
                    Else
                        .TextMatrix(n, .ColIndex("��˱�־")) = ""
                    End If
                End If
            Next
        End With
    Else
        With vsfList(0)
            If .rows <= 1 Then Exit Sub
            If .TextMatrix(1, .ColIndex("�������")) = "" Then Exit Sub
    
            For n = 1 To .rows - 1
                If .TextMatrix(n, .ColIndex("�������")) <> "" Then
                    If intType = 1 Then
                        If Val(.TextMatrix(n, .ColIndex("׼������"))) >= Val(.TextMatrix(n, .ColIndex("��������"))) Then
                            .TextMatrix(n, .ColIndex("��˱�־")) = "��"
                        Else
                            .TextMatrix(n, .ColIndex("��˱�־")) = "��"
                        End If
                    Else
                        .TextMatrix(n, .ColIndex("��˱�־")) = ""
                    End If
                End If
            Next
        End With
    End If
    
    With mrsDetail
        .Filter = ""
        .MoveFirst
        
        Do While Not .EOF
            If intType = 1 Then
                If !׼������ >= !�������� Then
                    !��˱�־ = 1
                Else
                    !��˱�־ = 2
                End If
            Else
                !��˱�־ = 0
            End If
            
            .Update
            
            'ͬ������������ϸ�б�
            mrsBatch.Filter = "����=" & !���� & _
                " And No='" & !NO & "' " & _
                " And ҩƷID=" & !ҩƷID & _
                " And ����ID=" & !����ID & _
                " And ����ʱ��='" & !����ʱ�� & "'"
            Do While Not mrsBatch.EOF
                mrsBatch!��˱�־ = !��˱�־
                mrsBatch.Update
                mrsBatch.MoveNext
            Loop
            
            .MoveNext
        Loop
    End With
End Sub

Public Function ShowForm(FrmMain As Form, ByVal lng�ⷿid As Long, ByVal strUnit As String, ByVal int����λ�� As Integer, _
    ByVal strCards As String, ByVal int��ӡ��ҩ�嵥 As Integer, ByVal strWriteOffMsg As String, _
    ByRef objSquareCard As Object, ByVal objPlugIn As Object) As String
    '���ر����ѽ��е����������Ϣ������ʱ��,����id|����ʱ��,����id...
    mlng�ⷿid = lng�ⷿid
    mstrUnit = strUnit
    mint����λ�� = int����λ��
    mstrCardType = strCards
    mint��ӡ��ҩ�嵥 = int��ӡ��ҩ�嵥
    mstrReceiveMsg = strWriteOffMsg
       
    Set mobjSquareCard = objSquareCard
    Set mobjPlugIn = objPlugIn
    
    If mstrCardType <> "" Then
        mintCardCount = UBound(Split(mstrCardType, ";")) + 1
    End If
    
    Me.Show vbModal, FrmMain
    
    ShowForm = mstrReturnWriteOffInfo
End Function
Private Sub Oper_Verify()
    '��ҩ����
    Dim i As Integer
    Dim strCurrent As String
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim bln�Ƿ�����ҩ As Boolean
    Dim str������� As String
    Dim strҩƷid As String
    Dim arrSql As Variant
    Dim blnBeginTrans As Boolean
    Dim Int��ҩ As Integer
    Dim strReturnInfo As String
    Dim strReserve As String
    
    arrSql = Array()
    
    On Error GoTo errHandle
    
    If optListType(0).Value = True Then
        If vsfMain(0).rows = 1 Then Exit Sub
        If vsfMain(0).TextMatrix(1, vsfMain(0).ColIndex("�շ�ϸĿid")) = "" Then Exit Sub
    Else
        If vsfList(0).rows = 1 Then Exit Sub
        If vsfList(0).TextMatrix(1, vsfList(0).ColIndex("ҩƷid")) = "" Then Exit Sub
    End If
    
    strCurrent = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    gclsInsure.InitOracle gcnOracle
    
    With mrsDetail
        If .State = 0 Then Exit Sub
        If .RecordCount = 0 Then Exit Sub
        
        '����Ƿ����
        .Filter = ""
        .Sort = "����, NO, ҩƷID, ����ʱ��"
        Do While Not .EOF
            mrsBatch.Filter = "����=" & !���� & _
                " And No='" & !NO & "' " & _
                " And ҩƷID=" & !ҩƷID & _
                " And ����ID=" & !����ID & _
                " And ����ʱ��='" & !����ʱ�� & "' " & _
                " And ��˱�־<>0 "
            If mrsBatch.RecordCount > 0 Then
                If mrsBatch!��˱�־ = 1 And !�������� <> 0 Then
                    If IsOutPatient(mstrPrivs, mrsBatch!����, mrsBatch!NO, 2, 2) = False Then Exit Sub
                    If IsReceiptBalance_Charge(1, mstrPrivs, mrsBatch!����, mrsBatch!NO, mrsBatch!�������, 2, 2) = False Then Exit Sub
                End If
            End If
            
            .MoveNext
        Loop
        
        '���ʴ�����ҩƷID����
        .Filter = ""
        .Sort = "ҩƷID,����,NO,����ʱ��"
        
        gclsInsure.InitOracle gcnOracle
        
        Screen.MousePointer = 11
                  
        Do While Not .EOF
            mrsBatch.Filter = "����=" & !���� & _
                " And No='" & !NO & "' " & _
                " And ҩƷID=" & !ҩƷID & _
                " And ����ID=" & !����ID & _
                " And ����ʱ��='" & !����ʱ�� & "' " & _
                " And ��˱�־<>0 "
            If mrsBatch.RecordCount > 0 Then
                '�������ʼ�¼����
                gstrSQL = "zl_���˷�������_Audit("
                '����ID
                gstrSQL = gstrSQL & mrsBatch!����ID
                '����ʱ��
                gstrSQL = gstrSQL & ",To_Date('" & mrsBatch!����ʱ�� & "','YYYY-MM-DD HH24:MI:SS')"
                '�����
                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                '���ʱ��
                gstrSQL = gstrSQL & ",To_Date('" & strCurrent & "','yyyy-MM-dd hh24:mi:ss')"
                '��˱�־
                gstrSQL = gstrSQL & "," & mrsBatch!��˱�־
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
                                
                                
                '��ҩ����
                Do While Not mrsBatch.EOF
                    If mrsBatch!��˱�־ = 1 And mrsBatch!�������� <> 0 Then
                        gstrSQL = "zl_ҩƷ�շ���¼_������ҩ("
                        '�շ�ID
                        gstrSQL = gstrSQL & mrsBatch!�շ�Id
                        '�����
                        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                        '���ʱ��
                        gstrSQL = gstrSQL & ",To_Date('" & strCurrent & "','yyyy-MM-dd hh24:mi:ss')"
                        '����
                        gstrSQL = gstrSQL & "," & IIf(IsNull(mrsBatch!����), "NULL", IIf(Mid(mrsBatch!����, 1, 1) = "(", "NULL", "'" & Mid(mrsBatch!����, 1, 8) & "'"))
                        'Ч��
                        gstrSQL = gstrSQL & "," & IIf(IsNull(mrsBatch!Ч��), "NULL", IIf(mrsBatch!Ч�� = "", "NULL", "To_Date('" & Format(mrsBatch!Ч��, "yyyy-MM-dd") & "','yyyy-MM-dd')"))
                        '����
                        gstrSQL = gstrSQL & "," & IIf(IsNull(mrsBatch!����), "NULL", "'" & mrsBatch!���� & "'")
                        '��ҩ��
                        gstrSQL = gstrSQL & "," & mrsBatch!��������
                        '��ҩ�ⷿ
                        gstrSQL = gstrSQL & ",NULL"
                        '��ҩ��
                        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                        '����λ��
                        gstrSQL = gstrSQL & "," & mint����λ��
                        '����
                        gstrSQL = gstrSQL & ",2"
                        '���ܷ�ҩ��
                        gstrSQL = gstrSQL & ",Null"
                        gstrSQL = gstrSQL & ")"
        
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = gstrSQL
                                
                        bln�Ƿ�����ҩ = True
                        
                        If InStr("," & strҩƷid & ",", "," & !ҩƷID & ",") = 0 Then
                            strҩƷid = IIf(strҩƷid = "", "", strҩƷid & ",") & !ҩƷID
                        End If
                        
                        strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & Val(mrsBatch!�շ�Id) & "," & mrsBatch!��������
                        
                        '��¼��ǰ������˵ļ�¼������ʱ��Ͳ���ID�����ڷ��ظ�������
                        If mstrReturnWriteOffInfo = "" Then
                            mstrReturnWriteOffInfo = Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss") & "," & !����ID
                        ElseIf InStr(mstrReturnWriteOffInfo & "|", Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss") & "," & !����ID & "|") = 0 Then
                            mstrReturnWriteOffInfo = mstrReturnWriteOffInfo & "|" & Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss") & "," & !����ID
                        End If
                    End If
                    
                    mrsBatch.MoveNext
                Loop
                
                mrsBatch.MoveFirst
                
                '���ʴ���
                If mrsBatch!��˱�־ = 1 And !�������� <> 0 Then
                    str������� = mrsBatch!������� & ":" & !��������
            
                    gstrSQL = "ZL_סԺ���ʼ�¼_Delete("
                    'NO
                    gstrSQL = gstrSQL & "'" & mrsBatch!NO & "'"
                    '��ţ�������
                    gstrSQL = gstrSQL & ",'" & str������� & "'"
                    '����Ա���
                    gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
                    '����Ա����
                    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                    '��¼����
                    gstrSQL = gstrSQL & "," & mrsBatch!��¼����
                    '����״̬
                    gstrSQL = gstrSQL & ",1"
                    gstrSQL = gstrSQL & ")"

                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
    
                    'ҽ������
                    If Not IsNull(mrsBatch!����) And InStr(1, strMCNO, mrsBatch!NO) = 0 Then
                        MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, , Val(mrsBatch!����))
                        MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , Val(mrsBatch!����))
                        strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & mrsBatch!NO & "," & mrsBatch!���� & _
                                "," & IIf(MCPAR.���������ϴ�, "1", "0") & "," & IIf(MCPAR.������ɺ��ϴ�, "1", "0")
                    End If
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    '��ʾͣ��ҩƷ
    If strҩƷid <> "" Then
        Int��ҩ = 1
        Call CheckStopMedi(strҩƷid, Int��ҩ)
        If Int��ҩ = 2 Then Exit Sub
    End If
 
     '���д�����ҩ��������
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "cmdVerify_Click")
    Next
                
    'ҽ�������������ϴ�������ʱ�ϴ�
    If strMCNO <> "" Then
        arrMCRec = Split(strMCNO, "|")
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    gcnOracle.RollbackTrans:
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        Next
    End If
                            
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    'ҽ�������������ϴ�����ɺ��ϴ�
    If strMCNO <> "" Then
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    MsgBox "����""" & CStr(arrMCPar(0)) & """������������ҽ������ʧ�ܣ��õ��������ʡ�", vbInformation, gstrSysName
                End If
            End If
        Next
    End If
    
    Screen.MousePointer = 0
    
    If bln�Ƿ�����ҩ = True Then
        frm���ŷ�ҩ����New.BlnRefresh = True
        If mint��ӡ��ҩ�嵥 = 2 Then
            If MsgBox("����Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "��ҩʱ��=" & strCurrent, "��װϵ��=" & IIf(mstrUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), 2)
            End If
        ElseIf mint��ӡ��ҩ�嵥 = 1 Then
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "��ҩʱ��=" & strCurrent, "��װϵ��=" & IIf(mstrUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), 2)
        End If
    End If
    
    '������ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing And bln�Ƿ�����ҩ Then
        On Error Resume Next
        mobjPlugIn.DrugReturnByID mlng�ⷿid, strReturnInfo, CDate(strCurrent), strReserve
        err.Clear: On Error GoTo 0
    End If

    Call GetRecord(Val(sstabList.Tab))
    
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub

Private Sub cboNode_Click()
    Call GetDept(IIf(opt����(0).Value = True, 0, 1))
End Sub


Private Sub cbo����_Click()
    If cbo����.ListIndex = -1 Then Exit Sub
    
    If Val(cbo����.Tag) <> cbo����.ItemData(cbo����.ListIndex) Then
        cbo����.Tag = cbo����.ItemData(cbo����.ListIndex)
        Call GetPres(IIf(opt����(0).Value, 0, 1))
    End If
End Sub

Private Sub Cbo����_KeyPress(KeyAscii As Integer)
    Dim sngX As Single
    Dim sngY As Single
    Dim sngH As Single
    Dim vRect As RECT
    Dim rstemp As ADODB.Recordset
    Dim StrNode As String
    Dim blnCancel As Boolean
    Dim i As Integer
    
    If KeyAscii = 13 Then
        If Trim(cbo����.Text) = "" Then Exit Sub
            
        If opt����(0).Value = True Then
            gstrSQL = " Select A.ID,b.���� As վ������, b.��� As վ��,A.����||'-'||A.���� ���� From ���ű� A, Zlnodelist B " & _
                " Where a.վ�� = b.���(+) And A.ID in (Select ����ID From ��������˵�� Where ��������='����' And ������� IN(2,3))" & _
                " And (A.����ʱ�� Is Null Or A.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                " And (A.���� Like [2] Or A.���� Like [2] Or A.���� Like [2])"
        Else
            gstrSQL = " Select A.ID,b.���� As վ������, b.��� As վ��,A.����||'-'||A.���� ���� From ���ű� A, Zlnodelist B " & _
                " Where a.վ�� = b.���(+) And A.ID in (Select ����ID From ��������˵�� Where �������� In ('���','����','����','����') And ������� IN(2,3))" & _
                " And (A.����ʱ�� Is Null Or A.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                " And (A.���� Like [2] Or A.���� Like [2] Or A.���� Like [2])"
        End If
        
        If cboNode.Visible Then
            If cboNode.ListIndex > 0 Then
                StrNode = cboNode.ItemData(cboNode.ListIndex)
            End If
        End If
        If StrNode <> "" Then
            gstrSQL = gstrSQL & " And A.վ�� = [1] "
        End If
        
        gstrSQL = gstrSQL & " Order By a.վ��, a.���� || '-' || a.���� "
        
        '�ж�����¼����ʾ����ѡ��
        vRect = zlControl.GetControlRect(cbo����.hWnd)
        sngX = vRect.Left
        sngY = vRect.Top
        sngH = cbo����.Height
        
        Set rstemp = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "ѡ���������", False, "", "ѡ���������", False, False, True, sngX, sngY, sngH, blnCancel, False, False, StrNode, UCase(cbo����.Text) & "%")
    
        If blnCancel = True Then Exit Sub
        
        If rstemp Is Nothing Then
            cbo����.Text = ""
            cbo����.Tag = ""
            cbo����.SetFocus
            Exit Sub
        Else
            For i = 1 To cbo����.ListCount - 1
                If cbo����.ItemData(i) = rstemp!Id Then
                    cbo����.ListIndex = i
                    Exit For
                End If
            Next
            
            If cbo����.ListIndex > 0 Then
                cbo����.Tag = cbo����.ItemData(cbo����.ListIndex)
                
                Call GetPres(IIf(opt����(0).Value, 0, 1))
            End If
        End If
    End If
End Sub

Private Sub cbo������_Click()
'    Exit Sub
End Sub

Private Sub cbo������_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo������.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cbo������_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim strText As String, strResult As String, strFilter As String

    If KeyAscii = 13 Then
        strText = UCase(cbo������.Text)
        If cbo������.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cbo������.List(cbo������.ListIndex) Then Call zlControl.CboSetIndex(cbo������.hWnd, -1)
        End If
        If strText = "" Then
            cbo������.ListIndex = -1
        ElseIf cbo������.ListIndex = -1 Then
            intIdx = -1

            For i = 1 To cbo������.ListCount - 1
                If Mid(cbo������.List(i), 1, InStr(1, cbo������.List(i), "-") - 1) = strText _
                    Or Mid(cbo������.List(i), InStr(1, cbo������.List(i), "-")) = strText Then
                    intIdx = i
                    Exit For
                End If
            Next

            If intIdx = -1 Then
                For i = 1 To cbo������.ListCount - 1
                    If UCase(cbo������.List(i)) Like strText & "*" Then
                        intIdx = i
                    End If
                Next
            End If

            cbo������.ListIndex = intIdx
            SendMessage cbo������.hWnd, CB_SHOWDROPDOWN, True, 0
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cbo������_Click
            Exit Sub
        End If
        If cbo������.ListIndex = -1 Then
            If cbo������.ListCount > 1 Then
                cbo������.ListIndex = 0
            Else
                cbo������.Text = "����������"
            End If
        Else
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
                Call cbo������_Click
            ElseIf intIdx <> cbo������.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cbo������.SetFocus
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
                Call cbo������_Click
            End If
        End If
    End If
End Sub

Private Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Private Sub chkNoTime_Click()
    If chkNoTime.Value = 0 Then
        Dtp��ʼʱ��.Enabled = True
        Dtp����ʱ��.Enabled = True
    Else
        Dtp��ʼʱ��.Enabled = False
        Dtp����ʱ��.Enabled = False
    End If
    
    If sstabList.Tab = 0 Then
        chkNoTime.Tag = chkNoTime.Value & Mid(chkNoTime.Tag, 2)
    Else
        chkNoTime.Tag = Mid(chkNoTime.Tag, 1, 2) & chkNoTime.Value
    End If
End Sub

Private Sub cmdAllSelect_Click()
    Call SetAllSelect(1)
End Sub

Private Sub cmdAllUnSelect_Click()
    Call SetAllSelect(0)
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub


Private Sub IniDate()
    Dim dateCurrent As Date
    
    dateCurrent = Sys.Currentdate
    
    Dtp��ʼʱ��.Value = CDate(Format(DateAdd("D", -1, dateCurrent), "yyyy-MM-dd 00:00:00"))
    Dtp����ʱ��.Value = CDate(Format(dateCurrent, "yyyy-MM-dd 23:59:59"))
End Sub

Private Sub GetDept(ByVal int�������� As Integer)
    'int�������ͣ�0-������1-ҽ������
    Dim rstemp As ADODB.Recordset
    Dim StrNode As String
    
    On Error GoTo errHandle
    Select Case int��������
        Case 0
            gstrSQL = " Select b.���� As վ������, b.��� As վ��,A.����||'-'||A.���� ����,A.ID From ���ű� A, Zlnodelist B " & _
                 " Where a.վ�� = b.���(+) And A.ID in (Select ����ID From ��������˵�� Where ��������='����' And ������� IN(2,3))" & _
                 " And (A.����ʱ�� Is Null Or A.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) "
        Case 1
            gstrSQL = " Select b.���� As վ������, b.��� As վ��,A.����||'-'||A.���� ����,A.ID From ���ű� A, Zlnodelist B " & _
             " Where a.վ�� = b.���(+) And A.ID in (Select ����ID From ��������˵�� Where �������� In ('���','����','����','����') And ������� IN(2,3))" & _
             " And (A.����ʱ�� Is Null Or A.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) "
    End Select
    
    If cboNode.Visible Then
        If cboNode.ListIndex > 0 Then
            StrNode = cboNode.ItemData(cboNode.ListIndex)
        End If
    End If
    If StrNode <> "" Then
        gstrSQL = gstrSQL & " And A.վ�� = [1] "
    End If
    
    gstrSQL = gstrSQL & " Order By a.վ��, a.���� || '-' || a.���� "
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ����", StrNode)
    
    cbo����.Clear
    cbo����.Text = ""
    cbo����.Tag = ""
    
    If int�������� = 0 Then
        cbo����.AddItem "���в���"
        cbo����.ItemData(cbo����.NewIndex) = 0
    Else
        cbo����.AddItem "���п���"
        cbo����.ItemData(cbo����.NewIndex) = 0
    End If
    
    Do While Not rstemp.EOF
        cbo����.AddItem rstemp!����
        cbo����.ItemData(cbo����.NewIndex) = rstemp!Id
        rstemp.MoveNext
    Loop
    
    cbo����.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetNode()
    Dim rstemp As ADODB.Recordset
    Dim strCurNode As String
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct b.���, b.���� " & _
        " From ���ű� A, Zlnodelist B " & _
        " Where a.վ�� = b.��� And a.Id In " & _
        " (Select ����id From ��������˵�� Where �������� In ('���', '����', '����', '����', '����') And ������� In (2, 3)) And " & _
        " (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
        " Order By b.��� "
    
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡվ����Ϣ")
    
    With cboNode
        .Visible = False
        .Clear
        
        If rstemp.RecordCount > 0 Then
            .Visible = True
            .AddItem "����վ��"
            
            Do While Not rstemp.EOF
                If strCurNode <> rstemp!��� Then
                    strCurNode = rstemp!���
                    .AddItem rstemp!����
                    .ItemData(.NewIndex) = rstemp!���
                End If
                rstemp.MoveNext
            Loop
            .ListIndex = 0
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub IniDept()
    If Lbl����.Tag = "" Then
        Lbl����.Tag = "-1"
        opt����_Click (0)
    End If
End Sub

Private Sub cmdPrint_Click()
'���ܣ���ӡ��ҩ֪ͨ��
    Dim StrDate As String
    
    If Trim(vsfDetail(1).TextMatrix(vsfDetail(1).Row, vsfDetail(1).ColIndex("����ʱ��"))) = "" Then Exit Sub
    StrDate = Format(vsfDetail(1).TextMatrix(vsfDetail(1).Row, vsfDetail(1).ColIndex("���ʱ��")), "yyyy-MM-dd HH:mm:ss")
    
    If Not IsDate(StrDate) Then
        MsgBox "�����м��б���ѡ����ϸ��¼��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "��ҩʱ��=" & StrDate, "��װϵ��=" & IIf(mstrUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), 2)
End Sub

Private Sub cmdRefresh_Click()
    If Trim(txtPati.Text) <> "" Then
        Call txtPati_KeyDown(vbKeyReturn, 0)
    Else
        Call GetRecord(Val(sstabList.Tab))
    End If
End Sub

Private Sub cmdVerify_Click()
    'ִ��Ԥ����
    Call setNOtExcetePrice
    
    If sstabList.Tab = 0 Then
        Call Oper_Verify
    Else
        Call Oper_ReVerify
    End If
End Sub

Private Sub Form_Activate()
    If mstrReceiveMsg <> "" Then
        cmdRefresh_Click
    End If
End Sub

Private Sub Form_Load()
    Dim objMenu As Menu
    Dim intCount As Integer
    Dim strCardName As String
    Dim int¼�뷽ʽ As Integer
    Dim int��ʾ��ʽ As Integer
    
    mblnStart = False
    
    mstrReturnWriteOffInfo = ""
    
    mstrPrivs = GetPrivFunc(glngSys, 1342)
    
    Call IniDate
    Call GetStockName
    Call GetNode
    Call IniDept
    
    mbln��˳�Ժ�������� = (Val(zldatabase.GetPara("��˳�Ժ���˵���������", glngSys, 1342, 0)) = 1)
    mint�Ƿ�������ʾܾ� = Val(zldatabase.GetPara("�Ƿ�������ʾܾ�", glngSys, 1342, 1))
    
    mintҩƷ���� = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ���ŷ�ҩ����", "ҩƷ������ʾ��ʽ", 0)))
    If mintҩƷ���� > 2 Or mintҩƷ���� < 0 Then mintҩƷ���� = 0
    
    cbo����.Tag = "-1"
    
    '���ѿ��˵�����
    If mintCardCount > 0 Then
        For intCount = 0 To UBound(Split(mstrCardType, ";"))
            'ȡ���п�����
            strCardName = Split(Split(mstrCardType, ";")(intCount), "|")(1)
            
            '��̬��Ӳ˵�
            Load Me.mnuPatiItem(Me.mnuPatiItem.UBound + 1)
            Set objMenu = Me.mnuPatiItem(Me.mnuPatiItem.UBound)
            objMenu.Caption = strCardName & "(" & 3 + intCount & ")"
            objMenu.Tag = Split(mstrCardType, ";")(intCount)
        Next
    End If
    
    If Val(zldatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        int¼�뷽ʽ = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\ҩƷ����", "¼�뷽ʽ", "0"))
        If int¼�뷽ʽ < 0 Or int¼�뷽ʽ > mnuPatiItem.count - 1 Then
            int¼�뷽ʽ = 0
        End If
        mnuPatiItem_Click int¼�뷽ʽ
        
        int��ʾ��ʽ = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\ҩƷ����", "��ʾ��ʽ", "0"))
        If int��ʾ��ʽ = 0 Then
            optListType(0).Value = True
        Else
            optListType(1).Value = True
        End If
    End If
    
    Call IniGrid(0, 0)
    Call IniGrid(1, 0)
    
    mblnStart = True
End Sub

Private Sub Form_Resize()
    Dim lngTmp As Long

    If WindowState = 1 Then Exit Sub
    On Error Resume Next
        
    If Me.Width < 13140 Then
        Me.Width = 13140
        Me.ScaleWidth = 12900
    End If
    If Me.Height < 8940 Then
        Me.Height = 8940
        Me.ScaleHeight = 8370
    End If
    
    cmdVerify.Left = Me.ScaleWidth - cmdVerify.Width - 50
    cmdExit.Left = cmdVerify.Left
    CmdHelp.Left = cmdVerify.Left
    cmdPrint.Left = cmdVerify.Left
    fraCondition.Left = Me.ScaleLeft + 20
    fraCondition.Width = Me.ScaleWidth - cmdVerify.Width - 100
    
    sstabList.Width = Me.ScaleWidth
    sstabList.Height = Me.ScaleHeight - fraCondition.Height
    
    'δ���
    vsfMain(0).Width = sstabList.Width - 90
    vsfDetail(0).Width = vsfMain(0).Width
    vsfBatch(0).Width = vsfMain(0).Width
    picHsc(0).Width = sstabList.Width
    picBatHsc(0).Width = sstabList.Width
    
    vsfBatch(0).Top = sstabList.Height - vsfBatch(0).Height - 50
    picBatHsc(0).Top = vsfBatch(0).Top - picBatHsc(0).Height - 50
    
    If picBatHsc(0).Top - vsfDetail(0).Top - 50 < 600 Then
        vsfDetail(0).Height = 2295
        vsfDetail(0).Top = picBatHsc(0).Top - vsfDetail(0).Height
        picHsc(0).Top = vsfDetail(0).Top - picHsc(0).Height - 50
    Else
        vsfDetail(0).Height = picBatHsc(0).Top - vsfDetail(0).Top - 50
    End If
    picHsc(0).Top = vsfDetail(0).Top - picHsc(0).Height - 50
    vsfMain(0).Height = picHsc(0).Top - vsfMain(0).Top - 50
    
    With vsfList(0)
        .Top = vsfMain(0).Top
        .Left = vsfMain(0).Left
        .Width = vsfMain(0).Width
        .Height = picBatHsc(0).Top - .Top - 50
    End With
    
    If optListType(0).Value = True Then
        vsfMain(0).Visible = True
        vsfDetail(0).Visible = True
        picHsc(0).Visible = True
        
        vsfList(0).Visible = False
    Else
        vsfMain(0).Visible = False
        vsfDetail(0).Visible = False
        picHsc(0).Visible = False
        
        vsfList(0).Visible = True
    End If
        
    '�����
    vsfMain(1).Width = sstabList.Width - 90
    vsfDetail(1).Width = vsfMain(1).Width
    vsfBatch(1).Width = vsfMain(1).Width
    picHsc(1).Width = sstabList.Width
    picBatHsc(1).Width = sstabList.Width
    
    vsfBatch(1).Top = sstabList.Height - vsfBatch(1).Height - 50
    picBatHsc(1).Top = vsfBatch(1).Top - picBatHsc(1).Height - 50
    
    If picBatHsc(1).Top - vsfDetail(1).Top - 50 < 600 Then
        vsfDetail(1).Height = 2295
        vsfDetail(1).Top = picBatHsc(1).Top - vsfDetail(1).Height
        picHsc(1).Top = vsfDetail(1).Top - picHsc(1).Height - 50
    Else
        vsfDetail(1).Height = picBatHsc(1).Top - vsfDetail(1).Top - 50
    End If
    picHsc(1).Top = vsfDetail(1).Top - picHsc(1).Height - 50
    vsfMain(1).Height = picHsc(1).Top - vsfMain(1).Top - 50
    
    With vsfList(1)
        .Top = vsfMain(1).Top
        .Left = vsfMain(1).Left
        .Width = vsfMain(1).Width
        .Height = picBatHsc(1).Top - .Top - 50
    End With
    
    If optListType(0).Value = True Then
        vsfMain(1).Visible = True
        vsfDetail(1).Visible = True
        picHsc(1).Visible = True
        
        vsfList(1).Visible = False
    Else
        vsfMain(1).Visible = False
        vsfDetail(1).Visible = False
        picHsc(1).Visible = False
        
        vsfList(1).Visible = True
    End If
    
    If cboNode.Visible Then
        lblNode.Visible = True
    Else
        lblNode.Visible = False
        lblDept.Left = lblNode.Left
        cbo����.Left = cboNode.Left
    End If
    
    Me.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mblnStart = False
    
    If Val(zldatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\ҩƷ����", "¼�뷽ʽ", Val(lblPatiInputType.Tag)
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\ҩƷ����", "��ʾ��ʽ", IIf(optListType(0).Value = True, 0, 1)
    End If
End Sub

Private Sub lblPatiInputType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        PopupMenu mnuPati, 2, lblPatiInputType.Left + lblPatiInputType.Width - 30, txtPati.Top
    End If
End Sub

Private Sub mnuPatiItem_Click(index As Integer)
    Dim i As Integer
    
    lblPatiInputType.Tag = index
    txtPati.Text = ""
    txtPati.PasswordChar = ""
    txtPati.MaxLength = 0
    
    Select Case index
        Case FindType.סԺ��
            lblPatiInputType.Caption = "סԺ�š�"
        Case FindType.Id
            lblPatiInputType.Caption = "ID��"
        Case FindType.����
            lblPatiInputType.Caption = "���š�"
        Case Else
            lblPatiInputType.Caption = Split(mnuPatiItem(index).Tag, "|")(gCardFormat.ȫ��) & "��"
            
            '�������ѿ�
            txtPati.MaxLength = Val(Split(mnuPatiItem(index).Tag, "|")(gCardFormat.���ų���))
            txtPati.PasswordChar = IIf(Trim(Split(mnuPatiItem(index).Tag, "|")(gCardFormat.��������)) <> "", "*", "")
    End Select
    
    For i = 0 To mnuPatiItem.count - 1
        mnuPatiItem(i).Checked = (i = index)
    Next
End Sub
    

Private Sub optListType_Click(index As Integer)
    Dim lngRow As Long
    
    If mblnStart = False Then Exit Sub
    
    If Not mrsDetail Is Nothing Then
        With mrsDetail
            .Filter = ""
            If .RecordCount > 0 Then
                Do While Not .EOF
                    !��˱�־ = IIf(index = 0, 1, 0)
                    .Update
                    
                    .MoveNext
                Loop
            End If
        End With
    End If
    
    If Not mrsBatch Is Nothing Then
        With mrsBatch
            .Filter = ""
            If .RecordCount > 0 Then
                Do While Not .EOF
                    !��˱�־ = IIf(index = 0, 1, 0)
                    .Update
                    
                    .MoveNext
                Loop
            End If
        End With
    End If
    
    DoEvents
    
    Call Form_Resize
    
    If index = 0 Then
        If vsfDetail(Val(sstabList.Tab)).rows > 1 Then
            mlngMainRow = 1
            mlngDetailRow = 1
            vsfMain(Val(sstabList.Tab)).Row = 1
            vsfMain(Val(sstabList.Tab)).SetFocus
        End If
    Else
        If vsfList(Val(sstabList.Tab)).rows > 1 Then
            If Val(sstabList.Tab) = 0 Then
                For lngRow = 1 To vsfList(Val(sstabList.Tab)).rows - 2
                    vsfList(Val(sstabList.Tab)).TextMatrix(lngRow, vsfList(Val(sstabList.Tab)).ColIndex("��˱�־")) = ""
                Next
            End If
            
            mlngListRow = 1
            vsfList(Val(sstabList.Tab)).Row = 1
            vsfList(Val(sstabList.Tab)).SetFocus
        End If
    End If
End Sub

Private Sub opt����_Click(index As Integer)
    If Val(Lbl����.Tag) <> index Then
        If index = 1 Then
            mnuPatiItem(2).Enabled = False
            If Val(lblPatiInputType.Tag) = FindType.���� Then
                Call mnuPatiItem_Click(0)
            End If
        Else
            mnuPatiItem(2).Enabled = True
        End If
        
        Call GetDept(index)
        Lbl����.Tag = index
    End If
End Sub






Private Sub picHsc_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsfMain(index).Height + y <= 500 Or vsfDetail(index).Height - y <= 500 Then Exit Sub
        
        picHsc(index).Top = picHsc(index).Top + y
        vsfMain(index).Height = vsfMain(index).Height + y
        vsfDetail(index).Height = vsfDetail(index).Height - y
        vsfDetail(index).Top = vsfDetail(index).Top + y
        
        Me.Refresh
    End If
End Sub


Private Sub sstabList_Click(PreviousTab As Integer)
    If sstabList.Tab = 0 Then
        lblʱ��.Caption = "�����ڼ�"
        If vsfMain(0).rows > 1 Then
            If vsfMain(0).TextMatrix(1, vsfMain(0).ColIndex("ҩƷ����")) <> "" Then
                cmdAllSelect.Enabled = True
                cmdAllUnSelect.Enabled = True
            End If
        End If
        cmdVerify.Caption = "��ҩ����(&V)"
        cmdVerify.Enabled = True
        chkNoTime.Value = Val(Mid(chkNoTime.Tag, 1, 1))
        cmdPrint.Enabled = False
    ElseIf sstabList.Tab = 1 Then
        lblʱ��.Caption = "����ڼ�"
        cmdAllSelect.Enabled = False
        cmdAllUnSelect.Enabled = False
        cmdVerify.Caption = "��������(&C)"
        cmdVerify.Enabled = False
        chkNoTime.Value = Val(Mid(chkNoTime.Tag, 3, 1))
        cmdPrint.Enabled = True
    End If
    
    If optListType(0).Value = True Then
        vsfMain(sstabList.Tab).Visible = True
        vsfDetail(sstabList.Tab).Visible = True
        vsfList(sstabList.Tab).Visible = False
    Else
        vsfMain(sstabList.Tab).Visible = False
        vsfDetail(sstabList.Tab).Visible = False
        vsfList(sstabList.Tab).Visible = True
    End If
End Sub

Private Sub txtPati_Change()
    If Val(lblPatiInputType.Tag) > 2 Then
        If Len(txtPati.Text) = txtPati.MaxLength Then
             Call txtPati_KeyDown(vbKeyReturn, 0)
        End If
    End If
End Sub
Private Sub txtPati_GotFocus()
    Call zlControl.TxtSelAll(txtPati)
End Sub
Private Sub txtPati_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim sngX As Single
    Dim sngY As Single
    Dim sngH As Single
    Dim vRect As RECT
    Dim strSqlCon As String
    Dim lng����ID As Long
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    txtPati.Text = Trim(txtPati.Text)
    
    If txtPati.Text = "" Then
        Call GetRecord(Val(sstabList.Tab))
        Exit Sub
    End If
    
    Select Case Val(lblPatiInputType.Tag)
        Case FindType.סԺ��
            If Not IsNumeric(txtPati.Text) Then Exit Sub
            strSqlCon = " And A.סԺ�� = [1] "
        Case FindType.Id
            strSqlCon = " And A.����ID = [2] "
        Case FindType.����
            If cbo����.ListIndex = 0 Then
                MsgBox "��ѡ������"
                Exit Sub
            End If
            strSqlCon = " And A.��ǰ����id = [3] And A.��ǰ���� = [1] "
        Case Else
            '�������ѿ���������ID����
            lng����ID = zlfuncCard_GetPatiID(mobjSquareCard, Val(Split(mnuPatiItem(Val(lblPatiInputType.Tag)).Tag, "|")(gCardFormat.�����ID)), txtPati.Text)
            strSqlCon = " And A.����ID = [4]"
    End Select
    
    gstrSQL = "Select A.����id As ID, A.����, A.סԺ��, B.����, A.��ǰ���� As ���� " & _
        " From ������Ϣ A, ���ű� B  Where A.��ǰ����id = B.Id" & IIf(mbln��˳�Ժ�������� = True, "(+)", "")
    
    gstrSQL = gstrSQL & strSqlCon & " Order By B.����, סԺ��"
    
    On Error GoTo errHandle
    
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", txtPati.Text, Val(txtPati.Text), Val(cbo����.ItemData(cbo����.ListIndex)), lng����ID)
    
    If rstemp.EOF Then
        txtPati.Text = ""
        txtPati.Tag = ""
        txtPati.SetFocus
        Exit Sub
    ElseIf rstemp.RecordCount = 1 Then
        'ֻ��һ����¼
        txtPati.Text = rstemp!����
        txtPati.Tag = rstemp!Id
    Else
        '�ж�����¼����ʾ����ѡ��
        vRect = zlControl.GetControlRect(txtPati.hWnd)
        sngX = vRect.Left
        sngY = vRect.Top
        sngH = txtPati.Height
        
        Set rstemp = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "ѡ����", False, "", "ѡ����", False, False, True, sngX, sngY, sngH, blnCancel, False, False, txtPati.Text, Val(txtPati.Text), Val(cbo����.ItemData(cbo����.ListIndex)))
    
        If blnCancel = True Then Exit Sub
        
        If rstemp Is Nothing Then
            txtPati.Text = ""
            txtPati.Tag = ""
            txtPati.SetFocus
            Exit Sub
        Else
            txtPati.Text = rstemp!����
            txtPati.Tag = rstemp!Id
        End If
    End If
    
    If Trim(txtPati.Text) <> "" Then Call GetRecord(Val(sstabList.Tab))
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

    Call SaveErrLog
End Sub


Private Sub txtPati_KeyPress(KeyAscii As Integer)
    If Val(lblPatiInputType.Tag) = FindType.סԺ�� Or Val(lblPatiInputType.Tag) = FindType.Id Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii <> vbKeyEscape Or KeyAscii = vbKeyBack Then Exit Sub
        KeyAscii = 0
    ElseIf Val(lblPatiInputType.Tag) > 2 Then
        '�����������ѿ�
        If InStr(":��;��?��''||" & Chr(22) & Chr(32), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
End Sub

Private Sub txtPati_Validate(Cancel As Boolean)
    If Trim(txtPati.Text) = "" Then txtPati.Tag = ""
End Sub

Private Sub vsfBatch_EnterCell(index As Integer)
    If index = 1 Then Exit Sub
    
    'ͨ��������ǲ��ܹ��޸����������ģ�����ͬһ�ŵ���ͬһҩƷ���ڶ�����ε�����£���Ȼ��Ҫ׼�������㹻
    With vsfBatch(0)
        .Editable = flexEDNone
        If .Row < 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("����")) = "" Then Exit Sub
        If .Col <> .ColIndex("��������") Then Exit Sub
        If mblnAllowChange = False Then Exit Sub
        If .rows = 2 Then Exit Sub
        
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Sub vsfBatch_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If index = 1 Then Exit Sub
    
    With vsfBatch(0)
        If .Row < 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("����")) = "" Then Exit Sub
        If .Col <> .ColIndex("��������") Then Exit Sub
        If KeyCode <> vbKeyReturn Then Exit Sub
        
        Call vsfBatch_ValidateEdit(0, .Row, .Col, True)
    End With
End Sub


Private Sub vsfBatch_KeyPressEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If index = 1 Then Exit Sub
    
    'ֻ����������
    If Col = vsfBatch(index).ColIndex("��������") Then
        If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub vsfBatch_RowColChange(index As Integer)
    '�ƶ���һ���ı�ǵ���ǰ�У�
    With vsfBatch(index)
        .Cell(flexcpText, 0, 0, .rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
        End If
    End With
End Sub

Private Sub vsfBatch_ValidateEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dblNewQuantity As Double
    Dim dblLeavingsQuantity As Double
    Dim dblQuantity As Double
    Dim i As Integer
    
    If index = 1 Then Exit Sub

    With vsfBatch(0)
        dblNewQuantity = Val(.EditText)
        
        If dblNewQuantity = Val(.TextMatrix(Row, .ColIndex("��������"))) Then Exit Sub
        
        If dblNewQuantity > Val(.TextMatrix(Row, .ColIndex("׼������"))) Or dblNewQuantity < 0 Then
            Cancel = True
            Exit Sub
        End If
        
        '����������
        dblLeavingsQuantity = Val(.TextMatrix(Row, .ColIndex("��������"))) - dblNewQuantity
        
        '�Ѳ���������䵽����������
        For i = 1 To .rows - 1
            If i <> Row Then
                dblQuantity = Val(.TextMatrix(i, .ColIndex("��������")))
                If dblQuantity + dblLeavingsQuantity <= Val(.TextMatrix(i, .ColIndex("׼������"))) And dblQuantity + dblLeavingsQuantity > 0 Then
                    .TextMatrix(i, .ColIndex("��������")) = dblQuantity + dblLeavingsQuantity
                    dblLeavingsQuantity = 0
                Else
                    .TextMatrix(i, .ColIndex("��������")) = Val(.TextMatrix(i, .ColIndex("׼������")))
                    dblLeavingsQuantity = dblLeavingsQuantity - (Val(.TextMatrix(i, .ColIndex("׼������"))) - dblQuantity)
                End If
                
                If dblLeavingsQuantity = 0 Then Exit For
            End If
        Next
        
        'ȷ�����ĵ�ǰ��������
        .EditText = FormatEx(dblNewQuantity + dblLeavingsQuantity, 5)
        .TextMatrix(Row, .ColIndex("��������")) = FormatEx(dblNewQuantity + dblLeavingsQuantity, 5)
        
        '���¼�¼���е���������
        For i = 1 To .rows - 1
            mrsBatch.Filter = "����=" & Val(.TextMatrix(i, .ColIndex("����"))) & _
                            " And No='" & .TextMatrix(i, .ColIndex("NO")) & "' " & _
                            " And ҩƷID=" & Val(.TextMatrix(i, .ColIndex("ҩƷid"))) & _
                            " And �շ����=" & Val(.TextMatrix(i, .ColIndex("�շ����"))) & _
                            " And ����ʱ��='" & .TextMatrix(i, .ColIndex("����ʱ��")) & "' "
            If mrsBatch.EOF Then Exit Sub
    
            mrsBatch!�������� = Val(.TextMatrix(i, .ColIndex("��������"))) * mrsBatch!��װ
            mrsBatch.Update
        Next
    End With
End Sub
Private Sub vsfDetail_Click(index As Integer)
    Dim bln���±�־ As Boolean
    Dim int��˱�־ As Integer
    
    With vsfDetail(index)
'        If index = 1 Then Exit Sub
        If .Row < 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("�������")) = "" Then Exit Sub

        If index = 0 And .MouseCol = .ColIndex("��˱�־") Then
            If .TextMatrix(.Row, .ColIndex("��˱�־")) = "��" Then
                .TextMatrix(.Row, .ColIndex("��˱�־")) = IIf(mint�Ƿ�������ʾܾ� = 1, "��", "")
                int��˱�־ = IIf(mint�Ƿ�������ʾܾ� = 1, 2, 0)
            ElseIf .TextMatrix(.Row, .ColIndex("��˱�־")) = "��" Then
                .TextMatrix(.Row, .ColIndex("��˱�־")) = ""
                int��˱�־ = 0
            Else
                If Val(.TextMatrix(.Row, .ColIndex("׼������"))) < Val(.TextMatrix(.Row, .ColIndex("��������"))) Then
                    If mint�Ƿ�������ʾܾ� = 1 Then
                        .TextMatrix(.Row, .ColIndex("��˱�־")) = "��"
                        int��˱�־ = 2
                    End If
                Else
                    .TextMatrix(.Row, .ColIndex("��˱�־")) = "��"
                    int��˱�־ = 1
                End If
            End If
            bln���±�־ = True
            
            '���¼�¼����˱��
            mrsDetail.Filter = "����=" & Val(.TextMatrix(.Row, .ColIndex("����"))) & _
                " And No='" & .TextMatrix(.Row, .ColIndex("NO")) & "' " & _
                " And ҩƷID=" & Val(.TextMatrix(.Row, .ColIndex("ҩƷid"))) & _
                " And ����ID=" & Val(.TextMatrix(.Row, .ColIndex("����id"))) & _
                " And ����ʱ��='" & .TextMatrix(.Row, .ColIndex("����ʱ��")) & "' "
            If mrsDetail.RecordCount > 0 Then
                mrsDetail!��˱�־ = int��˱�־
                mrsDetail.Update
            End If
        End If
        
        mblnAllowChange = False
        If Val(.TextMatrix(.Row, .ColIndex("׼������"))) > Val(.TextMatrix(.Row, .ColIndex("��������"))) Then
            mblnAllowChange = True
        End If
        
        '��ȡ������ϸ����
        If mlngDetailRow <> .Row Or bln���±�־ = True Then
            mlngDetailRow = .Row
            If .rows > 1 Then
                If index = 0 Then
                    Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("����"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("ҩƷid"))), .TextMatrix(.Row, .ColIndex("����ʱ��")), Val(.TextMatrix(.Row, .ColIndex("����id"))), bln���±�־, int��˱�־)
                Else
                    Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("����"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("ҩƷid"))), .TextMatrix(.Row, .ColIndex("���ʱ��")), Val(.TextMatrix(.Row, .ColIndex("����id"))), bln���±�־, int��˱�־)
                End If
            End If
        End If
    End With

End Sub

Private Sub vsfDetail_EnterCell(index As Integer)
    
    If index = 1 Then
        If vsfDetail(index).TextMatrix(vsfDetail(index).Row, vsfDetail(index).ColIndex("��˱�־")) = "��" Then
            vsfBatch(index).TextMatrix(0, vsfBatch(index).ColIndex("��������")) = "����������"
        Else
            vsfBatch(index).TextMatrix(0, vsfBatch(index).ColIndex("��������")) = "��������"
        End If
    
        With vsfDetail(index)
            If .Row = 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("��˱�־")) = "��" Then
                cmdVerify.Enabled = True
            Else
                cmdVerify.Enabled = False
            End If
        End With
    End If
End Sub


Private Sub vsfDetail_RowColChange(index As Integer)
    '�ƶ���һ���ı�ǵ���ǰ�У�
    With vsfDetail(index)
        .Cell(flexcpText, 0, 0, .rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
            
            '��ȡ������ϸ����
            If mlngDetailRow <> .Row Then
                mlngDetailRow = .Row
                If .rows > 1 Then
                    If index = 0 Then
                        Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("����"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("ҩƷid"))), .TextMatrix(.Row, .ColIndex("����ʱ��")), Val(.TextMatrix(.Row, .ColIndex("����id"))), False, 0)
                    Else
                        Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("����"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("ҩƷid"))), .TextMatrix(.Row, .ColIndex("���ʱ��")), Val(.TextMatrix(.Row, .ColIndex("����id"))), False, 0)
                    End If
                End If
            End If
        End If
    End With
End Sub


Private Sub vsfList_AfterSort(index As Integer, ByVal Col As Long, Order As Integer)
    If index = 0 Then
        If vsfList(0).rows < 2 Then Exit Sub
        '��ʾ���ʽ��ϼ���Ϣ
        vsfList(0).rows = vsfList(0).rows + 1
        vsfList(0).Cell(flexcpText, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = "���˽��ϼƣ�" & FormatEx(mdblSum, 5)
        vsfList(0).Cell(flexcpFontBold, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = True
        vsfList(0).Cell(flexcpForeColor, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = vbRed
        vsfList(0).Cell(flexcpAlignment, vsfList(0).rows - 1, 1, vsfList(0).rows - 1, vsfList(0).Cols - 1) = flexAlignLeftCenter
        vsfList(0).MergeCells = flexMergeRestrictRows
        vsfList(0).MergeRow(vsfList(0).rows - 1) = True
    End If
End Sub

Private Sub vsfList_BeforeSort(index As Integer, ByVal Col As Long, Order As Integer)
    If index = 0 Then
        If vsfList(0).rows > 2 Then vsfList(0).RemoveItem vsfList(0).rows - 1
    End If
End Sub

Private Sub vsfList_Click(index As Integer)
    Dim bln���±�־ As Boolean
    Dim int��˱�־ As Integer
    
    With vsfList(index)
        If index = 1 Then Exit Sub
        If .Row = .rows - 1 Then Exit Sub
        If .Row < 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("�������")) = "" Then Exit Sub
        
        If index = 0 And .MouseCol = .ColIndex("��˱�־") Then
            If .TextMatrix(.Row, .ColIndex("��˱�־")) = "��" Then
                .TextMatrix(.Row, .ColIndex("��˱�־")) = IIf(mint�Ƿ�������ʾܾ� = 1, "��", "")
                int��˱�־ = IIf(mint�Ƿ�������ʾܾ� = 1, 2, 0)
            ElseIf .TextMatrix(.Row, .ColIndex("��˱�־")) = "��" Then
                .TextMatrix(.Row, .ColIndex("��˱�־")) = ""
                int��˱�־ = 0
            Else
                If Val(.TextMatrix(.Row, .ColIndex("׼������"))) < Val(.TextMatrix(.Row, .ColIndex("��������"))) Then
                    If mint�Ƿ�������ʾܾ� = 1 Then
                        .TextMatrix(.Row, .ColIndex("��˱�־")) = "��"
                        int��˱�־ = 2
                    End If
                Else
                    .TextMatrix(.Row, .ColIndex("��˱�־")) = "��"
                    int��˱�־ = 1
                End If
            End If
            bln���±�־ = True
            
            '���¼�¼����˱��
            mrsDetail.Filter = "����=" & Val(.TextMatrix(.Row, .ColIndex("����"))) & _
                " And No='" & .TextMatrix(.Row, .ColIndex("NO")) & "' " & _
                " And ҩƷID=" & Val(.TextMatrix(.Row, .ColIndex("ҩƷid"))) & _
                " And ����ID=" & Val(.TextMatrix(.Row, .ColIndex("����id"))) & _
                " And ����ʱ��='" & .TextMatrix(.Row, .ColIndex("����ʱ��")) & "' "
            If mrsDetail.RecordCount > 0 Then
                mrsDetail!��˱�־ = int��˱�־
                mrsDetail.Update
            End If
        End If
        
        mblnAllowChange = False
        If Val(.TextMatrix(.Row, .ColIndex("׼������"))) > Val(.TextMatrix(.Row, .ColIndex("��������"))) Then
            mblnAllowChange = True
        End If
        
        '��ȡ������ϸ����
        If mlngListRow <> .Row Or bln���±�־ = True Then
            mlngListRow = .Row
            If .rows > 1 Then
                If index = 0 Then
                    Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("����"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("ҩƷid"))), .TextMatrix(.Row, .ColIndex("����ʱ��")), Val(.TextMatrix(.Row, .ColIndex("����id"))), bln���±�־, int��˱�־)
                Else
                    Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("����"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("ҩƷid"))), .TextMatrix(.Row, .ColIndex("���ʱ��")), Val(.TextMatrix(.Row, .ColIndex("����id"))), bln���±�־, int��˱�־)
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfList_EnterCell(index As Integer)
    If index = 1 Then
        With vsfList(index)
            If .Row = 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("��˱�־")) = "��" Then
                cmdVerify.Enabled = True
            Else
                cmdVerify.Enabled = False
            End If
        End With
    End If
End Sub

Private Sub vsfList_RowColChange(index As Integer)
    '�ƶ���һ���ı�ǵ���ǰ�У�
    With vsfList(index)
        If .Row < 1 Then Exit Sub
        .Cell(flexcpText, 0, 0, .rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
            
            '��ȡ������ϸ����
            If mlngListRow <> .Row Then
                mlngListRow = .Row
                If .rows > 1 Then
                    If index = 0 Then
                        Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("����"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("ҩƷid"))), .TextMatrix(.Row, .ColIndex("����ʱ��")), Val(.TextMatrix(.Row, .ColIndex("����id"))), False, 0)
                    Else
                        Call LoadBatchList(index, Val(.TextMatrix(.Row, .ColIndex("����"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("ҩƷid"))), .TextMatrix(.Row, .ColIndex("���ʱ��")), Val(.TextMatrix(.Row, .ColIndex("����id"))), False, 0)
                    End If
                End If
            End If
        End If
    End With
End Sub


Private Sub vsfMain_AfterSort(index As Integer, ByVal Col As Long, Order As Integer)
    If index = 0 Then
        If vsfMain(index).rows < 2 Then Exit Sub
        vsfMain(index).rows = vsfMain(index).rows + 1
        vsfMain(index).Cell(flexcpText, vsfMain(index).rows - 1, 1, vsfMain(index).rows - 1, vsfMain(index).Cols - 1) = "���˽��ϼƣ�" & FormatEx(mdblSum, 5)
        vsfMain(index).Cell(flexcpFontBold, vsfMain(index).rows - 1, 1, vsfMain(index).rows - 1, vsfMain(index).Cols - 1) = True
        vsfMain(index).Cell(flexcpForeColor, vsfMain(index).rows - 1, 1, vsfMain(index).rows - 1, vsfMain(index).Cols - 1) = vbRed
        vsfMain(index).MergeCells = flexMergeRestrictRows
        vsfMain(index).MergeRow(vsfMain(index).rows - 1) = True
    End If
End Sub

Private Sub vsfMain_BeforeSort(index As Integer, ByVal Col As Long, Order As Integer)
    If vsfMain(index).rows <= 2 Then Exit Sub
    If index = 0 Then
        vsfMain(index).RemoveItem vsfMain(index).rows - 1
    End If
End Sub

Private Sub vsfMain_EnterCell(index As Integer)
    With vsfMain(index)
        If .Row < 1 Or (.Row = .rows - 1 And index = 0) Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("�շ�ϸĿid")) = "" Then Exit Sub
        
        '��ȡ��ϸ����
        If mlngMainRow = .Row Then Exit Sub
        mlngMainRow = .Row
        Call LoadDetailList(index, Val(.TextMatrix(.Row, .ColIndex("�շ�ϸĿid"))))
    End With
    
    '��ȡ������ϸ����
    If vsfDetail(index).rows >= 2 Then
        If index = 0 Then
            Call LoadBatchList(index, Val(vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("����"))), vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("NO")), Val(vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("ҩƷid"))), vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("����ʱ��")), Val(vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("����id"))), False, IIf(optListType(0).Value = True, 1, 0))
        Else
            Call LoadBatchList(index, Val(vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("����"))), vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("NO")), Val(vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("ҩƷid"))), vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("���ʱ��")), Val(vsfDetail(index).TextMatrix(1, vsfDetail(index).ColIndex("����id"))), False, IIf(optListType(0).Value = True, 1, 0))
        End If
    End If
End Sub


Private Sub vsfMain_RowColChange(index As Integer)
    '�ƶ���һ���ı�ǵ���ǰ�У�
    With vsfMain(index)
        .Cell(flexcpText, 0, 0, .rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
        End If
    End With
End Sub


