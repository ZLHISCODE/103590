VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPacsInterfaceDemo 
   Caption         =   "PACS�ӿڵ���"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10200
   Icon            =   "frmPacsInterfaceDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   10200
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtTemp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   9120
      TabIndex        =   103
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComDlg.CommonDialog dlgImage 
      Left            =   9480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "��ȡ��鲿λ"
      TabPicture(0)   =   "frmPacsInterfaceDemo.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vsf��鲿λ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmd��鲿λ"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "��ȡ������Ϣ"
      TabPicture(1)   =   "frmPacsInterfaceDemo.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vsf������Ϣ"
      Tab(1).Control(1)=   "cmd������Ϣ"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "��ȡ���뵥"
      TabPicture(2)   =   "frmPacsInterfaceDemo.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmd�鿴������Ϣ(1)"
      Tab(2).Control(1)=   "cmd�鿴������Ϣ(0)"
      Tab(2).Control(2)=   "txtQueryValue"
      Tab(2).Control(3)=   "dtpStartTime"
      Tab(2).Control(4)=   "cboQueryStyle"
      Tab(2).Control(5)=   "vsf����"
      Tab(2).Control(6)=   "vsf������Ŀ"
      Tab(2).Control(7)=   "vsf���뵥"
      Tab(2).Control(8)=   "dtpEndTime"
      Tab(2).Control(9)=   "Label13"
      Tab(2).Control(10)=   "Label12"
      Tab(2).Control(11)=   "Label11"
      Tab(2).Control(12)=   "Line1"
      Tab(2).Control(13)=   "Label10"
      Tab(2).Control(14)=   "Label9"
      Tab(2).Control(15)=   "Label8"
      Tab(2).ControlCount=   16
      TabCaption(3)   =   "��ȡ������Ϣ"
      TabPicture(3)   =   "frmPacsInterfaceDemo.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label14"
      Tab(3).Control(1)=   "Label15"
      Tab(3).Control(2)=   "vsfGetPatInfo"
      Tab(3).Control(3)=   "cboQueryType"
      Tab(3).Control(4)=   "txtQueryValues"
      Tab(3).Control(5)=   "cmd������Ϣ"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "�������"
      TabPicture(4)   =   "frmPacsInterfaceDemo.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).Control(1)=   "Frame2"
      Tab(4).Control(2)=   "Frame1"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "PACS���汣��"
      TabPicture(5)   =   "frmPacsInterfaceDemo.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdDelReport"
      Tab(5).Control(1)=   "cmdSaveReport"
      Tab(5).Control(2)=   "txtReportDoc"
      Tab(5).Control(3)=   "txtAdviceID1"
      Tab(5).Control(4)=   "Frame6"
      Tab(5).Control(5)=   "Frame5"
      Tab(5).Control(6)=   "Frame4"
      Tab(5).Control(7)=   "Label29"
      Tab(5).Control(8)=   "Label28"
      Tab(5).ControlCount=   9
      TabCaption(6)   =   "�ĵ籨�汣��"
      TabPicture(6)   =   "frmPacsInterfaceDemo.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "txtReportTitle"
      Tab(6).Control(1)=   "cmdɾ������"
      Tab(6).Control(2)=   "cmd���汨��"
      Tab(6).Control(3)=   "txt����ҽ��"
      Tab(6).Control(4)=   "txtAdviceID2"
      Tab(6).Control(5)=   "Frame8"
      Tab(6).Control(6)=   "Frame7"
      Tab(6).Control(7)=   "Label34"
      Tab(6).Control(8)=   "Label33"
      Tab(6).Control(9)=   "Label32"
      Tab(6).ControlCount=   10
      Begin VB.TextBox txtReportTitle 
         Height          =   300
         Left            =   -73860
         TabIndex        =   98
         Top             =   5160
         Width           =   5055
      End
      Begin VB.CommandButton cmdɾ������ 
         Caption         =   "ɾ������(&D)"
         Height          =   300
         Left            =   -66480
         TabIndex        =   96
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmd���汨�� 
         Caption         =   "���汨��(&S)"
         Height          =   300
         Left            =   -67920
         TabIndex        =   95
         Top             =   5880
         Width           =   1215
      End
      Begin VB.TextBox txt����ҽ�� 
         Height          =   300
         Left            =   -70620
         TabIndex        =   94
         Top             =   5880
         Width           =   1815
      End
      Begin VB.TextBox txtAdviceID2 
         Height          =   300
         Left            =   -74040
         TabIndex        =   92
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Frame Frame8 
         Caption         =   "����ͼ��"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   88
         Top             =   3240
         Width           =   9735
         Begin VB.TextBox txt����ͼ�� 
            Height          =   900
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   90
            Top             =   300
            Width           =   8175
         End
         Begin VB.CommandButton cmd���ͼ��1 
            Caption         =   "���ͼ��(&I)"
            Height          =   975
            Left            =   8400
            TabIndex        =   89
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "�����Ϣ"
         Height          =   2715
         Left            =   -74880
         TabIndex        =   83
         Top             =   420
         Width           =   9735
         Begin VB.TextBox txt������� 
            Height          =   1020
            Left            =   1080
            TabIndex        =   85
            Top             =   360
            Width           =   8535
         End
         Begin VB.TextBox txt������ 
            Height          =   1020
            Left            =   1080
            TabIndex        =   84
            Top             =   1560
            Width           =   8535
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "���������"
            Height          =   180
            Left            =   120
            TabIndex        =   87
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Left            =   120
            TabIndex        =   86
            Top             =   1560
            Width           =   900
         End
      End
      Begin VB.CommandButton cmdDelReport 
         Caption         =   "ɾ������(&D)"
         Height          =   300
         Left            =   -66480
         TabIndex        =   78
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveReport 
         Caption         =   "���汨��(&S)"
         Height          =   300
         Left            =   -67920
         TabIndex        =   77
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox txtReportDoc 
         Height          =   300
         Left            =   -70620
         TabIndex        =   76
         Top             =   6240
         Width           =   1815
      End
      Begin VB.TextBox txtAdviceID1 
         Height          =   300
         Left            =   -74040
         TabIndex        =   74
         Top             =   6240
         Width           =   1935
      End
      Begin VB.Frame Frame6 
         Caption         =   "���渽��"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   72
         Top             =   4680
         Width           =   9735
         Begin VB.CommandButton cmd���ͼ�� 
            Caption         =   "���ͼ��(&I)"
            Height          =   975
            Left            =   8400
            TabIndex        =   82
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txt���渽�� 
            Height          =   900
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   81
            Top             =   300
            Width           =   8175
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "����ͼ��"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   71
         Top             =   3240
         Width           =   9735
         Begin VB.CommandButton cmdAddImg 
            Caption         =   "���ͼ��(&I)"
            Height          =   975
            Left            =   8400
            TabIndex        =   80
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtReportImage 
            Height          =   900
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   79
            Top             =   300
            Width           =   8175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "�����Ϣ"
         Height          =   2715
         Left            =   -74880
         TabIndex        =   66
         Top             =   420
         Width           =   9735
         Begin VB.TextBox txtAvice 
            Height          =   1020
            Left            =   1080
            TabIndex        =   70
            Top             =   1560
            Width           =   8535
         End
         Begin VB.TextBox txtView 
            Height          =   1020
            Left            =   1080
            TabIndex        =   68
            Top             =   360
            Width           =   8535
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Left            =   120
            TabIndex        =   69
            Top             =   1560
            Width           =   900
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "���������"
            Height          =   180
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "�ֲ�λִ��ҽ��"
         Height          =   855
         Left            =   -74880
         TabIndex        =   64
         Top             =   3360
         Width           =   9735
         Begin VB.TextBox txtAdviceID 
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   102
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdAccept 
            Caption         =   "��������(&R)"
            Height          =   300
            Index           =   1
            Left            =   4440
            TabIndex        =   101
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdModfiy 
            Caption         =   "�޸�����(&R)"
            Height          =   300
            Index           =   1
            Left            =   5760
            TabIndex        =   100
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "��������(&R)"
            Height          =   300
            Index           =   1
            Left            =   7080
            TabIndex        =   99
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "ҽ��ID��"
            Height          =   180
            Left            =   480
            TabIndex        =   65
            Top             =   405
            Width           =   720
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "����ִ��ҽ��"
         Height          =   855
         Left            =   -74880
         TabIndex        =   58
         Top             =   2280
         Width           =   9735
         Begin VB.CommandButton cmdCancel 
            Caption         =   "��������(&R)"
            Height          =   300
            Index           =   0
            Left            =   7080
            TabIndex        =   63
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdModfiy 
            Caption         =   "�޸�����(&R)"
            Height          =   300
            Index           =   0
            Left            =   5760
            TabIndex        =   62
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdAccept 
            Caption         =   "��������(&R)"
            Height          =   300
            Index           =   0
            Left            =   4440
            TabIndex        =   61
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtAdviceID 
            Height          =   300
            Index           =   0
            Left            =   1200
            TabIndex        =   60
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "ҽ��ID��"
            Height          =   180
            Left            =   480
            TabIndex        =   59
            Top             =   405
            Width           =   720
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1635
         Left            =   -74880
         TabIndex        =   41
         Top             =   420
         Width           =   9735
         Begin VB.TextBox txtDoctor 
            Height          =   300
            Left            =   7800
            TabIndex        =   57
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtDevice 
            Height          =   300
            Left            =   7800
            TabIndex        =   56
            Text            =   "Philips CT"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtExplain 
            Height          =   300
            Left            =   4440
            TabIndex        =   52
            Top             =   1200
            Width           =   5055
         End
         Begin VB.TextBox txtWeight 
            Height          =   300
            Left            =   4440
            TabIndex        =   50
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtStudyNum 
            Height          =   300
            Left            =   4440
            TabIndex        =   49
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtHeight 
            Height          =   300
            Left            =   1200
            TabIndex        =   45
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtExeRoom 
            Height          =   300
            Left            =   1200
            TabIndex        =   44
            Top             =   240
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker dtpExecutionTime 
            Height          =   300
            Left            =   1200
            TabIndex        =   53
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Format          =   99418113
            CurrentDate     =   41816
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "���ҽ����"
            Height          =   180
            Left            =   6840
            TabIndex        =   55
            Top             =   765
            Width           =   900
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "����豸��"
            Height          =   180
            Left            =   6840
            TabIndex        =   54
            Top             =   285
            Width           =   900
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "ִ��˵����"
            Height          =   180
            Left            =   3480
            TabIndex        =   51
            Top             =   1245
            Width           =   900
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "��    �أ�"
            Height          =   180
            Left            =   3480
            TabIndex        =   48
            Top             =   765
            Width           =   900
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "�� �� �ţ�"
            Height          =   180
            Left            =   3480
            TabIndex        =   47
            Top             =   285
            Width           =   900
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "ִ�����ڣ�"
            Height          =   180
            Left            =   240
            TabIndex        =   46
            Top             =   1245
            Width           =   900
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "��    �ߣ�"
            Height          =   180
            Left            =   240
            TabIndex        =   43
            Top             =   765
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "ִ�п��ң�"
            Height          =   180
            Left            =   240
            TabIndex        =   42
            Top             =   285
            Width           =   900
         End
      End
      Begin VB.CommandButton cmd������Ϣ 
         Caption         =   "��ѯ������Ϣ(&P)"
         Height          =   300
         Left            =   -67200
         TabIndex        =   40
         Top             =   6290
         Width           =   1695
      End
      Begin VB.TextBox txtQueryValues 
         Height          =   300
         Left            =   -70800
         TabIndex        =   39
         Top             =   6290
         Width           =   2655
      End
      Begin VB.ComboBox cboQueryType 
         Height          =   300
         ItemData        =   "frmPacsInterfaceDemo.frx":03CE
         Left            =   -73920
         List            =   "frmPacsInterfaceDemo.frx":03E7
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   6285
         Width           =   1695
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfGetPatInfo 
         Height          =   5715
         Left            =   -74880
         TabIndex        =   35
         Top             =   420
         Width           =   9735
         _cx             =   17171
         _cy             =   10081
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
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
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
      Begin VB.CommandButton cmd�鿴������Ϣ 
         Caption         =   "�鿴������Ϣ"
         Height          =   300
         Index           =   1
         Left            =   -67680
         TabIndex        =   34
         Top             =   6290
         Width           =   1695
      End
      Begin VB.CommandButton cmd�鿴������Ϣ 
         Caption         =   "�鿴������Ϣ"
         Height          =   300
         Index           =   0
         Left            =   -67680
         TabIndex        =   33
         Top             =   5760
         Width           =   1695
      End
      Begin VB.TextBox txtQueryValue 
         Height          =   300
         Left            =   -70800
         TabIndex        =   32
         Top             =   5760
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   300
         Left            =   -73980
         TabIndex        =   28
         Top             =   6290
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   99418113
         CurrentDate     =   41816
      End
      Begin VB.ComboBox cboQueryStyle 
         Height          =   300
         ItemData        =   "frmPacsInterfaceDemo.frx":0425
         Left            =   -73980
         List            =   "frmPacsInterfaceDemo.frx":0441
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   5760
         Width           =   1695
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf���� 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   24
         Top             =   4440
         Width           =   9735
         _cx             =   17171
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
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
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
      Begin VSFlex8Ctl.VSFlexGrid vsf������Ŀ 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   22
         Top             =   2880
         Width           =   9735
         _cx             =   17171
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
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
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
      Begin VSFlex8Ctl.VSFlexGrid vsf���뵥 
         Height          =   2115
         Left            =   -74880
         TabIndex        =   20
         Top             =   420
         Width           =   9735
         _cx             =   17171
         _cy             =   3731
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
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
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
      Begin VB.CommandButton cmd������Ϣ 
         Caption         =   "��ȡ��������Ϣ(&D)"
         Height          =   300
         Left            =   -74880
         TabIndex        =   19
         Top             =   6240
         Width           =   1935
      End
      Begin VB.CommandButton cmd��鲿λ 
         Caption         =   "��ȡ��鲿λ(&S)"
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   6240
         Width           =   1695
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf������Ϣ 
         Height          =   5715
         Left            =   -74880
         TabIndex        =   18
         Top             =   420
         Width           =   9735
         _cx             =   17171
         _cy             =   10081
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
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
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
      Begin MSComCtl2.DTPicker dtpEndTime 
         Height          =   300
         Left            =   -70800
         TabIndex        =   30
         Top             =   6290
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   99483649
         CurrentDate     =   41816
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf��鲿λ 
         Height          =   5715
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Width           =   9735
         _cx             =   17171
         _cy             =   10081
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
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
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
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "������⣺"
         Height          =   180
         Left            =   -74760
         TabIndex        =   97
         Top             =   5205
         Width           =   900
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "����ҽ����"
         Height          =   180
         Left            =   -71520
         TabIndex        =   93
         Top             =   5925
         Width           =   900
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��ID��"
         Height          =   180
         Left            =   -74760
         TabIndex        =   91
         Top             =   5925
         Width           =   720
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "����ҽ����"
         Height          =   180
         Left            =   -71520
         TabIndex        =   75
         Top             =   6285
         Width           =   900
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��ID��"
         Height          =   180
         Left            =   -74760
         TabIndex        =   73
         Top             =   6285
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "��ѯֵ��"
         Height          =   180
         Left            =   -71520
         TabIndex        =   38
         Top             =   6320
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "��ѯ���ͣ�"
         Height          =   180
         Left            =   -74880
         TabIndex        =   36
         Top             =   6320
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "��ѯֵ��"
         Height          =   180
         Left            =   -71580
         TabIndex        =   31
         Top             =   5805
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "�������ڣ�"
         Height          =   180
         Left            =   -71760
         TabIndex        =   29
         Top             =   6320
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "��ʼ���ڣ�"
         Height          =   180
         Left            =   -74880
         TabIndex        =   27
         Top             =   6320
         Width           =   900
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         DrawMode        =   5  'Not Copy Pen
         X1              =   -74880
         X2              =   -65160
         Y1              =   6160
         Y2              =   6160
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "��ѯ���ͣ�"
         Height          =   180
         Left            =   -74880
         TabIndex        =   25
         Top             =   5800
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "������ϸ��"
         Height          =   180
         Left            =   -74880
         TabIndex        =   23
         Top             =   4200
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "������Ŀ��ϸ��"
         Height          =   180
         Left            =   -74880
         TabIndex        =   21
         Top             =   2640
         Width           =   1260
      End
   End
   Begin VB.Frame fraDataBaseInfo 
      Caption         =   "��ʼ�����ݿ�����"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.CommandButton cmdConnOra 
         Caption         =   "�������ݿ�(&C)"
         Height          =   300
         Left            =   8040
         TabIndex        =   14
         Top             =   1280
         Width           =   1695
      End
      Begin VB.TextBox txtOwner 
         Height          =   300
         Left            =   8040
         TabIndex        =   13
         Text            =   "zlhis"
         Top             =   320
         Width           =   1695
      End
      Begin VB.TextBox txtPassW 
         Height          =   300
         Left            =   4800
         TabIndex        =   12
         Text            =   "aqa"
         Top             =   800
         Width           =   1695
      End
      Begin VB.TextBox txtSys 
         Height          =   300
         Left            =   4800
         TabIndex        =   11
         Text            =   "100"
         Top             =   320
         Width           =   1695
      End
      Begin VB.TextBox txtDeptID 
         Height          =   300
         Left            =   1680
         TabIndex        =   6
         Text            =   "64"
         Top             =   1280
         Width           =   1695
      End
      Begin VB.TextBox txtUser 
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Text            =   "zlhis"
         Top             =   800
         Width           =   1695
      End
      Begin VB.TextBox txtServer 
         Height          =   300
         Left            =   1680
         TabIndex        =   3
         Text            =   "zlhis"
         Top             =   320
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "���ڲ��ſ���¼�뵱ǰ���Ҷ�Ӧ�Ĳ���ID"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3480
         TabIndex        =   10
         Top             =   1320
         Width           =   3240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "���ݿ������ߣ�"
         Height          =   180
         Left            =   6720
         TabIndex        =   9
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Oracle���룺"
         Height          =   180
         Left            =   3720
         TabIndex        =   8
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ϵͳ�ţ�"
         Height          =   180
         Left            =   4080
         TabIndex        =   7
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ����ID��"
         Height          =   180
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Oracle�û����ƣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Oracleʵ�����ƣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmPacsInterfaceDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'������ʾ����
Public Enum TErrorShowType
    estNoDisplay = 1
    estShowMsg = 2
End Enum

Private mobjPacsInterface As Object
Private mstrErrInfo As String

Private Sub Form_Load()
    If cboQueryStyle.ListIndex < 0 And cboQueryStyle.ListCount > 0 Then cboQueryStyle.ListIndex = 0
    
    If cboQueryType.ListIndex < 0 And cboQueryType.ListCount > 0 Then cboQueryType.ListIndex = 0
End Sub

Public Function IsInit() As Boolean
    If mobjPacsInterface Is Nothing Then
        Call ShowErrMessage("���ݿ�δ���ӳɹ���ʧ��")
        Exit Function
    End If
    
    IsInit = True
End Function

Public Function ShowErrMessage(ByVal strMsg As String) As Boolean
    If strMsg <> "" Then
        MsgBox strMsg, vbExclamation, "PACS�ӿڲ��Թ���"
        ShowErrMessage = True
        Exit Function
    End If
    
    mstrErrInfo = mobjPacsInterface.GetLastError
    
    If mstrErrInfo = "" Then
        Exit Function
    Else
        ShowErrMessage = True
        MsgBox mstrErrInfo, vbExclamation, "PACS�ӿڲ��Թ���"
    End If
    
    mstrErrInfo = ""
End Function

Public Sub LoadDataToVSF(ByVal blnLoad As Boolean, ByVal vsfTable As VSFlexGrid)
    If Not blnLoad Then Exit Sub
    
    If Not ShowErrMessage("") Then
        Call ReadQueryData(mobjPacsInterface, vsfTable)
    End If
End Sub

Public Sub ReadQueryData(ByVal objPacsInterface As Object, ByVal vsfTable As VSFlexGrid)
    Dim lngColumnCount As Long
    Dim lngRecordCount As Long
    Dim i As Long, j As Long
    
    If objPacsInterface Is Nothing Then Exit Sub
    
    lngColumnCount = objPacsInterface.GetCurColumnCount
    lngRecordCount = objPacsInterface.GetCurRecordCount
    
    With vsfTable
        '���������
        .Cols = 0
        .Rows = 0
        
        If lngColumnCount <= 0 Then Exit Sub
        
        .Cols = lngColumnCount
        .Rows = lngRecordCount + 1
        .FixedRows = 1
        .ExtendLastCol = True
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeColumns
        .GridColor = &H80000010
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        
        '������ͷ��ָ���п�
        For i = 0 To lngColumnCount - 1
            .TextMatrix(0, i) = objPacsInterface.GetCurColumnName(i)
            .ColWidth(i) = 800
        Next
        
        If lngRecordCount <= 0 Then Exit Sub
        
        '��������
        For i = 0 To lngRecordCount - 1
            For j = 0 To lngColumnCount - 1
                .TextMatrix(i + 1, j) = objPacsInterface.GetCurValueByColumnName(i, vsfTable.TextMatrix(0, j))
            Next
        Next
    End With
End Sub

Private Sub cmdConnOra_Click()
'��ʼ��ipacs�ӿ�
    Dim strResult As String
    
    If Trim(txtDeptID.Text) = "" Then
        MsgBox "�����뵱ǰ����ID", vbExclamation, ""
        txtDeptID.SetFocus
        Exit Sub
    End If
    
    Set mobjPacsInterface = CreateObject("zlPacsInterface.clsPacsInterface")
    
    Call mobjPacsInterface.InitInterface(Val(txtDeptID.Text), Trim(txtServer.Text), Trim(txtUser.Text), Trim(txtPassW.Text), Val(txtSys.Text), Trim(txtOwner.Text), "", "~", estNoDisplay)
    
    strResult = mobjPacsInterface.GetLastError
    
    If strResult <> "" Then
        Call ShowErrMessage(strResult)
        Exit Sub
    End If
    
    Call ShowErrMessage("�ѳɹ����ӵ�Oracle���ݿ�:" & IIf(Trim(txtServer.Text) = "", "Local", Trim(txtServer.Text)))
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��ȡ��鲿λ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmd��鲿λ_Click()
'��ȡ��鲿λ
    If Not IsInit Then Exit Sub
    
    Call LoadDataToVSF(mobjPacsInterface.GetPacsItems(""), vsf��鲿λ)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��ȡ������Ϣ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmd������Ϣ_Click()
'��ȡ��������Ϣ
    If Not IsInit Then Exit Sub
    
    Call LoadDataToVSF(mobjPacsInterface.GetDeptItems(""), vsf������Ϣ)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��ȡ���뵥''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmd�鿴������Ϣ_Click(Index As Integer)
'�����ݿ��в�ѯ���뵥��Ϣ
    If Not IsInit Then Exit Sub
    
    If Index = 0 Then   '����ѯ������Ϣ
        Call LoadDataToVSF(mobjPacsInterface.GetRequestInfo(Trim(txtQueryValue.Text), cboQueryStyle.ItemData(cboQueryStyle.ListIndex)), vsf���뵥)
    
    Else                '����ѯ������Ϣ
        Call LoadDataToVSF(mobjPacsInterface.GetRequestInfo1(dtpStartTime.Value, dtpEndTime.Value, "�ĵ�"), vsf���뵥)
    End If
    
    If vsf���뵥.Rows >= 2 Then vsf���뵥.RowSel = 1
End Sub

Private Sub vsf���뵥_SelChange()
    If Not IsInit Then Exit Sub
    
    '��д������Ŀ��ϸ
    Call LoadDataToVSF(mobjPacsInterface.GetAdviceItems(vsf���뵥.TextMatrix(vsf���뵥.RowSel, 0)), vsf������Ŀ)
    
    '��д���ü�¼
    Call LoadDataToVSF(mobjPacsInterface.GetAdviceFees(vsf���뵥.TextMatrix(vsf���뵥.RowSel, 0)), vsf����)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��ȡ������Ϣ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmd������Ϣ_Click()
'��ȡ������Ϣ
    If Not IsInit Then Exit Sub
    
    Call LoadDataToVSF(mobjPacsInterface.GetPatientInfo(Trim(txtQueryValues.Text), cboQueryType.ItemData(cboQueryType.ListIndex)), vsfGetPatInfo)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�������''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAccept_Click(Index As Integer)
'��������
    Dim lngAdviceID As Long
    
    If Not IsInit Then Exit Sub
    
    If Trim(txtAdviceID(Index).Text) = "" Then
        MsgBox "��¼����Ч��ҽ��ID��", vbExclamation, ""
        txtAdviceID(Index).SetFocus
        Exit Sub
    End If
    
    lngAdviceID = Val(txtAdviceID(Index).Text)
    
    Call mobjPacsInterface.RecevieRequest(lngAdviceID, Trim(txtExeRoom.Text), CLng(Val(txtStudyNum.Text)), Trim(txtDevice.Text), CLng(Val(txtHeight.Text)), CLng(Val(txtWeight.Text)), Trim(txtDoctor.Text), dtpExecutionTime.Value, Trim(txtExplain.Text), Index)
    
    Call ShowErrMessage("���������ղ���!")
End Sub

Private Sub cmdModfiy_Click(Index As Integer)
'�޸�����
    Dim lngAdviceID As Long
    
    If Not IsInit Then Exit Sub
    
    If Trim(txtAdviceID(Index).Text) = "" Then
        MsgBox "��¼����Ч��ҽ��ID��", vbExclamation, ""
        txtAdviceID(Index).SetFocus
        Exit Sub
    End If
    
    lngAdviceID = Val(txtAdviceID(Index).Text)
    
    Call mobjPacsInterface.ModifyRequest(lngAdviceID, Trim(txtExeRoom.Text), CLng(Val(txtStudyNum.Text)), Trim(txtDevice.Text), CLng(Val(txtHeight.Text)), CLng(Val(txtWeight.Text)), Trim(txtDoctor.Text), dtpExecutionTime.Value, Trim(txtExplain.Text), Index)
    
    Call ShowErrMessage("��������޸Ĳ���!")
End Sub

Private Sub cmdCancel_Click(Index As Integer)
'��������
    Dim lngAdviceID As Long
    
    If Not IsInit Then Exit Sub
    
    If Trim(txtAdviceID(Index).Text) = "" Then
        MsgBox "��¼����Ч��ҽ��ID��", vbExclamation, ""
        txtAdviceID(Index).SetFocus
        Exit Sub
    End If
    
    lngAdviceID = Val(txtAdviceID(Index).Text)
    
    Call mobjPacsInterface.CancelRequest(lngAdviceID, Index)

    Call ShowErrMessage("������볷������!")
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PACS���汣��''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddImg_Click()
'���ͼ��,һ�ο�����Ӷ���ͼ
    Dim i As Integer
    Dim strImgPath As String, strImgFiles As String, strImgFile() As String
    Dim strSplit As String
    
    dlgImage.Filter = "(*.bmp)|*.bmp|(*.jpg)|*.jpg|(*.*)|*.*"
    dlgImage.DefaultExt = "*.bmp"
    dlgImage.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer
    dlgImage.ShowOpen
    
    If dlgImage.FileName = "" Then Exit Sub
    
    strImgPath = CurDir
    
    strImgFiles = Replace(dlgImage.FileName, strImgPath, "")
    strSplit = Mid(strImgFiles, 1, 1)
    
    ReDim strImgFile(UBound(Split(strImgFiles, strSplit))) As String
   
    strImgFile = Split(strImgFiles, strSplit)
    
    txtTemp.Text = ""
    For i = 1 To UBound(strImgFile)
        txtTemp.Text = txtTemp.Text & "," & strImgPath & "\" & strImgFile(i)
    Next
    txtTemp.Text = Mid(txtTemp.Text, 2)
    
    txtReportImage.Text = txtReportImage.Text & "," & txtTemp.Text
    
    If Mid(txtReportImage.Text, 1, 1) = "," Then txtReportImage.Text = Mid(txtReportImage.Text, 2)
End Sub

Private Sub cmd���ͼ��_Click()
'��ӱ��渽��
    dlgImage.Filter = "(*.bmp)|*.bmp|(*.jpg)|*.jpg|(*.*)|*.*"
    dlgImage.DefaultExt = "*.bmp"
    dlgImage.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer
    dlgImage.ShowOpen
    
    If dlgImage.FileName = "" Then Exit Sub
    
    txt���渽��.Text = dlgImage.FileName
End Sub

Private Sub cmdSaveReport_Click()
'���汨��
    If Not IsInit Then Exit Sub
    
    If Trim(txtAdviceID1.Text) = "" Then
        Call ShowErrMessage("��¼����Ч��ҽ��ID��")
        Exit Sub
    End If
    
    '��ɾ���Ѿ����ڵı�������,
    Call mobjPacsInterface.DeleteReport(CLng(txtAdviceID1.Text))
    
    '���汨���ı���Ϣ
    Call mobjPacsInterface.SendReport(CLng(txtAdviceID1.Text), txtView.Text, txtAvice.Text, txtReportDoc.Text, "")
    If ShowErrMessage("") Then Exit Sub
    
    '��ӱ���ͼ�� (ʹ�ø÷��������ȵ���SendReport)
    Call mobjPacsInterface.SendReportImages(CLng(txtAdviceID1.Text), txtReportImage.Text)
    Call ShowErrMessage("")
    
    '��ӱ��渽�� (ʹ�ø÷��������ȵ���SendReport)
    Call mobjPacsInterface.SendReportAffix(CLng(txtAdviceID1.Text), txt���渽��.Text)
    Call ShowErrMessage("")
End Sub

Private Sub cmdDelReport_Click()
'ɾ������
    If Not IsInit Then Exit Sub
    
    If Trim(txtAdviceID1.Text) = "" Then
        Call ShowErrMessage("��¼����Ч��ҽ��ID��")
        Exit Sub
    End If
    
    'ɾ���Ѿ����ڵı�������,
    Call mobjPacsInterface.DeleteReport(CLng(txtAdviceID1.Text))
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�ĵ籨�汣��''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmd���ͼ��1_Click()
'���ͼ��,һ�ο�����Ӷ���ͼ
    Dim i As Integer
    Dim strImgPath As String, strImgFiles As String, strImgFile() As String
    Dim strSplit As String
    
    dlgImage.Filter = "(*.bmp)|*.bmp|(*.jpg)|*.jpg|(*.*)|*.*"
    dlgImage.DefaultExt = "*.bmp"
    dlgImage.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer
    dlgImage.ShowOpen
    
    If dlgImage.FileName = "" Then Exit Sub
    
    strImgPath = CurDir
    
    strImgFiles = Replace(dlgImage.FileName, strImgPath, "")
    strSplit = Mid(strImgFiles, 1, 1)
    
    ReDim strImgFile(UBound(Split(strImgFiles, strSplit))) As String
   
    strImgFile = Split(strImgFiles, strSplit)
    
    txtTemp.Text = ""
    For i = 1 To UBound(strImgFile)
        txtTemp.Text = txtTemp.Text & "," & strImgPath & "\" & strImgFile(i)
    Next
    txtTemp.Text = Mid(txtTemp.Text, 2)
    
    txt����ͼ��.Text = txt����ͼ��.Text & "," & txtTemp.Text
    
    If Mid(txt����ͼ��.Text, 1, 1) = "," Then txt����ͼ��.Text = Mid(txt����ͼ��.Text, 2)
End Sub

Private Sub cmd���汨��_Click()
'���汨��
    If Not IsInit Then Exit Sub
    
    If Trim(txtAdviceID2.Text) = "" Then
        Call ShowErrMessage("��¼����Ч��ҽ��ID��")
        Exit Sub
    End If
    
    '��ɾ���Ѿ����ڵı�������,
    Call mobjPacsInterface.DeleteElectrocardioReport(CLng(txtAdviceID2.Text))
    
    '���汨���ı���Ϣ
    Call mobjPacsInterface.SendElectrocardioReport(CLng(txtAdviceID2.Text), txtReportTitle.Text, txt����ͼ��.Text, txt�������.Text, txt������.Text, txt����ҽ��.Text, "")
    Call ShowErrMessage("")
End Sub

Private Sub cmdɾ������_Click()
'ɾ������
    If Not IsInit Then Exit Sub
    
    If Trim(txtAdviceID2.Text) = "" Then
        Call ShowErrMessage("��¼����Ч��ҽ��ID��")
        Exit Sub
    End If
    
    'ɾ���Ѿ����ڵı�������,
    Call mobjPacsInterface.DeleteElectrocardioReport(CLng(txtAdviceID2.Text))
End Sub
