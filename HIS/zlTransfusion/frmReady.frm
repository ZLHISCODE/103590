VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "zlIDKind.ocx"
Begin VB.Form frmReady 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ӵ�"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12795
   Icon            =   "frmReady.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   426
   ScaleMode       =   2  'Point
   ScaleWidth      =   639.75
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   195
      Picture         =   "frmReady.frx":000C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   8075
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   180
      Picture         =   "frmReady.frx":0156
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   83
      Top             =   7755
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   180
      Picture         =   "frmReady.frx":06E0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   81
      Top             =   7420
      Width           =   240
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "ȫѡ(&H)"
      Height          =   300
      Index           =   1
      Left            =   10440
      TabIndex        =   88
      Top             =   2320
      Width           =   990
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "ȫ��(&N)"
      Height          =   300
      Index           =   0
      Left            =   11520
      TabIndex        =   87
      Top             =   2320
      Width           =   990
   End
   Begin VB.Frame Fra 
      Caption         =   "��������"
      Height          =   840
      Left            =   150
      TabIndex        =   0
      Top             =   45
      Width           =   12495
      Begin zlIDKind.IDKindNew idkSelect 
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         ShowSortName    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "����"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
         SaveRegType     =   4
      End
      Begin VB.CommandButton cmdReadCard 
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   345
         Left            =   3450
         TabIndex        =   2
         Top             =   270
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtNo 
         Height          =   300
         Left            =   1455
         TabIndex        =   1
         Top             =   300
         Width           =   1965
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   8970
         TabIndex        =   6
         Top             =   300
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   168361987
         CurrentDate     =   38082
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   6675
         TabIndex        =   5
         Top             =   300
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   168361987
         CurrentDate     =   38082
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ��(&R)"
         Height          =   345
         Left            =   11085
         TabIndex        =   7
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��                         ��"
         Height          =   180
         Left            =   5790
         TabIndex        =   4
         Top             =   345
         Width           =   3150
      End
   End
   Begin MSComctlLib.ImageList imgPic 
      Left            =   9960
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReady.frx":0C6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReady.frx":1204
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReady.frx":179E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReady.frx":1D38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTransfusion 
      Height          =   3780
      Left            =   180
      TabIndex        =   80
      Top             =   3510
      Width           =   12375
      _cx             =   21828
      _cy             =   6667
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
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
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReady.frx":1E92
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
   Begin VB.Frame fraPrint 
      Caption         =   "��ѡ��Ҫ��ӡ�ĵ���"
      Height          =   1000
      Left            =   2400
      TabIndex        =   60
      Top             =   7440
      Width           =   7170
      Begin VB.CheckBox chkWristband 
         Caption         =   "��Һ���"
         Height          =   250
         Left            =   2520
         TabIndex        =   43
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkLabel 
         Caption         =   "��Һƿǩ"
         Height          =   250
         Left            =   1320
         TabIndex        =   42
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkView 
         Caption         =   "Ԥ��"
         Height          =   250
         Left            =   6240
         TabIndex        =   44
         Top             =   600
         Width           =   700
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "���Ƶ�"
         Height          =   250
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "Ƥ�Ե�"
         Height          =   250
         Index           =   3
         Left            =   2520
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "ע�䵥"
         Height          =   250
         Index           =   2
         Left            =   1320
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "��Һ��"
         Height          =   250
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�˳�(&C)"
      Height          =   350
      Left            =   11545
      TabIndex        =   47
      Top             =   7960
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   10320
      TabIndex        =   46
      Top             =   7960
      Width           =   1100
   End
   Begin VB.ComboBox cboOperator 
      Height          =   300
      Index           =   0
      Left            =   10725
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   7440
      Width           =   1920
   End
   Begin VB.Frame frmBaseInfo 
      Caption         =   "������Ϣ"
      Height          =   1365
      Left            =   135
      TabIndex        =   50
      Top             =   930
      Width           =   12525
      Begin VB.ComboBox cboSeating 
         Height          =   300
         Left            =   9525
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   210
         Width           =   2670
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   9
         Left            =   3660
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   49
         Top             =   915
         Width           =   8530
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   8
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   48
         Top             =   915
         Width           =   1680
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   7
         Left            =   9525
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   45
         Top             =   540
         Width           =   2670
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   6
         Left            =   6540
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   37
         Top             =   555
         Width           =   1700
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   5
         Left            =   3660
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   14
         Top             =   585
         Width           =   1700
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   4
         Left            =   945
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   13
         Top             =   585
         Width           =   1700
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   3
         Left            =   6540
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   12
         Top             =   210
         Width           =   1700
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   1
         Left            =   3660
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   11
         Top             =   255
         Width           =   1700
      End
      Begin VB.TextBox txtBase 
         Height          =   300
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   9
         Top             =   255
         Width           =   1700
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "���"
         Height          =   240
         Index           =   9
         Left            =   3105
         TabIndex        =   61
         Top             =   930
         Width           =   525
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "ҽ��"
         Height          =   240
         Index           =   8
         Left            =   210
         TabIndex        =   59
         Top             =   930
         Width           =   720
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "���˿���"
         Height          =   240
         Index           =   7
         Left            =   8730
         TabIndex        =   58
         Top             =   615
         Width           =   780
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         Height          =   240
         Index           =   6
         Left            =   5895
         TabIndex        =   57
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "�Ա�"
         Height          =   240
         Index           =   5
         Left            =   3000
         TabIndex        =   56
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         Height          =   240
         Index           =   4
         Left            =   300
         TabIndex        =   55
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "����ʱ��"
         Height          =   240
         Index           =   3
         Left            =   5745
         TabIndex        =   54
         Top             =   270
         Width           =   780
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "��λ��"
         Height          =   240
         Index           =   2
         Left            =   8880
         TabIndex        =   53
         Top             =   285
         Width           =   630
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "˳���"
         Height          =   240
         Index           =   1
         Left            =   3000
         TabIndex        =   52
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Caption         =   "�Һŵ���"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   51
         Top             =   300
         Width           =   750
      End
   End
   Begin TabDlg.SSTab stabType 
      Height          =   5010
      Left            =   90
      TabIndex        =   8
      Top             =   2370
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   8837
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "(&0)����"
      TabPicture(0)   =   "frmReady.frx":1F2D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblForecastTime"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblOperator(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl����"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtpForecastTime(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboOperator(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtժҪ(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "(&1)��Һ"
      TabPicture(1)   =   "frmReady.frx":1F49
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblOperator(2)"
      Tab(1).Control(1)=   "lblTransfusion(1)"
      Tab(1).Control(2)=   "lblTransfusion(2)"
      Tab(1).Control(3)=   "lblTransfusion(0)"
      Tab(1).Control(4)=   "Label1"
      Tab(1).Control(5)=   "lblOperator(3)"
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(7)=   "dtpForecastTime(1)"
      Tab(1).Control(8)=   "cboOperator(2)"
      Tab(1).Control(9)=   "txtTransfusion(0)"
      Tab(1).Control(10)=   "txtTransfusion(1)"
      Tab(1).Control(11)=   "txtTransfusion(2)"
      Tab(1).Control(12)=   "cboOperator(3)"
      Tab(1).Control(13)=   "txtժҪ(1)"
      Tab(1).Control(14)=   "chk��Һ����"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "(&2)ע��"
      TabPicture(2)   =   "frmReady.frx":1F65
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(1)=   "lblOperator(4)"
      Tab(2).Control(2)=   "lblOperator(5)"
      Tab(2).Control(3)=   "Label11"
      Tab(2).Control(4)=   "dtpForecastTime(2)"
      Tab(2).Control(5)=   "cboOperator(4)"
      Tab(2).Control(6)=   "cboOperator(5)"
      Tab(2).Control(7)=   "txtժҪ(2)"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "(&3)Ƥ��"
      TabPicture(3)   =   "frmReady.frx":1F81
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblScratchTest(0)"
      Tab(3).Control(1)=   "Label3"
      Tab(3).Control(2)=   "lblOperator(6)"
      Tab(3).Control(3)=   "Label9"
      Tab(3).Control(4)=   "dtpForecastTime(3)"
      Tab(3).Control(5)=   "txtScratchTest"
      Tab(3).Control(6)=   "chkƤ������"
      Tab(3).Control(7)=   "cboOperator(6)"
      Tab(3).Control(8)=   "txtժҪ(3)"
      Tab(3).ControlCount=   9
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Index           =   0
         Left            =   1230
         MaxLength       =   100
         TabIndex        =   17
         Top             =   765
         Width           =   9735
      End
      Begin VB.CheckBox chk��Һ���� 
         Caption         =   "��ʱ����"
         Height          =   330
         Left            =   -66405
         TabIndex        =   20
         Top             =   405
         Width           =   1050
      End
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Index           =   2
         Left            =   -73770
         MaxLength       =   100
         TabIndex        =   29
         Top             =   765
         Width           =   9735
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Index           =   5
         Left            =   -71070
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   420
         Width           =   1830
      End
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Index           =   3
         Left            =   -73770
         MaxLength       =   100
         TabIndex        =   34
         Top             =   765
         Width           =   9735
      End
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Index           =   1
         Left            =   -73770
         MaxLength       =   100
         TabIndex        =   22
         Top             =   765
         Width           =   7260
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Index           =   1
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   765
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Index           =   4
         Left            =   -73770
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   765
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Index           =   6
         Left            =   -73770
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   765
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Index           =   3
         Left            =   -65850
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   765
         Width           =   1830
      End
      Begin VB.TextBox txtTransfusion 
         Height          =   300
         Index           =   2
         Left            =   -67365
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   25
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox txtTransfusion 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   -69540
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   420
         Width           =   990
      End
      Begin VB.TextBox txtTransfusion 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   -71265
         MaxLength       =   2
         TabIndex        =   19
         ToolTipText     =   "��ϵ����ȡֵΪ10��15��20"
         Top             =   420
         Width           =   495
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Index           =   2
         Left            =   -73770
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   765
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.CheckBox chkƤ������ 
         Caption         =   "��ʱ����"
         Height          =   330
         Left            =   -69690
         TabIndex        =   32
         Top             =   405
         Width           =   1050
      End
      Begin VB.TextBox txtScratchTest 
         Height          =   300
         Left            =   -70725
         MaxLength       =   5
         TabIndex        =   31
         Top             =   420
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpForecastTime 
         Height          =   300
         Index           =   0
         Left            =   1230
         TabIndex        =   15
         Top             =   420
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   168361987
         CurrentDate     =   38082
      End
      Begin MSComCtl2.DTPicker dtpForecastTime 
         Height          =   300
         Index           =   1
         Left            =   -73770
         TabIndex        =   18
         Top             =   420
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   168361987
         CurrentDate     =   38082
      End
      Begin MSComCtl2.DTPicker dtpForecastTime 
         Height          =   300
         Index           =   2
         Left            =   -73770
         TabIndex        =   26
         Top             =   420
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   168361987
         CurrentDate     =   38082
      End
      Begin MSComCtl2.DTPicker dtpForecastTime 
         Height          =   300
         Index           =   3
         Left            =   -73770
         TabIndex        =   30
         Top             =   420
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   168361987
         CurrentDate     =   38082
      End
      Begin VB.Label lbl���� 
         Caption         =   "ִ��ժҪ"
         Height          =   255
         Left            =   450
         TabIndex        =   63
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label11 
         Caption         =   "ִ��ժҪ"
         Height          =   255
         Left            =   -74550
         TabIndex        =   79
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblOperator 
         Caption         =   "��ҩ��"
         Height          =   240
         Index           =   5
         Left            =   -71670
         TabIndex        =   78
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label9 
         Caption         =   "ִ��ժҪ"
         Height          =   255
         Left            =   -74550
         TabIndex        =   77
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label8 
         Caption         =   "ִ��ժҪ"
         Height          =   255
         Left            =   -74550
         TabIndex        =   76
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblOperator 
         Caption         =   "ִ����"
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   75
         Top             =   840
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblOperator 
         Caption         =   "ִ����"
         Height          =   240
         Index           =   4
         Left            =   -74400
         TabIndex        =   74
         Top             =   840
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblOperator 
         Caption         =   "ִ����"
         Height          =   240
         Index           =   6
         Left            =   -74400
         TabIndex        =   73
         Top             =   840
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblOperator 
         Caption         =   "��ҩ��"
         Height          =   240
         Index           =   3
         Left            =   -66450
         TabIndex        =   72
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label3 
         Caption         =   "Ԥ�ƿ�ʼʱ��"
         Height          =   255
         Left            =   -74895
         TabIndex        =   71
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "Ԥ�ƿ�ʼʱ��"
         Height          =   255
         Left            =   -74895
         TabIndex        =   70
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Ԥ�ƿ�ʼʱ��"
         Height          =   255
         Left            =   -74895
         TabIndex        =   69
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label lblForecastTime 
         Caption         =   "Ԥ�ƿ�ʼʱ��"
         Height          =   255
         Left            =   105
         TabIndex        =   68
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label lblTransfusion 
         Caption         =   "��ϵ��"
         Height          =   255
         Index           =   0
         Left            =   -71880
         TabIndex        =   67
         Top             =   495
         Width           =   555
      End
      Begin VB.Label lblTransfusion 
         Caption         =   "Ԥ��ʱ��(��)"
         Height          =   255
         Index           =   2
         Left            =   -68475
         TabIndex        =   66
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label lblTransfusion 
         Caption         =   "Һ������(ml)"
         Height          =   255
         Index           =   1
         Left            =   -70665
         TabIndex        =   65
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label lblOperator 
         Caption         =   "ִ����"
         Height          =   240
         Index           =   2
         Left            =   -74400
         TabIndex        =   64
         Top             =   840
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblScratchTest 
         Caption         =   "Ԥ��ʱ��(��)"
         Height          =   255
         Index           =   0
         Left            =   -71835
         TabIndex        =   62
         Top             =   495
         Width           =   1095
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      Height          =   180
      Left            =   435
      TabIndex        =   86
      Top             =   8080
      Width           =   540
   End
   Begin VB.Label Label5 
      Caption         =   "����Ҽ��ܾ�"
      Height          =   240
      Left            =   435
      TabIndex        =   84
      Top             =   7770
      Width           =   1185
   End
   Begin VB.Label Label4 
      Caption         =   "�������ӵ�"
      Height          =   240
      Left            =   435
      TabIndex        =   82
      Top             =   7440
      Width           =   1080
   End
   Begin VB.Label lblOperator 
      Caption         =   "�ӵ���"
      Height          =   240
      Index           =   0
      Left            =   10050
      TabIndex        =   35
      Top             =   7500
      Width           =   630
   End
   Begin VB.Menu popMenu 
      Caption         =   "�˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuGNo 
         Caption         =   "�Һŵ�"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuJZK 
         Caption         =   "���￨"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuMNo 
         Caption         =   "�����"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuName 
         Caption         =   "�ա���"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSFZ 
         Caption         =   "���֤"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuICK 
         Caption         =   "�ɣÿ�"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCardSquare 
         Caption         =   "һ��ͨ"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmReady"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Base
    �Һŵ��� = 0
    ˳��� = 1
    ��λ�� = 2
    ����ʱ�� = 3
    ���� = 4
    �Ա� = 5
    ���� = 6
    ���˿��� = 7
    ҽ�� = 8
    ��� = 9
End Enum

Private Enum vsCol
    col_ѡ�� = 0
    col_ִ��˳�� = 1
    col_�ϴ�˳�� = 2
    col_��� = 3
    col_ҽ������ = 4
    col_���� = 5
    col_��λ = 6
    col_��Ŀ��� = 7
    col_ִ��Ƶ�� = 8
    col_�÷� = 9
    col_���� = 10
    col_���� = 11
    col_ʱ�� = 12
    col_ʣ����� = 13
    col_�շѽ�� = 14
    col_ҽ������ = 15
    col_BillKey = 16
    col_groupkey = 17
    col_ִ�мƷ�״̬ = 18
    col_��ϸ�Ʒ�״̬ = 19
End Enum

Private Enum Trans
    ��ϵ�� = 0
    Һ������ = 1
    Ԥ��ʱ�� = 2
End Enum

Private Enum cbo
    �ӵ��� = 0
    ����ִ���� = 1
    ��Һִ���� = 2
    ��Һ��ҩ�� = 3
    ע��ִ���� = 4
    ע����ҩ�� = 5
    Ƥ��ִ���� = 6
End Enum
Private Enum sType
    ���� = 0
    ��Һ = 1
    ע�� = 2
    Ƥ�� = 3
End Enum

Private mPatients As cPatients   '���˼�¼��
Private mPatient As cPatient     '����

Private mSeatings As Seatings    '��λ��¼��
Private mOutNurses As OutNurses  '��ʿ��¼��
Private mGps���� As Groups
Private mGps��Һ As Groups
Private mGpsע�� As Groups
Private mGpsƤ�� As Groups

Private mlng����ID As Long
Private mstrִ�п��� As String
Private mstr��λ As String      '����������ѡ�е���λ

Private mDateBegin As Date
Private mdateEnd As Date
Private mblnHaveData As Boolean '�Ƿ�������
Private mblnOk As Boolean

Private mstrPrivs As String                 'Ȩ��
Private mobjSquareCard  As Object           'һ��ͨ���� add by 2011-12-23

'Private mintFindType As Integer             '�������� 0-���￨,1-�����,2-�Һŵ�,3-����,4-���֤,5-IC��
'Private mstrIDCard As String                '����Զ�ˢ���������֤��
'Private WithEvents mobjIDCard As clsIDCard  '���֤����
Private mobjICCard As Object                'IC������
Private mblnLoad As Boolean                  '�Ƿ��һ����������
Private mblnLiquid As Boolean               '�Ƿ�����Һ����
Private mstrKeyType As String               '�ӵ���ѯ����
Private marrKey As Variant
Private mblnImmediatePuncture As Boolean    '�ӵ���ֱ�ӽ��봩��״̬
Private mblnActivate As Boolean             '�Ƿ��Ѿ�ִ��Activate�¼�
Private mstrSquareCards As String           'һ��ͨ��Ϣ
Private mintLabelState As Integer           '�ϴ���Һƿǩ��״̬
Private mintWristband As Integer            '�ϴ���Һ�����ǩ��״̬
Private mblnReadCard As Boolean
Private mbytType As Byte                    '�ӵ���ʽ
Private mptiInfo As PatiIdentify            '������Ϣ�ؼ�

'�ӵ���ʽ
'��ʽ˵����
Private Const MSTR_MODE As String = "��|�Һŵ�|0;��|���￨|1;��|�����|0;��|����|0;��|���֤��|0;IC|IC��|1"

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChoose_Click(Index As Integer)
    If vsTransfusion.Rows <= 1 Then Exit Sub
    
    Dim i As Integer, intRow As Integer
    
    With vsTransfusion
        intRow = .Row
        .Redraw = False
        For i = 1 To .Rows - 1
            .Row = i
            If .RowData(i) = 0 Or .RowData(i) = 3 Then
                If Index = 1 Then
                    'ͨ��ִ��˳���ж��Ƿ�ѡ��
                    If Val(.TextMatrix(i, col_ִ��˳��)) <= 0 Then
                        Call vsTransfusion_KeyDown(vbKeySpace, 0)
                    End If
                Else
                    If Val(.TextMatrix(i, col_ִ��˳��)) > 0 Then
                        Call vsTransfusion_KeyDown(vbKeySpace, 0)
                    End If
                End If
            End If
        Next
        .Redraw = True
        .Row = intRow
    End With
    
    vsTransfusion.SetFocus
End Sub

Private Sub cmdOk_Click()
    Dim strStat As String      '׼����Ϊ����״̬
    Dim blnNoCall As Boolean
    Dim strSeqNo As String, strErr As String
    Dim blnUpdate As Boolean
    
    '������λ
    If cboSeating.List(cboSeating.ListIndex) <> "<����>" And (mPatient.��λ�� = "" Or mPatient.��λ�� = "��") Then
        If mSeatings.SetSeating(mPatient.����ID, mPatient.�Һŵ�, cboSeating.List(cboSeating.ListIndex)) Then
            cboSeating.RemoveItem cboSeating.ListIndex '�����ռ����λ
        Else
            Exit Sub
        End If
    End If
    
    '���ŶӼ�¼
'    If txtBase(Base.˳���).Tag <> "" Then
'        '��Ҫд��
'        Call mPatient.AddQueue(mlng����ID)
'    End If

    On Error GoTo errHandle
    
    '��дִ�м�¼
    If Not mGps���� Is Nothing Then
        If mGps����.ѡ������ > 0 Then
            
            mGps����.pִ��ժҪ = txtժҪ(0)
            mGps����.p����ִ��ʱ�� = dtpForecastTime(0).Value
            mGps����.p�ӵ��� = cboOperator(�ӵ���)
            mGps����.SelectGroupThingNew mlng����ID, chkPrint(0).Value = 1, 0, Me, chkView.Value = 1
            strStat = "7-ִ����" '������ĸ�Ϊ7��ִ����
        End If
    End If
    
    If Not mGpsƤ�� Is Nothing Then
        If mGpsƤ��.ѡ������ > 0 Then
            mGpsƤ��.pִ��ժҪ = txtժҪ(3)
            mGpsƤ��.p����ִ��ʱ�� = dtpForecastTime(3).Value
            mGpsƤ��.p�ӵ��� = cboOperator(�ӵ���)
            mGpsƤ��.p��ʱ = txtTransfusion(Ԥ��ʱ��)
            If chkƤ������.Value = 1 Then
                If mGpsƤ��.p��ʱ <= 0 Then mGpsƤ��.p��ʱ = 5
                mGpsƤ��.p���� = Val(zldatabase.GetPara("Ƥ��������ǰʱ��", glngSys, 1264))
                If mGpsƤ��.p���� < 0 Or mGpsƤ��.p���� > 60 Then mGpsƤ��.p���� = 0
                mGpsƤ��.p���� = mGpsƤ��.p����
            Else
                mGpsƤ��.p���� = -1
            End If
            mGpsƤ��.SelectGroupThingNew mlng����ID, chkPrint(3).Value = 1, 3, Me, chkView.Value = 1
            strStat = "7-ִ����" 'Ƥ����ĸ�Ϊ7��ִ����
        End If
    End If
    
    
    If Not mGpsע�� Is Nothing Then
        If mGpsע��.ѡ������ > 0 Then
            mGpsע��.pִ��ժҪ = txtժҪ(2)
            mGpsע��.p����ִ��ʱ�� = dtpForecastTime(2).Value
            mGpsע��.p�ӵ��� = cboOperator(�ӵ���)
            mGpsע��.SelectGroupThingNew mlng����ID, chkPrint(2).Value = 1, 2, Me, chkView.Value = 1
            
            strStat = "7-ִ����" 'ע����ĸ�Ϊ7��ִ����
        End If
    End If
    
    If Not mGps��Һ Is Nothing Then
        If mGps��Һ.ѡ������ > 0 Then
            
            mGps��Һ.pִ��ժҪ = txtժҪ(1)
            mGps��Һ.p����ִ��ʱ�� = dtpForecastTime(1).Value
            mGps��Һ.p��ϵ�� = txtTransfusion(��ϵ��)
            
            If Not mblnLiquid Then
                If mblnImmediatePuncture Then
                    'ֱ�ӽ��봩��״̬
                    strStat = "7-ִ����"
                Else
                    blnNoCall = CurDayNoCall(mlng����ID, mPatients, mPatient)
                    If Not blnNoCall Then
                        strStat = "1-����Һ"  '��Һ��ĸ�Ϊ 5��������
                    Else
                        strStat = "7-ִ����"  '��Һ��ĸ�Ϊ 7��ִ����
                    End If
                End If
                mGps��Һ.p��ҩ�� = cboOperator(��Һ��ҩ��)
            Else
                strStat = "1-����Һ"  '����Һ���̣���1-����Һ  ������ҩ��
            End If
            
            mGps��Һ.p�ӵ��� = cboOperator(�ӵ���)
            mGps��Һ.p��ʱ = txtTransfusion(Ԥ��ʱ��)
            If chk��Һ����.Value = 1 Then
                If mGps��Һ.p��ʱ <= 0 Then mGps��Һ.p��ʱ = 5
                mGps��Һ.p���� = Val(zldatabase.GetPara("��Һ������ǰʱ��", glngSys, 1264))
                If mGps��Һ.p���� < 0 Or mGps��Һ.p���� > 60 Then mGps��Һ.p���� = 3
                mGps��Һ.p���� = mGps��Һ.p����
            Else
                mGps��Һ.p���� = -1
            End If
            mGps��Һ.SelectGroupThingNew mlng����ID, chkPrint(1).Value = 1, 1, Me, chkView.Value = 1, chkLabel.Value = 1, chkWristband.Value = 1
    
        End If
    End If
    
    On Error GoTo 0

    '����ӹ���Һ�ĵ����Ͳ��� ״̬��   2014-08-21 ȡ��
    'ӦΪ�ӹ�������û��ִ�н����Ĳ��ˣ�������״̬��ִ�н����Ĳ���������״̬������

    blnUpdate = Not CurDayHaveItem(mPatient, mlng����ID)
    
    'blnUpdate=True����ǰû�нӹ������Ŷ�״̬=4��������=3���˺ţ�=2������
    If blnUpdate Or Val(mPatient.�Ŷ�״̬) = 4 Or Val(mPatient.�Ŷ�״̬) = 3 Or Val(mPatient.�Ŷ�״̬) = 2 Then
        mPatient.UpdateState strStat, mlng����ID, False
        
        SaveOperLog mlng����ID, mPatient, QUEUE, "�ӵ�����Ķ���״̬Ϊ" & strStat
        
        '�к������̵ģ����ޡ���Һ�����̾����䴩��̨
        If strStat = "1-����Һ" Then
            Call AllocationDesks(mlng����ID, mPatient, strSeqNo, strErr)
        End If
        If Not mblnLiquid And strStat = "1-����Һ" Then
            '�ޡ���Һ������ʱ�Զ�ִ����Һ����������״̬
            strStat = Liquid(mlng����ID, mPatient.Key, mPatients, strErr)
            If Trim(strStat) = "" Then
                strStat = "5-������"
            End If
            mPatient.UpdateState strStat, mlng����ID, False
            SaveOperLog mlng����ID, mPatient, QUEUE, "�ӵ�����Ķ���״̬Ϊ" & strStat
            If strStat = "5-������" Then
                Call QueueCall("��Һ��", mlng����ID, mPatient)
            End If
        End If
    End If
   
    If mbytType = Val("1-�Զ��ӵ�") Then
        Unload Me
    Else
        'ˢ������
        Call initObject
        mblnOk = True
    End If
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        Call SaveErrLog
    End If
End Sub

Private Sub cmdReadCard_Click()
    If Not mobjICCard Is Nothing Then
        txtNo.Text = mobjICCard.Read_Card(Me)
        If txtNo.Text <> "" Then Call cmdRefresh_Click
    End If
End Sub

Private Sub cmdRefresh_Click()
    Dim strFindType As String, strFindTxt As String
    Dim strFindNo As String, blnFind As Boolean
    Dim objPoint As POINTAPI
    Dim sglX As Single, sglY As Single, iRow As Integer
    Dim rsTmp As ADODB.Recordset
    Dim rsVariable As ADODB.Recordset, strSQL As String
    Dim objSeating As Seating, i As Integer
    Dim strInfo As String, strTmp As String
    Dim intFrom As Integer
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    If Trim(txtNo.Text) = "" Then
        MsgBox "����д��" & idkSelect.GetCurCard.���� & "����", vbInformation, gstrSysName
        txtNo.SetFocus
        Exit Sub
    End If
    
    On Error GoTo hErr
    
    '���¸�ҳ��ġ�Ԥ�ƿ�ʼʱ�䡱
    For i = dtpForecastTime.LBound To dtpForecastTime.UBound
        dtpForecastTime(i).Value = zldatabase.Currentdate
    Next
    
    'If mDateBegin <> dtpBegin.Value Or mdateEnd <> DtpEnd.Value Then
    Me.cmdRefresh.Enabled = False
    Set mPatient = Nothing
    Call initObject '��ʼ�����ݼ�
    Call RefPatiData
    cmdOk.Enabled = False
    
    '--��ʼ��ѡ����Ҫ�õļ�¼��
'    SaveLog "1-��ʼ��ѡ����Ҫ�õļ�¼��"
'    Set rsTmp = New ADODB.Recordset
'    With rsTmp
'        .Fields.Append "ID", adVarChar, 20
'        .Fields.Append "Key", adVarChar, 40
'        .Fields.Append "�Һŵ�", adVarChar, 20
'        .Fields.Append "�Һ�ʱ��", adVarChar, 20
'        .Fields.Append "����ʱ��", adVarChar, 20
'        .Fields.Append "���˿���", adVarChar, 100
'        .CursorLocation = adUseClient
'        .LockType = adLockOptimistic
'        .CursorType = adOpenStatic
'        .Open
'    End With

    mDateBegin = dtpBegin.Value
    mdateEnd = dtpEnd.Value
    
    '��ȡ����
    LogWrite "��Һ�ӵ��ĵ�����־", "" & glngModul, "cmdRefresh_Click", "2-��ȡ����"
    strFindType = idkSelect.GetCurCard.����
    If strFindType = "�Һŵ�" Then txtNo.Text = GetFullNO(txtNo.Text, 12)   '12���Һ��վݺ�
    strFindTxt = Trim(txtNo)
    strFindNo = ""
    
'    If strFindType = "�Һŵ���" Then
'        '���Һŵ����ң�����ȷ��ʱ��,������ʽ��Ҫ�ֹ�ָ��ʱ���
'        strSQL = "Select ִ��ʱ�� From ���˹Һż�¼ Where No=[1]"
'        Set rsVariable = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFindTxt)
'        If Not rsVariable.EOF Then
'            mDateBegin = Format(rsVariable!ִ��ʱ��, "yyyy-MM-dd 00:00:00")
'            mdateEnd = Format(mDateBegin, "yyyy-MM-dd 23:59:59")
'        End If
'    End If
    
'    If InStr(";���￨;�����;���ݺ�;����;���֤;IC��;�ɣÿ�;", ";" & Replace(strFindType, "��", "") & ";") = 0 Then

    'strInfo��ʽ�������(1-6�̶�����7����Ϊһ��ͨ��)|����|һ��ͨ���ID
    '�����
    Select Case strFindType
        Case "���￨"
            strInfo = "1"
        Case "�����"
            strInfo = "2"
        Case "�Һŵ�", ""
            strInfo = "3"
        Case "����"
            strInfo = "4"
        Case "���֤��", "�������֤"
            strInfo = "5"
        Case Else
            'һ��ͨ
            strInfo = "6"
    End Select
    '����
    strInfo = strInfo & "|" & strFindTxt
    'һ��ͨ���ID
    If Val(strInfo) >= 6 Or Val(strInfo) = 1 Then
        strTmp = GetSquareCardInfo(mstrSquareCards, strFindType, enuCardProperty.�����ID)
        strInfo = strInfo & "|" & strTmp
    Else
        strInfo = strInfo & "|"
    End If
    
    LogWrite "��Һ�ӵ��ĵ�����־", "" & glngModul, "cmdRefresh_Click", "3-��ȡ���ݿ�ʼ"
    Call mPatients.FetchPatients(mlng����ID, mDateBegin, mdateEnd, True, strInfo, , , mobjSquareCard)
    LogWrite "��Һ�ӵ��ĵ�����־", "" & glngModul, "cmdRefresh_Click", "4-��ȡ���ݽ���"
    
    For Each mPatient In mPatients
        blnFind = False
        If strFindType = "�Һŵ�" Then
            If mPatient.�Һŵ� = strFindTxt Then blnFind = True
        ElseIf strFindType = "�����" Then
            If mPatient.����� = strFindTxt Then blnFind = True
        ElseIf strFindType = "����" Then
            If mPatient.���� = strFindTxt Then blnFind = True
        ElseIf strFindType = "���֤��" Or strFindType = "�������֤" Then
            If mPatient.���֤�� = strFindTxt Then blnFind = True
        Else
            'һ��ͨ
            blnFind = True
        End If
        
        If blnFind Then
            LogWrite "��Һ�ӵ��ĵ�����־", "" & glngModul, "cmdRefresh_Click", "5-�ҵ���Ӧ�ĹҺŵ�"
            '��ȡδִ����ɣ�����δ��ִ����ֹʱ�ڵĹҺŵ�
            If ExecutionComplete(mlng����ID, mPatient) Then
                strFindNo = strFindNo & "," & mPatient.Key
'                rsTmp.AddNew
'                iRow = iRow + 1
'                rsTmp.Fields("ID").Value = iRow
'                rsTmp.Fields("Key").Value = mPatient.Key
'                rsTmp.Fields("�Һŵ�").Value = mPatient.�Һŵ�
'                rsTmp.Fields("�Һ�ʱ��").Value = Format(mPatient.�Һ�ʱ��, "yyyy-MM-dd hh:mm")
'                rsTmp.Fields("����ʱ��").Value = Format(mPatient.�Һ�ʱ��, "yyyy-MM-dd hh:mm")
'                rsTmp.Fields("���˿���").Value = GetClinicDept(mPatient.����ID, mPatient.Key)
'                rsTmp.Update
            End If
            intFrom = mPatient.������Դ
        End If
    Next
    
    If strFindNo = "" Then
        LogWrite "��Һ�ӵ��ĵ�����־", "" & glngModul, "cmdRefresh_Click", "5-δ�ҵ���Ӧ�ĹҺŵ�"
        Set mPatient = Nothing
        MsgBox "�ڱ�����δ�ҵ����ˣ�", vbInformation, gstrSysName
        GoTo hExit
    Else
        strFindNo = Mid(strFindNo, 2)
        
        '���ﲡ�ˣ��������۲��˲����ڶ��ŹҺŵ�
        If InStr(strFindNo, ",") > 0 And intFrom <> 1 Then
'            '���������ϵļ�¼,Ҫѡ��һ���Һŵ�
'            SaveLog "6-���������ϵļ�¼��Ҫѡ��һ���Һŵ�"
'            Call ClientToScreen(txtNo.hwnd, objPoint)
'            sglX = objPoint.X * 15 - 30
'            sglY = objPoint.Y * 15 + 300
'
'            SaveLog "7-��ʾѡ����"
'            If intFrom = 1 Then
'                strTmp = "id,0,0,0;KEY,0,0,0;����ʱ��,1500,0,0;���˿���,2200,0,0"
'            Else
'                strTmp = "id,0,0,0;KEY,0,0,0;�Һŵ�,1200,0,0;�Һ�ʱ��,1500,0,0;���˿���,2200,0,0"
'            End If
'            If frmSelect.ShowSelect(Me, rsTmp, strTmp, sglX, sglY, 5000, 3000, Me.Name & "\ѡ��", "��ѡ") Then
'                strFindNo = Trim("" & rsTmp!Key)
'                SaveLog "8-��ʾѡ����" & strFindNo
'            Else
'                SaveLog "8-δѡ��"
'                GoTo hExit
'            End If

            '���������ϵļ�¼,Ҫѡ��һ���Һŵ�
            vRect = zlControl.GetControlRect(txtNo.hwnd)
            
            '������Null�����ѡ�������ֵ�һ��ΪNull���л���������ʾ
            strSQL = "Select a.Id, Null �ϼ�id, 0 ĩ��, a.No ����, '' ��������, '' ����ҽ��, '' �÷�, '' ҽ������, 0 ����, " & _
                     "    '' ��λ, '' ִ��Ƶ��, 0 ҽ��id, 0 ʣ������ " & vbNewLine & _
                     "From ���˹Һż�¼ A, Table(f_Str2list([1], ',')) B " & vbNewLine & _
                     "Where a.No = b.Column_Value " & vbNewLine & _
                     "Union All " & vbNewLine & _
                     "Select a.*, b.ʣ������ " & vbNewLine & _
                     "From (Select Rownum * -1 ID, c.Id �ϼ�id, 1 ĩ��, a.�Һŵ�, e.���� ��������, a.����ҽ��, b.ҽ������ �÷�, " & _
                     "          a.ҽ������, a.��������, d.���㵥λ, a.ִ��Ƶ��, a.���id " & vbNewLine & _
                     "      From ����ҽ����¼ A, ����ҽ����¼ B, ���˹Һż�¼ C, ������ĿĿ¼ D, ���ű� E, Table(f_Str2list([1], ',')) F " & vbNewLine & _
                     "      Where a.���id = b.Id And a.�Һŵ� = c.No And a.������Ŀid = d.Id And a.��������id = e.Id " & vbNewLine & _
                     "          And a.�Һŵ� = f.Column_Value And a.������� In ('5', '6', '7')) A," & vbNewLine & _
                     "     (Select a.Id, Nvl(Avg(b.��������), 0) - Nvl(Sum(c.��������), 0) ʣ������ " & vbNewLine & _
                     "      From ����ҽ����¼ A, ����ҽ������ B, ����ҽ��ִ�� C, Table(f_Str2list([1], ',')) D " & vbNewLine & _
                     "      Where a.�Һŵ� = d.Column_Value And a.Id = b.ҽ��id And b.ҽ��id = c.ҽ��id(+) And b.���ͺ� = c.���ͺ�(+) " & _
                     "          And a.������� = 'E' " & vbNewLine & _
                     "      Group By ID) B " & vbNewLine & _
                     "Where a.���id = b.Id "
            Set rsTmp = zldatabase.ShowSQLSelect(Me, strSQL, 2, "ѡ��Һŵ�", False, False, "", False, False, True, _
                                                    vRect.Left, vRect.Bottom, 0, blnCancel, False, False, _
                                                    strFindNo)
            If blnCancel Then
                LogWrite "��Һ�ӵ��ĵ�����־", "" & glngModul, "cmdRefresh_Click", "ȡ��ѡ��Һŵ�"
                GoTo hExit
            End If
            If rsTmp.EOF Then
                LogWrite "��Һ�ӵ��ĵ�����־", "" & glngModul, "cmdRefresh_Click", "δ��ѯ���Һŵ�����"
                GoTo hExit
            End If
            
            strFindNo = rsTmp!����
                        
        End If
    End If
    
    Set mPatient = mPatients(strFindNo)
    
    '-- ��ʼ����λ
    LogWrite "��Һ�ӵ��ĵ�����־", "" & glngModul, "cmdRefresh_Click", "9-��ʼ����λ"
    cboSeating.Clear
    If mPatient.��λ�� = "" Or mPatient.��λ�� = "��" Then
        cboSeating.AddItem "<����>"
        
        For Each objSeating In mSeatings
            If objSeating.����ID = 0 Then
                cboSeating.AddItem objSeating.��� & "-" & objSeating.���
            End If
        Next
        mstr��λ = Replace(mstr��λ, "_", "-")
        If cboSeating.ListCount > 0 Then
            For i = 0 To cboSeating.ListCount - 1
                If mstr��λ = "" Then Exit For
                If cboSeating.List(i) = mstr��λ Then
                    Exit For
                End If
            Next
            If i < cboSeating.ListCount Then
                cboSeating.ListIndex = i
            Else
                cboSeating.ListIndex = 0
            End If
        End If
    Else
        cboSeating.AddItem mPatient.��λ��
        cboSeating.ListIndex = 0
        cboSeating.Enabled = False
    End If
    If cboSeating.Enabled Then
        cboSeating.Enabled = InStr(";" & gstrPrivs & ";", ";" & "��λ����" & ";") > 0
    End If
    LogWrite "��Һ�ӵ��ĵ�����־", "" & glngModul, "cmdRefresh_Click", "10-ˢ����ʾ"
    Call InceptBill 'ˢ����ʾ
    
hExit:
    LogWrite "��Һ�ӵ��ĵ�����־", "" & glngModul, "cmdRefresh_Click", "9-�˳�"
    Me.cmdRefresh.Enabled = True
    Exit Sub
    
hErr:
    LogWrite "��Һ�ӵ��ĵ�����־", "" & glngModul, "cmdRefresh_Click", "����ˢ�£���" & CStr(Erl()) & "�У�" & Err.Description
    Exit Sub
    
errSQL:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    
    If mblnActivate Then Exit Sub
    
    If mblnLoad Then
        If mbytType = Val("1-�Զ��ӵ�") Then
            idkSelect.IDKind = mptiInfo.IDKindIDX
            txtNo.Text = mptiInfo.Text
        Else
            For i = 1 To idkSelect.ListCount
                If idkSelect.Cards(i).���� = mstrKeyType Then
                    idkSelect.IDKind = i
                    Exit For
                End If
            Next
            txtNo.Text = marrKey(0)
        End If
        
        mblnLoad = False
        If txtNo <> "" Then cmdRefresh_Click
    End If
    
    mblnActivate = True
    
    Exit Sub
    
errHandle:
    mblnActivate = True
    Call ErrCenter
End Sub


Private Sub RefPatiData()
    '��ʾҪ�ӵ����˵���Ϣ
    
    If mPatient Is Nothing Then
        txtBase(Base.�Һŵ���) = ""
        txtBase(Base.����ʱ��) = ""
        txtBase(Base.����) = ""
        txtBase(Base.�Ա�) = ""
        txtBase(Base.����) = ""
        txtBase(Base.ҽ��) = ""
        txtBase(Base.���˿���) = ""
        txtBase(Base.���) = ""
        
        '--  ��ʼ��˳���
        txtBase(Base.˳���).Tag = ""
        txtBase(Base.˳���) = ""
        Call GroupToVsFlex(1)
        Exit Sub
    Else
        txtBase(Base.�Һŵ���) = mPatient.�Һŵ�
        txtBase(Base.����ʱ��) = mPatient.�Һ�ʱ��
        txtBase(Base.����) = mPatient.����
        txtBase(Base.�Ա�) = mPatient.�Ա�
        txtBase(Base.����) = mPatient.����
        txtBase(Base.ҽ��) = mPatient.ҽ��
        txtBase(Base.���˿���) = mPatient.���˿���
        txtBase(Base.���) = mPatient.�������
        
        '--  ��ʼ��˳���
        If mPatient.˳��� = "0" Then
            txtBase(Base.˳���).Tag = mPatient.Get˳���
            txtBase(Base.˳���) = txtBase(Base.˳���).Tag
        Else
            txtBase(Base.˳���) = mPatient.˳���
        End If
    End If
    
    If mGps����.Count <= 0 Then
        stabType.TabVisible(0) = False
        chkPrint(0).Value = 0
        chkPrint(0).Visible = False
    Else
        stabType.TabVisible(0) = True
        chkPrint(0).Visible = True
        stabType.Tab = 0
    End If

    If mGps��Һ.Count <= 0 Then
        stabType.TabVisible(1) = False
        chkPrint(1).Value = 0
        chkPrint(1).Visible = False
        chkLabel.Visible = False
        chkWristband.Visible = False
    Else
        stabType.TabVisible(1) = True
        chkPrint(1).Value = 1
        chkPrint(1).Visible = True
        chkLabel.Visible = True
        chkWristband.Visible = True
        stabType.Tab = 1
    End If

    If mGpsע��.Count <= 0 Then
        stabType.TabVisible(2) = False
        chkPrint(2).Value = 0
        chkPrint(2).Visible = False
    Else
        stabType.TabVisible(2) = True
        chkPrint(2).Visible = True
        stabType.Tab = 2
    End If

    If mGpsƤ��.Count <= 0 Then
        stabType.TabVisible(3) = False
        chkPrint(3).Value = 0
        chkPrint(3).Visible = False
    Else
        stabType.TabVisible(3) = True
        chkPrint(3).Visible = True
        stabType.Tab = 3
    End If

    If mblnHaveData Then
        
        Call stabType_Click(-1)
    End If
    
End Sub
Private Sub Form_Load()
    Dim curDate As Date, i As Integer, ObjOutNurse As OutNurse, Y As Integer
    Dim strPara As String
    
    mstrPrivs = gstrPrivs
    
    '����Ϊ��24Сʱ�ڵ�ҽ��
    dtpBegin = mDateBegin
    dtpEnd = mdateEnd
    
    mstrKeyType = Trim(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�ӵ���ѯ����", ""))
    mintLabelState = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ӡ��Һƿǩ", "0"))
    mintWristband = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ӡ��Һ���", "0"))
    
    'Ԥ�ƿ�ʼʱ��
    curDate = zldatabase.Currentdate
    For i = 0 To dtpForecastTime.Count - 1
        dtpForecastTime(i).Value = curDate
    Next


    '--- ��ʼ����ʿ�б�
    For i = 0 To cboOperator.Count - 1
        cboOperator(i).Clear
    Next
    
    For Each ObjOutNurse In mOutNurses
        For i = 0 To cboOperator.Count - 1
            cboOperator(i).AddItem ObjOutNurse.����
        Next
    Next
    For Y = 0 To cboOperator.Count - 1
        If cboOperator(Y).ListCount > 0 Then
            For i = 0 To cboOperator(Y).ListCount - 1
                If cboOperator(Y).List(i) = UserInfo.���� Then
                    cboOperator(Y).ListIndex = i
                    Exit For
                Else
                    cboOperator(Y).ListIndex = 0
                End If
            Next
            
        End If
    Next
    
    '�ӵ���ֱ�Ӵ���
    mblnImmediatePuncture = zldatabase.GetPara("�ӵ�ֱ�Ӵ���", glngSys, glngModul, "0") = "1"
            
    '85046
    'mblnLiquid = GetDeptInListPara("������Һ_��Һ�����б�", mlng����ID)
    strPara = zldatabase.GetPara("����Һ�����б�", glngSys, 1264, "")
    mblnLiquid = InStr("," & strPara & ",", "," & mlng����ID & ",") > 0
    
    If mblnLiquid Then
        Me.cboOperator(3).Visible = False
        Me.lblOperator(3).Visible = False
    Else
        Me.cboOperator(3).Visible = True
        Me.lblOperator(3).Visible = True
    End If
    Me.Caption = "�ӵ� " & mstrִ�п���
    '��ò���,�ÿ���,ָ��ʱ���ڵ� ҽ�����͵�
    
'    mnuGNo.Checked = True
'    mnuMNo.Checked = False
'    mnuName.Checked = False
'    mnuSFZ.Checked = False
'    mnuICK.Checked = False
'    mnuJZK.Checked = False
    
    cmdOk.Enabled = False
    '����/��ʼ��һ��ͨ����
    Err = 0: On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Not mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle) Then
        Set mobjSquareCard = Nothing
        MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
    Else
        mstrSquareCards = mobjSquareCard.zlGetIDKindStr(mstrSquareCards)
    End If

'    '��������
'    Set mobjIDCard = New clsIDCard
'    On Error Resume Next
'    Set mobjICCard = CreateObject("zlICCard.clsICCard")
'    On Error GoTo 0
'
'    '��ʼ�������˵���������mstrSquareCards����
'    If mstrSquareCards <> "" Then
'        Dim arrVal As Variant
'        Dim strName As String
'        Dim objMenuPopItem As Object
'        Dim blnAdd As Boolean
'
'        arrVal = Split(mstrSquareCards, ";")
'        For i = LBound(arrVal) To UBound(arrVal)
'            strName = Split(arrVal(i), "|")(enuCardProperty.ȫ��)
'            If InStr(";���￨;�����;���ݺ�;����;���֤;IC��;�ɣÿ�;", ";" & strName & ";") = 0 Then
'                If blnAdd = False Then
'                    '��һ���˵���
'                    blnAdd = True
'                Else
'                    Load mnuCardSquare(mnuCardSquare.UBound + 1)
'                End If
'                '���ò˵���
'                With mnuCardSquare(mnuCardSquare.UBound)
'                    .Caption = strName
'                    .Tag = arrVal(i)
'                    .Visible = True
'                End With
'            End If
'        Next
'    End If
    
    '��������IDKindNew�ؼ�
    idkSelect.zlInit Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquareCard, MSTR_MODE, txtNo
    idkSelect.IDKind = 1
    For i = 1 To idkSelect.ListCount
        If idkSelect.Cards(i).���� = mstrKeyType Then
            idkSelect.IDKind = i
            Exit For
        End If
    Next
    
    chkLabel.Value = IIf(mintLabelState = 1, 1, 0)
    chkWristband.Value = IIf(mintWristband = 1, 1, 0)
    
    mblnLoad = True
End Sub

Public Function ShowIncepBill(ByVal bytType As Byte, ByVal lng����ID As Long, ByVal strִ�в��� As String, ByVal str��λ As String, _
                              ByVal DateBegin As Date, ByVal DateEnd As Date, ByRef objPatients As cPatients, _
                              objOutNurses As OutNurses, frmMain As Form, _
                              Optional strNO As String, Optional strJZK As String, _
                              Optional strName As String, Optional ByVal ptiVar As PatiIdentify) As Boolean
'���ܣ��ӵ����ܽӿ�
'������
'  bytType��0-����ӵ���ť��ʽ��1-�Զ����ýӵ������ӵ����ˣ�

    mbytType = bytType
    Set mptiInfo = ptiVar
    
    '�µĽӵ����
    Set mPatients = objPatients
    Set mSeatings = objPatients.mSeatings
    Set mOutNurses = objOutNurses
    
    mlng����ID = lng����ID
    mstrִ�п��� = strִ�в���
    mstr��λ = str��λ
    If DateDiff("d", DateBegin, DateEnd) > 7 Then
        '����7�죬ǿ��10��������������Ϊ���������ʱ�䣬�������ܳ�30�졣
        mDateBegin = Format(DateEnd - 9, "yyyy-MM-dd 00:00:00")
    Else
        mDateBegin = DateBegin
    End If
    mdateEnd = Format(DateEnd, "yyyy-MM-dd 23:59:59")
    mblnOk = False
    
    ReDim marrKey(3)
    marrKey(0) = strNO          '������Ϣ�ı�
    marrKey(1) = strJZK         '���￨��
    marrKey(2) = strName        '��������
    
    Me.Show vbModal, frmMain
    
    ShowIncepBill = mblnOk
    
End Function

Public Function InceptBill() As Boolean

    '�ӵ�������������
    Dim strPar As String
    Dim dateS As Date, dateE As Date

    mblnHaveData = False
    strPar = zldatabase.GetPara("��ʾ��������", glngSys, 1264, "1,1,1,1")
    
    dateS = dtpBegin.Value
    dateE = dtpEnd.Value
    
    If mGps���� Is Nothing Then Set mGps���� = New Groups
    If Val(Split(strPar, ",")(0)) = 1 Then
        Call mGps����.GetGroups(mPatient.����ID, mlng����ID, 0, dateS, dateE, mPatient.�Һŵ�, mPatient.Key, mPatient.������Դ)
        If mGps����.Count > 0 Then
            mblnHaveData = True
        End If
    End If
    
    If mGps��Һ Is Nothing Then Set mGps��Һ = New Groups
    If Val(Split(strPar, ",")(1)) = 1 Then
        Call mGps��Һ.GetGroups(mPatient.����ID, mlng����ID, 1, dateS, dateE, mPatient.�Һŵ�, mPatient.Key, mPatient.������Դ)
        If mGps��Һ.Count > 0 Then
            mblnHaveData = True
        End If
    End If
    
    If mGpsע�� Is Nothing Then Set mGpsע�� = New Groups
    If Val(Split(strPar, ",")(2)) = 1 Then
        Call mGpsע��.GetGroups(mPatient.����ID, mlng����ID, 2, dateS, dateE, mPatient.�Һŵ�, mPatient.Key, mPatient.������Դ)
        If mGpsע��.Count > 0 Then
            mblnHaveData = True
        End If
    End If
    
    If mGpsƤ�� Is Nothing Then Set mGpsƤ�� = New Groups
    If Val(Split(strPar, ",")(3)) = 1 Then
        Call mGpsƤ��.GetGroups(mPatient.����ID, mlng����ID, 3, dateS, dateE, mPatient.�Һŵ�, mPatient.Key, mPatient.������Դ)
        If mGpsƤ��.Count > 0 Then
            mblnHaveData = True
        End If
    End If
    
    Call RefPatiData    'ˢ�½�����ʾ����
    
    If Not mblnHaveData Then
        MsgBox "û�пɽӵ��ݣ�", vbInformation, gstrSysName
    End If
    
    cmdOk.Enabled = mblnHaveData

End Function

Private Sub initObject()
    Set mGps��Һ = Nothing
    Set mGps���� = Nothing
    Set mGpsע�� = Nothing
    Set mGpsƤ�� = Nothing
    If Not mPatient Is Nothing Then Call GroupToVsFlex(-1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strMode As String

    Call initObject
    
    strMode = idkSelect.GetCurCard.����
    If mbytType = 0 Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�ӵ���ѯ����", strMode
    End If
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ӡ��Һƿǩ", IIf(chkLabel.Value = 1, "1", "0")
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ӡ��Һ���", IIf(chkWristband.Value = 1, "1", "0")
    
    Erase marrKey
    Set mPatient = Nothing
    Set mSeatings = Nothing
    Set mOutNurses = Nothing
    Set mobjSquareCard = Nothing
'    Set mobjIDCard = Nothing
    mblnHaveData = False
    mblnActivate = False
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then PopupMenu popMenu
End Sub

Private Sub idkSelect_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtNo.Enabled And txtNo.Visible Then
        txtNo.Text = ""
        txtNo.SetFocus
    End If
End Sub

Private Sub idkSelect_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtNo.Text = objPatiInfor.����
    mblnReadCard = True
    Call txtNo_KeyPress(0)
End Sub

Private Sub stabType_Click(PreviousTab As Integer)
    Call GroupToVsFlex(stabType.Tab)
End Sub

Private Sub GroupToVsFlex(ByVal intType As Integer)
    Dim ObjGroups As Groups
    Dim objGroup As Group
    Dim objBIll As Bill
    Dim lng��� As Long, lng������ As Long
    Dim strHead As String
    Dim dateS As Date, dateE As Date
    
    If InStr(",10,15,20,", "," & Val(txtTransfusion(��ϵ��).Text) & ",") <= 0 Then
        txtTransfusion(��ϵ��).Text = Val(zldatabase.GetPara("Ĭ�ϵ�ϵ��", glngSys, 1264))
        If InStr(",10,15,20,", "," & Val(txtTransfusion(��ϵ��).Text) & ",") <= 0 Then txtTransfusion(��ϵ��).Text = 20
    End If
    
    strHead = ",280,4;˳��,450,4;�ϴ�,450,4;���,450,4;����,2800,1;����,800,7;��λ,450,4;���,800,7;ִ��Ƶ��,1200,1;�÷�,1000,1;����,450,7;����(ml),500,7;ʱ��(��),500,4;ʣ�����,450,7;��Һ��,700,7;��ע,1200,1;billKey,0,1;GroupKey,0,1;�Ʒ�״̬,0,1;��ϸ�Ʒ�״̬,0,1"
    If Not mPatient Is Nothing Then
        dateS = dtpBegin.Value
        dateE = dtpEnd.Value
        
        Select Case intType
        Case 0
            If mGps���� Is Nothing Then
                Set mGps���� = New Groups
                If mGps����.GetGroups(mPatient.����ID, mlng����ID, 0, dateS, dateE, mPatient.�Һŵ�, mPatient.Key, mPatient.������Դ) = True Then
                    Set ObjGroups = mGps����
                End If
            Else
                Set ObjGroups = mGps����
            End If
            '1 ����� 4 ���� 7 �Ҷ���
            strHead = ",280,4;˳��,0,4;�ϴ�,0,4;���,450,4;����,2800,1;����,800,7;��λ,450,4;���,0,7;ִ��Ƶ��,1200,1;�÷�,1000,1;����,0,7;����(ml),0,7;ʱ��(��),0,4;ʣ�����,450,7;���Ʒ�,700,7;��ע,1200,1;billKey,0,1;GroupKey,0,1;�Ʒ�״̬,0,1;��ϸ�Ʒ�״̬,0,1"
        Case 1
            If mGps��Һ Is Nothing Then
                Set mGps��Һ = New Groups
                mGps��Һ.p��ϵ�� = Val(txtTransfusion(��ϵ��))
                If mGps��Һ.GetGroups(mPatient.����ID, mlng����ID, 1, dateS, dateE, mPatient.�Һŵ�, mPatient.Key, mPatient.������Դ) = True Then
                    Set ObjGroups = mGps��Һ
                End If
            Else
                Set ObjGroups = mGps��Һ
            End If
            '1 ����� 4 ���� 7 �Ҷ���
            strHead = ",280,4;˳��,450,4;�ϴ�,450,4;���,450,4;����,2800,1;����,800,7;��λ,450,4;���,800,7;ִ��Ƶ��,1200,1;�÷�,1000,1;����,450,7;����(ml),500,7;ʱ��(��),500,4;ʣ�����,450,7;��Һ��,700,7;��ע,1200,1;billKey,0,1;GroupKey,0,1;�Ʒ�״̬,0,1;��ϸ�Ʒ�״̬,0,1"
        
        Case 2
            If mGpsע�� Is Nothing Then
                Set mGpsע�� = New Groups
                If mGpsע��.GetGroups(mPatient.����ID, mlng����ID, 2, dateS, dateE, mPatient.�Һŵ�, mPatient.Key, mPatient.������Դ) = True Then
                    Set ObjGroups = mGpsע��
                End If
            Else
                Set ObjGroups = mGpsע��
            End If
            '1 ����� 4 ���� 7 �Ҷ���
            strHead = ",280,4;˳��,0,4;�ϴ�,0,4;���,450,4;����,2800,1;����,800,7;��λ,450,4;���,800,7;ִ��Ƶ��,1200,1;�÷�,1000,1;����,0,7;����(ml),0,7;ʱ��(��),0,4;ʣ�����,450,7;ע���,700,7;��ע,1200,1;billKey,0,1;GroupKey,0,1;�Ʒ�״̬,0,1;��ϸ�Ʒ�״̬,0,1"
        
        Case Else
            If mGpsƤ�� Is Nothing Then
                Set mGpsƤ�� = New Groups
                If mGpsƤ��.GetGroups(mPatient.����ID, mlng����ID, 3, dateS, dateE, mPatient.�Һŵ�, mPatient.Key, mPatient.������Դ) = True Then
                    Set ObjGroups = mGpsƤ��
                End If
            Else
                Set ObjGroups = mGpsƤ��
            End If
            '1 ����� 4 ���� 7 �Ҷ���
            strHead = ",280,4;˳��,0,4;�ϴ�,0,4;���,450,4;����,2800,1;����,0,7;��λ,0,4;���,0,7;ִ��Ƶ��,1200,1;�÷�,1000,1;����,0,7;����(ml),0,7;ʱ��(��),0,4;ʣ�����,450,7;Ƥ�Է�,700,7;��ע,1200,1;billKey,0,1;GroupKey,0,1;�Ʒ�״̬,0,1;��ϸ�Ʒ�״̬,0,1"
        End Select
    End If
    vsTransfusion.Redraw = flexRDNone
    vsTransfusion.Rows = 2
    vsTransfusion.Clear
    Call SetVsFlexGridHead(strHead, vsTransfusion)
    
    
    lng������ = 0
    
    If ObjGroups Is Nothing Then Exit Sub
    For Each objGroup In ObjGroups
        lng��� = 0
        With vsTransfusion
            For Each objBIll In objGroup.BillsItem(objGroup.ִ��ҽ��ID & "_" & objGroup.���ͺ�)
                lng��� = lng��� + 1

                .TextMatrix(.Rows - 1, col_ִ��˳��) = IIf(objGroup.��� = 0, "", objGroup.���)
                .RowData(.Rows - 1) = objGroup.ִ��״̬
                '״̬ 0-δִ��;1-��ȫִ��;2-�ܾ�ִ��;3-����ִ��
                Call ShowPic(.Rows - 1, objGroup.���)   ' objGroup.��� �̶���ʾδѡ�����û�������ѡ��һ�Σ��Ա���֤һ��ͨ�շ�
                .TextMatrix(.Rows - 1, col_�ϴ�˳��) = objGroup.�ϴ����
                
                .TextMatrix(.Rows - 1, col_���) = lng���
                .TextMatrix(.Rows - 1, col_ҽ������) = objBIll.ҽ������
                .TextMatrix(.Rows - 1, col_����) = IIf(Left(CStr(objBIll.����), 1) = ".", "0", "") & CStr(objBIll.����)
                .TextMatrix(.Rows - 1, col_��λ) = objBIll.��λ
                .TextMatrix(.Rows - 1, col_��Ŀ���) = IIf(Format(objBIll.���, "0.00") = 0, "", Format(objBIll.���, "0.00"))
                
                If objBIll.��ϸ�Ʒ�״̬ = -1 Then
                    .TextMatrix(.Rows - 1, col_��Ŀ���) = "���Ʒ�"
                ElseIf objBIll.��ϸ�Ʒ�״̬ = -2 Then
                    If objBIll.��� = 0 Then .TextMatrix(.Rows - 1, col_��Ŀ���) = "�����"
                ElseIf objBIll.��ϸ�Ʒ�״̬ = -3 Then
                    .TextMatrix(.Rows - 1, col_��Ŀ���) = "���˷�"
                End If
                .TextMatrix(.Rows - 1, col_ִ��Ƶ��) = objGroup.ִ��Ƶ�� 'Ҫ�ϲ���Ԫ
                .TextMatrix(.Rows - 1, col_�÷�) = objGroup.�÷� 'Ҫ�ϲ���Ԫ
                .TextMatrix(.Rows - 1, col_����) = objGroup.���� 'Ҫ�ϲ���Ԫ
                .TextMatrix(.Rows - 1, col_����) = objBIll.����
                
                 objBIll.ʱ�� = CacleTransTime(objBIll.����, Val(txtTransfusion(��ϵ��)), ObjGroups.Item(objGroup.ִ��ҽ��ID & "_" & objGroup.���ͺ�).����)
                .TextMatrix(.Rows - 1, col_ʱ��) = objBIll.ʱ��
                .TextMatrix(.Rows - 1, col_ҽ������) = objBIll.ҽ������
                .TextMatrix(.Rows - 1, col_ʣ�����) = objGroup.�������� - objGroup.��ִ������ 'Ҫ�ϲ���Ԫ
                
                If objGroup.�շѽ�� > 0 Then .TextMatrix(.Rows - 1, col_�շѽ��) = Format(objGroup.�շѽ��, "0.00") ''Ҫ�ϲ���Ԫ
                If objGroup.�Ʒ�״̬ = -1 Then
                    .TextMatrix(.Rows - 1, col_�շѽ��) = "���Ʒ�"
                ElseIf objGroup.�Ʒ�״̬ = -2 Then
                    If objGroup.�շѽ�� = 0 Then .TextMatrix(.Rows - 1, col_�շѽ��) = "�����"
                ElseIf objGroup.�Ʒ�״̬ = -3 Then
                     .TextMatrix(.Rows - 1, col_�շѽ��) = "���˷�"
                End If
                .TextMatrix(.Rows - 1, col_BillKey) = objGroup.ִ��ҽ��ID & "_" & objBIll.ҽ��ID
                .TextMatrix(.Rows - 1, col_groupkey) = objGroup.ִ��ҽ��ID & "_" & objGroup.���ͺ�
                
                .TextMatrix(.Rows - 1, col_ִ�мƷ�״̬) = objGroup.�Ʒ�״̬
                .TextMatrix(.Rows - 1, col_��ϸ�Ʒ�״̬) = objBIll.��ϸ�Ʒ�״̬
                lng������ = lng������ + objBIll.����
                .Rows = .Rows + 1
            Next
            
        End With
    Next
    
    If vsTransfusion.Rows > 2 Then
        vsTransfusion.RemoveItem vsTransfusion.Rows - 1
    End If
    
    txtTransfusion(Һ������) = lng������     '�����û��޸�
    txtTransfusion(Һ������).Tag = lng������ '����ԭʼ��������,�����ڻָ�
    
    '�ϼ���ʱ��
    txtTransfusion(Ԥ��ʱ��) = vsTransfusion.Aggregate(flexSTSum, 1, col_ʱ��, vsTransfusion.Rows, col_ʱ��)
    
    '��Ԫ��ϲ�
    With vsTransfusion
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(col_����)
        
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 1, .Cols - 1
        .Editable = flexEDKbdMouse
        .Cell(flexcpBackColor, 1, col_����, .Rows - 1, col_����) = VsModiBackColor
        .Cell(flexcpBackColor, 1, col_ҽ������, .Rows - 1, col_ҽ������) = VsModiBackColor
        .Redraw = True
    End With
    
End Sub

Private Sub txtNo_Change()
    idkSelect.SetAutoReadCard Trim(txtNo.Text) = ""
End Sub

'Private Sub txtNo_Change()
'    If Not mobjIDCard Is Nothing Then
'        mobjIDCard.SetEnabled txtNo.Text = "" And Me.ActiveControl Is txtNo
'    End If
'End Sub

Private Sub txtNo_GotFocus()
    Call zlControl.TxtSelAll(txtNo)
    idkSelect.SetAutoReadCard Trim(txtNo.Text) = ""
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
    '���س�
    Dim strCard As String
    
    On Error GoTo hErr
    
    strCard = idkSelect.Cards(idkSelect.IDKind).����
   
    If mblnReadCard Or KeyAscii = 13 Then
'        If KeyAscii <> 13 Then
'            txtNo.Text = txtNo.Text & Chr(KeyAscii)
'            txtNo.SelStart = Len(txtNo.Text)
'        End If
'        KeyAscii = 0
        Call cmdRefresh_Click
        Call zlControl.TxtSelAll(txtNo)
    Else
        Select Case strCard
            Case "�����"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "�Һŵ�"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (txtNo.Text = "" Or txtNo.SelLength = Len(txtNo.Text)) _
                       And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "���֤��", "�������֤"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case Else
                If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                End If
        End Select
    End If
    mblnReadCard = False
    
    Exit Sub

hErr:
    mblnReadCard = False
    LogWrite "��Һ�ӵ��ĵ�����־", "" & glngModul, "txtNo_KeyPress", "��������򣬵�" & CStr(Erl()) & "�У�" & Err.Description
End Sub

Private Sub txtNo_LostFocus()
    idkSelect.SetAutoReadCard False
End Sub

'Private Sub txtNo_LostFocus()
'    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
'End Sub

Private Sub txtTransfusion_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = ��ϵ�� Then
        If KeyAscii = vbKeyReturn Then
            Call zlcommfun.PressKey(vbKeyTab)
        ElseIf InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtTransfusion_LostFocus(Index As Integer)
    If Index = ��ϵ�� Then
        Call stabType_Click(-1)
    End If
End Sub

Private Sub txtTransfusion_Validate(Index As Integer, Cancel As Boolean)
'    If Index = ��ϵ�� Then
'        If txtTransfusion(Index).Text < 10 Or txtTransfusion(Index).Text > 50 Then
'            txtTransfusion(Index).Text = 20
'        End If
'    End If
End Sub

Private Sub vsTransfusion_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strGroupKey As String, strBillKey As String
    Dim objBIll As Bill
    
    With vsTransfusion
    strGroupKey = .TextMatrix(Row, col_groupkey)
    strBillKey = .TextMatrix(Row, col_BillKey)
    
    Select Case Col
    
    Case col_����
        If stabType.Tab = 1 Then
            If mGps��Һ.Count > 0 Then
                mGps��Һ.Item(strGroupKey).���� = Val(.TextMatrix(Row, Col))
                For Each objBIll In mGps��Һ.Item(strGroupKey).BillsItem(strGroupKey)
                    objBIll.ʱ�� = CacleTransTime(objBIll.����, Val(txtTransfusion(��ϵ��)), mGps��Һ.Item(strGroupKey).����)
                    strBillKey = mGps��Һ.Item(strGroupKey).ִ��ҽ��ID & "_" & objBIll.ҽ��ID
                    mGps��Һ.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).ʱ�� = objBIll.ʱ��
                Next
                Call stabType_Click(-1)
                
            End If
        End If
    Case col_����
        If stabType.Tab = 1 Then
            If mGps��Һ.Count > 0 Then
                mGps��Һ.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).���� = Val(.TextMatrix(Row, Col))
                mGps��Һ.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).ʱ�� = _
                CacleTransTime(mGps��Һ.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).����, Val(txtTransfusion(��ϵ��)), mGps��Һ.Item(strGroupKey).����)
                Call stabType_Click(-1)
            End If
        End If
    Case col_ҽ������
        Select Case stabType.Tab
        Case 0
            If mGps����.Count > 0 Then
                mGps����.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).ҽ������ = .TextMatrix(Row, Col)
            End If
        Case 1
            If mGps��Һ.Count > 0 Then
                mGps��Һ.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).ҽ������ = .TextMatrix(Row, Col)
            End If
        Case 2
            If mGpsע��.Count > 0 Then
                mGpsע��.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).ҽ������ = .TextMatrix(Row, Col)
            End If
        Case Else
            If mGpsƤ��.Count > 0 Then
                mGpsƤ��.Item(strGroupKey).BillsItem(strGroupKey).Item(strBillKey).ҽ������ = .TextMatrix(Row, Col)
            End If
        End Select
    End Select
    End With
End Sub

Private Sub vsTransfusion_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    vsTransfusion.AutoSize 1, col_groupkey
End Sub

Private Sub vsTransfusion_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = col_���� Or Col = col_���� Or Col = col_ҽ������) Then Cancel = True
End Sub

Private Sub vsTransfusion_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)

    Dim LeftCol As Long, RightCol As Long, topRow As Long, BottomRow As Long
    
    If Not MergeRow(Row, topRow, BottomRow) Then Exit Sub '�Ǻϲ���,�˳�
    If topRow = BottomRow Then Exit Sub
    
    LeftCol = col_ִ��˳��: RightCol = col_�ϴ�˳��
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsTransfusion)
    
    LeftCol = col_ִ��Ƶ��: RightCol = col_����
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsTransfusion)
    
    LeftCol = col_ʣ�����: RightCol = col_ҽ������
    Call vfgDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done, LeftCol, RightCol, topRow, BottomRow, vsTransfusion)
    
End Sub

Private Sub vsTransfusion_EnterCell()
    If vsTransfusion.Col = col_���� Or vsTransfusion.Col = col_���� Or vsTransfusion.Col = col_ҽ������ Then
        Call vsTransfusion.CellBorder(vsTransfusion.GridColor, 2, 2, 3, 3, 0, 0)
    End If
End Sub

Private Sub vsTransfusion_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsTransfusion
        If Not (.Col = col_���� Or .Col = col_���� Or .Col = col_ҽ������) Then
            If KeyCode = vbKeySpace And .Row > 0 Then
                Call CheckGroup(.Row, col_ѡ��, 1)
            ElseIf KeyCode = vbKeyDelete Then
                Call CheckGroup(.Row, col_ѡ��, 2)
            End If
        End If
    End With
End Sub

Private Sub vsTransfusion_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim NextCol As Long
    If KeyAscii = vbKeyReturn Then
        NextCol = EditCol(Col)
        With vsTransfusion
            If NextCol = -1 Then
                If Row + 1 <= .Rows - 1 Then
                    If .TextMatrix(Row + 1, col_���) = 1 Then
                        .Select Row + 1, col_����
                    Else
                        .Select Row + 1, col_����
                    End If
                Else
                    If .TextMatrix(Row, col_���) = 1 Then
                        .Select Row, col_����
                    Else
                        .Select Row, col_����
                    End If
                End If
            Else
                .Select Row, NextCol
            End If
        End With
    End If
End Sub

Private Function EditCol(ByVal Col As Long) As Long
    '���ص�ǰ��֮��Ŀɱ༭��,-1��ʾ��ǰ��֮���޿ɱ༭��.
    Dim lngCol As Long
    
    If Col + 1 > vsTransfusion.Cols - 1 Then
        EditCol = -1
        Exit Function
    End If
    
    For lngCol = Col + 1 To vsTransfusion.Cols - 1
        If InStr(",8,9,13,", "," & CStr(lngCol) & ",") > 0 Then
            EditCol = lngCol
            Exit Function
        End If
    Next
    If EditCol = 0 Then EditCol = -1
    
End Function

Private Sub vsTransfusion_LeaveCell()
    If vsTransfusion.Col = col_���� Or vsTransfusion.Col = col_���� Or vsTransfusion.Col = col_ҽ������ Then
        On Error Resume Next
        Call vsTransfusion.CellBorder(vsTransfusion.GridColor, 0, 0, 0, 0, 0, 0)
    End If
End Sub

Private Sub vsTransfusion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsTransfusion
        Call CheckGroup(.MouseRow, .MouseCol, Button)
    End With
End Sub

Private Sub CheckGroup(ByVal Row As Long, ByVal Col As Long, ByVal Button As Integer)
    Dim blnCheck As Boolean, StrKey As String, lngRows As Long, lngCurRow As Long, i As Long
    Dim lngCol As Long, lngRow As Long, strTmpKey As String
    Dim lngҽ��ID As Long, lng���ͺ� As Long, intReturn As Integer    'һ��ͨ����ʱʹ��
    Dim blnOK As Boolean
    
    With vsTransfusion
    lngCol = Col: lngRow = Row
    If lngCol = col_ѡ�� And lngRow > 0 Then
        
        StrKey = .TextMatrix(lngRow, col_groupkey)
        If InStr(StrKey, "_") > 0 Then
            If Button = 1 And InStr(",0,3,", "," & .RowData(lngRow) & ",") > 0 Then
                '�ӵ�
                'blnCheck = .Cell(flexcpPicture, lngRow, col_ѡ��) = imgPic.ListImages(2).Picture
                
                blnCheck = Val(.TextMatrix(lngRow, col_ִ��˳��)) = 0
                'һ��ͨ�շѼ��
                Select Case stabType.Tab
                Case 0
                    lngҽ��ID = mGps����.Item(StrKey).ִ��ҽ��ID
                    lng���ͺ� = mGps����.Item(StrKey).���ͺ�

                Case 1
                    lngҽ��ID = mGps��Һ.Item(StrKey).ִ��ҽ��ID
                    lng���ͺ� = mGps��Һ.Item(StrKey).���ͺ�

                Case 2
                    lngҽ��ID = mGpsע��.Item(StrKey).ִ��ҽ��ID
                    lng���ͺ� = mGpsע��.Item(StrKey).���ͺ�

                Case Else
                    lngҽ��ID = mGpsƤ��.Item(StrKey).ִ��ҽ��ID
                    lng���ͺ� = mGpsƤ��.Item(StrKey).���ͺ�
                   
                End Select
                '2012-11-08 ���ʣ���ִ�д���
                If Not CheckRun(lngҽ��ID, lng���ͺ�) Then Exit Sub

                
                If .TextMatrix(lngRow, col_��ϸ�Ʒ�״̬) = -3 Then
                    MsgBox "��ϸ��Ŀ���˷ѣ�����ִ�д˲�����", vbInformation, Me.Caption
                    Exit Sub
                End If
                
                intReturn = OneCardCheck(lngҽ��ID, lng���ͺ�, Me, mobjSquareCard)
                If intReturn = 0 Then
                    '������
                    If InStr(mstrPrivs, "ִ����Ŀδ�շѽӵ�") <= 0 And blnCheck Then
                        If .TextMatrix(lngRow, col_ִ�мƷ�״̬) <> -2 Then '����ã������Ѿ��չ�����2012��07��16
                            If (.TextMatrix(lngRow, col_ִ�мƷ�״̬) > -1) And Val(.TextMatrix(lngRow, col_�շѽ��)) = 0 Then
                                MsgBox .TextMatrix(0, col_�շѽ��) & "δ��ȡ�����շѺ��ٲ�����", vbInformation, Me.Caption
                                Exit Sub
                            End If
                        End If
                    End If
                    If InStr(mstrPrivs, "��ϸ��Ŀδ�շѽӵ�") <= 0 And blnCheck Then
                        If .TextMatrix(lngRow, col_ִ�мƷ�״̬) <> -2 Then '����ã������Ѿ��չ�����2012��07��16
                            If (.TextMatrix(lngRow, col_��ϸ�Ʒ�״̬) > -1) And Val(.TextMatrix(lngRow, col_��Ŀ���)) = 0 Then
                                MsgBox "��ϸ��Ŀδ�շѣ����շѺ��ٲ�����", vbInformation, Me.Caption
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    
                ElseIf intReturn = 2 Then
                    'һ��ͨ������ʧ��,�����ڲ�����ʾ��ֱ���˳�
                    Exit Sub
                End If
                
                
                Select Case stabType.Tab
                Case 0
                    lngRows = mGps����.Item(StrKey).BillsItem(StrKey).Count
                    Call mGps����.CheckGroup(StrKey, blnCheck)
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .TextMatrix(lngCurRow, col_ִ��˳��) = IIf(mGps����.Item(strTmpKey).��� = 0, "", mGps����.Item(strTmpKey).���)
                        If Val(.TextMatrix(lngCurRow, col_ִ��˳��)) > 0 Then
                            If blnOK = False Then blnOK = True
                        End If
                    Next
                Case 1
                    lngRows = mGps��Һ.Item(StrKey).BillsItem(StrKey).Count
                    Call mGps��Һ.CheckGroup(StrKey, blnCheck)
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .TextMatrix(lngCurRow, col_ִ��˳��) = IIf(mGps��Һ.Item(strTmpKey).��� = 0, "", mGps��Һ.Item(strTmpKey).���)
                        If Val(.TextMatrix(lngCurRow, col_ִ��˳��)) > 0 Then
                            If blnOK = False Then blnOK = True
                        End If
                    Next
                Case 2
                    lngRows = mGpsע��.Item(StrKey).BillsItem(StrKey).Count
                    Call mGpsע��.CheckGroup(StrKey, blnCheck)
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .TextMatrix(lngCurRow, col_ִ��˳��) = IIf(mGpsע��.Item(strTmpKey).��� = 0, "", mGpsע��.Item(strTmpKey).���)
                        If Val(.TextMatrix(lngCurRow, col_ִ��˳��)) > 0 Then
                            If blnOK = False Then blnOK = True
                        End If
                    Next
                Case Else
                    lngRows = mGpsƤ��.Item(StrKey).BillsItem(StrKey).Count
                    Call mGpsƤ��.CheckGroup(StrKey, blnCheck)
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .TextMatrix(lngCurRow, col_ִ��˳��) = IIf(mGpsƤ��.Item(strTmpKey).��� = 0, "", mGpsƤ��.Item(strTmpKey).���)
                        If Val(.TextMatrix(lngCurRow, col_ִ��˳��)) > 0 Then
                            If blnOK = False Then blnOK = True
                        End If
                    Next
                End Select
                chkPrint(stabType.Tab).Value = IIf(blnOK, 1, 0)
                
            ElseIf Button = 2 And Val(.TextMatrix(lngRow, col_ִ��˳��)) = 0 Then
                '�Ҽ��ܾ�
                Select Case stabType.Tab
                Case 0
                    If .RowData(lngRow) = 2 Then
                        'ȡ���ܾ�
                        blnOK = mGps����.Item(StrKey).FuncExecRestore
                    Else
                        '�ܾ�
                        blnOK = mGps����.Item(StrKey).FuncExecRefuse
                    End If
                    If blnOK = False Then Exit Sub
                    
                    lngRows = mGps����.Item(StrKey).BillsItem(StrKey).Count
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .RowData(lngCurRow) = mGps����.Item(strTmpKey).ִ��״̬
                    Next
                Case 1
                    If .RowData(lngRow) = 2 Then
                        'ȡ���ܾ�
                        blnOK = mGps��Һ.Item(StrKey).FuncExecRestore
                    Else
                        '�ܾ�
                        blnOK = mGps��Һ.Item(StrKey).FuncExecRefuse
                    End If
                    If blnOK = False Then Exit Sub
                    
                    lngRows = mGps��Һ.Item(StrKey).BillsItem(StrKey).Count
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .RowData(lngCurRow) = mGps��Һ.Item(strTmpKey).ִ��״̬
                    Next
                Case 2
                    If .RowData(lngRow) = 2 Then
                        'ȡ���ܾ�
                        blnOK = mGpsע��.Item(StrKey).FuncExecRestore
                        If blnOK = False Then Exit Sub
                        .RowData(lngRow) = 0
                    Else
                        '�ܾ�
                        blnOK = mGpsע��.Item(StrKey).FuncExecRefuse
                        If blnOK = False Then Exit Sub
                        .RowData(lngRow) = 2
                    End If
                    lngRows = mGpsע��.Item(StrKey).BillsItem(StrKey).Count
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .RowData(lngCurRow) = mGpsע��.Item(strTmpKey).ִ��״̬
                    Next
                Case 3
                    If .RowData(lngRow) = 2 Then
                        'ȡ���ܾ�
                        blnOK = mGpsƤ��.Item(StrKey).FuncExecRestore
                        If blnOK = False Then Exit Sub
                        .RowData(lngRow) = 0
                    Else
                        '�ܾ�
                        blnOK = mGpsƤ��.Item(StrKey).FuncExecRefuse
                        If blnOK = False Then Exit Sub
                        .RowData(lngRow) = 2
                    End If
                    lngRows = mGpsƤ��.Item(StrKey).BillsItem(StrKey).Count
                    For lngCurRow = 1 To .Rows - 1
                        strTmpKey = .TextMatrix(lngCurRow, col_groupkey)
                        .RowData(lngCurRow) = mGpsƤ��.Item(strTmpKey).ִ��״̬
                    Next
                End Select
            End If
            '--- ����ͼƬ
            lngCurRow = Val(.TextMatrix(lngRow, col_���))
            
            If lngRows = lngCurRow Then
                If lngRows = 1 Then
                    Call ShowPic(lngRow, Val(.TextMatrix(lngRow, col_ִ��˳��)))
                    Exit Sub
                Else
                    For i = 1 To lngRows
                       Call ShowPic(lngRow - (lngCurRow - i), Val(.TextMatrix(lngRow - (lngCurRow - i), col_ִ��˳��)))
                    Next
                End If
            Else
                For i = 1 To lngCurRow
                    Call ShowPic(lngRow - (lngCurRow - i), Val(.TextMatrix(lngRow - (lngCurRow - i), col_ִ��˳��)))
                Next
                
                For i = lngCurRow To lngRows
                    Call ShowPic(lngRow + (lngRows - i), Val(.TextMatrix(lngRow + (lngRows - i), col_ִ��˳��)))
                Next
            End If
        End If
    End If
    .Refresh
    End With
End Sub

Private Sub vsTransfusion_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case col_����
        If Val(vsTransfusion.EditText) < 10 Or Val(vsTransfusion.EditText) > 100 Then
            Cancel = True
        End If
    Case col_����
        If Val(vsTransfusion.EditText) < 0 Or Val(vsTransfusion.EditText) > 10000 Then
            Cancel = True
        End If
    End Select
End Sub

Private Function MergeRow(ByVal Row As Long, topRow, BottomRow As Long) As Boolean
    '�Ƿ�ϲ���
    Dim strGroupKey As String, lngRow As Long
    With vsTransfusion
        If .Cols < col_groupkey Then Exit Function
        strGroupKey = .TextMatrix(Row, col_groupkey)
        topRow = Row: BottomRow = Row
        For lngRow = Row To 0 Step -1
            If .TextMatrix(lngRow, col_groupkey) <> strGroupKey Then
                topRow = lngRow + 1
                Exit For
            Else
                topRow = lngRow
            End If
        Next
        
        For lngRow = Row To .Rows - 1
            If .TextMatrix(lngRow, col_groupkey) <> strGroupKey Then
                BottomRow = lngRow - 1
                Exit For
            Else
                BottomRow = lngRow
            End If
        Next
    End With

    If topRow > 0 And BottomRow > 0 Then MergeRow = True
End Function

Private Sub ShowPic(ByVal Row As Long, ByVal ��� As Long)
    '����ָ���е�ͼƬ
    If Row <= 0 Then Exit Sub
    With vsTransfusion
        '״̬ 0-δִ��;1-��ȫִ��;2-�ܾ�ִ��;3-����ִ��
    
        If .RowData(Row) = 0 Or .RowData(Row) = 3 Then
            '0 δѡ�� 3-����ִ�� Ҳ����δִ�д���,��Ϊ�� �ֳɼ�����ִ�е����
            If ��� > 0 Then
                Set .Cell(flexcpPicture, Row, col_ѡ��) = imgPic.ListImages(2).Picture
            Else
                Set .Cell(flexcpPicture, Row, col_ѡ��) = imgPic.ListImages(1).Picture
            End If
        ElseIf .RowData(Row) = 1 Then
            '1 ���
            Set .Cell(flexcpPicture, Row, col_ѡ��) = imgPic.ListImages(4).Picture
        ElseIf .RowData(Row) = 2 Then
            Set .Cell(flexcpPicture, Row, col_ѡ��) = imgPic.ListImages(3).Picture

        End If
    End With
End Sub
Private Function Get�������(ByVal bln����ִ�� As Boolean, ByVal lngҽ��ID As Long, ByVal lng��ID As Long, ByVal str������� As String) As Long
'���ܣ���ȡĳ��ҽ������ĳ��ҽ������������ʵ�ҽ��ִ�д���
'       bln����ִ�� �Ƿ񵥶�ִ�У�����������ڵ��ݵ�ҽ���ĵ���ִ��ĳһ��λ��ĳһ���ּ��
'       lngҽ��ID ����ҽ��ID
'       lng��ID û�и�ҽ�������߸�ҽ��ʱΪҽ��ID,��ҽ��Ϊ���ID
'       str������� ��ҽ�����������
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If bln����ִ�� Then
        lng��ID = lngҽ��ID
        strSQL = "Select -1 * Sum(Nvl(a.����, 1) * a.���� / b.����) As ���������" & vbNewLine & _
                "From ������ü�¼ A, ����ҽ���Ƽ� B" & vbNewLine & _
                "Where a.ҽ����� = [1] And b.ҽ��id = a.ҽ����� And b.�շ�ϸĿid = a.�շ�ϸĿid And Nvl(B.��������,0)=0 And a.��¼״̬ = 2 And mod(a.��¼����,10) in(1,2) And a.�۸񸸺� Is Null And" & vbNewLine & _
                "      a.�շ���� Not In ('5', '6', '7') And Not Exists" & vbNewLine & _
                " (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1)"

    Else
        strSQL = "Select Max(c.������) ���������" & vbNewLine & _
                "From (Select -1 * Sum(Nvl(a.����, 1) * a.���� / b.����) As ������" & vbNewLine & _
                "       From ������ü�¼ A, ����ҽ���Ƽ� B" & vbNewLine & _
                "       Where a.ҽ����� In (Select ID From ����ҽ����¼ Where (ID = [1] Or ���id = [1]) And ������� = [2]) And b.ҽ��id = a.ҽ����� And" & vbNewLine & _
                "             b.�շ�ϸĿid = a.�շ�ϸĿid And Nvl(B.��������,0)=0 And a.��¼״̬ = 2 And mod(a.��¼����,10) in(1,2) And a.�۸񸸺� Is Null And a.�շ���� Not In ('5', '6', '7') And" & vbNewLine & _
                "             Not Exists" & vbNewLine & _
                "        (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) " & vbNewLine & _
                "       Group By  a.ҽ�����,a.�շ�ϸĿid) C"
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ID, str�������)
    If rsTmp.RecordCount <> 0 Then
        Get������� = Val(rsTmp!��������� & "")
    End If
    
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckRun(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long) As Boolean
    '���ҽ���Ƿ񻹿���ִ��
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lng��ִ������ As Long, lng�˷����� As Long
    On Error GoTo errH
    strSQL = "Select " & _
        " Max(ִ��ʱ��) as LastDate," & _
        " Max(Ҫ��ʱ��) as curDate," & _
        " Count(Ҫ��ʱ��) as curCount," & _
        " Sum(��������) as curNum" & _
        " From ����ҽ��ִ��" & _
        " Where ҽ��ID=[1] And ���ͺ�=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
    If Not rsTmp.EOF Then
        lng��ִ������ = Val("" & rsTmp!curNum)
    End If
    
    '���㱾��ִ��Ӧ�õ�Ҫ��ʱ��
    strSQL = "Select A.��������,Nvl(B.���id, B.ID) ��ID,C.���㵥λ,A.�״�ʱ��,A.ĩ��ʱ��,Decode(B.������Դ, 2, Decode(A.��¼����, 1, 1, Decode(A.�������, 1, 1, 2)), 1) ��������," & _
        " B.��ʼִ��ʱ��,B.ִ����ֹʱ��,B.�ϴ�ִ��ʱ��,B.ִ��ʱ�䷽��," & _
        " B.ִ��Ƶ��,B.Ƶ�ʴ���,B.Ƶ�ʼ��,B.�����λ,B.����ID,b.��ҳID,c.���,c.��������,c.ִ�з���" & _
        " From ����ҽ������ A,����ҽ����¼ B,������ĿĿ¼ C" & _
        " Where A.ҽ��ID=B.ID And B.������ĿID=C.ID" & _
        " And A.ҽ��ID=[1] And A.���ͺ�=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
    If rsTmp!�������� = 1 Then
        lng�˷����� = Get�������(False, lngҽ��ID, lngҽ��ID, "" & rsTmp!���)
    
    
        '��ǰʵ���Ѿ�ִ����Ҫ��Ĵ���,��׼��ִ��
        If lng��ִ������ + lng�˷����� >= Val("" & rsTmp!��������) Then
            MsgBox "��ҽ�����η�������ִ�� " & "" & rsTmp!�������� & IIf(lng�˷����� <> 0, " �Σ�" & "��ص����Ѿ��˷ѻ�����" & lng�˷�����, "") & "�Σ���ǰ�Ѿ�ִ���� " & lng��ִ������ & " �Σ������ٽӵ���", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckRun = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExecutionComplete(ByVal lngDeptID As Long, ByVal objPati As cPatient) As Boolean
'���ܣ�����ҽ������Һ��Ƥ�ԡ�ע�䣩ִ�������Ƿ����
'������
'  lngDeptID��ִ�в���ID
'  objPati��������Ϣ��
'���أ�Trueδ�꣬False����
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim StrKey As String
    
    '2014-03-19������ȡ��ִ����ֹʱ�䡱δ���ڵĲ���ҽ����¼
    'ִ����ֹʱ��=NULL���ÿ�ʼִ��ʱ��+7�졣ͬʱ��Ϊ�˽������ҽ������ִ����ֹʱ��Ƚ϶̵��������ִ����ֹʱ��ǿ�Ƽ�1�졣
    On Error GoTo errHandle
    
    If objPati.������Դ = 1 Then
        '��������
        StrKey = objPati.Key
        strSQL = "select c.��������, Sum(d.��������) ִ������" & vbNewLine & _
                "From ������ҳ a, ����ҽ����¼ B, ����ҽ������ C, ����ҽ��ִ�� D, ������ĿĿ¼ E" & vbNewLine & _
                "where a.����id=b.����id and a.��ҳid=b.��ҳid and b.id=c.ҽ��id and c.ҽ��id=d.ҽ��id(+) and c.���ͺ�=d.���ͺ�(+) " & vbNewLine & _
                "   And b.������Ŀid = e.Id And b.������� = 'E' And (e.ִ�з��� In (0, 1, 2, 3) or e.ִ�з��� is null) " & vbNewLine & _
                "   And c.ִ�в���id = [1] And a.��Ժ���� is null " & vbNewLine & _
                "   And b.����id = [2] and b.��ҳid = [3] " & vbNewLine & _
                "Group By c.ҽ��id, c.��������" & vbNewLine & _
                "Having Nvl(c.��������, 0) - Nvl(Sum(d.��������), 0) > 0 "
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "ҽ��ִ������", lngDeptID, Val(StrKey), Val(Split(StrKey, "_")(1)))
    Else
        StrKey = objPati.�Һŵ�
        strSQL = "Select c.��������, Sum(d.��������) ִ������ " & _
                 "From ����ҽ����¼ B, ����ҽ������ C, ����ҽ��ִ�� D, ������ĿĿ¼ E " & _
                 "Where b.Id = c.ҽ��id And c.ҽ��id = d.ҽ��id(+) And c.���ͺ� = d.���ͺ�(+) " & _
                 "  And b.������Ŀid = e.Id And b.������� = 'E' And (e.ִ�з��� In (0, 1, 2, 3) or e.ִ�з��� is null) " & _
                 "  And b.��ʼִ��ʱ�� + 7 > Sysdate " & _
                 "  And c.ִ�в���id = [1] And b.�Һŵ� = [2] " & _
                 "Group By c.ҽ��id, c.�������� " & _
                 "Having Nvl(c.��������, 0) - Nvl(Sum(d.��������), 0) > 0 "
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "ҽ��ִ������", lngDeptID, StrKey)
    End If
    Do While Not rsTemp.EOF
        If zlcommfun.NVL(rsTemp!��������, 0) - zlcommfun.NVL(rsTemp!ִ������, 0) > 0 Then
            ExecutionComplete = True
            Exit Do
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

'Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, _
'    ByVal datBirthDay As Date, ByVal strAddress As String)
''���ܣ����֤ʶ��ɹ��󼤻�
'    mstrIDCard = strID
'
'    If idkSelect.GetCurCard.���� = "���֤��" Then
'        txtNo.Text = mstrIDCard
'    Else
'        txtNo.Text = "" '�������(Ŀǰ�������������²��ܼ���)��
'    End If
'
'    If txtNo.Text <> "" Then Call cmdRefresh_Click
'End Sub
'
'
Private Function GetClinicDept(ByVal lngPatiId As Long, ByVal StrKey As String) As String
'���ܣ���ȡ����ҽ���Ŀ�������
'������
'  lngPatiID������ID
'  strKey���Һŵ��򡰲���id_��ҳid��
'���أ���������

    Dim lngPageID As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If StrKey Like "*_*" Then
        lngPageID = Val(Split(StrKey, "_")(1))
        strSQL = "Select Distinct b.���� " & vbNewLine & _
                 "From ����ҽ����¼ A, ���ű� B " & vbNewLine & _
                 "Where a.��������id = b.Id And a.����id = [1] And ��ҳid = [2] "
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "��ȡ���˵Ŀ�������", lngPatiId, lngPageID)
    Else
        strSQL = "Select Distinct b.���� " & vbNewLine & _
                 "From ����ҽ����¼ A, ���ű� B " & vbNewLine & _
                 "Where a.��������id = b.Id And a.�Һŵ� = [1] And a.����id = [2] "
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "��ȡ���˵Ŀ�������", StrKey, lngPatiId)
    End If
    If rsTemp.EOF = False Then
        GetClinicDept = zlcommfun.NVL(rsTemp!����)
    End If
    rsTemp.Close
    
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume

End Function
