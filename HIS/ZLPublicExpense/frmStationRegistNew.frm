VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmStationRegistNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ��վ�Һ�"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8460
   Icon            =   "frmStationRegistNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdPrice 
      Caption         =   "���滮�۵�(&J)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3720
      TabIndex        =   49
      Top             =   5625
      Width           =   1725
   End
   Begin VB.CheckBox chkAll 
      Height          =   360
      Left            =   8070
      Picture         =   "frmStationRegistNew.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "��ʾ�������"
      Top             =   45
      Width           =   345
   End
   Begin VB.CommandButton cmdNewPati 
      Height          =   345
      Left            =   2940
      Picture         =   "frmStationRegistNew.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "��������(F4)"
      Top             =   600
      Width           =   350
   End
   Begin VB.PictureBox picPayMoney 
      BackColor       =   &H80000005&
      Height          =   435
      Left            =   6300
      ScaleHeight     =   375
      ScaleWidth      =   1995
      TabIndex        =   37
      Top             =   4942
      Width           =   2055
      Begin VB.Label lblPayMoney 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   38
         Top             =   15
         Width           =   720
      End
   End
   Begin VB.PictureBox picInfo 
      Height          =   2925
      Left            =   15
      ScaleHeight     =   2865
      ScaleWidth      =   8310
      TabIndex        =   31
      Top             =   1950
      Width           =   8370
      Begin VB.CommandButton cmdOther 
         Height          =   345
         Left            =   4440
         Picture         =   "frmStationRegistNew.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "����ҽ���ű�"
         Top             =   45
         Width           =   345
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����"
         Height          =   255
         Left            =   7590
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   570
         Width           =   930
      End
      Begin VB.CommandButton cmdReg 
         Height          =   345
         Left            =   4005
         Picture         =   "frmStationRegistNew.frx":1788
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "��ǰҽ���ű�"
         Top             =   45
         Width           =   345
      End
      Begin VB.TextBox txtReg 
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
         Left            =   660
         TabIndex        =   4
         Top             =   45
         Width           =   3360
      End
      Begin VB.CheckBox chkBook 
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7020
         TabIndex        =   9
         Top             =   2543
         Width           =   1485
      End
      Begin VB.ComboBox cboDoctor 
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
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   525
         Width           =   3390
      End
      Begin VB.ComboBox cboAppointStyle 
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
         Left            =   5490
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   525
         Width           =   2055
      End
      Begin VB.ComboBox cboRemark 
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
         Left            =   660
         TabIndex        =   8
         Top             =   2490
         Width           =   6180
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMoney 
         Height          =   1440
         Left            =   75
         TabIndex        =   32
         Top             =   975
         Width           =   8205
         _cx             =   1985886345
         _cy             =   1985874412
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmStationRegistNew.frx":218A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
      Begin VB.Label lblSn 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���:22 "
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   7425
         TabIndex        =   50
         Top             =   90
         Width           =   960
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��"
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
         Left            =   120
         TabIndex        =   40
         Top             =   585
         Width           =   480
      End
      Begin VB.Label lblAppointStyle 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ��ʽ"
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
         Left            =   4365
         TabIndex        =   36
         Top             =   585
         Width           =   960
      End
      Begin VB.Label lblArrangeNO 
         AutoSize        =   -1  'True
         Caption         =   "�ű�"
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
         Left            =   120
         TabIndex        =   35
         Top             =   105
         Width           =   480
      End
      Begin VB.Label lblLimit 
         AutoSize        =   -1  'True
         Caption         =   "�޺�:"
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
         Left            =   4980
         TabIndex        =   34
         Top             =   120
         Width           =   600
      End
      Begin VB.Label lblRemark 
         AutoSize        =   -1  'True
         Caption         =   "��ע"
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
         Left            =   105
         TabIndex        =   33
         Top             =   2550
         Width           =   480
      End
   End
   Begin VB.PictureBox picTotal 
      BackColor       =   &H80000005&
      Height          =   420
      Left            =   795
      ScaleHeight     =   360
      ScaleWidth      =   1575
      TabIndex        =   29
      Top             =   4950
      Width           =   1635
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   825
         TabIndex        =   30
         Top             =   15
         Width           =   720
      End
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   3
      Left            =   -45
      TabIndex        =   25
      Top             =   5430
      Width           =   11000
   End
   Begin VB.ComboBox cboPayMode 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4335
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4935
      Width           =   1950
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   2
      Left            =   -30
      TabIndex        =   20
      Top             =   1440
      Width           =   11000
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   1
      Left            =   -30
      TabIndex        =   19
      Top             =   480
      Width           =   11000
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
      Height          =   435
      Left            =   7080
      TabIndex        =   12
      Top             =   5625
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5730
      TabIndex        =   11
      Top             =   5625
      Width           =   1300
   End
   Begin VB.CommandButton cmdHelp 
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
      Height          =   435
      Left            =   30
      TabIndex        =   13
      Top             =   5625
      Width           =   1300
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Index           =   0
      Left            =   -60
      TabIndex        =   17
      Top             =   960
      Width           =   11000
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   360
      Left            =   705
      TabIndex        =   16
      Top             =   600
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   635
      Appearance      =   2
      IDKindStr       =   "��|��������￨|0|0|0|0|0|;ҽ|ҽ����|0|0|0|0|0|;��|���֤��|1|0|0|0|0|;��|�����|0|0|0|0|0|"
      BorderStyle     =   1
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
      DefaultCardType =   "0"
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
      Height          =   360
      Left            =   1290
      TabIndex        =   1
      Top             =   600
      Width           =   1650
   End
   Begin VB.ComboBox cboNO 
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
      Left            =   6510
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   1575
   End
   Begin VB.CommandButton cmdTime 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4125
      TabIndex        =   28
      Top             =   1568
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   360
      Left            =   3030
      TabIndex        =   3
      Top             =   1560
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   93323266
      CurrentDate     =   42121
   End
   Begin VB.PictureBox picRoom 
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
      Left            =   5745
      ScaleHeight     =   300
      ScaleWidth      =   2580
      TabIndex        =   44
      Top             =   1560
      Width           =   2640
      Begin VB.Label lblRoomName 
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   45
         TabIndex        =   45
         Top             =   15
         Width           =   120
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   360
      Left            =   675
      TabIndex        =   2
      Top             =   1560
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   93323267
      CurrentDate     =   42121
   End
   Begin VB.PictureBox picDept 
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
      Left            =   675
      ScaleHeight     =   300
      ScaleWidth      =   3330
      TabIndex        =   42
      Top             =   1560
      Width           =   3390
      Begin VB.Label lblDeptName 
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   45
         TabIndex        =   43
         Top             =   15
         Width           =   120
      End
   End
   Begin VB.Label lblʱ�� 
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
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   4440
      TabIndex        =   51
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lbl�� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   135
      TabIndex        =   46
      Top             =   45
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblPayMode 
      AutoSize        =   -1  'True
      Caption         =   "֧����ʽ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2970
      TabIndex        =   24
      Top             =   4995
      Width           =   1320
   End
   Begin VB.Label lblSum 
      AutoSize        =   -1  'True
      Caption         =   "�ϼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   23
      Top             =   4995
      Width           =   660
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      Caption         =   "����Ԥ�����:0.00     "
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3615
      TabIndex        =   18
      Top             =   645
      Width           =   2880
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
      Height          =   240
      Left            =   135
      TabIndex        =   15
      Top             =   645
      Width           =   480
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "���ݺ�"
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
      Left            =   5790
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�:     ����:       �����:              �ѱ�: "
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   135
      TabIndex        =   39
      Top             =   1110
      Width           =   5880
   End
   Begin VB.Label lblDate 
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
      Height          =   240
      Left            =   135
      TabIndex        =   26
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "ʱ��"
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
      Left            =   2550
      TabIndex        =   27
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lblRoom 
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
      Height          =   240
      Left            =   5130
      TabIndex        =   21
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label lblDept 
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
      Height          =   240
      Left            =   135
      TabIndex        =   22
      Top             =   1620
      Width           =   480
   End
End
Attribute VB_Name = "frmStationRegistNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModul As Long
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset, mblnStartFactUseType As Boolean
Private mblnCard As Boolean, mintSysAppLimit As Integer
Private mfrmPatiInfo As frmPatiInfo
Private mstrYBPati As String, mlng�Һ�ID As Long, mlng����ID As Long
Private mblnOlnyBJYB As Boolean, mblnSharedInvoice As Boolean
Private mstr���� As String, mblnAppointment As Boolean, mblnChangeFeeType As Boolean
Private mstrAge As String, mstrFeeType As String, mstrGender As String, mstrClinic As String
Private mstrPassWord As String, mblnUnload As Boolean, mstrInsure As String
Private mlngDept As Long
Private Const SNCOLS = 10
Private Const SnArgCols = 7
Private mrsPlan As ADODB.Recordset, mblnInit As Boolean
Private mrsSNState As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset
Private mrsItems As ADODB.Recordset
Private mrsʱ��� As ADODB.Recordset
Private mrsInComes As ADODB.Recordset
Private mcolCardPayMode As Collection
Private mcolArrangeNo As Collection, mblnUpdateAge As Boolean
Private mlng����ID As Long, mintIDKind As Integer
Private mcur������� As Currency, mblnIntact As Boolean
Private mblnOK As Boolean, mstrCardPass As String
Private mstrNO As String, mlngSN As Long
Private mintInsure As Integer, mstrUseType As String
Private mdatLast As Date
Private mblnChangeByCode As Boolean
Private mstrCardNO As String
Private mcur����͸֧ As Currency
Private mlng���ż�¼ID As Long '�Һ�����ʱ�ļ�¼ID
Private Enum EM_REGISTFEE_MODE  '�Һŷ�����ȡ��ʽ
        EM_RG_���� = 0
        EM_RG_���� = 1
        EM_RG_���� = 2
End Enum
Private Enum EM_PATI_CHARGE_MODE    '�����շ�ģʽ
    EM_�Ƚ�������� = 0
    EM_�����ƺ���� = 1
End Enum
Private mRegistFeeMode As EM_REGISTFEE_MODE '�Һŷ�����ȡ��ʽ
Private mPatiChargeMode As EM_PATI_CHARGE_MODE    '�����շ�ģʽ

Private Type TYPE_MedicarePAR
    ҽ���ӿڴ�ӡƱ�� As Boolean
    ʹ�ø����ʻ�   As Boolean  'support�Һ�ʹ�ø����ʻ�
    �����Һ�  As Boolean    'support�����Һ�
    ���ղ����� As Boolean   'support�ҺŲ���ȡ������
    �Һż����Ŀ As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

Private Enum ViewMode
     V_��ͨ��
     v_ר�Һ�
     v_ר�Һŷ�ʱ��
     V_��ͨ�ŷ�ʱ��
End Enum
Private mViewMode As ViewMode

Private Type ty_ModulePara
    bln����ģ������ As Boolean
    lng������������ As Long
    blnĬ�Ϲ����� As Boolean
    blnĬ������ժҪ As Boolean
    byt�Һ�ģʽ As Byte
    bln�Һű���ˢ�� As Boolean
    bln����ʹ��Ԥ�� As Boolean
    blnסԺ���˹Һ� As Boolean
    bln�ҺŰ������Ұ��� As Boolean
    blnԤԼ�������Ұ��� As Boolean
    int�Һŷ�Ʊ��ӡ As Integer
    int�Һ�ƾ����ӡ As Integer
    intԤԼ�ҺŴ�ӡ As Integer
    bln������ѡ�� As Boolean
    lngԤԼ��Чʱ�� As Long
    bln�����շ�Ʊ�� As Boolean
    bln�˺����� As Boolean
    blnԤԼʱ�տ� As Boolean
    dblԤ��������鿨 As Double 'Ԥ�������ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤
    bln����ҽ�� As Boolean
    int��ԤԼ���� As Integer
    intͬ����Լ��           As Integer  'ͬ������Լ
    intͬ���޹���           As Integer
    blnͬ���޹Ҽ���         As Boolean
    int����ԤԼ������       As Integer
    int���˹Һſ�����       As Integer
    intר�ҺŹҺ�����       As Integer
    intר�Һ�ԤԼ����       As Integer
    strStationRegOrder As String  'ҽ��վ�Һ������ַ���
    blnShowAllPlan      As Boolean   ' �Ƿ���ʾ������ű�
End Type

Private mty_Para As ty_ModulePara
Private mstrPriceGrade As String, mintPriceGradeStartType As Integer
Private mobjRegister As clsRegist
Private mstrDef�ѱ� As String   'ȱʡ�ѱ�

Public Sub zlShowMe(ByVal frmMain As Object, ByVal objRegister As clsRegist, ByVal lngModul As Long, ByVal strDeptIDs As String, _
                    ByVal blnAppointment As Boolean, ByVal lng����ID As Long, ByRef strOutNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ��վ�Һ����
    '���:strDeptIDs-�Һſ���,֧�ֶ��,�ö��ŷָ�
    '     blnAppointment-�Ƿ�ԤԼ����
    '����:strOutNO-�Һųɹ���,�����Һŵ��ݺ�
    '����:������
    '����:2016-7-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    mblnAppointment = blnAppointment
    mlngModul = lngModul
    mlng����ID = lng����ID
    Set mobjRegister = objRegister
    If frmMain Is Nothing Then
        Me.Show
    Else
        Me.Show 1, frmMain
    End If
    If mblnOK = True Then
        strOutNO = mstrNO
        Unload Me
    End If
End Sub

Private Sub InitPara()
    Dim strValue As String
    On Error GoTo errH
    With mty_Para
        .bln����ģ������ = Val(gobjDatabase.GetPara("����ģ������", glngSys, 9000, "0")) = 1
        .lng������������ = Val(gobjDatabase.GetPara("������������", glngSys, 9000, 0))
        .blnĬ�Ϲ����� = Val(gobjDatabase.GetPara("Ĭ�Ϲ�����", glngSys, 9000, "0")) = 1
        .blnĬ������ժҪ = Val(gobjDatabase.GetPara("Ĭ������ժҪ", glngSys, 9000, "1")) = 1
        .byt�Һ�ģʽ = Val(gobjDatabase.GetPara("�Һ�ģʽ", glngSys, 9000, "0"))
        .bln����ʹ��Ԥ�� = Val(gobjDatabase.GetPara("����ʹ��Ԥ��", glngSys, 9000, "0")) = 1
        .blnסԺ���˹Һ� = Val(gobjDatabase.GetPara("����סԺ���˹Һ�", glngSys, 9000, "0")) = 1
        .int�Һŷ�Ʊ��ӡ = Val(gobjDatabase.GetPara("�Һŷ�Ʊ��ӡ��ʽ", glngSys, 9000, "0"))
        .int�Һ�ƾ����ӡ = Val(gobjDatabase.GetPara("�Һ�ƾ����ӡ��ʽ", glngSys, 9000, "0"))
        .intԤԼ�ҺŴ�ӡ = Val(gobjDatabase.GetPara("ԤԼ�Һŵ���ӡ��ʽ", glngSys, 9000, "0"))
        .bln������ѡ�� = Val(gobjDatabase.GetPara("������ѡ��", glngSys, 9000, "0")) = 1
        .bln�����շ�Ʊ�� = Val(gobjDatabase.GetPara("�ҺŹ����շ�Ʊ��", glngSys, 1121)) = 1
        .bln�˺����� = Val(gobjDatabase.GetPara("�����������Һ�", glngSys, 1111)) = 1
        .blnԤԼʱ�տ� = Val(gobjDatabase.GetPara("ԤԼʱ�տ�", glngSys, 9000, "0")) = 1
        .bln�Һű���ˢ�� = Val(gobjDatabase.GetPara("�Һű���ˢ��", glngSys, 9000)) = 1
        strValue = gobjDatabase.GetPara(28, glngSys, , "1|0")
        If InStr(strValue, "|") = 0 Then strValue = "1|0"
        .dblԤ��������鿨 = Val(Split(strValue, "|")(0))
        .bln����ҽ�� = Val(gobjDatabase.GetPara("����ҽ��", glngSys, 9000)) = 1
        .int��ԤԼ���� = Val(gobjDatabase.GetPara(66, glngSys, , 15))
        .intͬ����Լ�� = Val(gobjDatabase.GetPara("����ͬ����ԼN����", glngSys, 1111, 0))
        .intͬ���޹��� = Val(Split(gobjDatabase.GetPara("����ͬ���޹�N����", glngSys, 1111, 0) & "|", "|")(0))
        .blnͬ���޹Ҽ��� = Split(gobjDatabase.GetPara("����ͬ���޹�N����", glngSys, 1111, 0) & "|", "|")(1) = "1"
        .int���˹Һſ����� = Val(gobjDatabase.GetPara("���˹Һſ�������", glngSys, 1111, 0))
        .int����ԤԼ������ = Val(gobjDatabase.GetPara("����ԤԼ������", glngSys, 1111, 0))
        .intר�ҺŹҺ����� = Val(gobjDatabase.GetPara("ר�ҺŹҺ�����", glngSys, , 0))
        .intר�Һ�ԤԼ���� = Val(gobjDatabase.GetPara("ר�Һ�ԤԼ����", glngSys, , 0))
        strValue = gobjDatabase.GetPara("�������Ұ���", glngSys, 9000, "0|0") & "|"
        .bln�ҺŰ������Ұ��� = Val(Split(strValue, "|")(0)) = 1
        .blnԤԼ�������Ұ��� = Val(Split(strValue, "|")(1)) = 1
        .strStationRegOrder = gobjDatabase.GetPara("ҽ��վ�Һ��������", glngSys, 9000, "ҽ��,1|ִ��ʱ��,1|����,1|�ű�,1|��Ŀ,1")
        
        If .blnĬ������ժҪ Then
            cboRemark.TabStop = True
        Else
            cboRemark.TabStop = False
        End If
        If mblnAppointment Then
            mRegistFeeMode = EM_RG_����
        Else
            If (.byt�Һ�ģʽ = 0 Or .byt�Һ�ģʽ = 2) And gSysPara.bln��Һ�ģʽ = False Then
                mRegistFeeMode = EM_RG_����
            Else
                mRegistFeeMode = EM_RG_����
            End If
        End If
        '�Ƿ���ʾ������ű�
        .blnShowAllPlan = Val(gobjDatabase.GetPara("��ʾ������ű�", glngSys, 9000, "0")) = 1
    End With
    'ˢ��Ҫ����������
    mstrCardPass = gobjDatabase.GetPara(46, glngSys, , "0000000000")
    Call gobjControl.PicShowFlat(picInfo)
    '�շѺ͹ҺŹ���Ʊ��
    mblnSharedInvoice = gobjDatabase.GetPara("�ҺŹ����շ�Ʊ��", glngSys, 1121) = "1"
    '���ع��ùҺ�����ID
    If mblnSharedInvoice Then
        mlng�Һ�ID = Val(gobjDatabase.GetPara("�����շ�Ʊ������", glngSys, 1121, ""))
    Else
        mlng�Һ�ID = Val(gobjDatabase.GetPara("���ùҺ�Ʊ������", glngSys, mlngModul, ""))
    End If
    mlngDept = Val(gobjDatabase.GetPara("�������", glngSys, 1260, ""))
    If mlng�Һ�ID > 0 Then
        If Not ExistBill(mlng�Һ�ID, IIf(mblnSharedInvoice, 1, 4)) Then
            If mblnSharedInvoice Then
                gobjDatabase.SetPara "�����շ�Ʊ������", "0", glngSys, 1121
            Else
                gobjDatabase.SetPara "���ùҺ�Ʊ������", "0", glngSys, mlngModul
            End If
            mlng�Һ�ID = 0
        End If
    End If
    'Ʊ���ϸ����
    strValue = gobjDatabase.GetPara(24, glngSys, , "00000")
    gblnBill�Һ� = (Mid(strValue, IIf(mblnSharedInvoice, 1, 4), 1) = "1")
    mintSysAppLimit = Val(gobjDatabase.GetPara("�Һ�����ԤԼ����", glngSys))
    If mblnSharedInvoice Then
        '�Һ�������Ʊ��:42703
        mblnStartFactUseType = zlStartFactUseType("1")
    End If
    
    '�۸�ȼ�
    mintPriceGradeStartType = GetPriceGradeStartType()
    If mintPriceGradeStartType > 0 Then
        Call GetPriceGrade(gstrNodeNo, 0, 0, "", "", "", mstrPriceGrade)
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Function zlStartFactUseType(ByVal intƱ�� As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�ʹ����ʹ������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-10 16:11:47
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    On Error GoTo errHandle
    strSql = "Select  1 as ���� From Ʊ�����ü�¼ where Ʊ��=[1] and nvl(ʹ�����,'LXH')<>'LXH' and Rownum=1"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "���Ʊ���Ƿ�������ʹ������", intƱ��)
    
    If rsTemp.EOF Then
        Set rsTemp = Nothing: Exit Function
    End If
    Set rsTemp = Nothing
    zlStartFactUseType = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zl_GetInvoiceUserType(ByVal lng����ID As Long, ByVal lng��ҳId As Long, Optional intInsure As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ��ʹ�����
    '����:��Ʊ��ʹ�����
    '����:���˺�
    '����:2011-04-29 11:03:35
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errHandle
    strSql = "Select  Zl_Billclass([1],[2],[3]) as ʹ����� From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "��ȡƱ��ʹ�����", lng����ID, lng��ҳId, intInsure)
    zl_GetInvoiceUserType = Nvl(rsTemp!ʹ�����)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Function ExistBill(lngID As Long, bytKind As Byte) As Boolean
'���ܣ��ж��Ƿ����ָ����Ʊ������
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    
    strSql = "Select ID From Ʊ�����ü�¼ Where ID=[1] And Ʊ��=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "�������ID", lngID, bytKind)
    ExistBill = Not rsTmp.EOF
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function RefreshFact(Optional ByRef strFact As String) As Boolean
'������blnNew=�Ƿ��µ�����ʱ����,��ʱ���ڷ��ϸ���Ƶ�Ʊ���Ǳ��浱ǰ��
    If mblnStartFactUseType Then
        mstrUseType = zl_GetInvoiceUserType(Val(mrsInfo!����ID), 0, mintInsure)
    End If
    If gblnBill�Һ� Then
        mlng����ID = CheckUsedBill(IIf(mblnSharedInvoice, 1, 4), IIf(mlng����ID > 0, mlng����ID, mlng�Һ�ID), , mstrUseType)
        If mlng����ID <= 0 Then
            Select Case mlng����ID
                Case 0 '����ʧ��
                Case -1
                    MsgBox "��û�����ú͹��õ�" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End Select
            strFact = "": Exit Function
        Else
            '�ϸ�ȡ��һ������
            strFact = GetNextBill(mlng����ID)
        End If
    Else
        If mblnSharedInvoice Then
            strFact = gobjDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, 1121)
        Else
            strFact = gobjDatabase.GetPara("��ǰ�Һ�Ʊ�ݺ�", glngSys, 1111)
        End If
        strFact = IncStr(strFact)
        If mblnSharedInvoice Then
            gobjDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", strFact, glngSys, 1121
        Else
            gobjDatabase.SetPara "��ǰ�Һ�Ʊ�ݺ�", strFact, glngSys, 1111
        End If
    End If
    RefreshFact = True
End Function

'��ʼ��IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long, strTemp As String
    If gobjSquare Is Nothing Then CreateSquareCardObject Me, glngModul
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "��|����|0;ҽ|ҽ����|0;��|���֤��|0;��|�����|0;ס|סԺ��|0;��|�ֻ���|0", txtPatient)
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
        Set gobjSquare.objDefaultCard = objCard
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If

    Call GetRegInFor(g˽��ģ��, Me.Name, "idkind", strTemp)
    mintIDKind = Val(strTemp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
End Function

Private Sub cboAppointStyle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cmdPrice_Click()
    Dim bytRegistFeeMode As EM_REGISTFEE_MODE
    bytRegistFeeMode = mRegistFeeMode
    
    mRegistFeeMode = EM_RG_����
    If SaveData = False Then mRegistFeeMode = bytRegistFeeMode: Exit Sub
    
    mRegistFeeMode = bytRegistFeeMode
    mblnUpdateAge = False
    Call ReloadPage
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdReg_Click()
    Call LoadRegPlans(3)
End Sub
Private Function GetWorkTimeDefualtTime(ByVal strWorkName As String, ByVal str���� As String, ByVal strRegDate As String, Optional ByVal strCurSysDate As String = "") As Date
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ʱ�ε�ȱʡʱ��
    '���:strWorkName-����ʱ��
    '    strRegDate-�Һ����ڣ�yyyy-mm-dd)
    '    strCurSysDate-��ǰ��ȱʡʱ��
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-02-02 15:26:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtSysDate As Date, strDate As String
    Dim rsTime As ADODB.Recordset
    Dim dtRegDate As Date, strFilter As String
    
    On Error GoTo errHandle
    If strCurSysDate = "" Then
        dtSysDate = gobjDatabase.Currentdate
    Else
        dtSysDate = CDate(strCurSysDate)
    End If
    dtRegDate = CDate(strRegDate)
    
    If Format(dtRegDate, "yyyy-mm-dd") = Format(dtSysDate, "yyyy-mm-dd") Then
        '����
       GetWorkTimeDefualtTime = dtSysDate
    End If
    
    If mobjRegister.zlGetRegisterWorkTime_Record(rsTime) = False Then
        '���첻����,ȡ��ǰʱ��
        GetWorkTimeDefualtTime = CDate(Format(dtRegDate, "yyyy-mm-dd" & " " & Format(dtSysDate, "hh:mm:ss")))
        Exit Function
    End If
    strFilter = "ʱ���='" & strWorkName & "'"
    strFilter = strFilter & IIf(str���� = "", " And ����=NULL", " And =����='" & str���� & "'")
    strFilter = strFilter & " and վ��='" & gstrNodeNo & "'"
    
    rsTime.Filter = strFilter
    If rsTime.EOF Then
        strFilter = "ʱ���='" & strWorkName & "'"
        strFilter = strFilter & " and վ��='" & gstrNodeNo & "'"
        rsTime.Filter = strFilter
        If rsTime.EOF Then
            strFilter = "ʱ���='" & strWorkName & "' And վ��=NULL "
            strFilter = strFilter & IIf(str���� = "", " And ����=NULL", " And =����='" & str���� & "'")
            rsTime.Filter = strFilter
            If rsTime.EOF Then
                rsTime.Filter = "ʱ���='" & strWorkName & "' and ����=NULL and վ��=NULL"
                If rsTime.EOF Then
                    rsTime.Filter = 0
                    GetWorkTimeDefualtTime = dtSysDate: Exit Function
                End If
            End If
        End If
    End If
    If IsNull(rsTime!ȱʡʱ��) Then
        strDate = Format(dtRegDate, "yyyy-mm-dd") & " " & Format(rsTime!��ʼʱ��, "hh:mm:ss")
    Else
        strDate = Format(dtRegDate, "yyyy-mm-dd") & " " & Format(rsTime!ȱʡʱ��, "hh:mm:ss")
    End If
    rsTime.Filter = 0
    GetWorkTimeDefualtTime = CDate(strDate)
    Exit Function
errHandle:
    GetWorkTimeDefualtTime = gobjDatabase.Currentdate
End Function



Private Sub SetDefultRegTime()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ�ĹҺ�ʱ��
    '����:2018-02-05 10:00:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, dtSysDate As Date
    Dim rsTmp As ADODB.Recordset, rsTime As ADODB.Recordset
    Dim lng��¼ID As Long, lng��� As Long, str����ʱ�� As String
    
    On Error GoTo errH
    
    
    dtSysDate = gobjDatabase.Currentdate
    
    lblSn.Caption = ""
        
    lng��¼ID = Val(Nvl(mrsPlan!��¼ID))
    If lng��¼ID = 0 Then
      dtpTime.Value = Format(GetWorkTimeDefualtTime("����", "", Format(dtpDate.Value, "yyyy-mm-dd")), "hh:mm:ss")
      Exit Sub
    End If
    
    strSql = "Select ��ʼʱ��,��ֹʱ��,ȱʡԤԼʱ�� As ȱʡʱ�� From �ٴ������¼ Where ID=[1]"
    Set rsTime = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng��¼ID)

    If mViewMode = v_ר�Һŷ�ʱ�� Or (mViewMode = V_��ͨ�ŷ�ʱ�� And mblnAppointment) Then
        
        strSql = "Select ���,��ʼʱ�� From �ٴ�������ſ��� Where ��¼ID=[1] And Nvl(�Һ�״̬,0) = 0 Order By ���"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng��¼ID)
        If Not rsTmp.EOF Then
            lblSn.Caption = "���:" & Val(Nvl(rsTmp!���))
            'ʱ�ε�����ʱ��,ȡ��Сʱ��
            dtpTime.Value = Format(Nvl(rsTmp!��ʼʱ��), "hh:mm:ss"): Exit Sub
        End If
                
        
        If Format(dtpDate.Value, "yyyy-mm-dd") = Format(dtSysDate, "yyyy-mm-dd") Then
            dtpTime.Value = Format(dtSysDate, "hh:mm:ss"): Exit Sub
        End If

        'ʱ�ε�����ʱ��,ȡ��ʼʱ��
        If rsTime.EOF Then
            dtpTime.Value = Format(dtSysDate, "hh:mm:ss")
        Else
            If IsNull(rsTime!ȱʡʱ��) Then
                dtpTime.Value = Format(Nvl(rsTime!��ʼʱ��), "hh:mm:ss")
            Else
                dtpTime.Value = Format(Nvl(rsTime!ȱʡʱ��), "hh:mm:ss")
            End If
        End If
        Exit Sub
    End If
    
    If mViewMode = v_ר�Һ� Then
    
        If mobjRegister.zlGetRegisterNextSn__Visits(lng��¼ID, Format(dtSysDate, "yyyy-mm-dd HH:MM:SS"), InStr(gstrPrivs, ";�Ӻ�;"), mblnAppointment, False, lng���, str����ʱ��) Then
            If lng��� <> 0 Then lblSn.Caption = "���:" & lng���
            If mblnAppointment Then
                If Format(dtpDate.Value, "yyyy-mm-dd") > Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
                    str����ʱ�� = GetWorkTimeDefualtTime("����", "", Format(dtpDate.Value, "yyyy-mm-dd"))
                End If
            End If
            If IsDate(str����ʱ��) Then dtpTime.Value = Format(CDate(str����ʱ��), "hh:mm:ss"): Exit Sub
        End If
    End If
    
    If Not mblnAppointment Then
        dtpTime.Value = Format(dtSysDate, "hh:mm:ss"): Exit Sub
    End If
         
    If rsTime.EOF Then
        dtpTime.Value = Format(dtSysDate, "hh:mm:ss"): Exit Sub
    End If
    
    If IsNull(rsTime!ȱʡʱ��) Then
        dtpTime.Value = Format(Nvl(rsTime!��ʼʱ��), "hh:mm:ss")
    Else
        If CDate(Format(dtpDate.Value, "yyyy-mm-dd") & " " & Format(Nvl(rsTime!ȱʡʱ��), "hh:mm:ss")) > CDate(Format(rsTime!��ֹʱ��, "yyyy-mm-dd hh:mm:ss")) Then
            dtpTime.Value = Format(Nvl(rsTime!��ֹʱ��), "hh:mm:ss")
        ElseIf CDate(Format(dtpDate.Value, "yyyy-mm-dd") & " " & Format(Nvl(rsTime!ȱʡʱ��), "hh:mm:ss")) < CDate(Format(rsTime!��ʼʱ��, "yyyy-mm-dd hh:mm:ss")) Then
            dtpTime.Value = Format(Nvl(rsTime!��ʼʱ��), "hh:mm:ss")
        Else
            dtpTime.Value = Format(Nvl(rsTime!ȱʡʱ��), "hh:mm:ss")
        End If
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub GetAllҽ��()
    Dim strSql As String
    On Error GoTo errH
    
    strSql = "Select a.Id, a.����, Upper(a.����) As ����,b.����id,a.���" & _
            " From ��Ա�� a, ������Ա b, ��Ա����˵�� c" & _
            " Where a.Id = b.��Աid And a.Id = c.��Աid And c.��Ա���� = [1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order By a.���� Desc"
    Set mrsDoctor = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, "ҽ��")
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub LoadDoctor()
    With cboDoctor
        .Clear
        If Nvl(mrsPlan!ҽ������) = "" Then
            If mty_Para.bln����ҽ�� Then
                mrsDoctor.Filter = "����id=" & Val(Nvl(mrsPlan!����ID))
                
                Do While Not mrsDoctor.EOF
                    .AddItem IIf(IsNull(mrsDoctor!����), "", mrsDoctor!���� & "-") & mrsDoctor!����
                    .ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
                    If Nvl(mrsDoctor!����) = UserInfo.���� Then .ListIndex = .NewIndex
                    mrsDoctor.MoveNext
                Loop
                If .ListIndex < 0 Then
                    .ListIndex = 0
                End If
                .Enabled = True
                lblDoctor.Enabled = True
            Else
                mrsDoctor.Filter = "����='" & UserInfo.���� & "'"
                If mrsDoctor.RecordCount <> 0 Then
                    .AddItem IIf(IsNull(mrsDoctor!����), "", mrsDoctor!���� & "-") & mrsDoctor!����
                    .ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
                    .ListIndex = 0
                End If
                .Enabled = False
                lblDoctor.Enabled = False
            End If
        Else
            mrsDoctor.Filter = "����='" & Nvl(mrsPlan!ҽ������) & "'"
            If mrsDoctor.RecordCount <> 0 Then
                .AddItem IIf(IsNull(mrsDoctor!����), "", mrsDoctor!���� & "-") & mrsDoctor!����
                .ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
                .ListIndex = 0
            End If
            .Enabled = False
            lblDoctor.Enabled = False
        End If
    End With
End Sub

Private Sub cboPayMode_Click()
    If MCPAR.���ղ����� And cboPayMode.Text = mstrInsure Then
        chkBook.Enabled = False
        chkBook.Value = 0
    Else
        chkBook.Enabled = True
    End If
End Sub

Private Sub cboPayMode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cboRemark_Change()
    cboRemark.Tag = ""
End Sub

Private Sub cboRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboRemark.Tag <> "" Then gobjCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(cboRemark.Text) = "" Then gobjCommFun.PressKey vbKeyTab: Exit Sub
    If SelectMemo(Trim(cboRemark.Text)) = False Then
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
End Sub

Private Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '     blnUpper-�Ƿ�ת���ڴ�д
    '����:���ؼ�ƥ�䴮%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String

    If Val(gobjDatabase.GetPara("����ƥ��")) = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Private Function SelectMemo(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ����ժҪ
    '���:strInput-���봮;Ϊ��ʱ,��ʾȫ��
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-04 16:06:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strSql As String, strWhere As String
    Dim rsInfo As ADODB.Recordset
    Dim vRect As RECT, strKey As String
    strKey = GetMatchingSting(strInput, False)
    If strInput <> "" Then
        If gobjCommFun.IsCharChinese(cboRemark.Text) Then
             strWhere = " And  ���� like [1] "
        ElseIf gobjCommFun.IsNumOrChar(cboRemark.Text) Then
             strWhere = " And (���� like upper([1]) or ���� like upper([1]))"
        End If
    End If
    
    strSql = "" & _
     "   Select RowNum AS ID,����,����,����  " & _
     "   From ���ùҺ�ժҪ " & _
     "   Where 1=1 " & strWhere & _
     "   Order by ȱʡ��־"
     vRect = GetControlRect(cboRemark.hWnd)
     On Error GoTo Hd
     Set rsInfo = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "���ùҺ�ժҪ", False, _
                    "", "", False, False, True, vRect.Left, vRect.Top, cboRemark.Height, blnCancel, True, False, strKey)
     If blnCancel Then Exit Function
     If rsInfo Is Nothing Then
        If strInput = "" Then
            MsgBox "û�����ó���" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "ժҪ,�����ֵ����������", vbOKOnly + vbInformation, gstrSysName
        End If
        gobjCommFun.PressKey vbKeyTab: Exit Function
     End If
     gobjControl.CboSetText Me.cboRemark, Nvl(rsInfo!����)
     cboRemark.Tag = Nvl(rsInfo!����)
     gobjCommFun.PressKey vbKeyTab
     SelectMemo = True
     Exit Function
Hd:
    If gobjComlib.ErrCenter() = 1 Then Resume
    gobjComlib.SaveErrLog
End Function

Private Sub chkAll_Click()
    mty_Para.blnShowAllPlan = chkAll.Value <> 0
End Sub

Private Sub chkBook_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cmdNewPati_Click()
    Call zlExcuteMorePatiInfor
End Sub
 

Private Sub ResetDefault����()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������ȱʡ������Ϣ
    '����:���˺�
    '����:2017-10-27 15:16:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, bln���� As Boolean
    Dim lng����id As Long
    
    On Error GoTo errHandle
    lng����id = 0
    If mrsInfo Is Nothing Then chk����.Value = 0: Exit Sub
    If mrsInfo.State <> 1 Then chk����.Value = 0: Exit Sub
    If mrsPlan Is Nothing Then GoTo ReSet:
    If mrsPlan.RecordCount = 0 Then GoTo ReSet:
    lng����id = Val(Nvl(mrsPlan!����ID))
    
ReSet:
    lng����ID = Val(Nvl(mrsInfo!����ID))
    bln���� = zlPatiIsReturnVisit(lng����ID, lng����id)
    chk����.Value = IIf(bln����, 1, 0)
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub zlExcuteMorePatiInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�и���Ĳ�����Ϣ�޸�
    '����:���˺�
    '����:2017-10-27 14:35:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytFun As Byte, lng����ID As Long
    Dim lngOut����ID As Long
    Dim lng����id As Long
    
    On Error GoTo errHandle
    If mfrmPatiInfo Is Nothing Then Set mfrmPatiInfo = New frmPatiInfo
    bytFun = 2: lng����ID = 0
    If Not mrsInfo Is Nothing Then
        bytFun = 0
        lng����ID = Val(Nvl(mrsInfo!����ID))
        If lng����ID = 0 Then Exit Sub
    End If
    If mrsPlan.RecordCount <> 0 Then lng����id = Val(Nvl(mrsPlan!����ID))
    If mfrmPatiInfo.ShowMe(Me, bytFun, lng����ID, lngOut����ID, lng����id) = False Then Exit Sub
    If Not mfrmPatiInfo Is Nothing Then Unload mfrmPatiInfo
    Set mfrmPatiInfo = Nothing
    
    txtPatient.Text = "-" & lngOut����ID
    GetPatient IDKind.GetCurCard, txtPatient.Text, False
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub


Private Sub cmdTime_Click()
    If SelectTimeSn = False Then Exit Sub

End Sub

Private Sub cmdOther_Click()
    Call LoadRegPlans(3, , True)
End Sub

Private Sub dtpDate_Change()
    Call LoadRegPlans(1)
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub dtpTime_GotFocus()
    Call cmdTime_Click
End Sub

Private Sub dtpTime_Validate(Cancel As Boolean)
    If Format(dtpDate.Value, "YYYY-MM-DD") = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") Then
        If Format(dtpTime.Value, "hh:mm:ss") < Format(gobjDatabase.Currentdate, "hh:mm:ss") Then
            MsgBox "ԤԼʱ�䲻��С�ڵ�ǰʱ��!", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Activate()
    If mblnUnload Then mblnUnload = False: Unload Me: Exit Sub
    If mblnInit And Not mrsInfo Is Nothing Then
        If txtReg.Enabled And txtReg.Visible Then txtReg.SetFocus
    End If
    mblnInit = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsInfo = Nothing
    If Not mobjIDCard Is Nothing Then
         Call mobjIDCard.SetEnabled(False)
         Set mobjIDCard = Nothing
     End If
     mstrDef�ѱ� = ""
     mstrFeeType = ""
     If Not mobjICCard Is Nothing Then
         Call mobjICCard.SetEnabled(False)
         Set mobjICCard = Nothing
     End If
     mintIDKind = IDKind.IDKind
     Call SaveRegInFor(g˽��ģ��, Me.Name, "idkind", mintIDKind)
     gobjDatabase.SetPara "��ʾ������ű�", IIf(mty_Para.blnShowAllPlan, 1, 0), glngSys, 9000
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strExpand As String
    Dim strOutCardNO As String, strOutPatiInforXML As String
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        'ϵͳIC��
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call GetPatient(objCard, txtPatient.Text, True)
            End If
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
'    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
'    txtPatient.Text = strOutCardNO
'
'    If txtPatient.Text <> "" Then
'        Call GetPatient(objCard, txtPatient.Text, True)
'    End If
End Sub

Private Sub txtRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub chkBook_Click()
    If mrsPlan Is Nothing Then Exit Sub
    If mrsPlan.RecordCount = 0 Then Exit Sub
    Call LoadFeeItem(Val(Nvl(mrsPlan!��ĿID)), chkBook.Value = 1, mstrPriceGrade)
End Sub

Private Sub cmdCancel_Click()
    If txtPatient.Text <> "" Then
        If MsgBox("�Ƿ���յ�ǰ������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ClearPatient
        End If
        Exit Sub
    End If
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp "zl9RegEvent", Me.hWnd, "frmRegistEdit"
    Exit Sub
End Sub

Private Function CheckBrushCard(ByVal dblMoney As Double, ByVal lngҽ�ƿ����ID As Long, ByVal bln���ѿ� As Boolean, _
                                ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ˢ��
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsMoney As ADODB.Recordset, str���� As String, lng����ID As Long
    On Error GoTo errHandle
    '68991
    If mRegistFeeMode <> EM_RG_���� Then CheckBrushCard = True: Exit Function
    If dblMoney = 0 Then
        CheckBrushCard = True: Exit Function
    End If
    If Not (cboPayMode.Visible And cboPayMode.Enabled) Then
        CheckBrushCard = True: Exit Function
    End If
    If cboPayMode.ItemData(cboPayMode.ListIndex) <> -1 Then
        CheckBrushCard = True: Exit Function
    End If
    If lngҽ�ƿ����ID = 0 Then
        MsgBox cboPayMode.Text & "�쳣,����!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If gobjSquare.objSquareCard Is Nothing Then
        MsgBox "ʹ��" & cboPayMode.Text & "֧�������ȳ�ʼ���ӿڲ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call zlGetClassMoney(rsMoney, rsItems, rsIncomes)
    
     '����ˢ������
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln���ѿ� As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl��� As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String, _
    Optional ByRef bln�˷� As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln���� As Boolean = False, _
    Optional ByVal bln�����ֹ As Boolean = True, _
    Optional ByRef varSquareBalance As Variant, _
    Optional ByVal blnתԤ�� As Boolean = False, _
    Optional ByVal blnAllPay As Boolean = False, _
    Optional ByVal strXmlIn As String = "", _
    Optional ByVal str������Դ As String, _
    Optional ByVal lng����ID As Long) As Boolean
    str���� = Trim(mstrAge)
    If Not mrsInfo Is Nothing Then lng����ID = Val(Nvl(mrsInfo!����ID))
   If gobjSquare.objSquareCard.zlBrushCard(Me, glngModul, rsMoney, lngҽ�ƿ����ID, bln���ѿ�, _
    txtPatient.Text, NeedName(mstrGender), str����, dblMoney, mstrCardNO, mstrPassWord, _
    False, True, False, True, Nothing, False, True, "", "1", lng����ID) = False Then Exit Function
    
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, glngModul, lngҽ�ƿ����ID, _
        bln���ѿ�, mstrCardNO, dblMoney, "", "") = False Then Exit Function

    CheckBrushCard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetClassMoney(ByRef rsMoney As ADODB.Recordset, ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ,��ʼ��֧�����(�շ����,ʵ�ս��)
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql  As String
    
    Err = 0: On Error GoTo Errhand:
    
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        If .State = adStateOpen Then .Close
        .Fields.Append "�շ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic

        rsItems.Filter = 0
        If rsItems.RecordCount <> 0 Then rsItems.MoveFirst
        Do While Not rsItems.EOF
            rsIncomes.Filter = "��ĿID=" & rsItems!��ĿID
            rsMoney.Filter = "�շ����='" & Nvl(rsItems!���, "��") & "'"
            If rsMoney.EOF Then
                .AddNew
            Else
                rsMoney.Filter = 0
            End If
            !�շ���� = Nvl(rsItems!���, "��")
            Do While Not rsIncomes.EOF
                !��� = Val(Nvl(!���)) + Val(Nvl(rsIncomes!ʵ��))
                rsIncomes.MoveNext
            Loop
            .Update
            rsItems.MoveNext
        Loop
    End With
    rsMoney.Filter = 0
    zlGetClassMoney = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    If SaveData = False Then
        If (mViewMode = v_ר�Һ� Or mViewMode = v_ר�Һŷ�ʱ��) And mlng���ż�¼ID <> 0 Then Call CancelRegNo(mlng���ż�¼ID)
        Exit Sub
    End If
    mblnUpdateAge = False
    Call ReloadPage
    mblnOK = True
    Unload Me
End Sub

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Һ�����
    '����:����ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-02-01 15:46:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnSlipPrint As Boolean, blnInvoicePrint As Boolean, int�۸񸸺� As Integer, blnBalance As Boolean
    Dim k As Integer, i As Integer, j As Integer, strNO As String, strFactNO As String
    Dim cllPro As New Collection, strSql As String, str�Ǽ�ʱ�� As String, str����ʱ�� As String
    Dim curԤ�� As Currency, cur���� As Currency, cur�ֽ� As Currency, str����NO As String
    Dim lngSN As Long, lng�Һſ���ID As Long, lng����ID As Long, byt���� As Byte
    Dim lngҽ�ƿ����ID As Long, bln���ѿ� As Boolean, blnNoDoc As Boolean, strBalanceStyle As String
    Dim rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset, cllProAfter As New Collection
    Dim blnTrans As Boolean, blnNotCommit As Boolean, strAdvance As String, rsTemp As ADODB.Recordset
    Dim lngҽ��ID As Long, blnOneCard As Boolean, rsTmp As ADODB.Recordset, dat����ʱ�� As Date
    Dim cllCardPro As Collection, cllTheeSwap As Collection, strNotValiedNos As String
    Dim strDay As String, blnAppointPrint As Boolean, str���ʽ As String, mstr����IDs As String
    Dim rs���ʽ As ADODB.Recordset, strҽ�� As String, blnAdd As Boolean, blnNotWork As Boolean
    Dim dat�Ǽ�ʱ�� As Date, rsʱ��� As ADODB.Recordset, strʱ��� As String
    Dim bytMode As Byte, rsCheck As ADODB.Recordset, datԤԼʱ�� As Date
    Dim strResult As String, blnר�Һ� As Boolean
    Dim dblTotal  As Double
    
    If CheckValied = False Then Exit Function
    If Not mrsInfo Is Nothing Then
        strSql = "Select Zl_Fun_���˹Һż�¼_Check([1],[2],[3],[4],[5],[6]) As ����� From Dual"
        If mblnAppointment Then
            bytMode = 1
            datԤԼʱ�� = CDate(Format(dtpDate.Value, "yyyy-mm-dd"))
        Else
            bytMode = 0
            datԤԼʱ�� = CDate(Format(gobjDatabase.Currentdate, "yyyy-mm-dd"))
        End If
        
        blnר�Һ� = Nvl(mrsPlan!ҽ������) <> ""
        Set rsCheck = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, bytMode, Val(Nvl(mrsInfo!����ID)), Trim(txtReg.Tag), Val(Nvl(mrsPlan!��¼ID)), datԤԼʱ��, IIf(blnר�Һ�, 1, 0))
        If Not rsCheck.EOF Then
            strResult = Nvl(rsCheck!�����)
            If Val(Mid(strResult, 1, 1)) <> 0 Then
                MsgBox Mid(strResult, 3), vbInformation, gstrSysName
                Exit Function
            End If
        Else
            MsgBox "��Ч�Լ��ʧ��,�޷�������", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strSql = "Select ���,����,ҽԺ����,���㷽ʽ From һ��ͨĿ¼ Where ���� = 1 And ���㷽ʽ = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, cboPayMode.Text)
    blnOneCard = rsTmp.RecordCount <> 0
    
    If mblnAppointment And mty_Para.blnԤԼʱ�տ� = False Then
        blnSlipPrint = False
    Else
        Select Case Val(mty_Para.int�Һ�ƾ����ӡ)
            Case 0    '����ӡ
                blnSlipPrint = False
            Case 1    '�Զ���ӡ
                If InStr(gstrPrivs, ";���˹Һ�ƾ��;") > 0 Then
                    blnSlipPrint = True
                Else
                    blnSlipPrint = False
                    MsgBox "��û��" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "ƾ����ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                End If
            Case 2    'ѡ���ӡ
                If MsgBox("Ҫ��ӡ" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "ƾ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    If InStr(gstrPrivs, ";���˹Һ�ƾ��;") > 0 Then
                        blnSlipPrint = True
                    Else
                        blnSlipPrint = False
                        MsgBox "��û��" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "ƾ����ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                    End If
                Else
                    blnSlipPrint = False
                End If
        End Select
    End If
    
    If mRegistFeeMode = EM_RG_���� Or mRegistFeeMode = EM_RG_���� Or (mblnAppointment And mty_Para.blnԤԼʱ�տ� = False) Then
        blnInvoicePrint = False
    Else
        If Not (mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��) Then
            Select Case Val(mty_Para.int�Һŷ�Ʊ��ӡ)
                Case 0    '����ӡ
                    blnInvoicePrint = False
                Case 1    '�Զ���ӡ
                    If InStr(gstrPrivs, ";�Һŷ�Ʊ��ӡ;") > 0 Then
                        blnInvoicePrint = True
                    Else
                        blnInvoicePrint = False
                        MsgBox "��û��" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "��Ʊ��ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                    End If
                Case 2    'ѡ���ӡ
                    If MsgBox("Ҫ��ӡ" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "��Ʊ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        If InStr(gstrPrivs, ";�Һŷ�Ʊ��ӡ;") > 0 Then
                            blnInvoicePrint = True
                        Else
                            blnInvoicePrint = False
                            MsgBox "��û��" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "��Ʊ��ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                        End If
                    Else
                        blnInvoicePrint = False
                    End If
            End Select
        End If
    End If
    
    If mblnAppointment And mty_Para.blnԤԼʱ�տ� = False Then
        Select Case Val(mty_Para.intԤԼ�ҺŴ�ӡ)
            Case 0
                blnAppointPrint = False
            Case 1
                If InStr(gstrPrivs, ";ԤԼ�Һŵ�;") > 0 Then
                    blnAppointPrint = True
                Else
                    blnAppointPrint = False
                    MsgBox "��û��ԤԼ" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "����ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                End If
            Case 2
                If InStr(gstrPrivs, ";ԤԼ�Һŵ�;") > 0 Then
                    If MsgBox("Ҫ��ӡԤԼ" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        blnAppointPrint = True
                    Else
                        blnAppointPrint = False
                    End If
                Else
                    MsgBox "��û��ԤԼ" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "����ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                    blnAppointPrint = False
                End If
        End Select
    Else
        blnAppointPrint = False
    End If
    
    If blnInvoicePrint Or (mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��) Then
        If RefreshFact(strFactNO) = False Then Exit Function
    End If
    
    If mblnAppointment Then
        If mRegistFeeMode = EM_RG_���� And mty_Para.blnԤԼʱ�տ� Then
            MsgBox "��֧�������ƺ���㲡�˵�ԤԼ�տ�Һţ�", vbInformation, gstrSysName
            Exit Function
        End If
        If mty_Para.blnԤԼʱ�տ� Then
            If Not mRegistFeeMode = EM_RG_���� Then
                If cboPayMode.Text = "Ԥ����" Then
                    curԤ�� = Val(lblTotal.Caption)
                Else
                    If cboPayMode.Text = mstrInsure Then
                        cur���� = Val(lblTotal.Caption)
                    Else
                        blnBalance = True
                        cur�ֽ� = Val(lblTotal.Caption)
                    End If
                End If
            End If
        Else
            blnBalance = False
        End If
    Else
        If Not mRegistFeeMode = EM_RG_���� Then
            If cboPayMode.Text = "Ԥ����" Then
                curԤ�� = Val(lblTotal.Caption)
            Else
                If cboPayMode.Text = "�����ʻ�" Then
                    cur���� = Val(lblTotal.Caption)
                Else
                    blnBalance = True
                    cur�ֽ� = Val(lblTotal.Caption)
                End If
            End If
        End If
    End If
    
    If Val(curԤ��) <> 0 Then
        If Not gobjDatabase.PatiIdentify(Me, glngSys, Nvl(mrsInfo!����ID), Val(curԤ��), mlngModul, 1, , _
                             IIf(-1 * mty_Para.dblԤ��������鿨 >= Val(curԤ��), False, True), True, mstr����IDs, (mty_Para.dblԤ��������鿨 <> 0), (mty_Para.dblԤ��������鿨 = 2)) Then Exit Function
    End If
    
    ReadRegistPrice Val(Nvl(mrsPlan!��ĿID)), chkBook.Value = 1, False, mstrFeeType, rsItems, rsIncomes, _
        Val(Nvl(mrsInfo!����ID)), mintInsure, txtReg.Tag, IIf(mblnAppointment, 1, 0), , mstrPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
    
    If mblnAppointment = False Or (mblnAppointment = True And mty_Para.blnԤԼʱ�տ�) Then
        If zlIsAllowPatiChargeFeeMode(ZVal(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!����ģʽ))) = False Then Exit Function
    End If

    If blnBalance Then
        For i = 1 To mcolCardPayMode.Count
            If cboPayMode.Text = mcolCardPayMode.Item(i)(1) Then
                lngҽ�ƿ����ID = mcolCardPayMode.Item(i)(3)
                bln���ѿ� = Val(mcolCardPayMode.Item(i)(5)) = 1
                strBalanceStyle = mcolCardPayMode.Item(i)(6)
            End If
        Next i
        If CheckBrushCard(Val(cur�ֽ�), lngҽ�ƿ����ID, bln���ѿ�, rsItems, rsIncomes) = False Then Exit Function
        If strBalanceStyle <> "" Then
            strBalanceStyle = strBalanceStyle & "," & Val(cur�ֽ�) & ",,1"
        Else
            strBalanceStyle = cboPayMode.Text & "," & Val(cur�ֽ�) & ",,0"
        End If
    End If
    
    str�Ǽ�ʱ�� = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
    dat�Ǽ�ʱ�� = gobjDatabase.Currentdate
    
    If mblnAppointment Then
        strDay = zlGet��ǰ���ڼ�(dtpDate.Value)
    Else
        strDay = zlGet��ǰ���ڼ�
    End If
    
    '��ȡ����ʱ��
    blnAdd = False
    If mblnAppointment Then
        mlngSN = 0
        If IsNull(mrsPlan!��ʼʱ��) = False Then
            str����ʱ�� = "To_Date('" & Format(dtpDate, "yyyy-mm-dd") & " " & Format(dtpTime, "hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')"
            dat����ʱ�� = CDate(Format(dtpDate, "yyyy-mm-dd") & " " & Format(dtpTime, "hh:mm:ss"))
        Else
            str����ʱ�� = "To_Date('" & Format(dtpDate, "yyyy-mm-dd") & " " & Format(dtpTime, "hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')"
            dat����ʱ�� = CDate(Format(dtpDate, "yyyy-mm-dd") & " " & Format(dtpTime, "hh:mm:ss"))
        End If
        If mViewMode = v_ר�Һŷ�ʱ�� Then
            strʱ��� = "Select Rownum As Id, ���, To_Char(��ʼʱ��, 'hh24') || ':00' As ʱ���, To_Char(��ʼʱ��, 'hh24:mi') As ��ʼʱ��," & vbNewLine & _
                    "       To_Char(��ֹʱ��, 'hh24:mi') As ����ʱ��, ��ʼʱ�� As ��ϸ��ʼʱ��, ��ֹʱ�� As ��ϸ����ʱ�� " & vbNewLine & _
                    "From �ٴ�������ſ���" & vbNewLine & _
                    "Where ��¼id = [1] And Nvl(�Һ�״̬,0) = 0 And Nvl(�Ƿ�ԤԼ,0)=1 And Trunc(��ʼʱ��) = [2]" & vbNewLine & _
                    "Order By ��ϸ��ʼʱ��"
            Set rsʱ��� = gobjDatabase.OpenSQLRecord(strʱ���, Me.Caption, Val(Nvl(mrsPlan!��¼ID)), CDate(Format(dtpDate.Value, "yyyy-mm-dd")))
            If rsʱ���.RecordCount = 0 Then
                MsgBox "��ǰѡ��ķ�ʱ�κű��޿���ʱ�Σ��޷�ԤԼ��", vbInformation, gstrSysName
                Exit Function
            End If
            strSql = "Select a.���,a.��ʼʱ�� From �ٴ�������ſ��� A Where a.��¼ID=[1] And Nvl(a.�Һ�״̬,0) = 0 And Not Exists (Select 1 From �ٴ�����Һſ��Ƽ�¼ Where ����=1 And ��¼ID=[1] And ���=a.��� And ���Ʒ�ʽ=3) Order By a.���"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!��¼ID)))
            If rsTmp.RecordCount <> 0 Then
                Dim blnFind As Boolean
                blnFind = False
                Do While blnFind = False
                    If rsTmp.EOF Then
                        Exit Do
                    Else
                        If str����ʱ�� = "To_Date('" & Format(Nvl(rsTmp!��ʼʱ��), "yyyy-mm-dd hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')" Then blnFind = True: Exit Do
                        rsTmp.MoveNext
                    End If
                Loop
                If blnFind Then
                    mlngSN = Val(Nvl(rsTmp!���))
                Else
                    If MsgBox("��Ϊ����ԭ��,ѡ���ʱ���Ѳ�����,�Ƿ��Զ���ȡ����ʱ�ν���" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
                        Exit Function
                    End If
                    rsTmp.MoveFirst
                    mlngSN = Val(Nvl(rsTmp!���))
                    str����ʱ�� = "To_Date('" & Format(Nvl(rsTmp!��ʼʱ��), "yyyy-mm-dd hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')"
                    dat����ʱ�� = CDate(Format(Nvl(rsTmp!��ʼʱ��), "yyyy-mm-dd hh:mm:ss"))
                End If
            Else
                blnAdd = True
                strSql = "Select Max(���) As ��� From �ٴ�������ſ��� Where ��¼ID=[1]"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!��¼ID)))
                mlngSN = Val(Nvl(rsTmp!���)) + 1
            End If
        End If
        If Val(Nvl(mrsPlan!��¼ID)) = 0 Then blnNotWork = True
    Else
        Select Case mViewMode
            Case V_��ͨ��
                str����ʱ�� = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                If Val(Nvl(mrsPlan!��¼ID)) = 0 Then blnNotWork = True
            Case V_��ͨ�ŷ�ʱ��
                str����ʱ�� = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                If Val(Nvl(mrsPlan!��¼ID)) = 0 Then blnNotWork = True
            Case v_ר�Һ�
                str����ʱ�� = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                If Val(Nvl(mrsPlan!��¼ID)) = 0 Then blnNotWork = True
            Case v_ר�Һŷ�ʱ��
                If Val(Nvl(mrsPlan!��¼ID)) = 0 Then
                    str����ʱ�� = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                    blnNotWork = True
                Else
                    'ȡ��С����ʱ���
                    strSql = "Select ���,��ʼʱ�� From �ٴ�������ſ��� Where ��¼ID=[1] And Nvl(�Һ�״̬,0) = 0 Order By ���"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!��¼ID)))
                    If rsTmp.RecordCount <> 0 Then
                        mlngSN = Val(Nvl(rsTmp!���))
                        str����ʱ�� = "To_Date('" & Format(Nvl(rsTmp!��ʼʱ��), "yyyy-mm-dd hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        blnAdd = True
                        strSql = "Select Max(���) As ��� From �ٴ�������ſ��� Where ��¼ID=[1]"
                        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!��¼ID)))
                        mlngSN = Val(Nvl(rsTmp!���)) + 1
                        str����ʱ�� = "To_Date('" & gobjDatabase.Currentdate & "','yyyy-mm-dd hh24:mi:ss')"
                    End If
                End If
        End Select
    End If
    
    If blnAdd = False And Not blnNotWork Then
        If Val(Nvl(mrsPlan!�ѹ�)) >= Val(Nvl(mrsPlan!�޺�)) And Val(Nvl(mrsPlan!�޺�)) <> 0 Then
            blnAdd = True
        End If
        If Val(Nvl(mrsPlan!��Լ)) >= Val(Nvl(mrsPlan!��Լ)) And Val(Nvl(mrsPlan!��Լ)) <> 0 Then
            blnAdd = True
        End If
    End If
    
    mlng���ż�¼ID = 0
    If (mViewMode = v_ר�Һ� Or mViewMode = v_ר�Һŷ�ʱ��) And Val(Nvl(mrsPlan!��¼ID)) <> 0 Then
        If ReserveRegNo(Nvl(mrsPlan!�ű�), True, mViewMode = v_ר�Һŷ�ʱ��, str����ʱ��, mlngSN, "ҽ��վ����", Val(Nvl(mrsPlan!��¼ID))) = False Then Exit Function
        mlng���ż�¼ID = Val(Nvl(mrsPlan!��¼ID))
    End If
    
    If blnAdd And InStr(gstrPrivs, ";�Ӻ�;") = 0 Then
        MsgBox "��û�мӺ�Ȩ�ޣ��޷��Ե�ǰ�ű����" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If blnAdd = False Then
        If InStr(gstrPrivs, ";�Ӻ�;") > 0 Then blnAdd = True
    End If
    If blnNotWork Then blnAdd = blnNotWork
    
    If cboAppointStyle.Visible And mblnAppointment And blnAdd = False Then
        strSql = "Select Zl_Fun_Get�ٴ�����ԤԼ״̬([1],[2],[3],[4]) As ԤԼ��� From Dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!��¼ID)), dat����ʱ��, mlngSN, NeedName(cboAppointStyle.Text))
        If rsTemp.EOF Then
            MsgBox "��ǰѡ���ԤԼ��ʽ�޷�ԤԼ,��ѡ������ԤԼ��ʽ!", vbInformation, gstrSysName
            If cboAppointStyle.Enabled And cboAppointStyle.Visible Then cboAppointStyle.SetFocus
            Exit Function
        Else
            If Val(Mid(Nvl(rsTemp!ԤԼ���), 1, 1)) <> 0 Then
                MsgBox "��ǰѡ���ԤԼ��ʽ�޷�ԤԼ,��ѡ������ԤԼ��ʽ!" & vbCrLf & "ԭ��:" & Mid(Nvl(rsTemp!ԤԼ���), InStr(Nvl(rsTemp!ԤԼ���), "|") + 1), vbInformation, gstrSysName
                If cboAppointStyle.Enabled And cboAppointStyle.Visible Then cboAppointStyle.SetFocus
                Exit Function
            End If
        End If
    End If
    
    strSql = "Select Zl_�ٴ���������_Check([1],[2],[3]) As �����Լ�� From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!��¼ID)), mstrGender, mstrAge)
    If rsTemp.EOF Then
        MsgBox "��ǰѡ��Ĳ��˲����øúű�!", vbInformation, gstrSysName
        Exit Function
    Else
        If Val(Mid(Nvl(rsTemp!�����Լ��), 1, 1)) <> 0 Then
            MsgBox "��ǰѡ��Ĳ��˲����øúű�!" & vbCrLf & "ԭ��:" & Mid(Nvl(rsTemp!�����Լ��), InStr(Nvl(rsTemp!�����Լ��), "|") + 1), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    lng�Һſ���ID = Val(Nvl(mrsPlan!����ID))
    lng����ID = gobjDatabase.GetNextId("���˽��ʼ�¼")
    byt���� = IIf(chk����.Value = 1, 1, 0)
    

    lngSN = mlngSN
    strNO = gobjDatabase.GetNextNo(12)
    
    rsItems.Filter = ""
    strҽ�� = NeedName(cboDoctor.Text)
    If cboDoctor.ListCount = 0 Then
        lngҽ��ID = 0
    Else
        lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    strSql = "Select ���� From ҽ�Ƹ��ʽ Where ���� = [1]"
    Set rs���ʽ = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, Nvl(mrsInfo!ҽ�Ƹ��ʽ))
    If rs���ʽ.RecordCount <> 0 Then
        str���ʽ = Nvl(rs���ʽ!����)
    Else
        strSql = "Select ���� From ҽ�Ƹ��ʽ Where ȱʡ��־ = 1"
        Set rs���ʽ = gobjDatabase.OpenSQLRecord(strSql, App.ProductName)
        If rs���ʽ.RecordCount <> 0 Then
            str���ʽ = Nvl(rs���ʽ!����)
        End If
    End If
    
    dblTotal = 0
    If mRegistFeeMode = EM_RG_���� Then
        dblTotal = GetRegistMoney(True, False)
        '�ҺŷѴ�Ϊ���ұ���Ϊ���۵����Ų�������NO
       If dblTotal <> 0 Then str����NO = gobjDatabase.GetNextNo(13)
    End If
    
    
    k = 1: rsItems.MoveFirst
    For i = 1 To rsItems.RecordCount
        int�۸񸸺� = k
        rsIncomes.Filter = "��ĿID=" & rsItems!��ĿID
        For j = 1 To rsIncomes.RecordCount
            strSql = _
            "zl_���˹Һż�¼_����_INSERT(" & ZVal(Nvl(mrsPlan!��¼ID)) & "," & ZVal(Nvl(mrsInfo!����ID)) & "," & IIf(mstrClinic = "", "NULL", mstrClinic) & ",'" & txtPatient.Text & "','" & mstrGender & "'," & _
                     "'" & mstrAge & "','" & str���ʽ & "','" & mstrFeeType & "','" & strNO & "'," & _
                     "'" & IIf(blnInvoicePrint = False, "", "") & "'," & k & "," & IIf(int�۸񸸺� = k, "NULL", int�۸񸸺�) & "," & IIf(rsItems!���� = 2, 1, "NULL") & "," & _
                     "'" & rsItems!��� & "'," & rsItems!��ĿID & "," & rsItems!���� & "," & rsIncomes!���� & "," & _
                     rsIncomes!������ĿID & ",'" & rsIncomes!�վݷ�Ŀ & "','" & IIf(blnBalance, IIf(strBalanceStyle = "", cboPayMode.Text, strBalanceStyle), "") & "'," & _
                     IIf(mRegistFeeMode = EM_RG_����, 0, rsIncomes!Ӧ��) & "," & IIf(mRegistFeeMode = EM_RG_����, 0, rsIncomes!ʵ��) & "," & _
                     lng�Һſ���ID & "," & lng�Һſ���ID & "," & IIf(rsItems!ִ�п���ID = 0, lng�Һſ���ID, rsItems!ִ�п���ID) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                     str����ʱ�� & "," & str�Ǽ�ʱ�� & "," & _
                     "'" & strҽ�� & "'," & ZVal(lngҽ��ID) & "," & IIf(rsItems!���� = 3, 1, IIf(rsItems!���� = 4, 2, 0)) & "," & IIf(lbl��.Visible, 1, 0) & "," & _
                     "'" & mrsPlan!�ű� & "','" & IIf(strҽ�� = UserInfo.����, lblRoomName.Caption, "") & "'," & ZVal(lng����ID) & "," & IIf(blnInvoicePrint = False, "NULL", ZVal(mlng����ID)) & "," & _
                     ZVal(IIf(k = 1, curԤ��, 0)) & "," & ZVal(IIf(k = 1, cur�ֽ�, 0)) & "," & _
                     ZVal(IIf(k = 1, cur����, 0)) & "," & ZVal(Nvl(rsItems!���մ���ID, 0)) & "," & _
                     ZVal(Nvl(rsItems!������Ŀ��, 0)) & "," & ZVal(Nvl(rsIncomes!ͳ����, 0)) & "," & _
                     "'" & Trim(cboRemark.Text) & "'," & IIf(mblnAppointment, IIf(mty_Para.blnԤԼʱ�տ�, 0, 1), 0) & "," & IIf(mty_Para.bln�����շ�Ʊ��, 1, 0) & ",'" & rsItems!���ձ��� & "'," & byt���� & "," & ZVal(lngSN) & ",Null," & _
                     IIf(mblnAppointment, 1, 0) & ",'" & IIf(cboAppointStyle.Visible, cboAppointStyle.Text, "") & "'," & _
                     0 & ","
            '�����id_In   ����Ԥ����¼.�����id%Type := Null,
            strSql = strSql & "" & IIf(lngҽ�ƿ����ID <> 0 And bln���ѿ� = False, lngҽ�ƿ����ID, "NULL") & ","
            '���㿨���_In ����Ԥ����¼.���㿨���%Type := Null,
            strSql = strSql & "" & IIf(lngҽ�ƿ����ID <> 0 And bln���ѿ�, lngҽ�ƿ����ID, "NULL") & ","
            '����_In       ����Ԥ����¼.����%Type := Null,
            strSql = strSql & "'" & mstrCardNO & "',"
            '������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
            strSql = strSql & " NULL,"
            '����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
            strSql = strSql & " NULL,"
            '������λ_In   ����Ԥ����¼.������λ%Type := Null
            strSql = strSql & " NULL,"
            '  ��������_In   Number:=0
            strSql = strSql & IIf(blnAdd, 1, 0) & ","
            '  ����_IN       ���˹Һż�¼.����%type:=null,
            strSql = strSql & IIf(mintInsure = 0, "NULL", mintInsure) & ","
            '  ����ģʽ_IN   NUMBER :=0,
            strSql = strSql & IIf(mPatiChargeMode = EM_�����ƺ����, 1, 0) & ","
            '  ���ʷ���_IN Number:=0
            strSql = strSql & IIf(mRegistFeeMode = EM_RG_����, 1, 0) & ","
            '  �˺�����_IN Number:=1
            strSql = strSql & IIf(mty_Para.bln�˺�����, 1, 0) & ","
            '  ��Ԥ������ids_In Varchar2 := Null
            strSql = strSql & "'" & Nvl(mrsInfo!����ID) & "," & mstr����IDs & "',"
            '  �������˷ѱ�_In Number := 0
            strSql = strSql & "" & IIf(mblnChangeFeeType, 1, 0) & ",Null,"
            '  ������������_In Number := 0
            strSql = strSql & "" & IIf(mblnUpdateAge, 1, 0) & ","
            '  �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null
            strSql = strSql & "'" & str����NO & "')"
            
            Call zlAddArray(cllPro, strSql)
            '����:31187:���ҺŻ��ܵ�������
            If Nvl(mrsPlan!�ű�) <> "" And k = 1 Then
                If Nvl(mrsPlan!ҽ������) = "" Then blnNoDoc = True
                strSql = "zl_���˹ҺŻ���_Update("
                '  ҽ������_In   �ҺŰ���.ҽ������%Type,
                strSql = strSql & IIf(blnNoDoc, "Null,", "'" & strҽ�� & "',")
                '  ҽ��id_In     �ҺŰ���.ҽ��id%Type,
                strSql = strSql & "" & IIf(blnNoDoc, "0,", ZVal(lngҽ��ID) & ",")
                '  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
                strSql = strSql & "" & Val(Nvl(rsItems!��ĿID)) & ","
                '  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
                strSql = strSql & "" & IIf(Val(Nvl(rsItems!ִ�п���ID)) = 0, lng�Һſ���ID, Val(Nvl(rsItems!ִ�п���ID))) & ","
                '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
                strSql = strSql & "" & str����ʱ�� & ","
                '  ԤԼ��־_In   Number := 0  --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����,3-�շ�ԤԼ
                strSql = strSql & IIf(mblnAppointment, IIf(mty_Para.blnԤԼʱ�տ�, 3, 1), 0) & ","
                '  ����_In       �ҺŰ���.����%Type := Null
                strSql = strSql & "'" & Nvl(mrsPlan!�ű�) & "',0,"
                strSql = strSql & "" & ZVal(Nvl(mrsPlan!��¼ID)) & ")"
                Call zlAddArray(cllProAfter, strSql)
            End If
            
            If mRegistFeeMode = EM_RG_���� And dblTotal <> 0 Then
                strSql = _
                "zl_���ﻮ�ۼ�¼_Insert('" & str����NO & "'," & k & "," & ZVal(Nvl(mrsInfo!����ID)) & ",NULL," & _
                         IIf(mstrClinic = "", "NULL", mstrClinic) & ",'" & str���ʽ & "'," & _
                         "'" & txtPatient.Text & "','" & mstrGender & "','" & mstrAge & "'," & _
                         "'" & mstrFeeType & "',NULL," & lng�Һſ���ID & "," & _
                         IIf(lng�Һſ���ID <> 0, lng�Һſ���ID, UserInfo.����ID) & ",'" & UserInfo.���� & "'," & IIf(rsItems!���� = 2, 1, "NULL") & "," & _
                         rsItems!��ĿID & ",'" & rsItems!��� & "','" & rsItems!���㵥λ & "'," & _
                         "NULL,1," & rsItems!���� & ",NULL," & IIf(rsItems!ִ�п���ID = 0, lng�Һſ���ID, rsItems!ִ�п���ID) & "," & IIf(int�۸񸸺� = k, "NULL", int�۸񸸺�) & "," & _
                         rsIncomes!������ĿID & ",'" & rsIncomes!�վݷ�Ŀ & "'," & rsIncomes!���� & "," & _
                         rsIncomes!Ӧ�� & "," & rsIncomes!ʵ�� & "," & str����ʱ�� & "," & str�Ǽ�ʱ�� & ",NULL,'" & UserInfo.���� & "','�Һ�:" & strNO & "')"
                Call zlAddArray(cllPro, strSql)
            End If
            k = k + 1
            rsIncomes.MoveNext
            Next j
        rsItems.MoveNext
    Next i
    
    If Not mblnAppointment Then
        If strҽ�� = UserInfo.���� Then
            strSql = "ZL_���˹Һż�¼_��������('" & strNO & "'," & Nvl(mrsInfo!����ID) & ",'" & lblRoomName.Caption & "','" & UserInfo.���� & "','','','" & zl_GetԤԼ��ʽByNo(strNO) & "')"    '�����:48350
            Call zlAddArray(cllPro, strSql)
            strSql = "zl_���˽���(" & Nvl(mrsInfo!����ID) & ",'" & strNO & "',NULL,'" & UserInfo.���� & "','" & lblRoomName.Caption & "')"
            Call zlAddArray(cllPro, strSql)
        End If
    End If
    
    Err = 0: On Error GoTo ErrFirt:
    
    If cllPro.Count > 0 Then
        Err = 0: On Error GoTo ErrFirt:
        zlExecuteProcedureArrAy cllPro, Me.Caption, True, False

        Err = 0: On Error GoTo errH:
        blnTrans = True
        If blnOneCard And lngҽ�ƿ����ID <> 0 And mRegistFeeMode = EM_RG_���� And cur�ֽ� <> 0 Then
            If Not mobjICCard.PaymentSwap(Val(cur�ֽ�), Val(cur�ֽ�), Val(lngҽ�ƿ����ID), 0, mstrCardNO, "", lng����ID, Nvl(mrsInfo!����ID)) Then
                gcnOracle.RollbackTrans
                MsgBox "һ��ͨ����Һŷ�ʧ��", vbInformation, gstrSysName
                Exit Function
            Else
                strSql = "zl_һ��ͨ����_Update(" & lng����ID & ",'" & cboPayMode.Text & "','" & mstrCardNO & "','" & lngҽ�ƿ����ID & "','" & "" & "'," & cur�ֽ� & ")"
                Call gobjDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
        End If

        'ҽ���Ķ�
        blnNotCommit = False
        If mintInsure <> 0 And mstrYBPati <> "" And cur���� <> 0 Then
            '68991:strAdvance:����ģʽ(0��1)|�Һŷ���ȡ��ʽ(0��1) |�Һŵ���
            strAdvance = ""
            If mRegistFeeMode = EM_RG_���� Or mPatiChargeMode = EM_�����ƺ���� Then
                strAdvance = IIf(mPatiChargeMode = EM_�����ƺ����, "1", "0")
                strAdvance = strAdvance & "|" & IIf(mRegistFeeMode = EM_RG_����, "1", "0")
                strAdvance = strAdvance & "|" & strNO
            End If
            If Not gclsInsure.RegistSwap(lng����ID, cur����, mintInsure, strAdvance) Then
                gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Function
            End If
            blnNotCommit = True
        End If
        '����:31187 ����ҽ���ɹ���,�����һЩ���ݸ���:�ڲ������������ύ���,���Բ�����д
        zlExecuteProcedureArrAy cllProAfter, Me.Caption, False, False
        Set cllCardPro = New Collection: Set cllTheeSwap = New Collection
        If mRegistFeeMode = EM_RG_���� And Not blnOneCard And Not mPatiChargeMode = EM_�����ƺ���� And cur�ֽ� <> 0 Then
            If zlInterfacePrayMoney(lng����ID, cllCardPro, cllTheeSwap, Val(cur�ֽ�), lngҽ�ƿ����ID, bln���ѿ�) = False Then
                gcnOracle.RollbackTrans: If cmdOK.Enabled = False Then cmdOK.Enabled = True
                Exit Function
            End If
            '������������
            zlExecuteProcedureArrAy cllCardPro, Me.Caption, False, False
        End If
        
        Err = 0: On Error GoTo OthersCommit:
        zlExecuteProcedureArrAy cllTheeSwap, Me.Caption, False, False
OthersCommit:
        gcnOracle.CommitTrans
        blnTrans = False
        On Error GoTo 0
        
        If mintInsure > 0 Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistSwap, True, mintInsure)
        '145198:���ϴ�,2019/12/26,�Һųɹ��������ҽӿڣ�Ŀǰ����ԤԼ�����֧����ά��
        Call zlSaveRgstAfterByPlugIn(mlngModul, strNO, (mblnAppointment And Not mty_Para.blnԤԼʱ�տ�))
    End If
    '��ӡ����
    If blnInvoicePrint Then
RePrint:
        If Not (mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��) And mRegistFeeMode = EM_RG_���� Then
            Dim blnEnterPrint As Boolean
            blnEnterPrint = True
            Load frmPrint
            Call frmPrint.ReportPrint(1, strNO, "", mlng����ID, mlng�Һ�ID, strFactNO, dat�Ǽ�ʱ��, , , , mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��, False, mstrUseType)
            If gblnBill�Һ� Then
                If zlIsNotSucceedPrintBill(4, strNO, strNotValiedNos) = True Then
                    If MsgBox("" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "����Ϊ[" & strNotValiedNos & "]Ʊ�ݴ�ӡδ�ɹ�,�Ƿ����½���Ʊ�ݴ�ӡ!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
                End If
            End If
        End If
    End If
    
    If blnSlipPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
    End If
    
    If blnAppointPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me, "NO=" & strNO, 2)
    End If
    
    If (blnSlipPrint Or blnInvoicePrint) And Not blnEnterPrint Then
        '��¼��ӡ��ƾ��
        gstrSQL = "Zl_ƾ����ӡ��¼_Update(4,'" & strNO & "',1,'" & UserInfo.���� & "')"
        gobjDatabase.ExecuteProcedure gstrSQL, ""
    End If
    mstrNO = strNO
    SaveData = True
    Exit Function
ErrFirt:
    gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    Exit Function
errH:
    If mintInsure > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistSwap, False, mintInsure)
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    Exit Function
ErrGo:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub ReloadPage()
    On Error GoTo errHandle
    Call LoadRegPlans(1)
    Call ClearPatient
    Call ClearRegInfo
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub
 

Private Function zlIsAllowPatiChargeFeeMode(ByVal lng����ID As Long, ByVal intԭ����ģʽ As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ�����ı䲡���շ�ģʽ
    '���:lng����ID-����ID
    '       intԭ����ģʽ-0��ʾ�Ƚ��������;1��ʾ�����ƺ����
    '����:��������շ�ģʽ,����true,���򷵻�False
    '����:���˺�
    '����:2013-12-25 10:06:49
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim dtDate As Date, intDay As Integer
    On Error GoTo errHandle
    
'    If mbytMode = 1 Then zlIsAllowPatiChargeFeeMode = True: Exit Function 'ԤԼ������
    'ģʽδ������ֱ�ӷ���true
    If intԭ����ģʽ = mPatiChargeMode Then zlIsAllowPatiChargeFeeMode = True: Exit Function
    
      
    If intԭ����ģʽ = 1 Then
        'ԭΪ�����ƺ�����Ҵ���δ����õ�,�������ü���ģʽ
        strSql = "" & _
        "   Select 1 " & _
        "   From ����δ����� " & _
        "   Where ����id = [1] And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
        If rsTemp.EOF = False Then
            MsgBox "ע��:" & vbCrLf & "  ��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�" & _
                                          vbCrLf & "����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ�" & _
                                          vbCrLf & "�ٹҺŻ򲻵������˵ľ���ģʽ", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        intDay = -1 * Val(Left(gobjDatabase.GetPara(21, glngSys, , "01") & "1", 1))
        dtDate = DateAdd("d", intDay, gobjDatabase.Currentdate)
        ' �ϴ�Ϊ"�����ƺ����",����Ϊ"�Ƚ��������"��,ͬʱ����δ����ҽ��ҵ�����ݵ� ,
        '   ��������ľ���ģʽ
        strSql = "Select 1 " & _
        " From ���˹Һż�¼ A, ����ҽ����¼ B " & _
        " Where a.����id + 0 = b.����id And a.No || '' = b.�Һŵ�  " & _
        "               And a.��¼״̬ = 1 And a.��¼���� = 1 And a.�Ǽ�ʱ�� - 0 >= [2] " & _
        "               And  a.����id = [1] And rownum<2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, dtDate)
        If rsTemp.EOF Then
            'δ����ҽ������
            MsgBox "ע��:" & vbCrLf & "  ��ǰ���˵ľ���ģʽΪ�����ƺ����," & vbCrLf & "  ����������ò��˵ľ���ģʽ!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    zlIsAllowPatiChargeFeeMode = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Sub ClearRegInfo()
    txtReg.Text = ""
    txtReg.Tag = ""
    lblDeptName.Caption = ""
    lblRoomName.Caption = ""
    cboRemark.Text = ""
    chkBook.Value = IIf(mty_Para.blnĬ�Ϲ�����, 1, 0)
    vsfMoney.Clear 1
    vsfMoney.Rows = 2
    lblTotal.Caption = "0.00"
    lblPayMoney.Caption = "0.00"
    txtPatient.SetFocus
    lbl��.Visible = False
End Sub

Private Function zlIsNotSucceedPrintBill(ByVal bytType As Byte, ByVal strNos As String, ByRef strOutValidNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���Ƿ��Ѿ�������ӡ
    '���:bytType-1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '       strNos-���δ�ӡƱ�ݵĵ���,�ö��ŷ���
    '����:strOutValidNos-��ӡʧ�ܵĵ��ݺ�
    '����:���ڲ��湦Ʊ�ݵĴ�ӡ,����true,���򷵻�False
    '����:���˺�
    '����:2012-01-16 18:06:01
    '����:44322,44326,44332,44330
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTempNos As String, rsTemp As ADODB.Recordset
    Dim strSql As String, strBillNos As String
    Dim bytBill As Byte
    On Error GoTo errHandle
    strBillNos = Replace(Replace(strNos, "'", ""), " ", "")
    strSql = "" & _
        "Select  /*+ rule */ distinct  B.NO " & _
        " From Ʊ��ʹ����ϸ A,Ʊ�ݴ�ӡ���� B,Table( f_Str2list([2])) J" & _
        " Where A.��ӡID =b.ID And B.��������=[1] And B.No=J.Column_value "
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "���Ʊ���Ƿ��ӡ", bytType, strBillNos)
    
    strTempNos = ""
    With rsTemp
        Do While Not .EOF
            If InStr(1, "," & strBillNos & ",", "," & !NO & ",") = 0 Then
                strTempNos = strTempNos & "," & !NO
            End If
            .MoveNext
        Loop
        If .RecordCount = 0 Then strTempNos = "," & strBillNos
    End With
    If strTempNos <> "" Then strTempNos = Mid(strTempNos, 2)
    rsTemp.Close: Set rsTemp = Nothing
    strOutValidNos = strTempNos
    zlIsNotSucceedPrintBill = strTempNos <> ""
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckValied() As Boolean
    Dim i As Integer
    '����ǰ���
    If mrsInfo Is Nothing Then
        MsgBox "�޷�ȷ��������Ϣ,����ѡ��һ�����ˣ�", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsInfo.RecordCount = 0 Then
        MsgBox "�޷�ȷ��������Ϣ,����ѡ��һ�����ˣ�", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsPlan Is Nothing Then
        MsgBox "�޷�ȷ���ű���Ϣ,����ѡ��һ���ű�", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsPlan.State = 0 Then
        MsgBox "�޷�ȷ���ű���Ϣ,����ѡ��һ���ű�", vbInformation, gstrSysName
        Exit Function
    End If
    If mrsPlan.RecordCount = 0 Then
        MsgBox "�޷�ȷ���ű���Ϣ,����ѡ��һ���ű�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mstrFeeType = "" Then
        MsgBox "���˷ѱ���Ϊ��,����ѡ��һ���ѱ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mRegistFeeMode <> EM_RG_���� Then
        If cboPayMode.Text = "" And cboPayMode.Visible And Val(lblTotal.Caption) <> 0 Then
            MsgBox "û��ȷ�����õĽ��㷽ʽ,������ɹҺ�!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If mblnAppointment And mty_Para.blnԤԼʱ�տ� = False Then
        If IsNull(mrsPlan!�Ű�) Then
            MsgBox "ԤԼ���տ�ģʽ��,���ܹҲ�����ĺű�!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If Nvl(mrsInfo!����) <> txtPatient.Text Then
        If MsgBox("��ǰ���������Ѿ������仯,�Ƿ����¶�ȡ������Ϣ?", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            Call GetPatient(IDKind.GetCurCard, txtPatient.Text, False)
            Exit Function
        Else
            txtPatient.Text = Nvl(mrsInfo!����)
        End If
    End If
    
    If InStr(gstrPrivs, ";�Һŷѱ����;") = 0 Then
        For i = 1 To vsfMoney.Rows - 1
            If Val(vsfMoney.TextMatrix(i, 2)) <> Val(vsfMoney.TextMatrix(i, 1)) Then
                MsgBox "��û��Ȩ�޸�����ʹ�ô��۷ѱ�,�������" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End If
    
    '���������
    If Not mrsItems Is Nothing Then
        mrsItems.MoveFirst
        Do While Not mrsItems.EOF
            If Val(Nvl(mrsItems!��ĿID)) <> 0 Then
                If CheckServeRange(0, Val(Nvl(mrsItems!��ĿID))) = False Then Exit Function
            End If
            mrsItems.MoveNext
        Loop
        mrsItems.MoveFirst
    End If
    
    CheckValied = True
End Function

Private Function CheckServeRange(intType As Integer, lng�շ�ϸĿID As Long, Optional intRow As Integer = 0) As Boolean
'����:����շ���Ŀ�ķ������,intType:0-�������;1-סԺ����
    Dim strSql As String, rsTmp As ADODB.Recordset
    strSql = "Select ����,Nvl(�������,0) As ������� From �շ���ĿĿ¼ Where ID = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "CheckServeRange", lng�շ�ϸĿID)
    If rsTmp.EOF Then
        MsgBox "����ȷ��" & IIf(intRow = 0, "", "��" & intRow & "��") & "�շ���Ŀ�ķ������,������Ŀ�Ƿ���ȷ¼��!", vbInformation, gstrSysName
        Exit Function
    Else
        Select Case intType
        Case 0
            If Val(rsTmp!�������) = 2 Or Val(rsTmp!�������) = 0 Then
                MsgBox IIf(intRow = 0, "", "��" & intRow & "��") & "�շ���Ŀ[" & rsTmp!���� & "]������������,����!", vbInformation, gstrSysName
                Exit Function
            End If
        Case 1
            If Val(rsTmp!�������) = 1 Or Val(rsTmp!�������) = 0 Then
                MsgBox IIf(intRow = 0, "", "��" & intRow & "��") & "�շ���Ŀ[" & rsTmp!���� & "]��������סԺ,����!", vbInformation, gstrSysName
                Exit Function
            End If
        Case Else
            If Val(rsTmp!�������) = 0 Then
                MsgBox IIf(intRow = 0, "", "��" & intRow & "��") & "�շ���Ŀ[" & rsTmp!���� & "]�������ڲ���,����!", vbInformation, gstrSysName
                Exit Function
            End If
        End Select
    End If
    CheckServeRange = True
End Function

Private Function zlInterfacePrayMoney(ByVal lng�ҺŽ���ID As Long, ByRef cllPro As Collection, _
    ByRef cllThreeSwap As Collection, dblMoney As Double, lngҽ�ƿ����ID As Long, bln���ѿ� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ӿ�֧�����
    '����:cllPro-�޸�������������
    '        cll��������-����������������
    '����:֧���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    
    If lngҽ�ƿ����ID = 0 Or dblMoney = 0 Then zlInterfacePrayMoney = True: Exit Function
    If cboPayMode.ItemData(cboPayMode.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln���ѿ� As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ��
    '���:frmMain-���õ�������
    '        lngModule-����ģ���
    '        strBalanceIDs-����ID,����ö��ŷ���
    '        strPrepayNos-��Ԥ��ʱ��Ч. Ԥ�����ݺ�,����ö��ŷ���
    '       strCardNo-����
    '       dblMoney-֧�����
    '����:strSwapGlideNO-������ˮ��
    '       strSwapMemo-����˵��
    '       strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '����:�ۿ�ɹ�,����true,���򷵻�Flase
    '˵��:
    '   ��������Ҫ�ۿ�ĵط����øýӿ�,Ŀǰ�滮��:�շ��ң��Һ���;������ѯ��;ҽ������վ��ҩ���ȡ�
    '   һ����˵���ɹ��ۿ�󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
    '   �ڿۿ�ɹ��󣬷��ؽ�����ˮ�ź���ر�ע˵���������������������Ϣ�����Է��ڽ���˵�����Ա��˷�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, lngҽ�ƿ����ID, bln���ѿ�, mstrCardNO, lng�ҺŽ���ID, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '����������������
     If lng�ҺŽ���ID <> 0 Then
        '����:58322
        'mbytMode As Integer '0-�Һ�,1-ԤԼ,2-����,3-ȡ��ԤԼ ,4-�˺� ԤԼ������ģʽ:0-�Һ�,��ʱԤԼҪ�շ�,1-ԤԼ,���շ�
        If Not bln���ѿ� Then
            '���ѿ��Ѿ��ڲ���Һż�¼ʱ,�Ѿ��ۿ�
            Call zlAddUpdateSwapSQL(False, lng�ҺŽ���ID, lngҽ�ƿ����ID, bln���ѿ�, mstrCardNO, strSwapGlideNO, strSwapMemo, cllPro)
        End If
        Call zlAddThreeSwapSQLToCollection(False, lng�ҺŽ���ID, lngҽ�ƿ����ID, bln���ѿ�, mstrCardNO, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlAddThreeSwapSQLToCollection(ByVal blnԤ���� As Boolean, _
    ByVal strIDs As String, ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    ByVal str���� As String, strExpend As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������������
    '���: blnԤ����-�Ƿ�Ԥ����
    '       lngID-�����Ԥ����,����Ԥ��ID,�������ID
    ' ����:cllPro-����SQL��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSql As String, varData As Variant, varTemp As Variant, i As Long
     
    Err = 0: On Error GoTo Errhand:
    '���ύ,�����������,�ٸ�����صĽ�����Ϣ
    'strExpend:������չ��Ϣ,��ʽ:��Ŀ����|��Ŀ����||...
    varData = Split(strExpend, "||")
    Dim str������Ϣ As String, strTemp As String
    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            If varTemp(0) <> "" Then
                strTemp = varTemp(0) & "|" & varTemp(1)
                If gobjCommFun.ActualLen(str������Ϣ & "||" & strTemp) > 2000 Then
                    str������Ϣ = Mid(str������Ϣ, 3)
                    'Zl_�������㽻��_Insert
                    strSql = "Zl_�������㽻��_Insert("
                    '�����id_In ����Ԥ����¼.�����id%Type,
                    strSql = strSql & "" & lng�����ID & ","
                    '���ѿ�_In   Number,
                    strSql = strSql & "" & IIf(bln���ѿ�, 1, 0) & ","
                    '����_In     ����Ԥ����¼.����%Type,
                    strSql = strSql & "'" & str���� & "',"
                    '����ids_In  Varchar2,
                    strSql = strSql & "'" & strIDs & "',"
                    '������Ϣ_In Varchar2:������Ŀ|��������||...
                    strSql = strSql & "'" & str������Ϣ & "',"
                    'Ԥ����ɿ�_In Number := 0
                    strSql = strSql & IIf(blnԤ����, "1", "0") & ")"
                    zlAddArray cllPro, strSql
                    str������Ϣ = ""
                End If
                str������Ϣ = str������Ϣ & "||" & strTemp
            End If
        End If
    Next
    If str������Ϣ <> "" Then
        str������Ϣ = Mid(str������Ϣ, 3)
        'Zl_�������㽻��_Insert
        strSql = "Zl_�������㽻��_Insert("
        '�����id_In ����Ԥ����¼.�����id%Type,
        strSql = strSql & "" & lng�����ID & ","
        '���ѿ�_In   Number,
        strSql = strSql & "" & IIf(bln���ѿ�, 1, 0) & ","
        '����_In     ����Ԥ����¼.����%Type,
        strSql = strSql & "'" & str���� & "',"
        '����ids_In  Varchar2,
        strSql = strSql & "'" & strIDs & "',"
        '������Ϣ_In Varchar2:������Ŀ|��������||...
        strSql = strSql & "'" & str������Ϣ & "',"
        'Ԥ����ɿ�_In Number := 0
        strSql = strSql & IIf(blnԤ����, "1", "0") & ")"
        zlAddArray cllPro, strSql
    End If
    zlAddThreeSwapSQLToCollection = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlAddUpdateSwapSQL(ByVal blnԤ�� As Boolean, ByVal strIDs As String, ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    str���� As String, str������ˮ�� As String, str����˵�� As String, _
    ByRef cllPro As Collection, Optional intУ�Ա�־ As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������������ˮ�ź���ˮ˵��
    '���: blnԤ����-�Ƿ�Ԥ����
    '       lngID-�����Ԥ����,����Ԥ��ID,�������ID
    '����:cllPro-����SQL��
    '����:���˺�
    '����:2011-07-27 10:13:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    strSql = "Zl_�����ӿڸ���_Update("
    '  �����id_In   ����Ԥ����¼.�����id%Type,
    strSql = strSql & "" & lng�����ID & ","
    '  ���ѿ�_In     Number,
    strSql = strSql & "" & IIf(bln���ѿ�, 1, 0) & ","
    '  ����_In       ����Ԥ����¼.����%Type,
    strSql = strSql & "'" & str���� & "',"
    '  ����ids_In    Varchar2,
    strSql = strSql & "'" & strIDs & "',"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    strSql = strSql & "'" & str������ˮ�� & "',"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type
    strSql = strSql & "'" & str����˵�� & "',"
    'Ԥ����ɿ�_In Number := 0
    strSql = strSql & "" & IIf(blnԤ��, 1, 0) & ","
    '�˷ѱ�־ :1-�˷�;0-����
    strSql = strSql & "0,"
    'У�Ա�־
    strSql = strSql & "" & IIf(intУ�Ա�־ = 0, "NULL", intУ�Ա�־) & ")"
    zlAddArray cllPro, strSql
End Function

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub SetControl()
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strTemp As String, i As Integer
    If mblnAppointment Then
        lblRoom.Visible = False
        picRoom.Visible = False
        lblDept.Left = lblRoom.Left
        picDept.Left = picRoom.Left
        picDept.Width = picRoom.Width
        chkBook.Value = 0
        chkBook.Visible = False
        cboRemark.Width = 7170
        If mty_Para.blnԤԼʱ�տ� Then
            lblPayMode.Visible = True
            cboPayMode.Visible = True
            picPayMoney.Visible = True
            cmdPrice.Visible = mty_Para.byt�Һ�ģʽ = 2
        Else
            lblPayMode.Visible = False
            cboPayMode.Visible = False
            picPayMoney.Visible = False
            cmdPrice.Visible = False
        End If
        cboAppointStyle.Clear
        strSql = "Select ����,ȱʡ��־ From ԤԼ��ʽ"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)
        Do While Not rsTmp.EOF
            cboAppointStyle.AddItem Nvl(rsTmp!����)
            If Val(Nvl(rsTmp!ȱʡ��־)) = 1 Then cboAppointStyle.ListIndex = cboAppointStyle.NewIndex
            rsTmp.MoveNext
        Loop
        strTemp = gobjDatabase.GetPara("ȱʡԤԼ��ʽ", glngSys, 9000, "")
        For i = 0 To cboAppointStyle.ListCount - 1
            If cboAppointStyle.List(i) = strTemp Then
                cboAppointStyle.ListIndex = i
            End If
        Next i
    Else
        lblDate.Visible = False
        lblTime.Visible = False
        dtpDate.Visible = False
        dtpTime.Visible = False
        cmdTime.Visible = False
        If (mty_Para.byt�Һ�ģʽ = 0 Or mty_Para.byt�Һ�ģʽ = 2) And gSysPara.bln��Һ�ģʽ = False Then
            lblPayMode.Visible = True
            cboPayMode.Visible = True
            picPayMoney.Visible = True
            cmdPrice.Visible = mty_Para.byt�Һ�ģʽ = 2
        Else
            lblPayMode.Visible = False
            cboPayMode.Visible = False
            picPayMoney.Visible = False
            cmdPrice.Visible = False
        End If
    End If
End Sub

Private Sub Form_Load()
    Err = 0
    mblnInit = True
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    lblSn.Caption = ""
    If mblnAppointment Then
        Me.Caption = "ҽ��վԤԼ"
        lblAppointStyle.Visible = True
        cboAppointStyle.Visible = True
    Else
        Me.Caption = "ҽ��վ" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & ""
        lblAppointStyle.Visible = False
        cboAppointStyle.Visible = False
    End If
    
    gobjDatabase.ExecuteProcedure "zl1_auto_buildingregisterplan", Me.Caption
    Call Init�ѱ�
    Call InitPara
    chkBook.Value = IIf(mty_Para.blnĬ�Ϲ�����, 1, 0)
    Call InitIDKind
    Call InitAppointmentTime
    Call GetAllҽ��
    chkAll.Value = IIf(mty_Para.blnShowAllPlan, 1, 0)
    If LoadRegPlans(1) = False Then
        mblnUnload = True
    End If
    Call LoadPayMode
    Call SetControl
    If mblnAppointment And mlng����ID <> 0 And mblnUnload = False Then
        Call GetPatient(IDKind.GetCurCard, "-" & mlng����ID, False)
    End If
    cmdNewPati.Enabled = InStr(gstrPrivs, ";�ҺŲ��˽���;") > 0
    cmdOther.Enabled = InStr(gstrPrivs, ";���������ҽ���ĺ�Դ;") > 0
    '137272:���ϴ�,2019/2/20,��ֹ���ź�ϵͳ������������
    Call CancelRegNo
End Sub

Private Sub InitAppointmentTime()
    '��ʼ��ԤԼʱ��
    Dim rsDay As ADODB.Recordset, strSql As String
    Dim intԤԼ���� As Integer
    Dim dtNow As Date
  
    On Error GoTo ErrHandler
    intԤԼ���� = mintSysAppLimit
    If mblnAppointment Then
        Call mobjRegister.zlGetRegisterMaxDaysFromDeptAndDoctor_Visits( _
            gstrDeptIDs, UserInfo.����, mty_Para.blnԤԼ�������Ұ���, intԤԼ����)
    End If
    
    dtNow = gobjDatabase.Currentdate
    dtpDate.MaxDate = Format(dtNow + intԤԼ����, "yyyy-mm-dd")
    dtpDate.minDate = Format(dtNow, "yyyy-mm-dd")
    dtpTime.Value = Format(dtNow, "hh:mm:ss")
    
    strSql = _
        "Select Nvl(Min(a.��������), Trunc(Sysdate) + 1) As ��������" & vbNewLine & _
        "From �ٴ������¼ A" & vbNewLine & _
        "Where a.�������� > Trunc(Sysdate) And a.����id = [1] And a.ҽ��id = [2]"
    Set rsDay = gobjDatabase.OpenSQLRecord(strSql, "", UserInfo.����ID, UserInfo.ID)
    dtpDate.Value = Format(rsDay!��������, "yyyy-mm-dd")
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtPatient.Text = objPatiInfor.����
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
End Sub

Private Sub txtPatient_Change()
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub txtPatient_GotFocus()
    Call gobjControl.TxtSelAll(txtPatient)
    Call gobjCommFun.OpenIme(True)
    If txtPatient.Text = "" And ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub zlInusreIdentify()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ҽ������鿨
    '���ƣ����˺�
    '���ڣ�2010-07-14 11:32:08
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    Dim str�������� As String
    Dim rsTmp As ADODB.Recordset
    Dim cur��� As Currency
    Dim curMoney As Currency
    Dim blnDeposit As Boolean, blnInsure As Boolean
    If mrsInfo Is Nothing Then
        lng����ID = 0
        str�������� = ""
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
        str�������� = Nvl(mrsInfo!��������)
    End If

    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard False

    Dim strAdvance As String    '����ģʽ(0-�Ƚ�������ƻ�1-�����ƺ����)|�Һŷ���ȡ��ʽ(0-���ջ�1-����)
    Dim varData As Variant
    mstrYBPati = gclsInsure.Identify(3, lng����ID, mintInsure, strAdvance)
    
    If mblnAppointment Then
        mRegistFeeMode = EM_RG_����
    Else
        If (mty_Para.byt�Һ�ģʽ = 0 Or mty_Para.byt�Һ�ģʽ = 2) And gSysPara.bln��Һ�ģʽ = False Then
            mRegistFeeMode = EM_RG_����
            picPayMoney.Visible = True
            cboPayMode.Visible = True
            lblPayMode.Visible = True
            cmdPrice.Visible = mty_Para.byt�Һ�ģʽ = 2
        Else
            mRegistFeeMode = EM_RG_����
            picPayMoney.Visible = False
            cboPayMode.Visible = False
            lblPayMode.Visible = False
            cmdPrice.Visible = False
        End If
    End If
    
    mPatiChargeMode = EM_�Ƚ��������
    If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard True
    End If
    
    If mstrYBPati = "" Then
        If Not txtPatient.Enabled Then txtPatient.Enabled = True
         mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
         Exit Sub
    End If
    
    '�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
    If UBound(Split(mstrYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mstrYBPati, ";")(8)) Then lng����ID = Val(Split(mstrYBPati, ";")(8))
    End If
        
    If lng����ID = 0 Then
        mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
        Exit Sub
    End If
    MCPAR.�Һż����Ŀ = gclsInsure.GetCapability(support�Һż����Ŀ, lng����ID, mintInsure)
    txtPatient.Text = "-" & lng����ID
    Call txtPatient_Validate(False)    '���е�Setfocus����ʹ���¼�(txtPatient_KeyPress)ִ�����,�����ٴ��Զ�ִ��txtPatient_Validate
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False)
    Call SetPatiColor(txtPatient, str��������, vbRed)
    txtPatient.BackColor = &HE0E0E0
    txtPatient.Locked = True

    If strAdvance <> "" Then
        varData = Split(strAdvance & "|", "|")
        mPatiChargeMode = IIf(Val(varData(0)) = 1, EM_�����ƺ����, EM_�Ƚ��������)
        mRegistFeeMode = IIf(Val(varData(1)) = 1, EM_RG_����, EM_RG_����)
    End If
    MCPAR.���ղ����� = gclsInsure.GetCapability(support�ҺŲ���ȡ������, lng����ID, mintInsure)
    MCPAR.ҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, mintInsure)
    mlng����ID = 0
    curMoney = GetRegistMoney
    Set rsTmp = GetMoneyInfoRegist(lng����ID, , , 1)
    cur��� = 0
    Do While Not rsTmp.EOF
        cur��� = cur��� + Val(Nvl(rsTmp!Ԥ�����))
        cur��� = cur��� - Val(Nvl(rsTmp!�������))
        rsTmp.MoveNext
    Loop
    If cur��� > 0 Then
        lblMoney.Caption = "����Ԥ�����:" & Format(cur���, "0.00")
        If cur��� >= curMoney Then
            blnDeposit = True
        Else
            blnDeposit = False
        End If
    End If
    
    mcur������� = gclsInsure.SelfBalance(lng����ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur����͸֧, mintInsure)
    lblMoney.Caption = lblMoney.Caption & "/�����ʻ����:" & Format(mcur�������, "0.00")
    If gclsInsure.GetCapability(support�Һ�ʹ�ø����ʻ�, lng����ID, mintInsure) = False Then
        blnInsure = False
    Else
        If mcur������� + mcur����͸֧ >= curMoney Then
            blnInsure = True
        Else
            blnInsure = False
        End If
    End If
    Call LoadPayMode(blnDeposit, blnInsure)
    If mRegistFeeMode = EM_RG_���� Then
        lblSum.Caption = "����"
        picPayMoney.Visible = False
        cboPayMode.Visible = False
        lblPayMode.Visible = False
    Else
        lblSum.Caption = "�ϼ�"
    End If
    If mRegistFeeMode = EM_RG_���� Then
        If mblnAppointment Then
            mRegistFeeMode = EM_RG_����
        Else
            If (mty_Para.byt�Һ�ģʽ = 0 Or mty_Para.byt�Һ�ģʽ = 2) And gSysPara.bln��Һ�ģʽ = False Then
                mRegistFeeMode = EM_RG_����
                picPayMoney.Visible = True
                cboPayMode.Visible = True
                lblPayMode.Visible = True
                cmdPrice.Visible = mty_Para.byt�Һ�ģʽ = 2
            Else
                mRegistFeeMode = EM_RG_����
                picPayMoney.Visible = False
                cboPayMode.Visible = False
                lblPayMode.Visible = False
                cmdPrice.Visible = False
            End If
        End If
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    '0-�����,1-����,2-�Һŵ�,3-���￨��,4-ҽ����
    Dim blnCard As Boolean
    Dim strKind As String, intLen As Integer
    Static sngBegin As Single
    Dim sngNow As Single
    'ҽ����֤
    If txtPatient.Text = "" And KeyAscii = 13 Then
        KeyAscii = 0
        Call zlInusreIdentify
    End If
    
    If KeyAscii <> 0 And KeyAscii > 32 And mty_Para.bln�Һű���ˢ�� Then
        sngNow = Timer
        If txtPatient.Text = "" Then
            sngBegin = sngNow
        ElseIf Format((sngNow - sngBegin) / (Len(txtPatient.Text) + 1), "0.000") >= 0.04 Then    '>0.007>=0.01
            txtPatient.Text = Chr(KeyAscii)
            txtPatient.SelStart = 1
            KeyAscii = 0
            sngBegin = sngNow
        End If
    End If
    
    strKind = IDKind.GetCurCard.����
    txtPatient.PasswordChar = IIf(IDKind.GetCurCard.�������Ĺ��� <> "", "*", "")
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    
    'ȡȱʡ��ˢ����ʽ
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
            '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
            '��7λ��,��ֻ��������,��Ȼȡ������
    Select Case strKind
    Case "����"
        blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, gobjSquare.blnȱʡ��������)
        intLen = gobjSquare.intȱʡ���ų���
    Case "�����"
        If InStr("0123456789-" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "���֤"
    Case "ҽ����"
    Case Else
            If IDKind.GetCurCard.�ӿ���� <> 0 Then
                blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.GetCurCard.�������Ĺ��� <> "")
                intLen = IDKind.GetCurCard.���ų���
            End If
    End Select
    
    'ˢ����ϻ���������س�
    If (blnCard And Len(txtPatient.Text) = intLen - 1 And KeyAscii <> 8) Or (KeyAscii = 13) Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0: mblnCard = True
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), blnCard)
        mblnCard = False
        gobjControl.TxtSelAll txtPatient
   End If
End Sub

Private Function CheckNoValied(ByVal lngRow As Long) As Boolean
    CheckNoValied = True
End Function

Private Function zl_GetԤԼ��ʽByNo(strNO As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:���ݹҺŵ��ݺŻ�ȡ����ԤԼ��ʽ
    '���:strNo-�Һŵ��ݺ�
    '����:ԤԼ��ʽ
    '����:����
    '����:2012-07-03
    '�����:48350
    '-----------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim strԤԼ��ʽ As String
    Dim rsTemp As Recordset
    strSql = "" & _
        "Select ԤԼ��ʽ From ���˹Һż�¼ Where ��¼״̬=1 And No=[1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "��ȡԤԼ��ʽ", strNO)
    If rsTemp Is Nothing Then zl_GetԤԼ��ʽByNo = "": Exit Function
    If rsTemp.RecordCount = 0 Then zl_GetԤԼ��ʽByNo = "": Exit Function
    While rsTemp.EOF = False
        strԤԼ��ʽ = Nvl(rsTemp!ԤԼ��ʽ)
        rsTemp.MoveNext
    Wend
    zl_GetԤԼ��ʽByNo = strԤԼ��ʽ
End Function

Public Function GetʧԼ��(ByVal str�ű� As String, ByVal datThis As Date) As Long
   '��ȡ������ĳһ��.ԤԼʧԼ��
    Dim strSql  As String
    Dim rsTmp   As ADODB.Recordset
    Dim strDat  As String
'    If mty_Para.blnʧԼ���ڹҺ� = False Or mty_Para.lngԤԼ��Чʱ�� <= 0 Then Exit Function
    strSql = "                " & " SELECT count(1) AS ʧԼ�� "
    strSql = strSql & vbNewLine & " FROM �Һ����״̬ "
    strSql = strSql & vbNewLine & " WHERE ����=[1] AND ״̬=2 AND ����-[3]/24/60 <SYSDATE AND To_Char(����,'yyyy-MM-dd')=[2]"
    strDat = Format(datThis, "yyyy-MM-dd")
    On Error GoTo Hd
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, str�ű�, strDat, mty_Para.lngԤԼ��Чʱ��)
    If rsTmp.EOF Then
        GetʧԼ�� = 0
        Set rsTmp = Nothing
        Exit Function
    End If
    GetʧԼ�� = Val(Nvl(rsTmp!ʧԼ��, 0))
    Set rsTmp = Nothing
   Exit Function
Hd:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub txtPatient_LostFocus()
    Call gobjCommFun.OpenIme
    IDKind.SetAutoReadCard False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    txtPatient.Text = Trim(txtPatient.Text)
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False)
'        gobjControl.TxtSelAll txtPatient
'    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIndex As Long
    Dim strSql As String
    Dim lng����ID As Long
    Dim strAge As String, strBirth As String
    Dim str�ѱ� As String, str���ʽ As String
    
    On Error GoTo errH
    If txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        txtPatient.Text = strID:
        If txtPatient.Text = "" Then
            Call mobjIDCard.SetEnabled(False) '��������Ϸ������������ü����Զ���ȡ
            Exit Sub
        End If
        lngPreIndex = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("���֤��")
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True, True, , True)
        IDKind.IDKind = lngPreIndex
        'δ�ҵ�����,�Զ�����
        If txtPatient.Text = "" And strName <> "" Then
            If InStr(gstrPrivs, ";�ҺŲ��˽���;") > 0 Then
                If MsgBox("δ�ҵ����֤��Ϊ[" & strID & "]�Ĳ���,�Ƿ��Զ�����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    lng����ID = gobjDatabase.GetNextNo(1)
                    If IsDate(datBirthDay) Then
                        strAge = ReCalcOld(datBirthDay, , , False)
                        strBirth = "To_Date('" & Format(datBirthDay, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                    Else
                        strBirth = "Null"
                    End If
                    str�ѱ� = GetDefault("�ѱ�")
                    str���ʽ = GetDefault("ҽ�Ƹ��ʽ")
                    strSql = "ZL_�ҺŲ��˲���_INSERT(1," & lng����ID & ",Null,Null,Null,'" & strName & "','" & strSex & "','" & strAge & "','" & _
                            str�ѱ� & "','" & str���ʽ & "','" & strNation & "',Null,Null,Null,'" & strID & "',Null,Null,Null,Null,'" & strAddress & "',Null," & _
                            "Null,Sysdate,Null," & strBirth & ",Null,Null,Null,Null,'" & strAddress & "')"
                    gobjDatabase.ExecuteProcedure strSql, Me.Caption
                    txtPatient.Text = "-" & lng����ID
                    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False, , , True)
                End If
            Else
                MsgBox "δ�ҵ����֤��Ϊ[" & strID & "]�Ĳ���,���ܼ���!", vbInformation, gstrSysName
            End If
        End If
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Function GetDefault(strItem As String) As String
    Dim strSql As String, rsTmp As ADODB.Recordset
    strSql = "Select ���� From " & strItem & " Where ȱʡ��־ = 1"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)
    If Not rsTmp.EOF Then GetDefault = Nvl(rsTmp!����)
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
        Case vbKeyF3
            If txtPatient.Visible = True And txtPatient.Enabled Then
                Call txtPatient.SetFocus
            End If
        Case vbKeyF4
            If cmdNewPati.Enabled And cmdNewPati.Visible Then Call cmdNewPati_Click
        Case Else
            IDKind.ActiveFastKey
    End Select
End Sub

Public Sub ActiveIDKindKey()
    IDKind.ActiveFastKey
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIndex As Long
    If txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        txtPatient.Text = strNO
        If txtPatient.Text = "" Then
            Call mobjICCard.SetEnabled(False) '��������Ϸ������������ü����Զ���ȡ
            Exit Sub
        End If
        lngPreIndex = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("IC����")
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True, True)
        IDKind.IDKind = lngPreIndex
    End If
End Sub

Private Sub GetPatient(objCard As zlIDKind.Card, ByVal strInput As String, ByVal blnCard As Boolean, _
                    Optional blnInputIDCard As Boolean = False, Optional ByRef Cancel As Boolean, _
                    Optional blnNoPrompt As Boolean = False)
    '���ܣ���ȡ������Ϣ
    '������blnCard=�Ƿ���￨ˢ��
    '
    '         blnInputIDCard-�Ƿ����֤ˢ��
    '����:Cancel-Ϊtrue��ʾ���صķ�����ȡ������Ϣ
    Dim strSql As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur��� As Currency, curMoney As Currency
    Dim strInputInfo As String '���洫��������ı� ������ʹ�����֤�� �Բ��˽��в��Һ� ���滻��"-" ����ID�����
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str����Ժ As String
    Dim blnҽ���� As Boolean, rsFeeType As ADODB.Recordset
    Dim IntMsg As VbMsgBoxResult
    Dim blnOtherType As Boolean '�Ƿ������

    strInputInfo = strInput
    
    On Error GoTo errH
    blnҽ���� = False
    
    If objCard Is Nothing Then Set objCard = IDKind.GetCurCard

    strSql = "Select  A.����ID,A.�����,A.סԺ��,A.���￨��,A.�ѱ�,A.ҽ�Ƹ��ʽ,A.����,A.�Ա�,A.����,A.��������,A.�����ص�,A.���֤��,A.����֤��,A.���,A.ְҵ,A.����,A.��������, " & _
             "A.����,A.����,A.����,A.ѧ��,A.����״��,A.��ͥ��ַ,A.��ͥ�绰,A.��ͥ��ַ�ʱ�,A.�໤��,A.��ϵ������,A.��ϵ�˹�ϵ,A.��ϵ�˵�ַ,A.��ϵ�˵绰,A.���ڵ�ַ, " & _
             "A.���ڵ�ַ�ʱ�,A.Email,A.QQ,A.��ͬ��λid,A.������λ,A.��λ�绰,A.��λ�ʱ�,A.��λ������,A.��λ�ʺ�,A.������,A.������,A.��������,A.����ʱ��,A.����״̬, " & _
             "A.��������,A.סԺ����,A.��ǰ����id,A.��ǰ����id,A.��ǰ����,A.��Ժʱ��,A.��Ժʱ��,A.��Ժ,A.IC����,A.������,A.ҽ����,A.����,A.��ѯ����,A.�Ǽ�ʱ��,A.ͣ��ʱ��,A.����,A.��ϵ�����֤��, " & _
             "B.���� ��������,A.��ѯ���� As ����֤��,A.����ģʽ,a.��ҳID From ������Ϣ A,������� B  Where A.���� = B.���(+) And A.ͣ��ʱ�� is NULL "

    If mty_Para.blnסԺ���˹Һ� = False Then
        str����Ժ = " And Not Exists(Select 1 From ������ҳ Where ����ID=A.����ID   And ��ҳID<>0 And ��ҳID=A.��ҳID And Nvl(��������,0)=0 And ��Ժ���� is Null)"
    End If
   
    If blnCard And objCard.���� Like "����*" And mstrYBPati = "" And InStr("-+*.", Left(strInput, 1)) = 0 Then     'ˢ��
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        ElseIf IDKind.GetCurCard.�ӿ���� > 0 Then
            lng�����ID = IDKind.GetCurCard.�ӿ����
        Else
            lng�����ID = -1
        End If
        
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        If IDKind.IsMobileNo(strInput) And lng����ID = 0 Then
            If gobjSquare.objSquareCard.zlGetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        End If
        If lng����ID <= 0 Then GoTo NewPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSql = strSql & " And A.����ID=[2] " & str����Ժ
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '�����
        strSql = strSql & " And A.�����=[2]" & str����Ժ
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '����ID
        strSql = strSql & " And A.����ID=[2]" & _
        IIf(mstrYBPati <> "", "", str����Ժ)
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then
        'סԺ��
        strSql = strSql & " And A.����ID=(Select Max(����ID) As ����ID From ������ҳ Where סԺ�� = [2])" & str����Ժ
    ElseIf blnInputIDCard Then  '���������֤ʶ��
        strInput = UCase(strInput)
        If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        strInput = "-" & lng����ID
        strSql = strSql & " And A.����ID=[2] " & str����Ժ
    ElseIf objCard.���� Like "����*" And IDKind.IsMobileNo(strInput) = True Then
        If gobjSquare.objSquareCard.zlGetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then Exit Sub
        strInput = "-" & lng����ID
        strSql = strSql & " And A.����ID=[2] " & str����Ժ
    Else
        Select Case objCard.����
            Case "����", "��������￨"
                If Not mty_Para.bln����ģ������ Or mty_Para.bln����ģ������ And Len(txtPatient.Text) < 2 Then
                    Set mrsInfo = Nothing: Exit Sub
                End If
                strPati = _
                    " Select distinct 1 as ����ID,A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����,A.�����,A.��������,A.���֤��,A.��ͥ��ַ,A.������λ" & _
                    " From ������Ϣ A " & _
                    " Where Rownum <101 And A.ͣ��ʱ�� is NULL And A.���� Like [1]" & str����Ժ & _
                    IIf(mty_Para.lng������������ = 0, "", " And Nvl(A.����ʱ��,A.�Ǽ�ʱ��)>Trunc(Sysdate-[2])")
                    
'                strPati = strPati & " Union ALL " & _
'                        "Select 0,0 as ID,-NULL,'[�²���]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL From Dual"
                strPati = strPati & " Order by ����ID,����"
                    
                vRect = GetControlRect(txtPatient.hWnd)
                Set rsTmp = gobjDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", mty_Para.lng������������)
                If Not rsTmp Is Nothing Then
                    If rsTmp!ID = 0 Then '�����²���
                        txtPatient.Text = ""
                        If Not blnNoPrompt Then MsgBox "û���ҵ���Ӧ�Ĳ�����Ϣ������������Ϣ�Ƿ���ȷ���߲����Ƿ񽨵���", vbInformation, gstrSysName
                        Set mrsInfo = Nothing: Exit Sub
                    Else '�Բ���ID��ȡ
                        strInput = rsTmp!����ID
                        strSql = strSql & " And A.����ID=[1]"
                    End If
                Else 'ȡ��ѡ��
                    txtPatient.Text = ""
                    Set mrsInfo = Nothing: Exit Sub
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                blnҽ���� = True
                If mblnOlnyBJYB And gobjCommFun.ActualLen(strInput) >= 9 Then
                    strSql = strSql & " And A.ҽ���� like [3] " & str����Ժ
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSql = strSql & " And A.ҽ����=[1]" & str����Ժ
                End If
            Case "�ֻ���"
                If IDKind.IsMobileNo(strInput) = False Then Exit Sub
                If gobjSquare.objSquareCard.zlGetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then Exit Sub
                strInput = "-" & lng����ID
                strSql = strSql & " And A.����ID=[2] " & str����Ժ
            Case "���֤��", "���֤", "�������֤"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strSql = strSql & " And A.����ID=[2] " & str����Ժ
                 
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strSql = strSql & " And A.����ID=[2] " & str����Ժ
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.�����=[1]" & str����Ժ
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.����ID=(Select Max(����ID) As ����ID From ������ҳ Where סԺ�� = [1])" & str����Ժ
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                    blnOtherType = True
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                strSql = strSql & " And A.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
ReadPati:
    If Mid(mstrCardPass, 1, 1) = "1" And strPassWord <> "" Then
        If Not gobjCommFun.VerifyPassWord(Me, "" & strPassWord) Then
            MsgBox "���������֤ʧ�ܣ�", vbInformation, gstrSysName
            ClearPatient
            Exit Sub
        End If
    End If
    Set mrsInfo = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, strInput, Mid(strInput, 2), strTemp)
    strInput = strInputInfo
    If Not mrsInfo.EOF Then
        txtPatient.Text = Nvl(mrsInfo!����) '�����Change�¼�
        txtPatient.BackColor = &H80000005
        lblSum.Caption = "�ϼ�"
        Call SetControl
        '�ڵ���txtPatient_Change�¼���������źͲ���������Ϊ�յ������ �޷�ʶ��ò�����Ϣ ���ִ���
        '���������ݿ����ݴ����ٽ��к����Ĵ���
        If mrsInfo Is Nothing Then Cancel = True: Exit Sub
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(Trim(mstr����) = "", txtPatient.ForeColor, vbRed))
        
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!����֤��)
        mstrGender = Nvl(mrsInfo!�Ա�)
        txtPatient.PasswordChar = ""
        
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        If Load�ѱ�(Nvl(mrsInfo!�ѱ�)) = False Then mstrFeeType = ""
        
        mstrAge = Nvl(mrsInfo!����)
        mblnUpdateAge = False
        If Not IsNull(mrsInfo!��������) Then
            strSql = "Select Zl_Age_Calc([1],[2],Null) As Old From Dual"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, lng����ID, CDate(mrsInfo!��������))
            If mstrAge <> Nvl(rsTmp!old) And Nvl(rsTmp!old) <> "" Then
                mblnUpdateAge = True
                mstrAge = Nvl(rsTmp!old)
            End If
        End If
        mstrClinic = Nvl(mrsInfo!�����)
        If mstrClinic = "" Then
            mstrClinic = gobjDatabase.GetNextNo(3)
            mblnChangeFeeType = True
        Else
            mblnChangeFeeType = False
        End If
        
        lblInfo.Caption = "�Ա�:" & mstrGender & "   ����:" & mstrAge & "   �����:" & mstrClinic & "   �ѱ�:" & mstrFeeType
        
        '����Ԥ������Ϣ
        Set rsTmp = GetMoneyInfoRegist(mrsInfo!����ID, , , 1)
        cur��� = 0
        Do While Not rsTmp.EOF
            cur��� = cur��� + Val(Nvl(rsTmp!Ԥ�����))
            cur��� = cur��� - Val(Nvl(rsTmp!�������))
            rsTmp.MoveNext
        Loop
        If cur��� > 0 Then
            lblMoney.Caption = "����Ԥ�����:" & Format(cur���, "0.00")
            curMoney = GetRegistMoney
            If cur��� >= curMoney Then
                Call LoadPayMode(True)
            Else
                Call LoadPayMode
            End If
        Else
            lblMoney.Caption = "����Ԥ�����:0.00"
            Call LoadPayMode
        End If
        
        Call ResetDefault����
        
        '���ݲ������¶�ȡ��Ŀ����
        If mintPriceGradeStartType >= 2 Then
            Call GetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID)), Nvl(mrsInfo!ҽ�Ƹ��ʽ, 0), , , mstrPriceGrade)
        End If
        
        If Not mrsPlan Is Nothing Then
            If Not mrsPlan.EOF Then Call LoadFeeItem(Val(Nvl(mrsPlan!��ĿID)), chkBook.Value = 1, mstrPriceGrade)
        End If
        cmdNewPati.ToolTipText = "��ϸ��Ϣ(F4)"
        cmdNewPati.Enabled = True
        If txtReg.Enabled And txtReg.Visible Then txtReg.SetFocus
    Else
NewPati:
        If Not blnNoPrompt Then MsgBox "û���ҵ���Ӧ�Ĳ�����Ϣ������������Ϣ�Ƿ���ȷ���߲����Ƿ񽨵���", vbInformation, gstrSysName
        ClearPatient
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub ClearPatient()
    txtPatient.Text = ""
    txtPatient.BackColor = &H80000005
    txtPatient.ForeColor = vbBlack
    txtPatient.Locked = False
    mstrGender = ""
    mstrAge = ""
    cmdNewPati.ToolTipText = "��������(F4)"
    cmdNewPati.Enabled = InStr(gstrPrivs, ";�ҺŲ��˽���;") > 0
    mstrClinic = ""
    mstrFeeType = ""
    lblInfo.Caption = "�Ա�:     ����:       �����:              �ѱ�:  "
    lblMoney.Caption = "����Ԥ�����:0.00  "
    lblSum.Caption = "�ϼ�"
    mintInsure = 0
    mlng����ID = 0
    chkBook.Enabled = True
    LoadPayMode False, False
    Set mrsInfo = Nothing
    If mblnAppointment Then
        mRegistFeeMode = EM_RG_����
    Else
        If (mty_Para.byt�Һ�ģʽ = 0 Or mty_Para.byt�Һ�ģʽ = 2) And gSysPara.bln��Һ�ģʽ = False Then
            mRegistFeeMode = EM_RG_����
            lblPayMode.Visible = True
            cboPayMode.Visible = True
            picPayMoney.Visible = True
            cmdPrice.Visible = mty_Para.byt�Һ�ģʽ = 2
        Else
            mRegistFeeMode = EM_RG_����
            lblPayMode.Visible = False
            cboPayMode.Visible = False
            picPayMoney.Visible = False
            cmdPrice.Visible = False
        End If
    End If
End Sub

Private Function GetTotalFromMshMoney(Optional ByVal str��Ŀ���� As String = "") As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ܽ��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-11-03 16:57:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, i As Long
    
    On Error GoTo errHandle
    With vsfMoney
        For i = 1 To .Rows - 1
            If str��Ŀ���� = "" Or Trim(.TextMatrix(i, 0)) = str��Ŀ���� Then
                dblMoney = dblMoney + Val(.TextMatrix(i, 2))
            End If
        Next
    End With
    GetTotalFromMshMoney = dblMoney
    Exit Function
errHandle:
    GetTotalFromMshMoney = 0
End Function



Private Function GetRegistMoney(Optional blnOnlyReg As Boolean = False, Optional blnNoBook As Boolean = False) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ�Һŵ��ĺϼƽ��
    '���:blnOnlyReg-�Ƿ������ȡ�Һŷ���
    '     blnNoBook-��ȡ������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-11-03 16:53:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl�ϼ� As Double, i As Integer
    Dim k As Integer
    
    If Not blnOnlyReg Then
        dbl�ϼ� = FormatEx(GetTotalFromMshMoney, 5)
    Else
        If mrsItems Is Nothing Then
             GetRegistMoney = FormatEx(GetTotalFromMshMoney, 3): Exit Function
        End If
        mrsItems.Filter = " ���� <> 4"
        If mrsItems.RecordCount = 0 Then
            mrsItems.Filter = 0
            GetRegistMoney = FormatEx(GetTotalFromMshMoney, 3): Exit Function
        End If
        With mrsItems
            Do While Not .EOF
                dbl�ϼ� = dbl�ϼ� + GetTotalFromMshMoney(Nvl(mrsItems!��Ŀ����, "-"))
                .MoveNext
            Loop
        End With
        mrsItems.Filter = 0
    End If
    If blnNoBook Then
        If Not mrsItems Is Nothing Then
            mrsItems.Filter = " ���� = 3"
            Do While Not mrsItems.EOF
                dbl�ϼ� = dbl�ϼ� + GetTotalFromMshMoney(Nvl(mrsItems!��Ŀ����, "-"))
                mrsItems.MoveNext
            Loop
            mrsItems.Filter = 0
        End If
    End If
    GetRegistMoney = FormatEx(dbl�ϼ�, 5)
End Function




Private Sub LoadPayMode(Optional ByVal blnPrepay As Boolean = False, Optional ByVal blnInsure As Boolean = False)
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSql As String, str���� As String
    
    strSql = _
        "Select B.����,B.����,Nvl(B.����,1) as ����,Nvl(A.ȱʡ��־,0) as ȱʡ" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where A.Ӧ�ó���=[1] And B.����=A.���㷽ʽ And Instr([2] ,','||B.����||',')>0" & _
        " Order by B.����"
    On Error GoTo errH
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, "�Һ�", ",3,7,8,")
    
    Set mcolCardPayMode = New Collection
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    If Not gobjSquare.objSquareCard Is Nothing Then
        strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    End If
    varData = Split(strPayType, ";")
    
    With cboPayMode
        .Clear: j = 0
'        Do While Not rsTemp.EOF
'            blnFind = False
'            For i = 0 To UBound(varData)
'                varTemp = Split(varData(i) & "|||||", "|")
'                If varTemp(6) = Nvl(rsTemp!����) Then
'                    blnFind = True
'                    Exit For
'                End If
'            Next
'
'            If Not blnFind Then
'                .AddItem Nvl(rsTemp!����)
'                mcolCardPayMode.Add Array("", Nvl(rsTemp!����), 0, 0, 0, 0, Nvl(rsTemp!����), 0, 0), "K" & j
'                If Val(Nvl(rsTemp!ȱʡ)) = 1 Then
'                    If .ListIndex = -1 Then
'                         .ItemData(.NewIndex) = 1: .ListIndex = .NewIndex
'                    End If
'                End If
'                j = j + 1
'            End If
'            rsTemp.MoveNext
'        Loop
     
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                rsTemp.Filter = "����='" & varTemp(6) & "'"
                If Not rsTemp.EOF Then
                    mcolCardPayMode.Add varTemp, "K" & j
                    .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                    j = j + 1
                End If
            End If
        Next
    End With
    
    If blnPrepay Then
        cboPayMode.AddItem "Ԥ����"
        If mty_Para.bln����ʹ��Ԥ�� Then
            cboPayMode.ListIndex = cboPayMode.NewIndex
        End If
    End If
    
    If blnInsure Then
        rsTemp.Filter = "���� = 3"
        If rsTemp.EOF Then
            mstrInsure = ""
            MsgBox "���ܼ���ҽ�����㷽ʽ,����!", vbInformation, gstrSysName
        Else
            cboPayMode.AddItem Nvl(rsTemp!����)
            mstrInsure = Nvl(rsTemp!����)
            If Not mty_Para.bln����ʹ��Ԥ�� Or blnPrepay = False Then
                cboPayMode.ListIndex = cboPayMode.NewIndex
            End If
            If (mintInsure <> 0 And MCPAR.���ղ�����) And cboPayMode.Text = mstrInsure And cboPayMode.Visible Then
                chkBook.Enabled = False
                chkBook.Value = 0
            Else
                chkBook.Enabled = True
            End If
        End If
    End If
    
    If cboPayMode.ListCount > 0 And cboPayMode.ListIndex = -1 Then
        cboPayMode.ListIndex = 0
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Function LoadRegPlans(ByVal intSelMode As Integer, Optional ByVal strFilter As String, Optional ByVal blnOtherDoctor As Boolean) As Boolean
    '����:��ȡ�ҺŰ���
    '���:intSelMode:��ȡģʽ=1-Ĭ�϶�ȡ;2-���˶�ȡ;3-���ж�ȡ
    '       blnOtherDoctor:�Ƿ��ȡ����ҽ���ű�
    Dim strTime As String, strState As String, strWhere As String
    Dim strSql As String, strIF As String, rsPlan As ADODB.Recordset
    Dim i As Integer, k As Integer, datNow As Date
    Dim DateThis As Date, strZero As String
    Dim str�ҺŰ��� As String, strViewSQL As String
    Dim str�ҺŰ��żƻ� As String, strCondition As String
    Dim str����         As String
    Dim vRect          As RECT
    Dim varTemp As Variant, varData As Variant
    On Error GoTo errH
    
    If chkAll.Value = 0 Then
        varTemp = Split(mty_Para.strStationRegOrder, "|")
        For i = 0 To UBound(varTemp)
            varData = Split(varTemp(i), ",")
            Select Case varData(0)
                Case "ҽ��"
                    str���� = str���� & ",Decode(ҽ������,Null,Decode(����ID," & mlngDept & ",3,4),Decode(����ID," & mlngDept & ",1,2)),ҽ������ " & IIf(varData(1) = 1, "", "desc")
                Case "����"
                    str���� = str���� & ",���� " & IIf(varData(1) = 1, "", "desc")
                Case "ִ��ʱ��"
                    str���� = str���� & ",��ʼʱ�� " & IIf(varData(1) = 1, "", "desc")
                Case "�ű�"
                    str���� = str���� & ",�ű� " & IIf(varData(1) = 1, "", "desc")
                Case "��Ŀ"
                    str���� = str���� & ",��Ŀ " & IIf(varData(1) = 1, "", "desc")
            End Select
        Next
        str���� = "Decode(��������,Null,2,1),�������� Desc" & str����
    Else
        str���� = "Decode(��������,Null,2,1),�������� Desc,Decode(ҽ������,'" & UserInfo.���� & "',1,2),Decode(����ID," & mlngDept & ",1,2),�ѹ�,�ű�,��Ŀ"
    End If
    
    If gstrDeptIDs <> "" And Not blnOtherDoctor Then strIF = " And Instr(','||[4]||',',','||a.����ID||',')>0"
    If mblnAppointment Then
        If mty_Para.blnԤԼ�������Ұ��� Then
            strIF = strIF & IIf(blnOtherDoctor, " And (a.ҽ������ <> [1] or a.ҽ������ Is Null)", " And (a.ҽ������ = [1] or a.ҽ������ Is Null)")
        Else
            strIF = strIF & IIf(blnOtherDoctor, " And (a.ҽ������ <> [1] )", " And (a.ҽ������ = [1])")
        End If
    Else
        If mty_Para.bln�ҺŰ������Ұ��� Then
            strIF = strIF & IIf(blnOtherDoctor, " And (a.ҽ������ <> [1] or a.ҽ������ Is Null)", " And (a.ҽ������ = [1] or a.ҽ������ Is Null)")
        Else
            strIF = strIF & IIf(blnOtherDoctor, " And (a.ҽ������ <> [1] )", " And (a.ҽ������ = [1])")
        End If
    End If
    
    If intSelMode = 2 Then
        strCondition = " And (b.���� Like [11] Or Upper(c.����) Like Upper([11]) Or Upper(zlSpellCode(c.����)) Like Upper([11]) Or Upper(a.ҽ������) Like Upper([11]) Or Upper(zlSpellCode(a.ҽ������)) Like Upper([11]))"
    End If
    
    strSql = "Select a.Id As ��¼ID, b.���� As �ű�, b.����, b.����id, c.���� As ����, a.��Ŀid, d.���� As ��Ŀ, Nvl(a.����ҽ��id,a.ҽ��id) As ҽ��id, Nvl(a.����ҽ������,a.ҽ������) As ҽ������, Nvl(a.�ѹ���, 0) As �ѹ�," & vbNewLine & _
            "       Nvl(a.��Լ��, 0) As ��Լ, a.�޺��� As �޺�, a.��Լ�� As ��Լ, Nvl(b.�Ƿ񽨲���, 0) As ����, Nvl(d.��Ŀ����, 0) As ����, a.���﷽ʽ As ����," & vbNewLine & _
            "       a.�Ƿ���ſ��� As ��ſ���, a.�ϰ�ʱ�� As �Ű�, a.��Դid, a.�Ƿ��ʱ�� As ��ʱ��, a.��ʼʱ��, a.��ֹʱ��, a.��������  " & vbNewLine & _
            "From �ٴ������¼ A, �ٴ������Դ B, ���ű� C, �շ���ĿĿ¼ D, ��Ա�� E" & vbNewLine & _
            "Where (a.�������� = [6] Or a.�������� = [8]) And Nvl(C.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And a.��Դid = b.Id And b.����id = c.Id And a.��Ŀid = d.Id And Nvl(a.�Ƿ�����, 0) = 0 " & vbNewLine & _
            "       And a.ҽ��id = e.Id(+) And (d.����ʱ�� is NULL Or d.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')) And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null) " & _
            "       And Nvl(a.�Ƿ񷢲�,0) = 1 "
    strSql = strSql & " And (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��,a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��,a.��ʼʱ��) Or Exists (Select 1 From �ٴ�������ſ��� C,�ٴ������¼ D Where D.ID=A.ID And C.��¼ID=D.ID And Nvl(C.�Ƿ�ͣ��,0) = 0 And D.�Ƿ���ſ��� =1 And D.�Ƿ��ʱ�� = 1 And C.��ʼʱ�� <> C.��ֹʱ��)) "
    If chkAll.Value <> 1 Then
        strSql = strSql & " And [5] Not Between  Nvl(a.ͣ�￪ʼʱ��,a.��ֹʱ��) And Nvl(a.ͣ����ֹʱ��,a.��ʼʱ��) "
    End If
    If mblnAppointment Then
        strSql = strSql & " And Nvl(a.ԤԼ����,0) <> 1 "
        DateThis = Format(dtpDate, "yyyy-mm-dd hh:mm:ss")
    Else
        DateThis = gobjDatabase.Currentdate
    End If
    datNow = gobjDatabase.Currentdate
    
    If mblnAppointment Then
        If Format(DateThis, "yyyy-mm-dd") = Format(datNow, "yyyy-mm-dd") Then
            strSql = strSql & "       And Not Exists (Select 1 From �ٴ������¼ Where Id=a.Id And ��ֹʱ�� < [9])"
        Else
            strSql = strSql & "       And Not Exists (Select 1 From �ٴ������¼ Where Id=a.Id And ��ֹʱ�� < [6])"

        End If
    Else
        strSql = strSql & " And Not Exists (Select 1 From �ٴ������¼ Where Id=a.Id And ��ֹʱ�� < [5]) "
    End If
    
    strSql = strSql & strIF & strCondition
    
    If intSelMode = 2 Then
        strCondition = " And (a.���� Like [11] Or Upper(c.����) Like Upper([11]) Or Upper(zlSpellCode(c.����)) Like Upper([11]) Or Upper(a.ҽ������) Like Upper([11]) Or Upper(zlSpellCode(a.ҽ������)) Like Upper([11]))"
    End If
    
    strTime = " Union All " & _
            "Select 0 As ��¼id, a.���� As �ű�, a.����, a.����id, c.���� As ����, a.��Ŀid, d.���� As ��Ŀ, a.ҽ��id, a.ҽ������, 0 As �ѹ�, 0 As ��Լ, Null As �޺�," & vbNewLine & _
            "       Null As ��Լ, Nvl(a.�Ƿ񽨲���, 0) As ����, Nvl(d.��Ŀ����, 0) As ����, 0 As ����, 0 As ��ſ���, Null As �Ű�, a.Id As ��Դid, 0 As ��ʱ��, Null As ��ʼʱ��, Null As ��ֹʱ��, Null As �������� " & vbNewLine & _
            "From �ٴ������Դ A, ���ű� C, �շ���ĿĿ¼ D" & vbNewLine & _
            "Where a.����id = c.Id And Nvl(C.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And a.��Ŀid = d.Id " & vbNewLine & _
            "      And Exists (Select 1 From �ٴ����ﰲ�� M,�ٴ������ N Where M.��ԴID=A.ID And M.����ID=N.ID And N.����ʱ�� Is Not Null) " & vbNewLine & _
            "      And Sysdate < Nvl(a.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) " & strIF & _
            "      And Not Exists(Select 1" & vbNewLine & _
            "                     From �ٴ������¼" & vbNewLine & _
            "                     Where ��Դid = a.Id And (�������� = [6] Or �������� = [8])" & vbNewLine & _
            "                           And [5] Between ��ʼʱ�� And ��ֹʱ��" & vbNewLine & _
            "                           And (��ʼʱ�� < Nvl(ͣ�￪ʼʱ��, ��ֹʱ��) Or ��ֹʱ�� > Nvl(ͣ����ֹʱ��, ��ʼʱ��))" & vbNewLine & _
            "                           And Nvl(�Ƿ�����, 0) = 0 And Nvl(�Ƿ񷢲�, 0) = 1)"
        
    If mblnAppointment Then
        'ԤԼ�Һ�
        strSql = strSql & " And (a.��Լ�� > 0 Or a.��Լ�� Is Null)"
        strSql = strSql & " And Nvl(a.ԤԼ����,0) <> 1 "
        strSql = strSql & " And Sysdate + zl_Fun_GetAppointmentDays + Decode(Nvl(B.ԤԼ����," & mty_Para.int��ԤԼ���� & "),0,15,Nvl(B.ԤԼ����," & mty_Para.int��ԤԼ���� & ")" & ") > [6] "
    Else
        '�Һ�
        If chkAll.Value = 1 Then strSql = strSql & strTime
    End If
    
    strViewSQL = "Select RowNum As Id,A.��¼id,A.�ű�,a.����,a.����,a.����id,a.��Ŀ,a.ҽ������,a.�ѹ�,a.��Լ,a.�޺�,a.��Լ," & _
                 "      Decode(nvl(a.����,0),1,'��','') As ����,Decode(nvl(a.����,0),1,'��','') As ����,Decode(nvl(a.��ſ���,0),1,'��','') As ��ſ���," & _
                 "      a.�Ű�,Decode(nvl(a.��ʱ��,0),1,'��','') As ��ʱ��,a.�������� From (" & strSql & ") A Order By " & str����
    strSql = "Select * From (" & strSql & ") Order By " & str����
                
    Set mrsPlan = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, _
            UserInfo.����, "%", "", gstrDeptIDs, DateThis, CDate(Format(DateThis, "yyyy-MM-dd")), CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60, CDate(Format(DateThis - 1, "yyyy-MM-dd")), gobjDatabase.Currentdate, gdatRegistTime, "%" & strFilter & "%")
                

 
    If mrsPlan.RecordCount <> 0 Then
        If intSelMode = 1 Or mrsPlan.RecordCount = 1 Then
            'Ĭ�϶�ȡ
            Call ReadLimit(Val(Nvl(mrsPlan!��¼ID)), Nvl(mrsPlan!�ű�))
        Else
            vRect = GetControlRect(txtReg.hWnd)
            Set rsPlan = gobjDatabase.ShowSQLSelect(Me, strViewSQL, 0, "����ѡ��", False, "", "����ѡ��", _
                                                False, False, True, vRect.Left, vRect.Top - 250, 600, False, True, False, _
                                                UserInfo.����, "%", "", gstrDeptIDs, DateThis, CDate(Format(DateThis, "yyyy-MM-dd")), _
                                                CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60, CDate(Format(DateThis - 1, "yyyy-MM-dd")), _
                                                gobjDatabase.Currentdate, gdatRegistTime, "%" & strFilter & "%")
            If rsPlan Is Nothing Then
                Call ReadLimit(Val(Nvl(mrsPlan!��¼ID)), Nvl(mrsPlan!�ű�))
            Else
                If Not rsPlan.EOF Then
                    Call ReadLimit(Val(Nvl(rsPlan!��¼ID)), Nvl(rsPlan!�ű�))
                Else
                    Call ReadLimit(Val(Nvl(mrsPlan!��¼ID)), Nvl(mrsPlan!�ű�))
                End If
            End If
        End If
        Call LoadDoctor
        Call ResetDefault����
        Call LoadFeeItem(Val(Nvl(mrsPlan!��ĿID)), chkBook.Value = 1, mstrPriceGrade)
        Call GetActiveView
        If mblnAppointment Then
            Select Case mViewMode
                Case V_��ͨ�ŷ�ʱ��, v_ר�Һŷ�ʱ��
                    dtpTime.Enabled = False
                    cmdTime.Visible = True
                Case Else
                    dtpTime.Enabled = True
                    cmdTime.Visible = False
            End Select
            Call SetDefultRegTime
        Else
            cmdTime.Visible = False
        End If
        lblDeptName.Caption = Nvl(mrsPlan!����)
        If txtReg.Visible And txtReg.Enabled Then txtReg.SetFocus
    End If
    LoadRegPlans = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub ReadLimit(ByVal lng��¼ID As Long, str���� As String)

    If mrsPlan Is Nothing Then Exit Sub
    If mrsPlan.State = 0 Then Exit Sub
    
    mrsPlan.Filter = "��¼ID=" & lng��¼ID & " And �ű�='" & str���� & "'"
    
    If mrsPlan.RecordCount = 0 Then Exit Sub
    mblnIntact = True
    mblnChangeByCode = True
    If Nvl(mrsPlan!ҽ������) = "" Then
        txtReg.Text = "[" & Nvl(mrsPlan!�ű�) & "]" & Nvl(mrsPlan!��Ŀ)
    Else
        txtReg.Text = "[" & Nvl(mrsPlan!�ű�) & "]" & Nvl(mrsPlan!��Ŀ) & "(" & Nvl(mrsPlan!ҽ������) & ")"
    End If
    mblnChangeByCode = False
    txtReg.Tag = Nvl(mrsPlan!�ű�)
    
    If mblnAppointment Then
        If Nvl(mrsPlan!��Լ) = "" Then
            lblLimit.Caption = "��Լ:" & Nvl(mrsPlan!��Լ, 0)
        Else
            lblLimit.Caption = "��Լ:" & Nvl(mrsPlan!��Լ) & "  ��Լ:" & Nvl(mrsPlan!��Լ, 0)
        End If
    Else
        If Nvl(mrsPlan!�޺�) = "" Then
            lblLimit.Caption = "�ѹ�:" & Nvl(mrsPlan!�ѹ�, 0)
        Else
            lblLimit.Caption = "�޺�:" & Nvl(mrsPlan!�޺�) & "  �ѹ�:" & Nvl(mrsPlan!�ѹ�, 0)
        End If
    End If

    If Val(Nvl(mrsPlan!����)) = 0 Then
        lbl��.Visible = False
    Else
        lbl��.Visible = True
    End If
    
    lblʱ��.Caption = Nvl(mrsPlan!�Ű�)
    lblʱ��.Visible = Nvl(mrsPlan!�Ű�) <> ""
    If Not mrsInfo Is Nothing Then Call Load�ѱ�(Nvl(mrsInfo!�ѱ�))
End Sub

Private Sub txtReg_Change()
    If mblnChangeByCode = True Then Exit Sub
    mblnIntact = False
End Sub

Private Sub txtReg_GotFocus()
    Call gobjControl.TxtSelAll(txtReg)
End Sub

Private Sub txtReg_KeyPress(KeyAscii As Integer)
    If mblnIntact Then
        If KeyAscii = 13 Then gobjCommFun.PressKeyEx vbKeyTab
    Else
        If KeyAscii = 13 Then Call LoadRegPlans(2, txtReg.Text)
    End If
End Sub

Private Function GetActiveView()
    '�õ���ǰ�Һ�ҵ��  ��ȡ�������͵�����
    Dim strSql          As String
    Dim rsTmp           As ADODB.Recordset
    If mrsPlan Is Nothing Then Exit Function
    If mrsPlan.RecordCount = 0 Then Exit Function
    On Error GoTo errH
    
    strSql = "Select 1 From �ٴ������¼ Where ID=[1] And Nvl(�Ƿ��ʱ��,0)=1 "
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsPlan!��¼ID))
    
    If rsTmp.RecordCount > 0 And Val(Nvl(mrsPlan!��ſ���)) = 1 Then
       '*********************
       'ר�Һŷ�ʱ��
       '*********************
       mViewMode = v_ר�Һŷ�ʱ��

    ElseIf rsTmp.RecordCount > 0 And Val(Nvl(mrsPlan!��ſ���)) = 0 Then
       '*********************
       '��ͨ�ŷ�ʱ��
       '*********************
       mViewMode = V_��ͨ�ŷ�ʱ��

    ElseIf Val(Nvl(mrsPlan!��ſ���)) = 1 And Nvl(mrsPlan!�޺�) <> "" Then
       '*********************
       'ר�ҺŲ���ʱ��
       '*********************
       mViewMode = v_ר�Һ�

     Else
       '*********************
       '��ͨ��
       '*********************
       mViewMode = V_��ͨ��

    End If
    Set rsTmp = Nothing
Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
         Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Function SelectTimeSn() As Boolean
    '**************************************
    '����ʱ��
    '����ʱ���Ƿ���سɹ����Ƿ��з�ʱ��
    '**************************************
     Dim strSql         As String
     Dim vRect          As RECT
     
   
    If Not mblnAppointment Then Exit Function
    
    If mrsPlan Is Nothing Then Exit Function
    If mrsPlan.State <> 1 Then Exit Function
    If mrsPlan.EOF Then Exit Function
    
    If Not (Val(Nvl(mrsPlan!��ſ���)) = 1 And Val(Nvl(mrsPlan!��ʱ��)) = 1) Then Exit Function
    
    lblSn.Caption = ""
    strSql = "" & _
    " Select Rownum As Id, ���, To_Char(��ʼʱ��, 'hh24') || ':00' As ʱ���, To_Char(��ʼʱ��, 'hh24:mi') As ��ʼʱ��," & vbNewLine & _
    "       To_Char(��ֹʱ��, 'hh24:mi') As ����ʱ��, ��ʼʱ�� As ��ϸ��ʼʱ��, ��ֹʱ�� As ��ϸ����ʱ�� " & vbNewLine & _
    " From �ٴ�������ſ���" & vbNewLine & _
    " Where ��¼id = [1] And Nvl(�Һ�״̬,0) = 0 And Nvl(�Ƿ�ԤԼ,0)=1 And Trunc(��ʼʱ��) = [2]" & vbNewLine & _
    "Order By ��ϸ��ʼʱ��"

    If strSql = "" Then Exit Function
    vRect = GetControlRect(dtpTime.hWnd)
    
    On Error GoTo errH
    
    Set mrsʱ��� = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "ԤԼʱ��ѡ��", False, "", "ԤԼʱ��ѡ��", _
                                                False, False, True, vRect.Left, vRect.Top - 300, 600, False, True, False, mrsPlan!��¼ID, CDate(Format(dtpDate.Value, "yyyy-mm-dd")))
    If mrsʱ��� Is Nothing Then Exit Function
    If mrsʱ���.EOF Then Exit Function
    
    lblSn.Caption = "���:" & Val(Nvl(mrsʱ���!���))
    dtpTime.Value = Format(mrsʱ���!��ʼʱ��, "hh:mm:ss")
    SelectTimeSn = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub LoadFeeItem(ByVal lngItemID As Long, ByVal blnBook As Boolean, ByVal strPriceGrade As String)
    Dim strSql As String, i As Integer, dblTotal As Double
    Dim rsIncomes As ADODB.Recordset, curӦ�� As Currency, curʵ�� As Currency
    Dim j As Integer, rsItems As ADODB.Recordset, lng����ID As Long
    If lngItemID = 0 Then Exit Sub
    '����:1-���Һŷ��� 2-������� 3-������
    ReadRegistPrice lngItemID, blnBook, False, mstrFeeType, rsItems, rsIncomes, , , , 1, Val(Nvl(mrsPlan!����ID)), strPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
    If mintInsure <> 0 Then
        If MCPAR.�Һż����Ŀ = True Then
            If gclsInsure.CheckItem(mintInsure, 2, 0, rsItems) = False Then
                MsgBox "ҽ�������շ���Ŀ���ʧ�ܣ����ܼ���" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            If Not mrsInfo.EOF Then lng����ID = Val(Nvl(mrsInfo!����ID))
        End If
    End If
    ReadRegistPrice lngItemID, blnBook, False, mstrFeeType, rsItems, rsIncomes, lng����ID, mintInsure, _
        txtReg.Tag, IIf(mblnAppointment, 1, 0), , strPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
    
    vsfMoney.Clear 1
    vsfMoney.Rows = 2
    lblTotal.Caption = Format(0, "0.00")
    lblPayMoney.Caption = Format(0, "0.00")
    dblTotal = 0
    If rsItems.RecordCount = 0 Then Exit Sub
    rsItems.MoveFirst
    For i = 1 To rsItems.RecordCount
        With vsfMoney
            .RowData(.Rows - 1) = Nvl(rsItems!��ĿID)
            .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = Nvl(rsItems!��Ŀ����)
            rsIncomes.Filter = "��ĿID=" & rsItems!��ĿID
            curӦ�� = 0: curʵ�� = 0
            For j = 1 To rsIncomes.RecordCount
                curӦ�� = curӦ�� + rsIncomes!Ӧ��
                curʵ�� = curʵ�� + rsIncomes!ʵ��
                rsIncomes.MoveNext
            Next j
            .TextMatrix(.Rows - 1, .ColIndex("Ӧ�ս��")) = Format(curӦ��, "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("ʵ�ս��")) = Format(curʵ��, "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(rsItems!����)
            .Rows = .Rows + 1
        End With
        rsItems.MoveNext
    Next i
    If vsfMoney.Rows > 2 Then vsfMoney.Rows = vsfMoney.Rows - 1
    For i = 1 To vsfMoney.Rows - 1
        dblTotal = dblTotal + Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("ʵ�ս��")))
    Next i
    vsfMoney.RowHeightMin = 350
    lblTotal.Caption = Format(dblTotal, "0.00")
    lblPayMoney.Caption = Format(dblTotal, "0.00")
    lblRoomName.Caption = gstrRooms
End Sub


Private Function GetSNState(lng��¼ID As Long) As ADODB.Recordset
    Dim strSql           As String
    On Error GoTo errH

    strSql = "    " & vbNewLine & " Select A.���,A.�Һ�״̬,A.����Ա����,Decode(A.�Һ�״̬,2,1,0) as ԤԼ,To_Char(B.��������,'hh24:mi:ss') as ����  "
    strSql = strSql & vbNewLine & " From �ٴ�������ſ��� A, �ٴ������¼ B "
    strSql = strSql & vbNewLine & " Where B.ID=[1] And B.ID=A.��¼ID"
    Set GetSNState = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng��¼ID)

    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Function zlGet��ǰ���ڼ�(Optional strDate As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���������ڼ�
    '����:���˺�
    '����:2010-02-04 14:42:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset, bln��ǰ���� As Boolean, strTemp As String
    If strDate = "" Then
        strSql = "Select Decode(To_Char(Sysdate,'D'),'1','��','2','һ','3','��','4','��','5','��','6','��','7','��',NULL) as ����  From dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)
    Else
        strSql = "Select Decode(To_Char([1],'D'),'1','��','2','һ','3','��','4','��','5','��','6','��','7','��','') As ���� From dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(strDate))
    End If
    
    If rsTemp.EOF = True Then
        Exit Function
    End If
    strTemp = Nvl(rsTemp!����)
    zlGet��ǰ���ڼ� = strTemp
End Function

Private Sub Init�ѱ�()
    '��ʼ��ȱʡ�ѱ�
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo Errhand
    
    strSql = "Select ���� From �ѱ� Where ȱʡ��־ = 1 And Nvl(�������, 3) In (1, 3)"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)
    If Not rsTmp.EOF Then
        mstrDef�ѱ� = Nvl(rsTmp!����)
    Else
        MsgBox "�޷���ȡȱʡ�ѱ�����ȱʡ�ѱ��Ƿ���ȷ���ã�", vbInformation, gstrSysName
    End If
    Exit Sub
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Function Load�ѱ�(Optional ByVal str�ѱ� As String) As Boolean
    '����:���ݿ��Ҽ��ز��˷ѱ�
    '����:str�ѱ�-�����ϴ�ʹ�õĵķѱ�
    '����:�ɹ�,����true,���򷵻�False
    
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo Errhand
    
    If mrsInfo Is Nothing Then Exit Function
    If mrsPlan Is Nothing Then Exit Function
    If mrsPlan.EOF Then Exit Function
    If str�ѱ� <> "" Then
        strSql = " Select 1 From �ѱ� A, �ѱ����ÿ��� B" & _
                 " Where a.���� = b.�ѱ�(+) And a.���� = 1" & _
                 "      And Trunc(Sysdate) Between Nvl(a.��Ч��ʼ, To_Date('1900-01-01', 'YYYY-MM-DD'))" & _
                 "      And Nvl(a.��Ч����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                 "      And Nvl(a.�������, 3) In (1, 3) And (B.����ID=[1] or B.����ID is NULL) and A.����=[2]"
        
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsPlan!����ID)), str�ѱ�)
        If Not rsTmp.EOF Then
            mstrFeeType = str�ѱ�
        Else
            mstrFeeType = mstrDef�ѱ�
        End If
    Else
        mstrFeeType = mstrDef�ѱ�
    End If
    If mstrFeeType = "" Then
        MsgBox "δ�ҵ������ڡ�" & Nvl(mrsPlan!����) & "����ȱʡ�ѱ�,���ڡ�������ϸ��Ϣ�����������ò��˷ѱ�", vbInformation, gstrSysName
        Load�ѱ� = False
        Exit Function
    End If
    lblInfo.Caption = "�Ա�:" & mstrGender & "   ����:" & mstrAge & "   �����:" & mstrClinic & "   �ѱ�:" & mstrFeeType
    Load�ѱ� = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function




