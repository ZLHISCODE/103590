VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmStationRegist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ��վ�Һ�"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8445
   Icon            =   "frmStationRegist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8445
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
      Left            =   3780
      TabIndex        =   49
      Top             =   5760
      Width           =   1725
   End
   Begin VB.CheckBox chkAll 
      Height          =   360
      Left            =   8055
      Picture         =   "frmStationRegist.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "��ʾ�������"
      Top             =   60
      Width           =   345
   End
   Begin VB.CommandButton cmdNewPati 
      Height          =   345
      Left            =   2940
      Picture         =   "frmStationRegist.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "��������(F4)"
      Top             =   600
      Width           =   350
   End
   Begin VB.PictureBox picPayMoney 
      BackColor       =   &H80000005&
      Height          =   420
      Left            =   6645
      ScaleHeight     =   360
      ScaleWidth      =   1695
      TabIndex        =   37
      Top             =   4942
      Width           =   1755
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
         Left            =   930
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
         Picture         =   "frmStationRegist.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "����ҽ���ű�"
         Top             =   45
         Width           =   345
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����"
         Height          =   255
         Left            =   7500
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   570
         Width           =   930
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
      Begin VB.CommandButton cmdReg 
         Height          =   345
         Left            =   4020
         Picture         =   "frmStationRegist.frx":1788
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "��ǰҽ���ű�"
         Top             =   45
         Width           =   345
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
         Left            =   6990
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
         Width           =   1875
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
         Width           =   6120
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
         FormatString    =   $"frmStationRegist.frx":218A
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
         Caption         =   "���:"
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
         Left            =   7635
         TabIndex        =   50
         Top             =   105
         Width           =   600
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
         Top             =   105
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
      Top             =   5490
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
      Left            =   4680
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
      Left            =   6975
      TabIndex        =   12
      Top             =   5760
      Width           =   1335
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
      Left            =   5595
      TabIndex        =   11
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
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
      Left            =   90
      TabIndex        =   13
      Top             =   5760
      Width           =   1335
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
      Left            =   4140
      TabIndex        =   28
      Top             =   1568
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   360
      Left            =   3045
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
      Left            =   5700
      ScaleHeight     =   300
      ScaleWidth      =   2595
      TabIndex        =   44
      Top             =   1560
      Width           =   2655
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
      Left            =   4500
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
      Left            =   3315
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
      Left            =   5760
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
      Left            =   2505
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
Attribute VB_Name = "frmStationRegist"
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
Private mstrAge As String, mstr�ѱ� As String, mstr�Ա� As String, mstr����� As String
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
Private mcolArrangeNo As Collection
Private mblnIntact As Boolean, mstrUseType As String
Private mlng����ID As Long, mintIDKind As Integer
Private mcur������� As Currency
Private mblnOK As Boolean, mstrCardPass As String
Private mstrNO As String, mlngSN As Long
Private mintInsure As Integer, mblnUpdateAge As Boolean
Private mdatLast As Date
Private mblnChangeByCode As Boolean
Private mstrCardNO As String
Private mcur����͸֧ As Currency
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
    intͬ����Լ��           As Integer  'ͬ������Լ
    intͬ���޹���           As Integer
    blnͬ���޹Ҽ���         As Boolean
    int����ԤԼ������       As Integer
    int���˹Һſ�����       As Integer
    intר�ҺŹҺ�����       As Integer
    intר�Һ�ԤԼ����       As Integer
    strStationRegOrder As String    'ҽ��վ�Һ������ַ���
    blnShowAllPlan      As Boolean   ' �Ƿ���ʾ������ű�
End Type
Private mty_Para As ty_ModulePara
Private mstr����IDs As String
Private mstrPriceGrade As String, mintPriceGradeStartType As Integer
Private mobjRegister As clsRegist
Private mstrDef�ѱ� As String  'ȱʡ�ѱ�

Public Sub zlShowMe(ByVal frmMain As Object, ByVal objRegister As clsRegist, ByVal lngModul As Long, ByVal strDeptIDs As String, _
                    ByVal blnAppointment As Boolean, ByVal lng����ID As Long, ByRef strOutNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ��վ�Һ����
    '���:strDeptIDs-�Һſ���,֧�ֶ��,�ö��ŷָ�
    '     blnAppointment-�Ƿ�ԤԼ����
    '     objRegister-clsRegist����
    '����:strOutNO-�Һųɹ���,�����Һŵ��ݺ�
    '����:������
    '����:2016-7-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    mblnAppointment = blnAppointment
    Set mobjRegister = objRegister
    mlngModul = lngModul
    mlng����ID = lng����ID
    
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
        Call GetPriceGrade(gstrNodeNo, 0, 0, "", , , mstrPriceGrade)
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
                    MsgBox "��û�����ú͹��õĹҺ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
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

Private Sub cmdOther_Click()
    Call LoadRegPlans(3, , True)
End Sub

Private Sub cmdPrice_Click()
    If SaveData_Price = False Then Exit Sub
    mblnUpdateAge = False
    Call ReloadPage
    mblnOK = True
    Unload Me
End Sub
Private Sub cmdReg_Click()
    Call LoadRegPlans(3)
End Sub

Private Sub SetDefultRegTime()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ�ĹҺ�ʱ��
    '����:2018-02-02 15:05:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtCurDate As Date, dtSysDate As Date, strNO As String, str���� As String, strSql As String
    Dim rsTmp As ADODB.Recordset, rsTime As ADODB.Recordset, str����ʱ�� As String
    Dim lng����ID As Long, lng�ƻ�ID As Long
    Dim lngSN As Long, blnAdd As Boolean, blnNotWork As Boolean
    On Error GoTo errH
    
    
    lblSn.Caption = ""
    str���� = zlGet��ǰ���ڼ�(dtpDate.Value)
    
    If Nvl(mrsPlan.Fields(str����)) = "" Then
        dtpTime.Value = Format(GetWorkTimeDefualtTime("����", Format(dtpDate.Value, "yyyy-mm-dd")), "hh:mm:ss")
        Exit Sub
    End If
   
    Select Case mViewMode
    Case V_��ͨ�ŷ�ʱ��, v_ר�Һŷ�ʱ��
        
        strSql = "Select Distinct a.��� As ID, To_Char(a.��ʼʱ��, 'hh24:mi') As ��ʼʱ��, To_Char(a.����ʱ��, 'hh24:mi') As ����ʱ��" & vbNewLine & _
                "From �ҺŰ���ʱ�� A, �ҺŰ��� B" & vbNewLine & _
                "Where a.����id = b.Id And b.���� = [1] And" & vbNewLine & _
                " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.��ʼʱ��, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
                "      Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����'," & vbNewLine & _
                "             Null) = a.����(+) And Not Exists" & vbNewLine & _
                " (Select Count(1)" & vbNewLine & _
                "       From �Һ����״̬" & vbNewLine & _
                "       Where Trunc(����) = [2] And ���� = b.���� And (��� = a.��� Or ��� Like a.��� || '__') Having" & vbNewLine & _
                "        Count(1) - a.�������� >= 0) And Not Exists" & vbNewLine & _
                " (Select 1" & vbNewLine & _
                "       From �ҺŰ��żƻ� E" & vbNewLine & _
                "       Where e.����id = b.Id And e.���ʱ�� Is Not Null And" & vbNewLine & _
                "             [2] Between Nvl(e.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
                "             e.ʧЧʱ��)"
        strSql = strSql & " Union " & _
                "Select Distinct a.��� As ID, To_Char(a.��ʼʱ��, 'hh24:mi') As ��ʼʱ��, To_Char(a.����ʱ��, 'hh24:mi') As ����ʱ��" & vbNewLine & _
                "From �Һżƻ�ʱ�� A, �ҺŰ��żƻ� B, �ҺŰ��� C," & vbNewLine & _
                "     (Select Max(a.��Чʱ��) ��Ч" & vbNewLine & _
                "       From �ҺŰ��żƻ� A, �ҺŰ��� B" & vbNewLine & _
                "       Where a.����id = b.Id And b.���� = [1] And a.���ʱ�� Is Not Null And" & vbNewLine & _
                "             [2] Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
                "             a.ʧЧʱ��) D" & vbNewLine & _
                "Where a.�ƻ�id = b.Id And b.����id = c.Id And c.���� = [1] And b.��Чʱ�� = d.��Ч And b.���ʱ�� Is Not Null And" & vbNewLine & _
                " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.��ʼʱ��, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
                "      [2] Between Nvl(b.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
                "      b.ʧЧʱ�� And Not Exists" & vbNewLine & _
                " (Select Count(1)" & vbNewLine & _
                "       From �Һ����״̬" & vbNewLine & _
                "       Where Trunc(����) = [2] And ���� = b.���� And (��� = a.��� Or ��� Like a.��� || '__') Having" & vbNewLine & _
                "        Count(1) - a.�������� >= 0) And Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5'," & vbNewLine & _
                "                                           '����', '6', '����', '7', '����', Null) = a.����(+)" & vbNewLine & _
                "Order By ��ʼʱ��"
    
        dtCurDate = Format(dtpDate, "yyyy-mm-dd")
        strNO = txtReg.Tag
        
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, dtCurDate)
        If Not rsTmp.EOF Then
            'ʱ�ε�����ʱ��,ȡ��Сʱ��
            dtpTime.Value = Format(Nvl(rsTmp!��ʼʱ��), "hh:mm:ss")
            lblSn.Caption = "���:" & Val(Nvl(rsTmp!ID))
            Exit Sub
        End If
        
        If GetRegData(lngSN, str����ʱ��, blnAdd, blnNotWork) Then
            If lngSN <> 0 Then lblSn.Caption = "���:" & lngSN
            If str����ʱ�� <> "" Then
                If IsDate(str����ʱ��) Then dtpTime.Value = Format(CDate(str����ʱ��), "hh:mm:ss"): Exit Sub
            End If
        End If
       dtpTime.Value = Format(GetWorkTimeDefualtTime(Nvl(mrsPlan.Fields(str����)), Format(dtpDate.Value, "yyyy-mm-dd")), "hh:mm:ss")
        Exit Sub
    Case v_ר�Һ�
        lng�ƻ�ID = Val(Nvl(mrsPlan!�ƻ�ID))
        lng����ID = Val(Nvl(mrsPlan!ID))
        
        dtCurDate = Format(dtpDate, "yyyy-mm-dd")
        If mobjRegister.zlGetRegisterNextSn__Tradition(lng����ID, lng�ƻ�ID, dtCurDate, InStr(gstrPrivs, ";�Ӻ�;"), mblnAppointment, False, lngSN, str����ʱ��) = False Then
            dtpTime.Value = Format(GetWorkTimeDefualtTime(Nvl(mrsPlan.Fields(str����)), Format(dtpDate.Value, "yyyy-mm-dd")), "hh:mm:ss"): Exit Sub
        End If
        If lngSN <> 0 Then lblSn.Caption = "���:" & lngSN
        If mblnAppointment Then
            If Format(dtpDate.Value, "yyyy-mm-dd") > Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
                str����ʱ�� = GetWorkTimeDefualtTime(Nvl(mrsPlan.Fields(str����)), Format(dtpDate.Value, "yyyy-mm-dd"))
            End If
        End If
        If str����ʱ�� <> "" Then
            If IsDate(str����ʱ��) Then dtpTime.Value = Format(CDate(str����ʱ��), "hh:mm:ss"): Exit Sub
        End If
        dtpTime.Value = Format(GetWorkTimeDefualtTime(Nvl(mrsPlan.Fields(str����)), Format(dtpDate.Value, "yyyy-mm-dd")), "hh:mm:ss")
    Case Else
        dtpTime.Value = Format(GetWorkTimeDefualtTime(Nvl(mrsPlan.Fields(str����)), Format(dtpDate.Value, "yyyy-mm-dd")), "hh:mm:ss")
    End Select
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetWorkTimeDefualtTime(ByVal strWorkName As String, ByVal strRegDate As String, Optional ByVal strCurSysDate As String = "") As Date
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
    Dim dtRegDate As Date
    
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
    rsTime.Filter = "ʱ���='" & strWorkName & "' and ����=NULL and վ��=NULL"
    If rsTime.EOF Then
        rsTime.Filter = 0
        GetWorkTimeDefualtTime = dtSysDate: Exit Function
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
        If Nvl(mrsPlan!ҽ��) = "" Then
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
            mrsDoctor.Filter = "����='" & Nvl(mrsPlan!ҽ��) & "'"
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
            MsgBox "û�����ó��ùҺ�ժҪ,�����ֵ����������", vbOKOnly + vbInformation, gstrSysName
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
     mstr�ѱ� = ""
     mstrDef�ѱ� = ""
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
    If mrsPlan.State = 0 Then Exit Sub
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
    txtPatient.Text, NeedName(mstr�Ա�), str����, dblMoney, mstrCardNO, mstrPassWord, _
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
Private Function CheckIsPatiBlacklist() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:�Ϸ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-31 11:21:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnר�Һ� As Boolean, bytMode As Byte
    Dim strSql As String, datԤԼʱ�� As Date
    Dim rsCheck As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mrsInfo Is Nothing Then Exit Function

    If mblnAppointment Then
        bytMode = 1
        datԤԼʱ�� = CDate(Format(dtpDate.Value, "yyyy-mm-dd"))
    Else
        bytMode = 0
        datԤԼʱ�� = CDate(Format(gobjDatabase.Currentdate, "yyyy-mm-dd"))
    End If
    
    blnר�Һ� = Nvl(mrsPlan!ҽ��) <> ""
    
    strSql = "Select Zl_Fun_���˹Һż�¼_Check([1],[2],[3],Null,[4],[5]) As ����� From Dual"
    Set rsCheck = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, bytMode, Val(Nvl(mrsInfo!����ID)), Trim(txtReg.Tag), datԤԼʱ��, IIf(blnר�Һ�, 1, 0))
    If rsCheck.EOF Then
        MsgBox "��Ч�Լ��ʧ��,�޷�������", vbInformation, gstrSysName
        Exit Function
    End If

   strSql = Nvl(rsCheck!�����)
   If Val(Mid(strSql, 1, 1)) <> 0 Then
       MsgBox Mid(strSql, 3), vbInformation, gstrSysName
       Exit Function
   End If
    CheckIsPatiBlacklist = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDataIsValied(ByVal bytMode As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵĺϷ���
    '���:bytMode-0-����;1-����;2-����
    '����:
    '����:���ݺϷ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-31 11:17:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    '1.����������ݵĺϷ���
    If CheckValied(bytMode) = False Then Exit Function
    
    '2.��������
    If CheckIsPatiBlacklist = False Then Exit Function
    
    '3.������ؼ��
    If bytMode = 1 Then
       If mRegistFeeMode = EM_RG_���� And mty_Para.blnԤԼʱ�տ� And mblnAppointment Then
           MsgBox "��֧�������ƺ���㲡�˵�ԤԼ�տ�Һţ�", vbInformation, gstrSysName
           Exit Function
       End If
    End If
    
    '4.����ģʽ���
    If mblnAppointment = False Or (mblnAppointment = True And mty_Para.blnԤԼʱ�տ�) Then
        If zlIsAllowPatiChargeFeeMode(ZVal(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!����ģʽ))) = False Then Exit Function
    End If
     
    CheckDataIsValied = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetPrintProofIsPrint(ByVal bytMode As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ӡƾ���Ƿ���Ҫ��ӡ
    '���:bytMode-0-����;1-����;2-����
    '����:
    '����:��Ҫ��ӡ����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-31 11:50:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    On Error GoTo errHandle
    
    'ԤԼʱ���տ�򲻴�ӡ�Һ�ƾ��
    If mblnAppointment And mty_Para.blnԤԼʱ�տ� = False Then Exit Function
    
    If mty_Para.int�Һ�ƾ����ӡ = 0 Then Exit Function '����ӡ
    
    If InStr(gstrPrivs, ";���˹Һ�ƾ��;") = 0 Then
        '����Ƿ����Ȩ��
        MsgBox "��û��" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "ƾ����ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mty_Para.int�Һ�ƾ����ӡ = 1 Then '�Զ���ӡ
        GetPrintProofIsPrint = True: Exit Function
    End If
    
    '��ʾ��ӡ
    If MsgBox("Ҫ��ӡ" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "ƾ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    GetPrintProofIsPrint = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetInvoiceIsPrint(ByVal bytMode As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ӡ��Ʊ�Ƿ���Ҫ��ӡ
    '���:bytMode-0-����;1-����;2-����
    '����:
    '����:��Ҫ��ӡ����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-31 11:50:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    On Error GoTo errHandle
    
    If bytMode = 1 Or bytMode = 2 Or mblnAppointment And mty_Para.blnԤԼʱ�տ� = False Then Exit Function '���ۼ����ʻ�ԤԼʱ���տ��ӡ��Ʊ
    If mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ�� Then Exit Function
    
    
    
    If mty_Para.int�Һŷ�Ʊ��ӡ = 0 Then Exit Function '����ӡ
    
    If InStr(gstrPrivs, ";�Һŷ�Ʊ��ӡ;") = 0 Then
        '����Ƿ����Ȩ��
        MsgBox "��û��" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "��Ʊ��ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mty_Para.int�Һŷ�Ʊ��ӡ = 1 Then '�Զ���ӡ
        GetInvoiceIsPrint = True: Exit Function
    End If
    
    '��ʾ��ӡ
    If MsgBox("Ҫ��ӡ" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "��Ʊ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    GetInvoiceIsPrint = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetDepositBillIsPrint(ByVal bytMode As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡԤԼ���Ƿ���Ҫ��ӡ
    '���:bytMode-0-����;1-����;2-����
    '����:
    '����:��Ҫ��ӡ����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-31 11:50:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    On Error GoTo errHandle
    
    If Not (mblnAppointment And mty_Para.blnԤԼʱ�տ� = False) Then Exit Function  'ֻ��ԤԼ��(δ�տ�)�Ŵ�ӡ
     
    
    If mty_Para.intԤԼ�ҺŴ�ӡ = 0 Then Exit Function '����ӡ
    
    If InStr(gstrPrivs, ";ԤԼ�Һŵ�;") = 0 Then
        '����Ƿ����Ȩ��
        MsgBox "��û��ԤԼ�Һŵ���ӡ��Ȩ�ޣ�����ϵ����Ա����", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mty_Para.intԤԼ�ҺŴ�ӡ = 1 Then '�Զ���ӡ
        GetDepositBillIsPrint = True: Exit Function
    End If
    
    '��ʾ��ӡ
    If MsgBox("Ҫ��ӡԤԼ�Һŵ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    
    GetDepositBillIsPrint = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetRegData(ByRef lngSn_Out As Long, ByRef str����ʱ��_Out As String, ByRef blnAdd_Out As Boolean, _
                            ByRef blnNotWork_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�Һ���ص�����
    '���:
    '����:lngSn_Out-�������
    '     str����ʱ��_Out-���ط���ʱ��
    '     blnAdd_Out-�Ƿ񷵻صļӺ�
    '     blnNotWork_Out-�Һ�ʱ���Ƿ����Ű�ʱ��(False-��Ч�Һ�ʱ��,True-������)
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-31 14:36:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String, blnAdd As Boolean, strWorkTimeName As String, blnValied As Boolean
    Dim str����ʱ�� As String, lngSN As Long, strWeekName As String
    Dim lng�ƻ�ID As Long, lng����ID As Long
    Dim dtRegDate  As Date
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    
    On Error GoTo errHandle

    str���� = zlGet��ǰ���ڼ�(IIf(mblnAppointment, dtpDate.Value, ""))
  
    
    lng�ƻ�ID = Val(Nvl(mrsPlan!�ƻ�ID))
    lng����ID = Val(Nvl(mrsPlan!ID))
  
    '��ȡ����ʱ��
    blnAdd = False: lngSN = 0
    
    If mblnAppointment Then 'ԤԼ����
        
        dtRegDate = CDate(Format(dtpDate, "yyyy-mm-dd") & " " & Format(dtpTime, "hh:mm:ss"))
        
        str����ʱ�� = Format(dtpDate, "yyyy-mm-dd")
        strWeekName = mobjRegister.zlGetWeekNameFromDate(dtRegDate)
        
        
        If mViewMode = v_ר�Һŷ�ʱ�� Then
            If lng�ƻ�ID <> 0 Then
                    strSql = "" & _
                    " Select Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ, 0 As ��Լ��, Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����," & vbNewLine & _
                    "       Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) As ʧЧʱ��" & vbNewLine & _
                    " From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����," & vbNewLine & _
                    "              To_Date('" & str����ʱ�� & "' || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��," & vbNewLine & _
                    "              To_Date('" & str����ʱ�� & "' || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��," & vbNewLine & _
                    "              Sd.��������, Sd.�Ƿ�ԤԼ" & vbNewLine & _
                    "       From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd" & vbNewLine & _
                    "       Where Jh.Id = Sd.�ƻ�id And Jh.Id = [1] And" & vbNewLine & _
                    "             Sd.���� =[3]) Jh," & vbNewLine & _
                    "     �Һ����״̬ Zt" & vbNewLine & _
                    " Where Zt.����(+) = Jh.��ʼʱ�� And Zt.����(+) = Jh.���� And Jh.��ʼʱ�� = [2] And Zt.���(+) = Jh.��� And Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1" & vbNewLine & _
                    " Order By ���"
                        
            Else
                    strSql = "" & _
                    " Select Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ, 0 As ��Լ��, Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����," & vbNewLine & _
                    "       Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) As ʧЧʱ��" & vbNewLine & _
                    " From (Select Sd.����id, Sd.���, Sd.����, Ap.����," & vbNewLine & _
                    "              To_Date('" & str����ʱ�� & "' ||' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��," & vbNewLine & _
                    "              To_Date('" & str����ʱ�� & "' ||' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��," & vbNewLine & _
                    "              Sd.��������, Sd.�Ƿ�ԤԼ" & vbNewLine & _
                    "       From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd" & vbNewLine & _
                    "       Where Ap.Id = Sd.����id And Ap.Id = [1] And" & vbNewLine & _
                    "             Sd.���� =[3] ) Ap, �Һ����״̬ Zt" & vbNewLine & _
                    " Where Zt.����(+) = Ap.��ʼʱ�� And Zt.����(+) = Ap.���� And Ap.��ʼʱ�� =[2] And Zt.���(+) = Ap.��� And Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1" & vbNewLine & _
                    " Order By ���"
            End If
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(lng�ƻ�ID <> 0, lng�ƻ�ID, lng����ID), dtRegDate, strWeekName)
            
            '124298�����ϴ���2018/4/13��������Ӻ�ʱ���ж�Ȩ��
            If Not rsTmp.EOF Then
                rsTmp.Filter = "ʣ���� <> 0"
            Else
                blnValied = True
            End If
            If rsTmp.RecordCount <> 0 Then
                lngSN = Val(Nvl(rsTmp!���))
            Else
                strSql = "Select Max(���) As ��� From �Һ����״̬ Where ���� = [1] And Trunc(����) = Trunc(" & str����ʱ�� & ")"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, Nvl(mrsPlan!�ű�))
                If rsTmp.RecordCount <> 0 Then lngSN = Val(Nvl(rsTmp!���))
                
                
                If lng�ƻ�ID <> 0 Then
                    strSql = "" & _
                    "   Select Max(���) As ��� From �Һżƻ�ʱ��  " & _
                    "   Where �ƻ�ID = [1] And ���� = [3]"
                Else
                    strSql = "" & _
                    "   Select Max(���) As ��� From �ҺŰ���ʱ�� " & _
                    "   Where ����ID = [1] And ���� =[3]"
                End If
                
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(lng�ƻ�ID <> 0, lng�ƻ�ID, lng����ID), dtRegDate, strWeekName)
                If lngSN = 0 Then
                    If rsTmp.RecordCount <> 0 Then lngSN = Val(Nvl(rsTmp!���))
                Else
                    If Val(Nvl(rsTmp!���)) > lngSN Then lngSN = Val(Nvl(rsTmp!���))
                End If
                
                lngSN = lngSN + 1
                blnAdd = True
            End If
        End If
        If IsNull(mrsPlan.Fields(str����).Value) Then blnValied = True
        
        str����ʱ�� = Format(dtRegDate, "yyyy-mm-dd HH:MM:SS")
        blnAdd_Out = blnAdd
        blnNotWork_Out = blnValied
        str����ʱ��_Out = str����ʱ��
        lngSn_Out = lngSN
        GetRegData = True
        Exit Function
    End If
    
    '�ҺŴ���
    If mViewMode <> v_ר�Һŷ�ʱ�� Or (mViewMode = v_ר�Һŷ�ʱ�� And IsNull(mrsPlan.Fields(str����).Value)) Then
        
        dtRegDate = gobjDatabase.Currentdate
        str����ʱ�� = Format(dtRegDate, "yyyy-mm-dd HH:MM:SS")
        
        If IsNull(mrsPlan.Fields(str����).Value) Then blnValied = True
        
        blnNotWork_Out = blnValied
        str����ʱ��_Out = str����ʱ��
        lngSn_Out = lngSN
        GetRegData = True
        Exit Function
    End If
    
    dtRegDate = gobjDatabase.Currentdate
    str����ʱ�� = Format(dtRegDate, "yyyy-mm-dd HH:MM:SS")
    
    '��鷢ʱ���Ƿ���Ч
    strWorkTimeName = Nvl(mrsPlan.Fields(str����).Value)
    If mobjRegister.zlCheckIsValiedWorkTimeFromWorkTimeName(str����ʱ��, strWorkTimeName, "", "", blnValied) = False Then Exit Function
    
    If blnValied Then   '������
        blnNotWork_Out = blnValied
        str����ʱ��_Out = str����ʱ��
        lngSn_Out = lngSN
        GetRegData = True
        Exit Function
    End If
    
    strWeekName = mobjRegister.zlGetWeekNameFromDate(dtRegDate)
    
    
    If lng�ƻ�ID <> 0 Then
            strSql = "Select Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ, 0 As ��Լ��, Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����," & vbNewLine & _
            "       Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) As ʧЧʱ��" & vbNewLine & _
            "From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����," & vbNewLine & _
            "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��," & vbNewLine & _
            "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��," & vbNewLine & _
            "              Sd.��������, Sd.�Ƿ�ԤԼ" & vbNewLine & _
            "       From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd" & vbNewLine & _
            "       Where Jh.Id = Sd.�ƻ�id And Jh.Id = [1] And" & vbNewLine & _
            "             Sd.���� =[2]) Jh," & vbNewLine & _
            "     �Һ����״̬ Zt" & vbNewLine & _
            "Where Zt.����(+) = Jh.��ʼʱ�� And Zt.����(+) = Jh.���� And Zt.���(+) = Jh.��� And Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1" & vbNewLine & _
            "Order By ���"
    Else
        strSql = "Select Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ, 0 As ��Լ��, Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����," & vbNewLine & _
            "       Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) As ʧЧʱ��" & vbNewLine & _
            "From (Select Sd.����id, Sd.���, Sd.����, Ap.����," & vbNewLine & _
            "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��," & vbNewLine & _
            "              To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��," & vbNewLine & _
            "              Sd.��������, Sd.�Ƿ�ԤԼ" & vbNewLine & _
            "       From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd" & vbNewLine & _
            "       Where Ap.Id = Sd.����id And Ap.Id = [1] And" & vbNewLine & _
            "             Sd.���� =  [2]) Ap, �Һ����״̬ Zt" & vbNewLine & _
            "Where Zt.����(+) = Ap.��ʼʱ�� And Zt.����(+) = Ap.���� And Zt.���(+) = Ap.��� And Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1" & vbNewLine & _
            "Order By ���"
    End If
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(lng�ƻ�ID <> 0, lng�ƻ�ID, lng����ID), strWeekName)
    If Not rsTmp.EOF Then rsTmp.Filter = "ʣ���� <> 0"
    
    'ȡ��С����ʱ���
    If rsTmp.RecordCount <> 0 Then
        lngSN = Val(Nvl(rsTmp!���))
        str����ʱ�� = Format(dtRegDate, "yyyy-mm-dd") & " " & Format(Nvl(rsTmp!��ʼʱ��), "hh:mm:ss")
    End If
     
    blnAdd_Out = blnAdd
    str����ʱ��_Out = str����ʱ��
    lngSn_Out = lngSN
    GetRegData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function Get���ʽ����(ByVal str���ʽ As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ʽ����
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-31 18:38:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rs���ʽ As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSql = "Select ���� From ҽ�Ƹ��ʽ Where ���� = [1]"
    Set rs���ʽ = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, str���ʽ)
    If rs���ʽ.RecordCount <> 0 Then
        Get���ʽ���� = Nvl(rs���ʽ!����)
    Else
        strSql = "Select ���� From ҽ�Ƹ��ʽ Where ȱʡ��־ = 1"
        Set rs���ʽ = gobjDatabase.OpenSQLRecord(strSql, App.ProductName)
        If rs���ʽ.RecordCount <> 0 Then
            Get���ʽ���� = Nvl(rs���ʽ!����)
        End If
    End If
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Sub PrintInvoic(ByVal strNO As String, ByVal strFactNO As String, ByVal dat�Ǽ�ʱ�� As Date)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ��Ʊ
    '���:strNo-���ݺ�
    '     strFactNo-��Ʊ��
    '����:���˺�
    '����:2018-02-01 11:12:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNotValiedNos As String
    'If mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ�� Then Exit Sub:��Ӧ���ж�(���˺�)RePrint:
RePrint:
    Load frmPrint
    Call frmPrint.ReportPrint(1, strNO, "", mlng����ID, mlng�Һ�ID, strFactNO, dat�Ǽ�ʱ��, , , , mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��, False, mstrUseType)
    If Not gblnBill�Һ� Then Exit Sub
    
    If zlIsNotSucceedPrintBill(4, strNO, strNotValiedNos) = True Then
        If MsgBox(IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "����Ϊ[" & strNotValiedNos & "]Ʊ�ݴ�ӡδ�ɹ�,�Ƿ����½���Ʊ�ݴ�ӡ!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
         Exit Sub
    End If
    
End Sub

Private Function SaveData_Cash() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������
    '���:
    '����:
    '����:����ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-31 11:13:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim i As Long, strFactNO As String, blnBalance As Boolean
    Dim blnProofPrint As Boolean  'ƾ����ӡ
    Dim blnInvoiceIsPrint  As Boolean '��Ʊ��ӡ
    Dim blnDepositBillIsPrint As Boolean 'ԤԼ�Һŵ���ӡ
    Dim curԤ�� As Currency, cur���� As Currency, cur�ֽ� As Currency
    Dim dat�Ǽ�ʱ�� As Date, str�Ǽ�ʱ�� As String, dt����ʱ�� As Date
    Dim str����ʱ�� As String, lngSN As Long, blnAdd As Boolean, blnNotWork As Boolean, lngValue As Long
    Dim lngҽ�ƿ����ID As Long, bln���ѿ� As Boolean, blnNoDoc As Boolean, str���㷽ʽ As String
    Dim blnTrans As Boolean, lng����ID As Long
    Dim cllProAfter As Collection, cllPro As Collection
    Dim str������ˮ�� As String, str����˵��   As String
    Dim strNO As String, blnOneCard As Boolean
    Dim strSql As String, blnNotCommit As Boolean, strAdvance As String
    Dim cllTheeSwap As Collection, cllTheeSwapOther As Collection
    
    On Error GoTo errHandle
    'bytMode-0-����;1-����;2-����
    If CheckDataIsValied(0) = False Then Exit Function
    
    blnProofPrint = GetPrintProofIsPrint(0) 'ƾ����ӡ
    blnInvoiceIsPrint = GetInvoiceIsPrint(0)    '��Ʊ��ӡ
    blnDepositBillIsPrint = GetDepositBillIsPrint(0)   'ԤԼ�Һŵ���ӡ
    
    'ȷ����Ʊ��
    If blnInvoiceIsPrint Or (mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��) Then
        If RefreshFact(strFactNO) = False Then Exit Function
    End If
    
    blnBalance = False
    
    If mblnAppointment = False Or mblnAppointment And mty_Para.blnԤԼʱ�տ� Then
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
    If Val(curԤ��) <> 0 Then
        If Not gobjDatabase.PatiIdentify(Me, glngSys, Nvl(mrsInfo!����ID), Val(curԤ��), mlngModul, 1, , _
                             IIf(-1 * mty_Para.dblԤ��������鿨 >= Val(curԤ��), False, True), True, mstr����IDs, (mty_Para.dblԤ��������鿨 <> 0), (mty_Para.dblԤ��������鿨 = 2)) Then Exit Function
    End If
    
    
    ReadRegistPrice Val(Nvl(mrsPlan!��ĿID)), chkBook.Value = 1, False, mstr�ѱ�, rsItems, rsIncomes, _
        Nvl(mrsInfo!����ID), mintInsure, txtReg.Tag, IIf(mblnAppointment, 1, 0), , mstrPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
        
    '�������
    str���㷽ʽ = ""
    If blnBalance Then
        For i = 1 To mcolCardPayMode.Count
            If cboPayMode.Text = mcolCardPayMode.Item(i)(1) Then
                lngҽ�ƿ����ID = mcolCardPayMode.Item(i)(3)
                bln���ѿ� = Val(mcolCardPayMode.Item(i)(5)) = 1
                str���㷽ʽ = mcolCardPayMode.Item(i)(6)
            End If
        Next i
        If CheckBrushCard(Val(cur�ֽ�), lngҽ�ƿ����ID, bln���ѿ�, rsItems, rsIncomes) = False Then Exit Function
        If str���㷽ʽ = "" Then str���㷽ʽ = cboPayMode.Text
    End If

    dat�Ǽ�ʱ�� = gobjDatabase.Currentdate
    str�Ǽ�ʱ�� = Format(dat�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
    
    '124431:���ϴ�,2018/5/17��������ؼ��
    lngValue = Val(Split(lblSn.Caption & ":", ":")(1))
    If lngValue <> 0 Then
        If mblnAppointment Then
            dt����ʱ�� = CDate(Format(dtpDate, "yyyy-mm-dd"))
        Else
            dt����ʱ�� = CDate(Format(gobjDatabase.Currentdate, "yyyy-mm-dd"))
        End If
        'ҽ��վ���ж����ź�Ԥ��
        strSql = "Select 1 From �Һ����״̬ " & _
                "  Where ���� = [1] And Trunc(����) = [2] And ��� = [3]" & vbNewLine & _
                IIf(mty_Para.bln�˺�����, " And ״̬ <> 4", "")
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "��ż��", Nvl(mrsPlan!�ű�), dt����ʱ��, lngValue)
        If rsTemp.RecordCount > 1 Then
            If MsgBox("���� " & lngValue & " �ѱ�����ʹ�ã��Ƿ��Զ���ȡ����ʱ�ν���" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    '124298�����ϴ���2018/4/13��������Ӻ�ʱ���ж�Ȩ��
    If GetRegData(lngSN, str����ʱ��, blnAdd, blnNotWork) = False Then Exit Function
    
    '137272:���ϴ�,2019/2/20,�Һ�����
    If mViewMode = v_ר�Һ� Or mViewMode = v_ר�Һŷ�ʱ�� Then
        If ReserveRegNo(Nvl(mrsPlan!�ű�), True, mViewMode = v_ר�Һŷ�ʱ��, str����ʱ��, lngSN, "ҽ��վ����") = False Then Exit Function
    End If
    
    If blnAdd = False And Not blnNotWork Then
        If Val(Nvl(mrsPlan!�ѹ�)) >= Val(Nvl(mrsPlan!�޺�)) And Val(Nvl(mrsPlan!�޺�)) <> 0 Then
            blnAdd = True
        End If
        If Val(Nvl(mrsPlan!��Լ)) >= Val(Nvl(mrsPlan!��Լ)) And Val(Nvl(mrsPlan!��Լ)) <> 0 And mblnAppointment Then
            blnAdd = True
        End If
    End If
    
    If blnAdd And InStr(gstrPrivs, ";�Ӻ�;") = 0 Then
        MsgBox "��û�мӺ�Ȩ�ޣ��޷��Ե�ǰ�ű����" & IIf(gSysPara.bln��Һ�ģʽ And mblnAppointment = False, "����", "�Һ�") & "��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If blnAdd = False Then
        If InStr(gstrPrivs, ";�Ӻ�;") > 0 Then blnAdd = True
    End If
    If blnNotWork Then blnAdd = blnNotWork
    
    blnOneCard = zlOldOneCardIsStart(cboPayMode.Text)
    If GetSaveRegDataSQL(0, rsItems, rsIncomes, False, str���㷽ʽ, curԤ��, cur����, cur�ֽ�, lngSN, blnAdd, str����ʱ��, str�Ǽ�ʱ��, cllPro, cllProAfter, lng����ID, strNO, "", lngҽ�ƿ����ID, bln���ѿ�, mstrCardNO, str������ˮ��, str����˵��) = False Then Exit Function
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True, False

    blnTrans = True
    If blnOneCard And lngҽ�ƿ����ID <> 0 And cur�ֽ� <> 0 Then
        If Not mobjICCard.PaymentSwap(Val(cur�ֽ�), Val(cur�ֽ�), Val(lngҽ�ƿ����ID), 0, mstrCardNO, "", lng����ID, Nvl(mrsInfo!����ID)) Then
            gcnOracle.RollbackTrans
            MsgBox "һ��ͨ����Һŷ�ʧ��", vbInformation, gstrSysName
            Exit Function
        End If
        strSql = "zl_һ��ͨ����_Update(" & lng����ID & ",'" & cboPayMode.Text & "','" & mstrCardNO & "','" & lngҽ�ƿ����ID & "','" & "" & "'," & cur�ֽ� & ")"
        Call gobjDatabase.ExecuteProcedure(strSql, Me.Caption)
    End If

    
    
    blnNotCommit = False
    If mintInsure <> 0 And mstrYBPati <> "" And cur���� <> 0 Then
        '68991:strAdvance:����ģʽ(0��1)|�Һŷ���ȡ��ʽ(0��1) |�Һŵ���
        strAdvance = ""
        If mPatiChargeMode = EM_�����ƺ���� Then
            strAdvance = IIf(mPatiChargeMode = EM_�����ƺ����, "1", "0")
            strAdvance = strAdvance & "|" & IIf(mRegistFeeMode = EM_RG_����, "1", "0")
            strAdvance = strAdvance & "|" & strNO
        End If
        If Not gclsInsure.RegistSwap(lng����ID, cur����, mintInsure, strAdvance) Then
            gcnOracle.RollbackTrans:   Exit Function
        End If
        blnNotCommit = True
    End If
        
        
    '����:31187 ����ҽ���ɹ���,�����һЩ���ݸ���:�ڲ������������ύ���,���Բ�����д
    zlExecuteProcedureArrAy cllProAfter, Me.Caption, False, False
    
    Set cllTheeSwap = New Collection: Set cllTheeSwapOther = New Collection
    If Not blnOneCard And Not mPatiChargeMode = EM_�����ƺ���� And cur�ֽ� <> 0 Then
        If zlInterfacePrayMoney(Me, mlngModul, lng����ID, cllTheeSwap, cllTheeSwapOther, Val(cur�ֽ�), mstrCardNO, lngҽ�ƿ����ID, bln���ѿ�) = False Then gcnOracle.RollbackTrans: Exit Function
        '������������
        zlExecuteProcedureArrAy cllTheeSwap, Me.Caption, False, False
    End If
    
    
    Err = 0: On Error GoTo OthersCommit:
    zlExecuteProcedureArrAy cllTheeSwapOther, Me.Caption, False, False
        
OthersCommit:
    gcnOracle.CommitTrans: blnTrans = False
    
    If mintInsure > 0 Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistSwap, True, mintInsure)
    
    '145198:���ϴ�,2019/12/26,�Һųɹ��������ҽӿڣ�Ŀǰ����ԤԼ�����֧����ά��
    Call zlSaveRgstAfterByPlugIn(mlngModul, strNO, (mblnAppointment And Not mty_Para.blnԤԼʱ�տ�))
    If blnInvoiceIsPrint Then
        Call PrintInvoic(strNO, strFactNO, dat�Ǽ�ʱ��)
    End If
    If blnProofPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
        '��¼��ӡ��ƾ��
        gstrSQL = "Zl_ƾ����ӡ��¼_Update(4,'" & strNO & "',1,'" & UserInfo.���� & "')"
        gobjDatabase.ExecuteProcedure gstrSQL, ""
    End If
    
    If blnDepositBillIsPrint Then Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me, "NO=" & strNO, 2)
    mstrNO = strNO
    
    SaveData_Cash = True
    Exit Function
errHandle:
    If mintInsure > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistSwap, False, mintInsure)
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function




Private Function SaveData_Accounting() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������
    '���:
    '����:
    '����:����ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-31 11:13:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset
    Dim i As Long
    Dim blnProofPrint As Boolean  'ƾ����ӡ
    Dim blnDepositBillIsPrint  As Boolean   'ԤԼ�Һŵ���ӡ
    Dim dat�Ǽ�ʱ�� As Date, str�Ǽ�ʱ�� As String
    Dim str����ʱ�� As String, lngSN As Long, blnAdd As Boolean, blnNotWork As Boolean
    Dim blnTrans As Boolean, lng����ID As Long
    Dim cllProAfter As Collection, cllPro As Collection
    Dim curԤ�� As Currency, cur���� As Currency, cur�ֽ� As Currency
    Dim str���㷽ʽ As String, blnBalance As Boolean
    Dim lngҽ�ƿ����ID As Long, bln���ѿ� As Boolean, blnNoDoc As Boolean
    Dim strNO As String, str������ˮ�� As String, str����˵�� As String
    Dim blnNotCommit As Boolean, strAdvance As String
    
    
    On Error GoTo errHandle
    
    'bytMode-0-����;1-����;2-����
    If CheckDataIsValied(1) = False Then Exit Function
    blnProofPrint = GetPrintProofIsPrint(1) 'ƾ����ӡ
    blnDepositBillIsPrint = GetDepositBillIsPrint(1)   'ԤԼ�Һŵ���ӡ
    

    ReadRegistPrice Val(Nvl(mrsPlan!��ĿID)), chkBook.Value = 1, False, mstr�ѱ�, rsItems, rsIncomes, _
        Nvl(mrsInfo!����ID), mintInsure, txtReg.Tag, IIf(mblnAppointment, 1, 0), , mstrPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
       
       
    dat�Ǽ�ʱ�� = gobjDatabase.Currentdate
    str�Ǽ�ʱ�� = Format(dat�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
        
    '124298�����ϴ���2018/4/13��������Ӻ�ʱ���ж�Ȩ��
    If GetRegData(lngSN, str����ʱ��, blnAdd, blnNotWork) = False Then Exit Function
    
    '137272:���ϴ�,2019/2/20,�Һ�����
    If mViewMode = v_ר�Һ� Or mViewMode = v_ר�Һŷ�ʱ�� Then
        If ReserveRegNo(Nvl(mrsPlan!�ű�), True, mViewMode = v_ר�Һŷ�ʱ��, str����ʱ��, lngSN, "ҽ��վ����") = False Then Exit Function
    End If
    
    If blnAdd = False And Not blnNotWork Then
        If Val(Nvl(mrsPlan!�ѹ�)) >= Val(Nvl(mrsPlan!�޺�)) And Val(Nvl(mrsPlan!�޺�)) <> 0 Then
            blnAdd = True
        End If
        If Val(Nvl(mrsPlan!��Լ)) >= Val(Nvl(mrsPlan!��Լ)) And Val(Nvl(mrsPlan!��Լ)) <> 0 And mblnAppointment Then
            blnAdd = True
        End If
    End If
    
    If blnAdd And InStr(gstrPrivs, ";�Ӻ�;") = 0 Then
        MsgBox "��û�мӺ�Ȩ�ޣ��޷��Ե�ǰ�ű����" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If blnAdd = False Then
        If InStr(gstrPrivs, ";�Ӻ�;") > 0 Then blnAdd = True
    End If
    If blnNotWork Then blnAdd = blnNotWork
  
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
    If Val(curԤ��) <> 0 Then
        If Not gobjDatabase.PatiIdentify(Me, glngSys, Nvl(mrsInfo!����ID), Val(curԤ��), mlngModul, 1, , _
                           IIf(-1 * mty_Para.dblԤ��������鿨 >= Val(curԤ��), False, True), True, mstr����IDs, (mty_Para.dblԤ��������鿨 <> 0), (mty_Para.dblԤ��������鿨 = 2)) Then Exit Function
    End If
    
    If blnBalance Then
        For i = 1 To mcolCardPayMode.Count
            If cboPayMode.Text = mcolCardPayMode.Item(i)(1) Then
                lngҽ�ƿ����ID = mcolCardPayMode.Item(i)(3)
                bln���ѿ� = Val(mcolCardPayMode.Item(i)(5)) = 1
                str���㷽ʽ = mcolCardPayMode.Item(i)(6)
            End If
        Next i
        If CheckBrushCard(Val(cur�ֽ�), lngҽ�ƿ����ID, bln���ѿ�, rsItems, rsIncomes) = False Then Exit Function
        If str���㷽ʽ = "" Then str���㷽ʽ = str���㷽ʽ = cboPayMode.Text
    End If
     
    If GetSaveRegDataSQL(1, rsItems, rsIncomes, False, str���㷽ʽ, curԤ��, cur����, cur�ֽ�, lngSN, blnAdd, str����ʱ��, str�Ǽ�ʱ��, cllPro, cllProAfter, lng����ID, strNO, "", lngҽ�ƿ����ID, bln���ѿ�, mstrCardNO, str������ˮ��, str����˵��) = False Then Exit Function

    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True, False
    
    
    blnNotCommit = False
    If mintInsure <> 0 And mstrYBPati <> "" And cur���� <> 0 Then
        '68991:strAdvance:����ģʽ(0��1)|�Һŷ���ȡ��ʽ(0��1) |�Һŵ���
        strAdvance = IIf(mPatiChargeMode = EM_�����ƺ����, "1", "0")
        strAdvance = strAdvance & "|" & "1"
        strAdvance = strAdvance & "|" & strNO
        If Not gclsInsure.RegistSwap(lng����ID, cur����, mintInsure, strAdvance) Then
            gcnOracle.RollbackTrans:   Exit Function
        End If
        blnNotCommit = True
    End If
        
    zlExecuteProcedureArrAy cllProAfter, Me.Caption, False, False
    gcnOracle.CommitTrans: blnTrans = False
    If mintInsure > 0 Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistSwap, True, mintInsure)
    
    '145198:���ϴ�,2019/12/26,�Һųɹ��������ҽӿڣ�Ŀǰ����ԤԼ�����֧����ά��
    Call zlSaveRgstAfterByPlugIn(mlngModul, strNO, (mblnAppointment And Not mty_Para.blnԤԼʱ�տ�))
    If blnProofPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
        '��¼��ӡ��ƾ��
        gstrSQL = "Zl_ƾ����ӡ��¼_Update(4,'" & strNO & "',1,'" & UserInfo.���� & "')"
        gobjDatabase.ExecuteProcedure gstrSQL, ""
    End If
    If blnDepositBillIsPrint Then Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me, "NO=" & strNO, 2)
    mstrNO = strNO
    SaveData_Accounting = True
    Exit Function
errHandle:
    If mintInsure > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistSwap, False, mintInsure)
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function


Private Function SaveData_Price() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ϊ���۵�
    '���:
    '����:
    '����:����ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-31 11:41:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset
    Dim blnProofPrint As Boolean  'ƾ����ӡ
    Dim blnDepositBillIsPrint As Boolean    'ԤԼ�Һŵ���ӡ
    Dim dat�Ǽ�ʱ�� As Date, str�Ǽ�ʱ�� As String
    Dim str����ʱ�� As String, lngSN As Long, blnAdd As Boolean, blnNotWork As Boolean
    Dim str����NO As String, blnTrans As Boolean, lng����ID As Long, strNO As String
    Dim cllProAfter As Collection, cllPro As Collection
    
    On Error GoTo errHandle
    
    'bytMode-0-����;1-����;2-����
    If CheckDataIsValied(2) = False Then Exit Function
    blnProofPrint = GetPrintProofIsPrint(2) 'ƾ����ӡ
    blnDepositBillIsPrint = GetDepositBillIsPrint(2)   'ԤԼ�Һŵ���ӡ
    

    ReadRegistPrice Val(Nvl(mrsPlan!��ĿID)), chkBook.Value = 1, False, mstr�ѱ�, rsItems, rsIncomes, _
     Nvl(mrsInfo!����ID), mintInsure, txtReg.Tag, IIf(mblnAppointment, 1, 0), , mstrPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
     
     
    dat�Ǽ�ʱ�� = gobjDatabase.Currentdate
    str�Ǽ�ʱ�� = Format(dat�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
    
    '124298�����ϴ���2018/4/13��������Ӻ�ʱ���ж�Ȩ��
    If GetRegData(lngSN, str����ʱ��, blnAdd, blnNotWork) = False Then Exit Function
    
    '137272:���ϴ�,2019/2/20,�Һ�����
    If mViewMode = v_ר�Һ� Or mViewMode = v_ר�Һŷ�ʱ�� Then
        If ReserveRegNo(Nvl(mrsPlan!�ű�), True, mViewMode = v_ר�Һŷ�ʱ��, str����ʱ��, lngSN, "ҽ��վ����") = False Then Exit Function
    End If
    
    If blnAdd = False And Not blnNotWork Then
        If Val(Nvl(mrsPlan!�ѹ�)) >= Val(Nvl(mrsPlan!�޺�)) And Val(Nvl(mrsPlan!�޺�)) <> 0 Then
            blnAdd = True
        End If
        If Val(Nvl(mrsPlan!��Լ)) >= Val(Nvl(mrsPlan!��Լ)) And Val(Nvl(mrsPlan!��Լ)) <> 0 And mblnAppointment Then
            blnAdd = True
        End If
    End If
    
    If blnAdd And InStr(gstrPrivs, ";�Ӻ�;") = 0 Then
        MsgBox "��û�мӺ�Ȩ�ޣ��޷��Ե�ǰ�ű����" & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If blnAdd = False Then
        If InStr(gstrPrivs, ";�Ӻ�;") > 0 Then blnAdd = True
    End If
    If blnNotWork Then blnAdd = blnNotWork
    
    mlngSN = lngSN
    
    If GetSaveRegDataSQL(2, rsItems, rsIncomes, False, "", 0, 0, 0, lngSN, blnAdd, str����ʱ��, str�Ǽ�ʱ��, cllPro, cllProAfter, lng����ID, strNO, str����NO) = False Then Exit Function
    
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True, False
    zlExecuteProcedureArrAy cllProAfter, Me.Caption, False, False
    gcnOracle.CommitTrans: blnTrans = False
    
    '145198:���ϴ�,2019/12/26,�Һųɹ��������ҽӿڣ�Ŀǰ����ԤԼ�����֧����ά��
    Call zlSaveRgstAfterByPlugIn(mlngModul, strNO, (mblnAppointment And Not mty_Para.blnԤԼʱ�տ�))
    If blnProofPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
        '��¼��ӡ��ƾ��
        gstrSQL = "Zl_ƾ����ӡ��¼_Update(4,'" & strNO & "',1,'" & UserInfo.���� & "')"
        gobjDatabase.ExecuteProcedure gstrSQL, ""
    End If
    If blnDepositBillIsPrint Then Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me, "NO=" & strNO, 2)
    mstrNO = strNO
    SaveData_Price = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function
 


Private Function GetSaveRegDataSQL(ByVal bytMode As Byte, ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset, ByVal blnInvoicePrint As Boolean, _
    ByVal str���㷽ʽ As String, ByVal dbl��Ԥ�� As Double, ByVal dbl�����ʻ� As Double, ByVal dbl�ֽ� As Double, ByVal lngSN As Long, ByVal blnAddNum As Boolean, _
    ByVal str����ʱ�� As String, ByVal str�Ǽ�ʱ�� As String, ByRef cllPro_out As Collection, ByRef cllProAffter_out As Collection, _
    Optional lng����ID_Out As Long, Optional strNO_Out As String, Optional strPriceNo_Out As String, _
    Optional ByVal lng֧�����ID As Long = 0, Optional ByVal bln���ѿ� As Boolean = False, Optional ByVal str���� As String = "", Optional ByVal str������ˮ�� As String = "", Optional ByVal str����˵�� As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�Һű�������
    '���:bytMode-0-����;1-����;2-����
    '     rsItem-��Ŀ��
    '     rsInComes-������Ŀ��
    '     blnInvoicePrint-�Ƿ�Ʊ��ӡ
    '     strBalances-���㷽ʽ
    '     blnAddNum:�Ƿ�Ӻ�
    '����:cllPro_out-�������ݱ��漯
    '     cllProAffter_out-���غ�ִ�е�SQL��
    '     lng����ID_Out_Out-����ID
    '     strNO_out-���ݺ�
    '     strPriceNo_Out-���۵���
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-31 17:46:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, k As Long, i As Long, j As Long, int�۸񸸺� As Integer
    Dim lng�Һſ���ID As Long, byt���� As Byte
    Dim lngҽ��ID As Long, strҽ������ As String
    Dim str���ʽ���� As String
    Dim dblTotal As Double, blnNoDoc As Boolean
    
    
    On Error GoTo errHandle
    strPriceNo_Out = ""
    
    rsItems.Filter = ""
    strҽ������ = NeedName(cboDoctor.Text)
    If cboDoctor.ListCount = 0 Then
        lngҽ��ID = 0
    Else
        lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    
    str���ʽ���� = Get���ʽ����(Nvl(mrsInfo!ҽ�Ƹ��ʽ))
    
    lng�Һſ���ID = Val(Nvl(mrsPlan!����ID))
    byt���� = IIf(chk����.Value = 1, 1, 0)
    
    If bytMode = 2 Then '��Ϊ���۵�
        dblTotal = GetRegistMoney(True, False)
        '�ҺŷѴ�Ϊ���ұ���Ϊ���۵����Ų�������NO
        If dblTotal <> 0 Then strPriceNo_Out = gobjDatabase.GetNextNo(13)
    End If
    
    If cllPro_out Is Nothing Then Set cllPro_out = New Collection
    If cllProAffter_out Is Nothing Then Set cllProAffter_out = New Collection
    
    lng����ID_Out = 0
    If bytMode <> 1 And (Not mblnAppointment Or (mblnAppointment And mty_Para.blnԤԼʱ�տ�)) Then      'ԤԼ�տ�
        lng����ID_Out = gobjDatabase.GetNextId("���˽��ʼ�¼")
    End If
    strNO_Out = gobjDatabase.GetNextNo(12)
    k = 1: rsItems.MoveFirst
    For i = 1 To rsItems.RecordCount
        int�۸񸸺� = k
        rsIncomes.Filter = "��ĿID=" & rsItems!��ĿID
        For j = 1 To rsIncomes.RecordCount
        
            strSql = "zl_���˹Һż�¼_INSERT("
            '  ����id_In        ������ü�¼.����id%Type,
            strSql = strSql & "" & ZVal(Nvl(mrsInfo!����ID)) & ","
            '  �����_In        ������ü�¼.��ʶ��%Type,
            strSql = strSql & "" & IIf(mstr����� = "", "NULL", mstr�����) & ","
            '  ����_In          ������ü�¼.����%Type,
            strSql = strSql & "'" & txtPatient.Text & "',"
            '  �Ա�_In          ������ü�¼.�Ա�%Type,
            strSql = strSql & "'" & mstr�Ա� & "',"
            '  ����_In          ������ü�¼.����%Type,
            strSql = strSql & "'" & mstrAge & "',"
            '  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
            strSql = strSql & "'" & str���ʽ���� & "',"
            '  �ѱ�_In          ������ü�¼.�ѱ�%Type,
            strSql = strSql & "'" & mstr�ѱ� & "',"
            '  ���ݺ�_In        ������ü�¼.No%Type,
            strSql = strSql & "'" & strNO_Out & "',"
            '  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
            strSql = strSql & "'" & IIf(blnInvoicePrint = False, "", "") & "',"
            '  ���_In          ������ü�¼.���%Type,
            strSql = strSql & "" & k & ","
            '  �۸񸸺�_In      ������ü�¼.�۸񸸺�%Type,
            strSql = strSql & "" & IIf(int�۸񸸺� = k, "NULL", int�۸񸸺�) & ","
            '  ��������_In      ������ü�¼.��������%Type,
            strSql = strSql & "" & IIf(rsItems!���� = 2, 1, "NULL") & ","
            '  �շ����_In      ������ü�¼.�շ����%Type,
            strSql = strSql & "'" & rsItems!��� & "',"
            '  �շ�ϸĿid_In    ������ü�¼.�շ�ϸĿid%Type,
            strSql = strSql & "" & rsItems!��ĿID & ","
            '  ����_In          ������ü�¼.����%Type,
            strSql = strSql & "" & rsItems!���� & ","
            '  ��׼����_In      ������ü�¼.��׼����%Type,
            strSql = strSql & "" & rsIncomes!���� & ","
            '  ������Ŀid_In    ������ü�¼.������Ŀid%Type,
            strSql = strSql & "" & rsIncomes!������ĿID & ","
            '  �վݷ�Ŀ_In      ������ü�¼.�վݷ�Ŀ%Type,
            strSql = strSql & "'" & rsIncomes!�վݷ�Ŀ & "',"
            '  ���㷽ʽ_In      ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ�������
            strSql = strSql & "'" & str���㷽ʽ & "',"
            '  Ӧ�ս��_In      ������ü�¼.Ӧ�ս��%Type,
            strSql = strSql & "" & IIf(bytMode = 2 And dblTotal <> 0, 0, Val(Nvl(rsIncomes!Ӧ��))) & ","
            '  ʵ�ս��_In      ������ü�¼.ʵ�ս��%Type,
            strSql = strSql & "" & IIf(bytMode = 2, 0, Val(Nvl(rsIncomes!ʵ��))) & ","
            '  ���˿���id_In    ������ü�¼.���˿���id%Type,
            strSql = strSql & "" & lng�Һſ���ID & ","
            '  ��������id_In    ������ü�¼.��������id%Type,
            strSql = strSql & "" & lng�Һſ���ID & ","
            '  ִ�в���id_In    ������ü�¼.ִ�в���id%Type,
            strSql = strSql & "" & IIf(rsItems!ִ�п���ID = 0, lng�Һſ���ID, rsItems!ִ�п���ID) & ","
            '  ����Ա���_In    ������ü�¼.����Ա���%Type,
            strSql = strSql & "'" & UserInfo.��� & "',"
            '  ����Ա����_In    ������ü�¼.����Ա����%Type,
            strSql = strSql & "'" & UserInfo.���� & "',"
            '  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
            strSql = strSql & "to_date('" & str����ʱ�� & "','yyyy-mm-dd hh24:mi:ss'),"
            '  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
            strSql = strSql & "to_date('" & str�Ǽ�ʱ�� & "','yyyy-mm-dd hh24:mi:ss'),"
            '  ҽ������_In      �ҺŰ���.ҽ������%Type,
            strSql = strSql & "'" & strҽ������ & "',"
            '  ҽ��id_In        �ҺŰ���.ҽ��id%Type,
            strSql = strSql & "" & ZVal(lngҽ��ID) & ","
            '  ������_In Number, --������¼�Ƿ���������
            strSql = strSql & "" & IIf(rsItems!���� = 3, 1, IIf(rsItems!���� = 4, 2, 0)) & ","
            '  ����_In          Number,
            strSql = strSql & "" & IIf(lbl��.Visible, 1, 0) & ","
            '  �ű�_In          �ҺŰ���.����%Type,
            strSql = strSql & "'" & txtReg.Tag & "',"
            '  ����_In          ������ü�¼.��ҩ����%Type,
            strSql = strSql & "'" & IIf(strҽ������ = UserInfo.����, lblRoomName.Caption, "") & "',"
            '  ����id_In        ������ü�¼.����id%Type,
            strSql = strSql & "" & ZVal(lng����ID_Out) & ","
            '  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
            strSql = strSql & "" & IIf(blnInvoicePrint = False, "NULL", ZVal(mlng����ID)) & ","
            '  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --ˢ���Һ�ʱʹ�õ�Ԥ�����,���Ϊ1����.
            strSql = strSql & "" & ZVal(IIf(k = 1, dbl��Ԥ��, 0)) & ","
            '  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����.
            strSql = strSql & "" & ZVal(IIf(k = 1, dbl�ֽ�, 0)) & ","
            '  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����,,���Ϊ1����.
            strSql = strSql & "" & ZVal(IIf(k = 1, dbl�����ʻ�, 0)) & ","
            '  ���մ���id_In    ������ü�¼.���մ���id%Type,
            strSql = strSql & "" & ZVal(Nvl(rsItems!���մ���ID, 0)) & ","
            '  ������Ŀ��_In    ������ü�¼.������Ŀ��%Type,
            strSql = strSql & "" & ZVal(Nvl(rsItems!������Ŀ��, 0)) & ","
            '  ͳ����_In      ������ü�¼.ͳ����%Type,
            strSql = strSql & "" & ZVal(Nvl(rsIncomes!ͳ����, 0)) & ","
            '  ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
            strSql = strSql & "'" & Trim(cboRemark.Text) & "',"
            '  ԤԼ�Һ�_In      Number := 0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز���
            strSql = strSql & "" & IIf(mblnAppointment, IIf(mty_Para.blnԤԼʱ�տ�, 0, 1), 0) & ","
            '  �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
            strSql = strSql & "" & IIf(mty_Para.bln�����շ�Ʊ��, 1, 0) & ","
            '  ���ձ���_In      ������ü�¼.���ձ���%Type,
            strSql = strSql & "'" & rsItems!���ձ��� & "',"
            '  ����_In          ���˹Һż�¼.����%Type := 0,
            strSql = strSql & "" & byt���� & ","
            '  ����_In          �Һ����״̬.���%Type := Null, --ԤԼʱ������ü�¼�ķ�ҩ�����ֶ�,�Һ�ʱ����Һż�¼
            strSql = strSql & "" & ZVal(lngSN) & ","
            '  ����_In          ���˹Һż�¼.����%Type := Null,
            strSql = strSql & "" & "NULL" & ","
            '  ԤԼ����_In      Number := 0,
            strSql = strSql & "" & IIf(mblnAppointment, 1, 0) & ","
            '  ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
            strSql = strSql & "'" & IIf(cboAppointStyle.Visible, cboAppointStyle.Text, "") & "',"
            '  ���ɶ���_In      Number := 0,
            strSql = strSql & "" & 0 & ","
            '  �����id_In      ����Ԥ����¼.�����id%Type := Null,
            strSql = strSql & "" & IIf(lng֧�����ID <> 0 And bln���ѿ� = False, lng֧�����ID, "NULL") & ","
            '  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
            strSql = strSql & "" & IIf(lng֧�����ID <> 0 And bln���ѿ�, lng֧�����ID, "NULL") & ","
            '  ����_In          ����Ԥ����¼.����%Type := Null,
            strSql = strSql & "'" & mstrCardNO & "',"
            '  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
            strSql = strSql & "'" & str������ˮ�� & "',"
            '  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
            strSql = strSql & "'" & str����˵�� & "',"
            '  ������λ_In      ����Ԥ����¼.������λ%Type := Null,
            strSql = strSql & " NULL,"
            '  ��������_In      Number := 0,
            strSql = strSql & IIf(blnAddNum, 1, 0) & ","
            '  ����_In          ���˹Һż�¼.����%Type := Null,
            strSql = strSql & IIf(mintInsure = 0, "NULL", mintInsure) & ","
            '  ����ģʽ_In      Number := 0,
            strSql = strSql & IIf(mPatiChargeMode = EM_�����ƺ����, 1, 0) & ","
            '  ���ʷ���_In      Number := 0,
            strSql = strSql & IIf(bytMode = 1, 1, 0) & ","
            '  �˺�����_In      Number := 1,
            strSql = strSql & IIf(mty_Para.bln�˺�����, 1, 0) & ","
            '  ��Ԥ������ids_In Varchar2 := Null,
            strSql = strSql & "'" & Nvl(mrsInfo!����ID) & "," & mstr����IDs & "',"
            '  �������˷ѱ�_In  Number := 0,
            strSql = strSql & "" & IIf(mblnChangeFeeType, 1, 0) & ","
            '  ������������_In  Number := 0,
            strSql = strSql & "" & IIf(mblnUpdateAge, 1, 0) & ","
            '  �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null,
            strSql = strSql & "'" & strPriceNo_Out & "')"
            '  ���½������_In Number:=1
            
            Call zlAddArray(cllPro_out, strSql)
            
            '����:31187:���ҺŻ��ܵ�������
            If txtReg.Tag <> "" And k = 1 Then
                If Nvl(mrsPlan!ҽ��) = "" Then blnNoDoc = True
                strSql = "zl_���˹ҺŻ���_Update("
                '  ҽ������_In   �ҺŰ���.ҽ������%Type,
                strSql = strSql & IIf(blnNoDoc, "Null,", "'" & strҽ������ & "',")
                '  ҽ��id_In     �ҺŰ���.ҽ��id%Type,
                strSql = strSql & "" & IIf(blnNoDoc, "0,", ZVal(lngҽ��ID) & ",")
                '  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
                strSql = strSql & "" & Val(Nvl(rsItems!��ĿID)) & ","
                '  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
                strSql = strSql & "" & IIf(Val(Nvl(rsItems!ִ�п���ID)) = 0, lng�Һſ���ID, Val(Nvl(rsItems!ִ�п���ID))) & ","
                '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
                strSql = strSql & "to_date('" & str����ʱ�� & "','yyyy-mm-dd hh24:mi:ss'),"
       
                '  ԤԼ��־_In   Number := 0  --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����,3-�շ�ԤԼ
                strSql = strSql & IIf(mblnAppointment, IIf(mty_Para.blnԤԼʱ�տ�, 3, 1), 0) & ","
                '  ����_In       �ҺŰ���.����%Type := Null
                strSql = strSql & "'" & txtReg.Tag & "')"
                Call zlAddArray(cllProAffter_out, strSql)
            End If
            
            If bytMode = 2 And dblTotal <> 0 Then
                strSql = _
                "zl_���ﻮ�ۼ�¼_Insert('" & strPriceNo_Out & "'," & k & "," & ZVal(Nvl(mrsInfo!����ID)) & ",NULL," & _
                         IIf(mstr����� = "", "NULL", mstr�����) & ",'" & str���ʽ���� & "'," & _
                         "'" & txtPatient.Text & "','" & mstr�Ա� & "','" & mstrAge & "'," & _
                         "'" & mstr�ѱ� & "',NULL," & lng�Һſ���ID & "," & _
                         IIf(lng�Һſ���ID <> 0, lng�Һſ���ID, UserInfo.����ID) & ",'" & UserInfo.���� & "'," & IIf(rsItems!���� = 2, 1, "NULL") & "," & _
                         rsItems!��ĿID & ",'" & rsItems!��� & "','" & rsItems!���㵥λ & "'," & _
                         "NULL,1," & rsItems!���� & ",NULL," & IIf(rsItems!ִ�п���ID = 0, lng�Һſ���ID, rsItems!ִ�п���ID) & "," & IIf(int�۸񸸺� = k, "NULL", int�۸񸸺�) & "," & _
                         rsIncomes!������ĿID & ",'" & rsIncomes!�վݷ�Ŀ & "'," & rsIncomes!���� & "," & _
                         rsIncomes!Ӧ�� & "," & rsIncomes!ʵ�� & ",to_Date('" & str����ʱ�� & "','yyyy-mm-dd hh24:mi:ss'),to_date('" & str�Ǽ�ʱ�� & "','yyyy-mm-dd hh24:mi:ss'),NULL,'" & UserInfo.���� & "','�Һ�:" & strNO_Out & "')"
                Call zlAddArray(cllPro_out, strSql)
            End If
            k = k + 1
            rsIncomes.MoveNext
            Next j
        rsItems.MoveNext
    Next i

    If Not mblnAppointment Then
        If strҽ������ = UserInfo.���� Then
            strSql = "ZL_���˹Һż�¼_��������('" & strNO_Out & "'," & Nvl(mrsInfo!����ID) & ",'" & lblRoomName.Caption & "','" & UserInfo.���� & "','','','" & zl_GetԤԼ��ʽByNo(strNO_Out) & "')"    '�����:48350
            Call zlAddArray(cllPro_out, strSql)
            strSql = "zl_���˽���(" & Nvl(mrsInfo!����ID) & ",'" & strNO_Out & "',NULL,'" & UserInfo.���� & "','" & lblRoomName.Caption & "')"
            Call zlAddArray(cllPro_out, strSql)
        End If
    End If
    
    GetSaveRegDataSQL = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    Select Case mRegistFeeMode
    Case EM_RG_����
        If SaveData_Price = False Then GoTo ToFail
    Case EM_RG_����
        If SaveData_Accounting = False Then GoTo ToFail
    Case EM_RG_����
        If SaveData_Cash = False Then GoTo ToFail
    End Select
    mblnUpdateAge = False
    Call ReloadPage
    mblnOK = True
    Unload Me
    Exit Sub
ToFail:
    If mViewMode = v_ר�Һ� Or mViewMode = v_ר�Һŷ�ʱ�� Then Call CancelRegNo
End Sub

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

Private Function CheckValied(ByVal bytMode As Byte) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ֵ �ĺϷ���
    '���:bytMode-0-����;1-����;2-����
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-02-02 16:18:48
    '---------------------------------------------------------------------------------------------------------------------------------------------


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
    
    If mstr�ѱ� = "" Then
        MsgBox "���˷ѱ���Ϊ��,����ѡ��һ���ѱ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If bytMode = 0 Then  'bytMode-0-����;1-����;2-����
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


Private Sub chkAll_Click()
    mty_Para.blnShowAllPlan = chkAll.Value <> 0
End Sub

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
            cmdPrice.Visible = True
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
        Me.Caption = "ҽ��վ" & IIf(gSysPara.bln��Һ�ģʽ, "ֱ�Ӿ���", "�Һ�")
        lblAppointStyle.Visible = False
        cboAppointStyle.Visible = False
    End If
    
    gobjDatabase.ExecuteProcedure "Zl_�ҺŰ���_Autoupdate", Me.Caption
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
    If mblnAppointment And mlng����ID <> 0 Then
        Call GetPatient(IDKind.GetCurCard, "-" & mlng����ID, False)
    End If
    cmdNewPati.Enabled = InStr(gstrPrivs, ";�ҺŲ��˽���;") > 0
    cmdOther.Enabled = InStr(gstrPrivs, ";���������ҽ���ĺ�Դ;") > 0
    '137272:���ϴ�,2019/2/20,��ֹ���ź�ϵͳ������������
    Call CancelRegNo
End Sub

Private Sub InitAppointmentTime()
    '��ʼ��ԤԼʱ��
    Dim intԤԼ���� As Integer
    Dim dtNow As Date
    
    On Error GoTo ErrHandler
    intԤԼ���� = mintSysAppLimit
    If mblnAppointment Then
        Call mobjRegister.zlGetRegisterMaxDaysFromDeptAndDoctor_Tradition( _
            gstrDeptIDs, UserInfo.����, mty_Para.blnԤԼ�������Ұ���, intԤԼ����)
    End If

    dtNow = gobjDatabase.Currentdate
    dtpDate.minDate = Format(dtNow, "yyyy-mm-dd")
    dtpDate.MaxDate = Format(dtNow + intԤԼ����, "yyyy-mm-dd")
    dtpDate.Value = Format(dtNow + 1, "yyyy-mm-dd")
    dtpTime.Value = Format(dtNow, "hh:mm:ss")
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

    If zlPatiCardCheck(1, lng����ID, CStr(Split(mstrYBPati, ";")(0)), 2) = False Then
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
    Set rsTmp = GetMoneyInfoRegist(lng����ID, , , 1, , , True)
    Dim dbl������� As Double
    cur��� = 0
    Do While Not rsTmp.EOF
        cur��� = cur��� + Val(Nvl(rsTmp!Ԥ�����))
        cur��� = cur��� - Val(Nvl(rsTmp!�������))
        If Val(Nvl(rsTmp!����)) = 1 Then
            dbl������� = Val(Nvl(rsTmp!Ԥ�����)) - Val(Nvl(rsTmp!�������))
        End If
        rsTmp.MoveNext
    Loop
    If cur��� > 0 Then
        lblMoney.Caption = "����Ԥ�����:" & Format(cur���, "0.00") & _
                        IIf(FormatEx(dbl�������, 6) <> 0, "(������:" & Format(dbl�������, "0.00") & ")", "")
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
        cmdPrice.Visible = False
        
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
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    
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
    '      blnInputIDCard-�Ƿ����֤ˢ��
    '����:Cancel-Ϊtrue��ʾ���صķ�����ȡ������Ϣ
    Dim strSql As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur��� As Currency, curMoney As Currency
    Dim strInputInfo As String '���洫��������ı� ������ʹ�����֤�� �Բ��˽��в��Һ� ���滻��"-" ����ID�����
    Dim i As Integer, strPati As String, dat��������
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
               "B.���� ��������,A.��ѯ���� As ����֤��,A.����ģʽ,a.��ҳID From ������Ϣ A,������� B Where A.���� = B.���(+) And A.ͣ��ʱ�� is NULL "

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
        mstr�Ա� = Nvl(mrsInfo!�Ա�)
        txtPatient.PasswordChar = ""
        
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        If Load�ѱ�(Nvl(mrsInfo!�ѱ�)) = False Then mstr�ѱ� = ""
        
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
        
        mstr����� = Nvl(mrsInfo!�����)
        If mstr����� = "" Then
            mstr����� = gobjDatabase.GetNextNo(3)
            mblnChangeFeeType = True
        Else
            mblnChangeFeeType = False
        End If
        
        lblInfo.Caption = "�Ա�:" & mstr�Ա� & "   ����:" & mstrAge & "   �����:" & mstr����� & "   �ѱ�:" & mstr�ѱ�
        
        '����Ԥ������Ϣ
        Set rsTmp = GetMoneyInfoRegist(mrsInfo!����ID, , , 1, , , True)
        Dim dbl������� As Double
        cur��� = 0
        Do While Not rsTmp.EOF
            cur��� = cur��� + Val(Nvl(rsTmp!Ԥ�����))
            cur��� = cur��� - Val(Nvl(rsTmp!�������))
            If Val(Nvl(rsTmp!����)) = 1 Then
                dbl������� = Val(Nvl(rsTmp!Ԥ�����)) - Val(Nvl(rsTmp!�������))
            End If
            rsTmp.MoveNext
        Loop
        If cur��� > 0 Then
            lblMoney.Caption = "����Ԥ�����:" & Format(cur���, "0.00") & _
                            IIf(FormatEx(dbl�������, 6) <> 0, "(������:" & Format(dbl�������, "0.00") & ")", "")
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
        
        Call ResetDefault���� 'ȱʡ��ȡ�����־
        
        '���ݲ������¶�ȡ��Ŀ����
        If mintPriceGradeStartType >= 2 Then
           Call GetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID)), Nvl(mrsInfo!ҽ�Ƹ��ʽ), , , mstrPriceGrade)
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
    mstr�Ա� = ""
    mstrAge = ""
    cmdNewPati.ToolTipText = "��������(F4)"
    cmdNewPati.Enabled = InStr(gstrPrivs, ";�ҺŲ��˽���;") > 0
    mstr����� = ""
    mstr�ѱ� = ""
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
    Dim i As Integer, k As Integer
    Dim DateThis As Date, strZero As String
    Dim str�ҺŰ��� As String, strViewSQL As String
    Dim str�ҺŰ��żƻ� As String, str�ҺŻ��ܼƻ� As String
    Dim str����         As String, str�ҺŻ��ܰ��� As String
    Dim vRect          As RECT
    Dim varTemp As Variant, varData As Variant
    On Error GoTo errH
    
    If chkAll.Value = 0 Then
        varTemp = Split(mty_Para.strStationRegOrder, "|")
        For i = 0 To UBound(varTemp)
            varData = Split(varTemp(i), ",")
            Select Case varData(0)
                Case "ҽ��"
                    str���� = str���� & ",Decode(ҽ��,Null,Decode(����ID," & mlngDept & ",3,4),Decode(����ID," & mlngDept & ",1,2)),ҽ�� " & IIf(varData(1) = 1, "", "desc")
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
        str���� = Mid(str����, 2)
    Else
        str���� = "Decode(ҽ��,'" & UserInfo.���� & "',1,2),Decode(����ID," & mlngDept & ",1,2),�ű�,��Ŀ,�ѹ�"
    End If
    
    If gstrDeptIDs <> "" And Not blnOtherDoctor Then strIF = " And Instr(','||[4]||',',','||P.����ID||',')>0"
    If mblnAppointment Then
        If mty_Para.blnԤԼ�������Ұ��� Then
            strIF = strIF & IIf(blnOtherDoctor, " And (p.ҽ������ <> [1] or p.ҽ������ Is Null )", " And (p.ҽ������ = [1] or p.ҽ������ Is Null )")
        Else
            strIF = strIF & IIf(blnOtherDoctor, " And (p.ҽ������ <> [1] )", " And (p.ҽ������ = [1])")
        End If
    Else
        If mty_Para.bln�ҺŰ������Ұ��� Then
            strIF = strIF & IIf(blnOtherDoctor, " And (p.ҽ������ <> [1] or p.ҽ������ Is Null)", " And (p.ҽ������ = [1] or p.ҽ������ Is Null)")
        Else
            strIF = strIF & IIf(blnOtherDoctor, " And (p.ҽ������ <> [1] )", " And (p.ҽ������ = [1])")
        End If
    End If
    
    If intSelMode = 2 Then
        strIF = strIF & " And (p.���� Like [8] Or Upper(b.����) Like Upper([8]) Or Upper(zlSpellCode(b.����)) Like Upper([8]) Or Upper(p.ҽ������) Like Upper([8]) Or Upper(zlSpellCode(p.ҽ������)) Like Upper([8]))"
    End If
     
    str�ҺŰ��� = "" & _
        "            Select A.ID, A.����, A.����, A.����id, A.��Ŀid, A.ҽ��id, A.ҽ������, A.��������, A. ����, A.��һ, A.�ܶ�, A.����, " & _
        "                   A.���� , A.����, A.����, A.���﷽ʽ,A.��ſ���, B.�޺���, B.��Լ��,a.ͣ������ " & IIf(chkAll.Value <> 1, ",c.��ʼʱ�� ", "") & vbNewLine & _
        "            From �ҺŰ��� A, �ҺŰ������� B " & IIf(chkAll.Value <> 1, ", ʱ��� C ", "") & vbNewLine & _
        "            Where a.ͣ������ Is Null And [5] Between Nvl(a.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
        "                 Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) " & _
        "                  And a.ID = B.����id(+) " & IIf(mblnAppointment, " And Trunc(Sysdate)+Nvl(A.ԤԼ����," & IIf(mintSysAppLimit = 0, 1, mintSysAppLimit) & ") >= [5]", "") & _
        "                  And Decode(To_Char([5], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) = B.������Ŀ(+)" & vbNewLine & _
        IIf(chkAll.Value <> 1, " And Decode(To_Char([5], 'D'), '1', a.����, '2', a.��һ, '3', a.�ܶ�, '4', a.����, '5', a.����, '6', a.����, '7',a.����, Null)  = c.ʱ��� And c.վ�� Is Null And c.���� Is Null ", "")
    
    '�ҺŰ��� �޺�����Լ�� �ҺŰ��������л�ȡ
    str�ҺŻ��ܰ��� = str�ҺŰ��� & " And Not Exists (Select 1 From �ҺŰ��żƻ� Where ����id = a.Id) "
    '�ҺŰ��żƻ� �޺�����Լ�� �Һżƻ������л�ȡ
    str�ҺŻ��ܼƻ� = " Union All " & _
        "            Select C.ID, A.����, C.����, C.����id, A.��Ŀid, A.ҽ��id, A.ҽ������, C.��������, A. ����, A.��һ, A.�ܶ�, A.����, " & _
        "                   A.���� , A.����, A.����, A.���﷽ʽ,A.��ſ���, B.�޺���, B.��Լ��,C.ͣ������ " & IIf(chkAll.Value <> 1, ",NULL as ��ʼʱ�� ", "") & vbNewLine & _
        "            From �ҺŰ��żƻ� A, �Һżƻ����� B,�ҺŰ��� C " & vbNewLine & _
        "            Where c.ͣ������ Is Null And [5] Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
        "                 a.ʧЧʱ�� And a.���ʱ�� Is Not Null And " & _
        "           a.��Чʱ�� = (Select Max(��Чʱ��)" & vbNewLine & _
        "                           From �ҺŰ��żƻ�" & vbNewLine & _
        "                           Where ����id = a.����id And [5] Between" & vbNewLine & _
        "                           Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And ʧЧʱ�� And" & vbNewLine & _
        "                           ���ʱ�� Is Not Null)" & _
        "                  And a.ID = B.�ƻ�id(+) And a.����id = c.Id " & IIf(mblnAppointment, "   And Trunc(Sysdate)+Nvl(C.ԤԼ����," & IIf(mintSysAppLimit = 0, 1, mintSysAppLimit) & ") >= [5]", "") & _
        "                  And Decode(To_Char([5], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) = B.������Ŀ(+)" & vbNewLine & _
        IIf(chkAll.Value <> 1, " And Decode(To_Char([5], 'D'), '1', a.����, '2', a.��һ, '3', a.�ܶ�, '4', a.����, '5', a.����, '6', a.����, '7',a.����, Null) Is Not Null", "")
    
    
    If mblnAppointment Then
        DateThis = Format(dtpDate, "yyyy-mm-dd hh:mm:ss")
    Else
        DateThis = gobjDatabase.Currentdate
    End If
    'ȡ��Ӧ���ڰ��ŵ�ʱ���
    strSql = "Decode(To_Char([5],'D'),'1',P.����,'2',P.��һ,'3',P.�ܶ�,'4',P.����,'5',P.����,'6',P.����,'7',P.����,NULL)"
    
    '�ò������ȡ��������Ӧ��ʱ���
    strTime = _
        "Select ʱ��� From ʱ��� Where ���� Is Null And վ�� Is Null And " & _
        "    ('3000-01-10 '||To_Char([5],'HH24:MI:SS') >" & _
        "               Decode(Sign(��ʼʱ��-��ֹʱ��),1,'3000-01-09 '||To_Char(Nvl(��ǰʱ��,��ʼʱ��),'HH24:MI:SS'),'3000-01-10 '||To_Char(Nvl(��ǰʱ��,��ʼʱ��),'HH24:MI:SS')))" & _
        " Or" & _
        " ('3000-01-10 '||To_Char([5],'HH24:MI:SS')  >" & _
        "   '3000-01-10 '||To_Char(Nvl(��ǰʱ��,��ʼʱ��),'HH24:MI:SS')) "
    
    If Not (mblnAppointment And Format(DateThis, "yyyy-mm-dd") > Format(gobjDatabase.Currentdate, "yyyy-mm-dd")) Then
        strWhere = IIf(chkAll.Value = 0, " And " & strSql & " IN(" & strTime & ")", "")
    End If

    '�ò�����䵱ʱ��ȡ���ְ��ŵĹҺ����
    strState = _
    "   Select A.ID as ����ID,B.�ѹ���,B.��Լ��" & _
    "   From (" & str�ҺŻ��ܰ��� & str�ҺŻ��ܼƻ� & ") A,���˹ҺŻ��� B" & _
    "   Where A.����ID = B.����ID And A.��ĿID = B.��ĿID" & _
    "               And Nvl(A.ҽ��ID,0)=Nvl(B.ҽ��ID,0) " & _
    "               And Nvl(A.ҽ������,'ҽ��')=Nvl(B.ҽ������,'ҽ��') " & _
    "               And (A.����=B.���� or B.���� is Null )  And B.����=[6]"
    
    If mblnAppointment Then
        str�ҺŰ��żƻ� = " " & _
            "             Select A.ID,A.ID as �ƻ�ID, A.����id, A.����, A.��Ŀid, A.������, A.����ʱ��, A. ����, A.��һ, A.�ܶ�, A.����, A.����, A.����," & _
            "                    A.���� , A.���﷽ʽ, A.��ſ���, B.�޺���, B.��Լ��, A.��Чʱ��, A.ʧЧʱ�� ,A.ҽ������,A.ҽ��ID" & IIf(chkAll.Value <> 1, ",D.��ʼʱ�� ", "") & _
            "             From �ҺŰ��żƻ� A, �Һżƻ����� B," & vbNewLine & _
            "                  (" & vbNewLine & _
            "                      Select Max(��Чʱ��) As ��Чʱ��, ����id" & _
            "                      From �ҺŰ��żƻ� " & vbNewLine & _
            "                      Where ���ʱ�� Is Not Null And  [5] Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
            "                          ʧЧʱ��  " & vbNewLine & _
            "                       Group By ����id" & vbNewLine & _
            "                   ) C" & IIf(chkAll.Value <> 1, ",ʱ��� D", "") & _
            "             Where A.���ʱ�� Is Not Null And ([5] Between  A.��Чʱ�� + 0 And A.ʧЧʱ��)" & _
            "                   And A.ID = B.�ƻ�id(+) And " & vbNewLine & _
            "                   Decode(To_Char([5], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6'," & _
            "                  '����', '7', '����', Null) = B.������Ŀ(+) And A.��Чʱ�� = C.��Чʱ�� And A.����id = C.����id " & _
             IIf(chkAll.Value <> 1, "And Decode(To_Char([5], 'D'), '1', a.����, '2', a.��һ, '3', a.�ܶ�, '4', a.����, '5', a.����, '6', a.����, '7',a.����, Null) = d.ʱ��� And d.վ�� Is Null And d.���� Is Null ", "")
    
        strSql = _
        " Select P.ID,0 as �ƻ�ID,P.���� ,P.����,P.����ID,P.��ĿID," & _
        "       P.ҽ��ID,P.ҽ������,P.�޺���,P.��Լ��,Nvl(P.��������,0) as ��������," & _
        "       P.����,P.��һ ,P.�ܶ� ,P.���� ,P.���� ,P.���� ,P.����,P.���﷽ʽ,P.��ſ���," & _
        "       Decode(To_Char([5],'D'),'1',P.����,'2',P.��һ,'3',P.�ܶ�,'4',P.����,'5',P.����,'6',P.����,'7',P.����,NULL)  as �Ű�" & IIf(chkAll.Value <> 1, ",p.��ʼʱ�� ", "") & _
        " From (" & str�ҺŰ��� & ") P" & _
        " Where    Not Exists(Select 1 From �ҺŰ��żƻ� where ����ID=P.id And ([5] BETWEEN ��Чʱ�� + 0 and ʧЧʱ��)  And ���ʱ�� is not NULL  ) " & _
        "          And Not Exists(Select 1 From �ҺŰ���ͣ��״̬ Where ����ID=P.ID and [5] between ��ʼֹͣʱ�� and ����ֹͣʱ�� )" & _
        " Union ALL " & _
        " Select   C.ID,P.�ƻ�ID,C.����,C.����,C.����ID,P.��ĿID," & _
        "       P.ҽ��ID,P.ҽ������,P.�޺���,P.��Լ��,Nvl(C.��������,0) as ��������," & _
        "       P.����,P.��һ ,P.�ܶ� ,P.���� ,P.���� ,P.���� ,P.����,P.���﷽ʽ,P.��ſ���," & _
        "       Decode(To_Char([5],'D'),'1',P.����,'2',P.��һ,'3',P.�ܶ�,'4',P.����,'5',P.����,'6',P.����,'7',P.����,NULL)  as �Ű�" & IIf(chkAll.Value <> 1, ",p.��ʼʱ�� ", "") & _
        " From (" & str�ҺŰ��żƻ� & ") P, �ҺŰ��� C" & _
        " Where P.����ID=C.ID  And C.ͣ������ Is  NULL  And Trunc(Sysdate)+Nvl(C.ԤԼ����," & IIf(mintSysAppLimit = 0, 1, mintSysAppLimit) & ") >= [5]  " & _
        "           And Not Exists(Select 1 From �ҺŰ���ͣ��״̬ Where ����ID=C.ID and [5] between ��ʼֹͣʱ�� and ����ֹͣʱ�� )"
        strSql = "(" & strSql & ") P"
    Else
        strSql = _
                    " (Select P.ID,0 as �ƻ�ID,P.���� ,P.����,P.����ID,P.��ĿID," & _
                    "       P.ҽ��ID,P.ҽ������,P.�޺���,P.��Լ��,Nvl(P.��������,0) as ��������," & _
                    "       P.����,P.��һ ,P.�ܶ� ,P.���� ,P.���� ,P.���� ,P.����,P.���﷽ʽ,P.��ſ���," & _
                    "       Decode(To_Char([5],'D'),'1',P.����,'2',P.��һ,'3',P.�ܶ�,'4',P.����,'5',P.����,'6',P.����,'7',P.����,NULL) as �Ű�" & IIf(chkAll.Value <> 1, ",p.��ʼʱ�� ", "") & _
                    " From (" & str�ҺŰ��� & ") P "
        strSql = strSql & vbNewLine & "  ) P"
    End If
    
    strViewSQL = _
                "Select Distinct " & _
                "       P.ID,P.���� as �ű�,P.����,P.����ID,B.���� As ����,C.���� As ��Ŀ," & _
                "       P.ҽ������ as ҽ��,Nvl(A.�ѹ���,0) as �ѹ�,Nvl(A.��Լ��,0) as ��Լ," & _
                "       P.�޺��� as �޺�,P.��Լ�� as ��Լ,Decode(Nvl(P.��������,0),1,'��','') as ����,Decode(Nvl(C.��Ŀ����,0),1,'��','') as ����," & _
                "       Decode(P.���﷽ʽ,1,'ָ��',2,'��̬',3,'ƽ��',NULL) as ����,Decode(Nvl(P.��ſ���,0),1,'��','') As ��ſ���,P.�Ű�" & IIf(chkAll.Value <> 1, ",p.��ʼʱ��", "") & _
                " From " & strSql & "," & vbCrLf & _
                "           (" & strState & ") A,���ű� B,�շ���ĿĿ¼ C" & _
                " Where P.ID=A.����ID(+) And Nvl(B.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And P.����ID=B.ID And P.��ĿID=C.ID" & strIF & strZero & _
                "           And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                "           And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & strWhere & _
                "           And (Nvl(P.ҽ��ID,0)=0 Or Exists(Select 1 From ��Ա�� Q Where P.ҽ��ID=Q.ID And (Q.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or Q.����ʱ�� Is Null)))" & _
                " Order by " & str����
    If chkAll.Value <> 1 Then
        strViewSQL = _
                    "Select  ID,�ű�,����,����ID,����,��Ŀ,ҽ��,�ѹ�, ��Լ,�޺�, ��Լ,����, ����," & _
                    "           ����, ��ſ���,�Ű� " & _
                    "From (" & strViewSQL & ")"
    End If
    
    strSql = _
                "Select Distinct " & _
                "       P.ID,p.�ƻ�ID,P.���� as �ű�,P.����,P.����ID,B.���� As ����,P.��ĿID,C.���� As ��Ŀ," & _
                "       P.ҽ��ID,P.ҽ������ as ҽ��,Nvl(A.�ѹ���,0) as �ѹ�,Nvl(A.��Լ��,0) as ��Լ," & _
                "       P.�޺��� as �޺�,P.��Լ�� as ��Լ,Nvl(P.��������,0) as ����,Nvl(C.��Ŀ����,0) as ����," & _
                "       P.���� as ��,P.��һ as һ,P.�ܶ� as ��,P.���� as ��,P.���� as ��,P.���� as ��,P.���� as ��," & _
                "       Decode(P.���﷽ʽ,1,'ָ��',2,'��̬',3,'ƽ��',NULL) as ����,P.��ſ���,P.�Ű�" & IIf(chkAll.Value <> 1, ",p.��ʼʱ��", "") & _
                " From " & strSql & "," & vbCrLf & _
                "           (" & strState & ") A,���ű� B,�շ���ĿĿ¼ C" & _
                " Where P.ID=A.����ID(+) And Nvl(B.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And P.����ID=B.ID And P.��ĿID=C.ID" & strIF & strZero & _
                "           And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                "           And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & strWhere & _
                "           And (Nvl(P.ҽ��ID,0)=0 Or Exists(Select 1 From ��Ա�� Q Where P.ҽ��ID=Q.ID And (Q.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or Q.����ʱ�� Is Null)))" & _
                " Order by " & str����
                
    Set mrsPlan = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, _
            UserInfo.����, "%", "", gstrDeptIDs, DateThis, CDate(Format(DateThis, "yyyy-MM-dd")), CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60, "%" & strFilter & "%")
 
    
    If mrsPlan.RecordCount <> 0 Then
        If intSelMode = 1 Or mrsPlan.RecordCount = 1 Then
            'Ĭ�϶�ȡ
            Call ReadLimit(Nvl(mrsPlan!�ű�))
        Else
            vRect = GetControlRect(txtReg.hWnd)
            Set rsPlan = gobjDatabase.ShowSQLSelect(Me, strViewSQL, 0, "����ѡ��", False, "", "����ѡ��", _
                                                False, False, True, vRect.Left, vRect.Top - 250, 600, False, True, False, _
                                                UserInfo.����, "%", "", gstrDeptIDs, DateThis, CDate(Format(DateThis, "yyyy-MM-dd")), CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60, "%" & strFilter & "%")
            If rsPlan Is Nothing Then
                Call ReadLimit(Nvl(mrsPlan!�ű�))
            Else
                If Not rsPlan.EOF Then
                    Call ReadLimit(Nvl(rsPlan!�ű�))
                Else
                    Call ReadLimit(Nvl(mrsPlan!�ű�))
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
                    cmdTime.Visible = True
                Case Else
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

Private Sub ReadLimit(ByVal strRegNo As String)
    If mrsPlan Is Nothing Then Exit Sub
    If mrsPlan.State = 0 Then Exit Sub
    
    mrsPlan.Filter = "�ű�='" & strRegNo & "'"
    
    If mrsPlan.RecordCount = 0 Then Exit Sub
    mblnIntact = True
    mblnChangeByCode = True
    If Nvl(mrsPlan!ҽ��) = "" Then
        txtReg.Text = "[" & Nvl(mrsPlan!�ű�) & "]" & Nvl(mrsPlan!��Ŀ)
    Else
        txtReg.Text = "[" & Nvl(mrsPlan!�ű�) & "]" & Nvl(mrsPlan!��Ŀ) & "(" & Nvl(mrsPlan!ҽ��) & ")"
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

Private Function GetActiveView()
    '�õ���ǰ�Һ�ҵ��  ��ȡ�������͵�����
    Dim strSql          As String
    Dim rsTmp           As ADODB.Recordset
    Dim str����         As String
    Dim dat            As Date
    
    On Error GoTo errH
    str���� = txtReg.Tag
    If mblnAppointment Then
        dat = dtpDate.Value
    Else
        dat = gobjDatabase.Currentdate
    End If
    
    strSql = _
    "       Select   Havedata, ����id" & vbNewLine & _
    "       From (" & vbNewLine & _
    "               Select 1 As Havedata, b.Id As ����id " & vbNewLine & _
    "               From �ҺŰ���ʱ�� A, �ҺŰ��� B" & vbNewLine & _
    "               Where B.����=[1] And A.����id = b.ID " & _
    "                And   Decode(To_Char([2], 'D'), '1', '����', '2'," & _
    "                   '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6','����', '7', '����', Null) =a.���� " & vbNewLine & _
    "                       And Not Exists" & vbNewLine & _
    "                     (Select 1 From �ҺŰ��żƻ� C " & vbNewLine & _
    "                         Where c.����id = b.Id And c.���ʱ�� Is Not Null And [2] Between " & _
    "                               Nvl(c.��Чʱ��, [2]) And" & _
    "                          c.ʧЧʱ��)" & vbNewLine & _
    "               Union All " & vbNewLine & _
    "               Select 1 As Havedata, c.Id As ����id" & vbNewLine & _
    "               From �Һżƻ�ʱ�� A, �ҺŰ��żƻ� B, �ҺŰ��� C,(" & vbNewLine & _
    "                   SELECT MAX(a.��Чʱ�� ) ��Ч FROM �ҺŰ��żƻ� a,�ҺŰ��� B  WHERE a.����Id=b.ID AND b.����=[1] AND a.���ʱ�� IS NOT NULL" & vbNewLine & _
    "             And [2] Between nvl(a.��Чʱ��,to_date('1900-01-01','yyyy-mm-dd')) And a.ʧЧʱ��" & vbNewLine & _
    "           ) D  " & vbNewLine & _
    "               Where  C.����=[1] And c.Id = b.����id And b.Id = a.�ƻ�id And b.��Чʱ��=d.��Ч And b.���ʱ�� Is Not Null" & _
    "                    And   Decode(To_Char([2], 'D'), '1', '����', '2'," & _
    "                   '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6','����', '7', '����', Null) =a.���� " & vbNewLine & _
    "                       And [2] Between Nvl(b.��Чʱ��,[2]) And b.ʧЧʱ��) B"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, str����, dat)
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
     Dim dateCur        As Date
     Dim strNO          As String
     Dim vRect          As RECT
    If Not mblnAppointment Then Exit Function
    
    strSql = "" & _
    " Select Distinct a.��� As ID, A.���,To_Char(a.��ʼʱ��, 'hh24:mi') As ��ʼʱ��, To_Char(a.����ʱ��, 'hh24:mi') As ����ʱ��" & vbNewLine & _
    " From �ҺŰ���ʱ�� A, �ҺŰ��� B" & vbNewLine & _
    " Where a.����id = b.Id And b.���� = [1] And" & vbNewLine & _
    " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.��ʼʱ��, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
    "      Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',Null) = a.����(+)  " & _
    "      And Not Exists (Select Count(1) From �Һ����״̬ Where Trunc(����) = [2] And ���� = b.���� And (��� = a.��� Or ��� Like a.��� || '__') Having Count(1) - a.�������� >= 0) " & _
    "      And Not Exists (Select 1 From �ҺŰ��żƻ� E Where e.����id = b.Id And e.���ʱ�� Is Not Null And [2] Between Nvl(e.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And e.ʧЧʱ��)"
    
    strSql = strSql & " Union " & _
    "Select Distinct a.��� As ID,A.���,To_Char(a.��ʼʱ��, 'hh24:mi') As ��ʼʱ��, To_Char(a.����ʱ��, 'hh24:mi') As ����ʱ��" & vbNewLine & _
    "From �Һżƻ�ʱ�� A, �ҺŰ��żƻ� B, �ҺŰ��� C," & vbNewLine & _
    "     (Select Max(a.��Чʱ��) ��Ч" & vbNewLine & _
    "       From �ҺŰ��żƻ� A, �ҺŰ��� B" & vbNewLine & _
    "       Where a.����id = b.Id And b.���� = [1] And a.���ʱ�� Is Not Null And" & vbNewLine & _
    "             [2] Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
    "             a.ʧЧʱ��) D" & vbNewLine & _
    "Where a.�ƻ�id = b.Id And b.����id = c.Id And c.���� = [1] And b.��Чʱ�� = d.��Ч And b.���ʱ�� Is Not Null And" & vbNewLine & _
    " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.��ʼʱ��, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
    "      [2] Between Nvl(b.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
    "      b.ʧЧʱ�� And Not Exists" & vbNewLine & _
    " (Select Count(1)" & vbNewLine & _
    "       From �Һ����״̬" & vbNewLine & _
    "       Where Trunc(����) = [2] And ���� = b.���� And (��� = a.��� Or ��� Like a.��� || '__') Having" & vbNewLine & _
    "        Count(1) - a.�������� >= 0) And Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5'," & vbNewLine & _
    "                                           '����', '6', '����', '7', '����', Null) = a.����(+)" & vbNewLine & _
            "Order By ��ʼʱ��"


    dateCur = Format(dtpDate, "yyyy-mm-dd")
    If strSql = "" Then Exit Function
    strNO = txtReg.Tag
    vRect = GetControlRect(dtpTime.hWnd)
    
    On Error GoTo errH
    
    Set mrsʱ��� = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "ԤԼʱ��ѡ��", False, "", "ԤԼʱ��ѡ��", _
                                                False, False, True, vRect.Left, vRect.Top - 300, 600, False, True, False, strNO, dateCur)
    
    If mrsʱ��� Is Nothing Then Exit Function
    If mrsʱ���.EOF Then Exit Function
    
    lblSn.Caption = ""
    dtpTime.Value = Format(mrsʱ���!��ʼʱ��, "hh:mm:ss")
    lblSn.Caption = "���:" & Val(Nvl(mrsʱ���!���))
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
    ReadRegistPrice lngItemID, blnBook, False, mstr�ѱ�, rsItems, rsIncomes, , , , 1, _
        Val(Nvl(mrsPlan!����ID)), strPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
    If mintInsure <> 0 Then
        If MCPAR.�Һż����Ŀ = True Then
            If gclsInsure.CheckItem(mintInsure, 2, 0, rsItems) = False Then
                MsgBox "ҽ�������շ���Ŀ���ʧ�ܣ����ܼ��� " & IIf(gSysPara.bln��Һ�ģʽ, "����", "�Һ�") & "��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            If Not mrsInfo.EOF Then lng����ID = Val(Nvl(mrsInfo!����ID))
        End If
    End If
    ReadRegistPrice lngItemID, blnBook, False, mstr�ѱ�, rsItems, rsIncomes, lng����ID, mintInsure, _
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


Private Function GetSNState(str�ű� As String, datThis As Date, Optional lngSN As Long) As ADODB.Recordset
    Dim strSql           As String
    Dim datStart         As Date
    Dim datEnd           As Date
    On Error GoTo errH
    datStart = CDate(Format(datThis, "yyyy-MM-dd"))
    datEnd = DateAdd("s", -1, DateAdd("d", 1, datStart))
    strSql = "    " & vbNewLine & " Select ���,״̬,����Ա����,Nvl(ԤԼ,0) as ԤԼ,TO_Char(����,'hh24:mi:ss') as ����  "
    strSql = strSql & vbNewLine & " From �Һ����״̬ "
    strSql = strSql & vbNewLine & " Where ����=[1]"
    strSql = strSql & vbNewLine & IIf(datThis = CDate(0), " And ���� Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 ", " And ���� Between  [2] And [3]")
    strSql = strSql & vbNewLine & IIf(lngSN > 0, " And ���=[4]", "")
    Set GetSNState = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, str�ű�, datStart, datEnd, lngSN)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
            mstr�ѱ� = str�ѱ�
        Else
            mstr�ѱ� = mstrDef�ѱ�
        End If
    Else
        mstr�ѱ� = mstrDef�ѱ�
    End If
    If mstr�ѱ� = "" Then
        MsgBox "δ�ҵ������ڡ�" & Nvl(mrsPlan!����) & "����ȱʡ�ѱ�,���ڡ�������ϸ��Ϣ�����������ò��˷ѱ�", vbInformation, gstrSysName
        Load�ѱ� = False
        Exit Function
    End If
    lblInfo.Caption = "�Ա�:" & mstr�Ա� & "   ����:" & mstrAge & "   �����:" & mstr����� & "   �ѱ�:" & mstr�ѱ�
    Load�ѱ� = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
