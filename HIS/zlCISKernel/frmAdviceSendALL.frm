VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdviceSendALL 
   AutoRedraw      =   -1  'True
   Caption         =   "סԺҽ������"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   270
   ClientWidth     =   11760
   Icon            =   "frmAdviceSendALL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmAdviceSendALL.frx":058A
   ScaleHeight     =   7905
   ScaleWidth      =   11760
   Begin MSComCtl2.DTPicker dkpExecTime 
      Height          =   300
      Left            =   5400
      TabIndex        =   54
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   216727555
      CurrentDate     =   40976
   End
   Begin VB.PictureBox picҩ�� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   5400
      ScaleHeight     =   1800
      ScaleWidth      =   3885
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3870
      Width           =   3885
      Begin VB.CheckBox chkDeptExchange 
         Caption         =   "���η�������ҩ���û�"
         Height          =   255
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Width           =   2535
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDept 
         Height          =   1545
         Left            =   0
         TabIndex        =   25
         Top             =   255
         Width           =   3840
         _cx             =   6773
         _cy             =   2725
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
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   280
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAdviceSendALL.frx":0B14
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
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
   Begin VB.PictureBox pic���� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   90
      ScaleHeight     =   6375
      ScaleWidth      =   3840
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   120
      Width           =   3840
      Begin VB.Frame fraҽ������ 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1590
         Left            =   75
         TabIndex        =   48
         Top             =   1155
         Width           =   3460
         Begin VB.ComboBox cboDrugType 
            Height          =   300
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1260
            Width           =   2715
         End
         Begin VB.CommandButton cmdִ�п��� 
            Height          =   240
            Left            =   3200
            Picture         =   "frmAdviceSendALL.frx":0B69
            Style           =   1  'Graphical
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��ִ�п���(F4)"
            Top             =   600
            Width           =   240
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            ItemData        =   "frmAdviceSendALL.frx":0C5F
            Left            =   750
            List            =   "frmAdviceSendALL.frx":0C61
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   0
            Width           =   900
         End
         Begin VB.ComboBox cboҩ�� 
            Height          =   300
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   915
            Width           =   2715
         End
         Begin VB.TextBox txtִ�п��� 
            Height          =   300
            Left            =   750
            TabIndex        =   10
            Top             =   585
            Width           =   2715
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   1645
            TabIndex        =   9
            Top             =   0
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   216727555
            CurrentDate     =   37953
         End
         Begin VB.Label lblDrugType 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҩ������"
            Height          =   180
            Left            =   0
            TabIndex        =   57
            Top             =   1305
            Width           =   720
         End
         Begin VB.Label lblEatDrug 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ڷ�ҩ����ʱ�䣺2015-11-25 14:59"
            Height          =   180
            Left            =   0
            TabIndex        =   56
            Top             =   345
            Width           =   3240
         End
         Begin VB.Label lblִ�п��� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ִ�п���"
            Height          =   180
            Left            =   0
            TabIndex        =   51
            Top             =   645
            Width           =   720
         End
         Begin VB.Label lbl����ʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            Height          =   180
            Left            =   0
            TabIndex        =   50
            Top             =   60
            Width           =   720
         End
         Begin VB.Label lblҩ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩҩ��"
            Height          =   180
            Left            =   0
            TabIndex        =   49
            Top             =   960
            Width           =   720
         End
      End
      Begin VB.Frame fraҩƷ 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   90
         TabIndex        =   47
         Top             =   0
         Width           =   3500
         Begin VB.CheckBox chkSendNow 
            Caption         =   "ҩƷ�����ӵ��쿪ʼ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   750
            TabIndex        =   55
            Top             =   300
            Width           =   2655
         End
         Begin VB.CheckBox chkҩƷ��Ч 
            Caption         =   "ҩƷ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   280
            Index           =   1
            Left            =   2220
            TabIndex        =   1
            Top             =   0
            Width           =   1200
         End
         Begin VB.CheckBox chkҩƷ��Ч 
            Caption         =   "ҩƷ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   280
            Index           =   0
            Left            =   750
            TabIndex        =   0
            Top             =   0
            Width           =   1300
         End
         Begin VB.Image imgCheck 
            Height          =   360
            Left            =   140
            Picture         =   "frmAdviceSendALL.frx":0C63
            Top             =   -30
            Width           =   360
         End
      End
      Begin VB.Frame fra���� 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   90
         TabIndex        =   46
         Top             =   600
         Width           =   3500
         Begin VB.CheckBox chk������Ч 
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   280
            Index           =   1
            Left            =   2220
            TabIndex        =   3
            Top             =   0
            Width           =   1200
         End
         Begin VB.CheckBox chk������Ч 
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   280
            Index           =   0
            Left            =   750
            TabIndex        =   2
            Top             =   0
            Width           =   1300
         End
      End
      Begin VB.Frame fra״̬ 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   75
         TabIndex        =   44
         Top             =   960
         Width           =   3490
         Begin VB.OptionButton opt״̬ 
            Caption         =   "��У��"
            Height          =   180
            Index           =   1
            Left            =   1660
            TabIndex        =   5
            Top             =   0
            Width           =   900
         End
         Begin VB.OptionButton opt״̬ 
            BackColor       =   &H00D0FFFF&
            Caption         =   "�¿�"
            Height          =   180
            Index           =   0
            Left            =   750
            TabIndex        =   4
            Top             =   0
            Width           =   660
         End
         Begin VB.OptionButton opt״̬ 
            Caption         =   "ȫ��"
            Height          =   180
            Index           =   2
            Left            =   2780
            TabIndex        =   6
            Top             =   0
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.Label lbl״̬ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "״̬"
            Height          =   180
            Left            =   350
            TabIndex        =   45
            Top             =   0
            Width           =   360
         End
      End
      Begin VB.Frame fraӤ�� 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   5835
         Visible         =   0   'False
         Width           =   3210
         Begin VB.OptionButton optӤ�� 
            Caption         =   "Ӥ��ҽ��"
            Height          =   180
            Index           =   2
            Left            =   2175
            TabIndex        =   19
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton optӤ�� 
            Caption         =   "����ҽ��"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optӤ�� 
            Caption         =   "����ҽ��"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   18
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.CheckBox chk�Ӱ�Ӽ� 
         Caption         =   "ҽ�������������ķ���ִ�мӰ�Ӽ�"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   6105
         Width           =   3180
      End
      Begin VB.Frame fraPati 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   3360
         Left            =   60
         TabIndex        =   42
         Top             =   2415
         Width           =   3495
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   765
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   0
            Width           =   2715
         End
         Begin VB.CommandButton cmdNoPati 
            Caption         =   "ȫ��"
            Height          =   370
            Left            =   2820
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + R"
            Top             =   2940
            Width           =   675
         End
         Begin VB.CommandButton cmdAllPati 
            Caption         =   "ȫѡ"
            Height          =   370
            Left            =   2115
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + A"
            Top             =   2940
            Width           =   675
         End
         Begin VB.CommandButton cmdQuick 
            Caption         =   "�ſ�Ƿ�Ѳ���"
            Height          =   370
            Left            =   30
            TabIndex        =   14
            Top             =   2940
            Width           =   1380
         End
         Begin MSComctlLib.ListView lvwPati 
            Height          =   2490
            Left            =   15
            TabIndex        =   13
            Top             =   375
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   4392
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "����"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "סԺ��"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "����"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "ʣ���"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "סԺҽʦ"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "�ѱ�"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "����ȼ�"
               Object.Width           =   2028
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "����"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "��Ժ����"
               Object.Width           =   2857
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "��������"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "���ۺ�"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Left            =   15
            TabIndex        =   43
            Top             =   60
            Width           =   720
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5000
         Y1              =   5790
         Y2              =   5790
      End
   End
   Begin VB.PictureBox pic��� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   105
      ScaleHeight     =   915
      ScaleWidth      =   3630
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6510
      Width           =   3630
      Begin VB.ListBox lstClass 
         Columns         =   4
         Height          =   900
         ItemData        =   "frmAdviceSendALL.frx":1CE5
         Left            =   0
         List            =   "frmAdviceSendALL.frx":1CE7
         Style           =   1  'Checkbox
         TabIndex        =   26
         Top             =   0
         Width           =   3450
      End
   End
   Begin VB.PictureBox pic��ҩ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   5430
      ScaleHeight     =   2280
      ScaleWidth      =   3630
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1485
      Width           =   3630
      Begin VB.CommandButton cmdUnWay 
         Caption         =   "��ѡ"
         Height          =   330
         Left            =   1440
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   1920
         Width           =   675
      End
      Begin VB.CommandButton cmdNoWay 
         Caption         =   "ȫ��"
         Height          =   330
         Left            =   2805
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   1920
         Width           =   675
      End
      Begin VB.CommandButton cmdAllWay 
         Caption         =   "ȫѡ"
         Height          =   330
         Left            =   2130
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   1920
         Width           =   675
      End
      Begin MSComctlLib.ListView lvwWay 
         Height          =   1905
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   3360
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��ҩ;��"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6480
      Left            =   4065
      MousePointer    =   9  'Size W E
      TabIndex        =   35
      Top             =   870
      Width           =   45
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   7350
      TabIndex        =   34
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   7665
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   2175
      TabIndex        =   30
      Top             =   7620
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   7545
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceSendALL.frx":1CE9
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17489
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   2
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
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   4155
      MousePointer    =   7  'Size N S
      TabIndex        =   33
      Top             =   5910
      Width           =   7530
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Height          =   1425
      Left            =   4155
      TabIndex        =   28
      Top             =   6030
      Width           =   7545
      _cx             =   13309
      _cy             =   2514
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
      BackColorSel    =   4210752
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceSendALL.frx":257D
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
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4215
      TabIndex        =   31
      Top             =   525
      Width           =   7425
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0FFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   60
         Width           =   90
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4980
      Left            =   4155
      TabIndex        =   27
      Top             =   825
      Width           =   7530
      _cx             =   13282
      _cy             =   8784
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
      BackColorSel    =   16444122
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
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
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceSendALL.frx":2618
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
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin XtremeSuiteControls.TaskPanel tkpMain 
      Height          =   7275
      Left            =   45
      TabIndex        =   36
      Top             =   180
      Width           =   3900
      _Version        =   589884
      _ExtentX        =   6879
      _ExtentY        =   12832
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   480
      Top             =   -30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAdviceSendALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mMainPrivs As String 'IN
Private mlng����ID As Long 'IN:���ڼ�¼������Ĳ������ϴη��Ͳ�����ѡ����ת������ʱ����ԭ����ID
Private mlng����ID As Long 'IN
Private mlng��ҳID As Long 'IN,�����˵���ʱ����
Private mblnSend As Boolean 'OUT:�Ƿ�ɹ����͹���
Private mblnRefresh As Boolean 'OUT�����ͺ��Ƿ�Ҫ��ˢ��
Private mblnOnePati As Boolean     '�����˻��Ƕಡ��ģʽ

'----------------------------------------------
Private mcolStock1 As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mcolStock2 As Collection '��Ÿ������Ŀⷿ�ĳ����鷽ʽ
Private mrsҩ�� As ADODB.Recordset

Private mrsBill As ADODB.Recordset
Private mrsWarn As ADODB.Recordset
Private mrsPrice As ADODB.Recordset '�����Ƽ۹�ϵ

'������ر���������������ȡҽ����Ҫʹ��
Private mstrEnd As String '���η��͵Ľ���ʱ��
Private mblnLimit As Boolean '���η��͸�ҩ;�������Ƿ��Խ���ʱ������

Private mlngNOSequence As Long
Private mlngҩƷ���ID As Long 'ҩƷ������ID
Private mlng�������ID As Long
Private mbln��ҩ�� As Boolean
Private mstr��ҩ�� As String
Private mstrAutoExe As String '����ִ���Զ����
Private mblnҽ������ As Boolean
Private mint���� As Integer
Private mstrLike As String
Private mstrRollNotify As String
Private mblnAutoVerify As Boolean   '����֮ǰ�Զ�У�ԣ�������ȡδУ�Ե�ҽ����
Private mblnChangeIF As Boolean     '�Ƿ�ı��˹ؼ������������¶�ȡҽ��
Private mblnCheckAdvice As Boolean  '����ת�Ƴ�Ժҽ��ǰ���δУ�Ի�δ��������
Private mlngRefModld As Long        '0����ҽ����1=����ҽ��
Private mobjCustom As CommandBarControlCustom
Private mobjlblMsg As CommandBarControl
Private mblnIsSendInterval As Boolean '�жϻ�ʿ�Ƿ���Ȩ���޸� ����ʱ����
Private mdatCurr As Date
Private mstrInfDepIDs As String  '����Һ�������ķ�ҩ�Ĳ��˲���
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mblnSendMsg As Boolean '��Ϣƽ̨�ǲ��ǿ���
Private mstr��Һ��ҩ;�� As String '��Һ�������Ĳ�������Һ�ĸ�ҩ;��
Private mint��Һ������Ч As Integer '������Һ�������ĵ�ҽ����Ч
Private mblnҩ���û� As Boolean '�Ƿ������ҩ���û�
Private mstrEndPoint As String '�����ڷ�ҩ���ͽ���ʱ��
Private mbln���鵥���������� As Boolean  '����ҽ��������������
Private mintסԺ��ҩ���� As Integer   '0-��ҩִ�п��ң�1-���˲���
Private mstrNoneIDs As String
Private mbln������ҩ As Boolean  'Ƥ��������ҩ �����������ô˲��������ж�Ƥ�Խ��������Ҫ��дƤ��������ҩ˵��
Private mstrAdDrugIDs As String '���һ���������˵����ҩƷ��ҽ��ID����
Private mstrҩƷ�۸�ȼ� As String '���˵�ҩƷ�۸�ȼ�
Private mstr���ļ۸�ȼ� As String '���˵����ļ۸�ȼ�
Private mstr��ͨ��Ŀ�۸�ȼ� As String '���˵���ͨ��Ŀ�۸�ȼ�
Private mbln�������Ѻ��� As Boolean
Private mintBnt As Integer  '�ٴ����������жϣ�-1��ʾδ��ѡ���.=vbCancelʱ��ʾ�жϣ�=vbIgnoreʱ��ʾ����

Private Enum COND_GROUP_IDX
    GP_�������� = 0
    GP_��ҩ;�� = 1
    GP_ҩ���û� = 2
    GP_������� = 3
End Enum
'----------------------------------------------
Private Enum COL_ADVICE
    COL_ѡ�� = 0
    COL_���� = 1
    COL_���� = 2
    COL_סԺ�� = 3
    COL_���� = 4
    COL_�ѱ� = 5
    COL_Ӥ�� = 6
    COL_ҽ����Ч = 7
    col_ҽ������ = 8
    COL_��� = 9
    COL_���� = 10
    COL_������λ = 11
    COL_���� = 12
    COL_������λ = 13
    COL_��� = 14
    COL_Ƶ�� = 15
    COL_�÷� = 16 '###
    COL_ҽ������ = 17 'Data���ڴ��ժҪ(ҽ��)
    COL_ִ��ʱ�� = 18   'ִ��ʱ�䷽����Data�д泤���Ŀ�ʼִ��ʱ��
    COL_�״�ʱ�� = 19
    COL_ĩ��ʱ�� = 20
    COL_ִ�п��� = 21
    COL_����ִ�� = 22
    COL_ִ������ = 23
    COL_����ID = 24 '������
    COL_��ҳID = 25
    col_�Ա� = 26
    COL_���� = 27
    COL_���� = 28
    COL_ID = 29
    COL_���ID = 30
    COL_���˲���ID = 31
    COL_���˿���ID = 32
    COL_��������ID = 33
    COL_����ҽ�� = 34
    COL_������� = 35
    COL_������ĿID = 36
    COL_�Ƽ����� = 37
    COL_ִ������ID = 38
    COL_ִ�п���ID = 39
    COL_ִ�б�� = 40
    COL_�շ�ϸĿID = 41
    COL_����ϵ�� = 42
    COL_סԺ��װ = 43
    COL_סԺ��λ = 44
    COL_�ɷ���� = 45
    COL_ҩ������ = 46 '###
    COL_�Ƿ��� = 47
    COL_��� = 48 '###
    COL_���� = 49
    COL_�ֽ�ʱ�� = 50
    COL_�������� = 51 '����ҽ��ר��
    COL_�Թܱ��� = 52
    COL_�걾��λ = 53
    COL_��鷽�� = 54
    COL_�������� = 55
    COL_������־ = 56
    COL_ҽ��״̬ = 57
    COL_ִ��Ƶ�� = 58
    COL_�¿�����ʱ�� = 59
    COL_���㷽ʽ = 60
    COL_ִ�а��� = 61
    COL_�������� = 62
    COL_����ʱ�� = 63
    COL_��ʼʱ�� = 64
    COL_������� = 65
    COL_ִ�з��� = 66
    COL_������� = 67
    COL_��ҩ���� = 68
    COL_����ҽ��ID = 69
End Enum
'-------------------------------------------------
Private Enum COL_PRICE
    COLP_�к� = 0
    COLP_�շ�ϸĿID = 1
    COLP_�̶� = 2
    COLP_��� = 3
    COLP_�Ƽ�ҽ�� = 4 '�ɼ���
    COLP_��� = 5
    COLP_�շ���Ŀ = 6
    COLP_�Ƽ����� = 7
    COLP_���� = 8
    COLP_���� = 9
    COLP_��λ = 10
    COLP_���� = 11
    COLP_Ӧ�ս�� = 12
    COLP_ʵ�ս�� = 13
    COLP_ִ�п��� = 14
    COLP_�������� = 15
    COLP_���� = 16
    COLP_�շѷ�ʽ = 17
    COLP_�շ���� = 18 '������
    COLP_ִ�п���ID = 19
    COLP_�������� = 20
    COLP_�������� = 21
End Enum

Private Const BackColorNew = &HD0FFFF   'ǳ��ɫ

Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strMainPrivs As String, _
    blnRefresh As Boolean, blnOnePati As Boolean, Optional ByVal lngҽ������ID As Long, Optional ByVal lngӤ������ID As Long, Optional ByRef objMip As Object) As Boolean
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    If lngӤ������ID <> 0 Then
        If lngӤ������ID = lngҽ������ID Then
            mlng����ID = lngӤ������ID
        End If
    End If
    mMainPrivs = strMainPrivs
    mblnOnePati = blnOnePati
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    
    If mclsMipModule Is Nothing Then
        mblnSendMsg = False
    ElseIf mclsMipModule.IsConnect Then
        mblnSendMsg = True
    Else
        mblnSendMsg = False
    End If
    
    On Error Resume Next
    Me.Show 1, frmParent
    
    blnRefresh = mblnRefresh
    ShowMe = mblnSend
End Function

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.value = vNewValue
        txtPer.Text = CInt(psb.value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property

Private Sub cboDrugType_Click()
    mblnChangeIF = True
End Sub

Private Sub cboTime_Click()
    Dim strTmp As String, lngTmp As Long
    
    dtpEnd.Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    mdatCurr = zlDatabase.Currentdate
    Select Case cboTime.ListIndex
    Case 0 '����
        dtpEnd.value = Format(mdatCurr, "yyyy-MM-dd 23:59:59")
    Case 1 '����
        dtpEnd.value = Format(mdatCurr + 1, "yyyy-MM-dd 23:59:59")
    Case 2 '����
        dtpEnd.value = Format(mdatCurr + 2, "yyyy-MM-dd 23:59:59")
    Case 3 '[ָ��..]
        strTmp = zlDatabase.GetPara("���ͽ���ʱ��", glngSys, pסԺҽ������, "23:59:59", Array(dtpEnd))
        lngTmp = Val(zlDatabase.GetPara("����ʱ����", glngSys, pסԺҽ������, "0", Array(dtpEnd)))
        mblnIsSendInterval = dtpEnd.Enabled
        dtpEnd.value = Format(mdatCurr + lngTmp, "yyyy-MM-dd " & strTmp)
        If dtpEnd.Enabled And dtpEnd.Visible Then dtpEnd.SetFocus
    End Select
    If mstrEndPoint <> "" And lblEatDrug.Visible Then
        lblEatDrug.Caption = "�����ڷ�ҩ����ʱ�䣺" & Format(dtpEnd.value, "yyyy-MM-dd") & " " & Mid(mstrEndPoint, 1, 5)
    End If
    mblnChangeIF = True
End Sub

Private Sub chkSendNow_Click()
    mblnChangeIF = True
End Sub

Private Sub cboUnit_Click()
'���ܣ���ȡָ����Χ�ڵĲ����б�
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str����IDs As String, lng����ID As Long
    Dim lngUnitID As Long
    Dim lngColor As Long
        
    On Error GoTo errH
    lvwPati.ListItems.Clear
    lngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
    
    If DeptIsWoman(0, Get����IDs(lngUnitID)) Then
        fraӤ��.Visible = True
        'ҽ������Χ
        optӤ��(Val(zlDatabase.GetPara("ҽ������Χ", glngSys, pסԺҽ������, "0"))).value = True
    Else
        fraӤ��.Visible = False
        optӤ��(0).value = True
    End If
    strSQL = "Select ���ò���,��������,����ֵ,������־1,������־2,������־3 From ���ʱ����� Where ����ID=[1]"
    Set mrsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUnitID)
        
    If Not mblnOnePati Then
        str����IDs = zlDatabase.GetPara("���Ͳ���", glngSys, pסԺҽ������)
       
        If str����IDs <> "" And InStr(str����IDs, ":") > 0 Then
            lng����ID = Val(Split(str����IDs, ":")(0))
            str����IDs = Split(str����IDs, ":")(1)
        End If
    End If
        
    If Me.Visible Then
        Set rsTmp = GetPatiRsByUnit(lngUnitID, 0, True, True, False)
    Else
        Set rsTmp = GetPatiRsByUnit(lngUnitID, mlng����ID, True, True, False)
    End If
    
    For i = 1 To rsTmp.RecordCount
        If Val(rsTmp!��˱�־ & "") < 1 Or gbyt������˷�ʽ <> 1 Then
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID, rsTmp!����)
            objItem.SubItems(1) = IIF(IsNull(rsTmp!סԺ��), "", rsTmp!סԺ��)
            objItem.SubItems(2) = IIF(IsNull(rsTmp!����), "", rsTmp!����)
            objItem.SubItems(3) = Format(NVL(rsTmp!ʣ���, 0), "0.00")
            objItem.SubItems(4) = IIF(IsNull(rsTmp!סԺҽʦ), "", rsTmp!סԺҽʦ)
            objItem.SubItems(5) = IIF(IsNull(rsTmp!�ѱ�), "", rsTmp!�ѱ�)
            objItem.SubItems(6) = IIF(IsNull(rsTmp!����ȼ�), "", rsTmp!����ȼ�)
            objItem.SubItems(7) = IIF(IsNull(rsTmp!����), "", rsTmp!����)
            objItem.SubItems(8) = Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm")
            objItem.SubItems(9) = NVL(rsTmp!��������)
            objItem.SubItems(10) = NVL(rsTmp!���ۺ�)
        
            '������Ϣ
            objItem.ListSubItems(1).Tag = NVL(rsTmp!���ò���)
            objItem.ListSubItems(2).Tag = NVL(rsTmp!������, 0)
            objItem.ListSubItems(3).Tag = NVL(rsTmp!����״̬, 0)
            objItem.ListSubItems(7).Tag = Val("" & rsTmp!����ID)
            objItem.ListSubItems(9).Tag = Val("" & rsTmp!��ҳID)
            
            '������ɫ
            lngColor = zlDatabase.GetPatiColor(NVL(rsTmp!��������))
            objItem.ListSubItems(1).ForeColor = lngColor
            objItem.ListSubItems(9).ForeColor = lngColor
            
            '�ϴ��Ƿ�ѡ��
            If lngUnitID = lng����ID And str����IDs <> "" Then
                If str����IDs = "ALL" _
                    Or Left(str����IDs, 1) <> "-" And InStr("," & str����IDs & ",", "," & rsTmp!����ID & ",") > 0 _
                    Or Left(str����IDs, 1) = "-" And InStr("," & Mid(str����IDs, 2) & ",", "," & rsTmp!����ID & ",") = 0 Then
                    objItem.Checked = True
                    If k = 0 Then 'Ϊ�˿�����ѡ���
                        objItem.EnsureVisible
                        objItem.Selected = True
                        k = 1
                    End If
                End If
            '��Ժ���˺���ת������ͨ��ҽ�����ѽ���
            ElseIf rsTmp!����ID = mlng����ID Then
                objItem.Checked = True 'ȱʡֻѡ��ǰ����
                objItem.EnsureVisible
                objItem.Selected = True
            End If
        End If
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncExecSend()
'���ܣ�����ҽ������
'������lngModle=0 ���ͳ���ҽ����=1���ͱ���ҽ��
    Dim lng���ͺ� As Long, i As Long
    Dim objCbo As CommandBarComboBox
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_ID)) <> 0 And .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                Exit For
            End If
        Next
        If i > .Rows - 1 Then
            MsgBox "��ǰû�п��Է��͵�ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    If mblnChangeIF Then
        MsgBox "ҽ�����͵������Ѹı䣬���Զ����¶�ȡ���ݣ�������ٷ��͡�", vbInformation, gstrSysName
        Call RefreshData(mlngRefModld)
        Exit Sub
    End If
    
    'ִ�з���
    lng���ͺ� = SendAdvice
    If lng���ͺ� <> 0 Then
        mblnSend = True
        
        '����������ҽ��ʱ��鲢���ѳ����ջ�(�Զ�)ֹͣ��ҽ��
        If mstrRollNotify <> "" Then
            Call ShowRollNotify(mstrRollNotify)
        End If
        
        'ʹ��������ҩ�ŵĴ���
        If mstr��ҩ�� <> "" Then
            Set objCbo = cbsMain.FindControl(, conMenu_View_Find)
            i = objCbo.FindItem(mstr��ҩ��)
            If i = 0 Then
                objCbo.AddItem mstr��ҩ��, 2
                objCbo.ListIndex = 2
            End If
        End If
        
        '��ӡ���Ƶ���
        Call frmSendBillPrint.ShowMe(lng���ͺ�, 2, Me)
    End If
End Sub

Private Sub FuncExecReStop()
'���ܣ�����ȷ��ֹͣ
    Dim str����IDs As String, i As Long
    
    If lvwPati.Visible Then
        mlng����ID = cboUnit.ItemData(cboUnit.ListIndex)
        For i = 1 To lvwPati.ListItems.Count
            If lvwPati.ListItems(i).Checked Then
                str����IDs = str����IDs & ";" & Mid(lvwPati.ListItems(i).Key, 2) & "," & Val(lvwPati.ListItems(i).ListSubItems(9).Tag)
            End If
        Next
        If str����IDs = "" Then
            MsgBox "������ѡ��һ����Ҫȷ��ֹͣҽ���Ĳ��ˡ�", vbInformation, gstrSysName
            If lvwPati.Visible And lvwPati.Enabled Then lvwPati.SetFocus: Exit Sub
        Else
            str����IDs = Mid(str����IDs, 2)
        End If
    End If
    
    Call frmAdviceOperate.ShowMe(Me, mMainPrivs, 2, mlng����ID, mlng��ҳID, mlng����ID, 0, True, , , , , str����IDs = "", str����IDs, True, , , mlng����ID, , mclsMipModule)
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Long
    
    If Not Control.Visible Then Exit Sub
    
    Select Case Control.ID
    Case conMenu_View_Refresh  '��ȡ����ҽ��
        If Me.ActiveControl Is dtpEnd Then cboTime.SetFocus
        Call RefreshData(0)
        mobjCustom.Visible = False
        mobjlblMsg.Visible = False
    Case conMenu_View_RefreshSpare  '��ȡ����ҽ��
        If Me.ActiveControl Is dtpEnd Then cboTime.SetFocus
        Call RefreshData(1)
        mobjCustom.Visible = True
        mobjlblMsg.Visible = True
    Case conMenu_Edit_ReStop    'ȷ��ֹͣ
        Call FuncExecReStop
    Case conMenu_Edit_Send      '����
        Call FuncExecSend
        
    Case conMenu_View_Show
        tkpMain.Visible = True
        fraLR.Visible = True
        Call Form_Resize
    Case conMenu_View_Hide
        tkpMain.Visible = False
        fraLR.Visible = False
        Call Form_Resize
    Case conMenu_Edit_SelAll
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 Then
                    If Not (InStr(mstrNoneIDs, "," & .TextMatrix(i, COL_ID) & ",") > 0 And Not mbln������ҩ) Then
                        If Not (Val(.TextMatrix(i, COL_��������)) = 5 And .TextMatrix(i, COL_�������) = "Z") Then
                            Set .Cell(flexcpPicture, i, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
                        End If
                    End If
                End If
            Next
        End With
        Call ShowSendTotal
    Case conMenu_Edit_ClsAll
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 Then
                    Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                End If
            Next
        End With
        Call ShowSendTotal
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim lngLW As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    Me.tkpMain.Left = lngLeft
    Me.tkpMain.Top = lngTop
    Me.tkpMain.Height = lngBottom - lngTop - stbThis.Height
    
    Me.fraLR.Left = lngLeft + tkpMain.Width
    Me.fraLR.Top = lngTop
    Me.fraLR.Height = lngBottom - lngTop - stbThis.Height
    
    If tkpMain.Visible Then
        lngLW = fraLR.Width + tkpMain.Width
    End If
    
    fraInfo.Top = lngTop
    fraInfo.Left = lngLeft + lngLW
    fraInfo.Width = lngRight - lngLeft - lngLW
    
    vsAdvice.Left = lngLeft + lngLW
    vsAdvice.Top = fraInfo.Top + fraInfo.Height
    vsAdvice.Width = lngRight - lngLeft - lngLW
    vsAdvice.Height = lngBottom - lngTop - fraInfo.Height - vsPrice.Height - fraUD.Height - stbThis.Height
    
    fraUD.Top = vsAdvice.Top + vsAdvice.Height
    fraUD.Left = vsAdvice.Left
    fraUD.Width = vsAdvice.Width
    
    vsPrice.Left = vsAdvice.Left
    vsPrice.Top = fraUD.Top + fraUD.Height
    vsPrice.Width = vsAdvice.Width
    
    psb.Top = stbThis.Top + Screen.TwipsPerPixelY * 4
    psb.Left = stbThis.Panels(2).Left + Screen.TwipsPerPixelX * 2
    psb.Width = stbThis.Panels(2).Width - txtPer.Width - Screen.TwipsPerPixelX * 7
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
       
    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Show
        Control.Visible = Not tkpMain.Visible
    Case conMenu_View_Hide
        Control.Visible = tkpMain.Visible
    Case conMenu_View_Find
        Control.Visible = (chkҩƷ��Ч(0).value = 1 Or chkҩƷ��Ч(1).value = 1) And mbln��ҩ��
    Case conMenu_Edit_ReStop
        If InStr(GetInsidePrivs(pסԺҽ������), ";ҽ��ȷ��ֹͣ;") = 0 Then Control.Visible = False
    End Select
End Sub

Private Sub chk������Ч_Click(Index As Integer)
    Call SetFaceByҽ����Ч
End Sub

Private Sub chkҩƷ��Ч_Click(Index As Integer)
    Call SetFaceByҽ����Ч
End Sub

Private Sub cmdAllPati_Click()
    Call SelectLVW(lvwPati, True)
    lvwPati.SetFocus
End Sub

'��ѡ
Private Sub cmdUnWay_Click()
    Dim i As Long
    For i = 1 To lvwWay.ListItems.Count
        lvwWay.ListItems(i).Checked = Not lvwWay.ListItems(i).Checked
    Next
    lvwWay.SetFocus
End Sub

Private Sub cmdAllWay_Click()
    Call SelectLVW(lvwWay, True)
    lvwWay.SetFocus
End Sub

Private Sub cmdNoPati_Click()
    Call SelectLVW(lvwPati, False)
    lvwPati.SetFocus
End Sub

Private Sub cmdNoWay_Click()
    Call SelectLVW(lvwWay, False)
    lvwWay.SetFocus
End Sub

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
End Sub

Private Sub cmdQuick_Click()
    Dim i As Long, blnDo As Boolean
    
    If mrsWarn Is Nothing Then Exit Sub
    
    With lvwPati
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked Then
                'ֻ�����ۼƱ����������д���
                mrsWarn.Filter = "��������=1 And ���ò���='" & .ListItems(i).ListSubItems(1).Tag & "'"
                If Not mrsWarn.EOF Then
                    blnDo = False
                    Select Case BeSureMode(NVL(mrsWarn!������־1), NVL(mrsWarn!������־2), NVL(mrsWarn!������־3))
                    Case 1 '���ڱ���ֵ(����Ԥ����ľ�)��ʾѯ�ʼ���
                        blnDo = Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) <= 0
                    Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
                        blnDo = Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) <= 0
                    Case 3 '���ڱ���ֵ��ֹ����
                        blnDo = Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) < NVL(mrsWarn!����ֵ, 0)
                    End Select
                    If blnDo Then
                        .ListItems(i).Checked = False
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub cmdִ�п���_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    strSQL = _
        " Select 0 as ID,'-' as ����,'���п���' as ����,NULL as ���� From Dual" & _
        " Union ALL" & _
        " Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.������� IN(2,3)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " Order by ����"
    vRect = zlControl.GetControlRect(txtִ�п���.hwnd)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "ִ�п���", , , , , , True, vRect.Left, vRect.Top, txtִ�п���.Height, blnCancel, , True)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "û�п��õĿ��ң����ȵ����Ź��������á�", vbInformation, gstrSysName
        End If
        txtִ�п���.Text = txtִ�п���.Tag
        Call zlControl.TxtSelAll(txtִ�п���)
    Else
        txtִ�п���.Text = rsTmp!����
        txtִ�п���.Tag = rsTmp!����
        cmdִ�п���.Tag = rsTmp!ID
    End If
    txtִ�п���.SetFocus
End Sub

Private Sub dkpExecTime_Change()
    Call RefreshData(1)
End Sub

Private Sub dtpEnd_Change()
    mblnChangeIF = True
End Sub

Private Sub dtpEnd_Validate(Cancel As Boolean)
    If mstrEndPoint <> "" And lblEatDrug.Visible Then
        lblEatDrug.Caption = "�����ڷ�ҩ����ʱ�䣺" & Format(dtpEnd.value, "yyyy-MM-dd") & " " & Mid(mstrEndPoint, 1, 5)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If Me.ActiveControl Is lvwPati Then
            Call cmdAllPati_Click
        ElseIf Me.ActiveControl Is lvwWay Then
            Call cmdAllWay_Click
        ElseIf Me.ActiveControl Is lstClass Then
            j = lstClass.ListIndex
            For i = 0 To lstClass.ListCount - 1
                lstClass.Selected(i) = True
            Next
            lstClass.ListIndex = j
        Else
            cbsMain.FindControl(, conMenu_Edit_SelAll).Execute
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If Me.ActiveControl Is lvwPati Then
            Call cmdNoPati_Click
        ElseIf Me.ActiveControl Is lvwWay Then
            Call cmdNoWay_Click
        ElseIf Me.ActiveControl Is lstClass Then
            j = lstClass.ListIndex
            For i = 0 To lstClass.ListCount - 1
                lstClass.Selected(i) = False
            Next
            lstClass.ListIndex = j
        Else
            cbsMain.FindControl(, conMenu_Edit_ClsAll).Execute
        End If
    ElseIf KeyCode = vbKeyQ And Shift = vbCtrlMask Then
        If cmdQuick.Visible And cmdQuick.Enabled Then Call cmdQuick_Click
    ElseIf KeyCode = 13 Then
        If Not ActiveControl Is vsDept And Not ActiveControl Is vsAdvice _
            And Not ActiveControl Is vsPrice And Not Me.ActiveControl Is txtִ�п��� Then
            Call zlcommfun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ActiveControl Is vsDept And Not ActiveControl Is vsAdvice _
            And Not ActiveControl Is vsPrice And Not Me.ActiveControl Is txtִ�п��� Then
            KeyAscii = 0
        End If
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub SetPic����_BackColor()
    Dim i As Integer
    
    fraҩƷ.BackColor = pic����.BackColor
    fra����.BackColor = pic����.BackColor
    fra״̬.BackColor = pic����.BackColor
    fraӤ��.BackColor = pic����.BackColor
    For i = 0 To 2
        If i <> 2 Then
            chkҩƷ��Ч(i).BackColor = pic����.BackColor
            chk������Ч(i).BackColor = pic����.BackColor
        End If
        opt״̬(i).BackColor = pic����.BackColor
        optӤ��(i).BackColor = pic����.BackColor
    Next
    chkSendNow.BackColor = pic����.BackColor
    fraҽ������.BackColor = pic����.BackColor
    chk�Ӱ�Ӽ�.BackColor = pic����.BackColor
    
    fraPati.BackColor = pic����.BackColor
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    
    '������----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlcommfun.GetPubIcons
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "����")
        objControl.IconId = conMenu_View_Show
        objControl.ToolTipText = "���ط�����������"
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, "��ʾ")
        objControl.IconId = conMenu_View_Hide
        objControl.ToolTipText = "��ʾ������������"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "ȫѡ")
        objControl.BeginGroup = True
        objControl.ToolTipText = "ѡ�����п��Է��͵�ҽ��(Ctrl+A)"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "ȫ��")
        objControl.ToolTipText = "���������ѡ����ҽ����ѡ��״̬(Ctrl+R)"
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "��ȡ����ҽ��"): objControl.BeginGroup = True
        objControl.ToolTipText = "���ݵ�ǰ������ȡ���淢��ҽ��"
        Set objControl = .Add(xtpControlButton, conMenu_View_RefreshSpare, "��ȡ����ҽ��")
        objControl.ToolTipText = "���ݵ�ǰ������ȡ���÷���ҽ��"
        Set mobjlblMsg = .Add(xtpControlLabel, conMenu_View_RefreshSpare * 100# + 1, "��ִ��ʱ�䡿:")
        mobjlblMsg.Visible = False
        Set mobjCustom = .Add(xtpControlCustom, conMenu_View_RefreshSpare * 100# + 2, "")
        mobjCustom.ToolTipText = "������ñ���ҽ��ִ�е�ʱ�䡣"
        dkpExecTime.value = zlDatabase.Currentdate
        mobjCustom.Handle = dkpExecTime.hwnd
        mobjCustom.Visible = False
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "����ҽ��"): objControl.BeginGroup = True
        objControl.ToolTipText = "����������ѡ���ҽ��"
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReStop, "ȷ��ֹͣ")
        objControl.ToolTipText = "ȷ����ֹͣ��ҽ��"
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyE, conMenu_Edit_Send
        .Add 0, vbKeyF1, conMenu_Help_Help
        .Add 0, vbKeyF5, conMenu_View_Refresh
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add 0, vbKeyEscape, conMenu_File_Exit
    End With
    
    '���˵��Ҳ����ҩ��
    objBar.EnableDocking xtpFlagStretched
    With objBar.Controls
        Set objCbo = .Add(xtpControlComboBox, conMenu_View_Find, "��ҩ��")
            objCbo.BeginGroup = True
            objCbo.Flags = xtpFlagRightAlign
            objCbo.Style = xtpComboLabel
            objCbo.Width = 200
    End With
End Sub

Private Sub Form_Load()
    Dim objPane As Pane
    Dim objGroup As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem, blnDo As Boolean, i As Long
    Dim strPar As String
    
    If Not PatiFeeUsable(mlng����ID, mlng��ҳID) Then Unload Me: Exit Sub
    Call InitAdviceTable
    Call InitPriceTable
    fraLR.BackColor = Me.BackColor
    fraUD.BackColor = Me.BackColor
    
    mblnChangeIF = False
    mblnSend = False
    mblnRefresh = False
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ")) '����ƥ�䷽ʽ��0-ƴ��,1-���
    mstrAutoExe = zlDatabase.GetPara("����ִ���Զ����", glngSys, pסԺҽ������)
    mblnҽ������ = Val(zlDatabase.GetPara("ҽ��ҽ����������", glngSys, pסԺҽ������)) <> 0
    mbln��ҩ�� = Val(zlDatabase.GetPara(27, glngSys)) <> 0
    mblnAutoVerify = Val(zlDatabase.GetPara("����ǰ�Զ�У��", glngSys, pסԺҽ������, 0)) = 1
    mblnCheckAdvice = Val(zlDatabase.GetPara("����ҽ������ǰ���δ��Чҽ��", glngSys, pסԺҽ������, 0)) = 1
    mblnLimit = Val(zlDatabase.GetPara("ҩ���������ƽ���ʱ��", glngSys, pסԺҽ������, 0)) = 1
    mbln���鵥���������� = Val(zlDatabase.GetPara("����ҽ��������������", glngSys, pסԺҽ������, "0")) = 1
    mintסԺ��ҩ���� = Val(zlDatabase.GetPara("סԺ��ҩ����", glngSys, pסԺҽ������, "0"))
    mstrInfDepIDs = zlDatabase.GetPara("��Դ����", glngSys, p��Һ��������, "")
    mstr��Һ��ҩ;�� = zlDatabase.GetPara("��Һ��ҩ;��", glngSys, p��Һ��������)
    mint��Һ������Ч = Val(zlDatabase.GetPara("ҽ������", glngSys, p��Һ��������, "1")) - 1
    
    mstrEndPoint = zlDatabase.GetPara("�����ڷ�ҩ���ͽ���ʱ��", glngSys, pסԺҽ������, "0")
    If Mid(mstrEndPoint, 1, 1) = 1 And InStr(mstrEndPoint, "|") > 0 Then
        mstrEndPoint = Split(mstrEndPoint, "|")(1)
    Else
        mstrEndPoint = ""
    End If
     
    mbln������ҩ = Val(zlDatabase.GetPara("Ƥ��������ҩ", glngSys, pסԺҽ���´�)) <> 0
    
    If DeptIsWoman(0, Get����IDs(mlng����ID)) Then
        fraӤ��.Visible = True
        'ҽ������Χ
        optӤ��(Val(zlDatabase.GetPara("ҽ������Χ", glngSys, pסԺҽ������, "0"))).value = True
    End If
    
    '������ģʽ
    '-----------------------------------------
    If mblnOnePati Then '����ʾ���ˣ������ڼӷ���֮ǰ�ı�pic�����ĸ߶�
        fraPati.Visible = False
        pic����.Height = pic����.Height - fraPati.Height + 60
    End If
    
    Call InitCommandBar
    
    '����ؼ�------------------------------------------
    'tkpMain.VisualTheme = xtpTaskPanelThemeListViewOffice2003
    
    Call tkpMain.SetMargins(0, 0, 0, 0, 0)
    Call tkpMain.SetItemInnerMargins(0, 0, 0, 0)
    Call tkpMain.SetItemOuterMargins(0, 0, 0, 0)
    Call tkpMain.SetGroupInnerMargins(0, 0, 0, 0)
    Call tkpMain.SetGroupOuterMargins(3, 5, 3, 0)
        
    Set objGroup = tkpMain.Groups.Add(GP_��������, "��������")
    objGroup.Expandable = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic����
    pic����.BackColor = objItem.BackColor
    Call SetPic����_BackColor
    
    Set objGroup = tkpMain.Groups.Add(GP_��ҩ;��, "��ҩ;��")
    objGroup.Expanded = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic��ҩ
    pic��ҩ.BackColor = objItem.BackColor
    
    Set objGroup = tkpMain.Groups.Add(GP_ҩ���û�, "ҩ���û�")
    objGroup.Expanded = True
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picҩ��
    picҩ��.BackColor = objItem.BackColor
    chkDeptExchange.BackColor = objItem.BackColor
    
    Set objGroup = tkpMain.Groups.Add(GP_�������, "�������")
    objGroup.Expanded = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic���
    pic���.BackColor = objItem.BackColor
    
        
        
    '��ʼ��ȡһЩ����---------------------------------
    '�����ⷿҩƷ�����鷽ʽ,�������ϲ���
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    Call Initҩ���û�
    
    'ҩƷ�������
    mlng�������ID = ExistIOClass(41) '����ȷ���Ƿ�ʹ���������շ�,�������ж�
    mlngҩƷ���ID = ExistIOClass(9)
    If mlngҩƷ���ID = 0 Then
        MsgBox "����ȷ��ҩƷ�������ݵ�������,���ȵ���������������ã�", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    '��ʼ����������-----------------------------------
    cboDrugType.AddItem "0-ȫ��"
    cboDrugType.AddItem "1-��Ʒ��"
    cboDrugType.AddItem "2-����;���I��"
    cboDrugType.AddItem "3-����(��1��2��)"
    cboDrugType.ListIndex = 0
    
    'ȱʡ����ʱ��
    cboTime.AddItem "1-����"
    cboTime.AddItem "2-����"
    cboTime.AddItem "3-����"
    cboTime.AddItem "4-ָ��"
    
    '��ȡ����ҩ��
    Call Loadҩ��
    
    'ȱʡִ�п���
    txtִ�п���.Text = "���п���"
    txtִ�п���.Tag = txtִ�п���.Text
    cmdִ�п���.Tag = ""
        
    '����/����
    Call InitUnits
                        
    '��ҩ;��
    Call Load��ҩ;��
    
    'ҩ���û�
    Call Showҩ��
        
    '�������
    Call Load�������
        
    '������ҽ������--------------------------------
    Call Initҽ����Ҫ����
    
    
    '��У�Լ�����ģʽ��˫��ҽ������ʱ���ã�
    If mblnAutoVerify Then
        If mblnOnePati Then
            '������˴�����������ҽ�������Զ�����У�Դ���:��������ȼ�,����/Σҽ��,��ǰ����ҽ��������,��¼�����,ת�ƣ���Ժ��תԺ������
            blnDo = CheckSpecialAdvice(mlng����ID, mlng��ҳID)
            If blnDo Then
                mblnRefresh = True  '����ȼ��򲡿�У�Ժ���Ҫˢ�²����嵥
                If frmAdviceOperate.ShowMe(Me, mMainPrivs, 3, mlng����ID, mlng��ҳID, mlng����ID, 0, True, False, 0, , , mblnOnePati, , , , , mlng����ID, True, mclsMipModule) = False Then
                    'ȡ���˳�
                    Unload Me: Exit Sub
                End If
            End If
            
            Call RefreshData(0)
            
            If blnDo Then   '����������ҽ��ʱ����ʹû�����ݣ�Ҳ���˳�����Ϊ������֮ǰ���������ֵ�����������
                If vsAdvice.Rows = vsAdvice.FixedRows + 1 Then
                    If Val("" & vsAdvice.TextMatrix(vsAdvice.Rows - 1, COL_ID)) = 0 Then
                        mblnRefresh = True
                        Unload Me: Exit Sub
                    End If
                End If
            End If
        Else
            For i = 1 To lvwPati.ListItems.Count
                If lvwPati.ListItems(i).Checked Then
                    Exit For
                End If
            Next
            If i <= lvwPati.ListItems.Count Then
                '�Զ���ȡ�¿�ҽ��
                Call RefreshData(0)
            End If
        End If
    End If

    '-------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    Me.WindowState = vbMaximized
    
End Sub

Private Function Load�������() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str���s As String
    
    On Error GoTo errH
    
    str���s = zlDatabase.GetPara("���������������", glngSys, pסԺҽ������, "", Array(lstClass))
    
    strSQL = "Select ����,���� From ������Ŀ��� Where ���� Not IN('5','6','7','8','9') Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        lstClass.AddItem rsTmp!����
        lstClass.ItemData(lstClass.NewIndex) = Asc(rsTmp!����)
        If str���s <> "" Then
            If InStr(str���s, "'" & rsTmp!���� & "'") > 0 Then
                lstClass.Selected(lstClass.NewIndex) = True
            End If
        Else
            lstClass.Selected(lstClass.NewIndex) = True
        End If
        rsTmp.MoveNext
    Next
    If lstClass.ListCount > 0 Then lstClass.ListIndex = 0
    Load������� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
    
    '��������۲���
    If InStr(mMainPrivs, "ȫԺ����") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
    Else
        '����Ȩ������ֱ�����ڲ���+���ڿ�����������
        strSQL = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.����,C.����,Nvl(B.ȱʡ,0) as ȱʡ" & _
            " From �������Ҷ�Ӧ A,������Ա B,���ű� C" & _
            " Where A.����ID=C.ID And B.����ID=A.����ID And B.��ԱID=[1]" & _
            " And Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=A.����ID)" & _
            " And Not Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=A.����ID)" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSQL & ") Group by ID,����,���� Order by ����"
    End If
    
    cboUnit.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng����ID Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Load��ҩ;��() As Boolean
    Dim strSQL As String, i As Long
    Dim objItem As ListItem, str��ҩIDs As String
    Dim rs��ҩ;�� As ADODB.Recordset
    
    On Error GoTo errH
    str��ҩIDs = zlDatabase.GetPara("ҩ�����͸�ҩ;��", glngSys, pסԺҽ������, "", Array(lvwWay))
    
    Set rs��ҩ;�� = New ADODB.Recordset
    strSQL = "Select ID,����,����,ִ�з��� From ������ĿĿ¼ Where ���='E' And ��������='2' And ������� in(2,3) And (վ��='" & gstrNodeNo & "' Or վ�� is Null) Order by ����"
    
    Call zlDatabase.OpenRecordset(rs��ҩ;��, strSQL, Me.Caption)
        
    With rs��ҩ;��
        lvwWay.ListItems.Clear
        For i = 1 To .RecordCount
            Set objItem = lvwWay.ListItems.Add(, "_" & !ID, !���� & "-" & !����)
            
            If str��ҩIDs <> "" Then
                If InStr("," & str��ҩIDs & ",", "," & !ID & ",") > 0 Then
                    objItem.Checked = True
                End If
            Else
                objItem.Checked = True
            End If
            .MoveNext
        Next
    End With
    
    Load��ҩ;�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Showҩ��()
    Dim strTmp As String, i As Long, j As Long
    Dim str�û� As String, arr�û� As Variant
    Dim lngԭҩ�� As Long
    
    str�û� = zlDatabase.GetPara("ҩ������ҩ���û�", glngSys, pסԺҽ������, "", Array(vsDept))
    arr�û� = Split(str�û�, ",")
    
    If UBound(arr�û�) >= 0 Then
        If arr�û�(0) Like "����-*" Then
            str�û� = Replace(str�û�, CStr(arr�û�(0)), "")
            If Mid(str�û�, 1, 1) = "," Then str�û� = Mid(str�û�, 2)
            arr�û� = Split(str�û�, ",")
        End If
    End If
    
    mrsҩ��.Filter = 0
    If Not mrsҩ��.EOF Then
        vsDept.Rows = vsDept.FixedRows + mrsҩ��.RecordCount
        For i = 1 To mrsҩ��.RecordCount
            vsDept.Cell(flexcpData, i, 0) = CLng(mrsҩ��!ID)
            vsDept.TextMatrix(i, 0) = mrsҩ��!���� & "-" & mrsҩ��!����
            strTmp = strTmp & "|#" & mrsҩ��!ID & ";" & mrsҩ��!���� & "-" & mrsҩ��!����
            
            mrsҩ��.MoveNext
        Next
        
        
        For i = 1 To vsDept.Rows - 1
            lngԭҩ�� = CLng(vsDept.Cell(flexcpData, i, 0))
            mrsҩ��.Filter = "ID=" & lngԭҩ��
            For j = 0 To UBound(arr�û�)
                If arr�û�(j) Like mrsҩ��!ID & "-*" Then Exit For
            Next
            If j <= UBound(arr�û�) Then
                mrsҩ��.Filter = "ID=" & Val(Split(arr�û�(j), "-")(1))
                If Not mrsҩ��.EOF Then
                    vsDept.Cell(flexcpData, i, 1) = CLng(mrsҩ��!ID)
                    mrsҩ��.Filter = "ID=" & lngԭҩ��
                    mrsҩ��!��ID = CLng(vsDept.Cell(flexcpData, i, 1))
                    mrsҩ��.Update
                Else
                    vsDept.Cell(flexcpData, i, 1) = CLng(mrsҩ��!��ID)
                End If
            Else
                vsDept.Cell(flexcpData, i, 1) = CLng(mrsҩ��!��ID)
            End If
            
            mrsҩ��.Filter = "ID=" & CLng(vsDept.Cell(flexcpData, i, 1))
            vsDept.TextMatrix(i, 1) = mrsҩ��!���� & "-" & mrsҩ��!����
            vsDept.Cell(flexcpBackColor, i, 1) = COLEditBackColor
        Next
        If strTmp <> "" Then vsDept.ColComboList(1) = Mid(strTmp, 2)
    Else
        vsDept.Rows = vsDept.FixedRows + 1
        vsDept.Editable = flexEDNone
    End If
    vsDept.Row = vsDept.FixedRows: vsDept.Col = 1
    Call vsDept_AfterRowColChange(-1, -1, vsDept.Row, vsDept.Col)
    chkDeptExchange.Enabled = InStr(GetInsidePrivs(pסԺҽ������), ";�����û�ҩ��;") > 0
    picҩ��.Enabled = chkDeptExchange.Enabled
End Sub

Private Function Loadҩ��() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    cboҩ��.Clear
    cboҩ��.AddItem "����ҩ��"
    cboҩ��.ListIndex = 0
    
    On Error GoTo errH
    
    strSQL = _
        "Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID And B.������� IN(2,3) and B.�������� in('��ҩ��','��ҩ��','��ҩ��')" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        cboҩ��.AddItem rsTmp!���� & "-" & rsTmp!����
        cboҩ��.ItemData(cboҩ��.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    Loadҩ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function TheStockCheck(ByVal lng�ⷿID As Long, ByVal str��� As String) As Integer
'���ܣ���ȡָ���ⷿ�ĳ������鷽ʽ
    Dim intStyle As Integer
    On Error Resume Next
    If InStr(",5,6,7,", str���) > 0 Then
        intStyle = mcolStock1("_" & lng�ⷿID)
    ElseIf str��� = "4" Then
        intStyle = mcolStock2("_" & lng�ⷿID)
    End If
    err.Clear: On Error GoTo 0
    TheStockCheck = intStyle
End Function

Private Function Initҩ���û�() As Boolean
'���ܣ�'��ʼ��ȡһЩ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    '��ȡ����ҩ����������:����ҩ���û�
    Set mrsҩ�� = New ADODB.Recordset
    mrsҩ��.Fields.Append "ID", adBigInt
    mrsҩ��.Fields.Append "����", adVarChar, 100
    mrsҩ��.Fields.Append "����", adVarChar, 200
    mrsҩ��.Fields.Append "��ID", adBigInt
    mrsҩ��.CursorLocation = adUseClient
    mrsҩ��.LockType = adLockOptimistic
    mrsҩ��.CursorType = adOpenStatic
    mrsҩ��.Open
    
    strSQL = _
        "Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " And B.����ID=A.ID And B.������� IN(2,3) and B.�������� in('��ҩ��','��ҩ��','��ҩ��')" & _
        " Order by A.����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        mrsҩ��.AddNew
        mrsҩ��!ID = rsTmp!ID
        mrsҩ��!���� = rsTmp!����
        mrsҩ��!���� = rsTmp!����
        mrsҩ��!��ID = rsTmp!ID
        mrsҩ��.Update
        rsTmp.MoveNext
    Next
    
    Initҩ���û� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    '�ͷ�˽�м�IN����
    mMainPrivs = ""
    mlng����ID = 0
    mlng����ID = 0
    mstrInfDepIDs = ""
    mstrEnd = ""
    mblnLimit = False
    
    mlngҩƷ���ID = 0
    mlng�������ID = 0
    Set mrsҩ�� = Nothing
    Set mrsBill = Nothing
    Set mrsWarn = Nothing
    Set mcolStock1 = Nothing
    Set mcolStock2 = Nothing
    Set mobjCustom = Nothing
    Set mobjlblMsg = Nothing
    Set mclsMipModule = Nothing
    gbln�Ӱ�Ӽ� = False
End Sub

Private Sub Refresh��ҩ��()
    Dim objCbo As CommandBarComboBox
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strPre As String
    
    On Error GoTo errH
    
    Set objCbo = cbsMain.FindControl(, conMenu_View_Find)
    
    If objCbo.ListIndex > 0 Then strPre = objCbo.List(objCbo.ListIndex)
    
    objCbo.Clear
    objCbo.AddItem "<ʹ���µ���ҩ��>"
    objCbo.ListIndex = 1
    
    strSQL = "Select Distinct ��ҩ�� From δ��ҩƷ��¼ Where ��������>=Trunc(Sysdate) And ����=9 And �Է�����ID=[1] And ��ҩ�� is Not NULL Order by ��ҩ�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    Do While Not rsTmp.EOF
        objCbo.AddItem rsTmp!��ҩ��
        If rsTmp!��ҩ�� = strPre Then
            objCbo.ListIndex = objCbo.ListCount
        End If
        rsTmp.MoveNext
    Loop

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get��ҩ��() As String
    Dim objCbo As CommandBarComboBox
    
    Set objCbo = cbsMain.FindControl(, conMenu_View_Find)
    If objCbo.ListIndex = 1 Then
        Get��ҩ�� = zlDatabase.GetNextNo(122, mlng����ID)
    ElseIf objCbo.ListIndex > 1 Then
        Get��ҩ�� = objCbo.List(objCbo.ListIndex)
    End If
End Function

Private Sub RefreshData(ByVal lngRefModld As Long)
'���ܣ����÷�������
'������lngModle=0 ��ȡ����ҽ����=1��ȡ����ҽ��
    Dim strSel As String, strUnSel As String
    Dim str����IDs, str��ҳIDs As String, str��ҩIDs As String
    Dim str���s As String, strTmp As String
    Dim i As Long
    Dim str���˿���IDs As String
    
    '���ͻ�ȡ����
    '--------------------------------------------------------------------------------
    mlngRefModld = lngRefModld
    If cboUnit.ListIndex = -1 Then
        MsgBox "��ѡ��һ��������", vbInformation, gstrSysName
        If cboUnit.Visible Then cboUnit.SetFocus: Exit Sub
    End If
    If chkҩƷ��Ч(0).value = 0 And chkҩƷ��Ч(1).value = 0 And chk������Ч(0).value = 0 And chk������Ч(1).value = 0 Then
        MsgBox "��ѡ��һ��ҽ����Ч��", vbInformation, gstrSysName
        Exit Sub
        '����û��Ȩ�ޣ�ĳЩ��Ч���ɼ����򲻿�����
    End If
    
    
    '��Ч��ʱ��
    If chkҩƷ��Ч(0).value = 1 Or chk������Ч(0).value = 1 Then
        mstrEnd = Format(dtpEnd.value, "yyyy-MM-dd HH:mm:ss")
    Else
        mstrEnd = ""
    End If
    
    'סԺ����
    mlng����ID = cboUnit.ItemData(cboUnit.ListIndex)
    str����IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            If Val(lvwPati.ListItems(i).ListSubItems(3).Tag & "") = psԤ�� Or Val(lvwPati.ListItems(i).ListSubItems(3).Tag & "") = ps��Ժ Then
                Call MsgBox("����""" & lvwPati.ListItems(i) & """��" & IIF(Val(lvwPati.ListItems(i).ListSubItems(3).Tag & "") = psԤ��, "Ԥ", "") & "��Ժ�����������ҽ�����ͣ�", vbInformation, gstrSysName)
                lvwPati.ListItems(i).Checked = False
                Exit Sub
            End If
            str����IDs = str����IDs & "," & Mid(lvwPati.ListItems(i).Key, 2)
            strSel = strSel & "," & Mid(lvwPati.ListItems(i).Key, 2)
            str��ҳIDs = str��ҳIDs & "," & lvwPati.ListItems(i).ListSubItems(9).Tag
            str���˿���IDs = str���˿���IDs & "," & lvwPati.ListItems(i).ListSubItems(7).Tag
        Else
            strUnSel = strUnSel & "," & Mid(lvwPati.ListItems(i).Key, 2)
        End If
    Next
    str����IDs = Mid(str����IDs, 2)
    str��ҳIDs = Mid(str��ҳIDs, 2)
    str���˿���IDs = Mid(str���˿���IDs, 2)
    If str����IDs = "" Then
        MsgBox "������ѡ��һ����Ҫ����ҽ�����ˡ�", vbInformation, gstrSysName
        If lvwPati.Visible And lvwPati.Enabled Then lvwPati.SetFocus: Exit Sub
    End If
    
    'У������ҽ��
    If mblnAutoVerify Then
        strTmp = str����IDs
        If CheckSpecialAdvice(0, 0, strTmp) Then
             '����ȼ��򲡿�У�Ժ���Ҫˢ�²����嵥
            mblnRefresh = True
            If frmAdviceOperate.ShowMe(Me, mMainPrivs, 3, mlng����ID, mlng��ҳID, mlng����ID, _
                0, True, True, 0, , , mblnOnePati, strTmp, True, , , mlng����ID, True, mclsMipModule) = False Then
                Exit Sub
            End If
            '�����Զ�У�Դ��ں������Զ�ѡ��Ϊ:ȫ��
            opt״̬(2).value = True
        End If
        strTmp = ""
        '����Ǳ���ҽ������Ĭ��ѡ��ȫ��
        If lngRefModld = 1 Then
            If Not opt״̬(2).value Then opt״̬(2).value = True
        End If
    End If
    
    strSel = Mid(strSel, 2)
    strUnSel = Mid(strUnSel, 2)
    If strSel = "" Or (UBound(Split(strSel, ",")) = 0 And Val(strSel) = mlng����ID) Then
        strSel = ""
    Else
        If strUnSel = "" Then
            strSel = cboUnit.ItemData(cboUnit.ListIndex) & ":ALL"
        ElseIf UBound(Split(strSel, ",")) > UBound(Split(strUnSel, ",")) Then
            strSel = cboUnit.ItemData(cboUnit.ListIndex) & ":-" & strUnSel
        Else
            strSel = cboUnit.ItemData(cboUnit.ListIndex) & ":" & strSel
        End If
    End If
    
    If chkҩƷ��Ч(0).value = 1 Or chkҩƷ��Ч(1).value = 1 Then
        '��ҩ;��
        str��ҩIDs = ""
        For i = 1 To lvwWay.ListItems.Count
            If lvwWay.ListItems(i).Checked Then
                str��ҩIDs = str��ҩIDs & "," & Mid(lvwWay.ListItems(i).Key, 2)
            End If
        Next
        str��ҩIDs = Mid(str��ҩIDs, 2)
        If str��ҩIDs = "" Then
            MsgBox "������ѡ��һ�ָ�ҩ;����", vbInformation, gstrSysName
            If lvwWay.Visible And lvwWay.Enabled Then lvwWay.SetFocus
            Exit Sub
        End If
        If UBound(Split(str��ҩIDs, ",")) + 1 = lvwWay.ListItems.Count Then
            str��ҩIDs = ""
        End If
    Else
        '�������
        str���s = ""
        For i = 0 To lstClass.ListCount - 1
            If lstClass.Selected(i) Then
                str���s = str���s & ",'" & Chr(lstClass.ItemData(i)) & "'"
            End If
        Next
        str���s = Mid(str���s, 2)
        If str���s = "" Then
            MsgBox "������ѡ��һ���������", vbInformation, gstrSysName
            If lstClass.Visible And lstClass.Enabled Then lstClass.SetFocus
            Exit Sub
        End If
        If UBound(Split(str���s, ",")) + 1 = lstClass.ListCount Then
            str���s = ""
        End If
    End If
    
    gbln�Ӱ�Ӽ� = chk�Ӱ�Ӽ�.value = 1
    
    '��ȡ����
    '--------------------------------------------------------------------------------
    Call InitPriceRecordset '�Ƽ۹�ϵ��
    
    Call LoadAdviceSend(str����IDs, str��ҳIDs, mstrEnd, str��ҩIDs, str���s, mlngRefModld, str���˿���IDs)
    mblnChangeIF = False
    
    
    '��������ֵ
    '--------------------------------------------------------------------------------
    If Me.Visible Then  '�Զ���ȡ(ҽ������)ʱ�����б������

        Call zlDatabase.SetPara("���ͽ���ʱ��", cboTime.ListIndex, glngSys, pסԺҽ������)
        If cboTime.ListIndex = cboTime.ListCount - 1 Then
            Call zlDatabase.SetPara("���ͽ���ʱ��", Format(dtpEnd.value, "HH:mm:ss"), glngSys, pסԺҽ������)
            Call zlDatabase.SetPara("����ʱ����", Int(CDate(Format(dtpEnd.value, "yyyy-MM-dd")) - CDate(Format(mdatCurr, "yyyy-MM-dd"))), glngSys, pסԺҽ������)
        End If
        
        '������ģʽ������
        If Not mblnOnePati Then
            Call zlDatabase.SetPara("���Ͳ���", strSel, glngSys, pסԺҽ������)
        End If
        
        'ͨ��ҽ�����Ѵ�ʱ��ֻ��ʾ�¿��ģ�������״̬������Ч
        If Not (mblnAutoVerify And mblnOnePati) Then
            'ҽ����Ч
            If chkҩƷ��Ч(0).value = 1 And chkҩƷ��Ч(1).value = 1 Then
                strTmp = "2"
            ElseIf chkҩƷ��Ч(0).value = 1 Then
                strTmp = "0"
            ElseIf chkҩƷ��Ч(1).value = 1 Then
                strTmp = "1"
            Else
                strTmp = "-1"
            End If
            Call zlDatabase.SetPara("ҩ������ҽ����Ч", strTmp, glngSys, pסԺҽ������)
            
            If chk������Ч(0).value = 1 And chk������Ч(1).value = 1 Then
                strTmp = "2"
            ElseIf chk������Ч(0).value = 1 Then
                strTmp = "0"
            ElseIf chk������Ч(1).value = 1 Then
                strTmp = "1"
            Else
                strTmp = "-1"
            End If
            Call zlDatabase.SetPara("��������ҽ����Ч", strTmp, glngSys, pסԺҽ������)
        End If
        
        
        If chkҩƷ��Ч(0).Visible And chkҩƷ��Ч(0).Enabled Or (chkҩƷ��Ч(1).Visible And chkҩƷ��Ч(1).Enabled) Then
            Call zlDatabase.SetPara("ҩ�����͸�ҩ;��", str��ҩIDs, glngSys, pסԺҽ������)
                        
            'ҩ���û�
            mrsҩ��.Filter = 0
            strTmp = ""
            For i = 1 To mrsҩ��.RecordCount
                strTmp = strTmp & "," & mrsҩ��!ID & "-" & mrsҩ��!��ID
                mrsҩ��.MoveNext
            Next
            strTmp = Mid(strTmp, 2)
            Call zlDatabase.SetPara("ҩ������ҩ���û�", strTmp, glngSys, pסԺҽ������)
        End If
        
        If chk������Ч(0).Visible And chk������Ч(0).Enabled Or (chk������Ч(1).Visible And chk������Ч(1).Enabled) Then
            If chk������Ч(0).value = 1 Or chk������Ч(1).value = 1 Then
                Call zlDatabase.SetPara("���������������", Replace(str���s, "'", "''"), glngSys, pסԺҽ������)
            End If
        End If
        
        If chkSendNow.Enabled Then Call zlDatabase.SetPara("ҩƷ�����ӵ��쿪ʼ����", chkSendNow.value, glngSys, pסԺҽ������)
    End If
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tkpMain.Width + X < 3000 Or vsAdvice.Width - X < 3000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        tkpMain.Width = tkpMain.Width + X
        
        fraInfo.Left = fraInfo.Left + X
        fraInfo.Width = fraInfo.Width - X
        
        vsAdvice.Left = vsAdvice.Left + X
        vsAdvice.Width = vsAdvice.Width - X
        
        vsPrice.Left = vsPrice.Left + X
        vsPrice.Width = vsPrice.Width - X
        
        fraUD.Left = fraUD.Left + X
        fraUD.Width = fraUD.Width - X
        
        Me.Refresh
    End If
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsAdvice.Height + Y < 1000 Or vsPrice.Height - Y < 500 Then Exit Sub
        fraUD.Top = fraUD.Top + Y
        vsAdvice.Height = vsAdvice.Height + Y
        vsPrice.Top = vsPrice.Top + Y
        vsPrice.Height = vsPrice.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub RowSelectSame(ByVal lngRow As Long, ByVal lngCol As Long, _
    Optional rsSQL As ADODB.Recordset, Optional rsTotal As ADODB.Recordset, Optional rsUpload As ADODB.Recordset)
'���ܣ����ݿɼ��е�ѡ��״̬,�����ҽ��һ��ѡ��
    Dim i As Long
    
    With vsAdvice
        If lngCol = COL_ѡ�� Then
            For i = lngRow + 1 To .Rows - 1
                If IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            For i = lngRow - 1 To .FixedRows Step -1
                If IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            
            'һ��������ŵĻ���ҽ��
            If Val(.TextMatrix(lngRow, COL_�������)) <> 0 And .TextMatrix(lngRow, COL_�������) = "Z" Then
                For i = lngRow - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(i, COL_�������)) = Val(.TextMatrix(lngRow, COL_�������)) Then
                        .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                        Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                    Else
                        Exit For
                    End If
                Next
                For i = lngRow + 1 To .Rows - 1
                    If Val(.TextMatrix(i, COL_�������)) = Val(.TextMatrix(lngRow, COL_�������)) Then
                        .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                        Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                    Else
                        Exit For
                    End If
                Next
            End If
            
            'ȡ��ѡ��ʱ
            If Not (.Cell(flexcpData, lngRow, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, lngRow, COL_ѡ��) Is Nothing) Then
                i = IIF(Val(.TextMatrix(lngRow, COL_���ID)) = 0, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_���ID)))
                '1.�����Ӧ�ķ��ü����ͼ�¼��д
                If Not rsSQL Is Nothing Then
                    rsSQL.Filter = "ҽ��ID=" & i
                    Do While Not rsSQL.EOF
                        rsSQL.Delete
                        rsSQL.Update
                        rsSQL.MoveNext
                    Loop
                    rsSQL.Filter = 0 '��ΪҪʹ��BookMark����˻ָ�
                End If
                '2.�����Ӧ�ķ��ͼƼ������ۼ�
                If Not rsTotal Is Nothing Then
                    rsTotal.Filter = "ҽ��ID=" & i
                    Do While Not rsTotal.EOF
                        rsTotal.Delete
                        rsTotal.Update
                        rsTotal.MoveNext
                    Loop
                End If
                '3.�����Ӧ��ҽ���ϴ����ݺ�
                If Not rsUpload Is Nothing Then
                    rsUpload.Filter = "ҽ��ID=" & i
                    Do While Not rsUpload.EOF
                        rsUpload.Delete
                        rsUpload.Update
                        rsUpload.MoveNext
                    Loop
                End If
            End If
        End If
    End With
End Sub

Private Function GetVisibleRow(ByVal lngRow As Long, Optional ByVal blnFirst As Boolean) As Long
'���ܣ�����ָ��ҽ���У����ظ�ҽ���пɼ�����
    Dim lng��ID As Long, i As Long
    
    GetVisibleRow = lngRow
    
    With vsAdvice
        If Not .RowHidden(lngRow) Then Exit Function
        
        'һ����ҩ�Ķ�λ����һҩƷ��
        If blnFirst Then
            If .TextMatrix(lngRow, COL_�������) = "E" And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 _
                And Val(.TextMatrix(lngRow, COL_���ID)) = 0 And Val(.TextMatrix(lngRow, COL_ID)) = Val(.TextMatrix(lngRow - 1, COL_���ID)) Then
                i = .FindRow(.TextMatrix(lngRow, COL_ID), , COL_���ID)
                If i <> -1 Then GetVisibleRow = i: Exit Function
            End If
        End If
        
        lng��ID = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            If lng��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            If lng��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
    End With
End Function

Private Function RowIn������(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ����ڼ�������е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then
            '�ɼ�������
            If .TextMatrix(lngRow - 1, COL_�������) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) Then
                RowIn������ = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, COL_�������) = "C" And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            '������Ŀ��
            RowIn������ = True: Exit Function
        End If
    End With
End Function

Private Function RowIn�䷽��(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ�������ҩ�䷽�е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "E" Then
            If Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then
                '�÷���
                If Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) _
                    And .TextMatrix(lngRow - 1, COL_�������) = "E" Then
                    RowIn�䷽�� = True: Exit Function
                End If
            Else
                '�巨��
                If .TextMatrix(lngRow - 1, COL_�������) = "7" _
                    And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    RowIn�䷽�� = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, COL_�������) = "7" And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            '��ҩ��
            RowIn�䷽�� = True: Exit Function
        End If
    End With
End Function

Private Function GetComboList(ByVal lngRow As Long) As String
'���ܣ����ݵ�ǰҽ���л�ȡ��ѡ��ļƼ�ҽ������
'������lngRow=�ɼ���(ҩ�ƻ��ҩ)
'˵����ע�������Ǹ��ݾ���ҽ����ȡ
    Dim strCombo As String
    Dim strTmp As String, lngTmp As Long
    Dim i As Long, j As Long
    
    With vsAdvice
        If Val(.Cell(flexcpData, lngRow, COL_�������)) = 3 Then
            '��ҩ�÷�����ҩ�÷�,��ҩ�巨
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_���ID)
            For i = lngTmp To lngRow
                If InStr(",2,3,", Val(.Cell(flexcpData, i, COL_�������))) > 0 Then
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                If NVL(mrsPrice!�̶�, 0) = 0 Then
                                    If Val(.Cell(flexcpData, i, COL_�������)) = 2 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�巨-" & .Cell(flexcpData, i, col_ҽ������)
                                    ElseIf Val(.Cell(flexcpData, i, COL_�������)) = 3 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�÷�-" & .Cell(flexcpData, i, col_ҽ������)
                                    End If
                                    If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                        strCombo = strCombo & "|#" & strTmp
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Next
                        Else
                            If Val(.Cell(flexcpData, i, COL_�������)) = 2 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�巨-" & .Cell(flexcpData, i, col_ҽ������)
                            ElseIf Val(.Cell(flexcpData, i, COL_�������)) = 3 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�÷�-" & .Cell(flexcpData, i, col_ҽ������)
                            End If
                            If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                strCombo = strCombo & "|#" & strTmp
                            End If
                        End If
                    End If
                End If
            Next
        ElseIf .TextMatrix(lngRow, COL_�������) = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 _
            And .TextMatrix(lngRow - 1, COL_�������) = "C" And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
            '�ɼ�������
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_���ID)
            For i = lngTmp To lngRow
                If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                    If Not mrsPrice.EOF Then
                        For j = 1 To mrsPrice.RecordCount
                            If NVL(mrsPrice!�̶�, 0) = 0 Then
                                If .TextMatrix(i, COL_�������) = "C" Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";������Ŀ-" & .Cell(flexcpData, i, col_ҽ������)
                                ElseIf .TextMatrix(i, COL_�������) = "E" Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";�ɼ�����-" & .Cell(flexcpData, i, col_ҽ������)
                                End If
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                            mrsPrice.MoveNext
                        Next
                    Else
                        If .TextMatrix(i, COL_�������) = "C" Then
                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";������Ŀ-" & .Cell(flexcpData, i, col_ҽ������)
                        ElseIf .TextMatrix(i, COL_�������) = "E" Then
                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";�ɼ�����-" & .Cell(flexcpData, i, col_ҽ������)
                        End If
                        If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                            strCombo = strCombo & "|#" & strTmp
                        End If
                    End If
                End If
            Next
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            '���г�ҩ����ҩ;��
            If Val(.TextMatrix(lngRow - 1, COL_���ID)) <> Val(.TextMatrix(lngRow, COL_���ID)) Then
                lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_���ID))), lngRow + 1, COL_ID)
                If Val(.TextMatrix(lngTmp, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngTmp, COL_ִ������ID))) = 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngTmp, COL_ID))
                    If Not mrsPrice.EOF Then
                        For j = 1 To mrsPrice.RecordCount
                            If NVL(mrsPrice!�̶�, 0) = 0 Then
                                strCombo = "|#" & Val(.TextMatrix(lngTmp, COL_ID)) & ";��ҩ;��-" & .Cell(flexcpData, lngTmp, col_ҽ������)
                                Exit For
                            End If
                            mrsPrice.MoveNext
                        Next
                    Else
                        strCombo = "|#" & Val(.TextMatrix(lngTmp, COL_ID)) & ";��ҩ;��-" & .Cell(flexcpData, lngTmp, col_ҽ������)
                    End If
                End If
            End If
        Else
            'һ���������飬����Ѫҽ���������ҽ��
            For i = lngRow To .Rows - 1
                If i = lngRow Or Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                If NVL(mrsPrice!�̶�, 0) = 0 Then
                                    If .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, col_ҽ������)
                                    ElseIf .TextMatrix(i, COL_�������) = "G" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, col_ҽ������)
                                    ElseIf .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��鲿λ-" & .TextMatrix(i, COL_�걾��λ) & "(" & .TextMatrix(i, COL_��鷽��) & ")"
                                    ElseIf .TextMatrix(i, COL_�������) = "E" And .TextMatrix(lngRow, COL_�������) = "K" Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";��Ѫ;��-" & .Cell(flexcpData, i, col_ҽ������)
                                    Else
                                        If mrsPrice!�������� <> 0 Then
                                            '���շ��ã�Ŀǰ�������Ĵ��Ժ����м���
                                            lngTmp = -1 * Val(mrsPrice!�������� & Val(.TextMatrix(i, COL_ID)))
                                            strTmp = lngTmp & ";" & .Cell(flexcpData, i, COL_�������) & "ҽ��-" & .Cell(flexcpData, i, col_ҽ������) & _
                                                "(" & decode(Val(.TextMatrix(i, COL_ִ�б��)), 1, "����", 2, "����", "") & "����)"
                                        Else
                                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";" & .Cell(flexcpData, i, COL_�������) & "ҽ��-" & .Cell(flexcpData, i, col_ҽ������)
                                        End If
                                    End If
                                    If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                        strCombo = strCombo & "|#" & strTmp
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Next
                        Else
                            'δ���üƼ۵ģ�����ѡ����ӼƼ���Ŀ
                            If .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, col_ҽ������)
                            ElseIf .TextMatrix(i, COL_�������) = "G" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, col_ҽ������)
                            ElseIf .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��鲿λ-" & .TextMatrix(i, COL_�걾��λ) & "(" & .TextMatrix(i, COL_��鷽��) & ")"
                            ElseIf .TextMatrix(i, COL_�������) = "E" And .TextMatrix(lngRow, COL_�������) = "K" Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";��Ѫ;��-" & .Cell(flexcpData, i, col_ҽ������)
                            Else
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";" & .Cell(flexcpData, i, COL_�������) & "ҽ��-" & .Cell(flexcpData, i, col_ҽ������)
                            End If
                            If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                strCombo = strCombo & "|#" & strTmp
                            End If
                            
                            '���շ��ã�Ŀǰ�������Ĵ��Ի����м���
                            If .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) = 0 _
                                And (Val(.TextMatrix(i, COL_ִ�б��)) = 1 Or Val(.TextMatrix(i, COL_ִ�б��)) = 2) Then
                                lngTmp = -1 * Val(1 & Val(.TextMatrix(i, COL_ID)))
                                strTmp = lngTmp & ";" & .Cell(flexcpData, i, COL_�������) & "ҽ��-" & .Cell(flexcpData, i, col_ҽ������) & _
                                    "(" & decode(Val(.TextMatrix(i, COL_ִ�б��)), 1, "����", 2, "����", "") & "����)"
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    GetComboList = Mid(strCombo, 2)
End Function

Private Function ShowAdvicePrice(ByVal lngRow As Long) As Boolean
'���ܣ�����ҽ���Ƽ۹�ϵ�����㲢��ʾָ��ҽ���ķ���(����ҽ�������ܶ���)
'������lngRow=�ɼ���(ҩ�ƻ��ҩ)
    Dim rsTmp As New ADODB.Recordset
    Dim rsExeDays As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngTopRow As Long, lngLeftCol As Long
    Dim lngPreRow As Long, lngPreCol As Long
    Dim blnFirst As Boolean, str�Ƽ�ҽ�� As String
    Dim str��λ As String, dbl���� As Double, int���� As Integer
    Dim bln�������� As Boolean, strCombo As String, str�к� As String, str�ֽ�ʱ�� As String
    Dim dbl���� As Double, curӦ�� As Currency, curʵ�� As Currency
    Dim dbl��ǰ���� As Double, dbl��ǰӦ�� As Double, cur��ǰӦ�� As Currency, cur��ǰʵ�� As Currency
    Dim lng�к� As Long, cur�ϼ� As Currency, bln���� As Boolean
    
    Dim rsMain As New ADODB.Recordset
    Dim rsClone As New ADODB.Recordset
    Dim strHaveSub As String, strNoneSub As String
    Dim strPriceType As String
        
    On Error GoTo errH
    
    '���ڻ��ܼ����ۿ۵���ʱ��¼��
    rsMain.Fields.Append "ҽ���к�", adBigInt
    rsMain.Fields.Append "��������", adInteger
    rsMain.Fields.Append "�����к�", adBigInt
    rsMain.Fields.Append "������ID", adBigInt
    rsMain.Fields.Append "ҽ���ϼ�", adCurrency, , adFldIsNullable
    rsMain.CursorLocation = adUseClient
    rsMain.LockType = adLockOptimistic
    rsMain.CursorType = adOpenStatic
    rsMain.Open
    
    With vsAdvice
        blnFirst = True
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnFirst = False 'һ����ҩ���Ƿ��һҩƷ��
            End If
        End If
        
        If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            If blnFirst Then
                mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                    " Or ҽ��ID=" & Val(.TextMatrix(lngRow, COL_���ID))
            Else
                mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngRow, COL_ID))
            End If
        Else
            mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                " Or ���ID=" & Val(.TextMatrix(lngRow, COL_ID))
        End If
        
        For i = 1 To mrsPrice.RecordCount
            '�Ƽ�ҽ��
            bln�������� = False
            lng�к� = .FindRow(CStr(mrsPrice!ҽ��ID), , COL_ID)
            If .TextMatrix(lng�к�, COL_�������) = "4" Then
                str�Ƽ�ҽ�� = "��������-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf InStr(",5,6,7", .TextMatrix(lng�к�, COL_�������)) > 0 Then
                str�Ƽ�ҽ�� = "ҩƷҽ��-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf Val(.Cell(flexcpData, lng�к�, COL_�������)) = 1 Then
                str�Ƽ�ҽ�� = "��ҩ;��-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf Val(.Cell(flexcpData, lng�к�, COL_�������)) = 2 Then
                str�Ƽ�ҽ�� = "��ҩ�巨-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf Val(.Cell(flexcpData, lng�к�, COL_�������)) = 3 Then
                str�Ƽ�ҽ�� = "��ҩ�÷�-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "E" And Val(.TextMatrix(lng�к�, COL_���ID)) = 0 _
                And .TextMatrix(lng�к� - 1, COL_�������) = "C" And Val(.TextMatrix(lng�к� - 1, COL_���ID)) = Val(.TextMatrix(lng�к�, COL_ID)) Then
                str�Ƽ�ҽ�� = "�ɼ�����-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "E" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 _
                And .TextMatrix(lng�к� - 1, COL_�������) = "K" And Val(.TextMatrix(lng�к� - 1, COL_ID)) = Val(.TextMatrix(lng�к�, COL_���ID)) Then
                str�Ƽ�ҽ�� = "��Ѫ;��-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "C" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                str�Ƽ�ҽ�� = "������Ŀ-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "F" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                bln�������� = True
                str�Ƽ�ҽ�� = "��������-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "G" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                str�Ƽ�ҽ�� = "��������-" & .Cell(flexcpData, lng�к�, col_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "D" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                str�Ƽ�ҽ�� = "��鲿λ-" & .TextMatrix(lng�к�, COL_�걾��λ) & "(" & .TextMatrix(lng�к�, COL_��鷽��) & ")"
            Else
                If NVL(mrsPrice!��������, 0) = 1 Then
                    '���Ի����м��շ���
                    str�Ƽ�ҽ�� = .Cell(flexcpData, lng�к�, COL_�������) & "ҽ��-" & .Cell(flexcpData, lng�к�, col_ҽ������) & _
                        "(" & decode(Val(.TextMatrix(lng�к�, COL_ִ�б��)), 1, "����", 2, "����", "") & "����)"
                Else
                    str�Ƽ�ҽ�� = .Cell(flexcpData, lng�к�, COL_�������) & "ҽ��-" & .Cell(flexcpData, lng�к�, col_ҽ������)
                End If
            End If
            str�Ƽ�ҽ�� = Replace(str�Ƽ�ҽ��, "'", "''")
            
            '����:ҩƷ��סԺ��λ������,��������������
            int���� = 1
            If InStr(",5,6,", .TextMatrix(lng�к�, COL_�������)) > 0 Then
                dbl���� = Val(.TextMatrix(lng�к�, COL_����))
            ElseIf .TextMatrix(lng�к�, COL_�������) = "7" Then
                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                int���� = Val(.TextMatrix(lng�к�, COL_����))
                If Val(.TextMatrix(lng�к�, COL_�ɷ����)) = 0 Then
                    dbl���� = Val(.TextMatrix(lng�к�, COL_����)) / Val(.TextMatrix(lng�к�, COL_����ϵ��)) / Val(.TextMatrix(lng�к�, COL_סԺ��װ))
                Else
                    dbl���� = IntEx(Val(.TextMatrix(lng�к�, COL_����)) / Val(.TextMatrix(lng�к�, COL_����ϵ��)) / Val(.TextMatrix(lng�к�, COL_סԺ��װ)))
                End If
            Else
                If InStr(",3,4,5,6,", Val("" & mrsPrice!�շѷ�ʽ)) > 0 Then 'һ��ֻ��һ�ε�
                     '�ֽ�ʱ��
                    If .TextMatrix(lng�к�, COL_�ֽ�ʱ��) <> "" Then
                        str�ֽ�ʱ�� = .TextMatrix(lng�к�, COL_�ֽ�ʱ��)
                    Else
                        str�ֽ�ʱ�� = .Cell(flexcpData, lng�к�, COL_�ֽ�ʱ��)    '��ʼִ��ʱ��
                    End If
                    
                    Set rsExeDays = GetExecDays(str�ֽ�ʱ��)
                    dbl���� = rsExeDays.RecordCount
                ElseIf InStr(",1,2,", Val("" & mrsPrice!�շѷ�ʽ)) > 0 Then 'һ�η���ֻ��һ��
                    dbl���� = 1
                Else
                    dbl���� = Val(.TextMatrix(lng�к�, COL_����))
                End If
            End If
            dbl���� = Format(dbl���� * NVL(mrsPrice!����, 0), "0.00000")
                        
            '���SQL
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                " Select " & i & " as ���," & mrsPrice!ҽ��ID & " as ҽ��ID,ID," & _
                NVL(mrsPrice!�̶�, 0) & " as �̶�,'" & str�Ƽ�ҽ�� & "' as �Ƽ�ҽ��,���,����,����,���," & _
                "���㵥λ as ��λ," & NVL(mrsPrice!����, 0) & " as �Ƽ�����," & int���� & " as ����," & dbl���� & " as ����," & _
                Format(NVL(mrsPrice!����, 0), gstrDecPrice) & " as ����,��������," & lng�к� & " as �к�," & _
                " �Ƿ���,�Ӱ�Ӽ�," & IIF(bln��������, 1, 0) & " as ��������," & mrsPrice!���� & " as ����," & _
                NVL(mrsPrice!ִ�п���ID, 0) & " as ִ�п���ID,���ηѱ�," & mrsPrice!�������� & " as ��������," & _
                mrsPrice!�շѷ�ʽ & " as �շѷ�ʽ From �շ���ĿĿ¼ Where ID=" & mrsPrice!�շ�ϸĿID
            mrsPrice.MoveNext
        Next
    End With
    
    With vsPrice
        lngPreRow = .Row: lngPreCol = .Col
        lngTopRow = .TopRow: lngLeftCol = .LeftCol
        .Editable = flexEDNone
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        '��Ҫ�Ƽ۵�ҽ��ѡ��
        '���ݴ�����ҽ��ȡ�ɼƼ�ҽ��(���ܴ�mrsPriceȡ,��Ϊ�������շѹ�ϵ����ɾ��,����Ҳ�����ڼƼ���ȫ��ɾ��)
        strCombo = GetComboList(lngRow)
        If strCombo <> "" Then
            .ColData(COLP_�Ƽ�ҽ��) = strCombo
            .Editable = flexEDKbdMouse '����ѡ������Ա༭
        Else
            .ColData(COLP_�Ƽ�ҽ��) = ""
        End If
        If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(lngRow, COL_����ID)), Val(vsAdvice.TextMatrix(lngRow, COL_��ҳID)), "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
        '��ʾ���еļƼ���Ŀ
        If strSQL <> "" Then
            strSQL = "Select A.�к�,A.ID AS �շ�ϸĿID,A.�̶�,A.����,A.�Ƽ�ҽ��,A.���,C.���� as �������,A.ִ�п���ID,G.���� as ִ�п���," & _
                " Nvl(E.����,A.����)||Decode(A.����,NULL,NULL,'('||A.����||')')||Decode(A.���,NULL,NULL,' '||A.���) as ����," & _
                " A.��λ,A.�Ƽ�����,A.����,A.����,D.סԺ��װ,D.סԺ��λ,Decode(A.�Ƿ���,1,A.����,B.�ּ�) as ����,F.��������," & _
                " A.��������,A.�շѷ�ʽ,A.��������,A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,B.ԭ��,B.�ּ�,A.��������,B.�����շ���,B.������ĿID" & _
                " From (" & strSQL & ") A,�շѼ�Ŀ B,�շ���Ŀ��� C,ҩƷ��� D,�շ���Ŀ���� E,�������� F,���ű� G" & _
                " Where A.ID=B.�շ�ϸĿID And A.���=C.���� And A.ID=D.ҩƷID(+)" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "1", "2", "3") & _
                " And A.ID=F.����ID(+) And A.ִ�п���ID=G.ID(+)" & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                " And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIF(gbytҩƷ������ʾ = 0, 1, 3) & _
                " Order by A.���"
                '��Ϊ������ǵ��ñ�����ˢ��,Ҫ���ֶ�̬��¼���м�¼˳��
                'Ҫ��֤��������ǰ��,LoadAdvicePriceʱ������������ǰ�棬���ұ༭��ֻ���ܼ��˴���
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�) 'û��
            
            If Not rsTmp.EOF And gbln��������ۿ� Then
                Set rsClone = rsTmp.Clone
            End If
            
            For i = 1 To rsTmp.RecordCount
                If str�к� <> rsTmp!�к� & "_" & rsTmp!�������� & "_" & rsTmp!�շ�ϸĿID Then
                    If str�к� <> "" Then
                        If Not (Val(.TextMatrix(.Rows - 1, COLP_���)) = 1 And dbl���� = 0) Then
                            .TextMatrix(.Rows - 1, COLP_����) = Format(dbl����, gstrDecPrice)
                            .Cell(flexcpData, .Rows - 1, COLP_����) = .TextMatrix(.Rows - 1, COLP_����) '��¼���ڻָ�����
                            .TextMatrix(.Rows - 1, COLP_Ӧ�ս��) = Format(curӦ��, gstrDec)
                            .TextMatrix(.Rows - 1, COLP_ʵ�ս��) = Format(curʵ��, gstrDec)
                        End If
                        cur�ϼ� = cur�ϼ� + Format(curʵ��, gstrDec)
                    End If
                    str�к� = rsTmp!�к� & "_" & rsTmp!�������� & "_" & rsTmp!�շ�ϸĿID
                    dbl���� = 0: curӦ�� = 0: curʵ�� = 0
                    .Rows = .Rows + 1
                    
                    '��ʶ�̶�����Ϊ��ɫ
                    If rsTmp!�̶� <> 0 Then
                        .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HE0E0E0
                    End If

                    .TextMatrix(.Rows - 1, COLP_�к�) = rsTmp!�к�
                    .TextMatrix(.Rows - 1, COLP_�շ�ϸĿID) = rsTmp!�շ�ϸĿID
                    .TextMatrix(.Rows - 1, COLP_�̶�) = rsTmp!�̶�
                    .TextMatrix(.Rows - 1, COLP_�Ƽ�ҽ��) = rsTmp!�Ƽ�ҽ��
                    .TextMatrix(.Rows - 1, COLP_��������) = rsTmp!��������
                    .TextMatrix(.Rows - 1, COLP_�շѷ�ʽ) = getChargeMode(Val(NVL(rsTmp!�շѷ�ʽ, 0)))
                        .Cell(flexcpData, .Rows - 1, COLP_�շѷ�ʽ) = Val(NVL(rsTmp!�շѷ�ʽ, 0))
                    .TextMatrix(.Rows - 1, COLP_���) = rsTmp!�������
                    .TextMatrix(.Rows - 1, COLP_�շ����) = rsTmp!���
                    .TextMatrix(.Rows - 1, COLP_�շ���Ŀ) = rsTmp!����
                    .TextMatrix(.Rows - 1, COLP_�Ƽ�����) = NVL(rsTmp!�Ƽ�����, 0) '�������
                    
                    int���� = NVL(rsTmp!����, 1)
                    
                    dbl���� = NVL(rsTmp!����, 0) '�ۼ��������ں��水�ɱ����ۼ���
                    If InStr(",5,6,7,", rsTmp!���) > 0 Then 'סԺ��װ
                        .TextMatrix(.Rows - 1, COLP_��λ) = NVL(rsTmp!סԺ��λ)
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                            .TextMatrix(.Rows - 1, COLP_����) = FormatEx(NVL(rsTmp!����, 0), 5)
                            dbl���� = dbl���� * NVL(rsTmp!סԺ��װ, 1)
                        Else
                            '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                            '��ҩ��ҩƷ�Ƽ�:��Ϊ����Ԥ�����ۼ�����,���ת��Ϊҩ����λ��ʾʱ���������㴦��
                            .TextMatrix(.Rows - 1, COLP_����) = FormatEx(NVL(rsTmp!����, 0) / NVL(rsTmp!סԺ��װ, 1), 5)
                        End If
                        
                        If rsTmp!��� = "7" Then
                            .TextMatrix(.Rows - 1, COLP_����) = int����
                            bln���� = True
                        End If
                    Else
                        .TextMatrix(.Rows - 1, COLP_��λ) = NVL(rsTmp!��λ)
                        .TextMatrix(.Rows - 1, COLP_����) = FormatEx(NVL(rsTmp!����, 0), 5)
                    End If
                    
                    .TextMatrix(.Rows - 1, COLP_ִ�п���) = NVL(rsTmp!ִ�п���)
                    .TextMatrix(.Rows - 1, COLP_ִ�п���ID) = NVL(rsTmp!ִ�п���ID, 0)
                    
                    '��ʾҽ����������
                    If Val(rsTmp!�շ�ϸĿID & "") <> 0 Then
                        strPriceType = GetPriceType(Val(vsAdvice.TextMatrix(lngRow, COL_����ID)), Val(rsTmp!�շ�ϸĿID & ""), Val(vsAdvice.TextMatrix(lngRow, COL_����)), False)
                    End If
                    '��������
                    If strPriceType = "" Then
                        .TextMatrix(.Rows - 1, COLP_��������) = NVL(rsTmp!��������)
                    Else
                        .TextMatrix(.Rows - 1, COLP_��������) = strPriceType
                    End If
                    
                    
                    .TextMatrix(.Rows - 1, COLP_����) = IIF(NVL(rsTmp!����, 0) = 0, "", "��")
                    .TextMatrix(.Rows - 1, COLP_��������) = NVL(rsTmp!��������, 0)
                    
                    '��¼��������ָ�
                    .Cell(flexcpData, .Rows - 1, COLP_�Ƽ�ҽ��) = .TextMatrix(.Rows - 1, COLP_�Ƽ�ҽ��)
                    .Cell(flexcpData, .Rows - 1, COLP_�շ���Ŀ) = .TextMatrix(.Rows - 1, COLP_�շ���Ŀ)
                    .Cell(flexcpData, .Rows - 1, COLP_�Ƽ�����) = .TextMatrix(.Rows - 1, COLP_�Ƽ�����)
                    .Cell(flexcpData, .Rows - 1, COLP_ִ�п���) = .TextMatrix(.Rows - 1, COLP_ִ�п���)
                    
                    '��¼�����������Ϣ���Ա����
                    If gbln��������ۿ� And rsTmp!���� = 0 Then
                        If InStr(strHaveSub & ",", "," & rsTmp!�к� & "_" & rsTmp!�������� & ",") = 0 _
                            And InStr(strNoneSub & ",", "," & rsTmp!�к� & "_" & rsTmp!�������� & ",") = 0 Then
                            rsClone.Filter = "�к�=" & rsTmp!�к� & " And ��������=" & rsTmp!�������� & " And ����=1"
                            If Not rsClone.EOF Then
                                rsMain.AddNew
                                rsMain!ҽ���к� = rsTmp!�к�
                                rsMain!�������� = rsTmp!��������
                                rsMain!�����к� = .Rows - 1
                                rsMain!������ID = rsTmp!������ĿID
                                rsMain.Update
                                strHaveSub = strHaveSub & "," & rsTmp!�к� & "_" & rsTmp!��������
                            Else
                                strNoneSub = strNoneSub & "," & rsTmp!�к� & "_" & rsTmp!��������
                            End If
                        End If
                    End If
                    
                    '��ҩƷ������ҽ����ҩƷ�͸������ļƼۣ���ʹ�̶�Ҳ�����޸�ִ�п���
                    If InStr(",5,6,7,", rsTmp!���) > 0 _
                        Or rsTmp!��� = "4" And NVL(rsTmp!��������, 0) = 1 Then
                        .Editable = flexEDKbdMouse
                    End If
                End If
                
                '���ۼ��㴦��
                If InStr(",5,6,7,", rsTmp!���) > 0 Then
                    If NVL(rsTmp!�Ƿ���, 0) = 0 Then
                        dbl��ǰ���� = NVL(rsTmp!����, 0)
                    Else
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                            dbl��ǰ���� = CalcDrugPrice(rsTmp!�շ�ϸĿID, NVL(rsTmp!ִ�п���ID, 0), Format(int���� * NVL(rsTmp!����, 0) * NVL(rsTmp!סԺ��װ, 1), gstrDecPrice), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                        Else
                            dbl��ǰ���� = CalcDrugPrice(rsTmp!�շ�ϸĿID, NVL(rsTmp!ִ�п���ID, 0), Format(int���� * NVL(rsTmp!����, 0), gstrDecPrice), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                        End If
                    End If
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                        dbl��ǰ���� = dbl��ǰ���� * NVL(rsTmp!סԺ��װ, 1)
                        dbl��ǰӦ�� = Format(int���� * NVL(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                    Else
                        dbl��ǰӦ�� = Format(int���� * NVL(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                        dbl��ǰ���� = dbl��ǰ���� * NVL(rsTmp!סԺ��װ, 1)
                    End If
                ElseIf rsTmp!��� = "4" And NVL(rsTmp!��������, 0) = 1 And NVL(rsTmp!�Ƿ���, 0) = 1 Then
                    '�������õ�ʱ�����ĺ�ҩƷһ������
                    dbl��ǰ���� = CalcDrugPrice(rsTmp!�շ�ϸĿID, NVL(rsTmp!ִ�п���ID, 0), Format(NVL(rsTmp!����, 0), "0.00000"), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                    dbl��ǰӦ�� = Format(NVL(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                Else
                    dbl��ǰ���� = NVL(rsTmp!����, 0) '�������Ϊ��������û������
                    dbl��ǰӦ�� = Format(NVL(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                    If NVL(rsTmp!�Ƿ���, 0) = 1 Then '��¼��ҩ��۷�Χ
                        .TextMatrix(.Rows - 1, COLP_���) = 1
                        .Cell(flexcpData, .Rows - 1, COLP_Ӧ�ս��) = CCur(NVL(rsTmp!ԭ��, 0))
                        .Cell(flexcpData, .Rows - 1, COLP_ʵ�ս��) = CCur(NVL(rsTmp!�ּ�, 0))
                        .Editable = flexEDKbdMouse '��ҩƷ���,��ʹ�̶�Ҳ���Զ���
                        .Cell(flexcpBackColor, .Rows - 1, COLP_����) = COLEditBackColor       'ǳ��
                    End If
                End If
                'Ӧ��
                If rsTmp!�������� = 1 Then
                    dbl��ǰӦ�� = dbl��ǰӦ�� * NVL(rsTmp!�����շ���, 100) / 100
                End If
                '����Ӱ�Ӽ�
                If gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                    dbl��ǰӦ�� = dbl��ǰӦ�� * (1 + NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                End If
                cur��ǰӦ�� = Format(dbl��ǰӦ��, gstrDec)
                
                'ʵ��
                If gbln��������ۿ� And (rsTmp!���� = 1 Or InStr(strHaveSub & ",", "," & rsTmp!�к� & "_" & rsTmp!�������� & ",") > 0) Then
                    cur��ǰʵ�� = Format(cur��ǰӦ��, gstrDec)
                    '�ۼ�ҽ���ϼ��������ۿ�
                    rsMain.Filter = "ҽ���к�=" & rsTmp!�к� & " And ��������=" & rsTmp!��������
                    rsMain!ҽ���ϼ� = NVL(rsMain!ҽ���ϼ�, 0) + cur��ǰʵ��
                    rsMain.Update
                ElseIf NVL(rsTmp!���ηѱ�, 0) = 0 And vsAdvice.TextMatrix(lngRow, COL_�ѱ�) <> "" Then
                    cur��ǰʵ�� = Format(ActualMoney(vsAdvice.TextMatrix(lngRow, COL_�ѱ�), rsTmp!������ĿID, cur��ǰӦ��, rsTmp!�շ�ϸĿID, NVL(rsTmp!ִ�п���ID, 0), _
                        int���� * dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1, NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                Else
                    cur��ǰʵ�� = Format(cur��ǰӦ��, gstrDec)
                End If
                
                dbl���� = dbl���� + dbl��ǰ����
                curӦ�� = curӦ�� + cur��ǰӦ��
                curʵ�� = curʵ�� + cur��ǰʵ��
                
                rsTmp.MoveNext
            Next
            If str�к� <> "" Then
                If Not (Val(.TextMatrix(.Rows - 1, COLP_���)) = 1 And dbl���� = 0) Then
                    .TextMatrix(.Rows - 1, COLP_����) = Format(dbl����, gstrDecPrice)
                    .Cell(flexcpData, .Rows - 1, COLP_����) = .TextMatrix(.Rows - 1, COLP_����) '��¼���ڻָ�����
                    .TextMatrix(.Rows - 1, COLP_Ӧ�ս��) = Format(curӦ��, gstrDec)
                    .TextMatrix(.Rows - 1, COLP_ʵ�ս��) = Format(curʵ��, gstrDec)
                End If
                cur�ϼ� = cur�ϼ� + Format(curʵ��, gstrDec)
            End If
        End If
        
        '���ܼ����ۿ�
        If gbln��������ۿ� And strHaveSub <> "" Then
            rsMain.Filter = 0
            Do While Not rsMain.EOF
                cur��ǰʵ�� = Format(ActualMoney(vsAdvice.TextMatrix(lngRow, COL_�ѱ�), rsMain!������ID, rsMain!ҽ���ϼ�), gstrDec)
                cur�ϼ� = cur�ϼ� - Val(.TextMatrix(rsMain!�����к�, COLP_ʵ�ս��))
                .TextMatrix(rsMain!�����к�, COLP_ʵ�ս��) = Format(Val(.TextMatrix(rsMain!�����к�, COLP_ʵ�ս��)) + (cur��ǰʵ�� - rsMain!ҽ���ϼ�), gstrDec)
                cur�ϼ� = cur�ϼ� + Val(.TextMatrix(rsMain!�����к�, COLP_ʵ�ս��))
                rsMain.MoveNext
            Loop
        End If
        
        '�����Ƿ���ʾ
        .ColHidden(COLP_����) = Not bln����
        
        '------------------------------------------------
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        '��λȱʡ��Ԫ
        If lngPreRow >= .FixedRows And lngPreRow <= .Rows - 1 Then
            .Row = lngPreRow
        Else
            .Row = .FixedRows
        End If
        If lngPreCol >= COLP_�Ƽ�ҽ�� And lngPreCol <= .Cols - 1 Then
            .Col = lngPreCol
        Else
            .Col = COLP_�Ƽ�ҽ��
        End If
        '��λ�������λ��
        If lngTopRow >= .FixedRows And lngTopRow <= .Rows - 1 Then
            .TopRow = lngTopRow
        End If
        If lngLeftCol >= COLP_�Ƽ�ҽ�� And lngLeftCol <= .Cols - 1 Then
            .LeftCol = lngLeftCol
        End If
        .Redraw = flexRDDirect
    End With
    
    '���»�����ʾ�ɼ��еķ���ҽ�����
    vsAdvice.TextMatrix(lngRow, COL_���) = Format(cur�ϼ�, gstrDec)
    ShowAdvicePrice = True
    
    Call ShowSendTotal
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub Initҽ����Ҫ����()
    Dim lngTmp As Long, strTmp As String
    Dim strPrivs As String
    
     '���Խ��������ʾ����������һ��
    strPrivs = GetInsidePrivs(pסԺҽ������)
    lblEatDrug.Visible = mstrEndPoint <> ""
    lblDrugType.Visible = gbln����ҩƷ�ֿ�����: cboDrugType.Visible = gbln����ҩƷ�ֿ�����
    If InStr(strPrivs, ";����ҩ������;") = 0 And InStr(strPrivs, ";����ҩ�Ƴ���;") = 0 Then
        fraҩƷ.Visible = False
        chkҩƷ��Ч(0).Enabled = False
        chkҩƷ��Ч(1).Enabled = False
        lblEatDrug.Visible = False
        lblDrugType.Visible = False: cboDrugType.Visible = False
        cboҩ��.Visible = False: lblҩ��.Visible = False
        txtִ�п���.Top = lbl����ʱ��.Top + 250
        lblִ�п���.Top = txtִ�п���.Top + 50
        cmdִ�п���.Top = txtִ�п���.Top + 30
        
        fraҽ������.Height = txtִ�п���.Top + txtִ�п���.Height + 50
    ElseIf InStr(strPrivs, ";������������;") = 0 And InStr(strPrivs, ";������������;") = 0 Then
        fra����.Visible = False
        chk������Ч(0).Enabled = False
        chk������Ч(1).Enabled = False
        cboҩ��.Top = lblEatDrug.Top + IIF(lblEatDrug.Visible, 230, 0)
        If gbln����ҩƷ�ֿ����� Then
            lblDrugType.Top = cboҩ��.Top + 50
            cboDrugType.Top = cboҩ��.Top
            cboҩ��.Top = cboDrugType.Top + cboDrugType.Height + 10
        End If
        lblҩ��.Top = cboҩ��.Top + 50
        txtִ�п���.Visible = False: lblִ�п���.Visible = False: cmdִ�п���.Visible = False
        fraҽ������.Height = cboҩ��.Top + cboҩ��.Height + 50
    Else
     
        txtִ�п���.Top = IIF(lblEatDrug.Visible, lblEatDrug.Top, lbl����ʱ��.Top) + 250
        If gbln����ҩƷ�ֿ����� Then
            lblDrugType.Top = txtִ�п���.Top + 50
            cboDrugType.Top = txtִ�п���.Top
            txtִ�п���.Top = cboDrugType.Top + cboDrugType.Height + 10
        End If
        lblִ�п���.Top = txtִ�п���.Top + 50
        cmdִ�п���.Top = txtִ�п���.Top + 30

        cboҩ��.Top = lblִ�п���.Top + 270
        lblҩ��.Top = cboҩ��.Top + 50
        fraҽ������.Height = cboҩ��.Top + cboҩ��.Height + 50
    End If
    '����ҩƷҽ����Ч������ѡ�������ʾ����������һ��
    If InStr(strPrivs, "����ҩ�Ƴ���") = 0 Then
        chkҩƷ��Ч(0).value = 0
        chkҩƷ��Ч(0).Enabled = False
    ElseIf InStr(strPrivs, "����ҩ������") = 0 Then
        chkҩƷ��Ч(1).value = 0
        chkҩƷ��Ч(1).Enabled = False
    End If
    '��������ҽ����Ч������ѡ�������ʾ����������һ��
    If InStr(strPrivs, "������������") = 0 Then
        chk������Ч(0).value = 0
        chk������Ч(0).Enabled = False
    ElseIf InStr(strPrivs, "������������") = 0 Then
        chk������Ч(1).value = 0
        chk������Ч(1).Enabled = False
    End If
    
    'ͨ��ҽ�������Զ���ȡ֮ǰ������Ȩ�޵Ķ��Զ�����
    If mblnOnePati And mblnAutoVerify Then
        If chkҩƷ��Ч(0).Enabled Then chkҩƷ��Ч(0).value = 1
        If chkҩƷ��Ч(1).Enabled Then chkҩƷ��Ч(1).value = 1
        If chk������Ч(0).Enabled Then chk������Ч(0).value = 1
        If chk������Ч(1).Enabled Then chk������Ч(1).value = 1
    Else
        'ȱʡҩƷҽ����Ч
        lngTmp = Val(zlDatabase.GetPara("ҩ������ҽ����Ч", glngSys, pסԺҽ������, "2", Array(chkҩƷ��Ч(0), chkҩƷ��Ч(1))))
        If lngTmp = 0 Or lngTmp = 2 Then
            If chkҩƷ��Ч(0).Enabled Then chkҩƷ��Ч(0).value = 1
        End If
        If lngTmp = 1 Or lngTmp = 2 Then
            If chkҩƷ��Ч(1).Enabled Then chkҩƷ��Ч(1).value = 1
        End If
        'ȱʡ����ҽ����Ч
        lngTmp = Val(zlDatabase.GetPara("��������ҽ����Ч", glngSys, pסԺҽ������, "2", Array(chk������Ч(0), chk������Ч(1))))
        If lngTmp = 0 Or lngTmp = 2 Then
            If chk������Ч(0).Enabled Then chk������Ч(0).value = 1
        End If
        If lngTmp = 1 Or lngTmp = 2 Then
            If chk������Ч(1).Enabled Then chk������Ч(1).value = 1
        End If
        '�쳣�����ٹ�һ��(��ʹû��Ȩ��),�Ա���������ؼ���״̬
        If chkҩƷ��Ч(0).value = 0 And chkҩƷ��Ч(1).value = 0 And chk������Ч(0).value = 0 And chk������Ч(1).value = 0 Then
            If InStr(strPrivs, ";����ҩ������;") = 0 And InStr(strPrivs, ";����ҩ�Ƴ���;") = 0 Then
                If chk������Ч(0).Enabled Then
                    chk������Ч(0).value = 1
                Else
                    chk������Ч(1).value = 1
                End If
            Else
                If chkҩƷ��Ч(0).Enabled Then
                    chkҩƷ��Ч(0).value = 1
                Else
                    chkҩƷ��Ч(1).value = 1
                End If
            End If
        End If
    End If
    
    lngTmp = Val(zlDatabase.GetPara("ҩƷ�����ӵ��쿪ʼ����", glngSys, pסԺҽ������, "0", Array(chkSendNow), InStr(strPrivs, ";��������;") > 0))
    chkSendNow.Tag = IIF(chkSendNow.Enabled, "", "0") '�Ƿ��ʼ����
    chkSendNow.value = lngTmp
    
    Call SetFaceByҽ����Ч
    
    'ҽ��״̬
    If mblnAutoVerify Then
        'ȱʡ��ȡ�¿���
        fra״̬.Visible = True
        opt״̬(0).value = True
    Else
        'ֻ��ȡ��У�Ե�
        fra״̬.Visible = False
        opt״̬(1).value = True
    End If
    
        
    '���ͽ���ʱ��
    strTmp = zlDatabase.GetPara("���ͽ���ʱ��", glngSys, pסԺҽ������, "0", Array(cboTime))
    cboTime.ListIndex = Val(strTmp)
    If cboTime.ListIndex = cboTime.ListCount - 1 Then
        mdatCurr = zlDatabase.Currentdate
        strTmp = zlDatabase.GetPara("���ͽ���ʱ��", glngSys, pסԺҽ������, "23:59:59", Array(dtpEnd))
        lngTmp = Val(zlDatabase.GetPara("����ʱ����", glngSys, pסԺҽ������, "0", Array(dtpEnd)))
        mblnIsSendInterval = dtpEnd.Enabled
        dtpEnd.value = Format(mdatCurr + lngTmp, "yyyy-MM-dd " & strTmp)
    End If
    Call SetFaceByҽ����Ч

End Sub

Private Sub SetFaceByҽ����Ч()
'���ܣ�����������Ч��ѡ�����ý���Ŀ�������
    Dim blnҩƷ As Boolean, bln���� As Boolean
    
    blnҩƷ = chkҩƷ��Ч(0).value = 1 Or chkҩƷ��Ч(1).value = 1
    bln���� = chk������Ч(0).value = 1 Or chk������Ч(1).value = 1
    cboDrugType.Enabled = blnҩƷ
    cboҩ��.Enabled = blnҩƷ
    tkpMain.Groups.Find(GP_��ҩ;��).Enabled = blnҩƷ
    tkpMain.Groups.Find(GP_ҩ���û�).Enabled = blnҩƷ
    If Not blnҩƷ Then
        tkpMain.Groups.Find(GP_��ҩ;��).Expanded = False
        tkpMain.Groups.Find(GP_ҩ���û�).Expanded = False
    End If
    
    txtִ�п���.Enabled = bln����
    cmdִ�п���.Enabled = bln����
    tkpMain.Groups.Find(GP_�������).Enabled = bln����
    If Not bln���� Then tkpMain.Groups.Find(GP_�������).Expanded = False
       
    '�����Ľ�ֹʱ��
    cboTime.Enabled = chkҩƷ��Ч(0).value = 1 Or chk������Ч(0).value = 1
    cboTime.BackColor = IIF(Not cboTime.Enabled, Me.BackColor, &H80000005)
    If mblnIsSendInterval Then
        '�ɽ�����ʱ�����������ó�ģ����������ɿ��ƻ�ʿ��෢�Ͷ�����ģ�С��3����û�����壬��Ϊѡ����켴Ϊ3��
        dtpEnd.Enabled = cboTime.Enabled And cboTime.ListIndex = cboTime.ListCount - 1
    End If
    chkSendNow.Enabled = chkSendNow.Tag <> "0" And chkҩƷ��Ч(0).value = 1 And chkҩƷ��Ч(0).Enabled = True
    '�����ҽ����Ч���ƣ��򲻿���ʱ��ղ���ֵ
    If Not chkSendNow.Enabled And chkSendNow.Tag <> "0" Then chkSendNow.value = 0
End Sub

Private Sub imgCheck_Click()
    If chkҩƷ��Ч(0).Visible And chkҩƷ��Ч(0).Enabled Then
        chkҩƷ��Ч(0).value = IIF(chkҩƷ��Ч(0).value = 0, 1, 0)
    End If
    If chkҩƷ��Ч(1).Visible And chkҩƷ��Ч(1).Enabled Then
        chkҩƷ��Ч(1).value = IIF(chkҩƷ��Ч(1).value = 0, 1, 0)
    End If
    
    If chk������Ч(0).Visible And chk������Ч(0).Enabled Then
        chk������Ч(0).value = IIF(chk������Ч(0).value = 0, 1, 0)
    End If
    If chk������Ч(1).Visible And chk������Ч(1).Enabled Then
        chk������Ч(1).value = IIF(chk������Ч(1).value = 0, 1, 0)
    End If
End Sub

Private Sub pic��ҩ_Resize()
    On Error Resume Next
    
    lvwWay.Left = 0
    lvwWay.Top = 0
    lvwWay.Width = pic��ҩ.ScaleWidth
    lvwWay.Height = pic��ҩ.ScaleHeight - cmdAllWay.Height - Screen.TwipsPerPixelY * 2
    lvwWay.ColumnHeaders(1).Width = lvwWay.Width - (GetSystemMetrics(SM_CXVSCROLL) + 4) * Screen.TwipsPerPixelX

    cmdNoWay.Left = pic��ҩ.ScaleWidth - cmdNoWay.Width - Screen.TwipsPerPixelX * 4
    cmdAllWay.Left = cmdNoWay.Left - cmdAllWay.Width
    cmdUnWay.Left = cmdAllWay.Left - cmdUnWay.Width
End Sub

Private Sub pic����_Resize()
    On Error Resume Next
    Dim lngTop As Long
    
    If Me.Visible = False Then Exit Sub
    
    If fraҩƷ.Visible Then lngTop = fraҩƷ.Height + 60
    If fra����.Visible Then
        fra����.Top = lngTop
        lngTop = lngTop + fra����.Height + 60
    End If
     
    If fra״̬.Visible Then
        fra״̬.Top = lngTop
        lngTop = lngTop + fra״̬.Height + 60
    End If
    fraҽ������.Top = lngTop
       
    
    fraPati.Top = fraҽ������.Top + fraҽ������.Height + 30
    Line1.Y1 = pic����.ScaleHeight - fraӤ��.Height - chk�Ӱ�Ӽ�.Height - 150
    Line1.Y2 = Line1.Y1
    fraPati.Height = Line1.Y1 - fraPati.Top - 60
    lvwPati.Height = fraPati.Height - lvwPati.Top - cmdAllPati.Height - 30
    cmdAllPati.Top = lvwPati.Top + lvwPati.Height + 30
    cmdNoPati.Top = cmdAllPati.Top
    cmdQuick.Top = cmdAllPati.Top
    
    
    fraӤ��.Top = Line1.Y1 + 60
    chk�Ӱ�Ӽ�.Top = fraӤ��.Top + fraӤ��.Height + 60
                
    fraҽ������.Width = pic����.ScaleWidth - fraҽ������.Left
    cboҩ��.Width = fraҽ������.Width - cboҩ��.Left - Screen.TwipsPerPixelX * 3
    cboDrugType.Left = cboҩ��.Left
    cboDrugType.Width = cboҩ��.Width
    txtִ�п���.Width = fraҽ������.Width - txtִ�п���.Left - Screen.TwipsPerPixelX * 3
    cmdִ�п���.Left = txtִ�п���.Left + txtִ�п���.Width - cmdִ�п���.Width - Screen.TwipsPerPixelX * 2
    
    
    fraPati.Width = pic����.ScaleWidth - fraPati.Left
    cboUnit.Width = fraPati.Width - cboUnit.Left - Screen.TwipsPerPixelX * 3
    lvwPati.Width = fraPati.Width - lvwPati.Left - Screen.TwipsPerPixelX * 3
    cmdNoPati.Left = fraPati.Width - cmdNoPati.Width - Screen.TwipsPerPixelX * 3
    cmdAllPati.Left = cmdNoPati.Left - cmdAllPati.Width
    
    Line1.X2 = pic����.ScaleWidth
    
End Sub

Private Sub pic���_Resize()
    On Error Resume Next
    
    lstClass.Left = 0
    lstClass.Top = 0
    lstClass.Width = pic���.ScaleWidth
    lstClass.Height = pic���.ScaleHeight
End Sub

Private Sub picҩ��_Resize()
    On Error Resume Next
    
    vsDept.Width = picҩ��.ScaleWidth
    vsDept.Height = picҩ��.ScaleHeight - chkDeptExchange.Height
End Sub

Private Sub txtִ�п���_GotFocus()
    Call zlControl.TxtSelAll(txtִ�п���)
End Sub

Private Sub txtִ�п���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call cmdִ�п���_Click
End Sub

Private Sub txtִ�п���_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtִ�п���.Text = txtִ�п���.Tag Then
            Call zlcommfun.PressKey(vbKeyTab)
        ElseIf txtִ�п���.Text = "" Then
            txtִ�п���.Text = "���п���"
            txtִ�п���.Tag = txtִ�п���.Text
            cmdִ�п���.Tag = ""
            Call zlcommfun.PressKey(vbKeyTab)
        Else
            strSQL = _
                " Select 0 as ID,'-' as ����,'���п���' as ����,NULL as ���� From Dual" & _
                " Union ALL" & _
                " Select Distinct A.ID,A.����,A.����,A.����" & _
                " From ���ű� A,��������˵�� B" & _
                " Where A.ID=B.����ID And B.������� IN(2,3)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"
            strSQL = "Select * From (" & strSQL & ")" & _
                " Where ���� Like [1] Or Upper(����) Like [2] Or Upper(����) Like [2]" & _
                " Order by ����"
            vRect = zlControl.GetControlRect(txtִ�п���.hwnd)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ִ�п���", False, "", "", False, False, True, _
                vRect.Left, vRect.Top, txtִ�п���.Height, blnCancel, False, True, _
                UCase(txtִ�п���.Text) & "%", mstrLike & UCase(txtִ�п���.Text) & "%")
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ��Ŀ��ҡ�", vbInformation, gstrSysName
                End If
                txtִ�п���.Text = txtִ�п���.Tag
                Call zlControl.TxtSelAll(txtִ�п���)
                txtִ�п���.SetFocus
            Else
                txtִ�п���.Text = rsTmp!����
                txtִ�п���.Tag = rsTmp!����
                cmdִ�п���.Tag = rsTmp!ID
                Call zlcommfun.PressKey(vbKeyTab)
            End If
        End If
    End If
End Sub

Private Sub txtִ�п���_Validate(Cancel As Boolean)
    If txtִ�п���.Text = "" Then
        txtִ�п���.Text = "���п���"
        txtִ�п���.Tag = txtִ�п���.Text
        cmdִ�п���.Tag = ""
    ElseIf txtִ�п���.Text <> txtִ�п���.Tag Then
        txtִ�п���.Text = txtִ�п���.Tag '�ָ���Ϊ�����
    End If
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'���ܣ����ĳ�ҩ���
    Dim rsDrug As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng���� As Long, lng��С���� As Long
    Dim dbl���� As Double, str�ֽ�ʱ�� As String
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim cur��� As Currency
    Dim bln�ڷ� As Boolean
    
    If Col = COL_ִ�п��� Or Col = COL_����ִ�� Then
        With vsAdvice
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsAdvice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
        End With
    ElseIf Col = COL_��� Then
        With vsAdvice
            If Val(.TextMatrix(Row, COL_�շ�ϸĿID)) = .ComboData Then Exit Sub
            'ҩƷ�����Ϣ
            .TextMatrix(Row, COL_�շ�ϸĿID) = .ComboData
            Set rsDrug = GetDrugInfo(Val(.TextMatrix(Row, COL_������ĿID)), Val(.TextMatrix(Row, COL_�շ�ϸĿID)), Val(.TextMatrix(Row, COL_ִ�п���ID)))
            .TextMatrix(Row, COL_���) = rsDrug!���� & IIF(Not IsNull(rsDrug!����), "(" & rsDrug!���� & ")", "") & IIF(Not IsNull(rsDrug!���), " " & rsDrug!���, "")
            .TextMatrix(Row, COL_����ϵ��) = rsDrug!����ϵ��
            .TextMatrix(Row, COL_סԺ��װ) = rsDrug!סԺ��װ
            .TextMatrix(Row, COL_סԺ��λ) = NVL(rsDrug!סԺ��λ)
            .TextMatrix(Row, COL_�Ƿ���) = rsDrug!�Ƿ���
            .TextMatrix(Row, COL_ҩ������) = rsDrug!ҩ������
            .TextMatrix(Row, COL_�ɷ����) = NVL(rsDrug!�ɷ����, 0)
            .TextMatrix(Row, COL_���) = Format(NVL(rsDrug!���, 0), "0.00000")
   
            'ҽ�������Ϣ
            strSQL = _
                " Select A.ID,a.���id as ��ID,A.�������,A.��ʼִ��ʱ��,A.�ϴ�ִ��ʱ��,A.ִ����ֹʱ��,A.ִ��ʱ�䷽��," & _
                " A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.��������,A.�ɷ����,B.��Ժ����,A.ҽ��״̬,A.�״�����,A.ҽ����Ч,A.������־,A.���״̬" & _
                " From ����ҽ����¼ A,������ҳ B" & _
                " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.ID=[1]"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(Row, COL_ID)))
            
            If mstrEndPoint <> "" Then
                If InStr(",5,6,7,", .TextMatrix(Row, COL_�������)) > 0 Then
                    For i = Row + 1 To .Rows - 1
                        If .TextMatrix(i, COL_�������) = "E" And (Val(.TextMatrix(i, COL_��������)) = 2 Or Val(.TextMatrix(i, COL_��������)) = 4) Then
                            If Val(.TextMatrix(i, COL_ִ�з���)) = 4 Then bln�ڷ� = True
                            Exit For
                        End If
                    Next
                End If
            End If
            
            '���¼�������,����,�ֽ�ʱ��
            Call Calc��������ʱ��(dbl����, lng����, str�ֽ�ʱ��, mstrEnd, rsTmp, rsDrug, bln�ڷ�)
            
            .TextMatrix(Row, COL_����) = FormatEx(dbl����, 5)
            .TextMatrix(Row, COL_������λ) = NVL(rsDrug!סԺ��λ)
            
            .TextMatrix(Row, COL_����) = lng����
            .TextMatrix(Row, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
            If str�ֽ�ʱ�� <> "" Then
                .TextMatrix(Row, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "yyyy-MM-dd HH:mm")
                .TextMatrix(Row, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "yyyy-MM-dd HH:mm")
            End If
                        
            'ͬ�����ĸ�ҩ;���Ĵ���
            i = .FindRow(.TextMatrix(Row, COL_���ID), , COL_ID)
            .TextMatrix(i, COL_����) = .TextMatrix(Row, COL_����)
            .TextMatrix(i, COL_����) = .TextMatrix(Row, COL_����) '��ͬ
            .TextMatrix(i, COL_�ֽ�ʱ��) = .TextMatrix(Row, COL_�ֽ�ʱ��)
            .TextMatrix(i, COL_�״�ʱ��) = .TextMatrix(Row, COL_�״�ʱ��)
            .TextMatrix(i, COL_ĩ��ʱ��) = .TextMatrix(Row, COL_ĩ��ʱ��)
                                    
            'һ����ҩ�İ���С�������㣺����ҩƷ����������
            If RowInһ����ҩ(Row, lngBegin, lngEnd) Then
                For i = lngBegin To lngEnd
                    If Val(.TextMatrix(i, COL_����)) < lng��С���� Or lng��С���� = 0 Then
                        lng��С���� = Val(.TextMatrix(i, COL_����))
                    End If
                Next
                For i = lngBegin To lngEnd + 1
                    If Val(.TextMatrix(i, COL_����)) > lng��С���� Then
                        .TextMatrix(i, COL_����) = lng��С����
                        .TextMatrix(i, COL_�ֽ�ʱ��) = Trim�ֽ�ʱ��(lng��С����, .TextMatrix(i, COL_�ֽ�ʱ��))
                        .TextMatrix(i, COL_�״�ʱ��) = Format(Split(.TextMatrix(i, COL_�ֽ�ʱ��), ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(i, COL_�ֽ�ʱ��), ",")(lng��С���� - 1), "yyyy-MM-dd HH:mm")
                    End If
                Next
            Else
                lngBegin = Row: lngEnd = Row
            End If
            
            '���¼��㲢��ʾ��ǰҩƷ����ҩ;���Ľ��ͼƼ�
            mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngBegin, COL_ID)) & " Or ҽ��ID=" & Val(.TextMatrix(lngEnd + 1, COL_ID))
            Do While Not mrsPrice.EOF
                mrsPrice.Delete
                mrsPrice.Update
                mrsPrice.MoveNext
            Loop
            
            '��ǼƼ����ݱ仯
            .Cell(flexcpData, lngBegin, COL_���) = 1
            .Cell(flexcpData, lngEnd + 1, COL_���) = 1
            
            cur��� = 0
            Call LoadAdvicePrice(lngBegin, cur���, rsDrug)
            .TextMatrix(lngBegin, COL_���) = Format(cur���, gstrDec)
            cur��� = 0
            Call LoadAdvicePrice(lngEnd + 1, COL_���)
            .TextMatrix(lngEnd + 1, COL_���) = Format(cur���, gstrDec)
            
            'һ����ҩ�ĵ�һ��(�����)����ʾ������ҩ;���Ľ��
            .TextMatrix(lngBegin, COL_���) = Format(Val(.TextMatrix(lngBegin, COL_���)) + Val(.TextMatrix(lngEnd + 1, COL_���)), gstrDec)
            
            '���ݿ�����ѡ��״̬
            If Val(.TextMatrix(Row, COL_����)) > Val(.TextMatrix(Row, COL_���)) Then
                If TheStockCheck(Val(.TextMatrix(Row, COL_ִ�п���ID)), .TextMatrix(Row, COL_�������)) = 2 _
                    Or Val(.TextMatrix(Row, COL_ҩ������)) = 1 Or Val(.TextMatrix(Row, COL_�Ƿ���)) = 1 Then
                    .Cell(flexcpData, Row, COL_ѡ��) = 1
                    Set .Cell(flexcpPicture, Row, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                ElseIf TheStockCheck(Val(.TextMatrix(Row, COL_ִ�п���ID)), .TextMatrix(Row, COL_�������)) = 1 Then
                    .Cell(flexcpData, Row, COL_ѡ��) = 0
                    Set .Cell(flexcpPicture, Row, COL_ѡ��) = Nothing
                ElseIf TheStockCheck(Val(.TextMatrix(Row, COL_ִ�п���ID)), .TextMatrix(Row, COL_�������)) = 0 Then
                    .Cell(flexcpData, Row, COL_ѡ��) = 0
                    Set .Cell(flexcpPicture, Row, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
                End If
            ElseIf Val(.TextMatrix(Row, COL_����)) <= Val(.TextMatrix(Row, COL_���)) Then
                .Cell(flexcpData, Row, COL_ѡ��) = 0
                Set .Cell(flexcpPicture, Row, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
            End If
            Call RowSelectSame(Row, COL_ѡ��)
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
            Call ShowSendTotal
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAdvice
        If .Redraw <> flexRDNone And Not .RowHidden(NewRow) Then
            '���ݿɷ�༭���ñ༭���Լ��������
            If NewCol = COL_��� Then
                .ComboList = .Cell(flexcpData, NewRow, NewCol)
                .FocusRect = flexFocusLight
            ElseIf CellEditable(NewRow, NewCol) Then
                .ComboList = "..."
                Set .CellButtonPicture = Me.Picture
                .FocusRect = flexFocusHeavy
            Else
                .ComboList = ""
                .FocusRect = flexFocusLight
            End If
            
            If OldRow <> NewRow Then
                If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
                    Call ShowAdvicePrice(NewRow)
                End If
            End If
        End If
        .ForeColorSel = .Cell(flexcpForeColor, NewRow, COL_Ƶ��)
    End With
End Sub

Private Function Should����ִ��(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ����ҽ����(�ɼ���)�Ƿ�������ø��ӵ�ִ�п���
    Dim lngRow2 As Long, i As Long
    
    If lngRow = 0 Then Exit Function
    
    lngRow2 = -1
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_ID)) = 0 Then Exit Function
        If .TextMatrix(lngRow, COL_�������) = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 _
            And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) _
            And InStr(",7,E,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
            '��ҩ�÷�
            lngRow2 = lngRow
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            '��ҩ;��
            lngRow2 = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1, COL_ID)
        ElseIf .TextMatrix(lngRow, COL_�������) = "F" Then
            '��������
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If .TextMatrix(i, COL_�������) = "G" Then
                        lngRow2 = i: Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        ElseIf .TextMatrix(lngRow, COL_�������) = "E" _
            And .TextMatrix(lngRow - 1, COL_�������) = "C" _
            And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
            '�ɼ���ʽ
            lngRow2 = lngRow
        End If
        
        '������Ժ��ִ��
        If lngRow2 <> -1 Then
            If InStr(",0,5,", Val(.TextMatrix(lngRow2, COL_ִ������ID))) = 0 Then
                Should����ִ�� = True
            End If
        End If
    End With
End Function

Private Sub vsAdvice_AfterUserFreeze()
    With vsAdvice
        If .FrozenCols < COL_ѡ�� + 1 - .FixedCols Then
            .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    With vsAdvice
        If Col = col_ҽ������ Or Col = COL_��� Then
            If Not .ColHidden(COL_���) Then
                .AutoSize col_ҽ������, COL_���
            Else
                .AutoSize col_ҽ������
            End If
            .RowHeight(0) = 320
        ElseIf Row = -1 Then
            lngW = Me.TextWidth(.TextMatrix(.FixedRows - 1, Col) & "A")
            If .ColWidth(Col) < lngW Then
                .ColWidth(Col) = lngW
            ElseIf .ColWidth(Col) > .Width * 0.5 Then
                .ColWidth(Col) = .Width * 0.5
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_ѡ�� Then Cancel = True
End Sub

Private Sub vsAdvice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim vPoint As PointAPI, blnCancel As Boolean
    
    strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.������� IN(2,3)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " Order by A.����"
    With vsAdvice
        vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "ִ�п���", , , , , , True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            Call SetDeptInput(Row, Col, rsTmp)
            Call vsAdvice_AfterRowColChange(-1, -1, Row, Col) '������ʾ�Ƽ�ִ�п���
        Else
            If Not blnCancel Then
                MsgBox "û�п��õĿ������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_ChangeEdit()
    If vsAdvice.Col = COL_��� Then
        Call vsAdvice_AfterEdit(vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Sub vsAdvice_DblClick()
    With vsAdvice
        If .MouseCol = COL_ѡ�� And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAdvice_KeyPress(32)
        End If
    End With
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = COL_Ƶ��: lngRight = COL_�÷�
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_����: lngRight = COL_ҽ����Ч
            If Not Between(Col, lngLeft, lngRight) Then
                Exit Sub
            End If
        End If
        
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '���б����±���(���������õ��±��ߴ�Ϊ2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            If Val(.TextMatrix(Row, COL_ҽ��״̬)) = 1 Then
                SetBkColor hDC, OS.SysColor2RGB(BackColorNew)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 127 Then '���ֱ�����뺺�ֵ�����
        Call vsAdvice_KeyPress(KeyCode)
    End If
End Sub

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
    Dim i As Long
    
    With vsAdvice
        For i = lngRow + 1 To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        If i > .Rows - 1 Then .Row = .FixedRows
        If .RowHidden(.Row) Then .Row = lngRow
        Call .ShowCell(.Row, .Col)
    End With
End Sub

Private Sub EnterNextCellPrice(ByVal lngRow As Long, ByVal lngCol As Long)
'���ܣ���λ���۱�����һ����������ĵ�Ԫ��
    Dim i As Long, j As Long
    
    With vsPrice
        '��ǰ��Ԫ�����δ��������,���˳�
        If CellEditablePrice(lngRow, lngCol) Then
            If lngCol = COLP_���� And Val(.TextMatrix(lngRow, lngCol)) = 0 Then
                Exit Sub
            ElseIf .TextMatrix(lngRow, lngCol) = "" Then
                Exit Sub
            End If
        End If
        
        '����һ��Ԫ��ʼѭ������
        For i = lngRow To .Rows - 1
            For j = IIF(i = lngRow, lngCol + 1, COLP_�Ƽ�ҽ��) To .Cols - 1
                If CellEditablePrice(i, j) Then Exit For
            Next
            If j <= .Cols - 1 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        Else
            '��ǰ�����û���ҵ���һ���ɱ༭��Ԫ,�������Ƽ�ҽ��,������һ����
            If CStr(.ColData(COLP_�Ƽ�ҽ��)) <> "" Then
                '��ǰ��δ��������,��λ����������Ԫ
                If .TextMatrix(lngRow, COLP_�Ƽ�ҽ��) = "" Then
                    .Col = COLP_�Ƽ�ҽ��
                ElseIf .TextMatrix(lngRow, COLP_�Ƽ�����) = "" Then
                    .Col = COLP_�Ƽ�����
                ElseIf .TextMatrix(lngRow, COLP_�շ���Ŀ) = "" Then
                    .Col = COLP_�շ���Ŀ
                ElseIf Val(.TextMatrix(lngRow, COLP_���)) = 1 _
                    And Val(.TextMatrix(lngRow, COLP_����)) = 0 _
                    And CellEditablePrice(lngRow, COLP_����) Then
                    .Col = COLP_����
                Else
                    .AddItem "", .Rows
                    .Row = .Rows - 1: .Col = COLP_�Ƽ�ҽ��
                    
                    'ȱʡѡ��Ƽ�ҽ��(�������)
                    Call ShowDefaultRow
                End If
            Else
                If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1 '���ɱ༭ʱ���ⶨһ��
            End If
        End If
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub ShowDefaultRow()
'���ܣ����ڿ��ԼƼ۵�ҽ��,ȱʡ����һ�в�����ȱʡ�Ƽ�ҽ��
'˵����ComboList="#ҽ��ID1;�Ƽ�ҽ��1|#ҽ��ID2;�Ƽ�ҽ��2|..."
'      ���ڵ�һ����ʾ�Ƽ۱�ͻس�������ʱ����
    Dim arrCombo As Variant, lngRow As Long, i As Long
    Dim lngҽ��ID As Long, lng�к� As Long, str�Ƽ�ҽ�� As String
    Dim blnFirst As Boolean, blnHave As Boolean
    
    With vsPrice
        If .ColData(COLP_�Ƽ�ҽ��) <> "" And .Editable <> flexEDNone Then
            arrCombo = Split(.ColData(COLP_�Ƽ�ҽ��), "|")
            
            If Val(.TextMatrix(.Rows - 1, COLP_�к�)) <> 0 _
                And Val(.TextMatrix(.Rows - 1, COLP_�շ�ϸĿID)) <> 0 Then
                '��һ����ʾʱȱʡ����һ��
                blnFirst = True
                .AddItem "", .Rows
                .Row = .Rows - 1
            End If
            lngRow = .Rows - 1
            
            '���ǵ�һ����ʾʱȱʡ�Ƽ�ҽ������һ����ͬ
            If lngRow > 1 And Not blnFirst Then
                If Val(.TextMatrix(lngRow - 1, COLP_�̶�)) = 0 _
                    And Val(.TextMatrix(lngRow - 1, COLP_�к�)) <> 0 Then
                    blnHave = True
                End If
            End If
            For i = 0 To UBound(arrCombo)
                lngҽ��ID = Val(Mid(Mid(arrCombo(i), 1, InStr(arrCombo(i), ";") - 1), 2))
                str�Ƽ�ҽ�� = Replace(arrCombo(i), "#" & lngҽ��ID & ";", "")
                lng�к� = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
                If blnHave Then
                    If lng�к� = Val(.TextMatrix(lngRow - 1, COLP_�к�)) Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
                        
            'ģ��ѡ������Ƽ�ҽ��
            .TextMatrix(lngRow, COLP_�к�) = lng�к�
            .TextMatrix(lngRow, COLP_�Ƽ�ҽ��) = str�Ƽ�ҽ��
            .Cell(flexcpData, lngRow, COLP_�Ƽ�ҽ��) = .TextMatrix(lngRow, COLP_�Ƽ�ҽ��)
            
            'ֻ��һ���Ƽ�ҽ��ʱ����ͣ��
            If UBound(arrCombo) = 0 Then
                .Col = COLP_�շ���Ŀ
            Else
                .Col = COLP_�Ƽ�ҽ��
            End If
        End If
        Call .ShowCell(.Row, .Col)
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        ElseIf KeyAscii = 32 And .Col = COL_ѡ�� Then
            KeyAscii = 0
            If .Cell(flexcpData, .Row, COL_ѡ��) = 0 Then
                If .Cell(flexcpPicture, .Row, COL_ѡ��) Is Nothing Then
                    If InStr(mstrNoneIDs, "," & .TextMatrix(.Row, COL_ID) & ",") > 0 And Not mbln������ҩ Then
                        MsgBox "��ҽ������Ч������Ƥ�Խ�����������ͣ�", vbInformation, gstrSysName
                        Exit Sub
                    Else
                        Set .Cell(flexcpPicture, .Row, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
                    End If
                Else
                    Set .Cell(flexcpPicture, .Row, COL_ѡ��) = Nothing
                End If
                Call RowSelectSame(.Row, .Col)
                Call ShowSendTotal
            End If
        Else
            If CellEditable(.Row, .Col) And .ComboList = "..." Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsAdvice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strInput As String
    Dim vPoint As PointAPI, blnCancel As Boolean
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Col = COL_��� Then
                Call vsAdvice_KeyPress(13)
            ElseIf Col = COL_����ִ�� And .EditText <> "" Then
                strInput = UCase(.EditText)
                strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
                    " From ���ű� A,��������˵�� B" & _
                    " Where A.ID=B.����ID And B.������� IN(2,3)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
                    " Order by A.����"
                With vsAdvice
                    vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ִ�п���", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
                    If Not rsTmp Is Nothing Then
                        Call SetDeptInput(Row, Col, rsTmp)
                        .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                        Call EnterNextCell(Row, Col)
                    Else
                        If Not blnCancel Then
                            MsgBox "û���ҵ�ƥ��Ŀ��ҡ�", vbInformation, gstrSysName
                        End If
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsAdvice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    End If
                End With
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAdvice.EditSelStart = 0
    vsAdvice.EditSelLength = zlcommfun.ActualLen(vsAdvice.EditText)
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not CellEditable(Row, Col) Then Cancel = True
End Sub

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'���ܣ��жϷ���ҽ���嵥�е�Ԫ���Ƿ���Ա༭
    Dim bln�ɼ� As Boolean, blnDo As Boolean, i As Long
    Dim bln�Ŀ��� As Boolean
    
    If lngRow = 0 Then Exit Function
    
    bln�Ŀ��� = InStr(";" & GetInsidePrivs(pסԺҽ������) & ";", ";�޸ķ�ҩҽ����ִ�п���;") > 0
    
    With vsAdvice
        CellEditable = .Editable
        If lngCol = COL_��� Then
            CellEditable = .ComboList <> ""
        ElseIf lngCol = COL_ִ�п��� Then
            If InStr("5,6,7", .TextMatrix(lngRow, COL_�������)) > 0 Then
                'ҩƷִ�п��Ҳ�����ֱ���������
                CellEditable = False
            ElseIf bln�Ŀ��� Then
                '���������ֻ����һ���������ã�������ѡ��
                If Val(.TextMatrix(lngRow, COL_���ID)) = 0 And .TextMatrix(lngRow, COL_�������) = "E" And .TextMatrix(lngRow - 1, COL_�������) <> "E" _
                    And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then bln�ɼ� = True
                
                If bln�ɼ� Then
                    blnDo = False
                    For i = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                            If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                                blnDo = True: Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                Else
                    blnDo = InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������ID))) = 0
                End If
                If Not blnDo Then CellEditable = False
            Else
                CellEditable = False
            End If
        ElseIf lngCol = COL_����ִ�� And bln�Ŀ��� Then
            CellEditable = Should����ִ��(lngRow)
        Else
            CellEditable = False
        End If
    End With
End Function

Private Function CellEditablePrice(ByVal lngRow As Long, ByVal lngCol As Long, Optional bln�Ǳ��� As Boolean) As Boolean
'���ܣ��жϼ۱��е�Ԫ���Ƿ���Ա༭
    Dim lng�к� As Long
    
    With vsPrice
        bln�Ǳ��� = False
        CellEditablePrice = .Editable
        lng�к� = Val(.TextMatrix(lngRow, COLP_�к�))
        If lngCol = COLP_ִ�п��� Then
            '�������õ�����,��ҩ��ҩƷ�Ƽ۵�ִ�п��ҿ����޸�
            If Not ((.TextMatrix(lngRow, COLP_�շ����) = "4" And Val(.TextMatrix(lngRow, COLP_��������)) = 1 _
                Or InStr(",5,6,7,", .TextMatrix(lngRow, COLP_�շ����)) > 0) And InStr(",4,5,6,7,", vsAdvice.TextMatrix(lng�к�, COL_�������)) = 0) Then
                CellEditablePrice = False
            End If
            If .TextMatrix(lngRow, COLP_�շ���Ŀ) = "" Or .TextMatrix(lngRow, COLP_�к�) = "" Then
                CellEditablePrice = False
            End If
        ElseIf Val(.TextMatrix(lngRow, COLP_�̶�)) <> 0 Then
            '�̶������н������޸ı��
            If Not (Val(.TextMatrix(lngRow, COLP_���)) = 1 And lngCol = COLP_����) Then
                CellEditablePrice = False
            End If
        Else
            If lngCol = COLP_���� Then
                If Val(.TextMatrix(lngRow, COLP_���)) <> 1 Then
                    CellEditablePrice = False
                Else
                    '�Ǳ���ִ�еı����Ŀ�������۸�
                    If lng�к� <> 0 Then
                        If Not Check����ִ��(Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))) Then
                            bln�Ǳ��� = True: CellEditablePrice = False
                        End If
                    End If
                End If
            ElseIf lngCol <> COLP_�Ƽ�ҽ�� And lngCol <> COLP_�Ƽ����� And lngCol <> COLP_�շ���Ŀ Then
                CellEditablePrice = False
            End If
        End If
    End With
End Function

Private Sub vsDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsDept.Cell(flexcpData, Row, Col) = CLng(vsDept.ComboData)
    mrsҩ��.Filter = "ID=" & CLng(vsDept.Cell(flexcpData, Row, 0))
    mrsҩ��!��ID = CLng(vsDept.Cell(flexcpData, Row, Col))
    mrsҩ��.Update
    mblnChangeIF = True
End Sub

Private Sub vsDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsDept.Editable <> flexEDNone And NewCol = 1 Then
        vsDept.FocusRect = flexFocusSolid
    Else
        vsDept.FocusRect = flexFocusLight
    End If
End Sub

Private Sub vsDept_ChangeEdit()
    Call vsDept_AfterEdit(vsDept.Row, vsDept.Col)
End Sub

Private Sub vsDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsDept.Col = 1 Then
            If vsDept.Row + 1 <= vsDept.Rows - 1 Then
                vsDept.Row = vsDept.Row + 1
            Else
                Call zlcommfun.PressKey(vbKeyTab)
                vsDept.Row = vsDept.FixedRows + 1
            End If
        Else
            vsDept.Col = 1
        End If
        Call vsDept.ShowCell(vsDept.Row, vsDept.Col)
    End If
End Sub

Private Sub vsDept_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsDept.ComboIndex <> -1 Then
            Call vsDept_KeyPress(13)
        End If
    End If
End Sub

Private Sub vsDept_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngԭ��ID As Long, lngҽ��ID As Long
    Dim int�������� As Integer, intԭ�������� As Integer
    Dim lng�շ�ϸĿID As Long, i As Long
    Dim blnHaveSub As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If Col = COLP_�Ƽ�ҽ�� Then
            '�������ComboData,TextMatrixȡֵ��ΪComboData
            If .Cell(flexcpTextDisplay, Row, Col) <> .Cell(flexcpData, Row, Col) Then
                lngҽ��ID = .ComboData
                If lngҽ��ID < 0 Then
                    int�������� = Val(Left(Abs(lngҽ��ID), 1))
                    lngҽ��ID = Val(Mid(Abs(lngҽ��ID), 2))
                End If
                lngԭ��ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_�к�)), COL_ID))
                intԭ�������� = Val(.TextMatrix(Row, COLP_��������))
                lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                                
                '���üƼ�ҽ���Ƿ�������ͬ�շ�ϸĿ
                If lng�շ�ϸĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                    If Not mrsPrice.EOF Then
                        MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """�Ѿ��������շ���Ŀ""" & .TextMatrix(Row, COLP_�շ���Ŀ) & """��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                'ԭ����ҽ������д�������Ҫ����һ��(�����ǹ̶����ɶ���)
                If lngԭ��ID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngԭ��ID & " And ��������=" & intԭ�������� & " And ����=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(Row, COLP_����) <> "" Then
                        MsgBox """" & .Cell(flexcpData, Row, Col) & """����Ҫ����һ�������Ƽ���Ŀ��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                '���������˵ļƼ�ҽ������
                i = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
                .TextMatrix(Row, COLP_�к�) = i
                .TextMatrix(Row, COLP_��������) = int��������
                .TextMatrix(Row, Col) = .Cell(flexcpTextDisplay, Row, Col)
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                If lng�շ�ϸĿID <> 0 Then
                    '��ѡ���ҽ���Ƿ��д�������޸ĺ����Ŀ�Ƿ����
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And ����=1"
                    If Not mrsPrice.EOF Then blnHaveSub = True
                    .TextMatrix(Row, COLP_����) = IIF(blnHaveSub, "��", "")
                
                    '���»����Ӽ�¼������
                    If lngԭ��ID = 0 Then
                        mrsPrice.AddNew '����
                    Else '����
                        mrsPrice.Filter = "ҽ��ID=" & lngԭ��ID & " And ��������=" & intԭ�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                    End If
                    mrsPrice!ҽ��ID = lngҽ��ID
                    If Val(vsAdvice.TextMatrix(i, COL_���ID)) <> 0 Then
                        mrsPrice!���ID = vsAdvice.TextMatrix(i, COL_���ID)
                    Else
                        mrsPrice!���ID = Null
                    End If
                    mrsPrice!�������� = int��������
                    mrsPrice!�շѷ�ʽ = 0
                    If lngԭ��ID = 0 Then
                        mrsPrice!�շ�ϸĿID = lng�շ�ϸĿID
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_�Ƽ�����))
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_����))
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_��������))
                        mrsPrice!��� = Val(.TextMatrix(Row, COLP_���))
                        mrsPrice!�̶� = 0
                    End If
                    mrsPrice!���� = IIF(blnHaveSub, 1, 0)
                    mrsPrice.Update
                    
                    '��ǼƼ����ݱ仯
                    If lngԭ��ID <> 0 Then
                        vsAdvice.Cell(flexcpData, vsAdvice.FindRow(CStr(lngԭ��ID), , COL_ID), COL_���) = 1
                    End If
                    vsAdvice.Cell(flexcpData, i, COL_���) = 1
                    
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
            End If
        ElseIf Col = COLP_�շ���Ŀ Or Col = COLP_ִ�п��� Then
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
        ElseIf Col = COLP_�Ƽ����� Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '���¼�¼��
            lngҽ��ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_�к�)), COL_ID))
            int�������� = Val(.TextMatrix(Row, COLP_��������))
            lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
            If lngҽ��ID <> 0 And lng�շ�ϸĿID <> 0 Then
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                mrsPrice!���� = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                '��ǼƼ����ݱ仯
                vsAdvice.Cell(flexcpData, Val(.TextMatrix(Row, COLP_�к�)), COL_���) = 1
                
                Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
            End If
        ElseIf Col = COLP_���� Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            If CheckScope(.Cell(flexcpData, Row, COLP_Ӧ�ս��), .Cell(flexcpData, Row, COLP_ʵ�ս��), .TextMatrix(Row, Col)) <> "" Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gstrDecPrice)
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '���¼�¼��
            lngҽ��ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_�к�)), COL_ID))
            int�������� = Val(.TextMatrix(Row, COLP_��������))
            lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
            If lngҽ��ID <> 0 And lng�շ�ϸĿID <> 0 Then
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                mrsPrice!���� = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                '��ǼƼ����ݱ仯
                vsAdvice.Cell(flexcpData, Val(.TextMatrix(Row, COLP_�к�)), COL_���) = 1
                
                Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngRow As Long
    
    '���ݿɷ�༭����
    If Not CellEditablePrice(NewRow, NewCol) Then
        vsPrice.ComboList = ""
        vsPrice.FocusRect = flexFocusLight
    Else
        vsPrice.FocusRect = flexFocusSolid
        If NewCol = COLP_�Ƽ�ҽ�� Then
            vsPrice.ComboList = vsPrice.ColData(NewCol)
        ElseIf NewCol = COLP_�շ���Ŀ Or NewCol = COLP_ִ�п��� Then
            vsPrice.ComboList = "..."
        Else
            vsPrice.ComboList = ""
        End If
    End If
        
    If NewRow <> OldRow Then
        With vsPrice
            stbThis.Panels(2).Text = ""
            lngRow = Val(.TextMatrix(NewRow, COLP_�к�))
            If lngRow <> 0 And .Cell(flexcpData, NewRow, COLP_���) <> "" Then
                If InStr(",5,6,7,", .Cell(flexcpData, NewRow, COLP_���)) > 0 _
                    Or .Cell(flexcpData, NewRow, COLP_���) = "4" And Val(.Cell(flexcpData, NewRow, COLP_��������)) = 1 Then
                    '��ʾҩƷ���������ĵĿ��
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                        If InStr(GetInsidePrivs(pסԺҽ������), "��ʾҩƷ���") = 0 Then
                            stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, COL_���) & "��" & vsAdvice.TextMatrix(lngRow, COL_ִ�п���) & IIF(Val(vsAdvice.TextMatrix(lngRow, COL_���)) > 0, "�п��", "�޿��")
                        Else
                            stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, COL_���) & "��" & vsAdvice.TextMatrix(lngRow, COL_ִ�п���) & "���ÿ�棺" & _
                                FormatEx(Val(vsAdvice.TextMatrix(lngRow, COL_���)), 5) & vsAdvice.TextMatrix(lngRow, COL_סԺ��λ)
                        End If
                    Else
                        'ͬһ������ȡ:ҩƷ��סԺ��λ,���İ��ۼ۵�λ
                        If InStr(GetInsidePrivs(pסԺҽ������), "��ʾҩƷ���") = 0 Then
                            If GetStock(Val(.Cell(flexcpData, NewRow, COLP_�շ���Ŀ)), Val(.Cell(flexcpData, NewRow, COLP_ִ�п���))) > 0 Then
                                stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "��" & .TextMatrix(NewRow, COLP_ִ�п���) & "�п��"
                            Else
                                stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "��" & .TextMatrix(NewRow, COLP_ִ�п���) & "�޿��"
                            End If
                        Else
                            stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "��" & .TextMatrix(NewRow, COLP_ִ�п���) & "���ÿ�棺" & _
                                FormatEx(GetStock(Val(.Cell(flexcpData, NewRow, COLP_�շ���Ŀ)), Val(.Cell(flexcpData, NewRow, COLP_ִ�п���))), 5) & .TextMatrix(NewRow, COLP_��λ)
                        End If
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng�к� As Long, i As Long
    Dim str��ĿIDs As String, blnCancel As Boolean
    Dim lngҽ��ID As Long, lngԭ��ĿID As Long
    Dim int�������� As Integer, vPoint As PointAPI
    Dim strSQL2 As String
    
    With vsPrice
        lng�к� = Val(.TextMatrix(Row, COLP_�к�))
        If Col = COLP_�շ���Ŀ Then
            '����ѡ�����е���Ŀ
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COLP_�к�)) = lng�к� And lng�к� <> 0 And i <> Row Then
                    str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COLP_�շ�ϸĿID))
                End If
            Next
            str��ĿIDs = Mid(str��ĿIDs, 2)
            If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(lng�к�, COL_����ID)), Val(vsAdvice.TextMatrix(lng�к�, COL_��ҳID)), "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
        
            
            strSQL = _
                " Select Distinct 0 as ĩ��,To_Number('999999999'||����) as ID,-NULL as �ϼ�ID," & _
                " CHR(13)||���� as ����,Decode(����,1,'����ҩ',2,'�г�ҩ',3,'�в�ҩ',7,'��������') as ����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������,NULL as ҽ������,NULL as ˵��,NULL as �۸�," & _
                " -NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as ȱʡ�۸�ID,-NULL as �Ƿ���ID,Null as ���ID,-Null as ��������ID" & _
                " From ���Ʒ���Ŀ¼ Where ���� in (1,2,3,7) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as ĩ��,-ID as ID,Nvl(-�ϼ�ID,To_Number('999999999'||����)) as �ϼ�ID,����,����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������,NULL as ҽ������,NULL as ˵��,NULL as �۸�," & _
                " -NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as ȱʡ�۸�ID,-NULL as �Ƿ���ID,Null as ���ID,-Null as ��������ID" & _
                " From ���Ʒ���Ŀ¼ Where ���� in (1,2,3,7) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������,NULL as ҽ������," & _
                " NULL as ˵��,NULL as �۸�,-NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as ȱʡ�۸�ID,-NULL as �Ƿ���ID,Null as ���ID,-Null as ��������ID" & _
                " From �շѷ���Ŀ¼ Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strSQL2 = _
                " Select ĩ��,ID,�ϼ�ID,����,����,��λ,���,����,���,��������,ҽ������,˵��," & _
                " Decode(Nvl(�Ƿ���,0),1,Decode(Instr('567',���ID),0,Sum(Nvl(ԭ��,0))||'-'||Sum(Nvl(�ּ�,0)),'ʱ��'),Sum(�ּ�)) as �۸�," & _
                " Sum(ԭ��) as ԭ��ID,Sum(�ּ�) as �ּ�ID,Sum(ȱʡ�۸�) as ȱʡ�۸�ID,�Ƿ��� as �Ƿ���ID,���ID,��������ID" & _
                " From (" & _
                " Select Distinct 1 as ĩ��,A.ID,Decode(Instr('567',A.���),0,A.����ID,-E.����ID) as �ϼ�ID,A.����,A.����," & _
                " A.���㵥λ as ��λ,A.���,A.����,C.���� as ���,A.��������,N.���� as ҽ������,A.˵��,B.ԭ��,B.�ּ�,B.ȱʡ�۸�,A.�Ƿ���," & _
                " A.��� as ���ID,-Null as ��������ID" & _
                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ��� C,ҩƷ��� D,������ĿĿ¼ E,����֧����Ŀ M,����֧������ N" & _
                " Where A.ID=B.�շ�ϸĿID [ѡ���滻�Ĺ�����1]  And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "4", "5", "6") & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                " And A.������� IN(2,3)" & IIF(str��ĿIDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                " And A.��� Not IN('4','J','1') And A.���=C.���� And A.ID=D.ҩƷID(+) And D.ҩ��ID=E.ID(+)" & _
                " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[2]" & _
                " And (Nvl(a.ִ�п���,0) <> 4 Or Exists (Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid = a.Id And (w.������Դ=2 or (w.������Դ is Null And Nvl(w.��������id,[3]) = [3]))))" & _
                " And (a.��� Not in ('5','6','7') Or Exists(Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid=a.Id And Nvl(w.��������id,[3])=[3]))"
            If DeptExist("���ϲ���", 2) Then
                strSQL2 = strSQL2 & " Union ALL " & _
                    " Select Distinct 1 as ĩ��,A.ID,-E.����ID as �ϼ�ID,A.����,A.����," & _
                    " A.���㵥λ as ��λ,A.���,A.����,C.���� as ���,A.��������,N.���� as ҽ������,A.˵��," & _
                    " B.ԭ��,B.�ּ�,B.ȱʡ�۸�,A.�Ƿ���,A.��� as ���ID,D.�������� as ��������ID" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ��� C,�������� D,������ĿĿ¼ E,����֧����Ŀ M,����֧������ N" & _
                    " Where A.ID=B.�շ�ϸĿID [ѡ���滻�Ĺ�����2]  And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "4", "5", "6") & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And A.������� IN(2,3)" & IIF(str��ĿIDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                    " And A.���='4' And A.���=C.���� And A.ID=D.����ID And D.����ID=E.ID And D.�������=0" & _
                    " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[2]" & _
                    " And Exists(Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid=a.Id And Nvl(w.��������id,[3])=[3])"
            End If
            strSQL2 = strSQL2 & " ) Group by ĩ��,ID,�ϼ�ID,���,����,����,��λ,���,����,��������,ҽ������,˵��,�Ƿ���,���ID,��������ID"
            '[ѡ���滻�Ĺ�����1],[ѡ���滻�Ĺ�����2],����������ѡ���д����
            'Ҫȷ�� "ռλ����" �����һλ���ò�����ѡ������ƴ�ӣ�Ҫ���4000���ȵ�����
            Set rsTmp = ShowSQLSelectCIS(Me, strSQL, strSQL2, 2, "�շ���Ŀ", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, "," & str��ĿIDs & ",", Val(vsAdvice.TextMatrix(lng�к�, COL_����)), Val(vsAdvice.TextMatrix(lng�к�, COL_���˿���ID)), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "ռλ����")
            If Not rsTmp Is Nothing Then
                '�Ǳ���ִ�е�ҽ����������������Ŀ
                If lng�к� <> 0 Then
                    If NVL(rsTmp!�Ƿ���ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!���ID) > 0 Or rsTmp!���ID = "4" And NVL(rsTmp!��������ID, 0) = 1) Then
                        If Not Check����ִ��(Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))) Then
                            MsgBox "��ҽ���Ǳ���ִ�У�������Ա����Ŀ""" & rsTmp!���� & """���ۡ��üƼ���Ŀ��Ҫ�ֹ��Ƽۡ�", vbInformation, gstrSysName
                            .SetFocus: Exit Sub
                        End If
                    End If
                End If
                
                'ҽ��������
                If CheckItemInsure(rsTmp, lng�к�) Then
                    .SetFocus: Exit Sub
                End If
                
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                Call SetItemInput(Row, rsTmp, lngҽ��ID, int��������, lngԭ��ĿID)
                If lng�к� <> 0 Then
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCellPrice(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "û�п��õ��շ���Ŀ�����ȵ��շ���Ŀ���������ã�", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        ElseIf Col = COLP_ִ�п��� Then
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            If .TextMatrix(Row, COLP_�շ����) = "4" Then
                '�������õ�����
                strSQL = _
                    " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                    " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                    " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
                    " And B.������� IN(2,3) And B.����ID=C.ID" & _
                    " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " And (A.������Դ is NULL Or A.������Դ=2)" & _
                    " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                    " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                    " And A.�շ�ϸĿID=[1]" & _
                    " Order by B.�������,C.����"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ϲ���", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)))
            ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_�շ����)) > 0 Then
                'ҩƷ
                'ҩƷ��ϵͳָ���Ĵ���ҩ������
                If Not Check�ϰల��(True) Then
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                        " And B.������� IN(2,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (A.������Դ is NULL Or A.������Դ=2)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And A.�շ�ϸĿID=[1]" & _
                        " Order by B.�������,C.����"
                Else
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                        " And B.������� IN(2,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And D.����ID=C.ID And D.����=To_Number(To_Char(Sysdate,'D'))-1" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                        " And (A.������Դ is NULL Or A.������Դ=2)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And A.�շ�ϸĿID=[1]" & _
                        " Order by B.�������,C.����"
                End If
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩ��", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), _
                    decode(.TextMatrix(Row, COLP_�շ����), "5", "��ҩ��", "6", "��ҩ��", "7", "��ҩ��"))
            End If
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, COLP_ִ�п���ID) = rsTmp!ID
                .TextMatrix(Row, Col) = rsTmp!����
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!ִ�п���ID = rsTmp!ID
                    mrsPrice.Update
                    
                    '��ǼƼ����ݱ仯
                    vsAdvice.Cell(flexcpData, lng�к�, COL_���) = 1
                    
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCellPrice(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ����õĿ��ҡ�", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        End If
    End With
End Sub

Private Function CheckItemInsure(rsInput As ADODB.Recordset, ByVal lngRow As Long) As Boolean
'���ܣ��������(ѡ��)�Ƽ���Ŀ�Ƿ�ҽ������
'���أ����δ���룬������ʾѡ�񲻼������򷵻��档
    
    If gintҽ������ = 0 Then Exit Function
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_����)) <> 0 Then
            If Not ItemExistInsure(Val(.TextMatrix(lngRow, COL_����ID)), rsInput!ID, Val(.TextMatrix(lngRow, COL_����))) Then
                If gintҽ������ = 1 Then
                    If MsgBox("��Ŀ""" & rsInput!���� & """û�����ö�Ӧ�ı�����Ŀ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        CheckItemInsure = True
                    End If
                ElseIf gintҽ������ = 2 Then
                    MsgBox "��Ŀ""" & rsInput!���� & """û�����ö�Ӧ�ı�����Ŀ��", vbInformation, gstrSysName
                    CheckItemInsure = True
                End If
            End If
        End If
    End With
End Function

Private Sub SetItemInput(lngRow As Long, rsInput As ADODB.Recordset, ByVal lngҽ��ID As Long, ByVal int�������� As Integer, ByVal lngԭ��ĿID As Long)
    Dim lngִ�п���ID As Long, lng���˿���ID As Long
    Dim lng�к� As Long, dbl���� As Double
    Dim blnHaveSub As Boolean
    
    With vsPrice
        '��¼������
        '�������:����ʱ��ʾ�����������Ŀ,Ҳ���Դ���Ϊδ���Ƽ�ҽ��������������Ŀ
        .TextMatrix(lngRow, COLP_���) = rsInput!���
        .TextMatrix(lngRow, COLP_�շ����) = rsInput!���ID
        .TextMatrix(lngRow, COLP_�շ�ϸĿID) = rsInput!ID
        .TextMatrix(lngRow, COLP_�շ���Ŀ) = rsInput!����
        If Not IsNull(rsInput!����) Then
            .TextMatrix(lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ) & "(" & rsInput!���� & ")"
        End If
        If Not IsNull(rsInput!���) Then
            .TextMatrix(lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ) & " " & rsInput!���
        End If
        .TextMatrix(lngRow, COLP_��λ) = NVL(rsInput!��λ) '�������۵�λ(������ҩ��ҩƷ�Ƽ�)
        .TextMatrix(lngRow, COLP_�Ƽ�����) = 1 'ȱʡ��ԼƼ�1,ҩƷΪ��1�����۵�λ
        
        'ִ�п���
        lng�к� = Val(.TextMatrix(lngRow, COLP_�к�))
        If lng�к� <> 0 Then
            lngִ�п���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))
            If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(lng�к�, COL_����ID)), Val(vsAdvice.TextMatrix(lng�к�, COL_��ҳID)), "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
            '��ҩ��ҩƷ�͸������õ�����ר����ִ�п���
            If rsInput!���ID = "4" And NVL(rsInput!��������ID, 0) = 1 Or InStr(",5,6,7,", rsInput!���ID) > 0 Then
                lng���˿���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_���˿���ID))
                lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsInput!���ID, rsInput!ID, 4, lng���˿���ID, 0, 2, lngִ�п���ID, , , 2)
            End If
        End If
        .TextMatrix(lngRow, COLP_ִ�п���) = sys.RowValue("���ű�", lngִ�п���ID, "����")
        .TextMatrix(lngRow, COLP_ִ�п���ID) = lngִ�п���ID
        
        '���ۼ��㴦��:ҩ����ҩƷ�Ƽ۲����������ﴦ��
        If InStr(",5,6,7,", rsInput!���ID) > 0 Then
            If NVL(rsInput!�Ƿ���ID, 0) = 0 Then
                dbl���� = NVL(rsInput!�ּ�ID, 0)
            ElseIf lng�к� <> 0 Then
                '��ÿ��ȱʡһ�����۵�λ,��ǰ�������μ���
                dbl���� = CalcDrugPrice(rsInput!ID, lngִ�п���ID, Val(vsAdvice.TextMatrix(lng�к�, COL_����)), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
            End If
            .TextMatrix(lngRow, COLP_����) = Format(dbl����, gstrDecPrice)
                        
            'ʱ��ҩƷ������۸�
            .TextMatrix(lngRow, COLP_���) = 0
            .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = 0
            .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = 0
        ElseIf rsInput!���ID = "4" And NVL(rsInput!��������ID, 0) = 1 And NVL(rsInput!�Ƿ���ID, 0) = 1 Then
            '�������õ�ʱ�����ĺ�ҩƷһ������
            dbl���� = 0
            If lng�к� <> 0 Then
                dbl���� = CalcDrugPrice(rsInput!ID, lngִ�п���ID, Val(vsAdvice.TextMatrix(lng�к�, COL_����)), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
            End If
            .TextMatrix(lngRow, COLP_���) = 0
            .TextMatrix(lngRow, COLP_����) = Format(dbl����, gstrDecPrice)
            .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = 0
            .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = 0
        Else
            If NVL(rsInput!�Ƿ���ID, 0) = 0 Then
                .TextMatrix(lngRow, COLP_���) = 0
                .TextMatrix(lngRow, COLP_����) = Format(NVL(rsInput!�ּ�ID, 0), gstrDecPrice)
                .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = 0
                .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = 0
            Else
                .TextMatrix(lngRow, COLP_���) = 1
                .TextMatrix(lngRow, COLP_����) = Format(NVL(rsInput!ȱʡ�۸�ID), gstrDecPrice)
                .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = NVL(rsInput!ԭ��ID, 0)
                .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = NVL(rsInput!�ּ�ID, 0)
            End If
        End If
        
        .TextMatrix(lngRow, COLP_��������) = NVL(rsInput!��������)
        .TextMatrix(lngRow, COLP_�̶�) = 0
        
        '��������ָ�
        .Cell(flexcpData, lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ)
        .Cell(flexcpData, lngRow, COLP_�Ƽ�����) = .TextMatrix(lngRow, COLP_�Ƽ�����)
        .Cell(flexcpData, lngRow, COLP_����) = .TextMatrix(lngRow, COLP_����)
        .Cell(flexcpData, lngRow, COLP_ִ�п���) = .TextMatrix(lngRow, COLP_ִ�п���)
        
        '��¼������
        If lngҽ��ID <> 0 Then
            If lngԭ��ĿID = 0 Then
                '��ǰҽ���Ƿ��д��������������Ŀ�Ƿ����
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And ����=1"
                If Not mrsPrice.EOF Then blnHaveSub = True
                .TextMatrix(lngRow, COLP_����) = IIF(blnHaveSub, "��", "")
            
                mrsPrice.AddNew '����
            Else '����
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
            End If
            If lngԭ��ĿID = 0 Then
                mrsPrice!ҽ��ID = lngҽ��ID
                lng�к� = Val(.TextMatrix(lngRow, COLP_�к�))
                If Val(vsAdvice.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                    mrsPrice!���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_���ID))
                Else
                    mrsPrice!���ID = Null
                End If
                mrsPrice!�������� = int��������
                mrsPrice!���� = IIF(blnHaveSub, 1, 0)
            End If
            mrsPrice!�շѷ�ʽ = 0
            mrsPrice!�շ���� = rsInput!���ID
            mrsPrice!�շ�ϸĿID = rsInput!ID
            If lngִ�п���ID <> 0 Then
                mrsPrice!ִ�п���ID = lngִ�п���ID
            Else
                mrsPrice!ִ�п���ID = Null
            End If
            mrsPrice!���� = NVL(rsInput!��������ID, 0)
            mrsPrice!��� = NVL(rsInput!�Ƿ���ID, 0)
            mrsPrice!���� = Val(.TextMatrix(lngRow, COLP_����))
            mrsPrice!���� = 1
            mrsPrice!�̶� = 0
            mrsPrice.Update
            
            '��ǼƼ����ݱ仯
            vsAdvice.Cell(flexcpData, lng�к�, COL_���) = 1
        End If
    End With
End Sub

Private Sub vsPrice_DblClick()
    Call vsPrice_KeyPress(32)
End Sub

Private Sub vsPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsPrice
        If KeyCode = vbKeyF4 Then
            If CellEditablePrice(.Row, .Col) And .Col = COLP_�Ƽ�ҽ�� Then
                Call zlcommfun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Editable And Val(.TextMatrix(.Row, COLP_�̶�)) = 0 Then
                If Val(.TextMatrix(.Row, COLP_�к�)) <> 0 And Val(.TextMatrix(.Row, COLP_�շ�ϸĿID)) <> 0 Then
                    'ҽ������д�������Ҫ����һ��(�����ǹ̶����ɶ���)
                    mrsPrice.Filter = "ҽ��ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_�к�)), COL_ID)) & _
                        " And ��������=" & Val(.TextMatrix(.Row, COLP_��������)) & " And ����=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(.Row, COLP_����) <> "" Then
                        MsgBox """" & .Cell(flexcpData, .Row, COLP_�Ƽ�ҽ��) & """����Ҫ����һ�������Ƽ���Ŀ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                
                    If MsgBox("ȷʵҪɾ����ǰ�Ƽ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    mrsPrice.Filter = "ҽ��ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_�к�)), COL_ID)) & _
                        " And ��������=" & Val(.TextMatrix(.Row, COLP_��������)) & " And �շ�ϸĿID=" & Val(.TextMatrix(.Row, COLP_�շ�ϸĿID))
                    mrsPrice.Delete
                End If
                
                .RemoveItem .Row
                If .Rows = .FixedRows Then
                    .Rows = .FixedRows + 1
                    .Row = .FixedRows: .Col = COLP_�Ƽ�ҽ��
                End If
                
                Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsPrice_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsPrice_KeyPress(KeyAscii As Integer)
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCellPrice(.Row, .Col)
        Else
            If CellEditablePrice(.Row, .Col) And (.Col = COLP_�շ���Ŀ Or .Col = COLP_ִ�п���) Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsPrice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng�к� As Long, i As Long
    Dim str��ĿIDs As String, int�������� As Integer
    Dim lngҽ��ID As Long, lngԭ��ĿID As Long
    Dim strTmp As String, blnCancel As Boolean
    Dim strInput As String, strMatch As String
    Dim vPoint As PointAPI
    Dim lng���˿���ID As Long
    Dim lng��ҩ�� As Long
    Dim lng��ҩ�� As Long
    Dim lng��ҩ�� As Long
    Dim lng���ϲ��� As Long
    Dim strStock As String
    
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            lng�к� = Val(.TextMatrix(Row, COLP_�к�))
            If Col = COLP_�Ƽ�ҽ�� Then
                '����ʱ�س�
                If .ComboIndex <> -1 Then
                    .TextMatrix(.Row, .Col) = .ComboItem(.ComboIndex) '��ȻEnterNextCellPrice����Ҫ�˳�
                    Call EnterNextCellPrice(Row, Col)
                End If
            ElseIf Col = COLP_�Ƽ����� Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "�Ƽ�����������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!���� = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    
                    '��ǼƼ����ݱ仯
                    vsAdvice.Cell(flexcpData, lng�к�, COL_���) = 1

                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCellPrice(Row, Col)
            ElseIf Col = COLP_���� Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "�շѵ���������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                '��������뷶Χ
                strTmp = CheckScope(.Cell(flexcpData, Row, COLP_Ӧ�ս��), .Cell(flexcpData, Row, COLP_ʵ�ս��), .EditText)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .EditText = Format(.EditText, gstrDecPrice)
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!���� = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    
                    '��ǼƼ����ݱ仯
                    vsAdvice.Cell(flexcpData, lng�к�, COL_���) = 1
                    
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCellPrice(Row, Col)
            ElseIf Col = COLP_�շ���Ŀ And .EditText <> "" Then
                '����ѡ�����е���Ŀ
                For i = .FixedRows To .Rows - 1
                    If Val(vsAdvice.TextMatrix(Val(.TextMatrix(i, COLP_�к�)), COL_ID)) = Val(vsAdvice.TextMatrix(lng�к�, COL_ID)) _
                        And Val(vsAdvice.TextMatrix(lng�к�, COL_ID)) <> 0 And i <> Row Then
                        str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COLP_�շ�ϸĿID))
                    End If
                Next
                str��ĿIDs = Mid(str��ĿIDs, 2)
                
                lng���˿���ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID))
                lng��ҩ�� = Val(zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�, , , , , lng���˿���ID))
                lng��ҩ�� = Val(zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�, , , , , lng���˿���ID))
                lng��ҩ�� = Val(zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�, , , , , lng���˿���ID))
                lng���ϲ��� = Val(zlDatabase.GetPara("סԺȱʡ���ϲ���", glngSys, pסԺҽ���´�, , , , , lng���˿���ID))
                
                If lng��ҩ�� <> 0 Or lng��ҩ�� <> 0 Or lng��ҩ�� <> 0 Or lng���ϲ��� <> 0 Then
                    strStock = _
                        "Select A.ҩƷID,Sum(Nvl(A.��������,0)) as ���" & _
                        " From ҩƷ��� A,�շ���ĿĿ¼ B" & _
                        " Where A.���� = 1 And (Nvl(A.����,0)=0 Or A.Ч�� Is Null Or A.Ч��>Trunc(Sysdate))" & _
                        " And A.�ⷿID=Decode(B.���,'5',[7],'6',[8],'7',[9],'4',[10],Null)" & _
                        " And A.ҩƷID=B.ID And B.��� IN('4','5','6','7')" & _
                        " Group by A.ҩƷID Having Sum(Nvl(A.��������,0))<>0"
                Else
                    strStock = "Select Null as ҩƷID,Null as ��� From Dual"
                End If
                
                '��ͬ������ƥ�䷽ʽ
                strInput = UCase(.EditText)
                strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=[3] Or C.���� Like [2] And C.���� IN([3],3))"
                If IsNumeric(strInput) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=3)"
                ElseIf zlcommfun.IsCharAlpha(strInput) Then         '01,11.����ȫ����ĸʱֻƥ�����
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.���� Like [2] And C.����=[3]"
                ElseIf zlcommfun.IsCharChinese(strInput) Then
                    strMatch = " And C.���� Like [2] And C.����=[3]"
                End If
                If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(lng�к�, COL_����ID)), Val(vsAdvice.TextMatrix(lng�к�, COL_��ҳID)), "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                
                strSQL = ""
                If Not DeptExist("���ϲ���", 2) Then strSQL = " And A.���<>'4'"
                strSQL = "Select * From (" & _
                    " Select A.ĩ��,A.ID,A.���,A.����,A.����,A.��λ,A.���,A.����," & _
                    " Decode(Instr('4567',A.���ID),0,NULL,1," & _
                    "   Decode(S.���,NULL,NULL,LTrim(To_Char(S.���,'999990.0000'))||A.��λ)," & _
                    "   Decode(S.���,NULL,NULL,LTrim(To_Char(S.���/Nvl(C.סԺ��װ,1),'999990.0000'))||C.סԺ��λ)) as ���," & _
                    "   A.��������,N.���� as ҽ������,A.˵��," & _
                    " Decode(Nvl(A.�Ƿ���,0),1,Decode(Instr('567',A.���ID),0,Sum(Nvl(A.ԭ��,0))||'-'||Sum(Nvl(A.�ּ�,0)),'ʱ��'),Sum(A.�ּ�)) as �۸�," & _
                    " Sum(A.ԭ��) as ԭ��ID,Sum(A.�ּ�) as �ּ�ID,Sum(A.ȱʡ�۸�) as ȱʡ�۸�ID,A.�Ƿ��� as �Ƿ���ID,A.���ID,B.�������� as ��������ID,B.�������" & _
                    " From (" & _
                    " Select Distinct 1 as ĩ��,A.ID,a.ִ�п���,A.��� as ���ID,D.���� as ���,A.����,A.����,A.���㵥λ as ��λ," & _
                    " A.���,A.����,A.��������,A.˵��,B.ԭ��,B.�ּ�,B.ȱʡ�۸�,A.�Ƿ���" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ���� C,�շ���Ŀ��� D" & _
                    " Where A.ID=B.�շ�ϸĿID And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "11", "12", "13") & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And A.������� IN(2,3)" & IIF(str��ĿIDs <> "", " And Instr([4],','||A.ID||',')=0", "") & _
                    " And A.ID=C.�շ�ϸĿID And A.���=D.���� And A.��� Not IN('J','1')" & strSQL & strMatch & _
                    " ) A,�������� B,ҩƷ��� C,����֧����Ŀ M,����֧������ N,(" & strStock & ") S" & _
                    " Where A.ID=B.����ID(+) And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[5]  And A.ID=C.ҩƷID(+) And A.ID=S.ҩƷID(+)" & _
                    " And (Nvl(a.ִ�п���,0) <> 4 Or Exists (Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid = a.Id And (w.������Դ=2 or (w.������Դ is Null And Nvl(w.��������id,[6]) = [6]))))" & _
                    " And (a.���id not in ('4','5','6','7') Or Exists(Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid=a.Id And Nvl(w.��������id,[6])=[6]))" & _
                    " Group by A.ĩ��,A.ID,A.���,A.����,A.����,A.��λ,A.���,A.����,A.��������,C.סԺ��λ,C.סԺ��װ,S.���,N.����,A.˵��,A.�Ƿ���,A.���ID,B.��������,B.�������" & _
                    " ) Where Nvl(�������,0) = 0 Order by ���, ����"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�շ���Ŀ", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", mstrLike & strInput & "%", mint���� + 1, "," & str��ĿIDs & ",", Val(vsAdvice.TextMatrix(lng�к�, COL_����)), lng���˿���ID, lng��ҩ��, lng��ҩ��, lng��ҩ��, lng���ϲ���, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                If Not rsTmp Is Nothing Then
                    '�Ǳ���ִ�е�ҽ����������������Ŀ
                    If lng�к� <> 0 Then
                        If NVL(rsTmp!�Ƿ���ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!���ID) > 0 Or rsTmp!���ID = "4" And NVL(rsTmp!��������ID, 0) = 1) Then
                            If Not Check����ִ��(Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))) Then
                                MsgBox "��ҽ���Ǳ���ִ�У�������Ա����Ŀ""" & rsTmp!���� & """���ۡ��üƼ���Ŀ��Ҫ�ֹ��Ƽۡ�", vbInformation, gstrSysName
                                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                                Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                                .SetFocus: Exit Sub
                            End If
                        End If
                    End If
                
                    'ҽ��������
                    If CheckItemInsure(rsTmp, lng�к�) Then
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                        .SetFocus: Exit Sub
                    End If
                
                    lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                    int�������� = Val(.TextMatrix(Row, COLP_��������))
                    lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                    Call SetItemInput(Row, rsTmp, lngҽ��ID, int��������, lngԭ��ĿID)
                    .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                    If lng�к� <> 0 Then
                        Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                    End If
                    Call EnterNextCellPrice(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "û���ҵ����õ��շ���Ŀ��", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    .SetFocus
                End If
            ElseIf Col = COLP_ִ�п��� And .EditText <> "" Then 'ִ�п���
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                If .TextMatrix(Row, COLP_�շ����) = "4" Then
                    '�������õ�����
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
                        " And B.������� IN(2,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (A.������Դ is NULL Or A.������Դ=2)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And A.�շ�ϸĿID=[1] And (C.���� Like [3] Or C.���� Like [4] Or C.���� Like [4])" & _
                        " Order by B.�������,C.����"
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ϲ���", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_�շ����)) > 0 Then
                    'ҩƷ��ϵͳָ���Ĵ���ҩ������
                    If Not Check�ϰల��(True) Then
                        strSQL = _
                            " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                            " And B.������� IN(2,3) And B.����ID=C.ID" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " And (A.������Դ is NULL Or A.������Դ=2)" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                            " And A.�շ�ϸĿID=[1] And (C.���� Like [4] Or C.���� Like [5] Or C.���� Like [5])" & _
                            " Order by B.�������,C.����"
                    Else
                        strSQL = _
                            " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                            " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                            " And B.������� IN(2,3) And B.����ID=C.ID" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " And D.����ID=C.ID And D.����=To_Number(To_Char(Sysdate,'D'))-1" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                            " And (A.������Դ is NULL Or A.������Դ=2)" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                            " And A.�շ�ϸĿID=[1] And (C.���� Like [4] Or C.���� Like [5] Or C.���� Like [5])" & _
                            " Order by B.�������,C.����"
                    End If
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩ��", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), _
                        decode(.TextMatrix(Row, COLP_�շ����), "5", "��ҩ��", "6", "��ҩ��", "7", "��ҩ��"), _
                        UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                End If
                If Not rsTmp Is Nothing Then
                    .TextMatrix(Row, COLP_ִ�п���ID) = rsTmp!ID
                    .TextMatrix(Row, Col) = rsTmp!����
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                    
                    '���¼�¼��
                    lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                    int�������� = Val(.TextMatrix(Row, COLP_��������))
                    lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                    If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                        mrsPrice!ִ�п���ID = rsTmp!ID
                        mrsPrice.Update
                        
                        '��ǼƼ����ݱ仯
                        vsAdvice.Cell(flexcpData, lng�к�, COL_���) = 1
                        
                        Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                    End If
                    Call EnterNextCellPrice(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "û���ҵ����õĿ��ҡ�", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    .SetFocus
                End If
            End If
        Else
            If Col = COLP_�Ƽ����� Or Col = COLP_���� Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'˵�������ص��кŷ�Χ��������ҩ;�����к�
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub InitAdviceTable()
'���ܣ���ʼ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = ",300,4;����,850,1;����,750,1;סԺ��,750,1;����,500,4;�ѱ�,750,1;Ӥ��,550,1;��Ч,550,1;ҽ������,2000,1;���,2000,1;����,600,7;��λ,450,1;" & _
        "����,600,7;��λ,450,1;���,850,7;Ƶ��,1000,1;�÷�,1000,1;ҽ������,1500,1;ִ��ʱ��,1000,1;�״�ʱ��,1530,1;ĩ��ʱ��,1530,1;ִ�п���,850,1;" & _
        "����ִ��,850,1;ִ������,850,1;����ID;��ҳID;�Ա�;����;����;ID;���ID;���˲���ID;���˿���ID;��������ID;����ҽ��;�������;������ĿID;�Ƽ�����;" & _
        "ִ������ID;ִ�п���ID;ִ�б��;ҩƷID;����ϵ��;סԺ��װ;סԺ��λ;�ɷ����;ҩ������;�Ƿ���;���;����;�ֽ�ʱ��;��������;�Թܱ���;�걾��λ;��鷽��;��������;" & _
        "������־;ҽ��״̬;ִ��Ƶ��;�¿�����ʱ��;���㷽ʽ;ִ�а���;��������;����ʱ��;��ʼʱ��;�������;ִ�з���;�������;��ҩ����;����ҽ��ID"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        .RowHeight(0) = 320
    End With
End Sub

Private Sub InitPriceTable()
'���ܣ���ʼ���Ƽ��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "�к�;�շ�ϸĿID;�̶�;���;�Ƽ�ҽ��,2000,1;���,650,1;�շ���Ŀ,2000,1;�Ƽ�����,900,7;" & _
        "����,450,4;����,800,7;��λ,500,1;����,1000,7;Ӧ�ս��,1050,7;ʵ�ս��,1050,7;ִ�п���,1000,1;��������,850,1;" & _
        "����,450,4;�շѷ�ʽ,1500,1;�շ����;ִ�п���ID;��������;��������"
    arrHead = Split(strHead, ";")
    With vsPrice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Function Decodeҩ���û�() As String
'���ܣ�����ҩ���û�����,�����ֶ�Decode���
'˵����SQL��ѯ��,"����ҽ����¼"����Ϊ"A"
    Dim i As Long, strTmp As String
    
    mblnҩ���û� = False
    If chkDeptExchange.value = 1 And InStr(GetInsidePrivs(pסԺҽ������), ";�����û�ҩ��;") > 0 Then
        mrsҩ��.Filter = 0
        If Not mrsҩ��.EOF Then
            For i = 1 To mrsҩ��.RecordCount
                If mrsҩ��!ID <> mrsҩ��!��ID Then
                    strTmp = strTmp & "," & mrsҩ��!ID & "," & mrsҩ��!��ID
                End If
                mrsҩ��.MoveNext
            Next
        End If
    End If
    If strTmp <> "" Then
        Decodeҩ���û� = "Decode(A.ִ�п���ID" & strTmp & ",A.ִ�п���ID)"
        mblnҩ���û� = True
    Else
        Decodeҩ���û� = "A.ִ�п���ID"
    End If
End Function

Private Sub DeleteCurRow(ByVal lngRow As Long, ByVal lng���ID As Long)
'���ܣ��ڴ���������嵥�Ĺ�����ɾ������������
    Dim i As Long
    With vsAdvice
        'ɾ����ǰ��
        .RemoveItem lngRow
        
        'ɾ���䷽��һ����ҩ���Ѿ��������
        If lng���ID <> 0 Then
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = lng���ID Then
                    .RemoveItem i
                End If
            Next
        End If
    End With
End Sub

Private Function CheckWaitExecute(rsPati As ADODB.Recordset, ByVal lngRow As Long, ByVal byt��Ŀ��鷽ʽ As Byte, ByVal bytҩƷ��鷽ʽ As Byte) As Boolean
'���ܣ�����ָ���ļ�鷽ʽ���Բ���δִ�е���Ȧ��δ��ҩƷ���м��
'������byt��鷽ʽ=0-�����,1-��鲢��ʾ,2-��鲢��ֹ
'���أ��Ƿ����
    Dim strTmp As String
        
    With vsAdvice
        If byt��Ŀ��鷽ʽ <> 0 Then
            strTmp = ExistWaitExe(rsPati!����ID, Val(.TextMatrix(lngRow, COL_��ҳID)), -1)
            If strTmp <> "" Then
                Call .ShowCell(lngRow, col_ҽ������): .Refresh
                If byt��Ŀ��鷽ʽ = 1 Then
                    If MsgBox("���ֲ���""" & rsPati!���� & """������δִ����ɵ����ݣ�" & _
                        vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "ȷʵҪ����""" & .TextMatrix(lngRow, col_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "���ֲ���""" & rsPati!���� & """������δִ����ɵ����ݣ�" & _
                        vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """���������͡�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        If bytҩƷ��鷽ʽ <> 0 Then
            strTmp = ExistWaitDrug(rsPati!����ID, Val(.TextMatrix(lngRow, COL_��ҳID)), -1)
            If strTmp <> "" Then
                Call .ShowCell(lngRow, col_ҽ������): .Refresh
                If bytҩƷ��鷽ʽ = 1 Then
                    If MsgBox("���ֲ���""" & rsPati!���� & """" & _
                        strTmp & vbCrLf & vbCrLf & "ȷʵҪ����""" & .TextMatrix(lngRow, col_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "���ֲ���""" & rsPati!���� & """" & _
                        strTmp & vbCrLf & vbCrLf & "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """����������", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End With
    
    CheckWaitExecute = True
End Function

Private Function CheckStock(ByVal lngRow As Long, Optional bln�����ʾ As Boolean, Optional blnʱ����ʾ As Boolean, Optional blnĬ�Ϸ��� As Boolean, Optional ByVal blnCurPati As Boolean) As Boolean
'���ܣ����ݿ���������鷢��ҩƷ�Ŀ��
'������lngRow=ҽ���к�
'      blnCurPati=�Ƿ�ֻ�Ե�ǰ���˽��л��ܼ��,���ڷ��͹�����,��Ϊ�ǰ������ύ,��ʱ������ȡ�Ŀ����׼ȷ��
'      bln�����ʾ,blnʱ����ʾ,blnĬ�Ϸ���=������ʾ�������ʾ����
'���أ�������ʾ���Ƿ��ѡ��״̬�����˴���
    Dim int����� As Integer, dbl���� As Double
    Dim dbl���ÿ�� As Double, dbl�ѷ���� As Double
    Dim bln����ʱ�� As Boolean, bln���� As Boolean, blnʱ�� As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        'ҩƷ�����(0-�����;1-���,��������;2-��飬�����ֹ)
        int����� = TheStockCheck(Val(.TextMatrix(lngRow, COL_ִ�п���ID)), .TextMatrix(lngRow, COL_�������))
        bln���� = Val(.TextMatrix(lngRow, COL_ҩ������)) = 1
        blnʱ�� = Val(.TextMatrix(lngRow, COL_�Ƿ���)) = 1
        
        '������ʱ��ҩƷ����Ҫ���㹻�Ŀ��,�������ݿ�����������
        If int����� <> 0 Or bln���� Or blnʱ�� Then
            strTmp = .TextMatrix(lngRow, COL_סԺ��λ) '������ʾ
            
            '������Ͳ����ֹʱ,����ʱ��Ͳ��ص�������
            bln����ʱ�� = int����� <> 2 And (bln���� Or blnʱ��)
            
            '��ǰҩƷ����:סԺ��װ
            If .TextMatrix(lngRow, COL_�������) = "7" Then
                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                If Val(.TextMatrix(lngRow, COL_�ɷ����)) = 0 Then
                    dbl���� = Val(.TextMatrix(lngRow, COL_����)) * Val(.TextMatrix(lngRow, COL_����))
                    dbl���� = dbl���� / Val(.TextMatrix(lngRow, COL_����ϵ��)) / Val(.TextMatrix(lngRow, COL_סԺ��װ))
                Else
                    dbl���� = IntEx(Val(.TextMatrix(lngRow, COL_����)) / Val(.TextMatrix(lngRow, COL_����ϵ��)) / Val(.TextMatrix(lngRow, COL_סԺ��װ)))
                    dbl���� = dbl���� * Val(.TextMatrix(lngRow, COL_����))
                End If
            Else
                dbl���� = Val(.TextMatrix(lngRow, COL_����))
            End If
            
            '��ǰ���ÿ��:סԺ��װ,��ȥǰ����ͬҩƷҪ���͵Ŀ��
            For i = lngRow - 1 To .FixedRows Step -1
                If blnCurPati And Val(.TextMatrix(i, COL_����ID)) = Val(.TextMatrix(lngRow, COL_����ID)) Or Not blnCurPati Then
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0
                    If blnDo Then
                        blnDo = Val(.TextMatrix(i, COL_�շ�ϸĿID)) = Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) _
                            And Val(.TextMatrix(i, COL_ִ�п���ID)) = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
                    End If
                    If blnDo Then
                        blnDo = .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing
                    End If
                    If blnDo Then
                        If .TextMatrix(i, COL_�������) = "7" Then
                            '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                            If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                                dbl�ѷ���� = dbl�ѷ���� + _
                                    Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)) _
                                    / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ))
                            Else
                                dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����)) _
                                    * IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ)))
                            End If
                        Else
                            dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����))
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
            dbl���ÿ�� = Val(.TextMatrix(lngRow, COL_���))
            dbl���ÿ�� = dbl���ÿ�� - dbl�ѷ����
            
            If dbl���� > dbl���ÿ�� Then
                If (Not bln����ʱ�� And int����� <> 0 And bln�����ʾ) Or (bln����ʱ�� And blnʱ����ʾ) Then
                    '��һ��û��ѡ������ʾ,����ʾ
                    If bln����ʱ�� Then
                        If InStr(GetInsidePrivs(pסԺҽ������), "��ʾҩƷ���") = 0 Then
                            strTmp = "ҩ��������ʱ��ҩƷ""" & .TextMatrix(lngRow, COL_���) & """��" & vbCrLf & vbCrLf & _
                                "��" & .TextMatrix(lngRow, COL_ִ�п���) & "��治��" & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                                "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                        Else
                            strTmp = "ҩ��������ʱ��ҩƷ""" & .TextMatrix(lngRow, COL_���) & """��治�㣺" & vbCrLf & vbCrLf & _
                                .TextMatrix(lngRow, COL_ִ�п���) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                                "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                        End If
                    Else
                        If InStr(GetInsidePrivs(pסԺҽ������), "��ʾҩƷ���") = 0 Then
                            strTmp = "ҩƷ""" & .TextMatrix(lngRow, COL_���) & """��" & vbCrLf & vbCrLf & _
                                "��" & .TextMatrix(lngRow, COL_ִ�п���) & "��治��" & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                                "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                        Else
                            strTmp = "ҩƷ""" & .TextMatrix(lngRow, COL_���) & """��治�㣺" & vbCrLf & vbCrLf & _
                                .TextMatrix(lngRow, COL_ִ�п���) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                                "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                        End If
                    End If
                    If .Cell(flexcpData, lngRow, COL_���) <> "" Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "������ڷ����嵥��ѡ���ҩƷ�����㹻�����������"
                    End If
                    If int����� = 1 And Not bln����ʱ�� Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "Ҫ���͸�ҩƷ��"
                    End If
                    
                    strTmp = "����" & .TextMatrix(lngRow, COL_����) & "��" & vbCrLf & vbCrLf & strTmp
                    
                    .Redraw = flexRDDirect:
                    Call .ShowCell(lngRow, COL_ѡ��)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int����� = 2 Or bln����ʱ��)
                    
                    If bln����ʱ�� Then
                        If vMsg = vbIgnore Then blnʱ����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckStock = True
                    ElseIf int����� = 2 Then '����ֹ
                        If vMsg = vbIgnore Then bln�����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckStock = True
                    ElseIf int����� = 1 Then '�������
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln�����ʾ = False
                            blnĬ�Ϸ��� = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln�����ʾ = False
                            blnĬ�Ϸ��� = False
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckStock = True
                        End If
                    End If
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '��һ��ѡ���˲�����ʾ
                    If int����� = 2 Or bln���� Or blnʱ�� Then
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckStock = True
                    ElseIf int����� = 1 Then
                        '������һ�εĽ������
                        If Not blnĬ�Ϸ��� Then
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckStock = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function CheckDrugStorage(ByVal lngRow As Long, Optional bln�洢�ⷿ��ʾ As Boolean) As Boolean
'���ܣ����ݿ���������鷢��ҩƷ�Ĵ洢�ⷿ
'������lngRow=ҽ���к�
'      bln�洢�ⷿ��ʾ=�Ƿ������ʾ
'���أ�������ʾ���Ƿ��ѡ��״̬�����˴���
    Dim lngҩƷID As Long, lngִ�п���ID As Long
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim strTmp As String
    Dim vMsg As VbMsgBoxResult
    
    With vsAdvice
        '���������δ��ѡ���򲻼��
        If .Cell(flexcpData, lngRow, COL_ѡ��) = 1 Then Exit Function
        '�������û�ҩ���Ĳż��
        If mblnҩ���û� = False Then Exit Function
        '��ȡҩƷID
        lngҩƷID = Val(.TextMatrix(lngRow, COL_�շ�ϸĿID))
        If lngҩƷID = 0 Then Exit Function
        lngִ�п���ID = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
        If lngִ�п���ID = 0 Then Exit Function
        strSQL = "select 1 from �շ�ִ�п��� where �շ�ϸĿID = [1]  And Nvl(������Դ,2) = 2 And ִ�п���id = [2] And Nvl(��������id, [3]) = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDrugStorage", lngҩƷID, lngִ�п���ID, Val(.TextMatrix(lngRow, COL_��������ID)))
        
        If rsTmp.RecordCount > 0 Then Exit Function
        strTmp = "�ⷿ""" & .TextMatrix(lngRow, COL_ִ�п���) & """��û�д洢ҩƷ""" & .TextMatrix(lngRow, COL_���) & """"
        strTmp = "����" & .TextMatrix(lngRow, COL_����) & "��" & vbCrLf & vbCrLf & strTmp
        
        .Redraw = flexRDDirect:
        Call .ShowCell(lngRow, COL_ѡ��)
        Screen.MousePointer = 0
        '���˲�����ʾ
        If bln�洢�ⷿ��ʾ = True Then
            vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, True)
            If vMsg = vbIgnore Then bln�洢�ⷿ��ʾ = False
        End If
       
        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
        CheckDrugStorage = True
    
        Screen.MousePointer = 11
        .Refresh: .Redraw = flexRDNone
    End With
End Function

Private Function CheckPriceStock(ByVal lngRow As Long, rsPrice As ADODB.Recordset, ByVal lng�ⷿID As Long, ByVal dbl���� As Double, _
    rsTotal As ADODB.Recordset, Optional bln�����ʾ As Boolean, Optional blnʱ����ʾ As Boolean, Optional blnĬ�Ϸ��� As Boolean) As Boolean
'���ܣ����͹�����ʱ���Է�ҩ��ҩƷ���������õ����ļƼ۽��п����(�ۼƼ��)
'������lngRow=ҽ���к�
'      dbl����=�Ѽ���õļƼ�����(�ۼ۵�λ)
'      rsTotal=��ǰ����ǰ�����ۼƷ��͵ļƼ�ҩƷ����������(�ۼ۵�λ)
'      bln�����ʾ,blnʱ����ʾ,blnĬ�Ϸ���=������ʾ�������ʾ����
'���أ�������ʾ���Ƿ��ѡ��״̬�����˴���
    Dim int����� As Integer, dbl���� As Double
    Dim dbl���ÿ�� As Double, dbl�ѷ���� As Double
    Dim bln����ʱ�� As Boolean, bln���� As Boolean, blnʱ�� As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        'ҩƷ�����(0-�����;1-���,��������;2-��飬�����ֹ)
        int����� = TheStockCheck(lng�ⷿID, rsPrice!���)
        bln���� = NVL(rsPrice!����, 0) = 1
        blnʱ�� = NVL(rsPrice!�Ƿ���, 0) = 1
        
        '������ʱ��ҩƷ����Ҫ���㹻�Ŀ��,�������ݿ�����������
        If int����� <> 0 Or bln���� Or blnʱ�� Then
            strTmp = NVL(rsPrice!סԺ��λ, NVL(rsPrice!���㵥λ)) '������ʾ
            
            '������Ͳ����ֹʱ,����ʱ��Ͳ��ص�������
            bln����ʱ�� = int����� <> 2 And (bln���� Or blnʱ��)
            
            '��ǰҩƷ����������:סԺ��װ
            dbl���� = Format(dbl���� / NVL(rsPrice!סԺ��װ, 1), "0.00000")
            
            '��ǰ���ÿ��:סԺ��װ,��ȥǰ����ͬҩƷҽ��Ҫ���͵Ŀ��
            If InStr(",5,6,7,", rsPrice!���) > 0 Then
                For i = lngRow - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(i, COL_����ID)) = Val(.TextMatrix(lngRow, COL_����ID)) Then
                        blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0
                        If blnDo Then
                            blnDo = Val(.TextMatrix(i, COL_�շ�ϸĿID)) = rsPrice!ID And Val(.TextMatrix(i, COL_ִ�п���ID)) = lng�ⷿID
                        End If
                        If blnDo Then
                            blnDo = .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing
                        End If
                        If blnDo Then
                            If .TextMatrix(i, COL_�������) = "7" Then
                                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                                If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                                    dbl�ѷ���� = dbl�ѷ���� + _
                                        Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)) _
                                        / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ))
                                Else
                                    dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����)) _
                                        * IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ)))
                                End If
                            Else
                                dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����))
                            End If
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If
            '�Ƽ۲���Ҫ���͵��ۼ�����
            rsTotal.Filter = "��ĿID=" & rsPrice!ID & " And �ⷿID=" & lng�ⷿID
            Do While Not rsTotal.EOF
                dbl�ѷ���� = dbl�ѷ���� + Format(rsTotal!���� / NVL(rsPrice!סԺ��װ, 1), "0.00000")
                rsTotal.MoveNext
            Loop
            
            dbl���ÿ�� = Format(GetStock(rsPrice!ID, lng�ⷿID, 2), "0.00000")
            dbl���ÿ�� = dbl���ÿ�� - dbl�ѷ����
            
            If dbl���� > dbl���ÿ�� Then
                If (Not bln����ʱ�� And int����� <> 0 And bln�����ʾ) Or (bln����ʱ�� And blnʱ����ʾ) Then
                    '��һ��û��ѡ������ʾ,����ʾ
                    If bln����ʱ�� Then
                        If InStr(GetInsidePrivs(pסԺҽ������), "��ʾҩƷ���") = 0 Then
                            strTmp = "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """�ķ�����ʱ�ۼƼ���Ŀ��" & vbCrLf & vbCrLf & _
                                """" & rsPrice!���� & """��" & sys.RowValue("���ű�", lng�ⷿID, "����") & "��治��" & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                        Else
                            strTmp = "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """�ķ�����ʱ�ۼƼ���Ŀ""" & rsPrice!���� & """��治�㣺" & _
                                vbCrLf & vbCrLf & sys.RowValue("���ű�", lng�ⷿID, "����") & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                        End If
                    Else
                        If InStr(GetInsidePrivs(pסԺҽ������), "��ʾҩƷ���") = 0 Then
                            strTmp = "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """�ļƼ���Ŀ��" & vbCrLf & vbCrLf & _
                                """" & rsPrice!���� & """��" & sys.RowValue("���ű�", lng�ⷿID, "����") & "��治��" & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                        Else
                            strTmp = "ҽ��""" & .TextMatrix(lngRow, col_ҽ������) & """�ļƼ���Ŀ""" & rsPrice!���� & """��治�㣺" & _
                                vbCrLf & vbCrLf & sys.RowValue("���ű�", lng�ⷿID, "����") & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                                IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                        End If
                    End If
                    If int����� = 1 And Not bln����ʱ�� Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "Ҫ���͸�ҽ����"
                    End If
                    strTmp = "����" & .TextMatrix(lngRow, COL_����) & "��" & vbCrLf & vbCrLf & strTmp
                    
                    .Redraw = flexRDDirect
                    .Row = GetVisibleRow(lngRow, True)
                    Call .ShowCell(.Row, COL_ѡ��)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int����� = 2 Or bln����ʱ��)
                    
                    If bln����ʱ�� Then
                        If vMsg = vbIgnore Then blnʱ����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 2 Then '����ֹ
                        If vMsg = vbIgnore Then bln�����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 1 Then '�������
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln�����ʾ = False
                            blnĬ�Ϸ��� = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln�����ʾ = False
                            blnĬ�Ϸ��� = False
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckPriceStock = True
                        End If
                    End If
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '��һ��ѡ���˲�����ʾ
                    If int����� = 2 Or bln���� Or blnʱ�� Then
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 1 Then
                        '������һ�εĽ������
                        If Not blnĬ�Ϸ��� Then
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckPriceStock = True
                        End If
                    End If
                End If
            End If
        End If
        
        '���δ��ʾ��Ҫ����,�����ۼƷ�������
        If Not CheckPriceStock Then
            rsTotal.AddNew
            If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                rsTotal!ҽ��ID = Val(.TextMatrix(lngRow, COL_���ID))
            Else
                rsTotal!ҽ��ID = Val(.TextMatrix(lngRow, COL_ID))
            End If
            rsTotal!��ĿID = rsPrice!ID
            rsTotal!�ⷿID = lng�ⷿID
            rsTotal!���� = dbl����
            rsTotal.Update
        End If
    End With
End Function

Private Sub DeleteDrugRow(rsSend As ADODB.Recordset, ByVal lngRow As Long, lngDel���ID As Long)
'���ܣ�ɾ����Ӧ��ҩƷ��,����ҩƷͣ�û�������ԭ���Ҳ�����Ч���ʱ
'���أ�lngDel���ID-��Ҫͬʱɾ�����������ҽ����ʶ
    Dim strMsg As String
    
    With vsAdvice
        If rsSend!������� = "7" Then
            strMsg = "���в�ҩ��Ӧ����ҩ�䷽�޷����ͣ�" & vbCrLf & vbCrLf & "����" & NVL(rsSend!ҽ������)
        Else
            strMsg = "��ҩƷ(��һ����ҩ������ҩƷ)�޷����ͣ�" & vbCrLf & vbCrLf & "����" & NVL(rsSend!ҽ������)
        End If
        strMsg = strMsg & vbCrLf & vbCrLf & "û�з�����Ч��ҩƷ�����Ϣ����ҩƷ�����Ѿ���ͣ�û�������סԺ���ˡ�"
        strMsg = strMsg & vbCrLf & "���ȵ�ҩƷĿ¼�����д�����[ȷ��]������������ҽ����"
        .Redraw = flexRDDirect
        Call .ShowCell(lngRow, COL_ѡ��)
        Screen.MousePointer = 0
        MsgBox strMsg, vbInformation, gstrSysName
        
        Screen.MousePointer = 11
        lngDel���ID = NVL(rsSend!���ID, 0)
        Call DeleteCurRow(lngRow, rsSend!���ID)
        .Refresh: .Redraw = flexRDNone
    End With
End Sub

Private Sub SeekMatchDrug(rsSend As ADODB.Recordset, rsDrug As ADODB.Recordset, ByVal dbl���� As Double, vBookMark As Variant, Optional strList As String)
'���ܣ�����ҩƷ�Ķ�����λȱʡ���ʵĹ��,���������ҩƷ��Ϣ�������
'������rsSend=Ҫ���͵�ҽ����Ϣ
'      rsDrug=ҩƷ��Ϣ
'      dbl����=Ҫ���͵�ҩƷ����,Ϊ0ʱ��ʾ��δ�������
'      vBookMark=�������ڶ�λ���λ�õ���ǩ
'      strList=������Ч�ɹ�ѡ��Ĺ��,������������������
    Dim vPreBookMark As Variant
    Dim lng���� As Long
        
    vPreBookMark = 0
    If Not rsDrug.EOF And Not rsDrug.BOF Then
        vPreBookMark = rsDrug.Bookmark
    End If
    
    rsDrug.MoveFirst
    vBookMark = 0: strList = ""
    Do While Not rsDrug.EOF
        '�ſ�ͣ�õ�ҩƷ
        If NVL(rsDrug!����ʱ��, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", NVL(rsDrug!�������, 0)) > 0 Then
            If CInt(NVL(rsSend!��������, 0)) <> 0 And (NVL(rsDrug!���, 0) > dbl���� Or NVL(rsDrug!���, 0) = dbl���� And dbl���� <> 0) Then
                'Ѱ�Ҽ�����λΪ��������С�����Ĺ��
                If rsDrug!����ϵ�� / rsSend!�������� = Int(rsDrug!����ϵ�� / rsSend!��������) Then
                    If rsDrug!����ϵ�� / rsSend!�������� < lng���� Or lng���� = 0 Then
                        vBookMark = rsDrug.Bookmark
                        lng���� = rsDrug!����ϵ�� / rsSend!��������
                    End If
                End If
            End If
            strList = strList & "|#" & rsDrug!ҩƷID & ";" & rsDrug!���� & IIF(Not IsNull(rsDrug!����), "(" & rsDrug!���� & ")", "") & IIF(Not IsNull(rsDrug!���), " " & rsDrug!���, "") & _
                vbTab & IIF(InStr(GetInsidePrivs(pסԺҽ������), "��ʾҩƷ���") = 0, _
                    IIF(NVL(rsDrug!���, 0) > 0, "�п��", "�޿��"), "���:" & NVL(rsDrug!���, 0) & rsDrug!סԺ��λ)
        End If
        rsDrug.MoveNext
    Loop
    If vBookMark = 0 Then
        rsDrug.MoveFirst
        Do While Not rsDrug.EOF
            If NVL(rsDrug!����ʱ��, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", NVL(rsDrug!�������, 0)) > 0 Then
                If NVL(rsDrug!���, 0) > dbl���� Or NVL(rsDrug!���, 0) = dbl���� And dbl���� <> 0 Then
                    vBookMark = rsDrug.Bookmark: Exit Do
                End If
                'ȷ���ܹ�ѡ��һ��δͣ�õĹ��������ù���涼Ϊ0����rsDrugԭ��λ�õļ�¼��ͣ�ù����ᵼ�½������ͣ�ù�񣬲��ܱ�����
                vBookMark = rsDrug.Bookmark
            End If
            rsDrug.MoveNext
        Loop
    End If
    strList = Mid(strList, 2)
    
    If vBookMark = 0 And vPreBookMark <> 0 Then 'û�ҵ�ʱ�ָ�ԭ��λ��
        rsDrug.Bookmark = vPreBookMark
    End If
End Sub

Private Function Calc��������ʱ��(dbl���� As Double, lng���� As Long, str�ֽ�ʱ�� As String, ByVal strEnd As String, rsSend As ADODB.Recordset, rsDrug As ADODB.Recordset, ByVal bln�ڷ� As Boolean) As Boolean
'���ܣ��Գ��ڳ�ҩҽ����������,ִ�д���,ִ��ʱ��ֽ�
'������rsDrug=����ҩƷ���������Ϣ
'      rsSend=������ǰҩƷҽ���������Ϣ
'      strEnd=���η��͵Ľ���ʱ��
'���أ�dbl����=סԺ��װ
'      lng����=ִ�д���(��Ϊ��ҩ;����ִ�д���)
'      str�ֽ�ʱ��=�����ִ��ʱ��ֽ�
    Dim datBegin As Date, datEnd As Date, strPause As String
    Dim datTmp As Date
    Dim intType As Integer
    Dim varArr As Variant
    Dim i As Long
    Dim strTmp As String
    Dim strToday As String
    Dim strTimRange As String
    
    Dim lng����Tmp As Long 'bln�ڷ� true �����Ѿ������˷ֽ�ʱ��ʱ���ڼ��������ﲻ������ lng���� str�ֽ�ʱ�䣬(90797)
    Dim str�ֽ�ʱ��Tmp As String
    
    On Error GoTo errH
    
    '��ǰҽ������ͣʱ���:"��ͣʱ��,��ʼʱ��;...."
    If rsSend!ҽ��״̬ <> 1 Then
        strPause = GetAdvicePause(rsSend!ID, Val(rsSend!��ID & ""))
    End If
    
    '��ǰҽ���ķ��ͼ���ʱ���
    datBegin = rsSend!��ʼִ��ʱ��
    If Not IsNull(rsSend!�ϴ�ִ��ʱ��) Then
        datBegin = Calc�����ڿ�ʼʱ��(rsSend!��ʼִ��ʱ��, rsSend!�ϴ�ִ��ʱ��, NVL(rsSend!Ƶ�ʼ��, 0), rsSend!�����λ & "")
        
        '����������ִ�е�ʱ�䲻�ټ���,����ͨ����ͣ��ʽ������
        '����ҽ����֧����ͣ
        If mlngRefModld = 0 Then
            strPause = strPause & ";" & Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "," & Format(rsSend!�ϴ�ִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
            If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
        End If
    End If
    
    If bln�ڷ� And Val(rsDrug!�Ƿ��ҩ & "") = 1 Then strEnd = Split(strEnd, " ")(0) & " " & mstrEndPoint
    
    datEnd = CDate(strEnd)
    If Not IsNull(rsSend!ִ����ֹʱ��) Then
        If rsSend!ִ����ֹʱ�� < CDate(strEnd) Then
            datEnd = rsSend!ִ����ֹʱ��
        End If
    End If
    
    'ҩƷ�����ӵ��쿪ʼ���� ֻ��׼ȷ���� ��ҩ���ڳ���Ϊ 1 �� ҽ����һ����ȡ���ַ�ʽ���ͺ���ǰδ���͵Ľ����ٱ���ȡ
    If mlngRefModld = 0 And chkҩƷ��Ч(0).value = 1 And chkSendNow.value = 1 Then
        
        strToday = Format(mdatCurr, "yyyy-MM-dd 00:00:00")
        
        If rsSend!�����λ & "" = "��" And Val(NVL(rsSend!Ƶ�ʼ��, 0)) = 1 Or InStr(",Сʱ,����,", "," & rsSend!�����λ & ",") > 0 Then
            datTmp = CDate(strToday)
            If datTmp > datBegin Then
                If rsSend!�����λ & "" = "Сʱ" Then
                    datBegin = CalcÿСʱ����ӵ��췢�Ϳ�ʼʱ��(datBegin, datTmp, NVL(rsSend!Ƶ�ʼ��, 0))
                Else
                    datBegin = datTmp
                End If
            End If
        Else
            intType = 1
        End If
    End If
     
    '�Ȱ���������ʱ��μ���ֽ�ʱ�估����
    str�ֽ�ʱ�� = Calc���ڷֽ�ʱ��(datBegin, datEnd, strPause, NVL(rsSend!ִ��ʱ�䷽��), NVL(rsSend!Ƶ�ʴ���, 0), NVL(rsSend!Ƶ�ʼ��, 0), NVL(rsSend!�����λ), rsSend!��ʼִ��ʱ��)
    
    If intType = 1 And str�ֽ�ʱ�� <> "" Then
        If Format(datBegin, "yyyy-MM-dd HH:mm:ss") < strToday Then
            varArr = Split(str�ֽ�ʱ��, ",")
            For i = UBound(varArr) To 0 Step -1
                If strToday <= CStr(varArr(i)) Then
                    strTmp = strTmp & "," & varArr(i)
                Else
                    Exit For
                End If
            Next
            str�ֽ�ʱ�� = Mid(strTmp, 2)
        End If
    End If
    
    If Val(rsSend!ҽ����Ч & "") = 0 And Val(rsSend!������־ & "") = 1 And Val(rsSend!���״̬ & "") = 1 Then
        datBegin = rsSend!��ʼִ��ʱ��
        datEnd = DateAdd("d", 1, datBegin)
        strTimRange = Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "," & Format(datEnd, "yyyy-MM-dd HH:mm:ss")
        str�ֽ�ʱ�� = GetTimPointsInRange(strTimRange, str�ֽ�ʱ��)
    End If
    
    If str�ֽ�ʱ�� = "" Then
        dbl���� = 0
        lng���� = 0
        Calc��������ʱ�� = True
        Exit Function
    End If

    lng���� = UBound(Split(str�ֽ�ʱ��, ",")) + 1
    
    If NVL(rsSend!�������) = "7" Then
        '��ҩ�䷽����
        dbl���� = lng����
    Else
        If bln�ڷ� And Val(rsDrug!�Ƿ��ҩ & "") = 1 Then
            lng����Tmp = lng����
            str�ֽ�ʱ��Tmp = str�ֽ�ʱ��
        End If
        
        '��ҩ���г�ҩ���ٰ�ҩƷ�������Լ�������(��סԺ��λ),��ʱ�����ͷֽ�ʱ���������
        dbl���� = Calc����ҩƷ����( _
            rsSend!��ʼִ��ʱ��, lng����, str�ֽ�ʱ��, rsSend!��������, _
            rsDrug!����ϵ��, rsDrug!סԺ��װ, NVL(rsSend!�ɷ����, NVL(rsDrug!�ɷ����, 0)), _
            NVL(rsSend!ִ����ֹʱ��, CDate("3000-01-01")), strPause, NVL(rsSend!ִ��ʱ�䷽��), _
            rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ & "", mblnLimit, NVL(rsSend!�״�����, 0), NVL(rsSend!�ϴ�ִ��ʱ��, CDate(0)))
            
        If bln�ڷ� And Val(rsDrug!�Ƿ��ҩ & "") = 1 Then
            lng���� = lng����Tmp
            str�ֽ�ʱ�� = str�ֽ�ʱ��Tmp
        End If
    End If
    
    Calc��������ʱ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetWhere(ByVal bytMode As Byte, ByRef bln���� As Boolean)
'���ܣ�����ҽ��У�Ի��͵Ŀɲ���ҽ�����������û��Ȩ��ʱ��ֻ�ܴ���ǰ������Ա���������������п��һ��߻�������´��ҽ����
'������0-У�ԣ�1=����
'       bln���� ���������Ƿ�Ҫ��ȡ����ҽ��IDs
    Dim strTmp As String
    Dim blnDo As Boolean
    
    If bytMode = 0 Then
        blnDo = InStr(GetInsidePrivs(pסԺҽ������), "ȫԺҽ��У��") = 0
    Else
        blnDo = InStr(GetInsidePrivs(pסԺҽ������), "ȫԺҽ������") = 0
    End If
    
    If blnDo Then
        If gbln��������´�ҽ������ Then
            strTmp = " And (A.��������ID In (Select /*+cardinality(E,10)*/ e.Column_Value From Table(f_Num2list([4])) E) and nvl(a.����ҽ��id,0)=0 or instr(','||[11]||',',','||nvl(a.����ҽ��id,0)||',')>0)"
            bln���� = True
        Else
            strTmp = " And A.��������ID In (Select /*+cardinality(E,10)*/ e.Column_Value From Table(f_Num2list([4])) E) "
        End If
    End If
    
    GetWhere = strTmp
End Function

Private Function CheckSendPrivs(ByVal lngҽ��ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ҽ��ID As Long) As Boolean
'���ܣ��жϵ�ǰҽ���еĿ��������Ƿ��Ǳ�������������
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strDepts As String
    
    strDepts = GetUser����IDs(True)   '��ǰ������Ա���������������п���'
    
    If gbln��������´�ҽ������ Then
        strSQL = " Select 1 From ����ҽ����¼ D Where D.ID = [1] And D.��������ID In (Select /*+cardinality(E,10)*/ e.Column_Value From Table(f_Num2list([2])) E) And nvl(D.����ҽ��id,0)=0" & _
            " union all Select 1 From ����ҽ����¼ D Where D.ID = [3] And D.��������ID In (Select /*+cardinality(E,10)*/ e.Column_Value From Table(f_Num2list([2])) E)"
    Else
        strSQL = " Select 1 From ����ҽ����¼ D Where D.ID = [1] And D.��������ID In (Select /*+cardinality(E,10)*/ e.Column_Value From Table(f_Num2list([2])) E)"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, strDepts, lng����ҽ��ID)
    CheckSendPrivs = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadAdviceSend(ByVal str����IDs As String, ByVal str��ҳIDs As String, ByVal strEnd As String, ByVal str��ҩIDs As String, ByVal str���s As String, ByVal lngModle As Long, ByVal str���˿���IDs As String)
'���ܣ������˶�ȡҽ�������嵥
    Dim rsSendҩƷ As ADODB.Recordset, rsSend���� As ADODB.Recordset
    Dim arrPati As Variant, arrPatiPage As Variant, arrPatiDept As Variant
    
    Dim blnҩƷ As Boolean, bln���� As Boolean, blnOnePati As Boolean, bln�ɽ��ղ��� As Boolean
    Dim lngҩ��ID As Long, lngִ�п���ID As Long, strִ�п��� As String, str������� As String
    
    Dim strSQLҩƷ As String, strSQL���� As String, str��Ҫ���� As String
    Dim strҩ������ As String, str��ҩ;�� As String, strҩ���û� As String
    Dim lng������ As Long, lng����ID As Long, str���� As String, blnƷ��ҩƷ As Boolean, lng������ As Long
    Dim i As Long, k As Long, datEnd As Date
    Dim str���� As String, str���� As String, strҽ����Ч As String
    Dim blnʱ����ʾ As Boolean, bln�����ʾ As Boolean, blnĬ�Ϸ��� As Boolean, bln�洢�ⷿ��ʾ As Boolean
    Dim strDepts As String, strTmp1 As String, strTmp2, strtmp3 As String
    Dim str����ҽ�� As String
    Dim str��ҺҩƷ�ų� As String '�Ƿ���Է�����ҺҩƷ
    Dim strAdDrugIDs As String
    Dim str����ҽ��IDs As String
    Dim bln���� As Boolean
    Dim str����ҽ���ų� As String
    
    mstrNoneIDs = ","
    mstrAdDrugIDs = ""
    
    Screen.MousePointer = 11
    stbThis.Panels(3).Text = "" ': Call Form_Resize
    Call GetAdvicePause(0) '����˷����еĻ���
    blnʱ����ʾ = True: bln�����ʾ = True: blnĬ�Ϸ��� = True: bln�洢�ⷿ��ʾ = True
    blnҩƷ = chkҩƷ��Ч(0).value = 1 Or chkҩƷ��Ч(1).value = 1
    bln���� = chk������Ч(0).value = 1 Or chk������Ч(1).value = 1
    
    If chkҩƷ��Ч(0).value = 1 Or chk������Ч(0).value = 1 Then
        lblInfo.Caption = "���η��ͣ�" & IIF(chkҩƷ��Ч(1).value = 1 Or chk������Ч(1).value = 1, "��ʱҽ����", "") & "����ҽ��������ʱ�䣺" & strEnd
    Else
        lblInfo.Caption = "���η��ͣ���ʱҽ��"
    End If
    
    vsPrice.Rows = vsPrice.FixedRows
    vsPrice.Rows = vsPrice.FixedRows + 1
    
    With vsAdvice
        .Rows = .FixedRows '��ɾ���й���
        If mblnOnePati Then
            .ColHidden(COL_����) = True
            .ColHidden(COL_סԺ��) = True
            .ColHidden(COL_����) = True
            .ColHidden(COL_�ѱ�) = True
        End If
        .ColHidden(COL_����) = True
        .ColHidden(COL_Ӥ��) = True
        .ColHidden(COL_����) = False
        .ColHidden(COL_������λ) = False
        .ColHidden(COL_�״�ʱ��) = chkҩƷ��Ч(1).value = 1 Or chk������Ч(1).value = 1
        .ColHidden(COL_ĩ��ʱ��) = chkҩƷ��Ч(1).value = 1 Or chk������Ч(1).value = 1
        
        .ColHidden(COL_���) = Not blnҩƷ
        .ColHidden(COL_ִ������) = Not blnҩƷ
        If chkҩƷ��Ч(0).value = 0 And chk������Ч(0).value = 0 Or chkҩƷ��Ч(1).value = 0 And chk������Ч(1).value = 0 Then
            .ColHidden(COL_ҽ����Ч) = True
        Else
            .ColHidden(COL_ҽ����Ч) = False
        End If
    End With
    Me.Refresh
    
    strDepts = GetUser����IDs(True)    '��ǰ������Ա���������������п���
        
    str��Ҫ���� = " And A.��ʼִ��ʱ�� is Not NULL And Nvl(A.ִ�б��,0)<>-1 And A.������Դ<>3"
    'Ӥ���Ĵ���
    If optӤ��(1).value Or optӤ��(2).value Then
        str��Ҫ���� = str��Ҫ���� & " And Nvl(A.Ӥ��,0)" & IIF(optӤ��(1).value, "=0", ">0")
    End If
    str��Ҫ���� = str��Ҫ���� & IIF(Not mblnҽ������, " And A.ǰ��ID is Null", "")
    
    If opt״̬(1).value Then    '��У��
        '��ǰ����Ա���������Ŀ��ҵ�����ҽ��
        str��Ҫ���� = str��Ҫ���� & GetWhere(1, bln����)
    Else
        If opt״̬(0).value Then    '�¿�
            str��Ҫ���� = str��Ҫ���� & " And Exists(" & _
                "Select M.���� From ��Ա�� M,ִҵ��� N" & _
                " Where M.����=Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1)," & _
            "2,Substr(A.����ҽ��,1,Decode(Instr(A.����ҽ��,'/'),0,length(A.����ҽ��),Instr(A.����ҽ��,'/')-1))," & _
            "Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))" & _
                " And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ')" & _
                " )"
                
            str��Ҫ���� = str��Ҫ���� & GetWhere(0, bln����)
        Else '����
            str��Ҫ���� = str��Ҫ���� & " And (Nvl(A.ҽ��״̬,0)<>1 Or A.ҽ��״̬=1 And Exists(" & _
                "Select M.���� From ��Ա�� M,ִҵ��� N" & _
                " Where M.����=Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1)," & _
            "2,Substr(A.����ҽ��,1,Decode(Instr(A.����ҽ��,'/'),0,length(A.����ҽ��),Instr(A.����ҽ��,'/')-1))," & _
            "Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))" & _
                " And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ')" & _
                " ))"
                
            strTmp1 = GetWhere(0, bln����)
            strTmp2 = GetWhere(1, bln����)
            If Not (strTmp1 = "" And strTmp2 = "") Then
                str��Ҫ���� = str��Ҫ���� & " And (Nvl(A.ҽ��״̬,0)<>1" & strTmp2 & " Or A.ҽ��״̬=1" & strTmp1 & ")"
            End If
        End If
    End If

    
    '�û���ҩ��ID
    If blnҩƷ Then
        
        lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
        strҩ���û� = Decodeҩ���û�
        
        'ֻ����ָ��ҩ����ҩƷ:ҩ���û�֮���Ϊ׼
        If lngҩ��ID <> 0 Then
            strҩ������ = "Select ID From ����ҽ����¼ X" & _
                " Where ������� IN('5','6','7') And (X.���ID=A.���ID Or X.���ID=A.ID)" & _
                " And ִ�п���ID+0=[3] And ����ID=[2]"
            strҩ������ = " And Exists(" & strҩ������ & ")"
        End If
        
        '����ĸ�ҩ;������(������Ӧ�ĳ�ҩ)
        If str��ҩIDs <> "" Then
            str��ҩ;�� = "Select ID From ����ҽ����¼ X" & _
                " Where �������='E' And (X.ID=A.ID Or X.ID=A.���ID)" & _
                " And ����ID=[2] And ������ĿID+0 IN(" & str��ҩIDs & ")"
            str��ҩ;�� = " And Exists(" & str��ҩ;�� & ")"
        End If
        
        str���� = ""
        str���� = ""
        strҽ����Ч = ""
        '��ͬ��Ч������
        If chkҩƷ��Ч(0).value = 1 Then    '����
            strTmp1 = _
                "A.��ʼִ��ʱ��<=[1] And (A.�ϴ�ִ��ʱ��<[1] Or A.�ϴ�ִ��ʱ�� is NULL)" & _
                " And (A.ִ����ֹʱ��>A.�ϴ�ִ��ʱ�� Or A.ִ����ֹʱ�� is NULL Or A.�ϴ�ִ��ʱ�� Is NULL)" & _
                " And (A.ִ����ֹʱ��>A.��ʼִ��ʱ�� Or A.ִ����ֹʱ�� is NULL) And A.ҽ����Ч=0"
                
            If opt״̬(1).value Then    '��У��
                str���� = strTmp1 & " And Nvl(A.ҽ��״̬,0) Not IN(-1,1,2,4)"
            Else
                If opt״̬(0).value Then    '�¿�(���ܽ���ʱ�䣬����ʱ����ʼִ��ʱ�����ָ���ķ��ͽ���ʱ��ŷ���)
                    str���� = "A.ҽ��״̬=1 And A.ҽ����Ч=0"
                Else '����
                    str���� = "(A.ҽ��״̬=1 And A.ҽ����Ч=0 Or (" & strTmp1 & " And Nvl(A.ҽ��״̬,0) Not IN(-1,1,2,4)))"
                End If
            End If
        End If
        If chkҩƷ��Ч(1).value = 1 Then    '����
            If opt״̬(1).value Then    '��У��
                str���� = "Nvl(A.ҽ��״̬,0) Not IN(-1,1,2,4,8,9) And A.ҽ����Ч=1"
            Else
                If opt״̬(0).value Then    '�¿�
                    str���� = "A.ҽ��״̬=1 And A.ҽ����Ч=1"
                Else '����
                    str���� = "(A.ҽ��״̬=1 And A.ҽ����Ч=1 Or Nvl(A.ҽ��״̬,0) Not IN(-1,2,4,8,9) And A.ҽ����Ч=1)"
                End If
            End If
        End If
        If str���� <> "" And str���� <> "" Then '������ͬʱΪ��
            strTmp1 = " And ((" & str���� & ") Or (" & str���� & "))"
            If strTmp1 = " And ((A.ҽ��״̬=1 And A.ҽ����Ч=0) Or (A.ҽ��״̬=1 And A.ҽ����Ч=1))" Then
                strTmp1 = " And A.ҽ��״̬=1 And A.ҽ����Ч In(0,1)"
            End If
            strҽ����Ч = strTmp1
        ElseIf str���� <> "" Then
            strҽ����Ч = " And " & str����
        ElseIf str���� <> "" Then
            strҽ����Ч = " And " & str����
        End If
        
        If gblnKSSStrict Then
            If opt״̬(0).value Or opt״̬(2).value Then
                strҽ����Ч = strҽ����Ч & " And (A.ҽ��״̬<>1 Or A.ҽ��״̬=1 And  (Nvl(A.���״̬,0) Not in(1,3) or a.ҽ����Ч=0 and a.���״̬=1 and a.������־=1 and (instr(',5,6,',A.�������)>0 or A.�������='E' and E.��������='2')))"
            End If
        End If
        
        '����ҽ��
        If lngModle = 1 Then
            str����ҽ�� = " And (NVL(a.ִ��Ƶ��,'��')='��Ҫʱ' Or NVL(a.ִ��Ƶ��,'��')='��Ҫʱ' And to_date([6],'yyyy-mm-dd hh24:mi') - a.��ʼִ��ʱ��<0.5) "
        Else
            str����ҽ�� = " And (NVL(a.ִ��Ƶ��,'��')<>'��Ҫʱ' And NVL(a.ִ��Ƶ��,'��')<>'��Ҫʱ') "
        End If
        
        '��ȡ������ϸ:(δ�ų�����������ҽ����)
        '����������(��ҩ;��,�÷�,�巨����Ϊ),�������ȶ�ȡ����
        strSQLҩƷ = "Select A.ID,A.���ID,Nvl(A.���ID,A.ID) as ��ID,Nvl(X.���,A.���) as ���," & _
            " A.�������,A.������ĿID,E.���� as ������Ŀ,A.�շ�ϸĿID,A.Ӥ��,B.��Ժ����," & _
            " A.����ID,A.��ҳID,B.סԺ��,B.��Ժ���� as ����,D.���� as ����,A.����,A.�Ա�,A.����,B.�ѱ�,B.����," & _
            IIF(mlngRefModld = 1, "Decode(Sign(A.��ʼִ��ʱ��-to_date([6],'yyyy-mm-dd hh24:mi')),-1,to_date([6],'yyyy-mm-dd hh24:mi'),A.��ʼִ��ʱ��) as ��ʼִ��ʱ��", "A.��ʼִ��ʱ��") & ",A.�ϴ�ִ��ʱ��,A.ҽ������,A.����,A.�ܸ�����,A.��������,E.���㵥λ,A.ִ����ֹʱ��," & _
            " A.ִ��Ƶ��,Decode(A.ִ��Ƶ��,'��Ҫʱ',1,A.Ƶ�ʴ���) As Ƶ�ʴ���,Decode(A.ִ��Ƶ��,'��Ҫʱ',1,A.Ƶ�ʼ��) As Ƶ�ʼ��,Decode(A.ִ��Ƶ��,'��Ҫʱ','��',A.�����λ) as �����λ,A.ҽ������," & _
            " Decode(A.ִ��Ƶ��,'��Ҫʱ',[7],A.ִ��ʱ�䷽��) As ִ��ʱ�䷽��,b.��������,a.����ʱ��,e.ִ�з���,e.��������," & _
            " [5] as ���˲���ID,A.���˿���ID,A.��������ID,A.����ҽ��," & IIF(mblnAutoVerify, "s.����ʱ�� as �¿�����ʱ��,", "") & _
            " A.�ɷ����,A.�Ƽ�����,A.ִ������,A.ִ�б��," & strҩ���û� & " as ִ�п���ID,Nvl(F.����,Decode(Nvl(A.ִ������,0),5,'-')) as ִ�п���,A.ժҪ,A.ҽ��״̬,A.ҽ����Ч,A.�״�����,E.���㷽ʽ,e.ִ�а���,g.�������,a.������־,a.���״̬,a.��ҩ����,a.����ҽ��ID" & _
            " From ����ҽ����¼ A,������ҳ B,������Ϣ C,���ű� D,������ĿĿ¼ E,���ű� F,ҩƷ���� G,����ҽ����¼ X" & IIF(mblnAutoVerify, ",����ҽ��״̬ S", "") & _
            " Where A.����ID=[2] And A.����ID=C.����ID And B.��Ժ����ID=D.ID" & IIF(mblnAutoVerify, " And  s.ҽ��ID=a.ID And s.��������=1 ", "") & _
            " And A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.��ҳID = C.��ҳID" & _
            " And A.���ID=X.ID(+) And A.������ĿID=E.ID And " & strҩ���û� & "=F.ID(+) and e.id=g.ҩ��id(+)" & _
            " And A.������� IN('5','6','7','E')" & str��Ҫ���� & strҩ������ & str��ҩ;�� & strҽ����Ч & str����ҽ�� & _
            " And (B.Ӥ������ID is null or B.Ӥ������ID is not null and B.Ӥ������ID=[5] and NVL(A.Ӥ��,0)<>0 or B.Ӥ������ID is not null and B.Ӥ������ID<>[5] and NVL(A.Ӥ��,0)=0) "
        strtmp3 = strSQLҩƷ
    End If
    
    If bln���� Then
    
        lngִ�п���ID = Val(cmdִ�п���.Tag)
        
        'ִ�п���(����Ҫҽ��Ϊ׼)
        If lngִ�п���ID <> 0 Then
            'һ����Ŀ�Լ��������,������;������Ŀ(���)
            strִ�п��� = _
                " And Exists(" & _
                " Select ID From ����ҽ����¼ X" & _
                " Where ���ID is Null And (X.ID=A.ID Or X.ID=A.���ID)" & _
                " And ����ID=[2] And ִ�п���ID+0=[3]" & _
                " Union ALL " & _
                " Select ID From ����ҽ����¼ X" & _
                " Where ���ID is Not Null And �������='C' And (X.���ID=A.���ID Or X.���ID=A.ID)" & _
                " And ����ID=[2] And ִ�п���ID+0=[3])"
        End If
        
        '�����������𲿷�(����Ҫҽ��Ϊ׼)
        If str���s <> "" Then
            'һ����Ŀ�Լ��������,������;������Ŀ(���)
            str������� = _
                " And Exists(" & _
                " Select ID From ����ҽ����¼ X" & _
                " Where ���ID is Null And (X.ID=A.ID Or X.ID=A.���ID)" & _
                " And ����ID=[2] And ������� IN(" & str���s & ")" & _
                " Union ALL" & _
                " Select ID From ����ҽ����¼ X" & _
                " Where ���ID is Not Null And �������='C' And (X.���ID=A.���ID Or X.���ID=A.ID)" & _
                " And ����ID=[2] And ������� IN(" & str���s & "))"
        End If
        
        '�ſ���ҩ;������ҩ�巨���÷�
        strSQL���� = _
            " And Not(A.�������='E' And A.���ID is Not NULL And E.��������='3')" & _
            " And Not Exists(Select ID From ����ҽ����¼ X" & _
            " Where ������� IN('5','6','7') And X.���ID=A.ID" & _
            " And ����ID=[2])"
            
        '����������(����,���,���鲻����Ϊ����,�ɼ���������Ѫ;������Ϊ����������Ҫ������,����󲻷���)
        '����ǲ�У�Լ�����ģʽ������ִ�еĲ�����(����,���,���鲻����Ϊ����,�ɼ���������Ѫ;������Ϊ����)
        If Not mblnAutoVerify Or opt״̬(1).value Then
            strSQL���� = strSQL���� & " And A.������� is Not Null And (Nvl(A.ִ������,0)<>0 Or A.�������='E' And E.�������� IN('6','8'))"
        ElseIf opt״̬(2).value Then
            strSQL���� = strSQL���� & " And (A.ҽ��״̬=1 Or Nvl(A.ҽ��״̬,0)<>1 And A.������� is Not Null" & _
                        " And (Nvl(A.ִ������,0)<>0 Or A.�������='E' And E.�������� IN('6','8')))"
        End If
                
                
        '��ͬ��Ч������
        str���� = ""
        str���� = ""
        strҽ����Ч = ""
        If chk������Ч(0).value = 1 Then    '����
            '��ǰ����ҽ��������(���Ѽ������У��)
            '�¿��Ļ��ⳤ��������ȼ�,����/Σҽ��,��¼�����ҽ�������͵�Ҫ������У��(���û�л���ҽ����֮ǰû�е�����У�ԣ���������Ҫ������У��)
            strTmp1 = _
                "A.��ʼִ��ʱ��<=[1] And (A.�ϴ�ִ��ʱ��<[1] Or A.�ϴ�ִ��ʱ�� is NULL)" & _
                " And (A.ִ����ֹʱ��>A.�ϴ�ִ��ʱ�� Or A.ִ����ֹʱ�� is NULL Or A.�ϴ�ִ��ʱ�� Is NULL)" & _
                " And (A.ִ����ֹʱ��>A.��ʼִ��ʱ�� Or A.ִ����ֹʱ�� is NULL) And A.ҽ����Ч=0" & _
                " And Not(Nvl(a.�������,'����')='H' And E.��������='1' And E.ִ��Ƶ��=2)" & _
                " And Not(Nvl(a.�������,'����')='Z' And E.�������� IN('4','14','9','10','12'))"
                
            If opt״̬(1).value Then    '��У��
                str���� = strTmp1 & " And Nvl(A.ҽ��״̬,0) Not IN(-1,1,2,4)"
            Else
                If opt״̬(0).value Then    '�¿�(���ܽ���ʱ�䣬����ʱ����ʼִ��ʱ�����ָ���ķ��ͽ���ʱ��ŷ���)
                    str���� = "A.ҽ��״̬=1 And A.ҽ����Ч=0"
                Else '����
                    str���� = "(A.ҽ��״̬=1 And A.ҽ����Ч=0 Or (" & strTmp1 & " And Nvl(A.ҽ��״̬,0) Not IN(-1,1,2,4)))"
                End If
            End If
        End If
        If chk������Ч(1).value = 1 Then    '����
            If opt״̬(1).value Then    '��У��
                str���� = "Nvl(A.ҽ��״̬,0) Not IN(-1,1,2,4,8,9) And A.ҽ����Ч=1"
            Else
                If opt״̬(0).value Then    '�¿�
                    str���� = "A.ҽ��״̬=1 And A.ҽ����Ч=1"
                Else '����
                    str���� = "(A.ҽ��״̬=1 And A.ҽ����Ч=1 Or Nvl(A.ҽ��״̬,0) Not IN(-1,2,4,8,9) And A.ҽ����Ч=1)"
                End If
            End If
        End If
        If str���� <> "" And str���� <> "" Then '������ͬʱΪ��
            strTmp1 = " And ((" & str���� & ") Or (" & str���� & "))"
            If strTmp1 = " And ((A.ҽ��״̬=1 And A.ҽ����Ч=0) Or (A.ҽ��״̬=1 And A.ҽ����Ч=1))" Then
                strTmp1 = " And A.ҽ��״̬=1 And A.ҽ����Ч In(0,1)"
            End If
            strҽ����Ч = strTmp1
        ElseIf str���� <> "" Then
            strҽ����Ч = " And " & str����
        ElseIf str���� <> "" Then
            strҽ����Ч = " And " & str����
        End If
        
         '����ҽ��
        If lngModle = 1 Then
            str����ҽ�� = " And (NVL(a.ִ��Ƶ��,'��')='��Ҫʱ' Or NVL(a.ִ��Ƶ��,'��')='��Ҫʱ'  And to_date([6],'yyyy-mm-dd hh24:mi') - a.��ʼִ��ʱ��<0.5) "
        Else
            str����ҽ�� = " And (NVL(a.ִ��Ƶ��,'��')<>'��Ҫʱ' And NVL(a.ִ��Ƶ��,'��')<>'��Ҫʱ') "
        End If
        
        '�����ּ�����/��Ѫ�ּ�����
        If gbln�����ּ����� Or gbln��Ѫ�ּ����� Or gblnѪ��ϵͳ Then
            If opt״̬(0).value Or opt״̬(2).value Then
                strҽ����Ч = strҽ����Ч & " And (A.ҽ��״̬<>1 Or A.ҽ��״̬=1 And Nvl(A.���״̬,0) Not in " & IIF(gblnѪ��ϵͳ = True, "(1,3,7)", "(1,3,4,5,7)") & ")"
            End If
        End If
        
        '��ȡ������ϸ:
        strSQL���� = "Select A.ID,A.���ID,Nvl(A.���ID,A.ID) as ��ID,Nvl(X.���,A.���) as ���," & _
            " A.�������,G.���� as �������,A.������ĿID,E.���� as ������Ŀ,A.�շ�ϸĿID,B.��Ժ����," & _
            " A.Ӥ��,A.����ID,A.��ҳID,B.סԺ��,B.��Ժ���� as ����,D.���� as ����,A.����,A.�Ա�,A.����,B.�ѱ�,B.����," & _
            IIF(mlngRefModld = 1, "Decode(Sign(A.��ʼִ��ʱ��-to_date([6],'yyyy-mm-dd hh24:mi')),-1,to_date([6],'yyyy-mm-dd hh24:mi'),A.��ʼִ��ʱ��) as ��ʼִ��ʱ��", "A.��ʼִ��ʱ��") & ",A.�ϴ�ִ��ʱ��,A.ҽ������,A.�ܸ�����,A.��������,E.�������," & _
            " Decode(A.�������,'4',H.���㵥λ,E.���㵥λ) as ���㵥λ,A.ִ����ֹʱ��,b.��������,a.����ʱ��,a.�������,e.ִ�з���," & _
            " A.ִ��Ƶ��,Decode(A.ִ��Ƶ��,'��Ҫʱ',1,A.Ƶ�ʴ���) As Ƶ�ʴ���,Decode(A.ִ��Ƶ��,'��Ҫʱ',1,A.Ƶ�ʼ��) As Ƶ�ʼ��,Decode(A.ִ��Ƶ��,'��Ҫʱ','��',A.�����λ) as �����λ,A.ҽ������," & _
            " Decode(A.ִ��Ƶ��,'��Ҫʱ',[7],A.ִ��ʱ�䷽��) As ִ��ʱ�䷽��," & _
            " [5] as ���˲���ID,A.���˿���ID,A.��������ID,A.����ҽ��,A.ժҪ," & _
            " A.�걾��λ,A.��鷽��,A.ִ�б��,A.�Ƽ�����,E.��������,E.�Թܱ���,E.ִ��Ƶ��," & IIF(mblnAutoVerify, "s.����ʱ�� as �¿�����ʱ��,", "") & _
            " A.ִ������,A.ִ�п���ID,Nvl(F.����,Decode(Nvl(A.ִ������,0),5,'-')) as ִ�п���,H.�Ƿ���,I.��������,A.������־,A.ҽ��״̬,A.ҽ����Ч,h.����ʱ��,A.�״�����,E.���㷽ʽ,e.ִ�а���,a.��ҩ����,a.����ҽ��ID" & _
            " From ����ҽ����¼ A,������ҳ B,������Ϣ C,���ű� D,������ĿĿ¼ E," & _
            " ���ű� F,������Ŀ��� G,�շ���ĿĿ¼ H,�������� I,����ҽ����¼ X" & IIF(mblnAutoVerify, ",����ҽ��״̬ S", "") & _
            " Where A.����ID=[2] And A.����ID=C.����ID And B.��Ժ����ID=D.ID And A.�շ�ϸĿID=H.ID(+) And A.�շ�ϸĿID=I.����ID(+)" & IIF(mblnAutoVerify, " And  s.ҽ��ID=a.ID And s.��������=1 ", "") & _
            " And A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.��ҳID = C.��ҳID And A.���ID=X.ID(+)" & _
            " And A.������ĿID=E.ID(+) And E.���=G.����(+) And A.ִ�п���ID=F.ID(+)" & strSQL���� & _
            " And Nvl(a.�������,'����') Not IN('5','6','7')" & str��Ҫ���� & strִ�п��� & str������� & strҽ����Ч & str����ҽ�� & _
            " And (B.Ӥ������ID is null or B.Ӥ������ID is not null and B.Ӥ������ID=[5] and NVL(A.Ӥ��,0)<>0 or B.Ӥ������ID is not null and B.Ӥ������ID<>[5] and NVL(A.Ӥ��,0)=0)" & _
            " Order by A.ҽ����Ч,A.Ӥ��,���,��ID,A.���"
    End If
    
    On Error GoTo errH
    arrPati = Split(str����IDs, ",")
    arrPatiPage = Split(str��ҳIDs, ",")
    arrPatiDept = Split(str���˿���IDs, ",")
    blnOnePati = UBound(arrPati) = 0
    datEnd = CDate(IIF(strEnd = "", "1990-01-01", strEnd))
    
    If cboUnit.ListIndex <> -1 Then
        bln�ɽ��ղ��� = (mstrInfDepIDs = "" Or InStr("," & mstrInfDepIDs & ",", "," & Val(cboUnit.ItemData(cboUnit.ListIndex)) & ",") > 0)
    End If
    For k = 0 To UBound(arrPati)
        If bln���� Then str����ҽ��IDs = Get����ҽ��IDs(Val(arrPati(k)), arrPatiPage(k), strDepts)
        If blnҩƷ Then
            str��ҺҩƷ�ų� = ""
            '�ų����͵��������ĵ���Һ��ҽ��,���������һ��ҩ�����͵���Һ�������ģ������������﷢��(�Ա�ҩ����ȡҩ����Ժ��ҩ����ȡҩ���������ʵ�ҩƷ�����������﷢�͡��������Ĳ����յĸ�ҩ;��ҩƷ���������﷢�͡���Ч�������Ҳ���������﷢��)
            If lngModle = 0 Then
                If gstr��Һ�������� <> "" Then
                    If bln�ɽ��ղ��� Then
                        str����ҽ���ų� = Get��Һ��ҽ��(Val(arrPati(k)), arrPatiPage(k), 0)
                        str��ҺҩƷ�ų� = " and instr(','||[10]|| ',',','||Nvl(A.���ID,A.ID)||',')=0"
                    Else
                        '��������������������ģ����ų����еľ������̵�ҩƷ���������û�����ã���ȫԺ�����ˣ���ֻ�ų�����Ӫ��ҽ��
                        str��ҺҩƷ�ų� = " And NVL(E.ִ�б��,0)<>2 And (Not Exists(Select 1 From ������ĿĿ¼ Y Where X.������Ŀid = y.Id And NVL(Y.ִ�б��,0)=2) OR x.������Ŀid is null)"
                    End If
                End If
            End If
            strAdDrugIDs = ""
            mstrNoneIDs = mstrNoneIDs & GetNoneSendID(Val(arrPati(k)), arrPatiPage(k), 2, , , strAdDrugIDs) & ","
            If strAdDrugIDs <> "" Then
                mstrAdDrugIDs = IIF(mstrAdDrugIDs = "", "", mstrAdDrugIDs & ",") & strAdDrugIDs
            End If
            strSQLҩƷ = strtmp3
            strSQLҩƷ = strSQLҩƷ & str��ҺҩƷ�ų� & " Order by A.ҽ����Ч,A.Ӥ��,���,��ID,A.���"
            Set rsSendҩƷ = zlDatabase.OpenSQLRecord(strSQLҩƷ, Me.Caption, datEnd, Val(arrPati(k)), lngҩ��ID, strDepts, mlng����ID, Format(dkpExecTime.value, "YYYY-MM-DD HH:mm"), Format(dkpExecTime.value, "HH:mm"), _
                gstr��Һ��������, mstr��Һ��ҩ;��, str����ҽ���ų�, str����ҽ��IDs)
        End If
        If bln���� Then
            Set rsSend���� = zlDatabase.OpenSQLRecord(strSQL����, Me.Caption, datEnd, Val(arrPati(k)), lngִ�п���ID, strDepts, mlng����ID, Format(dkpExecTime.value, "YYYY-MM-DD HH:mm"), Format(dkpExecTime.value, "HH:mm"), , , , str����ҽ��IDs)
        End If
        
        '����ʾ�¿���
        If mblnAutoVerify Then
            If blnҩƷ Then
                rsSendҩƷ.Filter = "ҽ��״̬=1"
                If rsSendҩƷ.RecordCount > 0 Then
                    Call LoadAdviceSendDrug(blnOnePati, strEnd, rsSendҩƷ, lng������, str����, blnƷ��ҩƷ, blnʱ����ʾ, bln�����ʾ, blnĬ�Ϸ���, lng����ID, bln�洢�ⷿ��ʾ)
                End If
            End If
            If bln���� Then
                rsSend����.Filter = "ҽ��״̬=1"
                If rsSend����.RecordCount > 0 Then
                    Call LoadAdviceSendOther(blnOnePati, strEnd, rsSend����, lng������, str����, lng������, lng����ID)
                End If
            End If
        End If
        
        If blnҩƷ Then
            If mblnAutoVerify Then rsSendҩƷ.Filter = "ҽ��״̬<>1"
            If rsSendҩƷ.RecordCount > 0 Then
                Call LoadAdviceSendDrug(blnOnePati, strEnd, rsSendҩƷ, lng������, str����, blnƷ��ҩƷ, blnʱ����ʾ, bln�����ʾ, blnĬ�Ϸ���, lng����ID, bln�洢�ⷿ��ʾ)
            End If
        End If
        If bln���� Then
            If mblnAutoVerify Then rsSend����.Filter = "ҽ��״̬<>1"
            If rsSend����.RecordCount > 0 Then
                Call LoadAdviceSendOther(blnOnePati, strEnd, rsSend����, lng������, str����, lng������, lng����ID)
            End If
        End If
        If Not blnOnePati Then
            Progress = k / (UBound(arrPati) + 1) * 100
        End If
    Next
    If Not blnOnePati Then Progress = 0

    If blnҩƷ Then
        If mbln��ҩ�� Then Call Refresh��ҩ��
    End If
    
    
    With vsAdvice
        If mblnOnePati Then
            If .Rows - 1 > .FixedRows Then
                lblInfo.Caption = "������" & .TextMatrix(.Rows - 1, COL_����) & ",סԺ�ţ�" & .TextMatrix(.Rows - 1, COL_סԺ��) & "�����ţ�" & .TextMatrix(.Rows - 1, COL_����) & "," & lblInfo.Caption & IIF(str���� = "", " ", "(" & Mid(str����, 2) & ") ")
            Else
                lblInfo.Caption = "û�ж�ȡ�κ�ҽ����"
            End If
        Else
            lblInfo.Caption = lblInfo.Caption & "������" & IIF(str���� = "", " ", "(" & Mid(str����, 2) & ") ") & lng������ & " �����˵�ҽ��"
        End If
        
        .Redraw = flexRDNone
        If blnҩƷ Then
            .ColHidden(COL_���) = Not blnƷ��ҩƷ Or chkҩƷ��Ч(0).value <> 1
        End If
        
        If bln���� And blnҩƷ = False Then
            .ColHidden(COL_����) = lng������ = 0
        Else
            .ColHidden(COL_����) = False
        End If
        .ColHidden(COL_������λ) = .ColHidden(COL_����)
        
        If Not .ColHidden(COL_���) Then
            .AutoSize col_ҽ������, COL_���
        Else
            .AutoSize col_ҽ������
        End If
        .RowHeight(0) = 320
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
               
        .Col = .FixedCols
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    
    If VsfOnlyOneRow(vsAdvice) Then
        'ֻ��һ��
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 1 Then
            vsAdvice.BackColorSel = BackColorNew
        Else
            vsAdvice.BackColorSel = vbWhite
        End If
    Else
        vsAdvice.BackColorSel = COLSelBackColor
    End If
    
    If vsAdvice.Visible Then vsAdvice.SetFocus
    Call ShowSendTotal
    Screen.MousePointer = 0
        
    Exit Sub
errH:
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        vsAdvice.Redraw = flexRDNone
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadAdviceSendDrug(ByVal blnOnePati As Boolean, ByVal strEnd As String, ByVal rsSend As ADODB.Recordset, ByRef lng������ As Long, ByRef str���� As String, _
    ByRef blnƷ��ҩƷ As Boolean, ByRef blnʱ����ʾ As Boolean, ByRef bln�����ʾ As Boolean, ByRef blnĬ�Ϸ��� As Boolean, ByRef lng����ID As Long, ByRef bln�洢�ⷿ��ʾ As Boolean) As Boolean
'���ܣ���ʾҪ���͵�ҩƷҽ���嵥
'������strEnd=���͵��Ľ���ʱ��(yyyy-MM-dd HH:mm:ss),����û��
'���أ�lng������=�д�����ҽ���Ĳ�����
'      str����=���в��˵�ǰ���Ҵ�
'      blnƷ��ҩƷ=�Ƿ����δȷ������Ʒ��ҩƷ
'˵����ע��CellData�д�ŵ��и�������
'   RowData��0-δ���͵�,-1-�ѳɹ����͵�
'   COL_ѡ��0-������ѡ���,1-��ֹ�ı�ѡ��״̬��
'   COL_Ӥ�������Ӥ�����
'   COL_�������1-��ҩ;����2-��ҩ�巨��3-��ҩ�÷���ֻ�ڱ�������ʹ��
'   COL_ҽ�����ݣ����������Ŀ����,������ʾ�Ƽ�ҽ��
'   COL_�ֽ�ʱ��:�����޷ֽ�ʱ��ʱ,��ŷ��÷���ʱ��
'   COL_��񣺴�ų�ҩ��ѡ��Ĺ����������(ComboList)
'   COL_����żƼ������Ƿ�����
    
    Dim rsDrug As New ADODB.Recordset
    Dim i As Long, j As Long, k As Long, lngTmp As Long, strTmp As String
    Dim lngRow As Long, lngDel���ID As Long, vBookMark As Variant
    Dim lng���� As Long, lng��С���� As Long, str�÷� As String
    Dim str�ֽ�ʱ�� As String, dbl���� As Double, cur��� As Currency
    Dim blnReCalc As Boolean
    Dim rsTmp As Recordset, strSQL As String, strIDs As String
    Dim bln�ڷ� As Boolean
    Dim lng��ҽ��ID As Long
    Dim lng��ҽ��IDTmp As Long
    Dim varBookmark As Variant
     
    On Error GoTo errH
    '���㲢��ʾ�����嵥
    '----------------------------------------------------------------------------------------------------------
    With vsAdvice
        .Redraw = flexRDNone
        For i = 1 To rsSend.RecordCount
            If rsSend!������� = "E" And IsNull(rsSend!���ID) And rsSend!ID <> Val(.TextMatrix(.Rows - 1, COL_���ID)) Then
                GoTo NextLoop '������ҩ����������ҽ�������ɼ�����
            ElseIf rsSend!������� = "E" And Not IsNull(rsSend!���ID) And NVL(rsSend!���ID, 0) <> Val(.TextMatrix(.Rows - 1, COL_���ID)) Then
                GoTo NextLoop '������Ѫ;��
            ElseIf (rsSend!ID = lngDel���ID Or NVL(rsSend!���ID, 0) = lngDel���ID) And lngDel���ID <> 0 Then
                GoTo NextLoop 'һ����ҩ���䷽�е�һ�������Ѿ����ܷ���,�����鲻�ܷ���
            Else
                lngDel���ID = 0
            End If
            
            If mstrEndPoint <> "" Then
                lng��ҽ��IDTmp = IIF(IsNull(rsSend!���ID), rsSend!ID, rsSend!���ID)
                If lng��ҽ��ID <> lng��ҽ��IDTmp Then
                    lng��ҽ��ID = lng��ҽ��IDTmp
                    bln�ڷ� = False
                    If InStr(",5,6,7,", "," & rsSend!������� & ",") > 0 Then
                        varBookmark = rsSend.Bookmark
                        rsSend.MoveNext
                        Do While Not rsSend.EOF
                            If rsSend!������� = "E" And (Val(rsSend!�������� & "") = 2 Or Val(rsSend!�������� & "") = 4) Then
                                If Val(rsSend!ִ�з��� & "") = 4 Then
                                    bln�ڷ� = True
                                End If
                                Exit Do
                            End If
                            rsSend.MoveNext
                        Loop
                        rsSend.Bookmark = varBookmark
                    End If
                End If
            End If
            
            '���뵱ǰ��
            .Rows = .Rows + 1: lngRow = .Rows - 1
            .Cell(flexcpPictureAlignment, lngRow, COL_ѡ��) = 4
            
            If InStr(mstrNoneIDs, "," & CStr(rsSend!ID) & ",") > 0 And Not mbln������ҩ Then
                Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing
            Else
                Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
            End If
            
            
            If mlngRefModld = 1 Then
                If Format(rsSend!�ϴ�ִ��ʱ��, "YYYY-MM-DD") >= Format(mdatCurr, "YYYY-MM-DD") Then '�����ѷ��͵�(�������ܷ��͵�����)
                    .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HA08000           '����
                End If
            End If
            
            '���ظ�����
            If rsSend!������� = "7" Then
                .RowHidden(lngRow) = True '�в�ҩ
            ElseIf rsSend!������� = "E" Then
                If Not IsNull(rsSend!���ID) Then
                    .RowHidden(lngRow) = True
                    .Cell(flexcpData, lngRow, COL_�������) = 2 '��ҩ�巨
                ElseIf Val(.TextMatrix(lngRow - 1, COL_���ID)) = rsSend!ID _
                    And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                    .RowHidden(lngRow) = True
                    .Cell(flexcpData, lngRow, COL_�������) = 1 '��ҩ;��
                Else
                    .Cell(flexcpData, lngRow, COL_�������) = 3 '��ҩ�÷�
                End If
            End If
            
            'һ���и�ֵ
            '---------------------------------------------------------------
            .Cell(flexcpData, lngRow, COL_Ӥ��) = CLng(NVL(rsSend!Ӥ��, 0))
            If NVL(rsSend!Ӥ��, 0) = 0 Then
                .TextMatrix(lngRow, COL_Ӥ��) = "����"
            Else
                .TextMatrix(lngRow, COL_Ӥ��) = "Ӥ��" & rsSend!Ӥ��
                .ColHidden(COL_Ӥ��) = False '��Ӥ��ҽ��ʱ����ʾ
            End If
            .TextMatrix(lngRow, COL_����) = rsSend!����
            If InStr(str���� & ",", "," & rsSend!���� & ",") = 0 Then
                If str���� <> "" Then .ColHidden(COL_����) = False
                str���� = str���� & "," & rsSend!����
            End If
            
            .TextMatrix(lngRow, COL_����ID) = rsSend!����ID
            .TextMatrix(lngRow, COL_��ҳID) = rsSend!��ҳID
            .TextMatrix(lngRow, COL_����) = rsSend!����
            .TextMatrix(lngRow, col_�Ա�) = NVL(rsSend!�Ա�)
            .TextMatrix(lngRow, COL_����) = NVL(rsSend!����)
            .TextMatrix(lngRow, COL_����) = NVL(rsSend!����)
            .TextMatrix(lngRow, COL_סԺ��) = NVL(rsSend!סԺ��)
            .TextMatrix(lngRow, COL_����) = NVL(rsSend!����)
            .TextMatrix(lngRow, COL_�ѱ�) = NVL(rsSend!�ѱ�)
            
            .TextMatrix(lngRow, COL_ID) = rsSend!ID
            .TextMatrix(lngRow, COL_���ID) = NVL(rsSend!���ID)
            .TextMatrix(lngRow, COL_�������) = rsSend!�������
            .TextMatrix(lngRow, COL_������ĿID) = rsSend!������ĿID
            .TextMatrix(lngRow, COL_ҽ����Ч) = IIF(rsSend!ҽ����Ч = 0, "����", "����")
            .Cell(flexcpData, lngRow, COL_ҽ����Ч) = Val(rsSend!ҽ����Ч)
            
            .TextMatrix(lngRow, col_ҽ������) = NVL(rsSend!ҽ������)
            .Cell(flexcpData, lngRow, col_ҽ������) = CStr(NVL(rsSend!������Ŀ)) '������ʾ�Ƽ�ҽ��
            
            .TextMatrix(lngRow, COL_ҽ������) = NVL(rsSend!ҽ������)
            .Cell(flexcpData, lngRow, COL_ҽ������) = CStr(NVL(rsSend!ժҪ))
            
            .TextMatrix(lngRow, COL_ִ��ʱ��) = NVL(rsSend!ִ��ʱ�䷽��)
            If Not IsNull(rsSend!��ʼִ��ʱ��) Then
                .Cell(flexcpData, lngRow, COL_ִ��ʱ��) = CStr(Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss"))
            End If
            
            .TextMatrix(lngRow, COL_Ƶ��) = NVL(rsSend!ִ��Ƶ��)
            .TextMatrix(lngRow, COL_���˲���ID) = NVL(rsSend!���˲���ID)
            .TextMatrix(lngRow, COL_���˿���ID) = NVL(rsSend!���˿���id)
            .TextMatrix(lngRow, COL_��������ID) = NVL(rsSend!��������id)
            .TextMatrix(lngRow, COL_����ҽ��) = NVL(rsSend!����ҽ��)
            
            .TextMatrix(lngRow, COL_�Ƽ�����) = NVL(rsSend!�Ƽ�����, 0)
            .TextMatrix(lngRow, COL_ִ������ID) = NVL(rsSend!ִ������, 0)
            .TextMatrix(lngRow, COL_ִ�б��) = NVL(rsSend!ִ�б��, 0)
            .TextMatrix(lngRow, COL_ִ�з���) = NVL(rsSend!ִ�з���, 0)
            .TextMatrix(lngRow, COL_��������) = NVL(rsSend!��������, 0)
            .TextMatrix(lngRow, COL_��ҩ����) = NVL(rsSend!��ҩ����)
            .TextMatrix(lngRow, COL_����ҽ��ID) = NVL(rsSend!����ҽ��ID, 0)
            'ҽ��״̬���ڷ���ǰ��δУ�Ե��Ƚ����Զ�У��
            .TextMatrix(lngRow, COL_ҽ��״̬) = rsSend!ҽ��״̬
            If rsSend!ҽ��״̬ = 1 Then
                .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = BackColorNew 'ǳ��ɫ
            End If
                                                
            '��ʾ��Ҫִ�п���
            .TextMatrix(lngRow, COL_ִ�п���) = NVL(rsSend!ִ�п���)
            
            '��ʾ����ִ�п���
            If rsSend!������� = "E" And IsNull(rsSend!���ID) Then
                If InStr(",7,E,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                    '��ҩ�÷�
                    .TextMatrix(lngRow, COL_����ִ��) = NVL(rsSend!ִ�п���)
                    .Cell(flexcpData, lngRow, COL_����ִ��) = CStr(NVL(rsSend!ִ�п���))
                ElseIf InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                    '��ҩ;��
                    For j = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, COL_���ID)) = rsSend!ID Then
                            .TextMatrix(j, COL_����ִ��) = NVL(rsSend!ִ�п���)
                            .Cell(flexcpData, j, COL_����ִ��) = CStr(NVL(rsSend!ִ�п���))
                        Else
                            Exit For
                        End If
                    Next
                End If
            End If
            
            .TextMatrix(lngRow, COL_ִ�п���ID) = NVL(rsSend!ִ�п���ID)
            If mblnAutoVerify Then .TextMatrix(lngRow, COL_�¿�����ʱ��) = Format(rsSend!�¿�����ʱ��, "yyyy-MM-dd HH:mm:ss")
            If InStr(",4,5,6,7,", "," & rsSend!������� & ",") = 0 Then .TextMatrix(lngRow, COL_���㷽ʽ) = NVL(rsSend!���㷽ʽ, 0)
            .TextMatrix(lngRow, COL_ִ�а���) = NVL(rsSend!ִ�а���, 0)
            .TextMatrix(lngRow, COL_��������) = NVL(rsSend!��������, 0)
            .TextMatrix(lngRow, COL_����ʱ��) = Format(rsSend!����ʱ��, "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(lngRow, COL_��ʼʱ��) = Format(NVL(rsSend!��ʼִ��ʱ��), "yyyy-MM-dd HH:mm:ss")
                                                
            '��ȡҩƷ�����Ϣ
            '---------------------------------------------------------------
            If InStr(",5,6,7", rsSend!�������) > 0 Then
                '��������ж�
                If gbln����ҩƷ�ֿ����� Then
                    strTmp = ""
                    Select Case cboDrugType.ListIndex
                    Case 1
                        If rsSend!������� & "" <> "����ҩ" Then strTmp = "1"
                    Case 2
                        If InStr(",����ҩ,����I��,", "," & rsSend!������� & ",") = 0 Then strTmp = "1"
                    Case 3
                        If InStr(",����ҩ,����ҩ,����I��,", "," & rsSend!������� & ",") > 0 Then strTmp = "1"
                    End Select
                    
                    If strTmp <> "" Then
                        lngDel���ID = NVL(rsSend!���ID, 0)
                        Call DeleteCurRow(lngRow, rsSend!���ID)
                        lng��С���� = 0: GoTo NextLoop
                    End If
                    .TextMatrix(lngRow, COL_�������) = NVL(rsSend!�������, "��")
                End If
                
                Set rsDrug = New ADODB.Recordset
                '�Ȱ���ͣ��ҩƷ,��ȷ��Ҫ���͵�ҽ���ټ��ͣ��
                Set rsDrug = GetDrugInfo(rsSend!������ĿID, NVL(rsSend!�շ�ϸĿID, 0), NVL(rsSend!ִ�п���ID, 0), 2, False)
                '�ų���ǰִ�п�����û�д洢�Ĺ��
                If NVL(rsSend!ִ�п���ID, 0) <> 0 And rsDrug.RecordCount > 1 And InStr("," & gstr��Һ�������� & ",", "," & NVL(rsSend!ִ�п���ID, 0) & ",") > 0 Then
                    strIDs = ""
                    Do While Not rsDrug.EOF
                        strIDs = strIDs & "," & rsDrug!ҩƷID
                        rsDrug.MoveNext
                    Loop
                    strSQL = "Select /*+ rule*/" & vbNewLine & _
                            "Distinct �շ�ϸĿid" & vbNewLine & _
                            "From �շ�ִ�п���" & vbNewLine & _
                            "Where (��������id = [2] Or ��������id Is Null) And ִ�п���ID = [3] And" & vbNewLine & _
                            "      �շ�ϸĿid In (Select Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)))"
    
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strIDs, 2), Val(rsSend!��������id & ""), Val(rsSend!ִ�п���ID & ""))
                    If rsDrug.RecordCount > 0 Then rsDrug.MoveFirst
                    strIDs = ""
                    Do While Not rsDrug.EOF
                        rsTmp.Filter = "�շ�ϸĿID=" & rsDrug!ҩƷID
                        If rsTmp.RecordCount = 0 Then
                           strIDs = strIDs & " or ҩƷID<>" & rsDrug!ҩƷID
                        End If
                        rsDrug.MoveNext
                    Loop
                    strIDs = Mid(strIDs, 4)
                    If strIDs <> "" Then rsDrug.Filter = strIDs
                    If rsDrug.RecordCount > 0 Then rsDrug.MoveFirst
                End If
                If rsDrug.EOF Then
                    'ҩƷû�ж�Ӧ�Ĺ����Ϣ
                    'ɾ����ǰ��(�������),��������һҽ��
                    Call DeleteDrugRow(rsSend, lngRow, lngDel���ID)
                    lng��С���� = 0: GoTo NextLoop
                ElseIf rsDrug.RecordCount > 1 Then
                    'Ѱ�Һ��ʵĹ��
                    Call SeekMatchDrug(rsSend, rsDrug, 0, vBookMark, strTmp)
                    If vBookMark <> 0 Then
                        rsDrug.Bookmark = vBookMark
                    Else
                        rsDrug.MoveFirst
                    End If
                    .Cell(flexcpData, lngRow, COL_���) = strTmp '��ѡ��Ĺ��
                    '���ȫ��(ָ��)���ͣ�õ�ҩƷ
                    If .Cell(flexcpData, lngRow, COL_���) = "" Then
                        Call DeleteDrugRow(rsSend, lngRow, lngDel���ID)
                        lng��С���� = 0: GoTo NextLoop
                    End If
                Else
                    '���ȫ��(ָ��)���ͣ�õ�ҩƷ������ҩƷҽ����ȷ��Ҫ����ʱ��ɾ������ʾ
                    If Not (rsSend!ҽ����Ч = 0 And InStr(",5,6,", rsSend!�������) > 0) _
                        And Not (NVL(rsDrug!����ʱ��, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", NVL(rsDrug!�������, 0)) > 0) Then
                        Call DeleteDrugRow(rsSend, lngRow, lngDel���ID)
                        lng��С���� = 0: GoTo NextLoop
                    ElseIf Val(rsSend!ҽ����Ч & "") = 0 And InStr(",5,6,", rsSend!�������) > 0 And Val(rsSend!ִ�п���ID & "") <> 0 And Val(rsSend!�շ�ϸĿID & "") <> 0 Then '����շ�ִ�п����Ƿ�ı�
                        strSQL = "Select 1 From �շ�ִ�п��� Where �շ�ϸĿid = [1] And Nvl(������Դ, 2) = 2 And Nvl(��������ID, [2]) = [2] And ִ�п���ID = [3]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsSend!�շ�ϸĿID & ""), Val(rsSend!��������ID & ""), Val(rsSend!ִ�п���ID & ""))
                        If rsTmp.EOF Then
                            Call DeleteDrugRow(rsSend, lngRow, lngDel���ID)
                            lng��С���� = 0: GoTo NextLoop
                        End If
                    End If
                End If
                .TextMatrix(lngRow, COL_���) = rsDrug!���� & IIF(Not IsNull(rsDrug!����), "(" & rsDrug!���� & ")", "") & IIF(Not IsNull(rsDrug!���), " " & rsDrug!���, "")
                .TextMatrix(lngRow, COL_�շ�ϸĿID) = rsDrug!ҩƷID
                .TextMatrix(lngRow, COL_���) = Format(NVL(rsDrug!���, 0), "0.00000") '��סԺ��װ
                .TextMatrix(lngRow, COL_����ϵ��) = NVL(rsDrug!����ϵ��, 1)
                .TextMatrix(lngRow, COL_סԺ��װ) = NVL(rsDrug!סԺ��װ, 1)
                .TextMatrix(lngRow, COL_סԺ��λ) = NVL(rsDrug!סԺ��λ)
                .TextMatrix(lngRow, COL_�ɷ����) = NVL(rsSend!�ɷ����, NVL(rsDrug!�ɷ����, 0))
                .TextMatrix(lngRow, COL_ҩ������) = NVL(rsDrug!ҩ������, 0)
                .TextMatrix(lngRow, COL_�Ƿ���) = NVL(rsDrug!�Ƿ���, 0)
                
                '�Ƿ����δȷ������Ʒ��ҩƷ
                If .Cell(flexcpData, lngRow, COL_���) <> "" Then
                    .Cell(flexcpForeColor, lngRow, COL_���) = vbBlue 'ͻ����ʾ
                    blnƷ��ҩƷ = True
                End If
            End If
                                                                    
            '���㷢�ʹ�����ִ�еķֽ�ʱ�䣬����
            '---------------------------------------------------------------
            If rsSend!ҽ����Ч = 0 Then
                '����---------------------------------------------
                If rsSend!������� = "7" Then
                    '��ǰҽ���ķ��ͼ���ʱ���
                    Call Calc��������ʱ��(dbl����, lng����, str�ֽ�ʱ��, strEnd, rsSend, rsDrug, bln�ڷ�)
                    If str�ֽ�ʱ�� = "" Then
                        If rsSend!ҽ��״̬ = 1 Then '��У��
                            lng��С���� = 0
                        Else
                            '�޷��ֽ�ʱ��(�类��ͣ��)
                            lngDel���ID = rsSend!���ID
                            Call DeleteCurRow(lngRow, rsSend!���ID)
                            lng��С���� = 0: GoTo NextLoop
                        End If
                    ElseIf Not (NVL(rsDrug!����ʱ��, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", NVL(rsDrug!�������, 0)) > 0) Then
                        'ȷ��Ҫ�������ͣ����ѱ������򲻷����ڲ��˵�ҩƷ
                        Call DeleteDrugRow(rsSend, lngRow, lngDel���ID)
                        lng��С���� = 0: GoTo NextLoop
                    End If
                    .TextMatrix(lngRow, COL_����) = lng����
                    If Len(str�ֽ�ʱ��) > 4000 Then
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = Mid(str�ֽ�ʱ��, 1, InStr(Mid(str�ֽ�ʱ��, 4001), ",") + 3999)
                    Else
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
                    End If
                    If str�ֽ�ʱ�� <> "" Then
                        .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "yyyy-MM-dd HH:mm")
                    End If
                    
                    .TextMatrix(lngRow, COL_����) = NVL(rsSend!��������)
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    .TextMatrix(lngRow, COL_����) = lng���� * NVL(rsSend!�ܸ�����, 1) '�ܸ���
                    
                    .TextMatrix(lngRow, COL_������λ) = "��"
                ElseIf InStr(",5,6,", rsSend!�������) > 0 Then
                    blnReCalc = False
ReCalc:
                    '��ǰҽ���ķ��ͼ���ʱ���
                    Call Calc��������ʱ��(dbl����, lng����, str�ֽ�ʱ��, strEnd, rsSend, rsDrug, bln�ڷ�)
                    If str�ֽ�ʱ�� = "" Then
                        If rsSend!ҽ��״̬ = 1 Then '��У��
                            lng��С���� = 0
                        Else
                            '�޷��ֽ�ʱ��(�类��ͣ��)
                            lngDel���ID = rsSend!���ID
                            Call DeleteCurRow(lngRow, rsSend!���ID)
                            lng��С���� = 0: GoTo NextLoop
                        End If
                    ElseIf Not (NVL(rsDrug!����ʱ��, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", NVL(rsDrug!�������, 0)) > 0) Then
                        'ȷ��Ҫ�������ͣ����ѱ������򲻷����ڲ��˵�ҩƷ
                        Call DeleteDrugRow(rsSend, lngRow, lngDel���ID)
                        lng��С���� = 0: GoTo NextLoop
                    End If
                    .TextMatrix(lngRow, COL_����) = lng����
                    If Len(str�ֽ�ʱ��) > 4000 Then
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = Mid(str�ֽ�ʱ��, 1, InStr(Mid(str�ֽ�ʱ��, 4001), ",") + 3999)
                    Else
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
                    End If
                    If str�ֽ�ʱ�� <> "" Then
                        .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "yyyy-MM-dd HH:mm")
                    End If
                    
                    .TextMatrix(lngRow, COL_����) = FormatEx(NVL(rsSend!��������), 5)
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    .TextMatrix(lngRow, COL_����) = FormatEx(dbl����, 5)
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsDrug!סԺ��λ)
                    If lng���� < lng��С���� Or lng��С���� = 0 Then lng��С���� = lng����
                    
                    '���ж������ѡ��ʱ�����ݿ���Ƿ��㹻�ٴζ�λ���
                    If Not blnReCalc And .Cell(flexcpData, lngRow, COL_���) <> "" _
                        And Val(.TextMatrix(lngRow, COL_����)) > Val(.TextMatrix(lngRow, COL_���)) Then
                        Call SeekMatchDrug(rsSend, rsDrug, Val(.TextMatrix(lngRow, COL_����)), vBookMark)
                        If vBookMark <> 0 Then
                            rsDrug.Bookmark = vBookMark
                            .TextMatrix(lngRow, COL_���) = rsDrug!���� & IIF(Not IsNull(rsDrug!����), "(" & rsDrug!���� & ")", "") & IIF(Not IsNull(rsDrug!���), " " & rsDrug!���, "")
                            .TextMatrix(lngRow, COL_�շ�ϸĿID) = rsDrug!ҩƷID
                            .TextMatrix(lngRow, COL_���) = Format(NVL(rsDrug!���, 0), "0.00000") '��סԺ��װ
                            .TextMatrix(lngRow, COL_����ϵ��) = NVL(rsDrug!����ϵ��, 1)
                            .TextMatrix(lngRow, COL_סԺ��װ) = NVL(rsDrug!סԺ��װ, 1)
                            .TextMatrix(lngRow, COL_סԺ��λ) = NVL(rsDrug!סԺ��λ)
                            .TextMatrix(lngRow, COL_ҩ������) = NVL(rsDrug!ҩ������, 0)
                            .TextMatrix(lngRow, COL_�Ƿ���) = NVL(rsDrug!�Ƿ���, 0)
                            blnReCalc = True: GoTo ReCalc
                        End If
                    End If
                Else
                    'һ����ҩ�İ���С��������(Ӱ���ҩ;���ƷѼ��ϴ�ִ��ʱ��)(������Ŀ����˷�)
                    If .Cell(flexcpData, lngRow, COL_�������) = 1 Then '��ҩ;��
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_���ID)) = rsSend!ID Then
                                If Val(.TextMatrix(j, COL_����)) > lng��С���� Then
                                    .TextMatrix(j, COL_����) = lng��С����
                                    .TextMatrix(j, COL_�ֽ�ʱ��) = Trim�ֽ�ʱ��(lng��С����, .TextMatrix(j, COL_�ֽ�ʱ��))
                                    .TextMatrix(j, COL_�״�ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(0), "yyyy-MM-dd HH:mm")
                                    .TextMatrix(j, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(lng��С���� - 1), "yyyy-MM-dd HH:mm")
                                End If
                            Else
                                Exit For
                            End If
                        Next
                        lng��С���� = 0
                    End If
                    
                    If InStr(",2,3,", .Cell(flexcpData, lngRow, COL_�������)) > 0 Then
                        '��ҩ�巨���÷�Ϊ����
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                    Else
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                    End If
                    .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                    If .Cell(flexcpData, lngRow, COL_�������) = 3 Then '��ҩ�÷�
                        .TextMatrix(lngRow, COL_������λ) = "��"
                    End If
                    
                    .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                    .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                    .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                End If
                If mlngRefModld = 1 Then
                    If (dkpExecTime.value >= rsSend!ִ����ֹʱ�� Or dkpExecTime.value <= rsSend!�ϴ�ִ��ʱ��) Then
                        '����ҽ��ɾ����Χ��ı���ҽ��
                        lngDel���ID = rsSend!���ID
                        Call DeleteCurRow(lngRow, rsSend!���ID)
                        lng��С���� = 0: GoTo NextLoop
                    End If
                End If
            Else
                '����---------------------------------------------
                If rsSend!������� = "7" Then
                    .TextMatrix(lngRow, COL_����) = NVL(rsSend!�ܸ�����, 0)
                    If Not IsNull(rsSend!ִ��ʱ�䷽��) Or NVL(rsSend!�����λ) = "����" Then
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = Calc�����ֽ�ʱ��(rsSend!�ܸ�����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", NVL(rsSend!ִ��ʱ�䷽��), rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                        .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(rsSend!�ܸ����� - 1), "yyyy-MM-dd HH:mm")
                    Else
                        '�޷ֽ�ʱ��(��������δ����ִ��ʱ����޷��ֽ�)
                        '��¼���÷���ʱ��(��ҽ����ʼִ��ʱ��)
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                    End If
                    .TextMatrix(lngRow, COL_����) = NVL(rsSend!��������) '����
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    .TextMatrix(lngRow, COL_����) = NVL(rsSend!�ܸ�����, 0) '����
                    .TextMatrix(lngRow, COL_������λ) = "��"
                ElseIf InStr(",5,6,", rsSend!�������) > 0 Then
                    '����������ҩ����
                    If NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Then
                        lng���� = 1 '����Ϊһ���Ե�����ҩƷ
                    ElseIf NVL(rsSend!����, 0) <> 0 And Not IsNull(rsSend!ִ��Ƶ��) Then
                        'һ��Ƶ�����ڵĴ���
                        If rsSend!�����λ = "��" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / 7))
                        ElseIf rsSend!�����λ = "��" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / rsSend!Ƶ�ʼ��))
                        ElseIf rsSend!�����λ = "Сʱ" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / rsSend!Ƶ�ʼ��) * 24)
                        ElseIf rsSend!�����λ = "����" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / rsSend!Ƶ�ʼ��) * (24 * 60))
                        End If
                    Else
                        '�ɷ���ҩƷʱ,�������Ե����ı��������ҩ;���Ĵ���,���ɷ�����һ����ʹ��ҩƷʱ���������ԣ����������ϵ����ֵȡ�����ı��������ҩ;���Ĵ�����
                        '����һ��Ƶ�����ڵĴ�������
                        If NVL(rsSend!�ɷ����, NVL(rsDrug!�ɷ����, 0)) = 0 And NVL(rsSend!��������, 0) <> 0 Then
                            lng���� = IntEx(rsSend!�ܸ����� * rsDrug!����ϵ�� / rsSend!��������)
                        ElseIf (NVL(rsSend!�ɷ����, NVL(rsDrug!�ɷ����, 0)) = 1 Or NVL(rsSend!�ɷ����, NVL(rsDrug!�ɷ����, 0)) = 2) And NVL(rsSend!��������, 0) <> 0 Then
                            lng���� = IntEx(rsSend!�ܸ����� / IntEx(rsSend!�������� / rsDrug!����ϵ��))
                        Else
                            lng���� = NVL(rsSend!Ƶ�ʴ���, 0)
                        End If
                    End If
                    If Not IsNull(rsSend!ִ��ʱ�䷽��) Or NVL(rsSend!�����λ) = "����" Then
                        str�ֽ�ʱ�� = Calc�����ֽ�ʱ��(lng����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", NVL(rsSend!ִ��ʱ�䷽��), rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                        If str�ֽ�ʱ�� <> "" Then
                            If Len(str�ֽ�ʱ��) > 4000 Then
                                .TextMatrix(lngRow, COL_�ֽ�ʱ��) = Mid(str�ֽ�ʱ��, 1, InStr(Mid(str�ֽ�ʱ��, 4001), ",") + 3999)
                            Else
                                .TextMatrix(lngRow, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
                            End If
                            .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "yyyy-MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "yyyy-MM-dd HH:mm")
                        End If
                    Else
                        '�޷ֽ�ʱ��(һ��������δ����ִ��ʱ����޷��ֽ�)
                        '��¼���÷���ʱ��(��ҽ����ʼִ��ʱ��)
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                    End If
                    .TextMatrix(lngRow, COL_����) = lng����
                    .TextMatrix(lngRow, COL_����) = FormatEx(NVL(rsSend!��������), 5)
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    .TextMatrix(lngRow, COL_����) = FormatEx(rsSend!�ܸ����� / rsDrug!סԺ��װ, 5) '��סԺ��λ��ʾ
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsDrug!סԺ��λ)
                    
                    If lng���� < lng��С���� Or lng��С���� = 0 Then lng��С���� = lng����
                Else
                    '������һ����ҩ�İ���С��������(Ӱ���ҩ;���Ʒ�)
                    If .Cell(flexcpData, lngRow, COL_�������) = 1 Then '��ҩ;��
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_���ID)) = rsSend!ID Then
                                If Val(.TextMatrix(j, COL_����)) > lng��С���� Then
                                    .TextMatrix(j, COL_����) = lng��С����
                                    If .TextMatrix(j, COL_�ֽ�ʱ��) <> "" Then
                                        .TextMatrix(j, COL_�ֽ�ʱ��) = Trim�ֽ�ʱ��(lng��С����, .TextMatrix(j, COL_�ֽ�ʱ��))
                                        .TextMatrix(j, COL_�״�ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(0), "yyyy-MM-dd HH:mm")
                                        .TextMatrix(j, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(lng��С���� - 1), "yyyy-MM-dd HH:mm")
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        Next
                        lng��С���� = 0
                    End If
                    
                    .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����) '���������
                    .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                    If .Cell(flexcpData, lngRow, COL_�������) = 3 Then '��ҩ�÷�
                        .TextMatrix(lngRow, COL_������λ) = "��"
                    End If
                    
                    .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                    .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��)
                    .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                    .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                End If
            End If
            
            '������Ŀ�Ľ��:���ڲ鿴�����ʱ���
            '---------------------------------------------------------------
            cur��� = 0
            Call LoadAdvicePrice(lngRow, cur���, rsDrug)
            .TextMatrix(lngRow, COL_���) = Format(cur���, gstrDec)
            
            '�����ʱ��һЩ�����ۼ���ʾ���,��ҩ;��,�÷�,ִ�п���,ִ������
            '---------------------------------------------------------------
            If InStr(",1,3,", Val(.Cell(flexcpData, lngRow, COL_�������))) > 0 Then '��ҩ;������ҩ�÷�
                cur��� = 0
                lngTmp = .FindRow(CStr(rsSend!ID), , COL_���ID)
                
                If .Cell(flexcpData, lngRow, COL_�������) = 1 Then '��ҩ;��
                    'һ����ҩʱ,��ҩ;���Ľ���ۼ���ʾ�ڵ�һ����ҩ��
                    .TextMatrix(lngTmp, COL_���) = Format(Val(.TextMatrix(lngTmp, COL_���)) + Val(.TextMatrix(lngRow, COL_���)), gstrDec)
                    
                    '��ʾ��ҩ;��,ִ������
                    For j = lngTmp To lngRow - 1
                        strTmp = ""
                        If Val(.TextMatrix(j, COL_ִ������ID)) = 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                            If Val(.TextMatrix(j, COL_ִ�б��)) = 2 Then
                                strTmp = "��ȡҩ"
                            Else
                                strTmp = "�Ա�ҩ"
                            End If
                        ElseIf Val(.TextMatrix(j, COL_ִ������ID)) <> 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) = 5 Then
                            strTmp = "��Ժ��ҩ"
                        Else
                            strTmp = IIF(Val(.TextMatrix(j, COL_ִ�б��)) = 1, "��ȡҩ", "")
                        End If
                        .TextMatrix(j, COL_ִ������) = strTmp
                        .TextMatrix(j, COL_�÷�) = rsSend!������Ŀ
                    Next
                Else
                    'ҩƷ��ִ������
                    strTmp = ""
                    If Val(.TextMatrix(lngTmp, COL_ִ������ID)) = 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                        If Val(.TextMatrix(lngTmp, COL_ִ�б��)) = 2 Then
                            strTmp = "��ȡҩ"
                        Else
                            strTmp = "�Ա�ҩ"
                        End If
                    ElseIf Val(.TextMatrix(lngTmp, COL_ִ������ID)) <> 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) = 5 Then
                        strTmp = "��Ժ��ҩ"
                    Else
                        strTmp = IIF(Val(.TextMatrix(lngTmp, COL_ִ�б��)) = 1, "��ȡҩ", "")
                    End If
                    
                    '��ҩ�÷�,�巨
                    str�÷� = rsSend!������Ŀ
                    If Val(.Cell(flexcpData, lngRow - 1, COL_�������)) = 2 Then
                        str�÷� = str�÷� & "|" & sys.RowValue("������ĿĿ¼", Val(.TextMatrix(lngRow - 1, COL_������ĿID)), "����")
                    End If
                    For j = lngTmp To lngRow
                        .TextMatrix(j, COL_�÷�) = str�÷� '������д�շ���¼
                        cur��� = cur��� + Val(.TextMatrix(j, COL_���))
                    Next
                    .TextMatrix(lngRow, COL_���) = Format(cur���, gstrDec)
                    '��ʾִ������
                    .TextMatrix(lngRow, COL_ִ������) = strTmp
                    '��ʾ�䷽ִ�п���
                    .TextMatrix(lngRow, COL_ִ�п���) = .TextMatrix(lngTmp, COL_ִ�п���)
                End If
                
                'ʹ���ҽ��ѡ��״̬��ͬ(��Ϊ����ԭ��)
                For j = lngTmp To lngRow
                    If .Cell(flexcpData, j, COL_ѡ��) <> 0 Then
                        Call RowSelectSame(j, COL_ѡ��)
                        Exit For 'һ����ֹ,ȫ����ֹ
                    End If
                Next
                If j > lngRow Then
                    For j = lngRow To lngTmp Step -1
                        If InStr(",5,6,7,", .TextMatrix(j, COL_�������)) > 0 Then
                            If .Cell(flexcpPicture, j, COL_ѡ��) Is Nothing Then
                                Call RowSelectSame(j, COL_ѡ��)
                                Exit For '���ѡ,ȫ����ѡ
                            End If
                        End If
                    Next
                End If
            End If
            
            'ҩƷ�����:�Ա�ҩ�����
            '---------------------------------------------------------------
            If InStr(",5,6,7,", rsSend!�������) > 0 And NVL(rsSend!ִ������, 0) <> 5 Then
                Call CheckStock(lngRow, bln�����ʾ, blnʱ����ʾ, blnĬ�Ϸ���)
                Call CheckDrugStorage(lngRow, bln�洢�ⷿ��ʾ)
            End If
            
            '��������
            '---------------------------------------------------------------
            '���˼������ָ�
            If rsSend!����ID <> lng����ID Then
                lng������ = lng������ + 1
                If lng����ID <> 0 Then
                    For j = lngRow - 1 To .FixedRows Step -1
                        If Not .RowHidden(j) Then
                            .CellBorderRange j, .FixedCols, j, .Cols - 1, vbBlack, 0, 0, 0, 2, 0, 0
                            Exit For
                        End If
                    Next
                End If
            End If
            lng����ID = rsSend!����ID

NextLoop:           '---------------------------------------------------------------
            If blnOnePati Then Progress = i / rsSend.RecordCount * 100
            rsSend.MoveNext
        Next
        .Redraw = flexRDDirect
    End With
    
    If blnOnePati Then Progress = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadAdviceSendOther(ByVal blnOnePati As Boolean, ByVal strEnd As String, ByVal rsSend As ADODB.Recordset, _
    ByRef lng������ As Long, ByRef str���� As String, ByRef lng������ As Long, ByRef lng����ID As Long) As Boolean
'���ܣ�����������ȡ����ʾҪ���͵�ҽ���嵥
'������strEnd=���͵��Ľ���ʱ��(yyyy-MM-dd HH:mm:ss),����û��
'���أ�lng������=�д�����ҽ���Ĳ�����
'      str����=���в��˵�ǰ���Ҵ�
'      lng������=�����Ƿ���ʾ������
'˵����ע��CellData�д�ŵ��и�������
'   RowData��0-δ���͵�,-1-�ѳɹ����͵�
'   COL_ѡ��0-������ѡ���,1-��ֹ�ı�ѡ��״̬��
'   COL_Ӥ�������Ӥ�����
'   COL_������𣺴������������ƣ�����ʾ�Ƽ�ҽ��ʱʹ��
'   COL_ҽ�����ݣ����������Ŀ���ƻ�걾��λ������ʾ�Ƽ�ҽ��ʱʹ��
'   COL_�ֽ�ʱ��:�����޷ֽ�ʱ��ʱ,��ŷ��÷���ʱ��
'   COL_�״�ʱ��,COL_ĩ��ʱ�䣺��ų����Գ�������ĩ��ִ��ʱ��
'   COL_Ƶ�ʣ�1-"һ����"������2-"������"����
'   COL_ִ�п��ң����ԭִ�п�������
'   COL_ִ�п���ID���Ƿ������ִ�п���
'   COL_����żƼ������Ƿ�����
   
    Dim strTmp As String, i As Long, j As Long, k As Long
    Dim datBegin As Date, datEnd As Date, strPause As String
    Dim lng���� As Long, dbl���� As Double, bln�ɼ����� As Boolean
    Dim str�ֽ�ʱ�� As String, str�״�ʱ�� As String, strĩ��ʱ�� As String
    Dim lngDelҽ��ID As Long, lngRow As Long
    Dim cur��� As Currency
    Dim strͣ�� As String
             
        
    '���㲢��ʾ�����嵥
    '----------------------------------------------------------------------------------------------------------
    With vsAdvice
        .Redraw = flexRDNone
        For i = 1 To rsSend.RecordCount
            If NVL(rsSend!���ID, 0) = lngDelҽ��ID And lngDelҽ��ID <> 0 Then
                GoTo NextLoop '�����ϻ��������,��Ѫ��Ŀ�е�һ�������Ѿ����ܷ���,�����鲻�ܷ���
            Else
                lngDelҽ��ID = 0
            End If
            
            bln�ɼ����� = False
            
            '���뵱ǰ��
            .Rows = .Rows + 1: lngRow = .Rows - 1
            .Cell(flexcpPictureAlignment, lngRow, COL_ѡ��) = 4
            If Not ("Z" = "" & rsSend!������� And Val("" & rsSend!��������) = 5) Then
                Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
            End If
            
            If mlngRefModld = 1 Then
                If Format(rsSend!�ϴ�ִ��ʱ��, "YYYY-MM-DD") >= Format(mdatCurr, "YYYY-MM-DD") Then '�����ѷ��͵�(�������ܷ��͵�����)
                    .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HA08000           '����
                End If
            End If
            
            '����:��������,��������,��鲿λ����,������Ŀ,��Ѫ;��
            .RowHidden(lngRow) = Not IsNull(rsSend!���ID)
            
            'һ���и�ֵ
            '---------------------------------------------------------------
            .Cell(flexcpData, lngRow, COL_Ӥ��) = CLng(NVL(rsSend!Ӥ��, 0))
            If NVL(rsSend!Ӥ��, 0) = 0 Then
                .TextMatrix(lngRow, COL_Ӥ��) = "����"
            Else
                .TextMatrix(lngRow, COL_Ӥ��) = "Ӥ��" & rsSend!Ӥ��
                .ColHidden(COL_Ӥ��) = False '��Ӥ��ҽ��ʱ����ʾ
            End If
            .TextMatrix(lngRow, COL_����) = rsSend!����
            If InStr(str���� & ",", "," & rsSend!���� & ",") = 0 Then
                If str���� <> "" Then .ColHidden(COL_����) = False
                str���� = str���� & "," & rsSend!����
            End If
            
            .TextMatrix(lngRow, COL_����ID) = rsSend!����ID
            .TextMatrix(lngRow, COL_��ҳID) = rsSend!��ҳID
            .TextMatrix(lngRow, COL_����) = rsSend!����
            .TextMatrix(lngRow, col_�Ա�) = NVL(rsSend!�Ա�)
            .TextMatrix(lngRow, COL_����) = NVL(rsSend!����)
            .TextMatrix(lngRow, COL_����) = NVL(rsSend!����)
            .TextMatrix(lngRow, COL_סԺ��) = NVL(rsSend!סԺ��)
            .TextMatrix(lngRow, COL_����) = NVL(rsSend!����)
            .TextMatrix(lngRow, COL_�ѱ�) = NVL(rsSend!�ѱ�)
            
            .TextMatrix(lngRow, COL_ID) = rsSend!ID
            .TextMatrix(lngRow, COL_���ID) = NVL(rsSend!���ID)
            .TextMatrix(lngRow, COL_�������) = "" & rsSend!�������
            .TextMatrix(lngRow, COL_������ĿID) = Val("" & rsSend!������ĿID)
            .TextMatrix(lngRow, COL_ҽ����Ч) = IIF(rsSend!ҽ����Ч = 0, "����", "����")
            .Cell(flexcpData, lngRow, COL_ҽ����Ч) = Val(rsSend!ҽ����Ч)
                        
            .TextMatrix(lngRow, COL_�걾��λ) = NVL(rsSend!�걾��λ)
            .TextMatrix(lngRow, COL_��鷽��) = NVL(rsSend!��鷽��)
            .TextMatrix(lngRow, COL_ִ�б��) = NVL(rsSend!ִ�б��, 0)
            
            '������ʾ�Ƽ�ҽ��
            .Cell(flexcpData, lngRow, COL_�������) = CStr(NVL(rsSend!�������))
            .TextMatrix(lngRow, col_ҽ������) = NVL(rsSend!ҽ������)
            .Cell(flexcpData, lngRow, col_ҽ������) = CStr(NVL(rsSend!������Ŀ))
            
            .TextMatrix(lngRow, COL_ҽ������) = NVL(rsSend!ҽ������)
            .Cell(flexcpData, lngRow, COL_ҽ������) = CStr(NVL(rsSend!ժҪ))
            
            .TextMatrix(lngRow, COL_ִ��ʱ��) = NVL(rsSend!ִ��ʱ�䷽��)
            If Not IsNull(rsSend!��ʼִ��ʱ��) Then
                .Cell(flexcpData, lngRow, COL_ִ��ʱ��) = CStr(Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss"))
            End If
            
            .TextMatrix(lngRow, COL_Ƶ��) = NVL(rsSend!ִ��Ƶ��)
            
            .TextMatrix(lngRow, COL_���˲���ID) = NVL(rsSend!���˲���ID)
            .TextMatrix(lngRow, COL_���˿���ID) = NVL(rsSend!���˿���id)
            .TextMatrix(lngRow, COL_��������ID) = NVL(rsSend!��������id)
            .TextMatrix(lngRow, COL_����ҽ��) = NVL(rsSend!����ҽ��)
            
            .TextMatrix(lngRow, COL_�Ƽ�����) = NVL(rsSend!�Ƽ�����, 0)
            .TextMatrix(lngRow, COL_��������) = NVL(rsSend!��������)
            .TextMatrix(lngRow, COL_ִ��Ƶ��) = NVL(rsSend!ִ��Ƶ��)
            .TextMatrix(lngRow, COL_ִ������ID) = NVL(rsSend!ִ������, 0)
            .TextMatrix(lngRow, COL_������־) = NVL(rsSend!������־, 0)
            If mblnAutoVerify Then .TextMatrix(lngRow, COL_�¿�����ʱ��) = Format(rsSend!�¿�����ʱ��, "yyyy-MM-dd HH:mm:ss")
            If InStr(",4,5,6,7,", "," & rsSend!������� & ",") = 0 Then .TextMatrix(lngRow, COL_���㷽ʽ) = NVL(rsSend!���㷽ʽ, 0)
            .TextMatrix(lngRow, COL_ִ�а���) = NVL(rsSend!ִ�а���, 0)
            .TextMatrix(lngRow, COL_��������) = NVL(rsSend!��������, 0)
            .TextMatrix(lngRow, COL_����ʱ��) = Format(rsSend!����ʱ��, "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(lngRow, COL_��ʼʱ��) = Format(NVL(rsSend!��ʼִ��ʱ��), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(lngRow, COL_�������) = NVL(rsSend!�������, 0)
            .TextMatrix(lngRow, COL_ִ�з���) = NVL(rsSend!ִ�з���, 0)
            .TextMatrix(lngRow, COL_��ҩ����) = NVL(rsSend!��ҩ����)
            .TextMatrix(lngRow, COL_����ҽ��ID) = NVL(rsSend!����ҽ��ID, 0)
            'ҽ��״̬���ڷ���ǰ��δУ�Ե��Ƚ����Զ�У��
            .TextMatrix(lngRow, COL_ҽ��״̬) = rsSend!ҽ��״̬
            If rsSend!ҽ��״̬ = 1 Then
                .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = BackColorNew 'ǳ��ɫ
            End If
            
            '�ɼ���ʽ�Ĺ�����һ���ĵ�һ��������ͬ
            If NVL(rsSend!�������) = "E" And IsNull(rsSend!���ID) _
                And .TextMatrix(lngRow - 1, COL_�������) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_���ID)) = rsSend!ID Then
                j = .FindRow(CStr(rsSend!ID), .FixedRows, COL_���ID)
                If j <> -1 Then
                    .TextMatrix(lngRow, COL_�Թܱ���) = .TextMatrix(j, COL_�Թܱ���)
                End If
            Else
                .TextMatrix(lngRow, COL_�Թܱ���) = NVL(rsSend!�Թܱ���)
            End If
            
            '��������ҽ��
            .TextMatrix(lngRow, COL_�շ�ϸĿID) = NVL(rsSend!�շ�ϸĿID, 0)
            .TextMatrix(lngRow, COL_�Ƿ���) = NVL(rsSend!�Ƿ���, 0)
            .TextMatrix(lngRow, COL_��������) = NVL(rsSend!��������, 0)
            
            '����Ŀִ�п�����ʾ
            If IsNull(rsSend!���ID) And rsSend!������� = "E" _
                And Val(.TextMatrix(lngRow - 1, COL_���ID)) = rsSend!ID Then
                '�ɼ�������ʾΪ������Ŀ��ִ�п���
                bln�ɼ����� = True
                .TextMatrix(lngRow, COL_ִ�п���) = .TextMatrix(lngRow - 1, COL_ִ�п���)
                .Cell(flexcpData, lngRow, COL_ִ�п���) = .Cell(flexcpData, lngRow - 1, COL_ִ�п���)
            Else
                .TextMatrix(lngRow, COL_ִ�п���) = NVL(rsSend!ִ�п���)
                .Cell(flexcpData, lngRow, COL_ִ�п���) = CStr(NVL(rsSend!ִ�п���))
            End If
            
            '������Ŀִ�п�����ʾ
            If NVL(rsSend!�������) = "E" And IsNull(rsSend!���ID) _
                And .TextMatrix(lngRow - 1, COL_�������) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_���ID)) = rsSend!ID Then
                '�ɼ���ʽ�ڵ�ǰ����ʾ����ִ�п���
                .TextMatrix(lngRow, COL_����ִ��) = NVL(rsSend!ִ�п���)
                .Cell(flexcpData, lngRow, COL_����ִ��) = CStr(NVL(rsSend!ִ�п���))
            ElseIf NVL(rsSend!�������) = "G" And Not IsNull(rsSend!���ID) Then
                '������������������ʾ����ִ�п���
                j = .FindRow(CStr(rsSend!���ID), .FixedRows, COL_ID)
                If j <> -1 Then
                    .TextMatrix(j, COL_����ִ��) = NVL(rsSend!ִ�п���)
                    .Cell(flexcpData, j, COL_����ִ��) = CStr(NVL(rsSend!ִ�п���))
                End If
            End If
            
            .TextMatrix(lngRow, COL_ִ�п���ID) = NVL(rsSend!ִ�п���ID)
                            
            '���㷢�ʹ�����ִ�еķֽ�ʱ�䣬����
            '---------------------------------------------------------------
            If rsSend!ҽ����Ч = 0 Then
                '����---------------------------------------------
                If (IsNull(rsSend!���ID) And Not bln�ɼ�����) _
                    Or (Not IsNull(rsSend!���ID) And rsSend!������� = "C") Then '��Ҫҽ����һ���ɼ��ļ�����Ŀ
                
                    '��ǰҽ������ͣʱ���:"��ͣʱ��,��ʼʱ��;...."
                    If rsSend!ҽ��״̬ <> 1 Then
                        strPause = GetAdvicePause(rsSend!ID)
                    End If
                    
                    '��ǰҽ���ķ��ͼ���ʱ���
                    datBegin = rsSend!��ʼִ��ʱ��
                    If Not IsNull(rsSend!�ϴ�ִ��ʱ��) Then
                        If IsNull(rsSend!ִ��ʱ�䷽��) And (NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Or IsNull(rsSend!�����λ)) Then
                            datBegin = DateAdd("s", 1, rsSend!�ϴ�ִ��ʱ��) '"������"����Ŀ
                        Else
                            datBegin = Calc�����ڿ�ʼʱ��(rsSend!��ʼִ��ʱ��, rsSend!�ϴ�ִ��ʱ��, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                            
                            '����������ִ�е�ʱ�䲻�ټ���,����ͨ����ͣ��ʽ������
                            If mlngRefModld = 0 Then
                                strPause = strPause & ";" & Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "," & Format(rsSend!�ϴ�ִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                                If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
                            End If
                        End If
                    End If
                    datEnd = CDate(strEnd)
                    If Not IsNull(rsSend!ִ����ֹʱ��) Then
                        If rsSend!ִ����ֹʱ�� < CDate(strEnd) Then
                            datEnd = rsSend!ִ����ֹʱ��
                        End If
                    End If
                    
                    '����ֽ�ʱ�估����
                    If IsNull(rsSend!ִ��ʱ�䷽��) And (NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Or IsNull(rsSend!�����λ)) Then
                        'ִ��Ƶ��Ϊ"������"����Ŀ,ÿ�췢��һ��(00:00)
                        lng���� = Calc�����Գ�������(datBegin, datEnd, _
                            Format(NVL(rsSend!�ϴ�ִ��ʱ��), "yyyy-MM-dd HH:mm:ss"), _
                            Format(NVL(rsSend!ִ����ֹʱ��), "yyyy-MM-dd HH:mm:ss"), _
                            strPause, str�״�ʱ��, strĩ��ʱ��, str�ֽ�ʱ��)
                        If lng���� = 0 Then '�������跢��
                            If rsSend!ҽ��״̬ = 1 Then '�����ͽ�У��
                                str�ֽ�ʱ�� = ""
                                str�״�ʱ�� = ""
                                strĩ��ʱ�� = ""
                            Else
                                lngDelҽ��ID = NVL(rsSend!ID, 0)
                                .RemoveItem lngRow
                                GoTo NextLoop
                            End If
                        End If
                        
                        '���Ϊ"������"����
                        .Cell(flexcpData, lngRow, COL_Ƶ��) = 2
                    Else
                        'ִ��Ƶ��Ϊ"��ѡƵ��"����Ŀ
                        str�ֽ�ʱ�� = Calc���ڷֽ�ʱ��(datBegin, datEnd, strPause, NVL(rsSend!ִ��ʱ�䷽��), rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ, rsSend!��ʼִ��ʱ��)
                        If str�ֽ�ʱ�� = "" Then '�޷��ֽ�ʱ��(�类��ͣ��)
                            If rsSend!ҽ��״̬ = 1 Then '�����ͽ�У��
                                lng���� = 0
                            Else
                                lngDelҽ��ID = NVL(rsSend!ID, 0)
                                .RemoveItem lngRow
                                GoTo NextLoop
                            End If
                        Else
                            lng���� = UBound(Split(str�ֽ�ʱ��, ",")) + 1
                        End If
                        
                    End If
                    If NVL(rsSend!�������, 0) = 1 Then
                        'ȡ������,�����Գ����޵���
                        dbl���� = IntEx(NVL(rsSend!��������, 1)) * lng����
                    Else
                        dbl���� = NVL(rsSend!��������, 1) * lng����
                    End If

                    .TextMatrix(lngRow, COL_����) = lng����
                    If Len(str�ֽ�ʱ��) > 4000 Then
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = Mid(str�ֽ�ʱ��, 1, InStr(Mid(str�ֽ�ʱ��, 4001), ",") + 3999)
                    Else
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
                    End If
                    If str�ֽ�ʱ�� <> "" Then   'Ҫ���͵ĳ�����Ϊ�գ�Ϊ�յ��ǽ�У�Ե�
                        .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "yyyy-MM-dd HH:mm")
                    End If
                    
                    .TextMatrix(lngRow, COL_����) = FormatEx(NVL(rsSend!��������), 5)
                    If Not IsNull(rsSend!��������) Then
                        .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    End If
                    .TextMatrix(lngRow, COL_����) = FormatEx(dbl����, 5)
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                ElseIf rsSend!������� = "E" And Not IsNull(rsSend!���ID) _
                    And .TextMatrix(lngRow - 1, COL_�������) = "K" And NVL(rsSend!���ID, 0) = Val(.TextMatrix(lngRow - 1, COL_ID)) Then
                    '��Ѫ;����ִ�д���
                    .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                    .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    
                    .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                    .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��)
                    .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                    .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                ElseIf Not IsNull(rsSend!���ID) Or bln�ɼ����� Then '����ҽ����걾�ɼ�����
                    '�����Ϻ�������ϲ�����Ϊ����,���Դ˶β���ִ��
                    .TextMatrix(lngRow, COL_����) = FormatEx(NVL(rsSend!��������), 5)
                    If Not IsNull(rsSend!��������) Then
                        .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    End If
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    
                    If lngRow > .FixedRows Then
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��)
                        .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                    End If
                End If
                If mlngRefModld = 1 Then
                    If (dkpExecTime.value >= rsSend!ִ����ֹʱ�� Or dkpExecTime.value <= rsSend!�ϴ�ִ��ʱ��) Then
                        '����ҽ��ɾ����Χ��ı���ҽ��
                        lngDelҽ��ID = NVL(rsSend!ID, 0)
                        .RemoveItem lngRow
                        GoTo NextLoop
                    End If
                End If
            Else
                '����---------------------------------------------
                If (IsNull(rsSend!���ID) And Not bln�ɼ�����) _
                    Or (Not IsNull(rsSend!���ID) And rsSend!������� = "C") Then '��Ҫҽ����һ���ɼ��ļ�����Ŀ
                    
                    If rsSend!������� = "K" Then
                        dbl���� = NVL(rsSend!�ܸ�����, 0)
                        
                        '��Ѫ;����ִ�д���
                        If IsNull(rsSend!ִ��ʱ�䷽��) And (NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Or IsNull(rsSend!�����λ)) Then
                            lng���� = 1 'ִ��Ƶ��Ϊ"һ����"����Ŀ
                        Else
                            lng���� = NVL(rsSend!Ƶ�ʴ���, 1) 'ִ��Ƶ��Ϊ"��ѡƵ��"����Ŀ
                        End If
                    Else
                        dbl���� = NVL(rsSend!�ܸ�����, 1)
                        lng���� = IntEx(dbl���� / NVL(rsSend!��������, 1))
                    End If
                    
                    If IsNull(rsSend!ִ��ʱ�䷽��) And (NVL(rsSend!Ƶ�ʴ���, 0) = 0 Or NVL(rsSend!Ƶ�ʼ��, 0) = 0 Or IsNull(rsSend!�����λ)) Then
                        'ִ��Ƶ��Ϊ"һ����"����Ŀ
                        str�ֽ�ʱ�� = "" '����Ҫ
                        .Cell(flexcpData, lngRow, COL_Ƶ��) = 1
                    Else
                        'ִ��Ƶ��Ϊ"��ѡƵ��"����Ŀ
                        If Not IsNull(rsSend!ִ��ʱ�䷽��) Or NVL(rsSend!�����λ) = "����" Then
                            str�ֽ�ʱ�� = Calc�����ֽ�ʱ��(lng����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", NVL(rsSend!ִ��ʱ�䷽��), rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                        Else
                            str�ֽ�ʱ�� = "" '����Ҳ��δ����ִ��ʱ��,�޷��ֽ�
                        End If
                    End If
                    .TextMatrix(lngRow, COL_����) = lng����
                    If Len(str�ֽ�ʱ��) > 4000 Then
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = Mid(str�ֽ�ʱ��, 1, InStr(Mid(str�ֽ�ʱ��, 4001), ",") + 3999)
                    Else
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
                    End If
                    If str�ֽ�ʱ�� <> "" Then
                        .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "yyyy-MM-dd HH:mm")
                    Else
                        '��¼���÷���ʱ��(���޷ֽ�ʱ��ʱ),��ҽ���Ŀ�ʼִ��ʱ��
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = CStr(Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss"))
                    End If
                    
                    .TextMatrix(lngRow, COL_����) = FormatEx(NVL(rsSend!��������), 5)
                    If Not IsNull(rsSend!��������) Then
                        .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    End If
                    .TextMatrix(lngRow, COL_����) = IIF(dbl���� = 0, "", FormatEx(dbl����, 5))
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                ElseIf rsSend!������� = "E" And Not IsNull(rsSend!���ID) _
                    And .TextMatrix(lngRow - 1, COL_�������) = "K" And NVL(rsSend!���ID, 0) = Val(.TextMatrix(lngRow - 1, COL_ID)) Then
                    '��Ѫ;����ִ�д���
                    .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                    .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    
                    .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                    .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��)
                    .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                    .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                ElseIf Not IsNull(rsSend!���ID) Or bln�ɼ����� Then
                    '����ҽ����걾�ɼ�����
                    .TextMatrix(lngRow, COL_����) = FormatEx(NVL(rsSend!��������), 5)
                    If Not IsNull(rsSend!��������) Then
                        .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    End If
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsSend!���㵥λ)
                    
                    If lngRow > .FixedRows Then
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��)
                        If .TextMatrix(lngRow, COL_�������) = "E" And .TextMatrix(lngRow, COL_��������) = "6" And .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��) <> .TextMatrix(lngRow, COL_��ʼʱ��) Then
                            .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
                        End If
                        .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                    End If
                End If
            End If
            
            If Not IsNull(rsSend!��������) Then
                lng������ = lng������ + 1 '�����Ƿ���ʾ������
            End If
            
            '�����ͣ�õģ�����ʾ���ܷ���
            If Format(NVL(rsSend!����ʱ��, "3000-1-1"), "YYYY-MM-DD") <> Format("3000-1-1", "YYYY-MM-DD") Then
                .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                If InStr(strͣ�� & ",", "," & rsSend!ҽ������ & ",") = 0 Then strͣ�� = strͣ�� & "," & rsSend!ҽ������
            End If
            
            '������Ŀ�Ľ��:���ڲ鿴�����ʱ���
            '---------------------------------------------------------------
            cur��� = 0
            If Not IsNull(rsSend!������ĿID) Then
                Call LoadAdvicePrice(lngRow, cur���)
            End If
            .TextMatrix(lngRow, COL_���) = Format(cur���, gstrDec)
            
            '�����ʱ��һЩ�����ۼ���ʾһ��ҽ���Ľ��
            '---------------------------------------------------------------
            If Not IsNull(rsSend!���ID) And rsSend!������� <> "C" Then
                '��������ҽ��
                For j = lngRow - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(j, COL_ID)) = rsSend!���ID Then
                        .TextMatrix(j, COL_���) = Format(Val(.TextMatrix(j, COL_���)) + Val(.TextMatrix(lngRow, COL_���)), gstrDec)
                        Exit For
                    End If
                Next
            ElseIf bln�ɼ����� Then
                '����걾�ɼ�����Ϊ��ʾ��
                For j = lngRow - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(j, COL_���ID)) = rsSend!ID Then
                        .TextMatrix(lngRow, COL_���) = Format(Val(.TextMatrix(lngRow, COL_���)) + Val(.TextMatrix(j, COL_���)), gstrDec)
                    Else
                        Exit For
                    End If
                Next
            End If
            
            '��������
            '---------------------------------------------------------------
            '���˼������ָ�
            If rsSend!����ID <> lng����ID Then
                lng������ = lng������ + 1
                If lng����ID <> 0 Then
                    For j = lngRow - 1 To .FixedRows Step -1
                        If Not .RowHidden(j) Then
                            .CellBorderRange j, .FixedCols, j, .Cols - 1, vbBlack, 0, 0, 0, 2, 0, 0
                            Exit For
                        End If
                    Next
                End If
            End If
            lng����ID = rsSend!����ID
            
NextLoop:           '---------------------------------------------------------------
            If blnOnePati Then Progress = i / rsSend.RecordCount * 100
            rsSend.MoveNext
        Next
        .Redraw = flexRDDirect
        
         '�����ͣ�õ���Ŀ������ʾ
        If strͣ�� <> "" Then
            Call MsgBox("������Ŀ��" & Mid(strͣ��, 2) & " �Ѿ�ͣ�ã����ܷ��͡�", vbInformation, Me.Caption)
        End If
    End With
    
    If blnOnePati Then Progress = 0
End Function

Private Function LoadAdvicePrice(ByVal lngRow As Long, cur�ϼ� As Currency, Optional ByVal rsDrug As ADODB.Recordset) As Boolean
'���ܣ���ȡָ��ҽ��(����ǰ��)�ļƼ۹�ϵ����ʱ��¼��,������ȱʡ���ͽ��(���ѱ����)
'������rsDrug=����������ҩƷ��Ϣ�ļ�¼��������ҩƷҽ��ʱ���롣��Ϊ���ܰ�����´ҽ���в�һ������ȷ��ҩƷID��
'���أ�cur�ϼ�=�������ҽ�����ͽ��(��ҩ���δ��,��Ҫ����۸�����)
    Dim rsTmp As New ADODB.Recordset
    Dim rsCur As New ADODB.Recordset
    Dim strSQL As String, strPrice As String
    Dim str�������� As String, arr�������� As Variant
    Dim blnDo As Boolean, i As Long, k As Long
    Dim dbl���� As Double, dbl���� As Double, dblӦ�� As Double
    Dim curӦ�� As Currency, curʵ�� As Currency
    Dim bln�������� As Boolean, lng��ĿID As Long
    Dim lng������ID As Long, blnHaveSub As Boolean
    Dim lngִ�п���ID As Long, cur��� As Currency
    Dim lng����ID As Long
    
    On Error GoTo errH
    
    cur��� = 0
    With vsAdvice
        If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(lngRow, COL_����ID)), Val(vsAdvice.TextMatrix(lngRow, COL_��ҳID)), "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
        
        If InStr(",5,6,7,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            '��ΪԺ��ִ��(�Ա�ҩ),ҩƷ������Ϊ����,�ҹ̶������Ƽ�
            If Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                mrsPrice.AddNew
                mrsPrice!ҽ��ID = Val(.TextMatrix(lngRow, COL_ID))
                If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                    mrsPrice!���ID = Val(.TextMatrix(lngRow, COL_���ID))
                End If
                mrsPrice!�������� = 0
                mrsPrice!�շѷ�ʽ = 0
                mrsPrice!�շ���� = .TextMatrix(lngRow, COL_�������)
                mrsPrice!�շ�ϸĿID = rsDrug!ҩƷID
                mrsPrice!ִ�п���ID = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
                mrsPrice!���� = 1
                mrsPrice!���� = 1
                mrsPrice!��� = NVL(rsDrug!�Ƿ���, 0)
                mrsPrice!�̶� = 1
                mrsPrice!���� = 0
                                
                '���͵���������
                If .TextMatrix(lngRow, COL_�������) = "7" Then
                    '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                    If Val(.TextMatrix(lngRow, COL_�ɷ����)) = 0 Then
                        dbl���� = Val(.TextMatrix(lngRow, COL_����)) * Val(.TextMatrix(lngRow, COL_����)) / NVL(rsDrug!����ϵ��, 1)
                    Else
                        dbl���� = Val(.TextMatrix(lngRow, COL_����)) _
                            * IntEx(Val(.TextMatrix(lngRow, COL_����)) / NVL(rsDrug!����ϵ��, 1) / NVL(rsDrug!סԺ��װ, 1)) * NVL(rsDrug!סԺ��װ, 1)
                    End If
                Else
                    dbl���� = Val(.TextMatrix(lngRow, COL_����)) * NVL(rsDrug!סԺ��װ, 1)
                End If
                dbl���� = Format(dbl����, "0.00000")
                                
                '��¼�ۼ۵���
                If NVL(rsDrug!�Ƿ���, 0) = 0 Then
                    mrsPrice!���� = Format(CalcPrice(rsDrug!ҩƷID, , , True, , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                Else '���ۼۼ���ҩƷʱ��,�Ա�ҩʱ�޶�Ӧҩ��
                    mrsPrice!���� = Format(CalcDrugPrice(rsDrug!ҩƷID, Val(.TextMatrix(lngRow, COL_ִ�п���ID)), dbl����, , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                End If
                mrsPrice.Update
                                
                '����ҽ�����ͽ��(���ѱ���۵�ʵ�ս��)
                If .TextMatrix(lngRow, COL_�ѱ�) <> "" Then
                    If NVL(rsDrug!�Ƿ���, 0) = 0 Then
                        cur��� = Format(CalcPrice(rsDrug!ҩƷID, .TextMatrix(lngRow, COL_�ѱ�), dbl����, , Val(.TextMatrix(lngRow, COL_ִ�п���ID)), , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDec)
                    Else
                        cur��� = Format(CalcDrugPrice(rsDrug!ҩƷID, Val(.TextMatrix(lngRow, COL_ִ�п���ID)), dbl����, .TextMatrix(lngRow, COL_�ѱ�), , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), "0.00000")
                    End If
                Else
                    If gbln�Ӱ�Ӽ� Then
                        '����Ӱ�Ӽ�
                        If NVL(rsDrug!�Ƿ���, 0) = 0 Then
                            dbl���� = Format(CalcPrice(rsDrug!ҩƷID, , , , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                        Else '���ۼۼ���ҩƷʱ��,�Ա�ҩʱ�޶�Ӧҩ��
                            dbl���� = Format(CalcDrugPrice(rsDrug!ҩƷID, Val(.TextMatrix(lngRow, COL_ִ�п���ID)), dbl����, , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                        End If
                        cur��� = Format(mrsPrice!���� * dbl���� * dbl����, gstrDec)
                    Else
                        cur��� = Format(mrsPrice!���� * dbl���� * mrsPrice!����, gstrDec)
                    End If
                End If
            End If
            
            cur�ϼ� = cur���
        ElseIf .TextMatrix(lngRow, COL_�������) = "4" Then
            '��ΪԺ��ִ��(�Ա�ҩ),ҩƷ������Ϊ����,�ҹ̶������Ƽ�
            If Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                mrsPrice.AddNew
                mrsPrice!ҽ��ID = Val(.TextMatrix(lngRow, COL_ID))
                If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                    mrsPrice!���ID = Val(.TextMatrix(lngRow, COL_���ID))
                End If
                mrsPrice!�������� = 0
                mrsPrice!�շѷ�ʽ = 0
                mrsPrice!�շ���� = .TextMatrix(lngRow, COL_�������)
                mrsPrice!�շ�ϸĿID = Val(.TextMatrix(lngRow, COL_�շ�ϸĿID))
                mrsPrice!ִ�п���ID = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
                mrsPrice!���� = 1
                mrsPrice!���� = Val(.TextMatrix(lngRow, COL_��������))
                mrsPrice!��� = Val(.TextMatrix(lngRow, COL_�Ƿ���))
                mrsPrice!�̶� = 1
                mrsPrice!���� = 0
                                
                '���͵���������
                dbl���� = Format(Val(.TextMatrix(lngRow, COL_����)), "0.00000")
                                
                '��¼�ۼ۵���
                If Val(.TextMatrix(lngRow, COL_�Ƿ���)) = 0 Then
                    '��������
                    mrsPrice!���� = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), , , True, , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                ElseIf Val(.TextMatrix(lngRow, COL_��������)) = 0 Then
                    '�Ǹ������õ�ʱ�����ģ��۸�����ѱ����ڲ���ҽ���Ƽ���
                    mrsPrice!���� = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), , , True, , Val(.TextMatrix(lngRow, COL_ID)), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                Else
                    '���������������ʱ��
                    mrsPrice!���� = Format(CalcDrugPrice(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), Val(.TextMatrix(lngRow, COL_ִ�п���ID)), dbl����, , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                End If
                mrsPrice.Update
                                
                '����ҽ�����ͽ��(���ѱ���۵�ʵ�ս��)
                If .TextMatrix(lngRow, COL_�ѱ�) <> "" Then
                    If Val(.TextMatrix(lngRow, COL_�Ƿ���)) = 0 Then
                        cur��� = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), .TextMatrix(lngRow, COL_�ѱ�), dbl����, , Val(.TextMatrix(lngRow, COL_ִ�п���ID)), , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDec)
                    ElseIf Val(.TextMatrix(lngRow, COL_��������)) = 0 Then
                        cur��� = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), .TextMatrix(lngRow, COL_�ѱ�), dbl����, , Val(.TextMatrix(lngRow, COL_ִ�п���ID)), Val(.TextMatrix(lngRow, COL_ID)), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDec)
                    Else
                        cur��� = Format(CalcDrugPrice(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), Val(.TextMatrix(lngRow, COL_ִ�п���ID)), dbl����, .TextMatrix(lngRow, COL_�ѱ�), , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), "0.00000")
                    End If
                Else
                    If gbln�Ӱ�Ӽ� Then
                        '����Ӱ�Ӽ�
                        If Val(.TextMatrix(lngRow, COL_�Ƿ���)) = 0 Then
                            dbl���� = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), , , , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                        ElseIf Val(.TextMatrix(lngRow, COL_��������)) = 0 Then
                            dbl���� = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), , , , , Val(.TextMatrix(lngRow, COL_ID)), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                        Else '���ۼۼ���ҩƷʱ��,�Ա�ҩʱ�޶�Ӧҩ��
                            dbl���� = Format(CalcDrugPrice(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), Val(.TextMatrix(lngRow, COL_ִ�п���ID)), dbl����, , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                        End If
                        cur��� = Format(mrsPrice!���� * dbl���� * dbl����, gstrDec)
                    Else
                        cur��� = Format(mrsPrice!���� * dbl���� * mrsPrice!����, gstrDec)
                    End If
                End If
            End If
            
            cur�ϼ� = cur���
        Else
            'ȡ�����շ� ��ϵ�еĶ���(����ʱ�Ŷ��Ƽ�):�����Ƽ�,��Ϊ������Ժ��ִ��
            If Val(.TextMatrix(lngRow, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������ID))) = 0 Then
                dbl���� = Format(Val(.TextMatrix(lngRow, COL_����)), "0.00000")
                bln�������� = (.TextMatrix(lngRow, COL_�������) = "F" And Val(.TextMatrix(lngRow, COL_���ID)) <> 0)
                
                '���ֶ�Ӧ�ļƼ����
                If .TextMatrix(lngRow, COL_�걾��λ) <> "" And .TextMatrix(lngRow, COL_��鷽��) <> "" Then
                    strPrice = " And ��鲿λ=[4] And ��鷽��=[5] And Nvl(��������,0)=0"
                ElseIf Val(.TextMatrix(lngRow, COL_ִ�б��)) = 0 Then
                    strPrice = " And ��鲿λ Is Null And ��鷽�� is Null And Nvl(��������,0)=0"
                Else 'Ŀǰ�������Ի����м��յ����
                    strPrice = " And ��鲿λ Is Null And ��鷽�� is Null And Nvl(��������,0) IN(0,1)"
                End If
                
                strPrice = "Select �շ���ĿID,���ж��� From (" & _
                    " Select c.�շ���ĿID, c.���ж���, c.���ÿ���id" & _
                    "   ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
                    " From �����շѹ�ϵ C Where C.������ĿID=[2]" & strPrice & _
                    "       And (C.���ÿ���ID is Null And C.������Դ = 0 or C.���ÿ���ID = Decode([3],0,[6],[3]) And C.������Դ = 2)" & _
                    " ) Where Nvl(���ÿ���id, 0) = Top"
                
                '�ȶ�ȡ���еļƼ�
                strSQL = _
                    " Select C.���,A.�շ�ϸĿID as �շ���ĿID,A.���� as �շ�����,Nvl(E.���ж���,0) as ���ж���," & _
                    " B.������ĿID,C.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,Decode(C.�Ƿ���,1,A.����,B.�ּ�)" & IIF(bln��������, "*Nvl(B.�����շ���,100)/100", "") & " as ����," & _
                    " C.�Ƿ���,Nvl(A.����,0) as ����,D.��������,Nvl(A.ִ�п���ID,[3]) as ִ�п���ID,C.���ηѱ�," & _
                    " Nvl(A.��������,0) as ��������,Nvl(A.�շѷ�ʽ,0) as �շѷ�ʽ" & _
                    " From ����ҽ���Ƽ� A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,�������� D,(" & strPrice & ") E" & _
                    " Where A.ҽ��ID=[1] And A.�շ�ϸĿID=0+E.�շ���ĿID(+)" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "B", "7", "8", "9") & _
                    " And A.�շ�ϸĿID=B.�շ�ϸĿID And A.�շ�ϸĿID=C.ID And A.�շ�ϸĿID=D.����ID(+)" & _
                    " And C.������� IN(2,3) And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " Order by ��������,����,A.�շ�ϸĿID"
                
                
                '����ȡĬ�ϵļƼ�(�����Ƿ���)
                '��ҪУ�Ե�ģʽ�£����﷢�͵Ķ��Ǿ���У�Եģ���ʵ����ȷ���ļƼ�����Ϊ׼�������ٶ�ȱʡ�Ƽۣ���Ϊ�п���У�Ի�Ƽ۵���ʱ��ɾ��ĳЩ��Ŀ
                '��У�Լ����͵�ģʽ��ֻ�����¿�״̬�²Ŷ�ȡ����Ϊ���ͺ�ͬ�ϡ�
                If mblnAutoVerify And Val(.TextMatrix(lngRow, COL_ҽ��״̬)) = 1 Then
                    lng����ID = 0 '�����Թܷ���,ֻ��ȡ�Թܶ�Ӧ�����ķ���
                    If .TextMatrix(lngRow, COL_�Թܱ���) <> "" Then
                        lng����ID = GetTubeMaterial(.TextMatrix(lngRow, COL_�Թܱ���))
                    End If
                
                    '���ֶ�Ӧ�ļƼ����
                    If .TextMatrix(lngRow, COL_�걾��λ) <> "" And .TextMatrix(lngRow, COL_��鷽��) <> "" Then
                        strPrice = " And c.��鲿λ=[3] And c.��鷽��=[4] And Nvl(c.��������,0)=0"
                    ElseIf Val(.TextMatrix(lngRow, COL_ִ�б��)) = 0 Then
                        strPrice = " And c.��鲿λ Is Null And c.��鷽�� is Null And Nvl(c.��������,0)=0"
                    Else 'Ŀǰ�������Ի����м��յ����
                        strPrice = " And c.��鲿λ Is Null And c.��鷽�� is Null And Nvl(c.��������,0) IN(0,1)"
                    End If
                    
                    strPrice = "Select * From (" & _
                        "Select C.������ĿID,C.�շ���ĿID,C.��鲿λ,C.��鷽��,C.��������,C.�շ�����,C.���ж���,C.������Ŀ,C.�շѷ�ʽ,c.���ÿ���id" & _
                        " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
                        " From �����շѹ�ϵ C Where C.������ĿID=[1]" & strPrice & _
                        "      And (C.���ÿ���ID is Null And C.������Դ = 0 or C.���ÿ���ID = Decode([2],0,[6],[2]) And C.������Դ = 2)" & _
                        " ) Where Nvl(���ÿ���id, 0) = Top"
                    
                    strSQL = _
                        " Select C.���,A.�շ���ĿID,A.�շ�����,A.���ж���,B.������ĿID," & _
                        " C.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,Decode(C.�Ƿ���,1,B.ȱʡ�۸�,B.�ּ�)" & IIF(bln��������, "*Nvl(B.�����շ���,100)/100", "") & " as ����," & _
                        " C.�Ƿ���,Nvl(A.������Ŀ,0) as ����,D.��������,[2] as ִ�п���ID,C.���ηѱ�," & _
                        " Nvl(A.��������,0) as ��������,Nvl(A.�շѷ�ʽ,0) as �շѷ�ʽ" & _
                        " From (" & strPrice & ") A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,�������� D" & _
                        " Where A.�շ���ĿID=B.�շ�ϸĿID And A.�շ���ĿID=C.ID And A.�շ���ĿID=D.����ID(+)" & _
                        GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "B", "7", "8", "9") & _
                        " And C.������� IN(2,3) And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                        " And (Nvl(A.�շѷ�ʽ,0)=1 And C.���='4' And A.�շ���ĿID=[5] Or Not(Nvl(A.�շѷ�ʽ,0)=1 And C.���='4' And [5]<>0))" & _
                        " Order by ��������,����,A.�շ���ĿID"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_������ĿID)), _
                        Val(.TextMatrix(lngRow, COL_ִ�п���ID)), .TextMatrix(lngRow, COL_�걾��λ), .TextMatrix(lngRow, COL_��鷽��), lng����ID, mlng����ID, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                Else
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_������ĿID)), _
                        Val(.TextMatrix(lngRow, COL_ִ�п���ID)), .TextMatrix(lngRow, COL_�걾��λ), .TextMatrix(lngRow, COL_��鷽��), mlng����ID, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                End If
                
                'ȷ���Ƽ�֮���Ƿ���������Լ���������ID
                arr�������� = Array()
                If Not rsTmp.EOF Then
                    Do While Not rsTmp.EOF
                        If InStr(str�������� & ",", "," & rsTmp!�������� & ",") = 0 Then
                            str�������� = str�������� & "," & rsTmp!��������
                        End If
                        rsTmp.MoveNext
                    Loop
                    arr�������� = Split(Mid(str��������, 2), ",")
                End If
                                
                For k = 0 To UBound(arr��������)
                    rsTmp.Filter = "��������=" & arr��������(k)
                    
                    lng��ĿID = 0: cur��� = 0
                    lng������ID = 0: blnHaveSub = False
                    If Not rsTmp.EOF And gbln��������ۿ� Then
                        Do While Not rsTmp.EOF
                            If NVL(rsTmp!����, 0) = 0 Then
                                'SQL����������ǰ��,ֻȡ����Ŀ�ĵ�һ������
                                If lng������ID = 0 Then lng������ID = rsTmp!������ĿID
                            ElseIf NVL(rsTmp!����, 0) = 1 Then
                                blnHaveSub = True: Exit Do
                            End If
                            rsTmp.MoveNext
                        Loop
                        rsTmp.MoveFirst
                    End If
                    
                    Do While True
                        blnDo = False
                        If rsTmp.EOF Then
                            If lng��ĿID <> 0 Then blnDo = True
                        Else
                            If rsTmp!�շ���ĿID <> lng��ĿID And lng��ĿID <> 0 Then blnDo = True
                        End If
                        If blnDo Then
                            If Not IsNull(mrsPrice!����) Then
                                mrsPrice!���� = Format(mrsPrice!����, gstrDecPrice)
                            End If
                            mrsPrice.Update
                            
                            'ҽ�����ͽ��
                            cur��� = cur��� + Format(curʵ��, gstrDec)
                        End If
                        If rsTmp.EOF Then Exit Do
                        
                        '------------------------------------
                        If rsTmp!�շ���ĿID <> lng��ĿID Then
                            curʵ�� = 0
                            mrsPrice.AddNew
                            mrsPrice!ҽ��ID = Val(.TextMatrix(lngRow, COL_ID))
                            If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                                mrsPrice!���ID = Val(.TextMatrix(lngRow, COL_���ID))
                            End If
                            mrsPrice!�������� = NVL(rsTmp!��������, 0)
                            mrsPrice!�շѷ�ʽ = NVL(rsTmp!�շѷ�ʽ, 0)
                            mrsPrice!�շ���� = rsTmp!���
                            mrsPrice!�շ�ϸĿID = rsTmp!�շ���ĿID
                            mrsPrice!���� = NVL(rsTmp!�շ�����, 0)
                            mrsPrice!���� = NVL(rsTmp!��������, 0)
                            mrsPrice!��� = NVL(rsTmp!�Ƿ���, 0)
                            mrsPrice!�̶� = NVL(rsTmp!���ж���, 0)
                            mrsPrice!���� = NVL(rsTmp!����, 0)
                            
                            If .TextMatrix(lngRow, COL_�������) = "E" And .TextMatrix(lngRow, COL_��������) = "1" And .TextMatrix(lngRow, COL_ִ�з���) = "5" And InStr(",5,6,", rsTmp!���) > 0 Then
                                'ԭҺƤ�����⡣�󶨵�ҩƷ�������û��ָ��������ԭ���߼�
                                If Val(.TextMatrix(lngRow, COL_��ҩ����)) <> 0 Then
                                    lngִ�п���ID = Val(.TextMatrix(lngRow, COL_��ҩ����))
                                Else
                                    lngִ�п���ID = NVL(rsTmp!ִ�п���ID, 0)
                                End If
                                lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsTmp!���, rsTmp!�շ���ĿID, 4, Val(.TextMatrix(lngRow, COL_���˿���ID)), 0, 2, lngִ�п���ID, , , 2)
                            Else
                                'ִ�п���:��ҩ��ҩƷ���������ĵ�ר��ȡ
                                lngִ�п���ID = NVL(rsTmp!ִ�п���ID, 0)
                                If rsTmp!��� = "4" And NVL(rsTmp!��������, 0) = 1 Or InStr(",5,6,7,", rsTmp!���) > 0 Then
                                    lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsTmp!���, rsTmp!�շ���ĿID, 4, Val(.TextMatrix(lngRow, COL_���˿���ID)), 0, 2, lngִ�п���ID, , , 2)
                                End If
                            End If
                            
                            If lngִ�п���ID <> 0 Then
                                mrsPrice!ִ�п���ID = lngִ�п���ID
                            Else
                                mrsPrice!ִ�п���ID = Null
                            End If
                        End If
                        lng��ĿID = rsTmp!�շ���ĿID
                        
                        '���㵥�ۺ�ʵ��
                        If NVL(rsTmp!�Ƿ���, 0) = 1 And InStr(",5,6,7,", rsTmp!���) > 0 Then
                            '��ҩ��ҩƷ�Ƽ۰�ʱ�ۼ���(��һ������),���������Ҫ��ҽ������
                            mrsPrice!���� = CalcDrugPrice(rsTmp!�շ���ĿID, NVL(mrsPrice!ִ�п���ID, 0), dbl���� * NVL(rsTmp!�շ�����, 0), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                            
                            dblӦ�� = Format(mrsPrice!���� * dbl����, "0.00000") * Format(mrsPrice!����, gstrDecPrice)
                            
                            '����Ӱ�Ӽ�
                            If gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                                dblӦ�� = dblӦ�� * (1 + NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                            End If
    
                            curӦ�� = Format(dblӦ��, gstrDec)
                            
                            If .TextMatrix(lngRow, COL_�ѱ�) <> "" And Not (gbln��������ۿ� And blnHaveSub) And NVL(rsTmp!���ηѱ�, 0) = 0 Then
                                curʵ�� = curʵ�� + Format(ActualMoney(.TextMatrix(lngRow, COL_�ѱ�), rsTmp!������ĿID, curӦ��, rsTmp!�շ���ĿID, lngִ�п���ID, _
                                    mrsPrice!���� * dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1, NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                            Else
                                curʵ�� = curʵ�� + curӦ��
                            End If
                        ElseIf NVL(rsTmp!�Ƿ���, 0) = 1 And rsTmp!��� = "4" And NVL(rsTmp!��������, 0) = 1 Then
                            '�������õ�ʱ�����ĺ�ҩƷһ������
                            mrsPrice!���� = CalcDrugPrice(rsTmp!�շ���ĿID, NVL(mrsPrice!ִ�п���ID, 0), dbl���� * NVL(rsTmp!�շ�����, 0), , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                            
                            dblӦ�� = Format(mrsPrice!���� * dbl����, "0.00000") * Format(mrsPrice!����, gstrDecPrice)
                            
                            '����Ӱ�Ӽ�
                            If gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                                dblӦ�� = dblӦ�� * (1 + NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                            End If
    
                            curӦ�� = Format(dblӦ��, gstrDec)
                            
                            If .TextMatrix(lngRow, COL_�ѱ�) <> "" And Not (gbln��������ۿ� And blnHaveSub) And NVL(rsTmp!���ηѱ�, 0) = 0 Then
                                curʵ�� = curʵ�� + Format(ActualMoney(.TextMatrix(lngRow, COL_�ѱ�), rsTmp!������ĿID, curӦ��, rsTmp!�շ���ĿID, lngִ�п���ID, _
                                    mrsPrice!���� * dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1, NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                            Else
                                curʵ�� = curʵ�� + curӦ��
                            End If
                        Else '�̶��۸����ͨ���(ֻ��һ��������Ŀ)
                            mrsPrice!���� = NVL(mrsPrice!����, 0) + NVL(rsTmp!����, 0)
                            
                            dblӦ�� = Format(mrsPrice!���� * dbl����, "0.00000") * Format(NVL(rsTmp!����, 0), gstrDecPrice)
                            
                            '����Ӱ�Ӽ�
                            If gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                                dblӦ�� = dblӦ�� * (1 + NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                            End If
                            
                            curӦ�� = Format(dblӦ��, gstrDec)
                            
                            If .TextMatrix(lngRow, COL_�ѱ�) <> "" And Not (gbln��������ۿ� And blnHaveSub) And NVL(rsTmp!���ηѱ�, 0) = 0 Then
                                curʵ�� = curʵ�� + Format(ActualMoney(.TextMatrix(lngRow, COL_�ѱ�), rsTmp!������ĿID, curӦ��, rsTmp!�շ���ĿID, lngִ�п���ID, _
                                    mrsPrice!���� * dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsTmp!�Ӱ�Ӽ�, 0) = 1, NVL(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                            Else
                                curʵ�� = curʵ�� + curӦ��
                            End If
                        End If
                        
                        rsTmp.MoveNext
                    Loop
                    
                    '������Ŀ���ܼ����ۿ�
                    If gbln��������ۿ� And blnHaveSub And lng������ID <> 0 Then
                        cur��� = Format(ActualMoney(.TextMatrix(lngRow, COL_�ѱ�), lng������ID, cur���), gstrDec)
                    End If
                    
                    cur�ϼ� = cur�ϼ� + cur���
                Next
            End If
        End If
    End With
    LoadAdvicePrice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitBillSet()
'���ܣ���ʼ��ҽ�����ʵ������ɼ�¼��
    Set mrsBill = New ADODB.Recordset
    
    mrsBill.Fields.Append "Key", adVarChar, 100
    mrsBill.Fields.Append "NO", adVarChar, 30
    mrsBill.Fields.Append "�������", adBigInt
    mrsBill.Fields.Append "�������", adBigInt
    mrsBill.CursorLocation = adUseClient
    mrsBill.LockType = adLockOptimistic
    mrsBill.CursorType = adOpenStatic
    mrsBill.Open
End Sub

Private Sub InitSeekSet(rsSeek As ADODB.Recordset)
'���ܣ���ʼ�����ڻ��ܼ����ۿ۵���ʱ��¼��
    Set rsSeek = New ADODB.Recordset
    rsSeek.Fields.Append "��������", adInteger
    rsSeek.Fields.Append "�����ǩ", adVariant
    rsSeek.Fields.Append "������ID", adBigInt
    rsSeek.Fields.Append "�ϼ�", adCurrency, , adFldIsNullable
    rsSeek.CursorLocation = adUseClient
    rsSeek.LockType = adLockOptimistic
    rsSeek.CursorType = adOpenStatic
    rsSeek.Open
End Sub

Private Sub InitPriceRecordset()
'���ܣ���ʼ��ҽ���Ƽۼ�¼��
    Set mrsPrice = New ADODB.Recordset
    
    mrsPrice.Fields.Append "ҽ��ID", adBigInt
    mrsPrice.Fields.Append "���ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "��������", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "�շѷ�ʽ", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "�շ����", adVarChar, 1
    mrsPrice.Fields.Append "�շ�ϸĿID", adBigInt
    mrsPrice.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "����", adDouble
    mrsPrice.Fields.Append "����", adDouble, , adFldIsNullable '��ۼ۸�
    mrsPrice.Fields.Append "����", adInteger '�����Ƿ��������
    mrsPrice.Fields.Append "���", adInteger
    mrsPrice.Fields.Append "����", adInteger
    mrsPrice.Fields.Append "�̶�", adInteger
    
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
End Sub

Private Sub InitRecordSet(rsSQL As ADODB.Recordset, rsTotal As ADODB.Recordset, rsUpload As ADODB.Recordset, _
    rsNumber As ADODB.Recordset, rsMoneyNow As ADODB.Recordset, rsItems As ADODB.Recordset, rsSendDays As ADODB.Recordset)
'��ʼ����¼��
    'SQL��¼��
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "����", adInteger '1-���ü�¼,2-ҽ����¼,3-���ͼ�¼,4-���ϼ�¼
    rsSQL.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsSQL.Fields.Append "��ĿID", adBigInt '�շ�ϸĿID
    rsSQL.Fields.Append "���", adBigInt '��������
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
    rsSQL.Fields.Append "NO", adVarChar, 30, adFldIsNullable '����NO�滻����ʱ����
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
    
    '�Ƽ������ۼƼ�¼��
    Set rsTotal = New ADODB.Recordset
    rsTotal.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsTotal.Fields.Append "��ĿID", adBigInt
    rsTotal.Fields.Append "�ⷿID", adBigInt
    rsTotal.Fields.Append "����", adDouble
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open
    
    'ҽ���ϴ����ʵ�
    Set rsUpload = New ADODB.Recordset
    rsUpload.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsUpload.Fields.Append "NO", adVarChar, 30
    rsUpload.CursorLocation = adUseClient
    rsUpload.LockType = adLockOptimistic
    rsUpload.CursorType = adOpenStatic
    rsUpload.Open
    
    '��¼�Թܱ���
    Set rsNumber = New ADODB.Recordset
    rsNumber.Fields.Append "����", adVarChar, 18
    rsNumber.Fields.Append "���ID", adBigInt
    rsNumber.Fields.Append "��������", adVarChar, 18
    rsNumber.Fields.Append "ִ�п���ID", adVarChar, 18
    rsNumber.Fields.Append "������ĿID", adVarChar, 18
    rsNumber.Fields.Append "Ӥ��", adBigInt
    rsNumber.Fields.Append "������־", adBigInt
    rsNumber.Fields.Append "�걾", adVarChar, 18
    rsNumber.Fields.Append "�ɼ�����ID", adBigInt
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open
    
    '��ǰ���˱���Ҫ���͵ķ���
    Set rsMoneyNow = New ADODB.Recordset
    rsMoneyNow.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsMoneyNow.Fields.Append "������ĿID", adBigInt
    rsMoneyNow.Fields.Append "�շ���ĿID", adBigInt
    rsMoneyNow.Fields.Append "�Թܱ���", adVarChar, 18, adFldIsNullable
    rsMoneyNow.Fields.Append "��������", adVarChar, 50, adFldIsNullable
    rsMoneyNow.Fields.Append "�շѷ�ʽ", adInteger
    rsMoneyNow.Fields.Append "�շ�ʱ��", adVarChar, 10
    rsMoneyNow.Fields.Append "ִ�в���ID", adBigInt
    rsMoneyNow.Fields.Append "��ҽ��ID", adBigInt '���ID��Ϊ�յ�ҽ���е�ҽ��ID
    rsMoneyNow.Fields.Append "��鲿λ", adVarChar, 100
    rsMoneyNow.Fields.Append "��鷽��", adVarChar, 100
    rsMoneyNow.Fields.Append "����", adDouble '�շ�����
    rsMoneyNow.CursorLocation = adUseClient
    rsMoneyNow.LockType = adLockOptimistic
    rsMoneyNow.CursorType = adOpenStatic
    rsMoneyNow.Open
    
    '��ǰ���˱��η��͵ķ�����Ŀ����
    Set rsItems = New ADODB.Recordset
    rsItems.Fields.Append "����ID", adBigInt
    rsItems.Fields.Append "��ҳID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "ҽ��ID", adBigInt
    rsItems.Fields.Append "�շ����", adVarChar, 1
    rsItems.Fields.Append "�շ�ϸĿID", adBigInt
    rsItems.Fields.Append "����", adDouble
    rsItems.Fields.Append "����", adDouble
    rsItems.Fields.Append "ʵ�ս��", adDouble
    rsItems.Fields.Append "������", adVarChar, 100, adFldIsNullable
    rsItems.Fields.Append "��������", adVarChar, 100, adFldIsNullable
    rsItems.CursorLocation = adUseClient
    rsItems.LockType = adLockOptimistic
    rsItems.CursorType = adOpenStatic
    rsItems.Open
    
    '������ҩ�����ַ��ͼ�¼
    Set rsSendDays = New ADODB.Recordset
    rsSendDays.Fields.Append "ҽ��ID", adBigInt
    rsSendDays.Fields.Append "���ͺ�", adBigInt
    rsSendDays.Fields.Append "����ʱ��", adVarChar, 100
    rsSendDays.Fields.Append "�״�ʱ��", adVarChar, 100
    rsSendDays.Fields.Append "ĩ��ʱ��", adVarChar, 100
    rsSendDays.Fields.Append "�ֽ�ʱ��", adVarChar, 4000
    rsSendDays.Fields.Append "NO", adVarChar, 30
    rsSendDays.Fields.Append "��������", adDouble
    rsSendDays.CursorLocation = adUseClient
    rsSendDays.LockType = adLockOptimistic
    rsSendDays.CursorType = adOpenStatic
    rsSendDays.Open
End Sub

Private Sub GetCurBillSet(ByVal strKey As String, strNO As String, lng������� As Long, lng������� As Long)
'���ܣ���ȡ��ǰ���ʵ��ݵ�NO�����
'������lng�������=���ü�¼�е����,Ϊ-1ʱ��ʾ��ȡ�������
'      lng�������=���ͼ�¼�е����,Ϊ-1ʱ��ʾ��ȡ�������
'˵����strKey=���ݼ��ʵ������ɹ��򶨵�Ψһ�ؼ���
'1.������ҩ��"����(����ID,��ҳID)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
'2.һ���䷽�е����в�ҩ����һ���������ݺ�
'3.����ҽ�����ҩ�ֺŹ�����ͬ��
'4.������ҩҽ��ÿ��ҽ��һ���������ݺ�(������ҩ;�����䷽�巨���÷�)
'5.��鲿λ�͸�����������Ҫҽ��������ͬ���ݺţ�����������䵥���ĵ��ݺš�
'6.һ���ɼ��ļ�����Ϸ�����ͬ�ĵ��ݺţ��걾�ɼ��������䵥���ĵ��ݺ�
    mrsBill.Filter = "Key='" & strKey & "'"
    If mrsBill.EOF Then
        mrsBill.AddNew
        mrsBill!Key = strKey
        
        'ȡ���ݺ�
        'mrsBill!NO = zlDatabase.GetNextNo(14)
        mlngNOSequence = mlngNOSequence + 1
        mrsBill!NO = "TemporaryNO=" & Format(mlngNOSequence, "00000")
        
        mrsBill!������� = IIF(lng������� = -1, 0, 1)
        mrsBill!������� = IIF(lng������� = -1, 0, 1)
        mrsBill.Update
    Else
        If lng������� <> -1 Then
            mrsBill!������� = mrsBill!������� + 1
        End If
        If lng������� <> -1 Then
            mrsBill!������� = mrsBill!������� + 1
        End If
        mrsBill.Update
    End If
    strNO = mrsBill!NO
    If lng������� <> -1 Then lng������� = mrsBill!�������
    If lng������� <> -1 Then lng������� = mrsBill!�������
End Sub

Private Sub ReplaceTrueNO(rsSQL As ADODB.Recordset, rsUpload As ADODB.Recordset)
'���ܣ�����ʱ������NO�滻�����ձ������ʵNO
    Dim strNO As String, strCur As String, strPre As String
    
    rsSQL.Filter = 0
    rsSQL.Sort = "NO"
    Do While Not rsSQL.EOF
        If Not IsNull(rsSQL!NO) Then
            strCur = Split(rsSQL!NO, "=")(1)
            If strCur <> strPre Then
                strPre = strCur
                strNO = zlDatabase.GetNextNo(14)
                            
                'rsUpload��һ��NOֻ��һ����¼
                rsUpload.Filter = "NO='" & rsSQL!NO & "'"
                If Not rsUpload.EOF Then
                    rsUpload!NO = strNO
                    rsUpload.Update
                End If
            End If
            
            rsSQL!Sql = Replace(rsSQL!Sql, rsSQL!NO, strNO)
            'rsSQL!NO = strNO '��������£����⵼��Sort��˳������
            rsSQL.Update
        End If
        rsSQL.MoveNext
    Loop
End Sub

Private Function CompletePatiSend(rsPati As ADODB.Recordset, rsSQL As ADODB.Recordset, _
    rsUpload As ADODB.Recordset, rsItems As ADODB.Recordset, ByVal cur�ϼ� As Currency, ByVal cur���ʺϼ� As Currency, ByVal str��� As String, _
    ByVal bln���� As Boolean, blnTran As Boolean, ByVal lng���ͺ� As Long) As Boolean
'���ܣ��ύһ�����˵�ҽ����������,����֮ǰ������ʱ���
'������rsPati=����������Ϣ�ļ�¼��,���ڼ��ʱ���
'      rsSQL=��������Ҫִ�е�SQL
'      rsUpload=����ҽ���ϴ��ļ��ʵ��ݺ�
'      rsItems=����ҽ���ܿؼ�����Ŀ���ܼ�¼��
'      cur�ϼ�=���˱���Ҫ����ҽ���ļ��ʽ��ϼ�,���ڼ��ʱ���
'      cur���ʺϼ�=���˱���Ҫ����ҽ���ļ��ʽ��ϼƣ���������ִ�к��Զ���˵Ļ��۷��ã������������۷���
'      bln����=�Ƿ�ȫ�����ö��ǻ���ģʽ�����ڱ��������⴦��
'      str���=���˱��η��ͼ��ʷ��õ��շ����,���ڼ��ʱ���
'      lng���ͺ�=���η��͵����ؼ���
'˵�����������,���ڵ��ú����д���,blnTran�����Ƿ�����������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intR As Integer, lng��ID As Long, strҽ��IDs As String, lngS As Long
    Dim i As Long, j As Long
    Dim strMsg As String, strAllmsg As String, strDiag As String, strAdviceInfo As String
    Dim strErr As String
    Dim str��ǰ���� As String
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    Dim blnClearPatiCache As Boolean
    Dim blnPlugIn As Boolean
    Dim rsAdviceRis As ADODB.Recordset
    Dim strAdvices��Ѫ As String
    Dim var��Ѫ As Variant
    Dim bln�������Ѻ��� As Boolean
    Dim intBnt As Integer
    
    lng����ID = Val(rsPati!����ID & "")
    lng��ҳID = Val(rsPati!��ҳID & "")
    blnClearPatiCache = True
    
    '������ҽӿڷ���ǰ���ҽ������
    If CreatePlugInOK(pסԺҽ������, 1) Then
        blnPlugIn = True
        On Error Resume Next
        blnPlugIn = gobjPlugIn.AdviceCheckSendFee(glngSys, pסԺҽ������, lng����ID, lng��ҳID, cur�ϼ�, 1)
        If Not blnPlugIn And err.Number <> 0 Then blnPlugIn = True
        Call zlPlugInErrH(err, "AdviceCheckSendFee")
        err.Clear: On Error GoTo 0
        If Not blnPlugIn Then
            Exit Function
        End If
    End If
     
    '���˷��ñ���
    If cur�ϼ� > 0 Then
        If InitObjPublicExpense Then
            For i = 1 To Len(str���)
                intBnt = mintBnt
                Call gobjPublicExpense.zlBillingWarn.zlBillingWarnCheck(Me, 1, IIF(bln����, 1, 0), lng����ID, lng��ҳID, mlng����ID, Mid(str���, i, 1), IIF(gbln�����������۷���, cur�ϼ�, cur���ʺϼ�), InStr(";" & GetInsidePrivs(pסԺҽ������) & ";", ";Ƿ��ǿ�Ƽ���;") > 0, False, blnClearPatiCache, intR, , , , True, True, bln�������Ѻ���, intBnt)
                blnClearPatiCache = False
                If bln�������Ѻ��� And Not mbln�������Ѻ��� Then
                    mbln�������Ѻ��� = True
                    mintBnt = IIF(InStr(",2,3,", intR) > 0, vbCancel, vbIgnore)
                End If
                If InStr(",2,3,", intR) > 0 Then Exit For
            Next
        End If
    End If
    
    If InStr(",2,3,", intR) = 0 Then
        'ҽ���ܿ�ʵʱ���
        If Not IsNull(rsPati!����) Then
            If gclsInsure.GetCapability(supportʵʱ���, rsPati!����ID, rsPati!����) Then
                rsItems.Filter = 0
                If Not rsItems.EOF Then
                    If Not gclsInsure.CheckItem(rsPati!����, 1, 2, rsItems) Then
                        CompletePatiSend = True: Exit Function '���Լ�����һ������
                    End If
                End If
            End If
        End If
        Call InitObjLis(pסԺ��ʿվ)
        'LIS����ӿ�
        If Not gobjLIS Is Nothing Then
            strAdviceInfo = Get����ҽ����Ϣ(lng����ID, lng��ҳID)
            If strAdviceInfo <> "" Then
                Set rsTmp = Get������ϼ�¼(lng����ID, lng��ҳID, "2")
                If rsTmp.RecordCount > 0 Then strDiag = rsTmp!�������
            End If
        End If
            
        If gblnѪ��ϵͳ Then
            If InitObjBlood(True) Then
                strAdvices��Ѫ = Get��Ѫҽ����Ϣ(lng����ID, lng��ҳID)
                If strAdvices��Ѫ <> "" Then
                    var��Ѫ = Split(strAdvices��Ѫ, ",")
                End If
            End If
        End If
        
        Call ReplaceTrueNO(rsSQL, rsUpload)
        
        'ִ��˳��:1-����,2-ҽ��ִ�п���,3-����,4-�Զ�����
        '1.����д����,��Ϊ����ʱ���ܴ������
        '2.�Է��ü�¼���շ�ϸĿID�������
        rsSQL.Filter = 0 '�ϲ㺯������ʹ�ù�,��ʹû�ù�ҲMoveFirst
        rsSQL.Sort = "����,��ĿID,���"
        rsUpload.Filter = 0 '�ϲ㺯������ʹ�ù�,��ʹû�ù�ҲMoveFirst
        
        If mblnSendMsg Then
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        If .TextMatrix(i, COL_�������) = "Z" And (Val(.TextMatrix(i, COL_��������)) = 9 Or Val(.TextMatrix(i, COL_��������)) = 10) Then
                            str��ǰ���� = Get���˵�ǰ����(Val(.TextMatrix(i, COL_����ID)), Val(.TextMatrix(i, COL_��ҳID)))
                            Exit For
                        End If
                    End If
                Next
            End With
        End If
        gcnOracle.BeginTrans: blnTran = True
        Do While Not rsSQL.EOF
            Call zlDatabase.ExecuteProcedure(rsSQL!Sql, Me.Caption)
            rsSQL.MoveNext
        Loop
                
        '����LIS����ӿ�
        If strAdviceInfo <> "" Then
            If gobjLIS.SendLisApplicationForm(strAdviceInfo, strDiag) = False Then
                gcnOracle.RollbackTrans: blnTran = False
                Screen.MousePointer = 0
                Call Del��������
                MsgBox "����ӿڵ���ʧ�ܣ����ܷ��ͼ���ҽ����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        'ҽ�������ϴ�
        strAllmsg = ""
        If Not IsNull(rsPati!����) Then
            If gclsInsure.GetCapability(supportҽ���ϴ�, rsPati!����ID, rsPati!����) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, rsPati!����ID, rsPati!����) Then
                Do While Not rsUpload.EOF
                    strMsg = "" '��Ϊ����һ��NO�ڿ϶�Ϊһ�����˵�,��������˲������Բ���
                    'strAdvance�д��롰�ܵ�����|��ǰ���������Ա�ҽ���ӿڴ���
                    If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , rsPati!����, rsUpload.RecordCount & "|" & rsUpload.AbsolutePosition) Then
                        'δ�ύǰ�ϴ�ʧ����ع�����ֹ����
                        gcnOracle.RollbackTrans: blnTran = False
                        Screen.MousePointer = 0
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName 'ÿ����ʾ
                        Else
                            MsgBox rsPati!���� & "�ķ����ϴ�ʧ�ܣ����Ͳ���������ֹ��", vbExclamation, gstrSysName
                        End If
                        Exit Function
                    Else
                        If strMsg <> "" Then strAllmsg = strAllmsg & rsUpload!NO & ":" & strMsg & vbCrLf
                    End If
                    rsUpload.MoveNext
                Loop
            End If
            
            'ҽ�������ϴ��ӿ�(������������)
            If gclsInsure.GetCapability(support�ϴ�סԺ����, rsPati!����ID, rsPati!����) Then
                If Not gclsInsure.TranElecDossier(2, rsPati!����ID, rsPati!��ҳID, rsPati!����) Then Exit Function
            End If
        End If
        If strAdvices��Ѫ <> "" Then
            strErr = ""
            For i = 0 To UBound(var��Ѫ)
                If gobjPublicBlood.AdviceOperation(pסԺҽ������, Val(var��Ѫ(i)), 5, False, strErr) = False Then
                    gcnOracle.RollbackTrans: blnTran = False
                    Screen.MousePointer = 0
                    MsgBox "Ѫ��ϵͳ�ӿڵ���ʧ�ܣ�" & strErr, vbInformation, gstrSysName
                    Exit Function
                End If
            Next
        End If
        gcnOracle.CommitTrans: blnTran = False
        If strAllmsg <> "" Then
            Screen.MousePointer = 0
            MsgBox strAllmsg, vbInformation, gstrSysName
        End If
        
        'ҽ�������ϴ�
        If Not IsNull(rsPati!����) Then
            If gclsInsure.GetCapability(supportҽ���ϴ�, rsPati!����ID, rsPati!����) And gclsInsure.GetCapability(support������ɺ��ϴ�, rsPati!����ID, rsPati!����) Then
                Do While Not rsUpload.EOF
                    strMsg = ""
                    Screen.MousePointer = 0
                    If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , rsPati!����, rsUpload.RecordCount & "|" & rsUpload.AbsolutePosition) Then
                        '�ύ���ϴ�ʧ��,����ʾ
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName
                        Else
                            MsgBox rsPati!���� & "�ļ��ʵ�""" & rsUpload!NO & """�ϴ�ʧ�ܣ�HIS���������ύ����ȷ���������͡�", vbExclamation, gstrSysName
                        End If
                    Else
                        If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                    End If
                    Screen.MousePointer = 11
                    rsUpload.MoveNext
                Loop
            End If
        End If
        
        '��Ϣ����
        If mblnSendMsg Then
            Call SendMsg����(lng���ͺ�, IIF(bln����, 1, 2), str��ǰ����)
        End If
        
        'RIS�ӿ�
        If HaveRIS Then
            If GetAdviceRis(rsAdviceRis, lng����ID, lng��ҳID) Then
                On Error Resume Next
                If gobjRis.HISSendAdvice(rsAdviceRis, 2, lng����ID, lng��ҳID, "", lng���ͺ�) <> 1 Then
                    MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISSendAdvice)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                End If
                err.Clear: On Error GoTo 0
            End If
        ElseIf gbln����Ӱ����Ϣϵͳ�ӿ� = True Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������RIS�ӿڴ���ʧ��δ����(HISSendAdvice)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
        
        '�ύ�ɹ�,������ҽ���б��Ϊ��ɾ��
        With vsAdvice
            lngS = .FindRow(CStr(rsPati!����ID), , COL_����ID)
            For i = lngS To .Rows - 1
                If Val(.TextMatrix(i, COL_����ID)) = rsPati!����ID Then
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        .RowData(i) = -1
                    End If
                Else
                    Exit For
                End If
            Next
            '������ҽӿ�
            If CreatePlugInOK(pסԺҽ������) Then
                On Error Resume Next
                Call gobjPlugIn.AdviceSend(glngSys, pסԺҽ������, lng����ID, lng��ҳID, lng���ͺ�)
                Call zlPlugInErrH(err, "AdviceSend")
                err.Clear: On Error GoTo 0
            End If
            If gobjExchange Is Nothing Then
                On Error Resume Next
                Set gobjExchange = CreateObject("zlExchange.clsExchange")
                If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
                err.Clear: On Error GoTo 0
            End If
            '�������ݽ���ƽ̨����LIS,PACS�������뵥
            If Not gobjExchange Is Nothing Then
                For i = lngS To .Rows - 1
                    If Val(.TextMatrix(i, COL_����ID)) = rsPati!����ID And Val(.TextMatrix(i, COL_ҽ��״̬)) <> 1 Then
                        If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                            'c-����,d-���
                            If .TextMatrix(i, COL_�������) = "C" Or .TextMatrix(i, COL_�������) = "D" Then
                                If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                    lng��ID = Val(.TextMatrix(i, COL_���ID))
                                Else
                                    lng��ID = Val(.TextMatrix(i, COL_ID))
                                End If
                                If InStr(strҽ��IDs & ",", "," & lng��ID & ",") = 0 Then
                                    strҽ��IDs = strҽ��IDs & "," & lng��ID
                                    Call gobjExchange.SendMsg(IIF(.TextMatrix(i, COL_�������) = "C", 1, 2), "����ID::" & rsPati!����ID & "||��ҳID::" & rsPati!����ID & "||ҽ��ID::" & lng��ID & "||��������::1")
                                End If
                            End If
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If
        End With
    End If
    CompletePatiSend = True
End Function

Private Sub SendMsg����(ByVal lng���ͺ� As Long, ByVal int�������� As Integer, ByVal str��ǰ���� As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strIDs As String
    Dim lngTmp As Long
    Dim strTmp1 As String
    Dim strTmp2 As String
    Dim str�������� As String
    Dim i As Long
    Dim j As Long
    On Error GoTo errH
    strSQL = "select ���� from ���ű� where id=[1]"
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If Val(.TextMatrix(i, COL_ִ�а���)) = 1 Then '���밲��
                    Call ZLHIS_CIS_004(mclsMipModule, Val(.TextMatrix(i, COL_����ID)), .TextMatrix(i, COL_����), .TextMatrix(i, COL_סԺ��), , IIF(Val(.TextMatrix(i, COL_��������)) = 1, 1, 2), _
                        .TextMatrix(i, COL_��ҳID), .TextMatrix(i, COL_���˿���ID), "", , .TextMatrix(i, COL_����), Val(.TextMatrix(i, COL_ID)), IIF("����" = .TextMatrix(i, COL_ҽ����Ч), 0, 1), .TextMatrix(i, COL_�������), .TextMatrix(i, COL_��������), _
                        lng���ͺ�, .TextMatrix(i, COL_ִ�п���ID))
                End If
                '����ҽ��
                If .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_��������)) = 6 Then
                    strIDs = ""
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, COL_���ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_�������) = "C" Then
                                strIDs = strIDs & "," & Val(.TextMatrix(j, COL_ID))
                                lngTmp = Val(.TextMatrix(j, COL_ִ�п���ID))
                            End If
                        End If
                    Next
                    strIDs = Mid(strIDs, 2)
                    If strIDs <> "" Then
                        Call ZLHIS_CIS_016(mclsMipModule, Val(.TextMatrix(i, COL_����ID)), .TextMatrix(i, COL_����), .TextMatrix(i, COL_סԺ��), , IIF(Val(.TextMatrix(i, COL_��������)) = 1, 1, 2), _
                        .TextMatrix(i, COL_��ҳID), Val(.TextMatrix(i, COL_���˿���ID)), , Val(.TextMatrix(i, COL_ID)), _
                            .TextMatrix(i, COL_�걾��λ), .TextMatrix(i, COL_������ĿID), , .TextMatrix(i, COL_ִ�п���ID), , strIDs, , lngTmp, , lng���ͺ�, "", _
                            int��������, .TextMatrix(i, COL_����ҽ��), .TextMatrix(i, COL_����ʱ��), .TextMatrix(i, COL_��������ID), , "")
                    End If
                '�������
                ElseIf .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    strTmp1 = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_���ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_�������) = "D" Then
                                strTmp1 = strTmp1 & "," & .TextMatrix(j, COL_�걾��λ)
                            End If
                        End If
                    Next
                    strTmp1 = Mid(strTmp1, 2)
                    Call ZLHIS_CIS_017(mclsMipModule, Val(.TextMatrix(i, COL_����ID)), .TextMatrix(i, COL_����), .TextMatrix(i, COL_סԺ��), , IIF(Val(.TextMatrix(i, COL_��������)) = 1, 1, 2), _
                        .TextMatrix(i, COL_��ҳID), Val(.TextMatrix(i, COL_���˿���ID)), "", Val(.TextMatrix(i, COL_ID)), .TextMatrix(i, COL_������ĿID), .TextMatrix(i, col_ҽ������), strTmp1, .TextMatrix(i, COL_ִ�п���ID), , lng���ͺ�, _
                        "", int��������, .TextMatrix(i, COL_����ҽ��), .TextMatrix(i, COL_����ʱ��), .TextMatrix(i, COL_��������ID), , "")
                 
                '��������
                ElseIf .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    strTmp1 = Getҽ����������(Val(.TextMatrix(i, COL_ID)), "����ҽ��")
                    strTmp2 = Getҽ����������(Val(.TextMatrix(i, COL_ID)), "����ҽ��")
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_���ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_�������) = "F" Then
                                strIDs = strIDs & "," & .TextMatrix(j, COL_ID)
                            ElseIf .TextMatrix(j, COL_�������) = "G" Then
                                lngTmp = Val(.TextMatrix(j, COL_ID))
                            End If
                        End If
                    Next
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_��������ID)))
                    If Not rsTmp.EOF Then str�������� = rsTmp!���� & ""
                    strIDs = Mid(strIDs, 2)
                    Call ZLHIS_CIS_018(mclsMipModule, Val(.TextMatrix(i, COL_����ID)), .TextMatrix(i, COL_����), .TextMatrix(i, COL_סԺ��), , IIF(Val(.TextMatrix(i, COL_��������)) = 1, 1, 2), _
                        .TextMatrix(i, COL_��ҳID), Val(.TextMatrix(i, COL_���˿���ID)), "", Val(.TextMatrix(i, COL_ID)), strIDs, , lngTmp, , strTmp1, strTmp2, .TextMatrix(i, COL_ִ�п���ID), .TextMatrix(i, COL_ִ�п���), lng���ͺ�, _
                        "", int��������, .TextMatrix(i, COL_����ҽ��), .TextMatrix(i, COL_����ʱ��), .TextMatrix(i, COL_��������ID), str��������)
                
                '��Ѫ����
                ElseIf .TextMatrix(i, COL_�������) = "K" Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_��������ID)))
                    If Not rsTmp.EOF Then str�������� = rsTmp!���� & ""
                    Call ZLHIS_CIS_019(mclsMipModule, Val(.TextMatrix(i, COL_����ID)), .TextMatrix(i, COL_����), .TextMatrix(i, COL_סԺ��), , IIF(Val(.TextMatrix(i, COL_��������)) = 1, 1, 2), _
                        .TextMatrix(i, COL_��ҳID), Val(.TextMatrix(i, COL_���˿���ID)), "", Val(.TextMatrix(i, COL_ID)), .TextMatrix(i, COL_ִ�п���ID), .TextMatrix(i, COL_ִ�п���), lng���ͺ�, _
                        "", int��������, .TextMatrix(i, COL_����ҽ��), .TextMatrix(i, COL_����ʱ��), .TextMatrix(i, COL_��������ID), str��������)
                 
                '��������
                ElseIf .TextMatrix(i, COL_�������) = "Z" And InStr(",7,8,11,", "," & .TextMatrix(i, COL_��������) & ",") > 0 _
                    Or .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_��������)) = 5 Then
                    If .TextMatrix(i, COL_��������) = "7" Then
                        strTmp1 = "ZLHIS_CIS_020"  '��������
                    ElseIf .TextMatrix(i, COL_��������) = "8" Then
                        strTmp1 = "ZLHIS_CIS_021"  '��������ҽ��
                    ElseIf .TextMatrix(i, COL_��������) = "11" Then
                        strTmp1 = "ZLHIS_CIS_022"  '��������ҽ��
                    ElseIf .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_��������)) = 5 Then
                        strTmp1 = "ZLHIS_CIS_023"  '������������ҽ��
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_��������ID)))
                    If Not rsTmp.EOF Then str�������� = rsTmp!���� & ""
                    Call SendMsg(strTmp1, mclsMipModule, Val(.TextMatrix(i, COL_����ID)), .TextMatrix(i, COL_����), .TextMatrix(i, COL_סԺ��), , IIF(Val(.TextMatrix(i, COL_��������)) = 1, 1, 2), _
                        .TextMatrix(i, COL_��ҳID), Val(.TextMatrix(i, COL_���˿���ID)), "", Val(.TextMatrix(i, COL_ID)), .TextMatrix(i, COL_ִ�п���ID), .TextMatrix(i, COL_ִ�п���), lng���ͺ�, _
                        "", int��������, .TextMatrix(i, COL_����ҽ��), .TextMatrix(i, COL_����ʱ��), .TextMatrix(i, COL_��������ID), str��������)
                 
                'סԺ����Ԥ��Ժ
                ElseIf .TextMatrix(i, COL_�������) = "Z" And Val(.TextMatrix(i, COL_��������)) = 5 Then
                    Call GetPatChange(Val(.TextMatrix(i, COL_ID)), 10, lngTmp, strTmp1)
                    Call ZLHIS_PATIENT_009(mclsMipModule, Val(.TextMatrix(i, COL_����ID)), .TextMatrix(i, COL_��ҳID), .TextMatrix(i, COL_����), .TextMatrix(i, col_�Ա�), .TextMatrix(i, COL_סԺ��), _
                        lngTmp, .TextMatrix(i, COL_����ʱ��), Val(.TextMatrix(i, COL_���˲���ID)), , Val(.TextMatrix(i, COL_���˿���ID)), "", , .TextMatrix(i, COL_����), Val(.TextMatrix(i, COL_ID)))
                
                'סԺ���߲�����
                ElseIf .TextMatrix(i, COL_�������) = "Z" And (Val(.TextMatrix(i, COL_��������)) = 9 Or Val(.TextMatrix(i, COL_��������)) = 10) Then
                    Call GetPatChange(Val(.TextMatrix(i, COL_ID)), 13, lngTmp, strTmp1)
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_���˿���ID)))
                    strTmp2 = "": If Not rsTmp.EOF Then strTmp2 = rsTmp!���� & ""
                    Call ZLHIS_PATIENT_005(mclsMipModule, Val(.TextMatrix(i, COL_����ID)), .TextMatrix(i, COL_��ҳID), .TextMatrix(i, COL_����), .TextMatrix(i, col_�Ա�), .TextMatrix(i, COL_סԺ��), _
                        Val(.TextMatrix(i, COL_���˲���ID)), , Val(.TextMatrix(i, COL_���˿���ID)), strTmp2, str��ǰ����, lngTmp, .TextMatrix(i, COL_����ʱ��), strTmp1, .TextMatrix(i, COL_����ҽ��), Val(.TextMatrix(i, COL_ID)))
                 
                'סԺ����ת������
                ElseIf .TextMatrix(i, COL_�������) = "Z" And Val(.TextMatrix(i, COL_��������)) = 3 Then
                    Call GetPatChange(Val(.TextMatrix(i, COL_ID)), 3, lngTmp, strTmp1)
                    Call ZLHIS_PATIENT_003(mclsMipModule, Val(.TextMatrix(i, COL_����ID)), .TextMatrix(i, COL_��ҳID), .TextMatrix(i, COL_����), .TextMatrix(i, col_�Ա�), .TextMatrix(i, COL_סԺ��), _
                        Val(.TextMatrix(i, COL_���˲���ID)), , Val(.TextMatrix(i, COL_���˿���ID)), "", , .TextMatrix(i, COL_סԺ��), _
                        lngTmp, .TextMatrix(i, COL_����ʱ��), , , Val(.TextMatrix(i, COL_ִ�п���ID)), , Val(.TextMatrix(i, COL_ID)))
                End If
            End If
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub DeleteSendRow()
'���ܣ���������ҽ���嵥���ѷ��ͳɹ��ĵ���ɾ��
    Dim i As Long, blnDel As Boolean
    
    With vsAdvice
        .Redraw = flexRDNone
        For i = .Rows - 1 To .FixedRows Step -1
            If .RowData(i) = -1 Then .RemoveItem i: blnDel = True
        Next
        .Redraw = flexRDDirect
        
        If blnDel Then
            If .Rows = .FixedRows Then .Rows = .FixedRows + 1
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: .Col = COL_ѡ��
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            
            vsPrice.Rows = vsPrice.FixedRows
            vsPrice.Rows = vsPrice.FixedRows + 1
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
        End If
    End With
End Sub

Private Function Getʵ�ս��(ByVal strSQL As String) As Currency
    Dim lngPos As Long, strMatch As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strSQL = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    strMatch = "End" & Chr(0) & Chr(1)
    strSQL = Left(strSQL, InStr(strSQL, strMatch) - 1)
    Getʵ�ս�� = CCur(strSQL)
End Function

Private Function Setʵ�ս��(ByVal strSQL As String, ByVal cur��� As Currency) As String
    Dim strLeft As String, strRight As String
    Dim strMatch As String, strVal As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strLeft = Mid(strSQL, 1, InStr(strSQL, strMatch) - 1)
    strMatch = "End" & Chr(0) & Chr(1)
    strRight = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    
    Setʵ�ս�� = strLeft & cur��� & strRight
End Function

Private Function GetMergeDrugStore(ByVal lngRow As Long) As Long
'���ܣ���ȡһ����ҩ�Ļ�׼ҩ�����������ɷ���NO��Keyֵ
'˵����һ����ҩ��ҩƷ���͵�һ�𣬰����Ա�ҩ�Ͳ�ͬҩ�������
    Dim lngҩ��ID As Long, lngBegin As Long, i As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_���ID)) <> Val(.TextMatrix(lngRow - 1, COL_���ID)) And Val(.TextMatrix(lngRow, COL_ִ�п���ID)) <> 0 Then
            lngҩ��ID = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
        Else
            lngBegin = .FindRow(.TextMatrix(lngRow, COL_���ID), , COL_���ID)
            For i = lngBegin To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    If Val(.TextMatrix(i, COL_ִ�п���ID)) <> 0 Then
                        lngҩ��ID = Val(.TextMatrix(i, COL_ִ�п���ID)): Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    GetMergeDrugStore = lngҩ��ID
End Function

Private Sub SeekPriceRow(ByVal lngRow As Long, ByVal lng��ĿID As Long, ByVal int�������� As Integer, ByVal lngCol As Long)
'���ܣ���λ������ʾָ��ҽ����ָ���Ƽ���
'������lngRow=ҽ���к�
'      lng��ĿID=�Ƽ���ĿID
'      lngCol=�Ƽ۱����ʾ��
    Dim k As Long
    
    With vsAdvice
        .Col = col_ҽ������ '�������Զ�ShowPrice,mrsPrice�����仯
        If Not .RowHidden(lngRow) Then
            .Row = lngRow
        Else
            If InStr(",F,D,G,C,", .TextMatrix(lngRow, COL_�������)) > 0 And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                '��������,��������,��鲿λ,���������Ŀ
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_���ID))), , COL_ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 1 Then
                '��ҩ;��
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_���ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 2 Then
                '��ҩ�巨
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_���ID))), lngRow + 1, COL_ID)
            End If
        End If
        For k = vsPrice.FixedRows To vsPrice.Rows - 1
            If Val(vsPrice.TextMatrix(k, COLP_�к�)) = lngRow _
                And Val(vsPrice.TextMatrix(k, COLP_��������)) = int�������� _
                And Val(vsPrice.TextMatrix(k, COLP_�շ�ϸĿID)) = lng��ĿID Then
                vsPrice.Row = k: vsPrice.Col = lngCol: Exit For
            End If
        Next
        Call .ShowCell(.Row, .Col)
        Call vsPrice.ShowCell(vsPrice.Row, vsPrice.Col)
    End With
End Sub

Public Function SendAdvice() As Long
'���ܣ�����ҽ������(��������м��ʱ���)
'˵����������˷����ύ
'���أ�����ɹ����򷵻ط��ͺ�
'rsSQL!����=1-У��(�������Ҫ��У��),2-ҽ���Ƽ�,3-סԺ����,4-ִ�п����滻��5-ҽ�����ͣ�6-�Զ�����,7-��Һ��ҩ
    Dim rsPati As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset
    
    Dim rsSQL As ADODB.Recordset '������֯SQL���Ķ�̬��¼��
    Dim rsTotal As ADODB.Recordset '���ڿ����ܼ��Ķ�̬��¼��
    Dim rsUpload As ADODB.Recordset '�����ռ�ҽ���ϴ����ݺŵĶ�̬��¼��
    Dim rsNumber As ADODB.Recordset '������������Ķ�̬��¼��
    Dim rsItems As ADODB.Recordset '����ҽ���ܿصķ��ü�¼��,��̬��¼��
    Dim rsMoneyNow As ADODB.Recordset '��ǰ���˱���Ҫ���͵ķ���,��̬��¼��
    Dim rsMoneyDay As ADODB.Recordset '��ǰ���˵����ѷ��͵ķ���,��̬��¼��
    Dim rsAudit As ADODB.Recordset     'ҽ��������¼��
    Dim rsExec As ADODB.Recordset  'ҽ��ִ�мƼ�
    Dim rsClone As ADODB.Recordset, rsSeek As ADODB.Recordset '�û��ܴ��ۼ���Ķ�̬��¼��
    Dim rsDays As ADODB.Recordset  '������ҩ�����ַ���
    Dim rsSendDays As ADODB.Recordset '������ҩ�����ַ��ͼ�¼
    Dim rsDayNo As ADODB.Recordset '���ͺţ������֡���һ���Ӧһ�����ͺ�
    Dim rsTmp As ADODB.Recordset
    
    Dim i As Long, j As Long
    Dim strSQL As String, strTmp As String
    Dim blnTran As Boolean, strCurDate As String, strCurDateTmp As String
    Dim str��� As String, str���� As String
    
    Dim lng����ID As Long, lng��ҳID As Long, lng�������� As Long
    Dim lng���ͺ� As Long, int�Ʒ�״̬ As Integer, bln���� As Boolean, int���� As Integer, strNO As String
    Dim str�շ���Ŀ As String, lng������� As Long, lng���ø��� As Long, lng������� As Long, lng��ID As Long, lngOld��ID As Long
    Dim int���� As Integer, dbl���� As Double, cur�ϼ� As Currency, cur���ʺϼ� As Currency
    Dim dbl���� As Double, dblӦ�� As Double, curӦ�� As Currency, curʵ�� As Currency
    Dim bln������Ŀ�� As Boolean, lng���մ���ID As Long, curͳ���� As Currency, str���ձ��� As String, str�������� As String
    Dim str�ֽ�ʱ�� As String, str�״�ʱ�� As String, strĩ��ʱ�� As String
    Dim int�䷽�� As Integer, strNOKey As String, str�Զ����� As String
    Dim str����ʱ�� As String, str�Ǽ�ʱ�� As String
    Dim dbl�������� As Double, blnFirst As Boolean '�䷽�����ֺŹؼ���
    Dim lng���˿���ID As Long, lngִ�п���ID As Long, intִ��״̬ As Integer
    Dim bln�������� As Boolean, intҩƷ���� As Integer, blnBool As Boolean
    
    Dim strHaveSub As String, strNoneSub As String
    Dim int����� As Integer, lng����ĿID As Long, strʵ�� As String
    
    Dim blnҩƷʱ����ʾ As Boolean, blnҩƷ�����ʾ As Boolean, blnҩƷĬ�Ϸ��� As Boolean
    Dim bln����ʱ����ʾ As Boolean, bln���Ŀ����ʾ As Boolean, bln����Ĭ�Ϸ��� As Boolean
    Dim blnҩƷ������ʾ As Boolean
    Dim str��ҩ�� As String, strCuvetteNumber As String '��������
    
    Dim strAudit As String
    Dim blnʵʱ��� As Boolean, blnSend As Boolean, blnOldSend As Boolean, blnSendPrivs As Boolean
    Dim lng���ô��� As Long 'һ��ֻ��һ��ʱ�����η���Ӧ��ȡ�ķ��ô���
    Dim strMsg As String, blnExists As Boolean, lngSpecialAdviceID As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim str���� As String, str����ҽ�� As String
    Dim lng��ҽ���� As Long
    Dim lng�ɼ�����ID As Long
    Dim lng���ͺ�Tmp As Long
    Dim strPlugIn���ͺ� As String '��ҽӿ���
    Dim lngLastPatiID As Long, str��ҩIDs As String, lngLastPageID As Long, lngLastPatiDeptID As Long
    Dim str��λ���� As String '�����Ŀ�Ĳ�λ�������̶���ʽ����鲿λ<sTab>��鷽�����磺"ͷ��<sTab>ƽɨ"
    Dim dblOther���� As Double '������Ŀ�շѴ���
    Dim str����ҩ��  As String '������ҩƷ��ҽ�� ,"Ƥ��ҽ��ID,ҩƷ��ҽ��ID"
    Dim rsƤ�� As ADODB.Recordset
    Dim strMinDate As String
    
    On Error GoTo errH
    
    '�ȼ�鲢��ʾ����ҽ��:3-ת��,5-��Ժ,6-תԺ,11-����
    With vsAdvice
        If InitObjRecipeAudit(pסԺҽ���´�) Then
            '�������ϵͳ������������
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    If (.TextMatrix(i, COL_�������) = "5" Or .TextMatrix(i, COL_�������) = "6") And Val(.TextMatrix(i, COL_ҽ��״̬)) = 1 Then
                        If lngLastPatiID <> Val(.TextMatrix(i, COL_����ID)) Then
                            If Mid(str��ҩIDs, 2) <> "" Then
                                Call gobjRecipeAudit.BuildData(Mid(str��ҩIDs, 2), lngLastPatiDeptID, 1, lngLastPatiID, lngLastPageID, strTmp)
                                str��ҩIDs = ""
                            End If
                        End If
                        lngLastPatiID = Val(.TextMatrix(i, COL_����ID))
                        lngLastPageID = Val(.TextMatrix(i, COL_��ҳID))
                        lngLastPatiDeptID = Val(.TextMatrix(i, COL_���˿���ID))
                        If InStr("," & str��ҩIDs & ",", "," & .TextMatrix(i, COL_���ID) & ",") = 0 Then str��ҩIDs = str��ҩIDs & "," & .TextMatrix(i, COL_���ID)
                    End If
                End If
            Next
            If Mid(str��ҩIDs, 2) <> "" Then
                Call gobjRecipeAudit.BuildData(Mid(str��ҩIDs, 2), lngLastPatiDeptID, 1, lngLastPatiID, lngLastPageID, strTmp)
            End If
        End If
        
        strTmp = ""
        strMinDate = "3000-01-01 00:00"
        blnExists = False
        Call InitExecRecordset(rsExec)   'ҽ��ִ�мƼ�
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If .TextMatrix(i, COL_�������) = "Z" And InStr(",3,5,6,11,", Val(.TextMatrix(i, COL_��������))) > 0 Then
                    strTmp = strTmp & vbCrLf & .TextMatrix(i, COL_����) & IIF(.Cell(flexcpData, i, COL_Ӥ��) <> 0, "(Ӥ��" & .Cell(flexcpData, i, COL_Ӥ��) & ")", "") & "��" & .TextMatrix(i, col_ҽ������)
                    
                    If mblnCheckAdvice Then
                        lngSpecialAdviceID = Val(.TextMatrix(i, COL_ID))
                        strMsg = CheckUnExecutedAdvice(Val(.TextMatrix(i, COL_����ID)), Val(.TextMatrix(i, COL_��ҳID)), lngSpecialAdviceID, Val(.Cell(flexcpData, i, COL_Ӥ��)))
                        If strMsg <> "" Then
                            If blnExists = False Then blnExists = True
                            
                            Call GetPatiRowScope(i, lngBegin, lngEnd)
                            'ȡ����ѡ
                            For j = lngBegin To lngEnd
                                If .Cell(flexcpData, j, COL_ѡ��) = 0 Then
                                    Set .Cell(flexcpPicture, j, COL_ѡ��) = Nothing
                                End If
                            Next
                        End If
                    End If

                    'ת��ҽ������ʱ�жϳ������Լ�������
                    If Val(.TextMatrix(i, COL_��������)) = 3 Then
                        If CheckCanSendAdvice(Val(.TextMatrix(i, COL_����ID)), Val(.TextMatrix(i, COL_��ҳID)), lngSpecialAdviceID, Val(.Cell(flexcpData, i, COL_Ӥ��))) Then
                            Call MsgBox("����ת��ҽ����" & vbCrLf & .TextMatrix(i, COL_����) & IIF(.Cell(flexcpData, i, COL_Ӥ��) <> 0, "(Ӥ��" & .Cell(flexcpData, i, COL_Ӥ��) & ")", "") & "��" & .TextMatrix(i, col_ҽ������) & vbCrLf & vbCrLf & "���뽫���Է��͵ĳ���ҽ���������ܷ��͡�", vbInformation, gstrSysName)
                            Exit Function
                        End If
                    End If
                    
                    'ת��ʱδ������ʵ��ݼ��
                    If Val(.TextMatrix(i, COL_��������)) = 3 Then
                        If CheckWaitQuittance(Val(.TextMatrix(i, COL_����ID)), Val(.TextMatrix(i, COL_��ҳID))) Then Exit Function
                    End If
                    
                End If
                
                 '3-ת��;4-����;5-��Ժ;6-תԺ,11-����,14-��ǰ
                If Mid(gstrESign, 2, 1) = "1" Then  'סԺҽ��վ�����˵���ǩ���ż��
                    If .TextMatrix(i, COL_�������) = "Z" And InStr(",3,4,5,6,11,14,", Val(.TextMatrix(i, COL_��������))) > 0 Then
                        If InStr(str���� & ",", "," & .TextMatrix(i, COL_����ID) & ":" & .TextMatrix(i, COL_��ҳID) & ",") = 0 Then
                            str���� = str���� & "," & .TextMatrix(i, COL_����ID) & ":" & .TextMatrix(i, COL_��ҳID)
                        End If
                        
                        str����ҽ�� = str����ҽ�� & vbCrLf & .TextMatrix(i, COL_����) & IIF(.Cell(flexcpData, i, COL_Ӥ��) <> 0, "(Ӥ��" & .Cell(flexcpData, i, COL_Ӥ��) & ")", "") & "��" & .TextMatrix(i, col_ҽ������)
                    End If
                End If
                
                '���������У��ģʽ�������¿�ҽ���Ƿ񲢷��޸���(Ϊ������ܣ�ֻ���һ���е�����¼����Ϊһ��ҽ�����޸�ʱ������ͬ��)
                If Val(.TextMatrix(i, COL_ҽ��״̬)) = 1 Then
                    If mblnAutoVerify And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                        If CheckAdviceUpdate(Val(.TextMatrix(i, COL_ID)), .TextMatrix(i, COL_�¿�����ʱ��)) Then
                            MsgBox "ҽ����" & .TextMatrix(i, col_ҽ������) & vbCrLf & "�Ѿ����޸ģ������¶�ȡҽ�����ٷ��͡�", vbInformation, "����ҽ������"
                            Exit Function
                        End If
                    End If
                End If
                
                '��������ж���Ϣ�ռ�
                If gbln����ҩƷ�ֿ����� Then
                    If cboDrugType.ListIndex = 0 Then
                        If InStr("," & str���� & ",", "," & .TextMatrix(i, COL_�������) & ",") = 0 Then
                            str���� = str���� & "," & .TextMatrix(i, COL_�������)
                        End If
                    ElseIf cboDrugType.ListIndex = 3 Then
                        str���� = ""
                    Else
                        str���� = ",����ҩ"
                    End If
                End If
                If .TextMatrix(i, COL_�״�ʱ��) < strMinDate Then
                    strMinDate = .TextMatrix(i, COL_�״�ʱ��)
                End If
            End If
        Next
        If strMinDate = "3000-01-01 00:00" Then strMinDate = ""
        
        If str���� <> "" And cboDrugType.ListIndex = 0 Then
            If Not (str���� = ",����ҩ" Or str���� = ",����I��" Or str���� = ",����ҩ" Or str���� = ",����ҩ,����I��" Or str���� = ",����I��,����ҩ") Then
                If Not (InStr(str���� & ",", ",����ҩ,") = 0 And InStr(str���� & ",", ",����ҩ,") = 0 And InStr(str���� & ",", ",����I��,") = 0) Then
                    MsgBox "���η��͵�ҽ���п��ܰ������龫��ҩƷ����ֱ��ͣ����޸Ĺ����������¶�ȡҽ�����ٷ��͡�", vbInformation, gstrSysName
                    Exit Function
                Else
                    str���� = ""
                End If
            End If
        End If
        
        If strTmp <> "" Then
            If mblnCheckAdvice And blnExists Then
                Call MsgBox("������������ҽ����" & vbCrLf & strTmp & vbCrLf & vbCrLf & "���뽫" & strMsg & "�������ܷ��͡�", vbInformation, gstrSysName)
                
            ElseIf MsgBox("Ҫ���͵�ҽ���а�����������ҽ����" & vbCrLf & strTmp & vbCrLf & vbCrLf & "ȷʵҪ���͵�ǰѡ���ҽ����", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            If MsgBox("ȷʵҪ���͵�ǰѡ���ҽ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        
        '��������˵���ǩ����������"��ֹͣ��δȷ��ֹͣ"��ҽ������ʾ��ʿ�Ƚ���ȷ��ֹͣ
        '��Ϊ����ҽ��У��ʱ�Ὣ"��ֹͣ��δȷ��ֹͣ"��ҽ����"ִ����ֹʱ��"����Ϊ����ҽ���Ŀ�ʼִ��ʱ�䣬ҽ��ֹͣ��ǩ��Դ�İ�����"ִ����ֹʱ��"����ᵼ��ǩ����֤�޷�ͨ��
        If str���� <> "" Then
            str���� = Mid(str����, 2)
            strTmp = ""
            If CheckStopedUnAffirm(str����, strTmp) Then
                MsgBox "Ҫ���͵�ҽ���а�����������ҽ����" & vbCrLf & str����ҽ�� & _
                    vbCrLf & vbCrLf & "���ͺ�Ὣδȷ��ֹͣ��ҽ������ֹͣ��Ϊ�˲�Ӱ��ǩ����֤�����ȶ����²��˽���ȷ��ֹͣ������" & strTmp, vbInformation, gstrSysName
                Exit Function
            End If
            strTmp = ""
        End If
        
        '������ҩ
        If mbln������ҩ Then
            blnBool = Set������ҩ()
            If Not blnBool Then
                GoTo FuncEnd
            End If
        End If
        
        If Not zlPluginAdviceBeforeSend Then
            Exit Function
        End If
    End With
    
    'RISԤԼ����ж���ʾ
    Call CheckRISScheduling
    
    Screen.MousePointer = 11
    
    blnSendPrivs = InStr(GetInsidePrivs(pסԺҽ������), "ȫԺҽ������") > 0
    mstrRollNotify = "": mstr��ҩ�� = ""
    blnҩƷʱ����ʾ = True: blnҩƷ�����ʾ = True: blnҩƷĬ�Ϸ��� = True
    bln����ʱ����ʾ = True: bln���Ŀ����ʾ = True: bln����Ĭ�Ϸ��� = True
    blnҩƷ������ʾ = True
    mbln�������Ѻ��� = False
    mintBnt = -1
    
    Call InitBillSet
    lng���ͺ� = zlDatabase.GetNextNo(10)        '���ȫ�����¿���������ָ������ʱ�����޷��ͣ�����Ϊ�㣩����ִ�з���ʱ���˷�һ����
    mlngNOSequence = 0 '���ݺ��������³�ʼ
    mdatCurr = zlDatabase.Currentdate
    strCurDateTmp = Format(mdatCurr, "yyyy-MM-dd HH:mm:ss")
    strCurDate = "To_Date('" & strCurDateTmp & "','YYYY-MM-DD HH24:MI:SS')"
    int�䷽�� = 1 '��ʾ���͵ĵڼ����䷽,���ڷֵ��ݺ�
    'һ��һ�����ͺ�
    Set rsDayNo = New ADODB.Recordset
    rsDayNo.Fields.Append "���ͺ�", adBigInt
    rsDayNo.Fields.Append "����ʱ��", adVarChar, 100
    rsDayNo.CursorLocation = adUseClient
    rsDayNo.LockType = adLockOptimistic
    rsDayNo.CursorType = adOpenStatic
    rsDayNo.Open
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                lng��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                                   
                '�¿��ĳ�������ȡҽ��ʱ������ָ���Ľ���ʱ�������跢�͵ģ�����Ϊ�㣩
                '����¼��������ͳ���������
                '���ⳤ��ֻУ�Բ�����:����ȼ�,����/Σҽ��,��¼�����ҽ��������(���û�л���ҽ����֮ǰû�е���Ҫ����У��)
                blnSend = True
                If Val(.TextMatrix(i, COL_ҽ��״̬)) = 1 Then   '�¿�ҽ��
                    If lng��ID = lngOld��ID Then
                        blnSend = blnOldSend
                    Else
                        If Val(.Cell(flexcpData, i, COL_ҽ����Ч)) = 0 And Val(.TextMatrix(i, COL_����)) = 0 Or _
                            .TextMatrix(i, COL_�������) = "" And Val("" & .TextMatrix(i, COL_������ĿID)) = 0 Then
                            blnSend = False
                        ElseIf Val(.Cell(flexcpData, i, COL_ҽ����Ч)) = 0 Then
                            If .TextMatrix(i, COL_�������) = "H" And .TextMatrix(i, COL_��������) = "1" And .TextMatrix(i, COL_ִ��Ƶ��) = "2" Then
                                blnSend = False
                            ElseIf .TextMatrix(i, COL_�������) = "Z" And InStr(",9,10,12,", "," & .TextMatrix(i, COL_��������) & ",") > 0 Then
                                blnSend = False
                            End If
                        End If
                        If Not blnSendPrivs And blnSend Then
                            If Not CheckSendPrivs(Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_����ID)), Val(.TextMatrix(i, COL_��ҳID)), Val(.TextMatrix(i, COL_����ҽ��ID))) Then
                                blnSend = False
                            End If
                        End If
                    End If
                End If
                blnOldSend = blnSend
                
                '�ύ��ǰ���˵�����
                '-----------------------------------------------------------------------------------------
                If Val(.TextMatrix(i, COL_����ID)) <> lng����ID Then
                    '�ύ��ǰ��������
                    If lng����ID <> 0 Then
                        If strAudit <> "" Then
                            MsgBox "����""" & rsPati!���� & """���·�����Ŀ��û�о�����������Ӧ��ҽ�����ܷ��ͣ�" & vbCrLf & strAudit, vbInformation, gstrSysName
                            GoTo errH
                        End If
                                    
                        
                         'ҽ��ִ�мƼ�
                        If rsExec.RecordCount > 0 Then
                            rsExec.MoveFirst
                            Do While Not rsExec.EOF
                                rsSQL.AddNew
                                rsSQL!���� = 8
                                rsSQL!��ĿID = 0
                                rsSQL!��� = 0
                                rsSQL!ҽ��ID = lng��ID
                                rsSQL!Sql = "Zl_ҽ��ִ�мƼ�_Insert(" & rsExec!ҽ��ID & "," & rsExec!���ͺ� & ",To_date('" & _
                                rsExec!Ҫ��ʱ�� & "','yyyy-MM-dd HH24:mi:ss')," & ZVal(Val(rsExec!�շ�ϸĿID & "")) & "," & rsExec!���� & "," & rsExec!�������� & ")"
                                rsExec.MoveNext
                            Loop
                        End If
                    
                        If Not CompletePatiSend(rsPati, rsSQL, rsUpload, rsItems, cur�ϼ�, cur���ʺϼ�, str���, bln����, blnTran, lng���ͺ�) Then GoTo errH
                        SendAdvice = lng���ͺ� 'ֻҪ�ύ�ɹ����ע
                        Call InitExecRecordset(rsExec)   'ҽ��ִ�мƼ�
                    End If
                    
                    '���ò�����ر���
                    str�Զ����� = ""
                    lng����ID = Val(.TextMatrix(i, COL_����ID))
                    lng��ҳID = Val(.TextMatrix(i, COL_��ҳID))
                    lng�������� = Val(.TextMatrix(i, COL_��������))
                    Set rsƤ�� = Nothing
                    Call InitRecordSet(rsSQL, rsTotal, rsUpload, rsNumber, rsMoneyNow, rsItems, rsSendDays) '����SQL����
                    cur�ϼ� = 0:  str��� = "":  cur���ʺϼ� = 0 '���ñ�������
                    Set rsMoneyDay = Nothing
                    
                    '��ȡ��ǰ������Ϣ
                    strSQL = _
                        " Select ����ID,Ԥ�����,�������,0 as Ԥ����� From ������� Where ����=1 And ���� = 2 And ����ID=[1]" & _
                        " Union ALL" & _
                        " Select A.����ID,0,0,Sum(���) From ����ģ����� A,������ҳ B" & _
                        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.���� Is Not Null And A.����ID=[1] And A.��ҳID=[2] Group by A.����ID"
                    strSQL = "Select ����ID,Nvl(Sum(Ԥ�����),0)-Nvl(Sum(�������),0)+Nvl(Sum(Ԥ�����),0) as ʣ��� From (" & strSQL & ") Group by ����ID"
                    
                    '״̬:0-������1-��δ��ƣ�2-����ת�ƣ�3-��Ԥ��Ժ
                    strSQL = "Select A.����ID,B.��ҳID,NVL(B.����,A.����) ����,B.����,B.״̬,Nvl(B.�Ա�,A.�Ա�) as �Ա�,Nvl(B.����,A.����) as ����,B.��Ժ���� as ����,a.�����," & _
                        " zl_PatiWarnScheme(A.����ID,B.��ҳID) as ���ò���,C.ʣ���,a.��ͥ�绰 as PhoneNO,a.��ͥ��ַ as Address,To_Char(A.��������,'YYYY-MM-DD HH24:MI:SS') as Birthdate," & _
                        " Decode(A.������,Null,Null,zl_PatientSurety(A.����ID,B.��ҳID)) as ������,b.��ǰ����ID as ����ID,b.��Ժ����ID as ����ID,B.סԺ��" & _
                        " From ������Ϣ A,������ҳ B,(" & strSQL & ") C" & _
                        " Where A.����ID=B.����ID And A.����ID=C.����ID(+) And A.����ID=[1] And B.��ҳID=[2]"
                    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
                    
                    If blnSend Then
                        '��ȡ��ǰ���˵�������Ŀ�嵥
                        strAudit = ""
                        If Not IsNull(rsPati!����) Then
                            If Val(zlDatabase.GetPara("���ҽ������", glngSys, pסԺҽ������, "1")) = 1 Then
                                Set rsAudit = GetAuditRecord(lng����ID, lng��ҳID)
                            Else
                                Set rsAudit = Nothing
                            End If
                            blnʵʱ��� = gclsInsure.GetCapability(supportʵʱ���, rsPati!����ID, rsPati!����)
                        Else
                            Set rsAudit = Nothing '��NothingΪ��־�ò��˲���Ҫ�ж�
                            blnʵʱ��� = False
                        End If
                        
                        '�����²���鵱ǰ����ҽ����ҩƷ���,�Ա�ҩ�����
                        '��Ȼ��ȡʱ�ѻ��ܼ�飬����Ʒ����ʱ������˹����ܷ����仯
                        For j = i To .Rows - 1
                            If Val(.TextMatrix(j, COL_����ID)) = lng����ID Then
                                '���ܸ���ǰ��������ʾ�Ľ�������Ѳ�����
                                If .Cell(flexcpData, j, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, j, COL_ѡ��) Is Nothing Then
                                    If InStr(",5,6,7,", .TextMatrix(j, COL_�������)) > 0 And Val(.TextMatrix(j, COL_ִ������ID)) <> 5 Then
                                        '�ڲ����ֹ�������,����������ʱ��ҩƷ
                                        If TheStockCheck(Val(.TextMatrix(j, COL_ִ�п���ID)), .TextMatrix(j, COL_�������)) = 2 _
                                            Or Val(.TextMatrix(j, COL_ҩ������)) = 1 Or Val(.TextMatrix(j, COL_�Ƿ���)) = 1 Then
                                            .TextMatrix(j, COL_���) = Format(GetStock(Val(.TextMatrix(j, COL_�շ�ϸĿID)), Val(.TextMatrix(j, COL_ִ�п���ID)), 2), "0.00000")
                                            If CheckStock(j, blnҩƷ�����ʾ, blnҩƷʱ����ʾ, blnҩƷĬ�Ϸ���, True) Then
                                                Call RowSelectSame(j, COL_ѡ��)
                                            End If
                                        End If
                                        If CheckDrug����(j, blnҩƷ������ʾ) Then
                                            Call RowSelectSame(j, COL_ѡ��)
                                        End If
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If
                                    
                '���ܸ���ǰ��������ʾ�Ľ�������Ѳ�����
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                                         
                    '����ҽ����3-ת��;5-��Ժ;6-תԺ,11-����
                    If .TextMatrix(i, COL_�������) = "Z" Then
                        'ת��,��Ժ,תԺ,����ҽ������ʱ������Ҫ��������״̬
                        If blnSend And .Cell(flexcpData, i, COL_Ӥ��) = 0 Then
                            If InStr(",3,5,6,11,", .TextMatrix(i, COL_��������)) > 0 And NVL(rsPati!״̬, 0) <> 0 Then
                                MsgBox "����""" & rsPati!���� & """��ǰ����""" & decode(NVL(rsPati!״̬, 0), 1, "�ȴ����", 2, "����ת��", 3, "��Ԥ��Ժ") & """״̬��" & _
                                    "���ܷ���""" & .TextMatrix(i, col_ҽ������) & """ҽ����", vbInformation, gstrSysName
                                Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                                Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                GoTo NextAdvice
                            End If
                        End If
                        
                        '�����ת�ơ���Ժ��תԺҽ��,��鲡���Ƿ���δִ�е�ҽ����Ŀ��δ��ҩƷ
                        If InStr(",3,", .TextMatrix(i, COL_��������)) > 0 Then
                            If Not CheckWaitExecute(rsPati, i, gbytת�Ƽ��δִ��, gbytת�Ƽ��δ��ҩ) Then
                                Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                                Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                GoTo NextAdvice
                            End If
                        End If
                        If InStr(",5,6,", .TextMatrix(i, COL_��������)) > 0 Then
                            If Not CheckWaitExecute(rsPati, i, gbyt��Ժ���δִ��, gbyt��Ժ���δ��ҩ) Then
                                Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                                Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                GoTo NextAdvice
                            End If
                        End If
                        
                        '��Ϊ�Զ�ֹͣҽ������Ҫ���г����ջ�����
                        If InStr(",3,5,6,11,", .TextMatrix(i, COL_��������)) > 0 Then
                            If InStr(mstrRollNotify, lng����ID & ":" & lng��ҳID) = 0 Then
                                mstrRollNotify = mstrRollNotify & "," & lng����ID & ":" & lng��ҳID
                            End If
                        End If
                    End If
                    
                    '����ҽ����ִ�п���
                    If .Cell(flexcpData, i, COL_ִ�п���ID) = 1 Then
                        rsSQL.AddNew
                        rsSQL!���� = 4
                        rsSQL!ҽ��ID = lng��ID
                        rsSQL!��ĿID = 0
                        rsSQL!��� = i
                        rsSQL!Sql = "ZL_ҽ��ִ�п���_Update(" & Val(.TextMatrix(i, COL_ID)) & "," & ZVal(.TextMatrix(i, COL_ִ�п���ID)) & ",1)"
                        rsSQL.Update
                    End If
                    
                    If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, lng����ID, lng��ҳID, "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                    
                    '����ҽ�����ʷ���:�����¼۸����
                    '-----------------------------------------------------------------------------------------
                    strSQL = "": str�շ���Ŀ = ""
                    If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                        'ҩƷȱʡ�̶�Ϊ�����Ƽ�,����ҽ��ʱָ����Ϊ�Ա�ҩ(Ժ��ִ��)�Ĳ���ȡ;ҩƷ������Ϊ����
                        If Val(.TextMatrix(i, COL_ִ������ID)) <> 5 Then
                            strSQL = _
                                " Select A.ID,A.���,D.���� as �������,A.����,A.���㵥λ,B.������ĿID," & _
                                " C.�վݷ�Ŀ,Y.סԺ��λ,Y.סԺ��װ,Y.����ϵ��,1 as ����,B.�ּ� as ����," & _
                                " A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,B.�����շ���,A.�Ƿ���,Y.ҩ������ as ����,0 as ��������," & _
                                " 0 as ����,[2] as ִ�п���ID,A.���ηѱ�,A.����ȷ��,0 as ��������,0 as �շѷ�ʽ,I.Ҫ������" & _
                                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ���Ŀ��� D,ҩƷ��� Y,����֧����Ŀ I" & _
                                " Where A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID And A.���=D.����" & _
                                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "4", "5", "6") & _
                                " And A.ID=Y.ҩƷID(+) And A.ID=[1] And A.ID=I.�շ�ϸĿID(+) And I.����(+)=[3]" & _
                                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                                " Order by A.����"
                        End If
                    Else
                        '���Ƽ�,�ֹ��Ƽۣ�����,Ժ��ִ�е�ҽ������ȡ
                        If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                            '��ɾ��ԭ��ҩҽ���ļƼ�(����ʱ��У�Ե�ģʽ��û�б�Ҫ��ɾ������Ϊ֮ǰû�в����Ƽ�)
                            If Val(.Cell(flexcpData, i, COL_���)) = 1 And Val(.TextMatrix(i, COL_ҽ��״̬)) <> 1 Then
                                rsSQL.AddNew
                                rsSQL!���� = 2: rsSQL!��ĿID = 0: rsSQL!��� = i
                                rsSQL!ҽ��ID = lng��ID
                                rsSQL!Sql = "ZL_����ҽ���Ƽ�_Delete(" & Val(.TextMatrix(i, COL_ID)) & ",1)"
                                rsSQL.Update
                            End If
                        
                            mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                            If Not mrsPrice.EOF Then
                                For j = 1 To mrsPrice.RecordCount
                                    If NVL(mrsPrice!�շ�ϸĿID, 0) <> 0 And NVL(mrsPrice!����, 0) <> 0 Then '��������Ϊ0���Զ����˵�
                                        '��ͨ��Ŀ�ı�۵���Ҫ�����룬�����Ǹ������õ�ʱ������ҽ��
                                        If NVL(mrsPrice!����, 0) = 0 And NVL(mrsPrice!���, 0) = 1 _
                                            And Not (InStr(",5,6,7,", mrsPrice!�շ����) > 0 Or mrsPrice!�շ���� = "4" And NVL(mrsPrice!����, 0) = 1) Then
                                            Call SeekPriceRow(i, mrsPrice!�շ�ϸĿID, mrsPrice!��������, COLP_����)
                                            Screen.MousePointer = 0
                                            MsgBox "����Ϊ��۵��շ���Ŀȷ��һ���շѼ۸�", vbInformation, gstrSysName
                                            vsPrice.SetFocus: GoTo FuncEnd
                                        End If
                                        
                                        '�Ƽ�ִ�п���:ֻ�����ҩƷ������ҽ���ģ�ҩƷ�����ļƼ۵�ִ�п���
                                        If InStr(",4,5,6,7,", .TextMatrix(i, COL_�������)) = 0 _
                                            And (InStr(",5,6,7,", mrsPrice!�շ����) > 0 Or mrsPrice!�շ���� = "4" And NVL(mrsPrice!����, 0) = 1) Then
                                            lngִ�п���ID = NVL(mrsPrice!ִ�п���ID, 0)
                                            
                                            '���ı�������ִ�п���
                                            If lngִ�п���ID = 0 And mrsPrice!�շ���� = "4" Then
                                                Call SeekPriceRow(i, mrsPrice!�շ�ϸĿID, mrsPrice!��������, COLP_ִ�п���)
                                                Screen.MousePointer = 0
                                                MsgBox "����""" & vsPrice.TextMatrix(vsPrice.Row, COLP_�շ���Ŀ) & """û��ȷ��ִ�п��ң����ֹ�������ȷ��ִ�п��ҡ�" & vbCrLf & _
                                                    "�������ȷ����ȷ��ִ�п��ң��뵽""����Ŀ¼����""�м��洢�ⷿ�����Ƿ���ȷ��", vbInformation, gstrSysName
                                                vsPrice.SetFocus: GoTo FuncEnd
                                            End If
                                        Else
                                            lngִ�п���ID = 0
                                        End If
                                        
                                        'ҩƷ������ҽ���ļƼ۹̶���Ӧ�����棻�Ǹ������õ�ʱ�����ĵı����Ҫ���룬���Ҫ���浽�Ƽ۱���
                                        If Val(.Cell(flexcpData, i, COL_���)) = 1 Or Val(.TextMatrix(i, COL_ҽ��״̬)) = 1 Then
                                            If InStr(",4,5,6,7,", .TextMatrix(i, COL_�������)) = 0 _
                                                Or .TextMatrix(i, COL_�������) = "4" And NVL(mrsPrice!����, 0) = 0 And NVL(mrsPrice!���, 0) = 1 Then
                                                rsSQL.AddNew
                                                rsSQL!���� = 2: rsSQL!��ĿID = mrsPrice!�շ�ϸĿID: rsSQL!��� = i
                                                rsSQL!ҽ��ID = lng��ID
                                                rsSQL!Sql = "ZL_����ҽ���Ƽ�_INSERT(" & _
                                                    mrsPrice!ҽ��ID & "," & mrsPrice!�շ�ϸĿID & "," & _
                                                    NVL(mrsPrice!����, 0) & "," & NVL(mrsPrice!����, 0) & "," & _
                                                    NVL(mrsPrice!����, 0) & "," & ZVal(lngִ�п���ID) & "," & _
                                                    NVL(mrsPrice!��������, 0) & "," & NVL(mrsPrice!�շѷ�ʽ, 0) & ")"
                                                rsSQL.Update
                                            End If
                                        End If
                                        
                                        '��ʱ����ҽ���Ƽ۱�
                                        If Val(.TextMatrix(i, COL_����)) <> 0 Then '��Ѫ����û������
                                            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                                                "Select " & mrsPrice!�շ�ϸĿID & " as �շ�ϸĿID," & _
                                                NVL(mrsPrice!ִ�п���ID, 0) & " as ִ�п���ID," & _
                                                NVL(mrsPrice!����, 0) & " as ����," & Format(NVL(mrsPrice!����, 0), gstrDecPrice) & " as ����," & _
                                                NVL(mrsPrice!����, 0) & " as ����," & NVL(mrsPrice!��������, 0) & " as ��������," & _
                                                NVL(mrsPrice!�շѷ�ʽ, 0) & " as �շѷ�ʽ From Dual"
                                        End If
                                    End If
                                    
                                    mrsPrice.MoveNext
                                Next
                            End If
                        End If
                        
                        If strSQL <> "" Then
                            strSQL = _
                                " Select A.ID,A.���,D.���� as �������,A.����,A.���㵥λ,A.�Ƿ���," & _
                                " A.���ηѱ�,A.����ȷ��,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,B.�����շ���,Y.סԺ��λ,Y.סԺ��װ,Y.����ϵ��," & _
                                " Decode(A.���,'4',E.���÷���,Y.ҩ������) as ����,E.��������,B.������ĿID," & _
                                " C.�վݷ�Ŀ,X.����,Decode(A.�Ƿ���,1,X.����,B.�ּ�) as ����,X.ִ�п���ID," & _
                                " X.����,X.��������,X.�շѷ�ʽ,I.Ҫ������" & _
                                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ���Ŀ��� D,�������� E,(" & strSQL & ") X,ҩƷ��� Y,����֧����Ŀ I" & _
                                " Where A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID And A.ID=E.����ID(+)" & _
                                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "4", "5", "6") & _
                                " And A.���=D.���� And X.�շ�ϸĿID=A.ID And A.ID=Y.ҩƷID(+)" & _
                                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                                " And A.ID=I.�շ�ϸĿID(+) And I.����(+)=[3]" & _
                                " Order by X.��������,X.����,X.�շѷ�ʽ Desc,A.ID"
                                'һ��Ҫ����������ǰ��,�Ա��ڼ�����ڷ��ü�¼�б������ӹ�ϵ
                        End If
                    End If
                    
                    'ҽ��У��,����ǰ�Զ�У��(һ��ҽ������һ�Σ����ж�����ҪУ��)
                    If mblnAutoVerify Then
                        If Val(.TextMatrix(i, COL_ҽ��״̬)) = 1 And lng��ID <> lngOld��ID Then
                            rsSQL.AddNew
                            rsSQL!���� = 1
                            rsSQL!ҽ��ID = lng��ID
                            rsSQL!��ĿID = 0
                            rsSQL!��� = i
                            rsSQL!Sql = "ZL_����ҽ����¼_У��(" & lng��ID & ",3," & strCurDate & ",Null,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
                        End If
                    End If
                    
                    
                    
                    'ִ�з��ͺͼ��ʷ���
                    '-----------------------------------------
                    If blnSend Then
                        '�����ۿ۱�����ʼ
                        strHaveSub = "": strNoneSub = ""
                        int����� = 0: lng����ĿID = 0
                        Call InitSeekSet(rsSeek)
                        
                        '��ǰ������������(����"ҽ����������������"û������ʱҲ����һ����������룬�����ж��Ƿ��ղ�Ѫ�ܷ���)
                        strCuvetteNumber = ""
                        If Val(.TextMatrix(i, COL_ִ������ID)) <> 0 Then
                            j = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                            If j > 0 Then lng�ɼ�����ID = Val(.TextMatrix(j, COL_ִ�п���ID))
                            strCuvetteNumber = GetCuvetteNumber(rsNumber, .TextMatrix(i, COL_�Թܱ���), _
                                Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)), .TextMatrix(i, COL_�������), Val(.TextMatrix(i, COL_��������)), _
                                Val(.TextMatrix(i, COL_ִ�п���ID)), Val(.TextMatrix(i, COL_Ӥ��)), Val(.TextMatrix(i, COL_������ĿID)), _
                                Val(.TextMatrix(i, COL_������־)), .TextMatrix(i, COL_�걾��λ), lng�ɼ�����ID)
                        End If
                        
                        int�Ʒ�״̬ = IIF(Val(.TextMatrix(i, COL_�Ƽ�����)) = 1, -1, 0) '����Ʒѻ�δ�Ʒ�
                    
                
                        '�������ݺŷ���ؼ���
                        '-----------------------------------------------------------------------------------------
                        If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                            '���ò���������ҩƷ�ֿ����� ʱ������ҩƷҽ����ҩƷ�е������ɵ��ݺţ�һ��ҽ������һ����
                            If str���� <> "" Then
                                strNOKey = "������ҩ_" & .TextMatrix(i, COL_���ID)
                            Else
                                '������ҩ��"����(����ID,��ҳID)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
                                'һ����ҩ�ģ����͵�һ�𣺰����Ա�ҩ�Ͳ�ͬҩ�������
                                strNOKey = "������ҩ_" & lng����ID & "_" & lng��ҳID & "_" & .TextMatrix(i, COL_ҽ����Ч) & "_" & _
                                    Val(.TextMatrix(i, COL_���˿���ID)) & "_" & Val(.TextMatrix(i, COL_��������ID)) & "_" & _
                                    .TextMatrix(i, COL_����ҽ��) & "_" & GetMergeDrugStore(i)
                                '�ٰ�Ҫ��ӡ�����Ƶ��ݷֺ�
                                strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_������ĿID)), 2)
                                '��ҩִ�п��Ҳ���ͬ������䲻ͬ��NO��
                                j = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                                If j > 0 Then strNOKey = strNOKey & "_" & Val(.TextMatrix(j, COL_ִ�п���ID))
                            End If
                        ElseIf .TextMatrix(i, COL_�������) = "7" Then
                            'һ���䷽�е����в�ҩ����һ���������ݺ�
                            strNOKey = "��ҩ�䷽_" & lng����ID & "_" & lng��ҳID & "_" & .TextMatrix(i, COL_ҽ����Ч) & "_" & int�䷽��
                        ElseIf InStr(",4,M,", .TextMatrix(i, COL_�������)) > 0 Then
                            '���ϰ�"����(����ID,��ҳID)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
                            strNOKey = "����ҽ��_" & lng����ID & "_" & lng��ҳID & "_" & .TextMatrix(i, COL_ҽ����Ч) & "_" & _
                                Val(.TextMatrix(i, COL_���˿���ID)) & "_" & Val(.TextMatrix(i, COL_��������ID)) & "_" & _
                                .TextMatrix(i, COL_����ҽ��) & "_" & Val(.TextMatrix(i, COL_ִ�п���ID))
                            '�ٰ�Ҫ��ӡ�����Ƶ��ݷֺ�
                            strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_������ĿID)), 2)
                            
                        ElseIf Val(.TextMatrix(i, COL_���ID)) <> 0 And .TextMatrix(i, COL_�������) = "C" Then
                            'һ���ɼ��ļ�����Ϸ�����ͬ�ĵ��ݺţ��걾�ɼ��������䵥���ĵ��ݺ�
                            'ͬһ��������ͣ�ͬһ������ִ�п��ң�ͬһ�ɼ��ܣ�ͬһ���ɼ���ʽ��ͬһ���ɼ�ִ�п��ҵļ��������ͬ�ĵ��ݺ�
                            If mbln���鵥���������� Then
                                strNOKey = "һ���ɼ�_" & Val(.TextMatrix(i, COL_���ID))
                            Else
                                lng��ҽ���� = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                                strNOKey = "һ���ɼ�_" & lng����ID & "_" & lng��ҳID & "_" & .TextMatrix(i, COL_�걾��λ) & "_" & _
                                    .TextMatrix(i, COL_ִ�п���ID) & "_" & .TextMatrix(i, COL_��������) & "_" & .TextMatrix(i, COL_�Թܱ���) & "_" & _
                                    .TextMatrix(lng��ҽ����, COL_������ĿID) & "_" & .TextMatrix(lng��ҽ����, COL_ִ�п���ID)
                            End If
                        ElseIf Val(.TextMatrix(i, COL_���ID)) <> 0 And InStr(",F,D,", .TextMatrix(i, COL_�������)) > 0 Then
                            '��鲿λ�͸�����������Ҫҽ��������ͬ���ݺţ�����������䵥���ĵ��ݺš�
                            strNOKey = "��ҩҽ��_" & Val(.TextMatrix(i, COL_���ID))
                        Else
                            '������ҩҽ��ÿ��ҽ��һ���������ݺ�(������ҩ;�����䷽�巨���÷����ɼ���ʽ������ʽ����Ѫҽ��/��Ѫ;��)
                            strNOKey = "��ҩҽ��_" & Val(.TextMatrix(i, COL_ID))
                        End If
                        
                                
                         '�ֽ�ʱ��
                        If .TextMatrix(i, COL_�ֽ�ʱ��) <> "" Then
                            str�ֽ�ʱ�� = .TextMatrix(i, COL_�ֽ�ʱ��)
                        Else
                            str�ֽ�ʱ�� = .Cell(flexcpData, i, COL_�ֽ�ʱ��)    '��ʼִ��ʱ��
                        End If
                        If str�ֽ�ʱ�� = "" Then str�ֽ�ʱ�� = .TextMatrix(i, COL_��ʼʱ��)
                        If Len(str�ֽ�ʱ��) > 4000 Then
                            Screen.MousePointer = 0
                            MsgBox "��ǰ���͵�ҽ��ʱ�䷶Χ̫��,����ִ��" & CStr(UBound(Split(str�ֽ�ʱ��, ",")) + 1) & "�Ρ�" & vbCrLf & _
                                "������֧�ֵ�������" & CStr(UBound(Split(Mid(str�ֽ�ʱ��, 1, 4000), ",")) + 1) & "��,���������ʱ������·��ͣ�", vbInformation, gstrSysName
                            Call DeleteSendRow: Call ShowSendTotal
                            Progress = 0: Exit Function
                        End If
                        
                        If Val(.TextMatrix(i, COL_���ID)) <> 0 And .TextMatrix(i, COL_�������) = "D" Then
                            str��λ���� = .TextMatrix(i, COL_�걾��λ) & "<sTab>" & .TextMatrix(i, COL_��鷽��)
                        Else
                            str��λ���� = ""
                        End If
                        
                        '�������ʷ���
                        '------------------------------------------------------
                        If strSQL <> "" Then
                            '�Ƿ���Ժ��ҩ
                            intҩƷ���� = 0
                            If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                                intҩƷ���� = decode(.TextMatrix(i, COL_ִ������), "��Ժ��ҩ", 3, "��ȡҩ", 4, intҩƷ����)
                            ElseIf .TextMatrix(i, COL_�������) = "7" Then
                                j = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                                If j <> -1 Then
                                    intҩƷ���� = decode(.TextMatrix(j, COL_ִ������), "��Ժ��ҩ", 3, "��ȡҩ", 4, intҩƷ����)
                                End If
                            End If
                        
                            Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_�շ�ϸĿID)), Val(.TextMatrix(i, COL_ִ�п���ID)), Val(NVL(rsPati!����, 0)), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                            If Not rsPrice.EOF Then
                                int�Ʒ�״̬ = 1 '�ѼƷ�
                                Set rsClone = rsPrice.Clone
                            End If
    
                            '����������Ŀ���ķ�����ϸ
                            bln�������� = .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) <> 0
                            Do While Not rsPrice.EOF
MoneyItemBegin:
                                'ִ�п���ID
                                lngִ�п���ID = NVL(rsPrice!ִ�п���ID, 0)
                                '��ԭֵ������ȡ��Ч�ķ�ҩ��ҩƷ���������ĵ�ִ�п���
                                If rsPrice!��� = "4" And NVL(rsPrice!��������, 0) = 1 _
                                    Or InStr(",5,6,7", rsPrice!���) > 0 And InStr(",5,6,7", .TextMatrix(i, COL_�������)) = 0 Then
                                    lng���˿���ID = Val(.TextMatrix(i, COL_���˿���ID))
                                    lngִ�п���ID = Get�շ�ִ�п���ID(rsPati!����ID, rsPati!��ҳID, rsPrice!���, rsPrice!ID, 4, lng���˿���ID, 0, 2, lngִ�п���ID, , , 2)
                                    
                                    '���ı�������ִ�п���
                                    If lngִ�п���ID = 0 And rsPrice!��� = "4" Then
                                        .Row = GetVisibleRow(i, True)
                                        Call .ShowCell(.Row, .Col)
                                        Screen.MousePointer = 0
                                        MsgBox "ϵͳ����Ϊ�Ƽ�����""" & rsPrice!���� & """ȷ�����ʵ�ִ�п��ҡ�" & vbCrLf & _
                                            "��ʹ�üƼ۵���������Ϊȷ������""����Ŀ¼����""�м��洢�ⷿ�����Ƿ���ȷ��", vbInformation, gstrSysName
                                        Call DeleteSendRow: Call ShowSendTotal
                                        Progress = 0: Exit Function
                                    End If
                                End If
                                
                                '----------------------------------------
                                '�����շѷ�ʽ��ȷ����ǰ�շ���Ŀ�Ƿ�Ӧ�շ�
                                If rsPrice!�������� & "_" & rsPrice!ID <> str�շ���Ŀ Then
                                    If Not AdviceMoneyMake(lng����ID, lng��ҳID, rsMoneyNow, rsMoneyDay, _
                                        lng��ID, Val(.TextMatrix(i, COL_������ĿID)), rsPrice!ID, lngִ�п���ID, .TextMatrix(i, COL_�Թܱ���), _
                                        rsPrice!���, NVL(rsPrice!�շѷ�ʽ, 0), str�ֽ�ʱ��, 2, lng���ô���, Val(.TextMatrix(i, COL_����)), _
                                        Val(.TextMatrix(i, COL_ID)), lng���ͺ�, Val(rsPrice!���� & ""), rsExec, Val(.TextMatrix(i, COL_���㷽ʽ)), _
                                        .TextMatrix(i, COL_Ƶ��), Val(.TextMatrix(i, COL_����)), .Cell(flexcpData, i, COL_ҽ����Ч), rsPrice!��������, .TextMatrix(i, COL_�������), strCuvetteNumber, str��λ����, dblOther����, strMinDate) Then
                                        '������ǰ�շ���Ŀ(���������Ŀ)
                                        str�շ���Ŀ = rsPrice!�������� & "_" & rsPrice!ID
                                        Do While rsPrice!�������� & "_" & rsPrice!ID = str�շ���Ŀ
                                            rsPrice.MoveNext
                                            If rsPrice.EOF Then Exit Do
                                        Loop
                                        If rsPrice.EOF Then Exit Do
                                        GoTo MoneyItemBegin
                                    End If
                                End If
                                '----------------------------------------
                                
                                '����Ƿ���Ҫ���Ѿ�����
                                If NVL(rsPrice!Ҫ������, 0) = 1 And Not rsAudit Is Nothing Then
                                    rsAudit.Filter = "��ĿID=" & rsPrice!ID
                                    If rsAudit.EOF Then
                                        If UBound(Split(strAudit, vbCrLf)) < 10 Then
                                            If InStr(strAudit, "��" & rsPrice!����) = 0 Then
                                                strAudit = strAudit & vbCrLf & "��" & rsPrice!����
                                            End If
                                        ElseIf UBound(Split(strAudit, vbCrLf)) = 10 Then
                                            strAudit = strAudit & vbCrLf & "�� ��"
                                        End If
                                    End If
                                End If
                                
                                If InStr(",5,6,7", rsPrice!���) > 0 Then
                                    If InStr(",5,6,7", .TextMatrix(i, COL_�������)) > 0 Then
                                        If .TextMatrix(i, COL_�������) = "7" Then
                                            int���� = Val(.TextMatrix(i, COL_����))
                                            '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                                            If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                                                dbl���� = Val(.TextMatrix(i, COL_����)) / NVL(rsPrice!����ϵ��, 1)
                                            Else
                                                dbl���� = IntEx(Val(.TextMatrix(i, COL_����)) / NVL(rsPrice!����ϵ��, 1) / NVL(rsPrice!סԺ��װ, 1)) * NVL(rsPrice!סԺ��װ, 1)
                                            End If
                                        Else
                                            int���� = 1
                                            dbl���� = Val(.TextMatrix(i, COL_����)) * NVL(rsPrice!סԺ��װ, 1)
                                            If rsƤ�� Is Nothing Then
                                                Set rsƤ�� = GetԭҺƤ��(lng����ID, lng��ҳID, "")
                                            End If
                                            rsƤ��.Filter = "ҩƷID=" & Val(rsPrice!ID & "")
                                            If Not rsƤ��.EOF Then
                                                If Val(rsƤ��!��� & "") = 0 Then
                                                    '���м���������
                                                    dbl���� = (Val(.TextMatrix(i, COL_����)) - 1) * NVL(rsPrice!סԺ��װ, 1)
                                                    rsƤ��!��� = Val(.TextMatrix(i, COL_ID))
                                                    
                                                    str����ҩ�� = "'" & rsƤ��!Ƥ��ҽ��ID & "," & rsƤ��!��� & "'"
                                                    rsƤ��.Update
                                                    If dbl���� <= 0 Then
                                                        rsPrice.MoveNext
                                                        If rsPrice.EOF Then Exit Do
                                                        GoTo MoneyItemBegin
                                                    End If
                                                End If
                                            End If
                                            
                                        End If
                                    Else
                                        int���� = 1
                                        '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                                        '��ҩ��ҩƷ�Ƽ�:��Ϊ����Ԥ�����ۼ�����,��˲��������㴦��
                                        '�����շѶ����е�ҩƷ����Ϊ����ֻ��ȡһ�Σ�����Ϊ���ô���*��������
                                        If InStr(",2,3,4,5,6,7,9,", Val("" & rsPrice!�շѷ�ʽ)) > 0 Then
                                            If dblOther���� > 0 Then
                                                dbl���� = Format(dblOther����, "0.00000")
                                            Else
                                                dbl���� = Format(lng���ô��� * NVL(rsPrice!����, 0), "0.00000")
                                            End If
                                        Else
                                            dbl���� = Val(.TextMatrix(i, COL_����)) * NVL(rsPrice!����, 0)
                                        End If
                                    End If
                                    dbl���� = Format(dbl����, "0.00000")
                                    
                                    If NVL(rsPrice!�Ƿ���, 0) = 1 Then
                                        dbl���� = Format(CalcDrugPrice(rsPrice!ID, lngִ�п���ID, int���� * dbl����, , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                                    Else
                                        dbl���� = Format(NVL(rsPrice!����, 0), gstrDecPrice)
                                    End If
                                ElseIf rsPrice!��� = "4" And NVL(rsPrice!��������, 0) = 1 Then
                                    '�����������������
                                    If mlng�������ID = 0 Then
                                        Screen.MousePointer = 0
                                        MsgBox "����ȷ���������ϵ��ݵ�������,���ȵ���������������ã�", vbInformation, gstrSysName
                                        Call DeleteSendRow: Call ShowSendTotal
                                        Progress = 0: Exit Function
                                    End If
                                    
                                    int���� = 1
                                    If InStr(",1,2,3,4,5,6,7,9,", Val("" & rsPrice!�շѷ�ʽ)) > 0 Then
                                        If dblOther���� > 0 Then
                                            dbl���� = Format(dblOther����, "0.00000")
                                        Else
                                            dbl���� = Format(lng���ô��� * NVL(rsPrice!����, 0), "0.00000")
                                        End If
                                    Else
                                        dbl���� = Format(Val(.TextMatrix(i, COL_����)) * NVL(rsPrice!����, 0), "0.00000")
                                    End If
                                    
                                    '����ʱ�����ĵ���
                                    If NVL(rsPrice!�Ƿ���, 0) = 1 Then
                                        dbl���� = Format(CalcDrugPrice(rsPrice!ID, lngִ�п���ID, dbl����, , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                                    Else
                                        dbl���� = Format(NVL(rsPrice!����, 0), gstrDecPrice)
                                    End If
                                Else
                                    '�������ڵ������������Ρ�һ��ֻ��һ��ʱ���ж�����Ҫִ�У����ն��ٴΣ����ܵ������������磺ÿ�����Σ�,��Ҫ���շѶ��յĴ���
                                    int���� = 1
                                    If InStr(",1,2,3,4,5,6,7,9,", Val("" & rsPrice!�շѷ�ʽ)) > 0 Then
                                        If dblOther���� > 0 Then
                                            dbl���� = Format(dblOther����, "0.00000")
                                        Else
                                            dbl���� = Format(lng���ô��� * NVL(rsPrice!����, 0), "0.00000")
                                        End If
                                    Else
                                        dbl���� = Format(Val(.TextMatrix(i, COL_����)) * NVL(rsPrice!����, 0), "0.00000")
                                    End If
                                    dbl���� = Format(NVL(rsPrice!����, 0), gstrDecPrice)
                                End If
                                
                                '��ҩ��ҩƷ���������ĵĿ����
                                If rsPrice!��� = "4" And NVL(rsPrice!��������, 0) = 1 _
                                    Or InStr(",5,6,7", rsPrice!���) > 0 And InStr(",5,6,7", .TextMatrix(i, COL_�������)) = 0 Then
                                    If TheStockCheck(lngִ�п���ID, rsPrice!���) <> 0 Or NVL(rsPrice!�Ƿ���, 0) = 1 Or NVL(rsPrice!����, 0) = 1 Then
                                        If rsPrice!��� = "4" Then
                                            blnBool = CheckPriceStock(i, rsPrice, lngִ�п���ID, int���� * dbl����, rsTotal, bln���Ŀ����ʾ, bln����ʱ����ʾ, bln����Ĭ�Ϸ���)
                                        Else
                                            blnBool = CheckPriceStock(i, rsPrice, lngִ�п���ID, int���� * dbl����, rsTotal, blnҩƷ�����ʾ, blnҩƷʱ����ʾ, blnҩƷĬ�Ϸ���)
                                        End If
                                        If blnBool Then
                                            Call RowSelectSame(i, COL_ѡ��, rsSQL, rsTotal, rsUpload)
                                            Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                            GoTo NextAdvice
                                        End If
                                    End If
                                End If
                                
                                '���ͽ��
                                dblӦ�� = int���� * dbl���� * dbl����
                                If bln�������� Then
                                    dblӦ�� = dblӦ�� * NVL(rsPrice!�����շ���, 100) / 100
                                End If
                                
                                '����Ӱ�Ӽ�
                                If gbln�Ӱ�Ӽ� And NVL(rsPrice!�Ӱ�Ӽ�, 0) = 1 Then
                                    dblӦ�� = dblӦ�� * (1 + NVL(rsPrice!�Ӱ�Ӽ���, 0) / 100)
                                End If
                                
                                curӦ�� = Format(dblӦ��, gstrDec)
                                                            
                                'NO,���---------------------------------------------------------------------
                                Call GetCurBillSet(strNOKey, strNO, lng�������, -1)
                                rsSQL.AddNew: blnBool = False
                                If rsPrice!�������� & "_" & rsPrice!ID <> str�շ���Ŀ Then
                                    lng���ø��� = lng�������
                                    If rsPrice!���� = 0 Then
                                        '��¼������Ϣ������϶��ڴ���ǰ
                                        '��ʹ�������ۿۣ�ҲҪ��¼�������ϵ
                                        If InStr(strHaveSub & ",", "," & rsPrice!�������� & ",") = 0 _
                                            And InStr(strNoneSub & ",", "," & rsPrice!�������� & ",") = 0 Then
                                            rsClone.Filter = "��������=" & rsPrice!�������� & " And ����=1"
                                            If Not rsClone.EOF Then
                                                int����� = lng�������
                                                lng����ĿID = rsPrice!ID
                                                
                                                rsSeek.AddNew
                                                rsSeek!�������� = rsPrice!��������
                                                rsSeek!�����ǩ = rsSQL.Bookmark 'Variant(Double)
                                                rsSeek!������ID = rsPrice!������ĿID
                                                rsSeek.Update
                                                strHaveSub = strHaveSub & "," & rsPrice!��������
                                                
                                                blnBool = True
                                            Else
                                                strNoneSub = strNoneSub & "," & rsPrice!��������
                                            End If
                                        End If
                                    End If
                                End If
                                
                                '��������ۿۺϼ�
                                If gbln��������ۿ� And (rsPrice!���� = 1 Or InStr(strHaveSub & ",", "," & rsPrice!�������� & ",") > 0) Then
                                    curʵ�� = curӦ��
                                    
                                    '�ۼ�ҽ���ϼ��������ۿ�
                                    rsSeek.Filter = "��������=" & rsPrice!��������
                                    rsSeek!�ϼ� = NVL(rsSeek!�ϼ�, 0) + curʵ��
                                    rsSeek.Update
                                ElseIf NVL(rsPrice!���ηѱ�, 0) = 0 Then
                                    curʵ�� = Format(ActualMoney(.TextMatrix(i, COL_�ѱ�), rsPrice!������ĿID, curӦ��, rsPrice!ID, lngִ�п���ID, _
                                        int���� * dbl����, IIF(gbln�Ӱ�Ӽ� And NVL(rsPrice!�Ӱ�Ӽ�, 0) = 1, NVL(rsPrice!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                                Else
                                    curʵ�� = curӦ��
                                End If
                                If gbln��������ۿ� And blnBool Then
                                    '�����ۿ�ʱ���������ʵ�ս�������⴦��
                                    strʵ�� = Chr(0) & Chr(1) & "Begin" & curʵ�� & "End" & Chr(0) & Chr(1)
                                Else
                                    strʵ�� = curʵ��
                                End If
                                '----------------------------------------------------------------------------
                                
                                'ҽ������ֶ�
                                bln������Ŀ�� = False: lng���մ���ID = 0: curͳ���� = 0: str���ձ��� = "": str�������� = ""
                                If Not IsNull(rsPati!����) Then
                                    strTmp = gclsInsure.GetItemInsure(lng����ID, rsPrice!ID, curʵ��, False, rsPati!����, .Cell(flexcpData, i, COL_ҽ������) & "||" & int���� * dbl����)
                                    If strTmp <> "" Then
                                        bln������Ŀ�� = Val(Split(strTmp, ";")(0)) <> 0
                                        lng���մ���ID = Val(Split(strTmp, ";")(1))
                                        curͳ���� = Format(Val(Split(strTmp, ";")(2)), gstrDec)
                                        str���ձ��� = CStr(Split(strTmp, ";")(3))
                                        If UBound(Split(strTmp, ";")) >= 5 Then
                                            If Split(strTmp, ";")(5) <> "" Then
                                                str�������� = Split(strTmp, ";")(5)
                                            End If
                                        End If
                                    End If
                                End If
                                
                                '�ռ����ʱ������
                                cur�ϼ� = cur�ϼ� + curʵ��
                                If InStr(str���, rsPrice!���) = 0 Then
                                    str��� = str��� & rsPrice!���
                                End If
                                                            
                                '�Ƿ񻮼�
                                strTmp = mlng����ID
                                If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                                    int���� = IIF(InStr(gstrסԺ���ͻ��۵�, "5") > 0, 1, 0)
                                    '�����72503
                                    If mintסԺ��ҩ���� = 0 Then
                                        j = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                                        If Val(.TextMatrix(j, COL_ִ�п���ID)) <> 0 Then strTmp = Val(.TextMatrix(j, COL_ִ�п���ID))
                                    End If
                                Else
                                    int���� = IIF(InStr(gstrסԺ���ͻ��۵�, .TextMatrix(i, COL_�������)) > 0, 1, 0)
                                End If
                                If int���� = 0 Then int���� = IIF(NVL(rsPrice!����ȷ��, 0) = 1, 1, 0)
                                
                                If int���� = 0 Or intִ��״̬ = 1 Then
                                    bln���� = False
                                    cur���ʺϼ� = cur���ʺϼ� + curʵ��
                                End If
                                '����ʱ��
                                If intҩƷ���� = 3 Then
                                    str����ʱ�� = strCurDate
                                ElseIf .TextMatrix(i, COL_�ֽ�ʱ��) <> "" Then
                                    str����ʱ�� = "To_Date('" & Split(.TextMatrix(i, COL_�ֽ�ʱ��), ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                                Else
                                    str����ʱ�� = "To_Date('" & .Cell(flexcpData, i, COL_�ֽ�ʱ��) & "','YYYY-MM-DD HH24:MI:SS')"
                                End If
                                
                                '�Ǽ�ʱ��
                                If int���� = 1 Then '��ǻ��۵�ʱ�������ֿ�
                                    str�Ǽ�ʱ�� = "To_Date('" & Format(DateAdd("s", 1, mdatCurr), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                Else
                                    str�Ǽ�ʱ�� = strCurDate
                                End If
                                
                                '������ҩ�����ַ���
                                Set rsDays = New ADODB.Recordset
                                rsDays.Fields.Append "�շ�ʱ��", adVarChar, 10
                                rsDays.Fields.Append "����", adDouble
                                rsDays.Fields.Append "ʵ�ս��", adDouble
                                rsDays.Fields.Append "Ӧ�ս��", adDouble
                                rsDays.Fields.Append "����ʱ��", adVarChar, 100
                                rsDays.Fields.Append "NO", adVarChar, 30
                                rsDays.CursorLocation = adUseClient
                                rsDays.LockType = adLockOptimistic
                                rsDays.CursorType = adOpenStatic
                                rsDays.Open
                                '������ʳ�����������ơ��������ƿɲ��
                                If .TextMatrix(i, COL_ҽ����Ч) = "����" And (InStr(",I,L,", .TextMatrix(i, COL_�������)) > 0 Or .TextMatrix(i, COL_�������) = "E" And InStr(",0,5,", "," & .TextMatrix(i, COL_��������) & ",") > 0) Then
                                    rsExec.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID)) & " And ���ͺ�=" & lng���ͺ� & " And �շ�ϸĿID=" & rsPrice!ID
                                    If rsExec.RecordCount > 0 Then rsExec.MoveFirst
                                    Do While Not rsExec.EOF
                                        If Val(rsExec!���� & "") <> 0 Then
                                            rsDays.Filter = "�շ�ʱ��='" & Format(rsExec!Ҫ��ʱ�� & "", "YYYY-MM-DD") & "'"
                                            If rsDays.RecordCount = 0 Then
                                                rsSendDays.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID)) & " And ����ʱ��='" & Format(rsExec!Ҫ��ʱ�� & "", "yyyy-MM-dd") & "'"
                                                If rsSendDays.RecordCount > 0 Then rsSendDays.MoveFirst
                                                rsDayNo.Filter = 0
                                                If rsDayNo.RecordCount = 0 Then
                                                    lng���ͺ�Tmp = lng���ͺ�
                                                    rsDayNo.AddNew
                                                    rsDayNo!���ͺ� = lng���ͺ�Tmp
                                                    rsDayNo!����ʱ�� = Format(rsExec!Ҫ��ʱ�� & "", "yyyy-MM-dd")
                                                    rsDayNo.Update
                                                Else
                                                    rsDayNo.Filter = "����ʱ��='" & Format(rsExec!Ҫ��ʱ�� & "", "yyyy-MM-dd") & "'"
                                                    If rsDayNo.RecordCount = 0 Then
                                                        lng���ͺ�Tmp = zlDatabase.GetNextNo(10)
                                                        rsDayNo.AddNew
                                                        rsDayNo!���ͺ� = lng���ͺ�Tmp
                                                        rsDayNo!����ʱ�� = Format(rsExec!Ҫ��ʱ�� & "", "yyyy-MM-dd")
                                                        rsDayNo.Update
                                                    Else
                                                        lng���ͺ�Tmp = rsDayNo!���ͺ�
                                                    End If
                                                End If
                                                Call GetCurBillSet(strNOKey & "_" & Format(rsExec!Ҫ��ʱ�� & "", "yyyy-MM-dd"), strNO, -1, -1)
                                                rsDays.AddNew
                                                rsDays!�շ�ʱ�� = Format(rsExec!Ҫ��ʱ�� & "", "YYYY-MM-DD")
                                                rsDays!���� = rsExec!����
                                                rsDays!ʵ�ս�� = Format(rsExec!���� / dbl���� * strʵ��, gstrDec)
                                                rsDays!Ӧ�ս�� = Format(rsExec!���� / dbl���� * curӦ��, gstrDec)
                                                rsDays!����ʱ�� = "To_Date('" & Format(rsExec!Ҫ��ʱ�� & "", "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                                rsDays!NO = strNO
                                                If rsSendDays.RecordCount = 0 Then
                                                    rsSendDays.AddNew
                                                    rsSendDays!ҽ��ID = Val(.TextMatrix(i, COL_ID))
                                                    rsSendDays!���ͺ� = lng���ͺ�Tmp
                                                    rsSendDays!����ʱ�� = Format(rsExec!Ҫ��ʱ�� & "", "yyyy-MM-dd")
                                                    rsSendDays!�ֽ�ʱ�� = Get����ֽ�ʱ��(str�ֽ�ʱ��, CDate(rsExec!Ҫ��ʱ��))
                                                    If rsSendDays!�ֽ�ʱ�� <> "" Then
                                                        rsSendDays!�״�ʱ�� = "To_Date('" & Split(rsSendDays!�ֽ�ʱ��, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                                                        rsSendDays!ĩ��ʱ�� = "To_Date('" & Split(rsSendDays!�ֽ�ʱ��, ",")(UBound(Split(rsSendDays!�ֽ�ʱ��, ","))) & "','YYYY-MM-DD HH24:MI:SS')"
                                                    End If
                                                    rsSendDays!NO = strNO
                                                    rsExec!���ͺ� = lng���ͺ�Tmp
                                                Else
                                                    rsExec!���ͺ� = lng���ͺ�Tmp
                                                End If
                                            Else
                                                '�滻���ͺ�
                                                rsDayNo.Filter = "����ʱ��='" & Format(rsExec!Ҫ��ʱ�� & "", "yyyy-MM-dd") & "'"
                                                If Not rsDayNo.EOF Then lng���ͺ�Tmp = rsDayNo!���ͺ�
                                                rsDays!���� = rsDays!���� + rsExec!����
                                                rsDays!ʵ�ս�� = Format(rsDays!ʵ�ս�� + rsExec!���� / dbl���� * strʵ��, gstrDec)
                                                rsDays!Ӧ�ս�� = Format(rsDays!Ӧ�ս�� + rsExec!���� / dbl���� * curӦ��, gstrDec)
                                                rsExec!���ͺ� = lng���ͺ�Tmp
                                            End If
                                            rsDays.Update
                                            rsSendDays.Update
                                            rsExec.Update
                                        End If
                                        rsDayNo.Filter = 0
                                        rsExec.MoveNext
                                    Loop
                                    rsExec.Filter = 0
                                    rsDays.Filter = 0
                                    '���� ���ͺ�,rsSendDays,rsExec����֤����Խ���ͺ�ҲԽ��
                                    rsSendDays.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                                    If rsSendDays.RecordCount > 1 Then
                                        Set rsTmp = zlDatabase.CopyNewRec(rsSendDays)
                                        rsTmp.Sort = "���ͺ�"
                                        rsSendDays.Sort = "����ʱ��"
                                        rsExec.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                                        rsExec.Sort = "Ҫ��ʱ��"
                                        Do While Not rsTmp.EOF
                                            Do While True
                                                If rsExec!���ͺ� = rsSendDays!���ͺ� Then
                                                    rsExec!���ͺ� = rsTmp!���ͺ�
                                                    rsExec.MoveNext
                                                Else
                                                    Exit Do
                                                End If
                                                If rsExec.EOF Then
                                                    Exit Do
                                                End If
                                            Loop
                                            rsSendDays!���ͺ� = rsTmp!���ͺ�
                                            rsSendDays.MoveNext
                                            rsTmp.MoveNext
                                        Loop
                                    End If
                                    rsExec.Sort = ""
                                    rsExec.Filter = 0
                                End If
                                rsDays.Filter = 0
'                               '���С�ڵ���һ�죬����һ�ŵ���
                                If rsDays.RecordCount = 1 Then
                                    rsDays!���� = dbl����
                                    rsDays!ʵ�ս�� = Format(strʵ��, gstrDec)
                                    rsDays!Ӧ�ս�� = Format(curӦ��, gstrDec)
                                    rsDays!����ʱ�� = str����ʱ��
                                    rsDays.Update
                                ElseIf rsDays.RecordCount = 0 Then
                                    rsDays.AddNew
                                    rsDays!���� = dbl����
                                    rsDays!ʵ�ս�� = Format(strʵ��, gstrDec)
                                    rsDays!Ӧ�ս�� = Format(curӦ��, gstrDec)
                                    rsDays!����ʱ�� = str����ʱ��
                                    rsDays!NO = strNO
                                    rsDays.Update
                                End If
                                
                                '��Ϊ���ڲ��Ƽ۵�ҽ������������,���Դ���ļƼ����Զ�Ϊ(0-�����Ƽ�)
                                Do While Not rsDays.EOF
                                    '�ռ�ҽ���ϴ����ݺ�:mrsBill�еĲ�һ�������˷���
                                    If int���� = 0 Then
                                        rsUpload.Filter = "NO='" & rsDays!NO & "'"
                                        If rsUpload.EOF Then
                                            rsUpload.AddNew
                                            rsUpload!ҽ��ID = lng��ID
                                            rsUpload!NO = rsDays!NO
                                            rsUpload.Update
                                        End If
                                    End If
                                    If rsSQL!Sql <> "" Then rsSQL.AddNew
                                    rsSQL!���� = 3
                                    rsSQL!ҽ��ID = lng��ID
                                    rsSQL!��ĿID = rsPrice!ID
                                    rsSQL!��� = i
                                    rsSQL!NO = rsDays!NO
                                    
                                    If lng�������� = 1 Then
                                        rsSQL!Sql = "zl_������ʼ�¼_INSERT(" & _
                                            "'" & rsDays!NO & "'," & lng������� & "," & lng����ID & "," & _
                                            "'" & rsPati!����� & "','" & .TextMatrix(i, COL_����) & "'," & _
                                            "'" & .TextMatrix(i, col_�Ա�) & "','" & .TextMatrix(i, COL_����) & "'," & "'" & .TextMatrix(i, COL_�ѱ�) & "',0," & Val(.Cell(flexcpData, i, COL_Ӥ��)) & "," & _
                                            ZVal(.TextMatrix(i, COL_���˿���ID)) & "," & ZVal(.TextMatrix(i, COL_��������ID)) & "," & _
                                            "'" & .TextMatrix(i, COL_����ҽ��) & "'," & IIF(rsPrice!���� = 1, ZVal(int�����), "NULL") & "," & _
                                            rsPrice!ID & ",'" & rsPrice!��� & "','" & rsPrice!���㵥λ & "'," & _
                                            int���� & "," & rsDays!���� & "," & IIF(bln��������, 1, 0) & "," & ZVal(lngִ�п���ID) & "," & _
                                            IIF(lng���ø��� = lng�������, "NULL", lng���ø���) & "," & rsPrice!������ĿID & "," & _
                                            "'" & rsPrice!�վݷ�Ŀ & "'," & dbl���� & "," & rsDays!Ӧ�ս�� & "," & rsDays!ʵ�ս�� & "," & _
                                            rsDays!����ʱ�� & "," & str�Ǽ�ʱ�� & "," & _
                                            "'ҽ������'," & int���� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                                            "Null,'" & .TextMatrix(i, col_ҽ������) & "'," & Val(.TextMatrix(i, COL_ID)) & ",'" & .TextMatrix(i, COL_Ƶ��) & "'," & _
                                            ZVal(.TextMatrix(i, COL_����)) & ",'" & .TextMatrix(i, COL_�÷�) & "'," & .Cell(flexcpData, i, COL_ҽ����Ч) & "," & _
                                            IIF(intҩƷ���� <> 0, intҩƷ����, Val(.TextMatrix(i, COL_�Ƽ�����))) & ",1,Null,0," & ZVal(Val(.TextMatrix(i, COL_��鷽��))) & "," & ZVal(lng��ҳID) & "," & Val(.TextMatrix(i, COL_���˲���ID)) & ")"
                                    Else
                                        rsSQL!Sql = "ZL_סԺ���ʼ�¼_Insert(" & _
                                            "'" & rsDays!NO & "'," & lng������� & "," & lng����ID & "," & ZVal(lng��ҳID) & "," & _
                                            IIF(.TextMatrix(i, COL_סԺ��) = "", "NULL", "'" & .TextMatrix(i, COL_סԺ��) & "'") & ",'" & .TextMatrix(i, COL_����) & "'," & _
                                            "'" & .TextMatrix(i, col_�Ա�) & "','" & .TextMatrix(i, COL_����) & "'," & _
                                            "'" & .TextMatrix(i, COL_����) & "','" & .TextMatrix(i, COL_�ѱ�) & "'," & _
                                            Val(.TextMatrix(i, COL_���˲���ID)) & "," & Val(.TextMatrix(i, COL_���˿���ID)) & ",0," & Val(.Cell(flexcpData, i, COL_Ӥ��)) & "," & _
                                            ZVal(.TextMatrix(i, COL_��������ID)) & ",'" & .TextMatrix(i, COL_����ҽ��) & "'," & _
                                            IIF(rsPrice!���� = 1, ZVal(int�����), "NULL") & "," & rsPrice!ID & "," & _
                                            "'" & rsPrice!��� & "','" & NVL(rsPrice!���㵥λ) & "'," & _
                                            IIF(bln������Ŀ��, 1, 0) & "," & ZVal(lng���մ���ID) & ",'" & str���ձ��� & "'," & _
                                            int���� & "," & rsDays!���� & "," & IIF(bln��������, 1, 0) & "," & ZVal(lngִ�п���ID) & "," & _
                                            IIF(lng���ø��� = lng�������, "NULL", lng���ø���) & "," & rsPrice!������ĿID & "," & _
                                            "'" & NVL(rsPrice!�վݷ�Ŀ) & "'," & dbl���� & "," & rsDays!Ӧ�ս�� & "," & rsDays!ʵ�ս�� & "," & _
                                            curͳ���� & "," & rsDays!����ʱ�� & "," & str�Ǽ�ʱ�� & "," & _
                                            "'ҽ������'," & int���� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "',0," & _
                                            IIF(rsPrice!��� = "4", mlng�������ID, mlngҩƷ���ID) & "," & _
                                            "NULL,'" & .TextMatrix(i, col_ҽ������) & "',NULL," & Val(.TextMatrix(i, COL_ID)) & "," & _
                                            "'" & .TextMatrix(i, COL_Ƶ��) & "'," & ZVal(.TextMatrix(i, COL_����)) & "," & _
                                            "'" & .TextMatrix(i, COL_�÷�) & "'," & .Cell(flexcpData, i, COL_ҽ����Ч) & "," & _
                                            IIF(intҩƷ���� <> 0, intҩƷ����, Val(.TextMatrix(i, COL_�Ƽ�����))) & "," & _
                                            "Null,'" & str�������� & "',Null," & strTmp & ",NULL,-1,0," & ZVal(Val(.TextMatrix(i, COL_��鷽��))) & ")"
                                    End If
                                    rsSQL.Update
                                    
                                
                                    '��¼�Զ����ϵ�SQL
                                    If (gbytסԺ�Զ����� = 1 Or gbytסԺ�Զ����� = 2 And lngִ�п���ID = Val(.TextMatrix(i, COL_��������ID))) And int���� = 0 And lngִ�п���ID <> 0 And rsPrice!��� = "4" And NVL(rsPrice!��������, 0) = 1 Then
                                        If InStr(str�Զ����� & ";", ";" & rsDays!NO & "," & lngִ�п���ID & ";") = 0 Then
                                            rsSQL.AddNew
                                            rsSQL!���� = 6
                                            rsSQL!ҽ��ID = lng��ID
                                            rsSQL!��ĿID = 0
                                            rsSQL!��� = i
                                            rsSQL!NO = rsDays!NO
                                            rsSQL!Sql = "zl_�����շ���¼_��������(" & lngִ�п���ID & ",25,'" & rsDays!NO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1,Sysdate)"
                                            rsSQL.Update
                                            str�Զ����� = str�Զ����� & ";" & rsDays!NO & "," & lngִ�п���ID
                                        End If
                                    End If
                                    rsDays.MoveNext
                                Loop
                                
                                'ҽ���ܿ�ʵʱ��⣺���ɷ�����Ŀ��¼��,���շ�ϸĿ����
                                If Not IsNull(rsPati!����) And blnʵʱ��� Then
                                    rsItems.Filter = "�շ�ϸĿID=" & rsPrice!ID
                                    If rsItems.EOF Then
                                        '�����շ���Ŀ��Ӧ��ԭʼ��Ϣ
                                        rsItems.AddNew
                                        rsItems!����ID = rsPati!����ID
                                        rsItems!��ҳID = rsPati!��ҳID
                                        rsItems!ҽ��ID = Val(.TextMatrix(i, COL_ID))
                                        rsItems!�շ���� = rsPrice!���
                                        rsItems!�շ�ϸĿID = rsPrice!ID
                                        rsItems!������ = .TextMatrix(i, COL_����ҽ��)
                                        rsItems!�������� = CStr(sys.RowValue("���ű�", Val(.TextMatrix(i, COL_��������ID)), "����"))
                                        
                                        rsItems!���� = int���� * dbl����
                                        rsItems!���� = dbl����
                                    Else
                                        '����һ��ҽ��(������Ŀ)���շѶ��ղ������ظ����շ�ϸĿ
                                        '������ͬһ�շ���Ŀ�Ĳ�ͬ������Ŀ��¼��ͬ
                                        If rsPrice!�������� & "_" & rsPrice!ID <> str�շ���Ŀ Then
                                            rsItems!���� = NVL(rsItems!����, 0) + int���� * dbl����
                                        End If
                                        '���ۣ�ͬһ�շ���Ŀ�Ĳ�ͬ������Ŀ�ۼ�
                                        If Val(.TextMatrix(i, COL_ID)) = rsItems!ҽ��ID Then
                                            rsItems!���� = NVL(rsItems!����, 0) + dbl����
                                        End If
                                    End If
                                    rsItems!ʵ�ս�� = NVL(rsItems!ʵ�ս��, 0) + curʵ��
                                    rsItems.Update
                                End If
                                    
                                str�շ���Ŀ = rsPrice!�������� & "_" & rsPrice!ID
                                rsPrice.MoveNext
                            Loop
                        End If
                        
                        '��ҽ�������л����ۿ۴���
                        If gbln��������ۿ� And strHaveSub <> "" Then
                            rsSeek.Filter = 0
                            Do While Not rsSeek.EOF
                                rsSQL.Bookmark = rsSeek!�����ǩ
                                curʵ�� = Format(ActualMoney(.TextMatrix(i, COL_�ѱ�), rsSeek!������ID, rsSeek!�ϼ�), gstrDec)
                                curʵ�� = curʵ�� - rsSeek!�ϼ� '���۲��
                                
                                'ҽ���ܿ�ʵʱ��⣺������Ŀ����滻
                                If Not IsNull(rsPati!����) And blnʵʱ��� Then
                                    rsItems.Filter = "�շ�ϸĿID=" & lng����ĿID
                                    If Not rsItems.EOF Then
                                        rsItems!ʵ�ս�� = NVL(rsItems!ʵ�ս��, 0) + curʵ��
                                        rsItems.Update
                                    End If
                                End If
                                
                                '����SQL�����滻
                                curʵ�� = Getʵ�ս��(rsSQL!Sql) + curʵ��
                                rsSQL!Sql = Setʵ�ս��(rsSQL!Sql, curʵ��)
                                rsSQL.Update
                            
                                rsSeek.MoveNext
                            Loop
                        End If
                                                
                        
                        '����ҽ�����ͼ�¼
                        '-----------------------------------------------------------------------------------------
                        If Val(.TextMatrix(i, COL_ִ������ID)) <> 0 Then  '����������(��ҩ;�����䷽�巨���÷����ɼ���������Ϊ)
                            '�����˳�Ժ,תԺ,����ҽ��
                            If .TextMatrix(i, COL_�������) = "Z" _
                                And InStr(",5,6,11,", Val(.TextMatrix(i, COL_��������))) > 0 Then
                                mblnRefresh = True
                            End If
                            
                            'һ��Ҫ��������NO
                            Call GetCurBillSet(strNOKey, strNO, -1, lng�������)
                                                                    
                            '�Ƿ�һ��ҽ���ĵ�һҽ����
                            blnFirst = False
                            If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                                If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                                    blnFirst = True 'ҩ�Ʒ���ʱ,ֻ�е�һҩƷ�в�Ϊ��һҽ����
                                End If
                            ElseIf .TextMatrix(i, COL_�������) = "C" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                                    blnFirst = True '��������еĵ�һ������
                                End If
                            ElseIf Val(.TextMatrix(i, COL_���ID)) = 0 Then '�ſ�����ҩ�巨����Ѫ;��
                                If Not (.TextMatrix(i, COL_�������) = "E" _
                                    And Val(.TextMatrix(i, COL_ID)) = Val(.TextMatrix(i - 1, COL_���ID))) Then '�ſ���ҩ;������ҩ�÷����ɼ�����
                                    blnFirst = True
                                End If
                            End If
                            
                            '����ִ�е��Զ�ִ�У�����ҽ��������
                            intִ��״̬ = 0
                            If Val(Mid(mstrAutoExe, IIF(.TextMatrix(i, COL_ҽ����Ч) = "����", 1, 0) + 1, 1)) <> 0 And Not (.TextMatrix(i, COL_�������) = "Z" And Val(.TextMatrix(i, COL_��������)) <> 0) _
                                And (Val(.TextMatrix(i, COL_ִ�п���ID)) = Val(.TextMatrix(i, COL_���˿���ID)) Or Val(.TextMatrix(i, COL_ִ�п���ID)) = Val(.TextMatrix(i, COL_���˲���ID))) Then
                                If CanAutoExeItem(Val(.TextMatrix(i, COL_ִ�п���ID)), .TextMatrix(i, COL_�������), .TextMatrix(i, COL_��������), Val(.TextMatrix(i, COL_ִ�з���))) Then
                                    intִ��״̬ = 1
                                End If
                                If intִ��״̬ = 1 Then
                                    'Ѫ��������⴦��
                                    If gblnѪ��ϵͳ Then
                                        strTmp = .TextMatrix(i, COL_�������) & .TextMatrix(i, COL_��������) & .TextMatrix(i, COL_ִ�з���)
                                        If strTmp = "E81" Or strTmp = "E90" Then
                                            If InitObjBlood() Then
                                                intִ��״̬ = 0
                                            End If
                                        End If
                                    End If
                                End If
                            End If
    
                            '��ĩʱ��(�����á�str�ֽ�ʱ�䡱�жϣ���Ϊһ����������¼�����״�ʱ��)
                            If .TextMatrix(i, COL_�ֽ�ʱ��) <> "" Then
                                str�״�ʱ�� = "To_Date('" & Split(str�ֽ�ʱ��, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                                strĩ��ʱ�� = "To_Date('" & Split(str�ֽ�ʱ��, ",")(UBound(Split(str�ֽ�ʱ��, ","))) & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                '�޷��ֽ��Ϊ"һ����"��������Ϊ��ʼִ��ʱ�䣨74366��
                                str�״�ʱ�� = "To_Date('" & .TextMatrix(i, COL_��ʼʱ��) & "','YYYY-MM-DD HH24:MI:SS')"
                                strĩ��ʱ�� = "To_Date('" & .TextMatrix(i, COL_��ʼʱ��) & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            
                           
                            '��������:ҩƷΪ������λ������,����Ϊ����
                            If .TextMatrix(i, COL_�������) = "7" Then
                                dbl�������� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����))
                            ElseIf InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                                dbl�������� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_סԺ��װ)) * Val(.TextMatrix(i, COL_����ϵ��))
                            Else
                                dbl�������� = Val(.TextMatrix(i, COL_����))
                            End If
                                       
                            '��ҩ��
                            str��ҩ�� = ""
                            If mbln��ҩ�� And InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                                If mstr��ҩ�� = "" Then mstr��ҩ�� = Get��ҩ��
                                str��ҩ�� = mstr��ҩ��
                            End If
                            
                            '��ͨ����ҳ�涼��������ҩ��¼��64615��
                            If Not gbln�������������� Then strCuvetteNumber = ""
                            
                            '������ҩ�����ַ��ͼ�¼
                            rsSendDays.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                            If rsSendDays.RecordCount = 1 Then
                                rsSendDays!�ֽ�ʱ�� = str�ֽ�ʱ��
                                rsSendDays!�������� = dbl��������
                                rsSendDays.Update
                            ElseIf rsSendDays.RecordCount = 0 Then
                                rsSendDays.AddNew
                                rsSendDays!ҽ��ID = Val(.TextMatrix(i, COL_ID))
                                rsSendDays!���ͺ� = lng���ͺ�
                                rsSendDays!NO = strNO
                                rsSendDays!�ֽ�ʱ�� = str�ֽ�ʱ��
                                rsSendDays!�������� = dbl��������
                                rsSendDays.Update
                            Else
                                If rsSendDays.RecordCount > 0 Then
                                    Do While Not rsSendDays.EOF
                                        rsSendDays!�������� = dbl�������� * ((UBound(Split(Get����ֽ�ʱ��(rsSendDays!�ֽ�ʱ��, rsSendDays!����ʱ��), ",")) + 1) / (UBound(Split(str�ֽ�ʱ��, ",")) + 1))
                                        rsSendDays.MoveNext
                                    Loop
                                    rsSendDays.MoveFirst
                                End If
                            End If
                            
                            rsSendDays.Sort = "����ʱ��"
                            Do While Not rsSendDays.EOF
                                dbl�������� = Format(rsSendDays!��������, "0.00000")
                                rsSQL.AddNew
                                rsSQL!���� = 5
                                rsSQL!ҽ��ID = lng��ID
                                rsSQL!��ĿID = 0
                                rsSQL!��� = i
                                rsSQL!NO = rsSendDays!NO
                                
                                If InStr("," & strPlugIn���ͺ� & ",", "," & rsSendDays!���ͺ� & ",") = 0 Then
                                    strPlugIn���ͺ� = strPlugIn���ͺ� & "," & rsSendDays!���ͺ�
                                End If
                                
                                rsSQL!Sql = "ZL_����ҽ������_Insert(" & _
                                    Val(.TextMatrix(i, COL_ID)) & "," & rsSendDays!���ͺ� & ",2,'" & rsSendDays!NO & "'," & _
                                    lng������� & "," & ZVal(dbl��������) & "," & IIF(rsSendDays!�״�ʱ�� & "" = "", str�״�ʱ��, rsSendDays!�״�ʱ��) & "," & IIF(rsSendDays!ĩ��ʱ�� & "" = "", strĩ��ʱ��, rsSendDays!ĩ��ʱ��) & "," & strCurDate & "," & _
                                    intִ��״̬ & "," & ZVal(.TextMatrix(i, COL_ִ�п���ID)) & "," & int�Ʒ�״̬ & "," & _
                                    IIF(blnFirst, 1, 0) & ",'" & strCuvetteNumber & "','" & UserInfo.��� & "'," & _
                                    "'" & UserInfo.���� & "','" & str��ҩ�� & "'," & IIF(lng�������� = 1, 1, "Null") & ",'" & rsSendDays!�ֽ�ʱ�� & "'," & IIF(InStr(str����ҩ��, "," & Val(.TextMatrix(i, COL_ID)) & "'") > 0, str����ҩ��, "Null") & ")"
                                rsSQL.Update
                                str����ҩ�� = "''"
                                rsSendDays.MoveNext
                            Loop
                            If gblnѪ��ϵͳ And .TextMatrix(i, COL_�������) = "K" Then
                                rsSQL.AddNew
                                rsSQL!���� = 9
                                rsSQL!��ĿID = 0
                                rsSQL!��� = 0
                                rsSQL!Sql = "Zl_ѪҺ��Ѫ����_Insert(" & Val(.TextMatrix(i, COL_ID)) & ")"
                                rsSQL.Update
                            End If
                            
                        End If
                    
                        '������ҩ�䷽��
                        If .Cell(flexcpData, i, COL_�������) = 3 Then '��ҩ�÷�
                            int�䷽�� = int�䷽�� + 1
                        End If
                    End If  'Ҫ���ͺͼ��ʵ�
                End If  '��ǰѡ���
            End If
NextAdvice:
            '----------------------------------------
            Progress = (i - .FixedRows + 1) / (.Rows - .FixedRows) * 100
            lngOld��ID = lng��ID
        Next
        
        '�ύ���һ�����˵�����
        '-----------------------------------------------------------------------------------------
        If lng����ID <> 0 Then
            If strAudit <> "" Then
                MsgBox "����""" & rsPati!���� & """���·�����Ŀ��û�о�����������Ӧ��ҽ�����ܷ��ͣ�" & vbCrLf & strAudit, vbInformation, gstrSysName
                GoTo errH
            End If
            
            
            'ҽ��ִ�мƼ�
            If rsExec.RecordCount > 0 Then
                rsExec.MoveFirst
                Do While Not rsExec.EOF
                    rsSQL.AddNew
                    rsSQL!���� = 8
                    rsSQL!��ĿID = 0
                    rsSQL!��� = 0
                    rsSQL!ҽ��ID = lng��ID
                    rsSQL!Sql = "Zl_ҽ��ִ�мƼ�_Insert(" & rsExec!ҽ��ID & "," & rsExec!���ͺ� & ",To_date('" & _
                    rsExec!Ҫ��ʱ�� & "','yyyy-MM-dd HH24:mi:ss')," & ZVal(Val(rsExec!�շ�ϸĿID & "")) & "," & rsExec!���� & "," & rsExec!�������� & ")"
                    rsExec.MoveNext
                Loop
            End If
        
            If Not CompletePatiSend(rsPati, rsSQL, rsUpload, rsItems, cur�ϼ�, cur���ʺϼ�, str���, bln����, blnTran, lng���ͺ�) Then GoTo errH
            SendAdvice = lng���ͺ� 'ֻҪ�ύ�ɹ����ע
        End If
    End With
    mstrRollNotify = Mid(mstrRollNotify, 2)
    SendAdvice = lng���ͺ�
    
    strPlugIn���ͺ� = Mid(strPlugIn���ͺ�, 2)
    '������ҽӿ�
    If CreatePlugInOK(pסԺҽ������) Then
        On Error Resume Next
        Call gobjPlugIn.AdviceSendEnd(glngSys, pסԺҽ������, strPlugIn���ͺ�)
        Call zlPlugInErrH(err, "AdviceSendEnd")
        err.Clear: On Error GoTo 0
    End If
    Call Make��ִ����Ϣ(strCurDateTmp)
FuncEnd:
    'ɾ�������ѳɹ����͵���
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then
        gcnOracle.RollbackTrans
    End If
    If err.Number <> 0 Then '��ҽ���ϴ�ʧ���˳�û�д���
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0
End Function

Private Function Get����ֽ�ʱ��(ByVal str�ֽ�ʱ�� As String, ByVal dat���� As Date) As String
'���ܣ���ȡָ����ķֽ�ʱ��
    Dim arrTmp As Variant
    Dim i As Long
    
    arrTmp = Split(str�ֽ�ʱ��, ",")
    For i = 0 To UBound(arrTmp)
        If Format(CDate(arrTmp(i)), "YYYY-MM-DD") = Format(dat����, "YYYY-MM-DD") Then
            Get����ֽ�ʱ�� = Get����ֽ�ʱ�� & "," & arrTmp(i)
        End If
    Next
    Get����ֽ�ʱ�� = Mid(Get����ֽ�ʱ��, 2)
End Function

Private Sub ShowSendTotal()
'���ܣ����ݵ�ǰѡ��Ҫ���͵�ҽ�������㲢��ʾ���͵�ҽ���ϼ�
    Dim curTotal As Currency, i As Long
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If Not .RowHidden(i) And .Cell(flexcpData, i, COL_ѡ��) = 0 _
                And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                curTotal = curTotal + Val(.TextMatrix(i, COL_���))
            End If
        Next
    End With
    stbThis.Panels(3).Text = "���ͷ��ã�" & Format(curTotal, gstrDec)
    Call Form_Resize
End Sub

Private Sub SetDeptInput(ByVal lngRow As Long, ByVal lngCol As Long, rsInput As ADODB.Recordset)
'���ܣ�����ִ�п�������ĵ�ֵ
    Dim strҽ��IDs As String
    Dim strִ�п���IDs As String
    Dim i As Long
        
    With vsAdvice
        If lngCol = COL_ִ�п��� Then
            '������ʾ�е�ִ�п�����ʾ
            .TextMatrix(lngRow, COL_ִ�п���) = rsInput!����
            .Cell(flexcpData, lngRow, COL_ִ�п���) = .TextMatrix(lngRow, COL_ִ�п���)
            
            '��������Ŀ��ִ�п���(�ſ���ǰ��ʾ��Ϊ�ɼ���ʽ����)
            If Not (.TextMatrix(lngRow, COL_�������) = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 _
                And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID))) Then
                
                strҽ��IDs = strҽ��IDs & "," & Val(.TextMatrix(lngRow, COL_ID))
                strִ�п���IDs = strִ�п���IDs & "," & Val(.TextMatrix(lngRow, COL_ִ�п���ID))
                
                .TextMatrix(lngRow, COL_ִ�п���ID) = rsInput!ID
                .Cell(flexcpData, lngRow, COL_ִ�п���ID) = 1
                
            End If
            
            '����������ϵĸ�������
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If .TextMatrix(i, COL_�������) <> "G" _
                        And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then  '���������������ִ�п���
                        
                        strҽ��IDs = strҽ��IDs & "," & Val(.TextMatrix(i, COL_ID))
                        strִ�п���IDs = strִ�п���IDs & "," & Val(.TextMatrix(i, COL_ִ�п���ID))
                        
                        .TextMatrix(i, COL_ִ�п���) = rsInput!����
                        .Cell(flexcpData, i, COL_ִ�п���) = .TextMatrix(i, COL_ִ�п���)
                        .TextMatrix(i, COL_ִ�п���ID) = rsInput!ID
                        .Cell(flexcpData, i, COL_ִ�п���ID) = 1
                    End If
                Else
                    Exit For
                End If
            Next
            
            '������ϵ�����
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                    
                        strҽ��IDs = strҽ��IDs & "," & Val(.TextMatrix(i, COL_ID))
                        strִ�п���IDs = strִ�п���IDs & "," & Val(.TextMatrix(i, COL_ִ�п���ID))
                        
                        .TextMatrix(i, COL_ִ�п���) = rsInput!����
                        .Cell(flexcpData, i, COL_ִ�п���) = .TextMatrix(i, COL_ִ�п���)
                        .TextMatrix(i, COL_ִ�п���ID) = rsInput!ID
                        .Cell(flexcpData, i, COL_ִ�п���ID) = 1
                        
                    End If
                Else
                    Exit For
                End If
            Next
        ElseIf lngCol = COL_����ִ�� Then
            '������ʾ�еĸ���ִ�п�����ʾ
            .TextMatrix(lngRow, COL_����ִ��) = rsInput!����
            .Cell(flexcpData, lngRow, COL_����ִ��) = .TextMatrix(lngRow, COL_����ִ��)
            
            '���ĸ�����Ŀ�е�ִ�п���
            If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
                '��ҩ;��
                i = .FindRow(CStr(.TextMatrix(lngRow, COL_���ID)), lngRow + 1, COL_ID)
                
                strҽ��IDs = strҽ��IDs & "," & Val(.TextMatrix(i, COL_ID))
                strִ�п���IDs = strִ�п���IDs & "," & Val(.TextMatrix(i, COL_ִ�п���ID))
                
                .TextMatrix(i, COL_ִ�п���ID) = rsInput!ID
                .Cell(flexcpData, i, COL_ִ�п���ID) = 1
                
                For i = lngRow + 1 To .Rows - 1
                    If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                        .TextMatrix(i, COL_����ִ��) = rsInput!����
                        .Cell(flexcpData, i, COL_����ִ��) = .TextMatrix(lngRow, COL_����ִ��)
                    Else
                        Exit For
                    End If
                Next
            ElseIf .TextMatrix(lngRow, COL_�������) = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 _
                And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) _
                And InStr(",7,E,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                
                strҽ��IDs = strҽ��IDs & "," & Val(.TextMatrix(lngRow, COL_ID))
                strִ�п���IDs = strִ�п���IDs & "," & Val(.TextMatrix(lngRow, COL_ִ�п���ID))
                        
                '��ҩ�÷�
                .TextMatrix(lngRow, COL_ִ�п���ID) = rsInput!ID
                .Cell(flexcpData, lngRow, COL_ִ�п���ID) = 1
            ElseIf .TextMatrix(lngRow, COL_�������) = "F" Then
                '��������
                For i = lngRow + 1 To .Rows - 1
                    If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                        If .TextMatrix(i, COL_�������) = "G" Then
                            If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                                
                                strҽ��IDs = strҽ��IDs & "," & Val(.TextMatrix(i, COL_ID))
                                strִ�п���IDs = strִ�п���IDs & "," & Val(.TextMatrix(i, COL_ִ�п���ID))
                                
                                .TextMatrix(i, COL_ִ�п���ID) = rsInput!ID
                                .Cell(flexcpData, i, COL_ִ�п���ID) = 1
                                
                            End If
                            Exit For 'ֻ��һ������
                        End If
                    Else
                        Exit For
                    End If
                Next
            ElseIf .TextMatrix(lngRow, COL_�������) = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 _
                And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                '�ɼ���ʽ
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������ID))) = 0 Then
                    
                    strҽ��IDs = strҽ��IDs & "," & Val(.TextMatrix(lngRow, COL_ID))
                    strִ�п���IDs = strִ�п���IDs & "," & Val(.TextMatrix(lngRow, COL_ִ�п���ID))
                    
                    .TextMatrix(lngRow, COL_ִ�п���ID) = rsInput!ID
                    .Cell(flexcpData, lngRow, COL_ִ�п���ID) = 1
                End If
            End If
        End If
        If strҽ��IDs <> "" Then
            strҽ��IDs = Mid(strҽ��IDs, 2)
            strִ�п���IDs = Mid(strִ�п���IDs, 2)
            
            For i = 0 To UBound(Split(strҽ��IDs, ","))
                'ͬ�����·���ִ�п��ң�ֻ���º�ԭҽ��ִ�п�����ͬ�ķ���ִ�п��ң�
                mrsPrice.Filter = "ҽ��ID=" & Val(Split(strҽ��IDs, ",")(i))
                If Not mrsPrice.EOF Then mrsPrice.MoveFirst
                Do Until mrsPrice.EOF
                    If Val(mrsPrice!ִ�п���ID & "") = Val(Split(strִ�п���IDs, ",")(i)) And Val(Split(strִ�п���IDs, ",")(i)) <> 0 Then
                        mrsPrice!ִ�п���ID = Val(rsInput!ID & "")
                        mrsPrice.Update
                    End If
                    mrsPrice.MoveNext
                Loop
            Next
        End If
    End With
End Sub

Private Sub vsPrice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsPrice.EditSelStart = 0
    vsPrice.EditSelLength = zlcommfun.ActualLen(vsPrice.EditText)
End Sub

Private Sub vsPrice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim bln�Ǳ��� As Boolean
    
    If Not CellEditablePrice(Row, Col, bln�Ǳ���) Then
        '�Ǳ���ִ�еı����Ŀ�������۸�
        If bln�Ǳ��� Then
            MsgBox "��ҽ���Ǳ���ִ�У�������Ա����Ŀ���ۡ��üƼ���Ŀ��Ҫ�ֹ��Ƽۡ�", vbInformation, gstrSysName
        End If
        Cancel = True
    Else
        If Col = COLP_�Ƽ����� Or Col = COLP_���� Or Col = COLP_ִ�п��� Then
            '������ȷ���շ���Ŀ
            If vsPrice.TextMatrix(Row, COLP_�շ���Ŀ) = "" Then Cancel = True
        End If
        If Col = COLP_���� Then
            '������ǰ������ȷ���Ƽ�ҽ��,�Ծ����Ƿ��������(����ִ��)
            If vsPrice.TextMatrix(Row, COLP_�Ƽ�ҽ��) = "" Then Cancel = True
        End If
    End If
    
    If Col = COLP_�Ƽ����� Or Col = COLP_���� Then
        vsPrice.EditMaxLength = 10
    Else
        vsPrice.EditMaxLength = 0
    End If
End Sub

Private Sub GetPatiRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
'���ܣ���ȡ��ID��ͬ��һ��ҽ���кŷ�Χ(ע�⿼��һ����ҩ�еĿ���)
    Dim lng����ID As Long, lng��ҳID As Long, lngӤ�� As Long, i As Long
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lng����ID = Val(.TextMatrix(lngRow, COL_����ID))
        lng��ҳID = Val(.TextMatrix(lngRow, COL_��ҳID))
        lngӤ�� = Val(.TextMatrix(lngRow, COL_Ӥ��))
        
        For i = lngRow - 1 To .FixedRows Step -1
            If lng����ID = Val(.TextMatrix(lngRow, COL_����ID)) And lng��ҳID = Val(.TextMatrix(lngRow, COL_��ҳID)) And lngӤ�� = Val(.Cell(flexcpData, lngRow, COL_Ӥ��)) Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            If Not (lng����ID = Val(.TextMatrix(lngRow, COL_����ID)) And lng��ҳID = Val(.TextMatrix(lngRow, COL_��ҳID)) And lngӤ�� = Val(.Cell(flexcpData, lngRow, COL_Ӥ��))) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
    End With
End Sub



Private Sub Del��������()
'���ܣ�ҽ������ʧ�ܣ�������˺󣬵��ü�������ɾ���ӿ�
    Dim i As Long, strҽ��IDs As String, strErr As String
        
    '�ռ��ɼ�����
    With vsAdvice
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_��������)) = 6 And .TextMatrix(i, COL_�������) = "E" Then
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    strҽ��IDs = strҽ��IDs & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    Call InitObjLis(pסԺ��ʿվ)
    If strҽ��IDs <> "" Then
        strҽ��IDs = Mid(strҽ��IDs, 2)
        If Not gobjLIS Is Nothing Then
            If gobjLIS.DelLisApplicationForm(strҽ��IDs, strErr) = False Then
                MsgBox "ɾ����������ʧ�ܣ�" & strErr, vbInformation, gstrSysName
            End If
        End If
    End If
End Sub

Private Function Get����ҽ����Ϣ(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ���ȡ����ҽ����Ϣ�����ݸ�����ӿڳ���
    Dim i As Long, strInfo As String
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_��������)) = 6 And .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_����ID)) = lng����ID And Val(.TextMatrix(i, COL_��ҳID)) = lng��ҳID Then
                '����ҽ��ID1,�ɼ�ҽ��ID1,ִ�п���ID1,�걾1;.....
                'LIS�ӿڲ����ļ��飬һ���ɼ���ʽֻ��һ������ҽ����û��һ���ɼ��������
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    strInfo = strInfo & ";" & .TextMatrix(i - 1, COL_ID) & "," & .TextMatrix(i, COL_ID) & "," & .TextMatrix(i - 1, COL_ִ�п���ID) & "," & .TextMatrix(i - 1, COL_�걾��λ)
                End If
            End If
        Next
    End With
    Get����ҽ����Ϣ = Mid(strInfo, 2)
End Function

Private Function Get��Ѫҽ����Ϣ(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ���ȡ����ҽ����Ϣ�����ݸ�����ӿڳ���
    Dim i As Long, strInfo As String
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, COL_�������) = "K" And Val(.TextMatrix(i, COL_����ID)) = lng����ID And Val(.TextMatrix(i, COL_��ҳID)) = lng��ҳID Then
                '����ҽ��ID1,�ɼ�ҽ��ID1,ִ�п���ID1,�걾1;.....
                'LIS�ӿڲ����ļ��飬һ���ɼ���ʽֻ��һ������ҽ����û��һ���ɼ��������
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    strInfo = strInfo & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    Get��Ѫҽ����Ϣ = Mid(strInfo, 2)
End Function

Private Function CheckAdviceUpdate(ByVal lngҽ��ID As Long, ByVal str�¿�����ʱ�� As String) As Boolean
'���ܣ����������У��ģʽ�������Ƿ��в����޸ġ�
    Dim rsTmp As Recordset, strSQL As String
    
    strSQL = "Select ����ʱ�� From ����ҽ��״̬ Where ҽ��ID=[1] And ��������=1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    If rsTmp.RecordCount > 0 Then
        If Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss") <> str�¿�����ʱ�� Then CheckAdviceUpdate = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitExecRecordset(rsExec As Recordset)
'���ܣ���ʼ��ҽ���Ƽۼ�¼��
    Set rsExec = New ADODB.Recordset
    
    rsExec.Fields.Append "ҽ��ID", adBigInt
    rsExec.Fields.Append "���ͺ�", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "Ҫ��ʱ��", adDate, , adFldIsNullable
    rsExec.Fields.Append "����", adDouble, , adFldIsNullable
    rsExec.Fields.Append "��������", adInteger, , adFldIsNullable
    
    rsExec.CursorLocation = adUseClient
    rsExec.LockType = adLockOptimistic
    rsExec.CursorType = adOpenStatic
    rsExec.Open
End Sub

Private Function zlPluginAdviceBeforeSend() As Boolean
'���ܣ�ҽ������ǰ������Һ�
    Dim i As Long, j As Long
    Dim strAdviceIDs As String, strMsg  As String
    Dim rsDataPlugIn As ADODB.Recordset
    Dim lng���� As Long
    Dim str�ֽ�ʱ�� As String, strTmp As String
    
    zlPluginAdviceBeforeSend = True
    
    '������ҽӿڣ�ҽ������ǰ�ļ��
    If CreatePlugInOK(pסԺҽ������) Then
        Call InitPlugInRs(rsDataPlugIn)
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    If .TextMatrix(i, COL_�ֽ�ʱ��) <> "" Then
                        str�ֽ�ʱ�� = .TextMatrix(i, COL_�ֽ�ʱ��)
                    Else
                        str�ֽ�ʱ�� = .Cell(flexcpData, i, COL_�ֽ�ʱ��)    '��ʼִ��ʱ��
                    End If
                    rsDataPlugIn.AddNew
                    rsDataPlugIn!����ID = Val(.TextMatrix(i, COL_����ID))
                    rsDataPlugIn!����ID = Val(.TextMatrix(i, COL_��ҳID))
                    rsDataPlugIn!ҽ��ID = Val(.TextMatrix(i, COL_ID))
                    rsDataPlugIn!���ID = Val(.TextMatrix(i, COL_���ID))
                    rsDataPlugIn!�շ�ϸĿID = Val(.TextMatrix(i, COL_�շ�ϸĿID))
                    rsDataPlugIn!�ֽ�ʱ�� = str�ֽ�ʱ��
                    rsDataPlugIn!���� = Val(.TextMatrix(i, COL_����))
                    rsDataPlugIn!���� = Val(.TextMatrix(i, COL_����))
                    rsDataPlugIn!������λ = .TextMatrix(i, COL_������λ)
                    rsDataPlugIn!���� = Val(.TextMatrix(i, COL_����))
                    rsDataPlugIn!������λ = .TextMatrix(i, COL_������λ)
                    rsDataPlugIn!���� = 1
                    rsDataPlugIn.Update
                End If
            Next
            If rsDataPlugIn.RecordCount > 0 Then rsDataPlugIn.MoveFirst
            strAdviceIDs = "": strMsg = ""
            On Error Resume Next
            Call gobjPlugIn.AdviceBeforeSend(mstrEnd, rsDataPlugIn, strAdviceIDs, strMsg)
            Call zlPlugInErrH(err, "AdviceBeforeSend")
            err.Clear
            On Error GoTo 0
             
            If strAdviceIDs <> "" Then
                strTmp = ""
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        If InStr("," & strAdviceIDs & ",", "," & Val(.TextMatrix(i, COL_ID)) & ",") > 0 Then
                            If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                                j = Val(.TextMatrix(i, COL_ID))
                            Else
                                j = Val(.TextMatrix(i, COL_���ID))
                            End If
                            
                            If InStr("," & strTmp & ",", "," & j & ",") = 0 Then
                                strTmp = strTmp & "," & j
                            End If
                        End If
                    End If
                Next
                strAdviceIDs = Mid(strTmp, 2)
                lng���� = 0
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            j = Val(.TextMatrix(i, COL_ID))
                        Else
                            j = Val(.TextMatrix(i, COL_���ID))
                        End If
                        lng���� = lng���� + 1
                        If InStr("," & strAdviceIDs & ",", "," & j & ",") > 0 Then
                            .Cell(flexcpData, i, COL_ѡ��) = 1
                            Set .Cell(flexcpPicture, i, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                            lng���� = lng���� - 1
                        End If
                    End If
                Next
                
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                If lng���� = 0 Then
                    MsgBox "��ǰû�п��Է��͵�ҽ����", vbInformation, gstrSysName
                    zlPluginAdviceBeforeSend = False
                End If
            End If
        End With
    End If
End Function

Private Function GetAdviceRis(ByRef rsData As ADODB.Recordset, ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ���ȡ���͵�RIS��ҽ����Ϣ
    Dim i As Long
    
    On Error GoTo errH
    
    Set rsData = New ADODB.Recordset
    
    rsData.Fields.Append "ҽ��ID", adBigInt
    rsData.Fields.Append "��������ID", adBigInt
    rsData.Fields.Append "ִ�п���ID", adBigInt
    rsData.Fields.Append "������ĿID", adBigInt
    rsData.Fields.Append "������Դ", adInteger '1-����;2-סԺ;
    rsData.Fields.Append "���", adVarChar, 10
    rsData.CursorLocation = adUseClient
    rsData.LockType = adLockOptimistic
    rsData.CursorType = adOpenStatic
    rsData.Open
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If Val(.TextMatrix(i, COL_����ID)) = lng����ID And Val(.TextMatrix(i, COL_��ҳID)) = lng��ҳID Then
                    If InStr(",D,F,", .TextMatrix(i, COL_�������)) > 0 Or _
                        .TextMatrix(i, COL_ҽ����Ч) = "����" And InStr(",0,5,", Val(.TextMatrix(i, COL_��������))) > 0 And .TextMatrix(i, COL_�������) = "E" Then
                        If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            rsData.AddNew
                            rsData!ҽ��ID = Val(.TextMatrix(i, COL_ID))
                            rsData!��������id = Val(.TextMatrix(i, COL_��������ID))
                            rsData!ִ�п���ID = Val(.TextMatrix(i, COL_ִ�п���ID))
                            rsData!������ĿID = Val(.TextMatrix(i, COL_������ĿID))
                            rsData!������Դ = 2
                            rsData!��� = .TextMatrix(i, COL_�������)
                            rsData.Update
                        End If
                    End If
                End If
            End If
        Next
    End With
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        GetAdviceRis = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckDrug����(ByVal lngRow As Long, ByRef bln��ʾ As Boolean) As Boolean
'���ܣ����͹����ж�����ҩƷ���м���ֹ
    Dim strTmp As String
    Dim blnTmp As Boolean
    Dim vMsg As VbMsgBoxResult
    
    With vsAdvice
        If 0 <> Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) And 0 <> Val(.TextMatrix(lngRow, COL_ִ�п���ID)) And .Cell(flexcpData, lngRow, COL_ѡ��) <> 1 Then
            If InitObjPublicDrug Then
                blnTmp = gobjPublicDrug.zlCheckPriceAdjustBySell(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), Val(.TextMatrix(lngRow, COL_ִ�п���ID)), False)
                If Not blnTmp Then
                    strTmp = "��(" & .TextMatrix(lngRow, COL_ִ�п���) & ")��ҩƷ""" & .TextMatrix(lngRow, col_ҽ������) & """" & vbCrLf & vbCrLf & _
                        "���������۹����Ҫ�󣺳ɱ��ۺ��ۼ۲�һ�£��������۳��⡣" & vbCrLf & vbCrLf & _
                        "����ϵҩ����ҩ���ƽ��е��۴���"
                    
                    If bln��ʾ Then
                        .Redraw = flexRDDirect:
                        Call .ShowCell(lngRow, COL_ѡ��)
                        Screen.MousePointer = 0
                        vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, True)
                        If vMsg = vbIgnore Then bln��ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        Screen.MousePointer = 11
                        .Refresh: .Redraw = flexRDNone
                    Else
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    End If
                    CheckDrug���� = True
                End If
            End If
        End If
    End With
End Function

Private Function CheckRISScheduling() As Boolean
'���ܣ������Ŀ�Ƿ��Ǳ���ԤԼ
    Dim i As Long
    Dim blnDo As Boolean
    Dim lngҽ��ID As Long
    Dim lng������ĿID As Long
    Dim lngRst As Long
    Dim strMsg As String
    
    CheckRISScheduling = True
    
    If HaveRIS Then
        If gbln����Ӱ����ϢϵͳԤԼ Then
            blnDo = True
        End If
    End If
    
    If Not blnDo Then Exit Function
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If Val(.TextMatrix(i, COL_������־)) <> 1 Then
                    If InStr(",D,F,", .TextMatrix(i, COL_�������)) > 0 Or InStr(",0,5,", Val(.TextMatrix(i, COL_��������))) > 0 And .TextMatrix(i, COL_�������) = "E" Then
                        If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            lngҽ��ID = Val(.TextMatrix(i, COL_ID))
                            lng������ĿID = Val(.TextMatrix(i, COL_������ĿID))
                            lngRst = -1
                            lngRst = gobjRis.HISScheduling(2, lngҽ��ID, lng������ĿID, False)
                            If lngRst <> 0 Then
                            '�ӿڷ���ʧ�ܸ�����ʾ
                                .Cell(flexcpData, i, COL_ѡ��) = 1 '��ǰ��ֹѡ��
                                Set .Cell(flexcpPicture, i, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("F").Picture
                                Call RowSelectSame(i, COL_ѡ��)
                                strMsg = IIF("" = strMsg, "", strMsg & "��") & .TextMatrix(i, col_ҽ������)
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End With
    If strMsg <> "" Then
        MsgBox "����������RISϵͳԤԼ���̣�" & vbCrLf & "��" & strMsg & "��" & _
                vbCrLf & "ҽ��û��ԤԼ��ԤԼ�ɹ�����ܷ��͡�", vbInformation, gstrSysName
        CheckRISScheduling = False
    End If
End Function

Private Function Set������ҩ() As Boolean
'���ܣ�����ҩƷҽ���е�������ҩ˵��
    Dim i As Long
    Dim strMsg As String
    Dim str������ҩ As String
    Dim strSQL As String
    Dim strҽ��IDs As String
    
    On Error GoTo errH
    If mstrAdDrugIDs = "" Then
        Set������ҩ = True
        Exit Function
    End If
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If InStr("," & mstrAdDrugIDs & ",", "," & .TextMatrix(i, COL_ID) & ",") > 0 Then
                    strMsg = strMsg & "," & .TextMatrix(i, col_ҽ������)
                    strҽ��IDs = strҽ��IDs & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    If strMsg = "" Then
        Set������ҩ = True
        Exit Function
    End If
    Call frmMsgDruExcess.ShowMe(Me, 1, Mid(strMsg, 2), str������ҩ)
    If str������ҩ = "*NULL*" Then
        Exit Function
    End If
    strSQL = "Zl_����ҽ����¼_������ҩ('" & Mid(strҽ��IDs, 2) & "','" & str������ҩ & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Set������ҩ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
