VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeBat 
   Caption         =   "��������"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15105
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChargeBat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15105
   StartUpPosition =   1  '����������
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   33
      Top             =   10590
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmChargeBat.frx":0442
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18150
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   88
            Key             =   "�������"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   2
            Key             =   "MedicareType"
            Object.ToolTipText     =   "���մ���"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   951
            MinWidth        =   951
            Picture         =   "frmChargeBat.frx":0CD6
            Key             =   "Drugstore"
            Object.Tag             =   "Drugstore"
            Object.ToolTipText     =   "ҩ������"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmChargeBat.frx":0FF0
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmChargeBat.frx":162A
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picBillList 
      BorderStyle     =   0  'None
      Height          =   9180
      Left            =   3045
      ScaleHeight     =   9180
      ScaleWidth      =   11985
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   915
      Width           =   11985
      Begin VB.CheckBox chk�Ӱ� 
         Caption         =   "�Ӱ�(&A)"
         Height          =   270
         Left            =   7125
         TabIndex        =   7
         Top             =   75
         Width           =   1170
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "�������"
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   8475
         TabIndex        =   8
         Top             =   75
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.ComboBox cbo������ 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4455
         TabIndex        =   6
         Top             =   30
         Width           =   2205
      End
      Begin VB.PictureBox picBillBottom 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2805
         Left            =   75
         ScaleHeight     =   2805
         ScaleWidth      =   11835
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   6390
         Width           =   11835
         Begin VSFlex8Ctl.VSFlexGrid vsMoney 
            Height          =   1665
            Left            =   15
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1110
            Width           =   3420
            _cx             =   6032
            _cy             =   2937
            Appearance      =   2
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
            BackColorSel    =   16771802
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   2000
            ColWidthMin     =   0
            ColWidthMax     =   5000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmChargeBat.frx":1C64
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
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ȡ��(&C)"
            Height          =   420
            Left            =   3945
            TabIndex        =   22
            ToolTipText     =   "�ȼ�:Esc"
            Top             =   1710
            Width           =   1560
         End
         Begin VB.Frame fraDrawDept 
            Height          =   1155
            Left            =   0
            TabIndex        =   32
            Top             =   -105
            Width           =   13575
            Begin VB.TextBox txtӦ�� 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   9585
               Locked          =   -1  'True
               TabIndex        =   13
               TabStop         =   0   'False
               Text            =   "0.00"
               Top             =   165
               Width           =   2175
            End
            Begin VB.ComboBox cboDrawDept 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   4140
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   720
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.TextBox txtMemo 
               BackColor       =   &H00E0E0E0&
               Height          =   360
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   210
               Width           =   7485
            End
            Begin VB.ComboBox cboִ������ 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   1155
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   720
               Width           =   1725
            End
            Begin MSMask.MaskEdBox txtDate 
               Height          =   360
               Left            =   9480
               TabIndex        =   19
               Top             =   720
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   635
               _Version        =   393216
               AutoTab         =   -1  'True
               HideSelection   =   0   'False
               MaxLength       =   19
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "yyyy-MM-dd hh:mm:ss"
               Mask            =   "####-##-## ##:##:##"
               PromptChar      =   "_"
            End
            Begin VB.Label lblӦ�� 
               AutoSize        =   -1  'True
               Caption         =   "Ӧ��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   9015
               TabIndex        =   12
               Top             =   240
               Width           =   510
            End
            Begin VB.Line lnSplitH 
               BorderColor     =   &H80000010&
               X1              =   30
               X2              =   18000
               Y1              =   615
               Y2              =   615
            End
            Begin VB.Line lnSplitB 
               BorderColor     =   &H80000014&
               X1              =   0
               X2              =   18000
               Y1              =   630
               Y2              =   630
            End
            Begin VB.Label lblDate 
               AutoSize        =   -1  'True
               Caption         =   "ʱ��"
               Height          =   240
               Left            =   8955
               TabIndex        =   18
               Top             =   780
               Width           =   480
            End
            Begin VB.Label lblDrawDrugDept 
               AutoSize        =   -1  'True
               Caption         =   "��ҩ����"
               Height          =   255
               Left            =   3090
               TabIndex        =   16
               Top             =   773
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label lbl���˱�ע 
               AutoSize        =   -1  'True
               Caption         =   "��ע"
               Height          =   240
               Left            =   585
               TabIndex        =   10
               Top             =   270
               Width           =   480
            End
            Begin VB.Label lblִ������ 
               AutoSize        =   -1  'True
               Caption         =   "ִ������"
               Height          =   240
               Left            =   120
               TabIndex        =   14
               Top             =   780
               Width           =   960
            End
         End
         Begin VB.CommandButton cmdOK 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ȷ��(&O)"
            Height          =   420
            Left            =   3960
            TabIndex        =   21
            ToolTipText     =   "�ȼ���F2"
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   360
         Left            =   1110
         TabIndex        =   4
         Text            =   "cbo��������"
         Top             =   30
         Width           =   2160
      End
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   5760
         Left            =   60
         TabIndex        =   9
         Top             =   585
         Width           =   13065
         _ExtentX        =   23045
         _ExtentY        =   10160
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         TxtCheck        =   -1  'True
         TxtCheck        =   -1  'True
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   360
         RowHeightMin    =   360
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   240
         Left            =   3690
         TabIndex        =   5
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         Caption         =   "��������"
         Height          =   240
         Left            =   30
         TabIndex        =   3
         Top             =   90
         Width           =   960
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1605
      Top             =   930
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":1CB5
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":224F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":27E9
            Key             =   "ǩ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":2B3B
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":939D
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":FBFF
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeBat.frx":10199
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picHead 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   45
      ScaleHeight     =   780
      ScaleWidth      =   15300
      TabIndex        =   23
      Top             =   30
      Width           =   15300
      Begin VB.Frame fraTitle 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   0
         TabIndex        =   24
         ToolTipText     =   "���:F6"
         Top             =   0
         Width           =   15015
         Begin VB.CommandButton cmdSaveWholeSet 
            Caption         =   "����Ϊ�����շ���Ŀ(&W)"
            Height          =   375
            Left            =   3300
            TabIndex        =   30
            Top             =   180
            Width           =   2790
         End
         Begin VB.ComboBox cboNO 
            ForeColor       =   &H00C00000&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   13485
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   195
            Width           =   1425
         End
         Begin VB.CheckBox chkIn 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "������ʵ�:F3"
            Top             =   180
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox txtIn 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   390
            Left            =   585
            MaxLength       =   8
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   180
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.CommandButton cmdSelWholeSet 
            Caption         =   "����(&T)"
            Height          =   375
            Left            =   2190
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   180
            Width           =   1080
         End
         Begin VB.Label lblNO 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "���ݺ�"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   12705
            TabIndex        =   29
            Top             =   255
            Width           =   720
         End
      End
   End
   Begin VB.PictureBox picPatiList 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9750
      Left            =   15
      ScaleHeight     =   9750
      ScaleWidth      =   2955
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1005
      Width           =   2955
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   7455
         Left            =   105
         TabIndex        =   0
         Top             =   360
         Width           =   2655
         _Version        =   589884
         _ExtentX        =   4683
         _ExtentY        =   13150
         _StockProps     =   0
         BorderStyle     =   2
         AutoColumnSizing=   0   'False
      End
      Begin VB.ComboBox cbo��������ѡ�� 
         Height          =   360
         Left            =   1965
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   15
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.ComboBox cbo������ѡ�� 
         Height          =   360
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   2160
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   30
      Top             =   795
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChargeBat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------
'����Ϊ��������
Private mlng����ID As Long
Private mlngDeptID  As Long
Private mlng����ID As Long
Private mlngModule As Long
Private mstrPrivs As String 'ģ��Ȩ�޴�
Private mbln���� As Boolean '33744
Private mblnNurseStation As Boolean
Private mbytUseType As Byte '���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����,2-ҽ�����Ҽ���

'----------------------------------------------------------------------
Private mstrPrivsOpt As String '���ʲ���Ȩ��
Private mstr������Ŀ As String '������Ŀ
'�����ϼ�,0,1;Ӧ�պϼ�,0,7;ʵ�պϼ�,0,7 ��ʱ������,��Ҫԭ���������ϼ�,Ӧ�պϼƵ���Ҫ���ݲ���ѡ������,���ܻ�Ӱ������.
Private Const STR_HEAD = "" & _
"��,450,4;���,750,1;��Ŀ,2175,1;��Ʒ��,1800,1;���,1105,1;��λ,520,4;����,520,1;����,570,1;����,1055,7;" & _
"Ӧ�ս��,1030,7;ִ�п���,1255,1;��־,520,4;����,520,4"

'�ڲ�����
Private mblnOK As Boolean    '���ݱ����Ƿ�湦
'���ݶ���
Private mrsClass As ADODB.Recordset '���ݲ�����ȡ�ĵ�ǰ���õ��շ����
Private mrsUnit As ADODB.Recordset '��ѡ���ִ�п���
Private mrsPati As New ADODB.Recordset '������Ϣ
Private mrsMedAudit As ADODB.Recordset  '�����������ķ�����Ŀ
Private mrsWork As New ADODB.Recordset '�����ϰ��ҩ��
Private mrsWarn As ADODB.Recordset  '����������
Private mrsMedPayMode As ADODB.Recordset '���п��õ�ҽ�Ƹ��ʽ
Private mrs�������� As ADODB.Recordset '��������
Private mrs�������� As ADODB.Recordset  '��ѡ�Ŀ�������
Private mrs������ As ADODB.Recordset    '��ѡҽ���ͻ�ʿ
Private mrs��ҩ���� As ADODB.Recordset
Private mobjItem As XtremeReportControl.IReportRecordItem
Private mobjName As XtremeReportControl.IReportRecordItem
Private mobjBaseItem As Object    '������Ŀ���ò���

'�������
Private mobjBill As ExpenseBill '������õ��ݶ������
Private mcolBillDetails As BillDetails '���ݵ��շ�ϸĿ��
Private mobjBillDetail As BillDetail '���ݵ��շ�ϸĿ����
Private mcolBillInComes As BillInComes '�շ�ϸĿ��������Ŀ��
Private mobjBillIncome As BillInCome '�շ�ϸĿ��������Ŀ����
Private mobjDetail As Detail '�������շ�ϸĿ����
Private mcolDetails As Details '�������շ�ϸĿ����
Private mcolMoneys As BillInComes  '���������Ŀ���ܼ���(��ʾ����ӡʱʹ��)���
'���ö������
Private Enum mPatiCol
    COL_����ID = 0
    COL_��ҳID = 1
    COL_ѡ�� = 2
    COL_���� = 3
    COL_���� = 4
    COL_�Ա� = 5
    COL_���� = 6
    COL_סԺ�� = 7
    COL_�ѱ� = 8
    COL_���� = 9
    COL_������� = 10
    COL_Ӥ�� = 11
    COL_ʣ��� = 12
    COL_Ԥ����� = 13
    COL_������� = 14
    COL_������ = 15
    COL_���ն� = 16
    COL_���ò��� = 17
    COL_ҽ�Ƹ��ʽ = 18
    COL_������ = 19
    COL_��������ID = 20
    COL_�������� = 21
End Enum
Private Enum mPanceIdx
    EM_HeadList = 1
    EM_PatiList = 2
    EM_BILLList = 3
End Enum
Private Enum BillColType       '���ݿؼ���������
    CheckBox = -1
    Text_UnModify = 0
    CommandButton = 1
    Date = 2
    ComboBox = 3
    Text = 4
    UnFocus = 5
End Enum

Private Enum BillCol
    �� = 0
    ��� = 1
    ��Ŀ = 2
    ��Ʒ�� = 3
    ��� = 4
    ��λ = 5
    ���� = 6
    ���� = 7
    ���� = 8
    Ӧ�ս�� = 9
    ִ�п��� = 10
    ��־ = 11
    ���� = 12
End Enum
'�������
Private mblncboEnterCell As Boolean '����ѭ������
Private mblncboClick  As Boolean    '����ѭ������
Private mlngPreRow As Long '��ǰ�к�,�����иı�ʱ�ж�
Private mcolStock1 As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mcolStock2 As Collection '��Ÿ������Ŀ�ĳ����鷽ʽ
Private mbln����ְ���� As Boolean     '�Ƿ���д���ְ����
Private mbln����������� As Boolean     '�Ƿ���д����������
Private mblnOne As Boolean '�Ƿ�ֻ��һ�������շ����
Private mblnWork As Boolean '��ǰ�Ƿ��������ϰ��ҩ��
Private mlngҩƷ���ID As Long '��ǰ���ݲ�����ҩƷ������ID
Private mlng�������ID As Long '��ǰ���ݲ���������������ID
Private mstrUnitIDs As String   '��ǰ����Ա�����в���ID
Private mstrWarn As String '�Ѿ���������ѡ����������
Private mblnSendMateria As Boolean  '���ʺ��Զ���ҩ
Private mblnFirst As Boolean
Private mlngX As Long, mlngY As Long
Private mblnDrop As Boolean '��KeyDown���ж�cbo�����˵�ǰ�Ƿ񵯳�
Private mblnValid As Boolean
Private mblnNewRow As Boolean
Private mdblItemNum As Double '���ݿ��е�ǰ�����Ŀ������
Private mblnSelect As Boolean '���ڿ����շ�ϸĿ�����Ƿ��������б�ѡ���ѡ����
Private mblnNotClick As Boolean
Private mstr����IDs As String   '��ǰѡ�еĲ���IDs
Private mlngSelPatiCount As Long  '��ǰѡ�еĲ�������
Private mstrInsures As String   '��ǰѡ�е�ҽ������,����ö��ŷ���
Private mblnPrintDrugList As Boolean '�Ƿ��ӡ��ҩ�嵥
Private mblnKeyReturn As Boolean '�Ƿ����˻س�

'ҽ����ز���
Private Type TYPE_MedicarePAR
    �������� As Boolean
    �����ϴ� As Boolean
    ������ɺ��ϴ� As Boolean
    ���������ϴ� As Boolean
    ʵʱ��� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR
Private Type TY_PATIINFOR
    ����ID As Long
    ��ҳID As Long
    Ӥ�� As Integer
    ���� As Integer
    ���� As String
    ���� As String
    �Ա� As String
    ���ò��� As String
    סԺ�� As String
    ���� As String
    �ѱ� As String
    ҽ�Ƹ��ʽ As String
    ��Ժ���� As String
    ��Ժ����  As String
    �������� As Integer
    ״̬ As Integer
    ������� As String
    ʣ��� As Currency
    Ԥ����� As Currency
    ������� As Currency
    ������ As Currency
    ���ն� As Currency
    ������ As String
    ��������ID As Long
End Type
Private mstr���ת��ʱ�� As String

Private Enum Pan
    C2��ʾ��Ϣ = 2
End Enum
Private mstrҩƷ�۸�ȼ� As String, mstr���ļ۸�ȼ� As String, mstr��ͨ�۸�ȼ� As String

Public Function ShowMe(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal bytUseType As Byte, ByVal lng����ID As Long, ByVal lngDeptID As Long, ByVal lng����ID As Long, _
    ByVal bln���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:frmMain-���õ�������
    '     lng����ID-����ID
    '     lng����ID-ָ���Ĳ���
    '     bln����-�Ƿ񲹷�
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-07-08 17:46:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mblnOK = False: mlngModule = lngModule: mstrPrivs = strPrivs
    mlng����ID = lng����ID: mlng����ID = lng����ID: mlngDeptID = lngDeptID
    mbln���� = False: mbytUseType = bytUseType
    
    If gblnNurseStation Then
        mblnNurseStation = True
    Else
        mblnNurseStation = False
    End If
    
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    
    ShowMe = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub rptPati_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    mlngX = X
    mlngY = Y
End Sub

Private Function GetRptPositionX(intTYPE As Integer) As Long
    On Error GoTo errH
    Dim i As Long
    If intTYPE = 1 Then
        For i = 0 To mlngX
            If rptPati.HitTest(i, 0).Column.Caption = "������" Then
                GetRptPositionX = i
                Exit For
            End If
        Next i
    Else
        For i = 0 To mlngX
            If rptPati.HitTest(i, 0).Column.Caption = "��������" Then
                GetRptPositionX = i
                Exit For
            End If
        Next i
    End If
    Exit Function
errH:
    Err.Clear
    GetRptPositionX = mlngX
End Function

Private Function GetRptPositionY() As Long
    On Error GoTo errH
    Dim i As Long
    For i = 0 To mlngY
        If rptPati.HitTest(mlngX, i).Row Is rptPati.SelectedRows(0) Then
            GetRptPositionY = i
            Exit For
        End If
    Next i
    Exit Function
errH:
    Err.Clear
    GetRptPositionY = mlngY
End Function

Private Sub LockedScreen(ByVal blnLocked As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ļ,�Ա��ڱ���ʱ,������صİ�ť
    '���:blnLocked-true:������Ļ;False-������
    '����:���˺�
    '����:2015-07-13 10:49:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    cmdOK.Enabled = Not blnLocked
    cmdCancel.Enabled = Not blnLocked
    picHead.Enabled = Not blnLocked
    picBillBottom.Enabled = Not blnLocked
    picPatiList.Enabled = Not blnLocked
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub
Private Sub cbo������_GotFocus()
    mblnKeyReturn = False
End Sub

Private Sub cmdOK_Click()
    
    If isValied() = False Then Exit Sub
    
    '���ݱ���
    Call LockedScreen(True)
    If SaveData = False Then
        Call LockedScreen(False)
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        Exit Sub
    End If
    Call LockedScreen(False)

    '�ָ���ǰվ��۸�ȼ�
    mstrҩƷ�۸�ȼ� = gstrҩƷ�۸�ȼ�
    mstr���ļ۸�ȼ� = gstr���ļ۸�ȼ�
    mstr��ͨ�۸�ȼ� = gstr��ͨ�۸�ȼ�
    
    Call ClearRows: Call Bill.ClearBill: Call SetColNum
    Call ClearMoney
    Call SetMoneyList
    Call NewBill
    Call SetDrawDrugDeptEnabled
    If rptPati.Visible Then rptPati.SetFocus
    mblnOK = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call LockedScreen(False)
End Sub

Private Sub cmdCancel_Click()
    If mobjBill.Details.Count = 0 Or Not Bill.Active Then Unload Me: Exit Sub
    
    If MsgBox("ȷʵҪ�����ǰ�����е�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '�������
    chk����.Value = 0: chk����.Visible = False
    txtӦ��.Text = gstrDec
    'txtʵ��.Text = gstrDec:
    Call ClearRows: Call Bill.ClearBill
    Call SetColNum: Call ClearMoney
    Call NewBill
    If Bill.Enabled And Bill.Visible Then Bill.SetFocus
End Sub

Private Sub dkpMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If sta.Visible Then Bottom = sta.Height
End Sub

Private Sub cbo��������ѡ��_Click()
    If mblnNotClick Then Exit Sub
    If Not mobjItem Is Nothing Then
        mobjItem.Value = cbo��������ѡ��.ItemData(cbo��������ѡ��.ListIndex)
        mobjName.Value = zlStr.NeedName(cbo��������ѡ��.Text)
    End If
    cbo��������ѡ��.Visible = False
    rptPati.Populate
End Sub

Private Sub cbo��������ѡ��_LostFocus()
    cbo��������ѡ��.Visible = False
End Sub

Private Sub cbo������ѡ��_Click()
    If mblnNotClick Then Exit Sub
    If Not mobjItem Is Nothing Then
        mobjItem.Value = zlStr.NeedName(cbo������ѡ��.Text)
    End If
    cbo������ѡ��.Visible = False
    rptPati.Populate
End Sub

Private Sub cbo������ѡ��_LostFocus()
    cbo������ѡ��.Visible = False
End Sub

Private Sub Form_Load()
    Dim tmpBill As ExpenseBill
    Dim i As Long, lngPre As Long, strPre As String, strTmp As String, strҩ��IDs As String
    glngFormW = 15345: glngFormH = 11520
    If Not OS.IsDesinMode Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    Call InitPanel
    
    RestoreWinState Me, App.ProductName, Me.Name
    sta.Visible = True
    
    mblnValid = False: mblnFirst = True: gbln�������� = False
    
    chkIn.Visible = True: txtIn.Visible = True
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
    
    Call zlLoadDrawDeptData(mbytUseType, mlngDeptID)
    
    '��ʼ����������
    Set mobjBill = New ExpenseBill

    mstrUnitIDs = GetUserUnits
    
    
    '���ز�����Ϣ
    If Not InitData Then Unload Me: Exit Sub
    
    mstrҩƷ�۸�ȼ� = gstrҩƷ�۸�ȼ�
    mstr���ļ۸�ȼ� = gstr���ļ۸�ȼ�
    mstr��ͨ�۸�ȼ� = gstr��ͨ�۸�ȼ�
    
    Call InitFace: Call NewBill
    Call LoadPatiInfo
    cbo��������.SelStart = 0
    cbo��������.SelLength = 0
    Call Auto��������
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    
    '������ҩ����
    Call SetDrawDrugDeptVisible
    
    mblnFirst = False
    On Error Resume Next
    'If Bill.Visible Then Bill.SetFocus

    Call SetDrawDrugDeptEnabled
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If ActiveControl Is cbo������ Then Call cbo������_KeyPress(vbKeyReturn)
            If cmdOK.Enabled And cmdOK.Visible Then
                Call cmdOK.SetFocus
                Call cmdOK_Click
            End If
        Case vbKeyF3    '���뵥��
            If chkIn.Visible And chkIn.Enabled Then chkIn.Value = IIf(chkIn.Value = 1, 0, 1)
        Case vbKeyF4
        Case vbKeyF6    '��λ������ѡ���
            If rptPati.Visible Then rptPati.SetFocus
        Case vbKeyF7    '�л����뷨
            If Not gbln�����л� Then Exit Sub
            If Not (sta.Panels("WB").Visible And sta.Panels("PY").Visible) Then Exit Sub
            
            If sta.Panels("WB").Bevel = sbrRaised Then
                Call sta_PanelClick(sta.Panels("WB"))
            Else
                Call sta_PanelClick(sta.Panels("PY"))
            End If
        Case vbKeyF9 '��λ�����ݺ������
            cboNO.SetFocus
            Call zlControl.TxtSelAll(cboNO)
        Case vbKeyF11
            'If cmd�䷽.Enabled And cmd�䷽.Visible Then Call cmd�䷽_Click
        Case vbKeyF12
            If Shift <> vbAltMask Then Exit Sub
            
            Call sta_PanelClick(sta.Panels("Drugstore"))
        Case vbKeyA, vbKeyR             'ȫѡ��ȫ��
        Case vbKeyQ
            If Shift <> vbCtrlMask Then Exit Sub
            Call LocateNewRow
        Case vbKeyEscape
            If Bill.TxtVisible Then
                Bill.Text = "": Bill.TxtVisible = False: Bill.SetFocus
            Else
                Call cmdCancel_Click
            End If
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
 
    If InStr("',|~" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If Me.ActiveControl Is Bill Or Me.ActiveControl Is txtMemo Then Exit Sub
    '����:29464
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub  '���ܴ������Ƶ�ˢ��:   ;1088029?
End Sub


Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub


Private Sub InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:������
    '����:2014-06-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngWidth As Long, strReg As String, panThis As Pane
    Dim panTop As Pane, panBottom As Pane, panRight As Pane
    Dim lngHeight As Long
    Dim strName As String
    If mlng����ID <> 0 Then
       strName = "��ǰ����:" & GetDeptName(mlng����ID)
    ElseIf mlngDeptID = 0 Then
       strName = "��ǰ����:" & GetDeptName(mlngDeptID)
    Else
        strName = "������Ϣ"
    End If
    
    Set panThis = dkpMain.CreatePane(mPanceIdx.EM_HeadList, 250, 580, DockTopOf, Nothing)
    lngHeight = picHead.Height / Screen.TwipsPerPixelY
    panThis.Title = ""
    panThis.Tag = mPanceIdx.EM_HeadList
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picHead.hWnd
    panThis.MaxTrackSize.Height = lngHeight
    panThis.MinTrackSize.Height = lngHeight
    
    lngWidth = 2955 / Screen.TwipsPerPixelX
    Set panThis = dkpMain.CreatePane(mPanceIdx.EM_PatiList, lngWidth, 300, DockBottomOf, panThis)
    
    panThis.Title = strName
    panThis.Tag = mPanceIdx.EM_PatiList
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picPatiList.hWnd
    
    lngWidth = (Me.ScaleWidth - 2955) / Screen.TwipsPerPixelX
    Set panThis = dkpMain.CreatePane(mPanceIdx.EM_BILLList, lngWidth, 580, DockRightOf, panThis)
    panThis.Title = ""
    panThis.Tag = mPanceIdx.EM_BILLList
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picBillList.hWnd
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    
    Set dkpMain.PaintManager.CaptionFont = lbl��������.Font
    'zlRestoreDockPanceToReg Me, dkpMan, "����"
End Sub
 
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mPanceIdx.EM_HeadList
        Item.Handle = picHead.hWnd
    Case mPanceIdx.EM_PatiList
        Item.Handle = picPatiList.hWnd
    Case mPanceIdx.EM_BILLList
        Item.Handle = picBillList.hWnd
    End Select
End Sub


Private Sub InitReportColumn()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Report�ؼ���
    '����:���˺�
    '����:2015-07-09 11:15:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCol As ReportColumn, lngIdx As Long, i As Long
    
    On Error GoTo errHandle
    
    With rptPati
        
        Set objCol = .Columns.Add(COL_����ID, "����ID", 0, False)
        Set objCol = .Columns.Add(COL_��ҳID, "��ҳID", 0, False)
        Set objCol = .Columns.Add(COL_ѡ��, "", 20, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("UnCheck").Index - 1
        Set objCol = .Columns.Add(COL_����, "����", 45, True)
        Set objCol = .Columns.Add(COL_����, "����", 120, True)
        Set objCol = .Columns.Add(COL_�Ա�, "�Ա�", 30, True)
        Set objCol = .Columns.Add(COL_����, "����", 30, True)
        Set objCol = .Columns.Add(COL_סԺ��, "סԺ��", 60, True)
        Set objCol = .Columns.Add(COL_�ѱ�, "�ѱ�", 60, True)
        
        Set objCol = .Columns.Add(COL_����, "����", 0, False)
        Set objCol = .Columns.Add(COL_�������, "�������", 0, False)
        Set objCol = .Columns.Add(COL_Ӥ��, "Ӥ��", 0, False)
        Set objCol = .Columns.Add(COL_ʣ���, "ʣ���", 80, True)
        Set objCol = .Columns.Add(COL_Ԥ�����, "Ԥ�����", 80, True)
        Set objCol = .Columns.Add(COL_�������, "�������", 80, True)
        Set objCol = .Columns.Add(COL_������, "������", 0, False)
        Set objCol = .Columns.Add(COL_���ն�, "���ն�", 0, False)
        Set objCol = .Columns.Add(COL_���ò���, "���ò���", 0, False)
        Set objCol = .Columns.Add(COL_ҽ�Ƹ��ʽ, "ҽ�Ƹ��ʽ", 0, False)
        If mblnNurseStation Then
            Set objCol = .Columns.Add(COL_������, "������", 90, True)
            Set objCol = .Columns.Add(COL_��������ID, "��������ID", 0, False)
            Set objCol = .Columns.Add(COL_��������, "��������", 150, True)
        End If

        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĳ���..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub LoadPatiInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�������ز�����Ϣ
    '����:���˺�
    '����:2015-07-08 10:35:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, strSqlBaby As String
    Dim objRecord As ReportRecord, objItem As ReportRecordItem
    Dim lngSelectRow As Long, lng����ID As Long, rsBaby As ADODB.Recordset
    Dim strSQL As String, objChild As ReportRecord
    Dim lngColor As Long
    
     On Error GoTo errH
 
    Set mrsPati = GetPatiRsByUnit(mlng����ID, mlng����ID, True, True, False)
    lngSelectRow = -1
    With rptPati
        .Columns.Column(3).TreeColumn = True
        .PaintManager.TreeStructureStyle = xtpTreeStructureNone
        For i = 1 To mrsPati.RecordCount
            If Val(mrsPati!��˱�־ & "") < 1 Or gTy_System_Para.byt������˷�ʽ <> 1 Then
                If Val(mrsPati!Ӥ����� & "") = 0 Then
                    Set objRecord = .Records.Add()
                    objRecord.Tag = "0"
                    Set objItem = objRecord.AddItem(mrsPati!����ID & "")
                    Set objItem = objRecord.AddItem(mrsPati!��ҳID & "")
                    Set objItem = objRecord.AddItem("")
                    Set objItem = objRecord.AddItem(mrsPati!���� & "")
                    Set objItem = objRecord.AddItem(mrsPati!���� & "")
                        objItem.Icon = img16.ListImages.Item(IIf(mrsPati!�Ա� & "" = "��", "Man", "Woman")).Index - 1
                    Set objItem = objRecord.AddItem(mrsPati!�Ա� & "")
                    Set objItem = objRecord.AddItem(mrsPati!���� & "")
                    Set objItem = objRecord.AddItem(mrsPati!סԺ�� & "")
                    Set objItem = objRecord.AddItem(mrsPati!�ѱ� & "")
                    Set objItem = objRecord.AddItem(mrsPati!���� & "")
                    Set objItem = objRecord.AddItem(mrsPati!������� & "")
                    Set objItem = objRecord.AddItem(Val(mrsPati!Ӥ����� & ""))
                    Set objItem = objRecord.AddItem(Format(Val(mrsPati!ʣ��� & ""), "0.00"))
                    Set objItem = objRecord.AddItem(Format(Val(mrsPati!Ԥ����� & ""), "0.00"))
                    Set objItem = objRecord.AddItem(Format(Val(mrsPati!������� & ""), "0.00"))
                    Set objItem = objRecord.AddItem(Format(Val(mrsPati!������ & ""), "0.00"))
                    Set objItem = objRecord.AddItem(Format(Val(mrsPati!���ն� & ""), "0.00"))
                    Set objItem = objRecord.AddItem(Trim(mrsPati!���ò��� & ""))
                    Set objItem = objRecord.AddItem(Trim(mrsPati!ҽ�Ƹ��ʽ & ""))
                    If mblnNurseStation Then
                        Set objItem = objRecord.AddItem(Trim(mrsPati!������ & ""))
                        Set objItem = objRecord.AddItem(Val(Nvl(mrsPati!��������ID)))
                        Set objItem = objRecord.AddItem(Nvl(mrsPati!������������))
                    End If
                Else
                    strSqlBaby = "Select Ӥ������, Ӥ���Ա�, Zl_Age_Calc(0, ����ʱ��, Sysdate) As ���� From ������������¼ Where ����id = [1] And ��ҳid = [2] And ��� = [3]"
                    Set rsBaby = zlDatabase.OpenSQLRecord(strSqlBaby, Me.Caption, Val(mrsPati!����ID & ""), Val(mrsPati!��ҳID & ""), Val(mrsPati!Ӥ����� & ""))
                    Set objChild = objRecord.Childs.Add
                    Set objItem = objChild.AddItem(mrsPati!����ID & "")
                    Set objItem = objChild.AddItem(mrsPati!��ҳID & "")
                    Set objItem = objChild.AddItem("")
                    Set objItem = objChild.AddItem(mrsPati!���� & "")
                    If Not rsBaby.EOF Then
                        Set objItem = objChild.AddItem(Nvl(rsBaby!Ӥ������))
                            objItem.Icon = img16.ListImages.Item(IIf(InStr(rsBaby!Ӥ���Ա�, "��") > 0, "Man", "Woman")).Index - 1
                        Set objItem = objChild.AddItem(IIf(InStr(rsBaby!Ӥ���Ա�, "��") > 0, "��", "Ů"))
                        Set objItem = objChild.AddItem(rsBaby!���� & "")
                    Else
                        Set objItem = objChild.AddItem(mrsPati!���� & "")
                            objItem.Icon = img16.ListImages.Item(IIf(mrsPati!�Ա� & "" = "��", "Man", "Woman")).Index - 1
                        Set objItem = objChild.AddItem(mrsPati!�Ա� & "")
                        Set objItem = objChild.AddItem(mrsPati!���� & "")
                    End If
                    Set objItem = objChild.AddItem(mrsPati!סԺ�� & "")
                    Set objItem = objChild.AddItem(mrsPati!�ѱ� & "")
                    Set objItem = objChild.AddItem(mrsPati!���� & "")
                    Set objItem = objChild.AddItem(mrsPati!������� & "")
                    Set objItem = objChild.AddItem(Val(mrsPati!Ӥ����� & ""))
                    Set objItem = objChild.AddItem(Format(Val(mrsPati!ʣ��� & ""), "0.00"))
                    Set objItem = objChild.AddItem(Format(Val(mrsPati!Ԥ����� & ""), "0.00"))
                    Set objItem = objChild.AddItem(Format(Val(mrsPati!������� & ""), "0.00"))
                    Set objItem = objChild.AddItem(Format(Val(mrsPati!������ & ""), "0.00"))
                    Set objItem = objChild.AddItem(Format(Val(mrsPati!���ն� & ""), "0.00"))
                    Set objItem = objChild.AddItem(Trim(mrsPati!���ò��� & ""))
                    Set objItem = objChild.AddItem(Trim(mrsPati!ҽ�Ƹ��ʽ & ""))
                    If mblnNurseStation Then
                        Set objItem = objChild.AddItem(Trim(mrsPati!������ & ""))
                        Set objItem = objChild.AddItem(Val(Nvl(mrsPati!��������ID)))
                        Set objItem = objChild.AddItem(Nvl(mrsPati!������������))
                    End If
                End If
                
                '������ɫ
                If Not IsNull(mrsPati!��������) Then
                    '���ղ�����ָ��ɫ��ʾ
                    lngColor = zlDatabase.GetPatiColor(mrsPati!��������)
                    For j = 0 To rptPati.Columns.Count - 1
                        objRecord.Item(j).ForeColor = lngColor
                    Next
                ElseIf Not IsNull(mrsPati!����) Then
                    'δָ���������͵ı��ղ����ú�ɫ��ʾ
                    For j = 0 To rptPati.Columns.Count - 1
                        objRecord.Item(j).ForeColor = vbRed
                    Next
                End If
                '�ϴ��Ƿ�ѡ��
                If mrsPati!����ID = mlng����ID Then
                    objRecord.Item(COL_ѡ��).Icon = img16.ListImages.Item("Check").Index - 1
                    objRecord.Tag = "1"
                    lngSelectRow = objRecord.Index
                    mlngSelPatiCount = mlngSelPatiCount + 1
                End If
            End If
            mrsPati.MoveNext
        Next
        .Populate
        If .Records.Count <> 0 Then Set .FocusedRow = .Rows(0)
        If lngSelectRow <> -1 Then Set .FocusedRow = .Rows(lngSelectRow)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
   SaveWinState Me, App.ProductName, Me.Name

    mlngҩƷ���ID = 0: mlng�������ID = 0
    Set mrs�������� = Nothing
    Set mrs�������� = Nothing
    Set mrs������ = Nothing
    Set mrsWarn = Nothing
    Set mrsMedAudit = Nothing
    Set mrsMedPayMode = Nothing
    Set mobjBaseItem = Nothing
    If Not OS.IsDesinMode Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    mlngSelPatiCount = 0
    mblnNurseStation = False
    mstrInsures = ""
    mstr����IDs = ""
End Sub

Private Sub picBillBottom_Resize()
    Err = 0: On Error Resume Next
    If Not lblִ������.Visible Then
        lblDrawDrugDept.Left = lblִ������.Left
        cboDrawDept.Left = cboִ������.Left
    Else
        lblDrawDrugDept.Left = cboִ������.Left + cboִ������.Width + 100
        cboDrawDept.Left = lblDrawDrugDept.Left + lblDrawDrugDept.Width + 10
    End If
    txtӦ��.Left = picBillBottom.ScaleWidth - txtӦ��.Width - 100
    lblӦ��.Left = txtӦ��.Left - lblӦ��.Width - 10
    
    txtMemo.Width = lblӦ��.Left - txtMemo.Left - 100
    fraDrawDept.Width = picBillBottom.ScaleWidth - fraDrawDept.Left + 10
    lnSplitB.X2 = fraDrawDept.Width
    lnSplitH.X2 = fraDrawDept.Width
    txtDate.Left = picBillBottom.ScaleWidth - txtDate.Width - 100
    lblDate.Left = txtDate.Left - lblDate.Width - 10
End Sub

Private Sub picBillList_Resize()
    Err = 0: On Error Resume Next
    With picBillList
        
        picBillBottom.Left = .ScaleLeft
        picBillBottom.Top = .ScaleHeight - picBillBottom.Height - 100
        picBillBottom.Width = .ScaleWidth
        
        Bill.Top = cbo��������.Top + cbo��������.Height + 50
        Bill.Left = 50
        Bill.Width = .ScaleWidth - Bill.Left * 2
        Bill.Height = picBillBottom.Top - Bill.Top - 50
    End With
End Sub
Private Sub picHead_Resize()
    Err = 0: On Error Resume Next
    With picHead
        fraTitle.Width = .ScaleWidth - fraTitle.Left
        cboNO.Left = fraTitle.Width - cboNO.Width - 100
        lblNO.Left = cboNO.Left - lblNO.Width
    End With
End Sub

Private Sub picPatiList_Resize()
    Err = 0: On Error Resume Next
    With picPatiList
        rptPati.Top = 50
        rptPati.Left = 50
        rptPati.Width = .ScaleWidth - rptPati.Left * 2
        rptPati.Height = .ScaleHeight - rptPati.Top * 2
    End With
End Sub
 
Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cbo��������.Enabled And cbo��������.Visible Then
            Debug.Print "cbo��������.SetFocus"
            cbo��������.SetFocus
            Debug.Print TypeName(Me.ActiveControl)
            Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    End If
    If KeyCode <> vbKeySpace Then Exit Sub
    If rptPati.SelectedRows.Count <= 0 Then Exit Sub
    Call rptPati_RowDblClick(rptPati.SelectedRows(0), rptPati.SelectedRows(0).Record.Item(COL_ѡ��))
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objColumn As ReportColumn, i As Long
    Dim j As Long
    '��������ͷ��ͼƬ����ѡ��ȫ��
    If Button <> 1 Then Exit Sub
    If rptPati.HitTest(X, Y).ht <> xtpHitTestHeader Then Exit Sub

    Set objColumn = rptPati.HitTest(X, Y).Column
    If objColumn Is Nothing Then Exit Sub
    If objColumn.Index <> COL_ѡ�� Then Exit Sub

    If objColumn.Caption = "" Then
        objColumn.Caption = "1"
        rptPati.Columns(COL_ѡ��).Icon = img16.ListImages("AllCheck").Index - 1
        For i = 0 To rptPati.Records.Count - 1
            rptPati.Records(i)(COL_ѡ��).Icon = img16.ListImages("Check").Index - 1
            For j = 0 To rptPati.Records(i).Childs.Count - 1
                rptPati.Records(i).Childs.Record(j).Item(COL_ѡ��).Icon = img16.ListImages("Check").Index - 1
            Next j
            rptPati.Rows(i).Record.Tag = "1"
        Next
    Else
        objColumn.Caption = ""
        rptPati.Columns(COL_ѡ��).Icon = img16.ListImages("UnCheck").Index - 1
        For i = 0 To rptPati.Records.Count - 1
            rptPati.Records(i)(COL_ѡ��).Icon = -1
            For j = 0 To rptPati.Records(i).Childs.Count - 1
                rptPati.Records(i).Childs.Record(j).Item(COL_ѡ��).Icon = -1
            Next j
            rptPati.Rows(i).Record.Tag = "0"
        Next
    End If
    mstr����IDs = GetPatiIDsBySel(mlngSelPatiCount)
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Select Case Item.Index
    Case 19
        '������
        cbo������ѡ��.Top = rptPati.Top + GetRptPositionY * 15
        cbo������ѡ��.Left = rptPati.Left + GetRptPositionX(1) * 15
        mblnNotClick = True
        cbo������ѡ��.ListIndex = 0
        zlControl.CboLocate cbo������ѡ��, Item.Value
        mblnNotClick = False
        Set mobjItem = Item
        cbo������ѡ��.ZOrder: cbo������ѡ��.Visible = True
        cbo������ѡ��.SetFocus
    Case 21
        '��������
        cbo��������ѡ��.Top = rptPati.Top + GetRptPositionY * 15
        cbo��������ѡ��.Left = rptPati.Left + GetRptPositionX(2) * 15
        mblnNotClick = True
        cbo��������ѡ��.ListIndex = 0
        zlControl.CboLocate cbo��������ѡ��, Item.Value
        mblnNotClick = False
        Set mobjItem = Row.Record.Item(COL_��������ID)
        Set mobjName = Item
        cbo��������ѡ��.ZOrder: cbo��������ѡ��.Visible = True
        cbo��������ѡ��.SetFocus
    Case Else
        If Row.Record.Tag = "1" Then
            Row.Record.Item(COL_ѡ��).Icon = -1
            Row.Record.Tag = "0"
        Else
            Row.Record.Item(COL_ѡ��).Icon = img16.ListImages.Item("Check").Index - 1
            Row.Record.Tag = "1"
        End If
        rptPati.Populate
        mstr����IDs = GetPatiIDsBySel(mlngSelPatiCount)
        Call Auto��������
    End Select
End Sub

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���浥������
    '����:���˺�
    '����:2015-07-08 14:17:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str����IDs As String '���ʳɹ��Ĳ���ID
    Dim strNotPatiIDs As String '���ʲ��ɹ��Ĳ���ID
    Dim tyPati As TY_PATIINFOR, dtdtCurDate As Date
    Dim blnSavePrice As Boolean '�Ƿ񱣴�Ϊ���۵�
    Dim i As Long, strMessage As String
    Dim lngBringCount As Long, lngNotBringCount As Long
    
    On Error GoTo errHandle
    dtdtCurDate = zlDatabase.Currentdate
 
    '���˺�:���»�ȡ��ҩ����
    Call zlReSetDrawDrugDept
    str����IDs = "": strNotPatiIDs = ""
    lngBringCount = 0: lngNotBringCount = 0
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows.Count > 1 Then
            zlControl.StaShowPercent i / (rptPati.Rows.Count - 1), sta.Panels(2), Me
        End If
        If rptPati.Rows(i).Record.Tag = "1" Then
            tyPati = GetPatiInforByReport(i, mblnNurseStation)
            blnSavePrice = False
            Call zlCommFun.ShowFlash("���ڶԲ���:" & tyPati.���� & "���м���,���Ժ�...")
            
            '��ʼ��ҽ������
            If tyPati.���� <> 0 Then Call InitInsurePara(tyPati.����ID, tyPati.����)
            
            '���¸����ݸ�ֵ
            Call reSetBillObject(tyPati, mobjBill)
            
            '��¼ҽ��ժҪ
            Call InputItemMemo(tyPati)
            
            '����ȡ�۸�ȼ�
            If gintPriceGradeStartType >= 2 Then
                Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, tyPati.����ID, tyPati.��ҳID, tyPati.ҽ�Ƹ��ʽ, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
            End If

            '���¼���ʵ�ս��
            Call CalcMoneys(tyPati)
            
            '�������
            If CheckPatiChargeWrang(tyPati, blnSavePrice) = False Then Exit Function
            mobjBill.�Ǽ�ʱ�� = dtdtCurDate      'ע��:��ӡ��ҩ��ʱҪ�õ����ʱ��
            
            '���¼���ҽ��ͳ��
            Call ReCalcInsure(tyPati)
            
            '�����˱��浥��
    
            If SaveBill(tyPati, blnSavePrice, strMessage) Then
                '��ӡƱ��
                Call BillPrint(blnSavePrice)
                str����IDs = str����IDs & "," & tyPati.����ID
                lngBringCount = lngBringCount + 1
            Else
                strNotPatiIDs = strNotPatiIDs & "," & tyPati.����ID & "|" & strMessage
                lngNotBringCount = lngNotBringCount + 1
            End If
  
        End If
    Next
    Call zlCommFun.StopFlash
       
    str����IDs = Mid(str����IDs, 2)
    strNotPatiIDs = Mid(strNotPatiIDs, 2)
    If lngNotBringCount = 0 Then
        MsgBox "�㹲ѡ����" & mlngSelPatiCount & "������,�ɹ�����:" & lngBringCount & "������!", vbInformation + vbOKOnly, gstrSysName
    Else
        If MsgBox("�㹲ѡ����" & mlngSelPatiCount & "������,���ɹ�����:" & lngBringCount & "������,δ�ɹ�����:" & lngNotBringCount & "������!" & vbCrLf & "�Ƿ�鿴δ���ʳɹ�������ϸ?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            frmChargeBatFailNote.ShowMe Me, strNotPatiIDs
        End If
    End If
    sta.Panels(2).Text = ""
    SaveData = str����IDs <> ""
    Exit Function
errHandle:
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetPatiIDsBySel(ByRef lngCount As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ҫ���ʵĲ���IDs
    '����:lngCount-��ǰѡ�е�����
    '����:���ز���ID, ����ö��ŷָ�
    '����:���˺�
    '����:2015-07-09 10:13:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str����IDs As String, lng����ID As Long, i As Long
    Dim lng���� As Long
    
    On Error GoTo errHandle
    lngCount = 0
    str����IDs = "": mstrInsures = ""
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            lng����ID = Val(rptPati.Rows(i).Record(COL_����ID).Value)
            lng���� = Val(rptPati.Rows(i).Record(COL_����).Value)
            If lng����ID <> 0 Then
                str����IDs = str����IDs & "," & lng����ID
                lngCount = lngCount + 1
            End If
            If lng���� <> 0 And InStr(mstrInsures & ",", "," & lng���� & ",") = 0 Then
                mstrInsures = mstrInsures & "," & lng����
            End If

        End If
    Next
    If str����IDs <> "" Then str����IDs = Mid(str����IDs, 2)
    If mstrInsures <> "" Then mstrInsures = Mid(mstrInsures, 2)
    
    GetPatiIDsBySel = str����IDs
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function SaveBill(tyPati As TY_PATIINFOR, _
    Optional blnSavePrice As Boolean, Optional ByRef strMessage As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���浱ǰ����ļ��ʵ���
    '���:tyPati-��ǰ������Ϣ
    '     blnSavePrice-��ǰ����Ϊ���۵�
    '����:���ݱ��淵��true,���򷵻�False
    '����:���˺�
    '����:2015-07-13 14:27:09
    '˵��:mobjBill=���ݶ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, arrSQL As Variant, arrSMSQL As Variant
    Dim int��� As Integer, int�к� As Integer, strNo As String, strTmp As String, str���ܺ� As String
    Dim intParent As Integer, intParentNO As Integer
    Dim str��Ϣ As String, intInsure As Integer
    Dim dbl���� As Double, dbl���� As Double
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim strSQL As String, strStuffDept As String '��¼���Ϸ��ϲ���
    Dim strAddDate As String '���ʷ���,�Զ���ҩ,���ϵ�ʱ��
    Dim blnTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim str��ҩ��̬ As String
    Dim cllSMSQL As Collection, cllExctPro As Collection
    Dim varData As Variant, varTemp As Variant
    Dim rsItems As ADODB.Recordset
    Dim lng��������ID As Long
    Err = 0: On Error GoTo ErrHand:
    mobjBill.NO = zlDatabase.GetNextNo(14)
    mobjBill.����ID = mlng����ID
    mobjBill.����ʱ�� = CDate(txtDate.Text)
    mobjBill.�Ǽ�ʱ�� = zlDatabase.Currentdate      'ע��:��ӡ��ҩ��ʱҪ�õ����ʱ��
    strAddDate = "To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    gstrModiNO = mobjBill.NO
    arrSQL = Array()
    Set cllSMSQL = New Collection
    Set cllExctPro = New Collection
    
    If zlGetSaveDataItems_Plugin(mobjBill, rsItems) = False Then strMessage = "���ʺϷ��Լ��ʧ��!": Exit Function
    If zlChargeSaveValied_Plugin(mlngModule, 2, False, gbytBilling = 1, "", rsItems) = False Then strMessage = "���ʺϷ��Լ��ʧ��!": Exit Function
    
    For Each mobjBillDetail In mobjBill.Details
        If mobjBillDetail.���� <> 0 Then
            intParent = 0: intParentNO = int���
            For Each mobjBillIncome In mobjBillDetail.InComes
                int��� = int��� + 1 '��ǰ��¼���
                '��������
                With mobjBill
                    If tyPati.�������� <> 1 Then
                        gstrSQL = "zl_סԺ���ʼ�¼_INSERT('" & .NO & "'," & int��� & "," & .����ID & "," & IIf(.��ҳID = 0, "NULL", .��ҳID) & "," & _
                            IIf(Val(.��ʶ��) = 0, "NULL", .��ʶ��) & "," & "'" & .���� & "','" & .�Ա� & "','" & .���� & "','" & .���� & "','" & .�ѱ� & "',"
                        If mblnNurseStation Then
                            gstrSQL = gstrSQL & IIf(.����ID = 0, tyPati.��������ID, .����ID) & "," & IIf(.����ID = 0, tyPati.��������ID, .����ID) & "," & .�Ӱ��־ & "," & .Ӥ���� & "," & tyPati.��������ID & ",'" & tyPati.������ & "',"
                            lng��������ID = tyPati.��������ID
                        Else
                            gstrSQL = gstrSQL & IIf(.����ID = 0, .��������ID, .����ID) & "," & IIf(.����ID = 0, .��������ID, .����ID) & "," & .�Ӱ��־ & "," & .Ӥ���� & "," & .��������ID & ",'" & .������ & "',"
                            lng��������ID = .��������ID
                        End If
                    Else
                        '�������۲��˼��������
                        gstrSQL = "Zl_������ʼ�¼_Insert('" & .NO & "'," & int��� & "," & .����ID & "," & ZVal(.��ʶ��) & "," & _
                            "'" & .���� & "','" & .�Ա� & "','" & .���� & "','" & .�ѱ� & "'," & .�Ӱ��־ & "," & .Ӥ���� & ","
                        If mblnNurseStation Then
                            gstrSQL = gstrSQL & IIf(.����ID = 0, tyPati.��������ID, .����ID) & "," & tyPati.��������ID & ",'" & tyPati.������ & "',"
                            lng��������ID = tyPati.��������ID
                        Else
                            gstrSQL = gstrSQL & IIf(.����ID = 0, .��������ID, .����ID) & "," & .��������ID & ",'" & .������ & "',"
                            lng��������ID = .��������ID
                        End If
                    End If
                End With
                
                '�շ�ϸĿ����
                With mobjBillDetail
                    '�����������
                    If .��� <> int�к� Then
                        int�к� = .���
                        
                        '���´����������
                        If mobjBill.Details(.���).�������� = 0 Then    'ֻ�д��ڸ���ʱ,�Ż���´�����
                            For i = .��� + 1 To mobjBill.Details.Count
                                If mobjBill.Details(i).�������� = .��� Then
                                    mobjBill.Details(i).�������� = int��� '������Ŀ�ж��������Ŀ(������)ʱ,ȡ��һ�����
                                End If
                            Next
                        End If
                    End If
                    gstrSQL = gstrSQL & .�������� & "," & .�շ�ϸĿID & ",'" & .�շ���� & "','" & .���㵥λ & "',"
                    
                    If tyPati.�������� <> 1 Then
                        gstrSQL = gstrSQL & IIf(.������Ŀ��, 1, 0) & "," & IIf(.���մ���ID = 0, "NULL", .���մ���ID) & ",'" & .���ձ��� & "',"
                    End If
                    
                    dbl���� = .����
                    If InStr(",5,6,7,", .�շ����) > 0 And gblnסԺ��λ Then
                        dbl���� = Format(.���� * .Detail.סԺ��װ, gstrFeePrecisionFmt)
                    End If
                    gstrSQL = gstrSQL & IIf(.���� = 0, 1, .����) & "," & dbl���� & "," & .���ӱ�־ & "," & IIf(.ִ�в���ID = 0, "NULL", .ִ�в���ID) & ","
                    
                    '�ռ����Ϸ��ϲ���,�Ա��Զ�����
                    If Not (gbytBilling = 1 Or blnSavePrice) And gint���ķ��Ͽ��� <> 0 Then
                        'gint���ķ��Ͽ���:0-���Զ����ϣ�1-�Զ����ϣ�2-�����ҿ���ʱ�Զ�����
                        If .ִ�в���ID <> 0 And .�շ���� = "4" And .Detail.�������� _
                            And ((gint���ķ��Ͽ��� = 2 And .ִ�в���ID = lng��������ID) Or gint���ķ��Ͽ��� = 1) Then
                            If InStr("," & strStuffDept, "," & .ִ�в���ID & ",") = 0 Then
                                strStuffDept = strStuffDept & "," & .ִ�в���ID
                            End If
                        End If
                    End If
                End With
                
                '������Ŀ����
                With mobjBillIncome
                    intParent = intParent + 1
                    dbl���� = .��׼����
                    If InStr(",5,6,7,", mobjBillDetail.�շ����) > 0 And gblnסԺ��λ Then
                        dbl���� = Format(.��׼���� / mobjBillDetail.Detail.סԺ��װ, gstrFeePrecisionFmt)
                    End If
                    gstrSQL = gstrSQL & IIf(intParent = 1, "Null", intParentNO + 1) & "," & .������ĿID & "," & _
                        "'" & .�վݷ�Ŀ & "'," & dbl���� & "," & .Ӧ�ս�� & "," & .ʵ�ս�� & ","
                    If tyPati.�������� <> 1 Then gstrSQL = gstrSQL & ZVal(.ͳ����) & ","
                End With
                
                If cboִ������.ListIndex < 0 Or cboִ������.Enabled = False Then
                    strTmp = "NULL,NULL"
                ElseIf cboִ������.ItemData(cboִ������.ListIndex) = 0 Then
                    strTmp = "NULL,NULL"
                Else
                    strTmp = "1," & cboִ������.ItemData(cboִ������.ListIndex)
                End If
               
                If mobjBillDetail.�շ���� = "7" Then
                    str��ҩ��̬ = "'" & mobjBillDetail.Detail.��ҩ��̬ & "'"
                Else
                    str��ҩ��̬ = "NULL"
                End If
                
                '��������
                gstrSQL = gstrSQL & "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    strAddDate & ",NULL," & IIf(gbytBilling = 1 Or blnSavePrice, 1, 0) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "',"
                If tyPati.�������� <> 1 Then
                    gstrSQL = gstrSQL & "0," & IIf(mobjBillDetail.�շ���� = "4", mlng�������ID, mlngҩƷ���ID) & "," & _
                        "NULL,'" & mobjBillDetail.ժҪ & "'," & chk����.Value & "," & ZVal(lngҽ��ID) & "," & _
                        "Null,Null,'|" & mobjBill.�巨 & "', " & strTmp & ",NULL,'" & mobjBillDetail.Detail.���� & "',0," & _
                        mobjBill.��ҩ����ID & "," & str��ҩ��̬ & ")"
                Else
                    gstrSQL = gstrSQL & "NULL,'" & mobjBillDetail.ժҪ & "'," & ZVal(lngҽ��ID) & ",Null,Null,'|" & mobjBill.�巨 & "'," & _
                        strTmp & ",1," & str��ҩ��̬ & ",0,NULL," & ZVal(mobjBill.��ҳID) & ","
                    If mblnNurseStation Then
                        gstrSQL = gstrSQL & IIf(mobjBill.����ID = 0, tyPati.��������ID, mobjBill.����ID) & ")"
                    Else
                        gstrSQL = gstrSQL & IIf(mobjBill.����ID = 0, mobjBill.��������ID, mobjBill.����ID) & ")"
                    End If
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mobjBillDetail.�շ�ϸĿID & ";" & gstrSQL
            Next
        End If
    Next

    If UBound(arrSQL) >= 0 Then
        '��SQL���а��շ�ϸĿID����
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
        For i = 0 To UBound(arrSQL)
            varData = Split(arrSQL(i), ";")
            zlAddArray cllExctPro, varData(1)
        Next
        
        'ִ���Զ�����
        If strStuffDept <> "" Then
            strStuffDept = Mid(strStuffDept, 2)
            For i = 0 To UBound(Split(strStuffDept, ","))
                strSQL = "zl_�����շ���¼_��������(" & Split(strStuffDept, ",")(i) & ",25,'" & mobjBill.NO & "','" & _
                    UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1," & strAddDate & ")"
                zlAddArray cllExctPro, strSQL
            Next
        End If
                    
                    
        'ִ��SQL���
        On Error GoTo errH
        blnTrans = True
        Call zlExecuteProcedureArrAy(cllExctPro, Me.Caption, True)
        
        '׼���Զ���ҩ(����ͨ����),�����������в��ܶ�������
        If mblnSendMateria Then
            Set rsTmp = Get����ҩ�嵥(mobjBill.NO, Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), False)
            If rsTmp.RecordCount > 0 Then
                str���ܺ� = zlDatabase.GetNextNo(20)
                For i = 0 To rsTmp.RecordCount - 1
                   strSQL = "ZL_ҩƷ�շ���¼_���ŷ�ҩ(" & rsTmp!�ⷿID & "," & rsTmp!ID & ",'" & UserInfo.���� & "'," & strAddDate & ",Null,Null,Null," & str���ܺ� & ")"
                    zlAddArray cllSMSQL, strSQL
                    rsTmp.MoveNext
                Next
            End If
            rsTmp.Close
        End If
        
        'ִ���Զ���ҩ
        zlExecuteProcedureArrAy cllSMSQL, Me.Caption, True, True
        
        '����ʵʱ�ϴ�
        If gbytBilling = 0 And Not blnSavePrice And tyPati.���� <> 0 Then
            'ҽ�����������ϸ
            If MCPAR.�����ϴ� And Not MCPAR.������ɺ��ϴ� Then
                str��Ϣ = ""
                If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str��Ϣ, , tyPati.����) Then
                    gcnOracle.RollbackTrans
                    If str��Ϣ <> "" Then MsgBox str��Ϣ, vbInformation, gstrSysName
                    strMessage = str��Ϣ
                    Exit Function
                End If
            End If
        End If
        gcnOracle.CommitTrans
        blnTrans = False
        '2.���ʺ�ʵʱ�ϴ�
        If gbytBilling = 0 And Not blnSavePrice And tyPati.���� <> 0 Then
            'ҽ�����������ϸ
            If MCPAR.�����ϴ� And MCPAR.������ɺ��ϴ� Then
                str��Ϣ = ""
                If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str��Ϣ, , tyPati.����) Then
                    If str��Ϣ <> "" Then
                        MsgBox str��Ϣ, vbInformation, gstrSysName
                    Else
                        MsgBox "����""" & mobjBill.NO & """��������ҽ������ʧ��,�õ����ѱ��棡", vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
        
        '���뵥����ʷ��¼(�������͵���)
        For i = 0 To cboNO.ListCount - 1
            strNo = strNo & "," & cboNO.List(i)
        Next
        strNo = mobjBill.NO & strNo
        cboNO.Clear
        For i = 0 To UBound(Split(strNo, ","))
            cboNO.AddItem Split(strNo, ",")(i)
            If i = 9 Then Exit For 'ֻ��ʾ10��
        Next
        'ҽ���ӿ�
        If str��Ϣ <> "" Then MsgBox str��Ϣ, vbInformation, gstrSysName
    End If
    Call zlChargeSaveAfter_Plugin(mlngModule, mobjBill.����ID, mobjBill.��ҳID, False, 2, mobjBill.NO)
    SaveBill = True
    Exit Function
ErrHand:
    strMessage = Err.Description
    If ErrCenter = 1 Then Resume
    Exit Function
errH:
    If Err.Description Like "*��ǰ���㵥�۲�һ��*" Then
       If blnTrans Then gcnOracle.RollbackTrans
       If MsgBox("ĳЩ����ҩƷ�۸��ѷ����仯��Ҫ�Զ�����۸���", vbYesNo + vbQuestion + vbDefaultButton1, App.ProductName) = vbYes Then
           Call CalcMoneys(tyPati)
           Call ShowDetails
           Call ShowMoney
           If InStr(Err.Description, "[ZLSOFT]") > 0 Then
                strMessage = Split(Err.Description, "[ZLSOFT]")(1)
           Else
                strMessage = Err.Description
           End If
           Exit Function
       End If
    Else
        If blnTrans Then gcnOracle.RollbackTrans
        If InStr(Err.Description, "[ZLSOFT]") > 0 Then
             strMessage = Split(Err.Description, "[ZLSOFT]")(1)
        Else
             strMessage = Err.Description
        End If
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Function

Private Sub Auto��������()

    Dim i As Integer, strSQL As String, rsTmp As ADODB.Recordset
    Dim blnFind As Boolean
    
    If mstr����IDs <> "" Then
        strSQL = "Select ��ǰ����ID From ������Ϣ Where ����ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Split(mstr����IDs, ",")(0)))
    Else
        strSQL = "Select ��ǰ����ID From ������Ϣ Where ����ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    End If
    
    blnFind = False
    If Not rsTmp.EOF Then
        For i = 1 To cbo��������.ListCount
            If Val(cbo��������.ItemData(i - 1)) = Val(Nvl(rsTmp!��ǰ����id)) Then
                blnFind = True
                cbo��������.ListIndex = i - 1
                Exit For
            End If
        Next i
    End If
    If blnFind = False Then cbo��������.ListIndex = 0
End Sub


Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-07-08 17:33:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dtdtCurDate As Date     '��������ǰʱ��
    On Error GoTo errH
    
    '��ȡ��ҩ������
    Call ReadABCNum(mstrPrivsOpt)
    
    '��ͬҩ��ҩƷ�����鷽ʽ
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    

    '------------------������ȡ------------------
    strSQL = " Select '����ְ��' As ����,count(ҩ��ID) As num From ҩƷ���� Where ����ְ��<>'00' Union All " & _
             " Select '��������' As ����,count(ҩ��ID) As num From ҩƷ���� Where ��������>0    "
    
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    rsTmp.Filter = "����='����ְ��'"
    If Not rsTmp.EOF Then mbln����ְ���� = (rsTmp!Num > 0)
    
    rsTmp.Filter = "����='��������'"
    If Not rsTmp.EOF Then mbln����������� = (rsTmp!Num > 0)

    
    '------------------������ȡ------------------
            
    If Init�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mstrPrivs, 0, mlngDeptID) = False Then Exit Function
    
    If mblnNurseStation Then
        If Init�����˿�������(cbo������ѡ��, cbo��������ѡ��, mrs������, mrs��������, mstrPrivs, 0, mlngDeptID) = False Then Exit Function
    End If
    
    If gstr�շ���� = "" Then
        strSQL = "Select ����,���� as ��� from �շ���Ŀ��� Where ����<>'1' Order by ���"
    Else
        strSQL = "" & _
        "   Select /*+ RULE */   A.����,A.���� as ��� " & _
        "   From �շ���Ŀ��� A," & _
        "          (Select Column_Value From Table(Cast(f_str2list([1]) As Zltools.t_strlist))) J " & _
        "   Where A.����=J. Column_Value " & _
        "   Order by ���"
    End If
    Set mrsClass = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(gstr�շ����, "'", ""))
    
    If mrsClass.EOF Then
        MsgBox "û�����ÿ��õ��շ����,�����ڱ��ز��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��ֻ��һ�ֿ�ѡ�շ����ʱ,�����û�ѡ��
    mblnOne = (mrsClass.RecordCount = 1)
    
    If InStr(gstr�շ����, "'5'") > 0 Or InStr(gstr�շ����, "'6'") > 0 _
        Or InStr(gstr�շ����, "'7'") > 0 Or gstr�շ���� = "" Then
        mlngҩƷ���ID = ExistIOClass(9)
        If mlngҩƷ���ID = 0 Then
            MsgBox "����ȷ���������ݵ�������,���ȵ����������������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If InStr(gstr�շ����, "'4'") > 0 Or gstr�շ���� = "" Then
        mlng�������ID = ExistIOClass(41)
        If mlng�������ID = 0 Then
            MsgBox "����ȷ�����ĵ��ݵ�������,���ȵ����������������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    'ִ�в���
    strSQL = _
    "Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.������� " & _
    " From ���ű� A,��������˵�� B " & _
    " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
    " And B.����ID=A.ID and B.������� IN(2,3) " & _
    " Order by B.�������,A.����"
    Set mrsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If mrsUnit.EOF Then
        MsgBox "û�г�ʼ��������Ϣ,�����޷�����ִ�в��š����ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    dtdtCurDate = zlDatabase.Currentdate
    txtDate.Text = Format(dtdtCurDate, "yyyy-MM-dd HH:mm:ss")
    
    '�Զ�ʶ��Ӱ�
    If OverTime(dtdtCurDate) Then chk�Ӱ�.Value = Checked
    Set mrsWarn = GetUnitWarn
 
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function GetDeptName(ByVal lngDeptID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '���:lngDeptID-����ID(����ID)
    '����:����ȡ��������
    '����:���˺�
    '����:2015-07-15 17:52:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select ���� From ���ű� where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngDeptID)
    If rsTemp.EOF = False Then GetDeptName = Nvl(rsTemp!����)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ�Ҫ��ɵĹ������ý��沼��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-07-08 16:07:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrHead() As String, i As Long
    With cboִ������
        .Clear
        .AddItem "����"
        .ItemData(.NewIndex) = 0: .ListIndex = .NewIndex
        .AddItem "��Ժ��ҩ"
        .ItemData(.NewIndex) = 3
        .AddItem "��ȡҩ"
        .ItemData(.NewIndex) = 4
    End With
                
    Call InitReportColumn
            
    '��ʼ�����
    arrHead = Split(STR_HEAD, ";")
    With Bill
        .Font.Size = 10.5
        .cboObj.Font.Size = 10.5
        
        .Cols = UBound(arrHead) + 1
        .MsfObj.FixedCols = 1
        .MsfObj.ScrollBars = flexScrollBarVertical
        .LocateCol = BillCol.��Ŀ
        .PrimaryCol = BillCol.��Ŀ
        .MsfObj.ColAlignmentFixed(0) = 4
        .TextMatrix(1, BillCol.��) = 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
                
        .ColData(BillCol.��) = BillColType.UnFocus
        
        .ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)
        If mblnOne Then .ColData(BillCol.���) = BillColType.UnFocus
        If .ColData(BillCol.���) <> BillColType.UnFocus Then
            .LocateCol = BillCol.���
        End If
        .ColData(BillCol.��Ŀ) = BillColType.CommandButton  '��Ŀ����,��Ť��ѡ
        .ColData(BillCol.����) = BillColType.Text '��/������
        .ColData(BillCol.��Ʒ��) = BillColType.UnFocus    '��Ʒ������
        .ColData(BillCol.���) = BillColType.UnFocus    '�������
        .ColData(BillCol.��λ) = BillColType.UnFocus  '��λ����
        .ColData(BillCol.����) = BillColType.UnFocus  '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
        .ColData(BillCol.����) = BillColType.UnFocus '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
        .ColData(BillCol.Ӧ�ս��) = BillColType.UnFocus  'Ӧ�ս������
'        .ColData(BillCol.�����ϼ�) = BillColType.UnFocus   '�����ϼ�����
'        .ColData(BillCol.ʵ�պϼ�) = BillColType.UnFocus   'ʵ�պϼ�����
'        .ColData(BillCol.Ӧ�պϼ�) = BillColType.UnFocus   'Ӧ�պϼ�����
         .ColData(BillCol.ִ�п���) = BillColType.ComboBox 'Ĭ��ȡ�������һ���һ����
        .ColData(BillCol.��־) = BillColType.UnFocus '��־ȱʡ����,��Ϊ����ʱ,��Ϊ��ѡ(-1)
        .ColData(BillCol.����) = BillColType.UnFocus  '����ȱʡ����
          
        .SetColColor BillCol.���, &HE7CFBA
        .SetColColor BillCol.��Ŀ, &HE7CFBA
        .SetColColor BillCol.����, &HE7CFBA
        .SetColColor BillCol.ִ�п���, &HE7CFBA
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.��־, &HE0E0E0
        .MsfObj.ScrollBars = 3
        
        ReDim marrColData(.Cols - 1)
        For i = 0 To .Cols - 1
            marrColData(i) = .ColData(i)
        Next
    End With
    
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name)
    If gTy_System_Para.bytҩƷ������ʾ <> 2 Then
        '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
        Bill.ColWidth(BillCol.��Ʒ��) = 0
    Else
        If Bill.ColWidth(BillCol.��Ʒ��) = 0 Then
             Bill.ColWidth(BillCol.��Ʒ��) = GetOrigColWidth(BillCol.��Ʒ��)
        End If
    End If
    
    Call SetMoneyList '��ʼ�������б�
     
    '��ȡ����ƥ�䷽ʽ
    sta.Panels("MedicareType").Visible = True
    sta.Panels("PY").Visible = gbln�����л� '35242
    sta.Panels("WB").Visible = gbln�����л�
    '����ƥ�䷽ʽ��0-ƴ��,1-���,2-����
    If gbytCode = 0 Then
        sta.Panels("PY").Bevel = sbrInset
        sta.Panels("WB").Bevel = sbrRaised
    ElseIf gbytCode = 1 Then
        sta.Panels("PY").Bevel = sbrRaised
        sta.Panels("WB").Bevel = sbrInset
    Else
        sta.Panels("PY").Bevel = sbrInset
        sta.Panels("WB").Bevel = sbrInset
    End If
    txtӦ��.Text = gstrDec ': txtʵ��.Text = gstrDec
 
    Call SetShowCol ' ���ø�����
    
    '��ͨ���ʺͿ��ҷ�ɢ���ʻ򻮼�ʱ,�������޸Ĳ���������������ҩ�䷽
    cboִ������.Visible = True: lblִ������.Visible = True
    cmdSelWholeSet.Visible = True
    cmdSaveWholeSet.Visible = True
    
    '�������������뿪����λ��
    If gblnFromDr Then
        Call ExChangeLocate(cbo��������, cbo������)
        Call ExChangeLocate(lbl��������, lbl������)
        cbo��������.TabStop = False
    End If
    
    If mblnNurseStation Then
        lbl��������.Visible = False
        cbo��������.Visible = False
        lbl������.Visible = False
        cbo������.Visible = False
        lblִ������.Visible = False
        cboִ������.Visible = False
        lblDrawDrugDept.Visible = False
        cboDrawDept.Visible = False
    End If
End Sub

Private Sub rptPati_SelectionChanged()
    If cbo��������ѡ��.Visible = True Then cbo��������ѡ��.Visible = False
    If cbo������ѡ��.Visible Then cbo������ѡ��.Visible = False
End Sub

 Private Sub SetMoneyList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ������Ŀ�����������п�
    '����:���˺�
    '����:2015-07-08 17:57:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngW As Long, i As Long
    
    With vsMoney
        .Clear
        .Cols = 2: .Rows = 2
        .TextMatrix(0, 0) = "��Ŀ"
        .TextMatrix(0, 1) = "���"
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, 0)
            .FixedAlignment(i) = 4
        Next
        lngW = .Width - 60
        If .Rows > .Height / .RowHeight(0) Then lngW = lngW - 250
        .ColWidth(0) = lngW * 0.5: .ColWidth(1) = lngW * 0.5
        .ColAlignment(0) = 1: .ColAlignment(1) = 7
        .Row = 1
    End With
End Sub
Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Select Case Panel.Key
    Case "PY", "WB"
        If Panel.Bevel = sbrRaised And gbln�����л� Then
            '�л����������ƥ�䷽ʽ
            Panel.Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            If Panel.Key = "PY" Then
                sta.Panels("WB").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            Else
                sta.Panels("PY").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            End If
            zlDatabase.SetPara "���뷽ʽ", IIf(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIf(sta.Panels("WB").Bevel = sbrInset, 1, 0))
            gbytCode = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))
        End If
    Case "Drugstore"
        With frmSetExpence
            .mlngModul = mlngModule
            .mstrPrivs = mstrPrivs
            '���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����,2-ҽ�����Ҽ���
            '           0:��ͨ����,1-���ҷ�ɢ����,2-ҽ�����Ҽ���
            .mbytInFun = 0
            .mbytUseType = 0
            .mblnOnlyDrugStock = True
            .Show 1, Me
        End With
    End Select
End Sub
 
Private Sub SetShowCol()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����еĿ���(���ʱչ��)
    '����:���˺�
    '����:2015-07-08 18:04:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mrsClass.Filter = "����='7'"
    If mrsClass.RecordCount = 0 Then
        Bill.ColWidth(BillCol.����) = 0
    ElseIf Bill.ColWidth(BillCol.����) = 0 Then
        Bill.ColWidth(BillCol.����) = 520
    End If
End Sub
Private Function GetOrigColWidth(ByVal intIdx As Integer) As Long
    '���ܣ���ȡָ���е�ԭʼ�п�
    GetOrigColWidth = Val(Split(Split(STR_HEAD, ";")(intIdx), ",")(1))
End Function

Private Sub cboDrawDept_Click()
    Dim lng��ҩ����ID As Long
    If cboDrawDept.ListIndex <> -1 Then lng��ҩ����ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
    If Not mobjBill Is Nothing Then
        If mobjBill.��ҩ����ID = lng��ҩ����ID Then Exit Sub
        mobjBill.��ҩ����ID = lng��ҩ����ID
    End If
End Sub

Private Sub cboDrawDept_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 And Not cboDrawDept.Locked Then
        lngIdx = zlControl.CboMatchIndex(cboDrawDept.hWnd, KeyAscii)
        If lngIdx = -1 And cboDrawDept.ListCount > 0 Then lngIdx = 0
        cboDrawDept.ListIndex = lngIdx
    ElseIf KeyAscii = 13 Then
        If cboDrawDept.ListIndex = -1 Then Beep: Exit Sub
        mobjBill.��ҩ����ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo��������_GotFocus()
    zlControl.TxtSelAll cbo��������
End Sub

Private Sub cbo��������_LostFocus()
    cbo��������.SelLength = 0
End Sub

Private Sub cbo��������_Validate(Cancel As Boolean)
    If cbo��������.Text <> "" And cbo��������.ListIndex < 0 Then cbo��������.Text = ""
End Sub

Private Sub cboִ������_Click()
    If mobjBill Is Nothing Then Exit Sub
    mobjBill.ִ������ = cboִ������.ItemData(cboִ������.ListIndex)
End Sub

Private Sub cboִ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdSaveWholeSet_Click()
    Dim i As Long, strItems As String, lngִ�п���ID As Long
    Dim rsTemp As ADODB.Recordset, dbl���� As Double, dbl�۸� As Double
    Dim strSQL As String
    
    '����Ϊ�����շ���Ŀ
    '����:27327
    
    Err = 0: On Error Resume Next
    If mobjBaseItem Is Nothing Then
        Set mobjBaseItem = CreateObject("zl9BaseItem.clsBaseItem")
    End If
    If mobjBaseItem Is Nothing Then Exit Sub
    'OpenEditWholeSetItem(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection,
    '      ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPrivs As String, ByVal strItems As String) As Boolean
    'strItems:���,����,�շ�ϸĿID,����,����,ִ�п���|���,����,�շ�ϸĿID,����,����,ִ�п���|��
    Err = 0: On Error GoTo ErrHand:

    With mobjBill
        strItems = ""
        For i = 1 To .Details.Count
             '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
            If InStr(1, ",4,5,6,7,", "," & .Details(i).Detail.���) > 0 Then
                lngִ�п���ID = 0
            ElseIf InStr(1, ",0,4", .Details(i).Detail.ִ�п���) > 0 Then
                lngִ�п���ID = .Details(i).ִ�в���ID
            Else
                lngִ�п���ID = 0
            End If
            '����:52349
            dbl���� = .Details(i).����
            dbl�۸� = IIf(.Details(i).Detail.���, .Details(i).InComes(1).��׼����, 0)
            If InStr(",5,6,7,", .Details(i).�շ����) > 0 And gblnסԺ��λ Then
                dbl���� = Format(dbl���� * .Details(i).Detail.סԺ��װ, gstrFeePrecisionFmt)
                dbl�۸� = Format(dbl�۸� / .Details(i).Detail.סԺ��װ, gstrFeePrecisionFmt)
            End If
            strItems = strItems & "|" & .Details(i).��� & "," & .Details(i).�������� & "," & .Details(i).�շ�ϸĿID & "," & .Details(i).���� & "," & dbl���� & "," & dbl�۸� & "," & lngִ�п���ID
         Next
         If strItems = "" Then
            MsgBox "����δ�����κ���Ϣ,���ܱ���Ϊ�����շ���Ŀ,����!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
            Exit Sub
        End If
        strItems = Mid(strItems, 2)
    End With
    Call mobjBaseItem.OpenEditWholeSetItem(Me, gcnOracle, glngSys, 1150, mstrPrivsOpt, strItems)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSelWholeSet_Click()
    'ѡ������Ŀ
    Dim rsSel As ADODB.Recordset, lng����ID As Long, lng��������ID As Long
    Dim tmpBill As New ExpenseBill, bytӤ���� As Byte, dtCurdate As Date
    Dim curTotal  As Currency, rsTmp As ADODB.Recordset, i As Long
    Dim intInsure As Integer
    Dim bln��ҩ As Boolean
    
    intInsure = 0
    If mlngSelPatiCount = 0 Then
        MsgBox "����ѡ����,����!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
            
    If mobjBill Is Nothing Then
        lng����ID = 0
        If cbo��������.ListIndex < 0 Then
            lng��������ID = 0
        Else
            lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
    Else
        lng����ID = mobjBill.����ID: lng��������ID = mobjBill.��������ID
    End If
    
    If mlngSelPatiCount = 0 Then
        MsgBox "����ѡ����,����!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    
    If zlSelectWholeItems(Me, mlngModule, mstrPrivsOpt, rsSel) = False Then Exit Sub
    If rsSel Is Nothing Then Exit Sub
    Err = 0: On Error GoTo ErrHand:
    Screen.MousePointer = 11
    
    Set tmpBill = ImportWholeSet(Me, intInsure, rsSel, lng����ID, gblnסԺ��λ, lng��������ID, bytӤ����, 2, chk�Ӱ�.Value = 1, _
        0, 2, UserInfo.����, zlStr.NeedName(cbo������.Text), , mblnNurseStation, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
        
    '��������
    '�������Ĳ�����Ϣ
    Set mobjBill = New ExpenseBill
    Set mobjBill = tmpBill
    bln��ҩ = False
    With mobjBill
        For i = 1 To .Details.Count - 1
            If .Details(i).�շ���� = "7" Then bln��ҩ = True: Exit For
        Next
    End With
    
    dtCurdate = zlDatabase.Currentdate
    mobjBill.NO = cboNO.Text
    mobjBill.�Ǽ�ʱ�� = dtCurdate
    mobjBill.����Ա��� = UserInfo.���
    mobjBill.����Ա���� = UserInfo.����
    mobjBill.�Ӱ��־ = chk�Ӱ�.Value
    txtDate.Text = Format(dtCurdate, "yyyy-MM-dd HH:mm:ss")
    
    
    Bill.Redraw = False
    Bill.ClearBill
    '�����:116774,����,2017/12/28,�������ʵ���ȫ��ҩƷ�ĳ�����Ŀ��,������ݻᱨ��
    Bill.Rows = IIf(mobjBill.Details.Count = 0, 2, mobjBill.Details.Count + 1)
    
    Call InitBillColumnColor
    '���ʷ��౨��
    mstrWarn = ""
        
    Call Set�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mobjBill.������, mobjBill.��������ID)
        
    '������Ķ����˺�ȷ���ѱ��,�ټ���۸�
    Dim tyPati As TY_PATIINFOR
    Call CalcMoneys(tyPati)   '��CalcMoneys�в�����ʵ�ս��(����Ҫ����յ�tyPati)
    
    Call ShowDetails
    Call ShowMoney
    With Bill
        For i = 1 To .Rows - 1
            .TextMatrix(i, BillCol.��) = i
        Next
    End With
    Bill.Redraw = True
    Call SetDrawDrugDeptEnabled
    Screen.MousePointer = 0
    Exit Sub
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ReSetDefaultִ�п���(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������ȱʡ��ִ�п���
    '����:���˺�
    '����:2015-07-09 10:10:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng���˿���ID As Long, lngDoUnit As Long, strҩ��IDs As String
    
    Dim dblStock As Double
    Err = 0: On Error GoTo ErrHand:
    With mobjBill.Details(lngRow)
         '���ĺ�ҩƷ����
        '���˿���ID
        lng���˿���ID = mobjBill.����ID
        If cbo��������.Visible Then
            If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
        Else
            If lng���˿���ID = 0 Then lng���˿���ID = mlngDeptID
            If lng���˿���ID = 0 Then lng���˿���ID = GetNurseStationFirstPatiDeptID '��ʿ����վ,ȡ��һ�����˿���ID
            If lng���˿���ID = 0 Then lng���˿���ID = mlng����ID
        End If
        
         '����ִ�п���ȱʡΪ���˲���,�������ָ����,��Ϊָ������
        If .Detail.��� = "4" Then
            lngDoUnit = IIf(glng���ϲ��� > 0, glng���ϲ���, mobjBill.����ID)
            If lngDoUnit = 0 Then lngDoUnit = IIf(cbo��������.Visible, Get��������ID, lng���˿���ID)
        End If
        
        lngDoUnit = Get�շ�ִ�п���ID(.Detail.���, .Detail.ID, _
             .Detail.ִ�п���, lng���˿���ID, IIf(cbo��������.Visible, Get��������ID, lng���˿���ID), 2, lngDoUnit, mobjBill.����ID, .ִ�в���ID)
       .ִ�в���ID = lngDoUnit
        
        If InStr(",5,6,7,", .Detail.���) > 0 Then
            '��ǰ��ҩƷ���
            If Not gbln���뷢ҩ Then
                dblStock = GetStock(.Detail.ID, lngDoUnit)
                If gblnסԺ��λ Then
                    dblStock = dblStock / .Detail.סԺ��װ
                End If
                  .Detail.��� = dblStock
                Call ShowStock(.Detail.����, .Detail.���)
            Else
                strҩ��IDs = Decode(.Detail.���, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
                If strҩ��IDs <> "" Then
                    dblStock = GetMultiStock(.Detail.ID, strҩ��IDs)
                    If gblnסԺ��λ Then
                        dblStock = dblStock / .Detail.סԺ��װ
                    End If
                    .Detail.��� = dblStock
                    Call ShowStock(.Detail.����, .Detail.���)
                End If
            End If
        ElseIf .Detail.��� = "4" And .Detail.�������� Then
            dblStock = GetStock(.Detail.ID, lngDoUnit)
            .Detail.��� = dblStock
            Call ShowStock(.Detail.����, .Detail.���)
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
 End Sub
 
Private Sub ShowStock(strҩƷ As String, dbl��� As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾҩƷ�����ĵĿ��
    '����:���˺�
    '����:2015-07-09 10:51:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
     On Error GoTo errHandle
    If InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0 Then
        sta.Panels(Pan.C2��ʾ��Ϣ).Text = "[" & strҩƷ & "]���ÿ��:" & dbl���
    Else
        sta.Panels(Pan.C2��ʾ��Ϣ).Text = "[" & strҩƷ & "]" & IIf(dbl��� > 0, "��", "��") & "���."
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Bill_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long, i As Long
    
    If KeyCode <> vbKeyReturn Then Exit Sub
 '   If Bill.cboStyle = DropOlnyDown Then Exit Sub
    
    
    '���
    If Bill.TextMatrix(0, Bill.Col) = "���" Then
        If Bill.ListIndex > 0 Then Exit Sub
        For i = 0 To Bill.cboObj.ListCount - 1
            If IsNumeric(Bill.CboText) Then
                If Split(Bill.cboObj.List(i), "-")(0) = Val(Bill.CboText) Then
                    Bill.ListIndex = i
                    Exit Sub
                     
                End If
            ElseIf zlCommFun.IsCharAlpha(Bill.CboText) Then
                If zlCommFun.SpellCode(Split(Bill.List(i) & "-", "-")(1)) Like UCase(Bill.CboText) & "*" Then
                    Bill.ListIndex = i
                    Exit Sub
                End If
            ElseIf Split(Bill.ItemData(i) & "-", "-")(1) Like "*" & UCase(Bill.CboText) & "*" Then
                Bill.ListIndex = i
                Exit Sub
            End If
        Next
        Exit Sub
    End If
    
    If Bill.TextMatrix(0, Bill.Col) <> "ִ�п���" Then Exit Sub
    If Bill.ListIndex <> -1 Then Exit Sub
    
    lngRow = Bill.Row
    If mobjBill.Details.Count < lngRow Then Exit Sub
    
    With mobjBill.Details(lngRow)
        If InStr(",4,5,6,7,", .�շ����) > 0 Then
            If mrsWork Is Nothing Then Exit Sub
            If mrsWork.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModule, Bill.cboObj, mrsWork, Bill.CboText, True, , False) = False Then Exit Sub
        Else
            If mrsUnit Is Nothing Then Exit Sub
            If mrsUnit.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModule, Bill.cboObj, mrsUnit, Bill.CboText, True, , False) = False Then Exit Sub
        End If
    End With
    Exit Sub
End Sub
Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytsubs As Byte
    Dim bln��������ۿ� As Boolean
    Dim lngMainRow As Long
    
    If mobjBill.Details.Count >= Row Then
        '��������Ŀ����ɾ��ȷ��
        For i = Row + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�������� = Row Then bytsubs = bytsubs + 1
        Next
        If bytsubs > 0 Then
            If MsgBox("����Ŀ���� " & bytsubs & " ��������Ŀ,ɾ������ĿҲ��ɾ�����Ĵ�����Ŀ,������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        ElseIf mobjBill.Details(Row).�������� <> 0 Then '������Ŀɾ��ȷ��
            If MsgBox("����Ŀ��[" & mobjBill.Details(mobjBill.Details(Row).��������).Detail.���� & "]�Ĵ�����Ŀ,ȷ��Ҫɾ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            Else
                bln��������ۿ� = gbln��������ۿ�
            End If
        ElseIf MsgBox("ȷʵҪɾ�����շ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
        
        If bln��������ۿ� Then lngMainRow = mobjBill.Details(Bill.Row).�������� '����Ǵ���,ɾ��֮ǰ���´���Ĵ�������,���������,����ɾ��,��������
        

        
        'ɾ������
        For i = mobjBill.Details.Count To Row + 1 Step -1
            If mobjBill.Details(i).�������� = Row Then
                Call DeleteDetail(i) '��˳��ɾ���������
            End If
        Next
        Call DeleteDetail(Row) 'ɾ������
        Call ShowDetails
        Call ShowMoney
                
        Bill.TxtVisible = False
        Bill.CmdVisible = False
        Bill.CboVisible = False
        
        Cancel = True '���ÿؼ�������ɾ��
        
        mlngPreRow = 0    '��ʾ�иı���
        Call Bill_EnterCell(Bill.Row, Bill.Col)
        Call SetDrawDrugDeptEnabled
    ElseIf Row = 1 Then
        For i = 1 To Bill.Cols - 1
            Bill.TextMatrix(Row, i) = ""
        Next
        Cancel = True
    End If
    Call SetColNum(Row)
End Sub

Private Sub Bill_cboClick(ListIndex As Long)
    Dim dblStock As Double, tyPati As TY_PATIINFOR
    Dim lngִ�п��� As Long, strִ�п��� As String
    If mblncboClick Then Exit Sub  '����ͬһ������������bill��ֵѭ������,ע�����κ�exit sub ֮ǰ����mblncboClick = False
    'ҩƷ�����
    If Not (ListIndex <> -1 And Bill.TextMatrix(0, Bill.Col) = "ִ�п���") Then Exit Sub
    If mobjBill.Details.Count < Bill.Row Then Exit Sub
    
    mblncboClick = True
    
    With mobjBill.Details(Bill.Row)
        If .ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
            lngִ�п��� = .ִ�в���ID: strִ�п��� = Bill.TextMatrix(Bill.Row, Bill.Col)
            .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
            
            Bill.TextMatrix(Bill.Row, Bill.Col) = Bill.CboText
            
            If InStr(",5,6,7,", .�շ����) > 0 Then
                'ȡ���
                dblStock = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                If gblnסԺ��λ Then
                    dblStock = dblStock / .Detail.סԺ��װ
                End If
                .Detail.��� = dblStock  '��¼��ǰ��ҩƷ���
                Call ShowStock(.Detail.����, .Detail.���)
                
                'ҩ���ı�,ʵ��ҩƷ���¼���۸�
                Call CalcMoneys(tyPati, Bill.Row)  'ʵ�ս��������,����typati����Ϊ��
                Call ShowDetails(Bill.Row)
                Call ShowMoney
                
            ElseIf .�շ���� = "4" And .Detail.�������� Then
                'ȡ���
                dblStock = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                .Detail.��� = dblStock
                Call ShowStock(.Detail.����, .Detail.���)
                
                '���ϲ��Ÿı�,ʱ���������¼���۸�
                If .Detail.��� Then
                    Call CalcMoneys(tyPati, Bill.Row)
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                End If
            ElseIf InStr(",4,5,6,7,", .�շ����) = 0 Then
                If CheckItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row) '������ڴ���,��ı��ҩƷ�е�ִ�п���
            End If
            If mobjBill.Details(Bill.Row).���� <> 0 Then
                If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                    MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                    Bill.Text = "": Bill.TxtVisible = False
                    Bill.cboObj.Text = strִ�п���: .ִ�в���ID = lngִ�п���
                    mblncboClick = False: Exit Sub
                End If
            End If
        End If
    End With
    mblncboClick = False
End Sub


Private Sub Bill_CellCheck(Row As Long, Col As Long)
    '˵��������ȫ��Ϊ��Ҫ����,������ȫ��Ϊ��������
    Dim i As Long, strCheck As String, bytTime As Byte
    Dim blnReSet As Boolean, tyPati As TY_PATIINFOR
    
    If Bill.TextMatrix(Row, BillCol.��Ŀ) = "" Then Bill.TextMatrix(Row, Col) = "": Exit Sub
    
    '������δ��������Ч
    If mobjBill.Details.Count < Row Then
        Bill.TextMatrix(Row, Col) = "": Exit Sub
    End If
    
    strCheck = Bill.TextMatrix(Row, Col)
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�շ���� = "F" And mobjBill.Details(i).���ӱ�־ = 0 And i <> Row Then bytTime = bytTime + 1
    Next
    
    blnReSet = bytTime > 0
    If blnReSet = False Then     '����ֻ���ڸ����������ָĳ���������,��Ҫ���¼ƴ���:25495
        blnReSet = (strCheck = "" And mobjBill.Details(Row).�շ���� = "F" And mobjBill.Details(Row).���ӱ�־ = 1)
    End If
    
    If blnReSet Then
        With mobjBill.Details(Row)
            .���ӱ�־ = IIf(strCheck = "", 0, 1)
            Call CalcMoneys(tyPati, Row)   'ʵ�ս��������,����typati����Ϊ��
            
            Call ShowDetails(Row)
        End With
        Call ShowMoney
    ElseIf strCheck <> "" Then
        Bill.TextMatrix(Row, Col) = ""
        MsgBox "�����б�Ȼ��һ���������Ǹ���������", vbInformation, gstrSysName
        Exit Sub
    End If
End Sub
Private Sub Bill_CommandClick()
    Dim lng��Ŀid As Long, blnCancel As Boolean, bln��ʿ As Boolean
    Dim str��� As String, str��׼��Ŀ As String
    Dim int������Դ As Integer, int���� As Integer
    Dim str�ų���� As String
    
    Call GetOperatorInfo(mrs������, mobjBill.������, bln��ʿ)
    If gbln�շ���� Then
        If Bill.RowData(Bill.Row) <> 0 Then
            str��� = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
        Else
            str��� = IIf(bln��ʿ, "'E','M','4'", gstr�շ����)
        End If
    Else
        str��� = IIf(bln��ʿ, "'E','M','4'", gstr�շ����)
    End If
    int������Դ = 2
    
    If zlCheckBill���ڷ�ɢװ��ҩ() = True Then mblnSelect = False: Exit Sub
    
    lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, int������Դ, int����, gblnסԺ��λ, str���, , , str��׼��Ŀ, _
        0, str�ų����, , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    If lng��Ŀid = 0 Then mblnSelect = False: Exit Sub
    
    Bill.Text = lng��Ŀid
    mblnSelect = True
    Call Bill_KeyDown(13, 0, blnCancel)
    Bill.SetFocus
    If Not blnCancel Then
        Bill.Text = "": Bill.TxtVisible = False
        Call zlCommFun.PressKey(13)
    End If
End Sub



Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    '���ܣ�����������
    Dim rsҩƷ��Ϣ As ADODB.Recordset
    Dim lng��Ŀid As Long, str��� As String, bln��ʿ As Boolean
    Dim str��׼��Ŀ As String, int������Դ As Integer, lng���˿���ID As Long, int���� As Integer
    Dim dblStock As Double, strScope As String, i As Long
    Dim dblPreTime As Double, dblPreMoney As Double, dblNum As Double, lngOld���� As Long
    Dim blnSkip As Boolean, curTotal As Currency
    Dim lngDoUnit As Long, strժҪ As String, blnInput As Boolean
    Dim strҩ��IDs As String, bln�������� As Boolean, cur��� As Currency
    Dim curItemMoney As Currency
    Dim colStock As Collection, str�ų���� As String
    Dim cllData As Collection, tyPati As TY_PATIINFOR
    
    On Error GoTo errH
    '�����:110693,����,2017/08/07,��ȡִ�п��Ҳ���ȷ
    mobjBill.����ID = mlng����ID
    If Not (KeyCode = 13 And Bill.Active) Then Exit Sub
    
    If Bill.ColData(Bill.Col) = BillColType.Text_UnModify Then Exit Sub
                        
    Select Case Bill.TextMatrix(0, Bill.Col)
        Case "���"
            If Bill.ListIndex = -1 Then Exit Sub
            '���������ʱ���ᶨλ�������
            If Bill.RowData(Bill.Row) <> Bill.ItemData(Bill.ListIndex) Then
                'һ���ĸ��շ����,�����(����)ԭ�и���Ŀ����
                For i = 2 To Bill.Cols - 1
                    Bill.TextMatrix(Bill.Row, i) = ""
                Next
                If mobjBill.Details.Count >= Bill.Row Then
                    Set mobjBill.Details(Bill.Row).Detail = New Detail
                    Set mobjBill.Details(Bill.Row).InComes = New BillInComes
                    With mobjBill.Details(Bill.Row)
                        .�շ�ϸĿID = 0: .�շ���� = ""
                    End With
                    Call CalcMoneys(tyPati) 'tyPati�����ʱ,ʵ�ս��������
                    Call ShowMoney
                End If
            End If
            Bill.RowData(Bill.Row) = Bill.ItemData(Bill.ListIndex) '��ʱ��RowData��¼��ѡ����շ����
        Case "��Ŀ"
            '����Ŀȷ��,���շ�ϸĿ��Ӧ�ĳ�����������,ͬʱ���ﴦ���շѴ�����Ŀ
            If Bill.Text <> "" Then
                '��������������Ŀ�ϰ��س�,��ѡ����ѡ��
                If mobjBill.Details.Count >= Bill.Row Then
                    'ͨ����ťѡ���Ƿ��ص�ID,�����������ı�,�����һ����,�򲻸ı�
                    If Bill.TextMatrix(Bill.Row, BillCol.��Ŀ) = Bill.Text Then
                        Bill.TxtVisible = False: Bill.CmdVisible = False: Exit Sub
                    End If
                End If
            
                sta.Panels(2).Text = "": sta.Panels("MedicareType").Text = ""
                blnInput = True
                If mblnSelect Then
                    mblnSelect = False '��������ñ�־
                    Set mobjDetail = GetInputDetail(Val(Bill.Text))
                Else
                    If gbln�շ���� Then
                        If Bill.RowData(Bill.Row) = 0 Then
                            sta.Panels(2) = "û��ȷ���������,�����������"
                            Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                        str��� = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
                    Else
                        Call GetOperatorInfo(mrs������, mobjBill.������, bln��ʿ)
                        str��� = IIf(bln��ʿ, "'E','M','4'", gstr�շ����)
                    End If
                    
                    int������Դ = 2
                    If zlCheckBill���ڷ�ɢװ��ҩ Then
                        '���ڷ�ɢװ��,�����оͲ��ܽ���¼��
                        Bill.Text = "": Bill.TxtVisible = False
                        Bill.SetFocus: Cancel = True: Exit Sub
                    End If
                    lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, int������Դ, int����, gblnסԺ��λ, str���, Bill.Text, _
                        Bill.TxtHwnd, str��׼��Ŀ, 0, , , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
                    If lng��Ŀid <> 0 Then
                        Set mobjDetail = GetInputDetail(lng��Ŀid)
                    Else
                        Bill.Text = "": Bill.TxtVisible = False
                        Bill.SetFocus: Cancel = True: Exit Sub
                    End If
                End If
                
                Bill.TxtVisible = False '(���Ӳ���)
                
                '�������ò��˲�������
                If InStr(",5,6,7,", mobjDetail.���) = 0 Then
                    If Not CheckFeeItemLimitDept(mobjDetail.ID, IIf(mbytUseType = 2, UserInfo.����ID, mobjBill.����ID), IIf(mbytUseType = 2, UserInfo.����ID, mobjBill.����ID)) Then
                        If mbytUseType = 2 Then
                            MsgBox "���շ���Ŀ�Ե�ǰ���˲����Ϳ��Ҳ����ã�", vbInformation, gstrSysName
                        Else
                            MsgBox "���շ���Ŀ�Ե�ǰ���˲����Ϳ��Ҳ����ã�", vbInformation, gstrSysName
                        End If
                        Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                    End If
                End If
                
                If InStr(",5,6,7,", mobjDetail.���) > 0 And mblnNurseStation Then
                    MsgBox "��ʿվ�������ʲ���¼��ҩƷ��Ŀ��", vbInformation, gstrSysName
                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                End If
                

                '��鶾�����ͼ�ֵ����Ȩ��
                If CheckDrugType(mobjDetail) = False Then
                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                End If
 
                
                '���ҩƷ�����Ƿ��ظ�:������ʱ��ͬһҩ���������ظ�(����ֻ����)
                If InStr(",5,6,7,", mobjDetail.���) > 0 Or _
                    (mobjDetail.��� = "4" And mobjDetail.��������) Then
                    If PhysicExist(mobjDetail, Bill.Row) Then
                        Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                    End If
                End If
                
                '��鴦��ְ��
                If InStr(",5,6,7,", mobjDetail.���) > 0 And mbln����ְ���� Then
                    mobjDetail.����ְ�� = Get����ְ��(mobjDetail.ID)
                    '���в���
                    If CheckDuty(mobjDetail, True) > 0 Then
                        Bill.TxtSetFocus: Cancel = True: Exit Sub
                    End If
                End If
                
                '��ȡҩƷ�����Ϣ
                
                '���˿���ID
                lng���˿���ID = mobjBill.����ID
                If cbo��������.Visible Then
                    If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
                Else
                    If lng���˿���ID = 0 Then lng���˿���ID = mlngDeptID
                    If lng���˿���ID = 0 Then lng���˿���ID = GetNurseStationFirstPatiDeptID '��ʿ����վ,ȡ��һ�����˿���ID
                    If lng���˿���ID = 0 Then lng���˿���ID = mlng����ID
                End If
                
                '����ִ�п���ȱʡΪ���˲���,�������ָ����,��Ϊָ������
                If mobjDetail.��� = "4" Then
                    lngDoUnit = IIf(glng���ϲ��� > 0, glng���ϲ���, mobjBill.����ID)
                    If lngDoUnit = 0 Then lngDoUnit = IIf(cbo��������.Visible, Get��������ID, lng���˿���ID)
                End If
                
                lngDoUnit = Get�շ�ִ�п���ID(mobjDetail.���, mobjDetail.ID, _
                    mobjDetail.ִ�п���, lng���˿���ID, IIf(cbo��������.Visible, Get��������ID, lng���˿���ID), 2, lngDoUnit, mobjBill.����ID)
                
                If InStr(",5,6,7,", mobjDetail.���) > 0 Then
                    '��ǰ��ҩƷ���
                    dblStock = GetStock(mobjDetail.ID, lngDoUnit)
                    If gblnסԺ��λ Then
                        dblStock = dblStock / mobjDetail.סԺ��װ
                    End If
                    mobjDetail.��� = dblStock
                    Call ShowStock(mobjDetail.����, mobjDetail.���)
          
                ElseIf mobjDetail.��� = "4" And mobjDetail.�������� Then
                    dblStock = GetStock(mobjDetail.ID, lngDoUnit)
                    mobjDetail.��� = dblStock
                    Call ShowStock(mobjDetail.����, mobjDetail.���)
                End If
                
                 '��������
                If InStr(",5,6,7,", mobjDetail.���) > 0 And mbln����������� Then
                    mobjDetail.�������� = Get��������(mobjDetail.ID)
                End If
                
                '������Ŀ��Ӧ���
                If CheckInsureTheCode(mobjDetail) = False Then
                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                End If
                
                '����ժҪ(ȡ���е����Ա��޸�)
                If mobjBill.Details.Count >= Bill.Row Then
                    If mobjBill.Details(Bill.Row).Detail.ID = mobjDetail.ID Then
                        strժҪ = mobjBill.Details(Bill.Row).ժҪ
                    End If
                End If
                
                '������޸ĸ��շ�ϸĿ��
                Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                '59051:�ȵ���GetItemInfor
                '����ժҪ(������������и���ժҪ)
                
                mobjBill.Details(Bill.Row).Tag = ""
                If mobjBill.Details(Bill.Row).Detail.����ժҪ Then
                    If frmInputBox.InputBox(Me, "ժҪ", "������""" & mobjBill.Details(Bill.Row).Detail.���� & """��ժҪ��Ϣ:", 200, 3, True, False, strժҪ) Then
                        mobjBill.Details(Bill.Row).ժҪ = strժҪ
                        mobjBill.Details(Bill.Row).Tag = strժҪ
                    End If
                End If
                
                Call CalcMoney(tyPati, Bill.Row)                         '��ʱ,��ʹ�������������,���û������
                
                '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                If mobjBill.Details.Count = Bill.Row Then
                    If CheckAllPatiChargeWrang(Bill.Row) = False Then
                         Bill.Text = "": Cancel = True: Exit Sub
                    End If
                End If
                
                If mobjBill.Details(Bill.Row).���� <> 0 Then
                    If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                        MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                        mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                        Bill.Text = "": Cancel = True: Exit Sub
                    End If
                End If
                
                Call ShowDetails(Bill.Row)
                Call ShowMoney
                
                '�������ͼ��
                Call Check��������(Bill.Row)
                Call SetDrawDrugDeptEnabled
                Bill.Text = "": Bill.SetFocus
            End If
            
            If mobjBill.Details.Count >= Bill.Row Then
                mlngPreRow = 0  '�޸�������ʱ,�ָ���ֵ,�Ա���ʾ���
                With mobjBill.Details(Bill.Row)
                    '��һ�е�����ȷ��
                    If .�շ���� = "7" And gblnPay Then Bill.ColData(BillCol.����) = BillColType.Text  '����
                    If .�շ���� = "F" Then Bill.ColData(BillCol.��־) = BillColType.CheckBox '���ӱ�־
                    
                    '���������������
                    If .Detail.��� And InStr(",5,6,7,", .�շ����) = 0 _
                        And Not (.�շ���� = "4" And .Detail.��������) Then
                        Bill.ColData(BillCol.����) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus) '����
                        Bill.ColData(BillCol.����) = BillColType.Text '����
                    Else
                        Bill.ColData(BillCol.����) = BillColType.Text '����
                        Bill.ColData(BillCol.����) = BillColType.UnFocus '����
                    End If
                    
                    'ִ�п���
                    '��FillBillComboBox������ListIndexʱ����CboClick�¼�
                    mblncboEnterCell = True: Bill.Col = BillCol.ִ�п���: mblncboEnterCell = False
                    Call FillBillComboBox(Bill.Row, BillCol.ִ�п���, Not blnInput)  'ֱ�ӻس�ʱ����ִ�п���
                    mblncboEnterCell = True: Bill.Col = BillCol.��Ŀ: mblncboEnterCell = False
                    
                    blnSkip = Bill.ListCount = 1
                    If Not blnSkip And InStr(",4,5,6,7,", .�շ����) > 0 Then
                        'ָ���˹̶�ҩ��ʱ,��������ѡ��
                        Select Case .�շ����
                            Case "4"
                                blnSkip = glng���ϲ��� > 0 And .ִ�в���ID = glng���ϲ���
                            Case "5"
                                blnSkip = glng��ҩ�� > 0 And .ִ�в���ID = glng��ҩ��
                            Case "6"
                                blnSkip = glng��ҩ�� > 0 And .ִ�в���ID = glng��ҩ��
                            Case "7"
                                blnSkip = glng��ҩ�� > 0 And .ִ�в���ID = glng��ҩ��
                        End Select
                    End If
                    If blnSkip Then
                        Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus: .Key = 1
                    Else
                        Bill.ColData(BillCol.ִ�п���) = BillColType.ComboBox: .Key = Bill.ListCount
                    End If
                    
                    '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                    If .�շ���� = "4" And .Detail.�������� Then
                        Call CheckValidity(.�շ�ϸĿID, .ִ�в���ID, .����, False) '��ȷ������,��������
                    End If
                    
                     '������Ŀ����,�������շ���Ŀ�д�����Ŀ����δȡ��ȡ,ҩƷ�����ж�,ҩƷ��������������
                    If Bill.TextMatrix(0, Bill.Col) = "��Ŀ" And InStr(",5,6,7,", .�շ����) = 0 Then
                        If (gbln��������ۿ� And mobjBill.Details(Bill.Row).�������� = 0) Or Not gbln��������ۿ� Then  '(����м���,ֻȡһ��)
                            If ShouldDO(Bill.Row) Then
                               Call SetSubItem
                               mlngPreRow = 0 'ͨ���б仯��־������ȷ��������
                            End If
                        End If
                    End If
                    
                End With
            End If
            
            'ֻ����һ�θ���
            If mobjBill.Details.Count >= Bill.Row And Bill.Row >= 2 And Bill.Active And Visible Then
                If mobjBill.Details(Bill.Row).�շ���� = "7" Then
                    For i = 1 To Bill.Row - 1
                        If mobjBill.Details(i).�շ���� = "7" Then
                            '����ִ�иù��̣�����ᶨλ��һ����Ԫ,�ȶ�λ������,����һ����Ԫ������
                            'ѡ����øù��̣����ú���͸��س������ﲻ���ٻس��������������س���Ч��(�ؼ�ԭ��)��
                            Bill.Col = BillCol.����: Exit For
                        End If
                    Next
                End If
            End If
            
        Case "����"
            If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                '���ֺϷ���
                If Not IsNumeric(Bill.Text) Then
                    MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                    Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                End If
                If Val(Bill.Text) <= 0 Or Val(Bill.Text) <> Int(Val(Bill.Text)) Then
                    MsgBox "����Ӧ��Ϊ����������", vbInformation, gstrSysName
                    Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                End If
                
                '�������
                If gcurMaxMoney > 0 Then
                    If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� * Bill.TextMatrix(Bill.Row, BillCol.����) > gcurMaxMoney Then
                        If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                            Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                        End If
                    End If
                End If
                
            
                '����ҩ���Ǵ�����Ŀ�ſɸ��ĸ���(������ı�,����Ҳ��)
                If mobjBill.Details(Bill.Row).�շ���� = "7" Then
                    '������ʱ��ҩƷ�����ֹ����(û�з�����ʱ��ҩƷ�����޸ĸ���������)
                    If mobjBill.Details(Bill.Row).Detail.���� Or mobjBill.Details(Bill.Row).Detail.��� Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� * IIf(mlngSelPatiCount = 0, 1, mlngSelPatiCount) > mobjBill.Details(Bill.Row).Detail.��� Then
                            MsgBox """" & mobjBill.Details(Bill.Row).Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                            Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                        End If
                    End If
                          
                    '�������ʱ�ۻ������ҩ���ĸ��������Ƿ��㹻
                    For i = 1 To mobjBill.Details.Count
                        If i <> Bill.Row And mobjBill.Details(i).�շ���� = "7" _
                            And (mobjBill.Details(i).Detail.��� Or mobjBill.Details(i).Detail.����) Then
                            If Val(Bill.Text) * mobjBill.Details(i).���� * IIf(mlngSelPatiCount = 0, 1, mlngSelPatiCount) > mobjBill.Details(i).Detail.��� Then
                                MsgBox "�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                        End If
                    Next
                                            
                    lngOld���� = mobjBill.Details(Bill.Row).����
                    '���㲢ˢ�¸���
                    mobjBill.Details(Bill.Row).���� = Bill.Text
                    
                    If mobjBill.Details(Bill.Row).���� <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                            mobjBill.Details(Bill.Row).���� = lngOld����
                            Call CalcMoneys(tyPati, Bill.Row)
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
                    Call CalcMoneys(tyPati, Bill.Row)
                    Call ShowDetails(Bill.Row)
                    
                    '����������ҩ����,����Ƕ�����,���޸������Ǵ����,����Ǵ���,���޸�ͬһ����Ĵ����.��Ϊ�޶�Ϊ�в�ҩ,������������
                    For i = 1 To mobjBill.Details.Count
                        If i <> Bill.Row And mobjBill.Details(i).�շ���� = "7" And mobjBill.Details(i).�������� = mobjBill.Details(Bill.Row).�������� Then
                            If mobjBill.Details(i).�������� = 0 Or (mobjBill.Details(i).�������� <> 0 And mobjBill.Details(i).Detail.���д��� = 0) Then     '1��2�̶��Ͱ������Ĳ���
                                mobjBill.Details(i).���� = Bill.Text
                                Call CalcMoneys(tyPati, i)
                                Call ShowDetails(i)
                            End If
                        End If
                    Next
                    Call ShowMoney
                Else
                    sta.Panels(2) = "������Ŀ�ĸ������ܸ��ģ�"
                    Bill.Text = mobjBill.Details(Bill.Row).����: Beep '�ָ�ԭ�и���ֵ
                End If
            End If
        Case "����"
            If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                 With mobjBill.Details(Bill.Row)
                     '��ҩ�������ת��
                    If .�շ���� = "7" Then Bill.Text = ConvertABCtoNUM(Bill.Text)
                    '���ֺϷ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = .����: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) = 0 Then
                        If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Bill.Text = .����: Cancel = True: Exit Sub
                        End If
                    End If
                    'ҩƷ����С��
                    If InStr(",5,6,7,", .�շ����) > 0 Then
                        If Val(Bill.Text) - Int(Val(Bill.Text)) <> 0 And InStr(mstrPrivsOpt, ";ҩƷ����С��;") = 0 Then
                            MsgBox "��û��Ȩ������С����", vbInformation, gstrSysName
                            Bill.Text = .����: Cancel = True: Exit Sub
                        End If
                    End If
                    '�������
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * .���� * Bill.TextMatrix(Bill.Row, BillCol.����) > gcurMaxMoney Then
                            If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = .����: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Bill.Text = FormatEx(Bill.Text, 5)
                    If InStr(",5,6,7,", .�շ����) > 0 And gblnסԺ��λ Then
                        dblNum = Val(Bill.Text) * .���� * .Detail.סԺ��װ
                    Else
                        dblNum = Val(Bill.Text) * .����
                    End If
                        
                    '�����Ϸ��Լ��
                    If Val(Bill.Text) * .���� < 0 Then
                        MsgBox "��������ʱ�����������ʣ�", vbInformation, gstrSysName
                        Bill.Text = .����: Cancel = True: Exit Sub
                    End If
                    
                    'ҩƷ�����
                    If Not CheckDrugStoreIsEnough(FormatEx(.���� * Val(Bill.Text), 6), mobjBill.Details(Bill.Row)) Then
                        Bill.Text = .����: Cancel = True: Exit Sub
                    End If
                    
                    dblPreTime = .����
                    .���� = Bill.Text
                    
                    '�����������
                    If mbln����������� And Not gbln�������� Then
                        If Not CheckLimit(mobjBill, Bill.Row, gblnסԺ��λ) Then
                            .���� = dblPreTime: Bill.Text = dblPreTime
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
                    If .Detail.¼������ > 0 And dblNum > .Detail.¼������ Then
                        If MsgBox("��������γ�����¼������" & FormatEx(.Detail.¼������ / IIf(gblnסԺ��λ, .Detail.סԺ��װ, 1), 5) & ",�Ƿ����?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                            .���� = dblPreTime: Bill.Text = dblPreTime
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
          
                    '���д������ܸ�������(����Ŀ���θı�,���д���������Ҳ��)
                    If .�������� <> 0 And .Detail.���д��� <> 0 Then
                        sta.Panels(2) = "����Ŀ�ǹ��д�����Ŀ,�����β��ܹ����ġ�"
                        .���� = dblPreTime: Bill.Text = dblPreTime
                        Exit Sub
                    End If
                                        
                    Call CalcMoneys(tyPati, Bill.Row)
                    
                    '����������(���Ѿ�������з��õ�δ��ʾǰ)
                    If MoneyOverFlow(mobjBill) Then
                        MsgBox "�����������µ��ݽ����������ʵ�������", vbInformation, gstrSysName
                        .���� = dblPreTime
                        Call CalcMoneys(tyPati, Bill.Row)
                        Bill.Text = "": Bill.TxtVisible = False
                        Cancel = True: Exit Sub
                    End If
                    
                    If .���� <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                            .���� = dblPreTime
                            Call CalcMoneys(tyPati, Bill.Row)
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ):����󱣴�ʱ����
                End With
                    
                Call ShowDetails(Bill.Row)
                '��������д���������
                For i = Bill.Row + 1 To mobjBill.Details.Count
                    If mobjBill.Details(i).�������� = Bill.Row Then
                        '28136
                        '���������ĸ���,��Ҫ���¼��еĸ������и��³ɸ���
                        With mobjBill.Details(i)
                            If .Detail.���д��� = 0 Then  '�ǹ��д���
                                If Abs(.����) <> Abs(.Detail.��������) Then GoTo NotCalc:
                                .���� = IIf(Val(Bill.Text) < 0, -1, 1) * .Detail.��������
                            ElseIf .Detail.���д��� = 1 Then '�̶��Ĺ��д���
                                .���� = IIf(Val(Bill.Text) < 0, -1, 1) * IIf(.Detail.�������� = 0, 1, .Detail.��������)
                            ElseIf .Detail.���д��� = 2 Then   '�������Ĺ��д���
                                .���� = Val(Bill.Text) * .Detail.��������
                            Else
                                 GoTo NotCalc:
                            End If
                        End With
                        
                        Call CalcMoneys(tyPati, i)
                        Call ShowDetails(i)
NotCalc:
                    End If
                Next

                
                Call ShowMoney
            ElseIf mobjBill.Details.Count >= Bill.Row Then
                If Val(Bill.TextMatrix(Bill.Row, Bill.Col)) = 0 Then
                    If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: Exit Sub
                    End If
                End If
            End If
                
            If Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus Then
                If CheckItemHaveSub(Bill.Row) Then
                    KeyCode = 0
                    Call LocateMainItemNextRow(Bill.Row)
                End If
            End If
            
        Case "����"
            If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                '���ֺϷ���
                If Not IsNumeric(Bill.Text) Then
                    MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                    Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                End If
                If Val(Bill.Text) < 0 Then
                    MsgBox "��Ŀ�۸�Ӧ��Ϊ������", vbInformation, gstrSysName
                    Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                End If
                
                '�������
                If gcurMaxMoney > 0 Then
                    If Val(Bill.Text) * mobjBill.Details(Bill.Row).���� * mobjBill.Details(Bill.Row).���� > gcurMaxMoney Then
                        If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                            Bill.Text = "": Cancel = True: Exit Sub
                        End If
                    End If
                End If

                Bill.Text = FormatEx(Bill.Text, 5)
                
                '���û�ж�Ӧ��������Ŀ,���޷�����
                If mobjBill.Details(Bill.Row).Detail.��� And mobjBill.Details(Bill.Row).InComes.Count > 0 Then
                    If Not (mobjBill.Details(Bill.Row).InComes(1).�ּ� = 0 And mobjBill.Details(Bill.Row).InComes(1).ԭ�� = 0) Then
                        strScope = CheckScope(mobjBill.Details(Bill.Row).InComes(1).ԭ��, mobjBill.Details(Bill.Row).InComes(1).�ּ�, CCur(Bill.Text))
                        If strScope <> "" Then
                            sta.Panels(2) = strScope
                            If Bill.TxtVisible And Len(Bill.Text) > 9 Then Bill.Text = mobjBill.Details(Bill.Row).InComes(1).��׼����
                            If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                            Cancel = True: Beep: Exit Sub
                        End If
                    End If
                    
                    dblPreMoney = mobjBill.Details(Bill.Row).InComes(1).��׼����
                    
                    mobjBill.Details(Bill.Row).InComes(1).��׼���� = Bill.Text '�����շ�ϸĿֻ�ܶ�Ӧһ��������Ŀ
                    Call CalcMoneys(tyPati, Bill.Row)
                    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ),����󱣴�ʱ����
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                Else
                    Bill.Text = "0"
                    sta.Panels(2) = "����Ŀ�������ö�Ӧ�ķ�Ŀ�������޷�������ã�"
                    Beep
                End If
            End If
        Case "ִ�п���"
            If mobjBill.Details.Count >= Bill.Row And Bill.ListIndex <> -1 Then
                With mobjBill.Details(Bill.Row)
                    If .ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                        .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                        If CheckItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row) '������ڴ���,��ı��ҩƷ�е�ִ�п���
                    End If
                    
                    'ҩƷ�����:��̬ҩ��,������ʱ��ҩƷҲҪ�����
                    If Not CheckDrugStoreIsEnough(FormatEx(.���� * .����, 6), mobjBill.Details(Bill.Row), True) Then
                        Cancel = True
                    End If
            
                    '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                    If .�շ���� = "4" And .Detail.�������� Then
                        Call CheckValidity(.�շ�ϸĿID, .ִ�в���ID, .����, False) '��ȷ������,��������
                    End If
                    
                    If CheckItemHaveSub(Bill.Row) Then
                        KeyCode = 0
                        Call LocateMainItemNextRow(Bill.Row)
                    End If
                    
                    Call CalcMoneys(tyPati, Bill.Row, True)
                    Call ShowDetails(Bill.Row)
                    If .�շ���� = "4" And .Detail.�������� And .�շ�ϸĿID <> 0 Then
                        Call ShowStock(.Detail.����, .Detail.���)
                    End If
                    If mobjBill.Details(Bill.Row).���� <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling, Bill.Row)) = False Then
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                End With
            End If
    Case "��־"
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Cancel = True
End Sub

Private Sub Bill_EnterCell(Row As Long, Col As Long)
    'ע��:���κ�exit sub ֮ǰ����mblncboClick = False,����,�޷�������
    Dim strStock As String, i As Long
    Dim strҩ��IDs As String
        
    If Not Bill.Active Then Exit Sub
    
    If Bill.ColData(Col) = BillColType.UnFocus Then Exit Sub
    
    If mblncboEnterCell Then Exit Sub  '����ͬһ������������bill��ֵ��ѭ������,ע�����κ�exit sub ֮ǰ����mblncboClick = False
    mblncboEnterCell = True
        
    '--------------------------------------------------------------------------
    '1.�иı��������ݴ��������     mlngPreRow    ��ǰ���Ƿ�ı�
    If zlCheckBill���ڷ�ɢװ��ҩ = True Then
        '��������д��ڷ�ɢװ��,��������
        Call SetBill�в�ҩEditEnabled
        mblncboEnterCell = False
         Exit Sub
    End If
   
    If mobjBill.Details.Count >= Bill.Row And mlngPreRow <> Row Then
        With mobjBill.Details(Bill.Row)
            '��ʾ���
            If InStr(",5,6,7,", .�շ����) > 0 And .�շ�ϸĿID <> 0 Then
                If gbln����ҩ�� Or gbln����ҩ�� Then
                    strStock = GetStockInfo(.�շ�ϸĿID, gbln����ҩ��, gbln����ҩ��, gblnסԺ��λ)
                    If strStock <> "" Then
                        If InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0 Then
                            sta.Panels(Pan.C2��ʾ��Ϣ) = "��" & Bill.Row & "�п��:" & strStock
                        Else
                            sta.Panels(Pan.C2��ʾ��Ϣ) = "��" & Bill.Row & "���п��."
                        End If
                    End If
                End If
                If strStock = "" Then
                    '���¿����ʾ
                    .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                    If gblnסԺ��λ Then
                        .Detail.��� = .Detail.��� / .Detail.סԺ��װ
                    End If
                    Call ShowStock(.Detail.����, .Detail.���)
                End If
            ElseIf .�շ���� = "4" And .Detail.�������� And .�շ�ϸĿID <> 0 Then
                .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                Call ShowStock(.Detail.����, .Detail.���)
            Else
                sta.Panels(2) = ""
            End If
                     
            Bill.ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)
            Bill.ColData(BillCol.��Ŀ) = BillColType.CommandButton
            
             '����Ǵ�����Ŀ������Ŀ�����,���������������Ŀ
            If CheckItemHaveSub(Row) Or .�������� > 0 Then
                Bill.ColData(BillCol.���) = BillColType.Text_UnModify
                Bill.ColData(BillCol.��Ŀ) = BillColType.Text_UnModify
            End If

            If .�շ���� = "7" And gblnPay Then
                Bill.ColData(BillCol.����) = BillColType.Text
            Else
                Bill.ColData(BillCol.����) = BillColType.UnFocus
            End If
            
            '���������������
            If .Detail.��� And InStr(",5,6,7,", .�շ����) = 0 _
                And Not (.�շ���� = "4" And .Detail.��������) Then
                Bill.ColData(BillCol.����) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus) '����
                Bill.ColData(BillCol.����) = BillColType.Text '���
            Else
                Bill.ColData(BillCol.����) = BillColType.Text
                Bill.ColData(BillCol.����) = BillColType.UnFocus
            End If
            
            If .Key = "1" Then    'ָ���˹̶�ҩ��ʱ,��������ѡ��ִ�п���
                Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus
            Else
                Bill.ColData(BillCol.ִ�п���) = BillColType.ComboBox
            End If
                
            If .�շ���� = "F" Then
                Bill.ColData(BillCol.��־) = BillColType.CheckBox
            Else
                Bill.ColData(BillCol.��־) = BillColType.UnFocus
            End If
            
            'ֻ����һ�����,������ѡ�����
            If mblnOne Then Bill.ColData(BillCol.���) = BillColType.UnFocus
        End With
    
        '��ʾժҪ
        If mobjBill.Details(Bill.Row).ժҪ <> "" Then
            sta.Panels(2) = sta.Panels(2) & "  ժҪ:" & mobjBill.Details(Bill.Row).ժҪ
        End If
    End If
    
    '������δ�������,��ָ��е�����
    If mobjBill.Details.Count < Bill.Row Then
        Bill.ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)  '�����,��������ʱ�ᱻ�ı�
        Bill.ColData(BillCol.��Ŀ) = BillColType.CommandButton   '��Ŀ��,��������ʱ�ᱻ�ı�
    End If
    
    
    '-----------------------------------------------------------------
    '2.�иı�������ݴ������ʾ����
    If Bill.ColData(Bill.Col) = BillColType.ComboBox Then  '���ص�ǰ�е�����������
        Call FillBillComboBox(Bill.Row, Bill.Col, True)
    End If
    
    If gbln�շ���� And Bill.TextMatrix(Row, BillCol.���) = "" And mblnOne Then
        mrsClass.Filter = "����=" & gstr�շ����
        Bill.TextMatrix(Row, BillCol.���) = mrsClass!���
        Bill.RowData(Row) = Asc(mrsClass!����)
    End If
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "���" '���������ʱ���ᶨλ�������
            SetWidth Bill.cboHwnd, 70
            If Bill.TextMatrix(Row, Col) = "" Then
                If mblnOne Then
                    mrsClass.Filter = "����=" & gstr�շ����
                    Bill.TextMatrix(Row, Col) = mrsClass!���
                    Bill.RowData(Row) = Asc(mrsClass!����)
                ElseIf Row > 1 Then
                    Bill.ListIndex = -1
                    For i = 0 To Bill.ListCount - 1
                        If InStr(Bill.List(i), Bill.TextMatrix(Row - 1, Col)) > 0 Then Bill.ListIndex = i: Exit For
                    Next
                End If
            ElseIf Row >= 1 And Bill.TextMatrix(Row, Col) <> "" Then
                For i = 0 To Bill.ListCount - 1
                    If InStr(Bill.List(i), Bill.TextMatrix(Row, Col)) > 0 Then
                        Bill.ListIndex = i: Exit For
                    End If
                Next
                If Bill.ListIndex = -1 Then
                    Bill.ListIndex = SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal Bill.TextMatrix(Row - 1, Col))
                End If
            End If
        Case "ִ�п���"
            SetWidth Bill.cboHwnd, 130
        Case "����"
            Bill.TextLen = 3
            Bill.TextMask = "0123456789" & Chr(8)
        Case "����"
            Bill.TextLen = 8
            Bill.TextMask = "0123456789." & Chr(8)
            
            If mobjBill.Details.Count >= Bill.Row Then
                If InStr(",5,6,7,", mobjBill.Details(Bill.Row).�շ����) > 0 Then
                    If InStr(mstrPrivsOpt, ";ҩƷ����С��;") = 0 Then
                        Bill.TextMask = Replace(Bill.TextMask, ".", "")
                    End If
                End If
                '��ҩ�������
                If mobjBill.Details(Bill.Row).�շ���� = "7" Then
                        Bill.TextMask = Bill.TextMask & gstrABC & LCase(gstrABC)
                End If
            End If
        Case "����"
            Bill.TextLen = 10
            Bill.TextMask = "0123456789." & Chr(8)
    End Select
    
    '����,����������е����ʱ,�������л�û�п�ʼ
    If Bill.TextMatrix(Row, BillCol.��Ŀ) = "" Then
        mlngPreRow = 0
    ElseIf mobjBill.Details.Count >= Row Then
        mlngPreRow = Row
    End If
    
    mblncboEnterCell = False
End Sub
Private Sub Bill_LostFocus()
    Bill.TxtVisible = False
    Bill.CmdVisible = False
    Bill.CboVisible = False
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Bill.ToolTipText = Bill.TextMatrix(Bill.MouseRow, Bill.MouseCol)
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    Dim i As Long
    With Bill
        '������ʱ,�������ÿ����Ѿ������ĵĿɱ������е���ֵ
        .ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)  '�����,��������ʱ�ᱻ�ı�
        .ColData(BillCol.��Ŀ) = BillColType.CommandButton    '��Ŀ��,��������ʱ�ᱻ�ı�
        .ColData(BillCol.����) = BillColType.UnFocus   '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
        .ColData(BillCol.����) = BillColType.UnFocus  '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
        .ColData(BillCol.��־) = BillColType.UnFocus  '��־ȱʡ����,��Ϊ����ʱ,��Ϊ��ѡ(-1)
        '����б༭����������ɫ
        .SetColColor BillCol.���, &HE7CFBA
        .SetColColor BillCol.��Ŀ, &HE7CFBA
        .SetColColor BillCol.����, &HE7CFBA
        .SetColColor BillCol.ִ�п���, &HE7CFBA
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.��־, &HE0E0E0
        
        .TextMatrix(Row, BillCol.��) = Row
        '����ط��ֶ����ò�ִ��
        If Visible And Bill.Active And Row > 0 And .ColData(BillCol.���) <> BillColType.UnFocus And Not mblnNewRow Then
            Call zlCommFun.PressKey(13)
        End If
    End With
End Sub

Private Sub SetDefaultDoctor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ������
    '����:���˺�
    '����:2015-07-10 15:20:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If cbo������.ListCount = 0 Then Exit Sub
    If cbo������.ListCount = 1 Then cbo������.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub cbo��������_Click()
    Dim i As Long, lng��������ID As Long
    If cbo��������.ListIndex <> -1 Then lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    If mobjBill.��������ID = lng��������ID Then Exit Sub
    If mrs��ҩ����.RecordCount <> 0 Then
        For i = 0 To cboDrawDept.ListCount - 1
             If cboDrawDept.ItemData(i) = lng��������ID Then
                mobjBill.��ҩ����ID = lng��������ID
                cboDrawDept.ListIndex = i: Exit For
             End If
        Next
    End If
    
    mobjBill.��������ID = lng��������ID
        
    '��������ȷ��ҽ��
    If Not gblnFromDr Then
        If cbo��������.ListIndex <> -1 Then
            If gbln������ Then
                Call FillDoctor(cbo������, mrs������)
            Else
                Call FillDoctor(cbo������, mrs������, lng��������ID)
            End If
            Call SetDefaultDoctor
        Else
            cbo������.Clear
        End If
        Call cbo������_Click
    End If
    
    
    '�������������Ŀ��ִ�п���
    If cbo��������.ListIndex <> -1 And cbo��������.Visible Then
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                '�������շ���Ŀ
                If InStr(",4,5,6,7,", .Detail.���) = 0 And .Detail.ִ�п��� = 6 Then '6-�����˿���
                    .ִ�в���ID = cbo��������.ItemData(cbo��������.ListIndex)
                    'ˢ����ʾ����ִ�п���
                    If i <= Bill.Rows - 1 And .ִ�в���ID <> 0 Then
                        mrsUnit.Filter = "ID=" & .ִ�в���ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(i, BillCol.ִ�п���) = mrsUnit!���� & "-" & mrsUnit!����
                        Else
                            Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.ִ�в���ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                    End If
                End If
            End With
        Next
    End If
End Sub

     
Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
    
    If KeyAscii <> 13 Then Exit Sub
    If cbo��������.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cbo������.ListIndex >= 0 Then lngҽ��ID = cbo������.ItemData(cbo������.ListIndex)
    If mrs�������� Is Nothing Then Call FillDept(cbo��������, mrs��������, mrs������, mstrPrivs, mbytUseType, mlngDeptID, lngҽ��ID)
    
    If zlSelectDept(Me, mlngModule, cbo��������, mrs��������, cbo��������.Text) = False Then
        Call Beep: mobjBill.��������ID = 0
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cbo������_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo������_Validate(Cancel As Boolean)
    If cbo������.Locked Then Exit Sub
    
    If cbo������.Text <> "" Then
        If cbo.FindIndex(cbo������, zlStr.NeedName(cbo������.Text), True) = -1 Then cbo������.ListIndex = -1: cbo������.Text = ""
    End If
'    If cbo������.Text = "" And mblnKeyReturn Then
'        Call cbo������_KeyPress(vbKeyReturn)
'    End If
    mblnKeyReturn = False
    '����������ȷ��������ʱ,���ܴ�ʱ��ѡ������,��ȥ�����������Һ�����ѡ
    If gblnFromDr And gbln������ And cbo������.ListIndex = -1 And mlngSelPatiCount <> 0 Then Cancel = True
End Sub

Private Sub cbo������_Click()
    Dim lng������ID As Long
 
    If mobjBill.������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text)) Then Exit Sub
    
    mobjBill.������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
    If gblnFromDr Then
        If cbo������.ListIndex <> -1 Then
            lng������ID = cbo������.ItemData(cbo������.ListIndex)
            
            Call FillDept(cbo��������, mrs��������, mrs������, mstrPrivs, mbytUseType, mlngDeptID, lng������ID)
            Call SetDefaultDept(cbo��������, mrs��������, mrs������, lng������ID)
        Else
            cbo��������.Clear
        End If
        Call cbo��������_Click
    End If
                        
    '��ʿ���
    If Bill.Active Then
        If mobjBill.Details.Count < Bill.Rows - 1 And Bill.Row = Bill.Rows - 1 _
            And Bill.RowData(Bill.Rows - 1) <> 0 Then
            '�����Ч����
            Bill.TextMatrix(Bill.Rows - 1, BillCol.���) = ""
            Bill.RowData(Bill.Rows - 1) = 0
        ElseIf Bill.Col = BillCol.��� Then
            Call Bill_EnterCell(Bill.Row, Bill.Col) 'ˢ��
        End If
    End If
    
    '��ʿ���:�жϷǷ�����
    If CheckInhibitiveByNurse(mobjBill, mrs������) Then
        MsgBox "��ʿֻ���������Ƽ�������Ŀ,�������д����������͵���Ŀ��", vbInformation, gstrSysName
    End If
End Sub
Private Sub cbo������_KeyDown(KeyCode As Integer, Shift As Integer)
    If cbo������.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo������.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cbo������_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim rsTemp As ADODB.Recordset
    Dim strAdded As String
    If KeyAscii = vbKeyTab Then
        mblnKeyReturn = True
    End If
    
    If Not KeyAscii = 13 Then Exit Sub
    
    If cbo������.Locked Then
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    
    
    strText = UCase(cbo������.Text)
    If cbo������.ListIndex <> -1 Then
        '�����б�ʱ,�����ı�������������
        If strText <> cbo������.List(cbo������.ListIndex) Then Call zlControl.CboSetIndex(cbo������.hWnd, -1)
    End If
    If strText = "" Then
        cbo������.ListIndex = -1
    ElseIf cbo������.ListIndex = -1 Then
        intIdx = -1
        strFilter = IIf(gbln��ʿ, "��Ա����<>''", "��Ա����<>'��ʿ'")
        '���˺�:22383
        '�ȸ��Ƽ�¼��
        Set rsTemp = zlDatabase.zlCopyDataStructure(mrs������)
        Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
        Dim strCompents As String 'ƥ�䴮
        
        strCompents = Replace(gstrLike, "%", "*") & strText & "*"
        
        If IsNumeric(strText) Then
            intInputType = 0
        ElseIf zlCommFun.IsCharAlpha(strText) Then
            intInputType = 1
        Else
            intInputType = 2
        End If
        
        mrs������.Filter = strFilter: iCount = 0
        With mrs������
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not mrs������.EOF
                Select Case intInputType
                Case 0  '�������ȫ����
                    '������������,��Ҫ���:
                    '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������
                    '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                    
                    
                    '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                    If Nvl(!���) = strText Then strResult = Nvl(!����): iCount = 0: Exit Do
                    
                    '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                    If Val(Nvl(!���)) = Val(strText) Then
                        If iCount = 0 Then strResult = Nvl(!����)
                        iCount = iCount + 1
                    End If
                    
                    '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                     If Val(Nvl(!���)) Like strText & "*" Then
                        If isCheck������Exists(Nvl(!����)) And InStr(strAdded, "," & Nvl(!���) & ",") = 0 Then
                            Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                            strAdded = strAdded & "," & Nvl(!���) & ","
                        End If
                     End If
                Case 1  '�������ȫ��ĸ
                    '����:
                    ' 1.����ļ������,��ֱ�Ӷ�λ
                    ' 2.���ݲ�����ƥ����ͬ����
                    
                    '1.����ļ������,��ֱ�Ӷ�λ
                    If Trim(Nvl(!����)) = strText Then
                        If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                        iCount = iCount + 1
                    End If
                    
                    '2.���ݲ�����ƥ����ͬ����
                    If Trim(Nvl(!����)) Like strCompents Then
                        If isCheck������Exists(Nvl(!����)) And InStr(strAdded, "," & Nvl(!���) & ",") = 0 Then
                            Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                            strAdded = strAdded & "," & Nvl(!���) & ","
                        End If
                    End If
                Case Else  ' 2-����
                    '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                    '1.����\�������,ֱ�Ӷ�λ
                    '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                    
                    '1.����\�������,ֱ�Ӷ�λ
                    If Trim(!���) = strText Or Trim(!����) = strText Or Trim(!����) = strText Then
                        If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                        iCount = iCount + 1
                    End If
                    
                    '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                    If Trim(!���) Like strText & "*" Or Trim(Nvl(!����)) Like strCompents Or Trim(Nvl(!����)) Like strCompents Then
                        If isCheck������Exists(Nvl(!����)) And InStr(strAdded, "," & Nvl(!���) & ",") = 0 Then
                            Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                            strAdded = strAdded & "," & Nvl(!���) & ","
                        End If
                    End If
                End Select
                mrs������.MoveNext
            Loop
        End With
         If iCount > 1 Then strResult = ""
        If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!����)
        '���˺�:ֱ�Ӷ�λ
        If strResult <> "" Then
            rsTemp.Close: Set rsTemp = Nothing
            If isCheck������Exists(strResult, True) Then zlCommFun.PressKey vbKeyTab: mblnKeyReturn = True
            Exit Sub
        End If
        
        '��Ҫ����Ƿ��ж������������ļ�¼
        If rsTemp.RecordCount <> 0 Then
            '�Ȱ�ĳ�ַ�ʽ��������
            Select Case intInputType
            Case 0 '����ȫ����
                rsTemp.Sort = "���"
            Case 1 '����ȫƴ��
                rsTemp.Sort = "����"
            Case Else
                '����ѡ������
                If gbyt��������ʾ = 1 Then '����
                    rsTemp.Sort = "����"
                Else
                    rsTemp.Sort = "���"
                End If
            End Select
            '����ѡ����
            Dim rsReturn As ADODB.Recordset
            If zlDatabase.zlShowListSelect(Me, glngSys, 1133, cbo������, rsTemp, True, "", "ȱʡ,ְ��,���ȼ���", rsReturn) Then
                If Not rsReturn Is Nothing Then
                    If rsReturn.RecordCount <> 0 Then
                        '���ж�λ
                        If isCheck������Exists(Nvl(rsReturn!����), True) Then
                            'zlCommFun.PressKey vbKeyTab
                            mblnKeyReturn = True
                        End If
                    End If
                End If
            End If
        Else
            'δ�ҵ�
            rsTemp.Close: Set rsTemp = Nothing
            KeyAscii = 0: zlControl.TxtSelAll cbo������: Exit Sub
        End If
        rsTemp.Close: Set rsTemp = Nothing
         
    ElseIf Not mblnDrop Then
        '�س���꾭��
        Call cbo������_Click
        Call zlCommFun.PressKey(vbKeyTab)
        mblnKeyReturn = True
        Exit Sub
    End If
    If cbo������.ListIndex = -1 Then
        cbo������.Text = ""
        mobjBill.������ = ""
        If gblnFromDr Then Exit Sub
    Else
        mobjBill.������ = zlStr.NeedName(cbo������.Text)
        If intIdx <> -1 And mblnDrop Then
            '�����س�-ǿ�м���Click
            Call cbo������_Click
        ElseIf intIdx <> cbo������.ListIndex And intIdx <> -1 Then
            '������ѡ��-�Զ�����Click
            cbo������.SetFocus
            Call zlCommFun.PressKey(vbKeyF4)
            Exit Sub
        ElseIf intIdx <> -1 Then
            'һ��������-ǿ�м���Click
            Call cbo������_Click
        End If
    End If
    Call zlCommFun.PressKey(vbKeyTab)
    mblnKeyReturn = True
End Sub
  
Private Function isCheck������Exists(ByVal str���� As String, _
    Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ��ڿ����������б���.
    '���:str����-����
    '     blnLocateItem:�Ƿ�ֱ�Ӷ�λ
    '����:���ڷ���gtrue,���򷵻�False
    '����:���˺�
    '����:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cbo������.ListCount - 1
        If zlStr.NeedName(cbo������.List(i)) = str���� Then
            If blnLocateItem Then cbo������.ListIndex = i
            isCheck������Exists = True
            Exit Function
        End If
    Next
End Function
Private Sub chk�Ӱ�_Click()
    Dim blnAdd As Boolean
    Dim tyPati As TY_PATIINFOR
    
    If Not chk�Ӱ�.Visible Then Exit Sub
    
    blnAdd = OverTime(zlDatabase.Currentdate)
    If chk�Ӱ�.Value = Unchecked And blnAdd Then
        If MsgBox("��ǰ���ڼӰ�ʱ�䷶Χ��,Ҫȡ���Ӱ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Checked
        End If
    End If
    
    If chk�Ӱ�.Value = Checked And Not blnAdd Then
        If MsgBox("��ǰ�����ڼӰ�ʱ�䷶Χ��,Ҫִ�мӰ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Unchecked
        End If
    End If
    
    mobjBill.�Ӱ��־ = IIf(chk�Ӱ�.Value = Checked, 1, 0)
    
    '���¼���۸�
    If Not mobjBill.Details.Count = 0 Then
        Call CalcMoneys(tyPati)
        Call ShowDetails
        Call ShowMoney
    End If
End Sub

Private Sub chk�Ӱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub


Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub
Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.����ʱ�� = CDate(txtDate.Text)
End Sub

Private Sub cboNO_GotFocus()
    zlControl.TxtSelAll cboNO
    cboNO.Locked = True
End Sub


 
Private Sub SetSubItem()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շ���Ŀ��,���ص�ǰ�շ���Ŀ�Ĵ�����Ŀ�����ü�����,����ʾ�ڵ��ݿؼ���
    '����:���˺�
    '����:2015-07-10 11:48:01
    '������:Bill_KeyDown��������Ŀ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, lngMainRow As Long
    Dim lngDoUnit As Long, lng���˿���ID As Long
    Dim bln��������ۿ� As Boolean
    Dim strժҪ As String, tyPati As TY_PATIINFOR
    Dim dblStock As Double
    Dim cllData As Collection
    
    lngMainRow = Bill.Row               '�������
    If gbln��������ۿ� Then            '����������ηѱ�,����ܼ����ۿ۲�����Ч,�����ܼ���
        bln��������ۿ� = Not mobjBill.Details(lngMainRow).Detail.���ηѱ�
    End If
    
    lng���˿���ID = mobjBill.����ID
    If cbo��������.Visible Then
        If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
    Else
        If lng���˿���ID = 0 Then lng���˿���ID = mobjBill.����ID
    End If
    
    With mobjBill.Details(lngMainRow)
        Set mcolDetails = New Details
        Set mcolDetails = GetSubDetails(.�շ�ϸĿID)
        For i = 1 To mcolDetails.Count
            If mobjBill.Details.Count >= Bill.Rows - 1 Then
                Bill.Rows = Bill.Rows + 1
                mblnNewRow = True
                Call bill_AfterAddRow(Bill.Rows - 1)
                mblnNewRow = False
            End If
            Bill.TextMatrix(Bill.Rows - 1, BillCol.���) = "" '�б�Ҫ����
            
            'a.������ĿΪ��ҩƷ��Ŀ��ִ�п���
            lngDoUnit = 0
            If InStr(",4,5,6,7,", mcolDetails(i).���) = 0 Then
                 If mcolDetails(i).��� = .�շ���� Or mcolDetails(i).ִ�п��� = 0 Then
                    '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                    '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                    lngDoUnit = .ִ�в���ID
                 Else
                    '3.������ҩ��Ŀ��ִ�п���
                    lngDoUnit = Get�շ�ִ�п���ID(mcolDetails(i).���, mcolDetails(i).ID, _
                        mcolDetails(i).ִ�п���, lng���˿���ID, Get��������ID, 2, , mobjBill.����ID)
                 End If
            'b.������ĿΪҩƷ,���ĵ�ִ�п���
            Else
                lngDoUnit = Get�շ�ִ�п���ID(mcolDetails(i).���, mcolDetails(i).ID, _
                    mcolDetails(i).ִ�п���, lng���˿���ID, Get��������ID, 2, .ִ�в���ID, mobjBill.����ID)  '���Ĵ���ȱʡ������ִ�п�����ͬ
            End If
            
            '���»�ȡ���
            Call SetDetailtStock(lngDoUnit, mcolDetails(i))
     
                       
            '������Ŀ��Ӧ���
            If CheckInsureTheCode(mcolDetails(i)) = False Then
                Exit Sub
            End If
             
            Call SetDetail(mcolDetails(i), Bill.Rows - 1, lngDoUnit, Bill.Row)
            
            Call CalcMoney(tyPati, Bill.Rows - 1, bln��������ۿ�)
            Call ShowDetails(Bill.Rows - 1)
            'CalcMoney���ȵ���GetuItemInsure���ܷ���ժҪ
             strժҪ = mobjBill.Details(Bill.Rows - 1).ժҪ
        Next
        
        If bln��������ۿ� Then
            Call CalcMoney(tyPati, lngMainRow, bln��������ۿ�) '�����������Ӧ����ʵ��,��Ϊ��û�м������ǰ�����ǰ������������.
            
            Call Calc��������ʵ��(lngMainRow)
        End If
        
        Call ShowMoney
    End With
End Sub



Private Sub LocateMainItemNextRow(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ŀ����һ��(�������
    '����:���˺�
    '����:2015-07-10 11:44:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�������� = lngRow Then
            If mobjBill.Details(i).Detail.���д��� = 0 Then Exit For
        End If
    Next
    
    If i <= mobjBill.Details.Count Then
        Bill.Col = BillCol.����
        Bill.Row = i: Bill.MsfObj.TopRow = i
    Else
        Call LocateNewRow
    End If
End Sub

Private Sub LocateNewRow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��λ������
    '����:���˺�
    '����:2015-07-10 11:46:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjBill.Details.Count >= Bill.Rows - 1 Then
        Bill.Rows = Bill.Rows + 1
        mblnNewRow = True
        Call bill_AfterAddRow(Bill.Rows - 1)
        mblnNewRow = False
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.���
    Else
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.���
    End If
    If Not Me.ActiveControl Is Bill Then
        If Bill.Active And Bill.Visible Then Bill.SetFocus
    End If
End Sub

Private Sub SetDetailtStock(ByVal lngִ�п���ID As Long, ByRef objDetail As Detail)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�������ϸ�Ŀ������
    '���ƣ����˺�
    '���ڣ�2010-07-12 14:27:51
    '˵����
    '      bug:31374
    '------------------------------------------------------------------------------------------------------------------------
    Dim strҩ��IDs As String, dblStock As Double
    '��ȡ���
    '�������ҩƷ������
    If InStr(1, "5,6,7,4", objDetail.���) = 0 Then Exit Sub
    If objDetail.��� = "4" And objDetail.�������� = False Then Exit Sub
    If objDetail.��� = "4" Then
        '����
        dblStock = GetStock(objDetail.ID, lngִ�п���ID)
        objDetail.��� = dblStock
        Exit Sub
    End If
    dblStock = GetStock(objDetail.ID, lngִ�п���ID)
    If gblnסԺ��λ Then
        dblStock = dblStock / objDetail.סԺ��װ
    End If
    objDetail.��� = dblStock  '��¼��ǰ��ҩƷ���
    Exit Sub
End Sub

Private Sub SetSubItemDept(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������ִ�п��ҵı仯,ˢ�·�ҩ�����ִ�п���
    '����:���˺�
    '����:2015-07-10 14:52:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lng���˿���ID As Long
    
    With mobjBill
        '��ȡ���д����ִ�п�������,������ȡ(��Ϊ�����ϵĴ�����Ϣ�������޸Ĺ���)
        Set mcolDetails = GetSubDetails(.Details(lngRow).�շ�ϸĿID)
        
        lng���˿���ID = .����ID
        If cbo��������.Visible Then
            If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
        Else
            If lng���˿���ID = 0 Then lng���˿���ID = .����ID
        End If

        For i = lngRow + 1 To .Details.Count
            If .Details(i).�������� = lngRow Then
                '������ΪҩƷ�����ĵ���Ŀ��ִ�п��Ҳ�������䶯
                If InStr(",4,5,6,7,", .Details(i).�շ����) = 0 Then
                    If .Details(i).�շ���� = .Details(lngRow).�շ���� Then
                        '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                        .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                    Else
                        For j = 1 To mcolDetails.Count
                            If mcolDetails.Item(j).ID = .Details(i).Detail.ID Then
                                Exit For
                            End If
                        Next
                        If j <= mcolDetails.Count Then
                            If mcolDetails.Item(j).ִ�п��� = 0 Then
                                '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                                 .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                            Else
                                '3.������ҩ��Ŀ��ִ�п���
                                .Details(i).ִ�в���ID = Get�շ�ִ�п���ID(mcolDetails(j).���, mcolDetails(j).ID, _
                                    mcolDetails(j).ִ�п���, lng���˿���ID, Get��������ID, 2, , mobjBill.����ID)
                            End If
                        End If
                    End If
                    
                    'ˢ����ʾ����ִ�п���
                    If .Details(i).ִ�в���ID <> 0 Then
                        mrsUnit.Filter = "ID=" & .Details(i).ִ�в���ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(i, BillCol.ִ�п���) = mrsUnit!���� & "-" & mrsUnit!����
                        Else
                            Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                    End If
                    
                End If
            End If
        Next
    End With
End Sub

Private Function CheckItemHaveSub(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϵ�ǰ�е���Ŀ�Ƿ���д�����Ŀ
    '���:lngRow- ָ����
    '����:���˺�
    '����:2015-07-10 14:53:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    If mobjBill.Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�������� = lngRow Then
                CheckItemHaveSub = True: Exit Function
            End If
        Next
    End If
End Function



Private Function CheckInsureVerfyItem(objDetail As Detail, rsVerfyItem As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҽ�����˷�������
    '���:objDetail-��ǰ��ϸ��Ϣ
    '     rsVerfyItem-��Ҫ��������Ŀ
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2015-07-09 11:08:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, int���� As Integer
    Dim str���� As String, i As Long, lng��ҳID As Long
    Dim strҽ�Ƹ��ʽ As String
    Dim blnҽ�� As Boolean, bln���� As Boolean
    
    On Error GoTo errHandle
    
    If rsVerfyItem Is Nothing Then CheckInsureVerfyItem = True: Exit Function
    If rsVerfyItem.State <> 1 Then CheckInsureVerfyItem = True: Exit Function
    rsVerfyItem.Filter = 0
    If rsVerfyItem.RecordCount = 0 Then CheckInsureVerfyItem = True: Exit Function
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            lng����ID = Val(rptPati.Rows(i).Record(COL_����ID).Value)
            lng��ҳID = Val(rptPati.Rows(i).Record(COL_��ҳID).Value)
            int���� = Val(rptPati.Rows(i).Record(COL_����).Value)
            str���� = rptPati.Rows(i).Record(COL_����).Value
            strҽ�Ƹ��ʽ = rptPati.Rows(i).Record(COL_ҽ�Ƹ��ʽ).Value
            
            If lng����ID <> 0 And int���� <> 0 Then
                blnҽ�� = False: bln���� = False
                Call zlIsCheckMedicinePayMode(strҽ�Ƹ��ʽ, blnҽ��, bln����)
                If blnҽ�� Then
                    Set mrsMedAudit = GetAuditRecord(lng����ID, lng��ҳID)
                Else
                    Set mrsMedAudit = Nothing
                End If
                If Not mrsMedAudit Is Nothing Then
                    rsVerfyItem.Filter = "�շ�ϸĿID=" & mobjDetail.ID & " and ����=" & int����
                    If rsVerfyItem.RecordCount = 0 Then CheckInsureVerfyItem = True: Exit Function
                    
                    mrsMedAudit.Filter = "��ĿID=" & mobjDetail.ID
                    If mrsMedAudit.RecordCount = 0 Then
                        MsgBox "����:" & str���� & " δ����׼ʹ��[" & mobjDetail.���� & "]��", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If Not IsNull(mrsMedAudit!��������) Then
                        If mrsMedAudit!�������� <= 0 Then
                            MsgBox "����:" & str���� & "��ʹ��[" & mobjDetail.���� & "]�Ѵﵽ��׼��ʹ������" & FormatEx(mrsMedAudit!ʹ������ / IIf(gblnסԺ��λ, mobjDetail.סԺ��װ, 1), 5) & "��", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    CheckInsureVerfyItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDrugType(objDetail As Detail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鶾�����ͼ�ֵ����Ȩ��
    '���:objDetail-��ǰ��ϸ��Ϣ
    '����:�Ϸ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-07-09 11:28:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsҩƷ��Ϣ As ADODB.Recordset
    On Error GoTo errHandle
    If InStr(",5,6,7,", objDetail.���) = 0 Then CheckDrugType = True: Exit Function
    
    Set rsҩƷ��Ϣ = ReadҩƷ��Ϣ(objDetail.ID)
    If Not rsҩƷ��Ϣ Is Nothing Then
        If IIf(IsNull(rsҩƷ��Ϣ!�������), "", rsҩƷ��Ϣ!�������) = "����ҩ" _
            And InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then
            MsgBox """" & mobjDetail.���� & """Ϊ����ҩƷ����û��Ȩ�޶Ը���ҩƷ���ʣ�", vbInformation, gstrSysName
            Exit Function
        ElseIf IIf(IsNull(rsҩƷ��Ϣ!�������), "", rsҩƷ��Ϣ!�������) = "����ҩ" _
            And InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then
            MsgBox """" & mobjDetail.���� & """Ϊ����ҩƷ����û��Ȩ�޶Ը���ҩƷ���ʣ�", vbInformation, gstrSysName
            Exit Function
        ElseIf (IIf(IsNull(rsҩƷ��Ϣ!��ֵ����), "", rsҩƷ��Ϣ!��ֵ����) = "����" _
            Or IIf(IsNull(rsҩƷ��Ϣ!��ֵ����), "", rsҩƷ��Ϣ!��ֵ����) = "����") _
            And InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then
            MsgBox """" & mobjDetail.���� & """Ϊ���ػ򰺹�ҩƷ����û��Ȩ�޶Ը���ҩƷ���ʣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckDrugType = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function CheckInsureTheCode(objDetail As Detail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ն�����
    '���:objDetail-��ǰ��ϸ��Ϣ
    '����:���ڶ��뷵��true,���򷵻�False
    '����:���˺�
    '����:2015-07-09 11:37:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lng����ID As Long, intInsure As Integer, str���� As String
    Dim strInsures As String, strPriceGrade As String
    On Error GoTo errHandle
    
    strInsures = ""
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            lng����ID = Val(rptPati.Rows(i).Record(COL_����ID).Value)
            intInsure = Val(rptPati.Rows(i).Record(COL_����).Value)
            str���� = rptPati.Rows(i).Record(COL_����).Value
            If lng����ID <> 0 And intInsure <> 0 Then
                If InStr(strInsures & ",", "," & intInsure & ",") = 0 Then
                    If InStr(",5,6,7,", objDetail.���) > 0 Then
                        strPriceGrade = mstrҩƷ�۸�ȼ�
                    ElseIf objDetail.��� = "4" Then
                        strPriceGrade = mstr���ļ۸�ȼ�
                    Else
                        strPriceGrade = mstr��ͨ�۸�ȼ�
                    End If
                    If Not CheckMediCareItem(objDetail.ID, intInsure, objDetail.����, objDetail.��� = False, True, strPriceGrade) Then Exit Function
                    strInsures = strInsures & "," & intInsure
                End If
            End If
        End If
    Next
    CheckInsureTheCode = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function InputItemMemo(tyPati As TY_PATIINFOR) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ժҪ
    '���:tyPati-������Ϣ
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-07-09 11:37:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strժҪ As String
    On Error GoTo errHandle
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            strժҪ = ""
'            If tyPati.���� <> 0 And .Detail.����ժҪ = False Then '90304
            If .Detail.����ժҪ = False Then
                strժҪ = gclsInsure.GetItemInfo(tyPati.����, tyPati.����ID, .Detail.ID, strժҪ, 2)
            End If
            .ժҪ = strժҪ
        End With
    Next
    InputItemMemo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function




Private Function CheckAllPatiChargeWrang(Optional ByVal lngRow As Long = -1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������еĲ��˵ı���
    '���:lngRow:��ǰ��,-1��ʾ����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-07-09 14:27:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str����IDs As String
    Dim tyPati As TY_PATIINFOR
    
    mrsWarn.Filter = ""
    If mrsWarn.RecordCount = 0 Then CheckAllPatiChargeWrang = True: Exit Function
    If mlngSelPatiCount = 0 Then CheckAllPatiChargeWrang = True: Exit Function
    On Error GoTo errHandle
    
    str����IDs = ""
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            tyPati = GetPatiInforByReport(i)
            If tyPati.����ID <> 0 Then
                If InStr(str����IDs & ",", "," & tyPati.����ID & ",") = 0 Then
                    If CheckPatiChargeWrang(tyPati) = False Then Exit Function
                    str����IDs = str����IDs & "," & tyPati.����ID
                End If
            End If
        End If
    Next
    CheckAllPatiChargeWrang = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckPatiChargeWrang(tyPati As TY_PATIINFOR, _
      Optional blnSavePriceBill As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ�����˵ı���
    '���:tyPati-������Ϣ
    '����:blnSavePriceBill-�Ƿ񱣴�Ϊ���۵�
    '����:�Ϸ�����true(������ʾѡ�����),���򷵻�False
    '����:���˺�
    '����:2015-07-09 14:27:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
    Dim curTotal As Currency, curItemMoney As Currency
    Dim cur��� As Currency, str��� As String, str������� As String
    Dim rsTmp As ADODB.Recordset, cur���ն� As Currency
    Dim i As Long
    
    
    blnSavePriceBill = False
    mrsWarn.Filter = ""
    If mrsWarn.RecordCount = 0 Then CheckPatiChargeWrang = True: Exit Function
    On Error GoTo errHandle
    
    curTotal = CalcGridToTal(False) ' GetAllPatiTotal(tyPati.����ID, -1, -1, mobjBill)
    If curTotal <= 0 Then CheckPatiChargeWrang = True: Exit Function
    
    'ˢ�²���Ԥ������Ϣ
    Set rsTmp = GetMoneyInfo(tyPati.����ID, 0, True, 2)
    
    If Not rsTmp Is Nothing Then
        cur��� = Val(Nvl(rsTmp!Ԥ�����)) - Val(Nvl(rsTmp!�������))
    End If
    
    '���¶�ȡ���ն�
    cur���ն� = GetPatiDayMoney(tyPati.����ID)
    
    If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(1, tyPati.����ID)
    
    '����ȷ���Ǽ��ʱ���ʱ,�������ķ�ʽ������
    '����ǻ���ģʽ,��Ϊ�ް�ť����,��������µķ�ʽ������
    For i = 1 To mobjBill.Details.Count
        
        gbytWarn = BillingWarn(mstrPrivsOpt, tyPati.���� & IIf(tyPati.סԺ�� = "", "", "(סԺ��:" & tyPati.סԺ�� & " ����:" & tyPati.���� & ")"), mlng����ID, tyPati.���ò���, mrsWarn, cur���, cur���ն�, curTotal, _
                     tyPati.������, mobjBill.Details(i).�շ����, mobjBill.Details(i).Detail.�������, mstrWarn, , gblnPrice And gbytBilling = 1)

        
        '����:0;û�б���,����
        '     1:������ʾ���û�ѡ�����
        '     2:������ʾ���û�ѡ���ж�
        '     3:������ʾ�����ж�
        '     4:ǿ�Ƽ��ʱ���,����
        '     5.������ʾ���û�ѡ�����,��ֻ�������Ϊ���۵�
        '     str�������="CDE":�����ڱ��α�����һ�����,"-"Ϊ������𡣸÷������ڴ����ظ�����
        
        Select Case gbytWarn
        Case 2, 3 '������ʾ���û�ѡ���жϺͱ�����ʾ�����ж�
            Exit Function
        Case 1, 4   '������ʾ���û�ѡ�����,ǿ�Ƽ��ʱ���,����
            CheckPatiChargeWrang = True: Exit Function
        Case 5 '������ʾ���û�ѡ�����,��ֻ�������Ϊ���۵�
            blnSavePriceBill = True:    CheckPatiChargeWrang = True
            Exit Function
        Case Else
        End Select
    Next
    CheckPatiChargeWrang = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
'
'Private Function GetAllPatiTotal(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngӤ�� As Long, _
'    objBill As ExpenseBill) As Currency
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:��ȡָ�����˵��ݷ�Ŀ�ϼƽ��
'    '���:lng����ID-����ID
'    '     lng��ҳID-��ҳID(-1 ʱ����������)
'    '     lngӤ��-�ڼ���Ӥ���ķ���(-1ʱ,��ʾ����(��Ӥ��),0ʱ,�����˱��� >0ʱ��ʾ�ڼ���Ӥ��)
'    '     objBill-���ݶ���
'    '����:
'    '����:����ָ�����˵ĺϼƽ��
'    '����:���˺�
'    '����:2015-07-09 14:50:57
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim objBillDetail As New BillDetail
'    Dim curMoney As Currency
'
'    On Error GoTo errHandle
'    For Each objBillDetail In objBill.Details
'        curMoney = GetPatiBillRowTotal(lng����ID, objBillDetail.InComes, lng��ҳID, lngӤ��)
'    Next
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function
'
'Private Function GetPatiBillRowTotal(ByVal lng����ID As Long, objBillInComes As BillInComes, _
'    Optional ByVal lng��ҳID As Long = 1, Optional ByVal lngӤ�� As Long = -1) As Currency
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:��ȡָ���еĻ��ܽ��
'    '���:lng����ID-����ID
'    '     lng��ҳID-��ҳID(-1 ʱ����������)
'    '     lngӤ��-�ڼ���Ӥ���ķ���(-1ʱ,��ʾ����(��Ӥ��),0ʱ,�����˱��� >0ʱ��ʾ�ڼ���Ӥ��)
'    '     objBillInComes-���ݶ����ж���
'    '����:���ص���ָ���еĺϼƽ��
'    '����:���˺�
'    '����:2015-07-09 15:37:00
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim objBillIncome As New BillInCome
'    Dim curMoney As Currency
'    On Error GoTo errHandle
'
'    For Each objBillIncome In objBillInComes
'        curMoney = GetPatiʵ�ս��(lng����ID, objBillIncome, lng��ҳID, lngӤ��)
'        GetPatiBillRowTotal = GetPatiBillRowTotal + curMoney
'    Next
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function
'
'
'Private Function GetPatiʵ�ս��(ByVal lng����ID As Long, _
'    objBillIncome As BillInCome, _
'    Optional ByVal lng��ҳID As Long = -1, _
'    Optional ByVal lngӤ�� As Long = -1) As Currency
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:��ȡָ���е��ݽ��
'    '���:lng����ID-����ID
'    '     lng��ҳID-��ҳID(-1��������)
'    '     lngӤ��-�ڼ���Ӥ���ķ���(-1ʱ,��ʾ����(��Ӥ��),0ʱ,�����˱��� >0ʱ��ʾ�ڼ���Ӥ��)
'    '     objBill-���ݶ���
'    '����:
'    '����:����ָ�����˵ĺϼƽ��
'    '����:���˺�
'    '����:2015-07-09 14:54:42
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cllData As Collection, i As Long
'    On Error GoTo errHandle
'    If objBillIncome Is Nothing Then Exit Function
'    If UCase(TypeName(objBillIncome.Tag)) = UCase("Empty") Then Exit Function
'    If UCase(TypeName(objBillIncome.Tag)) <> UCase("Collection") Then Exit Function
'
'    '����ID,��ҳID,�ڼ���Ӥ��(0ʱ�����˱���),ʵ�ս��
'    Set cllData = objBillIncome.Tag
'    For i = 1 To cllData.Count
'       If cllData(i)(0) = lng����ID _
'            And (cllData(i)(1) = lng��ҳID Or lng��ҳID = -1) _
'            And (cllData(i)(2) = lngӤ�� Or lngӤ�� = -1) Then
'            GetPatiʵ�ս�� = GetPatiʵ�ս�� + Val(cllData(i)(3))
'       End If
'    Next
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������������Ч��
    '����: ������Ч����true,���򷵻�False
    '����:���˺�
    '����:2015-07-10 10:58:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, dbl���� As Double, strTmp As String
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim colStock As Collection, dblTotal As Double
    Dim str�շ�ϸĿIDs As String, rsVerfyItem As ADODB.Recordset
    Dim tyPati  As TY_PATIINFOR
    
    On Error GoTo errHandle
    
    mstr����IDs = GetPatiIDsBySel(mlngSelPatiCount)
    '����Ƿ�ѡ���˲���
    If mlngSelPatiCount = 0 Then
        MsgBox "û��ѡ��Ҫ�������ʵĲ���,��ѡ����ٵ��ȷ���ݣ�", vbInformation, gstrSysName
        If rptPati.Visible Then rptPati.SetFocus
        Exit Function
    End If
    
    If mobjBill.Details.Count = 0 Then
        MsgBox "������û���κ�����,����ȷ���뵥�����ݣ�", vbInformation, gstrSysName
        If Bill.Visible And Bill.Enabled Then Bill.SetFocus
        Exit Function
    End If
            

    i = Checkִ�п���
    If i <> 0 Then
        MsgBox "�����е� " & i & " ����Ŀû��ָ��ִ�п��ң�", vbInformation, gstrSysName
        If Bill.Visible And Bill.Enabled Then Bill.SetFocus
        Exit Function
    End If
    
    If mblnNurseStation Then
        For i = 0 To rptPati.Rows.Count - 1
            If rptPati.Rows(i).Record.Tag = "1" Then
                tyPati = GetPatiInforByReport(i, mblnNurseStation)
                If tyPati.������ = "" And gbln������ Then
                    MsgBox "��ȷ������" & tyPati.���� & "�Ŀ����ˣ�", vbInformation, gstrSysName
                    Exit Function
                End If
                If Val(tyPati.��������ID) = 0 Then
                    MsgBox "��ȷ������" & tyPati.���� & "�Ŀ������ң�", vbInformation, gstrSysName
                    Exit Function
                End If
                If mbln���� Then
                    If mlngDeptID <> Val(tyPati.��������ID) Then
                        MsgBox "ע��:" & vbCrLf & "    �������Ҳ��ǲ���" & tyPati.���� & "ת�ƵĿ���,���ܽ��в��Ѳ���!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    
    If mobjBill.��������ID = 0 And cbo��������.Visible Then
        MsgBox "��ȷ���������ң�", vbInformation, gstrSysName
        If cbo��������.Enabled And cbo��������.Visible Then cbo��������.SetFocus
        Exit Function
    End If
    If mobjBill.������ = "" And gbln������ And cbo������.Visible Then
        MsgBox "�����뿪���ˣ�", vbInformation, gstrSysName
        If cbo������.Enabled And cbo������.Visible Then cbo������.SetFocus
        Exit Function
    End If
    
    '��ʿ���:�жϷǷ�����
    If CheckInhibitiveByNurse(mobjBill, mrs������) Then
        MsgBox "��ʿֻ���������Ƽ�������Ŀ,�������д����������͵���Ŀ��", vbInformation, gstrSysName
        If Bill.Visible And Bill.Enabled Then Bill.SetFocus
        Exit Function
    End If
        
    '����ʱ����
    If Not IsDate(txtDate.Text) Then
        MsgBox "��������ȷ�ķ������ڣ�", vbInformation, gstrSysName
        If txtDate.Enabled And txtDate.Visible Then txtDate.SetFocus
        Exit Function
    End If
    
    '�����˼���������
    If CheckAllPatiIsValied = False Then Exit Function
    
    '���Ѽ��
    If mbln���� And cbo��������.Visible Then
        If cbo��������.ItemData(cbo��������.ListIndex) <> mlngDeptID Then
            MsgBox "ע��:" & vbCrLf & "    �������Ҳ��ǲ���ת�ƵĿ���,���ܽ��в��Ѳ���!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
            
    
    '����ĸ������
    If CheckIsExistsNegativeNums = False Then Exit Function

    mblnSendMateria = False
    dbl���� = 0: strTmp = ""
    For i = 1 To mobjBill.Details.Count
        If InStr(1, str�շ�ϸĿIDs & ",", "," & mobjBill.Details(i).�շ�ϸĿID & ",") = 0 Then
            str�շ�ϸĿIDs = str�շ�ϸĿIDs & "," & mobjBill.Details(i).�շ�ϸĿID
        End If
        If mobjBill.Details(i).���� <> 0 And dbl���� = 0 Then
            dbl���� = mobjBill.Details(i).����
        End If
        If mobjBill.Details(i).�շ�ϸĿID = 0 Then
            MsgBox "�����е� " & i & " ��û����ȷ��������,��������ɾ�����У�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Function
        ElseIf InStr(1, ",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
            '�ռ�ҩƷ�ķ�ҩҩ��
            strTmp = strTmp & "," & mobjBill.Details(i).�շ�ϸĿID
        End If
        
        
        '�������ò��˲�������
        If InStr(",5,6,7,", mobjBill.Details(i).�շ����) = 0 Then
            If CheckItemHaveSub(i) Then
                If Not CheckFeeItemLimitDept(mobjBill.Details(i).�շ�ϸĿID, IIf(mbytUseType = 2, UserInfo.����ID, mobjBill.����ID), IIf(mbytUseType = 2, UserInfo.����ID, mobjBill.����ID)) Then
                    If mbytUseType = 2 Then
                        MsgBox "��" & i & "�е��շ���Ŀ�������ڵĿ��Ҳ����ã�", vbInformation, gstrSysName
                    Else
                        MsgBox "��" & i & "�е��շ���Ŀ�Ե�ǰ���˲����Ϳ��Ҳ����ã�", vbInformation, gstrSysName
                    End If
                    Bill.Row = i: Bill.MsfObj.TopRow = i
                    Bill.Col = BillCol.��Ŀ: Bill.SetFocus
                    Exit Function
                End If
            End If
        End If
        
        '��������ʱ��ҩƷͬһҩ���Ƿ����ظ�����
        With mobjBill.Details(i)
            If (.Detail.���� Or .Detail.���) _
                And (InStr(",5,6,7,", .�շ����) > 0 Or .�շ���� = "4" And .Detail.��������) Then
                For j = 1 To mobjBill.Details.Count
                    If i <> j And .�շ�ϸĿID = mobjBill.Details(j).�շ�ϸĿID And .ִ�в���ID = mobjBill.Details(j).ִ�в���ID Then
                        If .�շ���� = "4" Then
                            MsgBox "�� " & j & " �еķ�����ʱ����������""" & .Detail.���� & """��ͬһ�����ϲ��ű��ظ����룬��ϲ���", vbInformation, gstrSysName
                        Else
                            MsgBox "�� " & j & " �еķ�����ʱ��ҩƷ""" & .Detail.���� & """��ͬһ��ҩ�����ظ����룬��ϲ���", vbInformation, gstrSysName
                        End If
                        Exit Function
                    End If
                Next
            End If
        End With
        
        '����Զ���ҩ
        If CheckAutoSendDrugAndStuff(mobjBill.Details(i), False, mblnSendMateria) = False Then Exit Function
    Next
    If InStr(mstrPrivsOpt, ";ҩƷ��ҩ;") = 0 Then mblnSendMateria = False
    
    '27467,52828
    If FormatEx(dbl����, 7) = 0 Then
        MsgBox "����������Ҫ��һ����Ϊ�������,���飡", vbInformation, gstrSysName
        Bill.SetFocus: Exit Function
    End If
    
    '���ҩƷ�ķ�ҩҩ����Ӧ�ķ������(�洢�ⷿ)
    If strTmp <> "" Then
        strTmp = Mid(strTmp, 2)
        Set rsTmp = GetServiceDept(strTmp)
        
        If Not rsTmp Is Nothing Then
            strTmp = ""
            For i = 1 To mobjBill.Details.Count
            
                
                If InStr(1, ",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
                    strInfo = mobjBill.Details(i).�շ�ϸĿID
                    '�ȼ���Ƿ�������Ĵ洢�ⷿ
                    rsTmp.Filter = "�շ�ϸĿID=" & strInfo & " And ִ�п���id=" & mobjBill.Details(i).ִ�в���ID
                    If rsTmp.RecordCount = 0 Then
                        strTmp = strTmp & "," & i
                    Else
                        '�ټ���Ƿ�������ķ������(û�����÷�����ҵ�,��������IDΪ��)
                        rsTmp.Filter = "(" & rsTmp.Filter & " And ��������ID=" & mobjBill.����ID & ") Or (" & rsTmp.Filter & " And ��������ID=0)"
                        If rsTmp.RecordCount = 0 Then
                            strTmp = strTmp & "," & i
                        End If
                    End If
                End If

 
            Next
            If strTmp <> "" Then
                strTmp = Mid(strTmp, 2)
                MsgBox "����,��" & strTmp & "��ҩƷ�Ƿ�Υ�����¹���:" & vbCrLf & vbCrLf & _
                    "A.ѡ���ִ�п��Ҳ���ҩƷ�Ĵ洢�ⷿ" & vbCrLf & _
                    "B.���˿���[" & GET��������(mobjBill.����ID, mrs��������) & "]������ҩƷ�ڴ˴洢�ⷿ�ķ������.", _
                    vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    '���в�����Ŀ
    i = CheckDuty(, True)
    If i > 0 Then
        Bill.Row = i: Bill.MsfObj.TopRow = i
        Bill.Col = BillCol.��Ŀ: Bill.SetFocus
        Exit Function
    End If
 

    
    'ҩƷ���ɼ��
    strInfo = CheckDisable(mobjBill)
    If strInfo <> "" Then
        If strInfo Like "*(�������)*" Then
            MsgBox strInfo, vbInformation, gstrSysName
            Exit Function
        End If
        If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
                    
    '�����������
    If Not gbln�������� And mbln����������� Then
        If Not CheckLimit(mobjBill, , gblnסԺ��λ) Then Exit Function
    End If
    
    '��ȡҽ��������Ҫ�����ķ���
    If str�շ�ϸĿIDs <> "" Then
        str�շ�ϸĿIDs = Mid(str�շ�ϸĿIDs, 2)
        Call GetҪ������(str�շ�ϸĿIDs, rsVerfyItem)
    End If
                
    
    'ҩƷ�����,71188:������,2014-04-03,�Բ������ѵ�ҲҪ���м��
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
            If InStr(",5,6,7,", .�շ����) > 0 Then
                If .Detail.���� Or .Detail.��� Then
                    dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                    .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                    If gblnסԺ��λ Then .Detail.��� = .Detail.��� / .Detail.סԺ��װ
             
                ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                    dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                    .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                    If gblnסԺ��λ Then .Detail.��� = .Detail.��� / .Detail.סԺ��װ
                    If dblTotal > .Detail.��� Then
                        MsgBox "�� " & i & " ��ҩƷ""" & .Detail.���� & _
                            """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                        Exit Function
                    End If
                ElseIf colStock("_" & .ִ�в���ID) = 1 Then
                    dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                    .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                    If gblnסԺ��λ Then .Detail.��� = .Detail.��� / .Detail.סԺ��װ
                    
                    If dblTotal > .Detail.��� Then
                        If MsgBox("�� " & i & " ��ҩƷ""" & .Detail.���� & _
                            """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """,Ҫ������?", vbInformation + vbYesNo, gstrSysName) = vbNo Then Exit Function
                    End If
                End If
            ElseIf .�շ���� = "4" And .Detail.�������� Then
                If .Detail.���� Or .Detail.��� Then
                    dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                    .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                    If dblTotal > .Detail.��� Then
                        MsgBox "�� " & i & " ��ʱ�ۻ������������""" & .Detail.���� & _
                            """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """��", vbInformation, gstrSysName
                        Exit Function
                    End If
                ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                    dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                    .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                    
                    If dblTotal > .Detail.��� Then
                        MsgBox "�� " & i & " ����������""" & .Detail.���� & _
                            """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                        Exit Function
                    End If
                ElseIf colStock("_" & .ִ�в���ID) = 1 Then
                    dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                    .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                    If dblTotal > .Detail.��� Then
                        If MsgBox("�� " & i & " ����������""" & .Detail.���� & _
                            """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """,Ҫ������?", vbInformation + vbYesNo, gstrSysName) = vbNo Then Exit Function
                    End If
                End If
            End If
            
            '���ҽ������
            If Not CheckInsureVerfyItem(.Detail, rsVerfyItem) Then Exit Function
        End With
    Next
    
    '���ۼ��,105875
    If Not gobjPublicDrug Is Nothing Then
        'Private Function zlCheckPriceAdjustBySell(ByVal lngҩƷid As Long, ByVal lngҩ��id As Long) As Boolean
        '���۹���ģʽʱ���жϼ۸��Ƿ��������۹���Ҫ���ɱ��ۺ��ۼ�һ�£�
        '����ҩƷ���ۼ��ǹ̶��ģ��Ƚ�����ҩ���ĳɱ��ۣ�������ڲ�һ�µľͲ������۳���
        'ʱ��ҩƷ���Ƚ�ҩ������¼�����ۼۺͳɱ��ۣ�������ڲ�һ�µľͲ������۳���
        '���۳���ʱֻ�ж�ҩ��
        '���أ�True-�����������۳��⣻false-���ܽ������۳���
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If InStr(",5,6,7,", .�շ����) > 0 Then
                    If gobjPublicDrug.zlCheckPriceAdjustBySell(.�շ�ϸĿID, .ִ�в���ID) = False Then
                        Exit Function
                    End If
                End If
            End With
        Next
    End If
         
    If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 1, _
        MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling)) = False Then
        Exit Function
    End If
        
    '���˺�:22441,����������͸����������
    If CheckMainOperation = False Then Exit Function
    If mblnSendMateria And gbytSendMateria = 2 Then
        If MsgBox("������ɺ��Զ�ִ�з�ҩ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            mblnSendMateria = False
        End If
    End If
    
    mblnPrintDrugList = False
    If mblnSendMateria Then
        mblnPrintDrugList = MsgBox("���ݷ�ҩ��ɣ�Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
    
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function CheckPatiIsValied(tyPati As TY_PATIINFOR, _
    objBill As ExpenseBill) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡�˵��ݵ���Ч��
    '���:typati-������Ϣ
    '     objBill-������Ϣ
    '����:��Ч����true,���򷵻�False
    '����:���˺�
    '����:2015-07-09 16:42:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim lngInsure As Long
    
    
    '1.����ʱ����ؼ��
    If CheckPatiFeeDateIsValied(tyPati) = False Then Exit Function
    
    '2.���Ѽ���Ƿ񳬹�ʱ��
    If mbln���� Then
        If zlCheckPatiFeeRenewValied(tyPati.����ID, tyPati.��ҳID, mobjBill.����ID, mobjBill.����ID, mstr���ת��ʱ��) = False Then Exit Function
        
        If txtDate.Text > mstr���ת��ʱ�� And mstr���ת��ʱ�� <> "" Then
            MsgBox "ע��:" & vbCrLf & _
                   "    ����:" & tyPati.���� & " ��¼�ķ���ʱ�䳬�������ת����ʱ��(" & mstr���ת��ʱ�� & "),���ܽ��в��Ѳ���!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
             Exit Function
        End If
    End If
    
    '3.ҽ�����
    If InsureItemCheck(tyPati, mobjBill) = False Then Exit Function
    
    '4.��鲡���Ƿ��ܽ��м���:��Ժǿ�Ƽ���Ȩ�޼��
    If Not PatiCanBilling(tyPati.����ID, tyPati.��ҳID, mstrPrivsOpt) Then Exit Function

    '5.��鲡���Ƿ�����䶯
    If zlIsAllowFeeChange(tyPati.����ID, tyPati.��ҳID, , tyPati.����) = False Then Exit Function
    '6.��鲡���Ƿ��Ѿ���Ŀ
    If zlPatiIS�����ѱ�Ŀ(tyPati.����ID, tyPati.��ҳID) = True Then Exit Function
    
    '7.����Ƿ�����
    '   Ҫ������,ҽ��������Ŀ�Ƿ��������,����ʱ�Ѽ�飬����ʱ�ټ������Ϊ��
    '   1).���䵥����ȷ��ҽ����ݣ�2).�������������ʱֻ��������3).���뵥��ʱδ���,
    '   4).ͨ���䷽����ʱδ���
    If tyPati.���� <> 0 And Not mrsMedAudit Is Nothing Then
        lngInsure = tyPati.����
        If Not CheckExamine(mobjBill.Details, mrsMedAudit, lngInsure, tyPati.����) Then Exit Function
    End If
    
    
    CheckPatiIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub reSetBillObject(tyPatiInfor As TY_PATIINFOR, _
    ByRef objBill As ExpenseBill)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������õ��ݶ���
    '���:tyPatiInfor-������Ϣ
    '����:���˺�
    '����:2015-07-09 17:17:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With tyPatiInfor
        objBill.��ʶ�� = tyPatiInfor.סԺ��
        objBill.����ID = tyPatiInfor.����ID
        objBill.��ҳID = tyPatiInfor.��ҳID
        objBill.���� = tyPatiInfor.����
        objBill.�ѱ� = tyPatiInfor.�ѱ�
        objBill.���� = tyPatiInfor.����
        objBill.�Ա� = tyPatiInfor.�Ա�
        objBill.���� = tyPatiInfor.����
        objBill.Ӥ���� = tyPatiInfor.Ӥ��
        '���¼���ʵ�ս��
        
    End With
End Sub

Private Function GetPatiInforByReport(ByVal lngRow As Long, Optional blnNurseStation As Boolean = False) As TY_PATIINFOR
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ���ָ�����л�ȡ������Ϣ
    '���:lngRow-ָ������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-07-09 17:22:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tyPati As TY_PATIINFOR
    
    On Error GoTo errHandle
    With rptPati.Rows(lngRow)
        tyPati.����ID = Val(.Record(COL_����ID).Value)
        tyPati.��ҳID = Val(.Record(COL_��ҳID).Value)
        tyPati.���� = Val(.Record(COL_����).Value)
        tyPati.Ӥ�� = Val(.Record(COL_Ӥ��).Value)
        tyPati.���� = .Record(COL_����).Value
        tyPati.�Ա� = .Record(COL_�Ա�).Value
        tyPati.���� = .Record(COL_����).Value
        tyPati.�ѱ� = .Record(COL_�ѱ�).Value
        tyPati.���ն� = Val(.Record(COL_���ն�).Value)
        tyPati.ʣ��� = Val(.Record(COL_ʣ���).Value)
        tyPati.������ = Val(.Record(COL_������).Value)
        tyPati.���ò��� = .Record(COL_���ò���).Value
        tyPati.סԺ�� = .Record(COL_סԺ��).Value
        tyPati.���� = .Record(COL_����).Value
        tyPati.������� = .Record(COL_�������).Value
        tyPati.ҽ�Ƹ��ʽ = .Record(COL_ҽ�Ƹ��ʽ).Value
        tyPati.��Ժ���� = ""
        tyPati.��Ժ���� = ""
        tyPati.״̬ = 0
        If Not mrsPati Is Nothing Then
            mrsPati.Filter = "����ID=" & tyPati.����ID
            If Not mrsPati.EOF Then
                tyPati.��Ժ���� = Format(mrsPati!��Ժ����, "yyyy-MM-DD HH:MM:SS")
                tyPati.��Ժ���� = Format(mrsPati!��Ժ����, "yyyy-MM-DD HH:MM:SS")
                tyPati.״̬ = Val(Nvl(mrsPati!״̬))
                tyPati.�������� = Val(Nvl(mrsPati!��������))
            End If
        End If
        If blnNurseStation = True Then
            tyPati.������ = .Record(COL_������).Value
            tyPati.��������ID = .Record(COL_��������ID).Value
        End If
    End With
    GetPatiInforByReport = tyPati
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InsureItemCheck(tyPati As TY_PATIINFOR, _
    ByVal objBill As ExpenseBill) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ����ؼ��
    '���:objBill-Ʊ�ݶ���
    '����:���˺�
    '����:2015-07-09 16:48:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If tyPati.���� = 0 Then InsureItemCheck = True: Exit Function
    If Not gclsInsure.GetCapability(supportʵʱ���, tyPati.����ID, tyPati.����) Then InsureItemCheck = True: Exit Function

    On Error GoTo errHandle
    If gclsInsure.CheckItem(tyPati.����, 1, 0, MakeDetailRecord(objBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, gbytBilling)) = False Then Exit Function
    InsureItemCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 Private Function CheckIsExistsNegativeNums() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:������true,���򷵻�False
    '����:���˺�
    '����:2015-07-10 10:36:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnIsNegative As Boolean '���ڸ���
    Dim dblNums As Double, lngTempRow As Long
    Dim bln�������� As Boolean
    Dim i As Long
    
    On Error GoTo errHandle

    lngTempRow = 0
    With mobjBill
        For i = 1 To .Details.Count
            blnIsNegative = .Details(i).���� * .Details(i).���� < 0
            lngTempRow = i: If blnIsNegative Then Exit For
        Next
    End With
    If blnIsNegative Then
        MsgBox "�������ʲ�������и�������(��" & lngTempRow & "��)��", vbInformation, gstrSysName
        If Bill.Rows - 1 >= Bill.Row Then Bill.Row = lngTempRow
    End If
    CheckIsExistsNegativeNums = True: Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDrugStoreIsEnough(ByVal dblNums As Double, _
    objBillDetail As BillDetail, Optional ByVal bln�ض���� As Boolean = False, _
    Optional lngRow As Long = -1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҩƷ�Ƿ����
    '���:dblNums-ҩƷ����
    '    objBillDetail-ҩƷ��ϸ
    '����:���㷵�ط���true,���򷵻�False
    '����:���˺�
    '����:2015-07-10 11:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTempNums As Double
    Dim colStock As Collection
    
    On Error GoTo errHandle
    
    dblTempNums = dblNums * IIf(mlngSelPatiCount = 0, 1, mlngSelPatiCount)
    
    
    With objBillDetail
        If Not (.�շ���� = "4" And .Detail.��������) Or (InStr(",5,6,7,", .�շ����) > 0) Then CheckDrugStoreIsEnough = True: Exit Function
        
        If dblNums = 0 Then
            dblNums = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
            dblTempNums = dblNums * IIf(mlngSelPatiCount = 0, 1, mlngSelPatiCount)
        End If
        
        If .Detail.���� Or .Detail.��� Then
            '������ʱ��ҩƷ�����ֹ����
            If bln�ض���� Then
                .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                If gblnסԺ��λ Then .Detail.��� = .Detail.��� / .Detail.סԺ��װ
            End If
            If dblTempNums <= .Detail.��� Then CheckDrugStoreIsEnough = True: Exit Function
            
            If .�շ���� = "4" Then
                If lngRow > 0 Then
                    MsgBox "�� " & lngRow & " ��ʱ�ۻ������������""" & .Detail.���� & _
                        """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTempNums & """��", vbInformation, gstrSysName
                    Exit Function
                Else
                    MsgBox """" & .Detail.���� & """Ϊ������ʱ����������,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                End If
            Else
                 If lngRow > 0 Then
                    MsgBox "�� " & lngRow & " ��ʱ�ۻ����ҩƷ""" & .Detail.���� & _
                        """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTempNums & """��", vbInformation, gstrSysName
                    Exit Function
                 Else
                    MsgBox """" & .Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                 End If
            End If
            Exit Function
        End If
    
        Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
        If colStock("_" & .ִ�в���ID) <> 0 _
            And Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus Then
            
            If dblTempNums <= .Detail.��� Then CheckDrugStoreIsEnough = True: Exit Function
        
            If colStock("_" & .ִ�в���ID) = 1 Then
                If MsgBox("""" & .Detail.���� & """�ĵ�ǰ���ÿ�治����������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                MsgBox """" & .Detail.���� & """�ĵ�ǰ���ÿ�治������������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End With
    CheckDrugStoreIsEnough = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckAllPati������(ByVal dblNum As Double, objBillDetail As BillDetail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������
    '���:dblNum-������
    '     objBillDetail-������ϸ����
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2015-07-09 14:27:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInsures As String, tyPati As TY_PATIINFOR
    Dim i As Long
    
    On Error GoTo errHandle
    strInsures = ""
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            tyPati = GetPatiInforByReport(i)
            If tyPati.���� <> 0 And InStr(strInsures & ",", "," & tyPati.���� & ",") = 0 Then
              
              If CheckPati������(dblNum, tyPati, objBillDetail) = False Then Exit Function
              strInsures = strInsures & "," & tyPati.����
    
            End If
        End If
    Next
    CheckAllPati������ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckPati������(ByVal dblNum As Double, tyPati As TY_PATIINFOR, objBillDetail As BillDetail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ�����˵�������
    '���:dblNum-������
    '     objBillDetail-������ϸ����
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2015-07-10 11:23:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If tyPati.���� = 0 Then CheckPati������ = True: Exit Function
    If mrsMedAudit Is Nothing Then CheckPati������ = True: Exit Function
    With objBillDetail
        If Not .Detail.Ҫ������ Then CheckPati������ = True: Exit Function
        mrsMedAudit.Filter = "��ĿID=" & .�շ�ϸĿID
        If mrsMedAudit.RecordCount = 0 Then CheckPati������ = True: Exit Function
        If IsNull(mrsMedAudit!��������) Then CheckPati������ = True: Exit Function
        If dblNum > mrsMedAudit!�������� Then
            MsgBox "��������γ�������׼�Ŀ�������" & FormatEx(mrsMedAudit!�������� / IIf(gblnסԺ��λ, .Detail.סԺ��װ, 1), 5) & "��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    CheckPati������ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckMainOperation() As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������������(�����������Ҫ����,�����ڸ�������,���ֹ
    '���:
    '����:lngRow-���ظ�����������
    '����:������������û�����븽������,����true,���򷵻�False
    '����:
    '�޸�:���˺�(�˺�ʱ,���Ӷ�λ����),���Ӳ���;strBackNo
    '����:2009/7/10
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long, lngRow As Long   'ָ����
    Dim i As Long
    
    lngCount = 0
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�շ���� = "F" Then
           If mobjBill.Details(i).���ӱ�־ = 0 Then CheckMainOperation = True: Exit Function     '������Ҫ����,�򲻼��,ֱ�ӷ���true
           lngCount = lngCount + 1  '��ʾ��������
           If lngRow <= 0 Then lngRow = i
        End If
    Next
    If lngCount <> 0 Then
          MsgBox "�����в�����Ҫ����,�����ڸ�������,���飡", vbInformation, gstrSysName
          If Bill.Rows > lngRow Then Bill.Row = lngRow
          If Bill.Visible Then Bill.SetFocus
          Exit Function
    End If
    CheckMainOperation = True
End Function





  
 

Private Sub Calc��������ʵ��(ByVal lngMainRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ۿ�ʱ,����ָ�����������ID�ĵ�һ��������Ŀ���������ʵ�ս��
    '���:lngMainRow-������ID
    '����:���˺�
    '����:2015-07-10 12:01:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim cur����ǰӦ�պϼ� As Currency     '��¼�����������Ӧ�պϼ�
    Dim cur���ۺ�ʵ�� As Currency
    
    With mobjBill
        For i = lngMainRow To .Details.Count
            If i = lngMainRow Or .Details(i).�������� = lngMainRow Then
                For j = 1 To .Details(i).InComes.Count
                    cur����ǰӦ�պϼ� = cur����ǰӦ�պϼ� + .Details(i).InComes(j).Ӧ�ս��
                Next
            End If
        Next
        
        cur���ۺ�ʵ�� = CCur(Format(ActualMoney(.�ѱ�, .Details(lngMainRow).InComes(1).������ĿID, cur����ǰӦ�պϼ�, 0, 0, 0, 0), gstrDec))
        cur���ۺ�ʵ�� = cur���ۺ�ʵ�� - cur����ǰӦ�պϼ� + .Details(lngMainRow).InComes(1).Ӧ�ս��
        .Details(lngMainRow).InComes(1).ʵ�ս�� = Format(cur���ۺ�ʵ��, gstrDec)
        
        Call ShowDetails(lngMainRow)
    End With
End Sub
Private Function CheckPatiFeeDateIsValied(tyPati As TY_PATIINFOR) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡�˵ķ��������Ƿ���Ч
    '���:tyPati-������Ϣ
    '����:��Ч����true,���򷵻�False
    '����:���˺�
    '����:2015-07-17 14:45:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strDate As String
    On Error GoTo errHandle
    strDate = Format(CDate(txtDate.Text), "yyyy-MM-dd HH:mm:ss")
    '��鷢��ʱ�䲻�����ڲ��˵���Ժʱ��
    If strDate < tyPati.��Ժ���� And tyPati.��Ժ���� <> "" Then
        MsgBox "���õķ���ʱ�䲻��С�ڲ���""" & tyPati.���� & """����Ժʱ��:" & tyPati.��Ժ���� & "��", vbInformation, gstrSysName
        If txtDate.Enabled And txtDate.Visible Then txtDate.SetFocus
        Exit Function
    End If
    '����ʱ����
    If strDate > tyPati.��Ժ���� And tyPati.��Ժ���� <> "" Then
        MsgBox "ǿ�ƶԳ�Ժ����(" & tyPati.���� & ")����ʱ������ʱ�䲻�ܴ��ڲ��˳�Ժʱ��:" & tyPati.��Ժ����, vbInformation, gstrSysName
        If txtDate.Enabled And txtDate.Visible Then txtDate.SetFocus
        Exit Function
    End If
    
    If tyPati.���� <> 0 And strDate < tyPati.��Ժ���� And tyPati.��Ժ���� <> "" Then
        MsgBox "���õķ���ʱ�䲻��С��ҽ�����˵���Ժʱ��(" & tyPati.���� & "):" & tyPati.��Ժ����, vbInformation, gstrSysName
        Exit Function
    End If
    CheckPatiFeeDateIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function CheckAllPatiIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ���鷢��ʱ���Ƿ�Ϸ�
    '���أ����ݺϷ�������true,���򷵻�False
    '����:���˺�
    '����:2015-07-10 15:47:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, varData As Variant
    Dim dtDate As Date, strMsg As String
    Dim strYBItemIDs As String, strGFItemIDs As String
    Dim strҽ�Ƹ��ʽ As String, tyPati As TY_PATIINFOR
    Dim blnҽ�� As Boolean, bln���� As Boolean
    
    On Error GoTo errH
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            tyPati = GetPatiInforByReport(i)
            '��ʼ��ҽ������
            If tyPati.���� <> 0 Then Call InitInsurePara(tyPati.����ID, tyPati.����)
            '���¸����ݸ�ֵ
            Call reSetBillObject(tyPati, mobjBill)
            
            '������صļ��
            blnҽ�� = False: bln���� = False
            Call zlIsCheckMedicinePayMode(tyPati.ҽ�Ƹ��ʽ, blnҽ��, bln����)
            If blnҽ�� Then
                Set mrsMedAudit = GetAuditRecord(tyPati.����ID, tyPati.��ҳID)
            Else
                Set mrsMedAudit = Nothing
            End If
                
            If CheckPatiIsValied(tyPati, mobjBill) = False Then Exit Function
                          
            If InStr(strҽ�Ƹ��ʽ & "','", "','" & tyPati.ҽ�Ƹ��ʽ & "','") = 0 And tyPati.ҽ�Ƹ��ʽ <> "" Then
                strҽ�Ƹ��ʽ = strҽ�Ƹ��ʽ & "','" & tyPati.ҽ�Ƹ��ʽ
                'ҽ���򹫷Ѳ�������:45605
                If zlIsCheckMedicinePayMode(tyPati.ҽ�Ƹ��ʽ) Then
                    '����ְ����
                    i = CheckDuty(, False, tyPati.����)
                    If i > 0 Then
                        Bill.Row = i: Bill.MsfObj.TopRow = i
                        Bill.Col = BillCol.��Ŀ: Bill.SetFocus
                        Exit Function
                    End If
                End If
                If Check��������(tyPati.ҽ�Ƹ��ʽ, , strYBItemIDs, strGFItemIDs, tyPati.����) = False Then Exit Function
            End If
            If Check�������(tyPati.��������, tyPati.����) > 0 Then Exit Function
        End If
    Next
    CheckAllPatiIsValied = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function Check��������(ByVal strҽ�Ƹ��ʽ As String, _
    Optional intRow As Integer, _
    Optional strYBItemIDs As String, _
    Optional strGFItemIDs As String, _
    Optional str���� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ���˵������ж�ָ���е���Ŀ�Ƿ��������,����������������Ŀ
    '���:intRow-ָ����
    '����: strYBItemIDs-�Ѿ�����ҽ��������Ŀ,����ö��ŷ���
    '      strGFItemIDs-�Ѿ����Ĺ��Ѳ�����Ŀ,����ö��ŷ���
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2015-07-10 17:24:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim strSQL As String, bytType As Byte
    Dim i As Integer
    Dim blnҽ�� As Boolean, bln���� As Boolean
    
    Check�������� = True
    
    On Error GoTo errHandle
    

    '�޷����
    If strҽ�Ƹ��ʽ = "" Then Exit Function
    
    'ҽ���򹫷Ѳ���
    '����:45605
    'ֻ���ҽ�����˺͹��Ѳ���
    If zlIsCheckMedicinePayMode(strҽ�Ƹ��ʽ, blnҽ��, bln����) = False Then Exit Function
    'ȷ����������
    bytType = IIf(blnҽ��, 1, 2)
    
    '��ȡ�������
    If mrs�������� Is Nothing Then
        strSQL = " Select 'ҽ��' As ���,����,���� From �������� Where ���� In(" & gstrҽ���������� & ") Union All " & _
                 " Select '����' As ���,����,���� From �������� Where ���� In(" & gstr���ѷ������� & ") "
        Set mrs�������� = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(mrs��������, strSQL, Me.Caption)
    End If
    mrs��������.Filter = ""
    If mrs��������.RecordCount = 0 Then Exit Function
        
    If bytType = 1 Then
        strSQL = " And ���='ҽ��'"
    Else
        strSQL = " And ���='����'"
    End If
    
    If intRow > 0 Then
        If bytType = 1 Then 'ҽ��
            If InStr("," & strYBItemIDs & ",", "," & mobjBill.Details(intRow).�շ�ϸĿID & ",") > 0 Then Exit Function
            strYBItemIDs = strYBItemIDs & "," & mobjBill.Details(intRow)
        Else
            If InStr("," & strGFItemIDs & ",", "," & mobjBill.Details(intRow).�շ�ϸĿID & ",") > 0 Then Exit Function
            strGFItemIDs = strGFItemIDs & "," & mobjBill.Details(intRow).�շ�ϸĿID
        End If
        
        If mobjBill.Details(intRow).Detail.���� = "" Then
            If InStr("," & strYBItemIDs & "," & strGFItemIDs & ",", "," & mobjBill.Details(intRow).�շ�ϸĿID & ",") > 0 Then Exit Function
            MsgBox """" & mobjBill.Details(intRow).Detail.���� & """�ķ�������δ���ã�", vbInformation, gstrSysName
            Check�������� = False
        Else
            mrs��������.Filter = "����='" & mobjBill.Details(intRow).Detail.���� & "'" & strSQL
            If mrs��������.EOF Then
                
                MsgBox """" & mobjBill.Details(intRow).Detail.���� & """�ķ�������Ϊ""" & _
                    mobjBill.Details(intRow).Detail.���� & """,����" & _
                    IIf(bytType = 1, "ҽ��", "����") & "��������" & IIf(str���� <> "", "(" & str���� & ")", "") & "��", vbInformation, gstrSysName
                Check�������� = False
            End If
        End If
    Else
        For i = 1 To mobjBill.Details.Count
            If mobjBill.Details(i).Detail.���� = "" Then
                If InStr("," & strYBItemIDs & "," & strGFItemIDs & ",", "," & mobjBill.Details(i).�շ�ϸĿID & ",") > 0 Then Exit Function
                strYBItemIDs = strYBItemIDs & "," & mobjBill.Details(i).�շ�ϸĿID
                
                If MsgBox("�����е� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """�ķ�������δ���ã�" & vbCrLf & "ȷʵҪ���浥����", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Check�������� = False: Exit For
                End If
            Else
                If bytType = 1 Then 'ҽ��
                    If InStr("," & strYBItemIDs & ",", "," & mobjBill.Details(i).�շ�ϸĿID & ",") > 0 Then Exit Function
                    strYBItemIDs = strYBItemIDs & "," & mobjBill.Details(i).�շ�ϸĿID
                Else
                    If InStr("," & strGFItemIDs & ",", "," & mobjBill.Details(i).�շ�ϸĿID & ",") > 0 Then Exit Function
                    strGFItemIDs = strGFItemIDs & "," & mobjBill.Details(i).�շ�ϸĿID
                End If
                
                mrs��������.Filter = "����='" & mobjBill.Details(i).Detail.���� & "'" & strSQL
                If mrs��������.EOF Then
                    If MsgBox("�����е� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """�ķ�������Ϊ""" & _
                        mobjBill.Details(i).Detail.���� & """,����" & _
                        IIf(bytType = 1, "ҽ��", "����") & "��������" & IIf(str���� <> "", "(" & str���� & ")", "") & "��" & vbCrLf & "ȷʵҪ���浥����", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Check�������� = False: Exit For
                    End If
                End If
            End If
        Next
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function Check�������(ByVal int�������� As Integer, _
    Optional str���� As String = "") As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ���˵ļ��ʷ�����Ŀ�ķ�������Ƿ�һ��
    '���:int��������-��������
    '���أ���һ�µķ�����,Ϊ0ʱ����
    '����:���˺�
    '����:2015-07-13 10:29:11
    '˵������Ϊ�������������۲���,�����д˼��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
        
    On Error GoTo errHandle
    For i = 1 To mobjBill.Details.Count
        If int�������� = 0 Or int�������� = 2 Then
            'סԺ���˻�סԺ���۲���,������ֻ�������������Ŀ
            If mobjBill.Details(i).Detail.������� = 1 Then
                If str���� = "" Then
                    MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """������������,�ò��˲���ʹ��.", vbInformation, gstrSysName
                Else
                    MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """������������,����:" & str���� & "����ʹ�ø���Ŀ.", vbInformation, gstrSysName
                End If
                Check������� = i: Exit Function
            End If
        ElseIf int�������� = 1 Or int�������� = -1 Then
            '������Ժ����(ҽ������)���������۲���,������ֻ������סԺ����Ŀ
            If mobjBill.Details(i).Detail.������� = 2 Then
                 If str���� = "" Then
                    MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """��������סԺ,�ò��˲���ʹ��.", vbInformation, gstrSysName
                Else
                    MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """��������סԺ,,����:" & str���� & "����ʹ�ø���Ŀ.", vbInformation, gstrSysName
                End If
                Check������� = i: Exit Function
            End If
        End If
    Next


    Check������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckAutoSendDrugAndStuff(ByVal objDetail As BillDetail, _
    ByVal blnSavePrice As Boolean, ByRef blnSendMaterial As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ��Զ���ҩ
    '���:objDetail-������ϸ
    '     blnSavePrice-�Ƿ񱣴�Ϊ���۵�
    '����:blnSendMaterial-�Ƿ��Զ���ҩ������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-07-13 10:37:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTotal As Double
    
    On Error GoTo errHandle
    With objDetail
        If .�շ���� = "4" And .Detail.�������� Then
            dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
            If Not CheckValidity(.�շ�ϸĿID, .ִ�в���ID, dblTotal) Then Exit Function
        End If
        If InStr(1, ",5,6,7,", .�շ����) > 0 Then
            '��ӡ��ҩ��,����ͨ����,�һ��۵�����
            If gbytSendMateria <> 0 And mbytUseType = 0 And gbytBilling = 0 And Not blnSavePrice Then
                'ȫ��ҩƷ��ȷ����ҩ���Ĳ��Զ���ҩ(���뷢ҩʱ,û��ȷ��ҩ��)
                blnSendMaterial = .ִ�в���ID <> 0
            End If
        End If
    End With
    CheckAutoSendDrugAndStuff = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub BillPrint(ByVal blnSavePrice As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡƱ��
    '���:blnSavePrice-�Ƿ񱣴�Ļ��۵�
    '����:���˺�
    '����:2015-07-13 10:46:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
     If gbytBilling = 0 And Not blnSavePrice And gbln���ʴ�ӡ Then
         Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_113" & 3 + mbytUseType, Me, "NO=" & mobjBill.NO, "�Ǽ�ʱ��=" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), "ҩƷ��λ=" & IIf(gblnסԺ��λ, 1, 0), "PrintEmpty=0", "�ش�=0", 2)
     ElseIf (gbytBilling = 1 Or blnSavePrice) And gbln���۴�ӡ Then
         Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & mobjBill.NO, "�Ǽ�ʱ��=" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), "ҩƷ��λ=" & IIf(gblnסԺ��λ, 1, 0), "PrintEmpty=0", "�ش�=0", 2)
     End If
    
    '��ӡ��ҩ��
    If mblnPrintDrugList Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "���ݺ�=" & mobjBill.NO, "�Ǽ�ʱ��=" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), 2)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetDrawDrugDeptEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ҩ���ŵ�Enabled����
    '����:���˺�
    '����:2015-07-13 10:57:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnHaveDrug As Boolean '����ҩƷ
    
    '���û�����ò��ŵ�ѡ��,��ֱ���˳�
    If cboDrawDept.Visible = False Then cboDrawDept.Enabled = False: lblDrawDrugDept.Enabled = False: Exit Sub
    blnHaveDrug = False
    For i = 1 To mobjBill.Details.Count
        If InStr(1, ",5,6,7,", "," & mobjBill.Details(i).�շ���� & ",") > 0 Then
            blnHaveDrug = True
            Exit For
        End If
    Next
    cboDrawDept.Enabled = blnHaveDrug: lblDrawDrugDept.Enabled = blnHaveDrug
End Sub

Private Sub SetBill�в�ҩEditEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����в�ҩ�ı༭״̬
    '����:���˺�
    '����:2015-07-13 11:02:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With Bill
        For i = 0 To .Cols - 1
            .ColData(i) = IIf(.TextMatrix(0, i) = "��Ŀ", 0, 5)
        Next
    End With
End Sub
 

Private Sub zlReSetDrawDrugDept()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ӧ�Ĺ���,���»�ȡ��ҩ����
    '����:���˺�
    '����:2015-07-13 11:39:26
    '3)  ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
    '4)  סԺ���ʡ����ҷ�ɢ���ʣ������ɲ���ʹ�ã�Ҳ������ҽ������ʹ�á�
    '    a)  �жϵ�ǰ����Ա�������ң����������ҽ�����ʵĿ��ң�����ҩ���Ź̶�Ϊ���˲�����(��顢���顢���������ơ�Ӫ��)
    '    b)  �������Ա����ҽ�����ʵĿ��ң����ڵ��ݽ���������"��ҩ����"ѡ��򣬿�ѡ��ΧΪ����Ա������ҽ�����ʵĿ���(���ܶ��)��ȱʡ�뿪��������ͬ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytUseType = 2 Then
        'ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
        mobjBill.��ҩ����ID = mlngDeptID: Exit Sub
    End If
    
    If mrs��ҩ����.RecordCount = 0 Then
        '�жϵ�ǰ����Ա�������ң����������ҽ�����ʵĿ��ң�����ҩ���Ź̶�Ϊ���˲�����(��顢���顢���������ơ�Ӫ��)
        mobjBill.��ҩ����ID = mobjBill.����ID: Exit Sub
    End If
    '�������Ա����ҽ�����ʵĿ��ң����ڵ��ݽ���������"��ҩ����"ѡ��򣬿�ѡ��ΧΪ����Ա������ҽ�����ʵĿ���(���ܶ��)��ȱʡ�뿪��������ͬ��
    If mrs��ҩ����.RecordCount = 1 Then
        'ֻ��һ������,�϶�����
        If mrs��ҩ����.EOF Then mrs��ҩ����.MoveFirst
         mobjBill.��ҩ����ID = Val(Nvl(mrs��ҩ����!ID)): Exit Sub
    End If
    'ѡ��Ŀ������ĸ������ĸ�
    With cboDrawDept
        If .ListIndex < 0 Then Exit Sub
        If mobjBill.��ҩ����ID <> .ItemData(.ListIndex) Then mobjBill.��ҩ����ID = .ItemData(.ListIndex): Exit Sub
    End With
End Sub

Private Sub zlLoadDrawDeptData(ByVal bytUseType As Byte, Optional ByVal lngDeptID As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�ҩ����
    '���:bytUseType:���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����,2-ҽ�����Ҽ���
    '����:24729,24731
    '����:���˺�
    '����:2009-07-29 15:05:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    '3)  ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
    '4)  סԺ���ʡ����ҷ�ɢ���ʣ������ɲ���ʹ�ã�Ҳ������ҽ������ʹ�á�
    '    a)  �жϵ�ǰ����Ա�������ң����������ҽ�����ʵĿ��ң�����ҩ���Ź̶�Ϊ���˲�����(��顢���顢���������ơ�Ӫ��)
    '    b)  �������Ա����ҽ�����ʵĿ��ң����ڵ��ݽ���������"��ҩ����"ѡ��򣬿�ѡ��ΧΪ����Ա������ҽ�����ʵĿ���(���ܶ��)��ȱʡ�뿪��������ͬ��
    
    On Error GoTo errHandle
    
    'ҽ������
    If bytUseType = 2 Then
        '3)  ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
        strSQL = "Select ID,����,���� From ���ű� where id=[2]"
    Else
        strSQL = _
            " Select distinct  A.ID, A.����,A.����   " & vbNewLine & _
            " From ���ű� A, ��������˵�� B,������Ա C" & vbNewLine & _
            " Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)  " & _
            "       And A.ID = B.����id and a.id=C.����ID and C.��Աid=[1] " & vbNewLine & _
            "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            "       AND B.�������� IN('���','����','����','����','Ӫ��') " & _
            " Order by ����"
    End If
    Set mrs��ҩ���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, lngDeptID)
    With mrs��ҩ����
        cboDrawDept.Clear
        Do While Not .EOF
            cboDrawDept.AddItem IIf(zlIsShowDeptCode, Nvl(!����) & "-", "") & Nvl(!����)
            cboDrawDept.ItemData(cboDrawDept.NewIndex) = Val(Nvl(!ID))
            If Val(Nvl(!ID)) = UserInfo.����ID Then cboDrawDept.ListIndex = cboDrawDept.NewIndex
            .MoveNext
        Loop
        If .RecordCount <> 0 And cboDrawDept.ListIndex < 0 Then cboDrawDept.ListIndex = 0
        If .RecordCount <> 0 Then .MoveFirst
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetDrawDrugDeptVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ҩ���ŵ�visibled����
    '����:���˺�
    '����:2009-07-29 19:07:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ' mbytUseType As Byte '���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����,2-ҽ�����Ҽ���
    '3)  ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
    If mblnNurseStation Then Exit Sub
    If mbytUseType = 2 Then
        cboDrawDept.Visible = False
    Else
        cboDrawDept.Visible = mrs��ҩ����.RecordCount > 1 And gbytBilling <> 2         '
    End If
    lblDrawDrugDept.Visible = cboDrawDept.Visible
End Sub


Private Function GetLastDeptID(ByVal str��� As String, ByVal lngRow As Long, _
    ByVal strDeptIDs As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����������ͬ�����Ŀ��ִ�п���ID
    '���:str���-�շ����
    '     lngRow-ָ����
    '     strDeptIDs-ִ�в���ID,����ö��ŷ���
    '����:�ɹ�,�������һ��ִ�в���ID,���򷵻�0
    '����:���˺�
    '����:2015-07-13 11:52:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    
    For i = lngRow - 1 To 1 Step -1
        If mobjBill.Details(i).�շ���� = str��� _
            And mobjBill.Details(i).ִ�в���ID <> 0 Then
            If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).ִ�в���ID & ",") > 0 Then
                GetLastDeptID = mobjBill.Details(i).ִ�в���ID
                Exit Function
            End If
        End If
    Next
    
    '�������������,��ȡ��������������ƥ���ִ�п���
    If str��� = "4" Then
        For i = lngRow - 1 To 1 Step -1
            If mobjBill.Details(i).ִ�в���ID <> 0 Then
                If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).ִ�в���ID & ",") > 0 Then
                    GetLastDeptID = mobjBill.Details(i).ִ�в���ID
                    Exit Function
                End If
            End If
        Next
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub FillBillComboBox(lngRow As Long, lngCol As Long, Optional blnEnter As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��������������б������
    '���:blnEnter=�Ƿ񰴽�����д���,����ִ�п��ұ��ֲ���
    '����:���˺�
    '����:2015-07-13 11:53:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String, bln��ʿ As Boolean
    Dim strSQL As String, strIDs As String, i As Long
    Dim lng����ID As Long, lng����ID As Long, j As Long
    Dim bln��ҩ��� As Boolean '�Ƿ����������ҩ���
    
    Bill.Clear
    
    On Error GoTo errHandle
    
    Select Case Bill.TextMatrix(0, lngCol)
        Case "���"
            Call GetOperatorInfo(mrs������, mobjBill.������, bln��ʿ)
            mrsClass.Filter = 0
            If mrsClass.RecordCount <> 0 Then
                mrsClass.MoveFirst
                j = 1
                For i = 1 To mrsClass.RecordCount
                    '��ʿ���:����
                    If Not (bln��ʿ And InStr(",E,M,4,", mrsClass!����) = 0) Then
                        Bill.AddItem j & "-" & mrsClass!���
                        Bill.ItemData(Bill.NewIndex) = Asc(mrsClass!����)  '����������ASCII��
                        j = j + 1
                    End If
                    mrsClass.MoveNext
                Next
            End If
            Bill.cboStyle = DropDownAndEdit  ' DropOlnyDown
        Case "ִ�п���"
            Bill.cboStyle = DropDownAndEdit
            '���ݵ�ǰ��Ŀִ�п�������,��̬���ÿ�ѡ����
            If mobjBill.Details.Count >= lngRow Then
                With mobjBill.Details(lngRow)
                    If InStr(",4,5,6,7,", .�շ����) > 0 Then
                        Call GetWorkUnit(.�շ�ϸĿID, .�շ����)
                        If mrsWork.RecordCount > 0 Then
                            'ȡ��һ��ҩ��ҩ��
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                strIDs = strIDs & "," & mrsWork!ID
                                mrsWork.MoveNext
                            Next
                            If Not blnEnter Then '�������ʱ������ȷ��ֵ����
                                lng����ID = GetLastDeptID(.�շ����, lngRow, Mid(strIDs, 2))
                            End If
                            If lng����ID = 0 Then lng����ID = .ִ�в���ID
                            
                            'ȷ����ǰ�е�ҩ��
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                Bill.AddItem IIf(zlIsShowDeptCode, mrsWork!���� & "-", "") & mrsWork!����
                                Bill.ItemData(Bill.NewIndex) = mrsWork!ID
                                If mrsWork!ID = lng����ID Then Bill.ListIndex = Bill.NewIndex
                                mrsWork.MoveNext
                            Next
                        End If
                    Else
                        Bill.TextMatrix(lngRow, lngCol) = ""
                        
                        lng����ID = mobjBill.����ID
                        If lng����ID = 0 Then lng����ID = Get��������ID
                        
                        lng����ID = mobjBill.����ID
                        If lng����ID = 0 Then lng����ID = Get����ID(lng����ID)
                        If lng����ID = 0 Then lng����ID = lng����ID
                        
                        '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
                        Select Case .Detail.ִ�п���
                            Case 0 '����ȷ
                                mrsUnit.Filter = 0
                            Case 1 '���˿���
                                mrsUnit.Filter = "ID=" & lng����ID & " Or ID=" & .ִ�в���ID
                            Case 2 '���˲���
                                mrsUnit.Filter = "ID=" & lng����ID & " Or ID=" & .ִ�в���ID
                            Case 3 '����Ա����
                                mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                            Case 4 'ָ������
                                strSQL = "" & _
                                "   Select Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                                "   From �շ�ִ�п��� A,���ű� C" & _
                                "   Where A.�շ�ϸĿID=[1]��And A.ִ�п���ID+0=C.ID " & _
                                "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                                "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
                                "       And (A.������Դ is NULL Or A.������Դ=[2])" & _
                                "       And (A.��������ID is NULL Or A.��������ID=[3])" & _
                                " Order by Decode(A.������Դ,Null,2,1)" 'Ĭ�Ͽ�������
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .�շ�ϸĿID, 2, lng����ID)
                                If Not rsTmp.EOF Then
                                    For i = 1 To rsTmp.RecordCount
                                        strTmp = strTmp & "ID=" & rsTmp!ִ�п���ID & " OR "
                                        rsTmp.MoveNext
                                    Next
                                    strTmp = strTmp & "ID=" & .ִ�в���ID & " OR "
                                    strTmp = Left(strTmp, Len(strTmp) - 4)
                                    mrsUnit.Filter = strTmp
                                Else
                                    mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                                End If
                            Case 5 'Ժ��ִ��(Ԥ��,������δ��)
                            Case 6 '�����˿���
                               mrsUnit.Filter = "ID=" & Get��������ID & " Or ID=" & .ִ�в���ID
                        End Select
                        If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                        If Not mrsUnit.EOF Then
                            For i = 1 To mrsUnit.RecordCount
                                strTmp = IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
                                '���˺�:28947
                                If zlCboFindItem(Bill.cboObj, Val(Nvl(mrsUnit!ID))) = False Then
                                'If Not (SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                    Bill.AddItem strTmp
                                    Bill.ItemData(Bill.ListCount - 1) = mrsUnit!ID
                                    
                                    '����ȱʡִ�п���
                                    If Not blnEnter Then '�������ʱ������ȷ��ֵ����
                                        If lngRow = 1 Then
                                            If mrsUnit!ID = lng����ID Then Bill.ListIndex = Bill.NewIndex
                                        ElseIf lngRow > 1 Then
                                            '����һ�з�ҩƷ��ͬ
                                            If mrsUnit!ID = mobjBill.Details(lngRow - 1).ִ�в���ID And mobjBill.Details(lngRow - 1).Detail.ִ�п��� = .Detail.ִ�п��� _
                                                And InStr(",5,6,7,", mobjBill.Details(lngRow - 1).�շ����) = 0 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            ElseIf mrsUnit!ID = lng����ID And Bill.ListIndex = -1 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            End If
                                        End If
                                    End If
                                End If
                                mrsUnit.MoveNext
                            Next
                            
                            If Not blnEnter And .Detail.ִ�п��� = 4 Then    'ִ�п���Ϊָ�����ҵ�,ȱʡΪ����Ա���ڿ���
                                For i = 0 To Bill.ListCount - 1
                                    If Bill.ItemData(i) = UserInfo.����ID Then Bill.ListIndex = i: Exit For
                                Next
                            End If
                            
                            If Bill.ListIndex = -1 Then '���û����ȡ���е�ִ�п���
                                For i = 0 To Bill.ListCount - 1
                                    If Bill.ItemData(i) = .ִ�в���ID Then Bill.ListIndex = i: Exit For
                                Next
                            End If
                            
                            If mblnNurseStation Then    '��ʿվȱʡ����һ�����˵�ִ�п�����ȱʡ.
                                Dim tyPati As TY_PATIINFOR
                                For i = 0 To rptPati.Rows.Count - 1
                                    If rptPati.Rows(i).Record.Tag = "1" Then
                                        tyPati = GetPatiInforByReport(i)
                                        Exit For
                                    End If
                                Next i
                                For i = 0 To Bill.ListCount - 1
                                    If Bill.ItemData(i) = tyPati.��������ID Then Bill.ListIndex = i: Exit For
                                Next
                            End If
                        End If
                        
                        If Bill.ListIndex = -1 And Bill.ListCount > 0 Then Bill.ListIndex = 0
                    End If
                    
                    If Bill.ListIndex <> -1 Then
                        .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                        Bill.TextMatrix(lngRow, lngCol) = Bill.List(Bill.ListIndex)
                    Else
                        .ִ�в���ID = 0
                    End If
                End With
            End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ϊ�����޸�״̬
    '����:���˺�
    '����:2015-07-13 11:54:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
        
    cboNO.Locked = Not bln
    cbo��������.Locked = Not bln
    cbo������.Locked = Not bln
    chk�Ӱ�.Enabled = bln
    txtDate.Enabled = bln
    Bill.Active = bln
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 

Private Function CalcMoneys(tyPati As TY_PATIINFOR, _
    Optional lngRow As Long = 0, Optional ByVal blnNoPrompt As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������¼���ָ���л������еĽ��
    '���:tyPati-��ǰ������Ϣ
    '     lngRow=ָ����,Ϊ0��ʾ����������
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-07-13 12:00:02
    '˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strMainRows As String
    Dim bln��������ۿ� As Boolean
    
    On Error GoTo errHandle
    
    If mobjBill.Details.Count = 0 Then CalcMoneys = True: Exit Function
    If lngRow > mobjBill.Details.Count Then CalcMoneys = True: Exit Function
    
    Call reSetBillObject(tyPati, mobjBill)   '�������ö���
    
    For i = IIf(lngRow = 0, 1, lngRow) To IIf(lngRow = 0, mobjBill.Details.Count, lngRow)
        bln��������ۿ� = False
        If gbln��������ۿ� Then                    '����������ηѱ�,����ܼ����ۿ۲�����Ч,�����ܼ���
            If mobjBill.Details(i).�������� > 0 Then    '����
                bln��������ۿ� = Not mobjBill.Details(mobjBill.Details(i).��������).Detail.���ηѱ�
                If bln��������ۿ� And lngRow <> 0 Then strMainRows = "," & mobjBill.Details(i).��������      '��������һ�е�ʱ��
            Else
                If CheckItemHaveSub(i) Then                            '����������
                     bln��������ۿ� = Not mobjBill.Details(i).Detail.���ηѱ�
                     If bln��������ۿ� Then strMainRows = strMainRows & "," & i  'һҳ�����ж��������,�ȼ�¼�����к�,���������������ۿ�
                End If
            End If
        End If
        Call CalcMoney(tyPati, i, bln��������ۿ�, blnNoPrompt)
    Next
    
    '������������,������bln��������ۿ۱���,��Ϊ�������������Ǵ������ʱ�Ѹı�
    If gbln��������ۿ� Then
        For i = 1 To UBound(Split(strMainRows, ","))
            Call Calc��������ʵ��(Split(strMainRows, ",")(i))
        Next
    End If
    CalcMoneys = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CalcMoney(ByRef tyPati As TY_PATIINFOR, _
    lngRow As Long, Optional bln��������ۿ� As Boolean, Optional ByVal blnNoPrompt As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������¼���ָ���еĽ��
    '���:tyPati-��ǰ����Ĳ�����Ϣ
    '     lngRow=ָ����
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-07-13 13:58:14
    '˵����1.ExpenseBill���ϵ�������Ӧ���ݵ��к�
    '      2.���ֻ�ܶ�Ӧһ��������Ŀ:mobjBill.Details(lngRow).InComes(1)
    '      3.������ϸĿδ�����������Ŀ(��һ�μ���),��ʹ��Ĭ���ּ�
    '      4.������ϸĿ�Ѿ������������Ŀ(����2��),���ֶ�����(Ҳ����δ��)�˵���,�򰴸õ��ۼ��㡣
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strInfo As String, strSQL As String, i As Long
    Dim dblMoney As Double '�û�����ı�۽��

    Dim dblAllTime As Double, dbl�Ӱ�Ӽ��� As Double
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dblʣ������ As Double
    Dim strPriceGrade As String, strWherePriceGrade As String
    
    On Error GoTo errH
    If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
        strPriceGrade = mstrҩƷ�۸�ȼ�
    ElseIf mobjBill.Details(lngRow).�շ���� = "4" Then
        strPriceGrade = mstr���ļ۸�ȼ�
    Else
        strPriceGrade = mstr��ͨ�۸�ȼ�
    End If
    
    If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
        Call AdjustCpt(mobjBill.Details(lngRow).�շ�ϸĿID)
    End If
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "       And (b.�۸�ȼ� = [2]" & vbNewLine & _
            "            Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "                And Not Exists(Select 1" & vbNewLine & _
            "                               From �շѼ�Ŀ" & vbNewLine & _
            "                               Where b.�շ�ϸĿId = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
            "                                     And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null"
    End If
    strSQL = _
        " Select B.������ĿID,C.����,C.�վݷ�Ŀ,B.�ּ�,B.ԭ��,B.�Ӱ�Ӽ���,B.�����շ���,B.ȱʡ�۸� " & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
        " Where B.�շ�ϸĿID = A.ID And C.ID = B.������ĿID " & _
        " And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD')) " & _
        " And A.ID=[1]" & vbNewLine & _
        strWherePriceGrade
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details(lngRow).�շ�ϸĿID, strPriceGrade)
    If rsTmp.EOF Then
        '���û��������Ŀ,�������Ӧ�ĳ������
        Set mobjBill.Details(lngRow).InComes = New BillInComes
        CalcMoney = True
        Exit Function
    End If
    
    '�Ȼ�ȡ����Ա��ǰ����ı�۽��
    With mobjBill.Details(lngRow)
        If InStr(",5,6,7,", .�շ����) > 0 Or (.�շ���� = "4" And .Detail.��������) Then
            '����ҩƷʱ��(�����򲻷���)
            '��Ȼ�м�¼(�������Ŀʱ���ж�)
            dblAllTime = .���� * .����
            If gblnסԺ��λ And InStr(",5,6,7,", .�շ����) > 0 Then
                dblAllTime = dblAllTime * .Detail.סԺ��װ '���ʱ�۰��ۼ��������м���
            End If
            If dblAllTime <> 0 Or Not .Detail.��� Then
                Set rsPrice = zlDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                            Me.Caption, .�շ�ϸĿID, .ִ�в���ID, dblAllTime)
                If rsPrice.EOF Then
                    '��ȡ�۸�ʧ��
                    If InStr(",5,6,7,", .�շ����) > 0 Then
                        MsgBox "�� " & lngRow & " ��ҩƷ""" & .Detail.���� & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                    Else
                        MsgBox "�� " & lngRow & " ����������""" & .Detail.���� & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                    End If
                Else
                    strPrice = Nvl(rsPrice!Price) & "|||"
                    varPrice = Split(strPrice, "|")
                    dblMoney = Val(varPrice(0))
                    dblʣ������ = Val(varPrice(2))
                    
                    If dblʣ������ <> 0 And .Detail.��� Then
                        '����δ�ֽ����
                        If Not blnNoPrompt Then
                            If InStr(",5,6,7,", .�շ����) > 0 Then
                                MsgBox "�� " & lngRow & " ��ʱ��ҩƷ""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                            Else
                                MsgBox "�� " & lngRow & " ��ʱ����������""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                            End If
                        End If
                        dblMoney = 0
                    End If
                End If
            Else
                dblMoney = 0
            End If
        Else
            If .Detail.��� Then
                If .InComes.Count = 0 Then '��һ�μ�����ȡȱʡֵ
                    dblMoney = Val(Nvl(rsTmp!ȱʡ�۸�))
                Else                        '��ȡ����Ա��ǰ����ı�۽��
                    dblMoney = .InComes(1).��׼����
                    '����û�����ı�۲������۷�Χ����ȡȱʡֵ
                    If CheckScope(Val(Nvl(rsTmp!ԭ��)), Val(Nvl(rsTmp!�ּ�)), dblMoney) <> "" Then
                        dblMoney = Val(Nvl(rsTmp!ȱʡ�۸�))
                    End If
                End If
            End If
        End If
    End With
    
    '�����ԭ�м�¼
    Set mobjBill.Details(lngRow).InComes = New BillInComes
    
    '��д���з��ü�¼
    For i = 1 To rsTmp.RecordCount
        Set mobjBillIncome = New BillInCome
        With mobjBillIncome
            .������ĿID = rsTmp!������ĿID
            .������Ŀ = rsTmp!����
            .�վݷ�Ŀ = Nvl(rsTmp!�վݷ�Ŀ)
            .ԭ�� = Val(Nvl(rsTmp!ԭ��))
            .�ּ� = Val(Nvl(rsTmp!�ּ�))
            
            If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
                If gblnסԺ��λ Then
                    .��׼���� = Format(dblMoney * mobjBill.Details(lngRow).Detail.סԺ��װ, gstrFeePrecisionFmt)
                Else
                    .��׼���� = Format(dblMoney, gstrFeePrecisionFmt)
                End If
            Else
                If mobjBill.Details(lngRow).Detail.��� Then
                    .��׼���� = Format(dblMoney, gstrFeePrecisionFmt)
                Else
                    .��׼���� = Format(Nvl(rsTmp!�ּ�, 0), gstrFeePrecisionFmt)
                End If
            End If
            
            'Ӧ�ս��=���� * ���� * ����
            .Ӧ�ս�� = .��׼���� * IIf(mobjBill.Details(lngRow).���� = 0, 1, mobjBill.Details(lngRow).����) * mobjBill.Details(lngRow).����
            
            '�������������ü���(����������Ŀ)
            If mobjBill.Details(lngRow).���ӱ�־ = 1 And mobjBill.Details(lngRow).�շ���� = "F" Then
                .Ӧ�ս�� = .Ӧ�ս�� * IIf(IsNull(rsTmp!�����շ���), 1, rsTmp!�����շ��� / 100)
            End If
            
            '�Ӱ�����ʼ���
            dbl�Ӱ�Ӽ��� = 0
            If mobjBill.�Ӱ��־ = 1 And mobjBill.Details(lngRow).Detail.�Ӱ�Ӽ� Then
                dbl�Ӱ�Ӽ��� = IIf(IsNull(rsTmp!�Ӱ�Ӽ���), 0, rsTmp!�Ӱ�Ӽ��� / 100)
                .Ӧ�ս�� = .Ӧ�ս�� + .Ӧ�ս�� * dbl�Ӱ�Ӽ���
            End If
            
            .Ӧ�ս�� = CCur(Format(.Ӧ�ս��, gstrDec))
            dblAllTime = mobjBill.Details(lngRow).���� * mobjBill.Details(lngRow).����
            If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
                If gblnסԺ��λ Then dblAllTime = dblAllTime * mobjBill.Details(lngRow).Detail.סԺ��װ
            End If
            
            If mobjBill.Details(lngRow).Detail.���ηѱ� _
                Or bln��������ۿ� Or .Ӧ�ս�� = 0 Or tyPati.����ID = 0 Then
                .ʵ�ս�� = .Ӧ�ս��
            Else
                If .Ӧ�ս�� = 0 Then
                    .ʵ�ս�� = 0
                    mobjBill.Details(lngRow).�ѱ� = mobjBill.�ѱ�
                Else
                     'ҩƷ���ɱ��ۼ���,��������
                    .ʵ�ս�� = CCur(Format(ActualMoney(mobjBill.�ѱ�, .������ĿID, .Ӧ�ս��, _
                         mobjBill.Details(lngRow).�շ�ϸĿID, mobjBill.Details(lngRow).ִ�в���ID, dblAllTime, dbl�Ӱ�Ӽ���), gstrDec))
                End If
            End If
            
            '��ȡ��Ŀ������Ϣ,ҽ�����˲Ŵ���,����Ҫ����ҽ��
            If tyPati.����ID <> 0 And tyPati.���� <> 0 Then
                strInfo = gclsInsure.GetItemInsure(tyPati.����ID, _
                    mobjBill.Details(lngRow).�շ�ϸĿID, .ʵ�ս��, False, tyPati.����, _
                     mobjBill.Details(lngRow).ժҪ & "||" & dblAllTime)
                     
                If strInfo <> "" Then
                    mobjBill.Details(lngRow).������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                    mobjBill.Details(lngRow).���մ���ID = Val(Split(strInfo, ";")(1))
                    .ͳ���� = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                    mobjBill.Details(lngRow).���ձ��� = CStr(Split(strInfo, ";")(3))
                    If UBound(Split(strInfo, ";")) >= 4 Then
                        If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(lngRow).ժҪ = CStr(Split(strInfo, ";")(4))
                        If UBound(Split(strInfo, ";")) >= 5 Then
                            If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(lngRow).Detail.���� = Split(strInfo, ";")(5)
                        End If
                    End If
                End If
            End If
            'ʵ�ս�����Key��,�Դ���ֱ�����(��Key�д��ԭʼʵ�ս��,����)
            mobjBill.Details(lngRow).InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��, .ͳ����
        End With
        rsTmp.MoveNext
    Next
    CalcMoney = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub ShowDetails(Optional lngRow As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ����ʾָ���л������е�����
    '���:lngRow=ָ����,Ϊ0��ʾ��ʾ������
    '����:���˺�
    '����:2015-07-13 14:11:09
    '˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, curTotal As Currency
    Bill.Redraw = False
    If lngRow = 0 Then
        For i = 1 To mobjBill.Details.Count
            ShowDetail i
        Next
    Else
        ShowDetail lngRow
    End If
    Bill.Redraw = True
End Sub
Private Sub ShowDetail(lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ����ʾָ���е�����
    '���:lngRow=ָ����
    '����:���˺�
    '����:2015-07-13 14:12:47
    '˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl���� As Double, cur��� As Currency
    Dim i As Long, j As Long
    
    If lngRow > Bill.Rows - 1 Then Exit Sub
    If lngRow > mobjBill.Details.Count Then Exit Sub
    
    '���������
    For i = 1 To Bill.Cols - 1
        '����ʱ�շ�������
        If Not (i = 1 And Bill.TextMatrix(lngRow, i) <> "") Then Bill.TextMatrix(lngRow, i) = ""
    Next
    
    If mobjBill.Details(lngRow).�շ���� <> "" Then
        Bill.RowData(lngRow) = Asc(mobjBill.Details(lngRow).�շ����)
    End If
    
    'ˢ�µ�����
    For i = 1 To Bill.Cols - 1
        Select Case Bill.TextMatrix(0, i)
            Case "���"
                '������ݻ������Ŀֻ(��)��ʾ����
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.�������
            Case "��Ŀ"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
            Case "���"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.���
            Case "��Ʒ��"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.��Ʒ��
            Case "��λ"
                If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 And gblnסԺ��λ Then
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.סԺ��λ
                Else
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.���㵥λ
                End If
            Case "����"
                Bill.TextMatrix(lngRow, i) = IIf(mobjBill.Details(lngRow).���� = 0, 1, mobjBill.Details(lngRow).����)
            Case "����"
                '�����ڵ�һ����ʾʱ��Ĭ������Ϊ1
                Bill.TextMatrix(lngRow, i) = FormatEx(mobjBill.Details(lngRow).����, 5)
            Case "����"
                '�����Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                '��һ�μ���ʱ����Ĭ������Ϊ1�Ļ����ϼ��������
                dbl���� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        dbl���� = dbl���� + mobjBill.Details(lngRow).InComes(j).��׼����
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(dbl����, gstrFeePrecisionFmt)
            Case "Ӧ�ս��"
                'Ӧ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                cur��� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur��� = cur��� + mobjBill.Details(lngRow).InComes(j).Ӧ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur���, gstrDec)
            Case "ʵ�ս��"
                'ʵ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                cur��� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur��� = cur��� + mobjBill.Details(lngRow).InComes(j).ʵ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur���, gstrDec)
            Case "ִ�п���"
                If mobjBill.Details(lngRow).ִ�в���ID <> 0 Then
                    mrsUnit.Filter = "ID=" & mobjBill.Details(lngRow).ִ�в���ID
                    If mrsUnit.RecordCount <> 0 Then
                        Bill.TextMatrix(lngRow, i) = mrsUnit!���� & "-" & mrsUnit!����
                    Else
                        Bill.TextMatrix(lngRow, i) = GET��������(mobjBill.Details(lngRow).ִ�в���ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(lngRow, i) = ""
                End If
            Case "��־"
                If mobjBill.Details(lngRow).�շ���� = "F" And mobjBill.Details(lngRow).���ӱ�־ = 1 Then
                    Bill.TextMatrix(lngRow, i) = "��"
                End If
            Case "����"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Public Sub ShowMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ����ʾ������Ŀ������
    '����:���˺�
    '����:2015-07-13 14:14:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, k As Long
    Dim blnExist As Boolean, curTotal As Currency, curӦ��Total As Currency
    
    vsMoney.Redraw = False
    
    '�������ܷ�Ŀ
    Set mcolMoneys = New BillInComes
    For i = 1 To mobjBill.Details.Count
        For j = 1 To mobjBill.Details(i).InComes.Count
            '�����Ƿ��Ѿ��������������Ŀ,������ϼ�,��������
            blnExist = False
            For k = 1 To mcolMoneys.Count
                If mcolMoneys(k).������ĿID = mobjBill.Details(i).InComes(j).������ĿID Then
                    blnExist = True: Exit For
                End If
            Next
            If blnExist Then
                mcolMoneys(k).ʵ�ս�� = mcolMoneys(k).ʵ�ս�� + mobjBill.Details(i).InComes(j).ʵ�ս��
                mcolMoneys(k).Ӧ�ս�� = mcolMoneys(k).Ӧ�ս�� + mobjBill.Details(i).InComes(j).Ӧ�ս��
            Else
                With mobjBill.Details(i).InComes(j)
                    mcolMoneys.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��
                End With
            End If
        Next
    Next
    
    'ˢ����ʾ
    If mcolMoneys.Count > 0 Then
        vsMoney.Rows = mcolMoneys.Count + 1
    End If
    If vsMoney.Rows < 5 Then vsMoney.Rows = 5

    Call SetMoneyList
    
    'ˢ����ʾ
    If mcolMoneys.Count > 0 Then
        vsMoney.Rows = mcolMoneys.Count + 1
    End If
    If vsMoney.Rows < 5 Then vsMoney.Rows = 5
    
    
    For i = 1 To mcolMoneys.Count
        vsMoney.TextMatrix(i, 0) = mcolMoneys(i).������Ŀ
        vsMoney.TextMatrix(i, 1) = Format(mcolMoneys(i).ʵ�ս��, gstrDec)
        curTotal = curTotal + mcolMoneys(i).ʵ�ս��
        curӦ��Total = curӦ��Total + mcolMoneys(i).Ӧ�ս��
    Next
    txtӦ��.Text = Format(curӦ��Total, gstrDec)
    'txtʵ��.Text = Format(curTotal, gstrDec)
    For i = 1 To vsMoney.Rows - 1
        vsMoney.TopRow = i
    Next
    vsMoney.Redraw = True
End Sub

Private Function GetCurӦ��() As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���˵�ǰ���ݺϼƽ��(�շѲ����ۼӵ���ʱ��)
    '����:���˺�
    '����:2015-07-13 14:15:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 1 To mcolMoneys.Count
        GetCurӦ�� = GetCurӦ�� + mcolMoneys(i).Ӧ�ս��
    Next
End Function

Private Function GetInputDetail(ByVal lng��Ŀid As Long) As Detail
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շ���Ŀ��Ϣ
    '����:���˺�
    '����:2015-07-13 14:15:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objDetail As New Detail
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, lngMediCareNO As Long
        
    If mstrInsures <> "" Then
        strSQL = _
        " Select A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ," & _
        "       A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.����ժҪ,A.�������,M.Ҫ������," & _
        "       Decode(A.���,'4',D.����ID,C.ҩ��ID) as ҩ��ID," & _
        "       Decode(A.���,'4',D.���÷���,C.ҩ������) as ����," & _
        "       Decode(A.���,'4',1,C.סԺ��װ) as סԺ��װ," & _
        "       Decode(A.���,'4',A.���㵥λ,C.סԺ��λ) as סԺ��λ,D.��������,A.¼������,C.��ҩ��̬,M1.���� as ��������,M1.���㵥λ as ������λ,C.����ϵ��" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,�������� D,�շ���Ŀ���� E,�շ���Ŀ���� E1,������ĿĿ¼ M1," & _
        "       (   Select A1.�շ�ϸĿID,max(A1.Ҫ������) as Ҫ������  " & _
        "           From ����֧����Ŀ A1,Table(f_Num2List([2])) B1 " & _
        "           Where A1.�շ�ϸĿID=[1] and a1.����=b1.Column_value " & _
        "           Group by A1.�շ�ϸĿID) M" & _
        " Where A.���=B.���� And A.ID=C.ҩƷID(+) And C.ҩ��ID=M1.id(+) And A.ID=D.����ID(+)" & _
        "       And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And A.ID=M.�շ�ϸĿID(+)  " & vbNewLine & _
        "       And A.ID=[1]"
    Else
        strSQL = _
        " Select A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ," & _
        "       A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.����ժҪ,A.�������,0 as Ҫ������," & _
        "       Decode(A.���,'4',D.����ID,C.ҩ��ID) as ҩ��ID," & _
        "       Decode(A.���,'4',D.���÷���,C.ҩ������) as ����," & _
        "       Decode(A.���,'4',1,C.סԺ��װ) as סԺ��װ," & _
        "       Decode(A.���,'4',A.���㵥λ,C.סԺ��λ) as סԺ��λ,D.��������,A.¼������,C.��ҩ��̬,M1.���� as ��������,M1.���㵥λ as ������λ,C.����ϵ��" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,�������� D,�շ���Ŀ���� E,�շ���Ŀ���� E1,������ĿĿ¼ M1" & _
        " Where A.���=B.���� And A.ID=C.ҩƷID(+) And C.ҩ��ID=M1.id(+) And A.ID=D.����ID(+)" & _
        "       And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And A.ID=[1]"
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, mstrInsures)
    With objDetail
        .ID = rsTmp!ID
        .ҩ��ID = Nvl(rsTmp!ҩ��ID, 0) '�����ж������ظ�
        .��� = rsTmp!���
        .������� = rsTmp!�������
        .���� = rsTmp!����
        .���� = rsTmp!����
        .��� = Nvl(rsTmp!���)
        .���㵥λ = Nvl(rsTmp!���㵥λ)
        .סԺ��λ = Nvl(rsTmp!סԺ��λ)
        .סԺ��װ = Nvl(rsTmp!סԺ��װ, 1)
        .���� = Nvl(rsTmp!����, 0) = 1 '�Ƿ�ҩ������
        .��� = Nvl(rsTmp!�Ƿ���, 0) = 1 '��ҩƷ�����Ƿ�ʱ��
        .���� = Nvl(rsTmp!��������)
        .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
        .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
        .ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
        .������� = Nvl(rsTmp!�������, 0)
        .����ժҪ = Nvl(rsTmp!����ժҪ, 0) = 1
        .�������� = Nvl(rsTmp!��������, 0) = 1
        .Ҫ������ = Nvl(rsTmp!Ҫ������, 0) = 1
        .¼������ = Val("" & rsTmp!¼������)
        .��ҩ��̬ = Val(Nvl(rsTmp!��ҩ��̬))
        .��Ʒ�� = Nvl(rsTmp!��Ʒ��)
        .�������� = Nvl(rsTmp!��������)
        .������λ = Nvl(rsTmp!������λ)
        .����ϵ�� = Val(Nvl(rsTmp!����ϵ��))
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetDetail(Detail As Detail, lngRow As Long, lngDoUnit As Long, _
    Optional bytParent As Byte = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ�����շ�ϸĿ�����趨����ָ�㶨�е��շ�ϸĿ��(�����Ļ��޸�)
    '����:���˺�
    '����:2015-07-13 14:18:31
    '˵����
    '      1.���������������շ�ϸĿ�У�����
    '      2.��bytParent<>0ʱ,��Ϊ���ô�����Ŀ,������Ŀһ����������,������Ŀһ������
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim tmpIncomes As New BillInComes
    Dim intPay As Integer, i As Long, dblTime As Double
    
    'ȡ������ҩ�ĸ���
    intPay = GetPay(lngRow)
    If Detail.��� <> "7" Then intPay = 1
    
    If mobjBill.Details.Count < lngRow Then
        '������ж�Ӧ�ĳ��������δ��ʼ,�����
        With Detail
            '���=�к�,����=0
            '����=1,������Ŀ�Ĵ������������ȷ��
            'ִ�в���ID:����ϸĿִ�п��ұ�־ȡ
            '���ӱ�־:�Ե�һ��Ϊ��,����Ϊ������Ȩ
            '���뼯=��
            If bytParent <> 0 Then
                '���ø���RowData
                Bill.RowData(lngRow) = Asc(Detail.���)
                '��ʼ����
                If Detail.���д��� = 0 Then '�ǹ��д���
                    dblTime = Detail.��������
                ElseIf Detail.���д��� = 1 Then '�̶��Ĺ��д���
                    dblTime = IIf(Detail.�������� = 0, 1, Detail.��������)
                ElseIf Detail.���д��� = 2 Then '�������Ĺ��д���
                    dblTime = Detail.�������� * mobjBill.Details(bytParent).����
                End If
            Else
                
                If InStr(",5,6,7,", Detail.���) > 0 Then
                    dblTime = 0
                                     
                Else
                    dblTime = 1
                End If
            End If
            mobjBill.Details.Add Detail, .ID, CByte(lngRow), CInt(bytParent), 0, 0, 0, 0, "", "", "", _
            0, 0, mobjBill.�ѱ�, 0, .���, .���㵥λ, "", intPay, dblTime, 0, lngDoUnit, tmpIncomes
        End With
    Else '��������Ѿ�����,���޸�
        
        If InStr(",5,6,7,", Detail.���) > 0 Then
            dblTime = 0
        Else
            dblTime = 1
        End If
        With mobjBill.Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .�ѱ� = mobjBill.�ѱ�
            .���� = intPay
            .���ӱ�־ = 0
            .���㵥λ = Detail.���㵥λ
            .�շ���� = Detail.���
            .�շ�ϸĿID = Detail.ID
            .���� = dblTime
            .��� = lngRow
            .�������� = 0
            .ִ�в���ID = lngDoUnit
        End With
    End If
End Sub

Private Function ShouldDO(lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϸ����Ƿ�Ӧ��ȡ������Ŀ
    '����:�д�����Ŀ����true,���򷵻�False
    '����:���˺�
    '����:2015-07-13 14:19:48
    '˵�����������շ���Ŀ�д�����Ŀ����δȡ��ȡ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select count(����ID) as NUM From �շѴ�����Ŀ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details(lngRow).�շ�ϸĿID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!Num) Then
            ShouldDO = False
        ElseIf rsTmp!Num = 0 Then
            ShouldDO = False
        Else
            blnExist = False
            For i = lngRow + 1 To mobjBill.Details.Count
                If mobjBill.Details(i).�������� = lngRow Then
                    blnExist = True: Exit For
                End If
            Next
            If Not blnExist Then
                ShouldDO = True
            Else
                ShouldDO = False
            End If
        End If
    Else
        ShouldDO = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetSubDetails(ByVal lng��Ŀid As Long) As Details
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ���շ�ϸĿ�Ĵ�����Ŀ��
    '���:lng��Ŀid-�շ�ϸĿID
    '����:����Details����
    '����:���˺�
    '����:2015-07-13 14:20:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objDetail As New Detail, lngMediCareNO As Long
    Dim dblStock As Double
    
    Set GetSubDetails = New Details
    If mstrInsures <> "" Then
        strSQL = _
        " Select A.ID,Decode(A.���,'4',E.����ID,D.ҩ��ID) as ҩ��ID,A.���,B.���� as �������," & _
        "       A.��������,A.����,Nvl(F.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ,A.���ηѱ�,G.Ҫ������," & _
        "       Decode(A.���,'4',E.���÷���,D.ҩ������) as ����,A.�Ƿ���," & _
        "       Decode(A.���,'4',1,D.סԺ��װ) as סԺ��װ,A.�������," & _
        "       Decode(A.���,'4',A.���㵥λ,D.סԺ��λ) as סԺ��λ," & _
        "       A.�Ӱ�Ӽ�,A.ִ�п���,C.���д���,C.��������,E.��������,D.��ҩ��̬,M1.���� as ��������,M1.���㵥λ as ������λ,D.����ϵ��" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�շѴ�����Ŀ C,ҩƷ��� D,�������� E,�շ���Ŀ���� F,�շ���Ŀ���� E1, ������ĿĿ¼ M1," & _
        "       (   Select A1.�շ�ϸĿID,max(A1.Ҫ������) as Ҫ������  " & _
        "           From ����֧����Ŀ A1,Table(f_Num2List([2])) B1 " & _
        "           Where A1.�շ�ϸĿID=[1] and a1.����=b1.Column_value " & _
        "           Group by A1.�շ�ϸĿID) G" & _
        " Where B.����=A.��� And C.����ID=A.ID And A.ID=D.ҩƷID(+) And D.ҩ��ID=M1.id(+) And A.ID=E.����ID(+)" & _
        "       And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        "       And A.ID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And C.����ID=[1] And A.ID=G.�շ�ϸĿID(+)   " & _
        " Order by ����"
    Else
        strSQL = _
        "Select A.ID,Decode(A.���,'4',E.����ID,D.ҩ��ID) as ҩ��ID,A.���,B.���� as �������," & _
        "       A.��������,A.����,Nvl(F.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ,A.���ηѱ�,0 as Ҫ������," & _
        "       Decode(A.���,'4',E.���÷���,D.ҩ������) as ����,A.�Ƿ���," & _
        "       Decode(A.���,'4',1,D.סԺ��װ) as סԺ��װ,A.�������," & _
        "       Decode(A.���,'4',A.���㵥λ,D.סԺ��λ) as סԺ��λ," & _
        "       A.�Ӱ�Ӽ�,A.ִ�п���,C.���д���,C.��������,E.��������,D.��ҩ��̬,M1.���� as ��������,M1.���㵥λ as ������λ,D.����ϵ��" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�շѴ�����Ŀ C,ҩƷ��� D,�������� E,�շ���Ŀ���� F,�շ���Ŀ���� E1,������ĿĿ¼ M1" & _
        " Where B.����=A.��� And C.����ID=A.ID And A.ID=D.ҩƷID(+) And D.ҩ��ID=M1.id(+)  And A.ID=E.����ID(+)" & _
        "   And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        "   And A.ID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "   And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "   And C.����ID=[1] " & _
        " Order by ����"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, lngMediCareNO)
    For i = 1 To rsTmp.RecordCount
        If mblnNurseStation And InStr(",5,6,7,", rsTmp!���) > 0 Then
            rsTmp.MoveNext
        Else
            Set objDetail = New Detail
            With objDetail
                .ID = rsTmp!ID
                .ҩ��ID = Nvl(rsTmp!ҩ��ID, 0)
                .���� = rsTmp!����
                .��� = Nvl(rsTmp!�Ƿ���, 0) = 1
                .��� = Nvl(rsTmp!���)
                .סԺ��װ = Nvl(rsTmp!סԺ��װ, 1)
                .סԺ��λ = Nvl(rsTmp!סԺ��λ)
                .���㵥λ = Nvl(rsTmp!���㵥λ)
                .���� = Nvl(rsTmp!����, 0) = 1
                .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
                .��� = rsTmp!���
                .������� = rsTmp!�������
                .���� = rsTmp!����
                .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
                .ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
                .������� = Nvl(rsTmp!�������, 0)
                .���д��� = Nvl(rsTmp!���д���, 0)
                .�������� = Nvl(rsTmp!��������, 1)
                .���� = Nvl(rsTmp!��������)
                .�������� = Nvl(rsTmp!��������, 0) = 1
                .Ҫ������ = Nvl(rsTmp!Ҫ������, 0) = 1
                .��ҩ��̬ = Val(Nvl(rsTmp!��ҩ��̬))
                .��Ʒ�� = Nvl(rsTmp!��Ʒ��)
                .�������� = Nvl(rsTmp!��������)
                .������λ = Nvl(rsTmp!������λ)
                .����ϵ�� = Val(Nvl(rsTmp!����ϵ��))
                GetSubDetails.Add .ID, .ҩ��ID, .���, .�������, .����, .����, .����, .����, .���, .���㵥λ, .˵��, .���ηѱ�, _
                    .סԺ��װ, .סԺ��λ, .����, .���, .�Ӱ�Ӽ�, .ִ�п���, .�������, .����, .����ժҪ, .���д���, .��������, .��������, , , , , , .Ҫ������, , .��ҩ��̬, .��Ʒ��, .��������, .������λ, .����ϵ��
            End With
            rsTmp.MoveNext
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub DeleteDetail(lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��ָ���շ���Ŀ��
    '����:���˺�
    '����:2015-07-13 14:22:03
    '˵������ʱ����������е�ɾ��,��Ҫ�����������д�����ϵ����Ӧ�ĵ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�������� <> 0 And mobjBill.Details(i).�������� > lngRow Then
            mobjBill.Details(i).�������� = mobjBill.Details(i).�������� - 1
        End If
        mobjBill.Details(i).��� = mobjBill.Details(i).��� - 1 '������кŶ�Ӧ
    Next
    mobjBill.Details.Remove lngRow
    If lngRow = 1 And mobjBill.Details.Count = 0 And Bill.Rows = 2 Then
        For i = 1 To Bill.Cols - 1
            Bill.TextMatrix(lngRow, i) = ""
            Bill.RowData(lngRow) = 0
        Next
    Else
        Bill.RemoveMSFItem lngRow
    End If
End Sub

Private Sub NewBill()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��һ���µĵ���(�������)
    '����:���˺�
    '����:2015-07-13 14:22:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnKeepDate As Boolean
    Dim dtdtCurDate As Date     '��������ǰʱ��
    
    mlngPreRow = 0
        
    sta.Panels(3).Text = ""
    Set mrsMedAudit = Nothing
    mstrWarn = ""
    cboNO.Text = ""
    Set mobjBill = New ExpenseBill
    Bill.ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)
    dtdtCurDate = zlDatabase.Currentdate
    chk�Ӱ�.Value = IIf(OverTime(dtdtCurDate), 1, 0)
    txtDate.Text = Format(dtdtCurDate, "yyyy-MM-dd HH:mm:ss")
    
    Call cbo��������_Click
    
    cmdOK.Visible = True
    
    
    With mobjBill
        .�����־ = 2
        .������ = UserInfo.����
        .������ = zlStr.NeedName(cbo������.Text)
        .����Ա��� = UserInfo.���
        .����Ա���� = UserInfo.����
        .����ʱ�� = CDate(txtDate.Text)
        .�Ӱ��־ = chk�Ӱ�.Value
        .Ӥ���� = 0
        
        If cbo��������.ListIndex = -1 Then
            .��������ID = 0
        Else
            .��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
        If cboDrawDept.ListIndex = -1 Then
            .��ҩ����ID = 0
        Else
            .��ҩ����ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
        End If
    End With
End Sub
Private Sub ClearMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ʾ��
    '����:���˺�
    '����:2015-07-13 14:25:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    On Error GoTo errHandle
    
    vsMoney.Redraw = flexRDNone
    For i = 1 To vsMoney.Rows - 1
        For j = 0 To vsMoney.Cols - 1
            vsMoney.TextMatrix(i, j) = ""
        Next
    Next
    vsMoney.Rows = 5
    vsMoney.Redraw = flexRDBuffered

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub InitInsurePara(ByVal lng����ID As Long, ByVal intInsure As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҽ������
    '����:���˺�
    '����:2015-07-13 15:36:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    MCPAR.�����ϴ� = gclsInsure.GetCapability(support�����ϴ�, , intInsure)
    MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub chkIn_Click()
    sta.Panels(2) = ""
    If chkIn.Value = Checked Then
        txtIn.Enabled = True
        txtIn.BackColor = &H80000005
        sta.Panels(2) = "������Ҫ����ļ��ʵ����ݺ���"
        txtIn.SetFocus
    Else
        txtIn.Text = ""
        txtIn.Enabled = False
        txtIn.BackColor = &HE0E0E0
        Bill.SetFocus
    End If
End Sub

Private Sub txtIn_KeyPress(KeyAscii As Integer)
    Dim tmpBill As New ExpenseBill
    Dim tyPati As TY_PATIINFOR
    Dim lng����ID As Long, i As Long
    Dim lngPre As Long, strPre As String
    Dim dtCurdate As Date     '��������ǰʱ��
 
    
    On Error GoTo errH
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtIn, KeyAscii)
        Exit Sub
    End If
    
    txtIn.Text = GetFullNO(txtIn.Text, 14)
   
    Set tmpBill = ImportBill(txtIn.Text, False, Me, False, gblnסԺ��λ, , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    If tmpBill.NO = "" Then
        MsgBox "��ȡ����ʧ�ܡ�", vbExclamation, gstrSysName
        txtIn.Text = "": txtIn.SetFocus: Exit Sub
    End If

    '�����޸ļ���ʾ
    Screen.MousePointer = 11
                
    lng����ID = tmpBill.����ID
    lngPre = tmpBill.��������ID
    strPre = tmpBill.������
    If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then strPre = ""
    
    '�������Ĳ�����Ϣ
    tmpBill.����ID = 0
    tmpBill.��ҳID = 0
    tmpBill.���� = ""
    tmpBill.��ʶ�� = 0
    tmpBill.���� = ""
    tmpBill.�Ա� = ""
    tmpBill.���� = ""
    tmpBill.�ѱ� = ""
    tmpBill.����ID = 0
    tmpBill.����ID = 0
    
    '���˺�:25882
    For i = 1 To tmpBill.Details.Count
        tmpBill.Details(i).����ID = 0
        tmpBill.Details(i).��ҳID = 0
        tmpBill.Details(i).���� = ""
        tmpBill.Details(i).�Ա� = ""
        tmpBill.Details(i).���� = ""
        tmpBill.Details(i).�ѱ� = ""
        tmpBill.Details(i).����ID = 0
        tmpBill.Details(i).����ID = 0
    Next
    
    '�������в�����Ϣ
    If Not mobjBill Is Nothing Then
        If mobjBill.����ID > 0 Then
            lng����ID = mobjBill.����ID
            lngPre = mobjBill.��������ID
            strPre = mobjBill.������
        End If
    End If
    
    Set mobjBill = New ExpenseBill
    Set mobjBill = tmpBill
    
    dtCurdate = zlDatabase.Currentdate
    mobjBill.NO = cboNO.Text
    mobjBill.�Ǽ�ʱ�� = dtCurdate
    mobjBill.����Ա��� = UserInfo.���
    mobjBill.����Ա���� = UserInfo.����
    mobjBill.�Ӱ��־ = chk�Ӱ�.Value
    mobjBill.Ӥ���� = 0
    
    'ȡ��ǰʱ��
    txtDate.Text = Format(dtCurdate, "yyyy-MM-dd HH:mm:ss")
 
    Bill.Redraw = False
    Bill.ClearBill
    Bill.Rows = mobjBill.Details.Count + 1
    
    Call InitBillColumnColor
    
    '���ʷ��౨��
    mstrWarn = ""
    
    mobjBill.��������ID = lngPre
    mobjBill.������ = strPre
    Call Set�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mobjBill.������, mobjBill.��������ID)
    '������Ķ����˺�ȷ���ѱ��,�ټ���۸�
    Call CalcMoneys(tyPati)
    Call ShowDetails
    Call ShowMoney
    
    Bill.Redraw = True
    chkIn.Value = 0
    Call SetDrawDrugDeptEnabled
    Call SetColNum
    
    Screen.MousePointer = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ReCalcInsure(tyPati As TY_PATIINFOR)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸ĵ���ʱ,���¼���ͳ������������Ϣ
    '���:tyPati-������Ϣ
    '����:���˺�
    '����:2015-07-13 15:44:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, dblAllTime As Double
    Dim strInfo As String, varTemp As Variant
    If tyPati.����ID = 0 Or tyPati.���� = 0 Then Exit Sub
    Err = 0: On Error GoTo ErrHand:
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            For j = 1 To .InComes.Count
                dblAllTime = .���� * .����
                If InStr(",5,6,7,", .�շ����) > 0 Then
                    If gblnסԺ��λ Then dblAllTime = dblAllTime * .Detail.סԺ��װ
                End If
                
                strInfo = gclsInsure.GetItemInsure(tyPati.����ID, .�շ�ϸĿID, .InComes(j).ʵ�ս��, False, tyPati.����, _
                     .ժҪ & "||" & dblAllTime)
                If strInfo <> "" Then
                    varTemp = Split(strInfo & ";;;;", ";")
                    
                    .������Ŀ�� = Val(varTemp(0)) <> 0
                    .���մ���ID = Val(varTemp(1))
                    .InComes(j).ͳ���� = Val(varTemp(2))
                    .���ձ��� = CStr(varTemp(3))
                    
                    If UBound(varTemp) >= 4 Then
                        If CStr(varTemp(4)) <> "" Then .ժҪ = CStr(varTemp(4))
                        If UBound(varTemp) >= 5 Then
                            If varTemp(5) <> "" Then .Detail.���� = varTemp(5)
                        End If
                    End If
                End If
            Next
        End With
    Next
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub

Private Function GetAdviceIDs(ByVal lngҽ��ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ��ҽ��������ҽ����¼ID��
    '���:lngҽ��ID=һ��ҽ����¼����ID:Nvl(���ID,ID)
    '����: ����һ��ҽ��ID,�ö��ŷָ�
    '����:���˺�
    '����:2015-07-13 15:52:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select ID From ����ҽ����¼ Where ID=[1] Or ���ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    
    GetAdviceIDs = Mid(strSQL, 2)
 
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ClearRows()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ڲ���־
    '����:���˺�
    '����:2015-07-13 15:59:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub

Private Function GetPay(lngRow As Long) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ������ҩ�ĸ���
    '����:���˺�
    '����:2015-07-13 16:00:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    GetPay = 1
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�շ���� = "7" And i <> lngRow Then
            GetPay = mobjBill.Details(i).����
            Exit For
        End If
    Next
End Function

Private Sub InitBillColumnColor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������ɫ
    '����:���˺�
    '����:2015-07-13 15:59:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Bill.SetColColor BillCol.���, &HE7CFBA
    Bill.SetColColor BillCol.��Ŀ, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE7CFBA
    Bill.SetColColor BillCol.ִ�п���, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE0E0E0
    Bill.SetColColor BillCol.����, &HE0E0E0
    Bill.SetColColor BillCol.��־, &HE0E0E0
End Sub

Private Function GetDetailNum(tyPati As TY_PATIINFOR, _
    ByVal lngRow As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ָ��ϸĿ���ܼ�������(����������)
    '���:lngRow=��ǰ������
    '����:����������
    '����:���˺�
    '����:2015-07-13 16:00:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim lngNum As Long, i As Long, lng�շ�ϸĿID As Long
    Dim strSQL As String
    If tyPati.����ID = 0 Then Exit Function
    If lngRow > mobjBill.Details.Count Then Exit Function
    
    lng�շ�ϸĿID = mobjBill.Details(lngRow).�շ�ϸĿID
    
    '��ǰ�����е�����
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If i <> lngRow And .�շ�ϸĿID = lng�շ�ϸĿID Then
                lngNum = lngNum + .���� * IIf(.���� = 0, 1, .����)
            End If
        End With
    Next
    
    '���ݿ��е�����
    strSQL = _
    " Select Sum(A.����*Nvl(A.����,1)" & IIf(gblnסԺ��λ, "/Nvl(B.סԺ��װ,1)", "") & ") as Num" & _
    " From סԺ���ü�¼ A,ҩƷ��� B" & _
    " Where A.�۸񸸺� is Null And A.���ʷ���=1" & _
            IIf(gbytBilling = 0, " And A.��¼״̬<>0", "") & _
    " And A.����ID=[1] And Nvl(A.��ҳID,0)=[2] And A.�շ�ϸĿID=B.ҩƷID(+) And A.�շ�ϸĿID+0=[3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, tyPati.����ID, tyPati.��ҳID, lng�շ�ϸĿID)
    If Not rsTmp.EOF Then
        lngNum = lngNum + Nvl(rsTmp!Num, 0)
    End If
    GetDetailNum = lngNum
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetWorkUnit(ByVal lngҩƷID As Long, ByVal str��� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ���пɹ�ѡ���ҩ��
    '���:lngҩƷID-ҩƷID
    '     str���-���
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2015-07-13 16:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim strSQL As String, strҩ�� As String, bytDay As Byte
    Dim int������� As Integer, str������� As String
    Dim int������Դ As Integer, lng��������ID As Long
    
    '������Ŀ��Ȩ��ȷ��ҩ���ķ������
    int������� = Get�������(lngҩƷID)
    
    If int������� = 1 Then
        str������� = "1,3"
    ElseIf int������� = 2 Then
        str������� = "2,3"
    ElseIf int������� = 3 Then
        If InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
            str������� = "1,2,3"
        Else
            str������� = "2,3"
        End If
    Else
            str������� = "2,3"
    End If
    
    'ȷ��������Դ
    int������Դ = 2
    
    lng��������ID = mobjBill.����ID
    If cbo��������.Visible Then
        If lng��������ID = 0 And cbo��������.ListIndex <> -1 Then lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    Else
        If lng��������ID = 0 Then lng��������ID = mlngDeptID
        If lng��������ID = 0 Then lng��������ID = GetNurseStationFirstPatiDeptID '��ʿ����վ,ȡ��һ�����˿���ID
        If lng��������ID = 0 Then lng��������ID = mlng����ID
    End If
       
    If str��� = "4" Then
        strSQL = _
        "Select Distinct c.Id, c.����, c.����, c.����, b.��������, b.�������" & vbNewLine & _
        "From �շ�ִ�п��� A, ��������˵�� B, ���ű� C" & vbNewLine & _
        "Where a.ִ�п���id + 0 = b.����id And b.�������� = '���ϲ���' And b.������� IN(" & str������� & ") And b.����id = c.Id And" & vbNewLine & _
        "      (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And (c.վ�� = '" & gstrNodeNo & "' Or c.վ�� Is Null) And" & vbNewLine & _
        "      (a.������Դ Is Null Or a.������Դ = [1]) And" & vbNewLine & _
        "      (a.��������id Is Null Or a.��������id = [2] Or Exists (Select 1 From �������Ҷ�Ӧ Where ����id = [2] And a.��������id = ����id)) And a.�շ�ϸĿid = [3]" & vbNewLine & _
        "Order By b.�������, c.����"
    Else
        '��ҩƷ����ȷ��ҩ������
        Select Case str���
            Case "5"
                strҩ�� = "��ҩ��"
            Case "6"
                strҩ�� = "��ҩ��"
            Case "7"
                strҩ�� = "��ҩ��"
        End Select
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If Not gblnҩ���ϰల�� Then
            strSQL = _
            " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[4]" & _
            "       And B.������� IN(" & str������� & ") And B.����ID=C.ID" & _
            "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
            "       And (A.������Դ is NULL Or A.������Դ=[1])" & _
            "       And (A.��������ID is NULL Or A.��������ID=[2])" & _
            "       And A.�շ�ϸĿID=[3]" & _
            " Order by B.�������,C.����"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
            " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[4]" & _
            "       And B.������� IN(" & str������� & ") And B.����ID=C.ID" & _
            "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
            "       And D.����ID=C.ID And D.����=[5]" & _
            "       And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
            "       And (A.������Դ is NULL Or A.������Դ=[1])" & _
            "       And (A.��������ID is NULL Or A.��������ID=[2])" & _
            "       And A.�շ�ϸĿID=[3]" & _
            " Order by B.�������,C.����"
        End If
    End If
    On Error GoTo errH
    Set mrsWork = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, int������Դ, lng��������ID, lngҩƷID, strҩ��, bytDay)
    GetWorkUnit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CalcGridToTal(Optional blnӦ�� As Boolean) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ӧ�ջ���ʵ�ܽ��
    '���:blnӦ��-�Ƿ�ȡӦ�ս��True-Ӧ��,False-ʵ��
    '����:�����ܽ��
    '����:���˺�
    '����:2015-07-13 16:08:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim i As Long, intCol As Integer

    On Error GoTo errHandle
    
    If mobjBill.Details.Count > 0 Then
        For Each objTmpDetail In mobjBill.Details
            For Each objTmpIncome In objTmpDetail.InComes
                If blnӦ�� Then
                    CalcGridToTal = CalcGridToTal + objTmpIncome.Ӧ�ս��
                Else
                    CalcGridToTal = CalcGridToTal + objTmpIncome.ʵ�ս��
                End If
            Next
        Next
        Exit Function
    End If

    For i = 1 To Bill.Cols - 1
        If blnӦ�� Then
            If Bill.TextMatrix(0, i) = "Ӧ�ս��" Then intCol = i: Exit For
        Else
            If Bill.TextMatrix(0, i) = "ʵ�ս��" Then intCol = i: Exit For
        End If
    Next

    For i = 1 To Bill.Rows - 1
        CalcGridToTal = CalcGridToTal + Val(Bill.TextMatrix(i, intCol))
    Next
    
    

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Sub SetColNum(Optional intRow As Long = 1)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ʾ���е��к�
    '���:intRow=�Ӹ��п�ʼ
    '����:���˺�
    '����:2015-07-13 16:11:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln As Boolean, i As Long
    
    Bill.Redraw = False
    For i = intRow To Bill.Rows - 1
        Bill.TextMatrix(i, BillCol.��) = i
    Next
    Bill.Redraw = True
End Sub

Private Function CheckDuty(Optional tmpDetail As Detail, _
    Optional blnCommon As Boolean = True, Optional strName As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ��ҩƷ�е�ְ���Ƿ��뵱ǰҽ����ְ����ƥ��
    '���:tmpDetail=�������Ŀ,����Ϊ������
    '     blnCommon=�Ƿ��������ж�,����Ϊҽ���򹫷Ѳ��˵��ж�
    '���أ���ƥ�����,0Ϊ��ȷ
    '����:���˺�
    '����:2015-07-13 16:12:25
    '˵����ְ��1=����,2=����,3=�м�,4=����/ʦ��,5=Ա/ʿ,9=��Ƹ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, intְ��A As Integer, intְ��B As Integer
    Dim strTmp As String, strAllDuty As String
    
    If cbo������.ListIndex = -1 Then Exit Function
    strAllDuty = "����,����,�м�,����/ʦ��,Ա/ʿ,,,,��Ƹ"
    Call GetOperatorInfo(mrs������, mobjBill.������, , intְ��A)
        
    If tmpDetail Is Nothing Then
        For i = 1 To mobjBill.Details.Count
            If InStr(",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
                If Not blnCommon Then
                    intְ��B = Val(Right(mobjBill.Details(i).Detail.����ְ��, 1))
                    If intְ��B > 0 Then
                        If intְ��A = 0 Then
                            If strName = "" Then
                                strTmp = "��ҽ���򹫷Ѳ���,�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                            Else
                                strTmp = "����:" & strName & "��ҽ���򹫷Ѳ���,���� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                            End If
                            CheckDuty = 1
                        ElseIf intְ��B < intְ��A Then
                            If strName = "" Then
                                strTmp = "��ҽ���򹫷Ѳ���,�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                            Else
                                strTmp = "����:" & strName & "��ҽ���򹫷Ѳ���,���� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                            End If
                            CheckDuty = i: Exit For
                        End If
                    End If
                Else
                    intְ��B = Val(Left(mobjBill.Details(i).Detail.����ְ��, 1))
                    If intְ��B > 0 Then
                        If intְ��A = 0 Then
                            strTmp = "�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                            CheckDuty = 1
                        ElseIf intְ��B < intְ��A Then
                            strTmp = "�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                            CheckDuty = i: Exit For
                        End If
                    End If
                End If
            End If
        Next
    Else
        If InStr(",5,6,7,", tmpDetail.���) = 0 Then Exit Function
        If Not blnCommon Then
            intְ��B = Val(Right(tmpDetail.����ְ��, 1))
            If intְ��B > 0 Then
                If intְ��A = 0 Then
                    If strName = "" Then
                        strTmp = "��ҽ���򹫷Ѳ���,ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                    Else
                        strTmp = "����:" & strName & "��ҽ���򹫷Ѳ���,��ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                    End If
                    CheckDuty = 1
                ElseIf intְ��B < intְ��A Then
                    If strName = "" Then
                        strTmp = "��ҽ���򹫷Ѳ���,ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                    Else
                        strTmp = "����:" & strName & "��ҽ���򹫷Ѳ���,��ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                    End If
                    CheckDuty = 1
                End If
            End If
        Else
            intְ��B = Val(Left(tmpDetail.����ְ��, 1))
            If intְ��B > 0 Then
                If intְ��A = 0 Then
                    strTmp = "ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                    CheckDuty = 1
                ElseIf intְ��B < intְ��A Then
                    strTmp = "ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                    CheckDuty = 1
                End If
            End If
        End If
    End If
    If CheckDuty > 0 Then MsgBox strTmp, vbInformation, gstrSysName
End Function

Private Function PhysicExist(objDetail As Detail, intRow As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ��ҩƷ�ڵ������Ƿ��Ѿ�����
    '���:objDetail=��Ŀ
    '     intRow=Ҫ�жϵ���
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2015-07-13 16:13:59
    '˵����ʱ�ۻ����ҩƷ��ͬһҩ����ֹ�ظ�����(�������ʾ,����ʱ��ֹ)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
    For i = 1 To mobjBill.Details.Count
        If i <> intRow And InStr(",4,5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
            If mobjBill.Details(i).Detail.ID = objDetail.ID Then
                If (mobjBill.Details(i).Detail.���� Or mobjBill.Details(i).Detail.���) _
                    And (objDetail.���� Or objDetail.���) Then
                    If objDetail.��� = "4" Then
                        If MsgBox("��������""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������" & _
                            vbCrLf & vbCrLf & "ע�⣺����������Ϊ������ʱ�۲���,�ظ�����ʱ���뱣֤���ǵķ��ϲ��Ų�ͬ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("ҩƷ""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������" & _
                            vbCrLf & vbCrLf & "ע�⣺��ҩƷΪ������ʱ��ҩƷ,�ظ�����ʱ���뱣֤���ǵ�ִ��ҩ����ͬ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                Else
                    If objDetail.��� = "4" Then
                        If MsgBox("��������""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("ҩƷ""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
End Function
Private Function Checkִ�п���() As Integer
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .ִ�в���ID = 0 Or Bill.TextMatrix(i, BillCol.ִ�п���) = "" Then
                If InStr(",5,6,7,", .�շ����) = 0 Then
                    Checkִ�п��� = i: Exit Function
                End If
            End If
        End With
    Next
End Function
Private Function Get��������ID() As Long
    If cbo��������.ListIndex <> -1 Then
        Get��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    Else
        Get��������ID = UserInfo.����ID
    End If
End Function

Public Function zl��ȡ��ҩ��̬(Optional ByVal lngRow As Long = -1, Optional blnOnly�г�ҩ As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����Ƿ�¼�����в�ҩ��
    '���:blnOnly�г�ҩ-���ж��Ƿ����г�ҩ(���䷽ʱ�ж���Ч):ԭ�����г�ҩ���䷽���Ѿ�����,�Ͳ���Ҫ���
    '     lngRow-��ǰ��������
    '����:
    '����:¼�����в�ҩ��,�򷵻���ҩ��̬����(0-ɢװ,1-��Ƭ,2-����),���򷵻�-1 ��ʾ��û��¼����ҩ��̬��Ŀ
    '����:���˺�
    '����:2010-02-02 11:44:17
    '����:27816
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    
    zl��ȡ��ҩ��̬ = -1
    '���δָ��ҳ,���õ�ǰҳ
    If mobjBill Is Nothing Then Exit Function
    strTemp = IIf(blnOnly�г�ҩ, ",6,", ",6,7,")
    With mobjBill.Details
        For i = 1 To .Count
            If InStr(1, strTemp, "," & .Item(i).�շ���� & ",") > 0 And .Item(i).�շ�ϸĿID <> 0 And i <> lngRow Then
                zl��ȡ��ҩ��̬ = .Item(i).Detail.��ҩ��̬
                Exit Function
            End If
        Next
    End With
End Function
Private Function zlCheckBill���ڷ�ɢװ��ҩ() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���д��ڷ�ɢװ��ҩ��̬
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-26 10:19:46
    '����:38328
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If mobjBill Is Nothing Then Exit Function
    If mobjBill.Details.Count = 0 Then Exit Function
    With mobjBill
        For i = 1 To mobjBill.Details.Count
            If .Details(i).�շ���� = "7" Then
                If .Details(i).Detail.��ҩ��̬ <> 0 Then    '0-ɢװ;1-��ҩ��Ƭ;2-����
                    zlCheckBill���ڷ�ɢװ��ҩ = True: Exit Function
                End If
            End If
        Next
    End With
End Function
Private Function GetҪ������(ByVal str�շ�ϸĿID As String, _
    ByRef rsItem As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҪ��������ҽ����Ŀ
    '����:������Ҫ�������շ�ϸĿ��¼��(����,�շ�ϸĿID,Ҫ������)
    '����:Ҫ����������true,���򷵻�False
    '����:���˺�
    '����:2015-07-14 17:22:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If str�շ�ϸĿID > 4000 Then
        strSQL = "" & _
        "   Select A1.�շ�ϸĿID,A1.����,A1.Ҫ������  " & _
        "   From ����֧����Ŀ A1,Table(f_Num2List([1])) B1  " & _
        "   Where  a1.����=b1.Column_value And A1.�շ�ϸĿID in (" & str�շ�ϸĿID & ")" & _
        "           And nvl(A1.Ҫ������,0)=1 "
    Else
        strSQL = "" & _
        "   Select A1.�շ�ϸĿID,A1.����,A1.Ҫ������  " & _
        "   From ����֧����Ŀ A1,Table(f_Num2List([1])) B1,Table(f_Num2List([2])) B2 " & _
        "   Where  a1.����=b1.Column_value And A1.�շ�ϸĿID=B2.Column_value " & _
        "           And nvl(A1.Ҫ������,0)=1 "
    End If
    Set rsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrInsures, str�շ�ϸĿID)
    GetҪ������ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetNurseStationFirstPatiDeptID() As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ի�ʿ����������ȡ��һ�����˵Ĳ��˲���ID
    '����:��Ի�ʿ������,���ص�һ������ID,���򷵻�0
    '����:���˺�
    '����:2017-11-14 17:22:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim lng����ID As Long, i As Long
    Dim lngFirstDeptID As Long, lngFirstSelDeptID As Long
    Dim lngFirtPatiID As Long
    
    On Error GoTo errHandle
     
    If mblnNurseStation = False Then GetNurseStationFirstPatiDeptID = 0: Exit Function
    lngFirtPatiID = 0
    For i = 0 To rptPati.Rows.Count - 1
        If lngFirtPatiID = 0 Then
            lngFirtPatiID = Val(rptPati.Rows(i).Record(COL_����ID).Value)
            lngFirstDeptID = Val(rptPati.Rows(i).Record(COL_��������ID).Value)
        End If
        If rptPati.Rows(i).Record.Tag = "1" Then
            GetNurseStationFirstPatiDeptID = Val(rptPati.Rows(i).Record(COL_��������ID).Value)
            Exit Function
        End If
    Next
    GetNurseStationFirstPatiDeptID = lngFirstDeptID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
