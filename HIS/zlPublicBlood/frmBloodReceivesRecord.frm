VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmBloodReceivesRecord 
   Caption         =   "ѪҺ���յǼ�"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16050
   Icon            =   "frmBloodReceivesRecord.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10155
   ScaleWidth      =   16050
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraType 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   2280
      TabIndex        =   32
      Top             =   0
      Width           =   2535
      Begin VB.OptionButton optOccasion 
         Caption         =   "����"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   34
         Top             =   -10
         Width           =   735
      End
      Begin VB.OptionButton optOccasion 
         Caption         =   "סԺ"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   33
         Top             =   -10
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label lblType 
         Caption         =   "ʹ�ó���"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox pic11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   4440
      ScaleHeight     =   3135
      ScaleWidth      =   11055
      TabIndex        =   28
      Top             =   3120
      Width           =   11055
      Begin VB.PictureBox picTransit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   10815
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   10815
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   960
            ScaleHeight     =   255
            ScaleWidth      =   2175
            TabIndex        =   37
            Top             =   120
            Width           =   2175
            Begin VB.OptionButton optTransit 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "ָ��ʱ��"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   39
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton optTransit 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "��ǰʱ��"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin MSComCtl2.DTPicker dtpTransit 
            Height          =   330
            Left            =   3240
            TabIndex        =   40
            Top             =   75
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   283049987
            CurrentDate     =   42618
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ת��ʱ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   150
            Width           =   735
         End
      End
      Begin VB.PictureBox pic4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   105
         ScaleHeight     =   1695
         ScaleWidth      =   3615
         TabIndex        =   29
         Top             =   240
         Width           =   3615
         Begin VB.CheckBox chk2 
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   120
            Width           =   255
         End
         Begin VSFlex8Ctl.VSFlexGrid VSF2 
            Height          =   1575
            Left            =   240
            TabIndex        =   31
            Top             =   480
            Width           =   2580
            _cx             =   4551
            _cy             =   2778
            Appearance      =   2
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
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483638
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   270
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
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
      End
   End
   Begin VB.PictureBox pic7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   4320
      ScaleHeight     =   2295
      ScaleWidth      =   3255
      TabIndex        =   23
      Top             =   240
      Width           =   3255
      Begin XtremeSuiteControls.TabControl tbcthis 
         Height          =   1335
         Left            =   720
         TabIndex        =   24
         Top             =   480
         Width           =   2175
         _Version        =   589884
         _ExtentX        =   3836
         _ExtentY        =   2355
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox Pic5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   7920
      ScaleHeight     =   2535
      ScaleWidth      =   7455
      TabIndex        =   14
      Top             =   360
      Width           =   7455
      Begin VB.PictureBox pic2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   10815
         TabIndex        =   16
         Top             =   0
         Width           =   10815
         Begin VB.PictureBox pic8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   960
            ScaleHeight     =   255
            ScaleWidth      =   2175
            TabIndex        =   25
            Top             =   120
            Width           =   2175
            Begin VB.OptionButton opt1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "��ǰʱ��"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   27
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton opt1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "ָ��ʱ��"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   26
               Top             =   0
               Width           =   1095
            End
         End
         Begin MSComCtl2.DTPicker DTP3 
            Height          =   330
            Left            =   3240
            TabIndex        =   17
            Top             =   75
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   283049987
            CurrentDate     =   42618
         End
         Begin VB.Label lbl6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ʱ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   150
            Width           =   735
         End
      End
      Begin VB.PictureBox pic3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   0
         ScaleHeight     =   2055
         ScaleWidth      =   5175
         TabIndex        =   15
         Top             =   1320
         Width           =   5175
         Begin VB.CheckBox chk3 
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   0
            Width           =   255
         End
         Begin VSFlex8Ctl.VSFlexGrid VSF1 
            Height          =   1575
            Left            =   360
            TabIndex        =   19
            Top             =   240
            Width           =   2580
            _cx             =   4551
            _cy             =   2778
            Appearance      =   2
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
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483638
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   270
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
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
      End
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   360
      ScaleHeight     =   7455
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   840
      Width           =   3735
      Begin zlPublicBlood.usrCardPeople UCP 
         Height          =   3975
         Left            =   240
         TabIndex        =   42
         Top             =   3240
         Width           =   3015
         _extentx        =   5318
         _extenty        =   7011
      End
      Begin VB.Frame Fra1 
         Height          =   2895
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   3855
         Begin VB.CommandButton cmdOper 
            Height          =   240
            Left            =   3330
            Picture         =   "frmBloodReceivesRecord.frx":07AA
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "�༭(F4)"
            Top             =   1710
            Width           =   255
         End
         Begin VB.TextBox txtOper 
            Height          =   300
            Left            =   960
            TabIndex        =   5
            Top             =   1680
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker dtpEdtime 
            Height          =   300
            Left            =   960
            TabIndex        =   4
            Top             =   1320
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   283049987
            CurrentDate     =   42635
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   960
            TabIndex        =   2
            Text            =   "������"
            Top             =   600
            Width           =   2655
         End
         Begin VB.CommandButton cmd2 
            Caption         =   "������ȡ"
            Enabled         =   0   'False
            Height          =   350
            Left            =   2520
            TabIndex        =   9
            Top             =   2430
            Width           =   1100
         End
         Begin VB.CheckBox chk1 
            Caption         =   "��������"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   2490
            Width           =   1095
         End
         Begin VB.CommandButton cmd1 
            Caption         =   "ˢ��"
            Height          =   350
            Left            =   2520
            TabIndex        =   7
            Tag             =   "��"
            Top             =   2040
            Width           =   1100
         End
         Begin VB.ComboBox cboDepart 
            Height          =   300
            Left            =   960
            TabIndex        =   1
            Top             =   240
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker dtpSttime 
            Height          =   300
            Left            =   960
            TabIndex        =   3
            Top             =   960
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   283049987
            CurrentDate     =   42593
         End
         Begin VB.Label lbl5 
            Caption         =   "ȡ Ѫ ��"
            Height          =   255
            Left            =   165
            TabIndex        =   22
            Top             =   1725
            Width           =   735
         End
         Begin VB.Label lbl1 
            Caption         =   "��    ��"
            Height          =   255
            Left            =   165
            TabIndex        =   13
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lbl2 
            Caption         =   "ȡѪʱ��"
            Height          =   255
            Left            =   165
            TabIndex        =   12
            Top             =   645
            Width           =   735
         End
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�����б�"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   3375
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   9795
      Width           =   16050
      _ExtentX        =   28310
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2461
            MinWidth        =   882
            Picture         =   "frmBloodReceivesRecord.frx":08A0
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22357
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "�༭"
            TextSave        =   "�༭"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   1200
      Top             =   120
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
            Picture         =   "frmBloodReceivesRecord.frx":1386
            Key             =   "�ܾ�ת��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBloodReceivesRecord.frx":1720
            Key             =   "��ɽ���"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBloodReceivesRecord.frx":7F82
            Key             =   "�ܾ�����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBloodReceivesRecord.frx":E7E4
            Key             =   "�ȴ���Ѫ"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   360
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpPeoPle 
      Bindings        =   "frmBloodReceivesRecord.frx":15046
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBloodReceivesRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngSys As Long   '����ģ���
Private mlngMoudle As Long 'ģ���
Private mstr��ʼʱ�� As String
Private mstr����ʱ�� As String
Private mstr��д�� As String
Private mstrPrivs As String 'Ȩ�޴�
Private mblnButtonChecked As Boolean
Private mblnTextChecked As Boolean
Private mblnSizeChecked As Boolean
Private mblnStatuChecked As Boolean
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private WithEvents mclsvsf1 As clsVsf
Attribute mclsvsf1.VB_VarHelpID = -1
Private mRsBR(0 To 1) As ADODB.Recordset '��ѯ���Ĳ��˵ļ�¼��
Private mtbcThisIndex As Long '��ǰѡ�е�ѡ�
Private mintOutPreTime As Long  '��ǰѡ�е�ʱ��
Private marrPreCardID(0 To 1) As String
Private mfrmMain As Object
Private mblnת�� As Boolean
Private mint���� As Integer '0:�����סԺ;1-����;2-סԺ
Public mblnBloodReceivesIsOpen As Boolean '��ģ̬״̬�£��жϴ����Ƿ���
Private mrs���� As ADODB.Recordset
Private mblnHavePrivs As Boolean    '�Ƿ������п��ҵ�Ȩ�ޣ��Ҳ������������סԺ����
'ˢ������ʱ�Ĺ�������
Private Type Type_Filter
    DeptID As Long
    TimeIndex As Integer
    BeginTime As String
    EndTime As String
    Oper As String
End Type
Private marrFilter(0 To 1) As Type_Filter

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ������
    
    Call CommandBarInit(cbsMain)
    '�˵�����:������������
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '�ļ�
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.id = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)
    
    '�༭
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.id = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "����", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "����")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Transfer, "ת��", True)
    '�鿴
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.id = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_FindNext, "������һ��(&N)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)
    
    mblnButtonChecked = True
    mblnTextChecked = True
    mblnSizeChecked = True
    mblnStatuChecked = True
    
    '����
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.id = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True
    End With
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Transfer, "ת��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        Set objCustom = .Add(xtpControlCustom, conMenu_View_FindType, "����")
        objCustom.Handle = fraType.hWnd
        objCustom.Flags = xtpFlagRightAlign
    End With
    For Each objControl In objBar.Controls
        If objControl.Type = xtpControlButton Then objControl.Style = xtpButtonIconAndCaption
    Next
    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
        .Add FCONTROL, vbKeyF, conMenu_View_Find           '����
        .Add 0, vbKeyF3, conMenu_View_FindNext         '��������
        .Add 0, vbKeyF5, conMenu_View_Refresh      'ˢ��
    End With
    
    Call gobjDatabase.ShowReportMenu(Me, 100, pѪҺ���յǼ�, mstrPrivs)
    InitCommandBar = True
    
    Exit Function
ErrHand:
End Function

Private Sub cboDepart_KeyPress(KeyAscii As Integer)
'    If KeyAscii <> vbKeyReturn Then KeyAscii = 0: Exit Sub
    Dim lngi As Long
    Dim blnisread  As Boolean
    Dim rs���� As ADODB.Recordset
    If KeyAscii = 39 Then KeyAscii = 0: Exit Sub '����������"'"
    blnisread = False
    If KeyAscii = vbKeyReturn Then
        For lngi = 0 To cboDepart.ListCount - 1
            If cboDepart.List(lngi) Like cboDepart.Text & "*" Or cboDepart.List(lngi) Like "*" & cboDepart.Text & "*" Then
                cboDepart.Text = cboDepart.List(lngi)
                cboDepart.Tag = cboDepart.ListIndex
                cboDepart.ListIndex = lngi
                blnisread = True
                Exit For
            End If
        Next
        If blnisread = False Then
            Call CopyRecord(mrs����, rs����)
            rs����.Filter = "���� like '" & cboDepart.Text & "%'"
            If rs����.RecordCount > 0 Then
                For lngi = 0 To cboDepart.ListCount - 1
                    If cboDepart.ItemData(lngi) = rs����!id Then
                        cboDepart.ListIndex = lngi
                        cboDepart.Tag = lngi
                        blnisread = True
                        Exit For
                    End If
                Next
            End If
        End If
        If blnisread = False Then
            cboDepart.ListIndex = IIf(Val(cboDepart.Tag) < 0, 0, Val(cboDepart.Tag))
            Exit Sub
        End If
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub cboDepart_Validate(Cancel As Boolean)
    '�������Ƿ���ȷ
    Dim lngi As Long
    Dim blnIsSelect As Boolean
    blnIsSelect = False
    If mblnHavePrivs = False Then Exit Sub '�������û�����Ȩ�����˳�
    For lngi = 0 To cboDepart.ListCount - 1
        If cboDepart.Text = cboDepart.List(lngi) Or cboDepart.Text = "���в���" Then
            blnIsSelect = True
            cboDepart.Tag = cboDepart.ListIndex
            Exit For
        End If
    Next
    If blnIsSelect = False Then
        cboDepart.ListIndex = IIf(Val(cboDepart.Tag) < 0, 0, Val(cboDepart.Tag))
        Exit Sub
    End If
End Sub

Private Sub cboTime_Click()
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    If cboTime.ListIndex < 0 Then Exit Sub
    If mintOutPreTime = cboTime.ListIndex And cboTime.ListIndex <> cboTime.ListCount - 1 Then Exit Sub
    intDateCount = cboTime.ItemData(cboTime.ListIndex)
    datCurr = Format(gobjDatabase.Currentdate, "yyyy-MM-dd")
    If Me.Visible Then
        If intDateCount = -1 Then
        ElseIf intDateCount = 0 Then
            mstr��ʼʱ�� = Format(datCurr, "yyyy-MM-dd 00:00:00")
            mstr����ʱ�� = Format(datCurr, "yyyy-MM-dd 23:59:59")
        Else
            mstr��ʼʱ�� = Format(datCurr - intDateCount, "yyyy-MM-dd 00:00:00")
            mstr����ʱ�� = Format(datCurr, "yyyy-MM-dd 23:59:59")
        End If
        dtpSttime.Value = Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm")
        dtpEdtime.Value = Format(mstr����ʱ��, "YYYY-MM-DD HH:mm")
        dtpSttime.Enabled = (cboTime.ItemData(cboTime.ListIndex) = -1): dtpEdtime.Enabled = dtpSttime.Enabled
    End If
    mintOutPreTime = cboTime.ListIndex
End Sub

Private Sub cboTime_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lng����ID As Long, lng��ҳid As Long
    Dim strTmp As String
    
    Select Case Control.id
        Case conMenu_File_PrintSet
            Call zlPrintSet
        Case conMenu_File_Preview
            Call zlRptPrint(2, IIf(mtbcThisIndex = 0, VSF1, VSF2), IIf(mtbcThisIndex = 0, "������ѪҺ", "�ѽ���ѪҺ"))
        Case conMenu_File_Print
            Call zlRptPrint(1, IIf(mtbcThisIndex = 0, VSF1, VSF2), IIf(mtbcThisIndex = 0, "������ѪҺ", "�ѽ���ѪҺ"))
        Case conMenu_Edit_Audit: '����
            If mblnת�� = False Then '��������
                If ExecuteCommand("��������") = True Then Call ExecuteCommand("ˢ������")
                chk3.Value = 0
            Else 'ת�ӽ���
                If ExecuteCommand("ת��") = True Then Call ExecuteCommand("ˢ������")
            End If
        Case conMenu_Edit_Untread: '����
            If ExecuteCommand("����") = True Then Call ExecuteCommand("ˢ������")
            chk2.Value = 0
        Case conMenu_View_Refresh: 'ˢ��
            cmd1_Click
         Case conMenu_View_Find, conMenu_View_FindNext '���ң���������
            Call UCP.FindPatiByVbKey(Control.id = conMenu_View_FindNext)
        Case conMenu_View_ToolBar_Button '��׼��ť
            mblnButtonChecked = Not mblnButtonChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_View_ToolBar_Text '�ı���ǩ
            mblnTextChecked = Not mblnTextChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_View_ToolBar_Size '��ͼ��
            mblnSizeChecked = Not mblnSizeChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_View_StatusBar '״̬��
            mblnStatuChecked = Not mblnStatuChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_Help_Help              '��������
            Call gobjComlib.ShowHelp(App.ProductName, Me.hWnd, Me.name, Int((100) / 100))
        Case conMenu_Help_Web_Home 'Web�ϵ�����
            Call gobjComlib.zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Forum         'Web�ϵ���̳
            Call gobjComlib.zlWebForum(Me.hWnd)
        Case conMenu_Help_Web_Mail '���ͷ���
            Call gobjComlib.zlMailTo(Me.hWnd)
        Case conMenu_Help_About '����
            Call gobjComlib.ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_File_Exit '�˳�
            Unload Me
        Case conMenu_Manage_Transfer 'ת��
            mblnת�� = Not mblnת�� 'ת������ģʽ
            If mblnת�� Then optTransit(0).Value = True
            Call SetVsf2State
        Case Else
            If Between(Control.id, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                strTmp = UCP.strReturn
                If strTmp <> "" Then
                    strTmp = Split(strTmp, "'")(1)
                    lng����ID = Split(strTmp, "-")(0)
                    lng��ҳid = Split(strTmp, "-")(1)
                    'ִ�з�������ǰģ��ı���
                    Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me, "����ID=" & lng����ID & ",����ID=" & lng��ҳid)
                End If
            End If
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case conMenu_Edit_Audit: '����
            If mblnת�� = False Then
                Control.Caption = "����"
                Control.Visible = IsPrivs(mstrPrivs, "��������")
                Control.Enabled = IIf(mtbcThisIndex = 0, True, False) And Control.Visible
            Else
                Control.Caption = "����"
                Control.Enabled = IsClick
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Untread: '����
            Control.Visible = IsPrivs(mstrPrivs, "��������")
            Control.Enabled = IIf(mtbcThisIndex = 0, False, True) And Control.Visible And Not mblnת��
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_Transfer: 'ת��
            Control.Visible = IsPrivs(mstrPrivs, "ת������")
            Control.Enabled = IIf(mtbcThisIndex = 0, False, True) And Control.Visible
            Control.Checked = mblnת��
            picTransit.Visible = mblnת��
            Call pic11_Resize
        Case conMenu_View_ToolBar_Button
            Control.Checked = mblnButtonChecked
        Case conMenu_View_ToolBar_Text
            Control.Checked = mblnTextChecked
        Case conMenu_View_ToolBar_Size
            Control.Checked = mblnSizeChecked
        Case conMenu_View_StatusBar
            stbThis.Visible = mblnStatuChecked
            Control.Checked = mblnStatuChecked
    End Select
End Sub

Private Sub chk1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub chk2_Click()
    'ȫѡ��ȫ��
    Dim lngi As Long
    
    For lngi = 1 To VSF2.Rows - 1
        If mblnת�� = False Then
            If Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("����id"))) <> 0 And (UserInfo.���� = VSF2.TextMatrix(lngi, VSF2.ColIndex("������")) Or UserInfo.���� = VSF2.TextMatrix(lngi, VSF2.ColIndex("������"))) Then
                VSF2.TextMatrix(lngi, VSF2.ColIndex("ѡ��")) = chk2.Value
            End If
        Else
            If Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("״̬"))) <> 1 And Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("״̬"))) <> 3 And Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("����id"))) <> 0 Then
                VSF2.TextMatrix(lngi, VSF2.ColIndex("ѡ��")) = chk2.Value
            End If
        End If
    Next
End Sub

Private Sub chk2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '��ʾ��Ϣ
    Dim strInfo As String
    If mblnת�� = True Then
        strInfo = "�ɶ��ѽ��յ�ѪҺ����ת��" & vbCrLf & "����������ʱ������ϴ�ʱ�䣻���˵�ǰ���ҷ����仯"
    Else
        strInfo = "�ɶ��Լ����ջ�˶Ե����ݽ��л���"
    End If
    gobjCommFun.ShowTipInfo chk2.hWnd, strInfo
End Sub

Private Sub chk3_Click()
    'ȫѡ��ȫ��
    Dim lngi As Long
    For lngi = 1 To VSF1.Rows - 1
        If Val(VSF1.TextMatrix(lngi, VSF1.ColIndex("����id"))) <> 0 Then
            VSF1.TextMatrix(lngi, VSF1.ColIndex("ѡ��")) = chk3.Value
        End If
    Next
End Sub

Private Sub cmd1_Click()
    'ˢ������
    If Format(dtpSttime.Value, "YYYY-MM-DD HH:mm") > Format(dtpEdtime.Value, "YYYY-MM-DD HH:mm") Then
        MsgBox "���������еĿ�ʼʱ�䲻�ܴ��ڽ���ʱ�䣬�������", vbInformation, gstrSysName
        Exit Sub
    End If
    If mblnHavePrivs = False Then Exit Sub
    mstr��ʼʱ�� = Format(dtpSttime.Value, "YYYY-MM-DD HH:mm")
    mstr����ʱ�� = Format(dtpEdtime.Value, "YYYY-MM-DD HH:mm")
    
    With marrFilter(mtbcThisIndex)
        .DeptID = cboDepart.ItemData(cboDepart.ListIndex)
        .TimeIndex = cboTime.ListIndex
        .BeginTime = mstr��ʼʱ��
        .EndTime = mstr����ʱ��
        .Oper = txtOper.Text
    End With
    '���VSF1��VSF2�ϵ�����
    If mtbcThisIndex = 0 Then
        VSF1.Rows = 2
        VSF1.RowData(1) = 0
        VSF1.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
    Else
        VSF2.Rows = 2
        VSF2.RowData(1) = 0
        VSF2.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
    End If
    
    Call ExecuteCommand("�������˲�ѯ") ','' as ����,'' as ����,'' as ����
End Sub

Private Sub ShowVsf(lngstatu As Long, Optional rs As ADODB.Recordset, Optional strP As String)
    '���ܣ���vsf����ʾ��Ӧ���ݣ���Ϊ��ѡ�͵�ѡ������ʽ��ͬ
    '������lngstatu:0-������ѪҺ��1-�ѽ���ѪҺ��rs��strp����ucp�ķ������ݣ��������ѡ��,������ѡ����ֻ��ѡ������һ�����������������ڣ���������
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strSql1 As String
    Dim lng����ID As Long, lng��ҳid As Long
    Dim strValues As String, strTmp As String
    Dim blnBatch As Boolean
    
    On Error GoTo ErrHand
    
    If chk1.Value = Checked Then  '��ѡ�������
        blnBatch = True
        If rs Is Nothing Then Exit Sub
        If rs.State = adStateClosed Then Exit Sub
        If rs.RecordCount = 0 Then Exit Sub
    Else
        blnBatch = False
        If strP = "" Then Exit Sub
    End If
    
    'ѡ���ͬ��ѯ���ݲ�ͬ
    If lngstatu = 0 Then '��ͬ��ѡ�������������ͬ
        strSql1 = " And e.����ʱ��+0 between [3] and [4] And nvl(e.����״̬,0) =0  "
        If Trim(marrFilter(mtbcThisIndex).Oper) <> "" Then
            strSql1 = strSql1 & " And e.ȡѪ��=[6]"
        End If
    Else
        strSql1 = " And e.����ʱ��+0 between [3] and [4] And nvl(e.����״̬,0)<>0 "
        If marrFilter(mtbcThisIndex).DeptID <> -1 Then
            strSql1 = strSql1 & " And E.ִ�п���id=[5]"
        End If
        If Trim(marrFilter(mtbcThisIndex).Oper) <> "" Then
            strSql1 = strSql1 & " And e.������=[6]"
        End If
    End If
    
    If optOccasion(0).Value = True Then
        strSQL = _
            " Select " & IIf(blnBatch = True, " /*+ CARDINALITY(T 10) */ ", "") & " e.�շ�id, d.����id, d.��ҳid, 0 As ѡ��, g.��Ժ���� As ����, g.����, g.�Ա�,g.��Ժ����ID ��ǰ����ID, a.���� As ѪҺ����, a.���, f.Ѫ�����, f.Abo, f.Rh," & vbNewLine & _
            "       Decode(e.����״̬, 0, '���ڽ���', 1, '�Ѿ�����', 2, '�ܾ�����', 'ת�ƽ���') As ����״̬, " & IIf(lngstatu = 1, "h.����", "''") & " As ִ�в���, e.������, to_char(e.����ʱ��,'yyyy-mm-dd HH24:mi:ss') as ����ʱ��, e.������, e.ִ�п���id," & vbNewLine & _
            "       e.ȡѪ��, e.�ܾ�ԭ��" & vbNewLine & _
            " From �շ���ĿĿ¼ a, ѪҺƷ�� b, ѪҺ��� c" & IIf(lngstatu = 1, ",���ű� h", "") & ", ѪҺ�շ���¼ f, ѪҺ���ͼ�¼ e, ѪҺ��Ѫ��¼ d, ������ҳ g" & IIf(blnBatch = True, ",Table(f_Num2list2([1])) T", "") & vbNewLine & _
            " Where a.Id = c.���id And c.Ʒ��id = b.Ʒ��id And c.���id = f.ѪҺid And f.Id = e.�շ�id" & IIf(lngstatu = 1, " And E.ִ�п���id=h.id(+) ", " ") & " And e.�䷢id = d.Id And d.����id = g.����id And" & vbNewLine & _
            "      d.��ҳid = g.��ҳid "
        If blnBatch = True Then
            strSQL = strSQL & " And g.����ID=T.C1 and g.��ҳID=T.C2"
        Else
            strSQL = strSQL & " And g.����id = [1] And g.��ҳid = [2]"
        End If
    Else
        strSQL = _
            " Select " & IIf(blnBatch = True, " /*+ CARDINALITY(T 10) */ ", "") & " e.�շ�id, d.����id, d.Id ��ҳid, 0 As ѡ��, '' As ����, g.����, g.�Ա�,g.ִ�в���ID ��ǰ����ID, a.���� As ѪҺ����, a.���, f.Ѫ�����, f.Abo, f.Rh," & vbNewLine & _
            "       Decode(e.����״̬, 0, '���ڽ���', 1, '�Ѿ�����', 2, '�ܾ�����', 'ת�ƽ���') As ����״̬, " & IIf(lngstatu = 1, "h.����", "''") & " As ִ�в���, e.������, to_char(e.����ʱ��,'yyyy-mm-dd HH24:mi:ss') as ����ʱ��, e.������, e.ִ�п���id," & vbNewLine & _
            "       e.ȡѪ��, e.�ܾ�ԭ��" & vbNewLine & _
            " From �շ���ĿĿ¼ a, ѪҺƷ�� b, ѪҺ��� c" & IIf(lngstatu = 1, ",���ű� h", "") & ", ѪҺ�շ���¼ f, ѪҺ���ͼ�¼ e, ѪҺ��Ѫ��¼ d, ����ҽ����¼ k, ���˹Һż�¼ g" & IIf(blnBatch = True, ",Table(f_Num2list2([1])) T", "") & vbNewLine & _
            " Where a.Id = c.���id And c.Ʒ��id = b.Ʒ��id And c.���id = f.ѪҺid And f.Id = e.�շ�id " & IIf(lngstatu = 1, " And E.ִ�п���id=h.id(+) ", " ") & " And e.�䷢id = d.Id And d.����id = k.Id And" & vbNewLine & _
            "       k.�Һŵ� = g.No And k.����id = g.����id And k.������� = 'K' "
        If blnBatch = True Then
            strSQL = strSQL & " And g.����ID=T.C1 and g.Id=T.C2"
        Else
            strSQL = strSQL & " And g.����id = [1] And g.Id = [2]"
        End If
    End If
    strSQL = strSQL & strSql1
    Screen.MousePointer = 11
    If blnBatch Then '��ѡ�������
        strValues = ""
        Do While Not rs.EOF
            strTmp = rs.Fields("ID").Value
            lng����ID = Val(Split(strTmp, "-")(0))
            lng��ҳid = Val(Split(strTmp, "-")(1))
            strValues = strValues & "," & lng����ID & ":" & lng��ҳid
            rs.MoveNext
        Loop
        If Left(strValues, 1) = "," Then strValues = Mid(strValues, 2)
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "ѪҺ��Ϣ", strValues, "", CDate(marrFilter(mtbcThisIndex).BeginTime), CDate(marrFilter(mtbcThisIndex).EndTime), marrFilter(mtbcThisIndex).DeptID, Trim(marrFilter(mtbcThisIndex).Oper))
    Else
        lng����ID = Val(Split(Split(strP, "'")(1), "-")(0))
        lng��ҳid = Val(Split(Split(strP, "'")(1), "-")(1))
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "ѪҺ��Ϣ", lng����ID, lng��ҳid, CDate(marrFilter(mtbcThisIndex).BeginTime), CDate(marrFilter(mtbcThisIndex).EndTime), marrFilter(mtbcThisIndex).DeptID, Trim(marrFilter(mtbcThisIndex).Oper))
    End If
    If lngstatu = 0 Then
        Call mclsVsf.LoadGrid(rsTemp)
    Else
        Call mclsvsf1.LoadGrid(rsTemp)
        Call SetVsf2State
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrHand:
    Screen.MousePointer = 0
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetVsf2State()
    '����vsf2�ϵ����ݵ������ȷ��������״̬�����и���״̬�ı�ÿ�����ݵ�ͼ�ꡣ
    Dim lngi As Long, lng״̬ As Long
    For lngi = 1 To VSF2.Rows - 1
        VSF2.TextMatrix(lngi, VSF2.ColIndex("ѡ��")) = 0
        VSF2.Cell(flexcpPicture, lngi, 4, lngi, 4) = Nothing
        If Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("����id"))) <> 0 Then
            If VSF2.TextMatrix(lngi, VSF2.ColIndex("����״̬")) = "�ܾ�����" Then
                lng״̬ = 3
            ElseIf VSF2.TextMatrix(lngi, VSF2.ColIndex("����״̬")) = "�Ѿ�����" Then
                lng״̬ = 2
            ElseIf VSF2.TextMatrix(lngi, VSF2.ColIndex("����״̬")) = "ת�ƽ���" Then
                lng״̬ = 4
            End If
            '������״̬Ϊ�������պ�ת�ƽ��գ���ִ�п���IDδ�ı�ʱ(����δת��)��������ת��
            If mblnת�� = True Then 'ת��ģʽ�Ὣ������ת�ӵ����ݱ�עΪ"!"
                If Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("ִ�п���id"))) = Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("��ǰ����ID"))) Then
                    lng״̬ = 1
                    VSF2.Cell(flexcpForeColor, lngi, 1, lngi, VSF2.Cols - 1) = vbGrayText
                End If
            Else
                VSF2.Cell(flexcpForeColor, lngi, 1, lngi, VSF2.Cols - 1) = vbBlack
            End If
            VSF2.TextMatrix(lngi, VSF2.ColIndex("״̬")) = lng״̬
            If lng״̬ <> 0 Then
                VSF2.Cell(flexcpPicture, lngi, VSF2.ColIndex("ͼ��"), lngi, VSF2.ColIndex("ͼ��")) = ils16.ListImages(lng״̬).Picture
            End If
        End If
    Next
End Sub

Private Sub GetSendorReceivePeople(ByVal objControl As TextBox, Optional ByVal StrInput As String = "")
    '���ܣ�����ʱ�䷶Χ����ȡѪ�˺ͽ�����
    Dim strSQL As String, strSQLNew As String
    Dim rsUser  As ADODB.Recordset
    Dim lngDeptID As Long, strWhere As String
    Dim vPoint As RECT, blnCancel As Boolean
    On Error GoTo ErrHand
    
    '�����������жϰ���ʲô��ʽ�������ݣ���������š�����
    If StrInput <> "" Then
         If IsNumeric(StrInput) Then
            strWhere = " And C.��� Like [4]"
         ElseIf gobjCommFun.IsCharAlpha(StrInput) Then
            strWhere = " And C.���� Like [4]"
            StrInput = UCase(StrInput)
         Else
            strWhere = " And C.���� Like [4]"
         End If
    End If
    
    '��ѯ���
    If mtbcThisIndex = 0 Then
        strSQL = "Select Distinct ȡѪ�� as ���� From ѪҺ���ͼ�¼ where ����ʱ�� between [1] And [2] And NVL(����״̬,0)=0"
    Else
        strSQL = "Select Distinct ������ as ���� From ѪҺ���ͼ�¼ where ����ʱ�� between [1] And [2] And NVL(����״̬,0)<>0"
    End If
    
    If IsPrivs(mstrPrivs, "���п���") And Val(cboDepart.ItemData(cboDepart.ListIndex)) = -1 Then
        lngDeptID = 0
        strSQLNew = _
        " Select distinct c.id,c.����,C.����" & vbNewLine & _
        " From ��Ա�� c,(" & strSQL & ") d" & vbNewLine & _
        " Where c.����=d.����" & strWhere
    Else
        lngDeptID = Val(cboDepart.ItemData(cboDepart.ListIndex))
        strSQLNew = _
        " Select distinct c.id,c.����,C.����" & vbNewLine & _
        " From ���ű� a, ������Ա b, ��Ա�� c,(" & strSQL & ") d" & vbNewLine & _
        " Where a.Id = b.����id And b.��Աid = c.Id And a.Id = [3] And c.����=d.����" & strWhere
    End If
    vPoint = GetControlRect(objControl.hWnd)
    Set rsUser = gobjDatabase.ShowSQLSelect(Me, strSQLNew, 0, "", False, "", "��ѡ��һ��" & IIf(mtbcThisIndex = 0, "ȡѪ", "����") & "��Ա", False, False, True, vPoint.Left, vPoint.Top, objControl.Height, blnCancel, False, False, CDate(Format(dtpSttime.Value, "YYYY-MM-DD HH:mm")), CDate(Format(dtpEdtime.Value, "YYYY-MM-DD HH:mm")), lngDeptID, StrInput & "%")
    If Not rsUser Is Nothing Then
        If blnCancel = False Then
            If rsUser.EOF Then Exit Sub
            objControl.Text = rsUser!���� & ""
            objControl.Tag = objControl.Text
        End If
    End If
    
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub


Public Sub BloodReceives(frmMain As Variant, ByVal lngSys As Long, ByVal lngMoudle As Long, Optional strPrivs As String, Optional lngisModul As Long = 0, Optional int���� As Integer = 0)
    '���ܣ�ѪҺ���յǼ���ں���
    Dim lngi As Long
    Dim strSQL As String
    Dim strȡѪ�� As String
    Dim rsȡѪ�� As ADODB.Recordset
    Dim rs��Ա
    Dim objPane As Pane

    If mblnBloodReceivesIsOpen = True Then GoTo TOSHOW
    
    mlngSys = lngSys
    mlngMoudle = lngMoudle
    mstrPrivs = strPrivs
    mint���� = int����
    mblnHavePrivs = False
    Set mfrmMain = frmMain
    
    InitCommandBar '��ʼ��commandbar
'    '��ʼ��DockingPane
    Call DockPannelInit(dkpPeoPle)
    dkpPeoPle.SetCommandBars cbsMain
    Set objPane = dkpPeoPle.CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "����": objPane.Options = PaneNoCaption
    Set objPane = dkpPeoPle.CreatePane(2, 800, 100, DockRightOf, Nothing): objPane.Title = "��¼": objPane.Options = PaneNoCaption
    
    '��ʼ��tbcthis
    mtbcThisIndex = 0
    With tbcthis
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .OneNoteColors = True
            .Position = xtpTabPositionTop
            .ShowIcons = False
        End With
        .InsertItem(0, "������ѪҺ", Pic5.hWnd, 0).Tag = "������ѪҺ"
        .InsertItem(1, "�ѽ���ѪҺ", pic11.hWnd, 0).Tag = "�ѽ���ѪҺ"
        .Item(0).Selected = True
        Call tbcThis_SelectedChanged(tbcthis.Selected)
    End With
TOSHOW:
    mblnBloodReceivesIsOpen = True
    If IsObject(frmMain) Then
        If frmMain Is Nothing Then
            Me.Show lngisModul
        Else
            Me.Show lngisModul, frmMain
        End If
    Else
        gobjComlib.os.ShowChildWindow Me.hWnd, Val(frmMain)
    End If
End Sub

Private Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
End Function


Private Sub RsTitelCopy(ByVal RsProm As ADODB.Recordset, ToRs As ADODB.Recordset)
    '���ܣ��½�ToRs��¼������RsProm�Ľṹ���Ƶ�ToRs��
    '������RsProm-ԭ��¼����ToRs-�½��ļ�¼��
    Dim lngi As Long
    Set ToRs = New ADODB.Recordset
    With ToRs '��ʼ��rsReturn
        For lngi = 0 To RsProm.Fields.Count - 1
            .Fields.Append RsProm.Fields(lngi).name, adLongVarChar, 100, adFldIsNullable
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    Dim intLoop As Integer
    Dim lngi As Long, lngj As Long
    Dim rsSAD As New ADODB.Recordset
    Dim StrSqlSAD As String
    Dim strSQL As String, strSql1 As String, strSql2 As String
    Dim rsTmp As ADODB.Recordset
    Dim lngColor As Long
    Dim strABORH As String
    Dim rsBR As ADODB.Recordset
    Dim blnSelect As Boolean, lngDeptID As Long, strOpter As String
    Dim strCurDate As String, strRows As String
    
    On Error GoTo Error
    
    Call SQLRecord(rsSAD)
    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
            Case "�������˲�ѯ":
                If mtbcThisIndex = 0 Then
                    strSql1 = "Select Distinct �䷢id From ѪҺ���ͼ�¼ Where ����ʱ�� Between [1] And [2] And nvl(����״̬,0)=0"
                    If Trim(marrFilter(mtbcThisIndex).Oper) <> "" Then
                        strSql1 = strSql1 & " And ȡѪ��=[3]"
                    End If
                Else
                    strSql1 = "Select Distinct �䷢id From ѪҺ���ͼ�¼ Where ����ʱ�� Between [1] And [2] And nvl(����״̬,0)<>0"
                    If Trim(marrFilter(mtbcThisIndex).Oper) <> "" Then
                        strSql1 = strSql1 & " And ������=[3]"
                    End If
                End If
                
                If optOccasion(0).Value = True Then 'סԺ����
                    If marrFilter(mtbcThisIndex).DeptID <> -1 And mtbcThisIndex = 0 Then
                        'ĳ�����Ų鿴���迼�ǲ���ת�Ƶ������ת��ǰ��ת���Ŀ��Ҷ����Խ���ѪҺ
                        strSql1 = "Select Distinct �䷢id,����ʱ�� From ѪҺ���ͼ�¼ Where ����ʱ�� Between [1] And [2] And nvl(����״̬,0)=0"
                        strSQL = _
                            " Select a.����id || '-' || a.��ҳID As Id,a.����ID, a.��ҳid, a.סԺ�� ������,  Decode(Nvl(a.��������, 0), 0, 'ס', '��')As סԺ���, a.����, a.�Ա� || '/' || a.���� As �Ա�����, a.��Ժ����id As ����id, d.����," & vbNewLine & _
                            "       a.��Ժ���� As ����, a.��Ժ���� As ����, a.����, a.�������� As ����, 255 As ��ɫ, '' As Aborh" & vbNewLine & _
                            " From ������ҳ a, ���ű� d," & vbNewLine & _
                            "     (Select Distinct b.����id, b.��ҳid" & vbNewLine & _
                            "       From ѪҺ��Ѫ��¼ b, (" & strSql1 & ") c" & vbNewLine & _
                            "       Where b.Id = c.�䷢id And b.��ҳid Is Not Null And Exists" & vbNewLine & _
                            "        (Select 1" & vbNewLine & _
                            "              From ���˱䶯��¼" & vbNewLine & _
                            "              Where b.����id = ����id And b.��ҳid = ��ҳid And ����id = [4] And" & vbNewLine & _
                            "                    Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) >= c.����ʱ��)) b" & vbNewLine & _
                            " Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.��Ժ����id = d.Id(+)"
                    Else
                        If marrFilter(mtbcThisIndex).DeptID <> -1 Then
                            strSql1 = strSql1 & " And ִ�п���ID=[4]"
                        End If
                        strSQL = _
                            " Select Distinct a.����id || '-' || a.��ҳID As Id,a.����ID, a.��ҳid, a.סԺ�� ������,  Decode(Nvl(a.��������, 0), 0, 'ס', '��')As סԺ���, a.����, a.�Ա� || '/' || a.���� As �Ա�����, a.��Ժ����id As ����id, d.����," & vbNewLine & _
                            "                a.��Ժ���� As ����, a.��Ժ���� As ����, a.����, a.�������� As ����, 255 As ��ɫ, '' As Aborh" & vbNewLine & _
                            " From ������ҳ a, ���ű� d, ѪҺ��Ѫ��¼ b, (" & strSql1 & ") c" & vbNewLine & _
                            " Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.��Ժ����id = d.Id(+) And b.Id = c.�䷢id And b.��ҳid Is Not Null"
                    End If
                Else '���ﲡ��
                    If marrFilter(mtbcThisIndex).DeptID <> -1 Then
                        If mtbcThisIndex = 0 Then
                            strSql2 = strSql2 & " And a.ִ�в���id=[4]"
                        Else
                            strSql1 = strSql1 & " And ִ�п���ID=[4]"
                        End If
                    End If
                    
                    strSQL = _
                        " Select Distinct a.����id || '-' || A.id As Id, A.����ID,A.ID ��ҳID, a.����� ������, Decode(a.����,1,'��',decode(a.����,1,'��','��')) As סԺ���, a.����, a.�Ա� || '/' || a.���� As �Ա�����, a.ִ�в���id As ����id, e.���� ��������," & vbNewLine & _
                        "                a.ִ��ʱ�� As ����, '' As ����, a.����, '' As ����, 255 As ��ɫ, a.ִ����, '' As Aborh" & vbNewLine & _
                        " From ���˹Һż�¼ a, ���ű� e, ����ҽ����¼ b, ѪҺ��Ѫ��¼ c,(" & strSql1 & ") d" & vbNewLine & _
                        " Where a.No = b.�Һŵ�" & strSql2 & " And a.ִ�в���id = e.Id(+) And b.Id = c.����id And c.Id = d.�䷢id And c.��ҳid Is Null"
                End If
                
                Set rsBR = gobjDatabase.OpenSQLRecord(strSQL, "������Ϣ", CDate(marrFilter(mtbcThisIndex).BeginTime), CDate(marrFilter(mtbcThisIndex).EndTime), Trim(marrFilter(mtbcThisIndex).Oper), marrFilter(mtbcThisIndex).DeptID)
                Call RsTitelCopy(rsBR, mRsBR(mtbcThisIndex))  '�½���¼��mRsBR����ṹ����rsBR
                '�����ǶԲ�ѯ�������ݽ��д��������������ݸ�ֵ��ͨ��RsTitelCopy�������ɵļ�¼��
                With mRsBR(mtbcThisIndex)
                    If rsBR.RecordCount > 0 Then
                        For lngi = 0 To rsBR.RecordCount - 1
                            .AddNew
                            For lngj = 0 To rsBR.Fields.Count - 1
                                .Fields(lngj).Value = rsBR.Fields(lngj).Value
                                
                                If .Fields(lngj).name = "����" Then
                                    .Fields(lngj).Value = Format(rsBR.Fields("����").Value, "YYYY-MM-DD HH:mm")
                                End If
                                
                                strABORH = ""
                                If .Fields(lngj).name = "ABORH" Then '���¸�ABORH��ֵ
                                    Set rsTmp = GetPatientOtherInfo(Val(rsBR.Fields("ID").Value), "ABO")
                                    If rsTmp.BOF = False Then strABORH = rsTmp("��Ϣֵ").Value
                                    If strABORH = "" Then '���ﲡ�˲�ѯѪ��
                                        Set rsTmp = GetPatientOtherInfo(Val(rsBR.Fields("ID").Value), "Ѫ��")
                                        If rsTmp.BOF = False Then strABORH = rsTmp("��Ϣֵ").Value
                                    End If
                                    Set rsTmp = GetPatientOtherInfo(Val(rsBR.Fields("ID").Value), "RH")
                                    If rsTmp.BOF = False Then strABORH = strABORH & rsTmp("��Ϣֵ").Value 'ABO&RH
                                     .Fields("ABORH").Value = strABORH
                                End If
                                
                                If .Fields(lngj).name = "��ɫ" Then '���¸������ͺ����������ɫ
                                    If Not IsNull(rsBR!����) And rsBR!���� & "" = "" Then
                                        '������ɫ
                                        lngColor = &HC0&
                                    Else
                                        lngColor = gobjDatabase.GetPatiColor(Nvl(rsBR!����))
                                    End If
                                    .Fields("��ɫ").Value = lngColor
                                End If
                            Next
                            .Update
                            rsBR.MoveNext
                        Next
                        .MoveFirst
                    End If
                End With
                If mblnHavePrivs = True Then
                    UCP.ShowPeople mRsBR(mtbcThisIndex)
                    If marrPreCardID(mtbcThisIndex) <> "" Then Call UCP.SetCardFocus("ID", marrPreCardID(mtbcThisIndex))
                End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "��ʼ���"
            
                Set mclsVsf = New clsVsf
                With mclsVsf
                    Call .Initialize(Me.Controls, VSF1, True, True)
                    Call .ClearColumn
                    Call .AppendColumn("�շ�id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("����id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("��ҳid", 0, flexAlignRightCenter, flexDTString, "", "", True, , , True)
                    
                    Call .AppendColumn("", 400, flexAlignLeftCenter, flexDTBoolean, "", "ѡ��", True)
                    Call .AppendColumn("", 0, flexAlignLeftCenter, flexDTString, "", "ͼ��", True)
                    Call .AppendColumn("", 0, flexAlignLeftCenter, flexDTString, "", "״̬", True)
                    Call .AppendColumn("����", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("����", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("�Ա�", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ѪҺ����", 1800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("���", 1000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("Ѫ�����", 1000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ABO", 600, flexAlignLeftCenter, flexDTString, , "ABO", True)
                    Call .AppendColumn("Rh", 600, flexAlignLeftCenter, flexDTString, , "RH", True)
                    Call .AppendColumn("����״̬", 1200, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ִ�в���", 1000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("������", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("����ʱ��", 1800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("������", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ִ�п���id", 0, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ȡѪ��", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("�ܾ�ԭ��", 2000, flexAlignLeftCenter, flexDTString, "�ܾ�ԭ��", "", True)
                    Call .AppendColumn("��ǰ����ID", 0, flexAlignRightCenter, flexDTString, "", "", True, , , True)
                    
                    .AppendRows = False
                    .SysHidden(.ColIndex("�շ�id")) = True
                    .SysHidden(.ColIndex("����id")) = True
                    .SysHidden(.ColIndex("��ҳid")) = True

                    Call .InitializeEdit(True, True, True)
                    Call .InitializeEditColumn(.ColIndex(""), True, vbVsfEditCheck)
                    Call .InitializeEditColumn(.ColIndex("����״̬"), True, vbVsfEditCombox, "���ڽ���|�ܾ�����")
                    Call .InitializeEditColumn(.ColIndex("�ܾ�ԭ��"), True, vbVsfEditText)
                    
                End With
                
                Set mclsvsf1 = New clsVsf
                With mclsvsf1
                    Call .Initialize(Me.Controls, VSF2, True, True) ', frmPubResource.GetImageList(16)����û��frmpubresource�������ｫ֮ȥ��
                    Call .ClearColumn
                    Call .AppendColumn("�շ�id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("����id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("��ҳid", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("", 400, flexAlignLeftCenter, flexDTBoolean, "", "ѡ��", True)
                    Call .AppendColumn("", 300, flexAlignLeftCenter, flexDTString, "ͼ��", "ͼ��", True)
                    Call .AppendColumn("", 0, flexAlignLeftCenter, flexDTString, "״̬", "״̬", True)
                    Call .AppendColumn("����", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("����", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("�Ա�", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ѪҺ����", 1800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("���", 1000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("Ѫ�����", 1000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ABO", 600, flexAlignLeftCenter, flexDTString, , "ABO", True)
                    Call .AppendColumn("Rh", 600, flexAlignLeftCenter, flexDTString, , "RH", True)
                    Call .AppendColumn("����״̬", 1200, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ִ�в���", 1000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("������", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("����ʱ��", 1800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("������", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ִ�п���id", 0, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ȡѪ��", 800, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("�ܾ�ԭ��", 2000, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("��ǰ����ID", 0, flexAlignRightCenter, flexDTString, "", "", True, , , True)
                    
                    .AppendRows = False
                    .SysHidden(.ColIndex("�շ�id")) = True
                    .SysHidden(.ColIndex("����id")) = True
                    .SysHidden(.ColIndex("��ҳid")) = True
                    
                    Call .InitializeEdit(True, True, True)
                    Call .InitializeEditColumn(.ColIndex(""), True, vbVsfEditCheck)
                    Call .InitializeEditColumn(.ColIndex("�ܾ�ԭ��"), True, vbVsfEditText)
                End With
                
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "��������"
                blnSelect = False
                With VSF1
                    '��ѯ����ѡ�е�����
                    For lngi = 1 To .Rows - 1
                        If Abs(Val(.TextMatrix(lngi, 3))) = 1 Then
                            blnSelect = True
                            Exit For
                        End If
                    Next
                    '��ѡ����������ʾ
                    If blnSelect = False Then
                        MsgBox "��ѡ��Ҫ���յ�ѪҺ���ݣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                    '�˶�����
                    If frmBloodVerification.ShowCheck(Me, True) = False Then ExecuteCommand = False: Exit Function
                    strOpter = frmBloodVerification.str������
                    '����ѡ������
                    For lngi = 1 To .Rows - 1
                        If Abs(Val(.TextMatrix(lngi, 3))) = 1 Then  '����ѡ��״̬
                            If marrFilter(mtbcThisIndex).DeptID = -1 Then
                                lngDeptID = Val(.TextMatrix(lngi, .ColIndex("��ǰ����id")))
                            Else
                                lngDeptID = marrFilter(mtbcThisIndex).DeptID
                            End If
                            If opt1(0).Value = True Then
                                StrSqlSAD = "Zl_ѪҺ���յǼ�_Receive(" & Val(.TextMatrix(lngi, .ColIndex("�շ�id"))) & ",'" & UserInfo.���� & "'," & IIf(.TextMatrix(lngi, .ColIndex("����״̬")) = "���ڽ���", 1, 2) & ",'" & .TextMatrix(lngi, .ColIndex("�ܾ�ԭ��")) & "',null,'" & strOpter & "'," & lngDeptID & ")"  '������Ϊ��½��
                            Else
                                StrSqlSAD = "Zl_ѪҺ���յǼ�_Receive(" & Val(.TextMatrix(lngi, .ColIndex("�շ�id"))) & ",'" & UserInfo.���� & "'," & IIf(.TextMatrix(lngi, .ColIndex("����״̬")) = "���ڽ���", 1, 2) & ",'" & .TextMatrix(lngi, .ColIndex("�ܾ�ԭ��")) & "',To_Date('" & DTP3.Value & "','YYYY-MM-DD hh24:mi'),'" & strOpter & "'," & lngDeptID & ")"   '������Ϊ��½��
                            End If
                            Call SQLRecordAdd(rsSAD, StrSqlSAD)
                        End If
                    Next
                    Call SQLRecordExecute(rsSAD)
                End With
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "����"
                With VSF2
                    For lngi = 1 To .Rows - 1
                        If Abs(Val(.TextMatrix(lngi, 3))) = 1 Then
                            blnSelect = True
                            Exit For
                        End If
                    Next
                    If blnSelect = False Then
                        MsgBox "��ѡ��Ҫ���˵�ѪҺ���ݣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                End With
                For lngi = 1 To VSF2.Rows - 1
                    '����ʱҪ���ǵ�ǰ��½�û��ͽ������Ƿ���ͬ�������ͬ��������ˣ�˭���գ�˭�ſ��Ի���
                    If Abs(Val(VSF2.TextMatrix(lngi, 3))) = 1 Then '����ѡ��״̬
                        
                        StrSqlSAD = "Zl_ѪҺ���յǼ�_fallback(" & Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("�շ�id"))) & ")" 'fallback,Unreceive
                        Call SQLRecordAdd(rsSAD, StrSqlSAD)
                    End If
                Next
                Call SQLRecordExecute(rsSAD)
            Case "ת��"
                blnSelect = False
                With VSF2
                    For lngi = 1 To .Rows - 1
                        If Abs(Val(.TextMatrix(lngi, 3))) = 1 Then
                            blnSelect = True
                            Exit For
                        End If
                    Next
                    If blnSelect = False Then
                        MsgBox "��ѡ��Ҫת�ӵ�ѪҺ���ݣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If frmBloodVerification.ShowCheck(Me, True) = False Then ExecuteCommand = False: Exit Function
                    strOpter = frmBloodVerification.str������
                    If optTransit(0).Value = True Then
                        strCurDate = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:MM")
                    Else
                        strCurDate = Format(dtpTransit.Value, "YYYY-MM-DD HH:MM")
                    End If
                    strRows = ""
                    For lngi = 1 To .Rows - 1
                        If Abs(Val(.TextMatrix(lngi, 3))) = 1 Then  '����ѡ��״̬
                            'ת��֮ǰҪ�ж�ת��ʱ���Ƿ���ڽ���ʱ�䣬���С�ڽ���ʱ�䣬��Ҫ��ʾ�Ҳ����Խ���
                            If strCurDate > Format(VSF2.TextMatrix(lngi, VSF2.ColIndex("����ʱ��")), "YYYY-MM-DD HH:mm") Then
                                lngDeptID = Val(.TextMatrix(lngi, .ColIndex("��ǰ����id")))
                                StrSqlSAD = "Zl_ѪҺ���յǼ�_Transfer(" & Val(.TextMatrix(lngi, .ColIndex("�շ�id"))) & ",3,'" & .TextMatrix(lngi, .ColIndex("������")) & "',To_Date('" & strCurDate & "','YYYY-MM-DD hh24:mi'),'" & strOpter & "'," & lngDeptID & ",NULL)"
                                Call SQLRecordAdd(rsSAD, StrSqlSAD)
                            Else
                                strRows = strRows & "," & lngi
                            End If
                        End If
                    Next
                    If strRows <> "" Then
                        If MsgBox("��[" & Mid(strRows, 2) & "]�����ݵĽ���ʱ��[" & strCurDate & "]С���ϴν���ʱ�䣬�������ݱ��ν�����ת�ӡ�" & vbCrLf & "�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
                    End If
                    Call SQLRecordExecute(rsSAD)
                End With
            Case "ˢ������"
                '���ݵ�ǰѡ�еĲ���ˢ��vsf�ϵ�����
                Select Case mtbcThisIndex
                    Case 0 'ѡ�1
                        If chk1.Value = Unchecked Then '��ѡ
                            Call ShowVsf(0, , UCP.strReturn)
                        Else '��ѡ
                            Call ShowVsf(0, UCP.GetCheckedData)
                        End If
                    Case 1 'ѡ�2
                        If chk1.Value = Unchecked Then '��ѡ
                            Call ShowVsf(1, , UCP.strReturn)
                        Else '��ѡ
                            Call ShowVsf(1, UCP.GetCheckedData)
                        End If
                End Select
                chk2.Value = Unchecked
                chk3.Value = Unchecked
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End Select
    Next

    ExecuteCommand = True
    Exit Function
Error:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    ExecuteCommand = False
End Function

Private Sub chk1_Click()
    'ת����ȡģʽ��һ���ǵ���������ȡ��һ���Ƕ�����˵�������ȡ��ͬʱҲҪ���vsf�ϵ�����
    If Me.Visible = True Then
        cmd2.Enabled = chk1.Value
        UCP.CanCheck = chk1.Value
        If mtbcThisIndex = 0 Then
            VSF1.Rows = 2
            VSF1.RowData(1) = 0
            VSF1.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
        Else
            VSF2.Rows = 2
            VSF2.RowData(1) = 0
            VSF2.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
        End If
    End If
End Sub

Private Sub cmd1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub cmd2_Click()
    Call ExecuteCommand("ˢ������")
End Sub

Private Sub cmd2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdOper_Click()
    If mblnHavePrivs = False Then Exit Sub
    Call GetSendorReceivePeople(txtOper)
End Sub

Private Sub dkpPeoPle_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
        Case 1
            Item.Handle = pic1.hWnd
        Case 2
            Item.Handle = pic7.hWnd
    End Select
End Sub

Private Sub dtpEdtime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpSttime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Set gobjFScrollBar = UCP.FScrollBar
    glngBooldPepWinProc = GetWindowLong(UCP.objPicBack.hWnd, GWL_WNDPROC)
    SetWindowLong UCP.objPicBack.hWnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub Form_Deactivate()
    SetWindowLong UCP.objPicBack.hWnd, GWL_WNDPROC, glngBooldPepWinProc
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHand
    marrPreCardID(0) = ""
    marrPreCardID(1) = ""
    optOccasion(0).Enabled = True
    optOccasion(1).Enabled = True
    If mint���� = 1 Or mint���� = 2 Then
        If mint���� = 1 Then
            optOccasion(1).Value = True
            optOccasion(0).Enabled = False
        Else
            optOccasion(0).Value = True
            optOccasion(1).Enabled = False
        End If
    End If
    mblnת�� = False
    Call InitCondFilter
    DTP3.Value = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    dtpTransit.Value = DTP3.Value
    
    Call ExecuteCommand("��ʼ���")
'    Set mfrmBloodPeoPle = New frmBloodPeoPle
    '��ʼ��UCP�ؼ�
    UCP.UserInit Me, "��ɫ|ID|1||||255;סԺ���|��ҳID;����;����;������;�Ա�����;����;ABORH;ͼ��", ils16, pѪҺ���յǼ�
    Call LoadDeptAndCard '��ʼ�����Ż��в�����Ϣ
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadDeptAndCard()
    Dim strSQL As String
    Dim bln���� As Boolean, i As Integer
'    Dim rs���� As New ADODB.Recordset
    On Error GoTo ErrHand
    '������Ϣ
    bln���� = optOccasion(1).Value
    Set mrs���� = GetDeptList("�ٴ�", IIf(bln���� = True, 1, 2), IsPrivs(mstrPrivs, "���п���"))
    If mrs����.RecordCount <= 0 And IsPrivs(mstrPrivs, "���п���") = False Then
        MsgBox "��������" & IIf(bln���� = True, "����", "סԺ") & "���ţ�", vbInformation, gstrSysName
        mblnHavePrivs = False
        Exit Sub
    End If
    cboDepart.Clear
    '���в���
    If InStr(";" & mstrPrivs & ";", ";���п���;") > 0 Then
        cboDepart.AddItem "���в���"
        cboDepart.ItemData(cboDepart.NewIndex) = -1
        cboDepart.Tag = cboDepart.Text
    End If
    
    For i = 1 To mrs����.RecordCount
        cboDepart.AddItem mrs����!���� & "-" & mrs����!����
        cboDepart.ItemData(cboDepart.NewIndex) = mrs����!id
        '����ȱʡ
        If IsPrivs(mstrPrivs, "���п���") = False Then
            If mrs����!ȱʡ = 1 Then
                Call gobjComlib.cbo.SetIndex(cboDepart.hWnd, cboDepart.NewIndex)
            End If
        End If
        mrs����.MoveNext
    Next
    If mrs����.RecordCount > 0 Then
        mrs����.MoveFirst
    End If
    
    If cboDepart.ListIndex = -1 And cboDepart.ListCount > 0 Then
        Call gobjComlib.cbo.SetIndex(cboDepart.hWnd, 0)
    End If
    cboDepart.Tag = cboDepart.Text
    '����usrCardPeople�ؼ�
    strSQL = "Select '' ��ɫ,'' ID,'' סԺ���, '' ��ҳID,'' ����,'' ����,'' ������,'' �Ա�����,'' ����,'' ABORH From dual where 1<>1"
    Set mRsBR(0) = gobjDatabase.OpenSQLRecord(strSQL, "ȡѪ�˻��������Ϣ")
    Set mRsBR(1) = gobjComlib.Rec.CopyNew(mRsBR(0))
    UCP.ShowPeople mRsBR(0)         '��usrCardPeoPle�������
    mblnHavePrivs = True
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitCondFilter()
    '���ܣ���ʼ��cbotime��dtpSttime��dtpEdtime�ؼ�
    Dim intDay As Long
    Dim intStart As Long
    
    mintOutPreTime = -1
    cboTime.Clear
    With cboTime
        .AddItem "����"
        .ItemData(.NewIndex) = 0
        .AddItem "2����"
        .ItemData(.NewIndex) = 1
        .AddItem "3����"
        .ItemData(.NewIndex) = 2
        .AddItem "һ����"
        .ItemData(.NewIndex) = 6
        .AddItem "[ָ��...]"
        .ItemData(.NewIndex) = -1
    End With
    
    intStart = Val(gobjDatabase.GetPara("����ȡѪʱ��ȱʡ��Χ", 100, pѪҺ���յǼ�, "0"))
    '�Զ���Ĭ�϶�λ������
    If InStr(1, ",0,1,2,6,", "," & intStart & ",") = 0 Then
        mstr��ʼʱ�� = GetDateTime(0, 1)
        mstr����ʱ�� = GetDateTime(0, 2)
        Call gobjComlib.cbo.SetIndex(cboTime.hWnd, 0)
    Else
        mstr��ʼʱ�� = GetDateTime(intStart, 1)
        mstr����ʱ�� = GetDateTime(intStart, 2)
        Select Case intStart
            Case 0
                Call gobjComlib.cbo.SetIndex(cboTime.hWnd, 0)
            Case 1
                Call gobjComlib.cbo.SetIndex(cboTime.hWnd, 1)
            Case 2
                Call gobjComlib.cbo.SetIndex(cboTime.hWnd, 2)
            Case 6
                Call gobjComlib.cbo.SetIndex(cboTime.hWnd, 3)
            Case Else
                Call gobjComlib.cbo.SetIndex(cboTime.hWnd, 4)
        End Select
    End If
    mintOutPreTime = cboTime.ListIndex
    dtpSttime.Value = Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm")
    dtpEdtime.Value = Format(mstr����ʱ��, "YYYY-MM-DD HH:mm")
    dtpSttime.Enabled = (cboTime.ItemData(cboTime.ListIndex) = -1): dtpEdtime.Enabled = dtpSttime.Enabled
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call SetPaneRange(dkpPeoPle, 1, 260, 100, 320, Me.ScaleHeight)
    Call SetPaneRange(dkpPeoPle, 2, 100, 100, Me.ScaleWidth, Me.ScaleHeight)
    dkpPeoPle.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call gobjDatabase.SetPara("����ȡѪʱ��ȱʡ��Χ", DateDiff("d", CDate(mstr��ʼʱ��), CDate(mstr����ʱ��)), 100, pѪҺ���յǼ�)
    
    mblnBloodReceivesIsOpen = False
    Set mRsBR(0) = Nothing
    Set mRsBR(1) = Nothing
    Set mclsVsf = Nothing
    Set mclsvsf1 = Nothing
    Set mrs���� = Nothing
'    Set marrFilter(0) = Nothing
'    Set marrFilter(1) = Nothing
End Sub

Private Sub UCP_CardChanged()
    Dim strReturn As String
    strReturn = UCP.strReturn
    If strReturn = "" Then Exit Sub
    If chk1.Value = Unchecked Then  '��ѡ��������ȡ����Ƭ�л�����������ˢ��
        marrPreCardID(mtbcThisIndex) = Split(strReturn, "'")(1)
        Call ExecuteCommand("ˢ������")
    End If
End Sub

Private Sub opt1_Click(Index As Integer)
    If Me.Visible = True Then
        DTP3.Enabled = opt1(1).Value
    End If
End Sub

Private Sub optOccasion_Click(Index As Integer)
    If Me.Visible = True Then
'        If mtbcThisIndex = 0 Then
'            VSF1.Rows = 2
'            VSF1.RowData(1) = 0
'            VSF1.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
'        Else
'            VSF2.Rows = 2
'            VSF2.RowData(1) = 0
'            VSF2.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
'        End If
        '�ڻ�סԺ/�������ģʽʱ������ҳ�涼Ҫ���������ת��tbcthis�ؼ���ҳ�����ת������ģʽ���ᵼ��ҳ�����ݲ���
        VSF1.Rows = 2
        VSF1.RowData(1) = 0
        VSF1.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
        VSF2.Rows = 2
        VSF2.RowData(1) = 0
        VSF2.Cell(flexcpText, 1, 0, 1, VSF1.Cols - 1) = ""
        Call LoadDeptAndCard
        mblnת�� = False
    End If
End Sub

Private Sub optTransit_Click(Index As Integer)
    If Me.Visible = True Then
        dtpTransit.Enabled = optTransit(1).Value
        If dtpTransit.Enabled = True Then
            dtpTransit.Value = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm")
        End If
    End If
End Sub

Private Sub pic1_Resize()
    On Error Resume Next
    lbl3.Left = pic1.ScaleLeft + 50
    lbl3.Top = pic1.ScaleTop
    lbl3.Width = pic1.ScaleWidth - 100
    lbl3.Height = 260
    
    Fra1.Left = lbl3.Left
    Fra1.Top = pic1.ScaleTop + lbl3.Height
    Fra1.Width = pic1.ScaleWidth - 100
    Fra1.Height = 2895
    
    UCP.Left = lbl3.Left
    UCP.Top = Fra1.Top + Fra1.Height + 100
    UCP.Width = Fra1.Width
    If pic1.ScaleHeight - Fra1.Top - Fra1.Height - 100 > 0 Then
        UCP.Height = pic1.ScaleHeight - Fra1.Top - Fra1.Height - 100
    End If
    
    'Fra1�пؼ�����
    cboDepart.Width = Fra1.Width - cboDepart.Left - 120
    cboTime.Width = cboDepart.Width
    dtpSttime.Width = cboDepart.Width
    dtpEdtime.Width = cboDepart.Width
    txtOper.Width = cboDepart.Width
    cmdOper.Left = txtOper.Left + txtOper.Width - cmdOper.Width - 30
    cmd1.Left = cboDepart.Width + cboDepart.Left - cmd1.Width
    cmd2.Left = cmd1.Left
End Sub

Private Sub pic11_Resize()
    On Error Resume Next
    If picTransit.Visible = False Then
        pic4.Move pic11.ScaleLeft, pic11.ScaleTop, pic11.ScaleWidth, pic11.ScaleHeight
    Else
        picTransit.Move pic11.ScaleLeft, pic11.ScaleTop, pic11.ScaleWidth, 500
        pic4.Move pic11.ScaleLeft, picTransit.Top + picTransit.Height, pic11.ScaleWidth, pic11.ScaleHeight - picTransit.Top - picTransit.Height
    End If
End Sub

Private Sub pic3_Resize()
    On Error Resume Next
    VSF1.Move pic3.ScaleLeft + 50, pic3.ScaleTop + 50, pic3.ScaleWidth - 100, pic3.ScaleHeight - 100
    chk3.Move VSF1.Left + 10 + VSF1.ColWidth(3) / 2 - 100, VSF1.Top + 10, VSF1.ColWidth(3) / 2 + 90
End Sub

Private Sub Pic4_Resize()
    On Error Resume Next
    VSF2.Move pic4.ScaleLeft + 50, pic4.ScaleTop + 50, pic4.ScaleWidth - 100, pic4.ScaleHeight - 100
    chk2.Move VSF2.Left + 10 + VSF2.ColWidth(3) / 2 - 100, VSF2.Top + 10, VSF2.ColWidth(3) / 2 + 90
End Sub

Private Sub Pic5_Resize()
    On Error Resume Next
    pic2.Move Pic5.ScaleLeft, Pic5.ScaleTop, Pic5.ScaleWidth, 500
    pic3.Move Pic5.ScaleLeft, pic2.Top + pic2.Height, Pic5.ScaleWidth, Pic5.ScaleHeight - pic2.Top - pic2.Height
End Sub

Private Sub pic7_Resize()
    On Error Resume Next
    tbcthis.Move pic7.ScaleLeft, pic7.ScaleTop, pic7.ScaleWidth, pic7.ScaleHeight
End Sub

Private Sub tbcThis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '��ת��ѡ�ʱˢ��ҳ��
    Dim i As Integer
    If Item.Tag = "" Then Exit Sub
    mtbcThisIndex = Item.Index
    If mtbcThisIndex = 0 Then
        lbl2.Caption = "ȡѪʱ��"
        lbl5.Caption = "ȡ Ѫ ��"
    Else
        lbl2.Caption = "����ʱ��"
        lbl5.Caption = "�� �� ��"
    End If
    For i = 0 To cboDepart.ListCount - 1
        If cboDepart.ItemData(i) = marrFilter(mtbcThisIndex).DeptID Then
            Call gobjComlib.cbo.SetIndex(cboDepart.hWnd, i)
            Exit For
        End If
    Next
    cboTime.ListIndex = marrFilter(mtbcThisIndex).TimeIndex
    If cboTime.ListIndex = cboTime.ListCount - 1 Then
        dtpSttime.Value = Format(marrFilter(mtbcThisIndex).BeginTime, "YYYY-MM-DD HH:mm")
        dtpEdtime.Value = Format(marrFilter(mtbcThisIndex).EndTime, "YYYY-MM-DD HH:mm")
    End If
    txtOper.Text = marrFilter(mtbcThisIndex).Oper
    cmd1_Click
    mblnת�� = False '��תҳ��������ת��״̬
    UCP.FindStart = True '��תҳ����ʼ����ѯ
    pic11_Resize
End Sub

Private Sub txtOper_KeyPress(KeyAscii As Integer)
'    If KeyAscii <> vbKeyReturn Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        If txtOper.Tag <> txtOper.Text And Trim(txtOper.Text) <> "" Then
            Call GetSendorReceivePeople(txtOper, txtOper.Text)
        End If
        txtOper.Tag = txtOper.Text
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub VSF1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Trim(VSF1.TextMatrix(Row, VSF1.ColIndex("�ܾ�ԭ��"))) <> "" And VSF1.TextMatrix(Row, VSF1.ColIndex("����״̬")) = "���ڽ���" Then
        VSF1.TextMatrix(Row, VSF1.ColIndex("����״̬")) = "�ܾ�����"
'    Else
'        VSF1.TextMatrix(Row, VSF1.ColIndex("����״̬")) = "���ڽ���"
    End If
End Sub

Private Sub VSF1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub VSF1_AfterScroll(ByVal OldtopRow As Long, ByVal OldLeftCol As Long, ByVal NewtopRow As Long, ByVal NewLeftCol As Long)
    If NewLeftCol > 3 Then
        chk3.Visible = False
    Else
        chk3.Visible = True
    End If
End Sub

Private Sub VSF1_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    chk3.Move VSF1.Left + 10 + VSF1.ColWidth(3) / 2 - 100, VSF1.Top + 10, VSF1.ColWidth(3) / 2 + 90
End Sub

Private Sub VSF1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0 '�����������"'"
End Sub

Private Sub VSF2_AfterScroll(ByVal OldtopRow As Long, ByVal OldLeftCol As Long, ByVal NewtopRow As Long, ByVal NewLeftCol As Long)
    If NewLeftCol > 3 Then
        chk2.Visible = False
    Else
        chk2.Visible = True
    End If
End Sub

Private Sub VSF2_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    chk2.Move VSF2.Left + 10 + VSF2.ColWidth(3) / 2 - 100, VSF2.Top + 10, VSF2.ColWidth(3) / 2 + 90
End Sub

Private Sub VSF1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSF1.TextMatrix(Row, VSF1.ColIndex("����id"))) = 0 Then Cancel = True: Exit Sub
    '����ĳЩ�в��ܱ༭
    If Col = VSF1.ColIndex("ѡ��") Then
        Cancel = False
    Else
        Cancel = True
    End If
    If Abs(Val(VSF1.TextMatrix(Row, VSF1.ColIndex("ѡ��")))) = 1 Then
        If Col = VSF1.ColIndex("����״̬") Or Col = VSF1.ColIndex("�ܾ�ԭ��") Then
            Cancel = False
        End If
'        If Col = VSF1.ColIndex("�ܾ�ԭ��") And VSF1.TextMatrix(Row, VSF1.ColIndex("����״̬")) = "�ܾ�����" Then
'            Cancel = False
'        End If
    End If
End Sub

Private Sub VSF1_DblClick()
    Call mclsVsf.DbClick
End Sub

Private Sub VSF2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsvsf1.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub VSF2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(VSF2.TextMatrix(Row, VSF2.ColIndex("����id"))) = 0 Then Cancel = True: Exit Sub
    If Col = VSF2.ColIndex("ѡ��") Then
        If mblnת�� = False Then
'            If (UserInfo.���� = VSF2.TextMatrix(Row, VSF2.ColIndex("������")) Or UserInfo.���� = VSF2.TextMatrix(Row, VSF2.ColIndex("������"))) Then
'                Cancel = False
'            Else
'                Cancel = True
'            End If
        Else
            If Val(VSF2.TextMatrix(Row, VSF2.ColIndex("״̬"))) = 1 Or Val(VSF2.TextMatrix(Row, VSF2.ColIndex("״̬"))) = 3 Then
                Cancel = True
            Else
                Cancel = False
            End If
        End If
    Else
        Cancel = True
    End If
End Sub

Private Function IsClick() As Boolean
    '���ܣ��ж���ת��ģʽ���Ƿ���ѪҺ�ǼǼ�¼��ѡ��
    Dim lngi As Long
    IsClick = False
    For lngi = 1 To VSF2.Rows - 1
        If Abs(Val(VSF2.TextMatrix(lngi, VSF2.ColIndex("ѡ��")))) = 1 Then
            IsClick = True
            Exit For
        End If
    Next
End Function

Private Sub CopyRecord(ByVal RsProm As ADODB.Recordset, ToRs As ADODB.Recordset)
    '���ܣ�����¼��RsProm�Ľṹ�������ݶ����Ƹ�ToRs
    '������RsProm-Ҫ��ֵ�ļ�¼����ToRs-Ŀ���¼��
    Dim lngi As Long
    Dim lngj As Long
    Call RsTitelCopy(RsProm, ToRs)
    With ToRs
        If RsProm.RecordCount > 0 Then '��ǰû�ж�rsbr���������жϻᱨ��
            For lngi = 0 To RsProm.RecordCount - 1
                .AddNew
                For lngj = 0 To RsProm.Fields.Count - 1
                    .Fields(lngj).Value = RsProm.Fields(lngj).Value
                Next
                .Update
                RsProm.MoveNext
            Next
            RsProm.MoveFirst
            If .RecordCount > 0 Then
                .MoveFirst
            End If
        End If
    End With
End Sub
