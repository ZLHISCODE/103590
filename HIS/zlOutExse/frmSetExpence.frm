VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSetExpence 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   ControlBox      =   0   'False
   Icon            =   "frmSetExpence.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "�豸����(&S)"
      Height          =   350
      Left            =   6485
      TabIndex        =   36
      Top             =   1290
      Width           =   1230
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6630
      TabIndex        =   38
      Top             =   4765
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6615
      TabIndex        =   35
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6615
      TabIndex        =   34
      Top             =   345
      Width           =   1100
   End
   Begin TabDlg.SSTab stab 
      Height          =   5355
      Left            =   45
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   60
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   9446
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   564
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "���ݿ���(&1)"
      TabPicture(0)   =   "frmSetExpence.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chk�ֽ��˿�ȱʡ��ʽ"
      Tab(0).Control(1)=   "chkҽ��������ȱʡ��λ"
      Tab(0).Control(2)=   "fraDoctor"
      Tab(0).Control(3)=   "fra����"
      Tab(0).Control(4)=   "chk�շ�ִ�п���"
      Tab(0).Control(5)=   "txt�շ�ִ�п���"
      Tab(0).Control(6)=   "cmd�շ�ִ�п���"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdPrintSetup(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fra�˷�ȱʡѡ��ʽ"
      Tab(0).Control(9)=   "chkLedDispDetail"
      Tab(0).Control(10)=   "chkLedWelcome"
      Tab(0).Control(11)=   "chkPayKey"
      Tab(0).Control(12)=   "fra���"
      Tab(0).Control(13)=   "cbo�ѱ�"
      Tab(0).Control(14)=   "cbo���㷽ʽ"
      Tab(0).Control(15)=   "fra������ҽ��"
      Tab(0).Control(16)=   "lbl�ѱ�"
      Tab(0).Control(17)=   "lbl���㷽ʽ"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Ʊ�ݿ���(&2)"
      TabPicture(1)   =   "frmSetExpence.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "chkDefaultPrintDays"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdPrintSetup(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdPrintSetup(5)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdPrintSetup(4)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkRegistInvoice"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdPrintSetup(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdPrintSetup(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdPrintSetup(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "fraTitle"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdPrintSetup(7)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtDefaultPrintDays"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "fraLine"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "ҩ������(&3)"
      TabPicture(2)   =   "frmSetExpence.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl���ϲ���"
      Tab(2).Control(1)=   "vsfDrugStore"
      Tab(2).Control(2)=   "cbo����"
      Tab(2).ControlCount=   3
      Begin VB.Frame fraLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   2685
         TabIndex        =   48
         Top             =   4260
         Width           =   285
      End
      Begin VB.TextBox txtDefaultPrintDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   210
         Left            =   2655
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "7"
         Top             =   4050
         Width           =   345
      End
      Begin VB.CheckBox chk�ֽ��˿�ȱʡ��ʽ 
         Caption         =   "�˷�ѡ��""�ֽ�""���㷽ʽʱȱʡ�˿���"
         Height          =   195
         Left            =   -74820
         TabIndex        =   18
         Top             =   3810
         Width           =   3585
      End
      Begin VB.CheckBox chkҽ��������ȱʡ��λ 
         Caption         =   "ҽ��������ȱʡ��λ����ҽ�����㡱��ť"
         Height          =   195
         Left            =   -74820
         TabIndex        =   17
         Top             =   3555
         Width           =   3735
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "�˷�Ʊ�ݴ�ӡ����(&0)"
         Height          =   350
         Index           =   7
         Left            =   4185
         TabIndex        =   47
         Top             =   3960
         Width           =   1950
      End
      Begin VB.Frame fraDoctor 
         Caption         =   "��ʾ������"
         Height          =   630
         Left            =   -74820
         TabIndex        =   7
         Top             =   1560
         Width           =   3945
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "������+������ʾ"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   2085
            TabIndex        =   9
            Top             =   285
            Width           =   1695
         End
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "������+������ʾ"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   255
            TabIndex        =   8
            Top             =   285
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame fra���� 
         Caption         =   "������Դ"
         Height          =   1095
         Left            =   -72240
         TabIndex        =   4
         Top             =   390
         Width           =   1365
         Begin VB.OptionButton opt���� 
            Caption         =   "���ﲡ��"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   165
            TabIndex        =   5
            Top             =   345
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "סԺ����"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   165
            TabIndex        =   6
            ToolTipText     =   "סԺ�����������ʱ,�����־Ϊ1(������ʵȲ���Ҳ������������ʵĹ�����)"
            Top             =   675
            Width           =   1020
         End
      End
      Begin VB.CheckBox chk�շ�ִ�п��� 
         Caption         =   "�����շ�ִ�п���"
         Height          =   210
         Left            =   -74820
         TabIndex        =   25
         Top             =   5040
         Width           =   1770
      End
      Begin VB.TextBox txt�շ�ִ�п��� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   -73050
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   4980
         Width           =   3975
      End
      Begin VB.CommandButton cmd�շ�ִ�п��� 
         Caption         =   "��"
         Height          =   280
         Left            =   -69060
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   4995
         Width           =   280
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "����֪ͨ����ӡ����(&1)"
         Height          =   350
         Index           =   3
         Left            =   -70830
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4350
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame fra�˷�ȱʡѡ��ʽ 
         Caption         =   "�˷�ȱʡѡ��ʽ"
         Height          =   840
         Left            =   -74820
         TabIndex        =   19
         Top             =   4065
         Width           =   3945
         Begin VB.OptionButton opt�˷�ȱʡѡ��ʽ 
            Caption         =   "ȱʡȫѡ���˷���Ŀ"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   555
            Width           =   2010
         End
         Begin VB.OptionButton opt�˷�ȱʡѡ��ʽ 
            Caption         =   "ȱʡ�����ݺŻ�Ʊ��ѡ���˷���Ŀ"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   300
            Value           =   -1  'True
            Width           =   3195
         End
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   -73620
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   4815
         Width           =   2355
      End
      Begin VB.Frame fraTitle 
         Caption         =   "���ع����շ�Ʊ��"
         Height          =   3105
         Left            =   150
         TabIndex        =   45
         Top             =   510
         Width           =   6000
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   2760
            Left            =   75
            TabIndex        =   28
            Top             =   255
            Width           =   5790
            _cx             =   10213
            _cy             =   4868
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
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
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSetExpence.frx":0060
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
            ExplorerBar     =   2
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
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "�շ�Ʊ�ݴ�ӡ����(&1)"
         Height          =   350
         Index           =   0
         Left            =   150
         TabIndex        =   44
         Top             =   4365
         Width           =   1950
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "�վ�֤����ӡ����(&2)"
         Height          =   350
         Index           =   1
         Left            =   2160
         TabIndex        =   43
         Top             =   4365
         Width           =   1950
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "�շ��嵥��ӡ����(&3)"
         Height          =   350
         Index           =   2
         Left            =   4185
         TabIndex        =   42
         Top             =   4365
         Width           =   1950
      End
      Begin VB.CheckBox chkRegistInvoice 
         Caption         =   "�Һ�ʱʹ�����շ���ͬ��Ʊ��"
         Height          =   195
         Left            =   165
         TabIndex        =   29
         Top             =   3780
         Width           =   2640
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "ҽ���ص���ӡ����(&4)"
         Height          =   350
         Index           =   4
         Left            =   150
         TabIndex        =   41
         Top             =   4770
         Width           =   1950
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "�˷ѻص���ӡ����(&5)"
         Height          =   350
         Index           =   5
         Left            =   2160
         TabIndex        =   40
         Top             =   4770
         Width           =   1950
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "ִ���嵥��ӡ����(&6)"
         Height          =   350
         Index           =   6
         Left            =   4185
         TabIndex        =   39
         Top             =   4770
         Width           =   1950
      End
      Begin VB.CheckBox chkLedDispDetail 
         Caption         =   "LED��ʾ�շ���ϸ"
         Height          =   225
         Left            =   -74835
         TabIndex        =   15
         ToolTipText     =   "�շѴ���,�����շ���Ŀ���Ƿ���ʾ��Ϣ"
         Top             =   3300
         Value           =   1  'Checked
         Width           =   1770
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED��ʾ��ӭ��Ϣ"
         Height          =   225
         Left            =   -73020
         TabIndex        =   16
         ToolTipText     =   "�շѴ������벡�˺�,�Ƿ���ʾ��ӭ��Ϣ������"
         Top             =   3300
         Value           =   1  'Checked
         Width           =   1770
      End
      Begin VB.CheckBox chkPayKey 
         Caption         =   "ʹ��С���̵ļӼ�(+-)���л�֧����ʽ"
         Height          =   195
         Left            =   -74835
         TabIndex        =   14
         Top             =   3045
         Width           =   3375
      End
      Begin VB.Frame fra��� 
         Caption         =   "�����շ����"
         Height          =   3885
         Left            =   -70800
         TabIndex        =   22
         Top             =   390
         Width           =   2055
         Begin VB.ListBox lst�շ���� 
            ForeColor       =   &H00C00000&
            Height          =   3420
            Left            =   60
            Style           =   1  'Checkbox
            TabIndex        =   23
            ToolTipText     =   "�븴ѡ����ʹ�õ��շ����"
            Top             =   345
            Width           =   1920
         End
      End
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   300
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2295
         Width           =   2235
      End
      Begin VB.ComboBox cbo���㷽ʽ 
         Height          =   300
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2655
         Width           =   2235
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
         Height          =   4140
         Left            =   -74835
         TabIndex        =   32
         Top             =   555
         Width           =   5970
         _cx             =   10530
         _cy             =   7302
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
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSetExpence.frx":013E
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
      Begin VB.Frame fra������ҽ�� 
         Caption         =   "������ҽ��"
         Height          =   1095
         Left            =   -74835
         TabIndex        =   0
         Top             =   390
         Width           =   2580
         Begin VB.OptionButton optUnit 
            Caption         =   "ͨ�����������ȷ��ҽ��"
            Height          =   180
            Left            =   210
            TabIndex        =   1
            Top             =   300
            Value           =   -1  'True
            Width           =   2280
         End
         Begin VB.OptionButton optDoctor 
            Caption         =   "ͨ������ҽ����ȷ������"
            Height          =   180
            Left            =   210
            TabIndex        =   2
            Top             =   540
            Width           =   2280
         End
         Begin VB.OptionButton optSelf 
            Caption         =   "���Һ�ҽ�������������"
            Height          =   195
            Left            =   210
            TabIndex        =   3
            Top             =   780
            Width           =   2280
         End
      End
      Begin VB.CheckBox chkDefaultPrintDays 
         Caption         =   "�����˲���Ʊ��ʱȱʡ��ӡ     ��ķ���"
         Height          =   195
         Left            =   165
         TabIndex        =   30
         Top             =   4050
         Width           =   3660
      End
      Begin VB.Label lbl���ϲ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡ���ϲ���"
         Height          =   180
         Left            =   -74820
         TabIndex        =   46
         Top             =   4875
         Width           =   1080
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡ���˷ѱ�"
         Height          =   180
         Left            =   -74835
         TabIndex        =   10
         Top             =   2355
         Width           =   1080
      End
      Begin VB.Label lbl���㷽ʽ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡ���㷽ʽ"
         Height          =   180
         Left            =   -74835
         TabIndex        =   12
         Top             =   2715
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSetExpence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mbytInFun As Byte '0=�շ�,1=����,2=�������
Public mstrPrivs As String
Public mlngModul As Long
Public mblnSetDrugStore As Boolean
Private mblnAutoAddItem As Boolean
Private mblnNotClick As Boolean

Private Sub chkDefaultPrintDays_Click()
    txtDefaultPrintDays.Enabled = (chkDefaultPrintDays.Value = vbChecked)
    txtDefaultPrintDays.BackColor = IIf(txtDefaultPrintDays.Enabled, vbWhite, vbButtonFace)
End Sub

Private Sub cmdDeviceSetup_Click()
    Dim lngModule As Long
    Select Case mbytInFun
    Case 0
        lngModule = 1121
    Case 1
        lngModule = 1120
    Case 2
        lngModule = 1122
    End Select
    Call zlCommFun.DeviceSetup(Me, glngSys, lngModule)
End Sub

Private Sub cbo�ѱ�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cbo�ѱ�.ListIndex = -1
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name & "1"
End Sub

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���淢Ʊ���Ʊ��
    '����:���˺�
    '����:2011-04-28 18:16:48
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    
    '���湲��Ʊ��
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.TextMatrix(i, .ColIndex("ʹ�����")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "�����շ�Ʊ������", strValue, glngSys, mlngModul, blnHavePrivs
End Sub
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч�Լ��
    '����:���Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-28 18:24:16
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str��� As String
    
    If mbytInFun <> 0 Then isValied = True: Exit Function
     
    isValied = False
    On Error GoTo errHandle
    '���ÿ��ʹ����ʽֻ��һ��ѡ��
    With vsBill
        str��� = "-"
        For i = 1 To vsBill.Rows - 1
            If str��� <> Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) Then
               str��� = Trim(.TextMatrix(i, .ColIndex("ʹ�����")))
               lngSelCount = 0
                For j = 1 To vsBill.Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) = Trim(.TextMatrix(j, .ColIndex("ʹ�����"))) Then
                        If Val(.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "ע��:" & vbCrLf & "    ʹ�����Ϊ��" & str��� & "����ֻ��ѡ��һ��Ʊ��,����!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    Dim strValue As String, i As Long
    Dim str��ҩ������ As String, str��ҩ������ As String, str��ҩ������ As String
    Dim lngȱʡ��ҩ�� As Long, lngȱʡ��ҩ�� As Long, lngȱʡ��ҩ�� As Long, lngȱʡ���ϲ��� As Long
    
    'a.���ݼ��
    '--------------------------------------------------------------
    'b.����ע���洢��ģ�����
    '------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    If isValied = False Then Exit Sub
     
    'c.���ݿ�洢��ģ�����
    '----------------------------------------------------------------------------------------
    
    If Not mblnSetDrugStore Then
        For i = lst�շ����.ListCount - 1 To 0 Step -1
            If lst�շ����.Selected(i) Then strValue = strValue & "'" & Chr(lst�շ����.ItemData(i)) & "',"
        Next
        
        If strValue <> "" Then strValue = Left(strValue, Len(strValue) - 1)
        zlDatabase.SetPara "�շ����", strValue, glngSys, mlngModul, blnHavePrivs
        
        If mbytInFun <> 2 Then
            zlDatabase.SetPara "ȱʡ�ѱ�", cbo�ѱ�.Text, glngSys, mlngModul, blnHavePrivs
        End If
        
        If mbytInFun = 0 Then
            Call SaveInvoice
            zlDatabase.SetPara "ȱʡ���㷽ʽ", cbo���㷽ʽ.Text, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "�ҺŹ����շ�Ʊ��", chkRegistInvoice.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "LED��ʾ�շ���ϸ", chkLedDispDetail.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "LED��ʾ��ӭ��Ϣ", chkLedWelcome.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "ҽ��������ȱʡ��λ", chkҽ��������ȱʡ��λ.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "�ֽ��˿�ȱʡ��ʽ", chk�ֽ��˿�ȱʡ��ʽ.Value, glngSys, mlngModul, blnHavePrivs
            
            '96357
            zlDatabase.SetPara "�����շ�ִ�п���", txt�շ�ִ�п���.Tag, glngSys, mlngModul, blnHavePrivs
            
            If chkDefaultPrintDays.Value = vbUnchecked Then
                strValue = "0"
            Else
                strValue = Val(txtDefaultPrintDays.Text)
            End If
            zlDatabase.SetPara "ȱʡ��Ʊ��ӡ����", strValue, glngSys, mlngModul, blnHavePrivs
        End If
    End If

    With vsfDrugStore
        For i = 1 To vsfDrugStore.Rows - 1
            If (mbytInFun = 0 Or mbytInFun = 1) And .TextMatrix(i, .ColIndex("����")) <> "�Զ�����" And .TextMatrix(i, .ColIndex("����")) <> "" Then
                Select Case .TextMatrix(i, 0)
                    Case "��ҩ��"
                        str��ҩ������ = str��ҩ������ & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("����"))
                    Case "��ҩ��"
                        str��ҩ������ = str��ҩ������ & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("����"))
                    Case "��ҩ��"
                        str��ҩ������ = str��ҩ������ & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("����"))
                End Select
            End If
            
            If Abs(Val(.TextMatrix(i, .ColIndex("ȱʡ")))) = 1 Then
                Select Case .TextMatrix(i, .ColIndex("���"))
                    Case "��ҩ��"
                        lngȱʡ��ҩ�� = .RowData(i)
                    Case "��ҩ��"
                        lngȱʡ��ҩ�� = .RowData(i)
                    Case "��ҩ��"
                        lngȱʡ��ҩ�� = .RowData(i)
                End Select
            End If
        Next
    End With
    
    If cbo����.ListIndex <> -1 Then
        lngȱʡ���ϲ��� = cbo����.ItemData(cbo����.ListIndex)
    End If
    
    
    If mbytInFun = 0 Or mbytInFun = 1 Then
        str��ҩ������ = Mid(str��ҩ������, 2)
        str��ҩ������ = Mid(str��ҩ������, 2)
        str��ҩ������ = Mid(str��ҩ������, 2)
        zlDatabase.SetPara "��ҩ������", str��ҩ������, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "��ҩ������", str��ҩ������, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "��ҩ������", str��ҩ������, glngSys, mlngModul, blnHavePrivs
    End If
    
    zlDatabase.SetPara "ȱʡ��ҩ��", lngȱʡ��ҩ��, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ȱʡ��ҩ��", lngȱʡ��ҩ��, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ȱʡ��ҩ��", lngȱʡ��ҩ��, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ȱʡ���ϲ���", lngȱʡ���ϲ���, glngSys, mlngModul, blnHavePrivs
    
    If Not mblnSetDrugStore Then
        zlDatabase.SetPara "����ҽ��", IIf(optDoctor.Value, 0, IIf(optUnit.Value, 1, 2)), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "��������ʾ��ʽ", IIf(optDoctorKind(0).Value, 1, 2), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "������Դ", IIf(opt����(0).Value, 1, 2), glngSys, mlngModul, blnHavePrivs
        If mbytInFun = 0 Or mbytInFun = 1 Then
            zlDatabase.SetPara "ʹ�üӼ��л�֧����ʽ", IIf(chkPayKey.Value = 1, "1", "0"), glngSys, mlngModul, blnHavePrivs
        End If
    End If
     '87489
    zlDatabase.SetPara "�˷�ȱʡѡ��ʽ", IIf(opt�˷�ȱʡѡ��ʽ(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
    
    Call InitLocPar(Choose(mbytInFun + 1, 1121, 1120, 1122))     '��Ҫ��Ҫ�ض��浽����ע���Ĳ���,�������ݿ�Ĳ����ڱ���ʱ���ض�
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdPrintSetup_Click(Index As Integer)
    Select Case Index
        Case 0 '����ҽ�Ʒ��շ�
            If gblnBillPrint Then
                Call gobjBillPrint.zlConfigure
            Else
                If glngSys Like "8??" Then
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_1", Me)
                Else
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_1", Me)
                End If
            End If
        Case 1 '�������֤��
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_2", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_2", Me)
            End If
        Case 2 '�����շ��嵥
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_3", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me)
            End If
        Case 3 '����֪ͨ��
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1120", Me)
        Case 4 'ҽ���ص�����
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me)
        Case 5  '�˷ѻص�����
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me)
        '62982:���ϴ�,2015/5/19,�շ�ִ�е�
        Case 6  '�շ�ִ�е�����
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_6", Me)
        Case 7  '�˷�ִ�е�����
            If gblnBillPrint Then
                Call gobjBillPrint.zlConfigure
            Else
                If glngSys Like "8??" Then
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_7", Me)
                Else
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_7", Me)
                End If
            End If
    End Select
End Sub
 
Private Sub SetDrugStore()
    Dim lngType As Long, strTmp As String, arrTmp As Variant
    Dim i As Long, j As Long, lngRow As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    With vsfDrugStore
        strTmp = "'��ҩ��','��ҩ��','��ҩ��','���ϲ���'"
        
        If stab.TabVisible(1) = True Then
            lngType = IIf(opt����(0).Value, 1, 2)
        Else
            lngType = gint������Դ
        End If
        Set rsTmp = GetDepartments(strTmp, lngType & ",3")
        .Rows = 1
        If mbytInFun = 2 Then .ColHidden(3) = True '������ʲ��贰��
        
        If rsTmp.RecordCount > 0 Then
            rsTmp.Filter = "��������<>'���ϲ���'"
            .Rows = rsTmp.RecordCount + 1
            .MergeCells = flexMergeFixedOnly
            .MergeCol(0) = True
            
            strTmp = "'��ҩ��','��ҩ��','��ҩ��'"
            arrTmp = Split(strTmp, ",")
            lngRow = 1
            For j = 0 To UBound(arrTmp)
                rsTmp.Filter = "��������=" & arrTmp(j)
                If rsTmp.RecordCount > 0 Then
                    For i = 1 To rsTmp.RecordCount
                        .TextMatrix(lngRow, 0) = Replace(arrTmp(j), "'", "")
                        .TextMatrix(lngRow, 1) = 0
                        .TextMatrix(lngRow, 2) = rsTmp!����
                        If mbytInFun <> 2 Then .TextMatrix(lngRow, 3) = "�Զ�����"
                        .RowData(lngRow) = Val(rsTmp!ID)
                        lngRow = lngRow + 1
                        rsTmp.MoveNext
                    Next
                    
                    If lngRow < .Rows - 1 Then  '���ָ���
                        .Select lngRow, .FixedCols, lngRow, .COLS - 1
                        .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                    End If
                End If
            Next
            
            cbo����.AddItem "�˹�ѡ��"
            rsTmp.Filter = "��������='���ϲ���'"
            For j = 1 To rsTmp.RecordCount
                cbo����.AddItem rsTmp!����
                cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Next
            cbo����.ListIndex = 0
        End If
    
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

 
Private Sub Loadҩ��ParaValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҩ����ز���ֵ
    '����:���˺�
    '����:2011-12-07 15:05:10
    '����:43775
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, blnParSet As Boolean
    Dim i As Long, k As Long, j As Long, intType As Integer
    Dim arrTmp  As Variant, arrWindow As Variant
    Dim str��ҩ������ As String, str��ҩ������ As String, str��ҩ������ As String
    Dim lngȱʡ��ҩ�� As Long, lngȱʡ��ҩ�� As Long, lngȱʡ��ҩ�� As Long, lngȱʡ���ϲ��� As Long
    blnParSet = InStr(1, mstrPrivs, ";��������;") > 0
    
    With vsfDrugStore
        arrTmp = Split("ȱʡ��ҩ��,ȱʡ��ҩ��,ȱʡ��ҩ��", ",")
        .Cell(flexcpData, 0, 0, .Rows - 1, .COLS - 1) = "0" '�洢�Ƿ��������.:0-������,1-����
        
        For j = 0 To UBound(arrTmp)
            '���˺�:���ڿ��ܲ���Ȩ�޷������,���,����ͳһ��������,��Ҫ����ĳһ����:
            '����:25132,intType-'���ز������ͣ�1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
            strTmp = zlDatabase.GetPara(CStr(arrTmp(j)), glngSys, mlngModul, "0", , blnParSet, intType)
            If Val(strTmp) > 0 Then
                Select Case arrTmp(j)
                    Case "ȱʡ��ҩ��"
                        lngȱʡ��ҩ�� = Val(strTmp)
                    Case "ȱʡ��ҩ��"
                        lngȱʡ��ҩ�� = Val(strTmp)
                    Case "ȱʡ��ҩ��"
                        lngȱʡ��ҩ�� = Val(strTmp)
                End Select
                Call SetDrugStockEdit(Replace(arrTmp(j), "ȱʡ", ""), intType, .ColIndex("ȱʡ"), Val(strTmp))
            Else
                Call SetDrugStockEdit(Replace(arrTmp(j), "ȱʡ", ""), intType, .ColIndex("ȱʡ"), "")
            End If
        Next
        
        strTmp = zlDatabase.GetPara("ȱʡ���ϲ���", glngSys, mlngModul, "0", Array(cbo����), blnParSet)
        zlControl.CboLocate cbo����, strTmp, True
        
        If mbytInFun <> 2 Then
                arrTmp = Split("��ҩ������,��ҩ������,��ҩ������", ",")
                For j = 0 To UBound(arrTmp)
                    strTmp = Trim(zlDatabase.GetPara(CStr(arrTmp(j)), glngSys, mlngModul, , , blnParSet, intType))
                    If strTmp <> "" Then
                        '����ɵ�����,���ڲ�����û�д洢ҩ��ID
                        If InStr(strTmp, ":") = 0 Then
                            Select Case arrTmp(j)
                                Case "��ҩ������"
                                    strTmp = lngȱʡ��ҩ�� & ":" & strTmp
                                Case "��ҩ������"
                                    strTmp = lngȱʡ��ҩ�� & ":" & strTmp
                                Case "��ҩ������"
                                    strTmp = lngȱʡ��ҩ�� & ":" & strTmp
                            End Select
                        End If
                        arrWindow = Split(strTmp, ",")
                        strTmp = Replace(arrTmp(j), "����", "")
                        For k = 0 To UBound(arrWindow)
                            Call SetDrugStockEdit(Replace(arrTmp(j), "����", ""), intType, .ColIndex("����"), Val(Split(arrWindow(k), ":")(0)), CStr(Split(arrWindow(k), ":")(1)))
                        Next
                    Else
                        Call SetDrugStockEdit(Replace(arrTmp(j), "����", ""), intType, .ColIndex("����"), "")
                    End If
                Next
            End If
        End With
End Sub
Private Sub LoadParaValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز���ֵ
    '����:���˺�
    '����:2011-09-12 15:03:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, blnParSet As Boolean, k As Long, rsTmp As ADODB.Recordset
    Dim i As Long, arrTmp As Variant, j As Long, intType As Integer, arrWindow As Variant
        
    
    blnParSet = InStr(1, mstrPrivs, ";��������;") > 0

    strTmp = zlDatabase.GetPara("�շ����", glngSys, mlngModul, , Array(lst�շ����), blnParSet)
    If strTmp = "" Then
        For i = 0 To lst�շ����.ListCount - 1
            lst�շ����.Selected(i) = True
        Next
    Else
        For i = 0 To lst�շ����.ListCount - 1
            If InStr(strTmp, Chr(lst�շ����.ItemData(i))) Then lst�շ����.Selected(i) = True
        Next
    End If
    If lst�շ����.ListCount > 0 Then lst�շ����.TopIndex = 0: lst�շ����.ListIndex = 0
    If mbytInFun <> 2 Then
        strTmp = zlDatabase.GetPara("ȱʡ�ѱ�", glngSys, mlngModul, , Array(cbo�ѱ�), blnParSet)
        zlControl.CboLocate cbo�ѱ�, strTmp
    End If
    
    i = IIf(zlDatabase.GetPara("������Դ", glngSys, mlngModul, , Array(opt����(0), opt����(1)), blnParSet) = "1", 0, 1)
    opt����(i).Value = True
    If mbytInFun <> 2 Then opt����(1).ToolTipText = ""
    
    Call opt����_Click(IIf(opt����(0).Value, 0, 1)) '����ҩƷ�ⷿ�����ķ��ϲ���
    Call Loadҩ��ParaValue
    Select Case mbytInFun
    Case 2 '����
    Case 1 '����
    Case Else
        chkRegistInvoice.Value = IIf(zlDatabase.GetPara("�ҺŹ����շ�Ʊ��", glngSys, mlngModul, 0, Array(chkRegistInvoice), blnParSet) = "1", 1, 0)
        chkLedDispDetail.Value = IIf(zlDatabase.GetPara("LED��ʾ�շ���ϸ", glngSys, mlngModul, 1, Array(chkLedDispDetail), blnParSet) = "1", 1, 0)
        chkLedWelcome.Value = IIf(zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", glngSys, mlngModul, 1, Array(chkLedWelcome), blnParSet) = "1", 1, 0)
        
        Dim objCards As Cards '���õ������˻���ҽ�ƿ�
        Set rsTmp = Get���㷽ʽ("�շ�", "1,2,7,8")
        If Not gobjSquare Is Nothing Then
            ' zlGetCards(ByVal BytType As Byte)
                '���:bytType-  0-����ҽ�ƿ�;
            '                        1-���õ�ҽ�ƿ�,
            '                        2-���д��������˻���������
            '                        3-���õ������˻���ҽ�ƿ�
           Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
        End If
        With cbo���㷽ʽ
            .Clear
            Do While Not rsTmp.EOF
                If Not (Val(NVL(rsTmp!����)) = 7 Or Val(NVL(rsTmp!����)) = 8 Or Val(NVL(rsTmp!Ӧ����)) = 1) Then
                    .AddItem NVL(rsTmp!����)
                End If
                rsTmp.MoveNext
            Loop
            '����ҽ�ƿ����㷽ʽ����Ӧ���㷽ʽδ���õĲ�����
            For i = 1 To objCards.Count
            rsTmp.Filter = "����='" & objCards(i).���㷽ʽ & "'"
                If Not rsTmp.EOF Then
                    .AddItem objCards(i).����
                End If
            Next
        End With
        '����:54923
        strTmp = zlDatabase.GetPara("ȱʡ���㷽ʽ", glngSys, mlngModul, , Array(cbo���㷽ʽ), blnParSet)
        For i = 0 To cbo���㷽ʽ.ListCount - 1
            If cbo���㷽ʽ.List(i) = strTmp Then cbo���㷽ʽ.ListIndex = i: Exit For
        Next
        
        '���ط�Ʊ���
        Call InitShareInvoice
        chkPayKey.Value = IIf(Val(zlDatabase.GetPara("ʹ�üӼ��л�֧����ʽ", glngSys, mlngModul, "1", Array(chkPayKey), blnParSet)) = 1, 1, 0)
        '87489
        strTmp = zlDatabase.GetPara("�˷�ȱʡѡ��ʽ", glngSys, mlngModul, "0", Array(opt�˷�ȱʡѡ��ʽ(0), opt�˷�ȱʡѡ��ʽ(1)), blnParSet)
        For i = 0 To 1
            If Val(strTmp) = i Then opt�˷�ȱʡѡ��ʽ(i).Value = True: Exit For
        Next
        chkҽ��������ȱʡ��λ.Value = IIf(zlDatabase.GetPara("ҽ��������ȱʡ��λ", glngSys, mlngModul, "0", Array(chkҽ��������ȱʡ��λ), blnParSet) = "1", 1, 0)
        chk�ֽ��˿�ȱʡ��ʽ.Value = IIf(zlDatabase.GetPara("�ֽ��˿�ȱʡ��ʽ", glngSys, mlngModul, "0", Array(chk�ֽ��˿�ȱʡ��ʽ), blnParSet) = "1", 1, 0)
        
        '96357
        strTmp = zlDatabase.GetPara("�����շ�ִ�п���", glngSys, mlngModul, , Array(chk�շ�ִ�п���, txt�շ�ִ�п���, cmd�շ�ִ�п���), blnParSet)
        mblnNotClick = True
        chk�շ�ִ�п���.Value = IIf(strTmp <> "", vbChecked, vbUnchecked)
        mblnNotClick = False
        cmd�շ�ִ�п���.Enabled = chk�շ�ִ�п���.Value = vbChecked
        txt�շ�ִ�п���.Text = GetDeptNameStr(strTmp)
        txt�շ�ִ�п���.Tag = strTmp
        
        strTmp = zlDatabase.GetPara("ȱʡ��Ʊ��ӡ����", glngSys, mlngModul, "0", Array(chkDefaultPrintDays, txtDefaultPrintDays), blnParSet)
        If Val(strTmp) <= 0 Then
            chkDefaultPrintDays.Value = vbUnchecked
            txtDefaultPrintDays.Text = 7
        Else
            chkDefaultPrintDays.Value = vbChecked
            txtDefaultPrintDays.Text = strTmp
        End If
        txtDefaultPrintDays.Enabled = (chkDefaultPrintDays.Value = vbChecked)
        txtDefaultPrintDays.BackColor = IIf(txtDefaultPrintDays.Enabled, vbWhite, vbButtonFace)
    End Select
End Sub
 
 

Private Sub InitShareInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���Ʊ
    '����:���˺�
    '����:2011-04-28 15:09:10
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '����Ʊ������,��ʽ:����,����
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    Dim lngTemp As Long, i As Long, strSQL As String
    Dim strPrintMode As String, blnHavePrivs As Boolean
    
    On Error GoTo errHandle
    
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '�ָ��п��
    zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "����Ʊ��������", False, False
    strShareInvoice = zlDatabase.GetPara("�����շ�Ʊ������", glngSys, mlngModul, , , True, intType)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
        fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsBill.ForeColor = &H80000008: vsBill.ForeColorFixed = &H80000008
        fraTitle.ForeColor = &H80000008
    End Select
    With vsBill
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then .Editable = flexEDNone
    End With
    
    '��ʽ:����ID1,ʹ�����1|����IDn,ʹ�����n|...
    varData = Split(strShareInvoice, "|")
    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(1)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = NVL(rsTemp!ʹ�����, " ")
            .TextMatrix(lngRow, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
            .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(NVL(rsTemp!ʣ������)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Trim(.TextMatrix(lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
 
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset, strSQL As String, objItem As ListItem, blnParSet As Boolean
    Dim strTmp As String, i As Integer, j As Long, k As Long, arrTmp As Variant, arrWindow As Variant, intType As Integer, blnSeted As Boolean '��������ȱʡֵ

    Dim str��ҩ������ As String, str��ҩ������ As String, str��ҩ������ As String
    Dim lngȱʡ��ҩ�� As Long, lngȱʡ��ҩ�� As Long, lngȱʡ��ҩ�� As Long, lngȱʡ���ϲ��� As Long
    
    gblnOK = False
    On Error GoTo errH
    If mbytInFun = 0 Then
         mblnAutoAddItem = InStr(zlDatabase.GetPara("�Զ����չҺŷ�", glngSys, mlngModul), ";") > 0
    End If
    
    blnParSet = InStr(1, mstrPrivs, "��������") > 0
    
    'a.��ʼ����
    '----------------------------------------------------------------------------------------
    '�շ����(�Һų���):���������
    strSQL = "Select ����,���� as ��� from �շ���Ŀ��� Where ����<>'1' Order by ���"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        lst�շ����.AddItem rsTmp!���
        lst�շ����.ItemData(lst�շ����.NewIndex) = Asc(rsTmp!����)
    
        rsTmp.MoveNext
    Loop
    If mbytInFun <> 2 Then
        strSQL = _
            " Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �ѱ�" & _
            " Where ����=1 And Nvl(���޳���,0)=0 And Nvl(�������,3) IN(1,3)" & _
            "       And Sysdate Between Nvl(��Ч��ʼ,To_Date('1900-01-01','yyyy-mm-dd')) And Nvl(��Ч����,To_Date('3000-01-01','yyyy-mm-dd'))+1-1/24/60/60" & _
            " Order by ����"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        For i = 1 To rsTmp.RecordCount
            cbo�ѱ�.AddItem rsTmp!����
            If rsTmp!ȱʡ = 1 Then cbo�ѱ�.ListIndex = cbo�ѱ�.NewIndex
            rsTmp.MoveNext
        Next
    End If
     
    'b.����ע���洢��ģ�����
    '----------------------------------------------------------------------------------------
    Call LoadParaValue
    'c.���ݿ�洢��ģ�����
    '----------------------------------------------------------------------------------------
    '--------------------------
    strTmp = zlDatabase.GetPara("����ҽ��", glngSys, mlngModul, , Array(optUnit, optDoctor, optSelf), blnParSet)
    If strTmp = "1" Then
        optUnit.Value = True
    ElseIf strTmp = "0" Then
        optDoctor.Value = True
    Else
        optSelf.Value = True
    End If
    
    i = IIf(zlDatabase.GetPara("��������ʾ��ʽ", glngSys, mlngModul, "1", Array(optDoctorKind(0), optDoctorKind(1)), blnParSet) = "1", 0, 1)
    optDoctorKind(i).Value = True
    
    
    'd.Ȩ�޿���
    '----------------------------------------------------------------------------------------
    chkLedDispDetail.Visible = mbytInFun = 0
    chkLedWelcome.Visible = mbytInFun = 0
    chkPayKey.Visible = mbytInFun = 0
    '87489
    fra�˷�ȱʡѡ��ʽ.Visible = mbytInFun = 0
    chkҽ��������ȱʡ��λ.Visible = mbytInFun = 0
    chk�ֽ��˿�ȱʡ��ʽ.Visible = mbytInFun = 0
    
    txt�շ�ִ�п���.Visible = mbytInFun = 0
    chk�շ�ִ�п���.Visible = mbytInFun = 0
    cmd�շ�ִ�п���.Visible = mbytInFun = 0

    lbl���㷽ʽ.Visible = mbytInFun = 0
    cbo���㷽ʽ.Visible = mbytInFun = 0
    
    cmdPrintSetup(3).Visible = mbytInFun = 1
    lbl�ѱ�.Visible = mbytInFun <> 2
    cbo�ѱ�.Visible = mbytInFun <> 2
    
    stab.TabVisible(1) = mbytInFun = 0
 
    If mblnSetDrugStore Then
        '56963
        stab.TabCaption(2) = "ҩ������"
        stab.TabVisible(0) = False
        stab.TabVisible(1) = False
    Else
        If mbytInFun = 1 Then
            stab.TabCaption(2) = "ҩ������(&2)"
        ElseIf mbytInFun = 2 Then
            stab.TabCaption(2) = "ҩ������(&2)"
        End If
    End If
    If stab.TabVisible(0) Then stab.Tab = 0

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDeptNameStr(ByVal strIDs As String) As String
    '������ID�ַ���װ���������ַ���
    '��Σ�
    '   strIDs ����ID����ʽ��ID1,ID2,ID3,...
    '���أ�
    '   ��������s����ʽ����������1;��������2;��������3;...
    Dim strSQL As String, rsTemp As Recordset
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandler
    If strIDs = "" Then Exit Function
    strSQL = "Select /*+cardinality(b,10) */a.����, a.����" & vbNewLine & _
            " From ���ű� A, Table(f_Str2list([1], ',')) B" & vbNewLine & _
            " Where a.Id = b.Column_Value"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݿ���ID��ȡ��������", strIDs)
    If rsTemp Is Nothing Then Exit Function
    
    Do While Not rsTemp.EOF
        strTemp = strTemp & ";" & NVL(rsTemp!����)
        rsTemp.MoveNext
    Loop
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    
    GetDeptNameStr = strTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    mblnSetDrugStore = False
    If mbytInFun = 0 Then
        zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����Ʊ��������", False, False
    End If
    
End Sub

Private Sub lst�շ����_ItemCheck(Item As Integer)
    If lst�շ����.SelCount = 0 And Not lst�շ����.Selected(Item) Then
        lst�շ����.Selected(Item) = True
    End If
End Sub

Private Sub opt����_Click(Index As Integer)
    
    Call SetDrugStore

End Sub
Private Sub stab_Click(PreviousTab As Integer)
    Select Case stab.Tab
        Case 0
            If optUnit.Enabled And optUnit.Visible And optUnit.Value Then optUnit.SetFocus
            If optSelf.Enabled And optSelf.Visible And optSelf.Value Then optSelf.SetFocus
            If optDoctor.Enabled And optDoctor.Visible And optDoctor.Value Then optDoctor.SetFocus
        Case 1
            If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
        Case 2
            If vsfDrugStore.Visible And vsfDrugStore.Enabled Then vsfDrugStore.SetFocus
    End Select
End Sub
      

Private Sub txtDefaultPrintDays_GotFocus()
    zlControl.TxtSelAll txtDefaultPrintDays
End Sub

Private Sub txtDefaultPrintDays_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtDefaultPrintDays, KeyAscii, m����ʽ
End Sub

Private Sub txtDefaultPrintDays_Validate(Cancel As Boolean)
    If Val(txtDefaultPrintDays.Text) < 1 Then txtDefaultPrintDays.Text = 1
End Sub

Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsBill
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.TextMatrix(Row, .ColIndex("ʹ�����"))) = Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
 
Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����Ʊ��������", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����Ʊ��������", False, False
End Sub

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsBill
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub

 
Private Sub vsfDrugStore_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("ȱʡ")
           Call SetDrugStockDeFault(Row)
        Case Else
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("ȱʡ"), .ColIndex("����")
            Cancel = Val(.Cell(flexcpData, Row, Col)) = 1
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick()
    Dim strTmp As String, i As Long
    With vsfDrugStore
        If Not (.Row > 0 And .Col = 1) Then Exit Sub
        If .Cell(flexcpData, .Row, .ColIndex("ȱʡ")) = 1 Then Exit Sub
        
        .TextMatrix(.Row, .Col) = IIf(Val(.TextMatrix(.Row, .Col)) = 0, 1, 0)
        Call SetDrugStockDeFault(.Row)
    End With
End Sub
Private Sub SetDrugStockDeFault(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҩ����ȱʡֵ
    '���:lngRow-ָ����
    '����:���˺�
    '����:2009-09-02 14:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngȱʡ As Long, strType As String
    With vsfDrugStore
        lngȱʡ = Abs(Val(.TextMatrix(lngRow, .ColIndex("ȱʡ"))))
        If lngȱʡ = 1 Then
            strType = .TextMatrix(lngRow, .ColIndex("���"))
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = strType And i <> lngRow Then
                    .TextMatrix(i, .ColIndex("ȱʡ")) = 0
                End If
            Next
        End If
    End With
End Sub
Private Sub SetDrugStockEdit(ByVal strType As String, ByVal intType As Integer, ByVal lngEditCol As Long, Optional strMachValue As String = "", Optional strDefaultValue As String = "")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҩ���ı༭����
    '���:strType-���
    '     intType-���ز������ͣ�1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    '     lngEditCol-���Ƶı༭��
    '����:
    '����:
    '����:���˺�
    '����:2009-09-02 14:53:10
    '����:25132
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnSetDefault As Boolean '������ȱʡֵ��,�����������ȱʡֵ
    Dim lngEditForColor As Long, blnAllowEdit As Boolean, bytLockEdit As Integer '1-����,0-������
    
    '���˺�:���ڿ��ܲ���Ȩ�޷������,���,����ͳһ��������,��Ҫ����ĳһ����:
    With vsfDrugStore
        blnSetDefault = False: blnAllowEdit = InStr(1, mstrPrivs, ";��������;") > 0
        bytLockEdit = 0
        If InStr(1, ",1,3,15,", "," & intType & ",") > 0 Then
            lngEditForColor = IIf(blnAllowEdit, vbBlue, &H8000000C)  '��Ȩ�޿���
            bytLockEdit = IIf(blnAllowEdit, 0, 1)
        ElseIf intType = 5 Then
            lngEditForColor = vbBlue    '����ģ��,������Ȩ�޿���
        Else
            lngEditForColor = &H80000008    '�����༭
        End If
        
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���")) = strType Then
                If lngEditCol = .ColIndex("ȱʡ") Then
                    '����ҩ��
                    If Val(.RowData(i)) = Val(strMachValue) And strMachValue <> "" And Not blnSetDefault Then
                        .TextMatrix(i, .ColIndex("ȱʡ")) = IIf(Val(strMachValue) > 0, 1, 0)
                        blnSetDefault = True
                    End If
                     .Cell(flexcpForeColor, i, .ColIndex("ȱʡ")) = lngEditForColor
                     .Cell(flexcpForeColor, i, .ColIndex("ҩ��")) = lngEditForColor:
                Else
                    If Val(.RowData(i)) = Val(strMachValue) And strMachValue <> "" And Not blnSetDefault Then
                        .TextMatrix(i, lngEditCol) = strDefaultValue
                    End If
                    '���ô���
                     .Cell(flexcpForeColor, i, .ColIndex("����")) = lngEditForColor
                End If
                .Cell(flexcpData, i, lngEditCol) = bytLockEdit
            End If
        Next
    End With
End Sub

Private Sub vsfDrugStore_EnterCell()
    Dim rsTmp As ADODB.Recordset, strList As String
    With vsfDrugStore
        If .Row > 0 Then
            If .Col = .ColIndex("����") Then
                Set rsTmp = Read��ҩ����(.RowData(.Row))
                strList = "�Զ�����|" & .BuildComboList(rsTmp, "����")
                .ColComboList(.Col) = strList
            Else
                .ColComboList(.Col) = ""
              '  .Editable = flexEDNone
            End If
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub chk�շ�ִ�п���_Click()
    If mblnNotClick Then Exit Sub
    If chk�շ�ִ�п���.Value = vbChecked Then
        cmd�շ�ִ�п���.Enabled = True
        Call cmd�շ�ִ�п���_Click
    Else
        txt�շ�ִ�п���.Text = ""
        txt�շ�ִ�п���.Tag = ""
        cmd�շ�ִ�п���.Enabled = False
    End If
End Sub

Private Sub cmd�շ�ִ�п���_Click()
    Dim rsDept As ADODB.Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandler
    '96357
    strSQL = "Select Distinct A.ID, A.����, A.����, A.����" & vbNewLine & _
            " From ���ű� A, ��������˵�� B" & vbNewLine & _
            " Where B.����ID=A.ID And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & vbNewLine & _
            "       And B.�������� In('��ҩ��', '��ҩ��', '��ҩ��', '���ϲ���')" & vbNewLine & _
            "       And B.������� In (1, 2, 3)" & vbNewLine & _
            "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " Order by A.����"
    vRect = zlControl.GetControlRect(txt�շ�ִ�п���.hWnd)
    Set rsDept = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "�����շ�ִ�п���", True, "", "", False, False, False, _
        vRect.Left, vRect.Top, txt�շ�ִ�п���.Height, blnCancel, False, True, "MultiCheckReturn=1")
    If blnCancel Then Exit Sub
    If rsDept Is Nothing Then Exit Sub
    
    txt�շ�ִ�п���.Text = ""
    txt�շ�ִ�п���.Tag = ""
    Do While Not rsDept.EOF
        txt�շ�ִ�п���.Text = txt�շ�ִ�п���.Text & ";" & NVL(rsDept!����)
        strTemp = strTemp & "," & NVL(rsDept!ID)
        rsDept.MoveNext
    Loop
    If txt�շ�ִ�п���.Text <> "" Then txt�շ�ִ�п���.Text = Mid(txt�շ�ִ�п���.Text, 2)
    If strTemp <> "" Then txt�շ�ִ�п���.Tag = Mid(strTemp, 2)
    
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

