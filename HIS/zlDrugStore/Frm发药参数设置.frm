VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm��ҩ�������� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "Frm��ҩ��������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   7380
   StartUpPosition =   1  '����������
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   3600
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab TabShow 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "����(&1)"
      TabPicture(0)   =   "Frm��ҩ��������.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "����(&2)"
      TabPicture(1)   =   "Frm��ҩ��������.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frm��ҩ��������"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "��ӡ(&3)"
      TabPicture(2)   =   "Frm��ҩ��������.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cbo��ҩ������ʽ"
      Tab(2).Control(1)=   "cbo��ҩ��ҩ��ʽ"
      Tab(2).Control(2)=   "cbo��ҩ������ʽ"
      Tab(2).Control(3)=   "cbo��ҩ��ҩ��ʽ"
      Tab(2).Control(4)=   "chkPreview"
      Tab(2).Control(5)=   "cboҩƷ��ǩ"
      Tab(2).Control(6)=   "cbo��ҩ��"
      Tab(2).Control(7)=   "Cbo��ҩ��"
      Tab(2).Control(8)=   "Fra��ӡ"
      Tab(2).Control(9)=   "Fraˢ��"
      Tab(2).Control(10)=   "lbl��ҩ������ʽ"
      Tab(2).Control(11)=   "lbl��ҩ��ҩ��ʽ"
      Tab(2).Control(12)=   "lbl��ҩ������ʽ"
      Tab(2).Control(13)=   "lbl��ҩ��ҩ��ʽ"
      Tab(2).Control(14)=   "lblҩƷ��ǩ"
      Tab(2).Control(15)=   "Lbl��ҩ"
      Tab(2).Control(16)=   "lbl��ҩ"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "Ʊ��(&4)"
      TabPicture(3)   =   "Frm��ҩ��������.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblƱ��"
      Tab(3).Control(1)=   "cmd��ӡ����"
      Tab(3).Control(2)=   "cboƱ������"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "��Դ����(&5)"
      TabPicture(4)   =   "Frm��ҩ��������.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Lvw��Դ����"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "��������(&6)"
      TabPicture(5)   =   "Frm��ҩ��������.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraSetColor"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "�Ŷӽк�(&7)"
      TabPicture(6)   =   "Frm��ҩ��������.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Fra�����豸����"
      Tab(6).Control(1)=   "frm��ʾ�豸����"
      Tab(6).Control(2)=   "chk�����Ŷӽк�"
      Tab(6).Control(3)=   "chkUseDisplay"
      Tab(6).ControlCount=   4
      Begin VB.ComboBox cbo��ҩ������ʽ 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -70320
         Style           =   2  'Dropdown List
         TabIndex        =   123
         Top             =   480
         Width           =   2280
      End
      Begin VB.ComboBox cbo��ҩ��ҩ��ʽ 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   480
         Width           =   2280
      End
      Begin VB.ComboBox cbo��ҩ������ʽ 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -70320
         Style           =   2  'Dropdown List
         TabIndex        =   119
         Top             =   840
         Width           =   2280
      End
      Begin VB.ComboBox cbo��ҩ��ҩ��ʽ 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   117
         Top             =   840
         Width           =   2280
      End
      Begin VB.CheckBox chkPreview 
         Caption         =   "��ӡ����ǩʱ��Ԥ���ٴ�ӡ"
         Height          =   195
         Left            =   -70920
         TabIndex        =   114
         Top             =   1608
         Width           =   2640
      End
      Begin VB.Frame Frm��ҩ�������� 
         Caption         =   "  ѡ�� "
         Height          =   4965
         Left            =   -74880
         TabIndex        =   84
         Top             =   480
         Width           =   6675
         Begin VB.ComboBox cbo�س���ʽ 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   125
            Top             =   3120
            Width           =   3360
         End
         Begin VB.CheckBox chkDispensing 
            Caption         =   "���������С����ܵ�ͬʱ֪ͨҩƷ�Զ����豸��׼����ҩ��"
            Height          =   225
            Left            =   120
            TabIndex        =   109
            Top             =   3960
            Width           =   5460
         End
         Begin VB.CheckBox chkɨ������ 
            Caption         =   "����ҩ����ɨ����Զ�����"
            Height          =   225
            Left            =   120
            TabIndex        =   108
            Top             =   3600
            Width           =   2460
         End
         Begin VB.ComboBox cbo�����ʾ 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4440
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   240
            Width           =   2160
         End
         Begin VB.CommandButton cmdDeviceSetup 
            Caption         =   "�豸����(&S)"
            Height          =   350
            Left            =   2160
            TabIndex        =   98
            Top             =   4320
            Width           =   1500
         End
         Begin VB.ListBox lst������ 
            Appearance      =   0  'Flat
            Columns         =   1
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   2280
            Style           =   1  'Checkbox
            TabIndex        =   95
            Top             =   2280
            Width           =   2760
         End
         Begin VB.Frame fraline1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   1560
            TabIndex        =   86
            Top             =   2040
            Width           =   650
         End
         Begin VB.ComboBox cboҩƷ������ʾ 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Top             =   240
            Width           =   2160
         End
         Begin VB.CheckBox chk��С��λ 
            Caption         =   "�����ֵ�λ��ʾҩƷ����"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1560
            TabIndex        =   91
            Top             =   840
            Width           =   2400
         End
         Begin VB.CheckBox Chk��ʾ������ 
            Caption         =   "��ʾ������"
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   840
            Width           =   1200
         End
         Begin VB.CheckBox chk��ҩɨ�� 
            Caption         =   "��ҩģʽ����ɨ����������ɨ��ȷ�ϣ�"
            Height          =   225
            Left            =   120
            TabIndex        =   89
            Top             =   1320
            Width           =   4140
         End
         Begin VB.CheckBox chkOverTime 
            Caption         =   "������ʾ"
            Height          =   225
            Left            =   120
            TabIndex        =   88
            Top             =   1800
            Width           =   1020
         End
         Begin VB.TextBox txtOverTime 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1665
            TabIndex        =   87
            Text            =   "1440"
            Top             =   1800
            Width           =   460
         End
         Begin VB.CheckBox chk��ҩˢ�� 
            Caption         =   "��ҩģʽ����ˢ����ҩ"
            Height          =   225
            Left            =   120
            TabIndex        =   85
            Top             =   2280
            Width           =   2100
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "����ʱϵͳ�Զ��س���ʽ"
            Height          =   180
            Left            =   120
            TabIndex        =   126
            Top             =   3180
            Width           =   1980
         End
         Begin VB.Label lbl�����ʾ 
            AutoSize        =   -1  'True
            Caption         =   "�����ʾ"
            Height          =   180
            Left            =   3600
            TabIndex        =   100
            Top             =   300
            Width           =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "���ܿ��������豸���� "
            Height          =   180
            Left            =   120
            TabIndex        =   99
            Top             =   4440
            Width           =   1890
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "ҩƷ������ʾ"
            Height          =   180
            Left            =   120
            TabIndex        =   94
            Top             =   300
            Width           =   1080
         End
         Begin VB.Label lblOverTime 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����       ����δ��ҩ��ҩƷ����"
            Height          =   180
            Left            =   1200
            TabIndex        =   93
            Top             =   1815
            Width           =   2790
         End
      End
      Begin VB.ComboBox cboҩƷ��ǩ 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   1560
         Width           =   2280
      End
      Begin VB.CheckBox chkUseDisplay 
         Caption         =   "��ʾ�ŶӶ���"
         Height          =   255
         Left            =   -74640
         TabIndex        =   77
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cbo��ҩ�� 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1200
         Width           =   2280
      End
      Begin VB.ComboBox Cbo��ҩ�� 
         ForeColor       =   &H80000012&
         Height          =   276
         IMEMode         =   3  'DISABLE
         Left            =   -70320
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   1200
         Width           =   2280
      End
      Begin VB.CheckBox chk�����Ŷӽк� 
         Caption         =   "�����Ŷӽк�"
         Height          =   255
         Left            =   -74640
         TabIndex        =   52
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame fraSetColor 
         Caption         =   " ������ӡ������ "
         Height          =   4965
         Left            =   -74880
         TabIndex        =   42
         Top             =   500
         Width           =   6795
         Begin VSFlex8Ctl.VSFlexGrid vsfPrinter 
            Height          =   3852
            Left            =   120
            TabIndex        =   116
            Top             =   600
            Width           =   6492
            _cx             =   11451
            _cy             =   6794
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
            Rows            =   50
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
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
            Editable        =   1
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
         Begin VB.CommandButton cmdDefaultPrinter 
            BackColor       =   &H00000000&
            Caption         =   "�ָ�Ĭ�ϴ�ӡ��(&P)"
            Height          =   300
            Left            =   3960
            MaskColor       =   &H00000000&
            TabIndex        =   51
            Top             =   4560
            Width           =   2655
         End
         Begin VB.Label Label6 
            Caption         =   "ѡ�񴦷���Ӧ�Ĵ�ӡ����������ҩ����ǩ����ҩ������"
            Height          =   432
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   4920
         End
      End
      Begin MSComctlLib.ListView Lvw��Դ���� 
         Height          =   4965
         Left            =   -74850
         TabIndex        =   40
         Top             =   495
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   8758
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.ComboBox cboƱ������ 
         Height          =   276
         Left            =   -73650
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   540
         Width           =   2565
      End
      Begin VB.CommandButton cmd��ӡ���� 
         Caption         =   "��ӡ����(&P)"
         Height          =   345
         Left            =   -74400
         TabIndex        =   35
         Top             =   990
         Width           =   3315
      End
      Begin VB.Frame Fra��ӡ 
         Caption         =   "  �Զ���ӡ"
         Height          =   1692
         Left            =   -74850
         TabIndex        =   39
         Top             =   2040
         Width           =   6705
         Begin VB.CheckBox chk���ķ��ϵ� 
            Caption         =   "��ӡ���ķ��ϵ�"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4320
            TabIndex        =   115
            Top             =   0
            Width           =   1920
         End
         Begin VB.CheckBox chkAllType 
            Caption         =   "�Զ���ӡ��ҩ��ʱ��ӡƱ�ݵ����и�ʽ"
            Height          =   195
            Left            =   2760
            TabIndex        =   112
            Top             =   315
            Width           =   3360
         End
         Begin VB.CheckBox chkҩƷ��ǩ 
            Caption         =   "��ӡҩƷ��ǩ"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2640
            TabIndex        =   41
            Top             =   0
            Width           =   1440
         End
         Begin VB.CheckBox chk���ʵ� 
            Caption         =   "��ӡʱ�������ʵ���"
            Height          =   195
            Left            =   345
            TabIndex        =   18
            Top             =   315
            Width           =   1920
         End
         Begin VB.ListBox lst��ӡ���� 
            Appearance      =   0  'Flat
            Columns         =   1
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   2760
            Style           =   1  'Checkbox
            TabIndex        =   22
            Top             =   600
            Width           =   2760
         End
         Begin VB.CheckBox Chk��ӡ��ҩ�� 
            Caption         =   "��ӡ��ҩ��"
            Height          =   210
            Left            =   1320
            TabIndex        =   17
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Opt��ӡ��ҩ�������� 
            Caption         =   "��ӡ�����ŵ���ҩ��"
            Enabled         =   0   'False
            Height          =   180
            Left            =   780
            TabIndex        =   19
            Top             =   615
            Width           =   1935
         End
         Begin VB.OptionButton Opt��ӡ��ҩ�������� 
            Caption         =   "��ӡ�����ڵ���ҩ��"
            Enabled         =   0   'False
            Height          =   180
            Left            =   780
            TabIndex        =   20
            Top             =   1027
            Width           =   1935
         End
         Begin VB.OptionButton Opt��ӡ��ҩ��ѡ�� 
            Caption         =   "��ӡָ��������ҩ��"
            Enabled         =   0   'False
            Height          =   180
            Left            =   780
            TabIndex        =   21
            Top             =   1440
            Width           =   2100
         End
      End
      Begin VB.Frame Fraˢ�� 
         Caption         =   "  �Զ�ˢ�� "
         Height          =   1680
         Left            =   -74850
         TabIndex        =   23
         Top             =   3840
         Width           =   6705
         Begin VB.TextBox txt��ӡ��� 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1620
            MaxLength       =   2
            TabIndex        =   47
            Top             =   240
            Width           =   1125
         End
         Begin VB.TextBox Txt��ӡ�˷ѵ��� 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1620
            MaxLength       =   2
            TabIndex        =   31
            Top             =   1320
            Width           =   1125
         End
         Begin VB.TextBox Txt�ӳٴ�ӡ 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1620
            MaxLength       =   2
            TabIndex        =   28
            Top             =   960
            Width           =   1125
         End
         Begin VB.TextBox Txtˢ�¼�� 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1620
            MaxLength       =   2
            TabIndex        =   25
            Top             =   600
            Width           =   1125
         End
         Begin VB.Label lblRefreshComment 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������Ϣ�������"
            Height          =   180
            Left            =   3120
            TabIndex        =   97
            Top             =   660
            Width           =   1620
         End
         Begin VB.Label lblPrintComment 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������Ϣ�������"
            Height          =   180
            Left            =   3120
            TabIndex        =   96
            Top             =   300
            Width           =   1620
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ӡ���"
            Height          =   180
            Left            =   840
            TabIndex        =   49
            Top             =   300
            Width           =   720
         End
         Begin VB.Label LblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   4
            Left            =   2760
            TabIndex        =   48
            Top             =   285
            Width           =   180
         End
         Begin VB.Label LblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   2760
            TabIndex        =   32
            Top             =   1380
            Width           =   180
         End
         Begin VB.Label LblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   2760
            TabIndex        =   29
            Top             =   1020
            Width           =   180
         End
         Begin VB.Label LblNote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   2760
            TabIndex        =   26
            Top             =   660
            Width           =   180
         End
         Begin VB.Label Lbl��ӡ�˷ѵ��� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ӡ�˷ѵ��ݼ��"
            Height          =   180
            Left            =   120
            TabIndex        =   30
            Top             =   1380
            Width           =   1440
         End
         Begin VB.Label Lbl�ӳٴ�ӡ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�ӳٴ�ӡ"
            Height          =   180
            Left            =   840
            TabIndex        =   27
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Lblˢ�¼�� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ˢ�¼��"
            Height          =   180
            Left            =   840
            TabIndex        =   24
            Top             =   660
            Width           =   720
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "  ѡ�� "
         Height          =   5025
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   6795
         Begin VB.CheckBox chkSame 
            Caption         =   "������ҩ�˺ͺ˲�����ͬ"
            Height          =   195
            Left            =   3480
            TabIndex        =   113
            Top             =   3413
            Width           =   2640
         End
         Begin VB.ComboBox cboCheck 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   110
            Top             =   3360
            Width           =   2280
         End
         Begin VB.ComboBox cbo����ҩ 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   4560
            Width           =   2280
         End
         Begin VB.CheckBox chkIsDosage 
            Caption         =   "��ǰҩ����Ҫ��ҩ����"
            Height          =   225
            Left            =   195
            TabIndex        =   105
            Top             =   1855
            Width           =   2940
         End
         Begin VB.CheckBox chkIsDosageOk 
            Caption         =   "��ǰҩ����Ҫ��ҩȷ��(����ǩ��)����"
            Height          =   225
            Left            =   195
            TabIndex        =   104
            Top             =   1560
            Width           =   3540
         End
         Begin VB.CheckBox chkSign 
            Caption         =   "ǩ��ʱ�Զ�������ҩ(ҩ������ǩ����Ч)"
            Height          =   180
            Left            =   195
            TabIndex        =   103
            Top             =   2150
            Width           =   3615
         End
         Begin VB.CheckBox chkCheckStuff 
            Caption         =   "��ҩ�������ķ������"
            Height          =   180
            Left            =   195
            TabIndex        =   102
            Top             =   2400
            Width           =   2295
         End
         Begin VB.ComboBox cbo�������� 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   1080
            Width           =   2280
         End
         Begin VB.TextBox txt��ҩʱ�� 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3150
            TabIndex        =   45
            Top             =   3060
            Width           =   525
         End
         Begin VB.CheckBox chk�Զ���ҩ 
            Caption         =   "�Զ���ҩģʽ"
            Height          =   195
            Left            =   240
            TabIndex        =   43
            Top             =   3120
            Width           =   1440
         End
         Begin VB.ComboBox cbo��λ 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   660
            Width           =   2280
         End
         Begin VB.ComboBox cbo���ʴ��� 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   4170
            Width           =   2280
         End
         Begin VB.TextBox txt��ѯ���� 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4275
            TabIndex        =   15
            Text            =   "1"
            Top             =   4560
            Width           =   885
         End
         Begin VB.ListBox lst��ҩ���� 
            Appearance      =   0  'Flat
            Columns         =   1
            ForeColor       =   &H80000012&
            Height          =   1710
            IMEMode         =   3  'DISABLE
            Left            =   4680
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   300
            Width           =   1800
         End
         Begin VB.ComboBox cbo�շѴ��� 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   3780
            Width           =   2280
         End
         Begin VB.ComboBox Cbo��ҩ�� 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2760
            Width           =   2280
         End
         Begin VB.ComboBox Cboҩ�� 
            ForeColor       =   &H80000012&
            Height          =   276
            IMEMode         =   3  'DISABLE
            Left            =   1320
            TabIndex        =   3
            Text            =   "Cboҩ��"
            Top             =   240
            Width           =   2280
         End
         Begin VB.Label lblCheck 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�˲���"
            Height          =   180
            Left            =   360
            TabIndex        =   111
            Top             =   3420
            Width           =   540
         End
         Begin VB.Label lbl��ҩ��ӡ״̬ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ҩ����"
            Height          =   180
            Left            =   60
            TabIndex        =   107
            Top             =   4620
            Width           =   900
         End
         Begin VB.Label lbl����סԺ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����סԺ����"
            Height          =   180
            Left            =   120
            TabIndex        =   81
            Top             =   1140
            Width           =   1080
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   3720
            TabIndex        =   46
            Top             =   3120
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�Զ���ҩʱ��"
            Height          =   180
            Left            =   2040
            TabIndex        =   44
            Top             =   3120
            Width           =   1080
         End
         Begin VB.Label lbl��λ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ҩ������"
            Height          =   180
            Left            =   480
            TabIndex        =   12
            Top             =   720
            Width           =   720
         End
         Begin VB.Label lbl���ʴ��� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���ʴ���"
            Height          =   180
            Left            =   240
            TabIndex        =   10
            Top             =   4230
            Width           =   720
         End
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Left            =   5280
            TabIndex        =   16
            Top             =   4620
            Width           =   180
         End
         Begin VB.Label lbl��ѯ���� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ѯ����"
            Height          =   180
            Left            =   3480
            TabIndex        =   14
            Top             =   4620
            Width           =   720
         End
         Begin VB.Label lbl�շѴ��� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�շѴ���"
            Height          =   180
            Left            =   240
            TabIndex        =   8
            Top             =   3840
            Width           =   720
         End
         Begin VB.Label Lbl��ҩ�� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩ��"
            Height          =   180
            Left            =   360
            TabIndex        =   6
            Top             =   2820
            Width           =   540
         End
         Begin VB.Label Lbl��ҩ���� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩ����"
            Height          =   180
            Left            =   3840
            TabIndex        =   4
            Top             =   300
            Width           =   720
         End
         Begin VB.Label Lblҩ�� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩҩ��"
            Height          =   180
            Left            =   480
            TabIndex        =   2
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame frm��ʾ�豸���� 
         Height          =   855
         Left            =   -74850
         TabIndex        =   53
         Top             =   840
         Width           =   6795
         Begin VB.CommandButton cmd��ʾ�豸���� 
            Caption         =   "�豸����"
            Height          =   300
            Left            =   4320
            TabIndex        =   55
            Top             =   300
            Width           =   1100
         End
         Begin VB.ComboBox cbo��ʾӲ����� 
            Height          =   300
            ItemData        =   "Frm��ҩ��������.frx":03CE
            Left            =   1560
            List            =   "Frm��ҩ��������.frx":03D0
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   300
            Width           =   2535
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "��ʾ�豸���"
            Height          =   180
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.Frame Fra�����豸���� 
         Height          =   3735
         Left            =   -74850
         TabIndex        =   57
         Top             =   1800
         Width           =   6795
         Begin VB.OptionButton optCallWay 
            Caption         =   "���ñ�������"
            Height          =   330
            Index           =   0
            Left            =   240
            TabIndex        =   76
            Top             =   320
            Width           =   1455
         End
         Begin VB.CheckBox chkUseSound 
            Caption         =   "������������"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optCallWay 
            Caption         =   "����Զ������"
            Height          =   330
            Index           =   1
            Left            =   240
            TabIndex        =   58
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Frame frm�����㲥���� 
            Height          =   1455
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Width           =   6495
            Begin VB.OptionButton optSoundType 
               Caption         =   "΢������"
               Height          =   255
               Index           =   1
               Left            =   2400
               TabIndex        =   71
               Top             =   338
               Width           =   1095
            End
            Begin VB.OptionButton optSoundType 
               Caption         =   "ϵͳ����"
               Height          =   255
               Index           =   0
               Left            =   1200
               TabIndex        =   70
               Top             =   338
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.TextBox txtSpeed 
               Height          =   270
               Left            =   1080
               TabIndex        =   66
               Text            =   "65"
               Top             =   685
               Width           =   495
            End
            Begin VB.CommandButton cmdTestSound 
               Caption         =   "��������"
               Height          =   350
               Left            =   4080
               TabIndex        =   65
               Top             =   645
               Width           =   1100
            End
            Begin VB.TextBox txtPlayCount 
               Height          =   270
               Left            =   1080
               TabIndex        =   64
               Text            =   "1"
               Top             =   1035
               Width           =   615
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Left            =   120
               TabIndex        =   69
               Top             =   375
               Width           =   720
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "�������٣�      (��Χ��0��100֮�䣬�Ƽ�65)"
               Height          =   180
               Left            =   120
               TabIndex        =   68
               Top             =   730
               Width           =   3780
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "���Ŵ���Ϊ        ��"
               Height          =   180
               Left            =   120
               TabIndex        =   67
               Top             =   1080
               Width           =   1800
            End
         End
         Begin VB.Frame FraԶ���������� 
            Height          =   1215
            Left            =   120
            TabIndex        =   60
            Top             =   2160
            Width           =   6495
            Begin VB.TextBox txtLoopQueryTime 
               Height          =   270
               Left            =   1800
               MaxLength       =   3
               TabIndex        =   78
               Text            =   "10"
               Top             =   840
               Width           =   615
            End
            Begin VB.ComboBox cboWorkStation 
               Height          =   300
               Left            =   1200
               TabIndex        =   61
               Top             =   360
               Width           =   3375
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "������ѯ���ʱ��Ϊ        ��"
               Height          =   180
               Left            =   120
               TabIndex        =   79
               Top             =   885
               Width           =   2520
            End
            Begin VB.Label labRemoteComputerName 
               Caption         =   "Զ��վ������"
               Height          =   255
               Left            =   120
               TabIndex        =   62
               Top             =   405
               Width           =   1215
            End
         End
      End
      Begin VB.Label lbl��ҩ������ʽ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ������ʽ"
         Height          =   180
         Left            =   -71460
         TabIndex        =   124
         Top             =   540
         Width           =   1080
      End
      Begin VB.Label lbl��ҩ��ҩ��ʽ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ��ҩ��ʽ"
         Height          =   180
         Left            =   -74940
         TabIndex        =   122
         Top             =   540
         Width           =   1080
      End
      Begin VB.Label lbl��ҩ������ʽ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ������ʽ"
         Height          =   180
         Left            =   -71460
         TabIndex        =   120
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label lbl��ҩ��ҩ��ʽ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ��ҩ��ʽ"
         Height          =   180
         Left            =   -74940
         TabIndex        =   118
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label lblҩƷ��ǩ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ��ǩ"
         Height          =   180
         Left            =   -74580
         TabIndex        =   83
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Lbl��ҩ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ǩ"
         Height          =   180
         Left            =   -70920
         TabIndex        =   75
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label lbl��ҩ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ��"
         Height          =   180
         Left            =   -74400
         TabIndex        =   74
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label lblƱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ��(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74370
         TabIndex        =   33
         Top             =   600
         Width           =   630
      End
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   38
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6000
      TabIndex        =   37
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4800
      TabIndex        =   36
      Top             =   6000
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ҩ��������.frx":03D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ҩ��������.frx":06EC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Frm��ҩ��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--ע�����ر���--
Private intDays As Integer
Private intUnit As Integer                              'ȱʡ��λ��0-����Ӧ;1-����ҩ����λ;2-סԺҩ����λ��
Private intPrint As Integer                             '����ӡδ��ҩ����(0)
Private intУ�鷽ʽ As Integer                          'У�鷽ʽ
Private intУ����ҩ�� As Integer                        '��ҩʱ�Ƿ�У����ҩ��
Private intУ�鷢ҩ�� As Integer                        '��ҩʱ�Ƿ�У�鷢ҩ��
Private mint���ʵ� As Integer                           '��ӡ��ҩ��ʱ�Ƿ�������ʵ�
Private mintҩƷ��ǩ As Integer                         '��ӡҩƷ��ǩ
Private mint���ķ��ϵ� As Integer                       '��ӡ���ķ��ϵ�
Private strPrintWindow As String                        '��ӡδ��ҩ����Ϊ3ʱ��Ч
'0-����ӡδ��ҩ����
'1-��ӡ����������δ��ҩ����
'2-��ӡ����������δ��ҩ����
'3-ѡ���ӡ(��ҩ����)

Private IntRefresh As Integer                           'ˢ�¼��(0)
Private intPrintDelay As Integer                        '�ӳٴ�ӡ(60)
Private intPrintHandbackNO As Integer                   '��ӡ�˷ѵ��ݺ�(0)
Private mintPrintInterval As Integer                    '��ӡ��ҩ�����(0)
Private lngҩ��ID As Long                               'ҩ��(���ñ�������Ӧ��ҩ��)
Private Str���� As String                               '��ҩ����(���ñ�������Ӧ�ķ�ҩ����)
Private str��ҩ�� As String                             '������ҩ��
Private mint�Զ���ҩ As Integer                         '�Ƿ�ʹ���Զ���ҩ���ܣ�0-��ʹ�ã�1-ʹ��
Private mint�Զ���ҩʱ�� As Integer                     '������ʱ�޾���Ҫ��֤��ҩ�ˣ�Ĭ��Ϊʼ�ղ���֤��ҩ��
Private mintˢ��֤ As Integer                           '��ҩ���Ƿ����ˢ����֤��0-��ˢ��;1-Ҫˢ��
Private mint��ҩɨ�� As Integer                         '��ҩģʽ����ɨ������0-������;1-����
Private mint�����Ŷӽк� As Integer                     '�Ƿ������ŶӽкŹ���
Private mintSign As Integer                             'ǩ��ʱ������ҩ
Private mblnLoadDrug As Boolean
Private mblnUseMsg As Boolean                           '�Ƿ���������Ϣ����
Private mstr����ˢ����ҩ As String                      '����ˢ����ҩ����ʽ�������1,�����2......����������ݱ�ʾ������
Private mint����ʱ����� As Integer                     'ҩƷҽ��������ʱ�䣨�״�ʱ�䣩���ˣ�0-������ʱ����ˣ�1-������ʱ�����
Private mint�����ʾ��ʽ As Integer                     '0-��ʾӦ�ս�1-��ʾʵ�ս�2-��ʾӦ�ս���ʵ�ս��
Private mint����ȡҩģʽ As Integer                     '����ȡҩģʽ��0-�����ã�1-����
Private mint��ҩ���� As Integer                       '��ҩ���Ƿ����б�ҩ���˲�����δ�������ĵ���
Private mintɨ������ As Integer                       '0-���Զ�����,1-ɨ����Զ���������
Private mstr�˲��� As String
Private mintRowNum As Integer


Private mintShowName As Integer                         'ҩƷ������ʾ��ʽ��0-���ƺͱ��룻1-�����룻2-������
Private mintType As Integer                             '�������ͣ�0-��ʾ�����סԺ������1-ֻ��ʾ���ﴦ����2-ֻ��ʾסԺ����

Private IntShowCol As Integer                           '�ڴ�����ϸ���Ƿ���ʾ����(0)
Private mintShowBill�շ� As Integer                     '�շѴ�����ʾ��Χ
Private mintShowBill���� As Integer                     '���ʴ�����ʾ��Χ
Private mintShowBill��ҩ As Integer                     '����ҩ����ӡ״̬��ʾ��Χ
Private IntAutoPrint As Integer                         '��ҩ���ӡ������(1)
Private mint��ҩ���Զ���ӡ As Integer                   '�Զ���ӡ��ҩ��
Private mint��ҩ���Զ���ӡҩƷ��ǩ As Integer           '��ҩ���Զ���ӡҩƷ��ǩ
Private mstrWin As String                               '��ҩ���ڴ�
Private mint�س���ʽ As Integer                         'ͨ��¼���ˢ������ʱϵͳ�Զ���ӻس�����ķ�ʽ��0-ϵͳ���Զ��س�,1-��¼��ﵽ��Ŀ�򿨺ų���ʱ�Զ��س�

Private mIntCol���� As Integer
Private mintCol��ʽ As Integer
Private mintCol��ӡ�� As Integer

Private Const mconstr���� = "��ͨ;����;����;�������;�������;����"
Private Const mconlng��ɫ = "&HFFFFFF;&HC0FFC0;&HC0FFFF;&HFFFFFF;&HC0C0FF;&HC0C0FF"

Public mstrPrivs As String                              'Ȩ�޴�
Private mblnSetPara As Boolean                          '�Ƿ���в�������Ȩ��
Private mstrRPTDefaultScheme_Recipt As String           '����ǩ�����Ĭ�ϸ�ʽ

'�Ŷӽк�ʹ�õĲ���

Private Type Type_Call
    int�����Ŷӽк� As Integer
    int�������� As Integer
'    int��ʾģʽ As Integer
    int��ʾ�ŶӶ��� As Integer
    int������������ As Integer
    int�кŷ�ʽ As Integer
    strԶ�˺���վ�� As String
    int�����㲥���� As Integer
    int�������Ŵ��� As Integer
    int��ѯʱ�� As Integer
End Type

Private mType_Call As Type_Call
'--��������ʹ�õĶ���--
Public RecPart As New ADODB.Recordset                   'ҩ��
Private RecPeople As New ADODB.Recordset                'ҩ����ҩ��
Private BlnStartUp  As Boolean                          '�Ƿ������ɹ�
Public strShow As String                                '��ʾ��
Private mstrSourceDep As String                         '��Դ���Ҵ�

Private mstrPrinters As String                          '���ش�ӡ���б���;�ָ�

'�������ͣ���ͨ��������ơ�������һ������
Private Enum ��������
    ��ͨ = 0
    ���� = 1
    ���� = 2
    ���� = 3
    ��һ = 4
    ���� = 5
End Enum

'Ĭ�ϴ�����ɫ����ͨ����ɫ���������ɫ�����ƣ�����ɫ��������һ������ɫ����������ɫ
Private Const mconlng��ͨ = &HFFFFFF
Private Const mconlng���� = &HC0FFC0
Private Const mconlng���� = &HC0FFFF
Private Const mconlng���� = &HFFFFFF
Private Const mconlng��һ = &HC0C0FF
Private Const mconlng���� = &HC0C0FF

Public Property Get In_���÷�ҩ() As Boolean
    In_���÷�ҩ = mblnLoadDrug
End Property

Public Property Let In_���÷�ҩ(ByVal vNewValue As Boolean)
    mblnLoadDrug = vNewValue
End Property

Public Property Get In_������Ϣ() As Boolean
    In_������Ϣ = mblnUseMsg
End Property

Public Property Let In_������Ϣ(ByVal vNewValue As Boolean)
    mblnUseMsg = vNewValue
End Property









Private Sub LoadList()
    Dim rs��ҩ��ʽ As New ADODB.Recordset
    Dim rs��ҩ��ʽ As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim str��ҩ��ʽ As String
    Dim str������ʽ As String
    Dim str��� As String
    Dim strPrinter As String
    Dim strPrinters As String
    Dim strColor As String
    Dim myPrinter As Printer
    Dim n As Integer
    Dim i As Integer
    
    On Error GoTo errHandle
    
    mIntCol���� = 0
    mintCol��ʽ = 1
    mintCol��ӡ�� = 2
    
    '��ȡ�����ʽ
    '--��ҩ
    str��� = "ZL1_BILL_1341_3"
    
    gstrSQL = "Select b.˵�� From zlReports A, zlRPTFMTs B Where a.Id = b.����id And a.��� = [1] order by b.���"
    
    Set rs��ҩ��ʽ = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ�����ʽ", str���)
    
    '--��ҩ
    str��� = "ZL1_BILL_1341_4"
    
    gstrSQL = "Select b.˵�� From zlReports A, zlRPTFMTs B Where a.Id = b.����id And a.��� = [1] order by b.���"
    
    Set rs��ҩ��ʽ = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ�����ʽ", str���)
    
    '��ȡ�������͵���ɫ����
    strColor = zldatabase.GetPara("������ɫ", glngSys, 1341, "", , mblnSetPara)
    
    '��ȡ����Ĵ�ӡ����������
    strPrinter = zldatabase.GetPara("������Ӧ�Ĵ�ӡ��", glngSys, 1341, "", , mblnSetPara)
    
    '��ȡ��Ӧ�Ĵ�ӡ��ʽ����
    str��ҩ��ʽ = zldatabase.GetPara("��ҩ����ӡ��ʽ", glngSys, 1341, "2;2", , mblnSetPara)
    str������ʽ = zldatabase.GetPara("����ǩ��ӡ��ʽ", glngSys, 1341, "1;1", , mblnSetPara)
    
    '��Ӵ�ӡ��ʽ�������б�
    With rs��ҩ��ʽ
        For n = 1 To .RecordCount
            cbo��ҩ��ҩ��ʽ.AddItem !˵��
            cbo��ҩ��ҩ��ʽ.ItemData(cbo��ҩ��ҩ��ʽ.NewIndex) = n
            cbo��ҩ������ʽ.AddItem !˵��
            cbo��ҩ������ʽ.ItemData(cbo��ҩ������ʽ.NewIndex) = n
            .MoveNext
        Next
    End With
    
    With rs��ҩ��ʽ
        For n = 1 To .RecordCount
            cbo��ҩ��ҩ��ʽ.AddItem !˵��
            cbo��ҩ��ҩ��ʽ.ItemData(cbo��ҩ��ҩ��ʽ.NewIndex) = n
            cbo��ҩ������ʽ.AddItem !˵��
            cbo��ҩ������ʽ.ItemData(cbo��ҩ������ʽ.NewIndex) = n
            .MoveNext
        Next
    End With
    
    '�����û����õĴ�ӡ��ʽ
    '--��ҩ
    For i = 0 To cbo��ҩ��ҩ��ʽ.ListCount - 1
        If Val(Split(str��ҩ��ʽ, ";")(0)) = cbo��ҩ��ҩ��ʽ.ItemData(i) Then
            cbo��ҩ��ҩ��ʽ.ListIndex = i
        End If
    Next
    
    For i = 0 To cbo��ҩ������ʽ.ListCount - 1
        If Val(Split(str������ʽ, ";")(0)) = cbo��ҩ������ʽ.ItemData(i) Then
            cbo��ҩ������ʽ.ListIndex = i
        End If
    Next
    '--��ҩ
    For i = 0 To cbo��ҩ��ҩ��ʽ.ListCount - 1
        If Val(Split(str��ҩ��ʽ, ";")(1)) = cbo��ҩ��ҩ��ʽ.ItemData(i) Then
            cbo��ҩ��ҩ��ʽ.ListIndex = i
        End If
    Next
    
    For i = 0 To cbo��ҩ������ʽ.ListCount - 1
        If Val(Split(str������ʽ, ";")(1)) = cbo��ҩ������ʽ.ItemData(i) Then
            cbo��ҩ������ʽ.ListIndex = i
        End If
    Next
    
    '���뱾�ش�ӡ���б�
    mstrPrinters = ""
    For Each myPrinter In Printers
        mstrPrinters = IIf(mstrPrinters = "", "", mstrPrinters & ";") & myPrinter.DeviceName
    Next
    
    For n = 0 To UBound(Split(mstrPrinters, ";"))
        If Split(mstrPrinters, ";")(n) <> "" Then
            strPrinters = strPrinters & "|" & Split(mstrPrinters, ";")(n)
        End If
    Next
    strPrinters = Mid(strPrinters, 2)
    
    'װ�ر��ؼ�¼��
    With rsData
        If .State = 1 Then .Close
        
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ʽ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ӡ��", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    '�ж����ݵĺϷ��ԣ�����������
    If UBound(Split(strPrinter, ";")) <> UBound(Split(mconstr����, ";")) Then
        For n = 0 To UBound(Split(mconstr����, ";"))
            strPrinter = strPrinter & ";"
        Next
    End If
    
    '�򱾵ؼ�¼�������û�����Ĵ�ӡ������
    For n = 0 To UBound(Split(mconstr����, ";"))
        rs��ҩ��ʽ.MoveFirst
        If InStr(strPrinter, "?") = 0 Then
            For i = 1 To rs��ҩ��ʽ.RecordCount
                rsData.AddNew
                
                rsData!���� = Split(mconstr����, ";")(n)
                rsData!��ʽ = rs��ҩ��ʽ!˵��
                rsData!��ӡ�� = Split(strPrinter, ";")(n)
    
                rsData.Update
                rs��ҩ��ʽ.MoveNext
            Next
        Else
            For i = 0 To UBound(Split(Split(strPrinter, ";")(n), ","))
                rsData.AddNew
                
                rsData!���� = Split(mconstr����, ";")(n)
                rsData!��ʽ = Mid(Split(Split(strPrinter, ";")(n), ",")(i), 1, InStr(Split(Split(strPrinter, ";")(n), ",")(i), "?") - 1)
                rsData!��ӡ�� = Mid(Split(Split(strPrinter, ";")(n), ",")(i), InStr(Split(Split(strPrinter, ";")(n), ",")(i), "?") + 1)
             
            Next
        End If
        rsData.Update
    Next
        
    With vsfPrinter
        .rows = rs��ҩ��ʽ.RecordCount * 6
        .Cols = 3
        .AllowSelection = False
        .ColAlignment(mIntCol����) = flexAlignCenterCenter
        .RowHeight(-1) = 250
        .ColWidth(mIntCol����) = 900
        .ColWidth(mintCol��ʽ) = 1500
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(mIntCol����) = True
        
        '���ش�ӡ��ѡ�������
        .ColComboList(mintCol��ӡ��) = strPrinters
        
        '����[����&��ɫ]��[��ʽ]
        For n = 0 To UBound(Split(mconstr����, ";"))
            rs��ҩ��ʽ.MoveFirst
            For i = 1 To rs��ҩ��ʽ.RecordCount
                .TextMatrix(n * rs��ҩ��ʽ.RecordCount + i - 1, mIntCol����) = Split(mconstr����, ";")(n)
                
                If strColor <> "" Then
                    .Cell(flexcpBackColor, n * rs��ҩ��ʽ.RecordCount + i - 1, mIntCol����) = Val(Split(strColor, ";")(n))
                Else
                    .Cell(flexcpBackColor, n * rs��ҩ��ʽ.RecordCount + i - 1, mIntCol����) = Split(mconlng��ɫ, ";")(n)
                End If
                
                .TextMatrix(n * rs��ҩ��ʽ.RecordCount + i - 1, mintCol��ʽ) = rs��ҩ��ʽ!˵��
                
                rs��ҩ��ʽ.MoveNext
            Next
        Next
        
        '�����û�����Ĵ�ӡ������
        For n = 0 To .rows - 1
            rsData.Filter = "���� = '" & .TextMatrix(n, mIntCol����) & "' and ��ʽ = '" & .TextMatrix(n, mintCol��ʽ) & "'"
            If rsData.RecordCount > 0 Then
                If InStr(strPrinters & "|", rsData!��ӡ�� & "|") > 0 Then   '���ô�ӡ�������Ƿ����
                    .TextMatrix(n, mintCol��ӡ��) = rsData!��ӡ��
                End If
            End If
        Next
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ReadFromReg()
    Dim strTmp As String
    Dim intOverTime As Integer
    Dim intParaType As Integer
    
    On Error Resume Next
    
    mblnSetPara = zlStr.IsHavePrivs(mstrPrivs, "��������")

    'ȡ˽�в���
    mintShowBill�շ� = Val(zldatabase.GetPara("�շѴ�����ʾ��ʽ", glngSys, 1341, 3, Array(lbl�շѴ���, cbo�շѴ���), mblnSetPara))
    mintShowBill���� = Val(zldatabase.GetPara("���ʴ�����ʾ��ʽ", glngSys, 1341, 3, Array(lbl���ʴ���, cbo���ʴ���), mblnSetPara))
    mintShowBill��ҩ = Val(zldatabase.GetPara("����ҩ���ݴ�ӡ��ʾ��ʽ", glngSys, 1341, 0, Array(lbl��ҩ��ӡ״̬, cbo����ҩ), mblnSetPara))
    intDays = Val(zldatabase.GetPara("��ѯ����", glngSys, 1341, 1, Array(lbl��ѯ����, txt��ѯ����, lbl����), mblnSetPara))
    mint���ʵ� = Val(zldatabase.GetPara("��ӡ�������ʵ�", glngSys, 1341, 0, Array(chk���ʵ�), mblnSetPara))
    intPrintHandbackNO = Val(zldatabase.GetPara("��ӡ�˷ѵ��ݼ��", glngSys, 1341, 0, Array(Lbl��ӡ�˷ѵ���, Txt��ӡ�˷ѵ���, lblNote(2)), mblnSetPara))
    intPrintDelay = Val(zldatabase.GetPara("��ӡ�ӳ�", glngSys, 1341, 60, Array(Lbl�ӳٴ�ӡ, Txt�ӳٴ�ӡ, lblNote(1)), mblnSetPara))
    IntRefresh = Val(zldatabase.GetPara("ˢ�¼��", glngSys, 1341, 0, Array(Lblˢ�¼��, Txtˢ�¼��, lblNote(0)), mblnSetPara))
    mintPrintInterval = Val(zldatabase.GetPara("��ӡ���", glngSys, 1341, 0, Array(Label3, txt��ӡ���, lblNote(4)), mblnSetPara))
    IntShowCol = Val(zldatabase.GetPara("��ʾ����", glngSys, 1341, 0, Array(Chk��ʾ������), mblnSetPara))
    IntAutoPrint = Val(zldatabase.GetPara("��ҩ���Զ���ӡ", glngSys, 1341, 0, Array(Lbl��ҩ, Cbo��ҩ��), mblnSetPara))
    intUnit = Val(zldatabase.GetPara("ҩ������", glngSys, 1341, 0, Array(lbl��λ, cbo��λ), mblnSetPara))
    mint��ҩ���Զ���ӡ = Val(zldatabase.GetPara("��ҩ���Զ���ӡ", glngSys, 1341, 2, Array(lbl��ҩ, cbo��ҩ��), mblnSetPara))
    mint��ҩ���Զ���ӡҩƷ��ǩ = Val(zldatabase.GetPara("��ҩ���ӡҩƷ��ǩ", glngSys, 1341, 2, Array(lblҩƷ��ǩ, cboҩƷ��ǩ), mblnSetPara))
    
    mint��ҩɨ�� = Val(zldatabase.GetPara("��ҩģʽɨ����ȷ��", glngSys, 1341, 0, Array(chk��ҩɨ��), mblnSetPara))
    intOverTime = Val(zldatabase.GetPara("��ʱδ��ҩƷ��ʾʱ����", glngSys, 1341, 0, Array(chkOverTime, lblOverTime, txtOverTime, fraline1), mblnSetPara))
    mintType = Val(zldatabase.GetPara("������סԺ����", glngSys, 1341, 0, Array(lbl����סԺ, cbo��������), mblnSetPara))
    mintSign = Val(zldatabase.GetPara("ǩ��ʱ������ҩ", glngSys, 1341, 0, Array(chkSign), mblnSetPara))
    mstr����ˢ����ҩ = zldatabase.GetPara("����ˢ����ҩ", glngSys, 1341, "", Array(chk��ҩˢ��, lst������), mblnSetPara)
    
    mint�����ʾ��ʽ = Val(zldatabase.GetPara("�����ʾ��ʽ", glngSys, 1341, 0, Array(lbl�����ʾ, cbo�����ʾ), mblnSetPara))
    
    mint��ҩ���� = zldatabase.GetPara("��ҩ�������ķ������", glngSys, 1341, 0, Array(chkCheckStuff), mblnSetPara)
    mintɨ������ = Val(zldatabase.GetPara("����ҩ����ɨ����Զ�����", glngSys, 1341, 0, Array(chkɨ������), mblnSetPara))
    mint�س���ʽ = Val(zldatabase.GetPara("����ʱϵͳ�Զ��س���ʽ", glngSys, 1341, 0, Array(cbo�س���ʽ), mblnSetPara))
   
    '0-����ӡδ��ҩ����
    '1-��ӡ����������δ��ҩ����
    '2-��ӡ����������δ��ҩ����
    '3-ѡ���ӡ(��ҩ����)
    intPrint = Val(zldatabase.GetPara("�����µ����Ƿ��ӡ", glngSys, 1341, 0, Array(Chk��ӡ��ҩ��), mblnSetPara))
    
    mintҩƷ��ǩ = Val(zldatabase.GetPara("��ӡҩƷ��ǩ", glngSys, 1341, 0, Array(chkҩƷ��ǩ), mblnSetPara))
    mint���ķ��ϵ� = Val(zldatabase.GetPara("��ӡ���ķ��ϵ�", glngSys, 1341, 0, Array(chk���ķ��ϵ�), mblnSetPara))
    lngҩ��ID = Val(zldatabase.GetPara("��ҩҩ��", glngSys, 1341, 0, Array(lblҩ��, , Cboҩ��), mblnSetPara))
    Str���� = zldatabase.GetPara("��ҩ����", glngSys, 1341, "", Array(Lbl��ҩ����, lst��ҩ����), mblnSetPara)
    str��ҩ�� = zldatabase.GetPara("��ҩ��", glngSys, 1341, "", Array(Lbl��ҩ��, cbo��ҩ��), mblnSetPara)
    strPrintWindow = zldatabase.GetPara("��ӡָ����ҩ����", glngSys, 1341, "", Array(Opt��ӡ��ҩ��ѡ��, lst��ӡ����), mblnSetPara)
    mstrSourceDep = zldatabase.GetPara("��Դ����", glngSys, 1341, "", Array(Lvw��Դ����), mblnSetPara)
    mint�Զ���ҩ = Val(zldatabase.GetPara("�Զ���ҩ", glngSys, 1341, 0, Array(chk�Զ���ҩ), mblnSetPara))
    mint�Զ���ҩʱ�� = Val(zldatabase.GetPara("�Զ���ҩʱ��", glngSys, 1341, 0, Array(Label1, txt��ҩʱ��, Label2), mblnSetPara))
    mstr�˲��� = zldatabase.GetPara("�˲���", glngSys, 1341, "", Array(lblCheck, cboCheck), mblnSetPara)
    
    If lngҩ��ID <> 0 Then
        Call SetDispense
    End If
    
    strTmp = zldatabase.GetPara("������", glngSys, 1341, "0", Array(Label4, cboҩƷ������ʾ), mblnSetPara)
    If InStr(1, strTmp, "|") > 0 Then
        mintShowName = Val(Mid(strTmp, 1, 1))
    Else
        mintShowName = Val(strTmp)
    End If
    If mintShowName > 2 Or mintShowName < 0 Then mintShowName = 0
    
    chk��С��λ.Value = Val(zldatabase.GetPara("��ʾ��С��λ", glngSys, 1341, 0, Array(chk��С��λ), mblnSetPara))
    chkAllType.Value = (zldatabase.GetPara("��ӡƱ�ݵ����и�ʽ", glngSys, 1341, 0, Array(chkAllType), mblnSetPara))
    chkSame.Value = (zldatabase.GetPara("����˲��˺���ҩ����ͬ", glngSys, 1341, 0, Array(chkSame), mblnSetPara))
    chkPreview.Value = zldatabase.GetPara("��ӡ����ǩʱ��Ԥ���ٴ�ӡ", glngSys, 1341, 0, Array(chkPreview), mblnSetPara)
    
    If intOverTime < 0 Or intOverTime > 1440 Then
        intOverTime = 0
    End If
    intOverTime = Int(intOverTime)
    chkOverTime.Value = IIf(intOverTime = 0, 0, 1)
    If chkOverTime.Value = 0 Then
        txtOverTime.Text = ""
        txtOverTime.Enabled = False
    Else
        txtOverTime.Text = intOverTime
        txtOverTime.Enabled = True
    End If
    
    '������ɫ�ʹ�ӡ������
    Call LoadList
    
    '�ŶӽкŲ��������Ŵ�ȡ
    With mType_Call
        .int�кŷ�ʽ = Val(zldatabase.GetPara("�кŷ�ʽ", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
        .int�����Ŷӽк� = Val(zldatabase.GetPara("�����Ŷӽк�", glngSys, 1341, 0, Array(chk�����Ŷӽк�), mblnSetPara, intParaType, lngҩ��ID))
        .int������������ = Val(zldatabase.GetPara("������������", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
'        .int��ʾģʽ = Val(zldatabase.GetPara("��ʾģʽ", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
        .int��ʾ�ŶӶ��� = Val(zldatabase.GetPara("��ʾ�ŶӶ���", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
        .int�������Ŵ��� = Val(zldatabase.GetPara("�������Ŵ���", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
        .int�����㲥���� = Val(zldatabase.GetPara("�����㲥����", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
        .int�������� = Val(zldatabase.GetPara("��������", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
        .strԶ�˺���վ�� = zldatabase.GetPara("Զ�˺���վ��", glngSys, 1341, "", Null, mblnSetPara, intParaType, lngҩ��ID)
        .int��ѯʱ�� = Val(zldatabase.GetPara("������ѯʱ��", glngSys, 1341, 10, Null, mblnSetPara, intParaType, lngҩ��ID))
    End With
End Function

Private Sub SetSourceDep()
    Dim rs As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select ���� || '-' || ���� ����, Id " & _
            " From ���ű� " & _
            " Where Id In (Select ����id From ��������˵�� Where �������� = '�ٴ�' And ������� In (1,2,3)) And " & _
            " (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By ���� || '-' || ���� "

    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, "SetSourceDep")
    Call SQLTest

    With rs
        If .EOF Then
            MsgBox "û�����ø��ಿ�ţ������Ź���", vbInformation, gstrSysName
            Exit Sub
        End If
        Lvw��Դ����.ListItems.Clear
        Do While Not .EOF
            Lvw��Դ����.ListItems.Add , "_" & !Id, !����, 1, 1
            If mstrSourceDep <> "" Then
                If InStr("," & mstrSourceDep & ",", "," & CStr(!Id) & ",") > 0 Then
                    Lvw��Դ����.ListItems("_" & !Id).Checked = True
                End If
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cboҩ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str�������� As String
    
    str�������� = "L,M,N"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Cboҩ��.ListCount = 0 Then Exit Sub
    
    If Cboҩ��.ListIndex >= 0 Then
        If Val(Cboҩ��.Tag) = Cboҩ��.ItemData(Cboҩ��.ListIndex) Then
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, Cboҩ��, Trim(Cboҩ��.Text), str��������, IIf(zlStr.IsHavePrivs(mstrPrivs, "����ҩ��"), False, True), "0,1,2,3") = False Then
        Exit Sub
    End If
    If Cboҩ��.ListIndex >= 0 Then
        Cboҩ��.Tag = Cboҩ��.ItemData(Cboҩ��.ListIndex)
    End If
End Sub

Private Sub Cboҩ��_KeyPress(KeyAscii As Integer)
    '�������뵥����
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Cboҩ��_Validate(Cancel As Boolean)
    If Cboҩ��.ListCount > 0 Then
        If Cboҩ��.ListIndex = -1 Then
            MsgBox "��ѡ��һ��ҩ�����ҩ����", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub chkIsDosage_Click()
    chkSign.Enabled = chkIsDosageOk.Value = 1 And chkIsDosage.Value = 1
    If chkSign.Enabled = False Then chkSign.Value = 0
    
    lblRefreshComment.Caption = IIf(chkIsDosage.Value = 0, "��������Ϣ�������", "����ҩ������������Ϣ��������Զ�ˢ��")
End Sub

Private Sub chkIsDosageOk_Click()
    chkSign.Enabled = chkIsDosageOk.Value = 1 And chkIsDosage.Value = 1
    If chkSign.Enabled = False Then chkSign.Value = 0
End Sub

Private Sub chkOverTime_Click()
    If chkOverTime.Value = 1 Then
        txtOverTime.Enabled = True
        If Int(Val(txtOverTime.Text)) = 0 Then
            txtOverTime.Text = "30"
        End If
    Else
        txtOverTime.Enabled = False
    End If
End Sub

Private Sub chkUseDisplay_Click()
    If Me.chkUseDisplay.Value = 0 Then
        frm��ʾ�豸����.Enabled = False
    Else
        frm��ʾ�豸����.Enabled = True
    End If
End Sub

Private Sub chkUseSound_Click()
    If Me.chkUseSound.Value = 1 Then
        frm�����㲥����.Enabled = True
        FraԶ����������.Enabled = True
        Me.optCallWay(0).Enabled = True
        Me.optCallWay(1).Enabled = True
    Else
        frm�����㲥����.Enabled = False
        FraԶ����������.Enabled = False
        Me.optCallWay(0).Enabled = False
        Me.optCallWay(1).Enabled = False
    End If
End Sub

Private Sub chk��ҩˢ��_Click()
    lst������.Enabled = (chk��ҩˢ��.Value = 1)
End Sub

Private Sub chk�Զ���ҩ_Click()
    If chk�Զ���ҩ.Value = 1 Then
        txt��ҩʱ��.Enabled = chk�Զ���ҩ.Enabled
    Else
        txt��ҩʱ��.Enabled = False
    End If
End Sub

Private Sub cmdDefaultPrinter_Click()
    Dim strDefault As String
    Dim n As Integer
    Dim i As Integer
    Dim rsData As ADODB.Recordset
    
    'ȡ����ĸ�ʽ���ƣ�Ĭ��ȡ��һ����ʽ��
    If mstrRPTDefaultScheme_Recipt = "" Then
        Set rsData = DeptSendWork_Get��ҩ����ʽ("ZL1_BILL_1341_3")
        If Not rsData.EOF Then mstrRPTDefaultScheme_Recipt = rsData!��ʽ
    End If
    
    '������ǰ�İ汾�����δӲ�ͬ��λ��ȡֵ
'    If mstrRPTDefaultScheme_Recipt <> "" Then strDefault = GetSetting("ZLSOFT", "˽��ģ��\zl9Report\LocalSet\ZL1_BILL_1341_3\" & mstrRPTDefaultScheme_Recipt, "Printer")
    If strDefault = "" Then strDefault = GetSetting("ZLSOFT", "˽��ģ��\zl9Report\LocalSet\ZL1_BILL_1341_3\���и�ʽ", "Printer")
    If strDefault = "" Then strDefault = GetSetting("ZLSOFT", "˽��ģ��\zl9Report\LocalSet\ZL1_BILL_1341_3", "Printer")
    If strDefault = "" Then strDefault = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\zl9Report\LocalSet\ZL1_BILL_1341_3", "Printer")
        
    If strDefault = "" Or InStr(1, ";" & mstrPrinters & ";", ";" & strDefault & ";") = 0 Then
        '���Ĭ�ϴ�ӡ��Ϊ�գ����߲��ڱ��ش�ӡ���б���ʱ
        MsgBox "û��������ҩ����ǩ��Ӧ�Ĵ�ӡ�������ڡ�Ʊ��(4)�������ã�", vbInformation, gstrSysName
        TabShow.Tab = 3
        Exit Sub
    Else
        '����Ĭ�ϵĴ�ӡ��
        For n = 0 To vsfPrinter.rows - 1
            vsfPrinter.TextMatrix(n, mintCol��ӡ��) = strDefault
        Next
        
    End If
End Sub

Private Sub cmdDeviceSetup_Click()
    Call FS.DeviceSetup(Me, 100, 1341)
End Sub

Private Sub cmdTestSound_Click()
    On Error GoTo errHandle
    If optSoundType(1).Value = True Then
        '΢������
        Call zlCall_MsSoundPlay("�롢" & "��־�ܡ�" & "��־�ܡ�" & "����һ�Ŵ���", Val(txtSpeed.Text))
    Else
        'ϵͳ����
        Call zlCall_SystemSoundPlay("�롢" & "��־�ܡ�" & "��־�ܡ�" & "����һ�Ŵ���", Val(txtSpeed.Text))
    End If
    Exit Sub
errHandle:
    Call SaveErrLog
End Sub
Private Sub cmd��ӡ����_Click()
    Dim strBill As String
    Select Case cboƱ������.ListIndex
    Case 0
        '��ҩ����ǩ
        strBill = "ZL1_BILL_1341_3"
    Case 1
        '��ҩ����ǩ
        strBill = "ZL1_BILL_1341_4"
    Case 2
        '������ҩ�嵥
        strBill = "ZL1_BILL_1341_2"
    Case 3
        '������ҩ֪ͨ��
        strBill = "ZL1_BILL_1341_1"
    Case 4
        '���ʴ���ͳ�Ʊ�
        strBill = "ZL1_INSIDE_1341"
    Case 5
        '��ҩҩƷ��ǩ
        strBill = "ZL1_BILL_1341_6"
    Case 6
        '�в�ҩҩƷ��ǩ
        strBill = "ZL1_BILL_1341_7"
    Case 7
        '���˷ѵ���
        strBill = "ZL1_BILL_1341_8"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub cmd��ʾ�豸����_Click()
    If gobjLEDShow Is Nothing Then
        If Not CreateObject_LED(Val(cbo��ʾӲ�����.ItemData(cbo��ʾӲ�����.ListIndex))) Then Exit Sub
    End If
        
    If Not gobjLEDShow Is Nothing Then
        Call gobjLEDShow.zlDrugSetup(Me, mstrWin)
    End If
End Sub

Private Sub lst��ӡ����_GotFocus()
    TabShow.Tab = 2
End Sub

Private Sub Cboҩ��_Click()
    Dim intDO As Integer
    Dim bln���� As Boolean, blnסԺ As Boolean
    Dim rstemp As New ADODB.Recordset
    Dim intParaType As Integer
    Dim n As Integer
    
    On Error GoTo errHandle
    
    '���ڼ��ع����е��µ�Click��ִ��
    If BlnStartUp = False Then Exit Sub
    
    '�����ܣ����û������ҩ���������涼������
    If Me.Cboҩ��.ListCount = 0 Then Exit Sub
    
    lngҩ��ID = Cboҩ��.ItemData(Cboҩ��.ListIndex)
    
    '���¶�ȡҩ����Ӧ�ķ�ҩ���ڼ�ҩ����Ա
    Call ReadWindowsAndPeople
    
    intUnit = Val(zldatabase.GetPara("ҩ������", glngSys, 1341))
    
    '������ҩ����
    SetDispense
    
    '����ҩ����ʾ��λ
    gstrSQL = " Select distinct ������� From ��������˵��" & _
              " Where ����ID=[1] And �������� like '%ҩ��'" & _
              " Order By ������� Desc"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡҩ���������]", Cboҩ��.ItemData(Cboҩ��.ListIndex))
    
    rstemp.Filter = "�������=3"
    If rstemp.RecordCount <> 0 Then bln���� = True: blnסԺ = True
    rstemp.Filter = "�������=2"
    If rstemp.RecordCount <> 0 Then blnסԺ = True
    rstemp.Filter = "�������=1"
    If rstemp.RecordCount <> 0 Then bln���� = True
    rstemp.Filter = 0
    
    With cbo��λ
        .Clear
        .AddItem "1-����Ӧ"
        .ItemData(.NewIndex) = 0
        If bln���� Then
            .AddItem "2-����ҩ��"
            .ItemData(.NewIndex) = 1
        End If
        If blnסԺ Then
            .AddItem "3-סԺҩ��"
            .ItemData(.NewIndex) = 2
        End If
        .ListIndex = 0
        
        For intDO = 0 To .ListCount - 1
            If .ItemData(intDO) = intUnit Then
                .ListIndex = intDO
                Exit For
            End If
        Next
    End With
    
    '������ȡ�ŶӽкŲ���
    With mType_Call
        .int�кŷ�ʽ = Val(zldatabase.GetPara("�кŷ�ʽ", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
        .int�����Ŷӽк� = Val(zldatabase.GetPara("�����Ŷӽк�", glngSys, 1341, 0, Array(chk�����Ŷӽк�), mblnSetPara, intParaType, lngҩ��ID))
        .int������������ = Val(zldatabase.GetPara("������������", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
'        .int��ʾģʽ = Val(zldatabase.GetPara("��ʾģʽ", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
        .int��ʾ�ŶӶ��� = Val(zldatabase.GetPara("��ʾ�ŶӶ���", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
        .int�������Ŵ��� = Val(zldatabase.GetPara("�������Ŵ���", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
        .int�����㲥���� = Val(zldatabase.GetPara("�����㲥����", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
        .int�������� = Val(zldatabase.GetPara("��������", glngSys, 1341, 0, Null, mblnSetPara, intParaType, lngҩ��ID))
        .strԶ�˺���վ�� = zldatabase.GetPara("Զ�˺���վ��", glngSys, 1341, "", Null, mblnSetPara, intParaType, lngҩ��ID)
        .int��ѯʱ�� = Val(zldatabase.GetPara("������ѯʱ��", glngSys, 1341, 10, Null, mblnSetPara, intParaType, lngҩ��ID))
    End With
    
   
    '���������Ŷӽкſؼ�״̬
    With mType_Call
        chk�����Ŷӽк�.Value = .int�����Ŷӽк�
        chkUseDisplay.Value = .int��ʾ�ŶӶ���
        chkUseSound.Value = .int������������
        
        If .int�кŷ�ʽ = 0 Then
            optCallWay(0).Value = True
        Else
            optCallWay(1).Value = True
        End If
        
        optSoundType(.int��������).Value = 1
        txtSpeed.Text = .int�����㲥����
        txtPlayCount.Text = .int�������Ŵ���
        Me.cboWorkStation.Text = .strԶ�˺���վ��
        txtLoopQueryTime.Text = .int��ѯʱ��
        
        chkUseDisplay_Click
        chkUseSound_Click
        
        If Me.optCallWay(0).Value = True Then
            optCallWay_Click 0
        Else
            optCallWay_Click 1
        End If
    End With


    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Chk��ӡ��ҩ��_Click()
    Dim ConState As Boolean
    
    ConState = (Chk��ӡ��ҩ��.Value = 1 And Chk��ӡ��ҩ��.Enabled = True)
    chkҩƷ��ǩ.Enabled = ConState
    If Not ConState And chkҩƷ��ǩ.Value = 1 Then
        chkҩƷ��ǩ.Value = 0
    End If
    
    chk���ķ��ϵ�.Enabled = ConState
    If Not ConState And chk���ķ��ϵ�.Value = 1 Then
        chk���ķ��ϵ�.Value = 0
    End If
    
    Opt��ӡ��ҩ��������.Enabled = ConState
    Opt��ӡ��ҩ��������.Enabled = ConState
    Opt��ӡ��ҩ��ѡ��.Enabled = ConState
    If Not ConState Then lst��ӡ����.Enabled = False
    
    If BlnStartUp = False Then Exit Sub
    
    If ConState Then
        If Opt��ӡ��ҩ��������.Enabled = True Then Opt��ӡ��ҩ��������.SetFocus
    End If
End Sub

Private Sub Chk��ӡ��ҩ��_GotFocus()
    TabShow.Tab = 2
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    Dim IntPrintStyle As Integer, i As Integer
    Dim strWin1 As String, strWin2 As String
    Dim intTemp As Integer
    Dim n As Integer
    Dim strPrinters As String
    Dim intSendCount As Integer
    Dim strCardType As String
    
    If Trim(txt��ѯ����.Text) = "" Then
        txt��ѯ����.Text = "1"
'        MsgBox "�������ѯ������1��-365�죩��", vbInformation, gstrSysName
'        txt��ѯ����.SetFocus
'        Exit Sub
    End If
    If Not IsNumeric(txt��ѯ����.Text) Then
        MsgBox "��ѯ�����к��зǷ��ַ���", vbInformation, gstrSysName
        If txt��ѯ����.Enabled = True Then txt��ѯ����.SetFocus
        Exit Sub
    End If
    If Val(txt��ѯ����.Text) < 1 Or Val(txt��ѯ����.Text) > 365 Then
        MsgBox "��ѯ��������С��1������365�죡", vbInformation, gstrSysName
        If txt��ѯ����.Enabled = True Then txt��ѯ����.SetFocus
        Exit Sub
    End If
    
    If Trim(Txtˢ�¼��) <> "" Then
        If Not IsNumeric(Txtˢ�¼��) Then
            MsgBox "ˢ�¼���к��зǷ��ַ���", vbInformation, gstrSysName
            If Txtˢ�¼��.Enabled = True Then Txtˢ�¼��.SetFocus
            Exit Sub
        End If
        If Val(Txtˢ�¼��) < 0 Or Val(Txtˢ�¼��) > 60 Then
            MsgBox "ˢ�¼��ֵ������Χ��0��60����", vbInformation, gstrSysName
            If Txtˢ�¼��.Enabled = True Then Txtˢ�¼��.SetFocus
            Exit Sub
        End If
        Txtˢ�¼�� = CInt(Txtˢ�¼��)
    End If
    If Trim(txt��ӡ���) <> "" Then
        If Not IsNumeric(txt��ӡ���) Then
            MsgBox "��ӡ����к��зǷ��ַ���", vbInformation, gstrSysName
            If txt��ӡ���.Enabled = True Then txt��ӡ���.SetFocus
            Exit Sub
        End If
        If Val(txt��ӡ���) < 0 Or Val(txt��ӡ���) > 60 Then
            MsgBox "��ӡ���ֵ������Χ��0��60����", vbInformation, gstrSysName
            If txt��ӡ���.Enabled = True Then txt��ӡ���.SetFocus
            Exit Sub
        End If
        txt��ӡ��� = CInt(txt��ӡ���)
    End If
    If Trim(Txt�ӳٴ�ӡ) <> "" Then
        If Not IsNumeric(Txt�ӳٴ�ӡ) Then
            MsgBox "�ӳٴ�ӡ�к��зǷ��ַ���", vbInformation, gstrSysName
            If Txt�ӳٴ�ӡ.Enabled = True Then Txt�ӳٴ�ӡ.SetFocus
            Exit Sub
        End If
        If Val(Txt�ӳٴ�ӡ) < 0 Or Val(Txt�ӳٴ�ӡ) > 60 Then
            MsgBox "�ӳٴ�ӡֵ������Χ��0��60����", vbInformation, gstrSysName
            If Txt�ӳٴ�ӡ.Enabled = True Then Txt�ӳٴ�ӡ.SetFocus
            Exit Sub
        End If
        Txt�ӳٴ�ӡ = CInt(Txt�ӳٴ�ӡ)
    End If
    If Trim(Txt��ӡ�˷ѵ���) <> "" Then
        If Not IsNumeric(Txt��ӡ�˷ѵ���) Then
            MsgBox "�˷ѵ����к��зǷ��ַ���", vbInformation, gstrSysName
            If Txt��ӡ�˷ѵ���.Enabled = True Then Txt��ӡ�˷ѵ���.SetFocus
            Exit Sub
        End If
        If Val(Txt��ӡ�˷ѵ���) < 0 Or Val(Txt��ӡ�˷ѵ���) > 60 Then
            MsgBox "��ӡ�˷ѵ�ֵ������Χ��0��60����", vbInformation, gstrSysName
            If Txt��ӡ�˷ѵ���.Enabled = True Then Txt��ӡ�˷ѵ���.SetFocus
            Exit Sub
        End If
        Txt��ӡ�˷ѵ��� = CInt(Txt��ӡ�˷ѵ���)
    End If
    
    '��鱾�����ܴ���:�����,����Ҫѡ��һ��
    For i = 0 To lst��ҩ����.ListCount - 1
        If lst��ҩ����.Selected(i) Then
            strWin1 = strWin1 & ",'" & lst��ҩ����.List(i) & "'"
            intSendCount = intSendCount + 1
        End If
    Next
    
    '��������Ŷӽкţ��򱾻�ֻ������һ����ҩ����
    If intSendCount > 1 And chk�����Ŷӽк�.Value = 1 Then
        MsgBox "�������Ŷӽкţ�ֻ������һ����ҩ���ڣ�", vbInformation, gstrSysName
        If lst��ҩ����.Enabled = True Then lst��ҩ����.SetFocus: Exit Sub
    End If
    
    If mblnLoadDrug And intSendCount > 1 Then
        MsgBox "�����������Զ���ҩ��ֻ������һ����ҩ���ڣ�", vbInformation, gstrSysName
        If lst��ҩ����.Enabled = True Then lst��ҩ����.SetFocus: Exit Sub
    End If
    
    strWin1 = Mid(strWin1, 2)
    If strWin1 = "" And lst��ҩ����.ListCount > 0 Then
        MsgBox "��ָ��������վ����Ӧ�ķ�ҩ���ڡ�", vbInformation, gstrSysName
        If lst��ҩ����.Enabled = True Then lst��ҩ����.SetFocus: Exit Sub
    End If
'    If UBound(Split(strWin1, ",")) + 1 = lst��ҩ����.ListCount Then strWin1 = ""
       
    
    '����ӡ��ҩ����:�����Ƿ���,����Ҫѡ��һ��
    For i = 0 To lst��ӡ����.ListCount - 1
        If lst��ӡ����.Selected(i) Then
            strWin2 = strWin2 & ",'" & lst��ӡ����.List(i) & "'"
        End If
    Next
    strWin2 = Mid(strWin2, 2)
    If strWin2 = "" And Chk��ӡ��ҩ��.Value = 1 And Opt��ӡ��ҩ��ѡ��.Value Then
        MsgBox "ѡ���ӡָ�����ڵ���ҩ��ʱ����Ҫ���ö�Ӧ�ķ�ҩ���ڣ�", vbInformation, gstrSysName
        If lst��ӡ����.Enabled = True Then lst��ӡ����.SetFocus: Exit Sub
    End If
    If UBound(Split(strWin2, ",")) + 1 = lst��ӡ����.ListCount Then strWin2 = ""
    
    '��Դ����
    mstrSourceDep = ""
    With Me.Lvw��Դ����
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked Then
                If mstrSourceDep = "" Then
                    mstrSourceDep = Mid(.ListItems(i).Key, 2)
                Else
                    mstrSourceDep = mstrSourceDep & "," & Mid(.ListItems(i).Key, 2)
                End If
            End If
        Next
    End With
    
    '������Ӧ�Ĵ�ӡ��
    With vsfPrinter
        intTemp = .rows / 6
        For n = 0 To .rows - 1
            If (n + 1) Mod intTemp = 1 And (n + 1) > intTemp Then strPrinters = strPrinters & ";"
            strPrinters = strPrinters & IIf(strPrinters = "" Or Right(strPrinters, 1) = ";", "", ",") & .TextMatrix(n, mintCol��ʽ) & "?" & .TextMatrix(n, mintCol��ӡ��)
        Next
    End With
        
    '����ˢ���Ŀ����
    If chk��ҩˢ��.Value = 1 Then
        If lst������.ListCount > 0 Then
            For i = 0 To lst������.ListCount - 1
                If lst������.Selected(i) Then
                    strCardType = IIf(strCardType = "", strCardType, strCardType & ",") & lst������.ItemData(i)
                End If
            Next
        End If
    End If
        
    On Error Resume Next
    
    '����˽�в���
    zldatabase.SetPara "������", Me.cboҩƷ������ʾ.ListIndex, glngSys, 1341

    zldatabase.SetPara "�շѴ�����ʾ��ʽ", cbo�շѴ���.ListIndex, glngSys, 1341
    zldatabase.SetPara "���ʴ�����ʾ��ʽ", cbo���ʴ���.ListIndex, glngSys, 1341
    zldatabase.SetPara "����ҩ���ݴ�ӡ��ʾ��ʽ", cbo����ҩ.ListIndex, glngSys, 1341
    zldatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, 1341
    zldatabase.SetPara "��ӡ�������ʵ�", IIf(chk���ʵ�.Value, 1, 0), glngSys, 1341
    zldatabase.SetPara "��ӡ�˷ѵ��ݼ��", Val(Txt��ӡ�˷ѵ���), glngSys, 1341
    zldatabase.SetPara "��ӡ�ӳ�", Val(Txt�ӳٴ�ӡ), glngSys, 1341
    zldatabase.SetPara "ˢ�¼��", Val(Txtˢ�¼��), glngSys, 1341
    zldatabase.SetPara "��ӡ���", Val(txt��ӡ���), glngSys, 1341
    
    zldatabase.SetPara "ҩ������", cbo��λ.ListIndex, glngSys, 1341
    zldatabase.SetPara "��ʾ����", Chk��ʾ������.Value, glngSys, 1341
    zldatabase.SetPara "��ҩ���Զ���ӡ", Me.Cbo��ҩ��.ListIndex, glngSys, 1341
    zldatabase.SetPara "��ҩ���Զ���ӡ", Me.cbo��ҩ��.ListIndex, glngSys, 1341
    zldatabase.SetPara "��ҩ���ӡҩƷ��ǩ", Me.cboҩƷ��ǩ.ListIndex, glngSys, 1341
    zldatabase.SetPara "��ʾ��С��λ", chk��С��λ.Value, glngSys, 1341
    zldatabase.SetPara "��ҩģʽɨ����ȷ��", chk��ҩɨ��.Value, glngSys, 1341
    zldatabase.SetPara "��ʱδ��ҩƷ��ʾʱ����", IIf(chkOverTime.Value = 0, 0, Int(Val(txtOverTime.Text))), glngSys, 1341
    zldatabase.SetPara "������סԺ����", Me.cbo��������.ListIndex, glngSys, 1341
    zldatabase.SetPara "����ˢ����ҩ", strCardType, glngSys, 1341
    zldatabase.SetPara "�����ʾ��ʽ", cbo�����ʾ.ListIndex, glngSys, 1341
    zldatabase.SetPara "��ҩ�������ķ������", chkCheckStuff.Value, glngSys, 1341
    zldatabase.SetPara "����ҩ����ɨ����Զ�����", chkɨ������.Value, glngSys, 1341
    zldatabase.SetPara "��ӡ����ǩʱ��Ԥ���ٴ�ӡ", chkPreview.Value, glngSys, 1341
    zldatabase.SetPara "����ʱϵͳ�Զ��س���ʽ", cbo�س���ʽ.ListIndex, glngSys, 1341
    
    If chkDispensing.Visible Then
        zldatabase.SetPara "����ʱ֪ͨ��ʼ��ҩ", Me.chkDispensing.Value, glngSys, 1341
    Else
        zldatabase.SetPara "����ʱ֪ͨ��ʼ��ҩ", "0", glngSys, 1341
    End If
     
    '��ӡ
    IntPrintStyle = Chk��ӡ��ҩ��.Value
    If IntPrintStyle = 1 Then IntPrintStyle = IIf(Opt��ӡ��ҩ��������.Value, 1, 1)
    If IntPrintStyle = 1 Then IntPrintStyle = IIf(Opt��ӡ��ҩ��������.Value, 2, 1)
    If IntPrintStyle = 1 Then IntPrintStyle = IIf(Opt��ӡ��ҩ��ѡ��.Value, 3, 1)
    zldatabase.SetPara "�����µ����Ƿ��ӡ", IntPrintStyle, glngSys, 1341
    zldatabase.SetPara "��ӡָ����ҩ����", strWin2, glngSys, 1341
    zldatabase.SetPara "��ӡҩƷ��ǩ", IIf(chkҩƷ��ǩ.Value, 1, 0), glngSys, 1341
    zldatabase.SetPara "��ӡ���ķ��ϵ�", IIf(chk���ķ��ϵ�.Value, 1, 0), glngSys, 1341
    zldatabase.SetPara "��ӡƱ�ݵ����и�ʽ", IIf(chkAllType.Value = 1, 1, 0), glngSys, 1341
            
    '��ҩ
    zldatabase.SetPara "��ҩҩ��", Cboҩ��.ItemData(Cboҩ��.ListIndex), glngSys, 1341
    zldatabase.SetPara "��ҩ����", strWin1, glngSys, 1341
    zldatabase.SetPara "��ҩ��", IIf(cbo��ҩ��.Text <> "��ǰ����Ա", cbo��ҩ��.Text, "|��ǰ����Ա|"), glngSys, 1341
    zldatabase.SetPara "�Զ���ҩ", IIf(chk�Զ���ҩ.Value = 1, 1, 0), glngSys, 1341
    zldatabase.SetPara "�Զ���ҩʱ��", Val(txt��ҩʱ��.Text), glngSys, 1341
    zldatabase.SetPara "ǩ��ʱ������ҩ", chkSign.Value, glngSys, 1341
    zldatabase.SetPara "�˲���", IIf(cboCheck.Text <> "��ǰ����Ա", cboCheck.Text, "|��ǰ����Ա|"), glngSys, 1341
    zldatabase.SetPara "����˲��˺���ҩ����ͬ", IIf(chkSame.Value = 1, 1, 0), glngSys, 1341
    
    '�����ŶӽкŵĲ���
    zldatabase.SetPara "�кŷ�ʽ", IIf(Me.optCallWay(0).Value = True, 0, 1), glngSys, 1341, mblnSetPara
    zldatabase.SetPara "�����Ŷӽк�", Me.chk�����Ŷӽк�.Value, glngSys, 1341, mblnSetPara, Cboҩ��.ItemData(Cboҩ��.ListIndex)
    zldatabase.SetPara "������������", Me.chkUseSound.Value, glngSys, 1341, mblnSetPara
    zldatabase.SetPara "��ʾ�ŶӶ���", chkUseDisplay.Value, glngSys, 1341, mblnSetPara
    zldatabase.SetPara "�������Ŵ���", Val(txtPlayCount.Text), glngSys, 1341, mblnSetPara
    zldatabase.SetPara "�����㲥����", Val(txtSpeed.Text), glngSys, 1341, mblnSetPara
    zldatabase.SetPara "��������", IIf(optSoundType(0).Value = True, 0, 1), glngSys, 1341, mblnSetPara
    zldatabase.SetPara "Զ�˺���վ��", Me.cboWorkStation.Text, glngSys, 1341, mblnSetPara, Cboҩ��.ItemData(Cboҩ��.ListIndex)
    zldatabase.SetPara "������ѯʱ��", Val(Me.txtLoopQueryTime.Text), glngSys, 1341, mblnSetPara
    zldatabase.SetPara "��ʾ�豸���", cbo��ʾӲ�����.ItemData(cbo��ʾӲ�����.ListIndex), glngSys, 1341, mblnSetPara
    
    
    '��Դ����
    zldatabase.SetPara "��Դ����", mstrSourceDep, glngSys, 1341
    
    '��ҩ��&����ǩ��ӡ��ʽ
    zldatabase.SetPara "��ҩ����ӡ��ʽ", cbo��ҩ��ҩ��ʽ.ItemData(cbo��ҩ��ҩ��ʽ.ListIndex) & ";" & cbo��ҩ��ҩ��ʽ.ItemData(cbo��ҩ��ҩ��ʽ.ListIndex), glngSys, 1341
    zldatabase.SetPara "����ǩ��ӡ��ʽ", cbo��ҩ������ʽ.ItemData(cbo��ҩ������ʽ.ListIndex) & ";" & cbo��ҩ������ʽ.ItemData(cbo��ҩ������ʽ.ListIndex), glngSys, 1341
    
    '������Ӧ�Ĵ�ӡ��
    zldatabase.SetPara "������Ӧ�Ĵ�ӡ��", strPrinters, glngSys, 1341
    
    frmҩƷ������ҩNew.BlnSetParaSuccess = True
    
    '������ҩ����ҩȷ�ϻ���
    gstrSQL = "Zl_ҩ����ҩ����_Update("
    gstrSQL = gstrSQL & Me.Cboҩ��.ItemData(Me.Cboҩ��.ListIndex)
    gstrSQL = gstrSQL & "," & Me.chkIsDosage.Value
    gstrSQL = gstrSQL & "," & Me.chkIsDosageOk.Value
    gstrSQL = gstrSQL & ")"
    
    Call zldatabase.ExecuteProcedure(gstrSQL, "cmdOK_Click")
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    BlnStartUp = False
    
    '��ʼ��chkDispensing
    Call InitDispensing
    
    '��ȡע���
    Call ReadFromReg
    '����������ʾ
    Call WriteCons
    '��Դ����
    Call SetSourceDep
    
    BlnStartUp = True
    RestoreWinState Me, App.ProductName
End Sub

Private Function ReadWindowsAndPeople()
    Dim intParaType As Integer
    
    '--��ȡ��ҩ���ķ�ҩ���ڼ���ҩ��--
    
    
        '��ҩ���ڣ�Ҫ��ӡ�ķ�ҩ�����������в�����"���з�ҩ����"��
'        If .State = 1 Then .Close
'        gstrSQL = " Select ���� From ��ҩ���� Where ҩ��ID=" & Cboҩ��.ItemData(Cboҩ��.ListIndex)
'        Call SQLTest(App.Title, Me.Caption, gstrSQL)
'        .Open gstrSQL, gcnOracle
'        Call SQLTest

    Dim lngLEDModal As Long
    
    On Error GoTo errHandle
    gstrSQL = " Select ���� From ��ҩ���� Where ҩ��ID=[1]"
    Set RecPeople = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Cboҩ��.ItemData(Cboҩ��.ListIndex))
    
    mstrWin = ""
    
    With RecPeople
        Me.lst��ҩ����.Clear
        Me.lst��ӡ����.Clear
        lst��ҩ����.Columns = 2 '��������Ч���ȽϺ�
        lst��ӡ����.Columns = 2
        Do While Not .EOF
            lst��ҩ����.AddItem !����
            lst��ӡ����.AddItem !����
            
            lst��ҩ����.Selected(lst��ҩ����.NewIndex) = True
            If Opt��ӡ��ҩ��ѡ��.Value Then
                lst��ӡ����.Selected(lst��ӡ����.NewIndex) = True
            End If
            
            mstrWin = IIf(mstrWin = "", "", mstrWin & ",") & !����
            
            .MoveNext
        Loop

        If lst��ҩ����.ListCount > 0 Then lst��ҩ����.ListIndex = 0
        If lst��ӡ����.ListCount > 0 Then lst��ӡ����.ListIndex = 0
    End With
    '��ҩ��
    gstrSQL = " Select ���� From ��Ա��  Where ID in " & _
             " (Select Distinct ��ԱID From ��Ա����˵�� Where ��Ա����='ҩ����ҩ��' " & _
             " And ��ԱID IN (Select ��ԱID From ������Ա Where ����ID=[1]))" & _
             " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
    Set RecPeople = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Cboҩ��.ItemData(Cboҩ��.ListIndex))
    
    With RecPeople
        Me.cbo��ҩ��.Clear
        Me.cbo��ҩ��.AddItem "��ǰ����Ա"
        Do While Not .EOF
            cbo��ҩ��.AddItem !����
            .MoveNext
        Loop
        cbo��ҩ��.ListIndex = 0
    End With
    
    With RecPeople
        If .RecordCount <> 0 Then
            .MoveFirst
        End If
        Me.cboCheck.Clear
        Me.cboCheck.AddItem "��ǰ����Ա"
        Do While Not .EOF
            cboCheck.AddItem !����
            .MoveNext
        Loop
        cboCheck.ListIndex = 0
    End With
    
    lngLEDModal = zldatabase.GetPara("��ʾ�豸���", glngSys, 1341, "101", Null, mblnSetPara, intParaType, lngҩ��ID)
    cbo��ʾӲ�����.Clear
    
    gstrSQL = "Select ��������,������,Nvl(����,0) AS ����,˵�� From �Ŷ�LED��ʾ����  "
    Set RecPeople = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��LED��ʾ�ӿڵ�ע����Ϣ")
    
    While RecPeople.EOF = False
        cbo��ʾӲ�����.AddItem zlStr.nvl(RecPeople!˵��)
        cbo��ʾӲ�����.ItemData(cbo��ʾӲ�����.ListCount - 1) = zlStr.nvl(RecPeople!��������, 0)
        If lngLEDModal = zlStr.nvl(RecPeople!��������, 0) Then
            cbo��ʾӲ�����.ListIndex = cbo��ʾӲ�����.ListCount - 1
        End If
        RecPeople.MoveNext
    Wend
    
    If cbo��ʾӲ�����.ListCount > 0 And cbo��ʾӲ�����.ListIndex = -1 Then
        cbo��ʾӲ�����.ListIndex = 0
    End If
    
    '���վ���б�
    ReadWorkStationInf
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteCons()
    Dim IntLocate As Integer
    Dim rsData As ADODB.Recordset
    
    '�����û�������ʾ
    
    RecPart.MoveFirst               '������Ϊ�գ������������涼�����ˣ�
    
    txt��ѯ����.Text = intDays
    'װ������������
    With Me.Cboҩ��
        Do While Not RecPart.EOF
            .AddItem RecPart!����
            .ItemData(.NewIndex) = RecPart!Id
            RecPart.MoveNext
        Loop
        .ListIndex = 0
    End With
    With Me.Cbo��ҩ��
        .AddItem "1-��ҩ����ʾ�Ƿ��ӡ"
        .AddItem "2-��ҩ���Զ���ӡ"
        .AddItem "3-��ҩ�󲻴�ӡ"
        .ListIndex = IntAutoPrint
    End With
    
    With Me.cbo��ҩ��
        .AddItem "1-��ҩ����ʾ�Ƿ��ӡ"
        .AddItem "2-��ҩ���Զ���ӡ"
        .AddItem "3-��ҩ�󲻴�ӡ"
        .ListIndex = mint��ҩ���Զ���ӡ
    End With
    
    With Me.cboҩƷ��ǩ
        .AddItem "1-��ҩ����ʾ�Ƿ��ӡ"
        .AddItem "2-��ҩ���Զ���ӡ"
        .AddItem "3-��ҩ�󲻴�ӡ"
        .ListIndex = mint��ҩ���Զ���ӡҩƷ��ǩ
    End With
    
    With cbo�շѴ���
        .Clear
        .AddItem "1-����ʾ�κδ���"
        .AddItem "2-��ʾδ�շѴ���"
        .AddItem "3-��ʾ���շѴ���"
        .AddItem "4-��ʾ���еĴ���"
        .ListIndex = 0
    End With
    With cbo���ʴ���
        .Clear
        .AddItem "1-����ʾ�κδ���"
        .AddItem "2-��ʾδ��˴���"
        .AddItem "3-��ʾ����˴���"
        .AddItem "4-��ʾ���еĴ���"
        .ListIndex = 0
    End With
    
    With cbo����ҩ
        .Clear
        .AddItem "0-��ʾ������ҩ��"
        .AddItem "1-��ʾδ��ӡ��ҩ��"
        .AddItem "2-��ʾ�Ѵ�ӡ��ҩ��"
        .ListIndex = 0
    End With
    
    With cboƱ������
        .Clear
        .AddItem "1-��ҩ����ǩ"
        .AddItem "2-��ҩ����ǩ"
        .AddItem "3-������ҩ�嵥"
        .AddItem "4-������ҩ֪ͨ��"
        .AddItem "5-���ʴ���ͳ�Ʊ�"
        .AddItem "6_��ҩҩƷ��ǩ"
        .AddItem "7_�в�ҩҩƷ��ǩ"
        .AddItem "8_���˷ѵ���"
        .ListIndex = 0
    End With
    
    With Me.cboҩƷ������ʾ
        .Clear
        .AddItem "0-��ʾҩƷ����������"
        .AddItem "1-����ʾҩƷ����"
        .AddItem "2-����ʾҩƷ����"
        .ListIndex = 0
    End With
    
    With Me.cbo�����ʾ
        .Clear
        .AddItem "0-��ʾӦ�ս��"
        .AddItem "1-��ʾʵ�ս��"
        .AddItem "2-��ʾӦ�պ�ʵ�ս��"
        .ListIndex = 0
    End With
    
    With Me.cbo��������
        .Clear
        .AddItem "0-��ʾ�����סԺ����"
        .AddItem "1-ֻ��ʾ���ﴦ��"
        .AddItem "2-ֻ��ʾסԺ����"
        .ListIndex = mintType
    End With
    
    With Me.cbo�س���ʽ
        .Clear
        .AddItem "0-ϵͳ���Զ��س�"
        .AddItem "1-��¼��ﵽ��Ŀ�򿨺ų���ʱ�Զ��س�"
    End With
    
    'װ���������
    cbo�շѴ���.ListIndex = mintShowBill�շ�
    cbo���ʴ���.ListIndex = mintShowBill����
    cbo����ҩ.ListIndex = mintShowBill��ҩ
    Chk��ʾ������.Value = IntShowCol
    
    Chk��ӡ��ҩ��.Value = IIf(intPrint = 0, 0, 1)
    
    cbo�����ʾ.ListIndex = mint�����ʾ��ʽ

    Opt��ӡ��ҩ��������.Value = IIf(intPrint = 1, True, False)
    Opt��ӡ��ҩ��������.Value = IIf(intPrint = 2, True, False)
    Opt��ӡ��ҩ��ѡ��.Value = IIf(intPrint = 3, True, False)
    
    Txtˢ�¼�� = Format(IntRefresh, "#####;-#####; ;")
    txt��ӡ��� = Format(mintPrintInterval, "#####;-#####; ;")
    Txt�ӳٴ�ӡ = Format(intPrintDelay, "#####;-#####; ;")
    Txt��ӡ�˷ѵ��� = Format(intPrintHandbackNO, "#####;-#####; ;")
    
    If txt��ӡ���.Enabled Then txt��ӡ���.Enabled = Not mblnUseMsg
    lblPrintComment.Visible = mblnUseMsg
    
    If Txtˢ�¼��.Enabled Then Txtˢ�¼��.Enabled = Not mblnUseMsg And chkIsDosage.Value = 0
    lblRefreshComment.Visible = mblnUseMsg
    lblRefreshComment.Caption = IIf(chkIsDosage.Value = 0, "��������Ϣ�������", "����ҩ������������Ϣ��������Զ�ˢ��")
    
    If lngҩ��ID <> 0 Then                                  '��λҩ��
        '�����ڸ�ҩ������ʾ
        For IntLocate = 0 To Me.Cboҩ��.ListCount - 1
            If Me.Cboҩ��.ItemData(IntLocate) = lngҩ��ID Then
                Me.Cboҩ��.ListIndex = IntLocate
                Exit For
            End If
        Next
        If IntLocate > (Cboҩ��.ListCount - 1) Then
            MsgBox "����������ҩ����ԭ�����õ�ҩ����ʧЧ����", vbInformation, gstrSysName
            If Cboҩ��.ListCount >= 1 Then Cboҩ��.ListIndex = 0
        End If
    End If
    BlnStartUp = True
    Cboҩ��_Click                                           '��������ҩ���񣬾���ȡ��ҩ��������ҩ���ڼ���ҩ��
    BlnStartUp = False
    
    '��λ��ҩ����
    If Str���� <> "" Then
        For IntLocate = 0 To lst��ҩ����.ListCount - 1
            If InStr(Str����, "'" & lst��ҩ����.List(IntLocate) & "'") > 0 Then
                lst��ҩ����.Selected(IntLocate) = True
            Else
                lst��ҩ����.Selected(IntLocate) = False
            End If
        Next
        If lst��ҩ����.ListCount > 0 Then lst��ҩ����.ListIndex = 0
    End If
    
    If str��ҩ�� <> "" Then                                 '��ʾ
        '�����ڸ���ҩ������ʾ
        If str��ҩ�� = "|��ǰ����Ա|" Then
            cbo��ҩ��.ListIndex = 0
        Else
            For IntLocate = 1 To cbo��ҩ��.ListCount - 1
                If cbo��ҩ��.List(IntLocate) = str��ҩ�� Then
                    cbo��ҩ��.ListIndex = IntLocate
                    Exit For
                End If
            Next
            If IntLocate > (cbo��ҩ��.ListCount - 1) Then
                MsgBox "������������ҩ�ˣ�ԭ�����õ���ҩ���Ѳ��ڱ����ţ���", vbInformation, gstrSysName
                If cbo��ҩ��.ListCount >= 1 Then cbo��ҩ��.ListIndex = 0
            End If
        End If
    End If
    
    If mstr�˲��� <> "" Then
        '�����ڸú˲�������ʾ
        If mstr�˲��� = "|��ǰ����Ա|" Then
            cboCheck.ListIndex = 0
        Else
            For IntLocate = 1 To cboCheck.ListCount - 1
                If cboCheck.List(IntLocate) = mstr�˲��� Then
                    cboCheck.ListIndex = IntLocate
                    Exit For
                End If
            Next
            If IntLocate > (cboCheck.ListCount - 1) Then
                MsgBox "���������ú˲��ˣ�ԭ�����õĺ˲����Ѳ��ڱ����ţ���", vbInformation, gstrSysName
                If cboCheck.ListCount >= 1 Then cboCheck.ListIndex = 0
            End If
        End If
    End If
    
    '��λ��ӡ��ҩ����
    If strPrintWindow <> "" Then
        For IntLocate = 0 To lst��ӡ����.ListCount - 1
            If InStr(strPrintWindow, "'" & lst��ӡ����.List(IntLocate) & "'") > 0 Then
                lst��ӡ����.Selected(IntLocate) = True
            Else
                lst��ӡ����.Selected(IntLocate) = False
            End If
        Next
        If lst��ӡ����.ListCount > 0 Then lst��ӡ����.ListIndex = 0
    End If
    
    Me.cboҩƷ������ʾ.ListIndex = mintShowName
    
    chk�Զ���ҩ.Value = IIf(mint�Զ���ҩ = 1, 1, 0)
    chk���ʵ�.Value = IIf(mint���ʵ� = 1, 1, 0)
    chkҩƷ��ǩ.Value = IIf(Chk��ӡ��ҩ��.Value = 1, IIf(mintҩƷ��ǩ = 1, 1, 0), 0)
    chk���ķ��ϵ�.Value = IIf(Chk��ӡ��ҩ��.Value = 1, IIf(mint���ķ��ϵ� = 1, 1, 0), 0)
    txt��ҩʱ��.Text = mint�Զ���ҩʱ��
    txt��ҩʱ��.Enabled = (mint�Զ���ҩ = 1 And chk�Զ���ҩ.Enabled = True)
    chk��ҩɨ��.Value = IIf(mint��ҩɨ�� = 1, 1, 0)
    chkSign.Value = IIf(mintSign = 1, 1, 0)
    Me.chkɨ������.Value = IIf(mintɨ������ = 1, 1, 0)
    
    If mint�س���ʽ >= 0 And mint�س���ʽ <= 1 Then
        cbo�س���ʽ.ListIndex = mint�س���ʽ
    Else
        cbo�س���ʽ.ListIndex = 0
    End If
    
    '�����ŶӽкŵĲ���
    With mType_Call
        chk�����Ŷӽк�.Value = .int�����Ŷӽк�
        chkUseDisplay.Value = .int��ʾ�ŶӶ���
        chkUseSound.Value = .int������������
        
        If .int�кŷ�ʽ = 0 Then
            optCallWay(0).Value = True
        Else
            optCallWay(1).Value = True
        End If
        
        optSoundType(.int��������).Value = 1
        txtSpeed.Text = .int�����㲥����
        txtPlayCount.Text = .int�������Ŵ���
        Me.cboWorkStation.Text = .strԶ�˺���վ��
        txtLoopQueryTime.Text = .int��ѯʱ��
    End With
    
    chkUseDisplay_Click
    chkUseSound_Click
    
    If Me.optCallWay(0).Value = True Then
        optCallWay_Click 0
    Else
        optCallWay_Click 1
    End If
    
    '����ˢ��ģʽ�Ϳ����
    chk��ҩˢ��.Value = IIf(mstr����ˢ����ҩ = "", 0, 1)
    lst������.Enabled = (chk��ҩˢ��.Value = 1)
    
    gstrSQL = "Select ID, ����, ���� From ҽ�ƿ���� Order By ����"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "WriteCons")
    If rsData.RecordCount > 0 Then
        lst������.Clear
        lst������.Columns = 2
        Do While Not rsData.EOF
            lst������.AddItem rsData!����
            lst������.ItemData(lst������.NewIndex) = rsData!Id
            
            If mstr����ˢ����ҩ <> "" Then
                If InStr(1, "," & mstr����ˢ����ҩ & ",", "," & rsData!Id & ",") > 0 Then
                    lst������.Selected(lst������.NewIndex) = True
                End If
            End If
            
            rsData.MoveNext
        Loop

        If lst������.ListCount > 0 Then lst������.ListIndex = 0
    Else
        chk��ҩˢ��.Enabled = False
        lst������.Enabled = False
    End If
    chkCheckStuff.Value = IIf(mint��ҩ���� = 1, 1, 0)
 End Function

Private Sub optCallWay_Click(index As Integer)
    If index = 0 Then
        FraԶ����������.Enabled = False
        frm�����㲥����.Enabled = True
    Else
        FraԶ����������.Enabled = True
        frm�����㲥����.Enabled = False
    End If
End Sub

Private Sub optSoundType_Click(index As Integer)
    If optSoundType(0).Value = True Then
        Label10.Caption = "�������٣�      (��Χ��0��100֮�䣬�Ƽ�65)"
        txtSpeed.Text = "65"
    Else
        Label10.Caption = "�������٣�      (��Χ��-10��10֮�䣬�Ƽ�-4)"
        txtSpeed.Text = "-4"
    End If
End Sub

Private Sub Opt��ӡ��ҩ��������_Click()
    lst��ӡ����.Enabled = False
End Sub

Private Sub Opt��ӡ��ҩ��������_GotFocus()
    TabShow.Tab = 2
End Sub

Private Sub Opt��ӡ��ҩ��������_Click()
    lst��ӡ����.Enabled = False
End Sub

Private Sub Opt��ӡ��ҩ��������_GotFocus()
    TabShow.Tab = 2
End Sub

Private Sub Opt��ӡ��ҩ��ѡ��_Click()
    lst��ӡ����.Enabled = Opt��ӡ��ҩ��ѡ��.Enabled
    If BlnStartUp = False Then Exit Sub
    
    If Opt��ӡ��ҩ��ѡ��.Value Then
        If lst��ӡ����.Enabled = True Then lst��ӡ����.SetFocus
    End If
End Sub
Private Sub Opt��ӡ��ҩ��ѡ��_GotFocus()
    TabShow.Tab = 2
End Sub



Private Sub tabShow_Click(PreviousTab As Integer)
    Select Case TabShow.Tab
    Case 0
        If Me.Cboҩ��.Enabled = True Then Me.Cboҩ��.SetFocus
    Case 2
        If Me.Chk��ӡ��ҩ��.Enabled = True Then Me.Chk��ӡ��ҩ��.SetFocus
    Case 3
        If Me.cboƱ������.Enabled = True Then Me.cboƱ������.SetFocus
    End Select
End Sub

Private Sub txtOverTime_Change()
    txtOverTime.Text = Int(Val(txtOverTime.Text))
    If Val(txtOverTime.Text) > 1440 Then
        txtOverTime.Text = "1440"
    End If
End Sub

Private Sub txtOverTime_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt��ӡ���_GotFocus()
    GetFocus txt��ӡ���
End Sub


Private Sub txt��ҩʱ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub


Private Sub Txt��ӡ�˷ѵ���_GotFocus()
    GetFocus Txt��ӡ�˷ѵ���
End Sub

Private Sub Txtˢ�¼��_GotFocus()
    GetFocus Txtˢ�¼��
End Sub

Private Sub Txt�ӳٴ�ӡ_GotFocus()
    GetFocus Txt�ӳٴ�ӡ
End Sub

Private Sub SetDispense()
'--------------------------------------
'������ҩ���Ƶ���ز���
'--------------------------------------
    Dim bln��ҩȷ�� As Boolean
    
    Me.chkIsDosage.Value = IIf(RecipeSendWork_DispensingMedi(Me.Cboҩ��.ItemData(Me.Cboҩ��.ListIndex), bln��ҩȷ��) = True, 1, 0)
    
    Me.chkIsDosageOk.Value = IIf(bln��ҩȷ�� = True, 1, 0)
End Sub

Private Sub ReadWorkStationInf()
'*****************************************************
'��ȡվ����Ϣ
'*****************************************************

    Dim strsql As String
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strsql = "select ����վ from zlClients where ��ֹʹ��<>1 order by ����վ"
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "��ȡվ����Ϣ")
    
    If rstemp.EOF Then Exit Sub
    
    While Not rstemp.EOF
        Call cboWorkStation.AddItem(rstemp("����վ"))
        rstemp.MoveNext
    Wend
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function NOCheck() As Boolean
    Dim strsql As String
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strsql = "select 1 from δ��ҩƷ��¼ where �ⷿid=[1] and (����=8 or ����=9 or ����=10)"
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "NOCheck", Val(Me.Cboҩ��.ItemData(Me.Cboҩ��.ListIndex)))
    
    If rstemp.EOF Then
        NOCheck = True
    Else
        NOCheck = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitDispensing()
'���ܣ���ʼ��chkDispensing�ؼ�

    Dim objMachine As Object
    
    err.Clear
    On Error Resume Next
    If Val(zldatabase.GetPara("����ҩƷ�Զ����豸�ӿ�", glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))) = 1 Then
        '�����½ӿ�
        Set objMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If err.Number <> 0 Then
            '��ξɽӿ�
            Set objMachine = CreateObject("zlDrugPacker.clsDrugPacker")
        End If
    Else
        '�ɽӿ�
        Set objMachine = CreateObject("zlDrugPacker.clsDrugPacker")
    End If
    On Error GoTo 0
    
    If objMachine Is Nothing Then
        'ҩƷ�Զ����豸�ӿڲ�����
        chkDispensing.Visible = False
        chkDispensing.Value = 0
    Else
        'ҩƷ�Զ����豸�ӿڴ���
        chkDispensing.Visible = True
        chkDispensing.Value = Val(zldatabase.GetPara("����ʱ֪ͨ��ʼ��ҩ", glngSys, 1341))
    End If
    
    lst������.Height = cbo�س���ʽ.Top - lst������.Top - 180
End Sub
