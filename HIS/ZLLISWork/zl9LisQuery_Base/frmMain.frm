VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "�ۺϲ�ѯ"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16755
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   16755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstSelect 
      Height          =   1680
      Left            =   6360
      TabIndex        =   96
      Top             =   8550
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Frame fraͳ�� 
      Height          =   900
      Left            =   60
      TabIndex        =   86
      Top             =   7785
      Visible         =   0   'False
      Width           =   5430
      Begin VB.CommandButton cmdCalc 
         Caption         =   "����(&J)"
         Height          =   350
         Left            =   4200
         TabIndex        =   95
         Top             =   135
         Width           =   1100
      End
      Begin VB.TextBox txtCount 
         Height          =   300
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   165
         Width           =   800
      End
      Begin VB.CheckBox chk˫�� 
         Caption         =   "˫���޳�����"
         Height          =   255
         Left            =   3300
         TabIndex        =   93
         Top             =   570
         Width           =   1380
      End
      Begin VB.TextBox txtSD 
         Height          =   300
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   165
         Width           =   800
      End
      Begin VB.TextBox txtDelSD 
         Height          =   300
         Left            =   615
         TabIndex        =   90
         Top             =   540
         Width           =   390
      End
      Begin VB.CommandButton cmd�޳� 
         Caption         =   "�޳�(&T)"
         Height          =   350
         Left            =   2145
         TabIndex        =   88
         Top             =   510
         Width           =   1100
      End
      Begin VB.TextBox txtAVG 
         Height          =   300
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   165
         Width           =   800
      End
      Begin VB.Label lbl��ֵSD 
         Caption         =   "ͳ������          ��ֵ           SD"
         Height          =   240
         Left            =   120
         TabIndex        =   91
         Top             =   225
         Width           =   3660
      End
      Begin VB.Label lbl�޳� 
         Caption         =   "�޳�>      SD������"
         Height          =   240
         Left            =   105
         TabIndex        =   89
         Top             =   585
         Width           =   1800
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "����(&R)"
      Height          =   900
      Index           =   4
      Left            =   11625
      Picture         =   "frmMain.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   7755
      Width           =   1500
   End
   Begin MSComctlLib.StatusBar stbBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   8685
      Width           =   16755
      _ExtentX        =   29554
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23715
            MinWidth        =   14111
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab ssTMain 
      Height          =   7725
      Left            =   345
      TabIndex        =   4
      Top             =   45
      Width           =   16365
      _ExtentX        =   28866
      _ExtentY        =   13626
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "�ճ�����(&R)"
      TabPicture(0)   =   "frmMain.frx":29FE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "pic(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "������ͳ��(&G)"
      TabPicture(1)   =   "frmMain.frx":2A1A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pic(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "���ͳ��(&T)"
      TabPicture(2)   =   "frmMain.frx":2A36
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "pic(2)"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   6540
         Index           =   0
         Left            =   225
         ScaleHeight     =   6540
         ScaleWidth      =   11955
         TabIndex        =   6
         Top             =   480
         Width           =   11955
         Begin VSFlex8Ctl.VSFlexGrid vfgData 
            Height          =   5700
            Index           =   0
            Left            =   30
            TabIndex        =   7
            Top             =   30
            Width           =   10905
            _cx             =   19235
            _cy             =   10054
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
            BackColorFixed  =   15790320
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16635590
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
            AllowUserFreezing=   1
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Frame fraData 
            Caption         =   "��������"
            Height          =   870
            Index           =   0
            Left            =   -30
            TabIndex        =   8
            Top             =   5880
            Width           =   11565
            Begin VB.ComboBox cbo���� 
               Height          =   300
               Index           =   0
               Left            =   4575
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   330
               Width           =   1500
            End
            Begin VB.ComboBox cboС�� 
               Height          =   300
               Index           =   0
               Left            =   6517
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   330
               Width           =   1800
            End
            Begin VB.ComboBox cbo���� 
               Height          =   300
               Index           =   0
               Left            =   8820
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   330
               Width           =   2600
            End
            Begin MSComCtl2.DTPicker dtpBegin 
               Height          =   300
               Index           =   0
               Left            =   1035
               TabIndex        =   12
               Top             =   330
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   39780355
               CurrentDate     =   40016
            End
            Begin MSComCtl2.DTPicker dtpEnd 
               Height          =   300
               Index           =   0
               Left            =   2640
               TabIndex        =   13
               Top             =   330
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   39780355
               CurrentDate     =   40016
            End
            Begin VB.Label lbl����ʱ�� 
               Caption         =   "����ʱ��                ��"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   17
               Top             =   390
               Width           =   2790
            End
            Begin VB.Label lbl���� 
               Caption         =   "����"
               Height          =   210
               Index           =   0
               Left            =   4110
               TabIndex        =   16
               Top             =   390
               Width           =   435
            End
            Begin VB.Label lblС�� 
               Caption         =   "С��"
               Height          =   210
               Index           =   0
               Left            =   6090
               TabIndex        =   15
               Top             =   390
               Width           =   435
            End
            Begin VB.Label lbl���� 
               Caption         =   "����"
               Height          =   210
               Index           =   0
               Left            =   8370
               TabIndex        =   14
               Top             =   390
               Width           =   435
            End
         End
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   7290
         Index           =   1
         Left            =   -74910
         ScaleHeight     =   7290
         ScaleWidth      =   12630
         TabIndex        =   18
         Top             =   390
         Width           =   12630
         Begin VB.Frame fraData 
            Caption         =   "��������"
            Height          =   1815
            Index           =   1
            Left            =   30
            TabIndex        =   21
            Top             =   5430
            Width           =   12555
            Begin VB.Frame fra�շ� 
               Caption         =   "�շ����"
               Height          =   1395
               Left            =   11415
               TabIndex        =   70
               Top             =   195
               Width           =   1050
               Begin VB.OptionButton opt�շ� 
                  Caption         =   "δ�շ�"
                  Height          =   180
                  Index           =   2
                  Left            =   90
                  TabIndex        =   73
                  Top             =   945
                  Width           =   855
               End
               Begin VB.OptionButton opt�շ� 
                  Caption         =   "���շ�"
                  Height          =   180
                  Index           =   1
                  Left            =   90
                  TabIndex        =   72
                  Top             =   630
                  Width           =   855
               End
               Begin VB.OptionButton opt�շ� 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   0
                  Left            =   90
                  TabIndex        =   71
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   675
               End
            End
            Begin VB.ComboBox cbo���� 
               Height          =   300
               Index           =   1
               Left            =   8805
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   300
               Width           =   2595
            End
            Begin VB.ComboBox cboС�� 
               Height          =   300
               Index           =   1
               Left            =   6495
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   300
               Width           =   1800
            End
            Begin VB.ComboBox cbo���� 
               Height          =   300
               Index           =   1
               Left            =   4560
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   300
               Width           =   1500
            End
            Begin VB.ComboBox cbo������� 
               Height          =   300
               Index           =   1
               Left            =   1020
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   675
               Width           =   1830
            End
            Begin VB.ComboBox cbo������ 
               Height          =   300
               Index           =   1
               Left            =   3495
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   675
               Width           =   1425
            End
            Begin VB.ComboBox cbo������ 
               Height          =   300
               Index           =   1
               Left            =   5565
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   675
               Width           =   1425
            End
            Begin VB.ComboBox cbo����� 
               Height          =   300
               Index           =   1
               Left            =   7635
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   675
               Width           =   1425
            End
            Begin VB.Frame fra������Դ 
               Caption         =   "������Դ"
               Height          =   1005
               Left            =   9105
               TabIndex        =   31
               Top             =   585
               Width           =   2295
               Begin VB.OptionButton opt��Դ 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   0
                  Left            =   120
                  TabIndex        =   36
                  Top             =   390
                  Value           =   -1  'True
                  Width           =   660
               End
               Begin VB.OptionButton opt��Դ 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   1
                  Left            =   810
                  TabIndex        =   35
                  Top             =   390
                  Width           =   660
               End
               Begin VB.OptionButton opt��Դ 
                  Caption         =   "סԺ"
                  Height          =   180
                  Index           =   2
                  Left            =   1500
                  TabIndex        =   34
                  Top             =   390
                  Width           =   660
               End
               Begin VB.OptionButton opt��Դ 
                  Caption         =   "Ժ��"
                  Height          =   180
                  Index           =   3
                  Left            =   120
                  TabIndex        =   33
                  Top             =   690
                  Width           =   660
               End
               Begin VB.OptionButton opt��Դ 
                  Caption         =   "���"
                  Height          =   180
                  Index           =   4
                  Left            =   825
                  TabIndex        =   32
                  Top             =   705
                  Width           =   660
               End
            End
            Begin VB.Frame frmͳ�Ʒ�ʽ 
               Caption         =   "ͳ�Ʒ�ʽ"
               Height          =   645
               Left            =   105
               TabIndex        =   22
               Top             =   945
               Width           =   8955
               Begin VB.OptionButton optͳ�Ʒ�ʽ 
                  Caption         =   "С��"
                  Height          =   180
                  Index           =   0
                  Left            =   180
                  TabIndex        =   30
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   1080
               End
               Begin VB.OptionButton optͳ�Ʒ�ʽ 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   1
                  Left            =   1260
                  TabIndex        =   29
                  Top             =   315
                  Width           =   1080
               End
               Begin VB.OptionButton optͳ�Ʒ�ʽ 
                  Caption         =   "������Ŀ"
                  Height          =   180
                  Index           =   2
                  Left            =   2325
                  TabIndex        =   28
                  Top             =   315
                  Width           =   1080
               End
               Begin VB.OptionButton optͳ�Ʒ�ʽ 
                  Caption         =   "�������"
                  Height          =   180
                  Index           =   3
                  Left            =   3405
                  TabIndex        =   27
                  Top             =   315
                  Width           =   1080
               End
               Begin VB.OptionButton optͳ�Ʒ�ʽ 
                  Caption         =   "������"
                  Height          =   180
                  Index           =   4
                  Left            =   4485
                  TabIndex        =   26
                  Top             =   315
                  Width           =   1080
               End
               Begin VB.OptionButton optͳ�Ʒ�ʽ 
                  Caption         =   "������"
                  Height          =   180
                  Index           =   5
                  Left            =   5565
                  TabIndex        =   25
                  Top             =   315
                  Width           =   1080
               End
               Begin VB.OptionButton optͳ�Ʒ�ʽ 
                  Caption         =   "�����"
                  Height          =   180
                  Index           =   6
                  Left            =   6630
                  TabIndex        =   24
                  Top             =   315
                  Width           =   1080
               End
               Begin VB.OptionButton optͳ�Ʒ�ʽ 
                  Caption         =   "������Դ"
                  Height          =   180
                  Index           =   7
                  Left            =   7710
                  TabIndex        =   23
                  Top             =   315
                  Width           =   1080
               End
            End
            Begin MSComCtl2.DTPicker dtpBegin 
               Height          =   300
               Index           =   1
               Left            =   1020
               TabIndex        =   44
               Top             =   300
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   39780355
               CurrentDate     =   40016
            End
            Begin MSComCtl2.DTPicker dtpEnd 
               Height          =   300
               Index           =   1
               Left            =   2625
               TabIndex        =   45
               Top             =   300
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   39780355
               CurrentDate     =   40016
            End
            Begin VB.Label lbl���� 
               Caption         =   "����"
               Height          =   210
               Index           =   1
               Left            =   8355
               TabIndex        =   53
               Top             =   360
               Width           =   435
            End
            Begin VB.Label lblС�� 
               Caption         =   "С��"
               Height          =   210
               Index           =   1
               Left            =   6075
               TabIndex        =   52
               Top             =   360
               Width           =   435
            End
            Begin VB.Label lbl���� 
               Caption         =   "����"
               Height          =   210
               Index           =   1
               Left            =   4050
               TabIndex        =   51
               Top             =   360
               Width           =   435
            End
            Begin VB.Label lbl����ʱ�� 
               Caption         =   "����ʱ��                ��"
               Height          =   255
               Index           =   1
               Left            =   225
               TabIndex        =   50
               Top             =   360
               Width           =   2790
            End
            Begin VB.Label lbl������� 
               Caption         =   "�������"
               Height          =   225
               Left            =   225
               TabIndex        =   49
               Top             =   690
               Width           =   780
            End
            Begin VB.Label Label1 
               Caption         =   "������"
               Height          =   225
               Left            =   2925
               TabIndex        =   48
               Top             =   720
               Width           =   780
            End
            Begin VB.Label Label2 
               Caption         =   "������"
               Height          =   225
               Left            =   4980
               TabIndex        =   47
               Top             =   720
               Width           =   780
            End
            Begin VB.Label Label3 
               Caption         =   "�����"
               Height          =   225
               Left            =   7050
               TabIndex        =   46
               Top             =   720
               Width           =   780
            End
         End
         Begin VB.Frame fraLR 
            Height          =   1875
            Index           =   1
            Left            =   4605
            MousePointer    =   9  'Size W E
            TabIndex        =   19
            Top             =   15
            Width           =   45
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgData 
            Height          =   5250
            Index           =   1
            Left            =   0
            TabIndex        =   20
            Top             =   90
            Width           =   4530
            _cx             =   7990
            _cy             =   9260
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
            BackColorFixed  =   15790320
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16635590
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
         Begin VSFlex8Ctl.VSFlexGrid vfgItem 
            Height          =   5250
            Index           =   1
            Left            =   4755
            TabIndex        =   54
            Top             =   75
            Width           =   7470
            _cx             =   13176
            _cy             =   9260
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
            BackColorFixed  =   15790320
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16635590
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   7260
         Index           =   2
         Left            =   -74490
         ScaleHeight     =   7260
         ScaleWidth      =   17070
         TabIndex        =   55
         Top             =   450
         Width           =   17070
         Begin VB.Frame fraData 
            Caption         =   "��������"
            Height          =   1350
            Index           =   2
            Left            =   840
            TabIndex        =   57
            Top             =   5865
            Width           =   16065
            Begin VB.TextBox txt���� 
               Height          =   300
               Left            =   11670
               TabIndex        =   85
               Top             =   765
               Width           =   1200
            End
            Begin VB.ComboBox cbo���� 
               Height          =   300
               Left            =   9150
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   765
               Width           =   1185
            End
            Begin VB.TextBox txt���� 
               Height          =   300
               Left            =   10395
               TabIndex        =   83
               Top             =   765
               Width           =   1200
            End
            Begin VB.CommandButton cmd��Ŀ 
               Caption         =   "��"
               Height          =   300
               Left            =   12620
               TabIndex        =   81
               Top             =   330
               Width           =   250
            End
            Begin VB.TextBox txt��Ŀ 
               Height          =   300
               Left            =   9135
               TabIndex        =   79
               Top             =   330
               Width           =   3465
            End
            Begin VB.ComboBox cbo�Ա� 
               Height          =   300
               Left            =   7530
               Style           =   2  'Dropdown List
               TabIndex        =   78
               Top             =   765
               Width           =   765
            End
            Begin VB.ComboBox cbo���� 
               Height          =   300
               Left            =   6045
               Style           =   2  'Dropdown List
               TabIndex        =   76
               Top             =   765
               Width           =   765
            End
            Begin VB.TextBox txt���� 
               Height          =   300
               Left            =   4665
               TabIndex        =   74
               ToolTipText     =   "֧������20-30�ķ�ʽָ����Χ"
               Top             =   765
               Width           =   1320
            End
            Begin VB.ComboBox cbo���� 
               Height          =   300
               Index           =   2
               Left            =   675
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   765
               Width           =   2985
            End
            Begin VB.ComboBox cboС�� 
               Height          =   300
               Index           =   2
               Left            =   6517
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   330
               Width           =   1800
            End
            Begin VB.ComboBox cbo���� 
               Height          =   300
               Index           =   2
               Left            =   4575
               Style           =   2  'Dropdown List
               TabIndex        =   58
               Top             =   330
               Width           =   1500
            End
            Begin MSComCtl2.DTPicker dtpBegin 
               Height          =   300
               Index           =   2
               Left            =   1035
               TabIndex        =   61
               Top             =   330
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   39780355
               CurrentDate     =   40016
            End
            Begin MSComCtl2.DTPicker dtpEnd 
               Height          =   300
               Index           =   2
               Left            =   2640
               TabIndex        =   62
               Top             =   330
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   39780355
               CurrentDate     =   40016
            End
            Begin VB.Label lbl�����Χ 
               Caption         =   "�����Χ"
               Height          =   225
               Left            =   8385
               TabIndex        =   82
               Top             =   810
               Width           =   900
            End
            Begin VB.Label lbl������Ŀ 
               Caption         =   "������Ŀ"
               Height          =   225
               Left            =   8370
               TabIndex        =   80
               Top             =   390
               Width           =   900
            End
            Begin VB.Label lbl�Ա� 
               Caption         =   "�Ա�"
               Height          =   225
               Left            =   6990
               TabIndex        =   77
               Top             =   810
               Width           =   450
            End
            Begin VB.Label lbl���� 
               Caption         =   "���䷶Χ"
               Height          =   225
               Left            =   3810
               TabIndex        =   75
               Top             =   810
               Width           =   795
            End
            Begin VB.Label lbl���� 
               Caption         =   "����"
               Height          =   210
               Index           =   2
               Left            =   225
               TabIndex        =   66
               Top             =   810
               Width           =   435
            End
            Begin VB.Label lblС�� 
               Caption         =   "С��"
               Height          =   210
               Index           =   2
               Left            =   6090
               TabIndex        =   65
               Top             =   390
               Width           =   435
            End
            Begin VB.Label lbl���� 
               Caption         =   "����"
               Height          =   210
               Index           =   2
               Left            =   4110
               TabIndex        =   64
               Top             =   390
               Width           =   435
            End
            Begin VB.Label lbl����ʱ�� 
               Caption         =   "����ʱ��                ��"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   63
               Top             =   390
               Width           =   2790
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgItem 
            Height          =   5250
            Index           =   2
            Left            =   5430
            TabIndex        =   68
            Top             =   195
            Width           =   6405
            _cx             =   11298
            _cy             =   9260
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
            BackColorFixed  =   15790320
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16635590
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
         Begin VSFlex8Ctl.VSFlexGrid vfgData 
            Height          =   5250
            Index           =   2
            Left            =   0
            TabIndex        =   56
            Top             =   540
            Width           =   5070
            _cx             =   8943
            _cy             =   9260
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
            BackColorFixed  =   15790320
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16635590
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
         Begin VB.Frame fraLR 
            Height          =   1875
            Index           =   2
            Left            =   4800
            MousePointer    =   9  'Size W E
            TabIndex        =   67
            Top             =   210
            Width           =   45
         End
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "�����Excel(&E)"
      Height          =   900
      Index           =   3
      Left            =   10125
      Picture         =   "frmMain.frx":2A52
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7755
      Width           =   1500
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "��ӡ����(&S)"
      Height          =   900
      Index           =   2
      Left            =   8655
      Picture         =   "frmMain.frx":5444
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "ҳ������"
      Top             =   7755
      Width           =   1500
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "��ӡ(&P)"
      Height          =   900
      Index           =   1
      Left            =   7050
      Picture         =   "frmMain.frx":7E36
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7755
      Width           =   1500
   End
   Begin VB.CommandButton cmdRun 
      Appearance      =   0  'Flat
      Caption         =   "����(&F)"
      Height          =   900
      Index           =   0
      Left            =   5475
      Picture         =   "frmMain.frx":A828
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "����"
      Top             =   7755
      Width           =   1500
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsHost As zl9LisQuery_Def.clsLisQueryHost
Private mlgIndex As Long
Private mintLastTab As Integer
Private Enum mCol_�ճ�
    ���� = 1:  С��: ����: ����: �ѽ���: �Ѻ���: ���: δ��
End Enum

Private mrs��Ŀ As ADODB.Recordset

Private Sub cbo����_Click()
    If cbo����.List(cbo����.ListIndex) = "��...֮��" Then
        txt����.Enabled = True
    Else
        txt����.Enabled = False
    End If
End Sub

Private Sub cmdCalc_Click()
    Call CalcData
End Sub

Private Sub cmd�޳�_Click()
    Dim lngRow As Long
    Dim strDeleteRow As String, varDelRow As Variant
    If txtDelSD.Text <> "" Then
        If Val(txtDelSD.Text) >= 1 And Val(txtDelSD.Text) <= 4 Then
            
            With vfgData(2)
                strDeleteRow = ""
                'lngCurrRow = .Row
                For lngRow = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(lngRow, .ColIndex("SD"))) = ">" & Val(txtDelSD.Text) & "S" Then
                        strDeleteRow = lngRow & "," & strDeleteRow
                    End If
                Next
                
                If strDeleteRow <> "" Then
                    varDelRow = Split(strDeleteRow, ",")
                    For lngRow = LBound(varDelRow) To UBound(varDelRow) - 1
                       .RemoveItem Val(varDelRow(lngRow))
                    Next
                End If
                
                'If lngCurrRow >= .FixedRows And lngCurrRow < .Rows Then
                Call vfgData_RowColChange(2)
                'End If
                
            End With
            Me.cmdCalc.Enabled = True
        Else
            MsgBox "����ķ�Χ��1-4�����飡", vbInformation, Me.Caption
        End If
    End If
End Sub

Private Sub cmd��Ŀ_Click()
    Call ShowSelect("")
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errH
    '--- ���¿�ʼ���ڣ����½�������
    For i = Me.dtpBegin.LBound To Me.dtpBegin.UBound
        Me.dtpBegin(i).Value = Format(Now, "yyyy-MM-01")
    Next
    For i = Me.dtpEnd.LBound To Me.dtpEnd.UBound
        Me.dtpEnd(i).Value = Format(Now, "yyyy-MM-dd")
    Next
    
    '--- ��ʼ������
    strSQL = "Select ���� From ���Ƽ������� Order By ����"
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cbo����.LBound To cbo����.UBound
        cbo����(i).Clear
        cbo����(i).AddItem ""
        
        Do Until rsTmp.EOF
            cbo����(i).AddItem "" & rsTmp.Fields("����")
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cbo����(i).ListIndex = 0
    Next
    
    '--- ��ʼ��С��
    strSQL = "Select a.Id, a.����, a.����" & vbNewLine & _
        "From ����С�� a, ����С���Ա b, �ϻ���Ա�� c" & vbNewLine & _
        "Where a.Id = b.С��id And b.��Աid = c.��Աid And �û��� = User" & vbNewLine & _
        "Order By a.����"

    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cboС��.LBound To cboС��.UBound
        cboС��(i).Clear
        cboС��(i).AddItem ""
        Do Until rsTmp.EOF
            cboС��(i).AddItem "" & rsTmp.Fields("����") & "-" & rsTmp.Fields("����")
            cboС��(i).ItemData(cboС��(i).NewIndex) = Val("" & rsTmp.Fields("ID"))
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cboС��(i).ListIndex = 0
    Next
    
    '--- ��ʼ������
    strSQL = "Select e.Id, e.����, e.����" & vbNewLine & _
        "From ����С�� a, ����С���Ա b, �ϻ���Ա�� c, ����С������ d, �������� e" & vbNewLine & _
        "Where Nvl(e.΢����,0) = 0 and a.Id = b.С��id And b.��Աid = c.��Աid And �û��� = User And a.Id = d.С��id And d.����id = e.Id" & vbNewLine & _
        "Order By e.����"
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cbo����.LBound To cbo����.UBound
        cbo����(i).Clear
        cbo����(i).AddItem ""
       
        Do Until rsTmp.EOF
            cbo����(i).AddItem "" & rsTmp.Fields("����") & "-" & rsTmp.Fields("����")
            cbo����(i).ItemData(cbo����(i).NewIndex) = Val("" & rsTmp.Fields("ID"))
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cbo����(i).ListIndex = 0
    Next
    '--- ��ʼ���������
    strSQL = "Select a.Id, a.����, a.����" & vbNewLine & _
            "From ��������˵�� b, ���ű� a" & vbNewLine & _
            "Where a.Id = b.����id And (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And" & vbNewLine & _
            "           Instr(',�ٴ�,���,', ',' || b.�������� || ',') > 0" & vbNewLine & _
            " Order by a.����"

    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cbo�������.LBound To cbo�������.UBound
        cbo�������(i).Clear
        cbo�������(i).AddItem ""
        Do Until rsTmp.EOF
            cbo�������(i).AddItem "" & rsTmp.Fields("����") & "-" & rsTmp.Fields("����")
            cbo�������(i).ItemData(cbo�������(i).NewIndex) = Val("" & rsTmp.Fields("ID"))
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cbo�������(i).ListIndex = 0
    Next
    
    '--- ��ʼ��������
    strSQL = "Select a.����" & vbNewLine & _
        "From ��Ա����˵�� b, ��Ա�� a" & vbNewLine & _
        "Where b.��Ա���� = 'ҽ��' And a.Id = b.��Աid And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd'))" & vbNewLine & _
        "Order By a.����"
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cbo������.LBound To cbo������.UBound
        cbo������(i).Clear
        cbo������(i).AddItem ""
        Do Until rsTmp.EOF
            cbo������(i).AddItem "" & rsTmp.Fields("����")
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cbo������(i).ListIndex = 0
    Next
    '--- ��ʼ�������, ������
    strSQL = "Select  Distinct B.����" & vbNewLine & _
        "From ��Ա�� B,����С���Ա A" & vbNewLine & _
        "Where A.��Աid=B.ID" & vbNewLine & _
        "Order By B.����"
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cbo�����.LBound To cbo�����.UBound
        cbo�����(i).Clear
        cbo�����(i).AddItem ""
        Do Until rsTmp.EOF
            cbo�����(i).AddItem "" & rsTmp.Fields("����")
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cbo�����(i).ListIndex = 0
    Next
    
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption)
    For i = cbo������.LBound To cbo������.UBound
        cbo������(i).Clear
        cbo������(i).AddItem ""
        Do Until rsTmp.EOF
            cbo������(i).AddItem "" & rsTmp.Fields("����")
            rsTmp.MoveNext
        Loop
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        cbo������(i).ListIndex = 0
    Next
    '����
    cbo����.Clear
    cbo����.AddItem "��"
    cbo����.AddItem "��"
    cbo����.AddItem "��"
    cbo����.AddItem "Сʱ"
    cbo����.AddItem "����"
    cbo����.AddItem "Ӥ��"
    cbo����.ListIndex = 0
    
    '�Ա�
    cbo�Ա�.Clear
    cbo�Ա�.AddItem ""
    cbo�Ա�.AddItem "1-��"
    cbo�Ա�.AddItem "2-Ů"
    cbo�Ա�.AddItem "3-δ֪"
    cbo�Ա�.AddItem "9-����"
    cbo�Ա�.ListIndex = 0
    
    '
    cbo����.Clear
    cbo����.AddItem "="
    cbo����.AddItem "<>"
    cbo����.AddItem ">"
    cbo����.AddItem "<"
    cbo����.AddItem ">="
    cbo����.AddItem "<="
    cbo����.AddItem "����"
    cbo����.AddItem "��...֮��"
    cbo����.ListIndex = 0
    '--- ��ʼ��
    
    
    Call initvfgDataTitle(0)
    Call initvfgDataTitle(1): Call initvfgItemTitle(1)
    Call initvfgDataTitle(2): Call initvfgItemTitle(2)
    mintLastTab = 0
    ssTMain.Tab = 0
    Me.Show
    Exit Sub
errH:
    MsgBox Err.Description
End Sub

Private Sub cmdRun_Click(Index As Integer)
    Dim strFileName As String
    Select Case Index
        Case 0 '����
            Me.cmdRun(0).Enabled = False
            Call DoQuery
            Me.cmdRun(0).Enabled = True
        Case 1 '��ӡ
            Me.cmdRun(1).Enabled = False
            If Not vsPrint Is Nothing Then Unload vsPrint
            Call vsPrint.vsPrint(Me.vfgData(Me.ssTMain.Tab).hWnd, Me.ssTMain.Tab)
            Me.cmdRun(1).Enabled = True
        Case 2 '��ӡ����
            Me.cmdRun(2).Enabled = False
            Call frmPrintSet.PageSetup(Me.ssTMain.Tab)
            Call DoQuery
            Me.cmdRun(2).Enabled = True
        Case 3 'Excel
            Me.cmdRun(3).Enabled = False
                strFileName = App.Path & "\Report" & Me.ssTMain.Tab & "_" & Format(Now, "yyyyMMddHHmmss") & ".xls"
                vfgData(Me.ssTMain.Tab).SaveGrid strFileName, flexFileExcel, flexXLSaveFixedCells
                MsgBox "�ѱ��浽" & strFileName, vbInformation, Me.Caption
            Me.cmdRun(3).Enabled = True
        Case 4 '����
            Unload Me
    End Select
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    Dim iTab As Integer
    On Error Resume Next
    With Me.ssTMain
        .Left = 45
        .Top = 45
        .Width = Me.ScaleWidth - 90
        .Height = Me.ScaleHeight - 90 - Me.cmdRun(0).Height - 45 - Me.stbBar.Height
    End With
    For i = pic.LBound To pic.UBound
        With Me.pic(i)
            .Left = Me.ssTMain.Left + 45
            .Top = Me.ssTMain.Top + 350
            .Width = Me.ssTMain.Width - 150
            .Height = Me.ssTMain.Height - 450
        End With
    Next
    
    With Me.cmdRun(0)
        .Left = Me.ScaleWidth - Me.cmdRun(0).Width * Me.cmdRun.Count - 90 * Me.cmdRun.Count
        .Top = Me.ssTMain.Top + Me.ssTMain.Height + 45
    End With
    
    For i = Me.cmdRun.LBound + 1 To Me.cmdRun.UBound
        Me.cmdRun(i).Left = Me.cmdRun(i - 1).Left + Me.cmdRun(i - 1).Width + 90
        Me.cmdRun(i).Top = Me.cmdRun(0).Top
    Next
    
    With fraͳ��
        .Left = Me.ssTMain.Left
        .Top = Me.cmdRun(0).Top
    End With
    
    If Me.lstSelect.Visible = True Then
        Call MoveSelect(txt��Ŀ)
    End If
    Me.Refresh
End Sub

Private Sub lstSelect_DblClick()

    If lstSelect.ListIndex >= 0 Then
        txt��Ŀ.Text = lstSelect.List(lstSelect.ListIndex)
        txt��Ŀ.Tag = lstSelect.ItemData(lstSelect.ListIndex)
    End If
    txt��Ŀ.SetFocus
    lstSelect.Visible = False
End Sub

Private Sub lstSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If lstSelect.ListIndex >= 0 Then
            txt��Ŀ.Text = lstSelect.List(lstSelect.ListIndex)
            txt��Ŀ.Tag = lstSelect.ItemData(lstSelect.ListIndex)
        End If
        txt��Ŀ.SetFocus
        lstSelect.Visible = False
    ElseIf KeyAscii = vbKeyEscape Then
        txt��Ŀ.SetFocus
        lstSelect.Visible = False
    End If
End Sub

Private Sub lstSelect_LostFocus()
    Me.lstSelect.Visible = False
End Sub

Private Sub pic_Resize(Index As Integer)
    On Error Resume Next
    '---- �ճ�

    With Me.vfgData(Index)
        .Left = Me.pic(Index).ScaleLeft
        .Top = Me.pic(Index).ScaleTop
        .Width = Me.pic(Index).ScaleWidth
        .Height = Me.pic(Index).Height - Me.fraData(Index).Height
    End With
    With Me.fraData(Index)
        .Left = Me.vfgData(Index).Left
        .Top = Me.vfgData(Index).Top + Me.vfgData(Index).Height
        .Width = Me.vfgData(Index).Width
    End With
    '--- ������
    If Index >= Me.vfgItem.LBound And Index <= Me.vfgItem.UBound Then
        With Me.vfgItem(Index)
             Me.vfgData(Index).Width = Me.vfgData(Index).Width - .Width - 45
            
            .Left = Me.vfgData(Index).Left + Me.vfgData(Index).Width + 45
            .Top = Me.vfgData(Index).Top
            .Height = Me.vfgData(Index).Height
            
            Me.fraLR(Index).Left = .Left - 45
            Me.fraLR(Index).Top = .Top
            Me.fraLR(Index).Height = .Height
        End With
    End If
End Sub

Private Sub ssTMain_Click(PreviousTab As Integer)
    Dim iTab As Integer
    iTab = PreviousTab
    If iTab >= pic.LBound And iTab <= pic.UBound Then
        pic(iTab).Visible = False
    End If
    
    iTab = ssTMain.Tab
    mintLastTab = iTab
    If iTab >= pic.LBound And iTab <= pic.UBound Then
        pic(iTab).Visible = True
    End If
    ssTMain.Tab = iTab
    fraͳ��.Visible = False
    If iTab = 2 Then fraͳ��.Visible = True
End Sub

Private Sub fraLR_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     On Error Resume Next
    If Button = vbLeftButton Then
        If Index >= Me.vfgItem.LBound And Index <= Me.vfgItem.UBound Then
            Me.vfgData(Index).Width = Me.vfgData(Index).Width + X
            Me.vfgItem(Index).Width = Me.vfgItem(Index).Width - X
            Me.fraLR(Index).Left = Me.fraLR(Index).Left + X
            Me.vfgItem(Index).Left = Me.fraLR(Index).Left + Me.fraLR(Index).Width
        End If
    End If
End Sub

Private Sub txt��Ŀ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(txt��Ŀ.Text) <> "" Then
            Call ShowSelect(Trim(txt��Ŀ.Text))
        End If
    End If
End Sub

Private Sub vfgData_DblClick(Index As Integer)
    Dim lngRow As Long
    
    If Index = 2 Then
        If chk˫��.Value = 1 Then
            lngRow = vfgData(Index).Row
            If lngRow >= vfgData(Index).FixedRows And lngRow < vfgData(Index).Rows Then
                vfgData(Index).RemoveItem (vfgData(Index).Row)
                Call vfgData_RowColChange(2)
                Me.cmdCalc.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub vfgData_RowColChange(Index As Integer)
    Dim strSQL As String
    Dim strValue As String, strKey As String
    On Error GoTo errH
    Select Case Index
    Case 1
        '������ͳ��
        With vfgData(Index)
             Call initvfgItemTitle(Index)
             If .ColIndex("ID") >= .FixedCols And .ColIndex("ID") <= .Cols - 1 Then
                strValue = Trim("" & .TextMatrix(.Row, .ColIndex("ID")))
                strKey = Trim("" & .TextMatrix(.FixedRows - 1, .ColIndex("С��")))
                If strValue <> "" And strKey <> "" Then
                    Call RefGrid_������Item(Index, strKey, strValue)
                End If
             End If
        End With
    Case 2  '���ͳ��
        With vfgData(Index)
            Call initvfgItemTitle(Index)
            If .ColIndex("�걾ID") >= .FixedCols And .ColIndex("�걾ID") <= .Cols - 1 Then
                strValue = Trim("" & .TextMatrix(.Row, .ColIndex("�걾ID")))
                If strValue <> "" Then
                    Call RefGrid_���Item(Index, strValue)
                End If
            End If
        End With
    End Select
    
    Exit Sub
errH:
    MsgBox Err.Description
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------
'�ⲿ���ù���

Public Function ShowMe(ByVal Index As Long, ShowMode As QueryShowMode, objHost As zl9LisQuery_Def.clsLisQueryHost) As Boolean
    mlgIndex = Index
    Set clsHost = objHost
    Me.Show ShowMode, objHost
End Function

'-----------------------------------------------------------------------------------------------------------------------------------
' �ڲ�����
Private Sub DoQuery()

    Dim curStart As Currency, curEnd As Currency
    On Error GoTo errH
    
    Me.MousePointer = vbHourglass
    Select Case Me.ssTMain.Tab
        Case 0
            Call RefGrid_�ճ�(0)
        Case 1
            Call RefGrid_������(1)
        Case 2
            Call RefGrid_���(2)
    End Select
    Me.MousePointer = vbDefault
    Exit Sub
errH:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbQuestion, Me.Caption
End Sub

Private Sub RefGrid_���(ByVal lngIndex As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strBegin As String, strEnd As String, lng���� As Long, str���� As String, lngС�� As Long
    Dim strWhere As String, iCol As Integer, strTitle As String
    Dim str���䵥λ As String, str���� As String, lng�������� As Long, lng�������� As Long, str������� As String, strRecord���� As String
    Dim str�Ա� As String, str������� As String, str������� As String, str������� As String
    Dim str������ As String, curSD As Currency, curAVG As Currency
    
    Dim lng��ĿID As Long ', str������� As String, strȡֵ���� As String
    Dim blnAdd As Boolean, lngƫ�� As Long, lngƫ�� As Long, lng��ʾ As Long
    Dim lngColor  As Long, lngForeColor As Long, str��־ As String
    On Error GoTo errH
    
    lngƫ�� = &H80FFFF: lngƫ�� = &H80C0FF: lng��ʾ = &H40C0&
    Call initvfgDataTitle(lngIndex)
    
    lng��ĿID = Val(txt��Ŀ.Tag)
    If lng��ĿID = 0 Then
        MsgBox "��������Ŀ����ִ�д˹���!", vbInformation, Me.Caption
        Exit Sub
    End If
'    mrs��Ŀ.Filter = ""
'    mrs��Ŀ.Filter = "��ĿID=" & lng��ĿID
'    str������� = "": str��Ŀ���� = ""
'    Do Until mrs��Ŀ.EOF
'        str������� = Trim("" & mrs��Ŀ!�������)
'        strȡֵ���� = Trim("" & mrs��Ŀ!ȡֵ����)
'        mrs��Ŀ.MoveNext
'    Loop
    
    strBegin = Format(dtpBegin(lngIndex).Value, "yyyy-MM-dd")
    strEnd = Format(dtpEnd(lngIndex).Value + 1, "yyyy-MM-dd")
    strWhere = ""
    
    strTitle = "���ڣ�" & strBegin & " �� " & strEnd
    lng���� = Val(cbo����(lngIndex).ItemData(cbo����(lngIndex).ListIndex))
    lngС�� = Val(cboС��(lngIndex).ItemData(cboС��(lngIndex).ListIndex))
    str���� = Trim(cbo����(lngIndex).List(cbo����(lngIndex).ListIndex))
    
    If str���� <> "" Then strWhere = strWhere & " And D.�������� =[6]"
    If lngС�� <> 0 Then strWhere = strWhere & " And C.ID=[5]"
    If lng���� <> 0 Then strWhere = strWhere & " And D.ID=[4]"
    
    str���䵥λ = cbo����.List(cbo����.ListIndex)
    
    str���� = Trim(txt����.Text)
    If str���� = "" Then str���䵥λ = ""
    
    str������� = "="
    If str���� Like "*-*" Then
        lng�������� = Val(Split(str����, "-")(0))
        lng�������� = Val(Split(str����, "-")(1))
         
        If lng�������� >= lng�������� Then
            MsgBox "�������޲��ܴ��ڻ�����������ޣ�", vbInformation, Me.Caption
            Exit Sub
        End If
        str������� = "Between"
    ElseIf str���� Like ">=*" Then
        lng�������� = Val(Mid(str����, 3))
        str������� = ">="
    ElseIf str���� Like "<=*" Then
        lng�������� = Val(Mid(str����, 3))
        str������� = "<="
    ElseIf str���� Like ">*" Then
        lng�������� = Val(Mid(str����, 2))
        str������� = ">"
    ElseIf str���� Like "<*" Then
        lng�������� = Val(Mid(str����, 2))
        str������� = "<"
    ElseIf str���� Like "<>*" Then
        lng�������� = Val(Mid(str����, 3))
        str������� = "<>"
    Else
        If Not IsNumeric(str����) Then str������� = "NO"
    End If
    
    If str���䵥λ <> "" Then
        If InStr("����,Ӥ��", str���䵥λ) <= 0 Then
            If lng�������� < 0 Or lng�������� < 0 Then
                MsgBox "�������޻��������޲���С��0��", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End If
    
    str�Ա� = Trim(cbo�Ա�.List(cbo�Ա�.ListIndex))
    If str�Ա� <> "" Then
        str�Ա� = Split(str�Ա�, "-")(1)
        strWhere = strWhere & " And A.�Ա� = [7] "
    End If
    
    str������� = cbo����.List(cbo����.ListIndex)
    str������� = Trim(txt����.Text)
    str������� = Trim(txt����.Text)
    
    strSQL = "Select /*+Rule */ g.����걾id, a.����ʱ��, c.���� As С��, a.�걾��� As ������, h.������ As ��Ŀ, g.������, a.����, a.�Ա�, a.����, nvl(a.���䵥λ,'��') as ���䵥λ," & vbNewLine & _
            " Decode(g.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') as �����־ " & vbNewLine & _
            "From ����������Ŀ h, ������ͨ��� g, �ϻ���Ա�� f, ����С���Ա e, �������� d, ����С�� c, ����С������ b, ����걾��¼ a" & vbNewLine & _
            "Where a.����ʱ�� Between [1] And [2] And a.����id = b.����id And" & vbNewLine & _
            "      a.������=g.��¼���� And b.С��id = c.Id And a.����id = d.Id And Nvl(a.΢����걾, 0) = 0 And c.Id = e.С��id And e.��Աid = f.��Աid And f.�û��� = User And" & vbNewLine & _
            "           a.����� Is Not Null And a.Id = g.����걾id And g.������Ŀid = h.Id And g.������Ŀid+0 = [3] " & strWhere
    
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption, CDate(strBegin), CDate(strEnd), lng��ĿID, lng����, lngС��, str����, str�Ա�)
    
    Do Until rsTmp.EOF
        blnAdd = True
        If Not (InStr("����,Ӥ��", str���䵥λ) > 0) Then
            If str���䵥λ = Trim("" & rsTmp!���䵥λ) And str���䵥λ <> "" Then
                Select Case str�������
                    Case "="
                        If Not (Val(Trim("" & rsTmp!����)) = lng��������) Then blnAdd = False
                    Case "Between"
                        If Not (Val(Trim("" & rsTmp!����)) >= lng�������� And Val(Trim("" & rsTmp!����)) <= lng��������) Then blnAdd = False
                    Case ">"
                        If Not (Val(Trim("" & rsTmp!����)) > lng��������) Then blnAdd = False
                    Case ">="
                        If Not (Val(Trim("" & rsTmp!����)) >= lng��������) Then blnAdd = False
                    Case "<"
                        If Not (Val(Trim("" & rsTmp!����)) < lng��������) Then blnAdd = False
                    Case "<="
                        If Not (Val(Trim("" & rsTmp!����)) <= lng��������) Then blnAdd = False
                    Case "<>"
                        If Not (Val(Trim("" & rsTmp!����)) <> lng��������) Then blnAdd = False
                    Case "NO"
                        blnAdd = False
                End Select
            Else
                If str���䵥λ <> "" Then blnAdd = False
            End If
        End If
        
        If blnAdd Then
            str������ = Trim("" & rsTmp!������)
            
            Select Case str�������
                Case "="
                    If IsNumeric(str������) Then
                        If Not (Val(str�������) = Val(str������)) Then blnAdd = False
                    Else
                        If Not (str������� = str������) Then blnAdd = False
                    End If
                Case "<>"
                    If IsNumeric(str������) Then
                        If Not (Val(str�������) <> Val(str������)) Then blnAdd = False
                    Else
                        If Not (str������� <> CStr(str������)) Then blnAdd = False
                    End If
                Case ">"
                    If IsNumeric(str������) Then
                        If Not (Val(str������) > Val(str�������)) Then blnAdd = False
                    Else
                        blnAdd = False
                    End If
                Case "<"
                    If IsNumeric(str������) Then
                        If Not (Val(str������) < Val(str�������)) Then blnAdd = False
                    Else
                        blnAdd = False
                    End If
                Case ">="
                    If IsNumeric(str������) Then
                        If Not (Val(str������) >= Val(str�������)) Then blnAdd = False
                    Else
                        blnAdd = False
                    End If
                Case "<="
                    If IsNumeric(str������) Then
                        If Not (Val(str������) <= Val(str�������)) Then blnAdd = False
                    Else
                        blnAdd = False
                    End If
                Case "����"
                    If Not (CStr(str������) Like "*" & str������� & "*") Then blnAdd = False
                Case "��...֮��"
                    If IsNumeric(str������) Then
                        If Not (Val(str������) >= Val(str�������) And Val(str������) <= Val(str�������)) Then blnAdd = False
                    Else
                        blnAdd = False
                    End If
            End Select
            
        End If
        If blnAdd Then
            With vfgData(lngIndex)
                .TextMatrix(.Rows - 1, .ColIndex("�걾ID")) = Val("" & rsTmp!����걾ID)
                .TextMatrix(.Rows - 1, .ColIndex("��������")) = Format("" & rsTmp!����ʱ��, "yy-MM-dd HH:mm")
                .TextMatrix(.Rows - 1, .ColIndex("����С��")) = Trim("" & rsTmp!С��)  'IIf(Val("" & rsTmp!����) = 0, "", Val("" & rsTmp!����))
                .TextMatrix(.Rows - 1, .ColIndex("������")) = Trim("" & rsTmp!������)
                .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = Trim("" & rsTmp!��Ŀ)
                
                .TextMatrix(.Rows - 1, .ColIndex("��Ŀ���")) = Trim("" & rsTmp!������)
                If IsNumeric(.TextMatrix(.Rows - 1, .ColIndex("��Ŀ���"))) Then
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("��Ŀ���")) = flexAlignRightCenter
                Else
                    .Cell(flexcpAlignment, .Rows - 1, .ColIndex("��Ŀ���")) = flexAlignLeftCenter
                End If
                
                str��־ = Trim("" & rsTmp!�����־)
                lngColor = .BackColor
                lngForeColor = .ForeColor
                If InStr("��", str��־) > 0 And str��־ <> "" Then     '2
                    lngColor = lngƫ��
                ElseIf InStr("��,�쳣", str��־) > 0 And str��־ <> "" Then '3,�쳣
                    lngColor = lngƫ��
                ElseIf InStr("����,����", str��־) > 0 And str��־ <> "" Then  '5,6
                    lngColor = lng��ʾ
                End If
                .Cell(flexcpBackColor, .Rows - 1, .ColIndex("��Ŀ���")) = lngColor
                .Cell(flexcpForeColor, .Rows - 1, .ColIndex("��Ŀ���")) = lngForeColor
                
                
                .TextMatrix(.Rows - 1, .ColIndex("����")) = Trim("" & rsTmp!����)
                .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = Trim("" & rsTmp!�Ա�)
                .TextMatrix(.Rows - 1, .ColIndex("����")) = Trim("" & rsTmp!����)
                .Rows = .Rows + 1
            End With
        End If
        
        rsTmp.MoveNext
    Loop
    
    With vfgData(lngIndex)
        '�ӱ����
        If .Rows > 2 Then .Rows = .Rows - 1
        Call CalcData
        .Select .FixedRows - 1, .FixedCols + 1, .Rows - 1, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select .FixedRows, .FixedCols
         
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("����С��"), .Rows - 1, .ColIndex("����С��")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("��Ŀ"), .Rows - 1, .ColIndex("��Ŀ")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("������"), .Rows - 1, .ColIndex("������")) = flexAlignRightCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = flexAlignRightCenter
    End With
    Exit Sub
errH:
    MsgBox "δ�ҵ����ݣ�", vbQuestion, Me.Caption
    Err.Clear
End Sub
Private Sub CalcData()
    Dim curSD As Currency, curAVG As Currency, curCount As Currency
    Dim lngRow As Long, str��� As String
        '���ֵ������
    With vfgData(2)
'        .Subtotal flexSTClear
'        .OutlineCol = 1   'ָ�������
'        .SubtotalPosition = flexSTBelow '�ϼ��ڵײ�
'
'        'SD
'        .Subtotal flexSTStd, -1, .ColIndex("��Ŀ���"), , , , , ""
'        curSD = Val(.TextMatrix(.Rows - 1, .ColIndex("��Ŀ���")))
'
'        'AVG
'        .Subtotal flexSTClear
'        .Subtotal flexSTAverage, -1, .ColIndex("��Ŀ���"), , , , , ""
'        curAVG = Val(.TextMatrix(.Rows - 1, .ColIndex("��Ŀ���")))
'
'        'Count
'        .Subtotal flexSTClear
'        .Subtotal flexSTCount, -1, .ColIndex("��Ŀ���"), , , , , ""
'        curCount = Val(.TextMatrix(.Rows - 1, .ColIndex("��Ŀ���")))
'
'        .Rows = .Rows - 1
        str��� = ""
        For lngRow = .FixedRows To .Rows - 1
            str��� = str��� & "," & Val(.TextMatrix(lngRow, .ColIndex("��Ŀ���")))
        Next
        If str��� <> "" Then
            curAVG = CalcSVG(str���)
            curSD = CalcSD(str���)
            curCount = UBound(Split(str���, ","))
        End If
        txtCount.Text = IIf(curCount = 0, "", CLng(curCount))
        txtAVG.Text = IIf(curAVG = 0, "", Format(curAVG, "0.00"))
        txtSD.Text = IIf(curSD = 0, "", Format(curSD, "0.000"))
        
        If curSD <> 0 Then
            For lngRow = .FixedRows To .Rows - 1
                str��� = Trim(.TextMatrix(lngRow, .ColIndex("��Ŀ���")))
                If IsNumeric(str���) Then
                    If Val(str���) > 4 * curSD Then
                        .TextMatrix(lngRow, .ColIndex("SD")) = ">4S"
                    ElseIf Val(str���) > 3 * curSD Then
                        .TextMatrix(lngRow, .ColIndex("SD")) = ">3S"
                    ElseIf Val(str���) > 2 * curSD Then
                        .TextMatrix(lngRow, .ColIndex("SD")) = ">2S"
                    ElseIf Val(str���) > curSD Then
                        .TextMatrix(lngRow, .ColIndex("SD")) = ">1S"
                    End If
                End If
            Next
        End If
    End With
    Me.cmdCalc.Enabled = False
End Sub

Private Sub RefGrid_���Item(ByVal Index As Long, ByVal strValue As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng�걾ID As Long, lngƫ�� As Long, lngƫ�� As Long, lng��ʾ As Long
    Dim lngColor  As Long, lngForeColor As Long, str��־ As String
    
    lngƫ�� = &H80FFFF: lngƫ�� = &H80C0FF: lng��ʾ = &H40C0&
    
    lng�걾ID = Val(strValue)
    If lng�걾ID = 0 Then Exit Sub
    strSQL = "Select c.������ As ��Ŀ, d.��д as Ӣ����, b.������, Decode(B.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') as �����־, b.����ο�" & vbNewLine & _
        "From ������Ŀ d,����������Ŀ c, ������ͨ��� b, ����걾��¼ a" & vbNewLine & _
        "Where b.������Ŀid=d.������Ŀid And a.Id = b.����걾id And a.������ = b.��¼���� And b.������Ŀid = c.Id And a.Id = [1]" & vbNewLine & _
        " Order by Nvl(to_number(d.�������),c.����)"
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption, lng�걾ID)
    Do Until rsTmp.EOF
        With vfgItem(Index)
            .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = Trim("" & rsTmp!��Ŀ)
            .TextMatrix(.Rows - 1, .ColIndex("Ӣ����")) = Trim("" & rsTmp!Ӣ����)
            .TextMatrix(.Rows - 1, .ColIndex("��Ŀֵ")) = Trim("" & rsTmp!������)
            If IsNumeric(.TextMatrix(.Rows - 1, .ColIndex("��Ŀֵ"))) Then
                .Cell(flexcpAlignment, .Rows - 1, .ColIndex("��Ŀֵ")) = flexAlignRightCenter
            Else
                .Cell(flexcpAlignment, .Rows - 1, .ColIndex("��Ŀֵ")) = flexAlignLeftCenter
            End If
            
            str��־ = Trim("" & rsTmp!�����־)
            lngColor = .BackColor
            lngForeColor = .ForeColor
            If InStr("��", str��־) > 0 And str��־ <> "" Then     '2
                lngColor = lngƫ��
            ElseIf InStr("��,�쳣", str��־) > 0 And str��־ <> "" Then '3,�쳣
                lngColor = lngƫ��
            ElseIf InStr("����,����", str��־) > 0 And str��־ <> "" Then  '5,6
                lngColor = lng��ʾ
            End If
            .Cell(flexcpBackColor, .Rows - 1, .ColIndex("��Ŀֵ")) = lngColor
            .Cell(flexcpForeColor, .Rows - 1, .ColIndex("��Ŀֵ")) = lngForeColor
            
            .TextMatrix(.Rows - 1, .ColIndex("״̬")) = str��־
            .TextMatrix(.Rows - 1, .ColIndex("�ο���Χ")) = Trim("" & rsTmp!����ο�)
            .Rows = .Rows + 1
        End With
        rsTmp.MoveNext
    Loop
    
    With vfgItem(Index)
        '�ӱ����
        If .Rows > 2 Then .Rows = .Rows - 1
        .Select .FixedRows - 1, .FixedCols, .Rows - 1, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select .FixedRows, .FixedCols
        
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("��Ŀ"), .Rows - 1, .ColIndex("��Ŀ")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("Ӣ����"), .Rows - 1, .ColIndex("Ӣ����")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("�ο���Χ"), .Rows - 1, .ColIndex("�ο���Χ")) = flexAlignLeftCenter
    End With
End Sub

Private Sub RefGrid_������Item(ByVal lngIndex As Long, ByVal strKey As String, ByVal strValue As String)
    '
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strBegin As String, strEnd As String, lng���� As Long, str���� As String, lngС�� As Long
    Dim strWhere As String, iCol As Integer
    Dim lng������� As Long, str������ As String, str������ As String, str����� As String
    Dim lng������Դ As Long, strͳ�Ʒ�ʽ As String
    Dim lng������ As Long, lng������Ŀid As Long
    
    'Call initvfgItemTitle(lngIndex)
    If InStr("С��,����,��Ŀ,�������,������,������,�����,������Դ", strKey) <= 0 Then Exit Sub
    strBegin = Format(dtpBegin(lngIndex).Value, "yyyy-MM-dd")
    strEnd = Format(dtpEnd(lngIndex).Value + 1, "yyyy-MM-dd")
    strWhere = ""
    
    lng���� = Val(cbo����(lngIndex).ItemData(cbo����(lngIndex).ListIndex))
    lngС�� = Val(cboС��(lngIndex).ItemData(cboС��(lngIndex).ListIndex))
    str���� = Trim(cbo����(lngIndex).List(cbo����(lngIndex).ListIndex))
    lng������� = Val(cbo�������(lngIndex).ItemData(cbo�������(lngIndex).ListIndex))
    str������ = Trim(cbo������(lngIndex).List(cbo������(lngIndex).ListIndex))
    str������ = Trim(cbo������(lngIndex).List(cbo������(lngIndex).ListIndex))
    str����� = Trim(cbo�����(lngIndex).List(cbo�����(lngIndex).ListIndex))

    
    lng������Դ = 0
    If opt��Դ(0).Value = True Then
        lng������Դ = 0
    ElseIf opt��Դ(1).Value = True Then
        lng������Դ = 1
    ElseIf opt��Դ(2).Value = True Then
        lng������Դ = 2
    ElseIf opt��Դ(3).Value = True Then
        lng������Դ = 3
    ElseIf opt��Դ(4).Value = True Then
        lng������Դ = 4
    End If
    
    Select Case strKey
        Case "С��"
            lngС�� = Val(strValue)
        Case "����"
            lng���� = Val(strValue)
        Case "��Ŀ"
            lng������Ŀid = Val(strValue)
        Case "�������"
            lng������� = Val(strValue)
        Case "������"
            str������ = Trim(strValue)
        Case "������"
            str������ = Trim(strValue)
        Case "�����"
            str����� = Trim(strValue)
        Case "������Դ"
            lng������Դ = Val(strValue)
    End Select
    
    If str���� <> "" Then strWhere = strWhere & " And D.�������� =[5]"
    If lngС�� <> 0 Then strWhere = strWhere & " And C.ID=[4]"
    If lng���� <> 0 Then strWhere = strWhere & " And N.����ID=[3]"
    If str����� <> "" Then strWhere = strWhere & " And N.�����=[9]"
    If str������ <> "" Then strWhere = strWhere & " And N.������=[8]"
    If str������ <> "" Then strWhere = strWhere & " And N.������=[7]"
    If lng������� <> 0 Then strWhere = strWhere & " And N.�������ID= [6]"
    If lng������Դ <> 0 Then strWhere = strWhere & " And N.������Դ=[10]"
    If lng������Ŀid <> 0 Then strWhere = strWhere & " And N.������Ŀid=[11]"
    
    strSQL = "Select n.������, n.��д, Count(n.Id) As ����" & vbNewLine & _
            "From �ϻ���Ա�� f, ����С���Ա e, �������� d, ����С�� c, ����С������ b," & vbNewLine & _
            "        (Select b.������Ŀid,a.������Դ, a.����id, a.�������id, a.������, a.�����, a.������, a.Id," & vbNewLine & _
            "                           b.������Ŀid, c.������, d.��д" & vbNewLine & _
            "            From ������Ŀ d, ����������Ŀ c, ������ͨ��� b, ����걾��¼ a" & vbNewLine & _
            "            Where a.Id = b.����걾id And b.������Ŀid = c.Id And c.Id = d.������Ŀid And" & vbNewLine & _
            "                  a.������=b.��¼���� And a.����ʱ�� Between [1] And [2] And nvl(a.΢����걾,0) = 0 And" & vbNewLine & _
            "                        a.����� Is Not Null) n" & vbNewLine & _
            "Where n.����id = b.����id And b.С��id = c.Id And n.����id = d.Id And c.Id = e.С��id And e.��Աid = f.��Աid And f.�û��� = User" & vbNewLine & _
            strWhere & vbNewLine & _
            "Group By n.������, n.��д"
        
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption, CDate(strBegin), CDate(strEnd), lng����, lngС��, str����, lng�������, str������, str������, str�����, lng������Դ, lng������Ŀid)
    With vfgItem(lngIndex)
        '�������
        lng������ = 0
        
        Do Until rsTmp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = Trim("" & rsTmp!������)
            .TextMatrix(.Rows - 1, .ColIndex("Ӣ����")) = Trim("" & rsTmp!��д)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(Val("" & rsTmp!����) = 0, "", Val("" & rsTmp!����))
            lng������ = lng������ + Val("" & rsTmp!����)
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        
        .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = "�ϼ�"
        .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(lng������ = 0, "", lng������)
        
        '�ӱ����
        .Select .FixedRows - 1, .FixedCols, .Rows - 1, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select .FixedRows, .FixedCols
         
        .Cell(flexcpAlignment, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Ӣ����")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("Ӣ����") + 1, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        
    End With

End Sub

Private Sub RefGrid_������(ByVal lngIndex As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strBegin As String, strEnd As String, lng���� As Long, str���� As String, lngС�� As Long
    Dim strWhere As String, iCol As Integer, strTitle As String
    Dim lng������� As Long, str������ As String, str������ As String, str����� As String
    Dim lng������Դ As Long, strͳ�Ʒ�ʽ As String
    Dim cur��� As Currency, lng������ As Long
    Call initvfgDataTitle(lngIndex)
    strBegin = Format(dtpBegin(lngIndex).Value, "yyyy-MM-dd")
    strEnd = Format(dtpEnd(lngIndex).Value + 1, "yyyy-MM-dd")
    strWhere = ""
    strTitle = "���ڣ�" & strBegin & " �� " & strEnd
    lng���� = Val(cbo����(lngIndex).ItemData(cbo����(lngIndex).ListIndex))
    lngС�� = Val(cboС��(lngIndex).ItemData(cboС��(lngIndex).ListIndex))
    str���� = Trim(cbo����(lngIndex).List(cbo����(lngIndex).ListIndex))
    
    If str���� <> "" Then
        strWhere = strWhere & " And D.�������� =[5]"
        strTitle = strTitle & "  ��������:" & Trim(cbo����(lngIndex).List(cbo����(lngIndex).ListIndex))
    End If
    
    If lngС�� <> 0 Then
        strWhere = strWhere & " And C.ID=[4]"
        strTitle = strTitle & "  С��:" & Trim(cboС��(lngIndex).List(cboС��(lngIndex).ListIndex))
    End If

    If lng���� <> 0 Then
        strWhere = strWhere & " And N.����ID=[3]"
        strTitle = strTitle & "  ����:" & Trim(cbo����(lngIndex).List(cbo����(lngIndex).ListIndex))
    End If
    
    If strWhere <> "" Then strWhere = strWhere & vbNewLine
    lng������� = Val(cbo�������(lngIndex).ItemData(cbo�������(lngIndex).ListIndex))
    If lng������� <> 0 Then
        strWhere = strWhere & " And N.�������ID= [6]"
        strTitle = strTitle & "  �������:" & Trim(cbo�������(lngIndex).List(cbo�������(lngIndex).ListIndex))
    End If
    
    str������ = Trim(cbo������(lngIndex).List(cbo������(lngIndex).ListIndex))
    If str������ <> "" Then
        strWhere = strWhere & " And N.������=[7]"
        strTitle = strTitle & "  ������:" & str������
    End If
    
    str������ = Trim(cbo������(lngIndex).List(cbo������(lngIndex).ListIndex))
    If str������ <> "" Then
        strWhere = strWhere & " And N.������=[8]"
        strTitle = strTitle & "  ������:" & str������
    End If
    
    str����� = Trim(cbo�����(lngIndex).List(cbo�����(lngIndex).ListIndex))
    If str����� <> "" Then
        strWhere = strWhere & " And N.�����=[9]"
        strTitle = strTitle & "  �����:" & str�����
    End If
    
    lng������Դ = 0
    If opt��Դ(0).Value = True Then
        lng������Դ = 0
    ElseIf opt��Դ(1).Value = True Then
        lng������Դ = 1
        strTitle = strTitle & "  ������Դ:����"
    ElseIf opt��Դ(2).Value = True Then
        lng������Դ = 2
        strTitle = strTitle & "  ������Դ:סԺ"
    ElseIf opt��Դ(3).Value = True Then
        lng������Դ = 3
        strTitle = strTitle & "  ������Դ:Ժ��"
    ElseIf opt��Դ(4).Value = True Then
        lng������Դ = 4
        strTitle = strTitle & "  ������Դ:���"
    End If
    
    If lng������Դ <> 0 Then
        strWhere = strWhere & " And N.������Դ=[10]"
    End If
    
    If opt�շ�(1).Value = True Then
        strWhere = strWhere & " And Nvl(N.��¼״̬,0) <> 0 "
        strTitle = strTitle & "  ���շ�"
    End If
    If opt�շ�(2).Value = True Then
        strWhere = strWhere & " And Nvl(N.��¼״̬,0) = 0 "
        strTitle = strTitle & "  δ�շ�"
    End If
    
    strͳ�Ʒ�ʽ = "С��"
    If optͳ�Ʒ�ʽ(0).Value = True Then
        strͳ�Ʒ�ʽ = "С��"
    ElseIf optͳ�Ʒ�ʽ(1).Value = True Then
        strͳ�Ʒ�ʽ = "����"
    ElseIf optͳ�Ʒ�ʽ(2).Value = True Then
        strͳ�Ʒ�ʽ = "��Ŀ"
    ElseIf optͳ�Ʒ�ʽ(3).Value = True Then
        strͳ�Ʒ�ʽ = "�������"
    ElseIf optͳ�Ʒ�ʽ(4).Value = True Then
        strͳ�Ʒ�ʽ = "������"
    ElseIf optͳ�Ʒ�ʽ(5).Value = True Then
        strͳ�Ʒ�ʽ = "������"
    ElseIf optͳ�Ʒ�ʽ(6).Value = True Then
        strͳ�Ʒ�ʽ = "�����"
    ElseIf optͳ�Ʒ�ʽ(7).Value = True Then
        strͳ�Ʒ�ʽ = "������Դ"
    End If
    
    
    With vfgData(lngIndex)
        If .FixedRows >= 2 Then
            For iCol = .FixedCols To .Cols - 1
                .TextMatrix(.FixedRows - 2, iCol) = strTitle
            Next
        End If
    End With
    If lng������Դ = 0 Then
        strSQL = "            From ���ű� g, �ϻ���Ա�� f, ����С���Ա e, �������� d, ����С�� c, ����С������ b," & vbNewLine & _
                "                       (Select Distinct c.������Ŀid,a.������Դ,a.�������id, a.������, a.������, a.�����, a.Id, a.ҽ��id, a.����id, a.����ʱ��, a.�걾���, e.����, e.���� as ��Ŀ, D.No, D.���, d.��¼����, d.��¼״̬, d.ʵ�ս�� " & vbNewLine & _
                "                           From ������ĿĿ¼ e, סԺ���ü�¼ d, ����ҽ����¼ c, ������Ŀ�ֲ� b, ����걾��¼ a" & vbNewLine & _
                "                           Where a.Id = b.�걾id And b.ҽ��id = c.���id And c.Id = d.ҽ�����(+) And c.������Ŀid = e.Id And" & vbNewLine & _
                "                                       a.����ʱ�� Between [1] And [2] And" & vbNewLine & _
                "                                       Nvl(a.΢����걾, 0) = 0 And a.����� Is Not Null" & _
                "                         union all             " & _
                "                        Select Distinct c.������Ŀid,a.������Դ,a.�������id, a.������, a.������, a.�����, a.Id, a.ҽ��id, a.����id, a.����ʱ��, a.�걾���, e.����, e.���� as ��Ŀ, D.No, D.���, d.��¼����, d.��¼״̬, d.ʵ�ս�� " & vbNewLine & _
                "                           From ������ĿĿ¼ e, ������ü�¼ d, ����ҽ����¼ c, ������Ŀ�ֲ� b, ����걾��¼ a" & vbNewLine & _
                "                           Where a.Id = b.�걾id And b.ҽ��id = c.���id And c.Id = d.ҽ�����(+) And c.������Ŀid = e.Id And" & vbNewLine & _
                "                                       a.����ʱ�� Between [1] And [2] And" & vbNewLine & _
                "                                       Nvl(a.΢����걾, 0) = 0 And a.����� Is Not Null" & _
                "                         ) n" & vbNewLine & _
                "            Where n.�������id = g.Id And n.����id = b.����id And b.С��id = c.Id And n.����id = d.Id And c.Id = e.С��id And" & vbNewLine & _
                "                        e.��Աid = f.��Աid And f.�û��� = User"
    
    ElseIf lng������Դ = 2 Then
    
        strSQL = "            From ���ű� g, �ϻ���Ա�� f, ����С���Ա e, �������� d, ����С�� c, ����С������ b," & vbNewLine & _
                "                       (Select Distinct c.������Ŀid,a.������Դ,a.�������id, a.������, a.������, a.�����, a.Id, a.ҽ��id, a.����id, a.����ʱ��, a.�걾���, e.����, e.���� as ��Ŀ, D.No, D.���, d.��¼����, d.��¼״̬, d.ʵ�ս�� " & vbNewLine & _
                "                           From ������ĿĿ¼ e, סԺ���ü�¼ d, ����ҽ����¼ c, ������Ŀ�ֲ� b, ����걾��¼ a" & vbNewLine & _
                "                           Where a.Id = b.�걾id And b.ҽ��id = c.���id And c.Id = d.ҽ�����(+) And c.������Ŀid = e.Id And" & vbNewLine & _
                "                                       a.����ʱ�� Between [1] And [2] And" & vbNewLine & _
                "                                       Nvl(a.΢����걾, 0) = 0 And a.����� Is Not Null) n" & vbNewLine & _
                "            Where n.�������id = g.Id And n.����id = b.����id And b.С��id = c.Id And n.����id = d.Id And c.Id = e.С��id And" & vbNewLine & _
                "                        e.��Աid = f.��Աid And f.�û��� = User"

    Else
        strSQL = "            From ���ű� g, �ϻ���Ա�� f, ����С���Ա e, �������� d, ����С�� c, ����С������ b," & vbNewLine & _
                "                       (Select Distinct c.������Ŀid,a.������Դ,a.�������id, a.������, a.������, a.�����, a.Id, a.ҽ��id, a.����id, a.����ʱ��, a.�걾���, e.����, e.���� as ��Ŀ, D.No, D.���, d.��¼����, d.��¼״̬, d.ʵ�ս�� " & vbNewLine & _
                "                           From ������ĿĿ¼ e, ������ü�¼ d, ����ҽ����¼ c, ������Ŀ�ֲ� b, ����걾��¼ a" & vbNewLine & _
                "                           Where a.Id = b.�걾id And b.ҽ��id = c.���id And c.Id = d.ҽ�����(+) And c.������Ŀid = e.Id And" & vbNewLine & _
                "                                       a.����ʱ�� Between [1] And [2] And" & vbNewLine & _
                "                                       Nvl(a.΢����걾, 0) = 0 And a.����� Is Not Null) n" & vbNewLine & _
                "            Where n.�������id = g.Id And n.����id = b.����id And b.С��id = c.Id And n.����id = d.Id And c.Id = e.С��id And" & vbNewLine & _
                "                        e.��Աid = f.��Աid And f.�û��� = User"
    
    End If
    Select Case strͳ�Ʒ�ʽ
    Case "С��"
        strSQL = "Select С��id As Id, С��, Count(Id) As ������, Sum(ʵ�ս��) As ���" & vbNewLine & _
                 "From (Select c.Id As С��id, c.���� As С��,n.Id, Sum(Nvl(n.ʵ�ս��, 0)) As ʵ�ս��" & vbNewLine & _
                 strSQL & strWhere & _
                "      Group By c.Id, c.����, n.Id)" & vbNewLine & _
                "Group By С��id, С��"
    Case "����"
        strSQL = "Select ����id As Id, ����, Count(Id) As ������, Sum(ʵ�ս��) As ���" & vbNewLine & _
                 "From (Select d.Id As ����id, d.���� As ����,n.Id, Sum(Nvl(n.ʵ�ս��, 0)) As ʵ�ս��" & vbNewLine & _
                 strSQL & strWhere & _
                "      Group By d.Id, d.����, n.Id)" & vbNewLine & _
                "Group By ����id, ����"

    Case "�������"
        strSQL = "Select �������id as id,�������, Count(Id) As ������, Sum(ʵ�ս��) As ���" & vbNewLine & _
                "From (Select n.�������id,g.���� As �������, n.Id, Sum(Nvl(n.ʵ�ս��, 0)) As ʵ�ս��" & vbNewLine & _
                 strSQL & strWhere & _
                "            Group By n.�������id,g.����, n.Id)" & vbNewLine & _
                "Group By �������id,�������"
    Case "��Ŀ"
        strSQL = "Select ������Ŀid as id," & strͳ�Ʒ�ʽ & ", Count(Id) As ������, Sum(ʵ�ս��) As ���" & vbNewLine & _
                "From (Select n.������Ŀid,n." & strͳ�Ʒ�ʽ & ", n.Id, Sum(Nvl(n.ʵ�ս��, 0)) As ʵ�ս��" & vbNewLine & _
                strSQL & strWhere & _
                "            Group By n.������Ŀid, n." & strͳ�Ʒ�ʽ & ", n.Id)" & vbNewLine & _
                "Group By ������Ŀid," & strͳ�Ʒ�ʽ
    
    Case "������Դ"
        strSQL = "Select ������Դ as id, decode(������Դ,1,'����',2,'סԺ',4,'���','Ժ��') as ������Դ, Count(Id) As ������, Sum(ʵ�ս��) As ���" & vbNewLine & _
                "From (Select n." & strͳ�Ʒ�ʽ & ", n.Id, Sum(Nvl(n.ʵ�ս��, 0)) As ʵ�ս��" & vbNewLine & _
                strSQL & strWhere & _
                "            Group By n.������Դ,decode(n.������Դ,1,'����',2,'סԺ',4,'���','Ժ��'), n.Id)" & vbNewLine & _
                "Group By " & strͳ�Ʒ�ʽ
                
    Case "������", "�����", "������"
        strSQL = "Select " & strͳ�Ʒ�ʽ & " as id," & strͳ�Ʒ�ʽ & ", Count(Id) As ������, Sum(ʵ�ս��) As ���" & vbNewLine & _
                "From (Select n." & strͳ�Ʒ�ʽ & ", n.Id, Sum(Nvl(n.ʵ�ս��, 0)) As ʵ�ս��" & vbNewLine & _
                strSQL & strWhere & _
                "            Group By n." & strͳ�Ʒ�ʽ & ", n.Id)" & vbNewLine & _
                "Group By " & strͳ�Ʒ�ʽ
    
    End Select
    
    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption, CDate(strBegin), CDate(strEnd), lng����, lngС��, str����, lng�������, str������, str������, str�����, lng������Դ)
    
    With vfgData(lngIndex)
        .TextMatrix(.FixedRows - 1, .ColIndex("С��")) = strͳ�Ʒ�ʽ
        '�������
        lng������ = 0: cur��� = 0
        
        Do Until rsTmp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = Trim("" & rsTmp!Id)
            .TextMatrix(.Rows - 1, .ColIndex("С��")) = Trim("" & rsTmp.Fields(1))
            .TextMatrix(.Rows - 1, .ColIndex("������")) = IIf(Val("" & rsTmp!������) = 0, "", Val("" & rsTmp!������))
            .TextMatrix(.Rows - 1, .ColIndex("���")) = IIf(Val("" & rsTmp!���) = 0, "", Format(Val("" & rsTmp!���), "0.00"))
            lng������ = lng������ + Val("" & rsTmp!������)
            cur��� = cur��� + Val("" & rsTmp!���)
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        
        .TextMatrix(.Rows - 1, .ColIndex("С��")) = "�ϼ�"
        .TextMatrix(.Rows - 1, .ColIndex("������")) = IIf(lng������ = 0, "", lng������)
        .TextMatrix(.Rows - 1, .ColIndex("���")) = IIf(cur��� = 0, "", Format(cur���, "0.00"))
        
        '�ӱ����
        .Select .FixedRows - 1, .FixedCols + 1, .Rows - 1, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select .FixedRows, .FixedCols + 1
         
        .Cell(flexcpAlignment, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("С��")) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, .ColIndex("С��") + 1, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        
    End With
    
    
End Sub

Private Sub RefGrid_�ճ�(ByVal lngIndex As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strBegin As String
    Dim strEnd As String
    Dim lng���� As Long, str���� As String, lngС�� As Long
    Dim strWhere As String, iCol As Integer, strTitle As String '������
    
    Call initvfgDataTitle(lngIndex)

    strBegin = Format(dtpBegin(lngIndex).Value, "yyyy-MM-dd")
    strEnd = Format(dtpEnd(lngIndex).Value + 1, "yyyy-MM-dd")
    strWhere = ""
    strTitle = "���ڣ�" & strBegin & " �� " & strEnd
    lng���� = Val(cbo����(lngIndex).ItemData(cbo����(lngIndex).ListIndex))
    lngС�� = Val(cboС��(lngIndex).ItemData(cboС��(lngIndex).ListIndex))
    str���� = Trim(cbo����(lngIndex).List(cbo����(lngIndex).ListIndex))
    
    If str���� <> "" Then
        strWhere = strWhere & " And D.�������� =[5]"
        strTitle = strTitle & "  ��������:" & Trim(cbo����(lngIndex).List(cbo����(lngIndex).ListIndex))
    Else
        strTitle = strTitle & "  ��������:��������"
    End If
    
    If lngС�� <> 0 Then
        strWhere = strWhere & " And C.ID=[4]"
        strTitle = strTitle & "  С��:" & Trim(cboС��(lngIndex).List(cboС��(lngIndex).ListIndex))
    Else
        strTitle = strTitle & "  С��:����С��"
    End If

    If lng���� <> 0 Then
        strWhere = strWhere & " And A.����ID = [3]"
        strTitle = strTitle & "  ����:" & Trim(cbo����(lngIndex).List(cbo����(lngIndex).ListIndex))
    Else
        strTitle = strTitle & "  ����:��������"
    End If
        

    With vfgData(lngIndex)
        If .FixedRows >= 2 Then
            For iCol = .FixedCols To .Cols - 1
                .TextMatrix(.FixedRows - 2, iCol) = strTitle
            Next
        End If
    End With
    
    strSQL = "Select D.�������� As ����, c.���� As С��, d.���� As ����, Sum(Decode(a.����, Null, 1, 0)) As ����," & vbNewLine & _
            "            Sum(Decode(a.������, Null, 0, Decode(a.�����, Null, 1, 0))) As �ѽ���, Count(a.Id) As �Ѻ���," & vbNewLine & _
            "            Sum(Decode(a.�����, Null, 0, 1)) As ���, Sum(Decode(a.�����, Null, 1, 0)) As δ��" & vbNewLine & _
            "From �ϻ���Ա�� f, ����С���Ա e, �������� D, ����С�� C, ����С������ b, ����걾��¼ a" & vbNewLine & _
            "Where a.����ʱ�� Between [1] And [2] And a.����id = b.����id And" & vbNewLine & _
            "           b.С��id = c.Id And a.����id = d.Id And Nvl(a.΢����걾,0)=0 And c.Id = e.С��id And e.��Աid = f.��Աid And f.�û��� = User" & vbNewLine & _
            strWhere & vbNewLine & _
            "Group By D.��������, c.����, d.����"

    Set rsTmp = clsHost.GetRecordSet(strSQL, Me.Caption, CDate(strBegin), CDate(strEnd), lng����, lngС��, str����)
    
    With vfgData(lngIndex)
        '�������
        Do Until rsTmp.EOF
            .TextMatrix(.Rows - 1, mCol_�ճ�.����) = Trim("" & rsTmp!����)
            .TextMatrix(.Rows - 1, mCol_�ճ�.С��) = Trim("" & rsTmp!С��)
            .TextMatrix(.Rows - 1, mCol_�ճ�.����) = Trim("" & rsTmp!����)
            .TextMatrix(.Rows - 1, mCol_�ճ�.����) = IIf(Val("" & rsTmp!����) = 0, "", Val("" & rsTmp!����))
            .TextMatrix(.Rows - 1, mCol_�ճ�.�ѽ���) = IIf(Val("" & rsTmp!�ѽ���) = 0, "", Val("" & rsTmp!�ѽ���))
            .TextMatrix(.Rows - 1, mCol_�ճ�.�Ѻ���) = IIf(Val("" & rsTmp!�Ѻ���) = 0, "", Val("" & rsTmp!�Ѻ���))
            .TextMatrix(.Rows - 1, mCol_�ճ�.���) = IIf(Val("" & rsTmp!���) = 0, "", Val("" & rsTmp!���))
            .TextMatrix(.Rows - 1, mCol_�ճ�.δ��) = IIf(Val("" & rsTmp!δ��) = 0, "", Val("" & rsTmp!δ��))
            
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop

        If .Rows > 4 Then .Rows = .Rows - 1
        
        '��ϼ�
        .Subtotal flexSTClear
        .OutlineCol = 1   'ָ�������
        .SubtotalPosition = flexSTBelow '�ϼ��ڵײ�
        .Subtotal flexSTSum, -1, .ColIndex("����"), , , , , "�ϼ�"
        .Subtotal flexSTSum, -1, .ColIndex("�Ѻ���"), , , , , "�ϼ�"
        .Subtotal flexSTSum, -1, .ColIndex("�ѽ���"), , , , , "�ϼ�"
        .Subtotal flexSTSum, -1, .ColIndex("�Ѻ���"), , , , , "�ϼ�"
        .Subtotal flexSTSum, -1, .ColIndex("���"), , , , , "�ϼ�"
        .Subtotal flexSTSum, -1, .ColIndex("δ��"), , , , , "�ϼ�"
        
        For iCol = .ColIndex("����") To .Cols - 1
            .TextMatrix(.Rows - 1, iCol) = Replace(.TextMatrix(.Rows - 1, iCol), ".00", "")
        Next
        
        '�ӱ����
        
        .Select .FixedRows - 1, .FixedCols, .Rows - 1, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select .FixedRows, .FixedCols
         
        .Cell(flexcpAlignment, .FixedRows, .FixedCols, .Rows - 1, mCol_�ճ�.����) = flexAlignLeftCenter
        .Cell(flexcpAlignment, .FixedRows, mCol_�ճ�.����, .Rows - 1, .Cols - 1) = flexAlignRightCenter
    End With

End Sub

Private Sub initvfgDataTitle(ByVal Index As Long)
    Dim strFiles As String, strTitle As String, strFont As String
    
    Select Case Index
    Case 0
        strFiles = ",100;����,900;С��,900;����,1800;����,900;�ѽ���,900;�Ѻ���,900;���,900;δ��,900"
        strTitle = Trim(ReadIni("Report0", "����", App.Path & "\PrintSetup.ini"))
        strFont = Trim(ReadIni("Report0", "��������", App.Path & "\PrintSetup.ini"))
        If strTitle = "" Then strTitle = "�ճ�����"
        
        vfgData(0).Rows = 4: vfgData(0).Cols = 9
        
        Call initVfg(vfgData(0), strFiles, strTitle, strFont)
        With vfgData(0)
            .Select .FixedRows - 1, .FixedCols, .Rows - 1, .Cols - 1
            .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        End With
        
    Case 1
        strFiles = ",100;ID,0;С��,2800;������,2000;���,2000"
        strTitle = Trim(ReadIni("Report1", "����", App.Path & "\zl9LisQuery_Base.ini"))
        strFont = Trim(ReadIni("Report1", "��������", App.Path & "\PrintSetup.ini"))
        If strTitle = "" Then strTitle = "������ͳ��"
        vfgData(1).Rows = 4: vfgData(1).Cols = 5
        Call initVfg(vfgData(1), strFiles, strTitle, strFont)
        With vfgData(1)
            .Select .FixedRows - 1, .FixedCols + 1, .Rows - 1, .Cols - 1
            .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        End With
    Case 2
        strFiles = ",100;�걾ID,0;��������,1800;����С��,1200;������,900;��Ŀ,2200;��Ŀ���,1200;����,900;�Ա�,800;����,800;SD,1000"
        strTitle = Trim(ReadIni("Report2", "����", App.Path & "\zl9LisQuery_Base.ini"))
        strFont = Trim(ReadIni("Report2", "��������", App.Path & "\PrintSetup.ini"))
        If strTitle = "" Then strTitle = "���ͳ��"
        
        vfgData(2).Rows = 4: vfgData(2).Cols = 11
        Call initVfg(vfgData(2), strFiles, strTitle, strFont)
        With vfgData(2)
            .Select .FixedRows - 1, .FixedCols + 1, .Rows - 1, .Cols - 1
            .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        End With
    End Select
End Sub

Private Sub initvfgItemTitle(ByVal Index As Long)
    Dim strFiles As String, strTitle As String
    Select Case Index
    Case 1
        strFiles = ",0;��Ŀ,2800;Ӣ����,1000;����,2000"
        strTitle = ""
        
        vfgItem(1).Rows = 2: vfgItem(1).Cols = 4
        Call initVfg(vfgItem(1), strFiles, strTitle, "")
        With vfgItem(1)
            .Select .FixedRows - 1, .FixedCols, .Rows - 1, .Cols - 1
            .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        End With
    Case 2
        strFiles = ",0;��Ŀ,1800;Ӣ����,1000;��Ŀֵ,1000;״̬,800;�ο���Χ,2000;"
        strTitle = ""
        vfgItem(2).Rows = 2: vfgItem(2).Cols = 6
        Call initVfg(vfgItem(2), strFiles, strTitle, "")
        With vfgItem(2)
            .Select .FixedRows - 1, .FixedCols, .Rows - 1, .Cols - 1
            .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        End With
    End Select
End Sub

Private Sub initVfg(objVfg As VSFlexGrid, ByVal str�ֶ� As String, ByVal strTitle As String, ByVal strFont As String)
    Dim iCol As Integer
    Dim varTmp As Variant, varTmp1 As Variant
    On Error GoTo errH
    varTmp = Split(str�ֶ�, ";")
    
    If UBound(Split(strFont, "|")) <> 2 Then strFont = "����|18"
    
    With objVfg
        .Clear
        .Editable = flexEDNone
        .GridLines = flexGridNone
        
        .MergeCells = flexMergeRestrictRows
        .BackColorFixed = .BackColor
        .ForeColorFixed = .ForeColor
        .GridColorFixed = .GridColor
        .GridLinesFixed = flexGridNone
        
                
        If strTitle <> "" Then
            If .Rows < 4 Then Exit Sub
            .FixedCols = 1: .FixedRows = 3
            '-- ��ͷ
            For iCol = 0 To 1
                .MergeRow(iCol) = True
            Next
            
            If strTitle <> "" Then
                For iCol = .FixedCols To .Cols - 1
                    .TextMatrix(0, iCol) = strTitle
                Next
            End If
            
            .Cell(flexcpFontName, 0, .FixedCols, 0, .Cols - 1) = Split(strFont, "|")(0)
            .Cell(flexcpFontSize, 0, .FixedCols, 0, .Cols - 1) = Split(strFont, "|")(1)
            .Cell(flexcpFontBold, 0, .FixedCols, 0, .Cols - 1) = True
            .RowHeight(0) = 600
            .RowHeight(1) = 500
            .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        Else
            If .Rows < 2 Then Exit Sub
            .FixedCols = 1: .FixedRows = 1
        End If
        
        For iCol = LBound(varTmp) To UBound(varTmp)
            If InStr(varTmp(iCol), ",") > 0 Then
                varTmp1 = Split(varTmp(iCol), ",")
                .TextMatrix(.FixedRows - 1, iCol) = Trim(varTmp1(0))
                If .TextMatrix(.FixedRows - 1, iCol) <> "" Then .ColKey(iCol) = .TextMatrix(.FixedRows - 1, iCol)
                .ColWidth(iCol) = Val(varTmp1(1))
                If .ColWidth(iCol) = 0 Then .ColHidden(iCol) = True
                .ColAlignment(iCol) = flexAlignCenterCenter
            End If
        Next
    End With
    Exit Sub
errH:
    MsgBox "initvfg" & vbCrLf & str�ֶ� & vbCrLf & Err.Description, vbQuestion, Me.Caption
End Sub

Private Sub ShowSelect(ByVal strInput As String)
    Dim strSQL As String
    
    Dim strP1 As String
    Dim strWhere As String
    
    If strInput <> "" Then
        strWhere = " And (D.���� Like '%" & UCase(strInput) & "%' Or Upper(D.������) Like '%" & UCase(strInput) & "%' Or Upper(C.��д) Like '%" & UCase(strInput) & "%')"
    End If
    strSQL = "Select c.������Ŀid As ��Ŀid, d.����, d.������, c.��д, d.��λ, c.��Ŀ���, c.�������, c.ȡֵ����" & vbNewLine & _
            "From ����������Ŀ d, ������Ŀ c" & vbNewLine & _
            "Where c.������Ŀid = d.Id " & strWhere & " Order by D.����"
            
    Set mrs��Ŀ = clsHost.GetRecordSet(strSQL, Me.Caption)
    
    lstSelect.Clear
    Do Until mrs��Ŀ.EOF
        lstSelect.AddItem "" & mrs��Ŀ!���� & "-" & mrs��Ŀ!������ & IIf(Trim("" & mrs��Ŀ!��д) = "", "", "(" & mrs��Ŀ!��д & ")")
        lstSelect.ItemData(lstSelect.NewIndex) = Val("" & mrs��Ŀ!��Ŀid)
        mrs��Ŀ.MoveNext
    Loop
    If lstSelect.ListCount > 0 Then
        Call MoveSelect(txt��Ŀ)
        lstSelect.ListIndex = 0
        lstSelect.Visible = True
        lstSelect.SetFocus
    End If
End Sub

Private Sub MoveSelect(ByVal ctrl As Control)
    
    Dim vRect As RECT
    Dim vRect1 As RECT
    
    vRect = GetControlRect(ctrl.hWnd)
    vRect1 = GetControlRect(lstSelect.hWnd)
    
    lstSelect.Top = lstSelect.Top + (vRect.Top - vRect1.Top) + ctrl.Height + 10
    lstSelect.Left = lstSelect.Left + (vRect.Left - vRect1.Left)
    lstSelect.Width = ctrl.Width

End Sub

Private Function CalcSVG(ByVal strVal As String) As Currency
'   ��ֵ
    Dim varInData As Variant, curX As Currency, i As Integer
    If Left(strVal, 1) = "," Then
        varInData = Split(Mid(strVal, 2), ",")
    Else
        varInData = Split(strVal, ",")
    End If
    For i = LBound(varInData) To UBound(varInData)
        curX = curX + Val(varInData(i))
    Next
    If i > 0 Then
        CalcSVG = curX / i
    End If
End Function
Private Function CalcSD(ByVal strVal As String) As Currency
    '��׼��
    Dim varInData As Variant, curX As Currency, i As Integer, cur��ֵ As Currency
    
    If Left(strVal, 1) = "," Then
        varInData = Split(Mid(strVal, 2), ",")
    Else
        varInData = Split(strVal, ",")
    End If
    cur��ֵ = CalcSVG(strVal)
    For i = LBound(varInData) To UBound(varInData)
        curX = curX + (Val(varInData(i)) - cur��ֵ) ^ 2
    Next
    If i - 1 > 0 Then
        CalcSD = Sqr(curX / (i - 1))
    End If
    'Sqr (��(xn - x��) ^ 2 / (N - 1))
End Function

